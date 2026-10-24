#!/usr/bin/env node
/**
 * What this host has said, question by question, across every round you have.
 *
 * `UNSTABLE_ANSWERS` in `host-baseline.mjs` exists because this host answers the
 * same question differently on different runs of one build, and an entry there
 * is a warning not to build on whichever answer a sheet happens to carry. The
 * entries are written by hand from whatever rounds the author had open, which
 * means they go stale silently: one said "ALTERNATES … it is a coin" about a
 * question that had by then answered the same way ten rounds running.
 *
 * Deciding whether to swap the committed fixture needs the same arithmetic —
 * CLAUDE.md says replace it when a new sheet answers AT LEAST AS MUCH, and
 * "count the ANSWERS" is a thing somebody has to actually count.
 *
 * So: point this at the rounds.
 *
 *   node scripts/host-history.mjs <round1.json> <round2.json> …
 *   node scripts/host-history.mjs --fixture <round*.json>   # also score each
 *                                                           # against the
 *                                                           # committed sheet
 *
 * Rounds are not in the repo — they are uploads. Keep them somewhere and pass a
 * glob. Exit 0 always: this reports, it does not gate.
 *
 * ORDERED BY THE ROUND'S OWN TIMESTAMP, never by the order the shell listed the
 * files. This tool's entire output is directional — "settled" counts from the
 * END, `latest` is the last column, and the fixture table's bottom line is read
 * as "the newest sheet" — and the docstring above tells you to pass a GLOB,
 * whose order is alphabetical by filename. The archive names rounds
 * `NNN-<sha>.json`, which happens to list oldest first until round 1000 — a
 * property of the naming, not a promise of the shell, and files a person lists by
 * hand have no such order. The sort is `byRoundTime` below.
 *
 * The cost was not hypothetical: the six rounds of 2026-08-11 in glob order made
 * `shapes-items-count-honest` read `steady lately — "unreadable" x 5` when the
 * NEWEST round had said `short-0`, and in true order the same six say
 * `UNSTABLE`. Opposite verdicts, same data. That is precisely the staleness this
 * tool was written to stop `UNSTABLE_ANSWERS` suffering by hand.
 */
import { readFileSync } from "fs";
import { FAKE_BASELINE, answersOf, diffAnswers, sheetOf } from "./host-baseline.mjs";
import { isMain } from "./is-main.mjs";

/**
 * A round's answers, the build that produced them, and when the driver started it.
 *
 * `startedAt` is the ROUND'S OWN clock (`driverRun.startedAt`, written by the
 * driver the moment the round begins), and it is on the round file itself, not
 * on the sheet inside it. Rounds archived before the driver recorded it carry
 * none, and `byRoundTime` copes with that.
 */
export function readRound(path) {
  const file = JSON.parse(readFileSync(path, "utf8"));
  const sheet = sheetOf(file) ?? file;
  return { path, build: sheet?.build ?? "?", startedAt: file?.driverRun?.startedAt, answers: answersOf(file) ?? {} };
}

/**
 * Oldest first: by the build's stamp, then by the round's own start time.
 *
 * `build` is `"<sha> · <YYYY-MM-DD HH:MMZ>"`, which is when vite BUILT the
 * bundle, to the minute — not when a round ran. It orders builds, and it is the
 * same string on every round of one build. A cycle takes two or three rounds per
 * build, so those tie, and `Array.prototype.sort` is stable: the tied rounds kept
 * whatever order the shell listed them in, which is exactly what the docstring
 * at the top of this file says must never happen. Handed the three rounds of one
 * build newest-first, `latest` and `streak` named the OLDEST (measured on rounds
 * 495-497: nine questions printed their sequence reversed; the whole archive
 * handed over newest-first moved two verdicts from UNSTABLE to "steady lately").
 * Within a build, the tie is therefore broken by `startedAt`.
 *
 * Round files are named `NNN-<sha>.json`, zero-padded and counting up, so a
 * glob already lists them oldest first and the archive never showed this. It
 * stops being true at round 1000, where `1000-…` sorts before `999-…`, and it
 * was never true of rounds a person lists by hand.
 *
 * A round whose stamp cannot be parsed sorts LAST and keeps its relative order,
 * which is the safe direction: an unparseable stamp is most likely a hand-made
 * or future file, and putting it at the end makes it visible in the `latest`
 * column rather than silently reordering real rounds around it. A round with no
 * `startedAt` sorts after the ones that have one, by the same rule.
 */
export function byRoundTime(a, b) {
  const stamp = (r) => {
    const m = /(\d{4}-\d{2}-\d{2}[T ][\d:]+Z?)/.exec(String(r.build ?? ""));
    const v = m ? Date.parse(m[1].replace(" ", "T")) : NaN;
    return Number.isFinite(v) ? v : Infinity;
  };
  const started = (r) => {
    const v = Date.parse(String(r.startedAt ?? ""));
    return Number.isFinite(v) ? v : Infinity;
  };
  // `-` is not safe on two Infinities (NaN), and a comparator returning NaN is
  // read as "equal" by accident rather than by decision.
  const cmp = (x, y) => (x === y ? 0 : x < y ? -1 : 1);
  return cmp(stamp(a), stamp(b)) || cmp(started(a), started(b));
}

/**
 * Per question, what each round said — and whether it has settled.
 *
 * "Settled" counts from the END, and deliberately ignores the words that mean
 * the question was never put: a run that could not set the probe up says
 * nothing about the host, and letting `no-scratch-slide` break a streak would
 * report every question as unstable on a bad night.
 */
export function history(rounds, neverAsked = new Set(["no-scratch-slide", "no-scratch-shape", "not-asked"])) {
  const ids = [...new Set(rounds.flatMap((r) => Object.keys(r.answers)))].sort();
  return ids.map((id) => {
    const seq = rounds.map((r) => r.answers[id] ?? "—");
    const real = seq.filter((a) => a !== "—" && !neverAsked.has(a));
    let streak = 0;
    for (let i = real.length - 1; i >= 0 && real[i] === real[real.length - 1]; i--) streak++;
    return {
      id,
      seq,
      asked: real.length,
      distinct: [...new Set(real)],
      latest: real[real.length - 1],
      streak,
    };
  });
}

function main(argv) {
  const files = argv.filter((a) => !a.startsWith("--"));
  if (!files.length) {
    console.error("usage: node scripts/host-history.mjs <round1.json> <round2.json> …");
    return;
  }
  const rounds = files.map(readRound).sort(byRoundTime);
  console.log(`\n  ${rounds.length} round(s):`);
  for (const r of rounds) console.log(`    ${r.build}`);

  console.log(`\n  WHAT THIS HOST SAYS, question by question\n`);
  for (const h of history(rounds)) {
    if (h.asked === 0) continue;
    const verdict =
      h.distinct.length === 1
        ? `SETTLED — ${h.asked} of ${h.asked} say "${h.latest}"`
        : h.streak >= 5
          ? `steady lately — "${h.latest}" × ${h.streak}, but ${h.distinct.length} faces overall (${h.distinct.join(", ")})`
          : `UNSTABLE — ${h.distinct.length} faces (${h.distinct.join(", ")})`;
    console.log(`    ${h.id}`);
    console.log(`      ${h.seq.join(" ")}`);
    console.log(`      ${verdict}\n`);
  }

  // The fixture-swap arithmetic CLAUDE.md asks for, done rather than eyeballed.
  console.log(`  AGAINST THE FAKE — how much each round would compare, as a CI fixture\n`);
  for (const r of rounds) {
    const d = diffAnswers(r.answers, FAKE_BASELINE);
    console.log(
      `    ${String(d.agree.length + d.differ.length).padStart(3)} answered · ` +
        `${String(d.notAsked.length).padStart(2)} never put · ` +
        `${String(d.onlyFake.length).padStart(2)} the sheet predates   ${r.build}`,
    );
  }
  console.log("\n  Swap the committed fixture only for a sheet that ANSWERS at least as much.\n");
}

if (isMain(import.meta.url, process.argv[1])) main(process.argv.slice(2));
