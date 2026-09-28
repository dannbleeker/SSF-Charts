#!/usr/bin/env node
/**
 * Is a question that flipped mid-round a COIN, or a function of the host's state?
 *
 * The probe asks every question several times across a round and stamps each
 * sample with the regime the host was in when it was taken (`healthy`,
 * `slide-trouble`, `collection-refused`, `unknown` — `regimeFrom` in
 * `src/render/host-probe.ts`, shipped in #390). The pane then reports any
 * question that CHANGED ITS ANSWER MID-ROUND, and such a question is written
 * into `UNSTABLE_ANSWERS` as a warning not to build on either face.
 *
 * Nothing has ever read the regime stamps mechanically. They have been read by
 * hand, one entry at a time, in prose — which is how `UNSTABLE_ANSWERS` came to
 * call a question "a coin" that had answered the same way ten rounds running.
 * `host-history.mjs` closed that for the ACROSS-rounds question and reads only
 * each round's final answer; this closes the WITHIN-round one.
 *
 * The distinction is not cosmetic. "A coin" means build on nothing. "Answers
 * `threw` while the collection is refused and `yes` while healthy" is not a coin
 * at all — it is a documented degradation, the caller can test for the regime,
 * and the entry is mislabelled.
 *
 *   node scripts/host-regimes.mjs <round1.json> <round2.json> …
 *
 * Rounds are uploads, not repo files. Exit 0 always: this reports, it does not
 * gate.
 *
 * WHY THERE ARE FOUR VERDICTS AND NOT TWO. With three passes landing in three
 * different regimes, "every regime maps to one answer" is true by construction —
 * it cannot fail, so it is not evidence. A claim that had no chance to be
 * refuted is reported as `untested`, never as `explained`. This repo has been
 * bitten by exactly that shape before: a text-versus-mark sweep that "found
 * nothing" because it matched no marker at all, and a frame gate whose
 * approximate metric invented two exceptions that were then written into a test
 * as a permanent allowance. A measurement that cannot come out the other way is
 * not a measurement.
 */
import { readFileSync } from "fs";
import { NEVER_ASKED, UNINFORMATIVE_ANSWERS, sheetOf } from "./host-baseline.mjs";
import { isMain } from "./is-main.mjs";

/** The samples a round recorded per question, with the build that produced them. */
export function readRoundSamples(path) {
  const file = JSON.parse(readFileSync(path, "utf8"));
  const sheet = sheetOf(file);
  const answers = Array.isArray(sheet?.answers) ? sheet.answers : [];
  return { path, build: sheet?.build ?? "?", answers };
}

/**
 * Did the host's state account for this question changing its answer?
 *
 * Samples whose answer means the question was never put are dropped first, for
 * the reason `history()` gives: a pass that could not set the probe up says
 * nothing about the host, and letting `no-scratch-slide` count as a face would
 * report every question as a flipper on a bad night.
 *
 * Returns one of:
 *   steady     — the real samples agree; nothing to explain
 *   explained  — every regime maps to a single answer, AND some regime was
 *                sampled more than once, so the mapping could have failed
 *   untested   — every regime maps to a single answer, but no regime repeated:
 *                unfalsifiable, so it is not evidence either way
 *   coin       — some regime carries two different answers, so host state does
 *                not account for the flip
 */
export function explainBy(samples, field = "regime", neverAsked = NEVER_ASKED, uninformative = UNINFORMATIVE_ANSWERS) {
  /**
   * BOTH TIERS OF NON-ANSWER, not just the first.
   *
   * `NEVER_ASKED` is "the question was never put". `UNINFORMATIVE_ANSWERS` is
   * "the question was put and produced nothing to name" — `other`, `unreadable`,
   * `silent`, `not-a-short-read`, `none-of-ours`, `no-shapes-to-list`. This
   * filtered only the first, so the second counted as a FACE: a question
   * answered `yes` twice and `unreadable` once looked like a question that had
   * changed its answer, and the machinery below then set about explaining a
   * flip that never happened.
   *
   * Measured 2026-09-28 against the committed archive: `host-regimes.mjs
   * rounds/302-aa459dd.json` reported `picture-then-shape-read` as EXPLAINED by
   * regime, and another question as COIN because "healthy gave threw and
   * unreadable". The host answered consistently; the non-answer was the second
   * face. Both verdicts were manufactured.
   *
   * This file's own docstring already states the rule — a word that says
   * "nothing about the host" must not count as a face — and `host-baseline.mjs`
   * applies both tiers at its line 1211. The rule was right and one of its two
   * halves was missing here.
   */
  const real = (samples ?? []).filter((s) => s && !neverAsked.has(s.answer) && !uninformative.has(s.answer));
  const faces = [...new Set(real.map((s) => s.answer))];
  const byRegime = new Map();
  for (const s of real) {
    const regime = s[field] ?? "unknown";
    if (!byRegime.has(regime)) byRegime.set(regime, []);
    byRegime.get(regime).push(s.answer);
  }
  const mapping = [...byRegime.entries()].map(([regime, answers]) => ({
    regime,
    answers: [...new Set(answers)],
    samples: answers.length,
  }));
  if (faces.length <= 1) return { verdict: "steady", faces, mapping, real: real.length };

  // The stamp never moved for this question, so it cannot explain anything it
  // did. Reported apart from `coin` because the two look identical in the data
  // — one bucket holding two answers — and mean opposite things: `coin` is a
  // fact about the HOST, `blind` is a fact about the STAMP. Conflating them is
  // how a useless field gets read as evidence the host is random.
  //
  // Not hypothetical. Tested retrospectively against round 17, the scratch-slide
  // stamp is `fresh-slide` on all three samples of the very question it was
  // added for — that question replaces its slide 0.2-0.5s before every ask — so
  // without this branch the tool would have answered `coin` and the stamp's own
  // blindness would have been invisible.
  if (mapping.length === 1) return { verdict: "blind", faces, mapping, real: real.length };

  const split = mapping.filter((m) => m.answers.length > 1);
  if (split.length) return { verdict: "coin", faces, mapping, real: real.length, split };

  // Every regime maps to one answer. That is only evidence if it could have
  // come out otherwise — i.e. if some regime was asked twice and agreed.
  const repeated = mapping.filter((m) => m.samples > 1);
  return {
    verdict: repeated.length ? "explained" : "untested",
    faces,
    mapping,
    real: real.length,
    repeated: repeated.map((m) => m.regime),
  };
}

/** The regime stamp, which is what this tool was written for. */
export const explainByRegime = (samples, neverAsked = NEVER_ASKED) => explainBy(samples, "regime", neverAsked);

/**
 * Two questions that both flipped: did they flip at the SAME pass boundary?
 *
 * A partner question is asked precisely because one answer admits two readings,
 * and `UNSTABLE_ANSWERS` already records a case where a partner was the coin
 * while its trigger held still. Flipping in lockstep is evidence of ONE
 * mechanism sampled twice; flipping at different boundaries is evidence of two.
 *
 * Compares only the passes where BOTH questions have a real answer — a pass one
 * of them never reached cannot witness agreement or disagreement.
 */
export function flippedTogether(aSamples, bSamples, neverAsked = NEVER_ASKED) {
  const byPass = (ss) => {
    const m = new Map();
    for (const s of ss ?? []) if (s && !neverAsked.has(s.answer)) m.set(s.pass, s.answer);
    return m;
  };
  const a = byPass(aSamples);
  const b = byPass(bSamples);
  const passes = [...a.keys()].filter((p) => b.has(p)).sort((x, y) => x - y);
  if (passes.length < 2) return { verdict: "not-comparable", shared: passes.length };

  const changes = [];
  for (let i = 1; i < passes.length; i++) {
    const movedA = a.get(passes[i]) !== a.get(passes[i - 1]);
    const movedB = b.get(passes[i]) !== b.get(passes[i - 1]);
    if (movedA || movedB) changes.push({ from: passes[i - 1], to: passes[i], movedA, movedB });
  }
  if (!changes.length) return { verdict: "neither-moved", shared: passes.length, changes };
  const together = changes.every((c) => c.movedA === c.movedB);
  const both = changes.some((c) => c.movedA && c.movedB);
  return {
    verdict: together && both ? "lockstep" : "independent",
    shared: passes.length,
    changes,
  };
}

/**
 * Questions that actually ask the shape collection, and what it said.
 *
 * The probe's own test (`aboutTheCollection` in `src/render/host-probe.ts`) and
 * its refusal vocabulary, duplicated here because a .mjs script cannot import
 * from the TypeScript source — the same unavoidable duplication `NEVER_ASKED`
 * already carries, and held in step the same way, by a source scan in
 * `test/host-regimes.test.ts`.
 */
export const COLLECTION_QUESTION = /^shapes?-/;
export const COLLECTION_REFUSALS = new Set(["unreadable", "short-0", "not-listed"]);

/**
 * What the shape collection did, in time order, across a whole round.
 *
 * This host's central defect is that it stops listing shapes a run just added,
 * and every expensive failure downstream — no group, no tag target, a chart
 * that is not re-editable — is that one thing. It has been reasoned about from
 * batch timings ("FAST IS THE BROKEN MODE"), from grouping traces, and from the
 * `regime` stamp. It had never simply been READ.
 *
 * Reading it is what found the stamp's own bug: in round 17 the collection
 * ANSWERED 14 of the 28 times it was asked, interleaved with the refusals
 * throughout, while 88% of that round's samples were stamped
 * `collection-refused`. So the collection oscillates question by question, and
 * the instrument was reporting a one-way slide.
 *
 * `recovered` is the fact worth having every round: a collection that answers
 * again after refusing is a different host from one that goes quiet and stays
 * quiet, and the two want different code.
 */
export function collectionTimeline(answers, neverAsked = NEVER_ASKED) {
  const events = [];
  for (const a of answers ?? []) {
    if (!COLLECTION_QUESTION.test(a.id)) continue;
    for (const s of a.samples ?? []) {
      const verdict = neverAsked.has(s.answer)
        ? "never-put"
        : COLLECTION_REFUSALS.has(s.answer)
          ? "refused"
          : "answered";
      events.push({ atMs: s.atMs, id: a.id, answer: s.answer, verdict });
    }
  }
  events.sort((x, y) => x.atMs - y.atMs);
  // A never-put pass says nothing about the collection — the run could not set
  // the question up — so it is shown but never counted, the rule every other
  // reader here applies.
  const real = events.filter((e) => e.verdict !== "never-put");
  const answered = real.filter((e) => e.verdict === "answered").length;
  const refused = real.filter((e) => e.verdict === "refused").length;

  // THE SPLIT THAT MATTERS, and the one a timeline alone invites you to miss.
  //
  // Read in time order, round 17 looks like a collection that refuses, comes
  // back, refuses again — three clean bursts. Grouped by QUESTION it is nothing
  // of the kind: seven of its eight collection questions gave the SAME verdict
  // on every pass. `shapes-items-count-honest` refused all three times,
  // `shape-add-fresh-slide-proxy` answered all three. The interleaving is the
  // fixed question order cycling through the passes, not the host changing.
  //
  // So "the collection is refusing" is, for the most part, a fact about WHICH
  // QUESTION was asked recently rather than about when. Only the questions that
  // actually varied are evidence of time-dependence, and this reports them
  // separately for that reason.
  const byQuestion = new Map();
  for (const e of real) {
    if (!byQuestion.has(e.id)) byQuestion.set(e.id, new Set());
    byQuestion.get(e.id).add(e.verdict);
  }
  const alwaysAnswered = [];
  const alwaysRefused = [];
  const varied = [];
  for (const [id, verdicts] of byQuestion) {
    if (verdicts.size > 1) varied.push(id);
    else if (verdicts.has("refused")) alwaysRefused.push(id);
    else alwaysAnswered.push(id);
  }
  return {
    events,
    answered,
    refused,
    firstRefusalAt: real.find((e) => e.verdict === "refused")?.atMs,
    alwaysAnswered,
    alwaysRefused,
    varied,
    // True only when some question gave BOTH verdicts across the round. A bare
    // "an answer came after a refusal" is satisfied by question order alone and
    // would report recovery on a host that never changed.
    variesOverTime: varied.length > 0,
  };
}

/**
 * How concentrated a stamp is across a whole round.
 *
 * "A field that nearly every sample shares cannot separate anything" is the
 * criterion `regimeFrom`'s own docstring sets, having been written after a
 * sticky flag put the same value on 55 of 65 samples. It is the mistake this
 * project has now recorded against `idleMs`, `afterAnswering`, the settle's
 * shared label, `listChartsInDeck` and `onSlide` — five fields, five times, and
 * each was caught by a human reading a round rather than by anything automatic.
 *
 * So it is measured. A stamp whose commonest value covers `SATURATED_AT` or
 * more of the samples is reported as saturated, whatever its verdicts say —
 * because `untested` and `blind` are what a saturated field produces, and both
 * read as caution rather than as "this field is not working".
 */
export const SATURATED_AT = 0.8;

export function stampSpread(answers, field, neverAsked = NEVER_ASKED) {
  const tally = new Map();
  let total = 0;
  for (const a of answers ?? []) {
    for (const s of a.samples ?? []) {
      if (!s || neverAsked.has(s.answer)) continue;
      const v = s[field] ?? "unknown";
      tally.set(v, (tally.get(v) ?? 0) + 1);
      total++;
    }
  }
  const byValue = [...tally.entries()].sort((x, y) => y[1] - x[1]);
  const top = byValue[0];
  const share = total ? (top?.[1] ?? 0) / total : 0;
  return { field, total, byValue, share, saturated: total > 0 && share >= SATURATED_AT };
}

/**
 * Of the questions this round never put, what state was the host in when it
 * tried? Answers "was the low yield the host being unwell, or something else",
 * which a bare count of never-put questions cannot.
 */
export function neverPutByRegime(answers, neverAsked = NEVER_ASKED) {
  const tally = new Map();
  const ids = [];
  for (const a of answers ?? []) {
    if (!neverAsked.has(a.answer)) continue;
    ids.push(a.id);
    for (const s of a.samples ?? []) {
      if (!neverAsked.has(s.answer)) continue;
      const regime = s.regime ?? "unknown";
      tally.set(regime, (tally.get(regime) ?? 0) + 1);
    }
  }
  return { ids, byRegime: [...tally.entries()].sort((x, y) => y[1] - x[1]) };
}

/**
 * The verdict as a line of prose.
 *
 * A `switch`, not a lookup table keyed by the verdict — because a table builds
 * EVERY branch's string before choosing one, and `split` is only present on a
 * `coin` while `repeated` is only present on the other two. The first version
 * was that table and it threw on the first real round it was pointed at, on a
 * question whose verdict was perfectly well formed. The unit tests could not
 * see it: they call the decision functions and never rendered a line, so the
 * whole reporting layer was unguarded. Exported so it is now testable at all.
 */
export function verdictLine(r) {
  switch (r.verdict) {
    case "coin":
      return `COIN — host state does NOT account for it; ${r.split.map((s) => `${s.regime} gave ${s.answers.join(" and ")}`).join(", ")}`;
    case "explained":
      return `EXPLAINED by regime — and it could have failed (${r.repeated.join(", ")} sampled more than once)`;
    case "untested":
      return `UNTESTED — the mapping fits, but no regime was sampled twice, so it could not have come out otherwise`;
    case "blind":
      return `BLIND — the stamp read "${r.mapping[0].regime}" on every sample, so it says nothing about this question either way`;
    default:
      return "steady";
  }
}

function describe(id, r) {
  const map = r.mapping
    .map((m) => `${m.regime}→${m.answers.join("/")}${m.samples > 1 ? ` ×${m.samples}` : ""}`)
    .join("  ");
  return `    ${id}\n      ${map}\n      ${verdictLine(r)}\n`;
}

function main(argv) {
  const files = argv.filter((a) => !a.startsWith("--"));
  if (!files.length) {
    console.error("usage: node scripts/host-regimes.mjs <round1.json> <round2.json> …");
    return;
  }
  for (const round of files.map(readRoundSamples)) {
    console.log(`\n  ${round.build}   (${round.path})`);

    const flippers = round.answers
      .map((a) => ({ id: a.id, r: explainByRegime(a.samples), scratch: explainBy(a.samples, "scratch") }))
      .filter((x) => x.r.verdict !== "steady");

    console.log(`\n  QUESTIONS THAT CHANGED ANSWER MID-ROUND — is host state the reason?\n`);
    if (!flippers.length) console.log("    none — every question held still within the round\n");
    for (const f of flippers) {
      console.log(describe(f.id, f.r));
      // Round 17 eliminated `regime` as the state that flips: a question can
      // agree perfectly with its partner at every instant and still answer two
      // ways inside one regime. `scratch` is the candidate that replaced it, so
      // both stamps are read side by side — a round where one says COIN and the
      // other EXPLAINED is the whole point of having added the second.
      if (f.scratch.mapping.some((m) => m.regime !== "unknown"))
        console.log(`      by scratch slide:  ${verdictLine(f.scratch)}\n`);
    }

    // Any two flippers may be one mechanism sampled twice. Say so, rather than
    // leaving it to be reasoned about from two entries in a table.
    if (flippers.length > 1) {
      console.log(`  DID THEY FLIP TOGETHER?\n`);
      const sampleOf = (id) => round.answers.find((a) => a.id === id)?.samples;
      for (let i = 0; i < flippers.length; i++) {
        for (let j = i + 1; j < flippers.length; j++) {
          const t = flippedTogether(sampleOf(flippers[i].id), sampleOf(flippers[j].id));
          console.log(`    ${flippers[i].id} × ${flippers[j].id}`);
          console.log(`      ${t.verdict} (${t.shared} pass(es) both answered)\n`);
        }
      }
    }

    // Before any verdict is believed, say whether each stamp could have
    // discriminated at all. A saturated field produces `untested` and `blind`,
    // and both of those read as caution rather than as a broken instrument.
    const ct = collectionTimeline(round.answers);
    if (ct.answered + ct.refused) {
      console.log(`  WHAT THE SHAPE COLLECTION DID — the defect everything else hangs off\n`);
      for (const e of ct.events) {
        const mark = e.verdict === "answered" ? "ok      " : e.verdict === "refused" ? "REFUSED " : "never-put";
        console.log(`    ${(e.atMs / 1000).toFixed(1).padStart(7)}s  ${mark} ${e.answer.padEnd(18)} ${e.id}`);
      }
      console.log(`\n    ${ct.answered} answered, ${ct.refused} refused across the round`);
      console.log(
        `    ${ct.alwaysAnswered.length} question(s) always answered, ${ct.alwaysRefused.length} always refused`,
      );
      if (!ct.variesOverTime) {
        console.log(`\n    NOT A TIME EFFECT — every collection question gave the SAME verdict on every pass.`);
        console.log(`    The split is by QUESTION, not by when it was asked, and a stamp that says`);
        console.log(`    "the collection is refusing" is reporting which question ran recently.\n`);
      } else {
        console.log(`\n    varied across passes: ${ct.varied.join(", ")}`);
        console.log(`    Only these are evidence the collection changed over time; the rest is question order.\n`);
      }
    }

    console.log(`  COULD THESE STAMPS SEPARATE ANYTHING?\n`);
    for (const field of ["regime", "scratch"]) {
      const sp = stampSpread(round.answers, field);
      if (!sp.total) continue;
      const dist = sp.byValue.map(([v, n]) => `${v} x${n}`).join("  ");
      console.log(`    ${field.padEnd(8)} ${dist}`);
      if (sp.saturated)
        console.log(
          `      SATURATED — "${sp.byValue[0][0]}" is ${Math.round(sp.share * 100)}% of samples; this stamp cannot separate much\n`,
        );
      else console.log(`      spread enough to discriminate (commonest ${Math.round(sp.share * 100)}%)\n`);
    }

    const np = neverPutByRegime(round.answers);
    if (np.ids.length) {
      console.log(`  NEVER PUT — ${np.ids.length} question(s), and the state the host was in\n`);
      for (const [regime, n] of np.byRegime) console.log(`    ${String(n).padStart(3)} attempt(s)  ${regime}`);
      console.log(`\n    ${np.ids.join(", ")}\n`);
    }
  }
  // Bound the claim in the output itself. Every verdict here is about ONE
  // round: `explained` means host state accounted for the flip in THIS round
  // and the mapping had a way to come out wrong, not that the question is
  // settled. This project has read a single round as a mechanism before — the
  // three consecutive `+108 byte` deltas that turned out to be a coincidence of
  // compression — so the footer says what a reader may take away.
  console.log("  Each verdict is about ONE round. EXPLAINED means host state accounted for");
  console.log("  the flip here and could have failed to; it is a reason to re-read the");
  console.log("  UNSTABLE_ANSWERS entry, not on its own a reason to rewrite it. The same");
  console.log("  verdict on a second round is what makes it a mechanism.\n");
  console.log("  COIN is the strong one: it survives however many regimes were sampled,");
  console.log("  because a single regime already answered two ways. UNTESTED is neither.\n");
}

if (isMain(import.meta.url, process.argv[1])) main(process.argv.slice(2));
