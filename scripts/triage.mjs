#!/usr/bin/env node
/**
 * Read a real run's .pptx and its run log together, and say where they differ.
 *
 * These two files are the whole evidence base for every host bug this project
 * has fixed, and joining them has always been done by hand: unzip the deck,
 * pretty-print 160 KB of JSON, line the slots up, squint. That is a slow job
 * and — twice now — a wrong one. One analysis read a truncated log and blamed
 * the wrong subsystem; another read a trace buffer that spanned three runs and
 * found a contradiction that was not there. Both cost a full round trip: a
 * deploy, a run in a real PowerPoint, an upload, and a re-read.
 *
 * The join itself is mechanical, so it belongs in a script. Slot by slot:
 * what the run believed it did, what the file actually holds, and the verdict
 * when those disagree.
 *
 *   node scripts/triage.mjs <deck.pptx> <run-log.json> [--all] [--json]
 *
 * Arguments may be given in either order — the extension says which is which.
 *
 * Two files, two authorities, and they are NOT equal: the deck is what the
 * user has, the log is what the add-in thought. Where they conflict the deck
 * wins, and the conflict is the finding. (A 39-slide run reported 20 tagged
 * charts; the file carried 31. The log agreed with itself and was false.)
 *
 * Exit code 0 when the two agree, 1 when they do not, 2 when a file cannot be
 * read — so it can gate as well as inform.
 */
import { readFileSync, readdirSync } from "fs";
import { checkClaims } from "./claims.mjs";
import { join } from "path";
import { readDeck, faultsIn } from "./verify-deck.mjs";
// THE ONE COPY, imported rather than repeated. `host-probe.ts` and
// `host-baseline.mjs` are pinned to each other by a test; this file kept a THIRD
// copy inside a function body, and on 2026-08-25 a new word — `no-named-slide` —
// was added to the two that are pinned and not to this one. A prediction reading
// that answer would have judged a question "put and answered" when the harness
// never asked it.
import { NEVER_ASKED as NOT_ASKED_ANSWERS, UNINFORMATIVE_ANSWERS } from "./host-baseline.mjs";

/**
 * Verdicts, worst first — the order they are printed and counted in.
 *
 * `repaired` and `foreign` are not faults: the first is the end-of-run repair
 * pass doing its job, the second is an earlier run's slide sitting in the same
 * deck, which is the case the run token exists to survive.
 */
const BAD = new Set(["lost", "duplicated", "tag-lost", "not-editable", "no-config", "blank", "orphan"]);

/**
 * What the deck holds for one slot, against what the run believed.
 *
 * `item.chart` — was this item MEANT to carry a config — is load-bearing here,
 * and its absence is why the first version of this file called seven perfectly
 * healthy slides broken. Not every demo item is a chart: the title page, the
 * contents pages and several elements are drawn as `PowerChart` objects with
 * no config by design, and a rule of "a chart object with no config is not
 * re-editable" flags every one of them. What the run *intended* cannot be
 * recovered from the file, and on the shape path it cannot be recovered from
 * `tagged` either — false there means both "never had one" and "the write was
 * lost". So the log states it, and a log too old to state it gets the weaker
 * `no-config` verdict rather than a confident wrong one.
 */
function verdictFor(item, rows) {
  if (rows.length === 0) return item.status === "skipped" ? "skipped" : "lost";
  if (rows.length > 1) return "duplicated";
  const row = rows[0];
  if (row.shapes === 0) return "blank";
  // `configOnChart`, NOT `config`. `verify-deck.mjs` draws that distinction on
  // purpose and says why on the field itself: `config` only reports that the
  // tag part EXISTS, which is equally true of a tag nothing in the shape tree
  // points at. The pane loads a chart by walking from the SSF Charts object to
  // its tag rid, so a config anchored on the SLIDE is a config the user cannot
  // reach — the chart is not re-editable, which is the whole thing this report
  // is about.
  //
  // triage was never migrated when that field was added, so for the exact case
  // `verify-deck` was extended to catch it printed `ok` and counted zero
  // disagreements, while verify-deck's own report on the same deck said
  // `orphan`. `--json` consumers saw `verdict: "ok"`.
  //
  // Falls back to `config` for a deck read by an older `verify-deck` that did
  // not report the anchor: unknown must not become a fault, and the two agree
  // wherever the anchor is known.
  const reachable = row.configOnChart ?? row.config;
  if (reachable) {
    // The repair pass got there after the run gave up on the tag — a success,
    // and one worth seeing, because it is the difference between a fix working
    // and a fix never having been needed.
    return item.tagged ? "ok" : "repaired";
  }
  // The run says it wrote a config tag and the file has none. On the file path
  // this is impossible by construction (the tag is written into the .pptx
  // before the host ever sees it), so it means the deck lost it afterwards; on
  // the shape path it means the write was reported as landing and did not.
  if (item.tagged) return "tag-lost";
  if (!row.chartObject) return "ok";
  if (item.chart === true) return "not-editable";
  if (item.chart === false) return "ok"; // never had a config to lose
  return "no-config"; // a log from before `chart` existed — intent unknown
}

/** A one-line description of what the file holds for a slot. */
function describeRow(row) {
  if (!row) return "—";
  const kind = row.picture ? "picture" : row.groups ? "group" : row.chartObject ? "shape" : "loose";
  // Names the orphan case rather than calling it `config`: a tag the pane
  // cannot reach is not the same fact as a tag it can, and the deck column is
  // where a reader looks to see which.
  const cfg = row.configOrphaned ? "config-ORPHANED" : (row.configOnChart ?? row.config) ? "config" : "no-config";
  return `${kind} ${row.shapes}sh ${cfg}${row.stamped ? " NOT-COMPLETE" : ""}`;
}

export function triage(deck, log) {
  // Logs written before the run token existed have no join key of their own.
  // The deck's commonest token is the best available stand-in and is very
  // probably right, but it IS a guess, so it gets labelled as one rather than
  // quietly presented as fact.
  const tally = new Map();
  for (const r of deck.rows) if (r.run) tally.set(r.run, (tally.get(r.run) ?? 0) + 1);
  const commonest = [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
  const run = log.run ?? commonest;
  const inferredRun = !log.run;

  const mine = new Map();
  const foreign = [];
  const unowned = [];
  for (const row of deck.rows) {
    if (row.run && row.run !== run) foreign.push(row);
    else if (row.slot === null) unowned.push(row);
    else {
      if (!mine.has(row.slot)) mine.set(row.slot, []);
      mine.get(row.slot).push(row);
    }
  }

  const items = log.items ?? [];
  const slots = items.map((item, i) => {
    const rows = mine.get(i) ?? [];
    return {
      slot: i,
      title: item.title,
      status: item.status,
      wroteTag: !!item.tagged,
      deck: describeRow(rows[0]),
      rows: rows.length,
      verdict: verdictFor(item, rows),
    };
  });
  // A slide carrying a slot this run never issued. Only possible when a tag
  // survived from a longer run into a shorter one, or when a slot tag was
  // written wrong — either way nothing owns it and nothing will clean it.
  // Outside the run's range in EITHER direction. Only `>= items.length` was
  // checked, so a slide tagged with a negative slot — a tag written wrong, or
  // carried over from something that was not this — matched no item, fell
  // outside the orphan test, and was reported by nothing at all.
  const orphans = [...mine.keys()]
    .filter((s) => !Number.isInteger(s) || s < 0 || s >= items.length)
    .map((s) => ({ slot: s, verdict: "orphan", indexes: mine.get(s).map((r) => r.index + 1) }));

  const counts = {};
  for (const s of [...slots, ...orphans]) counts[s.verdict] = (counts[s.verdict] ?? 0) + 1;
  const disagreements = slots.filter((s) => BAD.has(s.verdict)).length + orphans.length;
  return { run, inferredRun, slots, orphans, foreign, unowned, counts, disagreements };
}

/**
 * The runs inside a log file, however that file is shaped.
 *
 * One click can now take both insert paths, so a log holds a list. Logs from
 * before that — every artifact captured so far — are a single run at the top
 * level, and are read as a list of one rather than being refused: the whole
 * point of this tool is reading the runs that already exist.
 */
export function runsIn(log) {
  return Array.isArray(log?.runs) ? log.runs : [log];
}

/**
 * The host self-test's verdicts, if this log is one.
 *
 * A self-test log carries no runs at all — the scenarios are not inserts and
 * own no slots — so `runsIn` gives an empty list and the whole report used to
 * come out blank, exit 0, and look like a clean deck. It is a file someone
 * hands this tool expecting an answer.
 */
export function selfTestIn(log) {
  return Array.isArray(log?.selftest) ? log.selftest : [];
}

/**
 * A crash log — the steps a run wrote before it stopped existing.
 *
 * Not a run log and not a self-test log: it holds no slots, no verdicts and no
 * structured trace, because it is written a line at a time to browser storage
 * by a run that may be about to die. It carries the one thing the other two
 * cannot: what a run that never ENDED was doing. Both files this tool was
 * built for are produced at the end of a run, and the runs worth reading are
 * the ones that do not get there.
 */
/**
 * One banked finding in a line: a scenario verdict as its verdict, an answer
 * sheet as its count, anything else as JSON.
 *
 * Kept shallow on purpose. This runs on the file from a round that died, and the
 * reader wants to know which scenarios got a verdict and whether the probe half
 * survived — not to read a sheet pretty-printed down the terminal.
 */
export function describeFinding(value) {
  if (value && typeof value === "object" && typeof value.name === "string" && "ok" in value) {
    const state = value.skipped ? "SKIPPED" : value.ok ? "passed" : "FAILED";
    return `${state}${value.detail ? ` — ${value.detail}` : ""}${value.ms ? ` (${(value.ms / 1000).toFixed(1)}s)` : ""}`;
  }
  if (value && typeof value === "object" && Array.isArray(value.answers)) {
    const asked = value.answers.filter((a) => a && typeof a.answer === "string").length;
    return `answer sheet, ${asked} of ${value.answers.length} question(s) answered`;
  }
  return typeof value === "string" ? value : JSON.stringify(value);
}

export function crashLogIn(log) {
  return (log?.kind === "ssf-charts-crash-log" || log?.kind === "powerchart-crash-log") && Array.isArray(log.steps)
    ? log
    : null;
}

/**
 * Failure signatures that belong to PowerPoint, not to SSF Charts.
 *
 * The first hour of the last diagnosis went into establishing that
 * `InvalidParam passed to GetItem(id)` was a host bug rather than ours. It is,
 * it is reported, and most of these are still open — so the trace can just say
 * so. A reader who sees an issue number stops looking for the mistake in this
 * repo.
 *
 * Matched on a substring of the reason string the trace recorded, FIRST MATCH
 * WINS, so the order below is part of the mapping rather than decoration.
 * Checked 2026-08-01; re-check before trusting the status (see RESEARCH.md 4b).
 */
const SELECTION_WEDGE =
  "office-js#3083 / #3698 family (open) — a programmatic setSelectedShapes wedges the web host's " +
  "selection subsystem; every selection call after it goes silent";

const KNOWN_HOST_BUGS = [
  // MOST SPECIFIC FIRST, and this is load-bearing rather than tidy. A wedged
  // selection subsystem does not throw — it goes quiet, so what the trace
  // records is a timeout, whose text is "did not respond while <phase>". The
  // generic sync-hang note below matches that too, and on a first-match-wins
  // table it would claim every one of them and send the reader to #5022: a
  // different bug, a different cause, and a fix that does not apply. The
  // phase name is the only thing that separates them, which is one more
  // reason every bounded wait carries one.
  ["stopped answering selection calls", SELECTION_WEDGE],
  ["did not respond while reading the selected chart", SELECTION_WEDGE],
  ["did not respond while selecting a shape", SELECTION_WEDGE],
  ["did not respond while clearing the shape selection", SELECTION_WEDGE],
  ["InvalidParam passed to GetItem", "office-js#2903 (closed: not planned) — stale shape proxy on web"],
  ["not available", "office-js#6363 (open: regression) — loaded property missing after sync"],
  [
    "did not respond",
    "office-js#5022 (closed: completed 2024-11-18, unverified on this host) — sync hangs after add/delete/re-read",
  ],
  [
    "Timed out",
    "office-js#5022 (closed: completed 2024-11-18, unverified on this host) — sync hangs after add/delete/re-read",
  ],
];

/** The known-bug note for a problem string, if there is one. */
export function knownBug(text) {
  const hit = KNOWN_HOST_BUGS.find(([sig]) => text.includes(sig));
  return hit ? hit[1] : null;
}

function pad(s, n) {
  return String(s ?? "—")
    .padEnd(n)
    .slice(0, n);
}

function report(deck, log, run, t, showAll) {
  const secs = ((run.totalMs ?? 0) / 1000).toFixed(1);
  console.log(
    `\n  run ${t.run ?? "unknown"}${t.inferredRun ? " (INFERRED from the deck — the log carries no run id)" : ""}` +
      `\n  build ${log.build ?? run.build ?? "?"} · ${log.host ?? run.host ?? "?"}` +
      `\n  path ${run.path ?? "?"} · ${secs}s\n`,
  );

  console.log(
    `  DECK ${deck.rows.length} slide(s): ${deck.rows.length - t.foreign.length - t.unowned.length} from this run, ` +
      `${t.foreign.length} from other run(s), ${t.unowned.length} carrying no slot tag`,
  );

  const shown = showAll ? t.slots : t.slots.filter((s) => BAD.has(s.verdict) || s.verdict === "repaired");
  console.log(
    `\n  SLOTS ${t.slots.length} expected · ${t.slots.length - (t.counts.lost ?? 0)} present · ` +
      Object.entries(t.counts)
        .sort((a, b) => b[1] - a[1])
        .map(([k, n]) => `${n} ${k}`)
        .join(" · "),
  );
  if (shown.length) {
    console.log(`\n  ${pad("#", 5)}${pad("title", 26)}${pad("log", 11)}${pad("tag?", 6)}${pad("deck", 30)}verdict`);
    for (const s of shown)
      console.log(
        `  ${pad(s.slot, 5)}${pad(s.title, 26)}${pad(s.status, 11)}${pad(s.wroteTag ? "wrote" : "no", 6)}` +
          `${pad(s.rows > 1 ? `${s.rows} slides claim it` : s.deck, 30)}${s.verdict}`,
      );
    if (!showAll) console.log(`  (agreeing slots hidden — pass --all to see all ${t.slots.length})`);
  }
  for (const o of t.orphans) console.log(`  orphan slot ${o.slot} on slide(s) ${o.indexes.join(", ")}`);

  reportTrace(run.trace);

  console.log(
    t.disagreements
      ? `\n  ${t.disagreements} slot(s) where the deck and the log disagree\n`
      : `\n  deck and log agree on every slot\n`,
  );
}

/**
 * The trace tallies, from wherever the trace hangs.
 *
 * A run carries its own; a self-test round is not a run and hangs one at the
 * top level, which is why this is a function rather than four lines inside
 * `report`. Both shapes have the same tallies and deserve the same reading.
 */
function reportTrace(trace) {
  if (trace?.summary) {
    console.log(`\n  TRACE ${trace.entries?.length ?? 0} entries${trace.dropped ? `, ${trace.dropped} dropped` : ""}`);
    for (const s of trace.summary.steps.slice(0, 8)) console.log(`    ${pad(s.n, 6)}${s.scope}  ${s.message}`);
    // Every phase an error escaped, whether or not it made the top 8 above.
    //
    // These are the most locating lines in the whole log and the least likely
    // to be common enough to rank: a run with one fatal read has one of them
    // against hundreds of ordinary steps. The `problems` tally below carries
    // what the HOST said, truncated to keep one debugInfo blob from swamping
    // it; this carries what the add-in was DOING, which is the half that used
    // to be missing entirely and took a session to reconstruct.
    const phases = trace.summary.steps.filter(
      (s) => s.scope === "error" && !trace.summary.steps.slice(0, 8).includes(s),
    );
    if (phases.length) {
      console.log(`  phases an error escaped:`);
      for (const s of phases) console.log(`    ${pad(s.n, 6)}${s.message}`);
    }
    if (trace.summary.problems.length) {
      console.log(`  problems:`);
      for (const p of trace.summary.problems.slice(0, 8)) {
        console.log(`    ${pad(p.n, 6)}${p.text}`);
        const known = knownBug(p.text);
        if (known) console.log(`    ${pad("", 6)}  ^ known host bug: ${known}`);
      }
    }
  } else if (trace) {
    console.log(`\n  TRACE ${trace.entries?.length ?? 0} entries (written before summaries existed — no tallies)`);
  }
}

/** The self-test verdicts, with the host bugs their own words give away. */
/**
 * Pool the rasterise/cheap-read arms across MANY rounds.
 *
 * `does a rasterise poison the next draw` is a counterbalanced control that
 * makes four draws a round — two after a rasterise, two after a cheap read,
 * interleaved so position cannot account for the difference. Four is enough to
 * spot a call that fails EVERY time and far too few for anything smaller, so
 * for eleven rounds it has correctly reported "no pattern" and correctly been
 * unable to say more.
 *
 * Pooled over eight rounds the same arms read 5 stalls in 16 draws after a
 * rasterise against 2 in 16 after a cheap read — 31% against 12.5%. That is a
 * direction, and it is still not an answer: Fisher's exact on those counts is
 * p≈0.39. Separating rates that close needs somewhere near 60-100 draws an arm,
 * which is 30-50 more rounds at two an arm each.
 *
 * So the verdict a single round can reach is not the interesting one, and the
 * evidence was accumulating in files nobody was adding up. This adds them up.
 *
 * A draw's outcome is NOT its `batch issued` line: that is logged before the
 * sync, deliberately, because the sync is where a bad host goes quiet — so
 * every stall has a `batch issued` of its own immediately before it. Pairing
 * the two is how a first pass at this produced 0 stalls in 32 draws, and then
 * an unrelated cut of the same data produced a 6x effect that was not there.
 * The outcome is whether a `gave up waiting` lands between this draw's step and
 * the next one.
 *
 * Both mistakes were made while that line was NAMED `batch committed`, which is
 * why it is not called that any more — but round files saved before 2026-08-11
 * still carry the old name, and this function reads neither, by design.
 */
/**
 * The two populations a draw batch falls into, computed rather than eyeballed.
 *
 * "FAST IS THE BROKEN MODE" rests on the batch times being BIMODAL with an
 * empty band between them, and every number supporting it has been read off a
 * round by hand. That is how the band's edge came to be quoted as 29.2s and
 * stayed quoted after a later round put three surviving batches at 29.3, 30.8
 * and 31.1 — a measurement drifting in the prose because nothing recomputed it.
 *
 * The split is taken at the LARGEST gap in the sorted times, which is what a
 * reader does by eye and cannot then mis-remember. A round with no gap worth
 * the name says so rather than inventing a boundary: `bimodal` is false unless
 * the gap is bigger than the whole spread of the fast group, which is the
 * property "an empty band" actually means.
 *
 * The other number this file re-derives by hand — the per-slide cost slope —
 * is deliberately NOT computed here. `onSlide` is a NET counter that resets as
 * each chart replaces itself, so one slide's points are a repeated ramp rather
 * than a series: slide 257 of round 17 carries onSlide 58 and 68 four times
 * each, non-monotonic. Fitting a line through that produced 0.29s per shape
 * against the 0.44-0.49 the rasterise arms give, which is a number belonging to
 * neither population. A slope needs an add-only path and this trace does not
 * mark one; read the arms by hand until it does.
 */
export function batchPopulations(log) {
  const entries = log?.trace?.entries;
  if (!Array.isArray(entries)) return null;
  const times = entries
    .filter((e) => e.message === "batch issued" && typeof e.data?.prevBatchMs === "number")
    .map((e) => e.data.prevBatchMs)
    .sort((a, b) => a - b);
  if (times.length < 4) return null;
  let at = 0;
  let gap = 0;
  for (let i = 0; i < times.length - 1; i++)
    if (times[i + 1] - times[i] > gap) {
      gap = times[i + 1] - times[i];
      at = i;
    }
  const fast = times.slice(0, at + 1);
  const slow = times.slice(at + 1);
  const spread = fast[fast.length - 1] - fast[0];
  // Both groups need at least two members. On evenly spread times every gap is
  // equal, the argmax picks the FIRST, and a one-member fast group has a spread
  // of zero — which any gap beats, so the rule would answer "bimodal" for the
  // most uniform data there is. Found by the negative test, not by reasoning.
  const bimodal = fast.length >= 2 && slow.length >= 2 && gap > spread;
  return { times, fast, slow, gap, bimodal, max: times[times.length - 1] };
}

export function poolRasteriseArms(logs) {
  const arms = { rasterise: { ok: 0, stall: 0 }, "cheap read": { ok: 0, stall: 0 } };
  let rounds = 0;
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    rounds++;
    const es = [...entries].sort((a, b) => (a.ms ?? 0) - (b.ms ?? 0));
    const draws = [];
    es.forEach((e, i) => {
      const m = /^after a (rasterise|cheap read) #\d+: drawing$/.exec(String(e.data?.what ?? ""));
      if (m) draws.push([i, m[1]]);
    });
    draws.forEach(([i, arm], k) => {
      const end = k + 1 < draws.length ? draws[k + 1][0] : es.length;
      let stalled = false;
      for (let z = i + 1; z < end; z++) {
        const e = es[z];
        if (e.scope === "host" && e.message === "gave up waiting" && /drawing shapes/.test(String(e.data?.what ?? "")))
          stalled = true;
        if (/scenario (passed|FAILED|skipped)/.test(String(e.message))) break;
      }
      arms[arm][stalled ? "stall" : "ok"]++;
    });
  }
  return { rounds, arms };
}

/**
 * Every draw in the round, not just the scenario's four.
 *
 * Round 28 is why this exists. `does a rasterise poison the next draw` PASSED,
 * and the same round skipped `the chart is actually visible` on "PowerPoint did
 * not respond while drawing shapes 1-9 of 9 (45s)" with the trace adding "the
 * last thing the host answered was 'rasterising a slide', 0s earlier". A draw
 * stalling straight after a rasterise, in the round whose rasterise scenario had
 * just reported no effect — because the scenario counts only the four draws it
 * makes itself and that one was not one of them. The evidence was being thrown
 * away by the thing built to collect it.
 *
 * A round issues about forty draws. Counting all of them is a tenfold bigger
 * sample per round, which turns "30-50 more rounds" into a handful:
 *
 *   rounds 23, 26, 27, 28 — 157 draws, against 16 the arms would have counted
 *   after a rasterise      1 stalled /  36
 *   after anything else    0 stalled / 121
 *
 * REPORTED SEPARATELY FROM THE ARMS, AND WEAKER THAN THEM, WHICH IS THE WHOLE
 * POINT OF KEEPING BOTH. The arms are counterbalanced — interleaved so position
 * cannot account for a difference — and that is what makes four draws worth
 * anything. This population is observational: draws that follow a rasterise
 * follow it because of which scenario they belong to, and those scenarios differ
 * in shape count and in what they ask of the host. So it can raise a suspicion
 * and it cannot settle one. Two populations, honestly labelled, beat one
 * population quietly mixing the two kinds of evidence.
 */
/**
 * Rasterise labels from rounds archived BEFORE `op` existed.
 *
 * NOT a guess at wording — an ENUMERATION of the archive. A rasterise names
 * itself unambiguously in two places that do not depend on the label at all:
 * the success line `rasterised a slide` carries `label`, and the visibility
 * scenario's `visibility step` carries `what`. Reading every round through
 * those gives the complete set, and it is closed: old archives do not change.
 *
 * WHAT THIS RECOVERS TODAY: NOTHING, AND THAT IS MEASURED, NOT ASSUMED.
 *
 * The first version of this comment claimed these four labels were 35 of 43
 * labelled rasterises and that 81% of the population was missing from the
 * pooled answer. THAT WAS WRONG, and it was wrong in the way this repo keeps
 * being wrong: a number counted against the wrong denominator. These labels do
 * lack the string "rasteris" — but `isRasterise` tests the label AND THE
 * MESSAGE, and the message on a successful rasterise is `rasterised a slide`,
 * which matches. They were classified correctly all along.
 *
 * Pooled over all 90 archived rounds, with and without this set:
 *
 *     rasterise      ok 449, stall 1      (identical both ways)
 *     anything else  ok 3302, stall 1     (identical both ways)
 *
 * So this set is belt-and-braces, not a repair. It is kept because it makes the
 * classifier independent of a message string that nobody has promised to keep,
 * and because the equality above is now a fact on the record rather than an
 * assumption. If it ever starts changing a number, something upstream renamed a
 * trace message and that is worth knowing.
 *
 * The REAL breakage this pair found was in `chartIsVisible`, which matches
 * `lastStall.what` alone — no message to fall back on — and therefore genuinely
 * did stop firing. See round 113.
 */
const RASTERISE_LABELS_BEFORE_OP = new Set([
  "an end-of-round slide shot",
  "the visibility BEFORE render",
  "the visibility AFTER render",
  "the visibility CONTROL render (same slide, back to back)",
]);

export function poolEveryDraw(logs) {
  const isDraw = (e) =>
    (e.scope === "draw" && e.message === "batch issued") ||
    /^after a (?:rasterise|cheap read) #\d+: drawing$/.test(String(e.data?.what ?? ""));
  // A rasterise EVENT is never itself a draw. The scenario's own arm markers say
  // "after a rasterise #0: drawing" — that is a draw which FOLLOWS a rasterise,
  // and reading it as one would tar the next draw with a rasterise that had
  // already been accounted for. Caught by the test below rather than by reading:
  // the untagged-draw case came out one short and the miscount was this.
  // ON `op` FIRST, then the enumerated legacy labels, then the prose.
  //
  // THIS ONE WAS NOT BROKEN — checked, and the check is the point. The sibling
  // in `chartIsVisible` was broken by call sites being given individual names,
  // so this classifier was the obvious next casualty: it also identifies a
  // rasterise by matching prose. It survives only because it happens to test
  // the MESSAGE as well as the label, and the message `rasterised a slide`
  // still contains "rasteris".
  //
  // That is luck, not design. `op` makes it design. The pooled numbers are
  // identical before and after (see `RASTERISE_LABELS_BEFORE_OP`), which is
  // exactly what a belt-and-braces change should look like and is recorded so
  // nobody later mistakes this for a fix that moved something.
  const isRasterise = (e) =>
    !isDraw(e) &&
    (e.data?.op === "rasterise" ||
      RASTERISE_LABELS_BEFORE_OP.has(String(e.data?.what ?? "")) ||
      RASTERISE_LABELS_BEFORE_OP.has(String(e.data?.label ?? "")) ||
      /rasteris/i.test(`${String(e.data?.what ?? "")} ${String(e.message ?? "")}`));
  const isStall = (e) =>
    e.scope === "host" && e.message === "gave up waiting" && /drawing shapes/.test(String(e.data?.what ?? ""));

  const after = { rasterise: { ok: 0, stall: 0 }, "anything else": { ok: 0, stall: 0 } };
  let rounds = 0;
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    rounds++;
    const es = [...entries].sort((a, b) => (a.ms ?? 0) - (b.ms ?? 0));
    const at = [];
    es.forEach((e, i) => {
      if (isDraw(e)) at.push(i);
    });
    let prev = 0;
    at.forEach((i, k) => {
      // A rasterise anywhere since the PREVIOUS draw, so the classification is
      // about what the host did immediately before this draw and nothing older.
      let rasterised = false;
      for (let z = prev; z < i; z++) if (isRasterise(es[z])) rasterised = true;
      const end = k + 1 < at.length ? at[k + 1] : es.length;
      let stalled = false;
      for (let z = i + 1; z < end; z++) if (isStall(es[z])) stalled = true;
      after[rasterised ? "rasterise" : "anything else"][stalled ? "stall" : "ok"]++;
      prev = i;
    });
  }
  return { rounds, after };
}

/**
 * Judge a prediction against a round.
 *
 * A change made because of a round should say what the NEXT round will show, and
 * then be judged by it in as many words. That discipline is worth having and it
 * is worth nothing if nobody checks: #468 staked itself on three questions
 * answering for the first time in eleven rounds, round 27 said no, and the only
 * reason the failure was recorded rather than quietly forgotten is that someone
 * remembered making the claim.
 *
 * Deliberately few claim kinds. A prediction that cannot be evaluated by a
 * machine is a paragraph, and paragraphs are what this replaces — but the prose
 * still travels with it, because "it failed" is the least interesting half and
 * `because` is where the thinking is.
 */
export function judgePrediction(prediction, log) {
  const NOT_ASKED = NOT_ASKED_ANSWERS;
  // A THIRD CATEGORY, and keeping it out of NOT_ASKED is the point. A
  // CONDITIONAL probe's question only exists when the host misbehaves in a
  // particular way, and on a round where it behaved there is nothing to
  // measure — which is neither "the harness could not set it up" (ours to fix)
  // nor "the host would not answer" (a fact about the host).
  //
  // `does-a-failed-group-poison-the-tag` is why this exists. It answers
  // `no-refusal` — "the host grouped through a slide handle two syncs old, so
  // the refusal was never provoked", its own words — in 94 of the 114 rounds
  // that carry it, and `tags-gone` in the other 20. `no-refusal` fell through
  // to the pattern test, matched nothing, and was reported FAILED: the house
  // defect exactly, an unmeasured thing reported as a negative measurement.
  const UNPROVOKED = new Set(["no-refusal"]);
  const answers = log?.hostAnswers?.answers;
  const byId = new Map((Array.isArray(answers) ? answers : []).map((a) => [a.id, a]));
  const c = prediction.claim ?? {};
  const seen = (id) => byId.get(id);

  if (c.kind === "probe-answers" || c.kind === "probe-starves") {
    const want = c.kind === "probe-answers";
    const missing = (c.ids ?? []).filter((id) => !seen(id));
    if (missing.length) return { verdict: "undetermined", why: `not on this sheet: ${missing.join(", ")}` };
    const wrong = (c.ids ?? []).filter((id) => NOT_ASKED.has(seen(id).answer) === want);
    return wrong.length
      ? { verdict: "FAILED", why: `${wrong.map((id) => `${id}=${seen(id).answer}`).join(", ")}` }
      : { verdict: "held", why: (c.ids ?? []).map((id) => `${id}=${seen(id).answer}`).join(", ") };
  }
  if (c.kind === "trace-line-present") {
    // A CLAIM ABOUT WHAT THE TRACE SAYS, which the ledger could not express.
    // Its four kinds were all about scenarios and probes, so a question whose
    // whole answer is "does this line appear" had to live in prose — and prose
    // is what this ledger exists to replace.
    //
    // `insteadOf` is what makes it judgeable rather than merely true-or-silent:
    // a round where NEITHER line appears did not ask the question, and that is
    // `undetermined`, not a refutation. Without it, a round that simply never
    // reached the code would read as evidence against the claim.
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries) || !entries.length)
      return { verdict: "undetermined", why: "the round carries no trace to look in" };
    // BY SCOPE AS WELL AS BY MESSAGE, because a claim naming one message is only
    // as good as the number of messages the thing can produce. Round 102: the
    // deck-style probe has TWO failure lines, `spending-the-bad-first-call-cures-
    // the-read` named one, the round produced the other — and the ledger read
    // that as a cure. A false HELD from a malformed claim, which is the same
    // defect this file keeps finding elsewhere.
    //
    // `scope` claims the whole family: no entry from that scope at all.
    const saw = (m) => entries.some((e) => String(e.message) === m);
    // What actually matched, so the verdict names the right reason. Reporting
    // `still carries <message>` for a SCOPE match would be a report describing
    // evidence it did not use.
    const sawScope = () => (c.scope ? entries.find((e) => String(e.scope) === c.scope) : undefined);
    // A CLAIM THAT A SYMPTOM IS GONE, which is what a fix predicts and what this
    // kind could not say. `absent: true` inverts the reading: the named line
    // must NOT appear, and `insteadOf` — a line every round carries — is what
    // proves the round ran at all, so silence from a round that never started
    // stays `undetermined` rather than counting as a cure.
    if (c.absent) {
      if (saw(c.message)) return { verdict: "FAILED", why: `the trace still carries \`${c.message}\`` };
      const other = sawScope();
      if (other)
        return {
          verdict: "FAILED",
          why: `the \`${c.scope}\` scope is still there, as \`${String(other.message)}\``,
        };
      if (c.insteadOf && saw(c.insteadOf)) return { verdict: "held", why: `no \`${c.message}\` in a round that ran` };
      return { verdict: "undetermined", why: "the round did not run far enough to show the line is gone" };
    }
    if (saw(c.message)) return { verdict: "held", why: `the trace carries \`${c.message}\`` };
    if (c.insteadOf && saw(c.insteadOf))
      return { verdict: "FAILED", why: `the trace carries \`${c.insteadOf}\` instead` };
    return {
      verdict: "undetermined",
      why: c.insteadOf
        ? `neither line appeared — the round never reached the code that writes them`
        : `\`${c.message}\` did not appear, and no alternative was named to tell absence from silence`,
    };
  }
  if (c.kind === "scenario-passes") {
    const st = log?.selftest ?? {};
    const all = Object.keys(st).map((k) => st[k]);
    const named = (c.names ?? []).map((n) => [n, all.find((x) => x?.name === n)]);
    // A QUESTION THE HOST NEVER ANSWERED IS NOT A REFUTATION, and this branch
    // called both of those FAILED until 2026-08-19 — while `probe-answers` and
    // `probe-detail-matches` directly above already answered `undetermined` for
    // exactly the same situation, and this file's own doctrine says "a skip is
    // not a flip" and "a miss is not a failure" in as many words.
    //
    // Round 088 is what makes it live rather than theoretical: `insert onto a
    // slide that already has content` came back `skipped: true, ok: false`
    // because PowerPoint stopped answering mid-draw. Any prediction naming that
    // scenario would have been recorded FAILED — the strongest verdict the
    // ledger has — on a round where the host declined to be asked.
    const missing = named.filter(([, s]) => !s).map(([n]) => n);
    if (missing.length) return { verdict: "undetermined", why: `not on this sheet: ${missing.join(", ")}` };
    const skipped = named.filter(([, s]) => s.skipped).map(([n]) => n);
    if (skipped.length) return { verdict: "undetermined", why: `the host never ran: ${skipped.join(", ")}` };
    const wrong = named.filter(([, s]) => !s.ok).map(([n]) => n);
    return wrong.length ? { verdict: "FAILED", why: wrong.join(", ") } : { verdict: "held", why: "" };
  }
  if (c.kind === "probe-detail-matches") {
    const a = seen(c.id);
    if (!a) return { verdict: "undetermined", why: `${c.id} not on this sheet` };
    // A never-put question has no detail worth matching, and calling that a
    // failure would blame the prediction for the host refusing the question.
    if (NOT_ASKED.has(a.answer)) return { verdict: "undetermined", why: `${c.id} was never put (${a.answer})` };
    if (UNPROVOKED.has(a.answer))
      return { verdict: "undetermined", why: `${c.id}'s precondition did not occur this round (${a.answer})` };
    // THE ANSWER KEY AS WELL AS THE PROSE. An `answer` is a stable enum the
    // probe picks; a `detail` is a sentence someone rewords. #520 was staked on
    // `InvalidParam|5010|GeneralException` — error codes this probe has never
    // once emitted in 114 rounds — so on the 20 rounds where it DID fire and
    // report `.tags was undefined after the refused group`, the exact outcome
    // the prediction claimed, the ledger would still have printed FAILED.
    //
    // Round 102 taught the same lesson from the trace side: a claim is only as
    // good as the set of strings the thing it watches can produce. Matching the
    // key gives a claim something to anchor on that a rewrite cannot move.
    const hit = new RegExp(c.pattern, "i").test(`${a.answer ?? ""} ${a.detail ?? ""}`);
    return { verdict: hit ? "held" : "FAILED", why: `${a.answer} — ${String(a.detail ?? "").slice(0, 90)}` };
  }
  return { verdict: "undetermined", why: `unknown claim kind ${String(c.kind)}` };
}

/**
 * The round a prediction may be judged on: the newest one taken AFTER the build
 * it was staked on, or nothing.
 *
 * The "nothing" is the half that was wrong, and it produced a false HELD the
 * first time a prediction was staked on a build with no round yet. The old rule
 * read "find the prompting build in the archive and take everything after it",
 * and when that build was NOT in the archive — which is the normal state for a
 * prediction written the moment a change lands — it fell back to judging against
 * every round, i.e. the newest one, which is OLDER than the change. `same scale
 * across the deck` passes in the recent archive, so a prediction about a change
 * that has never been run came out `held` on evidence recorded before it
 * existed.
 *
 * That is the same defect `roundsToJudgeOn` already describes ("round 27
 * standing in for a round 29 that does not exist yet"), reached through the
 * not-found branch instead of through inequality. A prediction whose build has
 * not been rounded has NO round to judge on, and saying so is the whole answer.
 */
export function roundToJudgeOn(logs, afterBuild, buildOf, madeOn) {
  // THE NEWEST OF THE ELIGIBLE ROUNDS, and everything about WHICH rounds are
  // eligible now lives in `roundsToJudgeOn` with the code that decides it.
  // Prefer that one: taking the last of a population is exactly the n=1 reading
  // that let #520 sit open for 130 rounds, and it survives here only because
  // the single-round claim kinds have no population to pool.
  return roundsToJudgeOn(logs, afterBuild, buildOf, madeOn).slice(-1)[0];
}

/**
 * EVERY round a prediction may be judged on, newest last — the population
 * `roundToJudgeOn` takes the last of.
 *
 * A single round is n=1, and this file spends a page arguing that a single
 * round's verdict is a fact about the code AND the host's mood that afternoon.
 * The prediction ledger was the one reader that ignored it: `roundToJudgeOn`
 * hands back the newest round and the verdict is whatever that afternoon said.
 *
 * For a CONDITIONAL probe that is not a rounding error, it is the whole answer.
 * `does-a-failed-group-poison-the-tag` puts its question in 20 of 114 rounds,
 * so the newest round is 82% likely to carry no measurement at all — and the
 * one time in five it does, that reading stands alone against nineteen others
 * nobody looks at. #520 sat OPEN for 130 rounds with its answer recorded
 * twenty times.
 */
export function roundsToJudgeOn(logs, afterBuild, buildOf, madeOn) {
  // THE LAST ROUND ON THAT BUILD, NOT THE FIRST. `findIndex` stopped at the
  // earliest one, and the cycle archives two rounds at 16:9 plus one at 4:3 on
  // a single build — so `slice(madeAt + 1)` still held the prompting build's own
  // siblings, and the newest of them could be handed back as the round that
  // judges a change made after all three were taken. The discipline this loop is
  // built on is "run the same build twice"; the judge has to know that a build
  // can appear more than once or it quietly rules on its own control.
  let madeAt = -1;
  for (let i = 0; i < logs.length; i++) if (buildOf(logs[i]) === afterBuild) madeAt = i;
  // Rounds are ordered oldest-first, so anything at or before the prompting
  // round cannot test the change — and a build with no round at all is past the
  // end of the archive, not the start of it.
  //
  // WHICH LEAVES A PREDICTION STAKED ON A BUILD NOBODY EVER ROUNDS, and that is
  // the normal case rather than the exotic one: a claim is written the moment a
  // change lands, and the next merge supersedes that commit before any round
  // runs. Matching the build exactly would then answer `no round yet` forever —
  // an entry that can never be judged, which is the same as no entry at all.
  //
  // A round's build stamp carries its DATE (`95170cf · 2026-08-17 09:03Z`), so
  // an entry that says when it was made can be judged on any round taken after
  // that day. The build match stays first: it is exact, and dates are only as
  // good as the stamp.
  return madeAt !== -1 ? logs.slice(madeAt + 1) : madeOn ? logs.filter((l) => stampDate(l, buildOf) > madeOn) : [];
}

/**
 * The verdict across every eligible round, and the split that produced it.
 *
 * THE SPLIT IS THE POINT, not the headline. A probe that answers sometimes
 * gives a population, and a population can disagree with itself — which is a
 * finding about the host, not a tie to be broken quietly. So when the decided
 * rounds say both things, this says so rather than picking one.
 */
export function judgeAcross(prediction, rounds, buildOf) {
  const held = [];
  const failed = [];
  let undecided = 0;
  for (const r of rounds) {
    const { verdict, why } = judgePrediction(prediction, r);
    if (verdict === "held") held.push({ build: buildOf(r), why });
    else if (verdict === "FAILED") failed.push({ build: buildOf(r), why });
    else undecided++;
  }
  const decided = held.length + failed.length;
  const verdict = !decided ? "undetermined" : held.length && failed.length ? "BOTH" : held.length ? "held" : "FAILED";
  const last = (failed.length ? failed : held).slice(-1)[0];
  return { verdict, held: held.length, failed: failed.length, undecided, decided, rounds: rounds.length, last };
}

/** When a round was taken, out of its build stamp, or "" when it does not carry one. */
function stampDate(log, buildOf) {
  // `buildOf` yields the hash alone, so read the raw stamp — the date is the
  // half this needs and the half the hash throws away.
  void buildOf;
  // THE TIME AS WELL AS THE DAY, because the comparison above is strict `>` and
  // a prediction is nearly always judged by a round taken the SAME day it was
  // staked. Date-only, `"2026-08-19" > "2026-08-19"` is false, so the round that
  // exists is refused and the report says `no round yet` — which is precisely
  // what round 088 did on 2026-08-19 to the #586 entry staked that morning.
  //
  // Lexicographic order does the rest with no extra branch: a stamp is
  // `hash · YYYY-MM-DD HH:MMZ`, so `"2026-08-19 13:58Z" > "2026-08-19"` is true
  // (same prefix, longer string), while `"2026-08-18 07:11Z" > "2026-08-19"` is
  // still false. A `madeOn` carrying its own time compares by time; one carrying
  // only a day means start of that day, which is what staking on a day means.
  const m = /·\s*(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2})?)/.exec(String(log?.build ?? ""));
  return m ? m[1] : "";
}

/**
 * A scenario whose verdict counts "N of M" — and what M was in the rounds before.
 *
 * ROUND 088 IS WHY THIS EXISTS. `same scale across the deck` passed with
 * `6 of 6 charts carry the shared scale` where all 63 rounds before it had run
 * EIGHT. Its verdict is `scaled === charts.length`, measured against a
 * population it discovers rather than one it is given — `probeCharts(prefix)`
 * returns whatever the deck scan finds, and the only guard is `< 2`. So a run
 * that finds six charts and scales six reads exactly like one that found eight
 * and scaled eight, and `scenarioRegressions` compares PASS to PASS and says
 * "no scenario regressed".
 *
 * That is the same defect as the suite-size high-water mark: a guard that cannot
 * see its own population shrink. A scenario could fall to `2 of 2` and every
 * reading in this file would still be green.
 *
 * The denominator is not a fault on its own — round 088's six is downstream of a
 * host stall that skipped the scenario which seeds the probe charts. It is a
 * REASON TO READ, and quoting the pass without it is quoting a ratio whose
 * bottom half moved.
 */
export function poolScenarioPopulations(logs) {
  const seen = new Map();
  // THE INDEX, NOT JUST THE VALUE — see `isTheRoundBeingJudged`. Without it
  // there is nothing to distinguish "the round being judged counted 7" from
  // "some round counted 7, once, a month ago".
  logs.forEach((log, at) => {
    const st = log?.selftest ?? [];
    for (const s of Object.keys(st).map((k) => st[k])) {
      if (!s?.name) continue;
      // The verdict's own "N of M" — the shape every counting scenario here uses.
      const m = /\b(\d+) of (\d+)\b/.exec(String(s.detail ?? ""));
      if (!m) continue;
      if (!seen.has(s.name)) seen.set(s.name, []);
      seen.get(s.name).push({ build: String(log.build ?? ""), of: Number(m[2]), ok: Boolean(s.ok), at });
    }
  });
  const shrunk = [];
  for (const [name, hist] of seen) {
    // THREE PRIORS MINIMUM, because "usually" needs more than one observation.
    // Round 112 fired this on `insert onto a slide that already has content —
    // 2 this round, usually 16 over 1 prior round(s)`: a scenario whose verdict
    // only recently started carrying an "N of M" count, so its entire history
    // was a single round. One number is not a norm, and this project's own noise
    // floor — one build run twice scoring 1 and 5 — is the reason to say so.
    if (hist.length < MIN_PRIORS_FOR_A_BASELINE + 1) continue;
    const now = hist[hist.length - 1];
    /**
     * "THIS ROUND" HAS TO BE THIS ROUND, and for three rounds it was not.
     *
     * `hist` holds only the rounds that CARRIED a count, so its last entry is
     * the newest round that counted — which is not the round being judged when
     * the newest round carried none. Rounds 306, 307 and 308 each printed
     *
     *     insert onto a slide that already has content — 7 this round,
     *     usually 16 over 9 prior round(s)
     *
     * where the 7 is ROUND 282's, twenty-four rounds and four builds earlier,
     * and this round had counted nothing at all.
     *
     * It is permanent, not a blip: 282 is the newest round with a count and
     * always will be, so the line fires on every future round with the same
     * stale number. A warning that cries every round and names the wrong round
     * is worse than no warning — it teaches the reader to skip the line that
     * exists to make them stop and read.
     *
     * Worth knowing WHY that scenario went quiet, because it also condemns the
     * baseline: only 10 of 285 archived rounds ever carried a count for it, and
     * every one of them is a failure mode — "the host stopped answering during
     * this scenario", "the deck scan could not see the whole deck". The count
     * appears when the scenario does NOT do its job. So `usually 16` was never
     * the healthy population; it was the shape of an abort.
     *
     * If this round did not count, there is no ratio of this round's to warn
     * about. Say nothing.
     */
    if (!isTheRoundBeingJudged(now.at, logs)) continue;
    const priors = hist.slice(0, -1).map((h) => h.of);
    // The population it has USUALLY had. Not the mean: a single small round
    // would drag the bar down and hide the next one.
    const usual = priors.sort((a, b) => b - a)[Math.floor(priors.length / 2)];
    if (now.of < usual) shrunk.push({ name, now: now.of, usual, ok: now.ok, rounds: priors.length });
  }
  return shrunk;
}

/**
 * Whether the charts this round drew ended up GROUPED, which no verdict reports.
 *
 * ROUNDS 092 AND 093 ARE WHY. One build, run twice, nothing changed between them:
 *
 *     092   20 charts grouped, 0 refusals, deck 0,4,2,5,1,1,1
 *     093   15 charts grouped, 4 refusals, deck 0,4,2,17,24,24,24
 *
 * Three slides ended holding TWENTY-FOUR shapes each instead of one — three
 * charts that did not group, seventy-two shapes loose on the deck — and both
 * rounds reported `13/13` and the byte-identical verdict line `8 of 8 charts
 * carry the shared scale ... 8 still re-editable`.
 *
 * The scenario is not lying. It asks whether the config survived, and it did:
 * the ungrouped fallback keeps the tag. It simply cannot see grouping, which is
 * what the last several changes here have been about — so `scenarioRegressions`
 * compares PASS to PASS across a round that left seventy-two loose shapes and a
 * round that left none.
 *
 * Counted from the TRACE, which is exact — `charts` on each `grouped the chart's
 * shapes` line, and every `not grouping:` line — with the deck printed beside it
 * as corroboration rather than as the measure. A slide's shape count needs a
 * threshold to interpret, and a guard sized by guesswork is how this instrument
 * has been wrong before.
 */
/**
 * How many earlier rounds a "usually" needs before it is allowed to be printed.
 *
 * Two emitters have now shipped a baseline computed from ONE prior round, and
 * the gate printed `usually 16 over 1 prior round(s)` in a real run before
 * anyone noticed. This project's own noise floor — one build run twice, scoring
 * 1 and 5 with nothing changed — is the argument: a single prior cannot
 * distinguish a trend from the host's mood.
 */
const MIN_PRIORS_FOR_A_BASELINE = 3;

/**
 * DID THE ROUND BEING JUDGED ACTUALLY CONTRIBUTE, or is `now` an older one?
 *
 * Three emitters here build a history by walking `logs` and SKIPPING the rounds
 * that carry nothing to measure, then take the last entry as "this round". That
 * is only true while the newest round contributed. When it did not, the last
 * entry is the newest round that DID — which can be any distance back.
 *
 * Found on 2026-08-29 in `poolScenarioPopulations`, which told three consecutive
 * rounds `insert onto a slide that already has content — 7 this round, usually
 * 16` where the 7 belonged to round 282, twenty-four rounds and four builds
 * earlier, and none of the three had counted anything at all. It was permanent
 * rather than a blip: 282 was the newest round with a count and would have
 * stayed so, so the line would have fired for ever with the same stale number.
 *
 * Each site records the index it came from and asks this before reporting. The
 * honest answer when the newest round contributed nothing is silence: there is
 * no figure OF THIS ROUND to compare against its own history.
 */
function isTheRoundBeingJudged(at, logs) {
  return at === (logs?.length ?? 0) - 1;
}

export function poolGroupingOutcome(logs) {
  const per = [];
  logs.forEach((log, at) => {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) return;
    let grouped = 0;
    let refused = 0;
    // THE THIRD OUTCOME, and it was dropped from the numerator AND the
    // denominator. Grouping succeeds (`grouped the chart's shapes`), declines by
    // rule (`not grouping: …`), or THROWS — `grouping the chart's shapes`
    // carrying an `error`, usually `InvalidArgument`. Only the first two were
    // counted, so a round where a group threw reported `8 of 8 attempt(s)
    // grouped, 0 refused` and read as perfect. **The missing attempt IS the
    // defect**, and the count that was supposed to expose it hid it.
    //
    // 183 thrown groups across 65 of 150 rounds. They match the loose-shape
    // slides exactly: rounds 159, 161, 167 and 174 each threw once and ended
    // with a slide holding 17, 17, 11 and 11 shapes; the twelve rounds between
    // them threw none and ended at 5. A chart whose group throws is left as its
    // shapes, which is what `poolFullestSlide` sees from the other side.
    //
    // Found one round after `poolFullestSlide` landed, by that instrument, in
    // this same function — which had been counting two of three outcomes since
    // it was written.
    let threw = 0;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m === "grouped the chart's shapes") grouped += Number(e.data?.charts) || 0;
      else if (/^not grouping/.test(m)) refused += 1;
      else if (m === "grouping the chart's shapes" && e.data?.error) threw += 1;
    }
    if (!grouped && !refused && !threw) return;
    const deck = (log?.deck?.inventory ?? []).map((sl) => sl.count ?? sl.shapes?.length ?? 0);
    per.push({ build: String(log.build ?? "").split(" ")[0], grouped, refused, threw, deck, at });
  });
  if (!per.length) return null;
  const now = per[per.length - 1];
  // The stale-`now` guard — see `isTheRoundBeingJudged`. This one prints the
  // headline grouping figure of every gate run, and the deck line beside it, so
  // a stale `now` here describes another round's deck as this round's.
  if (!isTheRoundBeingJudged(now.at, logs)) return null;
  const priors = per.slice(0, -1);
  // THREE PRIORS MINIMUM — the same rule `poolScenarioPopulations` needed, and
  // the same defect it had. A median of one observation is not a "usually", and
  // a median of ZERO observations used to be reported here as `0`, which is this
  // project's house defect exactly: UNREADABLE PRINTED AS A NEGATIVE. A reader
  // seeing "usually 0 refused" cannot tell a clean history from no history.
  // `null` is the honest value and the printer must say so out loud.
  const refusedMedian =
    priors.length >= MIN_PRIORS_FOR_A_BASELINE
      ? priors.map((p) => p.refused).sort((a, b) => a - b)[Math.floor(priors.length / 2)]
      : null;
  // THE DENOMINATOR, AND THE ONLY READING THAT SURVIVES A CHANGE OF POPULATION.
  //
  // `grouped` per round ran 15-20 for the whole archive and then halved to 9 at
  // round 153, and nothing here could say why — or that it had happened. The
  // cause is benign and complete: the in-place update started working. Six more
  // charts per round are updated in place, an in-place update never redraws, and
  // a chart that is not redrawn is never regrouped. 15 - 9 = 6 and 11 - 5 = 6,
  // exactly, in the round it changed.
  //
  // Benign, and it silently rebased every grouping figure in this file. "0
  // refused (usually 2)" reads as an improvement when half the attempts stopped
  // happening — fewer refusals out of fewer tries. The RATE is the reading that
  // holds across the change: round 160 grouped 9 of 9 attempts, round 141
  // grouped 10 of 19. Same instrument, and only one of those two numbers can be
  // compared to the other.
  //
  // This is the `same scale across the deck` trap again — "6 of 6" and "8 of 8"
  // are both a pass — and this file argues it at length one screen up while
  // reporting a bare count here.
  const attempts = now.grouped + now.refused + (now.threw ?? 0);
  const recent = per.slice(-RECENT_IN_A_ROW).map((p) => p.grouped + p.refused + (p.threw ?? 0));
  return { now, refusedMedian, rounds: priors.length, attempts, recent };
}

/**
 * The host-cost meter every scenario verdict has carried since round 023, which
 * no script has ever read.
 *
 * Four counters — `errors`, `idRefusals`, `generalExceptions`, `emptyReReads` —
 * are recorded per scenario, per round, as a delta across that scenario. 86 of
 * 86 archived rounds carry them. `grep friction scripts/` returned nothing.
 *
 * It answers per scenario what `poolScenarioPopulations` can only ask per round:
 * whether a scenario passed on an easier host than it used to.
 *
 * TWO COUNTERS ARE NOT SIGNALS AND THE REPORT MUST SAY SO, or this becomes one
 * more number read as meaning something:
 *
 *   - `generalExceptions` has never been non-zero. Not once, in any scenario, in
 *     any round.
 *   - `stop a run part-way` reports exactly one `error` every round — the
 *     deliberate abort, counted as an error. A constant is not a measurement.
 *
 * Both are DERIVED here rather than written down, because a hardcoded conclusion
 * keeps printing after it stops being true — which is exactly how the deck-style
 * probe lied for three rounds.
 */
export function poolScenarioFriction(logs) {
  /**
   * THE KEY LIST IS READ OFF THE DATA, because the docstring above says so and
   * the code did the opposite.
   *
   * `KEYS` was a literal naming four counters. The friction object carries
   * EIGHT — `errors`, `idRefusals`, `generalExceptions`, `emptyReReads`,
   * `reReadsRepaired`, `shortReReads`, `unmatchedReReads`, `settledByBinding` —
   * enumerated over all 4,285 friction records in the archive. Four never
   * reached a reader anywhere, and one of them is not idle: `reReadsRepaired` is
   * non-zero in 331 of those records. A hardcoded list goes stale silently,
   * which is the reason given three lines above for deriving everything else.
   *
   * A union over what is actually present, so a counter added tomorrow is pooled
   * the day it first appears rather than the day someone remembers.
   */
  const KEYS = [];
  const seenKeys = new Set();
  for (const log of logs ?? [])
    for (const e of log?.trace?.entries ?? [])
      for (const k of Object.keys(e?.data?.friction ?? {}))
        if (!seenKeys.has(k)) {
          seenKeys.add(k);
          KEYS.push(k);
        }
  const per = new Map();
  let rounds = 0;
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    let any = false;
    for (const e of entries) {
      const f = e?.data?.friction;
      const name = e?.data?.name;
      if (!f || !name) continue;
      any = true;
      if (!per.has(name)) per.set(name, { name, n: 0, seen: new Map(), sum: {} });
      const row = per.get(name);
      row.n += 1;
      for (const k of KEYS) {
        const v = Number(f[k]) || 0;
        row.sum[k] = (row.sum[k] ?? 0) + v;
        // Every distinct per-round value, so a CONSTANT can be told from a
        // number that merely happens to be large.
        if (!row.seen.has(k)) row.seen.set(k, new Set());
        row.seen.get(k).add(v);
      }
    }
    if (any) rounds += 1;
  }
  // Ranked on the two counters the report leads with. Defaulted because the key
  // list is derived now: a set of logs carrying no `errors` at all must sort,
  // not produce NaN and a scrambled order.
  const lead = (r) => (r.sum.errors ?? 0) + (r.sum.idRefusals ?? 0);
  const rows = [...per.values()].sort((a, b) => lead(b) - lead(a));
  // A counter that has never once been non-zero anywhere carries no information.
  const dead = KEYS.filter((k) => rows.every((r) => r.sum[k] === 0));
  // A scenario/counter pair with exactly one distinct value across every round
  // is a constant — true of the deliberate abort, and worth naming as such.
  const constant = [];
  for (const r of rows)
    for (const k of KEYS) {
      const vals = r.seen.get(k);
      if (r.n > 2 && vals && vals.size === 1 && [...vals][0] !== 0)
        constant.push({ name: r.name, key: k, value: [...vals][0] });
    }
  return { rounds, rows, dead, constant };
}

/** Open predictions, judged against the newest round given. */
/**
 * Is this entry stamped with a day and no time, on a claim that has no build in
 * the archive to anchor it?
 *
 * `stampDate` compares lexicographically and a bare `2026-08-22` means the
 * START of that day, so a claim staked at 22:55 is judged against every round
 * taken earlier the same day — including the rounds that motivated it. The
 * first rule this ledger is built on is that a prediction is never judged
 * against the round that prompted it, and a date-only stamp walks around it.
 *
 * ONLY WHEN THE BUILD IS ABSENT. A `afterBuild` that IS in the archive pins the
 * position exactly and the date is never consulted, so warning there would be
 * noise on every correctly-staked entry.
 */
export function stakedWithoutATime(prediction, logs, buildOf) {
  const madeOn = String(prediction?.madeOn ?? "");
  // A DAY WITH NO TIME, or a time with no ZONE. Both make the comparison mean
  // something other than what the writer meant, and both fail quietly.
  //
  // The zone half was found immediately after the day half was fixed, by
  // stepping straight into it. #684 was re-stamped `2026-08-22 23:00` — local
  // time — while every round stamp is UTC (`c7e2876 · 2026-08-22 21:14Z`).
  // Lexicographically "21:14Z" < "23:00", so the two rounds that were taken to
  // judge it read as older than the claim and the entry said `no round yet`.
  // That is the worst direction to fail in: `no round yet` reads as patience.
  const bare = /^\d{4}-\d{2}-\d{2}$/.test(madeOn);
  const zoneless = /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}$/.test(madeOn);
  if (!bare && !zoneless) return false;
  // Only when the build is absent: a build that IS in the archive pins the
  // position exactly and the date is never consulted.
  return !(logs ?? []).some((l) => buildOf(l) === prediction.afterBuild);
}

function reportPredictions(logs, load = readFileSync) {
  let ledger;
  try {
    ledger = JSON.parse(load("rounds/predictions.json", "utf8"));
  } catch {
    return; // no ledger, nothing to say
  }
  const open = ledger.filter((p) => p.outcome === "open");
  if (!open.length || !logs.length) return;
  const buildOf = (log) => String(log?.build ?? "").split(" ")[0];
  console.log(`\n  OPEN PREDICTIONS`);
  for (const p of open) {
    // NEVER judged against the round that prompted it. A prediction is made
    // BECAUSE of a round, and the change it predicts about is not in that
    // round's build — so judging there fails every prediction the moment it is
    // written. #472's came out "FAILED / value=undefined" against round 28,
    // which is the round that raised the question and could not carry the
    // instrument that answers it.
    // By POSITION in the archive, not by "any build that is not this one".
    // Filtering on inequality picked the newest round that merely differed,
    // which for a prediction made on the newest round meant judging it against
    // an OLDER one — round 27 standing in for a round 29 that does not exist
    // yet. Rounds are ordered oldest-first, so anything at or before the
    // prompting round cannot test the change.
    // A DATE WITHOUT A TIME MEANS THE START OF THAT DAY, and that is a foot-gun
    // rather than a convenience. `stampDate` compares lexicographically, so a
    // claim staked at 22:55 with a bare `2026-08-22` is judged against every
    // round taken earlier the same day — INCLUDING the rounds that motivated it.
    // The first rule this ledger is built on is that a prediction is never
    // judged against the round that prompted it, and a date-only stamp walks
    // straight around it. #684 did exactly that on the evening it was written
    // and came out `BOTH`, on evidence recorded before the change existed.
    //
    // Warned rather than refused: the entries from 2026-08-16 onward carry
    // date-only stamps and are correctly judged, because the rounds that would
    // confuse them were taken on other days.
    if (stakedWithoutATime(p, logs, buildOf))
      console.log(
        `    ^ NOTE  ${p.id} is stamped \`${p.madeOn}\`, which the judge cannot compare safely.\n` +
          `            A day with no time is read from the START of that day, so rounds taken earlier —\n` +
          `            including the ones that prompted the claim — count as evidence. A time with no\n` +
          `            zone is compared against UTC round stamps, so a local-time stamp silently\n` +
          `            back-dates or forward-dates itself. Stake as \`YYYY-MM-DD HH:MMZ\`, in UTC.`,
      );
    const eligible = roundsToJudgeOn(logs, p.afterBuild, buildOf, p.madeOn);
    if (!eligible.length) {
      console.log(`    no round yet   ${p.id}  (${p.madeIn}, made on ${p.afterBuild})`);
      continue;
    }
    // ACROSS THE POPULATION, not on whichever round happened to be newest.
    const t = judgeAcross(p, eligible, buildOf);
    const where = t.decided
      ? `judged on ${t.rounds} round(s), ${t.decided} of which measured it`
      : `no round measured it`;
    console.log(`    ${t.verdict.padEnd(13)} ${p.id}  (${p.madeIn}) — ${where}`);
    if (t.decided) console.log(`                  held ${t.held} · FAILED ${t.failed} · never put ${t.undecided}`);
    if (t.last) console.log(`                  latest reading, ${t.last.build}: ${t.last.why}`);
    if (t.verdict === "BOTH")
      console.log(
        `                  the archive says BOTH about code that did not change. Read the split, not a headline.`,
      );
  }
  console.log(
    `    A prediction that came out is only half of it — record what happened in\n` +
      `    rounds/predictions.json and say so in the change that acts on it.`,
  );
}

/**
 * What each self-test scenario has said, round by round.
 *
 * A single round's verdict is not a fact about the code — it is a fact about the
 * code AND the host's mood that afternoon, and the two are not separable from
 * one round. `explode a degraded picture` FAILED in round 23, PASSED in 26,
 * FAILED in 27 and 28, with no change to it in between. Read one of those as a
 * regression and you go looking for a bug that is not there; read the next as a
 * fix and you close a question that is still open. Both happened here before
 * this existed, and the only reason it was caught is that someone laid four
 * rounds side by side by hand.
 *
 * A SKIP IS NOT A FLIP, and keeping the distinction is the point. `the chart is
 * actually visible` reads `pass pass pass skip` — it has never once disagreed
 * with itself; the host simply stopped answering during the fourth. A scenario
 * that has only ever passed-or-skipped is sometimes-unmeasured, which is a
 * completely different thing from one that has genuinely said both pass and fail
 * about the same code.
 */
export function scenarioHistory(logs) {
  const hist = new Map();
  for (const log of logs) {
    const st = log?.selftest;
    if (!st) continue;
    for (const key of Object.keys(st)) {
      const s = st[key];
      if (!s?.name) continue;
      if (!hist.has(s.name)) hist.set(s.name, []);
      hist.get(s.name).push(s.skipped ? "skip" : s.ok ? "pass" : "FAIL");
    }
  }
  return [...hist.entries()].map(([name, verdicts]) => {
    const measured = verdicts.filter((v) => v !== "skip");
    return {
      name,
      verdicts,
      skips: verdicts.length - measured.length,
      // Disagreement about the same code, which is the thing worth flagging.
      // Computed over MEASURED rounds only, so a skip can never manufacture one.
      flips: new Set(measured).size > 1,
    };
  });
}

/** Scenario verdicts across rounds, so one round is not mistaken for a trend. */
function reportStability(logs) {
  if (logs.length < 2) return;
  const hist = scenarioHistory(logs);
  if (!hist.length) return;
  const flipping = hist.filter((h) => h.flips);
  const skipped = hist.filter((h) => !h.flips && h.skips);
  console.log(`\n  SCENARIO VERDICTS ACROSS ${logs.length} ROUNDS`);
  for (const h of hist) {
    const mark = h.flips ? "FLIPS" : h.skips ? "skips" : "     ";
    console.log(`    ${mark} ${h.name.padEnd(46)} ${h.verdicts.join(" ")}`);
  }
  if (flipping.length)
    console.log(
      `\n    ${flipping.length} scenario(s) have said BOTH pass and fail about code that did not\n` +
        `    change between them. Do not read a single round's verdict on those as a\n` +
        `    regression or as a fix — they are reporting the host's mood.`,
    );
  if (skipped.length)
    console.log(
      `    ${skipped.length} more were SKIPPED in some rounds: never a disagreement, just\n` +
        `    rounds where the host stopped answering before the question was put.`,
    );
}

/**
 * The tag-failure faults, per BUILD, across every round.
 *
 * WHY THIS EXISTS, and it is the only reason: on 2026-08-15 a fix to the tag
 * anchor was nearly reported as working because round 043 (with it) looked like
 * round 042 (without it). It did — and so did rounds 041 and 042, which have NO
 * renderer change between them at all:
 *
 *     041  ca866e3   tags-undefined 1   group-5010 1   no-queue 1   tagging-failed 4
 *     042  a54401c   tags-undefined 5   group-5010 5   no-queue 5   tagging-failed 8
 *
 * A five-fold swing across a merge of a probe and some documentation. These
 * counts track the host's REGIME, not the code, so comparing two rounds by eye
 * is measuring mood — the same trap the rasterise arms exist to avoid, in a
 * place nobody had noticed it.
 *
 * Grouped by build rather than by round so repeat rounds on one build pool
 * instead of competing, and the spread WITHIN a build is printed, because that
 * spread is the noise floor any claim about a fix has to clear.
 */
export function poolTagFaults(logs) {
  const KINDS = [
    [/Cannot read properties of undefined \(reading 'add'\)/, "tags-undefined"],
    [/at=writing the chart's config tag/, "cfg-tag-5010"],
    [/at=grouping the chart's shapes/, "group-5010"],
    [/no chart's tag could be queued/, "no-queue"],
  ];
  const byBuild = new Map();
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    const build = String(log.build ?? "?").split(" ")[0];
    const counts = Object.fromEntries(KINDS.map(([, name]) => [name, 0]));
    counts["tagging-failed"] = 0;
    for (const e of entries) {
      if (/^tagging failed/.test(String(e.message ?? ""))) counts["tagging-failed"]++;
      const blob = JSON.stringify(e.data ?? {});
      for (const [re, name] of KINDS) if (re.test(blob)) counts[name]++;
    }
    if (!byBuild.has(build)) byBuild.set(build, []);
    byBuild.get(build).push(counts);
  }
  return byBuild;
}

/**
 * Does a chart that GETS GROUPED keep its config?
 *
 * THE QUESTION NOBODY ASKED FOR ELEVEN ROUNDS, and the archive had the answer
 * the whole time. Pooled over every round on 2026-08-15:
 *
 *     grouped      64 charts, 1 lost its tag    (1.6%)
 *     NOT grouped  62 charts, 41 lost their tag (66%)
 *
 * Per round it is almost mechanical — three grouped and none lost, two or three
 * ungrouped and two lost, round after round.
 *
 * It reframes the whole tag effort. When grouping succeeds the tag goes onto the
 * GROUP, a handle made in that batch and never resolved, and it lands. When
 * grouping is skipped the tag falls back to a `created` handle and is refused
 * two times in three. Which handle the fallback uses — the question four rounds
 * and a renderer change went into — is a question about the losing path.
 * `not grouping: no member handle this host will accept` carries `refreshed: 0`,
 * so what decides a chart's config is whether the pre-grouping RE-READ returned
 * anything.
 *
 * Reported every round from now on, because the cost of not reporting it was
 * eleven rounds of looking one level too low.
 */
/**
 * How much of the round's grouping this pool can actually SEE.
 *
 * `poolGroupVsTag` joins a chart's messages by `data.chart`, and that field is
 * written in exactly one place: the `traceAbout({ chart })` inside the deck-wide
 * rescale in `src/taskpane/selftest.ts`. Every other grouping in a round — the
 * ordinary inserts, the probe, anything a single-chart batch does — carries no
 * such key and is invisible here.
 *
 * MEASURED ARCHIVE-WIDE, the blindness is uneven and that is what makes it
 * dangerous rather than merely partial: across 57 rounds the pool sees 229 of
 * 638 `grouped the chart` events (36%) but 109 of 127 `not grouping` events
 * (86%). The two columns of a comparison are sampled at wildly different rates,
 * so the RATIO between them is biased, not just small.
 *
 * This is the shape of the 333/333 incident, in which a pooled figure was
 * reported as 100% because the window collecting it silently dropped every
 * declining case. The remedy is the same: print the denominator the reader would
 * otherwise assume.
 *
 * AND WIDENING IT IS NOT AVAILABLE, which is worth writing down so the next
 * reader does not spend the evening finding out again. Closing the gap would
 * mean joining each `grouped the chart` with the `tagging failed` for the SAME
 * chart, and that second line is emitted per BATCH on purpose — its own comment
 * says "the batch covers several charts, so one id would be a guess about
 * which", because a tag write fails for a whole sync rather than for one chart
 * in it. Attributing a batch failure to individual charts would manufacture
 * precision the host never gave, which is a worse fault than a narrow number
 * honestly labelled.
 *
 * Checked, not assumed: `data.charts` is 1 on every one of these entries across
 * the archive, so entries and charts do coincide here and the fractions below
 * compare like with like.
 */
export function poolGroupVsTagCoverage(logs) {
  const out = { groupedSeen: 0, groupedTotal: 0, ungroupedSeen: 0, ungroupedTotal: 0 };
  for (const log of logs) {
    for (const e of log?.trace?.entries ?? []) {
      const m = String(e.message ?? "");
      const keyed = (e.data ?? {}).chart !== undefined;
      if (m.startsWith("grouped the chart")) {
        out.groupedTotal++;
        if (keyed) out.groupedSeen++;
      } else if (m.startsWith("not grouping")) {
        out.ungroupedTotal++;
        if (keyed) out.ungroupedSeen++;
      }
    }
  }
  return out;
}

export function poolGroupVsTag(logs) {
  const out = { grouped: 0, groupedLost: 0, ungrouped: 0, ungroupedLost: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    const per = new Map();
    for (const e of entries) {
      const chart = (e.data ?? {}).chart;
      if (!chart) continue;
      if (!per.has(chart)) per.set(chart, []);
      per.get(chart).push(String(e.message ?? ""));
    }
    for (const msgs of per.values()) {
      const lost = msgs.some((m) => m.startsWith("tagging failed"));
      // A chart can only be counted once, and `grouped` wins: the two messages
      // are mutually exclusive per chart by construction, but a future retry
      // that produced both would otherwise be counted in both columns and quietly
      // flatten the very difference this exists to show.
      if (msgs.some((m) => m.startsWith("grouped the chart"))) {
        out.grouped++;
        if (lost) out.groupedLost++;
      } else if (msgs.some((m) => m.startsWith("not grouping"))) {
        out.ungrouped++;
        if (lost) out.ungroupedLost++;
      }
    }
  }
  return out;
}

/**
 * Did the chart land on a slide that already had shapes, or on a fresh one?
 *
 * **SUPERSEDED 2026-08-25: that 1% is now 36 of 36.** The measurement below is
 * — see `scripts/claims.mjs`, claim `fresh-slides-group`, re-checked every round.
 * kept because it is why the code exists, not because it describes the host
 * today — see `docs/WHAT-WE-KNOW.md`, checked every round by `scripts/claims.mjs`.
 *
 * THE ROOT, found 2026-08-15, and the cleanest separation this project has:
 *
 *     slide already had shapes  82 chart(s), 81 grouped = 99%
 *     freshly added, empty      74 chart(s),  1 grouped =  1%
 *
 * Not a tendency — a switch. And it completes the chain: a chart on a freshly
 * added slide gets a short or empty pre-grouping re-read, so it is not grouped,
 * so its tag falls back to a `created` handle, which this host refuses about
 * seven times in ten. Charts on an established slide group and keep their config.
 *
 * It is not a new problem either. It is THE problem, one level below everything
 * the tag work was aimed at, and this repo already knew a freshly-added slide is
 * special: `shape-add-held-slide-proxy` answers `threw`, its id does not
 * round-trip, and the #108-#111 saga was four attempts at drawing on one.
 *
 * `onSlide` is the shape count the DRAW recorded for the slide before it began,
 * which is why this pools over rounds that were archived long before anyone
 * thought to ask.
 */
export function poolFreshVsEstablished(logs) {
  const out = { established: 0, establishedGrouped: 0, fresh: 0, freshGrouped: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    const onSlide = new Map();
    const grouped = new Set();
    const decided = new Set();
    for (const e of entries) {
      const d = e.data ?? {};
      if (!d.chart) continue;
      // The FIRST batch's reading: what was on the slide before this chart.
      if (typeof d.onSlide === "number" && !onSlide.has(d.chart)) onSlide.set(d.chart, d.onSlide);
      if (/^grouped the chart/.test(String(e.message))) {
        grouped.add(d.chart);
        decided.add(d.chart);
      }
      if (/^not grouping/.test(String(e.message))) decided.add(d.chart);
    }
    for (const [chart, n] of onSlide) {
      // Only charts whose grouping was DECIDED. A chart the round never reached
      // has no outcome to attribute to its slide.
      if (!decided.has(chart)) continue;
      if (n > 0) {
        out.established++;
        if (grouped.has(chart)) out.establishedGrouped++;
      } else {
        out.fresh++;
        if (grouped.has(chart)) out.freshGrouped++;
      }
    }
  }
  return out;
}

/**
 * How long a window counts as "now".
 *
 * A GUESS ABOUT WHERE THE LAST REGIME BOUNDARY IS, and it should be read as
 * one. 20 was chosen to exclude everything before the settled retry (rounds
 * 064/065) — and it does, while sitting straight across a second boundary
 * nobody had found: rounds 137-142 failed to group 1,4,4,5,9,2 charts and
 * rounds 143-160 failed 3 in eighteen. So the "current rate" this produced was
 * itself smeared by a dead era, which is the exact defect it was written to
 * fix, one level down and in my own hand.
 *
 * Hence `recentSequence` below. A window is a guess; a sequence is evidence.
 */
export const RECENT_ROUNDS = 20;

/**
 * The fresh-versus-established split over recent rounds only.
 *
 * WHY IT EXISTS. The all-time number is known to be unreadable and the report
 * said so in prose and printed it anyway: "pooled over 39 rounds that predate
 * the retry, so it will climb slowly and should not be read as the current
 * rate". A figure a reader is told to mentally discount is a figure nobody can
 * use, and the answer to a stale window is not a caveat — it is a second
 * window.
 *
 * READ IT WITH `recentFreshSequence`, NEVER ALONE. Over successive windows this
 * same population reads 80% at 30 rounds, 91% at 20, 100% at 12 — and at 8
 * there are no fresh-slide charts in it at all, because the in-place update now
 * handles those charts and a chart that is not redrawn never lands on a fresh
 * slide. A percentage over an emptying population is the thing to watch for
 * here, and only the sequence shows it emptying.
 *
 * `null` below the window rather than a short-window figure, because a rate
 * from five rounds is the thing this exists to stop.
 */
export function recentFreshVsEstablished(logs, n = RECENT_ROUNDS) {
  return (logs?.length ?? 0) > n ? poolFreshVsEstablished(logs.slice(-n)) : null;
}

/**
 * Fresh-slide charts per round, and how many of them grouped — in sequence.
 *
 * The reading no window can give. A rate needs a population, and this one is
 * shrinking for a reason that has nothing to do with grouping: the in-place
 * update took the charts. `0/0` rounds are the signal, not a gap.
 */
export function recentFreshSequence(logs, n = RECENT_IN_A_ROW) {
  return (logs ?? []).slice(-n).map((log) => {
    const f = poolFreshVsEstablished([log]);
    return `${f.freshGrouped}/${f.fresh}`;
  });
}

/**
 * Charts that cannot follow a drag — the failure a passing scenario was hiding.
 *
 * WHY IT NEEDED ITS OWN NUMBER. `an update follows a moved chart` passes, and it
 * tests ONE chart. Rounds 073 and 074 lost the origin tag on **9 of 19 and 8 of
 * 17** charts in the same rounds — roughly half the population, every one of
 * which would fail to follow a user's drag, while the scenario reported the
 * round trip holding.
 *
 * That is the shape of thing this project keeps finding: a green verdict over a
 * sample, with the population telling a different story. `does a rasterise
 * poison the next draw` counted only its own four draws; the fresh-slide split
 * sat unqueried for eleven rounds. A scenario samples; a pooled count does not.
 *
 * NOT A RATE, deliberately. Only the FAILURES are traced — a successful origin
 * write says nothing — so there is no honest denominator here, and inventing one
 * by guessing at the chart count would be the kind of number this file has
 * already had to correct once. A count that climbed from 0 to 8-9 a round is the
 * signal; anyone wanting the ratio should read `grouped the chart's shapes`
 * beside it and say so out loud.
 */
export function poolOriginTagLosses(logs) {
  const out = { rounds: 0, charts: 0, worst: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    let here = 0;
    for (const e of entries) {
      if (!/^origin tag lost/.test(String(e.message))) continue;
      here += Number(e.data?.charts) || 0;
    }
    if (!here) continue;
    out.rounds++;
    out.charts += here;
    out.worst = Math.max(out.worst, here);
  }
  return out;
}

/**
 * A counter that did not move, said at the scope it was actually measured at.
 *
 * THIS LINE USED TO SAY "in any scenario in any round", from whatever logs it
 * was handed. Run over the archive that is nearly true and harmless. Run on ONE
 * round — `node scripts/triage.mjs rounds/207-6314a58.json`, which is how a
 * fresh round is read — it asserted a fact about 184 rounds from n=1.
 *
 * It was wrong. `emptyReReads` is non-zero in 72 of 184 archived rounds, 136
 * events; round 207 simply had none. The counter is gated behind the settle
 * retry running out (see `hostFriction.emptyReReads`), so a round where every
 * retry rescued its chart reports a TRUE zero — the thing the blind-gauge
 * warning elsewhere in this file exists to distinguish, printed here as though
 * the distinction had been made.
 *
 * A single round cannot tell a dead counter from a quiet one, so it no longer
 * claims to. Below two rounds it reports silence and names the alternative
 * rather than picking one.
 */
export function deadCounterNote(dead, rounds) {
  const names = dead.join(", ");
  if (rounds < 2)
    return `${names} did not move in this round — which a single round cannot tell from a counter that never moves.`;
  return (
    `${names} was never non-zero in the ${rounds} round(s) pooled here — either a dead gauge or a fault ` +
    `that did not occur. Pool more rounds before calling it either.`
  );
}

/** What each scenario cost the host — the meter nothing read for 63 rounds. */
function reportScenarioFriction(logs) {
  const o = poolScenarioFriction(logs);
  if (!o.rounds) return;
  console.log(`
  WHAT EACH SCENARIO COST THE HOST — pooled over ${o.rounds} round(s)`);
  console.log("    scenario                                    n    errors  idRefusals  emptyReReads");
  for (const r of o.rows.slice(0, 8))
    console.log(
      "    " +
        r.name.slice(0, 40).padEnd(42) +
        String(r.n).padStart(3) +
        String(r.sum.errors).padStart(9) +
        String(r.sum.idRefusals).padStart(12) +
        String(r.sum.emptyReReads).padStart(14),
    );
  // DERIVED, NOT DECLARED. A counter that has never moved and a counter that
  // never varies are both worthless, and printing them beside real numbers is
  // how a reader comes to trust one of them.
  if (o.dead.length) console.log("    " + deadCounterNote(o.dead, o.rounds));
  // NAME THE CONSTANT, DO NOT EXPLAIN IT. This line used to append "(its
  // deliberate abort)" to whatever it found, and it finds two things: `stop a
  // run part-way`, which does abort deliberately, and `explode a degraded
  // picture`, which does not — its `generalExceptions=1` is a real host throw.
  // So the report attached a false explanation to a genuine exception and told
  // the reader it was nothing, every round, for as long as anyone has looked.
  // A constant is worth naming; its cause belongs to whoever knows the scenario.
  for (const c of o.constant)
    console.log(`    \`${c.name}\` reports ${c.key}=${c.value} EVERY round — a constant, not a signal.`);
  console.log('    This is the per-scenario answer to "did it pass on an easier host", and it has');
  console.log("    been recorded since round 023 without anything reading it.");
}

/** Charts that would not follow a drag, which no scenario counts. */
function reportUpdateShortfalls(logs) {
  const o = poolUpdateShortfalls(logs);
  if (!o.updates) return;
  console.log(`
  WHAT AN UPDATE LEFT ON THE SLIDE — pooled over ${logs.length} round(s)`);
  console.log(
    `    ${o.updates} in-place update(s) measured across ${o.rounds} round(s); ${o.blind} touched a chart
` +
      `    AT RISK — ungrouped and with no parts list, which is what a refused \`reading back
` +
      `    an ungrouped chart's shape ids\` leaves behind.`,
  );
  // THE POPULATION `atRisk` CANNOT SEE, printed beside it so a zero above is
  // never read alone. #586 groups the majority the host names and leaves the
  // remainder out of the parts tag on purpose; that chart reports `group` at
  // update time, counts as safe, and still has loose shapes in its own box.
  if (o.subsetGroups)
    console.log(
      `    PLUS ${o.strandedByDesign} shape(s) left loose ON PURPOSE across ${o.subsetGroups} subset group(s) — #586's
` +
        `    trade, not a fault, and invisible to \`atRisk\` because the host calls a subset
` +
        `    group a group. Count these before calling any zero above an all-clear.`,
    );
  else
    console.log(
      `    #586's subset branch has not run in this pool: 0 partially-grouped chart(s). It is
` +
        `    reached only by a re-read still SHORT after the settled retry, which no round
` +
        `    has recorded. Read that as a FLOOR, not a count: until 2026-08-20 the COLD
` +
        `    read's outcome was never traced at all, so "none since the retry" compared a
` +
        `    post-settle read against 42 archived COLD ones. Different units. The cold read
` +
        `    is traced now, and the next rounds can say whether the fault went or is hidden.`,
    );
  if (o.unsettledKept)
    console.log(
      `    ${o.unsettledKept} reading(s) the two host reads DISAGREED on, kept and flagged rather than
` +
        `    dropped (${o.unsettledGrowth} shape(s) of growth between them, not pooled above). The second
` +
        `    read is taken: across the archive the deck adjudicates every one of these and backs the
` +
        `    second 48 times against the first's zero. A fifth of this instrument used to vanish here.`,
    );
  // THE INSTRUMENT'S OWN ERROR RATE, which is a number worth printing: this
  // reading has produced four false positives in two days, every one of them a
  // host lag the deck later contradicted.
  if (o.deckContradicted)
    console.log(
      `    ${o.deckContradicted} reading(s) DISCARDED because the deck disagreed — they claimed more shapes
` +
        `    on a slide than it finished the round holding, which a round that only adds
` +
        `    cannot do. Host lag outlasting the settle delay; two agreeing reads inside one
` +
        `    lag window agree on the stale number.`,
    );
  const usable = o.updates - o.unitMismatch - o.deckContradicted;
  if (usable > 0) {
    console.log(`    the slide GREW by ${String(o.blindGrowth).padStart(5)} shape(s), worst single update ${o.worst}`);
    console.log(`    on charts that HAD a list ${String(o.sightedGrowth).padStart(4)} — an ordinary change of size`);
    // ONLY WITH SOMETHING TO CONCLUDE FROM. This fired on a pool whose every
    // reading had just been excluded as unit-mismatched, announcing that the
    // growth "is not happening" on the strength of no measurements at all —
    // which is a report lying in a new direction rather than the old one.
    // GATED ON THE POPULATION THE QUESTION IS ABOUT, which is charts that were
    // ungrouped AND had no parts list — the only ones that can strand anything,
    // counted per update from the host's own `type`. The round-wide
    // `not grouping` count is a poor stand-in: it says a round contained such a
    // chart, not that an UPDATE ever touched one.
    if (o.blindGrowth === 0) {
      if (o.atRisk === 0)
        console.log(
          `    NOT AN ALL-CLEAR: no update here touched a chart that could strand anything
` +
            `    (0 at risk — every one was grouped, or had its parts list). Zero growth over
` +
            `    zero at-risk charts is the question never being put.
` +
            `    The round held ${o.ungroupedCharts} ungrouped chart(s); wait for one an update actually edits.`,
        );
      else
        console.log(
          `    ZERO, over ${o.atRisk} chart(s) that COULD have stranded: "the chart grows by a whole
` + `    chart on every edit" is a danger the code describes, not something happening.`,
        );
    }
  } else {
    console.log(`    no usable reading yet — every measurement here predates the instrument's fix.`);
  }
  if (o.unitMismatch)
    console.log(
      `    ${o.unitMismatch} update(s) are NOT counted above, from the two builds whose readings
` +
        `    cannot be believed. Round 082 subtracted mismatched units — "23 stranded" beside
` +
        `    its own before:3 after:3. Round 084 measured in one unit but read the host ONCE,
` +
        `    and every non-zero number it produced was the host lagging an addGroup it had
` +
        `    already committed: four slides "grew" by 23 while the deck ended with one grouped
` +
        `    chart on each. A reading now carries \`settled\` only when two reads agreed.`,
    );
}

function reportOriginTagLosses(logs) {
  const o = poolOriginTagLosses(logs);
  if (!o.charts) return;
  console.log(`
  CHARTS THAT CANNOT FOLLOW A DRAG — pooled over ${logs.length} round(s)`);
  console.log(
    `    origin tag lost  ${String(o.charts).padStart(4)} chart(s) across ${o.rounds} round(s), worst round ${o.worst}`,
  );
  console.log(
    `    The chart is re-editable and its config is intact — what it loses is the ability to
` +
      `    redraw where the USER left it, so an update snaps it back to where it was inserted.
` +
      `    \`an update follows a moved chart\` passes while this climbs, because a scenario tests
` +
      `    ONE chart and this counts them all. Rounds 073/074 lost 9 of 19 and 8 of 17.
` +
      `    No rate is printed: only failures are traced, so there is no honest denominator.`,
  );
}

/**
 * Which slide size a round ran at — its PROFILE, as a comparable string.
 *
 * `16:9` for everything filed before 2026-08-16, and that default is a fact
 * rather than a guess: all 53 rounds archived before the field existed were run
 * on a widescreen deck, which `docs/ROUNDS.md` states once so no reader has to
 * infer it per file.
 *
 * The profile exists because a 4:3 round and a 16:9 round are DIFFERENT
 * EXPERIMENTS, and pooling them is the rounds 24-and-25 mistake — "differed only
 * in this, and were compared as though they did not". Round 077 scored 10 of 13
 * where 16:9 scored 13 of 13 twice, so the difference is large enough to swamp
 * anything a pooled number would say.
 */
/**
 * Rounds whose two independent slide-size readings disagree.
 *
 * THE ROUND THAT MADE THIS NECESSARY: 115 and 116 recorded `720x540` from the
 * pane while the driver had read `960x540` off live `PageSetup`, twice, and
 * printed it before each round started. Every profile comparison in this file
 * groups by `slideSize`, so both rounds were filed into the wrong arm — the
 * exact failure `PW_EXPECT_SIZE` exists to prevent, sailing past it because the
 * guard and the archive read different sources and nobody compared them.
 *
 * A ROUND WITH NO DRIVER READING IS NOT A ROUND THAT AGREES. It is a round with
 * one opinion, and it is skipped rather than counted as a pass — every round
 * archived before `driverSlideSize` existed is in that state, and reporting 116
 * of them as "consistent" would be a lie told by a denominator.
 */
/**
 * Is the SECOND round of a pair systematically worse than the first?
 *
 * The pair exists to separate a real fault from the host's mood, and both halves
 * have been read as two samples of ONE condition. They are not. Pooled over
 * every build this archive has run twice — 31 pairs on 2026-08-20 — the second
 * round had more post-retry failures in 15, fewer in 2, and tied in 14.
 *
 * TIES ARE NOT COIN FLIPS. Most are older rounds whose counters sat at zero
 * both times, so counting them as evidence of symmetry is how this nearly got
 * waved away as noise: 15 against 2 among the pairs that moved is not a mood.
 *
 * Only the first two rounds of a build count, so a build run three times cannot
 * vote twice — and if it is ever run three times deliberately, that is the
 * experiment this finding asks for.
 */
/**
 * How long a round's own trace says it took, in seconds — or null.
 *
 * ALREADY IN EVERY ROUND FILE AND SURFACED BY NOTHING. Every entry carries an
 * `ms` offset, so the last one is the round's span; no reader has ever printed
 * it, and it turns out to be the strongest single predictor of a bad round in
 * this archive. The slower half of the rounds whose instruments existed average
 * 5.1 post-retry failures against 2.8, and 80 deck shapes against 48.
 *
 * It matters because the second round of a pair is 2.0-2.4x slower than the
 * first in all four pairs measured, and the likeliest reason is that the first
 * round gets mined WHILE THE SECOND RUNS. A reader who cannot see the span
 * cannot tell a degraded round from a clean one, and will read the difference
 * as evidence about the code.
 *
 * Null when the trace carries no usable offsets — an unreadable span must not
 * become a fast one.
 */
export function roundSpanSeconds(log) {
  const es = log?.trace?.entries;
  if (!Array.isArray(es) || !es.length) return null;
  const ms = es.map((e) => Number(e.ms) || 0);
  // LAST MINUS FIRST, NOT THE LAST. `ms` counts from the PANE'S load, not the
  // round's start, and the pane is not reloaded between rounds — so a round that
  // inherits its pane starts its clock wherever the previous round left off.
  //
  // The first version of this took `Math.max(...)` and therefore reported every
  // reused-pane round's duration as ITS OWN plus the previous round's. That
  // produced "the second round of a pair is 2.0-2.4x slower", which was
  // published twice before the arithmetic was checked. The real figure is about
  // 1.35x. See `paneAgeAtStartSeconds` for the fact the inflated number was
  // accidentally encoding.
  const span = Math.max(...ms) - Math.min(...ms);
  return span > 0 ? Math.round(span / 1000) : null;
}

/**
 * How long the pane had already been alive when this round started, in seconds.
 *
 * THE VARIABLE THAT ACTUALLY PREDICTS A BAD ROUND, and it was hiding inside a
 * metric that had been reported as duration. `ms` counts from the pane's load,
 * so the FIRST entry's offset is the pane's age at the moment the round began.
 *
 * Rounds 110-123, split on it:
 *
 *     fresh pane (< 200s)   post-retry 0, 2, 0, 0, 0, 1, 0   deck 14-45, mostly 16
 *     reused pane           post-retry 0, 5, 7, 3, 8, 7, 2   deck 16-97, mostly 60+
 *
 * Mean post-retry 0.43 against 4.57. This is what "the second round of a pair is
 * worse" always was: the second round is the one that inherits a pane. Position,
 * profile and observer load were all stand-ins for it.
 *
 * Null when unreadable — an unknown pane age must not read as a fresh one.
 */
export function paneAgeAtStartSeconds(log) {
  const es = log?.trace?.entries;
  if (!Array.isArray(es) || !es.length) return null;
  const ms = es.map((e) => Number(e.ms)).filter((n) => Number.isFinite(n));
  return ms.length ? Math.round(Math.min(...ms) / 1000) : null;
}

/** Under this many seconds old at a round's start, a pane counts as fresh. */
export const FRESH_PANE_SECONDS = 200;

/**
 * The fallback and repair signals the trace records and nothing reads.
 *
 * 71 of the 87 distinct messages in this archive are never named by triage, the
 * gate or the cycle. Most are narrative and that is fine. These four are not:
 * each records a path the code took because its FIRST choice failed, thousands
 * of times, with nobody watching the rate. Banded by round, mean per round:
 *
 *     rounds      tagging-failed  redraw-instead  scratch-retry  scratch-wrecked
 *       1- 40          5.2             9.3            9.3            10.7
 *      41- 80          6.6             9.1           14.1            11.1
 *      81-110          0.4            12.9           14.8            11.2
 *     111-141          0.2            13.0           14.6            10.7
 *
 * TWO THINGS WERE HAPPENING AND NOBODY COULD SEE EITHER. Tagging failures
 * collapsed thirtyfold around round 81 — a real win, never verified, and if it
 * ever regresses nothing will say so. Meanwhile in-place updates began falling
 * back to a full redraw 40% more often, which is slower and touches more of the
 * deck, and that drift went unremarked across sixty rounds.
 *
 * A count that no one reads is not observability. This makes the gate say them.
 */
export const FALLBACK_SIGNALS = {
  "tagging failed — charts are not re-editable until repaired": "charts left un-tagged",
  "not updating in place — redrawing instead": "in-place update fell back to a redraw",
  // ITS OWN SIGNAL, not folded into the line above. The in-place update either
  // declines by rule or THROWS, and this tracked only the first — the same
  // two-of-three gap `poolInPlaceUpdates` had, missed when that one was fixed.
  // Kept separate rather than merged because a write the HOST threw out is a
  // different event from the differ declining work a redraw does better, and
  // drift in one must not be absorbed by the other: rounds 144 and 145 carried
  // 2 and 3 of these against a flat zero everywhere else, and nothing said so.
  "in-place update refused — redrawing instead": "in-place update refused BY THE HOST",
  "took another scratch slide after giving up on the last": "scratch slide retried",
  "giving up the scratch slide this question wrecked": "scratch slide wrecked",
};

/**
 * Each fallback's count this round against the median of its priors.
 *
 * THREE PRIORS MINIMUM, the same rule `poolScenarioPopulations` and
 * `poolGroupingOutcome` needed: a "usually" from one observation is not a
 * baseline, and this project's own noise floor is the argument.
 */
/**
 * Has the in-place chart update EVER worked?
 *
 * IT DOES NOW, AND THIS PARAGRAPH IS KEPT BECAUSE IT IS WHY. It was added in
 * #405 ("Change one thing, write one shape") and #406 was titled "The in-place
 * update fired zero times and would not say why" — that PR added the fallback
 * trace to diagnose it. THE DIAGNOSTIC HAD BEEN ANSWERING EVER SINCE AND NOBODY
 * HAD READ IT: across 117 archived rounds the success line `updated only the
 * shapes that changed` appeared ZERO times while the fallback fired 12-13 times
 * a round, one reason 12 times of 13 — "the chart has no parts list, so its
 * nodes cannot be mapped".
 *
 * Reading it is what fixed it. A grouped chart never has a parts list by
 * design, and this host groups nearly every chart; `groupMembersInOrder` reads
 * the group's members instead. **The update now runs 11 of 13 attempts every
 * round**, and the fallback count fell 13 → 2 at round 153.
 *
 * A feature that has never once run in production is not a feature; it is a
 * branch that costs a fallback every time. This makes the gate say so — and it
 * is the reason the gate now prints the fallback SEQUENCE as well as its
 * median, because a median cannot see the step this fix produced.
 */
/**
 * The in-place update has THREE outcomes and this counted two.
 *
 * `tryInPlaceUpdate` declines by rule (`not updating in place`, carrying a
 * `why`) or THROWS (`in-place update refused`, carrying an `error` and no
 * `why`). Only the first was counted, so every host-side refusal was invisible
 * here — including the three `InvalidArgument | errorLocation=Shape.textFrame`
 * in round 145 that turned out to be the last thing standing between this
 * feature and working. They sat in round 144's trace too, uncounted.
 *
 * The reason breakdown exists for the same reason. A total says a feature fell
 * back; only the reasons say whether that was the differ declining work a
 * redraw does better (correct) or the host refusing a write (a defect).
 */
function inPlaceErrorKey(err) {
  const code = err.split(" | ")[0]?.trim() || "threw";
  const where = /"errorLocation":"([^"]+)"/.exec(err)?.[1];
  return where ? `${code} at ${where}` : code;
}

export function poolInPlaceUpdates(logs) {
  let ok = 0,
    fell = 0,
    threw = 0,
    rounds = 0;
  const reasons = new Map();
  // ONE EXAMPLE, not just a tally. A residual bucket reported as a number is
  // read as noise; reported as a line it gets opened. That distinction cost two
  // rounds of a host defect sitting in plain sight.
  const unexplained = [];
  const note = (key) => reasons.set(key, (reasons.get(key) ?? 0) + 1);
  for (const log of logs ?? []) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    rounds++;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m === "updated only the shapes that changed") ok++;
      else if (m === "not updating in place — redrawing instead") {
        fell++;
        const why = e.data?.why;
        if (why) note(String(why));
        else unexplained.push(m);
      } else if (m === "in-place update refused — redrawing instead") {
        threw++;
        const err = String(e.data?.error ?? "");
        if (err) note(inPlaceErrorKey(err));
        else unexplained.push(m);
      }
    }
  }
  return {
    ok,
    fell,
    threw,
    rounds,
    unexplained,
    reasons: [...reasons].sort((a, b) => b[1] - a[1]).map(([why, n]) => ({ why, n })),
  };
}

/**
 * What the driver had to do before each round would start.
 *
 * A successful recovery erases its own evidence — the round that follows looks
 * like one that never needed rescuing — so `driverRun` records the attempts and
 * the stops behind them. This pools that across the archive.
 *
 * REPORTED AS COUNTS, NOT A RATE, and deliberately. `driverRun` is only as old
 * as round 150, so most of the archive carries nothing, and a percentage over a
 * denominator of four would be a number with a decimal point and no meaning.
 * `rounds` is the rounds that CARRY the field; anything older is not a clean
 * round, it is an unmeasured one.
 */
/**
 * Trace signals a round RECORDS that no tool READS.
 *
 * This repo has found the same thing by hand twice. `poolFallbackRates` exists
 * because the fallback lines had been "recorded thousands of times and read by
 * nothing"; `poolInPlaceUpdates` because the in-place diagnostic "has been
 * answering ever since and nobody has read it". Both were discovered by someone
 * scrolling a round file, months after the data started arriving.
 *
 * Round 153 records 49 distinct messages and 35 of them are read by nothing.
 * Most of those are noise and should stay unread — the point is not to pool
 * them all, it is to stop the next `poolFallbackRates` waiting months for
 * someone to notice it by hand.
 *
 * A FLOOR, NOT A COUNT, and the gate says so. "Read" here means the tool source
 * mentions the message verbatim, so a matcher built by concatenation reads as
 * unread. That errs toward offering too much rather than hiding something, and
 * a detector that reports a floor must say which it is — see
 * `poolFallbackRates` and the dead-detector guard in the tests.
 */
export function unreadSignals(log, toolSource, top = 6) {
  const entries = log?.trace?.entries;
  if (!Array.isArray(entries)) return [];
  const counts = new Map();
  for (const e of entries) {
    const m = String(e?.message ?? "");
    if (m) counts.set(m, (counts.get(m) ?? 0) + 1);
  }
  const src = String(toolSource ?? "");
  return [...counts]
    .filter(([m]) => !src.includes(m))
    .sort((a, b) => b[1] - a[1])
    .slice(0, top)
    .map(([message, n]) => ({ message, n }));
}

export function poolDriverRuns(logs) {
  let rounds = 0,
    clean = 0;
  const causes = new Map();
  const attempts = new Map();
  // THE ARM OF THE EXPERIMENT — SPLIT ON THE PANE, NOT ON THE FLAG.
  //
  // The first version of this split on `driverRun.fresh`, and round 167 showed
  // within one pair why that is wrong: 166 ran WITHOUT `--fresh` and started on
  // a 69-second pane anyway, because a merge preceded it and recovery reloads a
  // stale pane. The flag and the pane are different variables, and the whole
  // PAIR POSITION argument is about the pane — `--fresh` is merely one way to
  // get a fresh one. Splitting on the flag files a fresh-pane round under
  // "aged" and makes both arms mean nothing.
  //
  // The house defect, in an instrument built four hours earlier to fix the same
  // defect: measuring the PROXY instead of the thing.
  //
  // `flagDisagreed` counts rounds where the two answers differ, so the proxy's
  // unreliability is on the page rather than assumed away. `unlabelled` counts
  // rounds from before either was recorded, so a rate over a handful is never
  // quoted as if it covered the archive.
  const arms = { fresh: 0, freshRefusedNone: 0, aged: 0, agedRefusedNone: 0, unlabelled: 0, flagDisagreed: 0 };
  /**
   * CRASHES PER ATTEMPT, SPLIT ON THE SLIDE SIZE — the denominator `BACKLOG`
   * says nobody had ever computed.
   *
   * The archive has carried every input to this number for months and no
   * instrument divided one by the other. Measured over 322 rounds it is
   * 40 crash events in 81 attempts at 4:3 against 11 in 365 at 16:9 — a
   * sixteen-fold gap that no report printed.
   *
   * ATTEMPTS, NOT ROUNDS, is the denominator on purpose: a round that crashed
   * three times and finally landed is one round and four attempts, and counting
   * rounds hides exactly the arm that is failing hardest.
   *
   * WHAT THIS NUMBER DOES NOT SAY is which variable it belongs to. `cyclePlan`
   * pairs 4:3 with one deck and 16:9 with another and has never crossed them,
   * so aspect ratio and deck file are perfectly confounded here. Session
   * position is NOT the explanation — recomputed from `startedAt` with the
   * driver's own 45-minute rule, 4:3 runs 47-51% at every position while 16:9
   * runs 6.5% at the first and 0% at every later one — but that leaves two
   * candidates, not one. Print the gap; do not name its cause.
   */
  const bySize = new Map();
  for (const log of logs ?? []) {
    const dr = log?.driverRun;
    if (!dr || typeof dr.attempts !== "number") continue;
    rounds++;
    attempts.set(dr.attempts, (attempts.get(dr.attempts) ?? 0) + 1);
    if (dr.attempts <= 1) clean++;
    const size = String(log?.driverSlideSize ?? "unrecorded");
    const bucket = bySize.get(size) ?? { size, rounds: 0, attempts: 0, crashes: 0, roundsWithCrash: 0 };
    bucket.rounds++;
    bucket.attempts += dr.attempts;
    const crashed = (dr.recovered ?? []).filter((c) => String(c).includes("crashed")).length;
    bucket.crashes += crashed;
    if (crashed) bucket.roundsWithCrash++;
    bySize.set(size, bucket);
    for (const c of dr.recovered ?? []) causes.set(String(c), (causes.get(String(c)) ?? 0) + 1);
    const age = paneAgeAtStartSeconds(log);
    if (typeof age !== "number") {
      arms.unlabelled++;
      continue;
    }
    const paneWasFresh = age < FRESH_PANE_SECONDS;
    if (typeof dr.fresh === "boolean" && dr.fresh !== paneWasFresh) arms.flagDisagreed++;
    const refusedNone = !(log?.trace?.entries ?? []).some((e) => /^not grouping/.test(String(e.message ?? "")));
    if (paneWasFresh) {
      arms.fresh++;
      if (refusedNone) arms.freshRefusedNone++;
    } else {
      arms.aged++;
      if (refusedNone) arms.agedRefusedNone++;
    }
  }
  return {
    rounds,
    clean,
    arms,
    attempts: [...attempts].sort((a, b) => a[0] - b[0]).map(([n, of]) => ({ attempts: n, rounds: of })),
    causes: [...causes].sort((a, b) => b[1] - a[1]).map(([cause, n]) => ({ cause, n })),
    // Worst first: the point of the row is to be impossible to skim past.
    bySize: [...bySize.values()].sort((a, b) => b.crashes / b.attempts - a.crashes / a.attempts),
  };
}

/**
 * How badly this host renumbers its slide ids when one is appended.
 *
 * RECORDED SINCE ROUND 041 AND READ BY NOTHING. `claimed the appended slide
 * though the id list churned` carries `before`, `after` and `fresh` on every
 * occurrence — 119 of them across 63 rounds — and no script has ever matched on
 * it. The comment in `powerpoint.ts` that describes the behaviour was written
 * before the archive existed, from seven hand-collected observations, and said
 * the count was ALWAYS two.
 *
 * It is not. 97 events read `fresh=2` and 22 read `fresh=3` — one append in
 * five renumbers two — and the two populations separate almost cleanly by deck
 * size: `fresh=2` has median `before` 12, `fresh=3` has median 77. The original
 * seven were all taken at before=3..37, where `fresh=3` is nearly absent.
 *
 * Which is why this is pooled rather than left in a paragraph: a hand-collected
 * sample cannot notice that its own conclusion is a function of deck size, and
 * the archive can.
 */
export function poolIdChurn(logs) {
  const byFresh = new Map();
  let rounds = 0;
  for (const log of logs ?? []) {
    let any = false;
    for (const e of log?.trace?.entries ?? []) {
      if (!/claimed the appended slide/.test(String(e.message ?? ""))) continue;
      const fresh = Number(e.data?.fresh);
      const before = Number(e.data?.before);
      if (!Number.isFinite(fresh)) continue;
      any = true;
      const t = byFresh.get(fresh) ?? { fresh, events: 0, befores: [] };
      t.events++;
      if (Number.isFinite(before)) t.befores.push(before);
      byFresh.set(fresh, t);
    }
    if (any) rounds++;
  }
  const median = (xs) => (xs.length ? [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)] : undefined);
  return {
    rounds,
    kinds: [...byFresh.values()]
      .sort((a, b) => a.fresh - b.fresh)
      .map((t) => ({ fresh: t.fresh, events: t.events, medianDeck: median(t.befores) })),
  };
}

/**
 * The fullest slide each round left behind.
 *
 * A CLEAN ROUND ENDS `0,4,1,2,5,1,1` — nothing above five. Eight of the last
 * thirty rounds ended with a slide holding between 11 and 48, and **every one
 * of them reported 13 of 13 scenarios passed**, including rounds 159, 161 and
 * 167. A chart that fails to group is left as its loose shapes, and no scenario
 * verdict looks at the deck.
 *
 * `does a rasterise poison the next draw` is the clearest case: it draws four
 * charts and its verdict is `all four draws landed`, which is literally true and
 * narrower than any reader takes it. It asks whether the CALL came back, not
 * whether a chart survived — so it passes with eight loose shapes sitting where
 * a chart should be.
 *
 * The gate has printed the inventory all along. Nothing compared it to
 * anything, which is the difference between a number being on screen and a
 * number being read.
 */
export function poolFullestSlide(logs, n = RECENT_IN_A_ROW) {
  return (logs ?? []).slice(-n).map((log) => {
    const counts = (log?.deck?.inventory ?? []).map((s) => Number(s?.count ?? s?.shapes?.length ?? 0));
    return counts.length ? Math.max(...counts) : 0;
  });
}

/**
 * Above this many shapes on one slide, a chart did not group.
 *
 * Five is what a clean round's fullest slide holds — one grouped chart per
 * probe slide, and the four one-batch charts the rasterise scenario adds beside
 * one of them. A loose chart contributes eight or more on its own, so there is
 * no near-miss band to argue about.
 */
export const CLEAN_SLIDE_CEILING = 6;

/**
 * Why a chart left the draw without a parts list.
 *
 * `tracePartsOutcome` was built to separate three faults that a day of archive
 * mining could not tell apart — never built because the chart was grouped;
 * built and lost to a throwing read-back; built and not found again. It has
 * recorded all four counters on every one of its 607 events across 29 rounds,
 * and **no script has ever read them.**
 *
 * Reading them answers the question it was built for, and the answer is mostly
 * reassuring:
 *
 *   346  grouped, so not loose — correct by design, a group needs no parts list
 *   223  no charts in the call at all
 *    36  the id read-back THREW — the real failure, ~1.2 per round
 *     2  reached the normal exit and still produced nothing
 *
 * WHICH REFUTES THE OBVIOUS READING. `gotPartsList` is 0 on all 607, and taken
 * alone that says a production path has never once worked. 569 of the 607 are
 * cases where a parts list is not wanted. The counter that matters is `where`,
 * and it was sitting in the same object the whole time.
 */
export function poolPartsListOutcome(logs) {
  const out = { grouped: 0, noCharts: 0, threw: 0, builtNothing: 0, gotList: 0, events: 0, rounds: 0 };
  for (const log of logs ?? []) {
    let any = false;
    for (const e of log?.trace?.entries ?? []) {
      if (!/parts list outcome/.test(String(e.message ?? ""))) continue;
      const d = e.data ?? {};
      any = true;
      out.events++;
      out.gotList += Number(d.gotPartsList) || 0;
      if (!Number(d.charts)) out.noCharts++;
      else if (Number(d.groupedSoNotLoose) === Number(d.charts)) out.grouped++;
      else if (/threw/.test(String(d.where))) out.threw++;
      else out.builtNothing++;
    }
    if (any) out.rounds++;
  }
  return out;
}

/**
 * Did the positional guess pick this chart's own shapes, or another chart's?
 *
 * The last-resort branch in `chooseGroupMembers` takes the TAIL of the host's
 * shape listing when no id matched. **The listing can be one grouping-event
 * stale**, and then its tail is the previous chart's shapes.
 *
 * When those shapes have already been absorbed into that chart's group they no
 * longer exist at top level, every `getItemOrNullObject` returns a null object,
 * and `addGroup` throws `InvalidArgument` — 29 times across rounds 068-175,
 * every one preceded by this branch. When they are still LOOSE, the same guess
 * SUCCEEDS on the wrong chart's shapes and feeds them to the parts tag as this
 * chart's own, throwing nothing and tracing nothing.
 *
 * That second case is why `mine` and `chose` are recorded. The throw is visible
 * in an error payload; the silent mis-group was not visible anywhere, and this
 * is the reading that separates them.
 */
export function poolPositionalGuess(logs) {
  const out = { events: 0, rounds: 0, mine: 0, other: 0, partial: 0, unreadable: 0, refused: 0 };
  for (const log of logs ?? []) {
    let any = false;
    for (const e of log?.trace?.entries ?? []) {
      // THE GUARD'S OWN LINE, counted alongside the guesses it judged. Without
      // it this section reports "3 picked ANOTHER chart's shapes entirely" over
      // a paragraph describing silent corruption, and every reader has to go to
      // the archive to find out whether anything was actually grouped. Two of
      // those three were REFUSED; the third is round 179, which predates the
      // guard and is the round whose write-up the guard was built from.
      if (/^not grouping: the positional guess named no shape of ours/.test(String(e.message ?? ""))) {
        out.refused++;
        any = true;
        // NO `continue` — it was here and it was dead. The refusal message does
        // not match the `picked the tail` test below either, so it is skipped
        // with or without it; the mutant that removed it could not be killed.
      }
      if (!/^the positional guess picked the tail/.test(String(e.message ?? ""))) continue;
      any = true;
      out.events++;
      const mine = e.data?.mine ?? [];
      const chose = e.data?.chose ?? [];
      if (!Array.isArray(mine) || !Array.isArray(chose) || chose.some((id) => id == null)) {
        out.unreadable++;
        continue;
      }
      const own = new Set(mine.filter((id) => id != null).map(String));
      const hit = chose.filter((id) => own.has(String(id))).length;
      if (hit === chose.length) out.mine++;
      else if (hit === 0) out.other++;
      else out.partial++;
    }
    if (any) out.rounds++;
  }
  return out;
}

/** How many rounds of a signal to print in sequence, so a step is visible. */
export const RECENT_IN_A_ROW = 8;

export function poolFallbackRates(logs) {
  const per = [];
  // The index of the newest log that CONTRIBUTED, kept beside `per` rather than
  // inside it: these entries are keyed by signal name and downstream code walks
  // their keys, so an `at` field would read as a signal called "at".
  let lastAt = -1;
  (logs ?? []).forEach((log, at) => {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) return;
    const c = {};
    for (const key of Object.keys(FALLBACK_SIGNALS)) c[key] = 0;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m in c) c[m]++;
    }
    per.push(c);
    lastAt = at;
  });
  if (per.length < MIN_PRIORS_FOR_A_BASELINE + 1) return [];
  // The stale-`now` guard — see `isTheRoundBeingJudged`. A round with no trace
  // at all is skipped above, and this said "now" about the newest round that
  // HAD one.
  if (!isTheRoundBeingJudged(lastAt, logs)) return [];
  const now = per[per.length - 1];
  const priors = per.slice(0, -1);
  // DRIFT, NOT JUST TODAY. Comparing this round against the median of ALL
  // priors cannot see a slow climb: the median absorbs it. The signal that
  // motivated this instrument rose from 9.3 to 13.0 per round over sixty
  // rounds, and by the time it was noticed "now" and "usually" were both 13 —
  // so a now-against-median check would have reported it as normal for as long
  // as it kept getting worse.
  //
  // The oldest third against the newest third catches exactly that shape, and
  // is why this returns both readings rather than one.
  const third = Math.max(1, Math.floor(per.length / 3));
  const med = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
  const out = [];
  for (const [key, label] of Object.entries(FALLBACK_SIGNALS)) {
    const median = med(priors.map((p) => p[key]));
    const oldest = med(per.slice(0, third).map((p) => p[key]));
    const newest = med(per.slice(-third).map((p) => p[key]));
    // THE LAST FEW VALUES IN SEQUENCE, because every summary above is a median
    // and A MEDIAN CANNOT SEE A STEP. `in-place update fell back to a redraw`
    // reads 13,13,11,10,10,8,8,2,2,2,2,2 over the last twelve rounds — the
    // feature started working and the fallback collapsed — and all three
    // readings miss it: all-time median 12, last-20 median 10, and the thirds
    // say "RISING, 8 to 13". A reader is told to chase a regression that ended
    // five rounds ago, next to a `now` of 2 that contradicts it.
    //
    // A wider window is not the fix; a median over ANY fixed window smears a
    // step by construction. The sequence is the evidence, and it costs eight
    // numbers.
    const recent = per.slice(-RECENT_IN_A_ROW).map((p) => p[key]);
    out.push({ key, label, now: now[key], median, oldest, newest, recent, rounds: priors.length, span: third });
  }
  return out;
}

export function poolPairPosition(logs) {
  const KIND = (m, d) => {
    const k = String(d.kind ?? "");
    if (k) return k;
    if (/came back empty/.test(m)) return "empty";
    if (/named none/.test(m)) return "zero-match";
    if (/matched only some/.test(m)) return "short";
    return null;
  };
  const byBuild = new Map();
  for (const log of logs ?? []) {
    const build = String(log?.build ?? "").split(" ")[0];
    if (!build) continue;
    let post = 0;
    for (const e of log?.trace?.entries ?? []) {
      const d = e.data ?? {};
      if (d.afterRetry === true && KIND(String(e.message ?? ""), d)) post++;
    }
    const deck = (log?.deck?.inventory ?? []).map((s) => s.count ?? s.shapes?.length ?? 0);
    if (!byBuild.has(build)) byBuild.set(build, []);
    byBuild.get(build).push({ post, age: paneAgeAtStartSeconds(log), deck: deck.reduce((a, b) => a + b, 0) });
  }
  let pairs = 0,
    worse = 0,
    better = 0,
    tied = 0,
    secondFresh = 0;
  for (const rounds of byBuild.values()) {
    if (rounds.length < 2) continue;
    const [a, b] = rounds;
    pairs++;
    if (b.post > a.post) worse++;
    else if (b.post < a.post) better++;
    else tied++;
    // WHETHER THE SECOND ROUND STARTED ON A FRESH PANE, which is the whole
    // reason this asymmetry exists. Counted so the report can distinguish
    // pairs that PREDATE the between-rounds reload from ones that do not,
    // instead of announcing a fixed problem forever from historical data.
    if (typeof b.age === "number" && b.age < FRESH_PANE_SECONDS) secondFresh++;
  }
  return { pairs, worse, better, tied, secondFresh };
}

/**
 * Charts that landed off the slide, or on top of each other.
 *
 * NEWLY ANSWERABLE, AND NEVER ASKED BEFORE. Until 2026-09-01 the deck inventory
 * recorded a shape's ORIGIN and no extent, so across 322 rounds and 10,362
 * archived shapes the only geometry fault it could express was an origin outside
 * the slide. "Did this chart land inside the slide" and "did two charts land on
 * top of each other" were unanswerable from the archive — and the second is not
 * hypothetical: the owner opened the deck the battery left him and found one
 * full-size chart drawn over another, then three in a heap. Both times a person
 * looking at a slide was the instrument.
 *
 * `items/width,items/height` ride along now, so both questions can be asked of
 * every round from here on.
 *
 * GROUPS ONLY, and that is not a shortcut. A chart the host grouped is ONE
 * top-level shape named `PowerChart`; a chart it refused to group is its loose
 * parts — `title`, `seg-0-0`, `category-1` — and those overlap each other by
 * design, because a label sits on the segment it names. Comparing everything
 * would report every ungrouped chart as a collision. Round 347 is exactly that
 * mix: 13 `PowerChart` groups and 7 loose parts.
 *
 * A shape with no size counts as UNMEASURED rather than clean: a round archived
 * before this landed, or a host that will not answer the two properties, must
 * never read as a deck with nothing wrong with it.
 */
export function deckGeometryFaults(log) {
  const slide = log?.slideSize;
  const inventory = log?.deck?.inventory;
  if (!slide || typeof slide.width !== "number" || typeof slide.height !== "number" || !Array.isArray(inventory))
    return null;
  // Half a point, for float noise off the host. A chart flush against the edge
  // is placed, not spilled.
  const TOL = 0.5;
  const offSlide = [];
  const collisions = [];
  let measured = 0;
  let unmeasured = 0;
  for (const s of inventory) {
    const groups = [];
    for (const sh of s.shapes ?? []) {
      if (typeof sh.width !== "number" || typeof sh.height !== "number") {
        unmeasured++;
        continue;
      }
      measured++;
      if (sh.name !== "PowerChart") continue;
      const box = { id: sh.id, x0: sh.left, y0: sh.top, x1: sh.left + sh.width, y1: sh.top + sh.height };
      const spill = [];
      if (box.x0 < -TOL) spill.push("left");
      if (box.y0 < -TOL) spill.push("top");
      if (box.x1 > slide.width + TOL) spill.push("right");
      if (box.y1 > slide.height + TOL) spill.push("bottom");
      if (spill.length)
        offSlide.push({
          slideId: s.slideId,
          id: sh.id,
          off: spill.join("+"),
          box: `${Math.round(sh.left)},${Math.round(sh.top)} ${Math.round(sh.width)}x${Math.round(sh.height)}`,
        });
      groups.push(box);
    }
    for (let i = 0; i < groups.length; i++)
      for (let j = i + 1; j < groups.length; j++) {
        const a = groups[i];
        const b = groups[j];
        const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
        const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
        if (w > TOL && h > TOL) collisions.push({ slideId: s.slideId, a: a.id, b: b.id, area: Math.round(w * h) });
      }
  }
  return { measured, unmeasured, offSlide, collisions };
}

/**
 * Probes that gave two runs of the SAME COMMIT different answers.
 *
 * NOTHING IN THE GATE CHAIN HAS EVER READ `hostAnswers`. `rounds-gate.mjs`
 * compares scenario verdicts and slide-size divergence and contains zero
 * references to the probe sheet — while 14 of the 15 scenarios pass in every
 * round and seven have a lifetime failure count of zero across 322 rounds. The
 * discriminating content of a round has moved into the half nothing gates.
 *
 * Measured 2026-09-01 over 82 builds carrying two or more archived rounds: 328
 * of 3,147 comparable probe slots — 10.4% — differ between rounds of the same
 * commit, and 34 of the 44 probes have flipped within a build at least once.
 * Same code, same host, different answer.
 *
 * `UNSTABLE_ANSWERS` exists to declare exactly this and was assembled BY HAND
 * across ten rounds, which is why most of the 34 are missing from it. This makes
 * the list derivable from evidence instead.
 *
 * A caveat worth carrying: some of these flips are the answer-ranking defects
 * fixed on 2026-08-31 and 2026-09-01 — `unreadable`, `silent` and a probe's own
 * `all` locking a row against a real answer that arrived on a later pass. Those
 * should stop flipping now, and the ones that keep flipping are the real
 * non-determinism. Re-derive after a few rounds rather than declaring today's
 * list permanent.
 */
export function probeFlipsWithinBuild(logs) {
  const byBuild = new Map();
  for (const log of logs ?? []) {
    const answers = log?.hostAnswers?.answers;
    // The sha alone: `build` reads "f6e3568 · 2026-08-31 07:59Z", and the
    // timestamp differs between two rounds of the same commit.
    const sha = String(log?.build ?? "").split(" ")[0];
    if (!sha || !Array.isArray(answers)) continue;
    if (!byBuild.has(sha)) byBuild.set(sha, []);
    byBuild.get(sha).push(Object.fromEntries(answers.map((a) => [a.id, a.answer])));
  }
  const flips = new Map();
  let builds = 0,
    slots = 0,
    differing = 0;
  for (const sheets of byBuild.values()) {
    if (sheets.length < 2) continue;
    builds++;
    for (const id of new Set(sheets.flatMap((s) => Object.keys(s)))) {
      /**
       * REAL ANSWERS ONLY, and it changes the number considerably.
       *
       * A probe absent from one sheet is a gap in coverage, not a disagreement —
       * counting it would report the day a probe was added as instability. And a
       * round answering `no-scratch-slide`, `unreadable` or `silent` did not
       * give a DIFFERENT answer; it gave none. Reading `yes` against
       * `unreadable` as the host contradicting itself is the same mistake
       * `stabilityOf` carried until 2026-08-31, and it is precisely how
       * `UNSTABLE_ANSWERS` would fill up with rows describing our own read
       * failures as host non-determinism.
       */
      const seen = sheets
        .map((s) => s[id])
        .filter((v) => v !== undefined && !NOT_ASKED_ANSWERS.has(v) && !UNINFORMATIVE_ANSWERS.has(v));
      if (seen.length < 2) continue;
      slots++;
      if (new Set(seen).size === 1) continue;
      differing++;
      const row = flips.get(id) ?? { id, builds: 0, answers: new Set() };
      row.builds++;
      for (const v of seen) row.answers.add(v);
      flips.set(id, row);
    }
  }
  return {
    builds,
    slots,
    differing,
    flips: [...flips.values()]
      .sort((a, b) => b.builds - a.builds)
      .map((f) => ({ id: f.id, builds: f.builds, answers: [...f.answers].sort() })),
  };
}

export function poolProfileDisagreements(logs) {
  const out = [];
  for (const log of logs ?? []) {
    const pane = log?.slideSize;
    const driver = log?.driverSlideSize;
    if (!driver || !pane || typeof pane.width !== "number") continue;
    // The driver records a PROFILE STRING ("16:9"), the pane records points.
    // Compare them as profiles, which is the unit every consumer groups by.
    const paneProfile = roundProfile(log);
    if (String(driver) !== paneProfile)
      out.push({
        build: String(log.build ?? "").split(" ")[0],
        pane: paneProfile,
        driver: String(driver),
        source: pane.source ?? "?",
      });
  }
  return out;
}

export function roundProfile(log) {
  const s = log?.slideSize;
  if (!s || typeof s.width !== "number" || typeof s.height !== "number") return "16:9";
  // Named ratios for the two PowerPoint offers, and the raw size for anything
  // else — a custom deck should be visibly its own profile, not silently folded
  // into whichever named one it is nearest.
  const r = s.width / s.height;
  if (Math.abs(r - 16 / 9) < 0.01) return "16:9";
  if (Math.abs(r - 4 / 3) < 0.01) return "4:3";
  return `${Math.round(s.width)}x${Math.round(s.height)}`;
}

/**
 * Scenarios that PASS at one slide size and FAIL at another, in the same cycle.
 *
 * A different question from the regression gate, and it needs its own answer.
 * The gate asks "did this fall against its own history"; this asks "did this
 * profile fail what another profile passed tonight". Round 077 was exactly the
 * second — 10 of 13 at 4:3 against 13 of 13 at 16:9 — and nothing would have
 * said so automatically.
 *
 * BUILD-SCOPED, because that is the only way the comparison is fair. Two rounds
 * on different builds differ for reasons that have nothing to do with slide
 * size, and this must never report those.
 *
 * A scenario failing in BOTH profiles is not divergence — it is an ordinary bug,
 * already visible everywhere else, and reporting it here would bury the one
 * signal this exists for.
 */
export function profileDivergence(logs) {
  const byBuild = new Map();
  for (const log of logs) {
    const build = String(log?.build ?? "").split(" ")[0];
    if (!build) continue;
    const prof = roundProfile(log);
    if (!byBuild.has(build)) byBuild.set(build, new Map());
    const profs = byBuild.get(build);
    if (!profs.has(prof)) profs.set(prof, new Map());
    const seen = profs.get(prof);
    for (const sc of log?.selftest ?? []) {
      if (!sc?.name) continue;
      // A scenario that DID NOT MEASURE says nothing about its profile — the
      // same rule the regression gate follows, for the same reason.
      if (sc.skipped) continue;
      // BOTH outcomes kept, not collapsed to the worst. A profile that
      // disagrees with ITSELF is flaky, not different from another profile, and
      // the two want opposite responses — see below.
      const at = seen.get(sc.name) ?? { pass: 0, fail: 0 };
      if (sc.ok) at.pass++;
      else at.fail++;
      seen.set(sc.name, at);
    }
  }
  const out = [];
  for (const [build, profs] of byBuild) {
    if (profs.size < 2) continue;
    const names = new Set([...profs.values()].flatMap((m) => [...m.keys()]));
    for (const name of names) {
      const passedIn = [];
      const failedIn = [];
      const unstableIn = [];
      for (const [prof, m] of profs) {
        const at = m.get(name);
        if (!at) continue;
        // A profile that both passed and failed this scenario on one build has
        // not told us anything about its slide size — it has told us the
        // scenario is flaky. Reporting that as divergence sends someone to
        // investigate an aspect ratio for a difference the profile produces on
        // its own, which is worse than saying nothing.
        if (at.pass && at.fail) unstableIn.push(prof);
        else if (at.pass) passedIn.push(prof);
        else failedIn.push(prof);
      }
      if (passedIn.length && failedIn.length) out.push({ build, name, passedIn, failedIn, unstableIn });
      else if (unstableIn.length && (passedIn.length || failedIn.length))
        // Worth saying, and NOT as divergence. This is the shape the check's
        // first live outing produced: `explode a degraded picture` passed at
        // 4:3 and then passed once and failed once at 16:9, on build 17a8204.
        // "Diverged between slide sizes" was true of the worst reading and
        // misleading about the cause.
        out.push({ build, name, passedIn, failedIn, unstableIn, flaky: true });
    }
  }
  /**
   * WHAT THE SCENARIO DOES AT EACH PROFILE OVER ITS WHOLE LIFE, attached to
   * every divergence.
   *
   * The comparison above can only see one build. It already refuses to call a
   * profile that disagrees with ITSELF a divergence — but it cannot see that a
   * scenario fails a quarter of the time at BOTH profiles, in which case a
   * single pass-here-fail-there split is the commonest thing chance produces.
   *
   * Not hypothetical. The first divergence this gate reported after it was
   * unblocked on 2026-09-08 was `stop a run mid-draw`, passed at 16:9 and
   * failed at 4:3 on build 31d358f, carrying a 4:3-flavoured detail — "the
   * slide this drew on could not be deleted". Its lifetime record is 10 of 39
   * at 4:3 against 9 of 41 at 16:9, and it has not failed that way since round
   * 383. Four points apart, fifty rounds stale, and the report was pointing at
   * an aspect ratio.
   *
   * So the rates travel with the finding. A reader who sees 25.6% against 22.0%
   * closes it in a second; one who does not spends an afternoon on it, and
   * `docs/ROUNDS.md` says what that does to a report people are meant to keep
   * reading.
   */
  // NESTED, not a joined string key. The first version built `name + separator
  // + profile`, and a NUL byte landed in the separator on the way to disk — in
  // both places that built one. `no-control-bytes.test.ts` caught it, which is
  // exactly what that test is for. A nested map has no separator to corrupt,
  // and none to collide with either: a scenario name is free text.
  const lifetime = new Map();
  for (const log of logs) {
    const prof = roundProfile(log);
    for (const sc of log?.selftest ?? []) {
      // Same rule as above: a scenario that did not measure says nothing, so it
      // must not enter the denominator either.
      if (!sc?.name || sc.skipped) continue;
      if (!lifetime.has(sc.name)) lifetime.set(sc.name, new Map());
      const byProfile = lifetime.get(sc.name);
      const at = byProfile.get(prof) ?? { ran: 0, failed: 0 };
      at.ran++;
      if (!sc.ok) at.failed++;
      byProfile.set(prof, at);
    }
  }
  for (const d of out) {
    d.history = {};
    for (const prof of [...d.passedIn, ...d.failedIn, ...d.unstableIn])
      d.history[prof] = lifetime.get(d.name)?.get(prof) ?? { ran: 0, failed: 0 };
  }
  return out.sort((a, b) => a.build.localeCompare(b.build) || a.name.localeCompare(b.name));
}

/**
 * Scenarios that were passing and have stopped — the only automatic check this
 * project has on a round's own result.
 *
 * WHY IT EXISTS. Rounds 070-072 took `same scale across the deck` from 35
 * consecutive failures to three consecutive passes. Nothing protected that: a
 * later build could take it back to 3 of 8 and no gate would fail, because every
 * round result in this repo is read by a person and then filed. Three rounds of
 * host time bought a result with no guard on it.
 *
 * ESTABLISHED MEANS PASSED IN ALL OF THE LAST `window` ROUNDS, and the threshold
 * is the project's own: three is what `docs/ROUNDS.md` asks for "where a claim
 * depends on it". This is what keeps the gate off the host's mood. A scenario
 * that fails half the time was never established, so its next failure is not a
 * regression and is not reported; one that has passed three times running and
 * then fails is exactly the thing a person would want stopped at.
 *
 * A scenario absent from the older rounds is NEW, not regressed — it cannot have
 * been established, and reporting it would make every added scenario look like a
 * fault on its first bad round.
 */
export function scenarioRegressions(rounds, window = 3) {
  if (!Array.isArray(rounds) || rounds.length < window + 1) return [];
  // WITHIN ONE PROFILE ONLY. A nightly cycle runs 16:9 twice and 4:3 once, so a
  // 4:3 round judged against three 16:9 rounds would be flagged for scoring
  // differently — which it does, by design. The gate would fire every night,
  // and a gate that cries wolf gets switched off; this file has already watched
  // that happen twice. See `roundProfile`.
  const profile = roundProfile(rounds[rounds.length - 1]);
  const sameProfile = rounds.filter((r) => roundProfile(r) === profile);
  if (sameProfile.length < window + 1) return [];
  rounds = sameProfile;
  // `true` passed, `false` failed, `null` DID NOT MEASURE. The third value is
  // what the gate got wrong on its first live outing: round 073 flagged `explode
  // a degraded picture` as having stopped passing, for a result whose own words
  // were "proves nothing either way". A scenario declining to conclude is an
  // absence of evidence, not a fall, and a gate that fires on one is a gate that
  // gets switched off — which this repo has already watched happen once.
  //
  // Read from the `skipped` FLAG, never from the detail text. The prose is
  // edited whenever a message is improved, and a gate keyed to it would go quiet
  // the first time someone reworded a sentence.
  const scenariosOf = (r) => {
    const out = new Map();
    for (const sc of r?.selftest ?? []) if (sc?.name) out.set(sc.name, sc.skipped ? null : !!sc.ok);
    return out;
  };
  const all = rounds.map(scenariosOf);
  // The DETAIL text alongside, read only for the report — see `sameDetailIn`
  // below, which spells out why keying anything else on prose is forbidden.
  const detailsOf = (r) => {
    const out = new Map();
    for (const sc of r?.selftest ?? []) if (sc?.name) out.set(sc.name, String(sc.detail ?? ""));
    return out;
  };
  const details = rounds.map(detailsOf);
  const newest = all[all.length - 1];
  // The `window` rounds BEFORE the newest — the newest is what is being judged.
  const before = all.slice(-1 - window, -1);
  const out = [];
  for (const [name, ok] of newest) {
    // Passed, or did not measure — neither is a regression.
    if (ok !== false) continue;
    // Established: present AND passing in every one of the previous rounds.
    const established = before.every((r) => r.get(name) === true);
    if (!established) continue;
    // LIFETIME, not the window. `passedIn` is the window SIZE — a constant that
    // is true of every regression this function can return, so printing it as
    // "had passed the previous 3 rounds running" said nothing at all. Whether a
    // scenario has failed once in 124 rounds or twelve times is the difference
    // between a blip and a habit, and it is what decides whether to chase it.
    let ran = 0,
      failed = 0;
    for (const r of all) {
      const v = r.get(name);
      // `undefined` never ran, `null` declined to conclude. Neither is evidence.
      if (v === undefined || v === null) continue;
      ran++;
      if (!v) failed++;
    }
    /**
     * HAS THIS EXACT FAILURE HAPPENED BEFORE, and on which builds?
     *
     * "failed 8 times in 332 rounds" says the scenario is unreliable. It does
     * not say whether TONIGHT's failure is one of the familiar ones or a new
     * shape wearing an old name, and those want opposite responses.
     *
     * Round 435 is why this is here. `two slides claiming one slot` stopped
     * passing with "4 slides inserted, 3 kept, 3 of 2 still re-editable; 2
     * queued as duplicates" — which is character-for-character what rounds 060,
     * 253 and 273 said. Round 273 is the round the bundle guard below was
     * written for. Everything needed to recognise it was in the archive and
     * nothing put it in front of the reader.
     *
     * A REPORT FIELD, NOT A GATE INPUT, and the difference is load-bearing.
     * `scenariosOf` above reads the `skipped` FLAG and never the prose,
     * precisely so that rewording a message cannot make the gate go quiet. This
     * keeps that rule: the verdict is decided before this runs, and a reworded
     * detail only makes this say "not seen before", which never suppresses
     * anything.
     */
    const detail = details[details.length - 1].get(name) ?? "";
    const sameDetailIn = [];
    for (let i = 0; i < all.length - 1; i++)
      if (all[i].get(name) === false && detail && details[i].get(name) === detail)
        sameDetailIn.push(String(rounds[i]?.build ?? "").split(" ")[0]);
    out.push({ name, passedIn: window, ran, failed, sameDetailIn });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Did the chart span sync batches, and did it group?
 *
 * THE SHARPEST SEPARATION IN THE ARCHIVE, found 2026-08-16 while chasing what
 * looked like a rasterise effect:
 *
 *     spanned batches   452 draw(s), 353 grouped = 78%
 *     one batch only    214 draw(s),  49 grouped = 23%
 *
 * And it is OURS, not the host's. `refreshShapes` is set from `spansBatches()`,
 * so only a multi-batch chart gets the pre-grouping re-read that resolves its
 * shapes by id. A single-batch chart hands `addGroup` the raw `created` proxies
 * — which this host refuses with InvalidParam 5010 — and the failed group then
 * takes the tag with it: `target.tags` comes back undefined, 155 times out of
 * 155 across the archive, every one immediately after a 5010 group.
 *
 * THE RASTERISE WAS A RED HERRING, recorded so nobody re-finds it. Draws after a
 * rasterise group 22% of the time against 27% for every other draw — nothing. It
 * looked like 22% against 93% until the arms were split by batch count, because
 * the scenarios that rasterise happen to draw small charts. That is the exact
 * confound `poolEveryDraw` already warns about, met from a different direction.
 */
export function poolBatchSpanVsGroup(logs) {
  const out = { multi: 0, multiGrouped: 0, single: 0, singleGrouped: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    entries.forEach((e, i) => {
      if (!/^batch issued/.test(String(e.message))) return;
      const d = e.data ?? {};
      const total = Number(d.total) || 0;
      // The LAST batch of a draw only, or a multi-batch chart is counted once
      // per batch and swamps the single-batch arm with copies of itself.
      if (Number(d.upTo) !== total || !total) return;
      const multi = total > (Number(d.perSync) || 10);
      // UNTIL THE NEXT DRAW, not a fixed number of entries away — and that is a
      // correction, not a refinement. This read `i + 4` until 2026-08-16, which
      // was true of the traces it was written against and stopped being true the
      // moment every groupable chart started re-reading the slide first: the
      // extra entries pushed the verdict out of the window and the report showed
      // ZERO single-batch draws, in the same change that altered the single-batch
      // path. An instrument that goes blind exactly where it is being used is
      // worse than no instrument.
      //
      // The outcome set was short too. `not grouping` — the honest decline — was
      // not counted at all, so a chart that declined looked like a chart that had
      // never been decided.
      for (let k = i + 1; k < entries.length; k++) {
        const m = String(entries[k].message);
        // The next draw begins: this one was never resolved either way.
        if (/^batch issued/.test(m)) break;
        const ok = /^grouped the chart/.test(m);
        const bad =
          /^not grouping/.test(m) ||
          (/grouping the chart/.test(m) && /5010/.test(JSON.stringify(entries[k].data ?? {})));
        if (!ok && !bad) continue;
        if (multi) {
          out.multi++;
          if (ok) out.multiGrouped++;
        } else {
          out.single++;
          if (ok) out.singleGrouped++;
        }
        break;
      }
    });
  }
  return out;
}

/** The batch-span split, which is where grouping is really decided. */
function reportBatchSpanVsGroup(logs) {
  const b = poolBatchSpanVsGroup(logs);
  if (!b.multi && !b.single) return;
  console.log(`\n  DID THE CHART SPAN BATCHES — pooled over ${logs.length} round(s)`);
  console.log(
    `    spanned batches  ${String(b.multi).padStart(4)} draw(s), ${String(b.multiGrouped).padStart(4)} grouped = ${rateOrSilence(b.multiGrouped, b.multi)}`,
  );
  console.log(
    `    one batch only   ${String(b.single).padStart(4)} draw(s), ${String(b.singleGrouped).padStart(4)} grouped = ${rateOrSilence(b.singleGrouped, b.single)}`,
  );
  console.log(
    `    OURS, not the host's. refreshShapes USED to be set from spansBatches(), so only a\n` +
      `    multi-batch chart got the pre-grouping re-read that resolves its shapes by id; a\n` +
      `    single-batch chart handed addGroup the raw created proxies, which this host refuses\n` +
      `    with 5010, and the failed group took the tag with it (target.tags undefined, 155 of\n` +
      `    155). Every groupable chart asks for the refresh since 2026-08-16. That removed the\n` +
      `    doomed attempt but did NOT make those charts group — rounds 066 and 067 both show\n` +
      `    5 of 5 declining with 'not grouping'. Their re-read comes back non-empty but\n` +
      `    matching none of our ids, and the settle retry only fires on empty or partial.\n` +
      `    A rasterise beforehand looked like the cause and is not: 22% vs 27% once the arms\n` +
      `    are split by batch count. The scenarios that rasterise just draw small charts.`,
  );
}

/**
 * Questions that produced NOTHING, round after round.
 *
 * Every round pays for the whole probe battery in host time and scratch slides,
 * and a question that never answers costs exactly as much as one that does. The
 * per-round report names the ones "never put" in that round; nothing said which
 * ones have never been put in ANY round, so a permanently dead question read as
 * bad luck twelve times running.
 *
 * It was found by hand on 2026-08-16 and the answer was worth the query: SIX
 * questions were 0-for-12, and four of them are the group cluster —
 * `addgroup-returns-usable`, `group-children-via-getcount`,
 * `grouped-child-by-id-from-slide`, `tag-on-group-survives`. The probe has been
 * blind on groups for twelve rounds, and rounds 064/065 then answered the most
 * important of them FROM PRODUCTION, twice, in one evening.
 *
 * SPLIT BY KIND, because the two want opposite fixes and pooling them hides
 * that:
 *
 *   never asked   `no-scratch-slide` / `no-scratch-shape` — the harness could
 *                 not set the question up. A HARNESS problem, ours to fix.
 *   unanswerable  `unreadable` — the question was put and the host would not
 *                 answer. A HOST fact, and a real (if annoying) finding.
 *
 * A question in the first group for many rounds should be moved into production
 * instrumentation or retired; the second is telling you something about the host
 * and should be left alone.
 */
export function poolStarvedQuestions(logs) {
  const NEVER = /^no-scratch/;
  // EVERY WAY THE HOST DECLINES, not just the one word.
  //
  // This was `/^unreadable/`, and everything else counted as an ANSWER — so a
  // probe that has never once answered stayed out of a report whose entire job
  // is to find probes that never answer, purely because its silence used a
  // different prefix. Measured over the last 40 rounds, 0 of 120 asks each:
  //
  //   no-creation-id   103x  the shape has no creationId, so the question about
  //                          creationId surviving anything cannot be put
  //   no-group-id      514x  "the host would not name the group it just made"
  //   no-refusal       486x  the probe's OWN comment: "the host grouped today,
  //                          so the question was never put. Not an answer."
  //
  // That last one is the sharpest: the probe documented the value as not an
  // answer and the reader of that value counted it as one.
  //
  // A probe that answers SOMETIMES still stays out of the report — the filter
  // below needs `answered === 0` — so widening this cannot bury a working
  // question, only surface a silent one.
  const UNANSWERABLE = /^(unreadable|no-creation-id|no-group-id|no-refusal)/;
  // WHICH QUESTIONS THE BUILD STILL ASKS. This report tells the reader to fix
  // or retire something, and for eight rounds it named two questions that had
  // ALREADY BEEN RETIRED — `grouped-child-by-id-from-slide` and
  // `tag-on-group-survives`, dropped on 2026-08-21 and last seen in round 149,
  // still listed at "125 round(s)" and still filed under OURS TO FIX. A pooled
  // count over the whole archive cannot tell a live starving probe from a dead
  // one, and a report that demands action on work already done is the same
  // defect as a conclusion hardcoded into an instrument: it keeps printing
  // after it stops being true.
  //
  // A WINDOW, not the newest round alone. A single sheet can be short because
  // the host died mid-probe, which would read a live question as retired; three
  // non-empty sheets is enough that a genuinely live probe appears in one, and
  // few enough that a retirement shows up within a round or two of landing.
  const recent = logs.filter((l) => (l?.hostAnswers?.answers ?? []).length).slice(-3);
  const live = new Set(recent.flatMap((l) => (l.hostAnswers.answers ?? []).map((a) => a?.id)));
  const seen = new Map();
  for (const log of logs) {
    for (const a of log?.hostAnswers?.answers ?? []) {
      if (!a?.id) continue;
      const v = a.answer == null ? "(none)" : String(a.answer);
      const t = seen.get(a.id) ?? { rounds: 0, never: 0, unanswerable: 0, answered: 0, last: "" };
      t.rounds++;
      if (NEVER.test(v)) {
        t.never++;
        t.last = v;
      } else if (UNANSWERABLE.test(v)) {
        t.unanswerable++;
        t.last = v;
      } else t.answered++;
      seen.set(a.id, t);
    }
  }
  // Only the ones that have NEVER produced an answer. A question that answers
  // sometimes is doing its job and does not belong in a report about waste.
  return (
    [...seen.entries()]
      .filter(([, t]) => t.answered === 0 && t.rounds > 1)
      // `retired` rather than dropped: the archive still holds the rounds these
      // starved in, and a reader comparing an old report to this one deserves to
      // see WHY a row moved rather than find it simply gone.
      .map(([id, t]) => ({ id, ...t, retired: !live.has(id) }))
      .sort((a, b) => b.rounds - a.rounds || a.id.localeCompare(b.id))
  );
}

/** The dead questions, named, so a starved probe stops reading as bad luck. */
/**
 * What PRODUCTION has seen of the question `shape-resolve-held-slide-proxy`
 * cannot ask.
 *
 * That probe has answered `no-scratch-shape` in all 133 archived rounds and
 * structurally cannot do better: it needs an id for a freshly added shape and
 * this host refuses to give one. The one production site that still resolves a
 * shape by id through a slide handle a sync old is `deleteShapesById`, so this
 * pools what that sweep saw.
 *
 * BOTH DIRECTIONS, and that is the point. The sweep used to trace only its
 * failures, so the archive held 133 rounds of silence that could mean either
 * "the host resolved everything" or "the sweep never ran" — and in fact it
 * means the second, which nothing could say until the positive line existed.
 */
export function poolAgedHandleResolves(logs) {
  let resolved = 0;
  let refused = 0;
  let rounds = 0;
  for (const log of logs ?? []) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    let any = false;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m === "resolved a shape by id through a slide handle a sync old") {
        resolved += Number(e.data?.resolved ?? 0) || 0;
        any = true;
      } else if (m === "wreckage the host would not resolve") {
        refused += Number(e.data?.unresolved ?? 0) || 0;
        any = true;
      }
    }
    if (any) rounds++;
  }
  return { resolved, refused, rounds, of: (logs ?? []).length };
}

/**
 * How badly the host renumbers, printed so the paragraph in `powerpoint.ts`
 * stops being the only account of it.
 */
function reportIdChurn(logs) {
  const c = poolIdChurn(logs);
  if (!c.kinds.length) return;
  const total = c.kinds.reduce((n, k) => n + k.events, 0);
  console.log(`
  IDS THAT LEAVE THE LIST WHEN A SLIDE IS APPENDED — ${total} event(s) over ${c.rounds} round(s)`);
  for (const k of c.kinds)
    console.log(
      `    ${String(k.events).padStart(4)} event(s) read ${k.fresh} id(s) as new` +
        (k.medianDeck === undefined ? "" : `, median deck ${k.medianDeck} slide(s) before the add`),
    );
  console.log(
    `    One id has to leave the list for "grew by 1, ${c.kinds[0]?.fresh ?? 2} read as new" to add up. The
` +
      `    comment in powerpoint.ts was written from SEVEN hand-collected observations before this
` +
      `    archive existed and said the count was always two. It is not, and the split is by deck
` +
      `    size — a hand sample cannot notice that its own conclusion is a function of the deck.`,
  );
}

/** Why a chart left the draw without a parts list — the four counters, read. */
function reportPartsListOutcome(logs) {
  const o = poolPartsListOutcome(logs);
  if (!o.events) return;
  console.log(`
  WHY A CHART HAD NO PARTS LIST — ${o.events} event(s) over ${o.rounds} round(s)`);
  console.log(
    `    ${String(o.grouped).padStart(4)}  grouped, so not loose — correct by design, a group needs no parts list`,
  );
  console.log(`    ${String(o.noCharts).padStart(4)}  no charts in the call at all`);
  console.log(`    ${String(o.threw).padStart(4)}  the id read-back THREW — the real failure`);
  console.log(`    ${String(o.builtNothing).padStart(4)}  reached the normal exit and still produced nothing`);
  console.log(
    `    \`gotPartsList\` is 0 on every event, and taken alone that reads as a path that has never once
` +
      `    worked. It is not: ${o.grouped + o.noCharts} of ${o.events} are cases where a parts list is not WANTED. The counter
` +
      `    that separates them is \`where\`, and it has been in the same object since round 142.`,
  );
}

/** Whether the positional guess grouped this chart's shapes or another's. */
/**
 * The REAL occupancy reading against what an update cost.
 *
 * `WHAT AN IN-PLACE UPDATE COSTS` says, in the archive's own words, that "there
 * is no reliable occupancy measure in a round" and falls back to `onSlide`,
 * which counts what THIS RUN drew and reports the same slide as 10 and as 42.
 *
 * There has been a real one since round 239. `same scale across the deck` calls
 * `slideOccupancy` before it touches a chart and traces `what each slide held
 * before the rescale`, ordered as the charts are — the host's own count, one
 * call, outside the timed path. Nothing read it. The renderer recorded the
 * answer for 29 rounds while the reader went on saying none existed.
 *
 * BUCKETED BY `of/changed`, and this section is the reason to insist on it. Read
 * unbucketed, the archive says charts 2 and 3 are the CHEAPEST in the run at
 * ~12.7s against ~18.4s for charts 4-8 — which reads as "a busier slide is
 * faster" and is nothing of the kind. Charts 2 and 3 are 16-node charts and the
 * rest are 24-node. The confound the surrounding section warns about swallowed
 * the result whole on the first pass.
 */
/**
 * WHERE THE HOST DIES, pooled — because nothing has ever read this directory.
 *
 * `keepCrashedRun` has been filing every unfinished round under `crashes/` for
 * weeks and no tool opens them. Fourteen sat there on 2026-08-27, read one at a
 * time by hand, and reading them one at a time is exactly why nobody saw that
 * the three most recent all stopped at the SAME step — a step none of the
 * eleven before them stopped at:
 *
 *     probe  re-asked what the empty deck could not answer
 *
 * Older crashes die mid-draw at ~400 steps (`parts list outcome`, `batch
 * issued`); those three die at 535-540 in the probe's end-of-phase cleanup,
 * across two builds and both slide sizes. That is a signature, and finding it
 * took an afternoon of hand-reading that this function does in a second.
 *
 * KEYED ON CHANNEL AND MESSAGE ONLY. The step line carries a timestamp and a
 * data blob, both of which differ every run; grouping on the raw line would put
 * every crash in its own bucket and report nothing.
 */
export function poolCrashLastSteps(crashes) {
  const groups = new Map();
  for (const c of crashes ?? []) {
    const steps = Array.isArray(c?.steps) ? c.steps : [];
    if (!steps.length) continue;
    const key = crashStepKey(steps[steps.length - 1]);
    if (!groups.has(key)) groups.set(key, { key, n: 0, at: [], steps: [] });
    const g = groups.get(key);
    g.n++;
    g.at.push(c.name);
    g.steps.push(steps.length);
  }
  return [...groups.values()].sort((a, b) => b.n - a.n || b.at[0].localeCompare(a.at[0]));
}

/** Channel and message, with the timing and the data blob stripped off. */
export function crashStepKey(line) {
  const withoutTime = String(line ?? "").replace(/^\s*[\d.]+s\s+/, "");
  const fields = withoutTime.split(/\s{2,}/).filter(Boolean);
  return fields.slice(0, 2).join("  ").slice(0, 72) || "(no steps recorded)";
}

/**
 * Which scenario was in flight when the host died — the FOURTH OUTCOME CLASS.
 *
 * A round file records three states per scenario: passed, failed, skipped. It
 * cannot record the fourth, because a scenario that KILLS THE HOST produces no
 * verdict and the round it belonged to files nothing. That is structural, not
 * sampling, and it has hidden the single worst finding in this archive:
 *
 *     same scale across the deck — 0 failures in 282 recorded verdicts,
 *     joint-safest in the suite, and the scenario the host died inside TEN
 *     times.
 *
 * A reader of `rounds/` alone ranks it perfectly safe. `rounds-salvaged/`
 * does not help either: a salvaged round carries VERDICTS, and a scenario that
 * killed the host has none, so it is simply absent there too. The only place
 * this fact exists is the crash record's own steps.
 *
 * ATTRIBUTION IS DELIBERATELY NARROW. A scenario counts as fatal only when its
 * `scenario starting` line has no matching `scenario passed|failed` after it —
 * it was open when the record ends. 26 of 83 sound records attribute this way;
 * the other 57 died in the probe phase or the deck-evidence scan and are
 * credited to NOTHING, which is correct and is why this returns counts rather
 * than a rate. A death with a scenario open is not proof that the scenario
 * caused it, and the budget in `host-baseline.mjs` is what turns the counts
 * into a judgement.
 *
 * Pure: takes parsed records so it can be tested without a filesystem, the
 * same split `host-baseline.mjs` keeps for the pane's sake.
 */
export function fatalScenarios(crashes) {
  const deaths = {};
  /** Scenario name → the crash-record filenames that credited it. */
  const credits = {};
  let attributed = 0;
  let unattributed = 0;
  for (const crash of crashes ?? []) {
    const steps = crash?.steps;
    if (!Array.isArray(steps) || !steps.length) continue;
    let open = null;
    for (const step of steps) {
      const started = /scenario starting\s+name=(.+?)\s+(?:deckSlides|atMs)=/.exec(String(step));
      if (started) {
        open = started[1];
        continue;
      }
      // CLOSED ONLY BY ITS OWN NAME. A nested or follow-up scenario line
      // closing someone else's would credit the death to whatever ran last
      // rather than to what was running.
      //
      // ALL THREE SPELLINGS, AND ONE OF THEM SHOUTS. `selftest.ts:3962` emits
      // `scenario skipped`, `scenario passed` or `scenario FAILED` — the last
      // one capitalised so a failure is legible in a wall of trace. The first
      // version of this line matched `passed|failed`, case-sensitively, which
      // in `crashes/` matches 905 lines and misses 15: FAILED appears 5 times,
      // skipped 10, and lowercase `failed` NEVER — zero occurrences in the
      // whole archive. A scenario closing either way was left OPEN forever,
      // and the record's death was credited to it.
      //
      // That was not hypothetical. It invented three deaths and put two
      // scenarios into `FATAL_SCENARIO_BUDGET` that have never killed
      // anything: `a chart of rotated shapes` (2) and `where a rotated shape
      // lands` (1). Both are now absent from the table, which is the promise
      // `host-baseline.mjs` makes about an absent name — that its FIRST death
      // is a rise from zero and is caught. Carrying a free death defeats it.
      //
      // A skip closes the scenario too, and should: a skipped scenario ended
      // and the host outlived it. That is the opposite of the thing being
      // counted.
      const ended = /scenario (?:passed|FAILED|failed|skipped)\s+name=(.+?)\s+detail=/.exec(String(step));
      if (ended && ended[1] === open) open = null;
    }
    if (open) {
      deaths[open] = (deaths[open] ?? 0) + 1;
      // WHICH RECORD credited it, not just how many did. `deathsAcknowledged`
      // keys a receipt to one crash file, so the pooled count alone cannot say
      // whether the death a person read is the death now breaching.
      //
      // `loadCrashRecords` stamps `_file` on every record it parses. A record
      // built by hand in a test has none, so it pushes `null` and can never be
      // matched by a ledger entry — which is the property that keeps a test
      // from acknowledging anything.
      (credits[open] ??= []).push(crash._file ?? null);
      attributed++;
    } else unattributed++;
  }
  return { deaths, attributed, unattributed, credits };
}

/**
 * How many times each scenario RAN — the denominator a death count needs.
 *
 * A scenario ran if a completed round carries an entry for it (passed, failed
 * or skipped all mean it was attempted), or if a crash record shows it opening.
 * Both halves are needed: rounds alone miss every run that ended in a crash,
 * which is precisely the population deaths come from.
 */
export function scenarioRuns(rounds, crashes) {
  const runs = {};
  for (const r of rounds ?? []) {
    for (const s of r?.selftest ?? []) if (s?.name) runs[s.name] = (runs[s.name] ?? 0) + 1;
  }
  for (const c of crashes ?? []) {
    for (const step of c?.steps ?? []) {
      const m = /scenario starting\s+name=(.+?)\s+(?:deckSlides|atMs)=/.exec(String(step));
      if (m) runs[m[1]] = (runs[m[1]] ?? 0) + 1;
    }
  }
  return runs;
}

/**
 * Scenarios killing the host MORE OFTEN than the archive says they do.
 *
 * A RATE, in deaths per 1000 runs, and the change from a raw count is the whole
 * point of this function.
 *
 * The count version shipped on 2026-09-03 and was red within hours. It could not
 * have been anything else. Deaths only ever accumulate — `crashes/` is
 * append-only — so a budget seeded at today's total breaches on the very next
 * death, and "the next death" is ordinary operation for a scenario that has been
 * dying all week. It fired on the episode continuing, not on a regression, and a
 * gate that cries wolf gets switched off. `docs/BACKLOG.md` records that
 * happening to an earlier one.
 *
 * The count also RANKED WRONG, because it conflated exposure with danger.
 * Measured against how often each scenario actually runs:
 *
 *     a big chart on a slide of its own    5 deaths /  11 runs = 455
 *     stop a run mid-draw                  8 deaths /  27 runs = 296
 *     same scale across the deck           9 deaths / 421 runs =  21
 *     every other listed scenario          1 death  / ~420     =   2.4
 *
 * `same scale` led on raw count and is a nine-times-baseline scenario. The two
 * that kill the host on a THIRD to a HALF of their runs looked smaller only
 * because they are young. The ranking was upside down.
 *
 * And a rate CAN FALL, which is the property this instrument needed and a
 * cumulative count can never have: runs keep accumulating, so a scenario that
 * stops dying sinks on its own and the fix is protected without anyone editing
 * a number. That is what `overlap-budget` does, and why the analogy finally
 * holds.
 *
 * A scenario absent from the ceilings has never killed the host, so its first
 * death is a rise from zero and is reported. That is still the case most worth
 * catching, and it is the one thing the count version got right.
 */
/**
 * The most deaths a scenario may show before the rise is more than noise.
 *
 * A RATE SEEDED AT THE POINT ESTIMATE IS STILL A HAIR TRIGGER, which is the
 * trap the count version fell into and the reason this is not simply
 * `rate > ceiling`. `stop a run mid-draw` sits at 8 deaths in 25 runs; a ninth
 * takes it to 346 per 1000 against a ceiling of 330 — red, on a scenario doing
 * exactly what it has done all week. Killing the host a third of the time
 * PRODUCES eight-and-nine-death stretches. That is the baseline, not a rise.
 *
 * So the gate asks the question a person would: is this more than a scenario at
 * the accepted rate would produce anyway? Normal approximation to the binomial,
 * two standard deviations:
 *
 *     allowed = p*n + 2*sqrt(p*(1-p)*n)      p = ceiling/1000, n = runs
 *
 * A ceiling of 0 gives p=0, sd=0 and allowed=0, so a scenario that has never
 * killed the host trips on its FIRST death. That case needs no statistics and
 * does not get any.
 *
 * THE COST IS HONEST AND WORTH STATING: at eleven runs this cannot tell 45% from
 * 80%. `a big chart on a slide of its own` would need nine deaths in eleven runs
 * to trip. That is a property of eleven samples, not a choice — a tighter bound
 * would not be more sensitive, it would just be wrong more often. The printed
 * table is the surveillance; the gate is for what the evidence can carry.
 */
export function fatalDeathsAllowed(ceilingPer1000, runs) {
  const p = Math.max(0, Math.min(1, (ceilingPer1000 ?? 0) / 1000));
  const n = Math.max(0, runs ?? 0);
  if (p === 0 || n === 0) return 0;
  return p * n + 2 * Math.sqrt(p * (1 - p) * n);
}

export function fatalRateBreaches(deaths, runs, ceilings) {
  const over = [];
  for (const [name, count] of Object.entries(deaths ?? {})) {
    const ran = runs?.[name] ?? 0;
    // No denominator means no rate. A death with no recorded run is a record
    // this tool cannot read, and inventing a rate of Infinity for it would fail
    // the gate on a parsing gap rather than on a host.
    if (!ran) continue;
    const allowed = ceilings?.[name] ?? 0;
    const cap = fatalDeathsAllowed(allowed, ran);
    if (count > cap) {
      over.push({ name, count, runs: ran, rate: (1000 * count) / ran, allowed, cap });
    }
  }
  return over.sort((a, b) => b.rate - a.rate);
}

/**
 * Which breaches a person has signed for, and which still want one.
 *
 * THE PROBLEM THIS EXISTS FOR. A scenario absent from `FATAL_SCENARIO_RATE` has
 * a ceiling of 0, and `fatalDeathsAllowed` returns 0 at p=0 for every
 * denominator. So its first host death breaches, and — because the crash record
 * stays in `crashes/` — it breaches forever. The gate meanwhile prints "falls on
 * its own as runs pile up… NOTHING NEEDS EDITING to make it green", which is
 * true at every seeded ceiling and false at zero. The nightly cycle stopped
 * after one round every night from round 428 on, and the 4:3 validation leg,
 * which is leg three, stopped running at all.
 *
 * WHAT THIS CLEARS AND WHAT IT REFUSES TO. It clears the gate's EXIT, never the
 * count. The breach still prints, the death still counts for ever, no ceiling
 * moves, and `FATAL_SCENARIO_RATE` is not touched. The distinction is the whole
 * design: a receipt says a person read one crash record, where a raised ceiling
 * says every future death of that shape is pre-authorised.
 *
 *   allowed > 0            ALWAYS STANDING. A rate that can fall has a green
 *                          path already and must never be signed away — that is
 *                          the case the table exists for, and letting a receipt
 *                          touch it would turn this into the ceiling edit its
 *                          own docstring forbids.
 *   allowed === 0, one     CLEARABLE, by naming the exact crash record. One
 *                          death, one receipt.
 *   allowed === 0, several NEVER CLEARABLE. A second death is the archive's own
 *                          signal that the first was not a one-off: across nine
 *                          scenarios that have died, every repeat death landed
 *                          within 25 rounds of its predecessor, median 1.
 *
 * STALE IS FATAL, and it is the guard that keeps a ledger from becoming a
 * silencer. An entry whose record no longer credits that scenario, or whose
 * scenario has since been given a real ceiling, is not a receipt any more — it
 * is a line nobody re-read. The gate exits 2 on it, which is the code for "the
 * gate could not judge this", not the code for a regression.
 */
export function deathsAcknowledged(breaches, credits, acknowledged, ceilings) {
  const standing = [];
  const cleared = [];
  const stale = [];
  const ledger = acknowledged ?? [];
  for (const b of breaches ?? []) {
    const records = credits?.[b.name] ?? [];
    if (b.allowed > 0 || b.count !== 1) {
      standing.push(b);
      continue;
    }
    // The one record that credited it. `records` and `count` agree by
    // construction — both are written by the same loop in `fatalScenarios` —
    // so a disagreement means the two were computed over different pools and
    // the receipt must not be honoured.
    const record = records.length === 1 ? records[0] : null;
    const signed = record && ledger.find((e) => e.record === record && e.scenario === b.name);
    if (signed) cleared.push({ ...b, receipt: signed });
    else standing.push(b);
  }
  for (const e of ledger) {
    const records = credits?.[e.scenario] ?? [];
    const ceiling = ceilings?.[e.scenario] ?? 0;
    if (!records.includes(e.record))
      stale.push({ ...e, why: `no crash record named ${e.record} credits \`${e.scenario}\` any more` });
    else if (ceiling > 0)
      stale.push({
        ...e,
        why: `\`${e.scenario}\` now carries a ceiling of ${ceiling}, so it has a green path without this`,
      });
  }
  return { standing, cleared, stale };
}

/**
 * Scenario pass/fail rates over a population, WITH A PER-SCENARIO DENOMINATOR.
 *
 * The denominator is the whole design. A crashed round ran scenarios 1..k and
 * never reached k+1..n, so pooling its verdicts against a global count of rounds
 * reads every scenario it never got to as a silent success — late scenarios look
 * rarer and safer, early ones more common, and each one is measured against a
 * different truth. That is why `rounds-salvaged/` has sat unread rather than
 * simply being concatenated onto `rounds/`.
 *
 * So a scenario's denominator is the number of rounds that RECORDED A VERDICT
 * for it, and nothing else. A round that never reached it contributes to
 * neither side. `salvagedPartial.reached` says how far a salvaged round got and
 * is reported alongside, so a reader can see how much of the population is
 * partial before trusting a rate built on it.
 *
 * Returns counts rather than percentages on purpose: 1 of 2 and 50 of 100 are
 * not the same claim and a percentage hides which one you have.
 */
export function scenarioRatesOver(rounds) {
  const per = {};
  let partial = 0;
  for (const r of rounds ?? []) {
    if (r?.salvagedPartial) partial++;
    for (const s of r?.selftest ?? []) {
      if (!s?.name || s.skipped) continue;
      const row = (per[s.name] ??= { ran: 0, failed: 0 });
      row.ran++;
      if (!s.ok) row.failed++;
    }
  }
  return { per, rounds: (rounds ?? []).length, partial };
}

/**
 * The one line every number over a mixed population must carry.
 *
 * A rate pooled from completed AND salvaged rounds is a different claim from one
 * pooled from completed rounds alone, and a reader cannot tell which they are
 * holding unless it says so. This repo has already paid for an unlabelled
 * population once — the 4:3 crash rate that turned out to be about a document —
 * so the label travels with the number rather than beside it.
 */
export function populationLine(complete, salvaged, partial) {
  const bits = [`${complete} complete round(s)`];
  if (salvaged) bits.push(`${salvaged} salvaged (${partial} of them partial)`);
  return bits.join(" + ");
}

function reportCrashes(dir = "crashes", list = readdirSync, read = readFileSync) {
  let files;
  try {
    files = list(dir).filter((f) => /-crashed-run\.json$/.test(f));
  } catch {
    return; // no crashes directory is a good state, not an error
  }
  if (!files.length) return;
  const crashes = [];
  for (const f of files) {
    try {
      const s = JSON.parse(read(`${dir}/${f}`, "utf8"));
      crashes.push({ name: f.slice(0, 19), steps: s.steps ?? [], build: String(s.build ?? "").slice(0, 7) });
    } catch {
      /* a truncated crash file is not worth taking the report down for */
    }
  }
  const groups = poolCrashLastSteps(crashes);
  if (!groups.length) return;
  console.log(`
  WHERE THE HOST DIED — ${crashes.length} crashed run(s) kept, pooled by their LAST step`);
  for (const g of groups.slice(0, 8)) {
    const span = `${Math.min(...g.steps)}-${Math.max(...g.steps)}`;
    console.log(`    ${String(g.n).padStart(3)}x  steps ${span.padEnd(9)} ${g.key}`);
    if (g.n > 1) console.log(`         ${g.at.slice(-4).join("  ")}`);
  }
  console.log("    A crash is the HOST falling over, not a scenario failing — none of these rounds");
  console.log("    reached a verdict. What makes them worth pooling is that a repeated last step is");
  console.log("    a place the host is reliably unhappy, and one crash never says that.");
  console.log("    THE SLIDE SIZE IS NOT HERE because a crashed run does not record it. Classifying");
  console.log("    these by profile still needs the cycle log, which is the next gap to close.");
}
export function poolOccupancyCost(logs) {
  const cells = new Map();
  let rounds = 0;
  for (const log of logs ?? []) {
    let occ = null;
    const ups = [];
    for (const e of log?.trace?.entries ?? []) {
      if (e.message === "what each slide held before the rescale" && Array.isArray(e.data?.slides)) occ = e.data.slides;
      const d = e.data;
      if (!d || typeof d.ms !== "number" || !/^\d+\/\d+$/.test(String(d.chart ?? ""))) continue;
      if (!/in.place|updated/i.test(String(e.message ?? ""))) continue;
      // BOTH HALVES OF `chart`, because "1/8" and "1/1" are not the same thing
      // and the first version of this read only the numerator. `one chart alone
      // on a warm deck` is a chart 1/1 — its own scenario, and the archive
      // already reports it as "alone in its own run: 18689ms, nearer a LATER
      // chart". Labelling it `first` pools the arm that ISOLATES the effect with
      // the arm that exhibits it, which is the confound this section exists to
      // keep apart.
      const [pos, total] = String(d.chart).split("/").map(Number);
      ups.push({ n: pos, total, ms: d.ms, of: d.of, changed: d.changed });
    }
    if (!occ || !ups.length) continue;
    rounds++;
    for (const u of ups) {
      const slide = occ[u.n - 1];
      if (!slide || typeof slide.shapes !== "number") continue;
      const where = u.total === 1 ? "alone" : u.n === 1 ? "first" : "later";
      const key = `${u.of}/${u.changed}|${slide.shapes}|${where}`;
      if (!cells.has(key)) cells.set(key, []);
      cells.get(key).push(u.ms);
    }
  }
  return { rounds, cells };
}

/** Report it, with every cell's n, because the interesting cells are thin. */
function reportOccupancyCost(logs) {
  const { rounds, cells } = poolOccupancyCost(logs);
  if (!rounds) return;
  const med = (a) => [...a].sort((x, y) => x - y)[Math.floor(a.length / 2)];
  console.log(`
  WHAT THE SLIDE ACTUALLY HELD, against what the update cost — ${rounds} round(s)`);
  const keys = [...cells.keys()].sort();
  for (const k of keys) {
    const [bucket, shapes, where] = k.split("|");
    const v = cells.get(k);
    console.log(
      `    ${bucket.padEnd(8)} slide held ${String(shapes).padStart(3)}  ${where.padEnd(5)}` +
        `  n=${String(v.length).padStart(3)}  ${med(v)}ms`,
    );
  }
  console.log("    The host's own count, taken by the scenario before the run — NOT `onSlide`, which");
  console.log("    counts what this run drew and reads the same slide as 10 and as 42.");
  console.log("    THE CLEAN COMPARISON is `slide held 3`, where first and alone are equally sampled:");
  console.log("    the first chart of an 8-chart run costs ~44.1s and THE SAME CHART ALONE in its own");
  console.log("    run costs ~17.6s, n=26 each, at one chart size, one changed count, and the SAME");
  console.log("    occupancy. Later charts sit at ~20.2s. What this DOES settle is that occupancy");
  console.log("    is not the difference — it is identical across the two rows.");
  console.log("    It does NOT establish that position is the cause: `alone` is a different");
  console.log("    SCENARIO, later in the round on a warmer deck, so it differs from `first` in");
  console.log("    more than run length. It narrows the candidates; it does not pick one.");
  console.log("    OCCUPANCY IS STILL NOT ESTABLISHED as a cause: the 1-shape and 17-shape cells are");
  console.log("    n=2 and do not move monotonically. Settling that needs a deck built to vary");
  console.log("    occupancy, which this one is not — every round builds the same slides.");
  console.log("    `alone` is `one chart alone on a warm deck`, chart 1/1. It is NOT a first chart:");
  console.log("    pooling it with one would hide the very effect these two rows measure.");
}
function reportPositionalGuess(logs) {
  const g = poolPositionalGuess(logs);
  if (!g.events) return;
  console.log(`
  THE POSITIONAL GUESS — ${g.events} event(s) over ${g.rounds} round(s)`);
  console.log(`    ${String(g.mine).padStart(4)}  picked this chart's OWN shapes`);
  console.log(`    ${String(g.other).padStart(4)}  picked ANOTHER chart's shapes entirely`);
  console.log(`    ${String(g.partial).padStart(4)}  picked a mixture`);
  console.log(`    ${String(g.unreadable).padStart(4)}  the host would not name what it listed`);
  console.log(
    `    ${String(g.refused).padStart(4)}  REFUSED by the guard — named no shape of ours, so nothing was grouped`,
  );
  if (g.other || g.partial)
    console.log(
      `    A guess that picks another chart's shapes GROUPS THEM under this chart's name and
` +
        `    feeds them to its parts tag. It throws nothing when those shapes are still loose,
` +
        `    which is why this line exists — the 29 archived addGroup throws are the same guess
` +
        `    on shapes that had already been absorbed into a group.`,
    );
  // WHAT ACTUALLY HAPPENED, not just what was risked. The paragraph above
  // describes the danger; on its own it reads as though the danger was realised.
  if ((g.other || g.partial) && g.refused >= g.other + g.partial)
    console.log(
      `    EVERY foreign pick above was refused — none was grouped. positionalGuessNamesOurs
` + `    is doing its job, and this section is the record of it working, not of it failing.`,
    );
  else if (g.other || g.partial)
    console.log(
      `    ${g.other + g.partial - g.refused} foreign pick(s) were NOT refused. Check whether they predate the guard
` + `    (round 179 does) before reading them as live corruption.`,
    );
}

/**
 * Which target route the config tag was written through — both halves.
 *
 * ONLY THE FAILURES WERE EVER COUNTED. `tagging failed` carries `from`, and
 * pooled over the archive it reads created 235, undefined 61, group 51, by-id
 * 26. That looks like a verdict on the creation proxy and is not one: a route
 * chosen ten times as often fails ten times as often at the same rate, and
 * nothing counted how often each was chosen.
 *
 * It is not academic. The probe sheet answers
 * `tag-the-creation-proxy-a-sync-later => yes` 148 of 148, so on that evidence
 * the creation proxy is the route to PREFER — while the raw failure counts say
 * the opposite. One of those is measured on a scratch shape and one has no
 * denominator, and a change made on either would be made blind.
 */
export function poolTagRoutes(logs) {
  const ok = {};
  const failed = {};
  let rounds = 0;
  for (const log of logs ?? []) {
    let any = false;
    for (const e of log?.trace?.entries ?? []) {
      const m = String(e.message ?? "");
      if (m === "config tags written, by target route") {
        for (const [k, n] of Object.entries(e.data?.from ?? {})) ok[k] = (ok[k] ?? 0) + (Number(n) || 0);
        any = true;
      } else if (/^tagging failed/.test(m)) {
        const k = String(e.data?.from ?? "undefined");
        failed[k] = (failed[k] ?? 0) + (Number(e.data?.charts) || 1);
        any = true;
      }
    }
    if (any) rounds++;
  }
  const routes = [...new Set([...Object.keys(ok), ...Object.keys(failed)])].sort();
  return { rounds, routes: routes.map((r) => ({ route: r, ok: ok[r] ?? 0, failed: failed[r] ?? 0 })) };
}

/** The tag routes, with the denominator that was missing. */
function reportTagRoutes(logs) {
  const t = poolTagRoutes(logs);
  if (!t.routes.length) return;
  console.log(`
  CONFIG TAG BY TARGET ROUTE — over ${t.rounds} round(s)`);
  for (const r of t.routes) {
    const n = r.ok + r.failed;
    const rate = r.ok ? `${Math.round((100 * r.ok) / n)}% written` : n ? "0% written" : "";
    console.log(
      `    ${r.route.padEnd(12)} ${String(r.ok).padStart(5)} written · ${String(r.failed).padStart(4)} failed  ${rate}`,
    );
  }
  if (t.routes.every((r) => !r.ok))
    console.log(
      `    NO SUCCESS COUNTS YET — every round above predates the line that records them, so these
` + `    are failures without a denominator and no route can be compared to another on them.`,
    );
}

/**
 * What the settle retry buys, split by which ask it was.
 *
 * THE RETRY IS LOAD-BEARING AND NOTHING MEASURED IT. Pooled over 161 rounds
 * before #709: 1,057 retry passes fired and 97 failures survived them — a 90.8%
 * rescue rate. The first read of the shape collection fails on roughly half the
 * charts this loop sees, and a pause plus a fresh slide handle fixes nine in
 * ten. The 4.3% quoted all over this project is what is LEFT after the retry,
 * not the failure rate of the read.
 *
 * #709 raised `REREAD_ATTEMPTS` from 1 to 2 to reach the 97. Whether that does
 * anything is only visible if the trace says WHICH ASK a pass was — five passes
 * in a round is either five first asks from five update calls, or four firsts
 * and a second, and those are opposite readings of the same number. Rounds
 * before the `attempt` field cannot be told apart and are counted as
 * `unlabelled` rather than folded in.
 */
/**
 * What a batch of shapes costs on this host, split by whether the slide is the
 * one the user is LOOKING at.
 *
 * `batch issued` has carried `prevBatchMs` for the whole archive — 2,913 samples
 * across 169 rounds — and nothing had ever pooled it. It is the only instrument
 * that prices the product's biggest latency cost: a chart redraw takes about
 * twenty seconds, and the scenario that redraws eight of them is 38% of a round.
 *
 * READ `upTo` CAREFULLY. It is a RUNNING TOTAL, not a batch size, and the first
 * cut of this reader treated it as a size:
 *
 *     {upTo: 10, total: 16}                    <- ten sent so far
 *     {upTo: 16, total: 16, prevBatchMs: 9731} <- six more; 9731ms timed the FIRST
 *
 * Bucketing by `upTo` produced a tidy table showing 10-shape and 20-shape batches
 * costing the same, and therefore that payload was free and the batch size
 * should go up. Both buckets were batches of ten; the second was simply later in
 * a sequence. **There is no size variation anywhere in the archive** — every
 * batch is ten — so this data cannot compare batch sizes and must not be used
 * to. An aggregate over a misread field is clean, confident, and empty.
 *
 * `prevBatchMs` times the batch described by the PRECEDING line, so the size it
 * belongs to is that line's delta, not this one's.
 */
/**
 * Numeric trace fields that NEVER VARY across the whole archive.
 *
 * Every serious error this project has made is one shape: an unmeasured thing
 * printed as a measurement. A field that is always 0 is the commonest form of
 * it, and nothing has ever detected the class automatically — each one was
 * caught by eye, late, and only when the number happened to look absurd.
 *
 * The case that produced this reader: a commit put `contextSyncs` on the
 * in-place update's trace line to ask whether the first chart of a run issues
 * more syncs or slower ones. Round 202 answered 0 for eight charts costing
 * 12-37 seconds each. A 37-second update that issues no syncs is not a
 * finding, it is a broken gauge — and it was visible in the very first round
 * that carried it.
 *
 * WHAT THIS CANNOT DO, said plainly: it cannot tell a broken gauge from a
 * thing that genuinely never happens. `partial: 0` on every grouped chart means
 * no chart was ever partial, which is a real and welcome answer.
 * `contextSyncs: 0` on a 37s update means the counter is blind. **Both look
 * identical here.** The report says so rather than ranking them, because
 * ranking them would be the same defect one level up.
 *
 * It flags CONSTANTS too, not only zeros. A field reporting the same non-zero
 * number in every round is a hardcoded conclusion wearing a measurement's
 * clothes, and this repo has shipped one of those before.
 */
export function poolFlatFields(logs, minSamples = 20) {
  const seen = new Map();
  for (const log of logs ?? []) {
    for (const e of log?.trace?.entries ?? []) {
      const message = String(e?.message ?? "");
      const data = e?.data;
      if (!data || typeof data !== "object") continue;
      for (const [field, value] of Object.entries(data)) {
        if (typeof value !== "number" || !Number.isFinite(value)) continue;
        const key = message + " :: " + field;
        let row = seen.get(key);
        if (!row) {
          row = { message, field, n: 0, only: value, varies: false };
          seen.set(key, row);
        }
        row.n++;
        if (value !== row.only) row.varies = true;
      }
    }
  }
  const flat = [...seen.values()].filter((r) => !r.varies && r.n >= minSamples);
  return {
    zeros: flat.filter((r) => r.only === 0).sort((a, b) => b.n - a.n),
    constants: flat.filter((r) => r.only !== 0).sort((a, b) => b.n - a.n),
    fieldsSeen: seen.size,
  };
}

/** Numeric fields that never move — suspect gauges, or genuinely quiet facts. */
function reportFlatFields(logs) {
  const f = poolFlatFields(logs);
  if (!f.zeros.length && !f.constants.length) return;
  console.log("\n  NUMERIC TRACE FIELDS THAT NEVER VARY — of " + f.fieldsSeen + " seen");
  const show = (rows, what) => {
    if (!rows.length) return;
    console.log("    " + what);
    for (const r of rows.slice(0, 12))
      console.log(
        "      " + (r.message.slice(0, 44) + " :: " + r.field).padEnd(64) + String(r.n).padStart(5) + "x  = " + r.only,
      );
    if (rows.length > 12) console.log("      … and " + (rows.length - 12) + " more");
  };
  show(f.zeros, "always 0:");
  show(f.constants, "always the same non-zero value:");
  console.log("    A broken gauge and a thing that never happens look IDENTICAL here, and this");
  console.log("    cannot tell them apart. Check any field whose zero would be surprising — a");
  console.log("    sync count of 0 on a 37-second update is what made this reader exist.");
  // THE VALUE IS PRINTED because without it every reader re-derives the same
  // thing. Working out that `charts` is constant because A BATCH HOLDS ONE CHART
  // — not because the counter is stuck — took six queries against the archive on
  // 2026-08-26, and the answer was one number this line already had in hand.
  console.log("    Most constants here are structural: a batch holds ONE chart, so `charts = 1`");
  console.log("    and a per-item `index = 0` are expected. `applying delete :: index` varies,");
  console.log("    which is what a field indexing something genuinely plural looks like.");
}
/**
 * What the IN-PLACE UPDATE path costs — the product's largest latency number,
 * and the one the batch pooler below cannot see.
 *
 * `same scale across the deck` is 166s median, 38% of a round, and issues not one
 * `batch issued` line: eight charts, all through `tryInPlaceUpdate`. The batch
 * timings therefore price the DRAW path and miss this entirely. `ms` arrived on
 * the update line in 05a27fd; round 198 is the first to carry it, and its
 * updates account for 156.5s of that scenario's 166s.
 *
 * BUCKETED BY `of`, BECAUSE CHART SIZE IS A CONFOUND. A 16-node chart changing 9
 * shapes and a 24-node chart changing 18 are different work, and pooling them
 * gives a per-shape figure describing neither. Every comparison here is made
 * within one `of`, and the first-chart one within one `changed` as well.
 */
/**
 * How many shapes THIS RUN put on each slide, from the run's own batch lines.
 *
 * Joined rather than measured, and that is the point: asking the host for a
 * slide's shape count inside the timed update would add a sync to the path being
 * timed, which is the "adding counting sites moves the baseline" hazard this
 * archive has been bitten by. `batch issued` already carries `onSlide` per
 * `onSlideKey`, written while drawing, so the occupancy is already in every
 * round — it had simply never been joined to the update rows.
 *
 * A FLOOR, not an occupancy: it counts what the run drew, so a slide that
 * arrived with content reads low. Adequate here because every slide in the
 * self-test deck starts empty.
 */
export function priorDrawsOnSlide(entries) {
  const per = new Map();
  for (const e of entries ?? []) {
    const d = e?.data ?? {};
    if (String(e?.message ?? "") !== "batch issued") continue;
    if (typeof d.onSlideKey !== "string" || typeof d.onSlide !== "number") continue;
    per.set(d.onSlideKey, Math.max(per.get(d.onSlideKey) ?? 0, d.onSlide));
  }
  return per;
}

export function poolUpdateCost(logs) {
  const rows = [];
  for (const log of logs ?? []) {
    const drawnBefore = priorDrawsOnSlide(log?.trace?.entries);
    for (const e of log?.trace?.entries ?? []) {
      if (String(e.message ?? "") !== "updated only the shapes that changed") continue;
      const d = e.data ?? {};
      if (typeof d.ms !== "number" || typeof d.changed !== "number" || typeof d.of !== "number") continue;
      // `chart` is "i/n" style. It USED to be present only when a run updated
      // several, and `first` was `/^1\//` on that assumption — which `one chart
      // alone on a warm deck` falsifies by design: it emits "1/1", and under the
      // old test that row would have been pooled into the FIRST-CHART population
      // it exists to be contrasted against. The two are the whole question.
      const inRun = typeof d.chart === "string";
      const [at, runLen] = inRun ? String(d.chart).split("/").map(Number) : [];
      const alone = inRun && runLen === 1;
      rows.push({
        ms: d.ms,
        changed: d.changed,
        of: d.of,
        inRun,
        // First OF SEVERAL. A run of one has no "rest" to be worse than.
        first: inRun && at === 1 && runLen > 1,
        alone,
        slideId: typeof d.slideId === "string" ? d.slideId : null,
        // WHAT THE FIRST-CHART COST TURNED OUT TO BE. Rounds 213-215: a chart
        // updated alone on the crowded slide cost 35033 and 39570ms, and the
        // same arm on an uncrowded slide cost 15235ms — later-chart cost. The
        // slide tracks it and the position does not.
        drawnThereBefore: typeof d.slideId === "string" ? (drawnBefore.get(d.slideId) ?? null) : null,
      });
    }
  }
  const median = (a) => {
    if (!a.length) return null;
    const b = [...a].sort((x, y) => x - y);
    return b[Math.floor(b.length / 2)];
  };
  const sizes = [...new Set(rows.map((r) => r.of))].sort((a, b) => a - b);
  const fits = [];
  for (const of of sizes) {
    // ALONE ROWS ARE OUT OF THE FIT, for the reason the comment below gives
    // about first rows: if a run of one turns out to carry the first-chart cost,
    // letting it into the baseline is the outlier defining itself away. No
    // archived round emits one, so every fit computed before this line existed
    // is unchanged by it.
    const here = rows.filter((r) => r.of === of && !r.first && !r.alone);
    const changes = [...new Set(here.map((r) => r.changed))].sort((a, b) => a - b);
    if (changes.length < 2) continue;
    const lo = changes[0];
    const hi = changes[changes.length - 1];
    const loMs = median(here.filter((r) => r.changed === lo).map((r) => r.ms));
    const hiMs = median(here.filter((r) => r.changed === hi).map((r) => r.ms));
    const perShape = (hiMs - loMs) / (hi - lo);
    fits.push({ of, lo, hi, perShape, fixed: loMs - perShape * lo, n: here.length });
  }
  const firstVsRest = [];
  for (const of of sizes) {
    const changedHere = [...new Set(rows.filter((r) => r.of === of).map((r) => r.changed))];
    for (const changed of changedHere) {
      const same = (r) => r.of === of && r.changed === changed;
      const f = rows.filter((r) => same(r) && r.first).map((r) => r.ms);
      const rest = rows.filter((r) => same(r) && r.inRun && !r.first && !r.alone).map((r) => r.ms);
      // The third arm, and the reason this scenario exists: a chart that is the
      // first of its run AND the only one in it. If it costs what a first chart
      // costs, the run is the unit and the cost is set-up the first chart
      // absorbs. If it costs what a later chart costs, the cost needs the rest
      // of the queue to exist.
      const alone = rows.filter((r) => same(r) && r.alone).map((r) => r.ms);
      if (!f.length || !rest.length) continue;
      firstVsRest.push({
        of,
        changed,
        firstN: f.length,
        restN: rest.length,
        first: median(f),
        rest: median(rest),
        aloneN: alone.length,
        alone: median(alone),
      });
    }
  }
  // EXCESS ABOVE THE FIT, which is the measurement that actually discriminates.
  //
  // Raw ms cannot: the first chart of the run changes 18 of 24 while the next two
  // change 9 of 16, so comparing them measures the work. The fit is built from
  // NON-FIRST rows only — otherwise the outlier sets its own baseline and
  // defines itself away — and every row is then scored against it.
  const fitFor = (of) => fits.find((f) => f.of === of);
  const excessOf = (r) => {
    const f = fitFor(r.of);
    return f ? r.ms - (f.fixed + f.perShape * r.changed) : null;
  };
  const firstRows = rows.filter((r) => r.first);
  const restRows = rows.filter((r) => r.inRun && !r.first);
  const firstExcess = firstRows.map(excessOf).filter((x) => x != null);
  const restExcess = restRows.map(excessOf).filter((x) => x != null);
  // AND THE SLIDE, HELD CONSTANT — the control, when the deck provides one.
  //
  // If the first chart's slide also carries later charts, those are same slide,
  // same run, later position. Round 200 is what this was written for: charts 1,
  // 2 and 3 all on slide 257, the first +19631ms above fit and the other two
  // +2768 and +3207. That is the whole argument that position, not the slide,
  // is the cause — and it was nearly missed, because the entry above this one
  // predicted the two would be permanently confounded on the assumption of one
  // chart per slide.
  const firstSlides = new Set(firstRows.map((r) => r.slideId).filter(Boolean));
  const sameSlideLater = restRows.filter((r) => r.slideId && firstSlides.has(r.slideId));
  const sameSlideExcess = sameSlideLater.map(excessOf).filter((x) => x != null);
  // COUNTED, NOT DROPPED. A same-slide chart with no fit for its size scores
  // null and used to vanish inside this filter — so the report announced
  // "no later chart shares the first one's slide, so slide and position are
  // still tied" while round 200 held two of them, on slide 257, and settled
  // the question.
  //
  // They are unscorable because a fit needs two distinct `changed` values for a
  // size, and every 16-node chart in the archive changes 9. That is a real
  // limit, and it is not the same fact as "there is no control". Silence from
  // a filter is not evidence of absence.
  const sameSlideUnscorable = sameSlideLater.length - sameSlideExcess.length;
  return {
    n: rows.length,
    // The rows themselves, so the occupancy split can be taken without a second
    // pass over every log — see `occupancySplit`.
    rows,
    sizes,
    fits,
    firstVsRest,
    firstExcess: median(firstExcess),
    firstExcessN: firstExcess.length,
    restExcess: median(restExcess),
    restExcessN: restExcess.length,
    sameSlideExcess: median(sameSlideExcess),
    sameSlideN: sameSlideExcess.length,
    sameSlideUnscorable,
    sameSlideRaw: sameSlideLater.map((r) => ({ ms: r.ms, changed: r.changed, of: r.of })),
  };
}

/**
 * What this project believes, checked against what the archive says.
 *
 * See `scripts/claims.mjs`. A STALE line is the most interesting thing in this
 * report — it means a stated fact stopped being true and nothing else would have
 * said so. On 2026-08-25 the fresh-slide failure had been fixed for an unknown
 * number of rounds while the figure describing it went on being quoted.
 */
function reportClaims(pooled) {
  const rows = checkClaims(pooled);
  if (!rows.length) return;
  console.log(`\n  WHAT THIS PROJECT BELIEVES — ${rows.length} claim(s), checked against the archive`);
  for (const r of rows) {
    // `staleIsGood` inverts the alarm, not the check: "no chart has ever carried
    // a parts list" going stale is the parts list starting to WORK.
    const mark = r.ok === null ? "  ?  " : r.ok ? " ok  " : r.staleIsGood ? "NEWS " : "STALE";
    console.log(`    ${mark} ${r.id.padEnd(26)} ${r.actual}`);
    if (r.ok === false) {
      console.log(`          claim: ${r.says}`);
      console.log(`          measured ${r.measured}`);
      console.log(`          The archive no longer agrees. Re-read what cites this before quoting it.`);
    }
  }
  console.log(`    A stale claim is not a failure — it is a fact that moved while nobody was looking.`);
}

/** Print the dormant instruments, loudest first, or say nothing when there are none. */
/**
 * Do draws stall more on a slide that already holds things?
 *
 * The question the scenario table raises and cannot answer, asked with the
 * per-draw covariate instead of the scenario as a proxy: `batch issued` carries
 * `onSlide`, the occupancy of the target when the batch went out.
 *
 * UNRESOLVED ROWS ARE EXCLUDED, and that is not tidying — it reverses the
 * answer. A chart's FIRST batch is issued before the slide has been named, and
 * carries `onSlide: 0, onSlideKey: "(visible)"`. Counted as an empty slide, 8 of
 * the 9 apparent empty-slide stalls in this archive are those rows, and the
 * table reads "empty slides stall MOST" — the exact opposite of what the
 * resolved rows say. This project renamed `priorDrawsOnSlide` for being a bad
 * occupancy proxy and then used it as one twice; this is the same proxy again.
 *
 * WHAT IT SAYS TODAY, whole archive, resolved rows only:
 *
 *     0 (empty)   1/1320  0.08%
 *     1-10        4/1338  0.30%
 *     11-25       4/1567  0.26%
 *     26-50       3/1613  0.19%
 *     51+         3/368   0.82%
 *
 * Fifteen stalls across six thousand draws. The trend is upward at both ends and
 * the buckets cannot be told apart — 1/1320 and 3/368 have thoroughly
 * overlapping intervals — so this prints counts and REFUSES the trend. It is
 * here to accumulate: the denominators grow every round, and the day the
 * separation is real the table will say so without anyone re-deriving it.
 */
export function stallsByOccupancy(logs, minPerBucket = 30) {
  const bucketOf = (n) => (n === 0 ? "0 (empty)" : n <= 10 ? "1-10" : n <= 25 ? "11-25" : n <= 50 ? "26-50" : "51+");
  const per = new Map();
  let unresolved = 0;
  for (const log of logs ?? []) {
    const es = log?.trace?.entries ?? [];
    const at = [];
    es.forEach((e, i) => {
      if (e.scope === "draw" && e.message === "batch issued") at.push(i);
    });
    at.forEach((i, k) => {
      const d = es[i].data ?? {};
      if (typeof d.onSlide !== "number") return;
      // See above: no slide key means the batch went out before the slide was
      // named, so its `onSlide` is a placeholder rather than a count.
      if (!d.onSlideKey || d.onSlideKey === "(visible)") {
        unresolved++;
        return;
      }
      const end = k + 1 < at.length ? at[k + 1] : es.length;
      let stalled = false;
      for (let z = i + 1; z < end; z++) {
        const e = es[z];
        if (e.scope === "host" && e.message === "gave up waiting" && /drawing shapes/.test(String(e.data?.what ?? "")))
          stalled = true;
      }
      const b = bucketOf(d.onSlide);
      if (!per.has(b)) per.set(b, { draws: 0, stalls: 0 });
      const row = per.get(b);
      row.draws++;
      if (stalled) row.stalls++;
    });
  }
  const order = ["0 (empty)", "1-10", "11-25", "26-50", "51+"];
  const rows = order
    .filter((b) => per.has(b) && per.get(b).draws >= minPerBucket)
    .map((b) => ({ bucket: b, ...per.get(b), pct: (100 * per.get(b).stalls) / per.get(b).draws }));
  // The bar for saying the buckets differ: every bucket needs enough STALLS to
  // be worth a rate, not merely enough draws. Ten is generous and still not met.
  const separable = rows.length > 1 && rows.every((r) => r.stalls >= 10);
  return { rows, unresolved, separable };
}

/** Print the occupancy table, and say plainly that it cannot yet answer. */
function reportOccupancyStalls(pooled) {
  const { rows, unresolved, separable } = stallsByOccupancy(pooled);
  if (rows.length < 2) return;
  console.log(`\n  DOES A BUSY SLIDE STALL A DRAW? — ${unresolved} unresolved row(s) held out, see below`);
  for (const r of rows)
    console.log(
      `    ${r.bucket.padEnd(11)}${String(r.stalls).padStart(3)}/${String(r.draws).padEnd(5)} ${r.pct.toFixed(2).padStart(5)}%`,
    );
  if (!separable) console.log("    NOT SEPARABLE at these counts — printed to accumulate, not to be read as a trend.");
  console.log("    A chart's FIRST batch goes out before its slide is named and carries `onSlide: 0`.");
  console.log("    Counting those as empty slides reverses the answer — 8 of 9 apparent empty-slide stalls");
  console.log("    are unresolved rows. Same bad proxy this project renamed once and then reused twice.");
}

/**
 * Which scenarios lose draws to a stall, and how often.
 *
 * The battery reports 14 of 14 nearly every round, so it reads as quiet. It is
 * not: across the last 60 rounds two scenarios skipped 7 times each, and a SKIP
 * means nothing was checked. The reason is in their skip text — "the host did
 * not finish a draw made while a shape was SELECTED", "the host stopped
 * answering during this scenario" — and the rate behind it was never pooled.
 *
 * Measured 2026-08-25 over 60 rounds:
 *
 *     a selected shape survives an insert     7 / 112   6.25%
 *     insert onto a slide that already has…   7 / 227   3.08%
 *     does a rasterise poison the next draw   1 / 240   0.42%
 *     explode a degraded picture              0 / 171   0%
 *     edit the chart the user selected        0 / 150   0%
 *
 * THE PATTERN IS NOT THE HEADLINE, AND THIS IS WHY. The two that stall INSERT
 * onto a slide that already holds something; the two that never stall UPDATE a
 * chart that is already there. Selection is the obvious story and it is one of
 * at least three — insert-versus-update, and how full the target slide is, sit
 * on top of it, and this archive has confounded position with load four times
 * already. So this reports rates and names the confound rather than a cause.
 *
 * What would settle it is an insert onto a busy slide with NOTHING selected,
 * against the same insert with a shape selected. Nothing in the battery does
 * that today.
 *
 * The denominator is `batch issued` in draw scope, which is the only per-draw
 * line that fires on SUCCESS. `drawing the chart's shapes` looks like a draw
 * counter and is a `boundedSync` label — it appears only when the sync fails, so
 * counting it gives every scenario a 100% stall rate. I did exactly that on the
 * first attempt.
 */
export function stallsByScenario(logs, minDraws = 20) {
  const per = new Map();
  for (const log of logs ?? []) {
    let current = "(outside any scenario)";
    for (const e of log?.trace?.entries ?? []) {
      if (/^scenario starting/.test(String(e.message ?? ""))) current = String(e.data?.name ?? "?");
      const isDraw = e.scope === "draw" && e.message === "batch issued";
      const isStall =
        e.scope === "host" && e.message === "gave up waiting" && /drawing shapes/.test(String(e.data?.what ?? ""));
      if (!isDraw && !isStall) continue;
      if (!per.has(current)) per.set(current, { draws: 0, stalls: 0 });
      const row = per.get(current);
      if (isDraw) row.draws++;
      else row.stalls++;
    }
  }
  return [...per.entries()]
    .filter(([, v]) => v.draws >= minDraws)
    .map(([name, v]) => ({ name, ...v, pct: (100 * v.stalls) / v.draws }))
    .sort((a, b) => b.pct - a.pct);
}

/** Print the stall rates, with the confound attached to them. */
function reportStalls(pooled) {
  const rows = stallsByScenario(pooled);
  if (!rows.length || !rows.some((r) => r.stalls)) return;
  console.log(`\n  WHICH SCENARIOS LOSE DRAWS — a stall means the scenario SKIPPED and checked nothing`);
  for (const r of rows.slice(0, 6))
    console.log(
      `    ${r.pct.toFixed(2).padStart(5)}%  ${String(r.stalls).padStart(3)}/${String(r.draws).padEnd(5)} ${r.name.slice(0, 44)}`,
    );
  console.log("    The two that stall INSERT onto an occupied slide; the two that never stall UPDATE a chart");
  console.log("    already there. Selection is the obvious story and it is one of at least three — this");
  console.log("    archive has confounded position with load four times. Rates here, not a cause.");
}

/**
 * A round's shape: how much of it is probe, how much battery, how much neither.
 *
 * Reconstructed by hand three times on 2026-08-25 to answer "why does a round
 * take fourteen minutes", and each reconstruction cost the same twenty minutes
 * and produced two wrong leads before the right accounting:
 *
 *   - a 37-second gap that looked like a wait and was the first chart's own
 *     update, the 2.2x this archive already claims;
 *   - 56 seconds before the first probe question that looked like a slow deck
 *     read and is the harness's lead-in — the entry carries
 *     `observedBeforeTheRound`, so it was captured before the round's trace
 *     window and replayed into it.
 *
 * Neither was a product cost and both were nearly "optimised". A standing
 * breakdown is cheaper than rediscovering the same two non-findings.
 *
 * OUTSIDE is deliberately not called overhead. It holds the lead-in, the deck
 * evidence, the rasterises and the sweep — real work, some of it the harness's
 * own, and naming it overhead would invite someone to cut it without looking.
 */
export function roundAnatomy(logs, minN = 5) {
  const rows = [];
  for (const log of logs ?? []) {
    const entries = log?.trace?.entries ?? [];
    if (!entries.length) continue;
    const span = entries[entries.length - 1]?.ms;
    if (typeof span !== "number" || span <= 0) continue;
    const probe = entries
      .filter((e) => e.message === "answered" || e.message === "partner answered")
      .reduce((a, e) => a + (e.data?.ms ?? 0), 0);
    const battery = (Array.isArray(log.selftest) ? log.selftest : []).reduce((a, s) => a + (s?.ms ?? 0), 0);
    rows.push({ span, probe, battery, outside: Math.max(0, span - probe - battery) });
  }
  if (rows.length < minN) return null;
  const med = (key) => {
    const s = rows.map((r) => r[key]).sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
  };
  return { n: rows.length, span: med("span"), probe: med("probe"), battery: med("battery"), outside: med("outside") };
}

/** Print the round's shape, so "rounds feel slow" has somewhere to start. */
function reportRoundAnatomy(pooled) {
  const a = roundAnatomy(pooled);
  if (!a) return;
  const pct = (v) => (a.span ? Math.round((100 * v) / a.span) : 0);
  console.log(`\n  WHAT A ROUND IS MADE OF — medians over ${a.n} round(s), ${Math.round(a.span / 1000)}s end to end`);
  console.log(
    `    battery  ${String(Math.round(a.battery / 1000)).padStart(4)}s  ${String(pct(a.battery)).padStart(2)}%   the scenarios, priced one by one above`,
  );
  console.log(
    `    probe    ${String(Math.round(a.probe / 1000)).padStart(4)}s  ${String(pct(a.probe)).padStart(2)}%   the question sheet`,
  );
  console.log(
    `    outside  ${String(Math.round(a.outside / 1000)).padStart(4)}s  ${String(pct(a.outside)).padStart(2)}%   lead-in, deck evidence, rasterises, the sweep`,
  );
  console.log("    `outside` is NOT overhead — it is the harness's own work plus the evidence a round exists to");
  console.log("    collect. Two things in it have already been mistaken for waste and were not.");
}

/**
 * What each scenario costs, pooled — the half of a round nothing had priced.
 *
 * The probe half has been measured all day and is now 60 seconds cheaper. The
 * BATTERY is the larger half and nothing has ever aggregated it: round 255 spent
 * 418s across 14 scenarios, and `same scale across the deck` alone took 133s of
 * that — a third of the battery in one scenario.
 *
 * A median per scenario is the useful shape rather than a total. Totals move
 * with how many scenarios ran; a median says what THIS scenario costs when it
 * runs, which is the number that tells you where a round's time goes and the one
 * that moves if a scenario starts drifting.
 *
 * Skips are excluded. A skipped scenario records the time it took to give up,
 * and pooling that with real runs drags the median toward a number describing
 * neither — the same mistake `selfTestHeadline` exists to prevent one level up.
 */
export function scenarioCost(logs, minN = 5) {
  const per = new Map();
  for (const log of logs ?? []) {
    for (const s of Array.isArray(log?.selftest) ? log.selftest : []) {
      if (!s?.name || typeof s.ms !== "number") continue;
      // Only scenarios that RAN to a verdict. `skipped` means nothing was
      // checked, and its duration is the cost of the giving up.
      if (!s.ok && s.skipped) continue;
      if (!per.has(s.name)) per.set(s.name, []);
      per.get(s.name).push(s.ms);
    }
  }
  const rows = [];
  for (const [name, all] of per) {
    if (all.length < minN) continue;
    const s = [...all].sort((a, b) => a - b);
    rows.push({
      name,
      n: s.length,
      median: s[Math.floor(s.length / 2)],
      q1: s[Math.floor(s.length * 0.25)],
      q3: s[Math.floor(s.length * 0.75)],
    });
  }
  const total = rows.reduce((a, r) => a + r.median, 0);
  return rows
    .map((r) => ({ ...r, share: total ? Math.round((100 * r.median) / total) : 0 }))
    .sort((a, b) => b.median - a.median);
}

/** Print where the battery's time goes, loudest first. */
function reportScenarioCost(pooled) {
  const rows = scenarioCost(pooled);
  if (!rows.length) return;
  const total = rows.reduce((a, r) => a + r.median, 0);
  console.log(
    `\n  WHERE THE BATTERY'S TIME GOES — median of ${Math.round(total / 1000)}s across ${rows.length} scenario(s)`,
  );
  for (const r of rows.slice(0, 6))
    console.log(
      `    ${String(Math.round(r.median / 1000)).padStart(4)}s  ${String(r.share).padStart(2)}%  ` +
        `${r.name.slice(0, 40).padEnd(42)}IQR ${Math.round(r.q1 / 1000)}-${Math.round(r.q3 / 1000)}s  n=${r.n}`,
    );
  console.log("    Medians, not totals: a total moves with how many scenarios ran, a median says what THIS");
  console.log("    one costs when it runs. Skips are held out — their duration is the cost of giving up.");
}

/**
 * What the probe run spends on slides it could not keep.
 *
 * `replaced the scratch slide` fires about 78 times a round and nothing has ever
 * counted it. Each one is a slide ADD on a host where that takes 0.2-4.0s, and
 * the retry after it usually succeeds — so the sheet shows a `no-scratch-slide`
 * answer roughly once in ten rounds while the run pays for it eighty times a
 * round. A cost that only appears in the raw trace is a cost nobody is deciding
 * about.
 *
 * Split by CAUSE, because the two want opposite fixes. `gone` means the host
 * said `isNullObject: true` and the slide really is deleted; `silent` means it
 * populated nothing, which is the same unanswered-proxy state the update path
 * recovers from by re-reading the collection rather than adding anything —
 * 105 of 105 on this host. `unrecorded` is every round before the cause was
 * carried out of the catch, and it must not be read as a third behaviour.
 */
export function scratchChurn(logs) {
  const perRound = [];
  const why = new Map();
  for (const log of logs ?? []) {
    let n = 0;
    for (const e of log?.trace?.entries ?? []) {
      if (!/^replaced the scratch slide/.test(String(e.message ?? ""))) continue;
      n++;
      const k = String(e.data?.why ?? "unrecorded");
      why.set(k, (why.get(k) ?? 0) + 1);
    }
    if (n) perRound.push(n);
  }
  perRound.sort((a, b) => a - b);
  return {
    rounds: perRound.length,
    total: perRound.reduce((a, b) => a + b, 0),
    median: perRound.length ? perRound[Math.floor(perRound.length / 2)] : null,
    min: perRound[0] ?? null,
    max: perRound[perRound.length - 1] ?? null,
    why: [...why.entries()].sort((a, b) => b[1] - a[1]),
  };
}

/** Print the scratch-slide churn, which is the run's own cost rather than the host's. */
function reportScratchChurn(pooled) {
  const c = scratchChurn(pooled);
  if (!c.rounds) return;
  console.log(
    `\n  WHAT THE PROBE SPENDS ON SLIDES IT COULD NOT KEEP — ${c.total} replacement(s) over ${c.rounds} round(s)`,
  );
  console.log(`    per round: min ${c.min}, median ${c.median}, max ${c.max} — every one of them a slide ADD`);
  for (const [k, n] of c.why) console.log(`      ${String(n).padStart(5)}  ${k}`);
  console.log("    The retry after a replacement usually succeeds, so the SHEET shows this about once in ten");
  console.log("    rounds while the run pays for it eighty times a round. `silent` is the same unanswered-proxy");
  console.log("    state the update path recovers from WITHOUT adding a slide — see the re-read claim.");
}

/** Print where the sheet's headline answer is a pass-1 answer that pass 1 makes special. */
function reportPassBias(pooled) {
  const rows = passPositionBias(pooled);
  if (!rows.length) return;
  console.log(`\n  THE SHEET REPORTS PASS 1, AND PASS 1 IS NOT LIKE THE OTHERS — ${rows.length} answer(s) shifted`);
  // One line per QUESTION, taking its largest shift: both halves of a two-answer
  // split appear in the rows and printing both says the same thing twice.
  const seen = new Set();
  for (const r of rows) {
    if (seen.has(r.id)) continue;
    seen.add(r.id);
    console.log(
      `    ${r.id.slice(0, 36).padEnd(38)}${r.answer.slice(0, 16).padEnd(17)}pass 1 ${r.p1.toFixed(0)}%  later ${r.later.toFixed(0)}%   n=${r.n1}/${r.n2}`,
    );
  }
  console.log("    `record` keeps the FIRST real answer, on purpose — so a sheet means today what it meant");
  console.log("    yesterday. The cost is that every headline answer is a COLD one, and cold is measurably");
  console.log("    different: collection reads are worse cold, positional slide reads are worse warm.");
  console.log("    This is the shape of the HARNESS, not of PowerPoint. Read a single answer accordingly.");
}

function reportDormant(pooled) {
  const rows = dormantInstruments(pooled);
  if (!rows.length) return;
  console.log(`\n  INSTRUMENTS THAT HAVE GONE QUIET — ${rows.length} trace line(s), none in 40+ rounds`);
  for (const r of rows.slice(0, 8))
    console.log(
      `    ${String(r.events).padStart(5)}x  last round ${String(r.lastRound).padStart(3)} (${r.gap} ago)  ${r.message.slice(0, 58)}`,
    );
  if (rows.length > 8) console.log(`    … and ${rows.length - 8} more`);
  // THE CAUSES FIRST, because they mean opposite things and the loudest rows are
  // not the interesting ones. A renamed line reads as a fixed fault and is a
  // live instrument; a removed one reads as a fixed fault and is nothing at all.
  const odd = rows.filter((r) => r.cause === "renamed" || r.cause === "removed");
  for (const r of odd) {
    console.log(
      `    ${String(r.events).padStart(5)}x  ${r.cause.toUpperCase()} — ${
        r.cause === "renamed"
          ? "the message was rewritten or is now interpolated; the instrument is ALIVE under another name"
          : "nothing in the source can emit this, so its zero says nothing about the host"
      }`,
    );
    console.log(`             ${r.message.slice(0, 70)}`);
  }
  console.log(
    `    ${rows.length - odd.length} more are still in the source verbatim — for those a zero means the fault stopped.`,
  );
  console.log("    Thread 3 waited on a counter that had been silent for 160 rounds and nothing said so.");
}

/**
 * THE NOISE FLOOR — how far apart two rounds can be with NOTHING changed.
 *
 * Every "is this a change or is this weather?" judgement in this archive has
 * rested on one line: `cabb357 scored 1 and 5 for tags-undefined with NOTHING
 * changed between them`. That is one build run twice — a range of two
 * observations — and it has been load-bearing for 200 rounds.
 *
 * IT CANNOT BE MEASURED BY RUNNING ROUNDS BACK TO BACK. Ten of those on
 * 2026-08-24 showed the median in-place update roughly DOUBLING across a
 * session (19321ms at 0 minutes in, 40067ms at 224), with scenarios starting to
 * skip past ~90 minutes. Pooling them measures the trend, not the floor. The
 * archive's own note half-saw this — "the second run is usually the worse one,
 * so a floor measured this way includes an effect as well as noise" — without
 * finding the cause.
 *
 * So the floor needs rounds at a CONSTANT session position: the FIRST round of
 * each session, which `sessionIndex === 1` now states rather than leaves to be
 * trusted. Rounds from any other position are excluded — not weighted, not
 * caveated, excluded, because a sample from minute 116 is measuring something
 * else entirely.
 *
 * Reports the SPREAD and not the mean. The question a floor answers is "how far
 * apart can two identical rounds be", and a mean cannot answer it.
 */
export function poolNoiseFloor(logs) {
  const per = new Map();
  for (const log of logs ?? []) {
    const dr = log?.driverRun;
    // Position 1 ONLY. Everything else is a different measurement wearing the
    // same name.
    if (!dr || dr.sessionIndex !== 1) continue;
    const build = String(log.build ?? "?").split(" ")[0];
    const entries = log?.trace?.entries ?? [];
    const later = [];
    for (const e of entries) {
      if (String(e.message ?? "") !== "updated only the shapes that changed") continue;
      const d = e.data ?? {};
      if (!d.chart || d.changed !== 18 || d.of !== 24 || typeof d.ms !== "number") continue;
      const [i, n] = String(d.chart).split("/").map(Number);
      if (i > 1 && n > 1) later.push(d.ms);
    }
    const selftest = Array.isArray(log.selftest) ? log.selftest : [];
    const row = {
      round: String(log.roundName ?? "").slice(0, 3),
      // The most repeated measurement this harness makes, and the one the
      // session drift moves most.
      laterMed: later.length ? [...later].sort((a, b) => a - b)[Math.floor(later.length / 2)] : null,
      skipped: selftest.filter((s) => !s.ok && s.skipped).length,
      failed: selftest.filter((s) => !s.ok && !s.skipped).length,
      lengthMs: entries.length ? entries[entries.length - 1].ms : null,
    };
    if (!per.has(build)) per.set(build, []);
    per.get(build).push(row);
  }
  return per;
}

/**
 * HOW MUCH THIS ROUND AGREED WITH ITSELF — the error bar on its own comparisons.
 *
 * The host's speed is a property of the ROUND, not of the chart. Across 32
 * rounds carrying four or more timed later charts:
 *
 *     WITHIN a round   median spread   8%   (range 1-51%)
 *     BETWEEN rounds                 164%   (between RESTED rounds, ~47%)
 *
 * Which means a comparison made INSIDE one round understates the real variance
 * by roughly six times against rested rounds, and twenty-one against the archive
 * as a whole. Every "chart A beat chart B this round" conclusion carries an error
 * bar six times wider than it looks, and nothing in this report has ever said so.
 *
 * The first-chart cost was read exactly that way for ten rounds. It survived only
 * because the effect was 2.2x — far outside even the between-round spread. A
 * smaller effect read the same way would have been noise in a finding's clothes.
 *
 * So the round's own agreement is printed beside the comparisons it licenses. It
 * is computed entirely inside the round and needs no baseline, which also makes
 * it a second symptom of a sick round: within-round spread GROWS with slowness —
 * 2% at 19.4s, 24% at 25.8s, 51% at the archive's worst.
 */
export function poolWithinRoundSpread(logs) {
  const spreads = [];
  for (const log of logs ?? []) {
    const later = [];
    for (const e of log?.trace?.entries ?? []) {
      if (String(e.message ?? "") !== "updated only the shapes that changed") continue;
      const d = e.data ?? {};
      if (!d.chart || d.changed !== 18 || d.of !== 24 || typeof d.ms !== "number") continue;
      const [i, n] = String(d.chart).split("/").map(Number);
      if (i > 1 && n > 1) later.push(d.ms);
    }
    // FOUR IS THE FLOOR FOR A SPREAD. Two charts give a range with no idea
    // whether either is typical, and a round that only managed one later chart
    // is telling a different story entirely.
    if (later.length < 4) continue;
    const s = [...later].sort((a, b) => a - b);
    spreads.push({
      round: String(log.roundName ?? "").slice(0, 3),
      pct: Math.round((100 * (s[s.length - 1] - s[0])) / s[0]),
    });
  }
  return spreads;
}

/**
 * WHERE AN UPDATE'S TIME WENT, sync by sync.
 *
 * `syncMs` is archived on every timed update and **`triage.mjs` has never read
 * it.** It is the field that settled the first-chart question — and that table
 * was rebuilt by hand, twice, from ad-hoc scripts, because the tool everyone
 * reads a round with could not show it.
 *
 * What it shows, at 18 of 24 across the archive:
 *
 *     sync        first    later    ratio
 *     write 1     11642     5510     2.11x
 *     write 2     11566     5176     2.23x
 *     write 3     12472     5469     2.28x
 *     tag           841      644     1.31x
 *
 * Three write syncs uniformly ~2.2x and a tag sync that does not share it. That
 * is the shape that excluded a single fixed stall — a timeout would appear in
 * ONE sync, not spread evenly across three — and it is the shape that survived
 * when the cause turned out to be slide occupancy rather than run position.
 *
 * Bucketed by `changed`/`of` like everything else here, because a sync's cost is
 * a function of how many shapes it writes.
 */
export function poolSyncBreakdown(logs, of = 24, changed = 18) {
  const first = [];
  const later = [];
  for (const log of logs ?? []) {
    for (const e of log?.trace?.entries ?? []) {
      if (String(e.message ?? "") !== "updated only the shapes that changed") continue;
      const d = e.data ?? {};
      if (!d.chart || d.changed !== changed || d.of !== of || !Array.isArray(d.syncMs)) continue;
      const [i, n] = String(d.chart).split("/").map(Number);
      // A run of ONE is neither: it is the arm that separates position from
      // slide, and pooling it into either side destroys that.
      if (n === 1) continue;
      (i === 1 ? first : later).push(d.syncMs);
    }
  }
  return { first, later };
}

/** The smallest number of sessions this report will call a floor. */
export const MIN_FLOOR_N = 5;

/**
 * The floor, or a refusal to state one — same discipline as `rateOrSilence`.
 *
 * A spread over two rounds IS the thing this replaces; printing one would
 * reproduce the defect with better provenance.
 */
export function reportNoiseFloor(logs) {
  const per = poolNoiseFloor(logs);
  if (!per.size) return;
  console.log(`\n  NOISE FLOOR — first round of a session only, so the session drift is held out`);
  // THE ANSWER FIRST, before the builds that cannot give one.
  //
  // This prints a block per build, and most builds have exactly one qualifying
  // round — so the section opens with several "too few to call a floor" lines
  // and buries the one build that HAS a floor somewhere in the middle. On
  // 2026-08-25 that cost an afternoon: the section was read as "the floor is
  // unmeasurable" and a four-hour plan to re-measure it was proposed against a
  // floor that had been sitting in the archive at n=9 the whole time.
  //
  // A report whose first line reads "too few" when an answer exists is not
  // merely unhelpful, it is wrong in the way a reader will act on.
  const withFloor = [...per.entries()].filter(([, r]) => r.length >= MIN_FLOOR_N);
  const ordered = [...per.entries()].sort((x, y) => y[1].length - x[1].length);
  if (withFloor.length) {
    const [build, rows] = ordered[0];
    console.log(`    BEST AVAILABLE: ${build} over ${rows.length} sessions — detailed first, thinner builds after.`);
  } else {
    const most = ordered.length ? ordered[0][1].length : 0;
    console.log(
      `    NO BUILD HAS ${MIN_FLOOR_N} FIRST-OF-SESSION ROUNDS — the best has ${most}. Nothing here is a floor.`,
    );
  }
  for (const [build, rows] of ordered) {
    const spread = (key) => {
      const vals = rows.map((r) => r[key]).filter((v) => typeof v === "number");
      if (!vals.length) return null;
      const s = [...vals].sort((a, b) => a - b);
      return {
        n: s.length,
        min: s[0],
        max: s[s.length - 1],
        med: s[Math.floor(s.length / 2)],
        q1: s[Math.floor(s.length * 0.25)],
        q3: s[Math.floor(s.length * 0.75)],
      };
    };
    const u = spread("laterMed");
    console.log(`    ${build}  ${rows.length} session(s): ${rows.map((r) => r.round).join(" ")}`);
    if (rows.length < MIN_FLOOR_N) {
      console.log(
        `      too few to call a floor (${rows.length}, needs ${MIN_FLOOR_N}) — a spread over two rounds is what this replaces`,
      );
    } else if (u) {
      const pct = u.min ? Math.round((100 * (u.max - u.min)) / u.min) : 0;
      console.log(`      in-place update, 18 of 24: ${u.min}-${u.max}ms, median ${u.med} — RANGE ${pct}% of the min`);
      // RANGE ONLY EVER GROWS WITH n, so it is a lower bound that looks like an
      // estimate. It went 66% at five sessions to 73% at eight purely because a
      // faster round arrived; nothing about the host changed. The IQR is what
      // does not drift with sample size, and it is the honest headline.
      const q = u.q1 != null && u.q3 != null ? Math.round((100 * (u.q3 - u.q1)) / u.q1) : null;
      if (q != null)
        console.log(`      middle half ${u.q1}-${u.q3}ms — IQR ${q}% of q1, and this does NOT grow with n`);
      console.log(
        `      a difference smaller than the IQR is not evidence of a change; the range is a floor on the floor`,
      );
    }
    const sk = rows.map((r) => r.skipped);
    const fa = rows.map((r) => r.failed);
    console.log(`      scenarios skipped per round: [${sk.join(" ")}]   failed: [${fa.join(" ")}]`);
  }
}

/**
 * The smallest sample this report will turn into a percentage.
 *
 * 5 of 5 is not 100%. Its 95% interval runs from roughly 48% to 100%, which is
 * most of the range the number is supposed to discriminate within — and the
 * baseline it would be read against is 1%.
 */
export const MIN_RATE_N = 20;

/**
 * A rate, or a refusal to state one.
 *
 * WHAT THIS FIXES, and it was live in the report: `freshly added, empty — 5
 * chart(s), 5 grouped = 100%`, printed directly above prose telling the reader
 * "the recent window is what to quote", against a documented baseline of 1% that
 * `BACKLOG.md` calls READ THIS FIRST. A flagship finding read as reversed on five
 * observations — and those five came from the rescale's redraw fallback, not from
 * the inserts, because the draw path carries no per-chart label for the pooling
 * to key on.
 *
 * The percentage was the wrong half to keep. `5 chart(s), 5 grouped` is the
 * measurement and it is honest at any n; `= 100%` is an inference the sample
 * cannot support. So the count stays and the rate goes silent, saying WHY rather
 * than printing a dash a reader will fill in themselves.
 */
export function rateOrSilence(count, total, minN = MIN_RATE_N) {
  if (!total) return "—";
  if (total < minN) return `too few to rate (n=${total}, needs ${minN})`;
  return `${Math.round((100 * count) / total)}%`;
}

/**
 * Trace lines this archive USED to see and has not seen for a long time.
 *
 * A zero in a report is two different facts — the fault stopped happening, or
 * nothing can produce that line any more — and the reader cannot tell them
 * apart. Thread 3 waited on `a chart's tag could not even be queued` to
 * "accumulate rounds"; the rounds accumulated and the line had last fired in
 * round 065, 160 rounds earlier. Nothing said so.
 *
 * ARCHIVE-ONLY ON PURPOSE. The obvious refinement is to grep `src/` and report
 * whether the build can still emit the line — and the first attempt at that said
 * `rasterising a slide` was still emittable when the only occurrences left are
 * three COMMENTS about it. A tool that answers "is this reachable?" by string
 * search is the same defect it exists to catch, so this one does not claim to
 * answer it. It reports the silence and names both readings.
 *
 * Counted in ROUNDS, not in events: a line that fired 155 times and then stopped
 * is the interesting shape, and an event count alone hides when it stopped.
 */
/**
 * Only what the source can actually SAY — the contents of its string literals.
 *
 * Comments are excluded, and that is the whole point rather than tidiness. This
 * codebase quotes retired trace messages in comments constantly, to record what
 * used to happen and when it stopped — `the binding could not name the chart`
 * survives only in a note reading "16, all in rounds 077-078". Searching raw
 * file text finds that note and reports the line as still emittable, which is
 * the DANGEROUS direction: it reads as "a fault that stopped" when the truth is
 * "nothing can emit this any more".
 *
 * Template literals keep their literal parts, so an interpolated message still
 * matches on its head and lands in `renamed` rather than `removed`.
 */
export function stringLiteralsIn(text) {
  const src = String(text ?? "");
  const out = [];
  let i = 0;
  // A SCANNER, not a regex sweep. Regexes cannot do this: comments here quote
  // retired messages in BACKTICKS — the house style for naming a trace line in
  // prose — so a template-literal pattern matches inside the very comments this
  // function exists to exclude, and every retired message reads as live.
  while (i < src.length) {
    const c = src[i];
    const next = src[i + 1];
    if (c === "/" && next === "/") {
      while (i < src.length && src[i] !== "\n") i++;
      continue;
    }
    if (c === "/" && next === "*") {
      i += 2;
      while (i < src.length && !(src[i] === "*" && src[i + 1] === "/")) i++;
      i += 2;
      continue;
    }
    if (c === '"' || c === "'" || c === "`") {
      const quote = c;
      i++;
      let body = "";
      while (i < src.length && src[i] !== quote) {
        // Skip the escaped character whole, so a `\"` does not end the literal.
        if (src[i] === "\\") {
          body += src[i + 1] ?? "";
          i += 2;
          continue;
        }
        body += src[i];
        i++;
      }
      i++;
      out.push(body);
      continue;
    }
    i++;
  }
  return out.join("\n");
}

/**
 * Every trace message the source can still produce, as one blob to search.
 *
 * Read from the tree rather than passed in, because the question — "can anything
 * still emit this?" — is about the code as it is now, and a caller that had to
 * supply it would be free to supply the wrong thing.
 */
export function traceSource(read = readFileSync, dir = "src") {
  const out = [];
  const walk = (d) => {
    for (const name of readdirSync(d, { withFileTypes: true })) {
      const full = join(d, name.name);
      if (name.isDirectory()) walk(full);
      else if (/\.ts$/.test(name.name)) out.push(stringLiteralsIn(read(full, "utf8")));
    }
  };
  try {
    walk(dir);
  } catch {
    // No tree to read — the caller gets an empty haystack and every line is
    // classified `unknown` rather than wrongly declared dead.
    return null;
  }
  return out.join("\n");
}

/**
 * WHY a line went quiet — which the count alone cannot say.
 *
 * Three causes, and they mean opposite things:
 *
 * - `stopped`  the message is still in the source verbatim, so something could
 *               emit it and nothing has. The fault stopped. GOOD NEWS.
 * - `renamed`  no verbatim match, but a line starting the same way exists. The
 *               instrument is ALIVE under another name, or the message is now
 *               built by interpolation and no rendered form can ever match the
 *               archive's. The zero is an artifact of the rename.
 * - `removed`  nothing resembling it in the source. Nothing can emit it, so the
 *               zero says nothing at all about the host.
 *
 * The audit that produced this found all three among the top nine, which is why
 * the report can no longer leave the question to the reader:
 * `a slide's shape count would not settle — not counting it` reads as a fault
 * that stopped 140 rounds ago and is in fact firing every round under
 * `— taking the second read`.
 *
 * The prefix is 30 characters because that is long enough to be specific and
 * short enough to survive a rewritten tail, which is where these messages get
 * edited — they are written as sentences, and the clause after the dash is the
 * part that changes.
 */
export const RENAME_PREFIX = 30;

export function dormancyCause(message, source) {
  if (source == null) return "unknown";
  if (source.includes(message)) return "stopped";
  const head = message.slice(0, RENAME_PREFIX);
  return head.length >= RENAME_PREFIX && source.includes(head) ? "renamed" : "removed";
}

/**
 * How far apart a question's first answer is from its later ones.
 *
 * THE SHEET REPORTS THE PASS-1 ANSWER. `record` keeps the first REAL answer and
 * lets nothing real displace it — deliberately, so a sheet means today what it
 * meant yesterday. The cost of that rule was never measured: if pass 1 is not
 * like passes 2 and 3, then every headline answer in the archive is a pass-1
 * answer and inherits whatever makes pass 1 special.
 *
 * It is not like them. Measured 2026-08-25 over the whole archive:
 *
 *     shape-add-positional-slide-proxy    yes 85% on pass 1, 67% later
 *     binding-names-shape-later           silent 15% on pass 1, ~0% later
 *     does-a-failed-group-poison-the-tag  tags-gone 12% on pass 1, 4% later
 *     shapes-items-via-positional-slide   short-0 9% on pass 1, 3% later
 *
 * The directions differ, which is why this reports a delta rather than a
 * verdict: the collection reads are WORSE cold and the positional slide reads
 * are worse warm. Both are facts about the harness's own shape, not about
 * PowerPoint, and neither can be read off a sheet that shows one answer.
 *
 * Never-asked answers are excluded — a question that could not be put says
 * nothing about pass position, and `no-scratch-slide` is far commoner on later
 * passes for reasons that are entirely ours.
 */
export function passPositionBias(logs, minN = MIN_RATE_N, minDelta = 5) {
  const per = new Map();
  for (const log of logs ?? []) {
    for (const row of log?.hostAnswers?.answers ?? []) {
      for (const s of row.samples ?? []) {
        if (NOT_ASKED_ANSWERS.has(s.answer)) continue;
        if (!per.has(row.id)) per.set(row.id, { p1: new Map(), later: new Map() });
        const side = per.get(row.id)[s.pass <= 1 ? "p1" : "later"];
        side.set(s.answer, (side.get(s.answer) ?? 0) + 1);
      }
    }
  }
  const share = (c, answer) => {
    const n = [...c.values()].reduce((a, b) => a + b, 0);
    return n ? { pct: (100 * (c.get(answer) ?? 0)) / n, n } : null;
  };
  const rows = [];
  for (const [id, { p1, later }] of per) {
    // Compared on the SAME answer, which is what makes the two percentages
    // commensurable. Taking each side's own top answer would compare different
    // questions and call it a difference.
    const answers = new Set([...p1.keys(), ...later.keys()]);
    for (const answer of answers) {
      const a = share(p1, answer);
      const b = share(later, answer);
      if (!a || !b || a.n < minN || b.n < minN) continue;
      const delta = a.pct - b.pct;
      if (Math.abs(delta) < minDelta) continue;
      rows.push({ id, answer, p1: a.pct, later: b.pct, delta, n1: a.n, n2: b.n });
    }
  }
  return rows.sort((x, y) => Math.abs(y.delta) - Math.abs(x.delta));
}

export function dormantInstruments(logs, minGap = 40, source = undefined) {
  const last = new Map();
  const total = new Map();
  let newest = 0;
  for (const log of logs ?? []) {
    const n = Number(String(log?.roundName ?? "").slice(0, 3));
    if (!Number.isFinite(n)) continue;
    newest = Math.max(newest, n);
    for (const e of log?.trace?.entries ?? []) {
      const m = String(e?.message ?? "");
      if (!m) continue;
      total.set(m, (total.get(m) ?? 0) + 1);
      last.set(m, Math.max(last.get(m) ?? 0, n));
    }
  }
  const rows = [];
  for (const [m, n] of last) {
    const gap = newest - n;
    if (gap > minGap)
      rows.push({
        message: m,
        lastRound: n,
        gap,
        events: total.get(m) ?? 0,
        cause: dormancyCause(m, source === undefined ? traceSource() : source),
      });
  }
  // Loudest first: a line that fired 155 times before going quiet is worth more
  // of a reader's attention than one that fired twice.
  return rows.sort((a, b) => b.events - a.events);
}

/**
 * Passed, skipped and failed — three outcomes, said as three.
 *
 * A skip means the scenario did not run to a verdict: the host stopped
 * answering, or the deck could not set the question up. NOTHING WAS CHECKED. It
 * is not a pass and it is emphatically not a failure of the code under test, and
 * folding it into "N of M passed" makes a host having a bad afternoon look like
 * a product regression — which is exactly how the 2026-08-24 block first read.
 */
export function selfTestHeadline(selftest) {
  const all = Array.isArray(selftest) ? selftest : [];
  const passed = all.filter((s) => s.ok).length;
  const skipped = all.filter((s) => !s.ok && s.skipped).length;
  const failed = all.filter((s) => !s.ok && !s.skipped).length;
  const parts = [`${passed} of ${all.length} scenarios passed`];
  // Named separately even at zero-of-one-kind, because the interesting case is
  // "12 of 14" and the reader needs to know WHICH kind the missing two were.
  if (skipped) parts.push(`${skipped} skipped (nothing was checked)`);
  if (failed) parts.push(`${failed} FAILED`);
  return parts.join(" · ");
}

/**
 * Recoveries that mean nothing — the ones a round needs BY CONSTRUCTION.
 *
 * A round starts with the pane closed (the browser was just closed, or never
 * opened) or a build behind (the merge under test deployed after the pane
 * loaded). The driver reopens or reloads it and carries on. Measured across
 * every round carrying `driverRun`: `pane-closed` 50%, `pane-stale` 32%.
 *
 * Deliberately NARROW. `host-silent`, `deck-dirty`, `browser-gone` and `crashed`
 * are all states something went wrong to reach, and none of them belong here
 * however often they occur — the point is not "common", it is "expected".
 */
export const ROUTINE_RECOVERIES = new Set(["not-ready:pane-closed", "not-ready:pane-stale"]);

/**
 * What the driver went through to produce this round, in one line.
 *
 * Every field here was already archived and none of it was ever printed. A round
 * that took three attempts and cleared a PowerPoint crash on the way rendered
 * identically to one that walked straight through — and on 2026-08-24 that cost
 * two consecutive wrong reports of the same ten-round block: first "nothing
 * failed, the host just slowed down" when the host had crashed three times, then
 * a correction that attributed the crashes to the wrong rounds and claimed the
 * round files did not record them. `recovered` says `crashed` in plain text on
 * exactly the three rounds concerned.
 *
 * The session fields are the other half — see `sessionPosition`. `laterMed`
 * doubles across a session, so a round's position in one is as load-bearing as
 * its build, and until now only the build was visible.
 *
 * Returns "" rather than a placeholder line for the rounds archived before any
 * of this existed: 190 of them, and a row of question marks on every one is
 * noise, not information.
 */
/**
 * WHICH DOCUMENT, WHICH PROFILE, AND HOW MUCH IT ASKED OF THE HOST.
 *
 * Three fields landed on 2026-09-02 and 09-03 and NOTHING printed any of them,
 * which is the state `driverRun` was in for its first fortnight — archived and
 * invisible, so a round rescued from a crash read exactly like one that never
 * needed rescuing. A field nobody prints is a field nobody uses.
 *
 *   driverDeck   the document actually in front, measured. 0 of 408 archived
 *                files named one before this, which is why "4:3 crashes more"
 *                and "that one file is sick" were the same sentence for a
 *                fortnight. Only names the operator configured are published —
 *                see `archivableDeck` — so `undisclosed` is a real value here
 *                and means "a document this archive will not name", never
 *                "unknown".
 *   syncs        `context.sync()` calls, end to end. office-js#6329 says each
 *                one forces a full presentation save on the web host, so this
 *                is the load figure that would matter if it is right — and
 *                every crash number in this archive is indexed by SHAPES.
 *   asked for    `driverRun.deck`, the deck the leg was TOLD to use. Printed
 *                only when it disagrees with what was in front, because that
 *                disagreement is the one state worth a reader's eye and the
 *                agreement is the ordinary case.
 *
 * Empty string for the rounds that predate all three, on the same reasoning
 * `driverRunLine` gives: a row of question marks on 400 files is noise.
 */
export function roundIdentityLine(log) {
  if (!log || typeof log !== "object") return "";
  const bits = [];
  if (log.driverDeck !== undefined) bits.push(`deck ${log.driverDeck ?? "unreadable"}`);
  const asked = log.driverRun?.deck;
  // Only when they differ. `deck-missing` is the refusal for the loud version
  // of this, but a round can run against the right deck by luck and nothing
  // would say the instruction had been wrong.
  if (asked && log.driverDeck && !String(log.driverDeck).startsWith(String(asked))) bits.push(`ASKED FOR ${asked}`);
  if (log.driverSlideSize) bits.push(String(log.driverSlideSize));
  if (typeof log.syncs === "number") bits.push(`${log.syncs} syncs`);
  return bits.length ? `\n  ${bits.join(" · ")}` : "";
}

export function driverRunLine(dr) {
  if (!dr || typeof dr !== "object") return "";
  const bits = [];
  if (typeof dr.attempts === "number") bits.push(`${dr.attempts} attempt(s)`);
  if (dr.fresh) bits.push("fresh browser");
  if (typeof dr.sessionIndex === "number") {
    // MINUTES INTO THE SESSION, because that is what the drift is indexed by —
    // rounds run 684s to 1085s, so a count is a proxy and the clock is the
    // variable. The gap since the last round is kept beside it because a rested
    // round and a back-to-back one at the same index are different samples.
    const into = typeof dr.sessionElapsedMs === "number" ? `, ${Math.round(dr.sessionElapsedMs / 60000)}m in` : "";
    // Sub-minute gaps read as "0m after the last", which is true and useless:
    // every back-to-back round is ~17s apart. Seconds below a minute.
    const gap =
      typeof dr.sincePrevRoundMs === "number"
        ? dr.sincePrevRoundMs < 60000
          ? `, ${Math.round(dr.sincePrevRoundMs / 1000)}s after the last`
          : `, ${Math.round(dr.sincePrevRoundMs / 60000)}m after the last`
        : "";
    bits.push(`round ${dr.sessionIndex} of this session${into}${gap}`);
  }
  if (dr.startedAt) bits.push(`started ${String(dr.startedAt).replace("T", " ").slice(0, 16)}Z`);
  const rec = Array.isArray(dr.recovered) ? dr.recovered : [];
  // NAMED, not counted. "recovered from 2" invites the reader to skim; "crashed"
  // is the word that stops them.
  //
  // AND SEPARATED BY WHETHER IT MEANS ANYTHING. Pooled over every round carrying
  // the field: `pane-closed` on 50% and `pane-stale` on 32% — 82% between them,
  // because the pane IS closed or a build behind when a round starts. A line
  // that says the same thing on four rounds in five teaches the reader to skim
  // it, and then `crashed` scrolls past inside it. That is not hypothetical:
  // three crashes did exactly that on 2026-08-24, and the correction that
  // followed misattributed them because the same line had been skimmed twice.
  const routine = rec.filter((r) => ROUTINE_RECOVERIES.has(String(r)));
  const notable = rec.filter((r) => !ROUTINE_RECOVERIES.has(String(r)));
  const recovered =
    (notable.length ? `\n  RECOVERED FROM: ${notable.join(", ")}` : "") +
    (routine.length ? `\n  (routine: ${routine.join(", ")})` : "");
  const drift =
    typeof dr.sessionIndex === "number" && dr.sessionIndex >= 5
      ? `\n  DEEP IN A SESSION — the in-place update cost roughly doubles by round 6-8 and scenarios start being skipped; compare against rounds at a similar position, not against round 1`
      : "";
  return (bits.length ? `\n  ${bits.join(" · ")}` : "") + recovered + drift;
}

/**
 * A run of ONE, read against the two arms it sits between.
 *
 * The first chart of a multi-chart run costs ~2.2x a later chart of the same
 * size, with the two distributions not overlapping across 14 and 70 samples. The
 * slide, a cold host, our sync count and the deck scan are all excluded. What is
 * left is that the cost attaches to a run's FIRST chart, and a run of one splits
 * the two readings of that:
 *
 *   pays like a FIRST chart  → the run is the unit; the cost is per-run set-up
 *                              that chart 1 happens to absorb.
 *   pays like a LATER chart  → the cost needs the rest of the queue to exist,
 *                              and "first" was standing in for "has work behind
 *                              it".
 *
 * Says which of the two it is nearer, and REFUSES to call it at n=1 — the same
 * discipline the first-chart line above learned. Nearness is judged on the log
 * scale so "2.2x above" and "2.2x below" are the same distance; on a linear
 * midpoint a 37s reading against arms at 17s and 37s would look ambiguous when
 * it is exactly one of them.
 */
/**
 * The same updates, split by how much this run had already drawn on their slide.
 *
 * NOT "occupancy", and the distinction is this repo's own: `onSlide` counts what
 * THIS RUN drew and reads 0 on a slide's first batch, so a slide that arrived
 * with content reads low. The first version of this line printed "0 shapes" for
 * slides the run had drawn ten shapes onto, which is the same misreading it was
 * written to expose.
 *
 * THIS IS THE ANSWER TO "the first chart of a run costs 2.2x". It does not, or
 * not for being first: in this harness the first chart of the deck-wide rescale
 * is always the chart on the VISIBLE slide, which earlier scenarios have already
 * piled onto — 42 shapes against 20-21 everywhere else. Position and occupancy
 * were perfectly confounded for 14 samples across 10 rounds, and the label on
 * the louder of the two variables stuck.
 *
 * Round 215 separated them by updating a lone chart on an UNCROWDED slide: it
 * cost 15235ms, later-chart cost, against 35033 and 39570ms for the identical
 * arm on the crowded slide in rounds 213 and 214. Visibility is excluded too,
 * because the arm calls `showSlide` on its own target — so 215's cheap chart was
 * the visible one.
 *
 * Prints nothing rather than one bucket: a split with only one side is the
 * confound restated, not resolved.
 */
export function occupancySplit(rows, of, changed) {
  const here = (rows ?? []).filter(
    (r) => r.of === of && r.changed === changed && typeof r.drawnThereBefore === "number",
  );
  if (here.length < 2) return [];
  const counts = [...new Set(here.map((r) => r.drawnThereBefore))].sort((a, b) => a - b);
  if (counts.length < 2) return [];
  const median = (a) => [...a].sort((x, y) => x - y)[Math.floor(a.length / 2)];
  const lo = counts[0];
  const hi = counts[counts.length - 1];
  const loMs = median(here.filter((r) => r.drawnThereBefore === lo).map((r) => r.ms));
  const hiMs = median(here.filter((r) => r.drawnThereBefore === hi).map((r) => r.ms));
  const ratio = loMs ? (hiMs / loMs).toFixed(1) + "x" : "?";
  return [
    `by SHAPES THIS RUN HAD ALREADY DRAWN THERE, at ${changed} of ${of}: ${hi} drawn ${hiMs}ms vs ${lo} drawn ` +
      `${loMs}ms (${ratio}) — which is what "first chart of a run" was standing in for`,
  ];
}

export function aloneVerdict(c) {
  const { alone, first, rest, aloneN } = c;
  const near = Math.abs(Math.log(alone / first)) < Math.abs(Math.log(alone / rest)) ? "a FIRST chart" : "a LATER chart";
  const head = "alone in its own run: " + Math.round(alone) + "ms — nearer " + near;
  if (aloneN < 2) return head + " (n=1, a lead; one round cannot tell this from weather)";
  return head + " (n=" + aloneN + ")";
}

/** What an in-place update costs, and whether the first one in a run is worse. */
function reportUpdateCost(logs) {
  const u = poolUpdateCost(logs);
  if (!u.n) return;
  console.log("\n  WHAT AN IN-PLACE UPDATE COSTS — the round's largest single cost");
  console.log("    " + u.n + " timed update(s). The batch report below does not cover this path at all.");
  for (const f of u.fits) {
    const label = "chart of " + f.of + " nodes";
    const shape = "~" + Math.round(f.fixed) + "ms fixed + ~" + Math.round(f.perShape) + "ms per changed shape";
    const range = "(" + f.lo + "-" + f.hi + " changed, n=" + f.n + ")";
    console.log("    " + label.padEnd(22) + shape + "   " + range);
  }
  for (const c of u.firstVsRest) {
    const ratio = c.rest ? (c.first / c.rest).toFixed(1) : null;
    const at = " at " + c.changed + " of " + c.of + ": ";
    const both = Math.round(c.first) + "ms vs " + Math.round(c.rest) + "ms";
    console.log("    first chart of a run vs the rest" + at + both + (ratio ? " (" + ratio + "x)" : ""));
    // ONE ROUND IS NOT A RATE. Round 198 showed 39355ms against ~18000ms for
    // identical work — a striking lead and a single observation. This line says
    // which it is, so a reader does not have to remember.
    if (c.firstN < 5)
      console.log(
        "      n=" + c.firstN + " first-chart sample(s) — a LEAD, not a rate. Needs rounds before it is anything.",
      );
    if (c.aloneN) console.log("      " + aloneVerdict(c));
    for (const line of occupancySplit(u.rows, c.of, c.changed)) console.log("      " + line);
  }
  // THE ERROR BAR ON EVERY COMPARISON ABOVE — see `poolWithinRoundSpread`.
  const sp = poolWithinRoundSpread(logs);
  if (sp.length) {
    const v = sp.map((x) => x.pct).sort((a, b) => a - b);
    const med = v[Math.floor(v.length / 2)];
    const where = sp.length === 1 ? `this round` : `${sp.length} round(s), median`;
    console.log(`    ROUND SELF-AGREEMENT: ${where} ${med}% spread across its own later charts`);
    console.log(`      Comparisons made INSIDE one round carry at least this much noise. Between rounds the`);
    console.log(`      same measurement varies ~47% rested and 164% across the archive, so a within-round`);
    console.log(`      difference understates real variance by roughly 6-21x. Only large effects survive it.`);
  }
  // WHERE THE TIME WENT — see `poolSyncBreakdown`. Archived on every timed
  // update since `syncMs` landed, and never once printed until now.
  const sb = poolSyncBreakdown(logs);
  if (sb.first.length && sb.later.length) {
    const med = (rows, k) => {
      const v = rows
        .filter((r) => r.length > k)
        .map((r) => r[k])
        .sort((a, b) => a - b);
      return v.length ? v[Math.floor(v.length / 2)] : null;
    };
    console.log(`    PER-SYNC, 18 of 24  (first n=${sb.first.length}, later n=${sb.later.length})`);
    for (let k = 0; k < 4; k++) {
      const a2 = med(sb.first, k);
      const b2 = med(sb.later, k);
      if (a2 == null || b2 == null) continue;
      const label = k < 3 ? `write ${k + 1}` : `tag`;
      console.log(
        `      ${label.padEnd(9)} first ${String(a2).padStart(6)}ms   later ${String(b2).padStart(6)}ms   ${(a2 / b2).toFixed(2)}x`,
      );
    }
  }
  if (u.firstExcessN) {
    const a = Math.round(u.firstExcess);
    const b = Math.round(u.restExcess);
    console.log(
      "    above the fit: first chart +" +
        a +
        "ms (n=" +
        u.firstExcessN +
        "), later charts +" +
        b +
        "ms (n=" +
        u.restExcessN +
        ")",
    );
    if (u.sameSlideN)
      console.log(
        "    on the FIRST CHART'S OWN SLIDE, later charts are +" +
          Math.round(u.sameSlideExcess) +
          "ms (n=" +
          u.sameSlideN +
          ") — this line has claimed POSITION is the cause and then LOAD is the cause." +
          "\n      Neither is established. There is no reliable occupancy measure in a round:" +
          "\n      `onSlide` counts what THIS RUN drew, so the same slide reads 10 and 42 in" +
          "\n      different rounds, and 18 of 50 first charts have no reading at all." +
          "\n      What IS observed: the lone arm cost 15235ms on slide 262 and 35033/39570ms on" +
          "\n      slide 257. The slides differ and the cost differs 2.3x. WHY they differ is open.",
      );
    else if (u.sameSlideUnscorable) {
      // The control EXISTS and cannot be scored against a fit. Show it raw,
      // with the reason, rather than printing a tie it does not support.
      const shown = u.sameSlideRaw.map((r) => r.ms + "ms at " + r.changed + " of " + r.of).join(", ");
      console.log(
        "    " +
          u.sameSlideUnscorable +
          " later chart(s) DO share the first one's slide: " +
          shown +
          " — against the first chart's +" +
          Math.round(u.firstExcess) +
          "ms.",
      );
      console.log("    Unscorable against a fit — their size has only one change-count in the archive — but the");
      console.log("    slide is held constant and the cost is not, which is the comparison that matters.");
    } else console.log("    no later chart shares the first one's slide, so slide and position are still tied here.");
  }
}

/**
 * What a batch of shapes costs, against HOW MUCH THIS RUN HAS ALREADY DRAWN on
 * the same slide.
 *
 * `batch issued` has carried `prevBatchMs` and `onSlide` for the whole archive
 * and nothing had ever pooled either. Together they price the product's biggest
 * latency number: a chart redraw takes about twenty seconds.
 *
 * TWO TRAPS ARE ENCODED HERE, BOTH PAID FOR.
 *
 * 1. `upTo` is a RUNNING TOTAL, not a batch size. Bucketing by it showed
 *    10-shape and 20-shape batches costing the same — therefore payload is
 *    free, therefore raise the batch size. Both buckets were batches of ten;
 *    the second was later in a sequence. Every batch in the archive is ten, so
 *    this data cannot compare batch sizes at all.
 *
 * 2. `onSlideKey === "(visible)"` DOES NOT MEAN THE SLIDE IS ON SCREEN. It is
 *    the sentinel `slideKeyFor` returns when the slide's id has not been
 *    loaded. A first pass at this split batches on it and reported "drawing
 *    where the user is looking costs 2.3x per shape" — which was committed,
 *    and is wrong. Checked afterwards: sentinel-keyed batches have median
 *    `onSlide` of 0 against 10 for id-keyed ones, so they draw onto EMPTIER
 *    targets, and first batches cost 5802ms against 5591ms for later ones —
 *    no setup effect either. The label never carried the meaning the finding
 *    put on it.
 *
 * What survives is the curve below, and it is the one the source already
 * described: `shapesDrawnOn` is exported precisely so a caller can avoid a
 * slide it has been filling.
 *
 * NOTE WHAT `onSlide` COUNTS: shapes THIS RUN drew on that slide, not the
 * slide's total occupancy. A slide arriving with content already on it reads
 * as 0 here. So this measures a run slowing itself down, which is a real and
 * documented effect, and it is NOT a measurement of drawing onto a busy slide.
 */
export function poolDrawCostCurve(logs) {
  const pts = [];
  for (const log of logs ?? []) {
    let prev = null;
    let cum = 0;
    for (const e of log?.trace?.entries ?? []) {
      if (String(e.message ?? "") !== "batch issued") continue;
      const d = e.data ?? {};
      if (typeof d.prevBatchMs === "number" && prev && prev.size > 0)
        pts.push({ ms: d.prevBatchMs, size: prev.size, drawnHere: prev.drawnHere });
      if (typeof d.upTo !== "number") {
        prev = null;
        cum = 0;
        continue;
      }
      // A new draw restarts the count, so a smaller upTo begins a fresh one.
      const size = d.upTo > cum ? d.upTo - cum : d.upTo;
      prev = { size, drawnHere: typeof d.onSlide === "number" ? d.onSlide : null };
      cum = d.upTo;
    }
  }
  const median = (a) => {
    if (!a.length) return null;
    const b = [...a].sort((x, y) => x - y);
    return b[Math.floor(b.length / 2)];
  };
  const bands = [
    [0, 0],
    [1, 20],
    [21, 50],
    [51, 100],
    [101, Infinity],
  ];
  const rows = [];
  for (const [lo, hi] of bands) {
    const a = pts.filter((p) => p.drawnHere != null && p.drawnHere >= lo && p.drawnHere <= hi).map((p) => p.ms);
    if (a.length < 10) continue;
    rows.push({ lo, hi, n: a.length, median: median(a) });
  }
  return {
    n: pts.length,
    sizes: [...new Set(pts.map((p) => p.size))].sort((a, b) => a - b),
    rows,
  };
}

/**
 * The empty round-trip, for scale.
 *
 * Without it a batch median is a number with nothing to be big compared to, and
 * this project spent an hour arguing about whether round-trips dominate drawing
 * while the answer was printed above every round.
 */
export const EMPTY_ROUNDTRIP_MS = 5; // driver ping: slides.getCount() + one sync

/**
 * The draw-cost curve. Not to be confused with `reportBatchCost` further down,
 * which takes ONE round and looks for two populations in its batch times.
 */
function reportDrawCostCurve(logs) {
  const c = poolDrawCostCurve(logs);
  if (!c.rows.length) return;
  // NOT THE PRODUCT'S BIGGEST LATENCY NUMBER, which is what this headline said
  // when it was committed earlier today. `same scale across the deck` is the
  // most expensive scenario in the round — 166s median, 38% of it — and it
  // issues NOT ONE `batch issued` line: eight charts, all of them through the
  // in-place update path. So everything below covers the DRAW path and misses
  // the largest cost entirely. Claiming otherwise was a scope error of exactly
  // the kind this section already carries two warnings about.
  console.log("\n  WHAT A BATCH OF SHAPES COSTS — on the DRAW path only");
  console.log("    (the in-place update path issues no batches; see `updated only the shapes that changed`)");
  console.log("    shapes this run already drew here    median batch    n");
  for (const r of c.rows) {
    const band = r.hi === Infinity ? r.lo + "+" : r.lo + "-" + r.hi;
    console.log("      " + band.padEnd(34) + String(Math.round(r.median) + "ms").padStart(8) + "   " + r.n);
  }
  const first = c.rows[0];
  const last = c.rows[c.rows.length - 1];
  if (first && last && first.median > 0 && last !== first)
    console.log(
      "    A run slows ITSELF down " +
        (last.median / first.median).toFixed(1) +
        "x by piling onto one slide. `shapesDrawnOn` is exported to avoid exactly this.",
    );
  console.log(
    "    An empty round-trip is " + EMPTY_ROUNDTRIP_MS + "ms, so batching is not a lever here — drawing is the cost.",
  );
  // AND SAY WHAT THIS CANNOT ANSWER, because the last two attempts at this
  // section both answered a question the data did not cover.
  if (c.sizes.length <= 1)
    console.log("    Every batch here is " + c.sizes[0] + " shapes, so this says NOTHING about batch size.");
  console.log("    `onSlide` counts what THIS RUN drew, not the slide's occupancy — a slide that");
  console.log("    arrived with content reads as 0, so this is not a measure of drawing onto a busy slide.");
}
/**
 * The build that first carried the SECOND SETTLE ASK (#709, `REREAD_ATTEMPTS` 1 -> 2).
 *
 * A round records its build, not the pull requests inside it, so this is the
 * only durable handle on "the era in which the change was live". Recheck it
 * with `git log --oneline -S "REREAD_ATTEMPTS = 2" -- src/render/powerpoint.ts`
 * if the constant ever stops matching a round in the archive.
 */
export const SECOND_ASK_BUILD = "2912802";

export function poolSettleAsks(logs) {
  const out = {
    rounds: 0,
    first: 0,
    second: 0,
    later: 0,
    unlabelled: 0,
    survivors: 0,
    charts: 0,
    // WHEN, not just how many — and TWO SEPARATE CLOCKS, because they are two
    // separate commits and reading one as the other is how this whole section
    // got written in the first place.
    sinceLastSurvivor: null, //  rounds since a failure last survived every ask
    changeLandedAgo: null, //    rounds since the SECOND ASK shipped (#709)
    labelledAgo: null, //        rounds since the trace could TELL ASKS APART (#710)
    survivorsSinceChange: 0,
    roundsSinceChange: 0,
    // THE BASELINE FOR `sinceLastSurvivor`. Without it a quiet stretch has nothing
    // to be long compared to, and this journal has already once called 12 quiet
    // rounds a ~1% event by picking the split point after seeing the zeros.
    longestPriorGap: null,
    survivorRounds: 0,
    // DID THE SECOND ASK WORK, not merely fire. Counting that it happened says
    // nothing about whether it helped, and the difference between "it fired"
    // and "it worked" is the distinction this whole section exists to draw.
    secondAskRescued: 0,
    secondAskLost: 0,
  };
  const all = logs ?? [];
  // The index of the LAST round that produced a survivor, of the first round on
  // the build that shipped the SECOND ASK, and of the first round whose asks
  // carry an `attempt` number.
  //
  // THE LAST TWO ARE NOT THE SAME ROUND, and the first draft of this reader
  // assumed they were:
  //
  //     #709  2912802  raised REREAD_ATTEMPTS 1 -> 2      first ran in round 187
  //     #710  2fd8401  added the `attempt` label to the trace  first ran in round 189
  //
  // Two rounds apart. Deriving "when the change landed" from the label reported
  // the change as two rounds old when it was four, under a field named for the
  // change — an instrument answering a question it was not measuring, which is
  // the exact defect this file exists to catch. Verify with:
  //
  //     git log --oneline -S "REREAD_ATTEMPTS = 2" -- src/render/powerpoint.ts
  //
  // The sha is a constant because nothing in a round's TRACE records which
  // source change it is running; the round records its BUILD, so the build is
  // what we match. The archive only moves forward, so the first round on that
  // build begins the era.
  let lastSurvivorAt = -1;
  const survivorAt = [];
  let labelledAt = -1;
  let changeLandedAt = -1;
  for (let i = 0; i < all.length; i++) {
    const log = all[i];
    let any = false;
    for (const e of log?.trace?.entries ?? []) {
      const m = String(e.message ?? "");
      if (/settle delay/.test(m)) {
        any = true;
        out.charts += Number(e.data?.charts) || 0;
        const a = e.data?.attempt;
        if (typeof a !== "number") out.unlabelled++;
        else {
          if (labelledAt < 0) labelledAt = i;
          if (a <= 1) out.first++;
          else if (a === 2) out.second++;
          else out.later++;
        }
      } else if (/re-read named none/.test(m)) {
        any = true;
        out.survivors++;
        if (lastSurvivorAt !== i) survivorAt.push(i);
        lastSurvivorAt = i;
      }
    }
    if (any) out.rounds++;
    if (changeLandedAt < 0 && String(log?.build ?? "").startsWith(SECOND_ASK_BUILD)) changeLandedAt = i;
  }
  const n = all.length;
  if (lastSurvivorAt >= 0) out.sinceLastSurvivor = n - 1 - lastSurvivorAt;
  // Every gap BETWEEN survivor rounds, so the current quiet stretch can be read
  // against the quiet stretches this archive has already produced.
  out.survivorRounds = survivorAt.length;
  // What FOLLOWED each second ask. The next decisive line wins: a group means it
  // rescued the chart, a survivor line means it did not. Anything else is
  // neither and is left uncounted rather than assumed either way.
  for (const log of all) {
    const es = log?.trace?.entries ?? [];
    for (let i = 0; i < es.length; i++) {
      const e = es[i];
      if (!/settle delay/.test(String(e?.message ?? ""))) continue;
      if (!(Number(e?.data?.attempt) >= 2)) continue;
      for (let j = i + 1; j < es.length; j++) {
        const m = String(es[j]?.message ?? "");
        if (/re-read named none/.test(m)) {
          out.secondAskLost++;
          break;
        }
        if (/grouped the chart's shapes/.test(m)) {
          out.secondAskRescued++;
          break;
        }
        // A further ask means this one is still in flight; stop looking.
        if (/settle delay/.test(m)) break;
      }
    }
  }
  for (let i = 1; i < survivorAt.length; i++) {
    const gap = survivorAt[i] - survivorAt[i - 1];
    if (out.longestPriorGap == null || gap > out.longestPriorGap) out.longestPriorGap = gap;
  }
  if (labelledAt >= 0) out.labelledAgo = n - 1 - labelledAt;
  if (changeLandedAt >= 0) {
    out.changeLandedAgo = n - 1 - changeLandedAt;
    out.roundsSinceChange = n - changeLandedAt;
    // Survivors are counted from the change forward, so "did it help" is asked
    // of the rounds that could possibly answer it.
    for (let i = changeLandedAt; i < n; i++)
      for (const e of all[i]?.trace?.entries ?? [])
        if (/re-read named none/.test(String(e.message ?? ""))) out.survivorsSinceChange++;
  }
  return out;
}

/** The settle asks, and whether the second one ever fires. */
function reportSettleAsks(logs) {
  const a = poolSettleAsks(logs);
  if (!a.rounds) return;
  console.log("\n  THE SETTLE RETRY — over " + a.rounds + " round(s)");
  console.log("    " + String(a.first).padStart(5) + "  first ask");
  console.log("    " + String(a.second).padStart(5) + "  SECOND ask (added in #709 to reach the survivors)");
  if (a.later) console.log("    " + String(a.later).padStart(5) + "  later asks");
  if (a.unlabelled)
    console.log("    " + String(a.unlabelled).padStart(5) + "  unlabelled — rounds before the attempt field");
  console.log("    " + String(a.survivors).padStart(5) + "  failures that survived every ask");
  // THE OUTCOME OF THE SECOND ASK, printed whenever one has fired. On
  // 2026-08-23 round 194 produced the first: the cold re-read fell short, the
  // first ask fell short again, the second ask fired, and the chart grouped
  // "by ids" with partial 0. Under the old REREAD_ATTEMPTS = 1 that chart would
  // have been a survivor. One event is not a rate — but it is the difference
  // between a change that works and a change nothing has ever exercised.
  if (a.second)
    console.log(
      "    of the " +
        a.second +
        " second ask(s): " +
        a.secondAskRescued +
        " rescued the chart, " +
        a.secondAskLost +
        " still lost it.",
    );

  // THE CLAIM, which needs every ask labelled to be worth making. An
  // unlabelled ask could have been a second one, so a slice containing any
  // cannot say the second never fired.
  if (!a.second && !a.unlabelled) console.log("    The second ask has NEVER FIRED — it has had nothing to rescue.");

  // THE TIMING, which needs no such thing — and MUST NOT SHARE THE GUARD.
  //
  // It did share it for one draft, and so it never printed: the archive holds
  // 2,144 unlabelled asks from before #710, so `!a.unlabelled` is false and
  // stays false for as long as those rounds are pooled. A diagnostic written
  // to catch a misreading, placed behind a condition that is never true, is
  // just a comment. The backlog of unlabelled rounds is the case that needs
  // the timing MOST, because the ask counts can say nothing at all there.
  if (a.sinceLastSurvivor == null) {
    if (a.survivors) return;
    console.log("    No round in this slice has ever produced a survivor.");
    return;
  }
  if (a.changeLandedAgo == null) {
    // No round here ran the build that shipped the second ask, so there is no
    // era to compare against and none is invented.
    console.log(
      "    Last survivor " +
        a.sinceLastSurvivor +
        " round(s) ago. No round in this slice ran build " +
        SECOND_ASK_BUILD +
        ", so it cannot say when the second ask arrived.",
    );
    return;
  }
  if (a.sinceLastSurvivor > a.changeLandedAgo) {
    const before = a.sinceLastSurvivor - a.changeLandedAgo;
    console.log(
      "    THE POPULATION DRIED UP BEFORE THE CHANGE: last survivor " +
        a.sinceLastSurvivor +
        " round(s) ago, the second ask shipped " +
        a.changeLandedAgo +
        " round(s) ago — " +
        before +
        " clean round(s) came FIRST.",
    );
    console.log("    So the quiet rounds since are NOT evidence for it. Something else moved this number.");
    // AND WHETHER THE QUIET IS EVEN UNUSUAL. "13 rounds since the last one" reads
    // as remarkable and may be ordinary; this archive's own gaps are the only
    // thing that can say. Reporting the stretch without the distribution is how
    // a long tail on a sporadic event gets written up as an inflection.
    if (a.longestPriorGap != null) {
      const unusual = a.sinceLastSurvivor > a.longestPriorGap;
      console.log(
        "    Quiet for " +
          a.sinceLastSurvivor +
          " round(s); the longest gap between survivors before this was " +
          a.longestPriorGap +
          " (over " +
          a.survivorRounds +
          " round(s) that produced one) — " +
          (unusual
            ? "so this is the longest on record, but NOT out of family."
            : "so this is ORDINARY. Not an inflection."),
      );
    }
    // AND NAME WHAT WE CANNOT RULE OUT, because "something else" invites the
    // reader to fill the blank with the nearest recent commit.
    //
    // Every round in this archive reports the same host string —
    // "PowerPoint · OfficeOnline · 0.0.0.0". Office Online does not version
    // itself to an add-in, so a host update is INVISIBLE HERE: it cannot be
    // confirmed and it cannot be excluded. A constant that is a constant
    // because nothing measures it reads exactly like a controlled variable,
    // and it is not one.
    const hosts = new Set((logs ?? []).map((l) => String(l?.host ?? "")).filter(Boolean));
    if (hosts.size === 1)
      console.log(
        "    The host is unversioned (" +
          [...hosts][0] +
          ") in all " +
          a.rounds +
          " round(s), so a host-side change cannot be ruled in OR out from this archive.",
      );
    return;
  }
  console.log(
    "    Last survivor " +
      a.sinceLastSurvivor +
      " round(s) ago; the second ask shipped " +
      a.changeLandedAgo +
      " round(s) ago. " +
      a.survivorsSinceChange +
      " survivor(s) in the " +
      a.roundsSinceChange +
      " round(s) since.",
  );
}

function reportStarvedQuestions(logs) {
  const dead = poolStarvedQuestions(logs);
  if (!dead.length) return;
  console.log(`\n  QUESTIONS THAT NEVER ANSWERED — pooled over ${logs.length} round(s)`);
  const show = (title, rows, note) => {
    if (!rows.length) return;
    console.log(`    ${title}`);
    for (const r of rows) console.log(`      ${r.id.padEnd(44)} ${String(r.rounds).padStart(3)} round(s)  ${r.last}`);
    console.log(`      ${note}`);
  };
  const live = dead.filter((d) => !d.retired);
  show(
    "never asked — the harness could not set the question up (OURS to fix)",
    live.filter((d) => d.never >= d.unanswerable),
    "Move it into production instrumentation or retire it; it costs a scratch slide either way.",
  );
  // THE PRODUCTION WITNESS for the one question above that will never answer
  // itself. Printed under the row that raises it, because a reader told to
  // "move it into production instrumentation" deserves to see whether that has
  // already happened and what it found.
  if (live.some((d) => d.id === "shape-resolve-held-slide-proxy")) {
    const w = poolAgedHandleResolves(logs);
    console.log(
      `      ^ production witness (deleteShapesById): ${w.resolved} resolved, ${w.refused} refused, ` +
        `across ${w.rounds} of ${w.of} round(s)`,
    );
    if (!w.rounds)
      console.log(`        The sweep only runs when a round leaves wreckage, and none has. Still unanswered.`);
  }
  show(
    "asked, and the host would not answer (a fact ABOUT the host)",
    live.filter((d) => d.unanswerable > d.never),
    "Leave it. An unanswerable question is a finding, not a failure.",
  );
  // ALREADY DONE, and said so. These rows sat in the OURS-TO-FIX bucket for
  // eight rounds after the work landed.
  show(
    "already retired — the archive remembers them, the build no longer asks",
    dead.filter((d) => d.retired),
    "Nothing to do. Here so a row that vanished reads as finished rather than as lost.",
  );
}

/** Fresh slide versus established — where a config is really decided. */
function reportFreshVsEstablished(logs) {
  const f = poolFreshVsEstablished(logs);
  if (!f.established && !f.fresh) return;
  console.log(`\n  WHICH SLIDE THE CHART LANDED ON — pooled over ${logs.length} round(s)`);
  console.log(
    `    slide already had shapes  ${String(f.established).padStart(3)} chart(s), ${String(f.establishedGrouped).padStart(3)} grouped = ${rateOrSilence(f.establishedGrouped, f.established)}`,
  );
  console.log(
    `    freshly added, empty      ${String(f.fresh).padStart(3)} chart(s), ${String(f.freshGrouped).padStart(3)} grouped = ${rateOrSilence(f.freshGrouped, f.fresh)}`,
  );
  // THE CURRENT RATE, because the pooled one is known to be unreadable and this
  // report said so in prose and printed it anyway: "pooled over 39 rounds that
  // predate the retry, so it will climb slowly and should not be read as the
  // current rate". A number a reader is told to mentally discount is a number
  // nobody can use — and the answer is not a caveat, it is a second window.
  //
  // 20 rounds: long enough to be a rate rather than an afternoon, short enough
  // that nothing before the settled retry (rounds 064/065) is in it.
  const r = recentFreshVsEstablished(logs);
  if (r && (r.established || r.fresh)) {
    console.log(`    last ${RECENT_ROUNDS} round(s) — the rate that is actually current:`);
    console.log(
      `      slide already had shapes  ${String(r.established).padStart(3)} chart(s), ${String(r.establishedGrouped).padStart(3)} grouped = ${rateOrSilence(r.establishedGrouped, r.established)}`,
    );
    console.log(
      `      freshly added, empty      ${String(r.fresh).padStart(3)} chart(s), ${String(r.freshGrouped).padStart(3)} grouped = ${rateOrSilence(r.freshGrouped, r.fresh)}`,
    );
    // A WINDOW IS A GUESS, A SEQUENCE IS EVIDENCE — and this population is
    // emptying, which no percentage can show. Over successive windows it reads
    // 80% at 30 rounds, 91% at 20, 100% at 12, and at 8 there are no fresh-slide
    // charts left to measure: the in-place update took them, and a chart that is
    // not redrawn never lands on a fresh slide. `0/0` rounds are the signal.
    const seq = recentFreshSequence(logs);
    if (seq.length) console.log(`      grouped/landed on a fresh slide, per round: [${seq.join(" ")}]`);
  }
  if (f.established && f.fresh)
    console.log(
      `    A chart on a freshly added slide USED NOT TO GROUP: its pre-grouping re-read came\n` +
        `    back short or empty, so it fell through ungrouped and lost its config. Since the\n` +
        `    settled retry those charts DO group — rounds 064 and 065, both, on the same two\n` +
        `    charts. The all-time percentage above is pooled over 39 rounds that predate the\n` +
        `    retry and cannot be read as the current rate. The recent window was what to quote\n` +
        `    UNTIL its population emptied — the in-place update took the redraws, and a chart\n` +
        `    that is not redrawn never lands on a fresh slide, so the window now rates a\n` +
        `    handful of samples from one unrepresentative path. Quote the per-round sequence.\n` +
        `    What such a chart still loses is the TAG, refused through the GROUP handle: the\n` +
        `    group hangs off a slide handle Office has rewritten to slides.getItem(id), and a\n` +
        `    freshly added slide's id does not round-trip on this host.\n` +
        `    See the #108-#111 saga and shape-add-held-slide-proxy.`,
    );
}

/** Grouped versus ungrouped, and what it costs a chart's config. */
function reportGroupVsTag(logs) {
  const g = poolGroupVsTag(logs);
  if (!g.grouped && !g.ungrouped) return;
  const cov = poolGroupVsTagCoverage(logs);
  console.log(`\n  DOES GROUPING SAVE THE CONFIG — over the charts of the DECK-WIDE RESCALE, ${logs.length} round(s)`);
  // THE DENOMINATOR THE READER WOULD OTHERWISE ASSUME. These two columns are
  // joined by `data.chart`, which only the rescale scenario writes, so this is
  // one scenario's charts and not the round's. Printed because the blindness is
  // UNEVEN — it drops far more of the grouped column than the ungrouped one —
  // and an uneven filter biases the ratio rather than merely shrinking it.
  console.log(
    `    scope: ${cov.groupedSeen}/${cov.groupedTotal} grouped and ${cov.ungroupedSeen}/${cov.ungroupedTotal} ungrouped` +
      ` events carry a chart key; the rest of the round is not in these numbers.`,
  );
  console.log(
    `    grouped      ${String(g.grouped).padStart(3)} chart(s), ${String(g.groupedLost).padStart(3)} lost the tag = ${rateOrSilence(g.groupedLost, g.grouped)}`,
  );
  console.log(
    `    NOT grouped  ${String(g.ungrouped).padStart(3)} chart(s), ${String(g.ungroupedLost).padStart(3)} lost the tag = ${rateOrSilence(g.ungroupedLost, g.ungrouped)}`,
  );
  // Said out loud, because the number's whole value is what it implies about
  // where to work, and that was missed for eleven rounds while it sat here.
  if (g.grouped && g.ungrouped && g.ungroupedLost / g.ungrouped > 2 * (g.groupedLost / g.grouped || 0.001))
    console.log(
      `    A chart that groups usually keeps its config; one that cannot usually loses it.\n` +
        `    READ THE SPLIT ABOVE BEFORE QUOTING THIS. Until round 064 a freshly-added slide's\n` +
        `    chart COULD NOT GROUP, so it could never appear in the grouped column — this ratio\n` +
        `    was measured on a population that excluded the hard case by construction. Round 064\n` +
        `    made those charts group and they lost their tag anyway (from: group, 5010, on the\n` +
        `    slide the run had just added). The rule is closer to "a slide that was already\n` +
        `    there saves the config", with grouping standing in for it.`,
    );
}

/** The tag-fault table, and the noise floor it implies. */
function reportTagFaults(logs) {
  const byBuild = poolTagFaults(logs);
  if (byBuild.size < 2) return;
  const kinds = ["tags-undefined", "cfg-tag-5010", "group-5010", "no-queue", "tagging-failed"];
  console.log(`\n  TAG FAULTS PER BUILD — pooled over ${logs.length} round(s)`);
  console.log(`    ${"build".padEnd(9)}${"n".padEnd(4)}${kinds.map((k) => k.padStart(16)).join("")}`);
  const spread = {};
  for (const [build, rounds] of byBuild) {
    const cell = (k) => {
      const vs = rounds.map((r) => r[k]);
      const lo = Math.min(...vs);
      const hi = Math.max(...vs);
      (spread[k] ??= []).push(lo, hi);
      return (lo === hi ? String(lo) : `${lo}-${hi}`).padStart(16);
    };
    console.log(`    ${build.padEnd(9)}${String(rounds.length).padEnd(4)}${kinds.map(cell).join("")}`);
  }
  // THE NOISE FLOOR, measured WITHIN one build wherever a build has been run
  // twice — which is the only measurement that owes nothing to an argument. A
  // spread across builds can always be answered with "but the code changed";
  // the same build twice cannot.
  const within = [];
  for (const [build, rounds] of byBuild) {
    if (rounds.length < 2) continue;
    for (const k of kinds) {
      const vs = rounds.map((r) => r[k]);
      const lo = Math.min(...vs);
      const hi = Math.max(...vs);
      if (hi > lo) within.push({ build, k, lo, hi });
    }
  }
  const worst = within.sort((a, b) => b.hi - b.lo - (a.hi - a.lo))[0];
  if (worst)
    console.log(
      `    NOISE FLOOR, from one build run ${byBuild.get(worst.build).length}× : ${worst.build} scored ` +
        `${worst.lo} and ${worst.hi} for ${worst.k}\n` +
        `    with NOTHING changed between them. A difference smaller than that is not evidence of a change.\n` +
        `    Judge a change on a count that did NOT move, or on a trace line that appears where none did.\n` +
        // CORRECTED 2026-08-20. This line said the spread WAS "the host's mood",
        // which is a claim about its CAUSE and it is wrong. The spread has a
        // direction: pooled over all 31 pairs this archive holds, the SECOND
        // round of a build is worse than the first in 15 of the 17 pairs that
        // moved at all. Mood is symmetric; this is not. Calling a directional
        // effect noise is how it stayed invisible through nineteen pairs.
        `    NOT symmetric: see PAIR POSITION below — the second run is usually the worse one,\n` +
        `    so a floor measured this way includes an effect as well as noise.`,
    );
  else
    console.log(
      `    No build has been run twice, so there is no measured noise floor and no difference\n` +
        `    in this table can be attributed to a change. Run one build twice before believing it.`,
    );
}

/** The pooled arms, printed — or nothing when no round carried the scenario. */
function reportPool(logs) {
  const { rounds, arms } = poolRasteriseArms(logs);
  const total = Object.values(arms).reduce((n, a) => n + a.ok + a.stall, 0);
  if (!total) return;
  console.log(`\n  DOES A RASTERISE POISON THE NEXT DRAW — pooled over ${rounds} round(s)`);
  for (const [name, a] of Object.entries(arms)) {
    const n = a.ok + a.stall;
    const rate = n ? `${((100 * a.stall) / n).toFixed(1)}%` : "—";
    console.log(
      `    after a ${name.padEnd(11)} ${String(a.stall).padStart(3)} stalled / ${String(n).padStart(3)} drawn = ${rate}`,
    );
  }
  const n = Math.min(...Object.values(arms).map((a) => a.ok + a.stall));
  if (n < 60)
    console.log(
      `    NOT an answer yet: ${n} draws in the smaller arm. Telling rates this close apart\n` +
        `    needs nearer 60-100 an arm, which is dozens more rounds at two an arm each.`,
    );

  // The same question asked of every draw the round made, which is ~10x the
  // sample and weaker evidence. See `poolEveryDraw`.
  const every = poolEveryDraw(logs);
  const seen = Object.values(every.after).reduce((t, a) => t + a.ok + a.stall, 0);
  if (!seen) return;
  console.log(`\n  EVERY DRAW IN THE ROUND — observational, not counterbalanced`);
  for (const [name, a] of Object.entries(every.after)) {
    const m = a.ok + a.stall;
    const rate = m ? `${((100 * a.stall) / m).toFixed(1)}%` : "—";
    const label = name === "rasterise" ? "after a rasterise" : "after anything else";
    console.log(
      `    ${label.padEnd(20)} ${String(a.stall).padStart(3)} stalled / ${String(m).padStart(3)} drawn = ${rate}`,
    );
  }
  console.log(
    `    Draws follow a rasterise because of WHICH SCENARIO they belong to, and those\n` +
      `    differ in shape count and in what they ask of the host. Suspicion, not verdict —\n` +
      `    the arms above are the controlled measurement, this is the one with the numbers.`,
  );
}

/** The batch-time split and the per-slide slope, printed when a round has them. */
export function reportBatchCost(log) {
  const pop = batchPopulations(log);
  if (pop) {
    console.log(`\n  DRAW BATCHES ${pop.times.length} timed, slowest ${(pop.max / 1000).toFixed(1)}s`);
    if (pop.bimodal)
      console.log(
        `    two populations: ${pop.fast.length} fast (${(pop.fast[0] / 1000).toFixed(1)}-${(pop.fast[pop.fast.length - 1] / 1000).toFixed(1)}s) ` +
          `and ${pop.slow.length} slow (${(pop.slow[0] / 1000).toFixed(1)}-${(pop.slow[pop.slow.length - 1] / 1000).toFixed(1)}s), ` +
          `nothing in the ${(pop.gap / 1000).toFixed(1)}s between`,
      );
    else console.log(`    one population — no empty band worth the name (largest gap ${(pop.gap / 1000).toFixed(1)}s)`);
  }
}

export function reportSelfTest(selftest) {
  if (!selftest.length) return 0;
  const failed = selftest.filter((s) => !s.ok && !s.skipped).length;
  // SKIPPED IS NOT FAILED, AND NEITHER IS PASSED. "12 of 14 scenarios passed"
  // reads as two failures; on 2026-08-24 every one of those was a SKIP — the
  // host stopped answering mid-scenario and NOTHING WAS CHECKED. That is a
  // third outcome and the headline had two, so a degrading host looked like a
  // regressing product. The per-line marks have always said `skip`; the line
  // people quote did not.
  console.log(`\n  SELF-TEST ${selfTestHeadline(selftest)}\n`);
  for (const s of selftest) {
    const mark = s.skipped ? "skip" : s.ok ? "ok" : "FAIL";
    console.log(`  ${pad(mark, 6)}${pad(s.name, 46)}${s.detail}`);
    // A scenario's own words are a problem string like any other, and the
    // one place a host bug is stated in plain language rather than in a
    // host error code. Annotating only the `problems` tally missed it.
    const known = knownBug(s.detail ?? "");
    if (known) console.log(`  ${pad("", 6)}${pad("", 46)}  ^ known host bug: ${known}`);
  }
  console.log("");
  return failed;
}

/**
 * What the round left in the deck, from the round's OWN evidence.
 *
 * The deck-evidence change made the file self-contained on purpose — "there is
 * no deck to save and no screenshot to take" — and then nothing read the part
 * that replaced them. Working the 2026-08-09 round out by hand took a dozen
 * ad-hoc queries to reach three numbers that are right here: how many added
 * slides carry shapes, how many read back empty, and how many of the empties
 * the host's own rasteriser AGREES are empty.
 *
 * That last column is the one worth having. An empty readback and a blank
 * picture are two witnesses; an empty readback alone is one, and this host is
 * known to answer a shape collection short without throwing
 * (`shapes-items-count-honest`). A slide with no shot is not evidence of
 * anything — the rasterise pass is capped — so it is counted apart rather than
 * folded in with the confirmed blanks.
 */
export function deckEvidence(deck) {
  if (!deck || !Array.isArray(deck.inventory) || !deck.inventory.length) return null;
  // A STRIPPED PICTURE IS NOT A BLANK ONE, and for 323 rounds this line could
  // not tell the difference. `stripImages` replaces every base64 payload with
  // the sentence "<image stripped for the archive — …" before a round is
  // written, so ALL 1,923 shots in the archive are a 53- or 55-character string.
  // That string is truthy, so it survived this filter as though a picture had
  // been taken, and `bytes()` scores it at 40-41 — comfortably under
  // `BLANK_PNG_CEILING`. Every added slide in every archived round therefore
  // came back with the rasteriser AGREEING it was blank: the second witness, the
  // one the docstring above calls "proven data loss". Ten empty added slides
  // across nine rounds read `confirmed`, `unseen: 0`, and the honest `unseen`
  // branch below had never once executed.
  //
  // Fixed HERE rather than in the stripper because the archive is append-only:
  // those 1,923 placeholders cannot be rewritten, and re-triaging the archive is
  // what the archive is for. Matched on the stable prefix — the sentence itself
  // has already changed once, from `rounds/README.md` to `docs/ROUNDS.md`.
  const stripped = (png) => typeof png === "string" && png.startsWith("<image stripped");
  const shots = new Map((deck.shots ?? []).filter((s) => s.png && !stripped(s.png)).map((s) => [s.slideId, s.png]));
  /**
   * How big a slide's PNG can be and still be nothing.
   *
   * This used to be `Math.min` over the very shots being classified, and that
   * cannot work: the smallest picture satisfies `<= itself` unconditionally, so
   * the tool could never report a round with NO blank in it. At least one added
   * slide was always printed under "read back empty AND rasterise blank" — the
   * two-witness line a maintainer reads as proven data loss — however much
   * chart was on it. It got the right answer on all ten real rounds only
   * because each contains genuine blanks; it breaks in exactly the regime the
   * column exists for, the round after the drawing bug is fixed, where every
   * slide has content and one answers its shape collection short (documented
   * host behaviour). The honest verdict there is "the readback is lying", and
   * the tool said "confirmed blank".
   *
   * A relative test cannot replace it either. Rounds 12 and 13 ran with the
   * picture cap at 12 and EVERY shot they took is a blank, so there is no
   * contrast in the round to measure against — judging by bimodality calls
   * those 23 genuinely blank slides content.
   *
   * So the assumption is made explicit instead of hidden in a `min`. Every
   * round this tool reads is a 960×540 deck from this add-in, and the measured
   * spread is not close: a blank is 1146 bytes across all ten rounds, and the
   * smallest picture with a chart on it is 5142. Two kilobytes sits 4.5× above
   * the one and 2.5× below the other.
   */
  const BLANK_PNG_CEILING = 2048;
  const bytes = (png) => Math.round((png.length * 3) / 4);
  const added = new Set(deck.newSlides ?? []);
  /**
   * An id whose second half is `0`, e.g. `256#0` against `288#3603562595`.
   *
   * Round 041 added seven slides and two came back shaped like that, in the same
   * round that reported `delete-by-id left slides behind and this host does not
   * list the ids they were added under` and left one added slide blank. Those
   * may be one fact or three; what is certain is that nobody saw it, because
   * both halves were in the log and nothing put them side by side. It took a
   * hand query to notice, which is the definition of a thing this tool should
   * have said.
   *
   * Reported as a SHAPE, not a diagnosis. What `#0` means to this host is not
   * known — only that it is not what a slide it has finished adding looks like.
   */
  const unfinishedId = (id) => /#0$/.test(String(id ?? ""));
  const out = {
    scanned: deck.inventory.length,
    added: added.size,
    withShapes: 0,
    confirmed: 0,
    unseen: 0,
    lying: 0,
    oddIds: [...added].filter(unfinishedId),
    oddAndBlank: 0,
  };
  for (const s of deck.inventory) {
    if (!added.has(s.slideId)) continue;
    // `count` is the host's own number and `shapes.length` is what the scan
    // managed to list; the SMALLER would call a partial listing empty, so the
    // larger is what decides whether anything is there.
    const n = Math.max(s.count ?? 0, s.shapes?.length ?? 0);
    if (n > 0) {
      out.withShapes++;
      continue;
    }
    const png = shots.get(s.slideId);
    if (!png) out.unseen++;
    else if (bytes(png) <= BLANK_PNG_CEILING) {
      out.confirmed++;
      // The join that had to be done by hand. If the blank slides ARE the
      // oddly-named ones, that is one finding rather than two — and if they are
      // not, the coincidence is dead and nobody spends a round on it.
      if (unfinishedId(s.slideId)) out.oddAndBlank++;
    }
    // A picture that is not blank and not obviously a chart cannot be called
    // either way — and calling it blank is the expensive mistake, because that
    // is the reading that alleges data loss.
    else out.lying++;
  }
  return out;
}

function reportDeckEvidence(deck) {
  const e = deckEvidence(deck);
  if (!e) return;
  console.log(
    `\n  DECK EVIDENCE ${e.scanned} slide(s) scanned, ${e.added} added by this round:` +
      `\n    ${e.withShapes} carry shapes` +
      `\n    ${e.confirmed} read back empty AND rasterise blank` +
      `\n    ${e.unseen} read back empty, no picture taken (the rasterise pass is capped — not evidence)` +
      (e.lying ? `\n    ${e.lying} read back empty but rasterise with CONTENT — the readback is lying` : "") +
      (e.oddIds.length
        ? `\n    ${e.oddIds.length} added with an id ending #0 (${e.oddIds.slice(0, 4).join(", ")}` +
          `${e.oddIds.length > 4 ? ", …" : ""}) — not the shape this host gives a slide it has finished adding` +
          `\n      of which ${e.oddAndBlank} also read back empty AND rasterised blank`
        : ""),
  );
  if (deck.gap) console.log(`    scan gap: ${deck.gap}`);
}

/**
 * Did an in-place update leave the rest of an ungrouped chart on the slide?
 *
 * The pooled reading for `shapes left on the slide after an in-place update`.
 * The question it settles has been open for 56 rounds: `reading back an
 * ungrouped chart's shape ids` is the last `GetItem(id)` refusal still firing,
 * each failure costs a chart its parts list, and the comment beside that read
 * says such a chart "grows by a whole chart on every edit" — which nothing has
 * ever observed, because no round recorded whether one of those charts was the
 * one an update touched.
 *
 * SPLIT BY WHETHER THE CHART COULD HAVE STRANDED ANYTHING, because that is the
 * whole point. Growth on a chart that had its parts list is an ordinary change
 * of size; growth on one that was UNGROUPED AND UNLISTED is the stranding.
 * Pooling the two would destroy the only distinction the instrument draws.
 *
 * `atRisk` is that population, counted from the host's own shape type, and the
 * report refuses to call zero growth an all-clear when it is zero — a grouped
 * chart is deleted whole and can strand nothing, so a round that grouped
 * everything never put the question. Round 082 was exactly that: 20 of 20
 * grouped, and it would have read as an all-clear.
 *
 * Readings from before 2026-08-16 carry `shortfall`/`unexplained` instead of
 * `growth` — a subtraction across three different units that summed to zero on
 * every line. They are counted as `unitMismatch` and never mixed in.
 */
export function poolUpdateShortfalls(logs) {
  const out = {
    rounds: 0,
    updates: 0,
    blind: 0,
    blindGrowth: 0,
    sightedGrowth: 0,
    worst: 0,
    unitMismatch: 0,
    ungroupedCharts: 0,
    atRisk: 0,
    /** Readings whose two host reads disagreed — kept, flagged, never pooled. */
    unsettledKept: 0,
    unsettledGrowth: 0,
    /** Shapes #586 left loose ON PURPOSE, and the charts that left them. */
    strandedByDesign: 0,
    subsetGroups: 0,
    deckContradicted: 0,
  };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    // THE DECK IS THE ONE SOURCE THAT HAS NEVER BEEN WRONG HERE. It is taken at
    // the end of the round, long after any host lag has settled, and it has
    // caught three phantom readings running — the 084 investigation and both of
    // round 086's.
    //
    // A round only ADDS shapes to the slides it keeps, so a reading claiming
    // MORE shapes than the slide finished with is claiming shapes that never
    // existed. Round 086 read `after: 24` on a slide the inventory then showed
    // holding 1, and both of its host reads agreed on the 24 — the lag outlasted
    // the settle delay, so `settled` was true and wrong.
    const finalCount = new Map();
    for (const s of log?.deck?.inventory ?? []) {
      const id = s.slideId ?? s.id;
      const n = s.count ?? s.shapes?.length;
      if (id && typeof n === "number") finalCount.set(id, n);
    }
    // WHETHER THE QUESTION COULD BE PUT AT ALL. A round that grouped everything
    // cannot answer this either way, and round 082 was exactly that: 20 grouped,
    // 0 not. Without this the report would read "no growth" from such a round
    // and sound like an all-clear.
    out.ungroupedCharts += entries.filter((e) => /^not grouping/.test(String(e.message))).length;
    // AND "A GROUP IS DELETED WHOLE" STOPPED BEING TRUE ON 2026-08-19. This
    // comment used to say stranding was possible only for an ungrouped chart.
    // #586 groups the majority the host will name and, in its own words, leaves
    // "the stranded remainder deliberately not written into the parts tag" —
    // so a chart can now be GROUPED and still leave shapes loose inside its own
    // box, and the next update deletes the group and walks past them.
    //
    // `atRisk` cannot see it. It is read from the host's own shape type at
    // update time (`powerpoint.ts`), where a subset group and a whole one are
    // the same word, and telling them apart there would cost a load per chart.
    // The ROUND file has both halves though — the draw pass records
    // `partial=N left=i:k` — so the join belongs here, and a zero from `atRisk`
    // can no longer be read as an all-clear on its own.
    for (const e of entries) {
      if (String(e.message) !== "grouped the chart's shapes") continue;
      // `left` is `index:count` per partially-grouped chart, comma-joined.
      for (const pair of String(e.data?.left ?? "").split(",")) {
        const k = Number(pair.split(":")[1]);
        if (Number.isFinite(k) && k > 0) {
          out.strandedByDesign += k;
          out.subsetGroups += 1;
        }
      }
    }
    let seen = false;
    for (const e of entries) {
      if (!/^shapes left on the slide/.test(String(e.message))) continue;
      const d = e.data ?? {};
      // `charts` and `withParts` are no longer read here. They said how many
      // charts an update touched and how many carried a parts list, and this
      // pool used the second as a stand-in for "could strand something" — which
      // round 086 disproved, because a grouped chart has no parts list either.
      // `atRisk` names the population directly. Both fields stay in the trace:
      // they are still worth having when reading a line by hand.
      seen = true;
      out.updates++;
      // ROUND 082 AND EARLIER SPOKE A DIFFERENT LANGUAGE. Those entries carry
      // `shortfall`/`unexplained` — a subtraction across three different units
      // that summed to zero on every line and measured nothing. Counted, never
      // mixed in: pooling them with a real growth would resurrect the artifact
      // this pool was rewritten to stop reporting.
      // Two shapes of unusable reading, counted together because the
      // response is the same: do not pool it. Pre-2026-08-16 entries speak the
      // mismatched-unit language; round 084 speaks `growth` but from a single
      // host read, and every non-zero number it produced was the host lagging a
      // group it had already committed.
      // ABSENT and FALSE are different now, and lumping them lost the
      // distinction. Absent means a pre-2026-08-16 build whose numbers are in
      // mismatched units and cannot be read at all. FALSE means a reading the
      // instrument deliberately kept: the two host reads disagreed, and the
      // SECOND one was taken because across the archive the deck adjudicates all
      // 76 such readings and backs the second 48 times against the first's zero.
      //
      // Still not pooled with settled readings — an unsettled number is weaker
      // evidence and mixing them would hide that. Counted and reported, because
      // a fifth of this instrument's output used to vanish in silence.
      if (d.growth === undefined || d.settled === undefined) {
        out.unitMismatch++;
        continue;
      }
      if (d.settled === false) {
        out.unsettledKept++;
        if (Number(d.growth) > 0) out.unsettledGrowth += Number(d.growth);
        continue;
      }
      // Checked BEFORE the reading is pooled, and counted rather than dropped
      // silently: an instrument's own error rate is a number worth reporting,
      // and this one has had four false readings in two days.
      const ended = finalCount.get(d.slideId);
      if (typeof ended === "number" && Number(d.after) > ended) {
        out.deckContradicted++;
        continue;
      }
      const atRisk = Number(d.atRisk) || 0;
      out.atRisk += atRisk;
      const growth = Number(d.growth) || 0;
      // AT RISK, not merely list-less. This bucketed on `withParts === 0`, and a
      // GROUPED chart has no parts list either — so every grouped chart landed
      // in the stranding column, where by construction it cannot belong: a group
      // is deleted whole and leaves nothing behind. Round 086 put a growth of 23
      // there from a chart whose own line said `atRisk: 0`.
      //
      // `atRisk` is the population the question is about: ungrouped AND with no
      // parts list, read from the host's own shape type. Growth anywhere else is
      // an ordinary change of size or an instrument artifact, and either way it
      // is not stranding.
      if (atRisk > 0) {
        out.blind++;
        out.blindGrowth += growth;
        out.worst = Math.max(out.worst, growth);
      } else out.sightedGrowth += growth;
    }
    if (seen) out.rounds++;
  }
  return out;
}

/**
 * One trace entry reduced to the SHAPE of the thing it says.
 *
 * Numbers and ids are the noise here: `repaired 3 tags on slide 7f2a91c4` and
 * `repaired 5 tags on slide 0b41ee02` are one event happening twice, and a
 * signature that keeps their digits counts them as two things that each
 * happened once. So digits collapse to `N` and hex runs to `#`.
 *
 * THE ERROR CLASS IS PART OF THE SHAPE, and leaving it out cost a real reading.
 * A first cut hashed `scope` and `message` alone, and round 077's 52
 * `UnexpectedError`s — the single loudest fact in that round — did not appear in
 * its report at all, because they live in `data.error` and every one of them
 * shared a message with its healthy counterpart. Folding the class in costs
 * almost nothing: the archive's whole vocabulary grew from 74 signatures to 81
 * across 58 rounds.
 *
 * The class only, never the full text — error strings carry ids and offsets, and
 * hashing those would give every failure its own signature and report a
 * vocabulary that grows forever.
 */
export function traceSignature(entry) {
  const norm = (s) =>
    String(s ?? "")
      .replace(/[0-9a-f]{8,}/gi, "#")
      .replace(/\d+/g, "N");
  const err = entry?.data?.error;
  const cls = err ? "!" + (norm(err).match(/[A-Za-z][A-Za-z0-9_]+/)?.[0] ?? "err") : "";
  return `${entry?.scope ?? "?"}|${norm(entry?.message).slice(0, 60)}${cls}`;
}

/**
 * What did the NEWEST round say that its 57 predecessors did not?
 *
 * The instrument this exists to replace is a person reading 95K characters of
 * trace. Measured on round 081: 512 entries, 44 distinct signatures, NONE of
 * them new against every prior round. That is the ordinary case, and paying
 * ~50k tokens to rediscover it is the waste — two rounds a night is ~100k tokens
 * of reading to learn nothing.
 *
 * THIS DOES NOT REPLACE READING THE TRACE. It routes it. `docs/ROUNDS.md` is
 * explicit that the headline is never all of a round says, and a summary that
 * stood in for the trace would be that mistake with a script behind it. What
 * this does is count, which a person reading 512 entries cannot do reliably —
 * rounds 077 and 078 are the same two signatures at the same two counts, and
 * establishing that by eye took an entire extra round.
 *
 * THREE BUCKETS, AND THE SPLIT IS THE POINT. A count that leaves a baseline it
 * had is a different event from a shape the archive has barely produced, and
 * that one from a shape it has never produced at all.
 *
 * The archive shows all three. At round 077 the `UnexpectedError` signatures
 * were NOVEL — a first appearance, 18 of them — while `settle pass: repaired
 * every config tag` beside them had been seen rarely and jumped to 14. At rounds
 * 079 and 081 the only report is `re-reading the slide's shapes again after a
 * settle delay`, 11 times against a median of 0: the re-read EXISTED before
 * build 17a8204 and was rare, and `needsPreGroupRefresh` widened which charts
 * reach it. That is a fix working, reported as such. Filing it under the same
 * heading as round 077's errors would mean every fix this project lands
 * announces itself as a fault on the night it starts working, and
 * `docs/BACKLOG.md` records what happens to a report like that: it gets ignored,
 * then switched off.
 *
 * It is the same mistake `profileDivergence` made on its first live outing —
 * collapsing two causes into one worst-case reading — caught there by reading
 * the output against the rounds it described. Split here from the start.
 */
export function traceNovelty(rounds, { minCount = 10, factor = 3, window = 5 } = {}) {
  const entriesOf = (r) => (Array.isArray(r?.trace?.entries) ? r.trace.entries : []);
  const countsOf = (r) => {
    const m = new Map();
    for (const e of entriesOf(r)) {
      const s = traceSignature(e);
      m.set(s, (m.get(s) ?? 0) + 1);
    }
    return m;
  };
  if (!rounds?.length) return { build: null, novel: [], sinceBuild: [], spikes: [], priors: 0, vocabulary: 0 };
  const newest = rounds[rounds.length - 1];
  const priors = rounds.slice(0, -1).map(countsOf);
  const vocabulary = new Set(priors.flatMap((m) => [...m.keys()]));
  const seen = new Set(vocabulary);
  const novel = [];
  const sinceBuild = [];
  const spikes = [];
  // A BASELINE NEEDS ROUNDS TO BE A BASELINE. With three prior rounds a median
  // is whatever the middle one happened to do, and every count looks like a
  // spike against it. The archive has 57, so this only guards the fresh-clone
  // case — but it guards it silently rather than reporting an alarm built on
  // nothing.
  const enough = priors.length >= 5;
  for (const [sig, n] of countsOf(newest)) {
    if (!seen.has(sig)) {
      novel.push({ sig, n });
      continue;
    }
    if (!enough || n < minCount) continue;
    // A RECENT WINDOW, not the whole archive. Taking the median over every
    // prior round meant a signature stayed "new behaviour" until it had
    // appeared in more than HALF the archive — so one that shipped twenty
    // rounds ago was still being announced as new, and the denominator kept
    // growing underneath it.
    //
    // Measured: `re-reading the slide's shapes again after a settle delay` first
    // appeared in round 064 and has sat at 10-11 ever since. It was reported as
    // NEW BEHAVIOUR in fifteen separate rounds and blamed on nine different
    // builds, the last of them a commit that only changed a slide counter. That
    // is the "cries wolf, gets switched off" failure this file's own header
    // warns about, and it drowned the one signature round 086 had actually
    // changed.
    //
    // Five rounds is the same order as this project's noise floor — 1-versus-5
    // for the same fault with nothing changed — so a signature absent from all
    // five and present now is genuinely new to recent history.
    const recent = priors.slice(-window);
    const vals = recent.map((m) => m.get(sig) ?? 0).sort((a, b) => a - b);
    const median = vals[Math.floor(vals.length / 2)];
    // The median is the honest centre for a host whose mood swings 4-of-5 to
    // 1-of-5 with nothing changed; a mean would let one bad night set the
    // baseline for every round after it.
    if (median === 0) {
      // THE BUILD IT STARTED IN, not the build being judged. This reported the
      // newest round's build unconditionally, which is how nine innocent
      // commits got named for one 064-era signature.
      const at = priors.findIndex((m) => (m.get(sig) ?? 0) > 0);
      const startedIn = at === -1 ? null : String(rounds[at]?.build ?? "").split(" ")[0] || null;
      sinceBuild.push({ sig, n, median, startedIn });
    } else if (n > factor * median) spikes.push({ sig, n, median });
  }
  const bySize = (a, b) => b.n - a.n;
  return {
    build: String(newest?.build ?? "").split(" ")[0] || null,
    novel: novel.sort(bySize),
    sinceBuild: sinceBuild.sort(bySize),
    spikes: spikes.sort(bySize),
    priors: priors.length,
    vocabulary: vocabulary.size,
  };
}

const invokedDirectly = process.argv[1] && process.argv[1].endsWith("triage.mjs");
if (invokedDirectly) {
  const args = process.argv.slice(2);
  const flags = args.filter((a) => a.startsWith("--"));
  const paths = args.filter((a) => !a.startsWith("--"));
  const deckPath = paths.find((p) => p.endsWith(".pptx"));
  // A DIRECTORY expands to the rounds inside it, which is how `npm run rounds`
  // stays correct as the archive grows. The obvious alternative — listing the
  // files in package.json, or `rounds/*.json` — fails twice over: the list has
  // to be edited for every round anyone adds, and npm scripts run through
  // cmd.exe on Windows, which does not expand a glob at all. Both failures are
  // silent, and both end with the pooled view quietly reading fewer rounds than
  // the archive holds, which is the exact state this archive exists to end.
  const logPaths = paths.flatMap((p) => {
    if (p.endsWith(".json")) return [p];
    let entries;
    try {
      entries = readdirSync(p);
    } catch {
      return [];
    }
    // ROUNDS ONLY — `NNN-<build>.json`, the naming `test/rounds.test.ts`
    // enforces. Taking every `.json` in the directory looked obviously right and
    // was wrong within the hour: `rounds/predictions.json` lives beside them, so
    // the ledger was counted as a round. The stability header said 5 rounds
    // while the pooling underneath said 4, and the open prediction was judged
    // against the ledger rather than against the newest round — reporting
    // "not on this sheet" for a question that was answered.
    return entries
      .filter((f) => /^\d{3}-.*\.json$/.test(f))
      .sort()
      .map((f) => join(p, f));
  });
  /**
   * The round REPORTED ON is the newest; every round is still pooled.
   *
   * A directory expands sorted, and this read `[0]` — the OLDEST round in the
   * archive. So `npm run rounds`, the one command a loop glances at between
   * rounds, printed two-day-old deck evidence and a two-day-old self-test above
   * a current grid, with nothing saying they came from different days. Nobody
   * noticed for nineteen rounds because the grid underneath was right.
   *
   * An explicit list of files keeps its order — someone naming two rounds means
   * the first one — so only the directory case flips.
   */
  const namedDirectly = paths.some((p) => p.endsWith(".json"));
  const logPath = namedDirectly ? logPaths[0] : logPaths[logPaths.length - 1];
  if (!logPath) {
    console.error("usage: node scripts/triage.mjs <deck.pptx> <run-log.json> [--all] [--json]");
    console.error("       node scripts/triage.mjs <crashed-run.json>            (no deck needed)");
    console.error("       node scripts/triage.mjs <round-*.json>                (pools the counterbalanced arms)");
    process.exit(2);
  }
  let log;
  try {
    log = JSON.parse(readFileSync(logPath, "utf8"));
  } catch (err) {
    console.error(`could not read the run: ${err.message}`);
    process.exit(2);
  }
  // Every round given, for the pooled view. Unreadable ones are skipped rather
  // than fatal: pooling is a bonus on top of reading the FIRST file, and one
  // bad archive should not stop the run in front of you being triaged.
  const pooled = logPaths
    .map((p) => {
      try {
        const parsed = JSON.parse(readFileSync(p, "utf8"));
        // WHICH ROUND THIS IS, from the filename, because nothing inside the
        // file says. `dormantInstruments` counts silence in rounds and cannot do
        // it from a build stamp shared by a whole pair. Attached here rather
        // than re-derived, so there is one place that knows the naming.
        parsed.roundName = p.replace(/^.*[\\/]/, "");
        return parsed;
      } catch {
        return null;
      }
    })
    .filter(Boolean);
  // A crashed run answers on its own. There is no deck to join it to — the run
  // never reached the point of producing one, which is the whole reason this
  // file exists — so requiring a .pptx would refuse the one artifact a crash
  // leaves behind.
  const crash = crashLogIn(log);
  if (crash && !deckPath) {
    const ran = crash.finishedAt ? "finished" : "NEVER REPORTED FINISHING";
    console.log(
      `\n  CRASHED RUN "${crash.label}" — ${ran}` +
        `\n  build ${crash.build} · ${crash.host}` +
        `\n  started ${crash.startedAt}${crash.dropped ? ` · ${crash.dropped} earlier step(s) dropped` : ""}` +
        `\n  ${crash.steps.length} step(s), OLDEST FIRST — the last line is where it stopped\n`,
    );
    for (const line of crash.steps) console.log(`    ${line}`);
    // What the run CONCLUDED, as opposed to what it narrated.
    //
    // A crashed round's steps say a scenario started; they never said what it
    // decided, because verdicts only existed in memory until the run ended and
    // these runs do not end. They are recorded as they land now, so a crash log
    // carries them — and this is the reader that has to show them, or they are
    // back to being invisible in a different place.
    if (crash.findings?.length) console.log(`\n  ${crash.findings.length} finding(s) banked before it stopped:\n`);
    for (const { key, value } of crash.findings ?? []) {
      console.log(`    ${key}: ${describeFinding(value)}`);
    }
    const known = crash.steps.map((l) => knownBug(l)).filter(Boolean);
    if (known.length) console.log(`\n  known host bug: ${[...new Set(known)].join("\n  known host bug: ")}`);
    console.log("");
    process.exit(crash.finishedAt ? 0 : 1);
  }
  // A round with no .pptx is the NORMAL case now, and this used to refuse it.
  //
  // The starred runbook step tells the owner to send one file and nothing else,
  // because the round carries every shape on every slide and the host's own
  // picture of each — "there is no deck to save and no screenshot to take". A
  // file produced by following that instruction landed here on 2026-08-09 with
  // 12 self-test verdicts and 246 trace entries, and got `usage:` and exit 2.
  // It was then read by hand, which is the job this script exists to remove.
  //
  // Only the slot join needs the .pptx. The self-test, the trace and the round's
  // own deck evidence are all readable without one, so they are reported and the
  // join is skipped — and the exit code still means what it meant.
  if (!deckPath) {
    const selftest = selfTestIn(log);
    if (!selftest.length && !log?.trace && !log?.deck) {
      console.error("usage: node scripts/triage.mjs <deck.pptx> <run-log.json> [--all] [--json]");
      console.error("       node scripts/triage.mjs <crashed-run.json>            (no deck needed)");
      console.error("       node scripts/triage.mjs <round.json>                  (self-test / trace half)");
      process.exit(2);
    }
    const runs = runsIn(log).filter((r) => r?.run);
    console.log(
      `\n  ROUND — no deck supplied, so slots are not joined` +
        `\n  build ${log.build ?? "?"} · ${log.host ?? "?"}` +
        // HOW THE ROUND WENT BEFORE IT PRODUCED ANYTHING. `driverRun` has been
        // archived since it was added and NEVER printed, so a round rescued from
        // a crash read exactly like one that never needed rescuing. On
        // 2026-08-24 that cost two wrong reports of the same block: three
        // crashes went unnoticed, and were then attributed to the wrong rounds
        // — while the field naming them sat in every file the whole time.
        driverRunLine(log.driverRun) +
        roundIdentityLine(log) +
        (runs.length ? `\n  ${runs.length} insert run(s) in this file — pass the .pptx to check their slots` : ""),
    );
    reportDeckEvidence(log.deck);
    reportTrace(log.trace);
    reportBatchCost(log);
    const failed = reportSelfTest(selftest);
    reportStability(pooled);
    reportStalls(pooled);
    reportOccupancyStalls(pooled);
    reportRoundAnatomy(pooled);
    reportScenarioCost(pooled);
    reportScratchChurn(pooled);
    reportPassBias(pooled);
    reportDormant(pooled);
    reportNoiseFloor(pooled);
    reportClaims(pooled);
    reportPredictions(pooled);
    reportFreshVsEstablished(pooled);
    reportGroupVsTag(pooled);
    reportBatchSpanVsGroup(pooled);
    reportOriginTagLosses(pooled);
    reportUpdateShortfalls(pooled);
    reportScenarioFriction(pooled);
    reportStarvedQuestions(pooled);
    reportIdChurn(pooled);
    reportPartsListOutcome(pooled);
    reportPositionalGuess(pooled);
    reportSettleAsks(pooled);
    reportCrashes();
    reportOccupancyCost(pooled);
    reportUpdateCost(pooled);
    reportFlatFields(pooled);
    reportDrawCostCurve(pooled);
    reportTagRoutes(pooled);
    reportTagFaults(pooled);
    reportPool(pooled);
    process.exit(failed ? 1 : 0);
  }
  let deck;
  try {
    deck = await readDeck(deckPath);
  } catch (err) {
    console.error(`could not read the run: ${err.message}`);
    process.exit(2);
  }
  const runs = runsIn(log);
  const results = runs.map((run) => ({ run, t: triage(deck, run) }));
  const selftest = selfTestIn(log);
  const faults = faultsIn(deck);
  if (flags.includes("--json")) {
    console.log(
      JSON.stringify({ faults, selftest, runs: results.map(({ run, t }) => ({ path: run.path, ...t })) }, null, 2),
    );
  } else {
    // Faults belong to the FILE, not to a run, so they are reported once
    // above the runs rather than repeated under each of them.
    if (faults.length) {
      console.log(`\n  ${faults.length} STRUCTURAL FAULT(S) — this repo wrote the file wrong:`);
      for (const f of faults) console.log(`    - ${f}`);
    }
    for (const { run, t } of results) report(deck, log, run, t, flags.includes("--all"));
    reportDeckEvidence(log.deck);
    // The trace belongs to the FILE, like the faults above, and it was reachable
    // only from the no-deck branch. So the tool's own documented invocation —
    // `triage.mjs <deck.pptx> <run-log.json>`, the one in its usage line and in
    // CLAUDE.md — printed 28 FEWER lines than the degraded one-file form and
    // said nothing about it. What went missing: the entry histogram, "phases an
    // error escaped", the problems tally, and every `known host bug: office-js#…`
    // annotation — the most locating lines in the whole log. A round carrying a
    // trace and no self-test went further and reported "this log holds no runs
    // and no self-test" over 186 entries, then exited 0.
    reportTrace(log.trace);
    reportBatchCost(log);
    reportSelfTest(selftest);
    reportStability(pooled);
    reportStalls(pooled);
    reportOccupancyStalls(pooled);
    reportRoundAnatomy(pooled);
    reportScenarioCost(pooled);
    reportScratchChurn(pooled);
    reportPassBias(pooled);
    reportDormant(pooled);
    reportNoiseFloor(pooled);
    reportClaims(pooled);
    reportPredictions(pooled);
    reportFreshVsEstablished(pooled);
    reportGroupVsTag(pooled);
    reportBatchSpanVsGroup(pooled);
    reportOriginTagLosses(pooled);
    reportStarvedQuestions(pooled);
    reportIdChurn(pooled);
    reportPartsListOutcome(pooled);
    reportPositionalGuess(pooled);
    reportTagRoutes(pooled);
    reportTagFaults(pooled);
    reportPool(pooled);
    if (!results.length && !selftest.length && !log?.trace)
      console.log("\n  this log holds no runs and no self-test\n");
  }
  const disagreements =
    results.reduce((n, { t }) => n + t.disagreements, 0) + selftest.filter((s) => !s.ok && !s.skipped).length;
  process.exit(disagreements || faults.length ? 1 : 0);
}
