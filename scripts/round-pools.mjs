/**
 * THE POOLED READERS: what the whole archive says, as opposed to one round.
 *
 * Every function here takes a list of parsed round logs and returns numbers.
 * None of them reads a file, asks the clock, or looks at `process`. That is not
 * a style preference — it is the entire reason this module exists.
 *
 * ── WHY IT IS ITS OWN FILE, IN THE CONFIG'S OWN WORDS ───────────────────────
 * `stryker.config.json` asked for this by name, and it recorded the cost of not
 * having it:
 *
 *   "The pooled readers in scripts/triage.mjs are outside `mutate`, and on
 *   2026-08-23 FOUR tests written for them passed against the code they were
 *   written to catch — a reset test that a fresh module made vacuous, a batch
 *   fixture whose two entries never reached the case that differs, a confound
 *   fixture whose contaminating rows could not outvote a median, and a
 *   same-slide control that a filter dropped."
 *
 * Every one of those four was found by stashing the source by hand. The repo's
 * flagship rule — a regression test must be proven to fail without its fix — is
 * enforced by whoever remembers to do that, and nobody has ever checked the
 * other two thousand assertions.
 *
 * The config also named both things that blocked automating it, and both were
 * real: `test/triage.test.ts` cannot run in Stryker's sandbox because it reads
 * repo files by a cwd-relative path, and `triage.mjs` was 6,655 lines, so
 * adding it whole risked the never-finishing run that `ignoreStatic` describes.
 * Splitting the pure half off answers both at once. The pools are now mutated
 * and their tests live in `test/round-pools.test.ts`, which is NOT excluded.
 *
 * ── SO KEEP IT PURE, AND THE TEST IS NOT A STYLE RULE ───────────────────────
 * One `readFileSync` in here and `test/round-pools.test.ts` stops being runnable
 * in the sandbox, which puts the whole file back outside `mutate` — silently,
 * because everything would still pass. `test/round-pools.test.ts` asserts the
 * absence, so that cannot happen quietly.
 *
 * The one reader that does reach disk stayed behind on purpose:
 * `dormantInstruments` takes its source as a default argument
 * (`source === undefined ? traceSource() : source`) and `traceSource` walks
 * `src/` relative to the working directory. It is in `triage.mjs`.
 *
 * ── AND THE THIRD PLACE THAT HAS TO AGREE ───────────────────────────────────
 * Adding this file to `mutate` in `stryker.config.json` is not enough on its
 * own. The weekly job runs `npx stryker run --mutate "$SCOPE"`, and a CLI
 * `--mutate` OVERRIDES the config — so the scope that actually applies comes
 * from `MUTATED` in `scripts/mutation-scope.mjs`. Widen the config and not that
 * regex and every change here is dropped, the job prints "no src/core changes",
 * and this file is never mutated once: a green run that measures nothing, which
 * is the exact failure the whole extraction exists to stop.
 * `test/mutation-scope.test.ts` pins it from the other side.
 *
 * ── .mjs, NOT .ts ───────────────────────────────────────────────────────────
 * `scripts/rounds-gate.mjs` imports these and is plain `.mjs` run by node
 * directly, with no build step. It cannot import a `.ts`.
 *
 * Nothing is imported here, and that is worth keeping: a module with no imports
 * cannot be part of a cycle and cannot pull the filesystem in by accident.
 */

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
 * Above this many shapes on one slide, the deck is worth LOOKING at.
 *
 * It used to say "a chart did not group", and that was wrong — corrected
 * 2026-09-16 against the archive rather than by argument.
 *
 * Five is what a clean round's fullest slide holds — one grouped chart per
 * probe slide, and the four one-batch charts the rasterise scenario adds beside
 * one of them. The old reasoning went on: "a loose chart contributes eight or
 * more on its own, so there is no near-miss band to argue about". The premise
 * is sound and the INFERENCE does not follow, because three different things
 * reach eight:
 *
 *   eight grouped charts      eight shapes named `PowerChart` — `GROUP_NAME` is
 *                             a LOOKUP KEY in powerpoint.ts, so a grouped chart
 *                             is ONE shape by that name. Rounds 443-455: eight
 *                             such rounds, every one benign.
 *   a title slide             `Title 1`, `Subtitle 2` and friends. Three rounds.
 *   an exploded chart         `explode a degraded picture` turns a picture into
 *                             native shapes ON PURPOSE. Round 457's title slide
 *                             holds `title, category-0..3, seg-0-0..seg-1-3,
 *                             baseline, series-label-0/1` for exactly that
 *                             reason.
 *
 * Only round 449 in that window was the real thing. Eleven of twelve firings
 * were benign, so a count here is a 92% false alarm — and NAMES cannot save it
 * either, since an exploded chart and a chart that failed to group leave a slide
 * that is identical in both count and naming.
 *
 * The question has a direct answer elsewhere: the grouping counter, attempts
 * against refusals, read from the trace. Use that. This constant now only says
 * "unusual enough to read the round".
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

/**
 * Each fallback's count this round against the median of its priors.
 *
 * THREE PRIORS MINIMUM, the same rule `poolScenarioPopulations` and
 * `poolGroupingOutcome` needed: a "usually" from one observation is not a
 * baseline, and this project's own noise floor is the argument.
 */
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
 * The build that first carried the SECOND SETTLE ASK (#709, `REREAD_ATTEMPTS` 1 -> 2).
 *
 * A round records its build, not the pull requests inside it, so this is the
 * only durable handle on "the era in which the change was live". Recheck it
 * with `git log --oneline -S "REREAD_ATTEMPTS = 2" -- src/render/powerpoint.ts`
 * if the constant ever stops matching a round in the archive.
 */
export const SECOND_ASK_BUILD = "2912802";

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
