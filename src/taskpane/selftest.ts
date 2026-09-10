/**
 * Eleven things that have never once run against a real PowerPoint, on one
 * click — plus two that only make sense on their own.
 *
 * The demo deck covers inserting onto slides it added BLANK. Nothing covers
 * what happens elsewhere — inserting on top of an earlier run, a slide
 * duplicated so two claim one slot, redrawing a chart the user is looking at,
 * editing the chart the user SELECTED, drawing onto a slide that already has
 * content, a deck-wide rescale, stopping part-way, turning a degraded picture
 * back into shapes, and whether any of it is visible. Every one of those paths
 * is guarded, and every guard has only ever been checked against a fake host.
 * The list sat in `docs/PUBLISHING.md` as separate things for a human to
 * remember to try, which in practice meant a session each: deploy, click
 * through it, save the deck, upload, read.
 *
 * The blank-slide bias is not incidental. `insertSceneIntoSlide` — the
 * everyday "put a chart on the slide I am looking at" — draws onto whatever
 * the user already has there, and no test anywhere touched that until
 * `insertOntoUsedSlide` below. That is exactly where the worst bug of the
 * session lived, and why a demo run could not have found it.
 *
 * Scenarios run in order and leave their slides in the deck, because the point
 * is a file someone can open, look at, and hand to `npm run triage`. The one
 * exception is `chartIsVisible`, which takes its control slide away again.
 *
 * **On selection.** This file used to say Office.js had no way to select a
 * shape, and that the pane's selection-driven entry points therefore could not
 * be scripted. It was wrong, in the most expensive way a comment can be: it
 * justified a hole with a fact, so nobody re-checked it for months.
 * `Slide.setSelectedShapes(shapeIds)` has been GA since **PowerPointApi 1.5** —
 * the same set this add-in already requires for `getSelectedShapes` and
 * `setSelectedSlides`. The mistaken belief was that it would live on
 * `Presentation` beside `setSelectedSlides`; it is one class down, on `Slide`.
 * `editViaSelection` is what that sentence had been preventing.
 *
 * **Order is load-bearing, because of two live web-host bugs.** On PowerPoint
 * on the web `setSelectedShapes([])` does not clear the selection
 * (office-js#3083), and a picture cannot be inserted while another shape is
 * selected (office-js#3698). So a scenario that leaves a chart selected breaks
 * the PICTURE scenario rather than its own — a failure that reads, in a run
 * log, as a bug in code that is fine. Every selecting scenario clears up after
 * itself through `clearShapeSelection`, which re-selects the slide rather than
 * trusting the empty-array clear.
 */
import type { ChartConfig, ChartKind } from "../core/types";
import type { Scene } from "../core/scene";
import { buildChart } from "../core/chart";
import { sampleConfig } from "../core/samples";
import { buildDeckBase64 } from "../render/pptx-deck";
import {
  deckSlideIds,
  hostFrictionCounts,
  canInsertPicture,
  canSelectShapes,
  clearShapeSelection,
  isStopped,
  isStopRequested,
  isTimeout,
  stepOf,
  readbackTimeoutMs,
  loadChartFromSelection,
  requestStop,
  resetStop,
  selectShape,
  selectionLadder,
  awaitSelectedChart,
  canWatchSelection,
  slideImageBase64,
  canInsertSlidesFromBase64,
  insertSceneIntoSlide,
  slideOccupancy,
  insertSlidesFromPptx,
  listChartsInDeck,
  chartOnItsOwnSlideAsFile,
  scanIsComplete,
  scanGap,
  newRunId,
  reconcileDeck,
  showSlide,
  slideCount,
  settledSlideCount,
  slideSize,
  slideShapeList,
  addSlideForChart,
  deleteSlideById,
  shapeGeometryByName,
  deleteShapesById,
  timeShapeRounds,
  updateChartInSlide,
  moveShapeBy,
  errorText,
  // A live binding, read fresh on every access — the whole point is to diff it
  // around a scenario.
  deadlinesFired,
  lastStall,
  RASTERISE_OP,
  rasterGap,
  lastLateSync,
  lastLateSyncSeq,
  waitForLateSync,
  shapesDrawnOn,
} from "../render/powerpoint";
import { trace, traceAbout, traceElapsed } from "../core/trace";

/**
 * How long to sit still between the deck scan and the first chart of a run.
 *
 * AN INSTRUMENT, NOT A FIX, and it is here to settle one question.
 *
 * Rounds 204-208 put the first chart of a multi-chart update at 36189-43837ms
 * against 16273-20967ms for the same size later in the same run — n=11, and the
 * two distributions DO NOT OVERLAP. Three explanations are already excluded by
 * the archive: not the slide (18233 against 17033 at 18 of 24, ranges
 * overlapping), not a cold host (an ordinary update on chart 1's OWN slide 87s
 * earlier cost 2330ms against a fit predicting 2491), and not our sync count
 * (four syncs either way, and the fourth is not the expensive one).
 *
 * What mining CANNOT exclude is the deck scan, because every run's chart 1 is
 * preceded by one and no later chart is — 1237ms and 1231ms ahead of it in the
 * two legs of the 207/208 pair. A variable that never varies cannot be split out
 * of an archive however it is read. Only sitting still separates it.
 *
 * OFF UNLESS ASKED, and read per run rather than cached, so both arms can be run
 * against one build without a rebuild:
 *
 *     localStorage.setItem("ssf-charts-scan-settle-ms", "3000")
 *
 * It only PAYS as a product change if the pause is shorter than the ~20s it
 * saves. That is not what this measures — this measures whether the pause moves
 * chart 1 at all. If it does not, the scan is excluded and the search moves on;
 * if it does, the next question is how little of it is enough.
 *
 * TRACED EVERY RUN INCLUDING ZERO. A round that does not say which arm it was in
 * cannot be pooled with one that does, and this archive is read by splitting
 * rounds on exactly this kind of field. An untraced default is how an experiment
 * quietly contaminates its own control.
 *
 * Clamped rather than trusted, because a typo that parks the pane for an hour
 * costs a round: `Number("")` is 0 but `Number("abc")` is NaN, and a NaN
 * setTimeout fires immediately rather than failing loudly.
 */
export const SCAN_SETTLE_KEY = "ssf-charts-scan-settle-ms";
export const SCAN_SETTLE_MAX_MS = 60_000;

export function scanSettleMs(read: (key: string) => string | null = readScanSettle): number {
  let raw: string | null;
  try {
    raw = read(SCAN_SETTLE_KEY);
  } catch {
    // A pane whose storage throws is not a pane that asked for a pause.
    return 0;
  }
  if (raw === null || raw.trim() === "") return 0;
  const n = Number(raw);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return Math.min(Math.round(n), SCAN_SETTLE_MAX_MS);
}

function readScanSettle(key: string): string | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage.getItem(key);
  } catch {
    return null;
  }
}

/** One scenario's verdict, as it goes into the run log. */
export interface ScenarioResult {
  name: string;
  /** False means the scenario ran and the host got it wrong. */
  ok: boolean;
  /**
   * True when the scenario could not run at all — the host lacks the API, or
   * a previous scenario left nothing to work with. Kept apart from a failure
   * on purpose: "we did not check" and "we checked and it is broken" are
   * different answers, and reporting the first as the second is how a
   * requirement-set gap gets diagnosed as a bug.
   */
  skipped?: boolean;
  /**
   * The skip is a HOST failure, not a missing API — nothing was learned.
   *
   * A third state, because the summary used to call every skip a host
   * capability gap and five of the scenarios skip on an empty scan instead. A
   * run in which the host refused every deck read therefore said, in green,
   * "2 of 2 scenarios passed · 6 skipped (host cannot run them)" — a total loss
   * of deck visibility reported as a feature gap, which is the sort of line
   * someone files and moves on from.
   *
   * Two things raise it now. A blinded deck scan, which is what it was written
   * for; and a scenario the host stopped answering part-way through, which used
   * to be reported as an ordinary failure. Both mean the same thing to a
   * reader — this verdict is not evidence about the product — which is why they
   * share a flag rather than getting one each.
   */
  blind?: boolean;
  /** What was actually observed — the sentence a diagnosis starts from. */
  detail: string;
  ms: number;
  /**
   * How much this host misbehaved DURING this scenario.
   *
   * Recorded on the trace line since it existed; carried on the result now
   * because the summary has to read it. `describeSelfTest` used to report every
   * red scenario the same way, so a round in which the host's shape collection
   * died read exactly like a round in which the add-in had a bug — and this
   * project has spent whole sessions on the difference. See `scenarioBlame`.
   */
  /**
   * ALL EIGHT COUNTERS, not the four that existed when this was written.
   *
   * The four-name list had gone stale in four places. A previous fix put the
   * whole snapshot on the `run finished drawing` trace instead — and that line
   * is on the DEMO-DECK path, which no self-test scenario exercises: it has
   * appeared in **0 of 180 archived rounds**, so the counters it carried
   * reached nothing. The commit that added it claimed it put all eight in the
   * round either way. It did not.
   *
   * This delta is the one thing that does reach every round, so it is where
   * the widening belongs. `scenarioBlame` reads three of these by name and is
   * unaffected by the extra five.
   */
  friction?: {
    errors: number;
    idRefusals: number;
    generalExceptions: number;
    emptyReReads: number;
    shortReReads: number;
    unmatchedReReads: number;
    reReadsRepaired: number;
  };
}

/**
 * Whose fault a scenario's outcome was, on the evidence the run already has.
 *
 * The headline counted every red scenario together, which made it useless for
 * the only question worth asking of a series of rounds: is the add-in getting
 * better? On this host most red is weather — `same scale across the deck` is
 * largely a measurement of when the shape collection stops answering — so the
 * number moved with the host's mood and never with our work.
 *
 * The split is evidence-based, never a judgement, and it defaults to OURS.
 * A failure counts as host-degraded only when the run recorded the host
 * actually refusing something inside that scenario: an id it would not resolve,
 * a collection that came back empty, a GeneralException. A scenario that failed
 * while the host answered everything is ours, and says so.
 *
 * That direction matters more than the split. Getting it backwards would turn
 * this into a way to make failures disappear, which is worse than no split at
 * all — so anything unproven lands on us.
 *
 * Checked against a round where the answer was known before the rule existed.
 * On `89675b6`, `explode a degraded picture` failed with `errors: 0,
 * idRefusals: 0, emptyReReads: 0` — the picture regression, a pure logic bug
 * with no host involvement — and `same scale across the deck` failed with
 * `idRefusals: 6, emptyReReads: 4`. Ours and the host's, separated correctly,
 * on data recorded before anyone was looking for it.
 */
export function scenarioBlame(r: ScenarioResult): "passed" | "not-run" | "ours" | "host" {
  if (r.skipped) return "not-run";
  if (r.ok) return "passed";
  const f = r.friction;
  /**
   * A REFUSAL THE CODE RECOVERED FROM IS NOT WHAT DEFEATED THE RUN.
   *
   * The rule above is sound and the round it was checked against proves it — for
   * `same scale across the deck`, whose failures carry this signature 100% of
   * the time and whose passes carry it 14%. That is a discriminator.
   *
   * For `explode a degraded picture` it discriminates NOTHING. Measured over all
   * 322 archived rounds: 265 of 269 passes (99%) and 46 of 47 failures (98%).
   * A condition that is simply always true there cannot be evidence of anything,
   * so every failure of the add-in's most-failing update path was being written
   * off as the host's weather on a reading a passing round shows just as often.
   * Getting the right answer from a signal that carries no information is luck,
   * and it runs out silently.
   *
   * Two changes, chosen by measuring candidates against the archive rather than
   * by reasoning:
   *
   * - A refusal that was REPAIRED does not count. The renderer already re-reads
   *   the slides when a by-id lookup refuses, and `reReadsRepaired` counts the
   *   times that worked. If the run recovered, that refusal is not what beat it.
   * - `generalExceptions` drops out of the test. For `explode` it is a CONSTANT
   *   — exactly 1 in every round the friction report names as "a constant, not a
   *   signal" — and a constant is not evidence by definition.
   *
   * The result on the two scenarios that actually fail:
   *
   *     explode a degraded picture   99% pass / 98% fail  ->  48% / 98%
   *     same scale across the deck   14% pass / 100% fail ->  14% / 100%
   *
   * The broken one gains a 50-point gap; the working one is untouched.
   *
   * HONEST LIMIT: this changes no archived verdict. All 109 recorded failures
   * carrying friction are blamed the same way by both rules, because every
   * `explode` failure so far had an UNREPAIRED refusal as well. The case this
   * catches — an `explode` failure whose only refusal was repaired — has not
   * happened yet. The point is that when it does, the old rule would still have
   * said "host", and nothing in the archive could have shown that it was wrong.
   */
  const unrepaired = !!f && (f.idRefusals ?? 0) > (f.reReadsRepaired ?? 0);
  const refused = !!f && (unrepaired || (f.emptyReReads ?? 0) > 0 || (f.unmatchedReReads ?? 0) > 0);
  return refused ? "host" : "ours";
}

/**
 * How long the selection scenario waits on a host that has stopped answering.
 *
 * Short on purpose. The default budget exists to break an infinite wait; this
 * one exists because we already KNOW what the wait ends in on the web, and
 * three minutes of a person's evening to re-learn it is not a test.
 *
 * Capped by the default rather than replacing it: ten seconds is shorter than
 * the ninety this normally sits under, and far longer than the milliseconds a
 * test shortens that to.
 */
const selectionBudgetMs = (): number => Math.min(10_000, readbackTimeoutMs());

const cfg = (title: string, kind: ChartKind = "clustered"): ChartConfig => ({
  ...sampleConfig(kind),
  title,
});

/** A small deck of tagged charts, built once and inserted however a scenario needs. */
async function buildProbe(
  run: string,
  titles: string[],
): Promise<{ base64: string; scenes: Scene[]; shapesPerSlide: number[] }> {
  const configs = titles.map((t) => cfg(t));
  const scenes = configs.map(buildChart);
  const built = await buildDeckBase64(
    configs.map((c, i) => ({
      scene: scenes[i],
      title: titles[i],
      configJson: JSON.stringify(c),
      slot: i,
      run,
    })),
    // At the destination's size, like every other deck build — the self-test
    // exists to exercise the real path, and a 16:9 file inserted into a 4:3
    // deck is exactly the mismatch it should be catching rather than creating.
    await slideSize(),
  );
  return { base64: built.base64, scenes, shapesPerSlide: built.shapesPerSlide };
}

/**
 * Charts this self-test put in the deck, by the title it gave them — AND
 * whether the scan that found them could see the whole deck.
 *
 * Five scenarios draw a conclusion from this coming back empty ("no probe chart
 * to edit", "fewer than two to scale together", and so on) and one draws a
 * conclusion from it coming back empty that it reports as a PASS. All six were
 * reading a blind scan as a fact about the deck. A real host produced
 * `unread=8 slides=8` between two scans that worked, so total blindness here is
 * transient and lands on whichever scenario happens to ask during it.
 */
async function probeCharts(prefix: string) {
  const scan = await listChartsInDeck();
  return {
    found: scan.charts
      .map((c) => ({ ...c, cfg: JSON.parse(c.configJson) as ChartConfig }))
      .filter((c) => typeof c.cfg.title === "string" && c.cfg.title.startsWith(prefix)),
    /** The scan could not see the whole deck — nothing it did NOT find is news. */
    blind: !scanIsComplete(scan),
    gap: scanGap(scan),
  };
}

/**
 * A skip caused by the deck scan going blind, phrased so it cannot be read as a
 * capability gap.
 *
 * `describeSelfTest` used to attribute every skip to "host cannot run them",
 * and five of the nine skip on an empty scan rather than on a missing API. A
 * run in which the host refused every deck read therefore reported, in green,
 * "2 of 2 scenarios passed · 6 skipped (host cannot run them)" — a total loss of
 * deck visibility filed as a feature gap.
 */
const blindSkip = (gap: string) => ({
  ok: false,
  skipped: true,
  blind: true,
  detail: `the deck scan could not see the whole deck (${gap}), so nothing it did not find means anything`,
});

/**
 * How many slides the deck gained between two counts — or null when either
 * count was never taken.
 *
 * `slideCount()` RETURNS UNDEFINED ON A HOST THAT WILL NOT ANSWER, and does so
 * on purpose: its own docstring says "No safe default here, deliberately …
 * Not knowing has to stay distinguishable from knowing." Three callers here
 * subtracted the two numbers anyway, and unknown arrived dressed as a
 * measurement in two different costumes.
 *
 * ROUND 360 IS THE FIRST COSTUME. `insertTwice` reported **"deck grew by NaN
 * (want 4); 0 of 4 charts re-editable"** as a hard FAILED verdict — a sentence
 * that reads as measured data loss and sends a maintainer after the insert
 * path. Nothing had been measured; the host would not say how big the deck was.
 *
 * The other two are worse, because they are silent. `after !== before` is
 * FALSE when both are undefined, so an unreadable count reads as "the deck did
 * not grow" — the exact thing those lines exist to detect — and the scenario
 * passes. A guard that cannot fire on no evidence is not a guard.
 *
 * THIS IS THE SWEEP THE SCAN VERSION NEVER GOT. `insertOntoUsedSlide` was
 * fixed for precisely this shape — the comment inside `insertTwice` says so —
 * and its siblings were left. That was about a blind SCAN; this is the same
 * mistake one measurement over, in the COUNTS, and it survived the first sweep
 * because the two look nothing alike at the call site.
 */
export function deckGrowth(before: number | undefined, after: number | undefined): number | null {
  return typeof before === "number" && typeof after === "number" && Number.isFinite(before) && Number.isFinite(after)
    ? after - before
    : null;
}

/** What a scenario says when the host would not give it a slide count. */
const countlessSkip = () => ({
  ok: false,
  skipped: true,
  /**
   * BLIND, exactly as a lost scan is — and the first version of this omitted
   * the flag. `blind`'s own docstring says what that costs: a skip without it
   * reads as "host cannot run them", a feature gap someone files and moves on
   * from, when what happened is that we lost sight of the deck. An unread COUNT
   * and an unread SCAN are the same loss of visibility and have to reach a
   * reader as the same thing.
   */
  blind: true,
  detail: "the host would not say how many slides the deck holds, so nothing about what it gained means anything",
});

type Scenario = (prefix: string) => Promise<{ ok: boolean; detail: string; skipped?: boolean; blind?: boolean }>;

/**
 * Inserting a demo deck on top of an earlier one.
 *
 * Before the run token existed this was destructive: every item matched two
 * slides, and the repair pass deleted a healthy run as the other one's
 * duplicate. The token is what makes it safe, and nothing has ever confirmed
 * that against a real host.
 */
const insertTwice: Scenario = async (prefix) => {
  if (!canInsertSlidesFromBase64()) return { ok: false, skipped: true, detail: "host has no insertSlidesFromBase64" };
  const titles = [`${prefix} twice A`, `${prefix} twice B`];
  const before = await slideCount();
  for (const _ of [0, 1]) {
    const { base64 } = await buildProbe(newRunId(), titles);
    await insertSlidesFromPptx(base64, titles.length);
  }
  const after = await slideCount();
  // A blind READBACK is not a finding. Every scenario here guards its first
  // scan and then draws its loudest conclusion from a second one it never
  // checked — and on the web a short deck scan is routine, which is why
  // `DeckScan.short` exists at all. `insertOntoUsedSlide` was fixed for exactly
  // this and its five siblings were never swept with it, so a host that answered
  // one page short produced a hard FAILED verdict asserting data loss the
  // add-in did not cause. Those sentences are what send a maintainer after the
  // tag-write path; this repo has already spent rounds on that hunt.
  const { found, blind, gap } = await probeCharts(`${prefix} twice`);
  if (blind) return blindSkip(gap);
  // An unread count is not a growth of NaN — see `deckGrowth`. Round 360 filed
  // this scenario FAILED with "deck grew by NaN (want 4)", which reads as
  // measured data loss and was nothing of the kind.
  const grew = deckGrowth(before, after);
  if (grew === null) return countlessSkip();
  const ok = grew === 4 && found.length === 4;
  return {
    ok,
    detail: `deck grew by ${grew} (want 4); ${found.length} of 4 charts re-editable`,
  };
};

/**
 * Two slides claiming one (run, slot) — what duplicating a demo slide produces.
 *
 * The same bytes inserted twice carry the same token, which is exactly what
 * PowerPoint's own Duplicate Slide does to a slot tag. The repair pass is
 * supposed to drop one and keep a working one; getting that backwards deletes
 * the user's chart.
 */
const duplicateSlot: Scenario = async (prefix) => {
  if (!canInsertSlidesFromBase64()) return { ok: false, skipped: true, detail: "host has no insertSlidesFromBase64" };
  const titles = [`${prefix} dup A`, `${prefix} dup B`];
  const run = newRunId();
  const { base64, shapesPerSlide } = await buildProbe(run, titles);
  const before = await slideCount();
  const landedA = await insertSlidesFromPptx(base64, titles.length);
  const landedB = await insertSlidesFromPptx(base64, titles.length);
  /**
   * THE RECONCILE'S WINDOW IS SIZED FROM A COUNT THIS HOST REPORTS LATE, and
   * until 2026-08-29 that produced a hard FAILED verdict about half the slides.
   *
   * Round 297 stopped a whole cycle. The first of the two inserts reported
   * `landed: 0` for slides that DID land, `slideCount()` answered 5 when the
   * deck held 7, and `reconcileDeck` was handed `[3,5]` — two slides of four.
   * Two distinct titles in that window, no duplicates, `0 queued`. Every step
   * correct about the wrong population, and the scenario called it a failure.
   *
   * Not the flake this scenario already has: all five earlier failures had the
   * dedup RUNNING and leaving slides behind. Not a blind host either — `unread`
   * was 0 on that read, so the `blindSkip` below could not see it. The host was
   * BEHIND, which is what `settledSlideCount` is for.
   */
  const wanted = titles.length * 2;
  const afterInsert = await settledSlideCount(before + wanted);
  // THE FOURTH SITE, missed by the first sweep of this exact defect — the
  // "fix written, one call site left behind" shape, committed the same day the
  // other three were repaired. `deckGrowth` returns null when either count was
  // never taken, and `null < wanted` is FALSE in JS, so an unread deck sailed
  // past this guard and into the arithmetic below. Round 360 filed the result:
  // "4 slides inserted, NaN kept".
  const landedCount = deckGrowth(before, afterInsert);
  if (landedCount === null) return countlessSkip();
  if (landedCount < wanted)
    // A SKIP, not a failure. This scenario asks whether the reconcile
    // deduplicates two slides claiming one slot; if the slides are not there to
    // be deduplicated then it has not been asked, and "failed" blames the dedup
    // for an insert that did not land. The same reasoning as `blindSkip`,
    // reached by the other door — there the host would not SAY, here it did not
    // DO.
    return {
      ok: false,
      skipped: true,
      detail:
        `only ${landedCount} of ${wanted} slides landed even after a settle ` +
        `(the two inserts reported ${landedA} and ${landedB}), so there was nothing to deduplicate`,
    };
  const outcome = await reconcileDeck(
    // The count the GENERATOR reported, not an estimate of what Office.js
    // would have drawn: the two disagree by design (a pie is one custGeom
    // wedge in a file and a sixteen-triangle fan through Office.js), and
    // estimating here measured five perfect charts as wreckage on a real run.
    titles.map((t, i) => ({ slot: i, title: t, shapes: shapesPerSlide[i], chart: true, wroteTag: true })),
    { before, after: afterInsert },
    () => undefined,
    { run },
  );
  const settled = await slideCount();
  // A blind READBACK is not a finding. Every scenario here guards its first
  // scan and then draws its loudest conclusion from a second one it never
  // checked — and on the web a short deck scan is routine, which is why
  // `DeckScan.short` exists at all. `insertOntoUsedSlide` was fixed for exactly
  // this and its five siblings were never swept with it, so a host that answered
  // one page short produced a hard FAILED verdict asserting data loss the
  // add-in did not cause. Those sentences are what send a maintainer after the
  // tag-write path; this repo has already spent rounds on that hunt.
  const { found: kept, blind, gap } = await probeCharts(`${prefix} dup`);
  if (blind) return blindSkip(gap);
  // Four slides in, two out, and both survivors still charts. A pass that
  // deleted both copies of one item would leave two slides and read as
  // success on the count alone — hence the second half of this.
  // The same guard again on the closing count: this is the line that produced
  // round 360's "4 slides inserted, NaN kept", and a verdict is exactly as
  // unfounded on an unread count here as it was on the opening one.
  const keptCount = deckGrowth(before, settled);
  if (keptCount === null) return countlessSkip();
  const ok = keptCount === 2 && kept.length === 2;
  return {
    ok,
    detail: `4 slides inserted, ${keptCount} kept, ${kept.length} of 2 still re-editable; ${outcome.plan.summary.duplicates} queued as duplicates`,
  };
};

/**
 * Of the probe charts on offer, the one on the slide this run has loaded least.
 *
 * Every scenario that needs "a chart to work with" took `found[0]`, which is
 * the same chart on the same slide every time — so the battery piled its whole
 * run onto one slide and then paid the per-slide cost curve for it. Round
 * `275a76a` measured that slide going 20 → 68 → 92 → 116 → 140 → 144 → 165
 * shapes while nothing else in the deck passed 34, and the draw at 144 stalled:
 * a NINE-shape chart, timed out at 45 seconds, which ~63 seconds of per-slide
 * overhead at the measured +0.44s per shape entirely accounts for.
 *
 * So a scenario ordered late was measurably harder to pass than the same
 * scenario ordered early, and several rounds of "this host stalls
 * intermittently" have some of that in them.
 *
 * STABLE, not merely least: ties break on the deck's own order, so a run where
 * nothing has been drawn yet picks exactly what `found[0]` picked and every
 * existing expectation about which chart a scenario takes still holds. It only
 * diverges once this run has actually loaded a slide, which is the case it
 * exists for.
 *
 * Pure, and takes the count as a function, so the rule can be checked without a
 * PowerPoint — the same reason `placeChart` and `chooseGroupMembers` are their
 * own functions.
 */
export function leastLoadedChart<T extends { target: { slideId: string } }>(
  charts: T[],
  drawnOn: (slideId: string) => number,
): T | undefined {
  let best: T | undefined;
  let bestLoad = Infinity;
  for (const c of charts) {
    const load = drawnOn(c.target.slideId);
    // Strictly less: the first chart at a given load wins, so ties keep the
    // deck's order.
    if (load < bestLoad) {
      best = c;
      bestLoad = load;
    }
  }
  return best;
}

/**
 * Redrawing a chart the user is looking at.
 *
 * The add-in's worst case: every shape replaced, on the one slide guaranteed
 * to be on screen. A real run died on its first batch here. The scenario puts
 * the view on the slide FIRST — redrawing off-screen would be testing the easy
 * path, which is the one that already works.
 */
const editOnVisibleSlide: Scenario = async (prefix) => {
  const { found, blind, gap } = await probeCharts(prefix);
  const chart = leastLoadedChart(found, shapesDrawnOn);
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to edit" };
  const shown = await showSlide(chart.target.slideId);
  const next = { ...chart.cfg, title: `${chart.cfg.title} (edited)` };
  const target = await updateChartInSlide(buildChart(next), chart.target, { tagData: JSON.stringify(next) });
  if (!target) return { ok: false, detail: "the chart was gone from the slide after the update" };
  // A blind READBACK is not a finding. Every scenario here guards its first
  // scan and then draws its loudest conclusion from a second one it never
  // checked — and on the web a short deck scan is routine, which is why
  // `DeckScan.short` exists at all. `insertOntoUsedSlide` was fixed for exactly
  // this and its five siblings were never swept with it, so a host that answered
  // one page short produced a hard FAILED verdict asserting data loss the
  // add-in did not cause. Those sentences are what send a maintainer after the
  // tag-write path; this repo has already spent rounds on that hunt.
  const { found: again, blind: againBlind, gap: againGap } = await probeCharts(prefix);
  if (againBlind) return blindSkip(againGap);
  const round = again.find((c) => c.cfg.title === next.title);
  return {
    ok: !!round,
    detail: round
      ? `redrawn ${shown ? "on the visible slide" : "(host would not move the view)"} and the config round-tripped`
      : "the edited chart is no longer re-editable — its config did not survive the redraw",
  };
};

/**
 * A deck-wide rescale, with a chart on screen.
 *
 * Every chart's shapes are deleted and redrawn in one request context. A chart
 * whose redraw stalls is left blank, so what matters is not only that it
 * finishes but that nothing is emptied on the way.
 */
/**
 * Which probe chart the lone-run arm should use — the LAST one that does not
 * share the first chart's slide, or nothing.
 *
 * Pure, and separate from the scenario, because the choice IS the experiment.
 * Rounds 213/214 ran this arm on `charts[0]` — the same chart the deck-wide
 * rescale updates first — so it landed on slide 257, which held 42 shapes
 * against 20-21 on every other slide in the deck. At 18 of 24 every expensive
 * sample in those rounds sat on that slide and every cheap one sat elsewhere, so
 * "alone costs what first costs" was equally well read as "both are on the
 * crowded slide". The arm reproduced the confound it existed to break, and a
 * number from a confounded arm is worse than no number.
 *
 * Returns null rather than a fallback: an arm that cannot separate the two
 * variables must SKIP and say why, not quietly measure the thing it was built
 * to rule out.
 */
export function pickLoneChart(slideIds: readonly string[], wantLoaded = false): number | null {
  if (!slideIds.length) return null;
  // THE OTHER ARM. Position is now demonstrated with load held constant (round
  // 239: charts 1, 2 and 3 on one slide holding three shapes, chart 1 +21800
  // above the fit and the others +3700). Load's INDEPENDENT effect needs the
  // mirror image — run-length held at one, the slide varied — and that means a
  // lone chart ON the busy slide, which is exactly what the default refuses.
  //
  // Index 0 deliberately: it is the 18-of-24 chart, so it is size-matched with
  // the clear-slide arm. Its slide-mates are 9-of-16 and would compare nothing.
  //
  // Null rather than a fallback when nothing shares that slide — a deck that
  // cannot pose the question must say so, the same rule the default arm follows.
  if (wantLoaded) return slideIds.some((s, i) => i > 0 && s === slideIds[0]) ? 0 : null;
  for (let i = slideIds.length - 1; i > 0; i--) if (slideIds[i] !== slideIds[0]) return i;
  return null;
}

/**
 * Which arm the lone-chart scenario runs — clear slide by default.
 *
 *     localStorage.setItem("ssf-charts-lone-chart-loaded", "1")
 *
 * Off unless asked, read per run, and traced either way so a round always says
 * which arm produced its number. Same discipline as `scanSettleMs`: an untraced
 * default is how an experiment contaminates its own control.
 */
export const LONE_LOADED_KEY = "ssf-charts-lone-chart-loaded";

export function wantsLoadedLoneChart(read: (key: string) => string | null = readLoneLoaded): boolean {
  try {
    return read(LONE_LOADED_KEY) === "1";
  } catch {
    // A pane whose storage throws is not a pane that asked for the other arm.
    return false;
  }
}

function readLoneLoaded(key: string): string | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage.getItem(key);
  } catch {
    return null;
  }
}

/**
 * One chart, updated alone, in its own run, on a deck that is already warm.
 *
 * THE THIRD ARM for the first-chart cost. The first chart of a multi-chart
 * update costs ~2.2x a later chart of the same size — n=14 against n=70, and the
 * distributions DO NOT OVERLAP. Four explanations are excluded by measurement:
 * not the slide, not a cold host, not our sync count, and not the deck scan,
 * which rounds 211/212 killed with a 10s pause that moved chart 1 by nothing.
 *
 * What survives is that the cost attaches to a run's FIRST chart, and that has
 * two readings this scenario separates:
 *
 *   pays ~37s  → the run is the unit. The cost is per-run set-up that chart 1
 *                absorbs, and it is paid even when nothing follows it.
 *   pays ~17s  → "first" was standing in for "has work queued behind it", and
 *                the cost belongs to the queue rather than to the position.
 *
 * RUNS AFTER THE DECK-WIDE RESCALE, deliberately: that leaves the deck, the
 * host and the context as warm as this harness ever gets them, so a slow reading
 * here cannot be blamed on start-up.
 *
 * A DIFFERENT HEADROOM from the deck rescale, and that is load-bearing. The
 * rescale leaves every chart at `max = tallest * 1.5`; re-applying that same
 * ceiling changes NOTHING, and an update that changes nothing measures nothing —
 * it would report a fast run of zero work as evidence about a run of eighteen
 * shapes. 2.25 moves the same bars again.
 *
 * The measurement is `updated only the shapes that changed`, labelled `1/1` so
 * the archive can tell a run of one from the first of eight. `poolUpdateCost`
 * was keying `first` off `/^1\//` and would have pooled this row INTO the
 * first-chart arm it exists to be contrasted with; it reads the run length now.
 */
const oneChartAlone: Scenario = async (prefix) => {
  const { found: charts, blind, gap } = await probeCharts(prefix);
  if (!charts.length)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to update on its own" };
  // NOT charts[0], AND THAT IS THE WHOLE POINT OF THIS SCENARIO.
  //
  // It picked charts[0] first — the same chart the deck-wide rescale updates
  // FIRST — and so landed on the same slide, which in rounds 213/214 was slide
  // 257 holding 42 shapes against 20-21 on every other slide. Every expensive
  // 18-of-24 sample in those rounds sat on that slide and every cheap one sat
  // elsewhere, so "alone costs what first costs" was equally well read as "both
  // are on the crowded slide". The arm reproduced the confound it was built to
  // break, and this file measures slide occupancy as a first-order cost — a
  // batch on a slide holding 21-50 shapes costs 14044ms against 5486ms at 1-20.
  //
  // The LAST probe chart sits on its own slide, uncrowded, so run-position and
  // slide-occupancy finally vary independently:
  //
  //   pays ~35s on an uncrowded slide  → position is the cause, slide excluded.
  //   pays ~15s on an uncrowded slide  → the SLIDE was the cause all along, and
  //                                      "first chart of a run" was standing in
  //                                      for "the chart on the busiest slide".
  //
  // Skipped rather than guessed when the deck will not provide the contrast:
  // a lone chart that shares the first chart's slide measures nothing this
  // scenario claims to measure.
  const loaded = wantsLoadedLoneChart();
  const pick = pickLoneChart(
    charts.map((x) => x.target.slideId),
    loaded,
  );
  // TRACED EITHER WAY, including the default, so a round always says which arm
  // produced its number rather than leaving it to be inferred from the slide id.
  trace("selftest", "the lone chart's arm", { arm: loaded ? "loaded slide" : "clear slide", pick });
  if (pick === null)
    return {
      ok: false,
      skipped: true,
      detail: loaded
        ? "no probe chart shares the first chart's slide — the loaded arm has nothing to measure"
        : "every probe chart shares the first chart's slide — the slide/position confound is not broken here",
    };
  const c = charts[pick];
  await showSlide(c.target.slideId);
  const values = (c.cfg.data?.series ?? [])
    .flatMap((s) => s.values)
    .map(Number)
    .filter((v) => Number.isFinite(v));
  if (!values.length) return { ok: false, skipped: true, detail: "the probe chart carries no numbers to scale" };
  const ALONE_HEADROOM = 2.25;
  const max = Math.round(Math.max(...values) * ALONE_HEADROOM);
  // The same guard the deck rescale carries, for the same reason: a ceiling that
  // redraws the identical scene makes this a timing of nothing.
  const wasInk = barInk(buildChart(c.cfg));
  const nowInk = barInk(buildChart({ ...c.cfg, scale: { max } }));
  if (wasInk > 0 && nowInk === wasInk)
    return { ok: false, detail: `scale.max=${max} redraws the chart identically — the run would time no work` };
  const next: ChartConfig = { ...c.cfg, scale: { max } };
  const back = await traceAbout({ chart: "1/1" }, () =>
    updateChartInSlide(buildChart(next), c.target, { tagData: JSON.stringify(next) }),
  );
  if (back === null) return { ok: false, detail: "the chart was gone before it could be updated alone" };
  // The TIMING is the deliverable and it is already traced by the update itself.
  // This verdict only has to say the run was real work that survived — a lost
  // config here would mean the timing described a redraw, not an update.
  return {
    ok: !back.lost,
    detail: back.lost
      ? `updated alone but lost its config (${back.lost}) — the timing describes a redraw, not an update`
      : `updated alone in its own run on a warm deck, config intact — see \`chart 1/1\` for the cost`,
  };
};

const sameScaleAcrossDeck: Scenario = async (prefix) => {
  const { found: charts, blind, gap } = await probeCharts(prefix);
  if (charts.length < 2)
    return blind
      ? blindSkip(gap)
      : { ok: false, skipped: true, detail: "fewer than two probe charts to scale together" };
  await showSlide(charts[0].target.slideId);
  // Every finite value across every probe chart. `Math.max()` of nothing is
  // -Infinity, and JSON.stringify turns that into `null` — so an empty or
  // all-blank data set wrote `{"scale":{"max":null}}` into the tag and then
  // compared it against -Infinity, which never matches. The scenario reported
  // a failure that was its own arithmetic.
  const values = charts
    .flatMap((c) => c.cfg.data?.series ?? [])
    .flatMap((s) => s.values)
    .map(Number)
    .filter((v) => Number.isFinite(v));
  if (!values.length) return { ok: false, skipped: true, detail: "probe charts carry no numbers to scale" };
  // ABOVE the tallest bar in the deck, not level with it.
  //
  // `Math.max(...values)` is the ceiling these charts are ALREADY drawn at:
  // every probe chart is the same `sampleConfig("clustered")`, whose axis tops
  // out at its own largest value. Writing that back as `scale.max` therefore
  // produced a config that rendered to the byte-identical scene, and a deck
  // saved out of a real PowerPoint proved it — a rescaled chart and an
  // untouched copy of it had the same bar geometry to the EMU. The scenario is
  // named for a deck-wide rescale and was measuring a deck-wide RE-TAG: it
  // could only ever catch a config that failed to round-trip, never a host
  // that took the config and redrew the old picture.
  //
  // Headroom fixes that at the source. Every bar has to shrink to fit a
  // ceiling nothing in the data reaches.
  const HEADROOM = 1.5;
  const max = Math.round(Math.max(...values) * HEADROOM);
  // And prove the ceiling bites, rather than trusting the arithmetic above to
  // stay true. The no-op survived for as long as it did because nothing ever
  // checked that the new config DREW anything different; a future sample whose
  // data reaches this ceiling would put the scenario straight back to
  // measuring nothing, silently, and the next person would find out from a
  // saved deck the way this one was found.
  const wasInk = barInk(buildChart(charts[0].cfg));
  const nowInk = barInk(buildChart({ ...charts[0].cfg, scale: { max } }));
  if (wasInk > 0 && nowInk === wasInk)
    return {
      ok: false,
      detail: `scale.max=${max} redraws the chart identically — the rescale would prove nothing about the host`,
    };
  // What the UPDATE said, kept rather than dropped on the floor.
  //
  // `updateChartInSlide` already reports a chart it redrew and could not tag
  // (`lost: "no-config"`) and one whose new shape it never got an id for
  // (`"unknown-shape"`). This scenario threw the return value away and then
  // re-counted the deck, so its verdict — "4 of 8 charts carry the shared
  // scale" — named a number and no cause, and the deck it produced took a
  // session to read afterwards. The renderer knew which four and why while it
  // was happening.
  const lost: Record<string, number> = {};
  const outcomes: (string | undefined)[] = [];
  // WHY, not just how many. This scenario has reported "3 of 8 charts carry the
  // shared scale … the host flipped at chart 4 of 8" in all 34 rounds it has run,
  // and the reason each chart was lost existed only in the trace — so reading a
  // round meant opening the log and joining it by hand, every time.
  //
  // The mechanism is known now (`docs/BACKLOG.md`): the pre-grouping re-read
  // comes back short or empty on a slide this run just added, so the chart is
  // not grouped, so its tag falls back to a handle this host refuses. The deltas
  // below put that on the verdict line itself.
  // WHAT EACH SLIDE HOLDS, before a single chart is touched.
  //
  // The position-vs-load question survived four corrections on 2026-08-25
  // because nothing recorded this. The first chart of this run is always the
  // chart on the deck's busiest slide — by construction, since earlier scenarios
  // draw onto the visible one — so run-position and slide-load have been
  // confounded across every sample in the archive. The only available proxy
  // reports the same slide as 10 and as 42.
  //
  // One host call, HERE, before the loop: it costs nothing inside the timed
  // per-chart path, which is the whole reason it is taken in the scenario rather
  // than in `updateChartsInSlides`.
  const occupancy = await slideOccupancy(charts.map((c) => c.target.slideId));
  trace("selftest", "what each slide held before the rescale", {
    charts: charts.length,
    // Ordered as the charts are, so `slides[0]` is the first chart's slide — the
    // one the whole question is about.
    slides: charts.map((c) => ({ slide: c.target.slideId, shapes: occupancy.get(c.target.slideId) ?? null })),
  });
  const frictionBeforeRescale = hostFrictionCounts();
  // THE SCAN/FIRST-CHART CONFOUND — see `scanSettleMs`. Between the deck scan
  // above and the first update below, which is the only gap where a pause can
  // tell the two apart. Traced at zero as well, so every archived round says
  // which arm it ran in.
  const settleMs = scanSettleMs();
  trace("selftest", "settling before the first chart of the run", { settleMs, charts: charts.length });
  if (settleMs > 0) await new Promise((resolve) => setTimeout(resolve, settleMs));
  for (const [n, c] of charts.entries()) {
    const next: ChartConfig = { ...c.cfg, scale: { max } };
    // Which chart of how many, on every line these calls write — the draw
    // batches, the grouping, the tag write, the settle and every error. Within
    // one chart's own update they all say `index: 0`, so the per-chart table
    // this scenario's result is made of used to be recovered by pairing trace
    // lines against timestamps by hand.
    const back = await traceAbout({ chart: `${n + 1}/${charts.length}` }, () =>
      updateChartInSlide(buildChart(next), c.target, { tagData: JSON.stringify(next) }),
    );
    const why = back === null ? "chart-gone" : back.lost;
    if (why) lost[why] = (lost[why] ?? 0) + 1;
    outcomes.push(why ?? undefined);
    if (rescaleShouldStop(outcomes)) {
      trace("selftest", "stopping the rescale — the host has already flipped", {
        done: n + 1,
        of: charts.length,
        flippedAt: rescaleFlipIndex(outcomes),
      });
      break;
    }
  }
  // A blind READBACK is not a finding. Every scenario here guards its first
  // scan and then draws its loudest conclusion from a second one it never
  // checked — and on the web a short deck scan is routine, which is why
  // `DeckScan.short` exists at all. `insertOntoUsedSlide` was fixed for exactly
  // this and its five siblings were never swept with it, so a host that answered
  // one page short produced a hard FAILED verdict asserting data loss the
  // add-in did not cause. Those sentences are what send a maintainer after the
  // tag-write path; this repo has already spent rounds on that hunt.
  const { found: after, blind: afterBlind, gap: afterGap } = await probeCharts(prefix);
  if (afterBlind) return blindSkip(afterGap);
  const scaled = after.filter((c) => c.cfg.scale?.max === max).length;
  const shrunk = wasInk > 0 ? `, bars redraw at ${Math.round((nowInk / wasInk) * 100)}% of their height` : "";
  const why = Object.entries(lost)
    .map(([k, n]) => `${n}×${k}`)
    .join(", ");
  return {
    ok: scaled === charts.length && after.length === charts.length,
    detail:
      `${scaled} of ${charts.length} charts carry the shared scale (max=${max}${shrunk}); ` +
      `${after.length} still re-editable${why ? `; the update reported ${why}` : ""}` +
      rescaleLossNote(outcomes, charts.length) +
      reReadNote(frictionBeforeRescale, hostFrictionCounts()),
  };
};

/**
 * The clause that makes a round read itself: WHY the charts were lost, and
 * whether the settled retry saved any.
 *
 * The verdict has named a count and a flip index since this scenario existed,
 * and never a cause — so every round was read by opening the trace and joining
 * it back to the summary by hand. The mechanism is settled now, so the summary
 * can carry it.
 *
 * Silent when nothing happened, which is the point: it must add words only to a
 * round that has something to say. And it reports the RETRY's own number
 * (`reReadsRepaired`) whether or not the scenario passed, because a round where
 * the retry repaired charts and the scenario still failed and a round where it
 * repaired none are two very different results that would otherwise read the
 * same.
 */
export function reReadNote(before: ReReadCounts, after: ReReadCounts): string {
  const empty = after.emptyReReads - before.emptyReReads;
  const short = after.shortReReads - before.shortReReads;
  const saved = after.reReadsRepaired - before.reReadsRepaired;
  if (!empty && !short && !saved) return "";
  const why = [
    short ? `${short} read short` : "",
    empty ? `${empty} read empty` : "",
    // Named even at zero once the other two are non-zero: "the pause repaired
    // none" is the finding that refutes the settling-slide theory, and a clause
    // that simply vanishes would leave the reader unable to tell it from a round
    // where the retry never ran.
    `the settled retry repaired ${saved}`,
  ].filter(Boolean);
  return `; of the re-reads before grouping, ${why.join(", ")}`;
}

/** Just the fields `reReadNote` differences — so a caller cannot pass halves of two snapshots. */
type ReReadCounts = Pick<ReturnType<typeof hostFrictionCounts>, "emptyReReads" | "shortReReads" | "reReadsRepaired">;

/**
 * Consecutive charts that must lose their config before the rescale gives up.
 *
 * Two, not one. A single loss is the flip's first chart and may still be
 * rescued by the settle — round 16's chart 4 was, and it counted toward the
 * score — so stopping on one would throw away the observation that decides
 * whether the round scores 3 or 4.
 */
const RESCALE_LOSSES_BEFORE_STOPPING = 2;

/**
 * Whether the deck-wide rescale has learned everything this host will teach it.
 *
 * The scenario is 212 seconds of an 818-second round, the longest thing in the
 * battery, and on a degrading host its tail is pure repetition: once two charts
 * in a row have lost their config, every remaining chart in three separate
 * rounds has done exactly the same thing, at ~12s each. Round 16 would have
 * stopped after chart 5 and skipped three charts and ~38 seconds, for the same
 * score and the same flip index.
 *
 * On a HEALTHY host this never fires — nothing is lost, so nothing is skipped
 * and the scenario still proves every chart in the deck takes the shared scale.
 * That is the property that makes the shortcut safe: it can only skip work the
 * host has already refused twice.
 *
 * Pure, and separate from the scenario, because the decision is the part worth
 * checking — same reason as `positionalSweepPlan` and `chooseGroupMembers`.
 */
export function rescaleShouldStop(outcomes: (string | undefined)[]): boolean {
  if (outcomes.length < RESCALE_LOSSES_BEFORE_STOPPING) return false;
  return outcomes.slice(-RESCALE_LOSSES_BEFORE_STOPPING).every(Boolean);
}

/** 1-based index of the first chart that lost its config, or null if none did. */
export function rescaleFlipIndex(outcomes: (string | undefined)[]): number | null {
  const i = outcomes.findIndex(Boolean);
  return i < 0 ? null : i + 1;
}

/**
 * What the verdict says about the losses — and specifically, when it may use
 * the word FLIPPED.
 *
 * Caught on this sentence's first outing. The round of `4feb5be` lost exactly
 * ONE chart of eight and scored 7 of 8, the best this scenario has ever
 * recorded, and the verdict announced `the host flipped at chart 5 of 8`. That
 * is the same word three earlier rounds use for a host that degraded and never
 * came back, applied to a host that dropped one chart and carried on — and
 * a reader comparing rounds would have read a regime change into the best round
 * on file.
 *
 * A flip is the thing the scenario stops for: two consecutive losses, the same
 * evidence `rescaleShouldStop` acts on. One loss is one loss, and the index is
 * still worth reporting — it is where the settle's coin landed — but it must
 * not be dressed as a boundary.
 */
export function rescaleLossNote(outcomes: (string | undefined)[], total: number): string {
  const first = rescaleFlipIndex(outcomes);
  if (first === null) return "";
  const losses = outcomes.filter(Boolean).length;
  if (!rescaleShouldStop(outcomes))
    return `; ${losses} of the ${outcomes.length} charts redrawn lost its config, the first at chart ${first} of ${total}`;
  const skipped = total - outcomes.length;
  return (
    `; the host flipped at chart ${first} of ${total}` +
    (skipped > 0 ? `, so the last ${skipped} were not attempted` : "")
  );
}

/**
 * Total height of a scene's bars — the one number a rescale has to move.
 *
 * Summed rather than taken off the tallest bar, because a chart that honours a
 * new ceiling for one series and ignores it for another is still drawn wrong,
 * and a single-bar reading would call that fixed.
 */
function barInk(scene: Scene): number {
  let total = 0;
  for (const n of scene.nodes) if (n.kind === "rect" && n.name?.startsWith("seg-")) total += n.h;
  return total;
}

/**
 * Turning a degraded picture back into native shapes.
 *
 * The picture path is what the web host falls back to, and Explode is the only
 * way out of it. The claim it rests on — that a picture keeps its config and
 * can become shapes again — has never been checked anywhere but a fake.
 */
/**
 * What a null return from an update MEANS, in words a verdict may use.
 *
 * `updateChartInSlide` answers null for two facts that are nothing alike: the
 * chart really is gone, or this host refused to name it. This scenario called
 * both "vanished" — a claim that the add-in destroyed the user's content — and
 * round `ee1741e` is the case that shows the cost. It reported `the picture
 * vanished while being exploded back to shapes`, and the deck inventory taken
 * at the end of the same run shows that slide holding one shape named
 * `PowerChart`. Nothing vanished. The update had just logged three
 * `InvalidParam passed to GetItem(id)` and an empty settle re-read, so the run
 * knew; only the verdict did not.
 *
 * Same defect as the empty-read one a few lines above, one step later in the
 * same scenario, and it wants the same treatment: decide from something outside
 * the answer. The friction counters are that something — they are kept for the
 * whole run anyway, so this costs a subtraction and no host call.
 *
 * Deliberately not a pass. Both outcomes are failures of the scenario; what
 * changes is which one a maintainer goes and investigates, and "the add-in
 * deleted a chart" and "the host would not answer for it" send them to
 * opposite ends of the codebase.
 */
/**
 * Is that shape still on its slide? Three answers, and the third is the point.
 *
 * `true` and `false` are what they look like. `undefined` is "the slide would
 * not say", which must never be reported as destruction — this host refuses
 * collection reads routinely, and an empty read of a slide the run has drawn on
 * is the refusal rather than an empty slide (see the collapse readback above,
 * and `slideHoldsOnlyChart`).
 */
/**
 * Delete the slide this scenario added, found by what the deck gained.
 *
 * WHY NOT JUST DELETE IT BY ITS ID. The id `addSlideForChart` returns is the
 * add-time one — `4123571152#123571113` — and this host RE-KEYS a slide once it
 * settles, to `258#347339495` and the like. So minutes later the deck holds the
 * slide under a name nobody upstream has, `deleteSlideById` cannot resolve it,
 * `deleteSlideByPosition` finds `indexOf === -1`, and it correctly refuses to
 * remove an index whose id does not match. That refusal is load-bearing and
 * must stay: round 124 reported 62 scratch slides swept and deleted none.
 *
 * The probe layer already works around this by sweeping positionally, and its
 * own trace says why — "delete-by-id left slides behind and this host does not
 * list the ids they were added under".
 *
 * So the slide is identified by a SET DIFFERENCE taken at delete time rather
 * than by a name captured at add time. Whatever the deck calls it now, it is
 * the id that is in the listing and was not there before. That is immune to
 * re-keying by construction, and it hands `deleteSlideById` an id the deck is
 * currently listing — the case it handles well.
 *
 * NOT EXACTLY ONE means something else changed the deck underneath this
 * scenario, and then nothing here can attribute a slide to this add. It
 * refuses rather than guessing, for the same reason the positional delete does.
 *
 * This became possible only on 2026-09-03. Before the slide-master fix the
 * added slide was rejected by the server and rolled back, so there was no
 * slide to find and no diff to take.
 *
 * AND IT ASKS MORE THAN ONCE, because the first version was immune to the id
 * having CHANGED and not to it not having changed YET.
 *
 * Round 378 caught that. `stop a run mid-draw` aborts early by design, so its
 * cleanup runs seconds after the add — before the host re-keys — and the diff
 * handed back the ADD-TIME id, which is exactly the name `deleteSlideById`
 * cannot resolve. Rounds 376 and 377 passed the same scenario only because
 * their cleanup ran later, after the re-key. Same code, opposite result, and
 * the difference was timing.
 *
 * So it retries, briefly. Each pass re-reads the deck and re-takes the
 * difference, because what changes between them is the NAME: the slide is
 * there the whole time, under a form that resolves a moment later. Three
 * passes over about a second, which is the settle window the rest of this file
 * uses for the same host.
 *
 * THE RETRY IS A HYPOTHESIS, AND NO TEST HERE PROVES IT. Two reasons, both
 * measured rather than assumed. `faults.newSlideIdSettlesAfter` settles an id
 * after N LOOKUPS, and a single `deleteSlideById` attempt makes several on its
 * own — so the id always settles inside one attempt, because the host's settle
 * is a matter of TIME and this fake cannot express that. And the call site
 * reads `deleteSlideAddedSince(...) || deleteSlideById(slideId)`, so a failed
 * diff falls through to the old route and the verdict looks identical either
 * way. Cutting this loop to a single pass changes no test in the suite; that
 * was checked, not supposed.
 *
 * A ROUND ADJUDICATES IT. `stop a run mid-draw` failed this cleanup in rounds
 * 378, 382 and 383. If it starts passing at 4:3 the retry is doing the work;
 * if it does not, revert this rather than lengthen it. Do not read the loop as
 * verified because the suite is green.
 *
 * DELIBERATELY NOT A POSITIONAL DELETE, though the diff knows the index and it
 * would work first time. Deleting an index whose id has not been confirmed is
 * the move round 124 punished — 62 scratch slides reported swept and none
 * gone — and buying three retries is cheaper than reopening that. If every
 * pass fails the caller is told false and reports it, which is honest and
 * costs a scenario rather than someone's deck.
 */
async function deleteSlideAddedSince(before: string[] | undefined): Promise<boolean> {
  if (!before) return false;
  const known = new Set(before);
  for (let pass = 0; pass < 3; pass++) {
    if (pass) await new Promise((resolve) => setTimeout(resolve, 400));
    const now = await deckSlideIds().catch(() => undefined);
    if (!now) continue;
    const fresh = now.filter((id) => !known.has(id));
    // Not exactly one means something else moved the deck underneath this
    // scenario, and no retry makes that attributable. Give up rather than guess.
    if (fresh.length !== 1) return false;
    if (await deleteSlideById(fresh[0])) return true;
  }
  return false;
}

async function stillThere(target: { slideId: string; shapeId?: string }): Promise<boolean | undefined> {
  if (!target.shapeId) return undefined;
  const shapes = await slideShapeList(target.slideId);
  if (!shapes) return undefined;
  // An empty read of a slide this run drew on is not an empty slide. Same
  // reasoning, and the same evidence, as the collapse readback.
  if (shapes.length === 0 && shapesDrawnOn(target.slideId) > 0) return undefined;
  return shapes.some((s) => s.id === target.shapeId);
}

/**
 * Did `updateLossNote` conclude anything, or did it say the round proves nothing?
 *
 * The note has always had one branch that DISCLAIMS its own evidence — the host
 * refused ids during the scenario, so the slide failing to name the chart says
 * nothing either way. That is not a failure, it is an absence of measurement,
 * and reporting it as `ok: false` made the two indistinguishable.
 *
 * It cost the regression gate its credibility on its first live outing: round
 * 073 flagged `explode a degraded picture` as having stopped passing when the
 * scenario's own words were "proves nothing either way". A gate that fires on a
 * scenario declining to conclude is a gate that gets switched off, which this
 * repo has already watched happen once.
 *
 * Kept as a predicate over the same inputs rather than a flag threaded through
 * the note, so the two can never disagree about which branch was taken.
 */
export function lossNoteIsInconclusive(refusalsDuring: number, stillOnSlide?: boolean): boolean {
  return stillOnSlide === false && refusalsDuring > 0;
}

export function updateLossNote(what: string, refusalsDuring: number, stillOnSlide?: boolean): string {
  // AN ID REFUSAL ANYWHERE IN THIS SCENARIO MAKES EVERY ID IN IT SUSPECT, so a
  // missing shape stops being evidence of a missing shape.
  //
  // Round `1789749` is why, and it is the fourth mechanism to produce a false
  // destruction claim from this one scenario. The collapse's readback was
  // refused (`reading back where the charts landed`, InvalidParam), so the
  // target handed on carried an id the host never confirmed — the settle had to
  // find that chart BY NAME, `withId: 0`, which is the tell. The picture then
  // landed under an id nobody here holds, `some(s => s.id === target.shapeId)`
  // answered false, and the verdict read `the picture is GONE from the slide`
  // while the deck inventory from the same run shows that slide holding one
  // shape named `PowerChart` — the picture, untouched.
  //
  // So "gone" now needs a clean scenario as well as a positive read. The
  // refusal count is measured from the scenario's start rather than from just
  // before the call for exactly this reason: the refusal that poisoned the id
  // happened in the step BEFORE, and a per-call window could not see it.
  if (stillOnSlide === false && refusalsDuring > 0)
    return (
      `the update would not work on the ${what} and this host refused ${refusalsDuring} id(s) in this scenario, so ` +
      `the slide not naming it proves nothing either way`
    );
  // Positive evidence first: the slide was asked, and it answered.
  if (stillOnSlide === true)
    return `the host would not work on the ${what} again, but it is STILL ON THE SLIDE — nothing was lost`;
  if (stillOnSlide === false) return `the ${what} is GONE from the slide — it was destroyed while being redrawn`;
  // No answer from the slide. A refusal during the call still explains it…
  if (refusalsDuring > 0)
    return (
      `the host would not name the ${what} afterwards (${refusalsDuring} id refusal(s) during the call), so it ` +
      `could not be worked on again`
    );
  // …and with neither, the honest answer is that we do not know. It must NOT
  // read as destruction. Round `eaddbf4` is why: the explode's update returned
  // null having logged nothing at all and refused nothing inside that call, so
  // the old wording printed `the picture vanished while being redrawn` while
  // the deck inventory from the same run showed a chart on every slide. That is
  // the third false destruction claim this scenario has produced, and the first
  // that survived a fix aimed at the previous two — because the fix keyed on
  // thrown id refusals and this host can also fail to resolve a target quietly,
  // with no throw to count.
  return `the update would not work on the ${what} and the slide would not say what became of it`;
}

/**
 * The drag round trip, without a drag.
 *
 * `CHART_ORIGIN_TAG` records where a chart was first drawn so a later update can
 * shift the redraw by however far the user has since moved it. `docs/PUBLISHING.md`
 * says that round trip "needs a real drag and so cannot be scripted at all", and
 * so nothing has ever checked it — which stopped being acceptable when rounds
 * 070-072 moved every remaining failure onto exactly that tag: `writing the
 * chart's origin tag` 5010s three to seven times a round, while the config tag
 * stopped failing altogether.
 *
 * **A drag is only a shape whose `left`/`top` changed**, so this moves one
 * programmatically and then updates the chart. The document state is what a
 * mouse would have produced, so the update path cannot tell the difference.
 *
 * WHAT IT DOES AND DOES NOT PROVE. It proves the ARITHMETIC — that an update
 * redraws at the moved position rather than snapping back to where the chart was
 * first inserted, which is the visible symptom of a lost origin tag. It does not
 * prove a mouse drag, and the manual check that does is still owed; test 4 of the
 * standing run stays.
 *
 * No selection call anywhere in it. `setSelectedShapes` wedges this host's whole
 * selection subsystem, and risking that to simulate a drag would cost more than
 * the check is worth.
 */
const dragThenUpdate: Scenario = async (prefix) => {
  const { found, blind, gap } = await probeCharts(prefix);
  const chart = leastLoadedChart(found, shapesDrawnOn);
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to move" };
  const before = { left: chart.target.left, top: chart.target.top };
  if (typeof before.left !== "number" || typeof before.top !== "number")
    return {
      ok: false,
      skipped: true,
      detail: "the host would not say where the chart is, so a move cannot be judged",
    };
  // Far enough that a rounding difference cannot be mistaken for a move, and
  // small enough to stay on the slide.
  const DX = 60;
  const DY = 40;
  if (!(await moveShapeBy(chart.target.slideId, chart.target.shapeId, DX, DY)))
    return {
      ok: false,
      skipped: true,
      detail: "the host would not move the shape, so the drag round trip is untested",
    };

  // The move has to be CONFIRMED before the update, or a failed move and a
  // failed redraw are indistinguishable — the trap `withId: 0` taught this
  // project about routes that never ran.
  const { found: moved, blind: movedBlind, gap: movedGap } = await probeCharts(prefix);
  if (movedBlind) return blindSkip(movedGap);
  const atMoved = moved.find((c) => c.cfg.title === chart.cfg.title);
  if (!atMoved || typeof atMoved.target.left !== "number")
    return { ok: false, detail: "the chart could not be found after moving it" };
  const shifted = Math.round(atMoved.target.left - before.left);
  if (Math.abs(shifted - DX) > 2)
    return { ok: false, skipped: true, detail: `the move did not land (x moved ${shifted}pt, wanted ${DX}pt)` };

  const next = { ...atMoved.cfg, title: `${atMoved.cfg.title} (moved)` };
  const target = await updateChartInSlide(buildChart(next), atMoved.target, { tagData: JSON.stringify(next) });
  if (!target) return { ok: false, detail: "the chart was gone from the slide after the update" };
  const { found: after, blind: afterBlind, gap: afterGap } = await probeCharts(prefix);
  if (afterBlind) return blindSkip(afterGap);
  const redrawn = after.find((c) => c.cfg.title === next.title);
  if (!redrawn || typeof redrawn.target.left !== "number")
    return { ok: false, detail: "the moved chart is no longer re-editable — its config did not survive the redraw" };

  // THE ASSERTION. A chart that kept its origin redraws where the user left it;
  // one that lost it snaps back to where it was first inserted, which is the
  // symptom a user would report as "it jumped".
  const drift = Math.round(redrawn.target.left - atMoved.target.left);
  const snappedBack = Math.abs(redrawn.target.left - before.left) < Math.abs(drift);
  return {
    ok: Math.abs(drift) <= 2,
    detail: snappedBack
      ? `the update SNAPPED THE CHART BACK to where it was first inserted (${Math.round(before.left)}pt), losing the ${DX}pt move — the origin tag did not survive`
      : Math.abs(drift) <= 2
        ? `moved ${DX}x${DY}pt and the update redrew it there, so the origin round trip held`
        : `the update redrew it ${drift}pt from where it was moved to`,
  };
};

const explodePicture: Scenario = async (prefix) => {
  if (!canInsertPicture()) return { ok: false, skipped: true, detail: "host has no picture fill (PowerPointApi 1.8)" };
  if (!rasterizer) return { ok: false, skipped: true, detail: "no rasteriser — cannot make a picture to explode" };
  const { found, blind, gap } = await probeCharts(prefix);
  const chart = leastLoadedChart(found, shapesDrawnOn);
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to explode" };
  // Collapse it to a picture first, the way a struggling host would, then ask
  // for it back as shapes. Both halves matter: a config lost on the way IN is
  // indistinguishable from one lost on the way OUT if only one is exercised.
  //
  // The picture comes from a real rasterisation. `render: "image"` on the
  // config does NOT make one — the renderer takes the picture path only when
  // handed `pictureBase64` — so an earlier version of this passed `undefined`
  // and quietly performed two ordinary shape updates while reporting that the
  // picture round-trip worked. A scenario that cannot fail is worse than one
  // that is missing: it reads as evidence.
  const asPicture: ChartConfig = { ...chart.cfg, render: "image" };
  const png = await rasterizer(buildChart(asPicture));
  if (!png) return { ok: false, skipped: true, detail: "the browser would not rasterise the chart" };
  // Read the slide BEFORE the collapse, so what the collapse produced can be
  // told apart from what was already there. See below for what that cost.
  const before = await slideShapeList(chart.target.slideId);
  // From the SCENARIO's start. See `updateLossNote`: the refusal that makes an
  // id untrustworthy is routinely in an earlier step than the call that fails.
  const refusalsAtStart = hostFrictionCounts().idRefusals;
  const pictured = await updateChartInSlide(buildChart(asPicture), chart.target, {
    tagData: JSON.stringify(asPicture),
    pictureBase64: png,
  });
  if (!pictured)
    return {
      ok: false,
      detail: updateLossNote(
        "chart",
        hostFrictionCounts().idRefusals - refusalsAtStart,
        await stillThere(chart.target),
      ),
    };
  // One shape is what a picture IS. More than one means the renderer drew
  // shapes instead and the rest of this scenario would prove nothing — the same
  // trap as the `undefined` png above, one step later, so it has to be a
  // verdict and not a note.
  //
  // Ask what the collapse ADDED, by shape id — never how many shapes the slide
  // holds. Two earlier versions asked the slide and both were wrong:
  //
  //  - `slideHoldsOnlyChart` is the slide-SWAP gate, which answers no for a
  //    slide it could not read, so a fine picture was logged as suspect
  //    whenever this host refused a collection (which is routinely);
  //  - counting the slide's shapes replaced that with a false FAILURE on its
  //    first real round. The slide held three, and all three were charts: the
  //    battery deliberately piles them onto one slide two scenarios earlier,
  //    and `a selected shape survives an insert` reports the same three. A
  //    neighbouring chart is not a broken picture.
  //
  // The id delta has neither problem. It is one shape when a picture lands
  // beside anything at all, and it is N when the renderer falls through to
  // native shapes.
  const afterRead = await slideShapeList(pictured.slideId);
  // An EMPTY read of a slide this run has just drawn on is the host refusing,
  // not the slide being empty — and the difference is a verdict claiming data
  // loss that did not happen.
  //
  // Round `957aca0` is the case. This scenario reported `the collapse added 0
  // shapes (none) — the slide went from 1 to 0`, and the deck inventory taken
  // at the end of the same run shows that slide holding one shape named
  // `PowerChart`. The chart never moved. `slideShapeList` is careful — it
  // corroborates `items` against `getCount()` and answers null when they
  // disagree — but here BOTH said zero, so nothing downstream could tell.
  // That is a hole in the corroboration rather than in this scenario: the two
  // signals can agree at zero and both be wrong.
  //
  // The floor argument is what makes this safe. `pictured` is non-null, so the
  // update handed back a target naming a shape it believes exists; a slide
  // cannot hold nothing while an operation that just reported success says
  // otherwise. Routed into the SAME path the scenario already has for "the
  // host would not say", which continues to the config round-trip below rather
  // than claiming anything about a picture.
  const after = afterRead && afterRead.length === 0 && before && before.length > 0 ? null : afterRead;
  if (afterRead && after === null)
    trace("selftest", "the slide read back EMPTY right after the collapse — not believing it", {
      slide: pictured.slideId,
      was: before?.length,
    });
  const delta =
    before && after
      ? {
          added: after.filter((s) => !new Set(before.map((b) => b.id)).has(s.id)),
          was: before.length,
          now: after.length,
        }
      : undefined;
  if (delta && delta.added.length !== 1)
    return {
      ok: false,
      // What was observed, not what caused it: a fall-through to native shapes
      // and a picture landing next to undeleted leftovers both show up here,
      // and the names are what a reader needs to tell them apart.
      detail:
        `the collapse added ${delta.added.length} shapes (${delta.added.map((s) => s.name).join(", ") || "none"}) — ` +
        `a picture is one; the slide went from ${delta.was} to ${delta.now}`,
    };
  // …and when the host will not say, the scenario still runs: the config
  // round-trip below is worth checking either way. It just may not have been a
  // picture making the trip, and the verdict says so rather than claiming one.
  const confirmed = delta !== undefined;
  if (!confirmed) trace("selftest", "the host would not confirm the picture is one shape", { slide: pictured.slideId });
  const asShapes: ChartConfig = { ...chart.cfg, render: "shapes" };
  const exploded = await updateChartInSlide(buildChart(asShapes), pictured, { tagData: JSON.stringify(asShapes) });
  if (!exploded) {
    const refusals = hostFrictionCounts().idRefusals - refusalsAtStart;
    const onSlide = await stillThere(pictured);
    return {
      ok: false,
      // SKIPPED, not failed, when the note disclaims its own evidence. See
      // `lossNoteIsInconclusive` — the host refused ids mid-scenario, so the
      // slide not naming the picture proves nothing, and a round that measured
      // nothing must not read as a round that measured a loss.
      skipped: lossNoteIsInconclusive(refusals, onSlide),
      detail: updateLossNote("picture", refusals, onSlide),
    };
  }
  // A blind READBACK is not a finding. Every scenario here guards its first
  // scan and then draws its loudest conclusion from a second one it never
  // checked — and on the web a short deck scan is routine, which is why
  // `DeckScan.short` exists at all. `insertOntoUsedSlide` was fixed for exactly
  // this and its five siblings were never swept with it, so a host that answered
  // one page short produced a hard FAILED verdict asserting data loss the
  // add-in did not cause. Those sentences are what send a maintainer after the
  // tag-write path; this repo has already spent rounds on that hunt.
  const rescan = await probeCharts(prefix);
  if (rescan.blind) return blindSkip(rescan.gap);
  const back = rescan.found.find((c) => c.target.shapeId === exploded.shapeId);
  return {
    ok: !!back,
    detail: back
      ? confirmed
        ? "collapsed to a picture and exploded back, config intact"
        : "collapsed and exploded back, config intact — but the host would not confirm it was a picture"
      : "exploded, but the config did not survive",
  };
};

/**
 * Drawing charts onto a slide that already has content on it.
 *
 * The everyday action — "insert a chart on the slide I am looking at" — and
 * until now the only path with no real-host coverage at all. Every other
 * scenario here, and the whole demo deck, works on slides the run added BLANK.
 *
 * That gap is not theoretical. The worst bug of the session lived in exactly
 * it: `insertSceneIntoSlide` could not ask for a fresh proxy re-fetch, because
 * the re-fetch identified a chart's shapes as "the last N on the slide" — true
 * of a blank slide a run just added, false of the slide a user is working on.
 * So a chart over ten shapes lost its group AND its config tag on the web,
 * silently, on the single most-used action in the add-in. A demo run would not
 * have caught it: the demo path opted into the re-fetch and the everyday path
 * could not.
 *
 * Two charts, onto a slide that starts with one. The second lands on a slide
 * holding two already, which is where a positional heuristic goes wrong.
 */
/**
 * Where a scenario may draw so it lands on top of neither the chart already
 * there nor another scenario's.
 *
 * Several scenarios share a slide on purpose — `insertOntoUsedSlide` needs one
 * that already has a chart, `chartIsVisible` needs a before-and-after on a
 * surface the run already owns — and all of them pick the FIRST probe chart, so
 * they pick the same slide. Left to themselves they drew full-size at the
 * default origin and the owner opened a deck with four charts in a heap.
 *
 * The first fix put them in a fixed column down the right-hand edge, sized as a
 * share of the slide. That worked on 16:9 and failed on the first 4:3 deck it
 * met, exactly as its own test admitted it would: 720pt across leaves nothing
 * wide enough to clear a 480pt chart, and the round put the column at x=498
 * over a chart running to x=600. The test said the guarantee did not hold
 * there; it did not make the guarantee hold.
 *
 * The reason it could not is that the column's position was GUESSED. Nothing
 * asked where the chart actually was — the code assumed an origin of 60 and the
 * host had used 120. So: take the occupied box, which every caller already
 * holds, and lay the slots out in the band below it. That band exists on both
 * deck shapes, is computed rather than assumed, and does not care what the host
 * decided to do with the first chart.
 *
 * Pure, so the arithmetic that keeps them apart can be checked without a
 * PowerPoint.
 */
/**
 * Five, not four: the four charts, plus the square of rectangles
 * `degradesOverTime` measures with. That grid used to be placed by hand at
 * (20, 430) with a comment claiming clearance from the old right-hand column —
 * true of a column, false of a band, and the band is what the slots are now. It
 * draws onto the same slides, so it takes a slot like everything else rather
 * than being reasoned about separately.
 */
export const SIDE_SLOTS = 5;

/** The slot the degradation experiment's measurement grid draws in. */
export const GRID_SLOT = SIDE_SLOTS - 1;

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * The gap this battery leaves around anything it places.
 *
 * ONE FUNCTION BECAUSE A SECOND COPY IS HOW TWO NUMBERS DRIFT. It was inline in
 * `sideSlot` and needed by a caller that places a FIXED-size box, which has to
 * know the same margin to keep that box on the slide.
 */
export const slotMargin = (slide: { height: number }): number => Math.round(slide.height * 0.037);

/**
 * Fit a fixed-size box into the band, without letting it hang off the slide.
 *
 * `sideSlot` sizes a box TO the band. A caller that must keep its own width and
 * height — `rotatedShapePlacement` needs segments steep enough to tell two
 * readings apart, which makes 160x220 an invariant of the measurement rather
 * than a property of the deck — cannot use the band's height, and was simply
 * taking the band's `top`.
 *
 * Measured on the decks this battery actually runs: under a 300pt chart the
 * band is 160pt at its tallest and 1pt at its shortest, so a 220pt box hung
 * 40 to 200pt off the BOTTOM of the slide in every round, at both aspect
 * ratios. Nothing reported it, because `sideSlot` clamps its height with
 * `Math.max(1, …)` — a repair for an impossible band that hands back a
 * plausible 1pt one, so the caller never learns the band could not hold it.
 *
 * Overlapping what is already in the slot was always accepted here and still
 * is; the chart is deleted before the scenario returns. Hanging off the canvas
 * is a different thing, and not one anybody chose.
 */
export function fitInSlot(slot: Box, size: { width: number; height: number }, slide: Box | { height: number }): Box {
  const margin = slotMargin(slide);
  return {
    ...size,
    left: slot.left,
    top: Math.max(margin, Math.min(slot.top, slide.height - size.height - margin)),
  };
}

export function sideSlot(n: number, slide: { width: number; height: number }, occupied: Box): Box {
  const margin = slotMargin(slide);
  // A row, not a column: the free band on a slide holding one landscape chart
  // is underneath it on every deck shape this battery has met, while the strip
  // beside it disappears as soon as the slide narrows.
  const top = occupied.top + occupied.height + margin;
  const width = Math.floor((slide.width - (SIDE_SLOTS + 1) * margin) / SIDE_SLOTS);
  return {
    width,
    height: Math.max(1, slide.height - top - margin),
    left: margin + n * (width + margin),
    top,
  };
}

/** The box a probe chart occupies, from its target and the config it carries. */
const boxOf = (c: { target: { left: number; top: number }; cfg: ChartConfig }): Box => ({
  left: c.target.left,
  top: c.target.top,
  width: c.cfg.width ?? 480,
  height: c.cfg.height ?? 300,
});

const insertOntoUsedSlide: Scenario = async (prefix) => {
  const { found: hosts, blind, gap } = await probeCharts(prefix);
  const [host] = hosts;
  if (!host) return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to insert alongside" };
  const shown = await showSlide(host.target.slideId);
  const before = await slideCount();
  const added = [`${prefix} onto A`, `${prefix} onto B`];
  const slide = await slideSize();
  for (const [n, title] of added.entries()) {
    // Slots 0 and 1 of the band below the chart already there — `chartIsVisible`
    // takes slot 2 on this same slide. The old code cascaded these by 24pt,
    // which on a 480x300 chart is not a cascade: both landed on each other AND
    // on the chart already there. See `sideSlot`.
    const { left, top, ...box } = sideSlot(n, slide, boxOf(host));
    const c: ChartConfig = { ...cfg(title), ...box };
    await insertSceneIntoSlide(buildChart(c), { tagData: JSON.stringify(c), left, top });
  }
  const after = await slideCount();
  const { found: mine, blind: mineBlind, gap: mineGap } = await probeCharts(`${prefix} onto`);
  // The chart that was already there must still be a chart. A run that swept
  // the slide's existing shapes into its own group would take this with it —
  // and delete it on the next edit, because the parts tag would claim it.
  //
  // Both readings take the scan's OWN account of whether it saw the deck.
  // Without that, "I did not find your chart" and "I could not look" are the
  // same sentence — and this scenario said the first while meaning the second
  // on a real host, on a run where the previous scenario alone had taken 102
  // seconds and the host was answering deck reads short. It is the trap
  // `stop a run part-way` already has a rung for: "found nothing, therefore
  // nothing was left behind" is an assertion a blind scan satisfies for free.
  //
  // A blind scan is reported as SKIPPED, not as a pass and not as a failure:
  // nobody has evidence either way, and a verdict that cannot be attributed
  // costs a whole real-host round to chase.
  const rescan = await probeCharts(prefix);
  if (mineBlind || rescan.blind) return blindSkip(rescan.blind ? rescan.gap : mineGap);
  /**
   * A SKIP, matching the rule three lines above rather than inventing a third
   * answer. `after !== before` was FALSE when the host answered neither, so an
   * unreadable deck read as "it did not grow" — a clean pass on no evidence.
   *
   * The first repair reported it as a PROBLEM instead, which makes a FAILED
   * verdict: unknown dressed as a measurement, the same costume in a new
   * colour, and against this scenario's own doctrine — "a blind scan is
   * reported as SKIPPED, not as a pass and not as a failure: nobody has
   * evidence either way, and a verdict that cannot be attributed costs a whole
   * real-host round to chase." An unread COUNT is that, exactly.
   */
  const grew = deckGrowth(before, after);
  if (grew === null) return countlessSkip();
  const survivor = rescan.found.find((c) => c.cfg.title === host.cfg.title);
  const problems = [
    grew !== 0 && `the deck grew by ${grew} — the charts did not go ON the slide`,
    mine.length !== added.length && `${mine.length} of ${added.length} new charts are re-editable`,
    !survivor && "the chart already on the slide is no longer re-editable",
  ].filter(Boolean);
  return {
    ok: problems.length === 0,
    detail: problems.length
      ? problems.join("; ")
      : `${added.length} charts drawn ${shown ? "onto the visible slide" : "(host would not move the view)"}, ` +
        `all re-editable, the one already there untouched`,
  };
};

/**
 * The pane's actual entry point: a chart the USER has selected.
 *
 * Everything above reaches the machinery through targets `listChartsInDeck`
 * hands over. That is not how anyone uses the add-in. A user clicks a chart and
 * presses "Edit it", and the pane answers from `loadChartFromSelection` — a
 * different read, on a different collection (`getSelectedShapes`), which no
 * scenario here has ever exercised because of a comment that said it could not
 * be done. It can, at PowerPointApi 1.5, and this is the whole round trip:
 * select the shape, read it back as the pane does, edit through THAT target,
 * and confirm what comes back is the same chart.
 *
 * Deselects at the end by re-selecting the slide, never by
 * `setSelectedShapes([])` — see `clearShapeSelection`. A chart left selected
 * breaks the picture scenario below rather than this one.
 */
const editViaSelection: Scenario = async (prefix) => {
  if (!canSelectShapes()) return { ok: false, skipped: true, detail: "host cannot select shapes (PowerPointApi 1.5)" };
  // The ladder runs immediately before this and has just asked, properly, which
  // selection call this host stops answering. If it found one, spending a
  // budget here to be told nothing is pure cost: the answer is already in the
  // report, one line up, in more detail than this scenario could produce.
  //
  // Reported as SKIPPED rather than failed, because that is what it is. A
  // wedged host is a fact about the host; calling it a failure of the pane's
  // most-used read is the misattribution the ladder exists to prevent.
  if (selectionWedged) {
    return { ok: false, skipped: true, detail: `not attempted — the ladder just found this host ${selectionWedged}` };
  }
  const { found, blind, gap } = await probeCharts(prefix);
  const chart = leastLoadedChart(found, shapesDrawnOn);
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to select" };
  if (!(await selectShape(chart.target.slideId, chart.target.shapeId, selectionBudgetMs()))) {
    return { ok: false, skipped: true, detail: "the host would not select the chart" };
  }
  try {
    // Exactly what the pane does when someone presses "Edit it".
    //
    // On a short budget, because on PowerPoint on the WEB this call does not
    // come back at all once a shape has been selected programmatically —
    // measured, twice: 90 seconds to the bound, then the very next
    // `setSelectedSlides` also silent. At the full budget this scenario costs
    // three minutes to learn a thing we already know, so it is given ten
    // seconds and a name for what it found.
    const picked = await loadChartFromSelection(selectionBudgetMs());
    if (!picked) return { ok: false, detail: "the selected chart did not read back as an SSF chart" };
    const was = JSON.parse(picked.configJson) as ChartConfig;
    if (was.title !== chart.cfg.title) {
      return { ok: false, detail: `the selection read back a different chart: "${was.title}"` };
    }
    // And edit through the target the SELECTION produced, not the one the deck
    // scan produced. The two are built by different code from different reads,
    // and only one of them is what a user's edit actually travels on.
    const next = { ...was, title: `${was.title} (via selection)` };
    const target = await updateChartInSlide(buildChart(next), picked.target, { tagData: JSON.stringify(next) });
    if (!target) return { ok: false, detail: "the chart was gone from the slide after editing the selection" };
    // A blind READBACK is not a finding — see the sweep note above.
    const rescan = await probeCharts(prefix);
    if (rescan.blind) return blindSkip(rescan.gap);
    const round = rescan.found.find((c) => c.cfg.title === next.title);
    return {
      ok: !!round,
      detail: round
        ? "selected, read back, edited through the selection's own target, still re-editable"
        : "edited through the selection, and the result is no longer re-editable",
    };
  } catch (err) {
    // Silence is a fact about the HOST, not a defect of ours, and the two must
    // not read alike in a report. `setSelectedShapes` is GA at PowerPointApi
    // 1.5 and takes the call, but on the web it leaves the selection subsystem
    // unable to answer anything afterwards — `getSelectedShapes` and
    // `setSelectedSlides` both went silent for the full budget in the same
    // run. Two selection bugs are already open against the web host
    // (office-js#3083, #3698); this is the same family and worse.
    //
    // Reported SKIPPED rather than FAILED, with the reason: nothing in the
    // add-in is broken by it, because nothing but this battery ever selects a
    // shape programmatically. A red line here would send the next diagnosis
    // after our own code.
    // ...but only when it was a SELECTION call that went silent.
    //
    // This used to accept any timeout, and `isTimeout` is just as true of a
    // draw that stalled. A round on 2026-08-08 skipped with the sentence below
    // while its own trace read `gave up waiting what=drawing shapes 1-10 of 24`
    // and `at=redrawing the chart's shapes` — the selection subsystem was fine,
    // the host stopped answering during a redraw, and the report sent the reader
    // to two selection bugs that had nothing to do with it.
    //
    // A diagnosis is worth having only if it can be wrong.
    const at = stepOf(err) ?? (err instanceof Error ? err.message : undefined);
    if (isTimeout(err) && wedgedSelection(err)) {
      return {
        ok: false,
        skipped: true,
        detail:
          "the host stopped answering selection calls after a programmatic select — known web-host limitation, " +
          "same family as office-js#3083 / #3698; the pane's own Edit-it path is unaffected",
      };
    }
    // A timeout somewhere else is a real FAILURE and says where. Skipping it
    // would hide a stall behind a known-limitation label, which is how a host
    // problem stops being counted.
    if (isTimeout(err)) {
      return { ok: false, detail: `the host stopped answering${at ? ` at: ${at}` : ""} — not a selection failure` };
    }
    throw err;
  } finally {
    // Also on the short budget: once the subsystem is wedged this call is
    // silent too, and waiting the full budget for a deselect we already know
    // will not answer is ninety seconds spent on nothing.
    await clearShapeSelection(chart.target.slideId, selectionBudgetMs());
  }
};

/**
 * Stopping a run part-way, and what the deck looks like afterwards.
 *
 * Stop is cooperative — Office.js has no abort, so a sync already handed to
 * PowerPoint runs to completion whatever anyone wants, and the only honest
 * promise is "nothing further after this batch". That makes the interesting
 * question not "did it stop" but "what did it leave": a stop is indistinguishable
 * from a stall to every layer above it, and both leave a partial chart on the
 * slide. This asserts the two things a user is owed — the call ends, and it
 * ends by SAYING it stopped rather than reporting a chart it did not finish.
 *
 * The stop is requested before the render rather than mid-flight, which is the
 * one part a script cannot time: the pane's Stop button is pressed by a human
 * some seconds in. What is exercised is the same cooperative path, taken at its
 * first batch boundary instead of its fifth.
 */
const stopPartWay: Scenario = async (prefix) => {
  // If the USER has already asked to stop, this scenario must not run at all —
  // and above all must not clear their flag. It is the only code in the add-in
  // that calls `resetStop()`, and doing that on the way out of a scenario the
  // user had already cancelled would resume a battery they had stopped. The
  // guard in `runSelfTest` means this branch should be unreachable; it is here
  // because "should be unreachable" is how the flag got clobbered in the first
  // place.
  if (isStopRequested()) return { ok: false, skipped: true, detail: "not reached — the run was stopped" };
  const before = await slideCount();
  const c = cfg(`${prefix} stopped`);
  requestStop();
  let outcome: string;
  // HOW FAR IT ACTUALLY GOT, counted rather than described. The verdict below
  // used to read "stopped at a batch boundary", and across 69 archived rounds
  // this scenario has committed ZERO batches — `requestStop()` runs before the
  // insert, and `throwIfStopped()` sits at the top of the batch loop, so it
  // throws on iteration zero and no shape is ever queued.
  //
  // So the promise being tested is "a stop asked for BEFORE a draw prevents it",
  // which is worth testing and is not what `CLAUDE.md` claimed this was ("aborts
  // a draw mid-flight"). The mid-flight promise has never been exercised.
  //
  // THE SEAM FOR THE REAL TEST IS RIGHT HERE: `onPhase("commit", …)` fires once
  // per batch, so calling `requestStop()` from the first commit would abort a
  // genuinely half-drawn chart. Deliberately not done in passing — a mid-draw
  // abort leaves a PARTIALLY DRAWN SLIDE in the deck, and `same scale across the
  // deck` discovers its chart population from that same deck, so it would
  // contaminate every scenario after it. It wants its own pair, not a slipped-in
  // change.
  let commits = 0;
  try {
    const target = await insertSceneIntoSlide(buildChart(c), { tagData: JSON.stringify(c) }, (phase) => {
      if (phase === "commit") commits += 1;
    });
    outcome = target ? "the insert ran to completion and reported a chart" : "the insert finished without a chart";
  } catch (err) {
    outcome = isStopped(err) ? "stopped" : `threw something other than a stop: ${errorText(err)}`;
  } finally {
    resetStop();
  }
  const after = await slideCount();
  // The third assertion is "found nothing, therefore nothing was left behind" —
  // which a blind scan satisfies for free. On a real host this scenario reported
  // PASS off a scan that had just read 1 of 8 slides (`unread=7 slides=8`, the
  // line immediately before its own verdict), while two other scenarios reported
  // SKIPPED under exactly the same blindness. An assertion that cannot fail is
  // not a guard, and this one is guarding a promise — that Stop is
  // non-destructive — which nothing else checks.
  const leftovers = await probeCharts(`${prefix} stopped`);
  if (leftovers.blind && leftovers.found.length === 0) return blindSkip(leftovers.gap);
  // Same silence as the sibling above, guarding a stronger promise: that Stop
  // is non-destructive, which nothing else in this suite checks. `after !==
  // before` cannot fire when neither count was taken, so the one assertion
  // nobody else makes passed on no evidence. A SKIP for the same reason as
  // there — a verdict nobody can attribute is worse than no verdict.
  const grew = deckGrowth(before, after);
  if (grew === null) return countlessSkip();
  const problems = [
    outcome !== "stopped" && outcome,
    grew !== 0 && `the deck grew by ${grew} — a stopped insert added a slide`,
    // A stop must not leave a chart the pane would offer to edit: half a chart
    // that claims to be whole is worse than a visible mess.
    leftovers.found.length > 0 && "a stopped insert left a re-editable chart behind",
  ].filter(Boolean);
  return {
    ok: problems.length === 0,
    detail: problems.length
      ? problems.join("; ")
      : `stopped before ${commits === 0 ? "the first batch — no shape was ever queued" : `batch ${commits + 1}, after ${commits} committed`}` +
        ", nothing added, nothing left claiming to be a chart",
  };
};

/**
 * Stop pressed while the chart is HALF DRAWN — the case a user actually meets.
 *
 * `stopPartWay` above asks for the stop before the insert starts, so
 * `throwIfStopped()` throws on iteration zero and no shape is ever queued. Its
 * own comment says so, and the archive agrees: across 322 rounds it reported
 * "stopped before the first batch — no shape was ever queued" in 251 of them and
 * "stopped at a batch boundary" in the other 71. The mid-draw promise — the one
 * `CLAUDE.md` claimed was being tested, and the one the Stop button exists for —
 * had been exercised ZERO times on a real host.
 *
 * The seam is `onPhase("commit", …)`, which fires once per batch of
 * `SHAPES_PER_SYNC`. Requesting the stop from the FIRST commit aborts a
 * genuinely part-drawn chart.
 *
 * WHY THIS IS ITS OWN SCENARIO AND NOT A CHANGE TO THAT ONE. `stopPartWay`'s
 * comment names the blocker: a mid-draw abort leaves a partially drawn slide,
 * and `same scale across the deck` discovers its chart population from that same
 * deck, so it would contaminate every scenario after it. So this one CLEANS UP
 * AFTER ITSELF — it records the slide's shape ids before the insert, and deletes
 * whatever the aborted draw added. A new name also starts its own regression
 * series rather than reinterpreting 322 rounds of the old one.
 *
 * WHAT IT ASSERTS is the contract the pane already states in words — "Stopped —
 * anything already drawn was kept". Partial shapes are allowed. What is not
 * allowed is a half-drawn chart that claims to be whole: if the abort left
 * something carrying a config tag, the pane would offer to re-edit a chart that
 * was never finished, and Same Scale would rescale it as though it were.
 */
const stopMidDraw: Scenario = async (prefix) => {
  // Same guard, same reason as `stopPartWay`: never clear a stop the USER asked
  // for. This is the other scenario that calls `resetStop()`.
  if (isStopRequested()) return { ok: false, skipped: true, detail: "not reached — the run was stopped" };
  const { found: hosts, blind, gap } = await probeCharts(prefix);
  const [host] = hosts;
  if (!host)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart, so the deck is not ready" };

  /**
   * ITS OWN SLIDE, AND THE FIRST VERSION OF THIS DID NOT HAVE ONE.
   *
   * That version drew beside a probe chart and cleaned up by diffing the slide's
   * shape ids before and after, deleting whatever was new. It ran once against a
   * real host, on 2026-09-01 in round 353, and reported `10 of 10 shape(s) from
   * the aborted draw could not be cleaned up` — the ten shapes stayed on the
   * deck.
   *
   * The reason is a host fact this repo already had written down and I did not
   * join up: ids read back from a slide LISTING are not ids this host will
   * accept in `getItemOrNullObject`. `host-probe.ts` records the two id spaces
   * side by side — scratch ids reading `4123571114#123571113` while the deck
   * listed `256#109857222`, "both from the SAME `slideIds()` projection minutes
   * apart". `deleteShapesById` resolves each id that way, so every delete
   * refused. The sibling rotation scenario cleans up fine because its ids come
   * from the INSERT's own return value, not from a listing.
   *
   * So this draws on a slide of its own and removes the SLIDE, which is the
   * deletion route the whole harness already leans on. It also makes the
   * scenario strictly safer: no id diff, so no chance of the ids churning under
   * it and taking a probe chart with them, and any wreckage that does survive is
   * on a slide nothing else reads.
   */
  // The deck's ids BEFORE the add, so the cleanup can find the slide by what
  // the deck gained rather than by a name that will have changed. See
  // `deleteSlideAddedSince`.
  const deckIdsBefore = await deckSlideIds().catch(() => undefined);
  const slideId = await addSlideForChart();
  if (!slideId)
    return { ok: false, skipped: true, detail: "the host would not add a slide to draw the stopped chart on" };
  const slideCountBefore = await slideCount();

  // Big enough to need several batches at SHAPES_PER_SYNC — a chart that fits in
  // one commit cannot be stopped half way through by definition.
  const c = { ...cfg(`${prefix} mid-draw`, "waffle"), width: 320, height: 220 };
  const scene = buildChart(c);
  let commits = 0;
  let outcome: string;
  try {
    const target = await insertSceneIntoSlide(scene, { slideId, tagData: JSON.stringify(c) }, (phase) => {
      if (phase !== "commit") return;
      commits += 1;
      // THE STOP, from inside the draw. One batch has landed by the time this
      // runs, so the abort is genuinely mid-chart.
      if (commits === 1) requestStop();
    });
    outcome = target ? "the insert ran to completion and reported a chart" : "the insert finished without a chart";
  } catch (err) {
    outcome = isStopped(err) ? "stopped" : `threw something other than a stop: ${errorText(err)}`;
  } finally {
    resetStop();
  }

  // What landed before the abort, purely to report it. Nothing is deleted by
  // these ids — see the note above the slide add for why that does not work
  // here.
  const after = await slideShapeList(slideId);
  const landed = after?.length;

  // CLEAN UP FIRST, REPORT SECOND. Whatever this scenario decides, the deck must
  // be left as it was found — every scenario after this one reads it. One slide
  // delete, not N shape deletes.
  const sweptSlide = (await deleteSlideAddedSince(deckIdsBefore)) || (await deleteSlideById(slideId));
  const slideCountAfter = await slideCount();
  /**
   * A stop that landed before any shape did is the case `stop a run part-way`
   * already covers, and it cannot answer the mid-draw question — so it SKIPS
   * rather than passing, because a scenario that cannot conclude must say so.
   *
   * BUT NOT WHEN THE INSERT THREW SOMETHING ELSE. A real error before the first
   * batch is the loudest thing this scenario can discover, and skipping on it
   * would file that under "nothing was drawn". Only an insert that ended
   * cleanly without reaching a commit is excused here; a genuine throw falls
   * through to the verdict below and is reported.
   */
  if (commits === 0 && (outcome === "stopped" || outcome.startsWith("the insert")))
    return {
      ok: false,
      skipped: true,
      detail: `the insert never reached its first commit, so nothing was drawn to stop half way through (${outcome})`,
    };

  const stillClaimsToBeAChart = await probeCharts(`${prefix} mid-draw`);
  const problems = [
    outcome !== "stopped" && outcome,
    // The one that matters: half a chart must not present itself as a whole one.
    stillClaimsToBeAChart.found.length > 0 && "the aborted draw left a re-editable chart behind",
    // THE SLIDE MUST GO. Its shapes cannot be deleted by a listed id on this
    // host — round 353 proved that, 10 of 10 refused — so the slide is the unit
    // of cleanup, and a slide that will not go is wreckage every later scenario
    // reads.
    //
    // A REFUSED DELETE IS NOT THE SAME AS A SLIDE LEFT BEHIND, and the deck
    // count already in hand tells them apart. Round 370 made the difference
    // matter: at 4:3 the slide this scenario adds is listed, reported landed,
    // and then not there — so `deleteSlideById` correctly refuses to remove an
    // index whose id does not match, returns false, and the old wording
    // announced "the aborted draw is still in the deck" about a deck that had
    // never been cleaner. Wrong in the one direction a self-test must not be:
    // inventing wreckage sends someone looking for it.
    //
    // So the DECK decides, not the return value. If it shrank, the slide is
    // gone, and how it went is not this scenario's subject.
    slideCountAfter >= slideCountBefore &&
      (sweptSlide
        ? `the slide was reported deleted but the deck still holds ${slideCountAfter} of ${slideCountBefore}`
        : "the slide this drew on could not be deleted, so the aborted draw is still in the deck"),
  ].filter(Boolean);
  return {
    ok: problems.length === 0,
    detail: problems.length
      ? problems.join("; ")
      : `stopped inside the draw after ${commits} batch(es) with ${landed ?? "?"} shape(s) down; ` +
        `the slide went with them, nothing left claiming to be a chart` +
        // Reported, not counted against it: the deck is clean either way, and a
        // reader chasing the own-slide defect wants to know which happened.
        (sweptSlide ? "" : ` (the delete was refused, but the deck shrank to ${slideCountAfter} anyway)`),
  };
};

/**
 * A BIG chart on a slide of its own — the shipped feature nothing has tested.
 *
 * `offerOwnSlide` puts a chart on a freshly added slide when the current one is
 * crowded, and it fires precisely when the chart is BIG: that is the whole
 * point of the offer. So the shipped path is "add a slide, then draw a chart
 * across many batches onto it", and no scenario has ever walked it.
 *
 * WHY IT IS SUSPECTED. `stop a run mid-draw` moved onto its own slide in
 * `ca138f8` and has thrown `GeneralException` at `SlideCollection.getItem` on
 * every one of the six rounds since — after passing twice on the build before,
 * when it drew on the VISIBLE slide. That is a clean build boundary with one
 * `src/` change in it, and the elapsed times overlap (a pass at 488s, a failure
 * at 400s), so it is not the host tiring.
 *
 * WHAT THIS ISOLATES. That scenario also requests a STOP, so its failure has
 * two candidate causes and cannot separate them. This one changes exactly one
 * thing: same freshly-added slide, same multi-batch chart, NO stop. A failure
 * here means the shipped offer is broken for the charts it exists to serve. A
 * pass means the stop is implicated and the feature is safe.
 *
 * `insertSceneIntoSlide` holds ONE slide proxy for the whole draw — see the
 * long note above `getTargetSlide` — and that proxy came from
 * `slides.getItem(id)`. Office.js replays the statement chain on every sync, so
 * a `getItem` the host will not honour fails at the FIRST sync that uses it and
 * every one after. `addSlideForChart` is built to return a settled, positional
 * id precisely so that cannot happen; whether it is enough across eleven syncs
 * is the open question, and it is open because nothing has ever asked.
 */
/**
 * The own-slide offer's OTHER route: a chart handed over as a generated file.
 *
 * `offerNativeInsteadOfPicture` gives a too-dense chart a slide of its own as
 * native shapes, built as a one-slide .pptx and inserted in ONE call — ~280ms
 * against 46.4s of shape-by-shape drawing whose per-shape cost climbs from
 * 161ms to 784ms as the slide fills.
 *
 * IT SHIPPED WITH NO ROUND THAT COULD REACH IT. The route is behind a button,
 * and the battery cannot press buttons, so the only user-facing insert path in
 * the pane that no scenario exercised was also the newest one. That is the
 * shape of defect this project has been caught by four times.
 *
 * So `chartOnItsOwnSlideAsFile` was lifted out of the click handler into
 * `powerpoint.ts` and this scenario drives THE SHIPPED FUNCTION. Re-writing
 * the two calls here instead would have tested a copy — the same "fake that
 * cannot fail the way the real thing fails" that cost this archive four
 * separate fixes, wearing a different hat.
 *
 * WHAT IT ASKS, in order of what would hurt most to get wrong:
 *   - did the deck gain exactly one slide (not zero, not two)
 *   - is the chart ON it, rather than the slide arriving empty
 *   - is it RE-EDITABLE — the whole argument for native over a picture is that
 *     the config tag survives, so a scan that cannot find it makes the feature
 *     pointless
 * and then it takes the slide back, because every scenario after this one
 * reads the deck.
 */
const chartAsItsOwnFile: Scenario = async (prefix) => {
  if (!canInsertSlidesFromBase64())
    return { ok: false, skipped: true, detail: "this host has no insertSlidesFromBase64, so there is no file route" };

  const before = await deckSlideIds().catch(() => undefined);
  const countBefore = await slideCount();
  const c = { ...cfg(`${prefix} own file`, "waffle"), width: 320, height: 220 };
  const landed = await chartOnItsOwnSlideAsFile(buildChart(c), c);
  const countAfter = await slideCount();

  if (landed !== 1)
    return {
      ok: false,
      detail: `the host took ${landed} slide(s) from a one-slide file; the deck went ${countBefore} to ${countAfter}`,
    };

  // The chart has to be FINDABLE, not merely present. `listChartsInDeck` is
  // what Same Scale, the repair pass and the pane's own reload all go through,
  // so a chart it cannot see is a chart the product has lost.
  const scan = await listChartsInDeck();
  const mine = scan.charts.filter((x) => String(x.configJson).includes(`${prefix} own file`));

  // CLEAN UP FIRST, REPORT SECOND — by what the deck gained, because the id an
  // add hands back is not the one the deck lists it under later. See
  // `deleteSlideAddedSince`.
  const swept = await deleteSlideAddedSince(before);
  const countSwept = await slideCount();

  const problems = [
    countAfter - countBefore !== 1 && `the deck grew by ${countAfter - countBefore}, not 1`,
    mine.length !== 1 && `${mine.length} of 1 charts from the file are re-editable`,
    !scanIsComplete(scan) && `the deck scan was short, so "not found" may mean "not looked at" — ${scanGap(scan)}`,
    !swept && countSwept > countBefore && "the slide the file added could not be taken back",
  ].filter(Boolean);
  return {
    ok: problems.length === 0,
    detail: problems.length
      ? problems.join("; ")
      : `one slide from a generated file, its chart re-editable, slide returned (deck ${countBefore} to ${countAfter} to ${countSwept})`,
  };
};

const bigChartOnItsOwnSlide: Scenario = async (prefix) => {
  /**
   * ITS OWN CONTROL, TAKEN SECONDS BEFORE, because without one this scenario
   * cannot tell what it thinks it can.
   *
   * Its first run failed on the first batch with `GeneralException` at
   * `SlideCollection.getItem` — which would condemn the shipped `offerOwnSlide`
   * path. But it runs fourteenth, ~1900 syncs into a round, on the deck that
   * crashes 70% of its attempts, and "a fresh slide's id is unusable" and "the
   * host was already going" predict exactly the same failure there. This
   * project has been caught by that confound twice this week.
   *
   * Moving the scenario to run first would settle it and costs four battery
   * tests — the single-batch draw, the scratch-slide accounting, the picker,
   * and the announce-before-running guard — which are real invariants, not
   * noise to be quieted.
   *
   * A CONTROL IS CHEAPER AND STRICTER. The same insert, the same size, onto the
   * VISIBLE slide, in the same seconds. Both fail: the host was going, and this
   * says nothing about slide ids. Control passes and the own-slide fails: the
   * host was healthy and the ID is the difference, which is the whole question.
   * Both pass: the offer is safe on this host at this point in a round.
   *
   * Small, because the control exists to prove the host is answering rather
   * than to re-test batching — one commit is enough to establish that, and
   * spending eleven more here would make the control the expensive half.
   */
  const controlCfg = { ...cfg(`${prefix} control`, "clustered"), width: 160, height: 120 };
  let control: string;
  try {
    const t = await insertSceneIntoSlide(buildChart(controlCfg), { tagData: JSON.stringify(controlCfg) });
    control = t ? "drew" : "finished without reporting a chart";
    // Swept before the subject runs, so the deck this scenario leaves is the
    // deck it found — every scenario after this one reads it. Same shape as
    // `rotatedShapePlacement`'s cleanup: the group id plus its parts.
    const ids = [t?.shapeId, ...(t?.partIds ?? [])].filter(Boolean) as string[];
    if (t && ids.length) await deleteShapesById(t.slideId, ids);
  } catch (err) {
    control = `threw: ${errorText(err)}`;
  }

  // Read BEFORE the add, so the cleanup below can tell a slide that would not
  // go from one that was never there. See the verdict for why that matters.
  // The id list is for the same cleanup, by set difference — the slide will be
  // called something else by the time it is deleted. See `deleteSlideAddedSince`.
  const slideCountBefore = await slideCount();
  const deckIdsBefore = await deckSlideIds().catch(() => undefined);
  const slideId = await addSlideForChart();
  if (!slideId)
    // The host dropping the add is `addSlides`' own finding and not this
    // scenario's — reporting it here would blame the draw for a slide that
    // never arrived.
    return { ok: false, skipped: true, detail: "the host would not add a slide to draw on" };

  // The same shape and size as the mid-draw scenario's chart, so the two differ
  // in exactly one thing. A waffle at this size is ~103 shapes: eleven batches
  // at SHAPES_PER_SYNC, which is what makes this a test of the multi-sync path
  // rather than of a single commit.
  const c = { ...cfg(`${prefix} own slide`, "waffle"), width: 320, height: 220 };
  const scene = buildChart(c);
  let batches = 0;
  let outcome: string;
  try {
    const target = await insertSceneIntoSlide(scene, { slideId, tagData: JSON.stringify(c) }, (phase) => {
      if (phase === "commit") batches += 1;
    });
    outcome = target ? "drew" : "finished without reporting a chart";
  } catch (err) {
    outcome = `threw: ${errorText(err)}`;
  }

  const landed = (await slideShapeList(slideId))?.length;
  // CLEAN UP FIRST, REPORT SECOND — every scenario after this one reads the
  // deck. One slide delete, not N shape deletes.
  const swept = (await deleteSlideAddedSince(deckIdsBefore)) || (await deleteSlideById(slideId));
  const slideCountAfter = await slideCount();

  /**
   * BATCHES MATTER TO THE VERDICT. A chart that happened to fit in one commit
   * would pass this scenario while testing nothing — the multi-sync path is the
   * whole subject — so too few batches is a SKIP rather than a pass.
   */
  if (outcome === "drew" && batches < 2)
    return {
      ok: false,
      skipped: true,
      detail: `the chart drew in ${batches} batch(es), so nothing crossed a sync boundary on the new slide`,
    };
  /**
   * THE CONTROL DECIDES WHAT A FAILURE MEANS, so it is read before anything
   * else and it can turn this into a SKIP.
   *
   * If the same insert onto the VISIBLE slide also failed, seconds earlier,
   * then the host was already going and this scenario has learned nothing about
   * slide ids. Reporting that as a failure of the own-slide path would be the
   * confound written down as a finding — which is exactly what one round of
   * this scenario did before the control existed.
   */
  if (control !== "drew")
    return {
      ok: false,
      skipped: true,
      blind: true,
      detail:
        `the control insert onto the visible slide ${control} — the host was not answering, so nothing ` +
        `here says anything about a freshly-added slide (subject: ${outcome})`,
    };
  const problems = [
    outcome !== "drew" && outcome,
    // THE DECK DECIDES, NOT THE RETURN VALUE — same correction as the mid-draw
    // scenario one screen up, and round 370 is why. At 4:3 the slide this adds
    // is listed, reported landed, and then not there; `deleteSlideById` rightly
    // refuses to remove an index whose id does not match and returns false. The
    // old wording turned that into "the scenario could not delete the slide it
    // added", which reads as a slide left in someone's deck. There was no
    // slide. A self-test may not invent wreckage.
    !swept &&
      slideCountAfter > slideCountBefore &&
      "the scenario could not delete the slide it added, and the deck still holds it",
    typeof landed === "number" && landed === 0 && "the draw reported success and the slide came back empty",
  ].filter(Boolean);
  return {
    ok: problems.length === 0,
    detail: problems.length
      ? // THE CONTROL IS QUOTED IN THE VERDICT, because it is what makes the
        // verdict mean anything: a failure beside a healthy control is a fact
        // about the SLIDE ID, and that sentence is the whole finding.
        `${problems.join("; ")} — while the control insert onto the visible slide DREW, seconds earlier ` +
        `(${batches} batch(es), ${landed ?? "unreadable"} shape(s) landed)`
      : `${batches} batches onto a slide added moments before, ${landed} shapes, slide swept; control also drew`,
  };
};

/**
 * Whether anything is actually VISIBLE.
 *
 * Every other assertion in this file counts shapes and reads tags. All of them
 * pass for a chart drawn entirely in white, or at zero size, or off the edge of
 * the slide — a chart that is structurally perfect and invisible. Nothing in
 * this project has ever checked otherwise except a human looking at a deck.
 *
 * `Slide.getImageAsBase64` (PowerPointApi 1.8) is the host's own rasteriser, so
 * this compares a slide holding a chart against the same slide before it had
 * one. Deliberately not a pixel comparison against the SVG renderer: two
 * different text shapers never agree, and a check that cries wolf gets ignored.
 * The question here is the crude one nobody was asking — did drawing the chart
 * change what the slide looks like at all.
 */
/**
 * Whether a timed-out step was a SELECTION call, and so the known web-host wedge.
 *
 * Pure and exported because it is the whole of a diagnosis, and a diagnosis
 * that cannot be tested is a label. The steps that count are the ones this
 * scenario makes against the selection subsystem; a draw or a redraw that
 * stalls is a different fact with a different cause, and calling it a selection
 * limitation buries a host stall under a citation.
 *
 * Unknown (`undefined`) is deliberately NOT a selection wedge. An error with no
 * step attached could have come from anywhere, and the honest report for that is
 * a failure that says so rather than a confident wrong answer.
 */
export function wedgedSelection(err: unknown): boolean {
  const step = stepOf(err);
  // The MESSAGE too, and it is not belt-and-braces: a bounded wait rejects with
  // `PowerPoint did not respond while <what>`, and on the path this scenario
  // actually takes nothing attaches a step to that error at all. Reading only
  // the step made every wedge look like "somewhere unknown", which fails the
  // scenario that this branch exists to skip.
  const msg = err instanceof Error ? err.message : typeof err === "string" ? err : "";
  return /select/i.test(step ?? "") || /select/i.test(msg);
}

/**
 * What two rasters and a naming outcome amount to.
 *
 * Pure and exported for the reason `targetWithNoTagResult` is: the state that
 * exposed the bug needs several simultaneous host failures to reach through the
 * fake, and every attempt to arm them overshot into a different failure. The
 * rule is the thing that was wrong; it should be checkable without a
 * PowerPoint.
 *
 * `named` is whether the insert handed back an edit target. It is NOT whether
 * anything was drawn, and reading it that way is the bug this replaces:
 * `insertSceneIntoSlide` returns null when it has no id to hand back, which on
 * PowerPoint web happens after a perfectly good draw. The 2026-08-08 round
 * committed all three batches (`upTo=24 total=24`) and then failed to name the
 * group — `shapes.getItem("51")`, annotated by the host as originally an
 * `addGroup`, refused with InvalidParam — and the scenario reported "nothing
 * was drawn, so there is
 * nothing to look at" over twenty-four shapes that were on the slide. Calling a
 * drawn chart undrawn is the one mistake a visibility check must not make.
 *
 * So the image decides, and `named` only colours the answer:
 *
 * - changed, named — the plain pass.
 * - changed, unnamed — still a pass. The chart is visible; that it carries no
 *   config is a real defect, and a different scenario's business.
 * - unchanged, named — the chart is on the slide and cannot be seen. This is
 *   the defect the scenario was written for.
 * - unchanged, unnamed — "nothing drew" and "it drew invisibly" both fit and
 *   the image cannot separate them. Say so rather than pick one.
 */
/**
 * How long to keep listening for a call the scenario gave up on.
 *
 * Short on purpose. The host's eventual answer routinely arrives minutes later
 * when it arrives at all, and a battery cannot spend minutes per stall — but
 * the DIFFERENCE between "answered at 46s" and "still nothing" is most of what
 * a stall report is worth, and three seconds buys it.
 */
const LATE_ANSWER_WAIT_MS = 3_000;

/**
 * What to say about a scenario the host stopped answering during.
 *
 * The measurement, not the conclusion. Seven real-host rounds have produced
 * thirteen abandoned calls and not one late answer, and the slowest of 327
 * batches that DID answer took 29.2 seconds against a 45-second budget — so the
 * gap between 29s and 45s is empty and a stalled call looks like death rather
 * than slowness. That reading came out of a trace line's absence, though, and
 * absence is the evidence this project has most often misread: `settleUntagged`
 * was diagnosed as "ran and failed" for two sessions when it had never run.
 *
 * So the report says which it was, in words, every time. A future round that
 * finally sees a late answer will say so on the line, instead of waiting for
 * somebody to notice a message that is no longer missing.
 *
 * `after` is the other half, and the one a round summary could not carry at
 * all. Every draw stall on record is the FIRST batch of a scenario's draw,
 * which makes "what did the host answer immediately before it" the question —
 * and the log had nothing between a scenario announcing itself and its first
 * batch, so the only available account was at scenario level. Scenario order is
 * fixed, so "which scenario ran before" and "which position in the battery" are
 * the same variable and no number of rounds can separate them. The call before
 * the stall is a different variable, and it costs nothing to record.
 */
export function stallDetail(
  errText: string,
  late: string | undefined,
  after?: { call: string; idleMs: number },
): string {
  const head = `the host stopped answering during this scenario, so nothing was checked — ${errText}`;
  const answered = late
    ? `; the abandoned call DID come back afterwards (${late}), so the host was slow rather than gone`
    : `; the abandoned call had still not answered ${LATE_ANSWER_WAIT_MS / 1000}s later`;
  const before = after
    ? `; the last thing the host answered was "${after.call}", ${Math.round(after.idleMs / 100) / 10}s earlier`
    : "";
  return `${head}${answered}${before}`;
}

/**
 * Below this share of the original rasterised size, a change is reported but
 * not trusted.
 *
 * Three passes on record: 10064 → 15652 (+55%), 15704 → 16580 (+5.6%) and, on
 * 2026-08-11, 14868 → 14976 — **+0.7%, a hundred and eight bytes**. The gate
 * asserts that the two renders DIFFER, which a re-encode can satisfy on its own,
 * and all three read identically in the round file. That is a pass whose margin
 * is invisible, and this project has a rule about numbers with no baseline.
 *
 * Deliberately not a failure. A change is still a change, and turning a thin one
 * red would fail the round on a judgement no measurement here supports; the
 * honest move is to keep the verdict and say how thin it was, so a reader can
 * tell 0.7% from 55% without opening the round file.
 */
export const THIN_VISIBILITY_RATIO = 0.02;

/**
 * WHERE two renderings of a slide first differ, and how much of them does.
 *
 * The length delta alone has now said `+108 bytes` on three consecutive rounds
 * — 14868→14976, 14988→15096, 14856→14964 — from three different starting
 * sizes. A chart appearing in a rasterised slide does not cost the same
 * hundred and eight bytes three times by coincidence; something constant is
 * moving and the chart may not be in the picture at all.
 *
 * A length cannot tell those apart, and the round file carried nothing else. So
 * measure the difference rather than its size: an encoder's header, a timestamp
 * or a counter differs EARLY and in a handful of places, where a chart drawn
 * into the image differs across the body of the data. This is not a judgement
 * with a threshold — it is the two numbers a reader needs to make one, and it
 * costs a single pass over a string the round already holds.
 */
export function renderDifference(before: string, after: string): { at: number; differing: number; of: number } {
  const n = Math.min(before.length, after.length);
  let at = -1;
  let differing = Math.abs(before.length - after.length);
  for (let i = 0; i < n; i++) {
    if (before[i] === after[i]) continue;
    if (at < 0) at = i;
    differing++;
  }
  return { at: at < 0 ? n : at, differing, of: Math.max(before.length, after.length) };
}

/**
 * @param stable Whether two renders of the UNCHANGED slide came back identical.
 *   `undefined` when the control could not be taken. This is what separates
 *   "the chart is visible" from "this host's rasteriser is not deterministic",
 *   and without it a before/after difference means neither.
 */
export function visibilityVerdict(
  before: string,
  after: string,
  named: boolean,
  stable?: boolean,
): { ok: boolean; detail: string; skipped?: true } {
  if (after !== before) {
    const delta = after.length - before.length;
    const share = before.length ? Math.abs(delta) / before.length : 1;
    const pct = Math.round(share * 1000) / 10;
    const diff = renderDifference(before, after);
    // Where the two renders part company, beside how much longer one is. A
    // difference that starts in the first few percent of the data and touches
    // little of it is a header; one spread through the body is a picture.
    const where = `, first differ at ${Math.round((diff.at / Math.max(1, diff.of)) * 100)}% in, ${diff.differing} byte(s) differing`;
    // The control decides what the difference is EVIDENCE of. A host whose
    // rasteriser answers differently for an unchanged slide makes this whole
    // gate meaningless, and it would look identical to a pass.
    const noise =
      stable === false
        ? " — but two renders of the UNCHANGED slide also differed, so this proves NOTHING about the chart"
        : stable === undefined
          ? // WHY THE CONTROL IS MISSING, not merely that it is. The
            // correlation held exactly over the first 14 rounds that reported
            // blind — every blind round carried one rasterise stall at its full
            // 20000ms budget, every sighted round carried none; 8 blind / 8
            // stalls, 6 sighted / 0 stalls.
            //
            // So this is not an unknown. The SECOND rasterise of the same slide
            // hangs — the first one answered, which is where `before` came from —
            // and that costs the project its only mechanical evidence that a
            // drawn chart is visible.
            //
            // RE-MEASURED 2026-08-20 over all 92 archived rounds: 36 blind, 56
            // sighted — a 39% blind rate. The comment said "more than half of
            // all rounds" until then, which was true of the sample it was
            // written from and has not been true for a while. A rate quoted in
            // a comment is a measurement with no re-measure attached; date it
            // and say what it was measured over, or it becomes folklore.
            //
            // MATCHED ON `op`, NOT ON WORDING, and this line is why the rule
            // exists. It used to read `/rasteris/i.test(lastStall.what)`, which
            // worked only while every rasterise traced the one string
            // "rasterising a slide". Once each call site was given its own name
            // the control render began stalling as "the visibility CONTROL
            // render (same slide, back to back)" — no "rasteris" in it — so this
            // branch stopped firing and the scenario reported the cause as
            // UNKNOWN while the cause sat one field away in its own trace.
            // Round 113 is the first archived proof. The regex is kept as a
            // fallback ONLY for rounds archived before `op` existed.
            lastStall && (lastStall.op === RASTERISE_OP || /rasteris/i.test(lastStall.what))
            ? ` (no control render: the second rasterise of the same slide stalled — ${lastStall.what})`
            : " (no control render, so an unstable rasteriser cannot be ruled out)"
          : "";
    return {
      // A DIFFERENCE WITHOUT A CONTROL IS NOT EVIDENCE, and this returned `ok:
      // true` for both branches its own `noise` text disclaims. `stable ===
      // false` means the host re-rendered an UNCHANGED slide differently, so a
      // before/after difference means nothing; `stable === undefined` means the
      // control was never taken. 23 of 69 archived rounds reported this scenario
      // green with no control at all — a third of the project's only mechanical
      // evidence that a drawn chart can be seen, resting on nothing.
      //
      // Skipped rather than failed: the chart may well be perfectly visible, and
      // calling that a failure would be the opposite lie. "A skip is not a flip"
      // is this repo's own rule, and the gate already honours it.
      ok: stable === true,
      ...(stable === true ? {} : { skipped: true as const }),
      detail:
        // "in PowerPoint's own render", not "on screen", and the distinction is
        // office-js#6498: on the web an inserted shape can appear in the slide
        // PREVIEW and not in the main view. `getImageAsBase64` renders exactly
        // that preview, so this gate measures the surface the issue says can
        // disagree with the canvas. It is still the only mechanical evidence
        // this project has that a chart it drew can be seen — it just cannot be
        // reported as "a human would see it", and no add-in API can close that
        // gap. Which is one more reason the battery leaves its slides behind.
        `drawing the chart changed PowerPoint's own render of the slide (${before.length} → ${after.length} bytes, ${
          delta >= 0 ? "+" : ""
        }${delta}, ${pct}%${where})${noise}` +
        (share < THIN_VISIBILITY_RATIO
          ? " — a THIN margin: this is a change, but too small to tell a drawn chart from a re-encode"
          : "") +
        (named ? "" : " — though the host would not name the chart afterwards, so it carries no config"),
    };
  }
  return {
    ok: false,
    detail: named
      ? "the slide renders identically with and without the chart — nothing is visible"
      : "the slide renders identically and the host would not name a chart on it — " +
        "cannot tell a chart that never drew from one that drew invisibly",
  };
}

/**
 * Say what is about to be tried, BEFORE trying it.
 *
 * `scenario starting` names the scenario and nothing finer, and a scenario that
 * makes several host calls can die in any of them. A real host proved the
 * difference twice, the same way both times: the run log's last line was the
 * scenario announcing itself and nothing after it was ever written, so the
 * evidence narrowed the cause to "somewhere in here" and stopped.
 *
 * The first time cost `the chart is actually visible` four rounds. Wrapping its
 * calls turned the fifth into `rasterising the empty slide`, which is a cause
 * rather than a location, and the scenario was fixed the same day. The second
 * time was `what makes a long run slow down`, on 2026-08-08, which announced
 * itself at 26.9s and took the tab with it — and could have been either of two
 * calls, because it was not wrapped.
 *
 * A handful of extra entries in a two-thousand-entry ring is a cheap price for
 * a log that names the call. Shared rather than copied, because the next
 * scenario to need it should not have to learn this a third time.
 */
const stepsOf =
  (message: string) =>
  async <T>(what: string, fn: () => Promise<T>): Promise<T> => {
    trace("selftest", message, { what });
    return fn();
  };

const chartIsVisible: Scenario = async (prefix) => {
  const attempt = stepsOf("visibility step");

  // A slide this run put in the deck EARLIER — never one added moments ago.
  //
  // It used to take its own scratch slide, and that is the call which has now
  // killed PowerPoint on the web five rounds running. The fifth was the
  // experiment that says so: picked alone, the run reached this scenario at
  // 61.5s with only the two inserts it depends on in front of it, added a
  // scratch slide at 61.5s, logged `rasterising the empty slide` at 61.8s, and
  // the tab died. The four rounds before it died ten minutes and nine
  // scenarios in, which is why "the scenario kills the host" and "ten minutes
  // of drawing kills the host, and this is what happened to be running" both
  // fit. They do not both fit any more.
  //
  // A fresh slide is the worst surface this host offers — its id does not
  // round-trip through `getItem`, its shapes will not list, and
  // `getImageAsBase64` has now failed on one in five distinct ways: a
  // GeneralException, a call taken that produced nothing, a sync never
  // answered, a ninety-second stall, and this. Every one of those is a fresh
  // slide; a pre-existing slide has never been asked.
  //
  // So ask the question that has not been asked, on the surface that has not
  // failed. Before-and-after on ONE slide is still what isolates the chart —
  // that part was never the problem — and the slide simply need not be empty
  // for the comparison to mean anything. It also drops the scratch add and the
  // delete, and the delete is separately implicated: one earlier round died on
  // it rather than on the rasterise.
  const { found, blind, gap } = await probeCharts(prefix);
  const host = leastLoadedChart(found, shapesDrawnOn);
  if (!host)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no slide from this run to draw a chart on" };
  const slideId = host.target.slideId;

  // Named apart from the old step on purpose. If a tab still dies here, the
  // crash log says `rasterising a slide that already existed` and that is the
  // remaining reading — this host cannot rasterise AT ALL — rather than another
  // repeat of the one now settled.
  const before = await attempt("rasterising a slide that already existed", () =>
    slideImageBase64(slideId, 640, "the visibility BEFORE render"),
  );
  if (!before) return { ok: false, skipped: true, detail: "host will not rasterise a slide (PowerPointApi 1.8)" };
  // The SAME slide, rasterised again, with nothing drawn in between.
  //
  // This gate asserts that the two renders differ, and until 2026-08-11 that
  // was all it could say. Its length delta read +108 bytes on three
  // consecutive rounds from three different starting sizes, which looked
  // exactly like a header or a counter; `renderDifference` then reported the
  // two renders differing across 97% of their bytes, which kills the header
  // reading and leaves two others standing. Either the chart is in the picture,
  // or this host's rasteriser does not produce the same bytes twice — and a
  // before/after pair cannot tell those apart, because both make a drawn chart
  // and an untouched slide look equally different.
  //
  // One extra rasterise decides it. If two renders of an UNCHANGED slide are
  // identical, then a difference afterwards is the chart and this gate means
  // what it says. If they differ, the gate has been passing on encoder noise
  // and every "the chart is visible" verdict on record is worth nothing.
  //
  // Costs one call on a slide the run has already rasterised safely, which is
  // the operation this scenario exists to have proven.
  // THE ONE THAT STALLS. Named so the next stall can say so instead of leaving
  // it to be inferred from where the line sits in the trace.
  //
  // AND NOW PRECEDED BY A GAP, because that stall is the whole reason this gate
  // is blind in 38% of rounds. Across 104 archived rounds the correlation is
  // exact: 40 blind rounds carry a rasterise stall, 64 sighted rounds carry
  // none — and all 40 stalls are this call, issued immediately after the BEFORE
  // render answered. See `rasterGap`.
  await rasterGap();
  const again = await attempt("rasterising the same slide a second time", () =>
    slideImageBase64(slideId, 640, "the visibility CONTROL render (same slide, back to back)"),
  );
  const control = again === undefined ? undefined : again === before;
  // Small, and tucked into the bottom-right corner.
  //
  // Sharing a slide is the price of never rasterising a fresh one, and the
  // first round after that trade produced a slide with two full-size charts
  // drawn over each other — legible to a rasteriser, ugly to a human, and this
  // battery leaves its slides in the deck for a human to look at. So the chart
  // this scenario draws is deliberately not a specimen: a quarter-size chart in
  // the corner changes the image just as decisively, draws fewer shapes, and
  // sits clear of a full-size chart placed anywhere the probe deck puts one.
  //
  // Measured against the slide rather than assumed: `slideSize` is cached after
  // its first read, so this costs nothing, and a 4:3 or 16:10 deck puts the
  // corner somewhere else.
  // Slot 2 — the two `insertOntoUsedSlide` draws take 0 and 1 on this same
  // slide, because both scenarios pick the first probe chart. See `sideSlot`.
  const { left, top, ...box } = sideSlot(2, await slideSize(), boxOf(host));
  // TWO categories, not the sample's full set — the box was already quarter
  // size and the shape count was not, which is what kept this scenario from
  // running.
  //
  // Nine rounds, and it completed ONE of them. Every other round skipped it on
  // `PowerPoint did not respond while drawing shapes 1-10 of 24 (45s)`. The
  // paragraph above already reached for "draws fewer shapes" and only shrank
  // the frame; `sampleConfig("clustered")` is 24 shapes over three batches
  // whatever size it is drawn at, and this host stalls the first batch of a
  // draw often enough that three of them is a coin flip repeated.
  //
  // The counterbalanced scenario made exactly this trade already and completes
  // eight rounds in nine — its `tiny` chart is one batch, and its note says
  // why: "what is under test is whether the FIRST sync of a draw comes back,
  // and one batch asks that exactly". Nothing here is about density either.
  // The question is whether the slide's PICTURE changes when a chart is drawn
  // on it, and seven shapes change it as decisively as twenty-four.
  const c: ChartConfig = {
    ...cfg(`${prefix} visible`),
    ...box,
    data: { categories: ["A", "B"], series: [{ name: "s", values: [1, 2] }] },
  };
  const drawn = await attempt("drawing the chart", () =>
    insertSceneIntoSlide(buildChart(c), { slideId, tagData: JSON.stringify(c), left, top }),
  );
  // A null target is NOT "nothing was drawn", so the rasterise happens either
  // way — see `visibilityVerdict`, which is where that used to be decided
  // wrongly. The measurement is the IMAGE, and the slide id came from the
  // caller rather than from the draw.
  const after = await attempt("rasterising the slide with the chart", () =>
    slideImageBase64(slideId, 640, "the visibility AFTER render"),
  );
  if (!after) return { ok: false, detail: "the host rasterised the slide before the chart but not after it" };
  // No cleanup. Every other scenario leaves its slides in the deck, and this
  // one only cleaned up because a scratch slide is a control surface nobody
  // would want to open. A chart on a slide the run already owns is an ordinary
  // result, so the delete — which cost one round on its own — simply goes.
  return visibilityVerdict(before, after, !!drawn, control);
};

/**
 * Does a rasterise poison the next draw?
 *
 * Eight rounds of stalls have had every one of their candidate causes killed by
 * a later round: the scenario, the scenario before it, the tab's age, and the
 * idle gap before the sync. Round 10 left exactly one thing standing. With the
 * predecessor call recorded on EVERY first batch rather than only on stalls,
 * its populations separate cleanly:
 *
 *   survivors followed  moving the view to a slide, counting the deck's slides,
 *                       writing the chart's origin tag, re-reading a slide to
 *                       tag the chart it would not tag, selecting a shape,
 *                       reading the selected chart
 *   the stall followed  rasterising a slide
 *
 * Twenty-nine surviving first batches across two rounds, not one of them after
 * a rasterise; two stalls after a rasterise, in consecutive rounds. And
 * `selecting a shape` — round 9's other stall — turns up among round 10's
 * survivors, which is what a non-cause looks like.
 *
 * It is still confounded, and exactly. Only `chartIsVisible` rasterises before
 * drawing, so "the draw followed a rasterise" and "the draw was that scenario"
 * are the same event, one per round. This project has now been caught by that
 * shape three times, and reasoning has never once broken it.
 *
 * So: two arms, one scenario, one slide, seconds apart. The CONTROL draws after
 * a cheap read; the TEST draws after a rasterise. Everything a scenario-level
 * account could appeal to — which scenario, what ran before it, how old the tab
 * is, how loaded the deck is — is held identical between them, and the only
 * difference is the call immediately before the draw.
 *
 * - both arms draw          → the rasterise is innocent, and `chartIsVisible`
 *                             is stalling for a reason still unnamed.
 * - only the test arm stalls → the call is the cause, on a surface that is not
 *                             `chartIsVisible`, and there is finally something
 *                             to fix rather than to watch.
 * - both stall              → the slide or the moment, not the rasterise.
 *
 * Deliberately small charts. This is a measurement, not a specimen, and it runs
 * every round — a full-size pair would add a minute to a battery that already
 * takes ten.
 */
/**
 * DOES THIS HOST PLACE A ROTATED SHAPE BY ITS ROTATED BOX, OR ITS UNROTATED ONE?
 *
 * After 333 rounds nobody could answer, because `Shape.rotation` had never once
 * been written on a real host. The battery draws only `clustered`, whose single
 * line node is the horizontal baseline, so `addSegment`'s rotated-rectangle
 * branch — the one every diagonal in the product goes through — had zero
 * executions outside the fake.
 *
 * It is not a curiosity. `addSegment` draws a solid diagonal as a rectangle of
 * the segment's LENGTH, placed at its midpoint, then rotated about the centre;
 * `arrowheadBox` offsets its box on the same assumption. If this host means the
 * post-rotation bounding box by `left`/`top`/`width`, then every diagonal in
 * every line, scatter, radar and violin chart is drawn in the wrong place and
 * the wrong size — and no round could have noticed.
 *
 * THE DISCRIMINATOR IS THE WIDTH, and it is chosen because it needs no
 * agreement about where the chart was placed. For a diagonal of length L whose
 * horizontal extent is W:
 *
 *     width ~= L   the host stores the box BEFORE rotation (what we assume)
 *     width ~= W   the host stores the bounding box AFTER it (we are wrong)
 *
 * L and W differ by more than a third on the segments this chart draws, so the
 * two readings cannot be confused. `left`/`top` would need the insertion offset
 * and give a weaker answer for more work.
 *
 * A host that cannot rotate at all (below PowerPointApi 1.10) SKIPS: it has no
 * opinion to report, and a scenario that cannot conclude must say so rather
 * than pass.
 *
 * ── WHICH LEAVES ONE BRANCH BELOW LOOKING LIKE IT BREAKS THAT RULE ──
 *
 * A mining pass over rounds 334-346 filed exactly that, and it is worth
 * answering here rather than in a commit nobody will find. The sentence above
 * was written when this scenario WAS the placement measurement, and it has not
 * been true since the measurement moved. There are two questions now:
 *
 *   does the rotated draw complete on a real host?   THIS scenario answers it
 *   where do the rotated shapes land?                the PROBE answers it
 *
 * The `ok: true` below is the first question and nothing else. It is a real
 * pass on a real path: `Shape.rotation` had never been written on a live host
 * in 333 rounds, so `addSegment`'s rotated branch — which every diagonal in the
 * product takes — had never once executed outside the fake. What it must never
 * be read as is an answer to the second question, which is why its detail says
 * so and names the probe.
 *
 * AND THE POINTER NOW RESOLVES, which it did not when that branch was written.
 * Until 2026-08-31 `rotation-keeps-the-unrotated-box` answered `unreadable` in
 * all eleven rounds it had run — because `unreadable` locked its row against
 * the `unrotated-box` its own samples carried four times, and the contradicting
 * answer never. See `UNINFORMATIVE` in host-probe.ts. Sending a reader to a
 * sheet that said "unreadable" is most of why this branch looked like a pass
 * over a void.
 */
const rotatedShapePlacement: Scenario = async (prefix) => {
  const { found: hosts, blind, gap } = await probeCharts(prefix);
  const [host] = hosts;
  if (!host)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to place a line chart beside" };
  const shown = await showSlide(host.target.slideId);
  const slide = await slideSize();
  const title = `${prefix} rotation`;
  // Slot 3 is shared with the rasterise arm, which is harmless here: this
  // scenario reads its shapes BY NAME, and `line-` is a name no other chart in
  // this battery emits.
  //
  // `fitInSlot` because the box below keeps its own 160x220 — see why in the
  // comment on it — and the band under a 300pt chart is 160pt at its tallest.
  // Taking the band's `top` unmodified hung this chart 40 to 200pt off the
  // bottom of the slide in every archived round.
  const { left, top } = fitInSlot(sideSlot(3, slide, boxOf(host)), { width: 160, height: 220 }, slide);
  /**
   * A FIXED, TALL box rather than the slot's — and deliberately not the slot's
   * shape at all.
   *
   * The discriminator needs segments steep enough that the length and the
   * horizontal extent cannot be confused. Taking the slot's aspect made that a
   * property of the DECK: on a 4:3 deck the slot is short and wide, the segments
   * flatten, and the scenario skipped with "no segment was steep enough" — which
   * is a scenario that reports nothing precisely when the deck is ordinary.
   *
   * Narrow and tall makes steep segments an invariant of the measurement instead.
   * Overlapping whatever is in that slot costs nothing here: this chart is
   * deleted before the scenario returns, which is the same reason it can borrow
   * an occupied slot at all.
   *
   * Three points, so the two segments slope hard in OPPOSITE directions — a
   * single slope could be matched by a host that happens to mirror the box.
   */
  const c: ChartConfig = {
    ...cfg(title, "line"),
    width: 160,
    height: 220,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [1, 9, 2] }] },
    decorations: { segmentLabels: false },
  };
  const scene = buildChart(c);
  const placed = await insertSceneIntoSlide(scene, { tagData: JSON.stringify(c), left, top });

  const drawn = await shapeGeometryByName(host.target.slideId, (n) => /^line-\d+-\d+$/.test(n));
  /**
   * A MEASUREMENT, NOT A SPECIMEN — taken away as soon as it has been read.
   *
   * Every side slot is already claimed by a scenario that leaves its chart
   * behind, and `leaves no two charts stacked on one slide` rightly fails a
   * battery that stacks one on another — it caught this on the first run. So
   * this chart borrows a slot and gives it back. Taking a SIXTH slot instead
   * would narrow every other one, and quietly change what every other scenario
   * draws and measures.
   */
  const clean = async (): Promise<void> => {
    if (!placed) return;
    const ids = [placed.shapeId, ...(placed.partIds ?? [])].filter(Boolean);
    if (ids.length) await deleteShapesById(placed.slideId, ids);
  };
  if (!drawn || !drawn.length) {
    await clean();
    /**
     * THE DRAW IS THE FINDING HERE, and it is not nothing.
     *
     * This scenario was built to measure where a rotated shape LANDS, and on
     * this host it cannot: reading a chart's parts means reading a group's
     * children, and this host refuses both routes into one — `threw` on
     * `group-reports-its-children` and `group-children-via-getcount` in 29 of
     * the last 30 rounds. The archive settled that before the read was written.
     * Rounds 334-335 skipped because the lookup was at slide level; 339 skipped
     * with the descent in place, for this reason instead.
     *
     * What it still does, and nothing else in the battery does, is DRAW a chart
     * whose every segment is a rotated rectangle. `Shape.rotation` had never
     * been written on a real host in 333 rounds, so `addSegment`'s rotated
     * branch — the path every diagonal in the product takes — had never once
     * executed outside the fake. Getting the chart in without an error is a
     * real answer about a real path, and it is reported as one.
     *
     * The measurement moved to `rotation-keeps-the-unrotated-box`, which asks
     * the host directly instead of through our own pipeline.
     */
    return {
      ok: true,
      /**
       * COUNTED FROM THE SCENE, not from `partIds`.
       *
       * Round 341 reported "drew 0 rotated segment(s) without error", which is
       * barely better than the skip it replaced: `partIds` is empty on this
       * host, so the count was measuring the same refusal the rest of the
       * sentence goes on to describe. The SCENE is what we handed the renderer,
       * and its line nodes are exactly the rotated rectangles `addSegment` was
       * asked to make — a number we know without asking the host anything.
       */
      detail: `drew a ${scene.nodes.filter((n) => n.kind === "line").length}-segment line chart without error; this host will not report a group's children, so where the segments landed is unmeasurable here — see the rotation-keeps-the-unrotated-box probe`,
    };
  }
  if (drawn.every((s) => s.rotation === null)) {
    await clean();
    return {
      ok: false,
      skipped: true,
      detail: `this host reports no rotation on ${drawn.length} segment(s) — below PowerPointApi 1.10, so it has no placement opinion to read`,
    };
  }

  // What the SCENE asked for, for the same named segments.
  const want = new Map(
    scene.nodes
      .filter((n): n is Extract<typeof n, { kind: "line" }> => n.kind === "line" && /^line-\d+-\d+$/.test(n.name ?? ""))
      .map((n) => [n.name as string, { len: Math.hypot(n.x2 - n.x1, n.y2 - n.y1), ext: Math.abs(n.x2 - n.x1) }]),
  );
  const judged = drawn
    .map((s) => ({ s, w: want.get(s.name) }))
    .filter((p): p is { s: (typeof drawn)[number]; w: { len: number; ext: number } } => !!p.w)
    // Only segments where the two readings are far enough apart to tell apart.
    .filter((p) => p.w.len - p.w.ext > 4);
  if (!judged.length) {
    await clean();
    return { ok: false, skipped: true, detail: "no segment was steep enough to distinguish the two placements" };
  }

  const preRotation = judged.filter((p) => Math.abs(p.s.width - p.w.len) < Math.abs(p.s.width - p.w.ext));
  const sample = judged[0];
  const detail =
    `${preRotation.length} of ${judged.length} segment(s) sized by the UNROTATED box` +
    ` (e.g. ${sample.s.name}: host ${sample.s.width.toFixed(1)}pt, length ${sample.w.len.toFixed(1)}pt,` +
    ` horizontal extent ${sample.w.ext.toFixed(1)}pt, rotation ${String(sample.s.rotation)})` +
    `${shown ? "" : " (host would not move the view)"}`;
  await clean();
  return {
    ok: preRotation.length === judged.length,
    detail:
      preRotation.length === judged.length
        ? detail
        : `${detail} — this host places rotated shapes by their POST-rotation box, so every diagonal in the product is drawn wrong`,
  };
};

const rasteriseThenDraw: Scenario = async (prefix) => {
  const attempt = stepsOf("rasterise-draw step");
  const { found, blind, gap } = await probeCharts(prefix);
  const host = leastLoadedChart(found, shapesDrawnOn);
  if (!host) return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to draw beside" };
  const slideId = host.target.slideId;
  const slide = await slideSize();
  const cell = sideSlot(GRID_SLOT, slide, boxOf(host));
  const w = Math.max(1, Math.floor(cell.width / 4) - 3);
  /**
   * Two categories, one series, no decorations — about eight shapes, one batch.
   *
   * The first version drew the battery's ordinary sample chart, twenty-four
   * shapes over three batches, and each arm cost most of a minute. Four of
   * those is three and a half minutes on a battery already at thirteen, and the
   * measurement does not need a specimen: what is under test is whether the
   * FIRST sync of a draw comes back, and one batch asks that exactly.
   */
  const tiny = (title: string): ChartConfig => ({
    kind: "clustered",
    title,
    width: w,
    height: cell.height,
    data: { categories: ["A", "B"], series: [{ name: "s", values: [1, 2] }] },
  });
  const arm = async (label: string, n: number, before: () => Promise<unknown>): Promise<DrawArm> => {
    await attempt(`${label} #${n}: the call before the draw`, before);
    const c = tiny(`${prefix} ${label} ${n}`);
    try {
      await attempt(`${label} #${n}: drawing`, () =>
        insertSceneIntoSlide(buildChart(c), {
          slideId,
          tagData: JSON.stringify(c),
          left: cell.left + n * (w + 3),
          top: cell.top,
        }),
      );
      return { drew: true, why: "" };
    } catch (err) {
      // A stall here is the RESULT, not a failure of the battery — so it is
      // caught per arm rather than allowed to reach the runner, which would
      // turn the whole scenario into a blind skip and discard the other arms.
      return { drew: false, why: isTimeout(err) ? "the host stopped answering" : errorText(err) };
    }
  };
  // COUNTERBALANCED, and the first version was not — which round 11 showed is
  // not a nicety. With the cheap arm first and the rasterise arm second, the
  // rasterise arm always ran later, on a fuller slide and an older tab, and the
  // scenario duly reported `the draw after a RASTERISE did not land` for a
  // difference that position explains just as well. The same round contained a
  // draw after a rasterise that SUCCEEDED 150 seconds earlier, so the verdict
  // was a claim the round's own log contradicted.
  //
  // Rasterise, cheap, cheap, rasterise: each call type runs once early and once
  // late, so position is held across the pair instead of confounded with it.
  const raster = () => slideImageBase64(slideId, 640, "the rasterise-poisons-the-next-draw arm");
  const cheap = () => slideCount();
  const rasterEarly = await arm("after a rasterise", 0, raster);
  const cheapEarly = await arm("after a cheap read", 1, cheap);
  const cheapLate = await arm("after a cheap read", 2, cheap);
  const rasterLate = await arm("after a rasterise", 3, raster);
  trace("selftest", "rasterise-draw arms", {
    rasterise: [rasterEarly.drew, rasterLate.drew],
    cheap: [cheapEarly.drew, cheapLate.drew],
    order: "raster, cheap, cheap, raster",
  });
  return rasteriseArmVerdict([rasterEarly, rasterLate], [cheapEarly, cheapLate]);
};

/** One arm's outcome: did the draw land, and if not, what the host said. */
export interface DrawArm {
  drew: boolean;
  why: string;
}

/**
 * What four counterbalanced arms mean together — pure, because the reading is
 * the thing that can be wrong, and the first version of it WAS.
 *
 * Two arms were not enough. With the cheap arm first and the rasterise arm
 * second, the rasterise arm always ran later, and round 11 duly reported "the
 * draw after a RASTERISE did not land" — for a difference that position
 * explains just as well, in a round whose own log contained a draw after a
 * rasterise succeeding 150 seconds earlier. A diagnostic that manufactures a
 * finding is worse than no diagnostic.
 *
 * So each call type runs once early and once late, and the verdict names the
 * CALL only when the call is what separates them: every rasterise arm failing
 * while every cheap arm draws is the call; both LATE arms failing while both
 * early ones draw is position or elapsed time; anything else is no separation,
 * which after eleven rounds of eliminated candidates is the expected answer.
 *
 * Each pair is [early, late], which is what makes the second reading available.
 */
export function rasteriseArmVerdict(raster: DrawArm[], cheap: DrawArm[]): { ok: boolean; detail: string } {
  const drew = (a: DrawArm[]) => a.filter((x) => x.drew).length;
  const why = [...raster, ...cheap].find((a) => !a.drew)?.why ?? "";
  const all = raster.length + cheap.length;
  const landed = drew(raster) + drew(cheap);
  if (landed === all)
    return {
      ok: true,
      detail: "all four draws landed — a draw straight after a rasterise is no worse than one after a cheap read",
    };
  if (landed === 0)
    return { ok: false, detail: `no draw landed at all (${why}) — the slide or the moment, not the call before them` };
  if (drew(raster) === 0 && drew(cheap) === cheap.length)
    return {
      ok: false,
      detail:
        `every draw after a RASTERISE failed (${why}) and every one after a cheap read landed, ` +
        "interleaved so position cannot account for it — the call before the draw is the difference",
    };
  const early = [raster[0], cheap[0]].filter(Boolean);
  const late = [raster[1], cheap[1]].filter(Boolean);
  if (late.length === 2 && drew(late) === 0 && drew(early) === early.length)
    return {
      ok: false,
      detail:
        `both LATER draws failed (${why}) and both earlier ones landed, whichever call preceded them — ` +
        "this is position or elapsed time, not the rasterise",
    };
  // No separation — and that is this control's EXPECTED result, not a failure.
  //
  // It reported `ok: false` here, and the sentence it printed while doing so
  // said "no pattern ... which is what eleven rounds of eliminated candidates
  // already said". A control cannot call its own documented answer a failure:
  // `CLAUDE.md` records that a few rounds of "no pattern" IS the finding and the
  // point at which to stop instrumenting.
  //
  // It also went red on a schedule. The stall this scenario runs alongside is
  // intermittent at roughly one or two draws in fifteen, and a round makes four
  // — so one round in three or four would fail for the known reason, on a
  // verdict whose own words said nothing was wrong. A red that appears on a
  // timer teaches a reader to stop reading reds.
  //
  // The three verdicts above stay `false`, because each of them IS a
  // separation: the call, the position, or a slide refusing everything. What
  // changes is only the case where the arms disagree with each other.
  return {
    ok: true,
    detail:
      `${landed} of ${all} draws landed (${why}), with no pattern in either the call before them or their ` +
      "position — the stall is intermittent, which is what eleven rounds of eliminated candidates already said",
  };
}

/**
 * WHAT A CHART KIND COSTS — the one question this archive is built wrong for.
 *
 * `src/core/insert-cost.ts` prices a draw as `batches x batchMs(occupancy)`,
 * which says every ten-shape batch costs the same. Measured 2026-09-05, they do
 * not: at a true prior occupancy of zero a 16-shape chart's first batch medians
 * 10,098ms against 3,617ms for a 24-shape one — backwards to the count, 2.8x.
 *
 * Every candidate explanation was eliminated against the archive: occupancy
 * (paired inside single draws), batch position, draw path, in-place updates,
 * and the call the host last answered before the draw. What defeated the last
 * one is why this scenario has to exist. Sorting batch-1 readings by that
 * preceding call separates them beautifully — ~3,400ms after a per-slide
 * operation, ~10,300ms after a deck-wide scan — and it is an artefact:
 * 103-shape charts follow `listing the deck's slides` in 23 of 23 readings and
 * are the CHEAPEST batch in the archive at 1,611ms.
 *
 * The reason no split works is structural. Each existing scenario draws ONE
 * chart in ONE context, so chart size, preceding call, scenario and slot are a
 * single variable wearing four names:
 *
 *      16-shape charts    273 of 274 after a deck scan     med 10,046ms
 *      24-shape charts      0 of 1,210 after a deck scan   med  3,615ms
 *     103-shape charts     23 of 23 after a deck listing   med  1,611ms
 *
 * No further reading answers that. It needs an experiment where the KIND is the
 * only thing that moves, which is this one.
 *
 * COUNTERBALANCED, for the reason `rasteriseArmVerdict` spells out at length:
 * the kinds are drawn in order and then in reverse, so each appears once early
 * and once late. Without it, kind would be confounded with position exactly as
 * it is confounded with scenario in the archive today — the same mistake, moved
 * into the instrument built to escape it.
 *
 * THE SLIDE THIS SCENARIO LEAVES IS NOT THE SLIDE IT FOUND, and it never has
 * been. An earlier version of this docstring promised the opposite — "each
 * specimen is deleted before the next, so the slide this scenario leaves is the
 * slide it found" — and then retracted it eight lines later. The promise is
 * retired here rather than repeated: `docs/BACKLOG.md` §"The sweep deleted
 * nothing seven times" set out the two ways forward, and this is the second of
 * them, taken 2026-09-10.
 *
 * WHY THE SWEEP DOES NOT LAND. `deleteShapesById` resolves each stray by id
 * through a slide handle a sync old, which is the one lookup this host refuses
 * for a shape drawn seconds earlier. Re-measured over the 55 archived rounds
 * that drew a specimen: **369 of the 419 sweeps that reported an outcome were
 * refused**, 34 of the 55 rounds had all eight refused, exactly one round took
 * all eight, and 20 fell in between. It is a coin, not a wall. (The figures
 * this paragraph replaced — "28 of 33 rounds" — were right at their own
 * denominator and had simply aged.)
 *
 * SO SAY WHAT IT LEAVES. The slide ends the round holding the probe chart this
 * scenario drew beside, plus one grouped chart per unswept specimen: a median
 * of 8 top-level shapes and 49 in the worst round. That litter is why the prior
 * occupancy the experiment measures at climbs 0, 7, 15, 24, 34, 44, 53, 61
 * across the eight slots; in the run that killed the tab, PowerPoint died on
 * the eighth.
 *
 * NOTHING DOWNSTREAM INHERITS IT, BECAUSE THE DRIVER DISPOSES OF IT. This
 * scenario runs LAST in every round that has carried it, and `sweepDeck`
 * (`scripts/round.mjs`) runs straight after the round is archived, deleting
 * every slide from the last down to index 1. **427 of 429 archived rounds begin
 * at `deckSlides: 1`.** The litter lives for the tail of one round and is gone
 * before the next reads anything.
 *
 * All four routes out are measured and closed — by-id, collection read matched
 * by id, binding retrieval, and name/positional matching. The remaining option,
 * drawing each specimen on a scratch slide deleted whole, is NOT taken and is
 * not a pending decision: a scratch ADD is the call with the worst kill record
 * in this archive — see `chartIsVisible`, which stopped making it after it
 * killed the host five rounds running.
 *
 * THE COUNTER IS RIGHT AND THE SWEEP'S ARITHMETIC IS NOT. Occupancy climbs in
 * the renderer's own counter, which does not decrement on a plain delete —
 * deliberate: `last batch settled` records `onSlideAfter` and `drew`, so the
 * true prior is exact for every reading. When a sweep is REFUSED the counter is
 * not merely conservative, it is CORRECT: the shapes are still there.
 *
 * **BUT WHEN A SWEEP LANDS ON A GROUP, THE COUNTER OVER-DEDUCTS THE OTHER WAY.**
 * `forgetShapesDrawnOn(slideId, gone)` subtracts the number of delete CALLS, and
 * a group delete takes 7-10 shapes in one call — the defect `replacedShapeCount`
 * exists to fix on the UPDATE path, which the sweep path never got. Measured
 * 2026-09-10: of 423 readings, **91 (21.5%) carry a prior the counter got
 * wrong**, and refitting the cost model on the repaired covariate moves the
 * occupancy slope from 198 to 219 ms/shape — about 1.9 standard errors.
 *
 * IT IS INVERTIBLE FROM THE ARCHIVE, so this is a note for the reader and not a
 * change to the instrument. Within one slide, `prior(n+1) === onSlideAfter(n) −
 * resolved(n)` holds in 364 of 368 transitions, and `resolved === 1` following a
 * chart of 7-10 shapes is a group delete whose true bite is that chart's
 * `total` — n=28, every one between 7 and 10. Whoever pools these readings
 * should apply that before fitting. **No reader is written**, deliberately:
 * `onSlideAfter` appears zero times in `triage.mjs`, `rounds-gate.mjs` and
 * `claims.mjs`, so there is no consumer to correct — building one would be an
 * instrument for an analysis nobody runs.
 *
 * THE VERDICT IS NOT THE MEASUREMENT. This returns pass/fail on whether the
 * draws LANDED; the costs live in the trace and are read from the archive. A
 * verdict that failed on a timing would go red on host weather, and this
 * project has already learned what a red on a schedule teaches a reader.
 */
const kindCostSpread: Scenario = async (prefix) => {
  const attempt = stepsOf("kind cost step");
  const { found, blind, gap } = await probeCharts(prefix);
  const host = leastLoadedChart(found, shapesDrawnOn);
  if (!host) return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to draw beside" };
  const slideId = host.target.slideId;
  const slide = await slideSize();
  const cell = sideSlot(GRID_SLOT, slide, boxOf(host));
  const w = Math.max(1, Math.floor(cell.width / 4) - 3);
  /**
   * ONE DATASET FOR EVERY KIND, so the only difference is how it is drawn. Two
   * categories and one series keeps each specimen inside a single batch, which
   * is what makes `last batch settled` the whole draw's cost rather than a
   * tail — the same reasoning as the rasterise arm's `tiny`.
   */
  const specimen = (kind: ChartKind, n: number): ChartConfig => ({
    kind,
    title: `${prefix} kind ${kind} ${n}`,
    width: w,
    height: cell.height,
    data: { categories: ["A", "B"], series: [{ name: "s", values: [1, 2] }] },
  });
  const drawn: KindDraw[] = [];
  for (let n = 0; n < KIND_COST_ORDER.length; n++) {
    const kind = KIND_COST_ORDER[n];
    const c = specimen(kind, n);
    try {
      const t = await attempt(`${kind} #${n}: drawing`, () =>
        insertSceneIntoSlide(buildChart(c), {
          slideId,
          tagData: JSON.stringify(c),
          left: cell.left + (n % 4) * (w + 3),
          top: cell.top,
        }),
      );
      drawn.push({ kind, drew: true, why: "" });
      // Swept immediately, so the deck this scenario leaves is the deck it
      // found. Same shape as the own-slide control's cleanup: group plus parts.
      const ids = [t?.shapeId, ...(t?.partIds ?? [])].filter(Boolean) as string[];
      if (t && ids.length) await deleteShapesById(t.slideId, ids);
    } catch (err) {
      // Per specimen, not per scenario: one kind the host will not draw must
      // not discard the readings of the seven it did.
      drawn.push({
        kind,
        drew: false,
        // A HOST THAT STOPPED ANSWERING IS NOT A STATEMENT ABOUT THE CHART, and
        // the verdict needs to know the difference. See `kindCostVerdict`.
        stalled: isTimeout(err),
        why: isTimeout(err) ? "the host stopped answering" : errorText(err),
      });
    }
  }
  trace("selftest", "chart kind cost arms", {
    order: KIND_COST_ORDER,
    drew: drawn.map((d) => d.drew),
    stalled: drawn.filter((d) => d.stalled).length,
  });
  return kindCostVerdict(drawn);
};

/** One specimen's outcome: did it land, and if not, what the host said. */
export interface KindDraw {
  kind: ChartKind;
  drew: boolean;
  why: string;
  /** The host stopped answering, as opposed to refusing THIS chart. */
  stalled?: boolean;
}

/**
 * The kinds, and then the same kinds backwards.
 *
 * Four chosen to span how a shape is MADE rather than how many there are, since
 * the count is the variable already known not to explain this: `clustered` is
 * plain rectangles and is the cheap population's own kind, so the readings
 * anchor to something; `funnel` is filled trapezoid bands; `waterfall` mixes
 * rectangles with connector lines; `line` is open polyline strokes.
 *
 * THE COUNTS HAVE TO BE CLOSE, AND THE FIRST VERSION'S WERE NOT. Rounds 398 and
 * 399 ran `clustered, line, area, pie` and the design did not survive contact.
 * `pie` draws THIRTY-SEVEN shapes at this size, so it was multi-batch — making
 * `last batch settled` its tail rather than its cost — and it drove the prior
 * occupancy to 60 before the second half began, then 74, 81, 91, where a draw
 * costs three to five times what it does at zero. The palindrome balances
 * POSITION. It can only balance OCCUPANCY when the specimens are the same size.
 *
 * Measured at four candidate boxes (50x60 through 160x120), shapes per kind:
 *
 *     clustered   7  7  9  9        line       10 10 10 10
 *     funnel      8  8  8  8        waterfall   9  9  9  9
 *     area        8 11 15 23        pie        37 37 37 37
 *
 * So `area` is out as well, for a reason the first pick missed: its count
 * varies THREE-FOLD with the box, so it cannot hold occupancy equal either. The
 * four kept are flat across every box and land within three shapes of one
 * another.
 *
 * What that buys, with priors running 0, 7, 15, 24, 34, 44, 53, 61 across the
 * eight slots: a mean prior per kind of 30.5, 30, 29.5, 29 — balanced to within
 * one and a half shapes, which is the property the palindrome is FOR.
 *
 * Exported so a reading of the archive can name the order without re-deriving
 * it from a round.
 */
export const KIND_COST_ORDER: readonly ChartKind[] = [
  "clustered",
  "funnel",
  "waterfall",
  "line",
  "line",
  "waterfall",
  "funnel",
  "clustered",
];

/**
 * Did the specimens land — deliberately NOT whether they differed in cost.
 *
 * The costs are the point of the scenario and are read from the archive, not
 * asserted here. A verdict that failed on a timing would go red on host
 * weather, and `rasteriseArmVerdict` records what that does to a reader.
 *
 * The one thing worth failing on is a kind the host will not draw AT ALL, in
 * either position — that is a product fact rather than weather, and it is
 * invisible in a battery where every other scenario draws `clustered`.
 */
export function kindCostVerdict(drawn: KindDraw[]): { ok: boolean; detail: string; skipped?: boolean } {
  if (!drawn.length) return { ok: false, skipped: true, detail: "no specimen was attempted" };
  const landed = drawn.filter((d) => d.drew).length;
  const kinds = [...new Set(drawn.map((d) => d.kind))];
  const never = (k: ChartKind) => drawn.filter((d) => d.kind === k).every((d) => !d.drew);
  // A KIND IS ONLY "DEAD" IF THE HOST REFUSED IT, NOT IF THE HOST WENT QUIET.
  //
  // This rule was `never(k)` alone, and it fired a false red on its first real
  // outing: 2026-09-05, on a 45-slide deck, `funnel` missed both slots while
  // clustered and waterfall took both, and the verdict called it "a kind this
  // host will not take". Re-run alone on the same deck twenty minutes later:
  // 8 of 8 specimens landed, funnel included. The kind is fine.
  //
  // The arithmetic says that was predictable rather than unlucky. With eight
  // slots in four mirrored pairs, if failures fall at random then SOME kind
  // loses both of its slots 14% of the time on two failures, 43% on three, and
  // 77% on four. That run had three. A red that arrives 43% of the time
  // whenever the host has a bad minute is a red on a timer, and
  // `rasteriseArmVerdict` records what those teach a reader to do.
  //
  // The discriminator is already in hand: `insertSceneIntoSlide` throwing a
  // TIMEOUT is the host not answering, which is not a statement about the
  // chart. A kind the host genuinely will not take refuses — it does not hang.
  // So a kind counts as dead only when at least one of its failures was a
  // refusal rather than a stall.
  const dead = kinds.filter((k) => never(k) && drawn.some((d) => d.kind === k && !d.drew && !d.stalled));
  const stalledOut = kinds.filter((k) => never(k) && !dead.includes(k));
  const why = drawn.find((d) => !d.drew)?.why ?? "";
  if (dead.length)
    return {
      ok: false,
      detail:
        `${dead.join(", ")} was REFUSED in either position (${why}) — a kind this host will not take, ` +
        "which no other scenario would show because they all draw clustered",
    };
  if (stalledOut.length)
    return {
      ok: true,
      detail:
        `${landed} of ${drawn.length} specimens landed; ${stalledOut.join(", ")} missed both slots, but every ` +
        "failure was the host not answering rather than a refusal, so this says nothing about the kind — " +
        "re-run it alone before believing it",
    };
  if (landed < drawn.length)
    return {
      ok: true,
      detail:
        `${landed} of ${drawn.length} specimens landed (${why}); every kind drew at least once, so this is ` +
        "the intermittent stall rather than a kind the host refuses",
    };
  return {
    ok: true,
    detail: `${landed} of ${landed} specimens landed across ${kinds.length} kinds, each drawn once early and once late`,
  };
}

/**
 * Order: the scenarios with a track record first, the new ones last.
 *
 * A battery is only worth its whole run if it survives to the end, and this one
 * has twice now failed to — a wedge at 1819s and a host killed outright at
 * 108s. When that happens every scenario AFTER the failure reports nothing, so
 * where an unproven scenario sits in this list decides how much of the run its
 * failure costs.
 *
 * The two added most recently — selection and stop — therefore run at the end.
 * The heaviest and least proven, `chartIsVisible`, ran dead last for the same
 * reason (it was 6th, and a crash there cost the verdicts of three scenarios
 * that had worked on a real host that morning). Last was not far enough: it
 * killed the tab in four consecutive rounds without ever returning a verdict,
 * and a battery that never returns never writes its report. It is `pickedOnly`
 * now — see its entry. New code must not be able to take the evidence for old
 * code with it, and "last" only buys that up to the point where the crash costs
 * the report itself.
 *
 * The first two remain first because everything else needs the probe charts
 * they insert. Beyond that the constraint is `explodePicture`: it inserts a
 * PICTURE, and on the web a picture cannot be inserted while a shape is
 * selected (office-js#3698) — so every scenario that selects must clear up
 * after itself, which `clearShapeSelection` does, wherever it runs.
 */
/**
 * Consecutive host-trouble scenarios before the battery stops asking.
 *
 * Three, not one: a single failure is usually the finding this whole file
 * exists to produce. Three in a row is not a finding about a scenario, it is a
 * finding about the host — and every real-host run so far has continued for
 * minutes past that point and then died, taking its own verdicts with it.
 */
const SICK_LIMIT = 3;

/**
 * Whether a finished scenario says the HOST is in trouble, as opposed to
 * having found something.
 *
 * Three facts, and `timedOut` is the one that matters: a bounded wait that hit
 * its deadline, a throw, or a deck scan that went blind. A plain FAILED verdict
 * is none of them — that is the battery working. Nor is a capability skip,
 * which says nothing about how the host is feeling.
 *
 * `timedOut` is passed IN, measured by diffing `deadlinesFired` around the
 * scenario, because the previous version read it out of the verdict's prose:
 * `/did not respond|gave up/` over `result.detail`. A real run walked straight
 * through that. Two scenarios timed out and used those words; a third timed out
 * after 49.8 seconds and reported *"the host stopped answering selection calls
 * after a programmatic select"* — true, specific, and matching neither phrase.
 * The counter reset one short of the limit, the battery spent two more
 * scenarios on a host that had been dead for three, and the tab died with the
 * remaining verdicts inside it.
 *
 * Adding that third phrase to the pattern would have been the same bug with a
 * longer list. A scenario chooses its own words; it does not choose whether a
 * deadline fired.
 *
 * A SKIP counts when a deadline fired inside it — that third scenario reported
 * skipped, and it was the sickest of the three: it had waited out the whole
 * budget to say so.
 */
export function hostSeemsSick(result: ScenarioResult, timedOut: boolean): boolean {
  return timedOut || result.blind === true || result.detail.startsWith("threw:");
}

/** How long the battery waits for a person to click a chart. */
let CLICK_WAIT_MS = 30_000;

/** Test-only: a suite cannot spend thirty seconds waiting for nobody. */
export function _setClickWaitForTest(ms: number): void {
  CLICK_WAIT_MS = ms;
}

/**
 * The one path a real user travels that nothing can script.
 *
 * `setSelectedShapes` is Office.js selecting a shape. A human clicking one is
 * the same call in theory, and on PowerPoint on the web demonstrably not the
 * same in practice — the programmatic version is taken and then the selection
 * subsystem stops answering, which is why `editViaSelection` reports *skipped*
 * there. That leaves the pane's most-used read — the one behind "Edit it" —
 * with no coverage at all on the host where all the bugs are.
 *
 * So ask a person. One click, and then the whole chain runs and is measured:
 * read the selection back, confirm it is the chart that was clicked, edit
 * through the target THAT read produced, and confirm the result is still
 * re-editable.
 *
 * That is not a workaround for being unable to script it. It is the more
 * faithful test, and this project has the measurement to say so: the scripted
 * version does not behave like the real one on the host that matters.
 *
 * Listens via `DocumentSelectionChanged` — a Common API event that does not go
 * through the wedging subsystem, and the same one the pane's own selection
 * banner has always used. Nothing here calls `setSelectedShapes`.
 *
 * Picked only, for the obvious reason: it blocks on a human.
 */
const editViaRealClick: Scenario = async () => {
  if (!canWatchSelection()) {
    return { ok: false, skipped: true, detail: "host does not raise DocumentSelectionChanged" };
  }
  prompt?.("Click an SSF chart on a slide now — the self-test is waiting for it.");
  trace("selftest", "WAITING FOR YOU: click a chart on the slide", { seconds: CLICK_WAIT_MS / 1000 });
  const {
    chart: picked,
    sawClick,
    readFailed,
  } = await awaitSelectedChart(CLICK_WAIT_MS, selectionBudgetMs(), (left) => {
    prompt?.(`Click an SSF chart on a slide — ${left}s left.`);
    trace("selftest", "still waiting for a click", { secondsLeft: left });
  });
  if (!picked) {
    // Two different findings, and only one of them is about the person. A
    // click the host would not describe is a HOST result and belongs in the
    // report as one; calling it "nobody clicked" blames the reader for a
    // failure they can see they did not cause, which is how a report stops
    // being trusted.
    if (readFailed) {
      return {
        ok: false,
        skipped: true,
        detail: "you clicked, and the host would not say what was selected — the selection read never came back",
      };
    }
    return {
      ok: false,
      skipped: true,
      detail: sawClick
        ? "you clicked, but what you clicked is not an SSF chart — nothing was checked"
        : `nobody clicked an SSF chart within ${CLICK_WAIT_MS / 1000}s — nothing was checked`,
    };
  }
  prompt?.("Got it — editing the chart you clicked.");
  const was = JSON.parse(picked.configJson) as ChartConfig;
  const title = typeof was.title === "string" ? was.title : "(untitled)";
  // The read is the half this scenario exists for, and it has just succeeded —
  // so it is stated first and unconditionally. Everything after it is the
  // ordinary edit path, which other scenarios already cover; what none of them
  // can cover is that the config came back from a HUMAN's click.
  const read = `read "${title}" back from a real click`;
  // Edit through the target the SELECTION produced. That target is built by
  // different code, from a different read, than the one the deck scan makes —
  // and only this one is what a user's edit actually travels on.
  const next = { ...was, title: `${title} (via a real click)` };
  const target = await updateChartInSlide(buildChart(next), picked.target, { tagData: JSON.stringify(next) });
  if (!target) return { ok: false, detail: `${read} — and the host would not edit through that target` };
  const round = await loadChartFromSelection(selectionBudgetMs()).catch(() => null);
  return {
    ok: true,
    detail:
      `${read}, edited through the selection's own target, and it is ` +
      (round?.configJson ? "still re-editable" : "no longer readable from the selection (the view may have moved)"),
  };
};

/**
 * Which selection call wedges the host — asked once, properly.
 *
 * Two accounts, different culprits. This project measured
 * `setSelectedShapes([id])` being taken and every selection call after it going
 * silent; office-js#3698 says it is `setSelectedShapes([])` whose promise never
 * resolves. Both are plausible, neither is verified, and the gate above was
 * written against my reading of one 159-second run.
 *
 * `selectionLadder` climbs from the least invasive call to the most and stops
 * at the first silence, because after a wedge every call is silent and a report
 * of four timeouts names nothing. What comes back is one sentence: the last
 * call the host answered, and the first one it did not.
 *
 * Reported `ok` whenever the ladder RAN, whatever it found. It is an
 * experiment, not an assertion — a host that answers every rung is a genuine
 * result (the wedge is gone, or is not on this host), and marking that as a
 * failure would be the same mistake as marking the wedge itself as one.
 */
/**
 * What the ladder found, in a form the scenario after it can act on.
 *
 * A sentence in a report is for a person. This is for `editViaSelection`, which
 * runs next and would otherwise spend a budget re-discovering it. Reset per run,
 * because a stale answer from the last round is worse than none.
 */
let selectionWedged: string | null = null;

const whichSelectionCallWedges: Scenario = async (prefix) => {
  if (!canSelectShapes()) return { ok: false, skipped: true, detail: "host cannot select shapes (PowerPointApi 1.5)" };
  const { found, blind, gap } = await probeCharts(prefix);
  const chart = leastLoadedChart(found, shapesDrawnOn);
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to select" };
  const rungs = await selectionLadder(chart.target.slideId, chart.target.shapeId, selectionBudgetMs());
  for (const r of rungs) trace("selftest", "ladder rung", { step: r.step, outcome: r.outcome, ms: r.ms });
  const stopped = rungs.find((r) => r.outcome === "silent");
  const answered = rungs.filter((r) => r.outcome === "ok").length;
  if (!stopped) {
    const refused = rungs.filter((r) => r.outcome === "refused").map((r) => r.step);
    return {
      ok: true,
      detail:
        `the host answered all ${rungs.length} rung(s) — nothing wedged` +
        (refused.length ? `; refused (but kept talking): ${refused.join(", ")}` : ""),
    };
  }
  // The rung before the silence is the last thing the host was willing to do,
  // and naming both is the whole point: "silent at X, after Y" is an answer,
  // "something in the selection API hangs" is what we already had.
  const before = rungs[rungs.indexOf(stopped) - 1];
  selectionWedged = `goes silent at "${stopped.step}"`;
  return {
    ok: true,
    detail:
      `SILENT at "${stopped.step}" after ${stopped.ms}ms` +
      (before ? `, the rung before it ("${before.step}") answered in ${before.ms}ms` : ", the very first rung") +
      ` — ${answered} rung(s) answered before it`,
  };
};

/**
 * Does drawing a chart destroy a shape the user had selected?
 *
 * office-js#2775, reported against PowerPoint on the web and still open: adding
 * a text box deletes whatever shape was selected, and the reporter notes it
 * "works fine on desktop". Every chart this add-in draws contains text boxes,
 * and the insert path deliberately leaves the user's selection alone because
 * that is how it learns where to put the chart — so if the bug is live on the
 * owner's build, selecting a picture and inserting a chart against it destroys
 * the picture. Silently, on the everyday path.
 *
 * `dropShapeSelection` is the guard, and it went in without waiting for an
 * answer because the cost is one selection call and the failure is a user
 * losing their own content. This is the question that says whether that guard
 * is load-bearing on THIS host or merely cheap insurance — and it is worth
 * knowing, because a host that does this will do it on every other add-in the
 * owner uses too.
 *
 * Asked on a scratch slide with a shape of our own, never on anything of the
 * user's. Placed after the ladder for the reason the ladder's own comment
 * gives: it calls `setSelectedShapes`, so it must not be the first thing in the
 * run that does.
 */
const selectionSurvivesInsert: Scenario = async (prefix) => {
  if (!canSelectShapes()) return { ok: false, skipped: true, detail: "host cannot select shapes (PowerPointApi 1.5)" };
  if (selectionWedged) {
    return { ok: false, skipped: true, detail: `not attempted — the ladder found this host ${selectionWedged}` };
  }
  // A probe chart on a settled slide, never a fresh scratch slide.
  //
  // The first draft added its own slide and selected a shape on it, and the
  // fake refused: `selectShape` resolves the slide, syncs, then reaches through
  // that same handle — and a freshly-added slide's handle is good for exactly
  // one sync. That is this repo's oldest gotcha biting a new caller, and the
  // scenario is better for the correction anyway: a shape on a slide the user
  // has had for a while is the situation #2775 actually describes.
  const { found, blind, gap } = await probeCharts(prefix);
  const [victim] = found;
  if (!victim) return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to select" };
  const slideId = victim.target.slideId;
  const inventory = async () =>
    (await listChartsInDeck({ withInventory: true })).inventory?.find((s) => s.slideId === slideId);
  const before = await inventory();
  if (!before) return { ok: false, skipped: true, detail: "the host would not say what is on the slide" };
  if (!(await selectShape(slideId, victim.target.shapeId, selectionBudgetMs()))) {
    return { ok: false, skipped: true, detail: "the host would not select the chart" };
  }
  try {
    // And now draw, with that shape still selected. Text boxes and all — which
    // is exactly #2775's repro, and exactly what the pane did on every insert
    // before `dropShapeSelection`.
    // Slot 3 — the last of the four. See `sideSlot`: three other charts share
    // this slide, and drawing full-size on top of them is what left the owner a
    // heap to look at.
    const { left, top, ...box } = sideSlot(3, await slideSize(), boxOf(victim));
    const extra: ChartConfig = { ...cfg(`${prefix} drawn while selected`), ...box };
    await insertSceneIntoSlide(buildChart(extra), { slideId, tagData: JSON.stringify(extra), left, top });
  } catch (err) {
    // Named, rather than left to the runner's generic "the host got in the way".
    //
    // This draw stalls FAR more than any other in the battery, but only
    // RECENTLY — and an earlier version of this note implied a standing
    // property of the host, which the full history does not support.
    //
    // Measured across all seventeen rounds on file, against `edit the chart the
    // user selected` — which selects a shape, DROPS the selection, then draws:
    //
    //     all 17 rounds        held 10 ok / 7 stalled   control 14 ok / 3 not ok
    //     since `957aca0` (8)  held  2 ok / 6 stalled   control  8 ok / 0
    //
    // THOSE ROUND COUNTS ARE HISTORY. Re-measured 2026-08-25 across all 235
    // rounds on file, counting DRAWS rather than rounds — `batch issued` in
    // draw scope, the only per-draw line that fires on success:
    //
    //     held selection      10 stalled / 457 drawn   2.19%
    //     dropped (control)    0 stalled / 516 drawn   0%
    //
    // They separate, and the sentence this note used to carry — "over the whole
    // history they do not separate" — was true of seventeen rounds and is not
    // true of two hundred. ZERO in 516 draws is the number that matters: with
    // the selection dropped this draw has never once stalled.
    //
    // So `dropShapeSelection` is not a precaution, it is the fix, and the
    // archive now says so. That also settles the older speculation below about
    // something changing at `957aca0`: the effect is visible across the whole
    // archive once draws are counted instead of rounds, so there is no era to
    // explain.
    //
    // Every one of those rounds reported the stall the same anonymous way — so the battery has been carrying a repeating, specific
    // observation and saying nothing about it.
    //
    // What the note may claim is bounded by what the round files support. The
    // preceding call is NOT it: every one of those stalls reads
    // `afterAnswering: "selecting a shape"`, and `selecting a shape` sits in
    // the SURVIVING population in all four of the same rounds, because `edit
    // the chart the user selected` draws after it and lands. What is left, and
    // all this says, is the one way this draw differs from every other draw in
    // the battery — it is the only one made with a selection still standing,
    // which is #2775's repro and what `dropShapeSelection` exists to avoid.
    //
    // A DEDICATED control arm was built to settle it and then removed: matching
    // this draw needs a same-size chart on the same slide, every slot is
    // allocated, and widening the band broke both the no-overlap invariant and
    // the degradation grid's fit. That decision still stands, and it turns out
    // not to have cost anything: `edit the chart the user selected` selects a
    // shape, DROPS the selection and then draws, which is the comparison that
    // was wanted, and 235 rounds of it now answer the question without
    // deforming the battery at all.
    //
    // The product still does not turn on it — the add-in already drops the
    // selection — so this remains an observation. It is now an observation with
    // 973 draws behind it rather than seventeen rounds.
    if (!isTimeout(err)) throw err;
    return {
      ok: false,
      skipped: true,
      blind: true,
      detail:
        `the host did not finish a draw made while a shape was SELECTED — the only draw in this battery made ` +
        `with a selection standing, and what dropShapeSelection exists to avoid: ${errorText(err)}`,
      ms: 0,
    };
  } finally {
    // Never leave a chart selected: on the web a picture cannot be inserted
    // while one is (office-js#3698), so the scenario after this would fail
    // instead of this one.
    await clearShapeSelection(slideId, selectionBudgetMs());
  }
  const after = await inventory();
  if (!after) return { ok: false, skipped: true, detail: "the host would not say what is on the slide afterwards" };
  // Both readings have to be WHOLE, and this scenario was the only one in the
  // file that could report `ok: true` off a blinded scan.
  //
  // `SlideInventory.count` is the slide's own count and exists for exactly this
  // corroboration — a short collection read is routine on the web, and this
  // repo has recorded a readback page asking about 19 shapes and getting 3.
  // Unchecked it goes wrong in both directions from one omission: a short
  // BEFORE empties the vanished-shape filter, so the scenario reports GREEN
  // ("all 0 shape(s) … survived") off a read that saw nothing; a short AFTER
  // makes every shape look deleted and the scenario declares office-js#2775
  // live on this host — the loudest claim it can make, from a host that merely
  // answered short.
  const partial = (inv: { shapes: unknown[]; count?: number }) =>
    typeof inv.count === "number" && inv.shapes.length < inv.count;
  if (partial(before) || partial(after)) {
    return blindSkip(
      `the slide's shape list came back short — ${before.shapes.length}/${before.count ?? "?"} before, ` +
        `${after.shapes.length}/${after.count ?? "?"} after`,
    );
  }
  const kept = new Set(after.shapes.map((s) => s.id));
  const lost = before.shapes.filter((s) => s.id && !kept.has(s.id));
  return {
    ok: lost.length === 0,
    detail:
      lost.length === 0
        ? `all ${before.shapes.length} shape(s) already on the slide survived an insert made while one was selected`
        : `${lost.length} of ${before.shapes.length} shape(s) VANISHED when a chart was drawn with one selected ` +
          `— office-js#2775 is live on this host, and dropShapeSelection is what stands between it and the user`,
  };
};

/** Fewer rounds than this and the slope is noise, not a curve. */
const MIN_ROUNDS_FOR_A_VERDICT = 4;

/**
 * How much a series' last third cost more than its first third, in ms.
 *
 * Thirds rather than last-minus-first, because one slow round is not a trend
 * and a two-sample comparison cannot tell the two apart. `null` when there is
 * not enough of a curve to say anything, which the caller reports rather than
 * rounding to "no change" — an unmeasured slope and a flat one are different
 * answers, and this file exists because they kept getting confused.
 *
 * **Median of each third, not the mean.** A third is three samples, and one
 * garbage collection is enough to move a mean of three. Run against a fake host
 * with no slowdown at all, the mean version read
 * `[6, 6, 6, 6, 10, 7, 17, 6]ms` — one 17ms hiccup — as "THE CONTEXT degrades,
 * +67%". A confident wrong answer from noise is the worst thing this experiment
 * could produce, because the whole reason it exists is that the cheap answers
 * were untrustworthy.
 */
function growth(ms: number[]): { head: number; delta: number } | null {
  if (ms.length < MIN_ROUNDS_FOR_A_VERDICT) return null;
  const k = Math.ceil(ms.length / 3);
  const median = (xs: number[]) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
  const head = median(ms.slice(0, k));
  return { head, delta: median(ms.slice(-k)) - head };
}

/**
 * How much growth counts as degradation, as a fraction of the reference round.
 *
 * Half again over the measured span, and chosen rather than derived — nothing
 * here has the sample size to justify a real threshold, and pretending
 * otherwise would be worse than saying so. It sits well above the round-to-round
 * noise of a healthy host and far below the tenfold blowups the real artefacts
 * show, which is all it has to separate.
 */
const GREW = 0.5;

/**
 * And how much growth counts in plain milliseconds, whatever the fraction says.
 *
 * The relative test alone is unsafe at the bottom of the scale: against a 6ms
 * round, 3ms of scheduler noise is "+50%". Both bars have to clear, so the
 * verdict needs a trend that is both proportionally large AND big enough to
 * matter — and 25ms is far below the hundreds of milliseconds every real
 * degraded run shows, while being comfortably above what a busy machine adds to
 * a round on its own.
 */
const GREW_MS = 25;

export interface DegradationVerdict {
  /** Growth across the long-context arm, as a fraction of the reference round. */
  oneContext: number;
  /** The same, for the arm that took a fresh context per round. */
  freshContext: number;
  suspect: "context" | "host" | "both" | "none" | "unknown";
  /** The sentence the report leads with. */
  headline: string;
}

/**
 * Read the two curves.
 *
 * Both arms grow the deck by the same amount in the same order, and both age
 * the tab. The ONLY thing that differs is whether the context is re-created
 * between rounds, so:
 *
 * - only the long-context arm grows → the request context is what degrades, and
 *   the fix is a shape of code we already know how to write (shorter contexts).
 * - the fresh-context arm grows too → the host itself is slowing as the deck
 *   grows or the tab ages, and shortening contexts will not help.
 * - both grow, the long one much harder → both, and the first is still worth
 *   fixing.
 * - neither grows → whatever kills a long run is not in this loop, which rules
 *   out three suspects at once and is the most useful answer of the four.
 *
 * **Growth in milliseconds against one shared reference, never each arm's own
 * ratio.** The first version of this compared tail/head within each arm, and it
 * gets the commonest case exactly wrong: under a host that slows linearly the
 * arms are equally affected in ms, but the second arm starts from a higher
 * baseline, so its RATIO is smaller. Worked through with the real numbers the
 * long arm read ×2.64 and the fresh arm ×1.47 for a slowdown neither arm caused
 * — a confident "the context did it" about a host. Both deltas divided by one
 * common reference are order-blind, and that case comes out equal.
 *
 * Residual limit, stated because the raw curves are in the report for exactly
 * this: the arms run one after the other, so a host whose cost grows STRONGLY
 * non-linearly is not fully cancelled — a logarithmic trend still hands the
 * first arm the steeper stretch. A linear trend, which is what a deck growing
 * by a fixed number of shapes per round should produce, cancels exactly.
 */
export function readDegradation(oneContextMs: number[], freshContextMs: number[]): DegradationVerdict {
  const one = growth(oneContextMs);
  const fresh = growth(freshContextMs);
  if (!one || !fresh) {
    return {
      oneContext: NaN,
      freshContext: NaN,
      suspect: "unknown",
      headline: `not enough rounds to read a slope (${oneContextMs.length} in one context, ${freshContextMs.length} in fresh ones)`,
    };
  }
  // The earliest, cheapest measurement in the experiment, and the same divisor
  // for both arms — that sameness is the whole point. Floored at a millisecond
  // because a round the clock reads as free is below the instrument, not free:
  // dividing by it would turn timer jitter into an infinite slowdown.
  const ref = Math.max(one.head, 1);
  const oneGrew = one.delta / ref;
  const freshGrew = fresh.delta / ref;
  const pct = (g: number) => `${g >= 0 ? "+" : ""}${Math.round(g * 100)}%`;
  const both = `one context ${pct(oneGrew)}, fresh contexts ${pct(freshGrew)}, against a ${Math.round(one.head)}ms round`;
  const verdict = (suspect: DegradationVerdict["suspect"], headline: string) => ({
    oneContext: oneGrew,
    freshContext: freshGrew,
    suspect,
    headline,
  });
  // Both bars, never either alone — see GREW_MS.
  const rose = (g: number, deltaMs: number) => g >= GREW && deltaMs >= GREW_MS;
  const oneRose = rose(oneGrew, one.delta);
  const freshRose = rose(freshGrew, fresh.delta);
  if (oneRose && !freshRose) return verdict("context", `THE CONTEXT degrades — ${both}`);
  if (oneRose && freshRose) {
    return oneGrew >= freshGrew * 1.5
      ? verdict("both", `THE HOST is slowing and the context makes it worse — ${both}`)
      : verdict("host", `THE HOST is slowing whatever the context does — ${both}`);
  }
  if (freshRose) {
    return verdict("host", `THE HOST is slowing, and holding a context open did not add to it — ${both}`);
  }
  // Phrased as a statement about the THRESHOLD, not about reality. "Neither
  // curve grew" is simply false in front of two curves that each climbed a
  // tenth, which is a thing this has already printed — and a summary that
  // contradicts the numbers under it is how a reader learns to stop reading
  // either.
  return verdict("none", `no degradation worth the name — ${both}`);
}

/** How big the degradation experiment is. Shrunk by tests, which need the shape and not the minutes. */
let degradeRounds = 8;
let degradePerRound = 12;

/** Test-only: run the experiment at a size a suite can afford. */
export function _setDegradeSizeForTest(rounds: number, perRound: number): void {
  degradeRounds = rounds;
  degradePerRound = perRound;
}

/**
 * Does a long run get slower because of the CONTEXT, the DECK, or the TAB?
 *
 * Every real-host artefact this project owns degrades — 496s to reach scenario
 * seven, a 38-item run whose later inserts cost multiples of its first, a tab
 * PowerPoint eventually killed — and not one of them can say which of the three
 * it was, because all three grow together in an ordinary run. Reasoning about it
 * has already cost this project two sessions; the rule the repo adopted from
 * that is to ask a question that separates the answers instead, which is what
 * this is.
 *
 * Two arms, each drawing the same shapes onto a slide of its own, differing in
 * one thing: whether the request context is re-created between rounds. Separate
 * slides on purpose — sharing one would make the second arm draw onto a slide
 * already holding the first arm's shapes, and a fat slide is a fourth variable.
 *
 * Picked only, and for both of the reasons that word exists here. It costs
 * `2 × rounds` syncs and a few hundred shapes, which is not something to put in
 * front of every routine run; and its subject is a host that has been fatigued,
 * so running it after eight other scenarios would measure them instead of it.
 */
const degradesOverTime: Scenario = async (prefix) => {
  // The steps answered in one round, which is what they were added for.
  //
  // This scenario killed the tab twice. The first time it announced itself at
  // 26.9s and wrote nothing else, so "taking a scratch slide" and "drawing
  // ninety-six shapes onto one just added" both fit. Wrapped, the second round
  // said:
  //
  //     33.2s  degradation step  what=adding the first scratch slide
  //     33.6s  degradation step  what=adding the second scratch slide
  //                                                     <- tab died
  //
  // The first add survived — the second step only gets written if it did — and
  // nothing after the second was ever reached. So the drawing is exonerated,
  // and it is the SECOND `slides.add()`, four tenths of a second after the
  // first, that this host does not survive. No reasoning was involved and none
  // was needed.
  //
  // So it takes no scratch slides. Two of the slides this run already added are
  // enough: the experiment needs two SEPARATE surfaces so neither arm draws on
  // top of the other's shapes, and it needs them EQUALLY loaded so the load is
  // not a fourth variable. Two probe-chart slides are both — one chart each, by
  // construction. Their absolute times will be higher than a bare slide's; the
  // curves are what this measures, and a constant offset does not bend a curve.
  const attempt = stepsOf("degradation step");
  const rounds = degradeRounds;
  const perRound = degradePerRound;
  const { found, blind, gap } = await probeCharts(prefix);
  // Keyed by slide rather than deduplicated separately, so every surface comes
  // with the chart that made it one — the grid below needs both and a second
  // lookup could miss.
  const surfaceBox = new Map<string, Box>();
  for (const c of found) if (!surfaceBox.has(c.target.slideId)) surfaceBox.set(c.target.slideId, boxOf(c));
  const surfaces = [...surfaceBox.keys()];
  const [slideA, slideB] = surfaces;
  if (!slideA || !slideB) {
    if (blind) return blindSkip(gap);
    return {
      ok: false,
      skipped: true,
      detail: `the run has only ${surfaces.length} slide(s) of its own to draw on, and the experiment needs two`,
    };
  }
  const deckBefore = await attempt("counting the deck", slideCount);
  // Each arm's grid goes in the grid slot of ITS OWN slide, measured from the
  // chart that slide actually carries. The arms are on different slides and
  // those slides' charts need not sit at the same origin, so one shared box
  // would be the guess this whole helper exists to remove.
  const slide = await slideSize();
  const gridOn = (box: Box) => {
    const { left, top } = sideSlot(GRID_SLOT, slide, box);
    return { left, top };
  };
  const one = await attempt("timing the one-context arm", () =>
    timeShapeRounds(slideA, {
      rounds,
      perRound,
      oneContext: true,
      label: `${prefix} one-context`,
      origin: gridOn(surfaceBox.get(slideA)!),
    }),
  );
  const fresh = await attempt("timing the fresh-context arm", () =>
    timeShapeRounds(slideB, {
      rounds,
      perRound,
      oneContext: false,
      label: `${prefix} fresh-context`,
      origin: gridOn(surfaceBox.get(slideB)!),
    }),
  );
  const oneMs = one.rounds.map((r) => r.ms);
  const freshMs = fresh.rounds.map((r) => r.ms);
  const verdict = readDegradation(oneMs, freshMs);
  trace("selftest", "degradation curves", {
    oneContext: oneMs,
    freshContext: freshMs,
    suspect: verdict.suspect,
    deckBefore,
  });
  // The raw curves go in the detail, not just the verdict. A ratio is a
  // conclusion and this file has been wrong about conclusions before; the
  // numbers are what someone can re-read when the threshold above turns out to
  // be the wrong one.
  const curves =
    `one context: [${oneMs.join(", ")}]ms` +
    (one.cutShort ? ` (cut short — ${one.cutShort})` : "") +
    `; fresh contexts: [${freshMs.join(", ")}]ms` +
    (fresh.cutShort ? ` (cut short — ${fresh.cutShort})` : "");
  return {
    ok: verdict.suspect === "none",
    // "unknown" means the host did not let the experiment run far enough, which
    // is a skip — the same distinction every other scenario here draws between
    // "we did not check" and "we checked and it is bad".
    ...(verdict.suspect === "unknown" ? { skipped: true } : {}),
    detail:
      `${verdict.headline} — ${rounds}×${perRound} shapes per arm, ` +
      `deck was ${deckBefore} slide(s) at the start. ${curves}`,
  };
};

const SCENARIOS: {
  name: string;
  run: Scenario;
  /**
   * Left out of a full run, offered by the picker.
   *
   * For a scenario whose result is only meaningful on a host nothing else has
   * touched yet. Running it alongside the rest does not merely cost time — it
   * produces a WRONG answer, which is worse than no answer.
   */
  pickedOnly?: boolean;
}[] = [
  { name: "insert on top of an earlier run", run: insertTwice },
  { name: "two slides claiming one slot", run: duplicateSlot },
  { name: "edit a chart on the visible slide", run: editOnVisibleSlide },
  { name: "insert onto a slide that already has content", run: insertOntoUsedSlide },
  { name: "same scale across the deck", run: sameScaleAcrossDeck },
  // AFTER the deck-wide rescale, so the deck it measures is as warm as this
  // harness gets — see `oneChartAlone`.
  { name: "one chart alone on a warm deck", run: oneChartAlone },
  // AFTER the scenarios that insert, because it needs a chart to move, and
  // BEFORE the selection ladder, because it must not run against a wedged host.
  { name: "an update follows a moved chart", run: dragThenUpdate },
  { name: "explode a degraded picture", run: explodePicture },
  // The ladder, and the position is the whole argument.
  //
  // It used to be picked-only, on the grounds that it must be the ONLY thing
  // that has touched the selection subsystem — a run of its own, not a position
  // in a list. That was too strong, and it cost the owner a separate five-minute
  // round every time.
  //
  // The property that actually matters is narrower: **the ladder has to be the
  // first `setSelectedShapes` in the run.** That is the call this project
  // measured wedging and the one office-js#3698 names; `setSelectedSlides`,
  // which three earlier scenarios call through `showSlide`, has never wedged
  // anything and the ladder's own second rung is there to notice if it starts.
  // Sitting immediately before `editViaSelection` gives exactly that, because
  // `editViaSelection` is the only other caller.
  //
  // Position 3 was considered and is worse than either. The ladder can wedge the
  // host, so putting it early means six scenarios run against a wedged host
  // instead of two — strictly less coverage than today, dressed up as a saving.
  { name: "which selection call wedges the host", run: whichSelectionCallWedges },
  // Also after the ladder, and for the ladder's own reason: it calls
  // `setSelectedShapes`, so it must not be the first thing in the run that does.
  { name: "a selected shape survives an insert", run: selectionSurvivesInsert },
  { name: "edit the chart the user selected", run: editViaSelection },
  { name: "stop a run part-way", run: stopPartWay },
  // AFTER it, deliberately. This one aborts a half-drawn chart and cleans up the
  // wreckage; running it before the cheaper before-the-first-batch case would
  // put the harder of the two first for no reason.
  /**
   * BEFORE the mid-draw stop, not after it, and the first attempt had it the
   * wrong way round.
   *
   * The two differ in exactly one thing and a reader comparing their verdicts
   * is the point — but `stop a run mid-draw` does not merely FAIL, it takes the
   * round with it often enough to matter: the round on build a69c7fd banked 12
   * verdicts and died at scenario 13, which is that one. Placed after it, this
   * scenario never got a turn, so the question it exists to answer went
   * unanswered by the round run to answer it.
   *
   * A scenario that cannot run because an earlier one killed the host is a
   * scenario that reports nothing, and ordering is the whole remedy. A NEW
   * name, so it opens its own series and disturbs no round of history —
   * `scenarioRegressions` compares by name.
   */
  { name: "a big chart on a slide of its own", run: bigChartOnItsOwnSlide },
  { name: "stop a run mid-draw", run: stopMidDraw },
  // Picked only because it is 0 for 4 and takes the tab with it every time.
  //
  // Four real-host rounds, four different builds (a5b032d, 618b8d8, cedbc6c,
  // cacf58a), and every one of them ends inside this scenario: at 602s, 631s,
  // 603s and 645s, always within a step or two of `adding a scratch slide`. It
  // has never once returned a verdict. Run 7 got as far as a 90-second rasterise
  // timeout; run 8 was handed a rasterise that "took the call and returned
  // nothing"; runs 6 and 9 simply stopped writing.
  //
  // Running last already protects the other scenarios' verdicts — that is why it
  // was put there — but it does not protect the ROUND. The battery's report is
  // written when the battery returns, and it has not returned yet, so every
  // round so far has reached the owner as a crash file and a screenshot rather
  // than a report. A scenario with no verdicts and a perfect record of killing
  // the tab is not providing coverage; the routine list is the wrong place for
  // it.
  //
  // Picked, it also became the experiment nobody had run. Every crash so far
  // happened around the ten-minute mark with nine scenarios' worth of load
  // behind it, so "this scenario kills the host" and "ten minutes of drawing
  // kills the host, and this is merely what was running" both fit — the same
  // two-readings trap this project keeps paying for. Running it ALONE, on a
  // fresh host, separates them in one short round.
  //
  // IT RAN, 2026-08-08, and it settled: the scenario. Picked alone it was
  // reached at 61.5 seconds with only its two inserts in front of it, took a
  // scratch slide at 61.5s, logged `rasterising the empty slide` at 61.8s, and
  // PowerPoint died. Elapsed time and volume of drawing are both out — the same
  // two inserts run at the head of every routine round and kill nothing. The
  // scenario no longer makes that call; see `chartIsVisible`.
  //
  // AND THE RE-RUN CAME BACK, 2026-08-08, on `e49cca8`: no crash. It rasterised
  // a pre-existing slide, drew its 24 shapes, rasterised again, and returned a
  // verdict — the first this scenario has ever produced in six rounds. So the
  // last reading is closed too: this host rasterises perfectly well, and it was
  // the FRESH slide all along, both directions now measured.
  //
  // The verdict it returned was wrong, and that was ours: the host refused to
  // name the group afterwards, `insertSceneIntoSlide` handed back null, and the
  // scenario called twenty-four committed shapes "nothing was drawn". See
  // `visibilityVerdict`.
  //
  // AND ON `c7d91d5` IT PASSED — routine again as of that round, which is the
  // criterion that was set for it in advance: comes back PASSING, not merely
  // comes back. `drawing the chart changed what the slide looks like (10064 →
  // 15652 bytes)`, on a real PowerPoint, through the host's own rasteriser.
  // That is the first time this project has confirmed a chart it drew is
  // VISIBLE anywhere but in a human's eyes, and it took six rounds to get.
  //
  // It carries a caveat rather than a clean pass — the host would not name the
  // chart afterwards, so it landed without a config — and that is the right
  // shape: a different defect, reported next to the verdict instead of
  // swallowing it.
  //
  // What its absence cost while it was picked-only: `npm run visible-charts`
  // rasterises every sample in a real browser on every CI run and fails on a
  // chart that is drawn but invisible, so the SVG side was never uncovered.
  // What was missing was the check against PowerPoint's own rasteriser, and
  // that is what comes back here.
  { name: "the chart is actually visible", run: chartIsVisible },
  // Immediately after it, and deliberately: `chartIsVisible` is the only other
  // draw in the battery that follows a rasterise, and this is its control. Last
  // in the routine list because it is the newest, which is the rule the comment
  // at the head of this array states.
  { name: "does a rasterise poison the next draw", run: rasteriseThenDraw },
  // LAST of the routine list, and only because its shapes share a slot with the
  // arm above. It reads its own segments by name, so an overlap costs it
  // nothing — but running after everything else means it cannot cost anyone
  // ELSE a readback either.
  { name: "a chart of rotated shapes", run: rotatedShapePlacement },
  /**
   * LAST of the routine list, and a NEW NAME so it opens its own series
   * without disturbing a round of history — `scenarioRegressions` compares by
   * name.
   *
   * Last for the ordinary reason, that it is the newest. But it is also the
   * cheapest thing in the battery: ONE host call that hands over a file,
   * against the eleven-batch draw its sibling makes. Putting it at the end
   * costs the scenarios in front of it nothing.
   */
  { name: "a chart handed over as a file", run: chartAsItsOwnFile },
  // Newest, so last: it draws eight specimens and this list's rule is that an
  // unproven scenario's failure should cost the fewest verdicts behind it.
  { name: "what a chart kind costs", run: kindCostSpread },
  // Picked only for the plainest reason there is: it blocks on a human.
  { name: "edit the chart YOU click", run: editViaRealClick, pickedOnly: true },
  { name: "what makes a long run slow down", run: degradesOverTime, pickedOnly: true },
];

/**
 * Run every scenario, in order, and report all of them.
 *
 * A scenario that throws is a FAILED scenario, not a failed run: the four
 * after it are exactly the ones nobody has data for, and stopping at the first
 * error would spend a real-host session to learn one thing. Each verdict is
 * traced as it lands, so a run that takes the tab down still leaves a record
 * of how far it got.
 */
/**
 * How the battery gets a picture, when it needs one.
 *
 * Rasterising needs a canvas, which lives in the pane — the same split
 * `insertDemoDeck` already uses for its degrade-to-picture path. Absent, the
 * picture scenario reports SKIPPED rather than pretending.
 */
let rasterizer: ((scene: Scene) => Promise<string | undefined>) | undefined;

/**
 * How a scenario speaks to the person running it.
 *
 * Only one scenario needs it, and it needs it badly: a battery that blocks
 * waiting for a click, without saying so, is indistinguishable from one that
 * has hung — and this project has spent two rounds failing to tell those
 * apart. The trace carries the same words to the Live steps list, so a pane
 * that never wires this up is quieter but not silent.
 */
let prompt: ((message: string) => void) | undefined;

/** Let the pane put a scenario's request in front of the user. */
export function setSelfTestPrompt(fn: (message: string) => void): void {
  prompt = fn;
}

export function setSelfTestRasterizer(fn: (scene: Scene) => Promise<string | undefined>): void {
  rasterizer = fn;
}

/**
 * Every scenario's name, in run order — for a picker that cannot drift.
 *
 * Derived from `SCENARIOS` rather than repeated beside it, because a list of
 * strings maintained by hand next to the list it describes is a rename away
 * from offering a scenario that does not exist.
 */
/**
 * Everything the Scenario picker offers — including what a full run leaves out.
 *
 * The picker's whole purpose is reaching one scenario without the ones in front
 * of it, so a `pickedOnly` scenario has to be here. It is the ONLY way to reach
 * one.
 */
export const SCENARIO_NAMES: readonly string[] = SCENARIOS.map((s) => s.name);

/** What a full run actually runs. A strict subset — see `pickedOnly`. */
export const ROUTINE_SCENARIO_NAMES: readonly string[] = SCENARIOS.filter((s) => !s.pickedOnly).map((s) => s.name);

/**
 * Run the battery, or ONE scenario of it.
 *
 * `only` exists because of what a full round costs. The first run that produced
 * a usable trace took 496 seconds to reach scenario seven and wedged there, and
 * the scenario before it left four charts on the deck as loose piles of shapes.
 * Iterating on the seventh scenario by running the six in front of it is eight
 * minutes and real damage per attempt, which is not iteration.
 *
 * The first two are inserted anyway whatever is picked: every other scenario
 * needs the probe charts they create, and one that finds none reports SKIPPED —
 * an honest answer, and a useless round. So "only" means "only this, plus what
 * it needs", and the report says which is which.
 */
export async function runSelfTest(
  prefix = `selftest ${newRunId()}`,
  only?: string,
  /**
   * Called with each verdict the moment it exists.
   *
   * The battery returns ONE array, after the last scenario, and the round
   * writes its file from that — so a tab that dies mid-battery has always taken
   * every finished verdict with it. That is not a hypothetical: it is why
   * `chartIsVisible` is `pickedOnly`, why the scenario order in `SCENARIOS` is
   * load-bearing, and why this file already carries the sentence "a battery
   * that never returns never writes its report".
   *
   * Ordering the list can only ever decide WHICH verdicts a crash costs. This
   * decides whether it costs any: the caller writes each one somewhere that
   * outlives the pane's JavaScript context. Optional, and wrapped by the
   * caller's own guard — a reporting sink must not be able to fail a run.
   */
  onResult?: (result: ScenarioResult) => void,
): Promise<ScenarioResult[]> {
  const wanted = only ? SCENARIOS.filter((s, i) => i < 2 || s.name === only) : SCENARIOS.filter((s) => !s.pickedOnly);
  // Per run, never carried between them: a stale "this host wedges" from the
  // last round would make the next one skip a scenario on evidence it no longer
  // has — which is exactly the sort of quiet, sticky wrong answer this file is
  // built to avoid.
  selectionWedged = null;
  const out: ScenarioResult[] = [];
  /**
   * Keep the verdict, and hand it to the caller in the same breath.
   *
   * One helper rather than a call beside each `push`, because there are three
   * push sites and the two that report "not reached" are exactly the ones a
   * future edit would forget — they are the verdicts of a run that is already
   * going wrong, which is when this matters most. A sink that throws must not
   * cost the verdict it was given, let alone the run.
   */
  const report = (result: ScenarioResult, into: ScenarioResult[]) => {
    into.push(result);
    try {
      onResult?.(result);
    } catch {
      /* a reporting sink is never a reason for the battery to stop */
    }
  };
  /** Consecutive scenarios that told us the host is in trouble. */
  let sick = 0;
  /** Set once the breaker has tripped, so the rest report why. */
  let abandoned: string | null = null;
  for (const { name, run } of wanted) {
    // Stop asking a host that has stopped answering.
    //
    // The battery had no rung for this, and every real-host artefact this
    // project owns ends the same way. The last one: at 261.6s a scenario failed
    // ("the chart was gone"), at 261.8s the deck scan read 0 of 8 slides — and
    // the battery then ran Same Scale for 95 seconds (six charts × 24 shapes,
    // six consecutive 5010 failures, two 90-second waits), then a picture
    // round-trip, then two more scenarios, until PowerPoint killed the tab at
    // ~396s. Roughly 130 seconds and 150 shapes were issued AFTER three
    // distinct signals that the host was already gone, and the tab took the
    // remaining verdicts with it.
    //
    // Three in a row rather than one, because a single failure is often the
    // finding — the battery exists to produce those. Three consecutive is not a
    // finding about a scenario, it is a finding about the host.
    if (!abandoned && sick >= SICK_LIMIT) {
      abandoned = `${sick} scenarios in a row could not get an answer out of the host`;
      trace("selftest", "giving up on the host", { after: out.length, why: abandoned });
    }
    if (abandoned) {
      report({ name, ok: false, skipped: true, blind: true, detail: `not reached — ${abandoned}`, ms: 0 }, out);
      continue;
    }
    // The battery had no stop check of its own — none at all. So even where a
    // scenario ended promptly, Stop could not end the RUN: the pane switched
    // its button to "Stopping…" and the next scenario started anyway. A
    // battery is the longest thing in the Testing panel and the likeliest to
    // be abandoned half way, which makes this the one loop that most needed it.
    //
    // Scenarios not reached are reported as skipped, with the reason, rather
    // than dropped: a report missing its last four lines looks like a battery
    // that crashed, and that is a different diagnosis from one that was
    // stopped.
    if (isStopRequested()) {
      report({ name, ok: false, skipped: true, detail: "not reached — the run was stopped", ms: 0 }, out);
      continue;
    }
    // Announced BEFORE it runs, not only after it finishes.
    //
    // Every verdict this battery emits is a past-tense record, so a run that
    // dies mid-scenario leaves the PREVIOUS scenario's line as its last word —
    // off by one, and pointing at the one thing that demonstrably worked. Twice
    // now a real-host failure has been diagnosed from a screenshot rather than
    // a log, because the log only exists once the run ends and these runs did
    // not. This line is what makes such a screenshot name the right scenario.
    // What the deck and the host looked like GOING IN. A scenario's verdict has
    // never carried either, so "the host stopped answering" has always been a
    // sentence with nothing beside it: no deck size, no elapsed time, no
    // indication of how the host had been behaving in the seconds before.
    //
    // Recorded on every scenario, not only the ones that fail — a value written
    // down only on failures cannot be compared against anything, which is the
    // mistake `idleMs`, `afterAnswering`, the settle's shared label and
    // `listChartsInDeck` each made in turn.
    const deckBefore = (await deckSlideIds().catch(() => undefined))?.length;
    const frictionBefore = hostFrictionCounts();
    trace("selftest", "scenario starting", { name, deckSlides: deckBefore ?? "unreadable", atMs: traceElapsed() ?? 0 });
    const t0 = Date.now();
    // Deadlines already fired before this scenario ran, so the diff below is
    // this scenario's own. Read here rather than once per run: what matters is
    // whether THIS one could not get an answer, not whether the run ever
    // couldn't.
    const deadlinesBefore = deadlinesFired;
    // Same reasoning as `deadlinesBefore`, for the other half of a stall: a
    // late answer that arrived during an EARLIER scenario must not be read as
    // this one's call coming back. See `lastLateSyncOwner` for the same bug one
    // level down.
    const lateSeqBefore = lastLateSyncSeq;
    let result: ScenarioResult;
    try {
      const r = await run(prefix);
      result = { name, ...r, ms: Date.now() - t0 };
    } catch (err) {
      // A host that stopped answering did not fail the scenario — it prevented
      // it. "We did not check" and "we checked and it is broken" is the
      // distinction this whole file is built on, and the generic catch was
      // collapsing them: on 2026-08-08 `the chart is actually visible` ran
      // eleventh, ten minutes into a round, and reported
      // `FAILED — threw: PowerPoint did not respond while drawing shapes 1-10
      // of 24 (45s)`. That is a fatigued host, said in the words of a broken
      // chart. Picked alone at 61 seconds the same build PASSED.
      //
      // `blind` because the verdict is not evidence about the product — the
      // same reason a blinded deck scan carries it — and it is what makes
      // `selfTestNeedsAttention` still surface the round.
      if (isTimeout(err)) {
        // Read BEFORE the wait below: `waitForLateSync` can let the abandoned
        // call finally answer, which would overwrite the record with the very
        // call that stalled and lose the predecessor this exists to name.
        const after =
          lastStall?.afterAnswering != null ? { call: lastStall.afterAnswering, idleMs: lastStall.idleMs } : undefined;
        // Then ask the one question the round log could not answer: did the
        // call we gave up on ever come back?
        //
        // It has never once been answered in the affirmative — 13 abandoned
        // calls across seven real-host rounds, not one late answer — but that
        // was read out of the ABSENCE of a trace line, which is precisely the
        // inference this project has been burned by. "Never asked" and "asked
        // and answered no" look identical in a log that only writes on yes.
        // So the scenario waits, briefly, and says which one happened.
        await waitForLateSync(LATE_ANSWER_WAIT_MS);
        const late = lastLateSyncSeq !== lateSeqBefore ? (lastLateSync ?? undefined) : undefined;
        result = {
          name,
          ok: false,
          skipped: true,
          blind: true,
          detail: stallDetail(errorText(err), late, after),
          ms: Date.now() - t0,
        };
      } else {
        result = { name, ok: false, detail: `threw: ${errorText(err)}`, ms: Date.now() - t0 };
      }
    }
    // How much the host misbehaved DURING this scenario, on the same terms
    // whether it passed or failed. `errors` is every host refusal the renderer
    // caught, `idRefusals` the `InvalidParam passed to GetItem(id)` this deck is
    // made of, and `emptyReReads` the shape collection coming back empty —
    // which is what decides whether a chart gets grouped and so whether it stays
    // re-editable. Recorded on passes too, because a value written only on
    // failures cannot be compared against anything.
    const f0 = frictionBefore;
    const f1 = hostFrictionCounts();
    // On the RESULT as well as the trace line, because `describeSelfTest` has
    // to tell our defects from this host's weather and the trace is not
    // available to it. Same numbers, one source.
    // Every counter the snapshot has, differenced. Written as a loop over the
    // snapshot's own keys rather than a list, because the list is what went
    // stale — the same reason `resetHostFriction` is a loop.
    result.friction = Object.fromEntries(
      Object.keys(f1).map((k) => [k, (f1 as Record<string, number>)[k] - (f0 as Record<string, number>)[k]]),
    ) as NonNullable<ScenarioResult["friction"]>;
    const deckAfter = (await deckSlideIds().catch(() => undefined))?.length;
    trace("selftest", result.skipped ? "scenario skipped" : result.ok ? "scenario passed" : "scenario FAILED", {
      name,
      detail: result.detail,
      ms: result.ms,
      deckSlides: deckAfter ?? "unreadable",
      ...(deckBefore !== undefined && deckAfter !== undefined ? { deckGrew: deckAfter - deckBefore } : {}),
      friction: result.friction,
      // Carried into the round log, because the regression gate cannot tell an
      // absence of measurement from a fall without it — and inferring it from
      // the detail TEXT would tie the gate to prose that gets edited.
      ...(result.skipped ? { skipped: true } : {}),
    });
    const timedOut = deadlinesFired > deadlinesBefore;
    if (timedOut) trace("selftest", "the host missed a deadline in this scenario", { name, sick: sick + 1 });
    sick = hostSeemsSick(result, timedOut) ? sick + 1 : 0;
    report(result, out);
  }
  return out;
}

/** Whether anything in this run needs looking at — a failure, or a blind scan. */
export function selfTestNeedsAttention(results: ScenarioResult[]): boolean {
  return results.some((r) => (!r.ok && !r.skipped) || r.blind);
}

/** The one-line summary the pane shows. */
export function describeSelfTest(results: ScenarioResult[]): string {
  const ran = results.filter((r) => !r.skipped);
  const failed = ran.filter((r) => !r.ok);
  const blind = results.filter((r) => r.skipped && r.blind);
  const skipped = results.length - ran.length - blind.length;
  // OUR defects lead, and they are the only number that can show whether the
  // add-in is improving. A red scenario on this host is usually the shape
  // collection dying part-way through, which moves with the host's mood and
  // never with our work — so a single "N of M passed" mixed the two and could
  // not answer the question anyone actually asks of a series of rounds.
  const ours = failed.filter((r) => scenarioBlame(r) === "ours");
  const host = failed.filter((r) => scenarioBlame(r) === "host");
  const parts = [
    ours.length ? `${ours.length} defect(s) of ours: ${ours.map((f) => f.name).join(", ")}` : "no defects of ours",
  ];
  parts.push(`${ran.length - failed.length} of ${ran.length} scenarios passed`);
  // Named, never hidden. The split exists to make the first number meaningful,
  // not to make failures go away — these still have to be read, they just are
  // not evidence about the product.
  if (host.length)
    parts.push(`${host.length} failed while the host was refusing: ${host.map((f) => f.name).join(", ")}`);
  if (skipped) parts.push(`${skipped} skipped (host cannot run them)`);
  // Counted apart, and never folded into the line above. These are not a
  // capability gap: the host got in the way, which is a finding, and the one
  // time it happened it was reported in green.
  //
  // And it says WHICH way. `blind` covers two things now — a deck scan that
  // could not see, and a host that stopped answering mid-scenario — and this
  // line called both of them "the deck scan went blind". The 2026-08-08
  // `1fd6aa3` round reported two scenarios that way when neither had scanned
  // anything: both had timed out drawing. Naming the scenarios is the rest of
  // it, because a summary that says two of eleven went wrong without saying
  // which is a summary someone has to open the file to use.
  if (blind.length)
    parts.push(`${blind.length} could not run — the host got in the way: ${blind.map((r) => r.name).join(", ")}`);
  return `Self-test — ${parts.join(" · ")}.`;
}
