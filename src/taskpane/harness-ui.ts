/**
 * THE TESTING PANEL, AND EVERYTHING ONLY IT USES.
 *
 * Lifted out of `app.ts` on 2026-09-29. It is ~1,890 lines — 30% of what that
 * file was — and it was never mixed in with the product: `wireInsert` split
 * cleanly at one line, with production controls above it and `demo-*` /
 * `experiment-*` below. The 1,369-line function that resulted was not one
 * tangled thing, it was two things sharing a brace.
 *
 * WHAT MOVED, and why these and nothing else. Every symbol here was measured to
 * be used ONLY by the testing panel: `lastRunLog`, `tidyable`,
 * `CROWDED_DECK_SLIDES`, `RunLogFile`, `RunLog`, `repairDeckSpan`,
 * `insertDemoDeckAsFile`, `WEB_SHAPE_BUDGET`, `MAX_SHOTS`, `describeLitter`,
 * `collectDeckEvidence`, `downloadJson`, `wireStepsPanel`, `describeHost`, and
 * the thirteen control wirings. `updateChartResilient`, `boundedRaster` and
 * `isWebHost` sat among them and did NOT move: the production insert path calls
 * all three.
 *
 * WHY IT TAKES ITS DEPENDENCIES AS AN ARGUMENT rather than importing them.
 * `app.ts` imports this module, so this module importing `app.ts` would close a
 * cycle — and an ESM cycle here hands one side an undefined binding at init, in
 * the bundle, with `tsc` and the whole suite green. `test/import-cycles.test.ts`
 * exists to catch exactly that and would fail on it. Six things cross the
 * boundary and they are passed in: `guard` and `keepDisabled` from
 * `wireInsert`'s own closure, and `note`, `setProgress`, `boundedRaster` and
 * `isWebHost` from `app.ts`'s module scope.
 *
 * THE MODULE-LEVEL `wired` IS DELIBERATE AND IT IS THE ONE THING TO DISLIKE.
 * Four of the moved helpers call `note` or `setProgress`, and they are
 * module-level declarations rather than closures, so they cannot see
 * `wireHarness`'s parameter. Threading a `deps` argument through them would
 * have changed four signatures and every call site for no behavioural gain, and
 * would have made this move a rewrite rather than a relocation — the code below
 * is otherwise byte-identical to what it replaced, which is what makes it
 * reviewable. The shims read through `wired` and throw if anything runs before
 * `wireHarness`, so the failure is loud rather than an undefined call.
 *
 * The honest follow-on is to extract `app.ts`'s status strip (`note`,
 * `setProgress`, the elapsed timer — one concern, about 190 lines) into a module
 * both files import. Then four of the six dependencies disappear and `wired`
 * shrinks to nothing. Not done here because this commit moves code and changes
 * none of it.
 *
 * AND WHAT THIS UNLOCKS, which is the reason it was worth doing first. The pane
 * ships its own test harness to every user — `app.ts` imported `./selftest` and
 * `../render/host-probe` statically, so ~12,700 lines of battery, probe and
 * `scripts/host-baseline.mjs` prose are in the 317 KB bundle every user loads.
 * That is deliberate and documented ("a build flag would mean the bundle users
 * get is not the bundle the round loop tests"). After this move those two
 * imports live HERE, and the only thing pulling the harness into the bundle is
 * one static import of this file. Gating it behind `?harness=1` — which
 * `manifest-harness.xml` already appends and `manifest-prod.xml` already does
 * not — becomes a one-line change that preserves the invariant that comment
 * defends, because a dynamic import is the same artifact at the same sha.
 * That is the owner's call and is not made here.
 */
import {
  type ReconcileOutcome,
  type SlideInventory,
  type SlideShot,
  applyReconcilePlan,
  canInsertSlidesFromBase64,
  deckSlideIds,
  deleteSlideById,
  enrichSnapshots,
  errorText,
  insertDemoDeck,
  insertSlidesFromPptx,
  isStopRequested,
  listChartsInDeck,
  namedShape,
  newRunId,
  readAddedSlides,
  refreshNamedShapeFromDeck,
  resetSyncCount,
  roundEnvironment,
  scanGap,
  scanIsComplete,
  slideCount,
  slideShots,
  slideSize,
  syncsSoFar,
  traceEnvironment,
} from "../render/powerpoint";
import {
  SCENARIO_NAMES,
  type ScenarioResult,
  describeSelfTest,
  runSelfTest,
  selfTestNeedsAttention,
  setSelfTestPrompt,
  setSelfTestRasterizer,
} from "./selftest";
import {
  type HostAnswerSheet,
  describeHostSheet,
  mergeHostSheets,
  reaskPlan,
  runHostProbes,
  sheetNeedsAttention,
} from "../render/host-probe";
import {
  beginCrashLog,
  clearCrashLog,
  endCrashLog,
  flushCrashLog,
  markCrashLogSaved,
  recordCrashFinding,
  recordCrashStep,
  recoverCrashLog,
} from "./crashlog";
import {
  type TraceSummary,
  formatTraceLine,
  onTrace,
  setTracing,
  trace,
  traceLog,
  traceMark,
  tracing,
} from "../core/trace";
import { type ResultRow, type ResultsSummary, buildResultsScenes, demoItems } from "../core/demo";
import { EXPERIMENTS, runExperiment } from "../render/experiments";
import { type ExpectedItem, type SlideSnapshot, describeReconcile, planReconcile } from "../core/reconcile";
import { type Scene, estimateOfficeShapes } from "../core/scene";
import { buildDeckBase64 } from "../render/pptx-deck";

/** Vite replaces this at build time; see `vite.config.ts`. Module-scoped, so it is declared again here. */
declare const __BUILD_STAMP__: string;

/**
 * Why a verification pass produced no verdict, when it produced none.
 *
 * A `null` here used to mean three different things — nothing tagged, a host
 * that would not answer, and a repair that threw — and the caller could tell
 * them apart from none of them.
 *
 * Moved from `app.ts` with `repairDeckSpan` on 2026-09-29: once that left,
 * nothing in `app.ts` referenced this type at all.
 */
type VerifyResult =
  { kind: "ok"; outcome: ReconcileOutcome } | { kind: "unidentified"; why: string } | { kind: "error"; why: string };

/** The pane's own element lookup, kept identical to `app.ts`'s. */
const $ = (id: string): HTMLElement => document.getElementById(id) as HTMLElement;

/** What the testing panel needs from the pane around it. See the header. */
export interface HarnessDeps {
  guard: (fn: () => Promise<void>) => (this: unknown, ev?: Event) => Promise<void>;
  keepDisabled: (btn: HTMLButtonElement) => void;
  note: (text: string, status?: "ok" | "err" | "busy" | "none", params?: Record<string, string | number>) => void;
  setProgress: (p: number | "busy" | null) => void;
  boundedRaster: (scene: Scene) => Promise<string | undefined>;
  isWebHost: () => boolean;
  /** Part of the status strip, like `note` — see the header's note on the follow-on extraction. */
  noteHostActivity: (now?: number) => void;
}

let wired: HarnessDeps | undefined;

/** Loud rather than undefined: a helper reached before `wireHarness` is a wiring bug, not a host one. */
function deps(): HarnessDeps {
  if (!wired) throw new Error("the testing panel was used before wireHarness() ran");
  return wired;
}

const note: HarnessDeps["note"] = (text, status, params) => deps().note(text, status, params);
const setProgress: HarnessDeps["setProgress"] = (p) => deps().setProgress(p);
const boundedRaster: HarnessDeps["boundedRaster"] = (scene) => deps().boundedRaster(scene);
const isWebHost: HarnessDeps["isWebHost"] = () => deps().isWebHost();
const noteHostActivity: HarnessDeps["noteHostActivity"] = (now) => deps().noteHostActivity(now);

/**
 * A one-line host descriptor for the demo title/results slides, e.g.
 * "PowerPoint · OfficeOnline · 16.0.1". Web vs desktop vs Mac behave differently
 * under the demo's shape budget, so a run's PDF should say which one it was.
 * Guarded — `Office.context.diagnostics` is absent outside a host.
 */
function describeHost(): string {
  try {
    const d = Office.context?.diagnostics;
    if (d?.host) return `${d.host} · ${d.platform} · ${d.version}`;
  } catch {
    /* Office.context unavailable — fall through */
  }
  return "unknown host";
}

/**
 * The last demo run, kept at module scope so it can be saved AFTER the fact.
 *
 * A run's own record used to live and die inside the click handler's closure:
 * when a run ended badly the summary slide was the first casualty, and the
 * only surviving evidence was whatever the user thought to copy out of a
 * one-line note. A run that reveals a host bug is worth more than that.
 */
let lastRunLog: RunLogFile | undefined;

/**
 * Slides the last round added, and the only thing Clean up will ever delete.
 *
 * A list of ids the round watched appear, not a rule for recognising a test
 * slide. The difference matters in someone's own deck: a rule can match a slide
 * they made, an id list cannot.
 */
let tidyable: string[] = [];

/**
 * What "Download run log" writes: one file, one or more runs inside it.
 *
 * A single click can now take both insert paths one after the other, and they
 * fail in completely different ways — so they are separate runs with separate
 * identities, sharing only the build and host that produced them. A
 * single-path click writes the same shape with one entry, so nothing reading
 * this file needs to care which it was.
 */
/**
 * How many slides make a deck too crowded to draw a whole demo onto shape by
 * shape.
 *
 * Not a measurement of where the host gives up — nobody has one, because every
 * attempt so far started on a deck the file half had already filled. It is a
 * threshold below which the run is worth attempting: a demo adds 38 slides, so
 * anything past a handful means the deck was not fresh.
 */
const CROWDED_DECK_SLIDES = 10;

interface RunLogFile {
  build: string;
  host: string;
  /**
   * The slide size this round ran at.
   *
   * Optional because 53 archived rounds predate it, and every one of those was
   * 16:9 — anything reading this must default to that rather than guess, and
   * `docs/ROUNDS.md` states it once so no reader has to infer it.
   */
  slideSize?: { width: number; height: number; source: string };
  /**
   * How many `context.sync()` calls the round made, end to end.
   *
   * Optional because every round archived before 2026-09-02 predates the
   * counter — and absent must read as "never counted", never as zero, which is
   * a number this field can genuinely take on a round that did nothing.
   *
   * See `syncsSoFar` in `src/render/powerpoint.ts` for why syncs rather than
   * shapes: office-js#6329 reports that each one forces a full presentation
   * save on the web host, which would make this the load figure that matters
   * and every shape count in this archive the wrong index.
   */
  syncs?: number;
  runs: RunLog[];
  /**
   * The host self-test's verdicts, when that is what produced this log.
   *
   * Kept beside the runs rather than inside one: the scenarios are not an
   * insert and have no slots, and folding them into a run's item list would
   * make them look like charts that failed to draw.
   */
  selftest?: ScenarioResult[];
  /**
   * The self-test's trace. A run carries its own; the self-test is not a run,
   * so its steps hang here. Declared rather than spread in untyped — an
   * undeclared field is one nothing downstream can be expected to read.
   */
  trace?: RunLog["trace"];
  /**
   * The host probe's answer sheet, when the round produced one.
   *
   * Declared rather than spread in untyped, for the reason the field above
   * gives: an undeclared field is one nothing downstream can be expected to
   * read. `npm run host-diff` takes either shape — a bare sheet, or a round's
   * file with this inside it — so one upload covers both halves.
   */
  hostAnswers?: HostAnswerSheet;
  /**
   * Set when the run did LESS than it was asked to, and why.
   *
   * "Both, one after the other" degrades to the file half on a deck that is
   * already large — see `CROWDED_DECK_SLIDES`. Without this the log shows a run
   * asked for two halves that produced one, with nothing to say the difference
   * was deliberate, and the obvious reading is that the shape half crashed.
   * Data rather than a trace line, because the trace is optional and this is
   * not.
   */
  refusedShapeHalf?: { slides: number; why: string };
  /**
   * What the deck actually held when the round finished, and what it looked like.
   *
   * This is the upload that used to be a person's job. Every diagnosis in this
   * project's history has needed three things — the run log, the deck, and a
   * screenshot — and the owner has been saving and sending all three by hand,
   * once per round, for as long as there have been rounds. Two of the three are
   * things the add-in can read for itself: `SlideInventory` is every shape on
   * every slide (names, ids, positions), and `SlideShot` is the host's own
   * rendering of the slides the round touched.
   *
   * `newSlides` is the id diff across the round, which is what makes the
   * pictures worth having: a picture of every slide in a 40-slide deck is
   * megabytes of things nobody changed.
   */
  deck?: {
    inventory: SlideInventory[];
    /** What the scan could not see — the same honesty every other scan reader owes. */
    gap?: string;
    newSlides: string[];
    /**
     * The before-list could not be read, so `newSlides` is empty for want of an
     * answer rather than because nothing was added. Present only in that case.
     */
    beforeUnknown?: boolean;
    shots: SlideShot[];
  };
}

interface RunLog {
  /**
   * This run's identity, the same token carried on every slide's slot tag.
   *
   * The join key between this file and the .pptx it produced. Every diagnosis
   * in this project's history has come from reading the two together, and
   * until now that read started by guessing which slides belonged to the run —
   * a deck holds whatever earlier runs left in it, and one recent file carried
   * 30 slides from the run under investigation plus a lone slide from another.
   * With the token the join is a lookup and the stray is labelled, not
   * mistaken for a loss. `npm run triage` does exactly that join.
   */
  run: string;
  totalMs: number;
  items: {
    title: string;
    status: string;
    shapes: number;
    ms: number;
    grouped: boolean;
    /**
     * What the run believes it wrote — NOT what the settled readback saw.
     * The two disagreeing is the whole point: a run reported 20 tagged charts
     * where the produced file carried 31, and establishing that took unzipping
     * the .pptx. Compare against `reconcile.snapshots[].tagged` for the same
     * slot: true here and false there is a readback fault; false here and true
     * there is impossible; false in both is a genuinely lost write.
     */
    tagged: boolean;
    /**
     * Whether this item was MEANT to carry a config — not whether it got one.
     *
     * Not every demo item is a chart: the title page, the contents pages and
     * several elements are drawn as `PowerChart` objects with no config by
     * design. Nothing in the produced file distinguishes those from a chart
     * whose tag write was lost, and on the shape path `tagged` cannot either
     * (false there means both "never had one" and "the write did not land").
     * Without this field a triage of a clean run called seven healthy slides
     * broken.
     */
    chart: boolean;
    /**
     * This item gave up on a host call. Separate from `lateOutcome` because
     * the host usually answers an abandoned call minutes later, long after the
     * run has moved on — so "which item stalled" is readable here even when
     * "how it ended" is not.
     */
    abandoned: boolean;
    lateOutcome: string;
  }[];
  deck: {
    slidesAdded: number;
    addsIssued: number;
    lost: number;
    blank: { position: number; title: string | null }[];
  };
  /** The settled truth: what the deck held once the host stopped moving. */
  reconcile?: ReconcileOutcome;
  /** Why there is no settled verdict, when there is none. */
  unverified?: string;
  /**
   * Which route the run took: one generated file, or shape by shape. They fail
   * in completely different ways, and a log that does not say which is being
   * read is a log that gets diagnosed as the wrong one.
   */
  path: "file" | "shapes";
  /**
   * Step-by-step record, when Verbose trace was on for the run — tallies
   * first, then the entries they were counted from.
   */
  trace?: {
    summary: TraceSummary;
    entries: { ms: number; scope: string; message: string; data?: Record<string, unknown> }[];
    dropped: number;
  };
}

/**
 * Verify and repair the demo slides wherever they are in the deck.
 *
 * By slot tag, never by position. The deck's own layout is the host's
 * business: `insertSlidesFromBase64` puts slides at the FRONT unless told
 * otherwise, and a run that assumed "the slides we just added are at the end"
 * read one slide short at the front and swept in the user's own title slide at
 * the back. The span between the first and last tagged slide is ours; nothing
 * outside it is read, let alone touched.
 *
 * "Tagged" means tagged BY THIS RUN. The span used to be drawn around every
 * SSF Charts slot tag in the presentation, which is only the same thing the
 * first time the deck is inserted. Insert it twice — or duplicate one demo
 * slide, which copies its tag — and the span covered both copies, every item
 * matched two slides, and the pass deleted one whole healthy run as the other
 * one's duplicate. Anything inside the span that is not this run's is reported
 * and left alone.
 */
async function repairDeckSpan(
  expected: ExpectedItem[],
  tagFor: (slot: number) => string | undefined,
  run: string,
): Promise<VerifyResult> {
  let snapshots: SlideSnapshot[];
  // Slides the readback could not see. Anything on one of them comes back
  // `lost`, so this is what stops the run reporting a chart the host kept as a
  // chart the host dropped.
  let unread: number;
  try {
    const count = await slideCount();
    // Pass A only, deck-wide. It has to be deck-wide — the span is discoverable
    // only by reading slot tags, and this run's slides are not necessarily at
    // the tail (`insertSlidesFromBase64` put them at the FRONT the first time
    // anyone tried it on a real host). Passes B and C run below, over the span
    // alone: their answers outside it are read, paid for, and then thrown away
    // by the `inSpan` filter twenty lines down — a tag read and two group-count
    // syncs per slide, spent on the user's own earlier work.
    ({ snapshots, unread } = await readAddedSlides(0, count, false));
  } catch (err) {
    return { kind: "error", why: errorText(err) };
  }
  const tagged = snapshots.filter((s) => s.run === run);
  // No slot tag anywhere means the read found slides but could not identify
  // one of them. Saying "verified" would be a lie and saying nothing is how a
  // whole verification pass went missing from a run report without anyone
  // noticing — so name it.
  if (!tagged.length) {
    return {
      kind: "unidentified",
      why: snapshots.length
        ? `read ${snapshots.length} slide(s), none carrying this run's slot tag`
        : "read no slides at all",
    };
  }
  const first = Math.min(...tagged.map((s) => s.index));
  const last = Math.max(...tagged.map((s) => s.index));
  const inSpan = snapshots.filter((s) => s.index >= first && s.index <= last);
  // Now that the span is known, enrich only it.
  await enrichSnapshots(inSpan);
  const plan = planReconcile(inSpan, expected, { dropOrphanBlanks: true, run });
  try {
    if (!plan.actions.length)
      return {
        kind: "ok",
        outcome: { snapshots: inSpan, plan, applied: { unstamped: 0, regrouped: 0, deleted: 0 }, refused: 0, unread },
      };
    return {
      kind: "ok",
      outcome: { ...(await applyReconcilePlan(plan, tagFor, { left: 60, top: 90 }, inSpan)), unread },
    };
  } catch (err) {
    return { kind: "error", why: errorText(err) };
  }
}

/**
 * Insert the demo deck as ONE generated .pptx instead of drawing it shape by
 * shape — see `src/render/pptx-deck.ts` for why that is worth doing.
 *
 * Returns null when the fast path did not run and it is SAFE to fall back to
 * the shape-by-shape path: the host has no `insertSlidesFromBase64`, the file
 * could not be built, or the call failed with nothing landed. It returns a
 * message — never null — once any slide has landed, because falling back after
 * a partial insert would draw the whole deck a second time on top of it.
 */
async function insertDemoDeckAsFile(items: { scene: Scene; title: string; configJson?: string }[]): Promise<{
  text: string;
  status: "ok" | "err";
  added: number;
  totalMs: number;
  verified: VerifyResult;
  run: string;
} | null> {
  const t0 = Date.now();
  const before = await slideCount();
  // Identity for this insert, carried on every slide's slot tag. Without it the
  // repair pass below cannot tell these slides from the ones an earlier insert
  // left in the same deck, and "cannot tell" ends in a delete.
  const run = newRunId();
  let built: { base64: string; shapesPerSlide: number[] };
  try {
    note("Building the deck…", "busy");
    built = await buildDeckBase64(
      items.map((it, i) => ({ scene: it.scene, title: it.title, configJson: it.configJson, slot: i, run })),
      // Build the file at the DESTINATION's slide size. A generated deck that
      // declares a different size is one PowerPoint rescales on insert, which
      // moves every chart on every slide — silently, and on every 4:3 deck.
      await slideSize(),
    );
  } catch (err) {
    console.warn("SSF Charts: could not build the deck file — falling back to shapes", err);
    return null;
  }
  let added: number;
  try {
    note("Handing the deck to PowerPoint…", "busy");
    setProgress("busy");
    added = await insertSlidesFromPptx(built.base64, items.length);
  } catch (err) {
    // A throw is not proof that nothing landed — the same lesson the
    // shape path learned the hard way. Measure before deciding.
    console.warn("SSF Charts: one-shot deck insert failed", err);
    // …and a failed MEASUREMENT is not proof either. This runs on a host that
    // has just failed once, milliseconds ago; the re-read can fail with it.
    // Treating that as "nothing landed" returns null, and the caller then
    // draws the entire deck a second time on top of however many slides did
    // arrive — the exact double-insert every other guard on this path exists
    // to prevent. When we cannot tell, we do not guess: say so and stop.
    let after: number | undefined;
    try {
      after = await slideCount();
    } catch (readErr) {
      console.warn("SSF Charts: could not measure the deck after the failed insert", readErr);
      return {
        text: "PowerPoint would not take the deck, and would not say how much of it landed. Check the deck before inserting again — running it now could add the slides twice.",
        status: "err",
        added: 0,
        totalMs: Date.now() - t0,
        verified: { kind: "error", why: errorText(readErr) },
        run,
      };
    }
    /**
     * A COUNT THAT CAME BACK UNREAD IS NOT A COUNT, and this line treated it as
     * one. Found 2026-09-29.
     *
     * The `catch` above handles the re-read THROWING. It does not handle the
     * re-read succeeding with nothing in it, which is the other way this host
     * declines to answer — `slideCount()` returns `c.value` straight off the
     * proxy and its own comment is explicit that there is "no safe default here,
     * deliberately… not knowing has to stay distinguishable from knowing". So
     * `after` can be `undefined` on a resolved promise.
     *
     * Then `added` is `NaN`, and every comparison below it is false:
     *
     *   NaN <= 0                → false, so "nothing landed" never returns
     *   NaN < items.length      → false, so the ⚠ "the host took N of M" is suppressed
     *   `Inserted ${added} of…` → the user reads "Inserted NaN of 12 slides"
     *
     * and `added: NaN` serialises into the round file as `added: null`. So the
     * one case where we know least produced the most confident-looking output,
     * with the warning that would have prompted a check turned off.
     *
     * This is the same shape `deckGrowth` exists for in `selftest.ts`, written up
     * there as "`null < wanted` is false in JavaScript" — the fix sailed through
     * one call site four times before. Kept inline rather than reaching for that
     * helper: it lives in the harness half of this pane, and a production path
     * must not depend on code that is on its way out of the bundle.
     */
    if (!Number.isFinite(after)) {
      return {
        text: "PowerPoint would not take the deck, and its answer about how much of it landed was unreadable. Check the deck before inserting again — running it now could add the slides twice.",
        status: "err",
        added: 0,
        totalMs: Date.now() - t0,
        verified: { kind: "error", why: `the deck count came back unread (${String(after)})` },
        run,
      };
    }
    added = after - before;
    if (added <= 0) return null;
  }
  if (added <= 0) return null;

  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  const expected: ExpectedItem[] = items.map((it, i) => ({
    slot: i,
    title: it.title,
    // What the FILE renderer drew, not what the Office.js one would have.
    // They disagree by design: a pie is one custGeom wedge here and a
    // sixteen-triangle fan there, so `estimateOfficeShapes` measured five
    // perfect charts as wreckage on the first real run.
    shapes: built.shapesPerSlide[i] ?? estimateOfficeShapes(it.scene),
    chart: !!it.configJson,
    // The generator writes the tag straight into the .pptx — there is no sync
    // to drop — so on this path the run always knows the tag is there. That is
    // what makes an unreadable slide safe to leave alone here, and unsafe to
    // leave alone on the shape path. See ExpectedItem.wroteTag.
    wroteTag: !!it.configJson,
  }));
  // Verify against the deck rather than trusting the count: the file carried
  // its own grouping and tags, so anything missing here is the host's doing.
  // Located by slot tag, not by position — where the host puts the slides is
  // its business, and on the first real run it put them at the FRONT.
  const verified = await repairDeckSpan(expected, (slot) => items[slot]?.configJson, run);
  const outcome = verified.kind === "ok" ? verified.outcome : undefined;

  let text = outcome
    ? `Inserted as one file in ${secs}s — ${describeReconcile(outcome.plan)}.`
    : `Inserted ${added} of ${items.length} slides as one file in ${secs}s.`;
  if (verified.kind !== "ok") text += ` (Not verified: ${verified.why}.)`;
  // A slide the readback could not see puts its item in the `lost` column for
  // want of evidence. Saying so is the difference between "the host dropped
  // your chart" and "we could not look" — and the first reading is what makes
  // someone insert the deck again.
  if (outcome?.unread) text += ` (Could not read ${outcome.unread} slide(s) — anything on them counts as lost here.)`;
  if (added < items.length) text += ` ⚠ the host took ${added} of ${items.length} slides.`;
  const clean = added >= items.length && !!outcome && outcome.plan.summary.lost === 0;
  return { text, status: clean ? "ok" : "err", added, totalMs: Date.now() - t0, verified, run };
}

/**
 * Shapes a web host may be asked to draw in one run before the add-in stops
 * asking. Every 12-item run in this project's history (~400 shapes) survived;
 * every 37-item one (~1850) took the whole client down. 600 sits between them,
 * nearer the side that lived.
 *
 * Only the web needs a number at all. Microsoft's documented runtime limits —
 * CPU, memory, four-crashes-per-session, five-seconds-unresponsive — are
 * scoped to Windows and Mac and explicitly NOT to a browser, so on the web
 * nothing throttles a runaway add-in: the tab dies and takes the session with
 * it. Desktop has that safety net and does not need this one.
 */
const WEB_SHAPE_BUDGET = 600;

/**
 * How many of a round's new slides get a picture.
 *
 * A cap rather than a guess about size: at 480px a slide PNG is on the order of
 * tens of kilobytes, and a round that added forty slides would otherwise turn a
 * readable JSON into something nobody opens. Twelve covers every round this
 * project has produced, and `slideShots` reports the ones it skipped rather than
 * dropping them, so a bigger round says so instead of looking smaller.
 */
const MAX_SHOTS = 12;

/**
 * The cap when **Picture every slide** is ticked — high enough not to bind.
 *
 * The default twelve is right for an ordinary round and wrong for one specific
 * question, which the 2026-08-09 round asked and could not answer: 29 of that
 * deck's 36 slides read back with zero shapes, and only 12 had a picture to
 * corroborate it. So 12 were confirmed blank on two witnesses and 17 were
 * unknowable — and "unknowable" is the answer the cap produced, not the host.
 *
 * A readback of zero is one witness, and this host answers a shape collection
 * short without throwing (`shapes-items-count-honest`), so one witness is not
 * enough to call a slide empty. The picture is the second. Ticking the box
 * spends the extra time and the extra megabytes to get it for every slide.
 *
 * **Now the default**, at the owner's call, and the rounds since have earned it:
 * the two-witness count is the only thing that has actually settled an
 * empty-slide question, and it was the capped rounds that could not. Shipped
 * opt-in first because the pictures are the heaviest call the add-in makes and
 * they run at the END of a round, on a host that has just been through the
 * self-test — the moment it is least able to take more. That cost is real and
 * has not changed; what changed is which side of it is worth defaulting to.
 * Untick the box for a struggling host or a smaller file. `slideShots` still
 * reports anything it skips, so even this cap can bind and say so.
 */
const MAX_SHOTS_ALL = 200;

/**
 * What the round left in the deck, in the sentence the owner reads at the end.
 *
 * A round leaves its slides behind ON PURPOSE — the scenarios' working slides
 * are the evidence, and `docs/REGRESSION.md` is written around a deck someone
 * can open. What was missing is that the pane never SAID so. The 2026-08-09
 * evening round added 43 slides of which 36 came back empty, said "Saved as one
 * file", and left the owner to discover a 44-slide deck by opening it.
 *
 * This repo has already paid for that once at the other end of the same run:
 * the host probe left 21 blank slides in a deck, reported nothing, and the only
 * way to find out was to look. That got a row in the answer sheet. The
 * self-test never got the equivalent.
 *
 * The count is what the READBACK said, and the wording says so rather than
 * claiming the slides are empty: this host reports a shape collection short
 * without throwing (`shapes-items-count-honest`), so a zero here is one witness.
 * `npm run triage` is where that gets cross-examined against the host's own
 * pictures — deliberately not duplicated into the pane, because two copies of
 * one claim is how the two stop agreeing.
 */
function describeLitter(deck: RunLogFile["deck"]): string {
  if (!deck?.newSlides?.length) return "";
  const added = new Set(deck.newSlides);
  // The larger of the two readings, so a partial listing is not called empty —
  // `count` is the host's own number, `shapes` is what the scan managed to list.
  const empty = (deck.inventory ?? []).filter(
    (s) => added.has(s.slideId) && Math.max(s.count ?? 0, s.shapes?.length ?? 0) === 0,
  ).length;
  const n = added.size;
  return (
    ` It left ${n} slide${n === 1 ? "" : "s"} in this deck${empty ? `, ${empty} of which read back empty` : ""} —` +
    " the scenarios' own working slides, kept so the deck is evidence." +
    " Press Clean up the last round when you have finished with them."
  );
}

/**
 * HOW LONG THE TAIL GETS BEFORE IT IS ABANDONED.
 *
 * `collectDeckEvidence` has always SAID it is best-effort — "a host too far gone
 * to describe its own deck still gets the verdicts out" — and nothing enforced
 * it. A failure was survivable because the function catches; a HANG was not,
 * because there was no bound, and a hang is what a dying host actually does.
 *
 * What that cost: the crash band is 441-572s, 9 builds of 13 die in this tail,
 * and 33 of 49 crash records hold a complete fourteen-scenario round that was
 * never filed. `readChartsPage` bounds its own sync at `READBACK_TIMEOUT_MS` —
 * ninety seconds — and `PowerPoint.run` need not settle at all once the host is
 * going, so the tail could sit there for minutes with every verdict already in
 * hand.
 *
 * 45 SECONDS, and the archive says that is enormously generous:
 *
 *     the whole tail, scanning -> done    31 completions   p50 5.6s   max 7.96s
 *     the deck scan alone               6,785 scans        p50 1.3s   max 22.1s
 *
 * Zero of either over 30s, ever. So this is 5.6x the worst complete tail on
 * record and 2x the worst single scan inside it — chosen to be impossible to hit
 * by a host that is merely slow, which is the only way a bound like this can be
 * wrong. (The tail has only 31 completions because per-phase tracing landed on
 * 2026-08-27; the scan figure is the whole archive.)
 *
 * WHAT IT CANNOT DO, said here so nobody credits it with more. This bound
 * catches ONE thing: a `PowerPoint.run` that never settles while every verdict
 * is already in hand. It does not catch the crash. Of the 77 crash reports on
 * file, 10 carry an ErrorName and it is the same one every time —
 * `errorLocalChangeLostSingleUser`, PowerPoint's SERVER-side lost edit — and the
 * other 67 carry none at all. A lost edit is not a hang, so a client-side
 * deadline has nothing to wait out: the dialog is already up and the round is
 * already over. Reading this bound as crash protection would be reading a
 * timeout as a cure for the one failure it cannot reach.
 */
export const DECK_EVIDENCE_TIMEOUT_DEFAULT_MS = 45_000;
let DECK_EVIDENCE_TIMEOUT_MS = DECK_EVIDENCE_TIMEOUT_DEFAULT_MS;

/**
 * Test-only, matching `_setSlideSizeTimeoutForTest`.
 *
 * The DEFAULT is exported and asserted, which is one step more than the sibling
 * takes, because the failure modes are not alike. A slide-size budget that
 * drifts high only wastes time; this one drifts back into the unbounded hang it
 * exists to prevent — and every test of it overrides the value, so nothing else
 * would ever notice.
 */
export function _setDeckEvidenceTimeoutForTest(ms: number): void {
  DECK_EVIDENCE_TIMEOUT_MS = ms;
}

/**
 * What landed on the slides — the two uploads a person has been making by hand.
 *
 * Best-effort by construction, and that is deliberate: this runs at the END of a
 * round, on a host that has just been through the self-test and may well be the
 * reason the round is worth reading. A failure here must cost the pictures and
 * nothing else — the verdicts are already in `lastRunLog`, and losing them to a
 * diagnostic's own tail would be the worst trade in the file.
 *
 * BOUNDED SINCE 2026-08-29, which is what makes the paragraph above true rather
 * than merely intended. On timeout this returns `undefined` — the same path a
 * throw already took — so the round finishes, the log is offered, and the round
 * file simply carries no `deck`. An absent `deck` is loud: `poolGroupingOutcome`
 * and `poolFullestSlide` both read it, and the trace names the abandonment.
 */
async function collectDeckEvidence(idsBefore: string[] | undefined): Promise<RunLogFile["deck"] | undefined> {
  let abandon: ReturnType<typeof setTimeout> | undefined;
  const started = Date.now();
  const bounded = await Promise.race([
    collectDeckEvidenceUnbounded(idsBefore),
    new Promise<"timeout">((resolve) => {
      abandon = setTimeout(() => resolve("timeout"), DECK_EVIDENCE_TIMEOUT_MS);
    }),
  ]);
  clearTimeout(abandon);
  if (bounded !== "timeout") return bounded;
  // NAMED, never silent. A round that files without deck evidence must say why,
  // or the next reader takes an absent inventory for an empty deck — the same
  // "unknown is not a no" mistake this repo has paid for three times.
  trace("pane", "gave up collecting deck evidence", {
    afterMs: Date.now() - started,
    budget: DECK_EVIDENCE_TIMEOUT_MS,
  });
  return undefined;
}

async function collectDeckEvidenceUnbounded(idsBefore: string[] | undefined): Promise<RunLogFile["deck"] | undefined> {
  try {
    // TRACED PER PHASE, because a crash in here has been blaming the wrong step.
    //
    // Four consecutive crashed runs on 2026-08-27 record the same last step,
    // `re-asked what the empty deck could not answer`, and none of them died
    // there — that is simply the last line anything WROTE. What runs next is
    // this function, which scans the whole deck, lists its slide ids and then
    // RENDERS A PICTURE OF EVERY NEW SLIDE, and until now it traced only in its
    // catch. So a host that fell over mid-screenshot filed a crash pointing at
    // an innocent probe re-ask, and `WHERE THE HOST DIED` pooled four of them
    // under that name.
    //
    // Cheap: three lines on a path that already takes seconds and runs once per
    // round, after every verdict is in.
    /**
     * THE SYNC COUNT GOES IN THE LINE THE ROUND DIES ON.
     *
     * 41 of the last 60 4:3 crashes have this trace line as their final entry —
     * the host never reaches the one below. `lastRunLog` is assembled after the
     * scan, so a crashed round's `syncs` field never exists; the trace does,
     * because `traceLog` is a local array read and survives a dead host.
     *
     * That makes this the only place a crashed round can report how much it had
     * asked of the save channel before it went, which is the whole question
     * office-js#6329 raises. Putting it on the same line as the scan
     * announcement rather than a line of its own keeps the crash-log tail
     * comparable with the 160 records already on file.
     *
     * IT WORKED, AND THE FIRST ANSWER CAME ON 2026-09-09. Rounds that died were
     * syncing about 15% faster than rounds that lived — 2.89 against 2.52 per
     * second at 4:3, 2.81 against 2.45 at 16:9, pooled p≈0.001 — with the arm
     * split done precisely because the 4:3 crash rate could otherwise have been
     * the whole effect. Correlation only, and the direction is not established,
     * but it is the sign #6329 predicts and it is the first evidence this
     * archive has produced about it. See `docs/BACKLOG.md`.
     *
     * `syncs` ON THIS LINE IS THE ROUND'S RUNNING TOTAL. The same key on
     * `updated only the shapes that changed` is that ONE chart's syncs, and a
     * query that does not scope to the message mixes the two into a bimodal
     * nonsense. Both sites carry this warning.
     */
    trace("pane", "collecting deck evidence — scanning", {
      knownBefore: idsBefore?.length ?? null,
      syncs: syncsSoFar(),
    });
    const scan = await listChartsInDeck({ withInventory: true });
    trace("pane", "collecting deck evidence — listing slide ids", { charts: scan.charts?.length ?? null });
    const idsAfter = await deckSlideIds();
    // Only slides that were not there before. Without the diff a picture of a
    // forty-slide deck is mostly slides nobody touched — and the id list is the
    // stronger question about a deck than any handle, which is why the diff is
    // taken from ids rather than from counts.
    const known = new Set(idsBefore ?? []);
    // AN UNREADABLE BEFORE-LIST IS NOT AN EMPTY DECK. This fell back to
    // `idsAfter`, which says every slide the deck holds was added by this round
    // — so a round that added four slides to the owner's forty-slide deck
    // reported forty, and `describeLitter` told them it had left forty behind,
    // twelve of which "read back empty". The twelve were their own blank slides.
    // Everything downstream inherits it: `triage.mjs` builds its added-set from
    // this field.
    //
    // The same mistake as reading a host's silence as a "no", which this repo
    // has now paid for in three separate places. Unknown is its own answer:
    // `newSlides` stays empty and `beforeUnknown` says why, so a reader can tell
    // "this round added nothing" from "nobody knows what it added".
    const beforeUnknown = !idsBefore || !idsAfter;
    const newSlides = beforeUnknown ? [] : idsAfter.filter((id) => !known.has(id));
    // Read at collection time, not at boot: the box is ticked for the round
    // about to be read, and a value captured when the pane loaded would be the
    // one from before the owner ticked it.
    const shotAll = ($("demo-shot-all") as HTMLInputElement | null)?.checked ?? false;
    // THE EXPENSIVE ONE — it renders an image per new slide — and it WAS the
    // prime suspect for those four crashes. It is not the culprit, and the
    // per-phase tracing added above is what settled it. Over 2026-08-29's four
    // builds, counted one vote per build the way this archive counts crashes:
    //
    //     e9222a8   4 of 4 crashes        557b10e   1 of 1
    //     3495fb8   5 of 5                4275306   4 of 6
    //
    // every one died at `scanning`, the FIRST phase, before this line was ever
    // reached. That takes the scan to 9 builds against 4 for everything else.
    // The suspect is `listChartsInDeck({ withInventory: true })` above.
    //
    // Still traced with the COUNT it is about to attempt: it remains the
    // expensive call, and a crash that DOES land here should say how many it was
    // asked for rather than leave the number to be inferred from the deck.
    trace("pane", "collecting deck evidence — shooting slides", {
      slides: newSlides.length,
      max: shotAll ? MAX_SHOTS_ALL : MAX_SHOTS,
    });
    const shots = await slideShots(newSlides, { max: shotAll ? MAX_SHOTS_ALL : MAX_SHOTS });
    trace("pane", "collecting deck evidence — done", { shots: shots.length });
    return {
      inventory: scan.inventory ?? [],
      ...(scanIsComplete(scan) ? {} : { gap: scanGap(scan) }),
      newSlides,
      // Carried so a reader can tell an empty `newSlides` that MEANS nothing was
      // added from one that means the question could not be asked. Omitted on
      // the ordinary path, so an existing round file reads exactly as it did.
      ...(beforeUnknown ? { beforeUnknown: true } : {}),
      shots,
    };
  } catch (err) {
    trace("pane", "could not collect deck evidence", { error: errorText(err) });
    return undefined;
  }
}

/** Save an object as a JSON file, via the same Blob dance as Download SVG. */
/**
 * Hand a file to the browser, and say whether the attempt even got that far.
 *
 * Three things were wrong with the two-line version, and a real round on
 * 2026-08-06 lost a whole evidence file to them.
 *
 * 1. **It revoked the object URL in the same tick as the click.** A download is
 *    asynchronous; revoking its source synchronously can cancel it before it
 *    starts. The revoke now waits, and the URL is released on a timer rather
 *    than never — a leaked blob in a pane that stays open for a session is the
 *    lesser of the two.
 * 2. **The anchor was never in the document.** Chromium tolerates a detached
 *    anchor and other engines do not, and a task pane is whatever the host
 *    embeds.
 * 3. **It reported nothing.** Callers said "Saved as one file" on the strength
 *    of having called this. It returns false when the browser refused outright,
 *    which is the case worth acting on: a blocked download that THROWS is now
 *    distinguishable from one that worked.
 *
 * A `false` is proof of failure; a `true` is not proof of success — a frame can
 * still swallow a download silently, and nothing in the DOM reports that. So no
 * caller may treat `true` as "the user has the file"; that is what the explicit
 * save button is for.
 */
function downloadJson(name: string, payload: unknown): boolean {
  let url: string | undefined;
  try {
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.style.display = "none";
    document.body.append(a);
    a.click();
    a.remove();
    return true;
  } catch (err) {
    trace("error", "the browser would not save the file", { name, error: String(err) });
    return false;
  } finally {
    // Long enough for the download to have been picked up, and not tied to the
    // click that starts it.
    const held = url;
    if (held) setTimeout(() => URL.revokeObjectURL(held), 60_000);
  }
}

/**
 * The live transcript panel: the step list, its copy button and its clear
 * button.
 *
 * Lifted whole out of `wireInsert`, which was 1113 lines and is the reason the
 * line formatter inside it went untested for months while it silently dropped
 * every array-valued payload. A panel that owns one DOM node, one bounded
 * buffer and three handlers is a thing that can be read in one screen; the same
 * code as a middle third of a thousand-line function is not.
 *
 * Returns `revealSteps`, which the run buttons call once as a run starts —
 * markup order is not position, and a log you must scroll to before you can
 * photograph it is a log a dying tab takes with it.
 */
function wireStepsPanel(): { revealSteps: () => void } {
  // The live transcript.
  //
  // Wired to the trace stream rather than to the run's own reporting, so it
  // costs no new instrumentation and shows exactly what the log would have —
  // including the `at=<phase>` on any error. It exists because the log does
  // NOT survive the failures worth explaining: it becomes downloadable only
  // when a run ends, and two real-host rounds have now been lost to a run
  // that never ended (a wedge at 1819s) and one PowerPoint killed outright
  // ("Sorry, we ran into a problem", at 108s). What is on screen survives
  // both, and can be copied or photographed before the reload.
  const steps = $("demo-steps");
  /** Newest FIRST, capped — a pane is not a heap, and the head is what is read. */
  const STEP_LINES = 300;
  const lines: string[] = [];
  // NEWEST FIRST, and that ordering is the whole point rather than a
  // preference.
  //
  // The first version appended and auto-scrolled to the bottom, which reads
  // better while a run is healthy and fails at the only moment this list
  // exists for. When PowerPoint dies, what you have is whatever pixels were
  // on screen — no scrolling, no clicking, often a modal dialog over half the
  // pane. An append-and-follow list puts the last thing that happened at the
  // bottom of a small scrolled box, which is exactly where it cannot be
  // relied on to be visible. Prepending puts it at a FIXED position, one line
  // below the header, and needs no scroll to have worked.
  //
  // It also removes the follow-the-tail logic entirely: there is no tail to
  // follow, so there is no "unless the user scrolled" case to get wrong.
  const paintSteps = (): void => {
    steps.textContent = lines.join("\n");
  };
  /**
   * Put the step list where it can be seen, once, as a run starts.
   *
   * Newest-first fixes WHERE in the box the last line is; it does nothing
   * about whether the box itself is on screen. Two things now answer that.
   * The list is the FIRST thing in the Testing section, above the buttons
   * that start a run — a real-host round crashed with the log still under
   * nine controls and a paragraph, which is a log you must scroll to before
   * you can photograph it. And this scrolls the panel to it anyway, because
   * markup order is not position: the Automation tab scrolls, and a run
   * started after reading the JSON section below would otherwise begin with
   * the box off the top. Once per run, on the click that starts it — never
   * while the run is going, because a pane that moves under the cursor
   * mid-run is its own problem.
   */
  const revealSteps = (): void => {
    try {
      steps.scrollIntoView({ block: "nearest" });
    } catch {
      /* an older host without scrollIntoView options — the list still fills */
    }
  };
  onTrace((e) => {
    // The data payload matters as much as the message for the lines that
    // locate a failure — `error`, `name`, `detail` are where the phase and
    // the verdict live. `formatTraceLine` is where that is decided, and it
    // lives in `trace.ts` rather than here because it used to live here: an
    // inline formatter in a DOM closure is a formatter nothing can test, and
    // for months it silently dropped every array-valued payload — including
    // the two timing series `degradation curves` exists to produce.
    const line = formatTraceLine(e);
    lines.unshift(line);
    // Drop the OLDEST, which is now the end of the array.
    if (lines.length > STEP_LINES) lines.length = STEP_LINES;
    paintSteps();
    // The same line, to storage, where it outlives this JavaScript context.
    // One formatter for both, so the file and the screen can never describe
    // the same run differently — and oldest-first there, because a file is
    // read from the top while a crashed screen is read from where it froze.
    recordCrashStep(line);
    // Anything traced is the host or the run still moving, which is what the
    // elapsed readout needs to tell "slow" from "gone".
    noteHostActivity();
  });
  // The last synchronous moment the pane gets. `pagehide` fires on the tab
  // close that ended the 1819-second run, so it buys back the final debounce
  // window — the part of a dying run that nothing else can reach.
  window.addEventListener("pagehide", flushCrashLog);
  $("demo-steps-copy").addEventListener("click", () => {
    if (!lines.length) {
      note("No steps to copy yet.", "err");
      return;
    }
    // Labelled, because the order is the opposite of what a log usually is
    // and a reader who assumes otherwise reads the run backwards.
    const text = [`SSF Charts steps — NEWEST FIRST (${lines.length} lines)`, ...lines].join("\n");
    // Two ways, because the first one does not work where this runs.
    //
    // `navigator.clipboard` needs a secure context, a user gesture AND the
    // `clipboard-write` permission — and an Office task pane is a nested
    // cross-origin iframe that is routinely refused it. Observed, on the run
    // this button exists for: "The browser would not give us the clipboard".
    // Telling the user to select the text by hand is not a fallback, it is
    // an apology, and it arrives at the moment they can least afford one.
    //
    // So: select the transcript and run the legacy copy command, which is
    // permitted from a user gesture in an iframe. It is deprecated and it
    // works. If even that is refused the text is at least now SELECTED, so
    // Ctrl+C finishes the job.
    const selectAndCopy = (): boolean => {
      try {
        const range = document.createRange();
        range.selectNodeContents(steps);
        const sel = window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(range);
        return document.execCommand("copy");
      } catch {
        return false;
      }
    };
    const done = (n: number) => note(`Copied ${n} step(s).`, "ok");
    void Promise.resolve()
      .then(() => navigator.clipboard?.writeText(text))
      .then(() => done(lines.length))
      .catch(() => {
        if (selectAndCopy()) done(lines.length);
        else note("The browser refused the clipboard — the steps are selected, press Ctrl+C.", "err");
      });
  });
  $("demo-steps-clear").addEventListener("click", () => {
    lines.length = 0;
    paintSteps();
  });
  return { revealSteps };
}

/** Wire the Testing panel. Called from `wireInsert` when the host is PowerPoint. */
export function wireHarness(d: HarnessDeps): void {
  wired = d;
  const { guard, keepDisabled } = d;
  // Testing aid: one demo slide per chart kind + feature/element highlights.
  // Turning the fast path OFF on the web is a decision worth flagging. At
  // volume the shape-by-shape path there does not merely stall: the full
  // 37-item deck took the whole web client down five seconds in on
  // 2026-07-31 — "Sorry, we ran into a problem. Please try again." The
  // twelve-item subset survived, so this is a warning and not a block.
  const pathSelect = $("demo-path") as HTMLSelectElement | null;
  pathSelect?.addEventListener("change", () => {
    if (pathSelect.value !== "shapes" || !canInsertSlidesFromBase64()) return;
    if (isWebHost()) {
      note(
        "Heads up: the full deck drawn shape by shape has crashed PowerPoint on the web. The fast path handles it in seconds.",
        "err",
      );
    }
  });
  // Enabled by a run, not by the host: with nothing to save it would only
  // ever produce an empty file.
  // ON by default for now. Nothing in this project has been diagnosed from
  // anything but an after-the-fact artifact, and the runs that matter happen
  // on a host nobody can attach a debugger to. When the add-in stops being
  // validated against real hosts, uncheck it in taskpane.html and drop the
  // `checked` — the module, its call sites and this toggle all keep working,
  // so a future investigation is one click away rather than a re-implementation.
  // The live transcript — see `wireStepsPanel`.
  const { revealSteps } = wireStepsPanel();

  const traceToggle = $("demo-trace") as HTMLInputElement | null;
  if (traceToggle?.checked) {
    setTracing(true);
    traceEnvironment(typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev");
  }
  traceToggle?.addEventListener("change", () => {
    setTracing(traceToggle.checked);
    if (traceToggle.checked) {
      traceEnvironment(typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev");
      note("Verbose trace on — it rides along in the run log.", "ok");
    } else note("Verbose trace off.", "ok");
  });
  $("demo-log").addEventListener("click", () => {
    if (!lastRunLog) {
      note("No run to save yet — insert the demo deck first.", "err");
      return;
    }
    if (!downloadJson("ssf-charts-run-log.json", lastRunLog)) {
      note("The browser would not save the file. Copy the Live steps instead — they carry the same run.", "err");
      return;
    }
    // A button the user pressed is the strongest evidence this pane can have
    // that they now hold the run. It is also the recovery from an auto-save
    // that was blocked, so it has to clear the stored record — otherwise the
    // pane would keep offering back a run they have just saved.
    markCrashLogSaved();
    note("Run log saved.", "ok");
  });
  /**
   * Offer the last run that never reported finishing.
   *
   * Checked once, on the open that follows the crash — which is the only
   * moment anyone is looking for it, and the moment before the natural next
   * action (run it again) would otherwise bury it. Hidden entirely when
   * there is nothing to recover, so a healthy pane carries no wreckage.
   */
  /**
   * Ask this host the fixed question list and save what it says.
   *
   * The one diagnostic here that is not about a run at all. Everything else
   * in this panel reports what the ADD-IN did; this reports what the HOST is,
   * so the fake that every test in the repo stands on can finally be checked
   * against the thing it stands for. One click, no deck changes — it works on
   * a scratch slide and takes it back.
   */
  /**
   * ONE QUESTION, ANSWERED IN SECONDS.
   *
   * The round is the wrong instrument for settling a single "does this host do
   * X?" that a decision is waiting on — fourteen minutes, and the question has
   * to earn a slot in a fixed sheet. `grouped-child-by-id-from-slide` waited
   * 125 rounds for a slot and never got one.
   *
   * The alternative was worse: putting a speculative host call in the drawing
   * batch to find out, which is where loading an id on a creation handle
   * poisons it and costs the tag that makes a chart re-editable.
   */
  {
    const pick = $("experiment-pick") as HTMLSelectElement;
    for (const e of EXPERIMENTS) {
      const option = document.createElement("option");
      option.value = e.id;
      // The QUESTION in the list, not the id. The id is for the archive; a
      // person choosing one wants to read what it asks.
      option.textContent = e.asks;
      pick.append(option);
    }
    $("experiment-run").addEventListener(
      "click",
      guard(async () => {
        revealSteps();
        const chosen = pick.value || EXPERIMENTS[0]?.id;
        note(`Asking: ${EXPERIMENTS.find((e) => e.id === chosen)?.asks ?? chosen}`, "busy");
        const r = await runExperiment(chosen);
        // The DETAIL beside the word, always. The vocabulary will be wrong for
        // something eventually and the detail is what survives that.
        note(`${r.id} — ${r.answer}${r.detail ? `: ${r.detail}` : ""} (${r.ms}ms)`, r.answer === "yes" ? "ok" : "err");
      }),
    );
  }

  $("demo-probe").addEventListener(
    "click",
    guard(async () => {
      revealSteps();
      note("Asking this PowerPoint what it actually does…", "busy");
      const sheet = await runHostProbes(describeHost(), typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev");
      const saved = downloadJson("ssf-charts-host-answers.json", sheet);
      // The diff, here, now — rather than after a round trip.
      //
      // Every probe run so far has been "download it, send it, wait for
      // someone to run `host-diff`, hear back". Most of those establish
      // nothing new: the answers are the same as last time. The comparison
      // table is a plain object, so the pane can do it and say whether this
      // run is worth sending at all.
      note(
        saved
          ? describeHostSheet(sheet)
          : `The browser would not save the file. ${describeHostSheet(sheet)} Copy the Live steps — they carry the answers.`,
        !saved || sheetNeedsAttention(sheet) ? "err" : "ok",
      );
    }),
  );
  /**
   * One click, one file: the probe and the self-test, back to back.
   *
   * What it saves is round trips rather than seconds. A round used to be
   * three clicks producing three downloads, uploaded separately and joined at
   * the other end — and most probe runs establish nothing, so a good share of
   * that traffic was to learn that the answers had not changed.
   *
   * The DEMO DECK is deliberately not in here, and not for want of effort.
   * Its two halves have to run on different decks: the file half fills the
   * deck, and the shape half then draws onto that same larger deck, which is
   * the one configuration that has ended in PowerPoint's crash dialog every
   * time it has been tried. A button cannot open a fresh deck, so chaining
   * the demo in would bake in exactly the arrangement the runbook splits up.
   *
   * The probe goes FIRST because it is the cheap one. If the host is already
   * unwell, seventeen short questions say so in seconds, and they are still
   * in the bundle when the long half dies.
   */
  const roundBtn = $("demo-round") as HTMLButtonElement;
  roundBtn.disabled = false;
  roundBtn.addEventListener(
    "click",
    guard(async () => {
      revealSteps();
      const buildStamp = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
      const host = describeHost();
      lastRunLog = undefined;
      ($("demo-log") as HTMLButtonElement).disabled = true;
      beginCrashLog({ build: buildStamp, host, label: "the whole round" });
      // ZEROED WITH THE ROUND, beside the trace mark and for the same reason:
      // a count that carries over from the pane's earlier life belongs to no
      // round, and the pane is reused across rounds far more often than it is
      // reloaded. See `syncsSoFar`.
      resetSyncCount();
      const traceFrom = traceMark();
      // AFTER THE MARK, OR IT IS NOT IN THE ROUND. `traceEnvironment` has been
      // called at pane-wiring time since #228 and its output has reached
      // exactly ZERO of the 69 archived rounds — `traceLog(traceFrom)` slices
      // off everything traced before the mark, so the host, the platform, the
      // Office version, the slide size and (as of yesterday) the deck-style
      // replay were all written into a window no round file ever sees.
      //
      // Measured, not reasoned: `host`/`environment` appears in 0 of 69 round
      // files, and 0 of `slide size`, on an archive where every build contains
      // the emitting line. The pre-mark window is real and large — round 092's
      // first entry is at ms 48245 with `dropped: 0`, so 48 seconds of trace
      // was sliced away.
      //
      // The replay had a second, earlier blocker of its own: `wireInsert()`
      // runs synchronously BEFORE `void readDeckStyle()` in `Office.onReady`,
      // so at wiring time no read has happened and `deckStyleVerdict` is null.
      // Here it has, which is the other half of why this call belongs here.
      traceEnvironment(buildStamp);
      // Read BEFORE the probe, because the probe adds scratch slides too — and
      // a slide the diagnostic created is exactly one worth a picture. An
      // unreadable deck here is not a reason to stop: it costs the id diff and
      // nothing else, and the round says so rather than reporting an empty
      // diff as "the round added nothing".
      const idsBefore = await deckSlideIds();
      // How loaded the deck was when the round started, as the first thing the
      // crash log carries.
      //
      // A run that dies leaves only its steps, and on 2026-08-09 a round died
      // sixteen seconds in — two 8-second stalls and then the tab — with no
      // way to tell an already-tired PowerPoint from a fresh one. The deck is
      // the difference: this project has documented since 2026-08-06 that
      // heavy work on a deck that is already large is what kills the tab, and
      // the runbook splits the demo halves across two decks for exactly that.
      // Whether a crashed round was on a fresh deck is the first question
      // anyone asks, and nothing recorded the answer.
      // The round's own environment, once, from local sources only. Half the
      // hypotheses this project has entertained are about the tab rather than
      // the deck — its age most of all, which has been a live candidate for
      // ten rounds with nothing measuring it.
      trace("selftest", "round starting", {
        deckSlides: idsBefore?.length ?? "unreadable",
        env: roundEnvironment(),
      });
      note("Round 1 of 2 — asking this PowerPoint what it actually does…", "busy");
      let sheet = await runHostProbes(host, buildStamp);
      // Written into the bundle before the long half starts. A self-test that
      // takes the tab down must not also lose the probe's answers, which are
      // complete, cheap, and the half most likely to be worth reading.
      lastRunLog = { build: buildStamp, host, runs: [], hostAnswers: sheet };
      // …and into the store that outlives this JavaScript context, for the
      // same reason and against a bigger loss. `lastRunLog` is a module
      // variable: the line above protects the probe's answers from a battery
      // that FAILS, and not at all from a tab that DIES, which is the way
      // these rounds actually end. The sheet was already complete and minutes
      // old at that point, and it went with the tab every time.
      recordCrashFinding("hostAnswers", sheet);
      note(`Probe done — ${describeHostSheet(sheet)} Now the self-test…`, "busy");
      if (isStopRequested()) {
        const saved = downloadJson("ssf-charts-round.json", lastRunLog);
        note(
          saved
            ? "Stopped after the probe. Its answers are saved."
            : "Stopped after the probe, but the browser would not save the file — press Download run log.",
          saved ? "ok" : "err",
        );
        endCrashLog(saved);
        return;
      }
      setSelfTestRasterizer(boundedRaster);
      setSelfTestPrompt((message) => note(message, "busy"));
      const results = await runSelfTest(undefined, scenarioPick?.value || undefined, (r) =>
        // Each verdict banked as it lands. A battery that never returns never
        // writes its report, and ordering `SCENARIOS` can only choose which
        // verdicts a crash costs — this is what makes it cost none of the
        // ones already reached.
        recordCrashFinding(`selftest:${r.name}`, r),
      );
      // The questions an empty deck could not be asked.
      //
      // The sheet is built at the top of a round, before the battery has drawn
      // anything — so a question needing an id THIS host has agreed to name is
      // put at the one moment no such id exists.
      // `shape-resolve-held-slide-proxy` answered `no-scratch-shape` in 216 of
      // 216 rounds for that reason and never once reached its question.
      //
      // Re-asked HERE, after the battery has drawn and tagged charts, rather
      // than by moving the sheet: the early write is what keeps the probe's
      // answers when the battery takes the tab down, and that is the failure
      // these rounds actually have. The merge re-banks the sheet, so the worst
      // a re-ask can cost is the re-ask.
      //
      // Skipped when nothing got tagged. Then there IS no named id, and
      // `no-scratch-shape` is the true answer rather than a deferral.
      // OBSERVED, not remembered. See `refreshNamedShapeFromDeck`: a chart
      // recorded when its tag was written is only as good as the minutes that
      // follow, and a round spends those redrawing and sweeping. Rounds 242,
      // 243, 246 and 247 all answered `unreadable` off an id whose shape was
      // no longer on the slide, and 247 proved it by listing the slide.
      //
      // Best-effort: a scan that throws must not cost the re-ask, which can
      // still run on whatever the record holds.
      let observed = namedShape();
      try {
        observed = await refreshNamedShapeFromDeck();
      } catch (err) {
        trace("probe", "could not scan the deck for a chart to name", { error: errorText(err) });
      }
      const plan = reaskPlan(sheet, observed);
      if (plan.ask.length) {
        note(`Re-asking ${plan.ask.length} question(s) now that a chart exists…`, "busy");
        const again = await runHostProbes(host, buildStamp, { only: plan.ask, passes: 1 });
        sheet = mergeHostSheets(sheet, again);
        recordCrashFinding("hostAnswers", sheet);
        trace("probe", "re-asked what the empty deck could not answer", {
          asked: plan.ask,
          answered: plan.ask.map((id) => sheet.answers.find((r) => r.id === id)?.answer ?? "missing"),
        });
      } else if (plan.skipped.length) {
        trace("probe", "left questions deferred — no chart this host would name", { deferred: plan.skipped });
      }
      /**
       * ASSEMBLED BEFORE THE STEP THAT KILLS THE HOST, and this ordering is
       * the whole point.
       *
       * `collectDeckEvidence` used to run FIRST and this object was built from
       * its result. Every verdict was already in by then — that function's own
       * docstring says so — but if the host died inside it, `lastRunLog` was
       * never assigned, `demo-log` was never enabled, and a complete round
       * evaporated. The driver found nothing to download and filed a crash.
       *
       * That is not rare and it is not cheap. **33 of 49 crash records in
       * `crashes/` hold all fourteen verdicts**, over 13 builds — about a tenth
       * of this archive again, thrown away for want of a deck scan. On
       * 2026-08-29 one 4:3 leg produced a full fourteen-scenario result FIVE
       * times, 14/14 four of them, and archived none.
       *
       * And the scan is where it dies: per-phase tracing puts 9 builds on
       * `collecting deck evidence — scanning` against 4 on everything else,
       * in a band of 441-572s.
       *
       * IT WORKS BECAUSE THE PANE OUTLIVES THE HOST. A crash takes Office.js
       * down, not this iframe — `keepCrashedRun` presses "Download the crashed
       * run" after exactly this crash and gets a file every time. Saving is
       * Blob and DOM, no host call, so a button enabled here still answers.
       *
       * `slideSize()` is a host call and stays on this side of the scan, where
       * the host is still alive to answer it.
       *
       * Same lesson as "SAVE FIRST, then end the record" below, one step
       * earlier in the same sequence.
       */
      lastRunLog = {
        build: buildStamp,
        host,
        // WHICH SLIDE SIZE THIS ROUND RAN AT, and it is load-bearing rather
        // than decorative. Until 2026-08-16 every round in the archive was
        // 16:9 and nothing said so — then the first 4:3 round was filed into
        // the same directory, where `npm run rounds` pools it with the rest.
        // Averaging two aspect ratios into one number is the rounds 24-and-25
        // mistake ("differed only in this, and were compared as though they
        // did not"), and a nightly 4:3 round would repeat it every night.
        //
        // Read through `slideSize()`, which resolves it from the host with
        // three fallback rungs, so this records what the ROUND actually ran
        // at rather than what anyone believed it would.
        slideSize: await slideSize(),
        runs: [],
        hostAnswers: sheet,
        selftest: results,
        ...(tracing() ? { trace: traceLog(traceFrom) } : {}),
      };
      /**
       * THE BUTTON STAYS DISABLED UNTIL THE ROUND IS ACTUALLY OVER, and
       * enabling it here — which the first version of this did — is a
       * regression that silently strips deck evidence from rounds.
       *
       * `scripts/round.mjs:2002` is the reason:
       *
       *     if (/button "Download run log"(?! \[disabled\])/.test(dl)) break;
       *
       * That button becoming enabled IS the driver's "the round has finished"
       * signal. Enabling it before `collectDeckEvidence` made the driver break
       * out of its wait loop mid-tail and download the banked log — verdicts
       * complete, `deck` absent, trace snapshotted before the scan.
       *
       * Rounds 313 and 318 are that race, both at 4:3, and I read 313 as proof
       * the new timeout had fired. It was not. Nothing had timed out; the
       * driver simply asked early. 314, 316 and 317 kept their deck only
       * because the tail beat the next poll.
       *
       * Durability does not need this. The verdicts survive a dying host
       * through the crash log — `runLogHead` below plus one `selftest:` finding
       * per scenario — which is the path `scripts/salvage-crashed.mjs` already
       * reads, and it works precisely because the pane outlives the host. The
       * button is a completion signal, not a safety net, and using it as both
       * made it a poor version of each.
       */
      /**
       * PERSISTED, because a variable does not survive what comes next.
       *
       * `lastRunLog` is a variable. Recovery RELOADS the tab after a crash, so
       * anything held only in memory is gone before the driver could press
       * anything — which is why the ordering above needs a durable partner.
       *
       * The crash log is that partner and it is already proven: every verdict
       * reaches `crashes/*.json` through `recordCrashFinding`, and all 49
       * records exist because the pane answered a button press AFTER the host
       * died. This adds the four fields a salvage needs that the verdicts do
       * not already carry, so a crashed run becomes a complete round rather
       * than a pile of verdicts missing their header.
       *
       * Deliberately NOT the whole log: `selftest` is already here one key per
       * scenario and the trace is the bulk of the record, so copying either
       * would double the file to say nothing new.
       */
      recordCrashFinding("runLogHead", {
        build: lastRunLog.build,
        host: lastRunLog.host,
        slideSize: lastRunLog.slideSize,
        runs: lastRunLog.runs,
      });
      // Gathered after the scenarios, before the file is written: this is the
      // upload that used to be the owner's job — save the deck, screenshot the
      // pane, attach both. It is a best-effort tail, so a host too far gone to
      // describe its own deck still gets the verdicts out.
      note("Collecting what landed on the slides…", "busy");
      const deck = await collectDeckEvidence(idsBefore);
      /**
       * THE TRACE IS RE-TAKEN WHETHER OR NOT THE TAIL CAME BACK, and the first
       * version of this got that wrong in the one case it was written for.
       *
       * Guarded on `deck`, the re-take never ran when the tail failed — so the
       * filed round carried a trace snapshotted BEFORE the scan, ending at the
       * last pre-tail line, with `deck` absent and nothing anywhere saying why.
       * Round 313 is the proof: 14/14 at 4:3, no `deck`, and not one
       * `collecting deck evidence` line in its trace, so the abandonment I had
       * just added a diagnostic for was the one thing it could not show.
       *
       * That is this repo's most repeated defect — UNKNOWN PRINTED AS NOTHING —
       * arrived at by guarding a diagnostic on the success it was meant to
       * explain the absence of. `traceLog` is a local array read, so it is
       * safe on a host that has stopped answering, which is exactly when it
       * matters most.
       */
      lastRunLog = {
        ...lastRunLog,
        ...(deck ? { deck } : {}),
        /**
         * HOW MANY TIMES THIS ROUND ASKED THE HOST TO SYNC.
         *
         * Taken here, after the deck scan, so it covers the whole round —
         * including the phase 41 of the last 60 4:3 crashes die in. See
         * `syncsSoFar` for why it is worth having: if office-js#6329 is right
         * that every sync forces a full presentation save, then the load this
         * add-in puts on the save channel is proportional to THIS number and
         * not to any shape count, and every crash figure this repo owns is
         * indexed by shapes.
         *
         * A crashed round never reaches this line, which is why the count is
         * traced before the scan as well: the trace survives into the crash
         * record, so a round that died still reports one.
         */
        syncs: syncsSoFar(),
        ...(tracing() ? { trace: traceLog(traceFrom) } : {}),
      };
      // NOW the round is over, and only now may the button say so — it is the
      // driver's completion signal (`scripts/round.mjs:2002`), so enabling it
      // any earlier makes the driver download a round that is still running.
      ($("demo-log") as HTMLButtonElement).disabled = false;
      // Only what THIS round added, and only what it could name. The button
      // stays disabled when the id diff came back empty, because a cleanup
      // with nothing to work from is one that would have to guess which
      // slides look like a test — and guessing about deletion in a user's own
      // deck is not a trade this pane makes.
      tidyable = deck?.newSlides ?? [];
      ($("demo-tidy") as HTMLButtonElement).disabled = tidyable.length === 0;
      const litter = describeLitter(deck);
      // SAVE FIRST, then end the record — and end it with what the save
      // actually did. The other order is what lost a real round: the run was
      // marked finished, which made it unrecoverable, and the download was
      // attempted afterwards. PowerPoint died, the pane reopened, and there
      // was nothing to offer back.
      const saved = downloadJson("ssf-charts-round.json", lastRunLog);
      endCrashLog(saved);
      const needed = sheetNeedsAttention(sheet) || selfTestNeedsAttention(results);
      note(
        `Round finished. ${describeSelfTest(results)} · Probe: ${describeHostSheet(sheet)} ` +
          (!saved
            ? "The browser would NOT save the file — press Download run log, or copy the Live steps."
            : needed
              ? "Saved as one file — send it over. If it is not in your downloads, press Download run log."
              : "Saved as one file; nothing in it is new.") +
          litter,
        needed || !saved ? "err" : "ok",
      );
    }),
  );
  /**
   * Put the deck back.
   *
   * A round leaves slides behind on purpose — the point is a file someone can
   * open and look at — and clearing them afterwards has been a manual chore
   * once per round, in a deck that also grows and skews the next round's
   * timings. This deletes exactly the ids the last round recorded adding, one
   * at a time, and reports what the host refused rather than claiming a clean
   * sweep it did not perform. `deleteSlideById` has a whole comment about why
   * a host saying "gone" is not proof; the count here is what it actually
   * confirmed.
   */
  $("demo-tidy").addEventListener(
    "click",
    guard(async () => {
      revealSteps();
      const ids = tidyable;
      note(`Removing the ${ids.length} slide(s) the last round added…`, "busy");
      let gone = 0;
      for (const id of ids) if (await deleteSlideById(id)) gone++;
      // Emptied whatever happened: a second press would re-ask about slides
      // the host has already refused once, and the honest state after a
      // partial sweep is "there is no longer a list I trust".
      tidyable = [];
      keepDisabled($("demo-tidy") as HTMLButtonElement);
      note(
        gone === ids.length
          ? `Cleaned up — ${gone} slide(s) removed.`
          : `Removed ${gone} of ${ids.length}. The host would not take the rest; delete those by hand.`,
        gone === ids.length ? "ok" : "err",
      );
    }),
  );
  const crashBtn = $("demo-crashlog") as HTMLButtonElement;
  const crashed = recoverCrashLog();
  if (crashed) {
    crashBtn.hidden = false;
    // Two different runs land here now, and telling the owner which one it is
    // is the difference between "the host died" and "the host was fine and
    // your file never arrived". Both are worth recovering; only one of them
    // means anything went wrong with the run itself.
    note(
      `A previous run ("${crashed.label}", build ${crashed.build}) ` +
        (crashed.finishedAt ? `finished, but its file was never saved` : `never reported finishing`) +
        ` — ${crashed.steps.length} step(s) were kept. Download the crashed run.`,
      "err",
    );
    crashBtn.addEventListener("click", () => {
      if (!downloadJson("ssf-charts-crashed-run.json", crashed)) {
        note("The browser would not save the file. Copy the Live steps instead.", "err");
        return;
      }
      clearCrashLog();
      crashBtn.hidden = true;
      note("Crashed run saved.", "ok");
    });
  }
  // The five paths the demo deck never touches. Its own button rather than a
  // mode of the demo run: it edits and deletes as well as inserting, and a
  // user reaching for "insert a demo deck" should not get that by accident.
  // Fill the picker from the battery's own list, so it cannot offer a
  // scenario that no longer exists or miss one that was added.
  const scenarioPick = $("demo-scenario") as HTMLSelectElement | null;
  if (scenarioPick) {
    for (const name of SCENARIO_NAMES) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      scenarioPick.append(opt);
    }
  }
  const selfTestBtn = $("demo-selftest") as HTMLButtonElement;
  selfTestBtn.disabled = false;
  selfTestBtn.addEventListener(
    "click",
    guard(async () => {
      revealSteps();
      const buildStamp = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
      lastRunLog = undefined;
      ($("demo-log") as HTMLButtonElement).disabled = true;
      beginCrashLog({ build: buildStamp, host: describeHost(), label: "host self-test" });
      const traceFrom = traceMark();
      // The same rasteriser the demo run degrades with — the picture
      // scenario needs a real PNG, not a config that merely says "image".
      setSelfTestRasterizer(boundedRaster);
      // A scenario that blocks on a person has to be able to ask. Routed to
      // the same note the rest of the pane speaks through, so the request is
      // where the user is already looking rather than buried in a step list.
      setSelfTestPrompt((message) => note(message, "busy"));
      const results = await runSelfTest(undefined, scenarioPick?.value || undefined);
      // No runs, but a log all the same — the scenarios ARE the record, and
      // the trace beside them is what says how each verdict was reached.
      lastRunLog = {
        build: buildStamp,
        host: describeHost(),
        runs: [],
        selftest: results,
        ...(tracing() ? { trace: traceLog(traceFrom) } : {}),
      };
      ($("demo-log") as HTMLButtonElement).disabled = false;
      // Only on the way out, and only here. A run that throws past this line
      // stays marked unfinished on purpose: it produced no downloadable run
      // log either, so the storage copy is the only record it has.
      //
      // Finished, NOT saved. This path writes no file — the user presses
      // *Download run log* — so until they do, the storage copy is still the
      // only copy, and `markCrashLogSaved` is what retires it.
      endCrashLog();
      note(describeSelfTest(results), selfTestNeedsAttention(results) ? "err" : "ok");
    }),
  );
  const demoBtn = $("demo-insert") as HTMLButtonElement;
  demoBtn.disabled = false;
  demoBtn.addEventListener(
    "click",
    guard(async () => {
      revealSteps();
      const buildStamp = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
      const host = describeHost();
      const items = demoItems({ buildStamp, host });
      // Drop the previous run's log before this one starts. It used to
      // survive, so a run that produced no log of its own left "Download run
      // log" enabled and handing out an OLDER run's file, with nothing on
      // screen to say the two were unrelated.
      lastRunLog = undefined;
      ($("demo-log") as HTMLButtonElement).disabled = true;
      beginCrashLog({ build: buildStamp, host, label: "demo deck" });
      /**
       * Runs this click produced, in the order they were taken.
       *
       * "Both" mode takes each path in turn, so one click can end with two.
       * They are kept apart rather than merged: the paths fail in completely
       * different ways, and a report that did not say which was being read
       * would be diagnosed as the wrong one.
       */
      const runs: RunLog[] = [];
      /**
       * Bank a run the moment it ends, rather than when the click does.
       *
       * In "both" mode the shape path runs second and can throw — the whole
       * deck lost to host errors is a real outcome and it is raised as one.
       * Banking at the end of the click would then discard the file run's
       * log along with it, and the failing run is precisely the one worth
       * having a file for. `runs` is the same array the log holds, so later
       * entries land in it without re-assigning anything.
       */
      const record = (r: RunLog) => {
        runs.push(r);
        lastRunLog = { build: buildStamp, host, runs, ...(refusedShapeHalf ? { refusedShapeHalf } : {}) };
        ($("demo-log") as HTMLButtonElement).disabled = false;
      };
      // Which path(s) to take. Both, one after the other, is what a change
      // touching the renderer wants: same session, same host, one deck, and
      // the two accounts directly comparable — instead of two separate runs
      // an hour apart with a deploy in between.
      let mode = (($("demo-path") as HTMLSelectElement | null)?.value ?? "file") as "file" | "shapes" | "both";
      let refusedShapeHalf: RunLogFile["refusedShapeHalf"];
      // …and it has never once survived on a deck that was not empty.
      //
      // Four attempts, four crash dialogs. The shape half always draws onto
      // whatever the file half has just built, and the last one asked the
      // host for a batch of FIVE shapes on a 40-slide deck and waited 45
      // seconds for nothing. Every crash this project has recorded from the
      // demo path has that same shape: heavy shape work on a deck that is
      // already large. `docs/PUBLISHING.md` splits the two into separate
      // tests on separate decks for exactly this reason, and the option here
      // quietly puts them back together.
      //
      // So it degrades rather than obeys: run the file half, and say what
      // the shape half needs. A twenty-minute run that ends in "Sorry, we ran
      // into a problem" costs more than the measurement was worth, and it
      // costs it after the tab has already eaten the log.
      // Only asked where the answer changes anything. A deck read is a host
      // round trip, and the other two modes would pay for it to learn
      // nothing — which is also how this first landed: it broke a test that
      // counts exactly how many reads an insert is allowed to spend.
      if (mode === "both") {
        const deckNow = await slideCount();
        if ((deckNow ?? 0) > CROWDED_DECK_SLIDES) {
          mode = "file";
          note(
            `This deck already holds ${deckNow} slides, so only the file half will run. ` +
              "The shape half needs a fresh deck — every attempt at both on one deck has ended in PowerPoint's crash dialog.",
            "err",
          );
          // In the LOG as well as on screen. The note is overwritten by this
          // run's own summary within seconds, and a reader of the log
          // otherwise sees a run that was asked for both halves and did one,
          // with nothing to say the difference was deliberate.
          trace("demo", "refused the shape half on a crowded deck", { slides: deckNow, needs: "a fresh deck" });
          refusedShapeHalf = {
            slides: deckNow ?? 0,
            why: "the shape half needs a fresh deck — both halves on one deck has crashed the host every time",
          };
        }
      }
      // Where THIS run's trace starts. The buffer keeps every operation since
      // tracing was switched on, so a log that carried all of it carried other
      // runs' entries too — and reading one run's numbers against another's
      // trace is a genuinely expensive mistake.
      let traceFrom = traceMark();
      // The slowest thing the pane can do — say where it has got to, or a
      // multi-minute run is indistinguishable from a hang.
      // Fast path first: one generated .pptx, one host call. Falls through
      // to the shape-by-shape renderer when the host cannot take it, or when
      // the attempt landed nothing — never after a partial insert, which
      // would draw the whole deck again on top of what is already there.
      // "Both" is the one case where drawing it again IS the intent: the two
      // runs carry different tokens, so nothing confuses one for the other.
      if (mode !== "shapes" && canInsertSlidesFromBase64()) {
        const outcome = await insertDemoDeckAsFile(items);
        if (outcome) {
          // A log for THIS path too. The fast path is the default — the
          // checkbox ships checked and every current host advertises
          // insertSlidesFromBase64 — so it is what a real run takes, and it
          // was the one path that produced no downloadable record at all.
          // Success or failure: the failing run is the one worth having.
          const settled = outcome.verified.kind === "ok" ? outcome.verified.outcome : undefined;
          const verdicts = settled?.plan.verdicts ?? [];
          record({
            run: outcome.run,
            totalMs: outcome.totalMs,
            items: items.map((it, i) => {
              const v = verdicts.find((x) => x.slot === i);
              return {
                title: it.title,
                // The settled verdict, which is the honest answer here —
                // there is no per-item render to report on a file insert.
                status: v ? v.status : "unverified",
                shapes: v?.shapes ?? 0,
                ms: 0,
                grouped: v?.tagged ?? false,
                // On this path the deck was BUILT with the tag, so intent is
                // simply whether the item had a config at all. The generator
                // writes it into the .pptx directly — there is no sync to
                // drop — so `tagged` true here against a false snapshot is
                // a readback fault and nothing else.
                tagged: !!it.configJson,
                chart: !!it.configJson,
                // The file path draws no item individually — there is one
                // insert for the whole deck — so no item can have abandoned
                // a call of its own.
                abandoned: false,
                lateOutcome: "",
              };
            }),
            deck: {
              slidesAdded: outcome.added,
              addsIssued: items.length,
              lost: Math.max(0, items.length - outcome.added),
              blank: [],
            },
            reconcile: settled,
            // Why there is no settled verdict, when there is none — the same
            // sentence the user sees, kept with the data it explains.
            unverified: outcome.verified.kind === "ok" ? undefined : outcome.verified.why,
            path: "file",
            trace: tracing() ? traceLog(traceFrom) : undefined,
          });
          if (mode === "file") {
            endCrashLog();
            note(outcome.text, outcome.status);
            return;
          }
          // Both: the shape path runs next, on top of what just landed, and
          // gets its own slice of the trace. Without a fresh mark the second
          // run's log would open with the first run's entries.
          note(`${outcome.text} Now drawing the same deck shape by shape…`, "busy");
          traceFrom = traceMark();
        } else if (mode === "both") {
          note("The host would not take a generated deck — running the shape path only.", "busy");
        } else {
          note("The host would not take a generated deck — drawing it shape by shape instead.", "busy");
        }
      }
      // NO budget exemption for the harness's own slides any more. It existed
      // because a large deck's contents page ran past the limit, and it is
      // what let a 79-shape text slide through as the SECOND thing a run
      // drew — five seconds before PowerPoint on the web crashed on
      // 2026-07-31. `buildIndexScenes` and `buildResultsScenes` now measure
      // their pages and split them instead, so there is nothing left to
      // exempt: a harness page that somehow still runs over is a page worth
      // skipping, exactly like any other.
      const {
        run,
        results,
        slidesAdded,
        addsIssued,
        blankSlides,
        blankItems,
        blanksRead,
        reconcile,
        degradedAt,
        degradeReason,
        totalMs,
      } = await insertDemoDeck(
        items.map((i) => ({
          scene: i.scene,
          tagData: i.configJson,
          title: i.title,
        })),
        (done, total) => {
          note("Inserting demo slides… {done} of {total}", "busy", { done, total });
          setProgress(done / total); // one slide per context, so a real bar
        },
        {
          // Close the run by reading the deck back and repairing it — the
          // per-item bookkeeping below is written while the host is still
          // committing, and has been observed calling a chart failed that
          // in fact landed twice.
          reconcile: true,
          // And give it somewhere to go when the host starts failing.
          // Rasterizing needs the pane's canvas, so the renderer asks and
          // this supplies; it decides when.
          // Web only. Degrading exists because a browser has no safety
          // net — Office's CPU/memory/crash-tolerance limits are scoped to
          // Windows and Mac — so on desktop, where the host throttles the
          // add-in rather than dying, drawing shapes remains the right
          // answer however long it takes.
          pictureFor: isWebHost() ? (i) => boundedRaster(items[i].scene) : undefined,
          shapeBudget: isWebHost() ? WEB_SHAPE_BUDGET : undefined,
        },
      );
      // Self-check: the deck is a regression harness, so report what the HOST
      // actually did, not what we asked for. The full table goes to the console.
      const named = (s: "skipped" | "failed") =>
        results.map((r, i) => (r.status === s ? items[i].title : "")).filter(Boolean);
      const skipped = named("skipped");
      const failedNames = named("failed");
      const rendered = results.filter((r) => r.status === "rendered").length;
      // Loss vs adds ISSUED, not vs items.length: a retry/fail stray inflates
      // slidesAdded, so measuring against items.length reads 0 during real
      // corruption when a stray cancels a lost slide. addsIssued − slidesAdded
      // counts adds that never landed (strays that landed cancel out).
      const lost = Math.max(0, addsIssued - slidesAdded);
      const secs = (totalMs / 1000).toFixed(1);
      console.log("SSF Charts demo self-check:");
      console.table(
        results.map((r, i) => ({
          chart: items[i].title,
          shapes: r.created,
          status: r.status,
          grouped: !!r.grouped,
          tagged: !!r.tagged,
          ms: r.ms,
          abandoned: !!r.abandoned,
          lateOutcome: r.lateOutcome ?? "",
        })),
      );
      console.log(
        `deck grew by ${slidesAdded}, issued ${addsIssued} adds${lost > 0 ? ` — ${lost} LOST` : ""}; blank slots ${blankSlides.length ? blankSlides.join(", ") : "none"}${blanksRead ? "" : " (blank check incomplete)"} · total ${secs}s`,
      );
      // The headline comes from the settled read when there is one. Every
      // number above was computed while the host was still committing, and
      // the run that produced Presentation_4.pptx got three of them wrong:
      // it blamed Gantt for failing (it landed twice), called a full Agenda
      // slide blank, and counted duplicate slides as successes.
      const verdicts = reconcile?.plan.verdicts ?? [];
      const missing = verdicts.filter((v) => v.status === "lost" || v.status === "empty").map((v) => v.title);
      let msg = reconcile
        ? `Deck settled: ${describeReconcile(reconcile.plan)} — in ${secs}s.`
        : `Inserted ${rendered} of ${items.length} in ${secs}s.`;
      if (skipped.length) msg += ` Skipped as too dense (stamped): ${skipped.join(", ")}.`;
      if (reconcile) {
        if (missing.length) msg += ` Never landed: ${missing.join(", ")}.`;
        if (reconcile.unread)
          msg += ` (Could not read ${reconcile.unread} slide(s) — anything on them counts as lost here.)`;
        const { deleted, unstamped, regrouped } = reconcile.applied;
        // Deletions are NOT all duplicates — most are usually empty slides a
        // lost add left behind. Calling nine removals "9 duplicate slide(s)"
        // in the same breath as the plan's own "1 duplicate slide removed ·
        // 8 orphan slides" made the run contradict itself in one sentence.
        const dupPlanned = reconcile.plan.actions.filter((a) => a.kind === "delete" && a.slot !== null).length;
        const emptyPlanned = reconcile.plan.actions.filter((a) => a.kind === "delete" && a.slot === null).length;
        if (deleted || unstamped || regrouped)
          msg +=
            ` Repaired: removed ${deleted} slide(s)` +
            (deleted ? ` (${dupPlanned} duplicate, ${emptyPlanned} empty)` : "") +
            `, cleared ${unstamped} false banner(s), re-grouped ${regrouped} chart(s).`;
        if (reconcile.refused) msg += ` ⚠ ${reconcile.refused} repair step(s) the host refused.`;
      } else if (failedNames.length) msg += ` Host failed on: ${failedNames.join(", ")}.`;
      if (degradedAt !== undefined)
        msg += ` Drew the last ${items.length - degradedAt} slide(s) as pictures — ${degradeReason}. Use "Explode to native shapes" on any of them to get real shapes back.`;
      // A rendered but ungrouped chart is not re-editable — flag them so
      // Phase 2 doesn't quietly count them as full successes.
      const ungrouped = reconcile
        ? verdicts.filter(
            (v) => !v.tagged && v.status !== "lost" && v.status !== "skipped" && items[v.slot]?.configJson,
          ).length
        : results.filter((r, i) => r.status === "rendered" && !r.grouped && items[i].scene.nodes.length > 1).length;
      if (ungrouped) msg += ` ⚠ ${ungrouped} chart${ungrouped === 1 ? "" : "s"} landed ungrouped (not re-editable).`;
      if (lost > 0)
        msg += ` ⚠ ${lost} add${lost === 1 ? "" : "s"} did not land — the host lost slides (issued ${addsIssued}, deck grew by ${slidesAdded}).`;
      // Blank slides carry the slot tag (item title) where the host has 1.3
      // slide tags; without them the entry has title null and only its deck
      // position is shown, same as before slot tags landed.
      if (blankSlides.length) {
        const named = blankItems.map((b) => (b.title ? `${b.title} (slide ${b.position})` : `slide ${b.position}`));
        msg += ` ⚠ ${blankSlides.length} slide${blankSlides.length === 1 ? "" : "s"} came back BLANK: ${named.join(", ")}.`;
      } else if (!blanksRead) msg += ` (Blank check did not finish.)`;
      // Keep the whole run, not just the sentence. A run that ends badly is
      // the one worth reporting, and it is also the one whose results slide
      // is most likely to be the thing the host drops.
      // Close the deck with a self-contained results slide so the exported PDF is
      // a complete run record. A second insertDemoDeck reuses the same add/render/
      // self-check machinery; wrap it so a host stall here can't swallow the run's
      // own summary (its failure is itself just another data point).
      const rows: ResultRow[] = results.map((r, i) => ({
        title: items[i].title,
        status: r.status,
        shapes: r.created,
        ms: r.ms,
      }));
      const summary: ResultsSummary = {
        buildStamp,
        items: items.length,
        rendered,
        skipped: skipped.length,
        failed: failedNames.length,
        lost,
        totalMs,
      };
      // Each results page inserts in its OWN insertDemoDeck call. A single
      // page failing must NOT drop the rest — insertDemoDeck throws when
      // every item in its batch failed, so batching all pages together
      // meant page 2's failure lost page 1. Presentation_3.pptx surfaced
      // this: "(results slide not added)" with zero pages landed.
      let resultsPages: Scene[] = [];
      try {
        resultsPages = buildResultsScenes(rows, summary);
      } catch (e) {
        console.warn("SSF Charts: results scene build failed", e);
      }
      let resultsLanded = 0;
      for (const [i, scene] of resultsPages.entries()) {
        try {
          await insertDemoDeck(
            [
              {
                scene,
                title: resultsPages.length === 1 ? "Results" : `Results (page ${i + 1} of ${resultsPages.length})`,
              },
            ],
            undefined,
            {
              // The same protections the run itself gets. This insert runs at
              // the WORST moment — right after a run that has just finished
              // exhausting the host — and it used to get none of them.
              //
              // Without `reconcile`, a page whose add landed but whose shapes
              // did not left a stamped, untagged slide at the end of the deck
              // that nothing ever cleaned: the main run's repair had already
              // finished, and its range stopped short of this slide. A real
              // 38-item run ended exactly that way.
              reconcile: true,
              // And without a picture to fall back on, the run's own summary
              // is the first casualty of a failure-heavy run — it is being
              // drawn precisely because things went badly.
              pictureFor: isWebHost() ? () => boundedRaster(scene) : undefined,
            },
          );
          resultsLanded += 1;
        } catch (e) {
          console.warn(`SSF Charts: results page ${i + 1} failed to insert`, e);
        }
      }
      if (!resultsPages.length) msg += " (results slide not added)";
      else if (resultsLanded === 0) msg += " (results slide not added)";
      else if (resultsLanded < resultsPages.length)
        msg += ` (${resultsLanded} of ${resultsPages.length} results pages added)`;
      // Written LAST, so the trace it carries covers the whole run including
      // the results pages. Taken before them, the log ended at the repair
      // read — and when a results page then failed, the file said
      // "(results slide not added)" with nothing in it about why.
      record({
        run,
        totalMs,
        items: results.map((r, i) => ({
          title: items[i].title,
          status: r.status,
          shapes: r.created,
          ms: r.ms,
          grouped: !!r.grouped,
          tagged: !!r.tagged,
          chart: !!items[i].configJson,
          abandoned: !!r.abandoned,
          lateOutcome: r.lateOutcome ?? "",
        })),
        deck: { slidesAdded, addsIssued, lost, blank: blankItems },
        reconcile,
        path: "shapes",
        trace: tracing() ? traceLog(traceFrom) : undefined,
      });
      // Reached only by a run that got all the way here. One that did not
      // stays marked unfinished, which is what offers it back on reopen —
      // and so does this one until *Download run log* is pressed, because
      // finishing and being saved are different facts.
      endCrashLog();
      note(
        runs.length > 1 ? `Both paths run. File: ${runs[0].deck.slidesAdded} slides. Shapes: ${msg}` : msg,
        lost > 0 || failedNames.length || blankSlides.length ? "err" : "ok",
      );
    }),
  );
}
