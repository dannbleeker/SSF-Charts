// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { deckGrowth } from "../src/taskpane/selftest";
import {
  installHost,
  makeSlide,
  makeShape,
  rasterised,
  hostSlideSize,
  applyWebProfile,
  faults,
  trips,
  stallSyncOn,
  userClicksShape,
  selectionHandlerCount,
  type FakeSlide,
} from "./helpers/office-host";
import {
  CHART_TAG,
  SHAPES_PER_SYNC,
  requestStop,
  resetStop,
  isStopRequested,
  _setReadbackTimeoutForTest,
  _setBatchTimeoutForTest,
  listChartsInDeck,
  timeShapeRounds,
  gridFootprint,
  shapesDrawnOn,
  _setLastStallForTest,
  _setCountSettleDelayForTest,
  RASTERISE_OP,
} from "../src/render/powerpoint";
import { sampleConfig } from "../src/core/samples";
import { buildChart } from "../src/core/chart";
import { estimateOfficeShapes } from "../src/core/scene";
import type { ChartConfig, ChartKind } from "../src/core/types";
import { setTracing, traceLog } from "../src/core/trace";
import {
  runSelfTest,
  describeSelfTest,
  scenarioBlame,
  selfTestNeedsAttention,
  setSelfTestRasterizer,
  SCENARIO_NAMES,
  ROUTINE_SCENARIO_NAMES,
  hostSeemsSick,
  readDegradation,
  _setClickWaitForTest,
  _setDegradeSizeForTest,
  wedgedSelection,
  leastLoadedChart,
  renderDifference,
  rescaleFlipIndex,
  rescaleLossNote,
  reReadNote,
  updateLossNote,
  rescaleShouldStop,
  visibilityVerdict,
  sideSlot,
  fitInSlot,
  SIDE_SLOTS,
  GRID_SLOT,
  stallDetail,
  rasteriseArmVerdict,
  kindCostVerdict,
  KIND_COST_ORDER,
  type KindDraw,
  type ScenarioResult,
} from "../src/taskpane/selftest";

/**
 * The host self-test — nine paths the demo deck never touches, plus one
 * experiment that a full run deliberately leaves out.
 *
 * The battery's own value is that it runs against a REAL PowerPoint, which
 * nothing here can do. What these cases pin is the property that makes it
 * worth clicking: that it comes back with a verdict for every scenario, that
 * the verdicts say what was observed, and above all that a scenario which
 * blows up does not take the other eight with it. A battery that stops at the
 * first error spends a whole real-host session to learn one thing — and the
 * scenarios after the failure are precisely the ones nobody has data for.
 */

afterEach(() => vi.unstubAllGlobals());

const byName = (rs: ScenarioResult[]) => Object.fromEntries(rs.map((r) => [r.name, r]));

const VISIBLE = "the chart is actually visible";

/**
 * The visibility scenario earned its way back into a routine round.
 *
 * It spent five rounds killing the browser tab — five builds, always within a
 * step or two of `adding a scratch slide`, never once returning a verdict — and
 * was parked as `pickedOnly` so that a crash cost a short round instead of the
 * whole report. Picked alone on `b998a2e` it named its own killer:
 * `getImageAsBase64` on a slide added 0.3 seconds earlier. It stopped doing
 * that, survived on `e49cca8`, and PASSED on `c7d91d5` — `10064 → 15652 bytes`
 * through PowerPoint's own rasteriser.
 *
 * What is pinned here is the shape of that history, not the outcome: parking a
 * scenario is a real decision with a real cost, so both directions have to be
 * deliberate. Moving it back out needs a reason as good as the one that brought
 * it back.
 */
describe("what a routine round covers", () => {
  it("runs the visibility scenario, which earned its place back", () => {
    expect(SCENARIO_NAMES, "the scenario was deleted rather than run").toContain(VISIBLE);
    expect(
      ROUTINE_SCENARIO_NAMES,
      "parked the visibility scenario again — it passed on a real host, so say why in its entry first",
    ).toContain(VISIBLE);
  });

  it("still offers it to the picker on its own", async () => {
    // The picker is how it was diagnosed in the first place: alone, on a fresh
    // host, a minute in, with only its two inserts in front of it. That is what
    // separated "this scenario kills the host" from "ten minutes of drawing
    // kills the host, and this is merely what was running".
    installHost([makeSlide("s1")]);
    const names = (await runSelfTest("probe", VISIBLE)).map((r) => r.name);
    expect(names, "the picker cannot reach it").toContain(VISIBLE);
  });
});

describe("the host self-test battery", () => {
  it("returns a verdict for every scenario, in order", async () => {
    installHost([makeSlide("s1")]);
    const results = await runSelfTest("probe");
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    expect(results.map((r) => r.name)).toEqual([
      "insert on top of an earlier run",
      "two slides claiming one slot",
      "edit a chart on the visible slide",
      "insert onto a slide that already has content",
      "same scale across the deck",
      // The third arm for the first-chart cost — a run of ONE. Directly after
      // the deck-wide rescale, because the reading is only worth anything on a
      // deck that is already warm. See `oneChartAlone`.
      "one chart alone on a warm deck",
      // The drag round trip, in the only form that can be scripted. After the
      // inserting scenarios because it needs a chart to move, and before the
      // ladder because it must not run against a wedged host.
      "an update follows a moved chart",
      "explode a degraded picture",
      // The ladder ahead of everything that selects a shape. Being the FIRST
      // `setSelectedShapes` in the run is the property that lets it be routine
      // at all — not being alone in a run, and not adjacency, which was only
      // ever a proxy for it. Pinned here, and as a property just below.
      "which selection call wedges the host",
      "a selected shape survives an insert",
      // The newest last — a crash in unproven code must not cost the verdicts
      // of scenarios that already work. Pinned, because the ordering is a
      // diagnostic property, not a detail.
      "edit the chart the user selected",
      "stop a run part-way",
      // Its mid-draw sibling, added 2026-09-01. `stop a run part-way` asks for
      // the stop BEFORE the insert, so it throws on iteration zero and no shape
      // is ever queued — 251 of 322 archived rounds say exactly that in their
      // own detail. This one requests the stop from the first `commit`, which is
      // the case the Stop button exists for and had never been exercised on a
      // real host. It cleans up the wreckage it makes, because the deck it draws
      // into is the deck every later scenario reads.
      "a big chart on a slide of its own",
      /**
       * Straight after it, and differing from it in exactly one thing: no stop.
       *
       * `stop a run mid-draw` moved onto a freshly-added slide in `ca138f8` and
       * has thrown `GeneralException` at `SlideCollection.getItem` on all six
       * rounds since, after passing twice on the build before, when it drew on
       * the visible slide. One `src/` change in that boundary, and the elapsed
       * times overlap — a pass at 488s against a failure at 400s — so it is not
       * the host tiring.
       *
       * But that scenario also requests a STOP, so it cannot say which of the
       * two caused it. This one holds everything else and drops the stop, which
       * makes it a test of the SHIPPED `offerOwnSlide` path: add a slide, draw
       * a big chart across eleven syncs onto it. That path fires precisely when
       * a chart is big, and nothing has ever walked it.
       */
      "stop a run mid-draw",
      // Routine again as of `c7d91d5`, where it PASSED on a real host after
      // five rounds of killing the tab — see the case below. Still last, for
      // the reason everything newest is last.
      "the chart is actually visible",
      // Its control, and the newest thing here. Immediately after it on
      // purpose: these are the only two draws in the battery that follow a
      // rasterise, and the pair is what separates the CALL from the SCENARIO.
      "does a rasterise poison the next draw",
      // Last, and it leaves nothing behind: it draws a line chart purely to read
      // back where the host put a ROTATED shape, then deletes it. Every side slot
      // is already occupied by a scenario whose chart stays, and taking a sixth
      // would narrow all the others.
      "a chart of rotated shapes",
      /**
       * Last, and it is the CHEAPEST thing here: one host call that hands over
       * a generated file, against the eleven-batch draw of its sibling above.
       *
       * It exists because the route it drives shipped with no round able to
       * reach it. `offerNativeInsteadOfPicture` puts a too-dense chart on a
       * slide of its own as native shapes, and it lives behind a button — and
       * the battery cannot press buttons, so the newest user-facing insert path
       * in the pane was also the only one nothing exercised.
       *
       * It drives `chartOnItsOwnSlideAsFile`, the shipped function, rather than
       * repeating its two calls here. A re-implementation would have been a
       * double that cannot fail the way the real thing fails, which is the
       * mistake this archive has paid for four times.
       */
      "a chart handed over as a file",
      /**
       * Runs last, and the list's own rule is why: newest, it draws eight
       * specimens, and an unproven scenario's failure should cost the fewest
       * verdicts behind it.
       *
       * It is also the only scenario in the battery that draws anything other
       * than `clustered`, which is the whole reason it exists — see
       * `kindCostSpread`.
       */
      "what a chart kind costs",
    ]);
    for (const r of results) {
      expect(typeof r.ok).toBe("boolean");
      expect(r.detail, `${r.name} reported no detail`).toBeTruthy();
      expect(r.ms).toBeGreaterThanOrEqual(0);
    }
  });

  it("keeps going when a scenario throws, and records the throw", async () => {
    // The property the whole battery rests on. A host that wedges on the first
    // scenario used to mean five unknowns instead of one known and four
    // answers — on a session someone had to sit through to produce.
    installHost([makeSlide("s1")]);
    const boom = new Error("the host wedged");
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw boom;
      },
      SlideLayoutType: { blank: "blank" },
    });
    const results = await runSelfTest("probe");
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    // Every one of them ran and reported; none is missing.
    for (const r of results) {
      expect(r.name).toBeTruthy();
      expect(r.detail).toBeTruthy();
    }
    const threw = results.filter((r) => r.detail.startsWith("threw:"));
    expect(threw.length).toBeGreaterThan(0);
    expect(threw[0].detail).toContain("the host wedged");
  });

  it("separates 'we did not check' from 'we checked and it is broken'", async () => {
    // A host below PowerPointApi 1.5 can insert nothing from base64 and can
    // hold no pictures. Reporting that as a failure would send a diagnosis
    // after a bug that is really a requirement-set gap.
    installHost([makeSlide("s1")], [], undefined, () => false);
    const results = await runSelfTest("probe");
    const skipped = results.filter((r) => r.skipped);
    expect(skipped.length).toBeGreaterThan(0);
    // A skip always says why. Two reasons are possible — the host cannot do
    // it, or an earlier scenario left nothing to work with — and at least one
    // here is the first, since this host advertises no requirement set at all.
    for (const s of skipped) expect(s.detail).toBeTruthy();
    expect(skipped.some((s) => /host/i.test(s.detail))).toBe(true);
    // A skip is never counted as a pass in the headline either.
    expect(describeSelfTest(results)).toContain("skipped");
  });

  /**
   * "Found nothing, therefore nothing was left behind" is an assertion a blind
   * scan satisfies for free.
   *
   * On a real host this scenario reported PASS off a scan that had just read
   * 1 of 8 slides — `unread=7 slides=8` is the line immediately before its own
   * verdict — while two other scenarios reported SKIPPED under exactly the same
   * blindness. An assertion that cannot fail is not a guard, and this one is
   * guarding the promise that Stop is non-destructive, which nothing else
   * checks.
   */
  it("does not pass 'nothing was left behind' on a scan that read nothing", async () => {
    installHost([makeSlide("s1")]);
    // Every deck scan comes back empty and SAYS so. `hollowReads` is per-read
    // and this scenario's scan is the last of many, so a large count is the
    // honest way to hold the whole run blind.
    faults.hollowReads = 500;
    try {
      // Targeted, so the host-sickness breaker cannot trip first and report
      // this scenario as "not reached" — which is also correct behaviour, and
      // would test the breaker instead of the assertion this case is about.
      const r = byName(await runSelfTest("probe", "stop a run part-way"))["stop a run part-way"];
      expect(r.ok, "passed on a scan that read nothing").toBe(false);
      expect(r.skipped, "reported a blind scan as a real verdict").toBe(true);
      expect(r.blind, "attributed the blind scan to something else").toBe(true);
      expect(r.detail).toMatch(/could not see the whole deck/i);
    } finally {
      faults.hollowReads = 0;
    }
  });

  /**
   * The same rule, swept across every scenario that draws a conclusion from a
   * READBACK.
   *
   * Each of these guards its first scan and then made its loudest claim off a
   * second one it never checked. `insertOntoUsedSlide` was fixed for exactly
   * this and its siblings were not swept with it, so a host answering one page
   * short produced a hard FAILED verdict asserting data loss the add-in did not
   * cause — "its config did not survive the redraw", "no longer re-editable",
   * "the config did not survive". Those are the sentences that send a
   * maintainer after the tag-write path.
   */
  it.each([
    "edit a chart on the visible slide",
    "explode a degraded picture",
    "insert on top of an earlier run",
    "two slides claiming one slot",
    "edit the chart the user selected",
    "same scale across the deck",
  ])("does not report data loss from a blind readback: %s", async (name) => {
    installHost([makeSlide("s1")]);
    faults.hollowReads = 500;
    try {
      const r = byName(await runSelfTest("probe", name))[name];
      // Either it never got far enough to read back (skipped for its own
      // reasons), or it read back blind — but it must never be a hard FAIL that
      // blames the add-in for what the host would not show it.
      if (r.ok === false && !r.skipped) {
        expect.fail(`reported a hard failure off a blinded host: ${r.detail}`);
      }
    } finally {
      faults.hollowReads = 0;
    }
  });

  /**
   * Stop asking a host that has stopped answering.
   *
   * Every real-host artefact this project owns ends the same way. The last one:
   * at 261.6s a scenario failed, at 261.8s the deck scan read 0 of 8 slides —
   * and the battery then ran Same Scale for 95 seconds (six charts, six
   * consecutive 5010 failures, two 90-second waits), a picture round-trip, and
   * two more scenarios, until PowerPoint killed the tab at ~396s. Roughly 130
   * seconds and 150 shapes were issued AFTER three distinct signals that the
   * host was already gone, and the tab took the remaining verdicts with it.
   */
  it("stops the run once the host has failed three scenarios in a row", async () => {
    installHost([makeSlide("s1")]);
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("the host wedged");
      },
      SlideLayoutType: { blank: "blank" },
    });
    const results = await runSelfTest("probe");
    // Still a verdict for every scenario — abandoning the run must not lose the
    // report, which is the only thing a run this bad produces.
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    const abandoned = results.filter((r) => r.detail.includes("in a row"));
    expect(abandoned.length, "kept asking a host that had stopped answering").toBeGreaterThan(0);
    // And it gave the host a fair number of tries first: three real attempts
    // before it gives up, or a single flaky scenario would end every run.
    const attempted = results.filter((r) => !r.detail.includes("in a row"));
    expect(attempted.length, "gave up too early to have found anything").toBeGreaterThanOrEqual(3);
    // Never green, and never filed as a capability gap.
    expect(selfTestNeedsAttention(results)).toBe(true);
  });

  /**
   * The three verdicts of a real run, in the order it produced them.
   *
   * PowerPoint on the web, 2026-08-05. Two scenarios hit the 45-second draw
   * deadline and said "did not respond"; the third hit the SAME deadline after
   * 49.8 seconds and reported "the host stopped answering selection calls after
   * a programmatic select — known web-host limitation". The breaker matched
   * prose, so the third reset the counter to zero one short of tripping. The
   * battery then ran two more scenarios on a host that had been dead for three
   * and the tab died, taking their verdicts with it.
   *
   * Every one of these is a verbatim `detail` from that run.
   */
  it("counts a scenario that waited out its deadline, however it words the verdict", () => {
    const timedOutAndSaidSo: ScenarioResult = {
      name: "same scale across the deck",
      ok: false,
      detail:
        "threw: PowerPoint did not respond while drawing shapes 1-10 of 24 (45s) | at=redrawing the chart's shapes",
      ms: 48954,
    };
    const timedOutAndSaidSomethingElse: ScenarioResult = {
      name: "edit the chart the user selected",
      ok: false,
      skipped: true,
      detail:
        "the host stopped answering selection calls after a programmatic select — known web-host limitation, " +
        "same family as office-js#3083 / #3698; the pane's own Edit-it path is unaffected",
      ms: 49840,
    };
    // The rule this replaced, verbatim, so the guard proves the DIFFERENCE
    // rather than merely agreeing with itself. Without this the test would pass
    // against any rule that happens to return true, including the broken one.
    const oldRule = (r: ScenarioResult) =>
      r.blind === true || r.detail.startsWith("threw:") || /did not respond|gave up/i.test(r.detail);
    expect(oldRule(timedOutAndSaidSo), "the old rule caught this one — that part was never wrong").toBe(true);
    expect(
      oldRule(timedOutAndSaidSomethingElse),
      "the old rule already handled this verdict, so this guard proves nothing",
    ).toBe(false);

    // The one the old rule caught, and it still counts.
    expect(hostSeemsSick(timedOutAndSaidSo, true)).toBe(true);
    // The one it did not. A skip, no "threw:", and its wording matches neither
    // phrase the pattern looked for — but a deadline fired inside it, which is
    // the only thing being asked.
    expect(
      hostSeemsSick(timedOutAndSaidSomethingElse, true),
      "reset the breaker on a scenario that had just waited out the full budget",
    ).toBe(true);
    // And the rule stays narrow: without a deadline that same verdict is an
    // ordinary capability skip and says nothing about the host's health.
    expect(hostSeemsSick(timedOutAndSaidSomethingElse, false)).toBe(false);
    // A plain failure is the battery working, not a sick host.
    expect(hostSeemsSick({ name: "x", ok: false, detail: "the chart was gone", ms: 1 }, false)).toBe(false);
    // A blind deck scan still counts on its own.
    expect(hostSeemsSick({ name: "x", ok: false, blind: true, detail: "could not see the deck", ms: 1 }, false)).toBe(
      true,
    );
  });

  it("counts a bounded wait that hit its deadline, as a fact rather than a phrase", async () => {
    // The fact the rule above rests on. One deadline fired, one counted — if
    // this stops being true the breaker goes quiet again and nothing else says
    // so.
    const { deadlinesFired, _resetDeadlinesFiredForTest, slideCount, _setReadbackTimeoutForTest } =
      await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    _resetDeadlinesFiredForTest();
    _setReadbackTimeoutForTest(20);
    faults.wedgeAfterSyncs = 0;
    try {
      expect(deadlinesFired).toBe(0);
      // The rejection is the point; what is being counted is that it happened.
      await slideCount().catch(() => undefined);
      const after = (await import("../src/render/powerpoint")).deadlinesFired;
      expect(after, "a deadline fired and nothing counted it").toBeGreaterThan(0);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
      _resetDeadlinesFiredForTest();
    }
  });

  it("counts a blind scan apart from a host that lacks the API", async () => {
    // Both are skips and they are not the same finding. The summary called
    // every skip "host cannot run them", so a run in which the host refused
    // every deck read reported, in green, a list of capability gaps.
    const results: ScenarioResult[] = [
      { name: "alpha", ok: true, detail: "fine", ms: 1 },
      { name: "beta", ok: false, skipped: true, detail: "host has no picture fill", ms: 1 },
      { name: "gamma", ok: false, skipped: true, blind: true, detail: "the deck scan could not see", ms: 1 },
    ];
    const text = describeSelfTest(results);
    expect(text).toContain("1 skipped (host cannot run them)");
    // `blind` covers two host failures now — a scan that could not see, and a
    // scenario the host stopped answering — so the line no longer names one of
    // them for both. The 2026-08-08 `1fd6aa3` round said "the deck scan went
    // blind" about two scenarios that had timed out drawing, having scanned
    // nothing at all.
    expect(text).toMatch(/1 could not run — the host got in the way/);
    // …and says which, because a summary that reports a count without a name
    // is one somebody has to open the file to use.
    expect(text).toContain("gamma");
    // And the run is not green: a deck nobody could read is something to look at.
    expect(selfTestNeedsAttention(results)).toBe(true);
    expect(selfTestNeedsAttention(results.slice(0, 2)), "a plain capability skip is not a problem").toBe(false);
  });

  it("says which scenarios failed, not just how many", async () => {
    const results: ScenarioResult[] = [
      { name: "alpha", ok: true, detail: "fine", ms: 1 },
      { name: "beta", ok: false, detail: "broke", ms: 1 },
      { name: "gamma", ok: false, skipped: true, detail: "host cannot", ms: 1 },
    ];
    const text = describeSelfTest(results);
    expect(text).toContain("1 of 2 scenarios passed");
    expect(text).toContain("beta");
    // The skipped one is neither a pass nor a named failure.
    expect(text).not.toContain("gamma");
    expect(text).toContain("1 skipped");
  });

  it("runs to completion against a host at its worst", async () => {
    // The point of pinning it to the hostile profile: whatever the verdicts
    // are, the battery must survive stale proxies, hollow reads and a refused
    // addGroup well enough to report them. A battery that throws here would
    // be unusable on the one host it exists for.
    installHost([makeSlide("s1")]);
    applyWebProfile();
    const results = await runSelfTest("probe");
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    const named = byName(results);
    expect(named["insert on top of an earlier run"].detail).toBeTruthy();
  });
});

/**
 * The three scenarios added when `setSelectedShapes` turned out to exist.
 *
 * Each is asserted twice: that it PASSES on a well-behaved host, and that it
 * FAILS when the thing it exists to catch is actually broken. Only the second
 * half makes it a test. A scenario that cannot fail is a line in a report that
 * always says "ok", which is worse than no scenario at all — it is evidence,
 * and it is false.
 */
describe("the scenarios the selection API unlocked", () => {
  it("edits through the selection, and notices when the host selects the wrong shape", async () => {
    installHost([makeSlide("s1")]);
    const good = byName(await runSelfTest("probe"))["edit the chart the user selected"];
    expect(good.skipped, good.detail).toBeFalsy();
    expect(good.ok, good.detail).toBe(true);

    // A host that takes the call and selects something else. The pane would
    // then edit a chart the user did not click — the failure this scenario is
    // for, and one no shape count or tag read can see.
    //
    // It lands on the "did not read back as an SSF chart" branch rather than
    // the "read back a different chart" one, because the shape it wrongly
    // selects is a chart PART and parts carry no config tag. Both branches are
    // the same defect — the selection went to the wrong shape — and the message
    // is pinned so a later change cannot quietly turn this into a pass for an
    // unrelated reason.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.selectionIgnoresIds = true;
    const bad = byName(await runSelfTest("probe"))["edit the chart the user selected"];
    faults.selectionIgnoresIds = false;
    expect(bad.ok, `reported ok against a host that selects the wrong shape: ${bad.detail}`).toBe(false);
    expect(bad.skipped, "reported as skipped rather than broken").toBeFalsy();
    expect(bad.detail).toMatch(/did not read back as an SSF chart|read back a different chart/);
  });

  it("calls a host that stops answering after a select a known limitation, not a failure", async () => {
    // The third of PowerPoint-on-the-web's selection bugs, and the one that
    // cost a whole real-host round: `setSelectedShapes` is GA at PowerPointApi
    // 1.5 and TAKES the call, then leaves the selection subsystem unable to
    // answer anything. Measured on build 55011a3 — `getSelectedShapes` ran out
    // a 90-second budget, and the `setSelectedSlides` behind it did too.
    //
    // Two properties are pinned here, and the second is the one that matters:
    //
    //   1. It reports SKIPPED, not FAILED. Nothing in the add-in is broken by
    //      this — the pane never selects a shape from code, it reads the
    //      selection the user made with a click — so a red line here sends the
    //      next diagnosis after our own code, which is what happened.
    //   2. It COMES BACK, in about a budget rather than in minutes. The
    //      scenario asks for ten seconds instead of the ninety it would
    //      otherwise inherit, which is the difference between a battery a
    //      person will run again and one they will not.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.selectionWedgesHost = true;
    // The scenario's budget is capped by this one, so shortening it shortens
    // both — that cap is why a bounded wait is testable at all here.
    _setReadbackTimeoutForTest(200);
    const started = Date.now();
    let wedged: ScenarioResult;
    try {
      wedged = byName(await runSelfTest("probe", "edit the chart the user selected"))[
        "edit the chart the user selected"
      ];
    } finally {
      faults.selectionWedgesHost = false;
      _setReadbackTimeoutForTest(90_000);
    }
    expect(wedged.skipped, `reported a host limitation as our failure: ${wedged.detail}`).toBe(true);
    expect(wedged.detail).toContain("stopped answering selection calls");
    expect(wedged.detail).toMatch(/office-js#3083/);
    // Not the point of the test, but it is the reason the gate exists: at the
    // inherited budget this is three minutes of a person's evening.
    expect(Date.now() - started, "waited far past its own budget").toBeLessThan(5_000);
  });

  it("names the exact selection call the host stopped answering", async () => {
    // Two accounts, different culprits. This project measured
    // `setSelectedShapes([id])` being taken and everything after it going
    // silent; office-js#3698 says it is `setSelectedShapes([])` whose promise
    // never resolves. Another blind round would produce the same ambiguity at
    // the same cost, so the ladder climbs from the least invasive call to the
    // most and STOPS at the first silence — because after a wedge every call is
    // silent, and four timeouts name nothing.
    //
    // The fake wedges on the non-empty select, so that is the answer the report
    // must reach. Against a fake that wedged on the empty one it would have to
    // reach the other, which is the whole point of asking rather than assuming.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    setTracing(true);
    faults.selectionWedgesHost = true;
    _setReadbackTimeoutForTest(200);
    let r: ScenarioResult;
    try {
      r = byName(await runSelfTest("probe", "which selection call wedges the host"))[
        "which selection call wedges the host"
      ];
    } finally {
      faults.selectionWedgesHost = false;
      _setReadbackTimeoutForTest(90_000);
    }
    // An experiment, not an assertion: it ran, so it reports ok whatever it found.
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail, "the ladder did not name the call that went silent").toContain("SILENT at");
    // The rung AFTER the wedging call is where silence shows up — the write
    // itself is taken. Naming the read alone would be the old, useless answer;
    // naming the rung before it is what makes this an attribution.
    expect(r.detail).toContain("after selecting a shape");
    expect(r.detail, "did not say what the host last answered").toContain("the rung before it");

    // And it STOPPED there — asserted against the rungs the ladder actually
    // traced, not against the sentence it wrote. The first version of this
    // check looked for the absence of a later rung's NAME in the summary, and
    // that summary only ever mentions two rungs, so it passed just as happily
    // against a ladder that climbed all the way to the top. It was decoration.
    const rungs = traceLog()
      .entries.filter((e) => e.message === "ladder rung")
      .map((e) => e.data as { step: string; outcome: string });
    setTracing(false);
    expect(rungs.length, "the ladder recorded nothing").toBeGreaterThan(0);
    expect(
      rungs.filter((x) => x.outcome === "silent"),
      "climbed past the first silence",
    ).toHaveLength(1);
    expect(rungs.at(-1)?.outcome, "kept going after the host went quiet").toBe("silent");
  });

  it("runs the ladder in the routine battery, ahead of everything that selects a shape", async () => {
    // The ladder used to cost a separate five-minute round because it had to be
    // the ONLY thing that touched the selection subsystem. The property that
    // actually matters is narrower: it has to be the FIRST `setSelectedShapes`
    // in the run.
    //
    // Stated as that property rather than as an adjacency, because adjacency
    // was only ever a proxy for it — and the proxy broke the moment a second
    // shape-selecting scenario arrived, while the property did not.
    installHost([makeSlide("s1")]);
    const order = [...ROUTINE_SCENARIO_NAMES];
    const ladder = order.indexOf("which selection call wedges the host");
    expect(ladder, "the ladder is still a separate run").toBeGreaterThanOrEqual(0);
    for (const name of ["a selected shape survives an insert", "edit the chart the user selected"]) {
      expect(order.indexOf(name), `${name} selects a shape before the ladder does`).toBeGreaterThan(ladder);
    }
  });

  it("does not spend a second budget re-learning what the ladder just found", async () => {
    // Adjacency buys this too: on a wedged host the ladder has already produced
    // a better answer than `editViaSelection` can, one line earlier in the same
    // report. Paying a second budget for a worse version of it is the cost the
    // old ordering could not avoid, because the two never ran together.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.selectionWedgesHost = true;
    _setReadbackTimeoutForTest(200);
    let all: ScenarioResult[];
    try {
      all = await runSelfTest("probe");
    } finally {
      faults.selectionWedgesHost = false;
      _setReadbackTimeoutForTest(90_000);
    }
    const ladder = byName(all)["which selection call wedges the host"];
    const edit = byName(all)["edit the chart the user selected"];
    expect(ladder.detail, "the ladder did not find the wedge, so this proves nothing").toContain("SILENT at");
    expect(edit.skipped, "attributed a wedged host to the pane's own read").toBe(true);
    expect(edit.detail, "did not say the ladder had already answered it").toMatch(/ladder just found/);
    expect(edit.ms, "spent a budget on a question already answered").toBeLessThan(50);
  });

  it("says so plainly when nothing wedges at all", async () => {
    // The other half. A host that answers every rung is a real result — the
    // wedge is gone, or is not on this host — and a report that could only ever
    // say "SILENT at …" would be evidence that cannot be contradicted, which is
    // not evidence.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    const r = byName(await runSelfTest("probe", "which selection call wedges the host"))[
      "which selection call wedges the host"
    ];
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail).toContain("nothing wedged");
    expect(r.detail).not.toContain("SILENT");
  });

  it("reads the chart back from a REAL click, and takes its listener away after", async () => {
    // The one path a user travels that nothing can script. `setSelectedShapes`
    // is Office.js selecting a shape; a human clicking one is the same call in
    // theory and demonstrably not in practice on the web. So the battery asks
    // for a click and listens on `DocumentSelectionChanged` — a Common API
    // event that does not go through the wedging subsystem.
    //
    // What is asserted here is the half CI can honestly speak to: the
    // listen → click → read chain, and that the listener is taken back off.
    // Whether the EDIT that follows succeeds is `editViaSelection`'s ground,
    // covered twice over there; duplicating it here would only assert that the
    // fake host edits, which is not what this scenario is for.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    const run = runSelfTest("probe", "edit the chart YOU click");
    // Click only once the scenario is actually listening — the ordering that
    // exercises the handler rather than a lucky first read.
    await vi.waitFor(() => expect(selectionHandlerCount()).toBeGreaterThan(0));
    const chart = (await listChartsInDeck()).charts[0];
    expect(chart, "no probe chart for the click to land on").toBeTruthy();
    userClicksShape(chart.target.shapeId);
    const r = byName(await run)["edit the chart YOU click"];

    expect(r.skipped, `a click arrived and it still reported skipped: ${r.detail}`).toBeFalsy();
    expect(r.detail, "did not read the clicked chart back").toMatch(/read ".*" back from a real click/);
    // And it took its listener back off. A battery that leaked one handler per
    // run would leave dead runs answering the user's clicks.
    expect(selectionHandlerCount(), "left a selection handler behind").toBe(0);
  });

  it("does not blame the user for a click the host would not describe", async () => {
    // Two different findings that used to read identically. If the selection
    // read goes silent after a click, the person did their part and the HOST
    // did not — reporting "nobody clicked" tells them they failed at the one
    // thing they can see they did, which is how a report stops being believed.
    //
    // It also used to be reachable by timing alone: the read inherited the
    // 30-second CLICK budget, so a click at 29s started a read with 30s of rope
    // and the deadline cut it off one second later. Two unrelated quantities
    // sharing a name.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    _setClickWaitForTest(400);
    let r: ScenarioResult;
    try {
      const run = runSelfTest("probe", "edit the chart YOU click");
      // The probe charts only exist once the scenario is listening, and the
      // fault has to be armed AFTER reading them or it breaks the deck scan
      // rather than the selection read this test is about.
      await vi.waitFor(() => expect(selectionHandlerCount()).toBeGreaterThan(0));
      const chart = (await listChartsInDeck()).charts[0];
      // A click, on a host that then refuses the shape-collection read. This
      // is the "e.load is not a function" shape a real web host produced.
      faults.selectionReadThrows = true;
      userClicksShape(chart.target.shapeId);
      r = byName(await run)["edit the chart YOU click"];
    } finally {
      faults.selectionReadThrows = false;
      _setClickWaitForTest(30_000);
    }
    expect(r.skipped, "a host that went silent was reported as a pass or a failure").toBe(true);
    expect(r.detail, "blamed the user for the host's silence").not.toContain("nobody clicked");
    expect(r.detail).toContain("would not say what was selected");
    expect(selectionHandlerCount(), "left a selection handler behind").toBe(0);
  });

  it("skips rather than fails when nobody clicks", async () => {
    // Nobody clicked is "we did not check", never "it is broken". Reporting a
    // person's absence as a red line is how a battery teaches its reader to
    // ignore red lines.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    _setClickWaitForTest(60);
    let r: ScenarioResult;
    try {
      r = byName(await runSelfTest("probe", "edit the chart YOU click"))["edit the chart YOU click"];
    } finally {
      _setClickWaitForTest(30_000);
    }
    expect(r.ok, r.detail).toBe(false);
    expect(r.skipped, "a missing human was reported as a failure").toBe(true);
    expect(r.detail).toContain("nobody clicked");
    expect(selectionHandlerCount(), "left a selection handler behind after timing out").toBe(0);
  });

  it("reports a stop honestly, and will not call a leftover chart a clean stop", async () => {
    installHost([makeSlide("s1")]);
    const good = byName(await runSelfTest("probe"))["stop a run part-way"];
    expect(good.ok, good.detail).toBe(true);
    // The wording changed because the old one was FALSE. "stopped at a batch
    // boundary" implied it reached one; across 69 archived rounds this scenario
    // has committed zero batches, because `requestStop()` runs before the insert
    // and `throwIfStopped()` throws on iteration zero. The detail now counts the
    // commits instead of describing them.
    expect(good.detail).toContain("no shape was ever queued");

    // A chart already on the slide under the name a stopped insert would use.
    // The scenario must read that as "the stop left something behind" — if it
    // only checked that the throw was a stop, a host that stopped AND littered
    // would report a clean pass.
    vi.unstubAllGlobals();
    const slide = makeSlide("s1");
    const litter = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 10, height: 10 });
    litter.name = "PowerChart";
    litter.tagStore.set(CHART_TAG, JSON.stringify({ ...sampleConfig("clustered"), title: "probe stopped" }));
    slide.created.push(litter);
    installHost([slide]);
    const bad = byName(await runSelfTest("probe"))["stop a run part-way"];
    expect(bad.ok, `called a stop clean with a chart left behind: ${bad.detail}`).toBe(false);
    expect(bad.detail).toContain("left a re-editable chart");
  });

  it("stops a chart HALF DRAWN, and takes the wreckage with it", async () => {
    /**
     * THE CASE THE STOP BUTTON EXISTS FOR, AND THE ONE NEVER EXERCISED.
     *
     * `stop a run part-way` asks for the stop before the insert begins, so
     * `throwIfStopped()` throws on iteration zero and no shape is ever queued —
     * its own detail says "no shape was ever queued" in 251 of 322 archived
     * rounds, and "stopped at a batch boundary" in the other 71. A user pressing
     * Stop is some seconds into a draw, which is neither.
     *
     * This one requests the stop from the first `onPhase("commit", …)`, so a
     * batch of shapes has already landed when the abort happens.
     *
     * IT MUST COMMIT SOMETHING. A verdict that passes having drawn nothing is
     * the old scenario wearing a new name, so the detail has to show batches —
     * and when the insert never reaches a commit the scenario SKIPS rather than
     * passing, because it cannot answer its own question.
     */
    vi.unstubAllGlobals();
    // The DECK, held by reference. `installHost` mutates the array it is given
    // as slides are added and removed, which is what makes the cleanup
    // observable rather than merely claimed.
    const deck = [makeSlide("s1")];
    installHost(deck);
    const r = byName(await runSelfTest("probe", "stop a run mid-draw"))["stop a run mid-draw"];
    expect(r.skipped, `could not run at all: ${r.detail}`).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail, "passed without ever getting inside the draw").toMatch(/stopped inside the draw after [1-9]/);
    expect(r.detail, "did not say how much had landed").toMatch(/with [1-9]\d* shape\(s\) down/);
    expect(r.detail).toContain("the slide went with them");
    expect(r.detail).toContain("nothing left claiming to be a chart");
    /**
     * READ THE DECK, and be exact about what this can and cannot prove.
     *
     * Measured both ways on this fixture: when the slide delete succeeds the
     * deck ends at 7 slides, and with `refuseSlideDelete` it ends at 10 — the
     * scenario's own slide plus two the probe cleanup could not take either. So
     * the deck size DOES move with the cleanup, but it does not move by exactly
     * one, and pinning 7 would be pinning the probe's behaviour as much as this
     * scenario's.
     *
     * What it can say honestly is that the deck did not grow past what a working
     * run leaves. The failure path is covered separately and discriminates
     * properly: under `refuseSlideDelete` this scenario reports ok=false.
     *
     * Four earlier versions of this assertion proved nothing at all — live shape
     * counts (the fake never shrinks `created`), the installed slide (the draw
     * lands elsewhere), and growth across two runs (the probe setup adds more
     * each time). Each passed against a scenario that deleted nothing.
     */
    expect(
      deck.length,
      `the deck grew to ${deck.length} slides — the scenario's slide was not taken back`,
    ).toBeLessThan(10);
  });

  it("fails loudly when it cannot clear the wreckage, rather than reporting a clean stop", async () => {
    /**
     * THE CLEANUP IS THE POINT OF THIS SCENARIO, so a cleanup that does not
     * happen has to be a failure and not a footnote. A partly drawn chart left
     * in the deck is read by `same scale across the deck` as part of its chart
     * population, which is precisely the contamination this design avoids.
     *
     * `refuseSlideDelete` is the fault that matters now that the slide is the
     * unit of cleanup. The first version of this scenario cleaned up by deleting
     * SHAPES by id, and round 353 showed what that is worth on a real host: 10
     * of 10 refused, ten shapes left in the deck. Ids read from a slide listing
     * are not ids this host accepts in `getItemOrNullObject`, which is written
     * down in `host-probe.ts` and which I did not join up until the round said
     * so.
     */
    installHost([makeSlide("s1")]);
    faults.refuseSlideDelete = true;
    try {
      const r = byName(await runSelfTest("probe", "stop a run mid-draw"))["stop a run mid-draw"];
      expect(r.ok, "reported a clean stop while its slide was still in the deck").toBe(false);
      expect(r.detail, "did not say the slide survived").toMatch(/could not be deleted|still holds/);
    } finally {
      faults.refuseSlideDelete = false;
    }
  });

  it("finds the slide it added by what the deck GAINED, once the id has changed under it", async () => {
    /**
     * THE LAST THING KEEPING THESE TWO SCENARIOS RED, and it only became
     * fixable on 2026-09-03.
     *
     * The id `addSlideForChart` returns is the add-time one; this host re-keys
     * a slide once it settles. So by cleanup time the deck holds the slide
     * under a name nobody upstream has: `deleteSlideById` cannot resolve it,
     * the positional fallback finds `indexOf === -1`, and it correctly refuses
     * to remove an index whose id does not match. That refusal must stay —
     * round 124 reported 62 scratch slides swept and deleted none.
     *
     * `deleteSlideAddedSince` sidesteps the name entirely: it diffs the deck
     * listing at DELETE time, so whatever the slide is called now, it is the id
     * that is there and was not before.
     *
     * Before the slide-master fix this could not have worked at all — the added
     * slide was rejected by the server and rolled back, so there was nothing to
     * find. `newSlideIdSettlesAfter` is what makes the fake re-key the way the
     * host does.
     */
    installHost([makeSlide("s1")]);
    faults.newSlideIdSettlesAfter = 1;
    try {
      const r = byName(await runSelfTest("probe", "stop a run mid-draw"))["stop a run mid-draw"];
      expect(
        r.detail,
        "the slide was left in the deck even though the deck was listing it under a new name",
      ).not.toMatch(/still in the deck/);
    } finally {
      faults.newSlideIdSettlesAfter = null;
    }
  });

  it("does not invent wreckage when the delete was refused because the slide had already gone", async () => {
    /**
     * THE OTHER HALF OF THE SAME QUESTION, and the scenario got it wrong until
     * round 370 showed the case existed.
     *
     * At 4:3 the slide this scenario adds is listed, reported `landed=1`, and
     * then not there. `deleteSlideById` is left holding an id the deck no longer
     * lists; it correctly refuses to remove an index whose id does not match and
     * returns false. The verdict read that `false` as "a slide is stuck in the
     * deck" and announced "the aborted draw is still in the deck" about a deck
     * that had never been cleaner.
     *
     * Wrong in the one direction a self-test must never be wrong. A missed
     * problem costs a round; an INVENTED one sends someone looking for wreckage
     * that was never there, and this project has burnt days on exactly that.
     *
     * So the DECK decides, not the return value. The fault reproduces a pairing
     * this fake could not previously express — refused, AND gone.
     */
    installHost([makeSlide("s1")]);
    faults.slideVanishesInsteadOfDeleting = true;
    try {
      const r = byName(await runSelfTest("probe", "stop a run mid-draw"))["stop a run mid-draw"];
      expect(
        r.detail,
        "claimed the aborted draw was still in the deck, about a slide the host had already dropped",
      ).not.toMatch(/still in the deck/);
      /**
       * AND ASSERTED POSITIVELY, because the line above passes for the wrong
       * reason on its own: a clean run does not say "still in the deck" either,
       * so it stays green even when the fault does nothing. The mutation run
       * proved that — disarming the fake's drop left it passing.
       *
       * This phrase appears only when `deleteSlideById` returned FALSE, so it
       * is the one observable separating "refused, and gone" from "deleted
       * normally" — and it is what a reader chasing the own-slide defect needs.
       */
      expect(r.detail, "the scenario did not report that the delete had been refused").toMatch(
        /the delete was refused, but the deck shrank/,
      );
      expect(r.ok, "a clean deck was reported as a failure").toBe(true);
    } finally {
      faults.slideVanishesInsteadOfDeleting = false;
    }
  });

  it("will not call a mid-draw stop clean when the wreckage claims to be a chart", async () => {
    // The half that matters. Partial shapes are allowed — the pane's own message
    // says "anything already drawn was kept". What is not allowed is half a
    // chart carrying a config tag, because the pane would then offer to re-edit
    // something that was never finished and Same Scale would rescale it as
    // though it were whole.
    vi.unstubAllGlobals();
    const slide = makeSlide("s1");
    const litter = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 10, height: 10 });
    litter.name = "PowerChart";
    litter.tagStore.set(CHART_TAG, JSON.stringify({ ...sampleConfig("clustered"), title: "probe mid-draw" }));
    slide.created.push(litter);
    installHost([slide]);
    const bad = byName(await runSelfTest("probe", "stop a run mid-draw"))["stop a run mid-draw"];
    expect(bad.ok, `called a mid-draw stop clean with a chart left behind: ${bad.detail}`).toBe(false);
    expect(bad.detail).toContain("left a re-editable chart behind");
  });

  it("sees the chart on the slide, and fails when the host renders the same bytes regardless", async () => {
    installHost([makeSlide("s1")]);
    const good = byName(await runSelfTest("probe", VISIBLE))[VISIBLE];
    expect(good.skipped, good.detail).toBeFalsy();
    expect(good.ok, good.detail).toBe(true);

    // A host whose rasteriser answers the same thing for an empty slide and a
    // slide with a chart on it. The comparison is then worthless, and saying so
    // is the only honest verdict — this is the one scenario whose whole value
    // is the difference between two images.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.constantSlideImage = true;
    const bad = byName(await runSelfTest("probe", VISIBLE))[VISIBLE];
    faults.constantSlideImage = false;
    expect(bad.ok, `passed while every render was identical: ${bad.detail}`).toBe(false);
    expect(bad.detail).toContain("nothing is visible");
  });

  /**
   * The failure this scenario actually hit in a real PowerPoint, on 2026-08-04.
   *
   * It got as far as "rasterising the empty slide" and the host answered
   * `GeneralException`, `errorLocation: SlideCollection.getItem`, `statement:
   * var slide = slides.getItem(...); slide.getImageAsBase64(...)`. The slide
   * was one `addScratchSlide` had just made and whose liveness check had just
   * passed — through the very same proxy, one sync earlier. Resolving a proxy
   * is what makes Office.js rewrite its object path to `getItem(id)`, and a new
   * slide's id does not round-trip through `getItem` on the web: the rule
   * `SlideThunk` is built on, in a function that held one handle across two
   * syncs.
   */
  it("rasterises a slide the host will only name once", async () => {
    installHost([makeSlide("s1")]);
    const r = byName(await runSelfTest("probe", VISIBLE))[VISIBLE];
    expect(r.skipped, `gave up on a host that answers by-id handles once: ${r.detail}`).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
  });

  /**
   * The healthy-host half of the blind-scan change, and ONLY that.
   *
   * A real host reported this scenario FAILED — *"the chart already on the
   * slide is no longer re-editable"* — on a build whose insert path had not
   * changed since the run where it passed. The scenario checked its first deck
   * scan for blindness and then made its most alarming claim off a second one
   * it never checked, so "your chart is gone" and "I could not look" came out
   * as the same sentence. It now skips when either scan is blind.
   *
   * WHAT THIS TEST DOES NOT DO: fail against the pre-fix code. Reproducing the
   * real sequence needs the first scan to answer and a later one to go short,
   * and `hollowReads` blinds from the first read — so the scenario's own
   * opening check catches it and the new rung is never reached. `hollowReadsAfter`
   * was added to model a host that degrades mid-run, which is what the real one
   * did, but landing it on exactly the right read means guessing a scan count
   * three scenarios deep — the fragile magic number CLAUDE.md warns about, and
   * a guard tuned until it goes red is not evidence.
   *
   * So this pins the half that IS checkable: the change must not cost the
   * healthy-host case its pass. The blind case is argued from the host's own
   * log, and is honestly unguarded until something can drive that window.
   */
  it("still passes on a host whose deck scans answer", async () => {
    installHost([makeSlide("s1")]);
    const r = byName(await runSelfTest("probe", "insert onto a slide that already has content"));
    const verdict = r["insert onto a slide that already has content"];
    expect(verdict.ok, `the healthy-host case stopped working: ${verdict.detail}`).toBe(true);
    expect(verdict.skipped, "skipped a scenario whose scans were fine").toBeFalsy();
  });

  /**
   * A rasteriser that answers with nothing has to SAY so.
   *
   * `slideImageBase64` has four ways of returning undefined — no 1.8, the
   * slide would not resolve, the call threw, the call came back empty — and
   * one `catch` covering all of them. A real run showed how little that is
   * worth: the visibility scenario went from "rasterising the empty slide"
   * straight to "removing the scratch slide", with no error line and no
   * drawing step, and the reason existed nowhere. The round before, the same
   * call at least reported `GeneralException`; a fresh slide handle stopped it
   * throwing without making it work, and traded a loud wrong answer for a
   * silent one. Both are failures. Only one can be diagnosed.
   */
  /**
   * The visibility chart fits in ONE batch, and that is the whole reason the
   * scenario ever runs.
   *
   * It completed one round in nine. The other eight skipped on
   * `PowerPoint did not respond while drawing shapes 1-10 of 24 (45s)` — this
   * host stalls the first batch of a draw often enough that needing three of
   * them is a coin flip repeated, and `sampleConfig("clustered")` is 24 shapes
   * however small the frame it is drawn into. The scenario's own note had
   * already reached for "draws fewer shapes" and only shrank the box.
   *
   * The counterbalanced scenario made this trade first and completes eight
   * rounds in nine. Nothing here is about density: the question is whether the
   * slide's PICTURE changes when a chart is drawn on it.
   *
   * Pinned on the batch TOTAL rather than the config, because that is the
   * property that matters — a future edit that restores the full sample, or
   * adds a series, costs nine more rounds of skips before anyone notices.
   */
  it("draws a chart small enough to land in a single batch", async () => {
    installHost([makeSlide("s1")]);
    setTracing(true);
    await runSelfTest("probe", "the chart is actually visible");
    const draws = traceLog()
      .entries.filter((e) => e.scope === "draw" && e.message === "batch issued")
      .map((e) => (e.data as { total?: number }).total ?? 0);
    expect(draws.length, "the scenario drew nothing at all — this no longer tests what it says").toBeGreaterThan(0);
    expect(
      Math.max(...draws),
      `the visibility chart needs ${Math.max(...draws)} shapes, so more than one batch`,
    ).toBeLessThanOrEqual(SHAPES_PER_SYNC);
  });

  it("says why it could not rasterise, instead of quietly reporting nothing", async () => {
    installHost([makeSlide("s1")]);
    setTracing(true);
    // The host takes the call and hands back an empty image — the quiet case,
    // which is the one that had no line at all.
    faults.constantSlideImage = false;
    faults.emptySlideImage = true;
    try {
      const r = byName(await runSelfTest("probe", "the chart is actually visible"))["the chart is actually visible"];
      expect(r.ok, "passed with no image to compare").toBe(false);
      const said = traceLog()
        .entries.filter((e) => e.scope === "host")
        .map((e) => e.message);
      expect(said, `nothing in the log says why: ${said.join(" | ")}`).toContain(
        "the host took the rasterise and returned nothing",
      );
    } finally {
      faults.emptySlideImage = false;
    }
  });

  /**
   * A real host died inside this scenario and the log's last line was the
   * scenario announcing itself — five host calls, and nothing to say which one
   * was outstanding. The scenario-level announcement was itself added for
   * exactly this reason one round earlier; it turned "somewhere in the battery"
   * into "somewhere in this scenario", and stopped there.
   *
   * The rasteriser is held open rather than made to fail, because a call that
   * throws leaves a verdict and a call that never returns leaves only the log.
   * Only the second one is the case this line exists for.
   */
  it("takes no scratch slide anywhere in the battery", async () => {
    // The rule the real host taught, twice over, and the second time by name:
    // `chartIsVisible` died rasterising a slide added 0.3s earlier, and
    // `degradesOverTime` died on its SECOND `slides.add()` 0.4s after its
    // first. Both stopped taking scratch slides, and no scenario should take
    // one again without a reason better than convenience.
    //
    // Asserted against the DECK, so it holds however the scenarios are written:
    // a full round adds slides only through the two inserts at its head.
    const { slideCount } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    _setDegradeSizeForTest(2, 2);
    try {
      await runSelfTest("probe", "two slides claiming one slot");
      const justTheInserts = await slideCount();

      installHost([makeSlide("s1")]);
      await runSelfTest("probe", "what makes a long run slow down");
      expect(await slideCount(), "the degradation experiment took a slide of its own").toBe(justTheInserts);
    } finally {
      _setDegradeSizeForTest(8, 12);
    }
  }, 60_000);

  it("calls a host that stopped answering a skip, not a failed scenario", async () => {
    // The 2026-08-08 `a546897` round. `the chart is actually visible` ran
    // eleventh, ten minutes in, and the host stalled 45s on its first draw
    // batch — so the battery reported `FAILED — threw: PowerPoint did not
    // respond while drawing shapes 1-10 of 24 (45s)`. Picked alone at 61
    // seconds, the same build PASSED. That verdict is about a fatigued host,
    // said in the words of a broken chart, and it is the exact distinction —
    // "we did not check" against "we checked and it is broken" — that the rest
    // of this file is built on.
    installHost([makeSlide("s1")]);
    const { _setBatchTimeoutForTest } = await import("../src/render/powerpoint");
    _setReadbackTimeoutForTest(20);
    _setBatchTimeoutForTest(20);
    faults.wedgeAfterSyncs = 4;
    try {
      const r = byName(await runSelfTest("probe", "edit a chart on the visible slide"))[
        "edit a chart on the visible slide"
      ];
      expect(r.detail, "a silent host was reported as a scenario failure").not.toMatch(/^threw:/);
      expect(r.skipped, r.detail).toBe(true);
      // …and still surfaced, because nothing was learned and the round should
      // say so rather than read as a clean pass.
      expect(r.blind, r.detail).toBe(true);
      expect(selfTestNeedsAttention([r])).toBe(true);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
      _setBatchTimeoutForTest(45_000);
    }
  }, 60_000);

  it("names the degradation experiment's host calls too", async () => {
    // The same lesson, learned twice. On 2026-08-08 this scenario was picked
    // alone, announced itself at 26.9s, and took the tab — with no finer step
    // than its own name, so "taking a scratch slide" and "drawing ninety-six
    // shapes onto a slide the run just added" both fit and neither could be
    // ruled out. That is precisely the state wrapping `chartIsVisible` got it
    // out of, and there was no reason for a second scenario to pay for it.
    installHost([makeSlide("s1")]);
    _setDegradeSizeForTest(2, 2);
    setTracing(true);
    try {
      await runSelfTest("probe", "what makes a long run slow down");
      const steps = traceLog()
        .entries.filter((e) => e.message === "degradation step")
        .map((e) => e.data?.what);
      // The two scratch adds are gone, and their absence is the finding: named
      // separately, they told a real host exactly which one it dies on — the
      // SECOND, four tenths of a second after the first — and the scenario
      // stopped taking them. The arms are what is left to name.
      expect(steps).toEqual(["counting the deck", "timing the one-context arm", "timing the fresh-context arm"]);
    } finally {
      setTracing(false);
      _setDegradeSizeForTest(8, 12);
    }
  }, 60_000);

  it("names the host call it is on, so a scenario that never ends is not a mystery", async () => {
    setTracing(true);
    let release!: () => void;
    try {
      // After installHost — it resets every fault, so arming one before it is
      // arming nothing.
      installHost([makeSlide("s1")]);
      faults.slideImageGate = new Promise<void>((r) => (release = r));
      const run = runSelfTest("probe", "the chart is actually visible");
      await vi.waitFor(() => {
        const last = traceLog()
          .entries.filter((e) => e.scope === "selftest")
          .at(-1);
        expect(last?.message).toBe("visibility step");
        expect(last?.data?.what).toBe("rasterising a slide that already existed");
      });
      release();
      await run;
      // And every call is named, in the order the scenario makes them.
      const steps = traceLog()
        .entries.filter((e) => e.message === "visibility step")
        .map((e) => e.data?.what);
      // No scratch add and no delete any more: both were calls that killed a
      // real PowerPoint, and neither was needed for the comparison.
      expect(steps).toEqual([
        "rasterising a slide that already existed",
        // The control: the same slide again with nothing drawn between. Without
        // it a before/after difference cannot tell a chart from a rasteriser
        // that does not answer the same way twice — and the two are
        // indistinguishable in a verdict.
        "rasterising the same slide a second time",
        "drawing the chart",
        "rasterising the slide with the chart",
      ]);
    } finally {
      release?.();
      faults.slideImageGate = null;
      setTracing(false);
    }
  });

  it("never asks the host to rasterise a slide it just added", async () => {
    // The call that killed PowerPoint on the web five rounds running, and the
    // fifth round is the one that pinned it: picked alone, the scenario was
    // reached at 61.5s with only its two inserts in front of it, took a scratch
    // slide at 61.5s, logged `rasterising the empty slide` at 61.8s, and the
    // tab died. The four rounds before died ten minutes and nine scenarios in,
    // so elapsed time and volume of drawing were live explanations until then.
    //
    // Asserted against the DECK rather than the trace, so it holds however the
    // steps are named: the scenario must not grow the deck at all.
    // A slide count cannot see this: the old code deleted its scratch slide
    // afterwards, so the deck ended the size it started. What the surface WAS
    // is the observable difference, and the fake's raster payload carries it —
    // an empty scratch slide reports `shapes=0`.
    installHost([makeSlide("s1")]);
    const visible = byName(await runSelfTest("probe", VISIBLE))[VISIBLE];
    expect(visible.skipped, visible.detail).toBeFalsy();
    expect(rasterised.length, "the scenario never rasterised anything").toBeGreaterThan(0);
    const empty = rasterised.filter((p) => /:shapes=0:/.test(p));
    expect(empty, `it rasterised a slide with nothing on it: ${empty.join(" | ")}`).toEqual([]);
  });

  it("runs one scenario plus the two inserts it needs, when asked for one", async () => {
    // A full round costs minutes and leaves slides behind: the first run that
    // produced a usable trace took 496 seconds to reach scenario seven and
    // wedged there, and the scenario before it left four charts on the deck as
    // loose piles of shapes. Iterating on the seventh by running the six in
    // front of it is not iteration.
    //
    // The first two run regardless, because every other scenario needs the
    // probe charts they insert — a targeted run that found none would report
    // SKIPPED, which is honest and useless.
    installHost([makeSlide("s1")]);
    const results = await runSelfTest("probe", "the chart is actually visible");
    expect(results.map((r) => r.name)).toEqual([
      "insert on top of an earlier run",
      "two slides claiming one slot",
      "the chart is actually visible",
    ]);
    const picked = results[2];
    expect(picked.skipped, picked.detail).toBeFalsy();
    expect(picked.ok, picked.detail).toBe(true);
  });

  it("offers exactly the scenarios it can run, in run order", async () => {
    // The picker is filled from this list. A hand-kept copy beside it would be
    // one rename away from offering a scenario that does not exist.
    installHost([makeSlide("s1")]);
    const results = await runSelfTest("probe");
    // A full run is the routine list exactly, in order.
    expect(results.map((r) => r.name)).toEqual([...ROUTINE_SCENARIO_NAMES]);
    // The picker offers those PLUS the ones a full run deliberately leaves out,
    // because being pickable is the only way such a scenario can ever run.
    expect(SCENARIO_NAMES.slice(0, ROUTINE_SCENARIO_NAMES.length)).toEqual([...ROUTINE_SCENARIO_NAMES]);
    for (const name of ROUTINE_SCENARIO_NAMES) expect(SCENARIO_NAMES).toContain(name);
    expect(SCENARIO_NAMES.length).toBeGreaterThan(ROUTINE_SCENARIO_NAMES.length);
  });

  it("runs everything when nothing is picked", async () => {
    installHost([makeSlide("s1")]);
    expect(await runSelfTest("probe", "")).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
  });

  it("announces a scenario BEFORE running it, so a crash names the right one", async () => {
    // Every verdict this battery emits is past tense. A run that dies mid
    // scenario therefore leaves the PREVIOUS scenario's line as its last word
    // — off by one, and pointing at the one thing that demonstrably worked.
    // That is not hypothetical: two real-host rounds have been diagnosed from
    // a screenshot of the pane rather than a log, because the log only exists
    // once a run ends and neither run ended.
    setTracing(true);
    try {
      installHost([makeSlide("s1")]);
      await runSelfTest("probe");
      const entries = traceLog().entries.filter((e) => e.scope === "selftest");
      const first = entries[0];
      expect(first?.message, "the first thing said about a scenario was its verdict").toBe("scenario starting");
      expect(first?.data?.name).toBe("insert on top of an earlier run");
      // And one announcement per scenario, each before its own verdict.
      const starts = entries.filter((e) => e.message === "scenario starting");
      expect(starts).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
      for (const s of starts) expect(s.data?.name, "an announcement with no name").toBeTruthy();
    } finally {
      setTracing(false);
    }
  });

  it("can be stopped, and says which scenarios were never reached", async () => {
    // The battery had no stop check of its own. A user who pressed Stop got a
    // button that said "Stopping…" and a run that carried on to the next
    // scenario regardless — observed on a real host at 1819 seconds, where the
    // only way out was closing the tab, which is also how that run's log was
    // lost.
    //
    // Not-reached is reported as SKIPPED with a reason, never dropped: a report
    // missing its last four lines looks like a battery that crashed, and that
    // is a different diagnosis from one that was stopped.
    installHost([makeSlide("s1")]);
    requestStop();
    let results: ScenarioResult[];
    try {
      results = await runSelfTest("probe");
    } finally {
      resetStop();
    }
    expect(results, "a stopped battery dropped scenarios instead of reporting them").toHaveLength(
      ROUTINE_SCENARIO_NAMES.length,
    );
    expect(
      results.every((r) => r.skipped),
      "a stopped battery ran a scenario anyway",
    ).toBe(true);
    expect(results[0].detail).toMatch(/stopped/i);
  });

  it("does not let the stop scenario clear a stop the USER asked for", async () => {
    // `stopPartWay` is the only code in the add-in that calls resetStop(). If
    // it runs while the user's own stop is pending, its finally clears the
    // flag — and the battery resumes after the user cancelled it. The guard
    // above makes this unreachable; this pins it anyway, because "unreachable"
    // is how the flag came to be clobbered in the first place.
    installHost([makeSlide("s1")]);
    requestStop();
    try {
      await runSelfTest("probe");
      expect(isStopRequested(), "the stop scenario cleared the user's stop").toBe(true);
    } finally {
      resetStop();
    }
  });

  it("skips rather than fails on a host below PowerPointApi 1.5", async () => {
    // "We could not check" and "we checked and it is broken" are different
    // answers, and reporting the first as the second is how a requirement-set
    // gap gets diagnosed as a bug.
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.5");
    const r = byName(await runSelfTest("probe"))["edit the chart the user selected"];
    expect(r.skipped, r.detail).toBe(true);
    expect(r.detail).toContain("1.5");
  });
});

describe("the self-test against a host that can take a generated deck", () => {
  it("actually runs the insert scenarios, and they pass on a well-behaved host", async () => {
    // Until the fake could take a .pptx these two always reported "skipped",
    // so the battery's own insert logic had never executed anywhere. A skip is
    // an honest answer to give a user and a useless one to give a test suite.
    installHost([makeSlide("s1")]);
    const named = byName(await runSelfTest("probe"));
    const twice = named["insert on top of an earlier run"];
    expect(twice.skipped, twice.detail).toBeFalsy();
    expect(twice.ok, twice.detail).toBe(true);
    expect(twice.detail).toContain("deck grew by 4");

    const dup = named["two slides claiming one slot"];
    expect(dup.skipped, dup.detail).toBeFalsy();
    expect(dup.ok, dup.detail).toBe(true);
  });

  it("notices when the host quietly drops half the deck", async () => {
    // The negative control. A run that inserts four slides and gets two must
    // not read as a pass — that is the exact failure the scenario exists to
    // catch, and it is invisible to a check that only asks whether the call
    // threw. A real 12-item file insert came back "11 of 12 complete · 1 lost"
    // with nothing raised at all.
    installHost([makeSlide("s1")]);
    expect(byName(await runSelfTest("probe"))["insert on top of an earlier run"].ok).toBe(true);

    installHost([makeSlide("s1")]);
    faults.swallowDecks = 1; // the host takes the first deck and lands nothing
    const after = byName(await runSelfTest("probe"))["insert on top of an earlier run"];
    expect(after.ok).toBe(false);
    expect(after.detail).toContain("deck grew by 2");
  });
});

describe("scenarios that must not be able to pass without proving anything", () => {
  const named = (rs: ScenarioResult[]) => Object.fromEntries(rs.map((r) => [r.name, r]));

  it("skips the picture scenario rather than faking one when there is no rasteriser", async () => {
    // `render: "image"` on a config does NOT make a picture — the renderer
    // takes the picture path only when handed `pictureBase64`. An earlier
    // version passed `undefined`, quietly performed two ordinary shape
    // updates, and reported that the picture round-trip worked. A scenario
    // that cannot fail is worse than one that is missing: it reads as
    // evidence. With no rasteriser the honest answer is "did not check".
    setSelfTestRasterizer(undefined as unknown as (s: unknown) => Promise<string | undefined>);
    installHost([makeSlide("s1")]);
    const explode = byName(await runSelfTest("probe"))["explode a degraded picture"];
    expect(explode.skipped, explode.detail).toBe(true);
    expect(explode.detail).toMatch(/rasteris/i);
  });

  it("hands the host a real picture when there is one", async () => {
    // The deck, not one slide: the insert scenarios append their own slides,
    // so the chart this one collapses is never on the slide we started with.
    const deck = [makeSlide("s1")];
    installHost(deck);
    let asked = 0;
    setSelfTestRasterizer(async () => {
      asked++;
      return "data:image/png;base64,UE5H";
    });
    const explode = named(await runSelfTest("probe"))["explode a degraded picture"];
    expect(asked, "the scenario never asked for a picture").toBeGreaterThan(0);
    expect(explode.skipped, explode.detail).toBeFalsy();
    // The picture really reached the host, rather than the renderer drawing
    // shapes and the scenario calling that a picture.
    const pictured = deck.flatMap((s) => s.created).filter((s) => s.imageBase64);
    expect(pictured.length, "no picture fill ever reached the host").toBeGreaterThan(0);
  });

  it("fails the picture scenario when the host quietly drew shapes instead", async () => {
    // `canInsertPicture` is a requirement-set check, so a host that advertises
    // 1.8 and then refuses the fill passes the scenario's own gate — the
    // renderer logs PC-IMG-REFUSED and draws native shapes. Everything after
    // that is a shape-to-shape update being reported as a picture round-trip,
    // which is the same "cannot fail" trap as the `undefined` png above, one
    // step later. It was only traced, so the scenario still said "collapsed to
    // a picture and exploded back, config intact".
    const deck = [makeSlide("s1")];
    installHost(deck);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    faults.refusePictureFill = true;
    try {
      const explode = byName(await runSelfTest("probe"))["explode a degraded picture"];
      expect(explode.skipped, explode.detail).toBeFalsy();
      expect(explode.ok, explode.detail).toBe(false);
      expect(explode.detail).toMatch(/a picture is one/);
    } finally {
      faults.refusePictureFill = false;
    }
  });

  it("does not blame the picture for the charts sitting next to it", async () => {
    // The false alarm the first real round produced. The check counted the
    // SLIDE's shapes and demanded exactly one, and the battery deliberately
    // piles charts onto a shared slide two scenarios earlier — so a perfectly
    // good picture failed with "the slide holds 3 shapes after the collapse".
    // `a selected shape survives an insert` reports the same three in the same
    // round, which is what identified them as neighbours rather than wreckage.
    //
    // Here the neighbours are put there outright: shapes already on the slide
    // that no part of this scenario touches.
    vi.unstubAllGlobals();
    const slide = makeSlide("s1");
    for (const name of ["Logo", "Footnote", "Rectangle 4"]) {
      const neighbour = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 10, height: 10 });
      neighbour.name = name;
      slide.created.push(neighbour);
    }
    installHost([slide]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    const explode = named(await runSelfTest("probe"))["explode a degraded picture"];
    expect(explode.detail, "a neighbouring shape was counted as a broken picture").not.toMatch(/a picture is one/);
    expect(explode.ok, explode.detail).toBe(true);
  });

  it("gives every scenario that shares a slide its own place on it", () => {
    // Three rounds of the same complaint from the owner, looking at the deck
    // the battery left him. The third was a 4:3 deck, where the first fix — a
    // fixed column down the right-hand edge — put charts at x=498 across one
    // running to x=600. Its own test had said the clearance could not hold at
    // 720pt wide, which was honest and did not stop it happening.
    //
    // The column's position was GUESSED: it assumed the chart began at x=60 and
    // the host had used 120. Placing from the box the caller actually holds is
    // what removes the guess, so the property is now checkable on any deck.
    const decks = [
      { width: 960, height: 540 }, // 16:9, as observed
      { width: 720, height: 540 }, // 4:3, as observed
    ];
    // The chart already on the slide, at the origins each deck really used.
    const occupied = [
      { left: 60, top: 90, width: 480, height: 300 },
      { left: 120, top: 120, width: 480, height: 300 },
    ];
    for (const [d, slide] of decks.entries()) {
      const chart = occupied[d];
      const slots = Array.from({ length: SIDE_SLOTS }, (_, n) => sideSlot(n, slide, chart));
      const apart = (a: typeof chart, b: typeof chart) =>
        a.left + a.width <= b.left ||
        b.left + b.width <= a.left ||
        a.top + a.height <= b.top ||
        b.top + b.height <= a.top;
      for (const [i, box] of slots.entries()) {
        expect(box.left, `slot ${i} runs off the left of a ${slide.width}pt deck`).toBeGreaterThanOrEqual(0);
        expect(box.left + box.width, `slot ${i} runs off the right`).toBeLessThanOrEqual(slide.width);
        expect(box.top, `slot ${i} starts above the slide`).toBeGreaterThanOrEqual(0);
        expect(box.top + box.height, `slot ${i} runs off the bottom`).toBeLessThanOrEqual(slide.height);
        expect(box.height, `slot ${i} has no room to draw in`).toBeGreaterThan(0);
        // Clear of the chart that is already there — on BOTH deck shapes now.
        // The old rule could not promise this at 720pt and said so; this one
        // can, because it is told where the chart is instead of assuming.
        expect(apart(box, chart), `slot ${i} overlaps the chart already on a ${slide.width}pt deck`).toBe(true);
      }
      for (const [i, a] of slots.entries())
        for (const b of slots.slice(i + 1))
          expect(apart(a, b), `slots overlap on a ${slide.width}x${slide.height} deck`).toBe(true);
      // And the measurement grid fits in the slot it was given. It used to be
      // placed by hand at (20, 430) under a comment asserting it cleared the
      // old right-hand column — which it did, and which stopped being the
      // question the moment the slots became a band along the bottom. Checked
      // against the renderer's own constants so the two cannot drift.
      const grid = slots[GRID_SLOT];
      const need = gridFootprint(96);
      expect(
        need.width,
        `the 96-shape grid does not fit across its slot on a ${slide.width}pt deck`,
      ).toBeLessThanOrEqual(grid.width);
      expect(
        need.height,
        `the 96-shape grid does not fit down its slot on a ${slide.width}pt deck`,
      ).toBeLessThanOrEqual(grid.height);
    }
  });

  it("keeps a FIXED-size box on the slide, which the slot's own top does not", () => {
    /**
     * THE TEST ABOVE CHECKS THE SLOT, AND THE SLOT WAS NEVER THE PROBLEM.
     *
     * `sideSlot` clamps its height with `Math.max(1, …)`, so the band it returns
     * fits by construction and `box.top + box.height <= slide.height` passes no
     * matter how impossible that band is. What is DRAWN in the band is not
     * checked by it at all — a gate only as wide as its sweep.
     *
     * `rotatedShapePlacement` keeps its own 160x220 on purpose: the segments
     * must be steep enough that a length and a horizontal extent cannot be
     * confused, which makes the size an invariant of the measurement rather than
     * a property of the deck. Under the 300pt chart these decks really carry the
     * band is 160pt at its tallest — so a 220pt box placed at the band's top
     * hung 40 to 200pt off the bottom of the slide in every archived round, at
     * both aspect ratios, and nothing said a word.
     */
    const decks = [
      { width: 960, height: 540 }, // 16:9, as observed
      { width: 720, height: 540 }, // 4:3, as observed
    ];
    const size = { width: 160, height: 220 };
    for (const slide of decks) {
      // Sweeping the chart DOWN the slide: the lower it sits the shorter the
      // band, and the further a fixed box would hang off the bottom.
      for (const top of [40, 90, 120, 150, 200, 260]) {
        const chart = { left: 60, top, width: 480, height: 300 };
        const slot = sideSlot(3, slide, chart);
        const placed = fitInSlot(slot, size, slide);
        const where = `${slide.width}pt deck, chart at top=${top}`;
        expect(placed.top, `box starts above the slide on a ${where}`).toBeGreaterThanOrEqual(0);
        expect(placed.top + size.height, `box hangs off the bottom on a ${where}`).toBeLessThanOrEqual(slide.height);
        expect(placed.left + size.width, `box hangs off the right on a ${where}`).toBeLessThanOrEqual(slide.width);
        // Still as low as it honestly can be: this is a fit, not a jump to the
        // top of the slide, so it keeps clear of the chart above wherever the
        // band genuinely has room.
        expect(placed.top, `box was not placed in the band on a ${where}`).toBeLessThanOrEqual(slot.top);
      }
      // The control: where the band DOES have room the box is left exactly where
      // the slot put it, so this narrows nothing it was not meant to.
      const roomy = sideSlot(3, slide, { left: 60, top: 0, width: 480, height: 200 });
      expect(fitInSlot(roomy, { width: 160, height: 100 }, slide).top, "moved a box that already fitted").toBe(
        roomy.top,
      );
    }
  });

  it("leaves no two charts stacked on one slide", async () => {
    // What the owner saw, twice, opening the deck the battery left him: first
    // one full-size chart drawn over another, then three in a heap. Both
    // scenarios that share a slide pick the FIRST probe chart, so they pick the
    // same one.
    //
    // Checked against the DECK rather than the arithmetic, because the geometry
    // helper is new and a test of a new function cannot fail against the file
    // that lacked it. This can: it reads back what was actually drawn.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    // On a 4:3 deck, because that is where the fixed right-hand column broke on
    // a real host: 720pt across leaves no strip wide enough beside a 480pt
    // chart, and the round put charts at x=498 over one running to x=600. 16:9
    // is covered by the arithmetic case above; this is the shape that caught a
    // rule its own test had already admitted could not hold.
    hostSlideSize.cx = 9144000; // 720pt
    hostSlideSize.cy = 6858000; // 540pt
    // The FULL battery, not a picked scenario: `insertOntoUsedSlide` is the one
    // that draws two charts onto an occupied slide, and the picker's "plus the
    // two inserts it needs" does not reach it.
    await runSelfTest("probe");
    const scan = await listChartsInDeck();
    const bySlide = new Map<string, { title: string; l: number; t: number; w: number; h: number }[]>();
    for (const c of scan.charts) {
      const cfg = JSON.parse(c.configJson) as ChartConfig;
      const box = { title: String(cfg.title), l: c.target.left, t: c.target.top, w: cfg.width!, h: cfg.height! };
      bySlide.set(c.target.slideId, [...(bySlide.get(c.target.slideId) ?? []), box]);
    }
    const shared = [...bySlide.values()].filter((v) => v.length > 1);
    expect(shared.length, "no slide ended up with two charts, so this proves nothing").toBeGreaterThan(0);
    for (const charts of shared)
      for (const [i, a] of charts.entries())
        for (const b of charts.slice(i + 1)) {
          const apart = a.l + a.w <= b.l || b.l + b.w <= a.l || a.t + a.h <= b.t || b.t + b.h <= a.t;
          expect(apart, `"${a.title}" and "${b.title}" are drawn on top of each other`).toBe(true);
        }
    hostSlideSize.cx = 12192000;
    hostSlideSize.cy = 6858000;
  }, 60_000);

  it("does not report a drawn chart as never drawn", () => {
    // The 2026-08-08 round, on the build that stopped it killing the tab. Every
    // batch committed — `upTo=24 total=24` — and then the host refused to name
    // the group, so `insertSceneIntoSlide` had no id to hand back and returned
    // null. The scenario read that as "nothing was drawn, so there is nothing
    // to look at" over twenty-four shapes that were on the slide.
    //
    // Asserted against the rule rather than through the fake, and that is a
    // deliberate second choice: three attempts to arm the fake into this state
    // (`refuseShapeById`, `refuseIdLeftTopLoads`, and both plus
    // `refuseShapeIdLoads`) each overshot into a different failure — the last
    // one threw during the draw itself, which is not the state at all. Same
    // reasoning as `targetWithNoTagResult`: the rule is what was wrong, so the
    // rule is what gets checked.
    // With the control, because the subject here is the naming caveat, not the
    // control — see "keeps the four visibility readings apart" for that.
    const unnamed = visibilityVerdict("PNG:before", "PNG:after-with-chart", false, true);
    expect(unnamed.detail, "called a drawn chart undrawn").not.toMatch(/nothing was drawn/);
    // The image moved, so the chart IS visible — that is the verdict, and the
    // naming failure rides along as a caveat rather than replacing it.
    expect(unnamed.ok, unnamed.detail).toBe(true);
    expect(unnamed.detail).toMatch(/would not name the chart/);
  });

  it("waits before re-rasterising the slide it has just rasterised", () => {
    // THE FAULT BEHIND EVERY BLIND VISIBILITY GATE, and the wiring is the fix.
    // Across 104 archived rounds the correlation is exact — 40 blind rounds
    // carry a rasterise stall, 64 sighted rounds carry none — and all 40 stalls
    // are the CONTROL render, issued immediately after the BEFORE render
    // answered, hanging for the full 20s budget. A successful render takes
    // 941ms at the median and 1466ms at its worst, so a call outstanding at 20s
    // is hung rather than slow.
    //
    // The gap is zeroed in `test/setup.ts` — a fake rasteriser cannot hang, and
    // three real seconds per scenario put four of them over budget — which
    // means NO BEHAVIOURAL TEST CAN SEE THIS CALL. Deleting it would leave the
    // suite green, exactly as deleting the pane refresh from `collectRound`
    // did. So this reads the source, and is labelled as the weaker check it is:
    // it catches the call being removed or reordered, not the gap being wrong.
    //
    // The real evidence is the blind rate across future rounds, against a 38%
    // baseline.
    // A cwd-relative path, not `import.meta.url`: this file runs under jsdom,
    // where `import.meta.url` is not a file:// URL and `readFileSync` refuses
    // it outright. Vitest runs from the repo root.
    const src = readFileSync("src/taskpane/selftest.ts", "utf8");
    const gapAt = src.indexOf("await rasterGap()");
    const controlAt = src.indexOf("rasterising the same slide a second time");
    expect(gapAt, "the gap before the control render is gone — the gate goes blind again").toBeGreaterThan(0);
    expect(gapAt, "the gap must come BEFORE the control render, or it waits after the damage").toBeLessThan(controlAt);
  });

  it("names the rasterise stall when the control is missing, instead of shrugging", () => {
    // ACROSS 14 ROUNDS THE CORRELATION IS EXACT: every blind round carries one
    // `rasterising a slide` stall at its full 20000ms budget, every sighted
    // round carries none. 8 blind / 8 stalls, 6 sighted / 0 stalls. So "no
    // control render" is not an unknown — the SECOND rasterise of the same
    // slide hangs, and the verdict can say so rather than leaving a reader to
    // join it against the trace by hand.
    _setLastStallForTest({ what: "rasterising a slide", afterAnswering: null, idleMs: 0, afterAnsweringMs: 0 });
    const blind = visibilityVerdict("PNG:before", "PNG:after", true);
    expect(blind.skipped).toBe(true);
    expect(blind.detail, "shrugged at a cause the archive has pinned 14 times").toMatch(/second rasterise/);
    // And it must NOT invent a cause it has no evidence for.
    _setLastStallForTest(null);
    const quiet = visibilityVerdict("PNG:before", "PNG:after", true);
    expect(quiet.detail).toMatch(/cannot be ruled out/);
    expect(quiet.detail).not.toMatch(/second rasterise/);

    // THE LABEL THIS FIXTURE FROZE IS NO LONGER THE ONE PRODUCTION WRITES, and
    // that is the whole failure. The case above passed continuously while the
    // real thing was broken, because it asserts against "rasterising a slide" —
    // the single string every rasterise used to trace. Once each call site was
    // named, the control render began stalling as the label below, `/rasteris/i`
    // stopped matching, and round 113 reported the cause as unknown with the
    // cause in its own trace.
    //
    // VERBATIM FROM `rounds/113-6a041de.json`, not paraphrased, so the fixture
    // cannot drift away from production a second time.
    _setLastStallForTest({
      what: "the visibility CONTROL render (same slide, back to back)",
      op: RASTERISE_OP,
      afterAnswering: null,
      idleMs: 0,
      afterAnsweringMs: 0,
    });
    const named = visibilityVerdict("PNG:before", "PNG:after", true);
    expect(named.detail, "the label carries no 'rasteris' — only `op` can classify it").toMatch(/second rasterise/);
    expect(named.detail).toContain("the visibility CONTROL render");

    // And `op` must be what decides it, not the new wording sneaking past the
    // old regex: a stall with neither marker is still an honest unknown.
    _setLastStallForTest({
      what: "waiting for something else entirely",
      afterAnswering: null,
      idleMs: 0,
      afterAnsweringMs: 0,
    });
    const other = visibilityVerdict("PNG:before", "PNG:after", true);
    expect(other.detail, "claimed a rasterise stall that never happened").toMatch(/cannot be ruled out/);
    _setLastStallForTest(null);
  });

  it("keeps the four visibility readings apart", () => {
    // The other three corners, so the rule cannot drift into always-passing.
    //
    // WITH THE CONTROL. This call used to omit `stable`, which made it the
    // no-control case — and it asserted `ok: true` and called it "a clean pass".
    // That is the vacuous pass this describe block is named after, written into
    // the test that was supposed to prevent it. A pass needs the control render.
    const named = visibilityVerdict("PNG:before", "PNG:after", true, true);
    expect(named.ok).toBe(true);
    expect(named.skipped, "a real pass is not a skip").toBeUndefined();
    expect(named.detail, "a clean pass should carry no caveat").not.toMatch(/would not name/);

    // NO CONTROL: the difference is real and means nothing, so it is blind
    // rather than green. 23 of 69 archived rounds reported this scenario green
    // in exactly this state.
    const noControl = visibilityVerdict("PNG:before", "PNG:after", true);
    expect(noControl.ok, "a difference without a control is not evidence").toBe(false);
    expect(noControl.skipped, "blind is a skip, not a failure — the chart may be fine").toBe(true);
    expect(noControl.detail).toMatch(/no control render/);

    // CONTROL SAYS THE RASTERISER IS UNSTABLE: the verdict's own text already
    // said this proves NOTHING, while returning ok.
    const unstable = visibilityVerdict("PNG:before", "PNG:after", true, false);
    expect(unstable.ok).toBe(false);
    expect(unstable.skipped).toBe(true);
    expect(unstable.detail).toMatch(/proves NOTHING/);

    // On the slide and invisible — the defect the scenario exists for.
    const invisible = visibilityVerdict("PNG:same", "PNG:same", true);
    expect(invisible.ok).toBe(false);
    expect(invisible.detail).toMatch(/nothing is visible/);

    // Nothing changed AND nothing was named: two readings, and the image cannot
    // separate them, so the verdict must not pretend it can.
    const blind = visibilityVerdict("PNG:same", "PNG:same", false);
    expect(blind.ok).toBe(false);
    expect(blind.detail).toMatch(/cannot tell/);
    expect(blind.detail, "picked one of two readings the image cannot separate").not.toMatch(/nothing is visible$/);
  });

  it("says WHERE the two renders differ, not only by how much", () => {
    // `+108 bytes` on three consecutive rounds, from three different starting
    // sizes. A chart appearing in a rasterised slide does not cost the same
    // hundred and eight bytes three times by coincidence — but a length cannot
    // tell a header from a picture, and the round file carried nothing else.
    //
    // A header, a timestamp or a counter differs EARLY and in few places; a
    // chart drawn into the image differs across the body of the data.
    const head = renderDifference("HEADERxxxxxxxxxxxxxxxxxxxxxx", "HEADYRxxxxxxxxxxxxxxxxxxxxxx");
    expect(head.at, "a difference in the first bytes was not reported as early").toBeLessThan(6);
    expect(head.differing).toBe(1);

    const body = renderDifference("aaaaaaaaaaaaaaaaaaaa", "aaaaabcbcbcbcbcbcbcb");
    expect(body.differing, "a difference through the body read as a handful of bytes").toBeGreaterThan(10);

    // A longer render counts its extra bytes as differing, or an image that
    // only grew would read as identical.
    const grew = renderDifference("aaaa", "aaaabbbb");
    expect(grew.differing).toBe(4);
    expect(grew.of).toBe(8);

    // Identical inputs: nothing differs, and `at` must not claim a position.
    const same = renderDifference("abcd", "abcd");
    expect(same.differing).toBe(0);
    expect(same.at).toBe(4);
  });

  it("refuses to claim anything when the rasteriser is not stable", () => {
    // The reading that would make every "the chart is visible" verdict on
    // record worth nothing, and which a before/after pair alone cannot rule
    // out: a host whose rasteriser answers differently for an UNCHANGED slide
    // produces exactly the same evidence as one that drew the chart.
    const unstable = visibilityVerdict("x".repeat(1000), "y".repeat(1000), true, false);
    expect(unstable.detail, `an unstable rasteriser was reported as proof: ${unstable.detail}`).toMatch(
      /proves NOTHING about the chart/,
    );

    // Control taken and clean: the difference IS the chart, and the verdict
    // says so by saying nothing extra.
    const stable = visibilityVerdict("x".repeat(1000), "y".repeat(1000), true, true);
    expect(stable.detail).not.toMatch(/proves NOTHING|no control render/);

    // Control not taken at all — the state every round before this one was in.
    // Absence must not read as a clean control.
    const unknown = visibilityVerdict("x".repeat(1000), "y".repeat(1000), true);
    expect(unknown.detail, "a missing control read as a passing one").toMatch(/no control render/);
  });

  it("carries that reading into the verdict a reader actually sees", () => {
    // Control passed: this test is about the SIZE wording in the detail.
    const v = visibilityVerdict("x".repeat(14856), "x".repeat(14856) + "y".repeat(108), true, true);
    expect(v.ok).toBe(true);
    expect(v.detail, `the verdict says nothing about where they differ: ${v.detail}`).toMatch(
      /first differ at \d+% in/,
    );
    expect(v.detail).toMatch(/108 byte\(s\) differing/);
  });

  it("says how big the visible change was, and calls a thin one thin", () => {
    // Three passes on record read identically in the round file — 10064 → 15652
    // (+55%), 15704 → 16580 (+5.6%) and 14868 → 14976, which is a hundred and
    // eight bytes. The gate asserts the two renders DIFFER, which a re-encode
    // satisfies on its own, and nothing in the verdict let a reader tell 0.7%
    // from 55% without opening the file.
    const fat = visibilityVerdict("x".repeat(10064), "y".repeat(15652), true, true);
    expect(fat.ok).toBe(true);
    expect(fat.detail, "a 55% change did not report its size").toMatch(/\+5588, 55\.5%/);
    expect(fat.detail, "a 55% change was called thin").not.toMatch(/THIN/);

    // The 2026-08-11 round, to the byte.
    const thin = visibilityVerdict("x".repeat(14868), "y".repeat(14976), true, true);
    expect(thin.ok, "a change is still a change — thinness is reported, not failed").toBe(true);
    expect(thin.detail, "a 0.7% change read exactly like a 55% one").toMatch(/\+108, 0\.7%/);
    expect(thin.detail).toMatch(/THIN margin/);
  });

  it("does not call an unreadable slide a failed picture", async () => {
    // The other half. This asked `slideHoldsOnlyChart` — the slide-SWAP gate,
    // which answers no for a slide it could not read, because it authorises
    // deleting one. PowerPoint on the web refuses shape collections routinely,
    // so a perfectly good picture logged "picture may not be a single shape"
    // on the strength of a host that never looked. The round-trip below is
    // still worth checking; the verdict just may not say "picture".
    const deck = [makeSlide("s1")];
    installHost(deck);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    // Every `items/name` read comes back empty while the slide's own count
    // does not, which is the corroboration failing — the host answering short
    // without throwing, rather than a slide that is genuinely bare.
    faults.hollowNameReads = 99;
    setTracing(true);
    try {
      const explode = named(await runSelfTest("probe"))["explode a degraded picture"];
      // The old sentence blamed the picture for the host's silence, and it is
      // the log that carried it — the verdict read "config intact" either way,
      // so only the trace can tell the two versions apart.
      const said = traceLog()
        .entries.filter((e) => e.scope === "selftest")
        .map((e) => e.message);
      expect(said, `the log blamed the picture: ${said.join(" | ")}`).not.toContain(
        "picture may not be a single shape",
      );
      expect(said).toContain("the host would not confirm the picture is one shape");
      // …and the verdict says which of the two round-trips it actually saw.
      expect(explode.detail).not.toMatch(/a picture is one/);
      expect(explode.detail).toMatch(/would not confirm/);
      expect(explode.ok, explode.detail).toBe(true);
    } finally {
      setTracing(false);
      faults.hollowNameReads = 0;
    }
  });

  it("skips the rescale rather than comparing against -Infinity", async () => {
    // `Math.max()` of nothing is -Infinity, and JSON.stringify turns that into
    // `null` — so an all-blank data set wrote {"scale":{"max":null}} and then
    // compared it against -Infinity, which never matches. The scenario
    // reported a failure that was its own arithmetic.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => undefined);
    const scale = named(await runSelfTest("probe"))["same scale across the deck"];
    expect(scale.detail).toBeTruthy();
    expect(scale.detail).not.toContain("Infinity");
    expect(scale.detail).not.toContain("null");
  });

  it("applies a ceiling the chart has to REDRAW for, not the one it already had", async () => {
    // The scenario is named for a deck-wide rescale and spent its whole life
    // measuring a deck-wide re-tag. `Math.max(...values)` is the ceiling the
    // sample is already drawn at, so `scale: { max }` rendered to the
    // byte-identical scene — a deck saved out of a real PowerPoint had the
    // rescaled chart and an untouched copy agreeing on bar geometry to the EMU.
    // A host that stored the config and redrew the old picture would have
    // passed this scenario every time.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    const scale = named(await runSelfTest("probe"))["same scale across the deck"];
    expect(scale.skipped, scale.detail).toBeFalsy();

    const scan = await listChartsInDeck();
    const scaled = scan.charts
      .map((c) => JSON.parse(c.configJson) as ChartConfig)
      .filter((c) => typeof c.scale?.max === "number");
    expect(scaled.length, "the scenario wrote no shared scale at all").toBeGreaterThan(0);

    const barInk = (cfg: ChartConfig) => {
      let total = 0;
      for (const n of buildChart(cfg).nodes) if (n.kind === "rect" && n.name?.startsWith("seg-")) total += n.h;
      return total;
    };
    for (const cfg of scaled) {
      const loose = barInk({ ...cfg, scale: undefined });
      expect(
        barInk(cfg),
        `scale.max=${cfg.scale?.max} draws the same bars as no scale at all — the rescale proves nothing`,
      ).toBeLessThan(loose);
    }
  });

  it("says WHY the charts it could not scale failed, not just how many", async () => {
    // The verdict from a real round was "4 of 8 charts carry the shared scale"
    // and nothing else. Reading the cause out of the deck afterwards took a
    // session; `updateChartInSlide` had it at the time and this scenario threw
    // the return value away.
    //
    // The host here is the one that round found: a tag write refused from the
    // drawing context AND from the settle pass that exists to repair it.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    faults.refuseTagWrites = 9999;
    try {
      const scale = named(await runSelfTest("probe"))["same scale across the deck"];
      expect(scale.ok, "a deck where nothing could be tagged reported success").toBe(false);
      expect(scale.detail, `the verdict counts but does not explain: ${scale.detail}`).toMatch(
        /the update reported \d+×no-config/,
      );
    } finally {
      faults.refuseTagWrites = 0;
    }
  });

  it("stops once the host has refused twice, and says where it flipped", async () => {
    // 212 seconds of an 818-second round, the longest thing in the battery, and
    // on a degrading host its tail is repetition: three rounds decomposed by
    // hand put the boundary in the same place and every chart past it did the
    // same thing at ~12s each. Round 16 would have stopped after chart 5.
    //
    // The flip INDEX is the number worth carrying. What moved a score between
    // 3 and 4 across rounds was one binary event — whether the settle caught
    // the first degraded chart — not a boundary sliding along the scenario, and
    // a score alone cannot say that.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    setTracing(true);
    faults.refuseTagWrites = 9999;
    try {
      const scale = named(await runSelfTest("probe"))["same scale across the deck"];
      expect(scale.detail, `the verdict names no flip point: ${scale.detail}`).toMatch(
        /the host flipped at chart 1 of \d+/,
      );
      expect(scale.detail, "every chart was attempted on a host that refused the first two").toMatch(
        /so the last \d+ were not attempted/,
      );
      const stopped = traceLog().entries.filter(
        (e) => e.message === "stopping the rescale — the host has already flipped",
      );
      expect(stopped.length, "the scenario gave up without saying so").toBe(1);
      expect(stopped[0].data).toMatchObject({ done: 2, flippedAt: 1 });
    } finally {
      faults.refuseTagWrites = 0;
      setTracing(false);
    }
  });

  it("attempts every chart on a host that behaves", async () => {
    // The property that makes the shortcut safe: it can only skip work the host
    // has already refused twice. A healthy deck loses nothing, so nothing is
    // skipped and the scenario still proves every chart takes the shared scale.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    setTracing(true);
    try {
      const scale = named(await runSelfTest("probe"))["same scale across the deck"];
      expect(scale.detail, `a healthy host was cut short: ${scale.detail}`).not.toMatch(/not attempted/);
      expect(scale.detail, "a healthy host was reported as having flipped").not.toMatch(/flipped at chart/);
      expect(
        traceLog().entries.some((e) => e.message === "stopping the rescale — the host has already flipped"),
        "the rescale gave up on a host that never refused anything",
      ).toBe(false);
    } finally {
      setTracing(false);
    }
  });
});

describe("when the deck-wide rescale has learned everything it will", () => {
  it("keeps going until two charts in a row have lost their config", () => {
    // Two, not one. A single loss is the flip's first chart and may still be
    // rescued by the settle — round 16's chart 4 was, and it counted toward the
    // score — so stopping on one throws away the observation that decides
    // whether a round scores 3 or 4.
    expect(rescaleShouldStop([])).toBe(false);
    expect(rescaleShouldStop(["no-config"])).toBe(false);
    expect(rescaleShouldStop([undefined, "no-config"])).toBe(false);
    expect(rescaleShouldStop(["no-config", "no-config"])).toBe(true);
    // Round 16's shape: three clean, then the flip. Chart 4 alone is not enough.
    expect(rescaleShouldStop([undefined, undefined, undefined, "unknown-shape"])).toBe(false);
    expect(rescaleShouldStop([undefined, undefined, undefined, "unknown-shape", "no-config"])).toBe(true);
  });

  it("does not stop on two losses that are not consecutive", () => {
    // A host that drops one chart and recovers has not flipped, and the tail is
    // still worth drawing.
    expect(rescaleShouldStop(["no-config", undefined])).toBe(false);
    expect(rescaleShouldStop(["no-config", undefined, "no-config", undefined])).toBe(false);
  });

  it("reports the first chart that lost its config, not the last", () => {
    expect(rescaleFlipIndex([])).toBeNull();
    expect(rescaleFlipIndex([undefined, undefined])).toBeNull();
    expect(rescaleFlipIndex([undefined, undefined, undefined, "unknown-shape", "no-config"])).toBe(4);
    expect(rescaleFlipIndex(["chart-gone"])).toBe(1);
  });

  it("does not call one lost chart a FLIP", () => {
    // Caught on the sentence's first outing. The `4feb5be` round lost exactly
    // one chart of eight and scored 7 of 8 — the best this scenario has ever
    // recorded — and the verdict announced `the host flipped at chart 5 of 8`,
    // the same word three earlier rounds use for a host that degraded and never
    // came back. A reader comparing rounds would have read a regime change into
    // the best round on file.
    const oneLost = [undefined, undefined, undefined, undefined, "no-config", undefined, undefined, undefined];
    const note = rescaleLossNote(oneLost, 8);
    expect(note, `a single loss was called a flip: ${note}`).not.toMatch(/flip/i);
    expect(note, "the index of the one loss went unreported").toMatch(/the first at chart 5 of 8/);
    expect(note).toMatch(/1 of the 8 charts redrawn/);

    // And nothing at all to say when nothing was lost.
    expect(rescaleLossNote([undefined, undefined], 2)).toBe("");
  });

  it("says WHY the rescale lost charts, and whether the pause saved any", () => {
    // Thirty-four rounds reported "3 of 8 charts carry the shared scale … the
    // host flipped at chart 4 of 8" and never once said why, so reading a round
    // meant opening the trace and joining it back by hand. The mechanism is
    // settled now — a short or empty pre-grouping re-read on a slide the run has
    // just added — so the verdict line can carry it.
    const zero = { emptyReReads: 0, shortReReads: 0, reReadsRepaired: 0 };
    const note = reReadNote(zero, { emptyReReads: 1, shortReReads: 1, reReadsRepaired: 2 });
    expect(note, "the two failure modes drive different branches and must not be pooled").toMatch(/1 read short/);
    expect(note).toMatch(/1 read empty/);
    expect(note).toMatch(/repaired 2/);

    // DIFFERENCED, not absolute. The counters are cumulative for the whole
    // session, so a scenario that reported them raw would report every earlier
    // scenario's friction as its own.
    expect(
      reReadNote(
        { emptyReReads: 5, shortReReads: 5, reReadsRepaired: 5 },
        { emptyReReads: 6, shortReReads: 5, reReadsRepaired: 5 },
      ),
      "read an absolute counter as this scenario's own",
    ).toMatch(/1 read empty/);

    // "REPAIRED NONE" IS THE FINDING THAT REFUTES THE FIX, so it must survive
    // being zero. A clause that vanished at 0 would read identically to a round
    // where the retry never ran at all, and those are opposite results.
    expect(
      reReadNote(zero, { emptyReReads: 2, shortReReads: 0, reReadsRepaired: 0 }),
      "a pause that saved nothing said nothing, so the refutation is invisible",
    ).toMatch(/repaired 0/);

    // Silent when there is nothing to report — it must not add words to a clean
    // round.
    expect(reReadNote(zero, zero)).toBe("");
  });

  it("says FLIPPED only for the two-in-a-row the scenario stops on", () => {
    // The word is reserved for the evidence `rescaleShouldStop` acts on, so the
    // verdict and the decision to stop can never disagree.
    const flipped = [undefined, undefined, undefined, "unknown-shape", "no-config"];
    const note = rescaleLossNote(flipped, 8);
    expect(note).toMatch(/the host flipped at chart 4 of 8/);
    expect(note, "a scenario that stopped early did not say what it skipped").toMatch(
      /so the last 3 were not attempted/,
    );

    // Two in a row at the very end: a flip, but nothing was skipped.
    expect(rescaleLossNote([undefined, "no-config", "no-config"], 3)).toMatch(/flipped at chart 2 of 3$/);
  });
});

describe("the scenario for a slide that already has content", () => {
  const onto = (rs: ScenarioResult[]) => rs.find((r) => r.name === "insert onto a slide that already has content")!;

  it("passes on a host that behaves, and on one that does not", async () => {
    // The everyday action — "insert a chart on the slide I am looking at" —
    // and until this scenario existed, the only path with no real-host
    // coverage at all. Every other scenario, and the whole demo deck, works on
    // slides the run added BLANK.
    for (const hostile of [false, true]) {
      installHost([makeSlide("s1")]);
      if (hostile) applyWebProfile();
      setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
      const r = onto(await runSelfTest("probe"));
      expect(r.skipped, r.detail).toBeFalsy();
      expect(r.ok, `${hostile ? "hostile" : "clean"}: ${r.detail}`).toBe(true);
      expect(r.detail).toContain("re-editable");
    }
  }, 60_000);

  it("reports a failure when the charts land un-re-editable", async () => {
    // The negative control, and the whole point of the scenario. A host with
    // no tag support (below PowerPointApi 1.3) cannot make a chart
    // re-editable, so the scenario must SAY so rather than count the shapes
    // and call it a pass.
    //
    // Proven against the real defect too: reverting #243 — the fresh-proxy
    // re-fetch on `insertSceneIntoSlide` — makes this scenario report
    // "0 of 2 new charts are re-editable" under `strictGroup` and under the
    // full web profile, while still passing on a clean host. That is exactly
    // the shape of the bug: invisible except on a host that refuses stale
    // proxies, which is the web.
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.3");
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    const r = onto(await runSelfTest("probe"));
    expect(r.ok, r.detail).toBe(false);
    expect(r.detail).toMatch(/re-editable/);
  }, 60_000);
});

describe("the experiment that asks what makes a long run slow down", () => {
  /**
   * Every real-host artefact this project owns degrades, and not one of them
   * can say WHY: a run that gets slower grows its deck, ages its tab and holds
   * a request context open all at the same time. The repo's own rule for that
   * situation is to ask a question that separates the answers rather than to
   * reason about which is likelier — this is the question, and these are the
   * cases that check it actually separates them.
   *
   * `faults.syncCostMs` is what makes that checkable at all. Given a cost that
   * follows the syncs within a context it stands for proxy accumulation; given
   * one that follows the syncs since the tab opened it stands for a host
   * slowing down. The verdict has to name a different suspect for each, and a
   * verdict that cannot is worse than none: it would send a session's work at
   * the wrong thing.
   */
  const series = async (id: string, oneContext: boolean, rounds = 9) =>
    (await timeShapeRounds(id, { rounds, perRound: 2, oneContext, label: "t", budgetMs: 30_000 })).rounds.map(
      (r) => r.ms,
    );

  it("uses one context for the long arm and one per round for the other", async () => {
    // The structural half, and deterministic: no clock, no threshold. If the
    // two arms were ever wired the same way — both long, both fresh, or
    // swapped — every timing case below would still pass, comparing a series
    // against itself and reporting whatever the noise said.
    installHost([makeSlide("s1")]);
    const before = trips.contexts;
    await series("s1", true);
    const long = trips.contexts - before;
    const mid = trips.contexts;
    await series("s1", false);
    expect(long).toBe(1);
    expect(trips.contexts - mid).toBe(9);
  });

  it("blames the context when only a held context gets slower", async () => {
    installHost([makeSlide("s1")]);
    // Flat 10ms plus 8ms per sync already spent IN THIS CONTEXT. The fresh arm
    // never sees the second term, so its cost is constant by construction.
    faults.syncCostMs = ({ syncsInContext }) => 10 + syncsInContext * 8;
    const one = await series("s1", true);
    const fresh = await series("s1", false);
    const v = readDegradation(one, fresh);
    // Timed arms: the experiment ran and the reader answered. Nothing about
    // WHICH arm grew — see the sibling below for why a wall-clock comparison
    // between these two arms cannot be trusted under load.
    expect(one.length, "the arms measured different numbers of rounds").toBe(fresh.length);
    expect(["none", "host", "context", "both"]).toContain(v.suspect);

    // The claim this test is named for, where the two arms cost the same by
    // construction: only the held arm climbs, and only the held arm is blamed.
    const heldGrows = [10, 18, 26, 34, 42, 50, 58, 66, 74];
    const freshFlat = [10, 10, 10, 10, 10, 10, 10, 10, 10];
    const exact = readDegradation(heldGrows, freshFlat);
    expect(exact.freshContext, "a flat fresh-context arm was reported as growing").toBeLessThan(0.5);
    expect(exact.oneContext, "a climbing held-context arm was not reported as growing").toBeGreaterThan(0.5);
    expect(exact.suspect).toBe("context");
    expect(exact.headline).toContain("THE CONTEXT");
  });

  it("blames the host when both arms get slower together", async () => {
    installHost([makeSlide("s1")]);
    // The same shape of cost, hung on the counter that does NOT reset at a
    // context boundary. A context boundary buys nothing here, which is exactly
    // what a growing deck or an ageing tab would look like.
    //
    // Rebased on the syncs already spent, so the case does not depend on where
    // in the file it runs — `trips.syncs` is a whole-suite counter and an
    // absolute cost read off it would make this test's answer depend on the
    // tests before it.
    const base = trips.syncs;
    faults.syncCostMs = ({ syncsTotal }) => 10 + (syncsTotal - base) * 6;
    const one = await series("s1", true);
    const fresh = await series("s1", false);
    const v = readDegradation(one, fresh);
    // What the timed arms can carry, and no more.
    //
    // "The verdict is robust because it is a choice between words" was the
    // earlier reasoning here, and it was wrong — this went red under full-suite
    // load after that fix, exactly like its sibling below. The mechanism is not
    // jitter: the FRESH arm pays for building a context every round, ~90ms on a
    // loaded machine against the 6ms this fault charges, and `growth` is a ratio,
    // so that fixed cost flattens the arm carrying it and the reader concludes —
    // correctly, from those numbers — that the held context is to blame. No
    // threshold and no word-set repairs a one-sided structural cost.
    //
    // Every cross-arm claim therefore lives on synthetic arrays, where the two
    // arms are equal by construction. What is left here is the end-to-end half:
    // the experiment ran, both arms measured, the reader answered.
    expect(one.length, "the arms measured different numbers of rounds").toBe(fresh.length);
    expect(
      ["none", "host", "context", "both"],
      `no verdict — one: ${one.join(",")} | fresh: ${fresh.join(",")}`,
    ).toContain(v.suspect);

    // The NUMBER moves onto synthetic arrays, for the reason its sibling below
    // already sets out — and this test is the evidence that the reason applies
    // here too.
    //
    // It was deliberately left alone when that one was fixed, on the grounds
    // that `> 0.5` is coarse enough to ride out the clock. It is not. Windows
    // resolves timers at ~15.6ms while the fault charges 6ms a sync, so the
    // series climbs in steps three rounds wide; under full-suite load this went
    // red having passed three times out of three in isolation.
    //
    // The load-bearing claim is that the FRESH arm grew — the one thing a
    // context-accumulation story cannot produce, and the one thing that tells a
    // reader shortening contexts will not help. Both arms climb 10 a round,
    // hit equally, which is what "the host is slowing whatever the context
    // does" looks like.
    const flatOne = [100, 110, 120, 130, 140, 150, 160, 170, 180];
    const flatFresh = [190, 200, 210, 220, 230, 240, 250, 260, 270];
    const exact = readDegradation(flatOne, flatFresh);
    expect(exact.freshContext, "a fresh-context arm that grew was not reported as growing").toBeGreaterThan(0.5);
    expect(["host", "both"]).toContain(exact.suspect);
  });

  it("does not read a host that slows LINEARLY as a context problem", async () => {
    // The case the first version of this got wrong, kept as its own test
    // because it is the commonest shape a real slowdown takes and the mistake
    // was invisible without it. Comparing each arm's tail to its OWN head, the
    // long arm read ×2.64 and the fresh arm ×1.47 for a slowdown neither arm
    // caused — the fresh arm only looks flatter because it starts from a higher
    // baseline. A confident "the context did it", about a host.
    //
    // Same fault as the case above; what this one pins is the NUMBER, not just
    // the label. Both arms are hit equally hard in milliseconds, so the two
    // growth figures must come out within a hair of each other.
    installHost([makeSlide("s1")]);
    const base = trips.syncs;
    faults.syncCostMs = ({ syncsTotal }) => 10 + (syncsTotal - base) * 6;
    const one = await series("s1", true);
    const fresh = await series("s1", false);
    const v = readDegradation(one, fresh);
    // All the timed arms can honestly carry: the pipeline ran, over real
    // measurements, and produced a legal verdict.
    //
    // Even `not.toBe("context")` had to come off them, and by a DIFFERENT
    // mechanism from the quantisation described further down — worth keeping
    // apart, because the earlier fix assumed one cause and left this behind.
    //
    // Caught under full-suite load, one run in four:
    //
    //     one:   23,31,31,48,59,67,67,67,74
    //     fresh: 98,84,110,119,100,107,117,117,124
    //
    // The fresh-context arm OPENS at ~98ms where the held arm opens at ~23ms.
    // That gap is not the injected fault: it is what building a new context per
    // round actually costs on a loaded machine, and it is fifteen times the 6ms
    // the fault charges. `growth` is a ratio, so a large fixed cost flattens the
    // arm carrying it — fresh reads ×1.2 against one's ×2.5 — and from those
    // numbers the reader correctly concludes the held context is the problem.
    // The numbers are the lie, not the reader.
    //
    // The two arms are therefore not comparable in wall-clock at this scale, and
    // no threshold repairs that: the noise is structural and one-sided. Anything
    // about WHICH arm grew is asserted on synthetic arrays below, where the
    // costs are equal by construction — including the claim this test is named
    // for, which is precisely what `exact.suspect === "host"` says.
    expect(
      ["none", "host", "context", "both"],
      `the reader returned no verdict — one: ${one.join(",")} | fresh: ${fresh.join(",")}`,
    ).toContain(v.suspect);
    expect(one.length, "the held-context arm measured nothing").toBe(fresh.length);

    // The NUMBER is pinned on synthetic arrays rather than on the timed arms
    // above, and that is a measurement rather than a preference.
    //
    // It used to read the timed arms with a tolerance of 0.35. That fails on
    // Windows five runs out of five — gaps of 0.355, 0.375, 0.387, 0.387,
    // 0.452 — and passes in CI, which is how it survived. The clock is the
    // cause, not the reader: the fault charges 6ms per sync and each round
    // spends exactly ONE sync (measured on both arms, so they really are hit
    // equally), while Windows' timer granularity is ~15.6ms. A series that
    // should climb 6ms a round instead climbs in ~16ms steps every third round
    // — `one` came back 26,32,31 · 48,47,46 · 63,64,79 — and `growth`'s
    // median-of-thirds lands the two arms on different phases of that
    // staircase. Re-run at 20ms per sync, comfortably over one tick, it passed
    // five out of five.
    //
    // Widening the tolerance would have buried that, and raising the modelled
    // cost until it clears the tick pushes this test into its own 30-second
    // budget (80ms per sync blows it). A round the clock cannot resolve is
    // below the instrument — the same thing `readDegradation`'s own
    // `Math.max(one.head, 1)` says one level down — so the arithmetic is
    // asserted where no clock can reach it, and the timed arms above keep the
    // end-to-end half.
    //
    // Both arms climb 10 a round; `fresh` merely starts higher, which is the
    // entire trap. A reader dividing each arm by its OWN head calls this +55%
    // against +30% and blames THE CONTEXT for a host-wide slowdown.
    const flatOne = [100, 110, 120, 130, 140, 150, 160, 170, 180];
    const flatFresh = [190, 200, 210, 220, 230, 240, 250, 260, 270];
    const exact = readDegradation(flatOne, flatFresh);
    expect(exact.suspect, "two arms climbing at one rate are a HOST slowdown").toBe("host");
    expect(exact.oneContext - exact.freshContext, "the arms were hit equally and must read equally").toBe(0);
  });

  it("blames nothing when the host is steady", async () => {
    installHost([makeSlide("s1")]);
    faults.syncCostMs = () => 6;
    const v = readDegradation(await series("s1", true), await series("s1", false));
    // Timed: the reader answered at all. A steady FAULT is not a steady clock —
    // under load the fresh arm's per-round context cost moves on its own, and
    // that alone can lift either arm off "none".
    expect(["none", "host", "context", "both"]).toContain(v.suspect);
    // Steady, where steady is guaranteed.
    expect(readDegradation([10, 10, 10, 10, 10, 10, 10, 10, 10], [10, 10, 10, 10, 10, 10, 10, 10, 10]).suspect).toBe(
      "none",
    );
  });

  it("says so rather than guessing when there are too few rounds to read", async () => {
    // Three rounds is not a curve. Reporting "no degradation" from it would be
    // the same mistake as reading a blind deck scan as an empty deck, which
    // this file already has a whole third result state for.
    installHost([makeSlide("s1")]);
    const v = readDegradation(await series("s1", true, 3), await series("s1", false, 3));
    expect(v.suspect).toBe("unknown");
    expect(v.headline).toContain("not enough rounds");
  });

  it("keeps the rounds it managed when the host stops answering", async () => {
    // A curve with four points is a finding; an exception is not. The long arm
    // is the one at risk, because its whole series sits under one deadline.
    installHost([makeSlide("s1")]);
    faults.wedgeAfterSyncs = 4;
    const got = await timeShapeRounds("s1", {
      rounds: 9,
      perRound: 2,
      oneContext: true,
      label: "t",
      budgetMs: 150,
    });
    expect(got.rounds).toHaveLength(4);
    expect(got.cutShort, "a series that lost five rounds must say so").toBeTruthy();
  });

  it("is offered by the picker, left out of a full run, and draws its arms on two slides", async () => {
    // The fake pushes `slides.add()` onto the array it was handed, so the deck
    // the experiment built its two scratch slides in is readable here — which
    // is the only way to check the arms went to different ones.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const deckShapeNames = () => deck.map((s) => s.created.map((c) => c.name ?? ""));
    _setDegradeSizeForTest(4, 2);
    try {
      expect(SCENARIO_NAMES).toContain("what makes a long run slow down");
      expect(ROUTINE_SCENARIO_NAMES).not.toContain("what makes a long run slow down");
      const r = (await runSelfTest("probe", "what makes a long run slow down")).find(
        (x) => x.name === "what makes a long run slow down",
      )!;
      expect(r.detail).toContain("one context:");
      expect(r.detail).toContain("fresh contexts:");
      // Two arms, two slides. Sharing one would make the second arm draw onto a
      // slide already holding the first arm's shapes, and a fat slide is a
      // fourth variable in an experiment built to have exactly one.
      const named = (frag: string) => deckShapeNames().filter((names) => names.some((n) => n.includes(frag))).length;
      expect(named("one-context")).toBe(1);
      expect(named("fresh-context")).toBe(1);
    } finally {
      _setDegradeSizeForTest(8, 12);
    }
  }, 60_000);

  it("draws its grid where the caller put it, not at a corner it picked itself", async () => {
    // The grid's origin used to be the constant (20, 430), under a comment
    // asserting it cleared `sideSlot`'s right-hand column. It did — and the
    // column became a band along the bottom of the slide, which is exactly
    // where those two numbers point. So the experiment's measurement artefact
    // would have been drawn across the first chart slot on the same slides the
    // other scenarios use.
    //
    // The fix is not a better constant: it is that the caller decides, because
    // only the caller knows where the chart on that slide actually is. This
    // checks the parameter is honoured at all, which a constant cannot do.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const origin = { left: 600, top: 300 };
    await timeShapeRounds("s1", { rounds: 2, perRound: 3, oneContext: true, label: "t", budgetMs: 30_000, origin });
    const need = gridFootprint(6);
    const boxes = deck[0].created.map((s) => s.box!);
    expect(boxes.length, "the arm drew nothing to check").toBe(6);
    for (const b of boxes) {
      expect(b.left, "a grid cell landed left of the origin it was given").toBeGreaterThanOrEqual(origin.left);
      expect(b.top, "a grid cell landed above the origin it was given").toBeGreaterThanOrEqual(origin.top);
      expect(b.left + b.width, "a grid cell ran off the right of its box").toBeLessThanOrEqual(
        origin.left + need.width,
      );
      expect(b.top + b.height, "a grid cell ran off the bottom of its box").toBeLessThanOrEqual(
        origin.top + need.height,
      );
    }
  });
});

describe("the scenario for a shape the user had selected", () => {
  /**
   * office-js#2775, open and web-only: adding a text box deletes whatever shape
   * was selected. Every chart drawn here contains text boxes, and the insert
   * path deliberately leaves the user's selection alone because that is how it
   * learns where to put the chart — so on a host where this is live, selecting
   * a picture and inserting a chart against it destroys the picture.
   */
  const pick = (rs: ScenarioResult[]) => byName(rs)["a selected shape survives an insert"];

  it("passes on a host that keeps the selected shape", async () => {
    installHost([makeSlide("s1")]);
    const r = pick(await runSelfTest("probe", "a selected shape survives an insert"));
    expect(r.skipped, r.detail).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail).toMatch(/survived an insert/);
  });

  it("catches a host that deletes it, and names the issue", async () => {
    // The negative control, and the whole reason the scenario exists. Without
    // the fault armed the assertion above passes against a fake that could not
    // lose a shape if it tried.
    installHost([makeSlide("s1")]);
    faults.textBoxDeletesSelection = true;
    try {
      const r = pick(await runSelfTest("probe", "a selected shape survives an insert"));
      expect(r.skipped, r.detail).toBeFalsy();
      expect(r.ok, "reported a host that destroyed the user's shape as a pass").toBe(false);
      expect(r.detail).toMatch(/VANISHED/);
      expect(r.detail, "did not name what a reader would search for").toContain("office-js#2775");
    } finally {
      faults.textBoxDeletesSelection = false;
    }
  });

  it("does not run before the ladder has asked whether selection works at all", async () => {
    // It calls `setSelectedShapes`, so on a wedged host it would spend a budget
    // to report silence the ladder has already attributed properly, one line up.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.selectionWedgesHost = true;
    _setReadbackTimeoutForTest(200);
    let all: ScenarioResult[];
    try {
      all = await runSelfTest("probe");
    } finally {
      faults.selectionWedgesHost = false;
      _setReadbackTimeoutForTest(90_000);
    }
    const r = pick(all);
    expect(r.skipped, "attributed a wedged host to this scenario").toBe(true);
    expect(r.detail).toMatch(/ladder found/);
  });
});

/**
 * A diagnosis is worth having only if it can be wrong.
 *
 * `edit the chart the user selected` catches a timeout and explains it: "the
 * host stopped answering selection calls after a programmatic select — known
 * web-host limitation, same family as office-js#3083 / #3698". It caught ANY
 * timeout, and `isTimeout` is just as true of a draw that stalled.
 *
 * A round on 2026-08-08 skipped with that sentence while its own trace read
 * `gave up waiting what=drawing shapes 1-10 of 24` and `at=redrawing the
 * chart's shapes`. The selection subsystem was fine; the host stopped answering
 * during a redraw; and the report sent the reader to two selection bugs that had
 * nothing to do with it — while quietly not counting a host stall as a failure.
 */
describe("which timeout counts as the known selection limitation", () => {
  it("does not call a draw stall a selection limitation", () => {
    const drawStall = new Error("PowerPoint did not respond while drawing shapes 1-10 of 24 (45s)");
    expect(wedgedSelection(drawStall), "a stalled draw was labelled a selection bug").toBe(false);
    const redraw = new Error("PowerPoint did not respond while redrawing the chart's shapes (45s)");
    expect(wedgedSelection(redraw), "a stalled redraw was labelled a selection bug").toBe(false);
  });

  it("still recognises the selection wedge it was written for", () => {
    // The negative control. A discriminator that says no to everything would
    // pass the case above and turn a known, harmless host limitation back into
    // a red line — which is the failure this branch was added to prevent.
    expect(wedgedSelection(new Error("PowerPoint did not respond while reading the selected chart (10s)"))).toBe(true);
    expect(wedgedSelection(new Error("PowerPoint did not respond while selecting a shape (10s)"))).toBe(true);
  });

  it("treats an unattributable timeout as a failure, not a known limitation", () => {
    // Unknown is not the same as innocent. An error with nothing to place it
    // could have come from anywhere, and the honest report is a failure that
    // says so rather than a confident wrong answer.
    expect(wedgedSelection(new Error("PowerPoint did not respond"))).toBe(false);
    expect(wedgedSelection(undefined)).toBe(false);
  });
});

describe("the experiment that asks whether a rasterise poisons the next draw", () => {
  const NAME = "does a rasterise poison the next draw";
  const pick = async () => byName(await runSelfTest("probe", NAME))[NAME];

  it("passes when every arm draws, and says so", async () => {
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    const r = await pick();
    expect(r.skipped, r.detail).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail).toMatch(/all four draws landed/);
  });

  it("names the CALL only when the call is what separates the arms", () => {
    // Pure, and extracted for the reason `visibilityVerdict` was: three
    // attempts to arm the fake into each of these states each overshot into a
    // different failure, exercising the fake's plumbing while the RULE — which
    // is the thing that can be wrong — went unchecked.
    const ok = { drew: true, why: "" };
    const dead = { drew: false, why: "the host stopped answering" };

    // All four landed.
    expect(rasteriseArmVerdict([ok, ok], [ok, ok]).ok).toBe(true);
    expect(rasteriseArmVerdict([ok, ok], [ok, ok]).detail).toMatch(/all four draws landed/);

    // The finding the experiment exists to produce: every rasterise arm
    // refused, every cheap arm drew, interleaved so position cannot explain it.
    const guilty = rasteriseArmVerdict([dead, dead], [ok, ok]);
    expect(guilty.ok).toBe(false);
    expect(guilty.detail, "did not name the call").toMatch(/every draw after a RASTERISE failed/);
    expect(guilty.detail, "did not say why position is excluded").toMatch(/interleaved/);

    // The reading the TWO-arm version could not reach, and got wrong on its
    // first real round: the late half failed and the early half did not,
    // whichever call preceded them.
    const positional = rasteriseArmVerdict([ok, dead], [ok, dead]);
    expect(positional.ok).toBe(false);
    expect(positional.detail, "blamed the rasterise for a late-in-the-round stall").toMatch(/both LATER draws failed/);
    expect(positional.detail).not.toMatch(/every draw after a RASTERISE failed/);

    // A slide that refuses everything looks identical to a poisoned call unless
    // the cheap arms are consulted.
    const dying = rasteriseArmVerdict([dead, dead], [dead, dead]);
    expect(dying.detail, "blamed the rasterise for a slide that refuses everything").toMatch(/no draw landed at all/);
    expect(dying.detail).not.toMatch(/every draw after a RASTERISE failed/);

    // And the shape round 11 actually produced — one rasterise arm each way —
    // which is no separation and must say so rather than pick a side.
    //
    // It PASSES, and that is the correction. This returned `ok: false` while
    // printing "no pattern ... which is what eleven rounds of eliminated
    // candidates already said" — a control failing on its own documented answer.
    // The 2026-08-09 round is what surfaced it: `3 of 4 draws landed`, reported
    // as a scenario failure, from the intermittent stall this control was built
    // to sit beside. At one or two draws in fifteen against four draws a round,
    // that red arrives roughly every third round for a reason nobody needs to
    // act on, and a red on a schedule is one people learn to skip.
    const mixed = rasteriseArmVerdict([ok, dead], [ok, ok]);
    expect(mixed.ok, "a control that finds no separation has done its job").toBe(true);
    expect(mixed.detail, "claimed a pattern from a mixed result").toMatch(/no pattern/);
    expect(mixed.detail).not.toMatch(/every draw after a RASTERISE failed/);
    expect(mixed.detail).not.toMatch(/both LATER draws failed/);
  });
});

/**
 * The scenario built because the archive cannot answer what a chart KIND costs:
 * every existing scenario draws one chart in one context, so kind, size,
 * preceding call and slot are one variable wearing four names.
 *
 * Its counterbalance is the part that can be silently wrong. If the kinds ran
 * in one pass instead of there-and-back, kind would be confounded with position
 * exactly as it is confounded with scenario today — the same defect, moved into
 * the instrument built to escape it, and nothing about a green round would say
 * so. So the ORDER is asserted as a property, not eyeballed.
 */
describe("what a chart kind costs", () => {
  it("draws each kind once early and once late, or it measures position", () => {
    const n = KIND_COST_ORDER.length;
    expect(n % 2, "an odd number of specimens cannot be balanced").toBe(0);
    // A palindrome is the balance: position i and position n-1-i are the same
    // kind, so every kind's two slots are mirrored about the middle.
    for (let i = 0; i < n; i++)
      expect(KIND_COST_ORDER[i], `position ${i} is not mirrored at ${n - 1 - i}`).toBe(KIND_COST_ORDER[n - 1 - i]);
    const seen = new Map<string, number>();
    for (const k of KIND_COST_ORDER) seen.set(k, (seen.get(k) ?? 0) + 1);
    for (const [k, c] of seen) expect(c, `${k} appears ${c} times, so its positions cannot balance`).toBe(2);
    expect(seen.size, "one kind is not a comparison").toBeGreaterThan(1);
    // The cheap population's own kind has to be in it, or there is nothing to
    // compare the others AGAINST: `clustered` is what every other scenario
    // draws and what the 3,615ms arm of the archive is made of.
    expect([...seen.keys()], "no clustered specimen, so the readings anchor to nothing").toContain("clustered");
  });

  /**
   * THE PALINDROME BALANCES POSITION. IT ONLY BALANCES OCCUPANCY IF THE
   * SPECIMENS ARE THE SAME SIZE — and the first version's were not.
   *
   * Rounds 398 and 399 ran `clustered, line, area, pie`. `pie` draws 37 shapes
   * at this size, so it was multi-batch — making `last batch settled` its tail
   * rather than its cost — and it drove the prior occupancy to 60 before the
   * second half began, reaching 91 by the end, where a draw costs three to five
   * times what it does at zero. The experiment built to remove a confound had
   * reintroduced it, and BOTH ROUNDS WERE GREEN, because the verdict is about
   * landing and not about cost.
   *
   * So the design constraint is asserted here instead of living in prose. It is
   * also the property `area` fails for a different reason: its count runs 8,
   * 11, 15, 23 across the boxes a specimen might get, so it cannot hold
   * occupancy equal even against itself.
   */
  it("draws specimens of near-equal size, or the palindrome cannot balance occupancy", () => {
    const at = (kind: ChartKind, width: number, height: number) =>
      estimateOfficeShapes(
        buildChart({
          kind,
          title: "t",
          width,
          height,
          data: { categories: ["A", "B"], series: [{ name: "s", values: [1, 2] }] },
        }),
      );
    const kinds = [...new Set(KIND_COST_ORDER)];
    // Every box a specimen plausibly gets: the slot is a quarter of a band, so
    // narrow, while the band's height varies with the deck.
    for (const [w, h] of [
      [50, 60],
      [70, 90],
      [100, 90],
      [160, 120],
    ]) {
      const counts = kinds.map((k) => at(k, w, h));
      const spread = Math.max(...counts) - Math.min(...counts);
      expect(
        spread,
        `at ${w}x${h} the specimens are ${kinds.map((k, i) => `${k} ${counts[i]}`).join(", ")} — ` +
          "too far apart to hold occupancy equal across the palindrome",
      ).toBeLessThanOrEqual(4);
      // Single-batch, or `last batch settled` reports a tail instead of a draw.
      expect(Math.max(...counts), `a specimen at ${w}x${h} spans more than one ten-shape batch`).toBeLessThanOrEqual(
        10,
      );
    }
  });

  /**
   * A REFUSAL IS A KIND FACT. A STALL IS WEATHER. The first version could not
   * tell them apart and fired a false red on its first real outing.
   *
   * 2026-09-05, on a 45-slide deck: `funnel` missed both slots while clustered
   * and waterfall took both, and the verdict called it "a kind this host will
   * not take". Re-run alone on the same deck twenty minutes later, 8 of 8
   * specimens landed, funnel included.
   *
   * That was predictable rather than unlucky. With eight slots in four mirrored
   * pairs, if failures fall at random then SOME kind loses both its slots 14% of
   * the time on two failures, 43% on three, 77% on four. That run had three.
   */
  it("fails on a kind the host REFUSED, and not on one the host stalled on", () => {
    const ok = (kind: ChartKind): KindDraw => ({ kind, drew: true, why: "" });
    const refused = (kind: ChartKind): KindDraw => ({ kind, drew: false, why: "GeneralException" });
    const stalled = (kind: ChartKind): KindDraw => ({
      kind,
      drew: false,
      stalled: true,
      why: "the host stopped answering",
    });

    const all = kindCostVerdict([ok("clustered"), ok("area"), ok("area"), ok("clustered")]);
    expect(all.ok).toBe(true);
    expect(all.detail).toMatch(/4 of 4 specimens landed across 2 kinds/);

    // THE FINDING: refused in BOTH positions. Invisible in a battery where every
    // other scenario draws clustered.
    const dead = kindCostVerdict([ok("clustered"), refused("area"), refused("area"), ok("clustered")]);
    expect(dead.ok, "a kind the host refuses twice is a product fact").toBe(false);
    expect(dead.detail, "did not name the kind").toMatch(/^area was REFUSED in either position/);

    // THE REGRESSION: the exact shape of the false red. Both slots missed, both
    // by the host going quiet, while other kinds drew.
    const weather = kindCostVerdict([ok("clustered"), stalled("area"), stalled("area"), ok("clustered")]);
    expect(weather.ok, "called the host's bad minute a broken chart kind").toBe(true);
    expect(weather.detail, "did not say the reading is worthless").toMatch(/says nothing about the kind/);
    expect(weather.detail, "did not say what to do about it").toMatch(/re-run it alone/);
    expect(weather.detail).not.toMatch(/REFUSED/);

    // One refusal among stalls is still a refusal — the kind never drew and at
    // least once the host said no rather than nothing.
    const mixed = kindCostVerdict([ok("clustered"), stalled("area"), refused("area"), ok("clustered")]);
    expect(mixed.ok, "a refusal hidden among stalls was let through").toBe(false);

    // NOT a failure: the intermittent stall that takes one specimen and leaves
    // the kind's other position standing.
    const flaky = kindCostVerdict([ok("clustered"), stalled("area"), ok("area"), ok("clustered")]);
    expect(flaky.ok, "went red on the intermittent stall").toBe(true);
    expect(flaky.detail).toMatch(/every kind drew at least once/);
    expect(flaky.detail).not.toMatch(/REFUSED/);

    // Nothing attempted is a skip, not a pass: no specimen means no reading.
    const none = kindCostVerdict([]);
    expect(none.skipped).toBe(true);
    expect(none.ok).toBe(false);
  });
});

describe("what a stalled scenario reports about the call it gave up on", () => {
  /**
   * Seven real-host rounds, thirteen abandoned calls, and not one of them ever
   * came back — the slowest of 327 batches that DID answer took 29.2s against a
   * 45-second budget, so the band between them is empty and a stall looks like
   * death rather than slowness.
   *
   * That was read out of a trace line's ABSENCE, which is the inference this
   * project has misread more than any other: `settleUntaggedCharts` was
   * diagnosed as "ran and failed" for two sessions when it had never run at
   * all. A log that only writes on yes cannot tell "no" from "not asked". So
   * the verdict says which, in words, on every stall.
   */
  const drew = "PowerPoint did not respond while drawing shapes 1-10 of 24 (45s)";

  it("says the call never came back when nothing answered", () => {
    const d = stallDetail(drew, undefined);
    expect(d).toContain("nothing was checked");
    expect(d, "a silent host must be reported as silent, not left ambiguous").toMatch(/still not answered/);
    expect(d, "no late answer arrived, so nothing may claim one did").not.toMatch(/DID come back/);
  });

  it("says the host was merely slow when the call answers late", () => {
    // The reading that has never yet been observed on a real host, and the
    // whole reason the question is asked out loud: the day it happens, the
    // round must say so rather than wait for someone to notice a message that
    // has stopped being missing.
    const d = stallDetail(drew, "drawing shapes 1-10 of 24: the host eventually SUCCEEDED after 61s");
    expect(d).toMatch(/DID come back/);
    expect(d).toContain("eventually SUCCEEDED after 61s");
    expect(d, "a call that answered must not also be reported as silent").not.toMatch(/still not answered/);
  });

  it("names the call the host last answered before the stall", () => {
    // The half no round file has ever carried, and the reason the stalls could
    // only ever be described at SCENARIO level.
    //
    // Every draw stall on record is the FIRST batch of a scenario's draw, eight
    // for eight — so what the host did immediately before that sync is the
    // question. The log answered it for nobody: between a scenario announcing
    // itself and its first `batch issued` there is not one entry, while
    // three to five seconds of probe reads, deck inventories and selection
    // calls go by.
    //
    // So the only account available was "it follows the selection ladder" — and
    // the battery's order is FIXED, which makes "which scenario ran before" and
    // "which position in the battery" the same variable. No number of rounds
    // separates those. The call before the stall is a different variable.
    const d = stallDetail(drew, undefined, { call: "selecting a shape", idleMs: 430 });
    expect(d).toContain('the last thing the host answered was "selecting a shape"');
    expect(d, "the gap must be the idle time, not the 45s budget").toContain("0.4s earlier");
  });

  it("says nothing about a predecessor when there was not one", () => {
    // A stall on the run's very first bounded call has no predecessor, and
    // inventing one — or printing an empty quotation — is how a reader ends up
    // diagnosing a call that never happened.
    const d = stallDetail(drew, undefined);
    expect(d).not.toContain("the last thing the host answered");
    expect(d).toMatch(/still not answered/);
  });

  it("carries the call before the stall through a real stalled scenario", async () => {
    // The plumbing half, on the DRAW path. `lastStall` is written by the
    // deadline handler in `withTimeout`, and it has to still hold the PREVIOUS
    // completed call — a call that never answered must not overwrite it with
    // itself, which would make every stall report that it followed itself.
    const { _setBatchTimeoutForTest, _resetStallContextForTest } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    _resetStallContextForTest();
    _setBatchTimeoutForTest(5);
    try {
      for (let k = 1; k <= 60; k++) stallSyncOn.add(trips.syncs + k);
      const r = byName(await runSelfTest("probe", "the chart is actually visible"))["the chart is actually visible"];
      stallSyncOn.clear();
      expect(r.skipped, `expected a stall, got: ${r.detail}`).toBe(true);
      expect(r.detail, "the stall did not name what the host last answered").toContain(
        "the last thing the host answered was",
      );
      expect(r.detail, "a stalled call was reported as its own predecessor").not.toContain(
        'answered was "drawing shapes 1-10',
      );
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncOn.clear();
    }
  }, 20_000);

  it("leaves no trace line claiming a batch landed when that batch is the one that stalled", async () => {
    // The per-batch line is written one statement BEFORE the sync it describes,
    // deliberately — the sync is where a bad host goes quiet, so the number has
    // to be on screen while you wait. For as long as it existed it was called
    // `batch committed`, which made that deliberate ordering a lie: every stall
    // on record left behind a line saying the batch it killed had committed.
    //
    // This is not pedantry about a word. Two hand analyses of the round files
    // died on it — one paired the lines with draws and reported 0 stalls in 32,
    // the other counted them as successes and manufactured a 6x rasterise
    // effect that was not there — and `scripts/triage.mjs` carries a comment
    // block whose entire job is to warn the next reader off it. The round of
    // 2026-08-11 has the clean instance: one line at 351.6s reporting
    // `upTo:10 total:24`, and 45 seconds later `gave up waiting` for that exact
    // batch, with the same `idleMs` and the same predecessor. Same batch, and
    // the log calls it committed.
    //
    // Asserted as the property rather than against the new spelling: whatever
    // this line is named, a batch the host never answered must not have left
    // one claiming it did.
    const { _setBatchTimeoutForTest, _resetStallContextForTest } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    _resetStallContextForTest();
    setTracing(true);
    _setBatchTimeoutForTest(5);
    try {
      for (let k = 1; k <= 60; k++) stallSyncOn.add(trips.syncs + k);
      await runSelfTest("probe", "the chart is actually visible");
      stallSyncOn.clear();
      const entries = traceLog().entries;
      const stalled = entries.filter((e) => e.message === "gave up waiting");
      expect(stalled.length, "nothing stalled, so this asserts nothing").toBeGreaterThan(0);
      // Every draw line the run left behind, from a run in which no batch ever
      // got an answer out of the host.
      const drawLines = entries.filter((e) => e.scope === "draw").map((e) => e.message);
      expect(drawLines.length, "the draw never reported a batch at all").toBeGreaterThan(0);
      for (const m of drawLines) {
        expect(
          m,
          `a batch the host never answered left the log saying "${m}" — a reader counting those counts a stall as a success`,
        ).not.toMatch(/committed|landed|done|succeeded/i);
      }
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncOn.clear();
    }
  }, 20_000);

  it("records the idle gap on the first batch of a draw, and only the first", async () => {
    // The baseline the stall record is worthless without.
    //
    // Round 7 was the first to name the call before a stall — `afterAnswering:
    // "rasterising a slide", idleMs: 1` — and 1ms looks damning until you ask
    // what the batches that SURVIVE report. Sequential code issues its next
    // call the instant the previous one answers, so a 1ms gap may be true of
    // every draw in the round. A number with no baseline is not a measurement.
    //
    // First batch only: a later batch's predecessor is always the batch before
    // it, which says nothing and would cost three lines a chart in the log.
    const { insertSceneIntoSlide } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    setTracing(true);
    try {
      await insertSceneIntoSlide(buildChart(sampleConfig("stacked")), { tagData: "{}" });
      const batches = traceLog().entries.filter((e) => e.message === "batch issued");
      expect(batches.length, "the draw did not batch, so there is nothing to check").toBeGreaterThan(1);
      expect(batches[0].data, "the first batch carries no idle gap to compare a stall against").toHaveProperty(
        "idleMs",
      );
      expect(typeof batches[0].data?.idleMs).toBe("number");
      // And the predecessor's NAME, which spent two rounds in exactly the
      // condition the gap was in: recorded on stalls, never on successes. Round
      // 9 produced two stalls naming two different calls while thirteen draws
      // survived without saying what they followed, so the comparison nobody
      // could make was the whole question.
      expect(
        batches[0].data,
        "the first batch does not say what the host last answered, so a stall's predecessor has nothing to be compared against",
      ).toHaveProperty("afterAnswering");
      // And how long that call TOOK. The name alone could not split the
      // 2026-08-11 control apart: its four arms rasterise a slide and then draw
      // seven shapes, and they came back 22.7s, 25.6s and 28.9s — one lump per
      // arm, with the rasterise and the draw inside it and no way to say which
      // half was growing. `withTimeout` already stamps both ends of every named
      // call, so the duration is a subtraction at a seam that exists.
      expect(
        batches[0].data,
        "the first batch names the call before it but not how long that call took, so a rasterise and the draw after it stay one number",
      ).toHaveProperty("afterAnsweringMs");
      expect(typeof batches[0].data?.afterAnsweringMs).toBe("number");
      for (const b of batches.slice(1)) {
        expect(b.data, "a later batch timed a predecessor that is always the batch before it").not.toHaveProperty(
          "afterAnsweringMs",
        );
        expect(b.data, "a later batch named a predecessor that is always the batch before it").not.toHaveProperty(
          "afterAnswering",
        );
        expect(
          b.data,
          "a later batch reported an idle gap that can only describe the batch before it",
        ).not.toHaveProperty("idleMs");
      }
    } finally {
      setTracing(false);
    }
  }, 20_000);

  it("carries the late answer through a real stalled scenario", async () => {
    const { _setBatchTimeoutForTest } = await import("../src/render/powerpoint");
    // The plumbing half. `stallDetail` is pure and cannot know whether the
    // battery ever hands it a late answer — and until this ran, the draw path's
    // late-answer trace had no test at all: every existing one goes through
    // `insertDemoDeck`. The conclusion above rests on the draw path reporting
    // late answers when they happen, so that has to be checked, not assumed.
    installHost([makeSlide("s1")]);
    _setBatchTimeoutForTest(5);
    try {
      // The fake settles a stalled sync 40ms in, well past the 5ms deadline —
      // so the scenario gives up and is then told how it went.
      for (let k = 1; k <= 60; k++) stallSyncOn.add(trips.syncs + k);
      const r = byName(await runSelfTest("probe", "the chart is actually visible"))["the chart is actually visible"];
      stallSyncOn.clear();
      expect(r.skipped, `expected a stall, got: ${r.detail}`).toBe(true);
      expect(r.blind).toBe(true);
      expect(r.detail, "the host answered late and the report did not say so").toMatch(/DID come back/);
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncOn.clear();
    }
  }, 20_000);
});

/**
 * The battery returns one array, after the last scenario, and the round writes
 * its file from that — so a tab that died mid-battery took every finished
 * verdict with it. Ordering `SCENARIOS` can only ever choose WHICH verdicts a
 * crash costs; reporting each one as it lands is what makes it cost none of the
 * ones already reached.
 */
describe("every verdict is reported as it lands, not only on return", () => {
  it("reports each scenario before the battery returns", async () => {
    const seen: string[] = [];
    const results = await runSelfTest("probe", undefined, (r) => seen.push(r.name));
    expect(results.length, "no scenarios ran, so this proves nothing").toBeGreaterThan(0);
    // Same verdicts, same order — the callback is the array being built, not a
    // summary of it.
    expect(seen).toEqual(results.map((r) => r.name));
  });

  it("reports the 'not reached' verdicts too", async () => {
    // These are the ones a future edit forgets, and they are the verdicts of a
    // run that is already going wrong — which is when this matters most.
    const seen: ScenarioResult[] = [];
    const results = await runSelfTest("probe", undefined, (r) => seen.push(r));
    const skipped = results.filter((r) => r.skipped);
    for (const s of skipped) {
      expect(seen.some((r) => r.name === s.name && r.skipped)).toBe(true);
    }
  });

  it("survives a sink that throws", async () => {
    // A reporting sink is never a reason for the battery to stop.
    const results = await runSelfTest("probe", undefined, () => {
      throw new Error("storage is full");
    });
    expect(results.length).toBeGreaterThan(0);
  });
});

/**
 * What a round records BESIDE its verdicts.
 *
 * Four fields in this project have been built and then thrown away because they
 * were written only when something went wrong — `idleMs`, `afterAnswering`, the
 * settle's shared label, `listChartsInDeck`. Each was a number with no
 * population to compare against. Everything asserted here is therefore asserted
 * on a scenario that PASSED: if a field is only present on failures it is not a
 * measurement, and this is the test that says so.
 */
describe("the context a round records around every scenario", () => {
  it("stamps each scenario with the deck it ran on and the friction it met", async () => {
    installHost([makeSlide("s1")]);
    setTracing(true);
    await runSelfTest("probe", "the chart is actually visible");
    const started = traceLog().entries.filter((e) => e.message === "scenario starting");
    const ended = traceLog().entries.filter((e) => String(e.message).startsWith("scenario "));
    expect(started.length, "no scenario announced itself").toBeGreaterThan(0);
    for (const e of started) {
      expect(e.data, "a scenario began without saying what deck it was on").toHaveProperty("deckSlides");
      expect(e.data, "a scenario began without saying how far into the round it was").toHaveProperty("atMs");
    }
    const verdicts = ended.filter((e) => e.message !== "scenario starting");
    expect(verdicts.length, "no scenario reported a verdict").toBeGreaterThan(0);
    for (const e of verdicts) {
      // On the verdict, whatever the verdict is — including "passed".
      expect(e.data, `${String(e.data?.name)} reported no friction count`).toHaveProperty("friction");
      const f = e.data?.friction as Record<string, number>;
      for (const k of ["errors", "idRefusals", "generalExceptions", "emptyReReads"]) {
        expect(typeof f[k], `friction.${k} is not a number`).toBe("number");
      }
      expect(e.data, `${String(e.data?.name)} reported no deck size`).toHaveProperty("deckSlides");
    }
  }, 30_000);

  it("records the on-slide shape count and the previous batch's duration on every batch", async () => {
    // The input to the quadratic per-slide cost, which is this project's main
    // performance claim and had never been recorded, plus the batch duration
    // that until now had to be differenced out of consecutive timestamps — and
    // so silently included every bit of inter-batch work.
    const { insertSceneIntoSlide } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    setTracing(true);
    await insertSceneIntoSlide(buildChart(sampleConfig("stacked")), { tagData: "{}", slideId: "s1" });
    const batches = traceLog().entries.filter((e) => e.message === "batch issued");
    expect(batches.length, "the draw did not batch").toBeGreaterThan(1);
    for (const b of batches) {
      expect(b.data, "a batch did not say how loaded its slide already was").toHaveProperty("onSlide");
      expect(typeof b.data?.onSlide).toBe("number");
      // And which slide it is counting, so a reader can tell an accumulation on
      // one slide from one that may span several.
      expect(b.data, "a batch counted shapes without saying what it counted them on").toHaveProperty("onSlideKey");
    }
    // The first batch has no predecessor, every later one does.
    expect(batches[0].data, "the first batch invented a previous batch").not.toHaveProperty("prevBatchMs");
    for (const b of batches.slice(1)) {
      expect(b.data, "a batch did not say how long the one before it took").toHaveProperty("prevBatchMs");
      expect(typeof b.data?.prevBatchMs).toBe("number");
    }
    // And it ACCUMULATES — the whole point is that the nth batch reports more
    // than the first, because that is the variable the cost grows with.
    const counts = batches.map((b) => b.data?.onSlide as number);
    expect(counts[counts.length - 1], "the on-slide count never grew, so it is not counting").toBeGreaterThan(
      counts[0],
    );
  }, 30_000);

  it("records the on-slide count on EVERY batch, not only when the caller named a slide", async () => {
    // It was conditional on `opts.slideId`, which most draws do not pass: in its
    // first real round the field appeared on SEVEN of forty-six batches, and on
    // exactly ONE alongside `prevBatchMs`. A number missing five times out of
    // six cannot answer the question it exists for — the per-slide cost curve
    // this project has asserted for weeks without ever measuring its input.
    const { insertSceneIntoSlide } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    setTracing(true);
    // No `slideId` — the ordinary case, and the one that was blank before.
    await insertSceneIntoSlide(buildChart(sampleConfig("stacked")), { tagData: "{}" });
    const batches = traceLog().entries.filter((e) => e.message === "batch issued");
    expect(batches.length).toBeGreaterThan(1);
    for (const b of batches) {
      expect(b.data, "an unnamed target left the count blank").toHaveProperty("onSlide");
      expect(typeof b.data?.onSlide).toBe("number");
    }
    const counts = batches.map((b) => b.data?.onSlide as number);
    expect(counts[counts.length - 1], "the count never grew, so it is not counting").toBeGreaterThan(counts[0]);
  }, 30_000);
});

/**
 * The battery used to pile its whole run onto one slide.
 *
 * Every scenario needing "a chart to work with" took `found[0]` — the same
 * chart on the same slide, every time. Drawing cost on this host grows with
 * what is already on the target (about +0.44s per shape present, measured), so
 * the battery made each of its own later draws slower than the last.
 *
 * Round `275a76a` put numbers on it: one slide went 20 → 68 → 92 → 116 → 140 →
 * 144 → 165 shapes while nothing else in the deck passed 34, and the draw at
 * 144 stalled — a NINE-shape chart, timed out at 45s, which ~63s of per-slide
 * overhead entirely accounts for. A scenario ordered late was measurably harder
 * to pass than the same scenario ordered early.
 */
describe("which chart a scenario picks to work with", () => {
  const chart = (slideId: string, name: string) => ({ target: { slideId }, name });

  it("takes the one on the slide this run has loaded least", () => {
    const charts = [chart("busy", "a"), chart("quiet", "b")];
    const load = (id: string) => (id === "busy" ? 144 : 0);
    expect(leastLoadedChart(charts, load)?.name).toBe("b");
  });

  it("keeps the deck's order when nothing has been drawn yet", () => {
    // The property that makes this safe to drop in: on a fresh run every load
    // is zero, so it picks exactly what `found[0]` picked and every existing
    // expectation about which chart a scenario takes still holds. It only
    // diverges once this run has actually loaded a slide.
    const charts = [chart("a", "first"), chart("b", "second"), chart("c", "third")];
    expect(leastLoadedChart(charts, () => 0)?.name).toBe("first");
    // And on a tie between two equally-loaded slides, still the earlier one.
    expect(leastLoadedChart(charts, (id) => (id === "a" ? 5 : 5))?.name).toBe("first");
  });

  it("has nothing to offer from an empty deck", () => {
    expect(leastLoadedChart([], () => 0)).toBeUndefined();
  });

  it("spreads the battery's draws across slides instead of onto one", async () => {
    // End to end: the shape of the defect was a single slide taking almost
    // every draw. Asserted as a share rather than a count, because the totals
    // depend on which scenarios the fake lets run.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    await runSelfTest("probe");
    const scan = await listChartsInDeck();
    const slides = [...new Set(scan.charts.map((c) => c.target.slideId))];
    expect(slides.length, "the run produced only one slide, so there was nothing to spread over").toBeGreaterThan(1);
    const loads = slides.map((id) => shapesDrawnOn(id));
    const total = loads.reduce((a, b) => a + b, 0);
    if (total === 0) return; // the fake drew nothing measurable; nothing to claim
    const worst = Math.max(...loads);
    // Measured, not invented. Against the fake, the same battery concentrates
    // 0.619 of its draws on one slide with `found[0]` and 0.369 with this rule
    // — 195 shapes on the worst slide versus 120, and the others rising off a
    // flat 24. Half sits between the two with room on both sides.
    //
    // The fake spreads more than the real host does to begin with (its probe
    // charts land one per slide, where a real deck had nine on one), so this
    // understates the effect it is guarding rather than overstating it.
    expect(
      worst / total,
      `one slide took ${worst} of ${total} shapes this run drew — the battery is concentrating again`,
    ).toBeLessThan(0.5);
  }, 60_000);
});

/**
 * A slide that reads back EMPTY right after a draw has not lost its shapes.
 *
 * Round `957aca0` failed `explode a degraded picture` with `the collapse added
 * 0 shapes (none) — the slide went from 1 to 0`, and the deck inventory taken
 * at the end of the same run shows that slide holding one shape named
 * `PowerChart`. The chart never moved: the host refused a collection read on a
 * slide the run had just drawn on, and the scenario reported the refusal as
 * data loss.
 *
 * `slideShapeList` is careful — it corroborates `items` against `getCount()`
 * and answers null when they disagree — so this is a hole in that strategy
 * rather than in the scenario: both signals can agree at ZERO and both be
 * wrong. Nothing downstream can tell from the return value alone.
 *
 * It is also the third time this scenario has produced a false FAILURE from
 * counting a slide, and its own comments record the other two.
 */
/**
 * A null return from an update is not a report that the chart was destroyed.
 *
 * Round `ee1741e` failed this scenario with `the picture vanished while being
 * exploded back to shapes`, and the deck inventory taken at the end of the same
 * run shows that slide holding one shape named `PowerChart`. Nothing vanished:
 * the host refused the shape id — three `InvalidParam passed to GetItem(id)` in
 * the same second — so the update could not work on it again.
 *
 * The same defect as the empty-read one below, one step later in the same
 * scenario, and it matters for the same reason: "the add-in deleted a chart"
 * and "the host would not answer for it" send a maintainer to opposite ends of
 * the codebase, and only one of them is true.
 */
/**
 * Telling our defects from this host's weather.
 *
 * The headline counted every red scenario together, which made it useless for
 * the only question worth asking of a series of rounds: is the add-in getting
 * better? On this host most red is the shape collection dying part-way through
 * — `same scale across the deck` is largely a measurement of WHEN that happens
 * — so the number moved with the host's mood and never with our work.
 *
 * The rule is checked here against a round whose answer was known before the
 * rule existed. On `89675b6`, `explode a degraded picture` failed with zero
 * friction (the picture regression: a pure logic bug, no host involvement) and
 * `same scale across the deck` failed with six id refusals and four empty
 * re-reads. Those are the two real cases, and they have to come out opposite.
 */
/**
 * The one recurring stall the battery had been reporting anonymously.
 *
 * `a selected shape survives an insert` stalled its first draw batch in four of
 * the last five rounds (`957aca0`, `ee1741e`, `89675b6`, `47a80c8`) after
 * passing eight running before that — and every round reported it with the
 * runner's generic "the host got in the way", so a specific repeating
 * observation was being thrown away each time.
 *
 * The note it carries now is bounded by what the round files support. It does
 * NOT blame the preceding call: `selecting a shape` precedes both this stall
 * and surviving draws in all four of those rounds. It says only the one way
 * this draw differs from every other in the battery — it is made with a
 * selection standing.
 */
describe("the draw made while a shape is selected", () => {
  it("says what stalled instead of reporting an anonymous host failure", async () => {
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    // Let the selection succeed and the draw time out, which is the shape the
    // real host produces: `drawing shapes 1-10 of 24`, first batch, gone.
    faults.stallDrawAfterSelect = true;
    _setBatchTimeoutForTest(10);
    try {
      const r = byName(await runSelfTest("probe"))["a selected shape survives an insert"];
      expect(r.detail, `the stall was reported anonymously: ${r.detail}`).toMatch(/while a shape was SELECTED/);
      expect(r.detail, "the note does not say what protects against it").toMatch(/dropShapeSelection/);
      expect(r.blind, "a stalled draw was reported as evidence about the product").toBe(true);
    } finally {
      faults.stallDrawAfterSelect = false;
      _setBatchTimeoutForTest(45_000);
    }
  }, 60_000);
});

describe("whose fault a red scenario was", () => {
  const res = (over: Partial<ScenarioResult>): ScenarioResult => ({
    name: "s",
    ok: false,
    detail: "",
    ms: 1,
    ...over,
  });
  // Every counter, zeroed. Explicit rather than derived, and safe to be:
  // `SelfTestResult.friction` requires all of them, so adding a ninth counter
  // fails typecheck here until this follows. The compiler maintains the list,
  // which is what the four-name version never had.
  const clean = {
    errors: 0,
    idRefusals: 0,
    generalExceptions: 0,
    emptyReReads: 0,
    shortReReads: 0,
    unmatchedReReads: 0,
    reReadsRepaired: 0,
  };

  it("blames US for a failure the host did not interfere with", () => {
    // Round `89675b6`, `explode a degraded picture` — the picture regression.
    expect(scenarioBlame(res({ friction: clean }))).toBe("ours");
  });

  it("blames the HOST when it refused something inside the scenario", () => {
    // Round `89675b6`, `same scale across the deck` — the case this rule was
    // built on, and it still reads the same way.
    expect(scenarioBlame(res({ friction: { ...clean, errors: 6, idRefusals: 6, emptyReReads: 4 } }))).toBe("host");
    expect(scenarioBlame(res({ friction: { ...clean, emptyReReads: 1 } }))).toBe("host");
    // An id refusal the code could NOT repair is still the host.
    expect(scenarioBlame(res({ friction: { ...clean, idRefusals: 2, reReadsRepaired: 1 } }))).toBe("host");
  });

  it("does not blame the host for a refusal the code recovered from", () => {
    /**
     * THE RULE WAS RIGHT ABOUT ONE SCENARIO AND MEANINGLESS ABOUT ANOTHER.
     * Measured over all 322 archived rounds:
     *
     *     same scale across the deck    14% of passes carry the signature, 100% of failures
     *     explode a degraded picture    99% of passes,                       98% of failures
     *
     * For `explode` the condition is simply always true, so every failure of
     * the add-in's most-failing update path was written off as the host's
     * weather on a reading a passing round shows just as often. Getting the
     * right answer from a signal that carries no information is luck.
     *
     * A refusal the renderer re-read and RECOVERED from is not what defeated
     * the run — `reReadsRepaired` counts the times that worked. With repaired
     * refusals discounted, `explode` moves to 48% of passes against 98% of
     * failures, a 50-point gap, while `same scale` does not move at all.
     */
    expect(
      scenarioBlame(res({ friction: { ...clean, idRefusals: 1, reReadsRepaired: 1 } })),
      "a refusal the code recovered from was blamed on the host",
    ).toBe("ours");

    /**
     * AND `generalExceptions` LEAVES THE TEST, which reverses an assertion this
     * file used to make.
     *
     * It read as principled — a general exception is the host misbehaving — and
     * for `explode a degraded picture` it is exactly 1 in every single round.
     * The friction report names it in as many words: "a constant, not a signal".
     * A constant cannot be evidence, and leaving it in is what made the whole
     * rule ambient for that scenario.
     *
     * Checked before changing: no archived verdict moves. All 109 recorded
     * failures carrying friction are blamed identically by both rules.
     */
    expect(
      scenarioBlame(res({ friction: { ...clean, generalExceptions: 1 } })),
      "a counter that is 1 in every round, pass or fail, was treated as evidence",
    ).toBe("ours");
  });

  it("defaults to US when there is no evidence either way", () => {
    // The direction matters more than the split: getting it backwards turns
    // this into a way to make failures disappear. Unproven lands on us.
    expect(scenarioBlame(res({}))).toBe("ours");
    // `errors` alone is not a refusal — a stall raises it, and a stall is
    // already reported as `blind`/not-run rather than as a failure.
    expect(scenarioBlame(res({ friction: { ...clean, errors: 3 } }))).toBe("ours");
  });

  it("keeps passes and unrun scenarios out of the blame counts", () => {
    expect(scenarioBlame(res({ ok: true, friction: clean }))).toBe("passed");
    expect(scenarioBlame(res({ skipped: true, friction: clean }))).toBe("not-run");
  });

  it("leads the summary with our defects, and never hides the host's", () => {
    const line = describeSelfTest([
      res({ name: "explode a degraded picture", friction: clean }),
      res({ name: "same scale across the deck", friction: { ...clean, idRefusals: 6, emptyReReads: 4 } }),
      res({ name: "insert on top of an earlier run", ok: true, friction: clean }),
    ]);
    expect(line, "our defect is not the headline").toMatch(/^Self-test — 1 defect\(s\) of ours: explode/);
    expect(line, "a host-degraded failure was hidden rather than named").toContain("same scale across the deck");
    expect(line).toMatch(/failed while the host was refusing/);
  });

  it("says so plainly when nothing was our fault", () => {
    const line = describeSelfTest([
      res({ name: "same scale across the deck", friction: { ...clean, idRefusals: 6 } }),
      res({ name: "edit a chart on the visible slide", ok: true, friction: clean }),
    ]);
    expect(line).toMatch(/^Self-test — no defects of ours/);
  });
});

describe("what a failed update is allowed to claim", () => {
  it("does not report a refused id as the chart being destroyed", () => {
    expect(updateLossNote("picture", 3)).toMatch(/would not name the picture/);
    expect(updateLossNote("picture", 3), "a refusal was reported as destruction").not.toMatch(/GONE|destroyed/);
  });

  /**
   * The case that survived the first fix, found on round `eaddbf4`.
   *
   * The explode's update returned null having logged nothing and refused
   * nothing INSIDE that call, so keying on thrown id refusals alone printed
   * `the picture vanished while being redrawn` — while the deck inventory from
   * the same run showed a chart on every slide. This host can fail to resolve a
   * target quietly, with no throw to count, and a count of throws cannot see it.
   *
   * So destruction is now claimed only on positive evidence: the slide was
   * asked and said the shape is not there.
   */
  it("never claims destruction without asking the slide", () => {
    // Nothing refused, nothing known — the honest answer, and NOT a loss claim.
    const blind = updateLossNote("picture", 0);
    expect(blind, "an unexplained null was reported as destruction").not.toMatch(/GONE|destroyed|vanish/);
    expect(blind).toMatch(/would not say what became of it/);
  });

  it("says so plainly when the slide confirms the shape is still there", () => {
    expect(updateLossNote("picture", 3, true)).toMatch(/STILL ON THE SLIDE/);
    expect(updateLossNote("picture", 0, true)).toMatch(/nothing was lost/);
  });

  it("reports a real destruction when the slide says the shape is gone", () => {
    // The one case that earns the loud wording — and it has to stay reachable,
    // or the guards would be satisfied by never claiming anything. Needs a
    // CLEAN scenario: no id refusal anywhere in it.
    expect(updateLossNote("picture", 0, false)).toMatch(/GONE from the slide/);
  });

  /**
   * The fourth mechanism, from round `1789749`.
   *
   * The collapse's readback was refused, so the target handed on carried an id
   * the host never confirmed — the settle had to find that chart BY NAME
   * (`withId: 0`), which is the tell. The picture then landed under an id
   * nobody held, the id comparison answered "not there", and the verdict read
   * `the picture is GONE from the slide` while the deck inventory from the same
   * run shows that slide holding one shape named `PowerChart`.
   *
   * So a missing shape is only evidence of a missing shape on a scenario where
   * this host refused no ids at all. The count is measured from the scenario's
   * start, because the refusal that poisons an id is routinely in an earlier
   * step than the call that fails.
   */
  it("will not call a shape destroyed when an id was refused in the same scenario", () => {
    const v = updateLossNote("picture", 3, false);
    expect(v, "a refused id made a shape look destroyed").not.toMatch(/GONE|destroyed/);
    expect(v).toMatch(/proves nothing either way/);
  });
});

describe("the collapse's readback when the host will not list a slide", () => {
  it("does not report data loss for a read that came back empty", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    // The collection AND its count both answer zero, which is what the real
    // host did on the slide it had just drawn onto — and the only shape of it
    // that defeats `slideShapeList`'s corroboration.
    //
    // Armed by the picture landing rather than from the first read, because a
    // host that is blind from the start starves `probeCharts`, the scenario
    // skips for want of a chart, and this guard passes against the unfixed
    // build. That was the first version of it.
    faults.slideReadsEmpty = "after-a-picture";
    try {
      const explode = byName(await runSelfTest("probe"))["explode a degraded picture"];
      // Not `ok`, which the scenario may still fail for its own reasons on a
      // host this hostile: the defect is the SENTENCE, a claim that shapes
      // disappeared, and that is what must not be reachable from a refusal.
      expect(explode.detail, `an empty readback was reported as data loss: ${explode.detail}`).not.toMatch(
        /went from \d+ to 0/,
      );
    } finally {
      faults.slideReadsEmpty = null;
    }
  }, 60_000);
});

describe("the host-friction counters", () => {
  it("resets every counter, not the four that existed when the reset was written", async () => {
    // WHAT WENT WRONG: `resetHostFriction` named four of eight fields by hand.
    // `shortReReads`, `unmatchedReReads`, `settledByBinding` and `reReadsRepaired`
    // were each added later, each with a docstring about the round-level question
    // it answers, and none was added to the reset. Only the test seam calls it, so
    // the cost was isolation — one test's re-read failures leaking into the next
    // one's counts, the kind of pollution that makes a suite pass in one order and
    // fail in another.
    //
    // ASSERTED ON THE SOURCE, and the first draft of this test is why.
    //
    // It called the reset and checked every counter was zero. It passed with the
    // bug put back — because a freshly imported module has never incremented
    // anything, so all eight are zero whether the reset touches them or not. A
    // test that cannot fail is not a test, and this one was caught only by
    // mutating the fix away and watching it stay green.
    //
    // Reaching the real thing at runtime would mean driving the host paths that
    // increment each counter, from a test about a four-line function. The
    // property actually wanted is not a runtime value at all: it is that the
    // reset CANNOT NAME FIELDS ONE BY ONE, because a list is what went stale.
    // So that is what is asserted.
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    const fn = src.slice(src.indexOf("function resetHostFriction(): void {"));
    const body = fn.slice(0, fn.indexOf("\n}"));
    expect(body, "the reset went back to a hand-maintained list").not.toMatch(/hostFriction\.\w+ = 0;/);
    expect(body, "the reset no longer covers the object's own keys").toContain("Object.keys(hostFriction)");

    // And the counter set is still the eight this is written about, so a shrink
    // to four somewhere else does not quietly satisfy the above.
    const mod = await import("../src/render/powerpoint");
    expect(
      Object.keys(mod.hostFrictionCounts()).length,
      // SEVEN, not eight: `settledByBinding` went with the route it counted,
      // which rescued 0 config tags in 180 rounds and was refused 16 times.
      // Asserted as a number rather than left open, so shrinking the set again
      // has to be a decision somebody writes down instead of a quiet drift.
      "expected the seven documented counters",
    ).toBeGreaterThanOrEqual(7);
  });

  it("puts every counter on the run-finished line, including the one nothing else reads", async () => {
    // `unmatchedReReads` is incremented where the re-read names none of a chart's
    // shapes and is referenced nowhere else in src/. It counts the same event as
    // the trace line beside it, which makes it a second source for a number this
    // project reasons about constantly — and it reached no archive, so the
    // cross-check could not be run at all.
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    expect(src, "the run-finished line no longer emits the counters").toContain("frictionSoFar: hostFrictionCounts(),");
    // The emitted value is the whole snapshot, so its keys are the real keys —
    // named individually here only because these four are the ones that were
    // missing from every hand-written list in the repo.
    const mod = await import("../src/render/powerpoint");
    const emitted = Object.keys(mod.hostFrictionCounts());
    expect(emitted, "unmatchedReReads is the counter this was written for").toContain("unmatchedReReads");
    expect(emitted).toContain("shortReReads");
    expect(emitted).toContain("reReadsRepaired");
  });
});

describe("the pause that breaks the scan/first-chart confound", () => {
  it("is off unless someone asks for it", async () => {
    const { scanSettleMs } = await import("../src/taskpane/selftest");
    // An instrument that is on by default is not a control arm, it is a
    // silent product change that every later round is measured against.
    expect(scanSettleMs(() => null)).toBe(0);
    expect(scanSettleMs(() => "")).toBe(0);
    expect(scanSettleMs(() => "   ")).toBe(0);
  });

  it("refuses a value that would park the pane rather than pausing it", async () => {
    const { scanSettleMs, SCAN_SETTLE_MAX_MS } = await import("../src/taskpane/selftest");
    // `Number("abc")` is NaN and a NaN setTimeout fires IMMEDIATELY — so an
    // unparseable value would silently produce the control arm while the round
    // recorded it as the treated one.
    expect(scanSettleMs(() => "abc")).toBe(0);
    expect(scanSettleMs(() => "-5")).toBe(0);
    expect(scanSettleMs(() => "0")).toBe(0);
    // A typo of an extra three zeros costs a round, not an hour.
    expect(scanSettleMs(() => "9999999")).toBe(SCAN_SETTLE_MAX_MS);
  });

  it("takes a value it can use", async () => {
    const { scanSettleMs } = await import("../src/taskpane/selftest");
    expect(scanSettleMs(() => "3000")).toBe(3000);
    expect(scanSettleMs(() => "2500.6")).toBe(2501);
  });

  it("treats storage that throws as nobody asking", async () => {
    const { scanSettleMs } = await import("../src/taskpane/selftest");
    // A pane with storage disabled must run the round, not fail it.
    expect(
      scanSettleMs(() => {
        throw new Error("SecurityError");
      }),
    ).toBe(0);
  });
});

describe("which chart the lone-run arm picks", () => {
  it("never picks a chart on the first chart's slide", async () => {
    const { pickLoneChart } = await import("../src/taskpane/selftest");
    // Rounds 213/214 ran this arm on charts[0] and so landed on slide 257,
    // which held 42 shapes against 20-21 elsewhere. Every expensive 18-of-24
    // sample sat on that slide and every cheap one sat elsewhere, so the arm
    // could not tell run-position from slide-occupancy.
    expect(pickLoneChart(["257", "257", "257", "258", "262"])).toBe(4);
    expect(pickLoneChart(["257", "258"])).toBe(1);
  });

  it("returns nothing when the deck cannot break the confound", async () => {
    const { pickLoneChart } = await import("../src/taskpane/selftest");
    // A number from a confounded arm is worse than no number.
    expect(pickLoneChart(["257", "257", "257"])).toBe(null);
    expect(pickLoneChart(["257"])).toBe(null);
    expect(pickLoneChart([])).toBe(null);
  });

  it("prefers the LAST differing slide, not the first it stumbles on", async () => {
    const { pickLoneChart } = await import("../src/taskpane/selftest");
    // The last chart of the run is the one furthest from the first chart in
    // both position and slide, which is the widest contrast the deck offers.
    expect(pickLoneChart(["257", "258", "259", "262"])).toBe(3);
  });
});

describe("the lone chart's other arm", () => {
  it("picks the busy slide when asked, size-matched", async () => {
    const { pickLoneChart } = await import("../src/taskpane/selftest");
    // Index 0 deliberately: it is the 18-of-24 chart, so it matches the
    // clear-slide arm. Its slide-mates are 9-of-16 and would compare nothing.
    expect(pickLoneChart(["257", "257", "257", "258", "262"], true)).toBe(0);
  });

  it("refuses the loaded arm when no chart shares that slide", async () => {
    const { pickLoneChart } = await import("../src/taskpane/selftest");
    // A deck that cannot pose the question must say so, the same rule the
    // default arm follows.
    expect(pickLoneChart(["257", "258", "259"], true)).toBe(null);
  });

  it("leaves the default arm exactly as it was", async () => {
    const { pickLoneChart } = await import("../src/taskpane/selftest");
    expect(pickLoneChart(["257", "257", "258", "262"])).toBe(3);
    expect(pickLoneChart(["257", "257"])).toBe(null);
  });

  it("stays on the clear arm unless the flag says otherwise", async () => {
    const { wantsLoadedLoneChart } = await import("../src/taskpane/selftest");
    expect(wantsLoadedLoneChart(() => null)).toBe(false);
    expect(wantsLoadedLoneChart(() => "0")).toBe(false);
    expect(wantsLoadedLoneChart(() => "yes")).toBe(false);
    expect(wantsLoadedLoneChart(() => "1")).toBe(true);
    expect(
      wantsLoadedLoneChart(() => {
        throw new Error("SecurityError");
      }),
    ).toBe(false);
  });
});

/**
 * `two slides claiming one slot` must not judge a window it cannot fill.
 *
 * Round 297 stopped a whole cycle on this. The scenario FAILED with `4 slides
 * inserted, 4 kept, 4 of 2 still re-editable; 0 queued as duplicates`, and it
 * is not the flake the scenario already has — all five earlier failures had the
 * dedup RUNNING and leaving slides behind. The trace says the first of two
 * inserts reported `landed: 0` for slides that did land, `slideCount()`
 * answered 5 when the deck held 7, and `reconcileDeck` was handed `[3,5]`: two
 * slides of four. Two distinct titles in that window, no duplicates, verdict
 * correct about the wrong population.
 *
 * `unread` was 0 on that read, which is the whole distinction. The host was not
 * refusing — `blindSkip` covers that and could not see this — it was BEHIND.
 */
describe("a deck count that is behind, not blind", () => {
  const NAME = "two slides claiming one slot";
  afterEach(() => {
    faults.slideCountCeiling = null;
    faults.slideCountCeilingBinds = 0;
    _setCountSettleDelayForTest(4_000);
  });

  /**
   * A CEILING OF SEVEN, TWO BINDS, reproduces round 297 step for step.
   *
   * `runSelfTest(prefix, only)` always runs the first two scenarios before the
   * named one, so the deck stands at 5 when this starts and reaches 9. A cap of
   * 7 therefore binds only here, and the two binds land where the round's did:
   * the first inside `insertSlidesFromPptx` — which reports `landed: 0` for
   * slides that DID land, the round's own signature — and the second on the
   * read that sizes the reconcile, which answers 7 for a deck of 9.
   *
   * Without the settle the window is `[5,7]`: two slides of four, and a
   * confident verdict about the wrong population.
   */
  it("still judges all four slides when the count settles late", () => {
    installHost([makeSlide("s1")]);
    faults.slideCountCeiling = 7;
    faults.slideCountCeilingBinds = 2;
    // Squashed so the test does not wait four real seconds for the settle.
    _setCountSettleDelayForTest(120);
    setTracing(true);
    return runSelfTest("probe", NAME).then((rs) => {
      setTracing(false);
      const r = byName(rs)[NAME];
      expect(r.skipped, `a recoverable lag was skipped: ${r.detail}`).toBeFalsy();
      // The population is the point: four inserted, and the verdict is about
      // four rather than about the two a short count would have shown.
      expect(r.detail, `judged the wrong population: ${r.detail}`).toContain("4 slides inserted");
      // AND THE BUG WAS ACTUALLY REPRODUCED. Without this the test passes on a
      // host that never lagged at all, which is how three earlier versions of
      // this fault looked correct while reproducing nothing — each one was
      // caught by the bug's own mutant surviving, not by a red test.
      const short = traceLog().entries.filter((e) => /count came back short/.test(e.message));
      expect(short.length, "the count never came back short, so nothing was reproduced").toBeGreaterThan(0);
      expect(short[0].data).toMatchObject({ counted: 7, wanted: 9 });
    });
  });

  /**
   * AND THE INSERT ITSELF STOPS REPORTING SLIDES THAT LANDED AS SLIDES THAT
   * DID NOT.
   *
   * The test above compensates one layer up: `runSelfTest` settles its own
   * count, so the scenario judged the right population while
   * `insertSlidesFromPptx` went on returning `landed: 0` — the docstring above
   * records that as the round's signature rather than as a defect. It is both.
   *
   * The archive says what the raw count costs a caller that has no second
   * layer. Rounds 148, 315, 375 and 422 each trace `landed: 0` here, then
   * `insert on top of an earlier run` failing with a GeneralException, then the
   * NEXT insert of two slides measuring `landed: 4` — the two extra being the
   * ones this call had just said did not arrive. Four of 1,641 traced inserts,
   * and 4 of the 4 GeneralException failures that scenario has ever had.
   *
   * A CEILING OF FIVE, ONE BIND, and both numbers are load-bearing.
   *
   * Five is the deck's size when this scenario starts, so `before` is read
   * accurately — the ceiling only binds on a deck that has grown past it — and
   * the read after the insert is pinned back to exactly `before`. That is what
   * makes `landed: 0` rather than merely short, which is the archive's actual
   * signature. A ceiling of seven produces `landed: 2` and this case would
   * assert nothing.
   *
   * One bind, because the repair has to be the thing that clears it. With two
   * the re-read is capped as well and the count stays short, which is the state
   * the test above pins.
   *
   * AND THE REPRODUCTION GUARD DOES NOT READ A TRACE THE FIX ITSELF EMITS.
   * `count came back short` is written by `settledSlideCount`, so a version of
   * this test that guarded on it failed under the mutant with "nothing was
   * reproduced" — the guard tripping, never the assertion. The allowance is
   * spent by the raw read too, so counting it down proves the lag happened
   * under either.
   */
  it("reports the slides a late deck really took, not the zero it first read", () => {
    installHost([makeSlide("s1")]);
    faults.slideCountCeiling = 5;
    faults.slideCountCeilingBinds = 1;
    _setCountSettleDelayForTest(120);
    setTracing(true);
    return runSelfTest("probe", NAME).then(() => {
      setTracing(false);
      expect(faults.slideCountCeilingBinds, "the ceiling never bound, so no lag was reproduced").toBe(0);
      const handed = traceLog().entries.filter((e) => e.message === "handed the host a generated deck");
      expect(handed.length, "no deck was inserted, so there is nothing to assert about").toBeGreaterThan(0);
      // THE ASSERTION. A deck that landed is never reported as having landed
      // nothing.
      const zeros = handed.filter((e) => (e.data as { landed?: number }).landed === 0);
      expect(
        zeros.length,
        `an insert reported landed: 0 for slides that landed — ${JSON.stringify(handed.map((e) => e.data))}`,
      ).toBe(0);
    });
  });

  it("skips, naming the shortfall, when the count never settles", () => {
    installHost([makeSlide("s1")]);
    // The same ceiling, but it never lifts — so the settle cannot rescue it and
    // the deck genuinely cannot be judged.
    faults.slideCountCeiling = 7;
    faults.slideCountCeilingBinds = 999;
    _setCountSettleDelayForTest(120);
    return runSelfTest("probe", NAME).then((rs) => {
      const r = byName(rs)[NAME];
      // A SKIP and not a FAILURE. Nothing about deduplication was tested, so
      // blaming the dedup would be the false verdict this exists to stop.
      expect(r.skipped, `a host that never settled produced a hard verdict: ${r.detail}`).toBe(true);
      expect(r.detail).toMatch(/slides landed even after a settle/);
      // And it says what it saw, so the next reader does not have to re-derive
      // it from a trace.
      expect(r.detail, `the skip did not name the shortfall: ${r.detail}`).toMatch(/of 4 slides landed/);
    });
  });

  it("leaves an ordinary run alone", () => {
    installHost([makeSlide("s1")]);
    return runSelfTest("probe", NAME).then((rs) => {
      const r = byName(rs)[NAME];
      expect(r.skipped, `a healthy host was skipped: ${r.detail}`).toBeFalsy();
      expect(r.detail).toContain("4 slides inserted");
    });
  });
});

/**
 * THE QUESTION NOBODY COULD ANSWER, and the scenario that now asks it.
 *
 * `Shape.rotation` has never been written on a real PowerPoint in 333 rounds,
 * because the battery draws only `clustered` and its one line node is the
 * horizontal baseline. So `addSegment`'s rotated-rectangle branch — the path
 * EVERY diagonal in the product takes — has zero real-host executions, and the
 * renderer's assumption that `left`/`top`/`width` mean the box BEFORE rotation
 * is untested against the host it is an assumption about.
 *
 * If the host means the post-rotation bounding box instead, every diagonal in
 * every line, scatter, radar and violin chart is drawn wrong. These tests are
 * what make the scenario capable of saying so.
 */
describe("a chart of rotated shapes", () => {
  const verdict = async () => {
    const results = await runSelfTest("probe");
    const r = results.find((x) => x.name === "a chart of rotated shapes");
    expect(r, "the scenario did not run at all").toBeTruthy();
    return r!;
  };

  it("passes on a host that means the box before rotation", async () => {
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    const r = await verdict();
    expect(r.ok, `expected a pass on a well-behaved host: ${r.detail}`).toBe(true);
    // Non-vacuous: it must have actually compared segments, not skipped.
    expect(r.skipped ?? false, `skipped instead of measuring: ${r.detail}`).toBe(false);
    expect(r.detail, "did not report what it measured").toMatch(/UNROTATED box/);
  });

  it("reports the DRAW on a host that will not open a group", async () => {
    /**
     * WHAT THIS SCENARIO CAN ACTUALLY ESTABLISH ON THE HOST WE RUN AGAINST.
     *
     * Reading a chart's parts means reading a group's CHILDREN, and this host
     * refuses both routes into one — `threw` on `group-reports-its-children`
     * and on `group-children-via-getcount` in 29 of the last 30 rounds. So
     * where the segments landed is unmeasurable here: rounds 334 and 335
     * skipped because the lookup was at slide level, and 339 skipped with the
     * descent in place, for this reason instead.
     *
     * The DRAW is still worth having, and nothing else in the battery does it.
     * `Shape.rotation` had never been written on a real host in 333 rounds, so
     * `addSegment`'s rotated branch — the path every diagonal in the product
     * takes — had never executed outside the fake. Getting the chart in without
     * an error is a real answer about a real path; a skip threw it away and
     * left a scenario that could only ever report nothing.
     */
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    faults.groupHidesChildren = true;
    faults.refuseGroupRead = true;
    try {
      const results = await runSelfTest("probe", "a chart of rotated shapes");
      const r = results.find((x) => x.name === "a chart of rotated shapes");
      expect(r, "the scenario did not run at all").toBeTruthy();
      expect(r!.skipped ?? false, "threw the draw away as a skip").toBe(false);
      expect(r!.ok, `reported the draw as a failure: ${r!.detail}`).toBe(true);
      expect(r!.detail, "did not say the measurement was unavailable").toMatch(/unmeasurable|will not report/i);
      /**
       * AND THE NUMBER IT REPORTS IS A REAL ONE.
       *
       * Round 341 said "drew 0 rotated segment(s) without error", because the
       * count came from `partIds` — which is empty on the very host the rest of
       * the sentence is about. A pass reporting zero of the thing it exists to
       * draw is barely better than the skip it replaced, so the count comes
       * from the SCENE now and this asserts it is not zero.
       */
      const drew = /drew a (\d+)-segment/.exec(String(r!.detail));
      expect(drew, `did not report a segment count: ${r!.detail}`).toBeTruthy();
      expect(Number(drew![1]), "reported drawing nothing and called it a pass").toBeGreaterThan(0);
    } finally {
      faults.refuseGroupRead = false;
      faults.groupHidesChildren = false;
    }
  });

  it("still concludes when the slide lists only the group, as a real one does", async () => {
    /**
     * THE REASON THIS SCENARIO REPORTED NOTHING TWICE.
     *
     * Rounds 334 and 335 both skipped it — "the host would not report the line
     * segments' geometry" — while the host had reported everything it was asked
     * for. `insertSceneIntoSlide` GROUPS a chart, so a real slide answers with
     * one shape called `PowerChart` per chart and no parts at all; both rounds
     * show exactly six of those and not one segment. The lookup read slide
     * level, and the message blamed the host for our own mistake.
     *
     * It passed here throughout, because this fake leaves grouped children in
     * the slide's collection — the fake being the optimistic one, which is the
     * direction that misleads. `faults.groupHidesChildren` is that divergence
     * modelled, and without the descent into groups this test skips.
     */
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    faults.groupHidesChildren = true;
    try {
      // THIS SCENARIO ALONE, because the fault is blunter than the host it
      // stands for: a real slide hides its group members from the collection
      // but still resolves one by id, and this hides them from everything. Run
      // over the whole battery it starves the readback scenarios and the run
      // gives up three scenarios before reaching this one. Narrowing the run is
      // honest here — the claim is about this scenario's lookup, not about the
      // battery surviving a host that behaves this way.
      const results = await runSelfTest("probe", "a chart of rotated shapes");
      const r = results.find((x) => x.name === "a chart of rotated shapes");
      expect(r, "the scenario did not run at all").toBeTruthy();
      expect(r!.skipped ?? false, `skipped on a slide that hides its group members: ${r!.detail}`).toBe(false);
      expect(r!.ok, `did not conclude: ${r!.detail}`).toBe(true);
      expect(r!.detail, "did not report what it measured").toMatch(/UNROTATED box/);
    } finally {
      faults.groupHidesChildren = false;
    }
  });

  it("FAILS on a host that means the box after rotation", async () => {
    // The whole point. A scenario that cannot report the bad answer is not
    // evidence, and this is the answer that would invalidate every diagonal the
    // product draws.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    faults.reportRotatedBounds = true;
    try {
      const r = await verdict();
      expect(r.ok, `did not notice a host placing rotated shapes by their bounding box: ${r.detail}`).toBe(false);
      expect(r.skipped ?? false, "called it a skip rather than a finding").toBe(false);
      expect(r.detail, "failed without naming what it found").toMatch(/POST-rotation/);
    } finally {
      faults.reportRotatedBounds = false;
    }
  });

  // CLEANUP IS GUARDED, but not from here. My own version of that test looked
  // for the measurement shapes by name across every slide and found none either
  // way — vacuous, and the cleanup mutant walked straight through it. The real
  // guard is `leaves no two charts stacked on one slide` above: it fails when
  // the cleanup is removed, because it asserts the CONSEQUENCE (a chart left on
  // top of another) rather than the mechanism. Verified by mutating
  // `await clean()` away and watching that test, not this one, go red.
});

describe("a slide count nobody took", () => {
  /**
   * `slideCount()` RETURNS UNDEFINED ON A HOST THAT WILL NOT ANSWER, on
   * purpose — its docstring says "No safe default here, deliberately … Not
   * knowing has to stay distinguishable from knowing." Three scenarios
   * subtracted the two numbers anyway, and unknown arrived dressed as a
   * measurement in two opposite costumes.
   */
  it("is not a growth of NaN, and not a growth of zero either", () => {
    expect(deckGrowth(3, 7)).toBe(4);
    expect(deckGrowth(7, 7)).toBe(0);
    expect(deckGrowth(7, 3), "a deck that shrank is still a measurement").toBe(-4);

    /**
     * ROUND 360, THE LOUD COSTUME. `insertTwice` filed **"deck grew by NaN
     * (want 4); 0 of 4 charts re-editable"** as a hard FAILED verdict. That
     * sentence reads as measured data loss and sends a maintainer after the
     * insert path; nothing had been measured at all.
     */
    expect(deckGrowth(undefined, 7), "an unread `before` produced a number").toBeNull();
    expect(deckGrowth(3, undefined), "an unread `after` produced a number").toBeNull();
    expect(deckGrowth(undefined, undefined)).toBeNull();
    // NaN is the shape the bug actually took — a subtraction that already
    // happened — so it is pinned separately from undefined.
    expect(deckGrowth(NaN, 7), "NaN was passed through as a growth").toBeNull();
    expect(deckGrowth(3, NaN)).toBeNull();
  });

  it("fires the guards it used to silence", () => {
    /**
     * THE QUIET COSTUME, and the worse one. `after !== before` is FALSE when
     * both are undefined, so an unreadable deck read as "it did not grow" —
     * precisely the condition those two lines exist to detect — and the
     * scenario passed on no evidence.
     *
     * One of them guards that Stop is non-destructive, which nothing else in
     * this suite checks. A guard that cannot fire on no evidence is not a
     * guard, and this is the same mistake `blindSkip` was written for one
     * measurement over: the scan version was swept and the COUNT version was
     * not, because at the call site they look nothing alike.
     *
     * Asserted as source because reaching those two lines needs a host that
     * answers a scan and refuses a count, and the fake has no fault for it.
     *
     * AND EVERY SITE ANSWERS THE SAME WAY. The first repair made two of them
     * report a PROBLEM instead — which is a FAILED verdict, so unknown came
     * back as a measurement in a fresh costume, against the rule those very
     * scenarios state: "a blind scan is reported as SKIPPED, not as a pass and
     * not as a failure: nobody has evidence either way, and a verdict that
     * cannot be attributed costs a whole real-host round to chase."
     */
    const src = readFileSync("src/taskpane/selftest.ts", "utf8");
    const unguarded = src.match(/after !== before && `the deck grew by/g) ?? [];
    expect(unguarded, "a scenario went back to reading an unread count as no growth").toHaveLength(0);
    // Each of the three call sites goes through the helper.
    expect((src.match(/deckGrowth\(before, after\)/g) ?? []).length).toBeGreaterThanOrEqual(3);
    // And the loud one skips rather than failing.
    expect(src, "an unread count no longer skips the scenario").toMatch(
      /const grew = deckGrowth\(before, after\);\s*\n\s*if \(grew === null\) return countlessSkip\(\);/,
    );

    /**
     * THE FOURTH SITE, and the reason this assertion counts rather than spot-
     * checks. The first sweep of this defect repaired three scenarios and left
     * `duplicateSlot` — which owns BOTH of round 360's failing sentences, and
     * whose opening guard `afterInsert - before < wanted` sailed through on an
     * unread deck because `null < wanted` is false in JavaScript.
     *
     * That is the "fix written, one call site left behind" shape this repo has
     * now shipped five times, committed on the same day as the fix it belongs
     * to. Counting every site is what a spot-check cannot do.
     */
    const subtractions = src.match(/\b(?:after|afterInsert|settled) - before\b/g) ?? [];
    expect(subtractions, `a raw slide-count subtraction survived: ${subtractions.join(", ")}`).toHaveLength(1);
    // The one survivor is inside `deckGrowth` itself, which is where the
    // subtraction is supposed to live.
    expect(src).toMatch(/Number\.isFinite\(after\)\s*\n?\s*\?\s*after - before/);
    expect((src.match(/deckGrowth\(/g) ?? []).length, "not every site goes through the helper").toBeGreaterThanOrEqual(
      6,
    );

    // ONE ANSWER TO UNKNOWN, everywhere. Every scenario that reads a growth
    // reaches `countlessSkip` when the count is missing — none of them turns it
    // into a problem string, which would file a FAILED verdict on no evidence.
    const skips = src.match(/if \(\w+ === null\) return countlessSkip\(\);/g) ?? [];
    expect(skips.length, "a site answers an unread count some other way").toBeGreaterThanOrEqual(4);
    expect(src, "an unread count is being reported as a problem — that is a FAILED verdict, not a skip").not.toMatch(
      /deckGrowth\([^)]*\) === null\s*\n?\s*\?/,
    );
  });

  it("reports a lost count as lost sight, not as a missing feature", () => {
    /**
     * `blind` exists to keep those two apart, and its own docstring says what
     * omitting it costs: the run reports "N skipped (host cannot run them)",
     * "a total loss of deck visibility reported as a feature gap, which is the
     * sort of line someone files and moves on from".
     *
     * `countlessSkip` shipped without the flag. An unread COUNT and an unread
     * SCAN are the same loss of visibility, and a reader has to see them as the
     * same thing — so it carries `blind` exactly as `blindSkip` does.
     */
    const src = readFileSync("src/taskpane/selftest.ts", "utf8");
    const body = src.slice(src.indexOf("const countlessSkip"), src.indexOf("type Scenario"));
    expect(body.length, "could not find countlessSkip").toBeGreaterThan(50);
    expect(body, "a lost count is reported as a feature gap, not as lost visibility").toMatch(/blind:\s*true/);
    expect(body, "a countless skip stopped being a skip").toMatch(/skipped:\s*true/);
    expect(body, "a countless skip started claiming success").toMatch(/ok:\s*false/);
  });
});
