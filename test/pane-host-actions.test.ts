// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import type { ChartConfig } from "../src/core/types";
import { buildChart } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";

/**
 * Pane ↔ host command handlers. `pane-state.test.ts` boots app.ts under jsdom
 * with no `Office` global, so it exercises everything EXCEPT the buttons that
 * talk to PowerPoint — Insert, Insert-new, Load-selection, Same-scale. Those
 * handlers (`doInsert`, `doLoadSelection`, `doSameScale`) were the pane's
 * largest untested surface: the branch that chooses a selected placeholder's
 * bounds over the tiled offset, the edit-in-place path, the union-extent maths
 * behind Same scale, and the "not an SSF chart" / "needs two charts" guards.
 *
 * The renderer primitives those handlers call (`insertSceneIntoSlide` et al.)
 * are covered against a fake host in `office-render.test.ts`; here the module is
 * mocked to spies so the test asserts the pane's ORCHESTRATION — which handler
 * fires, with what config — rather than re-testing the shape emitter.
 */

/** Shared mailbox the mocked renderer writes to; reset before each boot. */
/**
 * What `isPowerPointHost()` answers, and WHEN.
 *
 * Hoisted beside `host` because the `vi.mock` factory closes over it. Default
 * `always: true` keeps every existing test unchanged; `lateOffice` reproduces
 * the real ordering — false while app.ts is evaluating at module scope, true by
 * the time `Office.onReady` fires.
 */
const hostReadiness = vi.hoisted(() => {
  let mode: "always" | "late" = "always";
  let ready = false;
  return {
    answer: () => (mode === "always" ? true : ready),
    reset: () => {
      mode = "always";
      ready = false;
    },
    lateOffice: () => {
      mode = "late";
      ready = false;
    },
    officeArrives: () => {
      ready = true;
    },
  };
});

/** Where the mocked `traceEnvironment` sends its line. Pointed at the real `trace`. */
const traceRef = vi.hoisted(() => ({
  fn: (_scope: string, _message: string, _data?: Record<string, unknown>) => {},
}));

const host = vi.hoisted(() => ({
  selectionBounds: null as null | { left: number; top: number; width: number; height: number },
  /** What is already on the slide the next insert would draw onto. */
  /**
   * What `getSlideShapeBounds` answers — `null` for a host that would not read
   * the slide at all, which is a different fact from an empty slide.
   */
  slideShapes: [] as { left: number; top: number; width: number; height: number }[] | null,
  deckCharts: [] as { configJson: string; target: unknown }[],
  /** Slides the deck scan could not read — the signal Same Scale now gates on. */
  deckScanUnread: 0,
  selectionCharts: [] as { configJson: string; target: unknown }[],
  loadSelectionResult: null as null | { configJson: string; target: unknown },
  // When set, insertSceneIntoSlide awaits this before resolving — lets a test
  // observe the pane's mid-flight state (buttons disabled) before the action ends.
  gate: null as null | Promise<void>,
  // insertSceneIntoSlide throws this once, if set — drives the guard's catch path.
  failInsertOnce: false,
  // When set, updateChartInSlide awaits this — holds an in-place update open
  // so a test can act while it is still running.
  updateGate: null as null | Promise<void>,
  insertGate: null as null | Promise<void>,
  // The selection-change listener app.ts registers via addHandlerAsync, captured
  // so a test can fire it the way PowerPoint would.
  selectionListener: null as null | (() => unknown),
  /** The budget each loadChartFromSelection call was given, in order. */
  selectionBudgets: [] as (number | undefined)[],
  /** When set, loadChartFromSelection awaits this — holds a selection read open. */
  selectionGate: null as null | Promise<void>,
  /** What `readDeckStyle` answers — the style this deck carries, or none. */
  deckStyle: null as null | Record<string, unknown>,
  /** When true, the deck-style READ fails — not the same as a deck carrying none. */
  deckStyleUnreadable: false,
  /** Every style handed to `writeDeckStyle`, in order. */
  deckStyleWrites: [] as unknown[],
  /** False for a host below PowerPointApi 1.7, which cannot store one at all. */
  deckStyleWritable: true,
  agendaSlides: [] as unknown[][],
  demoRuns: 0,
  demoDeckCalls: [] as unknown[][],
  /** The runOpts each insertDemoDeck call was given, in order. */
  demoDeckOpts: [] as ({ reconcile?: boolean } | undefined)[],
  /** insertDemoDeck awaits this on the Nth call — lets a test act mid-run. */
  demoDeckGateOn: 0,
  demoDeckGate: null as null | Promise<void>,
  demoDeckPageFailures: new Set<number>(),
  demoDeckStatusOverride: null as null | "rendered" | "failed" | "skipped",
  /** What canInsertPicture() reports — false models a host below PowerPointApi 1.8. */
  canPicture: true,
  /**
   * What updateChartInSlide resolves to. Defaults to undefined because the
   * existing "target is gone" test depends on that; the explode tests opt in to
   * a live EditTarget, which is what a real successful update hands back.
   */
  updateResult: undefined as undefined | { slideId: string; shapeId: string; left: number; top: number; lost?: string },
  /** What a fresh insert hands back — an EditTarget now, so a recovery stays editable. */
  insertResult: null as null | { slideId: string; shapeId: string; left: number; top: number; lost?: string },
  /** Whether the host advertises insertSlidesFromBase64 — off by default. */
  canInsertFile: false,
  slideHoldsOnlyChart: false,
  /** Whether the renderer says this chart is too dense to draw as shapes. */
  autoPicture: false,
  /**
   * Marks this host cannot draw at all — a sub-1.10 PowerPoint, where a wedge
   * and an arrowhead both draw NOTHING. A different question from `autoPicture`
   * and a different host: that one is the web, which is too slow for a dense
   * chart; this is desktop and every volume-licensed build, which is fast and
   * simply has no `Shape.rotation`.
   */
  cannotDraw: { what: [] as string[], nodes: 0 },
  updateChartThrows: false,
  /**
   * Which items of a multi-chart update stall, keyed by index, and the stray
   * shape ids each one left on its slide.
   *
   * The real `updateChartsInSlides` deletes a chart's old shapes before it
   * redraws them, so a stall leaves a partial chart behind and reports it via
   * `onFailed` with the debris attached to the error. Modelling that is the
   * only way a test can see whether the caller sweeps.
   */
  updateChartsStalls: new Map<number, string[]>(),
  /**
   * How many charts `updateChartsInSlides` drops WITHOUT reporting them.
   *
   * The failure the renderer's early filters produce and no callback carries:
   * a chart whose slide or shape the host would not resolve never reaches
   * `onFailed`, so the only way a caller can notice is by counting what came
   * back. A real host dropped five of six this way and the pane said
   * "Same scale applied to 6 charts".
   */
  updateChartsDrops: 0,
  /** Whether the user has pressed Stop — the flag the render loops read. */
  stopRequested: false,
  /** What slideSize() reports — 16:9 unless a test says otherwise. */
  slideSize: { width: 960, height: 540, source: "pageSetup" as const },
  demoReconcile: undefined as unknown,
  /** What `replaceSlideWithDeck` answers: "failed" | "swapped" | "duplicated". */
  swapOutcome: "failed" as "failed" | "swapped" | "duplicated",
  /** How many slides the one-shot insert reports landing; null = all of them. */
  insertFileLands: null as null | number,
  insertFileError: null as null | Error,
  buildFileError: null as null | Error,
  slideCount: 1,
  /**
   * Number of slide-count reads that succeed before the rest throw — a host
   * that answers, then wedges. `null` means it always answers.
   */
  slideCountThrowsAfter: null as null | number,
  slideCountCalls: 0,
  /**
   * Make the round's very first statement throw a NATIVE Error.
   *
   * Every other fault in this file models the HOST refusing — politely, at a
   * call the trace already names. This one models OUR OWN code failing, which
   * is the case the record could not place: `resetSyncCount` runs at the top of
   * the round with nothing wrapping it but the outer handler, so a throw here
   * arrives there and nowhere else.
   */
  syncResetThrows: false,
  /** What a deck scan reports holding, when the caller asked for the inventory. */
  deckInventory: [] as { slideId: string; index: number; shapes: { id: string; name?: string }[] }[],
  /**
   * The deck's slide ids. `undefined` models a host that will not list them —
   * which is the case the round's id diff has to survive, not an edge case: it
   * is what every blind deck scan in this project's history looked like.
   */
  deckSlideIds: undefined as undefined | string[],
  deckSlideIdsBeforeFails: false,
  /**
   * What `addSlideForChart` answers when the pane takes up the slow-path offer.
   *
   * `null` is not an edge case here either: this host DROPS `slides.add()`, so
   * "accepted, and no slide appeared" is a real outcome and the pane's fallback
   * to inserting where the user asked has to be reachable from a test.
   */
  addSlideResult: "own-slide-1" as string | null,
  addSlideCalls: 0,
  /** Whether the host will draw a slide — PowerPointApi 1.8, absent on plenty of hosts. */
  canRaster: true,
  /**
   * The deck scan throws rather than answering short.
   *
   * A host that will not answer and a host that raises are different things,
   * and only the second one can take a round's verdicts down with it. The
   * first version of the guard for that used the first condition and passed
   * with the protection removed.
   */
  /** Which platform the pane believes it is on — drives every web-only guard. */
  platform: "OfficeOnline" as string,
  deckScanThrows: false,
  /**
   * THE SCAN THAT NEVER ANSWERS — what a host CRASH looks like, and a different
   * fault from the one above.
   *
   * `collectDeckEvidence` catches, so a scan that RAISES was always survivable.
   * A scan that never settles is not: the `await` in front of it never returns,
   * so everything after it — assembling the run log, enabling its button —
   * never runs at all. That is how 33 of 49 crash records came to hold all
   * fourteen verdicts while no round was ever filed from any of them.
   *
   * COUNTED, not a plain flag, because `withInventory` scans happen during the
   * scenarios too — a blunt flag hangs the first one and stops the round before
   * it has any verdicts to lose, which tests the opposite of what is meant. The
   * test discovers the number a clean round makes rather than hard-coding it,
   * so adding a scenario that scans cannot quietly disarm this.
   */
  deckInventoryScans: 0,
  /** Hang on the Nth `withInventory` scan; null never hangs. */
  deckScanHangsOnScan: null as number | null,
  deckSlideIdCalls: 0,
  /** Slides the round is modelled as having added, seen by the second id read. */
  roundAddsSlides: ["probe-a", "probe-b"] as string[],
  /**
   * How many charts had been drawn each time the shape selection was dropped.
   *
   * The ORDER is the whole assertion — see office-js#2775. Dropping the
   * selection after the draw would be no protection at all, and a mock that only
   * recorded "it happened" could not tell the two apart.
   */
  selectionDropped: [] as number[],
  /** Slide ids `deleteSlideById` was asked about, in order. */
  deletedSlides: [] as string[],
  /** Slide ids the host will refuse to delete — a refusal is not the same as a delete. */
  refuseSlideDelete: [] as string[],
  reconcileOutcome: undefined as unknown,
  calls: {
    insertScene: [] as { tagData?: string; left?: number; top?: number; slideId?: string; pictureBase64?: string }[],
    updateChart: [] as { target: unknown; opts: { tagData?: string; pictureBase64?: string } }[],
    updateCharts: [] as { scene: unknown; target: unknown; opts?: { tagData?: string } }[][],
    insertFile: [] as { b64: string; expected: number }[],
    deselected: [] as string[][],
    /** Sweeps of the litter a stalled redraw left behind, per recovery. */
    swept: [] as { slideId: string; ids: string[] }[],
  },
}));

vi.mock("../src/render/powerpoint", () => ({
  // FLIPPABLE, to reproduce the ordering the real host has. app.ts asks this at
  // MODULE SCOPE, long before `Office.context` exists, and asked it only once.
  isPowerPointHost: () => hostReadiness.answer(),
  // The round-start line evaluates this as an argument, so a mock without it
  // throws and the line vanishes — which is how eight tests in this file went
  // red at once when it was added. Kept minimal on purpose: the pane must not
  // depend on any field of it.
  roundEnvironment: () => ({ requirementSets: [] }),
  // Same reason: the battery stamps every scenario with the host-friction
  // delta, so it is on the round's critical path now. (`deckSlideIds` is
  // already mocked further down, with the round's own growth modelled.)
  // FOUR NAMES HERE, EIGHT IN THE REAL THING. This listed the counters that
  // existed when it was written, so a caller reading any of the four added
  // since got `undefined` from the double and a number from production — a
  // fake that cannot fail, which this repo has been bitten by before.
  //
  // A `vi.mock` factory is hoisted and cannot read the real module, so the list
  // stays literal. `the pane's friction double carries every counter` below
  // compares it against `vi.importActual` and goes red the day a ninth is added.
  hostFrictionCounts: () => ({
    errors: 0,
    idRefusals: 0,
    generalExceptions: 0,
    emptyReReads: 0,
    shortReReads: 0,
    unmatchedReReads: 0,
    reReadsRepaired: 0,
  }),
  canInsertPicture: vi.fn(() => host.canPicture),
  getSelectionBounds: vi.fn(async () => host.selectionBounds),
  dropShapeSelection: vi.fn(async () => {
    host.selectionDropped.push(host.calls.insertScene.length);
    return true;
  }),
  getSlideShapeBounds: vi.fn(async () => host.slideShapes),
  addSlideForChart: vi.fn(async () => {
    host.addSlideCalls++;
    return host.addSlideResult;
  }),
  insertSceneIntoSlide: vi.fn(
    async (
      scene: { nodes: unknown[] },
      opts: { tagData?: string; left?: number; top?: number },
      onPhase?: (phase: string, detail?: string) => void,
    ) => {
      // Report phases the way the real renderer does. The mock used to ignore
      // `onPhase` entirely, and that silence is what hid a live bug for good:
      // with no phase notes the pane's note stayed exactly "Working…", which is
      // the one string the old `guard` recognised as "nothing was said", so
      // every test saw the "Done." a real PowerPoint never printed.
      onPhase?.("context");
      onPhase?.("queue", `${scene.nodes.length} nodes`);
      onPhase?.("commit", `${scene.nodes.length} of ${scene.nodes.length} shapes`);
      onPhase?.("group");
      if (host.gate) await host.gate;
      if (host.failInsertOnce) {
        host.failInsertOnce = false;
        throw new Error("host refused the insert");
      }
      host.calls.insertScene.push(opts);
      // The sibling of `updateGate`, for holding an INSERT open. Needed to show
      // two guarded actions overlapping in either order: the update path had a
      // gate and the insert path did not, so a test could hold the outer action
      // but never the nested one.
      if (host.insertGate) await host.insertGate;
      // The last thing a successful insert ever says — and it is still "busy".
      onPhase?.("done");
      return host.insertResult;
    },
  ),
  updateChartInSlide: vi.fn(
    async (_scene: unknown, target: unknown, opts: { tagData?: string; pictureBase64?: string }) => {
      host.calls.updateChart.push({ target, opts });
      if (host.updateGate) await host.updateGate;
      // A stop taken while this was in flight surfaces the way the renderer
      // surfaces it: a marked error carrying whatever the halted redraw had
      // already committed. Stopping mid-batch leaves the same debris a stall
      // does, so the wreckage is not optional.
      if (host.stopRequested) {
        const err = new Error("Stopped.") as Error & Record<string, unknown>;
        err.__powerchartStopped = true;
        err.__powerchartWreckage = { slideId: "s1", at: { left: 40, top: 50 }, strayIds: ["stray-1"] };
        throw err;
      }
      // Models the live-canvas stall: the in-place redraw refuses, a picture
      // update (which draws one shape) does not.
      if (host.updateChartThrows && !opts.pictureBase64) {
        // Carrying wreckage, because the real one always does: an update deletes
        // the old chart and COMMITS that before it draws a single new shape, so
        // a stall mid-redraw is never a no-op. A double that threw a bare error
        // modelled a rollback the host does not have, and every fallback below
        // was written against that fiction.
        const err = new Error("did not respond while drawing shapes 1-10") as Error & Record<string, unknown>;
        err.__powerchartWreckage = { slideId: "s1", at: { left: 40, top: 50 }, strayIds: ["stray-1", "stray-2"] };
        throw err;
      }
      return host.updateResult;
    },
  ),
  wreckageOf: (err: unknown) =>
    err && typeof err === "object" && Object.prototype.hasOwnProperty.call(err, "__powerchartWreckage")
      ? (err as Record<string, unknown>).__powerchartWreckage
      : undefined,
  deleteShapesById: vi.fn(async (slideId: string, ids: string[]) => {
    host.calls.swept.push({ slideId, ids });
    return ids.length;
  }),
  // The cooperative stop. Modelled with the real semantics — a flag the render
  // loops read — so a test can press Stop mid-action and see what the pane does
  // with it, rather than only that the button exists.
  requestStop: vi.fn(() => {
    host.stopRequested = true;
  }),
  resetStop: vi.fn(() => {
    host.stopRequested = false;
  }),
  isStopRequested: vi.fn(() => host.stopRequested),
  // The destination deck's slide size, which the deck builder needs so the
  // generated file declares the size it is being inserted into.
  slideSize: vi.fn(async () => host.slideSize),
  isStopped: (err: unknown) =>
    !!err && typeof err === "object" && Object.prototype.hasOwnProperty.call(err, "__powerchartStopped"),
  updateChartsInSlides: vi.fn(
    async (
      items: { scene: unknown; target: { slideId?: string }; opts?: { tagData?: string } }[],
      onFailed?: (item: unknown, err: unknown) => void,
    ) => {
      host.calls.updateCharts.push(items);
      // Report the stalls the test asked for, exactly as the renderer does:
      // the error carries the wreckage, because that is the only channel the
      // caller has for finding out what was left on the slide.
      for (const [i, strayIds] of host.updateChartsStalls) {
        const item = items[i];
        if (!item) continue;
        const err = new Error("host stalled") as Error & Record<string, unknown>;
        err.__powerchartWreckage = {
          slideId: item.target?.slideId ?? `s${i}`,
          at: { left: 0, top: 0 },
          strayIds,
        };
        onFailed?.(item, err);
      }
      // One target per chart the host RESOLVED — the real contract, and the
      // only channel a caller has for "this chart was dropped without a word".
      // The renderer's `live`/`alive` filters drop a chart whose slide or shape
      // the host declines to answer for, and they are early returns rather than
      // throws, so `onFailed` never fires for them. A stalled chart is still in
      // this array (it comes back carrying its OLD target), which is why the two
      // signals do not double-count. `updateChartsDrops` is how a test arms the
      // silent case.
      return items.slice(host.updateChartsDrops).map((it) => it.target);
    },
  ),
  listChartsInDeck: vi.fn(async (opts: { withInventory?: boolean } = {}) => {
    if (opts.withInventory) {
      host.deckInventoryScans++;
      if (host.deckScanHangsOnScan !== null && host.deckInventoryScans >= host.deckScanHangsOnScan)
        return new Promise(() => {});
    }
    if (host.deckScanThrows) throw new Error("the host refused to describe the deck");
    return {
      charts: host.deckCharts,
      unread: host.deckScanUnread,
      short: 0,
      tagsUnread: 0,
      slides: host.slideCount,
      // Only when asked, exactly like the real one — a mock that always returns
      // the inventory would let a caller that forgot to ask still find it.
      ...(opts.withInventory ? { inventory: host.deckInventory } : {}),
    };
  }),
  // The deck GROWS across a round, because that is the only thing the id diff
  // is measuring. A double that answered the same list twice would make the
  // diff empty and every assertion about it vacuous.
  deckSlideIds: vi.fn(async () => {
    const first = host.deckSlideIdCalls++ === 0;
    // ONLY THE BEFORE-READ FAILS, which is the state that matters and the one a
    // double answering `undefined` to BOTH calls cannot express. With both
    // unreadable the old code and the fixed code agree on an empty list, so a
    // test built that way passes with the fix deleted — it did, before this
    // existed.
    if (host.deckSlideIdsBeforeFails && first) return undefined;
    if (!host.deckSlideIds) return undefined;
    return first ? host.deckSlideIds : [...host.deckSlideIds, ...host.roundAddsSlides];
  }),
  slideShots: vi.fn(async (ids: string[], opts: { max?: number } = {}) =>
    ids.map((slideId, i) => (i < (opts.max ?? 12) && host.canRaster ? { slideId, png: "UE5H" } : { slideId })),
  ),
  scanIsComplete: (s: { unread: number; short: number; tagsUnread: number }) =>
    s.unread === 0 && s.short === 0 && s.tagsUnread === 0,
  scanGap: (s: { unread: number; slides: number }) =>
    s.unread ? `${s.unread} of ${s.slides} slide(s) would not answer` : "",
  listChartsInSelection: vi.fn(async () => host.selectionCharts),
  loadChartFromSelection: vi.fn(async (budgetMs?: number) => {
    host.selectionBudgets.push(budgetMs);
    if (host.selectionGate) await host.selectionGate;
    return host.loadSelectionResult;
  }),
  insertAgendaSlides: vi.fn(async (scenes: unknown[][]) => {
    host.agendaSlides.push(scenes);
  }),
  insertDemoDeck: vi.fn(
    async (
      items: { title?: string; scene: { nodes: unknown[] } }[],
      _onProgress?: unknown,
      runOpts?: { reconcile?: boolean },
    ) => {
      host.demoRuns++;
      const call = host.demoRuns;
      host.demoDeckCalls.push(items);
      host.demoDeckOpts.push(runOpts);
      if (host.demoDeckGate && call === host.demoDeckGateOn) await host.demoDeckGate;
      if (host.demoDeckPageFailures.has(call)) throw new Error(`page ${call} refused`);
      // Fabricated report — each item counted as failed so the pane's results
      // path has rows to render. Callers that need a specific status set
      // demoDeckStatusOverride.
      const status = host.demoDeckStatusOverride ?? "failed";
      const results = items.map(() => ({ created: 5, status, ms: 100 }));
      return {
        // The run's identity, as the real renderer reports it — the join key
        // between the downloaded log and the .pptx it produced. Fabricated
        // per call so a test can tell two runs apart.
        run: `fake-run-${call}`,
        results,
        slidesAdded: items.length,
        addsIssued: items.length,
        blankSlides: [],
        blankItems: [],
        blanksRead: true,
        reconcile: host.demoReconcile,
        totalMs: items.length * 100,
      };
    },
  ),
  loadThemePalette: vi.fn(async () => null),
  // The deck's own style file. `null` is the ordinary answer — a deck nobody
  // has branded — and the pane must be usable before this resolves, which is
  // what the startup read being fire-and-forget is for.
  readDeckStyle: vi.fn(async () => host.deckStyle),
  // The same read, saying why it came back empty. The button awaits THIS one,
  // because "the read failed" and "the deck carries none" send the person at
  // the pane to opposite conclusions about their own document.
  // The pane warms the custom-XML surface before reading it — a call whose
  // failure is the point. Absent from this mock, app.ts throws at boot and every
  // test in this file dies, which is how 110 of them went red at once.
  // Found by `pane-mock-covers-the-renderer` on its first run: app.ts:2864 calls
  // this and the mock had no such key, so any pane test reaching that line would
  // have taken all 110 down with it. A landmine, not yet a crater.
  enrichSnapshots: vi.fn(async () => {}),
  warmCustomXmlSurface: vi.fn(async () => {}),
  readDeckStyleWithReason: vi.fn(async () => ({
    style: host.deckStyleUnreadable ? null : host.deckStyle,
    unreadable: host.deckStyleUnreadable,
  })),
  writeDeckStyle: vi.fn(async (style: unknown) => {
    host.deckStyleWrites.push(style);
    return host.deckStyleWritable;
  }),
  onLateSync: vi.fn(),
  errorText: (e: unknown) => String(e),
  // The one-shot deck path. Off by default so the existing cases keep
  // exercising the shape-by-shape renderer they were written for.
  canInsertSlidesFromBase64: vi.fn(() => host.canInsertFile),
  // EMITS, rather than doing nothing. A no-op here is why 69 rounds shipped with
  // `traceEnvironment` on the wrong side of the round's trace mark and no test
  // could see it: the defect is WHERE this is called, and a stub that traces
  // nothing cannot express where. `traceRef.fn` is pointed at the real `trace`
  // by the test that cares.
  traceEnvironment: vi.fn((build: string) => traceRef.fn("host", "environment", { build })),
  wantsAutoPicture: vi.fn(() => host.autoPicture),
  // What this host cannot draw AT ALL — empty for every test that does not say
  // otherwise, which is a host that can draw everything. A real export has to
  // be listed here or it arrives `undefined` and the call throws: this module
  // is mocked whole, and adding `marksThisHostWillDrop` to the renderer took 55
  // tests down before it appeared in this object.
  marksThisHostWillDrop: vi.fn(() => host.cannotDraw),
  // Selection juggling around an in-place redraw. The fake host has no view,
  // so it just runs the body — with `deselected` false, which is the honest
  // answer for a host that cannot move the selection.
  withSlideDeselected: vi.fn(async (ids: string[], fn: (d: boolean) => Promise<unknown>) => {
    host.calls.deselected.push(ids);
    return fn(false);
  }),
  slideHoldsOnlyChart: vi.fn(async () => host.slideHoldsOnlyChart),
  // Three outcomes, not two: "duplicated" is the swap that inserted the new
  // slide but could not remove the old one.
  replaceSlideWithDeck: vi.fn(async () => host.swapOutcome),
  newRunId: vi.fn(() => "run-under-test"),
  OFFSCREEN_BATCH: 40,
  insertSlidesFromPptx: vi.fn(async (b64: string, expected: number) => {
    host.calls.insertFile.push({ b64, expected });
    if (host.insertFileError) throw host.insertFileError;
    return host.insertFileLands ?? expected;
  }),
  /**
   * The own-slide-as-a-file route, which moved OUT of `app.ts` and into the
   * renderer so a self-test scenario could drive the shipped code rather than a
   * copy of it. Mocked here because it is now a renderer import.
   *
   * SWALLOWS, LIKE THE REAL ONE. `chartOnItsOwnSlideAsFile` catches its own
   * errors and returns 0, because every failure on that path is a fallback and
   * the caller's floor is the picture. A mock that threw instead would make the
   * pane's fallback look broken in exactly the tests written to prove it works.
   */
  chartOnItsOwnSlideAsFile: vi.fn(async () => {
    host.calls.insertFile.push({ b64: "generated-one-slide", expected: 1 });
    if (host.insertFileError) return 0;
    return host.insertFileLands ?? 1;
  }),
  reconcileDeck: vi.fn(async () => host.reconcileOutcome),
  applyReconcilePlan: vi.fn(async () => host.reconcileOutcome),
  snapshotAddedSlides: vi.fn(async () => []),
  readAddedSlides: vi.fn(async () => ({ snapshots: [], unread: 0 })),
  slideCount: vi.fn(async () => {
    host.slideCountCalls++;
    if (host.slideCountThrowsAfter !== null && host.slideCountCalls > host.slideCountThrowsAfter) {
      throw new Error("host would not report the slide count");
    }
    return host.slideCount;
  }),
  // What the host probe needs, and only that. The probe's own answers are
  // covered exhaustively in `host-probe.test.ts` against the real fake host;
  // what this file cares about is that "Run the whole round" puts a COMPLETE
  // sheet and the self-test's verdicts into ONE file. A host that will not give
  // out a scratch slide produces a complete sheet of `no-scratch-slide`, which
  // is exactly enough to test the bundling and nothing more.
  addScratchSlide: vi.fn(async () => null),
  // No chart this host has agreed to name.
  //
  // This mock's whole point is a host that refuses a scratch slide, so every
  // question comes back `no-scratch-slide` — not `no-scratch-shape` — and the
  // round's re-ask has nothing to re-ask. Which is the honest state for a mock
  // that never draws: `namedShape` is set by a chart whose TAG WROTE, and
  // nothing here writes one.
  namedShape: vi.fn(() => null),
  // No chart to observe either. This mock refuses a scratch slide and draws
  // nothing, so a scan finds no chart and the round's re-ask has nothing to ask
  // with — which is the honest state rather than a stubbed convenience.
  refreshNamedShapeFromDeck: vi.fn(async () => null),
  // The sync counter. A real number rather than a `vi.fn()` returning
  // undefined, because `app.ts` writes `syncs: syncsSoFar()` straight into the
  // round file — a mock answering undefined would let a round ship with the
  // field silently missing and every test here still pass.
  syncsSoFar: vi.fn(() => 0),
  resetSyncCount: vi.fn(() => {
    if (host.syncResetThrows) throw new Error("the pane's own code fell over");
  }),
  deleteSlideById: vi.fn(async (id: string) => {
    host.deletedSlides.push(id);
    return !host.refuseSlideDelete.includes(id);
  }),
  isTimeout: vi.fn(() => false),
  requirementSets: vi.fn(() => ["1.1", "1.5"]),
  ScratchSlideUnavailable: class ScratchSlideUnavailable extends Error {},
  withProbeContext: vi.fn(async () => {
    throw new Error("no scratch slide in this harness");
  }),
  deadlinesFired: 0,
  // The other half of a stall report: a scenario that gives up asks whether the
  // call ever came back. This harness has no host to answer late, so the double
  // says "nothing arrived" — which is the reading every real round has produced.
  lastLateSync: null,
  lastLateSyncSeq: 0,
  waitForLateSync: vi.fn(async () => false),
}));

// The deck builder is a real pptxgenjs run; the pane's own tests care about
// which path was taken, not about the bytes (test/pptx-deck.test.ts covers those).
vi.mock("../src/render/pptx-deck", () => ({
  buildDeckBase64: vi.fn(async (items: unknown[]) => {
    if (host.buildFileError) throw host.buildFileError;
    // The real builder reports what it actually drew per slide — the pane
    // verifies against THAT, not against the Office.js shape estimate.
    return { base64: "UEsDBBQA-fake-base64", shapesPerSlide: items.map(() => 7) };
  }),
}));

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Let the clicked handler's async chain (and its busy→done note) settle. */
const settle = () => new Promise((r) => setTimeout(r, 0));

/**
 * Boot app.ts with a PowerPoint host present. app.ts wires the host buttons only
 * inside `Office.onReady`, so the stub must both look like a host (mocked
 * `isPowerPointHost`) and fire onReady synchronously at import.
 */
async function bootHostPane(opts?: { deckStyle?: Record<string, unknown> | null; keepPicturePref?: boolean }) {
  host.selectionBounds = null;
  host.slideShapes = [];
  host.deckCharts = [];
  host.deckScanUnread = 0;
  host.deckInventory = [];
  host.deckSlideIds = undefined;
  host.canRaster = true;
  host.deckScanThrows = false;
  host.platform = "OfficeOnline";
  host.deckInventoryScans = 0;
  host.deckScanHangsOnScan = null;
  host.deletedSlides = [];
  host.selectionDropped = [];
  host.deckSlideIdCalls = 0;
  host.deckSlideIdsBeforeFails = false;
  host.addSlideResult = "own-slide-1";
  host.addSlideCalls = 0;
  host.roundAddsSlides = ["probe-a", "probe-b"];
  host.refuseSlideDelete = [];
  host.selectionCharts = [];
  host.loadSelectionResult = null;
  host.gate = null;
  host.failInsertOnce = false;
  host.selectionListener = null;
  host.selectionBudgets = [];
  host.selectionGate = null;
  host.agendaSlides = [];
  // Read at BOOT (the pane asks the deck on `Office.onReady`), so a test that
  // wants a branded deck has to say so here rather than after the fact. The
  // parameter is optional because `beforeEach(bootHostPane)` hands this a test
  // context, which carries no `deckStyle` and so resets it like everything else.
  host.deckStyle = opts?.deckStyle ?? null;
  host.deckStyleWrites = [];
  host.deckStyleWritable = true;
  // Reset with its siblings, or the one test that sets it leaves every later
  // test in this file reading a host that will not answer — an ordering-
  // dependent failure, which is the worst kind to debug.
  host.deckStyleUnreadable = false;
  host.demoRuns = 0;
  host.demoDeckCalls = [];
  host.demoDeckOpts = [];
  host.demoDeckGate = null;
  host.demoDeckGateOn = 0;
  host.demoDeckPageFailures = new Set();
  host.demoDeckStatusOverride = null;
  host.canInsertFile = false;
  host.slideHoldsOnlyChart = false;
  host.updateChartThrows = false;
  host.demoReconcile = undefined;
  host.swapOutcome = "failed";
  host.autoPicture = false;
  host.cannotDraw = { what: [], nodes: 0 };
  host.updateGate = null;
  host.insertGate = null;
  host.slideCountThrowsAfter = null;
  host.slideCountCalls = 0;
  host.syncResetThrows = false;
  host.insertFileLands = null;
  host.insertFileError = null;
  host.buildFileError = null;
  host.slideCount = 1;
  host.reconcileOutcome = undefined;
  host.calls.insertFile.length = 0;
  host.calls.deselected.length = 0;
  host.canPicture = true;
  host.updateResult = undefined;
  host.insertResult = null;
  host.calls.swept.length = 0;
  host.calls.insertScene = [];
  host.calls.updateChart = [];
  host.calls.updateCharts = [];
  host.updateChartsStalls.clear();
  host.updateChartsDrops = 0;
  host.stopRequested = false;
  host.slideSize = { width: 960, height: 540, source: "pageSetup" };

  /**
   * A FRESH BROWSER'S PREFERENCES, not the previous test's.
   *
   * `Insert as picture` became a REMEMBERED default on 2026-09-02, so
   * `tickPicture()` now writes to localStorage and every later boot reads it
   * back. That is exactly the behaviour asked for — tick it once and new charts
   * stay pictures — and it leaks straight through a shared jsdom localStorage
   * into every test that runs after it.
   *
   * It surfaced as eleven failures in tests with nothing to do with pictures:
   * `insertScene` came back empty, or carried a payload nothing had asked for,
   * because the pane they booted had inherited a preference from a test earlier
   * in the file. Every one of them passed when run alone.
   *
   * Only this key is cleared. Templates and the deck style live in localStorage
   * too, and their own tests set those up deliberately.
   *
   * `keepPicturePref` is the opt-out, for the one test whose SUBJECT is that the
   * preference survives a reload — it needs the leak this prevents.
   */
  if (!opts?.keepPicturePref)
    try {
      localStorage.removeItem("ssf-charts-render-image");
    } catch {
      /* a jsdom without storage is still a valid pane */
    }

  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;

  vi.stubGlobal("Office", {
    // OFFICE BECOMES AVAILABLE HERE, not before — which is the whole ordering
    // the real host has and this harness used to paper over.
    onReady: (cb: () => void) => {
      hostReadiness.officeArrives();
      cb();
    },
    EventType: { DocumentSelectionChanged: "DocumentSelectionChanged" },
    // `isWebHost()` lives in app.ts and is NOT part of the mocked renderer, so
    // without these the pane always believed it was on the desktop and every
    // web-only guard was untestable. Reverted once as noise; a surviving mutant
    // (`web: false`) proved it is not.
    PlatformType: { OfficeOnline: "OfficeOnline", PC: "PC", Mac: "Mac" },
    context: {
      host: "PowerPoint",
      displayLanguage: "en-US",
      // A GETTER: this object is built once per boot, so a test setting
      // `host.platform` afterwards would otherwise be silently ignored.
      get diagnostics() {
        return { platform: host.platform };
      },
      document: {
        // watchSelection() registers a selection listener; capture it so a test
        // can fire it the way PowerPoint does when the user clicks a shape.
        addHandlerAsync: (_type: string, handler: () => unknown) => {
          host.selectionListener = handler;
        },
      },
    },
  });

  vi.resetModules();
  await import("../src/taskpane/app");
  await settle();
}

/** A value-axis chart config with a known extent, as a JSON tag would carry it. */
const chartJson = (values: number[]): string =>
  JSON.stringify({
    kind: "stacked",
    width: 480,
    height: 320,
    data: { categories: values.map((_, i) => `C${i}`), series: [{ name: "S", values }] },
  } satisfies ChartConfig);

afterEach(() => vi.unstubAllGlobals());

/**
 * Stand in for the browser rasteriser. jsdom decodes no SVG, `getContext("2d")`
 * returns null and `toDataURL` returns null, so every step of `rasterizeScene`
 * needs a double. Pass `{ decode: false }` to model a browser that cannot decode
 * the SVG, or `{ encode: false }` for one whose canvas cannot produce a PNG.
 * Returns a restore function.
 */
function stubRaster({ decode = true, encode = true } = {}) {
  class FakeImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_v: string) {
      Promise.resolve().then(() => (decode ? this.onload?.() : this.onerror?.()));
    }
  }
  vi.stubGlobal("Image", FakeImage);
  const ctx = vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue({
    scale() {},
    drawImage() {},
  } as unknown as CanvasRenderingContext2D);
  const url = vi
    .spyOn(HTMLCanvasElement.prototype, "toDataURL")
    .mockReturnValue(encode ? "data:image/png;base64,RASTER" : (null as unknown as string));
  // Capture the SVG each rasterisation was handed. Without this a test can only
  // see THAT a payload was produced, never OF WHAT — and the stub returns the
  // same bytes whatever it is given, so "a payload exists" passes even when the
  // wrong scene was rasterised.
  const svgs: Blob[] = [];
  const c = vi.spyOn(URL, "createObjectURL").mockImplementation((blob: Blob | MediaSource) => {
    if (blob instanceof Blob) svgs.push(blob);
    return "blob:x";
  });
  const r = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
  return {
    rasterizedSvgs: () => Promise.all(svgs.map((b) => b.text())),
    restore: () => {
      for (const s of [ctx, url, c, r]) s.mockRestore();
    },
  };
}

/** Tick "Insert as picture", the way a user does. */
const tickPicture = () => {
  const box = $("render-image") as HTMLInputElement;
  box.checked = true;
  box.dispatchEvent(new Event("change"));
};

/** Capture whatever `downloadJson` hands to the browser. */
function captureDownloads() {
  const blobs: Blob[] = [];
  const c = vi.spyOn(URL, "createObjectURL").mockImplementation((b: Blob | MediaSource) => {
    if (b instanceof Blob) blobs.push(b);
    return "blob:x";
  });
  const r = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
  return {
    lastJson: async () => JSON.parse(await blobs[blobs.length - 1].text()),
    count: () => blobs.length,
    restore: () => {
      c.mockRestore();
      r.mockRestore();
    },
  };
}

describe("image render mode", () => {
  beforeEach(bootHostPane);

  it("hands the insert a rasterised PNG when Insert as picture is ticked", async () => {
    const raster = stubRaster();
    try {
      tickPicture();
      $("insert").click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(1);
      // The payload reached the renderer, bare — the pane strips the data: URI.
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBe("RASTER");
      // And the config it tagged says image mode, so a re-edit knows what it is.
      expect(JSON.parse(host.calls.insertScene[0].tagData!).render).toBe("image");
    } finally {
      raster.restore();
    }
  });

  it("sends no payload at all in shapes mode, so nothing pays for the canvas", async () => {
    const raster = stubRaster();
    try {
      $("insert").click();
      await settle();
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
      expect(JSON.parse(host.calls.insertScene[0].tagData!).render).toBeUndefined();
    } finally {
      raster.restore();
    }
  });

  it("skips the raster and SAYS SO on a host below PowerPointApi 1.8", async () => {
    // The note must survive to the end: insertSceneIntoSlide's first act is
    // onPhase("context"), which overwrites the pane note, so a warning posted
    // BEFORE the insert would be wiped. It is posted after, on purpose.
    const raster = stubRaster();
    try {
      host.canPicture = false;
      tickPicture();
      $("insert").click();
      await settle();
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
      expect($("host-note").textContent).toMatch(/can't insert pictures/i);
    } finally {
      raster.restore();
    }
  });

  it("inserts a picture, and says why, when this host cannot draw the marks", async () => {
    /**
     * WARN AND PICTURE — the owner's call, 2026-08-31, for the 18 of 123
     * shipped charts that lose ink below PowerPointApi 1.10 and the 8 that lose
     * their subject outright: a pie whose wedges cannot be drawn inserts as
     * four labels around empty space.
     *
     * The picture is the whole chart, correct, on a host that cannot draw it.
     * The message has to say WHAT would have been missing — "this chart could
     * not be drawn" gives a user nothing to act on.
     */
    const raster = stubRaster();
    try {
      host.canPicture = true;
      host.cannotDraw = { what: ["4 pie slices"], nodes: 4 };
      $("insert").click();
      await settle();
      expect(
        (host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64,
        "a chart this host cannot draw went in as shapes",
      ).toBeTruthy();
      const note = String($("host-note").textContent);
      expect(note, "did not say what was missing").toMatch(/4 pie slices/);
      // NOT an invitation to explode: exploding is the door back to the version
      // this host cannot draw, which is the mistake the density guard's own
      // comment records the product having made once already.
      expect(note, "offered Explode, which would put the chart back the way it could not be drawn").not.toMatch(
        /explode/i,
      );
    } finally {
      raster.restore();
    }
  });

  it("still says what is missing when it cannot make a picture either", async () => {
    // The chart goes in as shapes, because that beats refusing it — but
    // silently dropping the marks is the one outcome this must never produce.
    const raster = stubRaster();
    try {
      host.canPicture = false;
      host.cannotDraw = { what: ["an arrow"], nodes: 1 };
      $("insert").click();
      await settle();
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
      const note = String($("host-note").textContent);
      expect(note).toMatch(/an arrow/);
      expect(note, "did not say the mark would be missing").toMatch(/missing/i);
    } finally {
      raster.restore();
    }
  });

  it("degrades to native shapes, visibly, when the browser cannot rasterise", async () => {
    for (const mode of [{ decode: false }, { encode: false }] as const) {
      await bootHostPane();
      const raster = stubRaster(mode);
      try {
        tickPicture();
        $("insert").click();
        await settle();
        // The chart still landed — as shapes, with no payload.
        expect(host.calls.insertScene, JSON.stringify(mode)).toHaveLength(1);
        expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
        expect($("host-note").textContent, JSON.stringify(mode)).toMatch(/rasterize/i);
      } finally {
        raster.restore();
      }
    }
  });

  it("batch insert honours render per config, not the checkbox", async () => {
    // #json-insert-batch calls insertSceneIntoSlide directly and used to ignore
    // the mode outright, so a pasted image-mode array silently drew shapes.
    const raster = stubRaster();
    try {
      ($("json-io") as HTMLTextAreaElement).value = JSON.stringify([
        { kind: "stacked", render: "image", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } },
        { kind: "stacked", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } },
      ]);
      $("json-insert-batch").click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(2);
      const payloads = host.calls.insertScene.map((o) => (o as { pictureBase64?: string }).pictureBase64);
      expect(payloads).toEqual(["RASTER", undefined]);
    } finally {
      raster.restore();
    }
  });

  it("Same scale rasterises the RESCALED chart, not the one it loaded", async () => {
    // The scale mutation happens inside the map that builds each scene, so
    // rasterising in the same pass captures the PRE-scale chart: the picture
    // shows the old axis while the pane reports success.
    //
    // Asserting "a payload exists" does NOT catch that — the stub returns the
    // same bytes for any input, and tagData reads the same mutated object either
    // way. So compare the SVG actually handed to the rasteriser against the two
    // candidate renders. A chart whose own extent is [1,5] rescaled to max 90
    // draws visibly different bars, so the two SVGs differ.
    const raster = stubRaster();
    try {
      const own = JSON.parse(chartJson([1, 5])) as ChartConfig;
      host.deckCharts = [
        { configJson: JSON.stringify({ ...own, render: "image" }), target: { shapeId: "a" } },
        { configJson: chartJson([2, 90]), target: { shapeId: "b" } },
      ];
      $("same-scale").click();
      await settle();

      const [rasterized] = await raster.rasterizedSvgs();
      expect(rasterized, "exactly one chart was rasterised").toBeTruthy();
      const svgOf = (cfg: ChartConfig) => sceneToSvg(buildChart(cfg));
      const preScale = svgOf({ ...own, render: "image" });
      const postScale = svgOf({ ...own, render: "image", scale: { min: undefined, max: 90 } });
      expect(preScale, "the two candidates must actually differ").not.toBe(postScale);
      expect(rasterized).toBe(postScale);
      expect(rasterized).not.toBe(preScale);
    } finally {
      raster.restore();
    }
  });

  it("names the rescued charts even when others fell back, and keeps the counts", async () => {
    /**
     * THREE NOTES, ONE SLOT. `doSameScale` posted the applied count, then a
     * rescued clause, then a degraded clause, straight into a channel that
     * holds one — so when both qualifiers fired the user saw only the last, and
     * lost the rescued count, the chart count and the shared max with it.
     *
     * The rescued charts are the ones that matter here: they went in as
     * PICTURES, and "Explode to native shapes" is the way back. Saying nothing
     * about them steers the user away from the one control that recovers
     * them — which is the same harm the test below this one guards against,
     * reached from the other side.
     *
     * A MIXED RUN NEEDS A RASTER THAT SUCCEEDS ONCE AND THEN FAILS. Both flags
     * that drive this are global to the fake, so the only per-chart lever is
     * the canvas: chart one gets a png (warn + png = rescued), chart two does
     * not (warn, no png = degraded).
     */
    class FakeImage {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      set src(_v: string) {
        Promise.resolve().then(() => this.onload?.());
      }
    }
    vi.stubGlobal("Image", FakeImage);
    const ctx = vi
      .spyOn(HTMLCanvasElement.prototype, "getContext")
      .mockReturnValue({ scale() {}, drawImage() {} } as unknown as CanvasRenderingContext2D);
    const url = vi
      .spyOn(HTMLCanvasElement.prototype, "toDataURL")
      .mockReturnValueOnce("data:image/png;base64,RASTER")
      .mockReturnValue(null as unknown as string);
    const co = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:x");
    const re = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    try {
      host.autoPicture = true;
      host.canPicture = true;
      host.deckCharts = [
        { configJson: chartJson([1, 5]), target: { shapeId: "a" } },
        { configJson: chartJson([2, 90]), target: { shapeId: "b" } },
      ];
      $("same-scale").click();
      await settle();

      const said = $("host-note").textContent ?? "";
      expect(said, `dropped the rescued charts entirely: ${said}`).toMatch(/went in as pictures/i);
      expect(said, `dropped the charts that fell back: ${said}`).toMatch(/fell back to native shapes/i);
      expect(said, `dropped the count and the shared max: ${said}`).toMatch(/max/i);
    } finally {
      for (const s of [ctx, url, co, re]) s.mockRestore();
    }
  });

  it("does not report a successful auto-picture rescue as a fallback to shapes", async () => {
    // `chartPicture` returns a warn WITH a png on its SUCCESS path: a chart too
    // dense for the web host, rasterised and inserted as a picture, the warn
    // being the explanation. Treating every warn as a failure printed, in red,
    // "N image chart(s) fell back to native shapes" — over charts that went in
    // as PICTURES, overwriting the true success line, claiming the dangerous
    // thing had happened when the guard against it had just worked, and
    // steering the user away from "Explode to native shapes".
    const raster = stubRaster();
    try {
      host.autoPicture = true;
      host.canPicture = true;
      host.deckCharts = [
        { configJson: chartJson([1, 5]), target: { shapeId: "a" } },
        { configJson: chartJson([2, 90]), target: { shapeId: "b" } },
      ];
      $("same-scale").click();
      await settle();
      // Both really are pictures — the premise of the complaint.
      const payloads = host.calls.updateCharts.at(-1) ?? [];
      expect(payloads.length, "no update was issued").toBeGreaterThan(0);
      expect(
        payloads.every((it) => !!(it.opts as { pictureBase64?: string } | undefined)?.pictureBase64),
        "the rescue did not produce pictures, so this proves nothing",
      ).toBe(true);
      expect($("host-note").textContent ?? "").not.toMatch(/fell back to native shapes/i);
      expect($("host-note").className, "a successful rescue was reported in red").not.toMatch(/status-err/);
    } finally {
      raster.restore();
    }
  });

  it("remembers the choice for the NEXT chart, and does not impose it on an existing one", async () => {
    /**
     * THE STANDING HALF OF THE CHOICE. `Insert as picture` has always been a
     * per-chart tick that reset on every new chart, so anyone who wanted
     * pictures had to remember every single time. It is a remembered default
     * now — and a DEFAULT is all it is.
     *
     * The second half is the one that could do damage. `stateFromConfig` also
     * runs when a chart is loaded off a slide, so seeding the preference there
     * would silently rewrite somebody's existing shapes chart into a picture the
     * moment they opened it to change a label. A chart's own `render` is a fact
     * about that chart; the preference is a wish about the next one.
     */
    const raster = stubRaster();
    try {
      tickPicture();
      expect(localStorage.getItem("ssf-charts-render-image"), "the choice was not remembered").toBe("1");

      // A NEW pane, as if the add-in were reopened tomorrow.
      await bootHostPane({ keepPicturePref: true });
      expect(
        ($("render-image") as HTMLInputElement).checked,
        "the remembered preference did not survive a reload",
      ).toBe(true);

      // And a chart that went in as SHAPES still opens as shapes, preference or
      // not — this is the assertion that stops the default becoming an override.
      // Imported through the JSON box, which is the same `stateFromConfig` path
      // that opening a chart off a slide takes.
      ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({
        kind: "clustered",
        data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
        title: "a chart that went in as shapes",
      });
      $("json-import").click();
      await settle();
      expect(
        ($("render-image") as HTMLInputElement).checked,
        "opening a shapes chart silently turned it into a picture",
      ).toBe(false);

      // Unticking is remembered too, or the preference would be one-way.
      const box = $("render-image") as HTMLInputElement;
      box.checked = false;
      box.dispatchEvent(new Event("change"));
      expect(localStorage.getItem("ssf-charts-render-image")).toBe("0");
    } finally {
      raster.restore();
    }
  });
});

describe("Explode to native shapes", () => {
  beforeEach(bootHostPane);

  it("re-draws the selected picture chart as shapes by OMITTING the payload", async () => {
    // The omission is the explode: no pictureBase64 means wantsPicture is false
    // in the renderer, so the node loop runs. Nothing else about the update
    // differs, which is why this needs no renderer code.
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20 };
    $("explode").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(1);
    const { opts } = host.calls.updateChart[0];
    expect((opts as { pictureBase64?: string }).pictureBase64).toBeUndefined();
    // The re-saved config says shapes, so the chart does not turn back into a
    // picture on the next ordinary update.
    expect(JSON.parse(opts.tagData!).render).toBe("shapes");
    expect($("host-note").textContent).toMatch(/native shapes/i);
    // The checkbox follows the exploded config — the pane must not still claim
    // the next insert will be a picture.
    expect(($("render-image") as HTMLInputElement).checked).toBe(false);
  });

  it("refuses politely when the selection is not an SSF chart", async () => {
    host.loadSelectionResult = null;
    $("explode").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(0);
    expect($("host-note").textContent).toMatch(/select an inserted chart first/i);
  });

  it("does not claim success when the host would not save the config back", async () => {
    // The same host answer the ordinary update path handles carefully — and on
    // the web it is ordinary: five `no-config` results in a single recorded
    // round. Explode used to print GREEN over a chart that could no longer be
    // opened, and then adopt the lost target as the live edit target, so the
    // next push resolved a dead id and did nothing at all.
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20, lost: "no-config" };
    $("explode").click();
    await settle();
    expect($("host-note").className, "a lost chart was reported as a success").toMatch(/status-err/);
    expect($("host-note").textContent).toMatch(/no longer editable/i);
    // And the dead target is not kept: the Insert button must not offer Update
    // against a chart the pane cannot resolve.
    expect($("insert").textContent).not.toMatch(/update/i);
  });

  it("does not claim success when the host would not say where the shapes landed", async () => {
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20, lost: "unknown-shape" };
    $("explode").click();
    await settle();
    expect($("host-note").className).toMatch(/status-err/);
    expect($("host-note").textContent).toMatch(/lost track/i);
  });

  it("says so when the picture is already gone from the slide", async () => {
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = undefined; // the update found nothing to replace
    $("explode").click();
    await settle();
    expect($("host-note").textContent).toMatch(/no longer on the slide/i);
  });
});

describe("Insert", () => {
  beforeEach(bootHostPane);

  it("tiles a new chart at the default offset when nothing is selected", async () => {
    host.selectionBounds = null;
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    const opts = host.calls.insertScene[0];
    // Default drop point, and the config round-trips as the shape's tag.
    expect(opts.left).toBeGreaterThanOrEqual(60);
    expect(opts.top).toBeGreaterThanOrEqual(90);
    expect(JSON.parse(opts.tagData!)).toMatchObject({ kind: expect.any(String) });
    expect(host.calls.updateChart).toHaveLength(0);
  });

  /**
   * Two charts onto one slide. The insert path never looked at what was
   * already there — its whole answer to "where does this go" was a 14pt
   * cascade, and against a 480x300 chart that is a 90%-plus overlap. The
   * second chart landed on the first, which is what a user sees as one chart
   * drawn over another.
   */
  it("puts a second chart clear of the first instead of on top of it", async () => {
    const first = { left: 60, top: 90, width: 480, height: 300 };
    host.slideShapes = [first];
    $("insert").click();
    await settle();
    const at = host.calls.insertScene.at(-1)!;
    const cfg = JSON.parse(at.tagData!) as { width?: number; height?: number };
    const box = { left: at.left!, top: at.top!, width: cfg.width ?? 480, height: cfg.height ?? 300 };
    const hits =
      box.left < first.left + first.width &&
      first.left < box.left + box.width &&
      box.top < first.top + first.height &&
      first.top < box.top + box.height;
    expect(hits).toBe(false);
  });

  it("places the second chart BESIDE the first on a 16:9 deck", async () => {
    // The slide WIDTH decides this, and it is the number the pane spent its
    // life assuming. On 16:9 there is a 390pt band beside the first chart.
    const first = { left: 60, top: 90, width: 480, height: 300 };
    host.slideShapes = [first];
    host.slideSize = { width: 960, height: 540, source: "pageSetup" };
    $("insert").click();
    await settle();
    const at = host.calls.insertScene.at(-1)!;
    expect(at.left, "did not sit beside the first chart").toBeGreaterThanOrEqual(first.left + first.width);
    expect(at.top).toBe(first.top);
    expect(at.left! + (JSON.parse(at.tagData!) as { width: number }).width).toBeLessThanOrEqual(960);
    expect($("host-note").textContent?.toLowerCase()).toContain("beside");
  });

  it("drops the second chart BELOW instead on a 4:3 deck", async () => {
    // Same two charts, 240pt less width. Beside leaves 150pt, which scales the
    // chart under the readability floor — so down it goes. A pane that assumed
    // 16:9 here would run a chart off the right edge of every 4:3 deck.
    //
    // A fresh pane, not a second insert into the previous test's: doInsert
    // feeds the placed size back into the config it tags, so a follow-on insert
    // starts from the SHRUNK size and would be measuring two variables at once.
    const first = { left: 60, top: 90, width: 480, height: 300 };
    host.slideShapes = [first];
    host.slideSize = { width: 720, height: 540, source: "pageSetup" };
    $("insert").click();
    await settle();
    const at = host.calls.insertScene.at(-1)!;
    expect(at.left, "squeezed a chart beside on a 4:3 deck").toBe(first.left);
    expect(at.top, "did not drop below the first chart").toBeGreaterThanOrEqual(first.top + first.height);
    // And it stays on the slide, which is the risk horizontal placement adds.
    expect(at.left! + (JSON.parse(at.tagData!) as { width: number }).width).toBeLessThanOrEqual(720);
  });

  /**
   * A host that will not describe the slide must get the CASCADE, and the
   * cascade has to actually move.
   *
   * `getSlideShapeBounds` used to swallow a refusal into `[]`, with a comment
   * claiming placement then "falls back to the cascade, which is what it always
   * did". It does not: `placeChart` reads an empty `occupied` as "there is room
   * everywhere", so `placeBeside` succeeds on its first pass and returns the
   * origin UNMOVED — its `fallback` argument is unreachable. Two inserts onto a
   * slide the host would not read therefore landed on exactly the same point,
   * which is the pile the placement rule exists to prevent and worse than the
   * fixed cascade it replaced. A real host refused every shape read on a whole
   * deck (`unread=8 slides=8`), so this is its ordinary behaviour.
   */
  it("cascades instead of stacking when the host will not say what is on the slide", async () => {
    host.slideShapes = null;
    $("insert").click();
    await settle();
    const first = host.calls.insertScene.at(-1)!;
    $("insert").click();
    await settle();
    const second = host.calls.insertScene.at(-1)!;
    expect(host.calls.insertScene).toHaveLength(2);
    expect({ left: second.left, top: second.top }, "two charts landed on exactly the same point").not.toEqual({
      left: first.left,
      top: first.top,
    });
  });

  it("fits the chart to a selected placeholder's bounds", async () => {
    host.selectionBounds = { left: 200, top: 150, width: 360, height: 240 };
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    const opts = host.calls.insertScene[0];
    expect(opts.left).toBe(200);
    expect(opts.top).toBe(150);
    // The placeholder's size overrides the config's own dimensions.
    const cfg = JSON.parse(opts.tagData!) as ChartConfig;
    expect(cfg.width).toBe(360);
    expect(cfg.height).toBe(240);
  });

  /**
   * The user's own shape, not destroyed by inserting a chart beside it.
   *
   * office-js#2775: on PowerPoint on the web, adding a text box deletes the
   * shape that was selected. Every chart drawn here contains text boxes, and
   * this path deliberately leaves the selection alone because that is how it
   * learns where to put the chart — so a picture selected to position a chart
   * against would be silently destroyed by the insert. office-js#3698 is the
   * same setup failing the other way: a picture cannot be inserted while a
   * shape is selected, and this path inserts a picture for a dense chart.
   */
  it("lets go of the user's selected shape before it draws anything", async () => {
    host.selectionBounds = { left: 200, top: 150, width: 360, height: 240 };
    $("insert").click();
    await settle();
    // The bounds still decide the placement — reading the selection and holding
    // onto it are different things, and only the second one is the hazard.
    expect(host.calls.insertScene[0].left).toBe(200);
    // BEFORE the draw, which is the entire protection. `[0]` is "no charts had
    // been drawn yet when the selection was dropped"; a `[1]` here would be a
    // pane that let go after the damage.
    expect(host.selectionDropped, "drew with the user's shape still selected").toEqual([0]);
  });

  it("ignores a selection too small to be a real placeholder", async () => {
    host.selectionBounds = { left: 5, top: 5, width: 20, height: 20 };
    $("insert").click();
    await settle();
    // Below the 40pt threshold → falls back to the tiled offset, not the bounds.
    expect(host.calls.insertScene[0].left).toBeGreaterThanOrEqual(60);
  });
});

/**
 * THE SLOW-INSERT OFFER, end to end from the pane.
 *
 * `insert-cost.ts` is well covered and `addSlideForChart` has its own block in
 * test/office-render.test.ts — and between the two NOTHING asserted that taking
 * the offer actually sends the chart to the new slide. Both halves were green
 * with the wire that joins them untested, which is the shape of gap this file
 * exists to close.
 *
 * The trigger is arithmetic, not a magic number: `batchMs` is flat at 18074ms
 * past 75 shapes present, so a crowded slide puts ANY chart over the 14s
 * threshold, and the empty-slide estimate is 3886ms — comfortably under half,
 * which is what `worthOwnSlide` asks for.
 */
describe("the slow-insert offer", () => {
  beforeEach(bootHostPane);

  /**
   * A slide holding `n` shapes, enough to price an insert onto it AND to move
   * it. The first version tiled a narrow column at x<10, which priced the
   * insert correctly and left `placeChart` returning the origin unmoved — so
   * the "placement is recomputed" assertion below passed against a mutant that
   * carried the crowded placement over. These shapes cover the region a chart
   * lands in, so the two placements are genuinely different numbers.
   */
  const crowd = (n: number) =>
    Array.from({ length: n }, (_, i) => ({
      left: 40 + (i % 4) * 120,
      top: 70 + Math.floor(i / 4) * 20,
      width: 110,
      height: 18,
    }));

  /**
   * Press Insert and let the pane get as far as the offer.
   *
   * `settle()` is one macrotask and the offer sits behind an awaited host read,
   * so a single tick is not reliably enough. Polling for the box to appear says
   * what the test means — "wait until the user is being asked" — and fails with
   * a readable message instead of a null-click when the offer never comes.
   */
  const untilOffered = async () => {
    for (let i = 0; i < 20 && $("slow-offer").hidden; i++) await settle();
    expect($("slow-offer").hidden, "the pane never offered a slide").toBe(false);
  };

  it("offers, and sends the chart to the new slide when the offer is taken", async () => {
    host.slideShapes = crowd(80);
    $("insert").click();
    await untilOffered();
    // The sentence quotes BOTH costs. A slogan ("nearly instant") is the thing
    // `offerSentence` exists to prevent, so assert the shape rather than the text.
    expect($("slow-offer-text").textContent).toMatch(/80 shapes/);
    $("slow-offer-own").click();
    await settle();
    await settle();

    expect(host.addSlideCalls, "the offer was taken and no slide was added").toBe(1);
    const at = host.calls.insertScene.at(-1)!;
    expect(at.slideId, "the chart did not go to the slide that was just added").toBe("own-slide-1");
    // AND THE PLACEMENT IS RECOMPUTED. `at` was cascaded around 80 shapes on the
    // crowded slide; carrying it over would drop the chart into a corner of a
    // blank one, offset from obstacles that are not there.
    expect({ left: at.left, top: at.top }).toEqual({ left: 60, top: 90 });
  });

  it("inserts where the user asked when the offer is declined", async () => {
    host.slideShapes = crowd(80);
    $("insert").click();
    await untilOffered();
    $("slow-offer-here").click();
    await settle();

    expect(host.addSlideCalls, "declining the offer still bought a slide").toBe(0);
    const at = host.calls.insertScene.at(-1)!;
    expect(at.slideId, "declining still redirected the chart").toBeUndefined();
    // The other half of the recompute assertion above: this crowd really does
    // move the chart, so `{60, 90}` there is the new slide's placement and not
    // the origin surviving by accident.
    expect({ left: at.left, top: at.top }, "the crowd did not displace the chart").not.toEqual({
      left: 60,
      top: 90,
    });
  });

  /**
   * The host drops `slides.add()`, so this is a Tuesday, not a defensive branch.
   * Losing the offer is a nuisance; losing the chart is not acceptable — so the
   * insert goes ahead on the slide the user picked, and says so.
   */
  it("still inserts, on the original slide, when the host drops the added slide", async () => {
    host.slideShapes = crowd(80);
    host.addSlideResult = null;
    $("insert").click();
    await untilOffered();
    $("slow-offer-own").click();
    await settle();
    await settle();

    expect(host.addSlideCalls).toBe(1);
    expect(host.calls.insertScene, "the chart was dropped when the slide add failed").toHaveLength(1);
    expect(host.calls.insertScene[0].slideId, "addressed a slide that was never created").toBeUndefined();
  });

  it("does not report a failed choice in green just because the placement worked", async () => {
    /**
     * THE BEHAVIOUR CHANGE MOST WORTH PINNING. On a tiled slide the placement
     * cascade reports "beside, scaled to fit" — an "ok" outcome — and that
     * green sentence used to be the ONLY thing the user saw, erasing a red
     * failure that had happened seconds earlier.
     *
     * Both facts now, and red: the chart did land, but not where the user
     * chose, and a choice that did not happen is not a success.
     */
    host.slideShapes = crowd(80);
    host.addSlideResult = null;
    $("insert").click();
    await untilOffered();
    $("slow-offer-own").click();
    await settle();
    await settle();

    const said = $("host-note").textContent ?? "";
    expect(said, `did not lead with the failed choice: ${said}`).toMatch(/^Could not add a slide/);
    expect(said, `dropped the placement outcome: ${said}`).toMatch(/placed beside/i);
    expect($("host-note").className, `a failed choice painted green: ${said}`).toMatch(/status-err/);
  });

  it("says the slide add failed, and does not end on a busy note", async () => {
    /**
     * THE SETBACK IS POSTED AND THEN DESTROYED. `note("Could not add a slide —
     * inserting here instead.")` runs before `insertSceneIntoSlide`, and that
     * call's first act is `phaseNote("context")` — a busy note into the same
     * single-slot channel. The rule is already written down one screen away:
     * "After the insert, never before: phaseNote would have overwritten it."
     *
     * TWO THINGS GO WRONG, and the second is the worse one. The setback is
     * non-busy, so it increments `settledNotes`; `guard` compares that counter
     * to decide whether to close with "Done."; and on this path no closing
     * branch applies — the chart is not lost, there is no `warn`, and nothing
     * moved or shrank. So "Done." is suppressed while the last text written was
     * `phaseNote("done")`, and the action ends showing "Working… done" in busy
     * blue with the progress bar still up. That is exactly the state
     * `settledNotes` was introduced to prevent.
     */
    /**
     * STACKED, not tiled. `crowd(80)` lays shapes across the whole slide, so
     * placement ends up "beside" and shrunk — a closing branch fires and the
     * setback is merely overwritten. Eighty shapes piled in one corner keep the
     * COUNT that makes `worthOwnSlide` fire while leaving the slide empty
     * enough that placement reports something the closing chain does not print
     * ("below", "cascade" or "none"), which is the case where nothing settles.
     */
    host.slideShapes = Array.from({ length: 80 }, () => ({ left: 40, top: 70, width: 110, height: 18 }));
    host.addSlideResult = null;
    $("insert").click();
    await untilOffered();
    $("slow-offer-own").click();
    await settle();
    await settle();

    const said = $("host-note").textContent ?? "";
    const cls = $("host-note").className;
    expect(said, `ended on a busy phase note: ${said}`).not.toMatch(/^Working…/);
    expect(cls, `ended in the busy state: ${cls}`).not.toMatch(/status-busy/);
    expect(said, `never told the user the slide add failed: ${said}`).toMatch(/could not add a slide/i);
  });

  /**
   * SLOW IS NOT THE SAME AS WORTH MOVING, and this is the case that separates
   * them. An empty slide is the fastest target there is, so however big the
   * chart, a new slide would be exactly as slow — `isSlowInsert` alone would
   * fire here and offer something it cannot deliver.
   */
  it("stays silent on an empty slide, however slow the chart", async () => {
    host.slideShapes = [];
    $("insert").click();
    await settle();
    await settle();
    expect($("slow-offer").hidden, "offered a new slide that would be no faster").toBe(true);
    expect(host.calls.insertScene).toHaveLength(1);
  });

  /**
   * `null` is the host refusing to say, not the slide being empty.
   *
   * AN EQUIVALENT MUTANT LIVES HERE, and it is worth recording rather than
   * chasing. Replacing the `occupied &&` guard with `occupied?.length ?? 0`
   * — pricing a refused read as an EMPTY slide — leaves every test green, and
   * no test could kill it: `worthOwnSlide(n, 0)` is false for every n, because
   * the empty-slide estimate can never be at most half of itself. So the two
   * readings of `null` produce identical behaviour today.
   *
   * The guard stays, and the comment in `app.ts` should be read as intent
   * rather than as protection: the day the offer's rule stops being a ratio
   * against the empty-slide cost, "unknown" and "empty" stop agreeing, and this
   * test starts telling the two apart on its own.
   */
  it("stays silent when the host will not say what is on the slide", async () => {
    host.slideShapes = null;
    $("insert").click();
    await settle();
    await settle();
    expect($("slow-offer").hidden, "priced a slide the host would not describe").toBe(true);
    expect(host.calls.insertScene).toHaveLength(1);
  });
});

describe("Insert updates in place after loading a chart", () => {
  beforeEach(bootHostPane);

  it("routes Insert to an update once a chart is loaded from the selection", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();

    // Now the primary Insert edits in place rather than dropping a new chart.
    $("insert").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(1);
    expect(host.calls.updateChart[0].target).toMatchObject({ shapeId: "shape-9" });
    expect(host.calls.insertScene).toHaveLength(0);
  });

  /**
   * `updateChartInSlide` returns null when the target slide or shape is gone —
   * deliberately, since "a chart whose slide is gone is not an error, it is
   * nothing to do". Nothing consumed the null: the guard saw an unchanged note
   * and printed "Done." in green, and the stale target kept the button reading
   * "Update chart", so every later push no-opped just as silently.
   */
  it("says so, instead of Done., when the target chart is gone", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    expect($("insert").textContent).toBe("Update chart");

    // The mocked update returns undefined — the "nothing to do" answer.
    $("insert").click();
    await settle();
    const note = $("host-note");
    expect(note.textContent).not.toBe("Done.");
    expect(note.className).toContain("status-err");
    // And the pane falls back to inserting, rather than pushing into thin air.
    expect($("insert").textContent).toBe("Insert into slide");
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
  });

  it("Insert-new always drops a fresh chart even with a chart loaded", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([4, 5]),
      target: { slideId: "s1", shapeId: "shape-3", left: 0, top: 0 },
    };
    $("load-selection").click();
    await settle();

    $("insert-new").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    expect(host.calls.updateChart).toHaveLength(0);
  });
});

describe("Load selection", () => {
  beforeEach(bootHostPane);

  it("loads an SSF chart and reveals the in-place edit affordance", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([7, 8, 9]),
      target: { slideId: "s1", shapeId: "shape-1", left: 0, top: 0 },
    };
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("loaded");
    // The "edit selected chart" banner is stale once loaded, so it hides.
    expect($("selection-banner").style.display).toBe("none");
  });

  it("reports a non-SSF-chart selection without touching the deck", async () => {
    host.loadSelectionResult = null;
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("not an ssf chart");
    expect(host.calls.updateChart).toHaveLength(0);
  });
});

describe("Same scale", () => {
  beforeEach(bootHostPane);

  it("pins every value-axis chart in the deck to the union extent", async () => {
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
    ];
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(1);
    const batch = host.calls.updateCharts[0];
    expect(batch).toHaveLength(2);
    // Both charts are re-tagged with the SAME max — the union of the two extents.
    const maxima = batch.map((b) => (JSON.parse(b.opts!.tagData!) as ChartConfig).scale?.max);
    expect(new Set(maxima).size).toBe(1);
    expect(maxima[0]).toBe(90);
  });

  it("refuses to apply with fewer than two value-axis charts", async () => {
    host.deckCharts = [{ configJson: chartJson([1, 2, 3]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } }];
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(0);
    expect($("host-note").textContent?.toLowerCase()).toContain("two");
  });

  /**
   * "The deck now shares one scale" is the whole feature, and it was being
   * printed in green over a scan the code already knew was short.
   *
   * `listChartsInDeck` computes `unread` carefully — it has a comment saying so
   * — and then sent it to a trace and handed back a bare array. Tracing is off
   * by default, so in ordinary use that distinction reached nobody, and the max
   * was taken over whatever slides happened to answer. A real host produced
   * `unread=8 slides=8` and, in the same run, `unread=7 slides=8`.
   */
  it("will not rescale a deck it could not fully read", async () => {
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
    ];
    host.slideCount = 8;
    host.deckScanUnread = 6;
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts, "rescaled a deck it could not see").toHaveLength(0);
    const said = $("host-note").textContent?.toLowerCase() ?? "";
    expect(said, "did not say the scan was short").toContain("whole deck");
    expect($("host-note").className).toContain("status-err");
  });

  /**
   * The count in the success note has to be what the host TOOK.
   *
   * `updateChartsInSlides` drops a chart whose slide or shape the host declines
   * to resolve, and it drops it silently — those are early filters, not throws,
   * so `onFailed` never fires and the pane's `stalled` list stays empty. The
   * note said `parsed.length`, the charts requested. A real host applied the
   * shared scale to one chart in six and the self-test caught it only by
   * re-reading the deck afterwards (`1 of 6 charts carry the shared scale`);
   * this path does no such re-read and would have reported six of six.
   */
  it("reports the charts the host took, not the charts it was asked for", async () => {
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
      { configJson: chartJson([1, 40]), target: { slideId: "s3", shapeId: "c", left: 0, top: 0 } },
    ];
    host.updateChartsDrops = 2; // two charts the host would not resolve, silently
    $("same-scale").click();
    await settle();
    const said = $("host-note").textContent ?? "";
    expect(said, "claimed the whole deck").not.toContain("applied to 3");
    expect(said).toContain("1 of 3");
    expect($("host-note").className, "reported a partial rescale in green").toContain("status-err");
  });

  it("scopes to the selection and guides the user when too few are selected", async () => {
    host.selectionCharts = [
      { configJson: chartJson([1, 2]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
    ];
    $("same-scale-sel").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(0);
    // The selection-scoped guard names Ctrl-click, not the deck message.
    expect($("host-note").textContent?.toLowerCase()).toContain("ctrl-click");
  });

  it("sweeps the debris every stalled redraw left behind, not just the first", async () => {
    // Same Scale deletes each chart's old shapes and redraws them. A stall
    // leaves whatever committed on the slide, and the single-chart path has
    // swept that since updateChartResilient learned to — this path never did.
    // It reported the charts "now empty" while half a chart sat on each one.
    //
    // TWO stalls on purpose. The renderer used to annotate only the FIRST
    // failure with its wreckage and hand `onFailed` the raw error regardless,
    // so a one-stall test would pass against a caller that swept nothing and a
    // renderer that reported nothing. The second chart is what proves both
    // halves are fixed.
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
      { configJson: chartJson([1, 40]), target: { slideId: "s3", shapeId: "c", left: 0, top: 0 } },
    ];
    host.updateChartsStalls.set(0, ["stray-a1", "stray-a2"]);
    host.updateChartsStalls.set(2, ["stray-c1"]);
    $("same-scale").click();
    await settle();
    // Both wrecks cleared, each against its OWN slide.
    expect(host.calls.swept).toEqual([
      { slideId: "s1", ids: ["stray-a1", "stray-a2"] },
      { slideId: "s3", ids: ["stray-c1"] },
    ]);
    // And the user is still told which charts went blank.
    expect($("host-note").textContent?.toLowerCase()).toContain("would not redraw");
  });

  it("sweeps nothing when every chart redraws", async () => {
    // The other half of the contract: a clean run must not go poking at the
    // slide. A sweep here would be deleting shapes that are the new chart.
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
    ];
    $("same-scale").click();
    await settle();
    expect(host.calls.swept).toHaveLength(0);
    expect($("host-note").textContent?.toLowerCase()).toContain("same scale applied");
  });
});

describe("Elements and batch insert", () => {
  beforeEach(bootHostPane);

  // The five element buttons (harvey balls, checkboxes, process flow, KPI row,
  // table) all drop a compact scene at the same fixed offset through the guard.
  for (const id of ["harvey-insert", "check-insert", "flow-insert", "kpi-insert", "table-insert"]) {
    it(`${id} inserts a compact element scene at the element offset`, async () => {
      $(id).click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(1);
      expect(host.calls.insertScene[0].left).toBe(120);
      expect(host.calls.insertScene[0].top).toBe(160);
    });
  }

  it("batch-inserts every config in the JSON box onto the current slide", async () => {
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify([
      { kind: "pie", data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 1] }] } },
      { kind: "stacked", data: { categories: ["A"], series: [{ name: "S", values: [3] }] } },
    ]);
    $("json-insert-batch").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(2);
    expect($("host-note").textContent?.toLowerCase()).toContain("2 chart");
  });
});

describe("guard — busy lockout and error surfacing", () => {
  beforeEach(bootHostPane);

  it("disables the button while the action runs and re-enables it after", async () => {
    // Hold the insert open so the mid-flight state is observable.
    let release!: () => void;
    host.gate = new Promise<void>((r) => (release = r));
    const insertBtn = $<HTMLButtonElement>("insert");

    insertBtn.click();
    await settle();
    expect(insertBtn.disabled).toBe(true); // locked out mid-action

    release();
    await settle();
    expect(insertBtn.disabled).toBe(false); // restored once the action settles
    expect(host.calls.insertScene).toHaveLength(1);
  });

  /**
   * "Done." in green over a chart that is no longer an SSF chart.
   *
   * The renderer has always known — `groupAndTagAll` returns `tagged` — and the
   * insert path discarded its whole return value, so nothing reached the user.
   * The chart is on the slide, clicking it says "the selection is not an
   * SSF chart", and reopening the deck loses the settings permanently. A real
   * host produced it four times in one run.
   */
  it("says when an inserted chart carries no config, instead of Done.", async () => {
    host.insertResult = { slideId: "s1", shapeId: "grp-1", left: 40, top: 50, lost: "no-config" };
    $("insert").click();
    await settle();
    const said = $("host-note").textContent ?? "";
    expect(said, "reported an unusable chart as Done.").not.toBe("Done.");
    expect(said.toLowerCase()).toContain("settings");
    expect($("host-note").className).toContain("status-err");
  });

  /**
   * An update that redrew the chart and lost track of it must not keep the
   * target — `shapeId` names the shape that same update deleted, so the next
   * push resolves a dead id (or a group member, and draws a second chart).
   */
  it("drops the edit target when an update loses track of the chart", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "grp-1", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20, lost: "unknown-shape" };
    $("insert").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase(), "kept quiet about losing the chart").toContain("lost track");
    // The button is back to Insert, so the next press cannot push to a dead id.
    expect($("insert").textContent?.toLowerCase()).not.toContain("update");
  });

  it("surfaces a host failure as a Failed note instead of throwing", async () => {
    host.failInsertOnce = true;
    $("insert").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("failed");
    expect(host.calls.insertScene).toHaveLength(0);
  });

  /**
   * The progress bar exists to say "the host is still working", so it has to
   * stop when the host stops. An insert's last phase note is "Working… done" —
   * still busy — and `guard` used to decide whether to print "Done." by asking
   * whether the note text had changed since it posted "Working…". It had, so
   * "Done." was skipped, nothing ever posted a settled note, and the bar kept
   * its `indeterminate` class: a finished insert under an animation that slid
   * forever. Asserting the bar, not just the words, is the point — the note
   * could read "done" while the strip below it still claimed to be busy.
   */
  it("stops the progress bar when the insert finishes, not just the phase notes", async () => {
    let release!: () => void;
    host.gate = new Promise<void>((r) => (release = r));
    $("insert").click();
    await settle();
    // Mid-flight the bar must actually be running, or the assertion after the
    // release would pass against a bar that never showed at all.
    expect($("status-bar").hasAttribute("hidden")).toBe(false);
    expect($("status-bar").classList.contains("indeterminate")).toBe(true);

    release();
    await settle();
    expect($("host-note").textContent).toBe("Done.");
    expect($("host-note").className).toContain("status-ok");
    expect($("status-bar").hasAttribute("hidden")).toBe(true);
    expect($("status-bar").classList.contains("indeterminate")).toBe(false);
  });

  it("closes out an element insert too — every phase-reporting action settles", async () => {
    $("harvey-insert").click();
    await settle();
    expect($("host-note").textContent).toBe("Done.");
    expect($("status-bar").classList.contains("indeterminate")).toBe(false);
  });

  it("gives an element the same picture rescue a chart gets", async () => {
    /**
     * ELEMENTS ARE WHERE THIS MATTERS MOST, and they never went near
     * `chartPicture`: this path calls `insertSceneIntoSlide` directly. A Harvey
     * ball is a WEDGE at every fraction between 1% and 99% — 0% and 100% are
     * ellipses and survive — so a host without rotation loses exactly the
     * informative range and inserts an empty ring. A table trend cell is an
     * arrowhead that vanishes while its text stays.
     *
     * Shipping the guard on the chart path alone left the two element kinds
     * that need it most dropping their glyphs in silence.
     */
    const raster = stubRaster();
    try {
      host.canPicture = true;
      host.cannotDraw = { what: ["a pie slice"], nodes: 1 };
      $("harvey-insert").click();
      await settle();
      expect(
        (host.calls.insertScene.at(-1) as { pictureBase64?: string }).pictureBase64,
        "an element this host cannot draw went in as shapes",
      ).toBeTruthy();
      expect(String($("host-note").textContent), "inserted a picture without saying why").toMatch(/a pie slice/);
    } finally {
      raster.restore();
    }
  });

  /**
   * The flip side: an action that DID report an end state keeps it. "Done." is
   * the fallback for silence, not an overwrite — a settlement counted per
   * action is what keeps those two apart.
   */
  it("leaves an action's own closing note alone instead of overwriting it with Done.", async () => {
    host.loadSelectionResult = null;
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent).not.toBe("Done.");
    expect($("host-note").textContent?.toLowerCase()).toContain("not an ssf chart");
    expect($("status-bar").classList.contains("indeterminate")).toBe(false);
  });
});

/**
 * THE AGENDA BUTTON'S OWN "Done." OVER NOTHING.
 *
 * `guard` closes an action that posted no end state of its own by printing
 * "Done." in green — the fallback `settledNotes` exists to make reliable. The
 * agenda handler answered an empty chapters box with a bare `return`, so the
 * fallback fired and the pane reported success for zero slides: the exact
 * "cannot tell did-the-thing from did-nothing" verdict this file has spent the
 * most effort on, on the one tab where it was still reachable.
 *
 * BOTH HALVES ARE ASSERTED TOGETHER, and neither is enough alone. The note
 * without the call count passes against a handler that says the right thing and
 * inserts anyway; the call count without the note is exactly what the bug did.
 */
describe("Insert agenda slides", () => {
  beforeEach(bootHostPane);

  const typeChapters = (text: string) => {
    ($("agenda-chapters") as HTMLTextAreaElement).value = text;
  };

  it("says the chapters box is empty instead of settling green over nothing", async () => {
    typeChapters("");
    $("agenda-insert").click();
    await settle();
    expect(host.agendaSlides, "inserted slides for an empty chapters box").toHaveLength(0);
    expect($("host-note").textContent, "reported an insert that never happened as Done.").not.toBe("Done.");
    expect($("host-note").className, "a no-op settled green").toContain("status-err");
    expect(String($("host-note").textContent).toLowerCase()).toContain("chapters box is empty");
  });

  it("says the same for a box holding nothing but blank lines", async () => {
    // The route a user actually takes to this branch: a stray newline or an
    // indented line left behind. `agendaChapters` trims and drops empties, so
    // the handler sees the same zero as an untouched box — and a pane that
    // answered one honestly and the other with "Done." would be worse than both.
    typeChapters("\n   \n\t\n");
    $("agenda-insert").click();
    await settle();
    expect(host.agendaSlides, "inserted slides for a box of blank lines").toHaveLength(0);
    expect($("host-note").className, "a no-op settled green").toContain("status-err");
  });

  it("still inserts one slide per chapter, and settles, when there are chapters", async () => {
    typeChapters("Introduction\nMarket overview\nStrategy");
    $("agenda-insert").click();
    await settle();
    expect(host.agendaSlides, "the real insert stopped happening").toHaveLength(1);
    expect(host.agendaSlides[0], "one agenda slide per chapter").toHaveLength(3);
    // "Done." is right HERE — work landed and the handler posts no note of its
    // own, which is precisely the case `guard`'s fallback is for.
    expect($("host-note").textContent).toBe("Done.");
    expect($("host-note").className).toContain("status-ok");
  });
});

describe("demo-insert results pages", () => {
  beforeEach(bootHostPane);

  it("attempts each results page in its own insertDemoDeck call — a failing page does not drop the rest", async () => {
    // Presentation_3.pptx: one failing results page threw insertDemoDeck's
    // "everything failed" guard and NO pages landed. Fix: pane loops over
    // pages, per-page try/catch, so page 2 still gets attempted after page 1
    // throws.
    //
    // With the fake mock returning "failed" for every item and demoItems()'s
    // 32-ish failure-count on the full deck, buildResultsScenes emits multiple
    // pages (≥2). Fail the FIRST results page's insertDemoDeck call; the
    // pane must still call insertDemoDeck at least once more for page 2.
    host.demoDeckPageFailures.add(2); // main deck = call #1, results page 1 = call #2
    $("demo-insert").click();
    await settle();
    // At least 3 calls: main deck + results page 1 (thrown) + results page 2.
    // If the pane batched all results into one call (the old behavior), only
    // 2 calls would happen and the failure at #2 would drop the rest.
    expect(host.demoRuns).toBeGreaterThanOrEqual(3);
    // Each results-page call carries exactly ONE item (per-page loop), NOT
    // the whole batch — the load-bearing property of the fix.
    const resultsCalls = host.demoDeckCalls.slice(1); // drop the main deck call
    for (const call of resultsCalls) expect(call).toHaveLength(1);
    // The pane's note reflects that some pages landed and some did not.
    const noteText = $("host-note").textContent ?? "";
    expect(noteText).toMatch(/results page/i);
  });

  it("reconciles each results page, so a half-landed one cannot be orphaned", async () => {
    // The results insert runs at the worst possible moment — right after a run
    // that has just finished exhausting the host — and used to get none of the
    // run's own protections. A page whose add landed but whose shapes did not
    // left a stamped, untagged slide at the END of the deck that nothing ever
    // cleaned: the main run's repair had already finished, and its range
    // stopped short of that slide. A real 38-item web run ended exactly so.
    $("demo-insert").click();
    await settle();
    const resultsOpts = host.demoDeckOpts.slice(1); // drop the main deck call
    expect(resultsOpts.length).toBeGreaterThan(0);
    for (const o of resultsOpts) expect(o?.reconcile).toBe(true);
  });

  it("writes the run log after the results pages, not before them", async () => {
    // Taken before, the log's trace ended at the repair read — so when a
    // results page then failed, the file said "(results slide not added)" with
    // nothing in it about why. The log has to outlast the run it describes.
    let release!: () => void;
    host.demoDeckGate = new Promise<void>((r) => (release = r));
    host.demoDeckGateOn = 2; // the first results page
    $("demo-insert").click();
    await settle();
    // Mid-run, held inside the results insert: no log yet.
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(true);
    release();
    host.demoDeckGate = null;
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false);
  });

  it("all results pages land ⇒ no 'results slide not added' warning", async () => {
    $("demo-insert").click();
    await settle();
    const noteText = $("host-note").textContent ?? "";
    expect(noteText).not.toMatch(/results slide not added/i);
    // At least the main deck + at least one results page.
    expect(host.demoRuns).toBeGreaterThanOrEqual(2);
  });
});

describe("reporting what a repair actually removed", () => {
  beforeEach(bootHostPane);

  it("counts duplicates and empty strays apart", async () => {
    // A real run said "1 duplicate slide removed · 8 orphan slides" and then,
    // one sentence later, "Repaired 9 duplicate slide(s)". Both numbers came
    // from the same pass; only the second was wrong.
    host.demoDeckStatusOverride = "rendered";
    host.demoReconcile = {
      snapshots: [],
      plan: {
        actions: [
          { kind: "delete", index: 5, slot: 3, reason: "dup" },
          { kind: "delete", index: 4, slot: null, reason: "empty" },
          { kind: "delete", index: 2, slot: null, reason: "empty" },
        ],
        verdicts: [],
        orphans: [],
        summary: {
          items: 1,
          rendered: 1,
          partial: 0,
          lost: 0,
          skipped: 0,
          wreckage: 0,
          empty: 0,
          duplicates: 1,
          falseStamps: 0,
          untagged: 0,
          orphans: 2,
        },
      },
      applied: { unstamped: 0, regrouped: 0, deleted: 3 },
      refused: 0,
    };
    $("demo-insert").click();
    await settle();
    const text = $("host-note").textContent ?? "";
    expect(text).toMatch(/removed 3 slide\(s\) \(1 duplicate, 2 empty\)/);
    expect(text).not.toMatch(/3 duplicate slide/);
  });
});

describe("updating a chart the live canvas will not redraw", () => {
  beforeEach(bootHostPane);

  /** Load a chart so Insert becomes Update, then push an edit. */
  async function loadThenUpdate() {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    $("insert").click();
    await settle();
  }

  it("does not start a second write to the same chart while one is in flight", async () => {
    // The auto-update timer calls doInsert DIRECTLY, so it never sees the
    // disabled buttons a click would. One resilient update can legitimately
    // take tens of seconds — a 45s stall, then a slide swap, then a bounded
    // raster — and a user editing through that could start a second update
    // against the SAME stale target. Whichever finished last won; if that was
    // the one started from the already-superseded target, the pane reported
    // "that chart is no longer on the slide" about a chart just written fine.
    vi.useFakeTimers();
    try {
      host.loadSelectionResult = {
        configJson: chartJson([1, 2, 3]),
        target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
      };
      host.updateResult = { slideId: "s1", shapeId: "shape-10", left: 10, top: 20 };
      $("load-selection").click();
      await vi.advanceTimersByTimeAsync(10);
      ($("auto-update") as HTMLInputElement).checked = true;

      // Hold the first update open, then let the debounce fire underneath it.
      let release!: () => void;
      host.updateGate = new Promise<void>((r) => (release = r));
      $("insert").click();
      await vi.advanceTimersByTimeAsync(10);
      const duringFirst = host.calls.updateChart.length;
      // Two edits land while the first write is still open.
      ($("chart-title") as HTMLInputElement).value = "edited";
      $("chart-title").dispatchEvent(new Event("input", { bubbles: true }));
      await vi.advanceTimersByTimeAsync(3000);
      // Nothing new was issued — the timer re-armed instead of racing.
      expect(host.calls.updateChart).toHaveLength(duringFirst);
      release();
      host.updateGate = null;
      await vi.advanceTimersByTimeAsync(3000);
      // …and once the first finished, the queued edit did go out.
      expect(host.calls.updateChart.length).toBeGreaterThan(duringFirst);
    } finally {
      host.updateGate = null;
      vi.useRealTimers();
    }
  });

  it("redraws with the slide deselected, at the off-screen batch size", async () => {
    // The whole point: a redraw is the add-in's worst case on the web, and it
    // is only bad because the slide is on screen. Looking away is free.
    await loadThenUpdate();
    expect(host.calls.deselected).toEqual([["s1"]]);
  });

  it("swaps the slide when the redraw stalls and the slide holds only the chart", async () => {
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.slideHoldsOnlyChart = true;
    host.swapOutcome = "swapped";
    await loadThenUpdate();
    expect(host.calls.insertFile).toHaveLength(0); // a slide swap, not a deck insert
    expect($("host-note").textContent).toMatch(/rebuilt that slide/i);
    // And it says what the rebuild COST. A swap replaces the slide, so the
    // chart comes back and anything on it that is not a shape does not:
    // speaker notes, transition, animations. The guard in front of the swap
    // can only see shapes — Office.js exposes no way to read notes at all
    // (office-js#3269) — so the add-in cannot ask first and cannot avoid it.
    // Telling the user afterwards is the difference between a loss they can
    // undo and one they discover in front of an audience.
    expect($("host-note").textContent, "a swap discarded the slide's notes without saying so").toMatch(
      /speaker notes|replaced/i,
    );
  });

  it("stops, and says so, when the swap left the original slide behind", async () => {
    // The swap inserts the new slide and THEN deletes the old one. When only
    // the delete fails, the chart is on two slides — and answering the caller
    // the same "false" as a swap that did nothing sent it on to the picture
    // layer, which rasterized the chart onto the surviving original. One
    // stall, the chart three times over, and a message saying it went fine.
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.slideHoldsOnlyChart = true;
    host.swapOutcome = "duplicated";
    await loadThenUpdate();
    // No picture update was attempted on the original.
    expect(host.calls.updateChart.filter((c) => c.opts.pictureBase64)).toHaveLength(0);
    expect($("host-note").textContent).toMatch(/two slides/i);
  });

  it("gives up on a canvas that never answers instead of hanging the pane", async () => {
    // jsdom never fires `img.onload`, which is exactly the failure being
    // guarded: a wedged image decode. Unbounded, the await never settles —
    // the button stays disabled, no error, no timeout, nothing to do but
    // reload the pane. Same Scale awaits one of these PER CHART, so one
    // wedged decode froze the whole deck-wide operation.
    vi.useFakeTimers();
    try {
      host.autoPicture = true; // the too-dense branch, which rasterizes
      $("insert").click();
      await vi.advanceTimersByTimeAsync(30_000);
      // The action finished: the chart went in as shapes, and the pane says why.
      expect(($("insert") as HTMLButtonElement).disabled).toBe(false);
      expect(host.calls.insertScene).toHaveLength(1);
      expect($("host-note").textContent).toMatch(/could not be turned into a picture/i);
    } finally {
      vi.useRealTimers();
    }
  });

  it("will not swap a slide that holds anything besides the chart", async () => {
    // The replacement is a NEW slide: notes, transitions and any other shape
    // on the old one do not come with it. So it is only ever offered where
    // there is demonstrably nothing else to lose.
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.swapOutcome = "swapped"; // the swap WOULD work — the guard is the only thing stopping it
    host.slideHoldsOnlyChart = false;
    await loadThenUpdate();
    expect($("host-note").textContent).not.toMatch(/rebuilt that slide/i);
  });

  /**
   * The floor of the ladder, on the failure it exists for.
   *
   * Layer 1 deletes the old chart, COMMITS that, and only then draws — so a
   * stall leaves half a chart on the slide and no shape at the target's id.
   * Every layer below assumed the chart was still there: layer 2 checks the
   * slide holds nothing but the chart and saw the litter, layer 3 re-resolved
   * the id layer 1 had just destroyed, got nothing back, and the pane announced
   * "that chart is no longer on the slide" over a half-drawn one. Three
   * fallbacks, all disabled by the damage they were meant to repair.
   */
  it("sweeps the wreckage and draws the chart back, rather than calling it gone", async () => {
    const raster = stubRaster();
    try {
      host.updateChartThrows = true; // layer 1 stalls — after the delete committed
      host.slideHoldsOnlyChart = false; // layer 2 out of the picture, so this is the floor
      host.insertResult = { slideId: "s1", shapeId: "grp-new", left: 40, top: 50 };
      await loadThenUpdate();

      // The half-drawn chart went first — otherwise the replacement lands on
      // top of it and the user keeps both.
      expect(host.calls.swept).toEqual([{ slideId: "s1", ids: ["stray-1", "stray-2"] }]);

      // Then the chart was drawn back: onto the slide it came from, at the
      // position it held, as one picture no live canvas can stall on.
      const drawn = host.calls.insertScene.at(-1);
      expect(drawn).toMatchObject({ slideId: "s1", left: 40, top: 50 });
      expect(drawn?.pictureBase64).toBeTruthy();

      // And NOT as a picture update against the shape layer 1 deleted — that
      // call can only ever resolve nothing.
      expect(host.calls.updateChart.filter((c) => c.opts.pictureBase64)).toHaveLength(0);
      expect($("host-note").textContent).not.toMatch(/no longer on the slide/i);

      /**
       * AND IT SAYS THE CHART WAS DESTROYED AND REBUILT — which it could not,
       * until 2026-09-04.
       *
       * `updateChartResilient` returns `{ next, picture, recovered }` together,
       * and the ladder that read them returned on the first match. `next` came
       * back here, so `recovered` was unreachable and the user was told only
       * "Drawn as a picture" — nothing about the in-place update having deleted
       * their chart and the pane having drawn it again from scratch.
       *
       * That is the fact a user would most want about their own document, and
       * it was the one fact this path could not deliver.
       */
      const said = $("host-note").textContent ?? "";
      expect(said, `never said the chart had been redrawn from scratch: ${said}`).toMatch(/redrawn from scratch/i);
      expect($("host-note").className, "reported a destroyed-and-rebuilt chart as success").toMatch(/status-err/);
    } finally {
      raster.restore();
    }
  });

  it("says the chart was rebuilt AND why it is a picture, not one or the other", async () => {
    /**
     * The update ladder used to pick one fact with an `else if`, so a
     * too-dense chart that was destroyed and rebuilt reported the rebuild and
     * swallowed `chartPicture`'s warn — the sentence that explains why it is a
     * picture and what can still be done with it. Both are true at once and
     * both are the user's business.
     */
    const raster = stubRaster();
    try {
      /**
       * A warn WITHOUT a png, deliberately. `autoPicture` would also produce a
       * warn, but it supplies a picture with it — and that picture sends the
       * resilient ladder down a different route, so `next` comes back null and
       * this branch never runs. A host that cannot insert pictures at all warns
       * and hands over nothing, which is the combination this needs.
       */
      host.canPicture = false;
      host.updateChartThrows = true; // wreckage, so the chart is rebuilt
      host.slideHoldsOnlyChart = false;
      host.insertResult = { slideId: "s1", shapeId: "grp-new", left: 40, top: 50 };
      await loadThenUpdate();

      const said = $("host-note").textContent ?? "";
      expect(said, `never said the chart was rebuilt: ${said}`).toMatch(/redrawn from scratch/i);
      expect(said, `swallowed the second fact: ${said}`).toMatch(/drawn as a picture/i);
      // REBUILT LEADS. What happened to their document outranks how it was
      // drawn, and the old ladder reported only the second of these.
      expect(said.indexOf("redrawn from scratch"), `led with the wrong fact: ${said}`).toBeLessThan(
        said.indexOf("Drawn as a picture"),
      );
    } finally {
      host.canPicture = true;
      raster.restore();
    }
  });

  it("keeps the chart editable after that recovery, so the next edit still lands", async () => {
    const raster = stubRaster();
    try {
      host.updateChartThrows = true;
      host.slideHoldsOnlyChart = false;
      host.insertResult = { slideId: "s1", shapeId: "grp-new", left: 40, top: 50 };
      await loadThenUpdate();
      // The button still reads Update, against the shape the recovery created —
      // not the dead one, and not "Insert into slide".
      expect(($("insert") as HTMLButtonElement).textContent).toMatch(/update/i);
      host.updateChartThrows = false;
      host.updateResult = { slideId: "s1", shapeId: "grp-newer", left: 40, top: 50 };
      $("insert").click();
      await settle();
      expect(host.calls.updateChart.at(-1)?.target).toMatchObject({ shapeId: "grp-new" });
    } finally {
      raster.restore();
    }
  });

  it("still says 'gone' when the chart really is gone — nothing was destroyed to recover", async () => {
    // The other side of the same coin: an update that found nothing to replace
    // destroyed nothing, so there is no wreckage, no sweep, and no chart to
    // draw back. Telling this user to insert again is right; doing it for the
    // case above would have given them the chart twice.
    host.updateResult = undefined; // resolved no shape — the user deleted it
    await loadThenUpdate();
    expect(host.calls.swept).toHaveLength(0);
    expect(host.calls.insertScene).toHaveLength(0);
    expect($("host-note").textContent).toMatch(/no longer on the slide/i);
  });
});

describe("demo-insert one-shot deck insert", () => {
  beforeEach(bootHostPane);

  it("hands the host a generated file and never draws the deck shape by shape", async () => {
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.calls.insertFile[0].b64).toMatch(/^UEsDBBQA/);
    // The whole point: not one slide is drawn through the shape renderer.
    expect(host.demoRuns).toBe(0);
    expect($("host-note").textContent).toMatch(/one file/i);
  });

  it("says WHY when it could not verify, instead of quietly reporting a raw count", async () => {
    // A real run reported "Inserted 12 of 12 slides as one file" — the
    // fallback wording — because the verification pass returned null and the
    // pane shrugged. Three different failures produced that same null and the
    // message named none of them.
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    // The fake reads no slot tags, so the pass cannot identify a single slide.
    expect($("host-note").textContent).toMatch(/not verified:/i);
    expect($("host-note").textContent).toMatch(/slot tag|no slides/i);
  });

  it("falls back to shapes when the host has no insertSlidesFromBase64", async () => {
    host.canInsertFile = false;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("falls back when the file could not be built at all", async () => {
    host.canInsertFile = true;
    host.buildFileError = new Error("pptxgenjs blew up");
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("falls back when the insert threw and nothing landed", async () => {
    host.canInsertFile = true;
    host.insertFileError = new Error("host refused the deck");
    host.slideCount = 1; // unchanged before and after — nothing landed
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("does NOT fall back after a partial insert — that would draw the deck twice", async () => {
    // The failure this exists to prevent: some slides landed, the pane decides
    // the fast path "failed", and the shape renderer then appends the whole
    // deck a second time on top of them.
    host.canInsertFile = true;
    host.insertFileLands = 3;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBe(0);
    expect($("host-note").textContent).toMatch(/host took 3 of/i);
  });

  /**
   * The run token has to reach the FILE, not just the run.
   *
   * It is the join key between the downloaded log and the .pptx the run
   * produced, and every diagnosis in this project's history is that join. Done
   * by hand it started with a guess at which slides were the run's, because a
   * deck holds whatever earlier runs left in it — one real file carried 30
   * slides from the run under investigation and one from another. A log
   * missing the token sends `npm run triage` straight back to guessing.
   */
  it("puts the run's identity in the log the fast path writes", async () => {
    const dl = captureDownloads();
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs).toHaveLength(1);
    expect(log.runs[0].run).toBe("run-under-test");
    // …and which items were ever meant to carry a config. Without this the
    // title and contents pages — `PowerChart` objects with no config by
    // design — read as charts that lost their tag, and a clean run triages as
    // a broken one. Seven false alarms on the first real deck this was run on.
    const items = log.runs[0].items;
    const chartFlags: boolean[] = items.map((i: { chart: boolean }) => i.chart);
    expect(chartFlags).toContain(true);
    expect(chartFlags).toContain(false);
    expect(items.find((i: { title: string }) => i.title === "Title").chart).toBe(false);
    dl.restore();
  });

  it("puts the run's identity in the log the shape path writes", async () => {
    // A different token from the fast path's, and taken from the renderer's
    // own report rather than minted again by the pane — the pane minting its
    // own would produce a log that joins to nothing in the deck.
    const dl = captureDownloads();
    host.canInsertFile = false;
    $("demo-insert").click();
    await settle();
    $("demo-log").click();
    expect((await dl.lastJson()).runs[0].run).toBe("fake-run-1");
    dl.restore();
  });

  it("leaves a downloadable run log behind — including when the run went badly", async () => {
    // The whole point of the log is diagnosing real host failures after the
    // fact, and the fast path is what a real run TAKES: the checkbox ships
    // checked and every current host advertises insertSlidesFromBase64. Yet
    // this path returned before `lastRunLog` was ever set, so the button that
    // downloads it stayed disabled — success or failure. The failing run is
    // exactly the one worth having a file for.
    host.canInsertFile = true;
    host.insertFileLands = 3; // a partial insert: the case most worth logging
    $("demo-insert").click();
    await settle();
    expect($("host-note").textContent).toMatch(/host took 3 of/i);
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false);
  });

  it("does not offer an older run's log after a run that produced none", async () => {
    // The button used to stay enabled from a previous run, handing out a file
    // about a completely different insert with nothing on screen to say so.
    host.canInsertFile = false;
    $("demo-insert").click();
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false); // shape path logs
    // Now a run that produces no log of its own: the deck build fails, and the
    // shape path it falls back to throws on its very first call.
    host.canInsertFile = true;
    host.buildFileError = new Error("pptxgenjs blew up");
    host.demoRuns = 0; // the failure indices below count from this run
    host.demoDeckPageFailures = new Set([1]);
    $("demo-insert").click();
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(true);
  });

  it("does not redraw the whole deck when it could not measure what landed", async () => {
    // The insert throws, and the rescue count read fails WITH it — the same
    // wedged host, milliseconds apart. Treating that as "nothing landed" sent
    // the shape renderer off to draw all 38 slides a second time, on top of
    // however many did arrive. When it cannot tell, it must not guess.
    host.canInsertFile = true;
    host.insertFileError = new Error("host refused the deck");
    host.slideCountThrowsAfter = 1; // the opening read answers; the rescue does not
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBe(0); // the deck was NOT drawn again
    expect($("host-note").textContent).toMatch(/twice/i);
  });

  it("appends the generated deck instead of letting the host front it", async () => {
    // The first real run put 37 generated slides AHEAD of the user's own title
    // slide, because insertSlidesFromBase64 inserts at the front unless it is
    // given a target. The pane must ask for the tail.
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.calls.insertFile[0].expected).toBeGreaterThan(0);
  });

  it("respects the fast-path opt-out", async () => {
    host.canInsertFile = true;
    ($("demo-path") as HTMLSelectElement).value = "shapes";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("takes both paths on one click, and logs them as two runs", async () => {
    // Every renderer change touches both paths, and they fail in completely
    // different ways — so testing them meant two runs an hour apart with a
    // deploy in between. One click, one deck, one log, same session.
    const dl = captureDownloads();
    host.canInsertFile = true;
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs.map((r: { path: string }) => r.path)).toEqual(["file", "shapes"]);
    // Separate identities, or the deck they share cannot be read apart.
    expect(log.runs[0].run).not.toBe(log.runs[1].run);
    dl.restore();
  });

  /**
   * One click, one file.
   *
   * A round used to be three clicks producing three downloads, uploaded
   * separately and joined at the other end — and most probe runs establish
   * nothing, so a good share of that traffic existed to learn that the answers
   * had not changed.
   *
   * The demo deck is deliberately NOT chained in: its two halves have to run on
   * different decks, and a button cannot open a fresh one.
   */
  it("says how loaded the deck was before it starts", async () => {
    // A round that dies leaves only its steps, and on 2026-08-09 one died
    // sixteen seconds in — two 8-second stalls and then the tab — with nothing
    // in the file to say whether PowerPoint had been fresh or already carrying
    // a self-test's worth of shapes. This project has documented since
    // 2026-08-06 that heavy work on an already-large deck is what kills the
    // tab, and `docs/PUBLISHING.md` splits the demo halves across two decks for
    // exactly that reason. Both readings fitted the crash file and neither
    // could be checked.
    const { setTracing, traceLog } = await import("../src/core/trace");
    host.deckSlideIds = ["a", "b", "c"];
    setTracing(true);
    try {
      $("demo-round").click();
      await settle();
      const said = traceLog().entries.filter((e) => e.message === "round starting");
      expect(said, "a round that dies cannot say what deck it died on").toHaveLength(1);
      // The number, not merely the line. A line carrying nothing would satisfy
      // a check for its own existence and answer the question no better.
      expect(said[0].data?.deckSlides).toBe(3);
    } finally {
      setTracing(false);
      host.deckSlideIds = undefined;
    }
  });

  it("records a stack when the round dies in OUR code, not the host's", async () => {
    /**
     * A HOST REFUSAL NAMES ITSELF; A BUG IN THIS PANE DOES NOT.
     *
     * `errorText` digs an Office.js code and `debugInfo` out of a RichApi
     * error, and the trace line before it says which call was in flight. When
     * the fault is our own, the record carries a message and nothing else.
     *
     * 2026-09-02 is what that costs. Round 360's second attempt died at 83s
     * with `Cannot read properties of undefined (reading 'answer')` — the first
     * occurrence in 161 crash records, on a host so sick that `deckSlides` was
     * unreadable from the first line. Reading the code ruled out the three
     * obvious sites and named none: `withTimeout` rejects rather than resolving
     * empty, `withProbeContext` always returns `fn(...)` or throws, and the
     * probe in flight returns on every path. There are 49 `.answer` reads in
     * `host-probe.ts` and nothing to say which one it was. It is still
     * unlocated.
     *
     * One frame would have settled it. This does not fix the fault — it makes
     * the next occurrence the last one anybody has to guess at.
     */
    const { setTracing, traceLog } = await import("../src/core/trace");
    // The round's first statement throws a native Error: our own code failing,
    // not the host refusing at a call the trace already names.
    host.syncResetThrows = true;
    setTracing(true);
    try {
      $("demo-round").click();
      await settle();
      const failed = traceLog().entries.filter((e) => e.message === "action failed");
      expect(failed.length, "a round died and nothing traced the failure").toBeGreaterThan(0);
      const stack = failed[0].data?.stack;
      expect(stack, "the round died in our own code and recorded no frame").toBeTruthy();
      /**
       * THE FRAMES, not merely a truthy field. `err.stack` opens with the
       * message, so a stack trimmed to its first line is a field that exists
       * and names nothing — and the first version of this assertion, matching
       * `/\bat\b/` on the joined string, PASSED against exactly that mutant.
       * The message it is built from ("the pane's own code fell over") has no
       * standalone "at" in it, so the regex was matching a frame that a
       * one-line slice would still have contained: `String(stack)` of a
       * single-element join is the message alone, and the match came from the
       * word boundary landing inside the file path Vitest appends. Either way
       * it did not test the property it claimed.
       *
       * Counting segments does. The handler joins frames with " | ", so a slice
       * of one produces no separator at all, and a stack worth recording has at
       * least the throw site plus one caller.
       */
      const frames = String(stack).split(" | ");
      expect(frames.length, "the stack was trimmed to its message and names no call site").toBeGreaterThanOrEqual(2);
      expect(frames.slice(1).join(" "), "the frames carry no call site").toMatch(/\bat\b/);
    } finally {
      setTracing(false);
      host.syncResetThrows = false;
    }
  });

  it("puts the environment INTO the round file, not into a window it slices off", async () => {
    // ZERO OF 69 ARCHIVED ROUNDS carry a `host`/`environment` line, or a
    // `slide size` line, on an archive where every single build contains the
    // code that emits them. `traceEnvironment` was called at pane-wiring time,
    // and the round's `traceLog(traceFrom)` slices off everything traced before
    // its mark — so the host, the platform, the Office version and the slide
    // size have never once reached a round file.
    //
    // The pre-mark window is not small: round 092's first entry sits at
    // ms 48245 with `dropped: 0`, so 48 seconds of trace was cut away.
    //
    // Asserted against the DOWNLOADED BUNDLE rather than `traceLog()`, because
    // the whole defect is the difference between the two. A test reading the
    // live buffer would have passed for all 69 rounds.
    const dl = captureDownloads();
    const { setTracing, trace } = await import("../src/core/trace");
    traceRef.fn = trace;
    setTracing(true);
    try {
      $("demo-round").click();
      await settle();
      const bundle = (await dl.lastJson()) as { trace?: { entries?: { scope?: string; message?: string }[] } };
      const entries = bundle?.trace?.entries ?? [];
      expect(
        entries.filter((e) => e.scope === "host" && e.message === "environment"),
        "the round file cannot say which host it ran on",
      ).toHaveLength(1);
    } finally {
      setTracing(false);
    }
  });

  it("runs the probe and the self-test on one click, and saves them as one file", async () => {
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    expect(bundle.hostAnswers?.answers?.length, "the round's file carries no probe answers").toBeGreaterThan(0);
    expect(bundle.selftest?.length, "the round's file carries no self-test verdicts").toBeGreaterThan(0);
    // One file, not two. The whole point is that there is one thing to send.
    expect(dl.count()).toBe(1);
    dl.restore();
  });

  /**
   * The two uploads that used to be a person's job.
   *
   * Every diagnosis in this project's history has taken three things — the run
   * log, the deck, and a screenshot — and the owner has been producing all
   * three by hand, once per round. Two of them the add-in can read for itself.
   */
  it("carries what landed on the slides, and pictures of the slides it added", async () => {
    host.deckSlideIds = ["s1"];
    host.deckInventory = [
      { slideId: "s1", index: 0, shapes: [{ id: "sh1", name: "PowerChart" }] },
      { slideId: "s2", index: 1, shapes: [{ id: "sh2", name: "bar 1" }] },
    ];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    // Shapes that are NOT charts are the point. The scan has always had them
    // and always dropped them, and they are what "41 shapes became 79" and
    // "the slide still holds what was there before" are questions about.
    expect(bundle.deck?.inventory?.[1]?.shapes?.[0]?.name).toBe("bar 1");
    dl.restore();
  });

  it("says what it left in the deck, and points at the button that clears it", async () => {
    // A round leaves its slides on purpose — they ARE the evidence, and
    // `docs/REGRESSION.md` is written around a deck someone can open. What it
    // never did was say so. The 2026-08-09 evening round added 43 slides, 36 of
    // which read back empty, reported "Saved as one file", and left the owner
    // to discover a 44-slide deck by opening it. `Clean up the last round` had
    // existed the whole time and nothing named it at the moment it mattered.
    //
    // The same gap cost a round at the other end of this one: the host probe
    // left 21 blank slides and reported nothing, which is why it carries a row
    // in the answer sheet now. The self-test never got the equivalent.
    host.deckSlideIds = ["s1"];
    host.roundAddsSlides = ["added-full", "added-empty"];
    // `s1` is EMPTY on purpose, and it is the whole point of the fixture: it is
    // the user's own blank slide, sitting in the deck before the round started.
    // With it full, dropping the added-slides filter changed no number and the
    // assertion below passed against the bug it names.
    host.deckInventory = [
      { slideId: "s1", index: 0, shapes: [] },
      { slideId: "added-full", index: 1, shapes: [{ id: "sh2", name: "bar 1" }] },
      { slideId: "added-empty", index: 2, shapes: [] },
    ];
    $("demo-round").click();
    await settle();
    const said = $("host-note").textContent ?? "";
    expect(said, "the round never said it left anything behind").toMatch(/left 2 slides in this deck/);
    // ONE of the two, and not the slide that was already there: an empty count
    // that swept in the deck the round landed in would report a user's own
    // blank slides as our litter.
    expect(said, "counted the wrong slides as empty").toMatch(/1 of which read back empty/);
    expect(said, "named no way to clear them").toContain("Clean up the last round");
  });

  it("photographs only the slides the round added, never the whole deck", async () => {
    // A picture of a forty-slide deck is mostly slides nobody touched. The id
    // diff is what makes the pictures worth the bytes — and ids rather than
    // counts, because a deck's own id list is the stronger question about it.
    host.deckSlideIds = ["s1"];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const before = await dl.lastJson();
    expect(before.deck?.newSlides, "photographed a slide that was there before the round").not.toContain("s1");
    dl.restore();
  });

  it("does not claim it added the whole deck when the before-list could not be read", async () => {
    // THE FALLBACK WAS `idsAfter`, which says every slide the deck holds was
    // added by this round. A round that added four slides to the owner's
    // forty-slide deck reported forty — and `describeLitter` then told them it
    // had left forty behind, some of which "read back empty". Those were their
    // own blank slides. `triage.mjs` builds its added-set from this field too.
    //
    // Unknown is its own answer, the same lesson as reading a host's silence as
    // a "no": `newSlides` stays empty and `beforeUnknown` says why.
    host.deckSlideIds = ["s1", "s2", "s3"];
    host.deckSlideIdsBeforeFails = true;
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    expect(bundle.deck?.newSlides, "claimed the user's own slides as this round's litter").toEqual([]);
    expect(
      bundle.deck?.beforeUnknown,
      "an empty list that cannot say why is indistinguishable from adding nothing",
    ).toBe(true);
    dl.restore();
  });

  /**
   * Putting the deck back.
   *
   * A round leaves slides behind on purpose, and clearing them has been a manual
   * chore once per round — in a deck that also grows and skews the next round's
   * timings. What makes it safe to automate is that it deletes an ID LIST the
   * round watched appear, never a rule for recognising a test slide: a rule can
   * match a slide the owner made, an id list cannot.
   */
  it("cleans up exactly the slides the round added, and nothing that was there before", async () => {
    host.deckSlideIds = ["s1"];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    dl.restore();
    const added = bundle.deck.newSlides as string[];
    expect(added.length, "the round added no slides, so this proves nothing").toBeGreaterThan(0);
    $("demo-tidy").click();
    await settle();
    expect(host.deletedSlides).toEqual(added);
    expect(host.deletedSlides, "deleted a slide that was in the deck before the round").not.toContain("s1");
  });

  it("stays disabled until a round has named something to clean up", async () => {
    // With no id list there is nothing to delete but a guess, and guessing about
    // deletion in someone's own deck is not a trade this pane makes.
    expect(($("demo-tidy") as HTMLButtonElement).disabled).toBe(true);
    $("demo-tidy").click();
    await settle();
    expect(host.deletedSlides).toEqual([]);
  });

  it("says how many the host refused rather than reporting a clean sweep", async () => {
    // `deleteSlideById` has a whole comment about why a host saying "gone" is
    // not proof. A cleanup that reported success for slides still sitting in the
    // deck would be the same class of claim.
    host.deckSlideIds = ["s1"];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    dl.restore();
    host.refuseSlideDelete = [(bundle.deck.newSlides as string[])[0]];
    $("demo-tidy").click();
    await settle();
    expect($("host-note").textContent ?? "").toMatch(/Removed \d+ of \d+/);

    // And it STAYS dead. The handler empties `tidyable` and disables itself on
    // purpose — after a partial sweep there is no longer a list it trusts — but
    // `guard()` captured the clicked button before running the handler and its
    // finally brought it back. A second press then printed a green
    // "Cleaned up — 0 slide(s) removed." over the slides the host had just
    // refused, which docs/MANUAL.md and the button's own title both deny.
    expect(($("demo-tidy") as HTMLButtonElement).disabled, "guard resurrected a button meant to stay dead").toBe(true);
    const deletedSoFar = [...host.deletedSlides];
    $("demo-tidy").click();
    await settle();
    expect(host.deletedSlides, "a second press asked the host again").toEqual(deletedSoFar);
    expect($("host-note").textContent ?? "").not.toMatch(/Cleaned up/);
  });

  it("still writes the file when the host will not describe its own deck", async () => {
    // The tail of a round runs on a host that has just been through the
    // self-test, and may well be the reason the round is worth reading. Losing
    // the verdicts to the diagnostic's own evidence-gathering would be the
    // worst trade in the file.
    host.deckScanThrows = true;
    host.canRaster = false;
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    expect(bundle.selftest?.length, "lost the verdicts because the deck would not answer").toBeGreaterThan(0);
    expect(bundle.hostAnswers?.answers?.length).toBeGreaterThan(0);
    dl.restore();
  });

  it("banks the round before the deck scan, so a host that dies in it loses nothing", async () => {
    /**
     * THE CRASH, MODELLED PROPERLY. The test above arms a scan that RAISES, and
     * `collectDeckEvidence` catches — so that path was always survivable and
     * proved nothing about a crash. A crashing host does not raise, it stops
     * answering: the `await` never returns and every line after it is dead.
     *
     * The cost, measured on 2026-08-29: 33 of 49 records in `crashes/` hold all
     * fourteen verdicts and not one became a round. One 4:3 leg produced a full
     * result five times that evening — 14/14 four of them — and archived none,
     * because the host died in the scan every time, between 441s and 572s.
     *
     * So the log is assembled and its header written to the CRASH LOG before the
     * scan. `slideSize()` stays on this side of it too: it is a host call, and
     * after the crash there is no host to ask.
     *
     * THE BUTTON IS NOT PART OF THAT, and the first version of this test said it
     * was. Enabling `demo-log` early looked like belt-and-braces and was a
     * regression: that button becoming enabled is the DRIVER's "round finished"
     * signal (`scripts/round.mjs:2002`), so enabling it before the tail made the
     * driver download a round that was still running — verdicts complete, `deck`
     * absent, trace snapshotted pre-scan. Rounds 313 and 318 are that race, and
     * I misread 313 as proof the new timeout had fired.
     *
     * Durability lives in the crash log, which is the path
     * `scripts/salvage-crashed.mjs` already reads and which works because the
     * pane outlives the host. The button is a completion signal, and a signal
     * used as a safety net is a poor version of both.
     */
    // One clean round first, to learn how many inventory scans a round makes.
    // The last of them is the evidence collector's, and hard-coding that number
    // is how this test would rot the day a scenario starts scanning.
    const clean = captureDownloads();
    $("demo-round").click();
    await settle();
    const perRound = host.deckInventoryScans;
    expect(perRound, "no inventory scan ran at all — the fixture cannot arm the fault").toBeGreaterThan(0);
    clean.restore();

    // Now hang the SECOND round's last scan: its verdicts are all in by then.
    host.deckScanHangsOnScan = perRound * 2;
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();

    // WHILE THE TAIL IS STILL RUNNING the button must stay disabled, or the
    // driver takes it as "finished" and downloads a half-made round. This is the
    // anti-regression, and it is the assertion the first version had backwards.
    expect(
      ($("demo-log") as HTMLButtonElement).disabled,
      "offered the log mid-tail — the driver reads that as the round having finished",
    ).toBe(true);
    // Nothing was offered, so nothing could have been downloaded.
    expect(await dl.lastJson().catch(() => null), "a round still running wrote a file").toBeNull();
    dl.restore();

    // WHERE THE DURABILITY ACTUALLY LIVES is the crash log — `runLogHead` plus
    // one `selftest:` finding per scenario, written before the tail — and it is
    // asserted where it belongs: `crashlog.test.ts` for the record surviving a
    // pane that dies, `salvage-crashed.test.ts` for turning one back into a
    // round. Re-reading it here would need a flush and a store reset mid-run,
    // which tests the harness rather than the pane.
  });

  it("abandons a tail that never answers, so the round FINISHES instead of hanging", async () => {
    /**
     * THE OTHER HALF, and banking the log alone does not get it.
     *
     * The test above proves the verdicts SURVIVE a tail that hangs — but the run
     * itself never completes: no "Round finished", no automatic save, and the
     * driver's `collectRound` is never reached, so it files a crash instead of a
     * round. That is how 33 of 49 crash records came to hold a complete round.
     *
     * `collectDeckEvidence` has always CLAIMED to be best-effort — "a host too
     * far gone to describe its own deck still gets the verdicts out" — and
     * nothing enforced it. A throw was survivable because it catches. A hang was
     * not, and a hang is what a dying host actually does.
     */
    const app = await import("../src/taskpane/app");
    // THE DEFAULT, asserted here because every test below overrides it. The
    // archive says 45s is 5.6x the worst complete tail ever recorded (7.96s over
    // 31 completions) and 2x the worst single deck scan (22.1s of 6,785), with
    // zero of either over 30s. A number that drifted upward would restore the
    // unbounded hang silently.
    expect(app.DECK_EVIDENCE_TIMEOUT_DEFAULT_MS, "the tail's budget moved").toBe(45_000);

    const clean = captureDownloads();
    $("demo-round").click();
    await settle();
    const perRound = host.deckInventoryScans;
    clean.restore();

    host.deckScanHangsOnScan = perRound * 2;
    app._setDeckEvidenceTimeoutForTest(1);
    try {
      const dl = captureDownloads();
      $("demo-round").click();
      await settle();
      // Real time, because the bound is a real timer and `settle` is a 0ms tick.
      await new Promise((r) => setTimeout(r, 60));
      await settle();

      // THE ASSERTION THAT SEPARATES THIS FROM THE TEST ABOVE: the run reached
      // its own end. Nothing was clicked to make that happen.
      expect($("host-note").textContent ?? "", "the round never finished — the tail is still unbounded").toMatch(
        /Round finished/,
      );
      const log = await dl.lastJson();
      expect(log.selftest?.length, "finished without its verdicts").toBeGreaterThan(0);
      expect(log.deck, "claimed deck evidence from a scan that never answered").toBeUndefined();
      /**
       * AND THE FILED ROUND SAYS WHY, which the first version did not.
       *
       * The trace was snapshotted before the tail and re-taken only `if (deck)`,
       * so a round that lost its deck carried a trace ending at the last
       * pre-tail line with nothing explaining the gap. Round 313 on 2026-08-29
       * is the live proof: 14/14 at 4:3, `deck` absent, and not one
       * `collecting deck evidence` line in it.
       *
       * An absent field with no reason is this repo's most repeated defect —
       * unknown printed as nothing — and it was reached here by guarding the
       * diagnostic on the very success it exists to explain the absence of.
       */
      const said = (log.trace?.entries ?? []).some(
        (e: { message?: string }) => e.message === "gave up collecting deck evidence",
      );
      expect(said, "the round lost its deck evidence and does not say why").toBe(true);
      dl.restore();
    } finally {
      app._setDeckEvidenceTimeoutForTest(app.DECK_EVIDENCE_TIMEOUT_DEFAULT_MS);
    }
  });

  /**
   * "Both, one after the other" has never survived a deck that was not empty.
   *
   * Four attempts, four crash dialogs, and every one has the same shape: the
   * file half fills the deck, the shape half then draws onto that larger deck,
   * and the host stops answering. The last one waited 45 seconds for a batch of
   * FIVE shapes on 40 slides. `docs/PUBLISHING.md` splits the two into separate
   * tests on separate decks for exactly this reason, and the dropdown quietly
   * puts them back together.
   *
   * So the option degrades instead of obeying: run the half that works, and say
   * what the other one needs.
   */
  it("refuses the shape half on a deck that is already full, and says why", async () => {
    const dl = captureDownloads();
    host.canInsertFile = true;
    host.slideCount = 40;
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    // The file half still runs — it is the half that works on a big deck, and
    // refusing both would cost the measurement entirely.
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns, "drew the deck shape by shape onto a 40-slide deck").toBe(0);
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs.map((r: { path: string }) => r.path)).toEqual(["file"]);
    // And the LOG says the difference was deliberate. The on-screen note is
    // overwritten by this run's own summary within seconds; a reader of the log
    // would otherwise see a run asked for both halves that quietly did one, and
    // read it as the shape half having crashed.
    expect(log.refusedShapeHalf?.slides, "the log does not say the shape half was refused").toBe(40);
    expect(log.refusedShapeHalf?.why).toMatch(/fresh deck/i);
    dl.restore();
  });

  it("still takes both halves when the deck is empty enough to be worth it", async () => {
    // The gate is a threshold, not a ban. A fresh deck is exactly where the
    // measurement is wanted, and that case has never been the one that crashes.
    host.canInsertFile = true;
    host.slideCount = 1;
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns, "the shape half was refused on an empty deck").toBeGreaterThanOrEqual(1);
  });

  it("does not run the shape path twice when the fast path falls back", async () => {
    // "Both" continues to the shape path after a SUCCESSFUL file insert, which
    // is the one case where drawing the deck twice is the intent. A fast path
    // that landed nothing already falls back to shapes on its own — running
    // the fall-back and then the second leg would draw it three times.
    const dl = captureDownloads();
    host.canInsertFile = true;
    host.buildFileError = new Error("pptxgenjs blew up");
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    // Counted by whole-deck calls, not by `demoRuns` — the shape path also
    // calls insertDemoDeck once per results page at the end of a run.
    expect(host.demoDeckCalls.filter((c) => c.length > 5)).toHaveLength(1);
    $("demo-log").click();
    expect((await dl.lastJson()).runs.map((r: { path: string }) => r.path)).toEqual(["shapes"]);
    dl.restore();
  });

  it("keeps the fast path's log when the shape path then throws", async () => {
    // The failing run is the one worth having a file for. Banking the log at
    // the end of the click would discard the file run along with the throw.
    const dl = captureDownloads();
    host.canInsertFile = true;
    ($("demo-path") as HTMLSelectElement).value = "both";
    host.demoRuns = 0;
    host.demoDeckPageFailures = new Set([1]);
    $("demo-insert").click();
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false);
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs.map((r: { path: string }) => r.path)).toEqual(["file"]);
    dl.restore();
  });
});

describe("watchSelection", () => {
  beforeEach(bootHostPane);

  it("offers the edit banner when an SSF chart is selected", async () => {
    expect(host.selectionListener).toBeTypeOf("function");
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "sel-1", left: 0, top: 0 },
    };
    await host.selectionListener!();
    await settle();
    expect($("selection-banner").style.display).toBe(""); // shown
  });

  it("hides the banner when the selection is not an SSF chart", async () => {
    host.loadSelectionResult = null;
    await host.selectionListener!();
    await settle();
    expect($("selection-banner").style.display).toBe("none");
  });

  /**
   * A real host answered two of these ninety seconds later — the run log reads
   * `gave up waiting what=reading the selected chart afterMs=90000`, twice,
   * during Same Scale. 90s is `READBACK_TIMEOUT_MS`, the budget sized for a
   * twenty-slide repair page; a background listener that decides whether one
   * banner is visible inherited it by passing no budget at all.
   */
  it("gives the host a deadline of its own rather than the readback budget", async () => {
    await host.selectionListener!();
    await settle();
    expect(host.selectionBudgets).toHaveLength(1);
    const budget = host.selectionBudgets[0];
    expect(typeof budget).toBe("number");
    expect(budget!).toBeLessThan(10_000);
  });

  /**
   * The pane moves the selection itself — `showSlide` on every deck pass,
   * `withSlideDeselected` before every off-screen redraw, a slide swap on every
   * resilient update. Each of those fires this handler, so a long run used to
   * queue a selection read per slide behind the work the user actually asked
   * for. The banner is worth exactly one read, and only the last one is current.
   */
  it("asks nothing while an action is running, then catches up once", async () => {
    let release!: () => void;
    host.gate = new Promise<void>((r) => (release = r));
    $("insert").click();
    await settle();

    // The pane's own navigation, three slides' worth.
    await host.selectionListener!();
    await host.selectionListener!();
    await host.selectionListener!();
    await settle();
    expect(host.selectionBudgets, "queued host work while the pane was busy").toHaveLength(0);

    release();
    await settle();
    expect(host.selectionBudgets, "did not catch up after the action ended").toHaveLength(1);
  });

  /** Nothing to catch up on means nothing to ask — an action that never moved
   *  the selection must not provoke a read merely by finishing. */
  it("does not read the selection after an action that never moved it", async () => {
    $("insert").click();
    await settle();
    expect(host.selectionBudgets).toHaveLength(0);
  });

  /**
   * Two clicks in the time one read takes are still one question.
   *
   * Fired without awaiting, which is what PowerPoint does: `addHandlerAsync`
   * takes a callback and never looks at what it returns. Awaiting it here would
   * serialise the two reads by hand and test the harness rather than the pane.
   */
  it("keeps only one selection read outstanding", async () => {
    let release!: () => void;
    host.selectionGate = new Promise<void>((r) => (release = r));
    void host.selectionListener!();
    await settle();
    void host.selectionListener!();
    await settle();
    expect(host.selectionBudgets).toHaveLength(1);
    release();
    await settle();
  });
});

describe("Stop", () => {
  beforeEach(bootHostPane);

  /**
   * Load a chart from the selection, then press Insert — which makes Insert an
   * in-place UPDATE of that chart rather than a fresh insert. The update path
   * is the one that is slow enough to need a stop, and the only one that goes
   * through the gate these tests hold open.
   */
  async function startGatedUpdate(): Promise<() => void> {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    let release!: () => void;
    host.updateGate = new Promise<void>((r) => (release = r));
    $("insert").click();
    await settle();
    return () => {
      release();
      host.updateGate = null;
    };
  }

  it("is offered only while an action is in flight", async () => {
    // The button lives in the status strip and must be absent the rest of the
    // time — a permanent Stop on an idle pane is a button that does nothing.
    const stop = $("status-stop") as HTMLButtonElement;
    expect(stop.hidden).toBe(true);
    const release = await startGatedUpdate();
    expect(stop.hidden, "no way out of a long action").toBe(false);
    release();
    await settle();
    expect(stop.hidden, "left offering a stop after the action ended").toBe(true);
  });

  it("says it is stopping before the host has come back", async () => {
    // The batch already handed to PowerPoint still has to return — up to
    // BATCH_TIMEOUT_MS. Without immediate feedback the pane looks like it
    // ignored the click for as long as that takes.
    const stop = $("status-stop") as HTMLButtonElement;
    const release = await startGatedUpdate();
    stop.click();
    await settle();
    expect(stop.disabled).toBe(true);
    expect(stop.textContent?.toLowerCase()).toContain("stopping");
    release();
    await settle();
  });

  it("keeps the user's stop when another action is started on top of it", async () => {
    /**
     * THE STOP THE USER PRESSED WAS CLEARED, AND THE CANCELLED WORK LANDED.
     *
     * `guard` leaves about thirteen controls enabled during a slow action, and
     * called `resetStop()` on the way IN. So: start an insert, press Stop — the
     * button disables and reads "Stopping…" — then touch any other live
     * control, and the first action reaches its next batch, asks
     * `isStopRequested()`, is told no, and finishes.
     *
     * The outermost action owns the stop flag now. This drives the real
     * sequence through the real handlers rather than calling `resetStop`
     * directly, because the defect was in WHO calls it, not in what it does.
     */
    const release = await startGatedUpdate();
    ($("status-stop") as HTMLButtonElement).click();
    expect(host.stopRequested, "the stop never registered, so this proves nothing").toBe(true);

    // A second action, on a control `guard` leaves live during the first.
    $("harvey-insert").click();
    await settle();

    expect(host.stopRequested, "a second action threw away the stop the user pressed").toBe(true);
    release();
    await settle();
    expect($("host-note").textContent?.toLowerCase() ?? "", "the cancelled action did not report stopping").toContain(
      "stopped",
    );
  });

  it("does not report Done over an action that is still running", async () => {
    /**
     * The same `finally`, the other half: it posted "Done." in green, hid Stop
     * and cleared the elapsed interval while the first action was still in
     * flight — and that interval is the only carrier of the SILENT_RUN_MS
     * watchdog, so the "PowerPoint died and the pane did not" check went with
     * it. Measured before the fix: elapsed read "" 2.3s in, against "2s" when
     * nothing else was clicked.
     */
    const release = await startGatedUpdate();
    $("harvey-insert").click();
    await settle();

    expect($("host-note").textContent, "claimed the pane was finished mid-action").not.toBe("Done.");
    expect($("status-stop").hasAttribute("hidden"), "hid Stop out from under a running action").toBe(false);
    expect($("status-elapsed").textContent, "stopped the clock that carries the silence watchdog").not.toBe("");
    release();
    await settle();
  });

  it("keeps the lights on when the OUTER action finishes first", async () => {
    /**
     * The same defect with the two actions swapped, and the reason ownership is
     * "the last one out turns off the lights" rather than "the one that turned
     * them on".
     *
     * My first fix captured ownership at ENTRY, which is right only while
     * actions finish in the order they started. Here the outer insert returns
     * while the element insert is still going: an entry-captured owner tidies
     * up on its way out — hiding Stop and stopping the clock that carries the
     * silence watchdog — over an action that is still running.
     */
    let releaseElement!: () => void;
    const elementGate = new Promise<void>((r) => (releaseElement = r));
    const releaseInsert = await startGatedUpdate();

    // The nested action, held open past the outer one's return.
    host.insertGate = elementGate;
    $("harvey-insert").click();
    await settle();

    // The OUTER action finishes first.
    releaseInsert();
    await settle();

    expect($("status-stop").hasAttribute("hidden"), "hid Stop over a still-running action").toBe(false);
    expect($("status-elapsed").textContent, "stopped the watchdog clock over a still-running action").not.toBe("");

    releaseElement();
    host.insertGate = null;
    await settle();
    // …and once nothing is left running, the lights DO go out.
    expect($("status-stop").hasAttribute("hidden"), "left Stop showing with nothing running").toBe(true);
  });

  it("reports a stop as a stop, not as a failure", async () => {
    // "Failed: Stopped." reads like the add-in broke. The user pressed the
    // button; the pane should say what happened, and say the work already
    // drawn was kept.
    const release = await startGatedUpdate();
    ($("status-stop") as HTMLButtonElement).click();
    release();
    await settle();
    const said = $("host-note").textContent?.toLowerCase() ?? "";
    expect(said).toContain("stopped");
    expect(said, "blamed the host for the user's own stop").not.toContain("failed");
  });

  it("still sweeps what the stopped redraw left on the slide", async () => {
    // A stop mid-batch leaves a partial chart exactly as a stall does. Leaving
    // it there would make Stop destructive — the one thing a cancel must not be.
    const release = await startGatedUpdate();
    ($("status-stop") as HTMLButtonElement).click();
    release();
    await settle();
    expect(host.calls.swept).toEqual([{ slideId: "s1", ids: ["stray-1"] }]);
  });

  it("does not run the recovery ladder for a stop", async () => {
    // Every rung below the in-place redraw exists to get the chart drawn some
    // other way — rebuild the slide, rasterize it. That is precisely what a
    // user who just pressed Stop has said they do not want, and the slowest
    // rungs would run AFTER the cancel.
    host.canInsertFile = true;
    host.slideHoldsOnlyChart = true;
    // The swap RUNG FAILS, deliberately: it is what makes the picture rung
    // reachable, and the picture rung is the one that leaves a trace here (an
    // update carrying pictureBase64). Asserting against a swap that succeeds
    // proves nothing — a successful swap records nothing either way.
    host.swapOutcome = "failed";
    const release = await startGatedUpdate();
    ($("status-stop") as HTMLButtonElement).click();
    release();
    await settle();
    // No rasterized redraw was attempted after the cancel.
    expect(host.calls.updateChart.filter((c) => c.opts.pictureBase64)).toHaveLength(0);
    // And the pane did not claim to have rebuilt anything.
    const said = $("host-note").textContent?.toLowerCase() ?? "";
    expect(said).toContain("stopped");
    expect(said).not.toMatch(/rebuil|picture/i);
  });
});

describe("the live step list", () => {
  beforeEach(bootHostPane);

  /**
   * The trace module the PANE is using, not the one this file imported.
   *
   * `bootHostPane` calls `vi.resetModules()` before importing the app, so a
   * static import at the top of this file is a different instance: writing to
   * it would leave the pane's subscriber untouched and every assertion below
   * would pass or fail on an empty box for the wrong reason.
   */
  const paneTrace = async () => {
    const mod = await import("../src/core/trace");
    mod.setTracing(true);
    return mod.trace;
  };

  it("puts the newest step at the TOP, where a crash leaves it visible", async () => {
    // Ordering is the whole feature, not a preference. When PowerPoint dies you
    // get whatever pixels were on screen — no scrolling, no clicking, often a
    // dialog over half the pane. A list that grows downwards puts the last
    // thing that happened at the bottom of a small scrolled box, which is
    // exactly where it cannot be relied on to be visible. Growing upwards puts
    // it one line under the header, always.
    const trace = await paneTrace();
    const steps = document.getElementById("demo-steps")!;
    trace("selftest", "scenario starting", { name: "first" });
    trace("selftest", "scenario starting", { name: "second" });
    const lines = steps.textContent!.trim().split("\n");
    expect(lines[0], `newest was not first — got:\n${steps.textContent}`).toContain("name=second");
    expect(lines[1]).toContain("name=first");
  });

  it("carries the data payload, which is where the phase and the verdict live", async () => {
    const trace = await paneTrace();
    const steps = document.getElementById("demo-steps")!;
    trace("error", "drawing the chart's shapes", { error: "PowerPoint did not respond | at=drawing" });
    expect(steps.textContent).toContain("error");
    expect(steps.textContent).toContain("at=drawing");
  });

  it("keeps the newest when it reaches its cap, not the oldest", async () => {
    // The tail is what is read. Truncating the wrong end would leave a list
    // that is full of the start of a run and silent about how it ended.
    const trace = await paneTrace();
    const steps = document.getElementById("demo-steps")!;
    for (let i = 0; i < 320; i++) trace("draw", "batch issued", { n: i });
    const lines = steps.textContent!.trim().split("\n");
    expect(lines.length).toBeLessThanOrEqual(300);
    expect(lines[0], "the newest step fell off the cap").toContain("n=319");
  });

  it("puts the whole box ABOVE the buttons that start a run", () => {
    // Newest-first settles where in the box the last line is. It says nothing
    // about whether the box is on screen, and a real-host round proved the
    // difference: PowerPoint died, the pane came back, and the log was sitting
    // under nine controls and a paragraph of prose — reachable only by
    // scrolling, which is the one thing a crash does not leave you.
    //
    // Asserted as document order rather than pixels, because jsdom has no
    // layout: `compareDocumentPosition` is what actually decides which of two
    // blocks a scrolled-to-top pane shows first.
    // ONE control is allowed above it, and only one.
    //
    // "Probe, then self-test" sits at the top of the section at the owner's
    // request: it is the starred row of the standing test run and the thing he
    // opens the pane to press, and it was three groups down behind two buttons
    // it supersedes. The rule this test defends is about a log buried under
    // "nine controls and a paragraph", and one button with no prose does not
    // bury anything — so the exception is measured rather than argued, and the
    // count is what stops it growing back into the thing the rule forbids.
    const steps = document.getElementById("demo-steps")!;
    const section = steps.closest("section")!;
    expect(section.querySelector("h2")!.textContent).toContain("Testing");
    const above = [...section.querySelectorAll(".actions")].filter(
      (a) => steps.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_PRECEDING,
    );
    expect(
      above.flatMap((a) => [...a.querySelectorAll("button, select, input")]).map((el) => el.id),
      "more than one control now sits above the live step list — a crash leaves the log off-screen",
    ).toEqual(["demo-round"]);
    // Everything else still follows it. FOLLOWING means `actions` comes after
    // `steps`, which is the order a scrolled-to-top pane draws them in.
    const rest = [...section.querySelectorAll(".actions")].filter((a) => !above.includes(a));
    expect(rest.length, "the run controls vanished from the Testing section").toBeGreaterThan(2);
    for (const a of rest)
      expect(
        steps.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING,
        `${a.querySelector("button")?.id ?? "a control group"} comes BEFORE the live step list`,
      ).toBeTruthy();
    // And the header goes with it, or the box arrives with no Copy button and
    // nothing saying which end is newest.
    const head = section.querySelector(".steps-head")!;
    expect(head.compareDocumentPosition(steps) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(head.compareDocumentPosition(rest[0]) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
  });
});

describe("the style a deck carries", () => {
  beforeEach(bootHostPane);

  /**
   * The palette the chart is actually INSERTED with — read off the config the
   * pane hands the renderer, which is the thing that ends up in the deck.
   *
   * Not off `style-export`, which was the first version of this and measured
   * nothing: that button reads the same resolver the chart does, so a test
   * probing it passed with the chart's own wiring reverted. A probe that cannot
   * fail is the shape this repo keeps finding in its own gates.
   */
  const insertedPalette = async () => {
    host.calls.insertScene = [];
    $("insert").click();
    await settle();
    const tag = (host.calls.insertScene[0] as { tagData?: string }).tagData ?? "{}";
    return (JSON.parse(tag) as ChartConfig).style?.palette;
  };
  /** What `style-export` puts in the box — the file a user would send on. */
  const exported = () => JSON.parse(($("json-io") as HTMLTextAreaElement).value) as { palette?: string[] };

  it("draws with the DECK's style when the deck carries one", async () => {
    await bootHostPane({ deckStyle: { palette: ["#dd1122", "#dd3344"] } });
    await settle();
    expect(await insertedPalette()).toEqual(["#dd1122", "#dd3344"]);
    // …and the file the user would send on says the same thing.
    $("style-export").click();
    expect(exported().palette).toEqual(["#dd1122", "#dd3344"]);
  });

  it("hands the browser's style back when one is imported, and the deck's back on request", async () => {
    await bootHostPane({ deckStyle: { palette: ["#dd1122"] } });
    await settle();
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    // An import is an explicit act by the person at the pane; the deck must not
    // silently undo it on the chart they are looking at.
    expect(await insertedPalette()).toEqual(["#22aa33"]);

    $("style-from-deck").click();
    await settle();
    // The note first: inserting a chart writes its own ("Done."), and asserting
    // after that would be asserting the insert's message, not this button's.
    expect($("host-note").textContent).toMatch(/deck/i);
    expect(await insertedPalette()).toEqual(["#dd1122"]);
  });

  it("enables the deck-style buttons when Office arrives AFTER the module ran", async () => {
    // PROVEN ON THE REAL HOST AFTER ROUND 091, in the live pane:
    //
    //     Office.context.host  "PowerPoint"     <- would answer true NOW
    //     style-from-deck      disabled: true
    //     style-to-deck        disabled: true
    //
    // app.ts asked `isPowerPointHost()` at MODULE SCOPE, where `Office.context`
    // does not exist yet, so it answered false for every load inside PowerPoint
    // and disabled both buttons for the whole session. Nothing asked again, so
    // #583's entire user-facing feature was unreachable — and no round could see
    // it, because the thirteen scenarios drive charts and none clicks these.
    hostReadiness.lateOffice();
    await bootHostPane();
    await settle();
    expect(($("style-from-deck") as HTMLButtonElement).disabled, "unreachable on a real host").toBe(false);
    expect(($("style-to-deck") as HTMLButtonElement).disabled).toBe(false);
    hostReadiness.reset();
  });

  it("does not claim the deck is unbranded when the read simply failed", async () => {
    // ROUND 089: `reading the deck's style` hung for its full 90s budget on the
    // real host. The button reported that as "This deck carries no style" and
    // switched the person to the browser's copy — an absence it had not
    // established, about a deck that may well carry one.
    await bootHostPane({ deckStyle: { palette: ["#dd1122"] } });
    await settle();
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    expect(await insertedPalette()).toEqual(["#22aa33"]);

    host.deckStyleUnreadable = true;
    $("style-from-deck").click();
    await settle();
    expect($("host-note").textContent, "reported a failed read as an absence").not.toMatch(/carries no style/i);
    expect($("host-note").textContent).toMatch(/could not read/i);
    // AND IT LEFT THE PERSON'S OWN STYLE ALONE. Switching them to the deck on a
    // read that never answered is the half of this that changes their chart.
    expect(await insertedPalette(), "switched away from the imported style on a failed read").toEqual(["#22aa33"]);
  });

  it("saves the style the pane is actually drawing with", async () => {
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"], fontSize: 12 });
    $("style-import").click();
    $("style-to-deck").click();
    await settle();
    expect(host.deckStyleWrites).toEqual([{ palette: ["#22aa33"], fontSize: 12 }]);
    expect($("host-note").textContent).toMatch(/travels with the file/i);
  });

  it("says the deck was NOT updated on a host that cannot store one", async () => {
    host.deckStyleWritable = false;
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    $("style-to-deck").click();
    await settle();
    expect($("host-note").textContent).toMatch(/cannot store a deck style/i);
  });

  it("leaves the browser's style in force when the deck carries none", async () => {
    await bootHostPane({ deckStyle: null });
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    await settle();
    expect(await insertedPalette()).toEqual(["#22aa33"]);
  });
});

describe("the pane's host-friction double", () => {
  it("carries every counter the real module has", async () => {
    // The double named four of eight. Comparing KEY SETS rather than a list of
    // names is the point: a list is what went stale in the first place.
    const actual = (await vi.importActual("../src/render/powerpoint")) as typeof import("../src/render/powerpoint");
    const mocked = (await import("../src/render/powerpoint")) as unknown as {
      hostFrictionCounts: () => Record<string, number>;
    };
    const real = Object.keys(actual.hostFrictionCounts()).sort();
    const fake = Object.keys(mocked.hostFrictionCounts()).sort();
    expect(fake, "the double has drifted from the real counter set").toEqual(real);
  });
});

/**
 * EVERY CONTROL IN THE PANE HAS A NAME, which until 2026-08-30 was untrue of
 * thirty-six of them.
 *
 * AppSource validation tests accessibility, so this is a shipping gate and not
 * only a courtesy. What the audit found, all of it in controls the pane builds
 * in JavaScript rather than in the HTML:
 *
 *   - the DATASHEET, the pane's primary data-entry surface: twenty bare
 *     `<input>`s in bare `<td>`s, no label, no `<th>`, no caption. A screen
 *     reader announced "edit, blank" twenty times.
 *   - eight `<label>`s each wrapping two to four controls. A label names only
 *     its FIRST labelable descendant — that is the spec, not a lint opinion —
 *     so thirteen more were anonymous: the axis-scale maximum, the label suffix
 *     and locale, the log-scale checkbox, both difference-arrow categories.
 *   - three static controls with nothing at all: the template picker, the JSON
 *     automation box, the agenda chapters box.
 *
 * The assertion is deliberately GENERIC rather than a list of ids. A named list
 * would pass the day someone adds the thirty-seventh unnamed control, which is
 * exactly how the first thirty-six arrived.
 */
describe("every control the pane builds can be named by a screen reader", () => {
  beforeEach(bootHostPane);

  /** The accessible name of a control, by the rules a screen reader uses. */
  const accessibleName = (el: Element): string => {
    const aria = el.getAttribute("aria-label");
    if (aria?.trim()) return aria.trim();
    const by = el.getAttribute("aria-labelledby");
    if (by) {
      const ref = by
        .split(/\s+/)
        .map((id) => document.getElementById(id)?.textContent ?? "")
        .join(" ")
        .trim();
      if (ref) return ref;
    }
    const title = el.getAttribute("title");
    if (title?.trim()) return title.trim();
    const forLabel = el.id ? document.querySelector(`label[for="${el.id}"]`) : null;
    if (forLabel?.textContent?.trim()) return forLabel.textContent.trim();
    // A WRAPPING label names its FIRST labelable descendant and no other. This
    // is the rule the eight multi-control rows fell foul of, so the test has to
    // model it rather than accept any ancestor label.
    const wrapping = el.closest("label");
    if (wrapping) {
      const first = wrapping.querySelector("input, select, textarea, button");
      if (first === el && wrapping.textContent?.trim()) return wrapping.textContent.trim();
    }
    return "";
  };

  const controls = (): Element[] =>
    [...document.querySelectorAll("input, select, textarea")].filter(
      (el) => (el as HTMLInputElement).type !== "hidden",
    );

  it("names every input, select and textarea in the pane", () => {
    const all = controls();
    expect(all.length, "the pane rendered no controls — the fixture is not booting it").toBeGreaterThan(20);
    const unnamed = all.map((el) => (accessibleName(el) ? null : describeControl(el))).filter(Boolean);
    expect(unnamed, `${unnamed.length} control(s) have no accessible name`).toEqual([]);
  });

  it("names every datasheet cell by its own row and column", () => {
    // Named from the sheet's headers, so the name tracks a renamed series rather
    // than being a static "cell 3, 2".
    const cells = [...document.querySelectorAll(".datasheet input")];
    expect(cells.length, "the datasheet rendered no cells").toBeGreaterThan(0);
    for (const c of cells) expect(accessibleName(c), `a datasheet cell is anonymous`).not.toBe("");
  });
});

/** Enough to find the offender in a failure message without opening the pane. */
function describeControl(el: Element): string {
  const tag = el.tagName.toLowerCase();
  const id = el.id ? `#${el.id}` : "";
  const cls = el.className ? `.${String(el.className).split(/\s+/)[0]}` : "";
  const type = el.getAttribute("type") ? `[${el.getAttribute("type")}]` : "";
  return `${tag}${type}${id}${cls}`;
}

/**
 * STRUCTURE AND STATE, the half of accessibility that naming does not cover.
 *
 * A control can have a perfect name and still be unusable: a screen-reader user
 * navigates by heading, and needs to be told which of twenty-five buttons is the
 * one currently in effect. Neither was true of the chart gallery.
 */
describe("the chart gallery can be navigated and its state can be heard", () => {
  beforeEach(bootHostPane);

  it("gives each chart family a real heading, not a div that looks like one", () => {
    // Six families — "Columns & bars", "Line & area" … — were
    // `<div class="group-label">`, so the gallery was one flat run of buttons
    // with nothing to jump between. h3 is the level the rest of the pane uses
    // for a sub-heading inside a section.
    const families = [...document.querySelectorAll(".group-label")];
    expect(families.length, "the gallery rendered no family headings").toBeGreaterThan(1);
    const notHeadings = families.filter((el) => !/^H[1-6]$/.test(el.tagName)).map((el) => el.tagName.toLowerCase());
    expect(notHeadings, "a chart family is styled as a heading but is not one").toEqual([]);
  });

  it("says which chart type is currently chosen", () => {
    const thumbs = [...document.querySelectorAll("button.thumb")];
    expect(thumbs.length, "the gallery rendered no chart types").toBeGreaterThan(5);
    // Every thumb must state its state, not only the chosen one — a button with
    // no `aria-pressed` at all reads as a plain button beside ones that toggle.
    const silent = thumbs.filter((t) => t.getAttribute("aria-pressed") === null);
    expect(silent.length, `${silent.length} chart types do not say whether they are selected`).toBe(0);
    // And exactly one is pressed: the visible `.active` class and the announced
    // state must be the same fact, or the two audiences are told different things.
    const pressed = thumbs.filter((t) => t.getAttribute("aria-pressed") === "true");
    const active = thumbs.filter((t) => t.classList.contains("active"));
    expect(pressed.length, "no chart type is announced as selected").toBe(1);
    expect(pressed[0], "the announced selection is not the one shown as selected").toBe(active[0]);
  });
});

/**
 * LIVE REGIONS — the third piece, after a name and a role. A region can be
 * perfectly named, perfectly polite, and still announce nothing.
 *
 * TWO DEFECTS, ONE SUBJECT.
 *
 * 1. A LIVE REGION WRITTEN WHILE IT IS STILL `display: none`. `note()` set
 *    `#host-note`'s text and only THEN lifted `hidden` from `#status-strip`
 *    (`.status-strip[hidden] { display: none }` — an author rule in
 *    taskpane.css, so the collapse is real). A live region announces a CHANGE
 *    made while it is rendered; content that arrives together with the region
 *    is its initial content and is not announced. The strip ships `hidden` and
 *    `note("")` at boot leaves it that way, so what this cost was the FIRST
 *    message of a session — "Working…" on the user's first insert. `#slow-offer`
 *    had the identical ordering and is covered here too, because the fix that
 *    misses its second call site is this repo's most repeated mistake.
 *
 * 2. A NAME THAT ARIA THROWS AWAY. `#demo-steps` is a bare `<pre>`, which maps
 *    to the `generic` role, and `generic` PROHIBITS `aria-label` — so the live
 *    log reached a screen reader unnamed, and axe-core's `aria-prohibited-attr`
 *    flags it. AppSource validation tests accessibility, so both are shipping
 *    gates and not courtesies.
 *
 * WHAT IS AND IS NOT PROVEN HERE. There is no screen reader in this harness, so
 * what these assert is the DOM — the ordering of the two mutations, and the
 * role/attribute pairing — which is the half that is ours to get right. The
 * announcement itself rests on the platform rule, not on an experiment run here.
 */
describe("the pane's live regions can actually announce", () => {
  beforeEach(bootHostPane);

  /**
   * Which of the two nodes the DOM touched first.
   *
   * A MutationObserver rather than reading state afterwards, because AFTERWARDS
   * the two orderings are indistinguishable: strip open, text written, either
   * way. The order of the records IS the defect.
   */
  const watchOrder = (container: Element, region: Element) => {
    const seen: string[] = [];
    const obs = new MutationObserver((records) => {
      for (const r of records) seen.push(r.target === container ? "container" : "region");
    });
    obs.observe(container, { attributes: true, attributeFilter: ["hidden"] });
    obs.observe(region, { childList: true, characterData: true, subtree: true });
    return { seen, stop: () => obs.disconnect() };
  };

  it("opens #status-strip before writing the first status message into it", async () => {
    const strip = $("status-strip");
    // Both preconditions, or the test proves nothing: an already-open strip
    // makes any ordering pass, and a note already on screen means the first
    // announcement of the session has been and gone before the watch started.
    expect(strip.hasAttribute("hidden"), "the strip was already open").toBe(true);
    expect($("host-note").textContent, "a message was posted before the watch started").toBe("");

    const watch = watchOrder(strip, $("host-note"));
    $("insert").click();
    await settle();
    watch.stop();
    expect(watch.seen.length, "neither the strip nor the note moved at all").toBeGreaterThan(1);
    expect(watch.seen[0], `the live region was written while hidden: ${watch.seen.join(" → ")}`).toBe("container");
  });

  it("opens #slow-offer before writing the question into it", async () => {
    const box = $("slow-offer");
    expect(box.hasAttribute("hidden"), "the offer was already open").toBe(true);

    const watch = watchOrder(box, $("slow-offer-text"));
    // Enough shapes on the slide to cross SLOW_INSERT_MS and raise the offer.
    host.slideShapes = Array.from({ length: 80 }, (_, i) => ({
      left: 40 + (i % 4) * 120,
      top: 70 + Math.floor(i / 4) * 20,
      width: 110,
      height: 18,
    }));
    $("insert").click();
    for (let i = 0; i < 20 && $("slow-offer").hasAttribute("hidden"); i++) await settle();
    expect($("slow-offer").hasAttribute("hidden"), "the pane never offered a slide").toBe(false);
    watch.stop();
    expect(watch.seen[0], `the live region was written while hidden: ${watch.seen.join(" → ")}`).toBe("container");
    // Answer it, so the insert this started does not run on into the next test.
    $("slow-offer-here").click();
    await settle();
  });

  /**
   * GENERIC, NOT A LIST OF IDS — the same reasoning as the naming sweep above.
   * A test that named `#demo-steps` would pass the day the next `aria-label`
   * lands on a `<div>`, which is how these arrive.
   *
   * The tags are the ones whose implicit role prohibits an accessible name in
   * ARIA 1.2: `generic` (div, span, pre, b, i, u, s, small, q, code, samp, kbd,
   * var), `paragraph` (p), `emphasis` (em), `strong` (strong). An explicit
   * `role` replaces the implicit one and settles the question, so carrying one
   * is the way out — that is what `role="log"` on `#demo-steps` is for.
   */
  it("never hangs an accessible name on an element whose role forbids one", () => {
    const NAME_FORBIDDEN = [
      "div",
      "span",
      "pre",
      "p",
      "b",
      "i",
      "u",
      "s",
      "small",
      "q",
      "code",
      "samp",
      "kbd",
      "var",
      "em",
      "strong",
    ];
    const named = [...document.querySelectorAll("[aria-label], [aria-labelledby]")];
    expect(named.length, "the pane rendered no named elements — the fixture is not booting it").toBeGreaterThan(10);
    const discarded = named
      .filter((el) => NAME_FORBIDDEN.includes(el.tagName.toLowerCase()) && !el.getAttribute("role"))
      .map((el) => describeControl(el));
    expect(discarded, `${discarded.length} element(s) carry a name ARIA discards`).toEqual([]);
  });

  it("names the live step log and gives it a role that keeps the name", () => {
    const steps = $("demo-steps");
    expect(steps.tagName, "the live log is no longer a <pre> — recheck the role mapping").toBe("PRE");
    expect(String(steps.getAttribute("aria-label") ?? "").trim(), "the live log has no name").not.toBe("");
    // `log` is the ARIA role for this exact thing — new information added in a
    // meaningful order — and unlike `region` it is not a landmark, so a
    // diagnostic panel does not gain a stop in the landmark list.
    expect(steps.getAttribute("role"), "a bare <pre> is role=generic, where the name is discarded").toBe("log");
    // Kept explicit although `log` implies it: an implicit value depends on the
    // AT mapping the role, and this one was here and working first.
    expect(steps.getAttribute("aria-live")).toBe("polite");
  });
});

/**
 * THE DOOR BACK THROUGH THE WALL, which had no lock on it.
 *
 * A chart over the web shape budget is auto-inserted as a picture, with a
 * message ending "…\"Explode to native shapes\" turns it back." `doExplode` then
 * rebuilt with `render: "shapes"` and went straight to `updateChartInSlide`,
 * never consulting `wantsAutoPicture` — whose only other call site in the whole
 * codebase is the insert path. The product invited the user through the one door
 * it had just told them was dangerous.
 *
 * 17 of the 123 shipped charts exceed the budget; the largest is 401 shapes,
 * which at this host's measured rate is many minutes of work — well past the
 * 425-470s window where PowerPoint has been dying all week. The user loses the
 * session and whatever was unsaved, having followed the add-in's own advice.
 */
describe("Explode respects the same budget the insert path enforces", () => {
  beforeEach(bootHostPane);

  const pictureChart = () => ({
    configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
    target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
  });

  it("offers a slide of its own, as shapes, instead of silently rasterising", async () => {
    /**
     * THE QUESTION THAT WAS NEVER ASKED. A chart over the web budget was
     * rasterised before anything was put to the user, and Explode then refused
     * on the same predicate — so the picture was not the fallback, it was the
     * only reachable outcome.
     *
     * The alternative is measured: the same chart as a one-slide .pptx is
     * ~280ms and ONE call, against 46.4s of shape-by-shape drawing whose
     * per-shape cost climbs 161ms to 784ms as the slide fills. It keeps its
     * config tag, so it stays editable — which a picture also does, but a
     * picture cannot be nudged, restyled by the ribbon, or printed as vector.
     */
    const raster = stubRaster();
    try {
      host.autoPicture = true;
      host.canInsertFile = true;
      $("insert").click();
      await settle();
      // The offer is up, and it names the cost rather than just asking.
      expect($("slow-offer").hidden, "the too-dense chart was rasterised without asking").toBe(false);
      expect($("slow-offer-text").textContent ?? "").toMatch(/slide of its own|editable shapes/i);
      $("slow-offer-own").click();
      await settle();
      // It went as a FILE, not as shapes drawn one batch at a time.
      expect(host.calls.insertFile, "the accepted offer did not hand over a generated slide").toHaveLength(1);
      expect(host.calls.insertScene, "it drew the shapes anyway, which is the 46-second path").toHaveLength(0);
      expect($("host-note").textContent ?? "").toMatch(/own|editable shapes/i);
    } finally {
      raster.restore();
    }
  });

  it("falls back to the picture when the generated slide does not land", async () => {
    /**
     * THE PICTURE IS THE FLOOR. `insertSlidesFromPptx` measures what the host
     * actually took, from a before-and-after slide count, because a host that
     * drops the call reports no error. Anything other than exactly one slide
     * means the chart is not where it was meant to be, and the user must still
     * end up with a chart.
     */
    const raster = stubRaster();
    try {
      host.autoPicture = true;
      host.canInsertFile = true;
      host.insertFileLands = 0; // the host took the call and did nothing
      $("insert").click();
      await settle();
      $("slow-offer-own").click();
      await settle();
      expect(host.calls.insertFile, "never tried the file route").toHaveLength(1);
      expect(host.calls.insertScene, "the chart was lost — no picture went in either").toHaveLength(1);
      /**
       * BOTH FACTS, SETBACK FIRST.
       *
       * This assertion used to read `/picture/i` alone, under a comment saying
       * the user was told the outcome but not that their choice had failed, and
       * that unpicking it meant restructuring how one note channel carries two
       * facts. That restructure is what `insertOutcomeSentence` is, so the
       * comment has been replaced by the assertion it was standing in for.
       */
      const said = $("host-note").textContent ?? "";
      expect(said, `did not lead with the failed choice: ${said}`).toMatch(/^Could not add the slide/);
      expect(said, `dropped the outcome: ${said}`).toMatch(/picture/i);
      expect($("host-note").className, "a failed choice reported as success").toMatch(/status-err/);
    } finally {
      host.insertFileLands = null;
      raster.restore();
    }
  });

  it("does not ask when the user chose a picture themselves", async () => {
    // Ticking "Insert as picture" is a request, not a refusal to draw. Asking
    // again would question a decision the user already made.
    const raster = stubRaster();
    try {
      host.autoPicture = true;
      host.canInsertFile = true;
      const box = $("render-image") as HTMLInputElement;
      box.checked = true;
      box.dispatchEvent(new Event("change", { bubbles: true }));
      $("insert").click();
      await settle();
      expect($("slow-offer").hidden, "questioned a picture the user asked for").toBe(true);
      expect(host.calls.insertFile).toHaveLength(0);
    } finally {
      const box = $("render-image") as HTMLInputElement;
      box.checked = false;
      box.dispatchEvent(new Event("change", { bubbles: true }));
      raster.restore();
    }
  });

  it("does not promise, at insert time, the door it refuses at explode time", async () => {
    /**
     * THE LOCK WENT ON THE DOOR AND THE SIGN STAYED UP.
     *
     * The guard tested below was added so `doExplode` could not walk a user
     * into a 401-shape draw. Nobody changed the sentence the INSERT path
     * prints, which ended: `Inserted as a picture; "Explode to native shapes"
     * turns it back.` Same predicate — `wantsAutoPicture`, same arguments —
     * with one branch promising the door and the other refusing it.
     *
     * Being told the way back exists and finding it locked is worse than never
     * being offered it, and it is the exact complaint a rival markets this
     * whole category against: "a colleague asks you to just tweak the Q3 bar,
     * and there's nothing to tweak."
     *
     * Pinned together here so a later change has to keep them agreeing.
     */
    /**
     * The raster has to SUCCEED, or this tests the wrong branch. jsdom decodes
     * no SVG, so without `stubRaster` the insert falls through to "could not be
     * turned into a picture either" — a different sentence, which never carried
     * the false promise. The first draft of this test asserted against that one
     * and proved nothing about the message it was written for.
     */
    const raster = stubRaster();
    let said: string;
    try {
      host.autoPicture = true;
      $("insert").click();
      await settle();
      said = $("host-note").textContent ?? "";
    } finally {
      raster.restore();
    }
    expect(said, `the too-dense insert said nothing about a picture: ${said}`).toMatch(/picture/i);
    expect(said, `still promising Explode while doExplode refuses it: ${said}`).not.toMatch(/turns it back/i);
    /**
     * AND IT SAYS WHAT IS TRUE, which was never said at all. The picture keeps
     * the chart's config tag, so this pane can still load it, change its data
     * and restyle it; only the conversion to native SHAPES needs a host that
     * can draw them. That is a better offer than the one being made, and the
     * user was not being given it.
     */
    expect(said, `did not tell the user what they CAN still do: ${said}`).toMatch(/edit it here|reloads its data/i);
  });

  it("refuses to explode a chart the budget would have made a picture", async () => {
    host.autoPicture = true; // the renderer says: too many shapes for this host
    host.loadSelectionResult = pictureChart();
    $("explode").click();
    await settle();
    expect(host.calls.updateChart, "exploded a chart the insert path would have refused").toHaveLength(0);
    const said = $("host-note").textContent ?? "";
    expect(said, `refused without saying why: ${said}`).toMatch(/shapes/i);
    // AND IT ASKED THE RIGHT QUESTION. The renderer is mocked here, so a wrong
    // argument is invisible to the return value — `alreadyPicture: true` would
    // disable this guard for ever and every assertion above would still pass.
    const pp = (await import("../src/render/powerpoint")) as unknown as {
      wantsAutoPicture: { mock: { calls: [number, Record<string, unknown>][] } };
    };
    const asked = pp.wantsAutoPicture.mock.calls.at(-1);
    expect(asked?.[1], "asked whether an ALREADY-picture wants to be one — the guard would never fire").toMatchObject({
      alreadyPicture: false,
      web: true,
      canPicture: true,
    });
  });

  it("refuses to explode into marks this host cannot draw", async () => {
    /**
     * THE SECOND DOOR, and the one a desktop host actually walks through.
     *
     * A chart inserted as a picture because this PowerPoint has no
     * `Shape.rotation` would explode into shapes with its wedges MISSING — the
     * user asks for native shapes and gets a pie with no slices, having pressed
     * a button the add-in offered. Same trap as the density guard above, a
     * different host: that one is the web, this one is everything else.
     */
    host.autoPicture = false; // density is NOT the reason here
    host.cannotDraw = { what: ["4 pie slices"], nodes: 4 };
    host.loadSelectionResult = pictureChart();
    $("explode").click();
    await settle();
    expect(host.calls.updateChart, "exploded a chart into marks this host cannot draw").toHaveLength(0);
    const said = $("host-note").textContent ?? "";
    expect(said, `refused without saying what would be lost: ${said}`).toMatch(/4 pie slices/);
  });

  it("explodes when the budget is not the problem", async () => {
    // The common case must be untouched: most charts are nowhere near the budget,
    // and a guard that refused everything would be worse than the bug.
    host.autoPicture = false;
    host.loadSelectionResult = pictureChart();
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20 };
    $("explode").click();
    await settle();
    expect(host.calls.updateChart, "refused an ordinary chart").toHaveLength(1);
  });
});
