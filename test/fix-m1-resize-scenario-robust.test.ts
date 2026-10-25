// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  RESIZE_SPECIMEN,
  ROUTINE_SCENARIO_NAMES,
  hostSeemsSick,
  runSelfTest,
  scenarioBlame,
  selfTestNeedsAttention,
  setSelfTestRasterizer,
  stalledResize,
  type ScenarioResult,
} from "../src/taskpane/selftest";
import {
  _resetDeadlinesFiredForTest,
  _resetStallContextForTest,
  _setBatchTimeoutForTest,
  _setReadbackTimeoutForTest,
  insertSceneIntoSlide,
  isTimeout,
  listChartsInDeck,
} from "../src/render/powerpoint";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { onTrace, setTracing, traceLog, type TraceEntry } from "../src/core/trace";
import type { ChartConfig } from "../src/core/types";
import { faults, installHost, makeSlide, stallSyncOn, trips, type FakeSlide } from "./helpers/office-host";

/**
 * Three switches on the product, for the cases no fake fault reaches on its own,
 * as `selftest-resize.test.ts` has. `refuseDelete`: the by-id delete the scenario
 * takes its chart away with, which the live host refuses for a shape drawn seconds
 * earlier. `undoTheResize`: the edit comes to a chart that has been put back to the
 * size it was drawn at, which is what a product that lost the person's resize
 * would meet. `updateThrows`: the update throws something that is not the host
 * going quiet.
 *
 * The module is rebuilt with GETTERS and not spread. `deadlinesFired`, `lastStall`
 * and `lastLateSync` are live bindings the runner diffs around a scenario, and a
 * `...real` spread copies each once, so a runner reading them through the mock sees
 * a host that never stalled and never answered late: the very evidence these tests
 * assert on the verdicts of a stall.
 */
const wrapped = vi.hoisted(() => ({
  refuseDelete: false,
  undoTheResize: false,
  updateThrows: null as Error | null,
}));
vi.mock("../src/render/powerpoint", async (importOriginal) => {
  const real = await importOriginal<typeof import("../src/render/powerpoint")>();
  const mocked: Record<string, unknown> = {};
  for (const key of Object.keys(real))
    Object.defineProperty(mocked, key, {
      enumerable: true,
      configurable: true,
      get: () => (real as Record<string, unknown>)[key],
    });
  Object.defineProperty(mocked, "updateChartInSlide", {
    enumerable: true,
    value: async (...args: Parameters<typeof real.updateChartInSlide>) => {
      if (wrapped.updateThrows) throw wrapped.updateThrows;
      if (wrapped.undoTheResize) await real.resizeShapeBy(args[1].slideId, args[1].shapeId, 0.5, 0.5);
      return real.updateChartInSlide(...args);
    },
  });
  Object.defineProperty(mocked, "deleteShapesById", {
    enumerable: true,
    value: async (...args: Parameters<typeof real.deleteShapesById>) =>
      wrapped.refuseDelete ? 0 : real.deleteShapesById(...args),
  });
  return mocked as typeof real;
});

/**
 * `an edit after the person resizes a chart` as an INSTRUMENT.
 *
 * Rounds 502, 504 and 505 each ended it `FAILED threw: PowerPoint did not respond
 * while drawing shapes 1-10 of 24 (45s)`, and 503 passed in 164s. The product's
 * logic worked every time it was reached (`redrawing at the size the user gave the
 * chart`, kept [600,360], scale 2x via members); the scenario was a 24-shape chart
 * drawn in three batches of 14-28s onto a slide already holding 48-59 shapes, last
 * in the round, straight after the scenario that wears the host most. What these
 * hold down is that it is a fair question now: small, early, drawn where the slide
 * is quiet, silent about a host that would not answer, and still loud about a
 * product that gets the resize wrong.
 */

const NAME = "an edit after the person resizes a chart";
const PREDECESSOR = "an update follows a moved chart";

beforeEach(() => {
  // Per-slide draw counts and the stall record are module state that outlives a
  // test; slide ids repeat from one fake deck to the next.
  _resetStallContextForTest();
  _resetDeadlinesFiredForTest();
});

afterEach(() => {
  vi.unstubAllGlobals();
  setTracing(false);
  onTrace(undefined);
  _setBatchTimeoutForTest(45_000);
  _setReadbackTimeoutForTest(90_000);
  stallSyncOn.clear();
  wrapped.refuseDelete = false;
  wrapped.undoTheResize = false;
  wrapped.updateThrows = null;
});

const freshDeck = (): FakeSlide[] => {
  const deck = [makeSlide("s1")];
  installHost(deck);
  faults.groupResizeScalesMembers = true;
  return deck;
};

const drive = async (
  onResult?: (r: ScenarioResult) => void,
): Promise<{ result: ScenarioResult; all: ScenarioResult[] }> => {
  const all = await runSelfTest("probe", NAME, onResult);
  return { result: all.find((r) => r.name === NAME)!, all };
};

const chartConfigs = (deck: FakeSlide[]): ChartConfig[] =>
  deck
    .flatMap((s) => s.created)
    .filter((s) => !s.deleted && s.tagStore.get("POWERCHART_CONFIG"))
    .map((s) => JSON.parse(s.tagStore.get("POWERCHART_CONFIG")!) as ChartConfig);

/** Every shape on the deck that has not been deleted, a group's members and the group itself included. */
const liveShapes = (deck: FakeSlide[]): number => deck.flatMap((s) => s.created).filter((s) => !s.deleted).length;

/** The draw batches the scenario issued, from its announcement to its verdict. */
const ownBatches = (): TraceEntry[] => {
  const entries = traceLog().entries;
  const from = entries.findIndex((e) => e.message === "scenario starting" && e.data?.name === NAME);
  const to = entries.findIndex(
    (e, i) => i > from && /^scenario (passed|FAILED|skipped)$/.test(e.message) && e.data?.name === NAME,
  );
  return entries.slice(from, to < 0 ? undefined : to).filter((e) => e.scope === "draw" && e.message === "batch issued");
};

/**
 * Make the NEXT draw batch stall, with a budget short enough for the draw to give
 * up on it. Armed from a trace watcher at the moment the test wants the stall, so
 * the test names where in the scenario the host goes quiet and not which sync
 * number that happens to be.
 */
const stallTheNextDraw = () => {
  faults.stallNextDrawBatch = true;
  _setBatchTimeoutForTest(15);
};
const stallAfter = (match: (e: TraceEntry) => boolean) => {
  setTracing(true);
  onTrace((e) => {
    if (match(e)) stallTheNextDraw();
  });
};

describe("the chart it draws", () => {
  it("is nine shapes in ONE batch, with the bars the judge compares", () => {
    const scene = buildChart(RESIZE_SPECIMEN);
    expect(scene.nodes.length, "more than one batch (SHAPES_PER_SYNC is 10)").toBeLessThanOrEqual(10);
    // `judgeResizeKept` wants three rects at least 8pt on a side read at both
    // moments; a fourth is what lets a host leave one unanswered and be judged.
    const bars = scene.nodes.filter((n) => n.kind === "rect" && n.w >= 8 && n.h >= 8);
    expect(bars.length).toBeGreaterThanOrEqual(4);
  });

  it("is drawn in one batch on the host, and redrawn in one", async () => {
    freshDeck();
    setTracing(true);
    const { result } = await drive();
    expect(result.ok, result.detail).toBe(true);
    const batches = ownBatches();
    // Setup draw and the redraw after the edit: one batch each. The sample chart
    // was three each, and the batches are what the live runs stalled on.
    expect(batches.map((b) => `${b.data?.upTo}/${b.data?.total}`)).toEqual(["9/9", "9/9"]);
  });

  it("still resizes, edits and PASSES with the members following, and says how many bars it read", async () => {
    freshDeck();
    const { result } = await drive();
    expect(result.skipped, result.detail).toBeFalsy();
    expect(result.ok, result.detail).toBe(true);
    expect(result.detail).toMatch(/redrawn at one scale and at the size the person made it/);
    expect(result.detail).toMatch(/4 bars matched the layout at that size; members followed the resize/);
  });

  it("carries no title, so a chart a refused sweep leaves behind is never one a later scenario is offered", async () => {
    // `probeCharts` offers a scenario only a chart whose title starts with the run's
    // prefix. This scenario now runs in the MIDDLE of the battery, where a leftover
    // would be found by `explodePicture` and the selection scenarios; the sweep is
    // refused for a shape drawn seconds earlier in most live rounds.
    const deck = freshDeck();
    wrapped.refuseDelete = true;
    await drive();
    const leftovers = chartConfigs(deck).filter((c) => !/twice|dup/.test(String(c.title)));
    expect(leftovers.length, "the refusal left nothing behind, so nothing was tested").toBeGreaterThan(0);
    for (const c of leftovers)
      expect(
        typeof c.title === "string" && c.title.startsWith("probe"),
        `a leftover titled "${c.title}" would be offered to the scenarios after it as a probe chart`,
      ).toBe(false);
  });
});

describe("where it draws", () => {
  it("takes the probe slide this run has drawn on least, and not the first one the scan lists", async () => {
    // The old pick was `hosts[0]`, the one slide every scenario that takes
    // `found[0]` piles onto: it drew on a slide holding 48-59 shapes while four
    // other probe slides held 9, 16, 24 and 28 (rounds 502-505).
    freshDeck();
    await runSelfTest("probe", "two slides claiming one slot");
    const probes = (await listChartsInDeck()).charts.filter((c) =>
      String((JSON.parse(c.configJson) as ChartConfig).title).startsWith("probe"),
    );
    expect(probes.length, "no probe charts, so there is nothing to choose between").toBeGreaterThan(1);
    const first = probes[0].target.slideId;
    // Load the first slide as the battery loads it: a chart drawn by the renderer.
    // Its tag carries a title that does not start with the prefix, so it is not a
    // probe chart itself.
    const loading = sampleConfig("clustered");
    await insertSceneIntoSlide(buildChart(loading), {
      slideId: first,
      tagData: JSON.stringify({ ...loading, title: "load" }),
    });
    setTracing(true);
    const { result } = await drive();
    expect(result.ok, result.detail).toBe(true);
    const slides = new Set(ownBatches().map((b) => b.data?.onSlideKey));
    expect([...slides], "it drew on more than one slide").toHaveLength(1);
    expect(slides.has(first), "it drew on the slide the run had loaded most, which is where the stalls were").toBe(
      false,
    );
  });
});

describe("a host that stops answering", () => {
  it("is a SKIP, saying nothing was asked, when it stalls the draw of the chart the scenario needs", async () => {
    const deck = freshDeck();
    let heldAtStart = -1;
    setTracing(true);
    onTrace((e) => {
      if (e.message === "scenario starting" && e.data?.name === NAME) {
        heldAtStart = liveShapes(deck);
        stallTheNextDraw();
      }
    });
    const { result, all } = await drive();
    expect(result.detail, "a host that went quiet was filed as a failure of ours").not.toMatch(/^threw:/);
    expect(result.ok).toBe(false);
    expect(result.skipped, result.detail).toBe(true);
    expect(result.blind, "nothing was learned about the product: the round must still be flagged").toBe(true);
    expect(result.detail).toMatch(/while the scenario drew the chart it needs, so nothing was asked/);
    expect(result.detail).toMatch(/never resized and no edit was made/);
    // The host's own sentence, which is also what files it as a stall.
    expect(result.detail).toMatch(/PowerPoint did not respond while drawing shapes 1-9 of 9/);
    // And the evidence the runner adds to every stalled scenario, because it is
    // the runner that files this one.
    expect(result.detail, "the abandoned call came back and the report did not say so").toMatch(/DID come back/);
    expect(result.detail).toMatch(/the last thing the host answered was/);
    expect(scenarioBlame(result)).toBe("not-run");
    expect(selfTestNeedsAttention(all)).toBe(true);
    expect(hostSeemsSick(result, true)).toBe(true);
    // It never got as far as a verdict about the product, and the slide holds what
    // it held. A chart whose draw the host gave up on was never grouped or tagged,
    // so the scenario has no id for its shapes and takes them by what the slide
    // gained: counted on the deck, since a filter on its config finds nothing
    // whether or not they are there.
    expect(heldAtStart, "the scenario's start was never seen, so this proves nothing").toBeGreaterThan(0);
    expect(liveShapes(deck), "a half-drawn chart is still on the slide").toBe(heldAtStart);
  });

  it("keeps the sweep's note when the stalled scenario's chart could not be taken away", async () => {
    freshDeck();
    wrapped.refuseDelete = true;
    stallAfter((e) => e.message === "scenario starting" && e.data?.name === NAME);
    const { result } = await drive();
    expect(result.skipped, result.detail).toBe(true);
    expect(result.blind).toBe(true);
    expect(result.detail).toMatch(/note, not part of the verdict: the sweep did not take this scenario's chart off/);
  });

  it("is a SKIP carrying the size it decided to keep, when the product's redraw stalls after the resize was detected", async () => {
    freshDeck();
    stallAfter((e) => e.message === "redrawing at the size the user gave the chart");
    const { result, all } = await drive();
    expect(result.detail, "a host that went quiet was filed as a failure of ours").not.toMatch(/^threw:/);
    expect(result.ok).toBe(false);
    expect(result.skipped, result.detail).toBe(true);
    expect(result.blind).toBe(true);
    expect(result.detail).toMatch(/the update had decided to keep 600x360pt \(2x, via members\)/);
    expect(result.detail).toMatch(/during the redraw, so the redraw itself was not observed/);
    // The chart WAS resized and the edit WAS made: the words of the setup stall
    // would be false here.
    expect(result.detail).not.toMatch(/so nothing was asked|never resized/);
    expect(result.detail).toMatch(/PowerPoint did not respond while drawing shapes 1-9 of 9/);
    expect(result.detail).toMatch(/DID come back/);
    expect(scenarioBlame(result)).toBe("not-run");
    expect(selfTestNeedsAttention(all)).toBe(true);
  });

  describe("the words, for each place it can stop", () => {
    const err = new Error("PowerPoint did not respond while drawing shapes 1-9 of 9 (45s)");
    const kept = (data: Record<string, unknown>): TraceEntry[] => [
      { ms: 1, scope: "draw", message: "redrawing at the size the user gave the chart", data },
    ];

    it("in its own setup draw: nothing was asked", () => {
      const said = stalledResize("drawing", err, []);
      expect(said).toMatch(/so nothing was asked/);
      expect(isTimeout(new Error(said)), "the runner would not file this as the host stopping").toBe(true);
    });

    it("in the redraw, once the update had read the resize: the size is kept in the words", () => {
      const said = stalledResize("editing", err, kept({ kept: [600, 360], scale: { x: 2, y: 2 }, via: "members" }));
      expect(said).toMatch(/keep 600x360pt \(2x, via members\)/);
      expect(said).toMatch(/the redraw itself was not observed/);
      expect(isTimeout(new Error(said))).toBe(true);
    });

    it("on one axis, says which", () => {
      const said = stalledResize("editing", err, kept({ kept: [600, 180], scale: { x: 2, y: 1 }, via: "group" }));
      expect(said).toMatch(/\(2x wide and 1x tall, via group\)/);
    });

    it("in the edit, before the update said anything about the size: claims neither that it kept nor that it did not", () => {
      const said = stalledResize("editing", err, []);
      expect(said).toMatch(/had not reported keeping the person's size/);
      expect(said).not.toMatch(/had decided to keep/);
      expect(said).not.toMatch(/nothing was asked/);
      expect(isTimeout(new Error(said))).toBe(true);
    });
  });
});

describe("a host that was already stalling", () => {
  /**
   * The scenario before this one is made sick the way a live one is: every
   * readback times out while it runs, and the host is let go again as soon as it
   * has reported. Armed from the first scenario's verdict so the probe charts the
   * second needs have landed, and cleared from the second's, so the only thing the
   * scenario under test meets is a runner that says the host is in trouble.
   */
  const sickPredecessor = (r: ScenarioResult) => {
    if (r.name === "insert on top of an earlier run") {
      _setReadbackTimeoutForTest(5);
      for (let k = 1; k <= 400; k++) stallSyncOn.add(trips.syncs + k);
    }
    if (r.name === "two slides claiming one slot") {
      stallSyncOn.clear();
      _setReadbackTimeoutForTest(90_000);
    }
  };

  it("declines without reading or drawing anything, and says so", async () => {
    const deck = freshDeck();
    setTracing(true);
    const { result, all } = await drive(sickPredecessor);
    const before = all.find((r) => r.name === "two slides claiming one slot")!;
    expect(hostSeemsSick(before, false), "the scenario before it was not sick, so this proves nothing").toBe(true);
    expect(result.ok).toBe(false);
    expect(result.skipped, result.detail).toBe(true);
    expect(result.blind, "nothing was learned, so the round must still be flagged").toBe(true);
    expect(result.ms).toBe(0);
    expect(result.detail).toMatch(
      /^not attempted — the scenario before this one told the battery the host was in trouble/,
    );
    expect(result.detail).toMatch(/Nothing was drawn and nothing was asked/);
    // Without drawing: no draw traced after the decline, not even the deck scan
    // that would have looked for a slide to draw on, and nothing of the
    // scenario's on the deck. The runner announces a scenario it starts, so a
    // declined one has no `scenario starting` line to be found.
    const entries = traceLog().entries;
    const declined = entries.findIndex(
      (e) => e.message === "declining a scenario — the host was already in trouble" && e.data?.name === NAME,
    );
    expect(declined, "the decline left no trace line").toBeGreaterThanOrEqual(0);
    expect(entries[declined].data?.sick).toBe(1);
    expect(entries.some((e) => e.message === "scenario starting" && e.data?.name === NAME)).toBe(false);
    const after = entries.slice(declined);
    expect(after.filter((e) => e.scope === "draw")).toEqual([]);
    expect(after.filter((e) => e.message === "scanned the deck for charts")).toEqual([]);
    expect(chartConfigs(deck).filter((c) => c.data.series[0]?.name === RESIZE_SPECIMEN.data.series[0].name)).toEqual(
      [],
    );
    expect(scenarioBlame(result)).toBe("not-run");
    expect(selfTestNeedsAttention(all)).toBe(true);
  });

  it("leaves the runner's count of sick scenarios exactly where it found it", async () => {
    // Two scenarios in a row are sick, this one declines, and the next is sick too.
    // What shows whether the decline touched the count is what follows. Left at two,
    // the next one runs and is the third strike, so the run is abandoned from the
    // scenario after it. Had the decline counted, the next would not have run at
    // all; had it reset the count, the scenario after that would have run.
    freshDeck();
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    const arm = () => {
      _setReadbackTimeoutForTest(5);
      for (let k = 1; k <= 600; k++) stallSyncOn.add(trips.syncs + k);
    };
    const disarm = () => {
      stallSyncOn.clear();
      _setReadbackTimeoutForTest(90_000);
    };
    const all = await runSelfTest("probe", undefined, (r) => {
      if (r.name === "same scale across the deck") arm();
      if (r.name === PREDECESSOR) disarm();
      if (r.name === NAME) arm();
    });
    const by = (name: string) => all.find((r) => r.name === name)!;
    expect(by("one chart alone on a warm deck").blind, "the first sick scenario was not sick").toBe(true);
    expect(by(PREDECESSOR).blind, "the second sick scenario was not sick").toBe(true);
    expect(by(NAME).detail).toMatch(/^not attempted — the 2 scenarios before this one told the battery/);
    const next = by("explode a degraded picture");
    expect(next.detail, "the decline was counted as a third strike").not.toMatch(/^not reached/);
    expect(hostSeemsSick(next, false), "the scenario after the decline was not sick, so this proves nothing").toBe(
      true,
    );
    expect(by("which selection call wedges the host").detail, "the decline reset the count").toMatch(
      /^not reached — 3 scenarios in a row could not get an answer out of the host/,
    );
  }, 60_000);

  it("does not decline on a host the scenario before it left well", async () => {
    // The control: with nothing wrong upstream it runs, and passes.
    freshDeck();
    const { result } = await drive();
    expect(result.detail).not.toMatch(/not attempted/);
    expect(result.ok, result.detail).toBe(true);
  });
});

describe("what stays a failure", () => {
  it("fails, naming the undone resize, when the edit gives the chart back at the size it was drawn at", async () => {
    freshDeck();
    wrapped.undoTheResize = true;
    const { result, all } = await drive();
    expect(result.skipped, `a lost resize was filed as a skip: ${result.detail}`).toBeFalsy();
    expect(result.ok).toBe(false);
    expect(result.detail).toMatch(/the edit UNDID the resize/);
    expect(scenarioBlame(result)).toBe("ours");
    expect(selfTestNeedsAttention(all)).toBe(true);
  });

  it("fails as a throw of ours when the update throws something that is not the host going quiet", async () => {
    freshDeck();
    wrapped.updateThrows = new Error("the update fell over");
    const { result } = await drive();
    expect(result.skipped).toBeFalsy();
    expect(result.ok).toBe(false);
    expect(result.detail).toMatch(/^threw: the update fell over/);
  });

  it("does not take an error that only mentions a deadline for the host going quiet", async () => {
    freshDeck();
    wrapped.updateThrows = new Error("the update did not respond to the request, which is a bug of ours");
    const { result } = await drive();
    expect(result.skipped).toBeFalsy();
    expect(result.detail).toMatch(/^threw: the update did not respond to the request, which is a bug of ours/);
  });
});

describe("where it runs", () => {
  it("is the neighbour of the other update-path scenario, and `what a chart kind costs` is still last", () => {
    const order = [...ROUTINE_SCENARIO_NAMES];
    expect(order.indexOf(NAME)).toBe(order.indexOf(PREDECESSOR) + 1);
    expect(order.at(-1)).toBe("what a chart kind costs");
  });
});
