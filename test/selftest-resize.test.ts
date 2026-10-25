// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  RESIZE_SPECIMEN,
  describeSelfTest,
  judgeResizeKept,
  runSelfTest,
  scenarioBlame,
  selfTestNeedsAttention,
  type ResizeReading,
  type ScenarioResult,
} from "../src/taskpane/selftest";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { sceneDrawnExtent } from "../src/core/scene-diff";
import { setTracing, traceLog } from "../src/core/trace";
import type { Scene } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { faults, installHost, makeSlide, type FakeSlide } from "./helpers/office-host";

/**
 * Two switches on the product, for the two cases no fake fault reaches on its own.
 * `lost`: an update that redraws a chart and comes back with a target the host
 * would not name, which the scenario must not read as the chart. `refuseDelete`:
 * the by-id delete the scenario takes its chart away with, which the live host
 * refuses for a shape drawn seconds earlier (369 of 419 sweeps in the archive) —
 * it answers 0 and deletes nothing, the way a refused sweep does. Off, both
 * wrappers are the real functions.
 */
const wrapped = vi.hoisted(() => ({ lost: false, refuseDelete: false }));
vi.mock("../src/render/powerpoint", async (importOriginal) => {
  const real = await importOriginal<typeof import("../src/render/powerpoint")>();
  return {
    ...real,
    updateChartInSlide: async (...args: Parameters<typeof real.updateChartInSlide>) => {
      const back = await real.updateChartInSlide(...args);
      return back && wrapped.lost ? { ...back, lost: "no-config" as const } : back;
    },
    deleteShapesById: async (...args: Parameters<typeof real.deleteShapesById>) =>
      wrapped.refuseDelete ? 0 : real.deleteShapesById(...args),
  };
});

/**
 * `an edit after the person resizes a chart`.
 *
 * The scenario draws a chart in the corner of a slide the deck already had, sets
 * its group to twice its size through Office.js, edits a value through the
 * product's update path, reads back what the host reports at each of three
 * moments, and takes its chart away. Its first run on a real host is the next
 * round, so what these hold down is the reading: that it passes only on positive
 * evidence of one scale, that it says what was missing when it cannot tell, and
 * that it leaves the slide holding what it held.
 */

afterEach(() => {
  vi.unstubAllGlobals();
  setTracing(false);
  wrapped.lost = false;
  wrapped.refuseDelete = false;
});

const NAME = "an edit after the person resizes a chart";

/** The config of every chart still tagged on the deck. */
const chartConfigs = (deck: FakeSlide[]): ChartConfig[] =>
  deck
    .flatMap((s) => s.created)
    .filter((s) => !s.deleted && s.tagStore.get("POWERCHART_CONFIG"))
    .map((s) => JSON.parse(s.tagStore.get("POWERCHART_CONFIG")!) as ChartConfig);

/** The titles of every chart still tagged on the deck. */
const chartTitles = (deck: FakeSlide[]): string[] => chartConfigs(deck).map((c) => String(c.title));

/**
 * The scenario's own charts still on the deck. It draws one with NO TITLE (see
 * `RESIZE_SPECIMEN`), so a filter on a title finds none whether or not it was taken
 * away: it was how these assertions read before the chart lost its title, and
 * `toEqual([])` over that filter is a pass that proves nothing. The series name is
 * the mark the specimen carries.
 */
const resizeCharts = (deck: FakeSlide[]): ChartConfig[] =>
  chartConfigs(deck).filter((c) => c.data.series[0]?.name === RESIZE_SPECIMEN.data.series[0].name);

/** A deck for one run, so the test can read what is left on it. */
const freshDeck = (): FakeSlide[] => {
  const deck = [makeSlide("s1")];
  installHost(deck);
  return deck;
};

const drive = async (): Promise<{ result: ScenarioResult; all: ScenarioResult[] }> => {
  const all = await runSelfTest("probe", NAME);
  return { result: all.find((r) => r.name === NAME)!, all };
};

describe("the scenario, driven against the fake host", () => {
  it("passes on a host that scales a group's members with it, and says the size was kept", async () => {
    const deck = freshDeck();
    faults.groupResizeScalesMembers = true;
    const { result } = await drive();
    expect(result.skipped, result.detail).toBeFalsy();
    expect(result.ok, result.detail).toBe(true);
    expect(result.detail).toMatch(/at one scale and at the size the person made it/);
    expect(result.detail).toMatch(/redrawn/);
    expect(result.detail, "a clean sweep needs no note").not.toMatch(/note, not part of the verdict/);
    // The slide holds what it held: the chart it drew went with it, and the probe
    // charts it drew beside were not taken with it.
    expect(resizeCharts(deck), "the scenario left its chart on the slide").toEqual([]);
    expect(
      chartTitles(deck).filter((t) => /twice/.test(t)),
      "it took the probe charts with its own",
    ).toHaveLength(4);
  });

  it("draws a chart that was really on the deck while it ran, so a clean sweep below is a sweep", async () => {
    // The control for every `resizeCharts(deck)` that expects nothing: the chart is
    // found by its series name, so that name has to be one the chart really carries
    // while it is there. With the sweep refused it stays, and the mark must find it;
    // an assertion that the chart is gone, over a mark the chart never had, passes
    // on a scenario that leaves it behind.
    const deck = freshDeck();
    faults.groupResizeScalesMembers = true;
    wrapped.refuseDelete = true;
    await drive();
    expect(resizeCharts(deck), "the mark finds nothing, so the sweep assertions here are vacuous").toHaveLength(1);
  });

  it("records what the host reported at all three moments, raw, in the trace", async () => {
    freshDeck();
    faults.groupResizeScalesMembers = true;
    setTracing(true);
    await drive();
    const said = traceLog().entries.find((e) => e.message === "an edit after a resize — what the host reported");
    expect(said, "the readings are the point of the first run, and none was recorded").toBeTruthy();
    const data = said!.data as Record<string, { group: { width: unknown }; first: unknown; listed: number }>;
    for (const moment of ["drawn", "resized", "after"]) {
      expect(data[moment].group.width, `${moment}: the group's raw width`).toEqual(expect.any(Number));
      expect(data[moment].first, `${moment}: the first member's raw reading`).toMatchObject({
        name: expect.stringMatching(/^string:/),
      });
      // Every member the chart has, which is the specimen's whole scene: this was
      // "more than ten" when the scenario drew the 24-shape sample chart, and that
      // is a number the fake would also have reported for a group it half-listed.
      expect(data[moment].listed).toBe(buildChart(RESIZE_SPECIMEN).nodes.length);
    }
    // The group is twice what it was, as the host reported it.
    expect((data.resized.group.width as number) / (data.drawn.group.width as number)).toBeCloseTo(2, 1);
    expect(said!.data).toMatchObject({ redrawn: true, kept: { width: 600, height: 360 } });
  });

  it("SKIPS, naming what it saw, when the group resizes and its members do not follow", async () => {
    // The fake's default: a group takes a size write as a change to its own box
    // alone. The host reports 2x over members at their drawn sizes — a finding
    // about the size guard, and not a pass for it.
    installHost([makeSlide("s1")]);
    const { result } = await drive();
    expect(result.ok).toBe(false);
    expect(result.skipped, result.detail).toBe(true);
    expect(result.detail).toMatch(/member sizes do not move with a group resize/);
    expect(scenarioBlame(result)).toBe("not-run");
  });

  it("passes on the group's size alone when the members are silent, and says the one scale is inferred", async () => {
    installHost([makeSlide("s1")]);
    faults.groupResizeScalesMembers = true;
    faults.groupMembersSilentOn = ["name", "width", "height"];
    const { result } = await drive();
    expect(result.ok, result.detail).toBe(true);
    expect(result.detail).toMatch(/one scale is inferred from a full redraw and not read/);
  });

  it("SKIPS as a host that got in the way when the group will not say how big it is", async () => {
    installHost([makeSlide("s1")]);
    faults.groupResizeScalesMembers = true;
    faults.groupMembersSilentOn = ["name", "width", "height"];
    faults.groupSilentOn = ["width", "height"];
    const { result, all } = await drive();
    expect(result.ok).toBe(false);
    expect(result.skipped, result.detail).toBe(true);
    expect(result.blind, "an unreadable group is the host in the way, not a capability gap").toBe(true);
    expect(result.detail).toMatch(/would not report the group's size after the draw/);
    expect(selfTestNeedsAttention(all)).toBe(true);
  });

  it("SKIPS rather than reads a shape that is gone, when the update redraws and the host will not name the new group", async () => {
    const deck = freshDeck();
    faults.groupResizeScalesMembers = true;
    wrapped.lost = true;
    const { result } = await drive();
    expect(result.ok).toBe(false);
    expect(result.skipped, result.detail).toBe(true);
    expect(result.blind, "an unnamed redraw is the host in the way").toBe(true);
    expect(result.detail).toMatch(/would not name the new group \(lost: no-config\)/);
    // And it still took its chart off the slide, by what the slide held before.
    expect(resizeCharts(deck)).toEqual([]);
  });

  it("takes its chart back on a skip and on a failure as well as on a pass", async () => {
    // The default fake: the group resizes and its members do not follow, which ends
    // the scenario at a skip — after the chart was drawn and resized.
    const deck = freshDeck();
    const { result } = await drive();
    expect(result.skipped, result.detail).toBe(true);
    expect(resizeCharts(deck)).toEqual([]);
    expect(chartTitles(deck).filter((t) => /twice/.test(t))).toHaveLength(4);
  });
});

describe("a host that refuses to take the scenario's chart away", () => {
  /**
   * The by-id delete is the call this host refuses for a shape drawn seconds
   * earlier, so on the first live round a leftover chart is the likeliest outcome
   * and not a defect of ours. The cleanup is best effort: what the question
   * answered stays the verdict, and the leftover is a note on it.
   */
  const leftOf = (deck: FakeSlide[]) => resizeCharts(deck);

  it("keeps a PASS a pass, with the leftover in a note", async () => {
    const deck = freshDeck();
    faults.groupResizeScalesMembers = true;
    wrapped.refuseDelete = true;
    const { result, all } = await drive();
    expect(leftOf(deck), "the refusal did not leave the chart behind, so nothing was tested").not.toEqual([]);
    expect(result.skipped, result.detail).toBeFalsy();
    expect(result.ok, result.detail).toBe(true);
    expect(result.detail).toMatch(/at one scale and at the size the person made it/);
    expect(result.detail).toMatch(/note, not part of the verdict: the sweep did not take this scenario's chart off/);
    // And no consumer reads the note as a defect.
    expect(selfTestNeedsAttention(all)).toBe(false);
    expect(describeSelfTest(all)).toMatch(/no defects of ours/);
  });

  it("keeps a SKIP a skip, with the leftover in a note", async () => {
    // The default fake: the group resizes and its members do not follow.
    const deck = freshDeck();
    wrapped.refuseDelete = true;
    const { result, all } = await drive();
    expect(leftOf(deck)).not.toEqual([]);
    expect(result.ok).toBe(false);
    expect(result.skipped, "a refused sweep turned a skip into a failure").toBe(true);
    expect(result.detail).toMatch(/member sizes do not move with a group resize/);
    expect(result.detail).toMatch(/note, not part of the verdict/);
    expect(scenarioBlame(result)).toBe("not-run");
    expect(selfTestNeedsAttention(all)).toBe(false);
  });

  it("keeps a skip that says the host got in the way exactly as blind as the question made it", async () => {
    freshDeck();
    faults.groupResizeScalesMembers = true;
    faults.groupMembersSilentOn = ["name", "width", "height"];
    faults.groupSilentOn = ["width", "height"];
    wrapped.refuseDelete = true;
    const { result } = await drive();
    expect(result.skipped, result.detail).toBe(true);
    expect(result.blind).toBe(true);
    expect(result.detail).toMatch(/would not report the group's size after the draw/);
  });

  it("puts the leftover on the trace, where a round can count it", async () => {
    freshDeck();
    faults.groupResizeScalesMembers = true;
    wrapped.refuseDelete = true;
    setTracing(true);
    await drive();
    const line = traceLog().entries.find((e) => e.message === "an edit after a resize left its chart on the slide");
    expect(line, "the leftover was not traced").toBeTruthy();
    const data = line!.data as { before: number; after: number };
    expect(data.after).toBeGreaterThan(data.before);
  });
});

describe("judgeResizeKept", () => {
  const frame = (w: number, h: number): ChartConfig => ({ ...sampleConfig("clustered"), width: w, height: h });
  const config: Scene = buildChart(frame(300, 180));
  const kept: Scene = buildChart(frame(600, 360));

  /** What the host reports for `scene` with each member's size run through `size`. */
  const reading = (
    scene: Scene,
    size: (i: number, w: number, h: number) => [number, number] = (_i, w, h) => [w, h],
    group = sceneDrawnExtent(scene)!,
  ): ResizeReading => ({
    group: { width: group.width, height: group.height },
    members: scene.nodes.map((n, i) => {
      const w = (n as { w?: number }).w ?? 20;
      const h = (n as { h?: number }).h ?? 20;
      const [width, height] = size(i, w, h);
      return { name: n.name, width, height };
    }),
  });
  const drawn = reading(config);
  const doubled = reading(config, (_i, w, h) => [w * 2, h * 2], {
    width: sceneDrawnExtent(config)!.width * 2,
    height: sceneDrawnExtent(config)!.height * 2,
  });
  const input = (over: Partial<Parameters<typeof judgeResizeKept>[0]>) => ({
    factor: 2,
    drawn,
    resized: doubled,
    after: reading(kept),
    redrawn: true,
    scenes: { config, kept },
    ...over,
  });

  it("passes a redraw at the person's size, laid out for it", () => {
    const v = judgeResizeKept(input({}));
    expect(v.ok, v.detail).toBe(true);
    expect(v.detail).toMatch(/bars matched the layout at that size/);
  });

  it("fails an edit that undid the resize, and says the chart went back to the config's size", () => {
    const v = judgeResizeKept(input({ after: reading(config) }));
    expect(v.ok).toBe(false);
    expect(v.skipped).toBeUndefined();
    expect(v.detail).toMatch(/UNDID the resize/);
  });

  it("fails a mix of two scales, and gives the range", () => {
    // The chimera: one bar written at the config's size beside bars at the person's.
    const rects = config.nodes.flatMap((n, i) => (n.kind === "rect" && n.w >= 8 && n.h >= 8 ? [i] : []));
    const edited = rects[1];
    const after = reading(config, (i, w, h) => (i === edited ? [w, h] : [w * 2, h * 2]), {
      width: sceneDrawnExtent(config)!.width * 2,
      height: sceneDrawnExtent(config)!.height * 2,
    });
    const v = judgeResizeKept(input({ after, redrawn: false }));
    expect(v.ok).toBe(false);
    expect(v.skipped).toBeUndefined();
    expect(v.detail).toMatch(/not at one scale after the edit/);
    expect(v.detail).toMatch(/1\.00x to 2\.00x as wide/);
  });

  it("fails when the members match the person's size and the group does not say so", () => {
    const v = judgeResizeKept(input({ after: reading(kept, undefined, { width: 300, height: 180 }) }));
    expect(v.ok).toBe(false);
    expect(v.detail).toMatch(/the group reports/);
  });

  it("skips, naming it, when the group was not resized at all", () => {
    const v = judgeResizeKept(input({ resized: drawn }));
    expect(v.skipped).toBe(true);
    expect(v.detail).toMatch(/was not resized/);
  });

  it("skips as blind when the group's size was never given", () => {
    const silent: ResizeReading = { group: {}, members: drawn.members };
    for (const over of [{ drawn: silent }, { resized: silent }, { after: silent }]) {
      const v = judgeResizeKept(input(over));
      expect(v.skipped, JSON.stringify(Object.keys(over))).toBe(true);
      expect(v.blind).toBe(true);
    }
  });

  it("skips, and does not pass, an in-place write it cannot read the members of", () => {
    const v = judgeResizeKept(input({ after: { group: doubled.group }, redrawn: false }));
    expect(v.ok).toBe(false);
    expect(v.skipped).toBe(true);
    expect(v.detail).toMatch(/cannot be told from a clean one/);
  });

  it("passes a redraw it cannot read the members of when the group is the person's size, and says so", () => {
    const v = judgeResizeKept(input({ after: { group: doubled.group }, redrawn: true }));
    expect(v.ok).toBe(true);
    expect(v.detail).toMatch(/inferred from a full redraw/);
  });

  it("fails a redraw it cannot read the members of when the group went back", () => {
    const v = judgeResizeKept(input({ after: { group: drawn.group }, redrawn: true }));
    expect(v.ok).toBe(false);
    expect(v.skipped).toBeUndefined();
    expect(v.detail).toMatch(/resize was undone/);
  });

  it("needs three comparable bars before it reads a member list, not one", () => {
    const few: ResizeReading = {
      group: kept ? reading(kept).group : {},
      members: reading(kept).members!.map((m, i) => (i < 2 ? m : { name: m.name })),
    };
    const v = judgeResizeKept(input({ after: few, redrawn: false }));
    expect(v.ok).toBe(false);
    expect(v.skipped, "two bars were taken for a chart").toBe(true);
  });

  /**
   * A host can answer a bar's size at one moment and not at another. The drawn and
   * resized readings were filtered on their own and then paired by position in the
   * filtered lists, so a bar missing from one of them threw a TypeError (a hard
   * FAIL, `threw: ...`) or, when the lists happened to be the same length, held one
   * bar's drawn width against another bar's resized one. Bars of different widths,
   * so that a wrong pairing cannot hide behind equal ones.
   */
  describe("when the host answers a bar's size at one moment and not at the other", () => {
    const BARS = 8;
    const stretch = (s: Scene, by: number): Scene => ({
      ...s,
      width: s.width * by,
      height: s.height * by,
      nodes: s.nodes.map((n) => (n.kind === "rect" ? { ...n, x: n.x * by, y: n.y * by, w: n.w * by, h: n.h * by } : n)),
    });
    const bars: Scene = {
      width: 400,
      height: 200,
      nodes: Array.from({ length: BARS }, (_, i) => ({
        kind: "rect" as const,
        x: i * 40,
        y: 10,
        w: 12 + 3 * i,
        h: 40 + 5 * i,
        fill: "#123456",
        name: `bar-${i}`,
      })),
    };
    const big = stretch(bars, 2);
    const at1 = reading(bars);
    const at2 = reading(bars, (_i, w, h) => [w * 2, h * 2], {
      width: sceneDrawnExtent(bars)!.width * 2,
      height: sceneDrawnExtent(bars)!.height * 2,
    });
    const final = reading(big);
    const call = (over: Partial<Parameters<typeof judgeResizeKept>[0]>) =>
      judgeResizeKept({
        factor: 2,
        drawn: at1,
        resized: at2,
        after: final,
        redrawn: true,
        scenes: { config: bars, kept: big },
        ...over,
      });
    /** The reading with one bar's width not given. */
    const without = (r: ResizeReading, node: number): ResizeReading => ({
      ...r,
      members: r.members!.map((m, i) => (i === node ? { name: m.name, height: m.height } : m)),
    });
    /** The reading with only these bars' sizes given. */
    const only = (r: ResizeReading, nodes: number[]): ResizeReading => ({
      ...r,
      members: r.members!.map((m, i) => (nodes.includes(i) ? m : { name: m.name })),
    });

    it("passes the clean readings, which is the control for the rest", () => {
      const v = call({});
      expect(v.ok, v.detail).toBe(true);
      expect(v.detail).toMatch(/members followed the resize/);
    });

    it("does not throw when the first bar's drawn width is missing, and judges the bars both moments gave", () => {
      const v = call({ drawn: without(at1, 0) });
      expect(v.detail).not.toMatch(/threw|undefined/);
      expect(v.skipped, v.detail).toBeFalsy();
      expect(v.ok, v.detail).toBe(true);
    });

    it("does not throw when the last bar's drawn width is missing", () => {
      const v = call({ drawn: without(at1, BARS - 1) });
      expect(v.skipped, v.detail).toBeFalsy();
      expect(v.ok, v.detail).toBe(true);
    });

    it("does not throw when the host answered the last bar at the drawing and not after the resize", () => {
      const v = call({ resized: without(at2, BARS - 1) });
      expect(v.skipped, v.detail).toBeFalsy();
      expect(v.ok, v.detail).toBe(true);
    });

    it("pairs bars on their node when different bars are missing at the two moments", () => {
      // Both lists are seven long, so the position in each is the only thing that
      // would line them up, and the bars after the first gap sit one place apart.
      const v = call({ drawn: without(at1, 1), resized: without(at2, BARS - 2) });
      expect(v.skipped, `bars were held against the wrong bars: ${v.detail}`).toBeFalsy();
      expect(v.ok, v.detail).toBe(true);
    });

    it("skips, naming what was missing, when fewer than three bars were read at both moments", () => {
      const v = call({ drawn: only(at1, [0, 1, 2]), resized: only(at2, [2, 3, 4, 5]) });
      expect(v.ok).toBe(false);
      expect(v.skipped, v.detail).toBe(true);
      expect(v.blind).toBeFalsy();
      expect(v.detail).toMatch(/3 bars' sizes when the chart was drawn and 4 after the resize/);
      expect(v.detail).toMatch(/only 1 of them at both moments/);
    });

    it("still skips a group whose members really do not follow it, over the bars both moments gave", () => {
      const stuck: ResizeReading = { ...at2, members: at1.members };
      const v = call({ resized: without(stuck, 0) });
      expect(v.skipped, v.detail).toBe(true);
      expect(v.detail).toMatch(/member sizes do not move with a group resize/);
    });
  });
});
