// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import {
  CHART_ORIGIN_TAG,
  CHART_PARTS_TAG,
  CHART_SCENE_TAG,
  CHART_TAG,
  canInsertSlidesFromBase64,
  deleteShapesById,
  insertSceneIntoSlide,
  listChartsInDeck,
  replaceSlideWithDeck,
  slideHoldsOnlyChart,
  slideSize,
  updateChartInSlide,
  updateChartsInSlides,
  withSlideDeselected,
  wreckageOf,
} from "../src/render/powerpoint";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { estimateOfficeShapes } from "../src/core/scene";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import { faults, installHost, makeSlide, trips, type FakeSlide } from "./helpers/office-host";

/**
 * The failed-update recovery ladder, one regression per defect.
 *
 * Every test here drives the real functions against the repo's fake host and
 * asserts what the DECK holds afterwards, never what a mock was told — the
 * defects below were all invisible to tests that stopped at a call count.
 *
 * Slide ids are unique per test on purpose: `shapesDrawnOnSlide` is a per-RUN
 * total that is not reset between tests, and the swap gate reads it.
 */

const live = (s: FakeSlide) => s.created.filter((x) => !x.deleted);

/** A 1x1 PNG, for a chart drawn as one picture shape (still named `PowerChart`). */
const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const titleOf = (tag: string | undefined): string | undefined => {
  if (!tag) return undefined;
  try {
    return (JSON.parse(tag) as { title?: string }).title;
  } catch {
    return undefined;
  }
};

/** The titles of the charts a slide holds, by their config tag. */
const chartsOn = (s: FakeSlide): string[] =>
  live(s)
    .map((x) => titleOf(x.tagStore.get(CHART_TAG)))
    .filter((t): t is string => !!t);

/** The same chart with one more category, so an update cannot be answered in place. */
const edited = (base: ChartConfig): ChartConfig => ({
  ...base,
  title: `${base.title} edited`,
  data: {
    categories: [...base.data.categories, "Q5"],
    series: base.data.series.map((s) => ({ ...s, values: [...s.values, 80] })),
  },
});

/**
 * Layer 1 of `updateChartResilient` (src/taskpane/app.ts), up to the swap gate:
 * a redraw whose adds the host refuses, then the sweep of its litter.
 *
 * Reproduced rather than imported because the ladder is module-private to the
 * pane; each call below is the one it makes, in its order.
 */
async function failedRedraw(
  scene: ReturnType<typeof buildChart>,
  target: Parameters<typeof updateChartInSlide>[1],
  tagData: string,
) {
  faults.refuseShapeAdds = true;
  let stall: unknown;
  try {
    await updateChartInSlide(scene, target, { tagData });
  } catch (err) {
    stall = err;
  } finally {
    // The fake's own deck decode goes through `addGeometricShape`; the refusal
    // models the live canvas stalling, which a real `insertSlidesFromBase64`
    // does not.
    faults.refuseShapeAdds = false;
  }
  const wreckage = wreckageOf(stall);
  if (wreckage?.strayIds.length) await deleteShapesById(wreckage.slideId, wreckage.strayIds);
  return { stall, wreckage };
}

/** Layer 2 of the same ladder: ask the gate, and swap the slide if it says yes. */
async function trySwap(slideId: string, scene: ReturnType<typeof buildChart>, tagData: string) {
  const gate = canInsertSlidesFromBase64() && (await slideHoldsOnlyChart(slideId));
  if (!gate) return { gate, swap: undefined };
  const built = await buildDeckBase64([{ scene, title: "Chart", configJson: tagData }], await slideSize());
  return { gate, swap: await replaceSlideWithDeck(slideId, built.base64) };
}

describe("r2-4: the slide swap must not replace a slide that still holds ANOTHER chart", () => {
  it("refuses the swap when the lone shape left is a second chart, and that chart survives", async () => {
    const other = makeSlide("r24-other");
    const two = makeSlide("r24-two");
    const deck = [other, two];
    installHost(deck, [], two);
    const cfgA = { ...sampleConfig("clustered"), title: "Chart A" };
    const cfgB = { ...sampleConfig("line"), title: "Chart B (picture)" };
    const tA = await insertSceneIntoSlide(buildChart(cfgA), {
      slideId: "r24-two",
      tagData: JSON.stringify(cfgA),
      left: 20,
      top: 40,
    });
    // A picture chart is ONE top-level shape named `PowerChart`, so the fake's
    // count and its list agree with no workaround at all.
    await insertSceneIntoSlide(buildChart(cfgB), {
      slideId: "r24-two",
      tagData: JSON.stringify(cfgB),
      pictureBase64: PNG,
      left: 480,
      top: 40,
    });
    expect(tA, "the setup produced no target for chart A").toBeTruthy();

    const next = edited(cfgA);
    const { wreckage } = await failedRedraw(buildChart(next), tA!, JSON.stringify(next));
    expect(wreckage, "layer 1 did not destroy chart A, so this test proves nothing").toBeTruthy();
    expect(
      live(two).map((s) => s.name),
      "the litter was not swept down to chart B",
    ).toEqual(["PowerChart"]);

    const { gate, swap } = await trySwap("r24-two", buildChart(next), JSON.stringify(next));
    expect(gate, "the gate took chart B for the chart being updated and offered to replace the slide").toBe(false);
    expect(swap).toBeUndefined();
    expect(chartsOn(two), "chart B was destroyed by the swap").toEqual(["Chart B (picture)"]);
  });

  it("does the same for a second chart that is a GROUP, which is the ordinary case", async () => {
    const other = makeSlide("r24-other-g");
    const two = makeSlide("r24-two-g");
    installHost([other, two], [], two);
    // A real slide counts top-level shapes; the fake counts a group's members
    // too unless told otherwise, so make its count agree with its list.
    faults.groupHidesChildren = true;
    const cfgA = { ...sampleConfig("clustered"), title: "Chart A" };
    const cfgB = { ...sampleConfig("line"), title: "Chart B" };
    const tA = await insertSceneIntoSlide(buildChart(cfgA), {
      slideId: "r24-two-g",
      tagData: JSON.stringify(cfgA),
      left: 20,
      top: 40,
    });
    await insertSceneIntoSlide(buildChart(cfgB), {
      slideId: "r24-two-g",
      tagData: JSON.stringify(cfgB),
      left: 480,
      top: 40,
    });
    const next = edited(cfgA);
    const { wreckage } = await failedRedraw(buildChart(next), tA!, JSON.stringify(next));
    expect(wreckage, "layer 1 did not destroy chart A").toBeTruthy();
    const hidden = new Set(
      live(two).flatMap((g) => ((g.grouped ?? []) as { id?: string }[]).map((c) => String(c?.id))),
    );
    const topLevel = live(two).filter((s) => !hidden.has(String(s.id)));
    expect(topLevel.map((s) => s.name)).toEqual(["PowerChart"]);
    faults.shapeCountLag = 1;
    faults.shapeCountLagBy = topLevel.length - live(two).length;

    const { gate } = await trySwap("r24-two-g", buildChart(next), JSON.stringify(next));
    expect(gate, "the gate took chart B for the chart being updated").toBe(false);
    expect(chartsOn(two)).toContain("Chart B");
  });

  it("still swaps when chart A was the only thing on the slide and the redraw wrecked it", async () => {
    // The case the fallback exists for: the delete landed, the draw did not, the
    // litter is swept, and the slide holds nothing the user put there.
    const solo = makeSlide("r24-solo");
    installHost([makeSlide("r24-other-s"), solo], [], solo);
    const cfgA = { ...sampleConfig("clustered"), title: "Solo" };
    const tA = await insertSceneIntoSlide(buildChart(cfgA), {
      slideId: "r24-solo",
      tagData: JSON.stringify(cfgA),
      left: 20,
      top: 40,
    });
    const next = edited(cfgA);
    const { wreckage } = await failedRedraw(buildChart(next), tA!, JSON.stringify(next));
    expect(wreckage).toBeTruthy();
    const { gate, swap } = await trySwap("r24-solo", buildChart(next), JSON.stringify(next));
    expect(gate, "the fallback refused the one slide it exists for").toBe(true);
    expect(swap).toBe("swapped");
  });

  it("still swaps when layer 1 failed BEFORE it deleted anything and chart A is the lone shape", async () => {
    // No wreckage: the host stalled on the very first sync, so the chart is
    // still there and the lone `PowerChart` really is the one being updated.
    // Refusing here would kill the swap for the stall it was written for.
    const solo = makeSlide("r24-intact");
    installHost([makeSlide("r24-other-i"), solo], [], solo);
    const cfgA = { ...sampleConfig("clustered"), title: "Intact" };
    // A picture chart: one top-level shape named `PowerChart`, which is what a
    // real slide lists for a chart of any kind.
    const tA = await insertSceneIntoSlide(buildChart(cfgA), {
      slideId: "r24-intact",
      tagData: JSON.stringify(cfgA),
      pictureBase64: PNG,
      left: 20,
      top: 40,
    });
    const next = edited(cfgA);
    faults.failSyncOn = trips.syncs + 1;
    let stall: unknown;
    try {
      await updateChartInSlide(buildChart(next), tA!, { tagData: JSON.stringify(next) });
    } catch (err) {
      stall = err;
    } finally {
      faults.failSyncOn = 0;
    }
    expect(stall, "the first sync did not fail, so layer 1 was not stalled").toBeTruthy();
    expect(wreckageOf(stall), "nothing had been deleted yet").toBeUndefined();
    expect(chartsOn(solo)).toEqual(["Intact"]);
    const { gate } = await trySwap("r24-intact", buildChart(next), JSON.stringify(next));
    expect(gate, "refused the swap for a chart that is intact and alone").toBe(true);
  });

  it("forgets a lost chart once a later update on the slide has run", async () => {
    // The gate's memory of a destroyed chart is about the failure it is being
    // asked about. A later update that works must not leave the slide refusing
    // swaps for good.
    const two = makeSlide("r24-forget");
    installHost([makeSlide("r24-other-f"), two], [], two);
    const cfgA = { ...sampleConfig("clustered"), title: "Chart A" };
    const cfgB = { ...sampleConfig("line"), title: "Chart B (picture)" };
    const tA = await insertSceneIntoSlide(buildChart(cfgA), {
      slideId: "r24-forget",
      tagData: JSON.stringify(cfgA),
      left: 20,
      top: 40,
    });
    await insertSceneIntoSlide(buildChart(cfgB), {
      slideId: "r24-forget",
      tagData: JSON.stringify(cfgB),
      pictureBase64: PNG,
      left: 480,
      top: 40,
    });
    const next = edited(cfgA);
    await failedRedraw(buildChart(next), tA!, JSON.stringify(next));
    expect(await slideHoldsOnlyChart("r24-forget")).toBe(false);

    // Chart B is now the only thing on the slide; update it and nothing fails.
    const found = (await listChartsInDeck()).charts.filter((c) => c.target.slideId === "r24-forget");
    expect(found).toHaveLength(1);
    const cfgB2 = JSON.parse(found[0].configJson) as ChartConfig;
    // As a picture again, so the slide still lists one top-level `PowerChart`.
    const updated = await updateChartInSlide(buildChart(cfgB2), found[0].target, {
      tagData: found[0].configJson,
      pictureBase64: PNG,
    });
    expect(updated, "the clean update gave up").toBeTruthy();
    expect(
      await slideHoldsOnlyChart("r24-forget"),
      "the slide kept refusing the swap after an update on it had worked",
    ).toBe(true);
  });
});

describe("r1-1: a refused resolve is re-read per slide, because shape ids repeat across slides", () => {
  /**
   * Two charts on two slides, optionally carrying the SAME shape id — PowerPoint
   * numbers shapes per slide, and 423 of 469 archived decks hold two `PowerChart`
   * groups on different slides with the same id.
   */
  async function sameScaleOverTwoSlides(sameId: boolean, mode: "groupRead" | "byId") {
    const s1 = makeSlide(`r11-a-${sameId}-${mode}`);
    const s2 = makeSlide(`r11-b-${sameId}-${mode}`);
    installHost([s1, s2], [], s1);
    const cfgA: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [3, 4] },
          { name: "S2", values: [1, 2] },
        ],
      },
    };
    const cfgB: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["X", "Y"], series: [{ name: "T1", values: [10, 20] }] },
    };
    await insertSceneIntoSlide(buildChart(cfgA), { slideId: s1.id, left: 100, top: 50, tagData: JSON.stringify(cfgA) });
    await insertSceneIntoSlide(buildChart(cfgB), {
      slideId: s2.id,
      left: 400,
      top: 300,
      tagData: JSON.stringify(cfgB),
    });
    const gA = s1.created.find((s) => s.type === "group")!;
    const gB = s2.created.find((s) => s.type === "group")!;
    if (sameId) (gB as unknown as { id: string }).id = gA.id;
    const scan = await listChartsInDeck();
    const targets = scan.charts.map((c) => c.target);
    expect(targets, "the scan did not find both charts").toHaveLength(2);
    const before = [s1.created.length, s2.created.length];
    // The refusal that sends the batch down the re-read path: `.group` on a
    // non-group, or a by-id lookup the host will not honour.
    if (mode === "groupRead") faults.refuseGroupReadTimes = 1;
    else faults.refuseShapeById = true;
    try {
      await updateChartsInSlides(
        scan.charts.map((c) => {
          const cfg = JSON.parse(c.configJson) as ChartConfig;
          cfg.scale = { max: 40 };
          return { scene: buildChart(cfg), target: c.target, opts: { tagData: JSON.stringify(cfg) } };
        }),
      );
    } finally {
      faults.refuseGroupReadTimes = 0;
      faults.refuseShapeById = false;
    }
    const origin = (xs: { box: { left: number; top: number } }[]) =>
      xs.length ? [Math.min(...xs.map((s) => s.box.left)), Math.min(...xs.map((s) => s.box.top))] : null;
    const fresh = (s: FakeSlide, n: number) => s.created.slice(n).filter((x) => x.type !== "group");
    return {
      aGone: gA.deleted,
      bGone: gB.deleted,
      groupsA: live(s1).filter((x) => x.type === "group").length,
      groupsB: live(s2).filter((x) => x.type === "group").length,
      newA: origin(fresh(s1, before[0])),
      newB: origin(fresh(s2, before[1])),
    };
  }

  for (const mode of ["groupRead", "byId"] as const) {
    it(`deletes each chart's OWN old group and redraws it in place when the ids collide (${mode})`, async () => {
      const control = await sameScaleOverTwoSlides(false, mode);
      expect(control.aGone && control.bGone, "the control (distinct ids) already misbehaves").toBe(true);
      const got = await sameScaleOverTwoSlides(true, mode);
      expect(got.aGone, "chart A's old group was left on its slide").toBe(true);
      expect(got.bGone, "chart B's old group was left on its slide").toBe(true);
      // However many groups each chart ends with is the host's business (a
      // by-id refusal also refuses the grouping re-read): the collision must not
      // change it.
      expect(got.groupsA, "chart A ended differently than it does with distinct ids").toBe(control.groupsA);
      expect(got.groupsB, "chart B ended differently than it does with distinct ids").toBe(control.groupsB);
      // Each redraw follows ITS OWN chart's live position — not the other
      // slide's, which is what put chart A below the bottom of its slide.
      expect(got.newA).toEqual(control.newA);
      expect(got.newB).toEqual(control.newB);
    });
  }
});

describe("r1-23: a refused resolve still deletes the whole of an ungrouped chart that has a parts list", () => {
  const stacked: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "S1", values: [3, 4] },
        { name: "S2", values: [1, 2] },
      ],
    },
  };
  // A different config so the update cannot be a no-op.
  const changed: ChartConfig = {
    ...stacked,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "S1", values: [5, 1] },
        { name: "S2", values: [2, 6] },
      ],
    },
  };

  it("leaves no part of the old chart on the slide when a by-id lookup refuses the resolve", async () => {
    const slide = makeSlide("r123-single");
    // No grouping: the web host's case, and the one that carries a parts list.
    installHost([slide], [], slide, (v) => v !== "1.8");
    await insertSceneIntoSlide(buildChart(stacked), { tagData: JSON.stringify(stacked) });
    expect(
      slide.created.some((s) => s.type === "group"),
      "the chart is grouped, so nothing can strand",
    ).toBe(false);
    const drawn = live(slide).length;
    const anchor = live(slide).find((s) => s.tagStore.get(CHART_TAG))!;
    const partIds = JSON.parse(anchor.tagStore.get(CHART_PARTS_TAG)!) as string[];
    expect(partIds.length, "the chart carries no parts list, which is the premise").toBeGreaterThan(0);
    const found = (await listChartsInDeck()).charts;
    expect(found).toHaveLength(1);
    expect(found[0].target.partIds?.length).toBe(partIds.length);

    faults.refuseShapeById = true;
    try {
      await updateChartsInSlides([
        { scene: buildChart(changed), target: found[0].target, opts: { tagData: JSON.stringify(changed) } },
      ]);
    } finally {
      faults.refuseShapeById = false;
    }
    const stranded = partIds.filter((id) => live(slide).some((s) => s.id === id));
    expect(stranded, "parts of the old chart were left under the redraw").toEqual([]);
    expect(anchor.deleted).toBe(true);
    expect(live(slide).length, "the chart did not come back the size it went").toBe(drawn);
  });

  it("does not strand an ungrouped chart on another slide when a different chart poisons the shared resolve", async () => {
    const s1 = makeSlide("r123-multi-a");
    const s2 = makeSlide("r123-multi-b");
    // 1.8 supported (the web host), grouping refused: ungrouped charts that
    // queue `.group` at the update.
    installHost([s1, s2], [], s1, () => true);
    faults.refuseGroups = 1000;
    await insertSceneIntoSlide(buildChart(stacked), { tagData: JSON.stringify(stacked), slideId: s1.id });
    await insertSceneIntoSlide(buildChart(stacked), { tagData: JSON.stringify(stacked), slideId: s2.id });
    faults.refuseGroups = 0;
    // A has no parts list — the state a refused id read-back at insert leaves —
    // so it queues `Shape.group`, which throws on a non-group and poisons the
    // resolve it shares with B.
    live(s1)
      .find((s) => s.tagStore.get(CHART_TAG))!
      .tagStore.delete(CHART_PARTS_TAG);
    const bAnchor = live(s2).find((s) => s.tagStore.get(CHART_TAG))!;
    const bParts = JSON.parse(bAnchor.tagStore.get(CHART_PARTS_TAG)!) as string[];
    const bBefore = live(s2).length;
    const found = (await listChartsInDeck()).charts;
    expect(found.map((c) => c.target.partIds?.length ?? 0).sort((a, b) => a - b)).toEqual([0, bParts.length]);

    faults.refuseGroupReadTimes = 1;
    try {
      await updateChartsInSlides(
        found.map((c) => ({
          scene: buildChart(changed),
          target: c.target,
          opts: { tagData: JSON.stringify(changed) },
        })),
      );
    } finally {
      faults.refuseGroupReadTimes = 0;
    }
    const stranded = bParts.filter((id) => live(s2).some((s) => s.id === id));
    expect(stranded, "chart B's old parts were left on its slide").toEqual([]);
    expect(bAnchor.deleted, "chart B's old anchor was left on its slide").toBe(true);
    // The redraw is what is left: the new chart's shapes, and the group the
    // grouping step puts round them — never the old chart's too. By count, not
    // by value: a failing assertion would have vitest print the shapes, and the
    // fake's `group` getter throws on a shape that is not one.
    expect(live(s2).length, "chart B grew a whole chart").toBe(estimateOfficeShapes(buildChart(changed)) + 1);
    expect(bBefore, "the premise: B was drawn whole before the update").toBe(bParts.length + 1);
  });
});

describe("r1-24: the wreckage names the origin the redraw used, not the anchor's position", () => {
  /**
   * Insert a chart at (100,120), optionally drag it, then fail an update after
   * its delete committed. Returns the wreckage and the picture the pane's
   * recovery draws from it — `updateChartResilient`'s floor, call for call.
   */
  async function recoverFromAFailedRedraw(name: string, cfg: ChartConfig, grouped: boolean, drag?: [number, number]) {
    const slide = makeSlide(`r124-${name}-${grouped}-${drag ? "drag" : "still"}`);
    installHost([slide]);
    if (!grouped) faults.refuseGroups = 1000;
    const tagData = JSON.stringify(cfg);
    const scene = buildChart(cfg);
    await insertSceneIntoSlide(scene, { tagData, left: 100, top: 120 });
    faults.refuseGroups = 0;
    const target = (await listChartsInDeck()).charts[0].target;
    const anchor = live(slide).find((s) => s.tagStore.get(CHART_TAG))!;
    if (drag) {
      anchor.left += drag[0];
      anchor.top += drag[1];
    }
    // Without a fingerprint the in-place route refuses and the update redraws.
    anchor.tagStore.delete(CHART_SCENE_TAG);
    faults.refuseShapeAdds = true;
    let stall: unknown;
    try {
      await updateChartInSlide(scene, target, { tagData });
    } catch (err) {
      stall = err;
    } finally {
      faults.refuseShapeAdds = false;
    }
    const wreckage = wreckageOf(stall);
    expect(wreckage, "the redraw did not fail after its delete committed").toBeTruthy();
    if (wreckage!.strayIds.length) await deleteShapesById(wreckage!.slideId, wreckage!.strayIds);
    await insertSceneIntoSlide(scene, {
      tagData,
      pictureBase64: PNG,
      slideId: wreckage!.slideId,
      left: wreckage!.at.left,
      top: wreckage!.at.top,
    });
    const picture = live(slide).find((s) => s.tagStore.get(CHART_TAG))!;
    return { wreckage: wreckage!, picture, scene };
  }

  // Charts whose node 0 is NOT a title at the frame's corner: the anchor sits
  // off the origin, and that offset is the whole defect. A titled chart hides
  // it, which is why only the untitled ones are listed.
  const untitled = (kind: "mekko" | "stacked"): ChartConfig => ({ ...sampleConfig(kind), title: undefined });
  const cases = [
    ["mekko", untitled("mekko"), true],
    ["mekko", untitled("mekko"), false],
    ["stacked", untitled("stacked"), false],
  ] as const;
  for (const [kind, cfg, grouped] of cases) {
    it(`puts the recovered untitled ${kind} (${grouped ? "grouped" : "ungrouped"}) back where the redraw would have drawn it`, async () => {
      const { wreckage, picture } = await recoverFromAFailedRedraw(kind, cfg, grouped);
      // Inserted at (100,120) and never moved: that frame origin is where the
      // replacement belongs, whatever the anchor shape's own box says.
      expect(wreckage.at).toEqual({ left: 100, top: 120 });
      expect([picture.left, picture.top]).toEqual([100, 120]);
      const origin = JSON.parse(picture.tagStore.get(CHART_ORIGIN_TAG)!) as number[];
      expect(origin.slice(0, 2), "the displaced origin was written into the recovered chart's tag").toEqual([100, 120]);
    });
  }

  it("follows a drag made since the chart was tagged", async () => {
    const { wreckage, picture } = await recoverFromAFailedRedraw("stacked", untitled("stacked"), false, [30, 20]);
    expect(wreckage.at).toEqual({ left: 130, top: 140 });
    expect([picture.left, picture.top]).toEqual([130, 140]);
  });
});

describe("r1-25: withSlideDeselected runs the redraw once, and a failure reaches the caller", () => {
  /** Wire a controllable selection onto the fake host. */
  function withSelection(ctx: ReturnType<typeof installHost>, initial: string[]) {
    const state = { selected: [...initial], sets: [] as string[][] };
    const p = ctx.presentation as unknown as Record<string, unknown>;
    p.getSelectedSlides = () => ({ items: state.selected.map((id) => ({ id })), load() {} });
    p.setSelectedSlides = (ids: string[]) => {
      state.sets.push([...ids]);
      state.selected = [...ids];
    };
    return state;
  }

  it("does not run a failing fn a second time on the live canvas, and rethrows the first failure", async () => {
    const ctx = installHost([makeSlide("r125-min")]);
    withSelection(ctx, ["r125-min"]);
    // A one-slide deck whose scratch slide the host swallows: the live canvas
    // is the only surface left, so fn runs once on it.
    faults.swallowAdds = 1;
    const first = new Error("first failure");
    const calls: boolean[] = [];
    let outcome: unknown;
    try {
      outcome = await withSlideDeselected(["r125-min"], async (deselected) => {
        calls.push(deselected);
        if (calls.length === 1) throw first;
        return "second attempt";
      });
    } catch (err) {
      outcome = err;
    }
    expect(calls, "fn was run again after it had already failed").toEqual([false]);
    expect(outcome, "the first failure was swallowed").toBe(first);
  });

  it("does the same when the scratch slide lands but the host will not look at it", async () => {
    const deck = [makeSlide("r125-show")];
    const ctx = installHost(deck);
    const sel = withSelection(ctx, ["r125-show"]);
    // The park itself is refused, after the scratch slide was added.
    const original = (ctx.presentation as unknown as { setSelectedSlides: (ids: string[]) => void }).setSelectedSlides;
    (ctx.presentation as unknown as { setSelectedSlides: (ids: string[]) => void }).setSelectedSlides = (ids) => {
      if (ids[0] !== "r125-show") throw new Error("the host would not select the scratch slide");
      original(ids);
    };
    const first = new Error("first failure");
    const calls: boolean[] = [];
    let outcome: unknown;
    try {
      outcome = await withSlideDeselected(["r125-show"], async (deselected) => {
        calls.push(deselected);
        if (calls.length === 1) throw first;
        return "second attempt";
      });
    } catch (err) {
      outcome = err;
    }
    expect(calls).toEqual([false]);
    expect(outcome).toBe(first);
    // The scratch slide that would not be looked at was still taken back out.
    expect(deck.map((s) => s.id)).toEqual(["r125-show"]);
    expect(sel.selected).toEqual(["r125-show"]);
  });

  it("hands a redraw that died after destroying the chart to the caller WITH its wreckage", async () => {
    // End to end: the pane's recovery ladder is built to receive this error, and
    // used to receive `null` — "that chart is no longer on the slide" — over a
    // slide with half a chart on it.
    const slide = makeSlide("r125-e2e");
    const ctx = installHost([slide]);
    withSelection(ctx, ["r125-e2e"]);
    const before = sampleConfig("stacked");
    await insertSceneIntoSlide(buildChart(before), { tagData: JSON.stringify(before) });
    const target = (await listChartsInDeck()).charts[0].target;
    const after = edited(before);
    faults.swallowAdds = 1;
    faults.refuseShapeAdds = true;
    let calls = 0;
    let outcome: unknown;
    try {
      outcome = await withSlideDeselected(["r125-e2e"], () => {
        calls++;
        return updateChartInSlide(buildChart(after), target, { tagData: JSON.stringify(after) });
      });
    } catch (err) {
      outcome = err;
    } finally {
      faults.refuseShapeAdds = false;
    }
    expect(calls, "the redraw was attempted twice").toBe(1);
    expect(outcome, "the failure was swallowed and the caller was told the chart is gone").not.toBeNull();
    expect(outcome).toBeInstanceOf(Error);
    expect(wreckageOf(outcome), "the error lost what the redraw destroyed").toBeTruthy();
  });
});
