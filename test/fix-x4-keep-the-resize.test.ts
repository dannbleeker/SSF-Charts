// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  CHART_TAG,
  insertSceneIntoSlide,
  keepResizedSize,
  listChartsInDeck,
  updateChartInSlide,
  updateChartsInSlides,
  type EditTarget,
  type SeenOnTheSlide,
} from "../src/render/powerpoint";
import { setTracing, traceLog } from "../src/core/trace";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { liveScaleOfScene, sceneFingerprint } from "../src/core/scene-diff";
import type { Scene } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { faults, installHost, makeSlide, type FakeShape, type FakeSlide } from "./helpers/office-host";

/**
 * x4: an edit after the person resized the chart keeps their size.
 *
 * r2-2 made `tryInPlaceUpdate` decline to write geometry onto a resized chart,
 * and the redraw it fell back to drew the chart at the size it was DRAWN at: the
 * person's resize was silently undone by any edit that moved a bar (Ctrl+Z
 * restored it). The owner's decision was that the edit should keep the size.
 *
 * Every assertion below is about the SLIDE afterwards, in the units a person
 * sees: how big the group is, whether the bars are the size a layout at that
 * size gives them, whether it is still one group. None pins which route the
 * update took, except where the point is that an unresized chart still takes
 * the fast one.
 */

afterEach(() => {
  vi.unstubAllGlobals();
  faults.refuseGroups = 0;
  setTracing(false);
});

const clustered = (): ChartConfig => JSON.parse(JSON.stringify(sampleConfig("clustered"))) as ChartConfig;
const edited = (cfg: ChartConfig, change: (c: ChartConfig) => void): ChartConfig => {
  const next = JSON.parse(JSON.stringify(cfg)) as ChartConfig;
  change(next);
  return next;
};
/** The same chart, asked for at another size. */
const sizedTo = (cfg: ChartConfig, width: number, height: number): ChartConfig =>
  edited(cfg, (c) => {
    c.width = width;
    c.height = height;
  });

const FRAME = { left: 60, top: 90 };

async function drawGrouped(cfg: ChartConfig) {
  const slide = makeSlide("s1");
  installHost([slide], [], slide, () => true);
  await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), ...FRAME });
  const group = slide.created.find((s) => s.type === "group" && !s.deleted)!;
  expect(group, "the fixture did not group").toBeTruthy();
  return { slide, group, members: group.grouped as FakeShape[] };
}

async function drawLoose(cfg: ChartConfig) {
  const slide = makeSlide("s1");
  installHost([slide], [], slide, () => true);
  faults.refuseGroups = 99;
  await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), ...FRAME });
  return slide;
}

/** A real group's box is the bounding box of its members, whenever one moves. */
function groupBoxFollowsMembers(group: FakeShape): void {
  const kids = () => (group.grouped as FakeShape[]).filter((k) => !k.deleted);
  const edge = (pick: (k: FakeShape) => number, fold: (...n: number[]) => number) => fold(...kids().map(pick));
  const box = {
    left: () => edge((k) => k.left, Math.min),
    top: () => edge((k) => k.top, Math.min),
    width: () => edge((k) => k.left + k.width, Math.max) - box.left(),
    height: () => edge((k) => k.top + k.height, Math.max) - box.top(),
  };
  for (const [key, read] of Object.entries(box))
    Object.defineProperty(group, key, { get: read, set: () => {}, configurable: true });
}

/** A user resizing the group with PowerPoint's handles: every member scales about the group's corner. */
function userResizes(group: FakeShape, sx: number, sy: number): void {
  const gl = group.left;
  const gt = group.top;
  for (const k of group.grouped as FakeShape[]) {
    k.left = gl + (k.left - gl) * sx;
    k.top = gt + (k.top - gt) * sy;
    k.width = k.width * sx;
    k.height = k.height * sy;
  }
  groupBoxFollowsMembers(group);
}

/**
 * A host that takes width and height writes and answers nothing when asked for
 * them — the proxy that loads and comes back without the property, which the
 * fake models for `id` (`shapeIdNeverPopulates`) and for rotated shapes
 * (`rotationBlindsTheRead`) and not for sizes.
 */
function goSilentAboutSizes(shapes: FakeShape[]): void {
  for (const s of shapes)
    for (const side of ["width", "height"] as const)
      Object.defineProperty(s, side, { get: () => undefined, set: () => {}, configurable: true });
}

const live = (slide: FakeSlide) => slide.created.filter((s) => !s.deleted && s.type !== "group");
const liveGroups = (slide: FakeSlide) => slide.created.filter((s) => !s.deleted && s.type === "group");
const byName = (slide: FakeSlide, name: string): FakeShape => {
  const found = live(slide).filter((s) => s.name === name);
  expect(found, `${name} is on the slide ${found.length} times`).toHaveLength(1);
  return found[0];
};

/** Every rect and text node of `scene` against the shape named for it, at the frame it was drawn at. */
function wrongShapes(slide: FakeSlide, scene: Scene, at = FRAME): string[] {
  const off: string[] = [];
  for (const n of scene.nodes) {
    if (n.kind !== "rect" && n.kind !== "text") continue;
    const s = byName(slide, n.name!);
    const floor = n.kind === "rect" ? 0.2 : 4;
    const want = [at.left + n.x, at.top + n.y, Math.max(floor, n.w), Math.max(floor, n.h)];
    const got = [s.left, s.top, s.width, s.height];
    if (want.some((v, i) => Math.abs(v - got[i]) > 0.5))
      off.push(`${n.name}: want ${want.map((v) => v.toFixed(1))} got ${got.map((v) => v.toFixed(1))}`);
  }
  return off;
}

const routeTaken = () =>
  traceLog().entries.some((e) => e.message === "updated only the shapes that changed") ? "in place" : "redraw";
const keptTrace = () => traceLog().entries.filter((e) => e.message === "redrawing at the size the user gave the chart");
const storedConfig = (shape: FakeShape): ChartConfig => JSON.parse(shape.tagStore.get(CHART_TAG)!) as ChartConfig;

/** The pane's side of the contract: take over the size the update says it kept. */
const adopt = (cfg: ChartConfig, back: EditTarget | null): ChartConfig => (back?.size ? { ...cfg, ...back.size } : cfg);

const aValueEdit = (cfg: ChartConfig) => edited(cfg, (c) => (c.data.series[0].values[1] = 30));
const anAddedRow = (cfg: ChartConfig) =>
  edited(cfg, (c) => {
    c.data.categories.push("Q5");
    for (const s of c.data.series) s.values.push(12);
  });

describe("x4: a redraw after a resize comes out at the person's size", () => {
  it("draws a chart the user doubled at twice the size, in one piece", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const drawnWidth = group.width;
    const drawnHeight = group.height;
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    // One group, holding every shape the chart has, and nothing of the old one.
    const groups = liveGroups(slide);
    expect(groups, "the redraw did not leave exactly one group").toHaveLength(1);
    // By id: `addGroup` is handed handles, not the shapes the slide lists.
    const memberIds = new Set((groups[0].grouped as FakeShape[]).filter((m) => !m.deleted).map((m) => m.id));
    expect(
      live(slide)
        .filter((s) => !memberIds.has(s.id))
        .map((s) => s.name),
      "shapes of the chart are loose beside its group",
    ).toEqual([]);
    expect(memberIds.size).toBe(buildChart(next).nodes.length);

    // At the size the person made it: the group is twice the drawn one, within the
    // couple of points the ink margin does not scale by.
    expect(Math.abs(groups[0].width - 2 * drawnWidth), `group is ${groups[0].width} wide`).toBeLessThan(2);
    expect(Math.abs(groups[0].height - 2 * drawnHeight), `group is ${groups[0].height} tall`).toBeLessThan(2);
    // …and laid out for that size, with the edit in it, not stretched about.
    expect(wrongShapes(slide, buildChart(sizedTo(next, 960, 600)))).toEqual([]);
    const edit = byName(slide, "seg-0-1");
    const neighbour = byName(slide, "seg-0-0");
    expect(Math.abs(edit.height / 30 - neighbour.height / 42), "the edited bar is on another scale").toBeLessThan(0.05);
    expect(keptTrace(), "the redraw did not say it kept the size").toHaveLength(1);

    // The config on the slide is the config it was drawn from, size included, and the
    // caller is told, so its own copy can take it over.
    expect(storedConfig(groups[0])).toMatchObject({ width: 960, height: 600 });
    expect(back?.size).toEqual({ width: 960, height: 600 });
    expect(groups[0].tagStore.get(CHART_TAG)).toBe(JSON.stringify(sizedTo(next, 960, 600)));
  });

  it("lets a second edit work, in place, once the caller has taken the size over", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const first = aValueEdit(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    const back = await updateChartInSlide(buildChart(first), target, { tagData: JSON.stringify(first) });
    const redrawn = liveGroups(slide)[0];
    groupBoxFollowsMembers(redrawn);
    const ids = (redrawn.grouped as FakeShape[]).map((m) => m.id);

    // The pane's copy now carries the size; an unrelated value changes.
    setTracing(true);
    const second = edited(adopt(first, back), (c) => (c.data.series[1].values[2] = 11));
    const again = await updateChartInSlide(buildChart(second), back!, { tagData: JSON.stringify(second) });
    expect(routeTaken(), "the second edit could not be written in place").toBe("in place");
    expect(again?.size, "nothing was resized, so nothing is adopted").toBeUndefined();
    expect(liveGroups(slide)).toHaveLength(1);
    expect(liveGroups(slide)[0]).toBe(redrawn);
    expect((redrawn.grouped as FakeShape[]).map((m) => m.id)).toEqual(ids);
    expect(wrongShapes(slide, buildChart(second))).toEqual([]);
    expect(storedConfig(redrawn)).toMatchObject({ width: 960, height: 600 });
  });

  it("lets a second edit work when the chart is loaded afresh from the slide", async () => {
    // The pane's other route in: select the chart, Edit selected chart. The size is
    // in the config tag, so the config it loads is already the person's.
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const first = aValueEdit(cfg);
    await updateChartInSlide(buildChart(first), (await listChartsInDeck()).charts[0].target, {
      tagData: JSON.stringify(first),
    });
    const loaded = (await listChartsInDeck()).charts[0];
    const reloaded = JSON.parse(loaded.configJson) as ChartConfig;
    expect([reloaded.width, reloaded.height]).toEqual([960, 600]);

    setTracing(true);
    const second = edited(reloaded, (c) => (c.data.series[0].values[3] = 33));
    await updateChartInSlide(buildChart(second), loaded.target, { tagData: JSON.stringify(second) });
    expect(routeTaken()).toBe("in place");
    expect(wrongShapes(slide, buildChart(second))).toEqual([]);
  });

  it("re-tags a chart whose first tag write was refused with the config it was DRAWN from", async () => {
    // The settle pass writes the config it is handed back onto the redrawn chart.
    // That has to be the one with the person's size in it: the chart's shapes are
    // at 960 wide, and a tag that said 480 would fail the next edit's fingerprint.
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    faults.refuseTagWrites = 1;
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(back?.lost, "the settle did not run, so this proves nothing about what it writes").toBeUndefined();
    const tagged = slide.created.filter((s) => !s.deleted && s.tagStore.get(CHART_TAG));
    expect(tagged).toHaveLength(1);
    expect(storedConfig(tagged[0])).toMatchObject({ width: 960, height: 600 });
    expect(back?.size).toEqual({ width: 960, height: 600 });
  });

  it("keeps the size through an edit the in-place path declines before it looks at any size", async () => {
    // A new category changes the node count, so `planSceneUpdate` refuses before
    // the size check is reached. It is the edit people make most.
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = anAddedRow(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("redraw");
    expect(back?.size).toEqual({ width: 960, height: 600 });
    expect(wrongShapes(slide, buildChart(sizedTo(next, 960, 600)))).toEqual([]);
    expect(liveGroups(slide)).toHaveLength(1);
  });

  it("keeps a stretch in one direction to that direction", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 1.5, 1);
    const target = (await listChartsInDeck()).charts[0].target;
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(back?.size).toEqual({ width: 720, height: 300 });
    expect(wrongShapes(slide, buildChart(sizedTo(next, 720, 300)))).toEqual([]);
  });

  it("does the same for a chart the host left ungrouped", async () => {
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    const all = live(slide);
    const gl = Math.min(...all.map((s) => s.left));
    const gt = Math.min(...all.map((s) => s.top));
    for (const k of all) {
      k.left = gl + (k.left - gl) * 1.5;
      k.top = gt + (k.top - gt) * 1.5;
      k.width = k.width * 1.5;
      k.height = k.height * 1.5;
    }
    const target = (await listChartsInDeck()).charts[0].target;
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(back?.size).toEqual({ width: 720, height: 450 });
    // Placed wherever the drag put the anchor; the SIZES are the claim.
    const want = buildChart(sizedTo(next, 720, 450));
    for (const n of want.nodes) {
      if (n.kind !== "rect") continue;
      const s = byName(slide, n.name!);
      expect(s.width, n.name).toBeCloseTo(Math.max(0.2, n.w), 1);
      expect(s.height, n.name).toBeCloseTo(Math.max(0.2, n.h), 1);
    }
  });

  it("applies to every chart in a batch, each at its own size", async () => {
    // Same Scale is the batch caller; it loads each chart's config from its tag, so
    // the redraw is what has to keep the size, per chart.
    const a = clustered();
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => true);
    await insertSceneIntoSlide(buildChart(a), { tagData: JSON.stringify(a), left: 40, top: 40 });
    await insertSceneIntoSlide(buildChart(a), { tagData: JSON.stringify(a), left: 40, top: 360 });
    const groups = liveGroups(slide);
    expect(groups).toHaveLength(2);
    for (const g of groups) groupBoxFollowsMembers(g);
    userResizes(groups[0], 1.5, 1.5);
    const charts = (await listChartsInDeck()).charts;
    expect(charts).toHaveLength(2);
    // A scale that moves every bar, so the resized chart has to be redrawn.
    const next = edited(a, (c) => (c.scale = { max: 200 }));
    const back = await updateChartsInSlides(
      charts.map((c) => ({ scene: buildChart(next), target: c.target, opts: { tagData: JSON.stringify(next) } })),
    );
    const sizes = back.map((t) => t.size);
    expect(sizes.filter(Boolean), "only the resized chart should report a kept size").toEqual([
      { width: 720, height: 450 },
    ]);
    expect(liveGroups(slide)).toHaveLength(2);
  });
});

describe("x4: nothing changes for a chart nobody resized, or whose shapes cannot say", () => {
  it("still takes the fast path for a value edit and reports no size", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const ids = (group.grouped as FakeShape[]).map((m) => m.id);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect((group.grouped as FakeShape[]).map((m) => m.id)).toEqual(ids);
    expect(back).toBeTruthy();
    expect("size" in back!, "an unresized chart came back carrying a size").toBe(false);
    expect(keptTrace()).toEqual([]);
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
  });

  it("redraws an unresized chart at the config's size, with the config it was given", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = anAddedRow(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("redraw");
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    // Byte for byte the tag the caller sent.
    expect(liveGroups(slide)[0].tagStore.get(CHART_TAG)).toBe(JSON.stringify(next));
  });

  it("does not take a resize over a size the caller asked for", async () => {
    // The person resized the chart on the slide AND typed a new width in the pane.
    // Two instructions; the explicit one wins, as it did before.
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const asked = sizedTo(aValueEdit(cfg), 600, 400);
    const back = await updateChartInSlide(buildChart(asked), target, { tagData: JSON.stringify(asked) });
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
    expect(wrongShapes(slide, buildChart(asked))).toEqual([]);
    expect(storedConfig(liveGroups(slide)[0])).toMatchObject({ width: 600, height: 400 });
  });

  it("does not call one dragged bar a resize", async () => {
    const cfg = clustered();
    const { slide, group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const bar = members.find((m) => m.name === "seg-1-2")!;
    bar.width *= 1.5;
    bar.height *= 1.5;
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("redraw");
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
  });

  it("behaves as before on a host that will not say how big its shapes are", async () => {
    // Positive evidence only: no size read means no factor, so this chart is
    // handled exactly as it was before sizes were looked at. Here that is the
    // fast path, with whatever it writes at the config's size.
    const cfg = clustered();
    const { group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    goSilentAboutSizes(members);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
    const said = traceLog().entries.find((e) => e.message === "updated only the shapes that changed")!;
    expect((said.data as { checked: { sizes: number } }).checked.sizes, "sizes were compared on a silent host").toBe(0);
  });

  it("redraws at the config's size, as before, when a silent host meets an edit that needs a redraw", async () => {
    const cfg = clustered();
    const { slide, group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    goSilentAboutSizes(members);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = anAddedRow(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("redraw");
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
  });

  it("keeps the stretched chart for an edit that writes no geometry, as r2-2 decided", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 1.5, 1.5);
    const wide = byName(slide, "seg-1-2").width;
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.title = "Renamed"));
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(byName(slide, "seg-1-2").width).toBeCloseTo(wide, 6);
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
  });
});

describe("liveScaleOfScene", () => {
  const scene = buildChart(clustered());
  type Read = { name?: string; width?: number; height?: number };
  const exact: Read[] = scene.nodes.map((n) => {
    const box = n as { name?: string; w?: number; h?: number };
    return { name: n.name, width: box.w, height: box.h };
  });
  const scaled = (sx: number, sy: number, keep: (i: number) => boolean = () => true): Read[] =>
    exact.map((l, i) =>
      keep(i)
        ? { ...l, width: (l.width ?? 0) * sx, height: (l.height ?? 0) * sy }
        : { ...l, width: undefined, height: undefined },
    );
  const rectIndexes = scene.nodes.flatMap((n, i) => (n.kind === "rect" ? [i] : []));

  it("reads the factor off every rect, per axis", () => {
    expect(liveScaleOfScene(scene, scaled(2, 2))).toMatchObject({ x: 2, y: 2, samples: rectIndexes.length });
    const wide = liveScaleOfScene(scene, scaled(1.5, 1))!;
    expect(wide.x).toBeCloseTo(1.5, 9);
    expect(wide.y, "an axis that agrees reads exactly 1").toBe(1);
  });

  it("is null for a chart whose shapes are as drawn, and for rounding", () => {
    expect(liveScaleOfScene(scene, scaled(1, 1))).toBeNull();
    expect(
      liveScaleOfScene(
        scene,
        exact.map((l) => ({ ...l, height: (l.height ?? 0) + 0.9 })),
      ),
    ).toBeNull();
  });

  it("is null when the host answered for fewer than two rects, or for none", () => {
    const only = (n: number) => (i: number) => rectIndexes.slice(0, n).includes(i);
    expect(liveScaleOfScene(scene, scaled(2, 2, only(2)))).toMatchObject({ x: 2, samples: 2 });
    expect(liveScaleOfScene(scene, scaled(2, 2, only(1)))).toBeNull();
    expect(
      liveScaleOfScene(
        scene,
        scaled(2, 2, () => false),
      ),
    ).toBeNull();
  });

  it("is null when one rect disagrees with the rest, however the rest agree", () => {
    const stray = scaled(2, 2);
    stray[rectIndexes[3]] = { ...stray[rectIndexes[3]], width: (stray[rectIndexes[3]].width ?? 0) * 1.3 };
    expect(liveScaleOfScene(scene, stray)).toBeNull();
  });

  it("is null when two shapes have traded places, because the pairing is gone", () => {
    const swapped = scaled(2, 2);
    [swapped[3], swapped[4]] = [swapped[4], swapped[3]];
    expect(liveScaleOfScene(scene, swapped)).toBeNull();
  });

  it("is null for a factor no person made", () => {
    expect(liveScaleOfScene(scene, scaled(40, 40))).toBeNull();
    expect(liveScaleOfScene(scene, scaled(0.01, 0.01))).toBeNull();
  });
});

describe("keepResizedSize", () => {
  const cfg = clustered();
  const prev = buildChart(cfg);
  const target = { slideId: "s1", shapeId: "g1", left: 0, top: 0 };
  const itFor = (c: ChartConfig, extra: Record<string, unknown> = {}) => ({
    scene: buildChart(c),
    target,
    opts: { tagData: JSON.stringify(c), ...extra },
  });
  const seen = (x: number, y: number): SeenOnTheSlide => ({ prev, scale: { x, y, samples: 8 } });

  it("builds the scene at the scaled size and stores the config that builds it", () => {
    const kept = keepResizedSize(itFor(cfg), seen(2, 1.5))!;
    expect(kept.size).toEqual({ width: 960, height: 450 });
    expect([kept.it.scene.width, kept.it.scene.height]).toEqual([960, 450]);
    expect(sceneFingerprint(kept.it.scene)).toBe(sceneFingerprint(buildChart(sizedTo(cfg, 960, 450))));
    expect(kept.it.opts?.tagData).toBe(JSON.stringify(sizedTo(cfg, 960, 450)));
    expect(kept.it.target).toBe(target);
  });

  it("rounds to a tenth of a point and leaves an unstretched axis exactly alone", () => {
    const kept = keepResizedSize(itFor(cfg), seen(1.2345678, 1))!;
    expect(kept.size).toEqual({ width: 592.6, height: 300 });
  });

  it("says nothing without evidence", () => {
    expect(keepResizedSize(itFor(cfg), {})).toBeNull();
    expect(keepResizedSize(itFor(cfg), { prev })).toBeNull();
    expect(keepResizedSize(itFor(cfg), { scale: { x: 2, y: 2, samples: 8 } })).toBeNull();
  });

  it("says nothing when the caller asked for a canvas of its own", () => {
    expect(keepResizedSize(itFor(sizedTo(cfg, 600, 400)), seen(2, 2))).toBeNull();
  });

  it("says nothing when the scene is not what the caller's config builds", () => {
    const it = { ...itFor(cfg), scene: buildChart(edited(cfg, (c) => (c.title = "Not what the tag says"))) };
    expect(keepResizedSize(it, seen(2, 2))).toBeNull();
    expect(keepResizedSize({ ...itFor(cfg), opts: { tagData: "{not json" } }, seen(2, 2))).toBeNull();
  });

  it("says nothing for a picture update, whose raster was taken at the old size", () => {
    expect(keepResizedSize(itFor(cfg, { pictureBase64: "AAAA" }), seen(2, 2))).toBeNull();
  });

  it("says nothing below the smallest chart the pane will take", () => {
    expect(keepResizedSize(itFor(cfg), seen(0.1, 1))).toBeNull();
    expect(keepResizedSize(itFor(cfg), seen(1, 0.1))).toBeNull();
    expect(keepResizedSize(itFor(cfg), seen(0.5, 0.5))).not.toBeNull();
  });
});
