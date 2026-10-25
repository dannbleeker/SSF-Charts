// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  CHART_ORIGIN_TAG,
  CHART_SCENE_TAG,
  CHART_TAG,
  insertSceneIntoSlide,
  listChartsInDeck,
  updateChartInSlide,
  updateChartsInSlides,
} from "../src/render/powerpoint";
import { setTracing, traceLog } from "../src/core/trace";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { checkShapesAgainstScene } from "../src/core/scene-diff";
import type { Scene } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { faults, installHost, makeSlide, type FakeShape, type FakeSlide } from "./helpers/office-host";

/**
 * The in-place update path (`tryInPlaceUpdate` and `applyNodeInPlace`) writes
 * only the shapes whose node changed, on the claim that the shapes on the slide
 * are the shapes the stored config draws. Six findings are about that claim
 * being less true than the code assumed, and each section below reproduces one
 * against the repo's fake host.
 *
 * Every assertion is about the SLIDE afterwards, not about which route the
 * update took: a redraw and a correct in-place write are both fine, and a test
 * that pinned the route would pass for the wrong reason the day the fast path
 * learnt to handle one of these cases.
 */

afterEach(() => {
  vi.unstubAllGlobals();
  faults.refuseGroups = 0;
  setTracing(false);
});

const clustered = (): ChartConfig => JSON.parse(JSON.stringify(sampleConfig("clustered"))) as ChartConfig;
/** Re-parsed, so an edit to the copy cannot reach the config the chart was drawn from. */
const edited = (cfg: ChartConfig, change: (c: ChartConfig) => void): ChartConfig => {
  const next = JSON.parse(JSON.stringify(cfg)) as ChartConfig;
  change(next);
  return next;
};

const FRAME = { left: 60, top: 90 };

/**
 * A grouped chart, which is what a healthy host produces and what the bugs below
 * need: a group is a thing a user can resize or reorder from inside.
 */
async function drawGrouped(cfg: ChartConfig) {
  const slide = makeSlide("s1");
  installHost([slide], [], slide, () => true);
  await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), ...FRAME });
  const group = slide.created.find((s) => s.type === "group" && !s.deleted)!;
  expect(group, "the fixture did not group").toBeTruthy();
  return { slide, group, members: group.grouped as FakeShape[] };
}

/** An ungrouped chart, which is what the web host produces for everything it cannot group. */
async function drawLoose(cfg: ChartConfig) {
  const slide = makeSlide("s1");
  installHost([slide], [], slide, () => true);
  faults.refuseGroups = 99;
  await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), ...FRAME });
  return slide;
}

/**
 * What PowerPoint does that the fake does not: a group's left, top, width and
 * height ARE the bounding box of its members, recomputed whenever one moves.
 * The fake fixes them at `addGroup` time, which is why no repo test could see
 * a chart whose box moved under an in-place edit.
 */
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

const live = (slide: FakeSlide) => slide.created.filter((s) => !s.deleted && s.type !== "group");
const byName = (slide: FakeSlide, name: string): FakeShape => {
  const found = live(slide).filter((s) => s.name === name);
  expect(found, `${name} is on the slide ${found.length} times`).toHaveLength(1);
  return found[0];
};

/**
 * Every rect and text node of `scene`, against the shape named for it: where the
 * frame puts it, and how big it is. The definition of "the chart is right",
 * independent of how it got there.
 */
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

describe("r2-2: an edit after the user resized the chart", () => {
  const seg = (slide: FakeSlide, s: number, c: number) => byName(slide, `seg-${s}-${c}`);

  it("never leaves some bars stretched and the edited bar not", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 1.5, 1.5);
    const target = (await listChartsInDeck()).charts[0].target;
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    // The bar's pixels per unit is the whole point: a 30 must be drawn at the
    // same scale as the 42 beside it, whatever that scale is.
    const edit = seg(slide, 0, 1);
    const neighbour = seg(slide, 0, 0);
    const perUnit = (s: FakeShape, v: number) => s.height / v;
    expect(
      Math.abs(perUnit(edit, 30) - perUnit(neighbour, 42)),
      `the edited bar is drawn at ${perUnit(edit, 30).toFixed(2)}pt a unit, its neighbour at ${perUnit(neighbour, 42).toFixed(2)}`,
    ).toBeLessThan(0.05);
    // …and it still sits on the axis its neighbour stands on.
    expect(Math.abs(edit.top + edit.height - (neighbour.top + neighbour.height))).toBeLessThan(1);
  });

  it("refuses to write bar geometry onto a chart whose bars are no longer the size the config draws", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 1.5, 1);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken(), "a stretched chart took the in-place path for a geometry change").toBe("redraw");
    const declined = traceLog().entries.find((e) => e.message === "not updating in place — redrawing instead");
    expect((declined!.data as { why: string }).why).toMatch(/size|resiz|stretch/i);
    // The redraw lands at the size the person made it (1.5 wide, as tall as it
    // was), not at the config's: `test/fix-x4-keep-the-resize.test.ts` is the
    // proof of that decision. What matters here is that it is ONE scale.
    expect(wrongShapes(slide, buildChart(edited(next, (c) => (c.width = 720))))).toEqual([]);
  });

  it("still takes the fast path for an edit that writes no geometry, and keeps the user's resize", async () => {
    const cfg = clustered();
    const { slide, group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 1.5, 1.5);
    const wide = seg(slide, 1, 2).width;
    const ids = members.map((m) => m.id);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.title = "Renamed"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken(), "a retitle was redrawn, throwing the user's resize away").toBe("in place");
    expect(byName(slide, "title").text).toBe("Renamed");
    expect(seg(slide, 1, 2).width, "the retitle undid the user's resize").toBeCloseTo(wide, 6);
    expect(members.map((m) => m.id)).toEqual(ids);
  });

  it("leaves an unresized chart on the fast path (the control)", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
  });

  it("holds for an ungrouped chart too, where the shapes are separate", async () => {
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    // The user selects every shape and drags a corner: each one scales about the
    // selection's top left, which is the first shape's corner.
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
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    const edit = seg(slide, 0, 1);
    const neighbour = seg(slide, 0, 0);
    expect(
      Math.abs(edit.height / 30 - neighbour.height / 42),
      `the edited bar is ${(edit.height / 30).toFixed(2)}pt a unit, its neighbour ${(neighbour.height / 42).toFixed(2)}`,
    ).toBeLessThan(0.05);
  });
});

describe("checkShapesAgainstScene, the check behind r2-2 and r2-14", () => {
  const scene = buildChart(clustered());
  const rects = scene.nodes.filter((n) => n.kind === "rect");
  /** What the host would read back if every shape were exactly as drawn. */
  const asDrawn = scene.nodes.map((n) => {
    const box = n as { name?: string; w?: number; h?: number };
    return { name: n.name, width: box.w, height: box.h };
  });

  it("agrees with a slide that is exactly as drawn, and counts what it compared", () => {
    const got = checkShapesAgainstScene(scene, asDrawn, { sizes: true });
    expect(got.off).toBeNull();
    expect(got.names).toBe(scene.nodes.length);
    expect(got.sizes).toBe(rects.length);
    expect(rects.length).toBeGreaterThan(0);
  });

  it("names the first shape that has the wrong name", () => {
    const swapped = [...asDrawn];
    [swapped[3], swapped[4]] = [swapped[4], swapped[3]];
    const got = checkShapesAgainstScene(scene, swapped, { sizes: false });
    expect(got.off).toMatchObject({ what: "name", index: 3 });
    expect(got.names, "it should stop counting at the first miss").toBe(3);
  });

  it("measures size only when asked to, and only to the point", () => {
    const stretched = asDrawn.map((l) => ({ ...l, height: (l.height ?? 0) * 1.5 }));
    expect(checkShapesAgainstScene(scene, stretched, { sizes: false }).off).toBeNull();
    expect(checkShapesAgainstScene(scene, stretched, { sizes: true }).off?.what).toBe("size");
    const nudged = asDrawn.map((l) => ({ ...l, height: (l.height ?? 0) + 0.9 }));
    expect(
      checkShapesAgainstScene(scene, nudged, { sizes: true }).off,
      "under a point is rounding, not a resize",
    ).toBeNull();
  });

  it("takes silence for no evidence, and reports that it compared nothing", () => {
    const mute = asDrawn.map(() => ({}));
    const got = checkShapesAgainstScene(scene, mute, { sizes: true });
    expect(got).toEqual({ names: 0, sizes: 0, off: null });
  });

  it("does not count a name the host answered with nothing for", () => {
    const blank = asDrawn.map((l) => ({ ...l, name: "" }));
    expect(checkShapesAgainstScene(scene, blank, { sizes: false })).toEqual({ names: 0, sizes: 0, off: null });
  });

  it("takes a name this scene has no node for as no evidence of order", () => {
    const foreign = asDrawn.map((l, i) => ({ ...l, name: `Shape ${i}` }));
    expect(checkShapesAgainstScene(scene, foreign, { sizes: false })).toEqual({ names: 0, sizes: 0, off: null });
    // …but ONE shape carrying another node's name is evidence, whatever the rest say.
    foreign[5] = { ...foreign[5], name: scene.nodes[9].name! };
    expect(checkShapesAgainstScene(scene, foreign, { sizes: false }).off).toMatchObject({ what: "name", index: 5 });
  });
});

describe("r2-14: a node is written to the shape that carries its name", () => {
  it("does not misroute an edit after one member is brought to the front", async () => {
    const cfg = clustered();
    const { slide, group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    // Select category-0 inside the group and Bring to Front: it leaves the
    // slot it was drawn in and becomes the last member.
    const i = members.findIndex((m) => m.name === "category-0");
    expect(i).toBeGreaterThan(0);
    members.push(...members.splice(i, 1));
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    expect(routeTaken(), "a reordered group was written to by position").toBe("redraw");
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    // No bar was handed a label's string and no label a bar's colour.
    expect(live(slide).filter((s) => /^label-/.test(s.name ?? "") && s.text === undefined)).toEqual([]);
  });

  it("refuses a text-only edit after a reorder too, where no size is written for the sizes to catch it", async () => {
    // Renaming the last category changes ONE node, by its string alone. Nothing
    // is measured for it (`sizes` is asked only when geometry is written), so the
    // names are the only evidence — and with category-0 brought to the front the
    // string lands on the first bar.
    const cfg = clustered();
    const { slide, group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    members.push(
      ...members.splice(
        members.findIndex((m) => m.name === "category-0"),
        1,
      ),
    );
    const bar = members.find((m) => m.name === "seg-0-0")!;
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.categories[3] = "ZZZ"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken(), "a string was routed by position through a reordered group").toBe("redraw");
    expect(bar.text, "a bar was handed a category's string").toBeUndefined();
    expect(byName(slide, "category-3").text).toBe("ZZZ");
  });

  it("does not let the damage spread to the next edit", async () => {
    const cfg = clustered();
    const { slide, group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    members.push(
      ...members.splice(
        members.findIndex((m) => m.name === "category-0"),
        1,
      ),
    );
    let current = cfg;
    for (const [s, c, v] of [
      [0, 1, 30],
      [1, 2, 20],
    ] as const) {
      const target = (await listChartsInDeck()).charts[0].target;
      current = edited(current, (x) => (x.data.series[s].values[c] = v));
      await updateChartInSlide(buildChart(current), target, { tagData: JSON.stringify(current) });
    }
    expect(wrongShapes(slide, buildChart(current))).toEqual([]);
  });

  it("trusts the order of a chart whose shapes carry names it does not know, and says it checked none", async () => {
    // The deck builder names nothing (`Shape 3`, `Text 12`), and a bar a person
    // renamed in the Selection Pane is still where it was. A name this scene has
    // no node for is no evidence of order, and refusing on it would send the
    // first edit of every chart inserted that way to a full redraw.
    const cfg = clustered();
    const { group, members } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    members.forEach((m, i) => (m.name = `Shape ${i}`));
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken(), "a chart with default names lost its fast first edit").toBe("in place");
    const said = traceLog().entries.find((e) => e.message === "updated only the shapes that changed")!;
    expect((said.data as { checked: { names: number } }).checked.names).toBe(0);
    // And the edit still landed where the scene puts it, by position.
    buildChart(next).nodes.forEach((n, i) => {
      if (n.kind !== "rect") return;
      expect(members[i].top, `member ${i}`).toBeCloseTo(FRAME.top + n.y, 4);
      expect(members[i].height, `member ${i}`).toBeCloseTo(n.h, 4);
    });
  });

  it("still updates in place when nothing was reordered (the control)", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
  });

  it("reports how many shapes it could check, so a silent host reads as unchecked", async () => {
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].values[1] = 30));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    const said = traceLog().entries.find((e) => e.message === "updated only the shapes that changed")!;
    expect(
      (said.data as { checked?: { names: number } }).checked?.names,
      "the success line does not say what was verified",
    ).toBe(24);
  });
});

describe("U1: a recolour from translucent to opaque clears the old transparency", () => {
  it("clears the fill transparency on a bar that was drawn translucent", async () => {
    const cfg = edited(clustered(), (c) => (c.data.series[1].color = "#1f77b480"));
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const bar = byName(slide, "seg-1-0");
    expect((bar.fill as unknown as { transparency?: number }).transparency).toBeCloseTo(0.498, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[1].color = "#1f77b4"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken(), "the recolour was redrawn, which proves nothing about the fast path").toBe("in place");
    for (let c = 0; c < 4; c++)
      expect(
        (byName(slide, `seg-1-${c}`).fill as unknown as { transparency?: number }).transparency,
        `seg-1-${c} kept its old translucency`,
      ).toBe(0);
  });

  it("clears the outline transparency of a hollow bar the same way", async () => {
    const cfg = edited(clustered(), (c) => {
      c.data.series[0].scenario = "PL";
      c.data.series[0].color = "#2a78d680";
    });
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const line = (n: string) => byName(slide, n).lineFormat as { transparency?: number; color?: string };
    expect(line("seg-0-0").transparency, "the fixture did not draw a translucent outline").toBeCloseTo(0.498, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].color = "#2a78d6"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(line("seg-0-0").transparency, "the outline kept its old translucency").toBe(0);
  });

  it("does not change what a NEW shape is sent: no transparency write when opaque", async () => {
    const cfg = clustered();
    const { slide } = await drawGrouped(cfg);
    for (const s of live(slide)) {
      expect("transparency" in (s.fill as object), `${s.name} was sent a fill transparency on insert`).toBe(false);
      expect("transparency" in s.lineFormat, `${s.name} was sent a line transparency on insert`).toBe(false);
    }
  });
});

describe("U2: a bar that gains an outline in place shows it", () => {
  it("turns the outline on when a series goes from actual to plan", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    expect(byName(slide, "seg-0-0").lineFormat.visible, "an actual bar starts with no outline").toBe(false);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.series[0].scenario = "PL"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken(), "the scenario change was redrawn").toBe("in place");
    for (let c = 0; c < 4; c++) {
      const lf = byName(slide, `seg-0-${c}`).lineFormat;
      expect(lf.visible, `seg-0-${c} kept its outline hidden`).toBe(true);
      expect(lf.weight).toBe(1.5);
    }
  });

  it("turns it back off on the way home", async () => {
    const cfg = edited(clustered(), (c) => (c.data.series[0].scenario = "PL"));
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    const next = edited(cfg, (c) => (c.data.series[0].scenario = "AC"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(byName(slide, "seg-0-0").lineFormat.visible).toBe(false);
  });
});

describe("r1-21: an in-place edit that moves the group's box moves the anchor with it", () => {
  const untitled = (): ChartConfig => edited(clustered(), (c) => (c.title = undefined));
  const raise = (cfg: ChartConfig, s: number, c: number, v: number) =>
    edited(cfg, (x) => (x.data.series[s].values[c] = v));

  it("keeps later edits where the origin puts them", async () => {
    const cfg = untitled();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const top0 = group.top;
    let target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);

    // Raising the tallest bar lifts the group's top edge: the box moved.
    const e1 = raise(cfg, 1, 3, 72.1);
    target = (await updateChartInSlide(buildChart(e1), target, { tagData: JSON.stringify(e1) }))!;
    expect(routeTaken(), "the first edit was redrawn, so the anchor was never at risk").toBe("in place");
    expect(group.top, "the fixture's box did not move").toBeLessThan(top0 - 1);

    // A second, unrelated edit must land on the same frame as the first.
    const e2 = raise(e1, 0, 0, 43);
    await updateChartInSlide(buildChart(e2), target, { tagData: JSON.stringify(e2) });
    expect(wrongShapes(slide, buildChart(e2))).toEqual([]);
  });

  it("writes the new corner to the origin tag and hands it back in the target", async () => {
    const cfg = untitled();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    const e1 = raise(cfg, 1, 3, 72.1);
    const back = (await updateChartInSlide(buildChart(e1), target, { tagData: JSON.stringify(e1) }))!;

    const tag = JSON.parse(group.tagStore.get(CHART_ORIGIN_TAG)!) as number[];
    expect(tag.slice(0, 2), "the frame origin moved").toEqual([FRAME.left, FRAME.top]);
    expect(tag[2]).toBeCloseTo(group.left, 6);
    expect(tag[3], "the tag still names the corner the box had before the edit").toBeCloseTo(group.top, 6);
    expect(back.origin).toMatchObject({ left: FRAME.left, top: FRAME.top });
    expect(back.origin!.anchorTop).toBeCloseTo(group.top, 6);
    expect(back.top).toBeCloseTo(group.top, 6);
  });

  it("does not compound across edits (Same Scale with headroom, then more edits)", async () => {
    const cfg = untitled();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    let current = cfg;
    let target = (await listChartsInDeck()).charts[0].target;
    for (const change of [
      (c: ChartConfig) => (c.scale = { ...(c.scale ?? {}), max: 105 }),
      (c: ChartConfig) => (c.data.series[0].values[0] = 43),
      (c: ChartConfig) => (c.data.series[1].values[2] = 20),
    ]) {
      current = edited(current, change);
      target = (await updateChartInSlide(buildChart(current), target, { tagData: JSON.stringify(current) }))!;
      expect(wrongShapes(slide, buildChart(current))).toEqual([]);
    }
  });

  it("leaves a titled chart alone: nothing moved, so nothing is rewritten", async () => {
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const before = group.tagStore.get(CHART_ORIGIN_TAG);
    const target = (await listChartsInDeck()).charts[0].target;
    const next = raise(cfg, 1, 3, 72.1);
    setTracing(true);
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(group.tagStore.get(CHART_ORIGIN_TAG)).toBe(before);
  });

  it("follows a drag made before the edit, and keeps following it after", async () => {
    const cfg = untitled();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    for (const k of group.grouped as FakeShape[]) {
      k.left += 40;
      k.top += 25;
    }
    let current = cfg;
    let target = (await listChartsInDeck()).charts[0].target;
    const at = { left: FRAME.left + 40, top: FRAME.top + 25 };
    for (const change of [
      (c: ChartConfig) => (c.data.series[1].values[3] = 72.1),
      (c: ChartConfig) => (c.data.series[0].values[0] = 43),
    ]) {
      current = edited(current, change);
      target = (await updateChartInSlide(buildChart(current), target, { tagData: JSON.stringify(current) }))!;
      expect(wrongShapes(slide, buildChart(current), at)).toEqual([]);
    }
  });

  /**
   * For an ungrouped chart the tagged shape IS node 0, so the corner moves only
   * if node 0's own box does. No layout the fast path accepts does that — a
   * longer first category shrinks its font, so its height changes and its corner
   * does not — which is why the branch cannot be reached by an honest edit. The
   * host is made to do it instead: shrinking the label's height drags its top
   * edge 9pt, as a host that re-anchors a text box would.
   */
  it("does the same for an ungrouped chart when its first shape is the one that moves", async () => {
    const cfg = untitled();
    const slide = await drawLoose(cfg);
    let target = (await listChartsInDeck()).charts[0].target;
    const anchor = slide.created.find((s) => s.id === target.shapeId)!;
    let drift = 0;
    const left = Object.getOwnPropertyDescriptor(anchor, "left")!;
    Object.defineProperty(anchor, "left", {
      get: () => (left.get as () => number).call(anchor) + drift,
      set: (v: number) => (left.set as (v: number) => void).call(anchor, v),
      configurable: true,
    });
    let height = anchor.height;
    Object.defineProperty(anchor, "height", {
      get: () => height,
      set: (v: number) => {
        height = v;
        drift = 9;
      },
      configurable: true,
    });
    setTracing(true);
    const e1 = edited(cfg, (c) => (c.data.categories[0] = "A much longer first category label"));
    target = (await updateChartInSlide(buildChart(e1), target, { tagData: JSON.stringify(e1) }))!;
    expect(routeTaken(), "the rename was redrawn, so the anchor was never at risk").toBe("in place");
    expect(drift, "the fixture's host never moved the anchor").toBe(9);
    expect(JSON.parse(anchor.tagStore.get(CHART_ORIGIN_TAG)!)).toEqual([
      FRAME.left,
      FRAME.top,
      anchor.left,
      anchor.top,
    ]);

    const e2 = raise(e1, 0, 0, 43);
    await updateChartInSlide(buildChart(e2), target, { tagData: JSON.stringify(e2) });
    // category-0 is the shape the fixture displaced; every other one is on its frame.
    expect(wrongShapes(slide, buildChart(e2)).filter((m) => !m.startsWith("category-0:"))).toEqual([]);
  });

  it("refuses rather than guess when the host will not say where the corner is after the edit", async () => {
    const cfg = untitled();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    // Silent from here on: the update falls back to the caller's snapshot for
    // where the chart is, and only the read AFTER the writes has no answer.
    Object.defineProperty(group, "left", {
      get: () => {
        throw new Error("PropertyNotLoaded");
      },
      configurable: true,
    });
    setTracing(true);
    const e1 = raise(cfg, 1, 3, 72.1);
    await updateChartInSlide(buildChart(e1), target, { tagData: JSON.stringify(e1) });
    expect(routeTaken(), "an origin tag nobody could re-measure was trusted").toBe("redraw");
    const why = traceLog().entries.find((e) => e.message === "not updating in place — redrawing instead");
    expect((why!.data as { why: string }).why).toMatch(/corner/);
    expect(wrongShapes(slide, buildChart(e1))).toEqual([]);
  });

  it("does the same through the batch entry point", async () => {
    const cfg = untitled();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    const e1 = raise(cfg, 1, 3, 72.1);
    const [back] = await updateChartsInSlides([
      { scene: buildChart(e1), target, opts: { tagData: JSON.stringify(e1) } },
    ]);
    const e2 = raise(e1, 0, 0, 43);
    await updateChartsInSlides([{ scene: buildChart(e2), target: back, opts: { tagData: JSON.stringify(e2) } }]);
    expect(wrongShapes(slide, buildChart(e2))).toEqual([]);
  });
});

describe("r1-17: an in-place edit refreshes the chart's alt text", () => {
  const alt = (s: FakeShape) => s as unknown as { altTextTitle?: string; altTextDescription?: string };

  it("renames a grouped chart's accessible title on a retitle", async () => {
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    expect(alt(group).altTextTitle, "the fixture drew no alt text").toBeTruthy();
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.title = "Renamed"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(alt(group).altTextTitle).toBe(buildChart(next).title);
    expect(alt(group).altTextTitle).toContain("Renamed");
  });

  it("refreshes the description when only a name nobody sees changed", async () => {
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.data.categories[0] = "ZZZ"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(alt(group).altTextDescription).toBe(buildChart(next).desc);
    expect(alt(group).altTextDescription).toContain("ZZZ");
  });

  it("does it for an ungrouped chart, where the text sits on the anchor", async () => {
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    const anchor = slide.created.find((s) => s.id === target.shapeId)!;
    const next = edited(cfg, (c) => (c.title = "Renamed"));
    setTracing(true);
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(alt(anchor).altTextTitle).toBe(buildChart(next).title);
  });

  it("writes none on a host that cannot take it, and still updates", async () => {
    const cfg = clustered();
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.10");
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), ...FRAME });
    const group = slide.created.find((s) => s.type === "group" && !s.deleted)!;
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = edited(cfg, (c) => (c.title = "Renamed"));
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(alt(group).altTextTitle).toBeUndefined();
    expect(group.tagStore.get(CHART_TAG)).toContain("Renamed");
    expect(group.tagStore.get(CHART_SCENE_TAG)).toBeTruthy();
  });
});
