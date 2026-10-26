// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import JSZip from "jszip";
import {
  CHART_ORIGIN_TAG,
  CHART_TAG,
  insertSceneIntoSlide,
  listChartsInDeck,
  updateChartInSlide,
  updateChartsInSlides,
  wreckageOf,
  type EditTarget,
} from "../src/render/powerpoint";
import { setTracing, traceLog } from "../src/core/trace";
import { buildChart } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { arrowheadBox } from "../src/core/geometry";
import { sceneDrawnExtent, sceneInkOffset } from "../src/core/scene-diff";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import type { Scene } from "../src/core/scene";
import type { ChartConfig, ChartKind } from "../src/core/types";
import {
  FRAME,
  byName,
  clustered,
  drawGrouped,
  edited,
  groupBoxFollowsMembers,
  liveGroups,
} from "./helpers/group-resize";
import { faults, installHost, makeSlide, type FakeShape, type FakeSlide } from "./helpers/office-host";

/**
 * v4: a deck the OLD generator wrote moves its charts on their first edit.
 *
 * Until 0e9500f the add-in's deck builder stored the frame origin as the
 * `POWERCHART_ORIGIN` anchor too, `[dx, dy, dx, dy]`, where the anchor is meant to
 * be the GROUP's own corner (the frame origin plus the scene's ink offset). The
 * update redraws at `origin + (live - anchor)`, so for a chart whose ink does not
 * start at the frame origin the first edit read the ink offset as a drag. v0.6.1
 * carries that writer, so its decks still hold those tags.
 *
 * Every assertion below is about the SLIDE afterwards (where the shapes are), not
 * about which function mended what. The legacy tag is made the way the old writer
 * made it: the chart is drawn by the current code, which measures the group, and
 * the tag is then overwritten with the frame origin twice.
 */

afterEach(() => {
  vi.unstubAllGlobals();
  faults.refuseGroups = 0;
  faults.refuseShapeAdds = false;
  setTracing(false);
});

const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const untitled = (kind: ChartKind = "clustered"): ChartConfig => ({ ...sampleConfig(kind), title: undefined });
/** One bar moved, which writes geometry in place and leaves the ink offset alone. */
const aValueEdit = (cfg: ChartConfig) => edited(cfg, (c) => (c.data.series[0].values[0] = 43));
/** A row more: a different node count, so the update has to redraw. */
const aRowMore = (cfg: ChartConfig) =>
  edited(cfg, (c) => {
    c.data.categories.push("Q5");
    for (const s of c.data.series) s.values.push(12);
  });

const legacyTag = (at = FRAME) => JSON.stringify([at.left, at.top, at.left, at.top]);

/** Where the live chart's group sits. One group, whichever route the update took. */
function theGroup(slide: FakeSlide): FakeShape {
  const groups = liveGroups(slide);
  expect(groups, "the slide holds more or fewer than one chart").toHaveLength(1);
  return groups[0];
}

/**
 * A chart on a slide, with the tag the old generator wrote: the frame origin
 * as the anchor. The group itself is where the host puts it, at the frame plus the
 * scene's ink offset, which is the whole defect.
 */
async function drawLegacy(cfg: ChartConfig, tag = legacyTag()) {
  const drawn = await drawGrouped(cfg);
  groupBoxFollowsMembers(drawn.group);
  const ink = { x: drawn.group.left - FRAME.left, y: drawn.group.top - FRAME.top };
  drawn.group.tagStore.set(CHART_ORIGIN_TAG, tag);
  return { ...drawn, ink };
}

/** The target a pane would hold: read back from the slide, as the deck scan does. */
const targetOf = async (): Promise<EditTarget> => (await listChartsInDeck()).charts[0].target;

/** A person dragging the group: every member moves with it, and the box follows. */
function userDrags(group: FakeShape, dx: number, dy: number): void {
  for (const k of group.grouped as FakeShape[]) {
    k.left += dx;
    k.top += dy;
  }
  groupBoxFollowsMembers(group);
}

/** Every rect and text node of `scene` against the shape named for it, at the frame it should have been drawn at. */
function wrongShapes(slide: FakeSlide, scene: Scene, at: { left: number; top: number } = FRAME): string[] {
  const off: string[] = [];
  // A node can only be matched to its shape by a name that names it alone.
  const names = new Map<string, number>();
  for (const n of scene.nodes) if (n.name) names.set(n.name, (names.get(n.name) ?? 0) + 1);
  for (const n of scene.nodes) {
    if (n.kind !== "rect" && n.kind !== "text") continue;
    if (!n.name || names.get(n.name) !== 1) continue;
    const s = byName(slide, n.name);
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
const mended = () =>
  traceLog().entries.filter((e) => e.message === "a legacy origin anchor, read as the group's own corner");

describe("a legacy tag on a chart whose ink starts inside the frame", () => {
  it("is the fixture's premise: the group sits off the frame origin, and the tag says it does not", async () => {
    const { group, ink } = await drawLegacy(untitled());
    // Without this the tests below prove nothing: a chart whose ink starts AT the
    // frame origin has a legacy tag that is right.
    expect(ink.y, "the untitled chart's ink starts at the frame, so it cannot show the defect").toBeGreaterThan(10);
    expect(JSON.parse(group.tagStore.get(CHART_ORIGIN_TAG)!)).toEqual([FRAME.left, FRAME.top, FRAME.left, FRAME.top]);
  });

  it("stays where it is on the first edit that rewrites a bar in place", async () => {
    const cfg = untitled();
    const { slide, group, ink } = await drawLegacy(cfg);
    const corner = [group.left, group.top];
    setTracing(true);
    const next = aValueEdit(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(routeTaken(), "the edit was redrawn, so the in-place consumers were not exercised").toBe("in place");
    // The edited bar is written at the frame the chart already has, not an ink
    // offset beyond it: every shape sits at FRAME + its node.
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    expect([group.left, group.top], "the chart moved").toEqual(corner);
    expect(group.left - FRAME.left).toBeCloseTo(ink.x, 6);
    expect(group.top - FRAME.top).toBeCloseTo(ink.y, 6);
    expect(mended()).toHaveLength(1);
  });

  it("stays where it is on the first edit that redraws it", async () => {
    const cfg = untitled();
    const { slide, ink } = await drawLegacy(cfg);
    setTracing(true);
    const next = aRowMore(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("redraw");
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    const group = theGroup(slide);
    // The new group is framed where the old one was: the frame, plus the ink offset
    // of what is drawn now (an added row changes neither).
    expect(group.left).toBeCloseTo(FRAME.left + ink.x, 6);
    expect(group.top).toBeCloseTo(FRAME.top + ink.y, 6);
  });

  it("is mended by what the chart DREW, so an edit that adds a title still lands on the frame", async () => {
    const cfg = untitled();
    const { slide } = await drawLegacy(cfg);
    // A title puts the ink at the frame origin: the chart drawn now has offset
    // zero, and the one on the slide had a nonzero one. The corner to read the tag
    // against is the old one (the stored config's), not the new one's.
    const next = edited(cfg, (c) => (c.title = "Now with a title"));
    expect(
      sceneInkOffset(buildChart(next)),
      "the edit left the ink offset alone, so it cannot tell the two apart",
    ).toEqual({
      x: 0,
      y: 0,
    });
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    expect([theGroup(slide).left, theGroup(slide).top]).toEqual([FRAME.left, FRAME.top]);
  });

  it("hands back a tag that is true, so the edit after it stays put too", async () => {
    const cfg = untitled();
    const { slide } = await drawLegacy(cfg);
    const e1 = aRowMore(cfg);
    const back = (await updateChartInSlide(buildChart(e1), await targetOf(), { tagData: JSON.stringify(e1) }))!;
    const group = theGroup(slide);
    // The redraw measured the group it made: frame origin, then its own corner.
    const tag = JSON.parse(group.tagStore.get(CHART_ORIGIN_TAG)!) as number[];
    [FRAME.left, FRAME.top, group.left, group.top].forEach((want, k) =>
      expect(tag[k], `tag[${k}]`).toBeCloseTo(want, 6),
    );
    expect(back.origin!.left).toBeCloseTo(FRAME.left, 6);
    expect(back.origin!.top).toBeCloseTo(FRAME.top, 6);
    expect(back.origin!.anchorLeft).toBeCloseTo(group.left, 6);
    expect(back.origin!.anchorTop).toBeCloseTo(group.top, 6);
    const e2 = aValueEdit(e1);
    await updateChartInSlide(buildChart(e2), back, { tagData: JSON.stringify(e2) });
    expect(wrongShapes(slide, buildChart(e2))).toEqual([]);
  });

  it("keeps the same answer for every edit while the tag on the slide is still the old one", async () => {
    const cfg = untitled();
    const { slide, group, ink } = await drawLegacy(cfg);
    const e1 = aValueEdit(cfg);
    // The pane's own target, handed back by the update, and a fresh read of the
    // slide: two ways into the next edit, over a tag that nothing has rewritten.
    const back = (await updateChartInSlide(buildChart(e1), await targetOf(), { tagData: JSON.stringify(e1) }))!;
    expect(group.tagStore.get(CHART_ORIGIN_TAG), "the edit rewrote a tag it had no cause to").toBe(legacyTag());
    const e2 = edited(e1, (c) => (c.data.series[1].values[1] = 31));
    await updateChartInSlide(buildChart(e2), back, { tagData: JSON.stringify(e2) });
    expect(wrongShapes(slide, buildChart(e2))).toEqual([]);
    const e3 = edited(e2, (c) => (c.data.series[0].values[2] = 29));
    await updateChartInSlide(buildChart(e3), await targetOf(), { tagData: JSON.stringify(e3) });
    expect(wrongShapes(slide, buildChart(e3))).toEqual([]);
    expect(group.left - FRAME.left).toBeCloseTo(ink.x, 6);
    expect(group.top - FRAME.top).toBeCloseTo(ink.y, 6);
  });

  it("is rewritten true by a first edit that lifts the group's own box, and holds after it", async () => {
    const cfg = untitled();
    const { slide, group } = await drawLegacy(cfg);
    const top0 = group.top;
    setTracing(true);
    // The tallest bar raised: the group's top edge rises with it, so the corner the
    // tag names is no longer the corner the chart has, which is what
    // `cornerAfterEdit` exists to record.
    const e1 = edited(cfg, (c) => (c.data.series[1].values[3] = 72.1));
    const back = (await updateChartInSlide(buildChart(e1), await targetOf(), { tagData: JSON.stringify(e1) }))!;
    expect(routeTaken(), "the edit was redrawn, so the corner check was not exercised").toBe("in place");
    expect(group.top, "the fixture's box did not move").toBeLessThan(top0 - 1);
    expect(wrongShapes(slide, buildChart(e1))).toEqual([]);
    const tag = JSON.parse(group.tagStore.get(CHART_ORIGIN_TAG)!) as number[];
    [FRAME.left, FRAME.top, group.left, group.top].forEach((want, k) =>
      expect(tag[k], `tag[${k}]`).toBeCloseTo(want, 6),
    );
    const e2 = aValueEdit(e1);
    await updateChartInSlide(buildChart(e2), back, { tagData: JSON.stringify(e2) });
    expect(wrongShapes(slide, buildChart(e2))).toEqual([]);
  });

  it("follows a drag made before the first edit, by the drag and no more", async () => {
    const cfg = untitled();
    const { slide, group, ink } = await drawLegacy(cfg);
    userDrags(group, 30, 0);
    const next = aValueEdit(cfg);
    setTracing(true);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(routeTaken()).toBe("in place");
    expect(wrongShapes(slide, buildChart(next), { left: FRAME.left + 30, top: FRAME.top })).toEqual([]);
    // 30 right of where it was: not 30 + the ink offset.
    expect(group.left).toBeCloseTo(FRAME.left + ink.x + 30, 6);
    expect(group.top).toBeCloseTo(FRAME.top + ink.y, 6);
  });

  it("follows a drag when the edit has to redraw, too", async () => {
    const cfg = untitled();
    const { slide, group } = await drawLegacy(cfg);
    userDrags(group, 30, -12);
    const next = aRowMore(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(wrongShapes(slide, buildChart(next), { left: FRAME.left + 30, top: FRAME.top - 12 })).toEqual([]);
  });

  it("is mended for a mekko, whose ink starts BEFORE the frame", async () => {
    const cfg = untitled("mekko");
    const { slide, ink } = await drawLegacy(cfg);
    expect(ink.x, "the mekko fixture has no negative ink offset").toBeLessThan(0);
    const next = aRowMore(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    expect(theGroup(slide).left).toBeCloseTo(FRAME.left + ink.x, 6);
  });

  it("is mended in a deck-wide update, which is how Same Scale reaches a chart", async () => {
    const cfg = untitled();
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => true);
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), ...FRAME });
    const [group] = liveGroups(slide);
    groupBoxFollowsMembers(group);
    group.tagStore.set(CHART_ORIGIN_TAG, legacyTag());
    const next = aRowMore(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    await updateChartsInSlides([{ scene: buildChart(next), target, opts: { tagData: JSON.stringify(next) } }]);
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
  });
});

describe("a legacy chart whose redraw stalls", () => {
  /**
   * The slide swap and the picture fallback put the chart back wherever the
   * wreckage says its frame was (`UpdateWreckage.at`), so a frame read off a
   * legacy tag as an ink offset too far would carry the displacement into the
   * recovery as well, and write it into the recovered chart's own tag.
   */
  it("names the frame it was drawn at, plus the drag, for the recovery to use", async () => {
    for (const drag of [undefined, [30, 20]] as const) {
      const cfg = untitled();
      const { group } = await drawLegacy(cfg);
      if (drag) userDrags(group, drag[0], drag[1]);
      const target = await targetOf();
      faults.refuseShapeAdds = true;
      let stall: unknown;
      const next = aRowMore(cfg);
      try {
        await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
      } catch (err) {
        stall = err;
      } finally {
        faults.refuseShapeAdds = false;
      }
      const wreckage = wreckageOf(stall);
      expect(wreckage, "the redraw did not fail after its delete committed").toBeTruthy();
      expect(wreckage!.at.left).toBeCloseTo(FRAME.left + (drag?.[0] ?? 0), 6);
      expect(wreckage!.at.top).toBeCloseTo(FRAME.top + (drag?.[1] ?? 0), 6);
      vi.unstubAllGlobals();
    }
  });
});

describe("a tag that is not the legacy one is read exactly as before", () => {
  it("leaves a titled chart alone: its ink starts at the frame, so the legacy tag is true", async () => {
    const cfg = clustered();
    const { slide, group } = await drawLegacy(cfg);
    expect(sceneInkOffset(buildChart(cfg))).toEqual({ x: 0, y: 0 });
    expect([group.left, group.top], "a titled chart's group is not at the frame origin").toEqual([
      FRAME.left,
      FRAME.top,
    ]);
    setTracing(true);
    const next = aValueEdit(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    expect(mended(), "a tag that was right was changed").toHaveLength(0);
    expect(group.tagStore.get(CHART_ORIGIN_TAG)).toBe(legacyTag());
  });

  it("still follows a drag of a titled legacy chart by the drag alone", async () => {
    const cfg = clustered();
    const { slide, group } = await drawLegacy(cfg);
    userDrags(group, 30, 0);
    const next = aValueEdit(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(wrongShapes(slide, buildChart(next), { left: FRAME.left + 30, top: FRAME.top })).toEqual([]);
  });

  it("leaves a chart whose anchor was measured exactly as it was, dragged or not", async () => {
    for (const drag of [0, 30]) {
      const cfg = untitled();
      const { slide, group, ink } = await drawGrouped(cfg).then((d) => {
        groupBoxFollowsMembers(d.group);
        return { ...d, ink: { x: d.group.left - FRAME.left, y: d.group.top - FRAME.top } };
      });
      const before = group.tagStore.get(CHART_ORIGIN_TAG);
      const written = JSON.parse(before!) as number[];
      [FRAME.left, FRAME.top, FRAME.left + ink.x, FRAME.top + ink.y].forEach((want, k) =>
        expect(written[k], `the current writer's tag[${k}]`).toBeCloseTo(want, 6),
      );
      expect(written[2], "the fixture's anchor is the frame, so it is a legacy tag").not.toBeCloseTo(written[0], 1);
      if (drag) userDrags(group, drag, 0);
      setTracing(true);
      const next = aValueEdit(cfg);
      await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
      expect(wrongShapes(slide, buildChart(next), { left: FRAME.left + drag, top: FRAME.top })).toEqual([]);
      expect(mended(), `drag ${drag}: a measured anchor was read as a legacy one`).toHaveLength(0);
      expect(group.tagStore.get(CHART_ORIGIN_TAG), `drag ${drag}: the tag was rewritten`).toBe(before);
      setTracing(false);
      vi.unstubAllGlobals();
    }
  });

  it("believes a measured anchor over what the stored config would draw today", async () => {
    // A chart drawn before an engine change, whose stored config now lays out with
    // another ink offset: the config is rewritten to a mekko's here (offset -2, 4.5)
    // while the group on the slide, and the tag measured off it, are the untitled
    // clustered chart's. Only the legacy signature (anchor equal to origin) is
    // mended, so this tag stands and the redraw is framed on the frame it records.
    const cfg = untitled();
    const { slide, group, ink } = await drawGrouped(cfg).then((d) => {
      groupBoxFollowsMembers(d.group);
      return { ...d, ink: { x: d.group.left - FRAME.left, y: d.group.top - FRAME.top } };
    });
    const drifted = untitled("mekko");
    const would = sceneInkOffset(buildChart(drifted))!;
    expect(Math.abs(would.y - ink.y), "the drifted config draws the same offset, so it cannot tell").toBeGreaterThan(
      10,
    );
    group.tagStore.set(CHART_TAG, JSON.stringify(drifted));
    setTracing(true);
    const next = aRowMore(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
    expect(routeTaken(), "the stale fingerprint should have sent this to a redraw").toBe("redraw");
    expect(wrongShapes(slide, buildChart(next))).toEqual([]);
    expect(mended()).toHaveLength(0);
  });

  it("reads a two-number origin and a missing one as it always did: where the group is", async () => {
    for (const tag of ["[60,90]", undefined]) {
      const cfg = untitled();
      const { slide, group, ink } = await drawLegacy(cfg, tag);
      if (tag === undefined) group.tagStore.delete(CHART_ORIGIN_TAG);
      setTracing(true);
      const next = aRowMore(cfg);
      await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) });
      // No recorded frame, so the redraw is framed on the group's corner: the old
      // behaviour, and nothing here is in a position to improve on it.
      expect(wrongShapes(slide, buildChart(next), { left: FRAME.left + ink.x, top: FRAME.top + ink.y })).toEqual([]);
      expect(mended(), `${tag}: there was no anchor to mend`).toHaveLength(0);
      setTracing(false);
      vi.unstubAllGlobals();
    }
  });

  it("does not walk a picture of an untitled chart, whose anchor is the frame by construction", async () => {
    const cfg = untitled("mekko");
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => true);
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), pictureBase64: PNG, ...FRAME });
    const picture = slide.created.find((s) => !s.deleted && s.tagStore.get(CHART_TAG))!;
    expect(picture.type, "the fixture drew a group, not a picture").not.toBe("group");
    expect(JSON.parse(picture.tagStore.get(CHART_ORIGIN_TAG)!)).toEqual([FRAME.left, FRAME.top, FRAME.left, FRAME.top]);
    expect(sceneInkOffset(buildChart(cfg))!.x, "the fixture has no ink offset, so it cannot show the walk").not.toBe(0);
    setTracing(true);
    const next = aValueEdit(cfg);
    await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next), pictureBase64: PNG });
    const now = slide.created.filter((s) => !s.deleted && s.tagStore.get(CHART_TAG));
    expect(now).toHaveLength(1);
    expect([now[0].left, now[0].top], "the picture was moved by the ink offset").toEqual([FRAME.left, FRAME.top]);
    expect(mended()).toHaveLength(0);
  });

  it("keeps the tag of a chart the update cannot read the stored config of", async () => {
    const cfg = untitled();
    const { slide, group, ink } = await drawLegacy(cfg);
    // The config tag the host would not give back: nothing to say what the group
    // on the slide drew, so the tag is read as it came (framing on the live corner,
    // as before) and not guessed at from the config being drawn.
    const target = await targetOf();
    group.tagStore.delete(CHART_TAG);
    setTracing(true);
    const next = aRowMore(cfg);
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(mended()).toHaveLength(0);
    expect(wrongShapes(slide, buildChart(next), { left: FRAME.left + ink.x, top: FRAME.top + ink.y })).toEqual([]);
  });
});

describe("a legacy chart whose scene has a node the host turns", () => {
  /**
   * `sceneDrawnExtent` abstains on these (a wedge is a fan of turned triangles on
   * the live host), and a legacy tag still needs mending for them: the group in a
   * FILE is the deck builder's, which frames a wedge on its circle's box, an
   * arrowhead on `arrowheadBox` and a diagonal line on its bounding box, unturned.
   * The fake host frames its own group differently for these, so the group is put
   * where the deck builder put it, which is the premise being tested.
   */
  const TURNED: ChartKind[] = ["pie", "doughnut", "stacked", "line", "combo", "sunburst", "violin"];

  /** The group a deck file holds for `cfg`: the frame plus the deck builder's ink offset. */
  function asTheDeckFramedIt(group: FakeShape, cfg: ChartConfig): { x: number; y: number } {
    const ink = sceneInkOffset(buildChart(cfg))!;
    for (const [key, value] of [
      ["left", FRAME.left + ink.x],
      ["top", FRAME.top + ink.y],
    ] as const)
      Object.defineProperty(group, key, { get: () => value, set: () => {}, configurable: true });
    return ink;
  }

  it("is mended by the group the deck builder framed, so the first edit lands on the frame", async () => {
    let offset = 0;
    for (const kind of TURNED) {
      const cfg = untitled(kind);
      expect(
        sceneDrawnExtent(buildChart(cfg)),
        `${kind}: the live extent is known, so this is not a turned scene`,
      ).toBeNull();
      const { slide, group } = await drawLegacy(cfg);
      const ink = asTheDeckFramedIt(group, cfg);
      offset += Math.abs(ink.x) + Math.abs(ink.y);
      setTracing(true);
      const next = aValueEdit(cfg);
      await updateChartInSlide(buildChart(next), await targetOf(), { tagData: JSON.stringify(next) }).catch((err) => {
        throw new Error(`${kind}: ${String(err)}`);
      });
      expect(mended(), `${kind}: the legacy tag was not mended`).toHaveLength(1);
      expect(wrongShapes(slide, buildChart(next)), `${kind}: the first edit moved the chart`).toEqual([]);
      setTracing(false);
      vi.unstubAllGlobals();
    }
    expect(offset, "no kind had an ink offset, so the loop proves nothing").toBeGreaterThan(50);
  });
});

describe("sceneInkOffset", () => {
  const KINDS = CHART_KINDS.map((k) => k.kind);

  it("answers for every kind, the ones whose live extent is unknown included", () => {
    let turned = 0;
    for (const kind of KINDS) {
      for (const title of [undefined, "A title"]) {
        const scene = buildChart({ ...sampleConfig(kind), title });
        expect(sceneInkOffset(scene), `${kind}${title ? " (titled)" : ""}: no offset`).not.toBeNull();
        if (sceneDrawnExtent(scene) === null) turned++;
      }
    }
    // Not vacuous: there ARE scenes the extent abstains on, and they are answered.
    expect(turned, "no scene had a turned node, so the claim above is untested").toBeGreaterThan(10);
  });

  it("is the corner of the group the fake host makes, wherever the extent is known", async () => {
    let measured = 0;
    const off: string[] = [];
    for (const kind of KINDS) {
      for (const title of [undefined, "A title"]) {
        const cfg: ChartConfig = { ...sampleConfig(kind), title };
        const want = sceneInkOffset(buildChart(cfg))!;
        if (sceneDrawnExtent(buildChart(cfg)) === null) continue;
        const { group } = await drawGrouped(cfg);
        measured++;
        if (Math.abs(group.left - FRAME.left - want.x) > 1e-6 || Math.abs(group.top - FRAME.top - want.y) > 1e-6)
          off.push(
            `${kind}${title ? " titled" : ""}: scene ${want.x},${want.y} group ${group.left - FRAME.left},${group.top - FRAME.top}`,
          );
        vi.unstubAllGlobals();
      }
    }
    expect(off, "the ink offset is not the corner of the group the live path builds").toEqual([]);
    expect(measured, "almost nothing was measured").toBeGreaterThan(20);
  });

  it("answers a turned node on the box the deck builder frames it on", () => {
    const at = (node: unknown) => sceneInkOffset({ width: 100, height: 100, nodes: [node] } as unknown as Scene);
    expect(at({ kind: "wedge", cx: 50, cy: 60, r: 20, innerR: 0, startAngle: 0, endAngle: 90, fill: "#000" })).toEqual({
      x: 30,
      y: 40,
    });
    expect(
      at({
        kind: "polygon",
        points: [
          { x: 9, y: 30 },
          { x: 40, y: 12 },
          { x: 25, y: 44 },
        ],
      }),
    ).toEqual({ x: 9, y: 12 });
    expect(at({ kind: "line", x1: 80, y1: 10, x2: 20, y2: 70, stroke: "#000" })).toEqual({ x: 20, y: 10 });
    const tip = arrowheadBox(40, 50, 6, 30);
    expect(at({ kind: "arrowhead", x: 40, y: 50, angle: 30, size: 6, fill: "#000" })).toEqual({
      x: tip.left,
      y: tip.top,
    });
    // A polygon of fewer than two points is one the deck builder does not write.
    expect(at({ kind: "polygon", points: [{ x: 1, y: 1 }] })).toBeNull();
  });

  it("is zero for a titled chart, positive for an untitled one, and negative for a mekko", () => {
    expect(sceneInkOffset(buildChart(sampleConfig("clustered")))).toEqual({ x: 0, y: 0 });
    const bare = sceneInkOffset(buildChart(untitled()))!;
    expect(bare.y).toBeGreaterThan(10);
    expect(sceneInkOffset(buildChart(untitled("mekko")))!.x).toBeLessThan(0);
  });

  it("answers the offset of a hand-made scene from its first node on each axis", () => {
    const scene: Scene = {
      width: 100,
      height: 100,
      nodes: [
        { kind: "rect", x: 30, y: 8, w: 20, h: 20, fill: "#000" },
        {
          kind: "text",
          x: 12,
          y: 40,
          w: 30,
          h: 10,
          text: "x",
          fontSize: 10,
          color: "#000",
          align: "left",
          valign: "top",
        },
      ],
    } as unknown as Scene;
    expect(sceneInkOffset(scene)).toEqual({ x: 12, y: 8 });
    expect(sceneInkOffset({ width: 10, height: 10, nodes: [] })).toBeNull();
  });
});

describe("sceneInkOffset against the group the deck builder frames", () => {
  /**
   * The deck builder measures its anchor off the group it makes, and the origin
   * it writes is the frame, so `anchor - origin` IS the offset of that group. The
   * legacy tag was that same group with the frame written twice, which makes this
   * the evidence that mending a legacy tag with `sceneInkOffset` lands on the
   * group the old file really had.
   */
  async function deckOffset(cfg: ChartConfig): Promise<{ x: number; y: number }> {
    const built = await buildDeckBase64([{ scene: buildChart(cfg), title: "t", configJson: JSON.stringify(cfg) }]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    for (const name of Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f))) {
      const part = await zip.file(name)!.async("string");
      const m = /<p:tag name="POWERCHART_ORIGIN" val="([^"]*)"\/>/.exec(part);
      if (m) {
        const [ol, ot, al, at] = JSON.parse(m[1]) as number[];
        return { x: al - ol, y: at - ot };
      }
    }
    throw new Error("the deck carries no origin tag");
  }

  /** Every sample kind at three frames, titled and not, and every chart in the showcase as authored. */
  const SHOWCASE: ChartConfig[] = JSON.parse(readFileSync("examples/showcase.json", "utf8"));
  const CASES: { label: string; cfg: ChartConfig }[] = [
    ...SHOWCASE.map((cfg, i) => ({ label: `showcase #${i} (${cfg.kind})`, cfg })),
    ...CHART_KINDS.flatMap(({ kind }) =>
      [undefined, "A title"].flatMap((title) =>
        (
          [
            [480, 300],
            [200, 150],
            [960, 540],
          ] as const
        ).map(([width, height]) => ({
          label: `${kind} ${width}x${height} ${title ? "titled" : "untitled"}`,
          cfg: { ...sampleConfig(kind), title, width, height },
        })),
      ),
    ),
  ].filter((c) => c.cfg.render !== "image");

  it("equals anchor minus origin for every chart, the turned ones included", async () => {
    let measured = 0;
    let turned = 0;
    const off: string[] = [];
    for (const { label, cfg } of CASES) {
      const want = sceneInkOffset(buildChart(cfg));
      if (!want) {
        off.push(`${label}: no offset`);
        continue;
      }
      const got = await deckOffset(cfg);
      measured++;
      if (sceneDrawnExtent(buildChart(cfg)) === null) turned++;
      // An EMU is 1/12700pt, and the origin is divided by the same constant.
      if (Math.abs(want.x - got.x) > 0.01 || Math.abs(want.y - got.y) > 0.01)
        off.push(`${label}: scene ${want.x},${want.y} deck ${got.x},${got.y}`);
    }
    expect(off, "the scene's ink offset is not the corner of the group the deck builder makes").toEqual([]);
    expect(measured, "almost nothing was measured").toBeGreaterThan(250);
    expect(turned, "no turned scene was measured, so the claim for them is untested").toBeGreaterThan(40);
  }, 120_000);
});
