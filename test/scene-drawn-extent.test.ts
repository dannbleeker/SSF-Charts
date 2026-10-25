// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import JSZip from "jszip";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { groupScaleOfScene, measureGroup, memberSizesRead, sceneDrawnExtent } from "../src/core/scene-diff";
import type { Scene, SceneNode } from "../src/core/scene";
import type { ChartConfig, ChartKind } from "../src/core/types";
import { drawGrouped } from "./helpers/group-resize";

/**
 * The box a chart's shapes fill, derived from the scene alone.
 *
 * `groupScaleOfScene` reads a hand resize off the group's own box when the
 * group's members will not say how big they are, and that is only as good as
 * "the box the scene draws" being the box the host's group takes. So this file is
 * the measurement the choice of margin rests on: the scene-derived extent against
 * the extent the group reports for the same drawn chart, for every chart in the
 * showcase and every sample kind at four frames, on the two drawers this repo
 * has — the fake host's `addGroup` (the live path) and the deck builder's
 * `groupSlideShapes` (the file path, whose `boundsOf` this reuses the idea of).
 */

afterEach(() => vi.unstubAllGlobals());

const SHOWCASE: ChartConfig[] = JSON.parse(readFileSync("examples/showcase.json", "utf8"));
const KINDS = [...new Set(SHOWCASE.map((c) => c.kind))] as ChartKind[];
const FRAMES: [number, number][] = [
  [200, 150],
  [320, 200],
  [480, 300],
  [960, 540],
];

/** Every chart to measure: the showcase as authored, and each sample kind at each frame. */
const CASES: { label: string; cfg: ChartConfig }[] = [
  ...SHOWCASE.map((cfg, i) => ({ label: `showcase #${i} (${cfg.kind})`, cfg })),
  ...KINDS.flatMap((kind) =>
    FRAMES.map(([width, height]) => ({
      label: `${kind} ${width}x${height}`,
      cfg: { ...sampleConfig(kind), width, height },
    })),
  ),
].filter((c) => c.cfg.render !== "image");

/** Whether a node is one the host turns, which the extent refuses to guess a box for. */
const turned = (n: SceneNode): boolean => {
  if (n.kind === "wedge" || n.kind === "polygon" || n.kind === "arrowhead") return true;
  if (n.kind !== "line") return false;
  return Math.abs(n.x2 - n.x1) >= 0.5 && Math.abs(n.y2 - n.y1) >= 0.5 && !(n.dash && n.dash.some((d) => d > 0));
};

describe("sceneDrawnExtent against the group the live path builds", () => {
  it("equals the fake host's group box for every chart it will name a box for", async () => {
    let measured = 0;
    let worst = 0;
    const off: string[] = [];
    for (const { label, cfg } of CASES) {
      const want = sceneDrawnExtent(buildChart(cfg));
      if (!want) continue;
      const { group } = await drawGrouped(cfg);
      measured++;
      const ratio = { x: want.width / group.width, y: want.height / group.height };
      worst = Math.max(worst, Math.abs(ratio.x - 1), Math.abs(ratio.y - 1));
      if (Math.abs(ratio.x - 1) > 1e-9 || Math.abs(ratio.y - 1) > 1e-9)
        off.push(`${label}: drawn ${want.width}x${want.height}, group ${group.width}x${group.height}`);
      vi.unstubAllGlobals();
    }
    expect(off, "the scene-derived box and the group's box part company").toEqual([]);
    // Not vacuous: most of the library is measurable, and the margin below is
    // compared with something. 142 of 223 on the day this was written.
    expect(measured, "almost nothing was measured").toBeGreaterThan(100);
    // The margin `groupScaleOfScene` leaves is the larger of 5pt and 5%. The fake's
    // error is rounding, ~1e-16, and the live host's, where it has been compared,
    // is nil (see `sceneDrawnExtent`): the margin is sized to what a wrong factor
    // costs, not to noise that was seen.
    expect(worst).toBeLessThan(1e-9);
  });

  it("answers null exactly when the scene has a node the host turns", () => {
    const wrong: string[] = [];
    let known = 0;
    let unknown = 0;
    for (const { label, cfg } of CASES) {
      const scene = buildChart(cfg);
      const got = sceneDrawnExtent(scene);
      const expectUnknown = scene.nodes.some(turned);
      if (got === null) unknown++;
      else known++;
      if ((got === null) !== expectUnknown) wrong.push(`${label}: ${got === null ? "abstained" : "named a box"}`);
    }
    expect(wrong).toEqual([]);
    // Both branches are exercised, or the equivalence above is one-sided.
    expect(known).toBeGreaterThan(50);
    expect(unknown).toBeGreaterThan(20);
  });
});

describe("sceneDrawnExtent against the group the file path builds", () => {
  async function deckGroup(cfg: ChartConfig): Promise<{ width: number; height: number } | null> {
    const scene = buildChart(cfg);
    const built = await buildDeckBase64([{ scene, title: "t", configJson: JSON.stringify(cfg) }]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
    const m =
      /<p:grpSp><p:nvGrpSpPr>.*?<p:grpSpPr><a:xfrm><a:off x="-?\d+" y="-?\d+"\/><a:ext cx="(\d+)" cy="(\d+)"/.exec(xml);
    return m ? { width: Number(m[1]) / 12700, height: Number(m[2]) / 12700 } : null;
  }

  it("equals the box `groupSlideShapes` frames its group on, to an EMU", async () => {
    let measured = 0;
    const off: string[] = [];
    for (const { label, cfg } of CASES) {
      const want = sceneDrawnExtent(buildChart(cfg));
      if (!want) continue;
      const made = await deckGroup(cfg);
      expect(made, `${label}: the deck wrote no group`).not.toBeNull();
      measured++;
      // An EMU is 1/12700 of a point; a width of 480 is exact and 213.34 is not.
      if (Math.abs(want.width - made!.width) > 0.001 || Math.abs(want.height - made!.height) > 0.001)
        off.push(`${label}: drawn ${want.width}x${want.height}, deck group ${made!.width}x${made!.height}`);
    }
    expect(off).toEqual([]);
    expect(measured).toBeGreaterThan(100);
  }, 60_000);
});

describe("groupScaleOfScene", () => {
  const scene: Scene = buildChart(sampleConfig("clustered"));
  const drawn = sceneDrawnExtent(scene)!;

  it("is the 480x300 sample's 480x299, which is the box the live host reports for it", () => {
    expect(drawn).toEqual({ width: 480, height: 299 });
  });

  it("is quiet for a chart nobody touched, however the host rounds it", () => {
    expect(groupScaleOfScene(scene, { width: drawn.width, height: drawn.height })).toBeNull();
    expect(groupScaleOfScene(scene, { width: drawn.width + 0.3, height: drawn.height - 0.3 })).toBeNull();
    // 23 of 480 is inside 5%, and 14 of 299 is inside 5pt-or-5% too.
    expect(groupScaleOfScene(scene, { width: drawn.width + 23, height: drawn.height + 14 })).toBeNull();
  });

  it("reads the factor per axis once a side is outside the margin", () => {
    expect(groupScaleOfScene(scene, { width: 960, height: 598 })).toEqual({ x: 2, y: 2, samples: 0, via: "group" });
    const wide = groupScaleOfScene(scene, { width: 720, height: 299 })!;
    expect(wide.x).toBeCloseTo(1.5, 9);
    expect(wide.y, "an axis inside the margin reads exactly 1").toBe(1);
    const shrunk = groupScaleOfScene(scene, { width: 240, height: 149.5 })!;
    expect(shrunk.x).toBeCloseTo(0.5, 9);
    expect(shrunk.y).toBeCloseTo(0.5, 9);
  });

  it("stands off at the edge of the margin and not before", () => {
    expect(groupScaleOfScene(scene, { width: 480 + 24, height: 299 })).toBeNull();
    expect(groupScaleOfScene(scene, { width: 480 + 25, height: 299 })).not.toBeNull();
    // The margin is the LARGER of 5pt and 5%: 5% on a 200pt side...
    const small = buildChart({ ...sampleConfig("clustered"), width: 200, height: 150 });
    const box = sceneDrawnExtent(small)!;
    expect(groupScaleOfScene(small, { width: box.width + 9, height: box.height })).toBeNull();
    expect(groupScaleOfScene(small, { width: box.width + 11, height: box.height })).not.toBeNull();
    // ...and 5pt on a 60pt one, where 5% is 3pt.
    const short: Scene = { width: 100, height: 100, nodes: [{ kind: "rect", x: 0, y: 0, w: 60, h: 60, fill: "#000" }] };
    expect(groupScaleOfScene(short, { width: 65, height: 60 })).toBeNull();
    expect(groupScaleOfScene(short, { width: 66, height: 60 })).not.toBeNull();
  });

  it("says nothing for a factor no person made, or a reading that is not a number", () => {
    expect(groupScaleOfScene(scene, { width: 480 * 5, height: 299 })).toBeNull();
    expect(groupScaleOfScene(scene, { width: 480 * 0.2, height: 299 })).toBeNull();
    expect(groupScaleOfScene(scene, {})).toBeNull();
    expect(groupScaleOfScene(scene, { width: 960 })).toBeNull();
    expect(groupScaleOfScene(scene, { width: Number.NaN, height: 299 })).toBeNull();
  });

  it("says nothing for a scene whose box it cannot know", () => {
    const pie = buildChart(sampleConfig("pie"));
    expect(sceneDrawnExtent(pie)).toBeNull();
    expect(groupScaleOfScene(pie, { width: 9999, height: 9999 })).toBeNull();
    expect(sceneDrawnExtent({ width: 10, height: 10, nodes: [] })).toBeNull();
  });

  it("reads a short side as unknown rather than as a factor", () => {
    const thin: Scene = {
      width: 100,
      height: 100,
      nodes: [{ kind: "rect", x: 0, y: 0, w: 200, h: 10, fill: "#000" }],
    };
    // 10 is under the 40pt floor: 5pt of 10 is half.
    expect(groupScaleOfScene(thin, { width: 400, height: 25 })).toMatchObject({ x: 2, y: 1 });
  });
});

describe("measureGroup and memberSizesRead", () => {
  const scene = buildChart(sampleConfig("clustered"));

  it("gives a ratio for a chart nobody touched, which is what a trace needs", () => {
    const m = measureGroup(scene, { width: 480, height: 299 });
    expect(m.drawn).toEqual({ width: 480, height: 299 });
    expect(m.ratio).toEqual({ x: 1, y: 1 });
  });

  it("gives no ratio when a side was not read", () => {
    expect(measureGroup(scene, { width: 480 }).ratio).toBeNull();
    expect(measureGroup(scene, { width: 480 }).drawn).not.toBeNull();
  });

  const sized = (n: SceneNode) => ({ width: 12, height: 9, name: n.name });

  it("counts the members the host gave both sides of, and a side alone is not a size", () => {
    const live = scene.nodes.map(sized);
    expect(memberSizesRead(live)).toBe(scene.nodes.length);
    expect(memberSizesRead(live.map((l) => ({ ...l, height: undefined })))).toBe(0);
    expect(memberSizesRead(live.map((l) => ({ ...l, width: Number.NaN })))).toBe(0);
    expect(memberSizesRead([])).toBe(0);
    // One answer among silence is an answer: the host says how big a member is.
    expect(memberSizesRead(live.map((l, i) => (i === 3 ? l : {})))).toBe(1);
  });

  /**
   * The host's answer is a property of the host and not of the node kind. The gate
   * was "no RECT size was read", so a text box or an ellipse that said how big it
   * was did not count, and a scene with no rect in it read zero on a host that
   * answered every member.
   */
  it("counts a text box and an ellipse as much as a rect", () => {
    const only = (kind: SceneNode["kind"]) => scene.nodes.map((n) => (n.kind === kind ? sized(n) : { name: n.name }));
    for (const kind of ["text", "rect", "line"] as const) {
      expect(
        scene.nodes.some((n) => n.kind === kind),
        `the clustered sample has no ${kind} node`,
      ).toBe(true);
      expect(memberSizesRead(only(kind)), kind).toBe(scene.nodes.filter((n) => n.kind === kind).length);
    }
    const scatter = buildChart(SHOWCASE[41]);
    expect(scatter.nodes.some((n) => n.kind === "ellipse")).toBe(true);
    expect(
      scatter.nodes.some((n) => n.kind === "rect"),
      "showcase #41 is meant to be a scene with no rect",
    ).toBe(false);
    const answers = scatter.nodes.map((n) => (n.kind === "ellipse" ? sized(n) : { name: n.name }));
    expect(memberSizesRead(answers)).toBe(scatter.nodes.filter((n) => n.kind === "ellipse").length);
  });
});
