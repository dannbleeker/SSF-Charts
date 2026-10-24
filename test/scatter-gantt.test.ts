import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutScatter, spreadCap, scatterDomain } from "../src/core/layout/scatter";
import { niceTicks } from "../src/core/format";
import { layoutGantt } from "../src/core/layout/gantt";
import { layoutColumns } from "../src/core/layout/column";
import { placeLabels } from "../src/core/labels";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig, MarkerSymbol, Series } from "../src/core/types";
import type { EllipseNode, RectNode, SceneNode, TextNode } from "../src/core/scene";
import { symbolPoints } from "../src/core/geometry";
import { sceneToSvg } from "../src/render/svg";

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

describe("label placer", () => {
  const bounds = { x: 0, y: 0, w: 200, h: 200 };
  it("places non-conflicting labels and skips impossible ones", () => {
    const reqs = [
      { cx: 100, cy: 100, r: 3, w: 30, h: 10 },
      { cx: 100, cy: 100, r: 3, w: 30, h: 10 }, // same anchor → different slot
      { cx: 100, cy: 100, r: 3, w: 500, h: 10 }, // wider than bounds → hidden
    ];
    const placed = placeLabels(reqs, bounds);
    expect(placed).toHaveLength(2);
    expect(placed[0].slot).not.toBe(placed[1].slot);
  });
  it("respects obstacles", () => {
    const placed = placeLabels(
      [{ cx: 10, cy: 10, r: 2, w: 20, h: 8 }],
      bounds,
      [{ x: 0, y: 0, w: 200, h: 200 }], // everything blocked
    );
    expect(placed).toHaveLength(0);
  });
});

describe("scatter & bubble", () => {
  it("maps X/Y rows to point positions", () => {
    const c = cfg({
      kind: "scatter",
      data: {
        categories: ["P1", "P2"],
        series: [
          { name: "X", values: [0, 100] },
          { name: "Y", values: [0, 100] },
        ],
      },
    });
    const { nodes, anchors } = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const pts = nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("point-"));
    expect(pts).toHaveLength(2);
    // P2 is right of and above P1.
    expect(pts[1].cx).toBeGreaterThan(pts[0].cx);
    expect(pts[1].cy).toBeLessThan(pts[0].cy);
    expect(anchors.plot.w).toBeGreaterThan(0);
  });

  it("scales bubble area by the Size row", () => {
    const c = cfg({
      kind: "bubble",
      data: {
        categories: ["Small", "Big"],
        series: [
          { name: "X", values: [10, 90] },
          { name: "Y", values: [10, 90] },
          { name: "Size", values: [25, 100] },
        ],
      },
    });
    const { nodes } = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR);
    // By name, not by position: markers paint back-to-front (largest first), so
    // the emission order is deliberately not the datasheet order.
    const pt = (i: number) => nodes.find((n): n is EllipseNode => n.kind === "ellipse" && n.name === `point-${i}`)!;
    // Area ∝ size → radius ratio = sqrt(100/25) = 2.
    expect(pt(1).rx / pt(0).rx).toBeCloseTo(2, 1);
  });

  it("paints bubbles back-to-front so a big one cannot bury a small one", () => {
    const c = cfg({
      kind: "bubble",
      data: {
        categories: ["Big", "Small"],
        series: [
          // Same point: the small bubble sits inside the big one's disc.
          { name: "X", values: [50, 50] },
          { name: "Y", values: [50, 50] },
          { name: "Size", values: [100, 4] },
        ],
      },
    });
    const { nodes } = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const order = nodes
      .filter((n) => n.kind === "ellipse" && n.name?.startsWith("point-"))
      .map((n) => (n as EllipseNode).name);
    // Datasheet order would emit the big one first and paint the small one
    // under it — invisible in all three renderers.
    expect(order).toEqual(["point-0", "point-1"]);
    const rx = (i: number) => (nodes.find((n) => n.name === `point-${i}`) as EllipseNode).rx;
    expect(rx(0)).toBeGreaterThan(rx(1)); // point-0 IS the big one, drawn first
  });

  it("labels points without overlaps", () => {
    const scene = buildChart(sampleConfig("bubble"));
    const labels = scene.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("label-"));
    expect(labels.length).toBeGreaterThan(3);
    for (let i = 0; i < labels.length; i++) {
      for (let j = i + 1; j < labels.length; j++) {
        const a = labels[i];
        const b = labels[j];
        const overlap = a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
        expect(overlap).toBe(false);
      }
    }
  });
});

describe("gantt", () => {
  it("draws activity bars spanning start to end", () => {
    const c = cfg({
      kind: "gantt",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Start", values: [0, 5] },
          { name: "End", values: [5, 10] },
          { name: "Milestone", values: [null, 10] },
        ],
      },
    });
    const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const bars = nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("bar-"));
    expect(bars).toHaveLength(2);
    // Equal durations → equal widths; B starts where A ends.
    expect(bars[0].w).toBeCloseTo(bars[1].w, 1);
    expect(bars[1].x).toBeCloseTo(bars[0].x + bars[0].w, 1);
    expect(nodes.find((n) => n.name === "milestone-1")).toBeTruthy();
  });

  describe('gutter columns ("Column <label>" rows)', () => {
    const withColumns = (extra: Record<string, unknown> = {}) =>
      cfg({
        kind: "gantt",
        width: 640,
        height: 260,
        data: {
          categories: ["Phase 1", "> Design", "> Build"],
          series: [
            { name: "Start", values: [null, 0, 4] },
            { name: "End", values: [null, 4, 10] },
            { name: "Column Cost", values: [250, 90.5, 159.5] },
            { name: "Column FTE", values: [3, 1, 2] },
          ],
        },
        ...extra,
      });

    it("renders each Column row as a headed, right-aligned gutter column", () => {
      const { nodes } = layoutGantt(withColumns(), DEFAULT_STYLE, DEFAULT_DECOR);
      const text = (name: string) => nodes.find((n) => n.name === name) as TextNode;
      expect(text("col-head-0").text).toBe("Cost");
      expect(text("col-head-1").text).toBe("FTE");
      expect(text("col-0-1").text).toBe("91"); // 90.5, at the Cost column's own precision
      expect(text("col-1-1").text).toBe("1.0"); // FTE resolves its own decimals
      for (const n of nodes.filter((x) => x.name?.startsWith("col-"))) {
        expect((n as TextNode).align).toBe("right");
      }
      // Columns sit side by side, left of the plot.
      expect(text("col-head-1").x).toBeGreaterThan(text("col-head-0").x);
    });

    it("does not turn a Column row into a bar, and shrinks the plot to make room", () => {
      const bars = (c: ChartConfig) =>
        layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR).nodes.filter(
          (n): n is RectNode => n.kind === "rect" && /^bar-\d+$/.test(n.name ?? ""),
        );
      const withCols = bars(withColumns());
      expect(withCols).toHaveLength(2); // Design + Build; Phase 1 has no span

      // Same chart without the Column rows: the bars must start further left.
      const bare = withColumns();
      bare.data.series = bare.data.series.filter((s) => !s.name.startsWith("Column"));
      const without = bars(bare);
      expect(withCols[0].x).toBeGreaterThan(without[0].x);
    });

    it("paints cells after the section band, which is full-width and would cover them", () => {
      // Phase 1 has no Start/End, so it is a section header: `section-0` is a
      // rect spanning the whole chart width, including the gutter.
      const { nodes } = layoutGantt(withColumns(), DEFAULT_STYLE, DEFAULT_DECOR);
      const band = nodes.findIndex((n) => n.name === "section-0");
      const cell = nodes.findIndex((n) => n.name === "col-0-0");
      expect(band).toBeGreaterThanOrEqual(0);
      expect(cell).toBeGreaterThan(band);
      // A header row still shows the value it carries — nothing is auto-summed.
      expect((nodes[cell] as TextNode).text).toBe("250");
    });
  });
});

describe("segment order & manual scale", () => {
  const data = {
    categories: ["A"],
    series: [
      { name: "S1", values: [10] },
      { name: "S2", values: [30] },
      { name: "S3", values: [20] },
    ],
  };
  it("descending puts the largest segment at the baseline", () => {
    const { nodes, anchors } = layoutColumns(cfg({ data, segmentOrder: "descending" }), DEFAULT_STYLE, DEFAULT_DECOR);
    const segs = nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("seg-"));
    const bottom = segs.reduce((a, b) => (a.y + a.h > b.y + b.h ? a : b));
    expect(bottom.name).toBe("seg-1-0"); // S2 = 30, the largest
    expect(bottom.y + bottom.h).toBeCloseTo(anchors.baselineY, 5);
  });
  it("pins the axis max", () => {
    const { anchors } = layoutColumns(cfg({ data, scale: { max: 120 } }), DEFAULT_STYLE, DEFAULT_DECOR);
    // Total 60 of max 120 → column fills half the plot.
    expect((anchors.baselineY - anchors.columnTop[0]) / anchors.plot.h).toBeCloseTo(0.5, 2);
  });
});

describe("gantt working-day timeline (gantt.workdays)", () => {
  const day = (iso: string) => Math.round(Date.parse(iso) / 86400000);
  /** Mon 5 Jan 2026 → Fri 9 (4 elapsed, 4 working) and → Mon 12 (7 elapsed, 5 working). */
  const twoBars = (gantt?: ChartConfig["gantt"], holiday?: string) =>
    cfg({
      kind: "gantt",
      width: 500,
      height: 200,
      ...(gantt ? { gantt } : {}),
      data: {
        categories: ["MonFri", "MonMon"],
        series: [
          { name: "Start", values: [day("2026-01-05"), day("2026-01-05")] },
          { name: "End", values: [day("2026-01-09"), day("2026-01-12")] },
          ...(holiday ? [{ name: "Holiday", values: [day(holiday), null] }] : []),
        ],
        dates: true,
      },
    });
  const widthRatio = (c: ChartConfig) => {
    const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const w = (i: number) => (nodes.find((n) => n.name === `bar-${i}`) as RectNode).w;
    return w(1) / w(0);
  };
  const count = (c: ChartConfig, prefix: string) =>
    layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR).nodes.filter((n) => n.name?.startsWith(prefix)).length;

  it("makes a bar's length the working days it contains", () => {
    // The ratio of the two bars is what discriminates: it is independent of the
    // timeline's padding, which is a whole number of weeks either way.
    expect(widthRatio(twoBars())).toBeCloseTo(7 / 4, 3); // elapsed days
    expect(widthRatio(twoBars({ workdays: true }))).toBeCloseTo(5 / 4, 3); // Mon→Mon is 5, not 7
  });

  it("takes Holiday rows out of the working days too", () => {
    // A Wednesday holiday costs both bars a day: 4→3 and 5→4.
    expect(widthRatio(twoBars({ workdays: true }, "2026-01-07"))).toBeCloseTo(4 / 3, 3);
  });

  it("drops weekend and holiday shading — those days have no width left", () => {
    const shaded = twoBars(undefined, "2026-01-07");
    expect(count(shaded, "weekend-")).toBeGreaterThan(0);
    expect(count(shaded, "holiday-")).toBeGreaterThan(0);
    const workweeks: (boolean | number[])[] = [true, [7, 1, 2, 3, 4]];
    for (const wd of workweeks) {
      // [7,1,2,3,4] (Sun–Thu) is the case the `x2 > x1` guard alone does NOT
      // catch: Saturday has width again, so the old block shaded a Sunday — a
      // working day there — and left the real non-working Friday unshaded.
      const c = twoBars({ workdays: wd }, "2026-01-07");
      expect(count(c, "weekend-"), String(wd)).toBe(0);
      expect(count(c, "holiday-"), String(wd)).toBe(0);
    }
  });

  it("reaches the right edge of the plot, exactly like the elapsed scale", () => {
    const c = twoBars({ workdays: true });
    const { nodes, anchors } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const right = anchors.plot!.x + anchors.plot!.w;
    const xs = nodes.filter((n) => n.name?.startsWith("bar-")).map((n) => (n as RectNode).x);
    for (const x of xs) expect(x).toBeGreaterThanOrEqual(anchors.plot!.x - 0.01);
    // The denominator is the working-day count of [t0, t1), so the far end maps
    // onto the right edge rather than short of it.
    const gridlines = nodes.filter((n) => n.name?.startsWith("gridline"));
    if (gridlines.length) {
      const maxX = Math.max(...gridlines.map((n) => (n as any).x1 ?? 0));
      expect(maxX).toBeLessThanOrEqual(right + 0.01);
    }
  });

  it("keeps a weekend-only task visible instead of collapsing it to nothing", () => {
    const c = cfg({
      kind: "gantt",
      width: 500,
      height: 200,
      gantt: { workdays: true },
      data: {
        categories: ["Weekend job", "Normal"],
        series: [
          { name: "Start", values: [day("2026-01-10"), day("2026-01-05")] }, // Sat
          { name: "End", values: [day("2026-01-11"), day("2026-01-09")] }, // Sun
        ],
        dates: true,
      },
    });
    const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const bar = nodes.find((n) => n.name === "bar-0") as RectNode;
    expect(bar.w).toBeGreaterThan(0); // zero working days, but still drawn
    expect(bar.w).toBeLessThan(3); // as a hairline, not a real span
  });

  it("is a no-op on a numeric timeline, where working days mean nothing", () => {
    const numeric = (gantt?: ChartConfig["gantt"]) =>
      layoutGantt(
        cfg({
          kind: "gantt",
          width: 500,
          height: 200,
          ...(gantt ? { gantt } : {}),
          data: {
            categories: ["A"],
            series: [
              { name: "Start", values: [0] },
              { name: "End", values: [5] },
            ],
          },
        }),
        DEFAULT_STYLE,
        DEFAULT_DECOR,
      ).nodes;
    expect(numeric({ workdays: true })).toEqual(numeric());
  });
});

describe("scatter marginal histograms (decorations.marginals)", () => {
  const pts = (n: number) =>
    cfg({
      kind: "scatter",
      width: 480,
      height: 300,
      data: {
        categories: Array.from({ length: n }, (_, i) => `p${i}`),
        series: [
          { name: "X", values: Array.from({ length: n }, (_, i) => (i * 97) % 100) },
          { name: "Y", values: Array.from({ length: n }, (_, i) => (i * 53) % 100) },
        ],
      },
    });
  const build = (c: ChartConfig, marginals?: "x" | "y" | "both") =>
    layoutScatter(c, DEFAULT_STYLE, { ...DEFAULT_DECOR, ...(marginals ? { marginals } : {}) });

  it("puts every axis tick on a bin edge", () => {
    // The point of a chart-adjacent histogram: a bar is read against the scale
    // beside it. A rule keyed off the sample size alone (Sturges,
    // Freedman-Diaconis) yields a count with no relation to the tick grid, so
    // its edges land between the ticks. Deriving the count FROM the ticks makes
    // alignment a theorem instead of a coincidence.
    const { nodes, anchors } = build(pts(40), "x");
    const bars = nodes.filter((n) => n.name?.startsWith("marginal-x-")) as RectNode[];
    expect(bars.length).toBeGreaterThan(0);
    const gridX = nodes.filter((n) => n.kind === "line" && n.name?.startsWith("gridline-x")).map((n) => (n as any).x1);
    const plot = anchors.plot!;
    const bw = plot.w / ((gridX.length - 1) * 2); // 40 points → 2 sub-bins per interval
    for (const gx of gridX) {
      // Every gridline must sit on a bin boundary (within float noise).
      const k = (gx - plot.x) / bw;
      expect(Math.abs(k - Math.round(k)), `tick at ${gx}`).toBeLessThan(1e-6);
    }
  });

  it("reserves real space: the plot shrinks by the gutter", () => {
    const base = build(pts(20)).anchors.plot!;
    const top = build(pts(20), "x").anchors.plot!;
    const right = build(pts(20), "y").anchors.plot!;
    const both = build(pts(20), "both").anchors.plot!;
    expect(top.h).toBeLessThan(base.h);
    expect(top.y).toBeGreaterThan(base.y);
    expect(right.w).toBeLessThan(base.w);
    expect(both.h).toBeLessThan(base.h);
    expect(both.w).toBeLessThan(base.w);
  });

  it("drops the marginals rather than the plot when there is no room", () => {
    const tiny = cfg({ ...pts(20), height: 90, width: 120 } as Partial<ChartConfig>);
    const { nodes, anchors } = build(tiny, "both");
    expect(nodes.filter((n) => n.name?.startsWith("marginal-"))).toHaveLength(0);
    // ...and the plot keeps the space it would have given away.
    expect(anchors.plot!.h).toBe(build(tiny).anchors.plot!.h);
  });

  it("changes nothing when it is off", () => {
    expect(build(pts(20)).nodes).toEqual(build(pts(20), undefined).nodes);
    expect(build(pts(20)).nodes.some((n) => n.name?.startsWith("marginal-"))).toBe(false);
  });

  it("keeps the plot height expression bit-exact at a non-default font size", () => {
    // Float subtraction is not associative: (H-t)-6-l is NOT H-(t+6+l). The
    // showcase deck is entirely fs=10, where the two agree — so the deck's
    // byte-identity gate CANNOT catch a "simplification" here. Pin the exact
    // value at a font size where they differ (128 would be the rewritten form).
    const c = cfg({
      kind: "scatter",
      width: 480,
      height: 180,
      title: "T",
      data: {
        categories: ["a", "b"],
        series: [
          { name: "X", values: [1, 2] },
          { name: "Y", values: [1, 2] },
          { name: "Group", values: [1, 2] }, // forces the legend into the height
        ],
      },
    });
    const { anchors } = layoutScatter(c, { ...DEFAULT_STYLE, fontSize: 8 }, DEFAULT_DECOR);
    expect(anchors.plot!.h).toBe(127.99999999999999);
  });

  it("does not let the colour legend fall into the top gutter", () => {
    // The gradient legend's min/max labels are the ones that spill: they sit
    // below the bar, so anchoring them to a lowered plot top would drop them
    // onto the marginal bars.
    const c = cfg({
      kind: "scatter",
      width: 480,
      height: 300,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 5, 9] },
          { name: "Y", values: [2, 6, 8] },
          { name: "Color", values: [10, 50, 90] },
        ],
      },
    });
    const { nodes } = layoutScatter(c, DEFAULT_STYLE, { ...DEFAULT_DECOR, marginals: "x" });
    const bars = nodes.filter((n) => n.name?.startsWith("marginal-x-")) as RectNode[];
    const top = Math.min(...bars.map((b) => b.y));
    for (const name of ["color-legend-min", "color-legend-max", "color-legend-title"]) {
      const n = nodes.find((x) => x.name === name) as TextNode;
      expect(n, name).toBeTruthy();
      expect(n.y + n.h, name).toBeLessThanOrEqual(top + 0.01);
    }
  });
});

describe("scatter overlap relief (scatter.spread)", () => {
  // Five bubbles piled on the SAME spot — a genuine overlap the relief must
  // break. (They used to be a point apart in y; that only overlapped because the
  // old forced-zero domain squashed them. With a data-driven domain five points
  // spanning a few units fill the plot ~50px apart and no longer touch, so a
  // real pile now means a real coincidence of coordinates.)
  const pile = (scatter?: ChartConfig["scatter"]) =>
    cfg({
      kind: "bubble",
      width: 480,
      height: 300,
      footnote: "src", // so footnoteH matches with and without spread
      data: {
        categories: ["a", "b", "c", "d", "e"],
        series: [
          { name: "X", values: [50, 50, 50, 50, 50] },
          { name: "Y", values: [50, 50, 50, 50, 50] },
          { name: "Size", values: [60, 60, 60, 60, 60] },
        ],
      },
      ...(scatter ? { scatter } : {}),
    });
  /** The nice-ticked span of the pile's y axis — the units spreadCap.limit is in. */
  const pileYRange = () => {
    const t = niceTicks(...scatterDomain(pile({ spread: "y" }), "y"), 5);
    return t[t.length - 1] - t[0];
  };
  const at = (c: ChartConfig, i: number) =>
    layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR).nodes.find((n) => n.name === `point-${i}`) as EllipseNode;

  it("never moves a marker further than the cap it discloses", () => {
    const c = pile({ spread: "y" });
    const cap = spreadCap(c)!;
    const plot = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR).anchors.plot!;
    const limitPx = (cap.limit / pileYRange()) * plot.h; // cap is in data units of the y domain
    for (let i = 0; i < 5; i++) {
      // The cap is the contract — the footnote quotes it, so it cannot be a
      // suggestion. Bounding a marker to the plot must never override it.
      expect(Math.abs(at(c, i).cy - at(pile(), i).cy), `point-${i}`).toBeLessThanOrEqual(limitPx + 1e-6);
    }
  });

  it("leaves the cross axis exact", () => {
    const c = pile({ spread: "y" });
    for (let i = 0; i < 5; i++) expect(at(c, i).cx).toBe(at(pile(), i).cx);
  });

  it("relieves the overlap, spreading a pile symmetrically about its middle", () => {
    const c = pile({ spread: "y", spreadLimit: 6 });
    const d = [0, 1, 2, 3, 4].map((i) => at(c, i).cy - at(pile(), i).cy);
    // The middle of a symmetric pile has no reason to move; the ends move most.
    expect(Math.abs(d[2])).toBeLessThan(0.01);
    expect(Math.abs(d[0])).toBeGreaterThan(Math.abs(d[1]));
    expect(d[0]).toBeCloseTo(-d[4], 6);
    expect(d[1]).toBeCloseTo(-d[3], 6);
    // Overlap is genuinely reduced: adjacent markers end further apart than the
    // 1pt of y they started with.
    const gapBefore = Math.abs(at(pile(), 0).cy - at(pile(), 1).cy);
    const gapAfter = Math.abs(at(c, 0).cy - at(c, 1).cy);
    expect(gapAfter).toBeGreaterThan(gapBefore);
  });

  it("does not move a marker that overlaps nothing, even at an axis extreme", () => {
    // Bounding the circle's EXTENT rather than its centre would shove a big
    // bubble sitting on the top tick inward by its whole radius — several times
    // the cap, on the first pass, for a marker with no overlap at all.
    const c = cfg({
      kind: "bubble",
      width: 480,
      height: 300,
      footnote: "src",
      scatter: { spread: "y" },
      data: {
        categories: ["top", "far"],
        series: [
          { name: "X", values: [20, 80] },
          { name: "Y", values: [80, 10] }, // 80 lands exactly on the top tick
          { name: "Size", values: [100, 10] },
        ],
      },
    });
    const bare = cfg({ ...c, scatter: undefined } as Partial<ChartConfig>);
    expect(at(c, 0).cy).toBe(at(bare, 0).cy);
    expect(at(c, 1).cy).toBe(at(bare, 1).cy);
  });

  it("is deterministic, and independent of the input order", () => {
    const a = layoutScatter(pile({ spread: "y" }), DEFAULT_STYLE, DEFAULT_DECOR).nodes;
    const b = layoutScatter(pile({ spread: "y" }), DEFAULT_STYLE, DEFAULT_DECOR).nodes;
    expect(a).toEqual(b);
  });

  it("yields to quadrants rather than blur which box a point is in", () => {
    const c = pile({ spread: "y" });
    const q = cfg({ ...c, decorations: { quadrants: { x: 50, y: 52 } } } as Partial<ChartConfig>);
    expect(spreadCap(q)).toBeNull();
    // ...and with no spread there is no approximation to disclose.
    const foot = buildChart(q).nodes.find((n) => n.name === "footnote") as TextNode;
    expect(foot.text).not.toMatch(/approximate/);
  });

  it("discloses the cap it enforces, in the axis's own units", () => {
    const foot = buildChart(pile({ spread: "y" })).nodes.find((n) => n.name === "footnote") as TextNode;
    expect(foot.text).toContain("Y positions approximate");
    expect(foot.text).toContain(String(spreadCap(pile({ spread: "y" }))!.limit));
  });

  it("changes nothing when it is off", () => {
    expect(layoutScatter(pile(), DEFAULT_STYLE, DEFAULT_DECOR).nodes).toEqual(
      layoutScatter(pile(undefined), DEFAULT_STYLE, DEFAULT_DECOR).nodes,
    );
  });
});

describe('combo marker series (Series.type: "marker")', () => {
  const combo = (markerType: "marker" | "line", extra: Partial<ChartConfig> = {}) =>
    cfg({
      kind: "combo",
      width: 480,
      height: 300,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Revenue", values: [40, 55, 50] },
          { name: "Consensus", type: markerType, values: [45, 45, 52] },
        ],
      },
      ...extra,
    });
  const names = (c: ChartConfig) => buildChart(c).nodes.map((n) => n.name ?? "");

  it("draws the points but not the segments between them", () => {
    // The values are per-category facts — a benchmark does not interpolate.
    const marker = names(combo("marker"));
    expect(marker.filter((n) => n.startsWith("combo-marker-0-"))).toHaveLength(3);
    expect(marker.filter((n) => n.startsWith("combo-line-0-"))).toHaveLength(0);
    // A line series is the same overlay WITH the segments.
    const line = names(combo("line"));
    expect(line.filter((n) => n.startsWith("combo-marker-0-"))).toHaveLength(3);
    expect(line.filter((n) => n.startsWith("combo-line-0-")).length).toBeGreaterThan(0);
  });

  it("keeps the columns as columns", () => {
    // The marker series must not be mistaken for another column, and the
    // column series must not be mistaken for the overlay.
    const scene = buildChart(combo("marker"));
    const cols = scene.nodes.filter((n) => n.kind === "rect" && /^seg-0-\d+$/.test(n.name ?? ""));
    expect(cols).toHaveLength(3);
  });

  it("measures the benchmark against the same scale as the columns it benchmarks", () => {
    // Q3: Revenue 50 vs Consensus 52 — the mark must sit just ABOVE the column
    // top. That only holds on the shared scale, which is why a marker series
    // has no business on a secondaryAxis.
    const scene = buildChart(combo("marker"));
    const col = scene.nodes.find((n) => n.name === "seg-0-2") as RectNode;
    const mark = scene.nodes.find((n) => n.name === "combo-marker-0-2") as RectNode;
    expect(mark.y).toBeLessThan(col.y); // 52 > 50 → higher on the chart
    expect(col.y - mark.y).toBeLessThan(col.h); // ...but only slightly
  });

  it("does not let a marker series rank a pareto", () => {
    // Written FIRST, which is the only order that trips it: the bar-series
    // lookup takes the first non-line series, so a marker would be picked and
    // the columns sorted by their own benchmark.
    const c = cfg({
      kind: "clustered",
      width: 480,
      height: 300,
      pareto: true,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Consensus", type: "marker", values: [99, 1, 50] },
          { name: "Revenue", values: [10, 90, 50] },
        ],
      },
    });
    const order = buildChart(c)
      .nodes.filter((n) => n.kind === "text" && n.name?.startsWith("category-"))
      .sort((a, b) => (a as TextNode).x - (b as TextNode).x)
      .map((n) => (n as TextNode).text);
    // Revenue descending: B(90), C(50), A(10). Ranking by Consensus would give A, C, B.
    expect(order).toEqual(["B", "C", "A"]);
  });
});

describe("gantt owner lanes (gantt.lanes)", () => {
  const day = (iso: string) => Math.round(Date.parse(iso) / 86400000);
  const plan = (lanes?: "owner") =>
    cfg({
      kind: "gantt",
      width: 560,
      height: 280,
      ...(lanes ? { gantt: { lanes } } : {}),
      data: {
        categories: ["Spec | Ana", "Build | Ben", "Review | Ana", "Ship | Ben", "Handover"],
        series: [
          {
            name: "Start",
            values: [day("2026-01-05"), day("2026-01-12"), day("2026-01-19"), day("2026-01-26"), day("2026-02-02")],
          },
          {
            name: "End",
            values: [day("2026-01-12"), day("2026-01-19"), day("2026-01-26"), day("2026-02-02"), day("2026-02-05")],
          },
          { name: "After", values: [null, 1, 2, 3, 4] },
        ],
        dates: true,
      },
    });
  const labels = (c: ChartConfig) =>
    buildChart(c)
      .nodes.filter((n) => n.kind === "text" && /^category-\d+$/.test(n.name ?? ""))
      .map((n) => (n as TextNode).text);

  it("groups tasks under a header per owner, keeping each lane's own order", () => {
    // A stable partition, not a sort: Spec still precedes Review inside Ana's
    // lane, because that order is the plan.
    // The axis renders the activity name; the ">" only drives the indent.
    // "Handover" has no owner, so it sits under a header of its own rather than
    // trailing Ben's tasks as if it were his (`test/fix-c3-config-shaping.test.ts`).
    expect(labels(plan("owner"))).toEqual(["Ana", "Spec", "Review", "Ben", "Build", "Ship", "Other", "Handover"]);
  });

  it("renumbers After so the arrows still join the same tasks", () => {
    // After values are 1-based ROW INDICES. Moving rows without renumbering
    // would leave every dependency pointing at whatever now sits at that index.
    const pairs = (c: ChartConfig) => {
      const nodes = buildChart(c).nodes;
      const name = new Map<number, string>();
      nodes
        .filter((n) => n.kind === "text" && /^category-\d+$/.test(n.name ?? ""))
        .forEach((t) =>
          name.set(
            +t.name!.split("-")[1],
            String((t as TextNode).text)
              .replace(/^>\s*/, "")
              .split("|")[0]
              .trim(),
          ),
        );
      const bars = nodes.filter((n) => /^bar-\d+$/.test(n.name ?? "")) as RectNode[];
      return nodes
        .filter((n) => /^dep-v-\d+$/.test(n.name ?? ""))
        .map((n) => {
          const succ = +n.name!.split("-")[2];
          const pred = bars.find((b) => (n as any).y1 >= b.y - 6 && (n as any).y1 <= b.y + b.h + 6);
          return `${name.get(succ)} after ${name.get(pred ? +pred.name!.split("-")[1] : -1)}`;
        })
        .sort();
    };
    expect(pairs(plan("owner"))).toEqual(pairs(plan()));
    expect(pairs(plan())).toContain("Review after Build");
  });

  it("leaves unowned tasks together at the end", () => {
    expect(labels(plan("owner")).at(-1)).toBe("Handover");
  });

  it("changes nothing when it is off — a plan's row order is the plan", () => {
    expect(buildChart(plan()).nodes).toEqual(buildChart(plan(undefined)).nodes);
    expect(labels(plan())).toEqual(["Spec", "Build", "Review", "Ship", "Handover"]);
  });
});

describe("scatter marker symbols", () => {
  const points = (markers?: ChartConfig["scatter"], kind: "scatter" | "bubble" = "scatter", extra: Series[] = []) =>
    buildChart(
      cfg({
        kind,
        width: 480,
        height: 300,
        data: {
          categories: ["a", "b", "c", "d", "e", "f"],
          series: [
            { name: "X", values: [1, 2, 3, 4, 5, 6] },
            { name: "Y", values: [2, 4, 3, 5, 4, 6] },
            { name: "Group", values: [1, 2, 3, 4, 5, 6] },
            ...extra,
          ],
        },
        scatter: markers,
      }),
    ).nodes;

  const named = (nodes: SceneNode[], prefix: string) => nodes.filter((n) => n.name?.startsWith(prefix));
  const shapeOf = (n: SceneNode) => (n.kind === "symbol" ? `symbol/${n.shape}` : n.kind);

  it("draws circles and nothing else by default — the chart it has always drawn", () => {
    // Guards the whole feature's blast radius: `markers` off must be the exact
    // scene, node for node, that shipped before symbols existed.
    const before = points(undefined);
    expect(before.some((n) => n.kind === "symbol")).toBe(false);
    expect(named(before, "point-").map((n) => n.kind)).toEqual(Array(6).fill("ellipse"));
    // An empty list is "off", not "cycle through nothing".
    expect(points({ markers: [] })).toEqual(before);
  });

  it("assigns a shape per Group and cycles when groups outnumber shapes", () => {
    const got = named(points({ markers: ["circle", "square", "diamond"] }), "point-").map(shapeOf);
    expect(got).toEqual(["ellipse", "rect", "symbol/diamond", "ellipse", "rect", "symbol/diamond"]);
  });

  it("keeps each point's datasheet index, so shape never renames a point", () => {
    const got = named(points({ markers: ["triangle", "plus"] }), "point-").map((n) => n.name);
    expect(got.sort()).toEqual(["point-0", "point-1", "point-2", "point-3", "point-4", "point-5"]);
  });

  it("draws the legend chip as the shape it explains", () => {
    const nodes = points({ markers: ["circle", "square", "diamond", "triangle", "plus"] });
    const chips = named(nodes, "legend-chip-").map(shapeOf);
    const pts = named(nodes, "point-").map(shapeOf);
    // Group g's chip and group g's points must be the same shape, or the
    // legend explains a chart that isn't on the slide.
    expect(chips).toEqual(pts);
  });

  it("keeps a legend chip clear of its own label, at any font size", () => {
    // An area-matched diamond reaches ~1.25x the chip slot. The chip scales
    // with fontSize but the 3pt gap after it does NOT, so the overhang is
    // swallowed at small type and only bites past fs≈34 — which is exactly why
    // this asserts across sizes instead of the default alone, and why the top
    // of the range has to clear that threshold or the test proves nothing.
    // (Every showcase config uses fs=10, so the deck's gate cannot see this.)
    for (const fontSize of [8, 10, 14, 20, 28, 40]) {
      const nodes = buildChart(
        cfg({
          kind: "scatter",
          width: 720,
          height: 400,
          style: { fontSize },
          data: {
            categories: ["a", "b", "c"],
            series: [
              { name: "X", values: [1, 2, 3] },
              { name: "Y", values: [2, 4, 3] },
              { name: "Group", values: [1, 2, 3] },
            ],
          },
          scatter: { markers: ["diamond", "diamond", "diamond"] },
        }),
      ).nodes;
      for (const g of [1, 2, 3]) {
        const chip = nodes.find((n) => n.name === `legend-chip-${g}`)!;
        const label = nodes.find((n) => n.name === `legend-${g}`)! as TextNode;
        expect(chip.kind).toBe("symbol");
        const right = chip.kind === "symbol" ? chip.cx + chip.size : NaN;
        expect(right, `fs=${fontSize} group ${g}`).toBeLessThanOrEqual(label.x);
      }
    }
  });

  it("brings the group legend back under a Color row, where color means something else", () => {
    const color: Series[] = [{ name: "Color", values: [1, 2, 3, 4, 5, 6] }];
    // Without markers a color scale supersedes group coloring: no chips.
    expect(named(points(undefined, "scatter", color), "legend-chip-")).toHaveLength(0);
    // With markers, group is on the shape channel and must be explained.
    const nodes = points({ markers: ["circle", "diamond"] }, "scatter", color);
    expect(named(nodes, "legend-chip-").length).toBeGreaterThan(1);
    // ...and the chips must clear the gradient bar rather than sit on it.
    const barRight = Math.max(...named(nodes, "color-legend-").map((n) => (n.kind === "rect" ? n.x + n.w : 0)));
    for (const c of named(nodes, "legend-chip-")) {
      const left = c.kind === "symbol" ? c.cx - c.size : c.kind === "ellipse" ? c.cx - c.rx : (c as RectNode).x;
      expect(left).toBeGreaterThan(barRight);
    }
  });

  it("gives equal Size values equal ink, whatever their shape", () => {
    // A bubble asserts area ∝ size. Shape is categorical and must not perturb
    // it — otherwise a group drawn as stars reads as a quarter the magnitude.
    const size: Series[] = [{ name: "Size", values: [10, 10, 10, 10, 10, 10] }];
    const nodes = points({ markers: ["circle", "square", "diamond", "triangle", "plus"] }, "bubble", size);
    const areas = named(nodes, "point-").map((n) => {
      if (n.kind === "ellipse") return Math.PI * n.rx * n.ry;
      if (n.kind === "rect") return n.w * n.h;
      if (n.kind !== "symbol") throw new Error(`unexpected ${n.kind}`);
      const p = symbolPoints(n.shape, n.cx, n.cy, n.size);
      return (
        Math.abs(p.reduce((a, c, i) => a + (c.x * p[(i + 1) % p.length].y - p[(i + 1) % p.length].x * c.y), 0)) / 2
      );
    });
    for (const a of areas) expect(a).toBeCloseTo(areas[0], 6);
  });
});

describe("scatter group ids are whatever a datasheet cell holds", () => {
  const withGroups = (groups: (number | null)[], markers?: MarkerSymbol[]): ChartConfig =>
    cfg({
      kind: "scatter",
      width: 400,
      height: 260,
      data: {
        categories: groups.map((_, i) => `p${i}`),
        series: [
          { name: "X", values: groups.map((_, i) => i + 1) },
          { name: "Y", values: groups.map((_, i) => i + 2) },
          { name: "Group", values: groups },
        ],
      },
      ...(markers ? { scatter: { markers } } : {}),
    });

  // A Group row is data, so it can hold anything — including values no index
  // should ever see.
  //
  // Measured against the shipped code, NaN and +Infinity were the two that
  // actually broke: `Math.max(0, x - 1)` passes both straight through, so the
  // index became NaN, and `markers[NaN]` is no shape at all — which threw in
  // the renderer, on a chart that drew fine before markers existed.
  // zero/negative/-Infinity were already clamped to the first group and are
  // kept here as regression guards on the normalisation, not as past bugs.
  const junk: [string, number][] = [
    ["NaN", NaN],
    ["Infinity", Infinity],
    ["-Infinity", -Infinity],
    ["zero", 0],
    ["negative", -3],
  ];

  for (const [name, g] of junk) {
    it(`renders, and picks a real shape and colour, for a ${name} group`, () => {
      for (const markers of [undefined, ["circle", "diamond", "triangle"] as MarkerSymbol[]]) {
        const scene = buildChart(withGroups([g, 2], markers));
        const svg = sceneToSvg(scene);
        expect(svg, `${name}/${markers ? "markers" : "default"}`).not.toContain("undefined");
        expect(svg, `${name}/${markers ? "markers" : "default"}`).not.toContain("NaN");
        const p = scene.nodes.find((n) => n.name === "point-0")!;
        expect(p).toBeTruthy();
        if (p.kind === "symbol") expect(p.size).toBeGreaterThan(0);
        expect((p as { fill: string }).fill).toMatch(/^#[0-9a-f]{6}$/i);
      }
    });
  }

  it("buckets an unusable id into group 1, where a missing Group row already lands", () => {
    // Same chart, said two ways: junk ids and no Group row at all.
    const junky = buildChart(withGroups([NaN, NaN], ["diamond", "triangle"])).nodes;
    const none = buildChart(
      cfg({
        kind: "scatter",
        width: 400,
        height: 260,
        data: {
          categories: ["p0", "p1"],
          series: [
            { name: "X", values: [1, 2] },
            { name: "Y", values: [2, 3] },
          ],
        },
        scatter: { markers: ["diamond", "triangle"] },
      }),
    ).nodes;
    expect(junky).toEqual(none);
  });

  it("rounds a fractional id rather than indexing between shapes", () => {
    const shapes = buildChart(withGroups([1.4, 2.6], ["circle", "diamond", "triangle"]))
      .nodes.filter((n) => n.name?.startsWith("point-"))
      .map((n) => (n.kind === "symbol" ? n.shape : n.kind));
    expect(shapes).toEqual(["ellipse", "triangle"]); // 1.4 -> 1, 2.6 -> 3
  });
});

describe("polynomial scatter trendlines", () => {
  const curved = (deg?: number): ChartConfig => ({
    kind: "scatter",
    width: 480,
    height: 320,
    scatter: deg ? { trendDegree: deg } : undefined,
    decorations: { valueAxis: true },
    data: {
      categories: ["1", "2", "3", "4", "5", "6", "7"],
      series: [
        { name: "X", values: [1, 2, 3, 4, 5, 6, 7] },
        { name: "Y", values: [10, 4, 1, 0, 1, 4, 10] }, // parabola
        { name: "Trend", values: [1, null, null, null, null, null, null] },
      ],
    },
  });

  it("draws a single straight trend line by default (degree 1, unchanged)", () => {
    const nodes = layoutScatter(curved(), DEFAULT_STYLE, DEFAULT_DECOR).nodes;
    expect(nodes.filter((n) => n.name === "trend")).toHaveLength(1);
    expect(nodes.some((n) => n.name?.startsWith("trend-seg-"))).toBe(false);
    const stats = nodes.find((n) => n.name === "trend-stats") as { text: string };
    expect(stats.text).toMatch(/^R² = /);
  });

  it("draws a sampled polynomial curve and names the degree at higher degrees", () => {
    const nodes = layoutScatter(curved(2), DEFAULT_STYLE, DEFAULT_DECOR).nodes;
    expect(nodes.some((n) => n.name === "trend")).toBe(false); // not the single straight line
    const segs = nodes.filter((n) => n.name?.startsWith("trend-seg-"));
    expect(segs.length).toBeGreaterThan(10); // a smooth curve of segments
    for (const s of segs) expect(s.kind).toBe("line");
    const stats = nodes.find((n) => n.name === "trend-stats") as { text: string };
    expect(stats.text).toContain("quadratic");
    expect(stats.text).toMatch(/R² = (0\.9\d|1\.00)/); // a parabola fits the parabola well
  });

  it("clamps the drawn degree to points − 2 (never interpolates)", () => {
    // 4 points, quartic requested → clamps to degree 2 (n−2), still a curve.
    // (A quartic through 4 points would interpolate exactly — meaningless.)
    const cfg: ChartConfig = {
      kind: "scatter",
      width: 480,
      height: 320,
      scatter: { trendDegree: 4 },
      data: {
        categories: ["a", "b", "c", "d"],
        series: [
          { name: "X", values: [0, 1, 2, 3] },
          { name: "Y", values: [0, 1, 4, 9] },
          { name: "Trend", values: [1, null, null, null] },
        ],
      },
    };
    const nodes = layoutScatter(cfg, DEFAULT_STYLE, DEFAULT_DECOR).nodes;
    const stats = nodes.find((n) => n.name === "trend-stats") as { text: string };
    expect(stats.text).toContain("quadratic"); // clamped from quartic
  });
});
