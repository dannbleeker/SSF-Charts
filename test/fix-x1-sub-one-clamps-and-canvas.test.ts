import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { alphaOf, layoutBackground, toRgb } from "../src/core/color";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { resolveStyle } from "../src/core/style";
import { sceneToSvg } from "../src/render/svg";
import type { EllipseNode, LineNode, PolygonNode, RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * What 47aa582 and 7ab702a left undone, finished.
 *
 * A (r1-9)  The combo, the small-multiples radar and the glyph tile map still
 *           floored their value range at 1 IN DATA UNITS. Every test below draws the
 *           SAME data at two units and compares the geometry, which is the invariant
 *           (0.08 and 8 are one chart), not a pinned number that is only today's output.
 * B (r1-4)  `rgba(0,0,0,0)`, `#0000`, `none`, a typo and `""` read as grey, then dark,
 *           at every direct `lerpColor(style.background, …)` site.
 * C (r2-35) The hollow forecast marker was a patch in the canvas colour.
 */

/** The same data at another unit. Nulls stay blank. */
const scaled = (cfg: ChartConfig, k: number): ChartConfig => ({
  ...cfg,
  data: {
    ...cfg.data,
    series: cfg.data.series.map((s) => ({ ...s, values: s.values.map((v) => (v == null ? v : v * k)) })),
  },
});

const named = <T extends SceneNode>(nodes: SceneNode[], kind: T["kind"], prefix: string): T[] =>
  nodes.filter((n): n is T => n.kind === kind && !!n.name?.startsWith(prefix));

/** Two lists of numbers agree to a micro-point. */
const sameNumbers = (a: number[], b: number[]) => {
  expect(a.length).toBeGreaterThan(0);
  expect(a.length).toBe(b.length);
  a.forEach((v, i) => expect(v).toBeCloseTo(b[i], 6));
};

const finite = (cfg: ChartConfig) => {
  for (const n of buildChart(cfg).nodes)
    for (const [k, v] of Object.entries(n))
      if (typeof v === "number") expect(Number.isFinite(v), `${n.name}.${k}`).toBe(true);
};

// ---------------------------------------------------------------------------
// A (r1-9) — a floor of 1 in data units does not decide the scale
// ---------------------------------------------------------------------------

describe("a combo draws data below 1 as it draws the same data at x100 (r1-9)", () => {
  const combo: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C", "D"],
      series: [
        { name: "X", values: [0.2, 0.3, 0.25, 0.1] },
        { name: "Y", values: [0.1, 0.2, 0.15, 0.2] },
        { name: "Line", values: [0.3, 0.4, 0.45, 0.6], type: "line" },
      ],
    },
  };
  const colHeights = (cfg: ChartConfig) => named<RectNode>(buildChart(cfg).nodes, "rect", "seg-").map((r) => r.h);
  const barBoxes = (cfg: ChartConfig) =>
    named<RectNode>(buildChart(cfg).nodes, "rect", "bar-").flatMap((r) => [r.y, r.h]);
  const markerYs = (cfg: ChartConfig) =>
    named<RectNode>(buildChart(cfg).nodes, "rect", "combo-marker-").map((r) => r.y + r.h / 2);

  it("stacked columns and the shared-axis line fill the plot", () => {
    // Column heights 54.8 / 27.4 / ... against 91.3 / 45.7 / ... before: `niceTicks(0,
    // Math.max(stackMax, lineMax, 1))` pinned the shared axis at 0..1 for a combo whose
    // tallest value is 0.6, so the line drew at 60% of the plot.
    sameNumbers(colHeights(combo), colHeights(scaled(combo, 100)));
    sameNumbers(markerYs(combo), markerYs(scaled(combo, 100)));
    expect(Math.min(...markerYs(combo)), "the tallest point reaches the top of the plot").toBeLessThan(20);
  });

  it("the shared axis follows a line that dips below the columns' base", () => {
    // `sharedFloor` carried the same floor of 1.
    const dips: ChartConfig = {
      ...combo,
      data: {
        ...combo.data,
        series: [
          { name: "X", values: [0.2, 0.3, 0.25, 0.1] },
          { name: "Line", values: [-0.3, 0.4, -0.45, 0.6], type: "line" },
        ],
      },
    };
    sameNumbers(colHeights(dips), colHeights(scaled(dips, 100)));
    sameNumbers(markerYs(dips), markerYs(scaled(dips, 100)));
  });

  it("a secondary axis is ticked in the line's own units", () => {
    // The strip read 0.0 / 0.5 / 1.0 for a line of 0.3-0.6 (`Math.max(1, lineMax)`).
    const secondary: ChartConfig = { ...combo, secondaryAxis: true };
    sameNumbers(markerYs(secondary), markerYs(scaled(secondary, 100)));
    const ticks = named<TextNode>(buildChart(secondary).nodes, "text", "secondary").map((t) => Number(t.text));
    expect(ticks.length).toBeGreaterThan(1);
    expect(Math.max(...ticks)).toBeGreaterThanOrEqual(0.6);
    expect(Math.max(...ticks)).toBeLessThan(0.9);
  });

  it("a waterfall under the line is stretched to the line's own top (overflow branch)", () => {
    // wf.max is 0.11 here and the line reaches 0.6, so the overflow stretch fires.
    const overflow: ChartConfig = {
      ...combo,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "X", values: [0.05, 0.02, 0.03, 0.01] },
          { name: "Line", values: [0.3, 0.4, 0.45, 0.6], type: "line" },
        ],
      },
    };
    sameNumbers(barBoxes(overflow), barBoxes(scaled(overflow, 100)));
    sameNumbers(markerYs(overflow), markerYs(scaled(overflow, 100)));
  });

  it("a waterfall over a line that dips below its trough (underflow branch)", () => {
    const underflow: ChartConfig = {
      ...combo,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "X", values: [0.05, -0.02, 0.03, 0.01] },
          { name: "Line", values: [0.03, -0.4, 0.04, 0.05], type: "line" },
        ],
      },
    };
    sameNumbers(barBoxes(underflow), barBoxes(scaled(underflow, 100)));
    sameNumbers(markerYs(underflow), markerYs(scaled(underflow, 100)));
  });

  it("a waterfall with a line both above its peak and below its trough (both branches)", () => {
    const both: ChartConfig = {
      ...combo,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "X", values: [0.05, 0.02, 0.03, 0.01] },
          { name: "Line", values: [0.3, 0.4, -0.05, 0.9], type: "line" },
        ],
      },
    };
    sameNumbers(barBoxes(both), barBoxes(scaled(both, 100)));
    sameNumbers(markerYs(both), markerYs(scaled(both, 100)));
  });

  it("an all-zero combo still draws a finite chart on a 0..1 axis", () => {
    // The guard the floor was thought to be: `niceTicks(0, 0)` answers [0, 1].
    const zero = scaled(combo, 0);
    finite(zero);
    finite({ ...zero, secondaryAxis: true });
    finite({ ...zero, combo: { columns: "waterfall" } });
    expect(markerYs(zero).length).toBeGreaterThan(0);
  });

  it("data that already reaches 1 is where it was", () => {
    // The change is a removed floor, so a combo at 1 and above cannot move: the
    // x100 data and the x10 data are the same chart.
    sameNumbers(colHeights(scaled(combo, 10)), colHeights(scaled(combo, 100)));
  });
});

describe("a small-multiples radar draws data below 1 as it draws the same data at x100 (r1-9)", () => {
  const radar: ChartConfig = {
    kind: "radar",
    ...DEFAULT_SIZE,
    multiples: {},
    data: {
      categories: ["Speed", "Power", "Range", "Cost", "Style"],
      series: [
        { name: "A", values: [0.2, 0.4, 0.3, 0.25, 0.35] },
        { name: "B", values: [0.1, 0.3, 0.2, 0.35, 0.15] },
      ],
    },
  };
  const points = (cfg: ChartConfig, name: string) =>
    named<PolygonNode>(buildChart(cfg).nodes, "polygon", name).flatMap((p) => p.points.flatMap((q) => [q.x, q.y]));

  it("both panels share a web as large as the x100 one", () => {
    // Rings 0.5 / 1.0 before: the panels were pinned to 0..max(globalMax, 1).
    sameNumbers(points(radar, "p0-series-0"), points(scaled(radar, 100), "p0-series-0"));
    sameNumbers(points(radar, "p1-series-0"), points(scaled(radar, 100), "p1-series-0"));
  });

  it("the ring labels come from the data", () => {
    const ticks = named<TextNode>(buildChart(radar).nodes, "text", "p0-tick-").map((t) => Number(t.text));
    expect(ticks.length).toBeGreaterThan(0);
    expect(Math.max(...ticks)).toBeLessThan(0.5);
  });

  it("an all-zero grid still draws a finite web", () => {
    finite(scaled(radar, 0));
  });

  it("a negative value is not clamped onto the centre of its panel", () => {
    // The scale was `{ min: 0, … }` whatever the data held, the pin the single radar
    // lost in r1-11: -20, -5 and 0 drew at one point.
    const neg: ChartConfig = {
      ...radar,
      data: {
        categories: radar.data.categories,
        series: [
          { name: "A", values: [-20, -5, 0, 10, 15] },
          { name: "B", values: [5, 10, 15, 20, 25] },
        ],
      },
    };
    const nodes = buildChart(neg).nodes;
    const markers = named<EllipseNode>(nodes, "ellipse", "p0-marker-0-");
    expect(markers.length).toBe(5);
    // Three different spokes, so compare each marker's distance from the web's centre,
    // which every spoke shares: the first spoke of the panel starts there.
    const spokeLine = named<LineNode>(nodes, "line", "p0-spoke-0")[0];
    expect(spokeLine, "the panel's first spoke").toBeDefined();
    const dist = [0, 1, 2].map((i) => {
      const m = markers.find((e) => e.name === `p0-marker-0-${i}`)!;
      return Math.hypot(m.cx - spokeLine.x1, m.cy - spokeLine.y1);
    });
    expect(dist[0], "-20 sits at the centre of a web that starts at -20").toBeCloseTo(0, 6);
    expect(dist[1], "-5 is further out than -20").toBeGreaterThan(dist[0] + 1);
    expect(dist[2], "0 is further out than -5").toBeGreaterThan(dist[1] + 1);
  });
});

describe("a glyph tile map draws data below 1 as it draws the same data at x100 (r1-9)", () => {
  const countries = ["DK", "SE", "NO", "FI", "DE", "NL", "FR", "GB", "ES", "IT", "PL"];
  const tiles: ChartConfig = {
    kind: "tilemap",
    ...DEFAULT_SIZE,
    map: "europe",
    tilemap: { glyph: "bars" },
    data: {
      categories: countries,
      series: [
        { name: "A", values: countries.map((_, i) => 0.05 + ((i * 7) % 40) / 100) },
        { name: "B", values: countries.map((_, i) => 0.04 + ((i * 3 + 5) % 30) / 100) },
      ],
    },
  };
  const bars = (cfg: ChartConfig) => named<RectNode>(buildChart(cfg).nodes, "rect", "glyph-").map((r) => r.h);

  it("the mini-bars are as tall for 0.05-0.44 as for 5-44", () => {
    // 5.15pt against 12.86pt before: `maxOf(glyphVals, 1)` scaled the tallest bar to
    // 0.44 of a 1-unit range.
    sameNumbers(bars(tiles), bars(scaled(tiles, 100)));
  });

  it("an all-zero map draws finite, zero-height bars", () => {
    const zero = scaled(tiles, 0);
    finite(zero);
    const hs = bars(zero);
    expect(hs.length).toBeGreaterThan(0);
    for (const h of hs) expect(h).toBe(0);
  });

  it("a map of negatives hangs its bars from the zero line at the same scale as x100", () => {
    const down = scaled(tiles, -1);
    sameNumbers(bars(down), bars(scaled(down, 100)));
  });
});

// ---------------------------------------------------------------------------
// B (r1-4) — every spelling of "no canvas" lays out like `transparent`
// ---------------------------------------------------------------------------

/** Spellings that draw nothing, then spellings no sink can read. */
const CLEAR = [
  "rgba(0,0,0,0)",
  "#0000",
  "#00000000",
  "hsla(0,0%,0%,0)",
  "rgb(0 0 0 / 0)",
  "TRANSPARENT",
  " transparent ",
];
const UNREADABLE = ["none", "off-white", "not-a-color", "", "#abc;", "hsl(0.5turn 50% 50%)"];
const white = "#ffffff";

describe("layoutBackground reads one answer for every spelling of no canvas (r1-4)", () => {
  it.each(CLEAR)("%j draws nothing: the transparent keyword, which reads as the white slide", (bg) => {
    expect(layoutBackground(bg)).toBe("transparent");
    expect(alphaOf(layoutBackground(bg))).toBe(0);
    expect(toRgb(layoutBackground(bg))).toEqual([255, 255, 255]);
  });

  it("none is the SVG word for no paint, so it is the keyword too", () => {
    expect(layoutBackground("none")).toBe("transparent");
    expect(layoutBackground("NONE")).toBe("transparent");
  });

  it.each(UNREADABLE.slice(1))("%j cannot be read by any sink: the canvas they draw for it", (bg) => {
    expect(layoutBackground(bg)).toBe(white);
  });

  it("a background that is not a string reads as unreadable, not as a crash", () => {
    for (const bad of [undefined, null, 5, {}, []]) expect(layoutBackground(bad)).toBe(white);
  });

  it("a colour that draws something is returned byte for byte, light or dark", () => {
    for (const c of ["#ffffff", "#1b1b1b", "#fff", "white", "black", "rgb(10, 20, 30)", "rgba(0,0,0,0.5)", "#00000080"])
      expect(layoutBackground(c), c).toBe(c);
    // The genuinely dark canvas must stay dark: the other half, or "treat it as
    // light" could be satisfied by never asking.
    expect(toRgb(layoutBackground("#1b1b1b"))).toEqual([27, 27, 27]);
  });

  it("resolveStyle applies it to the style a layout is handed, and nothing else", () => {
    expect(resolveStyle({ background: "rgba(0,0,0,0)" }).background).toBe("transparent");
    expect(resolveStyle({ background: "off-white" }).background).toBe(white);
    expect(resolveStyle({ background: "#101010", fontSize: 14 })).toMatchObject({
      background: "#101010",
      fontSize: 14,
    });
    expect(resolveStyle().background).toBe(white);
    // An explicit undefined overrides the default in a spread; it must not reach a layout.
    expect(resolveStyle({ background: undefined }).background).toBe(white);
  });
});

describe("every chart kind lays out the same on any no-canvas spelling as on transparent (r1-4)", () => {
  const withBackground = (base: ChartConfig, background: string) =>
    buildChart({ ...base, style: { ...(base.style ?? {}), background } } as ChartConfig);

  /**
   * Charts whose direct `lerpColor(style.background, …)` sites the sample of a kind
   * does not reach: the column scenario marks, the line ribbons and the sparkline wash,
   * the grouped treemap, the constant-data heatmap and tile map, the glyph tile
   * backdrop and a negative bubble.
   */
  const variants: Record<string, ChartConfig> = {
    "column scenario marks (PY, FC)": {
      kind: "clustered",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "AC", values: [10, 12, 14] },
          { name: "PY", values: [9, 11, 12], scenario: "PY" },
          { name: "FC", values: [11, 13, 15], scenario: "FC" },
        ],
      },
    },
    "line fill-between and band ribbon": {
      kind: "line",
      ...DEFAULT_SIZE,
      decorations: { fillBetween: [0, 1], forecastFrom: 2 },
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "Plan", values: [10, 12, 14, 16] },
          { name: "Actual", values: [9, 13, 13, 17] },
        ],
      },
    },
    "line confidence band": {
      kind: "line",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Forecast", values: [50, 56, 63] },
          { name: "Band low", values: [50, 52, 55] },
          { name: "Band high", values: [50, 60, 72] },
        ],
      },
    },
    "sparkline area wash": {
      kind: "area",
      ...DEFAULT_SIZE,
      decorations: { sparkline: true },
      data: { categories: ["A", "B", "C", "D"], series: [{ name: "S", values: [4, 7, 5, 9] }] },
    },
    "grouped treemap": {
      kind: "treemap",
      ...DEFAULT_SIZE,
      data: {
        categories: ["North | A", "North | B", "South | A", "South | B", "South | C"],
        series: [{ name: "Revenue", values: [40, 30, 25, 15, 10] }],
      },
    },
    "constant-data heatmap": {
      kind: "heatmap",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "North", values: [5, 5] },
          { name: "South", values: [5, 5] },
        ],
      },
    },
    "constant-data tile map": {
      kind: "tilemap",
      ...DEFAULT_SIZE,
      map: "europe",
      data: { categories: ["DK", "SE", "NO", "FI"], series: [{ name: "V", values: [7, 7, 7, 7] }] },
    },
    "glyph tile map backdrop": {
      kind: "tilemap",
      ...DEFAULT_SIZE,
      map: "europe",
      tilemap: { glyph: "bars" },
      data: {
        categories: ["DK", "SE", "NO", "FI"],
        series: [
          { name: "A", values: [4, 6, 3, 8] },
          { name: "B", values: [5, 2, 7, 1] },
        ],
      },
    },
    "bubble with a negative size": {
      ...sampleConfig("bubble"),
      data: {
        categories: ["Alpha", "Bravo", "Core", "Delta"],
        series: [
          { name: "X", values: [12, 25, 40, 55] },
          { name: "Y", values: [30, 55, 42, 70] },
          { name: "Size", values: [10, -40, 90, -25] },
        ],
      },
    },
  };

  /**
   * Every colour property of the scene, in order. The ONLY paint allowed to differ is
   * one that IS the background literal: `transparent` in the reference build, and in
   * the subject `transparent`, the spelling as authored, or `#ffffff` (what the canvas
   * is drawn as for a paint no sink can read). A tint, a ramp stop, an ink or a
   * no-data grey that differs is the bug: the maths answers a `#rrggbb`, never the
   * spelling, so nothing a mix produced can hide behind this allowance.
   */
  const differences = (base: ChartConfig, spelling: string, label: string) => {
    const ref = withBackground(base, "transparent");
    const got = withBackground(base, spelling);
    const diffs: string[] = [];
    let compared = 0;
    if (ref.nodes.length !== got.nodes.length) return { diffs: [`${label}: node count`], compared };
    ref.nodes.forEach((a, i) => {
      const b = got.nodes[i] as unknown as Record<string, unknown>;
      for (const p of ["fill", "stroke", "color"]) {
        const av = (a as unknown as Record<string, unknown>)[p];
        compared++;
        if (av === b[p]) continue;
        if (av === "transparent" && [spelling, "transparent", white].includes(b[p] as string)) continue;
        diffs.push(`${label}/${a.name ?? i}.${p}: ${String(av)} vs ${String(b[p])}`);
      }
    });
    return { diffs, compared };
  };

  it.each([...CLEAR, ...UNREADABLE])("%j: the sample of every kind paints as transparent does", (spelling) => {
    const all: string[] = [];
    let compared = 0;
    for (const { kind } of CHART_KINDS) {
      const r = differences(sampleConfig(kind), spelling, kind);
      all.push(...r.diffs);
      compared += r.compared;
    }
    expect(compared, "the sweep compared nothing").toBeGreaterThan(3000);
    expect(all.slice(0, 12), `${all.length} paints differ`).toEqual([]);
  });

  it.each([...CLEAR, ...UNREADABLE])("%j: the sites the samples do not reach paint as transparent does", (spelling) => {
    const all: string[] = [];
    let compared = 0;
    for (const [label, cfg] of Object.entries(variants)) {
      const r = differences(cfg, spelling, label);
      all.push(...r.diffs);
      compared += r.compared;
    }
    expect(compared, "the sweep compared nothing").toBeGreaterThan(500);
    expect(all.slice(0, 12), `${all.length} paints differ`).toEqual([]);
  });

  it("the variants are real: on a dark canvas they differ from the light one", () => {
    // Otherwise the sweep above could be passing over charts that never read the
    // background at all.
    for (const [label, cfg] of Object.entries(variants)) {
      const light = withBackground(cfg, white);
      const dark = withBackground(cfg, "#1b1b1b");
      expect(JSON.stringify(dark.nodes), label).not.toBe(JSON.stringify(light.nodes));
    }
  });

  it("the canvas paint is still what the config says", () => {
    // Only the style a layout is handed is resolved. The pane and the deck read the
    // canvas from the config, so a transparent chart stays transparent in the SVG and
    // on the slide, and the config object is not rewritten.
    for (const spelling of [...CLEAR, ...UNREADABLE]) {
      const cfg = { ...sampleConfig("stacked"), style: { background: spelling } } as ChartConfig;
      const before = JSON.stringify(cfg);
      const scene = buildChart(cfg);
      expect(JSON.stringify(cfg), `${JSON.stringify(spelling)}: buildChart rewrote its input`).toBe(before);
      const svg = sceneToSvg(scene, { background: cfg.style?.background });
      const canvas = /<rect width="100%" height="100%" fill="([^"]*)"/.exec(svg)?.[1];
      // A paint that draws nothing is written to the canvas as authored, not as the
      // white the layout was resolved against.
      if (CLEAR.includes(spelling)) expect(canvas?.trim().toLowerCase(), spelling).toBe(spelling.trim().toLowerCase());
    }
  });
});

// ---------------------------------------------------------------------------
// C (r2-35) — the hollow forecast marker is a hollow, not a patch in the canvas colour
// ---------------------------------------------------------------------------

describe("a forecast marker is hollow without painting the canvas colour (r2-35)", () => {
  const line = (horizontal: boolean, background: string, extra: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "line",
    ...DEFAULT_SIZE,
    horizontal,
    style: { background },
    decorations: { forecastFrom: 3, segmentLabels: false },
    data: {
      categories: ["A", "B", "C", "D", "E"],
      series: [{ name: "S", values: [10, 12, 15, 17, 20] }],
    },
    ...extra,
  });
  const markers = (cfg: ChartConfig) => named<RectNode>(buildChart(cfg).nodes, "rect", "marker-0-");

  it.each([false, true])("horizontal=%s: forecast points are fill none with the series colour as the ring", (h) => {
    for (const bg of [white, "#123456", "transparent"]) {
      const ms = markers(line(h, bg));
      expect(ms.length).toBe(5);
      const solid = ms[0];
      for (const m of ms.slice(3)) {
        expect(m.fill, `${bg}: forecast marker ${m.name}`).toBe("none");
        expect(m.stroke, `${bg}: ring of ${m.name}`).toBe(solid.fill);
      }
    }
  });

  it.each([false, true])("horizontal=%s: no marker is filled with the canvas colour on any background", (h) => {
    for (const bg of [white, "#123456", "#fafaf0"]) {
      for (const m of markers(line(h, bg))) expect(m.fill, `${bg}: ${m.name}`).not.toBe(bg);
    }
  });

  it.each([false, true])("horizontal=%s: actuals keep their solid marker and the halo ring", (h) => {
    const ms = markers(line(h, "#123456"));
    for (const m of ms.slice(0, 3)) {
      expect(m.fill).not.toBe("none");
      expect(m.stroke, "the halo that keeps a crossing legible is a ring in the canvas colour").toBe("#123456");
    }
  });

  it.each([false, true])("horizontal=%s: a highlighted forecast point keeps its colour", (h) => {
    const cfg = line(h, white);
    cfg.data.series[0].colors = [null, null, null, null, "#e34948"];
    const last = markers(cfg).find((m) => m.name === "marker-0-4")!;
    expect(last.fill).toBe("#e34948");
  });
});

describe("the axis-break band is still an occluder (r2-35, left on purpose)", () => {
  // Pinned so a half-measure does not slip in unnoticed: the band hides the part of
  // every column that crosses the cut, and `fill: "none"` would draw the bars straight
  // through it. See the note on the band in `layout/frame.ts` for what a real gap needs.
  it("fills the cut with the chart's background", () => {
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      style: { background: "#123456" },
      axisBreak: { from: 30, to: 90 },
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [20, 40, 120] }] },
    };
    const band = buildChart(cfg).nodes.find((n): n is RectNode => n.name === "axis-break");
    expect(band?.fill).toBe("#123456");
    const sceneNodes: SceneNode[] = buildChart(cfg).nodes;
    const tall = sceneNodes.find((n): n is RectNode => n.name === "seg-0-2");
    expect(tall && band && tall.y < band.y && tall.y + tall.h > band.y + band.h, "a column crosses the cut").toBe(true);
  });
});
