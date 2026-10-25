import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { textWidth } from "../src/core/scene";
import { endLabelFontSize, onCanvas, roomForEnds, stripWidth, tickWidthFit } from "../src/core/layout/frame";
import type { LineNode, RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * A number the chart prints is printed whole, or not at all.
 *
 * One family: a label sits in a box somebody sized with a constant — 34pt for the
 * value axis, a 48pt box centred on a tick, `fs * 4` for the totals beside a
 * sideways bar, `bar + 12` for a waterfall's value — and the frame clip
 * (`clipTextToFrame`) cuts whatever the constant did not foresee into digits and an
 * ellipsis. "100,0…" reads as a different, smaller number than 100,000, and the
 * reader has no way to know. Nothing was wrong with the DATA in any of these;
 * ordinary six-figure revenue did it.
 *
 * Every test below states the property rather than pinning a coordinate that would
 * only be today's output: the text that comes out is the text that went in, and it
 * stands on the chart. The `fits` guards are the other half — a chart whose labels
 * already fit keeps the geometry it has always had, because the cure for a cut
 * number must not move the charts that never had one.
 */

/**
 * Half a point: what the frame clip lets a label's ink run past the chart before it
 * acts. Written out here rather than imported, so the tests stay an independent
 * statement of the property the code under test has to meet.
 */
const SLACK = 0.5;

const named = <T extends SceneNode>(nodes: SceneNode[], kind: T["kind"], name: string | RegExp): T[] =>
  nodes.filter(
    (n): n is T => n.kind === kind && (typeof name === "string" ? n.name === name : !!n.name && name.test(n.name)),
  );

const textsNamed = (cfg: ChartConfig, name: string | RegExp): TextNode[] =>
  named<TextNode>(buildChart(cfg).nodes, "text", name);

/** The ink a text node puts on the chart: where it begins and ends along x. */
const inkSpan = (t: TextNode): { x0: number; x1: number } => {
  const w = textWidth(t.text, t.fontSize, t.bold);
  const x0 = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
  return { x0, x1: x0 + w };
};

const baselineX = (cfg: ChartConfig): number => {
  const line = named<LineNode>(buildChart(cfg).nodes, "line", "baseline")[0];
  return Math.min(line.x1, line.x2);
};

/** A label is cut when it ends in the ellipsis the frame clip leaves. */
const cut = (t: TextNode) => /…$/.test(t.text);

const column = (values: number[], extra: Partial<ChartConfig> = {}): ChartConfig => ({
  kind: "clustered",
  ...DEFAULT_SIZE,
  data: { categories: ["A", "B", "C"], series: [{ name: "Revenue", values }] },
  decorations: { valueAxis: true, categoryAxis: true },
  ...extra,
});

describe("the value axis's column is sized to the labels it holds (r3-4, r3-7, r3-15)", () => {
  const ticks = (cfg: ChartConfig) => textsNamed(cfg, "value-axis").map((t) => t.text);

  it("prints a tick of six figures whole, where it was cut to '100,0…'", () => {
    // The column was a flat 34pt and the box 30, so any label wider than 30pt at
    // the 9pt tick font — 100,000 is 34 — was cut by the frame clip.
    expect(ticks(column([120000, 340000, 210000]))).toEqual(["0", "100,000", "200,000", "300,000", "400,000"]);
  });

  it("does not shrink the labels to do it: the column gives up the width", () => {
    // The preferred remedy, in the order the owner set: size the box from the
    // measured text and let the plot give up the width, shrinking only when the
    // cap on that cannot hold the label.
    const cfg = column([120000, 340000, 210000]);
    for (const t of textsNamed(cfg, "value-axis")) expect(t.fontSize).toBe(9);
    // The widest label plus the box's own 4pt margin and two points of air, no more.
    expect(baselineX(cfg)).toBeCloseTo(textWidth("400,000", 9) + 6, 9);
  });

  it("covers a five-digit negative, which is seven characters", () => {
    expect(ticks(column([-80000, -45000, -60000]))).toEqual(["-80,000", "-60,000", "-40,000", "-20,000", "0"]);
  });

  it("covers a unit suffix, which was lost on 4-digit ticks", () => {
    // " kr" is three characters: 1,000 kr is 38.9pt, and it printed '1,000…'.
    const t = ticks(column([1200, 3400, 2100], { numberFormat: { suffix: " kr" } }));
    expect(t).toEqual(["0 kr", "1,000 kr", "2,000 kr", "3,000 kr", "4,000 kr"]);
  });

  it("covers a locale whose cut reads as a different number: '100.0…' for 100.000", () => {
    const t = ticks(column([120000, 340000, 210000], { numberFormat: { locale: "de-DE" } }));
    expect(t).toEqual(["0", "100.000", "200.000", "300.000", "400.000"]);
  });

  it("covers the Tufte dashes, whose label box is two points narrower", () => {
    const t = ticks(column([120000, 150000], { decorations: { valueAxis: "datamarks", categoryAxis: true } }));
    expect(t).toEqual(["0", "50,000", "100,000", "150,000"]);
  });

  it("covers a share axis at two decimals", () => {
    const cfg: ChartConfig = {
      kind: "stacked100",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "x", values: [3, 4] },
          { name: "y", values: [5, 6] },
        ],
      },
      decorations: { valueAxis: true, categoryAxis: true },
      numberFormat: { decimals: 2 },
    };
    expect(ticks(cfg)).toEqual(["0.00%", "25.00%", "50.00%", "75.00%", "100.00%"]);
  });

  it("covers the waterfall and the combo's primary axis, which size their frames themselves", () => {
    const waterfall: ChartConfig = {
      kind: "waterfall",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Start", "Up", "Down", "End"],
        series: [{ name: "Delta", values: [500000, 300000, -120000, 0] }],
      },
      waterfall: { totalIndices: [3] },
      decorations: { valueAxis: true, categoryAxis: true },
    };
    expect(ticks(waterfall).some((t) => /,000$/.test(t) && t.length >= 7)).toBe(true);
    expect(ticks(waterfall).filter((t) => cut({ text: t } as TextNode))).toEqual([]);
    const combo: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Revenue", values: [120000, 340000, 210000] },
          { name: "Margin", values: [0.1, 0.2, 0.3], type: "line" },
        ],
      },
      decorations: { valueAxis: true, categoryAxis: true },
    };
    expect(ticks(combo)).toEqual(["0", "100,000", "200,000", "300,000", "400,000"]);
  });

  it("keeps the 34pt column while every label fits it, to the last bit", () => {
    // The guard on the other side: 10,000 is 29.2pt and always fitted. The cure
    // for a cut number must not move the chart that never had one.
    const cfg = column([12000, 34000, 21000]);
    expect(ticks(cfg)).toEqual(["0", "10,000", "20,000", "30,000", "40,000"]);
    expect(baselineX(cfg)).toBe(34);
    for (const t of textsNamed(cfg, "value-axis")) {
      expect(t.w).toBe(30);
      expect(t.fontSize).toBe(9);
    }
  });

  it("holds at every size and font: the column is capped, and a label the cap cannot hold is left out", () => {
    // 12 figures on a 120pt chart: the column may take 40% of the width (48pt),
    // which 1,000,000,000,000 (63pt at 9pt) does not fit. The axis shrinks together
    // down to the floor, and what still does not fit is not drawn — never cut.
    const cfg = column([1.2e12, 3.4e12, 2.1e12], { width: 120, height: 150 });
    const labels = textsNamed(cfg, "value-axis");
    expect(labels.length).toBeGreaterThan(0);
    for (const t of labels) {
      expect(cut(t)).toBe(false);
      expect(inkSpan(t).x0).toBeGreaterThanOrEqual(-SLACK);
      expect(t.fontSize).toBeGreaterThanOrEqual(5);
    }
    expect(baselineX(cfg)).toBeLessThanOrEqual(120 * 0.4 + 1e-9);
  });

  it("holds for every kind that draws the axis, at every magnitude the starter data can be", () => {
    const kinds: ChartConfig["kind"][] = ["clustered", "stacked", "waterfall", "line", "area", "boxplot"];
    for (const kind of kinds) {
      for (const k of [1, 1e3, 1e5, 1e6, 1e9]) {
        for (const valueAxis of [true, "datamarks"] as const) {
          const series =
            kind === "boxplot"
              ? ["Min", "Q1", "Median", "Q3", "Max"].map((name, i) => ({
                  name,
                  values: [10, 20].map((v) => v * (i + 1) * k),
                }))
              : [{ name: "S", values: [2, 5, 3].map((v) => v * 100 * k) }];
          const cfg: ChartConfig = {
            kind,
            ...DEFAULT_SIZE,
            data: { categories: kind === "boxplot" ? ["A", "B"] : ["A", "B", "C"], series },
            waterfall: kind === "waterfall" ? { totalIndices: [2] } : undefined,
            decorations: { valueAxis, categoryAxis: true },
          };
          for (const t of textsNamed(cfg, "value-axis")) {
            expect(cut(t), `${kind} x${k} ${valueAxis}: ${t.text}`).toBe(false);
            expect(inkSpan(t).x0, `${kind} x${k} ${valueAxis}: ${t.text}`).toBeGreaterThanOrEqual(-SLACK);
          }
        }
      }
    }
  });

  it("covers the scatter's y axis, whose column was the same flat 34", () => {
    const cfg: ChartConfig = {
      kind: "scatter",
      ...DEFAULT_SIZE,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 2, 3] },
          { name: "Y", values: [100000, 500000, 800000] },
        ],
      },
    };
    const labels = textsNamed(cfg, "y-axis").map((t) => t.text);
    // The topmost tick (800,000) would poke above the canvas on an untitled chart and
    // is left out by the axis's own title/canvas rule; the rest are whole.
    expect(labels).toEqual(["0", "200,000", "400,000", "600,000"]);
    expect(baselineX(cfg)).toBeCloseTo(textWidth("600,000", 9) + 6, 9);
  });

  it("covers the combo's secondary axis, whose strip was pinned 34pt wide at the canvas edge", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      secondaryAxis: true,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Revenue", values: [10, 20, 30] },
          { name: "Line", type: "line", values: [12000000, 34000000, 21000000] },
        ],
      },
      // No series names beside the plot: it runs to the canvas, and the strip is
      // pinned at `width - 34`, which is where the number was cut.
      decorations: { valueAxis: true, categoryAxis: true, seriesLabels: false },
    };
    const strip = textsNamed(cfg, "secondary-axis");
    expect(strip.length).toBeGreaterThan(1);
    for (const t of strip) {
      expect(cut(t)).toBe(false);
      expect(inkSpan(t).x1).toBeLessThanOrEqual(cfg.width + SLACK);
    }
    expect(strip.map((t) => t.text)).toContain("40,000,000");
  });
});

describe("a tick label at the end of an axis stays on the chart (r3-3, r3-9, r3-16)", () => {
  const sideways = (values: number[], extra: Partial<ChartConfig> = {}): ChartConfig =>
    column(values, { horizontal: true, ...extra });
  const scatter = (xs: number[], extra: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "scatter",
    ...DEFAULT_SIZE,
    data: {
      categories: xs.map((_, i) => `P${i}`),
      series: [
        { name: "X", values: xs },
        { name: "Y", values: xs.map((_, i) => i + 1) },
      ],
    },
    ...extra,
  });
  const onChart = (cfg: ChartConfig, name: string) => {
    for (const t of textsNamed(cfg, name)) {
      const { x0, x1 } = inkSpan(t);
      expect(cut(t), `${name} ${t.text}`).toBe(false);
      expect(x0, `${name} ${t.text}`).toBeGreaterThanOrEqual(-SLACK);
      expect(x1, `${name} ${t.text}`).toBeLessThanOrEqual(cfg.width + SLACK);
    }
  };

  it("prints a sideways axis's maximum whole: '4,0…' for 4,000", () => {
    // The last tick sits at the plot's right edge, 10pt from the canvas, and a
    // 48pt box centred on it hangs half a label past the chart.
    const cfg = sideways([1200, 3400]);
    expect(textsNamed(cfg, "value-axis").map((t) => t.text)).toEqual(["0", "1,000", "2,000", "3,000", "4,000"]);
    onChart(cfg, "value-axis");
  });

  it("holds at every width and with the dashes", () => {
    for (const width of [320, 480, 960]) {
      for (const valueAxis of [true, "datamarks"] as const) {
        const cfg = sideways([1200, 3400], { width, height: 300, decorations: { valueAxis, categoryAxis: true } });
        expect(textsNamed(cfg, "value-axis").at(-1)!.text).toBe("4,000");
        onChart(cfg, "value-axis");
      }
    }
  });

  it("prints a sideways axis's FIRST tick whole when the category gutter is narrow", () => {
    // Without category names the gutter is 2pt, so a centred '-100,000' hung 17pt
    // off the left edge — a negative-data chart printed '-100…'.
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      horizontal: true,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "x", values: [-80000, 20000] },
          { name: "y", values: [-5000, 40000] },
        ],
      },
      decorations: { valueAxis: true, categoryAxis: false },
    };
    expect(textsNamed(cfg, "value-axis").map((t) => t.text)).toContain("-100,000");
    onChart(cfg, "value-axis");
  });

  it("covers the sideways boxplot, which shares the chrome", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      horizontal: true,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Min", values: [100, 200] },
          { name: "Q1", values: [1000, 2000] },
          { name: "Median", values: [5000, 6000] },
          { name: "Q3", values: [10000, 11000] },
          { name: "Max", values: [14000, 15000] },
        ],
      },
      decorations: { valueAxis: true, categoryAxis: true },
    };
    expect(textsNamed(cfg, "value-axis").at(-1)!.text).toBe("15,000");
    onChart(cfg, "value-axis");
  });

  it("prints the scatter's x maximum whole, with no option switched on", () => {
    // Only the FIRST tick was nudged (off the y axis); the last hung 4pt past the
    // canvas and printed '8,…'. Any x range reaching 1,000 did it.
    const cfg = scatter([1200, 3000, 7400]);
    expect(textsNamed(cfg, "x-axis").map((t) => t.text)).toEqual(["0", "2,000", "4,000", "6,000", "8,000"]);
    onChart(cfg, "x-axis");
    onChart({ ...cfg, kind: "bubble" }, "x-axis");
  });

  it("prints the scatter's x maximum at two decimals: '80…' for 80.00", () => {
    const cfg = scatter([12.25, 40, 74.5], { numberFormat: { decimals: 2 } });
    expect(textsNamed(cfg, "x-axis").at(-1)!.text).toBe("80.00");
    onChart(cfg, "x-axis");
  });

  it("moves an end label by exactly its overhang, and no other label at all", () => {
    const cfg = scatter([1200, 3000, 7400]);
    const labels = textsNamed(cfg, "x-axis");
    const last = labels.at(-1)!;
    // On the canvas edge, not a point further: the ink ends where the chart does.
    expect(inkSpan(last).x1).toBeCloseTo(cfg.width, 6);
    // The inner ticks stay on their own ticks: equal spacing is the scale.
    const centres = labels.slice(1, -1).map((t) => t.x + t.w / 2);
    const steps = centres.slice(1).map((c, i) => c - centres[i]);
    for (const s of steps) expect(s).toBeCloseTo(steps[0], 6);
  });

  it("keeps a label that fits exactly on its tick (sideways, three digits)", () => {
    const cfg = sideways([100, 340], { decorations: { valueAxis: true, categoryAxis: true, gridlines: true } });
    const grid = named<LineNode>(buildChart(cfg).nodes, "line", "gridline").map((l) => l.x1);
    const labels = textsNamed(cfg, "value-axis");
    // Every tick but the zero baseline has a gridline, and every label is centred on it.
    for (const x of grid) expect(labels.some((t) => Math.abs(t.x + t.w / 2 - x) < 1e-9)).toBe(true);
  });

  it("still clears its neighbour after the nudge: the clamp does not undo the fit", () => {
    // A 5-digit axis on a small chart: the last label moves left, toward the one
    // beside it, so the strip has to be fitted where the label ENDS UP.
    for (const cfg of [
      scatter([12000, 30000, 74000], { width: 320, height: 200 }),
      sideways([12000, 34000], { width: 320, height: 200 }),
    ]) {
      const name = cfg.kind === "scatter" ? "x-axis" : "value-axis";
      const spans = textsNamed(cfg, name).map(inkSpan);
      for (let i = 1; i < spans.length; i++) expect(spans[i].x0 - spans[i - 1].x1).toBeGreaterThanOrEqual(1 - 1e-6);
      onChart(cfg, name);
    }
  });

  it("covers a sideways secondary axis that began on the canvas and ran off the right", () => {
    // The strip is moved onto the canvas as a block, but it was asked whether it
    // fit only when it hung off BOTH ends — so one that began 2pt in and ran 7pt off
    // the right was moved left by the overhang, and its first label went off the
    // left: '-10…' for -100.
    const cfg: ChartConfig = {
      kind: "combo",
      width: 960,
      height: 300,
      horizontal: true,
      secondaryAxis: true,
      numberFormat: { suffix: " m", locale: "fr-FR" },
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "A", values: [-89.51, 77.03] },
          { name: "B", values: [-75.69, -69.52] },
          { name: "Line", type: "line", values: [98.71, -71.29] },
        ],
      },
      decorations: { valueAxis: false, categoryAxis: true },
    };
    for (const t of textsNamed(cfg, "secondary-axis")) expect(cut(t), t.text).toBe(false);
  });
});

describe("an upright waterfall's value label is boxed as wide as it was fitted (r3-5, r3-18)", () => {
  const bridge = (values: number[], extra: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "waterfall",
    ...DEFAULT_SIZE,
    data: { categories: values.map((_, i) => `C${i}`), series: [{ name: "Delta", values }] },
    waterfall: { totalIndices: [values.length - 1] },
    decorations: { categoryAxis: true },
    ...extra,
  });
  const labels = (cfg: ChartConfig) => textsNamed(cfg, /^label-\d+$/);
  const eight = [8600000, 1400000, 900000, -1200000, -400000, 1000000, 600000, 0];

  it("prints every label of an eight-column bridge whole: '+1,400,0…' for +1,400,000", () => {
    // The font was fitted to the SLOT (59pt, so 57 of room) and the label then
    // clipped to a box of bar + 12 (51.5pt) — a fit and a clamp that disagree, and
    // the clamp is what the reader sees. '+1,400,000' is 54pt at 10pt.
    expect(labels(bridge(eight)).map((t) => t.text)).toEqual([
      "8,600,000",
      "+1,400,000",
      "+900,000",
      "-1,200,000",
      "-400,000",
      "+1,000,000",
      "+600,000",
      "10,900,000",
    ]);
  });

  it("boxes a label that needed it as wide as the label, and never wider than the slot", () => {
    const cfg = bridge(eight);
    const scene = buildChart(cfg);
    const slot = 59.25;
    const wide = named<TextNode>(scene.nodes, "text", "label-1")[0];
    // The whole label, not what is left of it: its width is measured below.
    expect(wide.text).toBe("+1,400,000");
    expect(wide.w).toBeGreaterThanOrEqual(textWidth(wide.text, wide.fontSize) + 2);
    expect(wide.w).toBeLessThanOrEqual(slot + 1e-9);
    // And its ink is inside its own box, which is what stops the clip.
    const { x0, x1 } = inkSpan(wide);
    expect(x0).toBeGreaterThanOrEqual(wide.x - 1e-9);
    expect(x1).toBeLessThanOrEqual(wide.x + wide.w + 1e-9);
  });

  it("does not change the size: the labels keep the chart font", () => {
    for (const t of labels(bridge(eight))) expect(t.fontSize).toBe(10);
  });

  it("leaves a label that already fits its box exactly where it was: bar + 12", () => {
    const cfg = bridge([86, 14, 9, -12, -4, 0]);
    const scene = buildChart(cfg);
    for (const c of [0, 1, 2, 3]) {
      const bar = named<RectNode>(scene.nodes, "rect", `bar-${c}`)[0];
      const label = named<TextNode>(scene.nodes, "text", `label-${c}`)[0];
      expect(label.w).toBeCloseTo(bar.w + 12, 9);
    }
  });

  it("leaves a number out rather than cutting it where even the floor does not fit the slot", () => {
    const thirty = Array.from({ length: 30 }, (_, i) => (i % 2 ? -1 : 1) * (100000000 + i * 1234567));
    const drawn = labels(bridge(thirty, { width: 320 }));
    for (const t of drawn) {
      expect(cut(t), t.text).toBe(false);
      expect(t.text).toMatch(/^[+-]?[\d,]+$/);
    }
  });

  it("covers a stacked waterfall, which draws through the same box", () => {
    const cfg: ChartConfig = {
      kind: "waterfall",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C", "D", "E", "F", "G", "H"],
        series: [
          { name: "x", values: [8000000, 1400000, 900000, 1200000, 400000, 1000000, 600000, 0] },
          { name: "y", values: [600000, 0, 0, 0, 0, 0, 0, 0] },
        ],
      },
      waterfall: { totalIndices: [7] },
      decorations: { categoryAxis: true },
    };
    for (const t of labels(cfg)) expect(cut(t), t.text).toBe(false);
  });
});

describe("a total at the end of a sideways bar has the room it needs (r3-19)", () => {
  const mekko = (extra: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "mekko",
    ...DEFAULT_SIZE,
    horizontal: true,
    data: {
      categories: ["EMEA", "Americas", "APAC"],
      series: [
        { name: "A", values: [1200000, 1500000, 800000] },
        { name: "B", values: [900000, 1100000, 700000] },
      ],
    },
    decorations: { totals: true, categoryAxis: true, segmentLabels: true },
    ...extra,
  });
  const totals = (cfg: ChartConfig) => textsNamed(cfg, /^total-\d+$/);

  it("prints every row total of a horizontal mekko whole: '2,100…' for 2,100,000", () => {
    // Every row of a %-axis mekko runs to the plot's edge, and the strip beside it
    // was `fs * 4` — 39pt of room for a 52pt number.
    const cfg = mekko();
    expect(totals(cfg).map((t) => t.text)).toEqual(["2,100,000", "2,600,000", "1,500,000"]);
    for (const t of totals(cfg)) {
      expect(inkSpan(t).x1).toBeLessThanOrEqual(cfg.width + SLACK);
      // The plot gave up the width; the number was not shrunk to fit a strip that stayed.
      expect(t.fontSize).toBe(10);
    }
  });

  it("gives up the width the totals ask for, and no more: the number ends 2pt inside the chart", () => {
    const cfg = mekko();
    const scene = buildChart(cfg);
    const edge = Math.max(...named<RectNode>(scene.nodes, "rect", /^seg-/).map((r) => r.x + r.w));
    const widest = Math.max(...totals(cfg).map((t) => textWidth(t.text, t.fontSize, true)));
    expect(edge + 3 + widest).toBeCloseTo(cfg.width - 2, 6);
  });

  it("leaves a mekko whose totals fit its plot exactly as it was", () => {
    const cfg = mekko({
      data: {
        categories: ["EMEA", "Americas", "APAC"],
        series: [
          { name: "A", values: [42, 55, 18] },
          { name: "B", values: [28, 30, 22] },
        ],
      },
    });
    const scene = buildChart(cfg);
    const edge = Math.max(...named<RectNode>(scene.nodes, "rect", /^seg-/).map((r) => r.x + r.w));
    // `cfg.width - totalsW - 2` with `totalsW = fs * 4`: the strip it always had.
    expect(edge).toBeCloseTo(cfg.width - 40 - 2, 9);
  });

  it("makes room for the one row that needs it in the mekko with units, and only that much", () => {
    // Rows are as long as their totals, so only the longest reaches the edge; the
    // shorter ones have room beside them already.
    const cfg = mekko({
      data: {
        categories: ["EMEA", "Americas", "APAC"],
        series: [{ name: "A", values: [1200000, 2500000, 300000] }],
        xExtent: [1, 1, 1],
      },
    });
    for (const t of totals(cfg)) expect(cut(t), t.text).toBe(false);
    expect(totals(cfg).map((t) => t.text)).toEqual(["1,200,000", "2,500,000", "300,000"]);
  });

  it("prints the longest bar's total of a sideways stacked chart whole", () => {
    // 4,990,000 on a 5,000,000 axis ends 0.2% short of the plot's edge, which is
    // 42pt from the canvas: the 52pt total began 3pt past it and was cut.
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      horizontal: true,
      data: { categories: ["A", "B", "C"], series: [{ name: "x", values: [4990000, 2600000, 1500000] }] },
      decorations: { totals: true, categoryAxis: true },
    };
    expect(totals(cfg).map((t) => t.text)).toEqual(["4,990,000", "2,600,000", "1,500,000"]);
    for (const t of totals(cfg)) {
      expect(inkSpan(t).x1).toBeLessThanOrEqual(cfg.width + SLACK);
      // At the chart's own size: the plot made the room, the number did not give way.
      expect(t.fontSize).toBe(10);
    }
  });

  it("does not move a sideways stacked chart whose longest total already has the room", () => {
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      horizontal: true,
      data: { categories: ["A", "B", "C"], series: [{ name: "x", values: [490000, 260000, 150000] }] },
      decorations: { totals: true, categoryAxis: true },
    };
    const scene = buildChart(cfg);
    const edge = Math.max(...named<RectNode>(scene.nodes, "rect", /^seg-/).map((r) => r.x + r.w));
    expect(edge).toBeLessThanOrEqual(cfg.width - 40 - 2 + 1e-9);
    // The scale tops out at 500,000, so the plot is the usual one to the last bit.
    expect(scene.nodes.some((n) => n.kind === "rect")).toBe(true);
  });

  it("shrinks or leaves out what the capped strip cannot hold, and never cuts it", () => {
    const cfg = mekko({
      width: 120,
      height: 150,
      data: {
        categories: ["EMEA", "Americas", "APAC"],
        series: [{ name: "A", values: [9.1e11, 8.2e11, 7.3e11] }],
      },
    });
    for (const t of totals(cfg)) {
      expect(cut(t), t.text).toBe(false);
      expect(inkSpan(t).x1).toBeLessThanOrEqual(cfg.width + SLACK);
    }
  });
});

describe("a sideways clustered-stacked chart totals each stack, not all of them at once (r3-8)", () => {
  const stacks = (horizontal: boolean, extra: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "stacked",
    ...DEFAULT_SIZE,
    horizontal,
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "A", values: [10, 12], stack: 0 },
        { name: "B", values: [5, 6], stack: 0 },
        { name: "Plan", values: [40, 44], stack: 1 },
      ],
    },
    decorations: { totals: true, categoryAxis: true },
    ...extra,
  });
  const totals = (cfg: ChartConfig) => textsNamed(cfg, /^total-/);

  it("prints one total per stack, as the upright chart does", () => {
    const up = totals(stacks(false)).map((t) => [t.name, t.text]);
    const side = totals(stacks(true)).map((t) => [t.name, t.text]);
    expect(up).toEqual([
      ["total-0-s0", "15"],
      ["total-0-s1", "40"],
      ["total-1-s0", "18"],
      ["total-1-s1", "44"],
    ]);
    expect(side).toEqual(up);
  });

  it("no longer prints the grand total (55) beside the tallest stack, a number that matched no bar", () => {
    expect(totals(stacks(true)).map((t) => t.text)).not.toContain("55");
    expect(totals(stacks(true)).map((t) => t.text)).not.toContain("62");
  });

  it("stands each total at the end of the bar it totals, on that bar's own row", () => {
    const cfg = stacks(true);
    const scene = buildChart(cfg);
    for (const t of totals(cfg)) {
      const sp = t.name!.endsWith("s0") ? 0 : 1;
      const c = t.name!.startsWith("total-0") ? 0 : 1;
      // The stack's last segment: series 1 ends stack 0, series 2 is stack 1.
      const bar = named<RectNode>(scene.nodes, "rect", sp === 0 ? `seg-1-${c}` : `seg-2-${c}`)[0];
      expect(t.x).toBeCloseTo(bar.x + bar.w + 3, 9);
      expect(t.y + t.h / 2).toBeCloseTo(bar.y + bar.h / 2, 6);
    }
  });

  it("leaves a blank stack unlabelled, as upright: a sum of nothing is not a measured zero", () => {
    const cfg = stacks(true, {
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "A", values: [10, 12], stack: 0 },
          { name: "Plan", values: [40, null], stack: 1 },
        ],
      },
    });
    expect(totals(cfg).map((t) => t.name)).toEqual(["total-0-s0", "total-0-s1", "total-1-s0"]);
  });

  it("leaves a single-stack sideways chart alone: one total per category", () => {
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      horizontal: true,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "A", values: [10, 12] },
          { name: "B", values: [5, 6] },
        ],
      },
      decorations: { totals: true, categoryAxis: true },
    };
    expect(totals(cfg).map((t) => [t.name, t.text])).toEqual([
      ["total-0", "15"],
      ["total-1", "18"],
    ]);
  });
});

describe("the fitting helpers", () => {
  it("stripWidth: the usual width while the label fits what the strip leaves it, the text plus the margin when not", () => {
    // 30pt of room, clip slack half a point: 30.5 still fits, 30.6 does not.
    expect(stripWidth(34, 30, 30.5, 4, 100)).toBe(34);
    // The ink, the strip's own 4pt margin, and two points of air at the other edge.
    expect(stripWidth(34, 30, 30.6, 4, 100)).toBeCloseTo(36.6, 9);
    expect(stripWidth(34, 30, 60, 4, 100)).toBe(66);
    // Capped at a share of the chart; never narrower than the usual column.
    expect(stripWidth(34, 30, 400, 4, 100)).toBe(100);
    expect(stripWidth(34, 30, 400, 4, 20)).toBe(34);
  });

  it("onCanvas: moves an end label by exactly its overhang, and leaves one the clip would leave alone", () => {
    expect(onCanvas(100, 12, 480)).toBe(100);
    expect(onCanvas(472, 12.15, 480)).toBeCloseTo(480 - 12.15, 9);
    expect(onCanvas(8, 12.15, 480)).toBeCloseTo(12.15, 9);
    // Within the clip's half-point of slack, it does not move.
    expect(onCanvas(2, 2.4, 480)).toBe(2);
    expect(onCanvas(478, 2.4, 480)).toBe(478);
  });

  it("endLabelFontSize: the font while the label stands on the chart, shrunk to the room, 0 below the floor", () => {
    expect(endLabelFontSize("52,000", true, 10, 100, 480)).toBe(10);
    const shrunk = endLabelFontSize("2,100,000,000", true, 10, 420, 480);
    expect(shrunk).toBeGreaterThanOrEqual(5);
    expect(shrunk).toBeLessThan(10);
    expect(420 + 3 + textWidth("2,100,000,000", shrunk, true)).toBeLessThanOrEqual(480 + SLACK);
    expect(endLabelFontSize("2,100,000,000", true, 10, 470, 480)).toBe(0);
  });

  it("roomForEnds: narrows the plot only for a label that does not fit, and returns the frame itself otherwise", () => {
    const cfg = { ...column([1, 2]), width: 480, height: 300 };
    const frame = { x: 20, y: 10, w: 418, h: 200 };
    expect(roomForEnds(cfg, frame, [{ ink: 30, reach: 1 }])).toBe(frame);
    const narrowed = roomForEnds(cfg, frame, [{ ink: 52, reach: 1 }]);
    expect(narrowed.w).toBeCloseTo(480 - 2 - 20 - 3 - 52, 9);
    expect(narrowed.x).toBe(20);
    expect(narrowed.y).toBe(10);
    // A bar that ends short of the edge has the room already.
    expect(roomForEnds(cfg, frame, [{ ink: 52, reach: 0.5 }])).toBe(frame);
  });

  it("tickWidthFit: 1 for every label that fits, shrinks together otherwise, drops what the floor cannot hold", () => {
    expect(tickWidthFit(["0", "10,000"], 9, 30)).toEqual({ scale: 1, keep: [true, true] });
    const fit = tickWidthFit(["0", "100,000"], 9, 30);
    expect(fit.keep).toEqual([true, true]);
    expect(fit.scale).toBeCloseTo(30 / textWidth("100,000", 9), 9);
    const dropped = tickWidthFit(["0", "1,000,000,000,000"], 9, 20);
    expect(dropped.keep).toEqual([true, false]);
  });
});
