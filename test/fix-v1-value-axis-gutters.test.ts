import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { textWidth } from "../src/core/scene";
import type { LineNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * The value axis's column is sized to the numbers it holds — in the line, area,
 * boxplot, candlestick and violin layouts too.
 *
 * `computeFrame` widens the value-axis column to the widest tick label when the
 * layout hands it the scale builder (9f5f591, r3-4 / r3-7 / r3-15: 100,000 was cut
 * to "100,0…" by a flat 34pt column). The column, waterfall, scatter and combo
 * layouts passed one; these five did not, so they got the fallback `chromeNodes`
 * gives a column that was never sized — shrink the whole axis until its widest
 * label fits the 30pt box, then drop what will not fit even at 5pt. No number was
 * cut, but a chart whose ticks read 2,000,000 drew them at 6pt next to a 34pt
 * column (a tick of seven figures is 44pt at the 9pt it should be) and, with more
 * digits on a narrow chart, drew only the "0". The column chart beside it printed
 * the same numbers whole at 9pt.
 *
 * Every test states the property rather than a coordinate that would only be
 * today's output: the labels that come out are the ticks the axis has, at the size
 * every axis draws at, standing on the chart and clear of the plot. The control is
 * the other half — a chart whose numbers already fit keeps the 34pt column it has
 * always had, because widening for a number that never needed it moves every
 * chart that was fine.
 */

/** Half a point: what the frame clip lets a label's ink run past its box before it acts. */
const SLACK = 0.5;
/** The value-axis column of a chart whose labels fit it — written out, not imported. */
const USUAL_COLUMN = 34;
/** The label box is `frame.x - 4` wide: 4pt of margin between the label and the plot. */
const MARGIN = 4;
/** The tick font at the default 10pt chart font: every axis draws at 0.9 of it. */
const TICK_FS = 9;
/** Below this the engine leaves a label out rather than draw it. */
const FLOOR_FS = 5;

const named = <T extends SceneNode>(nodes: SceneNode[], kind: T["kind"], name: string | RegExp): T[] =>
  nodes.filter(
    (n): n is T => n.kind === kind && (typeof name === "string" ? n.name === name : !!n.name && name.test(n.name)),
  );

const axisLabels = (cfg: ChartConfig): TextNode[] => named<TextNode>(buildChart(cfg).nodes, "text", "value-axis");

/** The ink a text node puts on the chart: where it begins and ends along x. */
const inkSpan = (t: TextNode): { x0: number; x1: number } => {
  const w = textWidth(t.text, t.fontSize, t.bold);
  const x0 = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
  return { x0, x1: x0 + w };
};

/** A label is cut when it ends in the ellipsis the frame clip leaves. */
const cut = (t: TextNode) => /…$/.test(t.text);

/**
 * The plot's left and right edge, read off the gridlines (which run the plot's
 * width) rather than off the label boxes the column is sized to, so the property
 * is measured on the thing the reader sees and not on the thing under test.
 */
const plot = (cfg: ChartConfig): { x: number; right: number } => {
  const grid = named<LineNode>(buildChart(cfg).nodes, "line", "gridline");
  expect(grid.length, "the chart drew no gridline to read the plot's edge from").toBeGreaterThan(0);
  return { x: Math.min(grid[0].x1, grid[0].x2), right: Math.max(grid[0].x1, grid[0].x2) };
};

const cats = ["A", "B", "C"];
const row = (name: string, values: number[], k: number) => ({ name, values: values.map((v) => v * k) });

/**
 * One chart per kind, written in thousands and scaled by `k`: at k = 1 the ticks
 * read up to 2,000 and fit the usual column, at k = 1,000 they read up to
 * 2,000,000, and at k = 1e10 they have fourteen digits.
 */
interface Subject {
  at: (k: number, over?: Partial<ChartConfig>) => ChartConfig;
  /** The ticks at k = 1 and at k = 1,000, as the axis prints them. */
  small: string[];
  wide: string[];
}

const make = (
  kind: ChartConfig["kind"],
  series: (k: number) => ChartConfig["data"]["series"],
  small: string[],
  wide: string[],
): Subject => ({
  small,
  wide,
  at: (k, over = {}) => ({
    kind,
    width: 480,
    height: 300,
    data: { categories: cats, series: series(k) },
    ...over,
    decorations: { valueAxis: true, categoryAxis: true, gridlines: true, ...over.decorations },
  }),
});

const zeroBased = ["0", "500", "1,000", "1,500", "2,000"];
const zeroBasedWide = ["0", "500,000", "1,000,000", "1,500,000", "2,000,000"];

const SUBJECTS: Record<string, Subject> = {
  line: make("line", (k) => [row("Revenue", [1200, 2000, 1600], k)], zeroBased, zeroBasedWide),
  // The stack, not the tallest series, reaches 2,000,000.
  area: make(
    "area",
    (k) => [row("North", [900, 1200, 1000], k), row("South", [600, 800, 700], k)],
    zeroBased,
    zeroBasedWide,
  ),
  boxplot: make(
    "boxplot",
    (k) => [
      row("Min", [200, 300, 100], k),
      row("Q1", [400, 500, 300], k),
      row("Median", [600, 800, 500], k),
      row("Q3", [1000, 1200, 800], k),
      row("Max", [1600, 2000, 1400], k),
    ],
    zeroBased,
    zeroBasedWide,
  ),
  // No forced zero: these two draw the data's own range.
  candlestick: make(
    "candlestick",
    (k) => [
      row("Open", [1800, 1850, 1900], k),
      row("High", [1900, 1950, 2000], k),
      row("Low", [1750, 1800, 1850], k),
      row("Close", [1850, 1900, 1950], k),
    ],
    ["1,700", "1,800", "1,900", "2,000"],
    ["1,700,000", "1,800,000", "1,900,000", "2,000,000"],
  ),
  violin: make(
    "violin",
    (k) => Array.from({ length: 8 }, (_, i) => row(`s${i}`, [900 + i * 40, 1200 + i * 60, 1500 + i * 70], k)),
    ["500", "1,000", "1,500", "2,000"],
    ["500,000", "1,000,000", "1,500,000", "2,000,000"],
  ),
};

for (const [kind, s] of Object.entries(SUBJECTS)) {
  describe(`${kind}: the value axis's column is sized to its ticks`, () => {
    it("prints every tick whole at the normal size, where it was shrunk to 6pt", () => {
      const labels = axisLabels(s.at(1e3));
      expect(labels.map((t) => t.text)).toEqual(s.wide);
      for (const t of labels) {
        expect(t.fontSize, t.text).toBe(TICK_FS);
        expect(cut(t), t.text).toBe(false);
        expect(inkSpan(t).x0, t.text).toBeGreaterThanOrEqual(-SLACK);
      }
    });

    it("gives the plot's width up for it: the left edge moves, the right edge does not", () => {
      const small = plot(s.at(1));
      const wide = plot(s.at(1e3));
      expect(wide.x).toBeGreaterThan(small.x);
      expect(wide.right).toBeCloseTo(small.right, 9);
      // And the labels stand clear of the plot by the box's own margin, which is
      // what the gutter was sized for.
      for (const t of axisLabels(s.at(1e3))) expect(inkSpan(t).x1, t.text).toBeLessThanOrEqual(wide.x - MARGIN + SLACK);
    });

    it("keeps the usual 34pt column, to the last bit, while every label fits it", () => {
      const cfg = s.at(1);
      const labels = axisLabels(cfg);
      expect(labels.map((t) => t.text)).toEqual(s.small);
      expect(plot(cfg).x).toBe(USUAL_COLUMN);
      for (const t of labels) {
        expect(t.w, t.text).toBe(USUAL_COLUMN - MARGIN);
        expect(t.fontSize, t.text).toBe(TICK_FS);
      }
    });

    it("covers the Tufte dashes, whose label box is two points narrower", () => {
      const cfg = s.at(1e3, { decorations: { valueAxis: "datamarks" } as ChartConfig["decorations"] });
      const labels = axisLabels(cfg);
      expect(labels.map((t) => t.text)).toEqual(s.wide);
      for (const t of labels) {
        expect(t.fontSize, t.text).toBe(TICK_FS);
        // The dash takes 4pt of the margin and the box stands 6pt left of the plot.
        expect(inkSpan(t).x1, t.text).toBeLessThanOrEqual(plot(cfg).x - 6 + SLACK);
      }
    });

    it("holds when the column hits its cap: shrunk together, none cut, and the cap stays 40%", () => {
      // Fourteen digits on a 200pt chart: the column may take 40% (80pt) and the
      // widest label is 87pt at 9pt. The cap cannot hold it at full size, so the
      // axis shrinks together — the fallback the widening leaves in place.
      const cfg = s.at(1e10, { width: 200 });
      const labels = axisLabels(cfg);
      expect(labels.length).toBeGreaterThan(0);
      expect(plot(cfg).x).toBeCloseTo(200 * 0.4, 9);
      const sizes = new Set(labels.map((t) => t.fontSize));
      expect(sizes.size, "the axis reads at one size").toBe(1);
      const [size] = [...sizes];
      expect(size).toBeLessThan(TICK_FS);
      expect(size).toBeGreaterThanOrEqual(FLOOR_FS);
      for (const t of labels) {
        expect(cut(t), t.text).toBe(false);
        expect(inkSpan(t).x0, t.text).toBeGreaterThanOrEqual(-SLACK);
      }
    });

    it("holds when even the cap cannot hold a label at the floor: it is left out, never cut", () => {
      // 120pt wide: the cap is 48pt and a fourteen-digit tick is 48.6pt at the 5pt
      // floor, past the 44pt box. What does not fit is not drawn.
      const cfg = s.at(1e10, { width: 120 });
      const labels = axisLabels(cfg);
      expect(plot(cfg).x).toBeCloseTo(120 * 0.4, 9);
      expect(labels.length).toBeLessThan(s.wide.length);
      for (const t of labels) {
        expect(cut(t), t.text).toBe(false);
        expect(inkSpan(t).x0, t.text).toBeGreaterThanOrEqual(-SLACK);
        expect(t.fontSize, t.text).toBeGreaterThanOrEqual(FLOOR_FS);
      }
    });

    it("leaves the plot alone when there is no value axis to make room for (a guard: passes either way)", () => {
      const off = { decorations: { valueAxis: false } as ChartConfig["decorations"] };
      expect(plot(s.at(1e3, off)).x).toBe(plot(s.at(1, off)).x);
    });
  });
}

describe("a line's log axis is sized from the decades it prints", () => {
  const logLine = (over: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "line",
    width: 480,
    height: 300,
    data: { categories: cats, series: [{ name: "Revenue", values: [1, 40, 2_000_000] }] },
    logScale: true,
    ...over,
    decorations: { valueAxis: true, categoryAxis: true, gridlines: true },
  });

  it("widens for 10,000,000, the top decade, not for the 2,000,000 a linear axis would print", () => {
    // The data tops out at 2,000,000, which a linear axis prints as its last tick
    // (43.7pt at 9pt). A log axis rounds up to the decade — 10,000,000, 48.6pt — so a
    // column sized from the linear ticks would hold the wrong axis and shrink this
    // one: the scale the gutter is asked of must be the one that is drawn.
    const cfg = logLine();
    const labels = axisLabels(cfg);
    expect(labels.map((t) => t.text)).toEqual([
      "1",
      "10",
      "100",
      "1,000",
      "10,000",
      "100,000",
      "1,000,000",
      "10,000,000",
    ]);
    for (const t of labels) expect(t.fontSize, t.text).toBe(TICK_FS);
    // The widest label plus the box's own 4pt margin and two points of air.
    expect(plot(cfg).x).toBeCloseTo(textWidth("10,000,000", TICK_FS) + MARGIN + 2, 9);
  });
});

describe("a grouped boxplot, whose legend takes rows off the plot after the column is sized", () => {
  const groups = ["Alpha group", "Beta group", "Gamma group", "Delta group", "Epsilon group", "Zeta group"];
  const grouped = (nG: number, k: number, over: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "boxplot",
    width: 480,
    height: 300,
    data: {
      categories: ["North", "South", "East"],
      series: groups
        .slice(0, nG)
        .flatMap((g, gi) =>
          ["Min", "Q1", "Median", "Q3", "Max"].map((r, ri) =>
            row(`${r} | ${g}`, [200 + ri * 300 + gi * 20, 300 + ri * 340 + gi * 20, 100 + ri * 300 + gi * 20], k),
          ),
        ),
    },
    ...over,
    decorations: { valueAxis: true, categoryAxis: true, gridlines: true },
  });

  /** How many rows the group legend wraps onto. */
  const legendRows = (cfg: ChartConfig) =>
    new Set(named<TextNode>(buildChart(cfg).nodes, "text", /^legend-\d+$/).map((t) => Math.round(t.y))).size;

  it("prints a one-row legend's ticks whole and widens the plot for them", () => {
    const small = grouped(2, 1);
    const wide = grouped(2, 1e3);
    expect(legendRows(wide)).toBe(1);
    const labels = axisLabels(wide);
    expect(labels.length).toBe(axisLabels(small).length);
    for (const t of labels) {
      expect(t.fontSize, t.text).toBe(TICK_FS);
      expect(cut(t), t.text).toBe(false);
    }
    expect(plot(wide).x).toBeGreaterThan(plot(small).x);
  });

  it("does the same where the legend wraps and the plot is shorter than the frame it was sized on", () => {
    const wide = grouped(6, 1e3);
    expect(legendRows(wide), "the legend must wrap for this to take the banded path").toBeGreaterThan(1);
    const labels = axisLabels(wide);
    expect(labels.length).toBe(axisLabels(grouped(6, 1)).length);
    for (const t of labels) {
      expect(t.fontSize, t.text).toBe(TICK_FS);
      expect(cut(t), t.text).toBe(false);
      expect(inkSpan(t).x1, t.text).toBeLessThanOrEqual(plot(wide).x - MARGIN + SLACK);
    }
    expect(plot(wide).x).toBeGreaterThan(plot(grouped(6, 1)).x);
  });

  it("never sizes the column for less than the axis prints, when the banded plot is too short to draw it at full size", () => {
    // 240x100: the three legend rows leave a plot whose ticks are 10.5pt apart
    // where a label of the full size wants 14, so the axis draws at the SPACING's
    // size (three quarters of 9, 6.75pt), not at 9. The column
    // is asked of the frame before the band comes off, so it errs wide — it may
    // pay for a bigger size than is drawn, never a smaller one. What it must not do
    // is shrink the labels again for width: the wide chart reads at the same size
    // as the one whose numbers are small.
    const small = axisLabels(grouped(6, 1, { width: 240, height: 100 }));
    const wideCfg = grouped(6, 1e3, { width: 240, height: 100 });
    const wide = axisLabels(wideCfg);
    expect(legendRows(wideCfg)).toBeGreaterThan(1);
    expect(small.length).toBeGreaterThan(0);
    expect(small[0].fontSize, "the plot is too short for the full size either way").toBeLessThan(TICK_FS);
    expect(wide.length).toBe(small.length);
    for (const t of wide) {
      expect(t.fontSize, t.text).toBeCloseTo(small[0].fontSize, 9);
      expect(cut(t), t.text).toBe(false);
      expect(inkSpan(t).x0, t.text).toBeGreaterThanOrEqual(-SLACK);
      expect(inkSpan(t).x1, t.text).toBeLessThanOrEqual(plot(wideCfg).x - MARGIN + SLACK);
    }
  });
});

describe("a sideways chart keeps its frame: its value axis runs along the bottom (a guard: passes either way)", () => {
  const sideways = (kind: "line" | "boxplot", k: number): ChartConfig => ({
    kind,
    width: 480,
    height: 300,
    horizontal: true,
    data: {
      categories: ["A", "B", "C"],
      series:
        kind === "boxplot"
          ? [
              row("Min", [200, 300, 100], k),
              row("Q1", [400, 500, 300], k),
              row("Median", [600, 800, 500], k),
              row("Q3", [1000, 1200, 800], k),
              row("Max", [1600, 2000, 1400], k),
            ]
          : [row("Revenue", [1200, 2000, 1600], k)],
    },
    decorations: { valueAxis: true, categoryAxis: true },
  });

  for (const kind of ["line", "boxplot"] as const) {
    it(`${kind}: the category column does not depend on how big the numbers are`, () => {
      const at = (k: number) => named<TextNode>(buildChart(sideways(kind, k)).nodes, "text", "category-0")[0];
      expect(at(1e3).x).toBe(at(1).x);
      expect(at(1e3).w).toBe(at(1).w);
    });
  }
});
