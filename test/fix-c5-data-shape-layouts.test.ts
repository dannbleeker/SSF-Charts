import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart, valueExtent } from "../src/core/chart";
import { detectLayout } from "../src/core/layout/tilemap-layouts";
import type { RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * Layouts that mis-detect their data: the tile map's auto-picked grid, the
 * boxplot's summary-vs-samples switch, the calendar heatmap's colour scale and
 * the tile map's glyph legend. Each describe block is one finding.
 */

const named = <T extends SceneNode>(nodes: SceneNode[], name: string): T | undefined =>
  nodes.find((n) => n.name === name) as T | undefined;

// ---------------------------------------------------------------------------
// r1-10 — the tile map's auto-detect
// ---------------------------------------------------------------------------

describe("tile map auto-detect keeps the codes it was given", () => {
  // 11 codes, ten of them EU members and GB not. EU_TILES is a strict subset of
  // EUROPE_TILES, so the Europe grid matches all eleven and the EU grid ten.
  const mixed = ["DK", "SE", "FI", "DE", "NL", "FR", "ES", "IT", "PL", "AT", "GB"];
  const config = (categories: string[], map?: ChartConfig["map"]): ChartConfig => ({
    kind: "tilemap",
    ...DEFAULT_SIZE,
    ...(map ? { map } : {}),
    data: { categories, series: [{ name: "Revenue", values: categories.map((_, i) => (i + 1) * 10) }] },
  });

  it("picks Europe, not the EU grid, when a few non-EU members ride along", () => {
    expect(detectLayout(mixed)).toBe("europe");
    // 9 EU + 1 non-EU is the smallest set the 0.9 bar used to send to the EU grid.
    expect(detectLayout(["DK", "SE", "FI", "NL", "FR", "ES", "IT", "PL", "AT", "GB"])).toBe("europe");
    // ...and two non-EU codes among eighteen.
    const eighteen = ["DK", "SE", "FI", "NL", "FR", "ES", "IT", "PL", "AT", "PT", "BE", "IE", "BG", "CZ", "EE", "GR"];
    expect(detectLayout([...eighteen, "LV", "LT", "GB", "NO"])).toBe("europe");
  });

  it("draws every code it was given: the tile, its value and its place on the scale", () => {
    const auto = buildChart(config(mixed));
    const gb = named<RectNode>(auto.nodes, "tile-GB");
    expect(gb, "GB lost its tile").toBeDefined();
    // The last value (110) is GB's, and it sets the top of the legend.
    expect(named<TextNode>(auto.nodes, "legend-max")?.text).toBe("110");
    // Auto-detect is the same chart as asking for the layout it should have picked.
    expect(auto.nodes).toEqual(buildChart(config(mixed, "europe")).nodes);
  });

  it("still takes the tighter EU grid for a set that is EU-only", () => {
    const euOnly = ["DK", "SE", "FI", "NL", "FR", "ES", "IT", "PL", "AT", "PT", "BE", "IE"];
    expect(detectLayout(euOnly)).toBe("eu");
    // An unrecognised code is in neither grid, so it cannot argue for the wider one.
    expect(detectLayout([...euOnly.slice(0, 9), "XX"])).toBe("eu");
  });
});

// ---------------------------------------------------------------------------
// r1-13 — the boxplot's summary-vs-samples switch
// ---------------------------------------------------------------------------

describe("boxplot reads a row named like a summary row as an observation unless the whole summary is there", () => {
  const rows = (names: string[], categories = ["Team A", "Team B"]): ChartConfig => ({
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: {
      categories,
      series: names.map((name, i) => ({ name, values: [3 + i, 5 + 2 * i] })),
    },
  });
  const boxes = (cfg: ChartConfig) => buildChart(cfg).nodes.filter((n) => n.name?.match(/^box-\d+$/));

  it("draws a box per category when one raw-sample row happens to be called Max", () => {
    // Anna, Ben, Max, Zoe, Ola: five people, five observations per team.
    const s = buildChart(rows(["Anna", "Ben", "Max", "Zoe", "Ola"]));
    expect(s.nodes.filter((n) => n.name?.match(/^box-\d+$/))).toHaveLength(2);
    // Renaming the one row changes nothing about the chart.
    expect(s.nodes).toEqual(buildChart(rows(["Anna", "Ben", "Maxi", "Zoe", "Ola"])).nodes);
  });

  it("matches the magic names the way the summary does (case and padding ignored)", () => {
    for (const name of ["Min", "max", "  Max  ", "MEDIAN", "Q1"])
      expect(boxes(rows(["Anna", name, "Bo", "Zoe", "Ola"])), name).toHaveLength(2);
  });

  it("does not take four quarters of observations for a summary", () => {
    // Q1 and Q3 are two of the five names; Min, Median and Max are not there.
    expect(boxes(rows(["Q1", "Q2", "Q3", "Q4"]))).toHaveLength(2);
  });

  it("reports the extent of the boxes it draws", () => {
    expect(valueExtent(rows(["Anna", "Ben", "Max", "Zoe", "Ola"]))).not.toBeNull();
  });

  it("decides per group: a summary group beside a raw-sample group draws both", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: {
        categories: ["North", "South"],
        series: [
          { name: "Min | 2024", values: [2, 3] },
          { name: "Q1 | 2024", values: [3, 5] },
          { name: "Median | 2024", values: [4, 7] },
          { name: "Q3 | 2024", values: [6, 9] },
          { name: "Max | 2024", values: [8, 12] },
          ...[4, 5, 6, 7, 8].map((v, i) => ({ name: `obs${i} | 2025`, values: [v, v + 3] })),
        ],
      },
    };
    const names = buildChart(cfg).nodes.map((n) => n.name);
    for (const box of ["box-0-g0", "box-1-g0", "box-0-g1", "box-1-g1"]) expect(names, box).toContain(box);
  });

  it("still reads the full five-number set as a summary", () => {
    const summary: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "Min", values: [2] },
          { name: "Q1", values: [3] },
          { name: "Median", values: [4] },
          { name: "Q3", values: [6] },
          { name: "Max", values: [30] },
        ],
      },
    };
    // Read as five observations, 30 would sit past the Tukey fence and be drawn
    // as an outlier dot. As a summary it is the top of the whisker.
    const nodes = buildChart(summary).nodes;
    expect(boxes(summary)).toHaveLength(1);
    expect(nodes.some((n) => n.name?.startsWith("outlier-"))).toBe(false);
    expect(named<{ y1: number } & SceneNode>(nodes, "cap-hi-0")).toBeDefined();
    expect(valueExtent(summary)?.max).toBe(30);
  });
});

// ---------------------------------------------------------------------------
// r1-14 — the calendar heatmap's scale
// ---------------------------------------------------------------------------

describe("calendar heatmap scales to the one series it draws", () => {
  const days = Array.from({ length: 14 }, (_, i) => new Date(Date.UTC(2025, 0, 6 + i)).toISOString().slice(0, 10));
  const visits = { name: "Visits", values: days.map((_, i) => 10 + i) };
  const calendar = (series: ChartConfig["data"]["series"], heatmap: ChartConfig["heatmap"] = {}): ChartConfig => ({
    kind: "heatmap",
    ...DEFAULT_SIZE,
    heatmap: { calendar: true, ...heatmap },
    data: { categories: days, series },
  });

  it("is not washed out by a second row on a different scale", () => {
    const one = buildChart(calendar([visits]));
    const two = buildChart(calendar([visits, { name: "Revenue", values: days.map((_, i) => 10000 + 1000 * i) }]));
    const fill = (nodes: SceneNode[], name: string) => named<RectNode>(nodes, name)?.fill;
    // The ramp runs across the 14 days it draws: the first is the palest step, the last the deepest.
    expect(fill(one.nodes, "cell-0")).not.toBe(fill(one.nodes, "cell-13"));
    expect(fill(two.nodes, "cell-13")).toBe(fill(one.nodes, "cell-13"));
    expect(two.nodes).toEqual(one.nodes);
  });

  it("does not turn diverging because another row has a negative in it", () => {
    const one = buildChart(calendar([visits]));
    const two = buildChart(calendar([visits, { name: "Delta", values: days.map((_, i) => (i === 3 ? -50 : 5)) }]));
    // The legend's first step would be the negative colour if the sheet-wide mode leaked in.
    expect(named<RectNode>(two.nodes, "legend-step-0")?.fill).toBe(named<RectNode>(one.nodes, "legend-step-0")?.fill);
    expect(two.nodes).toEqual(one.nodes);
  });

  it("marks no cell as positive on behalf of a row that is not drawn", () => {
    const delta = { name: "Delta", values: days.map((_, i) => (i === 3 ? -50 : 5)) };
    const signs = (series: ChartConfig["data"]["series"]) =>
      buildChart(calendar(series, { symbols: "sign" })).nodes.filter((n) => n.name?.startsWith("cell-sign-"));
    // All-positive data: the option stays silent. A negative elsewhere must not break that.
    expect(signs([visits])).toHaveLength(0);
    expect(signs([visits, delta])).toHaveLength(0);
  });

  it("still goes diverging, and still marks, when the DRAWN series spans zero", () => {
    const net = { name: "Net", values: days.map((_, i) => (i % 2 ? 1 : -1) * (i + 1)) };
    const cfg = calendar([net, visits], { symbols: "sign" });
    expect(buildChart(cfg).nodes.filter((n) => n.name?.startsWith("cell-sign-")).length).toBeGreaterThan(5);
    expect(buildChart(cfg).nodes).toEqual(buildChart(calendar([net], { symbols: "sign" })).nodes);
  });

  it("leaves the matrix heatmap on one scale across every row", () => {
    // The shared scale is the point of the matrix; only the calendar draws one row.
    const matrix = (second: number): ChartConfig => ({
      kind: "heatmap",
      ...DEFAULT_SIZE,
      data: {
        categories: ["a", "b"],
        series: [
          { name: "r0", values: [1, 2] },
          { name: "r1", values: [second, second] },
        ],
      },
    });
    // The value 2 is the top of the scale when the other row tops out at 2, and
    // near the bottom of it when the other row reaches 1000.
    const cell = (c: ChartConfig) => named<RectNode>(buildChart(c).nodes, "cell-0-1")?.fill;
    expect(cell(matrix(1000))).not.toBe(cell(matrix(2)));
  });
});

// ---------------------------------------------------------------------------
// r2-13 — the tile map's glyph legend wraps
// ---------------------------------------------------------------------------

describe("tile map glyph legend stays on the chart", () => {
  const countries = ["DK", "SE", "NO", "FI", "DE", "NL", "FR", "GB", "ES", "IT", "PL"];
  const glyph = (names: string[], size = { ...DEFAULT_SIZE } as { width: number; height: number }): ChartConfig => ({
    kind: "tilemap",
    ...size,
    map: "europe",
    tilemap: { glyph: "bars" },
    data: {
      categories: countries,
      series: names.map((name, si) => ({ name, values: countries.map((_, i) => 10 + ((i * 7 + si * 3) % 40)) })),
    },
  });
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const longNames = [
    "Revenue change 2022",
    "Revenue change 2023",
    "Margin change 2024 -> 2025",
    "Margin change 2025 -> 2026",
    "Volume change 2025",
    "Volume change 2026",
  ];
  const chips = (nodes: SceneNode[]) => nodes.filter((n): n is RectNode => !!n.name?.startsWith("legend-chip-"));

  it("keeps every chip on a chart of the default size and names every series", () => {
    for (const names of [months, longNames]) {
      const s = buildChart(glyph(names));
      const cs = chips(s.nodes);
      expect(cs).toHaveLength(names.length);
      for (const c of cs) {
        expect(c.x, c.name).toBeGreaterThanOrEqual(0);
        expect(c.x + c.w, c.name).toBeLessThanOrEqual(s.width);
        expect(c.y + c.h, c.name).toBeLessThanOrEqual(s.height);
      }
      // A name is drawn whole or the chart is wrong: not clipped, not dropped.
      names.forEach((name, i) => expect(named<TextNode>(s.nodes, `legend-${i}`)?.text, name).toBe(name));
    }
  });

  it("wraps onto a second row, and no chip is over a tile", () => {
    const s = buildChart(glyph(longNames));
    const cs = chips(s.nodes);
    expect(new Set(cs.map((c) => c.y)).size, "rows of chips").toBeGreaterThan(1);
    const tiles = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.match(/^tile-[A-Z]+$/));
    const gridBottom = Math.max(...tiles.map((t) => t.y + t.h));
    // The legend hangs from the grid's foot: no chip is over a tile.
    for (const c of cs) expect(c.y, c.name).toBeGreaterThanOrEqual(gridBottom - 0.01);
  });

  it("leaves a legend that fits on one row exactly where it was: hung from the grid's left edge", () => {
    const s = buildChart(glyph(["2024", "2025", "2026"]));
    const cs = chips(s.nodes);
    expect(new Set(cs.map((c) => c.y)).size).toBe(1);
    const left = Math.min(...s.nodes.filter((n) => n.name?.match(/^tile-[A-Z]+$/)).map((n) => (n as RectNode).x));
    expect(cs[0].x).toBeCloseTo(left, 6);
  });

  it("never costs the grid its bars: where the wrapped key cannot be paid for, the key is what goes", () => {
    // Letters name six series on one row at any width, so they pay for no extra
    // legend row and show what the grid draws when the key costs it nothing.
    const letters = longNames.map((_, i) => String.fromCharCode(65 + i));
    const drawn = (names: string[], height: number) => {
      const nodes = buildChart(glyph(names, { width: 480, height })).nodes;
      return {
        bars: nodes.filter((n) => n.name?.startsWith("glyph-")).length,
        chips: chips(nodes).length,
      };
    };
    const scan = [];
    for (let height = 160; height <= 420; height += 20) {
      const long = drawn(longNames, height);
      expect(long.bars, `480x${height}`).toBe(drawn(letters, height).bars);
      scan.push(long);
    }
    // The scan crosses both sides of the line, so it cannot pass by never reaching it.
    expect(
      scan.some((r) => r.chips > 0 && r.bars > 0),
      "a frame that pays for the wrapped key",
    ).toBe(true);
    expect(
      scan.some((r) => r.chips === 0 && r.bars > 0),
      "a frame that keeps its bars and drops the key",
    ).toBe(true);
  });

  it("never draws a chip outside a small or cramped frame: it fits or it is dropped", () => {
    const frames: [number, number][] = [
      [80, 60],
      [120, 90],
      [160, 120],
      [200, 150],
      [300, 60],
      [60, 300],
    ];
    const bad: string[] = [];
    for (const [width, height] of frames)
      for (const fontSize of [10, 18])
        for (const names of [months.slice(0, 3), longNames, months]) {
          const cfg = { ...glyph(names, { width, height }), style: { fontSize } } as ChartConfig;
          const s = buildChart(cfg);
          for (const c of chips(s.nodes))
            if (c.x < -0.01 || c.y < -0.01 || c.x + c.w > width + 0.01 || c.y + c.h > height + 0.01)
              bad.push(
                `${width}x${height}@${fontSize} ${names.length} series: ${c.name} at ${c.x.toFixed(1)},${c.y.toFixed(1)}`,
              );
        }
    expect(bad).toEqual([]);
  });
});
