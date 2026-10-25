import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart, normalizeConfig } from "../src/core/chart";
import { textWidth } from "../src/core/scene";
import type { LineNode, RectNode, SceneNode, TextNode } from "../src/core/scene";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { sceneToSvg } from "../src/render/svg";
import type { ChartConfig, ChartKind, Series } from "../src/core/types";

/**
 * Stream k2 of the third bug hunt: blank categories in the layouts the first two
 * sweeps never reached, wrong-typed option lists, and degenerate shapes.
 *
 * Each describe block names the finding it closes and states the PROPERTY, not
 * today's coordinates — the oracle is what the chart would say if the thing in
 * question were not there (a blank category is a category that was not measured,
 * so the chart is the one without it).
 */

const named = <T extends SceneNode>(nodes: SceneNode[], kind: T["kind"], name: string | RegExp): T[] =>
  nodes.filter(
    (n): n is T => n.kind === kind && (typeof name === "string" ? n.name === name : !!n.name && name.test(n.name)),
  );

const textOf = (cfg: ChartConfig, name: string): string | undefined =>
  named<TextNode>(buildChart(cfg).nodes, "text", name)[0]?.text;

const drawn = (cfg: ChartConfig, name: string): boolean => buildChart(cfg).nodes.some((n) => n.name === name);

/** A config with category `index` dropped from every series — the chart that never had it. */
function withoutCategory(cfg: ChartConfig, index: number): ChartConfig {
  const keep = <T>(xs: T[]) => xs.filter((_, i) => i !== index);
  return {
    ...cfg,
    data: {
      ...cfg.data,
      categories: keep(cfg.data.categories),
      series: cfg.data.series.map((s) => ({ ...s, values: keep(s.values) })),
      ...(cfg.data.xExtent ? { xExtent: keep(cfg.data.xExtent) } : {}),
    },
  };
}

const series = (name: string, values: (number | null)[]): Series => ({ name, values });

/**
 * The four kinds the blank-category fix (beb9c80) never reached: each builds
 * `columnValue` with `?? 0` and published no `columnHasData`, so the mean averaged
 * the blank in and the arrows were drawn to it and printed "-100%".
 *
 * `blank` is the category that carries nothing measured; the other four are
 * ordinary.
 */
const BLANK_KINDS: { kind: ChartKind; label: string; cfg: ChartConfig; blank: number }[] = [
  {
    kind: "candlestick",
    label: "candlestick (blank Close)",
    blank: 2,
    cfg: {
      kind: "candlestick",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        series: [
          series("Open", [42, 44, 43, 46, 45]),
          series("High", [45, 46, 47, 47, 48]),
          series("Low", [41, 42, 42, 44, 44]),
          series("Close", [44, 43, null, 45, 47]),
        ],
      },
    },
  },
  {
    kind: "boxplot",
    label: "boxplot (summary rows)",
    blank: 2,
    cfg: {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: {
        categories: ["North", "South", "East", "West"],
        series: [
          series("Min", [2, 3, null, 2]),
          series("Q1", [3, 5, null, 4]),
          series("Median", [4, 7, null, 5]),
          series("Q3", [6, 9, null, 7]),
          series("Max", [8, 12, null, 10]),
        ],
      },
    },
  },
  {
    kind: "boxplot",
    label: "boxplot (raw samples)",
    blank: 1,
    cfg: {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          series("s1", [70, null, 80, 90]),
          series("s2", [80, null, 90, 100]),
          series("s3", [90, null, 100, 110]),
        ],
      },
    },
  },
  {
    kind: "violin",
    label: "violin",
    blank: 1,
    cfg: {
      kind: "violin",
      ...DEFAULT_SIZE,
      data: {
        categories: ["API", "Web", "Batch"],
        series: [
          series("s1", [120, null, 640]),
          series("s2", [135, null, 700]),
          series("s3", [128, null, 580]),
          series("s4", [142, null, 660]),
        ],
      },
    },
  },
  {
    kind: "mekko",
    label: "mekko",
    blank: 1,
    cfg: {
      kind: "mekko",
      ...DEFAULT_SIZE,
      data: {
        categories: ["EMEA", "Americas", "APAC"],
        series: [series("Enterprise", [42, null, 18]), series("SMB", [28, null, 22])],
      },
    },
  },
];

describe("r3-6 / r3-17: candlestick, boxplot, violin and mekko publish which categories were measured", () => {
  const withDecor = (cfg: ChartConfig, last: number): ChartConfig => ({
    ...cfg,
    decorations: {
      ...cfg.decorations,
      valueLines: [{ mode: "mean" }],
      difference: { from: 0, to: last },
      cagr: { from: 0, to: last },
    },
  });

  it.each(BLANK_KINDS)("$label: an arrow to the blank category is not drawn", ({ cfg, blank }) => {
    const c = withDecor(cfg, blank);
    for (const input of [c, normalizeConfig(c)]) {
      expect(drawn(input, "diff-line")).toBe(false);
      expect(drawn(input, "diff-label")).toBe(false);
      expect(drawn(input, "cagr-line")).toBe(false);
      expect(drawn(input, "cagr-label")).toBe(false);
    }
  });

  it.each(BLANK_KINDS.filter((k) => k.kind !== "mekko"))(
    "$label: the mean is of what was measured, the chart that never had the blank",
    ({ cfg, blank }) => {
      const mean = (c: ChartConfig) =>
        textOf({ ...c, decorations: { valueLines: [{ mode: "mean" }] } }, "value-line-label-0");
      expect(mean(cfg)).toMatch(/^Ø /);
      expect(mean(cfg)).toBe(mean(withoutCategory(cfg, blank)));
    },
  );

  it.each(BLANK_KINDS)(
    "$label: a measured end still draws the arrow (the gate is the blank, not the kind)",
    ({ cfg, blank }) => {
      // Between two categories that were measured the arrow is drawn, as it always was.
      const last = cfg.data.categories.length - 1;
      const c = { ...cfg, decorations: { difference: { from: 0, to: last } } };
      expect(blank).not.toBe(0);
      expect(blank).not.toBe(last);
      expect(drawn(c, "diff-line")).toBe(true);
    },
  );
});

describe("r3-43: a Pareto ranks the measure, not a Target or Error row typed above it", () => {
  const texts = (cfg: ChartConfig) => named<TextNode>(buildChart(cfg).nodes, "text", /./).map((t) => t.text);
  const pareto = (rows: Series[]): ChartConfig => ({
    kind: "clustered",
    ...DEFAULT_SIZE,
    pareto: true,
    data: { categories: ["Scratch", "Dent", "Paint", "Other"], series: rows },
  });
  const defects = series("Defects", [5, 40, 12, 3]);

  it.each(["Target", "Error", "Error+", "Error −", "Band Low"])(
    "a %s row typed first changes nothing: the chart is the one without it",
    (name) => {
      const special = series(name, [10, 10, 10, 10]);
      const clean = texts(pareto([defects]));
      expect(texts(pareto([special, defects]))).toEqual(clean);
      // And it is the measure that was ranked: Dent (40) leads, Other (3) is last.
      expect(clean.slice(0, 4)).toEqual(["Dent", "Paint", "Scratch", "Other"]);
    },
  );

  it("with nothing but special rows there is no bar to rank and the config is left alone", () => {
    const cfg = pareto([series("Target", [10, 10, 10, 10])]);
    expect(texts(cfg)).not.toContain("Cumulative %");
  });
});

/** The whole scene as a string: two charts are the same chart when this is equal. */
const scene = (cfg: ChartConfig): string => JSON.stringify(buildChart(cfg));
/** Config-shaped junk, typed away: the whole point is a value the types forbid. */
const junk = (cfg: object): ChartConfig => cfg as unknown as ChartConfig;

describe("r3-41: multiples.columns that is no number is read as not given", () => {
  const grid = (columns: unknown): ChartConfig =>
    junk({
      kind: "stacked",
      ...DEFAULT_SIZE,
      width: 600,
      height: 300,
      multiples: columns === undefined ? {} : { columns },
      data: {
        categories: ["Q1", "Q2", "Q3", "Q4"],
        series: [series("A", [1, 2, 3, 4]), series("B", [2, 3, 4, 5]), series("C", [3, 4, 5, 6])],
      },
    });
  const panels = (cfg: ChartConfig) =>
    named<TextNode>(buildChart(cfg).nodes, "text", /./).filter((t) => t.text === "Q1").length;

  it.each(["auto", "50%", "abc", {}])("columns %j still draws the grid it would with no columns", (bad) => {
    expect(panels(grid(undefined))).toBe(3);
    expect(panels(grid(bad))).toBe(3);
    expect(scene(grid(bad))).toBe(scene(grid(undefined)));
  });

  it("a number, or a numeric string, still says how many columns", () => {
    expect(scene(grid("2"))).toBe(scene(grid(2)));
    expect(scene(grid(2))).not.toBe(scene(grid(undefined)));
  });
});

describe("r3-50 / r3-71: a list option written as the wrong type is repaired, not thrown", () => {
  const waterfall = (waterfallOpts: object): ChartConfig => {
    const base = sampleConfig("waterfall");
    return junk({ ...base, waterfall: waterfallOpts });
  };
  const pie = (kind: "pie" | "doughnut", pieOpts: object): ChartConfig => junk({ ...sampleConfig(kind), pie: pieOpts });

  it("a bare index is a list of one: totalIndices, spacerIndices", () => {
    expect(scene(waterfall({ totalIndices: 5 }))).toBe(scene(waterfall({ totalIndices: [5] })));
    expect(scene(waterfall({ totalIndices: [5], spacerIndices: 2 }))).toBe(
      scene(waterfall({ totalIndices: [5], spacerIndices: [2] })),
    );
    // And it is a real total bar, not a quietly ignored option.
    expect(scene(waterfall({ totalIndices: 5 }))).not.toBe(scene(waterfall({})));
  });

  it("an option with no sensible reading is dropped, as if it were not written", () => {
    for (const bad of [true, {}, "5", null, Number.NaN]) {
      expect(scene(waterfall({ totalIndices: bad, spacerIndices: bad, detailGroups: bad }))).toBe(scene(waterfall({})));
    }
  });

  it("a list with a non-number in it keeps the numbers", () => {
    expect(scene(waterfall({ totalIndices: [5, "x", null, Number.NaN] }))).toBe(
      scene(waterfall({ totalIndices: [5] })),
    );
  });

  it("detailGroups: a bare group is dropped, a group without an index list is dropped, a bare index is a list of one", () => {
    const withGroups = (detailGroups: unknown) =>
      junk({
        kind: "waterfall",
        ...DEFAULT_SIZE,
        data: {
          categories: ["Start", "Total", "of A", "of B", "End"],
          series: [series("Delta", [100, 50, 20, 30, 0])],
        },
        waterfall: { totalIndices: [4], detailGroups },
      });
    const clean = scene(withGroups(undefined));
    expect(scene(withGroups(3))).toBe(clean);
    expect(scene(withGroups([{ of: 1 }]))).toBe(clean);
    expect(scene(withGroups([null, "x", { indices: [2] }]))).toBe(clean);
    expect(scene(withGroups([{ of: 1, indices: 2 }]))).toBe(scene(withGroups([{ of: 1, indices: [2] }])));
    expect(scene(withGroups({ of: 1, indices: [2, 3] }))).toBe(scene(withGroups([{ of: 1, indices: [2, 3] }])));
    expect(scene(withGroups([{ of: 1, indices: [2, 3] }]))).not.toBe(clean);
  });

  it.each(["pie", "doughnut"] as const)("%s: explode and breakout take a bare index as a list of one", (kind) => {
    expect(scene(pie(kind, { explode: 1 }))).toBe(scene(pie(kind, { explode: [1] })));
    expect(scene(pie(kind, { explode: [1] }))).not.toBe(scene(pie(kind, {})));
    expect(scene(pie(kind, { explode: true, breakout: {} }))).toBe(scene(pie(kind, {})));
  });

  it("pie.breakout: a bare index is a list of one", () => {
    expect(scene(pie("pie", { breakout: 2 }))).toBe(scene(pie("pie", { breakout: [2] })));
    expect(scene(pie("pie", { breakout: [2] }))).not.toBe(scene(pie("pie", {})));
    expect(scene(pie("pie", { breakout: "x" }))).toBe(scene(pie("pie", {})));
  });

  it("a combo reads the same waterfall lists", () => {
    const combo = (totalIndices: unknown): ChartConfig =>
      junk({
        kind: "combo",
        ...DEFAULT_SIZE,
        combo: { columns: "waterfall" },
        data: {
          categories: ["A", "B", "C", "D"],
          series: [series("Delta", [10, 5, -3, 0]), { name: "Trend", type: "line", values: [10, 15, 12, 12] }],
        },
        waterfall: { totalIndices },
      });
    expect(scene(combo(3))).toBe(scene(combo([3])));
  });

  it("scatter.spread is an axis or nothing", () => {
    const scatter = (spread: unknown): ChartConfig =>
      junk({ ...sampleConfig("scatter"), scatter: spread === undefined ? {} : { spread } });
    const none = scene(scatter(undefined));
    for (const bad of [true, 1, {}, [], { axis: "x" }, "z", ""])
      expect(scene(scatter(bad)), JSON.stringify(bad)).toBe(none);
    expect(scene(scatter("x"))).not.toBe(none);
  });

  it("style.palette: null is a palette that was not given — on every kind", () => {
    for (const { kind } of CHART_KINDS) {
      const base = sampleConfig(kind);
      expect(scene(junk({ ...base, style: { palette: null } })), kind).toBe(scene({ ...base, style: {} }));
    }
  });
});

describe("r3-58: a category index that is no whole number never reaches a coordinate", () => {
  const withDecor = (kind: ChartKind, decorations: object): ChartConfig =>
    junk({ ...sampleConfig(kind), decorations: { ...sampleConfig(kind).decorations, ...decorations } });
  const NUMERIC = ["x", "y", "x1", "y1", "x2", "y2", "w", "h"];
  /** Every geometry field a node HAS is a finite number: an `undefined` own key is the hole `finiteNodes` cannot see. */
  const holes = (cfg: ChartConfig): string[] =>
    buildChart(cfg).nodes.flatMap((n) =>
      NUMERIC.filter((k) => k in n && typeof (n as unknown as Record<string, unknown>)[k] !== "number").map(
        (k) => `${n.name}:${k}`,
      ),
    );

  it.each(["stacked100", "mekko", "treemap", "sunburst"] as const)(
    "%s: a fractional or non-numeric arrow end leaves no hole",
    (kind) => {
      for (const end of [0.5, 1.5, "abc", {}, Number.NaN]) {
        const cfg = withDecor(kind, { difference: { from: end, to: 2 }, cagr: { from: 0, to: end } });
        expect(holes(cfg), `${kind} ${JSON.stringify(end)}`).toEqual([]);
        expect(sceneToSvg(buildChart(cfg)), `${kind} ${JSON.stringify(end)}`).not.toContain("NaN");
      }
    },
  );

  it("a fractional end is the category it rounds to", () => {
    const at = (from: number) => scene(withDecor("stacked100", { difference: { from, to: 2 } }));
    expect(at(0.6)).toBe(at(1));
    expect(at(0.4)).toBe(at(0));
    // ...and the arrow is there, so this is not two empty charts agreeing.
    expect(drawn(withDecor("stacked100", { difference: { from: 0.6, to: 2 } }), "diff-line")).toBe(true);
  });

  it("an end that is no number draws no arrow", () => {
    // Not compared with a chart that never asked for one: the request reserves its
    // margin whether or not an arrow can be drawn.
    for (const end of ["abc", {}, Number.NaN]) {
      const diff = withDecor("clustered", { difference: { from: 0, to: end } });
      const rate = withDecor("clustered", { cagr: { from: end, to: 2 } });
      for (const name of ["diff-line", "diff-head", "diff-label", "diff-level"])
        expect(drawn(diff, name), name).toBe(false);
      for (const name of ["cagr-line", "cagr-head", "cagr-label"]) expect(drawn(rate, name), name).toBe(false);
    }
  });

  it("a callout and a band are placed on the category they round to", () => {
    expect(scene(withDecor("clustered", { callouts: [{ text: "hi", category: 1.6 }] }))).toBe(
      scene(withDecor("clustered", { callouts: [{ text: "hi", category: 2 }] })),
    );
    expect(scene(withDecor("clustered", { bands: [{ axis: "x", from: 0.4, to: 1.6 }] }))).toBe(
      scene(withDecor("clustered", { bands: [{ axis: "x", from: 0, to: 2 }] })),
    );
    // No number, no place to put it.
    expect(drawn(withDecor("clustered", { callouts: [{ text: "hi", category: "abc" }] }), "callout-box-0")).toBe(false);
    expect(drawn(withDecor("clustered", { bands: [{ axis: "x", from: "abc", to: 1 }] }), "band-0")).toBe(false);
    expect(drawn(withDecor("clustered", { callouts: [{ text: "hi", category: 1.6 }] }), "callout-box-0")).toBe(true);
  });
});

describe("r3-10: a calendar heatmap's size does not depend on how far a mistyped date is from the rest", () => {
  const days = (n: number, start = "2026-01-"): string[] =>
    Array.from({ length: n }, (_, i) => `${start}${String(i + 1).padStart(2, "0")}`);
  const calendar = (categories: string[], calendarOn = true): ChartConfig =>
    junk({
      kind: "heatmap",
      ...DEFAULT_SIZE,
      heatmap: { calendar: calendarOn },
      data: {
        categories,
        series: [
          series(
            "Activity",
            categories.map((_, i) => (i * 7) % 11),
          ),
        ],
      },
    });
  /** A month of January 2026 with one date replaced. */
  const typo = (date: string): ChartConfig => calendar([...days(30), date]);
  const count = (cfg: ChartConfig) => buildChart(cfg).nodes.length;

  it.each(["20260-01-15", "99999-01-15", "9999-01-15", "3026-01-15"])(
    "a stray %s does not decide the node count",
    (date) => {
      expect(count(typo(date))).toBeLessThan(500);
      // Past what a calendar can span it is read as the ordinary grid, which is
      // bounded by the categories it was given.
      expect(scene(typo(date))).toBe(scene(calendar([...days(30), date], false)));
    },
  );

  it("the extremes the Date object can hold build in a blink and a few nodes", () => {
    const t0 = performance.now();
    const n = count(calendar(["-271821-04-20", "+275760-09-13"]));
    expect(n).toBeLessThan(500);
    expect(performance.now() - t0).toBeLessThan(2000);
  });

  it("a real multi-year daily calendar is still a calendar", () => {
    // CLAUDE.md names ~1-4k days as a real shape; four years of days is 1461.
    const cats = Array.from({ length: 1461 }, (_, i) =>
      new Date(Date.UTC(2022, 0, 1) + i * 86400000).toISOString().slice(0, 10),
    );
    const nodes = buildChart(calendar(cats)).nodes;
    expect(nodes.some((x) => x.name === "weekday-0")).toBe(true);
    expect(nodes.filter((x) => x.name?.startsWith("cell-") && !x.name.startsWith("cell-sign"))).toHaveLength(1461);
  });
});

describe("r3-59: a dense gantt carries no text below the floor — owner, remark and gutter cells go with the name", () => {
  /** The floor every fitted label in the engine keeps (MIN_LABEL_FS in frame.ts). */
  const FLOOR = 5;
  const gantt = (rows: number, size = { width: 480, height: 300 }): ChartConfig =>
    junk({
      kind: "gantt",
      ...DEFAULT_SIZE,
      ...size,
      data: {
        categories: Array.from({ length: rows }, (_, i) => `Task ${i} | Ann | note ${i}`),
        series: [
          series(
            "Start",
            Array.from({ length: rows }, (_, i) => i * 2),
          ),
          series(
            "End",
            Array.from({ length: rows }, (_, i) => i * 2 + 5),
          ),
          series(
            "Column Cost",
            Array.from({ length: rows }, (_, i) => 10 + i),
          ),
        ],
      },
    });
  const fonts = (cfg: ChartConfig) =>
    named<TextNode>(buildChart(cfg).nodes, "text", /./)
      .filter((t) => t.text.trim() !== "")
      .map((t) => ({ name: t.name ?? "", fontSize: t.fontSize }));

  it.each([20, 30, 35, 38, 40, 60])("%i rows: no text node is drawn below the floor", (rows) => {
    const small = fonts(gantt(rows)).filter((t) => t.fontSize < FLOOR);
    expect(small).toEqual([]);
  });

  it("40 rows: the names are gone and so are the columns that sat beside them", () => {
    const names = fonts(gantt(40)).map((t) => t.name);
    expect(names.some((n) => n.startsWith("category-"))).toBe(false);
    for (const prefix of ["owner-", "remark-", "col-0-"])
      expect(
        names.some((n) => n.startsWith(prefix)),
        prefix,
      ).toBe(false);
  });

  it("20 rows: everything that fits is still drawn", () => {
    const names = fonts(gantt(20)).map((t) => t.name);
    for (const prefix of ["category-", "owner-", "remark-", "col-0-"])
      expect(
        names.some((n) => n.startsWith(prefix)),
        prefix,
      ).toBe(true);
  });
});

describe("r3-60: a hierarchical treemap draws a group's header and tiles only inside the group's cell", () => {
  const SLACK = 0.5;
  const treemap = (rows: [string, number][], width = 480, height = 300): ChartConfig =>
    junk({
      kind: "treemap",
      ...DEFAULT_SIZE,
      width,
      height,
      data: {
        categories: rows.map(([c]) => c),
        series: [
          series(
            "Revenue",
            rows.map(([, v]) => v),
          ),
        ],
      },
    });
  /** What a node puts on the canvas: its box, for a rect; the text's ink, for a label. */
  const inkOf = (n: SceneNode) => {
    if (n.kind === "rect") return { x0: n.x, y0: n.y, x1: n.x + n.w, y1: n.y + n.h };
    if (n.kind !== "text") return null;
    const w = textWidth(n.text, n.fontSize, n.bold);
    const x0 = n.align === "right" ? n.x + n.w - w : n.align === "center" ? n.x + (n.w - w) / 2 : n.x;
    // `valign: middle`/`top` boxes carry their baseline inside the box.
    return { x0, y0: n.y, x1: x0 + w, y1: n.y + n.h };
  };
  const outside = (cfg: ChartConfig) =>
    buildChart(cfg).nodes.flatMap((n) => {
      const b = inkOf(n);
      if (!b || !n.name || n.name === "title") return [];
      const bad = b.y1 > cfg.height + SLACK || b.y0 < -SLACK || b.x1 > cfg.width + SLACK || b.x0 < -SLACK;
      return bad ? [`${n.name} y=${b.y0.toFixed(1)}..${b.y1.toFixed(1)}`] : [];
    });

  it("the reported charts: a 1% last group no longer hangs off the bottom", () => {
    const tiny = treemap(
      [
        ["EMEA | DK", 400],
        ["EMEA | SE", 200],
        ["APAC | JP", 250],
        ["APAC | CN", 140],
        ["LATAM | BR", 10],
      ],
      480,
      300,
    );
    expect(outside(tiny)).toEqual([]);
    const vanishing = treemap(
      [
        ["Big | A", 602935],
        ["Mid | B", 105017],
        ["Tiny | C", 10],
      ],
      600,
      400,
    );
    expect(outside(vanishing)).toEqual([]);
  });

  it("a group's header sits inside the cell it names, and a cell too thin for one has none", () => {
    // Between ~5 and 14pt tall mid-column the 15pt box also overran the next group's header.
    for (const tail of [400, 40, 20, 12, 6, 3, 1]) {
      const cfg = treemap(
        [
          ["A | a1", 500],
          ["A | a2", 300],
          ["B | b1", 250],
          ["C | c1", tail],
        ],
        480,
        300,
      );
      const nodes = buildChart(cfg).nodes;
      for (const g of named<RectNode>(nodes, "rect", /^group-\d+$/)) {
        const header = named<TextNode>(nodes, "text", `group-label-${g.name!.slice(6)}`)[0];
        if (!header) continue;
        expect(header.y + header.h, `${g.name} tail=${tail}`).toBeLessThanOrEqual(g.y + g.h + SLACK);
      }
      expect(outside(cfg), `tail=${tail}`).toEqual([]);
    }
  });

  it("a group that keeps its header is laid out as it always was", () => {
    const cfg = treemap(
      [
        ["A | a1", 500],
        ["A | a2", 300],
        ["B | b1", 250],
        ["B | b2", 150],
      ],
      480,
      300,
    );
    const nodes = buildChart(cfg).nodes;
    const label = named<TextNode>(nodes, "text", "group-label-0")[0];
    const group = named<RectNode>(nodes, "rect", "group-0")[0];
    expect(label.y).toBe(group.y + 1);
    expect(label.h).toBe(15);
  });

  it("no hierarchical treemap in a seeded sweep leaves its frame", () => {
    let seed = 20261025;
    const rnd = () => (seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296;
    const bad: string[] = [];
    for (let k = 0; k < 150; k++) {
      const groups = 2 + Math.floor(rnd() * 4);
      const rows: [string, number][] = [];
      for (let g = 0; g < groups; g++) {
        const members = 1 + Math.floor(rnd() * 3);
        // One group in three is a sliver.
        const scale = rnd() < 0.34 ? 0.002 + rnd() * 0.02 : 1;
        for (let m = 0; m < members; m++) rows.push([`G${g} | m${m}`, Math.max(1, Math.round(rnd() * 800 * scale))]);
      }
      const cfg = treemap(rows, rnd() < 0.5 ? 480 : 360, rnd() < 0.5 ? 300 : 220);
      const out = outside(cfg);
      if (out.length) bad.push(`${JSON.stringify(rows)} -> ${out.join(", ")}`);
    }
    expect(bad).toEqual([]);
  });
});

describe("r3-39: an x band stays inside the plot, whatever the category width means", () => {
  const banded = (cfg: ChartConfig, to: number): ChartConfig =>
    junk({ ...cfg, decorations: { ...cfg.decorations, bands: [{ axis: "x", from: 0, to, label: "Phase one" }] } });
  const extent = (cfg: ChartConfig) => {
    const nodes = buildChart(cfg).nodes;
    return {
      band: named<RectNode>(nodes, "rect", "band-0")[0],
      label: named<TextNode>(nodes, "text", "band-label-0")[0],
    };
  };
  const cases: [string, ChartConfig, number][] = [
    ["mekko", sampleConfig("mekko"), 2],
    ["violin", sampleConfig("violin"), 2],
    ["line", sampleConfig("line"), 4],
    ["area", sampleConfig("area"), 4],
    ["line, no axes", junk({ ...sampleConfig("line"), decorations: undefined }), 4],
    ["treemap", sampleConfig("treemap"), 2],
    ["stacked, gapWidth 0", junk({ ...sampleConfig("stacked"), gapWidth: 0 }), 2],
    ["stacked, gapWidth 20", junk({ ...sampleConfig("stacked"), gapWidth: 20 }), 2],
    ["line at 960x540", junk({ ...sampleConfig("line"), width: 960, height: 540 }), 4],
    ["horizontal mekko", junk({ ...sampleConfig("mekko"), horizontal: true, width: 300, height: 300 }), 2],
  ];

  it.each(cases)("%s: the band and its label are on the canvas", (_name, cfg, to) => {
    const { band, label } = extent(banded(cfg, to));
    expect(band, "the band is drawn").toBeDefined();
    expect(band.x).toBeGreaterThanOrEqual(0);
    expect(band.y).toBeGreaterThanOrEqual(0);
    expect(band.x + band.w).toBeLessThanOrEqual(cfg.width + 0.01);
    expect(band.y + band.h).toBeLessThanOrEqual(cfg.height + 0.01);
    if (label) {
      expect(label.x).toBeGreaterThanOrEqual(0);
      expect(label.y).toBeGreaterThanOrEqual(0);
    }
  });

  it("the plot bounds it: a band over every category covers the plot and no more", () => {
    // Along the category axis the baseline runs the width of the plot.
    const cfg = banded(sampleConfig("mekko"), 2);
    const nodes = buildChart(cfg).nodes;
    const base = named<LineNode>(nodes, "line", "baseline")[0];
    const band = named<RectNode>(nodes, "rect", "band-0")[0];
    expect(band.x).toBeGreaterThanOrEqual(Math.min(base.x1, base.x2) - 0.01);
    expect(band.x + band.w).toBeLessThanOrEqual(Math.max(base.x1, base.x2) + 0.01);
  });

  it("a default column chart keeps the band it has always drawn", () => {
    const { band } = extent(banded(sampleConfig("stacked"), 2));
    // Measured before the clamp existed: the plot starts at x=2 and three of the
    // sample's four slots end at 308.
    expect(band.x).toBeCloseTo(2.0, 1);
    expect(band.x + band.w).toBeCloseTo(308.0, 1);
  });
});

describe("r3-38: a sideways combo line labels the point at the plot's right edge", () => {
  const pareto = (horizontal: boolean): ChartConfig =>
    junk({
      kind: "clustered",
      ...DEFAULT_SIZE,
      pareto: true,
      horizontal,
      data: { categories: ["A", "B", "C", "D"], series: [series("Defects", [30, 50, 5, 15])] },
    });
  const combo = (line: number[], horizontal: boolean): ChartConfig =>
    junk({
      kind: "combo",
      ...DEFAULT_SIZE,
      horizontal,
      data: {
        categories: ["Q1", "Q2", "Q3"].slice(0, line.length),
        series: [series("Revenue", [10, 20, 15].slice(0, line.length)), { name: "Margin", type: "line", values: line }],
      },
    });
  const labels = (cfg: ChartConfig) =>
    named<TextNode>(buildChart(cfg).nodes, "text", /^combo-label-0-\d+$/).map((t) => t.text);
  const markerX = (cfg: ChartConfig, c: number) => {
    const m = named<RectNode>(buildChart(cfg).nodes, "rect", `combo-marker-0-${c}`)[0];
    return m.x + m.w / 2;
  };

  it("a horizontal Pareto keeps its cumulative 100", () => {
    expect(labels(pareto(false))).toEqual(["50", "80", "95", "100"]);
    expect(labels(pareto(true))).toEqual(["50", "80", "95", "100"]);
  });

  it("any line value whose label cannot stand to the right of its mark takes the left", () => {
    // 0-30 axis: 30 and 29.5 were dropped, 29 kept.
    expect(labels(combo([12, 30], true))).toEqual(["12", "30"]);
    expect(labels(combo([12, 29.5], true))).toHaveLength(2);
    expect(labels(combo([12, 29], true))).toEqual(["12", "29"]);
    // Upright was never in doubt, and is where this is compared.
    expect(labels(combo([12, 30], false))).toEqual(["12", "30"]);
  });

  it("the flipped label stands left of its mark, clear of it, and on the canvas", () => {
    const cfg = pareto(true);
    const nodes = buildChart(cfg).nodes;
    const t = named<TextNode>(nodes, "text", "combo-label-0-3")[0];
    const mx = markerX(cfg, 3);
    const ink = textWidth(t.text, t.fontSize);
    expect(t.align).toBe("right");
    expect(t.x + t.w).toBeLessThan(mx);
    expect(t.x + t.w - ink).toBeGreaterThanOrEqual(0);
  });

  it("a label that fits on the right is exactly where it was: beside the mark, left-aligned", () => {
    const cfg = combo([12, 30], true);
    const t = named<TextNode>(buildChart(cfg).nodes, "text", "combo-label-0-0")[0];
    expect(t.align).toBe("left");
    expect(t.x).toBeGreaterThan(markerX(cfg, 0));
  });
});

describe("r3-70: a cell of the X extent or 100% row that is no number is a blank one", () => {
  const mekko = (xExtent: unknown[], horizontal: boolean): ChartConfig =>
    junk({
      kind: "mekko",
      ...DEFAULT_SIZE,
      ...(horizontal ? { horizontal } : {}),
      data: {
        categories: ["EMEA", "Americas", "APAC"],
        series: [series("A", [42, 55, 18]), series("B", [28, 30, 22])],
        xExtent,
      },
    });

  it.each([false, true])("horizontal=%s: 'n/a', {} and Infinity draw the chart the null cell draws", (horizontal) => {
    const clean = scene(mekko([50, 30, null], horizontal));
    for (const bad of ["n/a", "—", {}, Number.POSITIVE_INFINITY, "1,200"])
      expect(scene(mekko([50, 30, bad], horizontal)), JSON.stringify(bad)).toBe(clean);
  });

  it("every node stays on the canvas", () => {
    for (const horizontal of [false, true]) {
      const s = buildChart(mekko([50, 30, "n/a"], horizontal));
      for (const n of s.nodes) {
        if (n.kind !== "rect") continue;
        expect(n.y + n.h, n.name).toBeLessThanOrEqual(s.height + 1);
        expect(n.x + n.w, n.name).toBeLessThanOrEqual(s.width + 1);
      }
    }
  });

  it("a numeric string is still the number it spells", () => {
    expect(scene(mekko(["50", "30", "20"], false))).toBe(scene(mekko([50, 30, 20], false)));
  });

  it("the 100% row is cleaned the same way", () => {
    const stacked100 = (hundredPercent: unknown[]): ChartConfig =>
      junk({
        kind: "stacked100",
        ...DEFAULT_SIZE,
        data: {
          categories: ["A", "B", "C"],
          series: [series("X", [3, 4, 5]), series("Y", [1, 2, 3])],
          hundredPercent,
        },
      });
    expect(scene(stacked100([8, Number.POSITIVE_INFINITY, 8]))).toBe(scene(stacked100([8, null, 8])));
    expect(scene(stacked100([8, "n/a", 8]))).toBe(scene(stacked100([8, null, 8])));
  });
});

describe("r3-42: a mekko column nobody filled in prints no total", () => {
  // The X extent row gives the empty column a width, so the label is not dropped
  // for want of room; EMEA is blank in every series.
  const mekko = (horizontal: boolean, xExtent: (number | null)[] | undefined): ChartConfig => ({
    kind: "mekko",
    ...DEFAULT_SIZE,
    ...(horizontal ? { horizontal } : {}),
    data: {
      categories: ["EMEA", "Americas", "APAC"],
      series: [series("A", [null, 50, 25]), series("B", [null, 35, 15])],
      ...(xExtent ? { xExtent } : {}),
    },
  });
  const totals = (cfg: ChartConfig) =>
    named<TextNode>(buildChart(cfg).nodes, "text", /^total-\d+$/).map((t) => `${t.name}=${t.text}`);

  it.each([false, true])("with an X extent, horizontal=%s: only the measured columns are totalled", (horizontal) => {
    const cfg = mekko(horizontal, [100, 200, 150]);
    for (const input of [cfg, normalizeConfig(cfg)]) {
      const got = totals(input);
      expect(got).toContain("total-1=85");
      expect(got).toContain("total-2=40");
      expect(got.some((t) => t.startsWith("total-0="))).toBe(false);
    }
  });

  it.each([false, true])("a typed zero in the empty column is still a measurement, horizontal=%s", (horizontal) => {
    // The other half of the line: a 0 somebody typed is data, and keeps its total.
    const cfg = mekko(horizontal, [100, 200, 150]);
    const typed = {
      ...cfg,
      data: { ...cfg.data, series: cfg.data.series.map((s) => ({ ...s, values: s.values.map((v) => v ?? 0) })) },
    };
    expect(totals(typed)).toContain("total-0=0");
  });
});
