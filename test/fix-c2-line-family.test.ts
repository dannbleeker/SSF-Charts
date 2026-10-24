import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { categoryRooms } from "../src/core/layout/frame";
import { splineSegments } from "../src/core/layout/line";
import { sampleConfig } from "../src/core/samples";
import { textWidth } from "../src/core/scene";
import type { LineNode, RectNode, Scene, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig, ChartKind } from "../src/core/types";

/**
 * Eight findings in the line, area and combo layouts: date spacing that moved
 * categories out of the order they were given, a category axis fitted to its
 * first gap, a sideways line that lost its largest label, a series with no name
 * anywhere, a combo line with no name anywhere, a smoothed curve that left the
 * plot, a bump chart that drew past the canvas, and a sparkline that threw its
 * scale away.
 *
 * Every block names its finding and carries the control that keeps it honest:
 * the case that has to be left alone, so a fix cannot pass by deleting the
 * feature it repaired.
 */

const nodesOf = (scene: Scene, prefix: string) => scene.nodes.filter((n) => n.name?.startsWith(prefix));
const rectX = (n: SceneNode) => (n as RectNode).x + (n as RectNode).w / 2;
const rectY = (n: SceneNode) => (n as RectNode).y + (n as RectNode).h / 2;

/* -------------------------------------------------------------------------- */

describe("r1-0: date spacing keeps the categories in the order they were given", () => {
  const line = (categories: string[], extra: Partial<ChartConfig> = {}): ChartConfig =>
    ({
      kind: "line",
      ...DEFAULT_SIZE,
      data: { categories, series: [{ name: "S", values: categories.map((_, i) => 10 + i) }] },
      ...extra,
    }) as ChartConfig;

  /** Marker centres along the category axis, in category order. */
  const markerAxis = (scene: Scene, axis: "x" | "y") =>
    nodesOf(scene, "marker-0-")
      .sort((a, b) => Number(a.name!.split("-")[2]) - Number(b.name!.split("-")[2]))
      .map(axis === "x" ? rectX : rectY);

  const increasing = (xs: number[]) => xs.every((x, i) => i === 0 || x > xs[i - 1]);

  // The space form is read by the parser as a DAY in V8's default year (a decision
  // test/dates.test.ts pins), and a year-less day label wraps at new year.
  const wrapped: [string, string[]][] = [
    ["Oct 23 .. Feb 24", ["Oct 23", "Nov 23", "Dec 23", "Jan 24", "Feb 24"]],
    ["weekly 14 Dec .. 11 Jan", ["14 Dec", "21 Dec", "28 Dec", "4 Jan", "11 Jan"]],
    ["Excel's d-mmm display", ["21-Dec", "28-Dec", "4-Jan", "11-Jan"]],
    ["out-of-order ISO months", ["2024-03", "2024-01", "2024-02", "2024-04"]],
    ["a duplicated month", ["2024-01", "2024-02", "2024-02", "2024-03"]],
    ["newest first", ["2024-04", "2024-03", "2024-02", "2024-01"]],
  ];

  it.each(wrapped)("upright, %s: every point sits right of the one before it", (_label, cats) => {
    const scene = buildChart(line(cats));
    const xs = markerAxis(scene, "x");
    expect(xs, `markers in category order ran backwards: ${xs.map((x) => x.toFixed(1)).join(", ")}`).toSatisfy(
      increasing,
    );
    // …and the line is drawn left to right with them, rather than slashing back
    // across the plot between two categories.
    const segs = nodesOf(scene, "line-0-").filter((n): n is LineNode => n.kind === "line");
    expect(segs).toHaveLength(cats.length - 1);
    for (const s of segs) expect(s.x2, s.name).toBeGreaterThan(s.x1);
  });

  it.each(wrapped)("sideways, %s: every point sits below the one before it", (_label, cats) => {
    const scene = buildChart(line(cats, { horizontal: true }));
    const ys = markerAxis(scene, "y");
    expect(ys, `markers in category order ran backwards: ${ys.map((y) => y.toFixed(1)).join(", ")}`).toSatisfy(
      increasing,
    );
    const segs = nodesOf(scene, "line-0-").filter((n): n is LineNode => n.kind === "line");
    for (const s of segs) expect(s.y2, s.name).toBeGreaterThan(s.y1);
  });

  it("falls back to the even spacing every other category list gets", () => {
    const xs = markerAxis(buildChart(line(["Oct 23", "Nov 23", "Dec 23", "Jan 24", "Feb 24"])), "x");
    const gaps = xs.slice(1).map((x, i) => x - xs[i]);
    for (const g of gaps) expect(g).toBeCloseTo(gaps[0], 6);
  });

  it("a confidence ribbon over out-of-order dates is never laid out backwards", () => {
    // The ribbon marches slab by slab from one centre to the next, so a centre
    // that is LEFT of the previous one gave it a negative span and negative-width
    // slabs.
    const cats = ["2024-04", "2024-03", "2024-02", "2024-01"];
    const scene = buildChart({
      kind: "line",
      ...DEFAULT_SIZE,
      data: {
        categories: cats,
        series: [
          { name: "Actual", values: [14, 12, 11, 10] },
          { name: "Band low", values: [12, 10, 9, 8] },
          { name: "Band high", values: [16, 14, 13, 12] },
        ],
      },
    } as ChartConfig);
    const slabs = nodesOf(scene, "band-ribbon-").filter((n): n is RectNode => n.kind === "rect");
    expect(slabs.length).toBeGreaterThan(0);
    for (const s of slabs) expect(s.w, s.name).toBeGreaterThan(0);
  });

  // CONTROLS: the feature this protects is still there.
  it("control: dates that run forward still space proportionally to time", () => {
    const xs = markerAxis(buildChart(line(["2025-01", "2025-02", "2025-12"])), "x");
    expect(xs[1] - xs[0]).toBeLessThan((xs[2] - xs[1]) / 5);
  });

  it("control: the hyphen and four-digit year forms still run oldest to newest", () => {
    for (const cats of [
      ["Oct-23", "Nov-23", "Dec-23", "Jan-24", "Feb-24"],
      ["Oct 2023", "Nov 2023", "Dec 2023", "Jan 2024", "Feb 2024"],
    ]) {
      expect(markerAxis(buildChart(line(cats)), "x")).toSatisfy(increasing);
    }
  });

  it("control: sideways dates that run forward still space proportionally to time", () => {
    const ys = markerAxis(buildChart(line(["2025-01", "2025-02", "2025-12"], { horizontal: true })), "y");
    expect(ys[1] - ys[0]).toBeLessThan((ys[2] - ys[1]) / 5);
  });
});

/* -------------------------------------------------------------------------- */

describe("r1-7: a date-spaced axis fits each name to the room ITS OWN neighbours leave it", () => {
  const months = (year: number) => Array.from({ length: 12 }, (_, i) => `${year}-${String(i + 1).padStart(2, "0")}`);
  const line = (categories: string[], size: [number, number], horizontal = false): ChartConfig =>
    ({
      kind: "line",
      width: size[0],
      height: size[1],
      horizontal,
      data: { categories, series: [{ name: "S", values: categories.map((_, i) => 10 + i) }] },
    }) as ChartConfig;

  const categoryNames = (scene: Scene) =>
    scene.nodes
      .filter((n): n is TextNode => n.kind === "text" && /^category-\d+$/.test(n.name ?? ""))
      .sort((a, b) => Number(a.name!.split("-")[1]) - Number(b.name!.split("-")[1]));
  const inkBox = (t: TextNode) => {
    const w = textWidth(t.text, t.fontSize);
    return { x0: t.x + (t.w - w) / 2, x1: t.x + (t.w + w) / 2 };
  };

  it("the helper: evenly spaced centres keep the one width they always had, to the bit", () => {
    const centers = Array.from({ length: 7 }, (_, i) => 10 + 33.3 * (i + 0.5));
    expect(categoryRooms(centers, 999, 10, 10 + 33.3 * 7)).toEqual(centers.map(() => centers[1] - centers[0]));
    expect(categoryRooms([5], 999, 0, 10)).toEqual([999]);
  });

  it("the helper: uneven centres give each name the narrower of its two gaps, and no more than the plot allows", () => {
    // gaps 3, 100, 20 -> 3, 3, 20, 20; the cap keeps the first and last inside the plot [0, 150].
    expect(categoryRooms([10, 13, 113, 133], 999, 0, 150)).toEqual([3, 3, 20, 20]);
    expect(categoryRooms([10, 13, 113, 133], 999, 9, 140)).toEqual([2, 3, 20, 14]);
  });

  it("a one-day first gap does not take the whole axis with it (960x540)", () => {
    const cats = ["2024-01-01", "2024-01-02", "2024-06-01", "2024-12-01"];
    const names = categoryNames(buildChart(line(cats, [960, 540])));
    // The two names a day apart cannot both be read; the other two have 100pt+
    // each and used to lose their axis to the pair.
    expect(
      names.map((n) => n.text),
      "the whole category axis was dropped",
    ).toEqual(expect.arrayContaining(["2024-06-01", "2024-12-01"]));
    // …at the chart's own size, not crushed to the floor by the two that cannot be drawn.
    for (const n of names) expect(n.fontSize, n.name).toBe(10);
  });

  it("a year-long first gap does not size the monthly names that follow it (960x540)", () => {
    const cats = ["2023-01", ...months(2024)];
    const names = categoryNames(buildChart(line(cats, [960, 540])));
    expect(names).toHaveLength(cats.length);
    for (let i = 1; i < names.length; i++) {
      expect(inkBox(names[i - 1]).x1, `${names[i - 1].text} runs into ${names[i].text}`).toBeLessThanOrEqual(
        inkBox(names[i]).x0 + 0.01,
      );
    }
    // The long gap's name is still there, and still on the chart.
    expect(names[0].text).toBe("2023-01");
    expect(names[0].x).toBeGreaterThanOrEqual(0);
  });

  it("sideways, a one-day first gap does not take the whole axis with it either", () => {
    const names = categoryNames(
      buildChart(line(["2024-01-01", "2024-01-02", "2024-06-01", "2024-12-01"], [480, 300], true)),
    );
    expect(names.map((n) => n.text)).toEqual(expect.arrayContaining(["2024-06-01", "2024-12-01"]));
  });

  it("sideways, a year-long first gap does not size the monthly rows that follow it", () => {
    const cats = ["2023-01", ...months(2024)];
    const names = categoryNames(buildChart(line(cats, [480, 300], true)));
    expect(names).toHaveLength(cats.length);
    for (let i = 1; i < names.length; i++) {
      const pitch = names[i].y - names[i - 1].y;
      expect(pitch, `${names[i - 1].text} is drawn through ${names[i].text}`).toBeGreaterThanOrEqual(
        names[i].fontSize * 1.5 - 0.01,
      );
    }
  });

  // CONTROLS
  it("control: a list that is not dates still gives every name the one slot width", () => {
    const names = categoryNames(buildChart(line(["North", "South", "East", "West"], [480, 300])));
    expect(names).toHaveLength(4);
    for (const n of names) expect(n.w).toBe(names[0].w);
    expect(names.map((n) => n.fontSize)).toEqual([10, 10, 10, 10]);
  });

  it("control: months a month apart keep their full-size names", () => {
    const names = categoryNames(buildChart(line(months(2024), [960, 540])));
    expect(names).toHaveLength(12);
    for (const n of names) expect(n.fontSize, n.name).toBe(10);
  });
});

/* -------------------------------------------------------------------------- */

describe("r1-6: a sideways line chart keeps the label of the point at the edge, and its labelContent", () => {
  const sideways = (values: number[], extra: Partial<ChartConfig> = {}, decorations: object = {}): ChartConfig =>
    ({
      kind: "line",
      horizontal: true,
      ...DEFAULT_SIZE,
      data: { categories: ["North", "South"], series: [{ name: "Sales", values }] },
      decorations,
      ...extra,
    }) as ChartConfig;

  const labelTexts = (scene: Scene) =>
    nodesOf(scene, "label-0-")
      .filter((n): n is TextNode => n.kind === "text")
      .map((n) => n.text);
  const inkOf = (t: TextNode) => {
    const ink = textWidth(t.text, t.fontSize);
    const x = t.align === "right" ? t.x + t.w - ink : t.align === "center" ? t.x + (t.w - ink) / 2 : t.x;
    return { x0: x, x1: x + ink };
  };

  // Each of these has its largest value ON the top tick, so its point is the
  // plot's right edge.
  it.each([
    [[12, 30], "30"],
    [[120, 300], "300"],
    [[12345, 30000], "30,000"],
    [[12, 40], "40"],
    [[30, 12], "30"],
  ])("%j: the label of the maximum (%s) is drawn", (values, text) => {
    const scene = buildChart(sideways(values));
    expect(labelTexts(scene), "the label of the largest value was dropped").toContain(text);
    for (const n of nodesOf(scene, "label-0-").filter((x): x is TextNode => x.kind === "text")) {
      const { x0, x1 } = inkOf(n);
      expect(x0, n.name).toBeGreaterThanOrEqual(-0.5);
      expect(x1, n.name).toBeLessThanOrEqual(DEFAULT_SIZE.width + 0.5);
    }
  });

  it("a value just under the maximum is whole, not cut to a fragment that reads as another number", () => {
    expect(labelTexts(buildChart(sideways([12345, 29500])))).toEqual(["12,345", "29,500"]);
    expect(labelTexts(buildChart(sideways([12345, 29000])))).toEqual(["12,345", "29,000"]);
  });

  it("the label that cannot fit on the right stands to the LEFT of its marker, clear of it", () => {
    const scene = buildChart(sideways([12, 30]));
    const marker = nodesOf(scene, "marker-0-1")[0] as RectNode;
    const label = nodesOf(scene, "label-0-1")[0] as TextNode | undefined;
    expect(label, "the label of the maximum was dropped").toBeDefined();
    expect(label!.align).toBe("right");
    expect(inkOf(label!).x1).toBeLessThan(marker.x);
  });

  it("holds at a larger canvas and font too", () => {
    const big = buildChart(sideways([12, 30], { width: 960, height: 540 }));
    expect(labelTexts(big)).toContain("30");
    const large = buildChart(sideways([12, 30], { style: { fontSize: 14 } as ChartConfig["style"] }));
    expect(labelTexts(large)).toContain("30");
  });

  it("labelContent is honoured sideways exactly as it is upright", () => {
    const content = (labelContent: string[]) => ({ labelContent });
    const read = (horizontal: boolean, labelContent: string[]) =>
      labelTexts(buildChart(sideways([12, 20], { horizontal }, content(labelContent))));
    for (const parts of [["category", "value"], ["series"], ["series", "value"], ["value"]]) {
      expect(read(true, parts), parts.join("+")).toEqual(read(false, parts));
    }
    expect(read(true, ["category", "value"])).toEqual(["North 12", "South 20"]);
  });

  // CONTROLS
  it("control: a label with room on the right stays where it was, to the right of its point", () => {
    const scene = buildChart(sideways([12, 18]));
    for (const n of nodesOf(scene, "label-0-").filter((x): x is TextNode => x.kind === "text")) {
      const marker = nodesOf(scene, n.name!.replace("label", "marker"))[0] as RectNode;
      expect(n.align, n.name).toBe("left");
      expect(n.x, n.name).toBeCloseTo(marker.x + marker.w / 2 + 4, 6);
      expect(n.w, n.name).toBe(30);
    }
  });

  it("control: totals reserve their gutter and the label stays on the right", () => {
    const scene = buildChart(sideways([12, 30], {}, { totals: true }));
    expect(labelTexts(scene)).toEqual(["12", "30"]);
    for (const n of nodesOf(scene, "label-0-").filter((x): x is TextNode => x.kind === "text")) {
      expect(n.align).toBe("left");
    }
  });
});

describe("r1-6 (sibling): the sideways dot, lollipop and range plots have the same edge and the same labelContent", () => {
  const dots = (barStyle: string, values: number[], labelContent?: string[]): ChartConfig =>
    ({
      kind: "clustered",
      horizontal: true,
      ...DEFAULT_SIZE,
      data: { categories: ["North", "South"], series: [{ name: "Sales", values }] },
      decorations: { barStyle, ...(labelContent ? { labelContent } : {}) },
    }) as ChartConfig;
  const labelTexts = (scene: Scene) =>
    nodesOf(scene, "label-0-")
      .filter((n): n is TextNode => n.kind === "text")
      .map((n) => n.text);

  it.each(["dot", "lollipop", "range"])("%s: the label of the maximum is drawn, and whole", (barStyle) => {
    expect(labelTexts(buildChart(dots(barStyle, [12, 30])))).toEqual(["12", "30"]);
    expect(labelTexts(buildChart(dots(barStyle, [12345, 29500])))).toEqual(["12,345", "29,500"]);
  });

  it.each(["dot", "lollipop"])("%s: labelContent is read", (barStyle) => {
    expect(labelTexts(buildChart(dots(barStyle, [12, 20], ["category", "value"])))).toEqual(["North 12", "South 20"]);
  });

  it("control: a dot with room on its right keeps its label there", () => {
    const scene = buildChart(dots("dot", [12, 18]));
    for (const n of nodesOf(scene, "label-0-").filter((x): x is TextNode => x.kind === "text")) {
      expect(n.align, n.name).toBe("left");
    }
  });
});

/* -------------------------------------------------------------------------- */

describe("r1-5: a series with no mark in the last category is still named", () => {
  const KINDS = ["stacked", "clustered", "stacked100", "line", "mekko", "combo"] as const;
  const chart = (
    kind: string,
    legacy: (number | null)[],
    size: [number, number] = [480, 300],
    series: { name: string; values: (number | null)[]; color?: string }[] = [],
    extra: object = {},
  ): ChartConfig =>
    ({
      kind,
      width: size[0],
      height: size[1],
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Hardware", values: [50, 55, 60] },
          { name: "Legacy", values: legacy },
          { name: "Services", values: [20, 25, 30] },
          ...series,
        ],
      },
      ...extra,
    }) as ChartConfig;

  /** Every text on the chart that carries a series name: the gutter and the combo line's own. */
  const gutter = (scene: Scene) =>
    scene.nodes.filter((n): n is TextNode => n.kind === "text" && /series-label/.test(n.name ?? ""));
  const named = (scene: Scene, name: string) => gutter(scene).filter((n) => n.text === name);

  describe.each(KINDS)("%s", (kind) => {
    it.each([[[480, 300]], [[960, 540]]] as [[number, number]][])(
      "a blank last cell: its earlier marks are drawn and its name is on the chart (%j)",
      (size) => {
        const scene = buildChart(chart(kind, [30, 20, null], size));
        expect(nodesOf(scene, "seg-1-").length + nodesOf(scene, "marker-1-").length, "its marks").toBeGreaterThan(0);
        expect(
          named(scene, "Legacy"),
          `a ${kind} series whose last cell is blank is drawn and named nowhere`,
        ).toHaveLength(1);
        // …and the others are still named, once each.
        for (const other of ["Hardware", "Services"]) expect(named(scene, other), other).toHaveLength(1);
      },
    );
  });

  // A measured 0 in the last category is a bar of no height: nothing to anchor
  // a name to on the kinds that draw bars. A line plots the 0 and was named.
  it.each(["stacked", "clustered", "stacked100", "mekko", "combo", "line"])(
    "%s: a 0 in the last cell does not take the name with it",
    (kind) => {
      expect(named(buildChart(chart(kind, [30, 20, 0])), "Legacy")).toHaveLength(1);
    },
  );

  it("the gutter keeps the stack's order: the unmarked series sits between the two it stacks between", () => {
    const scene = buildChart(chart("stacked", [30, 20, null]));
    const y = (name: string) => named(scene, name)[0].y;
    // Hardware is the bottom of the stack, Services the top; the screen's y runs down.
    expect(y("Services")).toBeLessThan(y("Legacy"));
    expect(y("Legacy")).toBeLessThan(y("Hardware"));
  });

  it("a line that ended early is named level with where it ended", () => {
    const scene = buildChart(chart("line", [30, 20, null]));
    const end = nodesOf(scene, "marker-1-1")[0] as RectNode;
    const label = named(scene, "Legacy")[0];
    expect(Math.abs(label.y + label.h / 2 - (end.y + end.h / 2))).toBeLessThan(16);
  });

  // CONTROLS
  it.each(KINDS)("control: %s with a value in the last cell names everyone exactly as it did", (kind) => {
    const scene = buildChart(chart(kind, [30, 20, 15]));
    for (const name of ["Hardware", "Legacy", "Services"]) expect(named(scene, name), name).toHaveLength(1);
  });

  it.each(["stacked", "clustered", "line", "mekko"])(
    "control: %s — a series with no data at all stays unnamed, as it was",
    (kind) => {
      const scene = buildChart(chart(kind, [null, null, null]));
      expect(named(scene, "Legacy")).toHaveLength(0);
      expect(named(scene, "Hardware")).toHaveLength(1);
    },
  );

  it("control: a transparent spacer with a blank last cell is not named", () => {
    const scene = buildChart({
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Base", values: [10, 20, null], color: "transparent" },
          { name: "Value", values: [30, 25, 20] },
        ],
      },
    } as ChartConfig);
    expect(named(scene, "Base")).toHaveLength(0);
    expect(named(scene, "Value")).toHaveLength(1);
  });

  it("control: a sideways chart still names every series through its legend", () => {
    const scene = buildChart(chart("clustered", [30, 20, null], [480, 300], [], { horizontal: true }));
    const legend = scene.nodes.filter((n): n is TextNode => n.kind === "text" && /^legend-\d+$/.test(n.name ?? ""));
    expect(legend.map((n) => n.text).sort()).toEqual(["Hardware", "Legacy", "Services"]);
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-11: an upright combo names its line wherever its column names are not merged with it", () => {
  const combo = (extra: Partial<ChartConfig>, size: [number, number]): ChartConfig =>
    ({ ...sampleConfig("combo"), width: size[0], height: size[1], ...extra }) as ChartConfig;
  const pareto = (size: [number, number]): ChartConfig =>
    ({ ...sampleConfig("clustered"), pareto: true, width: size[0], height: size[1] }) as ChartConfig;

  const SIZES: [number, number][] = [
    [300, 200],
    [480, 300],
    [960, 540],
  ];
  /** The line's own name on the chart: `combo-series-label-<n>` carries it upright. */
  const lineName = (scene: Scene) =>
    scene.nodes.filter((n): n is TextNode => n.kind === "text" && /^combo-series-label-\d+$/.test(n.name ?? ""));

  // Every one of these drew `combo-series-label-0` before commit a2b87db and
  // none has since: the loop handed the name to a spread that then refused the
  // chart.
  const refused: [string, (size: [number, number]) => ChartConfig, string][] = [
    ["a secondary axis", (s) => combo({ secondaryAxis: true }, s), "Margin %"],
    ["a 100% base", (s) => combo({ combo: { columns: "stacked100" } }, s), "Margin %"],
    ["a mekko base", (s) => combo({ combo: { columns: "mekko" } }, s), "Margin %"],
    ["a waterfall base", (s) => combo({ combo: { columns: "waterfall" } }, s), "Margin %"],
    ["an area base", (s) => combo({ combo: { columns: "area" } }, s), "Margin %"],
    ["a pareto", pareto, "Cumulative %"],
  ];

  describe.each(refused)("%s", (_label, build, name) => {
    it.each(SIZES)("draws the line's name (%ix%i)", (w, h) => {
      const scene = buildChart(build([w, h]));
      const named = lineName(scene);
      expect(
        named.map((n) => n.name),
        `no text names the line "${name}"`,
      ).toEqual(["combo-series-label-0"]);
      // A small frame may clip the name, never lose it.
      expect(name.startsWith(named[0].text.replace(/…$/, ""))).toBe(true);
      expect(named[0].text.length).toBeGreaterThan(2);
    });
  });

  it("the secondary axis keeps its own numbers beside the name", () => {
    const scene = buildChart(combo({ secondaryAxis: true }, [480, 300]));
    expect(nodesOf(scene, "secondary-axis").length).toBeGreaterThan(2);
    expect(["Services", "Product"].every((n) => scene.nodes.some((x) => x.kind === "text" && x.text === n))).toBe(true);
  });

  // CONTROLS: the charts the merged spread DOES take are exactly as they were.
  it.each([
    ["stacked", {}],
    ["clustered", { combo: { columns: "clustered" } }],
    ["independent axes", { combo: { lineAxes: "independent" } }],
  ] as [string, Partial<ChartConfig>][])("control: %s is named once, by the merged spread", (_label, extra) => {
    const scene = buildChart(combo(extra, [480, 300]));
    expect(lineName(scene).map((n) => n.text)).toEqual(["Margin %"]);
  });

  it("control: a sideways combo names its line on the plot, through its legend too", () => {
    const scene = buildChart(combo({ secondaryAxis: true, horizontal: true }, [480, 300]));
    expect(lineName(scene).map((n) => n.text)).toEqual(["Margin %"]);
    expect(
      scene.nodes.some((n) => n.kind === "text" && /^legend-\d+$/.test(n.name ?? "") && n.text === "Margin %"),
    ).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-11, the other half: the names it draws again do not land on the text beside them", () => {
  /**
   * Naming those lines again put 90 pairs of text over text into the overlap
   * sweep (785 -> 875) in three families: a pareto at 300x60 (46), a pareto with
   * long series names at 480x300 (13, a tick number under the name), and a
   * ten-series dual-axis combo over five frames (31). Each is built here the way
   * `overlap-budget.test.ts` builds it, and measured by the same rule: ink that
   * shares more than a point squared.
   */
  const inkOf = (t: TextNode) => {
    const w = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    const base =
      t.valign === "top"
        ? t.y + t.fontSize
        : t.valign === "bottom"
          ? t.y + t.h - t.fontSize * 0.25
          : t.y + t.h / 2 + t.fontSize * 0.36;
    return { x0: x, y0: base - t.fontSize * 0.8, x1: x + w, y1: base + t.fontSize * 0.21 };
  };
  const texts = (scene: Scene) =>
    scene.nodes.filter((n): n is TextNode => n.kind === "text" && !!String(n.text).trim());
  /** Every pair of texts on the chart with a gutter name in it that shares ink, as "a / b". */
  const gutterPairs = (scene: Scene) => {
    const ts = texts(scene);
    const out: string[] = [];
    for (let i = 0; i < ts.length; i++)
      for (let j = i + 1; j < ts.length; j++) {
        if (!/series-label/.test(`${ts[i].name} ${ts[j].name}`)) continue;
        const a = inkOf(ts[i]);
        const b = inkOf(ts[j]);
        const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
        const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
        if (w > 0 && h > 0 && w * h > 1) out.push(`${ts[i].name} "${ts[i].text}" / ${ts[j].name} "${ts[j].text}"`);
      }
    return out;
  };
  const named = (scene: Scene, re: RegExp) => texts(scene).filter((t) => re.test(t.name ?? ""));
  const lineNames = (scene: Scene) => named(scene, /^combo-series-label-\d+$/);
  const columnNames = (scene: Scene) => named(scene, /^series-label-\d+$/);
  const ticks = (scene: Scene) => named(scene, /^secondary-axis$/);

  /** The sweep's own data shapes. */
  const tenSeries = (c: ChartConfig, rise: (i: number) => number = (i) => i): ChartConfig => ({
    ...c,
    data: {
      ...c.data,
      series: Array.from({ length: 10 }, (_, i) => ({
        ...c.data.series[i % c.data.series.length],
        name: `Series ${i + 1}`,
        values: c.data.series[i % c.data.series.length].values.map((v) => (v == null ? v : v + rise(i))),
      })),
    },
  });
  const longNames = (c: ChartConfig): ChartConfig => ({
    ...c,
    data: { ...c.data, series: c.data.series.map((s) => ({ ...s, name: `${s.name} including allocations` })) },
  });
  const at = (c: ChartConfig, w: number, h: number, fontSize: number): ChartConfig =>
    ({ ...c, width: w, height: h, style: { fontSize } }) as ChartConfig;
  const sample = (kind: ChartKind, extra: Partial<ChartConfig> = {}) =>
    ({ ...sampleConfig(kind), ...extra }) as ChartConfig;

  describe("a pareto on a frame too short for its font (300x60, 18pt)", () => {
    // `pareto` rewrites whatever kind it is on into a combo, which is why the
    // sweep counted this once per kind; two of them are enough to say so.
    it.each<ChartKind>(["stacked", "mekko"])("%s: the line's name and the column's name share no ink", (kind) => {
      const scene = buildChart(at(sample(kind, { pareto: true }), 300, 60, 18));
      expect(lineNames(scene).map((t) => t.text)).toEqual(["Cumulative %"]);
      expect(columnNames(scene).length, "the column's name is still there").toBeGreaterThan(0);
      expect(gutterPairs(scene)).toEqual([]);
    });
  });

  describe("a pareto whose top tick number sits where the name hangs (480x300, 18pt, long series names)", () => {
    it.each<ChartKind>(["cascade", "mekko"])(
      "%s: the name stands clear of the tick numbers and every one is kept",
      (kind) => {
        const cfg = at(longNames(sample(kind, { pareto: true })), 480, 300, 18);
        const scene = buildChart(cfg);
        expect(lineNames(scene).map((t) => t.text)).toEqual(["Cumulative %"]);
        expect(gutterPairs(scene)).toEqual([]);
        // The tick the name hung under is the top of the axis — the 100% the line
        // ends on. Moving the name aside must not cost the axis that number.
        expect(ticks(scene).map((t) => t.text)).toContain("100");
      },
    );
  });

  describe("ten series on a dual-axis combo, three of them lines that end 3pt apart", () => {
    const build = (w: number, h: number, fs: number) =>
      buildChart(at(tenSeries(sample("combo", { secondaryAxis: true })), w, h, fs));

    it.each([
      [160, 120, 10],
      [200, 150, 10],
      [200, 150, 18],
      [480, 300, 18],
    ])("%ix%i at %ipt: every name is drawn and no two share ink", (w, h, fs) => {
      const scene = build(w, h, fs);
      expect(columnNames(scene).map((t) => t.text)).toHaveLength(7);
      expect(
        lineNames(scene)
          .map((t) => t.text)
          .sort(),
      ).toEqual(["Series 3", "Series 6", "Series 9"]);
      expect(gutterPairs(scene)).toEqual([]);
    });

    /**
     * The names that are drawn read top to bottom in the order their lines end. A
     * name is left out before it is drawn out of that order: a name that names
     * someone else's line is worse than none, which is the failure the settle pass
     * orders its labels to prevent.
     */
    const inEndOrder = (scene: Scene) => {
      const endY = (li: number) => {
        const marks = scene.nodes.filter((n) => new RegExp(`^combo-marker-${li}-\\d+$`).test(n.name ?? ""));
        const last = marks[marks.length - 1] as RectNode;
        return last.y + last.h / 2;
      };
      const drawn = lineNames(scene)
        .sort((a, b) => a.y - b.y)
        .map((t) => Number(t.name?.replace("combo-series-label-", "")));
      return { drawn, byEnd: [...drawn].sort((a, b) => endY(a) - endY(b)) };
    };

    it("the lines are named in the order they end, top to bottom", () => {
      const { drawn, byEnd } = inEndOrder(build(200, 150, 18));
      expect(drawn).toHaveLength(3);
      expect(drawn).toEqual(byEnd);
    });

    it.each([
      [120, 90, 10],
      [160, 120, 10],
      [160, 120, 18],
      [200, 150, 10],
      [300, 60, 10],
    ])("%ix%i at %ipt: a name that finds no room is left out, not put out of order", (w, h, fs) => {
      const { drawn, byEnd } = inEndOrder(build(w, h, fs));
      expect(drawn).toEqual(byEnd);
    });

    // The same strip with the lines the other way up, so the FIRST line drawn is the
    // one that ends highest and every later one has to be placed below its name.
    it.each([
      [120, 90, 10],
      [160, 120, 10],
      [160, 120, 18],
      [200, 150, 10],
      [200, 150, 18],
      [300, 60, 10],
    ])("%ix%i at %ipt, lines the other way up: the same order", (w, h, fs) => {
      const cfg = at(
        tenSeries(sample("combo", { secondaryAxis: true }), (i) => 9 - i),
        w,
        h,
        fs,
      );
      const { drawn, byEnd } = inEndOrder(buildChart(cfg));
      expect(drawn.length, "some line is named").toBeGreaterThan(0);
      expect(drawn).toEqual(byEnd);
    });

    it.each([
      [120, 90, 10],
      [160, 120, 18],
    ])("%ix%i at %ipt, where seven column names already fill the strip: nothing is drawn over anything", (w, h, fs) => {
      const scene = build(w, h, fs);
      // No room is no room. A name that cannot stand anywhere in the order its line
      // ends is left out, as the merged gutter leaves out every name it cannot lay
      // out — never drawn across another one — and the column names are not taken
      // with it.
      expect(columnNames(scene)).toHaveLength(7);
      expect(lineNames(scene).length, "at least the line that has room is named").toBeGreaterThan(0);
      expect(gutterPairs(scene)).toEqual([]);
    });

    it("a name with no room at its size is made smaller until it has some (160x120, 18pt)", () => {
      const scene = build(160, 120, 18);
      const [nearest, shrunk] = lineNames(scene).sort((a, b) => a.name!.localeCompare(b.name!));
      // "Series 3" ends lowest and keeps the chart's size; "Series 6" ends just above
      // it and stands in the little room left over the strip, at the size that fits.
      expect([nearest.text, shrunk.text]).toEqual(["Series 3", "Series 6"]);
      expect(shrunk.fontSize).toBeLessThan(nearest.fontSize);
      expect(shrunk.fontSize).toBeGreaterThanOrEqual(5);
    });
  });

  describe("two lines that end a few points apart", () => {
    // Found by drawing 3,000 random dual-axis charts and asking whether the names
    // read in the order the lines end: 210 of them did not once the first line
    // drawn was allowed to be named below one that ends under the second. The
    // first drawn is not the highest here, which is what the ten-series data above
    // never exercises.
    const lines = (cats: string[], endsA: (number | null)[], endsB: (number | null)[]) => [
      { name: "Product", values: cats.map((_, i) => 30 + i * 10) },
      { name: "Target", type: "line", values: endsA },
      { name: "Cumulative %", type: "line", values: endsB },
    ];
    const cats = ["C1", "C2", "C3", "C4", "C5"];
    const chart = (extra: Partial<ChartConfig>, size: [number, number], fs: number): ChartConfig =>
      ({
        kind: "combo",
        width: size[0],
        height: size[1],
        title: "A rather long chart title that names the measure",
        data: { categories: cats, series: lines(cats, [21, 15, 38, 80, null], [43, 71, 44, 38, 76]) },
        decorations: { seriesLabels: true, categoryAxis: true },
        style: { fontSize: fs },
        ...extra,
      }) as ChartConfig;

    it.each([
      ["a 100% base", { combo: { columns: "stacked100" } }, [480, 300], 10],
      ["a secondary axis at 160x120, 24pt", { secondaryAxis: true }, [160, 120], 24],
      ["a secondary axis at 120x90, 8pt", { secondaryAxis: true }, [120, 90], 8],
    ] as [string, Partial<ChartConfig>, [number, number], number][])(
      "%s: the names read in the order the lines end",
      (_label, extra, size, fs) => {
        const scene = buildChart(chart(extra, size, fs));
        const endY = (li: number) => {
          const marks = scene.nodes.filter((n) => new RegExp(`^combo-marker-${li}-\\d+$`).test(n.name ?? ""));
          const last = marks[marks.length - 1] as RectNode;
          return last.y + last.h / 2;
        };
        const drawn = lineNames(scene)
          .sort((a, b) => a.y - b.y)
          .map((t) => Number(t.name?.replace("combo-series-label-", "")));
        expect(drawn.length, "both lines are named").toBe(2);
        expect(drawn).toEqual([...drawn].sort((a, b) => endY(a) - endY(b)));
        expect(gutterPairs(scene)).toEqual([]);
      },
    );
  });

  describe("a footnote's band is reserved for the footnote", () => {
    const chart = (size: [number, number], fs: number): ChartConfig =>
      ({
        kind: "combo",
        width: size[0],
        height: size[1],
        title: "Revenue and margin by segment",
        footnote: "Source: an internal model",
        secondaryAxis: true,
        data: {
          categories: ["C1", "C2", "C3", "C4", "C5"],
          series: [
            { name: "Product", values: [48, 11, 56, 37, 13] },
            { name: "Legacy", values: [87, 50, 12, 64, 29] },
            { name: "Target", type: "line", values: [11, 79, null, 54, 81] },
          ],
        },
        decorations: { seriesLabels: true, categoryAxis: true, valueAxis: true },
        style: { fontSize: fs },
      }) as ChartConfig;

    it("a line's name does not stand where the footnote is (300x60, 24pt)", () => {
      const scene = buildChart(chart([300, 60], 24));
      expect(gutterPairs(scene).filter((p) => p.includes('"Source'))).toEqual([]);
    });

    it("control: with room, the same chart names its line and shares no ink with the footnote", () => {
      const scene = buildChart(chart([480, 300], 10));
      expect(lineNames(scene).map((t) => t.text)).toEqual(["Target"]);
      expect(texts(scene).some((t) => t.name === "footnote")).toBe(true);
      expect(gutterPairs(scene)).toEqual([]);
    });
  });

  describe("controls: a name with clear air keeps the spot it always had", () => {
    it.each([
      [480, 300],
      [960, 540],
    ])("the sample combo with a secondary axis, %ix%i: above its line's last point", (w, h) => {
      const scene = buildChart(at(sample("combo", { secondaryAxis: true }), w, h, 10));
      const [name] = lineNames(scene);
      const marks = scene.nodes.filter((n) => /^combo-marker-0-\d+$/.test(n.name ?? "")) as RectNode[];
      const end = marks[marks.length - 1];
      expect(name.y).toBeCloseTo(end.y + end.h / 2 - 10 * 1.6, 5);
      expect(gutterPairs(scene)).toEqual([]);
    });
  });
});

/* -------------------------------------------------------------------------- */

describe("r1-39: a smoothed line never leaves the plot or draws a value its data does not hold", () => {
  const smooth = (values: number[], extra: Partial<ChartConfig> = {}): ChartConfig =>
    ({
      kind: "line",
      ...DEFAULT_SIZE,
      data: { categories: values.map((_, i) => String.fromCharCode(65 + i)), series: [{ name: "S", values }] },
      decorations: { smooth: true, segmentLabels: false },
      ...extra,
    }) as ChartConfig;
  const curve = (scene: Scene) =>
    nodesOf(scene, "line-0-").filter((n): n is LineNode => n.kind === "line" && /-s\d+$/.test(n.name ?? ""));
  const ys = (c: LineNode[]) => c.flatMap((s) => [s.y1, s.y2]);
  const xs = (c: LineNode[]) => c.flatMap((s) => [s.x1, s.x2]);
  const markers = (scene: Scene) => nodesOf(scene, "marker-0-").filter((n): n is RectNode => n.kind === "rect");

  it("a spike in non-negative data does not dip below the zero baseline (480x300)", () => {
    const scene = buildChart(smooth([0, 0, 100, 0, 0]));
    const base = (scene.nodes.find((n) => n.name === "baseline") as LineNode).y1;
    const c = curve(scene);
    expect(c.length).toBeGreaterThan(20);
    expect(Math.max(...ys(c)), "the curve ran below the zero line").toBeLessThanOrEqual(base + 1e-6);
    expect(Math.max(...ys(c))).toBeLessThanOrEqual(DEFAULT_SIZE.height);
  });

  it("…at a larger canvas too", () => {
    const scene = buildChart(smooth([0, 0, 100, 0, 0], { width: 800, height: 450 }));
    const base = (scene.nodes.find((n) => n.name === "baseline") as LineNode).y1;
    expect(Math.max(...ys(curve(scene)))).toBeLessThanOrEqual(base + 1e-6);
  });

  it("a dip does not rise past the level it dips from", () => {
    const scene = buildChart(smooth([100, 100, 0, 100, 100]));
    const top = Math.min(...markers(scene).map((m) => m.y + m.h / 2));
    expect(Math.min(...ys(curve(scene))), "the curve ran above its highest point").toBeGreaterThanOrEqual(top - 1e-6);
    expect(Math.min(...ys(curve(scene)))).toBeGreaterThanOrEqual(0);
  });

  it("launch-shaped data stays on or above the baseline", () => {
    const scene = buildChart(smooth([0, 0, 0, 120, 15, 10, 8, 8]));
    const base = (scene.nodes.find((n) => n.name === "baseline") as LineNode).y1;
    expect(Math.max(...ys(curve(scene)))).toBeLessThanOrEqual(base + 1e-6);
  });

  it("sideways, a spike does not run left of zero or off the canvas", () => {
    const scene = buildChart(smooth([0, 0, 100, 0, 0], { horizontal: true }));
    const zero = Math.min(...markers(scene).map((m) => m.x + m.w / 2));
    const c = curve(scene);
    expect(c.length).toBeGreaterThan(20);
    expect(Math.min(...xs(c)), "the curve ran left of the zero line").toBeGreaterThanOrEqual(zero - 1e-6);
    expect(Math.min(...xs(c))).toBeGreaterThanOrEqual(0);
    expect(Math.max(...xs(c))).toBeLessThanOrEqual(DEFAULT_SIZE.width);
  });

  it("sideways, a dip does not run past the level it dips from", () => {
    const scene = buildChart(smooth([100, 100, 0, 100, 100], { horizontal: true }));
    const right = Math.max(...markers(scene).map((m) => m.x + m.w / 2));
    expect(Math.max(...xs(curve(scene)))).toBeLessThanOrEqual(right + 1e-6);
  });

  it("a gap splits the curve and each piece keeps to its own points", () => {
    const gapped = smooth([]);
    gapped.data = {
      categories: "ABCDEFGHI".split(""),
      series: [{ name: "S", values: [0, 0, 100, 0, null, 20, 60, 20, 20] }],
    };
    const scene = buildChart(gapped);
    const base = (scene.nodes.find((n) => n.name === "baseline") as LineNode).y1;
    expect(Math.max(...ys(curve(scene)))).toBeLessThanOrEqual(base + 1e-6);
    // The second piece (20, 60, 20, 20) is bounded by ITS points, not the whole series':
    // it may not rise past its own 60, which is far below the first piece's 100.
    const second = curve(scene).filter((s) => Number(s.name!.split("-")[2]) >= 5);
    const own = nodesOf(scene, "marker-0-6")[0] as RectNode;
    expect(second.length).toBeGreaterThan(0);
    expect(Math.min(...ys(second))).toBeGreaterThanOrEqual(own.y + own.h / 2 - 1e-6);
  });

  // CONTROLS: the curve is still the curve.
  it("control: straight data is still drawn as the straight line it is", () => {
    const pts = [10, 20, 30, 40, 50].map((v, c) => ({ x: 100 + 50 * c, y: 300 - v, c }));
    const out = splineSegments(pts, 4);
    expect(out).toHaveLength(16);
    for (const seg of out) {
      const slope = (seg.y2 - seg.y1) / (seg.x2 - seg.x1);
      expect(slope).toBeCloseTo(-10 / 50, 9);
    }
  });

  it("control: a swing that stays inside its points' extent is the Catmull-Rom spline, untouched", () => {
    const y = [100, 80, 90, 60, 70];
    const pts = y.map((v, c) => ({ x: 50 * c, y: v, c }));
    const out = splineSegments(pts, 8);
    // Segment 1 -> 2 at t = 4/8, evaluated by hand: P0..P3 = 100, 80, 90, 60.
    const t = 0.5;
    const hand =
      0.5 *
      (2 * 80 + (-100 + 90) * t + (2 * 100 - 5 * 80 + 4 * 90 - 60) * t * t + (-100 + 3 * 80 - 3 * 90 + 60) * t ** 3);
    const sample = out.find((s) => s.c === 2 && s.k === 4)!;
    expect(sample.y2).toBeCloseTo(hand, 9);
    // …which does swing past a knot (it is not simply flat), so the control means something.
    expect(hand).not.toBeCloseTo((80 + 90) / 2, 3);
  });

  it("control: every knot is still passed through", () => {
    const scene = buildChart(smooth([0, 0, 100, 0, 0]));
    const ends = curve(scene)
      .filter((s) => s.name!.endsWith("-s16"))
      .map((s) => s.y2);
    const marks = markers(scene).map((m) => m.y + m.h / 2);
    expect(ends).toHaveLength(4);
    ends.forEach((e, i) => expect(e).toBeCloseTo(marks[i + 1], 6));
  });
});

/* -------------------------------------------------------------------------- */

describe("r1-40: a bump chart keeps every rank on the chart, whatever the lowest one is", () => {
  const bump = (rows: number[][]): ChartConfig =>
    ({
      kind: "line",
      ...DEFAULT_SIZE,
      data: {
        categories: ["P1", "P2", "P3"].slice(0, rows[0].length),
        series: rows.map((values, i) => ({ name: `E${i + 1}`, values })),
      },
      decorations: { bump: true },
    }) as ChartConfig;
  const segments = (scene: Scene) => nodesOf(scene, "bump-").filter((n): n is LineNode => n.kind === "line");
  const dots = (scene: Scene) =>
    nodesOf(scene, "bump-marker-").filter((n): n is Extract<SceneNode, { kind: "ellipse" }> => n.kind === "ellipse");
  const offCanvas = (scene: Scene) =>
    segments(scene).filter((l) => [l.y1, l.y2].some((y) => y < 0 || y > DEFAULT_SIZE.height));

  it.each([
    [
      "a 0-based ranking",
      [
        [0, 1, 2],
        [1, 0, 1],
        [2, 2, 0],
      ],
    ],
    [
      "one typo'd 0 among 1-based ranks",
      [
        [1, 2, 3],
        [2, 1, 3],
        [3, 3, 1],
        [4, 4, 2],
        [5, 5, 0],
      ],
    ],
    [
      "a negative rank",
      [
        [1, -3, 2],
        [2, 1, 1],
      ],
    ],
    [
      "line data with negatives, bump switched on",
      [
        [40, 0, 60],
        [-10, 20, 30],
      ],
    ],
  ] as [string, number[][]][])("%s: no line leaves the canvas and every marker is drawn", (_label, rows) => {
    const scene = buildChart(bump(rows));
    expect(
      offCanvas(scene).map((l) => `${l.name}: y ${l.y1.toFixed(0)} -> ${l.y2.toFixed(0)}`),
      "ink above or below the chart",
    ).toEqual([]);
    expect(dots(scene)).toHaveLength(rows.flat().length);
  });

  it("a 0-based ranking reads top-down from 0, each rank on its own row", () => {
    const scene = buildChart(
      bump([
        [0, 1, 2],
        [1, 0, 1],
        [2, 2, 0],
      ]),
    );
    const rowOf = (name: string) => (dots(scene).find((d) => d.name === name) as { cy: number }).cy;
    // E1 is ranked 0 in P1, E2 is ranked 0 in P2, E3 is ranked 0 in P3: the same top row.
    const top = rowOf("bump-marker-0-0");
    expect(rowOf("bump-marker-1-1")).toBeCloseTo(top, 9);
    expect(rowOf("bump-marker-2-2")).toBeCloseTo(top, 9);
    const mid = rowOf("bump-marker-0-1"); // rank 1
    const bottom = rowOf("bump-marker-0-2"); // rank 2
    expect(top).toBeLessThan(mid);
    expect(mid - top).toBeCloseTo(bottom - mid, 9);
  });

  it("a 0 and a 1 are different ranks, not the same row", () => {
    const scene = buildChart(
      bump([
        [0, 1],
        [1, 0],
      ]),
    );
    const ys = new Set(dots(scene).map((d) => d.cy.toFixed(3)));
    expect(ys.size).toBe(2);
  });

  // CONTROLS
  it("control: ranks that start at 1 are drawn exactly where they always were", () => {
    const scene = buildChart(
      bump([
        [1, 2, 3],
        [2, 1, 3],
        [3, 3, 1],
      ]),
    );
    const ys = [...new Set(dots(scene).map((d) => d.cy))].sort((a, b) => a - b);
    expect(ys).toHaveLength(3);
    // Rank 1 on the plot's top edge — which starts where the period headers end —
    // and the rows a constant pitch apart.
    const header = scene.nodes.find((n) => n.name === "period-0") as TextNode;
    expect(ys[0]).toBeCloseTo(header.y + header.h, 9);
    expect(ys[1] - ys[0]).toBeCloseTo(ys[2] - ys[1], 9);
  });

  it("control: a ranking that is all 1s still sits on the plot's midline", () => {
    const scene = buildChart(
      bump([
        [1, 1],
        [1, 1],
      ]),
    );
    expect(new Set(dots(scene).map((d) => d.cy)).size).toBe(1);
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-38: the charts with no value axis still honour the scale they are given", () => {
  const kpis = (extra: Partial<ChartConfig> = {}, decorations: object = { sparkline: true }): ChartConfig =>
    ({
      kind: "line",
      width: 480,
      height: 200,
      data: {
        categories: ["Q1", "Q2", "Q3", "Q4"],
        series: [
          { name: "KPI 1", values: [10, 12, 11, 14] },
          { name: "KPI 2", values: [20, 24, 22, 28] },
          { name: "KPI 3", values: [30, 36, 33, 42] },
        ],
      },
      decorations,
      multiples: {},
      ...extra,
    }) as ChartConfig;
  const one = (extra: Partial<ChartConfig> = {}): ChartConfig =>
    ({
      kind: "line",
      width: 480,
      height: 200,
      data: { categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "KPI", values: [10, 12, 11, 14] }] },
      decorations: { sparkline: true },
      ...extra,
    }) as ChartConfig;

  const lines = (scene: Scene, prefix: string) =>
    nodesOf(scene, prefix).filter((n): n is LineNode => n.kind === "line");
  /** How far the drawn line travels vertically. */
  const travel = (ls: LineNode[]) =>
    Math.max(...ls.flatMap((l) => [l.y1, l.y2])) - Math.min(...ls.flatMap((l) => [l.y1, l.y2]));

  describe("sparkline", () => {
    it("a table of sparklines put on one scale by the author draws their real proportions", () => {
      const scene = buildChart(kpis({ scale: { min: 0, max: 100 } }));
      const t = [0, 1, 2].map((i) => travel(lines(scene, `p${i}-spark-0`)));
      // KPI 1, 2 and 3 move 4, 8 and 12 of the 100 the scale spans, and a table
      // of KPIs read side by side on one scale is only honest if the lines say so.
      // Every panel drew the whole of its plot (136pt) whatever the scale said.
      t.forEach((x, i) =>
        expect(x, `panel travel ${t.map((y) => y.toFixed(1)).join(" / ")}`).toBeCloseTo(((i + 1) * 4 * 136) / 100, 6),
      );
    });

    it("control: rows of unlike units, with no scale pinned, still scale to themselves", () => {
      // The shipped "KPI sparklines" slide: a churn row beside a signups row
      // would be all but a flat line on one scale. Only an explicit scale joins them.
      const scene = buildChart(kpis());
      const t = [0, 1, 2].map((i) => travel(lines(scene, `p${i}-spark-0`)));
      t.forEach((x) => expect(x).toBeCloseTo(136, 6));
    });

    it("a scale the author pinned is drawn to", () => {
      const free = travel(lines(buildChart(one()), "spark-0"));
      const pinned = travel(lines(buildChart(one({ scale: { min: 0, max: 100 } })), "spark-0"));
      // The data moves 4 of the 100 the scale spans.
      expect(pinned / free).toBeCloseTo(4 / 100, 6);
    });

    it("a value past a pinned end sits on the end of the plot, not off the chart", () => {
      const scene = buildChart(one({ scale: { min: 0, max: 10 } }));
      const ys = lines(scene, "spark-0").flatMap((l) => [l.y1, l.y2]);
      const padY = Math.max(2, 200 * 0.16);
      expect(Math.min(...ys)).toBeGreaterThanOrEqual(padY - 1e-9);
      expect(Math.max(...ys)).toBeLessThanOrEqual(200 - padY + 1e-9);
    });

    it("control: a pin the data cannot live in is ignored, as valueScale ignores one", () => {
      const free = travel(lines(buildChart(one()), "spark-0"));
      expect(travel(lines(buildChart(one({ scale: { min: 20 } })), "spark-0"))).toBeCloseTo(free, 9);
    });

    it("control: with no scale a sparkline spans its own data, as it always did", () => {
      const ls = lines(buildChart(one()), "spark-0");
      expect(travel(ls)).toBeCloseTo(200 - 2 * Math.max(2, 200 * 0.16), 9);
    });
  });

  describe("slope (the same range, taken the same way)", () => {
    const slope = (extra: Partial<ChartConfig> = {}): ChartConfig =>
      ({
        kind: "line",
        width: 480,
        height: 300,
        data: {
          categories: ["2023", "2024"],
          series: [
            { name: "A", values: [10, 30] },
            { name: "B", values: [20, 40] },
          ],
        },
        decorations: { slope: true },
        ...extra,
      }) as ChartConfig;

    it("a scale the author pinned is drawn to", () => {
      const free = travel([...lines(buildChart(slope()), "line-0-"), ...lines(buildChart(slope()), "line-1-")]);
      const pinnedScene = buildChart(slope({ scale: { min: 0, max: 100 } }));
      const pinned = travel([...lines(pinnedScene, "line-0-"), ...lines(pinnedScene, "line-1-")]);
      // The data spans 10..40, 30 of the 100.
      expect(pinned / free).toBeCloseTo(30 / 100, 6);
    });

    it("panels of a slope table share the scale buildMultiples pins them to, like any other line panel", () => {
      const scene = buildChart(kpis({}, { slope: true }));
      const t = [0, 1, 2].map((i) => travel(lines(scene, `p${i}-line-0-`)));
      // KPI 1, 2 and 3 move 4, 8 and 12 of the same range.
      expect(t[1] / t[0], `panel travel ${t.map((x) => x.toFixed(1)).join(" / ")}`).toBeCloseTo(2, 6);
      expect(t[2] / t[0]).toBeCloseTo(3, 6);
    });

    it("control: with no scale a slope spans its own data, top to bottom of its rails", () => {
      const scene = buildChart(slope());
      const ys = [...lines(scene, "line-0-"), ...lines(scene, "line-1-")].flatMap((l) => [l.y1, l.y2]);
      const rails = lines(scene, "slope-rail-");
      const h = rails[0].y2 - rails[0].y1;
      expect(Math.max(...ys) - Math.min(...ys)).toBeCloseTo(h * (1 - 0.16), 9);
    });
  });
});
