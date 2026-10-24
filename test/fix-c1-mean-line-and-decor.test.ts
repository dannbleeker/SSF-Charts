import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { resolveLabelCollisions, tightBox } from "../src/core/collide";
import { layoutLine } from "../src/core/layout/line";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, Scene, SceneNode, TextNode } from "../src/core/scene";

/**
 * Seven findings in one corner of the engine — the mean value line, the
 * difference and CAGR arrows, the callouts and the variance tier — that share
 * one habit: reading `columnValue`, `seriesLevels` or a series INDEX without
 * asking whether it names a measurement, or the series it was written against.
 *
 * Each describe block below names its finding and carries the control that
 * keeps it honest: the test that the fix did not simply delete the feature.
 */

type Series = { name: string; values: (number | null)[]; type?: string };

const chart = (kind: string, series: Series[], decorations: Record<string, unknown>, extra: object = {}): ChartConfig =>
  ({
    kind,
    width: 480,
    height: 300,
    data: { categories: ["A", "B", "C"].slice(0, series[0].values.length), series },
    decorations,
    ...extra,
  }) as unknown as ChartConfig;

const node = (scene: Scene, name: string): SceneNode | undefined => scene.nodes.find((n) => n.name === name);
const text = (scene: Scene, name: string) => (node(scene, name) as TextNode | undefined)?.text;
const line = (scene: Scene, name: string) => node(scene, name) as unknown as { y1: number; y2: number } | undefined;
const names = (scene: Scene, prefix: string) => scene.nodes.filter((n) => n.name?.startsWith(prefix));

/* -------------------------------------------------------------------------- */

describe("r2-0: the mean value line averages what was measured on EVERY kind that draws one", () => {
  const blank = [100, null, 200];
  const mean = { valueLines: [{ mode: "mean" }] };
  const at150 = { valueLines: [{ mode: "value", value: 150 }] };

  it.each(["line", "area", "clustered", "stacked"])("%s: [100, blank, 200] has a mean of 150", (kind) => {
    const scene = buildChart(chart(kind, [{ name: "S", values: blank }], mean));
    expect(
      text(scene, "value-line-label-0"),
      `a blank category was averaged in as a zero on a ${kind} chart: the mean of the two numbers supplied is 150`,
    ).toBe("Ø 150");
    // …and the line is DRAWN there, not just captioned there.
    const fixed = buildChart(chart(kind, [{ name: "S", values: blank }], at150));
    expect(line(scene, "value-line-0")!.y1).toBeCloseTo(line(fixed, "value-line-0")!.y1, 6);
  });

  it("a year-to-date line (six actual months, six blank) is not halved", () => {
    const ytd = [100, 110, 105, 120, 115, 110, null, null, null, null, null, null];
    const cats = Array.from({ length: 12 }, (_, i) => `M${i + 1}`);
    const build = (kind: string) =>
      buildChart({
        kind,
        width: 520,
        height: 320,
        data: { categories: cats, series: [{ name: "S", values: ytd }] },
        decorations: mean,
      } as unknown as ChartConfig);
    expect(text(build("line"), "value-line-label-0")).toBe("Ø 110");
    expect(text(build("line"), "value-line-label-0")).toBe(text(build("clustered"), "value-line-label-0"));
  });

  it("the sparkline and slope layouts, which build columnValue the same way, agree", () => {
    for (const mode of ["sparkline", "slope"]) {
      const scene = buildChart(chart("line", [{ name: "S", values: blank }], { ...mean, [mode]: true }));
      expect(text(scene, "value-line-label-0"), mode).toBe("Ø 150");
    }
  });

  it("an area counts a category as measured wherever ANY series has a cell, like a stack", () => {
    const series = [
      { name: "S1", values: blank },
      { name: "S2", values: [10, 20, 30] },
    ];
    const area = buildChart(chart("area", series, mean));
    const stacked = buildChart(chart("stacked", series, mean));
    // (110 + 20 + 230) / 3 — the blank in S1 is a zero in the sum, not a missing category.
    expect(text(area, "value-line-label-0")).toBe("Ø 120");
    expect(text(area, "value-line-label-0")).toBe(text(stacked, "value-line-label-0"));
  });

  it("a line's mean is its FIRST series' mean, so only that series' blanks are left out", () => {
    const scene = buildChart(
      chart(
        "line",
        [
          { name: "S1", values: blank },
          { name: "S2", values: [10, 20, 30] },
        ],
        mean,
      ),
    );
    expect(text(scene, "value-line-label-0")).toBe("Ø 150");
  });

  it("every anchor the line family builds with `?? 0` says which categories were measured", () => {
    const data = {
      categories: ["A", "B", "C"],
      series: [{ name: "S", values: [100, null, 200] as (number | null)[] }],
    };
    const anchors = (over: Partial<ChartConfig>, decor: object = {}) =>
      layoutLine({ kind: "line", width: 480, height: 300, data, ...over } as ChartConfig, DEFAULT_STYLE, {
        ...DEFAULT_DECOR,
        ...decor,
      }).anchors.columnHasData;
    expect(anchors({}), "line").toEqual([true, false, true]);
    expect(anchors({ kind: "area" }), "area").toEqual([true, false, true]);
    expect(anchors({}, { sparkline: true }), "sparkline").toEqual([true, false, true]);
    expect(anchors({}, { slope: true }), "slope").toEqual([true, false, true]);
    expect(anchors({}, { bump: true }), "bump").toEqual([true, false, true]);
    expect(anchors({ horizontal: true }), "horizontal line").toEqual([true, false, true]);
  });

  it("a difference arrow on a BUMP chart, which draws no value map, still refuses a blank end", () => {
    const scene = buildChart(
      chart("line", [{ name: "S", values: [3, 2, null] }], { bump: true, difference: { from: 0, to: 2 } }),
    );
    expect(text(scene, "diff-label")).toBeUndefined();
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-7: an arrow to a category nobody filled in is not drawn", () => {
  const diff = (kind: string, values: (number | null)[], extra: Record<string, unknown> = {}) =>
    buildChart(chart(kind, [{ name: "S", values }], { difference: { from: 0, to: 2 }, ...extra }));
  const arrowNodes = (scene: Scene) =>
    ["diff-label", "diff-line", "diff-level", "diff-head"].filter((n) => node(scene, n));

  it.each(["clustered", "stacked", "line"])("%s: [50, 60, blank] draws no difference arrow", (kind) => {
    expect(arrowNodes(diff(kind, [50, 60, null])), `an arrow was drawn to the blank end of a ${kind} chart`).toEqual(
      [],
    );
  });

  it("…and the same for a blank START, which used to print the other end as an absolute figure", () => {
    expect(arrowNodes(diff("stacked", [null, 60, 50], { totals: true }))).toEqual([]);
  });

  it("a MEASURED zero is still a zero: -100% is the right answer for it", () => {
    expect(text(diff("clustered", [50, 60, 0]), "diff-label")).toBe("-100%");
    expect(text(diff("clustered", [50, 60, 40]), "diff-label")).toBe("-20%");
  });

  it("a series-level arrow is withheld when THAT series is blank at an end", () => {
    const series = [
      { name: "A", values: [50, 60, 70] },
      { name: "B", values: [20, 30, null] },
    ];
    const level = (si: number) =>
      buildChart(chart("clustered", series, { difference: { from: 0, to: 2, series: si } }));
    expect(text(level(0), "diff-label")).toBe("+40%");
    expect(text(level(1), "diff-label")).toBeUndefined();
  });

  it("a stacked category blank in EVERY series has no level to point at", () => {
    const scene = buildChart(
      chart(
        "stacked",
        [
          { name: "A", values: [30, 40, null] },
          { name: "B", values: [20, 20, null] },
        ],
        { difference: { from: 0, to: 2, series: 0 } },
      ),
    );
    expect(text(scene, "diff-label")).toBeUndefined();
  });

  it("an arrow anchored at a value line only needs its END measured", () => {
    const mean = { valueLines: [{ mode: "mean" }] };
    const anch = (values: (number | null)[]) =>
      buildChart(
        chart("clustered", [{ name: "S", values }], { ...mean, difference: { from: 0, to: 2, fromValueLine: 0 } }),
      );
    // The start category is only where the dashed baseline begins — blank is fine.
    expect(text(anch([null, 100, 200]), "diff-label")).toBeTypeOf("string");
    expect(text(anch([100, 200, null]), "diff-label")).toBeUndefined();
  });

  it("the CAGR arrow is not drawn from or to the baseline of a blank category", () => {
    const cagrNodes = (values: (number | null)[]) =>
      ["cagr-line", "cagr-head", "cagr-label"].filter((n) =>
        node(buildChart(chart("stacked", [{ name: "S", values }], { cagr: { from: 0, to: 2 } })), n),
      );
    expect(cagrNodes([10, 20, null]), "blank end").toEqual([]);
    expect(cagrNodes([null, 20, 30]), "blank start").toEqual([]);
    // A measured zero end has no rate either, and says so with an arrow that
    // really does reach the baseline — that is the honest picture of it.
    expect(cagrNodes([10, 20, 0])).toEqual(["cagr-line", "cagr-head", "cagr-label"]);
    expect(cagrNodes([10, 20, 40])).toEqual(["cagr-line", "cagr-head", "cagr-label"]);
  });

  it("a series CAGR is withheld when that series is blank at an end", () => {
    const series = [
      { name: "A", values: [10, 20, 30] },
      { name: "B", values: [5, 6, null] },
    ];
    const cagr = (si: number) => buildChart(chart("clustered", series, { cagr: { from: 0, to: 2, series: si } }));
    expect(node(cagr(1), "cagr-line")).toBeUndefined();
    expect(text(cagr(0), "cagr-label")).toBe("+73.2% p.a.");
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-6: a series-qualified decoration names the series it was written for", () => {
  const rev = { name: "Revenue", values: [100, 150, 200] as (number | null)[] };
  const cost = { name: "Cost", values: [80, 80, 80] as (number | null)[] };

  it("line: series 1 is Cost's level (flat), not Revenue's", () => {
    const level = (series?: number) =>
      buildChart(chart("line", [rev, cost], { difference: { from: 0, to: 2, ...(series == null ? {} : { series }) } }));
    expect(text(level(1), "diff-label")).toBe("0%");
    // The default and series 0 are unchanged.
    expect(text(level(0), "diff-label")).toBe("+100%");
    expect(text(level(), "diff-label")).toBe("+100%");
  });

  it("line: a callout on series 1 points at Cost's marker", () => {
    const scene = buildChart(chart("line", [rev, cost], { callouts: [{ text: "flat cost", category: 2, series: 1 }] }));
    const marker = node(scene, "marker-1-2") as RectNode;
    const tail = line(scene, "callout-tail-0")!;
    expect(marker, "re-point this test: no marker-1-2").toBeTruthy();
    expect(tail.y2).toBeCloseTo(marker.y + marker.h / 2 - 2, 0);
  });

  it("area: a series below the top one reports ITS level, not the stack total", () => {
    const level = (series?: number) =>
      buildChart(chart("area", [rev, cost], { difference: { from: 0, to: 2, ...(series == null ? {} : { series }) } }));
    expect(text(level(0), "diff-label")).toBe("+100%");
    expect(text(level(1), "diff-label")).toBe("+56%");
    expect(text(level(), "diff-label")).toBe("+56%");
  });

  it("combo: a LINE series' index reads the line, not the column total", () => {
    const combo = (series: number) =>
      buildChart(chart("combo", [rev, { ...cost, type: "line" }], { difference: { from: 0, to: 2, series } }));
    expect(text(combo(1), "diff-label")).toBe("0%");
    expect(text(combo(0), "diff-label")).toBe("+100%");
  });

  it("combo: a line written FIRST does not shift the column indices under the decoration", () => {
    const series = [
      { name: "Cost", values: [80, 80, 80], type: "line" },
      { name: "ColA", values: [100, 150, 200] },
      { name: "ColB", values: [10, 10, 10] },
    ];
    const combo = (si: number) => buildChart(chart("combo", series, { difference: { from: 0, to: 2, series: si } }));
    expect(text(combo(0), "diff-label"), "the line").toBe("0%");
    expect(text(combo(1), "diff-label"), "ColA alone").toBe("+100%");
    expect(text(combo(2), "diff-label"), "ColA + ColB").toBe("+91%");
  });

  it("combo: a line on its OWN axis has no place on the columns' scale, so no level arrow is drawn", () => {
    const scene = buildChart(
      chart(
        "combo",
        [rev, { ...cost, type: "line" }],
        { difference: { from: 0, to: 2, series: 1 } },
        { secondaryAxis: true },
      ),
    );
    expect(text(scene, "diff-label"), "an arrow was measured on the wrong scale").toBeUndefined();
    // A callout is a comment, not a claim: it keeps its text and points at the column.
    const callout = buildChart(
      chart(
        "combo",
        [rev, { ...cost, type: "line" }],
        { callouts: [{ text: "n", category: 2, series: 1 }] },
        { secondaryAxis: true },
      ),
    );
    const plain = buildChart(
      chart(
        "combo",
        [rev, { ...cost, type: "line" }],
        { callouts: [{ text: "n", category: 2 }] },
        { secondaryAxis: true },
      ),
    );
    expect(line(callout, "callout-tail-0")!.y2).toBeCloseTo(line(plain, "callout-tail-0")!.y2, 6);
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-9: a series CAGR on a stack is anchored on that series' segment", () => {
  const base = { name: "Base", values: [100, 100, 100] };
  const added = { name: "New", values: [20, 30, 40] };
  const inside = (r: RectNode, x: number, y: number) => x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h;

  it.each(["stacked", "area"])("%s: the arrow is not drawn through the Base series", (kind) => {
    const scene = buildChart(chart(kind, [base, added], { cagr: { from: 0, to: 2, series: 1 } }));
    const arrow = line(scene, "cagr-line") as unknown as { x1: number; y1: number; x2: number; y2: number };
    const label = node(scene, "cagr-label") as TextNode;
    expect(arrow, "no CAGR arrow was drawn — re-point this test").toBeTruthy();
    const baseInk = scene.nodes.filter((n): n is RectNode => n.kind === "rect" && /^(seg|area)-0-/.test(n.name ?? ""));
    expect(baseInk.length).toBeGreaterThan(0);
    const hits = baseInk.filter(
      (r) =>
        inside(r, arrow.x1, arrow.y1) ||
        inside(r, arrow.x2, arrow.y2) ||
        inside(r, label.x + label.w / 2, label.y + label.h / 2),
    );
    expect(
      hits.map((r) => r.name),
      "the arrow, its head or its caption sit inside the series beneath the one it is about",
    ).toEqual([]);
  });

  it("…and it sits where the difference arrow for the same series sits, lifted clear of it", () => {
    const scene = buildChart(
      chart("stacked", [base, added], {
        cagr: { from: 0, to: 2, series: 1 },
        difference: { from: 0, to: 2, series: 1 },
      }),
    );
    // Both are anchored on the New segment's top; the CAGR one is then lifted.
    expect(line(scene, "cagr-line")!.y1).toBeLessThanOrEqual(line(scene, "diff-line")!.y1);
    expect(line(scene, "cagr-line")!.y2).toBeLessThanOrEqual(line(scene, "diff-line")!.y2);
  });

  it("a clustered series is unchanged: its level IS its own value", () => {
    const withSeries = buildChart(chart("clustered", [base, added], { cagr: { from: 0, to: 2, series: 1 } }));
    const own = line(withSeries, "cagr-line")!;
    // The New bars are 20 and 40 against a 100-high Base: the arrow starts at
    // the SHORT bar, far below the Base tops, and rises by the same amount.
    const baseTop = (node(withSeries, "seg-0-0") as RectNode).y;
    expect(own.y1).toBeGreaterThan(baseTop);
    expect(own.y2).toBeLessThan(own.y1);
  });
});

/* -------------------------------------------------------------------------- */

describe("r1-16: a decoration's series index survives the transforms that rewrite the series list", () => {
  const plan = { name: "Plan", values: [10, 20, 30, 40] };
  const actual = { name: "Actual", values: [10, 14, 22, 26] };
  const target = { name: "Target", values: [12, 18, 24, 30] };
  const cats4 = (series: Series[], decorations: Record<string, unknown>, extra: object = {}) =>
    ({
      kind: "clustered",
      width: 480,
      height: 300,
      data: { categories: ["A", "B", "C", "D"], series },
      decorations,
      ...extra,
    }) as unknown as ChartConfig;

  /** The decorations under test, written against a list whose "Actual" is at `idx`. */
  const decorated = (idx: number) => ({
    cagr: { from: 0, to: 3, series: idx },
    difference: { from: 0, to: 3, series: idx },
    variance: { actual: idx, reference: idx - 1 },
    callouts: [{ text: "c", category: 3, series: idx }],
  });
  const read = (scene: Scene) => ({
    cagr: text(scene, "cagr-label"),
    diff: text(scene, "diff-label"),
    variance: names(scene, "variance-label").map((n) => (n as TextNode).text),
    callout: line(scene, "callout-tail-0")?.y2,
  });

  it("a Target row typed ABOVE the series names the same series as one typed below", () => {
    const last = read(buildChart(cats4([plan, actual, target], decorated(1))));
    const first = read(buildChart(cats4([target, plan, actual], decorated(2))));
    expect(last.cagr, "control: the Target-last chart must produce a CAGR").toMatch(/p\.a\./);
    expect(last.variance.length, "control: the variance tier must have labels").toBeGreaterThan(0);
    expect(first).toEqual(last);
  });

  it("an Error row above does the same for a stacked callout's level", () => {
    const err = { name: "Error", values: [1, 1, 1, 1] };
    const build = (series: Series[], idx: number) =>
      buildChart({
        ...cats4(series, { callouts: [{ text: "c", category: 3, series: idx }] }),
        kind: "stacked",
      } as ChartConfig);
    // PLAN's level, not the top of the stack: an index that slid onto Actual would
    // land on the stack top, which is also where an unqualified callout points.
    const lastTail = line(build([plan, actual, err], 0), "callout-tail-0")!.y2;
    const firstTail = line(build([err, plan, actual], 1), "callout-tail-0")!.y2;
    const stackTop = line(build([plan, actual, err], 1), "callout-tail-0")!.y2;
    expect(lastTail, "control: Plan's level must differ from the stack top").toBeGreaterThan(stackTop + 1);
    expect(firstTail).toBeCloseTo(lastTail, 6);
  });

  it("a series folded into Other no longer names a series; the survivors that moved up still do", () => {
    const series = [
      { name: "Small1", values: [1, 1, 1, 1] },
      { name: "BigA", values: [10, 20, 30, 40] },
      { name: "Small2", values: [1, 1, 1, 1] },
      { name: "BigB", values: [100, 110, 120, 130] },
    ];
    // Stacked, so that a CAGR of the TOTALS and a CAGR of one series are different numbers.
    const other = { kind: "stacked", otherBucket: { max: 3 } };
    const bigB = buildChart(cats4(series, { cagr: { from: 0, to: 3, series: 3 } }, other));
    const alone = buildChart(cats4([series[3]], { cagr: { from: 0, to: 3, series: 0 } }));
    expect(text(bigB, "cagr-label"), "BigB moved from index 3 to 1").toBe(text(alone, "cagr-label"));
    // Small1 was folded away: it must not come to name whichever series slid into slot 0.
    const folded = buildChart(cats4(series, { cagr: { from: 0, to: 3, series: 0 } }, other));
    const totals = buildChart(cats4(series, { cagr: { from: 0, to: 3 } }, other));
    expect(text(folded, "cagr-label")).toBe(text(totals, "cagr-label"));
    expect(text(folded, "cagr-label")).not.toBe(text(alone, "cagr-label"));
  });

  it("a pareto keeps only the bar: index 1 must not come to name the cumulative line it adds", () => {
    const pareto = (series?: number) =>
      buildChart(
        cats4(
          [
            { name: "Defects", values: [40, 30, 20, 10] },
            { name: "Cost", values: [1, 2, 3, 4] },
          ],
          { difference: { from: 0, to: 2, ...(series == null ? {} : { series }) } },
          { pareto: true },
        ),
      );
    expect(text(pareto(0), "diff-label"), "the bar itself is still addressable").toBeTypeOf("string");
    expect(text(pareto(1), "diff-label"), "Cost was dropped, so the arrow compares totals").toBe(
      text(pareto(), "diff-label"),
    );
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-36: a flipped point label stays within reach of its marker", () => {
  const label = (over: Partial<TextNode>): TextNode => ({
    kind: "text",
    x: 0,
    y: 0,
    w: 60,
    h: 14,
    text: "66",
    fontSize: 10,
    color: "#000",
    align: "center",
    valign: "bottom",
    name: "combo-label-0-1",
    ...over,
  });
  /** A column of immovable text from `from` down to `to`, so that nothing there is clear. */
  const wall = (from: number, to: number): TextNode[] => {
    const out: TextNode[] = [];
    for (let y = from; y < to; y += 8)
      out.push(label({ y, h: 12, text: "WWWWWWWW", valign: "middle", name: `category-${y}` }));
    return out;
  };

  it("does not park a point label several em below the marker it names", () => {
    // The label starts at y=103 (tight box 104.5..117, so its marker is at 117).
    // Everything is taken down to y=140 and free below that: the first clear spot
    // is ~38pt of travel away and its top lands 26pt under the marker.
    const point = label({ y: 103 });
    resolveLabelCollisions([...wall(60, 140), point], 300);
    expect(point.y, "the label was parked past the reach of its own marker").toBe(103);
  });

  it("still flips to the other side of its marker when that is within reach", () => {
    const point = label({ y: 103 });
    resolveLabelCollisions([...wall(60, 120), point], 300);
    expect(point.y).toBeGreaterThan(103);
    // Its top is no more than two em under the marker.
    expect(tightBox(point).y - 117).toBeLessThanOrEqual(2 * point.fontSize);
  });

  it("leaves the CAGR caption's flip alone: its mark is a sloped arrow this pass cannot see", () => {
    const caption = label({ y: 103, text: "+9% p.a.", name: "cagr-label" });
    resolveLabelCollisions([...wall(60, 140), caption], 300);
    expect(caption.y, "the caption lost the flip it has always had").toBeGreaterThan(103 + 2 * caption.fontSize);
  });

  it("combo: every point label drawn sits within two em of its own marker", () => {
    // The report's chart: the CAGR caption flips under the 'C1' marker first, and
    // the point label '66' used to be pushed past it into the column beneath.
    const scene = buildChart({
      kind: "combo",
      width: 300,
      height: 200,
      title: "Revenue and margin",
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "Cols", values: [38, 68, 80, 3] },
          { name: "Cols2", values: [18, 22, 24, 36] },
          { name: "Line", values: [20, 66, 76, 57], type: "line" },
        ],
      },
      decorations: { segmentLabels: true, totals: true, seriesLabels: true, cagr: { from: 0, to: 3 } },
    } as unknown as ChartConfig);
    const labels = scene.nodes.filter((n): n is TextNode => n.kind === "text" && /^combo-label-/.test(n.name ?? ""));
    expect(labels.length, "re-point this test: no point label was drawn").toBeGreaterThan(0);
    for (const l of labels) {
      const marker = node(scene, l.name!.replace("combo-label-", "combo-marker-")) as RectNode;
      const gap = tightBox(l).y - (marker.y + marker.h);
      expect(
        gap,
        `${l.name} "${l.text}" sits ${(gap / l.fontSize).toFixed(1)} em under ${marker.name}`,
      ).toBeLessThanOrEqual(2 * l.fontSize);
    }
  });
});

/* -------------------------------------------------------------------------- */

describe("r2-37: a variance delta carries ONE sign when forceSign is on", () => {
  const variance = (decorations: Record<string, unknown>, numberFormat?: object) =>
    buildChart(
      chart(
        "clustered",
        [
          { name: "PY", values: [100, 100, 100] },
          { name: "AC", values: [120, 90, 100] },
        ],
        { variance: { actual: 1, reference: 0 }, ...decorations },
        numberFormat ? { numberFormat } : {},
      ),
    );
  const labels = (scene: Scene) => names(scene, "variance-label").map((n) => (n as TextNode).text);

  it("does not double the plus", () => {
    expect(labels(variance({}, { decimals: "auto", forceSign: true }))).toEqual(["+20", "-10", "+0"]);
  });

  it("…with a suffix either", () => {
    expect(labels(variance({}, { decimals: "auto", forceSign: true, suffix: "€m" }))).toEqual([
      "+20€m",
      "-10€m",
      "+0€m",
    ]);
  });

  it("is byte-identical to what it printed without forceSign", () => {
    expect(labels(variance({}, { decimals: "auto", forceSign: true }))).toEqual(labels(variance({})));
  });

  it("percent mode, which never reached formatNumber, is unchanged", () => {
    expect(
      labels(
        variance({ variance: { actual: 1, reference: 0, mode: "percent" } }, { decimals: "auto", forceSign: true }),
      ),
    ).toEqual(["+20%", "-10%", "+0%"]);
  });
});
