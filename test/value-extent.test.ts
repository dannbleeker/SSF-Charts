import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart, valueExtent } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { alphaOf } from "../src/core/color";
import { textWidth } from "../src/core/scene";
import type { LineNode, RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Cross-kind value-extent / auto-scale / layout-indexing invariants. */

/** Helper: the vertical span of every rect in a scene. */
const rectSpan = (scene: { nodes: any[] }) => {
  const ys = scene.nodes.flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : []));
  return { top: Math.min(...ys), bottom: Math.max(...ys) };
};

describe("degenerate inputs do not throw", () => {
  it("boxplot with an empty series", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
  it("area chart with a single empty category set", () => {
    const cfg: ChartConfig = {
      kind: "area",
      ...DEFAULT_SIZE,
      decorations: { seriesLabels: true },
      data: { categories: [], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
});

describe("value extents and auto-scales", () => {
  it("area extent honours the negative stack, like stacked (was floored at 0)", () => {
    const data = {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "A", values: [10, -40, 20] },
        { name: "B", values: [5, -30, 10] },
      ],
    };
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual({ min: -70, max: 30 });
    // identical data must give the same extent as the stacked sibling
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual(
      valueExtent({ kind: "stacked", ...DEFAULT_SIZE, data } as ChartConfig),
    );
  });

  it("negative small-multiples areas stay inside the scene", () => {
    const scene = buildChart({
      kind: "area",
      width: 480,
      height: 300,
      multiples: {},
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "A", values: [10, -40, 20] },
          { name: "B", values: [5, -30, 10] },
        ],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1);
    expect(bottom).toBeLessThanOrEqual(301); // was ~775 — rendered far below the canvas
  });

  it("a target below the data range widens the scale down, not just up", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Actual", values: [10, 20, 30] },
          { name: "Target", values: [-50, -50, -50] },
        ],
      },
    } as ChartConfig);
    const ticks = scene.nodes.filter((n: any) => n.kind === "line" && n.name?.startsWith("target-"));
    expect(ticks.length).toBe(3);
    for (const t of ticks as any[]) {
      expect(t.y1).toBeGreaterThanOrEqual(0);
      expect(t.y1).toBeLessThanOrEqual(300); // target tick used to land outside the plot
    }
  });

  it("a multi-series waterfall combo keeps its columns on canvas", () => {
    const scene = buildChart({
      kind: "combo",
      ...DEFAULT_SIZE,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "W1", values: [10, 10, 10] },
          { name: "W2", values: [10, 10, 10] },
          { name: "L", type: "line", values: [40, 40, 40] },
        ],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1); // was -129 — stacked peak understated
    expect(bottom).toBeLessThanOrEqual(scene.height + 1);
  });

  it("small-multiples panels share one category order under categorySort", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      multiples: { columns: 2 },
      categorySort: "descending",
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "S1", values: [1, 9, 5] },
          { name: "S2", values: [9, 1, 5] },
        ],
      },
      decorations: { categoryAxis: true },
    } as ChartConfig);
    const axisOf = (p: string) =>
      scene.nodes
        .filter((n: any) => n.kind === "text" && n.name?.startsWith(`${p}-category-`))
        .sort((a: any, b: any) => a.x - b.x)
        .map((n: any) => n.text);
    const p0 = axisOf("p0"),
      p1 = axisOf("p1");
    expect(p0.length).toBeGreaterThan(0);
    expect(p1).toEqual(p0); // panels used to rank by their own series and disagree
  });
});

describe("layout indexing", () => {
  it("grouped treemap tiles get the rectangle for their own value", () => {
    const scene = buildChart({
      kind: "treemap",
      width: 600,
      height: 400,
      data: { categories: ["G | a", "G | b", "G | c"], series: [{ name: "S", values: [1, 50, 100] }] },
    } as ChartConfig);
    const area = (name: string) => {
      const t = scene.nodes.find((n: any) => n.kind === "rect" && n.name === name) as any;
      return t ? t.w * t.h : 0;
    };
    // values 1 < 50 < 100 must give areas tile-0 < tile-1 < tile-2.
    // (Members are listed ascending, so the sort reorders them — which is
    // exactly when the post-sort-index lookup handed over the wrong rect.)
    expect(area("tile-0")).toBeLessThan(area("tile-1"));
    expect(area("tile-1")).toBeLessThan(area("tile-2"));
  });

  it("stacked connectors join the same series across a zero segment", () => {
    const scene = buildChart({
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [0, 10] },
          { name: "S2", values: [20, 30] },
        ],
      },
      decorations: { connectors: true },
    } as ChartConfig);
    const conns = scene.nodes.filter((n: any) => n.kind === "line" && n.name?.startsWith("connector-")) as any[];
    // S1 is absent in category A, so only S2's boundary can be joined: exactly one
    // connector, and it must link the two S2 tops (20 -> 30+... ), never S1<->S2.
    expect(conns).toHaveLength(1);
    expect(conns[0].name).toBe("connector-0-1"); // series index 1 = S2, not push-order 0
  });

  it("a tight-spread violin category still renders", () => {
    const scene = buildChart({
      kind: "violin",
      width: 480,
      height: 300,
      data: {
        categories: ["Wide", "Tight"],
        series: [
          { name: "o1", values: [0, 3] },
          { name: "o2", values: [30, 3.1] },
          { name: "o3", values: [60, 3] },
          { name: "o4", values: [90, 3.05] },
          { name: "o5", values: [120, 3.02] },
        ],
      },
    } as ChartConfig);
    // Both categories have valid observations; neither may be silently dropped.
    expect(scene.nodes.some((n: any) => n.name === "violin-0")).toBe(true);
    expect(scene.nodes.some((n: any) => n.name === "violin-1")).toBe(true);
  });

  it("horizontal mekko hides labels in rows thinner than the font", () => {
    const cats = Array.from({ length: 16 }, (_, i) => `C${i + 1}`);
    const cfg = {
      kind: "mekko",
      width: 600,
      height: 160,
      horizontal: true,
      data: {
        categories: cats,
        series: [
          { name: "S1", values: cats.map(() => 50) },
          { name: "S2", values: cats.map(() => 50) },
        ],
      },
      decorations: { segmentLabels: true },
    } as ChartConfig;
    const scene = buildChart(cfg);
    const segs = scene.nodes.filter((n: any) => n.kind === "rect" && n.name?.startsWith("seg-")) as any[];
    const labels = scene.nodes.filter((n: any) => n.kind === "text" && n.name?.startsWith("label-")) as any[];
    // Rows here are ~6.6pt thick — far under the 11pt font. The old gate measured
    // the segment's 250pt value-axis length instead, so it stamped a label into
    // every one of them.
    expect(segs.length).toBe(32);
    expect(segs.every((s) => s.h < 11 * 1.25)).toBe(true);
    expect(labels).toHaveLength(0);
    // Vertical mekko is unaffected: its gate always measured r.h.
    const tall = buildChart({ ...cfg, horizontal: false } as ChartConfig);
    expect(tall.nodes.some((n: any) => n.kind === "text" && n.name?.startsWith("label-"))).toBe(true);
  });

  it("horizontal mekko keeps labels out of segments shorter than the font", () => {
    // Thick rows, hairline-short segments — the case the thickness gate alone
    // does not catch. The fit check tolerates 2pt of bleed (the text box is
    // drawn 4pt wider than the segment), which is fine for a vertical mekko's
    // wide columns but let three 3.8pt segments each print a 5.4pt "4",
    // overlapping their neighbours by 1.6pt.
    const scene = buildChart({
      kind: "mekko",
      width: 400,
      height: 300,
      horizontal: true,
      data: {
        categories: ["EMEA"],
        series: [
          { name: "A", values: [4] },
          { name: "B", values: [4] },
          { name: "C", values: [4] },
          { name: "D", values: [300] },
        ],
      },
    } as ChartConfig);
    const ink = (n: any) => {
      const w = textWidth(n.text, n.fontSize);
      const cx = n.x + n.w / 2;
      return { lo: cx - w / 2, hi: cx + w / 2 };
    };
    const labels = scene.nodes.filter((n: any) => n.kind === "text" && n.name?.startsWith("label-")) as any[];
    // Only D is wide enough to carry a label.
    expect(labels.map((l) => l.text)).toEqual(["300"]);
    // And no two labels may ever overlap.
    const inks = labels.map(ink).sort((a, b) => a.lo - b.lo);
    for (let i = 1; i < inks.length; i++) expect(inks[i].lo).toBeGreaterThanOrEqual(inks[i - 1].hi);
  });
});

describe("valueExtent reports what the layout draws", () => {
  /** What "Same scale" does: write the extent back as a hard scale override. */
  const underSameScale = (cfg: ChartConfig) => {
    const e = valueExtent(cfg)!;
    return buildChart({ ...cfg, scale: { min: e.min < 0 ? e.min : undefined, max: e.max } });
  };
  const inkSpan = (scene: { nodes: any[] }) => {
    const ys = scene.nodes.flatMap((n) =>
      n.kind === "line" ? [n.y1, n.y2] : n.kind === "rect" ? [n.y, n.y + n.h] : [],
    );
    return { top: Math.min(...ys), bottom: Math.max(...ys) };
  };

  it("treats an Error row as a whisker, not as a data point", () => {
    const cfg = {
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [10] },
          { name: "Error", values: [30] },
        ],
      },
    } as ChartConfig;
    // The whisker spans 10±30. The old extent was {0,30}: 30 was the Error row's
    // own magnitude mistaken for a value — neither the data range nor the drawn one.
    const ext = valueExtent(cfg)!;
    expect(ext.max).toBeGreaterThanOrEqual(40);
    expect(ext.min).toBeLessThanOrEqual(-20);
    const { top, bottom } = inkSpan(underSameScale(cfg));
    expect(top).toBeGreaterThanOrEqual(-1); // whisker used to reach y=-83
    expect(bottom).toBeLessThanOrEqual(301); // and y=465
  });

  it("covers a waterfall's Target row", () => {
    const cfg = {
      kind: "waterfall",
      width: 480,
      height: 300,
      waterfall: { totalIndices: [2] },
      data: {
        categories: ["Start", "Up", "End"],
        series: [
          { name: "V", values: [100, 20, 0] },
          { name: "Target", values: [null, null, 200] },
        ],
      },
    } as ChartConfig;
    expect(valueExtent(cfg)!.max).toBeGreaterThanOrEqual(200); // was 120 — the running total only
    const { bottom } = inkSpan(underSameScale(cfg));
    expect(bottom).toBeLessThanOrEqual(301);
  });

  it("does not sum a Target row into a stack", () => {
    const cfg = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [10] },
          { name: "Target", values: [5] },
        ],
      },
    } as ChartConfig;
    // The Target is a tick at 5, not another 5pt of stack: the column totals 10.
    expect(valueExtent(cfg)).toEqual({ min: 0, max: 10 });
  });

  it("covers an explicit threshold line above the data", () => {
    const cfg = {
      kind: "clustered",
      width: 480,
      height: 300,
      decorations: { valueLines: [{ mode: "value", value: 500 }] },
      data: { categories: ["A"], series: [{ name: "S", values: [10] }] },
    } as ChartConfig;
    expect(valueExtent(cfg)!.max).toBeGreaterThanOrEqual(500); // was 10
  });

  it("a mean value line needs no widening — it is inside the data by construction", () => {
    const plain = {
      kind: "clustered",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 20] }] },
    } as ChartConfig;
    const withMean = { ...plain, decorations: { valueLines: [{ mode: "mean" }] } } as ChartConfig;
    expect(valueExtent(withMean)).toEqual(valueExtent(plain));
  });
});

describe("valueExtent (Same Scale)", () => {
  const data2 = (a: number[], b: number[]): ChartConfig["data"] => ({
    categories: ["A", "B", "C"].slice(0, a.length),
    series: [
      { name: "S1", values: a },
      { name: "S2", values: b },
    ],
  });
  const c = (partial: Partial<ChartConfig>): ChartConfig => ({
    kind: "stacked",
    data: data2([10, 20, 30], [5, 5, 5]),
    ...DEFAULT_SIZE,
    ...partial,
  });

  it("stacked sums positives and negatives per category", () => {
    expect(valueExtent(c({ data: data2([10, -4, 30], [5, -6, 5]) }))).toEqual({ min: -10, max: 35 });
  });
  it("clustered/line use the raw value range", () => {
    expect(valueExtent(c({ kind: "clustered" }))).toEqual({ min: 0, max: 30 });
    expect(valueExtent(c({ kind: "line", data: data2([-2, 8, 4], [1, 1, 1]) }))).toEqual({ min: -2, max: 8 });
  });
  it("area stacks positive values from zero", () => {
    expect(valueExtent(c({ kind: "area" }))).toEqual({ min: 0, max: 35 });
  });
  it("waterfall tracks the running level", () => {
    expect(
      valueExtent(
        c({
          kind: "waterfall",
          data: { categories: ["Start", "Up", "Down", "End"], series: [{ name: "S", values: [50, 20, -30, 0] }] },
          waterfall: { totalIndices: [3] },
        }),
      ),
    ).toEqual({ min: 0, max: 70 });
  });
  it("returns null when there is nothing to measure", () => {
    expect(valueExtent(c({ data: { categories: [], series: [] } }))).toBeNull();
    expect(valueExtent(c({ kind: "pie" }))).toBeNull();
  });
});

/**
 * The two seams where "what the extent says" and "what the chart draws" parted.
 *
 * This is the recurring family in this codebase — `boxplotBoxes`, `drawnExtent`,
 * `waterfallChain` were all the same shape. It matters because the pane's Same
 * Scale writes `valueExtent`'s answer back as a HARD `cfg.scale`, which
 * suppresses every auto-widen downstream: anything under-reported here renders
 * off the shape, and the Office renderer applies no clamp, so the rects land on
 * the slide above the chart.
 */
describe("the extent measures what is drawn", () => {
  it("collapses the Other bucket first, as buildChart does", () => {
    // `collapseOther` sums the tail into one synthesized series. On a clustered
    // chart that bar is taller than any series `dataExtent` can see, and
    // `valueExtent` never ran the collapse — so it reported the largest single
    // series and Same Scale pinned the chart to a scale it does not fit.
    const cfg = {
      ...DEFAULT_SIZE,
      kind: "clustered",
      otherBucket: { max: 3 },
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "A", values: [50, 50] },
          { name: "B", values: [40, 40] },
          { name: "C", values: [30, 30] },
          { name: "D", values: [30, 30] },
          { name: "E", values: [30, 30] },
          { name: "F", values: [30, 30] },
        ],
      },
    } as unknown as ChartConfig;
    const ext = valueExtent(cfg)!;
    // C+D+E+F = 120, and that bar is on the chart.
    expect(ext.max, "the extent did not see the Other bar it will be asked to fit").toBeGreaterThanOrEqual(120);
    // And the chart drawn under Same Scale's own transform stays inside its frame.
    const scaled = buildChart({ ...cfg, scale: { min: undefined, max: ext.max } } as ChartConfig);
    const { top } = rectSpan(scaled);
    expect(top, "Same Scale drew the chart off the top of the shape").toBeGreaterThanOrEqual(0);
  });
});

/**
 * The anchor a decoration hangs its ink on must name a mark that exists.
 */
describe("columnValue on a clustered chart", () => {
  const clustered = (extra: Record<string, unknown> = {}, series?: unknown[]) =>
    ({
      ...DEFAULT_SIZE,
      kind: "clustered",
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: series ?? [
          { name: "2024", values: [20, 30, 25] },
          { name: "2025", values: [18, 22, 24] },
        ],
      },
      ...extra,
    }) as unknown as ChartConfig;

  const named = (scene: { nodes: { name?: string }[] }, re: RegExp) => scene.nodes.filter((n) => re.test(n.name ?? ""));

  it("anchors an Error row's whiskers on the bars, not on their sum", () => {
    // `columnValue` published the SUM of every series while `columnTop` is the
    // tallest single bar — two anchors naming different marks, and the sum is
    // drawn nowhere. `widenForAnatomy` widens to (tallest bar + maxPlus), so it
    // structurally cannot contain a whisker hung off the sum: measured at
    // y1 = -3.0, -79.7 and -63.2 on a 300pt canvas, real line shapes ABOVE the
    // chart, because the Office renderer applies no clamp.
    const scene = buildChart(
      clustered({}, [
        { name: "2024", values: [20, 30, 25] },
        { name: "2025", values: [18, 22, 24] },
        { name: "Error", values: [4, 4, 4] },
      ]),
    );
    const whiskers = named(scene, /^error-/) as { y1?: number; y2?: number }[];
    expect(whiskers.length, "no error whiskers were drawn, so this proves nothing").toBeGreaterThan(0);
    const ys = whiskers.flatMap((w) => [w.y1 ?? 0, w.y2 ?? 0]);
    expect(Math.min(...ys), "a whisker was drawn above the top of the shape").toBeGreaterThanOrEqual(0);
  });

  it("draws a mean value-line inside the plot, at the mean of the BARS", () => {
    const scene = buildChart(clustered({ decorations: { valueLines: [{ mode: "mean" }] } }));
    const lines = named(scene, /^value-line-\d/) as { y1?: number }[];
    expect(lines.length, "no value line was drawn, so this proves nothing").toBeGreaterThan(0);
    expect(Math.min(...lines.map((l) => l.y1 ?? 0)), "the mean line was drawn off the top").toBeGreaterThanOrEqual(0);
    // mean(20, 30, 25) = 25, the mean of what is drawn. The mean of the SUMS is
    // 46, against a value axis topping out near 30 — a label naming a level the
    // chart does not have.
    const label = named(scene, /^value-line-label-/)[0] as { text?: string } | undefined;
    expect(label?.text, "the mean line lost its label, so its number is unchecked").toBeTruthy();
    expect(label!.text).toContain("25");
  });
});

/**
 * A bullet graph is conventionally horizontal, and until now that was the one
 * orientation where it silently lost its data.
 *
 * `extractErrorRows` pulls `Target` and `Error*` rows out of the series for
 * every kind that supports them, regardless of orientation. The blocks that
 * draw them back were gated on `skipDecor` — which began with `cfg.horizontal`
 * — and on `valueToY`, which the layouts leave undefined when horizontal. So on
 * a bar chart the row was removed and never redrawn: the numbers the user typed
 * vanished, which is worse than undecorated, because the row would at least
 * have rendered as a bar.
 */
describe("Target and Error rows follow the value axis, whichever way it runs", () => {
  const KINDS = ["stacked", "clustered", "waterfall", "line", "area"] as const;
  const withRows = (kind: string, horizontal: boolean) =>
    ({
      ...DEFAULT_SIZE,
      kind,
      horizontal,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Actual", values: [20, 30, 25] },
          { name: "Target", values: [28, 34, 31] },
          { name: "Error", values: [4, 4, 4] },
        ],
      },
    }) as unknown as ChartConfig;

  const marks = (scene: { nodes: { name?: string; kind: string }[] }, re: RegExp) =>
    scene.nodes.filter((n) => n.kind === "line" && re.test(n.name ?? "")) as unknown as {
      x1: number;
      y1: number;
      x2: number;
      y2: number;
    }[];

  it.each(KINDS)("%s draws its target ticks in both orientations", (kind) => {
    for (const horizontal of [false, true]) {
      const scene = buildChart(withRows(kind, horizontal));
      const ticks = marks(scene, /^target-\d/);
      expect(ticks.length, `${kind} horizontal=${horizontal}: the Target row was deleted and never drawn`).toBe(3);
      // A tick runs ACROSS the bar: along x on a column chart, along y on a bar.
      for (const t of ticks) {
        const acrossX = Math.abs(t.x2 - t.x1) > Math.abs(t.y2 - t.y1);
        expect(acrossX, `${kind} horizontal=${horizontal}: the tick is turned the wrong way`).toBe(!horizontal);
      }
    }
  });

  it.each(KINDS)("%s keeps every mark inside the frame, both ways", (kind) => {
    // The widening that makes room for a whisker was ALSO gated on
    // `!cfg.horizontal` — the third place that assumption was written down.
    // Unwidened, a whisker for 34 landed at x = 530.9 on a 480-wide canvas.
    for (const horizontal of [false, true]) {
      const cfg = withRows(kind, horizontal);
      const scene = buildChart(cfg);
      const all = marks(scene, /^(target|error)/);
      expect(all.length, `${kind} horizontal=${horizontal}: nothing drawn, so this proves nothing`).toBeGreaterThan(0);
      for (const m of all) {
        const xs = [m.x1, m.x2];
        const ys = [m.y1, m.y2];
        expect(Math.min(...xs), `${kind} horizontal=${horizontal}: a mark ran off the left`).toBeGreaterThanOrEqual(0);
        expect(Math.max(...xs), `${kind} horizontal=${horizontal}: a mark ran off the right`).toBeLessThanOrEqual(
          cfg.width!,
        );
        expect(Math.min(...ys), `${kind} horizontal=${horizontal}: a mark ran off the top`).toBeGreaterThanOrEqual(0);
        expect(Math.max(...ys), `${kind} horizontal=${horizontal}: a mark ran off the bottom`).toBeLessThanOrEqual(
          cfg.height!,
        );
      }
    }
  });

  it("draws the bullet's range bands on a bar chart too", () => {
    // The other half of the documented recipe, dropped by the same gate. A
    // value band is a range of VALUES whichever way the chart is turned, so on
    // a bar it is a vertical strip spanning the plot's height.
    // Inside the value domain — a band entirely outside the plot is clipped to
    // nothing and dropped, which is correct and would make this vacuous.
    const bands = {
      bands: [
        { axis: "y", from: 0, to: 15 },
        { axis: "y", from: 15, to: 30 },
      ],
    };
    const cfg = { ...withRows("stacked", true), decorations: bands } as unknown as ChartConfig;
    const drawn = buildChart(cfg).nodes.filter((n) => /^band-\d/.test(n.name ?? "")) as unknown as {
      w: number;
      h: number;
    }[];
    expect(drawn.length, "the range bands were dropped on a bar chart").toBe(2);
    // Strips ACROSS the value axis: full plot height, narrower than they are tall.
    for (const b of drawn) expect(b.h, "a value band was laid out as if the chart were vertical").toBeGreaterThan(b.w);
  });
});

/**
 * What ELSE the rotation toggle used to drop.
 *
 * The Target/Error fix found the "assumes a vertical value axis" assumption in
 * four places. It was in more. Each of these is a documented feature the chart
 * silently omitted sideways — silently being the operative word: the output
 * looks entirely correct, so nothing prompts anyone to check.
 */
describe("the rotation toggle keeps the decorations that mean something", () => {
  const line = (horizontal: boolean, decorations: Record<string, unknown>) =>
    ({
      ...DEFAULT_SIZE,
      kind: "line",
      horizontal,
      decorations: { categoryAxis: true, valueAxis: true, ...decorations },
      data: {
        categories: ["2022", "2023", "2024", "2025F", "2026F"],
        series: [{ name: "revenue", values: [40, 46, 52, 58, 65] }],
      },
    }) as unknown as ChartConfig;

  it("draws a forecast as a forecast in both orientations", () => {
    // A projection rendered with the same solid stroke and the same filled
    // marker as measured data shows the reader a guess as if it were fact.
    for (const horizontal of [false, true]) {
      const scene = buildChart(line(horizontal, { forecastFrom: 3 }));
      const dashed = scene.nodes.filter(
        (n) => n.kind === "line" && /^line-/.test(n.name ?? "") && (n as { dash?: number[] }).dash,
      );
      const hollow = scene.nodes.filter(
        (n) => /^marker-/.test(n.name ?? "") && (n as { fill?: string }).fill === "none",
      );
      const divider = scene.nodes.filter((n) => n.name === "forecast-divider");
      expect(dashed.length, `horizontal=${horizontal}: forecast segments are not dashed`).toBe(2);
      expect(hollow.length, `horizontal=${horizontal}: forecast markers are not hollow`).toBe(2);
      expect(divider.length, `horizontal=${horizontal}: no actuals/forecast divider`).toBe(1);
    }
  });

  it("honours `stepped` in both orientations", () => {
    // A step series drawn as straight interpolation claims the value slid
    // gradually where the data says it jumped — the chart asserts something
    // about the world that the numbers do not.
    for (const horizontal of [false, true]) {
      const stepped = JSON.stringify(buildChart(line(horizontal, { stepped: "after" })));
      const plain = JSON.stringify(buildChart(line(horizontal, {})));
      expect(stepped, `horizontal=${horizontal}: \`stepped\` was a silent no-op`).not.toBe(plain);
    }
  });

  it("keeps lollipop, dot and dumbbell-range marks when the chart is turned", () => {
    // A Cleveland dot plot and a dumbbell are NORMALLY horizontal — long
    // category labels down the left is the whole point. Sideways, `range` gave
    // two full-length bars from zero: the connector carrying the entire meaning
    // gone, and the endpoints reading as two independent magnitudes.
    for (const barStyle of ["range", "dot", "lollipop"]) {
      for (const horizontal of [false, true]) {
        const scene = buildChart({
          ...DEFAULT_SIZE,
          kind: "clustered",
          horizontal,
          decorations: { categoryAxis: true, valueAxis: true, barStyle },
          data: {
            categories: ["N", "S", "E", "W"],
            series: [
              { name: "2024", values: [20, 30, 25, 18] },
              { name: "2025", values: [28, 34, 31, 26] },
            ],
          },
        } as unknown as ChartConfig);
        const dots = scene.nodes.filter((n) => n.kind === "ellipse");
        expect(dots.length, `${barStyle} horizontal=${horizontal}: collapsed back to plain bars`).toBe(8);
        for (const d of dots as unknown as { cx: number; cy: number }[]) {
          expect(d.cx, `${barStyle} horizontal=${horizontal}: a dot ran off the frame`).toBeGreaterThanOrEqual(0);
          expect(d.cy).toBeGreaterThanOrEqual(0);
          expect(d.cx).toBeLessThanOrEqual(DEFAULT_SIZE.width!);
          expect(d.cy).toBeLessThanOrEqual(DEFAULT_SIZE.height!);
        }
        if (barStyle === "range") {
          expect(
            scene.nodes.filter((n) => /^range-/.test(n.name ?? "")).length,
            `range horizontal=${horizontal}: the connector is what carries the meaning`,
          ).toBe(4);
        }
      }
    }
  });
});

/**
 * `valueExtent`'s own comment claims it runs "the SAME normalisation buildChart
 * runs before it lays anything out, in the same order" — and it skipped
 * `normalizeConfig`, the first thing buildChart does. That is what guarantees
 * `Series.name` is a string and drops a null series entry, and the three passes
 * after it call `s.name.trim()` unguarded. So a series written as
 * `{ values: [...] }` with no name — which the skill writes verbatim into a
 * POWERCHART_CONFIG tag — rendered fine and made this throw. `doSameScale` maps
 * it over the deck without a guard, so one such chart aborted the whole
 * Same-Scale operation and nothing was rescaled.
 */
describe("valueExtent survives everything buildChart survives", () => {
  const cases: [string, unknown][] = [
    ["a series with no name", { categories: ["A", "B"], series: [{ values: [1, 2] }] }],
    ["a null series entry", { categories: ["A", "B"], series: [null, { name: "S", values: [1, 3] }] }],
    ["a numeric series name", { categories: ["A", "B"], series: [{ name: 7, values: [1, 2] }] }],
  ];
  for (const [label, data] of cases) {
    it(`measures ${label} instead of throwing`, () => {
      const cfg = { kind: "clustered", ...DEFAULT_SIZE, data } as unknown as ChartConfig;
      // The control: buildChart already copes, so the two must not disagree.
      expect(() => buildChart(cfg), `buildChart broke on ${label}`).not.toThrow();
      const ext = valueExtent(cfg);
      expect(ext, label).not.toBeNull();
      expect(Number.isFinite(ext!.max), label).toBe(true);
    });
  }
});

/**
 * The block that widens the auto scale so error whiskers and target ticks stay
 * inside the plot was entered only when `cfg.scale?.max == null` — while the
 * block itself already honoured a pinned min. So the widening was asymmetric:
 * pinning only `scale.min` still widened the max, and pinning only `scale.max`
 * skipped the min widening entirely. `toY` then clipped the out-of-range mark
 * to the plot edge (deliberate, for bars), so a Target below the auto minimum
 * was drawn exactly on the zero baseline — a mark that asserts a value,
 * printed at the wrong one, with nothing to say it was clipped.
 */
describe("pinning one end of the scale still widens the other", () => {
  const withTarget = (scale?: { min?: number; max?: number }) =>
    buildChart({
      kind: "clustered",
      ...DEFAULT_SIZE,
      ...(scale ? { scale } : {}),
      decorations: { valueAxis: true },
      data: {
        categories: ["A"],
        series: [
          { name: "V", values: [10] },
          { name: "Target", values: [-30] },
        ],
      },
    } as unknown as ChartConfig).nodes;

  const marks = (scale?: { min?: number; max?: number }) => {
    const nodes = withTarget(scale);
    const line = (name: string) => nodes.find((n): n is LineNode => n.kind === "line" && n.name === name)!;
    return { baseline: line("baseline").y1, target: line("target-0").y1 };
  };

  it("does not print a below-range Target on the zero baseline", () => {
    const pinnedMax = marks({ max: 100 });
    expect(Math.abs(pinnedMax.target - pinnedMax.baseline), "target clipped onto the baseline").toBeGreaterThan(1);
    // Both ends already widen when nothing is pinned, and pinning only the min
    // always worked — the controls that say the fix did not simply disable the
    // pin.
    const auto = marks();
    expect(Math.abs(auto.target - auto.baseline)).toBeGreaterThan(1);
    const pinnedMin = marks({ min: -40 });
    expect(Math.abs(pinnedMin.target - pinnedMin.baseline)).toBeGreaterThan(1);
  });

  it("still honours a pin that the data does fit inside", () => {
    // The negative control: the widen must not overwrite a usable pinned max.
    const nodes = withTarget({ min: -40, max: 100 });
    const axis = nodes.filter((n) => n.name === "value-axis").map((n) => (n as unknown as { text: string }).text);
    expect(axis.some((t) => t.includes("100"))).toBe(true);
  });
});

/**
 * A chart may be zoomed. It may not be drawn off the slide.
 */
describe("a manual scale narrower than the data", () => {
  it("clips at the axis instead of extrapolating off the canvas", () => {
    // The pane invites this — "Axis scale min / max" are free-text boxes — and
    // pinning a revenue chart to 0-100 (thinking percent) put the bars at
    // y = -231558 on a 300pt canvas. The SVG preview hides it behind its
    // viewBox, so the first anyone saw was a .pptx whose group extent exceeds
    // the OOXML coordinate limit and which PowerPoint offers to repair.
    const cfg = {
      ...DEFAULT_SIZE,
      kind: "clustered",
      scale: { min: 0, max: 100 },
      decorations: { valueAxis: true, categoryAxis: true },
      data: { categories: ["Q1", "Q2"], series: [{ name: "Rev", values: [80000, 92000] }] },
    } as unknown as ChartConfig;
    const rects = buildChart(cfg).nodes.filter((n) => n.kind === "rect") as unknown as {
      y: number;
      h: number;
    }[];
    expect(rects.length, "nothing was drawn, so this proves nothing").toBeGreaterThan(0);
    for (const r of rects) {
      expect(r.y, "a bar was drawn above the canvas").toBeGreaterThanOrEqual(0);
      expect(r.y + r.h, "a bar was drawn below the canvas").toBeLessThanOrEqual(DEFAULT_SIZE.height! + 1);
    }
  });

  it("clips on a LOG axis, which returns before the clamp was declared", () => {
    /**
     * The clamp above was written for the linear branch. `valueScale`'s log
     * branch sits thirty lines EARLIER and returns its own `toY`, so it never
     * saw the fix — and it survived the test beside it because that test never
     * sets `logScale`.
     *
     * `scale.max: 100` over data reaching 10,000 put the top bar at y = -540 on
     * a 300pt canvas: 540 points above the slide. `clip` is declared above both
     * branches now, so they cannot answer this differently again.
     */
    const cfg = {
      ...DEFAULT_SIZE,
      kind: "clustered",
      logScale: true,
      scale: { max: 100 },
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 10000] }] },
    } as unknown as ChartConfig;
    const rects = buildChart(cfg).nodes.filter((n) => n.kind === "rect") as unknown as { y: number; h: number }[];
    expect(rects.length, "nothing was drawn, so this proves nothing").toBeGreaterThan(0);
    for (const r of rects) {
      expect(r.y, "a bar was drawn above the canvas").toBeGreaterThanOrEqual(0);
      expect(r.y + r.h, "a bar was drawn below the canvas").toBeLessThanOrEqual(DEFAULT_SIZE.height! + 1);
    }
  });

  it("clips SIDEWAYS too, on every kind that rotates", () => {
    // The clamp lived in `valueScale.toY`, and only the vertical branch of each
    // layout routed through it — the horizontal branch is its own linear map and
    // had none. The same config that drew a full-height bar upright drew a
    // 30,128pt one sideways, off the slide and past the OOXML coordinate limit.
    for (const kind of ["clustered", "stacked", "waterfall", "boxplot", "line", "area"] as const) {
      const cfg = {
        width: 400,
        height: 300,
        kind,
        horizontal: true,
        scale: { min: 0, max: 100 },
        data: {
          categories: ["A", "B"],
          series: [
            { name: "Rev", values: [5000, 8000] },
            { name: "Two", values: [6000, 9000] },
          ],
        },
      } as unknown as ChartConfig;
      const drawn = buildChart(cfg).nodes.filter(
        (n): n is RectNode => n.kind === "rect" && !n.name?.startsWith("band-"),
      );
      expect(drawn.length, `${kind}: nothing was drawn, so this proves nothing`).toBeGreaterThan(0);
      for (const r of drawn) {
        expect(r.x, `${kind}: drawn left of the canvas`).toBeGreaterThanOrEqual(-1);
        expect(r.x + r.w, `${kind}: drawn past the right edge`).toBeLessThanOrEqual(401);
      }
    }
  });

  it("caps a chart dimension below the size pptxgenjs stops rounding", () => {
    // pptxgenjs treats any number >= 100 as EMU already, so AT 100 inches and
    // past it the value is written through unrounded — and
    // `ST_PositiveCoordinate` is an xsd:long, so that part is schema-invalid and
    // the user meets the repair dialog. The ceiling is fifteen times the widest
    // slide, so it costs nothing real.
    //
    // ASSERTED AS THE PROPERTY, not as the number, because the number was wrong
    // and this test pinned it. 7200pt is EXACTLY 100 inches — `7200 / 72 === 100`
    // with no float slack — so the old ceiling did not avoid pptxgenjs's
    // threshold, it landed precisely on it, and `Math.min` meant every oversize
    // request was clamped TO the one broken value. A chart at `width: 10000`
    // shipped its title box as `<a:ext cx="100"/>`, a ten-thousandth of an inch,
    // beside a sibling frame at `cx="91440000"`.
    const capped = buildChart({
      ...DEFAULT_SIZE,
      width: 7300,
      kind: "clustered",
      data: { categories: ["A"], series: [] },
    } as unknown as ChartConfig).width;
    expect(capped / 72, "the cap lands ON pptxgenjs's >=100in threshold, not below it").toBeLessThan(100);
    // And every oversize request lands there, so the boundary is the only value
    // that matters — one point either side is the whole bug.
    for (const w of [7200, 7300, 10000, 1e6]) {
      const got = buildChart({
        ...DEFAULT_SIZE,
        width: w,
        kind: "clustered",
        data: { categories: ["A"], series: [] },
      } as unknown as ChartConfig).width;
      expect(got / 72, `width ${w} clamped onto the 100-inch threshold`).toBeLessThan(100);
    }
    expect(
      buildChart({
        ...DEFAULT_SIZE,
        width: 900,
        kind: "clustered",
        data: { categories: ["A"], series: [] },
      } as unknown as ChartConfig).width,
    ).toBe(900);
  });
});

/**
 * Four smaller ways the output disagreed with itself.
 */
describe("labels, webs and the two paint sinks", () => {
  it("gives a sub-unit chart labels that are not all '0.00'", () => {
    // The auto ladder stopped at 2 decimals however small the data got, so
    // rates, yields, defect fractions and probabilities — everything under 1 —
    // printed "0.00" on every label for values that plainly differ. And the
    // chart said so out loud: `resolveAxisFormat` widens on the tick STEP, so
    // the same chart's axis read 0.000 / 0.001 / 0.002 beside those bars.
    const scene = buildChart({
      ...DEFAULT_SIZE,
      kind: "clustered",
      decorations: { segmentLabels: true },
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [0.0012, 0.0031, 0.0024] }] },
    } as unknown as ChartConfig);
    const labels = scene.nodes.filter((n) => /^label-/.test(n.name ?? "")) as unknown as { text: string }[];
    expect(labels.length, "no labels drawn, so this proves nothing").toBe(3);
    expect(new Set(labels.map((l) => l.text)).size, "three different values printed the same text").toBe(3);
    for (const l of labels) expect(l.text).not.toBe("0.00");
  });

  it("keeps a radar vertex inside its own web when a value exceeds the scale", () => {
    // The scale clamped values below `min` onto the centre and let anything
    // above `max` run past the outer radius unbounded — and `sampleConfig`
    // ships `{min:0,max:5}`, so no bad scale has to be typed: edit one cell
    // from 4 to 8, as anyone would for a 1-10 maturity scale, and the vertex
    // leaves the web and the canvas. A web is a picture of a scale; a point
    // outside it is not on the scale.
    const cfg = JSON.parse(JSON.stringify(sampleConfig("radar")));
    cfg.data.series[0].values[0] = 8;
    const scene = buildChart({ ...DEFAULT_SIZE, ...cfg });
    const ys = scene.nodes.filter((n) => n.kind === "ellipse") as unknown as { cy: number }[];
    expect(ys.length, "no markers drawn, so this proves nothing").toBeGreaterThan(0);
    expect(Math.min(...ys.map((e) => e.cy)), "a vertex left the canvas").toBeGreaterThanOrEqual(0);
  });

  it("treats `transparent` as transparent in the live renderer's colour sink", () => {
    // The documented floating-segment idiom. Only two layouts guarded it before
    // it reached a renderer, and mekko's guard spells out why — "Office.js hands
    // 'transparent' to setSolidColor, which it rejects" — but the guard was
    // never swept to the SINK, so every other kind sent the bare word to a live
    // host. `alphaOf` now knows the keyword its own sibling in the skill's
    // paint module has always known.
    expect(alphaOf("transparent")).toBe(0);
    expect(alphaOf("TRANSPARENT")).toBe(0);
    expect(alphaOf("#2a78d6"), "an ordinary colour lost its opacity").toBe(1);
  });
});

/**
 * The last two the rotation toggle dropped — the two that needed real geometry
 * rather than a turned coordinate.
 */
describe("smooth and fillBetween, sideways", () => {
  const line = (horizontal: boolean, decorations: Record<string, unknown>) =>
    ({
      ...DEFAULT_SIZE,
      kind: "line",
      horizontal,
      decorations: { categoryAxis: true, valueAxis: true, ...decorations },
      data: {
        categories: ["Q1", "Q2", "Q3", "Q4"],
        series: [
          { name: "Plan", values: [10, 18, 14, 22] },
          { name: "Actual", values: [12, 15, 19, 20] },
        ],
      },
    }) as unknown as ChartConfig;

  it("curves a smoothed series in both orientations", () => {
    for (const horizontal of [false, true]) {
      const curved = buildChart(line(horizontal, { smooth: true }));
      const straight = buildChart(line(horizontal, {}));
      expect(JSON.stringify(curved), `horizontal=${horizontal}: \`smooth\` was a silent no-op`).not.toBe(
        JSON.stringify(straight),
      );
      // Sampled segments, not one line per category gap.
      const sampled = curved.nodes.filter((n) => /-s\d+$/.test(n.name ?? ""));
      expect(sampled.length, `horizontal=${horizontal}: no spline samples`).toBeGreaterThan(20);
    }
  });

  it("fills the gap between two series in both orientations", () => {
    for (const horizontal of [false, true]) {
      const scene = buildChart(line(horizontal, { fillBetween: [0, 1] }));
      const band = scene.nodes.filter((n) => /^fill-between-/.test(n.name ?? "")) as unknown as {
        x: number;
        y: number;
        w: number;
        h: number;
      }[];
      expect(band.length, `horizontal=${horizontal}: the plan-vs-actual band was not drawn`).toBeGreaterThan(10);
      for (const r of band) {
        expect(r.x, `horizontal=${horizontal}: a band slab left the frame`).toBeGreaterThanOrEqual(0);
        expect(r.y).toBeGreaterThanOrEqual(0);
        expect(r.x + r.w).toBeLessThanOrEqual(DEFAULT_SIZE.width! + 1);
        expect(r.y + r.h).toBeLessThanOrEqual(DEFAULT_SIZE.height! + 1);
      }
      // Turned: the slabs march along the CATEGORY axis, which is y here.
      const spread = (vals: number[]) => Math.max(...vals) - Math.min(...vals);
      const alongY = spread(band.map((r) => r.y));
      const alongX = spread(band.map((r) => r.x));
      expect(horizontal ? alongY > alongX : alongX > alongY, "the band marched along the wrong axis").toBe(true);
    }
  });
});
