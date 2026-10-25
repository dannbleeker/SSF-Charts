import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { textWidth } from "../src/core/scene";
import type { EllipseNode, LineNode, Scene, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { withOptionData } from "./helpers/sweep-options";

/**
 * The overflows the repaired sweeps now see, and the two label siblings next to
 * them: the butterfly's stacked legend, the boxplot's mean marker, the mekko's
 * label content, and a sideways lollipop's label on its own stem.
 *
 * Every test here was run against the code without its fix and went red on the
 * assertion it names; the controls beside each stay green either way.
 */

/** The frame gate's ink rule, so these measure what `frame-fit` and `overlap-budget` measure. */
const ink = (t: TextNode) => {
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
type Ink = ReturnType<typeof ink>;

const overlapArea = (a: Ink, b: Ink) => {
  const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
  const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
  return w > 0 && h > 0 ? w * h : 0;
};

const texts = (scene: Scene) => scene.nodes.filter((n): n is TextNode => n.kind === "text" && !!String(n.text).trim());
const named = (scene: Scene, re: RegExp) => scene.nodes.filter((n) => re.test(n.name ?? ""));

/** The frames, fonts and orientations the frame and overlap sweeps run. */
const FRAMES: [number, number][] = [
  [80, 60],
  [120, 90],
  [160, 120],
  [200, 150],
  [300, 60],
  [60, 300],
  [480, 300],
  [960, 540],
];
const cells = FRAMES.flatMap(([w, h]) =>
  [10, 18].flatMap((fontSize) => [false, true].map((horizontal) => ({ w, h, fontSize, horizontal }))),
);
const label = (c: { w: number; h: number; fontSize: number; horizontal: boolean }) =>
  `${c.w}x${c.h}@${c.fontSize}${c.horizontal ? " rotated" : ""}`;

/* -------------------------------------------------------------------------- */
/* A. The butterfly's stacked legend                                          */
/* -------------------------------------------------------------------------- */

describe("a butterfly with stacked flanks draws its legend on the chart or not at all", () => {
  const stackedFlanks = (w: number, h: number, fontSize: number, extra: object = {}): ChartConfig =>
    ({
      ...withOptionData("butterfly.split", sampleConfig("butterfly")),
      width: w,
      height: h,
      style: { fontSize },
      butterfly: { split: 1 },
      ...extra,
    }) as ChartConfig;

  const legend = (scene: Scene) => named(scene, /^legend(-chip)?-\d+$/) as (TextNode | { y: number; h: number })[];
  const legendBottom = (scene: Scene) => Math.max(0, ...legend(scene).map((n) => n.y + n.h));

  it("no legend row hangs below the chart, over the frame sweep's cells", () => {
    const bad: string[] = [];
    for (const cell of cells) {
      const scene = buildChart({ ...stackedFlanks(cell.w, cell.h, cell.fontSize), horizontal: cell.horizontal });
      const past = legendBottom(scene) - scene.height;
      if (past > 0.01) bad.push(`${label(cell)}: ${past.toFixed(1)}pt below the chart`);
    }
    expect(bad, "the stacked legend wraps to more rows than the frame has room for").toEqual([]);
  });

  it("a legend that is drawn never has bars drawn through it", () => {
    // The reservation is only honest while the plot it leaves reaches its floor:
    // below that `fitPlot` grows the plot back UP and the bars rise into the
    // legend. Swept finely, because the frame-fit cells miss most of this band.
    const bad: string[] = [];
    for (let w = 60; w <= 240; w += 20)
      for (let h = 60; h <= 240; h += 20)
        for (const fontSize of [8, 10, 14, 18]) {
          const scene = buildChart(stackedFlanks(w, h, fontSize));
          if (!legend(scene).length) continue;
          const barTop = Math.min(...named(scene, /^seg-/).map((n) => (n as { y: number }).y));
          if (barTop < legendBottom(scene) - 1e-6)
            bad.push(
              `${w}x${h}@${fontSize}: bars start at ${barTop.toFixed(1)}, legend ends at ${legendBottom(scene).toFixed(1)}`,
            );
        }
    expect(bad).toEqual([]);
  });

  it("a legend that cannot be paid for is dropped and the chart is not: every bar is still drawn", () => {
    const cfg = stackedFlanks(80, 60, 18);
    const scene = buildChart(cfg);
    expect(legend(scene)).toHaveLength(0);
    expect(named(scene, /^seg-/)).toHaveLength(cfg.data.categories.length * cfg.data.series.length);
  });

  it("control: at the sizes a chart is placed at the legend is where it always was", () => {
    for (const [w, h] of [
      [480, 300],
      [960, 540],
    ]) {
      const scene = buildChart(stackedFlanks(w, h, 10));
      expect(named(scene, /^legend-chip-\d+$/), `${w}x${h}`).toHaveLength(3);
      // One row, hung from the title's band.
      expect(
        new Set(
          texts(scene)
            .filter((t) => /^legend-\d+$/.test(t.name ?? ""))
            .map((t) => t.y),
        ).size,
      ).toBe(1);
    }
  });

  it("control: the shipped showcase slide (480x300, split 2, four series) keeps its legend", () => {
    const scene = buildChart({
      kind: "butterfly",
      width: 480,
      height: 300,
      title: "Butterfly with stacked flanks",
      butterfly: { split: 2 },
      data: {
        categories: ["<30", "30–39", "40–49", "50–59", "60+"],
        series: [
          { name: "Women FT", values: [220, 340, 300, 180, 60] },
          { name: "Women PT", values: [200, 270, 240, 140, 60] },
          { name: "Men FT", values: [240, 420, 400, 280, 110] },
          { name: "Men PT", values: [140, 230, 200, 130, 50] },
        ],
      },
      decorations: { segmentLabels: true },
    } as ChartConfig);
    expect(named(scene, /^legend-chip-\d+$/)).toHaveLength(4);
  });

  it("control: the classic two-series butterfly keeps its two headers on a frame that can carry them", () => {
    const scene = buildChart({
      ...sampleConfig("butterfly"),
      width: 480,
      height: 300,
      butterfly: undefined,
    } as ChartConfig);
    expect(named(scene, /^header-\d$/)).toHaveLength(2);
    expect(named(scene, /^legend/)).toHaveLength(0);
  });

  it("a footnote counts against the legend too: a legend that is kept ends above the source line", () => {
    const bad: string[] = [];
    let kept = 0;
    for (let w = 100; w <= 300; w += 20)
      for (let h = 80; h <= 200; h += 20)
        for (const fontSize of [10, 14]) {
          const scene = buildChart(stackedFlanks(w, h, fontSize, { footnote: "Source: an internal model" }));
          if (!legend(scene).length) continue;
          kept++;
          const noteTop = scene.height - fontSize * 1.3;
          if (legendBottom(scene) > noteTop + 1e-6) bad.push(`${w}x${h}@${fontSize}`);
        }
    expect(kept, "the sweep kept no legend at all, so it measured nothing").toBeGreaterThan(20);
    expect(bad).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* B. The boxplot's mean marker                                               */
/* -------------------------------------------------------------------------- */

describe("a boxplot's mean markers collide with nothing", () => {
  const rawWithMean = (cell: { w: number; h: number; fontSize: number; horizontal: boolean }): ChartConfig =>
    ({
      ...withOptionData("boxplot", sampleConfig("boxplot")),
      width: cell.w,
      height: cell.h,
      horizontal: cell.horizontal,
      style: { fontSize: cell.fontSize },
      boxplot: { notch: true, showMean: true, jitter: true },
    }) as ChartConfig;

  it("no mean marker overlaps another label, over the overlap sweep's cells", () => {
    const bad: string[] = [];
    for (const cell of cells) {
      const ts = texts(buildChart(rawWithMean(cell)));
      const boxes = ts.map(ink);
      for (let i = 0; i < ts.length; i++)
        for (let j = i + 1; j < ts.length; j++)
          if (/^mean-/.test(ts[i].name ?? "") || /^mean-/.test(ts[j].name ?? ""))
            if (overlapArea(boxes[i], boxes[j]) > 1) bad.push(`${label(cell)}: ${ts[i].name} / ${ts[j].name}`);
    }
    expect(bad, "a mean marker was drawn over a neighbour, the title, a category name or a tick").toEqual([]);
  });

  it("two neighbouring markers abut at worst: the marker never outgrows its slot", () => {
    // 60x300 at 18pt: 5.5pt slots for an 18pt glyph.
    const scene = buildChart(rawWithMean({ w: 60, h: 300, fontSize: 18, horizontal: false }));
    const marks = named(scene, /^mean-/) as TextNode[];
    expect(marks.length).toBeGreaterThan(1);
    for (let i = 1; i < marks.length; i++) {
      const a = ink(marks[i - 1]);
      const b = ink(marks[i]);
      expect(b.x0 - a.x1, `${marks[i - 1].name} / ${marks[i].name}`).toBeGreaterThanOrEqual(-1e-6);
    }
    // It gave up size rather than being dropped: every category still has one.
    expect(marks).toHaveLength(4);
    for (const m of marks) expect(m.fontSize).toBeLessThan(18);
  });

  it("a marker with no room for a readable size is dropped, and the box it belongs to is not", () => {
    const scene = buildChart(rawWithMean({ w: 300, h: 60, fontSize: 18, horizontal: false }));
    expect(named(scene, /^mean-/)).toHaveLength(0);
    expect(named(scene, /^box-\d$/)).toHaveLength(4);
  });

  it("a marker stays above the foot of the plot, where the category names start", () => {
    const scene = buildChart(rawWithMean({ w: 120, h: 90, fontSize: 18, horizontal: false }));
    const base = named(scene, /^baseline$/)[0] as LineNode;
    const marks = named(scene, /^mean-/) as TextNode[];
    expect(marks.length).toBeGreaterThan(0);
    for (const m of marks) expect(ink(m).y1, m.name).toBeLessThanOrEqual(base.y1 + 1e-6);
  });

  it("control: at an ordinary size and font the marker is drawn at the chart font, centred on its category", () => {
    for (const [w, h, fontSize, horizontal] of [
      [480, 300, 10, false],
      [480, 300, 18, false],
      [960, 540, 10, true],
      [200, 150, 10, false],
    ] as const) {
      const scene = buildChart(rawWithMean({ w, h, fontSize, horizontal }));
      const marks = named(scene, /^mean-/) as TextNode[];
      expect(marks, `${w}x${h}@${fontSize}`).toHaveLength(4);
      marks.forEach((m, c) => {
        expect(m.fontSize).toBe(fontSize);
        expect(m.w).toBe(fontSize * 1.2);
        expect(m.h).toBe(fontSize * 1.4);
        // The median line is drawn at the category's centre line.
        const med = named(scene, new RegExp(`^median-${c}$`))[0] as LineNode;
        const centre = horizontal ? (med.y1 + med.y2) / 2 : (med.x1 + med.x2) / 2;
        expect(horizontal ? m.y + m.h / 2 : m.x + m.w / 2, `${m.name} at ${w}x${h}`).toBeCloseTo(centre, 9);
      });
    }
  });

  it("control: a precomputed Mean row is a marker like any other", () => {
    const scene = buildChart({
      kind: "boxplot",
      width: 480,
      height: 300,
      data: {
        categories: ["North", "South"],
        series: [
          { name: "Min", values: [2, 3] },
          { name: "Q1", values: [3, 5] },
          { name: "Median", values: [4, 7] },
          { name: "Q3", values: [6, 9] },
          { name: "Max", values: [8, 12] },
          { name: "Mean", values: [4.5, 7.2] },
        ],
      },
    } as ChartConfig);
    const marks = named(scene, /^mean-/) as TextNode[];
    expect(marks).toHaveLength(2);
    for (const m of marks) expect(m.fontSize).toBe(10);
  });
});

/* -------------------------------------------------------------------------- */
/* C1. The mekko's label content                                              */
/* -------------------------------------------------------------------------- */

describe("a mekko's segment labels honour decorations.labelContent", () => {
  const mekko = (labelContent?: unknown, horizontal = false): ChartConfig =>
    ({
      kind: "mekko",
      width: 480,
      height: 300,
      horizontal,
      data: {
        categories: ["Alpha", "Beta"],
        series: [
          { name: "Hardware", values: [30, 10] },
          { name: "Services", values: [10, 30] },
        ],
      },
      decorations: { segmentLabels: true, ...(labelContent === undefined ? {} : { labelContent }) },
    }) as unknown as ChartConfig;

  const byName = (scene: Scene) =>
    Object.fromEntries(
      (named(scene, /^label-\d+-\d+$/) as TextNode[]).map((n) => [n.name!.replace("label-", ""), n.text]),
    );

  it.each([false, true])("percent prints each segment's share of its column (horizontal: %s)", (horizontal) => {
    expect(byName(buildChart(mekko(["percent"], horizontal)))).toEqual({
      "0-0": "75%",
      "1-0": "25%",
      "0-1": "25%",
      "1-1": "75%",
    });
  });

  it("series and category parts are read, in the order given", () => {
    const scene = buildChart(mekko(["series", "value"]));
    expect(byName(scene)["0-0"]).toBe("Hardware 30");
    expect(byName(buildChart(mekko(["category", "percent"])))["1-1"]).toBe("Beta 75%");
  });

  it("an empty list means no label, and draws no empty box", () => {
    expect(named(buildChart(mekko([])), /^label-/)).toHaveLength(0);
  });

  it("a bare string is read as a list of one, as everywhere else", () => {
    expect(byName(buildChart(mekko("percent")))["0-0"]).toBe("75%");
  });

  it("control: with no labelContent the labels are the values, as they always were", () => {
    expect(byName(buildChart(mekko()))).toEqual({ "0-0": "30", "1-0": "10", "0-1": "10", "1-1": "30" });
  });

  it("a label that does not fit its segment is dropped, not drawn over its neighbour", () => {
    // "Alpha Hardware 30" fits a 480pt chart's segments and not a 120pt one's.
    const long = ["category", "series", "value"];
    expect(named(buildChart(mekko(long)), /^label-/)).toHaveLength(4);
    const small = buildChart({ ...mekko(long), width: 120, height: 90 } as ChartConfig);
    for (const t of named(small, /^label-/) as TextNode[]) {
      const seg = named(small, new RegExp(`^seg-${t.name!.slice("label-".length)}$`))[0] as { w: number };
      expect(textWidth(t.text, t.fontSize), t.name).toBeLessThanOrEqual(seg.w + 2);
    }
    expect(named(small, /^label-/).length).toBeLessThan(4);
  });
});

/* -------------------------------------------------------------------------- */
/* C2. A sideways lollipop's label is not drawn on its stem                   */
/* -------------------------------------------------------------------------- */

describe("a sideways dot's label is not drawn on the line the dot hangs from", () => {
  const dots = (
    barStyle: "lollipop" | "dot" | "range",
    series: number[][],
    size: [number, number] = [480, 300],
    horizontal = true,
  ): ChartConfig =>
    ({
      kind: "clustered",
      width: size[0],
      height: size[1],
      horizontal,
      data: {
        categories: ["North", "South", "East", "West"].slice(0, series[0].length),
        series: series.map((values, i) => ({ name: `S${i}`, values })),
      },
      decorations: { barStyle, segmentLabels: true, categoryAxis: true, valueAxis: true },
    }) as ChartConfig;

  const labels = (scene: Scene) => named(scene, /^label-\d+-\d+$/) as TextNode[];
  const lines = (scene: Scene) => named(scene, /^(stem|range)-/) as LineNode[];

  /** Every (label, line) pair whose ink the line is drawn through. */
  const crossings = (scene: Scene) => {
    const out: string[] = [];
    for (const l of labels(scene)) {
      const b = ink(l);
      for (const line of lines(scene)) {
        const half = (line.strokeWidth ?? 1) / 2;
        const lo = Math.min(line.x1, line.x2);
        const hi = Math.max(line.x1, line.x2);
        if (line.y1 >= b.y0 - half && line.y1 <= b.y1 + half && lo < b.x1 && hi > b.x0)
          out.push(`${l.name} / ${line.name}`);
      }
    }
    return out;
  };

  it("the label of the maximum, which flips left of its dot onto the stem, is lifted clear of it", () => {
    const scene = buildChart(dots("lollipop", [[12, 30]]));
    const max = labels(scene).find((l) => l.text === "30")!;
    // It did flip: the canvas ends before the label would.
    expect(max.align).toBe("right");
    expect(crossings(scene)).toEqual([]);
    // …and rests above the stem, not below it.
    const stem = lines(scene).find((l) => l.name === "stem-0-1")!;
    expect(ink(max).y1).toBeLessThanOrEqual(stem.y1 - 0.75);
  });

  it.each([
    ["lollipop [12,30,20,5]", dots("lollipop", [[12, 30, 20, 5]])],
    ["lollipop below zero, whose stem runs back to the label's side", dots("lollipop", [[-12, -30, 20, 5]])],
    [
      "two lollipop series",
      dots("lollipop", [
        [12, 30],
        [25, 10],
      ]),
    ],
    [
      "a dumbbell, whose connector runs from one dot to the other",
      dots("range", [
        [12, 20],
        [30, 25],
      ]),
    ],
    ["a lollipop in a shorter frame", dots("lollipop", [[12, 30, 20, 5]], [480, 160])],
    ["a lollipop at the larger preview size", dots("lollipop", [[12, 30, 20, 5]], [960, 540])],
  ] as [string, ChartConfig][])("no label has its line drawn through it: %s", (_name, cfg) => {
    const scene = buildChart(cfg);
    expect(labels(scene).length).toBeGreaterThan(0);
    expect(lines(scene).length).toBeGreaterThan(0);
    expect(crossings(scene)).toEqual([]);
  });

  it("a label lifted off its stem still names its own dot: it ends beside the dot, not past it", () => {
    const scene = buildChart(dots("lollipop", [[12, 30]]));
    const dot = named(scene, /^seg-0-1$/)[0] as EllipseNode;
    const max = labels(scene).find((l) => l.text === "30")!;
    // Right-aligned, ending a dot radius and two points left of the dot's centre.
    expect(max.x + max.w).toBeCloseTo(dot.cx - (dot.rx + 2), 6);
  });

  it("control: a label on the far side of its dot stays centred on the row, unchanged", () => {
    const scene = buildChart(dots("lollipop", [[12, 30]]));
    const small = labels(scene).find((l) => l.text === "12")!;
    const dot = named(scene, /^seg-0-0$/)[0] as EllipseNode;
    expect(small.align).toBe("left");
    expect(small.y + small.h / 2).toBeCloseTo(dot.cy, 9);
    expect(small.fontSize).toBe(10);
  });

  it("control: a dot has no line, so its flipped label stays centred on the row", () => {
    const scene = buildChart(dots("dot", [[12, 30]]));
    const max = labels(scene).find((l) => l.text === "30")!;
    const dot = named(scene, /^seg-0-1$/)[0] as EllipseNode;
    expect(max.align).toBe("right");
    expect(max.y + max.h / 2).toBeCloseTo(dot.cy, 9);
  });

  it("control: upright, the stem is vertical and the label hangs beside it, centred on the dot", () => {
    const scene = buildChart(dots("lollipop", [[12, 30]], [480, 300], false));
    for (const l of labels(scene)) {
      const dot = named(scene, new RegExp(`^seg-${l.name!.slice("label-".length)}$`))[0] as EllipseNode;
      expect(l.y + l.h / 2, l.name).toBeCloseTo(dot.cy, 9);
    }
  });

  it("where the row has no room to lift into, the label keeps its place on the line rather than vanishing", () => {
    // 14pt rows: a lifted label would need a 4.7pt font, below the floor.
    const scene = buildChart({ ...dots("lollipop", [[12, 30, 20, 5]], [480, 80]), title: undefined } as ChartConfig);
    const max = labels(scene).find((l) => l.text === "30")!;
    const dot = named(scene, /^seg-0-1$/)[0] as EllipseNode;
    expect(max).toBeDefined();
    expect(max.y + max.h / 2).toBeCloseTo(dot.cy, 9);
  });

  it("a lifted label is fitted to the room between its row and the next, and stays on the chart", () => {
    const scene = buildChart(dots("lollipop", [[12, 30, 20, 5]], [480, 120]));
    for (const l of labels(scene)) {
      const b = ink(l);
      expect(b.y0, l.name).toBeGreaterThanOrEqual(0);
      expect(b.y1, l.name).toBeLessThanOrEqual(scene.height);
    }
    expect(crossings(scene)).toEqual([]);
  });
});
