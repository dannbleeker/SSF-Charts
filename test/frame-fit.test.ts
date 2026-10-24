import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { fitPlot } from "../src/core/layout/frame";
import { polar } from "../src/core/geometry";
import { textWidth } from "../src/core/scene";
import type { SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { SCENE_NEUTRAL_OPTIONS, optionsThatChangeNothing, withOptionData } from "./helpers/sweep-options";

/**
 * A chart is a box on a slide, and nothing it draws may leave that box.
 *
 * Neither PowerPoint renderer wraps or clips a text box, so a label wider than
 * its frame is not truncated — it is drawn across whatever the chart happens to
 * be sitting next to. The same is true of a plot: a bar hanging below the frame
 * lands on the slide, not on the chart.
 *
 * The sweep is the point of this file. Every earlier fix in this area was found
 * by looking at one chart at one size, and each time the same defect was sitting
 * in four other layouts — the title alone overflowed 13 of 17 kinds. So the
 * property is asserted over the whole product of kinds and frame sizes, and the
 * unit tests below it name the mechanisms that sweep would otherwise only
 * report as a number.
 */

/** Half a point, for the rounding in the layouts' own arithmetic. */
const SLACK = 0.5;

/**
 * The ink a node actually puts on the slide — NOT its box.
 *
 * For text the two are very different: a label's box is routinely wider than
 * its text and is anchored by `align`, so measuring boxes reports overflow that
 * is not there (a right-aligned axis label whose box starts at x=0) and misses
 * overflow that is (a centred label in a 9pt box carrying 102pt of text). An
 * early version of this sweep measured boxes and produced four false positives
 * and one false negative in a single run.
 */
function inkBox(n: SceneNode): { x0: number; y0: number; x1: number; y1: number } | null {
  const a = n as unknown as Record<string, number>;
  if (n.kind === "text") {
    const t = n as TextNode;
    const w = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    // Baseline by vertical alignment, then the em box around it: cap height up,
    // descender down.
    const base =
      t.valign === "top"
        ? t.y + t.fontSize
        : t.valign === "bottom"
          ? t.y + t.h - t.fontSize * 0.25
          : t.y + t.h / 2 + t.fontSize * 0.36;
    return { x0: x, y0: base - t.fontSize * 0.8, x1: x + w, y1: base + t.fontSize * 0.21 };
  }
  if (n.kind === "rect" || n.kind === "chevron") return { x0: a.x, y0: a.y, x1: a.x + a.w, y1: a.y + a.h };
  if (n.kind === "line")
    return {
      x0: Math.min(a.x1, a.x2),
      y0: Math.min(a.y1, a.y2),
      x1: Math.max(a.x1, a.x2),
      y1: Math.max(a.y1, a.y2),
    };
  if (n.kind === "ellipse") return { x0: a.cx - a.rx, y0: a.cy - a.ry, x1: a.cx + a.rx, y1: a.cy + a.ry };
  if (n.kind === "wedge") {
    // THE ARC, not the disc that contains it. A wedge's bounding circle is a
    // conservative bound, and a conservative bound does not merely miss defects
    // — it invents them, which is a lesson this file already carries about the
    // arrowhead. A semi-doughnut gauge draws only the top half of its circle,
    // so the disc reported 291 points of "overflow" past the bottom of a
    // 960x540 chart where the ink stops at the centre line.
    //
    // The extremes of an arc are its two ends, the centre (a solid wedge is
    // drawn from there), the inner arc's ends where there is a hole, and
    // whichever of the four compass points the sweep happens to cross.
    const from = a.startAngle;
    const to = a.endAngle;
    const pts = [polar(a.cx, a.cy, a.r, from), polar(a.cx, a.cy, a.r, to)];
    if (a.innerR > 0) pts.push(polar(a.cx, a.cy, a.innerR, from), polar(a.cx, a.cy, a.innerR, to));
    else pts.push({ x: a.cx, y: a.cy });
    for (let q = -720; q <= 720; q += 90) if (q >= from && q <= to) pts.push(polar(a.cx, a.cy, a.r, q));
    return {
      x0: Math.min(...pts.map((p) => p.x)),
      y0: Math.min(...pts.map((p) => p.y)),
      x1: Math.max(...pts.map((p) => p.x)),
      y1: Math.max(...pts.map((p) => p.y)),
    };
  }
  if (n.kind === "symbol") return { x0: a.cx - a.size, y0: a.cy - a.size, x1: a.cx + a.size, y1: a.cy + a.size };
  // An arrowhead's (x, y) is its TIP and its body runs 1.8*size back along
  // `angle`. Read the angle: the disc that bounds it at EVERY rotation is what
  // this used to use, and a conservative bound does not merely miss defects, it
  // invents them — it reported the scatter's trajectory glyphs as 4.2pt off the
  // canvas while their actual triangles were inside it. The three vertices are
  // the ones the SVG renderer draws, which is the shape the scene describes;
  // both PowerPoint sinks name a slightly broader native preset, and that
  // difference is declared in the parity contract in `src/core/scene.ts`.
  if (n.kind === "arrowhead") {
    const rad = (a.angle * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const vs = [
      [0, 0],
      [-1.8 * a.size, -0.7 * a.size],
      [-1.8 * a.size, 0.7 * a.size],
    ].map(([px, py]) => ({ x: a.x + px * cos - py * sin, y: a.y + px * sin + py * cos }));
    return {
      x0: Math.min(...vs.map((v) => v.x)),
      y0: Math.min(...vs.map((v) => v.y)),
      x1: Math.max(...vs.map((v) => v.x)),
      y1: Math.max(...vs.map((v) => v.y)),
    };
  }
  if (n.kind === "polygon") {
    const p = (n as { points?: { x: number; y: number }[] }).points ?? [];
    if (!p.length) return null;
    return {
      x0: Math.min(...p.map((q) => q.x)),
      y0: Math.min(...p.map((q) => q.y)),
      x1: Math.max(...p.map((q) => q.x)),
      y1: Math.max(...p.map((q) => q.y)),
    };
  }
  return null;
}

/**
 * The worst distance any node's ink reaches past an edge of the chart, and which
 * node. `only` restricts it to the nodes it accepts by name, which is how an
 * overflow recorded as open (see `KNOWN_OPEN`) is kept out of a sweep without
 * keeping the rest of that chart out with it.
 */
function worstOverflow(
  cfg: ChartConfig,
  only: (name: string) => boolean = () => true,
): { pt: number; node: string; side: string } {
  const scene = buildChart(cfg);
  let pt = 0;
  let node = "";
  let side = "";
  for (const n of scene.nodes) {
    if (!only(n.name ?? n.kind)) continue;
    const b = inkBox(n);
    if (!b) continue;
    const sides: [string, number][] = [
      ["left", -b.x0],
      ["top", -b.y0],
      ["right", b.x1 - scene.width],
      ["bottom", b.y1 - scene.height],
    ];
    for (const [which, over] of sides) {
      if (over > pt) {
        pt = over;
        node = n.name ?? n.kind;
        side = which;
      }
    }
  }
  return { pt, node, side };
}

/**
 * Frames a chart is asked to draw in. The middle three are the sizes the pane's
 * gallery and the deck's thumbnails use; the outliers are there because every
 * defect this file guards against was invisible at a comfortable size and
 * obvious at a cramped one — and because a chart's size is a number a caller
 * types, so nothing stops one arriving.
 */
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

describe("no chart draws outside its own frame", () => {
  for (const [w, h] of FRAMES) {
    it(`every kind fits ${w}x${h}`, () => {
      const bad: string[] = [];
      for (const { kind } of CHART_KINDS) {
        const o = worstOverflow({ ...sampleConfig(kind), width: w, height: h } as ChartConfig);
        if (o.pt > SLACK) bad.push(`${kind}: ${o.node} ${o.pt.toFixed(1)}pt past the ${o.side}`);
      }
      expect(bad).toEqual([]);
    });
  }
});

/**
 * The same property over the DECORATIONS, which both sweeps above hold at
 * whatever the sample happens to switch on.
 *
 * This gate claims to cover charts and covered `sampleConfig(kind)` — so every
 * decoration a sample does not use was swept by nothing at all, on a product
 * whose whole selling point is think-cell's annotations. That is the mistake
 * this file has already recorded twice under other names: a gate whose name is
 * wider than its coverage.
 *
 * Turning each one on in turn found five defect families and one of them at an
 * ORDINARY size, which is the part worth remembering — 480x300 is not a corner:
 *
 *   - a difference arrow on a kind with no category axis (`categoryX` is a
 *     placeholder there, every category at the middle of the plot) put itself
 *     ten points past the right edge and labelled a zero-length span;
 *   - a callout whose text is wider than the chart was CENTRED by a clamp
 *     written for the vertical case, so it hung off both edges at once;
 *   - a quadrant label took `Math.max(20, zone - 8)` — a floor that RAISES a
 *     width, the scatter legend's own bug in another layout;
 *   - a band's label was priced in font sizes with nothing bounding it;
 *   - a lollipop's value label never got the row-pitch bound the rest of the
 *     horizontal labels were given.
 *
 * Every value below is a shape a caller can write today and the manual
 * documents. The sweep is two fonts rather than seven: the defects it found
 * were bounds that were missing outright, and the font axis is already swept in
 * full above.
 */
const DECORATIONS: Record<string, unknown> = {
  segmentLabels: true,
  seriesLabels: true,
  totals: true,
  grandTotal: true,
  variance: { actual: 1, reference: 0 },
  categoryAxis: true,
  valueAxis: true,
  gridlines: true,
  tickMode: "data",
  gridShape: "circle",
  fillOpacity: 0.5,
  labelContent: ["value", "percent", "series", "category"],
  cagr: { from: 0, to: 2 },
  difference: { from: 0, to: 2 },
  valueLines: [{ mode: "mean" }],
  valueLine: { mode: "mean" },
  connectors: true,
  // A sentence somebody typed, which is what a callout is — and at the default
  // font it is wider than three of the frames below.
  callouts: [{ text: "A rather long callout label", category: 1 }],
  bands: [{ axis: "y", from: 0, to: 50, label: "Target band" }],
  hundredPercentNote: true,
  forecastFrom: 1,
  quadrants: { x: 50, y: 50, labels: ["Top left quadrant", "Top right", "Bottom left", "Bottom right"] },
  marginals: "both",
  barStyle: "lollipop",
  fillBetween: [0, 1],
  stepped: "after",
  smooth: true,
  bridgeGaps: true,
  bump: true,
  slope: true,
  trajectory: true,
  summaryBars: true,
  criticalPath: true,
  sparkline: true,
};

/**
 * And the same over the TOP-LEVEL options, which are neither a decoration nor a
 * font.
 *
 * `sampleConfig(kind)` sets a handful of them and the sweeps above inherit
 * whatever it happens to switch on, so `axisBreak`, `multiples`, `pareto`,
 * `pie.semi`, `pie.explode`, `logScale`, `secondaryAxis` and the rest were swept
 * by nothing — the same hole the decorations were in, one level up. Seven more
 * defect families, and the two largest are not small-frame corners:
 *
 *   - the SEMI-doughnut gauge drew its labels through a clamp that cannot place
 *     a box wider than the frame, its leader lines past the arc's own edge, and
 *     its headline total below the chart;
 *   - an EXPLODED slice moves `r * 0.08` along its midline and takes its outside
 *     label with it, past a band sized for the un-exploded ring — 9.6pt below a
 *     960x540 doughnut;
 *   - `axisBreak`'s zigzag overhangs the plot by a fixed 4 points, which is more
 *     than the plot's own inset on a small chart;
 *   - `pareto`'s secondary axis labels are centred on their ticks with nothing
 *     stopping the topmost one leaving the chart;
 *   - `multiples` divides the frame into one chart per series, so a titled
 *     300x60 chart asks for a 150x25 panel — a size no sweep here had ever
 *     asked for, whose own TITLE then overflowed, and at three panels a 7.6pt
 *     one whose category axis was drawn 21.6pt below the chart;
 *   - a `footnote` or an explicit `scale` squeezes the plot until a line's or a
 *     combo's markers, drawn AROUND their points, cross the frame's edge.
 *
 * The wedge measurement had to be fixed before any of this could be read: see
 * `inkBox`, which used to bound a wedge by the disc that contains it and
 * therefore reported 291pt of "overflow" on a gauge that draws half a circle.
 */
const OPTIONS: Record<string, Record<string, unknown>> = {
  scale: { scale: { min: -50, max: 500 } },
  segmentOrder: { segmentOrder: "reverse" },
  categorySort: { categorySort: "descending" },
  secondaryAxis: { secondaryAxis: true },
  axisBreak: { axisBreak: { from: 20, to: 60 } },
  valueAxisTitle: { valueAxisTitle: "Revenue in millions of euro" },
  logScale: { logScale: true },
  gapWidth: { gapWidth: 0.8 },
  overlap: { overlap: 0.5 },
  footnote: { footnote: "Source: an internal model, restated for the 2024 perimeter" },
  pareto: { pareto: true },
  "multiples 2 columns": { multiples: { columns: 2 } },
  "multiples 3 columns": { multiples: { columns: 3 } },
  otherBucket: { otherBucket: { max: 3 } },
  "pie.explode": { pie: { explode: [0, 1] } },
  "pie.semi": { pie: { semi: true } },
  "pie.variableRadius": { pie: { variableRadius: true } },
  "tilemap.hex": { tilemap: { shape: "hex" } },
  "tilemap.glyph": { tilemap: { glyph: "bars" } },
  "butterfly.split": { butterfly: { split: 1 } },
  boxplot: { boxplot: { notch: true, showMean: true, jitter: true } },
  "render image": { render: "image" },
  numberFormat: { numberFormat: { locale: "de-DE", currency: "EUR", decimals: 2 } },
  title: { title: "A rather long chart title that names the measure and the period" },
};

/**
 * EVERY ENTRY ABOVE HAS TO DO SOMETHING TO THE CHART IT IS SWEPT OVER.
 *
 * It is laid over `sampleConfig(kind)` by a shallow spread, so an option that
 * needs data the sample does not have changed nothing: six of the entries did
 * not (`withOptionData` gives five of them the data they need), and the one
 * this was found through, `tilemap.glyph`, was green over 32 cells that drew no
 * glyph at all — it passed with the glyph bars thrown a chart-width off the
 * frame. The check below is that none of them is inert, so a table entry cannot
 * go back to meaning nothing without a test naming it.
 */
describe("every option in the sweep takes effect", () => {
  it("changes the scene of at least one kind", () => {
    const inert = optionsThatChangeNothing(OPTIONS).filter((key) => !(key in SCENE_NEUTRAL_OPTIONS));
    expect(
      inert,
      "these options are swept over a sample they do nothing to, so the sweep passes them for free — see `OPTION_DATA` in test/helpers/sweep-options.ts",
    ).toEqual([]);
  });

  it("leaves only the scene-neutral ones alone, so the exemption cannot go stale", () => {
    const inert = optionsThatChangeNothing(OPTIONS);
    expect(
      Object.keys(SCENE_NEUTRAL_OPTIONS).filter((key) => !inert.includes(key)),
      "an option declared scene-neutral now changes a scene: sweep it like the others",
    ).toEqual([]);
  });
});

/**
 * OVERFLOWS THE SWEEP FOUND WHEN TWO OF ITS OPTIONS BEGAN TO TAKE EFFECT, which no
 * layout has been fixed for yet — kept out of the sweep by NODE, never by option,
 * so the rest of those charts is still measured, and kept honest by the test
 * after this one, which fails the day each stops overflowing.
 *
 * It is the series legend of a mode the sweep had never drawn, and the same
 * mistake the sweep exists to catch:
 *
 *   - `butterfly.split`: with the flanks stacked, `layoutButterfly` reserves
 *     header rows for the legend by `legendRowCount(…, valueW, width - 4)` and
 *     then draws it from `plot.x`, so the legend wraps onto more rows than it
 *     reserved and the last is drawn up to 55pt below the chart; 6 of 32.
 *
 * It does not show at 480x300 or 960x540, the sizes a chart is normally placed
 * at, which is why this is recorded and not blocking. Fix the layout, delete the
 * entry, and the sweep covers it.
 *
 * (`tilemap.glyph` was the other: the glyph legend advanced a chip, a gap and a
 * measured name per series along one row and never asked whether the next fit.
 * It wraps now, so its entry is gone and the sweep holds the whole chart.)
 */
const KNOWN_OPEN: Record<string, { nodes: RegExp; what: string }> = {
  "butterfly.split": {
    nodes: /^legend(-chip)?-\d+$/,
    what: "the stacked legend wraps below the chart",
  },
};

/** The nodes of `key`'s sweep that are held to the frame: everything but its recorded open overflow. */
const outsideKnownOpen = (key: string) => (name: string) => !KNOWN_OPEN[key]?.nodes.test(name);

/**
 * And the third axis: the DATA, which every sweep here holds at whatever
 * `sampleConfig` ships — four to six short category names, one to three series,
 * two-digit values.
 *
 * A user's data is none of those things. This transforms the sample rather than
 * replacing it, so each kind keeps the shape its layout requires (a scatter's
 * points, a gantt's tasks, a heatmap's matrix) while the thing under test moves.
 *
 * It found ONE family, which is worth as much as a long list: a funnel with 24
 * stages on a 60pt-tall chart drew its last stage 10.4pt below the chart,
 * because `Math.max(1, …)` floored each band's height and 24 floored bands plus
 * their gaps measured 36 points of a 24-point plot. The floor comes out of the
 * GAP now — the bands are the chart, the gaps are chrome for a label — and where
 * even hairlines will not fit, the floor is abandoned rather than the stages.
 */
const DATA_SHAPES: Record<string, (c: Record<string, any>) => Record<string, any>> = {
  "long category names": (c) => ({
    ...c,
    data: { ...c.data, categories: c.data.categories.map((x: string, i: number) => `${x} enterprise segment ${i}`) },
  }),
  "long series names": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: Record<string, any>) => ({ ...s, name: `${s.name} including allocations` })),
    },
  }),
  "24 categories": (c) => ({
    ...c,
    data: {
      ...c.data,
      categories: Array.from({ length: 24 }, (_, i) => `C${i + 1}`),
      series: c.data.series.map((s: Record<string, any>) => ({
        ...s,
        values: Array.from({ length: 24 }, (_, i) => (s.values[i % s.values.length] ?? 0) + i),
      })),
    },
  }),
  "10 series": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: Array.from({ length: 10 }, (_, i) => ({
        ...c.data.series[i % c.data.series.length],
        name: `Series ${i + 1}`,
        values: c.data.series[i % c.data.series.length].values.map((v: number | null) => (v == null ? v : v + i)),
      })),
    },
  }),
  "one category": (c) => ({
    ...c,
    data: {
      ...c.data,
      categories: c.data.categories.slice(0, 1),
      series: c.data.series.map((s: Record<string, any>) => ({ ...s, values: s.values.slice(0, 1) })),
    },
  }),
  "one series": (c) => ({ ...c, data: { ...c.data, series: c.data.series.slice(0, 1) } }),
  "values in the billions": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: Record<string, any>) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : v * 1234567)),
      })),
    },
  }),
  "every value negative": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: Record<string, any>) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : -v)),
      })),
    },
  }),
  "mixed signs": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: Record<string, any>) => ({
        ...s,
        values: s.values.map((v: number | null, i: number) => (v == null ? v : i % 2 ? -v : v)),
      })),
    },
  }),
  "tiny fractions": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: Record<string, any>) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : v / 100000)),
      })),
    },
  }),
};

describe("no data shape draws outside the chart's frame", () => {
  for (const [key, shape] of Object.entries(DATA_SHAPES)) {
    it(`${key} stays inside every frame`, () => {
      const bad: string[] = [];
      for (const [w, h] of FRAMES) {
        for (const fontSize of [10, 18]) {
          for (const horizontal of [false, true]) {
            for (const { kind } of CHART_KINDS) {
              const cfg = {
                ...shape(sampleConfig(kind) as unknown as Record<string, any>),
                width: w,
                height: h,
                horizontal,
                style: { fontSize },
              } as ChartConfig;
              const o = worstOverflow(cfg);
              if (o.pt > SLACK)
                bad.push(
                  `${kind} ${w}x${h}@${fontSize}${horizontal ? " rotated" : ""}: ${o.node} ${o.pt.toFixed(1)}pt past the ${o.side}`,
                );
            }
          }
        }
      }
      expect(bad).toEqual([]);
    });
  }
});

describe("no top-level option draws outside the chart's frame", () => {
  for (const [key, value] of Object.entries(OPTIONS)) {
    it(`${key} stays inside every frame`, () => {
      const bad: string[] = [];
      for (const [w, h] of FRAMES) {
        for (const fontSize of [10, 18]) {
          for (const horizontal of [false, true]) {
            for (const { kind } of CHART_KINDS) {
              const cfg = {
                ...withOptionData(key, sampleConfig(kind)),
                width: w,
                height: h,
                horizontal,
                style: { fontSize },
                ...value,
              } as ChartConfig;
              const o = worstOverflow(cfg, outsideKnownOpen(key));
              if (o.pt > SLACK)
                bad.push(
                  `${kind} ${w}x${h}@${fontSize}${horizontal ? " rotated" : ""}: ${o.node} ${o.pt.toFixed(1)}pt past the ${o.side}`,
                );
            }
          }
        }
      }
      expect(bad).toEqual([]);
    });
  }
});

describe("the overflows recorded as open are still open", () => {
  for (const [key, open] of Object.entries(KNOWN_OPEN)) {
    it(`${key}: ${open.what}`, () => {
      let worst = 0;
      for (const [w, h] of FRAMES)
        for (const fontSize of [10, 18])
          for (const horizontal of [false, true])
            for (const { kind } of CHART_KINDS) {
              const cfg = {
                ...withOptionData(key, sampleConfig(kind)),
                width: w,
                height: h,
                horizontal,
                style: { fontSize },
                ...OPTIONS[key],
              } as ChartConfig;
              worst = Math.max(worst, worstOverflow(cfg, (name) => open.nodes.test(name)).pt);
            }
      expect(
        worst,
        `${key} no longer overflows through ${open.nodes}: delete its KNOWN_OPEN entry so the sweep holds it to the frame`,
      ).toBeGreaterThan(SLACK);
    });
  }
});

describe("no decoration draws outside the chart's frame", () => {
  for (const [key, value] of Object.entries(DECORATIONS)) {
    it(`${key} stays inside every frame`, () => {
      const bad: string[] = [];
      for (const [w, h] of FRAMES) {
        for (const fontSize of [10, 18]) {
          for (const horizontal of [false, true]) {
            for (const { kind } of CHART_KINDS) {
              const base = sampleConfig(kind) as unknown as Record<string, unknown>;
              const cfg = {
                ...base,
                width: w,
                height: h,
                horizontal,
                style: { fontSize },
                decorations: { ...(base.decorations as object), [key]: value },
              } as ChartConfig;
              const o = worstOverflow(cfg);
              if (o.pt > SLACK)
                bad.push(
                  `${kind} ${w}x${h}@${fontSize}${horizontal ? " rotated" : ""}: ${o.node} ${o.pt.toFixed(1)}pt past the ${o.side}`,
                );
            }
          }
        }
      }
      expect(bad).toEqual([]);
    });
  }
});

/**
 * The same property over the FONT, which the sweep above holds fixed.
 *
 * Every layout prices its chrome in font sizes, so the font is the other axis a
 * chart can be squeezed along — and it was never swept. Seven overflows were
 * sitting at 24 and 32pt, all of one shape: a label centred on a row, a ring or
 * a legend line, in a box `fontSize * 1.2` to `* 1.5` tall. Once the font
 * outgrew the spacing, the labels overlapped EACH OTHER at any frame size and
 * the last one left the chart at a small one. `CLAUDE.md` records the funnel
 * being fixed for exactly this; the butterfly, the gantt, the radar's ticks and
 * the scatter's legend had never been.
 *
 * A big font on a small chart is not a corner case anyone should have to defend:
 * `style.fontSize` is a number a caller types.
 */
describe("no chart draws outside its own frame at any font", () => {
  const FONTS = [6, 8, 10, 14, 18, 24, 32];
  for (const fontSize of FONTS) {
    it(`every kind at ${fontSize}pt`, () => {
      const bad: string[] = [];
      for (const [w, h] of FRAMES) {
        for (const { kind } of CHART_KINDS) {
          const o = worstOverflow({ ...sampleConfig(kind), width: w, height: h, style: { fontSize } } as ChartConfig);
          if (o.pt > SLACK) bad.push(`${kind} at ${w}x${h}: ${o.node} ${o.pt.toFixed(1)}pt past the ${o.side}`);
        }
      }
      expect(bad).toEqual([]);
    });
  }
});

/**
 * Text that is drawn ON TOP of other text.
 *
 * A label inside the frame can still be unreadable, and nothing here could see
 * that: the frame sweeps above pass a chart whose every label is stacked on its
 * neighbour. A sweep of the ink boxes against each other found 73 kind/font/frame
 * combinations with overlapping text, and one shape accounted for 30 of them —
 * adjacent CATEGORY AXIS labels, which is one defect rather than seven because
 * that axis is shared by every cartesian kind.
 *
 * This gate was deliberately narrow — it pinned the two defects that were fixed
 * rather than asserting no chart anywhere overlaps, which was not true yet. It
 * is no longer narrow in the way that matters: as of 2026-08-27 it carries NO
 * exception, so within the frames and fonts it sweeps, no chart draws any text
 * over any other text. What it still does not sweep is option and data-shape
 * VARIANTS of each kind, which is where the "24 categories" family was found and
 * measured; see docs/BACKLOG.md.
 */
describe("labels are not drawn on top of each other", () => {
  /**
   * THE ONE OVERLAP THIS GATE ALLOWED, and there is no exception here now —
   * 2026-08-27. Both sweeps below assert against EVERY pair.
   *
   * A scatter or bubble point label was permitted to touch an axis TICK number.
   * Two fixes had been tried and measured and both were refused, each removing
   * the overlaps by DROPPING point labels — 56 of 301 on charts as roomy as
   * 480x300 — because the y axis owns the left margin. A point's label is data
   * and a tick number is chrome, so the trade was refused.
   *
   * The verdict was right and the conclusion drawn from it was not. "The chrome
   * yields" does not mean both are drawn through each other; it means the tick
   * number goes. It does now — `tickLabelsUnderPointLabels` — and placement runs
   * a dodging pass first so the axis keeps 341 of its 423 numbers rather than
   * 310. Point labels are not paid: 266 become 265 on a scatter, 265 become 267
   * on a bubble.
   *
   * The note is kept although the code is gone, because HOW the exception was
   * written is worth more than the exception was: it stated itself, rather than
   * being a frame the gate quietly avoided. The arrangement before it was the
   * second kind, and it left 146 real overlapping pairs outside the two frames
   * it happened to check.
   */
  // (no exception function — the sweeps below allow nothing.)

  /**
   * Every font this engine claims to draw at, 6 to 32.
   *
   * 24 and 32 were outside this gate until 2026-08-19 on the argument that at
   * that size the chrome genuinely exceeds the frame, so the remaining pairs
   * were a decision about how small a chart SSF Charts claims to draw rather
   * than bounds anyone had forgotten. Measured, that turned out to be wrong:
   * 185 pairs, 165 of them on ONE frame (300x60), and the shape was not a chart
   * with more chrome than room — it was `fitPlot` growing a squeezed plot UP
   * from its bottom edge and carrying every band placed from that edge into the
   * title. The bands yield to the title now (`printsOnTitle`), and no
   * `MIN_READABLE` ratio was needed.
   */
  const FONTS_NEAR_DEFAULT = [6, 8, 10, 14, 18, 24, 32];

  /** The ink of a text node, as the frame sweep measures it. */
  const inkOf = (t: TextNode) => inkBox(t)!;
  const overlap = (a: ReturnType<typeof inkOf>, b: ReturnType<typeof inkOf>) => {
    const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
    const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
    return w > 0 && h > 0 ? w * h : 0;
  };
  const textsNamed = (cfg: ChartConfig, re: RegExp) =>
    buildChart(cfg).nodes.filter(
      (n): n is TextNode => n.kind === "text" && !!n.name && re.test(n.name) && !!n.text.trim(),
    );

  it("the shared category axis fits each name to its slot", () => {
    // Every label is centred in a slot and none was fitted to it, so a name
    // wider than its slot ran into its neighbours on both sides.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const { kind } of CHART_KINDS) {
        const cfg = { ...sampleConfig(kind), width: 200, height: 150, style: { fontSize } } as ChartConfig;
        const cats = textsNamed(cfg, /^category-\d+$/);
        for (let i = 1; i < cats.length; i++) {
          if (overlap(inkOf(cats[i - 1]), inkOf(cats[i])) > 1) {
            bad.push(`fs=${fontSize} ${kind}: ${cats[i - 1].name} over ${cats[i].name}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("no chart overlaps its own text at any font it draws", () => {
    // The strongest form this file can assert today. A sweep of the text ink
    // boxes against each other found 237 overlapping pairs across the kinds and
    // fonts; fitting each label to the space it actually has took that to 76,
    // and to ZERO at the font every chart is drawn at unless someone changes it.
    //
    // The FONTS EITHER SIDE of the default came next, and they were not a
    // formality: 98 pairs at 6, 8, 14 and 18pt, over eleven kinds. Most of them
    // were one shape — a label drawn above a mark, clamped to y=0 when the mark
    // reached the top of the plot, which is where the TITLE is (see
    // `aboveMarkFontSize`). The rest were fits that stopped at 6pt whatever the
    // band could hold (the funnel's rows), a legend clamped back into the grid
    // it explains (the heatmap's), a size key pushed off one edge to clear the
    // other (the bubble's), a tick fit measuring a gap the rings did not have
    // (the radar's), and the two corner labels where the x axis meets the y.
    //
    // 24 and 32pt joined it on 2026-08-19, and what they held was not what the
    // note here predicted: 185 pairs, 165 of them at 300x60 alone, and nearly
    // all of them one mechanism — the category strip, the axis ticks and the
    // line's name climbing into the TITLE's band behind a plot `fitPlot` had
    // squeezed. `printsOnTitle` is the shared answer; the title is what chrome
    // yields to, because it is the one label that says what the chart is.
    const bad: string[] = [];
    for (const fontSize of FONTS_NEAR_DEFAULT)
      for (const [w, h] of [
        // 160x120 was added after a sweep found this gate was checking TWO of the
        // eight frames the overflow sweep uses, while its name claims the default
        // font generally. See the block comment above.
        //
        // 60x300 followed: a tall sidebar is an ordinary shape for a chart on a
        // slide, and it carried 16 overlapping pairs at the default font — a
        // heatmap's column headers drawn across each other, a butterfly's two
        // series names meeting in the middle, a combo point label pinned to x=0
        // over the category names, a bubble size legend spilling onto the y axis.
        //
        // 120x90 followed once the sunburst's adjacent OUTSIDE labels were fitted
        // to the vertical gap between neighbours instead of to each wedge's own
        // arc. That one pair was the only thing left on this frame upright — it is
        // NOT in the rotated sweep below, which still has four (see there).
        //
        // 300x60 and 80x60 closed last, and each needed a different bound rather
        // than the de-collision work this comment used to predict: the upright
        // column totals fitted to their category slot, the combo line's series
        // name fitted to the gutter it actually has instead of a width floor it
        // does not, and the CAGR caption shrunk-then-dropped against the TITLE's
        // ink (clamping it to the title's bottom was tried and measured and is
        // still refused — it moves the overlap onto the totals).
        //
        // Every frame in the overflow sweep is now covered here, in both
        // orientations.
        [60, 300],
        [80, 60],
        [120, 90],
        [300, 60],
        [160, 120],
        [200, 150],
        [480, 300],
      ] as [number, number][]) {
        for (const { kind } of CHART_KINDS) {
          const ts = buildChart({
            ...sampleConfig(kind),
            width: w,
            height: h,
            style: { fontSize },
          } as ChartConfig).nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text.trim());
          const boxes = ts.map((t) => inkBox(t)!);
          for (let i = 0; i < boxes.length; i++) {
            for (let j = i + 1; j < boxes.length; j++) {
              if (overlap(boxes[i], boxes[j]) > 1) {
                bad.push(`${kind} at ${w}x${h} ${fontSize}pt: ${ts[i].name} over ${ts[j].name}`);
              }
            }
          }
        }
      }
    expect(bad).toEqual([]);
  });

  it("fits the X axis to the WIDTH of its numbers, not to a line height", () => {
    // The y axis and the x axis were fitted by one rule, and only one of them is
    // about a line height. Down the Y the labels stack, so the gap between ticks
    // has to hold `fs * 1.4`. Across the X they sit side by side, so what the gap
    // has to hold is the WIDEST LABEL — and `1,234,567,890` is 60 points wide at
    // the default font, which sails through a 14-point test and is then drawn
    // straight through its neighbour.
    //
    // It arrives with the MAGNITUDE of the data rather than with the frame,
    // which is why every sweep in this file missed it: `sampleConfig` deals in
    // two-digit numbers.
    const bad: string[] = [];
    for (const kind of ["scatter", "bubble"] as const) {
      const base = sampleConfig(kind) as unknown as Record<string, any>;
      const cfg = {
        ...base,
        width: 200,
        height: 150,
        data: {
          ...base.data,
          series: base.data.series.map((s: Record<string, any>) => ({
            ...s,
            values: s.values.map((v: number | null) => (v == null ? v : v * 1234567)),
            points: s.points?.map((p: Record<string, number>) => ({ ...p, x: p.x * 1234567 })),
          })),
        },
      } as ChartConfig;
      const ts = textsNamed(cfg, /^x-axis$/);
      for (let i = 1; i < ts.length; i++)
        if (overlap(inkOf(ts[i - 1]), inkOf(ts[i])) > 1) bad.push(`${kind}: x-axis pair ${i} (${ts[i].text})`);
    }
    expect(bad).toEqual([]);
  });

  it("axis tick labels are fitted to the spacing between ticks", () => {
    // One label per tick, each centred on its own tick, and none was fitted to
    // the gap — 71 of the 237 pairs were axis labels drawn through each other,
    // the worst shape in the engine after the category axis.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const kind of ["stacked", "line", "scatter", "bubble", "boxplot"] as const) {
        const cfg = { ...sampleConfig(kind), width: 200, height: 150, style: { fontSize } } as ChartConfig;
        for (const axis of [/^value-axis$/, /^y-axis$/, /^x-axis$/]) {
          const ts = textsNamed(cfg, axis);
          for (let i = 1; i < ts.length; i++) {
            if (overlap(inkOf(ts[i - 1]), inkOf(ts[i])) > 1)
              bad.push(`fs=${fontSize} ${kind}: ${ts[i].name} pair ${i}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a heatmap's row labels are fitted to their row", () => {
    // The label's BOX is the row so the boxes never overlap, but its INK is the
    // font and is centred in that box — so once the font outgrew the row the
    // names were drawn through each other. Twenty pairs, and invisible to a
    // box-based check, which is why this file measures ink.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      const ts = textsNamed(
        { ...sampleConfig("heatmap"), width: 200, height: 150, style: { fontSize } } as ChartConfig,
        /^row-\d+$/,
      );
      for (let i = 1; i < ts.length; i++) {
        if (overlap(inkOf(ts[i - 1]), inkOf(ts[i])) > 1)
          bad.push(`fs=${fontSize}: ${ts[i - 1].name} over ${ts[i].name}`);
      }
    }
    expect(bad).toEqual([]);
  });

  it("a sunburst's inside labels are fitted to their wedge", () => {
    // Drawn verbatim at the ring font and bounded by nothing, so a group name
    // wider than its own arc ran into the next group's — the same defect the
    // pie's inside labels had, in the same shape. The bound is the chord of the
    // wedge at the label's radius.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      const ts = textsNamed(
        { ...sampleConfig("sunburst"), width: 200, height: 150, style: { fontSize } } as ChartConfig,
        /^group-label-\d+$/,
      );
      for (let i = 0; i < ts.length; i++) {
        for (let j = i + 1; j < ts.length; j++) {
          if (overlap(inkOf(ts[i]), inkOf(ts[j])) > 1) bad.push(`fs=${fontSize}: ${ts[i].name} over ${ts[j].name}`);
        }
      }
    }
    expect(bad).toEqual([]);
  });

  /**
   * The same property, sideways — which for a year nothing checked.
   *
   * Every fit above was written for the upright chart and several of the
   * horizontal twins had none at all: `horizontalChrome` drew its category names
   * at the chart font in a box `fs * 1.5` tall centred on the row, the mekko's
   * own fit answered `false` outright when horizontal, and the totals and the
   * combo's point labels were bounded by nothing. Rotating a chart rotates which
   * side of a label is crowded; it does not excuse the label from being fitted.
   *
   * Zero, the same as upright — which was not a foregone conclusion when this
   * was written: a coarser estimate of text width put two scatter point labels
   * in the x-axis strip, and the engine's own `textWidth` says they clear it.
   * Measure with the metric the layouts measure with.
   */
  it("no horizontal chart overlaps its own text at any font it draws", () => {
    // Widened with the upright sweep above, and it found ONE node: the combo
    // line's series name, drawn at the full chart font while every label around
    // it had been fitted to its row — 22 pairs across four fonts, all of them
    // "Margin %" lying across the numbers of two or three categories. Sideways
    // that name has nowhere to be nudged to (every row is occupied), so it takes
    // half a row and is dropped below the floor.
    const bad: string[] = [];
    for (const fontSize of FONTS_NEAR_DEFAULT)
      for (const [w, h] of [
        // 160x120 was added after a sweep found this gate was checking TWO of the
        // eight frames the overflow sweep uses, while its name claims the default
        // font generally. See the block comment above.
        //
        // 60x300 followed: a tall sidebar is an ordinary shape for a chart on a
        // slide, and it carried 16 overlapping pairs at the default font — a
        // heatmap's column headers drawn across each other, a butterfly's two
        // series names meeting in the middle, a combo point label pinned to x=0
        // over the category names, a bubble size legend spilling onto the y axis.
        //
        // 120x90, 300x60 and 80x60 all joined once the horizontal mekko's legend
        // was gated on `horizontalLegendFits` — the predicate `computeFrameHorizontal`
        // and `horizontalChrome` already shared, which the mekko's own draw never
        // asked. On a frame the predicate refuses, the reservation was zero rows
        // and the legend was drawn into it anyway; that is the state this repo
        // measured as worse than not gating either side.
        //
        // Note this list is now WIDER than the upright one, which stops at 120x90.
        // Rotating a chart rotates which side of a label is crowded, so the two
        // sweeps are not expected to cover the same frames — what is left upright
        // at 300x60 and 80x60 is de-collision work on other kinds entirely.
        [60, 300],
        [80, 60],
        [120, 90],
        [300, 60],
        [160, 120],
        [200, 150],
        [480, 300],
      ] as [number, number][]) {
        for (const { kind } of CHART_KINDS) {
          const ts = buildChart({
            ...sampleConfig(kind),
            width: w,
            height: h,
            horizontal: true,
            style: { fontSize },
          } as ChartConfig).nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text.trim());
          const boxes = ts.map((t) => inkBox(t)!);
          for (let i = 0; i < boxes.length; i++) {
            for (let j = i + 1; j < boxes.length; j++) {
              if (overlap(boxes[i], boxes[j]) > 1)
                bad.push(`${kind} at ${w}x${h} ${fontSize}pt: ${ts[i].name} over ${ts[j].name}`);
            }
          }
        }
      }
    expect(bad).toEqual([]);
  });

  it("a horizontal chart's category names and totals fit the row they belong to", () => {
    // Both are centred on their row in a box `fs * 1.5` tall and neither was
    // bound by the row PITCH, so at 18pt on a 200x150 frame every name ran into
    // its neighbours — nine kinds, one shape, because they share this chrome.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const { kind } of CHART_KINDS) {
        const cfg = {
          ...sampleConfig(kind),
          width: 200,
          height: 150,
          horizontal: true,
          style: { fontSize },
        } as ChartConfig;
        for (const re of [/^category-\d+$/, /^total-\d+$/]) {
          const ts = textsNamed(cfg, re);
          for (let i = 1; i < ts.length; i++) {
            if (overlap(inkOf(ts[i - 1]), inkOf(ts[i])) > 1)
              bad.push(`fs=${fontSize} ${kind}: ${ts[i - 1].name} over ${ts[i].name}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a waterfall's value labels fit the slot their bar owns", () => {
    // Upright the box was `colThick + 12` however wide the slot was, so every
    // label could bleed 6pt into each neighbour; sideways it was `fs * 1.5` tall
    // on a row that may be thinner than that. Both orientations, one bound.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const horizontal of [false, true]) {
        const ts = textsNamed(
          { ...sampleConfig("waterfall"), width: 200, height: 150, horizontal, style: { fontSize } } as ChartConfig,
          /^label-\d+$/,
        );
        for (let i = 0; i < ts.length; i++) {
          for (let j = i + 1; j < ts.length; j++) {
            if (overlap(inkOf(ts[i]), inkOf(ts[j])) > 1)
              bad.push(`fs=${fontSize} h=${horizontal}: ${ts[i].name} over ${ts[j].name}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a combo line's point labels fit their row when the chart is horizontal", () => {
    // Sideways a point label is centred on its category's row rather than
    // floating above its point, and it was bounded by nothing — the largest
    // single group of text collisions in the engine when this was measured.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      const ts = textsNamed(
        { ...sampleConfig("combo"), width: 200, height: 150, horizontal: true, style: { fontSize } } as ChartConfig,
        /^combo-label-\d+-\d+$/,
      );
      for (let i = 0; i < ts.length; i++) {
        for (let j = i + 1; j < ts.length; j++) {
          if (overlap(inkOf(ts[i]), inkOf(ts[j])) > 1) bad.push(`fs=${fontSize}: ${ts[i].name} over ${ts[j].name}`);
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("no tick label is drawn too small to read — it is dropped instead", () => {
    // Every axis fit shrinks a label to the room it has and none had a floor, so
    // a plot squeezed to a few points answered with ONE-POINT labels: ink rather
    // than text, stacked in the gutter, from a fit reporting success. Six y ticks
    // 1.6pt apart on a 200x150 scatter at a 26pt font were the case that named
    // it. Asserted over every kind, since the floor is shared chrome.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const { kind } of CHART_KINDS) {
        for (const horizontal of [false, true]) {
          const cfg = {
            ...sampleConfig(kind),
            width: 200,
            height: 150,
            horizontal,
            style: { fontSize },
          } as ChartConfig;
          for (const t of textsNamed(cfg, /^(value-axis|x-axis|y-axis)$/)) {
            if (t.fontSize < 5) bad.push(`fs=${fontSize} ${kind} h=${horizontal}: ${t.name} at ${t.fontSize}pt`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("keeps the series labels and the legend out of the title band", () => {
    // Both settle into the gutter above or beside the plot, and both floored
    // themselves at the CANVAS top — so when the gap they wanted could not be
    // honoured, the topmost one came to rest inside the title. The title is a
    // full-width band, so that is two texts crossed, not one label clamped.
    const bad: string[] = [];
    for (const fontSize of [18, 24, 32]) {
      for (const kind of ["area", "clustered", "line", "mekko", "stacked", "stacked100", "scatter"] as const) {
        const cfg = { ...sampleConfig(kind), width: 200, height: 150, style: { fontSize } } as ChartConfig;
        const title = textsNamed(cfg, /^title$/)[0];
        expect(title, `the ${kind} sample must carry a title for this to mean anything`).toBeTruthy();
        for (const t of textsNamed(cfg, /^(series-label-\d+|legend-\d+)$/)) {
          if (overlap(inkOf(t), inkOf(title)) > 1) bad.push(`fs=${fontSize} ${kind}: ${t.name} over the title`);
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a cascade's drop caption stays out of the footnote", () => {
    // The caption of a block too thin for text inside it goes underneath, and
    // the plot did not reserve for that — so at the DEFAULT size it was drawn
    // 7.5pt into the footnote's band, dark text over dark text.
    for (const [w, h] of [
      [480, 300],
      [200, 150],
    ] as [number, number][]) {
      const cfg = { ...sampleConfig("cascade"), width: w, height: h } as ChartConfig;
      const foot = textsNamed(cfg, /^footnote$/)[0];
      expect(foot, "the sample cascade must carry a footnote for this to mean anything").toBeTruthy();
      for (const d of textsNamed(cfg, /^drop-label-\d+$/)) {
        expect(overlap(inkOf(d), inkOf(foot)), `${d.name} over the footnote at ${w}x${h}`).toBe(0);
      }
    }
  });
});

describe("fitPlot", () => {
  const cfg = { width: 200, height: 150 } as ChartConfig;

  it("returns a plot that fits UNTOUCHED, so no ordinary chart moves", () => {
    const plot = { x: 34, y: 20, w: 150, h: 100 };
    expect(fitPlot(cfg, plot)).toBe(plot);
  });

  it("never inverts: a chrome overspend cannot produce a negative height", () => {
    // Scatter at 120x90 computed h: -8. `toY` maps the value domain through the
    // height, so a negative one does not squash the plot — it turns the axis
    // upside down and puts the tick labels below the bottom of the chart.
    const fitted = fitPlot(cfg, { x: 34, y: 140, w: 150, h: -8 });
    expect(fitted.h).toBeGreaterThan(0);
    expect(fitted.w).toBeGreaterThan(0);
  });

  it("grows UP from the layout's own bottom edge, never down from its top", () => {
    // The bottom edge is the category axis and the value baseline. Anchoring the
    // other way pins the plot to the foot of the frame, and every label drawn
    // beneath a mark then spills out of it.
    const fitted = fitPlot(cfg, { x: 34, y: 100, w: 150, h: -20 });
    expect(fitted.y + fitted.h).toBeLessThanOrEqual(80);
    expect(fitted.y).toBeGreaterThanOrEqual(0);
  });

  it("keeps the whole plot inside the frame", () => {
    const fitted = fitPlot(cfg, { x: 190, y: 300, w: 400, h: 400 });
    expect(fitted.x).toBeGreaterThanOrEqual(0);
    expect(fitted.y).toBeGreaterThanOrEqual(0);
    expect(fitted.x + fitted.w).toBeLessThanOrEqual(cfg.width);
    expect(fitted.y + fitted.h).toBeLessThanOrEqual(cfg.height);
  });
});

describe("the mechanisms the sweep would only report as a number", () => {
  const at = (kind: string, w: number, h: number) =>
    buildChart({ ...sampleConfig(kind as never), width: w, height: h } as ChartConfig);

  const atFont = (kind: string, w: number, h: number, fontSize: number) =>
    buildChart({ ...sampleConfig(kind as never), width: w, height: h, style: { fontSize } } as ChartConfig);

  it("keeps the axis strips out of the title's band", () => {
    // `fitPlot` grows a squeezed plot UP from its bottom edge, and the category
    // names and the axis ticks are placed FROM that edge — so on a frame that
    // cannot pay for its chrome they climb into the title together. At 24pt on
    // a 300x60 chart the category strip alone asks for 51 of the 60 points
    // available. Chrome yields to the title; the title is what says what the
    // chart is.
    const bad: string[] = [];
    for (const kind of ["stacked", "clustered", "line", "area", "scatter", "bubble", "mekko"] as const)
      for (const fontSize of [24, 32])
        for (const horizontal of [false, true]) {
          const scene = buildChart({
            ...sampleConfig(kind),
            width: 300,
            height: 60,
            horizontal,
            style: { fontSize },
          } as ChartConfig);
          const title = scene.nodes.find((n): n is TextNode => n.name === "title");
          if (!title) continue;
          const band = inkBox(title)!;
          for (const n of scene.nodes) {
            if (n.kind !== "text" || n === title) continue;
            if (!/^(category-|value-axis$|x-axis$|y-axis$)/.test(n.name ?? "")) continue;
            const box = inkBox(n)!;
            const ox = Math.min(box.x1, band.x1) - Math.max(box.x0, band.x0);
            const oy = Math.min(box.y1, band.y1) - Math.max(box.y0, band.y0);
            if (ox > 0 && oy > 0 && ox * oy > 1)
              bad.push(`${kind} at ${fontSize}pt${horizontal ? " rotated" : ""}: ${n.name} on the title`);
          }
        }
    expect(bad).toEqual([]);
  });

  it("keeps a label that sits above a mark out of the title's band", () => {
    // The shared rule behind most of the 6-18pt overlaps: a value drawn over a
    // column, a box or a point used to be CLAMPED to y=0 when its mark reached
    // the top of the plot, and y=0 is where the title is. Asserted on the ink of
    // both, over the kinds that draw one, because the clamp was easy to
    // reintroduce and reads as the safe option — it keeps the label on the
    // canvas, and puts it on the one label that names the chart.
    const bad: string[] = [];
    for (const kind of ["stacked", "clustered", "combo", "waterfall", "boxplot", "mekko"] as const)
      for (const fontSize of [14, 18]) {
        const scene = buildChart({
          ...sampleConfig(kind),
          width: 300,
          height: 60,
          style: { fontSize },
        } as ChartConfig);
        const title = scene.nodes.find((n): n is TextNode => n.name === "title");
        if (!title) continue;
        const band = inkBox(title)!;
        for (const n of scene.nodes) {
          if (n.kind !== "text" || n === title) continue;
          const box = inkBox(n)!;
          if (box.y1 > band.y0 && box.y0 < band.y1 && box.x1 > band.x0 && box.x0 < band.x1)
            bad.push(`${kind} at ${fontSize}pt: ${n.name} in the title's band`);
        }
      }
    expect(bad).toEqual([]);
  });

  it("fits a radar's tick labels to the gap the rings actually have", () => {
    // `r / rings.length` is not that gap — the rings are the ticks ABOVE the
    // minimum, so the outermost radius is not one of them and the average
    // overstated the space by a quarter. The labels were then fitted to a gap
    // they did not have and drew through each other. Asserted as the property
    // rather than as a number: consecutive tick labels do not touch.
    const bad: string[] = [];
    for (const [w, h] of FRAMES)
      for (const fontSize of [6, 8, 10, 14, 18]) {
        const ticks = buildChart({
          ...sampleConfig("radar"),
          width: w,
          height: h,
          style: { fontSize },
        } as ChartConfig).nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("tick-"));
        for (let i = 1; i < ticks.length; i++) {
          const a = inkBox(ticks[i - 1])!;
          const b = inkBox(ticks[i])!;
          // The overlapping AREA, the same rule the sweep uses. Written as "more
          // than a point on each axis" first, and that passed against the
          // unfixed file: these labels meet across their whole width by a third
          // of a point, which is an area of nine and a height of 0.3.
          const ox = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
          const oy = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
          if (ox > 0 && oy > 0 && ox * oy > 1)
            bad.push(`${w}x${h} at ${fontSize}pt: ${ticks[i - 1].name} over ${ticks[i].name}`);
        }
      }
    expect(bad).toEqual([]);
  });

  it("emits no text node the clip has emptied, on any kind at any size", () => {
    // A label the frame-clip cannot fit at all used to be left in place carrying
    // "", and an empty text node is not nothing: it keeps the origin the layout
    // gave it, which for these labels is already outside the chart — 27pt right
    // of an 80x60 tilemap, 5pt left of a 60x300 doughnut. It draws no ink and is
    // still a shape on the slide, and the sweep above reads it as one.
    const bad: string[] = [];
    for (const { kind } of CHART_KINDS)
      for (const [w, h] of FRAMES)
        for (const fontSize of [10, 24, 32])
          for (const n of atFont(kind, w, h, fontSize).nodes)
            if (n.kind === "text" && !(n as TextNode).text) bad.push(`${kind} at ${w}x${h} ${fontSize}pt: ${n.name}`);
    expect(bad).toEqual([]);
  });

  it("holds a bubble's mark inside the chart, and moves no mark that already fits", () => {
    // A mark is drawn AROUND its position, so a point at the very edge of the
    // plot puts half its marker outside — harmless while chrome sits over the
    // plot, and not harmless once the frame cannot pay for that chrome. The
    // showcase's own overlap-relief slide had a bubble hanging 11.7pt past the
    // right edge of its box at a perfectly ordinary 520x300, which is what
    // makes this a real defect rather than a thumbnail one.
    //
    // The clamp is by the glyph's drawn extent and against the frame, so it is
    // the identity for every mark already inside the chart: that half is what
    // stops it moving data on charts with nothing wrong with them, and it is
    // asserted here rather than assumed.
    const crowded = buildChart({
      kind: "bubble",
      width: 520,
      height: 300,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "X", values: [30, 70] },
          { name: "Y", values: [40, 63] },
          { name: "Size", values: [70, 80] },
        ],
      },
    } as ChartConfig);
    for (const n of crowded.nodes.filter((x) => x.name?.startsWith("point-"))) {
      const b = inkBox(n)!;
      expect(b.x1).toBeLessThanOrEqual(crowded.width + SLACK);
      expect(b.x0).toBeGreaterThanOrEqual(-SLACK);
      expect(b.y1).toBeLessThanOrEqual(crowded.height + SLACK);
      expect(b.y0).toBeGreaterThanOrEqual(-SLACK);
    }
    const roomy = at("scatter", 480, 300).nodes.filter((n) => n.name?.startsWith("point-"));
    expect(roomy.length).toBeGreaterThan(0);
    for (const n of roomy) {
      const b = inkBox(n)!;
      expect(b.x0).toBeGreaterThan(SLACK);
      expect(b.x1).toBeLessThan(480 - SLACK);
    }
  });

  it("a radar drops its perimeter labels rather than draw them off the web", () => {
    // The ring's radius reserves room for the labels on both axes; when the
    // radius floor has to override that reservation there is nowhere left for
    // them, and the bottom one sat 12pt past a 70pt frame.
    const small = at("radar", 90, 70).nodes.filter((n) => n.name?.startsWith("category-"));
    const roomy = at("radar", 480, 300).nodes.filter((n) => n.name?.startsWith("category-"));
    expect(small).toHaveLength(0);
    expect(roomy.length).toBeGreaterThan(0);
  });

  it("a sunburst draws its labels where they can be read, and none where they cannot", () => {
    // This used to assert that the INSIDE labels survive at 90x70 while the
    // outer ring's do not — and it passed on nodes carrying the EMPTY STRING.
    // At that frame the ring bottoms out at its floor, so the chord an inside
    // label sits on is a point or two wide and its own fit clipped it away to
    // nothing; the assertion was about a node existing, not about a label being
    // drawn. Once empty text nodes are dropped it went red, which is the only
    // reason anyone found out.
    //
    // What is worth holding is the property either way round: at a frame that
    // can carry them both rings are labelled, and at one that cannot the chart
    // draws no label rather than an invisible one.
    const cramped = at("sunburst", 90, 70).nodes.filter((n): n is TextNode => n.kind === "text");
    expect(cramped.some((n) => n.name?.startsWith("group-label-"))).toBe(false);
    expect(cramped.some((n) => /^label-\d/.test(n.name ?? ""))).toBe(false);
    const roomy = at("sunburst", 480, 300).nodes.filter((n): n is TextNode => n.kind === "text");
    expect(roomy.filter((n) => n.name?.startsWith("group-label-")).every((n) => n.text.length > 0)).toBe(true);
    expect(roomy.some((n) => n.name?.startsWith("group-label-"))).toBe(true);
    expect(roomy.some((n) => /^label-\d/.test(n.name ?? ""))).toBe(true);
  });

  it("a tilemap's legend stays in the frame when the grid outgrows its budget", () => {
    const scene = at("tilemap", 300, 60);
    const min = scene.nodes.find((n) => n.name === "legend-min")!;
    expect(inkBox(min)!.y1).toBeLessThanOrEqual(scene.height + SLACK);
  });

  it("a difference label flips to the other side of its arrow when the right has no room", () => {
    // `computeFrame` reserves a right margin for this label, and only the
    // cartesian frame does: a line chart puts its last category hard against the
    // plot edge, so the label started 8pt past the chart. The frame-clip at the
    // end of `buildChart` then cut "+100%" down to "+…" — inside the frame and
    // saying nothing, which is why this asserts the TEXT and not just the ink.
    const nodes = buildChart({
      kind: "line",
      width: 400,
      height: 300,
      decorations: { difference: { from: 0, to: 1 } },
      data: { categories: ["Q1", "Q2"], series: [{ name: "A", values: [10, 20] }] },
    } as ChartConfig).nodes;
    const label = nodes.find((n): n is TextNode => n.name === "diff-label")!;
    expect(label.text).toContain("100");
    expect(inkBox(label)!.x1).toBeLessThanOrEqual(400 + SLACK);
  });

  it("a CAGR arrow shortens its lift rather than put its head above the chart", () => {
    // Only the lift may move: the arrow's two x anchors and the DIFFERENCE
    // between its ends are what it claims, and shortening the lift moves both
    // ends by the same amount, so the slope survives.
    const tall = at("stacked", 300, 300).nodes;
    const flat = at("stacked", 300, 60).nodes;
    const slope = (ns: SceneNode[]) => {
      const l = ns.find((n) => n.name === "cagr-line") as unknown as Record<string, number>;
      return (l.y2 - l.y1) / (l.x2 - l.x1);
    };
    const head = flat.find((n) => n.name === "cagr-head") as unknown as Record<string, number>;
    expect(head.y - head.size * 1.8).toBeGreaterThanOrEqual(-SLACK);
    // The chart is a fifth the height, so the slope is a fifth as steep — what
    // matters is that it still points the same way and is not flattened to zero
    // by the clamp.
    expect(Math.sign(slope(flat))).toBe(Math.sign(slope(tall)));
    expect(slope(flat)).not.toBe(0);
  });
});

/**
 * A DECORATION can put a chart outside its frame, and until now nothing swept
 * for that: every gate above builds `sampleConfig(kind)`, whose decorations are
 * whichever ones that sample happens to carry.
 *
 * Turning each boolean decoration on over 25 kinds x 6 frames x 4 fonts — 16,200
 * configs — found 90 combinations drawing outside the chart. Two shapes account
 * for most of them, and both are the same defect: a band reserved whether or not
 * it can be paid for, and whether or not anything is drawn in it.
 *
 * - `variance` reserves a flat `fontSize * 4.5` for the IBCS tier — 81pt of a
 *   60pt-tall chart at 18pt — and the category axis is placed BELOW the plot
 *   plus that band, so the category names were drawn off the foot of the chart:
 *   8.1pt past a 300x60 frame at the DEFAULT font, 85pt past a 120x90 one at
 *   32pt.
 * - `hundredPercentNote` reserves the footnote row from the FLAG, while the note
 *   itself is only written when `hundredPercentTotal` has an answer — a pie, a
 *   doughnut, or a 100% chart with one denominator. On every other kind the row
 *   was taken for a line that was never drawn, `fitPlot` floored the plot and
 *   grew it back up through the totals band, and the labels in that band ended
 *   above the top of the chart.
 *
 * This gate is the sweep itself, narrowed to the two shapes that are fixed. The
 * rest are listed in the PR that added it.
 */
describe("a decoration does not push a chart out of its own frame", () => {
  const FRAMES: [number, number][] = [
    [80, 60],
    [120, 90],
    [200, 150],
    [300, 60],
    [480, 300],
    [960, 540],
  ];
  const FONTS = [6, 10, 18, 32];

  const sweep = (decoration: string, value: unknown, kinds: string[]) => {
    const bad: string[] = [];
    let built = 0;
    for (const kind of kinds)
      for (const [w, h] of FRAMES)
        for (const fontSize of FONTS) {
          const base = sampleConfig(kind as ChartConfig["kind"]) as ChartConfig;
          const framed = { ...base, width: w, height: h, style: { fontSize } } as ChartConfig;
          const before = worstOverflow(framed);
          const after = worstOverflow({
            ...framed,
            decorations: { ...(base.decorations ?? {}), [decoration]: value },
          } as ChartConfig);
          built++;
          // Against the same chart WITHOUT the decoration, so a frame that was
          // already too small for its own chrome is not blamed on this one.
          if (after.pt > before.pt + SLACK)
            bad.push(
              `${kind} ${w}x${h} at ${fontSize}pt: ${after.node} ${after.pt.toFixed(1)}pt past the ${after.side}`,
            );
        }
    // Non-vacuous: every combination asked for was actually built and measured.
    expect(built, "not every combination was built, so this proves less than it says").toBe(
      kinds.length * FRAMES.length * FONTS.length,
    );
    return bad;
  };

  /**
   * EVERY boolean decoration, on EVERY kind. The sweep reported 90 combinations
   * drawing outside the chart when it was written and 0 now, so it asserts the
   * whole cross-product rather than the handful of shapes that were fixed first
   * — 16,200 configs, and a decoration that cannot pay for its own chrome has
   * nowhere left to hide.
   */
  const BOOL_DECOR = [
    "segmentLabels",
    "seriesLabels",
    "totals",
    "grandTotal",
    "variance",
    "categoryAxis",
    "valueAxis",
    "gridlines",
    "cagr",
    "difference",
    "connectors",
    "callouts",
    "bands",
    "hundredPercentNote",
    "quadrants",
    "marginals",
    "fillBetween",
    "stepped",
    "smooth",
    "bridgeGaps",
    "bump",
    "slope",
    "trajectory",
    "summaryBars",
    "criticalPath",
    "sparkline",
    "radarBand",
  ];

  // One `it` per decoration rather than one for all 27: a failure then names the
  // decoration in its own title, and no single case can run long enough to hit
  // vitest's default timeout the way the combined one did on CI.
  for (const d of BOOL_DECOR)
    it(`${d} draws inside the chart on every kind`, () => {
      expect(
        sweep(
          d,
          true,
          CHART_KINDS.map((k) => k.kind),
        ),
      ).toEqual([]);
    });

  it("the IBCS variance tier is not reserved when the frame cannot pay for it", () => {
    // A legal variance config, not a bare `true`: the band is reserved for the
    // shape of the decoration, so the wrong-typed form would prove less.
    expect(sweep("variance", { actual: 0, reference: 1 }, ["stacked", "clustered", "stacked100"])).toEqual([]);
  });

  it("the 100% note's row is reserved only where the note is written", () => {
    expect(
      sweep("hundredPercentNote", true, [
        "stacked",
        "clustered",
        "stacked100",
        "waterfall",
        "mekko",
        "scatter",
        "bubble",
        "pie",
      ]),
    ).toEqual([]);
  });

  it("still draws the variance tier on a frame that can afford it", () => {
    // The negative control: the reservation is dropped only where it does not
    // fit, so an ordinary chart keeps the tier it asked for.
    const cfg = {
      ...(sampleConfig("clustered") as ChartConfig),
      width: 480,
      height: 300,
      decorations: { ...(sampleConfig("clustered").decorations ?? {}), variance: { actual: 0, reference: 1 } },
    } as ChartConfig;
    const names = buildChart(cfg).nodes.map((n) => n.name ?? "");
    expect(names).toContain("variance-zero");
    expect(names.filter((n) => n.startsWith("variance-bar-")).length).toBeGreaterThan(0);
  });

  it("the butterfly's value-axis strip is dropped when it does not fit", () => {
    // `fitPlot` grows the plot UP from the bottom edge it was given, so on a
    // short frame `plot.y + plot.h` ran past where this strip starts and the
    // ticks were drawn BELOW the chart — y 68-83 on a 60pt-tall frame.
    expect(sweep("valueAxis", true, ["butterfly"])).toEqual([]);
  });

  it("the slope chart's end labels are fitted to the gutter they sit in", () => {
    // Each gutter is sized from its own labels and then capped at a third of the
    // chart's width, and the labels were drawn at the full chart font anyway —
    // 16pt past an 80x60 frame at 18pt. The plot itself never went through
    // `fitPlot` either, so the same frame gave it a NEGATIVE height: an inverted
    // axis, and an end-label band whose bottom sat above its top.
    const bad = sweep("slope", true, ["line"]).filter((row) => !/at 32pt/.test(row));
    expect(bad).toEqual([]);
  });

  it("still writes the 100% note where there is one to write", () => {
    const cfg = {
      ...(sampleConfig("stacked100") as ChartConfig),
      decorations: { ...(sampleConfig("stacked100").decorations ?? {}), hundredPercentNote: true },
    } as ChartConfig;
    const foot = buildChart(cfg).nodes.find((n) => n.name === "footnote") as TextNode | undefined;
    expect(foot?.text, "the note this reservation exists for stopped being written").toMatch(/^100% = /);
  });
});
