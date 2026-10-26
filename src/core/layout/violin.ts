import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { type SceneNode } from "../scene";
import { resolveFormat } from "../format";
import { maxOf, minOf } from "../agg";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import { thinOutline } from "../geometry";
import { chromeNodes, computeFrame, valueScale, type Frame } from "./frame";
import type { LayoutResult } from "./column";

const quantile = (sorted: number[], p: number): number => {
  /* v8 ignore next -- defensive: every caller guards samples.length >= 2 first */
  if (!sorted.length) return 0;
  const idx = p * (sorted.length - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
};

/**
 * Violin chart: a mirrored kernel-density shape per category, from raw samples
 * (every datasheet row is an observation, like the boxplot's raw mode). Shows
 * the full distribution, not just the five-number summary. The body is a
 * filled polygon — solid in SVG/pptx, outline-only in the live add-in (no
 * freeform fills there).
 */
export function layoutViolin(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  // Each category's observations get read repeatedly — the KDE loop, the flattened
  // allSamples, and both anchor builders below all ask for the same column. Memoize
  // per category so the series scan (map + filter over every row) runs once each.
  const sampleCache = new Map<number, number[]>();
  const samplesOf = (c: number): number[] => {
    let s = sampleCache.get(c);
    if (!s) {
      s = data.series.map((ser) => ser.values[c]).filter((v): v is number => v != null);
      sampleCache.set(c, s);
    }
    return s;
  };
  const allSamples = data.categories.flatMap((_, c) => samplesOf(c));
  const fmt = resolveFormat(allSamples, cfg.numberFormat);
  void fmt;

  const dataMin = minOf(allSamples.length ? allSamples : [0]);
  const dataMax = maxOf(allSamples.length ? allSamples : [1]);
  // A function of the frame, because the value-axis column is sized from the ticks
  // it prints (see `computeFrame`). No forced zero: a distribution's domain is its
  // data range.
  const scaleOf = (f: Frame) => valueScale(f, dataMin, dataMax, cfg.scale, undefined, undefined, false);
  const { frame } = computeFrame(cfg, style, { ...decor, seriesLabels: false }, [], scaleOf);
  const scale = scaleOf(frame);
  const slotLen = frame.w / Math.max(1, n);
  const halfW = slotLen * 0.42;
  const centers = data.categories.map((_, c) => frame.x + slotLen * (c + 0.5));

  const nodes: SceneNode[] = chromeNodes(cfg, style, { ...decor, seriesLabels: false }, frame, centers, scale);
  const vmin = scale.min;
  const vmax = scale.max;
  const M = 40;

  data.categories.forEach((_, c) => {
    const samples = samplesOf(c);
    if (samples.length < 2) return;
    const sorted = samples.slice().sort((a, b) => a - b);
    const mean = samples.reduce((a, b) => a + b, 0) / samples.length;
    const sd = Math.sqrt(samples.reduce((a, b) => a + (b - mean) ** 2, 0) / samples.length);
    // Silverman's rule-of-thumb bandwidth. A category whose observations are all
    // equal has sd 0, and the axis floor below is then the whole bandwidth — a
    // spike a sixtieth of the axis tall, whatever the unit. There was a fallback
    // here, `|| 1`, in DATA UNITS: it set the bandwidth to ~0.8 for any constant
    // category, which covers the whole axis once that spans a few units or less, so
    // a category with zero spread was drawn as a rectangle across the full plot —
    // the widest-looking distribution on the chart (y 8..282 on a 0.40-0.60 axis,
    // against 140..150 for the same data x100).
    const h = Math.max((vmax - vmin) / 60, 1.06 * sd * Math.pow(samples.length, -0.2));
    const lo = sorted[0];
    const hi = sorted[sorted.length - 1];

    // Evaluate the KDE across THIS category's own support (clamped to the axis).
    // Sampling the chart-wide range and skipping points outside the category
    // silently dropped any category whose spread was small relative to the whole
    // range — fewer than 2 of the M+1 grid points landed in it, so it rendered
    // no violin and no median tick at all despite having valid observations.
    const gLo = Math.max(vmin, lo - h);
    const gHi = Math.min(vmax, hi + h);
    /* v8 ignore next -- unreachable: valueScale guarantees vmax > vmin and h > 0, so the window never collapses */
    if (!(gHi > gLo)) return;
    const levels: { y: number; d: number }[] = [];
    let maxD = 0;
    for (let k = 0; k <= M; k++) {
      const y = gLo + ((gHi - gLo) * k) / M;
      let d = 0;
      for (const s of samples) d += Math.exp(-0.5 * ((y - s) / h) ** 2);
      d /= samples.length * h * Math.sqrt(2 * Math.PI);
      levels.push({ y, d });
      if (d > maxD) maxD = d;
    }
    /* v8 ignore next -- unreachable: the loop always fills M+1 levels and a Gaussian KDE sum is > 0 */
    if (levels.length < 2 || maxD <= 0) return;

    const cxc = centers[c];
    const points: { x: number; y: number }[] = [];
    for (const lv of levels) points.push({ x: cxc + (lv.d / maxD) * halfW, y: scale.toY(lv.y) });
    for (let k = levels.length - 1; k >= 0; k--)
      points.push({ x: cxc - (levels[k].d / maxD) * halfW, y: scale.toY(levels[k].y) });

    const color = seriesColor(style, c, data.series.find((s) => s.color)?.color);
    nodes.push({
      kind: "polygon",
      /**
       * THINNED, because on the add-in every edge of this outline is a SHAPE.
       *
       * `M` is a fixed 41-level KDE grid, mirrored to 82 points, and Office.js
       * has no freeform fill — it draws a polygon as one line per edge. Three
       * bodies came to 246 of this chart's 259 office shapes, which made the
       * violin the heaviest thing in the shipped deck, and most of those points
       * sit in the TAILS where consecutive samples land on top of one another
       * and each still costs a shape.
       *
       * A quarter of a point is sub-pixel at any size this is read at, and it
       * takes 246 points to 46. The grid stays at 41: that is what makes the
       * curve smooth where it BENDS. This drops only the samples that were never
       * distinguishable from the line through their own neighbours.
       */
      points: thinOutline(points, 0.25),
      fill: lerpColor(style.background, color, 0.32),
      stroke: color,
      strokeWidth: 1,
      name: `violin-${c}`,
    });
    // Median tick.
    const med = quantile(sorted, 0.5);
    nodes.push({
      kind: "line",
      x1: cxc - halfW * 0.35,
      y1: scale.toY(med),
      x2: cxc + halfW * 0.35,
      y2: scale.toY(med),
      stroke: style.text,
      strokeWidth: 1.75,
      name: `median-${c}`,
    });
  });

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => halfW * 2),
      // A category with no observations must NOT stand in value 0: this scale
      // is data-driven (no forced zero), so toY(0) can be hundreds of points
      // below the frame and would take any callout on that category with it.
      // Anchor an empty category on the plot floor instead.
      columnTop: data.categories.map((_, c) => {
        const s = samplesOf(c);
        return s.length ? scale.toY(Math.max(...s)) : frame.y + frame.h;
      }),
      columnValue: data.categories.map((_, c) => {
        const s = samplesOf(c);
        return s.length
          ? quantile(
              s.slice().sort((a, b) => a - b),
              0.5,
            )
          : 0;
      }),
      // The 0 above is the empty category's stand-in, not a median: averaged in,
      // it dragged the mean line down and an arrow to the category printed "-100%".
      columnHasData: data.categories.map((_, c) => samplesOf(c).length > 0),
      baselineY: frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: scale.toY,
    },
  };
}
