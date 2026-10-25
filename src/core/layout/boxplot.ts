import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode, type TextNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { maxOf, minOf } from "../agg";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import {
  titleInkBottom,
  aboveMarkFontSize,
  bandFontSize,
  baselineNode,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  legendRowCount,
  titleHeight,
  valueScale,
} from "./frame";
import { horizontalChrome, legendRow, type LayoutResult } from "./column";

/** The five-number-summary datasheet rows (think-cell's own boxplot recipe). */
const SUMMARY_ROWS: [keyof FiveNum, RegExp][] = [
  ["min", /^min$/i],
  ["q1", /^q1$/i],
  ["median", /^median$/i],
  ["q3", /^q3$/i],
  ["max", /^max$/i],
];
const MEAN_ROW = /^mean$/i;
const OUTLIER_ROW = /^outliers?\b/i;

/**
 * The ink of the mean marker `×`, as a multiple of the size it is drawn at.
 *
 * Read off the node `layoutBoxplot` draws it in — a `valign: middle` glyph in a
 * box `fs * 1.4` tall centred on the mean — by the rule the frame gate measures
 * ink with: the ink runs from 0.44 above the mean's position to 0.57 below it
 * (the 0.57 the median number's own fit below already uses), and is as wide as
 * the glyph is.
 */
const MEAN_INK_W = textWidth("×", 1, true);
const MEAN_INK_UP = 0.44;
const MEAN_INK_DOWN = 0.57;

/**
 * The mean marker for the mean at `at`, or null when the plot has no room for one
 * that can be read.
 *
 * A mark is drawn AROUND its position, and this one sits at the one place on the
 * chart a layout cannot move it from: the mean's value. So it gives up size
 * instead, the answer `markInFrame` gives a line's markers, and is dropped below
 * `MIN_LABEL_FS` — the box and its whiskers still say where the data is. It was
 * drawn at the chart font whatever the plot could hold, and over raw
 * observations that put it through four neighbours at once (53 overlapping
 * pairs in the overlap sweep): the next category's `×`, the title, the category
 * names under the plot, and the value-axis ticks under a sideways one.
 *
 * Four bounds, each by `bandFontSize` so a marker that already fits keeps the
 * chart font to the last bit, and its box scales by a RATIO that is exactly 1
 * then (the median number's rule: a ratio keeps every ordinary chart's geometry
 * identical, where re-deriving the box from the new size moves it by a rounding):
 *   - the PLOT, every side of it: the glyph is centred on the mean, so its ink must
 *     stay inside the rectangle the data is drawn in, below the title's ink — the
 *     band a squeezed plot is grown up into.
 *   - the NEIGHBOUR along the category axis: `pitch` is the distance to the next
 *     mean marker, the slot (or a group's sub-slot), and what the glyph needs of
 *     it is its width upright and its height sideways. Two markers abut at worst.
 */
function meanMark(
  cfg: ChartConfig,
  style: ChartStyle,
  frame: { x: number; y: number; w: number; h: number },
  at: { x: number; y: number },
  pitch: number,
  sideways: boolean,
  name: string,
): TextNode | null {
  const fs = style.fontSize;
  const top = Math.max(frame.y, titleInkBottom(cfg, style));
  const fits = [
    bandFontSize(fs, 2 * Math.min(at.x - frame.x, frame.x + frame.w - at.x), MEAN_INK_W),
    bandFontSize(fs, at.y - top, MEAN_INK_UP),
    bandFontSize(fs, frame.y + frame.h - at.y, MEAN_INK_DOWN),
    bandFontSize(fs, pitch, sideways ? MEAN_INK_UP + MEAN_INK_DOWN : MEAN_INK_W),
  ];
  if (!fits.every((f) => f > 0)) return null;
  const size = Math.min(...fits);
  const k = size / fs;
  return {
    kind: "text",
    x: at.x - fs * 0.6 * k,
    y: at.y - fs * 0.7 * k,
    w: fs * 1.2 * k,
    h: fs * 1.4 * k,
    text: "×",
    fontSize: size,
    bold: true,
    color: style.neutral,
    align: "center",
    valign: "middle",
    name,
  };
}

interface FiveNum {
  min: number;
  q1: number;
  median: number;
  q3: number;
  max: number;
}

interface Box extends FiveNum {
  mean?: number;
  outliers: number[];
  /** Raw observations (raw-sample mode only), for the optional jitter overlay. */
  samples?: number[];
}

/** Excel-style exclusive/inclusive quartile on a sorted sample. */
function quartile(sorted: number[], p: number, method: "exclusive" | "inclusive"): number {
  const n = sorted.length;
  const rank = method === "exclusive" ? p * (n + 1) - 1 : p * (n - 1);
  const lo = Math.max(0, Math.min(n - 1, Math.floor(rank)));
  const hi = Math.max(0, Math.min(n - 1, Math.ceil(rank)));
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (rank - lo);
}

/**
 * Box-and-whisker chart, one box per category. Two datasheet modes:
 *  - Precomputed: rows named Min/Q1/Median/Q3/Max (+ optional Mean,
 *    Outlier …) carry the summary directly — whiskers run to Min/Max. A group
 *    is a summary only when it has all five of those rows.
 *  - Raw samples: every row is an observation; quartiles are computed
 *    (exclusive method) and whiskers use Tukey 1.5×IQR fences with
 *    outliers drawn as dots.
 * Grouped boxplots: suffix rows with "| <group>" ("Min | 2024",
 * "Q1 | 2024", … "Min | 2025", …) — each group gets a side-by-side box
 * per category, colored from the palette, with a legend.
 */
/**
 * Every box the chart draws, in group-major order, plus the flat list of values
 * they span. The layout scales to these, and boxplotExtent reports them, so the
 * two can't drift apart — the mean±SD variant's whiskers reach beyond the raw
 * samples, and an extent taken from the samples alone understates it.
 */
function boxplotBoxes(cfg: ChartConfig): {
  groupNames: string[];
  grouped: (Box | null)[][];
  all: number[];
  drawn: number[];
} {
  const { data } = cfg;
  const opts = cfg.boxplot ?? {};
  // Grouped boxplots: "Min | 2024"-style suffixes split rows into groups.
  const baseName = (name: string) => name.split("|")[0].trim();
  const groupOf = (name: string) => (name.includes("|") ? name.split("|").slice(1).join("|").trim() : "");
  const groupNames: string[] = [];
  for (const srs of data.series) {
    const g = groupOf(srs.name);
    if (!groupNames.includes(g)) groupNames.push(g);
  }
  const rowsOf = (g: string) => data.series.filter((srs) => groupOf(srs.name) === g);
  const find = (g: string, re: RegExp) => rowsOf(g).find((srs) => re.test(baseName(srs.name)));
  // A group is a five-number SUMMARY only when it carries all five rows. It was
  // "any one row is named like one, and then the whole CHART is", and in raw-
  // sample mode a row's name is whatever the observation is called: Anna, Ben,
  // Max, Zoe, Ola switched the chart to summary mode, `boxFor` then wanted
  // Min/Q1/Median/Q3/Max, found only the one, and drew no box in any category —
  // axes and nothing else, with no message. Quarterly samples named Q1..Q4 hit
  // the same switch with two of the five names. Asking for the full set is the
  // one question a real summary always answers yes to and a stray name never
  // does. The price is a half-typed summary (Median misspelt): its rows are now
  // read as observations instead of drawing nothing.
  //
  // PER GROUP, because the rows of "| 2024" and "| 2025" are separate sheets that
  // share a chart: one summary group used to flip every raw-sample group beside
  // it to precomputed mode, where none of its rows could satisfy `find`.
  const summaryGroups = new Set(groupNames.filter((g) => SUMMARY_ROWS.every(([, re]) => find(g, re))));

  const boxFor = (g: string, c: number): Box | null => {
    if (summaryGroups.has(g)) {
      const summary = {} as FiveNum;
      for (const [key, re] of SUMMARY_ROWS) {
        const v = find(g, re)?.values[c];
        if (v == null) return null;
        summary[key] = v;
      }
      const mean = find(g, MEAN_ROW)?.values[c] ?? undefined;
      const outliers = rowsOf(g)
        .filter((srs) => OUTLIER_ROW.test(baseName(srs.name)))
        .map((srs) => srs.values[c])
        .filter((v): v is number => v != null);
      return { ...summary, mean: mean ?? undefined, outliers };
    }
    // Raw samples: this category's column across the group's rows.
    const sample = rowsOf(g)
      .map((srs) => srs.values[c])
      .filter((v): v is number => v != null)
      .sort((a, b) => a - b);
    if (sample.length < 2) return null;
    // Mean±SD variant: box = mean ± 1·SD, centre = mean, whiskers = mean ± 2·SD.
    if (opts.meanSd) {
      const mean = sample.reduce((a, b) => a + b, 0) / sample.length;
      const variance = sample.reduce((a, b) => a + (b - mean) ** 2, 0) / (sample.length - 1);
      const sd = Math.sqrt(variance);
      return {
        min: mean - 2 * sd,
        q1: mean - sd,
        median: mean,
        q3: mean + sd,
        max: mean + 2 * sd,
        mean: undefined,
        outliers: [],
        samples: sample,
      };
    }
    const method = opts.quartileMethod ?? "exclusive";
    const q1 = quartile(sample, 0.25, method);
    const q3 = quartile(sample, 0.75, method);
    const median = quartile(sample, 0.5, "inclusive");
    const whiskers = opts.whiskers ?? "tukey";
    let lo = sample[0];
    let hi = sample[sample.length - 1];
    let outliers: number[] = [];
    if (whiskers === "tukey") {
      const k = opts.iqrMultiplier ?? 1.5;
      const loFence = q1 - k * (q3 - q1);
      const hiFence = q3 + k * (q3 - q1);
      const inside = sample.filter((v) => v >= loFence && v <= hiFence);
      lo = inside[0] ?? q1;
      hi = inside[inside.length - 1] ?? q3;
      outliers = sample.filter((v) => v < loFence || v > hiFence);
    }
    const mean = opts.showMean ? sample.reduce((a, b) => a + b, 0) / sample.length : undefined;
    return { min: lo, q1, median, q3, max: hi, mean, outliers, samples: sample };
  };

  // boxes[g][c]
  const grouped = groupNames.map((g) => data.categories.map((_, c) => boxFor(g, c)));
  const all = grouped
    .flat()
    .flatMap((b) => (b ? [b.min, b.max, ...b.outliers, ...(b.mean != null ? [b.mean] : [])] : []));
  // The jitter overlay plots every raw observation. Normally the box already
  // spans them, but the mean±SD variant's whiskers stop at mean ± 2·SD and it
  // reports no outliers, so a sample past 2·SD has nothing covering it. The
  // scale needs those dots; the number format does not (it describes the
  // summary, and the samples it is derived from can only add noise).
  const drawn = opts.jitter ? [...all, ...grouped.flat().flatMap((b) => b?.samples ?? [])] : all;
  return { groupNames, grouped, all, drawn };
}

export function layoutBoxplot(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const opts = cfg.boxplot ?? {};
  const { groupNames, grouped, all, drawn } = boxplotBoxes(cfg);
  const nG = groupNames.length;
  const boxes = grouped[0] ?? []; // group 0 keeps the single-group code path's shape
  const fmt = resolveFormat(all, cfg.numberFormat);
  const H = !!cfg.horizontal;
  let frame = H
    ? computeFrameHorizontal(cfg, style, { ...decor, seriesLabels: false, totals: false })
    : computeFrame(cfg, style, { ...decor, seriesLabels: false }, []).frame;
  // Grouped boxplots draw a wrapping group legend at the top (below); the frame
  // above only budgets its usual single-row headroom, so reserve one extra row
  // per wrapped row and push the plot down. A one-row legend reserves nothing
  // extra, keeping existing decks byte-identical.
  if (nG > 1) {
    const extraRows = legendRowCount(groupNames, fs, frame.x, cfg.width - 4) - 1;
    if (extraRows > 0) {
      const band = extraRows * fs * 1.6;
      frame = { ...frame, y: frame.y + band, h: frame.h - band };
    }
  }
  // One shared value scale across every box — that is the point of putting
  // them in one chart.
  // Data-driven domain (no forced zero): a boxplot of scores 40–95 must not be
  // squashed against 0 — matches violin/candlestick. cfg.scale still overrides.
  const lo = drawn.length ? minOf(drawn) : 0;
  const hi = drawn.length ? maxOf(drawn) : 1;
  const scale = valueScale(frame, lo, hi, cfg.scale, undefined, undefined, false); // no forced zero: data-driven domain
  // Value coordinate along the value axis (x when horizontal, y otherwise).
  // Clamped like `scale.toY` is — see the note on the same map in column.ts.
  const qOf = H
    ? (v: number) =>
        frame.x + Math.max(0, Math.min(frame.w, ((v - scale.min) / (scale.max - scale.min || 1)) * frame.w))
    : scale.toY;

  const catStart = H ? frame.y : frame.x;
  const catLen = H ? frame.h : frame.w;
  const slotW = catLen / Math.max(1, n);
  const boxW = slotW * 0.55;
  const centers = data.categories.map((_, c) => catStart + slotW * (c + 0.5));

  // Summary rows (Min/Q1/…) are anatomy, not series — never a legend.
  const nodes: SceneNode[] = H
    ? horizontalChrome(
        cfg,
        style,
        { ...decor, totals: false, seriesLabels: false },
        frame,
        centers,
        scale,
        (v) => qOf(v) - frame.x,
      )
    : chromeNodes(cfg, style, { ...decor, seriesLabels: false }, frame, centers, scale);
  const columnTop: number[] = [];

  /** Category-axis point p + value-axis point q → chart coordinates. */
  const pt = (p: number, q: number) => (H ? { x: q, y: p } : { x: p, y: q });
  const segLine = (
    p1: number,
    q1: number,
    p2: number,
    q2: number,
    weight: number,
    stroke: string,
    nm: string,
  ): SceneNode => {
    const a = pt(p1, q1);
    const b = pt(p2, q2);
    return { kind: "line", x1: a.x, y1: a.y, x2: b.x, y2: b.y, stroke, strokeWidth: weight, name: nm };
  };

  // Group legend (grouped mode only) — wrapping, via the shared row.
  if (nG > 1) {
    nodes.push(
      ...legendRow(cfg, style, frame.x, titleHeight(cfg, style), {
        maxX: cfg.width - 4,
        entries: groupNames.map((g, gi) => ({ label: g, color: seriesColor(style, gi), name: `legend-${gi}` })),
      }),
    );
  }

  grouped.forEach((groupBoxes, gi) =>
    groupBoxes.forEach((b, c) => {
      // Sub-slot offset: groups sit side by side within the category slot.
      const subW = nG > 1 ? (slotW * 0.7) / nG : boxW;
      const p0 = nG > 1 ? centers[c] - (slotW * 0.7) / 2 + subW * (gi + 0.5) : centers[c];
      const gBoxW = nG > 1 ? Math.max(4, subW - 3) : boxW;
      const gCapW = gBoxW * 0.45;
      const p = p0;
      if (!b) {
        if (gi === 0) columnTop.push(frame.y + frame.h);
        return;
      }
      const gSuffix = nG > 1 ? `-g${gi}` : "";
      const fill = nG > 1 ? seriesColor(style, gi) : seriesColor(style, 0, data.series.find((s) => s.color)?.color);
      const boxFill = lerpColor(style.background, fill, 0.22);
      const qQ1 = qOf(b.q1);
      const qQ3 = qOf(b.q3);
      const qMed = qOf(b.median);
      if (gi === 0) columnTop.push(H ? p - gBoxW / 2 : Math.min(qOf(b.max), ...b.outliers.map(qOf)));
      const boxLo = Math.min(qQ1, qQ3);
      nodes.push(
        // Whiskers with caps.
        segLine(p, qOf(b.max), p, qQ3, 0.75, style.axis, `whisker-hi-${c}${gSuffix}`),
        segLine(p, qQ1, p, qOf(b.min), 0.75, style.axis, `whisker-lo-${c}${gSuffix}`),
        segLine(p - gCapW / 2, qOf(b.max), p + gCapW / 2, qOf(b.max), 0.75, style.axis, `cap-hi-${c}${gSuffix}`),
        segLine(p - gCapW / 2, qOf(b.min), p + gCapW / 2, qOf(b.min), 0.75, style.axis, `cap-lo-${c}${gSuffix}`),
      );
      // Q1–Q3 box (notched at the median CI when raw samples give us n).
      const notch = !!opts.notch && !opts.meanSd && !!b.samples && b.samples.length >= 3;
      let medHalf = gBoxW / 2;
      if (notch) {
        const ci = (1.57 * (b.q3 - b.q1)) / Math.sqrt(b.samples!.length);
        const qMedHi = qOf(Math.min(b.q3, b.median + ci));
        const qMedLo = qOf(Math.max(b.q1, b.median - ci));
        const nx = gBoxW * 0.25;
        medHalf = nx;
        nodes.push({
          kind: "polygon",
          points: [
            pt(p - gBoxW / 2, qQ3),
            pt(p + gBoxW / 2, qQ3),
            pt(p + gBoxW / 2, qMedHi),
            pt(p + nx, qMed),
            pt(p + gBoxW / 2, qMedLo),
            pt(p + gBoxW / 2, qQ1),
            pt(p - gBoxW / 2, qQ1),
            pt(p - gBoxW / 2, qMedLo),
            pt(p - nx, qMed),
            pt(p - gBoxW / 2, qMedHi),
          ],
          fill: boxFill,
          stroke: fill,
          strokeWidth: 1,
          name: `box-${c}${gSuffix}`,
        });
      } else {
        nodes.push(
          H
            ? {
                kind: "rect",
                x: boxLo,
                y: p - gBoxW / 2,
                w: Math.abs(qQ1 - qQ3),
                h: gBoxW,
                fill: boxFill,
                stroke: fill,
                strokeWidth: 1,
                name: `box-${c}${gSuffix}`,
              }
            : {
                kind: "rect",
                x: p - gBoxW / 2,
                y: boxLo,
                w: gBoxW,
                h: Math.abs(qQ1 - qQ3),
                fill: boxFill,
                stroke: fill,
                strokeWidth: 1,
                name: `box-${c}${gSuffix}`,
              },
        );
      }
      nodes.push(segLine(p - medHalf, qMed, p + medHalf, qMed, 1.75, fill, `median-${c}${gSuffix}`));
      // Settled before the median number below, which has to clear it. The nearest
      // other marker along the category axis is the next slot, or the next group's
      // sub-slot, and a lone box has none.
      const pitch = nG > 1 ? subW : n > 1 ? slotW : Infinity;
      const mean =
        b.mean != null ? meanMark(cfg, style, frame, pt(p, qOf(b.mean)), pitch, H, `mean-${c}${gSuffix}`) : null;
      if (mean) nodes.push(mean);
      const meanFs = mean?.fontSize ?? 0;
      // Jitter overlay: every raw observation as a deterministically offset dot
      // over the box. Golden-ratio spacing keeps it byte-deterministic (no RNG).
      const jitter = !!opts.jitter && !!b.samples;
      if (jitter) {
        const jw = gBoxW * 0.8;
        b.samples!.forEach((v, j) => {
          const off = ((((j + 1) * 0.6180339887) % 1) - 0.5) * jw;
          const d = pt(p + off, qOf(v));
          nodes.push({
            kind: "ellipse",
            cx: d.x,
            cy: d.y,
            rx: 1.5,
            ry: 1.5,
            fill,
            stroke: style.background,
            strokeWidth: 0.5,
            name: `dot-${c}${gSuffix}-${j}`,
          });
        });
      }
      // Separate outlier dots — skipped when jitter already draws every point.
      if (!jitter) {
        b.outliers.forEach((v, i) => {
          const o = pt(p, qOf(v));
          nodes.push({
            kind: "ellipse",
            cx: o.x,
            cy: o.y,
            rx: 2.2,
            ry: 2.2,
            fill,
            name: `outlier-${c}-${i}${gSuffix}`,
          });
        });
      }
      if (decor.segmentLabels && nG === 1) {
        const label = formatNumber(b.median, fmt);
        // Upright the number sits ABOVE the median line, and a box whose median
        // is near the top of the plot has nothing above it but the edge of the
        // chart — 10pt over the top of a 300x60 frame at 24pt. Fitted to the
        // band the median leaves and dropped when that band cannot carry a
        // readable number, which is what `gBoxW` already does sideways.
        /**
         * AND THE MEAN MARKER IS IN THAT BAND TOO, when the mean sits above the
         * median.
         *
         * The label was fitted against the title alone, and the shipped
         * showcase is what caught it: `Boxplot from raw samples` draws both
         * marks, and three of its boxes put the `×` exactly where the median's
         * number goes. The two are pushed by the same loop a few lines apart and
         * neither knew about the other.
         *
         * Its INK bottom, not its box. The `×` is a `valign: middle` glyph in a
         * `fs * 1.4` box, so its ink runs to `y + fs * 0.57`; measuring the box
         * would drop labels that clear the mark perfectly well. That is the
         * distinction the gauge total and the gantt nudge each got wrong once.
         *
         * Only when the mean is ABOVE the median. Below it the marker is inside
         * the box and nowhere near this label, and taking the max with it would
         * collapse the band and drop a number for no reason.
         *
         * A MUTANT SURVIVES HERE AND IT IS NOT EQUIVALENT — the gates simply
         * cannot see what it costs. Removing this line leaves every overlap
         * count identical, because the drop below catches the same cases. What
         * changes is how many labels SURVIVE: 19 median numbers in the shipped
         * deck with it, 17 without. This shrinks a label that the drop would
         * delete, which is the whole difference between fitting and refusing,
         * and no overlap gate can measure it. Counted by hand rather than left
         * to a mutation run's verdict.
         */
        // A marker that was not drawn (see `meanMark`) is nothing to clear.
        const meanY = b.mean != null && meanFs > 0 ? qOf(b.mean) : null;
        const meanInkBottom = meanY != null && meanY < qMed ? meanY + meanFs * MEAN_INK_DOWN : -Infinity;
        const medBandTop = Math.max(titleInkBottom(cfg, style), meanInkBottom);
        const medFs = H ? fs * 0.9 : aboveMarkFontSize(fs * 0.9, qMed, medBandTop, 1.3);
        const medScale = medFs / (fs * 0.9);
        /**
         * AND THE FIT IS TESTED ON THE PLACED LABEL, not on the band.
         *
         * Shrinking to the band clears the marker where the two are far enough
         * apart, and cannot where they are not: the label is anchored just
         * above the median line and the `×` sits at the mean, so a mean within
         * a line-height of the median leaves nowhere to be. Two of the
         * showcase's three collisions went to the shrink and the third did not.
         *
         * So the drop is the answer, and it is the median NUMBER that goes —
         * the median LINE is still drawn, and the `×` is the only thing that
         * says where the mean is. This is remedy three from the pattern this
         * repo keeps meeting: where a fit and a placement disagree, test the fit
         * on the placed position.
         */
        const medInkClearsMean = (() => {
          if (H || meanY == null || medFs <= 0) return true;
          const boxY = qMed - fs * 1.5 * medScale;
          const base = boxY + fs * 1.3 * medScale - medFs * 0.25;
          return base + medFs * 0.21 <= meanY - meanFs * 0.8 || base - medFs * 0.8 >= meanY + meanFs * MEAN_INK_DOWN;
        })();
        if (medFs > 0 && medInkClearsMean && gBoxW >= fs * 1.2 && (H || gBoxW >= textWidth(label, medFs) + 4)) {
          nodes.push(
            H
              ? {
                  kind: "text",
                  x: qMed - 30,
                  y: p - gBoxW / 2 - fs * 1.35,
                  w: 60,
                  h: fs * 1.3,
                  text: label,
                  fontSize: medFs,
                  color: style.text,
                  align: "center",
                  valign: "bottom",
                  name: `median-label-${c}`,
                }
              : {
                  kind: "text",
                  x: p - gBoxW / 2,
                  // As a RATIO of the unshrunk size — exactly 1 where nothing
                  // needed shrinking, so the box and the font move together and
                  // the geometry is byte-identical on every chart that already
                  // fits. Written as `medFs * 1.667` first, and that rounding
                  // moved the boxplot's snapshot by a hundredth of a point on a
                  // chart nothing was wrong with.
                  y: qMed - fs * 1.5 * medScale,
                  w: gBoxW,
                  h: fs * 1.3 * medScale,
                  text: label,
                  fontSize: medFs,
                  color: style.text,
                  align: "center",
                  valign: "bottom",
                  name: `median-label-${c}`,
                },
          );
        }
      }
    }),
  );

  if (H) {
    nodes.push({
      kind: "line",
      x1: frame.x,
      y1: frame.y,
      x2: frame.x,
      y2: frame.y + frame.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
  } else {
    nodes.push(baselineNode(frame, frame.y + frame.h, style));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => boxW),
      columnTop,
      columnValue: boxes.map((b) => b?.median ?? 0),
      baselineY: H ? frame.x : frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: H ? undefined : qOf,
      // Publish the sideways map too, for the reason the column and waterfall
      // layouts publish theirs: `bandNodes` reads whichever axis carries the
      // value, and with only the vertical one here a rotated boxplot dropped
      // every `decorations.bands` zone in silence — the band was computed,
      // found no map, and returned nothing.
      valueToX: H ? qOf : undefined,
    },
  };
}

/** Value extent for Same Scale, including whisker ends and outliers. */
export function boxplotExtent(cfg: ChartConfig): { min: number; max: number } | null {
  // Report the boxes, not the raw rows: mean±SD whiskers sit at mean ± 2·SD,
  // outside the samples, and "Same scale" turns this extent into a hard scale
  // override — understating it pushed those whiskers off the plot.
  const { drawn } = boxplotBoxes(cfg);
  return drawn.length ? { min: minOf(drawn), max: maxOf(drawn) } : null;
}
