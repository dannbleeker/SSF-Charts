import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors, Series } from "../types";
import { contrastInk, textWidth, type SceneNode, type TextNode } from "../scene";
import { clipToWidth } from "../elements";
import {
  authoredDecimals,
  formatNumber,
  niceTicks,
  resolveFormat,
  resolveAxisFormat,
  segmentLabel,
  axisTickLabel,
} from "../format";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import {
  printsOnTitle,
  titleInkBottom,
  aboveMarkFontSize,
  baselineNode,
  breakMarkerNodes,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  legendWrapWalk,
  seriesLegendLabels,
  legendLabelsFor,
  logFloor,
  titleHeight,
  valueScale,
  bandFontSize,
  tickGapScale,
  markInFrame,
  MIN_LABEL_FS,
  type Frame,
  type ValueScale,
  horizontalLegendFits,
} from "./frame";
// Combo base kinds. These modules import back from column (LayoutResult /
// horizontalChrome), but the calls happen at runtime so the ESM cycle resolves.
import { layoutWaterfall, detailParents } from "./waterfall";
import { layoutMekko } from "./mekko";
import { layoutLine } from "./line";
import { columnHasData, columnNegativeTotal, columnPositiveTotal, columnSignedTotal } from "./totals";
import { maxOf, minOf } from "../agg";
import { tightBox } from "../collide";

/**
 * The nearest drawn category before / after `c`, or -1.
 *
 * A line skips its gaps — a null value draws no mark and no label — so "the
 * label beside this one" is not `c - 1`. Measuring against a category that was
 * never drawn would shrink a label to clear something that is not there.
 */
function prevDrawn(drawn: boolean[], c: number): number {
  for (let i = c - 1; i >= 0; i--) if (drawn[i]) return i;
  return -1;
}
function nextDrawn(drawn: boolean[], c: number): number {
  for (let i = c + 1; i < drawn.length; i++) if (drawn[i]) return i;
  return -1;
}

export interface LayoutResult {
  nodes: SceneNode[];
  anchors: LayoutAnchors;
}

/** Minimum segment thickness (relative to font size) before its label is hidden. */
const LABEL_FIT = 1.25;

/** How one series' rectangles are painted, once IBCS notation is applied. */
interface MarkPaint {
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  pattern?: Series["pattern"];
}

/**
 * IBCS scenario notation restyles a series by the data's nature: PY a lighter
 * solid, PL/BU an outlined/hollow bar, FC a hatch. AC/none stay solid, carrying
 * only the series' own `pattern`.
 *
 * Shared by the plotted segment and its legend chip: a hollow PL bar keyed by a
 * solid block — or two same-coloured series told apart only by a hatch, keyed by
 * two identical squares — hides exactly the redundant encoding the notation
 * exists to provide for greyscale and colour-blind readers.
 */
function markPaint(
  style: ChartStyle,
  fill: string,
  s: Pick<Series, "scenario" | "pattern">,
  baseStroke?: string,
  baseStrokeWidth?: number,
): MarkPaint {
  const paint: MarkPaint = { fill, stroke: baseStroke, strokeWidth: baseStrokeWidth, pattern: s.pattern };
  if (s.scenario === "PY") {
    paint.fill = lerpColor(style.background, fill, 0.5);
  } else if (s.scenario === "PL" || s.scenario === "BU") {
    paint.fill = "none";
    paint.stroke = fill;
    paint.strokeWidth = 1.5;
  } else if (s.scenario === "FC") {
    // SVG keeps the true IBCS hatch, but rect.pattern is SVG-only — both
    // PowerPoint renderers drop it, which made FC pixel-identical to AC in
    // the actual deliverable. Back the hatch with a fill/border encoding all
    // three renderers can express: a light tint (distinct from PY's heavier
    // 0.5) plus a series-coloured edge (distinct from PL/BU's hollow bar).
    paint.pattern = "diagonal";
    paint.fill = lerpColor(style.background, fill, 0.75);
    paint.stroke = fill;
    paint.strokeWidth = 1;
  }
  return paint;
}

/**
 * The value a CLUSTERED category publishes to every decoration that reads one.
 *
 * This was `Math.max(0, ...values)`, and the zero seed is the bug: a category
 * whose bars are all negative published 0. `columnValue` is what mean lines,
 * difference arrows and the rest print, so a cash-flow chart of
 * [-40, -20, -10, 60] drew its mean line ABOVE the baseline and labelled it
 * "Ø 15" — the mean of [0, 0, 0, 60] — while the same data as `stacked`, `line`
 * or `area` correctly said "Ø -2.5". A difference arrow between two all-negative
 * categories came out zero-length and labelled "0", sitting beside the very
 * totals that contradict it.
 *
 * Zero is right for the drawn column's TOP — a bar hanging below the baseline
 * tops out at the baseline, and `columnTop` still clamps that way. It is wrong
 * for the column's VALUE, which is what a reader is being told the category is
 * worth. So: the extreme in the direction the column actually goes — the
 * highest bar when anything rises, the lowest when nothing does.
 */
export function clusteredTopValue(values: number[]): number {
  const finite = values.filter((v) => Number.isFinite(v));
  if (!finite.length) return 0;
  const hi = Math.max(...finite);
  return hi > 0 ? hi : Math.min(...finite);
}

/**
 * Stacked / clustered / 100% column charts — and, following think-cell's
 * "a bar chart is a rotated column chart" model, the same layouts in
 * horizontal orientation when cfg.horizontal is set.
 */
export function layoutColumns(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data, kind } = cfg;
  const n = data.categories.length;
  const stacked = kind !== "clustered";
  const pct = kind === "stacked100";
  const H = !!cfg.horizontal;
  // Clustered-stacked: series carry stack indices (blank datasheet rows).
  const stackIds = [...new Set(data.series.map((s) => s.stack ?? 0))].sort((a, b) => a - b);
  const nStacks = stacked && !pct ? stackIds.length : 1;
  const stackPos = new Map(stackIds.map((id, i) => [id, i]));

  const vframe = H ? null : computeFrame(cfg, style, decor, decor.seriesLabels ? data.series.map((s) => s.name) : []);
  const frame = H ? computeFrameHorizontal(cfg, style, decor) : vframe!.frame;
  const fs = style.fontSize;

  // Category slots run along x (vertical) or y (horizontal).
  const catStart = H ? frame.y : frame.x;
  const catLen = H ? frame.h : frame.w;
  const slotLen = catLen / Math.max(1, n);
  // Excel-style gap width: gap between columns as a % of column width.
  // Default 50 reproduces think-cell's 2/3-of-slot columns.
  const gapWidth = Math.max(0, Math.min(500, cfg.gapWidth ?? 50));
  const colThick = slotLen / (1 + gapWidth / 100);
  // Excel-style clustered overlap (−100…100): fraction each bar overlaps its
  // neighbour. 0 = edge to edge (the historical default).
  const overlapFrac = Math.max(-100, Math.min(100, cfg.overlap ?? 0)) / 100;
  const centers = Array.from({ length: n }, (_, i) => catStart + slotLen * (i + 0.5));

  const posTotals = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  const negTotals = data.categories.map((_, c) => columnNegativeTotal(data.series, c));
  const signedTotals = data.categories.map((_, c) => columnSignedTotal(data.series, c));
  // Per-category denominator for 100% charts (think-cell's "100%=" row).
  const denominators = data.categories.map((_, c) => {
    const d = data.hundredPercent?.[c];
    if (d != null && d > 0) return d;
    // An all-negative category has no positive total; normalise against the
    // negative magnitude so its segments fill down to -100% instead of
    // collapsing to zero (v / 0).
    return posTotals[c] > 0 ? posTotals[c] : -negTotals[c];
  });
  const fmt = resolveFormat(
    [...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)), ...signedTotals],
    cfg.numberFormat,
  );

  // Per-stack totals (clustered-stacked scales to the tallest single stack).
  const stackPosTotals = (c: number, id: number) =>
    data.series.reduce((a, s) => a + ((s.stack ?? 0) === id ? Math.max(0, s.values[c] ?? 0) : 0), 0);
  const stackNegTotals = (c: number, id: number) =>
    data.series.reduce((a, s) => a + ((s.stack ?? 0) === id ? Math.min(0, s.values[c] ?? 0) : 0), 0);

  const allValues = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const logOn = !stacked && !H && !!cfg.logScale;
  let dataMin: number, dataMax: number;
  if (pct) {
    dataMin = 0;
    dataMax = 1;
  } else if (stacked && nStacks > 1) {
    // Folded, not spread: this array is categories x stacks (up to 4096 x 256),
    // which overflows the argument list and throws RangeError. The sibling
    // branches below spread at most MAX_CATEGORIES or MAX_SERIES values, so they
    // stay safe as-is.
    dataMin = minOf(
      data.categories.flatMap((_, c) => stackIds.map((id) => stackNegTotals(c, id))),
      0,
    );
    dataMax = maxOf(
      data.categories.flatMap((_, c) => stackIds.map((id) => stackPosTotals(c, id))),
      0,
    );
  } else if (stacked) {
    dataMin = Math.min(0, ...negTotals);
    dataMax = Math.max(0, ...posTotals);
  } else {
    dataMin = minOf(allValues, 0);
    dataMax = maxOf(allValues, 0);
  }
  // 100% charts: negatives are shares below the zero line. The axis drops to
  // the most-negative column share (0 when all data is positive → unchanged).
  const pctNegMin = pct
    ? Math.min(0, ...data.categories.map((_, c) => (denominators[c] > 0 ? negTotals[c] / denominators[c] : 0)))
    : 0;
  // …and the ceiling is not always 100%. An authored "100% =" denominator can be
  // SMALLER than the column it normalises — a multi-select survey ("100% = 500
  // respondents", answers totalling 180%), or a typo. Pinning max at 1 painted
  // those segments outside the plot and outside the chart frame; `waffle` already
  // handles the same input. Left at exactly 1 whenever no column overflows, so
  // every well-formed 100% chart is untouched.
  const pctPosMax = pct
    ? Math.max(1, ...data.categories.map((_, c) => (denominators[c] > 0 ? posTotals[c] / denominators[c] : 0)))
    : 1;
  const scale: ValueScale = pct
    ? {
        min: pctNegMin,
        max: pctPosMax,
        percent: true,
        // Clamped to the domain, the way valueScale already filters its own:
        // niceTicks rounds OUTWARD, so it returned ticks below pctNegMin, and
        // toY mapped those past the plot — a gridline at y=301.6 on a 300pt
        // canvas, with its label sitting on the category row.
        ticks:
          pctNegMin < 0 || pctPosMax > 1
            ? niceTicks(pctNegMin, pctPosMax, 5).filter((t) => t >= pctNegMin - 1e-9 && t <= pctPosMax + 1e-9)
            : [0, 0.25, 0.5, 0.75, 1],
        toY: (v: number) => frame.y + frame.h - ((v - pctNegMin) / (pctPosMax - pctNegMin)) * frame.h,
      }
    : valueScale(
        frame,
        logOn ? logFloor(allValues, dataMin) : dataMin,
        dataMax,
        cfg.scale,
        H ? undefined : cfg.axisBreak,
        logOn,
      );

  // Value coordinate: distance along the value axis from the scale minimum.
  // Vertical charts route through toY so axis breaks apply; horizontal stays linear.
  //
  // CLAMPED either way. `toY` clips to the plot and says why (frame.ts: "every
  // charting tool draws a value above the axis maximum as a bar reaching the top
  // of the plot"), so the vertical branch inherited that and the horizontal one
  // — its own linear map — did not. Pin the axis to 0–100 on revenue data, which
  // the pane invites since "Axis scale min / max" are free-text boxes, and the
  // same config that draws a full-height bar upright drew a 30,128pt one
  // sideways: off the slide, and past the OOXML coordinate limit that makes
  // PowerPoint offer to repair the file.
  const valLen = H ? frame.w : frame.h;
  const qOf = H
    ? (v: number) => Math.max(0, Math.min(valLen, ((v - scale.min) / (scale.max - scale.min || 1)) * valLen))
    : (v: number) => frame.y + frame.h - scale.toY(v);
  /** Rect spanning [v0, v1] on the value axis at category position/thickness. */
  const segRect = (catPos: number, thick: number, v0: number, v1: number) => {
    const q0 = Math.min(qOf(v0), qOf(v1));
    const q1 = Math.max(qOf(v0), qOf(v1));
    return H
      ? { x: frame.x + q0, y: catPos - thick / 2, w: q1 - q0, h: thick }
      : { x: catPos - thick / 2, y: frame.y + frame.h - q1, w: thick, h: q1 - q0 };
  };

  const nodes: SceneNode[] = H
    ? horizontalChrome(cfg, style, decor, frame, centers, scale, qOf)
    : chromeNodes(cfg, style, decor, frame, centers, scale);
  const zeroQ = qOf(0);
  const y0 = H ? frame.x + zeroQ : frame.y + frame.h - zeroQ;
  const columnTop: number[] = [];
  /**
   * The DATA value `columnTop` is drawn at, per category.
   *
   * The pixel anchor and the number a decoration prints have to describe the
   * same mark, and on a clustered chart they did not: `columnValue` published
   * the SUM of every series while `columnTop` is the tallest single bar, which
   * is drawn nowhere near it. Its sibling `seriesLevels` was fixed for exactly
   * this ("clustered, each bar stands on the baseline") and this one was not
   * swept with it.
   */
  const drawnTopValue: number[] = [];
  const seriesLevels: number[][] = [];
  /** Segment mid-position of the last category per series, for series labels. */
  const lastSegMid: (number | null)[] = data.series.map(() => null);
  /** Cumulative segment boundaries per column (value units), for connectors. */
  const posBounds: number[][] = [];
  const negBounds: number[][] = [];

  for (let c = 0; c < n; c++) {
    // Running positive/negative levels per stack group (value units).
    const ups = stackIds.map(() => 0);
    const downs = stackIds.map(() => 0);
    const levels: number[] = data.series.map(() => 0);
    const stackThick = colThick / nStacks;
    // Clustered bars fill the column; overlap widens each bar and shrinks the
    // stride so they overlap (or gap). At overlap 0 this is colThick / nBars.
    const nBars = Math.max(1, data.series.length);
    const barW = stacked ? stackThick : colThick / (1 + (nBars - 1) * (1 - overlapFrac));
    const barStep = stacked ? stackThick : barW * (1 - overlapFrac);
    const barThick = barW;
    /**
     * A bar's drawn thickness, after the 1pt gap that separates it from its
     * neighbour.
     *
     * That gap is chrome, and a bar too thin to pay for it went NEGATIVE:
     * `barThick - 1` is -0.3 for a clustered chart 14 points wide, or one 12x9,
     * and a negative-width rect is not a thin bar — SVG drops it and PowerPoint
     * clamps it to a sliver, so the chart loses its bars entirely while every
     * label around them still says what they are worth.
     *
     * The gap yields instead, down to half the bar. Identical for any bar 2pt or
     * wider, which is every chart at a sane size.
     */
    const withGap = (t: number) => Math.max(t * 0.5, t - 1);
    // think-cell's Segment Order: stacking order within this column.
    const order = data.series.map((_, i) => i);
    if (cfg.segmentOrder === "reverse") order.reverse();
    else if (cfg.segmentOrder === "ascending" || cfg.segmentOrder === "descending") {
      const sign = cfg.segmentOrder === "ascending" ? 1 : -1;
      order.sort((a, b) => sign * ((data.series[a].values[c] ?? 0) - (data.series[b].values[c] ?? 0)));
    }

    order.forEach((si, position) => {
      const s = data.series[si];
      const raw = s.values[c];
      let v = raw ?? 0;
      if (pct) v = denominators[c] > 0 ? v / denominators[c] : 0;
      let r: { x: number; y: number; w: number; h: number } | null = null;
      const fill = seriesColor(style, si, s.colors?.[c] ?? s.color);

      const sp = stackPos.get(s.stack ?? 0) ?? 0;
      // The `!H` that used to be here forced every clustered chart back to
      // plain rectangles the moment the rotation toggle went on — and a
      // Cleveland dot plot and a dumbbell range chart are NORMALLY horizontal,
      // which is the whole point of them: long category labels run down the
      // left. `barStyle: "range"` sideways gave two full-length bars from zero
      // instead of dot–line–dot, so the connector carrying the entire meaning
      // was gone and the endpoints read as two independent magnitudes. Stacked
      // is still excluded, because a stack has no single point to mark.
      const barStyle = stacked ? "bar" : (decor.barStyle ?? "bar");
      /**
       * A MEASURED ZERO IS A POINT; A ZERO-HEIGHT BAR IS NOTHING TO DRAW.
       *
       * `v !== 0` is right for a rectangle — a bar of no height is invisible
       * whether or not it is emitted, so skipping it costs nothing. It is wrong
       * for a DOT or a LOLLIPOP, where the mark is a point AT the value and
       * zero is an ordinary place on the axis.
       *
       * Measured 2026-09-28, `[10, 0, 30]` against `[10, null, 30]` at
       * `barStyle: "dot"` and `"lollipop"`: both produced the identical scene
       * for category 1 — the axis heading and nothing else. No dot, no stem, no
       * label. So a zero somebody measured and a cell nobody filled in were
       * indistinguishable on the chart, which is this repo's blank-is-not-zero
       * rule pointing the other way: the harm there is inventing data, the harm
       * here is silently discarding it.
       *
       * `raw != null` still carries the blank, and stacked still forces "bar",
       * so a stack's zero segment is unaffected.
       */
      const marksAPoint = barStyle === "dot" || barStyle === "lollipop";
      if (raw != null && (v !== 0 || marksAPoint)) {
        if (stacked) {
          const catPos = nStacks > 1 ? centers[c] - colThick / 2 + (sp + 0.5) * stackThick : centers[c];
          const thick = nStacks > 1 ? withGap(stackThick) : colThick;
          if (v >= 0) {
            r = segRect(catPos, thick, ups[sp], ups[sp] + v);
            ups[sp] += v;
            // Key the boundary by SERIES, not push order: a zero/null segment
            // pushes nothing, which used to shift every later boundary down an
            // index and join mismatched series between columns.
            if (nStacks === 1) (posBounds[c] ??= [])[si] = ups[sp];
          } else {
            r = segRect(catPos, thick, downs[sp] + v, downs[sp]);
            downs[sp] += v;
            if (nStacks === 1) (negBounds[c] ??= [])[si] = downs[sp];
          }
        } else {
          const pos = centers[c] - colThick / 2 + barW / 2 + position * barStep;
          r = segRect(pos, withGap(barThick), 0, v);
        }
      }
      // Where THIS series' mark tops out, in value units — what a level-anchored
      // decoration (difference arrow, callout) points at. Stacked, that is the
      // cumulative level; clustered, each bar stands on the baseline, so it is
      // the bar's own value. Publishing the cumulative sum for clustered too left
      // it 0 for every series (ups/downs only advance in the stacked branch), so
      // the arrow silently collapsed onto the axis instead of the bar top.
      levels[si] = stacked ? ups[sp] + downs[sp] : v;
      if (!r) return;

      if (barStyle !== "bar") {
        // Lollipop / dot / dumbbell-range rendering for clustered charts:
        // the value point is a dot; lollipops add a stem from the baseline;
        // range connects the two series' dots with a line (drawn once).
        // Two coordinates rather than an x and a y: `cross` is the position on
        // the CATEGORY axis (a bar's x on a column chart, its y on a bar) and
        // `val` is the position on the VALUE axis. The rest of the block places
        // marks in those terms and maps them once, which is what lets the same
        // dot, stem and connector serve both orientations.
        const cross = barStyle === "range" ? centers[c] : H ? r.y + r.h / 2 : r.x + r.w / 2;
        const val = H ? (v >= 0 ? r.x + r.w : r.x) : v >= 0 ? r.y : r.y + r.h;
        const dotR = 4;
        /** A line along the value axis, at `cross` on the category axis. */
        const alongValue = (v1: number, v2: number) =>
          H ? { x1: v1, y1: cross, x2: v2, y2: cross } : { x1: cross, y1: v1, x2: cross, y2: v2 };
        if (barStyle === "lollipop") {
          nodes.push({
            kind: "line",
            // `y0` is the baseline in whichever direction the value axis runs.
            ...alongValue(y0, val),
            stroke: fill,
            strokeWidth: 1.5,
            name: `stem-${si}-${c}`,
          });
        }
        if (barStyle === "range" && si === 1 && data.series[0].values[c] != null) {
          const other = H
            ? frame.x + qOf(data.series[0].values[c]!)
            : frame.y + frame.h - qOf(data.series[0].values[c]!);
          nodes.push({
            kind: "line",
            ...alongValue(other, val),
            stroke: style.mutedText,
            strokeWidth: 1.5,
            name: `range-${c}`,
          });
        }
        nodes.push({
          kind: "ellipse",
          cx: H ? val : cross,
          cy: H ? cross : val,
          rx: dotR,
          ry: dotR,
          fill,
          stroke: style.background,
          strokeWidth: 1,
          name: `seg-${si}-${c}`,
        });
        if (c === n - 1) lastSegMid[si] = H ? cross : val;
        if (decor.segmentLabels) {
          const label = formatNumber(raw!, fmt);
          /**
           * BOUND BY THE ROW, sideways — the bound every other horizontal label
           * in this file got and this one did not.
           *
           * The label is centred on its dot in a box `fs * 1.4` tall. Turned
           * sideways that box sits on the category's row, so once the font
           * outgrows the row the labels of adjacent categories are drawn through
           * each other and the last one is drawn off the foot of the chart —
           * 5.4pt below a 300x60 lollipop at 18pt, and the same on three other
           * frames. Upright the label sits BESIDE its dot with the next category
           * a column away, so nothing there wants this and `fs` stands.
           *
           * Dropped rather than clamped past the floor: `bandFontSize` answers 0
           * below `MIN_LABEL_FS`, and a two-point number over a dot is ink, not
           * a value anyone can read.
           */
          const dotFs = H ? bandFontSize(fs, r.h, 1.4) : fs;
          if (dotFs > 0)
            nodes.push({
              kind: "text",
              // Clear of the dot along the value axis, centred on it across.
              x: (H ? val : cross) + dotR + 2,
              y: (H ? cross : val) - dotFs * 0.7,
              w: textWidth(label, dotFs) + 4,
              h: dotFs * 1.4,
              text: label,
              fontSize: dotFs,
              color: style.text,
              align: "left",
              valign: "middle",
              name: `label-${si}-${c}`,
            });
        }
        return;
      }

      // Transparent "no-fill" segment: it still occupies the stack (the level
      // was already advanced) but draws nothing, floating the segments above.
      if (fill === "transparent") return;
      const seg = markPaint(style, fill, s, style.background, stacked ? 0.75 : 0);
      nodes.push({
        kind: "rect",
        ...r,
        fill: seg.fill,
        stroke: seg.stroke,
        strokeWidth: seg.strokeWidth,
        pattern: seg.pattern,
        name: `seg-${si}-${c}`,
      });
      if (c === n - 1) lastSegMid[si] = H ? r.x + r.w : r.y + r.h / 2;

      if (decor.segmentLabels) {
        // think-cell's label-content dropdown: value / % / series / category.
        const label = segmentLabel(decor.labelContent ?? (pct ? ["percent"] : ["value"]), {
          value: raw!,
          /**
           * SIGNED, because the segment is drawn signed.
           *
           * `Math.max(0, raw!)` sent every negative segment's fraction to 0, so
           * a returns row of -25 against a positive total of 100 was DRAWN at a
           * quarter of the stack's height and labelled "0%". The picture was
           * right and the caption said the outflow never happened.
           *
           * `posTotals[c]` stays the denominator: it is the height the stack
           * actually reaches, so -25/100 is exactly the share the eye measures
           * off the drawing. The clamp belongs to geometry, where a negative
           * must not shorten the positive stack — it had no business reaching a
           * label, which is the same split `columnHasData` makes for totals.
           */
          fraction: pct ? v : posTotals[c] > 0 ? raw! / posTotals[c] : null,
          series: s.name,
          category: data.categories[c],
          fmt,
          // An authored decimals count reaches the value axis through
          // `resolveAxisFormat` and has to reach these too, or a 100% chart
          // labels its axis "25.0%" and its segments "33%" — see `segmentLabel`.
          percentDecimals: authoredDecimals(cfg.numberFormat),
        });
        const along = H ? r.w : r.h; // extent along the value axis
        const across = H ? r.h : r.w;
        // The upright test has NO slack, and the two points it used to allow are
        // the whole defect. The label is centred on the bar in a box `r.w + 8`
        // wide, so `textWidth <= across + 2` lets its ink stand a point proud of
        // the mark on each side — and in a CLUSTER the next bar's label is doing
        // the same thing from the other direction. At 160x120 the bars in a
        // cluster are a point or two apart and the two labels meet.
        //
        // Same shape as the mekko's `width + 8` against a 2pt column gap: a box
        // wider than the mark it labels bleeds into the neighbour, and the room
        // a label has is the mark itself. A label that no longer fits is dropped
        // rather than drawn over its neighbour — one unreadable pair is worse
        // than one missing number, which is the answer every other fit here
        // gives.
        const fits = H
          ? along >= textWidth(label, fs) + 2 && across >= fs * LABEL_FIT
          : along >= fs * LABEL_FIT && textWidth(label, fs) <= across;
        if (fits) {
          nodes.push({
            kind: "text",
            x: r.x - 4,
            y: r.y + r.h / 2 - fs * 0.75,
            w: r.w + 8,
            h: fs * 1.5,
            text: label,
            fontSize: fs,
            // Against the painted fill, not the original series colour: a hollow
            // PL/BU bar shows the canvas, and a PY/FC tint is far lighter than `fill`.
            color: contrastInk(seg.fill === "none" ? style.background : seg.fill),
            align: "center",
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
      }
    });
    seriesLevels.push(levels);

    const topV = pct
      ? denominators[c] > 0
        ? posTotals[c] / denominators[c]
        : 0
      : stacked
        ? Math.max(...ups)
        : clusteredTopValue(data.series.map((s) => s.values[c] ?? 0));
    const topQ = qOf(Math.max(0, topV));
    columnTop.push(H ? frame.x + topQ : frame.y + frame.h - topQ);
    drawnTopValue.push(topV);

    // Clustered-stacked: one total per stack sub-column (vertical only).
    if (decor.totals && !pct && nStacks > 1 && !H) {
      stackIds.forEach((id, sp) => {
        /**
         * THE SAME GATE THE BRANCH BELOW HAS, AND THIS ONE WENT WITHOUT IT.
         *
         * The `else if` below reads `columnHasData(data.series, c)`, and its
         * comment says it "gates BOTH branches below ... so a third one cannot
         * be added past it". This branch is not below it — it is ABOVE it, and
         * was written afterwards, so it slipped past the guard that comment
         * promised.
         *
         * Measured 2026-09-28, four series in two stacks with Q2 blank in all
         * of them: `total-1-s0 = "0"` and `total-1-s1 = "0"`, while the
         * single-stack control over the same blank column correctly printed no
         * `total-1` at all. It is the defect `totals.ts` already records — a sum
         * of nothing is 0, which is right for the stack's reach and wrong for a
         * label, because it tells the reader the business measured zero where
         * the truth is that nobody filled the cells in.
         *
         * Gated PER SUB-STACK, because that is the column this label sits over:
         * one stack can be blank while its neighbour is not, and only the blank
         * one should go unlabelled.
         */
        if (
          !columnHasData(
            data.series.filter((s) => (s.stack ?? 0) === id),
            c,
          )
        )
          return;
        const subX = centers[c] - colThick / 2 + sp * stackThick;
        const subTopQ = qOf(Math.max(0, ups[sp]));
        const total = data.series.reduce((a, s) => a + ((s.stack ?? 0) === id ? (s.values[c] ?? 0) : 0), 0);
        // Same band as the per-category total below, for the same reason: this
        // one sits over a sub-column, which reaches the top of the plot just as
        // readily.
        const subFs = aboveMarkFontSize(fs * 0.95, frame.y + frame.h - subTopQ, titleInkBottom(cfg, style), 1.53);
        // As a RATIO of the unshrunk size, so the box and the font move together
        // and a label that needed no shrinking is byte-identical. Written as
        // `subFs * 1.53` first, and 0.95 * 1.53 is 1.4535 against the 1.45 that
        // was there — four labels on a showcase slide moved three hundredths of
        // a point, which is the radar's tick lesson repeated one file over.
        const subScale = subFs / (fs * 0.95);
        if (subFs > 0)
          nodes.push({
            kind: "text",
            x: subX - 4,
            y: frame.y + frame.h - subTopQ - fs * 1.45 * subScale,
            w: stackThick + 8,
            h: fs * 1.4 * subScale,
            text: formatNumber(total, fmt),
            fontSize: subFs,
            bold: true,
            color: style.text,
            align: "center",
            valign: "bottom",
            name: `total-${c}-s${sp}`,
          });
      });
    } else if (decor.totals && !pct && columnHasData(data.series, c)) {
      // `columnHasData` gates BOTH branches below, because both print the same
      // number and a category with nothing in it has no total to print. It sums
      // to 0 — "missing cells count as 0", which is right for the stack's reach
      // and wrong for a label — and `total-3 = "0"` told the reader the business
      // measured zero where the truth was an empty column. Gated here rather
      // than inside each branch so a third one cannot be added past it.
      if (H) {
        // A total is centred on its row in a box `fs * 1.5` tall, so once the
        // font outgrows the row pitch the totals overlap each OTHER down the
        // end of the bars — the same defect the butterfly's value labels had,
        // and it wants the same bound: the ROW the total belongs to. Last
        // resort, so at any font that already fits its row this is `fs`.
        const totalFs = bandFontSize(fs, slotLen, 1.5);
        if (totalFs > 0)
          nodes.push({
            kind: "text",
            x: frame.x + topQ + 3,
            y: centers[c] - totalFs * 0.75,
            w: cfg.width - (frame.x + topQ) - 3,
            h: totalFs * 1.5,
            text: formatNumber(signedTotals[c], fmt),
            fontSize: totalFs,
            bold: true,
            color: style.text,
            align: "left",
            valign: "middle",
            name: `total-${c}`,
          });
      } else {
        // Upright, a total is centred on its column SLOT and was drawn at the
        // chart font however narrow that slot is — so on a small frame a
        // three-digit total ran into its neighbour's. `"113"` is 17.4pt at the
        // default font in a 13pt slot at 80x60, and the totals are the one
        // family here that never scaled while the title, the category names and
        // the series labels all did.
        //
        // Fit to the slot, then drop past the floor: the same two-step the
        // clustered in-bar labels take, and the horizontal branch above already
        // takes against its row pitch. Last resort — a total that already fits
        // keeps `fs` exactly, so no ordinary chart moves.
        const text = formatNumber(signedTotals[c], fmt);
        // Two bounds, in order: the SLOT sideways, and the band between the
        // title and the column's own top upright. The second used to be a clamp
        // to y=0, which is where the TITLE is — so a total that could not fit
        // over its column was drawn across the chart's own title instead of off
        // the top of it. Both defects, and the clamp traded one for the other.
        let tf = aboveMarkFontSize(fs, frame.y + frame.h - topQ, titleInkBottom(cfg, style), 1.45);
        while (tf > MIN_LABEL_FS && textWidth(text, tf, true) > slotLen) tf -= 0.5;
        // The box and the font move TOGETHER: `tf === fs` reproduces the old
        // geometry byte for byte, and a smaller label needs less clearance over
        // the bar it names.
        if (tf >= MIN_LABEL_FS && textWidth(text, tf, true) <= slotLen)
          nodes.push({
            kind: "text",
            x: centers[c] - slotLen / 2,
            // `aboveMarkFontSize` has already guaranteed this lands under the
            // title and inside the chart, so there is no clamp here any more.
            // The total sits `tf * 1.45` over its column's top, and where that
            // band cannot carry a readable number the total is not drawn — an
            // overlapping label does not "still read" when what it overlaps is
            // the title.
            y: frame.y + frame.h - topQ - tf * 1.45,
            w: slotLen,
            h: tf * 1.4,
            text,
            fontSize: tf,
            bold: true,
            color: style.text,
            align: "center",
            valign: "bottom",
            name: `total-${c}`,
          });
      }
    }
  }

  // Grand total (think-cell 14): one label at the top-right showing the sum of
  // every category total. A FIXED anchor in the de-collision pass, so a tall
  // right-hand column's own (movable) total nudges around it. Vertical only, and
  // never on a 100% chart, where every column totals the same 100%.
  // Drawn only where the band above the plot is ON the canvas. `frame.y` is the
  // title plus the totals row, and this sits `fs * 1.5` above it with no floor —
  // so at a 32pt font on a short frame the whole label was drawn above y=0,
  // 35.8pt past the top of an 80x60 chart. Clamping it down instead would put it
  // in the title, which is the move this engine already measured and reverted
  // for the CAGR caption: a fixed anchor that lands on other text is worse than
  // one that is not drawn. It is a summary of numbers the chart already shows.
  const grandY = frame.y - fs * 1.5;
  if (decor.grandTotal && !pct && !H && n > 0 && grandY >= 0) {
    const grand = signedTotals.reduce((a, b) => a + b, 0);
    const gtext = formatNumber(grand, fmt);
    const gw = Math.min(frame.w, textWidth(gtext, fs, true) + 8);
    nodes.push({
      kind: "text",
      x: frame.x + frame.w - gw,
      y: grandY,
      w: gw,
      h: fs * 1.4,
      text: gtext,
      fontSize: fs,
      bold: true,
      color: style.text,
      align: "right",
      valign: "bottom",
      name: "grand-total",
    });
  }

  // IBCS variance tier: a strip below the columns showing an actual series'
  // deviation from a reference (plan / previous year) as signed bars from a zero
  // line. Vertical only; the band height was reserved in computeFrame.
  if (decor.variance && !H && vframe && vframe.res.varianceH > 0) {
    const { actual, reference, mode = "absolute", goodIsUp = true } = decor.variance;
    const av = data.series[actual]?.values ?? [];
    const rv = data.series[reference]?.values ?? [];
    const deltas = data.categories.map((_, c) => {
      const a = av[c];
      const r = rv[c];
      if (a == null || r == null) return null;
      return mode === "percent" ? (r !== 0 ? ((a - r) / Math.abs(r)) * 100 : null) : a - r;
    });
    const maxAbs = maxOf(
      deltas.filter((d): d is number => d != null).map((d) => Math.abs(d)),
      0,
    );
    const bandTop = frame.y + frame.h;
    const bandH = vframe.res.varianceH;
    const zeroY = bandTop + bandH * 0.5;
    const halfH = bandH * 0.24; // leaves room for the delta label at the bar tip
    const scale = maxAbs > 0 ? halfH / maxAbs : 0;
    nodes.push({
      kind: "line",
      x1: frame.x,
      y1: zeroY,
      x2: frame.x + frame.w,
      y2: zeroY,
      stroke: style.axis,
      strokeWidth: 1,
      name: "variance-zero",
    });
    const GOOD = "#0ca30c";
    deltas.forEach((d, c) => {
      if (d == null) return;
      const favorable = goodIsUp ? d >= 0 : d <= 0;
      const color = d === 0 ? style.mutedText : favorable ? GOOD : style.negative;
      const barLen = Math.abs(d) * scale;
      const barW = Math.min(colThick, slotLen * 0.5);
      const up = d >= 0;
      const barY = up ? zeroY - barLen : zeroY;
      nodes.push({
        kind: "rect",
        x: centers[c] - barW / 2,
        y: barY,
        w: barW,
        h: barLen,
        fill: color,
        name: `variance-bar-${c}`,
      });
      const text = (d >= 0 ? "+" : "") + (mode === "percent" ? `${Math.round(d)}%` : formatNumber(d, fmt));
      nodes.push({
        kind: "text",
        x: centers[c] - slotLen / 2,
        y: up ? barY - fs * 1.05 : barY + barLen + 1,
        w: slotLen,
        h: fs * 1.05,
        text,
        fontSize: fs * 0.85,
        bold: true,
        color,
        align: "center",
        valign: up ? "bottom" : "top",
        name: `variance-label-${c}`,
      });
    });
  }

  // Connector lines between adjacent stacked columns: one per segment
  // boundary, so the development of each segment is easy to follow.
  if (decor.connectors && stacked && nStacks === 1) {
    const edge = (c: number, q: number, side: 1 | -1) =>
      H
        ? { x: frame.x + q, y: centers[c] + (side * colThick) / 2 }
        : { x: centers[c] + (side * colThick) / 2, y: frame.y + frame.h - q };
    for (let c = 0; c < n - 1; c++) {
      for (const bounds of [posBounds, negBounds]) {
        const a = bounds[c] ?? [];
        const b = bounds[c + 1] ?? [];
        // Sparse by series: only join a boundary that exists on both columns.
        for (let i = 0; i < Math.max(a.length, b.length); i++) {
          if (a[i] == null || b[i] == null) continue;
          const p1 = edge(c, qOf(a[i]), 1);
          const p2 = edge(c + 1, qOf(b[i]), -1);
          nodes.push({
            kind: "line",
            x1: p1.x,
            y1: p1.y,
            x2: p2.x,
            y2: p2.y,
            stroke: style.mutedText,
            strokeWidth: 0.75,
            name: `connector-${c}-${i}${bounds === negBounds ? "n" : ""}`,
          });
        }
      }
    }
  }

  if (!H) nodes.push(...breakMarkerNodes(frame, scale, style, cfg.width));

  // Zero baseline: horizontal line (vertical charts) or vertical line (bars).
  if (H) {
    nodes.push({
      kind: "line",
      x1: y0,
      y1: frame.y,
      x2: y0,
      y2: frame.y + frame.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
  } else {
    nodes.push(baselineNode(frame, y0, style));
  }

  if (decor.seriesLabels && !H) {
    nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid as (number | null)[]));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => colThick),
      columnTop,
      // Stacked keeps the signed total, which IS the drawn column top there.
      // Clustered takes the value `columnTop` was placed at — the tallest bar —
      // because the sum of a clustered category is drawn nowhere on the chart.
      // Two consumers put ink at whatever this says: the Error-row whiskers
      // (`base = a.columnValue[c]`) and `decorations.valueLines: [{mode:"mean"}]`.
      // With the sum, a 3-category 2-series chart anchored its whiskers at
      // y = -3, -80 and -63 on a 300pt canvas, and its mean line at y = -141
      // labelled "Ø 46" against a value axis topping out near 30 — real shapes
      // on the slide ABOVE the chart, since the Office renderer applies no clamp.
      // `decorations.difference`/`cagr` were hit more quietly: their Y came from
      // `columnTop` and their NUMBER from here, so the arrow sat over a +25% rise
      // and read "+29%".
      columnValue: pct ? posTotals : stacked ? signedTotals : drawnTopValue,
      // WHICH OF THOSE ARE MEASUREMENTS. All three arrays above are built with
      // `?? 0`, so a category nobody filled in is a zero here, indistinguishable
      // from a measured one — and the mean value line averaged it in, and
      // counted it in the denominator. See LayoutAnchors. The column family is
      // the one that offers `valueLines`, so it is the one that has to answer.
      columnHasData: data.categories.map((_, c) => columnHasData(data.series, c)),
      // What the right-margin names were ANCHORED to, before the spread moved
      // them. The combo re-lays this gutter with its own line names added, and
      // it needs where the columns started, not where they settled.
      seriesLabelAnchors: lastSegMid as (number | null)[],
      seriesLevels,
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: pct || H ? undefined : scale.toY,
      valueToX: H && !pct ? (v: number) => frame.x + qOf(v) : undefined,
    },
  };
}

/**
 * Combo chart, think-cell style: stacked columns plus line series drawn over
 * them on the same value axis. Series with `type: "line"` become lines; if
 * none is marked, the last series does.
 */
export function layoutCombo(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  /** Drawn over the columns rather than as one: a line, or bare markers. */
  const isOverlay = (s: (typeof cfg.data.series)[number]) => s.type === "line" || s.type === "marker";
  const marked = cfg.data.series.some(isOverlay);
  const nSeries = cfg.data.series.length;
  // Unmarked combo: the last series is the line, the rest are columns. A lone
  // series is a plain column, not both a column *and* a line (which would
  // double-render it).
  const lines = marked ? cfg.data.series.filter(isOverlay) : nSeries > 1 ? cfg.data.series.slice(-1) : [];
  const cols = marked
    ? cfg.data.series.filter((s) => !isOverlay(s))
    : nSeries > 1
      ? cfg.data.series.slice(0, nSeries - 1)
      : cfg.data.series;

  // Column mode: stacked (default), clustered, 100%, waterfall, or mekko.
  const columnsKind = cfg.combo?.columns ?? "stacked";
  // Independent axes: each line series gets its own scale (labelled, no shared
  // secondary-axis ticks) — for dashboards mixing unlike units.
  const independent = cfg.combo?.lineAxes === "independent" && lines.length >= 1;
  // Mekko (% normalised) and 100% columns expose no value→y map, so a line
  // needs its own axis; waterfall and column bases carry a shared scale.
  const noPrimaryAxis = columnsKind === "stacked100" || columnsKind === "mekko";
  // One shared scale: whichever of column extent / line values reaches higher.
  const stackMax =
    columnsKind === "clustered"
      ? maxOf(
          cols.flatMap((s) => s.values.filter((v): v is number => v != null)),
          0,
        )
      : maxOf(
          cfg.data.categories.map((_, c) => cols.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0)),
          0,
        );
  const lineMax = maxOf(
    lines.flatMap((s) => s.values.filter((v): v is number => v != null)),
    0,
  );
  // A shared-axis line can also dip BELOW the column base — e.g. a negative
  // overlay over all-positive or all-zero bars. The column scale floors at its
  // own data only, so without this the line plots far off the bottom of the plot
  // (a fuzz-found overshoot). Mirror the max-overflow fix: drop the shared floor
  // to the line, but only when it actually reaches lower than the columns would,
  // so combos whose bars already run more negative are left untouched.
  const lineMin = minOf(
    lines.flatMap((s) => s.values.filter((v): v is number => v != null)),
    0,
  );
  const colMin =
    columnsKind === "clustered"
      ? minOf(
          cols.flatMap((s) => s.values.filter((v): v is number => v != null)),
          0,
        )
      : minOf(
          cfg.data.categories.map((_, c) => cols.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0)),
          0,
        );
  const sharedFloor = lineMin < colMin ? niceTicks(lineMin, Math.max(stackMax, lineMax, 1))[0] : undefined;
  // Waterfall columns reach their running cumulative total, not the per-category
  // positive sum, so `stackMax` understates them; track the cumulative peak so a
  // shared-axis line taller than it isn't clipped off the top of the plot.
  const wf = (() => {
    const totals = new Set(cfg.waterfall?.totalIndices ?? []);
    const spacers = new Set(cfg.waterfall?.spacerIndices ?? []);
    const detailOf = detailParents(cfg);
    let running = 0;
    let max = 0;
    let min = 0;
    cfg.data.categories.forEach((_, c) => {
      // Match waterfallChain exactly: totals redraw the running level (no advance),
      // spacers carry it across unchanged, and detail-group columns decompose their
      // parent IN PLACE — none advance the chain. Counting spacer/detail columns
      // here overstated the peak, so the line-overflow stretch never fired and a
      // taller shared-axis line clipped off the top. Every column series
      // contributes to a real step (layoutWaterfall stacks them all).
      if (!totals.has(c) && !spacers.has(c) && !detailOf.has(c)) {
        running += cols.reduce((a, s) => a + (s.values[c] ?? 0), 0);
      }
      max = Math.max(max, running);
      min = Math.min(min, running);
    });
    return { max, min };
  })();
  // Secondary axis: line series get their own right-hand scale. A 100% / mekko
  // base forces it. Independent axes replace the shared secondary axis.
  const secondary = (!!cfg.secondaryAxis || noPrimaryAxis) && !independent;
  // A shared-axis line that overflows the waterfall's cumulative peak needs the
  // column scale stretched to fit it (only then — otherwise leave the
  // waterfall's own auto scale untouched to preserve existing layouts).
  const waterfallLineOverflow =
    columnsKind === "waterfall" && !secondary && !independent && cfg.scale?.max == null && lineMax > wf.max;
  // The mirror case #157 missed: a shared-axis line dipping BELOW the waterfall's
  // running trough (or below 0) needs the floor stretched down, or its low points
  // plot off the bottom of the plot.
  const waterfallLineUnderflow =
    columnsKind === "waterfall" && !secondary && !independent && cfg.scale?.min == null && lineMin < wf.min;
  // Overflow keeps its exact prior form (snapshot-neutral); underflow only adds a
  // min. When both fire, one niceTicks over the combined range keeps ticks aligned.
  const wfScale =
    waterfallLineOverflow && waterfallLineUnderflow
      ? (() => {
          const t = niceTicks(Math.min(lineMin, wf.min), Math.max(lineMax, wf.max, 1));
          return { ...cfg.scale, min: t[0], max: t[t.length - 1] };
        })()
      : waterfallLineOverflow
        ? { ...cfg.scale, max: niceTicks(0, Math.max(lineMax, wf.max, 1)).pop() }
        : waterfallLineUnderflow
          ? { ...cfg.scale, min: niceTicks(Math.min(lineMin, wf.min), Math.max(lineMax, wf.max, 1))[0] }
          : cfg.scale;
  const colCfg: ChartConfig = {
    ...cfg,
    kind: columnsKind,
    data: { ...cfg.data, series: cols },
    scale:
      columnsKind === "stacked100"
        ? undefined
        : columnsKind === "mekko"
          ? cfg.scale
          : columnsKind === "waterfall"
            ? wfScale
            : cfg.scale?.max != null || secondary
              ? cfg.scale
              : {
                  ...cfg.scale,
                  min: cfg.scale?.min ?? sharedFloor,
                  max: niceTicks(0, Math.max(stackMax, lineMax, 1)).pop(),
                },
  };
  /**
   * The lines, declared to the legend the column layout is about to draw.
   *
   * `colCfg` has them removed — they are drawn here, afterwards — and the legend
   * is built from `colCfg`, so a combo's legend named its columns and never its
   * lines. At every size, not only crowded ones. Where the line's own end-of-line
   * name also could not be drawn (a 120x90 horizontal combo at 6pt), the series
   * had no name anywhere on the chart.
   *
   * Coloured the way the line itself is: `seriesColor(style, cols.length + li)`
   * is exactly what the overlay below paints with, so the chip is a miniature of
   * the line rather than a second opinion about it.
   */
  const decorWithLines: Decorations = {
    ...decor,
    legendAlso: decor.seriesLabels
      ? lines.map((s, li) => ({
          label: s.scenario ? `${s.name} (${s.scenario})` : s.name,
          color: seriesColor(style, cols.length + li, s.color),
        }))
      : undefined,
  };
  const result =
    columnsKind === "waterfall"
      ? layoutWaterfall(colCfg, style, decorWithLines)
      : columnsKind === "mekko"
        ? layoutMekko(colCfg, style, decorWithLines)
        : columnsKind === "area"
          ? layoutLine(colCfg, style, decorWithLines)
          : layoutColumns(colCfg, style, decorWithLines);
  const { anchors, nodes } = result;
  // Horizontal (bar) base: the value axis runs left-to-right, and `categoryX`
  // holds each category's Y centre. The overlay is the same line drawn against
  // the other axis — before this it was drawn with the category's Y in the X
  // slot, so a horizontal Pareto (`pareto: true` keeps `horizontal`) rendered
  // its bars correctly and its cumulative-% line at ninety degrees to them.
  const H = !!cfg.horizontal;
  // No value→q map on the base and no line axis at all → nothing to overlay.
  // Read the map that matches the ORIENTATION: `valueToY` is undefined on every
  // horizontal base, which used to drop each overlay series silently, while the
  // chart's accessible description still listed it.
  const baseMap = H ? anchors.valueToX : anchors.valueToY;
  if (!baseMap && !secondary && !independent) return result;

  const fs = style.fontSize;
  /**
   * The size a line's point labels are drawn at, when horizontal.
   *
   * Sideways the label is centred on its category's ROW, in a box `fs * 1.4`
   * tall, and was bounded by nothing — so once the font outgrew the row pitch
   * the labels of adjacent categories were drawn through each other, which was
   * the largest single group of text collisions left in the engine. Bound by
   * the row, the same way the totals and the category names beside them now
   * are. Upright a point label sits ABOVE its point rather than on a row, so
   * nothing there wants this and `fs` is returned unchanged.
   */
  const rowPitch =
    anchors.categoryX.length > 1 ? Math.abs(anchors.categoryX[1] - anchors.categoryX[0]) : anchors.plot.h;
  const pointFs = H ? bandFontSize(fs, rowPitch, 1.4) : fs;
  /** Value → coordinate along the value axis (x when horizontal, y when not). */
  let lineToY = baseMap ?? ((v: number) => (H ? anchors.plot.x + v : anchors.plot.y + anchors.plot.h - v));
  /** Point for a value at a category, in the base chart's orientation. */
  const pointAt = (q: number, c: number) => (H ? { x: q, y: anchors.categoryX[c] } : { x: anchors.categoryX[c], y: q });
  if (secondary) {
    // Span down to any negative line value, not just up from 0 — anchoring the
    // secondary axis at 0 mapped a negative point below the plot floor. For an
    // all-positive line min(0, lineMin) is 0, so the domain and mapping are
    // unchanged (byte-identical); only a negative overlay is rescued.
    const ticks2 = niceTicks(Math.min(0, lineMin), Math.max(1, lineMax), 5);
    const min2 = ticks2[0];
    const max2 = ticks2[ticks2.length - 1];
    const plot = anchors.plot;
    lineToY = H
      ? (v: number) => plot.x + ((v - min2) / (max2 - min2 || 1)) * plot.w
      : (v: number) => plot.y + plot.h - ((v - min2) / (max2 - min2 || 1)) * plot.h;
    const fmt2 = resolveAxisFormat(ticks2, cfg.numberFormat);
    /**
     * How much of its size this strip may keep, from the gap between its own
     * ticks. Zero drops the numbers and leaves the axis to its gridlines.
     *
     * IT HAD NO FIT AT ALL, which made it the largest single shape left in this
     * project's overlap sweep — 617 pairs of `secondary-axis` on
     * `secondary-axis`, plus 320 more onto the category names it was drawn
     * across. Every combo, pareto and dual-axis column draws this strip, and the
     * scatter a few hundred lines away had the very fit it needed.
     *
     * The two orientations need DIFFERENT dimensions, which is the whole reason
     * `tickGapScale` takes a `want`. Sideways the labels sit at one y and spread
     * along x, so what one needs is its WIDTH; upright they stack down the right
     * edge, so it is a line height. Measuring the wrong one is how a scatter's x
     * axis came to be drawn through itself.
     */
    const tickScale2 = tickGapScale(
      fs * 0.9,
      ticks2,
      lineToY,
      H ? plot.w : plot.h,
      H ? (t) => textWidth(formatNumber(t, fmt2), fs * 0.9) + 2 : () => fs * 1.4,
    );
    /**
     * One shift for the WHOLE sideways strip, rather than a clamp per label.
     *
     * The clamp below keeps a label on the canvas, and sideways that is not a
     * free move: the labels are spaced along x, so pulling the last one left
     * closes the gap the fit above had just opened. A horizontal pareto at 80x60
     * had its ticks fitted to a 19.8pt gap and then squeezed into 14.8, and the
     * numbers touched anyway — the fit was right and the clamp undid it.
     *
     * Moving all of them by the same amount keeps every gap exactly as measured.
     * Where the strip cannot be brought onto the canvas whole — it is wider than
     * the chart — the numbers are dropped rather than crushed, which is the
     * answer this strip already gives when its ticks will not fit.
     *
     * Upright there is nothing to do: every label shares one x, so clamping it
     * moves the strip as a block already.
     */
    const boxW = fs * 3.4 * tickScale2;
    /**
     * THE STRIP YIELDS TO WHATEVER IS ALREADY THERE.
     *
     * No gutter is reserved for it — the note at the push below says so — and
     * that is the whole of this family of collisions. Sideways it sits above the
     * plot at `plot.y - fs * 1.5`, clamped by `Math.max(0, …)`, so a chart whose
     * chrome has squeezed the plot to the ceiling pins the whole strip onto the
     * TITLE. Upright it sits two points right of the plot, where a narrow chart
     * has the category names, the primary value axis, the legend and the series
     * names. 184 pairs across the variant sweep, spread thin over every kind
     * that can carry an overlay.
     *
     * Reserving a gutter was the other option and it is the worse one: it would
     * shrink the plot on every dual-axis chart, including the roomy ones that
     * have no collision to fix.
     *
     * So the secondary strip is the one that gives way. It is the add-on — the
     * base chart's own axis, its category names and its legend were placed
     * first and describe the bars — and this is the same answer its own note
     * already gives for the title: a tick that cannot be labelled inside the
     * chart keeps its gridline and loses its number.
     *
     * Tested against `nodes` as it stands right now, which is exactly the base
     * chart plus the title. Anything drawn AFTER this point is the overlay's
     * own, and those are fitted against each other by the rules above.
     */
    const placedText = nodes.filter((n): n is TextNode => n.kind === "text" && !!String(n.text).trim()).map(tightBox);
    const landsOnSomething = (x: number, y: number, w: number, h: number) =>
      placedText.some((b) => x < b.x + b.w && b.x < x + w && y < b.y + b.h && b.y < y + h);
    const stripShift = (() => {
      if (!H || tickScale2 <= 0) return 0;
      const lefts = ticks2.map((t) => lineToY(t) - fs * 1.7 * tickScale2);
      const under = Math.min(...lefts);
      const over = Math.max(...lefts) + boxW - cfg.width;
      if (under < 0 && over > 0) return NaN; // wider than the canvas — see below
      return under < 0 ? -under : over > 0 ? -over : 0;
    })();
    for (const t of tickScale2 > 0 && !Number.isNaN(stripShift) ? ticks2 : []) {
      // The base chart's own axis sits below (horizontal) or left (vertical), so
      // the secondary strip goes on the opposite side: above the plot for bars,
      // right of it for columns. Drawn at plot right on a bar chart it ran off
      // the canvas edge. On a bar chart it is also pinned to the canvas top when
      // the chrome left no room — an untitled one plots from y≈0, and
      // plot.y − 1.5em put the whole strip off the canvas.
      const q = lineToY(t);
      // NOT ON THE TITLE, and not off the top of the chart. Upright this strip's
      // labels are centred on their own tick, so the topmost one is drawn half a
      // box above the plot's ceiling — and on a chart whose chrome has squeezed
      // the plot to the frame's edge that is 7.1pt above the canvas at an 18pt
      // font. The same per-tick answer the primary value axis already gives: a
      // tick that cannot be labelled inside the chart keeps its gridline and
      // loses its number.
      if (!H && printsOnTitle(cfg, style, q - fs * 0.7)) continue;
      // And not on anything the base chart already drew — see `landsOnSomething`.
      const tx = H
        ? q - fs * 1.7 * tickScale2 + stripShift
        : Math.max(0, Math.min(plot.x + plot.w + 2, cfg.width - boxW));
      const ty = H ? Math.max(0, plot.y - fs * 1.5 * tickScale2) : q - fs * 0.7 * tickScale2;
      if (landsOnSomething(tx, ty, boxW, fs * 1.4 * tickScale2)) continue;
      nodes.push({
        kind: "text",
        // No gutter is reserved for this strip, so on a column chart whose plot
        // runs to the canvas edge it started 2pt from it and every tick label
        // was drawn off the chart — where the frame-clip then removed it and the
        // axis had no labels at all. Overlapping the plot's last few points is
        // the cheaper failure.
        //
        // Sideways the correction is `stripShift`, applied to every label alike;
        // clamping them one at a time is what closed the gaps. Upright they all
        // share one x, so the clamp there already moves the strip as a block.
        x: tx,
        y: ty,
        w: boxW,
        h: fs * 1.4 * tickScale2,
        text: formatNumber(t, fmt2),
        fontSize: fs * 0.9 * tickScale2,
        color: style.mutedText,
        align: H ? "center" : "left",
        valign: "middle",
        name: "secondary-axis",
      });
    }
  }
  const fmt = resolveFormat(
    lines.flatMap((s) => s.values.filter((v): v is number => v != null)),
    cfg.numberFormat,
  );
  /**
   * ONE PRECISION PER SERIES WHEN THE SERIES HAVE THEIR OWN AXES.
   *
   * `fmt` above is resolved over every line series at once, which is right for a
   * shared axis: the points are read against one scale and one tick strip, so
   * they must round the same way. It is wrong for `combo.lineAxes:
   * "independent"`, where each line is zoomed to its own range and NO numeric
   * axis is drawn — the block below says so itself: "the point labels carry the
   * real values since there is no shared numeric axis to read".
   *
   * A union format resolves from the LARGEST series, so the small one lost every
   * digit it had. Measured on `Revenue [1200…1500]` beside `Margin
   * [0.31…0.45]` — exactly the pairing independent axes exist for — the margin
   * line drew perfectly and all four of its labels read "0". The same margin
   * series alone reads 0.31 / 0.34 / 0.38 / 0.45.
   *
   * Per-series only in the independent branch, so a shared-axis combo keeps the
   * one format its tick strip is labelled in and nothing that renders today
   * moves. An explicit `numberFormat.decimals` still wins either way —
   * `resolveFormat` returns it untouched.
   */
  const lineFmts = lines.map((s) =>
    independent
      ? resolveFormat(
          s.values.filter((v): v is number => v != null),
          cfg.numberFormat,
        )
      : fmt,
  );
  /**
   * Each line's value→coordinate map, hoisted out of the draw loop.
   *
   * It was computed inside `lines.forEach`, which was fine while nothing else
   * needed it. The sideways row fit below has to know where EVERY series lands
   * at a category before any of them is drawn, and re-deriving the mapping there
   * would be the same arithmetic written twice — two copies that agree until one
   * is edited, which is how a fit comes to measure a label the chart does not
   * draw.
   */
  const toYs = lines.map((s) => {
    // Independent axis: zoom this line to its own value range (a nice-ticked
    // [min,max]) so its shape is visible whatever its units; the point labels
    // carry the real values since there is no shared numeric axis to read.
    const nums = s.values.filter((v): v is number => v != null);
    const ownTicks = independent && nums.length ? niceTicks(Math.min(...nums), Math.max(...nums)) : [0, 1];
    const lo = ownTicks[0];
    const hi = ownTicks[ownTicks.length - 1];
    return independent
      ? H
        ? (v: number) => anchors.plot.x + ((v - lo) / (hi - lo || 1)) * anchors.plot.w
        : (v: number) => anchors.plot.y + anchors.plot.h - ((v - lo) / (hi - lo || 1)) * anchors.plot.h
      : lineToY;
  });
  const labelOnAll = decor.segmentLabels || independent;

  /**
   * Sideways, how big each point label may be given the OTHER SERIES' labels on
   * the same category row. Zero means it is not drawn.
   *
   * `pointFs` bounds these by the row pitch, which is the right bound for one
   * label per row and no bound at all for ten. A horizontal combo with ten line
   * series draws ten numbers along one row, positioned by value, and they ran
   * through each other — the entire remaining `combo-label` residue in the
   * variant sweep, 97 pairs on a single chart shape.
   *
   * Upright this does not arise: there each series' label hangs above its own
   * mark and the crowding is between CATEGORIES, which `neighbourFs` handles.
   *
   * Shrink, then drop, and nothing moves — as everywhere else. Two series whose
   * values nearly coincide cannot both be labelled at any size, and that is the
   * honest answer: their marks are on top of each other too, so a number beside
   * them names neither.
   */
  const rowFs: number[][] = (() => {
    const blank = lines.map((s) => s.values.map(() => pointFs));
    if (!labelOnAll || pointFs <= 0 || lines.length < 2) return blank;
    const out = lines.map((s) => s.values.map(() => 0));
    for (let c = 0; c < anchors.categoryX.length; c++) {
      // Every series that labels this category, ordered along the axis their
      // labels are spread over: x sideways (they sit on one row), y upright
      // (they sit in one column, above their own marks).
      const here = lines
        .map((s, li) => ({ li, v: s.values[c] }))
        .filter((e): e is { li: number; v: number } => e.v != null)
        .map((e) => ({ li: e.li, q: toYs[e.li](e.v), em: textWidth(formatNumber(e.v, lineFmts[e.li]), 1) }))
        .sort((a, b) => a.q - b.q);
      for (let i = 0; i < here.length; i++) {
        const me = here[i];
        // SIDEWAYS: only the neighbour to the RIGHT, and only my own width.
        // Every label starts AT its mark and runs rightward, so what must fit
        // between two marks is the LEFT label's width — mine. Constraining a
        // label by its left neighbour's width was the first version: it shrinks
        // the wrong label, since the left one is the one in the way.
        //
        // UPRIGHT: the labels share their category's x and are stacked by value,
        // so what must fit between two marks is a LINE HEIGHT. `q` ascends down
        // the canvas here, so the neighbour that matters is the next one down.
        // Anchored on the MARKS rather than the drawn label tops, which each sit
        // 1.65 of their own size above their mark — a bound that depended on the
        // size it is computing could not be solved in one pass, and the mark gap
        // is the same number for equal sizes and conservative for unequal ones.
        const next = here[i + 1];
        const room = next ? Math.abs(next.q - me.q) : Infinity;
        out[me.li][c] = next ? bandFontSize(pointFs, room, H ? me.em : 1.4) : pointFs;
      }
    }
    return out;
  })();

  /** Upright line names, collected for one shared spread down the gutter. */
  const gutterExtras: GutterName[] = [];
  lines.forEach((s, li) => {
    const color = seriesColor(style, cols.length + li, s.color);
    const toY = toYs[li];
    const labelOn = labelOnAll;
    // A marker series is this same overlay minus the connecting segments: the
    // values are per-category facts (a benchmark, a target, a consensus), and a
    // line between them would claim they interpolate. The mark is a little
    // larger, since it has no line to carry it.
    const markersOnly = s.type === "marker";

    /**
     * How big each of this line's point labels may be, given the ones beside it.
     * Zero means there is no room and the label is not drawn, which is the same
     * "no room" answer `pointFs` and `uprightPointFs` already give.
     *
     * WHY THIS EXISTS. Upright, a point label is centred on its mark and its only
     * fits were VERTICAL — the band between the title and the mark — plus a
     * horizontal clamp to the plot's edge. Neither can see the label one category
     * across. With four or six categories that never showed; at twenty-four it is
     * 306 of this project's remaining text overlaps, every one of them a number
     * drawn through the number beside it.
     *
     * The room a label really has is the CATEGORY PITCH, shared with whichever
     * neighbour is closer, and the pair clears when their half-widths together
     * fit the gap between their marks. Expressed through `bandFontSize` so the
     * legibility floor is the one every other fit in this engine uses.
     *
     * Sideways this stays untouched: there the labels sit BESIDE their marks, one
     * per row, and `pointFs` is already fitted to the row pitch. The crowding is
     * a property of the direction the categories run in, so the fix belongs on
     * one axis only.
     *
     * SHRINK, THEN DROP, and nothing moves — the same order the pie's and the
     * radar's names take. `resolveLabelCollisions` would be the other candidate
     * and it is the wrong one twice over: it moves labels VERTICALLY, which does
     * not answer a horizontal collision, and a point label nudged off its mark
     * names the wrong category.
     */
    const neighbourFs: number[] = (() => {
      if (H) return s.values.map(() => pointFs);
      const drawn = s.values.map((v, c) => v != null && c < anchors.categoryX.length);
      // Width at one point, so the pair constraint solves for size directly:
      // two centred labels clear when their half-widths together fit the gap,
      // i.e. size <= 2 * gap / (emWidth(a) + emWidth(b)).
      const em = s.values.map((v, c) => (drawn[c] ? textWidth(formatNumber(v as number, lineFmts[li]), 1) : 0));
      return s.values.map((_v, c) => {
        if (!drawn[c]) return 0;
        let worst = Infinity;
        for (const j of [prevDrawn(drawn, c), nextDrawn(drawn, c)]) {
          if (j < 0) continue;
          const gap = Math.abs(anchors.categoryX[j] - anchors.categoryX[c]);
          worst = Math.min(worst, bandFontSize(pointFs, 2 * gap, em[c] + em[j]));
        }
        // A lone point has no neighbour to yield to.
        return worst === Infinity ? pointFs : worst;
      });
    })();

    let prev: { x: number; y: number } | null = null;
    let last: { x: number; y: number } | null = null;
    s.values.forEach((v, c) => {
      if (v == null || c >= anchors.categoryX.length) {
        prev = null;
        return;
      }
      const pt = pointAt(toY(v), c);
      if (prev && !markersOnly)
        nodes.push({
          kind: "line",
          x1: prev.x,
          y1: prev.y,
          x2: pt.x,
          y2: pt.y,
          stroke: color,
          strokeWidth: 2,
          name: `combo-line-${li}-${c}`,
        });
      // Shrunk where the frame's edge is closer than the marker's own half
      // extent, and dropped below a point of it — see `markInFrame`.
      const r = markInFrame(cfg, pt.x, pt.y, markersOnly ? 3.2 : 2.4);
      if (r > 0)
        nodes.push({
          kind: "rect",
          x: pt.x - r,
          y: pt.y - r,
          w: r * 2,
          h: r * 2,
          fill: color,
          stroke: style.background,
          strokeWidth: 1,
          name: `combo-marker-${li}-${c}`,
        });
      // `pointFs` is 0 when the row is too thin to carry a legible label, which
      // is the same "no room" answer the label placers give elsewhere — drawn
      // anyway it would be a zero-size font, which OOXML rejects.
      // The room a beside-the-mark label actually has, capped at the 60pt it
      // always used, and measured to the PLOT's right edge rather than the
      // canvas. Measuring to `cfg.width` merely moved the collision: the label
      // came off the category names on the left and landed on the totals strip
      // on the right, which `computeFrameHorizontal` reserves outside the plot.
      // A clamp moves a label whether or not the destination is free — the
      // lesson the CAGR caption already cost this repo — so the bound is the
      // room inside the plot, and a label that cannot fit there is not drawn.
      const comboLabelW = Math.max(0, Math.min(60, anchors.plot.x + anchors.plot.w - (pt.x + r + 2)));
      // Upright the label hangs above the mark, so the room it has is whatever
      // is over the mark — `pt.y * (1 / 1.3)` is where its top ink lands.
      // Upright the label hangs above its mark AND sits between the marks either
      // side of it, so it has to clear both — the smaller of the two bounds is
      // the one it can actually take.
      const uprightPointFs = Math.min(
        aboveMarkFontSize(fs, pt.y, titleInkBottom(cfg, style), 1.3),
        neighbourFs[c],
        rowFs[li][c],
      );
      if (
        labelOn &&
        pointFs > 0 &&
        (H
          ? rowFs[li][c] >= MIN_LABEL_FS && comboLabelW >= textWidth(formatNumber(v, lineFmts[li]), rowFs[li][c])
          : uprightPointFs >= MIN_LABEL_FS)
      ) {
        // Categories run down a bar chart, so a label ABOVE its point would sit
        // on the neighbouring category's row; put it beside the mark instead.
        nodes.push(
          H
            ? {
                kind: "text",
                // Bounded by the room to the RIGHT of the mark, not a fixed 60pt
                // box clamped to `cfg.width - 60`. On a 60pt-wide chart that
                // clamp evaluates to x=0 and the label spans the whole frame,
                // landing on the category names down the left — the label placed
                // beside its mark to avoid the neighbouring ROW, put on the
                // neighbouring COLUMN instead.
                x: Math.min(pt.x + r + 2, Math.max(0, cfg.width - comboLabelW)),
                y: pt.y - rowFs[li][c] * 0.7,
                w: comboLabelW,
                h: rowFs[li][c] * 1.4,
                text: formatNumber(v, lineFmts[li]),
                fontSize: rowFs[li][c],
                color: independent ? color : style.text,
                align: "left",
                valign: "middle",
                name: `combo-label-${li}-${c}`,
              }
            : {
                kind: "text",
                x: pt.x - 30,
                // Upright, the label sits ABOVE its mark and the only thing over
                // a mark near the top of the plot is the edge of the chart: at
                // 32pt it was drawn 36pt above the top of an 80x60 chart. Fitted
                // to the band the mark actually leaves, and not drawn at all
                // below the legibility floor — `uprightPointFs` is 0 there, which
                // is the same "no room" answer the row fit gives sideways.
                y: pt.y - uprightPointFs * 1.65,
                w: 60,
                h: uprightPointFs * 1.4,
                text: formatNumber(v, lineFmts[li]),
                fontSize: uprightPointFs,
                color: independent ? color : style.text,
                align: "center",
                valign: "bottom",
                name: `combo-label-${li}-${c}`,
              },
        );
      }
      prev = pt;
      last = pt;
    });
    if (decor.seriesLabels && last != null) {
      const end: { x: number; y: number } = last;
      // The combo line's own series label, with the gutter fit `seriesLabelNodes`
      // now does for the column one — this is the same node under a different
      // name, and it overflowed the same way.
      // Same missing gutter as the secondary axis above: a plot that reaches the
      // canvas edge left this strip ONE POINT wide, so the fit below shrank the
      // name to nothing. Take a readable strip at the edge instead — it names a
      // line whose end point is right there, so a little overlap costs less than
      // the name does.
      //
      // The floor is LOAD-BEARING on real configs, not just on thumbnails: the
      // showcase's combo slide has a plot that reaches the right edge, and
      // removing it dropped "Margin %" from that deck outright. Measure this on
      // more than `sampleConfig("combo")` before touching it.
      const gutter = cfg.width - (anchors.plot.x + anchors.plot.w) - 4;
      // Use the REAL gutter when it is wide enough to hold a readable name, and
      // fall back to the overhanging strip only when it is not. The floor was
      // being taken even where a perfectly good gutter existed, so this label
      // stayed at the chart font and was drawn across the column series labels
      // that share the same margin — two names for two different series, on top
      // of each other.
      const lw = H ? 80 : gutter >= MIN_LABEL_FS * 3 ? gutter : Math.max(fs * 3.4, gutter);
      const lx = H
        ? Math.max(0, Math.min(end.x - 40, cfg.width - 80))
        : Math.max(0, Math.min(anchors.plot.x + anchors.plot.w + 4, cfg.width - lw));
      // Sideways this name sits ON the plot beside the line's last point, and it
      // was the one label there still drawn at the full chart font: at 18pt on a
      // 120x90 bar chart every number around it had been fitted to the row and
      // shrunk to 6.7pt while "Margin %" stayed at 18 and was drawn across four
      // of them. It takes the same row bound as the point labels it sits among —
      // `pointFs` — and is not drawn at all when the row cannot carry it.
      // Upright it keeps `fs` exactly, so nothing there moves.
      // Half the row, not the whole of it: the point labels own their rows, and
      // this name sits BETWEEN two of them (`horizontalNameY` puts it under the
      // line's last point, which is another category's row). At the full row
      // bound it still met the neighbours' numbers on a 120x90 bar chart; at
      // half it clears them, and where half a row is unreadable the name is not
      // drawn — a line whose name cannot be read is still the line the legend
      // names, and `collide.ts` has nowhere to move this one sideways: every row
      // above it is occupied.
      // Sideways: the CLEAR AIR between two rows' own labels, not a share of the
      // pitch. Every row carries a category name and a total, both centred on
      // it, so what the line's name actually has is the pitch less the height of
      // one of those labels — and where that is not a readable size, the name is
      // not drawn. At 32pt on a 200x150 chart the rows are 16pt apart and their
      // numbers are 12pt of that, which is the case this is measuring: there is
      // no position for it, so no size for it either.
      const nameFs = H ? bandFontSize(fs, Math.max(0, rowPitch - pointFs * 1.4), 1.05) : fs;
      let lf = nameFs;
      while (lf > 5 && textWidth(s.name, lf) > lw) lf -= 0.5;
      const nameY = H
        ? horizontalNameY(end.y, nameFs, cfg.height, rowPitch)
        : uprightNameY(end.y, fs, cfg.height, titleInkBottom(cfg, style));
      /**
       * WHERE THE LEGEND ALREADY NAMES THIS SERIES, the name at the end of the
       * line is said twice and yields to whatever it collides with.
       *
       * TRIED ONCE BEFORE AND REVERTED, on the same reasoning, because the
       * premise was false: a combo's legend did not contain its lines at all, so
       * the rule dropped a line's only identification. Fixing the legend to name
       * every series it draws is what makes this available — and the ordering
       * matters, because a rule that is right only after another fix looks
       * identical to a rule that is wrong.
       *
       * BY NAME, never by "a legend exists" — and that distinction is EQUIVALENT
       * TODAY, which is worth writing down rather than deleting. `legendRow`
       * wraps rather than truncating, and `horizontalLegendFits` refuses the
       * whole legend or none of it, so a drawn legend currently holds every
       * entry and the two checks agree. A mutant reducing this to "a legend
       * exists" passes the suite.
       *
       * It stays by name because it is the honest statement of the rule, and
       * because the first attempt at this rule was reverted for precisely the
       * case the loose check cannot see: a legend that names some series and not
       * others. Nothing produces one now; the day something does, this is what
       * keeps a line from losing its only name.
       *
       * Upright there is no legend at all — `decor.seriesLabels` draws the
       * column series' names down the right margin instead — so nothing is said
       * twice and every name stays. Yield what is said twice, keep what is said
       * once.
       */
      /**
       * UPRIGHT, THIS NAME IS NOT DRAWN HERE AT ALL — it is handed to the gutter
       * layout below and spread with the column names, in one pass over every
       * name in the strip. See `gutterExtras`.
       *
       * Sideways it stays: there the name sits ON the plot beside the line's last
       * point, in a row of its own, and shares nothing with the column names.
       */
      if (!H) {
        gutterExtras.push({
          name: s.scenario ? `${s.name} (${s.scenario})` : s.name,
          color,
          y: end.y,
          nodeName: `combo-series-label-${li}`,
        });
        return;
      }
      const legendNamesIt = nodes.some(
        (n) => n.kind === "text" && /^legend-\d+$/.test(n.name ?? "") && String((n as TextNode).text ?? "") === s.name,
      );
      const nameBox = { x: lx, y: nameY ?? 0, w: lw, h: nameFs * 1.4 };
      const saidTwiceAndInTheWay =
        legendNamesIt &&
        nodes.some((n) => {
          if (n.kind !== "text") return false;
          const t = n as TextNode;
          if (!String(t.text ?? "").trim() || /^legend-/.test(t.name ?? "")) return false;
          const b = tightBox(t);
          return (
            nameBox.x < b.x + b.w && b.x < nameBox.x + nameBox.w && nameBox.y < b.y + b.h && b.y < nameBox.y + nameBox.h
          );
        });
      if (nameFs > 0 && nameY != null && !saidTwiceAndInTheWay)
        nodes.push({
          kind: "text",
          x: lx,
          // Sideways the name sits BELOW the line's last point, and the floor at 0
          // was the only bound it had: `end.y` is wherever that point landed, so
          // on a short frame the whole label was drawn past the foot of the chart
          // — y 91.5 with a 14pt box on a 90pt canvas, and the same at 300x60 and
          // 200x150, all at the DEFAULT font. Off the chart is onto whatever sits
          // under it on the slide, since neither PowerPoint renderer clips.
          // Bounded at both ends now; `combo-series-label-` is movable, so
          // de-collision can still lift it off whatever it lands on.
          y: nameY,
          w: lw,
          h: nameFs * 1.4,
          text: clipToWidth(s.name, lf, lw),
          fontSize: lf,
          color: style.text,
          align: H ? "center" : "left",
          valign: "middle",
          name: `combo-series-label-${li}`,
        });
    }
  });
  /**
   * ONE SPREAD OVER THE WHOLE GUTTER, replacing the column-only one.
   *
   * `layoutColumns` has already spread its own names down this strip and emitted
   * them, knowing nothing of the lines. Those nodes are removed and the strip is
   * laid out again with every name in it — columns from the anchors the base
   * exposes, lines from the points collected above. The result is a pitch that
   * accounts for all of them, which is what stops the global de-collision pass
   * having anything to do, which was the whole intent of the original spread.
   *
   * Only when there is something to add. With no upright lines this is skipped
   * entirely and the base's own nodes stand, so an ordinary column chart is
   * untouched — and so is a horizontal combo, whose line names sit on the plot.
   */
  // NOT WHEN THE SECONDARY AXIS IS IN THIS GUTTER TOO. The spread fills the
  // strip top to bottom, and the secondary tick numbers occupy the same two
  // points of x — so on a dual-axis or pareto chart the merged spread crosses
  // the whole strip and the yield below then deletes it. Measured: 34 overlaps
  // fixed, 335 tick numbers lost, and a comfortable 480x300 pareto left with
  // two. That is the wrong way round, so those charts keep the old arrangement
  // and their 30 pairs; the gutter there holds three families and wants a
  // design, not a third pass.
  if (gutterExtras.length && result.anchors.seriesLabelAnchors && !secondary) {
    for (let i = nodes.length - 1; i >= 0; i--) if (/^series-label-\d+$/.test(nodes[i].name ?? "")) nodes.splice(i, 1);
    const gutter = seriesLabelNodes(
      colCfg,
      style,
      result.anchors.plot,
      result.anchors.seriesLabelAnchors,
      gutterExtras,
    );
    nodes.push(...gutter);
    /**
     * AND THE SECONDARY STRIP YIELDS AGAIN, because it shares this gutter and
     * its own rule ran too early to see these names.
     *
     * That strip already gives way to whatever was drawn before it — the rule is
     * a few hundred lines up — but it is emitted BEFORE the lines exist, so it
     * could only yield to the column names. Moving the line names into the same
     * strip put fourteen new pairs of tick number on line name: the shape of
     * trade this file has already made twice, a family fixed and a neighbour
     * grown.
     *
     * Same verdict as before. A tick that cannot be labelled inside the chart
     * keeps its gridline and loses its number; a line name upright has no legend
     * behind it and is the only thing naming that series.
     */
    const gutterBoxes = gutter.map((n) => tightBox(n as TextNode));
    for (let i = nodes.length - 1; i >= 0; i--) {
      const n = nodes[i];
      if (n.name !== "secondary-axis" || n.kind !== "text") continue;
      const b = tightBox(n as TextNode);
      if (gutterBoxes.some((g) => b.x < g.x + g.w && g.x < b.x + b.w && b.y < g.y + g.h && g.y < b.y + b.h))
        nodes.splice(i, 1);
    }
  }
  return result;
}

/** Chrome for horizontal (bar) orientation: title, legend, left category labels, bottom axis. */
export function horizontalChrome(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  frame: Frame,
  centers: number[],
  scale: ValueScale,
  qOf: (v: number) => number,
): SceneNode[] {
  const nodes: SceneNode[] = chromeNodes(
    cfg,
    style,
    { ...decor, categoryAxis: false, valueAxis: false, gridlines: false },
    frame,
    centers,
  );
  const fs = style.fontSize;
  if (decor.gridlines) {
    for (const t of scale.ticks) {
      if (t === 0) continue;
      const x = frame.x + qOf(t);
      nodes.push({
        kind: "line",
        x1: x,
        y1: frame.y,
        x2: x,
        y2: frame.y + frame.h,
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: "gridline",
      });
    }
  }
  if (decor.valueAxis) {
    // Same labeller as the vertical chrome: a bar chart's value axis is the same
    // tick strip rotated, so it earns the same step-derived precision and the
    // same unitless share branch (see axisTickLabel).
    const axisLabel = axisTickLabel(scale.ticks, scale.percent, cfg.numberFormat);
    // Rotating the axis rotates its crowding: the vertical strip's labels run
    // into each other when the font outgrows the tick SPACING, and this one's
    // when it outgrows the tick spacing measured the other way — as width. The
    // box is a fixed 48pt regardless, so nothing here was ever bound by the gap
    // it sits in. Bound by the gap, shrunk together so the axis reads at one
    // size, then clipped. At any font whose labels already fit between two
    // ticks this is `fs * 0.9` and nothing moves.
    const tickGap =
      scale.ticks.length > 1
        ? Math.min(...scale.ticks.slice(1).map((t, i) => Math.abs(qOf(t) - qOf(scale.ticks[i]))))
        : frame.w;
    const room = Math.max(1, tickGap - 2);
    const tickFs = (() => {
      let f = fs * 0.9;
      while (f > MIN_LABEL_FS && scale.ticks.some((t) => textWidth(axisLabel(t), f) > room)) f -= 0.5;
      return f;
    })();
    // The loop floors rather than shrinking to nothing, so a gap this tight
    // leaves labels that still overlap — and a label too small to read is not
    // worth the collision. Dropped whole, gridlines kept, as on the other axes.
    const tickFits = scale.ticks.every((t) => textWidth(axisLabel(t), tickFs) <= room);
    // …and only where the strip is not inside the TITLE. Sideways it hangs off
    // the foot of the plot, and `fitPlot` grows a squeezed plot up from that
    // foot, so on a 300x60 chart at 24pt the whole row of numbers was drawn
    // across the chart's own name.
    const stripClear = !printsOnTitle(cfg, style, frame.y + frame.h + 2);
    for (const t of tickFits && stripClear ? scale.ticks : []) {
      const x = frame.x + qOf(t);
      nodes.push({
        kind: "text",
        x: x - 24,
        y: frame.y + frame.h + 2,
        w: 48,
        h: fs * 1.4,
        text: clipToWidth(axisLabel(t), tickFs, room),
        fontSize: tickFs,
        color: style.mutedText,
        align: "center",
        valign: "top",
        name: "value-axis",
      });
    }
  }
  if (decor.categoryAxis) {
    /**
     * One size for the whole axis, bounded by BOTH the row it names and the
     * gutter it sits in.
     *
     * The upright category axis was fitted to its slot width and this one — the
     * same axis rotated — was fitted to nothing at all: drawn at `fs` in a box
     * `fs * 1.5` tall, centred on its row. So once the font outgrew the row
     * pitch every name overlapped its neighbours (nine horizontal kinds at 18pt
     * on a 200x150 frame), and a name wider than the gutter ran back across the
     * bars it was naming, since `computeFrameHorizontal` caps that gutter at
     * 30% of the width however long the names are.
     *
     * Height first, then width, then clip — the two bounds are independent and
     * the smaller wins. Last resort as everywhere else: at any size where the
     * names already fit their row and their gutter this is `fs` and nothing
     * moves.
     */
    const slotH = centers.length > 1 ? Math.abs(centers[1] - centers[0]) : frame.h;
    const gutter = Math.max(1, frame.x - 4);
    const catFs = (() => {
      let f = bandFontSize(fs, slotH, 1.5);
      while (f > MIN_LABEL_FS && cfg.data.categories.some((c) => textWidth(String(c ?? ""), f) > gutter)) f -= 0.5;
      return f;
    })();
    // Zero means the rows cannot carry a legible name at any size, so the axis
    // is dropped whole rather than drawn at a size OOXML will not accept.
    if (catFs > 0)
      cfg.data.categories.forEach((cat, i) => {
        nodes.push({
          kind: "text",
          x: 0,
          y: centers[i] - catFs * 0.75,
          w: gutter,
          h: catFs * 1.5,
          text: clipToWidth(String(cat ?? ""), catFs, gutter),
          fontSize: catFs,
          color: style.text,
          align: "right",
          valign: "middle",
          name: `category-${i}`,
        });
      });
  }
  // Gated by the SAME predicate the reservation uses — see `horizontalLegendFits`.
  // Gating one without the other is the bug this repo has now written twice.
  // `legendLabelsFor`, not `cfg.data.series.length`, because a combo's lines are
  // not in this config at all — see `Decorations.legendAlso`. A one-column combo
  // with one line has a single series HERE and two on the chart, and the old
  // count sent it away with no legend for either.
  if (decor.seriesLabels && legendLabelsFor(cfg, decor).length > 1 && horizontalLegendFits(cfg, style, decor)) {
    nodes.push(
      ...legendRow(cfg, style, frame.x, (cfg.title ? fs * 1.6 + 6 : 0) + 2, {
        maxX: cfg.width - 4,
        entries: legendEntriesFor(cfg, style, decor),
      }),
    );
  }
  return nodes;
}

/** Horizontal legend row: color chip + series name, left to right. */
/** One legend entry: a coloured chip and a label. */
export interface LegendEntry {
  label: string;
  /** Chip fill, as the mark it labels is painted ("none" for a hollow bar). */
  color: string;
  /** Chip outline / hatch, so the chip is a miniature of that mark. */
  stroke?: string;
  strokeWidth?: number;
  pattern?: Series["pattern"];
  /** Node name for the text (defaults to `legend-${index}`). */
  name?: string;
}

/**
 * Horizontal legend of coloured chips, wrapping to new rows so a chart with many
 * series/groups never marches its chips off the right edge (`opts.maxX`). Custom
 * entries (group names, a "Peer range" swatch) come via `opts.entries`; without
 * them it legends the multi-series set from `seriesLegendLabels`. Called with no
 * opts it is byte-identical to the old single-row version (maxX defaults to no
 * wrap).
 */
/**
 * The legend's entries: this config's own series, then anything the layout has
 * added through `decor.legendAlso`.
 *
 * Shares `legendLabelsFor`'s list so the row drawn here and the band reserved
 * for it cannot disagree about how many entries there are.
 */
export function legendEntriesFor(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LegendEntry[] {
  const own: LegendEntry[] = seriesLegendLabels(cfg).map((label, si) => {
    const paint = markPaint(style, seriesColor(style, si, cfg.data.series[si].color), cfg.data.series[si]);
    return { label, color: paint.fill, stroke: paint.stroke, strokeWidth: paint.strokeWidth, pattern: paint.pattern };
  });
  // `seriesLegendLabels` answers EMPTY for a single series, which is right when
  // that is the whole chart and wrong when a combo has taken the others out —
  // one column and one line is a two-series chart, and both want naming.
  const base =
    own.length === 0 && (decor.legendAlso?.length ?? 0) > 0
      ? cfg.data.series.map((s, si) => {
          const paint = markPaint(style, seriesColor(style, si, s.color), s);
          return {
            label: s.scenario ? `${s.name} (${s.scenario})` : s.name,
            color: paint.fill,
            stroke: paint.stroke,
            strokeWidth: paint.strokeWidth,
            pattern: paint.pattern,
          };
        })
      : own;
  return [...base, ...(decor.legendAlso ?? [])].map((e, i) => ({ ...e, name: `legend-${i}` }));
}

export function legendRow(
  cfg: ChartConfig,
  style: ChartStyle,
  x0: number,
  y: number,
  opts: { maxX?: number; entries?: LegendEntry[] } = {},
): SceneNode[] {
  const fs = style.fontSize;
  const nodes: SceneNode[] = [];
  const chip = fs * 0.7;
  const rowH = fs * 1.6;
  const maxX = opts.maxX ?? Infinity;
  const entries: LegendEntry[] =
    opts.entries ??
    seriesLegendLabels(cfg).map((label, si) => {
      // Key the chip with the paint the segments actually get, not the raw
      // series colour — pattern and IBCS scenario are half the encoding.
      const paint = markPaint(style, seriesColor(style, si, cfg.data.series[si].color), cfg.data.series[si]);
      return {
        label,
        color: paint.fill,
        stroke: paint.stroke,
        strokeWidth: paint.strokeWidth,
        pattern: paint.pattern,
        name: `legend-${si}`,
      };
    });
  // Shared wrap walk (frame.ts) so the row count here matches what the frame
  // reserved via legendRowCount.
  const slots = legendWrapWalk(
    entries.map((e) => e.label),
    fs,
    x0,
    maxX,
  );
  entries.forEach((e, si) => {
    const wLabel = textWidth(e.label, fs);
    const { x, row } = slots[si];
    const ry = y + row * rowH;
    nodes.push(
      {
        kind: "rect",
        x,
        y: ry + fs * 0.35,
        w: chip,
        h: chip,
        fill: e.color,
        stroke: e.stroke,
        strokeWidth: e.strokeWidth,
        pattern: e.pattern,
        name: `legend-chip-${si}`,
      },
      {
        kind: "text",
        x: x + chip + 3,
        y: ry,
        w: wLabel + 6,
        h: fs * 1.4,
        text: e.label,
        fontSize: fs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: e.name ?? `legend-${si}`,
      },
    );
  });
  return nodes;
}

/**
 * Right-hand series labels at the last column's segment midpoints,
 * greedily pushed apart so they never overlap (think-cell placement).
 */
/**
 * Where a horizontal combo's line name sits relative to the line's last point.
 *
 * Below it by preference, which is where it has always gone — but the floor at
 * 0 was the only bound it had, and `end.y` is wherever that point landed. On a
 * short frame the whole label was therefore drawn past the foot of the chart:
 * y 91.5 with a 14pt box on a 90pt canvas, the same at 300x60 and 200x150, all
 * at the DEFAULT font, and off the chart is onto whatever sits under it on the
 * slide since neither PowerPoint renderer clips.
 *
 * ABOVE the point when below does not fit, rather than clamped into the canvas:
 * clamping was tried first and put the name straight onto the line's own last
 * POINT LABEL on four frames — a trade the frame gate caught, and one nothing
 * downstream can undo now that a point label may not travel more than two em
 * from its mark. Above and below a point are equally true of it, which is the
 * argument `collide.ts` already makes for flipping a point label.
 *
 * The clamp stays as the last resort for a frame too short for either.
 *
 * AIMED AT THE GUTTER BETWEEN ROWS when the row pitch is known, which is the
 * correction 24 and 32pt forced. "Just below the point" is `fs * 0.9` away, and
 * sideways every row carries a category name on the left and a total on the
 * right, both CENTRED on the row — so a name a few points under its point is
 * inside the next row's numbers however small it is shrunk. The clear air is
 * the boundary between two rows, and that is a position rather than a size.
 */
export function horizontalNameY(pointY: number, fs: number, canvasH: number, rowPitch = 0): number {
  const boxH = fs * 1.4;
  // Half a row down lands on the boundary; without a pitch this is the old
  // `fs * 0.9`, so a caller that does not know its rows behaves as before.
  const below = pointY + (rowPitch > 0 ? rowPitch / 2 - boxH / 2 : fs * 0.9);
  if (below >= 0 && below + boxH <= canvasH) return below;
  const above = pointY - (rowPitch > 0 ? rowPitch / 2 + boxH / 2 : fs * 0.9 + boxH);
  if (above >= 0) return above;
  return Math.max(0, Math.min(Math.max(0, below), canvasH - boxH));
}

/**
 * Where an UPRIGHT combo's line name sits relative to the line's last point.
 *
 * Above it by preference, which is where it has always gone — and `end.y` is
 * wherever that point landed, so a line ending near the top of the plot put the
 * name outside the chart: 13pt above a 300x60 frame at 24pt. Below the point
 * when above does not fit, for the reason `horizontalNameY` gives for flipping
 * the other way, and clamped only as the last resort on a frame too short for
 * either.
 */
export function uprightNameY(pointY: number, fs: number, canvasH: number, titleInk = 0): number | null {
  const boxH = fs * 1.4;
  const above = pointY - fs * 1.6;
  // Clear of the TITLE, not merely of the canvas. A line ending near the top of
  // a squeezed plot has the title above it and nothing else, and "on the
  // canvas" put the name across it — the same trade `aboveMarkFontSize` refuses
  // for every value drawn over a mark.
  if (above >= titleInk) return above;
  // Below the point, and never higher than the title's own ink — a point near
  // the top of a squeezed plot is itself inside the title's band, so "just
  // below it" is too.
  const below = Math.max(pointY + fs * 0.2, titleInk);
  if (below + boxH <= canvasH) return below;
  // NEITHER position is available, so there is nowhere to put it. Null rather
  // than the old clamp: the clamp's destination on a frame this short is the
  // title, and a name that reads as part of the chart's title names nothing.
  return null;
}

/** One name in the right-hand gutter, before it is spread. */
export interface GutterName {
  name: string;
  color: string;
  y: number | null;
  /** Node name to emit. Omitted for a config's own series — see below. */
  nodeName?: string;
}

export function seriesLabelNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  frame: { x: number; y: number; w: number; h: number },
  midYs: (number | null)[],
  /**
   * Names that belong in this gutter but are NOT in `cfg.data.series`.
   *
   * A combo's lines are the whole of it. They are drawn after the columns and
   * share this strip of canvas, and until 2026-08-28 they were placed by their
   * own code afterwards — so this function spread the columns to fill a gutter
   * it thought it had to itself, and the global de-collision pass then inherited
   * ten labels in a strip sized for seven with only "up" available. It pushed
   * four of them 24 points and closed one pitch to 2.8, which is an overlap this
   * function had deliberately prevented.
   *
   * Passing them here is the fix: one spread over every name in the gutter, so
   * the pitch is honest and de-collision has nothing to do — which is what the
   * comments below always claimed.
   */
  extra: GutterName[] = [],
): SceneNode[] {
  const fs = style.fontSize;
  const lineH = fs * 1.35;
  /** Font the labels are actually drawn at — reduced only if they must be spread. */
  let labelFs = fs;
  /**
   * NAMED BEFORE SORTING, and that is deliberate. A config's own series are
   * emitted as `series-label-<n>` where n is their position DOWN THE GUTTER, not
   * their index in the datasheet — which is what this function has always done,
   * so numbering them here keeps every existing name identical. The extras carry
   * their own names, so interleaving them cannot renumber the columns.
   */
  const own: GutterName[] = cfg.data.series
    .map((s, i) => ({
      name: s.scenario ? `${s.name} (${s.scenario})` : s.name,
      color: seriesColor(style, i, s.color),
      y: midYs[i],
    }))
    .filter((e) => e.y != null)
    .sort((a, b) => (a.y as number) - (b.y as number))
    .map((e, i) => ({ ...e, nodeName: `series-label-${i}` }));
  const entries = [...own, ...extra]
    .filter((e): e is GutterName & { y: number } => e.y != null)
    .sort((a, b) => a.y - b.y);
  // Push overlapping labels apart, then clamp back into the frame.
  for (let i = 1; i < entries.length; i++) {
    if (entries[i].y - entries[i - 1].y < lineH) entries[i].y = entries[i - 1].y + lineH;
  }
  const overflow = entries.length ? entries[entries.length - 1].y + lineH / 2 - (frame.y + frame.h) : 0;
  if (overflow > 0) {
    for (const e of entries) e.y -= overflow;
    for (let i = entries.length - 2; i >= 0; i--) {
      if (entries[i + 1].y - entries[i].y < lineH) entries[i].y = entries[i + 1].y - lineH;
    }
    // Only the BOTTOM was clamped, so the upward propagation walked straight
    // past the canvas top: twelve series on a 240×160 chart put two labels at
    // negative y, and thirty on a 400×300 put the topmost at −123. `collide.ts`
    // refuses to nudge a label off the top for the reason that applies here too
    // — an overlapping label still reads, an off-canvas one is lost — but it
    // only moves labels UP, so it cannot rescue one already emitted above the
    // canvas. Spread evenly over the band when the gap cannot be honoured,
    // which is what `layoutSlope.place()` does after the same discovery.
    //
    // The floor is the TITLE's bottom, not the canvas top. These labels sit in
    // the right-hand gutter, so moving them down cannot land them on the plot —
    // but the title is a FULL-WIDTH band, and on six kinds the upward
    // propagation walked the topmost label into the middle of it, at a size
    // where a reader sees the two texts crossed rather than one clamped. It is
    // the same "cannot honour the gap" case the canvas-top branch already
    // handles, so it takes the same answer: spread over the band that IS
    // available, and shrink to it.
    const bottom = Math.max(lineH / 2, Math.min(frame.y + frame.h, cfg.height - lineH / 2));
    // Clamped at `bottom` so the band can never invert. With no title this is
    // `lineH / 2` and both the test and the spread are exactly as before.
    const top = Math.min(titleHeight(cfg, style) + lineH / 2, bottom);
    if (entries[0].y < top) {
      const step = entries.length > 1 ? (bottom - top) / (entries.length - 1) : 0;
      // The shrink below is floored at MIN_LABEL_FS, so a band too short to give
      // every label a line of its own does not get one by shrinking — it gets
      // labels pitched CLOSER than their own height, which is the overlap the
      // shrink exists to prevent. At the limit the band has no height at all
      // (`top === bottom`, a short frame whose title has eaten the gutter) and
      // every label is placed at the SAME POINT: three series names drawn on
      // top of one another, the reader seeing one of them and having no way to
      // know which line it names. `collide.ts` cannot rescue them either — its
      // nudge only goes up, all of them want the same place, and they exhaust
      // the budget still stacked.
      //
      // So past that floor the labels are DROPPED, which is the answer every
      // other reservation in this engine already gives when it cannot be paid
      // for (the radar's ticks, the sunburst's ring, the pie's outside labels):
      // a label that cannot be told apart from its neighbour is not a label,
      // and a chart visibly missing its series names is honest where a stack of
      // hidden ones is not. Only for two or more — a single label overlaps
      // nothing, so the pitch says nothing about it.
      if (entries.length > 1 && step < MIN_LABEL_FS * 1.25) return [];
      entries.forEach((e, i) => (e.y = top + i * step));
      // Spread AND shrunk, because a spread alone trades one defect for a worse
      // one. The step here is by definition below the gap the labels wanted, so
      // every neighbouring pair overlaps — and `series-label-` is in
      // `collide.ts`'s MOVABLE list, whose nudge only goes UP. It therefore
      // pushed each label past the one above it and returned them REORDERED:
      // every label on the canvas, each naming the wrong line. An unreadably
      // small label is a bad chart; a legible label naming someone else's line
      // is a wrong one. Shrink to the step so nothing overlaps and the
      // de-collision pass has nothing to do.
      labelFs = Math.max(MIN_LABEL_FS, Math.min(fs, step / 1.25));
    }
  }
  const x = frame.x + frame.w + 4;
  // The gutter these sit in is reserved from the series names, but the
  // reservation is capped — so on a chart too narrow to pay for it the names
  // ran off the RIGHT edge, by 20pt on a 120pt-wide stacked column. This
  // function already shrinks for vertical crowding a few lines up; it never
  // asked whether a name FITS the gutter it is being put in.
  //
  // Shrunk together and then clipped, for the reason the vertical case gives
  // and one more: labels that differ in size read as a hierarchy that is not
  // there. Taken as a minimum with the crowding font so neither fix can undo
  // the other, and last-resort in both directions, so a chart whose names
  // already fit is untouched.
  const room = Math.max(1, cfg.width - x);
  let nameFs = labelFs;
  while (nameFs > 5 && entries.some((e) => textWidth(e.name, nameFs) > room)) nameFs -= 0.5;
  return entries.map((e) => ({
    kind: "text" as const,
    x,
    y: e.y - lineH / 2,
    w: room,
    h: lineH,
    text: clipToWidth(e.name, nameFs, room),
    fontSize: nameFs,
    color: style.text,
    align: "left" as const,
    valign: "middle" as const,
    // Assigned above, before the sort, so an extra cannot renumber a column.
    name: e.nodeName ?? "series-label",
  }));
}
