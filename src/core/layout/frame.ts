import type { ChartConfig, ChartStyle, Decorations, NumberFormat } from "../types";
import { LINE_PITCH, extraLinesHeight, lineCount, maxLineCount, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { niceTicks, axisTickLabel, segmentLabel } from "../format";

export interface Frame {
  /** Plot rectangle in chart coordinates. */
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * The smallest plot a layout may be handed, in points. It exists to keep the
 * plot POSITIVE, not to make it readable — a chart this small is not readable
 * on any arithmetic.
 */
export const MIN_PLOT_SIDE = 8;

/**
 * The smallest any label may be drawn; below it the label is dropped.
 *
 * Every fit in this engine shrinks a label to the room it has, and none of them
 * had a floor — so a band squeezed to a few points answered with labels of one
 * or two, which is ink rather than text. The house floor for a shrink loop is 5
 * (`while (f > 5 …)` everywhere); this is the same number for the fits
 * expressed as arithmetic rather than as a loop.
 */
export const MIN_LABEL_FS = 5;

/**
 * The size a label centred in a band of `band` points may be drawn at, in a box
 * `band / boxRatio` tall — or ZERO, meaning there is no room and the caller
 * must not draw it.
 *
 * `Math.min(fs, band / ratio)` is the whole of "fit the label to the mark it
 * sits on", and written inline it has no floor: a band of zero returns a font
 * of zero. That is not a small label, it is an INVALID one — OOXML's `sz` is
 * hundredths of a point with a minimum of 100, so a chart whose data is all
 * zeroes (a template nobody has filled in yet, the commonest thing there is)
 * writes a deck PowerPoint offers to repair. Found by sweeping degenerate data
 * for non-positive font sizes the day the inline form was introduced.
 *
 * Returning 0 rather than clamping to the floor is deliberate and is the answer
 * the radar, sunburst, tilemap and pie reservations already give: a label too
 * small to read is not there, so do not spend the band drawing it. Callers
 * branch on it, which is also what stops the invalid size reaching a renderer.
 */
/**
 * How much of its size a strip of TICK LABELS may keep, given the gap between
 * its own ticks. `0` means there is no room and the strip is not drawn.
 *
 * `want` is what one label needs in the direction the ticks are spaced. Down a
 * vertical axis the labels stack, so that is a line height. Across a horizontal
 * one they sit side by side, so it is the WIDEST label's width — and using the
 * line height there fits the wrong dimension entirely: `1,234,567,890` is 60
 * points wide at the default font and clears a 14-point test with ease, so the
 * numbers were drawn straight through each other. Ordinary values hide it; the
 * defect arrives with the magnitude, not with the frame.
 *
 * SHARED, and it was not. This lived inside `layoutScatter` and the secondary
 * value axis — which every combo, pareto and dual-axis column draws — had no fit
 * of any kind. Its ticks overlapped each other in 617 pairs of the variant
 * sweep, the largest single shape left in it, on a strip a few points from the
 * one this helper was already protecting.
 *
 * `drawnFs` is the size the labels are actually drawn at, not the chart font:
 * both callers draw their ticks at `fs * 0.9`, and comparing the chart font to
 * the floor would keep a strip that is drawn below it.
 */
export function tickGapScale(
  drawnFs: number,
  vals: number[],
  to: (v: number) => number,
  span: number,
  want: (t: number) => number,
): number {
  const gap = vals.length > 1 ? Math.min(...vals.slice(1).map((t, i) => Math.abs(to(t) - to(vals[i])))) : span;
  const need = Math.max(1, ...vals.map(want));
  const scale = Math.min(1, gap / need);
  return drawnFs * scale < MIN_LABEL_FS ? 0 : scale;
}

export function bandFontSize(fs: number, band: number, boxRatio: number): number {
  const f = Math.min(fs, band / boxRatio);
  return Number.isFinite(f) && f >= MIN_LABEL_FS ? f : 0;
}

/**
 * The size a label drawn ABOVE a mark may take — the band between the bottom of
 * the title and the top of the mark it names, or ZERO when that band cannot
 * carry a readable one.
 *
 * Every kind that puts a number over a column, a box or a point had the same two
 * bounds missing, in the same order. The first is the frame: a label hung
 * `fs * 1.45` over a mark that reaches the top of the plot is drawn off the
 * chart, and clamping it to y=0 — which is what the upright column totals did —
 * only exchanges one defect for another, because y=0 is where the TITLE is. At
 * 18pt on a 300x60 chart that put a total, a waterfall value, a boxplot median
 * and a mekko total straight across the title of their own chart.
 *
 * So the room is measured from the title's own band, and where it cannot be met
 * the label is not drawn: the answer the radar, sunburst, tilemap and pie
 * reservations already give. A title is the one label on a chart that names what
 * the reader is looking at, and it is drawn first — a value written over it
 * costs both.
 *
 * `titleInk` is `titleInkBottom` — the title's ink, NOT the band reserved for
 * it, which is eight points lower at the default font and shrinks labels that
 * were never on the title. It is 0 on a chart with no title, so an untitled
 * chart keeps the whole space above its marks.
 */
export function aboveMarkFontSize(fs: number, markTop: number, titleInk: number, boxRatio: number): number {
  return bandFontSize(fs, markTop - titleInk, boxRatio);
}

/**
 * Clamp a plot rectangle into the chart frame.
 *
 * Every layout computes its plot by subtracting its chrome — title, legend,
 * axis, footnote — from the frame, and each of those is a fixed number of
 * points. On a frame too small to pay for them the subtraction goes NEGATIVE,
 * and a negative height is not merely a small one: `toY` maps the value domain
 * through it, so the axis INVERTS and the plot's own bottom edge lands below the
 * frame. Scatter at 120x90 — a thumbnail — computed `h: -8`, which put its y
 * tick labels 5pt past the bottom of the chart and mapped larger values
 * downward.
 *
 * The floor is what stops the inversion. Where the recovered height comes from
 * is the other half: the plot grows UP from the bottom edge the layout gave it,
 * never down. That edge is the category axis and the value baseline — moving it
 * moves what the chart claims — while everything above it is chrome, and chrome
 * is exactly what a frame this small cannot pay for. Anchoring the other way
 * round was tried and is worse: it pins the plot to the bottom of the frame and
 * every label drawn beneath a mark spills out of it (14pt on a 120x90 bubble).
 *
 * The plot then overlaps its own title, which is ugly and honest — the
 * alternative is a chart drawn upside down and spilling onto the slide.
 *
 * A plot that already fits is returned with every field bit-identical, so this
 * is unreachable for any chart at a sane size.
 */
/**
 * The half-extent a mark centred at (cx, cy) may be drawn at, or 0.
 *
 * A MARK IS DRAWN AROUND ITS POSITION, so a point at the very edge of a plot
 * puts half its marker outside — harmless while chrome sits between the plot
 * and the frame, and not harmless once `fitPlot` brings the two together. A
 * line's own markers were 1.6pt above an 80x60 chart with a footnote on it, and
 * a combo's 1.2pt to the left of a 60x300 one with an explicit scale.
 *
 * Scatter and bubble answer this by moving the MARK inside (`markerExtent`), and
 * that is not available here: these markers sit on a line drawn through the same
 * points, so moving one detaches it from its own series. The mark keeps its
 * position and gives up size instead — and below a point of it there is nothing
 * worth drawing, so it is dropped and the line goes on saying where the point
 * was.
 */
export function markInFrame(cfg: ChartConfig, cx: number, cy: number, r: number): number {
  const room = Math.min(cx, cy, cfg.width - cx, cfg.height - cy);
  const fit = Math.min(r, room);
  return fit >= 1 ? fit : 0;
}

/**
 * The text of a point label: the pane's "Label content" parts applied to one
 * value, as the upright line and the stacked column already apply them. A point
 * has no share of anything, so a `percent` part has nothing to print.
 */
export function pointLabelText(
  decor: Decorations,
  v: number,
  seriesName: string,
  category: string,
  fmt: Partial<NumberFormat>,
): string {
  return segmentLabel(decor.labelContent ?? ["value"], { value: v, fraction: null, series: seriesName, category, fmt });
}

/**
 * The box of a point label drawn BESIDE its point along x: to the right of it —
 * and to the LEFT when the right has not the room for the whole of it. `gap` is
 * the clearance from the point and `minW` the width the box has on the right
 * (the ink and a little, unless the caller says otherwise).
 *
 * It was always to the right, with no gutter reserved there
 * (`computeFrameHorizontal` reserves one only for totals). So the point at the
 * scale maximum — which is the plot's right edge, and is a round number whenever
 * the data's largest value is on a nice tick (30, 40, 100, 3000, 30000) — had 6pt
 * of canvas for its label, and `clipTextToFrame` dropped it outright. A value
 * just under the maximum fared worse: it was cut to a fragment that reads as
 * another number, 29,500 as "2…" and 0.97 as "0.…". Measured 2026-10-24 at
 * 480x300 on `[12, 30]`: the 30 was not drawn at all, on the sideways line chart
 * and on the sideways dot, lollipop and range plots alike.
 *
 * Flipped rather than reserved for, because a gutter is paid by every sideways
 * chart and the problem is the one point at the edge: the label of an upright
 * point is also kept on the canvas, and "an overlapping label still reads where
 * an off-canvas one is lost" is the rule `layoutLine` already applies. Only a
 * point near the right edge ever flips; when neither side has the room the right
 * is kept, and the clip does what it has always done.
 */
export function labelBesidePoint(
  ptX: number,
  text: string,
  fs: number,
  gap: number,
  canvasW: number,
  minW?: number,
): { x: number; w: number; align: "left" | "right" } {
  const ink = textWidth(text, fs);
  const right = canvasW - (ptX + gap);
  const onRight = minW ?? ink + 4;
  if (ink <= right || ptX - gap <= right) return { x: ptX + gap, w: onRight, align: "left" };
  const w = Math.max(onRight, ink + 4);
  return { x: ptX - gap - w, w, align: "right" };
}

export function fitPlot<T extends Frame>(cfg: ChartConfig, plot: T): T {
  const w = Math.max(MIN_PLOT_SIDE, Math.min(plot.w, cfg.width));
  const h = Math.max(MIN_PLOT_SIDE, Math.min(plot.h, cfg.height));
  const x = Math.max(0, Math.min(plot.x + plot.w - w, cfg.width - w));
  const y = Math.max(0, Math.min(plot.y + plot.h - h, cfg.height - h));
  if (x === plot.x && y === plot.y && w === plot.w && h === plot.h) return plot;
  return { ...plot, x, y, w, h };
}

export interface CategorySlots {
  /** Center x of each slot. */
  centers: number[];
  /** Width of each slot (before the gap is removed). */
  slotWidth: number;
  /** Width of the column/cluster within a slot. */
  colWidth: number;
}

/** think-cell default: column width : gap ≈ 2 : 1 within each slot. */
export function categorySlots(frame: Frame, n: number, gapRatio = 1 / 3): CategorySlots {
  const slotWidth = frame.w / Math.max(1, n);
  const colWidth = slotWidth * (1 - gapRatio);
  const centers = Array.from({ length: n }, (_, i) => frame.x + slotWidth * (i + 0.5));
  return { centers, slotWidth, colWidth };
}

/**
 * The room each category name has along its axis: the narrower of the two gaps
 * either side of its centre, so two neighbouring names that each take their own
 * room can never meet, and no more than reaches the nearer end of the plot
 * (`lo`..`hi`), so the box of a name at the end of the axis stays on the chart.
 * `whole` is the answer for a lone category.
 *
 * THE AXIS USED TO ASSUME EVERY SLOT IS THE SAME WIDTH and read the first gap,
 * `centers[1] - centers[0]`, as the width of all of them. That holds for every
 * chart but a date-spaced line, whose centres are proportional to time. Measured
 * 2026-10-24 at 960x540: a first gap of one day (`2024-01-01, 2024-01-02,
 * 2024-06-01, 2024-12-01`) gave every name 1pt and `clipTextToFrame` dropped the
 * whole axis, with 110pt free for each of the others; and `2023-01` followed by
 * twelve months of 2024 sized every name for the year-long first gap, so 11 of
 * the 12 adjacent pairs overlapped, on an axis `collide.ts` leaves fixed.
 *
 * Evenly spaced centres — every chart that is not date-spaced — answer with the
 * one width they always had, to the last bit: the gaps of a fixed pitch differ in
 * their last digits, and the minimum of them is not the first one.
 */
export function categoryRooms(centers: number[], whole: number, lo: number, hi: number): number[] {
  if (centers.length < 2) return centers.map(() => whole);
  const gaps = centers.slice(1).map((c, i) => Math.abs(c - centers[i]));
  if (gaps.every((g) => Math.abs(g - gaps[0]) < 1e-6)) return centers.map(() => gaps[0]);
  return centers.map((c, i) =>
    Math.min(i > 0 ? gaps[i - 1] : Infinity, i < gaps.length ? gaps[i] : Infinity, 2 * (c - lo), 2 * (hi - c)),
  );
}

export interface ValueScale {
  min: number;
  max: number;
  ticks: number[];
  toY: (v: number) => number;
  /** Present when an axis break compresses part of the range: the band's y extent. */
  breakBand?: { yLow: number; yHigh: number };
  /**
   * The axis measures SHARES (a 100% chart), so its ticks are fractions that
   * must be labelled as percentages. Without this the axis of a chart whose
   * segments read "60%" was itself labelled 0.00-1.00.
   */
  percent?: boolean;
}

/**
 * Linear scale mapping [min, max] onto the plot height, snapped to nice ticks.
 * `override` pins either end manually (think-cell's axis-handle dragging).
 */
/**
 * The floor a LOG axis should be given.
 *
 * Every caller derives `dataMin` with a zero seed (`minOf(all, 0)`) because a
 * linear value axis baselines at zero — which means valueScale's `dataMin > 0`
 * branch is unreachable from real code and EVERY log axis fell back to
 * `dataMax / 1000`: three decades below the top, whatever the data does. Values
 * of 200–300 got the axis 0.1…1,000 and drew inside the top 4% of the plot,
 * with two decades of gridlines under them that no datum can reach.
 *
 * Return the smallest POSITIVE value instead, or `dataMin` when there is none
 * (valueScale then declines the log branch anyway, since dataMax ≤ 0).
 */
export const logFloor = (values: readonly number[], dataMin: number): number => {
  const lo = values.reduce((m, v) => (v > 0 && v < m ? v : m), Infinity);
  return Number.isFinite(lo) ? lo : dataMin;
};

export function valueScale(
  frame: Frame,
  dataMin: number,
  dataMax: number,
  override?: { min?: number; max?: number },
  axisBreak?: { from: number; to: number },
  logScale?: boolean,
  // Column charts baseline at zero, so their axis must always include it. But a
  // distribution chart (boxplot/violin/candlestick) of e.g. scores 40–95 forcing
  // 0 into the range squashes the data into the top of the plot — those callers
  // pass zeroFloor:false to keep the domain data-driven. cfg.scale still overrides.
  zeroFloor = true,
): ValueScale {
  /**
   * DECLARED HERE, above both branches, because it was written for one of them
   * and the other returns before it.
   *
   * The full reasoning is at its old position below: a manual scale.max under
   * the data — which the pane invites, "Axis scale min / max" being free-text
   * boxes — drove `toY` clean off the canvas, and clamping is what every
   * charting tool does with a value above the axis maximum. The linear branch
   * got that fix. The LOG branch sits thirty lines earlier and returns its own
   * `toY`, so it never saw it: `logScale` with `scale.max: 100` over data
   * reaching 10,000 put the top bar at y = -540 on a 300pt canvas — 540 points
   * above the slide — while the same numbers on a linear axis were fine.
   *
   * One definition above both, so they cannot answer this differently again.
   */
  const clip = (y: number) => Math.max(frame.y, Math.min(frame.y + frame.h, y));
  // Logarithmic axis: decade ticks; requires positive data (falls back otherwise).
  if (logScale && dataMax > 0) {
    const minPos = Math.max(dataMin > 0 ? dataMin : dataMax / 1000, 1e-12);
    let lo10 = Math.floor(Math.log10(override?.min && override.min > 0 ? override.min : minPos));
    // Never let the top decade fall below the bottom one: a manual scale.min set
    // above the data would give an empty tick list → undefined min → NaN toY for
    // the whole axis. Clamp to at least one decade.
    const hi10 = Math.max(lo10, Math.ceil(Math.log10(override?.max && override.max > 0 ? override.max : dataMax)));
    // "At least one decade" is what the line above says and is not what it did:
    // `max(lo10, …)` clamps to at least ZERO decades. Data that is all one exact
    // power of ten — `logScale` with values [100, 100], or a single 1000 —
    // yields lo10 === hi10, one tick, and a `max - min || 1` span of 1 log unit.
    // Every value then maps to the axis floor: a chart whose every bar is
    // 0pt high, drawn under a one-tick axis. Give it the decade BELOW, because
    // the floor is where a bar starts, so lowering it is what makes the data
    // visible; raising the top would leave the bars on the baseline.
    if (hi10 === lo10) lo10 -= 1;
    const ticks: number[] = [];
    for (let e = lo10; e <= hi10; e++) ticks.push(Math.pow(10, e));
    const min = ticks[0];
    const max = ticks[ticks.length - 1];
    const span = Math.log10(max) - Math.log10(min) || 1;
    const toY = (v: number) =>
      clip(frame.y + frame.h - ((Math.log10(Math.max(v, min)) - Math.log10(min)) / span) * frame.h);
    return { min, max, ticks, toY };
  }
  const autoLo = zeroFloor ? Math.min(0, dataMin) : dataMin;
  const autoHi = zeroFloor ? Math.max(0, dataMax) : dataMax;
  const lo = override?.min ?? autoLo;
  const hi = override?.max ?? autoHi;
  let ticks = niceTicks(lo, hi, 5).filter(
    (t) => (override?.min == null || t >= override.min - 1e-9) && (override?.max == null || t <= override.max + 1e-9),
  );
  let min = override?.min ?? ticks[0];
  let max = override?.max ?? ticks[ticks.length - 1];
  // A manual scale the data cannot live in — scale.min at or above the auto max
  // (the filter then leaves a single tick, so max collapses onto min), or an
  // outright degenerate {min:0,max:0} — makes the `max - min || 1` fallback below
  // map ONE data unit to ONE point: a 300pt chart emitted its columns ~83 canvas
  // heights down the slide. normalizeConfig already repairs an inverted or
  // non-finite scale; repair an unusable one the same way, by falling back to the
  // range the data itself asks for.
  if (!(max > min)) {
    ticks = niceTicks(autoLo, autoHi, 5);
    min = ticks[0];
    max = ticks[ticks.length - 1];
  }
  // CLIPPED to the plot, which is what a manual scale narrower than the data
  // has to mean.
  //
  // The repair above only catches a scale that is fully unusable. One that is
  // merely far too narrow passed straight through and `toY` extrapolated with
  // nothing to stop it: pin the axis to 0-100 on revenue data — which the pane
  // invites, "Axis scale min / max" being free-text boxes — and the bars land at
  // y = -231558 on a 300pt canvas. The SVG preview hides it behind its viewBox,
  // so the first anyone sees of it is a .pptx whose group extent exceeds the
  // OOXML coordinate limit, which PowerPoint offers to repair.
  //
  // Clamping is not a compromise here, it is the behaviour: every charting tool
  // draws a value above the axis maximum as a bar reaching the top of the plot.
  // Anything inside the scale is untouched, so this changes only what was
  // previously drawn off the slide.
  // (declared at the top of this function, above the log branch that also needs it)
  let toY = (v: number) => clip(frame.y + frame.h - ((v - min) / (max - min || 1)) * frame.h);
  let breakBand: ValueScale["breakBand"];

  // think-cell axis break: the [from, to] range is compressed into a small
  // fixed band so out-of-scale columns fit the plot.
  if (axisBreak && axisBreak.from > min && axisBreak.to < max && axisBreak.to > axisBreak.from) {
    const { from, to } = axisBreak;
    const gapFrac = 0.06;
    const below = from - min;
    const above = max - to;
    const span = below + above || 1;
    const belowFrac = (below / span) * (1 - gapFrac);
    const aboveFrac = (above / span) * (1 - gapFrac);
    const frac = (v: number) =>
      v <= from
        ? (below ? (v - min) / below : 0) * belowFrac
        : v >= to
          ? belowFrac + gapFrac + (above ? (v - to) / above : 0) * aboveFrac
          : belowFrac + ((v - from) / (to - from)) * gapFrac;
    // Through the same `clip`. This branch REPLACES the clamped `toY` above, and
    // replacing it threw the clamp away: `frac` extrapolates freely outside
    // [min, max], and `min`/`max` are whatever the author pinned — so a manual
    // scale narrower than the data put a bar at y = -21026 with a height of
    // 21308 on a 300pt canvas, and the .pptx carried `<a:off y="-243779040"/>`.
    // That is the exact blow-up the clamp was added to prevent, reachable by
    // setting an axis break and an axis scale at the same time, both of them
    // free-text boxes in the pane.
    toY = (v: number) => clip(frame.y + frame.h - frac(v) * frame.h);
    ticks = ticks.filter((t) => t <= from + 1e-9 || t >= to - 1e-9);
    breakBand = { yLow: toY(from), yHigh: toY(to) };
  }
  return { min, max, ticks, toY, breakBand };
}

/** Slanted-band break marker drawn across the plot (over the columns). */
export function breakMarkerNodes(frame: Frame, scale: ValueScale, style: ChartStyle, canvasW?: number): SceneNode[] {
  if (!scale.breakBand) return [];
  const { yLow, yHigh } = scale.breakBand;
  const skew = 2.5;
  // The zigzag OVERHANGS the plot on purpose — that is the convention that says
  // the axis is cut — but the overhang is a fixed 4 points and the plot's own
  // left edge is only a few points in on a small chart, so the mark was drawn
  // off the side of the CHART: 2pt past the left edge of an 80x60 or 120x90
  // frame, on four kinds, at every font. Overhang what the frame can pay for.
  //
  // `canvasW` is optional so a caller that has no canvas to measure against
  // keeps the old right-hand overhang; both callers in this engine pass it.
  const left = Math.min(4, frame.x);
  const right = canvasW == null ? 4 : Math.min(4, Math.max(0, canvasW - (frame.x + frame.w)));
  return [
    {
      kind: "rect",
      x: frame.x - Math.min(2, left),
      y: yHigh,
      w: frame.w + Math.min(2, left) + Math.min(2, right),
      h: yLow - yHigh,
      // LEFT AS A KNOCK-OUT ON PURPOSE (the r2-35 sweep of `fill: style.background`).
      // The doughnut's disc and the forecast markers had a real hollow to draw
      // instead, `fill: "none"`; this one has not. The band is an OCCLUDER: it exists
      // to hide the part of every column, segment, connector and label that crosses
      // the cut, and none of those nodes knows about the cut — an empty fill would
      // draw the bars straight through it. A real gap means splitting every rect that
      // spans [yHigh, yLow] in `layoutColumns` and `layoutWaterfall` (and dropping the
      // labels and connectors inside it), or a clip node the scene does not have
      // (rect, polygon, line, text, ellipse, wedge, chevron, symbol, arrowhead).
      // Either is an owner's call, not a one-line swap, so what it costs is recorded
      // instead: on a slide that is not the chart's own canvas colour this shows as
      // a patch, and on a TRANSPARENT chart `style.background` is the keyword, which
      // every sink draws as no fill — so the cut hides nothing there and the two
      // zigzag lines are all that marks it.
      fill: style.background,
      name: "axis-break",
    },
    {
      kind: "line",
      x1: frame.x - left,
      y1: yLow + skew,
      x2: frame.x + frame.w + right,
      y2: yLow - skew,
      stroke: style.mutedText,
      strokeWidth: 1,
      name: "axis-break-lo",
    },
    {
      kind: "line",
      x1: frame.x - left,
      y1: yHigh + skew,
      x2: frame.x + frame.w + right,
      y2: yHigh - skew,
      stroke: style.mutedText,
      strokeWidth: 1,
      name: "axis-break-hi",
    },
  ];
}

export interface FrameReservations {
  titleH: number;
  totalsH: number;
  categoryAxisH: number;
  valueAxisW: number;
  seriesLabelsW: number;
  /** Height of the IBCS variance tier below the plot (0 when off). */
  varianceH: number;
}

/**
 * Height of the band above the plot that carries the totals row, or the gap
 * that stands in for it. Extracted so the frame and everything that has to
 * measure the height budget against it read one number.
 */
export function totalsBandHeight(decor: Decorations, style: ChartStyle): number {
  return decor.totals || decor.cagr || decor.difference || decor.grandTotal
    ? style.fontSize * 1.5 + 4
    : style.fontSize * 0.8;
}

/**
 * Height of the category-name strip below the plot, or the gap that replaces it.
 *
 * `lines` is how many lines the TALLEST name draws on (`categoryLines`): the strip
 * is one band for the whole axis and each name hangs from its top, so a name with
 * a line break (Alt+Enter in a datasheet cell) needs the pitch of every line past
 * the first on top of the one-line band. It was always one line, and the second
 * line of such a name was drawn below the plot's band — off the foot of the chart
 * in the preview, over whatever sits under the chart on a slide. Priced at the
 * chart font, not the size the names end up drawn at: the shrink that picks that
 * size runs later, in `chromeNodes`, and can only make a name smaller than this.
 */
export function categoryAxisHeight(decor: Decorations, style: ChartStyle, lines = 1): number {
  return decor.categoryAxis ? style.fontSize * 1.5 + 3 + extraLinesHeight(lines, style.fontSize) : 4;
}

/**
 * The most lines any category name draws on (1 when none holds a line break).
 *
 * ONE question asked by the reservation (`computeFrame`, `varianceBandHeight`) and
 * by the draw (`chromeNodes`), the way `horizontalLegendFits` is, so the strip
 * that is reserved and the names that are drawn cannot disagree about how tall
 * the tallest is.
 */
export function categoryLines(cfg: ChartConfig): number {
  return maxLineCount(cfg.data.categories ?? []);
}

/** `categoryAxisHeight` for THIS chart's names — the number a frame reserves. */
function categoryStripHeight(cfg: ChartConfig, decor: Decorations, style: ChartStyle): number {
  return categoryAxisHeight(decor, style, decor.categoryAxis ? categoryLines(cfg) : 1);
}

/**
 * Height of the value-axis strip under a SIDEWAYS plot, or the gap that replaces
 * it. The reservation `computeFrameHorizontal` and `horizontalLegendFits` both
 * make, so they cannot price it two ways.
 *
 * A tick label is a number and the number format's suffix, so a suffix with a line
 * break (a free-text field) puts every label on more than one line, and the strip
 * is the pitch of each line past the first taller. Priced at the largest size the
 * ticks are ever drawn at (`fs * 0.9`) and from the format alone, because the
 * frame is fixed before the scale and its ticks exist. `"0"` stands in for the
 * number, as the one label every axis carries and never more than one line.
 */
export function horizontalValueAxisHeight(cfg: ChartConfig, style: ChartStyle, decor: Decorations): number {
  const lines = lineCount(`0${cfg.numberFormat?.suffix ?? ""}`);
  return decor.valueAxis ? style.fontSize * 1.6 + extraLinesHeight(lines, style.fontSize * 0.9) : 4;
}

/**
 * Height reserved below the plot for the IBCS variance tier. Only the vertical
 * column family (stacked/clustered/100%, drawn by layoutColumns) actually paints
 * it, so gate the reservation on that — otherwise a line/area/waterfall/boxplot/
 * violin/candlestick chart with `decorations.variance` lost 4.5×fontSize of plot
 * to a strip that drew nothing.
 *
 * And ZERO when the frame cannot pay for it, which is the rule every other
 * reservation in this engine follows and this one did not. The band is a flat
 * `fontSize * 4.5` however tall the chart is — 81pt of a 60pt-tall chart at an
 * 18pt font — so on a small frame it took more height than existed. `fitPlot`
 * floors the plot and grows it UP from the bottom edge it was given, but the
 * category axis is placed BELOW that edge plus this band, so the names were
 * drawn off the foot of the chart: 8.1pt past a 300x60 frame at the DEFAULT
 * font, 85pt past a 120x90 one at 32pt, and on into whatever sits under the
 * chart on the slide.
 *
 * The room it is measured against is what the frame has left once the bands
 * that cannot be dropped have taken theirs, and it is computed from the same
 * three helpers `computeFrame` reserves with, so the two cannot drift. Both
 * call sites — the reservation and the category axis that sits under the band —
 * ask this one function, which is what keeps them agreeing.
 */
export function varianceBandHeight(cfg: ChartConfig, decor: Decorations, style: ChartStyle): number {
  const drawsTier =
    !cfg.horizontal && (cfg.kind === "stacked" || cfg.kind === "clustered" || cfg.kind === "stacked100");
  if (!decor.variance || !drawsTier) return 0;
  const want = style.fontSize * 4.5;
  const room =
    cfg.height -
    titleHeight(cfg, style) -
    totalsBandHeight(decor, style) -
    categoryStripHeight(cfg, decor, style) -
    footnoteH(cfg, style, decor);
  return room - want >= MIN_PLOT_SIDE ? want : 0;
}

/**
 * The first series' values, twice: clamped for GEOMETRY, and raw for TEXT.
 *
 * Three single-series kinds — funnel, waffle and cascade — opened with the same
 * line, `Math.max(0, data.series[0]?.values[c] ?? 0)`, and then formatted their
 * labels off the clamped array. So a blank cell and a negative both arrived at
 * the label as the number zero, and the chart ASSERTED it. Measured on
 * `[1000, null, 250]`:
 *
 *     funnel    stage-value-1  = "0"
 *     waffle    legend-label-1 = "Signups  0%"
 *     cascade   drop-label-1   = "Other: 1,000 (100.0%)"   <- a drop that never happened
 *
 * Every other kind — clustered, line, pie, treemap, sunburst — omits a blank
 * correctly. A missing bar reads as missing; a printed "0" reads as measured,
 * and cascade went on to compute a 100% collapse out of nothing.
 *
 * The split is the fix. `plot` is what geometry needs and may be clamped, since
 * a band cannot have negative width. `raw` keeps `null` for "no value" and keeps
 * a negative as itself, and it is what any TEXT must be formatted from — with
 * the rule that a figure DERIVED from an unknown is itself unknown and must not
 * be printed at all.
 */
export function firstSeriesValues(data: ChartConfig["data"], n: number): { plot: number[]; raw: (number | null)[] } {
  const raw = Array.from({ length: n }, (_, c) => {
    const v = data.series[0]?.values[c];
    return typeof v === "number" && Number.isFinite(v) ? v : null;
  });
  return { raw, plot: raw.map((v) => Math.max(0, v ?? 0)) };
}

/** Height reserved above the plot for the chart title (0 when untitled). */
export function titleHeight(cfg: ChartConfig, style: ChartStyle): number {
  return cfg.title ? style.fontSize * 1.6 + 6 : 0;
}

/** The size the title is actually drawn at, once it has been fitted to the width. */
export function titleFontSize(cfg: ChartConfig, style: ChartStyle): number {
  const text = String(cfg.title ?? "");
  // The HEIGHT bounds it too, and only this bound is not about the text. A
  // title is drawn from y=0 and its ink reaches `fontSize * 1.21` down, so a
  // chart shorter than that draws its own name off the bottom. That is not a
  // hypothetical frame: SMALL MULTIPLES build one chart per series at
  // `(height - titleH - gaps) / rows`, so a 300x60 chart with a title and two
  // panels asks for a 150x25 one, and its title overflowed by 0.9pt at an 18pt
  // font. Bounded here rather than in `titleNode` so `titleInkBottom` and
  // `printsOnTitle` — which decide what other chrome may be drawn — agree with
  // what is actually drawn.
  let tf = Math.min(style.fontSize * 1.2, cfg.height / 1.21);
  while (tf > 6 && textWidth(text, tf, true) > cfg.width) tf -= 0.5;
  return tf;
}

/**
 * How far down the title's INK reaches — not the band reserved for it.
 *
 * The two differ by more than they look: `titleHeight` is `fs * 1.6 + 6` while
 * the title is drawn at `fs * 1.2` from y=0, so the reservation runs some eight
 * points below the last of the title's ink at the default font. Anything asking
 * "would this label be drawn ON the title" has to ask about the ink, and the
 * first version of `aboveMarkFontSize` asked about the band instead: it shrank a
 * combo's top point label from 10pt to 6.2pt on two showcase slides where the
 * label had never touched the title at all. The deck diff is what caught it —
 * an ordinary chart at an ordinary size must not move.
 *
 * The reservation is still the right question for a plot or a band, which is
 * why both exist.
 */
/**
 * Would a label whose ink starts at `inkTop` be printed ON the title?
 *
 * `fitPlot` grows a squeezed plot UP from its bottom edge, and the bands below
 * the plot — the category names, the value axis's own ticks — are placed from
 * that edge. So on a frame that cannot pay for its chrome they all climb into
 * the title's band together: at 24pt on a 300x60 chart the category strip asks
 * for 51 of the 60 points available, and every name was drawn across the title
 * of its own chart. 165 of the 185 remaining text collisions were this, in one
 * frame.
 *
 * The answer is the one every other reservation here gives — chrome that cannot
 * be paid for is not drawn — and the title is what it yields to, because the
 * title is the one label that says what the reader is looking at. Measured
 * against the title's INK rather than its reserved band, so an ordinary chart,
 * whose axis sits well below both, is untouched.
 */
export function printsOnTitle(cfg: ChartConfig, style: ChartStyle, inkTop: number): boolean {
  return inkTop < titleInkBottom(cfg, style);
}

export function titleInkBottom(cfg: ChartConfig, style: ChartStyle): number {
  return cfg.title ? titleFontSize(cfg, style) * 1.21 : 0;
}

/**
 * The chart-title text node, or null when the chart has no title. Every layout
 * that draws its own title emitted this exact node inline; sharing it keeps the
 * title's size/weight/placement from drifting between chart kinds.
 */
export function titleNode(cfg: ChartConfig, style: ChartStyle): SceneNode | null {
  if (!cfg.title) return null;
  const fs = style.fontSize;
  // The title box spans the whole frame and the title was drawn at `fs * 1.2`
  // whatever the frame could hold, so a title longer than its chart ran off the
  // right-hand edge. Because this node is shared, that was ONE defect wearing
  // seventeen faces: at a 120x90 thumbnail the title is the worst-overflowing
  // node in thirteen of the seventeen kinds that overflow at all, by as much as
  // 124pt on a 120pt-wide chart — a title four times wider than its own chart,
  // lying across whatever sits beside it on the slide.
  //
  // Shrink to fit, then clip: the same two-step the agenda, the process flow,
  // the mekko axis and the funnel's rows all use. Last-resort, so a title that
  // already fits keeps `fs * 1.2` exactly and no ordinary chart moves.
  //
  // `titleHeight` keeps reserving the full `fs * 1.6` and is deliberately not
  // shrunk with this. Reserving MORE room than the title now needs cannot push
  // anything off the frame, where reserving less could, and leaving it alone
  // means the plot below starts where it always did.
  const text = String(cfg.title ?? "");
  const tf = titleFontSize(cfg, style);
  return {
    kind: "text",
    x: 0,
    y: 0,
    w: cfg.width,
    h: fs * 1.6,
    text: clipToWidth(text, tf, cfg.width, true),
    fontSize: tf,
    bold: true,
    color: style.text,
    align: "left",
    valign: "top",
    name: "title",
  };
}

/**
 * The footnote line, or null when there is nothing to print on it.
 *
 * Extracted for the reason `titleNode` above was: both call sites emitted this
 * exact node inline, byte for byte apart from the text, so a fix to one was a
 * fix to one. It also had the title's defect — the box spans the frame and the
 * text was drawn at `fs * 0.85` whatever the frame could hold, so a long source
 * note ran off the right-hand edge (43pt on a 120pt-wide cascade).
 *
 * Shrink to fit, then clip, last-resort: a footnote that already fits keeps its
 * size and position exactly.
 */
export function footnoteNode(cfg: ChartConfig, style: ChartStyle, text: string): SceneNode | null {
  if (!text) return null;
  const fs = style.fontSize;
  const room = Math.max(1, cfg.width - 4);
  let ff = fs * 0.85;
  while (ff > 5 && textWidth(text, ff) > room) ff -= 0.5;
  return {
    kind: "text",
    x: 2,
    y: cfg.height - fs * 1.15,
    w: room,
    h: fs * 1.1,
    text: clipToWidth(text, ff, room),
    fontSize: ff,
    color: style.mutedText,
    align: "left",
    valign: "bottom",
    name: "footnote",
  };
}

/**
 * The denominator behind a "100% = N" note: the series total for pies, the
 * uniform per-category denominator for 100% charts (null when categories
 * have different denominators — the note would be a lie then).
 */
export function hundredPercentTotal(cfg: ChartConfig): number | null {
  const { data, kind } = cfg;
  if (kind === "pie" || kind === "doughnut") {
    const total = data.categories.reduce((a, _, c) => a + Math.max(0, data.series[0]?.values[c] ?? 0), 0);
    return total > 0 ? total : null;
  }
  if (kind === "stacked100") {
    const denominators = data.categories.map((_, c) => {
      const d = data.hundredPercent?.[c];
      return d != null && d > 0 ? d : data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0);
    });
    if (!denominators.length || denominators[0] <= 0) return null;
    return denominators.every((d) => Math.abs(d - denominators[0]) < 1e-9) ? denominators[0] : null;
  }
  return null;
}

/** Height reserved at the bottom for the footnote / "100% =" line. */
export function footnoteH(cfg: ChartConfig, style: ChartStyle, decor: Decorations): number {
  // scatter.spread prints its cap on the footnote line, so it needs the row
  // reserved even when the author wrote no footnote of their own.
  //
  // `hundredPercentNote` asks for a note that only EXISTS for a pie, a doughnut
  // or a 100% chart with one denominator — `hundredPercentTotal` answers null
  // for everything else, and the note is then not drawn. The reservation asked
  // the flag instead, so switching it on took a row off every other kind for a
  // line that was never written: the plot lost `fontSize * 1.3`, `fitPlot`
  // floored it and grew it back UP through the totals band, and the labels in
  // that band were drawn above the top of the chart — 19.8pt above a mekko at
  // 80x60, and the same on stacked, waterfall, scatter and bubble.
  //
  // The reservation and the draw ask one question now, which is the rule the
  // horizontal legend already follows and the reason it stopped drifting.
  const note = decor.hundredPercentNote && hundredPercentTotal(cfg) != null;
  return cfg.footnote || note || cfg.scatter?.spread ? style.fontSize * 1.3 : 0;
}

/**
 * One legend entry's placement: its top-left x, and the row it starts on.
 *
 * `row` is a distance DOWN the legend in row pitches (`fs * 1.6`, which every
 * drawer and every reservation multiplies it by), 0 for the first. For a legend
 * of names with no line break that is the 0-based wrap row it always was; a name
 * that holds a break is taller than one pitch, so the rows under it start `units`
 * pitches lower rather than one.
 */
export interface LegendSlot {
  x: number;
  row: number;
  /** Pitches this entry's own label occupies: 1 for a one-line name. */
  units: number;
}

/** A legend row's pitch as a multiple of the font size — what the drawers price a row at. */
const LEGEND_ROW_PITCH = 1.6;

/**
 * The whole row pitches a legend label of this text occupies. A name with a
 * line break draws `1.4 * fs` for its first line and `LINE_PITCH * fs` for each
 * further one, which the row units round UP to — never down, so a reservation
 * can only be larger than the ink, and 1 for a name with no break.
 */
function legendUnits(label: string): number {
  const lines = lineCount(label);
  return lines > 1 ? 1 + Math.ceil(((lines - 1) * LINE_PITCH) / LEGEND_ROW_PITCH) : 1;
}

/**
 * The wrap walk shared by `legendRow` (which places the chips) and
 * `legendRowCount` (which reserves the vertical space they need). Extracting it
 * means the drawer and the reserver can never disagree on how many rows a legend
 * occupies. Chips march left-to-right; one that would cross `maxX` starts a new
 * row (never the first entry of a row, so a single over-wide label still draws).
 */
export function legendWrapWalk(labels: string[], fs: number, x0: number, maxX: number): LegendSlot[] {
  const chip = fs * 0.7;
  const slots: LegendSlot[] = [];
  let x = x0;
  let row = 0;
  /** Pitches the current wrap row takes: its tallest label's. */
  let tall = 1;
  for (const label of labels) {
    const wLabel = textWidth(label, fs);
    if (x > x0 && x + chip + 3 + wLabel > maxX) {
      x = x0;
      row += tall;
      tall = 1;
    }
    const units = legendUnits(label);
    tall = Math.max(tall, units);
    slots.push({ x, row, units });
    x += chip + 3 + wLabel + 12;
  }
  return slots;
}

/**
 * Row pitches the wrapping legend occupies (0 when there are no labels): the
 * number of wrap rows when no name holds a line break, which is every legend
 * that has ever been drawn here, and the height of the rows in pitches when one
 * does — every drawer and reservation multiplies this by `fs * 1.6`.
 */
export function legendRowCount(labels: string[], fs: number, x0: number, maxX: number): number {
  return legendWrapWalk(labels, fs, x0, maxX).reduce((end, s) => Math.max(end, s.row + s.units), 0);
}

/**
 * Labels for the default series legend (the one `horizontalChrome` and the
 * horizontal mekko draw): a scenario-qualified series name, and only when there
 * is more than one series to tell apart. The height-reserving frame code and the
 * chip-drawing `legendRow` both derive their labels here so they legend the same
 * set and their wrap walks land on the same row count.
 */
export function seriesLegendLabels(cfg: ChartConfig): string[] {
  if (cfg.data.series.length <= 1) return [];
  return cfg.data.series.map((s) => (s.scenario ? `${s.name} (${s.scenario})` : s.name));
}

/**
 * Every label the legend must carry — the config's own series, plus any the
 * layout has told us about through `decor.legendAlso`.
 *
 * ONE FUNCTION, because three places need this list and they must agree: the
 * predicate that decides whether a legend fits, the band the frame reserves for
 * it, and the row that draws it. `seriesLegendLabels` was called directly by all
 * three, which was correct while nothing was ever missing from `cfg.data.series`
 * — and a combo's line series always is. See `Decorations.legendAlso`.
 *
 * The single-series short-circuit lives in `seriesLegendLabels` and is deliberately
 * NOT repeated here: a one-column combo with one line is two series between them
 * and wants a legend naming both, which is exactly the case that would be lost
 * by asking `cfg.data.series.length <= 1` after the split.
 */
export function legendLabelsFor(cfg: ChartConfig, decor: Decorations): string[] {
  const extra = (decor.legendAlso ?? []).map((e) => e.label);
  if (!extra.length) return seriesLegendLabels(cfg);
  return [...cfg.data.series.map((s) => (s.scenario ? `${s.name} (${s.scenario})` : s.name)), ...extra];
}

/**
 * Compute the plot rectangle by reserving margins for the enabled decorations
 * (title, totals row, category labels, value axis, right-hand series labels).
 */
export function computeFrame(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  seriesNames: string[],
): { frame: Frame; res: FrameReservations } {
  const fs = style.fontSize;
  const titleH = titleHeight(cfg, style);
  const totalsH = totalsBandHeight(decor, style);
  const categoryAxisH = categoryStripHeight(cfg, decor, style);
  const valueAxisW = decor.valueAxis ? 34 : 2;
  const seriesLabelsW = decor.seriesLabels
    ? Math.min(cfg.width * 0.3, Math.max(0, ...seriesNames.map((s) => textWidth(s, fs))) + 14)
    : 2;
  // Extra headroom when a difference arrow is drawn past the last column.
  const diffW = decor.difference ? 26 : 0;
  const varianceH = varianceBandHeight(cfg, decor, style);
  const frame: Frame = fitPlot(cfg, {
    x: valueAxisW,
    y: titleH + totalsH,
    w: cfg.width - valueAxisW - Math.max(seriesLabelsW, diffW + 2) - 2,
    h: cfg.height - titleH - totalsH - varianceH - categoryAxisH - footnoteH(cfg, style, decor),
  });
  return { frame, res: { titleH, totalsH, categoryAxisH, valueAxisW, seriesLabelsW, varianceH } };
}

/**
 * Plot rectangle for horizontal (bar) orientation: category labels on the
 * left, value axis at the bottom, totals to the right of the bar ends,
 * series legend row at the top.
 */
/**
 * Can this horizontal frame pay for its series legend?
 *
 * ONE predicate, read by the reservation in `computeFrameHorizontal` and by the
 * draw in `horizontalChrome`, so the two cannot disagree. They have disagreed
 * twice already — the mekko drew its legend at a widened `frame.x` the
 * reservation had not counted for, and a scatter fix that gated the reservation
 * alone left the legend drawn over a band reserved as zero, which measured
 * strictly worse than drawing it with room. A shared function is the only
 * version of "they agree" that cannot rot.
 *
 * On a frame with room the answer is always yes, so every ordinary chart is
 * untouched: this fires when the legend's own rows plus the title, the value
 * axis and the footnote leave less than a plot.
 */
export function horizontalLegendFits(cfg: ChartConfig, style: ChartStyle, decor: Decorations): boolean {
  const fs = style.fontSize;
  const labels = decor.seriesLabels ? legendLabelsFor(cfg, decor) : [];
  if (!labels.length) return true;
  const catW = decor.categoryAxis
    ? Math.min(cfg.width * 0.3, Math.max(0, ...cfg.data.categories.map((c) => textWidth(c, fs))) + 8)
    : 2;
  const rows = legendRowCount(labels, fs, catW, cfg.width - 4);
  const legendH = rows > 0 ? rows * (fs * 1.6) + 4 : fs * 0.6;
  const left =
    cfg.height -
    titleHeight(cfg, style) -
    legendH -
    horizontalValueAxisHeight(cfg, style, decor) -
    footnoteH(cfg, style, decor);
  // A plot thinner than a couple of rows of text is not a chart, and a legend
  // that has eaten it is naming bars nobody can see. Same judgement the scatter
  // legend and the cascade's group headers make.
  return left >= fs * 2.5;
}

export function computeFrameHorizontal(cfg: ChartConfig, style: ChartStyle, decor: Decorations): Frame {
  const fs = style.fontSize;
  const titleH = titleHeight(cfg, style);
  const catW = decor.categoryAxis
    ? Math.min(cfg.width * 0.3, Math.max(0, ...cfg.data.categories.map((c) => textWidth(c, fs))) + 8)
    : 2;
  // Reserve one row per wrapped legend row, not a fixed single row: many/long
  // series names on a narrow chart make legendRow wrap (same walk), and each
  // extra row must push the plot down or the legend draws on top of the bars.
  // The walk runs at the legend's own x0 (frame.x === catW) and maxX
  // (cfg.width - 4), so this reservation and legendRow agree on the row count. A
  // one-row legend keeps the old fs*1.6+4 exactly, so snapshots stay identical.
  const legendLabels = decor.seriesLabels && horizontalLegendFits(cfg, style, decor) ? legendLabelsFor(cfg, decor) : [];
  const legendRows = legendLabels.length ? legendRowCount(legendLabels, fs, catW, cfg.width - 4) : 0;
  const legendH = legendRows > 0 ? legendRows * (fs * 1.6) + 4 : fs * 0.6;
  const valueAxisH = horizontalValueAxisHeight(cfg, style, decor);
  const totalsW = decor.totals ? fs * 4 : fs * 0.8;
  return fitPlot(cfg, {
    x: catW,
    y: titleH + legendH,
    w: cfg.width - catW - totalsW - 2,
    h: cfg.height - titleH - legendH - valueAxisH - footnoteH(cfg, style, decor),
  });
}

/** Title, category labels, value axis, gridlines — shared chrome for all cartesian charts. */
export function chromeNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  frame: Frame,
  centers: number[],
  scale?: ValueScale,
): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;
  const t = titleNode(cfg, style);
  if (t) nodes.push(t);
  if (decor.gridlines && scale) {
    for (const t of scale.ticks) {
      if (t === 0) continue;
      const y = scale.toY(t);
      nodes.push({
        kind: "line",
        x1: frame.x,
        y1: y,
        x2: frame.x + frame.w,
        y2: y,
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: "gridline",
      });
    }
  }
  if (decor.valueAxis && scale) {
    // "datamarks": Tufte-style tick dashes + labels, no axis line, no
    // gridlines; tickMode "data" places them at the data extremes instead
    // of nice round values (the range frame).
    const marks = decor.valueAxis === "datamarks";
    const ticks = marks && decor.tickMode === "data" ? [...new Set([scale.min, scale.max])] : scale.ticks;
    // Tick labels are read against each other, so their precision comes from the
    // tick step, not the tick magnitude — otherwise a narrow axis prints the same
    // label at several heights. See axisTickLabel for that and for the share case.
    const axisLabel = axisTickLabel(ticks, scale.percent, cfg.numberFormat);
    /**
     * The size the tick labels are drawn at.
     *
     * One label per tick, each centred on its own tick, so the room each has is
     * the SPACING between adjacent ticks — and none was fitted to it. At a font
     * large relative to the plot the labels were drawn over each other, which is
     * the single commonest text collision left in this engine after the category
     * axis: 71 of the 237 overlapping pairs a sweep found, here and in the
     * scatter's own axes.
     *
     * Bound by that spacing, the way the radar's ring ticks are bound by their
     * ring gap. Last resort: on any chart whose ticks already clear each other
     * this is `fs * 0.9` and nothing moves.
     *
     * Dropped outright when the spacing cannot pay for a LEGIBLE label. A fit
     * with no floor answers whatever the arithmetic says, and on a plot squeezed
     * to a few points that is a stack of one-point labels — ink no reader can
     * resolve, and a fit that reports success. Same answer the radar, sunburst,
     * tilemap and pie reservations give when their band cannot be met: a label
     * that cannot be read is not there. The datamarks and gridlines stay, since
     * those still carry the scale.
     */
    const tickGap =
      ticks.length > 1
        ? Math.min(...ticks.slice(1).map((t, i) => Math.abs(scale.toY(t) - scale.toY(ticks[i]))))
        : frame.h;
    /**
     * A tick label can hold a line break — only through the number format's
     * prefix or suffix, but those are free text — and then it is as many pitches
     * tall as it has lines. The room it wants between ticks grows with the most
     * of them, so the whole axis shrinks together exactly as it does for a gap
     * that is too tight for one line; a label with no break wants `fs * 1.4`,
     * unchanged.
     */
    const tickTexts = ticks.map((t) => axisLabel(t));
    const tickScale = Math.min(1, tickGap / (fs * 1.4 + extraLinesHeight(maxLineCount(tickTexts), fs * 0.9)));
    const tickLegible = fs * 0.9 * tickScale >= MIN_LABEL_FS;
    for (const [ti, t] of ticks.entries()) {
      const y = scale.toY(t);
      if (marks) {
        nodes.push({
          kind: "line",
          x1: frame.x - 4,
          y1: y,
          x2: frame.x,
          y2: y,
          stroke: style.axis,
          strokeWidth: 1,
          name: "datamark",
        });
      }
      if (!tickLegible) continue;
      // Per TICK, not per axis: a squeezed plot puts its topmost ticks in the
      // title's band while the ones below it are still perfectly placed, and
      // each tick is a number in its own right.
      // This tick's own extra lines: the box grows by them and stays centred on
      // the tick, so the label still reads against the mark it names.
      const grow = extraLinesHeight(lineCount(tickTexts[ti]), fs * 0.9 * tickScale);
      if (printsOnTitle(cfg, style, y - fs * 0.7 * tickScale - grow / 2)) continue;
      nodes.push({
        kind: "text",
        x: 0,
        y: y - fs * 0.7 * tickScale - grow / 2,
        w: marks ? frame.x - 6 : frame.x - 4,
        h: fs * 1.4 * tickScale + grow,
        text: tickTexts[ti],
        fontSize: fs * 0.9 * tickScale,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "value-axis",
      });
    }
  }
  if (cfg.valueAxisTitle) {
    /**
     * A SHORT UNIT AT THE TOP OF THE VALUE AXIS, which is what this option is
     * and what it says it is — `ChartConfig.valueAxisTitle` is documented as
     * "units label shown at the top of the value axis (e.g. `€m`)", and the two
     * uses in the shipped showcase are `€m` and `$m (log)`.
     *
     * That mattered, because the family this label produced was measured with a
     * twenty-seven-character sentence and the count tracks LENGTH:
     *
     *     €m                             75 pairs
     *     $m (log)                      189
     *     EUR millions                  224
     *     Revenue in millions of euro   286
     *
     * SO THE WIDTH IS CLIPPED rather than allowed to span the chart. Clipping
     * is what this engine already does to gantt task names and category names,
     * and it is the one remedy that keeps AUTHOR TEXT: `€m` is untouched by it,
     * where the three remedies tried on 2026-08-28 all DELETED the unit. A share
     * of the chart rather than the axis column, because a chart drawn without a
     * value axis has no column at all — which is exactly what sank the
     * 2026-08-19 attempt to fit this to `frame.x`. Measured across the variant
     * sweep: **1,327 overlapping pairs to 1,014**, and no shape that was not
     * already there.
     *
     * AND THE `y` IS LEFT ALONE, which is the part worth reading. `title` is 205
     * of the family and is length-INDEPENDENT — identical for two characters and
     * for twenty-seven — because `Math.max(0, …)` parks the unit in the title's
     * band on any chart whose plot starts high. That looks exactly like the
     * clamp bug this engine has recorded five times, and flooring it at the
     * title's ink was written, measured, and REFUSED:
     *
     *     clip alone                1,014 pairs, no new shapes
     *     clip + floor at the ink   1,156 pairs, and TWO new families —
     *                               `value-axis-title / category#` at 310
     *
     * The floor takes the title collisions to zero and buys them by moving the
     * unit down into the category names on short charts. The CAGR caption's own
     * note two files over says the same thing about the same move: "a clamp
     * moves a label whether or not the destination is free". A clamp that has a
     * free destination is a fix; this one does not.
     *
     * WHAT THIS DOES NOT DO is decide where a unit belongs on a crowded chart.
     * The 205 on the title stay, and the rest is the unit against the topmost
     * tick and the totals row — genuinely contested space. The three relocations
     * in `docs/BACKLOG.md` — end of the axis, folded into the top tick, its own
     * gutter — all redesign every chart to accommodate long text this option
     * does not support, and none of them is bought by this change.
     */
    const unitFs = fs * 0.95;
    /** At most this share of the chart, so a long unit cannot span it. */
    const room = cfg.width * 0.4;
    /**
     * THE UNIT YIELDS TO THE TITLE, like everything else in this band already
     * does — and this is the whole of `title / value-axis-title`.
     *
     * That pair is LENGTH-INDEPENDENT: 205 of it across the sweep whether the
     * unit is two characters or twenty-seven. Both nodes are `align: "left"` at
     * `x: 0`, so their ink always shares the x range, and only the y decides.
     * `Math.max(0, …)` then parks the unit inside the title's band on a chart
     * whose plot cannot start below both. No width remedy — clip, gutter-fit,
     * shrink — can move that number, and all three were tried.
     *
     * MEASURED, on the 176 charts that draw a title and a unit: 22 overlap, and
     * the split is clean. Every overlapping pair has an ink gap of -2.56pt or
     * worse; every clear one has +3.31pt or better. There is no borderline case
     * to get wrong, so the rule is simply whether the ink clears.
     *
     * ON INK, NOT ON BOXES, and the box version was written first and was wrong.
     * `bandTop < titleHeight(cfg, style)` reads plausibly — and drops the unit
     * from a clustered chart at 480x300, a size people present at. That is the
     * exact objection that sank the 2026-08-28 ink-overlap remedy, arrived at by
     * the exact mistake this repo keeps making: measuring the box when the
     * question is about the ink.
     *
     * WHY DROPPING IS RIGHT HERE, when dropping author text is normally refused.
     * The 22 are all `80x60` and `300x60` at 18pt, and on those charts the
     * engine ALREADY drops the category names, the axis strip and the legend —
     * "Chrome yields to the title" in docs/MANUAL.md, whose worked example is a
     * 300x60 banner. The unit was the one thing in that band still drawn on top
     * of the title. It is not being singled out; it was the exception.
     *
     * On every frame anybody presents at, the unit is untouched: 154 of 176 kept,
     * and 100% at 480x300 and above.
     *
     * THE CONSTANTS HERE ARE NOT LOAD-BEARING, which is worth saying because it
     * looks like they should be. Three mutants survive `value-axis-title.test.ts`
     * and all three are equivalent, not gaps: `* 1.21` moved to `* 1.9` or
     * `* 1.0`, and the `Math.max(0, …)` taken back out of the decision, each
     * leave the drawn/dropped sets identical at 154/22. That follows from the
     * measurement — the dead zone between the worst clear gap (+3.31pt) and the
     * best overlapping one (-2.56pt) is 5.9pt wide, so every threshold inside it
     * decides every chart the same way. The test pins the OUTCOME (the counts
     * and the two frames) rather than the arithmetic, which is why it survives
     * them and would still catch a rule that reached further.
     */
    const bandTop = Math.max(0, frame.y - fs * 1.5);
    /**
     * A unit with a LINE BREAK rises. The box is anchored at its bottom, so the
     * last line stays exactly where a one-line unit sits and each earlier line
     * stacks above it, and the box grows upward by the pitch of the lines past
     * the first. Zero for a unit with no break, so that unit's box is untouched.
     */
    const unitRise = extraLinesHeight(lineCount(cfg.valueAxisTitle), unitFs);
    // The ink of the box as it will ACTUALLY be drawn — the clamped `y`, not the
    // raw one. Deciding on the unclamped position drops the unit from short
    // charts that have no title to collide with, which is a worse bug than the
    // one this fixes and is what the first version of this did.
    const unitInkTop = bandTop + fs * 1.4 - unitFs * 1.05 - unitRise;
    const titleInkBottom = titleFontSize(cfg, style) * 1.21;
    /** Only ever against a title that is drawn: nothing else is up there. */
    const yieldsToTitle = !!cfg.title && unitInkTop < titleInkBottom;
    /**
     * And a unit whose raised lines would leave the canvas yields the same way,
     * for the same reason it yields to the title: a line printed above the chart
     * is on the slide, not on the chart, and a clamp back inside would only move
     * it onto the top tick (see above — "a clamp moves a label whether or not the
     * destination is free"). It cannot happen to a one-line unit, whose ink top
     * is `bandTop + 0.4 * fs` at the least, so only a unit with a break asks.
     */
    const offCanvas = unitRise > 0 && unitInkTop < 0;
    const unitText = yieldsToTitle || offCanvas ? "" : clipToWidth(cfg.valueAxisTitle, unitFs, room);
    // Nothing legible fits — drop rather than push an empty box into the scene,
    // which the de-collision pass and every readback would then have to carry.
    if (unitText)
      nodes.push({
        kind: "text",
        x: 0,
        y: bandTop - unitRise,
        w: Math.min(Math.max(frame.x - 4, textWidth(unitText, fs)), room),
        h: fs * 1.4 + unitRise,
        text: unitText,
        fontSize: unitFs,
        color: style.mutedText,
        align: "left",
        valign: "bottom",
        name: "value-axis-title",
      });
  }
  const catY = frame.y + frame.h + varianceBandHeight(cfg, decor, style) + 3;
  // The whole strip or none of it: these names share one y, so a chart short
  // enough to push them into the title pushes all of them.
  //
  // AND THE SAME AT THE OTHER END. `fitPlot` floors the plot at
  // `MIN_PLOT_SIDE`, so on a frame shorter than that floor the plot's own bottom
  // edge — which is where this strip is placed from — is already past the foot
  // of the chart. That is not a hypothetical size: SMALL MULTIPLES divide the
  // frame into one chart per series, and a titled 300x60 chart with three of
  // them asks for panels 7.6 points tall, whose category names were drawn 21.6
  // points below the chart. A name printed under the chart is on the slide, not
  // on the chart, which is the same reason it may not be printed on the title.
  // The ink of the LAST line of the tallest name: a name with a line break hangs
  // its further lines below the first, `categoryAxisHeight` has priced them into
  // the strip, and the names are all drawn or none for the reason above.
  const catInk = catY + fs * 1.21 + extraLinesHeight(categoryLines(cfg), fs);
  if (decor.categoryAxis && !printsOnTitle(cfg, style, catY) && catInk <= cfg.height) {
    // Each name's own slot — the same width for all of them unless the centres
    // are uneven (a date-spaced line); see `categoryRooms`.
    const slots = categoryRooms(centers, frame.w, frame.x, frame.x + frame.w);
    const names = cfg.data.categories.map((c) => String(c ?? ""));
    const rooms = slots.map((s) => Math.max(1, s - 2));
    /**
     * One size for the whole axis, small enough that each name fits its slot.
     *
     * Every label is CENTRED in a slot `slots[i]` wide and none was ever fitted to
     * it, so a name wider than its slot runs into its neighbours on both sides.
     * That is the commonest text collision in this engine — 30 of the 73
     * kind/font/frame combinations a sweep turned up — and it is one defect
     * rather than seven because this axis is shared by every cartesian kind.
     *
     * Shrunk together, so the axis reads at one size rather than as a ransom
     * note, then clipped for the name no floor can fit. Last resort as
     * everywhere else: where the names already fit their slots this is `fs` and
     * nothing moves.
     *
     * A name whose slot is narrower than an ellipsis at the floor is gone at
     * every size, so it has no say in which size that is. With one slot width for
     * the whole axis that excuses all of the names or none of them, and changes
     * nothing; with a date-spaced line two points a day apart each get a slot of
     * a few points, and letting them choose took every other name on the axis
     * down to the 5pt floor with them.
     */
    const catFs = (() => {
      const counts = rooms.map((room) => room >= textWidth("…", 5));
      let f = fs;
      while (f > 5 && names.some((c, i) => counts[i] && textWidth(c, f) > rooms[i])) f -= 0.5;
      return f;
    })();
    names.forEach((name, i) => {
      nodes.push({
        kind: "text",
        x: centers[i] - slots[i] / 2,
        y: catY,
        w: slots[i],
        // Top-anchored, so a name with a line break grows DOWN by the pitch of
        // its further lines; a name with none keeps the box it always had, even
        // beside one that has.
        h: catFs * 1.4 + extraLinesHeight(lineCount(name), catFs),
        text: clipToWidth(name, catFs, rooms[i]),
        fontSize: catFs,
        color: style.text,
        align: "center",
        valign: "top",
        name: `category-${i}`,
      });
    });
  }
  return nodes;
}

/** Baseline (zero line) — drawn on top of columns, think-cell style. */
export function baselineNode(frame: Frame, y: number, style: ChartStyle): SceneNode {
  return {
    kind: "line",
    x1: frame.x,
    y1: y,
    x2: frame.x + frame.w,
    y2: y,
    stroke: style.axis,
    strokeWidth: 1,
    name: "baseline",
  };
}
