import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors, NumberFormat } from "./types";
import { textWidth, type SceneNode } from "./scene";
import { cagr, formatNumber, formatPercent, resolveFormat } from "./format";
import { MIN_LABEL_FS, titleInkBottom, titleNode } from "./layout/frame";

/**
 * THE MEAN OF WHAT WAS MEASURED, not of what was left after coercion.
 *
 * Every layout builds `columnValue` with `?? 0`, so a blank category arrives
 * here as a zero AND as a denominator: the mean of `[100, null, 200]` comes out
 * 100 where the mean of the numbers actually supplied is 150. Both halves are
 * wrong, which is why dropping the blank from the SUM alone does not fix it.
 *
 * `columnHasData` is optional and absent means every category counts — the
 * honest default for a layout whose `columnValue` has no blank to coerce (a
 * waterfall's running level carries across an empty step) or whose kind never
 * reaches this file with a mean or an arrow. It is NOT "a close, a median or a
 * column total is never blank": it was assumed for those four and each of them
 * can be, so candlestick, boxplot, violin and mekko answer it too. See
 * LayoutAnchors.
 *
 * ONE DEFINITION, BECAUSE THERE WERE TWO CALLERS AND THEY DISAGREED. The filter
 * was written for the value line, and the difference arrow's `fromValueLine`
 * anchor kept an unfiltered copy — so a chart carrying both drew TWO means.
 * Measured 2026-09-28 on `[100, null, 200]`: the value line sat at 150 labelled
 * "Ø 150" while the arrow's baseline sat at 100, and the arrow printed "+100%"
 * where the truth is "+33%". A wrong number in ink is the worst thing this
 * renderer can do. The commit that added the filter (`35840f6`) closed with
 * "the second time this week the fix was already written and one call site was
 * missed"; a shared function is what stops the third.
 */
function meanOfMeasured(a: LayoutAnchors): number {
  const measured = a.columnValue.filter((_, i) => a.columnHasData?.[i] ?? true);
  return measured.reduce((s, v) => s + v, 0) / (measured.length || 1);
}

/**
 * The format for a DERIVED number drawn beside the chart's own labels — a mean
 * line, a difference — at the finer of two precisions: the one the surrounding
 * data is labelled in, and the one the derived number itself asks for.
 *
 * Both halves are load-bearing, and each was wrong on its own.
 *
 * ONLY THE VALUE'S OWN. This is what these two labels used to do, by handing
 * `cfg.numberFormat` straight to `formatNumber` — whose "auto" ladder stops at
 * two decimals below 1, where `resolveFormat` widens to six. On bars labelled
 * 0.0040 / 0.0050 / 0.0060 the mean line, drawn correctly at 0.005, captioned
 * itself "Ø 0.01": twice the mean, and outside the range of every bar.
 *
 * ONLY THE CHART'S. A mean is fractional where its columns are whole, so
 * [-40, -20, -10, 60] — labelled at no decimals, because 60 needs none — turned
 * its true mean of -2.5 into "Ø -3". Nothing on the chart says -3.
 *
 * The finer of the two is right for both, and it is the rule
 * `resolveAxisFormat` already applies to a tick strip (`max(magnitude
 * precision, step precision)`). An authored `decimals` collapses both sides
 * onto itself, so it still wins outright.
 */
function finerFormat(around: number[], value: number, fmt: ChartConfig["numberFormat"]): NumberFormat {
  const context = resolveFormat(around, fmt);
  const own = resolveFormat([value], fmt);
  const dec = (f: NumberFormat) => (typeof f.decimals === "number" ? f.decimals : 0);
  return { ...context, decimals: Math.max(dec(context), dec(own)) };
}

/**
 * Where series `si` tops out at category `c`, in value units — the level a
 * SERIES-QUALIFIED decoration (a level arrow, a CAGR arrow, a callout) stands on.
 *
 * Three answers, and the caller needs all three:
 *  - a number: the series has a mark here, drawn at `valueToY(level)`;
 *  - `NaN`: the layout knows this series and has no place on this chart's value
 *    scale for it (a combo's line on its own axis) — a level read as a number
 *    would be drawn on the wrong scale;
 *  - `undefined`: this layout publishes no levels, or the index is off the end
 *    of them. The decoration keeps its category-wide anchor, which is what it
 *    always did for those.
 *
 * ONE READER, because the difference arrow and the callout each carried their
 * own copy of this guard and the CAGR arrow carried none, so the one decoration
 * with no guard was the one that anchored a stacked series inside the series
 * beneath it.
 */
function levelOf(a: LayoutAnchors, c: number, si: number | undefined): number | undefined {
  return si != null && si >= 0 && a.valueToY != null ? a.seriesLevels?.[c]?.[si] : undefined;
}

/**
 * Did anything get MEASURED at category `c`? — the question `columnValue`
 * cannot answer, because every layout builds it with `?? 0`.
 *
 * An arrow reaching for a blank end read it as a measured zero and PRINTED
 * "-100%" (or "+50") for a value nobody entered: `[50, 60, null]` with an
 * arrow 0→2 was byte-identical to `[50, 60, 0]`. A blank is not a zero to
 * compare against, so the arrow is not drawn — the same answer the column
 * total gives (`total-3` is withheld, not printed as "0"), and for the same
 * reason. Absent `columnHasData` means every category counts; see
 * LayoutAnchors.
 */
const categoryMeasured = (a: LayoutAnchors, c: number) => a.columnHasData?.[c] ?? true;

/**
 * Does the end of a LEVEL arrow at category `c` stand on a measurement?
 *
 * Asked of the series itself when the layout publishes a level for the series
 * the arrow names (`si` indexes `cfg.data.series`), and of the category when the
 * arrow compares totals. A series with no place on this chart's scale (NaN, see
 * `levelOf`) cannot be pointed at either.
 */
function levelEndMeasured(cfg: ChartConfig, a: LayoutAnchors, c: number, si: number | undefined): boolean {
  const level = levelOf(a, c, si);
  if (level === undefined) return categoryMeasured(a, c);
  return Number.isFinite(cfg.data.series[si!]?.values[c]) && !Number.isNaN(level);
}

/**
 * Where one end of a CAGR arrow is anchored, and whether it stands on a
 * measurement. `si` names the series the rate is for, or nothing for the totals.
 *
 * Anchored on the value it PRINTS, for the reason spelled out on the difference
 * arrow in `decorationNodes` — and when a series is named, the rate is that
 * series' and the arrow used to sit on the drawn top of whichever series happened
 * to be highest. Then on the MARK that value is drawn as: a series' own value is
 * where its bar tops out only when it stands on the baseline (clustered, line);
 * stacked, its segment sits on top of the series beneath it, so
 * `valueToY(own value)` is a height at which the series has no ink at all, and the
 * arrow, its head and its caption were drawn through the segments below — a
 * rising arrow across a series that is flat. `levelOf` is the segment's top in
 * every case and equals the own value wherever the two agree (clustered, line),
 * so those charts do not move.
 *
 * MEASURED at both ends, or there is no growth to draw: a blank is not a zero to
 * grow from or to. `cagr()` already refused the number ("CAGR n/a"), but the
 * arrow was still drawn from the baseline of a category with no data, under a
 * caption that said it had none. A series with no place on this chart's scale
 * (NaN) is the same answer — an arrow measured on the wrong scale is not an
 * arrow.
 */
function cagrEnd(
  cfg: ChartConfig,
  a: LayoutAnchors,
  c: number,
  si: number | undefined,
): { y: number; measured: boolean } {
  const sVals = si != null ? cfg.data.series[si]?.values : undefined;
  if (!sVals)
    return { y: a.valueToY ? a.valueToY(a.columnValue[c]) : a.columnTop[c], measured: categoryMeasured(a, c) };
  const level = levelOf(a, c, si);
  const at = level !== undefined && Number.isFinite(level) ? level : (sVals[c] ?? 0);
  return {
    y: a.valueToY ? a.valueToY(at) : a.columnTop[c],
    measured: Number.isFinite(sVals[c]) && !Number.isNaN(level),
  };
}

/**
 * think-cell's signature annotations, computed from layout anchors so they
 * work identically across chart types.
 */
export function decorationNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  a: LayoutAnchors,
): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;

  // Every decoration here anchors to the category geometry (categoryX / columnTop
  // / seriesLevels). A chart that drew NO segments — a single-0 pie or sunburst,
  // say — has empty anchor arrays, so any anchored read comes back undefined and
  // the decoration would emit NaN geometry. There is nothing to annotate.
  if (!a.categoryX.length) return nodes;

  /**
   * Do these two categories sit at DIFFERENT places on the chart?
   *
   * A CAGR arrow and a difference arrow both say "between here and there", and
   * both are drawn from `categoryX`. Seven kinds — pie, doughnut, funnel,
   * sunburst, tilemap, treemap, waffle — have no category axis to anchor to and
   * publish a PLACEHOLDER instead: every category at the middle of the plot,
   * `categoryWidth` the whole plot. The decorations that only need somewhere to
   * point (callouts, bands) are fine with that; an arrow between two of them is
   * not, because "here" and "there" are the same point.
   *
   * What it produced was worse than meaningless. The difference arrow puts
   * itself at `categoryX[to] + categoryWidth[to] / 2 + 10`, which on a
   * placeholder anchor is the middle of the plot plus half the plot plus ten —
   * ten points past the RIGHT EDGE of the chart. On a 480x300 treemap, an
   * entirely ordinary size, it drew a zero-length arrow at x=488 with its label
   * clipped to `-5…`: a claim about growth, drawn off the chart, in a chart that
   * has no such comparison to make.
   *
   * So the arrow is not drawn — the same answer every reservation in this engine
   * gives when it cannot be met. It also covers `{ from: 1, to: 1 }` on a
   * perfectly ordinary column chart, which is the same degenerate arrow reached
   * by a different route.
   *
   * ONE ARROW LEGITIMATELY HAS BOTH ENDS ON ONE CATEGORY, and the test suite
   * said so before this shipped: a difference arrow anchored at a VALUE LINE
   * (`fromValueLine`) measures a single column against that line, so its two
   * ends are the same x and different values. That is a real comparison drawn on
   * a real anchor, and it is why the difference arrow asks a slightly different
   * question from the CAGR arrow below rather than sharing one predicate.
   */
  const anchorsDiffer = (from: number, to: number) => a.categoryX[from] !== a.categoryX[to];

  // --- CAGR arrow: diagonal arrow between two column tops with "+x.x% p.a." ---
  const cagrPair = decor.cagr ? clampPair(decor.cagr, a) : null;
  if (decor.cagr && cagrPair && anchorsDiffer(cagrPair.from, cagrPair.to)) {
    const { from, to } = cagrPair;
    // Per-series CAGR when requested (think-cell computes on totals by default).
    const si = decor.cagr.series;
    const sVals = si != null ? cfg.data.series[si]?.values : undefined;
    const vFrom = sVals ? (sVals[from] ?? 0) : a.columnValue[from];
    const vTo = sVals ? (sVals[to] ?? 0) : a.columnValue[to];
    const rate = cagr(vFrom, vTo, to - from);
    // Where each end is anchored, and whether it stands on a measurement — see
    // `cagrEnd`, which carries the reasons for both.
    const fromEnd = cagrEnd(cfg, a, from, si);
    const toEnd = cagrEnd(cfg, a, to, si);
    /** How far the arrowhead reaches back from its tip — see the SVG renderer. */
    const ARROW = 5;
    // Clear the column totals row and difference arrows when shown — but never
    // by more than the chart above the columns can pay for. The lift is
    // legibility, not data: what the arrow claims is its two category anchors and
    // the DIFFERENCE between its ends, and shortening the lift moves both ends by
    // the same amount, so the slope and the anchors both survive. Letting it
    // stand put the arrowhead 4pt above the top of a 60pt-tall stacked column.
    const wantLift = fs * 1.6 + (decor.totals ? fs * 1.5 : 0) + (decor.difference ? fs * 1.2 : 0);
    // BOUNDED AT BOTH ENDS, and dropped when the two bounds cross.
    //
    // The lift is subtracted, so a NEGATIVE one pushes the arrow down onto the
    // columns — which is what a chart with no headroom needs. Floored at zero
    // the arrow sat on a column top that was itself within the arrowhead's reach
    // of the frame, and the head was drawn 1.4pt above the chart; a footnote is
    // enough to squeeze a 300x60 one that far. Moving BOTH ends by the same
    // amount is what this lift already does, and the comment above says why that
    // is safe: the slope and the two anchors survive it.
    //
    // The floor that replaces zero is the FOOT of the chart, because a negative
    // lift can only push the head down. Where the two bounds cross there is no
    // height at which this arrow fits — a small-multiples panel a few points
    // tall — and it is not drawn, like every other piece of chrome that cannot
    // be paid for.
    const highest = Math.min(fromEnd.y, toEnd.y);
    const lowest = Math.max(fromEnd.y, toEnd.y);
    const liftMax = highest - ARROW * 1.8;
    const liftMin = lowest + ARROW * 1.8 - cfg.height;
    // A CONDITION ON THE BLOCK, never an early return. `decorationNodes` builds
    // one list for every decoration, so returning here would silently take the
    // difference arrow, the callouts and the bands with it — the same trap the
    // pie's slice loop records, one file over.
    const arrowFits = fromEnd.measured && toEnd.measured && liftMin <= liftMax;
    const lift = Math.max(liftMin, Math.min(wantLift, liftMax));
    const x1 = a.categoryX[from];
    const y1 = fromEnd.y - lift;
    const x2 = a.categoryX[to];
    const y2 = toEnd.y - lift;
    const angle = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
    if (arrowFits)
      nodes.push(
        { kind: "line", x1, y1, x2, y2, stroke: style.text, strokeWidth: 1.25, name: "cagr-line" },
        { kind: "arrowhead", x: x2, y: y2, angle, size: ARROW, fill: style.text, name: "cagr-head" },
      );
    // The arrow is lifted clear of the column tops and the caption sits above the
    // arrow, so on a frame with no headroom the caption meets the TITLE. Its y is
    // decorative — it captions the arrow, where the arrow's endpoints are anchored
    // data — so it may be shrunk and dropped where they may not.
    //
    // CLAMPING it to the title's bottom was tried and measured and is still
    // refused: it turned five `title x cagr-label` overlaps into EIGHT against
    // the column totals, because a clamp moves a label whether or not the
    // destination is free. What that measurement did not try is making the
    // caption SMALLER, and it was the last label in the engine still drawn at
    // the full chart font in a FIXED 90pt box — wider than an 80pt chart.
    //
    // So: fit it to the gap between the title and the arrow, and past the floor
    // drop the caption while KEEPING the arrow. The arrow still shows the growth
    // and its two anchors; a rate drawn through the title is not readable
    // anyway. Same answer the ring, radar, tilemap and pie reservations give.
    const capText = rate == null ? "CAGR n/a" : `${formatPercent(rate, 1, true, cfg.numberFormat?.locale)} p.a.`;
    // Measured against the title's INK, not against `titleHeight`. The reserved
    // band is taller than the text in it, and the caption has always been free to
    // sit inside that band — bounding by the reservation dropped the caption on
    // 120x90 and 160x120 charts where nothing was overlapping at all.
    const tNode = titleNode(cfg, style) as { fontSize: number; text: string } | null;
    // A valign-top title's ink runs from its baseline (y + fontSize) down by the
    // descender, so its bottom is `fontSize * 1.21`.
    const titleInkBottom = tNode ? tNode.fontSize * 1.21 : 0;
    const titleInkRight = tNode ? textWidth(tNode.text, tNode.fontSize, true) : 0;
    // The caption is valign-bottom in a `cf * 1.4` box sitting `cf * 1.6` above
    // the arrow, so its ink top is `min(y1, y2) - cf * 1.25`. Clearing the title
    // means keeping that at or below the title's ink bottom.
    const capFits = (f: number) => Math.min(y1, y2) - f * 1.25 >= titleInkBottom;
    let cf = fs;
    while (cf > MIN_LABEL_FS && !capFits(cf)) cf -= 0.5;
    const capW0 = Math.min(90, cfg.width);
    const capX0 = Math.max(0, Math.min((x1 + x2) / 2 - capW0 / 2, cfg.width - capW0));
    // A caption sitting entirely to the RIGHT of the title's ink cannot collide
    // with it however tall it is, so it keeps its full size.
    const clearOfTitleX = capX0 + (capW0 - textWidth(capText, fs, true)) / 2 >= titleInkRight;
    if (clearOfTitleX) cf = fs;
    // The caption goes with the arrow: it captions something that is not there
    // otherwise.
    if (arrowFits && (clearOfTitleX || capFits(cf))) {
      // Bounded by the frame, not a fixed 90pt box, and re-centred on the arrow
      // at whatever width it actually gets.
      const capW = capW0;
      nodes.push({
        kind: "text",
        x: capX0,
        // The box and the font move together, so a caption that needed no
        // shrinking lands exactly where it always did.
        y: Math.max(0, Math.min(y1, y2) - cf * 1.6),
        w: capW,
        h: cf * 1.4,
        text: capText,
        fontSize: cf,
        bold: true,
        color: style.text,
        align: "center",
        valign: "bottom",
        name: "cagr-label",
      });
    }
  }

  // --- Difference arrow: dashed level line + vertical arrow ---
  // Total (column totals) by default; a level difference arrow when `series`
  // names a series — it compares the cumulative stack level at that series.
  const diffPair = decor.difference ? clampPair(decor.difference, a) : null;
  /** A value-line anchor gives the arrow a second END without a second category. */
  const spansValueLine =
    decor.difference?.fromValueLine != null &&
    !!a.valueToY &&
    !!(decor.valueLines ?? (decor.valueLine ? [decor.valueLine] : []))[decor.difference.fromValueLine];
  const diffSi = decor.difference?.series;
  /** Does the layout publish a level for the named series? Else the arrow compares totals. */
  const diffByLevel = diffPair != null && levelOf(a, diffPair.from, diffSi) !== undefined;
  // Both ends must stand on a MEASUREMENT. The arrow reads `columnValue` (or a
  // series level), which every layout builds with `?? 0`, so a blank end was
  // drawn at the baseline and printed as a measured zero — "-100%" for
  // `[50, 60, null]`, "+50" for `[null, 60, 50]`, indistinguishable from a typed
  // 0. See `categoryMeasured` and `levelEndMeasured`. An arrow anchored at a value
  // line has only one category end to check: its FROM end is the line, which is a
  // number the author typed or the mean of what was measured, not a column that
  // may be blank.
  const diffMeasured =
    diffPair != null &&
    (spansValueLine || levelEndMeasured(cfg, a, diffPair.from, diffSi)) &&
    levelEndMeasured(cfg, a, diffPair.to, diffSi);
  if (decor.difference && diffPair && diffMeasured && (anchorsDiffer(diffPair.from, diffPair.to) || spansValueLine)) {
    const { from, to } = diffPair;
    // The pixel anchor and the number the arrow prints have to describe the same
    // mark, so take BOTH from `columnValue` wherever the layout publishes a
    // value→y map. `columnTop` is the drawn top, and the two part company more
    // often than they look:
    //
    //  - a multi-series LINE chart publishes the topmost point as `columnTop`
    //    and the FIRST series' value as `columnValue`, so the arrow spanned one
    //    series' points and read the other's growth (+100% over a span that is
    //    +200%);
    //  - a STACKED column with a negative segment puts `columnTop` at the
    //    positive total and `columnValue` at the net, so +10/−4 → +10/−8 drew a
    //    ZERO-LENGTH arrow labelled −67%, and +10/−8 → +5/0 drew one pointing
    //    DOWN labelled +150%.
    //
    // Anchoring on the printed value fixes both without touching what any
    // layout publishes — and it is an identity wherever the two already agree,
    // which is every chart in the showcase.
    const anchorY = (c: number) => (a.valueToY ? a.valueToY(a.columnValue[c]) : a.columnTop[c]);
    let vFrom = diffByLevel ? levelOf(a, from, diffSi)! : a.columnValue[from];
    let yFrom = diffByLevel ? a.valueToY!(vFrom) : anchorY(from);
    const vTo = diffByLevel ? levelOf(a, to, diffSi)! : a.columnValue[to];
    const yTo = diffByLevel ? a.valueToY!(vTo) : anchorY(to);
    // Anchor the arrow's start at a value line instead of a column.
    const vlIdx = decor.difference.fromValueLine;
    const vls = decor.valueLines ?? (decor.valueLine ? [decor.valueLine] : []);
    if (vlIdx != null && vls[vlIdx] && a.valueToY) {
      const vl = vls[vlIdx];
      // The SAME basis the value line itself is drawn on — see `meanOfMeasured`.
      // These were two computations and they disagreed on any blank category.
      vFrom = vl.mode === "mean" ? meanOfMeasured(a) : vl.value;
      yFrom = a.valueToY(vFrom);
    }
    const x = a.categoryX[to] + a.categoryWidth[to] / 2 + 10;
    nodes.push(
      {
        kind: "line",
        x1: a.categoryX[from] + a.categoryWidth[from] / 2 + 2,
        y1: yFrom,
        x2: x + 4,
        y2: yFrom,
        stroke: style.mutedText,
        strokeWidth: 0.75,
        dash: [2, 2],
        name: "diff-level",
      },
      { kind: "line", x1: x, y1: yFrom, x2: x, y2: yTo, stroke: style.text, strokeWidth: 1.25, name: "diff-line" },
      { kind: "arrowhead", x, y: yTo, angle: yTo < yFrom ? -90 : 90, size: 5, fill: style.text, name: "diff-head" },
    );
    const usePct = decor.difference.percent ?? true;
    /**
     * A PERCENTAGE NEEDS A POSITIVE BASE, and `vFrom !== 0` was not enough.
     *
     * `vTo / vFrom - 1` inverts its sign when `vFrom` is negative, so the
     * caption contradicted the arrow drawn two lines above it — and the arrow
     * was the half that was right:
     *
     *   -100 ->  -50   a loss HALVED     arrow UP     label "-50%"
     *   -100 -> -200   a loss DOUBLED    arrow DOWN   label "+100%"
     *    -50 ->   25   into profit       arrow UP     label "-150%"
     *
     * And there is no percentage to rescue: growth from a negative base is not
     * a meaningful ratio at all — a swing from -50 to 25 is no more "+150%"
     * than it is "-150%" — which is why `cagr()` already returns null on a
     * non-positive base and the chart prints "CAGR n/a".
     *
     * So the guard that already caught zero now catches negatives too, for the
     * same reason, and falls back to the ABSOLUTE difference: true on every
     * base, and it still says which way the number went.
     */
    // RESOLVED against the two ends it measures, not handed the raw
    // `numberFormat` — see the note on the value-line label below, which had the
    // identical defect. `numberFormat.decimals` defaults to "auto", and
    // `formatNumber`'s own auto ladder stops at two decimals under 1, so an
    // arrow between columns the chart labels 0.0040 and 0.0060 captioned itself
    // "+0.01". `chart.ts` resolves the target-gap label over the same two ends.
    const diffFmt = finerFormat([vFrom, vTo], vTo - vFrom, cfg.numberFormat);
    const label =
      usePct && vFrom > 0
        ? formatPercent(vTo / vFrom - 1, 0, true, cfg.numberFormat?.locale)
        : formatNumber(vTo - vFrom, { ...diffFmt, forceSign: true });
    // The label reads to the RIGHT of the arrow, in a margin `computeFrame`
    // reserves for it — but only the cartesian frame reserves one. A line chart
    // puts its last category hard against the plot edge, so on an ordinary
    // 400x300 two-point line the label started 8pt past the chart and the
    // frame-clip at the end of `buildChart` cut "+100%" down to "+…". Flip it to
    // the other side of the arrow when the right has no room: the ARROW is what
    // anchors the claim, and the caption says the same thing from either side.
    const labelW = textWidth(label, fs, true) + 2;
    const flip = x + 3 + labelW > cfg.width;
    nodes.push({
      kind: "text",
      x: flip ? Math.max(0, x - 3 - labelW) : x + 3,
      y: (yFrom + yTo) / 2 - fs * 0.75,
      w: flip ? labelW : Math.max(30, cfg.width - x - 3),
      h: fs * 1.5,
      text: label,
      fontSize: fs,
      bold: true,
      color: style.text,
      align: "left",
      valign: "middle",
      name: "diff-label",
    });
  }

  // --- Value lines: dashed horizontals at fixed values or the mean of totals ---
  const valueLines = decor.valueLines ?? (decor.valueLine ? [decor.valueLine] : []);
  if (valueLines.length && a.valueToY) {
    valueLines.forEach((vl, i) => {
      // The mean of what was MEASURED — see `meanOfMeasured` at the head of
      // this file, which the difference arrow's anchor shares.
      const value = vl.mode === "mean" ? meanOfMeasured(a) : vl.value;
      const y = a.valueToY!(value);
      /**
       * THE CHART'S OWN PRECISION, not `formatNumber`'s per-value fallback.
       *
       * `cfg.numberFormat` was handed straight to `formatNumber`, whose "auto"
       * ladder stops at two decimals below 1 — while `resolveFormat`, which
       * every data label in the engine goes through, widens to six so a chart of
       * rates, yields or defect fractions keeps its digits. The two ladders
       * disagreeing put a WRONG NUMBER in ink, not merely a coarse one:
       *
       *     bars 0.004 / 0.005 / 0.006   labelled 0.0040 / 0.0050 / 0.0060
       *     mean line drawn at 0.005     labelled "Ø 0.01"
       *
       * — twice the mean, above every bar on the chart, on a line whose y is
       * correct. On any chart where the two ladders already agree this is the
       * string it always was.
       */
      const vlFmt = finerFormat(a.columnValue, value, cfg.numberFormat);
      nodes.push(
        {
          kind: "line",
          x1: a.plot.x,
          y1: y,
          x2: a.plot.x + a.plot.w,
          y2: y,
          stroke: style.mutedText,
          strokeWidth: 1,
          dash: [3, 2],
          name: `value-line-${i}`,
        },
        {
          kind: "text",
          x: a.plot.x + 2,
          y: y - fs * 1.5,
          w: 100,
          h: fs * 1.4,
          text: (vl.mode === "mean" ? "Ø " : "") + formatNumber(value, vlFmt),
          fontSize: fs * 0.95,
          color: style.mutedText,
          align: "left",
          valign: "bottom",
          name: `value-line-label-${i}`,
        },
      );
    });
  }

  // --- Speech-bubble callouts: a comment anchored to a column or level ---
  decor.callouts?.forEach((co, i) => {
    const c = categoryAt(co.category, a.categoryX.length);
    // A callout about no category has nowhere to stand (its text would be drawn
    // from `categoryX[NaN]`, every coordinate NaN and every node dropped).
    if (c == null) return;
    const ax = a.categoryX[c];
    // A callout is a comment, not a claim, so a series with no mark here (NaN —
    // a line on its own axis) keeps its text and points at the column instead of
    // being withheld the way a level ARROW is: the sentence is the author's, and
    // the category is still the thing it is about.
    const lv = levelOf(a, c, co.series);
    const ay = lv !== undefined && Number.isFinite(lv) ? a.valueToY!(lv) : a.columnTop[c];
    /**
     * FITTED TO THE CHART, then dropped — a callout's text is the one string on
     * a chart that nothing else bounds.
     *
     * Its box is `textWidth(text) + fs * 1.2` wide, and a callout is a sentence
     * somebody typed: "A rather long callout label" at the default font is 157
     * points, which is wider than a 120-point chart. The clamp below was written
     * for the vertical case and answers the CENTRE of the frame when the box
     * cannot fit between its own bounds — so an oversized bubble was centred and
     * hung off BOTH edges, 18.9 points each side on twelve kinds at 120x90.
     * Nothing was clipped, because a box is not text: PowerPoint drew the border
     * and the sentence straight across whatever the chart was sitting beside.
     *
     * The shrink-then-drop every other label in this engine gets. Below
     * `MIN_LABEL_FS` the callout is not drawn at all — a bubble at four points
     * is not a comment anyone can read, and the tail would point at a chart the
     * reader cannot see the note for.
     */
    let cf = fs;
    while (cf > MIN_LABEL_FS && textWidth(co.text, cf) + cf * 1.2 > cfg.width) cf -= 0.5;
    if (textWidth(co.text, cf) + cf * 1.2 > cfg.width) return;
    const w = textWidth(co.text, cf) + cf * 1.2;
    const h = cf * 1.9;
    // Bubble center defaults to hovering above the anchor — but "above" is only
    // available when the anchor has room above it. A 100% or Mekko chart exposes
    // no valueToY, so EVERY callout falls back to columnTop, which on those kinds
    // is the plot ceiling for every column by construction; the same happens on
    // any chart when the callout names the tallest column. Lifting 4.2 font sizes
    // from there put the box, its text and its tail entirely off the top of the
    // canvas. Keep the bubble on the canvas — a bubble overlapping its column
    // still reads; one at y = -30 is simply lost.
    const clamp = (v: number, lo: number, hi: number) => (lo > hi ? (lo + hi) / 2 : Math.min(hi, Math.max(lo, v)));
    const bx = clamp(ax + (co.dx ?? 0), w / 2, cfg.width - w / 2);
    /**
     * The ceiling is the TITLE's ink, not the top of the canvas.
     *
     * The clamp above keeps the bubble on the chart, and y=0 is where the title
     * is — the same mistake the column totals made before `aboveMarkFontSize`,
     * and the one the CAGR caption's floor was reverted for. It bites on exactly
     * the kinds this decoration is hardest on: a 100% chart, a mekko, a treemap
     * or a sunburst publishes no `valueToY`, so every callout anchors on the
     * plot ceiling, lifts 4.2 font sizes from there, and lands on the chart's
     * own name. Six kinds at 200x150, two of them at 480x300 as well.
     *
     * A bubble pushed DOWN is still a bubble with a tail pointing at its anchor;
     * a title with a sentence across it is unreadable and takes the chart's name
     * with it. Where the frame cannot hold the bubble under the title at all the
     * clamp's own `lo > hi` arm centres it, which is the pre-existing behaviour
     * for a bubble taller than its chart.
     */
    const by = clamp(ay - cf * 4.2 + (co.dy ?? 0), titleInkBottom(cfg, style) + h / 2, cfg.height - h / 2);
    nodes.push(
      {
        kind: "line",
        x1: bx,
        y1: by + h / 2 - 1,
        x2: ax,
        y2: ay - 2,
        stroke: style.text,
        strokeWidth: 0.75,
        name: `callout-tail-${i}`,
      },
      {
        kind: "rect",
        x: bx - w / 2,
        y: by - h / 2,
        w,
        h,
        fill: style.background,
        stroke: style.text,
        strokeWidth: 1,
        name: `callout-box-${i}`,
      },
      {
        kind: "text",
        x: bx - w / 2,
        y: by - h / 2,
        w,
        h,
        text: co.text,
        fontSize: cf,
        color: style.text,
        align: "center",
        valign: "middle",
        name: `callout-text-${i}`,
      },
    );
  });

  return nodes;
}

/**
 * Shaded background bands highlighting an axis region — drawn BEHIND the
 * data (the caller prepends these to the scene). axis "y" spans a value
 * range; axis "x" spans category indices.
 */
export function bandNodes(cfg: ChartConfig, style: ChartStyle, decor: Decorations, a: LayoutAnchors): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;
  decor.bands?.forEach((band, i) => {
    const fill = band.color ?? "#f2f1ec";
    let r: { x: number; y: number; w: number; h: number } | null = null;
    // `axis` names the DATA axis the band spans, not a screen direction: a "y"
    // band is a range of VALUES and a "x" band is a range of CATEGORIES,
    // whichever way the chart is turned. Reading `valueToY` alone meant the
    // range zones of a bullet chart — conventionally horizontal — were dropped
    // silently along with its target tick.
    const H = !!cfg.horizontal;
    const valueToPos = H ? a.valueToX : a.valueToY;
    if (band.axis === "y" && valueToPos) {
      // Clip to the plot: band.from/to are data values that may fall outside the
      // value domain, and the map extrapolates past the axis, so an unclamped
      // band renders off-frame. (The category branch below clamps its indices;
      // a band entirely outside the plot collapses to w/h <= 0 and is dropped.)
      const v1 = valueToPos(band.from);
      const v2 = valueToPos(band.to);
      if (H) {
        const left = Math.max(a.plot.x, Math.min(v1, v2));
        const right = Math.min(a.plot.x + a.plot.w, Math.max(v1, v2));
        r = { x: left, y: a.plot.y, w: right - left, h: a.plot.h };
      } else {
        const top = Math.max(a.plot.y, Math.min(v1, v2));
        const bot = Math.min(a.plot.y + a.plot.h, Math.max(v1, v2));
        r = { x: a.plot.x, y: top, w: a.plot.w, h: bot - top };
      }
    } else if (band.axis === "x" && a.categoryX.length) {
      const from = categoryAt(band.from, a.categoryX.length);
      const to = categoryAt(band.to, a.categoryX.length);
      if (from == null || to == null) return;
      const c1 = Math.min(from, to);
      const c2 = Math.max(from, to);
      // `categoryX` is the category's centre on the CROSS axis — x on a column
      // chart, y on a bar.
      //
      // CLIPPED TO THE PLOT, as the value-range branch above is. The 0.75 is half
      // a SLOT only where `categoryWidth` is a bar two thirds of one (a column at
      // the default gap); everywhere else it is a guess, and a wrong one drew the
      // band and its label outside the frame — a mekko's columns touch and the
      // violin's width is its half-spread, so the band overshot each end; a line's
      // width is the whole slot, so it ran 1.5 slots wide; and a stacked chart at
      // `gapWidth: 0` reached 36pt past the left edge. Ink outside the frame is
      // drawn on the slide, over whatever the chart sits beside. The plot is the
      // bound a category can have, and a band over every category exactly fills it.
      const lo = H ? a.plot.y : a.plot.x;
      const hi = lo + (H ? a.plot.h : a.plot.w);
      const m1 = Math.max(lo, a.categoryX[c1] - a.categoryWidth[c1] * 0.75);
      const m2 = Math.min(hi, a.categoryX[c2] + a.categoryWidth[c2] * 0.75);
      r = H ? { x: a.plot.x, y: m1, w: a.plot.w, h: m2 - m1 } : { x: m1, y: a.plot.y, w: m2 - m1, h: a.plot.h };
    }
    if (!r || r.w <= 0 || r.h <= 0) return;
    nodes.push({ kind: "rect", ...r, fill, name: `band-${i}` });
    if (band.label) {
      /**
       * FITTED TO THE CHART, not to the band — and dropped when even that
       * cannot be met.
       *
       * A band's label was the last string in `decor.ts` bounded by nothing. It
       * is drawn at `fs * 0.9` in a box `fs * 1.3` tall from the top of the
       * band, and `fitPlot` grows a squeezed plot UP from its bottom edge, so on
       * a chart with no room the band starts near the foot of the frame and the
       * label is drawn below it: 6.2 points under a 300x60 bar chart at 18pt,
       * four kinds, both cramped frames.
       *
       * Bounded by the FRAME rather than by the band, deliberately. A band can
       * legitimately be a sliver — two percent of a value axis — and fitting the
       * label to that would drop the name of every thin band, where today it
       * merely spills a little over its own edge onto the data it names. The
       * frame is the bound that is not a matter of taste: ink outside it is
       * drawn on the slide, over whatever the chart is sitting beside.
       */
      const room = cfg.width - (r.x + 3);
      const below = cfg.height - (r.y + 1);
      let bf = Math.min(fs * 0.9, below / 1.3);
      while (bf > MIN_LABEL_FS && textWidth(band.label, bf) > room) bf -= 0.5;
      if (bf >= MIN_LABEL_FS && room > 0 && textWidth(band.label, bf) <= room) {
        nodes.push({
          kind: "text",
          x: r.x + 3,
          y: r.y + 1,
          // The box moves with the font, so a label that needed no shrinking
          // lands exactly where it always did.
          w: Math.max(20, r.w - 6),
          h: bf * (1.3 / 0.9),
          text: band.label,
          fontSize: bf,
          color: style.mutedText,
          align: "left",
          valign: "top",
          name: `band-label-${i}`,
        });
      }
    }
  });
  return nodes;
}

/**
 * The category a decoration's index names: ROUNDED to a whole one and clamped
 * into range, or null when it is no number at all.
 *
 * Every anchor array is read by index, and a fractional index reads
 * `undefined`. Where a layout publishes a value→y map that comes out as NaN and
 * `finiteNodes` drops the node; where it does not (a 100% chart, a mekko, a
 * treemap, a sunburst) the arrow's `y1` was read straight off `columnTop[0.5]`
 * and the node left the engine with `y1: undefined` — a key `finiteNodes` cannot
 * see, which the SVG printed as `y1="NaN"`. The pane's difference spinner is a
 * plain number input, so typing 1.5 reached it.
 *
 * Rounded, because the author pointed at a category and the nearest one is the
 * only reading there is. A number that is out of range still clamps (`to: 99`
 * means "the last"), and a numeric string reads as the number it spells; what is
 * not a number is no category, and the decoration is not drawn.
 */
function categoryAt(index: unknown, n: number): number | null {
  const v = Number(index);
  return Number.isNaN(v) ? null : Math.max(0, Math.min(n - 1, Math.round(v)));
}

function clampPair(p: { from: number; to: number }, a: LayoutAnchors): { from: number; to: number } | null {
  const n = a.categoryX.length;
  const from = categoryAt(p.from, n);
  const to = categoryAt(p.to, n);
  if (from == null || to == null) return null;
  return from <= to ? { from, to } : { from: to, to: from };
}
