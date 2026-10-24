import type { ChartConfig, ChartData, ChartStyle, Decorations, Series } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, parseDateToken, resolveFormat, segmentLabel } from "../format";
import { maxOf, minOf } from "../agg";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import { clipToWidth } from "../elements";
import {
  MIN_LABEL_FS,
  baselineNode,
  categorySlots,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  fitPlot,
  labelBesidePoint,
  markInFrame,
  footnoteH,
  pointLabelText,
  logFloor,
  titleHeight,
  titleNode,
  valueScale,
} from "./frame";
import { horizontalChrome, seriesLabelNodes, type LayoutResult } from "./column";
import { columnHasData } from "./totals";

/**
 * Slab count for a filled area/ribbon segment.
 *
 * The renderers have no polygon fill, so a segment is tiled with thin rects.
 * This was a flat 24 regardless of on-screen width, which put sub-pixel rects on
 * narrow charts and multiplied the Office.js shape count ~24× per segment (a
 * 30×5 stacked area is ~3,500 rects, right against the host's shape budget).
 * Scale to the segment's pixel span (~4px per slab, the eye can't resolve a
 * finer staircase) and cap at the old 24 so a wide segment never GAINS shapes;
 * a stepped (staircase) segment has a flat top and needs no tessellation.
 *
 * AND THE SPAN WAS ONLY HALF OF IT. A segment is tiled to approximate a SLOPING
 * edge, so what the tiling has to resolve is how far that edge MOVES — but the
 * count was taken from width alone. A perfectly level segment 80px wide got
 * twenty identical rectangles laid side by side, which is one rectangle drawn
 * twenty times.
 *
 * Measured against the real engine on 2026-08-30, taking the rise into account:
 *
 *     sample charts      400 slabs -> 304   (-24%)
 *     the shipped deck   652 slabs -> 412   (-37%)
 *     "Revenue mix (stacked area) + margin %"
 *                        300 shapes -> 128  (-57%)
 *
 * That last one matters beyond speed: round 150 recorded 400-500 shapes in one
 * context CRASHING PowerPoint on all seven attempts, and a chart we ship sits at
 * 300. Shapes are also the unit of cost — an in-place update runs about 767ms
 * PER SHAPE on this host, measured over 178 updates — so every slab removed is
 * paid back on insert, on edit and on every deck-wide rescale.
 */
function slabSteps(spanPx: number, stepped?: "before" | "after" | "center", risePx = Infinity): number {
  if (stepped) return stepped === "center" ? 2 : 1;
  const bySpan = Math.max(2, Math.min(24, Math.ceil(Math.abs(spanPx) / 4)));
  // A caller that cannot say how far its edge moves gets exactly the old answer.
  if (!Number.isFinite(risePx)) return bySpan;
  /**
   * The staircase error of a straight edge is `slope * slabWidth / 2`. Put
   * `steps = rise / 2` into that and it is `(rise/span) * (span/(rise/2)) / 2`
   * = **1pt**, whatever the span and whatever the slope. So the bound is
   * arithmetic rather than a taste call.
   *
   * WITH ONE HONEST CAVEAT: that holds while THIS is the binding rule. On a
   * steep edge `bySpan` caps first — 24 slabs, the historical ceiling — and the
   * error is then whatever the old rule already gave, because on those segments
   * this changes nothing at all. The 1pt claim is about the segments this rule
   * actually decides, not about every segment.
   */
  const byRise = Math.max(1, Math.ceil(Math.abs(risePx) / 2));
  // The span rule stays as the CEILING, so no segment can ever GAIN a slab.
  return Math.min(bySpan, byRise);
}

// Band low/high rows shade an uncertainty ribbon instead of drawing lines.
const BAND_LOW = /^band\s*low$/i;
const BAND_HIGH = /^band\s*high$/i;
const isBandRow = (s: Series) => BAND_LOW.test(s.name.trim()) || BAND_HIGH.test(s.name.trim());

/**
 * Split a `Band low` / `Band high` pair out of the series list.
 *
 * The rows are documented as never drawing as lines — they are the shaded
 * confidence ribbon — so the DRAWN data is the rest of the list. Shared by the
 * vertical and horizontal layouts because the horizontal one did not do it at
 * all: sideways, a chart with band rows drew them as two ordinary lines with
 * their own legend chips and no ribbon, which is neither what the reference
 * documents nor what the same config renders upright.
 *
 * Returns the input untouched when there is no band row, so the ordinary chart
 * is byte-identical.
 */
function splitBandRows(raw: ChartData): {
  data: ChartData;
  bandLow?: (number | null)[];
  bandHigh?: (number | null)[];
} {
  const bandLow = raw.series.find((s) => BAND_LOW.test(s.name.trim()))?.values;
  const bandHigh = raw.series.find((s) => BAND_HIGH.test(s.name.trim()))?.values;
  if (!bandLow && !bandHigh) return { data: raw };
  return { data: { ...raw, series: raw.series.filter((s) => !isBandRow(s)) }, bandLow, bandHigh };
}

/**
 * WHICH CATEGORIES `columnValue` IS A MEASUREMENT FOR, when it is the first
 * series' value — the question every `?? 0` in the anchors below cannot answer.
 *
 * `decorations.valueLines: [{ mode: "mean" }]` averages `columnValue`, and a
 * blank category arrived there as a zero AND as a denominator: a year-to-date
 * line with six actual months and six blank ones printed "Ø 55" where the
 * clustered column over the same data printed "Ø 110". `layoutColumns` has
 * published the flag since `35840f6`; its comment said "the column family is the
 * one that offers `valueLines`", which was never true — `buildChart` draws them
 * on every kind with a value map, the pane offers the control on every kind, and
 * the shipped "KPI trend (line)" template switches the mean on. So every
 * `columnValue` in this file is paired with the flag, not just the one that was
 * reported.
 */
const firstSeriesHasData = (data: ChartData): boolean[] =>
  data.categories.map((_, c) => columnHasData(data.series.slice(0, 1), c));

/**
 * Levels for series that each stand on the baseline: a series' level is its own
 * value, because nothing stacks on anything. What a level-qualified decoration
 * (`difference.series`, `callouts[].series`, `cagr.series`) anchors on — see
 * `LayoutAnchors.seriesLevels`. Indexed by the CONFIG's series list, band rows
 * included, because that is the list the decoration's index was written against.
 */
const ownValueLevels = (data: ChartData): number[][] =>
  data.categories.map((_, c) => data.series.map((s) => s.values[c] ?? 0));

/**
 * The levels `layoutLine` starts from, for the decorations that name a series
 * (`difference.series`, `callouts[].series`, `cagr.series`).
 *
 * THIS LAYOUT PUBLISHED NONE, so on a line chart `series: 1` was ignored and the
 * arrow compared series 0 — "+100%" for a flat series — and on an area chart
 * every series below the top one was given the stack total. Lines stand on the
 * baseline, so a line's level is its own value; an area stacks, so its level is
 * the running sum, which the area loop fills in as it stacks. NaN until then, and
 * for a row that draws no mark at all (a band row on an area chart) — see
 * `levelOf` in `decor.ts`.
 */
const lineLevels = (cfg: ChartConfig, area: boolean): number[][] =>
  area ? cfg.data.categories.map(() => cfg.data.series.map(() => NaN)) : ownValueLevels(cfg.data);

/**
 * `columnHasData` and `seriesLevels` for `layoutLine`'s anchors. An area's value
 * is the stack's sum, which exists wherever ANY series has a cell; a line's is the
 * first series' own, which exists only where that one does. Each answers for the
 * number it sits beside.
 */
const measuredAnchors = (data: ChartData, area: boolean, seriesLevels: number[][]) => ({
  columnHasData: area ? data.categories.map((_, c) => columnHasData(data.series, c)) : firstSeriesHasData(data),
  seriesLevels,
});

/**
 * Where date categories sit along the axis, as a 0..1 fraction of its span — or
 * null, which means "evenly spaced, in the order given".
 *
 * Date spacing is only a time axis while the dates RUN FORWARD. Every other
 * thing a line chart does with its categories is indexed in the order they were
 * written — the line joins point `c` to point `c + 1`, the forecast divider sits
 * between `forecastFrom - 1` and `forecastFrom`, a ribbon slab marches from one
 * centre to the next — so a list whose days do not strictly increase drew its
 * line backwards across the plot and gave the ribbons a negative width.
 *
 * Measured 2026-10-24 on `Oct 23, Nov 23, Dec 23, Jan 24, Feb 24`: the parser
 * reads "Jan 24" as 24 January in V8's default year, so the new year's two
 * months landed at x 45 and 79 and the three before them at 343 to 410 — the
 * line ran 413 -> 48 across the whole plot between December and January.
 * `format.ts` fixed that for "Jan-24" only, and its comment is explicit that the
 * space form is left to `Date.parse` because it is ambiguous with a day. Labels
 * with no year at all are not ambiguous and break the same way: Excel's own
 * d-mmm display ("28-Dec", "4-Jan") and "28 Dec" / "4 Jan" put the December
 * points of a weekly series into 14pt at the right edge. Out-of-order dates
 * zigzag, and a duplicate date puts two points on one x.
 *
 * So the test is on the PARSED DAYS rather than on the spelling: this layout
 * cannot tell a wrapped year from a mistyped one, and does not need to. Where the
 * days do not strictly increase it falls back to the even spacing every other
 * category list gets, which keeps the categories in the order the author gave
 * them. A newest-first list is in that group on purpose — it used to be laid out
 * oldest-first with its line drawn right to left, a reordering nothing else in
 * the chart agreed with.
 */
function dateSpacing(categories: string[]): number[] | null {
  if (categories.length < 2) return null;
  const days = categories.map((c) => parseDateToken(c));
  if (!days.every((d): d is number => d != null)) return null;
  for (let c = 1; c < days.length; c++) if ((days[c] as number) <= (days[c - 1] as number)) return null;
  const d0 = days[0] as number;
  const d1 = days[days.length - 1] as number;
  return days.map((d) => ((d as number) - d0) / (d1 - d0));
}

/** Line and area charts over categories. Lines are 2pt with ≥3pt markers. */
export function layoutLine(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  if ((cfg.kind === "line" || cfg.kind === "area") && decor.sparkline) {
    return layoutSparkline(cfg, style, decor);
  }
  if (cfg.kind === "line" && decor.slope && cfg.data.categories.length >= 2) {
    return layoutSlope(cfg, style, decor);
  }
  if (cfg.kind === "line" && decor.bump && cfg.data.categories.length >= 2) {
    return layoutBump(cfg, style, decor);
  }
  if ((cfg.kind === "line" || cfg.kind === "area") && cfg.horizontal) {
    return layoutLineHorizontal(cfg, style, decor);
  }
  const { data, bandLow, bandHigh } = splitBandRows(cfg.data);
  // The config as the plot sees it: band rows are not drawn, so anything that
  // indexes series positionally must index THIS list, not the raw one.
  const drawn: ChartConfig = data === cfg.data ? cfg : { ...cfg, data };
  const n = data.categories.length;
  const area = cfg.kind === "area";
  const fs = style.fontSize;

  const all = [
    ...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)),
    ...(bandLow ?? []).filter((v): v is number => v != null),
    ...(bandHigh ?? []).filter((v): v is number => v != null),
  ];
  // Area charts stack (positives above zero, negatives below); lines share
  // one scale. Negative areas dip under the baseline — think-cell parity for
  // P&L-over-time where a series can go negative.
  const stackedPos = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
  const stackedNeg = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
  const dataMax = area ? maxOf(stackedPos, 0) : maxOf(all, 0);
  const dataMin = area ? minOf(stackedNeg, 0) : minOf(all, 0);
  const fmt = resolveFormat(all, cfg.numberFormat);

  const { frame } = computeFrame(cfg, style, decor, decor.seriesLabels ? data.series.map((s) => s.name) : []);
  const slots = categorySlots(frame, n, 0);
  // Date categories ("2024-03", "Jan 2025", …) space x proportionally to time —
  // while they run forward; see `dateSpacing`.
  const dateAt = area ? null : dateSpacing(data.categories);
  if (dateAt) {
    const inset = slots.slotWidth / 2;
    for (let c = 0; c < n; c++) slots.centers[c] = frame.x + inset + dateAt[c] * (frame.w - inset * 2);
  }
  const logOn = !area && !!cfg.logScale;
  const scale = valueScale(frame, logOn ? logFloor(all, dataMin) : dataMin, dataMax, cfg.scale, undefined, logOn);
  const y0 = scale.toY(0);
  /**
   * A value a log axis has no place for, read as a gap.
   *
   * `valueScale`'s log `toY` clamps with `Math.max(v, min)`, which is right for
   * a bar — it starts at the floor, so a value under the axis draws nothing —
   * and a LIE for a point: zero and every negative landed exactly on the bottom
   * decade line, drawn as a marker at `10` and joined to its neighbours by a
   * line, indistinguishable from a genuine 10. A series that starts at zero is
   * an ordinary thing to log-scale, and the chart said it started at the axis
   * minimum instead.
   *
   * A gap rather than a clamp, because that is what this layout already does
   * with a value it cannot place (`null`), and a break in the line is visible
   * where a point on the floor is not. `bridgeGaps` still bridges it, which is
   * the caller saying they would rather have the connection.
   */
  const plottable = (v: number | null | undefined): number | null => (v == null || (logOn && v <= 0) ? null : v);

  const nodes: SceneNode[] = chromeNodes(cfg, style, decor, frame, slots.centers, scale);
  const lastSegMid: (number | null)[] = data.series.map(() => null);
  const columnTop: number[] = data.categories.map(() => y0);
  // Where each series tops out, for the decorations that name one — see
  // `lineLevels`. An area fills its in below, as it stacks.
  const seriesLevels = lineLevels(cfg, area);

  /** Shaded ribbon between two per-category value arrays (slab technique). */
  const ribbon = (lows: (number | null)[], highs: (number | null)[], fill: string, name: string) => {
    for (let c = 0; c < n - 1; c++) {
      const l0 = lows[c];
      const l1 = lows[c + 1];
      const h0 = highs[c];
      const h1 = highs[c + 1];
      if (l0 == null || l1 == null || h0 == null || h1 == null) continue;
      const span = slots.centers[c + 1] - slots.centers[c];
      // In PIXELS, not values: `scale.toY` is what the rule is bounded in, and a
      // value delta says nothing about how far the edge actually travels.
      const steps = slabSteps(
        span,
        undefined,
        Math.max(Math.abs(scale.toY(h1) - scale.toY(h0)), Math.abs(scale.toY(l1) - scale.toY(l0))),
      );
      const w = span / steps;
      for (let k = 0; k < steps; k++) {
        const t = (k + 0.5) / steps;
        const yT = scale.toY(h0 + (h1 - h0) * t);
        const yB = scale.toY(l0 + (l1 - l0) * t);
        nodes.push({
          kind: "rect",
          x: slots.centers[c] + k * w,
          y: Math.min(yT, yB),
          w: w + 0.5,
          h: Math.abs(yB - yT),
          fill,
          name: `${name}-${c}-${k}`,
        });
      }
    }
  };
  // Confidence/uncertainty band from Band low / Band high rows.
  if (!area && bandLow && bandHigh) {
    ribbon(
      bandLow,
      bandHigh,
      lerpColor(style.background, seriesColor(style, 0, data.series[0]?.color), 0.18),
      "band-ribbon",
    );
  }
  // Filled gap between two named series (plan-vs-actual ribbon).
  if (!area && decor.fillBetween) {
    const [ai, bi] = decor.fillBetween;
    const sa = data.series[ai]?.values;
    const sb = data.series[bi]?.values;
    if (sa && sb)
      ribbon(sa, sb, lerpColor(style.background, seriesColor(style, ai, data.series[ai]?.color), 0.22), "fill-between");
  }

  if (area) {
    // Stacked areas drawn as per-category slabs (renderers have no polygon fill),
    // bottom-up so later series sit on top visually. Positives stack above the
    // zero baseline and negatives below it, so a negative series dips under zero.
    const posBase = data.categories.map(() => 0);
    const negBase = data.categories.map(() => 0);
    data.series.forEach((s, si) => {
      const fill = seriesColor(style, si, s.color);
      // This series' band boundaries per category (value units).
      const lower: number[] = [];
      const upper: number[] = [];
      for (let c = 0; c < n; c++) {
        const v = s.values[c] ?? 0;
        if (v >= 0) {
          lower[c] = posBase[c];
          upper[c] = posBase[c] + v;
          posBase[c] += v;
        } else {
          upper[c] = negBase[c];
          lower[c] = negBase[c] + v;
          negBase[c] += v;
        }
        // The stack's reach AFTER this series — the same sum the column layout
        // publishes, so a level arrow on an area chart reads like one on a stack.
        seriesLevels[c][cfg.data.series.indexOf(s)] = posBase[c] + negBase[c];
      }
      for (let c = 0; c < n - 1; c++) {
        // Approximate the trapezoid with a rect at the average height.
        const yTop0 = scale.toY(upper[c]);
        const yTop1 = scale.toY(upper[c + 1]);
        const yBot0 = scale.toY(lower[c]);
        const yBot1 = scale.toY(lower[c + 1]);
        const span = slots.centers[c + 1] - slots.centers[c];
        // BOTH edges: a band whose top is level but whose bottom climbs still
        // needs tiling, so the tessellation is sized by whichever moves more.
        const steps = slabSteps(span, decor.stepped, Math.max(Math.abs(yTop1 - yTop0), Math.abs(yBot1 - yBot0)));
        const w = span / steps;
        for (let k = 0; k < steps; k++) {
          const t = (k + 0.5) / steps;
          // Stepped areas hold a flat top across the interval (staircase);
          // "after" carries the left value, "before" the right, "center" both.
          let yT: number;
          let yB: number;
          if (decor.stepped === "after" || (decor.stepped === "center" && t < 0.5)) {
            yT = yTop0;
            yB = yBot0;
          } else if (decor.stepped === "before" || decor.stepped === "center") {
            yT = yTop1;
            yB = yBot1;
          } else {
            yT = yTop0 + (yTop1 - yTop0) * t;
            yB = yBot0 + (yBot1 - yBot0) * t;
          }
          nodes.push({
            kind: "rect",
            x: slots.centers[c] + k * w,
            y: Math.min(yT, yB),
            w: w + 0.5,
            h: Math.abs(yB - yT),
            fill,
            name: `area-${si}-${c}-${k}`,
          });
        }
      }
      lastSegMid[si] = n > 0 ? scale.toY((lower[n - 1] + upper[n - 1]) / 2) : null;
    });
    for (let c = 0; c < n; c++) columnTop[c] = scale.toY(posBase[c]);
  } else {
    // Forecast boundary: categories from this index on draw dashed with
    // hollow markers; a subtle divider marks the actuals/forecast split.
    const fc = decor.forecastFrom;
    if (fc != null && fc > 0 && fc < n) {
      const dx = slots.centers[fc - 1] + (slots.centers[fc] - slots.centers[fc - 1]) / 2;
      nodes.push({
        kind: "line",
        x1: dx,
        y1: frame.y,
        x2: dx,
        y2: frame.y + frame.h,
        stroke: style.gridline,
        strokeWidth: 1,
        dash: [2, 3],
        name: "forecast-divider",
      });
    }
    // Smooth (Catmull-Rom) curves, sampled to a dense polyline. Ignored when
    // stepped is set (mutually exclusive shapes).
    const smooth = !!decor.smooth && !decor.stepped;
    data.series.forEach((s, si) => {
      const color = seriesColor(style, si, s.color);
      if (smooth) {
        // The sampler lives at the top of this file now, because the horizontal
        // layout needs the same curve from points built the other way round.
        const pts = s.values.map((v, c) => {
          const pv = plottable(v);
          return pv == null ? null : { x: slots.centers[c], y: scale.toY(pv), c };
        });
        for (const seg of splineSegments(pts)) {
          nodes.push({
            kind: "line",
            x1: seg.x1,
            y1: seg.y1,
            x2: seg.x2,
            y2: seg.y2,
            stroke: color,
            strokeWidth: 2,
            ...(fc != null && seg.c >= fc ? { dash: [4, 3] } : {}),
            name: `line-${si}-${seg.c}-s${seg.k}`,
          });
        }
      }
      let prev: { x: number; y: number } | null = null;
      for (let c = 0; c < n; c++) {
        const v = plottable(s.values[c]);
        if (v == null) {
          // Bridge gaps: keep the previous point so the next value connects
          // straight across the missing categories instead of breaking.
          if (!decor.bridgeGaps) prev = null;
          continue;
        }
        const forecast = fc != null && c >= fc;
        const pt = { x: slots.centers[c], y: scale.toY(v) };
        columnTop[c] = Math.min(columnTop[c], pt.y);
        if (prev && !smooth) {
          const p = prev;
          const dashOpt = forecast ? { dash: [4, 3] } : {};
          const seg = (x1: number, y1: number, x2: number, y2: number, suffix: string) =>
            nodes.push({
              kind: "line",
              x1,
              y1,
              x2,
              y2,
              stroke: color,
              strokeWidth: 2,
              ...dashOpt,
              name: `line-${si}-${c}${suffix}`,
            });
          if (decor.stepped === "after") {
            seg(p.x, p.y, pt.x, p.y, "a");
            seg(pt.x, p.y, pt.x, pt.y, "b");
          } else if (decor.stepped === "before") {
            seg(p.x, p.y, p.x, pt.y, "a");
            seg(p.x, pt.y, pt.x, pt.y, "b");
          } else if (decor.stepped === "center") {
            const mx = (p.x + pt.x) / 2;
            seg(p.x, p.y, mx, p.y, "a");
            seg(mx, p.y, mx, pt.y, "b");
            seg(mx, pt.y, pt.x, pt.y, "c");
          } else {
            seg(p.x, p.y, pt.x, pt.y, "");
          }
        }
        // Marker: small square with a background ring so crossings stay legible.
        // A per-cell color override highlights the point (max/min/last…) with
        // a larger, recolored marker. Forecast points render hollow.
        const cellColor = s.colors?.[c];
        // Shrunk where the frame's edge is closer than the marker's own half
        // extent, and dropped below a point of it — see `markInFrame`.
        const r = markInFrame(cfg, pt.x, pt.y, cellColor ? 3.4 : 2.4);
        if (r > 0)
          nodes.push({
            kind: "rect",
            x: pt.x - r,
            y: pt.y - r,
            w: r * 2,
            h: r * 2,
            fill: forecast && !cellColor ? style.background : (cellColor ?? color),
            stroke: forecast && !cellColor ? color : style.background,
            strokeWidth: 1,
            name: `marker-${si}-${c}`,
          });
        if (decor.segmentLabels) {
          nodes.push({
            kind: "text",
            x: pt.x - 30,
            // Never above the canvas. The label sits `fs * 1.65` over its own
            // point, so a series touching the top of the plot put it off the
            // chart — 39.2pt past the top at a 32pt font. Clamped rather than
            // dropped, on the same reasoning as a column total: this is the
            // value of ITS point and appears nowhere else, and an overlapping
            // label still reads where an off-canvas one is lost.
            y: Math.max(0, pt.y - fs * 1.65),
            w: 60,
            h: fs * 1.4,
            text: segmentLabel(decor.labelContent ?? ["value"], {
              value: v,
              fraction: null,
              series: s.name,
              category: data.categories[c],
              fmt,
            }),
            fontSize: fs,
            bold: !!cellColor,
            color: cellColor ?? style.text,
            align: "center",
            valign: "bottom",
            name: `label-${si}-${c}`,
          });
        }
        prev = pt;
        // The LAST DRAWN point, which is the last category only when the line
        // reaches it. Taken from the last category alone, a line that stopped one
        // period early (a blank last cell) was named nowhere — upright there is no
        // legend — see `namesWithoutMark`. A gap is not a value a stand-in could
        // be drawn at, so the name is level with where the line ended.
        lastSegMid[si] = pt.y;
      }
    });
  }

  nodes.push(baselineNode(frame, y0, style));
  // The FILTERED series list, because `lastSegMid` is indexed by it. Handing
  // over `cfg` paired the drawn line's end-y with `cfg.data.series[0].name`, so
  // a chart whose `Band low` / `Band high` rows came first labelled its one
  // drawn line "Band low" — a row the reference says is never drawn as a line
  // at all.
  if (decor.seriesLabels) nodes.push(...seriesLabelNodes(drawn, style, frame, lastSegMid));

  return {
    nodes,
    anchors: {
      categoryX: slots.centers,
      categoryWidth: data.categories.map(() => slots.colWidth || 10),
      columnTop,
      columnValue: area
        ? data.categories.map((_, c) => stackedPos[c] + stackedNeg[c])
        : data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      ...measuredAnchors(data, area, seriesLevels),
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: scale.toY,
    },
  };
}

/**
 * The value range a chart with no value axis maps onto its plot: the scale it
 * was handed where it was handed one, otherwise the range of its own data.
 *
 * The sparkline and the slope chart each took `min` and `max` from their own
 * values and never looked at `cfg.scale`, so a scale the author pinned — what
 * the pane's "Same scale" writes — drew exactly as no scale at all: a sparkline
 * of `[10, 12, 11, 14]` spanned 40.8pt with `scale: { min: 0, max: 100 }` and
 * with none. It is also how `buildMultiples` hands a panel the shared scale
 * ("so panels compare honestly"), which the slope chart now takes like any other
 * line panel; a table of sparklines is deliberately not given it there, because
 * its rows are KPIs of unlike units.
 *
 * A pinned range the data cannot live in — `min` at or above the data's top, or
 * an empty span — falls back to the data's own, as `valueScale` does for the
 * charts that have an axis. Unpinned, the answer is the data's range exactly.
 */
function valueRange(cfg: ChartConfig, all: number[]): { lo: number; hi: number } {
  const ownLo = all.length ? minOf(all) : 0;
  const ownHi = all.length ? maxOf(all) : 1;
  const lo = cfg.scale?.min ?? ownLo;
  const hi = cfg.scale?.max ?? ownHi;
  return hi > lo ? { lo, hi } : { lo: ownLo, hi: ownHi };
}

/**
 * Sparkline: a compact, chrome-less trend line sized to sit inline (think
 * Tufte's "word-sized graphic"). No axes, gridlines or category labels — just
 * a thin line, an optional leading label (the title / series name) and a
 * trailing value, with dots on the min (red), max (green) and last points.
 * Pair with `multiples` to get a table of sparklines, one per series.
 */
function layoutSparkline(cfg: ChartConfig, style: ChartStyle, _decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const area = cfg.kind === "area";
  const single = data.series.length === 1;
  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const fmt = resolveFormat(all, cfg.numberFormat);
  const { lo, hi } = valueRange(cfg, all);
  const span = hi - lo || 1;

  // Leading label (title/series name) and trailing last value reserve gutters.
  const label = cfg.title ?? (single ? data.series[0].name : "");
  const labelW = label ? Math.min(cfg.width * 0.38, textWidth(label, fs) + 8) : 0;
  const lastVal = single ? ([...data.series[0].values].reverse().find((v): v is number => v != null) ?? null) : null;
  const endText = lastVal != null ? formatNumber(lastVal, fmt) : "";
  const endW = endText ? textWidth(endText, fs) + 8 : 0;
  const padY = Math.max(2, cfg.height * 0.16);
  const plot = {
    x: labelW,
    y: padY,
    w: Math.max(10, cfg.width - labelW - endW - 4),
    h: Math.max(4, cfg.height - padY * 2),
  };
  const xs = data.categories.map((_, c) => plot.x + (n === 1 ? plot.w / 2 : (c / (n - 1)) * plot.w));
  // Clipped to the range, as `valueScale`'s `toY` clips to the plot: a scale
  // narrower than the data is a scale the data runs off the end of. The value is
  // clipped rather than the result so that a value inside the range maps through
  // exactly the arithmetic it always did.
  const toY = (v: number) => plot.y + (1 - (Math.max(lo, Math.min(hi, v)) - lo) / span) * plot.h;

  const nodes: SceneNode[] = [];
  if (label) {
    nodes.push({
      kind: "text",
      x: 0,
      y: plot.y + plot.h / 2 - fs * 0.75,
      w: labelW - 6,
      h: fs * 1.5,
      text: label,
      fontSize: fs,
      color: style.text,
      align: "left",
      valign: "middle",
      name: "spark-label",
    });
  }

  data.series.forEach((s, si) => {
    const color = seriesColor(style, si, s.color);
    const pts = s.values
      .map((v, c) => (v == null ? null : { x: xs[c], y: toY(v), v, c }))
      .filter((p): p is { x: number; y: number; v: number; c: number } => p != null);
    // Light area fill under the line (area kind only), per-segment rects to the floor.
    if (area) {
      const fill = lerpColor(style.background, color, 0.16);
      const floor = plot.y + plot.h;
      for (let i = 0; i < pts.length - 1; i++) {
        const span = pts[i + 1].x - pts[i].x;
        // Cap at the sparkline's pre-#128 fixed count: slabSteps' 24 is the
        // column/area budget, so a wide sparkline segment tripled its shape count
        // and blew the Office.js budget a sparkline exists to stay well under.
        const steps = Math.min(8, slabSteps(span));
        const w = span / steps;
        for (let k = 0; k < steps; k++) {
          const t = (k + 0.5) / steps;
          const y = pts[i].y + (pts[i + 1].y - pts[i].y) * t;
          nodes.push({
            kind: "rect",
            x: pts[i].x + k * w,
            y,
            w: w + 0.5,
            h: Math.max(0, floor - y),
            fill,
            name: `spark-fill-${si}-${i}-${k}`,
          });
        }
      }
    }
    for (let i = 0; i < pts.length - 1; i++) {
      nodes.push({
        kind: "line",
        x1: pts[i].x,
        y1: pts[i].y,
        x2: pts[i + 1].x,
        y2: pts[i + 1].y,
        stroke: color,
        strokeWidth: 1.25,
        name: `spark-${si}-${pts[i + 1].c}`,
      });
    }
    // Min / max / last dots (single-series only, to stay uncluttered).
    if (single && pts.length) {
      let minP = pts[0];
      let maxP = pts[0];
      for (const p of pts) {
        if (p.v < minP.v) minP = p;
        if (p.v > maxP.v) maxP = p;
      }
      const last = pts[pts.length - 1];
      nodes.push(
        { kind: "ellipse", cx: minP.x, cy: minP.y, rx: 1.9, ry: 1.9, fill: style.negative, name: `spark-min-${si}` },
        { kind: "ellipse", cx: maxP.x, cy: maxP.y, rx: 1.9, ry: 1.9, fill: "#1a9e6e", name: `spark-max-${si}` },
        { kind: "ellipse", cx: last.x, cy: last.y, rx: 2.1, ry: 2.1, fill: color, name: `spark-last-${si}` },
      );
      if (endText) {
        nodes.push({
          kind: "text",
          x: plot.x + plot.w + 4,
          y: last.y - fs * 0.75,
          w: endW,
          h: fs * 1.5,
          text: endText,
          fontSize: fs,
          bold: true,
          color,
          align: "left",
          valign: "middle",
          name: `spark-end-${si}`,
        });
      }
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: xs,
      categoryWidth: data.categories.map(() => plot.w / Math.max(1, n)),
      columnTop: data.categories.map(() => plot.y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      columnHasData: firstSeriesHasData(data),
      seriesLevels: ownValueLevels(data),
      baselineY: plot.y + plot.h,
      plot,
      valueToY: toY,
    },
  };
}

/**
 * Slope chart: the before/after comparison. No value axis or gridlines —
 * two vertical hairlines carry the periods, every series is a straight
 * line (or polyline for >2 categories) with a "Name value" label at both
 * ends, colored like its line. Labels de-overlap vertically per side.
 */
function layoutSlope(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const last = n - 1;
  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const fmt = resolveFormat(all, cfg.numberFormat);
  const { lo, hi } = valueRange(cfg, all);
  const span = hi - lo || 1;

  const endLabel = (s: (typeof data.series)[number], c: number) =>
    s.values[c] == null ? "" : `${s.name} ${formatNumber(s.values[c]!, fmt)}`;
  const gutterL = Math.min(
    cfg.width * 0.34,
    Math.max(fs, ...data.series.map((s) => textWidth(endLabel(s, 0), fs))) + 10,
  );
  const gutterR = Math.min(
    cfg.width * 0.34,
    Math.max(fs, ...data.series.map((s) => textWidth(endLabel(s, last), fs))) + 10,
  );
  /**
   * The size both gutters' labels are drawn at.
   *
   * Each gutter is sized from its own labels and then CAPPED at a third of the
   * chart's width, and the labels were drawn at the full chart font regardless —
   * so on a narrow chart a name-plus-value was wider than the room it had been
   * given and ran off the edge of the chart: 16pt past an 80x60 frame at an 18pt
   * font, 67pt at 32pt, and neither PowerPoint renderer clips a text box.
   *
   * Shrunk together and then clipped, for the reason the column chart's series
   * names give: labels at differing sizes read as a hierarchy that is not there.
   * Both sides take one size for the same reason. Where the cap does not bite —
   * every comfortable chart — the gutter is the widest label plus 10, the loop
   * does not run, and nothing moves.
   */
  const endFits = (f: number) =>
    data.series.every(
      (s) => textWidth(endLabel(s, 0), f) <= gutterL - 6 && textWidth(endLabel(s, last), f) <= gutterR - 6,
    );
  let endFs = fs;
  while (endFs > MIN_LABEL_FS && !endFits(endFs)) endFs -= 0.5;

  const titleH = titleHeight(cfg, style);
  const headerH = fs * 1.5; // period labels above the rails
  // Through `fitPlot`, like every other layout: this one subtracted its chrome
  // and used the answer, so a frame too short to pay for a title, a period
  // header and a footnote row gave it a NEGATIVE height — and a negative height
  // is not a small plot, it is an inverted axis, with `toY` mapping larger
  // values downward and the end labels placed by a band whose bottom sat above
  // its top. An 80x60 slope chart at an 18pt font had exactly that.
  const plot = fitPlot(cfg, {
    x: gutterL,
    y: titleH + headerH + 4,
    w: cfg.width - gutterL - gutterR,
    h: cfg.height - titleH - headerH - 4 - footnoteH(cfg, style, decor) - 6,
  });
  const pad = plot.h * 0.08;
  // Clipped to the range, for the reason the sparkline's is.
  const toY = (v: number) => plot.y + pad + (1 - (Math.max(lo, Math.min(hi, v)) - lo) / span) * (plot.h - pad * 2);
  const xs = data.categories.map((_, c) => plot.x + (n === 1 ? plot.w / 2 : (c / (n - 1)) * plot.w));

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Rails and period labels at the two ends only. The labels are drawn where
  // their band is ON the canvas: `titleH + headerH` is priced in the font, so
  // at a 32pt font on a 60pt-tall chart it ends below the foot of the chart and
  // both were drawn there. Same call the bump chart's period headers take — the
  // rails and the slopes are the chart, the period names are chrome.
  const headerFits = titleH + headerH <= cfg.height;
  for (const c of [0, last]) {
    nodes.push({
      kind: "line",
      x1: xs[c],
      y1: plot.y,
      x2: xs[c],
      y2: plot.y + plot.h,
      stroke: style.gridline,
      strokeWidth: 1,
      name: `slope-rail-${c}`,
    });
    if (headerFits)
      nodes.push({
        kind: "text",
        x: xs[c] - 60,
        y: titleH,
        w: 120,
        h: headerH,
        text: data.categories[c],
        fontSize: fs,
        color: style.mutedText,
        align: "center",
        valign: "middle",
        name: `category-${c}`,
      });
  }

  // Per-side label placement: keep each label at its line end, then push
  // apart to a minimum gap and clamp back inside the plot.
  const place = (ys: (number | null)[]): (number | null)[] => {
    const idx = ys
      .map((y, i) => ({ y, i }))
      .filter((e): e is { y: number; i: number } => e.y != null)
      .sort((a, b) => a.y - b.y);
    const minGap = fs * 1.5;
    // `y` is the label's CENTRE and the box is fs*1.5 tall, so half a box of
    // headroom keeps it on the canvas.
    const top = Math.max(plot.y, fs * 0.75);
    const bottom = Math.min(plot.y + plot.h, cfg.height - fs * 0.75);
    for (let k = 1; k < idx.length; k++) idx[k].y = Math.max(idx[k].y, idx[k - 1].y + minGap);
    if (idx.length) {
      idx[idx.length - 1].y = Math.min(idx[idx.length - 1].y, bottom);
      for (let k = idx.length - 2; k >= 0; k--) idx[k].y = Math.min(idx[k].y, idx[k + 1].y - minGap);
      // Only the BOTTOM was clamped, and the upward propagation then walked
      // straight past the plot top, past the title and off the canvas — on a
      // dense slope chart the topmost labels were simply lost. Spread evenly
      // over the plot when the gap cannot be honoured: an overlapping label
      // still reads, an off-canvas one does not (collide.ts's own rule).
      if (idx[0].y < top) {
        const step = idx.length > 1 ? (bottom - top) / (idx.length - 1) : 0;
        idx.forEach((e, k) => (e.y = top + k * step));
      }
    }
    const out: (number | null)[] = ys.map(() => null);
    idx.forEach((e) => (out[e.i] = e.y));
    return out;
  };
  const leftYs = place(data.series.map((s) => (s.values[0] == null ? null : toY(s.values[0]!))));
  const rightYs = place(data.series.map((s) => (s.values[last] == null ? null : toY(s.values[last]!))));

  const columnTop: number[] = data.categories.map(() => plot.y + plot.h);
  data.series.forEach((s, si) => {
    const color = seriesColor(style, si, s.color);
    let prev: { x: number; y: number } | null = null;
    for (let c = 0; c < n; c++) {
      const v = s.values[c];
      if (v == null) {
        prev = null;
        continue;
      }
      const pt = { x: xs[c], y: toY(v) };
      columnTop[c] = Math.min(columnTop[c], pt.y);
      if (prev) {
        nodes.push({
          kind: "line",
          x1: prev.x,
          y1: prev.y,
          x2: pt.x,
          y2: pt.y,
          stroke: color,
          strokeWidth: 2,
          name: `line-${si}-${c}`,
        });
      }
      const r = 2.4;
      const mr = markInFrame(cfg, pt.x, pt.y, r);
      if (mr > 0)
        nodes.push({
          kind: "rect",
          x: pt.x - mr,
          y: pt.y - mr,
          w: mr * 2,
          h: mr * 2,
          fill: color,
          stroke: style.background,
          strokeWidth: 1,
          name: `marker-${si}-${c}`,
        });
      prev = pt;
    }
    if (leftYs[si] != null) {
      nodes.push({
        kind: "text",
        x: 0,
        y: leftYs[si]! - fs * 0.75,
        w: gutterL - 6,
        h: fs * 1.5,
        text: clipToWidth(endLabel(s, 0), endFs, gutterL - 6),
        fontSize: endFs,
        color,
        align: "right",
        valign: "middle",
        name: `slope-left-${si}`,
      });
    }
    if (rightYs[si] != null) {
      nodes.push({
        kind: "text",
        x: plot.x + plot.w + 6,
        y: rightYs[si]! - fs * 0.75,
        w: gutterR - 6,
        h: fs * 1.5,
        text: clipToWidth(endLabel(s, last), endFs, gutterR - 6),
        fontSize: endFs,
        color,
        align: "left",
        valign: "middle",
        name: `slope-right-${si}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: xs,
      categoryWidth: data.categories.map(() => fs * 2),
      columnTop,
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      columnHasData: firstSeriesHasData(data),
      seriesLevels: ownValueLevels(data),
      baselineY: plot.y + plot.h,
      plot,
      valueToY: toY,
    },
  };
}

/**
 * Bump chart: rank-over-time. Categories are periods (x); each series is an
 * entity whose values are ranks (1 = best). Ranks map onto an inverted integer
 * axis (rank 1 at the top) with thick lines, round markers and a "Name" label
 * at both ends of every line.
 */
function layoutBump(cfg: ChartConfig, style: ChartStyle, _decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const titleH = titleHeight(cfg, style);
  const headerH = fs * 1.6;
  const ranks = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const maxRank = maxOf(ranks, 1);
  /**
   * The top row is rank 1 — unless the data holds something lower.
   *
   * `toY` below measured every rank from 1, so a rank below it went ABOVE the
   * plot by a whole row pitch or more: a 0-based ranking, a typo'd 0 among five
   * 1-based ranks, a negative, or bump switched on over ordinary line data with
   * negatives in it. The markers were dropped by `markInFrame` and the line
   * segments were kept at the raw y. Measured 2026-10-24 at 480x300: a 0-based
   * ranking of three drew lines from y = -252, a -3 to y = -1056, and one 0 among
   * five ranks to y = -51 — ink on the slide above the chart, which
   * `frame-fit.test.ts` says nothing a chart draws may be.
   *
   * Not clamped to 1, which would put 0 and 1 on the same row and call two
   * different ranks the same. The range simply starts at the lowest value there
   * is, so a 0-based ranking reads top-down from 0 and every rank stays on the
   * chart in its own row. Seeded with 1, so for every ranking that starts at 1
   * — the documented one — `minRank` is 1 and nothing moves.
   */
  const minRank = minOf(ranks, 1);
  /**
   * The two name gutters, and the size the names in them are drawn at.
   *
   * `max(fs * 3, …widest name) + fs` was uncapped, and the plot is what is left
   * after TWO of them — so at a 32pt font on an 80pt-wide chart the gutters
   * wanted 256 points of 80 and the plot came out NEGATIVE. Every rank marker
   * then landed to the right of the chart: 77pt past the edge, drawn onto
   * whatever sits beside it on the slide.
   *
   * Capped at a third of the width each, as the slope chart's gutters already
   * are, with the names shrunk to what the cap leaves and clipped past the
   * floor. Where the cap does not bite — every comfortable chart — the gutter is
   * the widest name plus `fs` and nothing moves.
   */
  const nameW = Math.min(cfg.width * 0.34, Math.max(fs * 3, ...data.series.map((s) => textWidth(s.name, fs))) + fs);
  const nameRoom = Math.max(1, nameW - fs * 0.5);
  let nameFs = fs;
  while (nameFs > MIN_LABEL_FS && data.series.some((s) => textWidth(s.name, nameFs) > nameRoom)) nameFs -= 0.5;
  const plot = fitPlot(cfg, {
    x: nameW,
    y: titleH + headerH,
    w: cfg.width - nameW * 2,
    h: cfg.height - titleH - headerH - fs * 1.6,
  });
  const xs = data.categories.map((_, c) => plot.x + (n === 1 ? 0 : (c / (n - 1)) * plot.w));
  const toY = (rank: number) =>
    plot.y + (maxRank === minRank ? plot.h / 2 : ((rank - minRank) / (maxRank - minRank)) * plot.h);

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Period headers along the top — where the band they sit in is ON the canvas.
  // `titleH + headerH` is a fixed number of points priced in the font, so at a
  // 32pt font on a 60pt-tall chart the header band alone ends 34pt below the
  // foot of the chart and every header was drawn there. Chrome that cannot be
  // paid for is not drawn; the ranks and their end labels are the chart.
  const headerFits = titleH + headerH <= cfg.height;
  if (headerFits)
    data.categories.forEach((cat, c) => {
      nodes.push({
        kind: "text",
        x: xs[c] - 40,
        y: titleH,
        w: 80,
        h: headerH,
        text: cat,
        fontSize: fs,
        bold: true,
        color: style.text,
        align: "center",
        valign: "middle",
        name: `period-${c}`,
      });
    });

  data.series.forEach((s, si) => {
    const color = seriesColor(style, si, s.color);
    let prev: { x: number; y: number } | null = null;
    let firstC = -1;
    let lastC = -1;
    s.values.forEach((v, c) => {
      if (v == null) {
        prev = null;
        return;
      }
      if (firstC < 0) firstC = c;
      lastC = c;
      const pt = { x: xs[c], y: toY(v) };
      if (prev)
        nodes.push({
          kind: "line",
          x1: prev.x,
          y1: prev.y,
          x2: pt.x,
          y2: pt.y,
          stroke: color,
          strokeWidth: 3.5,
          name: `bump-${si}-${c}`,
        });
      prev = pt;
    });
    s.values.forEach((v, c) => {
      if (v == null) return;
      const br = markInFrame(cfg, xs[c], toY(v), 4);
      if (br > 0)
        nodes.push({
          kind: "ellipse",
          cx: xs[c],
          cy: toY(v),
          rx: br,
          ry: br,
          fill: color,
          stroke: style.background,
          strokeWidth: 1.5,
          name: `bump-marker-${si}-${c}`,
        });
    });
    // "Name" labels at both ends of the line.
    if (firstC >= 0) {
      const y = toY(s.values[firstC]!);
      nodes.push({
        kind: "text",
        x: 0,
        // Clamped into the canvas: the label is centred on its line's rank, and
        // the top rank sits at the plot's own top edge, so `- fs * 0.75` put it
        // above the chart at a large font.
        y: Math.max(0, Math.min(y - fs * 0.75, cfg.height - fs * 1.5)),
        w: nameRoom,
        h: fs * 1.5,
        text: clipToWidth(s.name, nameFs, nameRoom),
        fontSize: nameFs,
        bold: true,
        color,
        align: "right",
        valign: "middle",
        name: `bump-label-l-${si}`,
      });
    }
    if (lastC >= 0) {
      const y = toY(s.values[lastC]!);
      nodes.push({
        kind: "text",
        x: plot.x + plot.w + fs * 0.5,
        // Clamped into the canvas: the label is centred on its line's rank, and
        // the top rank sits at the plot's own top edge, so `- fs * 0.75` put it
        // above the chart at a large font.
        y: Math.max(0, Math.min(y - fs * 0.75, cfg.height - fs * 1.5)),
        w: nameRoom,
        h: fs * 1.5,
        text: clipToWidth(s.name, nameFs, nameRoom),
        fontSize: nameFs,
        bold: true,
        color,
        align: "left",
        valign: "middle",
        name: `bump-label-r-${si}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: xs,
      categoryWidth: data.categories.map(() => fs * 2),
      columnTop: data.categories.map(() => plot.y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      columnHasData: firstSeriesHasData(data),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}

/**
 * Horizontal "profile chart": line / area rotated 90° — categories run down
 * the left axis and values extend to the right (think-cell parity). Kept
 * separate from the vertical path so that stays byte-identical.
 */
/**
 * A Catmull-Rom spline through a series' points, as line segments.
 *
 * Nulls break the line: the points are split into contiguous runs first, and
 * each run is sampled at STEPS points per segment. The maths is symmetric in x
 * and y, which is the whole reason this is shared — a horizontal chart's points
 * are (value, category) where a vertical chart's are (category, value), and
 * nothing else about the curve changes. It was inlined in the vertical path,
 * and the horizontal path simply had no curve at all.
 *
 * `c` is the category index the segment ENDS at, which is what decides whether
 * it belongs to the forecast; `k` distinguishes the samples within a segment so
 * every node gets its own name.
 */
export function splineSegments(
  pts: ({ x: number; y: number; c: number } | null)[],
  steps = 16,
): { x1: number; y1: number; x2: number; y2: number; c: number; k: number }[] {
  const runs: { x: number; y: number; c: number }[][] = [];
  let cur: { x: number; y: number; c: number }[] = [];
  for (const p of pts) {
    if (p) cur.push(p);
    else if (cur.length) {
      runs.push(cur);
      cur = [];
    }
  }
  if (cur.length) runs.push(cur);
  const out: { x1: number; y1: number; x2: number; y2: number; c: number; k: number }[] = [];
  for (const run of runs) {
    /**
     * THE CURVE MAY NOT LEAVE THE BOX ITS OWN POINTS SPAN.
     *
     * A uniform Catmull-Rom spline is not monotone: next to a spike it swings
     * past the neighbouring knot, so `[0, 0, 100, 0, 0]` dipped 7.4 value units
     * BELOW the zero it was drawn from and `[100, 100, 0, 100, 100]` rose above
     * the 100. Measured 2026-10-24 at 480x300: the curve reached y = 302.2 on a
     * 300pt canvas, through the category names and off the chart; sideways it
     * reached x = -20.3, 20pt off the left edge, with zero at x = 13.4. Launch-
     * shaped data (`[0, 0, 0, 120, 15, 10, 8, 8]`) dipped 16pt below the baseline
     * into the category labels. Each of those is a value the data never held,
     * drawn where a reader reads values — and a mark off the chart, which
     * `frame-fit.test.ts` says nothing a chart draws may be.
     *
     * Clamped to the extent of the run's own knots, in whichever axis: the knots
     * are already inside the plot (`toY` and `toX` clip them), so the curve is
     * too, and a curve that stays within its points never shows a value outside
     * the range of the ones supplied. Where the spline already stays inside, as it
     * does for most data, nothing moves; at a spike it runs flat along the
     * neighbour's level for the part of the swing that went past it, which reads
     * as the curve levelling off rather than as a negative number.
     */
    const span = {
      x0: minOf(run.map((p) => p.x)),
      x1: maxOf(run.map((p) => p.x)),
      y0: minOf(run.map((p) => p.y)),
      y1: maxOf(run.map((p) => p.y)),
    };
    for (let i = 0; i < run.length - 1; i++) {
      const p0 = run[Math.max(0, i - 1)];
      const p1 = run[i];
      const p2 = run[i + 1];
      const p3 = run[Math.min(run.length - 1, i + 2)];
      let pp = { x: p1.x, y: p1.y };
      for (let k = 1; k <= steps; k++) {
        const t = k / steps;
        const t2 = t * t;
        const t3 = t2 * t;
        const sx =
          0.5 *
          (2 * p1.x +
            (-p0.x + p2.x) * t +
            (2 * p0.x - 5 * p1.x + 4 * p2.x - p3.x) * t2 +
            (-p0.x + 3 * p1.x - 3 * p2.x + p3.x) * t3);
        const sy =
          0.5 *
          (2 * p1.y +
            (-p0.y + p2.y) * t +
            (2 * p0.y - 5 * p1.y + 4 * p2.y - p3.y) * t2 +
            (-p0.y + 3 * p1.y - 3 * p2.y + p3.y) * t3);
        const cx = Math.max(span.x0, Math.min(span.x1, sx));
        const cy = Math.max(span.y0, Math.min(span.y1, sy));
        out.push({ x1: pp.x, y1: pp.y, x2: cx, y2: cy, c: p2.c, k });
        pp = { x: cx, y: cy };
      }
    }
  }
  return out;
}

function layoutLineHorizontal(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data, bandLow, bandHigh } = splitBandRows(cfg.data);
  // The chrome legends `cfg.data.series`, so it has to see the same list the
  // plot draws — otherwise a band row that is (correctly) not drawn still gets
  // a legend chip. Untouched when there is no band row.
  const drawn: ChartConfig = data === cfg.data ? cfg : { ...cfg, data };
  const n = data.categories.length;
  const area = cfg.kind === "area";
  const fs = style.fontSize;

  const all = [
    ...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)),
    ...(bandLow ?? []).filter((v): v is number => v != null),
    ...(bandHigh ?? []).filter((v): v is number => v != null),
  ];
  const stackedPos = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
  const stackedNeg = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
  const dataMax = area ? maxOf(stackedPos, 0) : maxOf(all, 0);
  const dataMin = area ? minOf(stackedNeg, 0) : minOf(all, 0);
  const fmt = resolveFormat(all, cfg.numberFormat);

  const frame = computeFrameHorizontal(drawn, style, decor);
  const scale = valueScale(frame, dataMin, dataMax, cfg.scale);
  // Clamped like `scale.toY` is — see the note on the same map in column.ts. A
  // manual scale narrower than the data extrapolated off the slide here too.
  const toX = (v: number) =>
    frame.x + Math.max(0, Math.min(frame.w, ((v - scale.min) / (scale.max - scale.min || 1)) * frame.w));
  const slotH = frame.h / Math.max(1, n);
  const centers = data.categories.map((_, c) => frame.y + slotH * (c + 0.5));
  // Date categories space proportionally to time — the same rule the manual
  // states for line charts without qualifying it by orientation, and the same
  // arithmetic the vertical path runs along x. Sideways it was simply absent,
  // so "2024-01, 2024-02, 2024-12" sat in three equal rows and a reader saw a
  // steady march where the data has a ten-month gap.
  // …and only while the dates run forward, for the reason `dateSpacing` gives.
  const dateAt = area ? null : dateSpacing(data.categories);
  if (dateAt) {
    const inset = slotH / 2;
    for (let c = 0; c < n; c++) centers[c] = frame.y + inset + dateAt[c] * (frame.h - inset * 2);
  }
  const x0 = toX(0);

  const nodes: SceneNode[] = horizontalChrome(drawn, style, decor, frame, centers, scale, (v) => toX(v) - frame.x);
  const columnTop: number[] = data.categories.map(() => x0);

  /**
   * A filled band between two value series — the same slab trick the vertical
   * path uses, turned: the renderers have no polygon fill, so the gap between
   * two lines is drawn as thin rects marching along the CATEGORY axis, which
   * runs down the side here rather than across the bottom.
   */
  const ribbon = (lows: (number | null)[], highs: (number | null)[], fill: string, name: string) => {
    for (let c = 0; c < n - 1; c++) {
      const l0 = lows[c];
      const l1 = lows[c + 1];
      const h0 = highs[c];
      const h1 = highs[c + 1];
      if (l0 == null || l1 == null || h0 == null || h1 == null) continue;
      const span = centers[c + 1] - centers[c];
      // In PIXELS via `toX`, for the same reason as the upright ribbon.
      const steps = slabSteps(span, undefined, Math.max(Math.abs(toX(h1) - toX(h0)), Math.abs(toX(l1) - toX(l0))));
      const h = span / steps;
      for (let k = 0; k < steps; k++) {
        const t = (k + 0.5) / steps;
        const xH = toX(h0 + (h1 - h0) * t);
        const xL = toX(l0 + (l1 - l0) * t);
        nodes.push({
          kind: "rect",
          x: Math.min(xH, xL),
          y: centers[c] + k * h,
          w: Math.abs(xH - xL),
          h: h + 0.5,
          fill,
          name: `${name}-${c}-${k}`,
        });
      }
    }
  };
  // The confidence band from `Band low` / `Band high` rows. Sideways those rows
  // were never split off the series list at all, so they drew as two ordinary
  // lines — the reference says they never do — and the ribbon they exist to
  // shade was missing.
  if (!area && bandLow && bandHigh) {
    ribbon(
      bandLow,
      bandHigh,
      lerpColor(style.background, seriesColor(style, 0, data.series[0]?.color), 0.18),
      "band-ribbon",
    );
  }
  // The plan-vs-actual ribbon. It was one of four decorations this branch
  // simply did not read, so `fillBetween` was a byte-identical no-op sideways
  // while changing the vertical scene — the band a reader is meant to see as
  // the GAP between plan and actual was not drawn at all.
  if (!area && decor.fillBetween) {
    const [ai, bi] = decor.fillBetween;
    const sa = data.series[ai]?.values;
    const sb = data.series[bi]?.values;
    if (sa && sb)
      ribbon(sa, sb, lerpColor(style.background, seriesColor(style, ai, data.series[ai]?.color), 0.22), "fill-between");
  }

  // NO legend block here. `horizontalChrome` already draws one (`legendRow`,
  // shared with every other sideways chart) under the same condition, so this
  // path had TWO: the shipped showcase slide "Horizontal profile chart
  // (stacked area)" carried both, 2.5pt apart, and rendered every series name
  // as a smear. The shared one is also the better one — it wraps to a second
  // row via `legendWrapWalk`, matching what `computeFrameHorizontal` reserved,
  // and its chips carry the series' pattern/scenario paint rather than the bare
  // colour.

  if (area) {
    const posBase = data.categories.map(() => 0);
    const negBase = data.categories.map(() => 0);
    data.series.forEach((s, si) => {
      const fill = seriesColor(style, si, s.color);
      const lower: number[] = [];
      const upper: number[] = [];
      for (let c = 0; c < n; c++) {
        const v = s.values[c] ?? 0;
        if (v >= 0) {
          lower[c] = posBase[c];
          upper[c] = posBase[c] + v;
          posBase[c] += v;
        } else {
          upper[c] = negBase[c];
          lower[c] = negBase[c] + v;
          negBase[c] += v;
        }
      }
      for (let c = 0; c < n - 1; c++) {
        const xL0 = toX(lower[c]);
        const xL1 = toX(lower[c + 1]);
        const xU0 = toX(upper[c]);
        const xU1 = toX(upper[c + 1]);
        const span = centers[c + 1] - centers[c];
        // Rotated, so the edge moves in x — same rule, other axis.
        const steps = slabSteps(span, decor.stepped, Math.max(Math.abs(xL1 - xL0), Math.abs(xU1 - xU0)));
        const h = span / steps;
        for (let k = 0; k < steps; k++) {
          const t = (k + 0.5) / steps;
          // Stepped areas hold a flat edge across the interval (staircase);
          // "after" carries the left value, "before" the right, "center" both.
          // The vertical path has done this since `stepped` shipped; sideways
          // the slab interpolated regardless, so a stepped AREA chart claimed
          // the value slid where the data says it jumped — the same silent
          // no-op that `stepped` on a sideways LINE already had fixed.
          let xL: number;
          let xU: number;
          if (decor.stepped === "after" || (decor.stepped === "center" && t < 0.5)) {
            xL = xL0;
            xU = xU0;
          } else if (decor.stepped === "before" || decor.stepped === "center") {
            xL = xL1;
            xU = xU1;
          } else {
            xL = xL0 + (xL1 - xL0) * t;
            xU = xU0 + (xU1 - xU0) * t;
          }
          nodes.push({
            kind: "rect",
            x: Math.min(xL, xU),
            y: centers[c] + k * h,
            w: Math.abs(xU - xL),
            h: h + 0.5,
            fill,
            name: `area-${si}-${c}-${k}`,
          });
        }
      }
    });
    for (let c = 0; c < n; c++) columnTop[c] = toX(posBase[c]);
  } else {
    // The actuals/forecast divider, mirroring the vertical path. Across the
    // plot at the boundary between two categories — which is a HORIZONTAL rule
    // here, because the categories run down the side.
    const fc = decor.forecastFrom;
    if (fc != null && fc > 0 && fc < n) {
      const dy = centers[fc - 1] + (centers[fc] - centers[fc - 1]) / 2;
      nodes.push({
        kind: "line",
        x1: frame.x,
        y1: dy,
        x2: frame.x + frame.w,
        y2: dy,
        stroke: style.gridline,
        strokeWidth: 1,
        dash: [2, 3],
        name: "forecast-divider",
      });
    }
    // `smooth` was the fourth silent no-op here. The Catmull-Rom itself is
    // symmetric in x and y, so nothing about the CURVE needed turning — only
    // the points it is built from, which are (value, category) here and
    // (category, value) there. Shared with the vertical path rather than
    // copied, because two spline implementations would drift.
    const smooth = !!decor.smooth && !decor.stepped;
    data.series.forEach((s, si) => {
      const color = seriesColor(style, si, s.color);
      if (smooth) {
        const pts = s.values.map((v, c) => (v == null ? null : { x: toX(v), y: centers[c], c }));
        for (const seg of splineSegments(pts)) {
          nodes.push({
            kind: "line",
            x1: seg.x1,
            y1: seg.y1,
            x2: seg.x2,
            y2: seg.y2,
            stroke: color,
            strokeWidth: 2,
            ...(fc != null && seg.c >= fc ? { dash: [4, 3] } : {}),
            name: `line-${si}-${seg.c}-s${seg.k}`,
          });
        }
      }
      let prev: { x: number; y: number } | null = null;
      for (let c = 0; c < n; c++) {
        const v = s.values[c];
        if (v == null) {
          if (!decor.bridgeGaps) prev = null;
          continue;
        }
        // Everything from here down was reading three decoration fields and
        // stopping, so `forecastFrom`, `stepped` and the per-cell marker colour
        // were byte-identical no-ops sideways. A projection drawn with the same
        // solid stroke and filled marker as measured data is the worst of them:
        // the reader is shown a forecast as if it were fact, on a chart that
        // looks entirely correct. `stepped` is next — a step series drawn as
        // straight interpolation claims the value slid where the data says it
        // jumped.
        const forecast = fc != null && c >= fc;
        const pt = { x: toX(v), y: centers[c] };
        columnTop[c] = Math.max(columnTop[c], pt.x);
        if (prev && !smooth) {
          const p = prev;
          const dashOpt = forecast ? { dash: [4, 3] } : {};
          const seg = (x1: number, y1: number, x2: number, y2: number, suffix: string) =>
            nodes.push({
              kind: "line",
              x1,
              y1,
              x2,
              y2,
              stroke: color,
              strokeWidth: 2,
              ...dashOpt,
              name: `line-${si}-${c}${suffix}`,
            });
          // Turned, not reinterpreted: "after" still means hold the VALUE to
          // the next category and jump there, which sideways is a move along y
          // followed by a move along x.
          if (decor.stepped === "after") {
            seg(p.x, p.y, p.x, pt.y, "a");
            seg(p.x, pt.y, pt.x, pt.y, "b");
          } else if (decor.stepped === "before") {
            seg(p.x, p.y, pt.x, p.y, "a");
            seg(pt.x, p.y, pt.x, pt.y, "b");
          } else if (decor.stepped === "center") {
            const my = (p.y + pt.y) / 2;
            seg(p.x, p.y, p.x, my, "a");
            seg(p.x, my, pt.x, my, "b");
            seg(pt.x, my, pt.x, pt.y, "c");
          } else {
            seg(p.x, p.y, pt.x, pt.y, "");
          }
        }
        // Forecast points render hollow, and a per-cell colour highlights one
        // (max/min/last) with a larger recoloured marker — both as vertical.
        const cellColor = s.colors?.[c];
        // Shrunk where the frame's edge is closer than the marker's own half
        // extent, and dropped below a point of it — see `markInFrame`.
        const r = markInFrame(cfg, pt.x, pt.y, cellColor ? 3.4 : 2.4);
        if (r > 0)
          nodes.push({
            kind: "rect",
            x: pt.x - r,
            y: pt.y - r,
            w: r * 2,
            h: r * 2,
            fill: forecast && !cellColor ? style.background : (cellColor ?? color),
            stroke: forecast && !cellColor ? color : style.background,
            strokeWidth: 1,
            name: `marker-${si}-${c}`,
          });
        if (decor.segmentLabels) {
          // `labelContent` is read here as it is upright: this emitter printed
          // `formatNumber(v)` whatever the pane's "Label content" said, so
          // rotating a chart silently discarded the setting.
          const text = pointLabelText(decor, v, s.name, data.categories[c], fmt);
          nodes.push({
            kind: "text",
            ...labelBesidePoint(pt.x, text, fs, 4, cfg.width, fs * 3),
            y: pt.y - fs * 0.75,
            h: fs * 1.5,
            text,
            fontSize: fs,
            color: style.text,
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
        prev = pt;
      }
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => slotH * 0.6),
      columnTop,
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      columnHasData: firstSeriesHasData(data),
      baselineY: x0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      // Published for the same reason the column and waterfall layouts publish
      // it: `Error` and `Target` rows are extracted from the data whatever the
      // orientation, and with no map to draw them back they simply vanish.
      valueToX: toX,
    },
  };
}
