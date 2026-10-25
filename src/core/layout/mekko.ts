import type { ChartConfig, ChartStyle, Decorations, NumberFormat } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { authoredDecimals, formatNumber, formatPercent, resolveFormat, segmentLabel } from "../format";
import { seriesColor } from "../style";
import {
  printsOnTitle,
  MIN_LABEL_FS,
  titleInkBottom,
  aboveMarkFontSize,
  bandFontSize,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  endLabelNode,
  fitPlot,
  horizontalLegendFits,
  roomForEnds,
  titleHeight,
  type EndLabel,
} from "./frame";
import { legendRow, namesWithoutMark, seriesLabelNodes, type LayoutResult } from "./column";
import { columnPositiveTotal } from "./totals";

/**
 * What a mekko segment prints: the pane's "Label content" menu applied to it, as it
 * is to a stacked column's segment.
 *
 * This printed `formatNumber` whatever the menu said, so picking "%" or
 * "Series + value" changed every other column kind and left the mekko showing bare
 * values. The share is the segment's part of ITS column — the proportion the drawn
 * height encodes; the caller has returned for a segment with no value or a column
 * with no total, so `total` is a real divisor. Value alone is the default, which is
 * what it always printed.
 *
 * An empty `labelContent` list means "no label" and prints nothing, which the
 * caller turns into no text node rather than an empty box PowerPoint has to create.
 */
function segmentText(
  cfg: ChartConfig,
  decor: Decorations,
  fmt: NumberFormat,
  value: number,
  total: number,
  series: string,
  category: string,
): string {
  return segmentLabel(decor.labelContent ?? ["value"], {
    value,
    fraction: value / total,
    series,
    category,
    fmt,
    percentDecimals: authoredDecimals(cfg.numberFormat),
  });
}

/**
 * The totals a SIDEWAYS mekko prints past its rows' ends, as the frame wants them:
 * how wide each is at the size its row can carry, and how far along the plot its
 * row ends — all of the way for the %-axis variant, in proportion to its total for
 * the one with units.
 *
 * The size and the text are the ones the draw below uses (`bandFontSize` of the
 * row, `formatNumber` of the total), so the strip that is reserved and the number
 * that stands in it cannot disagree.
 */
function rowEnds(
  totals: number[],
  extents: number[],
  grand: number,
  catLen: number,
  units: boolean,
  maxTotal: number,
  fmt: NumberFormat,
  fs: number,
): EndLabel[] {
  return totals.flatMap((total, c) => {
    const rowFs = bandFontSize(fs, (extents[c] / grand) * catLen, 1.5);
    if (rowFs <= 0) return [];
    return [{ ink: textWidth(formatNumber(total, fmt), rowFs, true), reach: units ? total / maxTotal : 1 }];
  });
}

/**
 * Mekko (Marimekko) chart, think-cell style. Two variants:
 * - %-axis (default): column widths proportional to column totals, columns
 *   normalized to full height, so segment area ∝ absolute value.
 * - "Mekko with units": explicit column widths from the datasheet's
 *   `X extent` row; column heights represent absolute totals on a value scale.
 * cfg.horizontal rotates the chart: categories become rows.
 */
export function layoutMekko(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const H = !!cfg.horizontal;
  const units = !!data.xExtent?.some((v) => v != null && v > 0);

  const totals = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  const extents = units ? data.categories.map((_, c) => Math.max(0, data.xExtent?.[c] ?? 0)) : totals;
  const maxTotal = Math.max(1e-9, ...totals);
  const fmt = resolveFormat(
    [...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)), ...totals],
    cfg.numberFormat,
  );

  const decorFull = { ...decor, totals: true };
  let frame = H
    ? computeFrameHorizontal(cfg, style, { ...decorFull, totals: true })
    : computeFrame(cfg, style, decorFull, decor.seriesLabels ? data.series.map((s) => s.name) : []).frame;
  // The x the legend reservation was computed for, captured before the gutter
  // widening below moves `frame.x`. See the legend push for why.
  const legendX = frame.x;
  if (H && !units && decor.categoryAxis) {
    // Row labels carry a share suffix ("EMEA (32%)") the generic frame
    // reservation doesn't know about — widen the left gutter for it.
    //
    // Back through `fitPlot`, because this widening happens AFTER the frame was
    // floored and takes a fixed number of points however narrow the chart is:
    // on a 14pt-wide mekko it left `w: -12.2`, and a negative-width segment is
    // not a thin one — SVG drops it and PowerPoint clamps it to a sliver, so
    // every row loses its bars while the totals beside them still print. Any
    // frame that can afford the gutter is returned untouched.
    const extra = textWidth(" (00%)", fs);
    frame = fitPlot(cfg, { ...frame, x: frame.x + extra, w: frame.w - extra });
  }
  // The legend is drawn at the x the RESERVATION was computed for, not at the
  // widened frame.
  //
  // `computeFrameHorizontal` counts legend rows by walking the labels from the
  // legend's own x0 and says so — "this reservation and legendRow agree on the
  // row count". The widening above moves `frame.x` right AFTER that count, so
  // drawing at the new x gives the walk less room, wraps it onto more rows than
  // were reserved, and the extra rows come down on the bars' own labels. Inside
  // the frame, so no overflow gate sees it; at 160x120 it put the legend across
  // the row labels and the totals.
  //
  // Kept as the pre-widening x rather than by re-reserving, because the gutter
  // exists for the share suffix on the ROW labels and the legend does not carry
  // one — it has always been free to start where the reservation assumed.
  const grand = extents.reduce((a, b) => a + b, 0) || 1;
  /**
   * The gutter between columns, and the share of the axis it may take.
   *
   * A flat 2pt subtracted before the columns are sized is chrome charged
   * whatever the chart can afford: at 230 categories on a 480pt frame the gaps
   * alone want 458 of 458 points, so `catLen` reached zero and every segment
   * came out with a NEGATIVE width — 460 of them, worst -0.012 — while the
   * totals, the category names and the legend all still printed. A chart that
   * has vanished under its own separators, with labels around the space where
   * it was.
   *
   * The columns are the chart and the gaps are chrome for reading them apart, so
   * the gaps are scaled into a budget when they cannot be paid for in full. Same
   * trade the funnel's row gaps and the butterfly's centre gutter already make.
   * Below the budget — every ordinary chart — this is exactly 2 and nothing
   * moves.
   */
  const GAP_SHARE = 0.5;
  const axisLen = H ? frame.h : frame.w;
  const gap = n > 1 ? Math.min(2, (axisLen * GAP_SHARE) / (n - 1)) : 2;

  const nodes: SceneNode[] = [];
  const centers: number[] = [];
  const widths: number[] = [];
  const columnTop: number[] = [];
  const lastSegMid: (number | null)[] = data.series.map(() => null);

  // Category extent runs along x (vertical) or y (horizontal).
  const catLen = (H ? frame.h : frame.w) - gap * (n - 1);
  // Sideways, the totals stand past the rows' ends in a strip of `fs * 4`, and a
  // %-axis mekko's rows ALL run to the plot's edge — so every total of six figures
  // was cut. The plot gives up the width they ask for. The rows run down y, so the
  // slots above and `catLen` stand whatever the plot's width.
  if (H) frame = roomForEnds(cfg, frame, rowEnds(totals, extents, grand, catLen, units, maxTotal, fmt, fs));
  let pos = H ? frame.y : frame.x;

  for (let c = 0; c < n; c++) {
    const ext = (extents[c] / grand) * catLen;
    centers.push(pos + ext / 2);
    widths.push(ext);
    // %-variant: every column fills the plot; units variant: length ∝ total.
    const colLen = units ? (totals[c] / maxTotal) * (H ? frame.w : frame.h) : H ? frame.w : frame.h;

    let acc = H ? frame.x : frame.y + frame.h;
    data.series.forEach((s, si) => {
      const v = Math.max(0, s.values[c] ?? 0);
      if (v === 0 || totals[c] === 0) {
        // No segment in the last column, and still a name — see `namesWithoutMark`.
        // At the running level, which is where the stacked column puts it too.
        if (c === n - 1 && namesWithoutMark(s, seriesColor(style, si, s.color))) lastSegMid[si] = acc;
        return;
      }
      const segLen = (v / totals[c]) * colLen;
      const r = H ? { x: acc, y: pos, w: segLen, h: ext } : { x: pos, y: acc - segLen, w: ext, h: segLen };
      acc = H ? acc + segLen : acc - segLen;
      const fill = seriesColor(style, si, s.color);
      // Transparent "no-fill" segment: it still occupies the column (`acc` has
      // moved past it) but draws nothing, floating the segments above — same
      // contract as the stacked column (column.ts). Emitting the rect made both
      // PowerPoint sinks paint it: the pptx maps a bare word to mid grey and
      // Office.js hands "transparent" to setSolidColor, which it rejects.
      if (fill === "transparent") return;
      nodes.push({ kind: "rect", ...r, fill, stroke: style.background, strokeWidth: 0.75, name: `seg-${si}-${c}` });
      if (c === n - 1) lastSegMid[si] = H ? r.x + r.w / 2 : r.y + r.h / 2;
      // The label is centred in r.h and spans r.w in BOTH orientations, so the
      // vertical fit is r.h either way. (Gating it on r.w for horizontal mekko
      // measured the value-axis length, not the row thickness, so labels
      // rendered in rows thinner than the font.)
      //
      // Horizontal segments also need a length floor. The fit check below allows
      // 2pt of bleed — the text box is deliberately 4pt wider than the segment —
      // which is harmless for a vertical mekko's wide columns but lets a row of
      // hairline-thin horizontal segments print their labels on top of each
      // other. Vertical keeps exactly the room it always had.
      const roomAcross = r.h >= fs * 1.25;
      const roomAlong = !H || r.w >= fs * 1.25;
      if (decor.segmentLabels && roomAcross && roomAlong) {
        const label = segmentText(cfg, decor, fmt, v, totals[c], s.name, data.categories[c]);
        if (label && textWidth(label, fs) <= r.w + 2) {
          nodes.push({
            kind: "text",
            x: r.x - 2,
            y: r.y + r.h / 2 - fs * 0.75,
            w: r.w + 4,
            h: fs * 1.5,
            text: label,
            fontSize: fs,
            color: contrastInk(fill),
            align: "center",
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
      }
    });
    columnTop.push(H ? frame.x + colLen : frame.y + frame.h - colLen);

    // Column total at the column's end — the Mekko signature.
    if (H) {
      // Bound by the row it ends, not by the chart's font: a Mekko's rows are
      // proportional to their totals, so a small column gets a thin row, and a
      // total centred in a box `fs * 1.5` tall then runs into its neighbours.
      // Per-row rather than one size for the axis, because the rows genuinely
      // differ in height here. Last resort: a row that can afford `fs` keeps it.
      //
      // A row can be ZERO here — the rows are shares of the total, so a chart
      // whose data is all zeroes has no rows at all — and `Math.min` answered a
      // font of zero, which OOXML rejects outright. `bandFontSize` is the floor.
      //
      // And by the room beyond the row's end. The frame made it (`rowEnds`), so
      // this is the size it was; where the strip was capped before it could hold the
      // number, the number shrinks to what there is or is left out — never cut.
      const total = endLabelNode(
        formatNumber(totals[c], fmt),
        frame.x + colLen,
        centers[c],
        bandFontSize(fs, ext, 1.5),
        cfg.width,
        style,
        `total-${c}`,
      );
      if (total) nodes.push(total);
    } else {
      // The upright total sits ABOVE the column, and on the %-variant every
      // column fills the plot — so the only room it has is the band the frame
      // reserved over the plot, `fs * 1.5 + 4`, which `fitPlot` spends when the
      // frame cannot pay for its chrome. At 24pt on an 80x60 chart that put the
      // number 34pt above the top of its own chart. Fitted to the band that is
      // actually there and dropped when there is none, like the horizontal
      // branch just above.
      const colTop = frame.y + frame.h - colLen;
      // Measured from the TITLE's ink rather than from the top of the chart:
      // clamping a total to y=0 keeps it on the canvas and puts it on the title,
      // which is the trade `aboveMarkFontSize` exists to refuse.
      //
      // And by the column's own WIDTH, which nothing bounded: a mekko's columns
      // are proportional to their totals, so a narrow column gets a narrow slot
      // and the number over it was drawn at the chart font whatever that slot
      // came to — adjacent totals met on an 80x60 chart at 14pt. The room is the
      // column, so neighbouring labels abut instead of overlapping, which is the
      // rule this file's own legend already follows.
      const totalText = formatNumber(totals[c], fmt);
      let totalFs = aboveMarkFontSize(fs, colTop, titleInkBottom(cfg, style), 1.45);
      while (totalFs > MIN_LABEL_FS && textWidth(totalText, totalFs, true) > ext) totalFs -= 0.5;
      if (totalFs >= MIN_LABEL_FS && textWidth(totalText, totalFs, true) <= ext)
        nodes.push({
          kind: "text",
          x: pos - 4,
          y: colTop - totalFs * 1.45,
          w: ext + 8,
          h: totalFs * 1.4,
          text: totalText,
          fontSize: totalFs,
          bold: true,
          color: style.text,
          align: "center",
          valign: "bottom",
          name: `total-${c}`,
        });
    }

    pos += ext + gap;
  }

  // Chrome: title only via chromeNodes, custom category labels per orientation.
  nodes.push(
    ...chromeNodes(
      cfg,
      style,
      { ...decorFull, categoryAxis: false, valueAxis: false, gridlines: false },
      frame,
      centers,
    ),
  );
  if (decor.categoryAxis) {
    const catLabels = data.categories.map((cat, c) =>
      units ? cat : `${cat} (${formatPercent(extents[c] / grand, 0, false, cfg.numberFormat?.locale)})`,
    );
    // A Mekko's category label sits under its own COLUMN and is as wide as it
    // needs to be, so a label wider than its column overflows the box
    // symmetrically — into the neighbouring label, and off the left edge of the
    // chart for the first one. At an 22pt font the three-category sample had two
    // overlapping pairs and `category-0` starting at x = -2, which reads as one
    // run-on string: "EMEA (32%)Americas (42%)APAC (27%)".
    //
    // Shrunk together until the widest fits its own column, then clipped —
    // together, because category labels that differ in size read as a hierarchy
    // that is not there. Same two-step the agenda and the process flow already
    // use, and `catFs` stays at `fs` whenever they fit, so an ordinary chart is
    // untouched.
    let catFs = fs;
    // The room is the column plus the gap it shares with its neighbours — half
    // on each side, so adjacent boxes ABUT rather than overlap. It used to be
    // `+ 8` against a 2pt gap, which let every label bleed 4pt into each
    // neighbour: "Americas (42%)" and "APAC (27%)" collided by 6pt on an
    // ordinary 200x150 mekko, at the default font.
    const catRoom = (c: number) => widths[c] + gap;
    // Rotating the chart rotates which side of the label is crowded, and this
    // predicate used to answer `false` outright when horizontal — so the whole
    // fit above was skipped and the names were drawn at `fs`, unclipped, in a
    // gutter `computeFrameHorizontal` caps at 30% of the width. Sideways the
    // room is that gutter, and it is the same for every row, so the axis still
    // shrinks together and still reads at one size.
    const gutter = Math.max(1, frame.x - 4);
    const overflows = (f: number) =>
      H ? catLabels.some((l) => textWidth(l, f) > gutter) : catLabels.some((l, c) => textWidth(l, f) > catRoom(c));
    while (catFs > 6 && overflows(catFs)) catFs -= 0.5;
    for (let c = 0; c < n; c++) {
      const label = catLabels[c];
      if (H) {
        // Bounded by the gutter (width, shared) AND by the row it names
        // (height, per-row — a Mekko's rows are proportional to their totals,
        // so a small column gets a thin one). The smaller wins, the same way
        // the butterfly's category names take the smaller of their gutter fit
        // and their row fit.
        const nameFs = bandFontSize(catFs, catRoom(c), 1.5);
        if (nameFs > 0)
          nodes.push({
            kind: "text",
            x: 0,
            y: centers[c] - nameFs * 0.75,
            w: gutter,
            h: nameFs * 1.5,
            text: clipToWidth(label, nameFs, gutter),
            fontSize: nameFs,
            color: style.text,
            align: "right",
            valign: "middle",
            name: `category-${c}`,
          });
      } else if (!printsOnTitle(cfg, style, frame.y + frame.h + 3)) {
        // Same rule as the shared category axis: a strip squeezed into the
        // title's band is not drawn at all.
        nodes.push({
          kind: "text",
          x: centers[c] - catRoom(c) / 2,
          y: frame.y + frame.h + 3,
          w: catRoom(c),
          h: catFs * 1.4,
          text: clipToWidth(label, catFs, catRoom(c)),
          fontSize: catFs,
          color: style.text,
          align: "center",
          valign: "top",
          name: `category-${c}`,
        });
      }
    }
  }
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
    // Gated by the SAME predicate the reservation uses. `computeFrameHorizontal`
    // asks `horizontalLegendFits` before it counts any legend rows, and this
    // draw did not ask at all — so on a frame the predicate refuses, the
    // reservation was ZERO rows and the legend was drawn into it anyway. That is
    // the worst of the three options and the one a half-fix produces: it
    // measured worse for the scatter than not gating either side. At 120x90
    // horizontal it put a three-row legend across the bars, their totals and
    // the row labels — four overlapping pairs, all inside the frame, so no
    // overflow gate could see it.
    if (decor.seriesLabels && data.series.length > 1 && horizontalLegendFits(cfg, style, decorFull)) {
      nodes.push(...legendRow(cfg, style, legendX, titleHeight(cfg, style) + 2, { maxX: cfg.width - 4 }));
    }
  } else {
    nodes.push({
      kind: "line",
      x1: frame.x,
      y1: frame.y + frame.h,
      x2: frame.x + frame.w,
      y2: frame.y + frame.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
    if (decor.seriesLabels) nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: widths,
      columnTop,
      columnValue: totals,
      baselineY: H ? frame.x : frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
    },
  };
}
