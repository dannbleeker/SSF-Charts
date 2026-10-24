import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { MIN_LABEL_FS, bandFontSize, firstSeriesValues, fitPlot, footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Funnel chart: pipeline stages as horizontally centered bands whose width
 * is proportional to the value (the Power BI convention — plain rectangles,
 * so it renders natively in every renderer). Between stages, a muted
 * conversion label states the % of the previous stage. Order the values
 * ascending for a pyramid.
 */
export function layoutFunnel(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const { plot: values, raw } = firstSeriesValues(data, n);
  // `|| 1` guards only the all-zero funnel (0 / 0). The floor that stood here —
  // `Math.max(1, ...values)` — guarded that too, and ALSO scaled any funnel whose
  // widest stage is below 1 to that fraction of the plot: conversion rates typed
  // as 0.62 / 0.41 / 0.18 drew 266pt wide where 62 / 41 / 18 drew 428, and 0.05
  // was a 21pt sliver. Below ~0.005 every band hit the 2pt minimum and the stages
  // stopped comparing at all. The width says "proportional to the value", so the
  // unit must not set how much of the plot it takes.
  const vMax = Math.max(0, ...values) || 1;
  const fmt = resolveFormat(values, cfg.numberFormat);

  const titleH = titleHeight(cfg, style);
  // Left gutter for stage names when category labels are on (default).
  const catW = decor.categoryAxis
    ? Math.min(cfg.width * 0.28, Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 10)
    : 2;
  const plot = fitPlot(cfg, {
    x: catW,
    y: titleH + 2,
    w: cfg.width - catW - 4,
    h: cfg.height - titleH - 2 - footnoteH(cfg, style, decor) - 4,
  });
  // Room for the conversion label between bands — but the BANDS are the chart
  // and the gaps are chrome for a label, so the split is not "a point per band,
  // gaps take the rest". That is what this was, and on a 120x90 frame four gaps
  // took 59 of 64 points: five HAIRLINE bands at the 1pt floor, which is a
  // degenerate geometry rather than a small chart.
  //
  // The gaps get at most HALF the plot between them. Below that the fixed
  // `fs * 1.5` still wins, so nothing at an ordinary size moves — and either
  // way the cumulative pitch stays inside the plot, which is what the clamp was
  // originally for: a fixed gap on a short frame with many stages drove
  // `bandH + gap` per stage past the bottom and the last bands rendered
  // off-frame.
  //
  // AND THE BAND'S OWN FLOOR COMES OUT OF THE GAP, not out of the plot. `Math.max(1, …)`
  // is a floor that RAISES a height — the same shape as the quadrant label's
  // width floor and the scatter legend's before it — so with 24 stages on a 60pt
  // chart the 24 floored bands plus 23 gaps measured 36 points of a 24-point
  // plot and the last stage was drawn 10.4pt below the chart. Solving the gap
  // against the floored band keeps the pitch inside the plot: the chart's own
  // marks win the argument with the chrome between them, which is what the
  // paragraph above already says the split is for.
  //
  // Where even the hairlines do not fit — more stages than the plot has points —
  // the floor is abandoned rather than the stages. A sub-point band is a thin
  // line; a band drawn past the foot of the chart is on the slide.
  const BAND_MIN = 1;
  const gapWant = Math.min(fs * 1.5, (plot.h * 0.5) / Math.max(1, n - 1));
  const gap = Math.max(0, Math.min(gapWant, (plot.h - n * BAND_MIN) / Math.max(1, n - 1)));
  // No floor of its own: the gap above already leaves `BAND_MIN` per band
  // wherever the plot can pay for it, and where it cannot the bands go thinner
  // rather than off the chart. `Math.max(1, …)` here was the floor that raised a
  // height past its own container.
  const bandH = Math.max(0, plot.h - gap * (n - 1)) / Math.max(1, n);
  const cx = plot.x + plot.w / 2;

  // The BANDS are solved for the plot above, so they always fit. Their LABELS
  // were not, and each escaped the frame a different way once the font outgrew
  // the chart:
  //
  //   - a label is centred in its band and its ink is about `f` tall, so once
  //     the font passed the band height the bottom row's label hung below the
  //     plot — 4.4pt at a 28pt font, 15.1 at 36;
  //   - a category name sits right-aligned in a gutter capped at 28% of the
  //     width, and a name wider than that cap runs off the LEFT edge of the
  //     chart: "Negotiation" by 35.9pt at 28, and 83.4pt at 36.
  //
  // Both are the same answer the agenda and the process flow already use, and
  // the one this file's value labels now use for the horizontal case: shrink
  // until it fits, then clip the remainder. Shrunk TOGETHER so a row's name and
  // its number stay the same size, and last-resort — at any font that already
  // fits, `labelFs` is `fs` and nothing moves.
  //
  // The loop used to stop at 6 whether or not the band could hold a 6pt label,
  // and a funnel's bands go to nothing long before its font does: at 14pt on an
  // 80x60 frame the bands are three points tall and every row's name was drawn
  // through the two either side of it, sixteen overlapping pairs across the
  // sweep. A floor that ignores its own reservation is the thing putting labels
  // where they cannot be read, so below `MIN_LABEL_FS` the whole set is dropped
  // — the same answer the radar, sunburst, tilemap and pie reservations give,
  // and the bands themselves still show the shape.
  const labelFs = (() => {
    const gutter = catW - 4;
    const tooWide = (f: number) => decor.categoryAxis && data.categories.some((c) => textWidth(c, f) > gutter);
    let f = Math.min(fs, bandH / 1.15);
    while (f > MIN_LABEL_FS && tooWide(f)) f -= 0.5;
    return f >= MIN_LABEL_FS ? f : 0;
  })();

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  const columnTop: number[] = [];
  values.forEach((v, c) => {
    const y = plot.y + c * (bandH + gap);
    columnTop.push(y);
    const w = Math.max(2, (v / vMax) * plot.w);
    const fill = data.series[0]?.colors?.[c] ?? style.palette[c % style.palette.length];
    nodes.push({ kind: "rect", x: cx - w / 2, y, w, h: bandH, fill, name: `stage-${c}` });

    if (decor.categoryAxis && labelFs > 0) {
      nodes.push({
        kind: "text",
        x: 0,
        y: y + bandH / 2 - labelFs * 0.75,
        w: catW - 4,
        h: labelFs * 1.5,
        text: clipToWidth(data.categories[c], labelFs, catW - 4),
        fontSize: labelFs,
        color: style.text,
        align: "right",
        valign: "middle",
        name: `category-${c}`,
      });
    }
    // `raw[c]`, not `v`: `v` is the CLAMPED value, so a blank cell and a negative
    // both reach here as 0 and the label states it as measured. No value, no label
    // — which is what clustered and line have always done.
    if (decor.segmentLabels && labelFs > 0 && raw[c] !== null) {
      const label = formatNumber(raw[c] as number, fmt);
      const labelW = textWidth(label, labelFs) + 6;
      // Outside means to the RIGHT of the band, and the widest band already
      // reaches the edge of the plot — so on a large font there is no room out
      // there and the label was drawn past the frame. `stage-value-0` landed at
      // x = 480.0 on a 480pt frame: the top stage of the funnel showed no value
      // at all while every stage below it did, because the one band big enough
      // to matter is the one with nothing to its right.
      //
      // So "outside" has to be a placement the frame can hold. When it is not,
      // the label goes inside after all: the band that cannot fit a label beside
      // it is the WIDEST one, which is exactly the band with the most room in
      // it. A cramped label on the bar beats a missing number.
      const roomOutside = cx + w / 2 + 4 + labelW <= cfg.width;
      const fitsInside = w >= textWidth(label, labelFs) + 8 && bandH >= labelFs * 1.3;
      const inside = fitsInside || !roomOutside;
      nodes.push({
        kind: "text",
        x: inside ? cx - w / 2 : cx + w / 2 + 4,
        y: y + bandH / 2 - labelFs * 0.75,
        w: inside ? w : labelW,
        h: labelFs * 1.5,
        text: label,
        fontSize: labelFs,
        bold: true,
        color: inside ? contrastInk(fill) : style.text,
        align: inside ? "center" : "left",
        valign: "middle",
        name: `stage-value-${c}`,
      });
    }
    // Conversion rate vs the previous stage, in the gap between bands. The
    // marker follows the direction: a fixed ▾ contradicted itself on the
    // ascending (pyramid) ordering this file recommends, printing "▾ 500.0%".
    // A conversion rate DERIVED from an unknown is unknown. Either end missing
    // means no percentage rather than a percentage of zero.
    // …AND NEITHER END IS NEGATIVE. `values` is the CLAMPED array, so a stage of
    // -50 arrives here as 0 and printed "▾ 0.0%" — a conversion rate of nothing,
    // stated as measured, from a number the stage label beside it renders
    // honestly as "-50". A funnel stage is a count: negative is not a small
    // conversion, it is data that cannot have a rate taken of it. The stage
    // value still shows -50, which is what tells the user their sheet is wrong.
    if (c > 0 && values[c - 1] > 0 && raw[c] !== null && raw[c - 1] !== null && raw[c]! >= 0 && raw[c - 1]! >= 0) {
      const marker = v > values[c - 1] ? "▴ " : v < values[c - 1] ? "▾ " : "";
      // Bounded by the GAP it sits in, which nothing did.
      //
      // The box is `h: gap` and the text was drawn at `labelFs * 0.85` whatever
      // that gap turned out to be — and `gap` is `min(fs * 1.5, half the plot
      // shared between the bands)`, so on a short frame it collapses while the
      // font does not. The ink then spills out of the gap onto the bands either
      // side and lands on their `stage-value`, which is the single commonest
      // overlapping pair at small frames.
      //
      // Dropped rather than floored when the gap cannot carry a legible one: a
      // conversion rate written across the number it is comparing is worse than
      // no conversion rate, and it is the secondary text here by design.
      const convFs = bandFontSize(labelFs * 0.85, gap, 1.2);
      if (convFs <= 0) return;
      nodes.push({
        kind: "text",
        x: cx - 40,
        y: y - gap,
        w: 80,
        h: gap,
        text: `${marker}${formatPercent(v / values[c - 1], 1, false, cfg.numberFormat?.locale)}`,
        // Off the row font, not the chart font: the conversion rate is muted
        // secondary text and must stay SMALLER than the stage names beside it.
        // Shrinking only the names inverted that — at a 28pt font the names
        // came down to 14 and these stayed at 23.8, so the percentages became
        // the loudest thing on the chart. At any font that fits, labelFs is fs
        // and this is the value it always had.
        fontSize: convFs,
        color: style.mutedText,
        align: "center",
        valign: "middle",
        name: `conversion-${c}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: values.map(() => cx),
      categoryWidth: values.map((v) => Math.max(2, (v / vMax) * plot.w)),
      columnTop,
      columnValue: values,
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
