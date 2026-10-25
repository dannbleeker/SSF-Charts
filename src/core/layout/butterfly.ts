import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { niceTicks } from "../format";
import { legendRow, type LayoutResult, type LegendEntry } from "./column";
import { MIN_PLOT_SIDE, bandFontSize, fitPlot, footnoteH, legendRowCount, titleHeight, titleNode } from "./frame";

/**
 * Butterfly (tornado) chart: think-cell models this as two bar charts placed
 * back-to-back, one rotated 180°, sharing the same scale. Uses the first two
 * series; category labels sit in the center gutter.
 */
export function layoutButterfly(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  // Flanks: with butterfly.split, the first `split` series stack on the left
  // and the rest on the right; otherwise the classic series 0 / series 1.
  const withIdx = data.series.map((s, si) => ({ s, si }));
  const split = cfg.butterfly?.split;
  const leftSeries = split != null ? withIdx.slice(0, split) : withIdx.slice(0, 1);
  const rightSeries = split != null ? withIdx.slice(split) : withIdx.slice(1, 2);
  const stacked = leftSeries.length > 1 || rightSeries.length > 1;
  const stackSum = (series: typeof withIdx, c: number) =>
    series.reduce((a, { s }) => a + Math.abs(s.values[c] ?? 0), 0);
  const signedSum = (series: typeof withIdx, c: number) => series.reduce((a, { s }) => a + (s.values[c] ?? 0), 0);

  const titleH = titleHeight(cfg, style);
  // Room for outside value labels on each flank. Scaled down below, with the
  // category gutter, when the two together would leave the BARS nothing.
  const valueW0 = fs * 3.4;
  const gutterW0 = Math.min(cfg.width * 0.3, Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 12);
  // The chrome — a value strip on each flank and the category gutter down the
  // middle — is priced in font sizes and label widths, so on a small frame it
  // takes the chart: 84 of 120 points at a thumbnail, leaving both sets of bars
  // 36 between them and the longest one 18pt. The bars are what a butterfly is
  // read by, so the chrome is scaled to fit inside HALF the width, the same way
  // the gantt scales its three text gutters into a budget. At any size where it
  // already fits, the scale is 1 and nothing moves.
  const chromeBudget = cfg.width * 0.5;
  const chrome = valueW0 * 2 + gutterW0;
  const chromeScale = chrome > chromeBudget ? chromeBudget / chrome : 1;
  const valueW = valueW0 * chromeScale;
  const gutterW = gutterW0 * chromeScale;
  // Stacked flanks legend the whole series set across the top; it wraps (shared
  // walk), so the header band has to be as tall as the rows it will occupy.
  const legendEntries: LegendEntry[] = stacked
    ? [...leftSeries, ...rightSeries].map(({ s, si }) => ({
        label: s.name,
        color: seriesColor(style, si, s.color),
        name: `legend-${si}`,
      }))
    : [];
  // Counted from `valueW` and drawn from `plot.x` below: the same place, to within
  // the last bit of a float, because `fitPlot` only moves a plot it has to floor
  // and the chrome budget above keeps the plot at least half the chart wide. Not
  // unified into one variable on purpose — `plot.x` is `valueW + plot.w - plot.w`
  // and differs from it in the last digit on 3,375 of 4,305 widths and fonts
  // tried, which would move every stacked-flank legend by that much in the deck.
  const legendRows = legendRowCount(
    legendEntries.map((e) => e.label),
    fs,
    valueW,
    cfg.width - 4,
  );
  const footH = footnoteH(cfg, style, decor);
  /**
   * The stacked legend is drawn only if the frame can pay for ALL of its rows.
   *
   * Reserving them was never the gap — the rows counted and the rows drawn agree,
   * to the row, in every cell the sweep flagged. What nothing asked is whether the
   * reservation fits the frame. The band is priced in rows of `fs * 1.6`, so a
   * narrow chart wraps three names onto three rows (48pt at the default font, 86
   * at 18) and the title above it takes another 22 to 35: on an 80x60 chart that
   * is 70pt of chrome for 60pt of chart, and the third row is drawn 8pt below the
   * foot of it (57.6 at 18pt, 27.6 on a 120x90) — six of the 32 cells of the
   * `butterfly.split` sweep, none of them at 480x300 or 960x540.
   *
   * A fit and a clamp, and the clamp wins: `fitPlot` floors the plot at
   * `MIN_PLOT_SIDE` and grows it back UP from the bottom edge, so the plot rises
   * through the legend that was reserved above it while the legend's rows keep
   * walking down. Chrome that cannot be paid for is not drawn — the answer the
   * tile map's glyph key, the scatter's group legend and the radar's key give. The
   * test is the one the value axis below already makes (the plot it leaves must
   * reach `MIN_PLOT_SIDE`), taken with the footnote subtracted as the plot itself
   * subtracts it, so a legend that is kept is a legend whose every row is on the
   * chart. A chart whose legend leaves a plot keeps it: over a grid of 11,020
   * widths, heights, fonts and series counts, the cells whose legend or bars moved
   * are exactly the ones that lost their legend.
   */
  const legendFits = stacked && cfg.height - titleH - fs * 1.6 * legendRows - 6 - footH >= MIN_PLOT_SIDE;
  // The classic two-series butterfly reserves one row for its two headers; a
  // stacked one reserves the rows it will draw, or none when it draws no legend.
  const headerH = stacked ? (legendFits ? fs * 1.6 * Math.max(1, legendRows) : 0) : fs * 1.6;
  /**
   * A value axis reserves a strip at the bottom for tick labels on both flanks —
   * and only when the frame can pay for it.
   *
   * `fitPlot` floors the plot's height and grows it back UP from the bottom edge
   * it was given, so on a frame too short for its own chrome `plot.y + plot.h`
   * runs past where this strip was meant to start, and the ticks were drawn
   * BELOW the chart: y 68-83 on a 60pt-tall frame at an 18pt font, y 103-131 on
   * a 90pt one at 32pt. Off the chart is off the slide's chart and onto whatever
   * sits under it, since neither PowerPoint renderer clips.
   *
   * Dropped rather than squeezed, which is the answer every other reservation in
   * this engine gives when it cannot be met. The draw below reads THIS value
   * rather than `decor.valueAxis`, so the strip that is reserved and the strip
   * that is drawn cannot come apart.
   */
  const axisH = decor.valueAxis && cfg.height - titleH - headerH - 6 - fs * 1.5 >= MIN_PLOT_SIDE ? fs * 1.5 : 0;
  const plot = fitPlot(cfg, {
    x: valueW,
    y: titleH + headerH + 2,
    w: Math.max(0, cfg.width - valueW * 2),
    // Floor at 0 for the same reason halfW is floored below: a chart too short
    // for its title/header/axis chrome would otherwise give every bar rect a
    // NEGATIVE height, which SVG drops and PowerPoint clamps to a sliver.
    //
    // AND THE FOOTNOTE IS CHROME TOO, which this was the one cartesian kind not
    // to say. `footnoteNode` is added centrally, so a butterfly with a footnote
    // drew one and reserved nothing for it: the bottom rows' names and values
    // were laid over the source line. 22 pairs in the variant sweep, all of them
    // this kind. Every sibling that builds its own plot — the funnel, the
    // waffle, the treemap, the gantt — already subtracts it here.
    h: Math.max(0, cfg.height - titleH - headerH - 6 - axisH - footnoteH(cfg, style, decor)),
  });
  // Floor at 0: a very narrow frame can drive plot.w below the gutter, which
  // would give the header texts and bar rects negative widths.
  const halfW = Math.max(0, (plot.w - gutterW) / 2);
  const leftEdge = plot.x + halfW; // right edge of the left half
  const rightEdge = leftEdge + gutterW; // left edge of the right half

  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null)).map((v) => Math.abs(v));
  const sums = data.categories.flatMap((_, c) => [stackSum(leftSeries, c), stackSum(rightSeries, c)]);
  // Floored at 0, not 1. `niceTicks(0, 0)` already answers [0, 1], so an all-zero
  // chart has its guard; a floor of 1 on top forced every chart whose data is
  // BELOW 1 onto a 0..1 axis. Shares and rates typed as 0.08 drew their longest
  // bar at 22pt where the same numbers as 8 drew 149 — the unit set the picture.
  const ticks = niceTicks(0, Math.max(0, ...sums), 4);
  const max = ticks[ticks.length - 1];
  const fmt = resolveFormat(all, cfg.numberFormat);
  const qOf = (v: number) => (Math.abs(v) / max) * halfW;

  const slotH = plot.h / Math.max(1, n);
  // A value label is centred on its row in a box `fs * 1.5` tall, so once the
  // font outgrows the row pitch the labels overlap each OTHER and the last one
  // leaves the plot — 10.7pt past a 200x150 frame at a 32pt font, and colliding
  // at any frame size once the font is big enough for the row count. Bound by
  // the row it labels, which is the same thing that stops both. Last resort: at
  // any font that already fits its row this is `fs` and nothing moves.
  const rowFs = bandFontSize(fs, slotH, 1.5);
  const barH = slotH * (2 / 3);

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  if (!stacked) {
    // Series headers above each half (classic two-series butterfly).
    (
      [
        [leftSeries[0], plot.x, leftEdge],
        [rightSeries[0], rightEdge, plot.x + plot.w],
      ] as const
    ).forEach(([entry, x0, x1], i) => {
      // Fitted to the HALF it names. The box is `x1 - x0` and the name was drawn
      // at the chart font whatever that came to — so on a 60pt-wide butterfly
      // each half is about twenty points and the two series names met in the
      // middle. Shrunk to the room, floored at 5 the way this file's category
      // gutter already is, and dropped below that rather than drawn across the
      // other series' header.
      //
      // The HALF was only one axis of it. The band is `fs * 1.6` whatever the
      // frame is, so a title and a header together outgrew a 60pt-tall chart at
      // 24pt and the two names were drawn 41pt below its foot — centred in a box
      // that starts inside the chart and ends outside it. Fitted to the band the
      // frame actually leaves as well, and dropped when that band cannot carry a
      // readable name: the same shrink-then-drop the heatmap's column headers
      // and the tilemap's legend take.
      const half = Math.max(0, x1 - x0);
      const name = entry?.s.name ?? "";
      // AND THE FOOTNOTE, for the same reason the plot subtracts it: this band is
      // what is left of the chart above the bars, and a footnote is drawn at the
      // foot whether or not anything reserved room for it. On an 80x60 butterfly
      // at 18pt the two names and the source line were laid over each other.
      const band = Math.max(0, Math.min(headerH, cfg.height - titleH - footnoteH(cfg, style, decor)));
      let hf = bandFontSize(fs, band, 1.6);
      if (!hf) return;
      while (hf > 5 && textWidth(name, hf) > half - 2) hf -= 0.5;
      if (textWidth(name, hf) > half - 2) return;
      nodes.push({
        kind: "text",
        x: x0,
        y: titleH,
        w: half,
        h: band,
        text: name,
        fontSize: hf,
        bold: true,
        color: seriesColor(style, entry?.si ?? i, entry?.s.color),
        align: "center",
        valign: "middle",
        name: `header-${i}`,
      });
    });
  } else if (legendFits) {
    // Stacked flanks: one legend of all series across the top, via the shared
    // wrapping row — the forked single-row walk marched chips off the canvas.
    // Gated by the predicate that reserved its rows (see `legendFits`).
    nodes.push(...legendRow(cfg, style, plot.x, titleH, { maxX: cfg.width - 4, entries: legendEntries }));
  }

  // Value gridlines mirrored on both flanks, drawn behind the bars.
  if (decor.gridlines) {
    for (const tk of ticks) {
      if (tk <= 0) continue;
      const q = qOf(tk);
      for (const x of [leftEdge - q, rightEdge + q]) {
        nodes.push({
          kind: "line",
          x1: x,
          y1: plot.y,
          x2: x,
          y2: plot.y + plot.h,
          stroke: style.gridline,
          strokeWidth: 1,
          name: `gridline-${tk}`,
        });
      }
    }
  }

  // The gutter holds the category names, and it may have been scaled down above
  // to keep the bars a chart — so the names have to come with it or they spill
  // onto the bars they are naming (they are CENTRED in the gutter, so a name
  // wider than it overhangs both flanks). Shrunk together so every row reads at
  // one size, then clipped; at any size where the names already fit, this is
  // `fs` and nothing moves.
  const catFs = (() => {
    let f = fs;
    while (f > 5 && data.categories.some((c) => textWidth(c, f) > gutterW)) f -= 0.5;
    return f;
  })();

  // A category name is bounded by BOTH the gutter it is centred in and the row it
  // names: `catFs` is the width, `rowFs` the height. Fitting only the width left
  // the names overlapping each other vertically at a big font, which is the same
  // defect as the value labels beside them and wants the same bound.
  const nameFs = rowFs > 0 ? Math.min(catFs, rowFs) : 0;

  const columnTop: number[] = [];
  for (let c = 0; c < n; c++) {
    const cy = plot.y + slotH * (c + 0.5);
    columnTop.push(cy - barH / 2);
    // Category label in the center gutter. Dropped when the row cannot carry a
    // legible one, rather than drawn at a size OOXML will not accept.
    if (nameFs > 0)
      nodes.push({
        kind: "text",
        x: leftEdge,
        y: cy - nameFs * 0.75,
        w: gutterW,
        h: nameFs * 1.5,
        text: clipToWidth(data.categories[c], nameFs, gutterW),
        fontSize: nameFs,
        color: style.text,
        align: "center",
        valign: "middle",
        name: `category-${c}`,
      });
    const drawSide = (series: typeof withIdx, dir: -1 | 1, edge: number) => {
      let offset = 0;
      for (const { s, si } of series) {
        const v = s.values[c];
        if (v == null) continue;
        const len = qOf(v);
        const x = dir < 0 ? edge - offset - len : edge + offset;
        const fill = seriesColor(style, si, s.color);
        nodes.push({ kind: "rect", x, y: cy - barH / 2, w: len, h: barH, fill, name: `seg-${si}-${c}` });
        if (decor.segmentLabels) {
          const label = formatNumber(v, fmt);
          const single = series.length === 1;
          const inside = len >= textWidth(label, fs) + 4 && barH >= fs * 1.25;
          // Stacked segments only label when the value fits inside; single
          // flanks fall back to an outside label (classic behaviour).
          if ((inside || single) && rowFs > 0) {
            nodes.push({
              kind: "text",
              x: inside ? x : dir < 0 ? x - fs * 3.4 - 2 : x + len + 2,
              y: cy - rowFs * 0.75,
              w: inside ? len : fs * 3.4,
              h: rowFs * 1.5,
              text: label,
              fontSize: rowFs,
              color: inside ? contrastInk(fill) : style.text,
              align: inside ? "center" : dir < 0 ? "right" : "left",
              valign: "middle",
              name: `label-${si}-${c}`,
            });
          }
        }
        offset += len;
      }
    };
    drawSide(leftSeries, -1, leftEdge);
    drawSide(rightSeries, 1, rightEdge);
  }

  // Center axis lines flanking the gutter.
  for (const x of [leftEdge, rightEdge]) {
    nodes.push({
      kind: "line",
      x1: x,
      y1: plot.y,
      x2: x,
      y2: plot.y + plot.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
  }

  // Value tick labels on both flanks, in the reserved bottom strip — gated on the
  // reservation, not on the flag, so a frame that could not pay for the strip
  // does not get the labels that live in it.
  if (axisH > 0) {
    const ty = plot.y + plot.h + 1;
    for (const tk of ticks) {
      const q = qOf(tk);
      const label = formatNumber(tk, fmt);
      // 0 sits at the inner edges (the gutter sides); other ticks mirror outward.
      const xs = tk === 0 ? [leftEdge, rightEdge] : [leftEdge - q, rightEdge + q];
      xs.forEach((x, side) => {
        nodes.push({
          kind: "text",
          x: x - valueW / 2,
          y: ty,
          w: valueW,
          h: axisH,
          text: label,
          fontSize: fs * 0.85,
          color: style.mutedText,
          align: "center",
          valign: "middle",
          name: `tick-${tk}-${side === 0 ? "l" : "r"}`,
        });
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.y + slotH * (c + 0.5)),
      categoryWidth: data.categories.map(() => barH),
      columnTop,
      columnValue: data.categories.map((_, c) => signedSum(rightSeries, c) - signedSum(leftSeries, c)),
      baselineY: leftEdge,
      plot,
    },
  };
}
