import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { titleInkBottom, firstSeriesValues, fitPlot, footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Cascade / decomposition chart: each stage's bar is a subset of the
 * previous one, read left to right (Total → Answered → With a case →
 * Solved…). Bars are top-aligned on one volume scale; the complement of
 * each split ("Dropped", "Without a case") hangs as a muted box at the
 * split point, labeled with value and % of the previous stage.
 *
 * Category syntax: "Stage | Drop label | Group header" — part 2 captions
 * the remainder box of the split INTO this stage; consecutive stages
 * sharing part 3 get one spanning header band.
 */
export function layoutCascade(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const parts = data.categories.map((c) => c.split("|").map((p) => p.trim()));
  const stages = parts.map((p) => p[0] ?? "");
  const dropLabels = parts.map((p) => p[1] ?? "");
  const groups = parts.map((p) => p[2] ?? "");
  const n = stages.length;
  const { plot: values, raw } = firstSeriesValues(data, n);
  // Valid cascades decrease, so the first stage IS the max; scaling by the
  // max keeps malformed (growing) data inside the plot instead of overflowing.
  //
  // `|| 1` guards only the all-zero case. A floor of 1 here — `Math.max(...values,
  // 1)` — also clamped any cascade whose data is below 1, and a cascade has no
  // value axis to show why: [0.09, 0.06, 0.03] drew its first stage, which is the
  // chart's whole reference, at 25.6pt of a 284pt plot.
  const v0 = Math.max(0, ...values) || 1;
  const fmt = resolveFormat(values, cfg.numberFormat);

  const titleH = titleHeight(cfg, style);
  const hasGroups = groups.some(Boolean);
  const groupH = hasGroups ? fs * 1.7 : 0;
  const plot = fitPlot(cfg, {
    x: 2,
    y: titleH + groupH + (hasGroups ? 4 : 0),
    w: cfg.width - 4,
    // The last term is the OUTSIDE drop label. A block too thin for text inside
    // it gets its caption underneath (see `outside` below), and the plot did not
    // reserve for that — so on an ordinary 480x300 cascade the caption was drawn
    // 7.5pt into the footnote's band, dark text over dark text.
    //
    // Reserved rather than clamped: clamping the label up to clear the footnote
    // puts it on its own block, and where two thin blocks end together it stacks
    // BOTH captions at the same y, which is a worse collision than the one it
    // fixes. A decomposition ends in thin blocks by its nature, so this is a row
    // the chart nearly always needs.
    h: cfg.height - titleH - groupH - (hasGroups ? 4 : 0) - footnoteH(cfg, style, decor) - 4 - fs * 1.2,
  });
  // The group band is chrome, and `fitPlot` is allowed to overrun it.
  //
  // The plot reserves `groupH` above itself, but `fitPlot` grows the plot UP
  // from the bottom edge it was given — deliberately, because that edge is the
  // baseline and moving it changes what the chart claims. On a frame too short
  // to pay for its chrome the floored plot therefore rises back through the
  // very band it reserved: at 300x60 the band occupies y 22-37 and the blocks'
  // captions land at 29.5-32, inside it.
  //
  // So the headers are DROPPED when the reservation did not survive, which is
  // the answer the radar, sunburst, tilemap and pie reservations already give.
  // A header drawn across the bars it names is not a header, and keeping the
  // bars is the right way round: they are the chart.
  const groupsFit = !hasGroups || plot.y >= titleH + groupH;
  const slotW = plot.w / Math.max(1, n);
  const barW = slotW * 0.91;
  const toH = (v: number) => (v / v0) * plot.h;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Spanning group header bands over consecutive same-group stages.
  if (groupsFit && hasGroups) {
    let start = 0;
    for (let c = 1; c <= n; c++) {
      if (c === n || groups[c] !== groups[start]) {
        if (groups[start]) {
          const x1 = plot.x + slotW * start + (slotW - barW) / 2;
          const x2 = plot.x + slotW * (c - 1) + (slotW + barW) / 2;
          nodes.push(
            { kind: "rect", x: x1, y: titleH, w: x2 - x1, h: fs * 1.5, fill: style.neutral, name: `group-${start}` },
            {
              kind: "text",
              x: x1,
              y: titleH,
              w: x2 - x1,
              h: fs * 1.5,
              // Centred in its band, so a header wider than the stages it spans
              // spilled out of both ends of the chart — 50pt on a 120pt-wide
              // cascade. Clipped to the band it names; a header that fits is
              // untouched.
              text: clipToWidth(groups[start], fs, x2 - x1, true),
              fontSize: fs,
              bold: true,
              color: contrastInk(style.neutral),
              align: "center",
              valign: "middle",
              name: `group-label-${start}`,
            },
          );
        }
        start = c;
      }
    }
  }

  const columnTop: number[] = [];
  const centers: number[] = [];
  values.forEach((v, c) => {
    const x = plot.x + slotW * c + (slotW - barW) / 2;
    const center = x + barW / 2;
    centers.push(center);
    columnTop.push(plot.y);
    const h = Math.max(2, toH(v));
    const fill = data.series[0]?.colors?.[c] ?? style.palette[c % style.palette.length];
    const ink = contrastInk(fill);
    nodes.push({ kind: "rect", x, y: plot.y, w: barW, h, fill, name: `stage-${c}` });

    // In-bar text: stage name near the top, value + % of previous centered.
    /**
     * KNOWN — and something a share can honestly be taken OF.
     *
     * Unknown in, unknown out: a blank cell used to make this a 100% drop, and
     * the null half of this gate fixed that. A NEGATIVE stage does the same
     * thing by the same route, because `values` is the CLAMPED array and -50
     * arrives here as 0: the stage got no bar and no label, and the remainder
     * box beside it read "Other: 1,000 (100.0%)" — a total wipe-out that never
     * happened, computed from a clamp, on a sheet whose only fault was a sign.
     *
     * A cascade stage is a surviving count. Negative is not a 100% drop; it is
     * a number no share can be taken of.
     */
    const known = raw[c] !== null && raw[c - 1] !== null && raw[c]! >= 0 && raw[c - 1]! >= 0;
    const pct = c > 0 && known && values[c - 1] > 0 ? values[c] / values[c - 1] : null;
    const lines = [
      { text: stages[c], y: plot.y + h * 0.18, bold: false, size: fs },
      { text: formatNumber(v, fmt), y: plot.y + h * 0.5 - fs * 0.75, bold: true, size: fs * 1.05 },
      ...(pct != null
        ? [
            {
              text: `(${formatPercent(pct, 1, false, cfg.numberFormat?.locale)})`,
              y: plot.y + h * 0.5 + fs * 0.7,
              bold: false,
              size: fs,
            },
          ]
        : []),
    ];
    // The bar-height guard below is a cheap pre-filter and is NOT the test that
    // matters: the lines sit at FRACTIONS of the bar's height, so the gap between
    // them shrinks with the bar while their ink does not. A bar could clear the
    // height guard and still stack its own name on its own value — four such
    // pairs at the DEFAULT font on a 200x150 cascade. So the real check is the
    // gap each line leaves the last one, measured rather than inferred.
    let lastCentre = -Infinity;
    for (const [i, line] of lines.entries()) {
      if (h < fs * (2.2 + i * 1.4)) break; // bar too short for more lines
      // Both are centred in a `fs * 1.4` box, so their centres carry the spacing.
      const centre = line.y + fs * 0.7;
      if (centre - lastCentre < line.size * 1.05) break;
      lastCentre = centre;
      nodes.push({
        kind: "text",
        x: x + 2,
        y: line.y,
        w: barW - 4,
        h: fs * 1.4,
        // Centred in the bar, so a stage name wider than its own bar spills out
        // of both sides — off the left of the chart for the first stage, 21pt on
        // a 120pt-wide cascade. Clipped to the bar it labels.
        text: clipToWidth(line.text, line.size, barW - 4, line.bold),
        fontSize: line.size,
        bold: line.bold,
        color: ink,
        align: "center",
        valign: "middle",
        name: `stage-label-${c}-${i}`,
      });
    }

    // Remainder box: what the previous stage lost at this split.
    // `known`, or the remainder is invented: with a blank stage clamped to 0 the
    // subtraction said the whole of the previous stage was lost and captioned it
    // "Other: 1,000 (100.0%)". If we do not know this stage, we do not know what
    // stopped here.
    if (c > 0 && known) {
      const rem = Math.max(0, values[c - 1] - v);
      if (rem > 0) {
        // The column is ONE bar split in two: the colored segment above is
        // what continues, this gray segment is what stops here. They are
        // flush, share the bar width, and their heights are exact — so the
        // block's span is identical to the previous column's continuing
        // segment, and a column can never outgrow what feeds it.
        const segY = plot.y + h;
        const segH = toH(rem);
        const remPct = values[c - 1] > 0 ? rem / values[c - 1] : null;
        const caption = dropLabels[c] || cfg.labels?.other || "Other";
        const numbers = `${formatNumber(rem, fmt)}${
          remPct != null ? ` (${formatPercent(remPct, 1, false, cfg.numberFormat?.locale)})` : ""
        }`;
        const oneLine = `${caption}: ${numbers}`;
        const ink = contrastInk(style.neutral);
        nodes.push({
          kind: "rect",
          x,
          y: segY,
          w: barW,
          h: segH,
          fill: style.neutral,
          stroke: style.background,
          strokeWidth: 0.75,
          name: `drop-${c}`,
        });
        // Labels adapt to the segment — never the other way around, and none of
        // them is drawn inside the TITLE's band.
        //
        // `fitPlot` grows the plot up from its bottom edge when the frame cannot
        // pay for its chrome, and this chart's own comment calls that overlap
        // "ugly and honest" — the bars end up over the title. What is neither is
        // a CAPTION there: at 18pt on a 300x60 frame three of them were drawn
        // across the title, at nearly the title's own size, over hairline blocks
        // a point tall. The blocks stay (they are the data); the words that
        // would be read as the title's do not.
        const fitsOneLine = textWidth(oneLine, fs * 0.9) <= barW - 6;
        const clearOfTitle = segY >= titleInkBottom(cfg, style);
        // A label under the block, centred on it — so a caption wider than the
        // block spills equally both ways, off the LEFT of the chart for the
        // first column and off the right for the last (21pt and 62pt on a
        // 120pt-wide cascade). Clipped to what the chart can hold, then the
        // centre nudged by exactly the overflow, which is the same pair of
        // moves the gantt's last tick label uses: a caption that already fits
        // is centred on its block as before.
        const outside = (text: string, name: string): SceneNode => {
          const boxW = barW + slotW * 0.09;
          const lf = fs * 0.85;
          const shown = clipToWidth(text, lf, Math.min(boxW, cfg.width));
          const half = textWidth(shown, lf) / 2;
          const mid = Math.min(Math.max(x + barW / 2, half), Math.max(half, cfg.width - half));
          return {
            kind: "text",
            x: mid - boxW / 2,
            y: segY + segH + 1,
            w: boxW,
            h: fs * 1.2,
            text: shown,
            fontSize: lf,
            color: style.text,
            align: "center",
            valign: "top",
            name,
          };
        };
        if (!clearOfTitle) {
          // nothing: the block is inside the title's band
        } else if (segH >= fs * 2.9 && !fitsOneLine) {
          // Tall enough for two lines: caption over numbers, inside.
          nodes.push(
            {
              kind: "text",
              x: x + 2,
              y: segY + segH / 2 - fs * 1.35,
              w: barW - 4,
              h: fs * 1.4,
              text: clipToWidth(caption, fs * 0.9, barW - 4),
              fontSize: fs * 0.9,
              color: ink,
              align: "center",
              valign: "middle",
              name: `drop-label-${c}`,
            },
            {
              kind: "text",
              x: x + 2,
              y: segY + segH / 2,
              w: barW - 4,
              h: fs * 1.4,
              text: clipToWidth(numbers, fs * 0.9, barW - 4),
              fontSize: fs * 0.9,
              color: ink,
              align: "center",
              valign: "middle",
              name: `drop-value-${c}`,
            },
          );
        } else if (segH >= fs * 1.3 && fitsOneLine) {
          // One comfortable line, inside.
          nodes.push({
            kind: "text",
            x: x + 2,
            y: segY,
            w: barW - 4,
            h: segH,
            text: clipToWidth(oneLine, fs * 0.9, barW - 4),
            fontSize: fs * 0.9,
            color: ink,
            align: "center",
            valign: "middle",
            name: `drop-label-${c}`,
          });
        } else if (segH >= fs * 1.3) {
          // Room for one line but the caption is long: numbers inside,
          // caption just below the block.
          nodes.push(
            {
              kind: "text",
              x: x + 2,
              y: segY,
              w: barW - 4,
              h: segH,
              text: clipToWidth(numbers, fs * 0.9, barW - 4),
              fontSize: fs * 0.9,
              color: ink,
              align: "center",
              valign: "middle",
              name: `drop-value-${c}`,
            },
            outside(caption, `drop-label-${c}`),
          );
        } else {
          // Segment too thin for any text: full label below the block.
          nodes.push(outside(oneLine, `drop-label-${c}`));
        }
      }
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: values.map(() => barW),
      columnTop,
      columnValue: values,
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
