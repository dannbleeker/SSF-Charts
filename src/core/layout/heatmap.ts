import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, parseDateToken, resolveFormat } from "../format";
import { divergingScale, lerpColor, noDataFill, sequentialScale, zoneFill } from "../color";
import { maxOf, minOf } from "../agg";
import { bandFontSize, fitPlot, footnoteH, MIN_LABEL_FS, titleHeight, titleNode, titleInkBottom } from "./frame";
import type { LayoutResult } from "./column";

/**
 * How far the gradient legend's ink reaches below its own top, in font sizes:
 * the strip is `fs * 0.9`, the end labels sit at `fs * 0.95` under it in a box
 * `fs * 1.2` tall, so the last ink is a shade over `fs * 2` down. Stated once
 * because the clamp that keeps the legend inside the frame and the band the end
 * labels are fitted to are both expressed in terms of it.
 */
const LEGEND_INK = 2.15;

/**
 * Heatmap: series = rows, categories = columns, value → color on ONE global
 * scale (comparability across rows is the whole point). Sequential scale for
 * one-signed data; diverging through white, symmetric around zero, when the
 * data spans zero. Cell labels shown when they fit; compact gradient legend.
 */
/**
 * The +/− mark for one cell, or null where a sign would be a claim the data
 * does not make.
 *
 * Zero gets NOTHING. The diverging scale paints exactly-zero its neutral
 * white, and a "+" over that white would have the two channels contradict each
 * other on the same cell — to the greyscale reader, the only one who needs the
 * mark, the colour is unreadable and the glyph would be the whole story. Zero
 * has no sign, so the honest mark is none.
 *
 * `box` is the cell AS DRAWN: sizeEncode shrinks it, and a mark sized off the
 * slot would spill across its neighbours. It also puts the largest mark on the
 * largest magnitude, which is where the sign matters most.
 *
 * Shared by the grid and calendar layouts so the two cannot drift apart.
 */
function signMark(
  v: number | null | undefined,
  box: { x: number; y: number; w: number; h: number },
  fill: string,
  name: string,
): SceneNode | null {
  if (v == null || v === 0) return null;
  const gs = Math.min(box.w, box.h) * 0.3;
  // Below this the mark is a smudge, not a sign.
  if (!(gs > 0.4)) return null;
  const ink = contrastInk(fill);
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  return v > 0
    ? { kind: "symbol", shape: "plus", cx, cy, size: gs, fill: ink, name }
    : // The minus IS the plus's horizontal arm — same bar, same weight, so the
      // pair matches by construction and needs no second preset.
      { kind: "rect", x: cx - gs, y: cy - gs / 2, w: gs * 2, h: gs, fill: ink, name };
}

export function layoutHeatmap(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const nCols = data.categories.length;
  const nRows = data.series.length;
  const fs = style.fontSize;
  const opts = cfg.heatmap ?? {};

  // Calendar layout: a single daily series with date categories becomes a
  // weekday × week grid (the GitHub-contributions view).
  const calDays = data.categories.map((c) => parseDateToken(c));
  const calendar = !!opts.calendar && data.series.length >= 1 && fitsCalendar(calDays);

  // The series the scale is read from: all of them for the matrix, whose one
  // shared scale is what makes its rows comparable, and the FIRST alone for the
  // calendar, which draws that row and no other. The scale, the diverging mode,
  // `constant` and the sign marks were all read off every row while the grid
  // drew one, so a second row on another scale washed the calendar out to its
  // palest step (14 of 14 days, a calendar that looked empty), and a negative
  // anywhere in it turned the drawn one diverging, put a red step on the legend
  // for a value nobody could see, and — under `symbols: "sign"` — marked every
  // day "+", the uniform noise `wantSign` below exists to avoid. Rows past the
  // first stay unread, as `types.ts` says of the calendar: "a single daily series".
  const scaled = calendar ? data.series.slice(0, 1) : data.series;
  const all = scaled.flatMap((s) => s.values.filter((v): v is number => v != null));
  const min = minOf(all);
  const max = maxOf(all);
  const positive = opts.color ?? style.palette[0];
  const negative = opts.negativeColor ?? style.negative;
  const mode = opts.mode && opts.mode !== "auto" ? opts.mode : min < 0 && max > 0 ? "diverging" : "sequential";
  const constant = !all.length || min === max;
  const colorOf = constant
    ? () => lerpColor(style.background, positive, 0.5)
    : mode === "diverging"
      ? divergingScale(min, max, positive, negative, style.background)
      : sequentialScale(min, max, positive, style.background);
  const fmt = resolveFormat(all, cfg.numberFormat);
  const maxAbs = Math.max(1e-9, Math.abs(min), Math.abs(max));
  const sizeEncode = !!opts.sizeEncode;
  // Sign marks only say something when the data actually has both signs. On
  // one-signed data every cell would carry the same mark — pure noise — so the
  // option stays inert there, the same way `mode` auto-detects diverging.
  const wantSign = opts.symbols === "sign" && min < 0 && max > 0;

  // Hierarchical row clustering: reorder rows by average-linkage similarity and
  // draw a dendrogram in a left gutter. Needs ≥3 rows and ≥2 columns. Capped at
  // MAX_CLUSTER_ROWS: the agglomerative sweep is ~O(rows³) and a dendrogram past a
  // few dozen leaves is unreadable anyway, so above the cap the heatmap renders in
  // input order rather than letting an authored row count spin the CPU.
  const MAX_CLUSTER_ROWS = 64;
  const clusterOn = !!opts.cluster && nRows >= 3 && nCols >= 2 && nRows <= MAX_CLUSTER_ROWS;
  const cl = clusterOn ? clusterRows(data.series.map((s) => data.categories.map((_, c) => s.values[c] ?? 0))) : null;
  const rows = cl ? cl.order.map((i) => data.series[i]) : data.series;

  if (calendar) {
    return calendarLayout(
      cfg,
      style,
      decor,
      calDays as number[],
      data.series[0].values,
      colorOf,
      min,
      max,
      constant,
      wantSign,
    );
  }

  const titleH = titleHeight(cfg, style);
  const headerH = decor.categoryAxis !== false ? fs * 1.5 : 2;
  const rowLabelW = Math.min(cfg.width * 0.28, Math.max(fs, ...data.series.map((s) => textWidth(s.name, fs))) + 8);
  const legendH = fs * 3;
  // Marginal totals reserve a strip on the right (row sums) / bottom (column sums).
  const wantRowTotals = opts.totals === "row" || opts.totals === "both";
  const wantColTotals = opts.totals === "column" || opts.totals === "both";
  const totalsW = wantRowTotals ? fs * 4 : 0;
  const totalsH = wantColTotals ? fs * 1.9 : 0;
  const dendroW = clusterOn ? fs * 4 : 0; // left gutter for the row dendrogram
  const plot = fitPlot(cfg, {
    x: dendroW + rowLabelW,
    y: titleH + headerH,
    w: cfg.width - dendroW - rowLabelW - 2 - totalsW,
    h: cfg.height - titleH - headerH - legendH - totalsH - footnoteH(cfg, style, decor),
  });
  const cw = plot.w / Math.max(1, nCols);
  const ch = plot.h / Math.max(1, nRows);

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  if (decor.categoryAxis !== false) {
    // Fitted to the COLUMN, which nothing did: the box is `cw` and the name was
    // drawn at the chart font whatever `cw` turned out to be. On a 60pt-wide
    // heatmap a column is a few points across and every header ran across its
    // neighbours — inside the frame, so no overflow gate saw it.
    //
    // Shrunk together so the row still reads as one header row, then clipped to
    // the column, then dropped as a set below the legibility floor. All three
    // are the answers this engine already gives elsewhere; what was missing was
    // any of them.
    //
    // The COLUMN was only half of it. `headerH` is `fs * 1.5` whatever the frame
    // is, so on a 300x60 chart at 32pt the title and the header band together
    // are taller than the chart and the header row was drawn 39pt below its own
    // foot. A band that cannot be paid for is not drawn: the same rule the plot,
    // the radar's ticks and the tilemap's scale ends follow.
    const widest = Math.max(0, ...data.categories.map((cat) => textWidth(cat, fs)));
    const wide = widest > cw - 2 && widest > 0 ? fs * ((cw - 2) / widest) : fs;
    // Anchored to the PLOT, not to the title. The header row was drawn at
    // `titleH` whatever `fitPlot` did with the grid, and on an 80x60 frame at
    // 18pt the plot is pushed up ABOVE that — so the headers sat below their own
    // columns, on top of the colour legend. A header names the column under it;
    // it has to move with the column.
    const headerBand = Math.max(0, Math.min(headerH, plot.y - titleH));
    const headerFs = bandFontSize(wide, headerBand, 1.5);
    if (headerFs >= MIN_LABEL_FS)
      data.categories.forEach((cat, c) => {
        nodes.push({
          kind: "text",
          x: plot.x + c * cw,
          y: plot.y - headerBand,
          w: cw,
          h: headerBand,
          text: clipToWidth(cat, headerFs, cw - 2),
          fontSize: headerFs,
          color: style.text,
          align: "center",
          valign: "middle",
          name: `col-${c}`,
        });
      });
  }

  // A row label's BOX is the row, so the boxes never overlap — but its INK is
  // `fs` tall and centred in that box, so once the font outgrows the row the
  // names are drawn through each other. Twenty of the 237 overlapping text pairs
  // a sweep found were this one label. Bound by the row it names, like every
  // other row label in this engine; at any font that already fits, nothing moves.
  const rowFs = bandFontSize(fs, ch, 1.15);
  rows.forEach((s, ri) => {
    if (rowFs > 0)
      nodes.push({
        kind: "text",
        x: dendroW,
        y: plot.y + ri * ch,
        w: rowLabelW - 4,
        h: ch,
        text: s.name,
        fontSize: rowFs,
        color: style.text,
        align: "right",
        valign: "middle",
        name: `row-${ri}`,
      });
    data.categories.forEach((_, c) => {
      const v = s.values[c];
      const fill = v == null ? noDataFill(style.background) : colorOf(v);
      const x = plot.x + c * cw;
      const y = plot.y + ri * ch;
      // The cell as DRAWN — sizeEncode shrinks it, and the sign mark has to sit
      // inside whatever is actually there.
      let box = { x, y, w: cw - 1, h: ch - 1 };
      if (sizeEncode) {
        // Centred square whose side ∝ √(|v|/maxAbs) so its area tracks magnitude.
        const frac = v == null ? 0 : Math.sqrt(Math.abs(v) / maxAbs);
        const side = Math.max(0, Math.min(cw - 1, ch - 1) * frac);
        box = { x: x + (cw - 1 - side) / 2, y: y + (ch - 1 - side) / 2, w: side, h: side };
      }
      nodes.push({ kind: "rect", ...box, fill, name: `cell-${ri}-${c}` });

      let labelled = false;
      if (v != null && decor.segmentLabels && !sizeEncode) {
        const label = formatNumber(v, fmt);
        if (cw >= textWidth(label, fs) + 4 && ch >= fs * 1.3) {
          labelled = true;
          nodes.push({
            kind: "text",
            x,
            y,
            w: cw - 1,
            h: ch - 1,
            text: label,
            fontSize: fs,
            color: contrastInk(fill),
            align: "center",
            valign: "middle",
            name: `cell-label-${ri}-${c}`,
          });
        }
      }
      // Sign marks, for the reader who cannot use the color. Only where the
      // value label is absent: the label already prints the minus sign, and a
      // second sign in the same cell is noise that would sit on top of it.
      if (wantSign && !labelled) {
        const mark = signMark(v, box, fill, `cell-sign-${ri}-${c}`);
        if (mark) nodes.push(mark);
      }
    });
  });

  // Marginal totals: neutral sum strips outside the color scale.
  const sum = (vals: (number | null)[]) => vals.reduce((a: number, v) => a + (v ?? 0), 0);
  /**
   * Whether a margin has anything to total — asked before the number is drawn.
   *
   * `sum` counts a null as zero, which is right for adding up what a row HAS
   * and wrong the moment the answer becomes a label. A row or column whose
   * every cell is blank sums to 0 and printed `row-total = "0"`, telling the
   * reader that quarter measured nothing when the truth is nobody filled the
   * cells in. The same defect as the stacked column's total, one file over —
   * and `test/blank-is-not-zero.test.ts` did not see it at first, because these
   * marginals are `heatmap: { totals }` and not `decorations.totals`.
   */
  const anyData = (vals: (number | null)[]) => vals.some((v) => typeof v === "number" && Number.isFinite(v));
  if (wantRowTotals) {
    rows.forEach((s, ri) => {
      const y = plot.y + ri * ch;
      nodes.push(
        {
          kind: "rect",
          x: plot.x + plot.w + 2,
          y,
          w: totalsW - 4,
          h: ch - 1,
          fill: zoneFill(style.background, "#f0efec"),
          name: `row-total-bg-${ri}`,
        },
        {
          kind: "text",
          x: plot.x + plot.w + 2,
          y,
          w: totalsW - 6,
          h: ch - 1,
          // The BACKGROUND still draws — the margin keeps its shape — and only
          // the figure is withheld. See `anyData`.
          text: anyData(s.values) ? formatNumber(sum(s.values), fmt) : "",
          fontSize: fs * 0.95,
          bold: true,
          color: style.text,
          align: "center",
          valign: "middle",
          name: `row-total-${ri}`,
        },
      );
    });
  }
  if (wantColTotals) {
    data.categories.forEach((_, c) => {
      const x = plot.x + c * cw;
      const column = data.series.map((s) => s.values[c]);
      const total = sum(column);
      nodes.push(
        {
          kind: "rect",
          x,
          y: plot.y + plot.h + 2,
          w: cw - 1,
          h: totalsH - 4,
          fill: zoneFill(style.background, "#f0efec"),
          name: `col-total-bg-${c}`,
        },
        {
          kind: "text",
          x,
          y: plot.y + plot.h + 2,
          w: cw - 1,
          h: totalsH - 4,
          text: anyData(column) ? formatNumber(total, fmt) : "",
          fontSize: fs * 0.95,
          bold: true,
          color: style.text,
          align: "center",
          valign: "middle",
          name: `col-total-${c}`,
        },
      );
    });
  }

  // Gradient legend: a strip of small steps with min/max (and 0) labels.
  //
  // The strip hangs off the bottom of the plot, and `fitPlot` grows the plot UP
  // from the bottom edge it was given — so on a frame that cannot pay for its
  // chrome the plot's foot is where the reservation put it and the legend below
  // it lands outside the chart: the end labels were drawn 30pt under an 80x60
  // chart at 32pt. Clamped to the last position its own ink allows, exactly as
  // the tilemap's legend already is, and the labels are then fitted to whatever
  // band is left and dropped below the legibility floor. An unlabelled gradient
  // says less; a labelled one drawn off the chart says it somewhere else.
  const foot = footnoteH(cfg, style, decor);
  // FLOORED AT THE TITLE'S INK, not at zero. The clamp above keeps the legend
  // on the canvas; `Math.max(0, …)` then put it at y=0 on a chart whose plot has
  // been squeezed to the ceiling, and y=0 is where the title is — a 300x60
  // heatmap at 18pt drew its colour scale across its own name. The fifth time
  // this engine has clamped a label to zero and found the title already there.
  //
  // `legendClearOfGrid` below is unchanged and still does the other half: where
  // the floor pushes the legend above the grid, there is nowhere for it and it
  // is not drawn at all.
  const ly = Math.max(
    titleInkBottom(cfg, style),
    Math.min(plot.y + plot.h + totalsH + fs * 0.6, cfg.height - foot - fs * LEGEND_INK),
  );
  /**
   * …and only while that clamp has not walked the legend back INTO the chart.
   *
   * On an 80x60 frame at 18pt the only position left on the canvas is above the
   * plot's own top edge, so the end labels were drawn through the column headers
   * — a colour key over the names of the columns it is meant to explain. The
   * grid and its headers are the chart; the key is chrome, and chrome that
   * cannot be paid for is not drawn.
   */
  const legendClearOfGrid = ly >= plot.y;
  if (constant && legendClearOfGrid) {
    const swatchFs = bandFontSize(fs * 0.9, cfg.height - foot - ly, 1.35);
    if (swatchFs > 0)
      nodes.push(
        { kind: "rect", x: plot.x, y: ly, w: fs * 1.6, h: fs * 0.9, fill: colorOf(min), name: "legend-swatch" },
        {
          kind: "text",
          x: plot.x + fs * 1.9,
          y: ly - fs * 0.25,
          w: fs * 8,
          h: fs * 1.4,
          text: all.length ? formatNumber(min, fmt) : "no data",
          fontSize: swatchFs,
          color: style.mutedText,
          align: "left",
          valign: "middle",
          name: "legend-min",
        },
      );
  } else if (legendClearOfGrid) {
    const lw = Math.min(plot.w * 0.5, fs * 14);
    const steps = 24;
    for (let i = 0; i < steps; i++) {
      const v = min + ((max - min) * i) / (steps - 1);
      nodes.push({
        kind: "rect",
        x: plot.x + (lw / steps) * i,
        y: ly,
        w: lw / steps + 0.5,
        h: fs * 0.9,
        fill: colorOf(v),
        name: `legend-step-${i}`,
      });
    }
    // Both ends or neither, for the reason the tilemap's pair gives: a gradient
    // labelled at one end says the wrong thing.
    //
    // Fitted on BOTH axes, which the tilemap's pair has always been and this one
    // was not: each end owns half the bar and is anchored to its outer edge, so
    // two numbers wider than `lw / 2` meet in the middle — 60x300 at 18pt, where
    // half the bar is ten points and the numbers are fifteen.
    const endBand = cfg.height - foot - (ly + fs * 0.95);
    const ends = [formatNumber(min, fmt), formatNumber(max, fmt)];
    const widestEnd = Math.max(...ends.map((t) => textWidth(t, fs * 0.85)));
    const endFs = Math.min(
      bandFontSize(fs * 0.85, endBand, 1.21),
      widestEnd > 0 ? (fs * 0.85 * (lw / 2)) / widestEnd : fs * 0.85,
    );
    if (endFs >= MIN_LABEL_FS)
      nodes.push(
        {
          kind: "text",
          x: plot.x,
          y: ly + fs * 0.95,
          w: lw / 2,
          h: Math.min(fs * 1.2, endBand),
          text: formatNumber(min, fmt),
          fontSize: endFs,
          color: style.mutedText,
          align: "left",
          valign: "top",
          name: "legend-min",
        },
        {
          kind: "text",
          x: plot.x + lw / 2,
          y: ly + fs * 0.95,
          w: lw / 2,
          h: Math.min(fs * 1.2, endBand),
          text: formatNumber(max, fmt),
          fontSize: endFs,
          color: style.mutedText,
          align: "right",
          valign: "top",
          name: "legend-max",
        },
      );
    // Only mark zero when it actually lies inside the range. A FORCED diverging
    // mode over single-signed data (auto-diverging can't reach here) would place
    // the "0" tick off the end of the legend strip.
    if (mode === "diverging" && min < 0 && max > 0) {
      const zx = plot.x + ((0 - min) / (max - min)) * lw;
      /**
       * AND ONLY WHEN IT CLEARS THE TWO ENDS, which it did not.
       *
       * The zero tick is placed proportionally along the strip, so a range that
       * is nearly all negative puts it hard against the max label and one that
       * is nearly all positive puts it on the min — 14 pairs in the variant
       * sweep. Both ends are placed first and both are worth more: they say what
       * the colours MEAN, where the zero tick only says where the middle is, and
       * a diverging scale shows that in its own neutral band anyway.
       *
       * Measured against the ends' INK, not their boxes. Each end label is given
       * half the strip to sit in and uses as much of it as the number needs, so
       * a box test would refuse the tick on almost every chart.
       */
      const zeroInk = { lo: zx - textWidth("0", fs * 0.85) / 2, hi: zx + textWidth("0", fs * 0.85) / 2 };
      const minInk = plot.x + textWidth(formatNumber(min, fmt), endFs);
      const maxInk = plot.x + lw - textWidth(formatNumber(max, fmt), endFs);
      if (zeroInk.lo > minInk + 2 && zeroInk.hi < maxInk - 2)
        nodes.push({
          kind: "text",
          x: zx - fs,
          y: ly + fs * 0.95,
          w: fs * 2,
          h: fs * 1.2,
          text: "0",
          fontSize: fs * 0.85,
          color: style.mutedText,
          align: "center",
          valign: "top",
          name: "legend-zero",
        });
    }
  }

  // Row dendrogram in the left gutter (leaves next to the labels, root at left).
  if (cl) {
    const maxH = cl.maxH || 1;
    const pos = new Map<number, number>();
    cl.order.forEach((leaf, idx) => pos.set(leaf, idx));
    const leafX = dendroW - 3;
    const nodeX = (h: number) => Math.max(2, leafX - (leafX - 2) * (h / maxH));
    const draw = (nd: ClusterNode): number => {
      if (nd.leaf != null) return plot.y + (pos.get(nd.leaf)! + 0.5) * ch;
      const yl = draw(nd.left!);
      const yr = draw(nd.right!);
      const x = nodeX(nd.height);
      const xl = nd.left!.leaf != null ? leafX : nodeX(nd.left!.height);
      const xr = nd.right!.leaf != null ? leafX : nodeX(nd.right!.height);
      nodes.push(
        { kind: "line", x1: x, y1: yl, x2: xl, y2: yl, stroke: style.mutedText, strokeWidth: 0.75, name: "dendro-h" },
        { kind: "line", x1: x, y1: yr, x2: xr, y2: yr, stroke: style.mutedText, strokeWidth: 0.75, name: "dendro-h" },
        { kind: "line", x1: x, y1: yl, x2: x, y2: yr, stroke: style.mutedText, strokeWidth: 0.75, name: "dendro-v" },
      );
      return (yl + yr) / 2;
    };
    draw(cl.root);
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.x + (c + 0.5) * cw),
      categoryWidth: data.categories.map(() => cw),
      columnTop: data.categories.map(() => plot.y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}

/** A node in the row-clustering tree: a leaf (row index) or an internal merge. */
interface ClusterNode {
  height: number;
  members: number[];
  leaf?: number;
  left?: ClusterNode;
  right?: ClusterNode;
}

/**
 * Agglomerative average-linkage clustering of row vectors (Euclidean distance).
 * Returns the leaf order (for reordering rows), the merge tree, and the tallest
 * merge height (for scaling the dendrogram).
 */
function clusterRows(vecs: number[][]): { order: number[]; root: ClusterNode; maxH: number } {
  const dist = (a: number[], b: number[]) => Math.sqrt(a.reduce((s, v, k) => s + (v - b[k]) ** 2, 0));
  // Precompute the row-to-row distance matrix ONCE. The linkage evaluation below
  // re-reads the same leaf-pair distances on every merge scan (an O(rows³) sweep);
  // recomputing each Euclidean distance there made every read cost another O(cols).
  // `dist` is symmetric bit-for-bit — (v-b)² === (b-v)² and the reduce runs over the
  // same indices — so caching D[i][j] leaves the linkage sums byte-identical.
  const n = vecs.length;
  const D: number[][] = vecs.map(() => new Array<number>(n));
  for (let i = 0; i < n; i++) {
    D[i][i] = 0;
    for (let j = i + 1; j < n; j++) {
      const d = dist(vecs[i], vecs[j]);
      D[i][j] = d;
      D[j][i] = d;
    }
  }
  let clusters: ClusterNode[] = vecs.map((_, i) => ({ height: 0, members: [i], leaf: i }));
  const linkage = (A: ClusterNode, B: ClusterNode) => {
    let s = 0;
    for (const i of A.members) for (const j of B.members) s += D[i][j];
    return s / (A.members.length * B.members.length);
  };
  let maxH = 0;
  while (clusters.length > 1) {
    let bi = 0;
    let bj = 1;
    let bd = Infinity;
    for (let i = 0; i < clusters.length; i++) {
      for (let j = i + 1; j < clusters.length; j++) {
        const d = linkage(clusters[i], clusters[j]);
        if (d < bd) {
          bd = d;
          bi = i;
          bj = j;
        }
      }
    }
    const merged: ClusterNode = {
      height: bd,
      members: [...clusters[bi].members, ...clusters[bj].members],
      left: clusters[bi],
      right: clusters[bj],
    };
    maxH = Math.max(maxH, bd);
    clusters = clusters.filter((_, k) => k !== bi && k !== bj);
    clusters.push(merged);
  }
  const order: number[] = [];
  const walk = (n: ClusterNode) => {
    if (n.leaf != null) order.push(n.leaf);
    else {
      walk(n.left!);
      walk(n.right!);
    }
  };
  walk(clusters[0]);
  return { order, root: clusters[0], maxH };
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

/**
 * The most week columns a calendar is drawn with: as many as the engine allows
 * CATEGORIES (`MAX_CATEGORIES` in chart.ts, 4096 — seventy-eight years), because
 * a calendar that spans more weeks than it could have days is not a calendar of
 * its data any more.
 *
 * The grid is as wide as its two farthest dates are apart, and the month-label
 * loop walks every week of it. One mistyped year (`20260-01-15`, or a
 * `9999-12-31` sentinel) among thirty real January days built a scene of 216,000
 * nodes and a 25MB SVG — 94,000 for the sentinel — with the thirty days drawn
 * at 0.0004pt each; two dates at the limits of `Date` made 6.5 million and took
 * the process out of memory. `parseDateToken` reads a five-digit or signed year
 * by design, so nothing upstream bounds the span.
 */
const MAX_CALENDAR_WEEKS = 4096;

/**
 * Do these categories lay out as ONE calendar: every one a date, and the dates no
 * more than `MAX_CALENDAR_WEEKS` apart? Past the cap `layoutHeatmap` reads the
 * chart as it reads one whose categories are not dates — the ordinary grid, one
 * column per category, which is bounded by the categories it was given and shows
 * the stray date as the column it is.
 */
function fitsCalendar(days: (number | null)[]): boolean {
  if (!days.every((d): d is number => d != null)) return false;
  return Math.ceil((Math.max(...days) - Math.min(...days)) / 7) + 1 <= MAX_CALENDAR_WEEKS;
}

/**
 * Calendar heatmap: a weekday (row) × week (column) grid. Days-since-epoch
 * map to a weekday (Mon=0 … Sun=6, since the epoch was a Thursday) and a week
 * index; month labels sit above the week where each month first appears.
 */
function calendarLayout(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  days: number[],
  values: (number | null)[],
  colorOf: (v: number) => string,
  min: number,
  max: number,
  constant: boolean,
  wantSign: boolean,
): LayoutResult {
  const fs = style.fontSize;
  const weekdayOf = (d: number) => (((d + 3) % 7) + 7) % 7; // Mon=0 … Sun=6
  const minDay = Math.min(...days);
  const maxDay = Math.max(...days);
  const weekStart = minDay - weekdayOf(minDay); // Monday of the first week
  const weekOf = (d: number) => Math.floor((d - weekStart) / 7);
  const nWeeks = weekOf(maxDay) + 1;

  const titleH = titleHeight(cfg, style);
  const wdLabelW = fs * 2.2;
  const monthH = fs * 1.4;
  const legendH = fs * 2.2;
  const gridX = wdLabelW;
  const gridY = titleH + monthH + 2;
  const availW = cfg.width - gridX - 4;
  const availH = cfg.height - gridY - legendH - footnoteH(cfg, style, decor);
  // The 4pt floor is a readability minimum, not a licence to leave the canvas.
  // `nWeeks` is unbounded — CLAUDE.md names "a multi-year daily calendar
  // heatmap is ~1-4k days" as a real chart shape — so past ~118 week columns the
  // floor took over from the fit and the grid marched off the right edge: a
  // four-year calendar drew 377pt beyond a 480pt canvas. A cell under 4pt is
  // hard to read; a cell outside the chart is not drawn at all.
  const weekW = availW / Math.max(1, nWeeks);
  const cell = Math.min(Math.max(4, Math.min(weekW, availH / 7)), weekW);

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Weekday labels (every other row to reduce clutter).
  for (let r = 0; r < 7; r++) {
    if (r % 2 === 0) {
      nodes.push({
        kind: "text",
        x: 0,
        y: gridY + r * cell,
        w: wdLabelW - 3,
        h: cell,
        text: WEEKDAYS[r],
        fontSize: fs * 0.8,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: `weekday-${r}`,
      });
    }
  }
  // Month labels above the week where each month first appears.
  let lastMonth = -1;
  for (let w = 0; w < nWeeks; w++) {
    const mon = new Date((weekStart + w * 7) * 86400000).getUTCMonth();
    if (mon !== lastMonth) {
      nodes.push({
        kind: "text",
        x: gridX + w * cell,
        y: titleH,
        w: cell * 5,
        h: monthH,
        text: MONTHS[mon],
        fontSize: fs * 0.85,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: `month-${w}`,
      });
      lastMonth = mon;
    }
  }
  // Day cells.
  days.forEach((d, i) => {
    const v = values[i];
    const x = gridX + weekOf(d) * cell;
    const y = gridY + weekdayOf(d) * cell;
    const fill = v == null ? noDataFill(style.background) : colorOf(v);
    const box = { x, y, w: cell - 1.5, h: cell - 1.5 };
    nodes.push({ kind: "rect", ...box, fill, name: `cell-${i}` });
    // A calendar never draws value labels, so colour is its ONLY sign carrier —
    // which makes it the layout that needs the marks most, not one to skip.
    if (wantSign) {
      const mark = signMark(v, box, fill, `cell-sign-${i}`);
      if (mark) nodes.push(mark);
    }
  });
  // Compact gradient legend (Less → More).
  const ly = gridY + 7 * cell + fs * 0.5;
  if (!constant) {
    const steps = 12;
    // The legend band is a CONSTANT reserve (legendH) while `cell` grows as the
    // range shortens, so an unclamped `cell * 0.8` swatch fell out of the frame
    // for any calendar under ~13 week-columns. Clamp to the room actually left
    // below the grid and to the width the strip plus its "More" label may use.
    const sw = Math.max(
      3,
      Math.min(
        cell * 0.8,
        cfg.height - footnoteH(cfg, style, decor) - ly - 1,
        (cfg.width - gridX - fs * 3.4) / steps - 1,
      ),
    );
    nodes.push({
      kind: "text",
      x: gridX - fs * 3,
      y: ly,
      w: fs * 2.6,
      h: sw,
      text: "Less",
      fontSize: fs * 0.8,
      color: style.mutedText,
      align: "right",
      valign: "middle",
      name: "legend-less",
    });
    for (let i = 0; i < steps; i++) {
      nodes.push({
        kind: "rect",
        x: gridX + i * (sw + 1),
        y: ly,
        w: sw,
        h: sw,
        fill: colorOf(min + ((max - min) * i) / (steps - 1)),
        name: `legend-step-${i}`,
      });
    }
    nodes.push({
      kind: "text",
      x: gridX + steps * (sw + 1) + 2,
      y: ly,
      w: fs * 3,
      h: sw,
      text: "More",
      fontSize: fs * 0.8,
      color: style.mutedText,
      align: "left",
      valign: "middle",
      name: "legend-more",
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: days.map((d) => gridX + weekOf(d) * cell + cell / 2),
      categoryWidth: days.map(() => cell),
      columnTop: days.map(() => gridY),
      columnValue: values.map((v) => v ?? 0),
      baselineY: gridY + 7 * cell,
      plot: { x: gridX, y: gridY, w: nWeeks * cell, h: 7 * cell },
    },
  };
}
