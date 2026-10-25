import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, resolveFormat } from "../format";
import { maxOf, minOf } from "../agg";
import { seriesColor } from "../style";
import {
  titleInkBottom,
  aboveMarkFontSize,
  bandFontSize,
  baselineNode,
  breakMarkerNodes,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  valueScale,
  type Frame,
} from "./frame";
import { horizontalChrome, type LayoutResult } from "./column";

/**
 * think-cell style waterfall: deltas build a running total; categories marked
 * as totals ("e" cells in the datasheet) draw a full bar from the baseline to
 * the running total. Dashed connectors join consecutive bar tops.
 *
 * Stacked waterfall: with several series, each column stacks its per-series
 * contributions from the incoming level and the running total moves by the
 * column sum. Set cfg.horizontal for the rotated (bar) variant.
 */
export interface Seg {
  series: number;
  from: number;
  to: number;
  value: number;
}
export interface Bar {
  segs: Seg[];
  isTotal: boolean;
  value: number;
  level: number;
}

/**
 * Walk the bridge: one bar per category, each carrying its stacked segments and
 * the running level it leaves behind.
 *
 * The single source of truth for what a waterfall draws. It used to be inlined
 * here and re-implemented, differently, by the value extent — which walked only
 * series[0] and ignored spacers, so a stacked bridge's extent stopped at the
 * first series' total. Anything that needs to know how high a waterfall reaches
 * calls this.
 */
/** Category → the chain column its detail group decomposes. */
export function detailParents(cfg: ChartConfig): Map<number, number> {
  const map = new Map<number, number>();
  for (const g of cfg.waterfall?.detailGroups ?? []) {
    for (const i of g.indices) map.set(i, g.of);
  }
  return map;
}

export function waterfallChain(cfg: ChartConfig): Bar[] {
  const { data } = cfg;
  const n = data.categories.length;
  const nSeries = Math.max(1, data.series.length);
  const totalSet = new Set(cfg.waterfall?.totalIndices ?? []);
  const spacerSet = new Set(cfg.waterfall?.spacerIndices ?? []);
  const detailOf = detailParents(cfg);
  const bars: Bar[] = [];
  let running = 0;
  for (let c = 0; c < n; c++) {
    if (spacerSet.has(c)) {
      // Blank grouping gap: no bar, running total carries across unchanged.
      bars.push({ segs: [], isTotal: false, value: 0, level: running });
      continue;
    }
    const parent = detailOf.get(c);
    if (parent != null && bars[parent]) {
      // "Of which": decompose the parent's delta as a sub-bridge rising from
      // the parent's own base, so the group reads as that column taken apart
      // rather than as more steps in the walk. Off the chain — `running` is
      // untouched, exactly like a total, so every downstream total and the
      // final one stay correct whatever these add up to.
      const p = bars[parent];
      const base = p.isTotal ? 0 : p.level - p.value;
      // Chain within the group: pick up where the previous detail left off.
      const prev = detailOf.get(c - 1) === parent ? bars[c - 1] : null;
      let level = prev ? prev.level : base;
      const segs: Seg[] = [];
      for (let si = 0; si < nSeries; si++) {
        const v = data.series[si]?.values[c];
        if (v == null || v === 0) continue;
        segs.push({ series: si, from: level, to: level + v, value: v });
        level += v;
      }
      bars.push({ segs, isTotal: false, value: level - (prev ? prev.level : base), level });
      continue;
    }
    if (totalSet.has(c)) {
      bars.push({
        segs: [{ series: 0, from: 0, to: running, value: running }],
        isTotal: true,
        value: running,
        level: running,
      });
      continue;
    }
    const segs: Seg[] = [];
    let level = running;
    for (let si = 0; si < nSeries; si++) {
      const v = data.series[si]?.values[c];
      if (v == null || v === 0) continue;
      segs.push({ series: si, from: level, to: level + v, value: v });
      level += v;
    }
    bars.push({ segs, isTotal: false, value: level - running, level });
    running = level;
  }
  return bars;
}

/** The value range a waterfall's bars actually occupy. */
export function waterfallExtent(cfg: ChartConfig): { min: number; max: number } {
  const vals = waterfallChain(cfg).flatMap((b) => [b.level, ...b.segs.flatMap((s) => [s.from, s.to])]);
  return { min: minOf(vals, 0), max: maxOf(vals, 0) };
}

/**
 * The font and box of a sideways value label that does not fit inside its bar and
 * is drawn to the RIGHT of it, given the room the row has from there to the
 * chart's edge.
 *
 * The box was a fixed `fs * 4` — 40pt at 10pt text, whatever the chart's width —
 * and the clip applied to it cut the number to that: "+125,000" drew as
 * "+125,0…" with 115pt of empty row beside it, and a wider chart did not help
 * because the box stayed 40pt. A truncated number reads as a different, plausible
 * one ("+1,250…" reads as 1,250). The upright branch shrinks the labels to the slot
 * and leaves out only the number no floor can fit — it cut them to a box narrower
 * than the slot until the box was sized to the label too (see the label's box in
 * `layoutWaterfall`).
 *
 * So the box is as wide as the label needs and no wider than the row has, and
 * where even that is short the label SHRINKS to the room before it is clipped —
 * down to the 5pt floor the upright fit uses. `fs * 4` stays the floor of the
 * box: a label that already fits it keeps its box byte for byte.
 *
 * A bar that ends at the frame's edge leaves the row LESS than that, and there is
 * no wider box to give it. It keeps the box and the font it had, and the chart's
 * own frame clip (`clipTextToFrame`) cuts the label to the edge as before —
 * shrinking toward a box the chart cannot hold would only make the same cut
 * smaller.
 */
function sideLabelFit(label: string, labelFs: number, fs: number, roomRight: number): { fs: number; w: number } {
  if (roomRight <= fs * 4) return { fs: labelFs, w: fs * 4 };
  // Whole points: the caller clips to `w - 2`, and `(width + 2) - 2` is not always
  // `width` in floating point, so a box sized to the label's exact width could
  // clip it by one ulp and cut a character.
  const need = (f: number) => Math.ceil(textWidth(label, f)) + 2;
  let f = labelFs;
  while (f > 5 && need(f) > roomRight) f -= 0.5;
  return { fs: f, w: Math.max(fs * 4, Math.min(need(f), roomRight)) };
}

/**
 * The width of the box an UPRIGHT value label is drawn in, given the width of the
 * label (`need`, at the size it is drawn at), of its bar (`barW`) and of the slot
 * the bar owns.
 *
 * The box is centred on the bar and capped at the slot, so adjacent boxes ABUT
 * instead of overlapping — and never narrower than the number it holds, up to that
 * cap. The font fit is against the SLOT (`labelRoom`), while the box was `barW + 12`
 * — two thirds of the slot and 12pt — and the label is clipped to the box. So a
 * label the fit had accepted as fitting was cut anyway, whenever it was wider than
 * the bar plus 10 and narrower than the slot minus 2: "+1,400,000" is 54pt at 10pt,
 * the slot at eight columns is 59, the box was 51.5, and the chart printed
 * "+1,400,0…" — a number that reads as another — while the same eight columns at
 * 960pt printed every label whole. A fit and a clamp that disagree are resolved by
 * the clamp, which is what the reader sees.
 *
 * The box is as wide as the label needs and no wider than the slot, and a label
 * that already fit the old box keeps it to the last bit. Whole points past the
 * label, as `sideLabelFit` does: the caller clips to `w - 2`, and `(width + 2) - 2`
 * is not always `width` in floating point.
 *
 * ZERO where the label does not fit the slot even then — the font fit floors at
 * 5pt, and a number wider than the slot at the floor was cut to a box that cannot
 * hold it. The caller leaves it out: the comment above `sideLabelFit` says the
 * upright branch clips only "the number no floor can fit", and a clip is a number
 * cut in half, which reads as a different one.
 */
function uprightLabelBox(need: number, barW: number, slotLen: number): number {
  const box = Math.min(need <= barW + 12 - 2 ? barW + 12 : Math.ceil(need) + 2, slotLen);
  return need <= Math.max(1, box - 2) ? box : 0;
}

export function layoutWaterfall(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const stacked = data.series.length > 1;
  const spacerSet = new Set(cfg.waterfall?.spacerIndices ?? []);
  const H = !!cfg.horizontal;
  const detailOf = detailParents(cfg);
  /** The next column the chain actually flows to, stepping over any details. */
  const nextChainColumn = (from: number) => {
    let j = from + 1;
    while (j < n && detailOf.has(j)) j++;
    return j;
  };

  const bars = waterfallChain(cfg);

  const allLevels = bars.flatMap((b) => b.segs.flatMap((s) => [s.from, s.to]));
  const hi = maxOf(allLevels, 0);
  const lo = minOf(allLevels, 0);
  const fmt = resolveFormat(
    bars.flatMap((b) => b.segs.map((s) => s.value)),
    cfg.numberFormat,
  );

  // A function of the frame, because the value-axis column is sized from the ticks
  // it prints (see `computeFrame`).
  const scaleOf = (f: Frame) => valueScale(f, lo, hi, cfg.scale, H ? undefined : cfg.axisBreak);
  const frame = H
    ? computeFrameHorizontal(cfg, style, { ...decor, seriesLabels: false })
    : computeFrame(cfg, style, { ...decor, seriesLabels: false }, [], scaleOf).frame;
  const scale = scaleOf(frame);
  const fs = style.fontSize;

  // Category slots along x (vertical) or y (horizontal); value axis the other way.
  const catStart = H ? frame.y : frame.x;
  const catLen = H ? frame.h : frame.w;
  const slotLen = catLen / Math.max(1, n);
  const colThick = slotLen * (2 / 3);
  const centers = Array.from({ length: n }, (_, i) => catStart + slotLen * (i + 0.5));
  const valLen = H ? frame.w : frame.h;
  // Clamped on both branches — see the note on the same map in column.ts. The
  // vertical one inherits `toY`'s clip; the horizontal one is its own linear
  // map and had none, so a manual scale narrower than the data extrapolated
  // straight off the slide.
  const qOf = H
    ? (v: number) => Math.max(0, Math.min(valLen, ((v - scale.min) / (scale.max - scale.min || 1)) * valLen))
    : (v: number) => frame.y + frame.h - scale.toY(v);
  const vRect = (catPos: number, v0: number, v1: number) => {
    const q0 = Math.min(qOf(v0), qOf(v1));
    const q1 = Math.max(qOf(v0), qOf(v1));
    return H
      ? { x: frame.x + q0, y: catPos - colThick / 2, w: Math.max(0.5, q1 - q0), h: colThick }
      : { x: catPos - colThick / 2, y: frame.y + frame.h - q1, w: colThick, h: Math.max(0.5, q1 - q0) };
  };

  const nodes: SceneNode[] = H
    ? horizontalChrome(cfg, style, { ...decor, seriesLabels: false }, frame, centers, scale, qOf)
    : chromeNodes(cfg, style, decor, frame, centers, scale);
  const zeroQ = qOf(0);
  const y0 = H ? frame.x + zeroQ : frame.y + frame.h - zeroQ;
  // The bridge's opening column rises from the zero baseline, so it reads as a
  // LEVEL and takes an unsigned label — but only it. Gating on `from === 0`
  // instead misread any mid-chain delta whose incoming running total happened
  // to be zero (…, -100, +50) as a base column and dropped its "+".
  const openingBar = bars.findIndex((b) => b.segs.length > 0);
  /** The text a segment is labelled with — shared so the fit below sees it too. */
  const labelFor = (isTotal: boolean, c: number, value: number) =>
    formatNumber(value, {
      ...fmt,
      // Base columns (the opening column and explicit totals) get unsigned
      // labels throughout; every other column is a delta and keeps its sign.
      forceSign: !isTotal && c !== openingBar && (cfg.numberFormat?.forceSign ?? true),
    });
  /**
   * One size for every value label, bounded by the slot its bar owns.
   *
   * Upright the labels sit side by side, so the bound is the slot's WIDTH — the
   * box was `colThick + 12` regardless, which bleeds 6pt into each neighbour
   * whenever the gap between bars is narrower than that. Horizontal they are
   * centred on their row in a box `fs * 1.5` tall, so the bound is the row
   * PITCH, the same one the totals beside them now take.
   *
   * Shrunk together so a row of labels reads at one size rather than as a
   * ransom note, then clipped for the number no floor can fit. Last resort as
   * everywhere else: a chart whose labels already fit their slot keeps `fs`.
   */
  const labelRoom = Math.max(1, slotLen - 2);
  const labelFs = (() => {
    if (H) return bandFontSize(fs, slotLen, 1.5);
    const labels = bars.flatMap((b, c) => b.segs.map((seg) => labelFor(b.isTotal, c, seg.value)));
    let f = fs;
    while (f > 5 && labels.some((l) => textWidth(l, f) > labelRoom)) f -= 0.5;
    return f;
  })();
  const columnTop: number[] = [];

  bars.forEach((b, c) => {
    let topQ = 0;

    b.segs.forEach((seg) => {
      const r = vRect(centers[c], seg.from, seg.to);
      topQ = Math.max(topQ, qOf(Math.max(seg.from, seg.to)));
      const fill = b.isTotal
        ? style.neutral
        : stacked
          ? seriesColor(style, seg.series, data.series[seg.series]?.color)
          : seg.value >= 0
            ? style.palette[0]
            : style.negative;
      nodes.push({
        kind: "rect",
        ...r,
        fill,
        stroke: stacked ? style.background : undefined,
        strokeWidth: stacked ? 0.75 : 0,
        name: `bar-${c}${stacked && !b.isTotal ? `-s${seg.series}` : ""}`,
      });

      const label = labelFor(b.isTotal, c, seg.value);
      const along = H ? r.w : r.h;
      const fits = H ? along >= textWidth(label, labelFs) + 2 : along >= labelFs * 1.25;
      // `labelFs` is 0 when the row is too thin for a legible label — the same
      // "no room" answer the other fits give, and what keeps a zero-size font
      // (which OOXML rejects) out of the scene.
      if ((fits || !stacked || b.isTotal) && labelFs > 0) {
        // Sideways, a label that cannot sit inside its bar sits to the right of it
        // in the room the row has left (see `sideLabelFit`).
        const side = H && !fits ? sideLabelFit(label, labelFs, fs, cfg.width - (r.x + r.w + 2)) : null;
        // A label that does not fit INSIDE its bar is drawn above it, and the
        // only thing above a bar that reaches the top of the plot is the edge of
        // the chart — 30pt over the top of a 300x60 chart at 32pt. Fitted to the
        // band the bar leaves, and dropped where that band cannot carry a
        // readable number: the same answer `labelFs` itself gives for a row too
        // thin to label.
        const sideFs = side?.fs ?? labelFs;
        const outsideFs = fits || H ? sideFs : aboveMarkFontSize(labelFs, r.y, titleInkBottom(cfg, style), 1.45);
        // Upright, the box is centred on the bar and capped at the slot (see
        // `uprightLabelBox`). `r.x - 6` is exactly `centers[c] - (r.w + 12) / 2`, so
        // an uncapped slot is byte-identical.
        const boxW = side ? side.w : H ? r.w + 12 : uprightLabelBox(textWidth(label, outsideFs), r.w, slotLen);
        // Left out, not cut, where even the smallest legible size does not fit the
        // slot (an upright box of 0) — the answer a total gives upright
        // (`column.ts`) for a number that cannot be read there, and the comment
        // above `sideLabelFit` promised it.
        if (outsideFs > 0 && boxW > 0)
          nodes.push({
            kind: "text",
            x: H ? (fits ? r.x - 6 : r.x + r.w + 2) : centers[c] - boxW / 2,
            y: fits || H ? r.y + r.h / 2 - outsideFs * 0.75 : r.y - outsideFs * 1.45,
            w: boxW,
            h: outsideFs * 1.5,
            text: clipToWidth(label, outsideFs, Math.max(1, boxW - 2)),
            fontSize: outsideFs,
            color: fits ? contrastInk(fill) : style.text,
            bold: b.isTotal,
            align: fits ? "center" : H ? "left" : "center",
            valign: "middle",
            name: `label-${c}${stacked && !b.isTotal ? `-s${seg.series}` : ""}`,
          });
      }
    });
    columnTop.push(b.segs.length ? (H ? frame.x + topQ : frame.y + frame.h - topQ) : y0);

    // Spacer slot: bridge the connector across the empty gap at the running
    // level so the dashed line reads as continuous through the section break.
    if (b.segs.length === 0 && spacerSet.has(c) && c > 0) {
      const levelQ = qOf(b.level);
      if (H) {
        const x = frame.x + levelQ;
        nodes.push({
          kind: "line",
          x1: x,
          y1: centers[c] - colThick / 2,
          x2: x,
          y2: centers[c] + colThick / 2,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `spacer-bridge-${c}`,
        });
      } else {
        const y = frame.y + frame.h - levelQ;
        nodes.push({
          kind: "line",
          x1: centers[c] - colThick / 2,
          y1: y,
          x2: centers[c] + colThick / 2,
          y2: y,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `spacer-bridge-${c}`,
        });
      }
    }

    // Dashed connector from this column's outgoing level to the next bar. A
    // detail column has no outgoing level to carry — the chain flows past the
    // whole group — so it draws none, and its parent's connector reaches over
    // the group to the next chain column instead. That step-over is what tells
    // the reader the group is a decomposition and not more of the walk.
    const isDetail = detailOf.has(c);
    const to = isDetail ? -1 : nextChainColumn(c);
    if (!isDetail && to > c && to < n) {
      const levelQ = qOf(b.level);
      if (H) {
        const x = frame.x + levelQ;
        nodes.push({
          kind: "line",
          x1: x,
          y1: centers[c] + colThick / 2,
          x2: x,
          y2: centers[to] - colThick / 2,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `connector-${c}`,
        });
      } else {
        const y = frame.y + frame.h - levelQ;
        nodes.push({
          kind: "line",
          x1: centers[c] + colThick / 2,
          y1: y,
          x2: centers[to] - colThick / 2,
          y2: y,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `connector-${c}`,
        });
      }
    }
  });

  if (!H) nodes.push(...breakMarkerNodes(frame, scale, style, cfg.width));
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

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => colThick),
      columnTop,
      columnValue: bars.map((b) => (b.isTotal ? b.value : b.level)),
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: H ? undefined : scale.toY,
      // The horizontal twin, published for the same reason `column.ts` publishes
      // it: a `Target` row is extracted from the data whatever the orientation,
      // and without a map to draw it back the values simply disappear.
      valueToX: H ? (v: number) => frame.x + qOf(v) : undefined,
    },
  };
}

/** Parse a datasheet column value for waterfalls: "e"/"=" marks a computed total. */
export function isTotalToken(raw: string): boolean {
  const t = raw.trim().toLowerCase();
  // "Σ".toLowerCase() is "σ" — compare against the lowercased form.
  return t === "e" || t === "=" || t === "σ";
}
