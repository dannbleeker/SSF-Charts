import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { lerpColor } from "../color";
import { fitPlot, footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";
import { PALETTE } from "../style";

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Worst aspect ratio in a candidate row (Bruls et al. squarified treemap).
 * Takes the row's max/min/sum directly so the caller can carry them incrementally
 * as the row grows — instead of re-slicing and re-scanning the row each step,
 * which made packing one strip O(L²) in allocations and comparisons.
 */
function worst(mx: number, mn: number, side: number, sum: number): number {
  const s2 = sum * sum;
  const side2 = side * side;
  return Math.max((side2 * mx) / s2, s2 / (side2 * mn));
}

/**
 * Squarified layout: pack `items` (each with an `area` already scaled to fill
 * `rect`) into rectangles whose aspect ratios stay close to 1. Returns a rect
 * per item, keyed by the item's `key`.
 */
function squarify<T extends { area: number; key: number }>(items: T[], rect: Rect): Map<number, Rect> {
  const out = new Map<number, Rect>();
  let { x, y, w: dx, h: dy } = rect;
  let i = 0;
  while (i < items.length) {
    const side = Math.min(dx, dy) || 1;
    let end = i + 1;
    let rowArea = items[i].area;
    // Carry the row's running max/min/sum; adding an item only folds in one value.
    let rowMax = items[i].area;
    let rowMin = items[i].area;
    while (end < items.length) {
      const a = items[end].area;
      const next = rowArea + a;
      const withNext = worst(Math.max(rowMax, a), Math.min(rowMin, a), side, next);
      const without = worst(rowMax, rowMin, side, rowArea);
      if (withNext <= without) {
        rowArea = next;
        rowMax = Math.max(rowMax, a);
        rowMin = Math.min(rowMin, a);
        end++;
      } else break;
    }
    const thickness = rowArea / side;
    let off = 0;
    for (let t = i; t < end; t++) {
      const len = items[t].area / (thickness || 1);
      out.set(
        items[t].key,
        dx >= dy ? { x, y: y + off, w: thickness, h: len } : { x: x + off, y, w: len, h: thickness },
      );
      off += len;
    }
    if (dx >= dy) {
      x += thickness;
      dx -= thickness;
    } else {
      y += thickness;
      dy -= thickness;
    }
    i = end;
  }
  return out;
}

/**
 * Treemap: one series, area ∝ value. Categories named "Group | Item" nest into
 * two levels — groups are squarified first, then each group's items within its
 * cell. Renders as native rects (no freeform paths needed).
 */
export function layoutTreemap(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const palette = cfg.style?.palette ?? PALETTE;
  const raw = data.categories.map((c, i) => ({ label: c, value: Math.max(0, data.series[0]?.values[i] ?? 0), i }));
  const items = raw.filter((r) => r.value > 0);
  const total = items.reduce((a, r) => a + r.value, 0) || 1;
  const fmt = resolveFormat(
    items.map((r) => r.value),
    cfg.numberFormat,
  );

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  const plot: Rect = fitPlot(cfg, { x: 2, y: titleH + 2, w: cfg.width - 4, h: cfg.height - titleH - footH - 4 });

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // "Group | Item" → two levels; otherwise a flat treemap.
  const grouped = items.some((r) => r.label.includes("|"));
  const groupOf = (label: string) => (label.includes("|") ? label.split("|")[0].trim() : "");
  const labelOf = (label: string) => (label.includes("|") ? label.split("|").slice(1).join("|").trim() : label);

  const drawTile = (r: Rect, fill: string, label: string, value: number, name: string, headerH = 0) => {
    nodes.push({
      kind: "rect",
      x: r.x,
      y: r.y,
      w: Math.max(0, r.w - 1),
      h: Math.max(0, r.h - 1),
      fill,
      stroke: style.background,
      strokeWidth: 1,
      name,
    });
    if (!decor.segmentLabels) return;
    const text = `${label}`;
    const valText = formatNumber(value, fmt);
    const ink = contrastInk(fill);
    if (r.w - 4 >= textWidth(text, fs * 0.85) && r.h - headerH >= fs * 2.2) {
      nodes.push({
        kind: "text",
        x: r.x + 3,
        y: r.y + headerH + 2,
        w: r.w - 6,
        h: fs * 1.3,
        text,
        fontSize: fs * 0.85,
        bold: true,
        color: ink,
        align: "left",
        valign: "top",
        name: `${name}-label`,
      });
      if (r.h - headerH >= fs * 3.4) {
        nodes.push({
          kind: "text",
          x: r.x + 3,
          y: r.y + headerH + fs * 1.4,
          w: r.w - 6,
          h: fs * 1.3,
          text: valText,
          fontSize: fs * 0.8,
          color: ink,
          align: "left",
          valign: "top",
          name: `${name}-value`,
        });
      }
    }
  };

  if (!grouped) {
    const sorted = items.map((r, k) => ({ ...r, key: k })).sort((a, b) => b.value - a.value);
    const scale = (plot.w * plot.h) / total;
    const rects = squarify(
      sorted.map((r) => ({ area: r.value * scale, key: r.key })),
      plot,
    );
    sorted.forEach((r) => {
      const rect = rects.get(r.key);
      if (rect) drawTile(rect, palette[r.i % palette.length], r.label, r.value, `tile-${r.i}`);
    });
  } else {
    // Level 1: squarify the groups (by total), preserving first-seen order.
    const groups: { name: string; total: number; members: typeof items; gi: number }[] = [];
    for (const r of items) {
      const g = groupOf(r.label);
      let entry = groups.find((e) => e.name === g);
      if (!entry) {
        entry = { name: g, total: 0, members: [], gi: groups.length };
        groups.push(entry);
      }
      entry.total += r.value;
      entry.members.push(r);
    }
    const gscale = (plot.w * plot.h) / total;
    const grects = squarify(
      groups.map((g) => ({ area: g.total * gscale, key: g.gi })),
      plot,
    );
    const headerH = fs * 1.5;
    groups.forEach((g) => {
      const gr = grects.get(g.gi);
      if (!gr) return;
      const gColor = palette[g.gi % palette.length];
      // Group header band + border.
      nodes.push({
        kind: "rect",
        x: gr.x,
        y: gr.y,
        w: Math.max(0, gr.w - 1),
        h: Math.max(0, gr.h - 1),
        fill: lerpColor(style.background, gColor, 0.14),
        stroke: gColor,
        strokeWidth: 1,
        name: `group-${g.gi}`,
      });
      // THE HEADER IS DRAWN ONLY IF THE CELL CAN HOLD IT. A group is squarified by
      // its share, so a small one is a hairline strip — 0.04pt at the foot of the
      // chart for a 0.01% group — and its header (a box `headerH` tall at the
      // cell's top) and its tiles (below that box) were placed regardless: the
      // label stood 14pt under the canvas and, between 5 and 14pt, ran down over
      // the next group's own header. Hairline cells draw no header and tile the
      // whole cell instead, so what a thin group has to say is its tiles, which the
      // flat path already draws that way.
      const headed = gr.h >= headerH + 2;
      if (headed)
        nodes.push({
          kind: "text",
          x: gr.x + 3,
          y: gr.y + 1,
          w: gr.w - 6,
          h: headerH,
          text: g.name,
          fontSize: fs * 0.9,
          bold: true,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `group-label-${g.gi}`,
        });
      // Level 2: squarify this group's members within the cell below the header.
      const inner: Rect = headed
        ? { x: gr.x + 2, y: gr.y + headerH, w: Math.max(1, gr.w - 4), h: Math.max(1, gr.h - headerH - 2) }
        : { x: gr.x + 1, y: gr.y + 1, w: Math.max(1, gr.w - 2), h: gr.h - 2 };
      // A cell thinner than its own gaps has no room for a tile: the group's outline
      // is all there is, and a zero-area tile is not an input `squarify` is for.
      if (inner.h < 1) return;
      const sorted = g.members.map((r, k) => ({ ...r, key: k })).sort((a, b) => b.value - a.value);
      const iscale = (inner.w * inner.h) / (g.total || 1);
      const rects = squarify(
        sorted.map((r) => ({ area: r.value * iscale, key: r.key })),
        inner,
      );
      sorted.forEach((r, k) => {
        // squarify keys its rects by the item's own key (the pre-sort index), so
        // look up by r.key — using the post-sort loop index handed each tile the
        // rectangle sized for a different member's value whenever the group's
        // members weren't already in descending order.
        const rect = rects.get(r.key);
        if (rect)
          drawTile(
            rect,
            lerpColor(gColor, style.background, 0.15 + 0.12 * (k % 4)),
            labelOf(r.label),
            r.value,
            `tile-${r.i}`,
          );
      });
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: raw.map(() => plot.x + plot.w / 2),
      categoryWidth: raw.map(() => plot.w),
      columnTop: raw.map(() => plot.y),
      columnValue: raw.map((r) => r.value),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
