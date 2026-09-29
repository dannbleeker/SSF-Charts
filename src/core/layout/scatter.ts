import type { ChartConfig, ChartStyle, Decorations, MarkerSymbol, NumberFormat } from "../types";
import { arrowheadFits, markerScale, markerSymbolOf } from "../geometry";
import { textWidth, type SceneNode, type TextNode } from "../scene";
import { tightBox } from "../collide";
import {
  formatNumber,
  formatP,
  histogramBins,
  niceTicks,
  polyTrend,
  resolveFormat,
  resolveAxisFormat,
  trendStats,
} from "../format";
import { placeLabels, type Box, type LabelRequest } from "../labels";
import { spreadAlongAxis } from "../spread";
import { PALETTE, paletteColor } from "../style";
import { lerpColor, sequentialScale, zoneFill } from "../color";
import {
  printsOnTitle,
  titleInkBottom,
  fitPlot,
  footnoteH,
  titleHeight,
  titleNode,
  legendRowCount,
  legendWrapWalk,
  MIN_PLOT_SIDE,
  MIN_LABEL_FS,
  tickGapScale,
  bandFontSize,
} from "./frame";
import type { LayoutResult } from "./column";

/**
 * Scatter / bubble chart, think-cell datasheet convention: categories are
 * point labels; rows named X, Y (and Size for bubbles) carry coordinates.
 * A row named Group (values 1..k) colors points by group.
 * Point labels are placed by the greedy collision-avoiding placer and hidden
 * when the chart gets too dense.
 */
/**
 * The overlap-relief cap, in DATA UNITS of the spread axis — or null when
 * spread is off. Exported so the footnote quotes exactly the number the layout
 * enforces: a disclosure that drifts from the cap is worse than no disclosure.
 */
export function spreadCap(cfg: ChartConfig): { axis: "x" | "y"; limit: number } | null {
  const axis = cfg.scatter?.spread;
  if (!axis || (cfg.kind !== "scatter" && cfg.kind !== "bubble")) return null;
  // Quadrants make a categorical claim — which box a point is in. A nudge, however
  // small and however well disclosed, could carry a marker across the crossing
  // line and change that claim. Overlap is the lesser problem, so spread yields.
  // Returning null here also suppresses the footnote, so the chart never
  // promises an approximation it did not make.
  if (cfg.decorations?.quadrants) return null;
  // Measure the cap against the SAME nice-ticked domain the plot maps, so the
  // disclosed number matches what the viewer sees. (This used to force a zero
  // baseline of its own, which — once the plot became data-driven — let the cap
  // exceed the whole visible range on a tight cluster.)
  const ticks = niceTicks(...scatterDomain(cfg, axis), 5);
  const range = ticks[ticks.length - 1] - ticks[0];
  if (!(range > 0)) return null;
  const limit = cfg.scatter?.spreadLimit ?? range * 0.02;
  return { axis, limit: Math.max(0, Math.min(limit, range * 0.1)) };
}

/**
 * The value domain the plot maps for one axis of a scatter/bubble: the plotted
 * points' coordinates plus the decorations that must stay on-plot (partition
 * lines, a quadrant crossing, axis bands). Shared by the layout and spreadCap.
 *
 * The domain follows the DATA and is NOT anchored to zero — a scatter is read by
 * point position, not by length-from-zero the way a bar chart is, so forcing a
 * zero baseline (as this once did) collapsed a tight cluster like x∈[1000,1050]
 * into a sliver.
 */
export function scatterDomain(cfg: ChartConfig, axis: "x" | "y"): [number, number] {
  const find = (re: RegExp) => cfg.data.series.find((s) => re.test(s.name.trim()));
  const xs = find(/^x$/i)?.values ?? [];
  const ys = find(/^y$/i)?.values ?? [];
  // Only points carrying BOTH coordinates are plotted; a half-specified point
  // contributes no extent.
  const vals: number[] = [];
  for (let i = 0; i < cfg.data.categories.length; i++) {
    const x = xs[i];
    const y = ys[i];
    if (x == null || y == null) continue;
    vals.push(axis === "x" ? x : y);
  }
  const extra = (find(axis === "x" ? /^x\s*line$/i : /^y\s*line$/i)?.values ?? []).filter(
    (v): v is number => v != null,
  );
  const q = cfg.decorations?.quadrants;
  if (q) extra.push(axis === "x" ? q.x : q.y);
  for (const b of cfg.decorations?.bands ?? []) if (b.axis === axis) extra.push(b.from, b.to);
  const all = [...vals, ...extra].filter((v) => Number.isFinite(v));
  if (!all.length) return [0, 1];
  let lo = Math.min(...all);
  let hi = Math.max(...all);
  if (lo === hi) {
    const pad = Math.abs(lo) > 1e-9 ? Math.abs(lo) * 0.1 : 1;
    lo -= pad;
    hi += pad;
  }
  return [lo, hi];
}

/** Width of the continuous-color gradient bar; the group legend clears it. */
const COLOR_BAR_W = 90;

/**
 * A point's group id, normalised to a whole number ≥ 1.
 *
 * Group ids come out of a datasheet row, so they are whatever a cell can hold —
 * blank, 0, negative, fractional, NaN. Every consumer indexes something with
 * them (the palette, the marker shape, the legend), and each used to re-derive
 * its own index from the raw value: `palette[NaN]` quietly yields `fill=
 * "undefined"`, and `markers[NaN]` yields no shape at all, which throws in the
 * renderer. Normalising once, here, makes all of them safe by construction.
 * Anything not a usable id reads as group 1 — the same bucket a point with no
 * Group row already falls into.
 */
function groupOf(v: number | null | undefined): number {
  const g = Math.round(Number(v ?? 1));
  return Number.isFinite(g) && g >= 1 ? g : 1;
}

/**
 * A point mark of the given shape, carrying the same ink as a circle of radius
 * `r`.
 *
 * `r` is the DATA radius — "area ∝ size" is the bubble's central claim — so
 * every shape is grown or shrunk by markerScale to match the circle's area.
 * Without that, a group drawn as stars would read as a quarter the magnitude of
 * an identical value drawn as squares.
 *
 * "circle" and "square" resolve to the ellipse and rect the scene already has,
 * so the default scatter emits exactly the node it always did (markerScale is
 * 1 for a circle) and its output cannot move; only the shapes that need preset
 * geometry become SymbolNodes.
 */
function markerNode(
  shape: MarkerSymbol,
  cx: number,
  cy: number,
  r: number,
  fill: string,
  stroke: string,
  strokeWidth: number,
  name: string,
): SceneNode {
  const m = markerExtent(shape, r);
  if (shape === "circle") return { kind: "ellipse", cx, cy, rx: m, ry: m, fill, stroke, strokeWidth, name };
  if (shape === "square")
    return { kind: "rect", x: cx - m, y: cy - m, w: m * 2, h: m * 2, fill, stroke, strokeWidth, name };
  return { kind: "symbol", shape, cx, cy, size: m, fill, stroke, strokeWidth, name };
}

/** Drawn half-extent of a marker whose data radius is `r`. */
const markerExtent = (shape: MarkerSymbol, r: number) => r * markerScale(shape);

/**
 * The part of a segment that lies between `top` and `bot`, or null when none of
 * it does. Clipping, not clamping: moving an endpoint onto the boundary bends
 * the line, and the segments this serves are straight fits whose slope is the
 * claim being made.
 */
function clipToPlotY(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  top: number,
  bot: number,
): { x1: number; y1: number; x2: number; y2: number } | null {
  const dy = y2 - y1;
  let t0 = 0;
  let t1 = 1;
  if (dy === 0) {
    if (y1 < top || y1 > bot) return null;
  } else {
    const ta = (top - y1) / dy;
    const tb = (bot - y1) / dy;
    t0 = Math.max(0, Math.min(ta, tb));
    t1 = Math.min(1, Math.max(ta, tb));
    if (!(t1 > t0)) return null;
  }
  const lerp = (a: number, b: number, t: number) => a + (b - a) * t;
  return {
    x1: lerp(x1, x2, t0),
    y1: lerp(y1, y2, t0),
    x2: lerp(x1, x2, t1),
    y2: lerp(y1, y2, t1),
  };
}

/**
 * One axis of the plot: its ticks, the scale that maps a value to a coordinate,
 * the domain those ticks span, and the format its labels are drawn in.
 *
 * ── WHY A BUNDLE, AND WHY THIS ONE ──────────────────────────────────────────
 * `layoutScatter` derives ten loose locals up front — `xTicks`, `yTicks`, `x0`,
 * `x1`, `y0`, `y1`, `toX`, `toY`, `xFmt`, `yFmt` — and every block that draws
 * against an axis needs three to five of them. Passing them individually is what
 * makes an extracted helper take nine arguments and read worse than the code it
 * replaced.
 *
 * A SINGLE CONTEXT OBJECT WOULD BE WORSE, and that is measured rather than
 * asserted. `{cfg, style, decor, plot, toX, toY, xTicks, yTicks, fs}` serves
 * exactly one caller fully and eight others at two to five of its nine fields,
 * while each of them still needs three to seven arguments besides. It renames
 * the problem.
 *
 * Two per-axis bundles instead: six of the ten candidate extractions take
 * `(x, y)` and nothing else about the scale. The precedent is already here —
 * `frame.ts:192`'s `ValueScale` is passed as one parameter to `chromeNodes` and
 * `horizontalChrome` for the same reason.
 */
interface ScatterAxis {
  /** The nice ticks, ascending. `ticks.length - 1` is the interval count. */
  ticks: number[];
  /** Value to coordinate. */
  to: (v: number) => number;
  /** First and last tick — the domain actually plotted, not the data's extent. */
  lo: number;
  hi: number;
  fmt: NumberFormat;
}

/**
 * The quadrant preset: one X/Y crossing, four tinted zones with corner labels,
 * and the crossing lines — BCG-matrix framing in one step.
 */
function quadrantNodes(
  q: NonNullable<Decorations["quadrants"]>,
  x: ScatterAxis,
  y: ScatterAxis,
  plot: { x: number; y: number; w: number; h: number },
  style: ChartStyle,
  fs: number,
): SceneNode[] {
  const out: SceneNode[] = [];
  const { x: qx, y: qy, labels } = q;
  const cx = Math.max(plot.x, Math.min(plot.x + plot.w, x.to(qx)));
  const cy = Math.max(plot.y, Math.min(plot.y + plot.h, y.to(qy)));
  const zones: { x: number; y: number; w: number; h: number }[] = [
    { x: plot.x, y: plot.y, w: cx - plot.x, h: cy - plot.y }, // TL
    { x: cx, y: plot.y, w: plot.x + plot.w - cx, h: cy - plot.y }, // TR
    { x: plot.x, y: cy, w: cx - plot.x, h: plot.y + plot.h - cy }, // BL
    { x: cx, y: cy, w: plot.x + plot.w - cx, h: plot.y + plot.h - cy }, // BR
  ];
  zones.forEach((z, i) => {
    if (z.w <= 0 || z.h <= 0) return;
    // Checkerboard tint so adjacent zones read as distinct regions.
    out.push({
      kind: "rect",
      ...z,
      fill: zoneFill(style.background, i === 0 || i === 3 ? "#f2f1ec" : "#faf9f6"),
      name: `quadrant-${i}`,
    });
    const label = labels?.[i];
    /**
     * FITTED TO ITS OWN ZONE, and dropped when the zone cannot carry it.
     *
     * The box was `Math.max(20, z.w - 8)` — a floor that RAISES a width, which
     * this file has already been caught by once: `a conservative bound must
     * stay conservative at the small end`, the scatter legend's own bug. A
     * zone narrower than 28 points got a 20-point box starting 4 points inside
     * it, so a right-aligned label's ink was placed past the zone's right edge
     * and, on a narrow chart, past the chart's: 9.3 points off a 60x300 frame
     * at every font size, which is the tell that no font-dependent bound was
     * involved at all.
     *
     * A quadrant label names its quadrant, so the room it has is the zone —
     * the same rule the pie's inside labels and the mekko's column labels
     * follow. Shrink to it, and below `MIN_LABEL_FS` draw nothing: the tint
     * still shows the four zones, and a two-point word in the corner of one
     * names nothing.
     */
    const room = z.w - 8;
    let qf = fs * 0.9;
    if (label) while (qf > MIN_LABEL_FS && textWidth(label, qf, true) > room) qf -= 0.5;
    if (label && room > 0 && textWidth(label, qf, true) <= room) {
      out.push({
        kind: "text",
        x: z.x + 4,
        y: z.y + 2,
        w: room,
        h: (fs * 1.3 * qf) / (fs * 0.9),
        text: label,
        fontSize: qf,
        bold: true,
        color: style.mutedText,
        align: i === 1 || i === 3 ? "right" : "left",
        valign: "top",
        name: `quadrant-label-${i}`,
      });
    }
  });
  out.push(
    {
      kind: "line",
      x1: cx,
      y1: plot.y,
      x2: cx,
      y2: plot.y + plot.h,
      stroke: style.mutedText,
      strokeWidth: 1,
      dash: [3, 2],
      name: "quadrant-x",
    },
    {
      kind: "line",
      x1: plot.x,
      y1: cy,
      x2: plot.x + plot.w,
      y2: cy,
      stroke: style.mutedText,
      strokeWidth: 1,
      dash: [3, 2],
      name: "quadrant-y",
    },
  );
  return out;
}

/** Background bands (both axes in value units), behind gridlines and points. */
function bandNodes(
  bands: NonNullable<Decorations["bands"]>,
  x: ScatterAxis,
  y: ScatterAxis,
  plot: { x: number; y: number; w: number; h: number },
  style: ChartStyle,
  fs: number,
): SceneNode[] {
  const out: SceneNode[] = [];
  bands.forEach((band, i) => {
    const clampX = (v: number) => Math.max(plot.x, Math.min(plot.x + plot.w, x.to(v)));
    const clampY = (v: number) => Math.max(plot.y, Math.min(plot.y + plot.h, y.to(v)));
    const r =
      band.axis === "x"
        ? {
            x: Math.min(clampX(band.from), clampX(band.to)),
            y: plot.y,
            w: Math.abs(clampX(band.to) - clampX(band.from)),
            h: plot.h,
          }
        : {
            x: plot.x,
            y: Math.min(clampY(band.from), clampY(band.to)),
            w: plot.w,
            h: Math.abs(clampY(band.to) - clampY(band.from)),
          };
    if (r.w <= 0 || r.h <= 0) return;
    out.push({ kind: "rect", ...r, fill: band.color ?? zoneFill(style.background, "#f2f1ec"), name: `band-${i}` });
    if (band.label) {
      out.push({
        kind: "text",
        x: r.x + 3,
        y: r.y + 1,
        w: Math.max(20, r.w - 6),
        h: fs * 1.3,
        text: band.label,
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "left",
        valign: "top",
        name: `band-label-${i}`,
      });
    }
  });
  return out;
}

/**
 * The trail through the points in datasheet order, with a direction arrowhead at
 * each segment midpoint — a Gapminder-style path of one entity over time.
 *
 * Returned rather than pushed. `nodes` is one z-ordered array and the SVG bytes
 * depend on its order, so a helper that appends to it can only be called in
 * exactly one place; one that returns is a function.
 */
function trajectoryNodes(
  pts: readonly { x: number; y: number }[],
  x: ScatterAxis,
  y: ScatterAxis,
  style: ChartStyle,
  cfg: ChartConfig,
): SceneNode[] {
  const out: SceneNode[] = [];
  for (let i = 0; i < pts.length - 1; i++) {
    const ax = x.to(pts[i].x);
    const ay = y.to(pts[i].y);
    const bx = x.to(pts[i + 1].x);
    const by = y.to(pts[i + 1].y);
    out.push({
      kind: "line",
      x1: ax,
      y1: ay,
      x2: bx,
      y2: by,
      stroke: style.mutedText,
      strokeWidth: 1.5,
      name: `trajectory-${i}`,
    });
    const angle = (Math.atan2(by - ay, bx - ax) * 180) / Math.PI;
    // The direction glyph, where it fits. Its tip is the segment's MIDPOINT
    // and its body runs back along the segment, so a path along the top of the
    // plot — markers already overhang that edge by design — pushed the
    // triangle off the canvas, 4pt above an 80x60 chart at a 32pt font, onto
    // whatever sits over it on the slide.
    //
    // Dropped rather than moved: the arrowhead's whole job is to say which way
    // the path runs, so an arrowhead somewhere other than on its own segment
    // says something false. The segment's LINE is drawn either way, so the
    // path is still there — only the direction glyph on that one hop is not.
    if (arrowheadFits((ax + bx) / 2, (ay + by) / 2, 4, angle, cfg.width, cfg.height))
      out.push({
        kind: "arrowhead",
        x: (ax + bx) / 2,
        y: (ay + by) / 2,
        angle,
        size: 4,
        fill: style.mutedText,
        name: `trajectory-head-${i}`,
      });
  }
  return out;
}

/**
 * Marginal distribution histograms in the reserved gutters.
 *
 * The bin count is a multiple of the axis's own tick intervals, so every tick is
 * a bin edge and a bar reads straight against the scale beside it. A rule keyed
 * off the sample size alone (Sturges, Freedman-Diaconis) puts edges BETWEEN the
 * ticks, which is exactly what a chart-adjacent histogram must not do. The
 * multiplier is the only freedom and it stays bounded: past ~2 sub-bins per
 * interval the bars are a few points wide and read as noise, not shape.
 *
 * `fill` is passed already mixed rather than derived here, which is what keeps
 * this to six parameters instead of eight — it is the only thing the block
 * wanted from `style` and `cfg` together.
 */
function marginalNodes(
  pts: readonly { x: number; y: number }[],
  x: ScatterAxis,
  y: ScatterAxis,
  plot: { x: number; y: number; w: number; h: number },
  gutter: { top: number; right: number },
  fill: string,
): SceneNode[] {
  const out: SceneNode[] = [];
  const binMult = pts.length >= 15 ? 2 : 1;
  if (gutter.top > 0) {
    const counts = histogramBins(
      pts.map((p) => p.x),
      x.lo,
      x.hi,
      (x.ticks.length - 1) * binMult,
    );
    const peak = Math.max(1, ...counts);
    const bw = plot.w / counts.length;
    counts.forEach((n, i) => {
      if (!n) return;
      const h = (n / peak) * (gutter.top - 5);
      out.push({
        kind: "rect",
        x: plot.x + i * bw,
        y: plot.y - 3 - h,
        w: Math.max(0.5, bw - 1),
        h,
        fill,
        name: `marginal-x-${i}`,
      });
    });
  }
  if (gutter.right > 0) {
    const counts = histogramBins(
      pts.map((p) => p.y),
      y.lo,
      y.hi,
      (y.ticks.length - 1) * binMult,
    );
    const peak = Math.max(1, ...counts);
    const bh = plot.h / counts.length;
    counts.forEach((n, i) => {
      if (!n) return;
      const w = (n / peak) * (gutter.right - 5);
      // Bin 0 is the bottom of the y axis, so it is the LAST band down the plot.
      out.push({
        kind: "rect",
        x: plot.x + plot.w + 3,
        y: plot.y + plot.h - (i + 1) * bh,
        w,
        h: Math.max(0.5, bh - 1),
        fill,
        name: `marginal-y-${i}`,
      });
    });
  }
  return out;
}

/**
 * Gridlines, tick labels and the two spines, for both axes.
 *
 * The largest of the five blocks lifted out of `layoutScatter`, and the last.
 * It is self-contained in a way none of the others had to be checked for: it
 * derives `gapScale`, `yTickScale`, `xTickScale` and `zeroSpineX`, and NOTHING
 * after it reads any of them — verified by grep across the whole file, because
 * "it looks local" is how a forward dependency survives a move.
 *
 * ── THE PARAMETERS ARE `xAxis` / `yAxis`, NOT `x` / `y` ─────────────────────
 * The y loop declares `const y = yAxis.to(t)`. A parameter called `y` would be
 * shadowed by it — `const y = y.to(t)` is a TDZ error at runtime and reads as a
 * typo. The names match the call site, which is also where a reader looks next.
 *
 * ── THE PROSE IS UNTOUCHED, DELIBERATELY ────────────────────────────────────
 * Only code lines were rewritten onto the bundles. The comment at the spine
 * QUOTES CODE THAT NO LONGER EXISTS — "`toX(0) >= plot.x ? … : plot.x` pinned the
 * spine to the left edge" — and rewriting a historical quote to name today's API
 * would destroy the record of what the bug was. Every comment line in this
 * function is byte-identical to the block it came from, checked rather than
 * intended.
 */
function axisNodes(
  xAxis: ScatterAxis,
  yAxis: ScatterAxis,
  plot: { x: number; y: number; w: number; h: number },
  style: ChartStyle,
  fs: number,
  cfg: ChartConfig,
): SceneNode[] {
  const out: SceneNode[] = [];
  /**
   * The size each axis's tick labels are drawn at.
   *
   * One label per tick, each centred on its tick, so the room each has is the
   * SPACING between adjacent ticks — and neither axis was fitted to it. On a
   * plot small relative to the font the labels were drawn over each other:
   * 60 of the 237 overlapping text pairs a sweep found were this axis alone,
   * which made it the single worst offender in the engine.
   *
   * Bound by that spacing, the same rule the shared value axis and the radar's
   * ring ticks now use. Last resort: where the ticks already clear each other
   * this is 1 and nothing moves.
   *
   * Zero when the spacing cannot pay for a LEGIBLE label, and the labels are
   * then dropped rather than drawn. A fit with no floor answers whatever the
   * arithmetic says: six y ticks 1.6pt apart on a 200x150 chart at a 26pt font
   * produced six ONE-POINT labels — ink no reader can resolve, stacked in the
   * axis gutter, from a fit that reported success. Same answer the radar,
   * sunburst, tilemap and pie reservations give when their band cannot be met:
   * a label that cannot be read is not there. The gridlines stay, since those
   * still carry the scale.
   */
  /**
   * How much of the chart font a tick strip may take, from the gap between its
   * own ticks. Lives in `frame.ts` now — the secondary value axis needs exactly
   * this and had no fit at all, which was the largest remaining shape in the
   * overlap sweep. See `tickGapScale`.
   */
  const gapScale = (vals: number[], to: (v: number) => number, span: number, want: (t: number) => number) =>
    tickGapScale(fs * 0.9, vals, to, span, want);
  const yTickScale = gapScale(yAxis.ticks, yAxis.to, plot.h, () => fs * 1.4);
  /**
   * The x strip's size, AFTER the origin nudge has been paid for.
   *
   * `gapScale` guarantees the tick spacing can carry the widest label. The nudge
   * below then spends part of that guarantee: the label on the axis's origin is
   * centred on its tick, so half of it hangs into the y axis's gutter, and it is
   * moved right by exactly the overhang. Moving it right moves it TOWARD its
   * neighbour, and the gap the fit had just opened closes by the same amount —
   * 84 pairs of `x-axis` on `x-axis` in the variant sweep, on a strip that had a
   * fit and looked correct.
   *
   * The secondary value axis had the same defect and took the other remedy: one
   * shift for the whole strip, which keeps every gap. That is not available here
   * because only the FIRST label is nudged, and shifting all of them right by
   * its overhang would push the last one off the canvas. So this strip pays in
   * SIZE instead: shrink until the nudged layout clears, and drop the numbers if
   * it cannot — the answer this file already gives when the ticks will not fit.
   *
   * Half a point at a time rather than solving it: the nudge depends on the
   * label width, which depends on the size, which is what is being solved for.
   * A dozen iterations at most, on a strip of five or six numbers.
   */
  const xTickScale = (() => {
    const want = (t: number) => textWidth(formatNumber(t, xAxis.fmt), fs * 0.9) + 2;
    let scale = gapScale(xAxis.ticks, xAxis.to, plot.w, want);
    if (scale <= 0 || xAxis.ticks.length < 2) return scale;
    // Where each label's ink actually lands at this size, nudge included.
    const spans = (s: number) =>
      xAxis.ticks.map((t) => {
        const half = textWidth(formatNumber(t, xAxis.fmt), fs * 0.9 * s) / 2;
        const centre = Math.max(xAxis.to(t), plot.x + half);
        return { lo: centre - half, hi: centre + half };
      });
    const clears = (s: number) => {
      const b = spans(s);
      for (let i = 1; i < b.length; i++) if (b[i].lo - b[i - 1].hi < 1) return false;
      return true;
    };
    while (scale > 0 && !clears(scale)) {
      scale -= 0.05;
      if (fs * 0.9 * scale < MIN_LABEL_FS) return 0;
    }
    return scale;
  })();
  for (const t of yAxis.ticks) {
    const y = yAxis.to(t);
    out.push({
      kind: "line",
      x1: plot.x,
      y1: y,
      x2: plot.x + plot.w,
      y2: y,
      stroke: style.gridline,
      strokeWidth: 0.75,
      name: "gridline-y",
    });
    // Not into the TITLE's band: `fitPlot` grows a squeezed plot up from its
    // bottom edge, so on a 300x60 chart at 24pt the axis and its ticks are
    // inside the title. Per tick, like the value axis in `frame.ts` — the ones
    // lower down are still where they belong.
    if (yTickScale > 0 && !printsOnTitle(cfg, style, y - fs * 0.7 * yTickScale)) {
      out.push({
        kind: "text",
        x: 0,
        y: y - fs * 0.7 * yTickScale,
        w: plot.x - 4,
        h: fs * 1.4 * yTickScale,
        text: formatNumber(t, yAxis.fmt),
        fontSize: fs * 0.9 * yTickScale,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "y-axis",
      });
    }
  }
  if (xTickScale > 0) {
    for (const t of xAxis.ticks) {
      const x = xAxis.to(t);
      // A tick label is CENTRED on its tick, so the one at the axis's origin
      // puts half its width to the LEFT of the plot — which is the strip the y
      // axis writes its own numbers in, and at 18pt on a 480x300 chart the two
      // corner labels met. Nudged right by exactly the overlap, the same move
      // the gantt's last tick label makes at the other end of its axis: a label
      // that already clears the gutter does not move at all.
      const half = textWidth(formatNumber(t, xAxis.fmt), fs * 0.9 * xTickScale) / 2;
      const at = Math.max(x, plot.x + half);
      // The x strip sits under the plot, and a plot squeezed into the title's
      // band takes the strip with it — the whole row, since these share one y.
      if (printsOnTitle(cfg, style, plot.y + plot.h + 2)) break;
      out.push({
        kind: "text",
        x: at - 24,
        y: plot.y + plot.h + 2,
        w: 48,
        h: fs * 1.4 * xTickScale,
        text: formatNumber(t, xAxis.fmt),
        fontSize: fs * 0.9 * xTickScale,
        color: style.mutedText,
        align: "center",
        valign: "top",
        name: "x-axis",
      });
    }
  }
  const zeroSpineX = Math.max(plot.x, Math.min(plot.x + plot.w, xAxis.to(0)));
  out.push(
    {
      kind: "line",
      x1: plot.x,
      y1: plot.y + plot.h,
      x2: plot.x + plot.w,
      y2: plot.y + plot.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    },
    {
      kind: "line",
      // Clamped on BOTH sides. `toX(0) >= plot.x ? … : plot.x` pinned the spine
      // to the left edge when zero fell left of the domain and left it free
      // when zero fell right of it — so an all-negative x axis (a variance or
      // drawdown scatter) put a full-height spine 109pt past the right edge of
      // the canvas. A guard that guards one direction is not a guard.
      x1: zeroSpineX,
      y1: plot.y,
      x2: zeroSpineX,
      y2: plot.y + plot.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "y-axis-line",
    },
  );
  return out;
}

export function layoutScatter(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()));
  const xs = find(/^x$/i)?.values ?? [];
  const ys = find(/^y$/i)?.values ?? [];
  const sizes = cfg.kind === "bubble" ? (find(/^size$/i)?.values ?? []) : [];
  const groups = find(/^group$/i)?.values ?? [];
  // A numeric "Color" row encodes a third/fourth variable on a sequential ramp.
  const colorVals = find(/^colou?r$/i)?.values ?? [];
  // Partition lines and trend line, think-cell's scatter decorations.
  const xLines = (find(/^x\s*line$/i)?.values ?? []).filter((v): v is number => v != null);
  const yLines = (find(/^y\s*line$/i)?.values ?? []).filter((v): v is number => v != null);
  const wantTrend = (find(/^trend$/i)?.values ?? []).some((v) => v != null);

  const pts = data.categories
    .map((label, i) => ({
      label,
      x: xs[i],
      y: ys[i],
      size: sizes[i] ?? null,
      group: groupOf(groups[i]),
      color: colorVals[i] ?? null,
    }))
    .filter((p): p is typeof p & { x: number; y: number } => p.x != null && p.y != null);

  // Shape per group, cycled like the palette. Off => every point a circle,
  // which is the ellipse the layout has always emitted.
  const markers = cfg.scatter?.markers?.length ? cfg.scatter.markers : null;
  // Narrowed, not trusted: `markers` is config, and MarkerSymbol is erased at
  // runtime — an unknown name reached the renderers as a SymbolNode shape they
  // have no geometry for and threw mid-render.
  const markerFor = (group: number): MarkerSymbol =>
    markers ? markerSymbolOf(markers[(group - 1) % markers.length]) : "circle";

  // Continuous color scale (a "Color" row): maps each point onto a sequential
  // ramp; supersedes group coloring and swaps the chip legend for a gradient.
  const colorNums = pts.map((p) => p.color).filter((v): v is number => v != null);
  const colorScale =
    colorNums.length > 0
      ? {
          min: Math.min(...colorNums),
          max: Math.max(...colorNums),
          of: sequentialScale(Math.min(...colorNums), Math.max(...colorNums), (cfg.style?.palette ?? PALETTE)[0]),
        }
      : null;

  const titleH = titleHeight(cfg, style);
  const axisW = 34;
  const multiGroup = !colorScale && new Set(pts.map((p) => p.group)).size > 1;
  // ROWS, not one row. This legend is hand-rolled — it does not go through
  // `legendRow`/`legendRowCount` the way mekko, boxplot, radar, butterfly and
  // column all do — and it never wrapped: entries marched right until they left
  // the frame. Eight groups on the DEFAULT 480pt frame ran to x=520. In SVG the
  // viewBox clips them (a group silently unexplained); in PowerPoint it is
  // worse, because the Office renderer applies no clamp and they become real
  // shapes sitting off the chart on the slide.
  //
  // Reserved from a deliberately CONSERVATIVE width — the frame is not computed
  // yet here, and the drawing walk below uses the real plot, which is never
  // narrower than this. Reserving with the narrower bound can only over-reserve,
  // and an extra empty row costs a few points where an unreserved one overlaps
  // the plot.
  //
  // The `max(40, …)` floor this line used to carry INVERTED that on a narrow
  // frame, which is the one case it mattered. At 80pt wide the estimate is 38
  // and the floor raised it to 40 — a WIDER bound than the walk gets, so the
  // count came out lower than the rows drawn, the legend wrapped past its band,
  // and its third row landed in the x-axis strip. A conservative bound has to
  // stay conservative at the small end; flooring a WIDTH up reserves less, not
  // more.
  const groupLegendRows =
    multiGroup && !colorScale
      ? legendRowCount(
          [...new Set(pts.map((p) => p.group))].filter((g) => g !== undefined).map((g) => `Group ${g}`),
          fs,
          0,
          Math.max(1, cfg.width - axisW - 8),
        )
      : 0;
  // A gutter costs real space. If what's left would stop being a chart, drop
  // the marginals rather than the plot. Declared here rather than beside its
  // first use below, because the legend decision needs the same floor: both are
  // asking "would what remains still be a chart".
  const MIN_PLOT = 60;
  // A wrapped group legend can eat the whole chart, so it is DROPPED rather
  // than drawn over what is left.
  //
  // `bodyH` below is what remains after this band, and on a small frame a
  // three-row legend takes more than the frame has: at 80x60 the rows ran to
  // y=44 on a 60pt chart and the third landed in the x-axis strip. `fitPlot`
  // then floors the plot and the axis comes up to meet the legend, so following
  // the fitted plot — which `legendTop` already does — cannot save it.
  //
  // The same answer the radar, sunburst, tilemap and pie reservations give, and
  // the cascade's group headers: chrome that cannot be paid for is not drawn.
  // The plot is the chart; a legend naming groups nobody can see is not.
  const legendWant = colorScale ? fs * 1.8 : multiGroup ? fs * 1.8 * Math.max(1, groupLegendRows) : 0;
  const roomForLegend = cfg.height - titleH - 6 - fs * 1.6 - footnoteH(cfg, style, decor) - MIN_PLOT;
  const showGroupLegend = !colorScale && legendWant > 0 ? legendWant <= roomForLegend : true;
  const legendH = showGroupLegend ? legendWant : 0;
  /** Where the plot starts before any marginal gutter — the legends' anchor. */
  const chromeTop = titleH + 6 + legendH;
  // Marginal gutters, in font-size units like the heatmap's already-shipped
  // marginal totals, so they scale with the style instead of being a magic 34.
  const GUT = fs * 3.4;
  const wantMx = decor.marginals === "x" || decor.marginals === "both";
  const wantMy = decor.marginals === "y" || decor.marginals === "both";
  // NB: bodyW/bodyH are today's w/h expressions character-for-character. Do not
  // "simplify" bodyH in terms of chromeTop — float subtraction is not
  // associative, and (H-t)-6-l differs from H-(t+6+l) for many font sizes. It
  // happens to agree at the default fs=10, which every showcase config uses, so
  // the deck's byte-identity gate would NOT catch the difference.
  const bodyW = cfg.width - axisW - 8;
  const bodyH = cfg.height - titleH - 6 - legendH - fs * 1.6 - footnoteH(cfg, style, decor);
  const mTop = wantMx && bodyH - GUT >= MIN_PLOT ? GUT : 0;
  const mRight = wantMy && bodyW - GUT >= MIN_PLOT ? GUT : 0;
  const plot = fitPlot(cfg, {
    x: axisW,
    y: chromeTop + mTop,
    w: bodyW - mRight,
    h: bodyH - mTop,
  });
  /**
   * Where the legends hang from. They sit ABOVE the plot's top edge, so they
   * have to follow the FITTED plot: `chromeTop` is the position the layout asked
   * for, and on a frame too short to pay for its own chrome `fitPlot` moves the
   * plot up while the legend, anchored to the request, stays behind — 9pt past
   * the bottom of a 70pt chart. Identical to `chromeTop` whenever the plot fits.
   */
  const legendTop = plot.y - mTop;

  // Data-driven axis domain (no forced zero baseline) shared with spreadCap;
  // folds in partition lines, a quadrant crossing and x/y bands so those
  // decorations never fall outside the plot. See scatterDomain.
  const xTicks = niceTicks(...scatterDomain(cfg, "x"), 5);
  const yTicks = niceTicks(...scatterDomain(cfg, "y"), 5);
  const x0 = xTicks[0];
  const x1 = xTicks[xTicks.length - 1];
  const y0 = yTicks[0];
  const y1 = yTicks[yTicks.length - 1];
  const toX = (v: number) => plot.x + ((v - x0) / (x1 - x0 || 1)) * plot.w;
  const toY = (v: number) => plot.y + plot.h - ((v - y0) / (y1 - y0 || 1)) * plot.h;

  const xFmt = resolveAxisFormat(xTicks, cfg.numberFormat);
  const yFmt = resolveAxisFormat(yTicks, cfg.numberFormat);
  // The ten loose locals above, as the two bundles every extracted block takes.
  // See `ScatterAxis`. The originals stay: the blocks still inline here use them
  // by name, and rewriting those call sites would make this a rewrite rather
  // than a relocation.
  const xAxis: ScatterAxis = { ticks: xTicks, to: toX, lo: x0, hi: x1, fmt: xFmt };
  const yAxis: ScatterAxis = { ticks: yTicks, to: toY, lo: y0, hi: y1, fmt: yFmt };

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Behind the gridlines and the points — push order is paint order.
  if (decor.quadrants) nodes.push(...quadrantNodes(decor.quadrants, xAxis, yAxis, plot, style, fs));
  if (decor.bands) nodes.push(...bandNodes(decor.bands, xAxis, yAxis, plot, style, fs));
  nodes.push(...axisNodes(xAxis, yAxis, plot, style, fs, cfg));

  // Partition lines (dashed) at fixed x / y values.
  for (const v of xLines) {
    const x = toX(v);
    nodes.push({
      kind: "line",
      x1: x,
      y1: plot.y,
      x2: x,
      y2: plot.y + plot.h,
      stroke: style.mutedText,
      strokeWidth: 1,
      dash: [3, 2],
      name: "x-line",
    });
  }
  for (const v of yLines) {
    const y = toY(v);
    nodes.push({
      kind: "line",
      x1: plot.x,
      y1: y,
      x2: plot.x + plot.w,
      y2: y,
      stroke: style.mutedText,
      strokeWidth: 1,
      dash: [3, 2],
      name: "y-line",
    });
  }

  // Trend line across all points, always stating the fit. Linear (default) is a
  // straight OLS line; a higher `scatter.trendDegree` fits a polynomial drawn as
  // a sampled curve.
  const trendDeg = Math.max(1, Math.min(4, Math.floor(cfg.scatter?.trendDegree ?? 1)));
  // The fallback has to cover THREE points as well as two. `polyTrend` clamps
  // its degree to `n - 2` to keep a residual degree of freedom, so three points
  // come back as a degree-1 fit, and the caller below discards anything under
  // degree 2 — so a `Trend` row with `trendDegree: 2` and exactly three points
  // drew no line, no R² and no diagnostic, while the same config with two
  // points drew a straight fit. Three points is not a degenerate input.
  if (wantTrend && pts.length >= 2 && (trendDeg <= 1 || pts.length < 4)) {
    const mx = pts.reduce((s, p) => s + p.x, 0) / pts.length;
    const my = pts.reduce((s, p) => s + p.y, 0) / pts.length;
    const sxx = pts.reduce((s, p) => s + (p.x - mx) ** 2, 0);
    if (sxx > 0) {
      const slope = pts.reduce((s, p) => s + (p.x - mx) * (p.y - my), 0) / sxx;
      const at = (x: number) => my + slope * (x - mx);
      // CLIPPED to the plot, which the polynomial branch below already does and
      // this one — the default — did not.
      //
      // The line is drawn across the padded NICE-TICK ends, not the data's own
      // x range, and `scatterDomain` deliberately folds in `X line` rows, x-axis
      // `bands` and a `quadrants` crossing. Any of those three documented
      // decorations therefore stretches x0 far past the data and the
      // extrapolated y explodes with it: an `X line` at −1000 put the trend's
      // endpoints at y = 93043 and −45957 on a 300pt canvas.
      //
      // Clipped rather than clamped: clamping the endpoints would bend a
      // straight line, and the whole claim of an OLS fit is that it is straight.
      // Only y needs it — `toX(x0)`/`toX(x1)` are the plot's own edges.
      const seg = clipToPlotY(toX(x0), toY(at(x0)), toX(x1), toY(at(x1)), plot.y, plot.y + plot.h);
      if (seg) {
        nodes.push({
          kind: "line",
          ...seg,
          stroke: style.negative,
          strokeWidth: 1.25,
          dash: [4, 2],
          name: "trend",
        });
      }
      const stats = trendStats(pts);
      if (stats) {
        const label = `R² = ${stats.r2.toFixed(2)}${stats.p != null ? `, p ${formatP(stats.p)}` : ""}`;
        const endY = toY(at(x1));
        nodes.push({
          kind: "text",
          x: plot.x + plot.w - textWidth(label, fs * 0.9) - 4,
          y: Math.max(plot.y, Math.min(plot.y + plot.h - fs * 1.3, endY + (slope >= 0 ? fs * 0.4 : -fs * 1.7))),
          w: textWidth(label, fs * 0.9) + 4,
          h: fs * 1.3,
          text: label,
          fontSize: fs * 0.9,
          color: style.negative,
          align: "right",
          valign: "middle",
          name: "trend-stats",
        });
      }
    }
  } else if (wantTrend && pts.length >= 3) {
    const fit = polyTrend(pts, trendDeg);
    if (fit && fit.degree >= 2) {
      // Sample within the data's x-range (extrapolating a polynomial past the
      // data is meaningless), clamping y to the plot so an overshoot rides the
      // edge instead of drawing over the rest of the chart.
      const lo = Math.max(x0, Math.min(...pts.map((p) => p.x)));
      const hi = Math.min(x1, Math.max(...pts.map((p) => p.x)));
      const clampY = (y: number) => Math.max(plot.y - 2, Math.min(plot.y + plot.h + 2, toY(y)));
      const SAMPLES = 40;
      let px = toX(lo);
      let py = clampY(fit.at(lo));
      for (let i = 1; i <= SAMPLES; i++) {
        const x = lo + ((hi - lo) * i) / SAMPLES;
        const nx = toX(x);
        const ny = clampY(fit.at(x));
        nodes.push({
          kind: "line",
          x1: px,
          y1: py,
          x2: nx,
          y2: ny,
          stroke: style.negative,
          strokeWidth: 1.25,
          dash: [4, 2],
          name: `trend-seg-${i}`,
        });
        px = nx;
        py = ny;
      }
      const names: Record<number, string> = { 2: "quadratic", 3: "cubic", 4: "quartic" };
      const label = `R² = ${fit.r2.toFixed(2)} · ${names[fit.degree] ?? `deg ${fit.degree}`}`;
      nodes.push({
        kind: "text",
        x: plot.x + plot.w - textWidth(label, fs * 0.9) - 4,
        y: plot.y + 2,
        w: textWidth(label, fs * 0.9) + 4,
        h: fs * 1.3,
        text: label,
        fontSize: fs * 0.9,
        color: style.negative,
        align: "right",
        valign: "top",
        name: "trend-stats",
      });
    }
  }

  // Continuous color legend: a discretized gradient bar with min/max labels
  // (renderer-safe — small rects, no SVG gradient).
  if (colorScale) {
    /**
     * THE COLOR ROW IS ITS OWN METRIC, so it takes only the format's
     * metric-neutral halves.
     *
     * `suffix` and `forceSign` are value-axis concerns — the rule
     * `axisTickLabel` applies to a share axis (which labelled a 100% axis
     * "25 m%" until it did) and the gantt's gutter columns state outright:
     * "suffix/forceSign are deliberately not picked up". This legend took the
     * whole format, so a scatter of revenue declaring `suffix: " €m"` with a
     * `Color` row of NPS or headcount labelled its gradient "1.0 €m" …
     * "4.0 €m" — the y axis's unit stamped on a third variable.
     *
     * `decimals` and `locale` still carry: an authored precision and the
     * chart's number system belong to the whole chart, not to one row.
     */
    const cFmt = resolveFormat([colorScale.min, colorScale.max], {
      decimals: cfg.numberFormat?.decimals,
      locale: cfg.numberFormat?.locale,
    });
    const steps = 24;
    const barW = COLOR_BAR_W;
    const cell = barW / steps;
    const bx = plot.x;
    // The min/max labels hang BELOW the gradient bar, so with a top gutter the
    // legend has to sit a little higher or they land on the marginal bars.
    // Only when the gutter exists, so no existing output moves.
    const by = legendTop - fs * (mTop > 0 ? 1.75 : 1.35);
    for (let i = 0; i < steps; i++) {
      const t = (i + 0.5) / steps;
      nodes.push({
        kind: "rect",
        x: bx + i * cell,
        y: by,
        w: cell + 0.5,
        h: fs * 0.7,
        fill: colorScale.of(colorScale.min + t * (colorScale.max - colorScale.min)),
        name: `color-legend-${i}`,
      });
    }
    const colorName = find(/^colou?r$/i)?.name ?? "Color";
    nodes.push(
      {
        kind: "text",
        x: bx - 40,
        y: by - fs * 0.15,
        w: 38,
        h: fs,
        text: colorName,
        fontSize: fs * 0.85,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "color-legend-title",
      },
      {
        kind: "text",
        x: bx,
        y: by + fs * 0.75,
        w: barW,
        h: fs,
        text: formatNumber(colorScale.min, cFmt),
        fontSize: fs * 0.75,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: "color-legend-min",
      },
      {
        kind: "text",
        x: bx,
        y: by + fs * 0.75,
        w: barW,
        h: fs,
        text: formatNumber(colorScale.max, cFmt),
        fontSize: fs * 0.75,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "color-legend-max",
      },
    );
  }

  // Group legend when points are colored by group. A color scale normally
  // supersedes group coloring and suppresses this — but markers put group on
  // the SHAPE channel, which the color legend says nothing about, so the legend
  // has to come back or the shapes stand unexplained.
  const groupIds = colorScale && !markers ? [] : [...new Set(pts.map((p) => p.group))].sort((a, b) => a - b);
  // `showGroupLegend` gates the DRAW as well as the reservation. Gating only the
  // reservation is the bug this file keeps finding in others — the band is
  // reserved as zero and the legend is drawn anyway, which is strictly worse
  // than drawing it with room. Measured: doing exactly that took the extreme-frame
  // overlap count from 55 to 63.
  /**
   * How far right the group legend's own ink reaches, or 0 when there is none.
   *
   * The bubble size key sits at the top-right of the plot and has to clear this
   * legend — but clearing the whole BAND moves the key on charts where the two
   * are nowhere near each other: an ordinary 480x300 bubble has "Group 1..3"
   * ending around x=190 and the key starting past x=380, and shifting it there
   * changed a chart with nothing wrong with it (the showcase's snapshot is what
   * said so). What the key has to clear is the legend's INK.
   */
  let legendRight = 0;
  if (groupIds.length > 1 && showGroupLegend) {
    // Clear the gradient bar when both legends are up — they share this row and
    // both anchor at plot.x, so without the offset the chips land on the ramp.
    let lx = plot.x + (colorScale ? COLOR_BAR_W + 16 : 0);
    // The same shared walk the reservation used, so the drawer and the reserver
    // cannot disagree about how many rows this legend takes.
    const labels = groupIds.map((g) => `Group ${g}`);
    const walk = (f: number) => legendWrapWalk(labels, f, lx, plot.x + plot.w);
    /** What the whole block occupies, top row's ascent to bottom row's descent. */
    const blockH = (f: number, rows: number) => (rows - 1) * f * 1.8 + f * 1.4;
    /**
     * One size for the whole legend, small enough that the block fits the chart.
     *
     * The rows are as tall as the font and there can be several of them, so a big
     * font on a small frame asks for more legend than there is chart: three
     * wrapped rows at a 32pt font want 173 points of a 150pt-tall scatter. No
     * placement can fix that — the floor below holds the top row on, and the
     * bottom row then leaves — so the size has to give first.
     *
     * Shrinking also un-wraps rows, which is why this re-walks rather than
     * scaling a row count: a narrower label may need one row where it needed
     * three. Last resort as everywhere else — at any font whose legend already
     * fits, this is `fs` and nothing moves.
     *
     * The budget is the canvas LESS the title band and a minimum plot, not the
     * whole canvas. Measured against the canvas, three wrapped rows at a 26pt
     * font were allowed to occupy 140 points of a 150-point chart — which fits,
     * and leaves the title and the plot underneath the legend. A legend is
     * chrome; it may not price itself as though it were the chart.
     */
    const legendBudget = Math.max(fs, cfg.height - titleH - MIN_PLOT_SIDE);
    const legendFs = (() => {
      let f = fs;
      while (f > 5 && blockH(f, walk(f)[labels.length - 1].row + 1) > legendBudget) f -= 0.5;
      return f;
    })();
    const slots = walk(legendFs);
    // Rows stack upward from the plot: the LAST row sits where the single row
    // always did, so a one-row legend is byte-identical to before.
    const rows = slots[slots.length - 1].row + 1;
    // ...and stacking upward is how the top row leaves the chart when the plot
    // has been pulled up to fit a short frame: 13pt above an 80x60 scatter with
    // two legend rows. Hold the block down far enough for its own top row, which
    // on any chart with room above the plot is no constraint at all — and no
    // further down than its own BOTTOM row can afford, which is what the shrink
    // above guarantees is reachable.
    //
    // The floor is the TITLE's bottom, not the canvas top: the title is a
    // full-width band, so a top row merely held on canvas was held straight
    // onto the words. The budget above is what makes this reachable, rather
    // than a clamp that trades one collision for another.
    const legendY = Math.min(
      Math.max(legendTop, titleH + (rows - 1) * legendFs * 1.8 + legendFs * 1.55),
      cfg.height - legendFs * 1.4 + legendFs * 1.55,
    );
    groupIds.forEach((g, gi) => {
      const chip = legendFs * 0.7;
      const label = `Group ${g}`;
      lx = slots[gi].x;
      const dy = -(rows - 1 - slots[gi].row) * legendFs * 1.8;
      // Under a color scale the chip's color would be a lie (color means the
      // Color row there), so the shape carries the legend in neutral ink.
      const chipFill = colorScale ? style.mutedText : paletteColor(cfg.style?.palette ?? PALETTE, g - 1);
      // The chip is drawn like the points it explains, area and all — so an
      // area-matched star is 1.67x wider than `chip` and would sit on its own
      // label. Advance by what was actually drawn. Without markers the drawn
      // width IS chip, so the default legend's spacing is untouched.
      const drawn = markers ? markerExtent(markerFor(g), chip / 2) * 2 : chip;
      nodes.push(
        markers
          ? markerNode(
              markerFor(g),
              lx + drawn / 2,
              legendY + dy - legendFs * 1.2 + chip / 2,
              chip / 2,
              chipFill,
              style.background,
              0,
              `legend-chip-${g}`,
            )
          : {
              kind: "rect",
              x: lx,
              y: legendY + dy - legendFs * 1.2,
              w: chip,
              h: chip,
              fill: chipFill,
              name: `legend-chip-${g}`,
            },
        {
          kind: "text",
          x: lx + drawn + 3,
          y: legendY + dy - legendFs * 1.55,
          w: textWidth(label, legendFs) + 6,
          h: legendFs * 1.4,
          text: label,
          fontSize: legendFs,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `legend-${g}`,
        },
      );
      legendRight = Math.max(legendRight, lx + drawn + 3 + textWidth(label, legendFs));
    });
  }

  // Bubble radius: area ∝ size, max radius 9% of the smaller plot dimension.
  const maxSize = Math.max(1e-9, ...pts.map((p) => Math.abs(p.size ?? 0)));
  const maxR = Math.min(plot.w, plot.h) * 0.09;
  const radius = (p: (typeof pts)[number]) =>
    cfg.kind === "bubble" && p.size != null ? Math.max(2.5, Math.sqrt(Math.abs(p.size) / maxSize) * maxR) : 3;

  // Overlap relief along one axis. The cross axis stays exact, the named one
  // moves by at most the disclosed cap, and every marker keeps its identity —
  // this only shifts where a marker is drawn, never which point it is.
  const cap = spreadCap(cfg);
  const spread = new Map<number, number>();
  if (cap) {
    const byX = cap.axis === "x";
    // The cap is quoted in data units; the relaxation works in px.
    const limitPx = byX ? (cap.limit / (x1 - x0 || 1)) * plot.w : (cap.limit / (y1 - y0 || 1)) * plot.h;
    const disp = spreadAlongAxis(
      pts.map((p) => ({ m: byX ? toX(p.x) : toY(p.y), c: byX ? toY(p.y) : toX(p.x), r: radius(p) })),
      byX
        ? { limit: limitPx, min: plot.x, max: plot.x + plot.w }
        : { limit: limitPx, min: plot.y, max: plot.y + plot.h },
    );
    disp.forEach((d, i) => spread.set(i, d));
  }
  /**
   * Drawn position: the exact one, plus any disclosed nudge on the spread axis,
   * held far enough inside the chart for the GLYPH to fit.
   *
   * A mark is drawn around its position, so a point at the very top of the plot
   * puts half its marker above the plot — harmless while the plot has chrome
   * over it, and not harmless when the frame cannot pay for that chrome and
   * `fitPlot` brings the plot's top edge up against the frame's: 2.2pt of a
   * point's disc outside an 80x60 chart at 32pt. The clamp is by the glyph's
   * own drawn extent (a star reaches 1.67x its data radius) and against the
   * FRAME, so a marker that already fits inside the chart does not move at all
   * and no ordinary chart changes.
   *
   * A clamped mark is a point drawn a couple of points off its value, which is
   * the same trade `fitPlot` makes for the plot itself; the alternative is a
   * mark drawn onto whatever sits beside the chart on the slide, since neither
   * PowerPoint renderer clips.
   */
  const inset = (p: (typeof pts)[number]) => markerExtent(markerFor(p.group), radius(p));
  const px = (p: (typeof pts)[number], i: number) => {
    const e = inset(p);
    const at = toX(p.x) + (cap?.axis === "x" ? (spread.get(i) ?? 0) : 0);
    return Math.min(Math.max(at, e), Math.max(e, cfg.width - e));
  };
  const py = (p: (typeof pts)[number], i: number) => {
    const e = inset(p);
    const at = toY(p.y) + (cap?.axis === "y" ? (spread.get(i) ?? 0) : 0);
    return Math.min(Math.max(at, e), Math.max(e, cfg.height - e));
  };

  // Bubble size legend: without a key, bubble AREA is unreadable. Two
  // outline reference circles (a nice maximum and its half), top-right.
  const legendBoxes: Box[] = [];
  // A `Size` row that is entirely zero has nothing to key: `maxSize` is floored
  // to a tiny epsilon so the ratios stay finite, so the two reference circles
  // came out different sizes and BOTH were labelled "0.00", over a plot where
  // every bubble sits at the 2.5pt floor. A legend that contradicts itself and
  // the marks it explains is worse than none.
  if (cfg.kind === "bubble" && pts.some((p) => Math.abs(p.size ?? 0) > 0)) {
    // The `Size` row is its own metric too — see the color legend above for why
    // `suffix`/`forceSign` are left behind. This key read "600 €m" / "300 €m"
    // for a headcount, and "+600" / "+300" under `forceSign`.
    const sizeFmt = resolveFormat(
      pts.map((p) => Math.abs(p.size ?? 0)),
      { decimals: cfg.numberFormat?.decimals, locale: cfg.numberFormat?.locale },
    );
    const refMax = niceTicks(0, maxSize, 3).pop()!;
    const refs = [refMax, refMax / 2];
    // A size legend needs room for its widest NUMBER, not just its circles. The
    // label box is `r * 2` centred on a reference circle, and a small circle
    // carries a number far wider than itself — so on a narrow plot the text
    // spills left out of the legend and across the y-axis ticks and the group
    // legend beside it.
    //
    // Dropped whole rather than shrunk: this legend is a reference scale, and
    // two circles whose values cannot be read explain nothing. Half the plot is
    // the bound because the legend sits at the plot's right edge — a label
    // reaching past the midpoint is into the chart, not beside it.
    const widestRef = Math.max(...refs.map((v) => textWidth(formatNumber(v, sizeFmt), fs * 0.8)));
    let sizeLegendFits = widestRef <= plot.w * 0.5;
    // The numbers hang `fs * 1.35` above the tallest reference circle, and
    // nothing held them inside the chart: on a 300x60 frame at 32pt they were
    // drawn 33pt above its top edge while the circles themselves sat inside it.
    // The whole key moves down by the shortfall rather than being dropped — it
    // already overlays the top of the plot by design, so a legend a few points
    // lower is the same legend, where a dropped one leaves the bubble areas
    // unreadable. Zero on any chart whose numbers already fit, so nothing that
    // fits moves.
    const refMaxR = Math.max(2.5, Math.sqrt(refMax / maxSize) * maxR);
    // Shifted clear of the TITLE, not merely onto the canvas. Holding the
    // numbers at y >= 0 stopped them leaving the chart and left them lying
    // across its title on a 120x90 frame at 18pt, which is the same trade the
    // upright column totals used to make with their clamp.
    // Clear of the TITLE always, and of the GROUP LEGEND only where the two
    // actually meet: the legend fills its row from the left and the key sits at
    // the right, so on a wide chart they never touch, and clearing the legend's
    // whole band there moved a key that was perfectly well placed. On a narrow
    // one the legend runs under the key and its numbers landed on "Group 3".
    const keyLeft = plot.x + plot.w - 4 - 4 * refMaxR;
    const clearTo = legendRight > keyLeft ? chromeTop : titleInkBottom(cfg, style);
    const legendShift = Math.max(0, clearTo + refMaxR + fs * 1.35 - fs * 0.36 - (plot.y + maxR * 1.1));
    // The shift has a floor as well as a ceiling. Pushed far enough to clear the
    // group legend it ran off the foot of a 300x60 chart at 24pt, and then onto
    // the x-axis tick labels at 14pt — one collision traded for the next. The
    // key belongs INSIDE the plot, so that is what is asked: where the shifted
    // circles do not fit between the plot's own top and bottom, the key is
    // dropped, exactly as the width bound above drops it.
    if (maxR * 1.1 + legendShift + refMaxR > plot.h) sizeLegendFits = false;
    let lx = plot.x + plot.w - 4;
    if (sizeLegendFits)
      refs.forEach((v, i) => {
        const r = Math.max(2.5, Math.sqrt(v / maxSize) * maxR);
        const cx = lx - r;
        const cy = plot.y + maxR * 1.1 + legendShift + (Math.sqrt(refMax / maxSize) * maxR - r); // bottom-aligned circles
        nodes.push({
          kind: "ellipse",
          cx,
          cy,
          rx: r,
          ry: r,
          fill: "none",
          stroke: style.mutedText,
          strokeWidth: 1,
          name: `size-legend-${i}`,
        });
        // THE NUMBER IS FITTED TO ITS OWN CIRCLE, and dropped below the floor.
        //
        // The box is `r * 2` — the circle's diameter — and the text was always
        // `fs * 0.8`. The circles are nested reference sizes, so the smallest is
        // a 5pt box, and a centred number wider than its box spills equally off
        // both sides onto the neighbouring key's number. 34 pairs in the variant
        // sweep, and the same defect the gauge's centre total had: text drawn
        // wider than the box it is centred in.
        //
        // Shrink, then drop, as everywhere else. What is lost is one reference
        // number on the smallest circle, which is the least informative of the
        // three — the circle itself still shows the size, and the largest and
        // middle keys still carry their values.
        const keyFs = bandFontSize(fs * 0.8, r * 2, textWidth(formatNumber(v, sizeFmt), 1));
        if (keyFs > 0)
          nodes.push({
            kind: "text",
            x: cx - r,
            y: cy - Math.sqrt(refMax / maxSize) * maxR - fs * 1.35,
            w: r * 2,
            h: fs * 1.2,
            text: formatNumber(v, sizeFmt),
            fontSize: keyFs,
            color: style.mutedText,
            align: "center",
            valign: "bottom",
            name: `size-legend-label-${i}`,
          });
        legendBoxes.push({ x: cx - r, y: cy - r - fs * 1.4, w: r * 2, h: r * 2 + fs * 1.4 });
        lx = cx - r - fs * 0.8;
      });
  }

  // Both drawn BEHIND the markers, which is why they sit here and not later:
  // `nodes` is z-ordered and the push order is the paint order.
  if (decor.trajectory && pts.length > 1) nodes.push(...trajectoryNodes(pts, xAxis, yAxis, style, cfg));
  if (mTop > 0 || mRight > 0)
    nodes.push(
      ...marginalNodes(
        pts,
        xAxis,
        yAxis,
        plot,
        { top: mTop, right: mRight },
        lerpColor(style.background, (cfg.style?.palette ?? PALETTE)[0], 0.35),
      ),
    );

  // Point labels treat the size legend as an obstacle.
  //
  // The AXIS labels are deliberately NOT obstacles, and that was measured rather
  // than assumed. Adding them removes 35 overlapping pairs at large fonts — and
  // makes the placer DROP point labels it can no longer position, including on a
  // comfortable 480x300 chart where a point sits near the left edge and the y
  // axis owns that margin. A point's label is data and a tick label is chrome;
  // losing the first to protect the second is the wrong way round, and hiding is
  // what this placer does when it runs out of room, so the cost lands exactly
  // where it hurts most.
  const markerBoxes: Box[] = [...legendBoxes];
  // Paint back-to-front. Emitting in datasheet order lets a large bubble drawn
  // early bury a smaller one drawn later — completely, in all three renderers,
  // with nothing to tell the reader a point is missing. Largest first puts the
  // big ones at the back. This is paint order only: every marker keeps its
  // datasheet index in its name and its exact position. Ties keep datasheet
  // order (Array.sort is stable), so the result stays deterministic. Labels
  // already run biggest-first below, for the same reason.
  const paintOrder = pts.map((_, i) => i).sort((a, b) => radius(pts[b]) - radius(pts[a]));
  for (const i of paintOrder) {
    const p = pts[i];
    const r = radius(p);
    const gi = p.group - 1;
    const fill =
      colorScale && p.color != null
        ? colorScale.of(p.color)
        : colorScale
          ? style.mutedText
          : paletteColor(cfg.style?.palette ?? PALETTE, gi);
    nodes.push(markerNode(markerFor(p.group), px(p, i), py(p, i), r, fill, style.background, 1, `point-${i}`));
    // Keep labels off the mark as DRAWN: an area-matched star reaches ~1.67x
    // its data radius, and a keep-out box built from `r` would let a label sit
    // on its points.
    const mr = markerExtent(markerFor(p.group), r);
    markerBoxes.push({ x: px(p, i) - mr, y: py(p, i) - mr, w: mr * 2, h: mr * 2 });
  }

  // Greedy label placement, biggest bubbles first so important points win.
  if (decor.segmentLabels !== false) {
    const order = pts.map((_, i) => i).sort((a, b) => radius(pts[b]) - radius(pts[a]));
    // Point label content: category (default), optionally with "(x, y)".
    const pointLabel = (p: (typeof pts)[number]) => {
      // A LIST, and a config that wrote a bare `"value"` instead of `["value"]`
      // threw `parts.map is not a function`. Same coercion as `segmentLabel`,
      // which is the other consumer of this key — scatter builds its own label
      // rather than going through it, so it needs its own guard.
      const raw = decor.labelContent ?? ["category"];
      const parts = Array.isArray(raw) ? raw : typeof raw === "string" ? [raw] : [];
      const fmt = resolveFormat([p.x, p.y], cfg.numberFormat);
      return parts
        .map((part) =>
          part === "category"
            ? p.label
            : part === "value"
              ? `(${formatNumber(p.x, fmt)}, ${formatNumber(p.y, fmt)})`
              : null,
        )
        .filter(Boolean)
        .join(" ");
    };
    // Anchored to the mark as DRAWN (px/py), not to the undisplaced position:
    // beeswarm relief exists precisely to separate co-located points, so the
    // spot a spread point vacated is typically taken by a NEIGHBOUR — and a
    // label placed there named one bubble while sitting on another.
    const reqs: LabelRequest[] = order.map((i) => ({
      cx: px(pts[i], i),
      cy: py(pts[i], i),
      r: radius(pts[i]),
      w: textWidth(pointLabel(pts[i]), fs) + 2,
      h: fs * 1.3,
    }));
    // The band is a line and a half TALLER than the plot on purpose, and that
    // extra strip is where the x-axis tick labels live — so a point label at
    // the foot of the plot can be drawn on one. Confining the band to the plot
    // was tried and MEASURED: it takes the overlapping-text count for these two
    // kinds from 889 to 599 and drops 56 of 301 point labels, on charts as
    // comfortable as 480x300. Same verdict as giving this placer the axis
    // labels as obstacles, and for the same reason — a point's label is DATA
    // and a tick label is chrome, so the chrome yields.
    // The band starts BELOW the title's ink. Confining it to the plot was tried
    // and measured and is refused (it drops 56 of 301 point labels on a
    // comfortable 480x300 — a point's label is data, a tick label is chrome);
    // this is the other bound, and it only bites when the plot itself has been
    // squeezed into the title, where the alternative is a number printed across
    // the chart's own name.
    const bandTop = Math.max(plot.y, titleInkBottom(cfg, style));
    const band = { x: 0, y: bandTop, w: cfg.width, h: plot.y + plot.h + fs * 1.5 - bandTop };
    /**
     * Tick numbers as SOFT obstacles: avoided if there is anywhere else to go,
     * overwritten if there is not.
     *
     * The note above records that giving this placer the axis numbers as
     * obstacles was tried and refused — it drops 56 of 301 point labels on a
     * comfortable chart, and a point's label is data where a tick number is
     * chrome. That verdict is about HARD obstacles, and it stands. This is the
     * third option neither pass had: place what can dodge a tick number first,
     * then place the rest with the ticks not counted at all.
     *
     * NEARLY nothing is lost by it, and the honest word is "nearly". Every
     * request that fails the first pass is offered the second, whose obstacles
     * are the old set plus what pass one has already put down — but pass one
     * puts labels in different SLOTS than a single pass would, so a later label
     * can find its spot taken, and can equally find one freed. Measured across
     * the whole sweep at seven fonts: 266 point labels become 265 on a scatter,
     * 265 become 267 on a bubble. Two either way out of five hundred.
     *
     * What it buys is a third of the loss back: of 423 tick numbers, dropping
     * the covered ones alone leaves 310, and dodging first leaves 341. Per
     * chart it is starker than the total — a 200x150 scatter at 18pt keeps 5 of
     * its 10 rather than 2, and a 120x90 keeps 10 rather than 5.
     */
    const tickBoxes: Box[] = nodes
      .filter((n): n is TextNode => n.kind === "text" && (n.name === "x-axis" || n.name === "y-axis") && !!n.text)
      .map(tightBox);
    const firstPass = placeLabels(reqs, band, [...markerBoxes, ...tickBoxes]);
    const dodged = new Set(firstPass.map((p) => p.index));
    const restIndex: number[] = [];
    const rest: LabelRequest[] = [];
    reqs.forEach((r, i) => {
      if (!dodged.has(i)) {
        restIndex.push(i);
        rest.push(r);
      }
    });
    const secondPass = placeLabels(rest, band, [...markerBoxes, ...firstPass.map((p) => p.box)]).map((p) => ({
      ...p,
      index: restIndex[p.index],
    }));
    for (const placed of [...firstPass, ...secondPass]) {
      const p = pts[order[placed.index]];
      nodes.push({
        kind: "text",
        x: placed.box.x,
        y: placed.box.y,
        w: placed.box.w,
        h: placed.box.h,
        text: pointLabel(p),
        fontSize: fs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `label-${order[placed.index]}`,
      });
    }
  }

  /**
   * THE FIT'S STATISTICS YIELD TO THE POINT LABELS, and here is the only place
   * they can.
   *
   * `trend-stats` is placed at the plot's right edge on the trend line's end and
   * is fitted against nothing: the point labels do not exist yet when it is
   * pushed, six hundred lines above. Nor can it fall back on the global
   * de-collision pass — scatter is a `decorlessKind`, so
   * `resolveLabelCollisions` never runs for this layout at all. Adding
   * `trend-stats` to that pass's MOVABLE list was tried first and is dead code
   * for exactly that reason.
   *
   * So it settles here, once everything it could hit exists. It is the label
   * that gives way because it is the only one in the plot naming no datum: a
   * point's label says which observation this is, where R² describes the fit and
   * reads correctly from anywhere inside the plot.
   *
   * Found in the shipped showcase rather than by a sweep — `Scatter — partition
   * lines, trend, groups` and the quadratic-fit chart both drew it through a
   * point label at 480x300. See `test/showcase-overlap.test.ts`.
   */
  const stat = nodes.find((n): n is TextNode => n.kind === "text" && n.name === "trend-stats");
  if (stat) {
    const boxes = nodes
      .filter((n): n is TextNode => n.kind === "text" && n !== stat && !!String(n.text).trim())
      .map(tightBox);
    const hits = () => {
      const b = tightBox(stat);
      return boxes.some((o) => b.x < o.x + o.w && b.x + b.w > o.x && b.y < o.y + o.h && b.y + b.h > o.y);
    };
    if (hits()) {
      const from = stat.y;
      // Both ways, nearest first: the box already picks a side from the trend's
      // slope, so neither direction is the wrong one. Clamped into the plot,
      // because a statistic outside the plot labels nothing.
      const room = [1, -1, 2, -2, 3, -3]
        .map((k) => from + k * fs * 1.3)
        .find((y) => {
          if (y < plot.y || y + stat.h > plot.y + plot.h) return false;
          stat.y = y;
          return !hits();
        });
      // Restored when nothing is clear, for the reason the nudge pass restores:
      // a label that moved and still collides has been put somewhere nobody
      // chose.
      stat.y = room ?? from;
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: pts.map((p) => toX(p.x)),
      categoryWidth: pts.map((p) => radius(p) * 2),
      columnTop: pts.map((p) => toY(p.y)),
      columnValue: pts.map((p) => p.y),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
