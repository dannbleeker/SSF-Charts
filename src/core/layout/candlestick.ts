import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { type SceneNode } from "../scene";
import { chromeNodes, computeFrame, valueScale } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Candlestick / OHLC chart. Datasheet rows named Open / High / Low / Close
 * give each period's prices; categories are the periods. A thin high–low wick
 * carries a body from open to close. Direction is encoded REDUNDANTLY: a rising
 * period is a hollow body (background fill, coloured outline) and a falling one
 * is solid — the classic convention, so up/down survives greyscale printing and
 * red-green colour blindness, not colour alone.
 */
export function layoutCandlestick(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()))?.values ?? [];
  const open = find(/^open$/i);
  const high = find(/^high$/i);
  const low = find(/^low$/i);
  const close = find(/^close$/i);

  const all = [...open, ...high, ...low, ...close].filter((v): v is number => v != null);
  const { frame } = computeFrame(cfg, style, { ...decor, seriesLabels: false }, []);
  const scale = valueScale(
    frame,
    Math.min(...(all.length ? all : [0])),
    Math.max(...(all.length ? all : [1])),
    cfg.scale,
    undefined,
    undefined,
    false, // no forced zero: OHLC prices rarely include 0
  );
  const slotLen = frame.w / Math.max(1, n);
  const colThick = Math.min(slotLen * 0.5, 24);
  const centers = data.categories.map((_, c) => frame.x + slotLen * (c + 0.5));

  const nodes: SceneNode[] = chromeNodes(cfg, style, { ...decor, seriesLabels: false }, frame, centers, scale);
  const up = "#1a9e6e";
  const down = style.negative;

  for (let c = 0; c < n; c++) {
    const hi = high[c];
    const lo = low[c];
    const x = centers[c];
    if (hi != null && lo != null) {
      nodes.push({
        kind: "line",
        x1: x,
        y1: scale.toY(hi),
        x2: x,
        y2: scale.toY(lo),
        stroke: style.text,
        strokeWidth: 1,
        name: `wick-${c}`,
      });
    }
    const o = open[c];
    const cl = close[c];
    if (o != null && cl != null) {
      const rising = cl >= o;
      const yTop = scale.toY(Math.max(o, cl));
      const yBot = scale.toY(Math.min(o, cl));
      // Rising = hollow (background fill, coloured outline); falling = solid.
      // Direction reads without colour, then colour reinforces it.
      nodes.push({
        kind: "rect",
        x: x - colThick / 2,
        y: yTop,
        w: colThick,
        h: Math.max(1, yBot - yTop),
        fill: rising ? style.background : down,
        stroke: rising ? up : down,
        strokeWidth: rising ? 1 : 0.5,
        name: `body-${c}`,
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => colThick),
      // Top of the period's drawn ink. A blank High must NOT fall back to
      // value 0: this scale is deliberately zero-free, so toY(0) lands
      // thousands of points below the frame and drags every callout / CAGR /
      // difference anchor off the canvas with it. Use the highest price the
      // period actually has; a wholly empty period sits on the plot floor.
      columnTop: data.categories.map((_, c) => {
        const priced = [high[c], open[c], close[c], low[c]].filter((v): v is number => v != null);
        return priced.length ? scale.toY(Math.max(...priced)) : frame.y + frame.h;
      }),
      columnValue: data.categories.map((_, c) => close[c] ?? 0),
      // A period with no Close is a period that has not closed (today's, a
      // holiday), and `columnValue` above answers 0 for it. The mean line and the
      // difference and CAGR arrows read that 0 as a close of nothing — "Ø 36" for
      // closes of 44, 43, -, 45, 47 whose mean is 44.75, and "-100%" on an arrow to
      // the open period — unless this says it was not measured. The Close is the
      // question because it is what `columnValue` is.
      columnHasData: data.categories.map((_, c) => close[c] != null),
      baselineY: frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: scale.toY,
    },
  };
}
