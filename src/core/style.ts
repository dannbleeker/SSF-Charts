import type { ChartStyle, Decorations } from "./types";
import { layoutBackground } from "./color";

/**
 * Default series palette: the validated brand-neutral categorical set from the
 * dataviz reference (worst adjacent CVD ΔE 24.2 on a white surface; the three
 * sub-3:1-contrast hues are relieved by the engine's always-on direct labels).
 * Swap for the presentation's theme colors via ChartConfig.style.palette.
 */
export const PALETTE = [
  "#2a78d6", // blue
  "#1baf7a", // aqua
  "#eda100", // yellow
  "#008300", // green
  "#4a3aa7", // violet
  "#e34948", // red
  "#e87ba4", // magenta
  "#eb6834", // orange
];

/**
 * Palette presets. "Default" is the validated categorical set; the others are
 * corporate looks — swap in the presentation's accent colors via a custom
 * palette when needed (Office.js exposes no theme-color API).
 */
export const PALETTES: Record<string, string[]> = {
  Default: PALETTE,
  Blues: ["#0d366b", "#1c5cab", "#2a78d6", "#5598e7", "#86b6ef", "#b7d3f6", "#52514e", "#898781"],
  Warm: ["#b3541e", "#eb6834", "#eda100", "#e34948", "#8a6d3b", "#d5a021", "#52514e", "#898781"],
  Grayscale: ["#1a1a19", "#3d3d3b", "#5f5e5b", "#82817d", "#a5a49e", "#c3c2b7", "#dcdbd2", "#8a3ffc"],
};

export const DEFAULT_STYLE: ChartStyle = {
  fontFamily: "Segoe UI, Arial, sans-serif",
  fontSize: 10,
  palette: PALETTE,
  text: "#0b0b0b",
  mutedText: "#52514e",
  axis: "#a5a49e",
  gridline: "#e1e0d9",
  negative: "#e34948",
  neutral: "#898781",
  background: "#ffffff",
};

/**
 * The style a layout is handed: the defaults under the chart's own `style`, with
 * `background` resolved to what a layout can lay itself out against.
 *
 * `background` goes through `layoutBackground`, so every spelling of "no canvas"
 * (`transparent`, `none`, a zero-alpha `rgba()`/`#0000`/`hsla()`, a typo, an empty
 * string) reads the same way at the dozen direct `lerpColor(style.background, …)`
 * sites in the layouts — see there for what it answers and why. The CANVAS paint
 * is not read from this style, so the chart's own `style.background` still decides
 * what the SVG and the slide are filled with.
 */
export function resolveStyle(over?: Partial<ChartStyle>): ChartStyle {
  const style: ChartStyle = { ...DEFAULT_STYLE, ...over };
  style.background = layoutBackground(style.background);
  return style;
}

export const DEFAULT_DECOR: Decorations = {
  segmentLabels: true,
  seriesLabels: true,
  totals: false,
  categoryAxis: true,
  valueAxis: false,
  gridlines: false,
};

/**
 * The colour a series is painted: its own override, else its slot in the
 * palette.
 *
 * Routed through `paletteColor` rather than indexing directly, so the guard
 * that function's docstring already advertises actually covers the PRIMARY
 * colour path. Indexing an empty palette is `index % 0` → NaN → `undefined`,
 * written straight into `RectNode.fill`, which the scene contract types as a
 * string — and the three renderers then disagree about what an absent fill is
 * (black in SVG and pptx, mid grey through `officeHex`). `style.palette: []`
 * arrives from the pane's style import, the JSON box, a POWERCHART_CONFIG tag
 * and the skill's caller; the hostile-input sweep already feeds it, but its bar
 * is "does not throw", so the wrong output went unpinned.
 */
export function seriesColor(style: ChartStyle, index: number, override?: string): string {
  return override ?? paletteColor(style.palette, index);
}

/**
 * Wrap `index` into `palette` by its ACTUAL length. Scatter's group colouring
 * hardcoded `% 8` (the default palette length), so a custom palette shorter than
 * 8 handed out `undefined` for high group ids, and a longer one wrapped early —
 * the same modulo-length bug fixed for the violin. Guards an empty palette.
 */
export function paletteColor(palette: string[], index: number): string {
  const n = palette.length;
  return n ? palette[((index % n) + n) % n] : "#888888";
}
