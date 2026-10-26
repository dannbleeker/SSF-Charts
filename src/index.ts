/**
 * SSF Charts library entry: the pure chart engine, consumable outside the
 * add-in (batch SVG rendering, .ppttc-style automation pipelines, tests).
 */
export { buildChart, valueExtent, DEFAULT_SIZE } from "./core/chart";
export { sceneToSvg } from "./render/svg";
export { buildAgendaScene, SLIDE } from "./core/agenda";
export { buildHarveyBall, buildCheckbox, buildProcessFlow, buildTableScene, buildKpiTile } from "./core/elements";
export { sampleConfig, CHART_KINDS } from "./core/samples";
export { PALETTE, PALETTES, DEFAULT_STYLE } from "./core/style";
export { formatNumber, formatPercent, parseDateToken, cagr } from "./core/format";
// Shared renderer geometry — the headless pptx renderer (skill/scripts/render-pptx.mjs)
// consumes these from the built lib instead of re-deriving them.
export {
  arrowheadBox,
  sceneToOoxmlPieAngle,
  annularSectorPoints,
  symbolPoints,
  SYMBOL_PRESET,
  symbolPreset,
  dashKind,
} from "./core/geometry";
// The OOXML post-processor. pptxgenjs can write neither a real shape group nor
// a tag part, so the headless renderer (skill/scripts/render-pptx.mjs) runs its
// output through the same injector the add-in uses — one implementation, so a
// chart Claude generates opens in the pane exactly like one the pane drew.
export { injectGroupsAndTags } from "./render/ooxml";
export type { SlideDressing } from "./render/ooxml";
export type { ChartConfig, ChartData, ChartKind, Decorations, Series } from "./core/types";
export type { Scene, SceneNode } from "./core/scene";
