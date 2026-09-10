import { buildChart, clampDim, DEFAULT_SIZE, valueExtent } from "../core/chart";
import { worthOwnSlide, offerSentence, insertOutcomeSentence } from "../core/insert-cost";
import { PALETTES } from "../core/style";
import type { ChartConfig, ChartKind, Decorations, Series } from "../core/types";
import { resolveStyleFile, isEmptyStyle, type DeckStyle, type StylePreference } from "../core/deck-style";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import { sceneToSvg } from "../render/svg";
import {
  roundEnvironment,
  canInsertPicture,
  getSelectionBounds,
  addSlideForChart,
  getSlideShapeBounds,
  insertAgendaSlides,
  insertDemoDeck,
  insertSceneIntoSlide,
  isPowerPointHost,
  newRunId,
  applyReconcilePlan,
  canInsertSlidesFromBase64,
  insertSlidesFromPptx,
  chartOnItsOwnSlideAsFile,
  OFFSCREEN_BATCH,
  replaceSlideWithDeck,
  slideHoldsOnlyChart,
  withSlideDeselected,
  slideCount,
  readAddedSlides,
  enrichSnapshots,
  traceEnvironment,
  wantsAutoPicture,
  marksThisHostWillDrop,
  listChartsInDeck,
  listChartsInSelection,
  scanIsComplete,
  scanGap,
  loadChartFromSelection,
  loadThemePalette,
  readDeckStyle,
  readDeckStyleWithReason,
  warmCustomXmlSurface,
  writeDeckStyle,
  updateChartInSlide,
  updateChartsInSlides,
  onLateSync,
  errorText,
  wreckageOf,
  deleteShapesById,
  requestStop,
  resetStop,
  isStopRequested,
  isStopped,
  slideSize,
  deckSlideIds,
  deleteSlideById,
  dropShapeSelection,
  slideShots,
  type EditTarget,
  type InsertPhase,
  type ReconcileOutcome,
  type SlideInventory,
  type SlideShot,
  type UpdateWreckage,
  namedShape,
  refreshNamedShapeFromDeck,
  syncsSoFar,
  resetSyncCount,
} from "../render/powerpoint";
import { placeChart, type Placement } from "../core/placement";
import { buildAgendaScene } from "../core/agenda";
import { demoItems, buildResultsScenes, type ResultRow, type ResultsSummary } from "../core/demo";
import type { Scene } from "../core/scene";
import { estimateOfficeShapes } from "../core/scene";
import { describeReconcile, planReconcile } from "../core/reconcile";
import {
  formatTraceLine,
  onTrace,
  setTracing,
  trace,
  traceLog,
  traceMark,
  tracing,
  type TraceSummary,
} from "../core/trace";
import {
  beginCrashLog,
  clearCrashLog,
  endCrashLog,
  flushCrashLog,
  markCrashLogSaved,
  recordCrashFinding,
  recordCrashStep,
  recoverCrashLog,
} from "./crashlog";
import {
  runSelfTest,
  describeSelfTest,
  selfTestNeedsAttention,
  setSelfTestRasterizer,
  setSelfTestPrompt,
  SCENARIO_NAMES,
  type ScenarioResult,
} from "./selftest";
import { buildDeckBase64 } from "../render/pptx-deck";
import type { ExpectedItem, SlideSnapshot } from "../core/reconcile";
import { buildTableScene } from "../core/elements";
import { localizePane, localizeTree, t } from "./i18n";
import { contiguousStacks, dataToSheet, mountDatasheet, sheetToData, type SheetModel } from "./datasheet";
import { BUILTIN_TEMPLATES } from "./templates";
import { harveyScene, checkScene, flowScene, kpiScene, wireElementPreviews } from "./elements-ui";
import { agendaChapters, wireAgendaPreview } from "./agenda-ui";
import { EXPERIMENTS, runExperiment } from "../render/experiments";
import {
  runHostProbes,
  describeHostSheet,
  sheetNeedsAttention,
  reaskPlan,
  mergeHostSheets,
  type HostAnswerSheet,
} from "../render/host-probe";

interface AppState {
  kind: ChartKind;
  sheet: SheetModel;
  decorations: Partial<Decorations>;
  horizontal: boolean;
  title: string;
  /**
   * Frame size in points. The single source of truth — the size inputs write
   * here and currentConfig() reads here. It used to be read straight off the DOM
   * inputs while every other field lived in state, and that split silently
   * resized loaded charts to the previous chart's dimensions.
   */
  width: number;
  height: number;
  segmentOrder: NonNullable<ChartConfig["segmentOrder"]>;
  scaleMin: string;
  scaleMax: string;
  breakFrom: string;
  breakTo: string;
  decimals: string; // "auto" | "0" | "1" | "2"
  suffix: string;
  locale: string;
  labelContent: string; // comma-joined parts, "" = default
  paletteName: string;
  /**
   * Style carried in from a loaded chart — fonts, negative/neutral, and a
   * palette matching no preset. Overrides the corporate style file's defaults;
   * `paletteName` still wins over `style.palette` once the user picks one.
   */
  style?: NonNullable<ChartConfig["style"]>;
  /** Per-series color overrides, by ROW INDEX (see seriesMeta). */
  seriesColors: (string | undefined)[];
  /**
   * Series fields the datasheet can't carry (combo `type`, hatch `pattern`,
   * per-point `colors`). The pane rebuilds ChartConfig from the sheet on every
   * preview/insert — without this these fields would be silently dropped from an
   * imported chart and destroyed on re-save.
   *
   * Indexed by ROW, not by series NAME: renaming a series in the datasheet is the
   * sheet's core edit, and a name key meant the rename found no entry, so the
   * overlay line collapsed back into a column and lost its colour/pattern/
   * scenario. (Two series sharing a name collapsed onto one entry for the same
   * reason.) The row index is what the datasheet actually preserves.
   */
  seriesMeta: (Pick<Series, "type" | "pattern" | "colors" | "scenario"> | undefined)[];
  axisTitle: string;
  logScale: boolean;
  /** `render: "image"` — insert one raster picture instead of native shapes.
   *  Owns the key outright now that #render-image exists, so it leaves `extras`. */
  renderImage: boolean;
  /** Footnote / source line ("Kilde: …"). */
  footnote: string;
  /** Comma-separated slice indices to explode (pie/doughnut), 1-based in the UI. */
  pieExplode: string;
  /**
   * Kind-specific / advanced config that has no pane control, preserved verbatim
   * across a load→edit→export (and the shape-tag re-save). Without this, importing
   * an authored chart and re-saving it silently strips these shipped features.
   * pie / waterfall / numberFormat are carried here too but MERGED with their
   * pane-driven parts (explode, total "e" tokens, decimals/suffix/locale) in
   * currentConfig, so the control still wins for the fields it owns.
   */
  extras: Pick<
    ChartConfig,
    | "boxplot"
    | "heatmap"
    | "map"
    | "combo"
    | "gapWidth"
    | "overlap"
    | "multiples"
    | "scatter"
    | "radar"
    | "gantt"
    | "butterfly"
    | "tilemap"
    | "otherBucket"
    | "pareto"
    | "categorySort"
    | "secondaryAxis"
    | "labelOffsets"
    | "pie"
    | "waterfall"
    | "numberFormat"
    | "labels"
  >;
  /** When set, "Update chart" replaces this shape in place. */
  editTarget: EditTarget | null;
}

/**
 * Whether NEW charts default to a picture — the standing half of the choice.
 *
 * `Insert as picture` has always existed as a per-chart tick, and it still is
 * one: loading a chart that went in as a picture shows it ticked, because that
 * is a fact about that chart rather than a preference. What was missing was
 * anywhere to say "I always want this", so the box reset on every new chart and
 * anyone who wanted pictures had to remember, every time.
 *
 * A DEFAULT, NOT AN OVERRIDE, and that distinction is the design. It seeds a new
 * chart's state and nothing else: open an existing chart and that chart's own
 * `render` still wins, because otherwise the preference would silently rewrite
 * work already sitting on a slide.
 *
 * Default OFF, because a picture is not editable in PowerPoint and a preference
 * that quietly turns every chart into a flat image is not one to acquire by
 * accident.
 *
 * DECLARED HERE, ABOVE ITS FIRST USE, and not down with the other storage keys.
 * `rememberedPicturePref` runs at module init while building `state`; a `const`
 * defined 2,400 lines lower is in its temporal dead zone at that moment, and the
 * `catch` below would have swallowed the ReferenceError and returned `false`
 * every time — a preference that silently never loaded, which is worse than one
 * that never existed.
 */
const PICTURE_PREF_KEY = "ssf-charts-render-image";

/**
 * The remembered "insert as picture" default, for a NEW chart only.
 *
 * Read here rather than inside `stateFromConfig`, and that placement is the
 * whole point: `stateFromConfig` also runs when a chart is LOADED off a slide or
 * out of a JSON config, and seeding the preference there would silently rewrite
 * an existing shapes chart into a picture the moment its owner opened it.
 *
 * Storage can throw outright — a private window, or a browser set to block site
 * data — and the pane must open regardless. An unreadable preference is simply
 * the default, which is off.
 */
function rememberedPicturePref(): boolean {
  try {
    return localStorage.getItem(PICTURE_PREF_KEY) === "1";
  } catch {
    return false;
  }
}

const state: AppState = {
  ...stateFromConfig(sampleConfig("stacked")),
  renderImage: rememberedPicturePref(),
  editTarget: null,
};

function stateFromConfig(cfg: ChartConfig): Omit<AppState, "editTarget"> {
  // ONE ordering for the sheet and for the two positional side-channels below.
  // `dataToSheet` regroups a non-contiguous stack so the grid can express it at
  // all; deriving `seriesColors`/`seriesMeta` from the ORIGINAL order would put
  // them out of step with the rows the user sees, which is the same class of
  // bug the undo history had.
  const data = contiguousStacks(cfg.data);
  const sheet = dataToSheet(data);
  if (cfg.kind === "waterfall") {
    // Show "e" tokens in the sheet where totals are computed.
    // Bounded as well as coerced. The write had no bound on `i`, so an
    // out-of-range index wrote past the row's end and left `cells[1]` SPARSE —
    // `mountDatasheet`'s `row.forEach` skips holes, so the grid rendered one
    // stray cell under no column header while every other row was short. The
    // index is also dropped in silence, which is the honest outcome for an
    // index naming a category that is not there.
    const cols = sheet.cells[0]?.length ?? 0;
    for (const i of asArray<number>(cfg.waterfall?.totalIndices)) {
      if (!Number.isInteger(i) || i < 0 || i + 1 >= cols) continue;
      if (sheet.cells[1]) sheet.cells[1][i + 1] = "e";
    }
  }
  return {
    kind: cfg.kind,
    sheet,
    decorations: { ...cfg.decorations },
    horizontal: !!cfg.horizontal,
    title: cfg.title ?? "",
    // Same clamp as the engine — the state these fields hold is written back
    // into a config, so a NaN here becomes a NaN in the user's saved template.
    width: clampDim(cfg.width, DEFAULT_SIZE.width),
    height: clampDim(cfg.height, DEFAULT_SIZE.height),
    segmentOrder: cfg.segmentOrder ?? "sheet",
    scaleMin: cfg.scale?.min != null ? String(cfg.scale.min) : "",
    scaleMax: cfg.scale?.max != null ? String(cfg.scale.max) : "",
    breakFrom: cfg.axisBreak ? String(cfg.axisBreak.from) : "",
    breakTo: cfg.axisBreak ? String(cfg.axisBreak.to) : "",
    decimals: cfg.numberFormat?.decimals != null ? String(cfg.numberFormat.decimals) : "auto",
    suffix: cfg.numberFormat?.suffix ?? "",
    locale: cfg.numberFormat?.locale ?? "en-US",
    labelContent: asArray<string>(cfg.decorations?.labelContent).join(","),
    paletteName: paletteNameFor(cfg.style?.palette),
    style: cfg.style ? { ...cfg.style } : undefined,
    seriesColors: data.series.map((s) => s.color),
    seriesMeta: data.series.map((s) =>
      s.type || s.pattern || s.colors || s.scenario
        ? { type: s.type, pattern: s.pattern, colors: s.colors, scenario: s.scenario }
        : undefined,
    ),
    axisTitle: cfg.valueAxisTitle ?? "",
    logScale: !!cfg.logScale,
    renderImage: cfg.render === "image",
    footnote: cfg.footnote ?? "",
    pieExplode: asArray<number>(cfg.pie?.explode)
      .map((i) => i + 1)
      .join(","),
    extras: {
      boxplot: cfg.boxplot,
      heatmap: cfg.heatmap,
      map: cfg.map,
      combo: cfg.combo,
      gapWidth: cfg.gapWidth,
      overlap: cfg.overlap,
      multiples: cfg.multiples,
      scatter: cfg.scatter,
      radar: cfg.radar,
      gantt: cfg.gantt,
      butterfly: cfg.butterfly,
      tilemap: cfg.tilemap,
      otherBucket: cfg.otherBucket,
      pareto: cfg.pareto,
      categorySort: cfg.categorySort,
      secondaryAxis: cfg.secondaryAxis,
      labelOffsets: cfg.labelOffsets,
      pie: cfg.pie,
      waterfall: cfg.waterfall,
      numberFormat: cfg.numberFormat,
      labels: cfg.labels,
    },
  };
}

/**
 * Corporate style file: persisted defaults merged into every chart.
 *
 * The same shape the DECK carries, deliberately — the pane's exported style
 * file, the browser's copy and the deck's part are one format, so a user can
 * paste between them and a second spelling cannot drift from the first.
 */
type StyleFile = DeckStyle;
let styleFile: StyleFile = {};
try {
  styleFile = JSON.parse(localStorage.getItem("ssf-charts-style") ?? localStorage.getItem("powerchart-style") ?? "{}");
} catch {
  /* corrupted style file — start fresh */
}

/**
 * The style THIS DECK carries, read once on load from its custom XML part.
 *
 * `styleFile` above follows the browser; this follows the presentation. A deck
 * sent to a colleague used to arrive with no style at all, so their charts drew
 * with their defaults and every chart they added drifted further — the sharing
 * gap `src/core/deck-style.ts` exists to close.
 */
let deckStyle: DeckStyle | null = null;
/**
 * Which of the two is in force. The deck by default, and the pane says so.
 *
 * Importing a style in the pane is an explicit act, so it flips this to "mine"
 * for the session; "Use deck style" flips it back. The reasoning behind the
 * default is in `resolveStyleFile`, which is where the decision belongs.
 */
let stylePrefer: StylePreference = "deck";

/** Deck theme accents loaded via "Use deck theme" (session-scoped). */
let themePalette: string[] | null = null;

/** Style-file defaults + the palette preset chosen in the pane. */
function mergedStyle(): ChartConfig["style"] {
  // The loaded chart's own style beats the corporate defaults; an explicit
  // palette pick beats both.
  const style = {
    ...resolveStyleFile(deckStyle, styleFile, stylePrefer).style,
    ...state.style,
  } as NonNullable<ChartConfig["style"]>;
  if (state.paletteName === "Theme" && themePalette) style.palette = themePalette;
  else if (state.paletteName !== "Default") style.palette = namedPalette(state.paletteName);
  return Object.keys(style).length ? style : undefined;
}

/**
 * A palette by name, or the default.
 *
 * `PALETTES[name]` alone reaches Object.prototype — and the name is not always
 * one of the dropdown's options: a template or an imported style file supplies
 * it, and both are user JSON. `PALETTES["constructor"]` is a function, so the
 * `??` fallback does not fire, and every colour lookup then indexes a function
 * by number and gets `undefined`.
 */
function namedPalette(name: string): string[] {
  return Object.prototype.hasOwnProperty.call(PALETTES, name)
    ? PALETTES[name as keyof typeof PALETTES]
    : PALETTES.Default;
}

/**
 * Whatever arrived where an array belongs, as an array.
 *
 * `stateFromConfig` is the pane's ingest boundary and it trusted its own types.
 * The engine does not — `test/chart-hostile-input.test.ts` pins that
 * `{ palette: "red" }` renders fine everywhere — so a config that draws
 * correctly in the preview, in the deck and in the skill's .pptx could not be
 * OPENED in the pane that has to edit it. The JSON box reported the user's
 * perfectly valid JSON as "Invalid JSON: palette.join is not a function",
 * sending them to hunt a syntax error that is not there, and through the
 * template picker and "Edit selected chart" the same throw is uncaught and
 * silent — no message at all.
 *
 * The pane WRITES these configs, which is what makes it more than a hostile-
 * input case: `style-import` accepts a colleague's style file with no shape
 * check and persists it to localStorage, after which every export, every saved
 * template and every POWERCHART_CONFIG shape tag carries it.
 */
function asArray<T>(v: unknown): T[] {
  return Array.isArray(v) ? (v as T[]) : [];
}

/**
 * The preset name for a palette, or "Default" when it matches none — including
 * a chart's own custom palette, which `state.style` carries instead.
 */
function paletteNameFor(palette?: string[]): string {
  const list = asArray<string>(palette);
  if (!list.length) return "Default";
  return Object.entries(PALETTES).find(([, p]) => p.join() === list.join())?.[0] ?? "Default";
}

function currentConfig(): ChartConfig {
  const totals = new Set<number>();
  const data = sheetToData(state.sheet, state.kind === "waterfall" ? totals : undefined);
  // Size is state now, not a DOM read — the inputs write state on edit.
  const size = { width: state.width, height: state.height };
  const min = state.scaleMin.trim() === "" ? undefined : Number(state.scaleMin);
  const max = state.scaleMax.trim() === "" ? undefined : Number(state.scaleMax);
  const bFrom = Number(state.breakFrom);
  const bTo = Number(state.breakTo);
  const axisBreak =
    state.breakFrom.trim() && state.breakTo.trim() && Number.isFinite(bFrom) && Number.isFinite(bTo) && bTo > bFrom
      ? { from: bFrom, to: bTo }
      : undefined;
  data.series.forEach((s, i) => {
    const c = state.seriesColors[i];
    if (c) s.color = c;
    // Re-attach the fields the datasheet can't carry (combo type, pattern,
    // per-point colors) so an imported chart survives the sheet round-trip.
    const m = state.seriesMeta[i];
    if (m) {
      if (m.type) s.type = m.type;
      if (m.pattern) s.pattern = m.pattern;
      if (m.colors) s.colors = m.colors;
      if (m.scenario) s.scenario = m.scenario;
    }
  });
  const labelParts = state.labelContent
    ? (state.labelContent.split(",") as NonNullable<Decorations["labelContent"]>)
    : undefined;
  const explode = state.pieExplode
    .split(",")
    .map((v) => Number(v.trim()) - 1)
    .filter((v) => Number.isInteger(v) && v >= 0);
  // pie / waterfall / numberFormat are half pane-driven, half carried in extras:
  // the control owns explode / total "e" tokens / decimals+suffix+locale, but an
  // imported chart's semi, breakout, variableRadius, detailGroups, spacerIndices
  // and forceSign live only in extras and would be lost if the control overwrote
  // the whole object. Merge so the pane wins its own fields and the rest survives.
  const pieBase = { ...(state.extras.pie ?? {}) };
  delete pieBase.explode;
  const pie =
    explode.length || Object.keys(pieBase).length ? { ...pieBase, ...(explode.length ? { explode } : {}) } : undefined;
  const numberFormat: ChartConfig["numberFormat"] =
    state.decimals !== "auto" || state.suffix || state.locale !== "en-US" || state.extras.numberFormat
      ? {
          ...(state.extras.numberFormat ?? {}),
          decimals: state.decimals === "auto" ? "auto" : Number(state.decimals),
          suffix: state.suffix || undefined,
          locale: state.locale !== "en-US" ? state.locale : undefined,
        }
      : undefined;
  // The engine's one injected word ("Other"); pass a localized override only when
  // a non-English pane is active, so English configs stay byte-identical.
  const otherLabel = t("Other");
  const labels =
    otherLabel !== "Other" || state.extras.labels
      ? { ...state.extras.labels, ...(otherLabel !== "Other" ? { other: otherLabel } : {}) }
      : undefined;
  return {
    kind: state.kind,
    data,
    ...state.extras,
    labels,
    horizontal: state.horizontal || undefined,
    footnote: state.footnote || undefined,
    pie,
    valueAxisTitle: state.axisTitle || undefined,
    logScale: state.logScale || undefined,
    // Omitted rather than written as "shapes": the key is optional and the
    // engine treats undefined as shapes, so spelling out the default would add
    // noise to every exported config.
    render: state.renderImage ? "image" : undefined,
    style: mergedStyle(),
    ...size,
    title: state.title || undefined,
    decorations: { ...state.decorations, labelContent: labelParts },
    // Only the WATERFALL kind owns this field, because only its grid encodes the
    // totals: `sheetToData` fills `totals` from the sheet's "e" tokens, and those
    // are seeded only when `state.kind === "waterfall"`. Writing it
    // unconditionally therefore replaced the author's list with an EMPTY one for
    // every other kind — and `combo` with `combo.columns: "waterfall"` is a
    // documented recipe that reads it (`src/core/layout/column.ts` uses it for
    // the bridge's column bases). Slide 90 of the committed showcase is exactly
    // that chart, and opening it in the pane and re-inserting dropped its
    // closing total column and the label on it. It was the only one of the 123
    // showcase configs that did not round-trip through the pane.
    waterfall:
      state.kind === "waterfall"
        ? { ...(state.extras.waterfall ?? {}), totalIndices: [...totals] }
        : state.extras.waterfall,
    segmentOrder: state.segmentOrder === "sheet" ? undefined : state.segmentOrder,
    axisBreak,
    scale:
      (min != null && Number.isFinite(min)) || (max != null && Number.isFinite(max))
        ? { min: Number.isFinite(min!) ? min : undefined, max: Number.isFinite(max!) ? max : undefined }
        : undefined,
    numberFormat,
  };
}

// --- UI wiring ------------------------------------------------------------

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/**
 * Ask whether to put this chart on its own slide, and wait for the answer.
 *
 * Only reached when the estimate crosses `SLOW_INSERT_MS`. Resolves to
 * `"own-slide"` or `"here"`; there is no third answer, because the user pressed
 * Insert and something has to happen.
 *
 * A PROMISE ROUND A PAIR OF BUTTONS rather than `confirm()`. The pane runs in an
 * Office iframe where a native modal blocks the host as well as the pane, and
 * this is a speed hint — it does not deserve to freeze PowerPoint. Both handlers
 * are removed on the way out so a second slow insert cannot resolve the first
 * one's promise.
 */
function offerOwnSlide(
  present: number,
  /**
   * Whether to show the third button.
   *
   * FALSE means "not available", never "not worth it". Two reasons it is
   * withheld: a host that cannot rasterise has no picture to offer, and a chart
   * ALREADY going in as a picture has nothing to switch to — offering it there
   * would be a button that does nothing, on the one screen where the user is
   * being asked to make a decision under time pressure.
   */
  canPicture = false,
): Promise<"own-slide" | "here" | "picture"> {
  const box = $("slow-offer");
  const text = $("slow-offer-text");
  const own = $<HTMLButtonElement>("slow-offer-own");
  const here = $<HTMLButtonElement>("slow-offer-here");
  const picture = $<HTMLButtonElement>("slow-offer-picture");
  // The wording lives in `insert-cost.ts`, where a test can read what it
  // actually says rather than grep for what it must not. It USED to be fed two
  // durations from `estimateInsertMs` as well; those came out on 2026-09-05
  // because the curve behind them fails out of sample. See BACKLOG item 20.
  text.textContent = offerSentence(present);
  picture.hidden = !canPicture;
  box.hidden = false;
  return new Promise((resolve) => {
    const done = (choice: "own-slide" | "here" | "picture") => () => {
      own.removeEventListener("click", onOwn);
      here.removeEventListener("click", onHere);
      picture.removeEventListener("click", onPicture);
      box.hidden = true;
      resolve(choice);
    };
    const onOwn = done("own-slide");
    const onHere = done("here");
    const onPicture = done("picture");
    own.addEventListener("click", onOwn);
    here.addEventListener("click", onHere);
    picture.addEventListener("click", onPicture);
  });
}

/**
 * The question a too-dense chart never got asked.
 *
 * A chart over the web shape budget was rasterised BEFORE anything was put to
 * the user: `chartPicture` runs first, `pricedShapes` collapses to 1, and the
 * own-slide offer then prices a PICTURE — so it does not fire. Explode refuses
 * on the same host and the same predicate. The picture was not the fallback,
 * it was the only reachable outcome, and nobody was asked.
 *
 * There is a real alternative, and it is measured. The same chart built as a
 * one-slide .pptx and handed over in ONE call takes ~280ms, against 46.4s of
 * shape-by-shape drawing whose per-shape cost CLIMBS as the slide fills — 161ms
 * to 784ms across a single 103-shape chart. It keeps `POWERCHART_CONFIG`, so
 * the pane can still edit it, and it draws filled polygons and exact arcs that
 * Office.js cannot. Round 379 settled the one risk that would have killed it:
 * four inserts on a ONE-master deck, `mastersAfter: 1` every time, so a
 * generated slide does not drag its own master into someone's template.
 *
 * WHY THIS ASKS RATHER THAN JUST DOING IT. The user pointed at a slide, and
 * `insertSlidesFromBase64` can only add slides — it can never place shapes on
 * one that already exists. So whatever is chosen, the chart cannot land where
 * they pointed. Moving someone's content to a different slide unasked is a
 * worse surprise than a flat picture, so the choice is theirs.
 *
 * The picture stays the floor: it is what "here" means, what a refusal means,
 * and what every failure below falls back to. That also matches the field —
 * every web-capable competitor ships a picture on purpose, and every
 * native-shapes product is desktop-only.
 *
 * Reuses the slow-offer box. `here` is hidden because for this chart there is
 * no "here, as shapes" — that is the premise — and it is restored on the way
 * out so the other offer is not left with a missing button.
 */
function offerNativeInsteadOfPicture(shapes: number): Promise<"own-slide" | "picture"> {
  const box = $("slow-offer");
  const text = $("slow-offer-text");
  const own = $<HTMLButtonElement>("slow-offer-own");
  const here = $<HTMLButtonElement>("slow-offer-here");
  const picture = $<HTMLButtonElement>("slow-offer-picture");
  text.textContent =
    `That chart is ${shapes} shapes — too many for PowerPoint on the web to draw onto this slide. ` +
    `It can go on a slide of its own as editable shapes, or stay on this one as a picture.`;
  here.hidden = true;
  picture.hidden = false;
  box.hidden = false;
  return new Promise((resolve) => {
    const done = (choice: "own-slide" | "picture") => () => {
      own.removeEventListener("click", onOwn);
      picture.removeEventListener("click", onPicture);
      here.hidden = false;
      box.hidden = true;
      resolve(choice);
    };
    const onOwn = done("own-slide");
    const onPicture = done("picture");
    own.addEventListener("click", onOwn);
    picture.addEventListener("click", onPicture);
  });
}

const gallery = $("gallery");
const preview = $("preview");
const optionsHost = $("options");
const hostNote = $("host-note");

// eslint-disable-next-line prefer-const -- forward-declared; assigned once after wiring
let sheetApi: { setSheet(next: SheetModel): void };

/**
 * Pending auto-update push. Declared up here because the boot render calls
 * maybeAutoUpdate() long before the wiring below runs, and that now clears the
 * timer before its guard — a `let` further down would be in its dead zone.
 */
let autoUpdateTimer: ReturnType<typeof setTimeout> | undefined;

/**
 * An insert or update is talking to PowerPoint right now.
 *
 * `guard()` disables the buttons a CLICK comes from, but the auto-update timer
 * calls `doInsert` directly and never sees that. See `maybeAutoUpdate`.
 */
let insertInFlight = false;

/**
 * How many pane actions are talking to PowerPoint right now.
 *
 * Broader than `insertInFlight`, which only covers writes and exists to stop
 * two updates racing on the same chart. This one covers EVERY host action —
 * every `guard()`ed click plus the auto-update timer — because one background
 * listener needs to know not to add work of its own while the host is busy.
 * See `onSelectionChanged`.
 *
 * A counter rather than a flag: an action can nest (the demo run drives the
 * same helpers a click does), and a nested `finally` clearing a boolean would
 * declare the host idle while the outer action is still mid-flight.
 */
let hostWork = 0;

/** A selection change arrived while the host was busy and was not acted on. */
let selectionMissed = false;

const hostBusy = (): boolean => hostWork > 0 || insertInFlight;

function hostWorkStarted(): void {
  hostWork++;
}

function hostWorkFinished(): void {
  hostWork = Math.max(0, hostWork - 1);
  // Catch up on the selection changes that were dropped while this ran. The
  // banner is the only thing the watcher touches, so dropping events costs
  // nothing as long as the LAST one is honoured — and one read after the run
  // is far cheaper than the dozens the run itself provoked.
  if (!hostBusy() && selectionMissed) {
    selectionMissed = false;
    void refreshSelectionBanner();
  }
}

/**
 * Write the host note together with its status colour. The colour is a
 * parameter rather than an afterthought because only guard() used to set it,
 * so every other message inherited whatever the previous action left behind —
 * an "Invalid JSON" error rendered in the success green.
 */
const statusStrip = document.getElementById("status-strip");
const statusBar = document.getElementById("status-bar");
const statusElapsed = document.getElementById("status-elapsed");
const statusStop = document.getElementById("status-stop") as HTMLButtonElement | null;

/**
 * Show or hide the Stop button for the action in flight.
 *
 * Reset on every show: the button is reused across actions, and one left
 * disabled reading "Stopping…" would greet the next action as already stopping.
 */
function showStop(on: boolean) {
  if (!statusStop) return;
  if (on) {
    statusStop.disabled = false;
    statusStop.textContent = t("Stop");
  }
  statusStop.toggleAttribute("hidden", !on);
}

/**
 * How many times the pane has SETTLED — posted a note that is not "busy".
 *
 * `guard()` prints "Done." only when the action it ran did not report an end
 * state of its own, and it used to detect that by comparing the note text
 * against the busy text it had posted. Any progress note broke the comparison,
 * and an insert always ends on one: `phaseNote("done")` writes "Working… done",
 * which is still busy. So the text no longer matched, "Done." was skipped, and
 * the pane was left showing a blue busy note above a progress bar that slid on
 * forever — the action had finished, and nothing said so.
 *
 * Counting settlements asks the question that actually matters — "did this
 * action reach an end state?" — instead of inferring it from wording, which
 * also makes it immune to the language the note is rendered in.
 */
let settledNotes = 0;

function note(text: string, status: "ok" | "err" | "busy" | "none" = "none", params?: Record<string, string | number>) {
  // Route status text through the runtime translator so a localized pane
  // announces it in the user's language. `text` is the English source string
  // (an EN catalogue key), optionally carrying {placeholders} that `params`
  // fills after translation. The aria-live host-note reads the change to a
  // screen reader.
  hostNote.textContent = t(text, params);
  hostNote.className = status === "none" ? "hint" : `hint status-${status}`;
  // The strip carries the note now, so it has to follow it: shown whenever
  // there is something to say, collapsed when there is not.
  statusStrip?.toggleAttribute("hidden", !text);
  statusBar?.toggleAttribute("hidden", status !== "busy");
  if (status !== "busy") {
    setProgress(null);
    settledNotes++;
  }
}

/**
 * How far along, when we honestly know: a fraction for work we complete in
 * steps, "busy" for work we hand to PowerPoint in one go.
 *
 * A single insert is ONE context.sync() and Office.js reports nothing until it
 * lands, so any percentage there would be invented — and a bar that sticks at
 * 99% is a worse lie than no bar. Chunked work (the demo deck) really does know,
 * so it gets a real one.
 */
function setProgress(p: number | "busy" | null) {
  if (!statusBar) return;
  const fill = statusBar.querySelector("i");
  if (p === "busy") {
    statusBar.classList.add("indeterminate");
    if (fill) fill.style.width = "";
  } else {
    statusBar.classList.remove("indeterminate");
    if (fill) fill.style.width = p == null ? "0" : `${Math.round(Math.max(0, Math.min(1, p)) * 100)}%`;
  }
}

/**
 * The last moment anything was heard from a run — see `noteHostActivity`.
 *
 * Module-level rather than passed in, because the two things that need it sit
 * in different parts of the file: the trace subscriber, which sees every step a
 * run produces, and the elapsed ticker, which is the only thing on screen while
 * a sync is outstanding.
 */
let lastHostActivity = 0;

/** A run just did something. Bumped per traced step. */
export function noteHostActivity(now = Date.now()): void {
  lastHostActivity = now;
}

/**
 * How long a run may go silent before the pane stops implying it is working.
 *
 * Generous on purpose. A single draw batch on PowerPoint web has been measured
 * at ~17 seconds and a stalled-but-alive sync at 45, so anything under a minute
 * would cry wolf on a host that is merely slow — the failure mode this replaces
 * cuts the other way and is worse, but a warning nobody believes is no warning.
 */
const SILENT_RUN_MS = 60_000;

/**
 * What the elapsed readout should say, given how long the run has been going
 * and how long since anything happened.
 *
 * Pure, and exported, because it is the whole of a decision this project has
 * now watched go wrong on a real host three times: PowerPoint dies, the task
 * pane survives — it is a separate frame, and the `PowerPoint.run` promise it
 * is waiting on simply never settles — and the pane counts upward forever under
 * the word "Working…". The owner is left watching a number climb with no way to
 * tell it from a slow chart. There is no error to catch here and no timeout
 * that helps, because nothing rejects; silence is the only evidence there is.
 */
export function elapsedLabel(elapsedMs: number, silentMs: number): string {
  const secs = `${Math.round(elapsedMs / 1000)}s`;
  return silentMs >= SILENT_RUN_MS ? `${secs} · silent for ${Math.round(silentMs / 1000)}s` : secs;
}

/**
 * Count the seconds while the host works. It is the only number we can report
 * mid-sync, and on a host that takes 20s to draw a chart, a number that moves
 * is the difference between "working" and "dead".
 *
 * Except when it is not, which is why the readout also carries silence. See
 * `elapsedLabel`.
 */
let elapsedTimer: ReturnType<typeof setInterval> | undefined;
let saidSilent = false;
function startElapsed() {
  const t0 = Date.now();
  stopElapsed();
  noteHostActivity(t0);
  saidSilent = false;
  const tick = () => {
    const now = Date.now();
    const silentMs = now - lastHostActivity;
    if (statusElapsed) statusElapsed.textContent = elapsedLabel(now - t0, silentMs);
    // Said once, not every second: the note area is shared, and a message that
    // rewrites itself every tick is one nobody reads. It also must not be an
    // "err" — nothing has failed as far as this pane knows, and claiming
    // otherwise over a merely slow host is the mistake in the other direction.
    if (!saidSilent && silentMs >= SILENT_RUN_MS) {
      saidSilent = true;
      note(
        "PowerPoint has not answered for {secs}s. Look at the slide area: if PowerPoint is showing *Sorry, we ran into a problem*, click Refresh there — nothing behind that dialog can answer. The run's steps are saved either way and *Download the crashed run* will offer them.",
        "busy",
        { secs: Math.round(silentMs / 1000) },
      );
    }
  };
  tick();
  elapsedTimer = setInterval(tick, 1000);
}
function stopElapsed() {
  clearInterval(elapsedTimer);
  elapsedTimer = undefined;
  saidSilent = false;
  if (statusElapsed) statusElapsed.textContent = "";
}

function applyConfig(cfg: ChartConfig, editTarget: EditTarget | null) {
  Object.assign(state, stateFromConfig(cfg), { editTarget });
  sheetApi.setSheet(state.sheet);
  const titleField = document.getElementById("chart-title") as HTMLInputElement | null;
  if (titleField) titleField.value = state.title;
  // currentConfig() reads the size straight off these fields, so leaving them
  // stale silently resized every loaded chart to the previous one's dimensions.
  const sizeField = (id: string, value: number) => {
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (el) el.value = String(value);
  };
  sizeField("chart-w", clampDim(cfg.width, DEFAULT_SIZE.width));
  sizeField("chart-h", clampDim(cfg.height, DEFAULT_SIZE.height));
  const renderBox = document.getElementById("render-image") as HTMLInputElement | null;
  if (renderBox) renderBox.checked = state.renderImage;
  resetHistory();
  renderGallery();
  renderOptions();
  renderPreview();
  renderActionState();
}

/** Miniature preview of a chart kind for the gallery (think-cell's Elements menu). */
function thumbnailSvg(kind: ChartKind): string {
  const cfg: ChartConfig = {
    ...sampleConfig(kind),
    width: 96,
    height: 58,
    title: undefined,
    decorations: {
      segmentLabels: false,
      seriesLabels: false,
      totals: false,
      categoryAxis: false,
      valueAxis: false,
      gridlines: false,
    },
    style: { fontSize: 4 },
  };
  try {
    return sceneToSvg(buildChart(cfg));
  } catch {
    return "";
  }
}

const thumbnails = new Map<ChartKind, string>();

/** Chart kinds grouped by family, so the picker is scannable (think-cell's
 *  Elements menu). Any CHART_KINDS entry not listed here still renders under a
 *  trailing "Other" group, so a new kind can never silently disappear. */
const CHART_GROUPS: { label: string; kinds: ChartKind[] }[] = [
  {
    label: "Columns & bars",
    kinds: ["stacked", "clustered", "stacked100", "waterfall", "mekko", "butterfly", "cascade", "funnel"],
  },
  { label: "Line & area", kinds: ["line", "area", "combo"] },
  { label: "Parts of a whole", kinds: ["pie", "doughnut", "treemap", "sunburst", "waffle"] },
  { label: "Distribution", kinds: ["boxplot", "violin", "candlestick"] },
  { label: "Correlation", kinds: ["scatter", "bubble"] },
  { label: "Matrix & spatial", kinds: ["heatmap", "tilemap", "radar", "gantt"] },
];

function thumbButton(kind: ChartKind, label: string): HTMLButtonElement {
  const b = document.createElement("button");
  const chosen = kind === state.kind;
  b.className = "thumb" + (chosen ? " active" : "");
  /**
   * WHICH ONE IS CHOSEN, said out loud. The selected chart type was carried by
   * a CSS class alone, so a screen-reader user tabbing the gallery heard
   * "Waterfall, button" twenty-five times with nothing to say which was the
   * current chart. Sighted users were fine — `.thumb.active` changes border,
   * background AND font-weight, so it was never colour-alone — which is exactly
   * why it went unnoticed.
   *
   * `aria-pressed` rather than a radiogroup: these are buttons that apply a
   * config, and `role="radio"` would promise arrow-key navigation between them
   * that this gallery does not implement.
   */
  b.setAttribute("aria-pressed", String(chosen));
  b.dataset.kind = kind;
  b.dataset.label = label.toLowerCase();
  if (!thumbnails.has(kind)) thumbnails.set(kind, thumbnailSvg(kind));
  const pic = document.createElement("span");
  pic.className = "thumb-pic";
  pic.innerHTML = thumbnails.get(kind)!;
  const cap = document.createElement("span");
  cap.className = "thumb-cap";
  cap.textContent = label;
  b.append(pic, cap);
  b.addEventListener("click", () => {
    applyConfig(sampleConfig(kind), null);
    // Auto-collapse the (tall) type grid once a kind is chosen — the summary
    // then shows the current kind, click to re-expand.
    const acc = document.getElementById("type-acc") as HTMLDetailsElement | null;
    if (acc) acc.open = false;
  });
  return b;
}

function renderGallery() {
  gallery.innerHTML = "";
  const labelOf = new Map(CHART_KINDS.map((k) => [k.kind, k.label] as const));
  const grouped = new Set<ChartKind>();
  const groups = CHART_GROUPS.map((g) => ({ label: g.label, kinds: g.kinds.filter((k) => labelOf.has(k)) }));
  groups.forEach((g) => g.kinds.forEach((k) => grouped.add(k)));
  const leftover = CHART_KINDS.filter((k) => !grouped.has(k.kind)).map((k) => k.kind);
  if (leftover.length) groups.push({ label: "Other", kinds: leftover });

  for (const g of groups) {
    if (!g.kinds.length) continue;
    const sec = document.createElement("div");
    sec.className = "type-group";
    /**
     * A REAL HEADING, because this is how a screen-reader user moves around.
     * The six chart families — "Columns & bars", "Line & area" and the rest —
     * were `<div class="group-label">`: they looked like headings, read like
     * headings, and were navigable as none, so the gallery was one flat run of
     * twenty-five buttons with no way to jump between families.
     *
     * `h3` is the level this pane already uses for a sub-heading inside a
     * section (`.run-label` in taskpane.html), and `.group-label` sets font-size,
     * weight and margin explicitly, so nothing moves on screen.
     */
    const heading = document.createElement("h3");
    heading.className = "group-label";
    heading.textContent = g.label;
    const grid = document.createElement("div");
    grid.className = "gallery";
    for (const kind of g.kinds) grid.appendChild(thumbButton(kind, labelOf.get(kind)!));
    sec.append(heading, grid);
    gallery.appendChild(sec);
  }
  localizeTree(gallery);
  updateTypeSummary();
  applyTypeFilter();
}

/** Filter the grouped picker by the search box; hide families with no match. */
function applyTypeFilter() {
  const input = document.getElementById("type-search-input") as HTMLInputElement | null;
  const q = (input?.value ?? "").trim().toLowerCase();
  let anyVisible = false;
  for (const sec of gallery.querySelectorAll<HTMLElement>(".type-group")) {
    let shown = 0;
    for (const btn of sec.querySelectorAll<HTMLButtonElement>(".thumb")) {
      const match = !q || (btn.dataset.label ?? "").includes(q);
      btn.style.display = match ? "" : "none";
      if (match) shown++;
    }
    sec.style.display = shown ? "" : "none";
    if (shown) anyVisible = true;
  }
  const noRes = document.getElementById("type-noresult");
  if (noRes) noRes.style.display = anyVisible ? "none" : "";
}

/** Reflect the selected chart kind in the collapsed "1 · Chart type" summary. */
function updateTypeSummary() {
  const sub = document.getElementById("type-sub");
  if (sub) sub.textContent = CHART_KINDS.find((k) => k.kind === state.kind)?.label ?? state.kind;
}

/** Top-level mode tabs: Chart / Elements / Agenda / Automation. */
function wireTabs() {
  const tabs = Array.from(document.querySelectorAll<HTMLButtonElement>(".tabs .tab"));
  const panels = Array.from(document.querySelectorAll<HTMLElement>(".tab-panel"));
  const bar = document.querySelector<HTMLElement>(".action-bar");
  /**
   * The action bar belongs to the Chart tab alone.
   *
   * Every one of its actions reads the CHART's state — "Insert into slide"
   * inserts currentConfig(), and the ⋯ menu edits, rescales or downloads
   * charts. On Elements it therefore offered a big primary button that inserts
   * a stacked column chart, sitting directly under the small "Insert" that
   * inserts the Harvey ball you are actually looking at: the prominent button
   * was the wrong one. Elements and Agenda already carry their own insert
   * buttons, so hiding it there removes a trap rather than a feature.
   */
  const showBarFor = (name?: string) => bar?.toggleAttribute("hidden", name !== "chart");
  // Select a tab and mirror it into the ARIA tab pattern: one tab is
  // aria-selected and in the tab order (tabindex 0), the rest are out of it
  // (tabindex -1) and reachable only via the arrow keys. Without this a screen
  // reader can't tell which mode is active and keyboard users tab through every
  // one. `focus` is false for the initial paint (don't steal focus on boot).
  const select = (tab: HTMLButtonElement, focus: boolean) => {
    const name = tab.dataset.tab;
    tabs.forEach((t) => {
      const on = t === tab;
      t.classList.toggle("active", on);
      t.setAttribute("aria-selected", String(on));
      t.tabIndex = on ? 0 : -1;
    });
    panels.forEach((p) => p.classList.toggle("active", p.dataset.panel === name));
    showBarFor(name);
    if (focus) tab.focus();
  };
  tabs.forEach((tab, i) => {
    tab.addEventListener("click", () => select(tab, false));
    tab.addEventListener("keydown", (e) => {
      // Left/Right move between tabs (wrapping); Home/End jump to the ends —
      // the WAI-ARIA tabs keyboard model.
      const delta = e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0;
      let next = -1;
      if (delta) next = (i + delta + tabs.length) % tabs.length;
      else if (e.key === "Home") next = 0;
      else if (e.key === "End") next = tabs.length - 1;
      if (next >= 0) {
        e.preventDefault();
        select(tabs[next], true);
      }
    });
  });
  // Seed the ARIA state to match the markup's default active tab, without
  // moving focus. A ?tab= deep link CLICKS its tab afterwards, re-running select.
  const initial = tabs.find((t) => t.classList.contains("active")) ?? tabs[0];
  if (initial) select(initial, false);
}

/** The footer "⋯" overflow menu holding the secondary actions (edit selected,
 *  same scale, download). Opens upward; closes on item click, outside click,
 *  or Escape. */
function wireActionsMenu() {
  const btn = document.getElementById("more-actions");
  const menu = document.getElementById("actions-menu");
  if (!btn || !menu) return;
  // Give the popup the ARIA menu role so a screen reader announces it as a menu
  // and its entries as menu items, reachable with the arrow keys.
  btn.setAttribute("aria-haspopup", "menu");
  menu.setAttribute("role", "menu");
  const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button"));
  items.forEach((it) => {
    it.setAttribute("role", "menuitem");
    it.tabIndex = -1;
  });
  const focusItem = (i: number) => items[(i + items.length) % items.length]?.focus();
  const setOpen = (open: boolean, restoreFocus = false) => {
    menu.hidden = !open;
    btn.setAttribute("aria-expanded", String(open));
    if (open)
      focusItem(0); // move into the menu so the keyboard lands there
    else if (restoreFocus) btn.focus(); // Escape/close returns focus to the trigger
  };
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    setOpen(menu.hidden);
  });
  menu.addEventListener("click", () => setOpen(false, true));
  menu.addEventListener("keydown", (e) => {
    const i = items.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === "ArrowDown") {
      e.preventDefault();
      focusItem(i + 1);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      focusItem(i - 1);
    } else if (e.key === "Home") {
      e.preventDefault();
      focusItem(0);
    } else if (e.key === "End") {
      e.preventDefault();
      focusItem(items.length - 1);
    }
  });
  document.addEventListener("click", (e) => {
    if (!menu.hidden && e.target !== btn && !menu.contains(e.target as Node)) setOpen(false);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !menu.hidden) setOpen(false, true);
  });
}

/** One collapsible Format group (Labels / Axes / Analysis / Layout / Colours). */
interface OptGroup {
  details: HTMLDetailsElement;
  togs: HTMLDivElement;
  body: HTMLDivElement;
}
const FGROUP_ICON: Record<string, string> = {
  labels: '<path d="M3 4h10M8 4v9" stroke-linecap="round"/>',
  axes: '<path d="M4 3v10h9M4 10l3-3 2 2 4-4" stroke-linecap="round" stroke-linejoin="round"/>',
  analysis: '<path d="M3 12l4-4 2 2 5-6M11 4h3v3" stroke-linecap="round" stroke-linejoin="round"/>',
  layout: '<rect x="3" y="3" width="10" height="10" rx="1.5"/><path d="M3 8h10M8 3v10" stroke-width="1.1"/>',
  colours: '<circle cx="8" cy="8" r="5"/><path d="M8 3a5 5 0 010 10z" fill="currentColor" stroke="none"/>',
};
function optGroup(name: string, iconKey: string): OptGroup {
  const details = document.createElement("details");
  details.className = "fgroup";
  const summary = document.createElement("summary");
  summary.innerHTML =
    `<svg class="fgroup-ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true">${FGROUP_ICON[iconKey]}</svg>` +
    `<span class="fgroup-name">${name}</span><span class="fgroup-count"></span><span class="fgroup-chev"></span>`;
  const body = document.createElement("div");
  body.className = "fgroup-body";
  const togs = document.createElement("div");
  togs.className = "togs";
  body.appendChild(togs);
  details.append(summary, body);
  return { details, togs, body };
}
/** Reflect each group's enabled-checkbox count in its "N on" pill. */
function updateGroupCounts() {
  for (const g of optionsHost.querySelectorAll<HTMLDetailsElement>(".fgroup")) {
    const on = g.querySelectorAll<HTMLInputElement>('input[type="checkbox"]:checked').length;
    const pill = g.querySelector(".fgroup-count");
    if (pill) {
      pill.textContent = on ? `${on} on` : "";
      pill.classList.toggle("zero", on === 0);
    }
  }
}

function renderOptions() {
  optionsHost.innerHTML = "";
  const d = state.decorations;
  const nCats = () => Math.max(0, state.sheet.cells[0].length - 1);
  // think-cell surfaces these controls contextually on the chart; in the pane
  // they're grouped so the long list stays scannable.
  const G = {
    labels: optGroup("Labels", "labels"),
    axes: optGroup("Axes & scale", "axes"),
    analysis: optGroup("Analysis", "analysis"),
    layout: optGroup("Layout", "layout"),
    colours: optGroup("Colours & style", "colours"),
  };
  G.labels.details.open = true;

  const toggles: { key: keyof Decorations; label: string; group: OptGroup }[] = [
    { key: "segmentLabels", label: "Segment labels", group: G.labels },
    { key: "seriesLabels", label: "Series labels", group: G.labels },
    { key: "totals", label: "Column totals", group: G.labels },
    { key: "grandTotal", label: "Grand total", group: G.labels },
    { key: "categoryAxis", label: "Category labels", group: G.labels },
    { key: "valueAxis", label: "Value axis", group: G.axes },
    { key: "gridlines", label: "Gridlines", group: G.axes },
    { key: "connectors", label: "Connector lines", group: G.layout },
    { key: "hundredPercentNote", label: "100% = note", group: G.labels },
  ];
  for (const t of toggles) {
    const label = document.createElement("label");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!d[t.key];
    cb.addEventListener("change", () => {
      (d as Record<string, unknown>)[t.key] = cb.checked;
      renderPreview();
    });
    label.append(cb, t.label);
    t.group.togs.appendChild(label);
  }

  // Tufte-style datamark axis: tick dashes + labels, no axis line.
  const dm = document.createElement("label");
  const dmCb = document.createElement("input");
  dmCb.type = "checkbox";
  dmCb.checked = d.valueAxis === "datamarks";
  dmCb.addEventListener("change", () => {
    d.valueAxis = dmCb.checked ? "datamarks" : false;
    renderOptions();
    renderPreview();
  });
  dm.append(dmCb, "Datamark axis (ticks only)");
  G.axes.togs.appendChild(dm);

  // think-cell's rotation handle, as a toggle: column ⇄ bar.
  const rot = document.createElement("label");
  const rotCb = document.createElement("input");
  rotCb.type = "checkbox";
  rotCb.checked = state.horizontal;
  rotCb.addEventListener("change", () => {
    state.horizontal = rotCb.checked;
    renderPreview();
  });
  rot.append(rotCb, "Horizontal (bar)");
  G.layout.togs.appendChild(rot);

  // pairControl only knows from/to, but decorations.cagr also carries the series
  // the arrow is anchored to — and core/decor.ts reads it. Emitting the bare pair
  // dropped that anchor the moment the spinner was touched, so the arrow silently
  // switched to the column totals and printed a different growth rate.
  const cagrSeries = d.cagr?.series;
  G.analysis.body.appendChild(
    pairControl("CAGR arrow", d.cagr, nCats(), (pair) => {
      d.cagr = pair ? { ...pair, series: cagrSeries } : undefined;
      renderPreview();
    }),
  );

  // Difference arrow: totals by default, or a level arrow at a series.
  const diff = document.createElement("label");
  diff.className = "wide";
  const diffCb = document.createElement("input");
  diffCb.type = "checkbox";
  diffCb.checked = !!d.difference;
  const dFrom = numInput((d.difference?.from ?? 0) + 1);
  const dTo = numInput((d.difference?.to ?? Math.max(0, nCats() - 1)) + 1);
  const dSeries = numInput((d.difference?.series ?? -1) + 1, 0);
  dSeries.title = "0 = column totals, 1+ = level of that series";
  const emitDiff = () => {
    const s = Number(dSeries.value) - 1;
    d.difference = diffCb.checked
      ? { from: Number(dFrom.value) - 1, to: Number(dTo.value) - 1, series: s >= 0 ? s : undefined }
      : undefined;
    renderPreview();
  };
  [diffCb, dFrom, dTo, dSeries].forEach((el) => el.addEventListener(el === diffCb ? "change" : "input", emitDiff));
  dFrom.setAttribute("aria-label", "Difference arrow from category");
  dTo.setAttribute("aria-label", "Difference arrow to category");
  dSeries.setAttribute("aria-label", "Difference arrow series");
  diff.append(diffCb, "Difference arrow from ", dFrom, " to ", dTo, " series ", dSeries);
  G.analysis.body.appendChild(diff);

  // Value lines: mean and/or comma-separated fixed values.
  const vl = document.createElement("label");
  vl.className = "wide";
  const vlMean = document.createElement("input");
  vlMean.type = "checkbox";
  const existing = d.valueLines ?? (d.valueLine ? [d.valueLine] : []);
  vlMean.checked = existing.some((v) => v.mode === "mean");
  const vlValues = document.createElement("input");
  vlValues.type = "text";
  vlValues.placeholder = "e.g. 50, 100";
  vlValues.style.width = "80px";
  vlValues.value = existing
    .filter((v): v is { mode: "value"; value: number } => v.mode === "value")
    .map((v) => v.value)
    .join(", ");
  const emitVl = () => {
    const lines: NonNullable<Decorations["valueLines"]> = [];
    if (vlMean.checked) lines.push({ mode: "mean" });
    for (const part of vlValues.value.split(",")) {
      const v = Number(part.trim());
      if (part.trim() && Number.isFinite(v)) lines.push({ mode: "value", value: v });
    }
    d.valueLines = lines.length ? lines : undefined;
    d.valueLine = undefined;
    renderPreview();
  };
  vlMean.addEventListener("change", emitVl);
  vlValues.addEventListener("input", emitVl);
  vlValues.setAttribute("aria-label", "Value line at fixed values");
  vl.append(vlMean, "Value line: mean Ø", " + values ", vlValues);
  G.analysis.body.appendChild(vl);

  // Segment order (think-cell's mini-toolbar menu).
  const so = document.createElement("label");
  so.className = "wide";
  const soSel = document.createElement("select");
  for (const [value, label] of [
    ["sheet", "Sheet order"],
    ["reverse", "Reversed"],
    ["ascending", "Ascending"],
    ["descending", "Descending"],
  ]) {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    soSel.appendChild(opt);
  }
  soSel.value = state.segmentOrder;
  soSel.addEventListener("change", () => {
    state.segmentOrder = soSel.value as AppState["segmentOrder"];
    renderPreview();
  });
  so.append("Segment order ", soSel);
  G.layout.body.appendChild(so);

  // Manual axis scale (think-cell's axis-handle dragging).
  const sc = document.createElement("label");
  sc.className = "wide";
  const scMin = document.createElement("input");
  const scMax = document.createElement("input");
  for (const [el, val] of [
    [scMin, state.scaleMin],
    [scMax, state.scaleMax],
  ] as const) {
    el.type = "text";
    el.style.width = "48px";
    el.placeholder = "auto";
    el.value = val;
  }
  const emitScale = () => {
    state.scaleMin = scMin.value;
    state.scaleMax = scMax.value;
    renderPreview();
  };
  scMin.addEventListener("input", emitScale);
  scMax.addEventListener("input", emitScale);
  scMax.setAttribute("aria-label", "Axis scale maximum");
  sc.append("Axis scale min ", scMin, " max ", scMax);
  G.axes.body.appendChild(sc);

  // Axis break (compresses the given value range).
  const ab = document.createElement("label");
  ab.className = "wide";
  const abFrom = document.createElement("input");
  const abTo = document.createElement("input");
  for (const [el, val] of [
    [abFrom, state.breakFrom],
    [abTo, state.breakTo],
  ] as const) {
    el.type = "text";
    el.style.width = "48px";
    el.placeholder = "none";
    el.value = val;
  }
  const emitBreak = () => {
    state.breakFrom = abFrom.value;
    state.breakTo = abTo.value;
    renderPreview();
  };
  abFrom.addEventListener("input", emitBreak);
  abTo.addEventListener("input", emitBreak);
  abTo.setAttribute("aria-label", "Axis break to value");
  ab.append("Axis break from ", abFrom, " to ", abTo);
  G.axes.body.appendChild(ab);

  // Number format.
  const nf = document.createElement("label");
  nf.className = "wide";
  const nfDec = document.createElement("select");
  const decOptions = ["auto", "0", "1", "2"];
  // An imported chart may carry decimals outside the four presets (3, say).
  // Without an option of its own the select falls to value "", and emitNf — which
  // writes BOTH controls — then rewrote decimals to 0 as soon as the suffix box
  // was touched. Carry the loaded value as its own option instead.
  if (!decOptions.includes(state.decimals)) decOptions.push(state.decimals);
  for (const v of decOptions) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v === "auto" ? "auto" : `${v} dp`;
    nfDec.appendChild(opt);
  }
  nfDec.value = state.decimals;
  const nfSuffix = document.createElement("input");
  nfSuffix.type = "text";
  nfSuffix.style.width = "48px";
  nfSuffix.placeholder = "e.g. €m";
  nfSuffix.value = state.suffix;
  const emitNf = () => {
    state.decimals = nfDec.value;
    state.suffix = nfSuffix.value;
    renderPreview();
  };
  nfDec.addEventListener("change", emitNf);
  nfSuffix.addEventListener("input", emitNf);
  const nfLoc = document.createElement("select");
  for (const v of ["en-US", "de-DE", "fr-FR", "da-DK"]) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    nfLoc.appendChild(opt);
  }
  nfLoc.value = state.locale;
  nfLoc.addEventListener("change", () => {
    state.locale = nfLoc.value;
    renderPreview();
  });
  nfSuffix.setAttribute("aria-label", "Label suffix");
  nfLoc.setAttribute("aria-label", "Label locale");
  nf.append("Labels: decimals ", nfDec, " suffix ", nfSuffix, " locale ", nfLoc);
  G.labels.body.appendChild(nf);

  // Footnote / source line — good charts always cite their source.
  const fn = document.createElement("label");
  fn.className = "wide";
  const fnInput = document.createElement("input");
  fnInput.type = "text";
  fnInput.placeholder = "e.g. Source: Statistics Denmark, 2024";
  fnInput.style.width = "180px";
  fnInput.value = state.footnote;
  fnInput.addEventListener("input", () => {
    state.footnote = fnInput.value;
    renderPreview();
  });
  fn.append("Footnote / source ", fnInput);
  G.colours.body.appendChild(fn);

  // Exploding slices (pie/doughnut only).
  if (state.kind === "pie" || state.kind === "doughnut") {
    const ex = document.createElement("label");
    ex.className = "wide";
    const exInput = document.createElement("input");
    exInput.type = "text";
    exInput.placeholder = "e.g. 1";
    exInput.style.width = "48px";
    exInput.value = state.pieExplode;
    exInput.addEventListener("input", () => {
      state.pieExplode = exInput.value;
      renderPreview();
    });
    ex.append("Explode slices ", exInput);
    G.layout.body.appendChild(ex);
  }

  // Label content (think-cell's label dropdown).
  const lc = document.createElement("label");
  lc.className = "wide";
  const lcSel = document.createElement("select");
  for (const [value, label] of [
    ["", "Default"],
    ["value", "Value"],
    ["percent", "%"],
    ["value,percent", "Value + %"],
    ["series,value", "Series + value"],
    ["category,percent", "Category + %"],
  ]) {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    lcSel.appendChild(opt);
  }
  lcSel.value = state.labelContent;
  lcSel.addEventListener("change", () => {
    state.labelContent = lcSel.value;
    renderPreview();
  });
  lc.append("Label content ", lcSel);
  G.labels.body.appendChild(lc);

  // Axis title + log scale.
  const ax = document.createElement("label");
  ax.className = "wide";
  const axTitle = document.createElement("input");
  axTitle.type = "text";
  axTitle.style.width = "56px";
  axTitle.placeholder = "e.g. €m";
  axTitle.value = state.axisTitle;
  axTitle.addEventListener("input", () => {
    state.axisTitle = axTitle.value;
    renderPreview();
  });
  const axLog = document.createElement("input");
  axLog.type = "checkbox";
  axLog.checked = state.logScale;
  axLog.addEventListener("change", () => {
    state.logScale = axLog.checked;
    renderPreview();
  });
  axLog.setAttribute("aria-label", "Logarithmic scale");
  ax.append("Axis title ", axTitle, " ", axLog, " log scale");
  G.axes.body.appendChild(ax);

  // Palette preset + per-series color overrides.
  const pal = document.createElement("label");
  pal.className = "wide";
  const palSel = document.createElement("select");
  const palNames = [...Object.keys(PALETTES), ...(themePalette ? ["Theme"] : [])];
  for (const name of palNames) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    palSel.appendChild(opt);
  }
  palSel.value = state.paletteName;
  palSel.addEventListener("change", () => {
    state.paletteName = palSel.value;
    // An explicit pick — "Default" included — replaces a custom palette the
    // loaded chart brought with it.
    delete state.style?.palette;
    renderPreview();
  });
  // Read the deck's theme accent colors (PowerPointApi 1.10) as a palette.
  const themeBtn = document.createElement("button");
  themeBtn.type = "button";
  // THE SAME TRAP AS THE TWO DECK-STYLE BUTTONS, and found by sweeping for it:
  // `renderOptions()` is called at module scope, so this gate is first evaluated
  // before `Office.context` exists and answers false inside PowerPoint. It is
  // less visible than those two only because an incidental re-render can undo
  // it — which makes it worse to diagnose, not better. An id puts it under
  // `syncHostOnlyButtons` with its siblings.
  themeBtn.id = "style-deck-theme";
  themeBtn.textContent = "Use deck theme";
  themeBtn.disabled = !isPowerPointHost();
  themeBtn.addEventListener("click", async () => {
    const loaded = await loadThemePalette();
    if (!loaded) {
      themeBtn.textContent = "Theme unavailable";
      return;
    }
    themePalette = loaded;
    if (![...palSel.options].some((o) => o.value === "Theme")) {
      const opt = document.createElement("option");
      opt.value = "Theme";
      opt.textContent = "Theme";
      palSel.appendChild(opt);
    }
    state.paletteName = "Theme";
    palSel.value = "Theme";
    renderOptions();
    renderPreview();
  });
  pal.append("Palette ", palSel, " ", themeBtn);
  G.colours.body.appendChild(pal);

  const colors = document.createElement("div");
  colors.className = "wide series-colors";
  // Resolve through mergedStyle so the swatches show the colors the chart
  // actually draws with — including a loaded chart's custom palette.
  const palette = mergedStyle()?.palette ?? PALETTES.Default;
  currentSeriesNames().forEach((name, i) => {
    const wrap = document.createElement("label");
    const input = document.createElement("input");
    input.type = "color";
    input.value = state.seriesColors[i] ?? palette[i % palette.length];
    input.addEventListener("input", () => {
      state.seriesColors[i] = input.value;
      renderPreview();
    });
    wrap.append(input, name);
    colors.appendChild(wrap);
  });
  G.colours.body.appendChild(colors);

  for (const g of [G.labels, G.axes, G.analysis, G.layout, G.colours]) optionsHost.appendChild(g.details);
  localizeTree(optionsHost);
  updateGroupCounts();
}

/** Series names from the sheet, excluding special rows. */
function currentSeriesNames(): string[] {
  return sheetToData(state.sheet).series.map((s) => s.name);
}

function numInput(value: number, min = 1): HTMLInputElement {
  const el = document.createElement("input");
  el.type = "number";
  el.min = String(min);
  el.value = String(value);
  return el;
}

function pairControl(
  label: string,
  current: { from: number; to: number } | undefined,
  nCats: number,
  onChange: (pair: { from: number; to: number } | undefined) => void,
): HTMLElement {
  const wrap = document.createElement("label");
  wrap.className = "wide";
  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = !!current;
  const from = numInput((current?.from ?? 0) + 1);
  const to = numInput((current?.to ?? Math.max(0, nCats - 1)) + 1);
  const emit = () => onChange(cb.checked ? { from: Number(from.value) - 1, to: Number(to.value) - 1 } : undefined);
  cb.addEventListener("change", emit);
  from.addEventListener("input", emit);
  to.addEventListener("input", emit);
  from.setAttribute("aria-label", "From category");
  to.setAttribute("aria-label", "To category");
  wrap.append(cb, label, " from ", from, " to ", to);
  return wrap;
}

/**
 * The canvas colour an SVG is painted on. Forcing white here put a dark-theme
 * config's light text on a white rect (1.13:1 contrast) in the preview AND in
 * both downloads, while insertSceneIntoSlide drops the same shapes onto the real
 * (dark) slide with no background rect at all — canvas != export. Default charts
 * are byte-identical: style.background already defaults to #ffffff.
 */
const canvasColor = (cfg: ChartConfig) => cfg.style?.background ?? "#ffffff";

function renderPreviewNow() {
  try {
    const cfg = currentConfig();
    const scene = buildChart(cfg);
    preview.innerHTML = sceneToSvg(scene, { background: canvasColor(cfg) });
  } catch (err) {
    // textContent, not innerHTML: the error message can carry config-derived
    // strings, and an innerHTML sink here would be a second injection path.
    preview.replaceChildren();
    const p = document.createElement("p");
    p.className = "hint";
    p.textContent = `Could not render: ${err instanceof Error ? err.message : String(err)}`;
    preview.appendChild(p);
  }
  maybeAutoUpdate();
}

/**
 * Coalesce preview renders onto a trailing timer. Every control and every
 * datasheet keystroke calls renderPreview(); each render is a full
 * buildChart()+sceneToSvg()+innerHTML reparse, so firing one per keystroke was
 * measurable jank on a large sheet. A short trailing debounce collapses a burst
 * of edits into a single render of the final state. (snapshot() stays
 * synchronous — undo granularity depends on it.)
 */
let renderTimer: ReturnType<typeof setTimeout> | undefined;
function renderPreview() {
  clearTimeout(renderTimer);
  renderTimer = setTimeout(renderPreviewNow, 80);
}

function renderActionState() {
  const insertBtn = $("insert") as HTMLButtonElement;
  insertBtn.textContent = state.editTarget ? "Update chart" : "Insert into slide";
  ($("insert-new") as HTMLButtonElement).style.display = state.editTarget ? "" : "none";
  // This label is rewritten every time the edit target changes — long after
  // localizePane translated the pane — so it has to be re-translated or it
  // reverts to English. localizeTree only looks at descendants, hence the
  // parent.
  if (insertBtn.parentElement) localizeTree(insertBtn.parentElement);
}

// --- boot ------------------------------------------------------------------

// Datasheet undo/redo (Ctrl+Z / Ctrl+Y while the pane has focus).
const history: string[] = [];
const redoStack: string[] = [];
/**
 * What an undo step has to restore — the grid AND the two arrays that hang off
 * it by position.
 *
 * It used to be the cells alone, and `seriesColors`/`seriesMeta` are kept in
 * step only by the `onStructure` handler below, which fires on the EDIT and has
 * no inverse. So undo rewound the grid and left the side-channels spliced: take
 * a three-series combo, delete the middle row, press Ctrl+Z, and the row comes
 * back wearing the colour and the combo type of the series that followed it —
 * a chart that now says something different, from a keystroke whose entire
 * promise is that it changes nothing.
 *
 * Transpose was worse: its `reset` branch empties both arrays, and undo brought
 * the cells back to an empty palette with no way to recover but re-importing
 * the original JSON.
 *
 * The unit of undo is now the whole thing an edit changes. Serialising is also
 * the deep copy this needs — `seriesMeta[i].colors` is spliced in place.
 */
function snapshot() {
  const snap = JSON.stringify({
    cells: state.sheet.cells,
    seriesColors: state.seriesColors,
    seriesMeta: state.seriesMeta,
  });
  if (history[history.length - 1] !== snap) {
    history.push(snap);
    if (history.length > 100) history.shift();
    redoStack.length = 0;
  }
}
/**
 * Start a fresh undo timeline at the current sheet — the same baseline boot
 * establishes, re-established for each newly loaded chart. Without it Ctrl+Z
 * replayed the previous chart's cells into the new one.
 */
function resetHistory() {
  history.length = 0;
  redoStack.length = 0;
  history.push(
    JSON.stringify({ cells: state.sheet.cells, seriesColors: state.seriesColors, seriesMeta: state.seriesMeta }),
  );
}
function restore(snap: string) {
  const { cells, seriesColors, seriesMeta } = JSON.parse(snap) as {
    cells: string[][];
    seriesColors: AppState["seriesColors"];
    seriesMeta: AppState["seriesMeta"];
  };
  state.sheet = { cells };
  state.seriesColors = seriesColors;
  state.seriesMeta = seriesMeta;
  sheetApi.setSheet(state.sheet);
  renderPreview();
}
document.addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  if (e.key === "z" && history.length > 1) {
    e.preventDefault();
    redoStack.push(history.pop()!);
    restore(history[history.length - 1]);
  } else if (e.key === "y" && redoStack.length) {
    e.preventDefault();
    const snap = redoStack.pop()!;
    history.push(snap);
    restore(snap);
  }
});

sheetApi = mountDatasheet(
  $("datasheet"),
  state.sheet,
  (sheet) => {
    state.sheet = sheet;
    snapshot();
    renderPreview();
  },
  // seriesColors / seriesMeta are POSITIONAL, so a row or column the grid
  // splices in or out has to move them too — otherwise every series below the
  // edit inherits its neighbour's colour and combo type, and a routine
  // datasheet edit silently changes what the chart draws.
  (change) => {
    const colorsOf = (m: AppState["seriesMeta"][number]) => m?.colors;
    switch (change.kind) {
      case "series-insert":
        state.seriesColors.splice(change.index, 0, undefined);
        state.seriesMeta.splice(change.index, 0, undefined);
        break;
      case "series-remove":
        state.seriesColors.splice(change.index, 1);
        state.seriesMeta.splice(change.index, 1);
        break;
      case "category-insert":
      case "category-remove":
        // Per-point colours are indexed by CATEGORY, so they follow the column.
        for (const m of state.seriesMeta) {
          const cs = colorsOf(m);
          if (!cs) continue;
          if (change.kind === "category-insert") cs.splice(change.index, 0, null);
          else cs.splice(change.index, 1);
        }
        break;
      case "reset":
        // A transpose turns series into categories: no positional mapping survives.
        state.seriesColors = [];
        state.seriesMeta = [];
        break;
    }
  },
  // The sheet's own voice — today only the European-paste reading, which
  // rewrites the user's cells and therefore has to say so.
  //
  // "none", not "ok" and not "busy". Not "ok" because a green tick over a
  // reinterpretation of the user's numbers reads as confirmation rather than as
  // the flag it is. Not "busy" either, which was the first choice here and was
  // picked for its colour without reading what it DOES: "busy" un-hides the
  // progress bar and skips the settled count, so the pane would have implied
  // work in flight for a paste that had already finished. A neutral hint says
  // the one true thing — this happened, look at it — and claims nothing else.
  (message, params) => note(message, "none", params),
);
snapshot();

const titleInput = $("chart-title") as HTMLInputElement;
titleInput.value = state.title;
titleInput.addEventListener("input", () => {
  state.title = titleInput.value;
  renderPreview();
});
wireTabs();
wireActionsMenu();
document.getElementById("type-search-input")?.addEventListener("input", applyTypeFilter);
optionsHost.addEventListener("change", updateGroupCounts);
/**
 * THE REMEMBERED PICTURE PREFERENCE, SHOWN.
 *
 * `state.renderImage` is seeded from storage when the pane boots, but the only
 * place the checkbox is synced from state is `applyConfig` — which runs when a
 * chart is LOADED, not at startup. Without this line the pane would open with
 * the preference on and the box unticked: it would insert pictures while
 * telling the user it was drawing shapes, which is a worse fault than not
 * remembering the choice at all.
 *
 * Found by the test for the preference, not by reading the code.
 */
{
  const renderBox = document.getElementById("render-image") as HTMLInputElement | null;
  if (renderBox) renderBox.checked = state.renderImage;
}
renderGallery();
renderOptions();
renderPreview();
renderActionState();

$("download").addEventListener("click", () => {
  const cfg = currentConfig();
  const svg = sceneToSvg(buildChart(cfg), { background: canvasColor(cfg) });
  const blob = new Blob([svg], { type: "image/svg+xml" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "ssf-charts.svg";
  a.click();
  URL.revokeObjectURL(a.href);
});

/** Oversample for rasterized output: 2× the scene's point size, i.e. 144 dpi.
 *  A point-for-pixel PNG looks fuzzy the moment a slide is zoomed. */
const RASTER_SCALE = 2;

/**
 * Rasterize a scene to a bare base64 PNG for `InsertOptions.pictureBase64`.
 *
 * Deliberately **transparent** — no background is painted. Shapes mode drops the
 * chart's ink straight onto the real slide with no background rect at all (see
 * `canvasColor`), so painting one here would make image mode land an opaque
 * white box over a themed or dark slide. Transparency is what keeps the two
 * modes visually interchangeable, which is the whole promise of the flag.
 *
 * The `download-png` handler below does the same SVG→Image→canvas dance but ends
 * in `toBlob` for a file download; this one ends in `toDataURL` because Office.js
 * wants base64. The shared middle is three lines, and factoring it would couple
 * an export button to the insert path for no real gain.
 *
 * Resolves with the payload after the comma; rejects on a decode failure, a
 * missing 2D context, or a `toDataURL` that hands back something that isn't a
 * PNG data URI (jsdom returns null). The caller degrades to native shapes.
 */
function rasterizeScene(scene: Scene): Promise<string> {
  const svg = sceneToSvg(scene);
  return new Promise<string>((resolve, reject) => {
    const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
    const img = new Image();
    img.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(scene.width * RASTER_SCALE);
        canvas.height = Math.round(scene.height * RASTER_SCALE);
        const ctx = canvas.getContext("2d");
        if (!ctx) throw new Error("no 2d canvas context");
        ctx.scale(RASTER_SCALE, RASTER_SCALE);
        ctx.drawImage(img, 0, 0, scene.width, scene.height);
        const uri = canvas.toDataURL("image/png");
        if (!uri || !uri.startsWith("data:image/png")) throw new Error("canvas could not encode a PNG");
        resolve(uri.slice(uri.indexOf(",") + 1));
      } catch (err) {
        reject(err instanceof Error ? err : new Error(String(err)));
      } finally {
        URL.revokeObjectURL(url);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("the browser could not decode the chart SVG"));
    };
    img.src = url;
  });
}

/**
 * The picture payload for one insert, plus the reason there isn't one.
 *
 * `warn` is returned rather than noted here on purpose: every insert path posts
 * `phaseNote` as its first act, which would immediately overwrite a note posted
 * before the insert. The caller notes it AFTER the insert resolves, where it
 * survives — and where it counts as this action's settlement, so `guard` leaves
 * it standing instead of closing with "Done."
 */
/**
 * A picture for a scene this host cannot draw, or null when it can draw it.
 *
 * SHARED, because ELEMENTS need it too and they do not go through
 * `chartPicture` at all — they call `insertSceneIntoSlide` directly. That is
 * where the marks in question actually live: a Harvey ball is a wedge at every
 * fraction between 1% and 99% (0% and 100% are ellipses and survive, so exactly
 * the informative range is lost), and a table `[up]`/`[down]` cell is an
 * arrowhead. Shipping this on the chart path alone would have left the two
 * element kinds that need it most silently dropping their glyphs.
 */
async function pictureForUndrawableMarks(scene: Scene): Promise<{ png?: string; warn: string } | null> {
  const dropped = marksThisHostWillDrop(scene);
  if (!dropped.nodes) return null;
  const missing = dropped.what.join(" and ");
  if (!canInsertPicture())
    return {
      warn:
        `This PowerPoint cannot draw ${missing}, and cannot insert a picture either (that needs PowerPointApi 1.8). ` +
        `Drawing it as shapes — ${missing} will be missing.`,
    };
  const png = await boundedRaster(scene);
  return png
    ? {
        png,
        warn:
          `This PowerPoint cannot draw ${missing} — that needs a newer Office (PowerPointApi 1.10). ` +
          `Inserted as a picture so it is complete; it stays an image here.`,
      }
    : {
        warn:
          `This PowerPoint cannot draw ${missing}, and it could not be turned into a picture either. ` +
          `Drawing it as shapes — ${missing} will be missing.`,
      };
}

async function chartPicture(cfg: ChartConfig, scene: Scene): Promise<{ png?: string; warn?: string }> {
  /**
   * MARKS THIS HOST CANNOT DRAW — asked before density, because this one is
   * about the chart being WRONG rather than about the host being slow.
   *
   * Below PowerPointApi 1.10 a wedge and an arrowhead each trace and draw
   * nothing, so a pie inserts as four labels around empty space. 18 of the 123
   * shipped charts lose ink there and 8 lose their subject outright.
   *
   * DELIBERATELY NOT `wantsAutoPicture`, and the difference is the whole point:
   * that one gates on `web`, because density is what kills PowerPoint on the
   * WEB. This is the opposite host — the web reports 1.10 and never reaches
   * here, while desktop, Mac and every volume-licensed build never have it.
   * Sharing the predicate would have made this dead code on the only hosts it
   * exists for.
   *
   * A picture is the whole chart, correct, on a host that cannot draw it. The
   * message says what would have been missing and does NOT offer Explode:
   * exploding is the door back to the broken version, which is the mistake
   * `doExplode`'s own comment records the product having made once already.
   */
  if (cfg.render !== "image") {
    const undrawable = await pictureForUndrawableMarks(scene);
    if (undrawable) return undrawable;
  }

  // A chart nobody asked to rasterize, on the one host that cannot survive
  // drawing it. The densest kinds are far past what PowerPoint on the web will
  // take as shapes — violin 253, area 176, tile map 122, waffle 103 — and the
  // budget below is the same number the demo deck has always used to decide a
  // chart "too dense for this host". It used to skip those and stamp a
  // placeholder; drawing a picture is strictly better than not drawing.
  const shapes = estimateOfficeShapes(scene);
  if (
    wantsAutoPicture(shapes, {
      web: isWebHost(),
      canPicture: canInsertPicture(),
      alreadyPicture: cfg.render === "image",
    })
  ) {
    // Bounded, like every other rasterize on a rescue path. A canvas that
    // never fires `onload` does not throw and does not resolve — it simply
    // stops, and an unbounded await here hung Insert, Same Scale and the JSON
    // batch insert forever: buttons disabled, no error, no way out but
    // reloading the pane. Same Scale awaits one of these per chart, so a
    // single wedged decode froze the whole deck's operation.
    const png = await boundedRaster(scene);
    if (png)
      return {
        png,
        /**
         * THIS USED TO PROMISE SOMETHING THE PRODUCT THEN REFUSED.
         *
         * It read: `Inserted as a picture; "Explode to native shapes" turns it
         * back.` And `doExplode`, reached by the button that sentence names,
         * asks `wantsAutoPicture` with THE SAME ARGUMENTS as the branch you are
         * reading — and on the web answers "so it stays a picture. Open it on
         * the desktop app to explode it, or make the chart simpler."
         *
         * One predicate, two contradictory promises. Being told the door back
         * exists and then finding it locked is worse than never being offered
         * it, and it is the exact complaint a rival markets this whole category
         * against: "a colleague asks you to just tweak the Q3 bar, and there's
         * nothing to tweak."
         *
         * What is TRUE is better than what was promised, and was never said:
         * the picture carries the chart's config tag, so this pane can still
         * load it, change its data and restyle it. Only the conversion to
         * native SHAPES needs a host that can draw them.
         */
        warn:
          `That chart is ${shapes} shapes — too many for PowerPoint on the web to draw, so it went in as a picture. ` +
          `You can still edit it here: select it and this pane reloads its data. ` +
          `Turning it into native shapes needs the desktop app.`,
      };
    // The rescue itself failed. Drawing the shapes anyway is still better than
    // refusing the user's chart, but this is precisely the case the budget
    // exists to intercept — so say so instead of silently doing the dangerous
    // thing. It used to fall through with no message at all.
    return {
      warn: `That chart is ${shapes} shapes — too many for PowerPoint on the web to draw, and it could not be turned into a picture either. Drawing it as shapes; if PowerPoint stops responding, that is why.`,
    };
  }
  if (cfg.render !== "image") return {};
  if (!canInsertPicture()) {
    return {
      warn: "This PowerPoint can't insert pictures (needs PowerPointApi 1.8) — inserted native shapes instead.",
    };
  }
  const png = await boundedRaster(scene);
  if (png) return { png };
  console.warn("PC-IMG-RASTER could not rasterize the chart in the browser.");
  return { warn: "Couldn't rasterize the chart — inserted native shapes instead." };
}

/**
 * PNG export: the native-shapes output is the real deliverable, but SVG doesn't
 * render in email/chat, so a rasterized fallback earns its place. Rasterize the
 * SAME preview SVG through a canvas at 2× for a crisp bitmap. The SVG carries no
 * foreignObject or external refs, so the canvas never tags as tainted and
 * toBlob() is allowed. Best-effort: a decode failure surfaces as an error note
 * rather than a silent no-op.
 */
$("download-png").addEventListener("click", () => {
  const cfg = currentConfig();
  const scene = buildChart(cfg);
  const svg = sceneToSvg(scene, { background: canvasColor(cfg) });
  const scale = 2;
  const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
  const img = new Image();
  img.onload = () => {
    try {
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(scene.width * scale);
      canvas.height = Math.round(scene.height * scale);
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("no 2d canvas context");
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, scene.width, scene.height);
      canvas.toBlob((png) => {
        if (!png) {
          note("Couldn't encode the PNG on this browser.", "err");
          return;
        }
        const a = document.createElement("a");
        a.href = URL.createObjectURL(png);
        a.download = "ssf-charts.png";
        a.click();
        URL.revokeObjectURL(a.href);
      }, "image/png");
    } catch (err) {
      note("Couldn't render PNG: {error}", "err", { error: err instanceof Error ? err.message : String(err) });
    } finally {
      URL.revokeObjectURL(url);
    }
  };
  img.onerror = () => {
    URL.revokeObjectURL(url);
    note("Couldn't render the preview to PNG.", "err");
  };
  img.src = url;
});

/**
 * Encode the current chart into a shareable deep link. The config rides in the
 * URL hash (base64 of the JSON) so it never hits a server log, and the hash is
 * decoded on boot — reopening the exact chart on the hosted gallery. Round-trips
 * are UTF-8-safe via encodeURIComponent before btoa.
 */
const CONFIG_HASH = "#c=";
$("copy-link").addEventListener("click", async () => {
  const json = JSON.stringify(currentConfig());
  const link = location.origin + location.pathname + location.search + CONFIG_HASH + btoa(encodeURIComponent(json));
  try {
    await navigator.clipboard.writeText(link);
    note("Shareable chart link copied to the clipboard.", "ok");
  } catch {
    // Clipboard blocked (no permission / not focused): drop the link into the
    // JSON box so it can still be copied by hand.
    ($("json-io") as HTMLTextAreaElement).value = link;
    note("Clipboard blocked — the link is in the JSON box, copy it from there.", "err");
  }
});

/** Cascading default insert position so repeated inserts don't pile up. */
let insertOffset = 0;

/**
 * Say which host phase we are in. A stalled Office.js sync never throws — it
 * simply never settles — so "Working…" alone cannot tell a slow host from a
 * dead one. Naming the phase makes a stall legible: whatever it says last is
 * where it stopped.
 */
function phaseNote(phase: InsertPhase, detail?: string) {
  const said: Record<InsertPhase, string> = {
    context: "opening PowerPoint…",
    queue: "building shapes…",
    commit: "sending to PowerPoint…",
    group: "grouping…",
    done: "done",
  };
  // The phase word is itself a catalogue key; the detail (a dynamic id) trails
  // untranslated in parentheses.
  note("Working… {phase}", "busy", { phase: t(said[phase]) + (detail ? ` (${detail})` : "") });
}

async function doInsert(asNew: boolean) {
  insertInFlight = true;
  hostWorkStarted();
  try {
    return await runInsert(asNew);
  } finally {
    insertInFlight = false;
    hostWorkFinished();
  }
}

async function runInsert(asNew: boolean) {
  let cfg = currentConfig();
  if (!asNew && state.editTarget) {
    const scene = buildChart(cfg);
    const { png, warn } = await chartPicture(cfg, scene);
    // Adopt the target the update hands back. Every shape was replaced, so the
    // one we just used is dead: keeping it made the SECOND update resolve a
    // shape id that no longer existed, get filtered out as "the user deleted
    // this chart", and do nothing — silently. With auto-update on, that meant
    // only the first debounced push ever landed.
    const { next, swapped, duplicated, picture, recovered } = await updateChartResilient(scene, state.editTarget, {
      tagData: JSON.stringify(cfg),
      pictureBase64: png,
    });
    if (duplicated) {
      // Both slides hold the chart: the rebuilt one and the original PowerPoint
      // would not remove. Nothing here can safely pick one — deleting the wrong
      // slide loses the user's notes or transitions — so name it and let them.
      state.editTarget = null;
      renderActionState();
      note(
        "Rebuilt that slide, but PowerPoint would not remove the original — the chart is now on two slides. Delete whichever you don't want.",
        "err",
      );
      return;
    }
    if (swapped) {
      // The chart is on a NEW slide, so the old target is dead. Say so rather
      // than leaving a stale target that makes every later push no-op.
      state.editTarget = null;
      renderActionState();
      // Say what was LOST, not just what was done.
      //
      // A swap replaces the slide: the chart comes back, and anything the
      // slide carried that is not a shape does not. Speaker notes, the
      // transition, animations, slide-level formatting. The guard in front of
      // this checks the slide holds no other SHAPES, and that is all it can
      // check — Office.js offers no way to read notes at all (office-js#3269,
      // in backlog), so the add-in cannot see what it is about to discard and
      // cannot ask first.
      //
      // Which leaves telling the user afterwards. Losing a slide's notes is
      // survivable if you know it happened and silent data loss if you do not,
      // and this is the whole difference between the two.
      note(
        "Rebuilt that slide — PowerPoint would not redraw it in place. The chart is back, but the slide was " +
          "REPLACED, so any speaker notes, transition or animation on it are gone (undo restores them). " +
          "Select the chart to keep editing.",
        "err",
      );
      return;
    }
    if (next?.lost) {
      // Redrawn, and not something the pane can come back to. Two different
      // reasons, and neither may be kept as the live edit target: an untagged
      // chart cannot be re-opened at all, and a target whose shape id was never
      // read back names the shape this very update deleted — pushing to it
      // again would resolve a dead id, or a group member, and draw a second
      // chart over the first.
      //
      // Both used to print "Done." in green over a chart that was quietly no
      // longer editable, and the `no-config` half is the one a real host
      // produced: four charts in one run lost their config tag and nothing
      // said so.
      state.editTarget = null;
      renderActionState();
      note(
        next.lost === "no-config"
          ? "Redrawn, but PowerPoint would not save the chart's settings back onto it — it is no longer editable " +
              "from the pane. Undo (Ctrl+Z) restores the version that was."
          : "Redrawn, but PowerPoint would not say where the new chart landed, so the pane has lost track of it. " +
              "Click the chart and press Edit it to carry on.",
        "err",
      );
      return;
    }
    if (next) {
      state.editTarget = next;
      /**
       * THREE FACTS, AND TWO OF THEM USED TO BE UNREACHABLE.
       *
       * `updateChartResilient` returns `{ next, picture, recovered }` together
       * (see its `return { next, picture: true, recovered: !!wreckage }`), but
       * this ladder returned on the first match. So `recovered` — the chart was
       * DESTROYED by the failed in-place update and redrawn from scratch — was
       * only ever said when `next` was null, and the `warn` explaining a
       * too-dense chart was swallowed by an `else if` whenever the ladder had
       * fallen to a picture.
       *
       * `recovered` leads: "your chart was destroyed and rebuilt" outranks how
       * it was drawn, and it is the one a user would want to know about their
       * own document.
       *
       * The standalone `if (recovered)` below is untouched — it still handles
       * the `next === null` case it was written for, and its comment explains
       * why that case must be told apart from "gone".
       */
      const setbacks = [
        recovered && t("PowerPoint would not redraw that chart in place, so it was redrawn from scratch."),
        picture && !png && t("Drawn as a picture — PowerPoint would not redraw the shapes."),
      ].filter((s): s is string => !!s);
      const said = insertOutcomeSentence(setbacks, warn ? t(warn) : "");
      if (said) note(said, "err");
      return;
    }
    if (recovered) {
      // The chart IS on the slide — redrawn from scratch after the in-place
      // update destroyed it — the host just would not tag it, so there is no
      // target to keep editing from. Distinguishing this from "gone" is the
      // whole point: telling the user to insert it again would give them two.
      state.editTarget = null;
      renderActionState();
      note("Redrawn as a picture — PowerPoint would not redraw it in place. Select the chart to keep editing.", "err");
      return;
    }
    // null means the target slide or shape is gone — nothing was written. The
    // caller's guard saw an unchanged note and printed "Done." in green, and the
    // stale target kept the button reading "Update chart", so every later push
    // (including every auto-update) no-opped just as silently. Say so, and fall
    // back to inserting a new chart.
    state.editTarget = null;
    renderActionState();
    note("That chart is no longer on the slide — insert it again.", "err");
    return;
  }
  // New chart: use the selected placeholder's bounds when one is selected.
  const bounds = await getSelectionBounds();
  const intoPlaceholder = !!bounds && bounds.width > 40 && bounds.height > 40;
  // Read the selection, then let go of it — see `dropShapeSelection`.
  //
  // Two published web bugs fire on exactly this flow, and both need a shape to
  // still be selected while we draw: office-js#2775 deletes the selected shape
  // when a text box is added (every chart here has text boxes), and
  // office-js#3698 refuses a picture insert while one is selected (this same
  // call inserts a picture for a chart too dense to draw). The bounds have
  // already been read by this line, so nothing downstream wants the selection.
  await dropShapeSelection();
  // Where the chart goes, and at what size — decided BEFORE the scene is built,
  // because both branches can change the size and the raster has to be of the
  // scene that sizes the picture's rect. A raster of a differently-sized scene
  // is a stretched chart.
  // `moved` is absent on the placeholder branch, and deliberately: the user
  // picked that position themselves, so there is nothing to report about it.
  // Hoisted so the slow-insert offer below can price the target slide from the
  // read this path already makes. `null` means the host would not say.
  let occupied: Awaited<ReturnType<typeof getSlideShapeBounds>> = null;
  let at: {
    left: number;
    top: number;
    width: number;
    height: number;
    shrunk?: boolean;
    moved?: Placement["moved"];
  };
  if (intoPlaceholder) {
    at = { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height };
  } else {
    // Nothing selected, so the chart has to be placed — beside or below
    // whatever is already on the slide, shrunk into the space left if it will
    // not fit at full size. A fixed 14pt cascade against a 480x300 chart
    // overlapped the previous one by better than 90%, which is what "they are
    // built on top of each other" looked like. The cascade survives as the last
    // resort, for a slide with room in neither direction.
    //
    // The slide's real size, not an assumed one: whether a second chart can sit
    // beside the first is entirely a question about the width, and a 4:3 deck is
    // 240pt narrower than the 16:9 this used to take for granted.
    //
    // A host that will not describe the slide gets the cascade, and gets it
    // HERE rather than inside `placeChart`. `placeChart` cannot tell the
    // difference: an empty `occupied` means "there is room everywhere", so
    // `placeBeside` succeeds on its first pass and returns the origin unmoved,
    // and its `fallback` argument is never reached. Two inserts onto a slide the
    // host would not read therefore landed on exactly the same point — the pile
    // this whole rule exists to prevent, and worse than the fixed cascade it
    // replaced. A real host refused every shape read on a whole deck.
    // The ENGINE's clamp, not `??`. A config arrives from the JSON box, a saved
    // template, a shape tag written in another deck and the skill's caller, so
    // `width` is a number in the types and anything at all in practice —
    // `??` passes NaN, Infinity, 0 and negatives straight through. `placeChart`
    // then shrinks a NaN width to fit and hands back 882 points, which is wider
    // than the slide and finite enough to survive every check after it.
    const size = {
      width: clampDim(cfg.width, DEFAULT_SIZE.width),
      height: clampDim(cfg.height, DEFAULT_SIZE.height),
    };
    occupied = await getSlideShapeBounds();
    at = occupied
      ? placeChart(
          occupied,
          size,
          { left: 60, top: 90 },
          { left: 60 + insertOffset, top: 90 + insertOffset },
          await slideSize(),
        )
      : { left: 60 + insertOffset, top: 90 + insertOffset, ...size, shrunk: false, moved: "cascade" };
    insertOffset = (insertOffset + 14) % 84;
  }
  cfg = { ...cfg, width: at.width, height: at.height };
  const scene = buildChart(cfg);
  /**
   * THE SLOW-INSERT OFFER. Adding to a loaded slide costs several times what
   * the same chart costs on an empty one, and nothing used to say so — the pane
   * simply looked hung for half a minute.
   *
   * Offered on `worthOwnSlide`, not on `isSlowInsert`. Slow is not the same as
   * worth moving: a big chart is slow on a blank slide too, and offering a new
   * slide there would cost a slide and save nothing.
   *
   * Priced from `occupied`, the read this path ALREADY made to place the chart.
   * No second host call: an earlier draft added one, and duplicating a read the
   * insert had just done is the "adding counting sites moves the baseline"
   * mistake this repo warns about.
   *
   * `occupied === null` means the host would not say. Treated as UNKNOWN rather
   * than empty and left silent: a warning that fires because a read failed is
   * worse than one that never fires.
   *
   * That guard is INTENT, not protection, and the mutation check is what said
   * so. Pricing a refused read as an empty slide instead leaves every test
   * green and cannot be caught by any — `worthOwnSlide(n, 0)` is false for
   * every n, since the empty-slide estimate is never at most half of itself, so
   * "unknown" and "empty" reach the same silence by different routes. It is
   * kept because the day this rule stops being a ratio against the empty-slide
   * cost, the two stop agreeing and the difference becomes real.
   */
  let ownSlideId: string | undefined;
  /**
   * WHAT WENT WRONG ON THE WAY, collected rather than posted.
   *
   * `note()` is one slot and the last write wins, so a setback posted here was
   * destroyed by `phaseNote` moments later — the insert's first act is a busy
   * note into the same element, which is why "After the insert, never before"
   * is already written down further down this function. The user's chosen route
   * had failed and the pane told them only the outcome.
   *
   * AN ARRAY, NOT A STRING, because two can stack: the generated-slide fallback
   * leaves `pricedShapes` at 1, and one shape onto a slide already holding
   * fifty-odd still prices past `SLOW_INSERT_MS` — so the own-slide offer fires
   * after it and can fail too. Two failed routes are one bad host, and the user
   * should hear about both.
   *
   * Translated at PUSH time: the first carries an interpolated count, so it has
   * stopped being a catalogue key by the time it is composed. See
   * `insertOutcomeSentence`.
   */
  const setbacks: string[] = [];
  const sceneShapes = estimateOfficeShapes(scene);
  /**
   * THE PICTURE IS DECIDED BEFORE THE OFFER, because the offer quotes a TIME and
   * the time depends on it.
   *
   * `chartPicture` used to run twenty lines below this, after the question had
   * already been asked and answered. So a user who had ticked "Insert as
   * picture" — a checkbox that has shipped for some time — was told a violin
   * onto a sixty-shape slide "takes about 7 minutes", when what was about to
   * happen was one `setImage` and one sync. The sentence was describing an
   * insert the pane had already decided not to perform.
   *
   * Rasterising first costs about 700ms before the question appears. That is the
   * price of the number being true.
   */
  let pic = await chartPicture(cfg, scene);
  // WHAT WILL ACTUALLY BE INSERTED, not what the scene contains. A picture is
  // one shape and one sync — `wantsPicture` in the renderer takes that path on
  // the PAYLOAD, so `png` being present is the honest test, not `cfg.render`.
  const pricedShapes = pic.png ? 1 : sceneShapes;
  /**
   * THE BUDGET FORCED A PICTURE — so ask, instead of just doing it.
   *
   * Guarded on `pic.png` and on the budget predicate together: this is only the
   * charts the renderer refused to draw here, not every picture. A user who
   * ticked "Insert as picture" asked for one and is not questioned about it,
   * which is why `cfg.render !== "image"` is part of the test.
   *
   * See `offerNativeInsteadOfPicture` for the measurements and for why the
   * choice is put to the user rather than taken for them.
   */
  if (
    pic.png &&
    cfg.render !== "image" &&
    canInsertSlidesFromBase64() &&
    wantsAutoPicture(sceneShapes, { web: isWebHost(), canPicture: canInsertPicture(), alreadyPicture: false }) &&
    (await offerNativeInsteadOfPicture(sceneShapes)) === "own-slide"
  ) {
    note("Building a slide for it…", "busy");
    const landed = await chartOnItsOwnSlideAsFile(scene, cfg);
    if (landed === 1) {
      state.editTarget = null;
      renderActionState();
      note("Added on a slide of its own, as editable shapes. Select it and this pane reloads its data.", "ok");
      return;
    }
    // ANY other outcome falls through to the picture, which is the floor. A
    // landed count that is not exactly 1 means the host took something other
    // than what was handed over, and `insertSlidesFromPptx` measures that from
    // a before-and-after slide count rather than trusting the call not to throw.
    setbacks.push(
      t("Could not add the slide{took} — inserting here as a picture.", {
        took: landed > 1 ? ` (the host took ${landed})` : "",
      }),
    );
  }
  if (occupied && worthOwnSlide(pricedShapes, occupied.length)) {
    const choice = await offerOwnSlide(
      occupied.length,
      // Offer the switch only when there is something to switch TO: a host that
      // can rasterise, and a chart not already going in as one.
      !pic.png && canInsertPicture(),
    );
    if (choice === "picture") {
      // The user chose it here rather than in the pane, so the CONFIG changes —
      // both auto-picture guards downstream test `cfg.render === "image"`
      // literally, and a sibling flag would slip past them.
      cfg.render = "image";
      pic = await chartPicture(cfg, scene);
      // NAMES THE CHOICE ONLY. `chartPicture` always returns a `warn` on this
      // path ("Couldn't rasterize the chart — inserted native shapes instead."),
      // and the old wording here repeated it almost verbatim — composing the two
      // would have said the same thing twice. This says which choice failed and
      // lets the outcome carry the remedy.
      if (!pic.png) setbacks.push(t("Could not render the picture you chose."));
    }
    if (choice === "own-slide") {
      note("Adding a slide…", "busy");
      // `null` when the host dropped the add even after `addSlides` retried. The
      // chart then goes where the user originally asked — slowly, but it goes.
      // Losing the offer is a nuisance; losing the chart is not acceptable.
      ownSlideId = (await addSlideForChart()) ?? undefined;
      if (!ownSlideId) setbacks.push(t("Could not add a slide — inserting here instead."));
    }
  }
  // PLACEMENT IS RECOMPUTED FOR AN EMPTY SLIDE. `at` was cascaded around the
  // shapes on the CROWDED slide, so carrying it over would drop the chart into
  // a corner of a blank one for no reason — offset from obstacles that are not
  // there.
  const place = ownSlideId ? { left: 60, top: 90 } : { left: at.left, top: at.top };
  // Rendered ABOVE, before the offer, so the time it quoted could be true. Not
  // re-rendered here: a second rasterise of the same scene is 700ms spent to
  // produce the same bytes.
  const { png, warn } = pic;
  const inserted = await insertSceneIntoSlide(
    scene,
    { tagData: JSON.stringify(cfg), ...place, pictureBase64: png, ...(ownSlideId ? { slideId: ownSlideId } : {}) },
    phaseNote,
  );
  state.editTarget = null;
  renderActionState();
  // After the insert, never before: phaseNote would have overwritten it.
  //
  // A chart with no config tag is on the slide and is not an SSF chart: click
  // it and the pane says so, reopen the deck and the settings are gone for
  // good. The renderer has always known — `groupAndTagAll` returns `tagged` —
  // and this path threw the whole return away, so `guard()` printed "Done." in
  // green over it. The demo path has a repair pass for exactly this; the
  // everyday insert has none, which makes saying so the only thing left.
  /**
   * CHOSEN HERE, POSTED ONCE BELOW, so a setback can lead it.
   *
   * Same four conditions and the same texts as before — only the verb changed,
   * from posting to assigning. `undefined` is the fourth outcome and always
   * was: an ordinary insert with nothing to report.
   */
  let outcome: { text: string; status: "ok" | "err" } | undefined;
  if (inserted?.lost === "no-config")
    outcome = {
      text:
        "The chart is on the slide, but PowerPoint would not save its settings onto it — the pane cannot re-open " +
        "this one. Insert it again if you need to keep editing it.",
      status: "err",
    };
  else if (warn) outcome = { text: warn, status: "err" };
  // Say where it went as well as whether it was scaled. A chart that appears
  // BESIDE the last one rather than under it is the one placement outcome a
  // user has no reason to expect, and silence about it reads as the add-in
  // having put the chart somewhere at random.
  else if (at.moved === "beside")
    outcome = {
      text: at.shrunk
        ? "Placed beside the last chart and scaled to fit — drag or resize it as you like."
        : "Placed beside the last chart — drag or resize it as you like.",
      status: "ok",
    };
  else if (at.shrunk)
    outcome = { text: "Scaled to fit the space left on the slide — drag or resize it as you like.", status: "ok" };
  /**
   * SETBACK FIRST, and RED whenever there is one.
   *
   * A setback beside an "ok" placement is an error overall — the user's choice
   * did not happen, and reporting that in green is the defect this replaces.
   *
   * The outcome is translated HERE and the setbacks were translated at push
   * time, so `insertOutcomeSentence` only joins. Passing the joined string to
   * `note` is safe precisely because both halves are already localised: `t()`
   * misses on the join and returns it unchanged, which is the documented
   * fallback rather than an accident.
   *
   * THE THIRD BRANCH IS LOAD-BEARING. A setback with no outcome must still
   * post, or nothing settles and `guard` paints "Done." in green over a
   * failure — which is worse than the bug being fixed. It is also the case that
   * used to end on "Working… done" in busy blue, because the setback settled,
   * suppressed "Done.", and left the phase note as the last thing written.
   */
  const said = insertOutcomeSentence(setbacks, outcome ? t(outcome.text) : "");
  if (said) note(said, setbacks.length || outcome?.status === "err" ? "err" : "ok");
}

/**
 * What Same Scale did, then what it cost — one sentence for one slot.
 *
 * THREE NOTES USED TO GO OUT BACK TO BACK: the applied count, then a rescued
 * clause, then a degraded clause, into a channel that holds one. Each
 * replacement re-said "Same scale applied" and dropped the numbers the one
 * before it carried, so a run with both qualifiers showed only the degraded
 * clause — losing the rescued count, the chart count and the shared max.
 *
 * The rescued charts are the ones that cost most to lose. They went in as
 * PICTURES, and "Explode to native shapes" is the way back; saying nothing
 * about them steers the user away from the one control that recovers them.
 * `pane-host-actions.test.ts` already guards that harm from the other
 * direction — "does not report a successful auto-picture rescue as a fallback
 * to shapes" — and this was the same harm by a different route.
 *
 * OUTCOME FIRST, THEN QUALIFIERS, which is why this is not
 * `insertOutcomeSentence`. That one puts setbacks first and its docstring
 * argues for it; the same join under the opposite rule would make one of the
 * two comments a lie.
 *
 * TAKES TRANSLATED TEXT. Each clause is its own catalogue key with its own
 * params, so `t()` runs at the call site and this only joins. Exported for the
 * same reason `elapsedLabel` is: so what it says can be asserted directly.
 *
 * A RESCUE IS NOT A FAILURE — the comment at the call site is emphatic that it
 * is the guard working — so only `degraded` makes this red.
 */
export function sameScaleNote(parts: { base: string; rescued?: string; degraded?: string }): {
  text: string;
  status: "ok" | "err";
} {
  return {
    text: [parts.base, parts.rescued, parts.degraded].filter(Boolean).join(" "),
    status: parts.degraded ? "err" : "ok",
  };
}

/**
 * think-cell's Set Same Scale: pin every value-axis chart (in the deck, or
 * just the selected ones) to the union of their extents and re-render them.
 */
async function doSameScale(scope: "deck" | "selection" = "deck") {
  // A deck-wide scan that could not see the whole deck cannot be the basis for
  // "the deck now shares one scale" — that sentence is the entire feature, and
  // it was being printed in green over a scan the code already knew was short.
  // A real host produced `unread=8 slides=8` and, later in the same run,
  // `unread=7 slides=8`; the self-test caught the consequence by re-reading the
  // deck afterwards (`1 of 6 charts carry the shared scale`) and this path,
  // which does not re-read, would have said 6 of 6.
  const scan = scope === "deck" ? await listChartsInDeck() : null;
  const charts = scan ? scan.charts : await listChartsInSelection();
  if (scan && !scanIsComplete(scan)) {
    note("Same scale needs to see the whole deck first — {gap}. Try again; nothing was changed.", "err", {
      gap: scanGap(scan),
    });
    return;
  }
  const parsed = charts
    .map((c) => ({ target: c.target, cfg: JSON.parse(c.configJson) as ChartConfig }))
    .map((c) => ({ ...c, extent: valueExtent(c.cfg) }))
    .filter((c): c is typeof c & { extent: { min: number; max: number } } => c.extent != null);
  if (parsed.length < 2) {
    note(
      scope === "deck"
        ? "Same scale needs at least two value-axis charts in the deck."
        : "Select two or more charts (Ctrl-click), then apply Same scale.",
      "err",
    );
    return;
  }
  const min = Math.min(...parsed.map((c) => c.extent.min));
  const max = Math.max(...parsed.map((c) => c.extent.max));
  // One request context for the whole deck, not one per chart: each chart's
  // update costs four round-trips to PowerPoint, and awaiting them in a loop
  // made Same Scale across 20 charts eighty of them.
  // Apply the shared scale FIRST, then build. Rasterising inside the same map
  // would capture the pre-scale scene — an image chart would be re-inserted
  // showing the old axis while the pane reported success.
  const rescaled = parsed.map((c) => {
    c.cfg.scale = { min: min < 0 ? min : undefined, max };
    return { ...c, scene: buildChart(c.cfg) };
  });
  // Rasterise every image-mode chart before opening the request context — one
  // context for the whole deck is the property doSameScale exists to protect,
  // and awaiting a canvas inside it would be both slower and unsafe.
  const pictures = await Promise.all(rescaled.map((c) => chartPicture(c.cfg, c.scene)));
  // A chart whose redraw stalls has already had its old shapes deleted, so it
  // is left blank and the user has to be told which. Silence here meant a
  // deck-wide operation could quietly empty charts the user never looked at.
  const stalled: string[] = [];
  /**
   * What each stalled redraw destroyed before it gave up.
   *
   * Same Scale deletes a chart's old shapes and redraws them, per chart. A
   * redraw that stalls part-way leaves whatever it committed on the slide —
   * a few axis lines, half a series — and the single-chart path has swept
   * that since `updateChartResilient` learned to. This path never did: it
   * told the user the charts were "now empty" and left the debris sitting on
   * them, so the one operation that touches every chart in the deck was also
   * the one that left the most behind.
   */
  const wreckage: UpdateWreckage[] = [];
  const applied = await updateChartsInSlides(
    rescaled.map((c, i) => ({
      scene: c.scene,
      target: c.target,
      opts: { tagData: JSON.stringify(c.cfg), pictureBase64: pictures[i].png },
    })),
    (item, err) => {
      const w = wreckageOf(err);
      if (w?.strayIds.length) wreckage.push(w);
      // A chart the USER stopped is not a chart the host refused. Its debris is
      // still swept below — stopping mid-batch leaves the same partial chart a
      // stall does — but naming it in "PowerPoint would not redraw…" would
      // blame the host for the user's own decision.
      if (isStopped(err)) return;
      stalled.push(item.scene.title || "an untitled chart");
    },
  );
  // Sweep AFTER the batch, never from inside the callback. `onFailed` fires
  // within updateChartsInSlides' own request context, and `deleteShapesById`
  // opens a fresh one on purpose — the context that just stalled is precisely
  // the one that cannot be trusted to carry the repair.
  if (wreckage.length) {
    let swept = 0;
    let strays = 0;
    for (const w of wreckage) {
      strays += w.strayIds.length;
      swept += await deleteShapesById(w.slideId, w.strayIds);
    }
    console.warn(
      `SSF Charts: same scale swept ${swept} of ${strays} shapes left by ${wreckage.length} stalled redraw(s)`,
    );
  }
  // Charts the host silently declined to resolve at all.
  //
  // `updateChartsInSlides` drops those — its `live`/`alive` filters are early
  // returns rather than throws, so `onFailed` never fires and `stalled` stays
  // empty. A stalled chart IS in the returned array (it comes back with its old
  // target), so these two counts never double-report the same chart. The
  // success note used to say `parsed.length`, the charts REQUESTED, which is
  // how a run that rescaled one of six reported six with nothing else said.
  const missed = rescaled.length - applied.length;
  if (stalled.length) {
    note(
      "Same scale: PowerPoint would not redraw {n} chart(s) — {which}. They are now empty; undo (Ctrl+Z) restores them." +
        (missed > 0 ? " {missed} more were never reached — the deck is not on one scale." : ""),
      "err",
      {
        n: stalled.length,
        which: stalled.join(", "),
        missed,
      },
    );
    return;
  }
  // Stopped part-way: the charts that were rescaled keep their new scale and the
  // rest keep their old one, so claiming the deck is now on one scale would be
  // false. `guard()` posts the "Stopped" note; this just declines to overwrite
  // it with a success message first.
  if (isStopRequested()) return;
  if (missed > 0) {
    note(
      "Same scale reached {n} of {total} charts — PowerPoint would not answer for the other {missed}. " +
        "The deck is not on one scale; try again.",
      "err",
      { n: applied.length, total: rescaled.length, missed },
    );
    return;
  }
  // `warn` alone is not a failure. `chartPicture` returns a warn WITH a png on
  // its SUCCESS path — the web auto-picture rescue, where a chart too dense for
  // this host was rasterised and inserted as a picture, and the warn is the
  // explanation. Counting every warn reported those in RED as "fell back to
  // native shapes", the opposite of what happened: it claimed the dangerous
  // thing had occurred when the guard against it had just worked, and steered
  // the user away from "Explode to native shapes" — the one control that turns
  // a picture back.
  const rescued = pictures.filter((p) => p.warn && p.png).length;
  const degraded = pictures.filter((p) => p.warn && !p.png).length;
  /**
   * ONE NOTE, NOT THREE. These used to be posted one after another into a
   * single slot, so the last silently erased the others — and each opened with
   * "Same scale applied" because each was written to stand alone. Composed,
   * the base says that once and the clauses qualify it, which is why both lost
   * their prefix. See `sameScaleNote`.
   */
  const said = sameScaleNote({
    base: t("Same scale applied to {n} charts (max {max}).", { n: applied.length, max }),
    rescued: rescued
      ? t(
          '{n} chart(s) were too dense for this host and went in as pictures — "Explode to native shapes" turns them back.',
          {
            n: rescued,
          },
        )
      : "",
    degraded: degraded ? t("{n} image chart(s) fell back to native shapes.", { n: degraded }) : "",
  });
  note(said.text, said.status);
}

async function doLoadSelection() {
  const found = await loadChartFromSelection();
  if (!found) {
    note("The selection is not an SSF chart — select an inserted chart group first.", "err");
    return;
  }
  note("Chart loaded — edits will update it in place.", "ok");
  applyConfig(JSON.parse(found.configJson) as ChartConfig, found.target);
  // The banner offers to load the selected chart; once it is loaded the offer
  // is stale. Hiding it here rather than in the banner's own click handler
  // covers the other way in — the "Edit selected chart" button.
  $("selection-banner").style.display = "none";
}

/**
 * Turn the selected picture chart back into native, editable shapes.
 *
 * The whole cost of image mode is that the chart stops being editable, so this
 * is the door back. It adds NO renderer code: read the config off the picture's
 * tag, force `render: "shapes"`, and hand it to the ordinary in-place update
 * with **no** `pictureBase64`. That omission is the explode — `wantsPicture`
 * goes false and the node loop runs instead.
 *
 * Works on a shapes-mode chart too, where it is an idempotent redraw rather than
 * a special case worth refusing. Acts on the SELECTION, which need not be the
 * chart the pane is currently editing.
 */
async function doExplode() {
  const found = await loadChartFromSelection();
  if (!found) {
    note("Select an inserted chart first — Explode re-draws it as native shapes.", "err");
    return;
  }
  const cfg = { ...(JSON.parse(found.configJson) as ChartConfig), render: "shapes" as const };
  const scene = buildChart(cfg);
  /**
   * THE SAME BUDGET THE INSERT PATH USES, because this is the door back through
   * the wall it put up — and it had no lock on it.
   *
   * A chart over the budget is auto-inserted as a picture with a message that
   * ends "…\"Explode to native shapes\" turns it back." This function then
   * rebuilt with `render: "shapes"` and went straight to `updateChartInSlide`,
   * never consulting `wantsAutoPicture` — whose only other call site in the
   * codebase is the insert path. So the product invited the user through the
   * one door it had just told them was dangerous.
   *
   * The cost is not theoretical: 17 of the 123 shipped charts exceed the budget
   * and the largest is 401 shapes, which at this host's measured rate is many
   * minutes of work, well past the 425-470s window where PowerPoint has been
   * dying all week. The user loses the session and whatever was unsaved, having
   * followed the add-in's own instruction.
   *
   * Refuse with the count. `canPicture` matters: a host that cannot rasterise
   * has no picture to fall back TO, so exploding is the only way to see the
   * chart at all and the guard correctly stands aside.
   *
   * THE SAME MISTAKE HAS A SECOND DOOR, and it is the one this host actually
   * walks through. A chart inserted as a picture because this PowerPoint cannot
   * draw its wedges would explode into shapes with the wedges MISSING — the
   * user asked for native shapes and would get a pie with no slices, having
   * followed a button the add-in offered. Refused for the same reason and in
   * the same shape as the density guard, and the insert-time message is careful
   * not to advertise Explode on this path at all.
   */
  const dropped = marksThisHostWillDrop(scene);
  if (dropped.nodes) {
    note(
      `This PowerPoint cannot draw ${dropped.what.join(" and ")} — that needs a newer Office ` +
        `(PowerPointApi 1.10) — so the chart stays a picture. Exploding it here would leave them out.`,
      "err",
    );
    return;
  }
  const shapes = estimateOfficeShapes(scene);
  if (
    wantsAutoPicture(shapes, {
      web: isWebHost(),
      canPicture: canInsertPicture(),
      alreadyPicture: false,
    })
  ) {
    note(
      `That chart is ${shapes} shapes — too many for PowerPoint on the web to draw, so it stays a picture. ` +
        `Open it on the desktop app to explode it, or make the chart simpler.`,
      "err",
    );
    return;
  }
  // Same live-canvas wall as an ordinary update: exploding a picture draws
  // every shape onto the slide in view. Look away while it happens.
  const next = await withSlideDeselected([found.target.slideId], (deselected) =>
    updateChartInSlide(scene, found.target, {
      tagData: JSON.stringify(cfg),
      ...(deselected ? { shapesPerSync: OFFSCREEN_BATCH } : {}),
    }),
  );
  if (!next) {
    // The picture is gone from the slide. Only clear the pane's edit target if
    // it was pointing at this same shape — Explode works on the selection.
    if (state.editTarget?.shapeId === found.target.shapeId) {
      state.editTarget = null;
      renderActionState();
    }
    note("That chart is no longer on the slide.", "err");
    return;
  }
  if (next.lost) {
    // The shapes are on the slide and the pane cannot come back to them —
    // exactly the case the ordinary update path handles at length, and this
    // door onto the same host answer used to print green and adopt the dead
    // target anyway. On a web host that refuses a shape tag write (the 5010
    // family, recorded five times in a single round) Explode claimed success
    // over a chart that could no longer be opened, and installed the lost
    // target as `state.editTarget` — so the next push resolved a dead id, was
    // filtered out as "the user deleted this chart", and did nothing. With
    // auto-update on, every debounced push after it did nothing too.
    state.editTarget = null;
    renderActionState();
    $("selection-banner").style.display = "none";
    note(
      next.lost === "no-config"
        ? "Exploded to native shapes, but PowerPoint would not save the chart's settings onto them — it is no " +
            "longer editable from the pane. Undo (Ctrl+Z) restores the picture."
        : "Exploded to native shapes, but PowerPoint would not say where they landed, so the pane has lost track " +
            "of them. Click the chart and press Edit it to carry on.",
      "err",
    );
    return;
  }
  // Adopt the returned target: every shape was replaced, so the old one is dead
  // (the trap doInsert documents — a stale target makes the NEXT update no-op).
  applyConfig(cfg, next);
  $("selection-banner").style.display = "none";
  note("Exploded to native shapes — edits now update them in place.", "ok");
}

// --- Elements (harvey balls, checkboxes, process flow, table) -----------------

wireElementPreviews();

// --- Templates & style file ----------------------------------------------------

const TEMPLATES_KEY = "ssf-charts-templates";
const STYLE_KEY = "ssf-charts-style";

/**
 * The keys these two used to have, read once when the new ones are empty.
 *
 * Renamed with the product on 2026-08-27. These hold WORK THE USER DID — saved
 * chart templates and a deck style — so a plain rename would not lose a
 * setting, it would lose their templates, silently, on the first load after an
 * update. The pane would come up looking correct and emptied.
 *
 * Read-through rather than a migration write: the old value is used when the
 * new key is absent, and the next save writes the new key. Nothing is deleted,
 * so a downgrade still finds its data.
 */
const LEGACY_TEMPLATES_KEY = "powerchart-templates";
// The style's own fallback is spelled out at its read site instead of using a
// constant from here: that read runs during MODULE EVALUATION, above this
// declaration, so a `const` reference would hit the temporal dead zone.

/** The new key's value, or the old key's if this browser predates the rename. */
function readStored(key: string, legacy: string): string | null {
  try {
    return localStorage.getItem(key) ?? localStorage.getItem(legacy);
  } catch {
    return null;
  }
}

/**
 * The saved templates, on an object with NO prototype.
 *
 * This table is keyed by whatever the user typed into a name box, which makes
 * it the third in this repo to need the same guard — the notes say to apply it
 * to every one, and this one was missed. Both directions were wrong:
 *
 * - **Writing.** `all[name] = config` is a plain assignment for every name but
 *   one. For `__proto__` it hits the setter inherited from `Object.prototype`
 *   and re-parents the object instead of storing anything. `Object.keys` then
 *   does not see it, `JSON.stringify` does not write it, and the template the
 *   user just saved is gone — with a list that redraws as if nothing happened.
 * - **Reading.** `all[name]` for a name nobody saved — `constructor`,
 *   `toString`, `valueOf` — walks up to `Object.prototype` and hands back a
 *   FUNCTION, which is truthy. Only the picker being built from `Object.keys`
 *   stops that reaching `applyConfig` today, which is an accident rather than
 *   a guard.
 *
 * A null prototype fixes both at the root instead of at each call site: there
 * is no inherited setter to hit and nothing above to walk up to.
 */
function loadTemplates(): Record<string, ChartConfig> {
  const empty = () => Object.create(null) as Record<string, ChartConfig>;
  try {
    // `JSON.parse` creates `__proto__` as an OWN property (it uses
    // CreateDataProperty, not assignment), so a template stored under that name
    // survives the round trip once the target cannot be re-parented.
    return Object.assign(empty(), JSON.parse(readStored(TEMPLATES_KEY, LEGACY_TEMPLATES_KEY) ?? "{}"));
  } catch {
    return empty();
  }
}

function renderTemplateList() {
  const sel = $("template-list") as HTMLSelectElement;
  sel.innerHTML = "<option value=''>— templates —</option>";
  const starters = document.createElement("optgroup");
  starters.label = "Starters";
  for (const t of BUILTIN_TEMPLATES) {
    const opt = document.createElement("option");
    opt.value = `builtin:${t.name}`;
    opt.textContent = t.name;
    starters.appendChild(opt);
  }
  sel.appendChild(starters);
  const names = Object.keys(loadTemplates()).sort();
  if (names.length) {
    const mine = document.createElement("optgroup");
    mine.label = "My templates";
    for (const name of names) {
      const opt = document.createElement("option");
      opt.value = `user:${name}`;
      opt.textContent = name;
      mine.appendChild(opt);
    }
    sel.appendChild(mine);
  }
}

$("template-save").addEventListener("click", () => {
  const name = prompt("Template name?", state.title || state.kind);
  if (!name) return;
  const all = loadTemplates();
  all[name] = currentConfig();
  localStorage.setItem(TEMPLATES_KEY, JSON.stringify(all));
  renderTemplateList();
});
$("template-list").addEventListener("change", () => {
  const value = ($("template-list") as HTMLSelectElement).value;
  if (!value) return;
  const sep = value.indexOf(":");
  const [source, name] = [value.slice(0, sep), value.slice(sep + 1)];
  const cfg = source === "builtin" ? BUILTIN_TEMPLATES.find((t) => t.name === name)?.config : loadTemplates()[name];
  if (cfg) applyConfig({ ...DEFAULT_SIZE, ...cfg }, null);
});
$("template-delete").addEventListener("click", () => {
  const value = ($("template-list") as HTMLSelectElement).value;
  if (!value.startsWith("user:")) return; // starters (and the placeholder) can't be deleted
  const name = value.slice("user:".length);
  const all = loadTemplates();
  delete all[name];
  localStorage.setItem(TEMPLATES_KEY, JSON.stringify(all));
  renderTemplateList();
});
renderTemplateList();

$("style-export").addEventListener("click", () => {
  // What the pane is DRAWING with, which since decks carry their own style is
  // not always the browser's copy. Exporting the copy while the preview shows
  // the deck's would hand the user a file that does not describe what they can
  // see — and this is the text they are about to send a colleague.
  const current: StyleFile = { ...resolveStyleFile(deckStyle, styleFile, stylePrefer).style };
  if (state.paletteName === "Theme" && themePalette) current.palette = themePalette;
  else if (state.paletteName !== "Default") current.palette = namedPalette(state.paletteName);
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(current, null, 2);
  note("Style exported — share the JSON as your corporate style file.", "ok");
});
$("style-import").addEventListener("click", () => {
  try {
    const parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
    if (parsed.kind) throw new Error("that is a chart config — use Import instead");
    styleFile = parsed;
    localStorage.setItem(STYLE_KEY, JSON.stringify(styleFile));
    // An import is an explicit act by the person at the pane, so it wins over
    // the deck's style for this session — otherwise the deck would silently
    // undo what they just did, on the very chart they are looking at. "Use deck
    // style" is the way back, which is why the flip needs no confirmation.
    stylePrefer = "mine";
    renderPreview();
    note("Style imported — applied to every chart from this pane.", "ok");
  } catch (err) {
    note("Style import failed: {error}", "err", { error: err instanceof Error ? err.message : String(err) });
  }
});

/**
 * Put the style in force into the DECK, where it travels with the file.
 *
 * Writes what the pane is actually drawing with — the resolved style, not the
 * browser's copy — because the button is beside a preview and what it saves has
 * to be what the user can see. On a host below PowerPointApi 1.7 the write
 * cannot happen at all, and the note says which of the two it was: a style
 * nobody stored and a style stored somewhere invisible fail the same way from
 * the outside, and only one of them is worth retrying.
 */
$("style-to-deck").addEventListener("click", async () => {
  const style = resolveStyleFile(deckStyle, styleFile, stylePrefer).style;
  const ok = await writeDeckStyle(isEmptyStyle(style) ? null : style);
  if (!ok) {
    note("This host cannot store a deck style (needs PowerPoint API 1.7).", "err");
    return;
  }
  deckStyle = isEmptyStyle(style) ? null : style;
  stylePrefer = "deck";
  note("Style saved into this deck — it travels with the file.", "ok");
});

// Host-only, like "Use deck theme" beside the palette: outside PowerPoint there
// is no deck to store a style in, and a button that can only ever report its own
// impossibility is worse than one that is visibly unavailable.
//
// AND IT HAS TO RUN AGAIN WHEN OFFICE IS READY. This line is at module scope, so
// it runs while `Office.context` is still undefined — `isPowerPointHost()` reads
// `!!Office.context?.host` and answers FALSE for every load inside PowerPoint.
// Both buttons were then disabled forever, because nothing ever asked again.
//
// Measured in the pane after round 091, on the real host:
//
//     Office.context.host  "PowerPoint"     <- would answer true NOW
//     style-from-deck      disabled: true
//     style-to-deck        disabled: true
//
// So #583's whole user-facing feature was unreachable on PowerPoint web, and no
// round could see it: thirteen scenarios drive charts, and none clicks these.
export function syncHostOnlyButtons(): void {
  for (const id of ["style-to-deck", "style-from-deck", "style-deck-theme"]) {
    const el = document.getElementById(id) as HTMLButtonElement | null;
    if (el) el.disabled = !isPowerPointHost();
  }
}
syncHostOnlyButtons();

/** Go back to whatever the deck says, after an import switched to the browser's copy. */
$("style-from-deck").addEventListener("click", async () => {
  const { style, unreadable } = await readDeckStyleWithReason();
  // A READ THAT FAILED IS NOT AN ABSENCE. This said "This deck carries no style"
  // whenever the read came back empty for ANY reason, and round 089 caught the
  // reason that makes it a lie: `reading the deck's style` hung for the full 90s
  // on the real host. Telling someone their deck is unbranded — and switching
  // them to the browser's style — on the strength of a timeout is the same
  // mistake as reading an unreadable deck scan as "these slides are new".
  if (unreadable) {
    note("Could not read this deck's style — leaving yours in place. Try again.", "err");
    return;
  }
  deckStyle = style;
  stylePrefer = "deck";
  renderPreview();
  note(deckStyle ? "Using this deck's own style." : "This deck carries no style — using yours.", "ok");
});

// --- Automation (JSON in / out, the open .ppttc idea) -------------------------

$("json-export").addEventListener("click", () => {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(currentConfig(), null, 2);
});
$("json-import").addEventListener("click", () => {
  /**
   * TWO FAILURES, TWO MESSAGES. One `try` covered both, so anything
   * `applyConfig` threw was reported as a SYNTAX error. Measured on configs
   * that are perfectly valid JSON:
   *
   *   {"kind":"clustered","data":null}
   *     -> Invalid JSON: Cannot read properties of null (reading 'series')
   *   {"kind":"clustered","data":{"categories":"A","series":"x"}}
   *     -> Invalid JSON: data.series.some is not a function
   *
   * That sends the user hunting a syntax error that is not there, in text any
   * validator will tell them is fine. `stateFromConfig`'s own comment records
   * this symptom ("palette.join is not a function") and `asArray` made the
   * common shapes survivable — it did not make the MESSAGE honest, and the
   * shapes it does not cover still arrive here.
   */
  let parsed: unknown;
  try {
    parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
  } catch (err) {
    note("Invalid JSON: {error}", "err", { error: err instanceof Error ? err.message : String(err) });
    return;
  }
  try {
    const one = (Array.isArray(parsed) ? parsed[0] : parsed) as ChartConfig;
    applyConfig({ ...DEFAULT_SIZE, ...one }, null);
    if (Array.isArray(parsed))
      note('Loaded chart 1 of {total} — use "Insert batch" for all.', "ok", { total: parsed.length });
    else note("Chart config loaded.", "ok");
  } catch (err) {
    // The JSON parsed; the chart it describes is not one this pane can open.
    note("That is valid JSON, but not a chart this pane can open: {error}", "err", {
      error: err instanceof Error ? err.message : String(err),
    });
  }
});

// --- Agenda ------------------------------------------------------------------

wireAgendaPreview();

// Auto-update: push edits to the slide shortly after each change.
function maybeAutoUpdate() {
  // Cancel any pending push BEFORE the guard. The guard returns early once the
  // edit target is gone — which is exactly what loading another chart does — and
  // a timer still armed against the previous one would go on to fire. doInsert
  // then finds no edit target, takes its new-chart branch, and drops a chart
  // nobody asked for onto the slide.
  clearTimeout(autoUpdateTimer);
  const on = ($("auto-update") as HTMLInputElement | null)?.checked;
  if (!on || !state.editTarget || !isPowerPointHost()) return;
  autoUpdateTimer = setTimeout(() => {
    // Never two writes to the same chart at once. This timer bypasses
    // `guard()`, so it does not see the disabled buttons a click would — and
    // one resilient update can now legitimately take tens of seconds (a 45s
    // stall, then a slide swap, then a bounded raster). A user who keeps
    // editing through that could start a second update against the SAME stale
    // edit target; whichever finished last won, and if that was the one
    // started from the already-superseded target it reported "that chart is
    // no longer on the slide" about a chart that had just been written fine.
    //
    // Re-arm rather than drop: the edit still needs to reach the slide, it
    // just has to wait its turn.
    if (insertInFlight) {
      maybeAutoUpdate();
      return;
    }
    void doInsert(false).catch(() => {});
  }, 900);
}
// Unticking has to cancel a push that is already in flight; ticking on its own
// shouldn't push anything until the next edit.
$("auto-update").addEventListener("change", () => {
  if (!($("auto-update") as HTMLInputElement).checked) clearTimeout(autoUpdateTimer);
});

// The picture/shapes choice changes nothing about the SCENE, so the preview is
// unaffected — it always shows the vector chart. It only changes what the next
// insert hands PowerPoint, which is why this touches state and nothing else.
$("render-image").addEventListener("change", () => {
  state.renderImage = ($("render-image") as HTMLInputElement).checked;
  // REMEMBERED, so it is a preference rather than something to re-tick on every
  // chart. It seeds a NEW chart only — see `PICTURE_PREF_KEY`; a chart loaded
  // off a slide still shows its own `render`, because that is a fact about that
  // chart and not a wish about the next one.
  //
  // Wrapped, because a browser with storage blocked must not take the pane down
  // over a checkbox. Failing to remember the choice is a small loss; failing to
  // make it is not.
  try {
    localStorage.setItem(PICTURE_PREF_KEY, state.renderImage ? "1" : "0");
  } catch {
    /* a pane that cannot remember still works */
  }
});

/**
 * How long the selection watcher gives the host before it stops waiting.
 *
 * Short on purpose, and much shorter than `READBACK_TIMEOUT_MS`. Nothing
 * depends on this read: it decides whether one banner is visible. The default
 * budget is sized for work a user asked for and is waiting on; a background
 * listener that inherits it holds a promise open for a minute and a half over
 * a hint. A real host produced exactly that, twice in one run —
 * `gave up waiting what=reading the selected chart afterMs=90000`, both of them
 * during Same Scale, which is not a selection feature at all.
 */
const SELECTION_WATCH_BUDGET_MS = 4_000;

/** A selection read is already outstanding. */
let selectionReadInFlight = false;

/** think-cell's "click the chart" feel: watch the slide selection. */
function watchSelection() {
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      void onSelectionChanged();
    });
  } catch {
    /* event unavailable on this host */
  }
}

/**
 * Decide whether to offer the selected chart — and, far more often, decide not
 * to ask.
 *
 * The event this answers is not only the user clicking a shape. The pane moves
 * the selection itself, constantly: `showSlide` on every deck-wide pass,
 * `withSlideDeselected` before every off-screen redraw, `selectShape` in the
 * self-test, and a slide swap on every resilient update. A deck run therefore
 * fires this handler dozens of times, and the old version answered every one of
 * them with an unbounded `PowerPoint.run` — queued behind work the user
 * actually asked for, on a host that was already the bottleneck. Two of those
 * reads were still outstanding ninety seconds later on a real host, on a run
 * that then killed the tab.
 *
 * So: while the pane is working, record that the banner is stale and do
 * nothing. `hostWorkFinished` runs exactly one read afterwards, which is the
 * only one whose answer was ever going to be current anyway.
 */
async function refreshSelectionBanner(): Promise<void> {
  if (selectionReadInFlight) return;
  selectionReadInFlight = true;
  try {
    const found = await loadChartFromSelection(SELECTION_WATCH_BUDGET_MS);
    const banner = $("selection-banner");
    banner.style.display = found && found.target.shapeId !== state.editTarget?.shapeId ? "" : "none";
  } catch {
    // A hiccup or a deadline. Leave the banner as it is rather than guessing:
    // hiding it loses a true offer, showing it makes a false one.
  } finally {
    selectionReadInFlight = false;
  }
}

async function onSelectionChanged(): Promise<void> {
  if (hostBusy()) {
    selectionMissed = true;
    return;
  }
  await refreshSelectionBanner();
}

/**
 * A one-line host descriptor for the demo title/results slides, e.g.
 * "PowerPoint · OfficeOnline · 16.0.1". Web vs desktop vs Mac behave differently
 * under the demo's shape budget, so a run's PDF should say which one it was.
 * Guarded — `Office.context.diagnostics` is absent outside a host.
 */
function describeHost(): string {
  try {
    const d = Office.context?.diagnostics;
    if (d?.host) return `${d.host} · ${d.platform} · ${d.version}`;
  } catch {
    /* Office.context unavailable — fall through */
  }
  return "unknown host";
}

/**
 * The last demo run, kept at module scope so it can be saved AFTER the fact.
 *
 * A run's own record used to live and die inside the click handler's closure:
 * when a run ended badly the summary slide was the first casualty, and the
 * only surviving evidence was whatever the user thought to copy out of a
 * one-line note. A run that reveals a host bug is worth more than that.
 */
let lastRunLog: RunLogFile | undefined;

/**
 * Slides the last round added, and the only thing Clean up will ever delete.
 *
 * A list of ids the round watched appear, not a rule for recognising a test
 * slide. The difference matters in someone's own deck: a rule can match a slide
 * they made, an id list cannot.
 */
let tidyable: string[] = [];

/**
 * What "Download run log" writes: one file, one or more runs inside it.
 *
 * A single click can now take both insert paths one after the other, and they
 * fail in completely different ways — so they are separate runs with separate
 * identities, sharing only the build and host that produced them. A
 * single-path click writes the same shape with one entry, so nothing reading
 * this file needs to care which it was.
 */
/**
 * How many slides make a deck too crowded to draw a whole demo onto shape by
 * shape.
 *
 * Not a measurement of where the host gives up — nobody has one, because every
 * attempt so far started on a deck the file half had already filled. It is a
 * threshold below which the run is worth attempting: a demo adds 38 slides, so
 * anything past a handful means the deck was not fresh.
 */
const CROWDED_DECK_SLIDES = 10;

interface RunLogFile {
  build: string;
  host: string;
  /**
   * The slide size this round ran at.
   *
   * Optional because 53 archived rounds predate it, and every one of those was
   * 16:9 — anything reading this must default to that rather than guess, and
   * `docs/ROUNDS.md` states it once so no reader has to infer it.
   */
  slideSize?: { width: number; height: number; source: string };
  /**
   * How many `context.sync()` calls the round made, end to end.
   *
   * Optional because every round archived before 2026-09-02 predates the
   * counter — and absent must read as "never counted", never as zero, which is
   * a number this field can genuinely take on a round that did nothing.
   *
   * See `syncsSoFar` in `src/render/powerpoint.ts` for why syncs rather than
   * shapes: office-js#6329 reports that each one forces a full presentation
   * save on the web host, which would make this the load figure that matters
   * and every shape count in this archive the wrong index.
   */
  syncs?: number;
  runs: RunLog[];
  /**
   * The host self-test's verdicts, when that is what produced this log.
   *
   * Kept beside the runs rather than inside one: the scenarios are not an
   * insert and have no slots, and folding them into a run's item list would
   * make them look like charts that failed to draw.
   */
  selftest?: ScenarioResult[];
  /**
   * The self-test's trace. A run carries its own; the self-test is not a run,
   * so its steps hang here. Declared rather than spread in untyped — an
   * undeclared field is one nothing downstream can be expected to read.
   */
  trace?: RunLog["trace"];
  /**
   * The host probe's answer sheet, when the round produced one.
   *
   * Declared rather than spread in untyped, for the reason the field above
   * gives: an undeclared field is one nothing downstream can be expected to
   * read. `npm run host-diff` takes either shape — a bare sheet, or a round's
   * file with this inside it — so one upload covers both halves.
   */
  hostAnswers?: HostAnswerSheet;
  /**
   * Set when the run did LESS than it was asked to, and why.
   *
   * "Both, one after the other" degrades to the file half on a deck that is
   * already large — see `CROWDED_DECK_SLIDES`. Without this the log shows a run
   * asked for two halves that produced one, with nothing to say the difference
   * was deliberate, and the obvious reading is that the shape half crashed.
   * Data rather than a trace line, because the trace is optional and this is
   * not.
   */
  refusedShapeHalf?: { slides: number; why: string };
  /**
   * What the deck actually held when the round finished, and what it looked like.
   *
   * This is the upload that used to be a person's job. Every diagnosis in this
   * project's history has needed three things — the run log, the deck, and a
   * screenshot — and the owner has been saving and sending all three by hand,
   * once per round, for as long as there have been rounds. Two of the three are
   * things the add-in can read for itself: `SlideInventory` is every shape on
   * every slide (names, ids, positions), and `SlideShot` is the host's own
   * rendering of the slides the round touched.
   *
   * `newSlides` is the id diff across the round, which is what makes the
   * pictures worth having: a picture of every slide in a 40-slide deck is
   * megabytes of things nobody changed.
   */
  deck?: {
    inventory: SlideInventory[];
    /** What the scan could not see — the same honesty every other scan reader owes. */
    gap?: string;
    newSlides: string[];
    /**
     * The before-list could not be read, so `newSlides` is empty for want of an
     * answer rather than because nothing was added. Present only in that case.
     */
    beforeUnknown?: boolean;
    shots: SlideShot[];
  };
}

interface RunLog {
  /**
   * This run's identity, the same token carried on every slide's slot tag.
   *
   * The join key between this file and the .pptx it produced. Every diagnosis
   * in this project's history has come from reading the two together, and
   * until now that read started by guessing which slides belonged to the run —
   * a deck holds whatever earlier runs left in it, and one recent file carried
   * 30 slides from the run under investigation plus a lone slide from another.
   * With the token the join is a lookup and the stray is labelled, not
   * mistaken for a loss. `npm run triage` does exactly that join.
   */
  run: string;
  totalMs: number;
  items: {
    title: string;
    status: string;
    shapes: number;
    ms: number;
    grouped: boolean;
    /**
     * What the run believes it wrote — NOT what the settled readback saw.
     * The two disagreeing is the whole point: a run reported 20 tagged charts
     * where the produced file carried 31, and establishing that took unzipping
     * the .pptx. Compare against `reconcile.snapshots[].tagged` for the same
     * slot: true here and false there is a readback fault; false here and true
     * there is impossible; false in both is a genuinely lost write.
     */
    tagged: boolean;
    /**
     * Whether this item was MEANT to carry a config — not whether it got one.
     *
     * Not every demo item is a chart: the title page, the contents pages and
     * several elements are drawn as `PowerChart` objects with no config by
     * design. Nothing in the produced file distinguishes those from a chart
     * whose tag write was lost, and on the shape path `tagged` cannot either
     * (false there means both "never had one" and "the write did not land").
     * Without this field a triage of a clean run called seven healthy slides
     * broken.
     */
    chart: boolean;
    /**
     * This item gave up on a host call. Separate from `lateOutcome` because
     * the host usually answers an abandoned call minutes later, long after the
     * run has moved on — so "which item stalled" is readable here even when
     * "how it ended" is not.
     */
    abandoned: boolean;
    lateOutcome: string;
  }[];
  deck: {
    slidesAdded: number;
    addsIssued: number;
    lost: number;
    blank: { position: number; title: string | null }[];
  };
  /** The settled truth: what the deck held once the host stopped moving. */
  reconcile?: ReconcileOutcome;
  /** Why there is no settled verdict, when there is none. */
  unverified?: string;
  /**
   * Which route the run took: one generated file, or shape by shape. They fail
   * in completely different ways, and a log that does not say which is being
   * read is a log that gets diagnosed as the wrong one.
   */
  path: "file" | "shapes";
  /**
   * Step-by-step record, when Verbose trace was on for the run — tallies
   * first, then the entries they were counted from.
   */
  trace?: {
    summary: TraceSummary;
    entries: { ms: number; scope: string; message: string; data?: Record<string, unknown> }[];
    dropped: number;
  };
}

/**
 * Verify and repair the demo slides wherever they are in the deck.
 *
 * By slot tag, never by position. The deck's own layout is the host's
 * business: `insertSlidesFromBase64` puts slides at the FRONT unless told
 * otherwise, and a run that assumed "the slides we just added are at the end"
 * read one slide short at the front and swept in the user's own title slide at
 * the back. The span between the first and last tagged slide is ours; nothing
 * outside it is read, let alone touched.
 *
 * "Tagged" means tagged BY THIS RUN. The span used to be drawn around every
 * SSF Charts slot tag in the presentation, which is only the same thing the
 * first time the deck is inserted. Insert it twice — or duplicate one demo
 * slide, which copies its tag — and the span covered both copies, every item
 * matched two slides, and the pass deleted one whole healthy run as the other
 * one's duplicate. Anything inside the span that is not this run's is reported
 * and left alone.
 */
async function repairDeckSpan(
  expected: ExpectedItem[],
  tagFor: (slot: number) => string | undefined,
  run: string,
): Promise<VerifyResult> {
  let snapshots: SlideSnapshot[];
  // Slides the readback could not see. Anything on one of them comes back
  // `lost`, so this is what stops the run reporting a chart the host kept as a
  // chart the host dropped.
  let unread: number;
  try {
    const count = await slideCount();
    // Pass A only, deck-wide. It has to be deck-wide — the span is discoverable
    // only by reading slot tags, and this run's slides are not necessarily at
    // the tail (`insertSlidesFromBase64` put them at the FRONT the first time
    // anyone tried it on a real host). Passes B and C run below, over the span
    // alone: their answers outside it are read, paid for, and then thrown away
    // by the `inSpan` filter twenty lines down — a tag read and two group-count
    // syncs per slide, spent on the user's own earlier work.
    ({ snapshots, unread } = await readAddedSlides(0, count, false));
  } catch (err) {
    return { kind: "error", why: errorText(err) };
  }
  const tagged = snapshots.filter((s) => s.run === run);
  // No slot tag anywhere means the read found slides but could not identify
  // one of them. Saying "verified" would be a lie and saying nothing is how a
  // whole verification pass went missing from a run report without anyone
  // noticing — so name it.
  if (!tagged.length) {
    return {
      kind: "unidentified",
      why: snapshots.length
        ? `read ${snapshots.length} slide(s), none carrying this run's slot tag`
        : "read no slides at all",
    };
  }
  const first = Math.min(...tagged.map((s) => s.index));
  const last = Math.max(...tagged.map((s) => s.index));
  const inSpan = snapshots.filter((s) => s.index >= first && s.index <= last);
  // Now that the span is known, enrich only it.
  await enrichSnapshots(inSpan);
  const plan = planReconcile(inSpan, expected, { dropOrphanBlanks: true, run });
  try {
    if (!plan.actions.length)
      return {
        kind: "ok",
        outcome: { snapshots: inSpan, plan, applied: { unstamped: 0, regrouped: 0, deleted: 0 }, refused: 0, unread },
      };
    return {
      kind: "ok",
      outcome: { ...(await applyReconcilePlan(plan, tagFor, { left: 60, top: 90 }, inSpan)), unread },
    };
  } catch (err) {
    return { kind: "error", why: errorText(err) };
  }
}

/**
 * Why a verification pass produced no verdict, when it produced none.
 *
 * A `null` here used to mean three different things — nothing tagged, a host
 * that would not answer, and a repair that threw — and the caller could tell
 * them apart from none of them. The first fast-path run reported a raw slide
 * count instead of a settled verdict and gave no hint why, which is precisely
 * the failure mode this whole line of work exists to eliminate.
 */
type VerifyResult =
  { kind: "ok"; outcome: ReconcileOutcome } | { kind: "unidentified"; why: string } | { kind: "error"; why: string };

/**
 * Update a chart in place without asking the live canvas to do the impossible.
 *
 * Redrawing a chart is the add-in's worst case on PowerPoint web: every shape
 * replaced, onto the one slide guaranteed to be on screen. A real run died on
 * the FIRST batch — "did not respond while drawing shapes 1-10 of 39". Three
 * attempts, least destructive first:
 *
 *  1. Look away and redraw. Selecting another slide puts the target
 *     off-screen, where the host swallows 40 shapes a sync instead of choking
 *     at 10, and the selection is restored afterwards. Nothing is lost.
 *  2. Swap the slide. Generate a one-slide .pptx and replace the slide with
 *     it — one host call, no drawing at all. Only offered when the slide holds
 *     NOTHING but the chart, because the replacement is a new slide and does
 *     not carry the old one's notes or transitions. The user is told when this
 *     happens, since their edit target moves with it.
 *  3. Draw it as a picture. Never stalls, keeps the config tag so the chart is
 *     still re-editable and can be exploded back to shapes — but it is a
 *     raster, and that is a real cost, so it is the floor and not the default.
 */
async function updateChartResilient(
  scene: Scene,
  target: EditTarget,
  opts: { tagData?: string; pictureBase64?: string },
): Promise<{
  next: EditTarget | null;
  swapped?: boolean;
  duplicated?: boolean;
  picture?: boolean;
  /** The chart was re-drawn from scratch after a failed update destroyed it. */
  recovered?: boolean;
}> {
  let stall: unknown;
  /**
   * What layer 1 destroyed before it failed, when it destroyed anything.
   *
   * Everything below used to assume the chart was still on the slide, and layer
   * 1's first act is to delete it — so once layer 1 had run at all, layer 2
   * disqualified itself on the litter and layer 3 re-resolved a shape id that
   * no longer existed and reported the chart "no longer on the slide". Three
   * fallbacks, none of them reachable, in exactly the case they are for.
   */
  let wreckage: UpdateWreckage | undefined;
  try {
    const next = await withSlideDeselected([target.slideId], (deselected) =>
      updateChartInSlide(scene, target, deselected ? { ...opts, shapesPerSync: OFFSCREEN_BATCH } : opts),
    );
    return { next };
  } catch (err) {
    console.warn("SSF Charts: in-place redraw stalled — trying the slide swap", err);
    stall = err;
    wreckage = wreckageOf(err);
  }

  // Clear the half-drawn chart before anything else draws over it. Whatever
  // comes next puts a whole chart back on this slide, and without the sweep the
  // user keeps the wreckage underneath it — the failure made visible twice.
  if (wreckage?.strayIds.length) {
    const swept = await deleteShapesById(wreckage.slideId, wreckage.strayIds);
    console.warn(`SSF Charts: swept ${swept} of ${wreckage.strayIds.length} shapes left by the stalled redraw`);
  }

  // A stop is not a stall, and the ladder below is for stalls. Every rung of it
  // exists to get the chart drawn by some other means — rebuild the slide,
  // rasterize it — which is the one thing a user who just pressed Stop has said
  // they do not want. Running it anyway would make Stop mean "draw this a
  // different way", and the slowest rungs would then run AFTER the cancel.
  //
  // The sweep above still happens: stopping mid-batch leaves the same debris a
  // stall does, and leaving it there would make Stop destructive.
  if (isStopped(stall)) throw stall;

  if (canInsertSlidesFromBase64() && opts.tagData && (await slideHoldsOnlyChart(target.slideId))) {
    try {
      note("Rebuilding that slide…", "busy");
      const built = await buildDeckBase64([{ scene, title: "Chart", configJson: opts.tagData }], await slideSize());
      const swap = await replaceSlideWithDeck(target.slideId, built.base64);
      if (swap === "swapped") return { next: null, swapped: true };
      // The new slide landed; only the old one's removal failed. Falling
      // through to the picture layer here would rasterize the chart onto that
      // surviving original — leaving the user with the chart twice, once as
      // shapes and once as a picture, and a message saying it went fine. Stop
      // and say what happened instead.
      if (swap === "duplicated") return { next: null, duplicated: true };
    } catch (err) {
      console.warn("SSF Charts: slide swap failed — falling back to a picture", err);
    }
  }

  try {
    // Bounded: a canvas that never fires `onload` would otherwise hang the
    // update forever, which is a worse outcome than the stall we are already
    // recovering from.
    const png = await Promise.race([
      rasterizeScene(scene),
      new Promise<never>((_, reject) => setTimeout(() => reject(new Error("rasterizing timed out")), 10_000)),
    ]);
    // UPDATE only while there is something left to update. Once layer 1 has
    // deleted the chart, an update resolves a dead shape id, decides the user
    // must have deleted the chart themselves, and does nothing — which is how
    // the floor of this ladder came to report "that chart is no longer on the
    // slide" about a chart it had just removed. With the slide and the position
    // both known, drawing it back is an insert, and the picture makes it one
    // shape that no live canvas can stall on.
    const next = wreckage
      ? await insertSceneIntoSlide(scene, {
          ...opts,
          pictureBase64: png,
          slideId: wreckage.slideId,
          left: wreckage.at.left,
          top: wreckage.at.top,
        })
      : await updateChartInSlide(scene, target, { ...opts, pictureBase64: png });
    return { next, picture: true, recovered: !!wreckage };
  } catch (err) {
    // Nothing left to try. Surface the host's own words rather than a
    // rasterizer error the user can do nothing about.
    console.warn("SSF Charts: picture fallback failed too", err);
    throw stall;
  }
}

/**
 * Insert the demo deck as ONE generated .pptx instead of drawing it shape by
 * shape — see `src/render/pptx-deck.ts` for why that is worth doing.
 *
 * Returns null when the fast path did not run and it is SAFE to fall back to
 * the shape-by-shape path: the host has no `insertSlidesFromBase64`, the file
 * could not be built, or the call failed with nothing landed. It returns a
 * message — never null — once any slide has landed, because falling back after
 * a partial insert would draw the whole deck a second time on top of it.
 */
async function insertDemoDeckAsFile(items: { scene: Scene; title: string; configJson?: string }[]): Promise<{
  text: string;
  status: "ok" | "err";
  added: number;
  totalMs: number;
  verified: VerifyResult;
  run: string;
} | null> {
  const t0 = Date.now();
  const before = await slideCount();
  // Identity for this insert, carried on every slide's slot tag. Without it the
  // repair pass below cannot tell these slides from the ones an earlier insert
  // left in the same deck, and "cannot tell" ends in a delete.
  const run = newRunId();
  let built: { base64: string; shapesPerSlide: number[] };
  try {
    note("Building the deck…", "busy");
    built = await buildDeckBase64(
      items.map((it, i) => ({ scene: it.scene, title: it.title, configJson: it.configJson, slot: i, run })),
      // Build the file at the DESTINATION's slide size. A generated deck that
      // declares a different size is one PowerPoint rescales on insert, which
      // moves every chart on every slide — silently, and on every 4:3 deck.
      await slideSize(),
    );
  } catch (err) {
    console.warn("SSF Charts: could not build the deck file — falling back to shapes", err);
    return null;
  }
  let added: number;
  try {
    note("Handing the deck to PowerPoint…", "busy");
    setProgress("busy");
    added = await insertSlidesFromPptx(built.base64, items.length);
  } catch (err) {
    // A throw is not proof that nothing landed — the same lesson the
    // shape path learned the hard way. Measure before deciding.
    console.warn("SSF Charts: one-shot deck insert failed", err);
    // …and a failed MEASUREMENT is not proof either. This runs on a host that
    // has just failed once, milliseconds ago; the re-read can fail with it.
    // Treating that as "nothing landed" returns null, and the caller then
    // draws the entire deck a second time on top of however many slides did
    // arrive — the exact double-insert every other guard on this path exists
    // to prevent. When we cannot tell, we do not guess: say so and stop.
    let after: number | undefined;
    try {
      after = await slideCount();
    } catch (readErr) {
      console.warn("SSF Charts: could not measure the deck after the failed insert", readErr);
      return {
        text: "PowerPoint would not take the deck, and would not say how much of it landed. Check the deck before inserting again — running it now could add the slides twice.",
        status: "err",
        added: 0,
        totalMs: Date.now() - t0,
        verified: { kind: "error", why: errorText(readErr) },
        run,
      };
    }
    added = after - before;
    if (added <= 0) return null;
  }
  if (added <= 0) return null;

  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  const expected: ExpectedItem[] = items.map((it, i) => ({
    slot: i,
    title: it.title,
    // What the FILE renderer drew, not what the Office.js one would have.
    // They disagree by design: a pie is one custGeom wedge here and a
    // sixteen-triangle fan there, so `estimateOfficeShapes` measured five
    // perfect charts as wreckage on the first real run.
    shapes: built.shapesPerSlide[i] ?? estimateOfficeShapes(it.scene),
    chart: !!it.configJson,
    // The generator writes the tag straight into the .pptx — there is no sync
    // to drop — so on this path the run always knows the tag is there. That is
    // what makes an unreadable slide safe to leave alone here, and unsafe to
    // leave alone on the shape path. See ExpectedItem.wroteTag.
    wroteTag: !!it.configJson,
  }));
  // Verify against the deck rather than trusting the count: the file carried
  // its own grouping and tags, so anything missing here is the host's doing.
  // Located by slot tag, not by position — where the host puts the slides is
  // its business, and on the first real run it put them at the FRONT.
  const verified = await repairDeckSpan(expected, (slot) => items[slot]?.configJson, run);
  const outcome = verified.kind === "ok" ? verified.outcome : undefined;

  let text = outcome
    ? `Inserted as one file in ${secs}s — ${describeReconcile(outcome.plan)}.`
    : `Inserted ${added} of ${items.length} slides as one file in ${secs}s.`;
  if (verified.kind !== "ok") text += ` (Not verified: ${verified.why}.)`;
  // A slide the readback could not see puts its item in the `lost` column for
  // want of evidence. Saying so is the difference between "the host dropped
  // your chart" and "we could not look" — and the first reading is what makes
  // someone insert the deck again.
  if (outcome?.unread) text += ` (Could not read ${outcome.unread} slide(s) — anything on them counts as lost here.)`;
  if (added < items.length) text += ` ⚠ the host took ${added} of ${items.length} slides.`;
  const clean = added >= items.length && !!outcome && outcome.plan.summary.lost === 0;
  return { text, status: clean ? "ok" : "err", added, totalMs: Date.now() - t0, verified, run };
}

/**
 * Shapes a web host may be asked to draw in one run before the add-in stops
 * asking. Every 12-item run in this project's history (~400 shapes) survived;
 * every 37-item one (~1850) took the whole client down. 600 sits between them,
 * nearer the side that lived.
 *
 * Only the web needs a number at all. Microsoft's documented runtime limits —
 * CPU, memory, four-crashes-per-session, five-seconds-unresponsive — are
 * scoped to Windows and Mac and explicitly NOT to a browser, so on the web
 * nothing throttles a runaway add-in: the tab dies and takes the session with
 * it. Desktop has that safety net and does not need this one.
 */
const WEB_SHAPE_BUDGET = 600;

/**
 * Rasterize, or give up. A canvas that never fires `onload` would otherwise
 * hang a run at the exact moment it is trying to rescue one.
 */
function boundedRaster(scene: Scene): Promise<string | undefined> {
  return Promise.race([
    rasterizeScene(scene),
    new Promise<undefined>((resolve) => setTimeout(() => resolve(undefined), 8_000)),
  ]).catch(() => undefined);
}

/**
 * How many of a round's new slides get a picture.
 *
 * A cap rather than a guess about size: at 480px a slide PNG is on the order of
 * tens of kilobytes, and a round that added forty slides would otherwise turn a
 * readable JSON into something nobody opens. Twelve covers every round this
 * project has produced, and `slideShots` reports the ones it skipped rather than
 * dropping them, so a bigger round says so instead of looking smaller.
 */
const MAX_SHOTS = 12;

/**
 * The cap when **Picture every slide** is ticked — high enough not to bind.
 *
 * The default twelve is right for an ordinary round and wrong for one specific
 * question, which the 2026-08-09 round asked and could not answer: 29 of that
 * deck's 36 slides read back with zero shapes, and only 12 had a picture to
 * corroborate it. So 12 were confirmed blank on two witnesses and 17 were
 * unknowable — and "unknowable" is the answer the cap produced, not the host.
 *
 * A readback of zero is one witness, and this host answers a shape collection
 * short without throwing (`shapes-items-count-honest`), so one witness is not
 * enough to call a slide empty. The picture is the second. Ticking the box
 * spends the extra time and the extra megabytes to get it for every slide.
 *
 * **Now the default**, at the owner's call, and the rounds since have earned it:
 * the two-witness count is the only thing that has actually settled an
 * empty-slide question, and it was the capped rounds that could not. Shipped
 * opt-in first because the pictures are the heaviest call the add-in makes and
 * they run at the END of a round, on a host that has just been through the
 * self-test — the moment it is least able to take more. That cost is real and
 * has not changed; what changed is which side of it is worth defaulting to.
 * Untick the box for a struggling host or a smaller file. `slideShots` still
 * reports anything it skips, so even this cap can bind and say so.
 */
const MAX_SHOTS_ALL = 200;

/**
 * What the round left in the deck, in the sentence the owner reads at the end.
 *
 * A round leaves its slides behind ON PURPOSE — the scenarios' working slides
 * are the evidence, and `docs/REGRESSION.md` is written around a deck someone
 * can open. What was missing is that the pane never SAID so. The 2026-08-09
 * evening round added 43 slides of which 36 came back empty, said "Saved as one
 * file", and left the owner to discover a 44-slide deck by opening it.
 *
 * This repo has already paid for that once at the other end of the same run:
 * the host probe left 21 blank slides in a deck, reported nothing, and the only
 * way to find out was to look. That got a row in the answer sheet. The
 * self-test never got the equivalent.
 *
 * The count is what the READBACK said, and the wording says so rather than
 * claiming the slides are empty: this host reports a shape collection short
 * without throwing (`shapes-items-count-honest`), so a zero here is one witness.
 * `npm run triage` is where that gets cross-examined against the host's own
 * pictures — deliberately not duplicated into the pane, because two copies of
 * one claim is how the two stop agreeing.
 */
function describeLitter(deck: RunLogFile["deck"]): string {
  if (!deck?.newSlides?.length) return "";
  const added = new Set(deck.newSlides);
  // The larger of the two readings, so a partial listing is not called empty —
  // `count` is the host's own number, `shapes` is what the scan managed to list.
  const empty = (deck.inventory ?? []).filter(
    (s) => added.has(s.slideId) && Math.max(s.count ?? 0, s.shapes?.length ?? 0) === 0,
  ).length;
  const n = added.size;
  return (
    ` It left ${n} slide${n === 1 ? "" : "s"} in this deck${empty ? `, ${empty} of which read back empty` : ""} —` +
    " the scenarios' own working slides, kept so the deck is evidence." +
    " Press Clean up the last round when you have finished with them."
  );
}

/**
 * HOW LONG THE TAIL GETS BEFORE IT IS ABANDONED.
 *
 * `collectDeckEvidence` has always SAID it is best-effort — "a host too far gone
 * to describe its own deck still gets the verdicts out" — and nothing enforced
 * it. A failure was survivable because the function catches; a HANG was not,
 * because there was no bound, and a hang is what a dying host actually does.
 *
 * What that cost: the crash band is 441-572s, 9 builds of 13 die in this tail,
 * and 33 of 49 crash records hold a complete fourteen-scenario round that was
 * never filed. `readChartsPage` bounds its own sync at `READBACK_TIMEOUT_MS` —
 * ninety seconds — and `PowerPoint.run` need not settle at all once the host is
 * going, so the tail could sit there for minutes with every verdict already in
 * hand.
 *
 * 45 SECONDS, and the archive says that is enormously generous:
 *
 *     the whole tail, scanning -> done    31 completions   p50 5.6s   max 7.96s
 *     the deck scan alone               6,785 scans        p50 1.3s   max 22.1s
 *
 * Zero of either over 30s, ever. So this is 5.6x the worst complete tail on
 * record and 2x the worst single scan inside it — chosen to be impossible to hit
 * by a host that is merely slow, which is the only way a bound like this can be
 * wrong. (The tail has only 31 completions because per-phase tracing landed on
 * 2026-08-27; the scan figure is the whole archive.)
 *
 * WHAT IT CANNOT DO, said here so nobody credits it with more. This bound
 * catches ONE thing: a `PowerPoint.run` that never settles while every verdict
 * is already in hand. It does not catch the crash. Of the 77 crash reports on
 * file, 10 carry an ErrorName and it is the same one every time —
 * `errorLocalChangeLostSingleUser`, PowerPoint's SERVER-side lost edit — and the
 * other 67 carry none at all. A lost edit is not a hang, so a client-side
 * deadline has nothing to wait out: the dialog is already up and the round is
 * already over. Reading this bound as crash protection would be reading a
 * timeout as a cure for the one failure it cannot reach.
 */
export const DECK_EVIDENCE_TIMEOUT_DEFAULT_MS = 45_000;
let DECK_EVIDENCE_TIMEOUT_MS = DECK_EVIDENCE_TIMEOUT_DEFAULT_MS;

/**
 * Test-only, matching `_setSlideSizeTimeoutForTest`.
 *
 * The DEFAULT is exported and asserted, which is one step more than the sibling
 * takes, because the failure modes are not alike. A slide-size budget that
 * drifts high only wastes time; this one drifts back into the unbounded hang it
 * exists to prevent — and every test of it overrides the value, so nothing else
 * would ever notice.
 */
export function _setDeckEvidenceTimeoutForTest(ms: number): void {
  DECK_EVIDENCE_TIMEOUT_MS = ms;
}

/**
 * What landed on the slides — the two uploads a person has been making by hand.
 *
 * Best-effort by construction, and that is deliberate: this runs at the END of a
 * round, on a host that has just been through the self-test and may well be the
 * reason the round is worth reading. A failure here must cost the pictures and
 * nothing else — the verdicts are already in `lastRunLog`, and losing them to a
 * diagnostic's own tail would be the worst trade in the file.
 *
 * BOUNDED SINCE 2026-08-29, which is what makes the paragraph above true rather
 * than merely intended. On timeout this returns `undefined` — the same path a
 * throw already took — so the round finishes, the log is offered, and the round
 * file simply carries no `deck`. An absent `deck` is loud: `poolGroupingOutcome`
 * and `poolFullestSlide` both read it, and the trace names the abandonment.
 */
async function collectDeckEvidence(idsBefore: string[] | undefined): Promise<RunLogFile["deck"] | undefined> {
  let abandon: ReturnType<typeof setTimeout> | undefined;
  const started = Date.now();
  const bounded = await Promise.race([
    collectDeckEvidenceUnbounded(idsBefore),
    new Promise<"timeout">((resolve) => {
      abandon = setTimeout(() => resolve("timeout"), DECK_EVIDENCE_TIMEOUT_MS);
    }),
  ]);
  clearTimeout(abandon);
  if (bounded !== "timeout") return bounded;
  // NAMED, never silent. A round that files without deck evidence must say why,
  // or the next reader takes an absent inventory for an empty deck — the same
  // "unknown is not a no" mistake this repo has paid for three times.
  trace("pane", "gave up collecting deck evidence", {
    afterMs: Date.now() - started,
    budget: DECK_EVIDENCE_TIMEOUT_MS,
  });
  return undefined;
}

async function collectDeckEvidenceUnbounded(idsBefore: string[] | undefined): Promise<RunLogFile["deck"] | undefined> {
  try {
    // TRACED PER PHASE, because a crash in here has been blaming the wrong step.
    //
    // Four consecutive crashed runs on 2026-08-27 record the same last step,
    // `re-asked what the empty deck could not answer`, and none of them died
    // there — that is simply the last line anything WROTE. What runs next is
    // this function, which scans the whole deck, lists its slide ids and then
    // RENDERS A PICTURE OF EVERY NEW SLIDE, and until now it traced only in its
    // catch. So a host that fell over mid-screenshot filed a crash pointing at
    // an innocent probe re-ask, and `WHERE THE HOST DIED` pooled four of them
    // under that name.
    //
    // Cheap: three lines on a path that already takes seconds and runs once per
    // round, after every verdict is in.
    /**
     * THE SYNC COUNT GOES IN THE LINE THE ROUND DIES ON.
     *
     * 41 of the last 60 4:3 crashes have this trace line as their final entry —
     * the host never reaches the one below. `lastRunLog` is assembled after the
     * scan, so a crashed round's `syncs` field never exists; the trace does,
     * because `traceLog` is a local array read and survives a dead host.
     *
     * That makes this the only place a crashed round can report how much it had
     * asked of the save channel before it went, which is the whole question
     * office-js#6329 raises. Putting it on the same line as the scan
     * announcement rather than a line of its own keeps the crash-log tail
     * comparable with the 160 records already on file.
     *
     * IT WORKED, AND THE FIRST ANSWER CAME ON 2026-09-09. Rounds that died were
     * syncing about 15% faster than rounds that lived — 2.89 against 2.52 per
     * second at 4:3, 2.81 against 2.45 at 16:9, pooled p≈0.001 — with the arm
     * split done precisely because the 4:3 crash rate could otherwise have been
     * the whole effect. Correlation only, and the direction is not established,
     * but it is the sign #6329 predicts and it is the first evidence this
     * archive has produced about it. See `docs/BACKLOG.md`.
     *
     * `syncs` ON THIS LINE IS THE ROUND'S RUNNING TOTAL. The same key on
     * `updated only the shapes that changed` is that ONE chart's syncs, and a
     * query that does not scope to the message mixes the two into a bimodal
     * nonsense. Both sites carry this warning.
     */
    trace("pane", "collecting deck evidence — scanning", {
      knownBefore: idsBefore?.length ?? null,
      syncs: syncsSoFar(),
    });
    const scan = await listChartsInDeck({ withInventory: true });
    trace("pane", "collecting deck evidence — listing slide ids", { charts: scan.charts?.length ?? null });
    const idsAfter = await deckSlideIds();
    // Only slides that were not there before. Without the diff a picture of a
    // forty-slide deck is mostly slides nobody touched — and the id list is the
    // stronger question about a deck than any handle, which is why the diff is
    // taken from ids rather than from counts.
    const known = new Set(idsBefore ?? []);
    // AN UNREADABLE BEFORE-LIST IS NOT AN EMPTY DECK. This fell back to
    // `idsAfter`, which says every slide the deck holds was added by this round
    // — so a round that added four slides to the owner's forty-slide deck
    // reported forty, and `describeLitter` told them it had left forty behind,
    // twelve of which "read back empty". The twelve were their own blank slides.
    // Everything downstream inherits it: `triage.mjs` builds its added-set from
    // this field.
    //
    // The same mistake as reading a host's silence as a "no", which this repo
    // has now paid for in three separate places. Unknown is its own answer:
    // `newSlides` stays empty and `beforeUnknown` says why, so a reader can tell
    // "this round added nothing" from "nobody knows what it added".
    const beforeUnknown = !idsBefore || !idsAfter;
    const newSlides = beforeUnknown ? [] : idsAfter.filter((id) => !known.has(id));
    // Read at collection time, not at boot: the box is ticked for the round
    // about to be read, and a value captured when the pane loaded would be the
    // one from before the owner ticked it.
    const shotAll = ($("demo-shot-all") as HTMLInputElement | null)?.checked ?? false;
    // THE EXPENSIVE ONE — it renders an image per new slide — and it WAS the
    // prime suspect for those four crashes. It is not the culprit, and the
    // per-phase tracing added above is what settled it. Over 2026-08-29's four
    // builds, counted one vote per build the way this archive counts crashes:
    //
    //     e9222a8   4 of 4 crashes        557b10e   1 of 1
    //     3495fb8   5 of 5                4275306   4 of 6
    //
    // every one died at `scanning`, the FIRST phase, before this line was ever
    // reached. That takes the scan to 9 builds against 4 for everything else.
    // The suspect is `listChartsInDeck({ withInventory: true })` above.
    //
    // Still traced with the COUNT it is about to attempt: it remains the
    // expensive call, and a crash that DOES land here should say how many it was
    // asked for rather than leave the number to be inferred from the deck.
    trace("pane", "collecting deck evidence — shooting slides", {
      slides: newSlides.length,
      max: shotAll ? MAX_SHOTS_ALL : MAX_SHOTS,
    });
    const shots = await slideShots(newSlides, { max: shotAll ? MAX_SHOTS_ALL : MAX_SHOTS });
    trace("pane", "collecting deck evidence — done", { shots: shots.length });
    return {
      inventory: scan.inventory ?? [],
      ...(scanIsComplete(scan) ? {} : { gap: scanGap(scan) }),
      newSlides,
      // Carried so a reader can tell an empty `newSlides` that MEANS nothing was
      // added from one that means the question could not be asked. Omitted on
      // the ordinary path, so an existing round file reads exactly as it did.
      ...(beforeUnknown ? { beforeUnknown: true } : {}),
      shots,
    };
  } catch (err) {
    trace("pane", "could not collect deck evidence", { error: errorText(err) });
    return undefined;
  }
}

/** True when the add-in is running inside PowerPoint on the web. */
function isWebHost(): boolean {
  try {
    return Office.context?.diagnostics?.platform === Office.PlatformType.OfficeOnline;
  } catch {
    return false;
  }
}

/** Save an object as a JSON file, via the same Blob dance as Download SVG. */
/**
 * Hand a file to the browser, and say whether the attempt even got that far.
 *
 * Three things were wrong with the two-line version, and a real round on
 * 2026-08-06 lost a whole evidence file to them.
 *
 * 1. **It revoked the object URL in the same tick as the click.** A download is
 *    asynchronous; revoking its source synchronously can cancel it before it
 *    starts. The revoke now waits, and the URL is released on a timer rather
 *    than never — a leaked blob in a pane that stays open for a session is the
 *    lesser of the two.
 * 2. **The anchor was never in the document.** Chromium tolerates a detached
 *    anchor and other engines do not, and a task pane is whatever the host
 *    embeds.
 * 3. **It reported nothing.** Callers said "Saved as one file" on the strength
 *    of having called this. It returns false when the browser refused outright,
 *    which is the case worth acting on: a blocked download that THROWS is now
 *    distinguishable from one that worked.
 *
 * A `false` is proof of failure; a `true` is not proof of success — a frame can
 * still swallow a download silently, and nothing in the DOM reports that. So no
 * caller may treat `true` as "the user has the file"; that is what the explicit
 * save button is for.
 */
function downloadJson(name: string, payload: unknown): boolean {
  let url: string | undefined;
  try {
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.style.display = "none";
    document.body.append(a);
    a.click();
    a.remove();
    return true;
  } catch (err) {
    trace("error", "the browser would not save the file", { name, error: String(err) });
    return false;
  } finally {
    // Long enough for the download to have been picked up, and not tied to the
    // click that starts it.
    const held = url;
    if (held) setTimeout(() => URL.revokeObjectURL(held), 60_000);
  }
}

/**
 * The live transcript panel: the step list, its copy button and its clear
 * button.
 *
 * Lifted whole out of `wireInsert`, which was 1113 lines and is the reason the
 * line formatter inside it went untested for months while it silently dropped
 * every array-valued payload. A panel that owns one DOM node, one bounded
 * buffer and three handlers is a thing that can be read in one screen; the same
 * code as a middle third of a thousand-line function is not.
 *
 * Returns `revealSteps`, which the run buttons call once as a run starts —
 * markup order is not position, and a log you must scroll to before you can
 * photograph it is a log a dying tab takes with it.
 */
function wireStepsPanel(): { revealSteps: () => void } {
  // The live transcript.
  //
  // Wired to the trace stream rather than to the run's own reporting, so it
  // costs no new instrumentation and shows exactly what the log would have —
  // including the `at=<phase>` on any error. It exists because the log does
  // NOT survive the failures worth explaining: it becomes downloadable only
  // when a run ends, and two real-host rounds have now been lost to a run
  // that never ended (a wedge at 1819s) and one PowerPoint killed outright
  // ("Sorry, we ran into a problem", at 108s). What is on screen survives
  // both, and can be copied or photographed before the reload.
  const steps = $("demo-steps");
  /** Newest FIRST, capped — a pane is not a heap, and the head is what is read. */
  const STEP_LINES = 300;
  const lines: string[] = [];
  // NEWEST FIRST, and that ordering is the whole point rather than a
  // preference.
  //
  // The first version appended and auto-scrolled to the bottom, which reads
  // better while a run is healthy and fails at the only moment this list
  // exists for. When PowerPoint dies, what you have is whatever pixels were
  // on screen — no scrolling, no clicking, often a modal dialog over half the
  // pane. An append-and-follow list puts the last thing that happened at the
  // bottom of a small scrolled box, which is exactly where it cannot be
  // relied on to be visible. Prepending puts it at a FIXED position, one line
  // below the header, and needs no scroll to have worked.
  //
  // It also removes the follow-the-tail logic entirely: there is no tail to
  // follow, so there is no "unless the user scrolled" case to get wrong.
  const paintSteps = (): void => {
    steps.textContent = lines.join("\n");
  };
  /**
   * Put the step list where it can be seen, once, as a run starts.
   *
   * Newest-first fixes WHERE in the box the last line is; it does nothing
   * about whether the box itself is on screen. Two things now answer that.
   * The list is the FIRST thing in the Testing section, above the buttons
   * that start a run — a real-host round crashed with the log still under
   * nine controls and a paragraph, which is a log you must scroll to before
   * you can photograph it. And this scrolls the panel to it anyway, because
   * markup order is not position: the Automation tab scrolls, and a run
   * started after reading the JSON section below would otherwise begin with
   * the box off the top. Once per run, on the click that starts it — never
   * while the run is going, because a pane that moves under the cursor
   * mid-run is its own problem.
   */
  const revealSteps = (): void => {
    try {
      steps.scrollIntoView({ block: "nearest" });
    } catch {
      /* an older host without scrollIntoView options — the list still fills */
    }
  };
  onTrace((e) => {
    // The data payload matters as much as the message for the lines that
    // locate a failure — `error`, `name`, `detail` are where the phase and
    // the verdict live. `formatTraceLine` is where that is decided, and it
    // lives in `trace.ts` rather than here because it used to live here: an
    // inline formatter in a DOM closure is a formatter nothing can test, and
    // for months it silently dropped every array-valued payload — including
    // the two timing series `degradation curves` exists to produce.
    const line = formatTraceLine(e);
    lines.unshift(line);
    // Drop the OLDEST, which is now the end of the array.
    if (lines.length > STEP_LINES) lines.length = STEP_LINES;
    paintSteps();
    // The same line, to storage, where it outlives this JavaScript context.
    // One formatter for both, so the file and the screen can never describe
    // the same run differently — and oldest-first there, because a file is
    // read from the top while a crashed screen is read from where it froze.
    recordCrashStep(line);
    // Anything traced is the host or the run still moving, which is what the
    // elapsed readout needs to tell "slow" from "gone".
    noteHostActivity();
  });
  // The last synchronous moment the pane gets. `pagehide` fires on the tab
  // close that ended the 1819-second run, so it buys back the final debounce
  // window — the part of a dying run that nothing else can reach.
  window.addEventListener("pagehide", flushCrashLog);
  $("demo-steps-copy").addEventListener("click", () => {
    if (!lines.length) {
      note("No steps to copy yet.", "err");
      return;
    }
    // Labelled, because the order is the opposite of what a log usually is
    // and a reader who assumes otherwise reads the run backwards.
    const text = [`SSF Charts steps — NEWEST FIRST (${lines.length} lines)`, ...lines].join("\n");
    // Two ways, because the first one does not work where this runs.
    //
    // `navigator.clipboard` needs a secure context, a user gesture AND the
    // `clipboard-write` permission — and an Office task pane is a nested
    // cross-origin iframe that is routinely refused it. Observed, on the run
    // this button exists for: "The browser would not give us the clipboard".
    // Telling the user to select the text by hand is not a fallback, it is
    // an apology, and it arrives at the moment they can least afford one.
    //
    // So: select the transcript and run the legacy copy command, which is
    // permitted from a user gesture in an iframe. It is deprecated and it
    // works. If even that is refused the text is at least now SELECTED, so
    // Ctrl+C finishes the job.
    const selectAndCopy = (): boolean => {
      try {
        const range = document.createRange();
        range.selectNodeContents(steps);
        const sel = window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(range);
        return document.execCommand("copy");
      } catch {
        return false;
      }
    };
    const done = (n: number) => note(`Copied ${n} step(s).`, "ok");
    void Promise.resolve()
      .then(() => navigator.clipboard?.writeText(text))
      .then(() => done(lines.length))
      .catch(() => {
        if (selectAndCopy()) done(lines.length);
        else note("The browser refused the clipboard — the steps are selected, press Ctrl+C.", "err");
      });
  });
  $("demo-steps-clear").addEventListener("click", () => {
    lines.length = 0;
    paintSteps();
  });
  return { revealSteps };
}

function wireInsert() {
  const insertBtn = $("insert") as HTMLButtonElement;
  const insertNewBtn = $("insert-new") as HTMLButtonElement;
  const loadBtn = $("load-selection") as HTMLButtonElement;
  const agendaBtn = $("agenda-insert") as HTMLButtonElement;
  if (isPowerPointHost()) {
    note("");
    insertBtn.disabled = false;
    loadBtn.disabled = false;
    /**
     * Run a host action with a busy note, and lock out re-entry while it runs.
     *
     * Both buttons matter. The one that was CLICKED has to go dead or a slow
     * action invites a second click that queues the whole job again — "Insert
     * demo deck" is 35 slides and ~1,700 shapes, and it stayed live throughout.
     * The primary Insert button has to go dead too, since it acts on the same
     * deck. Disabling only the primary was the worst of both: the clicked
     * button re-entered freely, while Insert went dead WITHOUT looking it (the
     * CSS greys `.el-insert:disabled`, not the primary), so a stuck action read
     * as "Insert does nothing" rather than "Insert is busy".
     *
     * The clicked button comes from the event, so no call site can forget it.
     */
    /**
     * Disable a button and mean it, from inside a guarded action.
     *
     * `guard()` captures which buttons it disabled BEFORE running the handler and
     * re-enables exactly those afterwards. A handler that disables its OWN button
     * during the action is therefore undone by the very next line of the finally —
     * which is how "Clean up the last round" came back to life after emptying its
     * list, so a second press reported a clean sweep of nothing in green.
     */
    function keepDisabled(btn: HTMLButtonElement): void {
      btn.disabled = true;
      btn.dataset.keepDisabled = "1";
    }

    /**
     * HOW MANY GUARDED ACTIONS ARE IN FLIGHT, and why one number fixes two bugs.
     *
     * `guard` owns three pieces of state that are MODULE-WIDE while it is not:
     * the stop flag, the elapsed interval, and the Stop button. It leaves about
     * thirteen controls enabled during a slow action, so two actions sit inside
     * it at once and the second one's bookkeeping runs over the first's.
     *
     * THE STOP THE USER PRESSED WAS CLEARED AND THE CANCELLED WORK LANDED.
     * Start an insert this code prices at "the better part of a minute", press
     * Stop — the button disables and reads "Stopping…" — then touch any other
     * live control. Its `guard` calls `resetStop()` on the way in, the insert
     * reaches its next batch, asks `isStopRequested()`, is told NO, and writes
     * its chart onto the slide. Reproduced against this module with the real
     * stop functions: `stopSeen: [false] landed: 1`, against `[true] landed: 0`
     * when nothing else is clicked.
     *
     * AND THE FINALLY LIED ABOUT THE OUTCOME. It posted "Done." in green, hid
     * Stop and cleared the elapsed interval while the first action was still
     * running — and that interval is the only carrier of the `SILENT_RUN_MS`
     * watchdog, so the "PowerPoint died and the pane did not" check was disarmed
     * for the rest of the action. Measured: elapsed reads "" 2.3s in, against
     * "2s" when the second action is not clicked.
     *
     * So the OUTERMOST action owns all three and a nested one leaves them alone.
     * NOT per-action stop tokens: `isStopRequested()` is asked from eleven
     * places inside `powerpoint.ts`, and threading a token through them is a far
     * larger change than this defect needs. One consequence worth naming — a
     * pending stop DOES cancel a concurrently started action, because the flag
     * stays global. That is the safe direction: Stop stops work, and nothing
     * completes against the user's wish.
     */
    let actionsInFlight = 0;

    const guard = (fn: () => Promise<void>) =>
      async function (this: unknown, ev?: Event) {
        const clicked = ev?.currentTarget as HTMLButtonElement | undefined;
        // Every host action in the pane comes through here, so this is the one
        // place that can log them all — inserts, updates, agenda, elements,
        // same-scale, explode, the demo deck. Normal use, not just the harness.
        const action = clicked?.id ?? "action";
        const startedAt = Date.now();
        trace("pane", "action started", { action });
        const lock = [insertBtn, clicked].filter((b): b is HTMLButtonElement => !!b && !b.disabled);
        for (const b of lock) b.disabled = true;
        // Before the first await: background listeners must see the host as
        // busy from the instant the action begins, not from its first sync.
        hostWorkStarted();
        note("Working…", "busy");
        // Mark AFTER the busy note: "Working…" is not a settlement, and the mark
        // has to sit at the boundary of this action so only what `fn()` itself
        // posts counts. See `settledNotes` — a phase note is busy by definition,
        // so an insert that ends on one still needs "Done." to close it out.
        const settledAt = settledNotes;
        setProgress("busy");
        // OWNS the stop flag, the elapsed timer and the Stop button — or does
        // not, because another action is already inside `guard`. See
        // `actionsInFlight`.
        const owns = actionsInFlight === 0;
        actionsInFlight++;
        if (owns) {
          startElapsed();
          // Clear any stop left by the PREVIOUS action before offering a new
          // one. A stop that survived into the next action would cancel it at
          // its first batch, and the user would never learn why. PREVIOUS, not
          // concurrent: doing this while another action is running is what
          // threw away the stop the user had just pressed.
          resetStop();
          showStop(true);
        }
        try {
          await fn();
          trace("pane", "action finished", { action, ms: Date.now() - startedAt });
          // A stopped action that ended tidily still has to SAY it stopped —
          // "Done." over work the user cancelled is the pane claiming credit
          // for something it did not do. This covers the paths that stop by
          // returning early (a deck run between items, Same Scale between
          // charts) rather than by throwing.
          if (isStopRequested()) {
            note("Stopped — anything already drawn was kept.", "err");
          } else if (settledNotes === settledAt && owns) {
            // `owns`, because "Done." is a claim about the PANE and not about
            // this handler: posted by a nested action it says the work is
            // finished while an insert is still drawing. The outer action will
            // post its own when it actually ends.
            note("Done.", "ok");
          }
        } catch (err) {
          // The user's own stop is not a failure, and reporting it as one
          // ("Failed: Stopped.") reads like the add-in broke.
          if (isStopped(err)) {
            trace("pane", "action stopped", { action, ms: Date.now() - startedAt });
            note("Stopped — anything already drawn was kept.", "err");
          } else {
            /**
             * THE STACK, FOR THE ONE CLASS OF FAILURE THIS TRACE CANNOT PLACE.
             *
             * A host refusal names itself — `errorText` digs the Office.js code
             * and `debugInfo` out, and the trace line before it says which call
             * was in flight. A fault in OUR OWN code says only what it tripped
             * over, and on 2026-09-02 that was `Cannot read properties of
             * undefined (reading 'answer')`: a whole round abandoned at 83s,
             * the first occurrence in 161 crash records, on a host so sick that
             * `deckSlides` was unreadable from the first line and the deck
             * refused to list its slides four times before the throw.
             *
             * It cost an afternoon and was not found. `withTimeout` rejects
             * rather than resolving empty, `withProbeContext` always returns
             * `fn(...)` or throws, and the probe that was in flight returns on
             * every path — so reading the code ruled out the three obvious
             * sites and named none. There are 49 `.answer` reads in
             * `host-probe.ts` and no way to tell which one it was.
             *
             * One frame would have. It is recorded only for a native `Error`
             * with a stack, and only the first few frames: a RichApi.Error's
             * stack is Office.js's own plumbing and says nothing, and the whole
             * record is capped at what a crash log can carry.
             *
             * NOT A FIX. The fault is still there and still unlocated; this is
             * so the next occurrence is the last one that has to be guessed at.
             */
            const stack =
              err instanceof Error && typeof err.stack === "string"
                ? err.stack.split("\n").slice(0, 6).join(" | ").slice(0, 600)
                : null;
            trace("pane", "action failed", {
              action,
              ms: Date.now() - startedAt,
              error: errorText(err),
              ...(stack ? { stack } : {}),
            });
            // errorText, not err.message: a RichApi.Error's message is generic
            // ("An internal error has occurred") and the useful part is in code
            // and debugInfo, which String(err) throws away.
            note("Failed: {error}", "err", { error: errorText(err) });
          }
        } finally {
          actionsInFlight = Math.max(0, actionsInFlight - 1);
          /**
           * THE LAST ONE OUT TURNS OFF THE LIGHTS — not the one that turned them
           * on.
           *
           * Gating this on `owns` (taken at entry) is only right while actions
           * finish in the order they started, and they need not: the outer
           * action here is a slow insert and the nested one a fast element, but
           * an insert that returned FIRST would tidy up while the element was
           * still running — hiding Stop and stopping the watchdog clock under
           * it, which is this same defect with the two swapped. Reading the
           * count after the decrement is order-independent.
           */
          if (actionsInFlight === 0) {
            stopElapsed();
            showStop(false);
            resetStop();
          }
          // Only re-enable what this call disabled — never resurrect a button
          // some other state (no host, no selection) means to keep dead.
          //
          // `lock` is captured BEFORE `fn()`, so a button the HANDLER itself
          // disabled during the action was still in it and came back to life.
          // "Clean up the last round" is the case: it empties `tidyable` and
          // disables itself on purpose, because after a partial sweep there is
          // no longer a list it trusts — and this line undid that, so a second
          // press printed a GREEN "Cleaned up — 0 slide(s) removed." over slides
          // the host had just refused. `keepDisabled` is how a handler says it
          // means it.
          for (const b of lock) {
            if (b.dataset.keepDisabled) {
              delete b.dataset.keepDisabled;
              continue;
            }
            b.disabled = false;
          }
          // Last: this releases the selection watcher, and it must not be let
          // loose while the lines above are still touching the pane.
          hostWorkFinished();
        }
      };
    // Not through `guard()`: this button must stay live precisely WHILE a
    // guarded action is running, which is the one state guard() disables things
    // in. It queues no host work of its own — it sets a flag the render loops
    // read at their next batch boundary.
    statusStop?.addEventListener("click", () => {
      requestStop();
      // Say so immediately. The batch already handed to PowerPoint still has to
      // come back — up to BATCH_TIMEOUT_MS — and without this the pane looks
      // like it ignored the click for as long as that takes.
      if (statusStop) {
        statusStop.disabled = true;
        statusStop.textContent = t("Stopping…");
      }
    });
    insertBtn.addEventListener(
      "click",
      guard(() => doInsert(false)),
    );
    insertNewBtn.addEventListener(
      "click",
      guard(() => doInsert(true)),
    );
    loadBtn.addEventListener("click", guard(doLoadSelection));
    $("selection-banner-load").addEventListener("click", guard(doLoadSelection));
    // Explode needs the same host surface as "Edit selected chart" (it reads the
    // selection's tag), so it lives or dies with the same gate.
    const explodeBtn = $("explode") as HTMLButtonElement;
    explodeBtn.disabled = false;
    explodeBtn.addEventListener("click", guard(doExplode));
    watchSelection();
    const sameScaleBtn = $("same-scale") as HTMLButtonElement;
    sameScaleBtn.disabled = false;
    sameScaleBtn.addEventListener(
      "click",
      guard(() => doSameScale("deck")),
    );
    const sameScaleSelBtn = $("same-scale-sel") as HTMLButtonElement;
    sameScaleSelBtn.disabled = false;
    sameScaleSelBtn.addEventListener(
      "click",
      guard(() => doSameScale("selection")),
    );
    const batchBtn = $("json-insert-batch") as HTMLButtonElement;
    batchBtn.disabled = false;
    batchBtn.addEventListener(
      "click",
      guard(async () => {
        const parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
        const configs: ChartConfig[] = Array.isArray(parsed) ? parsed : [parsed];
        let degraded = 0;
        let rescued = 0;
        for (const c of configs) {
          const cfg = { ...DEFAULT_SIZE, ...c };
          const scene = buildChart(cfg);
          // Batch insert honours render:"image" per config — without this a
          // pasted image-mode array silently drew native shapes.
          const { png, warn } = await chartPicture(cfg, scene);
          // A warn WITH a png is the rescue succeeding, not failing — see the
          // same distinction in doSameScale.
          if (warn && !png) degraded++;
          else if (warn) rescued++;
          await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), pictureBase64: png });
        }
        note("Inserted {n} chart(s) on the current slide.", "ok", { n: configs.length });
        if (rescued) {
          note(
            '{r} of them were too dense for this host and went in as pictures — "Explode to native shapes" turns them back.',
            "ok",
            { r: rescued },
          );
        }
        if (degraded) {
          note("Inserted {n} chart(s), but {d} image chart(s) fell back to native shapes.", "err", {
            n: configs.length,
            d: degraded,
          });
        }
      }),
    );
    // Elements insert at a small default offset (they're compact shapes).
    for (const [id, scene] of [
      ["harvey-insert", harveyScene],
      ["check-insert", checkScene],
      ["flow-insert", flowScene],
      ["kpi-insert", kpiScene],
      [
        "table-insert",
        () => buildTableScene(state.sheet.cells, 480, { totalRow: ($("table-total") as HTMLInputElement).checked }),
      ],
    ] as const) {
      const btn = $(id) as HTMLButtonElement;
      btn.disabled = false;
      btn.addEventListener(
        "click",
        guard(async () => {
          const built = scene();
          /**
           * ELEMENTS GET THE SAME TREATMENT AS CHARTS, and they are where it
           * matters most: a Harvey ball is a WEDGE at every fraction between 1%
           * and 99% — 0% and 100% are ellipses and survive, so a host without
           * rotation loses exactly the informative range and draws an empty
           * ring — and a table `[up]`/`[down]` cell is an arrowhead that
           * vanishes while its text stays.
           *
           * This path calls `insertSceneIntoSlide` directly and never went near
           * `chartPicture`, so shipping the guard on the chart path alone left
           * the two element kinds that need it most dropping their glyphs in
           * silence.
           */
          const undrawable = await pictureForUndrawableMarks(built);
          await insertSceneIntoSlide(
            built,
            { left: 120, top: 160, ...(undrawable?.png ? { pictureBase64: undrawable.png } : {}) },
            phaseNote,
          );
          // AFTER the insert, like every other warning here: the insert's first
          // act is `onPhase("context")`, which overwrites a note posted before
          // it. See `chartPicture`'s own comment.
          // A picture is a SUCCESS — the element went in complete — so it settles as
          // "ok"; only the shapes-with-marks-missing case is an error.
          if (undrawable) note(undrawable.warn, undrawable.png ? "ok" : "err");
        }),
      );
    }
    agendaBtn.disabled = false;
    agendaBtn.addEventListener(
      "click",
      guard(async () => {
        const chapters = agendaChapters();
        if (!chapters.length) return;
        await insertAgendaSlides(chapters.map((_, i) => buildAgendaScene(chapters, { highlight: i })));
      }),
    );
    // Testing aid: one demo slide per chart kind + feature/element highlights.
    // A call we gave up on may still answer. Whatever it says is the only
    // evidence we get about a host that went quiet, so surface it even though
    // the action has already failed — the note is stale by then, but the
    // information is what a bug report needs.
    onLateSync((msg) => note("Host answered late — {message}", "err", { message: msg }));
    // Turning the fast path OFF on the web is a decision worth flagging. At
    // volume the shape-by-shape path there does not merely stall: the full
    // 37-item deck took the whole web client down five seconds in on
    // 2026-07-31 — "Sorry, we ran into a problem. Please try again." The
    // twelve-item subset survived, so this is a warning and not a block.
    const pathSelect = $("demo-path") as HTMLSelectElement | null;
    pathSelect?.addEventListener("change", () => {
      if (pathSelect.value !== "shapes" || !canInsertSlidesFromBase64()) return;
      if (isWebHost()) {
        note(
          "Heads up: the full deck drawn shape by shape has crashed PowerPoint on the web. The fast path handles it in seconds.",
          "err",
        );
      }
    });
    // Enabled by a run, not by the host: with nothing to save it would only
    // ever produce an empty file.
    // ON by default for now. Nothing in this project has been diagnosed from
    // anything but an after-the-fact artifact, and the runs that matter happen
    // on a host nobody can attach a debugger to. When the add-in stops being
    // validated against real hosts, uncheck it in taskpane.html and drop the
    // `checked` — the module, its call sites and this toggle all keep working,
    // so a future investigation is one click away rather than a re-implementation.
    // The live transcript — see `wireStepsPanel`.
    const { revealSteps } = wireStepsPanel();

    const traceToggle = $("demo-trace") as HTMLInputElement | null;
    if (traceToggle?.checked) {
      setTracing(true);
      traceEnvironment(typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev");
    }
    traceToggle?.addEventListener("change", () => {
      setTracing(traceToggle.checked);
      if (traceToggle.checked) {
        traceEnvironment(typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev");
        note("Verbose trace on — it rides along in the run log.", "ok");
      } else note("Verbose trace off.", "ok");
    });
    $("demo-log").addEventListener("click", () => {
      if (!lastRunLog) {
        note("No run to save yet — insert the demo deck first.", "err");
        return;
      }
      if (!downloadJson("ssf-charts-run-log.json", lastRunLog)) {
        note("The browser would not save the file. Copy the Live steps instead — they carry the same run.", "err");
        return;
      }
      // A button the user pressed is the strongest evidence this pane can have
      // that they now hold the run. It is also the recovery from an auto-save
      // that was blocked, so it has to clear the stored record — otherwise the
      // pane would keep offering back a run they have just saved.
      markCrashLogSaved();
      note("Run log saved.", "ok");
    });
    /**
     * Offer the last run that never reported finishing.
     *
     * Checked once, on the open that follows the crash — which is the only
     * moment anyone is looking for it, and the moment before the natural next
     * action (run it again) would otherwise bury it. Hidden entirely when
     * there is nothing to recover, so a healthy pane carries no wreckage.
     */
    /**
     * Ask this host the fixed question list and save what it says.
     *
     * The one diagnostic here that is not about a run at all. Everything else
     * in this panel reports what the ADD-IN did; this reports what the HOST is,
     * so the fake that every test in the repo stands on can finally be checked
     * against the thing it stands for. One click, no deck changes — it works on
     * a scratch slide and takes it back.
     */
    /**
     * ONE QUESTION, ANSWERED IN SECONDS.
     *
     * The round is the wrong instrument for settling a single "does this host do
     * X?" that a decision is waiting on — fourteen minutes, and the question has
     * to earn a slot in a fixed sheet. `grouped-child-by-id-from-slide` waited
     * 125 rounds for a slot and never got one.
     *
     * The alternative was worse: putting a speculative host call in the drawing
     * batch to find out, which is where loading an id on a creation handle
     * poisons it and costs the tag that makes a chart re-editable.
     */
    {
      const pick = $("experiment-pick") as HTMLSelectElement;
      for (const e of EXPERIMENTS) {
        const option = document.createElement("option");
        option.value = e.id;
        // The QUESTION in the list, not the id. The id is for the archive; a
        // person choosing one wants to read what it asks.
        option.textContent = e.asks;
        pick.append(option);
      }
      $("experiment-run").addEventListener(
        "click",
        guard(async () => {
          revealSteps();
          const chosen = pick.value || EXPERIMENTS[0]?.id;
          note(`Asking: ${EXPERIMENTS.find((e) => e.id === chosen)?.asks ?? chosen}`, "busy");
          const r = await runExperiment(chosen);
          // The DETAIL beside the word, always. The vocabulary will be wrong for
          // something eventually and the detail is what survives that.
          note(
            `${r.id} — ${r.answer}${r.detail ? `: ${r.detail}` : ""} (${r.ms}ms)`,
            r.answer === "yes" ? "ok" : "err",
          );
        }),
      );
    }

    $("demo-probe").addEventListener(
      "click",
      guard(async () => {
        revealSteps();
        note("Asking this PowerPoint what it actually does…", "busy");
        const sheet = await runHostProbes(
          describeHost(),
          typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev",
        );
        const saved = downloadJson("ssf-charts-host-answers.json", sheet);
        // The diff, here, now — rather than after a round trip.
        //
        // Every probe run so far has been "download it, send it, wait for
        // someone to run `host-diff`, hear back". Most of those establish
        // nothing new: the answers are the same as last time. The comparison
        // table is a plain object, so the pane can do it and say whether this
        // run is worth sending at all.
        note(
          saved
            ? describeHostSheet(sheet)
            : `The browser would not save the file. ${describeHostSheet(sheet)} Copy the Live steps — they carry the answers.`,
          !saved || sheetNeedsAttention(sheet) ? "err" : "ok",
        );
      }),
    );
    /**
     * One click, one file: the probe and the self-test, back to back.
     *
     * What it saves is round trips rather than seconds. A round used to be
     * three clicks producing three downloads, uploaded separately and joined at
     * the other end — and most probe runs establish nothing, so a good share of
     * that traffic was to learn that the answers had not changed.
     *
     * The DEMO DECK is deliberately not in here, and not for want of effort.
     * Its two halves have to run on different decks: the file half fills the
     * deck, and the shape half then draws onto that same larger deck, which is
     * the one configuration that has ended in PowerPoint's crash dialog every
     * time it has been tried. A button cannot open a fresh deck, so chaining
     * the demo in would bake in exactly the arrangement the runbook splits up.
     *
     * The probe goes FIRST because it is the cheap one. If the host is already
     * unwell, seventeen short questions say so in seconds, and they are still
     * in the bundle when the long half dies.
     */
    const roundBtn = $("demo-round") as HTMLButtonElement;
    roundBtn.disabled = false;
    roundBtn.addEventListener(
      "click",
      guard(async () => {
        revealSteps();
        const buildStamp = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
        const host = describeHost();
        lastRunLog = undefined;
        ($("demo-log") as HTMLButtonElement).disabled = true;
        beginCrashLog({ build: buildStamp, host, label: "the whole round" });
        // ZEROED WITH THE ROUND, beside the trace mark and for the same reason:
        // a count that carries over from the pane's earlier life belongs to no
        // round, and the pane is reused across rounds far more often than it is
        // reloaded. See `syncsSoFar`.
        resetSyncCount();
        const traceFrom = traceMark();
        // AFTER THE MARK, OR IT IS NOT IN THE ROUND. `traceEnvironment` has been
        // called at pane-wiring time since #228 and its output has reached
        // exactly ZERO of the 69 archived rounds — `traceLog(traceFrom)` slices
        // off everything traced before the mark, so the host, the platform, the
        // Office version, the slide size and (as of yesterday) the deck-style
        // replay were all written into a window no round file ever sees.
        //
        // Measured, not reasoned: `host`/`environment` appears in 0 of 69 round
        // files, and 0 of `slide size`, on an archive where every build contains
        // the emitting line. The pre-mark window is real and large — round 092's
        // first entry is at ms 48245 with `dropped: 0`, so 48 seconds of trace
        // was sliced away.
        //
        // The replay had a second, earlier blocker of its own: `wireInsert()`
        // runs synchronously BEFORE `void readDeckStyle()` in `Office.onReady`,
        // so at wiring time no read has happened and `deckStyleVerdict` is null.
        // Here it has, which is the other half of why this call belongs here.
        traceEnvironment(buildStamp);
        // Read BEFORE the probe, because the probe adds scratch slides too — and
        // a slide the diagnostic created is exactly one worth a picture. An
        // unreadable deck here is not a reason to stop: it costs the id diff and
        // nothing else, and the round says so rather than reporting an empty
        // diff as "the round added nothing".
        const idsBefore = await deckSlideIds();
        // How loaded the deck was when the round started, as the first thing the
        // crash log carries.
        //
        // A run that dies leaves only its steps, and on 2026-08-09 a round died
        // sixteen seconds in — two 8-second stalls and then the tab — with no
        // way to tell an already-tired PowerPoint from a fresh one. The deck is
        // the difference: this project has documented since 2026-08-06 that
        // heavy work on a deck that is already large is what kills the tab, and
        // the runbook splits the demo halves across two decks for exactly that.
        // Whether a crashed round was on a fresh deck is the first question
        // anyone asks, and nothing recorded the answer.
        // The round's own environment, once, from local sources only. Half the
        // hypotheses this project has entertained are about the tab rather than
        // the deck — its age most of all, which has been a live candidate for
        // ten rounds with nothing measuring it.
        trace("selftest", "round starting", {
          deckSlides: idsBefore?.length ?? "unreadable",
          env: roundEnvironment(),
        });
        note("Round 1 of 2 — asking this PowerPoint what it actually does…", "busy");
        let sheet = await runHostProbes(host, buildStamp);
        // Written into the bundle before the long half starts. A self-test that
        // takes the tab down must not also lose the probe's answers, which are
        // complete, cheap, and the half most likely to be worth reading.
        lastRunLog = { build: buildStamp, host, runs: [], hostAnswers: sheet };
        // …and into the store that outlives this JavaScript context, for the
        // same reason and against a bigger loss. `lastRunLog` is a module
        // variable: the line above protects the probe's answers from a battery
        // that FAILS, and not at all from a tab that DIES, which is the way
        // these rounds actually end. The sheet was already complete and minutes
        // old at that point, and it went with the tab every time.
        recordCrashFinding("hostAnswers", sheet);
        note(`Probe done — ${describeHostSheet(sheet)} Now the self-test…`, "busy");
        if (isStopRequested()) {
          const saved = downloadJson("ssf-charts-round.json", lastRunLog);
          note(
            saved
              ? "Stopped after the probe. Its answers are saved."
              : "Stopped after the probe, but the browser would not save the file — press Download run log.",
            saved ? "ok" : "err",
          );
          endCrashLog(saved);
          return;
        }
        setSelfTestRasterizer(boundedRaster);
        setSelfTestPrompt((message) => note(message, "busy"));
        const results = await runSelfTest(undefined, scenarioPick?.value || undefined, (r) =>
          // Each verdict banked as it lands. A battery that never returns never
          // writes its report, and ordering `SCENARIOS` can only choose which
          // verdicts a crash costs — this is what makes it cost none of the
          // ones already reached.
          recordCrashFinding(`selftest:${r.name}`, r),
        );
        // The questions an empty deck could not be asked.
        //
        // The sheet is built at the top of a round, before the battery has drawn
        // anything — so a question needing an id THIS host has agreed to name is
        // put at the one moment no such id exists.
        // `shape-resolve-held-slide-proxy` answered `no-scratch-shape` in 216 of
        // 216 rounds for that reason and never once reached its question.
        //
        // Re-asked HERE, after the battery has drawn and tagged charts, rather
        // than by moving the sheet: the early write is what keeps the probe's
        // answers when the battery takes the tab down, and that is the failure
        // these rounds actually have. The merge re-banks the sheet, so the worst
        // a re-ask can cost is the re-ask.
        //
        // Skipped when nothing got tagged. Then there IS no named id, and
        // `no-scratch-shape` is the true answer rather than a deferral.
        // OBSERVED, not remembered. See `refreshNamedShapeFromDeck`: a chart
        // recorded when its tag was written is only as good as the minutes that
        // follow, and a round spends those redrawing and sweeping. Rounds 242,
        // 243, 246 and 247 all answered `unreadable` off an id whose shape was
        // no longer on the slide, and 247 proved it by listing the slide.
        //
        // Best-effort: a scan that throws must not cost the re-ask, which can
        // still run on whatever the record holds.
        let observed = namedShape();
        try {
          observed = await refreshNamedShapeFromDeck();
        } catch (err) {
          trace("probe", "could not scan the deck for a chart to name", { error: errorText(err) });
        }
        const plan = reaskPlan(sheet, observed);
        if (plan.ask.length) {
          note(`Re-asking ${plan.ask.length} question(s) now that a chart exists…`, "busy");
          const again = await runHostProbes(host, buildStamp, { only: plan.ask, passes: 1 });
          sheet = mergeHostSheets(sheet, again);
          recordCrashFinding("hostAnswers", sheet);
          trace("probe", "re-asked what the empty deck could not answer", {
            asked: plan.ask,
            answered: plan.ask.map((id) => sheet.answers.find((r) => r.id === id)?.answer ?? "missing"),
          });
        } else if (plan.skipped.length) {
          trace("probe", "left questions deferred — no chart this host would name", { deferred: plan.skipped });
        }
        /**
         * ASSEMBLED BEFORE THE STEP THAT KILLS THE HOST, and this ordering is
         * the whole point.
         *
         * `collectDeckEvidence` used to run FIRST and this object was built from
         * its result. Every verdict was already in by then — that function's own
         * docstring says so — but if the host died inside it, `lastRunLog` was
         * never assigned, `demo-log` was never enabled, and a complete round
         * evaporated. The driver found nothing to download and filed a crash.
         *
         * That is not rare and it is not cheap. **33 of 49 crash records in
         * `crashes/` hold all fourteen verdicts**, over 13 builds — about a tenth
         * of this archive again, thrown away for want of a deck scan. On
         * 2026-08-29 one 4:3 leg produced a full fourteen-scenario result FIVE
         * times, 14/14 four of them, and archived none.
         *
         * And the scan is where it dies: per-phase tracing puts 9 builds on
         * `collecting deck evidence — scanning` against 4 on everything else,
         * in a band of 441-572s.
         *
         * IT WORKS BECAUSE THE PANE OUTLIVES THE HOST. A crash takes Office.js
         * down, not this iframe — `keepCrashedRun` presses "Download the crashed
         * run" after exactly this crash and gets a file every time. Saving is
         * Blob and DOM, no host call, so a button enabled here still answers.
         *
         * `slideSize()` is a host call and stays on this side of the scan, where
         * the host is still alive to answer it.
         *
         * Same lesson as "SAVE FIRST, then end the record" below, one step
         * earlier in the same sequence.
         */
        lastRunLog = {
          build: buildStamp,
          host,
          // WHICH SLIDE SIZE THIS ROUND RAN AT, and it is load-bearing rather
          // than decorative. Until 2026-08-16 every round in the archive was
          // 16:9 and nothing said so — then the first 4:3 round was filed into
          // the same directory, where `npm run rounds` pools it with the rest.
          // Averaging two aspect ratios into one number is the rounds 24-and-25
          // mistake ("differed only in this, and were compared as though they
          // did not"), and a nightly 4:3 round would repeat it every night.
          //
          // Read through `slideSize()`, which resolves it from the host with
          // three fallback rungs, so this records what the ROUND actually ran
          // at rather than what anyone believed it would.
          slideSize: await slideSize(),
          runs: [],
          hostAnswers: sheet,
          selftest: results,
          ...(tracing() ? { trace: traceLog(traceFrom) } : {}),
        };
        /**
         * THE BUTTON STAYS DISABLED UNTIL THE ROUND IS ACTUALLY OVER, and
         * enabling it here — which the first version of this did — is a
         * regression that silently strips deck evidence from rounds.
         *
         * `scripts/round.mjs:2002` is the reason:
         *
         *     if (/button "Download run log"(?! \[disabled\])/.test(dl)) break;
         *
         * That button becoming enabled IS the driver's "the round has finished"
         * signal. Enabling it before `collectDeckEvidence` made the driver break
         * out of its wait loop mid-tail and download the banked log — verdicts
         * complete, `deck` absent, trace snapshotted before the scan.
         *
         * Rounds 313 and 318 are that race, both at 4:3, and I read 313 as proof
         * the new timeout had fired. It was not. Nothing had timed out; the
         * driver simply asked early. 314, 316 and 317 kept their deck only
         * because the tail beat the next poll.
         *
         * Durability does not need this. The verdicts survive a dying host
         * through the crash log — `runLogHead` below plus one `selftest:` finding
         * per scenario — which is the path `scripts/salvage-crashed.mjs` already
         * reads, and it works precisely because the pane outlives the host. The
         * button is a completion signal, not a safety net, and using it as both
         * made it a poor version of each.
         */
        /**
         * PERSISTED, because a variable does not survive what comes next.
         *
         * `lastRunLog` is a variable. Recovery RELOADS the tab after a crash, so
         * anything held only in memory is gone before the driver could press
         * anything — which is why the ordering above needs a durable partner.
         *
         * The crash log is that partner and it is already proven: every verdict
         * reaches `crashes/*.json` through `recordCrashFinding`, and all 49
         * records exist because the pane answered a button press AFTER the host
         * died. This adds the four fields a salvage needs that the verdicts do
         * not already carry, so a crashed run becomes a complete round rather
         * than a pile of verdicts missing their header.
         *
         * Deliberately NOT the whole log: `selftest` is already here one key per
         * scenario and the trace is the bulk of the record, so copying either
         * would double the file to say nothing new.
         */
        recordCrashFinding("runLogHead", {
          build: lastRunLog.build,
          host: lastRunLog.host,
          slideSize: lastRunLog.slideSize,
          runs: lastRunLog.runs,
        });
        // Gathered after the scenarios, before the file is written: this is the
        // upload that used to be the owner's job — save the deck, screenshot the
        // pane, attach both. It is a best-effort tail, so a host too far gone to
        // describe its own deck still gets the verdicts out.
        note("Collecting what landed on the slides…", "busy");
        const deck = await collectDeckEvidence(idsBefore);
        /**
         * THE TRACE IS RE-TAKEN WHETHER OR NOT THE TAIL CAME BACK, and the first
         * version of this got that wrong in the one case it was written for.
         *
         * Guarded on `deck`, the re-take never ran when the tail failed — so the
         * filed round carried a trace snapshotted BEFORE the scan, ending at the
         * last pre-tail line, with `deck` absent and nothing anywhere saying why.
         * Round 313 is the proof: 14/14 at 4:3, no `deck`, and not one
         * `collecting deck evidence` line in its trace, so the abandonment I had
         * just added a diagnostic for was the one thing it could not show.
         *
         * That is this repo's most repeated defect — UNKNOWN PRINTED AS NOTHING —
         * arrived at by guarding a diagnostic on the success it was meant to
         * explain the absence of. `traceLog` is a local array read, so it is
         * safe on a host that has stopped answering, which is exactly when it
         * matters most.
         */
        lastRunLog = {
          ...lastRunLog,
          ...(deck ? { deck } : {}),
          /**
           * HOW MANY TIMES THIS ROUND ASKED THE HOST TO SYNC.
           *
           * Taken here, after the deck scan, so it covers the whole round —
           * including the phase 41 of the last 60 4:3 crashes die in. See
           * `syncsSoFar` for why it is worth having: if office-js#6329 is right
           * that every sync forces a full presentation save, then the load this
           * add-in puts on the save channel is proportional to THIS number and
           * not to any shape count, and every crash figure this repo owns is
           * indexed by shapes.
           *
           * A crashed round never reaches this line, which is why the count is
           * traced before the scan as well: the trace survives into the crash
           * record, so a round that died still reports one.
           */
          syncs: syncsSoFar(),
          ...(tracing() ? { trace: traceLog(traceFrom) } : {}),
        };
        // NOW the round is over, and only now may the button say so — it is the
        // driver's completion signal (`scripts/round.mjs:2002`), so enabling it
        // any earlier makes the driver download a round that is still running.
        ($("demo-log") as HTMLButtonElement).disabled = false;
        // Only what THIS round added, and only what it could name. The button
        // stays disabled when the id diff came back empty, because a cleanup
        // with nothing to work from is one that would have to guess which
        // slides look like a test — and guessing about deletion in a user's own
        // deck is not a trade this pane makes.
        tidyable = deck?.newSlides ?? [];
        ($("demo-tidy") as HTMLButtonElement).disabled = tidyable.length === 0;
        const litter = describeLitter(deck);
        // SAVE FIRST, then end the record — and end it with what the save
        // actually did. The other order is what lost a real round: the run was
        // marked finished, which made it unrecoverable, and the download was
        // attempted afterwards. PowerPoint died, the pane reopened, and there
        // was nothing to offer back.
        const saved = downloadJson("ssf-charts-round.json", lastRunLog);
        endCrashLog(saved);
        const needed = sheetNeedsAttention(sheet) || selfTestNeedsAttention(results);
        note(
          `Round finished. ${describeSelfTest(results)} · Probe: ${describeHostSheet(sheet)} ` +
            (!saved
              ? "The browser would NOT save the file — press Download run log, or copy the Live steps."
              : needed
                ? "Saved as one file — send it over. If it is not in your downloads, press Download run log."
                : "Saved as one file; nothing in it is new.") +
            litter,
          needed || !saved ? "err" : "ok",
        );
      }),
    );
    /**
     * Put the deck back.
     *
     * A round leaves slides behind on purpose — the point is a file someone can
     * open and look at — and clearing them afterwards has been a manual chore
     * once per round, in a deck that also grows and skews the next round's
     * timings. This deletes exactly the ids the last round recorded adding, one
     * at a time, and reports what the host refused rather than claiming a clean
     * sweep it did not perform. `deleteSlideById` has a whole comment about why
     * a host saying "gone" is not proof; the count here is what it actually
     * confirmed.
     */
    $("demo-tidy").addEventListener(
      "click",
      guard(async () => {
        revealSteps();
        const ids = tidyable;
        note(`Removing the ${ids.length} slide(s) the last round added…`, "busy");
        let gone = 0;
        for (const id of ids) if (await deleteSlideById(id)) gone++;
        // Emptied whatever happened: a second press would re-ask about slides
        // the host has already refused once, and the honest state after a
        // partial sweep is "there is no longer a list I trust".
        tidyable = [];
        keepDisabled($("demo-tidy") as HTMLButtonElement);
        note(
          gone === ids.length
            ? `Cleaned up — ${gone} slide(s) removed.`
            : `Removed ${gone} of ${ids.length}. The host would not take the rest; delete those by hand.`,
          gone === ids.length ? "ok" : "err",
        );
      }),
    );
    const crashBtn = $("demo-crashlog") as HTMLButtonElement;
    const crashed = recoverCrashLog();
    if (crashed) {
      crashBtn.hidden = false;
      // Two different runs land here now, and telling the owner which one it is
      // is the difference between "the host died" and "the host was fine and
      // your file never arrived". Both are worth recovering; only one of them
      // means anything went wrong with the run itself.
      note(
        `A previous run ("${crashed.label}", build ${crashed.build}) ` +
          (crashed.finishedAt ? `finished, but its file was never saved` : `never reported finishing`) +
          ` — ${crashed.steps.length} step(s) were kept. Download the crashed run.`,
        "err",
      );
      crashBtn.addEventListener("click", () => {
        if (!downloadJson("ssf-charts-crashed-run.json", crashed)) {
          note("The browser would not save the file. Copy the Live steps instead.", "err");
          return;
        }
        clearCrashLog();
        crashBtn.hidden = true;
        note("Crashed run saved.", "ok");
      });
    }
    // The five paths the demo deck never touches. Its own button rather than a
    // mode of the demo run: it edits and deletes as well as inserting, and a
    // user reaching for "insert a demo deck" should not get that by accident.
    // Fill the picker from the battery's own list, so it cannot offer a
    // scenario that no longer exists or miss one that was added.
    const scenarioPick = $("demo-scenario") as HTMLSelectElement | null;
    if (scenarioPick) {
      for (const name of SCENARIO_NAMES) {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        scenarioPick.append(opt);
      }
    }
    const selfTestBtn = $("demo-selftest") as HTMLButtonElement;
    selfTestBtn.disabled = false;
    selfTestBtn.addEventListener(
      "click",
      guard(async () => {
        revealSteps();
        const buildStamp = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
        lastRunLog = undefined;
        ($("demo-log") as HTMLButtonElement).disabled = true;
        beginCrashLog({ build: buildStamp, host: describeHost(), label: "host self-test" });
        const traceFrom = traceMark();
        // The same rasteriser the demo run degrades with — the picture
        // scenario needs a real PNG, not a config that merely says "image".
        setSelfTestRasterizer(boundedRaster);
        // A scenario that blocks on a person has to be able to ask. Routed to
        // the same note the rest of the pane speaks through, so the request is
        // where the user is already looking rather than buried in a step list.
        setSelfTestPrompt((message) => note(message, "busy"));
        const results = await runSelfTest(undefined, scenarioPick?.value || undefined);
        // No runs, but a log all the same — the scenarios ARE the record, and
        // the trace beside them is what says how each verdict was reached.
        lastRunLog = {
          build: buildStamp,
          host: describeHost(),
          runs: [],
          selftest: results,
          ...(tracing() ? { trace: traceLog(traceFrom) } : {}),
        };
        ($("demo-log") as HTMLButtonElement).disabled = false;
        // Only on the way out, and only here. A run that throws past this line
        // stays marked unfinished on purpose: it produced no downloadable run
        // log either, so the storage copy is the only record it has.
        //
        // Finished, NOT saved. This path writes no file — the user presses
        // *Download run log* — so until they do, the storage copy is still the
        // only copy, and `markCrashLogSaved` is what retires it.
        endCrashLog();
        note(describeSelfTest(results), selfTestNeedsAttention(results) ? "err" : "ok");
      }),
    );
    const demoBtn = $("demo-insert") as HTMLButtonElement;
    demoBtn.disabled = false;
    demoBtn.addEventListener(
      "click",
      guard(async () => {
        revealSteps();
        const buildStamp = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
        const host = describeHost();
        const items = demoItems({ buildStamp, host });
        // Drop the previous run's log before this one starts. It used to
        // survive, so a run that produced no log of its own left "Download run
        // log" enabled and handing out an OLDER run's file, with nothing on
        // screen to say the two were unrelated.
        lastRunLog = undefined;
        ($("demo-log") as HTMLButtonElement).disabled = true;
        beginCrashLog({ build: buildStamp, host, label: "demo deck" });
        /**
         * Runs this click produced, in the order they were taken.
         *
         * "Both" mode takes each path in turn, so one click can end with two.
         * They are kept apart rather than merged: the paths fail in completely
         * different ways, and a report that did not say which was being read
         * would be diagnosed as the wrong one.
         */
        const runs: RunLog[] = [];
        /**
         * Bank a run the moment it ends, rather than when the click does.
         *
         * In "both" mode the shape path runs second and can throw — the whole
         * deck lost to host errors is a real outcome and it is raised as one.
         * Banking at the end of the click would then discard the file run's
         * log along with it, and the failing run is precisely the one worth
         * having a file for. `runs` is the same array the log holds, so later
         * entries land in it without re-assigning anything.
         */
        const record = (r: RunLog) => {
          runs.push(r);
          lastRunLog = { build: buildStamp, host, runs, ...(refusedShapeHalf ? { refusedShapeHalf } : {}) };
          ($("demo-log") as HTMLButtonElement).disabled = false;
        };
        // Which path(s) to take. Both, one after the other, is what a change
        // touching the renderer wants: same session, same host, one deck, and
        // the two accounts directly comparable — instead of two separate runs
        // an hour apart with a deploy in between.
        let mode = (($("demo-path") as HTMLSelectElement | null)?.value ?? "file") as "file" | "shapes" | "both";
        let refusedShapeHalf: RunLogFile["refusedShapeHalf"];
        // …and it has never once survived on a deck that was not empty.
        //
        // Four attempts, four crash dialogs. The shape half always draws onto
        // whatever the file half has just built, and the last one asked the
        // host for a batch of FIVE shapes on a 40-slide deck and waited 45
        // seconds for nothing. Every crash this project has recorded from the
        // demo path has that same shape: heavy shape work on a deck that is
        // already large. `docs/PUBLISHING.md` splits the two into separate
        // tests on separate decks for exactly this reason, and the option here
        // quietly puts them back together.
        //
        // So it degrades rather than obeys: run the file half, and say what
        // the shape half needs. A twenty-minute run that ends in "Sorry, we ran
        // into a problem" costs more than the measurement was worth, and it
        // costs it after the tab has already eaten the log.
        // Only asked where the answer changes anything. A deck read is a host
        // round trip, and the other two modes would pay for it to learn
        // nothing — which is also how this first landed: it broke a test that
        // counts exactly how many reads an insert is allowed to spend.
        if (mode === "both") {
          const deckNow = await slideCount();
          if ((deckNow ?? 0) > CROWDED_DECK_SLIDES) {
            mode = "file";
            note(
              `This deck already holds ${deckNow} slides, so only the file half will run. ` +
                "The shape half needs a fresh deck — every attempt at both on one deck has ended in PowerPoint's crash dialog.",
              "err",
            );
            // In the LOG as well as on screen. The note is overwritten by this
            // run's own summary within seconds, and a reader of the log
            // otherwise sees a run that was asked for both halves and did one,
            // with nothing to say the difference was deliberate.
            trace("demo", "refused the shape half on a crowded deck", { slides: deckNow, needs: "a fresh deck" });
            refusedShapeHalf = {
              slides: deckNow ?? 0,
              why: "the shape half needs a fresh deck — both halves on one deck has crashed the host every time",
            };
          }
        }
        // Where THIS run's trace starts. The buffer keeps every operation since
        // tracing was switched on, so a log that carried all of it carried other
        // runs' entries too — and reading one run's numbers against another's
        // trace is a genuinely expensive mistake.
        let traceFrom = traceMark();
        // The slowest thing the pane can do — say where it has got to, or a
        // multi-minute run is indistinguishable from a hang.
        // Fast path first: one generated .pptx, one host call. Falls through
        // to the shape-by-shape renderer when the host cannot take it, or when
        // the attempt landed nothing — never after a partial insert, which
        // would draw the whole deck again on top of what is already there.
        // "Both" is the one case where drawing it again IS the intent: the two
        // runs carry different tokens, so nothing confuses one for the other.
        if (mode !== "shapes" && canInsertSlidesFromBase64()) {
          const outcome = await insertDemoDeckAsFile(items);
          if (outcome) {
            // A log for THIS path too. The fast path is the default — the
            // checkbox ships checked and every current host advertises
            // insertSlidesFromBase64 — so it is what a real run takes, and it
            // was the one path that produced no downloadable record at all.
            // Success or failure: the failing run is the one worth having.
            const settled = outcome.verified.kind === "ok" ? outcome.verified.outcome : undefined;
            const verdicts = settled?.plan.verdicts ?? [];
            record({
              run: outcome.run,
              totalMs: outcome.totalMs,
              items: items.map((it, i) => {
                const v = verdicts.find((x) => x.slot === i);
                return {
                  title: it.title,
                  // The settled verdict, which is the honest answer here —
                  // there is no per-item render to report on a file insert.
                  status: v ? v.status : "unverified",
                  shapes: v?.shapes ?? 0,
                  ms: 0,
                  grouped: v?.tagged ?? false,
                  // On this path the deck was BUILT with the tag, so intent is
                  // simply whether the item had a config at all. The generator
                  // writes it into the .pptx directly — there is no sync to
                  // drop — so `tagged` true here against a false snapshot is
                  // a readback fault and nothing else.
                  tagged: !!it.configJson,
                  chart: !!it.configJson,
                  // The file path draws no item individually — there is one
                  // insert for the whole deck — so no item can have abandoned
                  // a call of its own.
                  abandoned: false,
                  lateOutcome: "",
                };
              }),
              deck: {
                slidesAdded: outcome.added,
                addsIssued: items.length,
                lost: Math.max(0, items.length - outcome.added),
                blank: [],
              },
              reconcile: settled,
              // Why there is no settled verdict, when there is none — the same
              // sentence the user sees, kept with the data it explains.
              unverified: outcome.verified.kind === "ok" ? undefined : outcome.verified.why,
              path: "file",
              trace: tracing() ? traceLog(traceFrom) : undefined,
            });
            if (mode === "file") {
              endCrashLog();
              note(outcome.text, outcome.status);
              return;
            }
            // Both: the shape path runs next, on top of what just landed, and
            // gets its own slice of the trace. Without a fresh mark the second
            // run's log would open with the first run's entries.
            note(`${outcome.text} Now drawing the same deck shape by shape…`, "busy");
            traceFrom = traceMark();
          } else if (mode === "both") {
            note("The host would not take a generated deck — running the shape path only.", "busy");
          } else {
            note("The host would not take a generated deck — drawing it shape by shape instead.", "busy");
          }
        }
        // NO budget exemption for the harness's own slides any more. It existed
        // because a large deck's contents page ran past the limit, and it is
        // what let a 79-shape text slide through as the SECOND thing a run
        // drew — five seconds before PowerPoint on the web crashed on
        // 2026-07-31. `buildIndexScenes` and `buildResultsScenes` now measure
        // their pages and split them instead, so there is nothing left to
        // exempt: a harness page that somehow still runs over is a page worth
        // skipping, exactly like any other.
        const {
          run,
          results,
          slidesAdded,
          addsIssued,
          blankSlides,
          blankItems,
          blanksRead,
          reconcile,
          degradedAt,
          degradeReason,
          totalMs,
        } = await insertDemoDeck(
          items.map((i) => ({
            scene: i.scene,
            tagData: i.configJson,
            title: i.title,
          })),
          (done, total) => {
            note("Inserting demo slides… {done} of {total}", "busy", { done, total });
            setProgress(done / total); // one slide per context, so a real bar
          },
          {
            // Close the run by reading the deck back and repairing it — the
            // per-item bookkeeping below is written while the host is still
            // committing, and has been observed calling a chart failed that
            // in fact landed twice.
            reconcile: true,
            // And give it somewhere to go when the host starts failing.
            // Rasterizing needs the pane's canvas, so the renderer asks and
            // this supplies; it decides when.
            // Web only. Degrading exists because a browser has no safety
            // net — Office's CPU/memory/crash-tolerance limits are scoped to
            // Windows and Mac — so on desktop, where the host throttles the
            // add-in rather than dying, drawing shapes remains the right
            // answer however long it takes.
            pictureFor: isWebHost() ? (i) => boundedRaster(items[i].scene) : undefined,
            shapeBudget: isWebHost() ? WEB_SHAPE_BUDGET : undefined,
          },
        );
        // Self-check: the deck is a regression harness, so report what the HOST
        // actually did, not what we asked for. The full table goes to the console.
        const named = (s: "skipped" | "failed") =>
          results.map((r, i) => (r.status === s ? items[i].title : "")).filter(Boolean);
        const skipped = named("skipped");
        const failedNames = named("failed");
        const rendered = results.filter((r) => r.status === "rendered").length;
        // Loss vs adds ISSUED, not vs items.length: a retry/fail stray inflates
        // slidesAdded, so measuring against items.length reads 0 during real
        // corruption when a stray cancels a lost slide. addsIssued − slidesAdded
        // counts adds that never landed (strays that landed cancel out).
        const lost = Math.max(0, addsIssued - slidesAdded);
        const secs = (totalMs / 1000).toFixed(1);
        console.log("SSF Charts demo self-check:");
        console.table(
          results.map((r, i) => ({
            chart: items[i].title,
            shapes: r.created,
            status: r.status,
            grouped: !!r.grouped,
            tagged: !!r.tagged,
            ms: r.ms,
            abandoned: !!r.abandoned,
            lateOutcome: r.lateOutcome ?? "",
          })),
        );
        console.log(
          `deck grew by ${slidesAdded}, issued ${addsIssued} adds${lost > 0 ? ` — ${lost} LOST` : ""}; blank slots ${blankSlides.length ? blankSlides.join(", ") : "none"}${blanksRead ? "" : " (blank check incomplete)"} · total ${secs}s`,
        );
        // The headline comes from the settled read when there is one. Every
        // number above was computed while the host was still committing, and
        // the run that produced Presentation_4.pptx got three of them wrong:
        // it blamed Gantt for failing (it landed twice), called a full Agenda
        // slide blank, and counted duplicate slides as successes.
        const verdicts = reconcile?.plan.verdicts ?? [];
        const missing = verdicts.filter((v) => v.status === "lost" || v.status === "empty").map((v) => v.title);
        let msg = reconcile
          ? `Deck settled: ${describeReconcile(reconcile.plan)} — in ${secs}s.`
          : `Inserted ${rendered} of ${items.length} in ${secs}s.`;
        if (skipped.length) msg += ` Skipped as too dense (stamped): ${skipped.join(", ")}.`;
        if (reconcile) {
          if (missing.length) msg += ` Never landed: ${missing.join(", ")}.`;
          if (reconcile.unread)
            msg += ` (Could not read ${reconcile.unread} slide(s) — anything on them counts as lost here.)`;
          const { deleted, unstamped, regrouped } = reconcile.applied;
          // Deletions are NOT all duplicates — most are usually empty slides a
          // lost add left behind. Calling nine removals "9 duplicate slide(s)"
          // in the same breath as the plan's own "1 duplicate slide removed ·
          // 8 orphan slides" made the run contradict itself in one sentence.
          const dupPlanned = reconcile.plan.actions.filter((a) => a.kind === "delete" && a.slot !== null).length;
          const emptyPlanned = reconcile.plan.actions.filter((a) => a.kind === "delete" && a.slot === null).length;
          if (deleted || unstamped || regrouped)
            msg +=
              ` Repaired: removed ${deleted} slide(s)` +
              (deleted ? ` (${dupPlanned} duplicate, ${emptyPlanned} empty)` : "") +
              `, cleared ${unstamped} false banner(s), re-grouped ${regrouped} chart(s).`;
          if (reconcile.refused) msg += ` ⚠ ${reconcile.refused} repair step(s) the host refused.`;
        } else if (failedNames.length) msg += ` Host failed on: ${failedNames.join(", ")}.`;
        if (degradedAt !== undefined)
          msg += ` Drew the last ${items.length - degradedAt} slide(s) as pictures — ${degradeReason}. Use "Explode to native shapes" on any of them to get real shapes back.`;
        // A rendered but ungrouped chart is not re-editable — flag them so
        // Phase 2 doesn't quietly count them as full successes.
        const ungrouped = reconcile
          ? verdicts.filter(
              (v) => !v.tagged && v.status !== "lost" && v.status !== "skipped" && items[v.slot]?.configJson,
            ).length
          : results.filter((r, i) => r.status === "rendered" && !r.grouped && items[i].scene.nodes.length > 1).length;
        if (ungrouped) msg += ` ⚠ ${ungrouped} chart${ungrouped === 1 ? "" : "s"} landed ungrouped (not re-editable).`;
        if (lost > 0)
          msg += ` ⚠ ${lost} add${lost === 1 ? "" : "s"} did not land — the host lost slides (issued ${addsIssued}, deck grew by ${slidesAdded}).`;
        // Blank slides carry the slot tag (item title) where the host has 1.3
        // slide tags; without them the entry has title null and only its deck
        // position is shown, same as before slot tags landed.
        if (blankSlides.length) {
          const named = blankItems.map((b) => (b.title ? `${b.title} (slide ${b.position})` : `slide ${b.position}`));
          msg += ` ⚠ ${blankSlides.length} slide${blankSlides.length === 1 ? "" : "s"} came back BLANK: ${named.join(", ")}.`;
        } else if (!blanksRead) msg += ` (Blank check did not finish.)`;
        // Keep the whole run, not just the sentence. A run that ends badly is
        // the one worth reporting, and it is also the one whose results slide
        // is most likely to be the thing the host drops.
        // Close the deck with a self-contained results slide so the exported PDF is
        // a complete run record. A second insertDemoDeck reuses the same add/render/
        // self-check machinery; wrap it so a host stall here can't swallow the run's
        // own summary (its failure is itself just another data point).
        const rows: ResultRow[] = results.map((r, i) => ({
          title: items[i].title,
          status: r.status,
          shapes: r.created,
          ms: r.ms,
        }));
        const summary: ResultsSummary = {
          buildStamp,
          items: items.length,
          rendered,
          skipped: skipped.length,
          failed: failedNames.length,
          lost,
          totalMs,
        };
        // Each results page inserts in its OWN insertDemoDeck call. A single
        // page failing must NOT drop the rest — insertDemoDeck throws when
        // every item in its batch failed, so batching all pages together
        // meant page 2's failure lost page 1. Presentation_3.pptx surfaced
        // this: "(results slide not added)" with zero pages landed.
        let resultsPages: Scene[] = [];
        try {
          resultsPages = buildResultsScenes(rows, summary);
        } catch (e) {
          console.warn("SSF Charts: results scene build failed", e);
        }
        let resultsLanded = 0;
        for (const [i, scene] of resultsPages.entries()) {
          try {
            await insertDemoDeck(
              [
                {
                  scene,
                  title: resultsPages.length === 1 ? "Results" : `Results (page ${i + 1} of ${resultsPages.length})`,
                },
              ],
              undefined,
              {
                // The same protections the run itself gets. This insert runs at
                // the WORST moment — right after a run that has just finished
                // exhausting the host — and it used to get none of them.
                //
                // Without `reconcile`, a page whose add landed but whose shapes
                // did not left a stamped, untagged slide at the end of the deck
                // that nothing ever cleaned: the main run's repair had already
                // finished, and its range stopped short of this slide. A real
                // 38-item run ended exactly that way.
                reconcile: true,
                // And without a picture to fall back on, the run's own summary
                // is the first casualty of a failure-heavy run — it is being
                // drawn precisely because things went badly.
                pictureFor: isWebHost() ? () => boundedRaster(scene) : undefined,
              },
            );
            resultsLanded += 1;
          } catch (e) {
            console.warn(`SSF Charts: results page ${i + 1} failed to insert`, e);
          }
        }
        if (!resultsPages.length) msg += " (results slide not added)";
        else if (resultsLanded === 0) msg += " (results slide not added)";
        else if (resultsLanded < resultsPages.length)
          msg += ` (${resultsLanded} of ${resultsPages.length} results pages added)`;
        // Written LAST, so the trace it carries covers the whole run including
        // the results pages. Taken before them, the log ended at the repair
        // read — and when a results page then failed, the file said
        // "(results slide not added)" with nothing in it about why.
        record({
          run,
          totalMs,
          items: results.map((r, i) => ({
            title: items[i].title,
            status: r.status,
            shapes: r.created,
            ms: r.ms,
            grouped: !!r.grouped,
            tagged: !!r.tagged,
            chart: !!items[i].configJson,
            abandoned: !!r.abandoned,
            lateOutcome: r.lateOutcome ?? "",
          })),
          deck: { slidesAdded, addsIssued, lost, blank: blankItems },
          reconcile,
          path: "shapes",
          trace: tracing() ? traceLog(traceFrom) : undefined,
        });
        // Reached only by a run that got all the way here. One that did not
        // stays marked unfinished, which is what offers it back on reopen —
        // and so does this one until *Download run log* is pressed, because
        // finishing and being saved are different facts.
        endCrashLog();
        note(
          runs.length > 1 ? `Both paths run. File: ${runs[0].deck.slidesAdded} slides. Shapes: ${msg}` : msg,
          lost > 0 || failedNames.length || blankSlides.length ? "err" : "ok",
        );
      }),
    );
  } else {
    insertBtn.disabled = true;
    loadBtn.disabled = true;
    ($("agenda-insert") as HTMLButtonElement).disabled = true;
    note("Not running inside PowerPoint — use Download SVG, or sideload the manifest to insert native shapes.");
  }
}

// Ribbon deep-link: taskpane.html?kind=waterfall preselects a chart type;
// ?tab=elements opens a tab and ?el=harvey focuses that element's card
// (the ribbon's "Insert element" menu uses these).
const deepLink = new URLSearchParams(location.search);

/**
 * Does the Testing section have to be asked for, or is it simply there?
 *
 * **THE PRODUCT SHIPS ITS OWN TEST HARNESS IN ITS UI.** The Automation tab's
 * first half is `Testing` — demo deck, self-test, host probe, one experiment,
 * download run log, clean up the last round — and every published build has
 * carried it, ungated, to every user. It is written for the person who runs the
 * round loop, and it reads that way: "the fake host every test runs against",
 * "a full round takes minutes and leaves its slides behind", "clean up the last
 * round". A stranger cannot parse it, and two of those buttons change their
 * document.
 *
 * That is fine for a sideloaded tool and is a problem for a store listing,
 * where a reviewer opens every tab cold.
 *
 * **WHY A URL PARAMETER AND NOT A BUILD FLAG.** A build flag would mean the
 * bundle users get is not the bundle the round loop tests, and this project's
 * whole validation rests on those being the same artifact — `round.mjs` refuses
 * to run when HEAD and the deployed stamp differ, for exactly that reason. One
 * build, one deployment; only the manifest differs. `build-manifest.mjs`
 * already emits more than one manifest, so the harness gets one carrying
 * `?harness=1` and users get one that does not.
 *
 * **HIDDEN, NOT REMOVED.** `hidden` takes the section out of the accessibility
 * tree, which is the tree `round.mjs` searches with `find` — so an un-opted-in
 * pane genuinely cannot be driven — while leaving the nodes in place so the
 * queries app.ts makes at init keep working. Removing the subtree would mean
 * auditing every one of those, for no gain.
 *
 * **DEFAULT `false`, WHICH IS TODAY'S BEHAVIOUR.** Flipping this changes what a
 * user receives, so it is the owner's call and not this file's. Flip it, and
 * regenerate the manifests, together: on its own this hides the section from the
 * round driver as well and the loop stops.
 */
const TESTING_UI_NEEDS_OPT_IN = false;
if (TESTING_UI_NEEDS_OPT_IN && deepLink.get("harness") !== "1") {
  const testing = document.getElementById("testing-section");
  if (testing) testing.hidden = true;
}

const requestedKind = deepLink.get("kind");
if (requestedKind && CHART_KINDS.some((k) => k.kind === requestedKind)) {
  applyConfig(sampleConfig(requestedKind as ChartKind), null);
}
// A shared chart link (#c=<base64 config>) reopens the exact chart. Applied
// after ?kind so an explicit link wins; malformed links are ignored silently.
if (location.hash.startsWith(CONFIG_HASH)) {
  try {
    const cfg = JSON.parse(decodeURIComponent(atob(location.hash.slice(CONFIG_HASH.length)))) as ChartConfig;
    if (cfg && typeof cfg === "object" && cfg.kind) {
      applyConfig({ ...DEFAULT_SIZE, ...cfg }, null);
      note("Chart loaded from a shared link.", "ok");
    }
  } catch {
    /* malformed share link — fall through to the default chart */
  }
}
const requestedTab = deepLink.get("tab");
if (requestedTab) {
  // Match on the dataset value rather than interpolating the parameter into a
  // selector: ?tab=chart"] made querySelector THROW at module top level, which
  // aborted the rest of boot — no build stamp, no size inputs, and no
  // wireInsert(), leaving the Insert button enabled with no click handler.
  // An unknown tab is now ignored silently, like a malformed share link.
  [...document.querySelectorAll<HTMLButtonElement>(".tabs .tab")]
    .find((el) => el.dataset.tab === requestedTab)
    ?.click();
}
const requestedEl = deepLink.get("el");
if (requestedEl) {
  const card = document.getElementById(`${requestedEl}-insert`)?.closest(".el-card");
  if (card) {
    card.scrollIntoView({ block: "center" });
    card.classList.add("el-flash");
    setTimeout(() => card.classList.remove("el-flash"), 1600);
  }
}

// Injected by vite (see vite.config.ts). Shown in the header so the running
// build is always identifiable — PowerPoint caches the pane, and a stale one is
// otherwise indistinguishable from a fixed one.
declare const __BUILD_STAMP__: string;
const stampEl = document.getElementById("build-stamp");
if (stampEl) stampEl.textContent = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";

/**
 * Say so when this pane is older than the site it came from.
 *
 * The stamp above has always been readable, and reading it is a habit the
 * runbook asks for — but it can only be checked against something you already
 * know, and the number you need is on GitHub. Twice now a round has been run on
 * a build that predated the fix it was meant to test, and the second time was
 * spent hard-reloading a pane that looked identical either way.
 *
 * The cause is not the reload. GitHub Pages serves the pane HTML with
 * `Cache-Control: max-age=600` and gives us no way to set headers, so for ten
 * minutes after a deploy PowerPoint hands back the cached page, which names the
 * previous hashed bundle — still in the browser's cache even after it 404s on
 * the server. `build.json` is written by `pages-postbuild.mjs` from the stamp
 * inside the built bundle, and is small enough to re-fetch every boot.
 *
 * **This is the first outbound request this add-in has ever made**, and that
 * matters beyond the feature: the standing read of the ~83 CSP `connect-src`
 * violations in the host console has been "not ours — we make no connections".
 * That argument is now weaker by one same-origin GET, so if those violations
 * change shape, look here first.
 *
 * Everything about it fails quiet. No `build.json` (a dev server, an older
 * deploy, a host CSP that blocks it) leaves the pane exactly as it was: this is
 * a convenience, and a diagnostic that breaks the pane it is diagnosing is
 * worse than the trap it replaces.
 */
async function warnIfStale(): Promise<void> {
  const running = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
  if (!stampEl || running === "dev") return;
  try {
    // Cache-busted, or the check inherits the very caching it exists to detect.
    const res = await fetch(`/build.json?t=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) return;
    const live = (await res.json())?.build;
    if (typeof live !== "string" || !live || live === running) return;
    stampEl.textContent = `${running} — SITE HAS ${live}`;
    stampEl.classList.add("stale");
    stampEl.title = "This pane is older than the deployed site. Hard-reload the whole PowerPoint tab (Ctrl+F5).";
    trace("pane", "the pane is older than the site", { running, live });
  } catch {
    /* offline, blocked, or no such file — say nothing rather than something wrong */
  }
}
void warnIfStale();

const chartWInput = $("chart-w") as HTMLInputElement | null;
const chartHInput = $("chart-h") as HTMLInputElement | null;
// Seed the inputs from state, then let them WRITE state on edit (state is the
// source of truth currentConfig reads). A sub-usable value is ignored so the
// preview holds its last good size instead of snapping to the default.
if (chartWInput) chartWInput.value = String(state.width);
if (chartHInput) chartHInput.value = String(state.height);
const syncSizeFromInputs = () => {
  const w = Number(chartWInput?.value);
  const h = Number(chartHInput?.value);
  if (Number.isFinite(w) && w >= 80) state.width = w;
  if (Number.isFinite(h) && h >= 60) state.height = h;
  renderPreview();
};
chartWInput?.addEventListener("input", syncSizeFromInputs);
chartHInput?.addEventListener("input", syncSizeFromInputs);

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    wireInsert();
    // ASK AGAIN NOW THAT THERE IS AN ANSWER. See `syncHostOnlyButtons`: the
    // module-scope call above ran before `Office.context` existed and disabled
    // both deck-style buttons for the whole session.
    syncHostOnlyButtons();
    // The deck's own style, if it carries one. Fired and forgotten: the pane is
    // usable before it answers, and a host that cannot store one (or a deck
    // that has none) simply leaves the browser's style in force. It re-renders
    // only when there is something to show, so an ordinary load does not flash.
    // WARM FIRST, THEN READ. The first custom-XML call after a pane loads fails
    // on this host and the second works — seven rounds, and the diagnostic probe
    // (always the second call) has answered on every one of them. `#602` tried
    // retrying inside the read and `#603` reverted it: the read's own retry
    // spent the good second call and left the probe third, so the probe stopped
    // answering too. Spending it HERE costs one thrown-away count and leaves the
    // real read as the second call.
    //
    // Still fired and forgotten — the pane is usable before any of this answers.
    void warmCustomXmlSurface()
      .then(() => readDeckStyle())
      .then((style) => {
        if (!style) return;
        deckStyle = style;
        renderPreview();
      });
    try {
      localizePane(Office.context.displayLanguage);
    } catch {
      /* no display language available */
    }
  });
} else {
  wireInsert();
  localizePane(new URLSearchParams(location.search).get("lang") ?? undefined);
}
