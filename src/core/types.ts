/**
 * A point shape — RE-EXPORTED, not restated.
 *
 * This was a second `"circle" | "square" | SymbolShape` beside the one in
 * ./geometry, and the two agreed only by luck. The file that owns `MARKER_AREA`
 * and `markerSymbolOf` is the one that decides what a marker may BE, so a copy
 * here could disagree with it and still typecheck. It nearly did: adding
 * `hexagon` to SymbolShape for the hex tile map widened both copies at once,
 * silently making a tile shape a public scatter marker in the config surface —
 * the accident the comment on the geometry-side type describes.
 */
import type { MarkerSymbol } from "./geometry";
export type { MarkerSymbol };

/** Chart kinds supported by the layout engine (think-cell equivalents in comments). */
export type ChartKind =
  | "stacked" // think-cell "Stacked" column chart
  | "clustered" // "Clustered"
  | "stacked100" // "100%" chart
  | "waterfall" // "Waterfall" (build up/down, "e" totals)
  | "mekko" // "Mekko" / Marimekko with %-axis
  | "line" // "Line"
  | "area" // "Area"
  | "butterfly" // two back-to-back bar charts sharing one scale
  | "scatter" // "Scatter" — X/Y rows, labelled points
  | "bubble" // "Bubble" — scatter + Size row
  | "gantt" // Gantt (Start/End rows), numeric or calendar timeline
  | "combo" // stacked columns + line series (Series.type === "line")
  | "pie" // pie chart (first series)
  | "doughnut" // doughnut variant of pie
  | "boxplot" // box-and-whisker: precomputed Min/Q1/Median/Q3/Max rows or raw samples
  | "radar" // spider chart: categories = spokes, series = polygons
  | "heatmap" // matrix: series = rows, categories = columns, value → color
  | "tilemap" // tile-grid cartogram: categories = region codes, value → color
  | "cascade" // decomposition: each stage is a subset of the previous one
  | "funnel" // pipeline stages as centered bands, width ∝ value
  | "waffle" // 10×10 unit grid, part-to-whole (categories = parts)
  | "treemap" // squarified rect packing, area ∝ value; "Group | Item" → 2 levels
  | "sunburst" // nested wedge rings from a "Group | Item" hierarchy
  | "violin" // distribution shape per category (mirrored density) from raw samples
  | "candlestick"; // OHLC financial bars (Open/High/Low/Close rows)

export interface Series {
  name: string;
  /** One value per category; null = blank cell. */
  values: (number | null)[];
  /** Override the palette color for this series. */
  color?: string;
  /**
   * Per-category fill overrides — highlight a single segment, point, or
   * slice ("color on a data point"). null cells keep the series color.
   */
  colors?: (string | null)[];
  /**
   * Hatch/dot pattern over the series fill — extends a tight color budget
   * and survives grayscale printing. Column-family segments; SVG/preview
   * renders the pattern, PowerPoint output degrades to the solid fill.
   */
  pattern?: "diagonal" | "crosshatch" | "dots" | "horizontal";
  /**
   * Combo charts: how this series renders over the columns. "line" connects the
   * points; "marker" draws unconnected points only — a per-category benchmark,
   * target or consensus, where a connecting line would falsely imply the values
   * interpolate between categories.
   *
   * A marker series shares whatever scale the overlays use, so it only means
   * what it looks like on the columns' own scale: don't put one on a
   * `secondaryAxis` or under `combo.lineAxes: "independent"`, where it would be
   * measured against a different axis than the columns it benchmarks.
   */
  type?: "column" | "line" | "marker";
  /**
   * Stack group for clustered-stacked charts (think-cell separates stacks
   * with blank datasheet rows). Series sharing a stack index stack together;
   * different indices sit side by side within each category.
   */
  stack?: number;
  /**
   * IBCS scenario (International Business Communication Standards) — encodes the
   * data's nature in the FILL, so actual vs plan vs forecast reads without a
   * legend lookup. Column family only:
   *   AC (actual)        → solid fill
   *   PY (previous year) → lighter solid
   *   PL / BU (plan/budget) → outlined / hollow (fill + a coloured border)
   *   FC (forecast)      → hatched (diagonal pattern; solid in the live add-in)
   * The two-letter code is appended to the series' legend/label.
   */
  scenario?: "AC" | "PY" | "PL" | "BU" | "FC";
}

export interface ChartData {
  categories: string[];
  series: Series[];
  /**
   * think-cell's "100%=" datasheet row: per-category denominators for 100%
   * charts. When set, columns whose series sum to less than this stay short
   * of full height. Defaults to the column sum.
   */
  hundredPercent?: (number | null)[];
  /**
   * think-cell's "X extent" row (Mekko with units): explicit column widths,
   * scaled so the total matches the chart width.
   */
  xExtent?: (number | null)[];
  /**
   * Set when datasheet cells were parsed as calendar dates (values are days
   * since the Unix epoch). Gantt charts render a month/year timeline then.
   */
  dates?: boolean;
}

export interface WaterfallOptions {
  /**
   * Categories rendered as cumulative totals (think-cell's "e" cell):
   * the bar is drawn from the baseline to the running total at that point,
   * regardless of the cell value.
   */
  totalIndices?: number[];
  /**
   * Blank spacer categories that draw no bar and leave a gap — used to group
   * a long bridge into sections. The running total carries across unchanged
   * and the dashed connector bridges the gap. (Give the category an empty
   * name so no axis label shows.)
   */
  spacerIndices?: number[];
  /**
   * "Of which" detail groups: `indices` are categories that decompose column
   * `of`'s delta into a sub-bridge, drawn in line with the others but OFF the
   * chain — they carry no running total, and the main dashed connector steps
   * over the whole group from `of` to the next chain column.
   *
   * The parent is named, never inferred: guessing it (say, "the nearest
   * preceding chain column") is what would make this look broken on the cases
   * where the guess is wrong. `indices` should be contiguous and sit after
   * `of`. A group whose values don't sum to the parent's delta renders exactly
   * as authored — the engine draws your numbers, it doesn't reconcile them.
   */
  detailGroups?: { of: number; indices: number[] }[];
}

/** think-cell style decorations, all computed from layout anchors. */
export interface Decorations {
  /**
   * Series the legend must name that are NOT in `cfg.data.series` — set by the
   * layout, never by a caller.
   *
   * A combo hands its column series to `layoutColumns` in a config it builds
   * itself (`colCfg`), with the LINE series taken out, because those are drawn
   * by the combo afterwards. The legend is drawn from that same config, so a
   * combo's legend named its columns and never its lines — at every size, not
   * only crowded ones. On a horizontal combo small enough that the line's own
   * name will not fit either, the series was left with no name anywhere on the
   * chart: 120x90 at 6pt drew "Product" and "Services" and nothing for
   * "Margin %".
   *
   * It goes on `Decorations` rather than on the config because `decor` is the
   * one thing threaded to BOTH the reservation (`horizontalLegendFits`, and the
   * band `computeFrameHorizontal` keeps) and the draw (`horizontalChrome` into
   * `legendRow`). Gating one without the other is the bug this file's own
   * comments say has been written twice.
   */
  legendAlso?: { label: string; color: string; stroke?: string; strokeWidth?: number }[];
  /** Value label inside each segment (hidden automatically when it doesn't fit). */
  segmentLabels: boolean;
  /** Series labels to the right of the last category (think-cell placement). */
  seriesLabels: boolean;
  /** Column totals above each column. */
  totals: boolean;
  /**
   * Grand total: one label at the top-right of a stacked/clustered column chart
   * showing the sum of every category total (think-cell 14). Independent of
   * `totals` — you can show the grand total without the per-column ones. Not
   * meaningful on a 100%-stacked chart (every column totals 100%), so ignored there.
   */
  grandTotal?: boolean;
  /**
   * IBCS variance tier: a strip below the columns showing the deviation of an
   * `actual` series from a `reference` (plan / previous year) per category, as
   * signed bars from a zero line — the IBCS "integrated variance" convention.
   * `actual`/`reference` are series indices. `mode` is the absolute Δ (default)
   * or relative Δ%. `goodIsUp` colours the sign (default true = up is favourable,
   * green; down unfavourable, red) — set false for cost-like metrics. Vertical
   * column family only.
   */
  variance?: {
    actual: number;
    reference: number;
    mode?: "absolute" | "percent";
    goodIsUp?: boolean;
  };
  /**
   * Combo charts only. Which set of value labels survives when the frame is too
   * small to carry both — the column totals and the line's point labels.
   *
   * On a short frame these two genuinely have nowhere to go. The de-collision
   * pass resolves every one of them if the canvas is unbounded, by flipping the
   * point label BELOW its point; on a 60pt-tall chart that destination is off
   * the bottom, so the label stays where it was and is drawn through the total.
   * Neither shrinking nor nudging helps, because the two labels are centred at
   * the same y whenever the column total and the line value coincide — which is
   * a fact about the DATA, not about the layout.
   *
   * `"columns"` (default) keeps the totals and drops the point labels that
   * collide, which is the ordering `MOVABLE` in `collide.ts` already encodes:
   * the total is anchored to the bar it labels, the point label is a floater.
   * `"line"` reverses it for charts whose subject is the line.
   *
   * Only labels that are ACTUALLY still colliding after de-collision are
   * dropped, so a roomy chart — or a short one whose data happens to separate
   * them — keeps both.
   */
  tightLabelPriority?: "columns" | "line";
  /** Category labels below the baseline. */
  categoryAxis: boolean;
  /**
   * Numeric value axis on the left (off by default, as in think-cell).
   * "datamarks" draws Tufte-style tick dashes + labels with no axis line.
   */
  valueAxis: boolean | "datamarks";
  /**
   * Tick placement for datamark axes: "nice" round values (default) or
   * "data" — ticks at the actual data min/max (Tufte's range frame).
   */
  tickMode?: "nice" | "data";
  /** Radar: gridline web shape (default "polygon", business style). */
  gridShape?: "polygon" | "circle";
  /** Radar: series fill opacity (default 0.18; 0.25 for a single series). */
  fillOpacity?: number;
  gridlines: boolean;
  /**
   * Segment label content, think-cell's label dropdown: any combination of
   * value, percent (of column), series name, and category name.
   * Default: ["value"] (["percent"] on 100% charts).
   */
  labelContent?: ("value" | "percent" | "series" | "category")[];
  /** CAGR arrow between two categories; `series` computes it on one series. */
  cagr?: { from: number; to: number; series?: number };
  /**
   * Difference arrow between two category indices. Without `series` it is a
   * total difference arrow (column totals); with `series` it is a level
   * difference arrow comparing the cumulative level at that series.
   */
  difference?: {
    from: number;
    to: number;
    percent?: boolean;
    series?: number;
    /** Anchor the arrow's start at valueLines[fromValueLine] instead of a column. */
    fromValueLine?: number;
  };
  /** Value lines: at fixed values or at the mean of column totals. */
  valueLines?: ({ mode: "mean" } | { mode: "value"; value: number })[];
  /** @deprecated legacy single value line; normalized into valueLines. */
  valueLine?: { mode: "mean" } | { mode: "value"; value: number };
  /**
   * Connector lines joining segment boundaries of adjacent stacked
   * columns/bars — they make the development of each segment much easier
   * to follow. Stacked/100% charts (single stack group).
   */
  connectors?: boolean;
  /**
   * Speech-bubble callouts commenting on a value: anchored to a column top,
   * or to the cumulative level of `series` within it. dx/dy nudge the bubble
   * from its default spot above the anchor.
   */
  callouts?: { text: string; category: number; series?: number; dx?: number; dy?: number }[];
  /**
   * Shaded background bands highlighting a region of an axis, drawn behind
   * the data. axis "y" spans a value range; axis "x" spans category indices
   * (scatter/bubble: both axes are in value units).
   */
  bands?: { axis: "x" | "y"; from: number; to: number; color?: string; label?: string }[];
  /**
   * Render a "100% = N" note in the footnote line — the classic annotation
   * telling readers what the percentages are of. Pie/doughnut (series total)
   * and 100% charts (uniform denominator).
   */
  hundredPercentNote?: boolean;
  /**
   * Line charts: categories from this index on are forecast — segments
   * turn dashed, markers hollow, and a subtle divider marks the boundary.
   */
  forecastFrom?: number;
  /**
   * Scatter/bubble: one X/Y crossing shades four quadrant zones with
   * optional corner labels ([top-left, top-right, bottom-left,
   * bottom-right]) — BCG-matrix framing in one step.
   */
  quadrants?: { x: number; y: number; labels?: [string?, string?, string?, string?] };
  /**
   * Scatter/bubble: distribution histograms in a gutter along the top (x)
   * and/or right (y) margin — where are the points concentrated, beside the
   * chart that shows how they relate. Bins SUBDIVIDE the axis's own nice
   * ticks, so every tick is a bin edge and a bar can be read straight against
   * the scale next to it. The gutter takes real space: the plot shrinks, and
   * the marginals are dropped rather than drawn if that would leave it too
   * small to be a chart.
   */
  marginals?: "x" | "y" | "both";
  /**
   * Clustered charts: draw columns as lollipops (stem + dot), plain dots
   * (Cleveland dot plot), or a two-series dumbbell range (dot–line–dot).
   */
  barStyle?: "bar" | "lollipop" | "dot" | "range";
  /** Line charts: shade the gap between two series (indices) as a ribbon. */
  fillBetween?: [number, number];
  /**
   * Line/area charts: draw stepped (staircase) segments instead of straight
   * lines. "before" jumps to the new value at the start of the interval
   * (VH), "after" holds then jumps at the next category (HV), "center"
   * steps at the interval midpoint (HVH). Area fills follow the same steps.
   */
  stepped?: "before" | "after" | "center";
  /**
   * Line charts: draw smooth (Catmull-Rom) curves through the points instead
   * of straight segments — sampled to a dense polyline since the renderers
   * have no freeform paths. Ignored when `stepped` is set.
   */
  smooth?: boolean;
  /**
   * Line charts: bridge missing data — connect straight across null points
   * instead of breaking the line into segments. (Applies to plain straight
   * lines, not smooth/stepped.)
   */
  bridgeGaps?: boolean;
  /**
   * Line charts: bump-chart mode — values are ranks (1 = best), drawn on an
   * inverted integer axis (rank 1 at the top) with thicker lines, markers and
   * "Name" labels at both ends. Rank-over-time comparison.
   */
  bump?: boolean;
  /**
   * Line charts: slope-chart mode — no value axis or gridlines, vertical
   * hairlines at the first and last category, and "Name value" labels at
   * both line ends (best with 2 categories).
   */
  slope?: boolean;
  /**
   * Scatter/bubble: connect the points in datasheet (row) order with a
   * trail and a direction arrowhead — a Gapminder-style trajectory showing
   * how one entity moved through the X/Y space over time.
   */
  trajectory?: boolean;
  /**
   * Gantt: draw a summary bar (with end caps) on each section-header row,
   * spanning min(start)→max(end) of the activities beneath it.
   */
  summaryBars?: boolean;
  /**
   * Gantt: highlight the critical path — the dependency chain (over the
   * "After" edges) with the greatest cumulative duration. Its bars get a red
   * outline and its dependency arrows are drawn thicker in red (MS-Project
   * convention). No-op without "After" rows.
   */
  criticalPath?: boolean;
  /**
   * Line/area: sparkline mode — a compact, axis-less, word-sized trend line
   * with no chrome, an optional leading label (the title/series name) and a
   * trailing value, plus min/max/last dots. Pair with `multiples` for a table
   * of sparklines, one row per series.
   */
  sparkline?: boolean;
  /**
   * Radar: shade the per-spoke min–max envelope of the "peer" series (all
   * but the last) as a band, and draw the last series prominently on top —
   * the "peer range + us" competitive-profiling view.
   */
  radarBand?: boolean;
}

export interface NumberFormat {
  /** Number of decimals; "auto" picks based on magnitude. */
  decimals: number | "auto";
  /** Append a suffix such as "%" or "€m". */
  suffix?: string;
  /** Show an explicit "+" on positive deltas (waterfall, difference arrows). */
  forceSign?: boolean;
  /** BCP-47 locale for separators, e.g. "de-DE" → 1.234,5. Default en-US. */
  locale?: string;
  /**
   * Group the thousands (1,234 / 1.234). Default true; `false` prints 1234.
   *
   * For a number that is not a quantity: a scatter or bubble chart whose X values
   * are YEARS printed "2,015" … "2,018" on its axis and beside its points, and
   * nothing in the data says a 2015 is a year and not two thousand and fifteen.
   * The author says so. Read strictly (`=== false`), so anything else a hand-edited
   * config writes there leaves the separator on, which is what it always did.
   *
   * Reaches every number the chart prints through `formatNumber` with the chart's
   * format — ticks, value labels, totals, a scatter's point labels. Percent labels
   * (a CAGR arrow, a share) keep their separator: they are quantities by
   * construction. NOT YET reached: the three formats that pick `decimals` and
   * `locale` out of this object by hand (the bubble Size key and the scatter Color
   * legend in `scatter.ts`, a Gantt `Column` cell in `gantt.ts`); they would take
   * `grouping` with them.
   */
  grouping?: boolean;
}

export interface ChartStyle {
  fontFamily: string;
  /** Base label font size in points. */
  fontSize: number;
  palette: string[];
  /** Ink colors. */
  text: string;
  mutedText: string;
  axis: string;
  gridline: string;
  /** Fill for waterfall decreases / negative deltas. */
  negative: string;
  /** Fill for waterfall totals. */
  neutral: string;
  background: string;
}

export interface ChartConfig {
  kind: ChartKind;
  data: ChartData;
  /**
   * Rotate column charts into bar charts (think-cell's rotation handle).
   * Applies to stacked/clustered/100%; decorations that assume a vertical
   * value axis are skipped in horizontal orientation.
   */
  horizontal?: boolean;
  /**
   * Manual value-axis scale (think-cell's axis-handle dragging). Either end
   * may be pinned; the other stays automatic.
   */
  scale?: { min?: number; max?: number };
  /**
   * think-cell's Segment Order menu: stacking order of series within each
   * column. `ascending`/`descending` sort per column by value.
   */
  segmentOrder?: "sheet" | "reverse" | "ascending" | "descending";
  /**
   * Sort categories by column total (think-cell's category sorting).
   * Column/Mekko/pie families only — order is meaningful for waterfall/Gantt.
   */
  categorySort?: "ascending" | "descending";
  /**
   * Combo charts: give line series their own right-hand value axis
   * (e.g. margin % over absolute columns).
   */
  secondaryAxis?: boolean;
  /**
   * think-cell axis break: compress the value range [from, to] into a small
   * band so out-of-scale columns fit. Vertical column/waterfall charts only.
   */
  axisBreak?: { from: number; to: number };
  /** Units label shown at the top of the value axis (e.g. "€m"). */
  valueAxisTitle?: string;
  /**
   * Manual label nudges by node name (e.g. {"label-0-2": {dx: 0, dy: -8}}) —
   * the pane-based stand-in for think-cell's label dragging. Applied to text
   * nodes after layout.
   */
  labelOffsets?: Record<string, { dx: number; dy: number }>;
  /** Logarithmic value axis (decade ticks). Clustered/line charts, positive data. */
  logScale?: boolean;
  /**
   * Column-family gap width, as Excel exposes it: the gap between columns as
   * a percentage of column width (0–500). 0 makes columns touch (histogram
   * look); higher values make them thinner. Default 50 (think-cell's spacing).
   */
  gapWidth?: number;
  /**
   * Clustered charts: overlap between bars within a category, as Excel
   * exposes it (−100…100). 0 places bars edge to edge; positive values
   * overlap them; negative values add a gap. Default 0.
   */
  overlap?: number;
  /**
   * Source/footnote line rendered bottom-left in small muted text
   * (e.g. "Kilde: Danmarks Statistik, 2024"). Good charts always cite
   * their source and period.
   */
  footnote?: string;
  /**
   * Pie/doughnut options: `explode` offsets the listed slice indices
   * radially; `breakout` (pie only) collapses the listed category indices
   * into one muted "Other" slice and details them in a stacked bar beside
   * the pie, joined by connector lines (Excel's bar-of-pie).
   */
  pie?: {
    explode?: number[];
    breakout?: number[];
    /** Doughnut only: render as a semi-circle (180°) gauge for scorecards. */
    semi?: boolean;
    /**
     * Variable-radius pie: each slice's angle still encodes the first series,
     * but its radius encodes a second metric — supply a `Radius` datasheet row
     * (or set this flag to use the second series). Pie only, no breakout.
     */
    variableRadius?: boolean;
  };
  /**
   * Pareto helper: sort categories by the (first non-line) series descending
   * and overlay a computed cumulative-% line on a secondary axis. Turns a
   * clustered/combo chart into the classic Pareto (80/20) view.
   */
  pareto?: boolean;
  /**
   * Small multiples: split a multi-series chart into a grid of
   * single-series panels titled by series name, sharing one value scale
   * (stacked/clustered/line/area/waterfall/radar). Special rows (Error,
   * Target, Band low/high) are carried into every panel.
   */
  multiples?: { columns?: number };
  /**
   * Boxplot options. Without precomputed Min/Q1/Median/Q3/Max rows, every
   * series row is a raw observation and quartiles/whiskers are computed.
   */
  boxplot?: {
    /** Whisker convention; auto: min/max for precomputed rows, Tukey for raw samples. */
    whiskers?: "tukey" | "minmax";
    /** Quartile method for raw samples (default "exclusive", Excel-style). */
    quartileMethod?: "exclusive" | "inclusive";
    /** Show the mean marker (default: only when a Mean row exists). */
    showMean?: boolean;
    /** Tukey fence multiplier (default 1.5). */
    iqrMultiplier?: number;
    /**
     * Overlay the raw observations as deterministically jittered dots over
     * each box (raw-sample mode only) — shows the underlying distribution
     * and sample size behind the summary.
     */
    jitter?: boolean;
    /**
     * Notch each box at the median ± 1.57·IQR/√n confidence interval
     * (raw-sample mode only, where n is known) — non-overlapping notches
     * suggest a significant difference in medians.
     */
    notch?: boolean;
    /**
     * Mean±SD box variant (raw-sample mode only): draw the box as mean ± one
     * standard deviation with the centre line at the mean and whiskers to
     * mean ± 2·SD, instead of the quartile/Tukey box. The scientific
     * "mean and spread" summary rather than the order-statistic one.
     */
    meanSd?: boolean;
  };
  /** Tile-grid map layout; omitted → auto-detected from the region codes. */
  map?: "us" | "eu" | "europe" | "world";
  /** Heatmap color options; mode "auto" picks diverging when data spans zero; totals adds sum strips. */
  heatmap?: {
    color?: string;
    negativeColor?: string;
    mode?: "sequential" | "diverging" | "auto";
    totals?: "row" | "column" | "both";
    /**
     * Mark each cell's sign (+ / −) for the reader who cannot use the color.
     * A diverging scale says "which way" in hue alone, and hue is exactly what
     * greyscale printing and red-green color blindness take away: the strongest
     * positive and the strongest negative land within ~1.1:1 of each other in
     * greyscale, i.e. the same tone.
     *
     * Drawn only where the cell has no value label — the label already prints
     * the minus sign — which is what makes this matter under `sizeEncode`,
     * where labels are suppressed entirely and color is the only sign carrier.
     * Inert on one-signed data, where every mark would be identical.
     *
     * Not "up/down": an arrow on a matrix of KPIs reads as movement against
     * last period, which is a claim this makes no attempt to check.
     */
    symbols?: "sign";
    /**
     * Calendar layout: a single daily series with date categories is laid out
     * as a weekday (row) × week (column) grid, with month labels — the
     * GitHub-contributions view. Ignored without parseable date categories.
     */
    calendar?: boolean;
    /**
     * Cell-size encoding: each cell shrinks to a centred square whose area
     * encodes the value's magnitude (colour still encodes the signed value) —
     * near-zero cells fade to dots, strong cells fill their slot (corrplot
     * style). Good for correlation / signed matrices.
     */
    sizeEncode?: boolean;
    /**
     * Hierarchical row clustering: reorder the rows by average-linkage
     * similarity and draw a dendrogram in a left gutter, so similar rows sit
     * together. Needs ≥3 rows.
     */
    cluster?: boolean;
  };
  /**
   * Column family: collapse the long tail of series into one "Other" segment,
   * keeping the `max` largest (by absolute total) and summing the rest. No-op
   * when there are already `max` or fewer series. think-cell's "Move to Other".
   */
  otherBucket?: { max?: number };
  /**
   * Tilemap options. `shape: "hex"` draws hexagonal tiles (offset rows) instead
   * of squares. `glyph: "bars"` draws a mini bar chart inside each tile from a
   * multi-series datasheet (one bar per series) instead of a single color.
   */
  tilemap?: { shape?: "square" | "hex"; glyph?: "bars" };
  /**
   * Butterfly options. `split` is the number of series on the left flank (the
   * rest go right); each flank stacks its series. Omitted → the classic
   * two-series butterfly (series 0 left, series 1 right).
   */
  butterfly?: { split?: number };
  /** Scatter/bubble options. */
  scatter?: {
    /**
     * Overlap relief: nudge markers along ONE axis so they stop covering each
     * other, leaving the other axis exact. Position on the NAMED axis becomes
     * approximate, and the chart says so in its footnote, quoting the cap. Off
     * by default. Free 2D repulsion is deliberately not offered — it would
     * corrupt both of a marker's readings at once, with nothing to bound it.
     */
    spread?: "x" | "y";
    /**
     * Hard cap on the displacement, in DATA UNITS of the `spread` axis — the
     * units the footnote quotes. Default: 2% of that axis's tick range, clamped
     * to at most 10%. A marker is never moved further than this, even if
     * overlap remains.
     */
    spreadLimit?: number;
    /**
     * Point shape per `Group`, cycled like the palette — the same channel
     * color already carries, in a form that survives greyscale printing and
     * red-green color blindness. `["circle", "diamond"]` gives group 1 circles
     * and group 2 diamonds. Off by default (every point a circle).
     *
     * Shape follows Group and nothing else: a scatter has no series to hang it
     * on (X/Y/Size/Group/Color are rows, points are categories), and encoding
     * it against anything else would claim a grouping the data doesn't state.
     * When this is set the legend draws the shapes, so the channel is always
     * explained — including under a `Color` row, where color means something
     * else and the group legend would otherwise be suppressed.
     */
    markers?: MarkerSymbol[];
    /**
     * Degree of the `Trend` row's least-squares fit: 1 (linear, default) through
     * 4 (quartic). A degree above 1 fits a polynomial and draws it as a sampled
     * curve, still reporting R². Clamped to points − 1 so it never interpolates
     * noise. Distinct from `decorations.smooth`, which visually smooths a line
     * series (Catmull-Rom) rather than regressing the scatter.
     */
    trendDegree?: number;
  };
  /** Gantt options — the timeline itself; task data stays in datasheet rows. */
  gantt?: {
    /**
     * Group tasks into lanes under a synthesized header per owner (the middle
     * part of an "Activity | Owner | Remark" category).
     *
     * Opt-in on purpose. Row order is meaningful in a plan — it is the
     * narrative — which is why `categorySort` refuses to touch a Gantt at all.
     * This reorders only because you asked, and it is a stable partition rather
     * than a sort: within a lane the rows keep the order you wrote them in, and
     * tasks with no owner stay together at the end. `After` dependencies are
     * renumbered to follow their rows.
     */
    lanes?: "owner";
    /**
     * Working-day timeline: non-working days collapse to zero width, so a bar's
     * LENGTH reads as working days rather than elapsed days (Mon→Mon is 5 units,
     * not 7). `true` = Mon–Fri; an array of ISO weekday numbers (1=Mon … 7=Sun)
     * sets a custom working week, e.g. [7,1,2,3,4] for Sun–Thu. `Holiday` rows
     * are excluded too. Weekend/holiday shading switches off with it — those
     * days have no width left to shade. Calendar timelines only (`data.dates`);
     * a no-op on a numeric one, where "working day" means nothing.
     */
    workdays?: boolean | number[];
  };
  /**
   * Radar options. `perSpoke` normalizes each spoke to its own maximum, so
   * spokes carrying different KPI units become comparable in shape (the shared
   * numeric ticks are dropped in favour of per-spoke rim maxima).
   */
  radar?: {
    perSpoke?: boolean;
    /**
     * Radial (polar) bar chart / coxcomb: draw each category as an equal-angle
     * wedge whose radius encodes its value, instead of connecting the spokes
     * into a polygon. Multi-series data stacks the wedges outward.
     */
    bars?: boolean;
    /**
     * Stacked radar: series stack cumulatively along each spoke (part-to-whole
     * across dimensions) — nested filled polygons instead of overlaid ones.
     */
    stacked?: boolean;
  };
  /** Combo: how the column series render under the lines (default stacked). */
  combo?: {
    /** Base mode under the lines (default stacked). "area" = stacked area. */
    columns?: "stacked" | "clustered" | "stacked100" | "waterfall" | "mekko" | "area";
    /**
     * "independent" gives each line series its own scale (labelled at the
     * points, no shared secondary axis) so unlike-unit KPIs all read on one
     * chart; "shared" (default) puts every line on one secondary axis.
     */
    lineAxes?: "shared" | "independent";
  };
  /** Frame size in points (PowerPoint native unit). */
  width: number;
  height: number;
  title?: string;
  decorations?: Partial<Decorations>;
  waterfall?: WaterfallOptions;
  numberFormat?: Partial<NumberFormat>;
  style?: Partial<ChartStyle>;
  /**
   * Overrides for the few labels the engine injects itself (everything else in a
   * chart is user data). Lets a localized caller pass a translated word in; the
   * engine stays language-agnostic and defaults to English.
   */
  labels?: {
    /** The auto-collapsed long-tail bucket in column/pie families. Default "Other". */
    other?: string;
  };
  /**
   * How to insert the chart:
   * - `"shapes"` (default): native, editable PowerPoint shapes — the think-cell
   *   pattern the engine was built for. Re-editable via `POWERCHART_CONFIG`.
   * - `"image"`: one raster picture per chart, sidesteps the web host's
   *   dense-shape wall (office-js #4272 / #5022 / #6498) and every
   *   `slides.add()` drop that a chart-shape's own render can't. Not editable
   *   in PowerPoint after insert; loses theme colours and blurs on rescale.
   *   The config tag round-trips onto the picture, so a future "Explode to
   *   shapes" command can re-render as native shapes on demand.
   *
   * Renderer support:
   * - Skill (`skill/scripts/render-pptx.mjs`): both modes.
   * - Office.js (`src/render/powerpoint.ts`): image mode ships in a follow-up.
   * - SVG (`src/render/svg.ts`): mode is ignored — the preview is always the
   *   vector scene.
   */
  render?: "shapes" | "image";
}

/** Geometry the decoration pass needs from a layout: where each column lives. */
export interface LayoutAnchors {
  /** Center x of each category slot. */
  categoryX: number[];
  /** Width of the column at each category. */
  categoryWidth: number[];
  /** y of the visual top of each column (min y across its segments). */
  columnTop: number[];
  /** Column total values (signed sum, or cumulative value for waterfall totals). */
  columnValue: number[];
  /**
   * Whether each category had ANY measured cell — the question `columnValue`
   * cannot answer, because every layout builds it with `?? 0`.
   *
   * A mean over `[100, null, 200]` was (100 + 0 + 200) / 3 = 100, and the line
   * was drawn and labelled at a number nobody measured; the true mean of what
   * was measured is 150. A blank is not a zero to average in, and it is not a
   * third of the denominator either.
   *
   * OPTIONAL, and absent means "every category counts" — the honest default for
   * a layout whose `columnValue` has no blank to coerce (a waterfall's running
   * level carries across an empty step). It was also written for "a boxplot's
   * median, a candlestick's close", and that was false: a category with no box
   * or no Close is ordinary input, and both layouts answered 0 for it. Every
   * layout that builds `columnValue` with `?? 0` has to answer, because the
   * readers are not limited to the column family: `buildChart` draws
   * `valueLines` on every kind that has a value map, and the difference and CAGR
   * arrows read the same array. The line family (line, area, sparkline, slope,
   * bump), candlestick, boxplot, violin and mekko answer alongside the columns.
   */
  columnHasData?: boolean[];
  /**
   * Per category, per series IN THE CONFIG'S INDEXING (the one a decoration's
   * `series` is written against): where that series tops out, in value units —
   * the level a level arrow, a series CAGR or a callout anchors on. Stacked
   * (columns, area) that is the cumulative value after the series; everywhere a
   * series stands on the baseline (clustered, line) it is the series' own value.
   *
   * NaN: the series has no place on this chart's value scale (a combo's line on
   * its own axis), so the decoration cannot be anchored on it. Absent: this
   * layout publishes no levels, and a series-qualified decoration compares the
   * category totals instead.
   */
  seriesLevels?: number[][];
  /**
   * Where each series' right-margin name was ANCHORED — its last segment's
   * midpoint — before `seriesLabelNodes` spread them apart. Absent when the
   * chart draws no such names.
   *
   * Exposed for the combo, which cannot reconstruct it. A combo's line names and
   * its column names share one gutter and are emitted by different code at
   * different times: the columns are spread to fill the gutter before the lines
   * exist, the lines are added at their own points afterwards, and the global
   * de-collision pass then inherits ten labels in a strip sized for seven with
   * only "up" to work with. Laying both out together is the fix, and it needs
   * the anchors the first pass started from — not the positions it settled on.
   */
  seriesLabelAnchors?: (number | null)[];
  /** y coordinate of the zero baseline. */
  baselineY: number;
  /** Plot rectangle. */
  plot: { x: number; y: number; w: number; h: number };
  /** Value → y mapping used by the layout (absent for 100% charts). */
  valueToY?: (v: number) => number;
  /**
   * Value → x mapping, for horizontal (bar) orientation — where the value axis
   * runs left-to-right and `categoryX` holds the category's Y centre. Absent
   * for vertical charts and for 100% bars.
   */
  valueToX?: (v: number) => number;
}
