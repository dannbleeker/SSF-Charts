# ChartConfig reference

Everything the SSF Charts engine accepts. All lengths in points (1pt = 1/72").

```ts
{
  kind: "stacked" | "clustered" | "stacked100" | "waterfall" | "mekko"
      | "line" | "area" | "butterfly" | "scatter" | "bubble" | "gantt"
      | "combo" | "pie" | "doughnut" | "boxplot" | "radar" | "heatmap"
      | "tilemap" | "cascade" | "funnel" | "waffle" | "treemap"
      | "sunburst" | "violin" | "candlestick",
  width?: 480, height?: 300,          // frame size in pt
  title?: string,
  horizontal?: boolean,               // rotate columns/waterfall/mekko/boxplot into bars/rows;
                                      // also line/area → a "profile chart" (categories down, values right)
  data: {
    categories: string[],             // column headers (x categories / activities / points)
    series: [{
      name: string,
      values: (number | null)[],      // one per category; null = blank
      color?: "#rrggbb",              // per-series override
      colors?: ("#rrggbb"|null)[],    // per-CELL override: highlight one segment/point/slice
      pattern?: "diagonal"|"crosshatch"|"dots"|"horizontal",
                                      // hatch over the fill (SVG/preview; solid in PPT output)
      type?: "line" | "marker",       // combo: draw this series over the columns as a line, or as
                                      // bare unconnected points ("marker") — a per-category
                                      // benchmark/target/consensus, where a line would imply the
                                      // values interpolate. A marker series shares the overlay
                                      // scale, so it only means what it looks like on the columns'
                                      // own scale: not with secondaryAxis or lineAxes:"independent",
                                      // and pareto keeps only the bars + the cumulative line.
      stack?: number,                 // clustered-stacked: stack group index
      scenario?: "AC"|"PY"|"PL"|"BU"|"FC"  // IBCS notation (column family): the fill encodes the
                                      // data's nature — AC solid, PY lighter, PL/BU outlined/hollow,
                                      // FC hatched; the code is appended to the legend/label. Set the
                                      // same `color` on each scenario series for the single-hue look.
    }],
    hundredPercent?: (number|null)[], // "100%=" row: denominators for 100% charts
    xExtent?: (number|null)[],        // "X extent" row: Mekko-with-units widths
    dates?: true                      // set when values are days-since-epoch (Gantt)
  },
  decorations?: {
    segmentLabels?: boolean,          // default true, auto-hidden when too small
    seriesLabels?: boolean,           // default true (right of last column / legend)
    totals?: boolean,                 // column totals (one per stack on a clustered-stacked chart, sideways too)
    grandTotal?: boolean,             // sum of all category totals, top-right (stacked/clustered columns)
    variance?: { actual: number, reference: number, mode?: "absolute"|"percent", goodIsUp?: boolean },
                                      // IBCS variance tier below the columns: signed Δ (or Δ%) of the
                                      // actual series vs a reference (plan/PY), green favourable / red not
    categoryAxis?: boolean,           // default true
    valueAxis?: boolean | "datamarks", gridlines?: boolean,  // default false;
                                      // "datamarks" = Tufte ticks-only axis, no line
    tickMode?: "nice" | "data",       // datamarks at round values or at data min/max
    tightLabelPriority?: "columns" | "line",  // combo: on a frame too small for both rows of
                                      // value labels, which survives. "columns" (default) keeps
                                      // the totals, "line" keeps the point labels. Only labels
                                      // still colliding after de-collision are dropped, so a
                                      // chart with room keeps both.
    gridShape?: "polygon" | "circle", // radar grid web (default polygon)
    fillOpacity?: number,             // radar series fill (default 0.18)
    labelContent?: ("value"|"percent"|"series"|"category")[],
    cagr?: { from: number, to: number, series?: number },        // category indices
    difference?: { from, to, percent?, series?, fromValueLine? },
    valueLines?: ({mode:"mean"} | {mode:"value", value:number})[],
    connectors?: boolean,             // lines joining stacked-segment boundaries between columns
    callouts?: [{ text, category, series?, dx?, dy? }],  // speech-bubble comments on a value
                                      // `series` (cagr / difference / callouts / variance) indexes
                                      // data.series as written, Error / Target rows counted. A blank
                                      // end draws no arrow, and the mean skips blank categories.
    bands?: [{ axis:"x"|"y", from, to, color?, label? }], // shaded background region
                                      // (y = value range; x = category indices;
                                      //  scatter/bubble: both axes in value units)
    hundredPercentNote?: boolean,     // "100% = N" note (pie/doughnut/stacked100)
    forecastFrom?: number,            // line: dashed segments + hollow markers from this index
    barStyle?: "bar"|"lollipop"|"dot"|"range",  // clustered: stems+dots / dots / dumbbell
    fillBetween?: [number, number],   // line: shade the gap between two series (plan vs actual)
    stepped?: "before"|"after"|"center",  // line/area: staircase segments (jump at start / end / midpoint)
    smooth?: boolean,                 // line: smooth Catmull-Rom curves (sampled polyline)
    bridgeGaps?: boolean,             // line: connect straight across null points instead of breaking
    slope?: boolean,                  // line: slope-chart mode — end rails + "Name value" labels, no axis
    bump?: boolean,                   // line: bump chart — ranks on an inverted axis, thick lines, end names
    sparkline?: boolean,              // line/area: word-sized trend, no chrome, min/max/last dots (pair with multiples)
    trajectory?: boolean,             // scatter/bubble: connect points in row order with a direction trail
    summaryBars?: boolean,            // gantt: summary bar on section rows (spans children min→max)
    criticalPath?: boolean,           // gantt: red-outline the longest "After"-dependency chain + its arrows
    radarBand?: boolean,              // radar: shade the peer min–max envelope, draw last series on top
    quadrants?: { x, y, labels? }     // scatter: 4 tinted zones + corner labels at one crossing
    marginals?: "x"|"y"|"both",       // scatter/bubble: distribution histograms in a top/right gutter
                                      // (bins subdivide the axis ticks, so every tick is a bin edge;
                                      //  the plot shrinks; dropped if that would leave too little)
  },
  footnote?: string,                  // source line, bottom-left ("Source: …, 2024")
  pie?: { explode?: number[],         // slice indices offset radially to highlight
          breakout?: number[],        // pie only: collapse these categories into one "Other"
                                      // slice detailed in a stacked bar beside the pie (bar-of-pie)
          semi?: boolean,             // doughnut only: 180° semi-circle gauge (scorecard)
          variableRadius?: boolean }, // pie: angle = 1st series, radius = a "Radius" row / 2nd series
  waterfall?: { totalIndices?: number[],   // categories drawn as cumulative totals (the "e" cell)
                spacerIndices?: number[],  // blank gap columns; the running total carries across
                detailGroups?: [{ of: number, indices: number[] }] },
                                      // "of which": `indices` decompose column `of`'s delta as a
                                      // sub-bridge from that column's own base. Off the chain — they
                                      // carry no running total, so every downstream total is
                                      // unaffected — and the dashed connector steps over the group.
                                      // The parent is named, never inferred. A group that does not
                                      // sum to the parent's delta renders as authored.
  pareto?: boolean,                   // clustered/combo: sort desc + cumulative-% line (80/20 view)
  scatter?: { spread?: "x"|"y",       // scatter/bubble: nudge markers along ONE axis to relieve overlap;
                                      // the other axis stays exact. The cap is printed in the footnote.
                                      // Ignored under decorations.quadrants (a nudge must not move a
                                      // point across a quadrant line). No free 2D repulsion: it would
                                      // corrupt both of a marker's readings at once.
              spreadLimit?: number,   // hard cap, in DATA UNITS of the spread axis (default 2% of its
                                      // tick range, max 10%)
              markers?: ("circle"|"square"|"diamond"|"triangle"|"plus")[],
              trendDegree?: number },  // Trend row fit: 1 linear (default) … 4 quartic; a polynomial
                                      // above 1 draws as a sampled curve, clamped to points−1, R² stated
                                      // point shape per Group row, cycled like the palette; the legend
                                      // draws the shapes. Off => every point a circle. Shape is a
                                      // channel color cannot carry: it survives greyscale printing and
                                      // red-green color blindness. Area is equalised across shapes, so
                                      // a bubble's "area proportional to Size" still holds.
  multiples?: { columns?: number },   // small multiples: one single-series panel per series,
                                      // shared value scale (stacked/clustered/line/area/waterfall/radar)
  boxplot?: { whiskers?: "tukey"|"minmax", quartileMethod?: "exclusive"|"inclusive",
              showMean?: boolean, iqrMultiplier?: number,
              jitter?: boolean,       // overlay raw observations as jittered dots (raw-sample mode)
              notch?: boolean,        // notch each box at the median CI (raw-sample mode)
              meanSd?: boolean },     // box = mean±SD, centre = mean, whiskers = mean±2SD (raw-sample mode)
  map?: "us" | "eu" | "europe" | "world",   // tilemap layout (us/eu/europe are auto-detected if omitted; world never is)
  tilemap?: { shape?: "square"|"hex",       // hexagonal tiles (offset rows) instead of squares
              glyph?: "bars" },   // mini bar chart per region from a multi-series datasheet
  heatmap?: { color?, negativeColor?, mode?: "sequential"|"diverging"|"auto",
              totals?: "row"|"column"|"both",    // marginal sum strips
              calendar?: boolean,     // weekday × week grid for a daily date series
              sizeEncode?: boolean,   // cell area encodes |value| (corrplot style); colour = signed value
              cluster?: boolean,      // reorder rows by similarity + draw a dendrogram (≥3 rows)
              symbols?: "sign" },     // mark each cell +/− for readers who cannot use the colour. A
                                      // diverging scale says "which way" in hue alone, and greyscale
                                      // printing collapses its strongest + and − to the same tone.
                                      // Only where a cell has no value label (the label already prints
                                      // the sign) — so it matters most under sizeEncode and on calendar
                                      // heatmaps, neither of which draws labels. Exactly-zero gets no
                                      // mark (no sign to state). Inert on one-signed data.
  otherBucket?: { max?: number },     // column family: collapse the long tail into one "Other" (keep max series)
  butterfly?: { split?: number },     // butterfly: series on the left flank (rest stack right)
  radar?: { perSpoke?: boolean,       // radar: normalise each spoke to its own max (mixed KPI units)
            bars?: boolean,           // radial bar chart / coxcomb: category wedges, radius = value (stacks)
            stacked?: boolean },      // stacked radar: series stack cumulatively along each spoke
  combo?: { columns?: "stacked"|"clustered"|"stacked100"|"waterfall"|"mekko"|"area",  // base under the lines
            lineAxes?: "shared"|"independent" },  // "independent": each line its own scale + labels
  // (waterfall? is documented once above, with detailGroups)
  scale?: { min?: number, max?: number },   // pin the value axis
  axisBreak?: { from: number, to: number }, // compress an out-of-scale range
  logScale?: boolean,                       // clustered/line, positive data
  gapWidth?: number,                        // column family: Excel gap width 0–500 (% of column width; default 50)
  overlap?: number,                         // clustered: Excel bar overlap −100…100 (default 0 = edge to edge)
  valueAxisTitle?: string,                  // units label, e.g. "€m"
  segmentOrder?: "sheet"|"reverse"|"ascending"|"descending",
  categorySort?: "ascending"|"descending",  // by column total; a pie/doughnut by its slices, the first series (not waterfall/gantt)
  secondaryAxis?: boolean,                  // combo: right-hand axis for line series
  labelOffsets?: { [nodeName]: {dx, dy} },  // manual label nudges
  numberFormat?: { decimals?: number|"auto", suffix?: string,
                   forceSign?: boolean, locale?: "de-DE"|... },
  style?: { palette?: string[], fontSize?, negative?, neutral?,   // series ink
            text?, mutedText?, axis?, gridline?, background? },   // canvas ink
  render?: "shapes"|"image"                 // "shapes" (default): native, editable PPT shapes.
                                            // "image": one raster picture per chart —
                                            //   sidesteps the PowerPoint-web dense-shape wall,
                                            //   but the chart is no longer editable.
                                            //   Needs @resvg/resvg-js installed.
}
```

## Special series-name conventions (the "datasheet rows")

| Row name | Effect |
|---|---|
| `100%=` | per-category denominator for `stacked100` (columns can stay short of 100%) |
| `X extent` | Mekko-with-units: explicit column widths |
| `X`, `Y`, `Size`, `Group` | scatter/bubble coordinates, bubble area (a negative `Size` draws as an outline at the size of its magnitude), point coloring (1..k) |
| `Color` | scatter/bubble: numeric third variable on a sequential color ramp (with a gradient legend); supersedes `Group` coloring |
| `X line`, `Y line` | scatter partition lines at those values |
| `Trend` | any value present → OLS trend line, labelled with its R² and p-value |
| `Start`, `End`, `Milestone` | Gantt bars & milestone markers (numbers or day values) |
| `% Complete` | Gantt: darker progress fill inside the bar (0-100, or 0-1 when no value in the row exceeds 1 — one unit per row) |
| `Baseline start`, `Baseline end` | Gantt: thin ghost bars showing the original plan vs actual |
| `After` | Gantt dependency: 1-based predecessor index → elbow arrow |
| `Today` | Gantt today line at the (single) value |
| `Holiday` | Gantt: shade these dates |
| `Bracket <label>` | Gantt: interval annotation spanning min→max of the row's values |
| `Column <label>` | Gantt: numeric table column beside the task labels (right-aligned, own precision; unit goes in the label) |

`gantt: { lanes?: "owner" }` — regroup tasks under one synthesized header per
Owner (the middle part of an `Activity | Owner | Remark` category). Opt-in: a
plan's row order is the plan, so nothing regroups on its own. A stable
partition rather than a sort — rows keep their order inside a lane, unowned
tasks stay together at the end under a header of their own (`Other`; a
localized word goes in `labels: { other }`), so they are never read as the last
owner's tasks or counted by that owner's summary bar — and `After` is
renumbered so the dependency arrows still join the same tasks.

`gantt: { workdays?: boolean | number[] }` — working-day timeline: non-working
days collapse to zero width so a bar's length reads as working days, not
elapsed days (Mon→Mon = 5 units, not 7). `true` = Mon–Fri; an array of ISO
weekday numbers (1=Mon … 7=Sun) sets a custom week, e.g. `[7,1,2,3,4]` for
Sun–Thu. `Holiday` rows drop out of the scale too, and weekend/holiday shading
switches off with it. Calendar timelines only; a no-op on a numeric one.

| `Error`, `Error+`, `Error-` | error bars at the column total / line point (± or asymmetric); stacked/clustered/line/area, either orientation |
| `Target` | bullet tick across each column or bar (set `horizontal: true` for a conventional bullet graph — `decorations.bands` follow); on waterfalls also a hatched gap-to-target segment + label |
| `Band low`, `Band high` | line charts: shaded confidence/uncertainty ribbon (rows never draw as lines) |
| `Min`, `Q1`, `Median`, `Q3`, `Max` | boxplot five-number summary (whiskers to Min/Max, think-cell style); suffix "\| group" ("Min \| 2024") for side-by-side grouped boxes |
| `Mean` | boxplot mean marker (×) |
| `Open`, `High`, `Low`, `Close` | candlestick OHLC prices per period (category) |
| `Outlier <n>` | boxplot: extra outlier dots in precomputed mode |

Gantt category conventions: `"Activity | Owner | Remark"` adds responsible and
remark columns; a leading `">"` indents; a category with no Start/End/Milestone
renders as a bold section-header band.

**Dates**: for Gantt rows, pass days-since-epoch numbers and set `data.dates:
true` (compute as `Math.round(Date.parse("2026-01-15")/86400000)`), or let the
interactive datasheet parse ISO strings. Calendar timelines pick weeks (with
weekend shading), months, or quarters by span automatically.

**Waterfall semantics**: values are deltas; `waterfall.totalIndices` lists
categories drawn from the baseline to the running total (think-cell's `e`).
Multiple series → stacked waterfall (contributions stack per column).

**Clustered-stacked**: give series `stack: 0`, `stack: 1`, … — same index
stacks together; different indices sit side by side per category.

**Boxplot without summary rows**: a group (the rows sharing a `| group`
suffix, or all of them) is read as a five-number summary only when it has all
five of `Min`/`Q1`/`Median`/`Q3`/`Max`. Otherwise every series row in it is a
raw observation — quartiles are computed (exclusive method) and whiskers use
Tukey 1.5×IQR fences with outliers drawn as dots — so a person called Max, or
quarterly samples named `Q1`…`Q4`, stay observations.

**Funnel**: one series of stage values (order ascending for a pyramid);
centered bands with width ∝ value, conversion % vs the previous stage in
the gaps, stage names on the left.

**Waffle**: a 10×10 unit grid (1 cell = 1%), filling from the bottom-left.
Categories are the parts (pie semantics, first series); shares round to
whole cells by largest remainder and unassigned cells stay gray. The
denominator is the `100%=` row when present, else the value sum — except a
single category ≤ 100 with no denominator, which reads as a literal
percentage (`68` → 68 cells + a big "68%" beside the grid).

**Stepped line/area** (`decorations.stepped`): draw values as a staircase
instead of sloped segments — `"after"` holds each value then jumps at the
next category (HV), `"before"` jumps immediately (VH), `"center"` steps at
the midpoint (HVH). Works for `line` and `area`; area fills follow the steps.

**Area with negative values**: `area` charts stack positives above the zero
baseline and negatives below it, so a series that goes negative (e.g. net
income over time) dips under the axis instead of being clamped to 0.

**Scatter/bubble trajectory** (`decorations.trajectory`): connects the points
in datasheet (row) order with a trail and a direction arrowhead at each
segment midpoint — a Gapminder-style path of one entity through the X/Y
space over time. Order the categories chronologically.

**Scatter/bubble continuous color** (a `Color` row): maps each point's numeric
value onto a sequential color ramp and swaps the group chip legend for a
gradient legend with min/max labels. Encodes a third variable (or a fourth,
on bubbles alongside `Size`). Supersedes `Group` coloring when both exist.

**Smoothed lines** (`decorations.smooth`): draws smooth Catmull-Rom curves
through the points instead of straight segments (sampled to a dense polyline
since the renderers have no freeform paths). Ignored when `stepped` is set.

**Waterfall spacers** (`waterfall.spacerIndices`): blank categories that draw
no bar and leave a gap, grouping a long bridge into sections. The running
total carries across and the dashed connector bridges the gap — give the
spacer category an empty name so no axis label shows.

**Boxplot jitter** (`boxplot.jitter`): in raw-sample mode, overlays every
observation as a deterministically jittered dot over its box, so the reader
sees the distribution and sample size behind the summary (outlier dots are
subsumed by the jittered points).

**Boxplot notch** (`boxplot.notch`): in raw-sample mode, pinches each box at
the median ± 1.57·IQR/√n confidence interval — boxes whose notches don't
overlap have significantly different medians. Precomputed boxes (no sample
size) stay rectangular.

**Mean±SD box** (`boxplot.meanSd`): in raw-sample mode, draws each box as
mean ± 1 standard deviation with the centre line at the mean and whiskers to
mean ± 2·SD — the scientific "mean and spread" summary rather than the
quartile/order-statistic one. No mean × marker (the centre line is the mean);
mutually exclusive with `notch`.

**Gantt summary bars** (`decorations.summaryBars`): draws a capped summary bar
on each section-header row (a category with no Start/End/Milestone), spanning
min(start)→max(end) of the activities beneath it up to the next header.

**Gantt critical path** (`decorations.criticalPath`): over the `After`
dependency edges, finds the chain with the greatest cumulative duration (the
critical path) and highlights its bars with a red outline and its dependency
arrows in thicker red — the MS-Project convention for "the tasks that decide
the finish date". No-op without `After` rows.

**Sparklines** (`decorations.sparkline`, line/area): a compact, axis-less,
word-sized trend line with no chrome — a thin line, an optional leading label
(the title / series name), a trailing value, and dots on the min (red), max
(green) and last points. Pair with `multiples: { columns: 1 }` for a table of
sparklines, one row per series (KPI dashboards) — `multiples: {}` lays three
series side by side and four as a 2x2. Each row scales to its own values, so
rows of unlike units read side by side; pin `scale` to put them all on one.

**Radar min–max band** (`decorations.radarBand`): shades the per-spoke min–max
envelope of the peer series (all series except the last) as a band and draws
the last series prominently on top — the "peer range + us" competitive
profile. The legend collapses the peers into one "Peer range" swatch.

**Other bucket** (`otherBucket.max`): for stacked/clustered/100% charts with
many series, keep the `max` largest (by absolute total) and sum the rest into
one trailing "Other" segment — think-cell's "Move to Other Series". No-op when
there are already `max` or fewer series.

**Calendar heatmap** (`heatmap.calendar`): a single daily series with date
categories (strings that read as dates — `2026-01-15`, `15.01.2026`, `Jan 2026`)
is laid out as a weekday (row) × week (column) grid with month labels — the
GitHub-contributions view. A bare day number is NOT a date here, `data.dates`
or not (that flag belongs to Gantt rows; the heatmap never reads it), so it
falls back to the plain matrix layout — as does any set with a category that
is not a date, with no warning. Only the first series is drawn, and the colour
scale follows it alone.

**Butterfly stacked flanks** (`butterfly.split`): the first `split` series
stack on the left flank and the rest stack on the right (a legend replaces the
two headers; it wraps onto more rows when the names do not fit one, and is left
out when the room for it would leave the bars no plot). Omit for the classic
two-series butterfly (series 0 vs series 1).

**Radar per-spoke scales** (`radar.perSpoke`): normalise each spoke to its own
maximum so spokes carrying different KPI units become comparable in shape; the
shared numeric ticks give way to fraction rings (25/50/75/100 %).

**Radial bar chart / coxcomb** (`radar.bars`): instead of connecting the spokes
into a polygon, draw each category as an equal-angle wedge whose radius encodes
its value (a Nightingale rose), from a small inner hole with concentric value
rings. A single series colours bars by category; multiple series stack outward
within each sector.

**Stacked radar** (`radar.stacked`): series stack cumulatively along each spoke
(part-to-whole across dimensions) — nested filled bands instead of overlaid
polygons; the scale reaches the per-spoke sums.

**Variable-radius pie** (`pie.variableRadius`, or add a `Radius` datasheet row):
each slice's angle still encodes the first series while its radius encodes a
second metric — a two-variable pie. Labels sit outside; pie only (no doughnut,
no breakout).

**Heatmap cell-size encoding** (`heatmap.sizeEncode`): each cell shrinks to a
centred square whose area encodes the value's magnitude (colour still encodes
the signed value) — near-zero cells fade to dots, strong cells fill their slot.
The corrplot view for correlation / signed matrices.

**Heatmap row clustering** (`heatmap.cluster`): reorder the rows by
average-linkage similarity (Euclidean over each row's values) so similar rows
sit together, and draw a dendrogram in a left gutter. Needs ≥3 rows.

**Combo stacked-area base** (`combo.columns: "area"`): the base under the line
series is a stacked area instead of columns — trend-of-mix plus a KPI line
(pair with `secondaryAxis` for a %-on-the-right line).

**Missing-data bridge** (`decorations.bridgeGaps`): line charts connect
straight across null categories instead of breaking into separate segments
(applies to plain straight lines, not `smooth`/`stepped`).

**Floating segments**: give a stacked-column series `color: "transparent"` and
it occupies the stack without drawing — the segments above it "float" clear of
the baseline (a lightweight floating-bar / range effect).

**Combo base modes** (`combo.columns`): the column base under the lines can be
`waterfall` (a bridge with a %-of-total line over it — set `waterfall.totalIndices`)
or `mekko` (variable-width columns with an overlaid line) as well as the
stacked / clustered / 100% modes. The lines use a secondary axis (forced for
mekko / 100%, which have no shared value axis).

**Combo independent line axes** (`combo.lineAxes: "independent"`): each line
series is zoomed to its own value range and labelled at every point, so several
KPIs in unlike units read on one chart without a shared secondary axis.

**Tilemap hex tiles** (`tilemap.shape: "hex"`): draws the cartogram with
hexagonal tiles (odd rows offset, nested rows) instead of squares — filled in
SVG/pptx, outlines in the live add-in (polygons have no freeform fill there).

**Tilemap mini-glyphs** (`tilemap.glyph: "bars"`): with a multi-series
datasheet (categories = regions, series = periods), each tile draws a mini bar
chart of that region's values with a series legend, instead of a single value
color. The legend wraps onto more rows when the names are wider than the chart,
and is left out when the room for it would cost the tiles their bars.

**100% charts with negatives**: `stacked100` no longer clamps negatives — a
negative segment renders below the zero line as its share of the (positive)
column total, so returns / adjustments stay visible.

**Semi-circle gauge** (`kind: "doughnut"` + `pie.semi: true`): a half-doughnut
scorecard — categories fill a 180° arc across the top with the total in the
open centre.

**Pareto** (`pareto: true` on a clustered/combo chart): sorts the categories by
the (first non-line) series descending and overlays a computed "Cumulative %"
line on a secondary axis — the 80/20 view. Rewrites the chart into a combo.

**Bump chart** (`kind: "line"` + `decorations.bump: true`): values are ranks
(1 = best) drawn on an inverted integer axis (rank 1 at the top; a ranking that
starts lower, such as 0-based, puts its lowest rank there) with thick lines,
round markers and a "Name" label at both ends — rank-over-time.

**Horizontal profile chart** (`kind: "line" | "area"` + `horizontal: true`):
rotates the line/area 90° — categories run down the left axis and values
extend to the right. Stacked areas stack outward to the right.

**Sunburst** (`kind: "sunburst"`): a two-ring hierarchical pie — categories
named `"Group | Item"` put groups on the inner ring and their items on the
outer ring; without `"|"` it degrades to a single-ring doughnut.

**Violin** (`kind: "violin"`): a mirrored kernel-density shape per category
from raw samples (every datasheet row is an observation, like the boxplot's
raw mode) with a median tick. The body is a filled polygon — solid in
SVG/pptx, outline-only in the live add-in.

**Candlestick** (`kind: "candlestick"`): OHLC financial bars from rows named
`Open` / `High` / `Low` / `Close`; a high–low wick carries a body from open to
close, green when the period rose and red when it fell.

**Gap width & overlap** (`gapWidth`, `overlap`): mirror Excel's two spacing
controls for the column family. `gapWidth` (0–500, default 50) is the gap
between columns as a percentage of column width — 0 makes them touch
(histogram look), higher makes them thinner. `overlap` (−100…100, default 0)
sets how much clustered bars within a category overlap: positive overlaps
them, negative opens a gap.

**Slope chart**: `kind: "line"` + `decorations.slope: true` with (ideally)
two categories — the before/after comparison. No value axis; vertical
rails carry the two periods and every series gets a colored "Name value"
label at both ends (labels de-overlap automatically).

**Bar-of-pie** (`pie.breakout`): the listed category indices collapse into
one muted "Other" slice (rotated to face 3 o'clock) and are detailed in a
stacked bar on the right with dashed connectors; detail labels show the
share of the grand total.

**Small multiples** (`multiples`): a multi-series chart becomes a grid of
single-series panels titled by series name, pinned to one shared value
scale so panels compare honestly (a table of sparklines is the exception: each
row scales to itself unless `scale` is pinned). Special rows (Error, Target, Band
low/high) are carried into every panel. `columns` overrides the grid
(default: one row up to 3 panels, else a near-square).

**Cascade** (decomposition): each stage's bar is a subset of the previous
one, top-aligned on one volume scale; the complement hangs as a muted
labeled box at each split. Category syntax: `"Stage | Drop label | Group
header"` — part 2 captions the remainder of the split into this stage,
consecutive stages sharing part 3 get one spanning header band.
series[0] = kept counts per stage.

**Radar**: categories = spokes (first at 12 o'clock, clockwise), series =
translucent polygons; keep 5–10 spokes and ≤3 series. With no `scale` the rings
follow the data from zero (or from the lowest negative value) up. **Heatmap**: series =
rows, categories = columns; one global color scale (diverging through white
when the data spans zero). **Tilemap**: categories = region codes (US postal,
ISO-2 for eu/europe, or NA/LATAM/EU/MEA/SSA/CIS/SA/EA/SEA/OCE for world) —
`us`, `eu` and `europe` are auto-detected from the codes when at least 90% of
them match one layout, or pinned with `map:`. **`world` is never
auto-detected** (its macro codes collide with ISO-2): always set
`map: "world"`. A set that matches no layout draws no tiles, only the hint
"No recognized region codes — set map: …".

## Chart-design formalia (baked-in defaults, per "the good chart" practice)

Every chart should carry: a message title, a `footnote` citing source and
period, sorted categories where order is free, labels on the data instead of
a value axis, and — on percentage charts — `hundredPercentNote` so readers
know what 100% is. Use `decorations.connectors` on stacked charts to make
segment development followable, `series.colors` to draw the eye to one data
point, `callouts` to comment on it, and `bands` to mark a reference region.

## Palette

Default 8-slot categorical palette (CVD-validated): `#2a78d6 #1baf7a #eda100
#008300 #4a3aa7 #e34948 #e87ba4 #eb6834`. Waterfall decreases use `#e34948`,
totals `#898781`. Override via `style.palette` or per-series `color`.

## Output

`scripts/render-pptx.mjs` emits one 13.33×7.5in slide per config with the chart
centered, every bar/label/line as a native PowerPoint shape, and exact pie
geometry (each wedge is a custom-geometry polygon sampled from the same polar
maths as the preview; the OOXML pie preset is not used, so a wedge has no
adjustment handles). The slide takes the chart's own `style.background`
(white by default, and white for a background that draws nothing — `transparent`,
`rgba(0,0,0,0)`, `#0000` — or that cannot be read; an agenda is always white), so
a dark-styled config stays readable.
`scripts/render-svg.mjs` emits SVGs for fast visual QA, on the chart's own
`style.background` too. Both accept an
`"agenda"` config and isolate a failing one, so a single bad entry cannot throw
away the rest of the batch.

Each chart lands as **one grouped object carrying its own config**, so a deck
rendered here opens in the SSF Charts task pane exactly like one the pane drew:
select a chart, click *Edit selected chart*, change the data, update it in
place. The grouping and the `POWERCHART_CONFIG` tag are injected into the OOXML
after pptxgenjs is done — it can express neither — by the same code the add-in
runs, so the two outputs cannot drift apart. Image-mode charts get the tag on
the picture, so they are re-editable too. The tag part also records the chart's
frame origin, so the first edit in the pane redraws the chart where it stands
(that first edit is a full redraw, which is what stamps the fingerprint later
edits patch against). An `"agenda"` slide is the exception: loose shapes, no
group and no tag, exactly as the pane's *Insert agenda slides* draws one — it is
text, not a chart, and there is no data in it to edit. It is always drawn as
native shapes, whatever `render` says.

The chart's text alternative travels too: the same one-line summary the SVG
puts in `<desc>` is written as the group's alt text, so PowerPoint shows it
under *Edit Alt Text* and a screen reader announces it. Image-mode charts carry
it on the picture. (pptxgenjs itself exposes alt text on pictures and native
charts only — the group's comes from the same OOXML post-pass that writes the
config tag.)

One known limit of the headless pptx: the deck font is Segoe UI (no config
field selects it — set the font in the deck's theme after inserting).
