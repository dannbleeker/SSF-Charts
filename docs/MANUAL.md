# SSF Charts User Manual

SSF Charts makes think-cell-style charts in PowerPoint as **native, editable
shapes** — every bar, label, and line is a real PowerPoint object you can
tweak after insertion. This manual covers the task pane; for driving
SSF Charts from JSON or from Claude, see the [automation](#automation-json)
section and `skill/reference.md`.

## Getting started

1. Sideload the add-in: in PowerPoint, **Insert → Add-ins → Upload My
   Add-in** → pick `manifest.xml` (the dev server must be running:
   `npm run dev`). The **SSF Charts** ribbon group has an **Insert chart**
   button and menu (direct entries for the most common chart types) plus an
   **Insert element** menu (Harvey ball, checkbox, process flow, KPI tile,
   table) that opens the pane on the Elements tab with that element highlighted.
2. Requirements: PowerPoint on Windows 2207+, Mac 16.62+, or the web
   (PowerPointApi 1.4+). Grouping needs 1.8+, chart re-editing (tags) 1.3+,
   pies 1.10+ (shape rotation), deck-theme colors 1.10+. Marker symbols need
   only 1.4 — they are preset geometry, not rotated shapes. Most missing
   capabilities degrade gracefully: grouping, tags and theme colours are simply
   skipped and the chart still inserts.

   **The pie family is the exception, and it is a hard one.** Pie, doughnut,
   sunburst and gauge are built from rotated shapes, and there is no way to
   build them without rotation. Below **1.10** they insert with their title,
   legend and labels and **no slices at all**. If you are on a build without
   1.10 — most desktop Office before 2026, and volume-licensed/LTSC builds —
   use a bar or column chart instead. The pane records the refusal in its trace.
3. No host? The pane also runs as a plain web page (`npm run dev` →
   `localhost:3000`) with everything except insertion, plus a demo gallery.

## Inserting and editing a chart

The pane is organised into four **tabs** — **Chart** (the main flow), **Elements**,
**Agenda**, and **Automation** — and the **Insert / Update** button is pinned to a
bar at the bottom so it's always reachable. Within the Chart tab, *Chart type*,
*Decorations* and *Preview & size* are collapsible sections (click the header);
picking a type auto-collapses the gallery to a compact "current type" summary.

1. **Pick a chart type** from the gallery — the types are grouped by family
   (columns & bars, line & area, parts of a whole, distribution, correlation,
   matrix & spatial), with a search box to filter by name; thumbnails are live
   previews.
2. **Type or paste your data** into the datasheet — row 1 is categories,
   column A is series names. Pasting a range from Excel (Ctrl+V) works.
3. Adjust options; the preview updates live.
4. **Insert into slide.** If a placeholder or shape is selected, the chart
   fills its bounds. Otherwise it is placed clear of whatever is already on the
   slide — **beside** it or **below** it, whichever leaves the bigger chart,
   scaled down to fit the space left if it will not go at full size. The pane
   says which it did.

   Beside usually wins on a 16:9 deck and usually loses on 4:3, and the
   difference is bigger than it sounds: a second default 480×300 chart gets a
   390×244 slot next to the first on 16:9, against 192×120 underneath it. On a
   4:3 deck there is only 150pt of width spare, which would squeeze the chart
   below the point of being readable, so it goes underneath instead. Your
   deck's real slide size is read from PowerPoint, not assumed, so both cases
   come out right.

   A slide with room in neither direction falls back to a small cascading
   offset, so the chart lands somewhere you can see and drag rather than off
   the edge.

   **If the slide is already busy, the pane may offer the chart a slide of its
   own.** Adding to a loaded slide is several times slower than adding to an
   empty one — the cost is PowerPoint's, and it grows with what is already
   there — so when the wait looks long the pane says how long it will take
   here, what it would take on a new slide, and lets you choose. Both figures
   are estimates from measured inserts, and neither is called instant: a big
   chart takes a while wherever it goes.

   You only see the offer when moving the chart would at least halve the wait.
   A chart that is slow because it is *large* is just as slow on an empty
   slide, and there is no point offering a slide that would not help. If you
   take the offer and PowerPoint fails to add the slide, the chart goes where
   you originally asked and the pane says so — you never lose a chart to a
   slide that did not appear.
5. **Re-edit later**: select the chart on the slide — the pane shows an
   "Edit it" banner. Or click **Edit selected chart** to load its data and
   options back into the pane, change anything, and **Update chart** to
   replace it in place. Charts remember their configuration in the file. The
   banner appears for every chart except the one already open in the pane —
   including its copy on another slide after **Duplicate Slide**.

   If you resized the chart on the slide by hand, an edit that leaves every bar
   and label where it is (a new title, new colours, a restyled outline) keeps
   your size. An edit that has to move or resize a bar or label redraws the chart
   at the size the pane's settings give it, so the hand resize is lost — the
   alternative was a chart with some bars stretched and the edited ones not,
   reading the wrong value. The same redraw happens if you have
   changed the stacking order of shapes inside the chart's group (Bring to
   Front on a single bar or label).

   Very occasionally PowerPoint draws the chart and then refuses to save its
   settings back onto it. The chart is on the slide and looks right; it just
   cannot be re-opened from the pane. **The pane now says so** instead of
   reporting Done — insert it again if you need to keep editing it, or undo
   (Ctrl+Z) to get back the version that was editable.
6. **Auto-update chart** pushes every pane edit to the slide live (debounced).
   Writes to the slide take turns: a push, an **Update chart** click, **Same
   scale** and **Explode** never run on the same chart at once — the later one
   waits for the earlier to finish. A write always goes to the chart it started
   from, even if you open another chart in the pane while it runs.
7. Ctrl+Z / Ctrl+Y undo and redo pane edits.

## Chart types

| Type | Use for | Notes |
|---|---|---|
| Stacked | composition over categories | one unified engine — a simple column chart is a stacked chart with one series |
| Clustered | side-by-side comparison | bar styles: plain bars, lollipop (stem + dot), dot plot, or a two-series dumbbell range |
| 100% | share of total | `100%=` row lets columns fall short of 100% |
| Waterfall | bridges (EBITDA, headcount) | values are deltas; type `e` in a cell for a computed total; multiple series stack |
| Mekko | two-dimensional composition | width ∝ column total; `X extent` row for explicit widths |
| Line | trends | date categories ("2025-01", "Jan 2025", Excel's "Jan-25") space proportionally to time; forecast periods draw dashed with hollow markers; `decorations.slope: true` turns two-period data into a slope chart (end rails, "Name value" labels both ends, no axis) |
| Area | stacked trends | |
| Butterfly | two-sided comparison (pyramids) | first two series, back to back, one scale |
| Scatter | correlation | `X`/`Y` rows; `Group` colors points; `X line`/`Y line` partitions; `Trend` fits an OLS line labelled with R² and p, or a **polynomial** curve (`scatter.trendDegree` 2–4, quadratic/cubic/quartic); quadrants shade a 2×2 matrix |
| Bubble | scatter + size | add a `Size` row; a size legend with reference circles appears automatically |
| Gantt | project plans | `Start`/`End`/`Milestone` rows; see [Gantt](#gantt) |
| Combo | columns + line | mark a series with type "line"; secondary axis optional; columns can be stacked, clustered, or 100% |
| Pie / Doughnut | simple shares | slices from 12 o'clock, clockwise; explode slices to highlight; `pie.breakout: [indices]` collapses categories into an "Other" slice detailed in a bar beside the pie (bar-of-pie) |
| Boxplot | distributions | `Min`/`Q1`/`Median`/`Q3`/`Max` rows, or raw samples → automatic Tukey whiskers and outliers; suffix rows with `\| group` for side-by-side grouped boxes |
| Radar | multi-dimension profiles | categories = spokes, series = translucent polygons; keep ≤ 3 series |
| Heatmap | value matrices | series = rows, categories = columns, one global color scale (diverging when data spans zero); marginal row/column totals optional |
| Tile map | values by geography | categories = region codes (US postal / ISO-2 / world regions); layout auto-detected |
| Cascade | stage-by-stage breakdown of a total | categories `"Stage \| Drop label \| Group"`; each bar is a subset of the previous, remainders hang as labeled boxes |
| Funnel | pipeline / conversion stages | centered bands, width ∝ value, conversion % between stages; order values ascending for a pyramid |
| Waffle | one dominant share ("68% of…") | 10×10 unit grid, 1 cell = 1%, filled from the bottom-left; a single category ≤ 100 reads as a literal % with a big-number legend; `100%=` overrides the denominator |
| Treemap | part-to-whole with many items | squarified rectangles, area ∝ value; categories named `"Group \| Item"` nest into two levels (group headers + child tiles) |
| Sunburst | a hierarchy as nested rings | `"Group \| Item"` categories → groups on the inner ring, items on the outer ring |
| Violin | full distribution per category | raw samples per column form a mirrored kernel-density shape with a median tick (fills degrade to outline in the live add-in) |
| Candlestick | financial OHLC / price action | datasheet rows `Open`/`High`/`Low`/`Close`; green when the period rose, red when it fell |

Nine kinds rotate into **horizontal** bars/rows with the "Horizontal (bar)"
toggle: the column family (stacked/clustered/100%), combo, waterfall, Mekko,
boxplot, and line/area (the profile chart, under Options and decorations). The
axes, the labels, the totals, the value bands and the Error/Target marks all
turn with them.

The analytical overlays — CAGR and difference arrows, value lines, callouts,
the grand total and the IBCS variance tier — are drawn only on an upright
chart: they are built around a vertical value axis. On the other sixteen kinds
the toggle does nothing at all, and leaves the chart exactly as it was.

## The datasheet

- **Row 1** = category headers; **column A** = series names; the corner cell
  is fixed.
- **Navigation**: Enter/↓, ↑, and ←/→ at the text edges move between cells,
  Excel-style. **+ Row / + Column** insert at the cursor; **− Row / −
  Column** delete at the cursor; **⇄ Transpose** swaps rows and columns.
- **Formulas**: cells starting with `=` evaluate A1-style references,
  `+ - * /`, and `SUM`/`AVG`/`MIN`/`MAX` over ranges (`=SUM(B2:E2)`). A
  reference reads a cell exactly as the chart does, so a percent-formatted
  cell (`35%`) counts as 35 and a date cell counts as its day — `=B3-B2`
  over two dates is a duration in days. A range may run past the end of the
  sheet (`=SUM(B2:B999)` totals the whole column), and **⇄ Transpose rewrites
  every reference** so formula cells keep the numbers they had.
- **Blank rows split stacks**: in a stacked chart, series separated by a
  blank row form side-by-side sub-stacks (clustered-stacked).
- **Dates**: ISO (`2026-01-15`), `15.01.2026`, and `Jan 2026` parse as
  calendar dates for Gantt timelines. A **slash** date that could be read either
  way round — `03/01/2026` is 3 January in Europe and 1 March in the US — is
  refused rather than guessed at, and the cell reads as empty. Write it as
  `2026-01-03` or `03.01.2026` to say which you meant. Slash dates with only one
  possible reading (`01/15/2026`, `2026/01/15`) are fine.

### Special rows (magic series names)

| Row name | Effect |
|---|---|
| `100%=` | per-category denominators for 100% charts |
| `X extent` | Mekko-with-units column widths |
| `X`, `Y`, `Size`, `Group` | scatter/bubble coordinates, bubble area, point groups |
| `X line`, `Y line` | scatter partition lines |
| `Trend` | OLS trend line with R² and p-value |
| `Error`, `Error+`, `Error-` | error bars at the column total / line point (symmetric or asymmetric) |
| `Target` | bullet tick across each column — or across each bar with **Horizontal** on, which is how a bullet graph is usually drawn; on waterfalls also a hatched gap-to-target segment |
| `Band low`, `Band high` | line charts: shaded confidence/uncertainty ribbon |
| `Min`, `Q1`, `Median`, `Q3`, `Max`, `Mean` | boxplot five-number summary (+ mean marker) |
| `Outlier <n>` | extra boxplot outlier dots |
| `Start`, `End`, `Milestone`, `After`, `Today`, `Holiday`, `Bracket <label>` | Gantt — see below |
| `% Complete`, `Baseline start`, `Baseline end` | Gantt progress fill and plan-vs-actual ghost bars |
| `Column <label>` | Gantt: a numeric table column beside the task labels (cost, FTE, days) |

### Gantt

`Start`/`End` values (numbers or dates) become bars; `Milestone` diamonds;
`After` (1-based row number) draws dependency arrows; `Today` a dashed red
line; `Holiday` shades dates; `Bracket <label>` spans an annotation.
Category cells support `Activity | Owner | Remark` columns, a leading `>`
for indenting, and section-header rows (no Start/End). Calendar timelines
pick weeks (with weekend shading), months, or quarters automatically.

**Waterfall "of which" columns** (`waterfall.detailGroups: [{ of, indices }]`):
break one step of a bridge into its parts, in line with the rest. The detail
columns decompose that step's delta starting from its own base, and sit OFF the
chain — they carry no running total, so every later total and the final one are
unaffected no matter what the parts add up to — while the dashed connector
steps over the whole group to the next real column. Name the parent (`of`); the
engine never guesses it. If the parts don't sum to the parent, you get exactly
the bars you typed — nothing is reconciled behind your back.

**Combo marker series** (`type: "marker"` on a combo series): draws the points
without a connecting line — a benchmark, target or consensus that belongs to
each category rather than to a trend. Keep it on the columns' own scale: on a
`secondaryAxis`, or under `combo.lineAxes: "independent"`, the marks would be
measured against a different axis than the columns they benchmark, which is
the one thing the mark is for. `pareto` keeps only the bars and its cumulative
line, so a marker series does not survive it.

**Overlap relief** (`scatter.spread: "x" | "y"` on scatter/bubble): when markers
cover each other, SSF Charts can nudge them along ONE axis you name — the other
stays exact. The shift is hard-capped (`scatter.spreadLimit`, in that axis's
own units; 2% of the range by default) and the cap is printed in the footnote,
so the chart discloses its own approximation. It is ignored under `quadrants`:
a nudge must never move a point across a quadrant line. Both axes carry data,
so SSF Charts will not move a marker freely in 2D — that would corrupt two
readings at once.

**Marker symbols** (`scatter.markers` on scatter/bubble): give each `Group` a
point shape — `["circle", "diamond", "triangle"]` makes group 1 circles, group
2 diamonds, group 3 triangles, cycling if there are more groups than shapes.
Shape carries what color cannot: a printed deck is often greyscale, and ~8% of
men cannot separate red from green, either of which flattens a palette back to
one group. The legend draws the shapes, so the channel is always explained —
including alongside a `Color` row, where color means something else. Shapes are
area-matched to the circle they replace, so a bubble's "area tracks Size"
survives: without that a group drawn as stars would read as a quarter the
magnitude of the same value drawn as squares. Available: `circle`, `square`,
`diamond`, `triangle`, `plus` — all PowerPoint preset geometry, so they stay
filled and editable in the deck on any 1.4+ host, and each reproduces its
preset exactly, so the preview and the deck agree on the ink. (A five-point
star and an X are deliberately absent: their presets reshape themselves in
ways the preview cannot reproduce, which would make the preview disagree with
the deck and quietly break the area match.)

**Heatmap sign marks** (`heatmap.symbols: "sign"`): draw a `+` or `−` in each
cell. A diverging heatmap states its direction in hue alone, and hue is what a
greyscale printer takes away: on the default palette the strongest positive
and the strongest negative land about 1.1:1 apart in grey, which is to say
the same tone. A palette chosen for contrasting lightness fares better, but
the two ends of a diverging ramp are the same distance from white by
construction, so the collapse is the rule rather than the exception. Marks are drawn only where the cell has no value label,
since the label already prints the minus sign; that makes them matter most
under `sizeEncode`, which suppresses labels entirely and leaves colour as the
only sign carrier — and on calendar heatmaps, which never draw labels at all.
A cell of exactly zero gets no mark: the scale already paints it the neutral
midpoint, and zero has no sign to state. The option is inert on one-signed
data, where every mark would say the same thing. Not up/down arrows: an arrow on a matrix of KPIs
reads as movement against last period, a claim SSF Charts does not check.

**Marginal histograms** (`decorations.marginals: "x" | "y" | "both"` on
scatter/bubble): distribution bars in a gutter along the top and/or right
margin — where the points bunch up, beside the chart that shows how they
relate. The bins subdivide the axis's own ticks, so every tick is a bin edge
and a bar can be read straight against the scale next to it. The gutter takes
real space, so the plot shrinks; on a chart too small to give any away, the
marginals are dropped instead.

Set `gantt.lanes: "owner"` to regroup the tasks under one header per owner
(the middle part of an `Activity | Owner | Remark` category). It is opt-in
because a plan's row order IS the plan — nothing regroups behind your back —
and it groups rather than sorts: inside a lane your rows keep the order you
wrote them, unowned tasks stay together at the end, and `After` dependencies
are renumbered so the arrows still join the same tasks. To control the
grouping yourself, type a section-header row and indent its tasks with `>` —
that has always worked and this changes nothing about it.

**Resource capacity vs load** is a different chart, not a Gantt option: make a
stacked column chart with the weeks as categories and one series per task, and
add a `Target` row for the capacity line. A load histogram needs a value axis,
which a timeline does not have.

Set `gantt.workdays` to make a bar's LENGTH mean working days rather than
elapsed days: a Mon→Mon task reads as 5 units, not 7, and weekends stop
inflating every bar. `true` is Mon–Fri; pass ISO weekday numbers for a custom
week (`[7,1,2,3,4]` = Sun–Thu). `Holiday` rows come out of the scale as well,
and weekend/holiday shading switches off — those days have no width left to
shade. Calendar timelines only.

`Column <label>` adds a numeric table column between the task labels and the
plan — the MS-Project look — right-aligned under `<label>` as its heading. Use
one row per column (`Column Cost €k`, `Column FTE`): each resolves its own
precision from its own values, so money and headcount don't share decimals. Put
the unit in the label. Section-header rows show whatever value you type in
them; nothing is summed for you. Unlike `Remark`, these align and format.

## Options and decorations

The *Format* section organises these into five collapsible groups — **Labels**,
**Axes & scale**, **Analysis**, **Layout**, and **Colours & style** — each with a
count of how many of its options are currently on.

- **Toggles**: Segment labels (auto-hidden when they don't fit),
  Series labels, Column totals, **Grand total** (one label at the top-right of
  a stacked/clustered column chart summing every category total — think-cell 14;
  independent of the per-column totals), Category labels, Value axis, Gridlines,
  **Connector lines** (joins stacked-segment boundaries between columns),
  **100% = note**, **Datamark axis (ticks only)** — Tufte-style tick dashes
  with no axis line.
- **CAGR arrow** and **Difference arrow** between two categories (the
  difference arrow can compare totals, one series' level, or start from a
  value line).
- **Value line**: at the mean (drawn as "Ø") and/or fixed values. The mean is
  of the categories that have data: an empty cell is left out of it, not averaged
  in as a zero. The arrows above follow the same rule — one is not drawn to or
  from a category with no data, because an empty cell is not a zero to compare
  against.
- **Segment order**: sheet / reversed / ascending / descending.
- **Axis scale** min/max pinning and **Axis break** (compresses an
  out-of-scale range, clearly marked).
- **Unit label** (`valueAxisTitle`) — a *short* unit at the top of the value
  axis: "€m", "%", "$m (log)". It shares the narrow band above the plot with
  the chart title and the topmost tick, so it is drawn at no more than 40% of
  the chart's width and anything longer is clipped. Put the explanation in the
  title or the footnote; this is a unit, not a subtitle. On a chart too short to
  hold both it yields to the title — see *Chrome yields to the title* below.
- **Labels**: decimals / suffix ("€m") / locale (1.234,5 for de-DE), and
  **label content** — any combination of value, %, series, category.
- **Footnote / source** — every good chart cites its source; renders
  bottom-left.
- **Explode slices** (pie/doughnut).
- **Stepped line/area** (`stepped: "before"|"after"|"center"`): draws a
  staircase instead of sloped segments — good for rates that hold then jump.
- **Area with negatives**: area charts now stack positives above the zero
  line and negatives below it, so P&L-style series that go negative dip
  under the baseline instead of being clamped to zero.
- **Scatter trajectory** (`trajectory`): connects scatter/bubble points in
  row order with a direction trail (Gapminder-style path over time).
- **Continuous color** (`Color` row): colors scatter/bubble points on a
  sequential ramp by a numeric value, with a gradient legend — a third
  variable beyond X/Y (and Size).
- **Smoothed lines** (`smooth`): draws line charts as smooth curves instead
  of straight segments.
- **Waterfall spacers** (`spacerIndices`): blank grouping gaps that split a
  long bridge into sections; the running total and connector bridge across.
- **Gantt summary bars** (`summaryBars`): a capped summary bar on each
  section-header row spanning its child activities' min start → max end.
- **Notched boxplots** (`notch`): pinch each box at the median confidence
  interval — non-overlapping notches flag a significant median difference.
- **Radar band** (`radarBand`): shade the peer min–max envelope and draw the
  last series on top — a "peer range + us" competitive profile.
- **Other bucket** (`otherBucket`): collapse a long tail of stack segments
  into one "Other" series, keeping the largest few.
- **Calendar heatmap** (`heatmap.calendar`): lay a daily date series out as a
  weekday × week grid with month labels (contributions view).
- **Butterfly stacked flanks** (`butterfly.split`): stack more than one
  series on each side of a butterfly chart.
- **Radar per-spoke scales** (`radar.perSpoke`): normalize each spoke to its
  own maximum so mixed-unit KPIs compare by shape.
- **Missing-data bridge** (`bridgeGaps`): connect a line straight across
  missing (null) points instead of breaking it.
- **Floating segments**: a stacked-column series with `color: "transparent"`
  reserves its space without drawing, floating the bars above it.
- **Combo base modes** (`combo.columns`): put the lines over a waterfall
  bridge or a Mekko as well as stacked / clustered / 100% columns.
- **Combo independent line axes** (`combo.lineAxes: "independent"`): give each
  line series its own scale so unlike-unit KPIs share one chart.
- **Combo label priority on a tight frame**
  (`decorations.tightLabelPriority: "columns" | "line"`): a short chart can be
  too small to carry both the column totals and the line's point labels — the
  two collide when a column total and a line value fall at the same height, and
  on a 60pt-tall frame neither has anywhere to move. `"columns"` (the default)
  keeps the totals; `"line"` keeps the point labels. Only labels still
  overlapping after the de-collision pass are dropped, so a chart with room
  keeps both and nothing changes at ordinary sizes.
- **Hex tile maps** (`tilemap.shape: "hex"`): draw the cartogram with
  hexagonal tiles instead of squares.
- **Tilemap mini-glyphs** (`tilemap.glyph: "bars"`): a mini bar chart per
  region from a multi-series datasheet.
- **100% charts with negatives**: `stacked100` now shows negative segments
  below the zero line instead of clamping them away.
- **Semi-circle gauge** (`doughnut` + `pie.semi`): a half-doughnut scorecard.
- **Pareto** (`pareto`): sorted bars + a cumulative-% line (80/20 view).
- **Bump chart** (`decorations.bump`): rank-over-time on an inverted axis.
- **Horizontal profile chart** (`horizontal` on line/area): rotates the chart
  so categories run down the left and values extend to the right. Every line
  decoration rotates with it — forecast styling, stepped, smooth, the
  plan-vs-actual ribbon, gap bridging and per-point highlights — and so do the
  `Band low`/`Band high` confidence ribbon and the date-proportional category
  spacing. (`logScale` is the one exception, as on horizontal columns.)
- **Critical path** (`decorations.criticalPath` on Gantt): red-outlines the
  longest chain of `After`-dependent tasks — the tasks that drive the finish
  date — and draws its dependency arrows thicker in red.
- **Mean±SD box** (`boxplot.meanSd`): draws each box as mean ± 1 SD with the
  centre line at the mean and whiskers to mean ± 2 SD (raw-sample mode) — the
  scientific summary rather than the quartile one.
- **Sparklines** (`decorations.sparkline` on line/area): compact, axis-less
  word-sized trend lines with min/max/last dots; pair with small multiples for
  a KPI dashboard table, one row per series.
- **Radial bar chart** (`radar.bars`): a circular bar chart / Nightingale rose —
  each category is an equal-angle wedge whose radius encodes its value; multiple
  series stack outward.
- **Stacked radar** (`radar.stacked`): series stack cumulatively along each
  spoke as nested filled bands (part-to-whole across dimensions).
- **Variable-radius pie** (`pie.variableRadius`, or a `Radius` datasheet row):
  slice angle encodes the first series while radius encodes a second metric.
- **Cell-size heatmap** (`heatmap.sizeEncode`): each cell's area encodes the
  value's magnitude (colour still encodes the signed value) — the corrplot view
  for correlation / signed matrices.
- **Heatmap clustering** (`heatmap.cluster`): reorder rows by average-linkage
  similarity so similar rows sit together, with a dendrogram in a left gutter.
- **Combo area base** (`combo.columns: "area"`): put the line series over a
  stacked-area base instead of columns — trend-of-mix plus a KPI line.
- **Boxplot jitter** (`jitter`): overlays the raw observations as jittered
  dots over each box (raw-sample mode).
- **Gap width & overlap** (`gapWidth`, `overlap`): Excel's two column-spacing
  controls. Gap width (0–500%) sets how wide the columns are — 0 makes them
  touch for a histogram look; overlap (−100…100) sets how much clustered bars
  within a category overlap or gap.
- **Same scale** (deck-wide or selection) re-renders charts on a common axis.
  It will refuse rather than half-do the job: if PowerPoint would not describe
  every slide, or would not answer for some of the charts it was asked to
  redraw, the pane says which and changes nothing rather than reporting a
  shared scale the deck does not have. If the chart open in the pane is one of
  them, the pane follows it to its redrawn shape and takes over the shared
  axis scale, so the next **Update chart** does not undo it.
- **Download SVG / Download PNG** (overflow menu) save the preview as a vector
  or a 2× raster image — handy for email or chat, where the native-shape output
  can't go. The real deliverable is still **Insert into slide** (editable shapes).
- **Copy chart link** (overflow menu) copies a URL that reopens the exact chart
  on the hosted gallery — the config rides in the link's `#c=` hash (never a
  server log). Good for support or sharing a config without a file.
- **Small multiples** (`multiples: {columns?}`, JSON): splits a multi-series
  chart into a grid of single-series panels titled by series name, on one
  shared value scale (column family, line/area, waterfall, radar).
- **IBCS scenario notation** (`series.scenario`, JSON, column family): encodes
  the data's nature in the fill so actual vs plan vs forecast reads without a
  legend — **AC** solid, **PY** lighter, **PL** / **BU** outlined/hollow, **FC**
  hatched (solid in the live add-in). The two-letter code is appended to the
  legend/label; give each scenario series the same `color` for the single-hue
  IBCS look.
- **IBCS variance tier** (`decorations.variance`, JSON, column family): a strip
  below the columns showing an `actual` series' deviation from a `reference`
  (plan / previous year) per category, as signed bars from a zero line — absolute
  Δ (default) or relative `mode:"percent"`. `goodIsUp` (default true) colours the
  sign green (favourable) / red (not); set it false for cost-like metrics.
- Advanced (JSON-only, survive re-editing): callouts (speech bubbles),
  background bands, per-cell highlight colors, pattern fills, log scale
  presets — see `skill/reference.md` for the full schema.

## Colors and style

- **Palette** presets, per-series color pickers, and **Use deck theme** —
  reads the presentation's Accent 1–6 theme colors (PowerPointApi 1.10).
- **Export style / Import style**: a corporate style file (JSON: palette,
  font, negative/total colors) applied to every chart you make. Export gives you
  the style the pane is currently drawing with — which is not always your own
  (see below).
- **Save style to deck / Use deck style**: the same style file, stored **in the
  presentation** instead of in your browser. Your imported style follows the
  browser, so a branded deck sent to a colleague used to arrive with none of it
  and every chart they added drifted further; a style saved into the deck
  travels with the file.
  - The **deck's style wins** when it has one, so a colleague opening your deck
    draws your brand without doing anything. The pane says which style is in
    force whenever it changes.
  - **Importing** a style in the pane overrides the deck for the rest of the
    session — it is an explicit act, and the deck should not silently undo it.
    **Use deck style** switches back.
  - Needs PowerPointApi 1.7. On an older host the button says the deck could not
    be updated rather than implying it was; your own style still applies.
  - Chart *data* is unaffected: each chart's config stays in its own shape tag,
    which is what lets a chart survive copy/paste into another deck.
- **Templates** — the dropdown groups built-in **Starters** (polished, ready-to-
  edit presets: revenue bridge, growth columns + CAGR, channel mix, market
  share, KPI trend, bullet vs target, programme Gantt) with **My templates**.
  **Save as template** stores the whole current chart (data + options) locally
  under *My templates*; **Delete** removes a saved one (starters stay).

## Small charts and big fonts

A chart's size and its font size are both numbers you type, and the two together
decide how much room the chart has for anything that is not a mark. SSF Charts
adapts rather than spilling, and it does so in a fixed order — so it is worth
knowing what you will see when you shrink a chart onto a busy slide or raise the
font for a poster.

**Nothing a chart draws leaves its own box.** Neither PowerPoint renderer wraps
or clips a text box, so a label wider than its frame would be drawn across
whatever sits next to the chart on the slide. That cannot happen: every kind is
swept against **8 frame sizes × 7 font sizes** (60x300 through 960x540, 6pt
through 32pt) in CI, and nothing may put ink outside its own frame.

That sweep covers the decorations too — every annotation on the Decorations tab,
switched on one at a time, across every kind, frame and orientation. A callout
whose text is wider than the chart, a band label on a chart with no room below
the plot, a quadrant label in a narrow zone: each is fitted or dropped by the
same rules as any other label, rather than drawn over the slide.

The options that change a chart's shape are swept the same way — a broken axis, a
semi-circle gauge, an exploded slice, a Pareto's second axis, and small multiples,
which divide the frame into one chart per series and can leave each of them very
little. A panel too short to carry its own title or category names draws neither
and keeps its marks.

The data is swept too, because a chart's own numbers change what it has to draw:
long category names, twenty-four of them, ten series, values in the billions, and
values that are all negative. A funnel with more stages than the frame has points
draws every stage thinner rather than dropping the last ones off the chart.

**Labels shrink to the room they have, then are dropped.** A name is fitted to
the mark it sits on — the column's slot, the row's pitch, the wedge's chord, the
gap between two rings — and below about 5pt it is not drawn at all. A label too
small to read is not a small label, it is ink; the marks it named are still
there, and the chart still says what it says.

**Text is never drawn on top of other text.** The same sweep runs a second way,
comparing every label's ink against every other label's, across **7 frames × 7
fonts in both orientations**, and it is zero. The one declared exception is a
scatter or bubble point label touching an axis tick: a point's label is data and
a tick label is chrome, and the alternatives were measured — both of them drop
point labels on charts as roomy as 480x300, which is the worse trade.

**Chrome yields to the title.** On a chart too small to carry a title, an axis
and a legend at the font you have chosen, the axis strip, the legend or the
colour key is dropped and the title stays: it is the one label that says what
the reader is looking at. A 300x60 banner at 24pt therefore shows its title and
its bars, and no category names — give it more height, or a smaller font, and
they come back.

The **unit label** now follows the same rule, and until 2026-08-29 it was the one
thing in that band that did not: on a chart with no room for both it was drawn
straight through the title, so the reader lost the title as well as the unit.
It is dropped only where its text would land in the title's — measured, that is
80x60 and 300x60 at 18pt and nothing else, with every chart at 480x300 and above
keeping its unit at every font size. As with the category names: more height, or
a smaller font, and it comes back.

**Marks stay inside the chart too.** A marker is drawn *around* its position, so
a bubble at the very edge of the plot would put half its circle outside the
chart. Scatter and bubble hold a mark far enough in for its whole glyph to fit,
which moves it by a point or two only when it would otherwise overflow.

None of this is triggered by a setting, and none of it changes an ordinary
chart: at the default font on a slide-sized frame every fit is the identity.

## Elements

- **Harvey ball** — fraction-filled circle, size slider.
- **Checkbox** — ✓ / ✗ / − status marks.
- **Process flow** — chevron step bar with a highlighted step.
- **KPI tile** — big number + delta arrow + caption; the arrow takes
  semantic colors from the delta's sign, and **↓ good** flips them for
  metrics where falling is better (costs, churn).
- **Table from datasheet** — rule-based table (lines top/header/bottom only,
  never side borders, a gap every 5 rows, optional bold **Total row**).
  Cells accept effect tokens: `[hb:0.75]` mini harvey ball, `[up]` `[down]`
  `[flat]` trend arrows, `[good]` `[bad]` semantic colors. A formula
  (`=SUM(B2:B3)`) and a waterfall's computed total (`e`) print as the number the
  chart shows, not as typed; every other cell prints exactly as typed.
- **Agenda** — inserts one chapter slide per line, each highlighting its own
  chapter.

## Automation (JSON)

**Export current** copies the chart's `ChartConfig` JSON; **Import** loads
one; **Insert batch** inserts an array of configs at once, each on a new slide
of its own and fitted to it (a chart wider than the slide is scaled down, as a
single Insert would). A config that cannot be drawn is skipped, and the note
names which one and why while the rest still go in. The same JSON
drives the CLI (`npm run render`) and the **Claude Agent Skill**
(`skill-dist/ssf-charts.zip` — upload at claude.ai → Settings →
Capabilities → Skills, then ask Claude for "a waterfall of …" on any
surface, including Claude for PowerPoint).

**`render` (per config, optional):** `"shapes"` (default) inserts native,
editable PowerPoint shapes — the think-cell pattern this add-in is built for.
`"image"` rasterises the whole chart into one picture object. In the pane it is
the **Insert as picture** checkbox beside Auto-update; in JSON it is the
`render` key, honoured by Insert, Update chart, Insert batch (per config),
Same scale and the skill CLI alike.

Reach for it only on the densest kinds (area, violin, sunburst, tile-map,
waffle) when a shapes-mode render would push the deck past the PowerPoint-web
dense-shape wall — one picture always lands, where 250 shapes may not. The
costs are real and worth reading before you tick it:

- The chart is **not editable** in PowerPoint. **Explode to native shapes**
  (⋯ menu) turns it back — it reads the config off the picture and re-draws the
  real shapes in place, so nothing is lost, but it is a manual step.
- It ignores the deck theme and blurs when rescaled (it is rasterised at 2×).
- Below PowerPointApi 1.10 it is **invisible to screen readers** — no alt text
  and no text shapes either. This is the one axis where picture mode is strictly
  worse than shapes.

If the host can't insert pictures (needs PowerPointApi 1.8) or the browser
can't rasterise, the chart still lands as native shapes and the pane says which
happened rather than failing.

### Testing

**Insert demo deck** appends a fresh slide for every chart kind plus a set of
feature and element highlights — a fast way to eyeball how everything renders
as native shapes in your PowerPoint. Each chart stays editable (click it and
**Edit selected chart**); delete the slides when you're done. The deck opens
with a title slide stamping the build and host, and closes with a results slide
summarising what rendered, what the host skipped or failed, and how long the run
took — so an exported PDF is a self-contained regression record.

**Path** chooses how the deck is drawn.

- **One file insert** (the default) builds the whole deck as a .pptx in the
  pane and hands it to PowerPoint in a single call, instead of drawing it shape
  by shape through hundreds of round trips. Charts arrive already grouped and
  already tagged, so they are editable exactly as before. The file is built at
  *your* deck's slide size — a generated deck that declared a different size
  was one PowerPoint rescaled on insert, which moved every chart on every
  slide, and 4:3 decks got that every time.
- **Shape by shape** draws it through Office.js, one chart per slide.
- **Both, one after the other** takes each path in turn into the same deck. The
  two fail in completely different ways, and comparing them used to mean two
  separate runs with a deploy in between; this is one session, one host, one
  downloaded log holding both runs. Each carries its own identity token, so
  nothing confuses one run's slides for the other's.

Whichever you pick, the add-in falls back to the shape-by-shape path on its own
if the host cannot take a generated deck, or if the attempt landed nothing.
Once any slide has landed it will not fall back — that would insert the deck
twice. ("Both" drawing the deck a second time is the one case where that is the
intent.)

Every run ends by reading the deck back and repairing it. This matters on
PowerPoint for the web, where a `context.sync()` can commit minutes after the
add-in has given up waiting for it — so the chart is on the slide, but the
"NOT COMPLETE" banner stamped while it still looked empty says otherwise. The
closing pass compares each slide against the item it was drawn for, deletes
duplicate slides, clears banners that contradict the shapes underneath them,
and restores the config tag on charts that lost it — re-grouping the ones that
landed as loose shapes, and re-tagging the ones that are already a single
object — so every chart is editable again. A chart the host would not let it
measure is reported rather than repaired: acting on a read that came back
empty once rewrote fourteen charts whose config was already correct. The run's
summary is written from what the deck actually holds, not from what the add-in
believed while the host was still catching up.

It only ever touches slides from the run that just finished. Every slide a run
adds carries a token identifying that run, so inserting the demo deck a second
time cannot make the second run's slides look like duplicates of the first's —
and a slide you added yourself, or copied from a demo slide with PowerPoint's
Duplicate Slide, is reported and left alone.

*Which selection call wedges the host* used to be one of them and is now part of
an ordinary run. It asks the host, one call at a time, which selection call it
stops answering, and it needs to be the first thing in the run that selects a
*shape* — so it sits immediately before *edit the chart the user selected*, the
only other place the add-in does that. If it finds a wedge, that scenario says
so and skips rather than spending another wait to learn less. It reports *ok*
whatever it finds, because it is a measurement rather than a check.

**Clean up the last round** deletes the slides the last round added, and only
those. It works from the list of slide ids the round watched appear, never from
a rule about which slides look like a test — a rule could match a slide you made.
It stays greyed out until a round has produced such a list, and if the host
refuses a delete it says how many it actually removed rather than claiming a
clean sweep.

**Probe, then self-test** does all of it on one click — the host probe, then the
self-test, then a look at what actually landed — and saves the lot as one JSON
file. That file is the whole upload. It carries every shape on every slide (name,
id, position), which is what questions like *did the old chart's shapes get left
behind* are really about, and it carries the host's own picture of the slides the
round added, so there is nothing to screenshot. If the host will not describe its
deck, the round still writes the file: you lose the pictures and keep the
verdicts.

**A round leaves its slides in the deck, and most of them look empty.** That is
deliberate — the scenarios' working slides are the evidence, and the deck is
meant to be something you can open and scroll. But there are a lot of them and
they are mostly blank: one real round added 43 slides of which 36 read back with
nothing on them. The round now says so when it finishes, and **Clean up the last
round** removes exactly the slides it added, by id, so nothing of yours is
touched. Nothing is lost by clearing them once the file is saved — the file
carries the shape list and the pictures.

If you want to look at them first, do that before pressing it: the button works
from the id list the round recorded, and reloading the pane drops that list.

**Picture every slide** is **on by default**, and it is what makes the pictures
worth having. Without it the first twelve added slides are photographed and the
rest are recorded as never asked about — and that limit decides what a round can
settle, because a slide that reads back with no shapes is only *probably* empty:
this host will report a shape collection short without saying so, and the
picture is what turns a readback into a fact. Untick it for a struggling host or
a smaller file; the pictures are the heaviest thing the add-in does and they run
at the end of a round.

**Verbose trace** records every step the add-in takes — each drawing batch, each
slide the host loses, each call it stops waiting for, each repair action, and
every pane action including ordinary inserts and updates. It is on by default
while the add-in is being validated against real hosts, and it rides along
inside the run log rather than in a separate file. Nothing is written while it
is off.

**Scenario** picks what **Run host self-test** actually runs. Left on *All* it
runs the whole battery; pick one and it runs that scenario plus the two inserts
every scenario needs to have probe charts to work on. A full round takes several
minutes and leaves slides in the deck, so when you are chasing one failure, pick
it — seconds instead of minutes, and far less to clean up afterwards.

Two entries in that menu are **only** reachable that way.

*Edit the chart YOU click* waits for you. Pick it and the pane asks you to click
an SSF chart, counting down for 30 seconds; click one and it does exactly what
pressing *Edit it* does — reads the chart back from your selection, edits it,
and checks the result is still editable — then reports what happened. It is left
out of a full run because it blocks on a person, and if nobody clicks it reports
*skipped* rather than failing. It exists because asking Office.js to select a
chart and clicking one yourself are not the same thing on PowerPoint on the web,
and only the second one is what you actually do.

*What makes a long run slow down* is a measurement, not a check. Long runs on
PowerPoint on the web get slower and eventually stop answering, and three things
could be causing it: the request context accumulating, the deck growing, or the
browser tab simply having been open a while. An ordinary run varies all three at
once, so no amount of reading a run log can separate them. This draws the same
shapes onto two slides — one slide inside a single long-lived request context,
the other with a fresh context per batch — and reports both timing curves. If
only the long-context curve climbs, the context is the culprit and shorter
contexts would fix it; if both climb together, the host itself is slowing and
they would not. It takes a minute or so and leaves two slides of small
rectangles behind. It reports *ok* only when neither curve climbs.

**Run host self-test** runs the twelve things the demo deck does not cover:
inserting on top of an earlier run, two slides claiming one slot, editing a
chart on the slide you are looking at, adding charts to a slide that already
has content, a deck-wide rescale, turning a degraded picture back into native
shapes, which selection call the host stops answering, whether a shape you had
selected survives an insert, editing the chart you have *selected*,
stopping a run part-way, and whether a rasterise poisons the draw that follows
it. They run in that order on
purpose: the longest-standing checks first, the newest last, so that if a run
does not survive to the end you still have the verdicts for everything with a
track record. Each reports what was
actually observed; one that your host cannot run is reported as skipped rather
than failed. It leaves its slides in the deck so you can look at them — delete
them when you are done, exactly like the demo deck.

One of the twelve — **whether a chart is actually visible** — was out of the
full run for a while, and the reason is worth knowing because it shapes what the
battery does. It killed the browser tab in five consecutive rounds on PowerPoint
on the web without once returning a verdict, and because the report is written
when the battery finishes, a scenario that ends the run takes every other
verdict's *report* with it even when it runs last.

Running it on its own is what identified the cause. It used to borrow a
brand-new slide and immediately ask PowerPoint to render it as an image, and on
the web that call kills the tab. It compares before and after on a slide the run
already added now, so it never makes that call — and it passes. Rendering a
slide is fine; rendering one created moments earlier is not.

The same turned out to be true of **adding** a slide: a second `slides.add()`
four tenths of a second after the first killed the tab twice in the degradation
experiment. Nothing in the battery takes a scratch slide any more.

Three of those are new, and are the ones worth understanding. **Editing the
chart you selected** goes through the same read the pane uses when you click a
chart and press *Edit it* — a different code path from every other scenario,
and the one an actual user travels on. On PowerPoint **on the web** it reports
*skipped*, and that is the correct answer rather than a gap: asking the host to
select a shape from code makes it stop answering selection questions
altogether, so the scenario waits ten seconds, names the limitation and moves
on. It is a bug in the host and it does not touch you — the pane only ever
reads the selection *you* made with a click, which works normally. **Whether a chart is visible** — the picked-only
one — asks the
host to render a slide before and after drawing, and compares: every other
check in the add-in counts shapes and reads tags, all of which pass happily for
a chart drawn in white, at zero size, or off the edge of the slide. It borrows
a slide to do this and takes it away again. On PowerPoint **on the web** it has
never got as far as an answer: the host has refused to rasterise a freshly-added
slide in three different ways across three rounds — a `GeneralException`, a
call that returned nothing, and once by never answering at all — and each round
ended with the tab gone. It gives up after twenty seconds rather than the
ninety a deck read gets, since a rasterise that is going to answer answers in
about a second, and the long wait once cost a whole round while the tab died
waiting. Note this is the only check of its kind that needs a host at all:
every chart shipped here is rasterised in a real browser on every CI run, and a
chart that draws but is not visible fails the build. **Stopping part-way**
confirms a stopped run adds nothing
and leaves nothing behind claiming to be a chart.

**Live steps** is the list at the top of the Testing section, above the
buttons. Every step of a run appears there as it happens, **newest at the top**
— which scenario started, which one finished and how, which slides were read
back, and the phase of anything that went wrong.

Both of those orderings are deliberate, and both are the opposite of how a log
usually reads. When PowerPoint dies you get whatever pixels were on screen: no
scrolling, no clicking, often a dialog over half the pane. A list that grows
downwards puts the last thing that happened at the bottom of a small scrolled
box, which is exactly where it cannot be relied on to be visible; growing
upwards puts it in a fixed place, one line under the header. And a box sitting
below nine controls and a paragraph is a box you have to scroll to before you
can photograph it — which is why it now comes first, and the buttons that start
a run come after it. The list is also scrolled into view once when a run
starts, so it cannot begin somewhere off-screen.

It is not a prettier version of the run log; it is the record that survives
things the log does not. The log can only be downloaded once a run **ends**, and
the runs worth explaining are the ones that do not end: a host that stops
answering, or a PowerPoint that puts up *"Sorry, we ran into a problem"* and
takes the pane's memory with it. Both have happened. What is already on screen
survives both. **Copy** puts the whole list on the clipboard, with a header line saying the
order so nobody reads the run backwards; failing that, screenshot it before you
reload. **Clear** empties it.

**Run host probe** asks this PowerPoint a fixed list of questions — does a
proxy survive a sync, does writing the same tag twice overwrite it, does a
deleted shape report itself gone — and saves what it answered. It works on
scratch slides it appends and deletes again, so it does not touch anything you
made.

It tries to leave the deck as it found it, and it now **says whether it
managed**. On PowerPoint on the web it sometimes cannot: a run on 6 August 2026
left 21 blank slides behind, an earlier one left 14, and neither answer sheet
mentioned it — the deletes were attempted, the failures were discarded. The
sheet carries a `scratch-slides-returned` row for exactly that, so a run that
littered says so instead of leaving you to find out by scrolling. Blank
leftovers are safe to delete by hand; nothing else in the deck depends on them.

It exists because of something worth being blunt about. Every automated test in
this project runs against a *fake* PowerPoint, and nobody has ever checked
whether that fake tells the truth. The faults it models are things a real host
taught us the hard way; everything else about it is an assumption. Where an
assumption is wrong, the tests resting on it are confidently wrong too, and a
green suite says nothing about that.

So: run it once on each host you care about, send the file, and
`npm run host-diff` lines it up against what the fake claims. Every
disagreement is either a fake that lies — in which case some tests are worth
less than they look — or something PowerPoint does that we did not know. Both
are worth the click.

**Download the crashed run** appears only when there is one to download. Every
step is written to your browser's storage as it happens, and the record stays on
offer until its file has actually been saved. Two runs qualify, and the banner
says which one you have:

- **A run that never reported finishing** — it crashed, or you closed the tab on
  it. This is the case a screenshot cannot cover: the pane is gone, you already
  reloaded, and you want the whole run rather than the twenty lines that
  happened to be visible.
- **A run that finished, but whose file was never saved.** Completing and
  handing you the file are different things, and the second one can fail on its
  own: a task pane is a nested frame, and a browser can refuse a download from
  one exactly as it already refuses the clipboard. A host that dies moments
  after the round ends takes an in-flight download with it too. This used to
  lose the run — the pane marked it finished, which retired the record, and then
  attempted the save.

The file is what to send; `npm run triage` reads it without needing the deck.

It keeps the last 2000 steps, so a very long run loses its beginning and never
its end, and it lives in browser storage — clearing site data clears it too.
Saving it, or pressing **Download run log**, dismisses the offer. If a round
says it saved a file and nothing appears in your downloads, press
**Download run log** — the run is still in the pane until you reload, and still
in storage after you do.

**Download run log** saves the last run as JSON: the run's identity token,
per-item timings, statuses, what the host did to each slide, and the repair
pass's verdicts, plus the trace — which opens with a tally of every step and
every failure reason, so a long run can be read without reading all of it.
Attach it to a bug report instead of retyping the summary line, and attach the
deck with it: the token joins the two, and `npm run triage` reads them together
to say where the run's account and the file disagree.

### Very dense charts on the web

A violin is 253 native shapes; an area chart 176, a tile map 122, a waffle 103.
PowerPoint on the web will not draw those reliably — it is the one host with no
resource limits of its own, so an add-in that asks too much takes the whole tab
down rather than being throttled. On the web only, and only past ~90 shapes,
SSF Charts inserts such a chart as a picture and says so. It still carries its
configuration, so **Edit selected chart** works as usual and **Explode to native
shapes** turns it back into shapes on a host that can take them. Ticking
**Insert as picture** yourself always wins; this never overrides your choice.

### When PowerPoint won't redraw a chart

Updating a chart in place is the heaviest thing the add-in does on the web: every
shape is replaced, on the one slide guaranteed to be on screen — and the live
canvas is where PowerPoint gives up first. If it stalls, the pane works down
three fallbacks on its own, least disruptive first:

1. **It looks away.** Another slide is selected so yours is off-screen, where the
   host accepts four times as many shapes per round trip, and your selection is
   put back afterwards. Nothing is lost. If the deck has no other slide to look
   at — a one-slide deck, the commonest case when you are building your first
   chart — a blank slide is added at the end for the duration and removed again
   when the redraw finishes. You may see it appear briefly. On the rare host
   that refuses to remove it, the pane says so; it is the last slide and safe to
   delete.
2. **It rebuilds the slide.** The chart is generated as a one-slide file and
   swapped in — no drawing at all. Only when the slide holds *nothing but the
   chart*, because the replacement is a new slide and does not carry the old
   one's speaker notes or transitions. The pane tells you when this happens;
   select the chart again to keep editing.
3. **It draws a picture.** Never stalls, and the chart stays re-editable from the
   pane (**Explode to native shapes** turns it back), but what is on the slide is
   a raster.

If all three fail you get PowerPoint's own message, not a substitute for it.

### Stopping a long operation

While the pane is working, a **Stop** button appears next to the progress bar.
Inserting a whole demo deck, or applying **Same scale** across many charts, can
run for minutes on the web; Stop is the way out.

It stops at the next safe point rather than instantly. PowerPoint cannot abort
a round trip already in flight, so the batch being drawn finishes first — the
button reads *Stopping…* until it does. What that safe point is depends on the
operation:

- **Inserting a deck** stops between slides. Every slide already finished is
  complete, grouped and tagged, and is kept.
- **Same scale** stops between charts. Charts already rescaled keep the new
  scale; the rest are untouched, so the deck is left on two scales — re-run it
  to finish the job. A chart a Stop catches mid-redraw is cleared exactly as a
  single-chart redraw is (below), and the pane says so.
- **Redrawing one chart** stops between batches. Because an update replaces
  every shape, a chart caught mid-redraw is left partly drawn — the pane clears
  what it had drawn so you are not left with debris under a half-chart, and
  **Ctrl+Z** restores the original. The pane says so — *Stopped — the chart was
  removed and not redrawn* — and goes back to **Insert into slide**, because
  there is no chart left on the slide for **Update chart** to write to; the data
  stays in the pane. A Stop that lands before the chart is touched leaves it
  exactly as it was, and the pane says *anything already drawn was kept*.

Stop never falls back to rebuilding the slide or drawing a picture. Those exist
to get a stalled chart drawn some other way, which is the opposite of what you
asked for.

## Excel companion

Sideload `manifest-excel.xml` in Excel, select a range, and **Generate**
turns it into SSF Charts JSON to paste into the pane's Automation box — the
practical substitute for live data links.

## Troubleshooting

- **Charts insert ungrouped** — host below PowerPointApi 1.8; everything
  still works, shapes just aren't grouped.
- **Pie slices missing** — host below 1.10 (no shape rotation).
- **"Theme unavailable"** — host below 1.10.
- **Chart not recognized for editing** — it must carry the SSF Charts tag;
  charts inserted before tagging existed can't be re-opened.
- **"SSF Charts has been updated"** — a new version was published while your
  pane was open, so the pane is holding a page whose files have been replaced.
  Close the pane and open it again. Your slides are untouched; nothing was lost,
  and the chart you were working on is still on the slide.
- **Other languages** — the pane is translation-ready and localizes itself from
  `Office.context.displayLanguage`, but **no dictionary ships today**: every
  string is English. Adding one is a single entry in `DICTS` in
  `src/taskpane/i18n.ts`. This line used to promise German, which the repo's
  own test asserts the pane does not have.
