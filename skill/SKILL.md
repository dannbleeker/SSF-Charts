---
name: ssf-charts
description: Create think-cell-style consulting charts as native, editable PowerPoint shapes — waterfalls (EBITDA bridges), Mekko/Marimekko, stacked/clustered columns, Gantt plans, combo, pie, scatter/bubble, butterfly — plus agenda slides, from a JSON chart config. Use when the user asks for professional/consulting/think-cell-style charts, a chart in PowerPoint (.pptx), a waterfall or bridge chart, a Mekko, a Gantt/project plan slide, or CAGR/difference-arrow annotations.
---

# SSF Charts: think-cell-style charts for PowerPoint

Turn a JSON **ChartConfig** into a `.pptx` slide made of **native, individually
editable PowerPoint shapes** (never pictures) — with think-cell's signature
decorations: CAGR arrows, difference arrows, value lines, automatic totals, and
collision-avoiding labels.

## Workflow

1. **Author a ChartConfig** (JSON) from the user's data. Full schema and all
   datasheet-row conventions: read `reference.md`. Minimal example:

```json
{
  "kind": "waterfall",
  "title": "EBITDA bridge FY24",
  "data": {
    "categories": ["FY23", "Volume", "Price", "Cost", "FX", "FY24"],
    "series": [{ "name": "Delta", "values": [86, 14, 9, -12, -4, 0] }]
  },
  "waterfall": { "totalIndices": [5] },
  "decorations": { "difference": { "from": 0, "to": 5 } }
}
```

2. **Render.** One-time setup: `npm install pptxgenjs jszip` (in this skill's folder).
   For `render: "image"` charts also `npm install @resvg/resvg-js`.

```bash
node scripts/render-pptx.mjs charts.json out.pptx   # native shapes, 1 chart/slide
                                                   # each chart is one group carrying its
                                                   # config — re-editable in the add-in
node scripts/render-svg.mjs charts.json out/        # quick SVG previews
```

`charts.json` may hold one config or an array (one slide each). A bad config in
an array is isolated: it becomes a visible error slide and the rest still render.

**`render` (per config, optional):** `"shapes"` (default) inserts native,
editable PowerPoint shapes — the think-cell pattern. `"image"` rasterises the
scene into one picture object; the chart stays visible on the slide but is no
longer editable in PowerPoint. Reach for `"image"` only on the densest kinds
(area, violin, sunburst, tile-map, waffle) when a shapes-mode render would
push the deck past the PowerPoint-web shape wall — a picture sidesteps the
`slides.add()` drops and the "we ran into a problem" crash. Everything else
should stay `"shapes"` so the recipient can edit the chart.

**Agenda slides** use the same JSON, with `kind: "agenda"`:

```json
{ "kind": "agenda", "chapters": ["Context", "Findings", "Recommendation"], "highlight": 1 }
```

`highlight` is the current chapter (0-based; omit or `-1` for an overview). For a
section-divider before each chapter, pass an array with a different `highlight`
per slide. An agenda has no `style` and is always drawn on a white slide.

3. **QA visually.** Render the SVGs and inspect them (or convert the pptx to
   images) before delivering: check label overlaps, totals, axis sanity. Fix
   the config, not the output.

## Choosing the chart kind

| User asks for | kind |
|---|---|
| Bridge / walk / build-up / EBITDA waterfall | `waterfall` (`"e"` semantics via `waterfall.totalIndices`; several series → stacked waterfall) |
| Market map with variable column widths | `mekko` (add an `X extent` row for explicit widths) |
| Share-of-total over time | `stacked100` (optionally a `100%=` row) |
| Revenue by segment / stacked bars | `stacked` (`horizontal: true` for bars; `stack` per series for clustered-stacked) |
| Project / timeline slide | `gantt` (`Start`/`End`/`Milestone`/`After`/`Today`/`Holiday`/`Bracket …` rows; ISO dates supported) |
| KPI line over columns | `combo` (`"type": "line"` on a series; `secondaryAxis: true` for %-on-right) |
| Positioning / portfolio map | `scatter` or `bubble` (`X`/`Y`/`Size`/`Group` rows, auto label placement) |
| Population pyramid / two-sided compare | `butterfly` |
| Simple share | `pie` / `doughnut` (exact wedge geometry in the pptx output) |
| Side-by-side comparison per category | `clustered` |
| Distribution per category | `boxplot` (Min/Q1/Median/Q3/Max rows, or raw samples → Tukey whiskers) |
| Multi-dimension profile / score | `radar` (translucent polygons, ≤3 series) |
| Matrix of values (region × period) | `heatmap` (one global color scale) |
| Values by geography | `tilemap` (tile-grid cartogram: `map: "us"/"eu"/"europe"/"world"`; `world` is never auto-detected, so always set it) |
| Stage-by-stage breakdown of a total (contacts → answered → solved) | `cascade` ("Stage | Drop label | Group" categories) |
| Pipeline / conversion stages | `funnel` (centered bands + conversion %) |
| One dominant share ("68% of…") | `waffle` (10×10 unit grid; single category reads as a literal %) |
| Part-to-whole with many items / 2 levels | `treemap` (area ∝ value; `"Group \| Item"` categories nest) |
| Hierarchy as nested rings | `sunburst` (`"Group \| Item"` → inner + outer ring) |
| Full distribution per category (not just quartiles) | `violin` (raw samples per column, mirrored density) |
| Financial OHLC / price action | `candlestick` (`Open`/`High`/`Low`/`Close` rows) |
| Share with a long tail ("top 3 + the rest in detail") | `pie` + `pie.breakout: [indices]` (bar-of-pie) |
| Same chart repeated per series ("one panel per region") | any column/line/area/waterfall/radar kind + `multiples: {}` (shared scale) |
| Before/after comparison of a few series | `line` + `decorations.slope: true` (slope chart: end rails, labels both ends) |
| Trend over time | `line` (date categories space proportionally) or `stacked100`-style share via `area` |
| Rates that hold then jump (policy rate, tiers) | `line`/`area` + `decorations.stepped: "after"` (staircase) |
| P&L / net value over time that can go negative | `area` (negatives dip below the zero baseline) |
| One entity's path through an X/Y space over time | `scatter`/`bubble` + `decorations.trajectory: true` (Gapminder trail) |
| A third variable on a scatter/bubble (by color) | add a numeric `Color` row (sequential ramp + gradient legend) |
| Smooth trend curve instead of straight segments | `line` + `decorations.smooth: true` |
| Grouping a long waterfall into sections | `waterfall.spacerIndices` (blank gap columns) |
| Breaking one bridge step into "of which" sub-columns | `waterfall.detailGroups: [{ of: 2, indices: [3,4,5] }]` (off-chain; totals unaffected) |
| Distribution behind a boxplot (show the raw points) | `boxplot` (raw samples) + `boxplot.jitter: true` |
| Is the median difference significant? | `boxplot` (raw samples) + `boxplot.notch: true` (non-overlapping notches) |
| Summary bar per Gantt section | `gantt` with section-header rows + `decorations.summaryBars: true` |
| Plan grouped by who owns each task | `gantt` (`Activity \| Owner` categories) + `gantt.lanes: "owner"` |
| Resource capacity vs load per week | NOT a gantt — a `stacked` column chart (weeks = categories, one series per task) + a `Target` row for capacity |
| Cost / effort table beside a Gantt | `gantt` + one `Column <label>` row per column (e.g. `Column Cost k€`) |
| Plan where a bar's length means working days | `gantt` (dates) + `gantt.workdays: true` (or ISO weekday numbers for a custom week) |
| Where are the points concentrated (distribution beside a scatter) | `scatter`/`bubble` + `decorations.marginals: "x" \| "y" \| "both"` |
| Bubbles covering each other | `bubble` + `scatter.spread: "y"` (nudges along ONE axis, capped and disclosed in the footnote) |
| Groups must survive greyscale printing or color blindness | `scatter.markers: ["circle", "diamond", "triangle"]` (shape per `Group` row, cycled like the palette) |
| Diverging heatmap must survive greyscale printing | `heatmap.symbols: "sign"` (+/− per cell, where no value label carries it) |
| Benchmark / consensus / target per category over columns | `combo` + `"type": "marker"` on that series (points, no line — keep it on the columns' scale) |
| "Us vs the peer range" on a radar | `radar` (peers then us) + `decorations.radarBand: true` |
| Too many stack segments (long tail) | `stacked`/`clustered` + `otherBucket: {max}` (collapse the rest to "Other") |
| Daily activity over months (contributions grid) | `heatmap` (one daily date series, categories as date strings such as `2026-01-15`, not day numbers) + `heatmap.calendar: true` |
| Back-to-back comparison with sub-parts each side | `butterfly` + `butterfly.split: k` (stacked flanks) |
| Radar over mixed KPI units (shape comparison) | `radar` + `radar.perSpoke: true` |
| Line series with holes in the data | `line` + `decorations.bridgeGaps: true` (connect across nulls) |
| Floating bar / range column | `stacked` with a `color: "transparent"` base series |
| Bridge (waterfall) with a %-of-total line | `combo` + `combo.columns: "waterfall"` + a `type:"line"` series |
| Mekko with an overlaid line | `combo` + `combo.columns: "mekko"` |
| Several KPIs in unlike units on one chart | `combo` + `combo.lineAxes: "independent"` |
| Hexagonal tile map | `tilemap` + `tilemap.shape: "hex"` |
| A trend per region on a map | `tilemap` (regions × periods) + `tilemap.glyph: "bars"` |
| Scorecard gauge (one metric split) | `doughnut` + `pie.semi: true` (semi-circle) |
| 80/20 breakdown (defects, revenue) | `clustered` + `pareto: true` (sorted bars + cumulative line) |
| Rank changes over time | `line` + `decorations.bump: true` (bump chart) |
| Profile across many named items | `line`/`area` + `horizontal: true` (profile chart) |
| Which tasks drive the project finish date | `gantt` (with `After` rows) + `decorations.criticalPath: true` |
| Distribution summary by mean and spread (not quartiles) | `boxplot` (raw samples) + `boxplot.meanSd: true` |
| Word-sized inline trend / KPI dashboard rows | `line`/`area` + `decorations.sparkline: true` (pair with `multiples`) |
| Circular bar chart / Nightingale rose | `radar` + `radar.bars: true` (radius = value; series stack) |
| Part-to-whole across several dimensions | `radar` + `radar.stacked: true` (stacked radar) |
| Two metrics on a pie (share + a second measure) | `pie` + `pie.variableRadius: true` (or a `Radius` row) |
| Correlation / signed matrix (size + colour) | `heatmap` + `heatmap.sizeEncode: true` |
| Group similar rows of a matrix together | `heatmap` + `heatmap.cluster: true` (dendrogram) |
| Trend of a mix plus a KPI line | `combo` + `combo.columns: "area"` + a `type:"line"` series |

Decorations (INSIDE the `decorations` object): `totals`, `cagr {from,to}`,
`difference {from,to,series?,fromValueLine?}`, `valueLines`, `labelContent`,
`connectors` (segment-boundary lines between stacked columns),
`callouts` (speech-bubble comments), `bands` (shaded background regions),
`hundredPercentNote`, `stepped` (line/area staircase), `smooth` (line
curves), `bridgeGaps` (line across nulls), `trajectory` (scatter/bubble trail),
`summaryBars` (gantt sections), `radarBand` (radar peer range). These go at the
TOP LEVEL of the config, beside `decorations`, not inside it: `scale {min,max}`,
`axisBreak {from,to}`, `logScale`, `gapWidth`/`overlap` (Excel-style column
spacing), `categorySort`, `segmentOrder`, `footnote` (source line) and
`pie {explode}`; per-cell `colors` highlights go on a series. A key put in the
wrong object is ignored without an error. Defaults are
think-cell-like: labels on, axis off; always set `footnote` with the data
source when you know it.

## Rules

- Emit **valid JSON only** for configs; validate values are numbers (dates as
  ISO strings only in Gantt rows).
- Keep one message of data → one chart. For decks, pass an array.
- In the headless `.pptx` each chart IS one group carrying its config, exactly
  as the add-in draws it — so a deck rendered here opens re-editable in the
  pane. (This bullet used to say the opposite, and contradicted the usage note
  124 lines above it: grouping shipped in #228, and the committed showcase deck
  has one `p:grpSp` named `PowerChart` per slide.) An agenda slide is not a
  chart: it is loose, untagged shapes, as the pane inserts one.
- Prefer `waterfall` over generic columns whenever the story is a change
  decomposition. Prefer `mekko` when both share and size matter.
