# Deep research: think-cell, and how to clone it

This report condenses a fan-out research run (5 search angles, 22 sources fetched,
59 claims extracted, top 25 adversarially verified 3-0 each — 0 refuted) into the
findings that shaped SSF Charts. Confidence is high for everything below unless
flagged; caveats at the end.

## 1. What think-cell actually is

- **A native C++ COM add-in, ~1M lines of C++**, for installed desktop
  PowerPoint/Excel only — there is no web/Office.js think-cell. Its own
  engineering page states a dedicated reverse-engineering team (IDA/Hex-Rays)
  built a function-hooking engine that **patches Office executables in memory at
  every startup**, locating assembly signatures to survive Office updates. That
  is why think-cell breaks on new Office builds until compatibility updates ship
  — and why no sandboxed add-in can replicate its integration depth.
  *(Sources: think-cell.com/en/career/tech; corroborated by their KB breakage
  pattern. LOC figure is self-reported.)*
- **The founding thesis is layout, not charting.** think-cell began (2002,
  Fraunhofer FIRST spin-off) as an algorithm for automatic slide layout; the
  company names its hard problems as layout quality, algorithms fast enough for
  interactive editing, and a UI to match. Its core IP is automatic label
  placement (US Patent 7,292,244; Müller & Schödl 2005 on scatter-chart
  labeling). **A clone's differentiating work is the layout/labeling engine,
  not drawing rectangles.**

## 2. Feature model (verified specifics)

- **One unified stacked chart type.** think-cell does not distinguish simple vs
  stacked column/bar charts — a simple chart is a stacked chart with one series.
  Bar charts are *rotated column charts*: users drag a rotation handle to
  convert column↔bar (works for Mekko and waterfall too), and butterfly charts
  are two back-to-back bar charts, one rotated 180°, with the same scale applied.
  → One stacked engine + an orientation transform covers column/bar/butterfly.
- **Mekko math.** Mekko with %-axis = a 100% stacked column chart whose column
  widths are proportional to column totals, so **segment area ∝ absolute
  value**. A second variant, "Mekko with units," takes explicit column widths
  from an `X extent` datasheet row scaled to the chart width.
- **Datasheet semantics.** The internal Excel-like datasheet opens automatically
  on insertion and reopens on double-click; supports **transpose rows/columns**;
  values are absolute by default with each column summing to 100% for percentage
  math; a special **`100%=` row** lets users enter percentages and have absolute
  values derived.
- **Labels and totals are automatic**, placed collision-free at creation with no
  user action; totals can be toggled via "Add Total"; segment stacking order is
  user-controllable (sheet order / reversed / ascending / descending).
- **Axes are direct-manipulation**: auto-scaled by default, rescaled by dragging
  handles at axis ends; **axis breaks** are inserted at the mouse position and
  resized by dragging the compressed range's bounding lines; **Same Scale**
  links scales across multiple charts, re-adjusting all of them whenever any
  chart's data changes.
- **Decorations are live data-bound objects, not static annotations.** Two kinds
  of difference arrows — *level* (between segments/data points) and *total*
  (between column totals). CAGR arrows compute from column totals, taking the
  date range from the datasheet cells behind the category labels. Value lines
  are horizontal lines at a value (several per chart allowed), and difference
  arrows may anchor to segments, data points, **or a value line**. Everything
  recomputes on every data edit.

## 3. Ecosystem: there is no open-source clone

- **No open-source reimplementation of think-cell's rendering exists.** The
  closest project, **ThinkcellBuilder** (Python, GitHub), only automates data
  entry: it emits `.ppttc` files — think-cell's official JSON automation format
  (IANA-registered `application/vnd.think-cell.ppttc+json`, published JSON
  schema) — and still requires a licensed think-cell install to render.
- Lesson 1: the rendering pipeline is closed; a real clone must build its own.
- Lesson 2: the **JSON-in → deck-out automation surface is a proven, valued
  interface** worth replicating (SSF Charts' `ChartConfig` is JSON-shaped for
  exactly this reason).
- Commercial alternatives (UpSlide, Empower, Aploris, Vizzlo, Zebra BI) exist,
  but no specific claims about their internals survived verification — treat
  anything written about them as unvetted.

## 4. Office.js feasibility (what SSF Charts is built on)

- **Shape construction is fully supported**: `ShapeCollection.addGeometricShape`
  (GeometricShapeType + position/size in points from the slide's top-left),
  `addLine` (ConnectorType; left/top is the start point, width/height the delta
  to the endpoint — negatives allowed), `addTextBox`, plus fill/line formatting
  and text via `textFrame.textRange`. Grouping: `addGroup` / `ShapeGroup.ungroup`.
- **Requirement sets**: shape creation/format/delete arrived only in
  **PowerPointApi 1.4** (Windows 2207+, Mac 16.62+, web; not iPad); **grouping
  needs 1.8**; **1.3 added persistent string tags** on presentations, slides and
  shapes — the mechanism for storing a serialized chart model on the generated
  group so charts stay re-editable. `insertSlidesFromBase64` (1.2) is a limited
  OOXML fallback for older hosts, but without in-place editing.
- **Key limitation**: Office.js lines are positioned geometry only — there is no
  API to glue endpoints to shape connection sites (unlike VBA's `AddConnector`),
  so "universal connectors" require the add-in to maintain its own anchor logic.
  No freeform/path shapes either (hence triangle-shape arrowheads).
- **Verified architecture conclusion**: render charts as grouped native shapes,
  persist the data model in tags, recompute layout and re-emit shapes on every
  edit. This is exactly SSF Charts' pipeline.

## 4b. Re-check, 2026-08: is the Office.js approach still right?

The original §4 was desk research before anything shipped. This is a re-check
after nine real-host runs, ~20 host-behaviour bugs and a targeted search of
Microsoft's own issue tracker. **Verdict: the architecture holds, and the
evidence is stronger than the original guess.**

### Every hard workaround maps to an open Microsoft bug

None of the defensive machinery in `powerpoint.ts` is compensating for our own
misuse. Each of these is a PowerPoint-on-the-web defect, reported by others,
still open:

| Behaviour we defend against | Issue | Status |
|---|---|---|
| `InvalidParam passed to GetItem(id)` — a shape proxy goes stale across a sync | [office-js#2903](https://github.com/OfficeDev/office-js/issues/2903) | closed **"not planned"** |
| `context.sync()` hangs forever after add → delete → re-read shapes | [office-js#5022](https://github.com/OfficeDev/office-js/issues/5022) | **WITHDRAWN, not fixed** — closed 2024-11-18 after the reporter found his own selection handler was opening a parallel `PowerPoint.run`. Not evidence of a host defect, and our probe has never seen the symptom (0 `silent` in 401 rounds) |
| Loaded properties silently unavailable after `sync()` (our "hollow reads") | [office-js#6363](https://github.com/OfficeDev/office-js/issues/6363) | labelled **regression + product bug**; reporter tried ten approaches, none worked |
| `sync()` hangs past ~51 items in one `load()` | [office-js#4272](https://github.com/OfficeDev/office-js/issues/4272) | open — this is why `READBACK_PAGE` is 20 |
| `setSelectedShapes([])` does not clear the selection on the web | [office-js#3083](https://github.com/OfficeDev/office-js/issues/3083) | open — why `clearShapeSelection` re-selects the *slide* instead |
| A picture cannot be inserted while a shape is selected | [office-js#3698](https://github.com/OfficeDev/office-js/issues/3698) | open — why every selecting scenario cleans up after itself |

| A deck insert's `sync()` never resolves — **and the slide lands anyway** | [office-js#1650](https://github.com/OfficeDev/office-js/issues/1650) | marked fixed, but see below — the *shape* of it drives the design |

The `hollowReads` fault in `test/helpers/office-host.ts` is not a paranoid
invention: it models a regression Microsoft has acknowledged and not yet fixed.

### A shape can be WRITTEN through its creation handle, but never read back

Measured on the real host, rounds `032-360309f` and `033-86eaf65`. Answers that
only mean something read together:

| question                              | 032          | 033          |
| ------------------------------------- | ------------ | ------------ |
| `tags-on-fresh-shape`                 | `yes`        | `yes`        |
| `tag-through-refetched-shape`         | `no-id`      | `no-id`      |
| `tag-the-creation-proxy-a-sync-later` | —            | **`yes`**    |
| `shape-proxy-survives-one-sync`       | `unreadable` | `unreadable` |
| `shapes-items-count-honest`           | `unreadable` | `unreadable` |

**Writing through an aged creation handle works. Reading through one does not,
and there is no id to re-fetch by.**

This section first said both routes out of the creating batch were closed, which
was wrong and lasted about an hour. Round 032 answered `no-id` and `unreadable`
and that was read as "the shape is unreachable once its batch is over". Round
033 asked the half `no-id` had left untried — tag through the handle that made
the shape, one sync later — and the write went through. Recorded because the
overstatement would have sent the fix at the wrong half of the path.

So the rule is narrower and more useful than "tag it while you are making it":

- `shape.tags.add(...)` through the handle that CREATED the shape works, in that
  batch or a later one.
- `shape.load(...)` through that same aged handle comes back unreadable.
- `shapes.getItemOrNullObject(id)` is not available at all, because the id never
  reads back.

That is the whole of the `same scale across the deck` failure, which three
consecutive rounds scored 4-5 of 8 on. `finishCharts` writes `POWERCHART_CONFIG`
in a LATER batch than the one that drew the shapes, reaching back by id or by
proxy, and this host has neither to offer. The 5010 it reports —
`InvalidParam passed to GetItem(id)`,
[office-js#2903](https://github.com/OfficeDev/office-js/issues/2903) — is the
same story from the host's side: Office.js rewrites a resolved proxy's object
path to `getItem(id)`, so even the fallback that never mentions an id has become
an id lookup by the time the host sees it.

It is not absolute. Half the charts do get their tag, and the probe records a
`regime` beside each answer, so the host moves in and out of the state where
collection reads work at all. Which is why no rewrite is here yet:
`tag-through-refetched-shape` is marked `resample`, and its partner
`tag-the-creation-proxy-a-sync-later` asks the half that `no-id` leaves untried.
Two more rounds saying the same thing turns one measurement into a rule worth
rewriting a critical path for.

### A hung sync says nothing about whether the work happened

This is the most useful thing the tracker taught this project, and it changed
how every bounded write is written. office-js#1650, verbatim: *"the first time
`context.sync()` is called the promise resolves, but in subsequent calls the
promise doesn't resolve, although **the slide still gets added successfully**."*
office-js#5022 reports the same shape for shape work after an image insert, and
its only known workaround is a pause: *"I had better result by adding a timer of
1-2 seconds between the `shape.delete()` and the next `await context.sync()`."*

Every timeout in `powerpoint.ts` used to throw. For a **read** that is correct —
an unread page is unread, and each caller already has an honest word for it
(`unread`, `undetermined`, `unmeasured`). For a **write** it is wrong twice: it
discards work that landed, and it sends the caller off to do the work again,
which is how one stalled insert becomes two copies of a chart on one slide.

So a bounded mutation now swallows its own silence and lets a fresh-context
re-read decide (`withTimeoutOrVerify`). Every site that uses it already measured
what actually happened — `slideCount()` before and after, a re-read of the
slide — and simply never reached the measurement on the runs that needed it.
A **refusal** still throws: the host said no and is still talking, which is a
different fact and a different diagnosis.

`#1650` is marked fixed upstream. It is kept here because the pattern is not:
the fix does not make silence mean failure, and #5022 is open with the same
shape.

### Why the file path cannot become the DEFAULT for editing a chart

It is tempting, and it was proposed. The shape path carries all four bugs
above; the file path carries none of them. `updateChartInSlide` redraws a chart
shape by shape, which is the add-in's worst case on the web, while
`replaceSlideWithDeck` is two host calls. On a slide holding nothing but our
chart, rebuilding the slide from a generated `.pptx` looks strictly better.

**It is not, and the reason is a platform gap rather than a preference.** A swap
*replaces the slide*. The chart comes back; everything the slide carried that is
not a shape does not — speaker notes, the transition, animations, slide-level
formatting. `slideHoldsOnlyChart` checks that no other **shapes** are present,
and that is the most it can ever check: **Office.js exposes no way to read
speaker notes at all** ([office-js#3269](https://github.com/OfficeDev/office-js/issues/3269),
in backlog; the feature has been asked for repeatedly). The add-in cannot see
what the swap would destroy, so it cannot warn first and cannot decline.

That is survivable where the swap lives today — a fallback reached only after
the in-place redraw has already stalled, where the alternative is rasterising
the chart to a picture, which is lossy too and worse. It would not be
survivable on the happy path: every ordinary chart edit would silently discard
the user's notes.

So the swap stays a fallback, and it now *says* what it cost. Revisit if
#3269 ever ships — reading notes is the only thing standing between this and a
much faster, much less bug-prone edit path.

**One more, measured here rather than found in the tracker.** On PowerPoint on
the web, a *programmatic* `Slide.setSelectedShapes([id])` — GA since
PowerPointApi 1.5 — is accepted without complaint and then leaves the selection
subsystem unable to answer anything: `getSelectedShapes` ran out a full
90-second budget, and the `setSelectedSlides` behind it did the same, in the
same run, twice. Nothing throws; the host simply stops replying, which is a
failure shape no amount of error handling catches and only a bounded wait
survives. It reads as the same family as #3083 and #3698 — the web host's
selection layer accepting writes it cannot then serve — and it is why the
battery's selection scenario reports **skipped** on that host rather than red
(`faults.selectionWedgesHost` models it; `docs/PUBLISHING.md` says what to
expect on screen). It costs the add-in nothing, because the pane never selects
a shape from code: it reads the selection the user made with a click, and that
path answers normally.

**The resource-limit reading is confirmed.** Microsoft's current
[resource limits doc](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/concepts/resource-limits-and-performance-optimization.md)
scopes the CPU / memory / four-crashes / five-second ceilings to *"Windows and
Mac only… not included: mobile apps or browser versions."* Nothing throttles a
runaway add-in on the web — the tab simply dies, which is what happened at
~1850 shapes. `DEMO_SHAPE_BUDGET` exists because no platform limit does.

### The file-first default is the decision that matters

All four bugs above live on the **shape-by-shape** path.
`insertSlidesFromBase64` is one call and touches none of them. The measured gap
is 5.8–6.5 s with zero loss versus 85–118 s and 3–8 items short, on the same
host in the same session — but the reliability gap matters more than the speed
gap: it is the difference between using four broken APIs and using none.

One caveat to carry: [office-js#2780](https://github.com/OfficeDev/office-js/issues/2780)
and [#5896](https://github.com/OfficeDev/office-js/issues/5896) report
`insertSlidesFromBase64` losing source formatting, both closed without
resolution. Reading them, the complaints concern *theme-inherited* formatting
when copying slides between decks. SSF Charts emits explicit formatting on
every shape, which is why its runs land clean — so **do not drift toward
theme-inherited styling in the generated deck**; that is what would expose us
to this.

### Alternatives, assessed and rejected

| Approach | Verdict |
|---|---|
| **OOXML injection** (`setSelectedDataAsync` + `CoercionType.Ooxml`) | **Does not exist for PowerPoint.** Word and Word Online only. The obvious escape hatch is not there |
| **SVG coercion** (`CoercionType.XmlSvg`) | Rejected in `BACKLOG.md`, and the reasons check out: writes to the *selection* only, [#2881](https://github.com/OfficeDev/office-js/issues/2881) renders complex SVG wrong, [#3309](https://github.com/OfficeDev/office-js/issues/3309) cannot read it back |
| **Freeform / custom geometry** | Still absent. The API offers exactly three creators — `addGeometricShape`, `addLine`, `addTextBox`. The triangle-fan pie is not a workaround chosen over something better; it is the only option |
| **Copying a shape between slides** | **No API.** `Shape.Duplicate` / `ShapeRange.Copy` are VBA only; the Office.js `shapes` collection is read-only and there is no `copyTo`. Any "build it on a temp slide and paste it" plan is unbuildable — see below |
| **Native PowerPoint chart objects** (`c:chart`) | Real, and pptxgenjs can emit them. Rejected on product grounds, not technical: a native chart cannot carry think-cell decorations, which is the entire point |
| **`customXmlParts`** | Available for PowerPoint, and the right tool for *deck-level* data — but not for chart config. See below |

Microsoft's own performance guidance — minimise `sync()`, batch loads,
`untrack()` proxies — is what `SHAPES_PER_SYNC`, `READBACK_PAGE` and the
untrack calls already implement. We are not fighting the platform's advice; we
are following it and hitting its bugs anyway.

### Putting a chart on the slide the user is already on

There is no shape copy/paste in Office.js, so the tempting workaround — build
the chart on a scratch slide via the reliable file path, copy it onto the
current slide, delete the scratch slide — **cannot be built**. There is no
paste step to write.

What remains for an existing slide:

1. **`insertSlidesFromBase64` with `targetSlideId`** places the generated slide
   immediately after the current one. Not *on* the slide, but the reliable path
   and the correct default for "add a chart here".
2. **Insert as a picture** — one shape, one sync, so it avoids all four bugs
   above by construction. "Explode to native shapes" is then the user's
   informed opt-in to the risky path.
3. **`insertSceneIntoSlide`** — drawing shape by shape onto the live slide. The
   only way to get native shapes onto a slide we did not generate, and
   consequently where the bugs concentrate.

### Should `customXmlParts` get a spike?

**Not for chart config.** A chart's config must travel *with the shape*, so it
survives copy/paste into another deck and PowerPoint's own Duplicate Slide.
Shape tags do that; a presentation-scoped XML part cannot.

**Not as a repair-pass manifest** either, though it is tempting: one read
instead of paging tags off every slide would dodge #4272 and #6363. But such a
manifest is a *cache*, and it goes stale the moment the user deletes a slide,
duplicates one, or edits the deck in another app. A stale cache driving a pass
that **deletes slides** is precisely the failure this project has spent its
time eliminating. It would still have to be verified against the slides — which
is the expensive part it was meant to avoid.

**Yes, for deck-level style.** This is a real gap with no alternative: the
imported style file and saved templates live in `localStorage`
(`powerchart-style`, `powerchart-templates`), so they follow the *browser*, not
the deck. Send a branded deck to a colleague and their charts do not match
yours. A presentation-scoped custom XML part is exactly the right home for
"this deck's chart style".

One trap if that spike happens: Office.js enumerates only custom XML parts
related from `ppt/presentation.xml`. Parts written at the package root
(`/customXml/itemN.xml`) are invisible to it — which is how a
[reported "missing parts" bug](https://learn.microsoft.com/en-us/answers/questions/2149356/powerpoint-office-js-customxml-api-does-not-return)
turned out to be a scoping rule rather than a defect. A part written into our
generated deck must be related from the presentation, not dropped at the root.

## 5. How the findings map to SSF Charts

| Finding | Status in SSF Charts |
|---|---|
| Layout/labeling engine is the core IP | Pure-TS engine in `src/core`, unit-tested |
| Native grouped shapes + tags persistence | `src/render/powerpoint.ts` (tags: `POWERCHART_CONFIG` on the group) |
| Unified stacked model | `layoutColumns` handles stacked/clustered/100% from one code path |
| Mekko: width ∝ total, area ∝ value | `layoutMekko` implements the %-axis variant |
| Waterfall with computed totals | `e` cells → running-total bars + dashed connectors |
| Live decorations from anchors | CAGR/difference/value-line computed from `LayoutAnchors`, re-derived on every render |
| Datasheet: auto data grid + paste | Task-pane grid with Excel TSV paste |
| JSON automation surface | `ChartConfig` is plain JSON in/out |

**Parity ledger (final).** Everything feasible from the research is built —
rotation (column↔bar for columns/waterfall/Mekko/boxplot), butterfly, Mekko
with units, datasheet transpose + `100%=`, level/total/value-line-anchored
difference arrows, multiple value lines, pinned scales, axis breaks, Same
Scale, segment order, Gantt, agenda, plus chart types think-cell lacks
(boxplot raw-sample mode, radar, heatmap, tile-grid maps). What remains
un-cloneable from a sandboxed add-in is the trio think-cell memory-patches
Office for: live Excel data links, in-canvas drag manipulation, and the
slide-layout engine. The README feature table is the authoritative list.

## Caveats

- think-cell.com blocks automated fetches (403), so manual quotes were verified
  via search-index snippets; some cited paths may have drifted.
- The C++/reverse-engineering claims come from think-cell's own recruiting pages
  (self-reported, behaviorally corroborated).
- Gantt/agenda/process-flow, smart-labeling internals, Excel data links, and
  commercial-competitor details produced no *verified* claims — absence of
  verification, not absence of the features.
- How think-cell physically represents charts inside `.pptx` (grouped shapes vs
  OLE vs custom XML parts) remains an open question; likewise the exact
  algorithms behind interactive-speed label placement, and how Excel data links
  could work from a sandboxed PowerPoint add-in.
- §4b's issue links were read 2026-08-01. All were open then except #2903,
  which is closed "not planned" — a decision rather than a fix. Re-check before
  assuming any of them has been resolved — and if #6363 or #5022 ever are, the
  shape path's cost/benefit changes and is worth re-opening.
- The selection wedge in §4b has **no issue of its own**. It was measured here,
  on build `55011a3`, and filed against no tracker — so unlike everything else
  in that table it rests on this project's own evidence. Worth reporting
  upstream, and worth re-measuring before treating it as permanent.

## Primary sources

- think-cell manual: column/line/area, Mekko, axes, chart decorations, data
  entry pages (think-cell.com/en/resources/manual/…)
- think-cell engineering: think-cell.com/en/career/tech
- Microsoft Learn: PowerPoint add-in shapes guide, `ShapeCollection` /
  `ShapeGroup` API reference, PowerPoint API requirement sets, tags guide
- github.com/Philistino/ThinkcellBuilder; static.think-cell.com/ppttc/ppttc-schema.json
- Slide Science tutorials (UX corroboration); Peltier Tech (Marimekko geometry)
