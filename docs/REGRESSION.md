# Real-host regression: the demo deck

The **Insert demo deck** action (Automation tab) is SSF Charts' real-host
regression harness. It drops one slide per chart kind + feature + element onto
fresh slides, so a human — or a diff — can see what PowerPoint *actually* draws.
This catches host-only bugs the vitest fakes cannot (ellipse repaint, wedge
tessellation, slide-timing races), which is the whole point: the fake asserts our
intent, the deck asserts the host's behaviour.

The deck is **self-identifying and self-summarising**: it opens with a **title
slide** stamping the running build *and* the host (`Office.context.diagnostics` →
`PowerPoint · OfficeOnline · 16.0.x`), a **contents** slide indexing every chart
with its office-shape count, and closes with a **results slide** — a summary line
plus a table of only the skipped/failed items. So an exported PDF is a complete,
comparable record of one run without opening the console: which build, which host,
what failed, how long it took.

A stalled slide is NOT retried. It was, once, and that retry is where every
duplicate slide in this project came from: the readback that judged the first
attempt "short" ran while the host was still committing, so it was routinely
wrong, and the retry drew the same chart again on a new slide. Both then
landed. The settled repair pass at the end does the same job with evidence
instead of a guess — see section 3.

## 0. The independent check — audit the deck itself

```bash
npm run verify-deck -- path/to/Presentation.pptx      # or --json
```

Reads the saved `.pptx` and reports, per slide: slot, run token, whether the
chart is a group or a degraded picture, how many shapes are INSIDE that chart
object (a 40-shape chart and a 1-shape degraded picture both have one shape at
the slide's top level, so the inner count is usually the one wanted), whether
it carries a `POWERCHART_CONFIG`, and whether it wears a `NOT COMPLETE` banner. Exit 0 when
the file is structurally sound, 1 on a fault, 2 when it cannot be read. A slide holding
several charts is audited chart by chart: every config tag is checked, and the `chart
objects` total counts objects, not slides (the table's `2/3` is re-editable out of objects).

**Run this before believing the run report.** Every hard diagnosis in this
project was settled by the bytes rather than by the add-in's own summary —
including one where the summary was the thing that was wrong: a 39-slide web
run reported 20 tagged charts where the file provably carried 31. The log
agreed with itself and was false. Nothing in the add-in can catch that, because
the add-in is the thing under test.

What counts as a **fault** is deliberately narrow: an unreferenced tag part, a
tag part missing its `[Content_Types]` override, a config that is not valid
JSON, a config with no `PowerChart` object to load it from, or two slides
claiming the same run *and* slot. Those are things this repo wrote wrong. A
deck missing eight charts because the host dropped them is a bad RUN and a
perfectly well-formed FILE, and is reported without being called a fault — as
are two different runs' slides in one deck, which is the case the run token
exists to survive.

## 0b. The join — deck against run log

```bash
npm run triage -- Presentation.pptx powerchart-run-log.json   # --all, --json
npm run triage -- powerchart-round.json                       # no deck to give
```

The two files may be given in either order — the extension says which is which.
Piping `--json` anywhere needs `npm run --silent` (or `node scripts/triage.mjs`
directly), because npm prints its own banner to stdout ahead of the script.

**A round file on its own is the normal case now**, and the second form is not a
degraded mode. The starred runbook step saves the probe, the self-test and the
round's own deck evidence as one file precisely so there is nothing else to
send — no deck to save, no screenshot to take — and only the SLOT JOIN needs a
`.pptx`. Given a round alone this reports the self-test verdicts, the trace
tallies with their known-host-bug annotations, and what the round left in the
deck; it skips the join and says so. It used to answer `usage:` and exit 2,
which sent the reader back to reading 150 KB of JSON by hand.

The deck-evidence block counts an added slide as blank only when **two**
witnesses agree — the readback said zero shapes AND the host's own picture of
that slide is blank. One witness is not enough here: this host is known to
answer a shape collection short without throwing (`shapes-items-count-honest`),
and the rasterise pass is capped, so a slide with no picture is counted apart as
*not evidence* rather than folded in. A slide that reads back empty while its
picture has CONTENT is called out separately — that is the readback lying, and
it is the one shape of this that a screenshot alone would never show.

A log holds a **list** of runs, because one click can take both insert paths.
Each is reported separately against the same deck — merging them would produce
a report true of neither, since the file run's slides are present while the
shape run's are still missing, in the same file, at the same moment. Logs from
before the list existed are read as a list of one.

The two files a real run produces are the entire evidence base, and reading
them *together* is where the findings are: slot by slot, what the run believed
it did against what the file actually holds.

```
  run ms9fcxg6-k93mmn · build ba3365b · PowerPoint · OfficeOnline
  path shapes · 85.2s

  DECK 32 slide(s): 30 from this run, 1 from other run(s), 1 carrying no slot tag
  SLOTS 38 expected · 30 present · 24 no-config · 8 lost · 3 repaired · 3 ok

  #    title           log        tag?  deck                     verdict
  3    Stacked         rendered   no    picture 2sh config       repaired
  4    Clustered       rendered   no    —                        lost
  7    Mekko           rendered   no    picture 1sh no-config    no-config

  TRACE 276 entries
     117  demo   item finished
      79  group  tagging failed — charts are not re-editable until repaired
  problems:
      79  InvalidParam passed to GetItem(id) | code=5010
```

The verdicts, and what each means:

| verdict         | the run said              | the file holds          |
| --------------- | ------------------------- | ----------------------- |
| `ok`            | drew it, tagged it        | it is there, tagged     |
| `lost`          | drew it                   | no slide for that slot  |
| `duplicated`    | drew it once              | two slides claim it     |
| `tag-lost`      | wrote the config tag      | no config tag           |
| `not-editable`  | never got the tag written | a chart with no config  |
| `no-config`     | (log too old to say)      | a chart with no config  |
| `repaired`      | gave up on the tag        | tagged — the repair won |
| `blank`         | drew it                   | a slide with no shapes  |
| `orphan`        | never issued that slot    | a slide carrying it     |
| `skipped`       | skipped it                | nothing — as expected   |

A failure whose reason matches a known PowerPoint bug is annotated with the
issue number, so a reader stops looking for the mistake in this repo:

```
      79  InvalidParam passed to GetItem(id) | code=5010
          ^ known host bug: office-js#2903 (closed: not planned) — stale shape proxy on web
```

The first hour of one diagnosis went into establishing exactly that. The list
lives in `scripts/triage.mjs` and was checked 2026-08-01 — see `RESEARCH.md`
§4b before trusting a status.

`repaired` and another run's slides are **not** disagreements; they are the
repair pass and the run token doing their jobs. Exit 0 when the two accounts
agree, 1 when they do not, 2 when a file cannot be read.

Both this and `verify-deck` are deliberately `.mjs` tools outside `src/`, so
they cannot inherit a bug from the code they audit.

**Why it exists.** This join used to be done by hand — unzip the deck,
pretty-print 160 KB of log, line the slots up, squint — and twice it was done
wrong: once from a truncated read of the log, once from a trace buffer that
spanned three runs and showed a contradiction that was not there. Each wrong
turn cost a full round trip: a deploy, a run in a real PowerPoint, an upload
and a re-read. The join is mechanical, so it belongs in a script.

## 1. The cheap pass — self-check (every run)

Insert the deck. When it finishes, the pane reports and the **console** (F12)
prints a per-chart table plus the run's integrity numbers:

```
chart        shapes  status    grouped   ms
Bubble         44    rendered     true   180
Combo          22    failed      false 45012   ← host stalled mid-draw (near the 45s timeout)
Doughnut       15    rendered     true   240
Area            0    skipped     false     2   ← too dense, stamped
deck grew by 33, issued 35 adds — 2 LOST; blank slots 24, 34 · total 78.4s
```

Read `insertDemoDeck`'s `DemoReport` (`src/render/powerpoint.ts`): `results[i]` is
`{created, status, ms, grouped}`; `slidesAdded` is the deck's ACTUAL growth (settled
`getCount`, after − before); `totalMs` is the whole run's wall-clock.

**Lost slides — measure against `addsIssued`, not `items.length`.** A failed item
can leave a **stray** slide, so `slidesAdded` can equal `items.length` even when
the host lost a real slide — the stray cancels it. So loss is
`addsIssued − slidesAdded`. A stray that LANDED cancels; a swallowed/lost add
does not — so this reads through the coincidence.
(A real run once lost 2 slides and reported 0 under the old `items.length` formula.)

**`blankSlides`** is the list of **1-based deck positions** of added slides that read
back with **zero shapes** — the host kept the slide but its content detached. It is
reported **by position, never by item name**: a blank slide has no content and no
config tag, so it cannot be attributed to an item, and under load the host reorders/
merges/loses slides, which breaks any positional item mapping anyway. Each `0` is
re-read once (a struggling host reports transient `0`s) before it counts, and
`blanksRead` is `false` if the readback faulted — so an empty list is never mistaken
for "no blanks" when it means "not fully measured".

Honest limits of `blankSlides`: it cannot see a **merge** (two items on one slide —
that slide isn't blank), loss *inside* a group, or a **paint-only** blank
(office-js#2699 — the shapes exist, so `getCount > 0`). Naming the missing/merged
charts by their config tag (`CHART_TAG`, deck-wide, order-independent) is a
**documented follow-up**, not yet built.

This pass needs no PDF and catches structural regressions: skipped, failed, lost
(via `addsIssued`), and empty slots (via `blankSlides`).

It does NOT catch *paint* bugs — a shape created but not rendered (office-js#2699)
still counts. That's what the visual pass is for.

## 2. The visual pass — PDF diff (when paint correctness matters)

Export the deck to PDF, then compare each page to its SVG reference (the SVG
renderer is the source of truth — see the byte-identical snapshot invariant).

```
# render the PDF pages
python -c "import fitz; d=fitz.open('deck.pdf'); [p.get_pixmap(matrix=fitz.Matrix(2,2)).save(f'p{i+1:02}.png') for i,p in enumerate(d)]"
# map pages -> chart by title
python -c "import fitz; [print(i+1, p.get_text().split(chr(10))[0]) for i,p in enumerate(fitz.open('deck.pdf'))]"
# render references: sceneToSvg(buildChart(sampleConfig(kind))) -> Chrome --screenshot
```

**Pair by TITLE, not page number** — a lost or misaligned slide shifts the page
order (map the title text on each page to the chart).

**Two artifacts that WILL bite you** (both flagged bogus "defects" on a real run):

- *Framing.* The rendered chart sits in a *sub-region* of a 960×540 slide (placed
  at ~60,90 with margins), while a raw SVG reference fills its own frame. So
  "rendered is smaller / shifted down-right" is not a defect — it flagged 15
  charts once. Either place the reference on a 960×540 canvas at the same offset,
  or tell the reviewer to ignore absolute scale/position and judge only data,
  proportions, colour, and completeness.
- *Title text.* `demoItems()` overrides each sample's title with a short label
  (`{...sampleConfig(kind), title: label}`), so a reference built from
  `sampleConfig(kind)` shows the sample's LONG title — a "title differs / colour
  lost" mismatch that isn't real. Build references with the SAME title override
  the demo uses.

Net from the run that produced this doc: of ~20 "findings", exactly ONE was a real
render bug (doughnut arc gaps); the rest were these two artifacts, known host
limits, or the stamp/stall. Distrust the diff; verify each finding against the
scene before filing.

## Known host limitations (not regressions)

- **Filled polygons render as outline only** (radar/violin fills missing) — Office.js
  has no freeform fill; the SVG reference fills them, the host can't.
- **Dense charts (>~90 shapes) are skipped + stamped** on web (area, tile map,
  waffle, smoothed line) — they exceed the per-slide shape budget.
- **Freshly-inserted shapes may not repaint** until zoom/navigate (office-js#2699).

Relevant OfficeDev/office-js issues: #2699 (repaint), #2903/#2474 (slide-add
timing / non-round-trip ids), #4272 (>50-item load hang). See `docs/repro/` for a
Script Lab repro of the ellipse case.
