# AppSource store listing — copy & submission checklist

Everything needed to submit SSF Charts to **Microsoft AppSource** (via Partner
Center). This is the *public store* path — heavier than sideloading or an
org-wide admin deploy (see `PUBLISHING.md` "Distribution beyond sideloading").

> **Trademark rule for anything store-facing:** the internal docs describe
> SSF Charts as a "think-cell clone / think-cell-style" tool, but the public
> listing, name, description, and screenshots must **not** use the "think-cell"
> mark as branding. Describe the features generically. A nominative disclaimer
> ("not affiliated with … think-cell") is fine and lives on `/terms.html`.

## Listing copy (paste into Partner Center)

**Name:** SSF Charts

**Subtitle / summary (≤ ~100 chars):**
> Consulting-grade charts as native, editable PowerPoint shapes.

**Short description (≤ ~1,000 chars):**
> SSF Charts turns a simple data table into the charts that consultants and
> analysts actually use — waterfall/bridge, Mekko/Marimekko, stacked and
> clustered columns, 100% charts, lines, areas, pie/doughnut, scatter/bubble,
> Gantt plans, and more — inserted onto your slide as **native, fully editable
> PowerPoint shapes** you can recolour, move and restyle by hand. Add the
> annotations that tell the story: CAGR arrows, difference arrows, value lines,
> automatic column totals, and collision-avoiding labels. Every chart stays
> re-editable: reopen the pane, change the data, and it updates in place. A few
> of the densest layouts — tile maps, large area charts — are more shapes than a
> browser will take, so in PowerPoint on the web those arrive as a picture
> instead; they stay just as re-editable from the pane, and the desktop apps
> draw them as shapes like everything else.

**Long description — feature bullets:**
- 25 chart kinds incl. waterfall bridges, Mekko/Marimekko, stacked/clustered/100%,
  cascade, funnel, butterfly, Gantt, combo, scatter/bubble, boxplot, violin,
  candlestick, radar, heatmap, tile map, treemap, sunburst and waffle.
- Native PowerPoint shapes — recolour, move or restyle any element by hand, in
  the desktop apps and, for all but the densest layouts, on the web.
- Signature annotations: CAGR & difference arrows, value lines, totals, smart
  labels with a global de-overlap pass.
- Re-editable charts, saved templates, and an import/export style file for a
  consistent corporate look.
- Runs entirely in your client — your data never leaves your device.

> **"NEVER FLAT PICTURES" WAS REMOVED FROM THE COPY ABOVE ON 2026-09-10, AND IT
> MUST NOT COME BACK.** It was not true on two of the four platforms Microsoft
> says it will test. Below PowerPointApi 1.10 the add-in deliberately inserts a
> **picture** for any chart carrying marks the host cannot draw, and tells the
> user so: *"This PowerPoint cannot draw … — that needs a newer Office
> (PowerPointApi 1.10). Inserted as a picture so it is complete; it stays an
> image here."* (`pictureForUndrawableMarks`, `src/taskpane/app.ts`). The repo's
> own measurement: **18 of 123 shipped charts lose ink below 1.10 and 9 lose
> their subject entirely** — pie 4/4, doughnut 2/2, sunburst 2/2, and the
> radial-bar radar, showcase #107, whose eight bars ARE eight wedges.
>
> **That ninth was published as an 8 until 2026-09-13.** Radar was filed under
> the charts that "lose annotation arrows and keep their marks"; it has no
> arrows at all. The count is now re-derivable rather than remembered —
> `npm run mac:webkit` prints it, and the recipe is `buildChart` over
> `examples/showcase.json` counting `wedge` and `arrowhead` nodes.
>
> **THE GROUND UNDER THIS NOTE MOVED ON 2026-09-13, AND THE BAN IS NOW THE
> OWNER'S TO LIFT RATHER THAN A STANDING FACT.** The floor went 1.4 → 1.8 →
> **1.10** that day. 1.10 is exactly the set that carries `Shape.rotation`, so
> no host the manifest now admits can take the undrawable-marks path at all:
> `canRotate()` is true everywhere, `marksThisHostWillDrop` always returns
> empty, and `pictureForUndrawableMarks` is unreachable. On every platform
> Microsoft will now certify, a chart arrives as native shapes.
>
> So the sentence is no longer false. It is also **not automatically restored**:
> whoever puts it back must say it in words the OTHER two picture paths do not
> contradict — the density rescue (`wantsAutoPicture`, which rasterises a chart
> with more shapes than the host can swallow, and has nothing to do with
> requirement sets) and the user's own "insert as image" choice. "Never flat
> pictures" is still wrong for those two. Something like *"charts arrive as
> native, editable PowerPoint shapes"* is true and does not promise that no
> picture can ever appear.
>
> Kept rather than deleted because the reasoning is the valuable part: the
> claim was struck for a real reason, and it is being reconsidered because the
> product changed, not because anyone got tired of the constraint.
>
> The BEHAVIOUR is good and worth describing plainly rather than hiding: the
> user gets a complete chart plus a message naming what would otherwise be
> missing, instead of a silently broken one.
>
> **AND THE FIRST FIX HERE WAS INCOMPLETE, which is worth recording.** Striking
> "never flat pictures" left two sentences making the same promise in other
> words — the short description's "inserted onto your slide as native, fully
> editable PowerPoint shapes" and the bullet "Native PowerPoint shapes —
> recolour, move, or restyle any element by hand". Both are now qualified. If
> the unqualified guarantee is ever wanted back, the honest route is to raise
> the manifest's `MinVersion` to 1.10 so it is TRUE, not to delete the
> qualifier: read every sentence in this file, not just the one that named the
> word "picture".
>
> **AND THE SECOND FIX QUALIFIED THE RIGHT SENTENCES ON THE WRONG AXIS.
> CORRECTED 2026-09-20.** Both qualifiers it added were written on 2026-09-10 —
> three days before the floor moved — and both keyed off HOW OLD THE BUILD IS:
> *"(on PowerPoint on the web and Microsoft 365 builds from 2026; older builds
> get a complete chart with a note about what they cannot draw)"* and *"Native
> PowerPoint shapes on the web and current Microsoft 365"*. At a 1.10 floor that
> is wrong twice over, and the second way is the one a certifier would hit.
>
> **An older build cannot install this add-in, so it cannot have the degraded
> experience the copy promised it.** The floor is 1.10 and
> `pictureForUndrawableMarks` is unreachable above it — the paragraphs above are
> the argument. So that clause was not merely inaccurate: there is no population
> it is about. It reads as a support commitment to users Microsoft will not let
> install the product, and it points a reviewer at a code path that cannot run.
>
> **And it named the web as a safe host when the web is the ONLY host with a
> picture path left.** `wantsAutoPicture` returns false unless `opts.web`
> (`src/render/powerpoint.ts`) and `isWebHost()` is `OfficeOnline` alone
> (`src/taskpane/app.ts`), so the density rescue fires on PowerPoint on the web
> and nowhere else. The old qualifier named exactly the wrong platform — and the
> web is where most reviewers test.
>
> **THE MEASUREMENT — and the two audit readings of it, "Area and Tile map" and
> "9 of 123", ARE BOTH RIGHT ABOUT DIFFERENT POPULATIONS.** Same predicate the
> pane uses either way: `estimateOfficeShapes(buildChart(cfg)) >
> DEMO_SHAPE_BUDGET` (105), gated on `isWebHost()`. Measured 2026-09-20:
>
> - **What a user or a reviewer actually meets: 2 of the 25 kinds.** The picker
>   loads `sampleConfig(kind)`, and at its 480×300 default exactly **Tile map
>   (122) and Area (111)** cross 105. Everything else is shapes — waffle 103 and
>   sunburst 101 sit just under. THIS is the set the listing copy should name,
>   and it is why the copy above names tile maps and area charts.
> - **What the shipped deck holds: 9 of 123 shipped charts insert as a picture.**
>   Over `examples/showcase.json` —
>   the same denominator the 1.10 census uses — tile map 4/4, area 2, waffle 1
>   (#61 at 107, not the 103 of a 10×10 grid), line 1 (the Catmull-Rom smoothed
>   one, 226) and combo 1 (stacked area + margin line, 164). The showcase carries
>   denser instances than the samples do, two of them in kinds the samples keep
>   well under budget.
>
> Neither figure is "the" answer, and the gap between them is the useful part:
> **the gate is a shape count, not a list of kinds**, so a user's own data can
> push a kind over that its sample never approaches. None of this happens on the
> desktop apps at all unless the user ticks *Insert as picture*.
> `test/web-density-census.test.ts` re-derives both numbers and asserts this file
> still says them, exactly as `below-1-10-census.test.ts` does for the 1.10 one.
>
> **THE TWO NINES ARE DIFFERENT NINES, and one is going to get quoted for the
> other.** Nine charts lose their SUBJECT below 1.10 — a population that can no
> longer install. Nine charts arrive as a PICTURE on the web — a population that
> is every reviewer. They share not one chart: the first nine are pies,
> doughnuts, sunbursts and a radial-bar radar, none of which reaches 105 shapes;
> the second nine contain no wedge at all. If you catch yourself writing "the
> nine charts", say which.

> **Where "25" comes from, so it cannot go stale unnoticed:** it is the number of
> entries in `CHART_KINDS` (`src/core/samples.ts`), which is what the picker
> offers — 25 unique kinds, re-counted 2026-09-10. It read **18** until then,
> which understated the product by seven and had been wrong for some time. A
> listing number that nobody can re-derive is a listing number that drifts; count
> it again before submitting.

**Categories:** Productivity; Data visualization
**Search keywords:** waterfall chart, bridge chart, Mekko, Marimekko, Gantt,
consulting charts, CAGR, editable charts, data visualization
**Support URL:** https://github.com/dannbleeker/SSF-Charts
**Privacy URL:** https://ssf-chart.struktureretsundfornuft.dk/privacy.html
**Terms URL:** https://ssf-chart.struktureretsundfornuft.dk/terms.html

## Assets to produce
- ✅ **Store logo** 300×300 PNG — `assets/store-logo-300.png`, generated by
  `npm run icons` from the same geometry and palette as the ribbon icons, so the
  listing image cannot drift from the mark it is meant to show.
- 🟡 **1–5 screenshots** 1366×768 of the pane + an inserted chart — **two slide
  images exist as of 2026-09-23 (`docs/store-shots/`); the pane image does
  not**, and the reason below was half wrong twice, corrected 2026-09-13 and
  again 2026-09-23.

  What is true: the agent drives PowerPoint through a headless browser, and the
  slide-editing canvas does not composite there. Every browser capture shows the
  ribbon and the pane correctly and a uniform `#F5F5F5` where the slide should
  be.

  **What was wrong: "a screenshot showing an inserted chart is not [capturable]".
  It is.** `slideImageBase64` (`src/render/powerpoint.ts`) wraps
  `Slide.getImageAsBase64` — PowerPointApi 1.8, and its own docstring calls it
  "the only way an add-in can see its own output". It renders through Office.js
  rather than through the screen, so the compositing problem does not reach it,
  and it takes a caller-chosen width. Every round already produces these: the
  archive strips them only because they are 48% of the bytes
  (`rounds/README.md`), and `PUBLISHING.md` says plainly that a round file
  carries "the host's own PNG of each slide the round added, so there is no deck
  to save and no screenshot to take."

  **THE SLIDE HALF IS DONE, 2026-09-23.** `scripts/store-shots.mjs` produced
  `docs/store-shots/01-waterfall.png` and `02-stacked.png` — both 1366×768, both
  rendered by PowerPoint itself through `getImageAsBase64` on the live web host,
  nothing composited and nothing edited. Each was looked at before being kept,
  which is the only reason they are usable: see "what the images caught" below.

  **AND THE SENTENCE BELOW WAS WRONG ABOUT THE PANE, corrected 2026-09-23.**
  "The pane is a browser screenshot, real, composites fine, it is just HTML" is
  true of the pixels and wrong about the product. The add-in iframe does not
  composite into a screenshot of the host page — an element capture of it comes
  back 318×1298 of pure white — so the only headless way to photograph the pane
  is to open `taskpane.html` directly, OUTSIDE PowerPoint. Do that and it
  renders perfectly and shows a banner reading **"Not running inside PowerPoint
  — use Download SVG, or sideload the manifest to insert native shapes."** That
  is the pane being honest, and it is a state no user in PowerPoint ever sees.
  `docs/evidence/pane-outside-powerpoint-2026-09-23.png` is that capture, kept
  as evidence and deliberately NOT placed in `docs/store-shots/` so that nobody
  uploads it by mistake. Cropping the banner away to make it read as a normal
  session is the same misrepresentation question as route 3 and is answered the
  same way: it is the owner's, not the agent's. **A pane image for the listing
  therefore still needs route 1.**

  So the slide-with-chart is capturable, the pane is capturable only out of
  host, and **what is not available headlessly is one AUTHENTIC frame containing
  both.** Three ways to close it, and the choice is the owner's:

  1. **A real PowerPoint window on AITEST**, which is where the 2026-09-13
     Windows reading was taken. This yields a genuine screenshot of ribbon, pane
     and chart together. It needs the owner to connect a viewer once — the same
     one-person step `docs/evidence/README.md` records for the desktop test.
  2. **Separate images**: the pane in one, the chart output in another. Both are
     real captures and many listings are shaped this way.
  3. **Compositing the host's slide PNG into the blank region of a browser
     capture.** Every pixel would be genuine product output — and it would still
     not be a screenshot, because no user ever saw that frame. Presenting it as
     one in a store listing is a misrepresentation question rather than a
     technical one, which is why it is listed and not done.

  **WHAT THE IMAGES CAUGHT, and why every capture gets looked at.** Three
  defects reached a written PNG before anyone saw one, and each passed every
  automated check the script had — shapes present, transfer length matched, PNG
  signature and dimensions right:

  - **A stacked column chart captioned "Revenue bridge FY25."** The script set a
    title and never chose a chart KIND, so it drew the pane's default. A bridge
    is a waterfall. A listing image whose caption contradicts its picture is the
    exact shape of an AppSource "functionality does not match the offer
    description". Choosing the kind has to come BEFORE typing the title, because
    the gallery tile applies a fresh sample config and overwrites it.
  - **A chart with its title drawn twice, at two sizes, axis labels out of
    frame.** `clearScript` is raced against a budget and returns `"clear-failed"`
    when it loses; the caller discarded that, so the chart drew on top of the
    previous one's leftovers. The clear must now say it worked AND the slide
    must read back empty before anything is drawn.
  - **An insert that never grouped.** The guard against this was removed on an
    argument that sounded good — the host traces "the host refused addGroup", a
    refusal costs re-editing rather than appearance, and a listing image is only
    an appearance — and the very next capture was the two-titled one above.
    Ungrouped did not mean "drew fine, failed to group". The guard is back, on
    evidence, against reasoning that had none.

  The lesson is the one `elements-probe.mjs` already carries: a capture tool's
  bug is indistinguishable from the failure it exists to report. Here the tool
  also spent two full runs reporting "the pane never settled" while the pane had
  been answering `"idle:Done."` throughout — it joined stdout to stderr, so the
  CLI's update banner landed after the value and the parser read null forever.
  **Look at the image. Nothing else in the pipeline can.**
- Optional short demo video.

## Submission checklist
- [x] Partner Center account created (free for Office Store apps). **Confirmed
      by the owner 2026-09-13.**
- [x] `manifest-prod.xml` validated: `npx office-addin-manifest validate manifest-prod.xml`
      — **all four manifests pass** as of 2026-08-27 (both PowerPoint, both Excel).
- [ ] Add-in works on **every** platform the manifest claims (web + Windows +
      Mac — testers check all of them). Do Phase 2 validation first.
      **Web: 425+ archived rounds. Windows: one reading, 2026-09-13,
      `docs/evidence/windows-desktop-2026-09-13.json`. Mac: the HOST surface is
      still unmeasured and the owner has no access to one** — what exists is
      `docs/evidence/mac-webkit-2026-09-13.json`, the chart engine on WebKit
      (Mac's engine family), 148/148 identical to Node. That removes the
      parse-and-format class of failure and nothing about Office.js.
      The 1.10 floor narrows what is being claimed: every certified host now
      runs the same code path, rather than Mac alone running the picture
      fallback.
- [x] Privacy + Terms pages live (they build to `/privacy.html`, `/terms.html`).
      **Re-verified 2026-09-16: both return 200, as do all 18 URLs across BOTH
      prod manifests.** It said 17 and `manifest-prod.xml`, which undercounted
      by one and named one file: the count predates the Excel companion's
      `src/excel/excel.html`, and a listing that ships two manifests has to
      check both. Re-run it with:
      `grep -ohE 'https://[^"<> ]+' manifest-prod.xml manifest-excel-prod.xml | sed 's/&amp;/\&/g' | sort -u`
- [ ] Listing copy above is trademark-clean; screenshots contain no competitor marks.
- [ ] Value is demonstrable **without a login** (SSF Charts needs none — good).
- [ ] **Notes for certification** written (Partner Center → the offer →
      **Properties → Notes for certification**). **This item was missing from the
      checklist until 2026-09-20**, which is worse than it sounds: it is the one
      field where you tell the reviewer how to exercise the add-in, and a blank
      one leaves a stranger to decide for themselves what the product was meant
      to do — with a pane of 25 chart kinds, a picture fallback and a ribbon
      entry that is easy to miss. Draft:

      > SSF Charts needs no account, licence key, sign-in or demo credentials.
      > Every feature works the moment the pane loads, and no document data
      > leaves the client — charts are built in the task pane and written to the
      > slide through Office.js.
      >
      > A one-minute pass, on PowerPoint on the web or the desktop app:
      >
      > 1. **Home** tab → **SSF Charts** → **Insert chart**. The pane opens on
      >    the Chart tab with a stacked column chart and sample data already
      >    filled in — nothing to type.
      > 2. Press **Insert into slide**. The chart lands as a group of native
      >    PowerPoint shapes (it is named `PowerChart` in PowerPoint's Selection
      >    pane). Click a column and restyle it with PowerPoint's own tools.
      > 3. Click that group on the slide. The pane shows *"An SSF chart is
      >    selected on the slide"* with an **Edit it** button: press it, change a
      >    number in the data grid, and insert again — the chart updates in place
      >    rather than adding a second one.
      > 4. Section **1 · Chart type** holds all 25 kinds. Two of them — **Tile
      >    map** and **Area** — are denser than PowerPoint on the web will draw
      >    shape by shape, so there they insert as a picture instead and the pane
      >    says so in words. They stay re-editable via step 3, and they insert as
      >    native shapes in the desktop apps. That is the behaviour the listing
      >    copy describes.
      >
      > Minimum requirement set: PowerPointApi 1.10 — PowerPoint on the web,
      > Microsoft 365 on Windows (2601 / 19610.20002 or later) and on Mac
      > (16.105 or later).

      Every claim in that draft is sourced rather than remembered: the ribbon
      path, the 1.10 floor and the build numbers are `manifest.xml`; the pane's
      opening chart is `sampleConfig("stacked")` and the banner wording is
      `taskpane.html` (`selection-banner`); the group name is `GROUP_NAME`, seen
      as one `msoGroup` of 40 shapes in
      `docs/evidence/windows-desktop-2026-09-13.json`; the picture message is
      `chartPicture` in `app.ts`; and "Tile map and Area" is the measured
      2-of-25 above, not a guess. **Walk the four steps once on the web before
      pasting it.** Nobody has driven them end to end as a script, and a reviewer
      who finds one click missing trusts the rest of the note less.
- [ ] Submit → respond to Microsoft validation feedback (days–weeks).

## Faster alternative (recommended first)
For BESTSELLER-internal use you don't need the store at all: a Microsoft 365
admin uploads `manifest-prod.xml` under **Admin Center → Settings → Integrated
apps → Upload custom app** and deploys it to chosen users/groups. No review,
appears automatically.
