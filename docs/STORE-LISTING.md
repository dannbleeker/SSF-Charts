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
> PowerPoint shapes** you can recolour, move and restyle by hand (on PowerPoint
> on the web and Microsoft 365 builds from 2026; older builds get a complete
> chart with a note about what they cannot draw). Add the
> annotations that tell the story: CAGR arrows, difference arrows, value lines,
> automatic column totals, and collision-avoiding labels. Every chart stays
> re-editable: reopen the pane, change the data, and it updates in place.

**Long description — feature bullets:**
- 25 chart kinds incl. waterfall bridges, Mekko/Marimekko, stacked/clustered/100%,
  cascade, funnel, butterfly, Gantt, combo, scatter/bubble, boxplot, violin,
  candlestick, radar, heatmap, tile map, treemap, sunburst and waffle.
- Native PowerPoint shapes on the web and current Microsoft 365 — recolour,
  move, or restyle any element by hand.
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
- ⛔ **1–5 screenshots** 1366×768 of the pane + an inserted chart — **blocked,
  and it needs a person.** The agent drives PowerPoint through a headless
  browser, and the slide-editing canvas does not composite there: every capture
  shows the ribbon and the pane correctly and a uniform `#F5F5F5` where the
  slide should be. The pane alone is capturable (and is, in the round archive);
  a screenshot showing an inserted chart is not. Take these by hand from a real
  PowerPoint window, or the listing will show a blank slide.
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
      **Verified 2026-09-13: both return 200, as do all 17 URLs in
      `manifest-prod.xml`.**
- [ ] Listing copy above is trademark-clean; screenshots contain no competitor marks.
- [ ] Value is demonstrable **without a login** (SSF Charts needs none — good).
- [ ] Submit → respond to Microsoft validation feedback (days–weeks).

## Faster alternative (recommended first)
For BESTSELLER-internal use you don't need the store at all: a Microsoft 365
admin uploads `manifest-prod.xml` under **Admin Center → Settings → Integrated
apps → Upload custom app** and deploys it to chosen users/groups. No review,
appears automatically.
