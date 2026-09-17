#!/usr/bin/env node
/**
 * Watch the office-js tracker for defects that land on code this add-in ships.
 *
 * Five of the bugs this project now guards against — a text box deleting the
 * selected shape, a sync that never returns after an image insert, `Shape.group`
 * throwing, layouts unreadable on custom-template decks — were found in a SINGLE
 * manual sweep of the tracker, on one afternoon, after months of not looking.
 * That is not a process, it is luck with a good afternoon attached. Every one of
 * them sat under code that ships, open upstream, for a year or more.
 *
 * So: look every week, automatically, and say only what is NEW.
 *
 * **`KNOWN_ISSUES` is the load-bearing half.** Without it the sweep reports the
 * same twenty issues forever and is ignored within a month. With it, the report
 * is exactly "here is something nobody here has looked at", and the table
 * doubles as the record of what this project did about each one — which is the
 * question a reader actually has when they meet an issue number in a comment.
 *
 * Usage:
 *   node scripts/office-js-watch.mjs                 # fetch and report
 *   node scripts/office-js-watch.mjs --from a.json   # report from a saved page
 *   node scripts/office-js-watch.mjs --json          # machine-readable output
 *
 * The fetch half cannot run in every environment — an agent session is bound to
 * its own repositories and the office-js API is out of reach from one. The
 * matching half is pure and is what the tests drive, through `--from`.
 */
import { isMain } from "./is-main.mjs";

/**
 * The Office.js surface this add-in actually calls.
 *
 * The bar is the same as the host probe's: not "this is interesting about
 * Office.js" but "if this broke, code in this repo would be wrong". A term that
 * matches nothing we call produces noise, and noise is how a weekly report stops
 * being read.
 */
export const WATCHED_APIS = [
  { term: "addGeometricShape", why: "every native shape a chart is made of" },
  { term: "addTextBox", why: "every label, title and axis tick" },
  { term: "addLine", why: "connectors, axes, difference arrows" },
  { term: "addGroup", why: "a chart IS a group; ungrouped charts are the failure mode" },
  { term: "shape.group", why: "how the repair pass counts a chart's children" },
  { term: "setSelectedShapes", why: "the selection round trip behind Edit it" },
  { term: "setSelectedSlides", why: "showSlide, and every off-screen redraw" },
  { term: "getSelectedShapes", why: "reading the chart the user clicked" },
  { term: "getImageAsBase64", why: "the only way the add-in can see its own output" },
  { term: "insertSlidesFromBase64", why: "the fast path for a whole generated deck" },
  { term: "slideMasters", why: "blankLayoutId — every slide the add-in creates" },
  { term: "SlideLayout", why: "same, and reported broken on custom templates" },
  { term: "getItemOrNullObject", why: "how every slide and shape is resolved" },
  { term: "shape.tags", why: "config tags — what makes a chart re-editable" },
  { term: "pageSetup", why: "the slide size every placement depends on" },
  { term: "fill.setImage", why: "the degrade-to-picture path" },
  {
    term: "context.sync",
    why: "the one call every failure this project has met goes through",
    // TRUE of this repo, and useless on its own. Both at once.
    //
    // Every Office.js issue in every host goes through `context.sync`, so as a
    // standalone signal it matches the TRACKER rather than our exposure. The
    // 2026-08-10 sweep is what that looks like: 63 issues to triage, 56 of them
    // found by this term alone, 37 naming Word, Excel or Outlook in the title,
    // and 2 naming PowerPoint. A weekly list of 63 to find 2 is a list nobody
    // opens — and `KNOWN_ISSUES` wants an entry for every one of them or they
    // all come back next Monday.
    //
    // So it corroborates rather than nominates: it counts when something else
    // already ties the issue to us, and never by itself.
    needsCorroboration: true,
  },
];

/**
 * The office-js issue template's Host line, and specifically its ANSWER.
 *
 * `* Host [Excel, Word, PowerPoint, etc.]: *Excel*`
 *
 * The menu names PowerPoint in EVERY templated issue, which is a trap worth
 * spelling out because the first version of this walked straight into it: a
 * plain "does the body say powerpoint" rescue kept `Excel Data Validation -
 * Whole Numbers have restricted integer values`, on the strength of the square
 * brackets. Only what the reporter typed after the colon counts.
 */
const HOST_ANSWER = /host\s*\[[^\]]*\]\s*:?\s*\**\s*([^\n*]{0,40})/i;

/**
 * Does this issue say it is about PowerPoint?
 *
 * The second signal a corroboration-only term looks for, and it never reads the
 * raw body — see `HOST_ANSWER`. A stated host is believed outright, in both
 * directions: "Excel" is a no whatever else the text mentions. With no host
 * line to go on it falls back to the TITLE, which is the one part of an issue
 * nobody fills in from a menu.
 *
 * Only ever used to RESCUE an issue a broad term found, never to reject one a
 * specific term did — so being wrong here costs a line in a weekly list, not a
 * missed defect.
 */
function mentionsPowerPoint(title, body) {
  const stated = HOST_ANSWER.exec(body ?? "")?.[1];
  if (stated && /\w/.test(stated)) return /powerpoint|\bppt\b/i.test(stated);
  return /powerpoint|\bslides?\b|\bpptx?\b/i.test(title ?? "");
}

/**
 * Issues this project has already read and responded to.
 *
 * The value is not a status, it is WHAT WE DID — so an issue number met in a
 * code comment can be traced to a decision without re-reading the thread. "No
 * exposure" is a legitimate entry and an important one: it records that someone
 * checked, which is otherwise indistinguishable from nobody having looked.
 */
export const KNOWN_ISSUES = {
  1650: "A slide add whose sync never resolves though the slide lands. Every slide-add in powerpoint.ts is bounded and verified by re-reading the deck rather than by trusting the promise.",
  2328: "SlideMaster.shapes throws GeneralException on the web. Asked by the `layouts-readable` probe; blankLayoutId already degrades to the inherited layout.",
  2699: "A blank slide that is only blank to the eye. Handled by the demo path's blank re-read.",
  2775: "Adding a text box deletes the SELECTED shape, web only. Guarded by `dropShapeSelection` on the insert path, and asked by the self-test's `a selected shape survives an insert`.",
  2172:
    "addGeometricShape refused on a slide that is completely blank, web only; the reporter's workaround was to add a text element first. " +
    "Closed `Status: fixed`. NO EXPOSURE, and worth saying why rather than leaving it to the label: every slide this repo draws onto gets " +
    "its shapes through `addAndRenderItem`, which stamps a title text box before the chart's shapes on the demo path and otherwise draws " +
    "onto a slide the user already has. Found on 2026-08-07 while searching for the ShapeCollection.getItem refusal; it is a different " +
    "failure (the ADD, not the lookup) and it is not this one.",
  2780: "Carried as a caveat in the docs; no code depends on the behaviour.",
  2881: "Complex SVG renders wrong through the picture path. Why charts are native shapes rather than an image by default.",
  2903:
    "A stale shape proxy answers InvalidParam passed to GetItem(id). The reason `refreshShapes` and `settleAndTagChart` exist. " +
    "Re-read 2026-08-07 and it says more than we had taken from it: the report is about a freshly ADDED slide, not a stale shape — on Online " +
    "a slide that has been added and synced is not usable yet (text does not render, images land on the first slide instead), and the " +
    "reporter's only workaround is to wait a couple of seconds. Microsoft closed it `not planned`, so the wait is the fix that exists. " +
    "TRIED AND REVERTED: `addScratchSlide` settled for 2s on 2026-08-07 and the next real-host round answered 1 of 25 questions against 19 of 26 before it — the add landed, the wait ran, and the liveness check after it found nothing. This host resolves a fresh slide id ONCE and refuses it ever after, so waiting spends that one resolution later instead of buying time. Do not re-apply the workaround from this issue.",
  3014:
    "PowerPoint's API has no grouping story: creating and reading groups is a known parity gap, grouped shapes come back from getItem() " +
    "as type `unknown`, and sub-shapes cannot be reached. CLOSED/COMPLETED — checked 2026-09-05, and the table said `Open since 2022, " +
    "Status: in backlog` until then. DO NOT STRIP THE WORKAROUND ON THE LABEL: a fix upstream is a claim about the service, not proof " +
    "it reached this host. What this host actually does is already measured — the rounds gate reports 12 of 12 attempts grouped, 0 " +
    "refused, over 358 rounds — so the parity gap does look shut here. Removing `ungroupedFallback` still needs a probe that FAILS " +
    "without it. It is why " +
    "`ungroupedFallback` and CHART_PARTS_TAG exist at all — a chart that cannot be grouped has to carry its parts some other way — and " +
    "why `chooseGroupMembers` prefers an ungrouped, tagged chart over a failed addGroup that takes the tagging down with it.",
  3083: "setSelectedShapes([]) does not clear the selection on the web. clearShapeSelection re-selects the slide instead.",
  3269: "Office.js cannot read speaker notes at all. Recorded as a limitation; nothing here reads them.",
  3309: "SVG cannot be read back out of a shape. Same reason as #2881.",
  3698: "A picture cannot be inserted while a shape is selected, and setSelectedShapes([]) may never resolve. Both covered by dropShapeSelection and by the scenario ordering.",
  3826: "A freshly-added slide's layout shapes throw GeneralException. Asked by the `slide-layout-readable` probe.",
  4272: "A load of more than ~50 items answers short. Why the deck scan is paged at READBACK_PAGE.",
  4906: "SlideLayout.shapes throws on decks built from a custom template — the owner's case. Asked by the `layouts-readable` probe.",
  5455:
    "PowerPoint throws GeneralException READING ParagraphFormat.horizontalAlignment on a text range with no alignment set " +
    "(no `align` on `a:pPr`). Closed. NO EXPOSURE, and the direction is the whole reason: this repo only ever WRITES that " +
    "property — `stampSlide`'s banner and the text renderer, both plain assignments — and never loads it. Checked on " +
    "2026-08-10 rather than assumed, because the property name matching ours is exactly what makes an issue look like a hit.",
  6079:
    "PowerPoint on the WEB uppercases tag keys internally, then requires the uppercased spelling to read them back; a lowercase " +
    "`tags.getItem` throws GeneralException. Desktop is case-insensitive, so an add-in written against desktop breaks online. " +
    "NO EXPOSURE, checked rather than assumed on 2026-08-11: every tag key in this repo is a constant and every one is already " +
    "upper case — POWERCHART_CONFIG, _PARTS, _ORIGIN, _SCENE, _DEMO_SLOT, and the probe's POWERCHART_PROBE pair. Worth keeping in " +
    "the table because the trap is invisible: it fires only on the web, only for a key somebody spells in lower case, and the " +
    "symptom is a GeneralException nowhere near the write.",
  2474:
    "`SlideRange.id` is not roundtrippable: the id read off a selection lacks the `#XYZ` suffix the same slide carries when " +
    "read from the collection, so `slides.getItem(id)` answers InvalidArgument while `slides.getItemAt(index)` works. Closed " +
    "`not planned`. INDEPENDENT CONFIRMATION of this repo's own finding — the 2026-08-11 rounds hold scratch ids like " +
    "`4123571114#123571113` while the deck lists `256#109857222`, and delete-by-id is therefore structurally impossible rather " +
    "than merely unreliable. Different numbers, same disease: two id spaces for one slide. The workaround the issue gives is " +
    "the one already built — go by POSITION (`deleteTrailingSlides` / `positionalSweepPlan`), which reclaimed 68 of 68 and 72 " +
    "of 72 scratch slides on its outings.",
  3565:
    "context.sync() taking progressively longer on every run, with a restart of the app resetting it — the reporter's guess is " +
    "a memory leak. WORD FOR MAC, not PowerPoint web, so this is NOT evidence about our host and must not be cited as though " +
    "it were. Recorded because the SHAPE matches what `what makes a long run slow down` measured here, and because it names a " +
    "cheap experiment nobody has run: if reloading the tab resets the per-slide cost curve, the accumulation is in the session " +
    "rather than in the deck. Closed `no recent activity`.",
  6329:
    "PowerPoint ON THE WEB forces a FULL PRESENTATION SAVE on every context.sync(), read-only syncs included; the independent " +
    "duplicate #6740 measured autosave firing once per iteration of a sync loop. This is our host and our exact call, so it is " +
    "live exposure rather than a recorded curiosity — the add-in draws in batches of SHAPES_PER_SYNC, which makes the number " +
    "of forced saves proportional to shape count over ten, and the deck scan adds two more per page while writing nothing. " +
    "STATUS IS CONTESTED AND THE LABEL SHOULD NOT BE TRUSTED: Microsoft marked it fixed on 2026-08-10 and the reporter " +
    "rebutted it with video thirty minutes later; it carries both `Status: fixed` and `Needs: author feedback`. Cited by " +
    "`syncsSoFar`, which exists to measure our own sync count per round — the archive indexes every crash figure by SHAPES, " +
    "and if this issue is right the load that matters was never counted.",
  6867:
    "Slide.exportAsBase64 omits modern comments and ppt/authors.xml from the exported deck. NO EXPOSURE: the add-in calls it " +
    "through `slideImageBase64` to get a PICTURE of a slide for the round's deck evidence, and a round has no comments in it " +
    "and no interest in them if it did. Kept in the table rather than dropped so the next sweep does not re-raise it.",
  3784:
    "Shape TAGS are lost when a shape is cut and pasted on PowerPoint WEB; desktop keeps them. Open, awaiting Microsoft, no " +
    "workaround offered. Triaged 2026-08-12 and it is the most direct user-facing exposure in this table: every chart's " +
    "config lives in a `POWERCHART_CONFIG` shape tag, so on the web a user who cuts a chart and pastes it back has a chart " +
    "that is no longer re-editable — silently, with nothing in any log, because the tag ANSWERED and said there is no " +
    "config. `tagsUnread` does not cover it; that counts tags the host would not answer either way. " +
    "NOT DETECTED ON PURPOSE. A count of shapes named `PowerChart` carrying no config would find these, and on this host it " +
    "would be dominated by the tag writes the host itself refuses every round (`tagging failed` appears in most rounds), so " +
    "the number would say 'this host is unwell' rather than 'your paste broke a chart'. `same scale across the deck` already " +
    "reports how many charts in a deck are re-editable, which is the same fact without the false precision. Recorded as a " +
    "limitation so the next person to meet an un-editable pasted chart does not go looking for it in our tag writer.",
  6266:
    "getImageAsBase64 behaves differently on Mac and Windows: content add-ins do not appear in the image Windows generates. " +
    "NO EXPOSURE — nothing here embeds a content add-in in a slide it rasterises — and it is registered because it was READ AND " +
    "REJECTED, not because it bites. Found 2026-08-21 while searching for the fault behind `rasterGap`: the host answers the " +
    "first rasterise of a slide and then hangs for the full 20s budget on a second rasterise of the SAME slide issued straight " +
    "after it, which is 40 of 40 blind visibility gates in the archive. THIS ISSUE IS NOT THAT. It is about the CONTENT of a " +
    "render that succeeds, not a render that never returns, and the nearest thing upstream to the real fault remains #5022. " +
    "Recorded so the next search does not have to re-read it to reach the same conclusion.",
  6498:
    "Shapes inserted on PowerPoint WEB do not reflect instantly: changes may not appear without a page refresh, and — the part" +
    "that matters here — they may appear IN THE SLIDE PREVIEW BUT NOT IN THE MAIN VIEW. Open, no Microsoft response, no " +
    "workaround offered. Triaged 2026-08-12 as a LIMIT ON WHAT ONE GATE PROVES rather than a bug to fix: `chartIsVisible` is " +
    "this project's only mechanical evidence that a chart it drew can be SEEN, and it works by rasterising the slide with " +
    "`getImageAsBase64` and diffing the bytes — which renders the preview, precisely the surface this issue says can disagree " +
    'with the canvas. So a pass there means "the chart is in PowerPoint\'s own render of the slide", which is strictly weaker ' +
    'than "a human looking at the slide would see it", and the verdict says so now instead of claiming the stronger thing. ' +
    "There is no read of the main view available to an add-in, so this cannot be closed by measuring harder — only by a human " +
    "looking, which is why the battery leaves its slides in the deck. The reporter's case is slide MASTERS and ours is slides, " +
    "so the match is on the mechanism rather than on the repro.",
  5022: "context.sync() runs indefinitely when shapes are re-read after an image insert. Asked by the `picture-then-shape-read` probe; drawDemoItem does exactly that sequence.",
  5101: "A placeholder keeps type: Placeholder when reused. NO EXPOSURE — this repo never reads a shape's type. Checked, not assumed.",
  5264: "A part of the object model Office.js cannot reach. Recorded as a limitation.",
  5849:
    "Shape.group throws GeneralException. Asked by the `group-of-existing-shape-readable` probe; countGroupChildrenPage reads groups exactly that way. " +
    "Re-read 2026-08-08: CLOSED as `Status: no recent activity`, not fixed, and the report is DESKTOP (Windows, 2505 build) — nobody has established " +
    "the web behaviour. That gap is expensive here: `contentShapes` returns UNKNOWN_CONTENT for every grouped slide, which is what makes the reconcile " +
    "report a slide complete without counting it (see SlotVerdict.measured). `ShapeGroup.shapes` with `getCount()` does exist in the API, so if the web " +
    "host honours it the verdicts could be measured instead of assumed. Both probe questions that would settle it — `group-reports-its-children` and " +
    "`group-of-existing-shape-readable` — have failed at SETUP in every round so far (`no-scratch-slide`), so this is unknown rather than answered.",
  5896: "Reported alongside another SVG defect; same handling.",
  6363:
    'PowerPoint.run\'s batching "fails to load properties reliably — properties not available after context.sync()", web only. ' +
    'Re-read 2026-08-07 and it is the closest published match to this repo\'s central failure: the reporter loads `slides.load("items")`, ' +
    'syncs, and reading `.items` throws "the property is not available", and the same happens for `shapes.items`, `shape.type` and ' +
    "`textFrame.hasText`. Identical code typed into the browser console works; only the batched form fails. `Status: under investigation`, " +
    "assigned, no root cause published, and the thread records TEN workarounds that all failed — including re-querying with `getItemAt()`, " +
    "`items/type` path loading, per-load syncs, and `context.trackedObjects.add()`. " +
    "Our shape of it is adjacent but not identical: `the re-read before grouping came back empty` is an EMPTY array, not a throw, and the " +
    "empty/threw split is traced precisely so the next round says which. No fix available; the recovery paths are the exposure management.",
  2714: "setSelectedDataAsync converts points to pixels. NO EXPOSURE — this repo never calls it. Checked, not assumed.",
  6658:
    "There is no supported way to declare PLATFORM scope in an add-in manifest, and the reporter calls it a Marketplace certification " +
    "blocker. Open, `Type: product question`, last activity 2026-06-04 — Microsoft has not treated it as a defect, so do not expect a " +
    'mechanism to arrive. NOT A BUG WE HIT AT RUNTIME; it is why this repo\'s manifest floor is what it is. `<Host Name="Presentation" />` ' +
    "means web + Windows + Mac + iPad as ONE unit and cannot be narrowed, and the issue names the consequence: Partner Center AUTO-DERIVES " +
    'the certified platform list from the `<Requirements>` block. With certification policy 1120.3 ("Add-ins must work across all platforms ' +
    'that support methods defined in the Requirements element"), that inverts the usual instinct — a LOWER MinVersion does not hedge, it ' +
    "enlarges the surface Microsoft tests. It is the argument behind raising PowerPointApi 1.4 -> 1.8 -> 1.10 on 2026-09-13; see the comment " +
    "in manifest.xml. The reporter's own attempt to narrow the block instead produced `No Supported Office Products`, so that escape does " +
    "not work either.",
  6183:
    "PowerPoint ONLINE does not update `left` and `top` on a shape simultaneously — set both in one sync and only the left " +
    "property is applied. Closed COMPLETED. THE MOST EXPENSIVE ENTRY IN THIS SWEEP, because it named the axis our own check " +
    "omitted: `an update follows a moved chart` moved a probe chart 60pt across and 40pt down and then compared x ALONE, " +
    "from the day it was written — the move check and the drift check after the redraw, both. A host doing exactly what " +
    "this issue describes would have been reported as 19 of 19 for as long as it did it. `moveShapeBy` is the call, and it " +
    "sets left and top in one sync. Fixed 2026-09-17, then fixed AGAIN the same day after a review found the first fix " +
    "inert: adding the y clause to the existing `skipped: true` return bought nothing, because `scenarioBlame` answers " +
    "`not-run` for anything skipped and `describeSelfTest` drops it out of `ran` — the headline came back `17 of 17 " +
    "scenarios passed · 2 skipped (host cannot run them)` with `selfTestNeedsAttention` FALSE, which the pane paints " +
    "green. The two cases are now split on the evidence that separates them: NEITHER axis moved is a skip (nothing was " +
    "measured), ONE axis moved is a failure (the host answered, and the answer was wrong). Guarded from both ends — " +
    "`faults.ignoresTopWrites` drives the partial move and asserts it is NOT skipped; `faults.ignoresLeftWrites` beside " +
    "it drives the declined move and asserts it IS; and `judgeOriginHeld` is exported so the redraw half can be tested " +
    "directly, which no fault can reach because the fake builds a shape's position from its creation box.",
  6948:
    "PowerPoint `customXmlParts.getByNamespace(...).getItemOrNullObject` THROWS for an id that is not there, instead of " +
    "returning a null object the way Excel and Word do. Open, `Needs: triage`, desktop. CORROBORATION FOR THIS REPO'S OWN " +
    "FINDING, which is why it earns an entry: the deck-style read calls `getByNamespace(...).getOnlyItemOrNullObject()`, " +
    "and rounds 096 and 097 recorded that on an EMPTY collection that call never returns at all — `getCount()` answers " +
    "`parts: 0` and the item read hangs for the whole budget. Same API family, same broken OrNullObject promise, two " +
    "symptoms and two hosts. `attemptDeckStyleRead` already counts first and never asks for the only item of an empty " +
    "collection, so nothing is owed; what this adds is that the contract is unreliable upstream generally, and the guard " +
    "must not be removed on the grounds that OrNullObject is documented not to throw.",
  4121:
    "`addLine` comes out bent when the box handed to it has a zero height or width, and the workaround Microsoft gave is to " +
    "RE-ASSIGN the dimension after creation (`const line = shapes.addLine(..., {height: 0, ...}); line.height = 0;`). Closed " +
    "for inactivity, not fixed. NOT EXPOSED, AND AN EARLIER VERSION OF THIS ENTRY SAID IT WAS — corrected 2026-09-17 by a " +
    "review that read the code the entry was describing. `addSegment` clamps before it calls anything: the box is " +
    "`{ width: Math.max(w, 0.5), height: Math.max(h, 0.5) }` (`powerpoint.ts:11794`, unchanged since 78afcc9 on " +
    "2026-07-06), seven lines above the branch, and `powerpoint.ts:11839` is the ONLY `shapes.addLine` call site in `src/`. " +
    "So no zero dimension is ever handed to `addLine` and the issue's trigger cannot occur. The struck claim also " +
    'misquoted the branch: it is `w < 0.5 || h < 0.5 || dashStyle !== "none" || !canRotate()`, not the first pair alone. ' +
    "What the entry got right and keeps: a horizontal rule really does reach `addLine` with a 0.5pt height, the workaround " +
    "if this ever bites is one re-assignment after creation, and the showcase deck is NOT evidence either way because it " +
    "is written by our own OOXML writer and never goes near `addLine`.",
  4222:
    "`getSelectedShapes()` sometimes answers an EMPTY array while a shape is selected on the slide, web and desktop. Closed " +
    "COMPLETED. Live exposure on the pane's newest path: `loadChartFromSelection` (`powerpoint.ts:3565`) reads the " +
    "selection the user made with a click, " +
    "and an empty answer there is indistinguishable from 'nothing is selected' — the user clicks a chart, presses Edit, and " +
    "is told to select a chart. Already reported rather than swallowed (`edit the chart the user selected` is routine and " +
    "says which of the two it saw), and no retry is added on the strength of a closed issue: the rule here is that a " +
    "workaround needs a probe that fails without it.",
  3552:
    "`setSelectedSlides` throws when the NOTES pane holds the selection — desktop only, web is fine. Closed COMPLETED. " +
    "Exposure is real and already contained: `showSlide` is the only caller, it is best-effort behind `supports('1.5')`, and " +
    "every path that uses it treats failure as 'the slide could not be brought into view' rather than as an error. It stays " +
    "in the table because it is DESKTOP-ONLY and this repo's whole evidence base is the web host — so if a desktop user ever " +
    "reports that the pane will not follow their chart, this is the first thing to read, and no round will have seen it.",
  3715:
    "`insertSlidesFromBase64` loses CENTRE ALIGNMENT on inserted text boxes. PowerPoint, web and desktop. Closed COMPLETED. " +
    "NO EXPOSURE ON THE PATH WE TAKE, and the distinction is one argument wide: the repro passes " +
    "`formatting: useDestinationTheme`, and both of our call sites use `KeepSourceFormatting` — the `insertSlidesFromPptx` " +
    "wrapper defaults to it (`powerpoint.ts:7626`) and `replaceSlideWithDeck` passes it explicitly (`powerpoint.ts:9575`). " +
    "Centred text is not incidental here — every chart title the fast path ships is centred — so if that default is ever " +
    "changed to follow the deck's theme, this issue is the cost of doing it. Checked rather than assumed on 2026-09-17.",
  6130:
    "PowerPoint TABLE cells will not report `cell.fill.color` or `cell.borders.*.color` — they read back undefined or null. " +
    "Closed COMPLETED. NO EXPOSURE: this add-in has no PowerPoint table in it anywhere. The `table` element is built from " +
    "native shapes — `rule-top`, `rule-header`, `cell-text-<r>-<c>` — and `addTable` appears zero times in `src/`. Checked " +
    "rather than assumed, because an issue about table cell colour is exactly what a table-drawing add-in should stop and " +
    "read.",
  6933:
    "Setting `TextRange.font.color` and `font.underline` does not visually override HYPERLINK theme formatting, though the " +
    "API reports the new values back. Open, `Area: PowerPoint`, desktop. NO EXPOSURE: nothing here adds or reads a " +
    "hyperlink — `hyperlinks`, `getLinkedTextRangeOrNullObject` and `font.underline` are all zero matches across `src/`. In " +
    "the table because the sweep will find it again every week until it is.",
  4988:
    "Excel table creation becomes slow when heavy Excel files are open. EXCEL, not PowerPoint. NO EXPOSURE. It reaches the " +
    "sweep on `getItemOrNullObject` plus `context.sync`, which is what a corroboration-only term looks like when it lands " +
    "on the wrong host — kept so it is not re-triaged next Monday.",
  5537:
    "`Document.insertFileFromBase64()` strips web extension settings. WORD ON MAC. NO EXPOSURE: this repo calls " +
    "`Presentation.insertSlidesFromBase64`, a different method on a different host, and stores nothing in document " +
    "settings. Same shape of false positive as #4988, recorded for the same reason.",
};

/** Lower-cased haystack for one issue. */
const haystack = (issue) => `${issue.title ?? ""}\n${issue.body ?? ""}`.toLowerCase();

/**
 * Which of these issues are worth a human's attention.
 *
 * Pure, and that is deliberate: the network half cannot run everywhere, and a
 * filter nobody can test is a filter nobody should trust. Pull requests are
 * dropped — the GitHub issues endpoint returns them too, and a PR against
 * office-js is not a defect report about it.
 */
export function freshIssues(issues, known = KNOWN_ISSUES, apis = WATCHED_APIS) {
  const out = [];
  for (const issue of issues ?? []) {
    if (issue.pull_request) continue;
    if (issue.number in known) continue;
    const text = haystack(issue);
    const hits = apis.filter((a) => text.includes(a.term.toLowerCase()));
    if (!hits.length) continue;
    // A hit that only ever corroborates cannot nominate an issue on its own —
    // see `needsCorroboration`. Something specific has to have matched too, or
    // the issue has to say it is about PowerPoint.
    if (hits.every((a) => a.needsCorroboration) && !mentionsPowerPoint(issue.title, issue.body)) continue;
    out.push({
      number: issue.number,
      title: issue.title,
      state: issue.state,
      url: issue.html_url ?? `https://github.com/OfficeDev/office-js/issues/${issue.number}`,
      updated: issue.updated_at,
      hits: hits.map((h) => h.term),
      why: hits.map((h) => h.why),
    });
  }
  // Newest activity first — a reader works down the list and stops when they
  // recognise everything below.
  return out.sort((a, b) => String(b.updated ?? "").localeCompare(String(a.updated ?? "")));
}

/** The issue body a run posts. Markdown, and short enough to read on a phone. */
export function reportBody(fresh, checked) {
  if (!fresh.length) {
    return (
      `No office-js issue mentioning an API this add-in calls has appeared since the last sweep.\n\n` +
      `Checked ${checked} recently-updated issue(s) against ${WATCHED_APIS.length} watched calls ` +
      `and ${Object.keys(KNOWN_ISSUES).length} already answered.\n`
    );
  }
  const rows = fresh
    .map(
      (f) =>
        `### [#${f.number}](${f.url}) — ${f.title}\n\n` +
        `\`${f.hits.join("`, `")}\` · ${f.state} · updated ${f.updated}\n\n` +
        f.why.map((w) => `- ${w}`).join("\n"),
    )
    .join("\n\n");
  return (
    `${fresh.length} office-js issue(s) mention an API this add-in calls and are not yet in ` +
    `\`KNOWN_ISSUES\`.\n\n` +
    `For each: decide whether this repo is exposed, then add it to \`KNOWN_ISSUES\` in ` +
    `\`scripts/office-js-watch.mjs\` **with what was done about it** — including "no exposure", ` +
    `which records that somebody checked. An issue left out of that table comes back next week.\n\n` +
    `${rows}\n`
  );
}

/** Recently-updated issues from the tracker. Only reachable where the API is. */
async function fetchIssues(pages = 3, perPage = 100) {
  const all = [];
  for (let page = 1; page <= pages; page++) {
    const url =
      `https://api.github.com/repos/OfficeDev/office-js/issues` +
      `?state=all&sort=updated&direction=desc&per_page=${perPage}&page=${page}`;
    const res = await fetch(url, {
      headers: {
        Accept: "application/vnd.github+json",
        ...(process.env.GITHUB_TOKEN ? { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` } : {}),
      },
    });
    if (!res.ok) throw new Error(`office-js issues page ${page}: HTTP ${res.status}`);
    const batch = await res.json();
    if (!Array.isArray(batch) || !batch.length) break;
    all.push(...batch);
  }
  return all;
}

async function main() {
  const args = process.argv.slice(2);
  const fromAt = args.indexOf("--from");
  const issues =
    fromAt >= 0
      ? JSON.parse(await (await import("fs/promises")).readFile(args[fromAt + 1], "utf8"))
      : await fetchIssues();
  const fresh = freshIssues(issues);
  if (args.includes("--json")) {
    process.stdout.write(JSON.stringify({ checked: issues.length, fresh }, null, 2));
    return;
  }
  process.stdout.write(reportBody(fresh, issues.length));
  // Non-zero when there is something to look at, so a workflow can branch on it
  // without parsing the body.
  if (fresh.length) process.exitCode = 3;
}

// Only when run directly — the tests import the pure half. The predicate is
// SHARED now (`is-main.mjs`), because this file is where the bug was found and
// fixed and its three siblings were left standing with it: the old form split
// argv[1] on "/" and asked whether `import.meta.url` ended with the last piece,
// which on Windows never splits at all, so the CLI silently did nothing on the
// owner's machine — `--from`, `--json`, the lot, exit 0 and no output.
if (isMain(import.meta.url, process.argv[1])) {
  await main();
}
