# Publishing runbook — get SSF Charts live in PowerPoint

Instructions for a Claude (Opus 4.8) session working with the repo owner to
take SSF Charts from "feature-complete on main" to "usable inside
PowerPoint, with the Claude skill active". Read `CLAUDE.md` first — the
working conventions there (lockstep rule, branch flow, auto-merge policy,
visual QA) apply to every change you make here.

Legend: **[agent]** = you can do it with repo access; **[owner]** = needs
the owner's click (GitHub settings, PowerPoint UI, claude.ai account).
Do the phases in order — later phases depend on the hosted URLs.

---

## Phase 0 — Preconditions ✅ done

1. **[owner] Make the repo public** — ✅ done.
2. **[agent] Pre-publication sweep** — ✅ done: no secrets/keys/tokens
   (`git grep` clean; every "token" hit is benign code), no env/credential
   files, all sample/showcase data is invented dummy data, `npm test` green.
3. **[owner] Post-public hygiene** — ✅ done (Dependabot + CodeQL, description
   + topics).
4. **[agent] Branch protection** — ✅ done: ruleset "main: require CI green"
   (active) requires the `test` check and blocks force-push and deletion of
   `main`. Only `test` is required because it is the only check that runs on
   `pull_request` — `build`/`deploy` (Pages) and `release` fire on push/tag
   only, so requiring them would deadlock every PR. Repo admins keep an
   `always` bypass, so the owner can never be locked out of his own default
   branch.

## Phase 1 — Host the add-in on GitHub Pages ✅ agent work landed

Office add-ins load from an HTTPS URL; the dev manifests point at
`https://localhost:3000`. The site is hosted on GitHub Pages under a **custom
domain**, `https://ssf-chart.struktureretsundfornuft.dk/`. Because a custom
domain serves the project site from its **root**, the bundle base is `/`
(no `/PowerChart/` path segment) — the prod-manifest URLs are just
`https://ssf-chart.struktureretsundfornuft.dk/…`.

1. **[agent] Build for Pages** — ✅ `npm run build:pages`
   (`scripts/pages-postbuild.mjs`): runs the prod-manifest gen, `tsc`, a
   root-base `vite build`, then copies the manifest-referenced ribbon icons
   into `dist/assets/`. Emits `index.html`, `src/taskpane/taskpane.html`,
   `src/excel/excel.html`, `assets/icon-*.png`, and the static `public/` files
   (`CNAME`, `privacy.html`, `terms.html`) which Vite copies verbatim.
   > Gotcha found & fixed: Vite doesn't bundle `assets/icon-*.png` (they're
   > referenced only by the manifests), so without the copy step the hosted
   > icon URLs 404. `pages-postbuild.mjs` copies them; the `CNAME` and legal
   > pages ride along from `public/`.
2. **[agent] Deploy workflow** — ✅ `.github/workflows/pages.yml`: on push to
   `main`, `npm ci` → `npm run build:pages` → `upload-pages-artifact` (path
   `dist`) → `deploy-pages`, with `pages: write` / `id-token: write`.
3. **[owner] Enable Pages + custom domain** — ✅ done (Source: GitHub Actions;
   domain `ssf-chart.struktureretsundfornuft.dk`). Confirm **Enforce HTTPS**
   is checked once the cert provisions.
4. **[agent] Production manifests** — ✅ `scripts/build-manifest.mjs` rewrites
   `https://localhost:3000` → the custom-domain origin into
   `manifest-prod.xml` / `manifest-excel-prod.xml` (committed; `--check` mode
   gates staleness in `ci.yml`; regenerated + attached to releases in
   `release.yml`). Both GUIDs (`b7f6d3a2…`, `c8a7e4b3…`) preserved; 0 localhost
   URLs survive. **Validated in CI** by the `manifest` job, which runs
   Microsoft's own `office-addin-manifest validate` over all four manifests. Its
   own job rather than part of `test`, because the validator calls a Microsoft
   service: `test` is the only check the branch ruleset requires, so an outage
   there must not be able to block every merge in the repo. The manifest is the
   one artifact CI could otherwise not judge — a broken one passes every test
   here and fails when you sideload it, which is the slowest feedback loop this
   project has.

   > **It found something on its first run.** All four manifests declared
   > `<Version>0.1.0</Version>`, and Microsoft's validator rejects that outright:
   > *"Manifest Version Too Low: The manifest has unsupported version number
   > less than 1.0."* An error, present since the day the files were written,
   > passing every test in this repo because nothing had ever looked. Bumped to
   > `1.0.0.0` — deliberately independent of the npm package version, which is
   > free to stay below 1.0. `test/manifest.test.ts` pins the rule offline so it
   > cannot come back when the validation service is unreachable.
5. **[agent/owner] Smoke-test the deployment**: after the first Pages run,
   `curl -sI https://ssf-chart.struktureretsundfornuft.dk/src/taskpane/taskpane.html`
   → 200, and the icons under `/assets/icon-*.png`. Load the demo gallery URL
   in a browser to confirm assets render.

## Phase 2 — Sideload in PowerPoint ([owner], agent assists)

Pick the platform(s); the manifest file is `manifest-prod.xml` from Phase 1
(attached to the latest release, or in the repo).

- **PowerPoint on the web** (fastest validation): open a deck on
  office.com → Home ▸ Add-ins → **More add-ins** → **My Add-ins → Upload My
  Add-in** → pick `manifest-prod.xml`.
- **Windows**: easiest supported route is the same Upload dialog (newer
  builds), else the shared-folder catalog: put the manifest in a folder,
  share it (`\\machine\manifests`), add it under File → Options → Trust
  Center → Trusted Add-in Catalogs, restart PowerPoint, Insert → My
  Add-ins → Shared Folder.
- **Mac**: copy the manifest to
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` and
  restart PowerPoint; it appears under Insert → My Add-ins.
- **Excel companion**: same procedure in Excel with
  `manifest-excel-prod.xml`.

### The standing test run

What to do on **every** build that lands a real-host fix. Four items, ordered
by risk. Two are one click each and run themselves; the other two are a scroll
and a single click-and-edit — about three minutes of attention in total.

**The probe asks every question three times, spread across the round.** That is
where the extra five to eight minutes went. A question asked ONCE has been
sampled, not answered — which is why `UNSTABLE_ANSWERS` had to be built by hand
across ten rounds — so a round now says on its own whether an answer held still,
and the pane's summary line names any question that **CHANGED ITS ANSWER
MID-ROUND**. That alone makes a round worth sending, whatever the diff says.

When the host starts refusing scratch slides — it does, in windows of about
fifteen seconds — the later passes drop to the shortlist of questions this
project does not yet trust, rather than bidding for slides against them. The
trace says which it did and why.

This used to be seven tests. Tests 1, 2, 3 and 7 were absorbed into the
self-test battery once `Slide.setSelectedShapes` turned out to have been
available since PowerPointApi 1.5, and the old test 4 (the `=SUM(A1:ZZ999)`
formula crash) was deleted outright: it is pure browser JavaScript, it never
needed PowerPoint, and it has been covered in CI all along by
`test/formula-blanks.test.ts`. The owner drives PowerPoint, the agent
fixes fallout — expect a real host to surface things the mocked tests can't,
because every Office.js assertion in this repo is against a fake.

**Before you start.** Wait ~2 minutes after the merge for the Pages deploy.
Open the pane and check the **build stamp** under the title is the commit you
mean to test — PowerPoint caches the pane aggressively, and a whole session can
otherwise go into testing code the host never fetched. Hard-reload if it is
older. Then tick **Verbose trace** in the Testing section and leave it on.

**Beside it sits "Picture every slide", now ON by default, and it decides what a
round can prove.** Left off, the round photographs the first twelve slides it
adds and past that records the slide and says it never asked. That was the
original default and the pictures are still the heaviest call the add-in makes,
run at the end of a round on a host that has just been through the self-test —
but the rounds since settled which side of that cost is worth defaulting to. On
2026-08-09, 29 of that deck's 36 slides read back with **zero shapes** and only
12 had a picture to check that against: 12 were confirmed empty on two
witnesses, and 17 stayed unknowable. The cap made them unknowable, not the host.
A readback of zero is one witness, and this host is known to report a shape
collection short without throwing; the picture is the second. **Leave it on** —
and expect a slower tail and a run log several times bigger. Untick it only for
a host that is already struggling, and know that you are giving up the second
witness when you do.

**"Merged to main" and "on the site" are different facts, and the stamp is the
only one that counts.** On 2026-08-06 they diverged for eight hours: four pull
requests merged green while every Pages deploy was cancelled by
`actions/deploy-pages@v4`'s ten-minute default timeout — GitHub was still
reporting `deployment_in_progress` when the action killed it, and re-running
restarted the same ten-minute race. The site went on serving `5d2c18a` and the
stamp said so; nothing else did. The timeout is now half an hour and
`workflows.test.ts` holds it there, but the habit is the real protection:
**read the stamp, don't trust the merge.** A stamp that will not change after a
hard reload means the deploy did not land — check the `Deploy Pages` run before
spending a session on it.

**The pane now checks itself, so the stamp turns RED when it is out of date.**
It reads `build.json` off the site on every boot and compares it to the build it
is running; if they differ it says so in place —
`abc1234 · … — SITE HAS def5678 · …` — and that is your signal to hard-reload
before doing anything else. Reading the stamp only ever helped if you already
knew which commit it should say, and that number is on GitHub, not in the pane.

The cache is why this exists rather than being a nicety. Pages serves the pane's
HTML with `Cache-Control: max-age=600` and gives us no way to set headers, so
for ten minutes after a deploy PowerPoint keeps handing back the cached page,
which names the previous hashed bundle — still in the browser's cache even after
that file 404s on the server. **Ctrl+F5 on the whole PowerPoint tab** is what
clears it; reopening the pane alone often does not, because the iframe re-asks
for a URL the browser still considers fresh. Two rounds were lost to this before
the check existed: one ran a fix that was not in the build under test.

A stamp with no warning on it means one of two things — you are current, or the
pane could not reach `build.json` at all. It stays silent when it cannot tell,
deliberately, because a warning that fires on every boot is one nobody reads.

**If a run dies, the evidence is no longer lost.** The run log is downloadable
only once a run ends, and two rounds were lost to runs that did not end: one
wedged at 1819 seconds, one killed by PowerPoint's own *"Sorry, we ran into a
problem"* at 108 seconds. Neither produced a log. Two things cover that now:

- **Live steps** is on screen the whole time, at the top of the Testing section.
  Copy or screenshot it before reloading — it is the only thing that catches the
  very last moments.
- Every step is also written to **browser storage** as it happens. Reopen the
  pane and a red note offers **Download the crashed run**. Send that file;
  `npm run triage` reads it on its own, with no deck
  (`node scripts/triage.mjs powerchart-crashed-run.json`).
- **The findings go to that same storage, not just the narration.** The probe's
  whole answer sheet is banked the moment the probe finishes, and each
  scenario's verdict the moment it lands. A round that dies half way through the
  battery now hands back every verdict it had already reached, plus the entire
  probe half — which used to die with the tab even though it had finished
  minutes earlier. `npm run triage` prints them under "finding(s) banked before
  it stopped".

  This also means **Verbose trace is no longer what makes a crashed round
  recoverable**. Leave it on anyway — the steps are what say where it stopped —
  but a round that crashes with it off is no longer worthless.

A third round was lost a different way, and it is worth knowing about because
nothing above would have caught it. The round **completed**, said "Saved as one
file", and the file never arrived — PowerPoint crashed seconds later. The pane
had already marked the run finished, which retired the stored record, and only
then attempted the download. Finishing and being saved are now separate facts:
a record stays on offer until its file is actually saved, and the banner says
whether the run crashed or merely never handed you its file. **If a round
reports success and nothing lands in your downloads, press *Download run log*** —
before reloading if you can, after reloading if you cannot.

The storage copy survives a closed tab and a reloaded pane. It does not survive
clearing site data, and it keeps the last ~2000 steps — a longer run loses its
beginning, never its end.

| # | test | what it catches |
| --- | --- | --- |
| ⭐ | **Probe, then self-test** — the button of that name, under the *Host probe* heading in the Testing section. (This row used to call it "Run the whole round", which is not a label that appears anywhere in the pane; the owner went looking for it and it was not there.) One click: the probe, then the self-test, then what landed on the slides — saved as **one file**. Send that, and nothing else: the file now carries every shape on every slide (names, ids, positions) and the host's own PNG of each slide the round added, so there is no deck to save and no screenshot to take. Replaces tests 0 and 1 below on any build where you do not need them separately. | Everything those two catch, in one click and one upload. The pane also diffs the probe's answers against the fake on screen, so a run that found nothing says so and needs no upload at all. |
| 0 | **Run host probe.** One click, nothing you made is touched — it works on scratch slides it appends, and tries to delete them again. On the web it does not always succeed (21 blank slides left behind on 2026-08-06, 14 in an earlier round), so the sheet now carries a `scratch-slides-returned` row saying how many came back; delete any leftovers by hand. Send the saved JSON. | Whether the FAKE POWERPOINT every test in this repo runs against is telling the truth. Its faults are things a real host taught us; its happy path is assumptions. `npm run host-diff` lines the two up, and each disagreement is either a fake that lies (so some tests are worth less than they look) or something the host does that we did not know. This used to say "once per host — it does not change between builds". It does change: two questions have now been observed ALTERNATING between two runs of the same build ninety minutes apart, and swapping back on a third (`UNSTABLE_ANSWERS` in `scripts/host-baseline.mjs` carries all four observations). So a repeat sheet is worth having even when nothing shipped, and **the sheet's own JSON is what the repo needs** — a pasted summary cannot replace `test/fixtures/host-answers-web.json`, which is what the CI contract gate diffs the fake against. |
| 1 | **Run host self-test.** One click, twelve scenarios (below), and the Live steps list names each one as it starts. Chasing one failure? Set **Scenario** to it first — it runs that plus the two inserts it needs, in seconds rather than minutes. Read the verdicts, and if the run does not finish, read the last line. | Nearly everything that used to be tests 1, 2, 3 and 7 of this table. The battery now selects shapes itself (`Slide.setSelectedShapes`, PowerPointApi 1.5), stops its own run, and asks the host to render a slide before and after drawing so it can tell a chart that is *there* from a chart that is *visible*. A verdict of **skipped** is not a failure. |
| 2a | **Demo deck — file path.** Path → **One file insert** → **Insert demo deck**. ~6 s. | The file half must report **38 of 38 complete** — anything less is a regression. |
| 2b | **Demo deck — shape path, on a FRESH deck.** Close the deck from 2a without saving, open a new one, then Path → **Shape by shape**. 1–2 minutes. | What the everyday path costs at 38× scale. Some items running short is the *measurement*, not a defect. This is the one thing no battery can stand in for: it is the only test at real scale, and scale is what crashes the tab. |
| 3 | **Look at the deck.** Scroll the 38 slides the demo run added. | The judgement a machine does not have. The battery's visibility check answers "did anything render"; it does not answer "is this the right chart, drawn well". Look for charts off the slide edge, overlapping labels, and anything that is visibly not what the gallery shows. |
| 4 | **One selection round trip, by hand.** Click a chart. The pane says "An SSF chart is selected." Press **Edit it**, change a number, **Update chart**. Drag the chart, then edit it again. | The battery drives the same machinery through `setSelectedShapes`, which is Office.js selecting a shape — not a human clicking one. Those are the same call in theory. This is the check that they are the same in practice, plus the drag-delta round trip (`POWERCHART_ORIGIN`), which needs a real drag and so cannot be scripted at all. |

**Why 2a and 2b are now two tests on two decks.** They used to be one click —
*Both, one after the other* — and that click has twice ended in PowerPoint's
own *"Sorry, we ran into a problem"*. The most recent run says why plainly: the
file half landed **38 of 38** and the downloaded deck was flawless (39 slides,
38 chart objects, 31 re-editable, nothing malformed), and then the shape half
started drawing onto that same 39-slide deck and got one batch of five shapes in
before the tab died. Both of this round's crashes have that shape — heavy shape
work on a deck that is already large — and running the two halves back to back
guarantees it. Separate decks cost one extra minute and remove the only
variable the two crashes shared.

**The dropdown now refuses "Both" on a deck that is not fresh.** It runs the
file half and says the shape half needs an empty deck, and the run log records
that it did so (`refusedShapeHalf`). Four attempts at both halves on one deck
have produced four crash dialogs; this stops the fifth from costing twenty
minutes. On a fresh deck the option behaves exactly as before.

**One extra click, once, on a fresh deck: set Scenario to *a selected shape
survives an insert* and run just that.** It takes about a minute.

This is an experiment with its answer written down in advance, not another
round. That scenario has stalled in every routine round on record — always on
its first batch of shapes, always after 45 seconds — and in the routine order it
always runs eighth and always straight after the selection ladder. Those are the
same fact, so no number of ordinary rounds can say whether the stall is caused
by the scenario, by what ran before it, or by how late in the round it is.
Picked alone it runs at about sixty seconds with only the two head-of-round
inserts in front of it, which separates all three:

- **It stalls** → the scenario itself. It is the only one that draws while a
  shape is still selected, and that would be the mechanism.
- **It passes** → the predecessor or the elapsed time, and the stall report now
  names the call the host last answered before giving up, which says which.

Send the run file either way — a pass is the finding here, not the absence of
one.

**Run it on a 4:3 deck at least once**, and on desktop PowerPoint at least once
per release — everything above is normally run on the web, which is where the
bugs have been, but it is not the only host and the battery is not a substitute
for a different one.

Then send back two files: **Download run log** (Testing section) and the deck
itself (File → Download a copy). The log carries the run's identity token, so
`npm run triage` joins it to the deck exactly rather than by guesswork.

**On the host probe (test 0).** It answers a fixed list of questions —
does a shape proxy survive a sync, does writing the same tag key twice
overwrite it, does a deleted shape report itself gone, does
`load('isNullObject')` populate the flag. Each one is a behaviour
`powerpoint.ts` depends on, so a divergence is never academic: the notes in
`scripts/host-diff.mjs` say, per question, what would be wrong if the real host
disagreed.

One divergence is already known without running it: the fake does not implement
`untrack`, and Office.js does. So every `untrack()` call in this repo is a no-op
under test, and the claim that SSF Charts releases proxies to avoid Office.js's
"too many proxy objects" warning is entirely unverified. That is the kind of
thing this probe exists to surface, and it surfaced one before ever reaching a
real host.

**What the first real run said, and what it cost.** PowerPoint on the web
answered question 1 and then reported *"the host would not resolve the scratch
slide"* for the other thirteen — which was our bug, not the host's. The slide
was in the deck the whole time; the host simply stopped answering
`getItemOrNullObject` for a freshly-added slide's id after the first lookup,
while still listing that same id in `slides.load("items/id")`. Three things
changed because of it: `addScratchSlide` now names its slide by diffing the
deck's id list rather than assuming the add went to the end, and proves the id
resolves before handing it out; the probe run replaces a scratch slide it has
lost so the remaining questions are actually asked; and `deleteSlideById` falls
back to deleting by position, because the old code read "the host will not
resolve this" as "already gone" and reported a clean-up it had not done. That
last one is the one that mattered beyond the diagnostic — the same call cleans
up after every off-screen redraw, on the user's own deck.

The one genuine answer from that run is worth keeping: `load('isNullObject')`
**does** populate the flag on PowerPoint web (it read back `false`), which
contradicts the negative claim `queueNullCheck` was written on. The workaround
stays — the host it was written for is real too — but it is harmless there
rather than necessary, and the comment now says so.

**What the second real run said (2026-08-04), and what it cost.** The same
lesson, one layer down. Fourteen questions came back; six were answered and
eight were not, and the eight all failed the same way — `GeneralException` on
the shape add, or a sync that never returned. The split is the finding: every
question that resolved a slide handle of its own was answered, and every
question that wrote through the handle `withProbeContext` had resolved for it
failed. A freshly-added slide's by-id handle is good for exactly one sync on
PowerPoint web, because resolving it is what makes Office.js rewrite its object
path to `getItem(id)` — which is the rule `SlideThunk` was already built on for
`getItemAt`, a thousand lines away in the same file.

`npm run host-diff` reported eight host divergences from a run that had asked
six questions, because `"threw"` and `"silent"` are real answers to those
questions. That is now impossible: a probe that cannot get its shapes answers
`no-scratch-shape`, which no probe can produce, and the diff reports never-put
questions in their own block instead of counting them as findings.

The same run's self-test failed its last scenario for the same reason —
`GeneralException`, `errorLocation: SlideCollection.getItem`, on
`slide.getImageAsBase64(...)` — so this one was never only about the
diagnostic: `slideImageBase64` held a slide handle across two syncs on the
user's own path. Three new probe questions
(`shape-add-fresh-slide-proxy`, `shape-add-held-slide-proxy`,
`shape-add-positional-slide-proxy`) ask the three ways of naming that slide
apart, so the next sheet says which of them this host will actually take.

**The third run (build a609c9c) answered all seventeen**, and settled that
fork: fresh by-id **yes**, by-index **yes**, held-across-a-sync **threw**. It
is the *holding* that fails — not the id, not the slide's newness — so the fake
now models a single-sync by-id slide handle unconditionally rather than behind
a fault, and `shape-add-held-slide-proxy` is the one baseline answer that is
expected to be a refusal.

Four of the previously-unasked eight are now genuinely answered:
`tags-on-fresh-shape` **yes** and `delete-then-lookup` **reports-gone** (both
agreeing with the fake), `getcount-populates-same-sync` **yes**, and
`shape-proxy-survives-one-sync` **unreadable** — office-js#2903, confirmed
first-hand at last.

The other four came back with answers that were about the probe rather than
the host: `tag-on-group-survives` **no**, `tags-add-same-key-twice` **other**,
`shapes-items-count-honest` **unreadable**, `addgroup-returns-usable`
**unreadable**. Each of those probes wrote or read through a shape or group
proxy from an earlier sync — the pattern the same sheet proves fails. A fourth
run (build `a2191fe`) reproduced all seventeen answers exactly, so they are
stable and reproducible, and still not about the host.

Taken at face value the first of them says no chart in any deck is
re-editable, which the same run disproves: its repair pass landed 23 retags on
grouped charts. So all four were **withdrawn and re-asked** rather than acted
on. The rule the rewrite follows is one line — *only an id crosses a sync,
never a handle* — and it now holds for every probe:

- group members are resolved in the batch that groups them, and the group's id
  is loaded in that same batch;
- tags are written and read through a shape resolved in each batch, exactly as
  `settleAndTagChart` does on the path that carries real charts;
- a probe that loses its scratch slide *part way through* a question says
  `no-scratch-shape` and gets asked again, instead of reporting whatever the
  next line did to an undefined collection.

`shapes-items-count-honest` is the one that could not be cleaned up this way —
a collection load is queued in one batch and read in the next by definition —
so it got a partner instead. `shapes-items-via-positional-slide` asks the same
question through `getItemAt(index)`. If that one reads back and the by-id form
does not, the collection was never the problem and the parent handle was,
which decides how every readback in `powerpoint.ts` should name its slide.
Read the pair together or neither.

What to read in the result: the `tagging failed` count (was 28 on the last slow
run; should be near zero), any line annotated `^ known host bug: office-js#…`
(Microsoft's, not ours — annotated automatically, don't chase it), the nine
self-test verdicts, and — new — the **`phases an error escaped`** block.

That block is the half of a failure that used to be missing. Office.js reports
what it refused (`errorLocation`, `code`); it cannot report what the add-in was
doing at the time, and reconstructing that from timestamps and call order is
what made the last three real-host failures cost a session each. Every error
now carries the phase it came out of — `at=drawing the chart's shapes`,
`at=writing the chart's config tag`, `at=reading slides 20-38 for charts` — in
its own message, in the trace as it happens, and tallied by `npm run triage`.
An error with no `at=` came from somewhere with no `step` around it, which is
itself worth reporting.

Record anything broken as issues; fix per the lockstep rules. Real-host
degradation paths that are *expected* (not bugs): radar fills are
outline-only, pattern fills render solid.

**A programmatic `setSelectedShapes` wedges the web host's selection
subsystem.** Measured on PowerPoint on the web, twice, on build `55011a3`:

- `Slide.setSelectedShapes([id])` is GA at PowerPointApi 1.5 and **takes the
  call** — no error, no refusal, `context.sync()` resolves.
- Every selection call after it goes silent. `getSelectedShapes` (the read
  behind *Edit it*) ran out its full 90-second budget; so did the
  `setSelectedSlides` that follows it. Nothing throws — the host simply never
  answers, and a run that would otherwise take seconds took 159.

Two selection bugs are already open against the web host — office-js#3083
(`setSelectedShapes([])` does not clear on web) and #3698 (no picture insert
while a shape is selected) — and this is the same family: the web host's
selection layer accepts writes it cannot then serve. It is Microsoft's, not
ours, and it is now **gated**: the *edit the chart the user selected* scenario
waits ten seconds rather than ninety, and reports **skipped** with the reason
rather than red. `npm run triage` annotates it as a known host bug
automatically.

**Nothing in the add-in is affected by it**, which is why it is a gate and not
a fix: the pane never selects a shape programmatically. It reads the selection
the *user* made, and a user's own click leaves the subsystem answering
normally. That is exactly what test 4 of the standing run checks by hand, and
on the web it is the only thing that can.

**Two more things that are expected and look alarming.** Neither is a defect:

- **A blank slide appears at the end of the deck and disappears again** while a
  chart redraws. That is deliberate. Updating a chart is the add-in's worst
  case on the web, and it is only bad because the slide is on screen — so the
  view is parked elsewhere for the redraw. A one-slide deck has nowhere to
  park, so one is made and removed afterwards. If a blank slide is ever left
  behind, *that* is the bug; the pane warns when it cannot remove it.
- **The view jumps to another slide and back** during an edit, for the same
  reason. Your own navigation wins: click away mid-redraw and you stay where
  you clicked.

**More on slide size.** It is read from the host now
(`PageSetup` at 1.10, an exported slide at 1.8, `getFileAsync` below that) and
drives both chart placement and the generated deck's declared size. Every
earlier session tested 16:9 only, where a wrong answer is invisible — on 4:3 it
shows up as charts off the right edge or a demo deck rescaled on insert. The
run log's `slide size` line records the value *and which rung produced it*;
`assumed` there means nothing answered.

**Wider feature sweep** — once per release, rather than per build:

1. Ribbon shows the SSF Charts menu; pane opens; gallery renders.
2. Pie chart on a 1.10+ host (triangle-fan rotation), grouping on 1.8+.
3. **Use deck theme** on a 1.10+ host pulls the template's accent colors.
4. Elements (harvey ball, table with a total row) and Agenda insert.
5. Excel: select a range → Generate → paste JSON into PowerPoint pane →
   Import → chart matches.

Four questions come from the **office-js issue tracker** rather than from
something this project measured: `picture-then-shape-read` (office-js#5022, a
`context.sync()` that runs indefinitely after an image is inserted — the shape
`drawDemoItem` takes for a degraded chart), `group-of-existing-shape-readable`
(#5849, `Shape.group` throwing — how `countGroupChildrenPage` reads every chart
it checks), `slide-layout-readable` (#3826) and `layouts-readable` (#4906,
#2328, reported only on decks built from a **custom template**, which is what the
owner's are). All four are open upstream, all four sit under code this add-in
ships, and none of them can be pinned on this host without asking.

Some questions now ask their **own follow-up**, in the same run. When an answer
admits two readings that lead opposite ways, the probe puts the partner question
immediately instead of waiting for someone to notice and ask for another round —
which is what the last two ambiguities cost, a session each. The sheet says which
answer triggered which follow-up, so the pair reads as a pair.

The one wired today: if `slides.getItem(id)` refuses a **freshly added** slide,
the probe immediately asks whether it resolves a slide that was already in the
deck. Refused-because-new means the everyday insert path is fine and only one
caller is at risk; refused-outright means it is broken for everyone. The
follow-up only READS — repeating the shape add on a pre-existing slide would draw
in the owner's own presentation, and a diagnostic that litters a deck is one
nobody clicks twice.

### The host self-test

Eight paths existed only as items on this list for a human to remember to try,
which in practice meant eight separate sessions. They are now one button:

| scenario | what it proves |
| --- | --- |
| insert on top of an earlier run | the run token keeps two runs' slides apart, instead of one being deleted as the other's duplicate |
| two slides claiming one slot | the repair pass drops one copy and keeps a working one — not both, not the wrong one |
| edit a chart on the visible slide | the live-canvas redraw survives with the slide genuinely on screen |
| insert onto a slide that already has content | the everyday action — a chart drawn onto a slide that is not blank stays grouped and re-editable, and does not swallow what was already there |
| same scale across the deck | every chart in the deck redraws to a shared ceiling and stays re-editable — the ceiling has headroom above the data on purpose, so a host that stores the new config and redraws the old picture is caught rather than counted as a pass |
| explode a degraded picture | a picture keeps its config and can become native shapes again |
| which selection call wedges the host | *which* call stops the web host answering — the one measurement that settles it |
| edit the chart the user selected | the read behind *Edit it* — the only entry point a real user travels on |
| stop a run part-way | a stopped run adds nothing and leaves nothing behind claiming to be a chart |
| a selected shape survives an insert | whether office-js#2775 is live here — on the web, adding a text box deletes the shape that was selected, and every chart drawn here has text boxes |

**Back in that list as of `c7d91d5`: the chart is actually visible.** It proves
the host's own render changed where the chart was drawn — not just that shapes
exist — and it spent five rounds killing the browser tab before it ever proved
anything. Every one of those stopped writing inside the scenario, within a step
or two of `adding a scratch slide`.

**The fifth round was picked alone, and that identified the cause.** The first
four crashed around 600 seconds with nine scenarios' worth of drawing behind
them, so "this scenario kills the host" and "ten minutes of drawing kills the
host, and this is merely what was running" both fit. Picked alone on `b998a2e`
it was reached at **61.5s** with only its two inserts in front of it, took a
scratch slide, logged `rasterising the empty slide` at 61.8s, and the tab died.
Those two inserts head every round and kill nothing.

The surface, precisely: `getImageAsBase64` on a slide added 0.3 seconds earlier
— the fifth distinct way that call has failed on a fresh slide and the first
fatal one. It does its before-and-after on a slide the run added earlier now, so
it never makes that call, and on `c7d91d5` it passed: `10064 → 15652 bytes`.

Two things that trade cost, both paid: the round after the change reported
*"nothing was drawn"* over twenty-four committed shapes (the host would not name
the group, and a null target was read as an empty slide), and the round after
THAT left two full-size charts drawn over each other. The verdict now comes from
the image with the naming failure as a caveat, and the chart it draws is a
quarter-size one in the corner.

The ladder — **which selection call wedges the host** — is in that list now, and
used to be a separate run. It runs ahead of every scenario that selects a shape —
*a selected shape survives an insert* and *edit the chart the user selected*.
Being the FIRST such call is the property it needs; being alone in a run was a
stronger condition than the question requires, and it cost a five-minute round
every time. When the ladder does find a wedge, both scenarios after it report
*skipped* with the ladder's own words instead of spending a budget each to learn
less.

It climbs from the least invasive selection call to the most — read the
selection, `setSelectedSlides`, read, `setSelectedShapes([id])`, read,
`setSelectedShapes([])`, read — and **stops at the first one that goes silent**,
reporting that call and the last one the host answered. It stops because after a
wedge every later call is silent too, so climbing on would report four timeouts
and name nothing, which is the ambiguity it exists to remove.

Position matters, and two earlier answers were wrong about it. Putting the
ladder LAST in the battery let `edit the chart the user selected` wedge the host
six scenarios earlier, so the ladder reported silence on its own first rung — an
answer about nothing. Putting it THIRD, right after the two inserts, is worse
still: the ladder can wedge the host itself, so six scenarios would then run
against a wedged one instead of two. Ahead of every other `setSelectedShapes` caller is the
position that gives the property without the cost — stated as that property in
the tests, because adjacency was only ever a proxy for it and the proxy broke
the moment a second such scenario arrived.

**Two more, and a full run deliberately leaves them out.** Pick them by name
from the Scenario menu:

| scenario | what it proves |
| --- | --- |
| edit the chart YOU click | the pane's most-used read, driven by a real human click instead of a call that behaves differently |
| what makes a long run slow down | whether a long run degrades because of the request CONTEXT, the DECK, or the TAB — the three every artefact so far confounds. Picked-only, and it killed the tab on its first outing (`25407ed`, 26.9s in) before measuring anything. Its host calls are traced individually now, so the next attempt names the one that does it. |

**What makes a long run slow down** is the newest, and it is the only one here
that is purely a measurement. Every real-host artefact this project owns
degrades — 496 seconds to reach scenario seven, a 38-item run whose later
inserts cost multiples of its first, a tab PowerPoint eventually killed — and
not one of them can say which of three things did it, because an ordinary run
grows its deck, ages its tab and holds a request context open all at once.

It draws the same shapes twice, onto a slide each: once inside ONE request
context held open across every batch, once with a fresh context per batch.
Separate slides on purpose, so neither arm draws onto a slide the other has
already filled. Then it reads the two curves:

- only the long-context curve climbs → **the context** is what degrades, and
  shortening contexts is a fix we know how to write.
- both climb together → **the host** is slowing as the deck grows or the tab
  ages, and shortening contexts would not help.
- neither climbs → the thing that kills a long run is not in this loop, which
  rules out three suspects in one click.

The verdict is a sentence, but the raw millisecond curves are in the detail and
in the run log too — the threshold that turns them into a sentence is a judgement
call, and the numbers are what someone can re-read when it turns out to be the
wrong one. It costs about a minute and leaves two slides of small rectangles
behind.

It reports **ok** whatever it finds. It is an experiment, not an assertion — a
host that answers every rung is a real result, and one worth knowing.

**Edit the chart YOU click** is the other picked-only one, and it is the answer
to a thing this project got wrong for months. `setSelectedShapes` is Office.js
selecting a shape; a human clicking one is the same call in theory and
demonstrably not the same in practice on the web. So the battery stops
pretending and asks: pick the scenario, and it waits 30 seconds for you to click
an SSF chart, counting down in the pane. Click one, and it runs the whole chain
the pane's *Edit it* button runs — read the selection back, edit through the
target that read produced, confirm the result is still re-editable — and reports
what happened.

It listens on `DocumentSelectionChanged`, a **Common API** event that does not
go through the subsystem the programmatic select wedges. Nothing in it calls
`setSelectedShapes`. If nobody clicks it reports **skipped**, never failed.

This is what old test 4 of the standing run was, with the guesswork removed: the
same click, and now a recorded verdict instead of "looked fine". Test 4 still
earns its place for the **drag**, which needs a real mouse and cannot be
observed any other way.

Each verdict says what was observed, not just pass/fail, and a scenario that
throws is recorded and the rest still run — a battery that stopped at the first
error would spend a whole session to learn one thing. A scenario the host
cannot run is reported as **skipped**, kept apart from a failure: "we did not
check" and "we checked and it is broken" send a diagnosis in different
directions.

It leaves its slides in the deck on purpose — save the file and hand it to
`npm run triage` with the log.

**What it cannot cover.** For a long time this paragraph said Office.js had no
way to select a *shape*. That was wrong — `Slide.setSelectedShapes` has been GA
since PowerPointApi **1.5** — and the belief cost this project four manual tests
that could have been buttons. It is scripted now.

What it still cannot cover is a **human** selecting: a real click, and the drag
that precedes an edit (`POWERCHART_ORIGIN`). Test 4 of the standing run is the
only coverage those get, which is why it is still manual. And on PowerPoint on
the **web** the scripted version does not work at all — see below — so on that
host the manual round trip is the *only* coverage the selection path has.

### Reading the demo-deck self-check (post-#212–#216)

The **Insert demo deck** button runs every chart kind, appends a results
slide, and posts a summary in the pane note. Ten harness-reliability PRs
since v0.2.0 mean the note now carries strictly more signal than "N of M
rendered":

- **`rendered`** — chart landed with every expected shape and was grouped.
- **`late-settled`** — a sync timed out but the shapes committed anyway;
  the harness read back the slide and trusted the count (no dup slide, no
  NOT COMPLETE stamp). Counted as rendered.
- **`rendered-partial`** — a sync threw with ≥85 % of the expected shapes
  on the slide. Counted as rendered; the fresh-context rescue groups what
  landed so the chart is still re-editable.
- **`ungrouped`** — chart shapes are on the slide but not one group. Not
  re-editable via the `POWERCHART_CONFIG` tag; a rescue attempt already
  ran. Flag for investigation.
- **`failed`** — under the 85 % gate after both attempts AND the
  unstamp+rescue path could not group what landed; slide carries the red
  NOT COMPLETE banner.
- **`BLANK: <title>`** — slide committed but readback showed zero shapes,
  and the slot tag names which item was on it (host lost the content).
- **`N of M results pages added`** — the run's own results slide
  paginated; each page is attempted independently, so a partial landing
  no longer drops the record.
- **`addsLostAtCommit=N`** — `addSlides` confirmed N `slides.add()` calls
  never landed even after its own retry. Correlates with the office-js
  bug documented in `OFFICE_JS_LOST_ADDS.md`. `addsIssued − slidesAdded`
  is the wider gap.

A clean run reports every chart rendered + grouped, no ungrouped/blank/
`addsLostAtCommit`. The full console.table dump under **F12** carries
`shapes`, `status`, `grouped`, `ms`, `abandoned` and `lateOutcome` per
item. `abandoned` means a deadline fired inside that item — that is the
column to scan for "which chart did the host stop answering for".
`lateOutcome` is how the abandoned call eventually ended, and it is only
filled in when the host answered soon enough to still be paired with the
item; observed answers often arrive minutes later, and those land in the
`a call we gave up on finally answered` trace instead. So an `abandoned`
item with an empty `lateOutcome` means "no answer yet", not "no stall".
Better
still, **Download run log** writes the whole run to JSON — both insert
paths, the settled repair verdicts, and the activity trace when Verbose
trace was on. That file is the right attachment for a Phase-2 regression.

## Phase 3 — Activate the Claude skill ([owner])

1. Download `ssf-charts.zip` from the latest release (the rolling
   [`skill-latest`](../../releases/tag/skill-latest) is rebuilt on every
   merge).
2. claude.ai → Settings → Capabilities → **Skills** → upload the zip.
3. Test from any Claude surface: *"Make me an EBITDA bridge: FY23 86,
   Volume +14, Price +9, Cost −12, FX −4, FY24 total"* → expect a .pptx
   with native shapes. Then test inside **Claude for PowerPoint** (the
   add-in from AppSource) — skills enabled in settings are available there,
   which closes the loop: Claude builds SSF Charts charts directly in the
   user's deck.

## Phase 4 — Cut the release ✅ v0.2.0 done

1. **[agent]** ✅ done — `gh workflow run release.yml -f version=v0.2.0`
   (the workflow creates the tag; the git proxy in remote sessions can't push
   tags). Trigger it from a green `main`: the release job runs `npm test` but
   **not** typecheck or the coverage thresholds, so it is a weaker gate than
   `ci.yml` and trusts the branch it builds from.
2. **[agent]** ✅ done — README carries the live gallery link and an install
   section pointing at the prod manifest download.
3. **[agent]** ✅ done — CLAUDE.md's "Pending / user-gated" list now names
   only what is genuinely still owner-gated.

**v0.1.0 shipped the DEV manifests** (`manifest.xml`, pointing at
`https://localhost:3000`) because it predates the Phase 1 change that attaches
the prod pair — so the README's install path was broken for anyone who tried
it: the file it names, `manifest-prod.xml`, was not in the release at all.
v0.2.0 fixes that. If a future release ever ships a manifest again, check the
asset list, not just the workflow file — the workflow was right for 12 days
while the published release stayed wrong.

## Publication readiness, audited 2026-09-10 — what is fixed and what is yours

A pass over everything between "main is green" and "a stranger installs this
from AppSource". The technical gates were already clean and stayed clean:
`office-addin-manifest validate` says valid, `vite build` and the manifest check
pass, CI is green, and every one of the 16 URLs the prod manifest references
returns 200. What follows is what those gates do not look at.

### Fixed, and the reason each survived every existing check

- **`HighResolutionIconUrl` was 80×80; Office requires 64×64.** Wrong in all
  four manifests since they were written, with the correct `icon-64.png` sitting
  in `assets/` and on the deployed site the whole time. Survived because
  `check-published-install.mjs` HEAD-checks that a URL *resolves*,
  `office-addin-manifest validate` checks the *schema*, and `manifest-rules.mjs`
  never opened an image. **A URL that 200s is not a URL that is right.** Fixed,
  plus a rule so it cannot drift back.
- **The store listing promised "never flat pictures or opaque objects".** Below
  PowerPointApi 1.10 the add-in deliberately inserts a picture and says so.
  "Functionality does not match the offer description" is the ordinary AppSource
  rejection. Copy corrected; see `STORE-LISTING.md` for the measurement and the
  right way to say it if the guarantee is wanted back.
- **Four Elements-tab previews shipped as `role="img"` with no accessible name**
  — a WCAG 1.1.1 Level A failure, found with axe-core against the live pane. The
  `desc` written for charts never reached the element family, and every test
  guarding it asked about a chart. All five builders fixed; the test now
  iterates the family.
- **"18 chart kinds"** in the listing; `CHART_KINDS` has 25.
- **The 60-second watchdog** pointed an ordinary user at *Download the crashed
  run*, a control that only exists for a harness run.

### Yours, and the first is the one that blocks everything else

1. **Cut a release.** `check-published-install.mjs` now exits 3: the manifests in
   v0.5.0 are not the committed ones, so **the icon fix reaches nobody until a
   release is cut**. A tag auto-publishes, which is why this is not done here.
2. **The pane ships its own test harness to every user.** Automation ▸ Testing —
   self-test, host probe, download run log, clean up the last round — is live in
   production, ungated, in harness vocabulary, with buttons that rewrite the
   reader's deck. `TESTING_UI_NEEDS_OPT_IN` (`src/taskpane/app.ts`) hides it
   behind `?harness=1` and is committed **false**, because flipping it changes
   what a user receives. Flip it *and* add a harness manifest together — a test
   fails if you do one without the other, because otherwise the round loop stops
   without saying so.
3. **The manifest admits hosts the product does not fully work on.** It declares
   PowerPointApi **1.4**; the product is whole only at **1.10**. Three tiers
   exist and only the top has ever been tested — 1.10 (web: everything works),
   1.8–1.9 (pie/doughnut/sunburst and any CAGR or difference arrow insert as a
   flat picture), and ≤1.5 including Office LTSC (**no picture fallback at all**
   — marks are simply missing). Raising `MinVersion` to 1.10 makes every promise
   true and costs LTSC and pre-2026 builds. That is a reach-versus-correctness
   trade and it is a product decision.
4. **The default chart is the degraded case.** The pre-selected tile is
   `stacked`, and its sample carries `cagr: { from: 0, to: 3 }` — an arrowhead,
   which is exactly what is dropped below 1.10. So the first thing a new user
   does, one click with no input, is the broken case on every host below 1.10.
   Changing the default sample is the cheap half of decision 3.
5. **Screenshots** remain a person's job — the slide canvas does not composite in
   a headless browser. Unchanged from `STORE-LISTING.md`.

### Verified fine, so nobody re-checks them

Privacy and terms pages are live, specific and accurate — they disclose hosting
logs and the Office.js CDN, and the "your data never leaves your device" claim
holds: the shipped code contains exactly **one** network call,
`fetch("/build.json")`, same-origin, sending nothing. The 300×300 store logo
exists at the right dimensions. User-facing language is clean — 94 `note()`
calls and 121 catalogue entries with no internal vocabulary leaking.

## Distribution beyond sideloading (later, optional)

- **Org-wide (BESTSELLER)**: a Microsoft 365 admin deploys the manifest
  centrally via Admin Center → Settings → Integrated apps → Upload custom
  app. No store review; appears for chosen users automatically. Fastest path
  for internal use — recommended before attempting the public store.
- **AppSource** (public store): requires a Partner Center account (free for
  Office Store apps) and Microsoft validation (works on every claimed platform,
  WCAG, privacy + terms + support URLs). Substantial process; only worth it if
  SSF Charts should be publicly installable. Prep is staged:
  - **[agent, done]** Hosted **privacy** (`/privacy.html`) + **terms**
    (`/terms.html`) pages (in `public/`, built to the site root), a
    trademark-clean store listing in `docs/STORE-LISTING.md`, and the
    store-facing manifest `<Description>` reworded off the "think-cell" mark.
  - **[owner]** Create the Partner Center account, produce the listing images
    (300×300 logo + screenshots), run `office-addin-manifest validate`, then
    submit. Full checklist in `docs/STORE-LISTING.md`.
  - ⚠️ **Trademark:** keep everything store-facing (name, description,
    screenshots) free of the "think-cell" mark — internal docs may keep it.

## Known constraints to keep in mind

- Requirement sets: shapes need PowerPointApi **1.4+** (Win 2207+, Mac
  16.62+, web; not iPad); grouping 1.8, re-edit tags 1.3, pie rotation 1.10,
  theme colors 1.10. The pane degrades gracefully below each. **This paragraph
  said otherwise until 2026-09-01** — that below 1.10 the pie family "inserts
  with no slices at all" — and that stopped being true when the warn-and-picture
  decision shipped on 2026-08-31. A chart carrying marks the host cannot draw
  now inserts as a **picture**, complete, with a message naming what would
  otherwise be missing; where no picture is possible it still draws and still
  says what is absent; and Explode refuses on the same test rather than handing
  back the version that cannot be drawn. Measured then: 18 of 123 shipped charts
  lose ink below 1.10 and 8 lose their subject entirely (pie 4/4, doughnut 2/2,
  sunburst 2/2). Diagonal lines used to fail there too, drawn as flat
  rectangles, and that was fixed on 2026-08-30.

  It still matters for a submission, but as something to describe rather than
  fear. Microsoft's validator confirms this add-in is tested on PowerPoint on
  the web, on Windows (Microsoft 365), and on Mac (2019+ and Microsoft 365) —
  exactly where 1.10 is most often absent. Both prod manifests validate clean:
  `npx office-addin-manifest validate manifest-prod.xml` and the same for
  `manifest-excel-prod.xml`, run 2026-09-01, "The manifest is valid".
- Pages is static HTTPS — exactly what an add-in needs; no server code, no
  auth, no cost. If the repo must stay private instead, any static HTTPS
  host works (Azure Static Web Apps free tier, Cloudflare Pages) — only the
  base URL in the prod manifests changes.
