# Round loop journal — started 2026-08-14 20:50, hard stop 2026-08-15 04:50

Brief: `loop-brief-v2.md` beside this file. One round per wake-up, five-part
protocol as a gate, journal before reporting.

## Round 29 — 3c09b6a — 10/12 pass (before the loop, kept for continuity)

Two attempts; first wedged 8s into question 4 (crash dialog), cleared in 90s.
Archived `rounds/029-3c09b6a.json`, merged as #488.

- **Mine** — `left=73 deckBefore=1 deckAfter=1 stillListed=0`: the scratch slides
  never landed. Id namespaces differ (`4123571114#…` vs `287#62081387`).
- **Research** — office-js#2903 confirmed as the 5010 source (18 hits).
- **Instrument** — none needed; the handback line already carried the answer.
- **Fix** — driver poll no longer ends a round on one failed CLI call;
  `archive()` writes 2-space JSON; `rounds/*.json` in `.prettierignore`.
- **Doctrine** — `docs/BACKLOG.md` scratch entry corrected: the counter is the
  bug, not the deck.

## Round 30 — 6d6e278 — started 20:57

Preconditions: host was silent (form 2, no dialog) AND the pane was a build
behind; one `pw reload` + reopen fixed both. Deck cleaned 7 → 1. Check green:
`host answered in 9ms · ready`.

WEDGED. Driver hit its 30-minute limit; the new poll behaved correctly (no false
"pane stopped answering"). Crash dialog up, deck still 1, `--check` named it.

- **Mine** — died at `shape-add-fresh-slide-proxy`, the fourth round to do so,
  but the recorded answer is `no-scratch-slide ms=90135`: the probe never got as
  far as asking. Three 90s timeouts on `listing the deck's slides` followed, then
  it gave up after 3 silent questions.
- **Research** — pulled the ULS batch (request 47998). Same signature as the two
  earlier crashes, and it corrects my own framing: the `PptApi Call` sequence in
  flight is GetSlides / GetId / GetItemOrNullObject — READS, the scratch-slide
  resolution. No `addGeometricShape` anywhere. `Failed to restore selection after
  load content` appears in all three.
- **Instrument** — none added; the ULS route is now documented and took two
  commands. The probe already records the scratch id and deck length.
- **Fix** — nothing add-in-side to fix: the host's log says the write never
  happened.
- **Doctrine** — `docs/ROUNDS.md` corrected: the crash fingerprint is a read
  sequence, and "the round's first write" was wrong.

Recovery: clicked Refresh by text, reopened pane, cleaned deck 1→1, check green
(`host answered in 4ms`). Round 31 started 21:3x.

## Round 31 — 6d6e278 — started 21:36

WEDGED, same signature, ~3 minutes in. Crash dialog up, deck 1.

- **Mine** — identical shape to round 30: `shape-add-fresh-slide-proxy` →
  `no-scratch-slide` after 90s, then three more 90s timeouts, then gave up.
- **Research** — nothing new to look up; the signature is now understood.
- **Instrument** — THE ONE THAT MATTERED. Rounds 30 and 31 each died at ~3
  minutes and held the driver's poll for the full 30. The poll now looks for the
  crash dialog every iteration and stops when it appears, reporting how far in.
  A DOM read, not a ping — a ping would interleave with the round's own batches.
  Merged as #490.
- **Fix** — same commit.
- **Doctrine** — already corrected after round 30 (#489).

## Round 32 — abcb7c5 — started 22:07

Preconditions green after a reload to the new build; deck cleaned to 1;
`host answered in 3ms`. First round to run with the mid-round crash watch
in the driver — though the watch only takes effect for rounds started after
#490 merged, so this one still has the old 30-minute behaviour.

WEDGED ~3 minutes in, same signature. Started before #490 merged, so it still
burned the full 30 minutes.

## Round 33 (first try) — 8b287e3 — 22:48

WEDGED **2 seconds in** — and the new crash watch caught it, live-verified: the
dialog was genuinely up and the host genuinely silent. 30 minutes saved.

Crashes are accelerating: 3 min, 3 min, 3 min, 2 s. Five of six attempts tonight.

- **Mine** — nothing new in the trace; it never got far enough to have one.
- **Research** — checked whether a fresh document would help. Dozens of scratch
  decks exist in OneDrive but the list gives no clean refs, and sideloading the
  add-in into a new one needs the file chooser, which is the owner's step. Not
  pursued.
- **Instrument + Fix** — `--retry N` (#491): the driver now clears the crash and
  starts again itself, encoding the fifteen-step recovery done by hand six times.
  Only a crash is retried; mutation-proven three ways.
- **Doctrine** — the recovery's traps are in the code's own comments: Refresh
  matched by TEXT, pane reopened after the reload, Automation before the run
  button, deck cleaned last.

## Round 33 (retrying) — cabb357 — started 22:57

`--retry 6`, unattended. Check green: `host answered in 3ms`.

FINISHED first try. Archived `rounds/030-cabb357.json` — 10 of 12.

## Round 34 — cabb357 — 23:10, archived as 031

Two crashes at 2s each, both cleared and restarted by `--retry` with no hand on
it, third attempt finished. First fully unattended recovery. 10 of 12.

## The thread worth pulling — `same scale across the deck`

Three rounds, same failure, and it is a real feature breaking:

    029  5 of 8 carry the scale; 2×no-config;                  flipped at 6
    030  4 of 8 carry the scale; 2×no-config;                  flipped at 5
    031  4 of 8 carry the scale; 1×unknown-shape, 1×no-config; flipped at 4

Every loss is `InvalidParam passed to GetItem(id)` (5010, office-js#2903) at
`writing the chart's config tag`. "Flipped" = two consecutive losses and the
scenario stops, so the score is a floor.

- **Mine** — traced it to `finishCharts` writing the tag through
  `shapes.getItemOrNullObject(id)` with an id read a sync earlier.
- **Research** — office-js#2903 already catalogued; nothing new to find.
- **Instrument** — `tag-through-refetched-shape` (#492): asks whether the id
  round trip is what the host refuses. `tags-on-fresh-shape` already says the
  fresh shape's own `.tags` works, so this is the untested half.
- **Fix** — deliberately NOT yet. A rewrite of this path on a theory is exactly
  what has been reverted before. The question decides it.
- **Doctrine** — PENDING_QUESTIONS entry states what each answer means.

## Round 35 — 360309f — started 00:05

First round carrying `tag-through-refetched-shape`. `--retry 6`, unattended.

FINISHED first try. Archived 032. 10 of 12.

**The new question answered immediately: `no-id`, every pass.** The host will not
report a freshly created shape's id at all, so the re-fetch was never reached.

Read with its neighbours that IS the answer — `tags-on-fresh-shape yes`,
`shape-proxy-survives-one-sync unreadable`, `shapes-items-count-honest
unreadable`. Wrote it up in RESEARCH.md as "a shape is reachable only in the
batch that created it", and added the partner `tag-the-creation-proxy-a-sync-later`
because production falls back to the creation proxy when the id is unreadable —
exactly the case this host produces. #493.

## Round 36 — 86eaf65 — 00:40, archived as 033

FINISHED first try. 10 of 12.

**The partner answered `yes`** — a tag CAN be written through the creation handle
a sync later. That corrects what I wrote an hour earlier:

    write through an aged creation handle   works
    read  through an aged creation handle   unreadable
    re-fetch by id                          no id to re-fetch by

- **Mine** — five rounds now, `same scale` fails identically every time.
- **Research** — none needed; the host answered directly.
- **Instrument** — the partner question, which paid for itself in one round.
- **Fix** — still not written. The evidence now points at "stop reaching for the
  id, keep the handle" rather than "tag at creation time", which is what the
  overstatement would have built.
- **Doctrine** — RESEARCH.md corrected in #494, with the wrong version kept
  visible rather than quietly replaced.

## Round 37 — 4d68e73 — started 01:20

Third sample of the `no-id` / `yes` pair. `--retry 6`, unattended.

FINISHED first try. Archived 034. 10 of 12.

The decisive field, found by reading the trace rather than the headline: every
settle-pass failure reports `withId: 0`. The repair that exists to rescue a lost
config tag resolves shapes BY ID, and this host never yields one.

- **Fix attempted and abandoned correctly** — wrote the obvious theory (the
  positional `load` shares a batch with the tag writes) as a test expecting to
  FAIL. It passed on the first run: the insert path already survives that. Kept
  the test, labelled behaviour-not-guard. One wrong theory caught by writing the
  test first. #495.

## Round 38 — 0e230db — 03:00, archived as 035

FINISHED first try. 10 of 12. Third sample: `no-id` ×4 rounds,
`tag-the-creation-proxy-a-sync-later: yes` ×3. The resample bar is met.

**Where the fix belongs, corrected:** NOT the settle pass. It opens a fresh
`PowerPoint.run`, so the creation handle does not exist by the time it runs —
its only routes are an id and a collection read, and this host refuses both.
The retry belongs in the DRAWING context, before the handle is thrown away.

Same goal as `binding-names-shape-later` (`unreadable` here, four rounds); the
creation-handle write is a second route that answers yes and needs no 1.8
surface. #496.

## Round 39 — cfbf434 — started 04:05

Final round of the night. `--retry 4`.

FINISHED first try. Archived 036. 10 of 12. Fourth sample of the pair:
`no-id`, `tag-the-creation-proxy-a-sync-later: yes`.

## Night's tally — 2026-08-14 20:50 to 2026-08-15 04:30

Nine rounds attempted, **six landed** (029-036 archived, less the wedged ones),
all 10 of 12. Twelve PRs merged, every one gated on exit codes.

Rasterise arms: 24 draws per arm pooled over 12 rounds. Still needs 60-100.

What the night actually bought, in order of what it cost to learn:

1. The wedge is the host's session dying — PowerPoint's own crash dialog or a
   severed session — not the add-in. A reload clears both, in seconds.
2. The scratch "leak" is a counter bug; the deck never grew.
3. A shape can be WRITTEN through its creation handle a sync later, but never
   read back, and its id never round-trips.
4. `same scale across the deck` fails because the settle pass resolves by id and
   this host has none — `withId: 0`, every failure, six rounds.
5. Therefore the retry belongs in the drawing context, not the settle pass.

Three wrong theories were killed by measurement rather than argument: the
burnsTheSlide question, the shared tag/load batch, and "both routes out of the
creating batch are closed".


## Resumed window — 2026-08-15 07:20 to 12:20

### Round 40 — 3c163c4 — archived as 037 — 10/12

**All three overnight instruments delivered on their first real outing.**

- `from: created×1` ×7 and `created` ×4. **Not one `refreshed`, `group` or
  `by-id`.** The suspect that started this — the pre-grouping re-read taking the
  tag target — is cleared on the real host. Swapping the target changes nothing;
  the ordering is what has to change.
- A SECOND fault was hiding under the first: 4× `Cannot read properties of
  undefined (reading 'add')` on a `created` handle. `.tags` GONE, not refused.
- Slide identity now reads `unreadable … UNKNOWN (?, ?, ?)` — the false
  `stable (?)` is gone.
- Scratch handback now reads `79 never landed — the host took the add and the
  deck never listed them`.

Instrument built in response: `how-many-syncs-a-creation-handle-survives` (#501).
`tags-on-fresh-shape` (creating batch) and `tag-the-creation-proxy-a-sync-later`
(one sync) both answer yes; production's handles are older because the renderer
chunks a chart across batches. The question ages a handle one sync at a time and
reports where it breaks — a NUMBER the ordering fix can be built against.

### Round 41 — 3c163c4 — archived as 038 — 10/12

Identical `from` distribution to 037. Two rounds agreeing is what the ordering
conclusion now rests on. #502.

### Round 42 — 5a2522e — WEDGED, then the browser died

Wedged in the QUIET form — no crash dialog, so the 30-minute timeout caught it
rather than the crash watch. ~~The lifetime question was never answered.~~
**Wrong, corrected 2026-08-15 14:10 — it was answered twice and the answer sat
in the pane's storage for six hours.** See round 43 below: the pane still had
the run on offer, `Download the crashed run` handed over all 270 steps, and
`how-many-syncs-a-creation-handle-survives` had answered `survives-8` at 50.5s
and again at 67.2s before the wedge. A wedged round is not an empty one, and
what it reached is worth downloading before the next run buries it.

Then the archive step INVENTED A ROUND. `Download run log` is disabled while a
round runs, so the click did nothing and the previous round's file was still at
the same path; it was filed as 039, byte-identical to 038. A checksum caught it.
Deleted, and `archive` now refuses a byte-identical log by name (#503).

**BLOCKED — the browser process is gone and the sign-in with it.** `pw list`
answers `(no browsers)`, no chrome processes. Reopening works, but the fresh
browser has no cookies: onedrive.live.com redirects to login.live.com. Signing in
is the owner's step and mine to stay out of.

TO RESUME: sign in to OneDrive, open Presentation63.pptx, sideload PowerChart
(Home ▸ Add-ins), then `node scripts/round.mjs --check --dir "$OD"` and carry on
from the brief. The loop was stopped rather than left waking into a blocked
state.

Open question still unanswered: `how-many-syncs-a-creation-handle-survives`. It
is merged and will be asked by the first round that runs.

## Resumed 2026-08-15 14:00 — the sign-in was never gone

### Round 43 — 04510e2 — archived as 039 — 10/12

**The lifetime question is answered: `survives-8`, and it is not the
constraint.**

Setup, because the blocker recorded above was not the one that existed: the
browser process was dead, but `--persistent --profile=C:/devtools/pw-profile`
still carried the OneDrive session, so no password was needed. The deck opened
into a THIRD tab and the CLI stays on the tab it opened — `pw tab-list` then
`tab-select 2` is the step that was missing, and without it `--check` reads a
healthy setup as a closed pane. First attempt crashed 2s in; `--retry` cleared
it and attempt 2 landed clean.

1. **Mine.** `how-many-syncs-a-creation-handle-survives` answered `survives-8`
   on all three passes, `stable: true`, and all three were sampled in the
   `collection-refused` regime — the handle took a tag write at every one of
   eight successive syncs while the host was refusing collection reads. Five
   samples now exist across two builds (three here, two recovered from round
   42). **The budget is not the constraint: an UNRESOLVED creation handle does
   not age out.**

   Production failed anyway, and the same way as 037 and 038: every
   `tagging failed` carried `from: created×1` — no `refreshed`, no `group`, no
   `by-id`, three rounds running. The host's own statement dump names the
   mechanism with the ids in it:

       var slide  = slides.getItem("288#2569279682") /* originally getItemOrNullObject(...) */;
       var shapes = slide.shapes;
       var shape  = shapes.getItem("27") /* originally addTextBox(...) */;   ← 5010
       var tags   = shape.tags;

   `27` is the title box, drawn in the first batch and written to in the last —
   the same id the comment at `powerpoint.ts:6956` already quotes from the
   2026-08-07 log. That line is the resolver: `created[k].load("id")` on each
   batch's own sync resolves EVERY created shape including the anchor, so by the
   time the tag is written there is no unresolved handle left. `ungroupedFallback`
   is not the culprit — it slices the anchor off deliberately.

   Rest of the sheet: `tags-add-same-key-twice` still `other` with the honest
   `slide unreadable … UNKNOWN (?, ?, ?)`; `scratch-slides-returned: some` —
   1 of 84 deleted, **83 never landed**, deck 7 slides with 6 added and 0 blank,
   so the counter bug reading holds; 37 draws, ZERO stalls, slowest batch 17.1s,
   one population.

2. **Research.** Nothing new to look up. The 5010 is office-js#2903 (closed, not
   planned) and triage names it inline; no unexplained host behaviour this round.

3. **Instrument.** The three overnight fields all delivered again, and the one
   gap this round exposes is in the probe rather than the pane: the lifetime
   question asks through `ctx.scratch()`, a slide resolved in the batch, while
   production's parent is a slide handle Office has rewritten from
   `getItemOrNullObject(id)` to `getItem(id)`. The probe therefore cannot see a
   refusal that comes from the PARENT path. A partner question that varies only
   the parent would settle it. Not built here — see below.

4. **Fix.** Nothing shipped in the renderer. The ordering change this evidence
   supports is the open BACKLOG item, and the repo's own note on it says it
   "wants a session of its own"; starting it inside a round-mining pass is how a
   wrong theory gets shipped. What the round buys it is the headroom: keep one
   handle unresolved and there are at least eight syncs to spend.

5. **Doctrine.** Round 42's "never answered" corrected above. `docs/BACKLOG.md`
   ordering item now carries the measured budget. This brief's FIRST THING TO
   READ is spent and replaced.

Verdict grid is now 15 rounds: `same scale across the deck` FAIL 15 of 15 —
deterministic, only the degree varies (4 of 8 charts carried the shared scale
here, host flipped at chart 5 of 8). `explode a degraded picture` is the one
scenario that has ever flipped, and its single pass is still 1 in 15.

Rasterise arms: 30 draws per arm pooled over 15 rounds, both 0.0%. Still short
of the 60-100 an arm it needs.

Recovered round-42 log kept at `.pw-session/crashed-5a2522e.json` — not filed
under `rounds/`, since a wedged partial has no self-test and would read as a
16th round in the ledger.

### Round 044 — b6582c7 — two attempts lost to the DRIVER, not the host

Both attempts died identically, and the message was false:

    PowerPoint crashed 2s in — its dialog is up …
    clearing the crash and starting again

    attempt 2 of 7
    HEAD b6582c7 · site b6582c7 · pane b6582c7 · deck 1 slide(s)
    host answered in 6ms
    NOT READY — playwright-cli could not be run — nothing below was actually measured.

Every value on that sheet was measured, by the tool it says could not be run,
one line above the refusal. Exit code 0 while doing it.

**Root cause, reproduced away from the round rather than guessed at.** Running
`collectCrashEvidence` directly against the live browser printed it in one pass:

    spawn errors: 2
      - requests => spawnSync …\node.exe ENOBUFS

`sh("requests")` on a live PowerPoint tab is bigger than `spawnSync`'s 1 MiB
default `maxBuffer`. Over the line Node throws the output away and returns an
`error`, and `cli` read any `error` as "the tool could not be run".

Three defects, all fixed here, each with a mutation-proven guard:

1. **The buffer.** `maxBuffer: 64e6`. The default is invisible — nothing about
   the call says 1 MiB, and the failure names the executable rather than the size.
2. **An overflow is not an unreachable tool.** They are opposite facts: one means
   the browser answered and the answer did not fit, the other means nobody asked.
   `isOverflow` splits them.
3. **The doubt outlived its sweep.** `unreachable` is deliberately sticky within
   one readiness sweep and was sticky for the whole PROCESS, so a failure during
   attempt 1's crash handling refused attempt 2 forever. `sh.startSweep()` at the
   top of each attempt. This is the round-29 poll bug — "do not end a round on one
   failed CLI call" — one call site further out, unswept.

**And the damage was worse than two lost attempts: it emptied both crash
reports.** They read

    ## Document channel, last 12
    (nothing)
    ## PowerPoint's own error log (ULS)
    (no fatal entry in the last 0 telemetry batches — the quiet form of the wedge
    looks like this)

Nothing was measured, and the second sentence names a diagnosis. Round 043's
report from the same crash carried twelve real channel entries and a ULS window.
`collectCrashEvidence` now reports a failed read AS a failed read; verified
against the real browser, the report went 1834 → 3167 characters.

**Why 043 survived this and 044 could not**, which is the part worth carrying:
the overflow is a function of TAB AGE. Round 043's crash was collected at request
768 of a freshly opened tab and fitted; 044's third attempt crashed at request
**16748** of a tab that had been alive for two hours. So the bug only bites a
long-lived tab — which is to say, only during a round LOOP, which is the only
thing this driver is for. A defect that cannot reproduce on the first run of the
day is worth naming as such.

### Round 044, third attempt — b6582c7 — archived as 040 — 10/12

Attempt 1 crashed 2s in as usual; attempt 2 reached `ready` and ran, which is the
driver fix above doing its job on its first outing.

1. **Mine.** Two things, and the second is the one that closes an argument.

   **`survives-8` replicated a third time** — three passes here, `stable: true`,
   all in the `collection-refused` regime, on a third build. Eight samples across
   three builds now say a creation handle nobody resolved does not age out.

   **`tag-through-refetched-shape` is NEW and answers `no-id`, stable ×3** —
   "the fresh shape would not report an id". Taken with the above, the last
   alternative to the ordering change is gone: you cannot re-fetch the tag target
   by id, because on this host a fresh shape has no id to re-fetch it BY. Not
   "the id is refused" — there is no id. Route by route: `refreshed` is resolved
   and refused, `by-id` cannot be built, `created` works and is the one the pass
   throws away. The ordering change is not the best option left, it is the only
   one.

   Otherwise identical to 039: 10 of 12, the same two failures, `same scale` 4 of
   8 with the host flipping at chart 5, 37 draws and ZERO stalls, slowest batch
   18.4s, scratch handback 82 of 83 never landed.

2. **Research.** None needed; the one unexplained thing this round is the 2s
   crash, and the host now explains it itself — see below.

3. **Instrument.** The crash report is the instrument, and it worked the moment
   it was fixed. From `crashes/2026-08-15T12-49-04.md`, PowerPoint's own log:

       319 | In OnDisconnect(), setting SlideViewNode.srcSlide to null
       321 | Failed to restore selection after load content.
       389 | OnServerFindSucceeded could not find target slide, time elapsed: 430 ms
       390 | GlobalErrorHandler:DisplayErrorDialog: 5341289

   **PowerPoint has now crashed 2s into the FIRST attempt four times running**
   (043, and all three 044 attempts) and never into an attempt that followed a
   recovery. The difference between them is ~80 seconds of reload and settling.
   The host's own account is that it could not find the target slide after
   restoring content — a document still settling when the round's first Office.js
   call lands. That is a theory with four samples and no test.

   **Built, and it is an experiment rather than a fix.** `--check` pinged the
   host but the ping touches no slide — `getCount` is a count, and this host
   answers it in single-digit ms while in exactly the state it then dies from.
   The pre-flight now resolves slide 1 and the check line carries `slide 1
   resolved` or `slide 1 REFUSED` every round, pass or fail, because the rounds
   where it says `resolved` are what make the others mean anything. The point is
   to MOVE the crash, not prevent it: trip it in a two-second check that
   `--retry` recovers from, instead of an attempt plus 80s. If it answers
   `resolved` and the round crashes anyway, the theory is refuted for the price
   of one call — which is the outcome worth having either way.

   The crash dialog is now re-read AFTER the slide touch as well as before it. A
   crash the check itself provoked would otherwise be carried into the round as
   `ready`.

4. **Fix.** The three driver defects above.

5. **Doctrine.** `docs/BACKLOG.md` ordering item records the closed route.

### Round 045 — ca866e3 — archived as 041 — 10/12, and the first clean start in five

    host answered in 4ms · slide 1 resolved
    ready
    running — this takes about ten minutes
    finished

**No crash, on the first attempt.** Rounds 043 and 044 crashed 2s into every
first attempt — four for four — and this is the first that did not.

**What that is worth, stated before it gets quoted as more.** n=1, and the
experiment is confounded BY DESIGN: the check resolves a slide, so a clean run is
equally consistent with "the touch settled the host" and with "the host was fine
today". The reading that would have been decisive is the other one — `resolved`
followed by a crash anyway would have refuted the theory outright — and it did
not happen. So this is one point of weak support for a theory that still has no
test, and the brief already says to record the word every round for exactly this
reason. Three or four more clean starts make it an effect; one does not.

1. **Mine.** Everything else is the fourth consecutive repeat, which is itself
   the finding — this host is now boringly reproducible:
   - `survives-8`, stable, on a fourth build. Eleven samples.
   - `tag-through-refetched-shape: no-id`, stable, replicated.
   - Every tag failure `from: created×1`. Four rounds, not one `refreshed`,
     `group` or `by-id`.
   - `same scale` 4 of 8, host flips at chart 5 of 8. 16 of 16 failures.
   - 37 draws, ZERO stalls, slowest batch 16.9s, one population.

   **Two things that are NOT repeats, both in the deck rather than the trace:**
   - **One added slide came back blank and the rasterise agrees** (7 added, 6
     carry shapes). Rounds 039 and 040 both had zero. One witness plus its
     picture is the standard this repo set for calling a slide empty, and this
     one meets it.
   - **Two of the seven new slide ids are `256#0` and `257#0`** against
     `288#3603562595` and friends for the rest. An id whose second half is `0` is
     not the shape this host gives a slide it has finished adding, and the round
     also reports `delete-by-id left slides behind and this host does not list the
     ids they were added under`. Those two facts are probably the same fact.

2. **Research.** Nothing unexplained beyond the `#0` ids, which are worth a
   search of the office-js tracker next round rather than a guess now.

3. **Instrument.** The next question is already framed: does a slide whose id
   reads `NNN#0` correspond to the blank one? The round log carries both halves —
   `deck.inventory` and `deck.newSlides` — but nothing joins them, so it took a
   hand query to notice. Joining them in `triage` is the cheap next instrument.

4. **Fix.** None. Nothing this round exposed is fixable without first knowing
   whether the `#0` id and the blank slide are the same event.

5. **Doctrine.** The brief's "first thing to read" now points at the experiment's
   own word, and this entry records what a single clean start does and does not
   license.

### Round 046 — a54401c — archived as 042 — 10/12, and it settled the ordering fix

`slide 1 resolved`, clean first attempt again — two in a row now, still not an
effect.

Getting there needed a hand recovery worth recording: `pwclean` answered
`deck-failed` and `--check` said why — the QUIET wedge, host silent for 8016ms
with no dialog, plus a pane still on `ca866e3` and an eight-slide deck. One
reload cleared all three. That is `recover`'s own sequence, done by hand, because
the driver would not do it — see the fix below.

1. **Mine. `collection-read-poisons-the-creation-handle` answers `yes`** — three
   passes, `stable: true`, every one taken while the host was refusing collection
   reads. **The pre-grouping re-read does NOT poison the creation handle.** The
   fake was wrong, the ordering fix is viable, and it was rebuilt the same hour.

   Otherwise the fifth consecutive repeat: `survives-8`, `tag-through-refetched-shape:
   no-id`, every tag failure `from: created×1`, `same scale` 4 of 8 flipping at
   chart 5, 30 draws with zero stalls. One added slide blank again, and the same
   two `#0` ids as round 041 — `256#0` and `257#0`, identical strings across two
   rounds, so they are not per-round noise.

2. **Research.** None owed; the one open question was put to the host and
   answered.

3. **Instrument.** `triage` now names added slides whose id ends `#0` and joins
   them to the confirmed blanks, because round 041's finding took a hand query
   and nearly went unseen. **It earned itself immediately, in the negative
   direction: 0 of the odd-id slides are the blank one, in both rounds.** The
   coincidence is dead and nobody spends a round on it.

4. **Fix.** Two, neither of them the renderer:
   - The driver stopped on states it knows how to fix. `shouldRetry` retried a
     crash and nothing else, so the quiet wedge (`silent`) ended a round and a
     `not-ready` was never retried at all — which is why the recovery above was
     done by hand. Every readiness stop now carries a CODE, and a check is
     retried when recovery addresses all of them and refused when even one stop
     it cannot touch is present (`site-behind`, `verbose-off`).
   - `origin tag lost` is its own trace line. Losing it costs drag tracking, not
     re-editability, and reporting it as "charts are not re-editable until
     repaired" was about to become the common case.

5. **Doctrine.** `docs/BACKLOG.md` records the answer and what it licenses.

### Round 047 — d2ca1c9 — archived as 043 — 10/12, and it does NOT validate the fix

The first round carrying the tag-anchor change, and the honest answer is that one
round cannot tell. Recorded before anything else because the temptation to read
it as a win was real.

**The driver fix worked, visibly.** Attempt 1 refused on `deck-dirty` alone,
retried, `recover` cleaned the deck 8 → 1, attempt 2 ran and finished. Under the
old rule that was a dead stop, and the two rounds before it were hand-recovered.
(Wart: it still prints "clearing the crash and starting again" when nothing
crashed. The message predates the change.)

**The renderer fix changed nothing measurable, and here is why that is not a
verdict.** Fault counts, three consecutive rounds:

    round  build     tags-undefined  cfg-tag-5010  group-5010  no-queue  tagging-failed
    041    ca866e3        1               6            1          1           4
    042    a54401c        5               6            5          5           8
    043    d2ca1c9        5               6            5          5           8

**042 and 043 are identical, and 041 → 042 is a five-fold jump with NO renderer
change between them** — a54401c added a probe and documentation and nothing else.
So the counts are dominated by the host's regime, they swing by 5× on their own,
and a single round comparing two builds is measuring mood. `same scale` moving
from 4-of-8 to 3-of-8 sits inside that same noise.

What IS informative: the config write still fails six times, exactly as before.
If the anchor were reaching the write unresolved, that number should have moved.
And `origin tag lost` — the line that exists for the case where the config lands
and only the origin fails — fired ZERO times, which says the same thing from the
other side.

**So the fix is unvalidated, not refuted, and the next hypothesis is better than
the last one.** `addGroup` fails 5× with 5010 in the same pass, and a failed sync
poisons its own context — this project already knows that and works around it
elsewhere. `no chart's tag could be queued` firing 5× is a context-level symptom,
not a handle-level one: the QUEUE is refusing, before any handle is exercised. If
the grouping attempt poisons the context the tag write sits in, then no choice of
handle can help and the whole ordering effort has been aimed one level too low.

That is testable and cheap: tag in a context that has not just tried to group.

**Instrument owed, and it is the lesson of this entry:** the rasterise arms are
pooled across rounds because one round could never answer them, and the tag
failure counts need exactly the same treatment. Comparing two builds by eye in a
regime that swings 5× is how a fix gets declared good. `npm run rounds` should
pool tag failures per build the way it already pools draws per arm.

Probes: `which-end-a-short-read-drops` answered `unreadable` — "the collection
would not list its items", which is this host's usual answer to anything asking a
collection to enumerate itself. The trade it prices stays unpriced; the question
is right and the host is not answering it yet.

### Round 048 — 355a37c — archived as 044 — 10/12, and the question missed

**The driver recovered the whole three-stop state on its own**, which is what
this morning cost three hand-recoveries:

    NOT READY — a silent host, a pane on d2ca1c9 while the site served 355a37c,
                and a seven-slide deck
    recovering from a silent host and a stale pane and a dirty deck, then starting again
    attempt 2 of 7 — ready

Both fixes from the last two commits doing their job on their first outing, and
the recovery line naming all three rather than announcing a crash.

1. **Mine. `does-a-failed-group-poison-the-tag` answered `no-refusal`, three
   times** — "the host grouped, so the question was never put". The probe was
   honest about missing rather than inventing an answer, and the miss is itself
   the finding: **this host groups two fresh shapes from one batch perfectly
   well.** What it refuses is PRODUCTION's grouping, whose members and slide
   handle are several syncs old by the time `addGroup` is called. Grouping is not
   refused as such; it is refused for aged handles.

   The rest is the sixth consecutive repeat, and the new table says so at a
   glance — `355a37c` scores 5/6/5/5/8, identical to `a54401c` and `d2ca1c9`.
   Three builds, one of them carrying the tag-anchor change, all indistinguishable
   against a noise floor of 1-vs-5 measured within a single build.

2. **Research.** None owed.

3. **Instrument.** The question now ages its handles two syncs before grouping,
   which is what it should have done first time. **And it comes with a warning
   that the fake found:** under `strictTags` the aged version answers
   `refused-after-group` for a reason that has nothing to do with grouping — the
   age rule refuses the write on its own. On a real host the same word could mean
   either thing, so it must be read beside `tag-the-creation-proxy-a-sync-later`,
   which answers `yes` here and excludes age. A question that can be right for
   the wrong reason is worth shipping only with the control named next to it.

4. **Fix.** None owed; the renderer is where it was.

5. **Doctrine.** `PENDING_QUESTIONS` carries the miss, the reason for the ageing,
   and the read-it-with-the-control rule.

## The long run — 2026-08-15 evening

### Pair 1 — 56eb477 × 2 — archived as 045 and 046

**THE FIRST SAME-BUILD PAIR, and it settles the tag-anchor question.** Both
rounds scored `tags-undefined 5 · cfg-tag-5010 6 · group-5010 5 · no-queue 5 ·
tagging-failed 8` — **identical, zero spread.** Four consecutive builds now score
the same, and one of them predates `tagAnchorIndex`:

    a54401c  5 6 5 5 8   ← before the anchor change
    d2ca1c9  5 6 5 5 8   ← anchor change lands
    355a37c  5 6 5 5 8
    56eb477  5 6 5 5 8   (twice, no spread)

Five rounds, four builds, no within-build variance on the current one. **The
anchor change did nothing.** That is not the ambiguous single-round reading of
043; it is a measurement with its own noise floor attached. Whether to revert it
is the owner's — it is merged renderer code.

1. **Mine — and the finding did not need a round at all.** Joining each chart's
   grouping outcome to its tag outcome, across the whole archive:

       grouped      64 chart(s),  1 lost the tag =  2%
       NOT grouped  62 chart(s), 41 lost the tag = 66%

   Per round it is nearly mechanical: three grouped and none lost, two or three
   ungrouped and two lost, eleven rounds running. **A chart that gets grouped
   keeps its config; one that cannot loses it two times in three.** When grouping
   succeeds the tag goes onto the GROUP — a handle made in that batch, never
   resolved — and it lands. When grouping is skipped the tag falls back to a
   `created` handle and is refused.

   So the entire handle question — four rounds, a merged renderer change, a fake
   corrected, three probes — has been about the LOSING path. It was aimed one
   level too low, and the evidence had been in the archive for eleven rounds
   unqueried. `npm run rounds` prints it every round now.

   `not grouping: no member handle this host will accept` carries `refreshed: 0`
   every time, so what decides a chart's config is whether the PRE-GROUPING
   RE-READ returned anything. That is where the next attempt belongs.

2. **Research.** None owed; the archive answered it.

3. **Instrument.** The three lines that decide a chart's fate now carry the
   slide (`not grouping`, `a chart's tag could not even be queued`, `tagging
   failed`). Rounds 043-045 each lost a chart's config and the one line that
   names a slide named `257#0` every time — an id whose second half is `0`, which
   is not the shape this host gives a slide it has finished adding. Next round
   says outright whether the charts that lose their config are the ones on those
   slides. If they are, this is one SLIDE fault rather than three tag faults.

4. **Fix — four, and the run exposed every one of them.**
   - **A mid-round wedge was not retryable.** Round B timed out at thirty
     minutes and the driver stopped; in an unattended run that is nine idle
     hours. `timeout` now recovers like a crash, bounded by `--retry`.
   - **Crash reports were STILL empty**, hours after the overflow fix, and this
     one was mine: `ask()` honoured `lastError` (spawn, overflow) but not
     `lastFailed` (ran, exited non-zero) — which is exactly how a wedged tab
     fails. Three sections read "(nothing)" over three failed reads.
   - **A dead browser read as a closed pane.** The wedge killed the browser
     process; `recover` then reloaded and reopened a pane inside a window that
     did not exist, SEVEN times, while the check asked "is the add-in open?"
     `noBrowser` names it, and recovery now reopens from the persistent profile
     — no password, because a dead browser is not a lost sign-in.
   - **`npm run rounds` reported the oldest round** in the archive rather than
     the newest.

5. **Doctrine.** `docs/BACKLOG.md` now LEADS with grouping-saves-the-config, and
   says in as many words that the three sections under it are about the losing
   path.

**Cost of the pair:** one wedge, one dead browser, seventy-one slides of litter
to clean. All four fixes above came out of it, which is the trade this loop
exists to make.

### Pair 2 — 5856d1d × 2 — archived as 047 and 048

**The `NNN#0` slide lead is dead, and it died from the archive rather than from
the instrument built for it.** The draw trace already carried `chart` and
`onSlideKey`, so the join was available on rounds already filed:

    1/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    2/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    3/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    4/8  slide 257#0            no group  CONFIG LOST    <- a #0 slide
    5/8  slide 288#1168146411   no group  CONFIG LOST    <- an ordinary id

Identical in 043-047. A `#0` slide carries the three best-behaved charts AND a
failure; an ordinary slide fails too. **The id shape predicts nothing**, and the
instrument shipped this build confirmed it independently: every failure line in
047 named `291#567483212` and `288#230020768`, both ordinary.

**What does predict it is POSITION IN THE DECK-WIDE UPDATE**, and six rounds now
show the same decay curve rather than a coin:

    charts 1-3   re-read matches all 24     grouped, config kept
    chart 4      re-read matches 20 of 24   partial, thrown away, config lost
    chart 5      re-read returns NOTHING    config lost
    charts 6-8   never attempted

`same scale across the deck` says it itself — "the host flipped at chart 4 of 8"
— seventeen rounds running.

1. **Mine.** Grouping-saves-the-config strengthened with the pair: **73 grouped,
   1 lost (1%); 68 ungrouped, 47 lost (69%)**.
2. **Research.** None owed.
3. **Instrument.** `how-many-collection-reads-a-context-survives` — how many
   times one context can re-read a slide's shapes before the answer goes short.
   That is the mechanism the decay curve implies, asked directly.
4. **Fix.** A comment, and it was asserting something false: the partial-match
   branch claimed an ungrouped chart "is still tagged, still re-editable". The
   69% above refutes it. **The branch itself is deliberately left alone** — the
   alternative strands shapes, and choosing between two harms is not a call to
   make from a trace.
5. **Doctrine.** `docs/BACKLOG.md` carries the refuted lead, the decay curve, and
   the experiment with its prediction.

**The suspect is our own perf work.** `updateChartsInSlides` was deliberately made
ONE context, four syncs, flat in N — the right shape for a host that can hold a
context, and this one appears to degrade as a context is used. `#112` already made
the opposite call for the demo deck. **Not acted on**: the fix is a 390-line
restructure of the live update path, and this project has three shipped-broken
fixes on record from changing that path on a theory. The probe decides it for the
price of one round.

### Pair 3 — 3f04df2 × 2 — archived as 049 and 050

**`how-many-collection-reads-a-context-survives` answered `unreadable-at-1`**,
three samples, stable: the scratch slide would not enumerate its collection even
ONCE, so there was never a baseline to degrade from. The question was built to
decide the strongest lead this project has, and it could not be put.

**That is the third question to die on the same harness limit** —
`shapes-items-count-honest` and `which-end-a-short-read-drops` are the other two,
both `unreadable` — and three is a pattern rather than three misses.
**The finding is about the harness, not the host: this scratch slide is strictly
worse at collection reads than a real one**, whose collection enumerates fine for
the first three charts of a deck update, every round. Collection questions cannot
be asked from a probe here. They have to be instrumented in production.

1. **Mine.** Grouping-saves-the-config, pooled over 26 rounds: **79 grouped, 1
   lost (1%); 72 ungrouped, 51 lost (71%)**. Six rounds tonight and the effect
   has only sharpened.
2. **Research.** None owed.
3. **Instrument — moved from the probe into the production path.** `boundedSync`
   now counts syncs per context (a WeakMap, so the count dies with the context),
   and the two re-read failure lines carry `contextSyncs`. The deck-wide update
   runs every chart through ONE context and the re-read decays chart by chart, so
   this is the x-axis the decay curve has never had. Next round says at which
   sync of a context the re-read starts going short — which is exactly what the
   probe could not ask.
4. **Fix.** None owed; the renderer is where it was, deliberately.
5. **Doctrine.** `PENDING_QUESTIONS` records the harness limit and tells the next
   reader not to wait on that question.

**A repair worth recording, because it was mine and the tables caught it.** The
patch that wrote the finding into `PENDING_QUESTIONS` hit `FAKE_BASELINE`
instead, replacing an answer with prose and inventing a
`…-ORIGINAL` key for the value it displaced. `host-contract.test.ts` failed on a
baseline entry for a question that does not exist — which is precisely the check
that exists so a table cannot quietly drift from the probes it describes.

### Pair 4 — 1a1e05f × 2 — archived as 051 and 052 — THE ROOT

**`contextSyncs` answered on its first outing, and it refuted the hypothesis it
was built for.** Both failing charts show `contextSyncs=1`: the re-read that goes
short is the FIRST sync of its context, not the thirtieth. The context is fresh.
**Chunking `updateChartsInSlides` would change nothing**, and that 390-line
restructure is ruled out before anyone started it — which is exactly what the
instrument was for.

**Then the same trace said what DOES separate them**, and it is a switch:

    chart 1/8  slide 256#0           32 shapes already on it  -> re-read OK, GROUPED
    chart 2/8  slide 256#0           40 shapes already on it  -> re-read OK, GROUPED
    chart 3/8  slide 256#0           40 shapes already on it  -> re-read OK, GROUPED
    chart 4/8  slide 257#0            0 shapes already on it  -> re-read SHORT 20/24
    chart 5/8  slide 288#807648421    0 shapes already on it  -> re-read EMPTY

Pooled over the whole archive:

    slide already had shapes   82 chart(s), 81 grouped = 99%
    freshly added, empty       74 chart(s),  1 grouped =  1%

**A chart drawn onto a slide this run has just added does not get grouped, and a
chart that is not grouped loses its config.** That is the root, and everything
this project has done to the tag path for four rounds sits one level above it.

It is also not new. `shape-add-held-slide-proxy` answers `threw`, a web-new-slide
id does not round-trip, and the #108-#111 saga was four attempts at drawing on a
freshly added slide. The evidence has been in every round since; nobody had
joined `onSlide` to the grouping outcome.

1. **Mine.** As above. `256#0` — the id flagged as malformed two pairs ago — is
   the slide that WORKS, because it is the established one.
2. **Research.** None owed; the archive answered it.
3. **Instrument.** `npm run rounds` prints **WHICH SLIDE THE CHART LANDED ON**
   every round now.
4. **Fix.** None shipped, deliberately. The matcher already knows the honest rule
   — "the positional rule is still right for a slide this run added blank" — but
   that branch is reachable only when NOTHING matched, so a chart matching 20 of
   24 falls past it. On a slide this run added blank we KNOW our shapes are the
   only ones there, so a short read is a host lie we can detect rather than obey.
   Contained, but still surgery on the grouping path, which carries three
   shipped-broken fixes on record. It wants a person awake.
5. **Doctrine.** `docs/BACKLOG.md` now leads with the root and lists what was
   ruled out: context wear, the `#0` ids, and the poisoned-context theory.

**Two of tonight's driver fixes proved themselves in this pair.** Round B wedged
at thirty minutes and the driver recovered and finished on attempt 3 — before
tonight that was a dead stop. And the crash report from that wedge reads "(could
not read — the call could not be run)" three times rather than "(nothing)",
which is the `lastFailed` fix refusing to report an empty host as evidence.

### Pair 5 — a004711 × 2 — archived as 053 and 054 — a question CLOSES

**`does a rasterise poison the next draw` is answered: NO.**

    after a rasterise     0 stalled /  60 drawn = 0.0%
    after a cheap read    0 stalled /  60 drawn = 0.0%

Sixty draws per arm is the bar the tool itself set, and the "NOT an answer yet"
line is gone. Open since round 33, counterbalanced the whole way, and it closes
in the negative — a rasterise is not what makes a draw stall. Nothing to fix,
which is the point: the suspicion is retired rather than carried.

The root cause holds and sharpens with two more rounds:

    slide already had shapes   91 chart(s), 90 grouped = 99%
    freshly added, empty       80 chart(s),  1 grouped =  1%

**A slip of mine worth recording.** Round B was launched with a shell `&`
instead of the harness's background mechanism, so its console output was lost and
no completion signal existed. The process survived and the round completed — the
run log and the archive are what matter, and both were intact — but the recovery
was luck rather than design. A waiter that watches for the process to exit
restored the signal without touching playwright-cli. Launch a round the same way
every time.

### Pair 6 — 61f431b × 2 — archived as 055 and 056

Accumulation, and the two things worth writing down are both about what has
STOPPED happening.

**The 2s crash has not recurred since the pre-flight shipped.** Four crash
reports say "2s into a round" and all four are from a forty-five-minute window
before `slideResolveScript` existed. In the twelve rounds since, `slide 1
resolved` has printed on every check and not one round has crashed 2s in:

    before the pre-flight   4 first attempts, 4 crashed
    after                  12 rounds,        0 crashed

**Observational, not counterbalanced**, and it must be read that way: the
pre-flight resolves a slide, which may settle the host rather than merely
measure it, and the host's mood is not controlled. But four-in-four against
zero-in-twelve is the strongest thing this experiment can say without an arm that
deliberately skips the touch — and skipping it would cost a round to learn
something the loop already gets for free.

**The root cause is unmoved by more data**, which is what a real effect does:

    slide already had shapes   97 chart(s), 96 grouped = 99%
    freshly added, empty       84 chart(s),  1 grouped =  1%

Nothing else new. `same scale` failed the same way, the same charts, in the same
order — which after twelve rounds is itself the finding: this host is
deterministic here, and the remaining work is a decision rather than a
measurement.

### Pair 7 — 2a11873 × 2 — archived as 057 and 058

Round B wedged at thirty minutes and the driver recovered and finished on attempt
3 — the second time tonight that fix has saved a round which, this morning, would
have been a dead stop.

    slide already had shapes  103 chart(s), 102 grouped = 99%
    freshly added, empty       88 chart(s),   1 grouped =  1%

Fourteen rounds and the ratio has not moved. Nothing new from the host.

**The work this pair actually bought was in the documents, not the rounds.**

- **`chooseGroupMembers`'s contract carried the same refuted premise** the
  partial-match branch did — "strictly worse than an ungrouped chart that is
  still re-editable". The conclusion survives (a throw costs every chart in the
  batch, grouping nothing costs one) but the premise does not, so it now says
  plainly that this is a choice between two losses and that the way out is
  upstream. Two other candidates in `PUBLISHING.md` were checked and are sound —
  both describe charts the rescue actually grouped.
- **The brief's queue was stale in all five items**, every one answered during
  this run, and item 5 said something now false: "it moves when the tag path
  moves". It moves when the fresh-slide re-read is fixed. Rewritten around what
  is settled, with nine questions listed as answered and their results, so the
  next session cannot spend a round re-deriving any of them.

**And the honest note at the top of it:** the mechanism is settled and more
rounds do not add to it. What is left is two decisions, both the owner's. A loop
that keeps running past the point where measurement helps is manufacturing
activity, and saying so is worth more than another pair.

### Pair 8 — 85277f5 × 2 — archived as 059 and 060

    slide already had shapes  109 chart(s), 108 grouped = 99%
    freshly added, empty       92 chart(s),   1 grouped =  1%

Sixteen rounds, unmoved. Both rounds clean, no wedge, nothing new from the host.

**Doc drift caused by this run's own fixes, found and corrected:**

- **Three rotting counts.** `docs/BACKLOG.md` (twice) and `host-probe.ts` asserted
  `same scale` had failed "seventeen rounds running". It is 34 of 34 now, and
  would be wrong again tomorrow — so they say "every round it has run" and point
  at `npm run rounds` for the number. A count frozen in prose is a comment with a
  half-life.
- **`ROUNDS.md` described a driver that no longer exists.** Its `--check` list was
  missing the slide-resolve and no-browser refusals, and said nothing about the
  eight states `--retry` now recovers on its own — including two it deliberately
  does not, which is the more useful half. Rewritten against
  `RECOVERABLE_STOPS`, which is derived from what `recover` actually does.

**The audit is out-yielding the rounds, and that is the honest read of where this
is.** Two refuted premises and four stale statements in three pairs, against
sixteen rounds that agree with each other. The host has said what it has to say
about this mechanism; the documents had not caught up.

### Pair 9 — 9dc3bc8 × 2 — archived as 061 and 062

    slide already had shapes  115 chart(s), 114 grouped = 99%
    freshly added, empty       96 chart(s),   1 grouped =  1%

Eighteen rounds. Both clean, no wedge, nothing new.

**`CLAUDE.md` was the last document pointing at the wrong problem.** It said the
red here is "the shape collection dying part-way through", which sends a reader
looking at POSITION — and position is what the archive refuted. It now carries
the split above and names the freshly-added slide, with the note that four rounds
of tag-path work aimed one level above it before anyone joined `onSlide` to the
grouping outcome. That matters more than another round: `CLAUDE.md` is what a
fresh session and the parallel stream read first.

Four documents now agree with the measurement — `CLAUDE.md`, `docs/BACKLOG.md`,
this brief and `docs/ROUNDS.md` — where at the start of the night three of them
described the tag writer as the problem.

### Pair 10 — e68147f — round A archived as 063; round B stopped on SIGN-IN

**The run ends here, on the one state that is the owner's.** Round B wedged at
thirty minutes, the driver recovered as it now does, and attempt 3 found a
`login.live.com` popup open beside the deck — the session's token had expired.
The check named it exactly:

    the browser is on a Microsoft sign-in page — there is no document to run a
    round in. … None of that is the agent's to do: it needs a password.

Nothing was touched: the popup is left exactly as Office opened it, because
dismissing an authentication prompt is not a workaround and the owner may want to
complete that prompt as it stands. **The readiness work paid for itself one last
time** — this stopped in one attempt with the right diagnosis, where the same
class of state cost seven attempts and an "is the add-in open?" earlier tonight.

**A precision note for whoever reads this next:** `signedOut` fires on any tab
matching `login.live.com`, and what actually happened was a POPUP beside a deck
tab that was still open. Treating that as signed-out is the right call — if
Office is asking for credentials, a round cannot be trusted — but the message
says "the browser is on a sign-in page", which is not quite what a reader will
see on screen.

**Coverage the audit found, and it was mine.** Two of tonight's own fixes
shipped without guards, on paths a green suite could not have covered because
nothing tested them at all:

- The `lastFailed` half of the crash-forensics fix — verified in production on
  the 18:06 wedge and never pinned.
- `syncsOf` / `contextSyncs` — the number that ruled out a 390-line restructure.

Both guarded and mutation-proven now. That is the third class of defect this run
produced, after round findings and doc drift: **fixes of mine, unguarded.** Worth
checking for deliberately, because the suite was green before I looked.

## Round 064 (`bcd5773`, 2026-08-16) — the retry works, and it exposed the next link

**The prediction was staked before the round and it held.** Charts 4 and 5 of
`same scale across the deck` — the two that had failed identically for five
rounds running, one on a short re-read and one on an empty one — **both grouped
for the first time.** The trace is causal rather than statistical, which is why
one round is enough to believe this much of it:

    4/8   re-reading the slide's shapes again after a settle delay  waitedMs=1500
    4/8   grouped the chart's shapes                                by=ids
    5/8   re-reading the slide's shapes again after a settle delay  waitedMs=1500
    5/8   grouped the chart's shapes                                by=ids

`the settled retry repaired 2`, on the scenario's own verdict line. **No chart in
the round went ungrouped** — `NOT grouped 0`, against a pooled 98 charts at 79%
tag loss. Freshly-added slides went 2 of 2 grouped where the archive had 1 of 98.

**AND THE SCENARIO STILL FAILED**, 4 of 8. The chain did not break; it moved one
link along, and the new link could never have been seen before, because these
charts had never once got as far as having a group:

    4/8   tagging failed   from: group×1   slide 258#4111159134   5010
    5/8   tagging failed   from: group×1   slide 259#3844610554   5010

**THE DOCTRINE THIS CORRECTS — read it before quoting the 2% again.** This file
and `docs/BACKLOG.md` have said "grouping is what saves a config" on the strength
of:

    grouped      123 chart(s),   3 lost the tag = 2%
    NOT grouped   98 chart(s),  77 lost the tag = 79%

That number was measured on a population that **excluded freshly-added slides by
construction** — a fresh-slide chart could not group, so it could never appear in
the grouped column. The moment one does, it loses its tag anyway: this round's
own split is `grouped 5, 2 lost = 40%`. The rule was never "a group saves the
config"; it was "a slide that was already there saves the config", and grouping
was standing in for it.

**Why the group handle is refused too, and it is already written down from the
other side.** A shape proxy carries its PARENT's object path. The group is made
in the grouping batch, but its slide handle is by then rewritten to
`slides.getItem(id)` — and a freshly-added slide's id does not round-trip on this
host (`shape-add-held-slide-proxy: threw`, the #108-#111 saga). The round says so
in as many words: *the host grouped through a slide handle two syncs old, so the
refusal was never provoked.* The members were never too old. The PARENT was.

**So the next question is the slide handle, not the shape handle** — and unlike
every previous "next question" here, it is not a guess: the failure names the
slide, both times, and both are freshly added.

Two other tag failures in the round, neither new and neither on this path:
`from: created×1` on slide `260#1686107471` beside an empty settle re-read, and
five × `Cannot read properties of undefined (reading 'add')` on `262#3659873566`
— the documented `target.tags` guard.

**Not yet a pair.** The causal lines (retry → group, same chart) are strong
enough to stand alone; the counts are not, and `40%` in particular is five
charts. A same-build second round is what makes the tag-loss number mean
anything.

## Round 065 (`bcd5773`, 2026-08-16) — the pair, and it replicates exactly

Second round on the same build, run without a merge between the two, which is
what makes it a pair rather than two sheets. **It is structurally identical to
064, chart for chart:**

    1-3/8   grouped by=ids                             no retry needed
    4/8     settle delay 1500ms → grouped by=ids        tagging failed  from: group×1
    5/8     settle delay 1500ms → grouped by=ids        tagging failed  from: group×1
    260     from: created×1, beside an empty settle re-read
    262     5× no chart's tag could be queued

`the settled retry repaired 2` in **both** rounds. The slide ids differ between
them (`258#4111159134` vs `258#2150477121`), so what repeats is the POSITION and
the freshness, not an id.

**This is the strongest replication this project has.** The noise floor — 1
versus 5 for the same fault with nothing changed — is about counts. These are not
counts: the same two charts take the same retry, group, and are refused through
the same handle kind, twice. Nothing here needs a third round.

**What still moved between the two, and it is the downstream half:** `same scale`
scored 4 of 8 then 3 of 8, and re-editable went 7 then 6. That is the part that
was always noisy and it stays noisy. The mechanism is deterministic; the score is
not, and reading the score as the result is the mistake this pair exists to
prevent.

**So both conclusions from 064 stand on two rounds now:**

1. The settled retry works, deterministically, on exactly the charts it was built
   for.
2. `from: group×1` is refused on a freshly-added slide — so **grouping does not
   save a config**, and the old 2%-vs-79% split was an artifact of a population
   that could not contain these charts.

**The next thing to build is the slide handle**, and it now has two witnesses
naming it rather than one.

### A numbering trap, found by walking into it

The second round archived as `064` as well. **Nothing was overwritten and no
evidence was lost** — `nextRoundNumber` is max+1, so it can never land on a file
it can see. What happened is subtler: round 064's file was committed to a branch,
`git checkout main` removed it from the working tree, and archiving from `main` —
where the directory ends at 063 — reissued the number.

The damage is not destruction, it is two different rounds both called 064 in two
git contexts: a merge collision, and a pooled report that reads one of them twice
or not at all. Every number in `npm run rounds` is a pool over that directory.

`--archive` now numbers from the union of the working tree and `git log --all`,
so a round committed on an unmerged branch still holds its number. Falling back
to the directory alone when git is unavailable, because refusing to archive a
real round over a numbering nicety is the worse trade.

**Worth recording how the wrong diagnosis was caught:** the first guard written
for this was "refuse to overwrite an existing file", and its test would not go
red. That is what proved the overwrite theory wrong — `nextRoundNumber` cannot
produce a colliding name. A guard that cannot fail is evidence about the
diagnosis, not a spare safety net.

## Rounds 066 + 067 (`d8ba7df`, 2026-08-16) — the prediction failed, and the instrument was wrong too

**The staked prediction was: single-batch grouping moves off 24% toward the
multi-batch 100%, and `tags-undefined` falls with it. Half of that happened, and
the half that did not is the more useful half.**

The pair is identical on every number, which is what makes it worth reading:

    round   1-batch draws  grouped  not-grouping  group-5010  tags-undefined  cfg-tag-5010
    065           5           0          0            5             5              3
    066           5           0          5            0             0              8
    067           5           0          5            0             0              8

**Grouping did not improve. It was already zero.** Widening `refreshShapes` did
not turn single-batch charts into grouped charts; it turned a doomed `addGroup`
into an honest decline. Ten spurious errors a round — five 5010s and five
`tags-undefined` — became five `not grouping` lines. `cfg-tag-5010` rose from 3
to 8 because those five charts now reach the tag write and are refused there
instead of failing before it. **For the user the outcome is unchanged; for anyone
reading a round it is much clearer.**

**Why it did not work, which is the next thread.** The single-batch charts show
**no settle-delay line**. Their re-read did not come back empty — it came back
with items that matched none of our ids, and the retry only fires on an empty or
partial read. A zero-match falls straight through to `use: "none"`.

A single-batch chart loads its shape ids in the same sync that creates them and
re-reads on the very next one, so the ids are plausibly not resolvable yet —
which is the same settling story that DID work for charts 4 and 5. **Retrying on
a zero-match is the obvious next experiment**, and it is small.

### THE 100% WAS MINE, AND IT WAS WRONG

`poolBatchSpanVsGroup` reported `333 of 333 = 100%` for the multi-batch arm. That
number is a bug in the pooling function, not a fact about the host. It looked for
the grouping verdict within a fixed few trace entries and did not count
`not grouping` among the outcomes — so **every honest decline was dropped from
both arms**, and only draws that either grouped or threw were counted at all.

Corrected, the archive says **353 of 452 (78%) against 49 of 214 (23%)**. The
separation is real and still the sharpest thing here; the absolutes were not.

Two things made it worse than an ordinary slip. It was quoted into a commit
message, a PR body, `docs/BACKLOG.md`, a code comment and two tests before anyone
checked it. And the re-read this very change added pushed the verdict further
from the draw, so the report showed **zero single-batch draws** in round 066 — the
instrument went blind in exactly the place it was being used to measure.

It was caught because the number was implausible, not because anything failed.
That is the same lesson as the starved-questions sweep, from the other side:
**pool your instrument's own behaviour, and distrust a clean 100%.**

## Rounds 068 + 069 (`2c7dcd8`, 2026-08-16) — the pair did NOT replicate, and that is the result

**Reading 068 alone would have shipped a false conclusion**, and this pair is the
case for the discipline added the same day.

    round   1-batch draws  grouped  not-grouping  retries  unmatched(traced)
    067           5           0          5           2           0
    068           5           4          0           7           4
    069           5           1          4           7           5

**068 read as a clean win: 0 to 4 of 5.** 069, same build, no merge between,
scored **1 of 5**. The retry fired identically in both (7), and the zero-match
persisted in both (4 and 5) — what moved is whether the charts grouped at all.

**One instrumented field explains the whole spread.** `listed` — how many shapes
the host named — on the traces that survived the retry:

    068    listed 10, 9, 16, 17     all >= drawn  ->  positional fallback fires  ->  4 grouped
    069    listed  1, 1, 1, 1, 31   four short    ->  fallback cannot fire       ->  4 declined

Four reads returned **ONE shape** for a slide holding seven to nine, and every
one of them carries `afterRetry: true` — the host read that short *after* a
1.5-second settle. More waiting is not the answer.

**So the mechanism is now fully named, and it is not what the fix assumed:**

1. The pre-grouping re-read **never matches our ids** on these charts.
   `withOwnId` is 7 of 7 and 9 of 9, so our handles are fine — the host lists the
   slide under ids that are not the ones it gave us at creation.
2. Grouping therefore depends entirely on the **positional fallback**, which is a
   guess ("the last N on the slide") and only legal when nothing matched.
3. That fallback needs `listed >= drawn`, and whether the host lists enough is
   **moody**: 9-17 one round, 1 the next.

**What the change is actually worth.** The old build grouped these charts 0 times
in two rounds (066, 067). The new build grouped them 4 and 1. That is a real
improvement — it never happened before and now sometimes does — but it is
unreliable, and it arrives through a positional guess rather than through the id
match the retry was meant to repair. **Recorded as partial, not as a win.**

Scenario level is unchanged in both: 10 of 12, `same scale` 4 of 8, and `the
chart is actually visible` still ends "the host would not name the chart
afterwards, so it carries no config". Grouping improved; the config did not
follow, which is consistent with the earlier correction that grouping is not what
saves it.

### The web search, recorded including its null result

No upstream issue reports this symptom — a collection listing shapes under ids
that do not match the ones returned at creation. The near relatives are all
already cited here: #5022 (sync hangs after add-then-read; delay workaround,
applied), #2903 (applied), #6498 (inserts not reflected instantly), #6363
(properties unavailable after sync, ten failed attempts). **Written down so the
next session does not spend the same twenty minutes finding the same absence.**

**The internal parallel is stronger than any of them.** This project already
measured the SLIDE version: scratch slides came back as `4123571114#123571113`
while the deck listed `287#62081387` — a freshly added slide reporting an id in a
namespace the deck will not answer to. Round 068/069's shape traces are the same
signature one level down.

**And the search surfaced something worth reopening**: Microsoft's shape-BINDING
documentation, unprompted, in two of three searches. `bindings.add` takes the
live Shape proxy inside the batch that created it — no id round trip, no
collection read, which is exactly the two things failing here. This repo parked
that route as *unanswerable from the probe* (`binding-names-shape-later` asked
eight times, never once reached its own question). That was a fact about the
PROBE. The retry proved the other route: measure it in production, where the
question can actually be asked.

## Rounds 070 + 071 + 072 (`01f3607`, 2026-08-16) — `same scale` passes, and not for the reason it was built

**THE SCENARIO THAT HAD FAILED 35 ROUNDS OUT OF 35 NOW PASSES, THREE TIMES:**

    070   12 of 12 scenarios   same scale 8 of 8, all 8 re-editable
    071   12 of 12 scenarios   same scale 8 of 8, all 8 re-editable
    072   12 of 12 scenarios   same scale 8 of 8, all 8 re-editable

`explode a degraded picture` passes in all three as well, and `the chart is
actually visible` has lost its "the host would not name the chart afterwards"
caveat. Twelve of twelve is the first clean sheet this project has recorded.

**THE FIX WORKED FOR A REASON IT WAS NOT DESIGNED FOR, and that is worth more
than the win.** The binding was added to give the SETTLE a durable handle. The
settle route has never once executed:

    round   grouped  notG  tagFail  cfg5010  origin5010  bind ok/fail  retries  unmatched
    069        11      4       8       8         0          0 / 0         7         5
    070        18      1       0       0         7          0 / 0        10         0
    071        19      0       0       0         7          0 / 0        10         0
    072        14      4       1       0         3          0 / 0        10         3

`bind 0/0` in every round: nothing fails, so the settle has nothing to rescue.
What changed is upstream of it. **The zero-match re-read is gone** — the defect
rounds 068/069 named exactly (`withOwnId` 7 of 7 with zero matches, the host
listing shapes under ids it never gave us). Calling `bindings.add` on the live
proxy in the drawing batch appears to **stabilise the shape's identity**, so the
re-read afterwards reports ids that match ours and the config tag write lands.

Attribution is clean: the only difference between `2c7dcd8` and `01f3607` is
that one commit.

**THE HOST IS STILL MOODY UNDERNEATH, and that is the honest reading of the
third round.** 072 grouped 14 where 071 grouped 19, declined 4 where 071 declined
0, and had 3 unmatched re-reads where 070 and 071 had none. Retries fired 10
times in all three. So the wobble did not go away — **the OUTCOME became robust
to it.** `same scale` is 8 of 8 whether the internals wobble or not, which is a
better result than a quieter host would have been.

`cfg5010` is **0 in all three** post-change rounds against 8 before. That is the
cleanest single signal in the set.

**What is now the top failure, and it is the one nothing here can check.**
`writing the chart's origin tag` 5010s 3-7 times a round. The origin tag is the
drag-delta round trip, and `docs/PUBLISHING.md` says it "needs a real drag and so
cannot be scripted at all". The failures have migrated to precisely the feature
no automated round reaches — see the backlog for the proposal to make the origin
ARITHMETIC scriptable even though a mouse drag is not.

### A driver gap this run exposed and closed

A browser that died UNDER a running round was invisible for 24 minutes.
`quietStreak` resets to zero on a failed CLI call — deliberately, so contention
cannot end a healthy round — but a dead browser makes every call fail forever, so
the counter never reaches its threshold and the loop polls a corpse to the
thirty-minute limit. `browserDiedMidRound` now needs a STREAK of failures AND an
affirmative `(no browsers)`; either alone re-makes a bug the other prevents.

## Rounds 075 + 076 (`aeefb93`, 2026-08-16) — the origin tag is fixed, and the binding is proven

**The prediction was staked before the pair and it held in both rounds:**

    round   scenarios  origin5010  charts that cannot follow a drag  cfg5010  grouped
    074        12/13        8                    8                     0       17
    075        13/13        0                    0                     0       19
    076        13/13        0                    0                     0       16

**Thirteen of thirteen, twice** — the first perfect sheets this project has
recorded, and the first with the drag check in the list at all.

**What was wrong.** The origin tag is the one write that cannot dodge a resolved
handle. The config tag is queued BEFORE `load("id,left,top")` and lands; the
origin tag needs that loaded position, so it went out after the load had
resolved the target into `shapes.getItem(id)` — the handle this host refuses.
Before the binding change the CONFIG tag took the 5010s and the origin tag none;
after it the config tag took none and the origin tag took eight or nine a round.
**The refusal never went away. It moved to the only write still going through a
resolved proxy**, and stayed invisible because a scenario samples one chart while
half the population was failing.

**And this refutes the alternative I staked against it.** `binding.getShape()` is
NOT `shapes.getItem(id)` under another name. A write through the binding lands
where the identical write through a resolved proxy was refused eight or nine
times a round. That closes the question the probe could never reach —
`binding-names-shape-later` asked eight times across eight rounds and never once
got to its own question, which was a fact about the probe.

**The host is still moody underneath and the outcome no longer cares.** Grouping
was 19 in one round and 16 in the other on the same build. Every scenario passed
in both. That is the same shape as the settled retry: the wobble did not go away,
the result became robust to it.

### What this leaves

`same scale across the deck` has now passed on **070, 071, 072, 073, 074, 075
and 076** — seven rounds across four builds, after failing 35 straight. It is no
longer a claim that needs a pair.

**The pooled `CHARTS THAT CANNOT FOLLOW A DRAG` count is now history rather than
news**: 34 charts across 5 rounds, all of them 072-074. It will not climb again
unless something regresses, which is exactly what it is there to notice.

**The settle-by-binding route has still never executed.** `settledByBinding` is 0
across seven rounds — the settle only runs when the config tag fails, and the
config tag stopped failing. It is now a fallback that has never fired, kept
deliberately: it costs nothing at rest and it is the only route left if the
config tag ever starts failing again. Recorded so nobody mistakes it for tested
code.

## Round 077 (`357632b`, 2026-08-16) — the first 4:3 round, and it is CONFOUNDED

**The first round in 53 not run at 16:9.** It scored **10 of 13** against 13 of
13 twice on the same build, and the failure is a class this project has never
seen:

    round   scenarios  UnexpectedError  tagging-failed  cfg5010  orig5010
    076        13/13          0               0            0        0
    077        10/13         52              18            0        0

    36 × UnexpectedError  at=writing the chart's config tag
    16 × UnexpectedError  at=settling the config tag through its binding

**The 5010s stayed fixed.** `cfg5010`, `orig5010` and `origLost` are all zero —
everything landed this week held. What appeared instead is `UnexpectedError`,
mostly on a write through the GROUP handle (`from: group×1`, 17 of 18).

**AND IT EXERCISED CODE THAT HAD NEVER RUN.** `settledByBinding` was 0 across
seven rounds because the settle only fires when the config tag fails and the
config tag had stopped failing. Here it fired sixteen times — and the binding
route was refused too, with the same `UnexpectedError`. So the fallback is no
longer untested; its first test says it does not save this case.

### WHY THIS IS NOT YET A FINDING ABOUT 4:3

**Two variables changed at once**, which is the exact mistake `docs/ROUNDS.md`
records about rounds 24 and 25 — "differed only in this, and were compared as
though they did not".

1. **The aspect ratio**, which is the thing under test.
2. **The deck.** Round 077 ran on `Presentation65`, not the `Presentation64` that
   carries all 52 previous rounds. That deck was created during a browser crash
   earlier the same day, and its whole ribbon — Present, Share, Slide Layout,
   Slide Size — was greyed out until a reload. A document whose session was
   already damaged is not a clean instrument.

`UnexpectedError` is also a different SHAPE of error from everything in this
archive: it is Office.js's generic failure, not the specific
`InvalidParam / 5010` that names a refused handle. That is as consistent with a
sick document as with an aspect ratio.

### THE CONTROL THAT SEPARATES THEM, and it costs one round

**Switch `Presentation65` back to 16:9 and run it again.**

    back to 13/13   the deck is fine and 4:3 is the cause
    still ~10/13    the DECK is the cause and 4:3 is exonerated

One variable moves, and it is the cheap one. The alternative — switching the
proven `Presentation64` to 4:3 — tests the same question but risks the deck
carrying 52 comparable rounds, which is not a trade worth making while a
cheaper control exists.

**Nothing about 4:3 support should be written down until that control has run.**

## Round 078 (`8e9227c`, 2026-08-16) — the control, and 4:3 is EXONERATED

**Same deck, same build, 16:9 instead of 4:3. One variable moved.**

    round   deck             size   scenarios  UnexpectedError  tagFail
    076     Presentation64   16:9      13/13          0            0
    077     Presentation65   4:3       10/13         52           18
    078     Presentation65   16:9      10/13         52           18

**077 and 078 are identical on every number.** The failure follows the DECK, not
the aspect ratio. `Presentation65` was created during a browser crash, has gone
into an unusable greyed ribbon state twice, and turned up at the start of this
round holding **100 slides**. It is a damaged document, and every one of those 52
`UnexpectedError`s belongs to it.

**Had the control not been run**, "4:3 breaks the config tag write" would have
gone into the docs on the strength of a single round, and every subsequent 4:3
result would have been read through it.

**Two things it also proved in passing.** The driver swept a 100-slide deck
rather than refusing — the heal added earlier the same day, on a deck far past
anything it was designed against. And `settledByBinding` fired sixteen times
after seven rounds of never running: the fallback is no longer untested, and its
first test says it does not save this case.

### What the 4:3 question still needs

**Nothing has yet been learned about 4:3.** Round 077 measured a sick deck. The
question needs a CLEAN 4:3 deck, which is now the first item of the nightly work
below.

## Rounds 079-081 (`17a8204`, 2026-08-16) — the first full cycle, and 4:3 is fine

**One 4:3 validation round and a 16:9 pair, all on one build.**

    round   deck             size      scenarios
    079     Presentation66   720x540     13/13
    080     Presentation64   960x540     12/13   explode a degraded picture
    081     Presentation64   960x540     13/13

**4:3 SCORED BETTER THAN 16:9 IN ITS OWN CYCLE.** Round 077's 52
`UnexpectedError`s belonged entirely to `Presentation65`, a deck created during
a browser crash — round 078 reproduced them at 16:9 on that same deck. On a
clean 4:3 deck the failure does not exist. **Nothing about 4:3 is broken.**

Both new mechanisms worked on their first real use: the size is recorded
(`{"width":720,"height":540,"source":"pageSetup"}`) and readiness confirmed the
profile before each round (`slide size 4:3 (want 4:3)`).

### THE DIVERGENCE CHECK GOT ITS FIRST OUTING WRONG, and it is fixed

It reported:

    explode a degraded picture — passed at 4:3, failed at 16:9

True of the worst reading and **wrong about the cause**. The 16:9 pair disagrees
with ITSELF: 080 failed and 081 passed. Collapsing a profile to its worst
outcome turned a flaky scenario into a slide-size difference, and would have sent
someone to investigate an aspect ratio for a scenario that is simply unreliable.

`profileDivergence` now keeps both outcomes per profile and separates them:

    DIVERGED between slide sizes      one profile passed, the other failed
    UNSTABLE WITHIN a slide size      one profile did both — flaky, not different

The gate says the second in its own words and does not call it divergence.
Mutation-proven.

**Worth noting what caught it**: not a test, but reading the first real output
and finding it did not match the rounds it was describing. A check built the same
afternoon, wrong on its first live data, is the ordinary case rather than a
surprise.

### `explode a degraded picture` is now instrumented rather than guessed at

Five failures across eight rounds and **no evidence**: the scenario reaches a
null from `updateChartInSlide` and the trace held nothing between the deck scan
and the verdict. There is a line now, carrying `asPicture` first because the
picture path is the one that fails intermittently.

**Upstream has no match.** office-js#3698 (image insertion refused while a shape
is selected) is the nearest and is a different API — `setSelectedDataAsync`, not
`ShapeFill.setImage` — with no evidence here of a selection. #225 (large base64
fails mid-insert) leaves a half-inserted image where ours leaves the chart
intact; #5022 hangs where ours returns. **A promising chain via #5443 was killed
by reading it**: that issue is Excel on Mac, not PowerPoint shape selection.

So no workaround was applied. Guessing one now would produce a change nothing
could evaluate; the next round that reproduces this will say which step failed.

## Rounds 082-087 (`9f175a8` → `95170cf`, 2026-08-16/17) — six perfect rounds, and nine bugs in the instruments watching them

    round   build      scenarios   note
    082     9f175a8      13/13     first perfect round in the archive
    083     9f175a8      13/13     the pair
    084     600ea90      13/13     the orphan instrument's first non-zero reading — a phantom
    085     fbdc49b      13/13     survived a real browser death, recovered unaided
    086     2c2547f      13/13     settle guard caught two phantoms, missed a third
    087     95170cf      13/13     4s settle held; zero contradicted by the deck

**Six consecutive 13/13 rounds, zero `UnexpectedError` in any of them.** `explode
a degraded picture` passed all six, having failed 47 of its first 57.

### The product was fine. The instruments were not.

Two product defects were found across the whole stretch — a blank slide shipping
in the finished deck, and an update dying on a refused id. **Nine defects were in
the reporting**, five of them introduced the same day they were found.

### The orphan instrument was wrong four times, each differently

1. **Mismatched units.** It subtracted inner chart shapes, delete CALLS and
   top-level slide shapes from one another and reported "283 stranded" beside its
   own `before: 3, after: 3`.
2. **An algebraic identity.** `shortfall` and `unexplained` summed to zero on
   every line, so the reading was the same on healthy and broken data — and its
   control test passed vacuously.
3. **One stale host read.** "92 shapes grew" was PowerPoint lagging an
   `addGroup` our own sync had already resolved; the deck showed one grouped
   chart per slide.
4. **Two agreeing stale reads.** With a 1.5s settle added, both reads landed
   inside one lag and agreed on the stale number, so the reading was marked
   `settled` and was still wrong.

**Every one was caught by the same second source** — `deck.inventory`, already in
the round file, taken long after any lag. It has never been wrong. It is now
cross-checked on every reading, and disagreements are counted rather than
discarded, because an instrument's own error rate belongs in its report.

The settle delay was then sized from data that had been in every round file for
weeks: the gap from a group commit to the count that followed it is stale up to
**3193ms**, so 1.5s was never enough. No new instrumentation was needed to learn
that — the timestamps simply had not been differenced.

### The gate spent fifteen rounds blaming nine innocent builds

`traceNovelty`'s "NEW BEHAVIOUR" bucket took its median over EVERY prior round,
so a signature stayed new until it appeared in more than half the archive — while
the denominator grew. `re-reading the slide's shapes again after a settle delay`
debuted in round 064 and has sat at 10-11 since; it was announced as new in
fifteen rounds and attributed to nine different builds, the last a commit that
only changed a slide counter. The signature round 086 had actually changed went
unmentioned.

Now a five-round window, and each entry names the build it first appeared in. On
the real archive the bucket reports **nothing**, which is the truth.

### A blank slide of ours was shipping in the finished deck

One `slides.add()` can land TWO slides. The branch that fires when a landed slide
cannot be named reported one — `after - before` is 2 in all ten archive
occurrences, never 1 — so the sweep's clamp ran one short and left the remainder.
Round 085's inventory carries it: `257#3837665135`, zero shapes, listed in
`newSlides`. Nine rounds are affected and five of the last thirteen began their
scenarios on a deck already dirty by one slide, invisible because the scenarios
measure GROWTH.

The comment above that branch cites a real 2026-08-11 measurement; what went
stale was the inference drawn from it, contradicted ten times since by data
nobody re-read.

### The stranding question has three observations

Round 087 finally failed to group something (15 grouped, 4 not). Pooled across
every round that can answer: **three at-risk charts, all zero growth.** Three is
not five. A round that groups everything cannot answer this either way, and the
report now says so rather than reading zero as an all-clear.

### The browser died three times, and the cause is not a crash

`ERR_NETWORK_IO_SUSPENDED` in the console tail, and Windows Event 507 — entering
connected standby — four times in one morning. The machine was sleeping. Recovery
survives it, and did so unaided in round 085. **A sideload does not always
survive** the browser process: once it was lost and once it was not, and an
`Upload My Add-in` appears to register against the account rather than the
session.

### Six vacuous tests

Written and caught in this stretch, all the same shape: built from what the
author expected rather than from what the code does. A callback passed as a
fourth argument to a three-parameter function; a fixture asserting a lookup
happened rather than a click; `ref=rCancel` against a `[a-z0-9]+` matcher; a
fixture modelling "rare" as "absent"; a control that could not fail; and a
novelty fixture with the signature in every prior, where the old and new readings
agree. **Mutation caught all six. Reading caught none.**

## Round 088 — 3056f91 — 12/13 — and a prediction that could not be answered

First round after the 14-commit merge (#576-#590). Browser was dead at `--check`;
the profile still held the sign-in, so it was reopened without a password and
**the sideload survived** — see the doctrine note below, because this file said
it would not. `HEAD 3056f91 · site 3056f91 · pane 3056f91 · deck 1 slide`.

The reading was pre-registered before the host was touched, which is the only
reason the round is legible. All three calls came out.

- **Mine.** `same scale across the deck` PASSED — its tenth consecutive pass, not
  its first. #586's branch was never entered: it runs only on a re-read that is
  SHORT after the settled retry, and round 088 recorded none. Across all 64
  rounds there are 42 short re-reads and `afterRetry: true` on **exactly none**
  of them — every one predates `REREAD_RETRY_MS`. What the round did record is
  the failure #586 deliberately does not rescue: 2 empty re-reads and 1
  zero-match, all after the retry, on charts that then hit `not grouping: no
  member handle this host will accept` and a 5010 on the id readback.

  Those two charts kept their config anyway. Tempting, and not a finding: the
  pooled table already prices an ungrouped chart at 32% keeping its tag, so
  two-for-two is inside the noise. Recorded so the next session does not spend a
  round on it.

  **The one real finding is the denominator.** The scenario ran **6 charts, not
  8** — the first time in 64 rounds. Its verdict is `scaled === charts.length`
  against a population it DISCOVERS (`probeCharts`, guarded only at `< 2`), so
  `6 of 6` is a pass and the gate compared PASS to PASS and said "no scenario
  regressed". Cause, consistent with four fields and not proven by one round:
  the upstream `insert onto a slide that already has content` stalled mid-draw
  (`shapes 1-10 of 16, 45s`), so fewer probe charts existed to find. Deck slide
  count was 7 — identical to nine rounds that found 8 — so deck size is NOT the
  cause; that hypothesis was formed and refuted against the archive.

- **Research.** office-js **#5022** (open, under investigation, product bug):
  `context.sync()` hangs after delete-then-re-read on the web; the reported
  workaround is "a timer of 1-2 seconds between the `shape.delete()` and the next
  `await context.sync()`" and the reporter says it **remained unreliable**. That
  is `REREAD_RETRY_MS` at 1500ms, independently corroborated as a mitigation
  rather than a cure — which is exactly round 088's two empty re-reads surviving
  it. **#6498** (open, needs attention, 2026-02-09): shapes inserted on the web
  not reflected without a refresh; same family, no workaround offered. Nothing
  upstream we are not already doing. Recorded so it is not rediscovered.

- **Instrument.** `poolScenarioPopulations` in `triage.mjs`, printed by the gate:
  it names a scenario passing on a smaller population than it usually runs. On
  round 088 it prints `same scale across the deck — 6 this round, usually 8 over
  63 prior round(s) (and it still reports PASS)`. Median, not mean, so one small
  round cannot lower the bar for the next — the mutation that turns it into a
  mean fails its own test.

- **Fix.** Three defects in the judge, each with a test that goes red when the
  fix is reverted:
  1. `scenario-passes` recorded a SKIPPED scenario as `FAILED`, while
     `probe-answers` and `probe-detail-matches` beside it already answered
     `undetermined` — and this file's doctrine says "a skip is not a flip". Round
     088 made it live: `insert onto a slide that already has content` came back
     `skipped: true, ok: false`.
  2. `roundToJudgeOn` took the FIRST round on a build, so with the cycle's three
     rounds per build a prediction could be judged on a sibling of its own
     control. Now the last.
  3. `stampDate` read the stamp down to the DAY and compared with strict `>`, so
     a round taken the same day a prediction was staked could not judge it. Round
     088 is that case exactly. Reading `HH:MM` too makes it work by lexicographic
     order with no extra branch.

  **The first draft of test 2 was vacuous** and mutation caught it: the fixture
  put a later build after the siblings, and both readings return that later
  build. The siblings have to be the newest rounds or the test proves nothing.

- **Doctrine.** Four documents asserted things this archive contradicts:
  - `docs/ROUNDS.md` said "a group is deleted whole", so only an ungrouped chart
    can strand shapes. #586 ended that: it groups a SUBSET and leaves the
    remainder out of the parts tag on purpose. **`atRisk` reads the host's shape
    type, so a subset group reports `group` and is counted SAFE** — the
    instrument is blind to the only stranding the code now creates deliberately.
    Latent today (the branch has never run), a landmine tomorrow.
  - `docs/round-prompt.md` and the #586 ledger entry both stated "failed 34 of 34
    rounds" and "chart 4 matched 20 of 24 in every round on record" as current
    fact. Neither is true; both corrected in place with the archive numbers.
  - `docs/BACKLOG.md` asserted the shipped change and its opposite, in two live
    sections of one file.
  - **A browser death does NOT always take the sideload.** It did on 2026-08-18
    and did not on 2026-08-19; the ribbon group was intact after the relaunch.

  The prediction is recorded `undetermined` on `088-3056f91`, not held. Once the
  judge was fixed it began printing **held** — a false positive that two bugs had
  been hiding between them: a claim that could not discriminate, behind a judge
  that refused to judge. Fixing one without the other manufactures the artifact.

## Round 089 — 3056f91 — 13/13 — the pair, and what it took back

Same build, no merge between, which is the discipline. **PowerPoint crashed 324s
into attempt 1**; the driver caught the dialog, wrote the host's own account to
`crashes/2026-08-19T14-46-31.md`, cleared it and ran again. Attempt 2 finished
clean. That machinery had been built and never yet met the thing it was for.

- **Mine — the pair takes back round 088's loudest reading.** 089 ran `same scale
  across the deck` at **8 of 8**, not 6. So the shrunken population does NOT
  replicate: it is what happens WHEN an upstream scenario is skipped, not a new
  steady state. Reported alone, round 088 would have read as "the scenario
  shrank". This is what pairs are for, and it is the second time the same
  denominator has needed one.

  The deck inventories say the same thing from the other side: 088 ended
  `0,2,2,11,24,24,1` — two slides carrying 24 LOOSE shapes, the two charts whose
  re-read came back empty — and 089 ended `0,4,2,11,1,1,1`, one shape per slide,
  every chart grouped.

  **#586 was not exercised in EITHER round.** 0 short re-reads in both. Two
  observations now, on one build, and 42 short re-reads in the archive with
  `afterRetry: true` on none of them.

  **The zero-match is the live one:** 1 in 088, 2 in 089, 3 in 087. It is the
  failure #586's strict-majority bound deliberately does not rescue, and it is
  the only one still happening.

- **Research — the crash was the network, not the host and not us.** The captured
  console is `ERR_NETWORK_CHANGED`, `ERR_NAME_NOT_RESOLVED`, `ERR_HTTP2_PING_FAILED`,
  `ERR_CONNECTION_CLOSED`, and a `wss://…augloop.office.com` that would not
  resolve. Same family as the dead browser at the start of this session, and the
  same window in which `gh` could not reach `api.github.com` and `curl` could not
  resolve a host from this machine. Nothing to fix in this repo; recorded so the
  next crash file is not read as a PowerPoint bug.

- **Instrument — the novelty view earned its keep.** It named **one trace
  signature never seen in 64 prior rounds**: `error|reading the deck's style`.
  That is #583, merged the same morning and never before run against a host.

- **Fix — a failed read was reporting an absence.** `reading the deck's style`
  hung for its full 90s budget while the host kept answering other calls
  throughout it (see the correction under round 090 — this line first had the
  ordering backwards). `readDeckStyle` caught that and returned `null` — the same value a deck
  carrying no style returns — and `style-from-deck` AWAITS it, so the pane would
  tell the user their deck is unbranded and switch them to the browser's style on
  the strength of a read that never happened. `readDeckStyleWithReason` now
  separates them; the fire-and-forget pane-load caller keeps the narrow contract
  it wants. The fake could not express a host that has the API and will not
  answer the READ — `refuseCustomXmlReads` closes that, beside the write switch
  that already existed.

  **This is the same defect this repo keeps finding in new places:** unreadable
  reported as negative. `collectDeckEvidence`'s `beforeUnknown`, "no history is
  not a spike", "a miss is not a failure", and now a deck style. Fourth time.

- **Doctrine.** The crash-recovery path works on a real crash, first evidence.
  The round-089 population is the counter-example the brief needs beside round
  088's: a scenario's denominator moves with what ran before it, so quote the
  denominator or quote nothing.

## Round 090 — f6cd580 — 13/13 — and the deck-style read fails every time

Four attempts, and the driver drove all of it: a stale pane after the merge, then
a browser death 51s into attempt 3, then a clean run. No password anywhere. The
round before it never started at all — see the refusal below, which was the
session's largest finding and was not in the round file.

- **Mine.** 13 of 13, `same scale` 8 of 8, and the cleanest re-read round on
  record: **0 short, 0 zero-match, 0 empty**, 20 grouped lines with `partial:0`.
  Deck ended `0,4,2,5,1,1,1` — one shape per chart slide, everything grouped.
  Nothing changed in the draw path between 089 and 090, so a clean round is
  weather: 2 zero-matches to 0 is inside the noise floor and is not a result.

  **#586 still not exercised — three rounds now.** 0 short re-reads after the
  settled retry in 088, 089 and 090.

- **THE DECK-STYLE READ FAILS ON THIS HOST, 2 FOR 2.** Round 089 flagged it as a
  signature never seen in 64 rounds; round 090 reproduced it exactly:

      089  afterMs 90000, idleMs -89057, "listing the deck's slides" answered in 44ms
      090  afterMs 90000, idleMs -89xxx, "listing the deck's slides" answered in 56ms

  **CORRECTED: those two lines were first read backwards, here and in #595.** The
  44ms and 56ms are how long that OTHER call took, not how long before the stall
  it ran. `idleMs` is `issued - lastAnswered` and `lastAnswered` is sampled when
  the deadline fires, so a NEGATIVE value means the host answered something else
  89 seconds INTO this call's wait — one second before it was abandoned.

  Two rounds, both since #583 landed, which is every chance it has had at that
  point. The host is alive throughout and the read never answers.

  **CORRECTED after round 092 — "has never once succeeded here" was wrong.** It
  succeeds in 428ms. See the round 092 entry: the failure is INTERMITTENT, and
  this line was written from the only two observations that existed.

- **Research — a null result, recorded as one.** No upstream issue describes a
  PowerPoint-web custom-XML READ that hangs. office-js #2937 ("CustomXMLPart
  can't sync") is Word on the desktop and bibliography-specific; #3936 is Excel
  on iPad and needs an unsaved file. Neither is this. Worth reporting upstream —
  the owner's call, not ours to post.

- **Instrument.** The pooled report now says, in as many words, that #586's
  subset branch has not run in the pool, instead of printing a zero at-risk count
  that reads as an all-clear. The population check stayed quiet, correctly: 8 of
  8 is the usual denominator.

- **Fix.** The deck-style read gets its own 10s budget instead of the 90s
  readback one. **The number is not sized from a successful read, because this
  host has never produced one** — it bounds the damage rather than fitting a
  distribution, and both callers degrade safely: the pane keeps the browser's
  style, the button says it could not read. `PW_DECK_STYLE_TIMEOUT_MS` overrides.
  What the 90s was costing: a `PowerPoint.run` context held open for a minute and
  a half on EVERY pane load, and a person waiting that long before the button
  admitted defeat.

- **Doctrine — the refusal that stopped the round before it started.** `--check`
  said this document had no add-in and that a reload would not bring it back: the
  one stop recovery may not retry. The tree said `button "Insert chart"
  [disabled]` and `status: Disconnected` the whole time. `refFor` returns the ref
  on the matching line and Playwright issues refs only for actionable elements,
  so ABSENT and DISABLED arrived as the same null and the driver chose the
  permanent explanation for a transient state. Now `host-disconnected`, and
  recoverable — proven on the same document, which then reloaded and came back.

  **Fifth time this codebase has reported UNREADABLE as NEGATIVE**, after
  `beforeUnknown`, "no history is not a spike", "a miss is not a failure", and
  the deck style. It is the house defect.

### After round 090 — the number nobody could read was the answer

Not a round. Mining round 090's own trace harder, before staking the next one.

**`idleMs: -89057` is not a glitch, it is the finding.** It reads as nonsense, so
two rounds carried it and nobody used it. `idleMs` is `issued - lastAnswered`,
and `lastAnswered` is sampled when the DEADLINE fires — so a negative value means
the host answered something else AFTER this call went out, while it was still
waiting. PowerPoint answered `listing the deck's slides` in 44ms, **89 seconds
into the deck-style read's 90-second wait**, one second before it was abandoned.

So the deck-style stall is **one call stuck on a healthy host**, not a host that
went quiet. Those are different faults wanting different fixes, and this project
had no way to say which it was looking at.

`stallShape` now says it in words on every stall, because the sign of a number is
not something a reader notices — and the comment on `lastStall` had assumed
sequential code, where a negative gap cannot happen. The pane's deck-style read
is fire-and-forget, so it breaks that assumption, which is exactly why this is
the call that exposed it.

**A cheaper question, asked only after the read has already failed.** The
namespace lookup, `getOnlyItemOrNullObject`, the `load` and `getXml` all go into
ONE batch, so one sync covers four calls and the trace can only report the sync.
`getCount()` on the same scoped collection needs the namespace lookup and nothing
else, so its answer splits the two explanations:

    it answers    -> the collection is reachable; the fault is in
                     getOnlyItemOrNullObject / load / getXml
    it also hangs -> the whole customXmlParts surface is dead here and #583
                     cannot work on this host at all

The next round decides it, which two rounds so far could not.

## Round 091 — de76543 — 13/13 — and the feature nobody could click

Two attempts. The driver cleared a silent host and a stale pane by itself, which
is the pair the previous session left it: `8012ms silent` before, `3ms` after.

- **Mine.** 13 of 13. The staked prediction came back **`undetermined — neither
  line appeared`**, and the ledger was right to say so rather than call it a
  refutation. That is what `insteadOf` was added for, one round earlier.

  **The reason is my own fix.** The deck-style read fires from `Office.onReady`,
  and dropping its budget from 90s to 10s moved its failure out of the round's
  traced window: round 090 recorded it at entry 135 of 528, round 091 has not one
  deck-style line in 529. The bug did not go away, the instrument stopped being
  able to see it. **A fix that hides its own subject is worse than the cost it
  removed.** The conclusion is remembered now and replayed at round start, so the
  next round can judge the claim it was staked for.

- **AND THE FEATURE WAS UNREACHABLE ANYWAY.** Chasing why the probe never ran, in
  the live pane after the round:

      Office.context.host  "PowerPoint"    <- isPowerPointHost() would say true NOW
      style-from-deck      disabled: true
      style-to-deck        disabled: true

  `app.ts` asked `isPowerPointHost()` at MODULE SCOPE, where `Office.context`
  does not exist yet, so it answered false on every load inside PowerPoint and
  disabled both buttons for the whole session. Nothing ever asked again. **#583's
  entire user-facing feature could not be clicked on PowerPoint web**, and
  thirteen scenarios could not see it because none of them clicks a button.

  Swept: `Use deck theme` has the same gate and `renderOptions()` is also called
  at module scope, so it was disabled for the same reason — hidden only by
  incidental re-renders, which makes it harder to diagnose, not safer. Both are
  under one `syncHostOnlyButtons()` now, called again from `Office.onReady`.

- **Research.** Nothing new to search: the question this round was staked to ask
  was never put. Recorded rather than skipped.

- **Instrument.** `stallShape` shipped this round and had nothing to describe —
  no stalls. The replay above is the round's real instrument change.

- **Doctrine — a correction to yesterday's correction.** "Playwright hands out a
  ref only for something it could act on" was too strong. Same session, same
  page: the ribbon's `button "Insert chart" [disabled]` has no ref, the pane's
  `button "Use deck style" [disabled] [ref=f19e328]` has one. Native `disabled`
  and whatever Office marks its ribbon with are not treated alike. The fix
  (`namePresent`) stands; the explanation did not, and is corrected in
  `round.mjs` and in memory.

## Round 092 — 0aa6f91 — 13/13 — the bug that stopped happening

Second attempt again; the driver cleared a silent host and a stale pane by itself.

- **Mine.** 13 of 13. The staked prediction came back `undetermined — neither
  line appeared` for the SECOND round running, and the reason is different this
  time. 091 could not ask because the 10s budget moved the failure out of the
  traced window. **092 did not ask because the failure did not happen**: zero
  deck-style entries, and the replay correctly stayed silent because there was no
  verdict to carry — which is the behaviour its second test pins.

- **THE READ WORKS. I said it never had, and that was wrong.** Driving the pane's
  own button on the real host after the round:

      click 1   428ms   "This deck carries no style — using yours."
      click 2-4  <1s    same

  That is the `isNullObject` path completing normally, not the timeout path. So
  the deck-style read is **intermittent**, not broken: 90s and never in rounds 089
  and 090, under half a second now. "#583's deck-style read has never once
  succeeded on this host" was written from the only two observations that existed
  and is corrected in the journal, in `powerpoint.ts` and in the ledger.

  **The 10s budget is now sized from evidence rather than from harm** — 428ms
  observed, so ten seconds is ~23x the success and still fails fast. Not tightened
  further: one sample from a healthy host is not a distribution.

  **A hypothesis at n=4, offered as one.** The two rounds that hung (089, 090) are
  also the two with a host crash and a browser death. The two that did not (091,
  092) booted cleanly. Four rounds is not an effect. The test is cheap though:
  when it hangs again, look at whether that round had a crash.

- **THE FEATURE IS REACHABLE FOR THE FIRST TIME, verified on the real host.**
  Before the fix, `button "Use deck style" [disabled]`; after it,
  `button "Use deck style" [ref=f25e328] [cursor=pointer]`, and clicking it
  answers. Everything above was only measurable BECAUSE that fix landed — the
  button could not be clicked at all until this round.

- **Research.** None this round: the question the ledger is staked on was not put,
  and inventing a search to fill the slot would be the padding this protocol
  exists to prevent.

- **Instrument.** The replay shipped and correctly emitted nothing. A test that
  only ever fires is not a test; the silent case is pinned too.

## Round 093 — 0aa6f91 — 13/13 — THE FIRST REAL PAIR, and it does not replicate

First attempt, no recovery needed. **And the first time in this stretch that two
rounds ran on ONE build** — 089 through 092 each had their own, because fixes
were landed between every round. The brief says not to do that, in bold, and the
reason is exactly what this pair shows.

- **Mine — same build, nothing changed, and the internals swing hard:**

      round   grouped  refused  errors  5010  deck inventory
      092        20       0        2      5   0,4,2,5,1,1,1
      093        15       4        9     11   0,4,2,17,24,24,24

  Three slides ended holding **24 shapes each instead of one**: three charts that
  did not group, **seventy-two shapes left loose on the deck**. And both rounds
  report `13/13` and the byte-identical verdict line `8 of 8 charts carry the
  shared scale (max=105 …); 8 still re-editable`.

  The trace and the deck agree exactly — 15 grouped and 4 refused in the trace,
  three big slides in the inventory — so both are trustworthy here. This is the
  authority confirming the instrument rather than contradicting it, which is the
  first time in this journal it has gone that way round.

  **The scenario is not lying.** It asks whether the config survived, and it did:
  the ungrouped fallback keeps the tag. It simply cannot see grouping — which is
  what #586, #520 and most of the last month have been about. So
  `scenarioRegressions` compares PASS to PASS across a round that left seventy-two
  loose shapes and one that left none.

  **Read as noise, not as a fall.** 0 and 4 refusals on the same build is inside
  this project's own floor (1 vs 5, nothing changed). The finding is not "093 is
  worse"; it is "the verdict cannot see the difference at all".

- **Research.** None: the pair answered a question about our own instrument, not
  about the host, and there is nothing to look up for it. Said rather than padded.

- **Instrument.** `poolGroupingOutcome`, printed by the gate every round —
  charts grouped, charts refused, the median refusals of the archive, and the
  deck line beside it. Counted from the trace, which is exact; the deck is
  printed as corroboration rather than used as the measure, because turning slide
  shape-counts into "ungrouped" needs a threshold and a guard sized by guesswork
  is how this instrument has been wrong before.

- **Fix — process, not code.** Stop landing between the two rounds of a pair.
  089-092 were four rounds on four builds and none of them can be compared with
  its neighbour. This pair took twenty minutes and produced the clearest result
  of the night.

## Rounds 094 + 095 — ab5d730 — 13/13 and 13/13 — the blindness replicates

A second pair, on a second build, run back to back with nothing landed between.

      build      round   grouped  refused  deck inventory
      0aa6f91     092      20        0     0,4,2,5,1,1,1
      0aa6f91     093      15        4     0,4,2,17,24,24,24
      ab5d730     094      20        0     0,4,2,5,1,1,1
      ab5d730     095      19        1     0,4,2,5,1,1,24

**Four rounds, two builds, every one 13/13 with a byte-identical verdict line** —
and two of them left charts ungrouped as loose shapes on the deck. The finding
from the first pair is not a one-off: `same scale` asks whether the config
survived, which it does either way, so grouping is invisible to it.

What the second pair adds is the SHAPE of the noise. Refusals are not rare and
not constant: 0, 4, 0, 1 across four rounds on two builds. Any single round's
refusal count is uninformative, and the difference between two rounds on one
build is not evidence of anything about the build. `poolGroupingOutcome` prints
it every round now precisely so nobody has to infer it from a verdict that cannot
see it.

**This is what pairing buys.** Four rounds run as two pairs answered a question
that eight rounds run singly could not have, because the comparison a single
round invites — against the previous round, on a different build — is exactly the
comparison the noise floor forbids.

## Rounds 096 + 097 — 9e81c14 — two staked claims held, and one of them closes a question

The richest pair of the run. Both predictions were written down before the host
was touched, and both came out.

- **`the-round-file-can-finally-say-which-host-it-ran-on` — HELD, both rounds.**
  The first two rounds in 71 to carry it:

      host PowerPoint · platform OfficeOnline · version 0.0.0.0
      requirementSets 1.1 … 1.10 · canInsertSlidesFromBase64 · canInsertPicture
      slide size 960x540, source documentFile

  So the trace mark was the whole story and there is no second slice. Seventy
  rounds were archived without any of this.

- **`the-deck-style-namespace-is-reachable` — HELD, both rounds, identically.**
  `the namespace IS reachable — the fault is further in`, `parts: 0`,
  `observedBeforeTheRound: true` — the replay carrying it across the mark, which
  is the whole reason it exists.

  **The expensive half is refuted**: `customXmlParts` is not dead on this host,
  so #583 needs a fix rather than a product decision. `getCount()` answers; it is
  `getOnlyItemOrNullObject` / `load` / `getXml` that hangs. And `parts: 0` names
  the call precisely — **asking for the only item of an EMPTY collection**, which
  is the state every unbranded deck is in, so every pane load in the wild.

  Fixed by counting first and never making that call unless the count is 1. It
  costs one round trip on a deck that carries a style and buys not hanging on a
  deck that does not.

  It also settles the intermittency from round 092: the manual clicks that
  answered in 428ms hit the same empty namespace and did not hang, so this is a
  RACE rather than a property of the empty case. The guard removes the call, not
  the race — worth remembering if it ever hangs at `count === 1`.

- **The honesty fixes are visible on the real host.** 096 came back **12/13**,
  skipping `the chart is actually visible` — the first honest reading that
  scenario has ever produced. 097 passed it. That difference is the fix working:
  the control render is available sometimes and not others, and the verdict now
  says which rather than reporting green either way. A lower score for a truer
  reason.

  `stallShape` also had its first outing, and correctly declined to over-read two
  stalls: `issued immediately after the previous answer — an ordinary sequential
  gap, says little`.

- **Grouping, third pair running: 19/0 then 15/4.** Pooled across three pairs the
  refusal counts are 0, 4, 0, 1, 0, 4. Neither rare nor constant, and invisible
  to a verdict that reports 13/13 either way.

## Rounds 098 + 099 — 3da2ef6 — a first-ever failure that the pair dissolved

      round  score  failed scenario                    scratch-unresolved  same scale  refused
      098    12/13  edit a chart on the visible slide          14            7 of 7        0
      099    13/13  —                                           0            8 of 8        4

Same build, nothing landed between them, and **none of round 098 replicated**.

`edit a chart on the visible slide` had passed 73 of 73 rounds and failed here for
the first time. The gate caught it and exited 1 — `had passed the previous 3
rounds running` — which is exactly its job. Then 099 passed it and the gate went
quiet again.

**The 14× spike is the round's real story, and it is weather.** `scratch slide
landed but its id will not resolve` and `could not remove the unusable scratch
slide` both ran 14 times against a baseline of 1, and both were 0 in the very next
round. The 7-chart population is downstream of that same churn — the probe charts
are discovered from a deck that fourteen unremovable scratch slides had polluted.

**I had a live reason to suspect my own change**, since 098 was the first round on
the count-first deck-style read. The pair says no: nothing in 099 differs except
the host's mood. That is the whole argument for pairing, and it is the second time
tonight a single round would have sent someone after a regression that does not
exist.

### The staked prediction failed, and the instrument was why

`counting-first-removes-the-hang` — FAILED. The probe line is still there, so the
read still fails with the guard in place.

**But the probe was lying, and it was mine.** Its `meaning` field was the fixed
sentence `getOnlyItemOrNullObject/load/getXml is what hung` — true while the read
was one batch, false the moment the read gained a `getCount` of its own. Round 098
printed it on a build where the count runs FIRST, describing a call it had not
observed. So the verdict refutes the prediction and not the diagnosis: the
diagnosis was never tested, because the instrument built to test it asserted its
answer.

The field now carries the operation the bounded sync names (`at=<what>`), read off
the actual failure. The count-first change is kept rather than reverted — it takes
one call off the path every pane load walks, costs a round trip only on a branded
deck, and nothing here shows it made anything worse. Its premise is simply
unconfirmed, and the next round can say.

### A note on what the gate can and cannot do

The gate judges the NEWEST round, so 098's alarm disappeared the moment 099
passed. That is correct for a regression gate and worth writing down anyway: an
unattended loop can erase its own alarm by continuing, and the only durable record
of round 098's failure is this entry and the archive. Not changed — a gate that
keeps firing about a round two back would be reporting history, and the pooled
per-scenario view already carries it.

## Rounds 100 + 101 — 504033c — the honest probe names the fault, and it is ONE call

Both 12/13, both skipping `the chart is actually visible`, and both reporting the
same thing byte for byte the first time the probe was allowed to measure rather
than assert:

    failedAt: "The value of the result object has not been loaded yet. Before reading the value"

**That is not a timeout.** `boundedSync` rejects on timeout, so this says the sync
RESOLVED and the value was still unloaded — a different fault from the 90-second
hangs, on the same line of code.

**Six rounds, three costumes, one fault.** Put side by side, what looked like two
separate bugs is one:

    089, 090   first call = getOnlyItemOrNullObject + load + getXml   hung, 90s
    096, 097   same                                                  hung, 10s
    098, 100, 101   first call = getCount (counting moved to front)  resolved,
                                                                     value NOT loaded
    the probe — always the SECOND call, fresh context                answered, always
    round 092, manual clicks on a long-lived pane                    428ms

**The first customXmlParts call after a pane loads does not work. The second
does.** The failure takes the shape of whichever call happens to be first, which
is exactly why moving `getCount` to the front changed the error message and not
the outcome.

And the diagnostic probe had been demonstrating the cure since the day it was
written, without anyone reading it that way: it opens a fresh context, asks the
same question, and has answered on every round it ran — immediately after the
read it was diagnosing had just failed.

So the read spends the bad call: attempt, and on failure attempt once more. Not a
timeout to tune.

- **Fix.** `attemptDeckStyleRead` twice. Guards proven red both ways — dropping
  the retry, and retrying without preserving `unreadable`, which would turn a
  genuinely dead read into a silent "no style" and is the exact lie that flag
  exists to prevent. The fake gained `refuseCustomXmlReadsOnce`, because a
  boolean cannot express "once" and without it a retry is indistinguishable from
  a read that gave up.

- **Instrument.** This is what the honest probe bought, on its first outing. The
  hardcoded `meaning` had been printing a conclusion about `getOnlyItemOrNullObject`
  on builds where that call no longer ran first.

- **The visibility skip is now in both rounds** where 096/097 split one-each. The
  control render's availability moves; the verdict reports which, which is all it
  was ever supposed to do.

## Rounds 102 + 103 — f3e3941 — the retry is taken back by the rounds that tested it

Both 12/13. The two scenario problems did not replicate — 102's
`insert onto a slide that already has content` and 103's visibility skip are one
each, which is weather. **The deck-style regime replicated exactly**, and it is
worse than the one before it.

    rounds     build      probe says
    096-099    pre-retry  the namespace IS reachable
    100, 101   504033c    the namespace IS reachable   (count-first, NO retry)
    102, 103   f3e3941    the namespace is UNREACHABLE too, at `counting the
                          deck's style parts`

**Two rounds on each side of one commit, and that commit is mine.** The probe had
answered on every round it ever ran — that was the whole evidence for "the second
call works" — and it stopped the moment the read began making two calls of its own
ahead of it.

The plain reading: the read's retry **spends the good second call**. The probe,
now third, gets the same nothing the first one got. A trace line appearing where
none did is one of the two readings this project treats as real, and this one has
a pair on each side.

**So the retry is reverted.** I wrote, one round earlier, "the retry stays; it is
simply not the whole cure" — that was too generous to my own change. It is not
merely not a cure: on this host it moved the failure from one call to the whole
surface, and on the interactive path it made a person wait two budgets instead of
one before being told.

What survives is the diagnosis it was built on — the first customXmlParts call
after a pane loads is the one that fails — and the instrument honest enough to
show the retry backfiring. Spending the bad call is still the right shape; doing
it INSIDE this read is not, because this read is what the probe measures.
Somewhere earlier in pane start, once, is the version worth trying next.

**And the ledger said HELD for round 102 before I read the round.** The claim
named one of the probe's two failure lines and the round produced the other, so
the named line was genuinely absent and the entry read as a cure while the read
had failed both attempts. A false HELD, in my own claim, of exactly the kind this
file spends its time finding elsewhere. `trace-line-present` takes a `scope` now,
so an absence claim can name the whole family, and the verdict says which of the
two it matched rather than quoting the one it did not use.

## Rounds 104 + 105 — 9b6a78f — the revert holds, and the diagnosis is now clean

    102, 103   WITH my retry     probe: the namespace is UNREACHABLE too
    104, 105   retry removed     probe: the namespace IS reachable   (both)

Two pairs bracketing one commit. The probe answers again on both rounds with the
retry gone, so it really was being starved of the good second call, and the
revert was not wasted motion.

**The pair also gives the cleanest statement of the fault the archive has
produced:** on both rounds the read fails at `counting the deck's style parts` —
its FIRST call — and the probe counts the same namespace successfully moments
later. First call fails, second works, twice, on one build.

105 skipped `the chart is actually visible` where 104 passed it; that is the
control render's availability moving, which is what the skip is there to report.
Grouping 20/0 then 17/2 — the sixth pair, and the refusal counts still swing on
one build.

- **Fix.** `warmCustomXmlSurface` spends the bad call before the read, in the
  pane's boot chain, where nothing waits on the answer and nothing else measures
  it. One thrown-away `getCount`. It cannot throw — failing is what it is for —
  and the guard for that is a mutation that rethrows, which takes the whole
  pane-load chain down.

- **Instrument.** The fake's once-fault only fired on the ITEM read, so it could
  not express "the first call, whatever it is" — and the warm-up sailed straight
  through it while the test failed for the wrong reason. It covers `getCount` now,
  which is what the real fault does.

- **The gate caught a real one.** Adding a renderer export the pane calls means
  adding it to the pane's MOCK in the same change: without it `app.ts` throws at
  boot and 110 tests die at once. Second time this session — the same thing
  happened with `readDeckStyleWithReason`.

## Rounds 106 + 107 — fa6448a — the first quiet pair, and why quiet is not an answer

    106   12/13   deck-style entries: 0
    107   13/13   deck-style entries: 0

**The first two rounds in this archive where the deck-style read recorded no
failure at all**, and the staked claim came back `held` on exactly that.

**It is weaker than it looks, and the claim was the wrong shape.** The probe
writes only when the read FAILS and the warm-up traces nothing, so an empty scope
is equally consistent with the read never having run — and the pane-load window
sits before the round's trace mark, so nothing in the file can settle it. A cure
and a silence have the same shape here.

That is this project's house defect pointed the other way: an absent failure line
read as evidence of success. Round 102 had pushed me toward claiming an ABSENCE
after a message-shaped claim misfired, and the right answer there was a POSITIVE
claim, not a broader negative one. Two mistakes about the same instrument, in
opposite directions, three rounds apart.

- **Fix.** The read records success as well as failure — `the deck-style read
  answered`, carrying the part count and the call it got past — and the verdict
  replays at round start like the failure one. A round can now SAY the read
  worked instead of leaving a gap that has to be interpreted.

- **Re-staked positively** as `the-read-answers-and-says-so`. HELD means the read
  is genuinely cured on this host and the round says so in a line that only
  exists if it completed. FAILED — `round starting` present, success line absent
  — means the read is not running in the pane-load chain at all, which would make
  this pair's quiet a worse story than a cure rather than a better one.

- **The warm-up is kept and is not yet credited.** It is cheap, it cannot throw,
  and nothing here argues against it. But two silent rounds do not show it is the
  reason they were silent, and saying so costs nothing.

Grouping, seventh pair: 20/0 then 17/2. 106 skipped `the chart is actually
visible` and 107 passed it — the control render moving again, which is what the
skip exists to report.

## Rounds 108 + 109 — f47a6d4 — the deck-style read is cured, and the round SAYS so

Both rounds carry `the deck-style read answered`, `parts: 0`, at `counting the
deck's style parts` — the call that had failed on every measured round since 089.
Positive evidence, twice, on one build.

**The whole arc, from the archive:**

    089, 090   —          hung 90s, no probe existed yet
    096-101    5 builds   probe: namespace reachable, read still failing
    102, 103   my retry   probe: UNREACHABLE too
    104, 105   reverted   probe: reachable again
    106, 107   warm-up    silent — cure or silence, indistinguishable
    108, 109   + success  THE READ ANSWERED

Twenty rounds. The fault was one sentence long the whole time: **the first
`customXmlParts` call after a pane loads does not work, and the second does.**

### What it cost to find, and why

Three wrong turns, each caught by a round:

1. **A retry inside the read** (#602), reverted by the next pair (#603). The read
   is what the probe measures, so a retry there spends the very call the
   measurement depends on. Two pairs bracketing one commit said so.
2. **A false HELD** — the claim named one of the probe's two failure lines and
   the round produced the other, so a genuine failure read as a cure.
3. **An absence claim** that could not tell a cure from a read that never ran,
   which is the same defect as (2) pointed the other way.

And the instrument only became useful when it stopped asserting: its `meaning`
field was a hardcoded sentence naming `getOnlyItemOrNullObject`, and it kept
printing that on builds where the count ran first. The moment it reported the
operation it had actually observed, the diagnosis fell out in one round.

**Nothing here fixes the host.** The bad first call is still bad; the warm-up
makes it the throwaway rather than the one a person waits on. Remove the warm-up
and the failure returns — that is what the guard pins and what 089-105 record.

Grouping, eighth pair: 20/0 then 18/2. Both rounds skipped `the chart is actually
visible`; the control render has been unavailable on this build across four
rounds now, which is worth watching but is exactly what the skip is for.

## Round 110 — 48cd380 — the cure holds, and the visibility gate's blindness has one cause

`the deck-style read answered` for the third consecutive round. The cure is not a
two-round coincidence.

**And the visibility gate's blindness is fully determined.** Across 15 rounds the
correlation is exact:

    blind rounds    8 (now 9)  — each with exactly ONE `rasterising a slide`
                                 stall at its full 20000ms budget
    sighted rounds  6          — zero stalls
    rasteriser actually unstable: 0 times in 85 rounds

The chart draws visibly in every one of them — the render moves ~1000 bytes each
time. What is lost is the CONTROL: the SECOND `getImageAsBase64` of the same
slide hangs. The first answered, which is where `before` came from.

Note the shape. The deck-style fault was the FIRST call failing and the second
working; this is the mirror — the first works and the second hangs. Two different
host surfaces, opposite orders, same class of fault.

- **Instrument.** The blind verdict names the stall now instead of shrugging "an
  unstable rasteriser cannot be ruled out" — true, but unhelpful when the archive
  has pinned the cause fifteen times. It says nothing when there is no stall to
  point at, which is the guard.

- **NO FIX SHIPPED, deliberately.** A retry is the obvious move and there is no
  pair behind it. The last obvious retry this session made things worse and the
  rounds took it back within two rounds — and this one would sit on the very
  call whose second attempt is the thing failing. It wants its own pair, not a
  05:00 guess.

  What that costs: the project's only mechanical evidence that a drawn chart is
  visible is unavailable in more than half of all rounds. Worth someone's next
  session.

## Round 111 — d56cf96 — 13/13 — the cold read never stopped failing

The first round carrying four new instruments, and one of them overturns a story
this project has been telling itself for twenty rounds.

**THE COLD RE-READ FAILS ELEVEN TIMES A ROUND.**

    cold re-reads that fell short   11   (1 short, 4 empty, 4 zero-match, 2 short)
    settle-delay retries fired      11
    post-retry failures              0
    charts grouped / refused      20 / 0

Every claim of the form "no re-read has come back short since the retry shipped"
was measuring **the retry's success rate**, not the host's behaviour. The fault
is exactly as common as it ever was; attempt 0 pushed the entry onto the retry
list and returned in silence, so nothing recorded it.

And there it is in the first line: `chart 4/8, kind: short, drew: 24, matched:
20` — the twenty-of-twenty-four case #586 was built for, still happening, on
every round, invisible until today.

**So #586's subset branch is starved because the RETRY NEVER FAILS**, not because
the host stopped producing short reads. That is a different fact with a different
response: the branch guards a regime this host enters constantly and is rescued
from every time. Keeping it is right; staking a prediction on it firing is not.

- **A successful rasterise costs under a second.** First durations ever recorded:

      the visibility BEFORE render    694ms
      the visibility CONTROL render   492ms
      the visibility AFTER render     803ms
      the rasterise-poisons arm       570-660ms
      an end-of-round slide shot      532-949ms  (n=6)

  Against a 20-second budget — twenty times the slowest of eleven. The budget can
  finally be sized from evidence rather than from harm, and this pair of rounds
  is the evidence. One round is a sample: pair it before cutting.

- **Zero stalls this round, and the visibility scenario passed** — no skip. The
  control render answered in 492ms. Consistent with the correlation: blind rounds
  have a stall, sighted rounds do not.

- **Four unsettled readings kept** that the old guard would have dropped in
  silence.

Every one of those four numbers was unmeasurable yesterday.

## Round 112 — 55562e1 — the pair, and the cold read replicates

    round  cold re-reads          retries  post-retry short/empty/zero  grouped/refused
    111    11 (3 short, 4 empty, 4 zero)  11       0 / 0 / 0                20 / 0
    112     8 (2 short, 3 empty, 3 zero)   8       0 / 1 / 1                18 / 1

**The cold read fails eight to eleven times a round, and that replicates.** The
correction stands: every "no short re-read since the retry shipped" was measuring
the retry's success rate, on a fault that never went away.

**And the pair sharpens it.** The retry does NOT repair everything — two of
round 112's eight survived it. But `short` is repaired **five times out of five**
across the pair, and `short` is exactly what #586's subset branch needs. So the
branch is starved by the ONE case the retry is best at, while the cases that do
survive (empty, zero-match) are the ones it cannot use.

That is a better answer than "the retry never fails", which is what one round
suggested. Two rounds say: the retry never fails *at the short read*.

**Trace and deck agree, again.** 112's trace reports one chart refused and its
deck reads `2,4,2,11,24,1,1` — one slide holding 24 loose shapes. 111 reports
none and its deck is all ones. The authority confirms the instrument.

**A rasterise is always under two seconds.** Pooled over the pair, n=22:

    min 492ms   p50 694ms   p90 1647ms   max 1835ms

The budget is 20000ms — eleven times the slowest ever observed. **Not cutting
it.** The measurement says the budget is generous, but cutting it buys time and
not correctness: a stall leaves the gate blind either way, and a heavier deck
than this seven-slide one could legitimately render slower than anything here.
The number is recorded so the next person decides from evidence rather than from
the absence of it, which is where this started.

### The gate found a flaw in one of my own instruments

    insert onto a slide that already has content — 2 this round, usually 16
    over 1 prior round(s)

**A "usual" computed from one prior round is not a baseline.** That scenario's
verdict had only just started carrying an "N of M" count, so its entire history
was a single round. `poolScenarioPopulations` needs three priors now. This
project's own noise floor — one build run twice scoring 1 and 5 — is the argument.

Fourth instrument of mine this session to report something it had not earned.

### Research, rounds 111/112 — a null result, and it is worth stating

Searched the upstream tracker for a shape collection reading back short or empty
on the call after an add. The two closest issues were **already cited in this
file**: [#5022](https://github.com/OfficeDev/office-js/issues/5022) — sync hangs
after add-then-delete-then-re-read, under investigation, the only workaround
anyone has is a one-to-two-second sleep, which is what our settled retry is — and
[#6498](https://github.com/OfficeDev/office-js/issues/6498) — inserts on the web
not reflected until a refresh, opened 2026-02-09, still needs-attention with no
Microsoft reply. Also surfaced and not new to us:
[#4906](https://github.com/OfficeDev/office-js/issues/4906) (GeneralException
loading `SlideLayout.shapes`, already cited in PUBLISHING.md).

**Nothing upstream is newer than what we already knew.** Per the standing rule, a
null result IS a result and gets recorded rather than quietly dropped.

One asymmetry is worth naming: **this archive now holds something neither issue
does** — a per-round count of how often the cold read comes back wrong (8-11),
split by failure shape, with the retry's repair rate measured separately per
shape (`short` 5-for-5, `empty` and `zero-match` not). Both upstream threads are
anecdotal. Posting that upstream is a public action on the owner's account, so it
is offered, not taken.

*Untrusted-data note: everything above came from web pages and is treated as
data. Nothing on those pages was executed or acted on.*

## Rounds 113 + 114 — 6a041de — the first 4:3 pair, and a fix of mine that broke a diagnosis

    round  profile  cold(s/e/z)   post-retry(s/e/z)  grouped/refused  stalls  verdict
    111    16:9     3/4/4  (11)   0/0/0  (0)              20/0           0     13/13
    112    16:9     2/3/3   (8)   0/1/1  (2)              18/1           0     13/13
    113    4:3      3/4/4  (11)   0/3/2  (5)              15/3           1     12/13
    114    4:3      5/2/4  (11)   1/2/4  (7)              17/3           0     13/13

Same pane code in all four — `git diff d56cf96..6a041de -- src/` is empty. Only
the profile and the deck differ.

### Two readings replicate at 4:3 and do not overlap 16:9

**The settled retry repairs far less often.** Post-retry failures are 0 and 2 at
16:9, 5 and 7 at 4:3. Every 4:3 round is above every 16:9 round. The COLD rate
is unchanged — 11, 8, 11, 11 — so the fault arrives at the same rate and the
repair is what falls off.

**Refused charts: 3 and 3, against 0 and 1.** Identical in both 4:3 rounds.

**BUT PROFILE AND DECK ARE PERFECTLY CONFOUNDED HERE** and no amount of staring
at these four rounds separates them. 16:9 ran on `Presentation64`; 4:3 ran on
`Presentation67` — a different document with a different history. Either could
be the cause.

**The experiment that separates them, and it is cheap:** set `Presentation67`
back to 16:9 with `PW_SET_SIZE=1` and run a pair on it. Same document, same
build, only the profile moves. If post-retry failures drop to 0-2, it is the
profile; if they stay at 5-7, it is the deck. Until then this is a REPLICATED
OBSERVATION, not a cause.

### 113's failure dissolved on its pair, and the trace says why

113 reported `the chart is actually visible` as its one failure; 114 passed it.
113 carries exactly one stall and 114 carries none, and the stall is the
scenario's own control render:

    gave up waiting  what=the visibility CONTROL render (same slide, back to back)

So the chart was almost certainly fine and the RASTERISER stalled — which is
what the pair is for. Nothing to fix in the chart path.

### The instrument had that answer and printed a shrug

113's verdict text read `(no control render, so an unstable rasteriser cannot be
ruled out)` — the branch for when the cause is UNKNOWN. The branch that names
the stall exists two lines above it and did not fire, because it tested:

    lastStall && /rasteris/i.test(lastStall.what)

**That worked only while every rasterise traced the one string "rasterising a
slide" — and I am the one who stopped that being true.** Giving each call site
its own label was a real improvement; it also silently broke the consumer that
matched on the old wording. The control render now stalls as "the visibility
CONTROL render (same slide, back to back)", which contains no "rasteris".

Fixed by marking the operation instead of describing it: `RASTERISE_OP` travels
beside the label, on the stall and on the success line, and both consumers match
the token. Prose for people, a token for code.

**And the test froze the old world.** `selftest.test.ts` asserted against
`what: "rasterising a slide"` and went on passing while production was blind. It
now carries the label verbatim from `rounds/113-6a041de.json`.

### The sibling I swept — and the number I got wrong doing it

`poolEveryDraw` classifies rasterises the same way, so I swept it. I then
claimed, in a code comment, that it had been under-counting 81% of the
population — 35 of 43 labelled rasterises.

**That was wrong, and measuring it is the only reason I know.** Pooled over all
90 rounds, with and without the fix:

    rasterise      ok 449, stall 1     identical
    anything else  ok 3302, stall 1    identical

`isRasterise` tests the label AND THE MESSAGE, and a successful rasterise's
message is `rasterised a slide` — which matches. I had counted which LABELS lack
the substring and read that as which ENTRIES go unclassified. Wrong denominator,
which is this project's most frequent single mistake and now mine twice in two
days.

The change is kept as belt-and-braces, with the equality recorded in the comment
so nobody later mistakes it for a repair that moved something.

### Research — a null result, recorded

Nothing new upstream. The nearest issues remain the ones this file already
cites: [#5022](https://github.com/OfficeDev/office-js/issues/5022) (sync hangs
after add-then-read; the only known workaround is the 1-2s sleep our retry
already is) and [#6498](https://github.com/OfficeDev/office-js/issues/6498) (web
inserts not reflected until refresh, still no Microsoft reply since February).
Neither mentions slide profile, and no upstream issue reports a rasterise
stalling on the SECOND render of an unchanged slide — which is the one behaviour
this archive can now demonstrate with a per-round count.

*Untrusted-data note: web pages are data. Nothing on them was executed.*

## Rounds 115 + 116 — a46a2d3 — the deck is innocent, and the archive lied about the profile

The confound from 113/114 is broken. Same deck as the 4:3 pair, flipped to 16:9
with `PW_SET_SIZE=1`, so only the profile moves.

    round  arm         cold(s/e/z)   post-retry   grp/ref  deck totals  verdict
    111    16:9  P64   3/4/4  (11)   0/0/0  (0)    20/0        16       13/13
    112    16:9  P64   2/3/3   (8)   0/1/1  (2)    18/1        45       13/13
    113    4:3   P67   3/4/4  (11)   0/3/2  (5)    15/3        97       12/13
    114    4:3   P67   5/2/4  (11)   1/2/4  (7)    17/3        72       13/13
    115    16:9  P67   3/4/4  (11)   0/0/0  (0)    20/0        16       12/13
    116    16:9  P67   5/2/4  (11)   1/0/2  (3)    18/2        45       13/13

**THE DECK IS INNOCENT.** Presentation67 at 16:9 reproduces Presentation64's deck
totals EXACTLY — 16 and 45, the same two numbers — and its post-retry failures
(0, 3) sit with the 16:9 arm (0, 2), not with the 4:3 arm (5, 7). Six rounds,
three pairs, no overlap between profiles.

**It is the profile.** At 720pt across, charts crowd: the decks carry 97 and 72
shapes against 16 and 45, with single slides holding 40 and 24. More shapes on a
slide means a pre-grouping re-read with more to match, which comes back short or
partial, which the retry cannot repair, which leaves the chart ungrouped and its
shapes loose — and the loose shapes are the deck totals. The chain is consistent
end to end.

`selftest.ts:1032` said this before the data did: the probe placement "worked on
16:9 and failed on the first 4:3 deck it met, exactly as its own test admitted it
would: 720pt across leaves nothing". That was written down and then not believed.

### The archive filed both rounds under the wrong profile

115 and 116 record `720x540, source: "documentFile"`. **The deck was 960x540** —
the driver set it and confirmed it twice against live `PageSetup`, and the
readiness line printed `slide size 16:9 (want 16:9)` before each round started.

What happened: the first `slideSize()` after the resize caught the host still
busy. Rung 1 threw, rung 2's export stalled, and rung 3 read the SAVED FILE,
which PowerPoint had not yet written the new size into. That fallback was then
CACHED, so one unlucky moment became the round's permanent answer.

**`PW_EXPECT_SIZE` cannot catch this.** The guard reads live `PageSetup`; the
archive records whatever rung the pane happened to reach. Two numbers, different
sources, never compared — so the guard passed while the round filed itself under
a profile it was not measured at, which is the exact failure the guard exists to
prevent.

**The remedy was already written and never wired up.** The comment on
`cachedSlideSize` names this hazard — "a cached 16:9 would then place charts off
the edge of a deck that is now 4:3" — and `slideSize({ refresh: true })` and
`_resetSlideSizeCache()` were built for it. `refresh: true` appears NOWHERE in
the codebase except that comment; `_resetSlideSizeCache` is called only by
`office-render.test.ts`. And `PW_SET_SIZE`, which I added, creates that exact
condition programmatically before every profile-flipping round.

**Fixed structurally rather than by remembering to call something.** Only
`pageSetup` and `exportedSlide` — readings of the LIVE deck — may end the ladder
for good. `documentFile` and `assumed` are still used, still cheap to reuse, but
no longer final: the cheap rung gets another chance on the next call and a
recovered host upgrades the answer. The expensive rung is not re-run to re-learn
what is already held.

A fallback is what you take when the measurement is unavailable. Caching it as
though it were the measurement is this project's oldest mistake in a new hat.

### What this does NOT establish

115 and 116 ran with the pane believing 720 while the deck was 960. That is a
third condition, not a clean 16:9 arm — and it happens to be a useful control:
the pane's BELIEF was 720 across all four P67 rounds while the actual geometry
varied, and the counters tracked the geometry. So the conclusion survives, but
the clean re-run belongs on a build that records what it measured.

### CORRECTION to the 115/116 entry above, same day, before anyone builds on it

Two claims in that entry are wrong. Both were mine, both were confident, and
neither was measured before I wrote it.

**1. The causal chain is backwards.** I wrote that more shapes on a slide means a
re-read with more to match, which comes back short, which leaves the chart
ungrouped. The arrow points the other way. Refused charts and dense slides are
the SAME EVENT counted twice — an ungrouped chart simply leaves its shapes loose:

    round  refused  slides holding >=10 shapes
    113       3          3
    115       0          0
    116       2          2
    112       1          2

The deck total is a CONSEQUENCE of refused grouping, not a cause of the re-read
failing. So the observation stands — 4:3 refuses more and the retry repairs less
— but **the mechanism remains unexplained**, and the entry above claimed to have
explained it.

**2. `selftest.ts:1032` did not predict this.** I quoted it as having said "720pt
across leaves nothing" before the data did. That comment is a HISTORY of the
first placement fix — a fixed right-hand column — explaining why it was replaced
by a band computed from the occupied box, which works on both deck shapes. It
describes a bug that was already fixed. I read a repaired defect as a standing
prediction, which is the most flattering possible misreading and should have
been checked against the code before it went into a commit message.

What is actually true: the 4:3 arm refuses more charts and repairs fewer
re-reads, replicated across a pair, with the deck exonerated. Why, is open.

## Rounds 117 + 118 — 796605c — the second round of a pair is worse, and always has been

**117 refuted the 4:3 claim I published this morning, and 118 rebuilt it smaller.**

    round  arm       post-retry  grp/ref  deck  position
    111    16:9 P64      0         20/0     16   1st
    112    16:9 P64      2         18/1     45   2nd
    113    4:3  P67      5         15/3     97   1st
    114    4:3  P67      7         17/3     72   2nd
    115    16:9 P67      0         20/0     16   1st
    116    16:9 P67      3         18/2     45   2nd
    117    4:3  P67      0         20/0     16   1st
    118    4:3  P67      8         16/4     95   2nd

117 is 4:3 and scored 0 post-retry, 0 refused, deck 16 — the 16:9 signature
exactly. "The profile degrades the retry, replicated, deck exonerated" does not
survive its third round at that profile.

### What was actually there: POSITION IN THE PAIR

Pooled over every build this archive has run twice — 31 pairs, not the eight
that suggested it:

    second round had MORE post-retry failures : 15   (tied 14, better 2)
    second round left a BIGGER deck           : 15   (tied 10, better 6)

**Of the 17 pairs that moved at all, the second round is worse in 15.** The ties
are mostly older rounds whose counters sat at zero both times; they are not
evidence of symmetry, and counting them as coin flips is how this nearly got
waved away as noise.

**Nothing in this project's method controls for it.** The pair exists to
separate a real fault from the host's mood, and the two halves are read as two
samples of one condition. They are not: the second is systematically degraded.
Every "one build run twice" noise-floor figure in these docs is measuring
position as well as noise, and every comparison that reads a 1st-round number
against a 2nd-round number is comparing different conditions.

### The profile effect survives, smaller and better bounded

Split by position, it is still there — and only in second rounds:

    2nd rounds   16:9  2, 3        4:3  7, 8      no overlap
    1st rounds   16:9  0, 0        4:3  5, 0      overlapping

So: **the second round is worse everywhere, and at 4:3 much worse.** That is a
weaker and more honest claim than the one it replaces, which pooled both
positions and read the compounded effect as a clean profile difference.

113 remains an outlier among first rounds — 5 post-retry and a 97-shape deck
against 0, 0, 0 and 16, 16, 16. Unexplained.

### The cross-check fired for the first time, and agreed

117 and 118 are the first rounds to carry `driverSlideSize`. Both record
`4:3` from the driver against `720x540` from the pane — agreement, stated
rather than assumed. The guard built this morning for the 115/116 failure now
has live data proving it can also say "these match".

### What is worth doing next, and it is not another 4:3 round

The position effect is the bigger prize and it is cheap to attack: run one build
THREE times. If the third is worse again, something accumulates across rounds
within a browser session — the proxy-memory exhaustion this project already
knows about is the obvious suspect. If the third matches the first, the sweep
between rounds is leaving something behind that one more sweep clears.

## Rounds 119-121 — b7c4196 — one build three times, and the observer is the variable

**The three-run experiment killed the accumulation theory, and then killed the
position theory that replaced it.**

    round  span    post-retry  grp/ref  deck
    119     758s        0        20/0     16
    120    1830s        7        16/4     91
    121     800s        1        19/0     22

**Not cumulative.** 121 ran on a pane that had been through two rounds, with no
recovery or reload between them, and came back at first-round speed and
first-round counters. Anything that "builds up over a session" is refuted.

**And position does not determine it either.** The archive's two earlier
three-run builds go the other way — 070/071/072 scored 1, 0, 6 and 079/080/081
scored 0, 0, 4, both spiking on the THIRD. Ours spiked on the second and
recovered. Same build, same deck, same driver: the position is not the cause.

### What actually tracks it is DURATION, and the pairs are unanimous

    pair          1st      2nd      ratio   post-retry
    113 -> 114   1249s -> 2488s     2.0x     5 -> 7
    115 -> 116    877s -> 2115s     2.4x     0 -> 3
    117 -> 118    912s -> 1991s     2.2x     0 -> 8
    119 -> 120    758s -> 1830s     2.4x     0 -> 7

Four pairs, four times slower, every time by roughly the same factor. Pooled
over the 30 rounds whose instruments existed, the slower half averages 5.1
post-retry failures against 2.8, and 80 deck shapes against 48.

### THE VARIABLE IS PROBABLY ME, AND THAT IS NOT A JOKE

There is one thing that reliably differs between the first and second round of a
pair on this machine, and it is not the host. **I launch round one and wait. I
mine round one WHILE ROUND TWO IS RUNNING** — parsing every file in a 120-round
archive, spawning triage, and during round 118 running the full 3,195-test suite
repeatedly to chase a flake. 118 is the worst round in the set. 121 ran while I
was writing a short script and doing nothing heavy, and it came back clean.

This is a hypothesis, not a proof: the load was never recorded, so it cannot be
recovered from the archive. But it explains every observation the other two
theories could not, it predicts the earlier three-run builds spiking on
whichever round happened to coincide with work, and the mitigation costs
nothing.

**`docs/ROUND-LOOP-BRIEF.md` has one rule about not disturbing a running round:
do not touch `playwright-cli` while the driver polls. That rule is too narrow.**
It protects the CLI's single-command channel and says nothing about the machine
the browser is running on. A 16-worker Vitest run is not "touching
playwright-cli" and is far more disruptive than a `find` would be.

### What this costs, and what to do

Every pair in this archive may be a fast round and a loaded round rather than
two samples of one condition — which is exactly the objection I raised against
position yesterday, one level deeper. The "second round is worse, 15 of 17"
finding stands as a STATISTIC and its causal story is now the third candidate in
two days.

- **Do nothing heavy while a round polls.** Mine the previous round before
  starting the next one, or after both land.
- **Record the round's own span in the receipt** so a future reader can tell a
  slow round from a fast one without recomputing it. It is already in the
  trace; nothing surfaces it.
- **Stop treating a pair as two samples of one condition** until a pair has been
  run with the machine deliberately idle for both halves.

## Rounds 122 + 123 — b321c65 — the idle pair refutes "the observer is the variable"

Run deliberately with nothing else happening on this machine: no mining between
the two, no triage, no test suite. The whole point was to remove the one
difference I had identified between a first and a second round.

    round  condition   span    post  grp/ref  deck
    122    1st, IDLE    862s      0    20/0     16
    123    2nd, IDLE   2076s      2    18/2     62

**2.4x slower, with the machine idle.** That is the same ratio as every loaded
pair — 2.0x, 2.4x, 2.2x, 2.4x — so THE SLOWDOWN IS NOT CAUSED BY MY LOAD. The
hypothesis published in #624 an hour earlier is refuted by the experiment it
asked for, which is the best thing that can happen to a hypothesis.

**Three theories down in one day:**

1. *Damage accumulates within a session* — killed by 121, which ran on the
   oldest pane of three and came back clean.
2. *Position in the pair determines it* — killed by the archive's earlier
   three-run builds, which spike on the third instead.
3. *The observer's load causes it* — killed here.

**What survives, and it is worth stating precisely.** The second round of a pair
is 2.0-2.4x slower than the first in FIVE of five pairs measured, idle or not.
That is now a fact about the HOST, not about the machine or the method.

Load may still MODULATE severity rather than cause it: the idle second round
scored 2 post-retry failures where loaded second rounds scored 7, 8, 3 and 7,
and left 62 deck shapes against 45, 72, 91 and 95. One observation, so it is a
lead and nothing more — but it is the difference between "my analysis invalidates
the round" and "my analysis makes a real effect worse", and those call for
different rules.

**So the brief's second rule stays, with an honest reason.** "Do nothing heavy
while a round polls" is still right — it is cheap, and the one idle second round
is the mildest second round on record. What must change is the CLAIM attached to
it: it does not prevent the slowdown, and anyone who obeys it expecting fast
second rounds will conclude the rule does not work.

**What is actually unexplained now:** why a round that follows another round is
half as fast, on a swept deck, on a pane that has not been reloaded, with no
load on the box — and why the third round recovers (758 -> 1830 -> 800). Nothing
in this project's model accounts for that, and the next experiment should
measure the HOST rather than the harness: whether the slowdown is in the
`context.sync()` round trips (service-side throttling) or in the pane's own
work.

## It was the PANE'S AGE all along — and the metric that hid it was mine

Asked to find where the extra time goes, the answer turned out to be that there
was no extra time. **`ms` counts from the PANE'S load, not the round's start, and
the pane is not reloaded between rounds.**

    round 122   first entry     67,116ms   last     862,019ms
    round 123   first entry    961,782ms   last   2,075,549ms

123's clock starts at 961s because it inherited 122's pane. `roundSpanSeconds`
took `Math.max(...)`, so every reused-pane round's "duration" silently included
the previous round's. Corrected, last-minus-first:

    pair          published   actual
    113 -> 114      1.99x      1.26x
    115 -> 116      2.41x      1.43x
    117 -> 118      2.18x      1.38x
    119 -> 120      2.41x      1.36x
    122 -> 123      2.41x      1.40x

**Every ratio in #624 and #625 is wrong.** The real effect is about 1.35x, not
2.0-2.4x, and I published the inflated figure twice without checking what the
number measured.

### And the inflated metric was accidentally encoding the real cause

Split rounds 110-123 on the pane's age when the round STARTED:

    fresh pane (<200s)   110, 112, 115, 117, 119, 121, 122
                         post-retry 0, 2, 0, 0, 0, 1, 0     deck mostly 16
    reused pane          111, 113, 114, 116, 118, 120, 123
                         post-retry 0, 5, 7, 3, 8, 7, 2     deck 45-97

Mean post-retry **0.43 against 4.57**. Deck 16 against 60+.

**"The second round of a pair is worse" was always this.** The second round is
the one that inherits a pane. Position, profile and observer load were three
stand-ins for a variable none of them named.

### Which means theory 1 was right, and I killed it with a broken instrument

    121   pane age at start: 71s   FRESH

121 is the round I used to refute "damage accumulates within a pane session" —
"it ran on the oldest pane of three and came back clean". **It did not. Its pane
was 71 seconds old.** It never tested accumulation, and the refutation was
worthless. `docs/ROUND-LOOP-JOURNAL.md` has carried that claim since this
morning and every conclusion built on it inherits the error.

The idle pair's finding survives intact and is now explained: 123 was slower and
worse than 122 with the machine idle **because it inherited a pane**, not because
of anything the observer did.

### The fix is cheap and the driver already knows how

`recover` reloads the tab and reopens the pane. Rounds should do that BETWEEN
runs rather than only sweeping the deck — a sweep clears the slides and leaves
whatever the pane accumulated. That is one call in `collectRound` and it turns
every round into a fresh-pane round.

Not shipped in this commit, deliberately: it changes what every future round
measures, and it deserves its own pair — run on a fresh pane both times, which
is exactly the thing this archive has never done on purpose.

### The instruments

`roundSpanSeconds` is last-minus-first now. `paneAgeAtStartSeconds` is new and is
printed by the gate above the counters, because it is the single best predictor
of what those counters will say. Null when unreadable, never 0 — the gate treats
under 200s as fresh, so a 0 for "unknown" would mark every unreadable round
clean, which is the same defect in a third costume.

## Rounds 124-126 — the fix works: a second round that scored a first round's numbers

    round  pane age   span   post  grp/ref  deck  verdict
    120      889s     942s     7    16/4     91   12/13   reused
    123      962s    1114s     2    18/2     62   13/13   reused
    124       68s     719s     0    20/0     16   13/13   FRESH
    125       68s     691s     0    20/0     16   13/13   FRESH
    126      119s     802s     0    20/0     16   13/13   FRESH

**126 is a SECOND round, and it is the first second round in this archive to
score a first round's numbers.** post-retry 0, 20 grouped and 0 refused, a
16-shape deck, 13 of 13. Every previous second round scored 2, 3, 7, 7 or 8 with
decks of 45 to 95.

Reloading the pane between rounds removes the effect. **Pane age was the cause**,
and the intervention that follows from it works on the first attempt.

Worth stating what that closes: the same observation was published as a property
of the SLIDE PROFILE, then of POSITION IN THE PAIR, then of the OBSERVER'S LOAD,
before anyone measured the pane's age — and the measurement only happened
because the metric that hid it (`max(ms)` read as duration) was itself wrong.
Four wrong answers, each internally consistent, from one unexamined number.

### Shipping the fix cost three defects, all of them mine

1. **A reload raises a beforeunload modal**, and a modal makes every
   playwright-cli command fail in a way that impersonates a dead browser. Only
   `screenshot` names the real state.
2. **`refreshPane` looked once after a fixed sleep** — the third instance of that
   exact shape in one day, in the function written to make the next round clean.
3. **The wiring was untested.** Deleting the call from `collectRound` left the
   suite green; mutation testing caught it and forced a real guard.

### And an instrument that would have kept announcing a fixed problem

The gate's PAIR POSITION line pools all history, so it will report "the second
round was worse 17x" for as long as the archive exists. That is true and it is
no longer ADVICE — it now names the cause, says the reload shipped, and counts
how many pairs had a fresh second round (14 of 34 today, rising).

A statistic about a fixed problem is exactly the kind of instrument this project
keeps having to correct: it goes on being true and stops being useful.

## Rounds 127 + 128 — the 4:3 profile effect was never real

The first profile comparison this project has ever made properly: same deck,
machine idle, and — for the first time — **every round started on a fresh pane**,
so a pair is two comparable rounds instead of a clean one and a degraded one.

    round  profile  pane-age  span   cold  post  grp/ref  deck  verdict
    124     4:3        68s    719s    11     0    20/0     16   13/13
    125     4:3        68s    691s    11     0    20/0     16   13/13
    126     4:3       119s    802s    11     0    20/0     16   13/13
    127     16:9       77s    839s    10     0    20/0     16   13/13
    128     16:9       72s    651s    11     0    20/0     16   13/13

**Identical on every counter.** post-retry 0 across the board, 20 grouped and 0
refused, a 16-shape deck, 13 of 13.

**"4:3 degrades the retry" was entirely an artefact of which rounds happened to
inherit a pane.** It was published as replicated, deck-exonerated, and
mechanistically explained; the mechanism was invented to fit a difference that
did not exist. Fourth theory dead, and the only one that survived contact with a
controlled experiment is pane age.

### What IS real, and profile-independent

The cold re-read fails **10 to 11 times a round in all five**, at both profiles,
on a fresh pane, with everything else clean. That fault has replicated across
every condition this project has tested and is the one finding from the whole
sequence that nothing has undermined.

### What I nearly claimed and did not

All five rounds got a control render, where the visibility gate has been blind in
38% of the archive. That looked like a second win for the pane reload. It is not:

    fresh panes    27 blind / 51 sighted    35% blind
    reused panes   13 blind / 13 sighted    50% blind

Pane age is ASSOCIATED with blindness and does not explain it. Five sighted in a
row on fresh panes has roughly a 12% chance at the historical 35% rate — which is
unremarkable, not evidence. The visibility gate's blindness remains unexplained
and is the obvious next target now that the confound underneath everything else
is gone.

## The visibility gate's blindness has one cause, and it is 40 for 40

The gate is blind — no control render, so a before/after difference proves
nothing — in 38% of the archive. Pooled over all 104 rounds that carry the
scenario:

                             blind      sighted
    any rasterise stall     40 100%     0   0%

**Perfect separation.** Every blind round carries a rasterise stall; no sighted
round does. The code claimed this from 14 rounds (8 blind / 8 stalls, 6 sighted
/ 0); it holds at 40 and 64.

And the stall trace says the same thing all 40 times:

    WHICH CALL STALLED             WHAT THE HOST LAST ANSWERED
     34x rasterising a slide        34x rasterising a slide
      6x the visibility CONTROL      6x the visibility BEFORE render

Two rows, one event — the labels only became specific this week. The stall shape
adds `issued immediately after the previous answer`.

**So: the host answers the first rasterise of a slide and then hangs on a second
rasterise of the SAME slide issued straight after it.** Not slowly — for the full
20s budget, never returning. A successful render takes 941ms at the median and
1466ms at its worst, so a call outstanding at 20s is hung, not slow.

### The fix, and it is a candidate rather than a cure

`rasterGap()` — three seconds before the control render, twice the slowest
render that has ever come back.

Nothing upstream describes this shape. The nearest are
[#5022](https://github.com/OfficeDev/office-js/issues/5022) (sync hangs after
add-then-read, already cited here) and
[#6266](https://github.com/OfficeDev/office-js/issues/6266) (a Mac/Windows
rendering difference, not a hang). The only guidance that exists for the family
is to space the calls out, which is what this does.

**The rounds after this are the test, and the baseline to beat is 38%.** Five
clean rounds would not settle it: at a 35% blind rate on fresh panes, five
sighted in a row happens 12% of the time by chance. That is the same arithmetic
that stopped me claiming the pane reload had fixed this yesterday, and it applies
to my own fix too.

### A test that cannot see the thing it guards

The gap is zeroed in `test/setup.ts` — a fake rasteriser cannot hang, and three
real seconds per scenario put four of them over budget — so no behavioural test
can observe this call at all. Deleting it leaves the suite green, exactly as
deleting the pane refresh from `collectRound` did.

The guard therefore reads the source and says so: it catches the call being
removed or reordered, not the gap being the wrong length. Second time this week
that mutation testing has caught a shipped behaviour with no guard behind it,
and both times the behaviour was the whole point of the change.

## Rounds 129-132 — the gap looks right, and the reload that carried it had to be pulled

**The four rounds that ran are clean, and so is the mechanism:**

    round  pane-age  visibility  rasterise-stalls  control-render
    129       66s    sighted            0            413ms
    130       93s    sighted            0            361ms
    131       61s    sighted            0            503ms
    132       65s    sighted            0            444ms

Against a 38% blind archive. **Not proof:** four sighted in a row has an 17.9%
probability at the 35% fresh-pane baseline, and the batch was meant to be seven.
The 40-for-40 correlation between blindness and a rasterise stall still holds —
zero stalls, zero blind.

The control render now returns in 361-503ms, faster than the 941ms median of the
sighted rounds before it. That is what a call that is not fighting for the
rasteriser looks like.

### THE BATCH DIED AT ROUND 3, AND THE CAUSE WAS MY OWN FIX

    batch round 3   the pane did NOT reopen
    batch round 4   this document has no PowerChart command in its ribbon
                    the deck holds 79 slides

One mechanism explains both. A reload of a tab with unsaved work raises
PowerPoint's beforeunload prompt, and **`dialog-accept` means LEAVE WITHOUT
SAVING**. The deck sweep immediately before it GUARANTEES unsaved work — it has
just deleted slides — so accepting discards them, and the deck comes back. It
appears to discard the per-document SIDELOAD too: the add-in was gone from the
ribbon after the reload following round 124, and again after 132.

**So the between-rounds reload is out.** It worked — 126 was the first second
round in this archive to score a first round's numbers — and it cost the session
roughly one round in four, on the one failure only the owner can reliably undo.

Dismissing the dialog instead would cancel the reload rather than save it. The
real fix is to WAIT FOR THE AUTOSAVE and reload a clean document so no prompt
appears at all, which needs a way to read PowerPoint's saved state that this
driver does not have — and which must not be guessed at with a sleep on the one
path whose failure costs the add-in.

`refreshPane` stays, because `recover` needs it and is correct there: the page it
reloads has usually already crashed and has nothing left to save.
`paneAgeAtStartSeconds` stays too — the gate still reports whether a round
inherited a pane, which is what makes its counters readable even when nothing
can be done about it automatically.

### What this leaves

- **The gap is the live candidate** for the blind gate, at 4 rounds and 17.9%.
- **Pane age is diagnosed, measured, reported — and no longer auto-corrected.**
  A person can reload the pane by hand between rounds; the driver will not.
- The guard that asserted the reload was wired in now asserts the opposite, and
  says what it costs, so it cannot drift back unnoticed.

## `sideloadAddIn`: two successes, one failure, and one real fix

Its record before today was "has never once succeeded on a real host". Today it
put the add-in back twice unaided, and failed once — the failure being the
current state of `Presentation67`, which needs a hand sideload.

**The fix with evidence behind it: reload before walking.** Twice today the walk
died at `the Add-ins menu did not open`, and both times the cause was a
`button "Add-ins" [expanded]` left over from a PREVIOUS failed attempt. Clicking
an open menu closes it, so every retry toggled the menu instead of walking it —
and the driver reported the ribbon as unopenable when it was merely already
open. A first failure therefore guaranteed every later one.

**Escape is not enough, and that is measured rather than assumed.** It returned
the button to `[active]` and the very next attempt still failed; only a reload
let the walk reach the upload, both times.

The reload can discard unsaved work, which is exactly the fault that got the
between-rounds reload pulled an hour ago. Acceptable HERE and nowhere else: this
path runs only when the add-in is already missing, and a document with no add-in
cannot produce a round whose work would be worth saving.

### What is NOT fixed, and why I am not touching it

`SIDELOAD_COMMAND_BUDGET_MS` is 90s and the command has twice taken longer. The
obvious move is to raise it — and the evidence refuses. On 2026-08-20 the
command surfaced roughly fourteen minutes after the upload; today it never
surfaced at all, through twelve minutes of polling and a further reload. A budget
long enough to catch the first case makes a genuine failure take a quarter of an
hour to report, and this one WAS a genuine failure.

So the number stays and the message carries the caveat instead: "the upload may
still land, so re-check the ribbon before sideloading by hand". A wait cannot
distinguish a slow success from a failure; only looking again can.

## The add-in was never missing: a narrow window hid it, four times

PowerPoint collapses trailing ribbon commands into a `...` overflow when the
window is narrow, and **a collapsed command is not in the accessibility tree**.
So the ribbon read answers `false` — truthfully, it is not rendered — and every
reader above it concludes the add-in is gone.

Measured 2026-08-21, same browser, same document, seconds apart:

    page width 1237   Insert chart: false   "the add-in is not loaded here"
    page width 2375   Insert chart: true    round runs

On that reading the driver ran the sideload walk **four times** against a deck
that already had the add-in, refused four rounds, and I told the owner to
sideload by hand. Nothing was ever missing, and two "fixes" — the pre-walk
reload and the document wait — were shipped for a failure that was never the
walk's fault.

**It became reachable through my own change.** The round browser moved to
`viewport: null` so the page would follow the window, which fixed the pane being
invisible to a PERSON and made commands invisible to the DRIVER. The window is
not a measurement surface; nothing in a round's result depends on how wide the
tab is. So it must not be allowed to vary.

### Two fixes, and the second is the one that matters

**`ensureRibbonRoom` widens the window** before any ribbon read, and reports the
width it actually measured rather than the one it asked for — the relationship
between the two depends on the device scale factor, which varies by machine and
by monitor.

**`ribbon-cramped` replaces `addin-missing` when the window is too narrow to
judge.** It is RECOVERABLE, unlike `addin-missing`: widening a window changes
nothing a round measures, so the driver corrects it instead of stopping. And a
width that could not be READ counts as too narrow, because a reader that cannot
measure the window cannot tell a missing add-in from a hidden one.

The real refusal survives: with room to render and still no command, the add-in
genuinely is not there, and that stays a hard stop.

### The pattern, stated plainly because I keep repeating it

Three times in two days I have read ABSENT FROM THE ACCESSIBILITY TREE as ABSENT
FROM THE PRODUCT: the `find` miss message matching a bare-name pattern, a
disabled control carrying no ref, and now a command collapsed into an overflow.
Each time a SCREENSHOT settled it in one call, and each time I reached for it
only after exhausting the text queries that cannot distinguish the two states.

The tree says what is rendered. It does not say what exists.

## Rounds 129-141 — the blind gate is fixed, and pane age is dose-dependent

**THE BLIND RATE IS SETTLED. 0 of 13.**

    before the gap   fresh panes  27 blind / 51 sighted = 35%
                     reused panes 13 blind / 13 sighted = 50%

    after the gap    fresh   0 blind of 8
                     reused  0 blind of 5

    If the gap changed nothing, this run has probability 0.10%.

Zero rasterise stalls across all thirteen, and the blindness/stall correlation
never broke — it stands at 40 blind rounds with a stall, 77 sighted without one.
`rasterGap` — three seconds before re-rasterising a slide the host has just
rasterised — closed the fault that made this project's only mechanical evidence
of a visible chart unavailable in 38% of rounds.

The five reused-pane rounds matter most: at a 50% historical rate, those alone
are a 3% coincidence. This is not a fresh-pane artefact.

### Pane age is dose-dependent, and yesterday's split understated it

    round   pane age   post-retry   grouped/refused
    136        844s         6           16/3
    138        891s         7           16/4
    139       1925s         8           15/4
    140       3037s         9           12/5
    141       4283s        10           10/9

**Monotonic on every counter.** Post-retry failures climb 6, 7, 8, 9, 10 as the
pane ages from 14 minutes to 71; grouped charts fall 16, 16, 15, 12, 10; refused
climb 3, 4, 4, 5, 9. Fresh-pane rounds in the same batch average 1.13 post-retry
against 8.00 for reused.

**I misread this mid-batch and said so at the time.** Seeing 139 and 140 score
13 of 13 on a 32- and 50-minute-old pane, I suggested the pane-age effect had
weakened. It had not: the SCENARIO VERDICTS survive because the settled retry
still repairs enough, and the damage shows in how hard it has to work. Verdicts
are a coarse instrument; the counters are the measurement.

### `a selected shape survives an insert` failed three times in thirteen

Rounds 130, 134 and 141 — at pane ages 93s, 64s and 4283s, so not pane age. It
is the scenario that drives `setSelectedShapes`, and this host's selection layer
has two open upstream issues against it (#3083, #3698). Three failures in
thirteen is a rate, not a flake, and it is the first scenario in weeks to fail
repeatedly without dissolving on its pair.

**Not chased here.** The gap and the pane are what these rounds were run to
settle, and starting a third investigation on the same data is how the profile
effect got published three times. Recorded as the next question.

## Chasing the selection scenario: it never failed, and I said it had

`a selected shape survives an insert` across the whole archive:

    rounds carrying it : 117
    pass 114   FAIL 0   skip 3

**Zero failures in 117 rounds.** I reported it yesterday as "the first scenario
in weeks to fail repeatedly", which was wrong: my verdict counter treated
`ok: false` as a failure, and a SKIPPED scenario is also `ok: false`. This repo
has an explicit rule that a skip is not a flip, and I broke it with a counter.

### What the skips actually say, and they are identical

    the host did not finish a draw made while a shape was SELECTED — the only
    draw in this battery made with a selection standing, and what
    dropShapeSelection exists to avoid: PowerPoint did not respond while
    drawing shapes 1-10 of 16 (45s)

All three word-for-word the same. This is office-js#3698 territory — the web
host cannot be relied on to insert while another shape is selected — and the
scenario exists precisely to probe it. **A skip here is the instrument working:
it declines to judge rather than calling a hung host a pass or a fail.**

### The budget is sound, so the hang is real

Pooled over 2485 timed draw batches:

    p50 5571ms   p90 15405ms   p99 21431ms   max 31086ms
    batches over 40s: 0 (0.00%)

The budget is 45000ms and the slowest draw ever recorded is 31s. A draw still
outstanding at 45s is hung, not slow.

**And I nearly reported the opposite.** My first query found "zero timed draws"
and I was about to record a missing instrument — the durations are keyed
`prevBatchMs`, attached to the FOLLOWING batch, and I had searched for `ms`. The
instrument was there; the query was wrong. Second time this week a conclusion of
mine rested on a field name.

### What is real and unexplained

Draw stalls across the archive: rounds 28, 88, 130, 134, 141.

    before round 129 :  2 of 104   (1.9%)
    round 129 onward :  3 of 13    (23%)

**`rasterGap` is NOT the cause**, and the ordering says so rather than an
argument: the draw stall lands at 761s, 799s and 5237s while the visibility
renders it enabled land at 815s, 871s and 5371s. The stall happens first, every
time.

No other candidate survives contact with the data, so **no mechanism is being
proposed**. Three occurrences, a twelvefold rate change, and nothing to attach
it to. Inventing an explanation here is exactly how the 4:3 profile effect got
published three times.

**One instrument gap, recorded not fixed:** `prevBatchMs` is carried by the NEXT
batch, so the LAST batch of every run has no duration. The pool above is missing
one measurement per draw sequence, which is why "max 31086ms" is a floor.

## What the archive records that nothing reads

Asked what the rounds could teach about observability, three questions the files
already answer:

**1. 71 of the 87 distinct trace messages are never named by triage, the gate or
the cycle.** Most are narrative and that is fine. Four are not — each records a
path the code took because its FIRST choice failed, thousands of times, with
nobody watching the rate. Banded by round, mean per round:

    rounds      tagging-failed  redraw-instead  scratch-retry  scratch-wrecked
      1- 40          5.2             9.3            9.3            10.7
     41- 80          6.6             9.1           14.1            11.1
     81-110          0.4            12.9           14.8            11.2
    111-141          0.2            13.0           14.6            10.7

**Two things were happening and nobody could see either.** Tagging failures
collapsed THIRTYFOLD around round 81 — a real win, never verified, and if it
regresses nothing would say so. Meanwhile in-place updates began falling back to
a full redraw 40% more often, which is slower and touches more of the deck, and
that drift went unremarked across sixty rounds.

`poolFallbackRates` pools all four now and the gate prints them.

### The instrument I nearly shipped blind

The first version reported each fallback against the MEDIAN OF ALL PRIORS. Run
against the archive it said:

    in-place update fell back to a redraw    13  (usually 13)

**A median absorbs a slow climb.** The signal that motivated the whole instrument
had risen 9 to 13 across sixty rounds, and by the time anyone looked, "now" and
"usually" were both 13 — so the check would have called it normal for exactly as
long as it kept getting worse. A detector blind to its own motivating case.

It reports the oldest third against the newest third as well now:

    charts left un-tagged                     0  (usually 1)   <- falling, 8 to 0
    in-place update fell back to a redraw    13  (usually 13)  <- RISING, 8 to 13

Guarded with a fixture whose median deliberately sits between the two thirds, so
a regression to median-only reading turns it red — and it does.

**2. Five scenarios have never discriminated in 117 rounds** — always pass, never
fail, never skip:

    insert on top of an earlier run
    which selection call wedges the host
    edit the chart the user selected
    stop a run part-way
    does a rasterise poison the next draw

Not necessarily vacuous: some may guard something that genuinely never breaks on
this host. But a gate that has never gone red is a gate nobody has seen work, and
this repo's own rule is that a test you have not watched fail is not yet
evidence. **Recorded, not acted on** — proving which is which means mutating the
production path each one guards, and that is its own piece of work.

**3. No starved probe questions.** Every host-probe question has answered at
least once, so nothing there is asking into the void.

## Chasing the redraw drift: it was never a drift, and the feature has never run

**FIRST, THE CORRECTION.** I reported `not updating in place — redrawing instead`
as a 40% regression, 9.3 to 13.0 per round. The per-round series says otherwise:

    rounds 23-44   ~8
    rounds 45-69    7, flat
    round  70      jumps to 11, and stays 11-12 for every round since

A step, not a drift — and my four-band averaging smeared it into a slope. The
step lands exactly on build `01f3607`, PR #547 "Bind the tag target, so the
settle has a handle that is not an id", which touched the tagging path.

**AND THE FALLBACK RATE NEVER CHANGED AT ALL, because it has always been 100%.**

    updated only the shapes that changed   : 0 times in 117 rounds
    not updating in place — redrawing      : 1301 times in 117 rounds

**The in-place chart update has never once succeeded.** What rose at round 70 is
the number of updates ATTEMPTED per round, not the share that fail. Reporting a
count without its denominator is the same mistake I made with the "81%
under-count" two days ago, and I made it again here.

### The project already knew, and stopped looking

    #405  Change one thing, write one shape — the in-place chart update
    #406  The in-place update fired zero times and would not say why

#406 added the fallback trace precisely to answer this. **The answer has been
sitting in every round file since, read by nothing** — and the docs record none
of it. Nothing in BACKLOG, RESEARCH or the brief says the feature has never run.

### The answer the trace has been giving

    the chart has no parts list, so its nodes cannot be mapped   12 of 13
    this update draws a picture, which is not in the scene        1 of 13

The picture case is legitimate — a picture is not in the scene, so the scene
cannot decide the update. The other twelve are the whole story: charts reach the
update path without a `CHART_PARTS_TAG`, so their nodes cannot be mapped to
shapes and the only option left is a full redraw.

**Root cause not established here.** Whether the tag is never written, written
and lost, or written and not found is a separate investigation, and the one thing
this day has taught repeatedly is that a mechanism invented to fit a number is
worse than an unexplained number.

### What shipped

`poolInPlaceUpdates`, and a gate line that cannot be missed:

    IN-PLACE UPDATE HAS NEVER SUCCEEDED — 0 successes against 1301 fallbacks
    over 117 round(s).

A count of zero is the hardest thing for a report to say, because nothing draws
attention to a line that is never printed. Guarded both ways: the test pins the
zero AND pins that a success would be noticed, so the day the feature starts
working the gate stops calling it dead.

## Chasing the parts list: the archive cannot answer, and that IS the finding

The in-place update falls back with "the chart has no parts list" 12 times out of
13, every round. Why that list is missing has three possible answers, and the
code makes all three reachable:

    never built     the chart was GROUPED, and `loose` is
                    `!grouped.has(i) && tagTargets[i]` — a grouped chart
                    collects no siblings at all
    built and lost  the id read-back sync threw, and the catch returns an
                    all-undefined list
    built, written, and not found again by the update path

**Three faults, three different fixes — and the archive cannot separate them.**

    trace entries carrying a `partIds` field : 0

The field exists in the code and is written to a structure nothing traces. A full
day of archive mining could not choose between the three, because the evidence
needed to choose was never recorded.

### What the archive DOES say, and where it stops

The read-back failure is real but partial: `reading back an ungrouped chart's
shape ids` fails 3.8 times per round on average, while `no parts list` is a flat
12 in every round from 070 onward. So the catch explains roughly a third of it at
most, and something structural accounts for the rest.

The obvious candidate is grouping — this host groups successfully (20 grouped, 0
refused, most rounds) and a grouped chart is not `loose`. **That is a hypothesis,
not a finding**, and it does not fit cleanly: round 141 grouped 10 and refused 9,
and its `no parts list` count was still exactly 12. A cause that varies while the
effect stays constant is not the cause.

So it is left as a hypothesis, and the instrument is shipped instead.

### The instrument

`tracePartsOutcome` reports, from every one of the three exits:

    charts, groupedSoNotLoose, noTagTarget, looseWithNoSiblings, gotPartsList

and `where` — which exit was taken — because two of the three are early returns
and a count without its exit cannot tell them apart either.

**Guarded on the count of exits, not on the presence of one call.** The specific
way this gap reopens is someone adding a fourth early return and not adding a
report to it, so the test asserts that the number of `tracePartsOutcome` calls
EQUALS the number of `return partsJson` statements. Mutation-checked: deleting
one report turns it red.

**The next round answers this.** Not a guess — the numbers will say which of the
three it is, on the first round after this ships.

## Round 142 — the instrument answered on its first run: it is grouping

    WHICH EXIT
      18x  no loose chart had siblings
       3x  the id read-back threw

    TOTALS
      charts                21
      groupedSoNotLoose     18
      looseWithNoSiblings    1
      gotPartsList           0

**Not one chart in the round got a parts list, and 18 of 21 because they were
GROUPED.** `loose` is `!grouped.has(i) && tagTargets[i]`, so a grouped chart
collects no siblings and no parts list is written — by design, because a grouped
chart is one shape and does not need one to be deleted as a unit.

But the in-place update path REQUIRES that list to map scene nodes to shapes. So:

**the in-place chart update is structurally incompatible with successful
grouping**, and on this host grouping nearly always succeeds. That is why the
feature has never run in 117 rounds. Not a bug in either path — a conflict
between two designs that were never reconciled.

The read-back failure is real but minor: 3 of 21, consistent with the 3.8 per
round the archive already showed.

### And the reason I rejected this yesterday was a denominator error

I argued grouping could not be the cause because grouping varies while `no parts
list` stays at exactly 12. The two counts do not share a denominator:

    round  grouped  refused  update attempts  no-parts-list
     139     15        4           13              12
     140     12        5           13              12
     141     10        9           13              12
     142     18        2           13              12

**Update attempts are a constant 13 whatever grouping does** — the battery makes
a fixed number of updates. "A cause that varies while the effect stays constant
is not the cause" was the right rule applied to the wrong pair of numbers.

Third denominator error in three days: the "81% under-count", the "40% redraw
regression", and this. The pattern is always the same — two counts compared
without asking whether they are counts of the same thing.

### The decision this leaves, which is not mine to make

Three ways out, and they are genuinely different products:

1. **Teach the update path to work on a grouped chart** — read the group's
   members instead of a stored parts list. Most work, keeps both features.
2. **Do not group charts that are meant to stay editable** — cheapest, but
   grouping is what makes a chart behave as one object for a user.
3. **Remove the in-place update** — it has never run; the fallback redraw is what
   has always happened, and deleting it would cost nothing that is working today.

Recorded rather than chosen. The gate now says the feature has never succeeded,
so whichever is picked, the evidence for picking is on screen every round.

## Rounds 143 + 144 — the mapping is solved; the wall moved

    round 143   no parts list 12 -> 1,  one-for-one mismatches 5
    round 144   no parts list      1,   one-for-one mismatches 0

**The group read works and the anchor fix landed.** 143 proved the members are
reachable — `ShapeGroup.shapes` at PowerPointApi 1.8, against office-js#3014's
2022 note that they "cannot be reached" — and it was off by exactly one, five
times out of five, because the anchor was being sought by the TAGGED SHAPE'S id
and a group is never among its own members. 144, with `items.slice(1)`, has zero
mismatches.

**13 of 13 both rounds, and that matters more than the counters here.** This is
the first change in this sequence that WRITES to a chart rather than measuring
one. The drawing-order assumption — that the group answers in the order the
parts tag would have recorded — is not provable from the type definitions, and a
wrong order would scramble a chart rather than refuse. Two rounds of `edit a
chart on the visible slide` and `an update follows a moved chart` passing is the
evidence that it holds on this host.

### Four of eleven declines are now the differ working

    6x  the chart carries no scene fingerprint
    3x  too much of the chart changed to be worth writing shape by shape
    1x  this update draws a picture, which is not in the scene
    1x  no parts list and no readable group members

"Too much changed" and the picture case are correct refusals — the fast path
declining work a redraw does better. Before this week none of these could even
be reached.

### The fingerprint blocker is not new, it was masked

    rounds 139-142   0 fingerprint declines
    rounds 143-144   6 each

Zero before, because the parts check fired first and short-circuited. Charts now
get past it and meet the next condition, which was always there.

**And it has a shape:** charts 1/8, 2/8 and 3/8 carry a fingerprint; 4/8 through
8/8 do not. Per-item, not per-batch — which points at the options rather than at
the tag write, since `sceneTag` is defaulted as `{ sceneTag: sceneFingerprint(scene), ...opts }`
and a caller-supplied `opts` overrides it.

**Not chased here.** Recorded with the numbers that will start it.

## The fingerprint blocker, chased — and the guess in the section above was wrong

Round 144's entry closed by saying the per-item split "points at the options
rather than at the tag write, since `sceneTag` is defaulted as
`{ sceneTag: sceneFingerprint(scene), ...opts }` and a caller-supplied `opts`
overrides it."

**That was the wrong mechanism.** No caller was overriding anything. The split
was per-item because the eight charts come from TWO DIFFERENT WRITERS, and only
one of them had ever written the tag:

    charts 1-3   insertSceneIntoSlide  ->  powerpoint.ts, stamps sceneTag
    charts 4-8   the generated deck    ->  pptx-deck.ts -> ooxml.ts, never did

`sceneFingerprint` was not imported in `pptx-deck.ts` at all. The default that
looked like the suspect is on the path that already worked.

**The lesson is about the shape of the evidence.** "Per-item, not per-batch"
was read as "the items differ in their options" when it equally meant "the items
took different code paths" — and the second reading was cheap to check and never
made. One grep for `sceneFingerprint` in `src/render/` would have settled it
before the theory was written down.

### The decline message was lying about itself

    "the chart carries no scene fingerprint - it was drawn by an older build"

These charts were drawn by the CURRENT build, which had never written one. The
message named a cause it could not know, and named it confidently enough that
the journal repeated the framing. A refusal that reports a condition should
report the condition, not a story about how the condition arose.

### Present is not fixed

`tryInPlaceUpdate` does `sceneFingerprint(buildChart(JSON.parse(tags.config)))`
and refuses on a mismatch, so writing SOME value would have moved the refusal
one reason along and looked like progress — the same way the parts check hid
this one for 117 rounds. The guard therefore asserts the value equals what the
update path recomputes, and was mutation-checked with a plausible-looking wrong
fingerprint, which a presence check would have passed.

Landed as #645. What round 145 should show: the six `no scene fingerprint`
declines gone, and — for the first time — `updated only the shapes that changed`
on a chart from the deck.

## Round 145 — the fingerprint fix held; the positive claim did not

The prediction was staked before the round landed, in two halves, because this
fix had a failure mode that looks like success.

    rung  4  the chart carries no scene fingerprint      6 -> 0   as predicted
    rung  7  no longer renders to the scene that was drawn  0 -> 0   as predicted
    rung  9  too much of the chart changed to be worth it  3 -> 8
    in-place SUCCESSES                                     0 -> 0   PREDICTION FAILED

**The half that mattered held.** Rung 7 is the mismatch rung: if the value
written into the deck were not the value `sceneFingerprint(buildChart(config))`
produces at update time, all six declines would have migrated there and the
total would not have moved. They did not migrate. The stored fingerprint is the
one the update path recomputes, and #645 is a real fix rather than a cosmetic
one.

**The positive claim failed outright.** "At least one `updated only the shapes
that changed`" was staked precisely so a smaller round could not satisfy it, and
it came back zero. The denominator is sound — 15 declines in each of 143, 144
and 145, and 578 / 571 / 580 trace entries — so this is a real zero, not a
different population.

### Where the six went, and what was waiting there

Five of the six landed on rung 9, which is the differ correctly declining work a
redraw does better: `changed 18 of 24` and `changed 9 of 16`, three quarters and
better than half of the chart. Those are right.

**The sixth kind reached the host and the host refused it** — three charts, all
`changed 1 of 24`, all with the same error:

    InvalidArgument | errorLocation=Shape.textFrame
    statement: var textFrame = shape.textFrame;

`changed 1 of 24` is the differ working perfectly: it isolated a single changed
node and went to write only that. The write went to the wrong object. For a
grouped chart `shapes = [old, ...parts]` uses the GROUP as node 0, and node 0 is
the group's first member. A `ShapeGroup` has `fill` and `lineFormat` — both
navigations succeeded in the host's own statement list — and no `textFrame`.

### The safety argument in #643 was wrong, and precisely wrong

That PR justified the positional mapping like this: *"the one-for-one guard still
refuses anything that doesn't line up"*. It does not. **A count guard checks how
many, and this was an error in which element means what.** `parts.length + 1`
came to 24 against 24 nodes both before and after the defect — the guard was
satisfied by the broken mapping and would have been satisfied by any permutation
of it.

The lesson generalises past this bug: a cardinality check is not an alignment
check, and quoting one as protection for the other is how a wrong mapping gets
called safe in writing.

### The suite had a test for this and it was green

`updates a GROUPED chart through its group members` edits the title — node 0,
the only node that was mis-targeted — and passed the whole time. The fake host
lets a group accept a name and a text frame, so the wrong target was invisible
in jsdom while the real host refused it three times out of three.

**The double was more permissive than the thing it doubles**, which is the same
shape as the fake that populated values on request and hid a 56-round defect. A
test that exercises the exact defect and passes is worse than no test: it was
counted as coverage.

The new guard asserts the TARGET rather than the outcome — after a title-only
update the group must still be named `PowerChart` — because the outcome was
indistinguishable in a host that accepts everything.

### And one underneath it

Fixing the mapping made every grouped write refuse with *"the host would not
confirm every shape that had to change"*. `isLive` asks `isNullObject === false`,
which is the protocol of an `…OrNullObject` lookup; a shape handed back as an
item of a loaded collection has no such flag and reads `undefined`. The guard
called a shape dead because it had arrived alive by a different route.

Landed as #646. What round 146 should show: the three `Shape.textFrame`
refusals gone, and the first `updated only the shapes that changed` in 119
archived rounds.

## Round 146 — the in-place update ran, for the first time in 119 rounds

    round 145   13 update attempts   0 in place   3 refused at Shape.textFrame
    round 146   13 update attempts   3 in place   0 refused at Shape.textFrame

    WINS: changed 1 of 24, saved 47   (x3)

Both halves of the staked prediction held. The three charts that the host had
refused are the same three that now succeed, and each one wrote a single shape
where a redraw would have written 48.

**The population is provably identical.** Thirteen update attempts in each
round — the constant established weeks ago — and the three did not appear from
anywhere, they moved out of the decline column. 145 was 13 declines and 0
successes; 146 is 10 and 3.

### The unclassified bucket was the interesting one, for two rounds

My decline breakdown bucketed by `data.why` and let everything without one fall
into a `?` bucket that I read as noise:

    round 143   2 unclassified
    round 144   4 unclassified   <- THREE of these were Shape.textFrame refusals
    round 145   5 unclassified   <- and these

I opened the bucket at round 145 and found the host refusals in it. **They were
in round 144's trace too, and I had reported that round without looking.** The
entries had no `why` for the ordinary reason that they had not declined by a
rule — they had THROWN, and a throw carries an `error`, not a `why`.

So the mining lesson is specific: **a residual bucket in a breakdown is not
noise until it has been opened once.** The named categories are the ones already
understood; the interesting thing is, almost by definition, the one that does
not fit them. Two rounds of a host-side defect sat in plain sight in a bucket I
had labelled and skipped.

**And the count in that bucket is not the count of anything.** Two of round
146's "unclassified declines" are a by-id `GeneralException` that RECOVERED
(`asked 1, recovered 1`) — matched only because the message contains the word
"refused". Excluding them is what makes the arithmetic exact at 13 and 13. A
filter that greps prose will pick up prose.

### What is left is the differ declining correctly

    8x  too much of the chart changed     18 of 24, 9 of 16
    1x  this update draws a picture
    1x  no parts list and no readable group members

Three quarters and better than half of a chart changed: a redraw genuinely does
those better. The remaining two are correct by construction.

13 of 13 scenarios passed. Friction — 3 errors, 2 id refusals, 1 general
exception — all recovered.

## Round 147 — the pair confirms it

    145   ok 0 | declined 10 | host-refused 3 | attempts 13 | scenarios 13/13
    146   ok 3 | declined 10 | host-refused 0 | attempts 13 | scenarios 13/13
    147   ok 3 | declined 10 | host-refused 0 | attempts 13 | scenarios 13/13

An exact repeat, reason breakdown included. `9ef98af` differs from `2521d23`
only by a comment and an archive, so this is the second run of a pair rather
than a new build — and the second run is usually the WORSE one. An identical
result is therefore the strongest form this evidence takes: three successes is
the rate, not a lucky round.

Thirteen attempts in all three rounds, which is the constant. Nothing here
depends on a denominator that moved.

### Both of these rounds were read through an instrument that could not see the failure

`poolInPlaceUpdates` — the pool behind the gate's "A FEATURE THAT HAS NEVER RUN"
banner — counted two of three outcomes:

    updated only the shapes that changed        -> ok
    not updating in place — redrawing instead   -> fell
    in-place update refused — redrawing instead -> COUNTED AS NEITHER

The third is the THROW path, which carries an `error` and no `why`. So every
host-side refusal registered as neither a success nor a fallback and left no
trace in the gate at all: the three `Shape.textFrame` refusals in round 145 and
two more in 144. The tool was structurally incapable of showing the thing that
was missed by hand, which is why looking harder would not have helped.

Pooled over the archive it now reads:

    in-place update: 3 succeeded, 1358 declined, 5 refused by the host over 122 round(s)
      1197x  the chart has no parts list, so its nodes cannot be mapped to shapes
       121x  this update draws a picture, which is not in the scene
        19x  too much of the chart changed to be worth writing shape by shape
         5x  InvalidArgument at Shape.textFrame

And the residual is NAMED rather than tallied. A bucket reported as a number
reads as noise and gets skipped; reported as a line it gets opened.

### A measurement fault found by making it twice

`attempt` re-read `git HEAD` on every retry, so committing in this clone while a
round retried moved the build under test mid-round — the driver would then
refuse its own round as `site-behind` against a commit that had never been
deployed and was not what it was measuring. Both times it was harmless only by
luck.

The habit is the operator's; the fault is the driver's. **A round tests one
build**, so `main` now reads HEAD once and hands it to every attempt.

## Round 149 — the confirmation: 148 was the host, not the build

    146   ok 3 | declined 10 | host-refused 0 | attempts 13 | scenarios 13/13
    147   ok 3 | declined 10 | host-refused 0 | attempts 13 | scenarios 13/13
    148   ok 3 | declined  8 | host-refused 0 | attempts 11 | scenarios 11/13
    149   ok 3 | declined 10 | host-refused 0 | attempts 13 | scenarios 13/13

Round 148's two scenario failures did not reproduce. Neither could have been
ours: `git diff --name-only 148..147 -- src/` is empty, so the app code was
byte-identical to the round before that scored 13/13.

**The host was having a bad spell across both rounds, and only one of them
said so.** 148 needed three attempts to start — a silent host, then a closed
pane. 149's BROWSER DIED 245 seconds in and the driver rebuilt it. The
difference is that 148's trouble happened before the round and left no mark on
the round file at all.

    in-place successes, three clean rounds running: 3 of 13

### What 148 could not tell me, and now can

A successful recovery erases its own evidence. Once `recover` works the round
looks exactly like one that never needed it, so for all 149 archived rounds a
run against a sick host was indistinguishable from a clean one — and "was the
host unwell?" is the FIRST question to ask of a first-ever scenario failure.
Rounds now carry `driverRun { attempts, recovered }`.

### Four faults of one family, all found by asking that question

1. **The archive threw the driver's own history away** — see above.
2. **`FALLBACK_SIGNALS` had the same two-of-three gap `poolInPlaceUpdates`
   had**, and the sibling sweep for that fix missed it. It tracked the
   rule-based decline and not the throw, so the drift detector understated
   every round in which the host refused a write. Given its own label rather
   than merged: a write the host threw out is a different event from the differ
   declining work a redraw does better, and drift in one must not absorb the
   other.
3. **Nothing checked that a detector still matches anything.** Every tool here
   finds evidence by comparing a trace message to a string literal. Rename the
   message in `src/` and the detector does not break — it reports zero, every
   round, forever, and zero is what a healthy round looks like. All the current
   literals are live; the guard is for the rename that has not happened yet.
4. **`saved` was a fabricated number.** `nodes * 2 - changed` asserted a cost
   model nobody measured — and the host's own statement list refutes it, since
   writing ONE node emits about twenty statements. Taken literally it says the
   in-place path always wins, which contradicts `UPDATE_SHARE_LIMIT` sitting
   beside it, itself documented as untuned. Dropped. `changed` of `of` is the
   measurement.

### The archive was carrying the failure twice

Round 148's file is 408 KB against ~197 KB for every clean round, and 111 KB of
that is one `scenario FAILED` entry — the same base64 payload also stored in
`selftest`. `trimDebugInfo` capped the COUNT of statements and not their
LENGTH, and `insertSlidesFromBase64` puts the whole deck in one of them.

## Round 151 — 0.6 holds, and the ladder has no more rungs

    149  0.5   ok 3 | declined 10 | 13/13 | same scale 237048 ms
    150  0.5   ok 3 | declined 10 | 13/13 | same scale 195928 ms   (control, host healthy)
    151  0.6   ok 5 | declined  8 | 13/13 | same scale 164518 ms

Every point of the staked prediction held: declines 8 to 6, successes 3 to 5,
13 of 13, and the round COMPLETED — which was the part that mattered, because
0.8 crashed PowerPoint on all seven attempts of the round before.

**164518 ms is 31 seconds below the MINIMUM of the five prior observations** of
that scenario (197426, 201182, 204085, 237048, 195928), and 31.4s faster than
the control run immediately before it. Same eight charts, same thirteen update
attempts, so nothing about the denominator moved.

### The ladder is finished, and not because it succeeded

The obvious next step would be 0.7. It would do NOTHING: the only two shares
this battery produces are 9-of-16 (0.5625) and 18-of-24 (0.75), and there is
nothing between 0.6 and 0.75 to admit.

And any limit that admits 18-of-24 also admits 9-of-16, because 0.5625 is below
it. So a limit of 0.76 hands the host exactly what 0.8 did — eight charts
writing eighteen shapes plus two writing nine — and 0.8's six crashes ARE the
answer for 0.75. There is no experiment left to run at this granularity.

    0.5   ->  0 extra shape-writes    safe, measured
    0.6   ->  18 extra                safe, measured, 31s faster
    0.75+ ->  162 extra               fatal, measured six times

**So 0.6 is the ceiling for a limit expressed as a share.** Not because 0.75 is
untested — because it is tested, and it kills the host.

### What would actually unlock the rest

The crash is a batch-size problem, and this project already solved a batch-size
problem once: the demo deck renders ONE `PowerPoint.run` PER SLIDE because a
chunk of four dense charts piled 400-500 shapes into a single context and the
host's per-slide budget refused it (#112).

The in-place update has no such chunking. It writes every changed shape of
every chart into one batch and syncs once. At 0.6 that is 18 extra writes and
fine; at 0.75 it is 162 and fatal.

**Chunking the in-place writes is the change that would make the share limit
irrelevant**, and it is a better fix than any threshold: a limit expressed as a
share is a proxy for a cost the host measures in shapes per batch. That is the
next product step, and it is worth doing properly rather than by raising a
number until something dies.

## Round 152 — chunking costs nothing, which is what makes the next round worth running

    151  0.6, unchunked   ok 5 | declined 8 | 13/13 | same scale 164518 ms
    152  0.6, CHUNKED     ok 5 | declined 8 | 13/13 | same scale 164626 ms

Every point of the staked prediction held. It completed, 13/13, and the success
and decline counts did not move — the point worth watching, because chunking
changes HOW a write is sent and not WHETHER it is attempted, so any movement
there would have meant an unintended change.

**108 milliseconds on a 164-second scenario: 0.07%.** Two 9-shape writes now go
out as 6+3 with an extra sync each, and the cost of those extra round trips is
lost in the noise.

### Why that matters more than it looks

It removes the obvious objection to chunking — that more syncs must cost more
time — and it makes the real experiment cheap. If bounding a batch is free at
this size, there is no reason not to bound it at every size, and the share
limit can be tested against the thing it was always a proxy for.

### The next round discriminates between two hypotheses, and both answers are useful

0.8 crashed PowerPoint six times at 255-284s. Two explanations survive that
round, and the existing data cannot separate them because both grew together:

    per-SYNC statements     8 charts x 18 shapes = ~360 statements per sync
    per-CONTEXT accumulation  every chart's proxies and queued work in one PowerPoint.run

Chunking fixes the first and does nothing about the second. So raising the
limit again, WITH the batch bounded at six, asks the question directly:

    it completes  ->  the crash was per-sync size, and chunking is the fix
    it crashes    ->  the crash is context accumulation, and the fix is a fresh
                      `PowerPoint.run` per chart — which is exactly what #112
                      did for the demo deck, one run PER SLIDE

A crash would cost a round and a recovery. It would also end the guessing, and
it is the only remaining way to find out.

## Round 153 — the crash was per-SYNC size, and the feature is now at its ceiling

    150  0.5  unchunked   ok  3 | declined 10 | 13/13 | same scale 195928 ms
    151  0.6  unchunked   ok  5 | declined  8 | 13/13 | same scale 164518 ms
    152  0.6  CHUNKED     ok  5 | declined  8 | 13/13 | same scale 164626 ms
    153  0.8  CHUNKED     ok 11 | declined  2 | 13/13 | same scale 157403 ms

**0.8 is the value that crashed PowerPoint six times at 255-284s, and it
completed.** The only thing that changed between those two rounds is that a
write now goes out six shapes at a time. So of the two explanations round 150
left standing, it was the first: per-SYNC statement volume, not per-context
accumulation. A fresh `PowerPoint.run` per chart is not needed.

### The decline column is finished

`too much of the chart changed` went from six a round to none. What is left is
two refusals that are correct by construction:

    1x  this update draws a picture, which is not in the scene the differ compares
    1x  the chart has no parts list and no readable group members

**Eleven of thirteen is every attempt this battery can produce.** The feature
that had never run once in 119 archived rounds now runs on everything it is
allowed to.

### What the numbers do and do not support

The counts are deterministic and not in doubt: 6 declines to 0, 5 successes to
11. The timing is weaker evidence and should be read as such — 157403 ms
against 164626, one observation each. The 0.6 pair is unusually tight (164518
and 164626, 108ms apart), which makes a 7.2-second drop look real, but n=1 is
n=1. The honest claim is the 195928 to 157403 span across the whole sequence:
**20% off the longest scenario in the battery**, with three intermediate points
that move in the right direction each time.

### The constant that mattered was never the one being tuned

`UPDATE_SHARE_LIMIT` spent this entire investigation being raised, crashed,
reverted, raised again — and it was a proxy the whole time. The thing holding
the host up was `IN_PLACE_WRITES_PER_SYNC`, which did not exist until round 152.
What survives in the share limit is a risk cap, not a cost model.

### Every round of this sequence needed two attempts

    150  {"attempts":2,"recovered":["not-ready"]}
    151  {"attempts":2,"recovered":["not-ready"]}
    152  {"attempts":2,"recovered":["not-ready"]}
    153  {"attempts":2,"recovered":["not-ready"]}

Four for four, always `not-ready`, always recovered. That is a minute or two
per round spent on a first attempt that never had a chance, and it has been
invisible for the whole archive — `driverRun` is three rounds old. Chased next.

## Round 155 — the first round in the archive to start first time

    153            ok 11 | declined 2 | 13/13 | driverRun {"attempts":2,"recovered":["not-ready"]}
    154            ok 11 | declined 2 | 13/13 | driverRun {"attempts":2,"recovered":["not-ready:host-silent+pane-stale"]}
    155  --fresh   ok 11 | declined 2 | 13/13 | driverRun {"attempts":1,"recovered":[]}

**Attempts 2 to 1, and `recovered` empty.** Every round before this one started
from whatever the last had left behind, gone stale, and paid a failed attempt
plus a recovery for it. The label added in round 154 is what made the fix
obvious: `host-silent` is the editing session dropping while the browser sits
idle, `pane-stale` is a build deployed since the pane loaded, and a session
opened a moment ago has neither problem.

The counts did not move — 11, 2, 13/13 — which is the requirement. A fresh
browser changes how a round STARTS, not what it measures, and any movement
there would have meant an unintended change.

### The timing is suggestive and is not being claimed

`same scale across the deck` came in at 150740 ms against 159662 and 157403.
That is 9 seconds under the round before it, and there is a plausible mechanism
— a fresh session carries no accumulated state, and pane age is the best
explanation this archive has for "the second round of a pair is worse". But it
is ONE observation, and the honest position is that the attempts claim is
established and the timing claim is not.

### What this fixes that is not speed

Every cross-round comparison in this archive rests on rounds beginning the same
way. They did not: each one began from a different amount of leftover state,
depending on how long the gap had been and what had been deployed. The
195928-to-157403 sequence, the 108ms chunking measurement and the 0.6 pair were
all read against that. **Making the start deterministic is worth more than the
nine seconds it appears to save.**

Not the default yet. One round is enough to show the mechanism works and not
enough to adopt on.

## Rounds 156-157 — a prediction that held twenty times and read as FAILED

    154            ok 11 | declined 2 | 13/13 | same scale 159662 ms | attempts 2
    155  --fresh   ok 11 | declined 2 | 13/13 | same scale 150740 ms | attempts 1
    156            ok 11 | declined 2 | 13/13 | same scale 153799 ms | attempts 3
    157            ok 11 | declined 2 | 13/13 | same scale 185389 ms | attempts 2

The verdicts are a flat line across four rounds and seven driver fixes, which is
what those fixes were for: they change how a round STARTS, not what it measures.

### The timing claim from round 155 is now refuted, by its own successor

Round 155 recorded 150740 ms and this journal wrote that the mechanism was
plausible and the claim was not being made. Round 157 came in at **185389 ms —
23% slower — on a build whose only diff from 156 is `scripts/round.mjs`, a test
file, and an archived round.** No src/ change whatever. So the same application
code, twice, produced 153799 and 185389.

**`same scale across the deck` has a noise band of at least ±20%.** Every timing
comparison in this archive that quoted a difference smaller than that quoted
noise, this journal's own nine-second observation included. Good that it was
staked as unclaimed; the discipline earned its keep four rounds later.

### The find: #520 held in 20 rounds and the ledger printed FAILED

`the-refused-group-is-what-kills-the-tag` was staked on 2026-08-18 and has sat
OPEN ever since, reported as FAILED whenever triage looked at it. Pooling the
probe across the whole archive rather than reading the newest round:

    does-a-failed-group-poison-the-tag   present in 114 of 133 rounds
      no-refusal   94   "the refusal was never provoked"    <- the question was never put
      tags-gone    20   ".tags was undefined after the refused group"  <- the claim, exactly

**Twenty confirmations, zero refutations, and a verdict of FAILED.** Three
separate defects stacked, and all three are this project's house defect wearing
different clothes:

1. **`no-refusal` is a never-put question scored as a refutation.** The probe
   says so in its own words. It is a THIRD category and is deliberately NOT
   folded into `NOT_ASKED`: not "the harness could not set it up" (ours to fix),
   not "the host would not answer" (a fact about the host), but "the precondition
   did not occur". A **conditional probe**, whose question exists only when the
   host misbehaves in a particular way. `UNPROVOKED` now names it.

2. **The claim matched on strings the probe cannot emit.**
   `InvalidParam|5010|GeneralException` against a probe whose vocabulary is
   `no-refusal`/`tags-gone` plus prose. Even the 20 firings would have read
   FAILED. Round 102 taught precisely this lesson from the trace side — *a claim
   is only as good as the set of strings the thing it watches can produce* — and
   it did not carry across to the probe side. Claims now match the answer KEY, an
   enum the probe picks, as well as the detail a later commit can reword.

3. **The reading was n=1.** `roundToJudgeOn` handed the judge the newest round.
   For a probe that fires 18% of the time, that is an 82% chance of being judged
   on a round that measured nothing. triage.mjs spends a page arguing that a
   single round's verdict is a fact about the code AND the host's mood that
   afternoon — and the prediction ledger was the one reader ignoring it.
   Predictions are judged across every eligible round now and print the split.

`BOTH` is a verdict in its own right. When the decided rounds disagree about
code that did not change, that is a finding about the host, not a tie to break
quietly.

### Sibling sweep

All 38 probe ids, every distinct `answer` value across 133 rounds. Only
`no-refusal` qualifies. `no-group-id` ("the host would not name the group it
just made"), `no-id` ("the fresh shape would not report an id"), `none` and
`unreadable` are the host DECLINING a question that WAS put — findings, not
starved probes — and they stay where they are.

### What it means for the product

The diagnosis behind #520 is **confirmed**: the grouping attempt this host
refuses 5010 is what takes the tag with it, and the tag-anchor change merged
beside it in August was aimed one level too low. The fix belongs at the context,
not at the handle. That has been true and recorded for 130 rounds.

### Doctrine

An instrument that reports a verdict for a question it never asked is the same
defect as one that hardcodes its conclusion. Both print a sentence with no
measurement under it, and both keep printing after they stop being true. The new
rule: **before believing a probe's verdict, pool its answers over the archive and
check the denominator.** A claim judged on one round is a claim about one
afternoon.

## Between 157 and 158 — two reports that kept demanding work already done

Neither of these is a round finding. Both are instruments that went on printing
after they stopped being true, which is the same defect #677 fixed in the
prediction ledger, and both were found by asking the same question of a
different report: **what is the denominator, and is the thing still there?**

### The starved-question report named two probes retired eight rounds earlier

`grouped-child-by-id-from-slide` and `tag-on-group-survives` were retired on
2026-08-21 and last appear in round 149. Since then the report has listed them
at "125 round(s)" under **OURS to fix** — telling the reader to retire two
probes that were already gone, in a bucket whose entire purpose is to name
actionable waste.

A count pooled over the whole archive cannot tell a live starving probe from a
dead one. `poolStarvedQuestions` now marks a question `retired` when the build
has stopped asking it, and the report gives them their own section rather than
dropping them — a row that vanishes reads as lost, a row that says "already
retired" reads as finished.

**A window of three non-empty sheets, not the newest round.** One short sheet
means the host died mid-probe; reading that as a retirement would quietly empty
the actionable bucket, which is the failure that matters here. Empty sheets do
not consume the window.

**What is left in that bucket is now exactly one live item:**
`shape-resolve-held-slide-proxy`, starved 133 of 133 rounds. That is the real
backlog entry, and it was previously buried among two closed ones.

### The Pages wait was hand-rolled every round, and tonight it lied twice at once

Every round is launched behind a "wait until the site serves HEAD" loop, retyped
at the shell each time. Tonight's copy carried two independent faults:

- it polled `dannbleeker.github.io`, which Pages has **301'd to the custom
  domain since July** — the driver's own `defaultFetchBuild` has used the
  canonical URL all along;
- it read `.commit` from a document whose only field is `build`.

Either fault alone prints `?`. So does a deploy that has not landed. **Ten
identical `?` lines are indistinguishable from patience**, and round 158 sat in
that loop for ten minutes after its deploy had already landed at 19:16Z.

`--wait-for-deploy` now lives in the driver, using the driver's own fetch and
its own `buildOf`. A caller that has to re-derive the URL or the field name can
get one of them wrong, and the failure looks like waiting. The poll prints the
build it actually read, so "could not read a stamp" and "read an old hash" are
different lines — the distinction the tick collapsed.

It does not throw on giving up. `readiness` is what judges a stale site: it
names the hash, names the fix, and `--retry` recovers from it. Dying here would
replace a diagnosis with a stack trace.

### Doctrine

Both of these were found the same way, and it is worth stating as a habit:
**when a report tells you to do something, check that the thing is still there.**
An instrument that names work is making a claim about the present, and nothing
about a stale claim looks stale.

## Rounds 158 and 159 — the first real pair in seven rounds

    155  --fresh   ok 11 | declined 2 | 13/13 | same scale 150740 ms | attempts 1
    156            ok 11 | declined 2 | 13/13 | same scale 153799 ms | attempts 3
    157            ok 11 | declined 2 | 13/13 | same scale 185389 ms | attempts 2
    158  2897681   ok 11 | declined 2 | 13/13 | same scale 152821 ms | attempts 2
    159  2897681   ok 11 | declined 2 | 13/13 | same scale 159320 ms | attempts 1

**Rounds 152-157 were each a single round on their own build.** The brief says
two per build minimum and warns that a merge between them makes the pair
incomparable, and that rule quietly lapsed for six rounds while driver fixes
were landing one per round. 158 and 159 are one build run twice with nothing
merged between them, and it paid for itself immediately.

### What the pair settled

**The timing noise floor.** 152821 against 159320 — a 4% spread on identical
code. Round 157's 185389 was an OUTLIER, not a trend: four builds with no
relevant src change read 150.7k, 153.8k, 185.4k, 152.8k. This retires round
155's suggestion that a fresh session saves nine seconds; nine seconds is
comfortably inside a spread that reaches twenty percent.

**A tag failure that is mood, not regression.** 159 carried
`tagging failed — charts are not re-editable until repaired`, 1 chart,
`ErrorPointer` at `BindingCollection.add` — the first non-zero
`charts left un-tagged` in eight rounds, and it would have read as a fresh
regression from the evening's work. **158 has none, on the same build.** The
signature is known: 19 of 134 rounds, always 2 or 4 entries, last seen in round
139. The pair is what makes that a shrug instead of an investigation.

**159 started FIRST TIME** — attempts 1, nothing recovered. Second such start in
the archive; 155 was the first, and it needed `--fresh` to get there.

### Two more instruments describing eras that had ended

**The grouping rate.** `WHICH SLIDE THE CHART LANDED ON` printed "66% of charts
on a freshly added slide group" and then told the reader not to believe it —
"pooled over 39 rounds that predate the retry ... should not be read as the
current rate". A figure a reader is told to mentally discount is a figure nobody
can use, and the answer to a stale window is a second window, not a caveat.

Over the last 20 rounds it is **81%**, and that number had been unavailable for
the seventy rounds since the settled retry landed. It also surfaces a live
defect the pooled figure was muddying: **13 of 68 charts on a freshly added
slide still fail to group**, in the current build, and the table directly below
says an ungrouped chart loses its config 44% of the time.

**The fallback drift arrow.** `in-place update fell back to a redraw` printed
`now 2` beside `usually 12` and `RISING, 8 to 13` — three summaries, all
describing an era that ended five rounds ago, two of them contradicting the
number sitting next to them. The signal actually read 13,13,11,10,10,8,8,2,2,2,2,2
across twelve rounds: the in-place update started working and the fallback
collapsed.

**A shorter window is not the fix.** The last-20 median is 10, which is just as
wrong — a median over ANY fixed window smears a step by construction. The gate
now prints the last eight values in sequence beside the summaries, and says so
outright when the recent run contradicts the arrow:

    in-place update fell back to a redraw  2  (usually 12)  <- RISING, 8 to 13 across
    45-round thirds — but the last 8 rounds median 2, so the thirds are describing an
    era that ended  [8,2,2,2,2,2,2,2]

### Doctrine

Three nights running, the same defect in three different reports: **a summary
that outlived the regime it summarised.** The prediction ledger judged on one
round, the starved-question report counted retired probes, the grouping rate
pooled a dead era, and the drift arrow read thirds of an archive whose newest
third was itself half history.

They share a tell worth naming: **the report knew.** Two of the four said so in
their own prose — "should not be read as the current rate", "Counts, not a rate"
— and printed the number anyway. A caveat next to a figure is an admission that
the figure is wrong, not a fix for it. Print the window that answers the
question, or print the sequence and let the reader see the shape.

## Round 160 — a correction to the entry above, and the population that halved

    160  e2b3f7d  ok 11 | declined 2 | 13/13 | 512s | attempts 2

### CORRECTION: "13 of 68 fresh-slide charts still fail to group" was wrong

That claim is in the entry above and in #680's commit message, and it named a
live product defect that is not live. Every one of those 13 failures comes from
rounds 137-142, an era that ended:

    137  1   138  4   139  4   140  5   141  9   142  2
    143-160: three failures in eighteen rounds

Over the last twelve rounds the fresh-slide rate is **20 of 20**, and over the
last eight there are no fresh-slide charts at all. The reason is the same one
running through everything else here: the in-place update took those charts, and
a chart that is not redrawn never lands on a fresh slide.

**The window I chose to fix a stale window was itself stale.** 20 rounds was
picked to exclude everything before the settled retry (rounds 064/065). It does
— and it sits straight across a second boundary I had not found. The same
population reads 80% at 30 rounds, 91% at 20, 100% at 12, and at 8 it is empty.

That is the lesson, sharper than the one in the entry above: **a window is a
guess about where the last regime boundary is, and you do not know where they
all are.** A sequence does not need the guess. So the fresh-slide block now
prints one, and `0/0` rounds are the signal rather than a gap:

    grouped/landed on a fresh slide, per round: [0/0 0/0 0/0 0/0 0/0 0/0 0/0 0/0]

A percentage over an emptying population is what to watch for, and only the
sequence shows it emptying.

### The grouping population halved at round 153 and nothing said so

Charts grouped per round ran 15-20 for the whole archive, then:

    round  151  152  153  154  155  156  157  158  159  160
    grouped 15   15    9    9    9    9    9    9    7    9
    in-place 5    5   11   11   11   11   11   11   11   11

**15 - 9 = 6 and 11 - 5 = 6, exactly, in the round it changed.** The cause is
complete and benign: an in-place update never redraws, and a chart that is not
redrawn is never regrouped.

Benign, and it silently rebased every grouping figure in triage. The gate line
read `9 chart(s) grouped, 0 refused (usually 2 over 135 prior round(s))` — and
"0 refused, usually 2" reads as an improvement when half the attempts stopped
happening. Fewer refusals out of fewer tries.

The rate is the only reading that survives a change of population: round 160
grouped **9 of 9 attempts**, round 141 grouped **10 of 19**. Those two are
comparable; 9 and 10 are not. This is the `same scale across the deck` trap
again — "6 of 6" and "8 of 8" are both a pass — argued at length one screen up
in this same file while a bare count was printed here.

### Doctrine, revised

The entry above said: print the window that answers the question, or print the
sequence. Round 160 says the first half of that is not safe advice. **Print the
sequence.** A window has to be chosen, and choosing it requires knowing where
the regime boundaries are — which is the thing you are trying to find out.

And when a count moves, ask what happened to its denominator before asking what
happened to the thing being counted. Twice tonight the answer was "a different
subsystem started working".

## Round 161 — the pair discipline and the fresh-pane discipline are in conflict

    160  e2b3f7d  ok 11 | declined 2 | 13/13 | 512s | grouped 9 of 9 | pane  70s
    161  e2b3f7d  ok 11 | declined 2 | 13/13 | 542s | grouped 7 of 8 | pane 666s

Second round of the pair, same build, nothing merged between them — and it
refused a group, left a slide holding **17 shapes** against 160's five, and
started with a pane **666 seconds old**.

### The measurement

    round  148 149 150 151 152 153 154 155 156 157 158 | 159 | 160 | 161
    pane    76  67  67  62  64  62  72  64  65  63  67 | 696 |  70 | 666

Rounds 148-158 were each taken after a merge. Rounds 159 and 161 are the only
true second-rounds in that stretch. **Both inherited a ~670-second pane, and
both were the worse half of their pair** — 159 lost a chart's tag
(`ErrorPointer` at `BindingCollection.add`), 161 refused a group. That is the
PAIR POSITION direction — second round worse 20x, better 2x — reproducing twice
in one evening with the cause visible in the same table.

### Why, and it is not what the gate said

The gate has been claiming *"The driver reloads it between rounds now, so this
count is mostly history."* **It does not reload it, and the reason is recorded
in `round.mjs` itself:** reloading the pane raises a beforeunload prompt over
unsaved work, and accepting that prompt appears to discard the per-document
SIDELOAD — the add-in was gone from the ribbon after round 124 and again after
round 132. The reload was deliberately removed. The note describing it was not.

What actually freshens the pane is **a merge**: it makes the pane stale, so
recovery reloads it as part of clearing `pane-stale`. Every fresh-pane round in
the archive got there that way.

Which puts the two disciplines in direct conflict. **The pair exists so that
nothing is merged between the two rounds — and the merge is the only thing
freshening the pane.** Run a proper pair and you guarantee the second round is
taken on an aged pane, i.e. you control for the build and lose control of the
confound that this archive says matters most.

Six rounds of merging-between-rounds hid it: 148-158 were all first rounds by
construction, so the pane was always fresh and PAIR POSITION looked like history.

### The fix

**Run the second round of a pair with `--fresh`.** It closes the BROWSER — which
the persistent profile survives, sign-in and sideload both — rather than
reloading the page that holds unsaved work. It is the only route to a fresh pane
that does not risk the add-in, and round 155 already demonstrated it: pane 64s,
attempts 1, nothing recovered.

Written into ROUND-LOOP-BRIEF as part of the pair rule rather than left as a
finding, because the next pair will otherwise repeat it.

### What this costs the two pairs taken tonight

They are still the best evidence in the archive for the timing question — 152821
vs 159320, and 512s vs 542s, are close enough that a 666-second pane did not move
them much. But **158/159 and 160/161 are not clean pairs**, and the two anomalies
in them (a lost tag, a refused group) both sit on the aged half. Neither is
evidence about the build. The next pair will be.

## Rounds 162 and 163 — the first clean pair in the archive

    round  158   159   160   161   162   163
    pane    67   696    70   666    70    65
    group  9/9   7/8   9/9   7/8   9/9   9/9
    ms     152k  159k  157k  158k  154k  166k

163 is the first true second-round in this archive to start on a fresh pane. It
ran with `--fresh`, per the rule written down after 161, and it grouped **9 of
9** where 159 and 161 — its counterparts in the two preceding pairs, both on a
~670-second pane — each grouped 7 of 8.

**Both aged-pane rounds refused a group. All four fresh-pane rounds refused
none.** That is 2 against 4 and it is not proof; what makes it worth acting on
is the prior, which was already written down: PAIR POSITION says the second
round is worse 20x against better 2x over 39 pairs, and pane age separates
post-retry 0.4 from 4.6.

So it is **staked, not claimed** — `a-fresh-second-round-stops-refusing-groups`,
open in the ledger, judged on the next `--fresh` second-round. FAILED would mean
pane age is not the cause and the PAIR POSITION explanation this project has
carried for dozens of rounds needs replacing rather than acting on. That is the
more valuable half.

### What `--fresh` costs

One extra attempt. The browser closes and reopens, the deck loads, and the pane
comes back CLOSED — recovery reopens it from the ribbon and the round starts on
attempt 2. The sideload survived, which is the whole reason this route was
chosen over reloading the pane. Cheap against a 670-second pane.

### The stake itself was staked wrongly, and the ledger caught it

Written at 22:55 with `madeOn: "2026-08-22"`. `stampDate` compares
lexicographically and a bare date means the START of that day — so the entry was
immediately judged against **all twelve rounds already taken that day, including
159 and 161, the two that motivated it**. It came out `BOTH` (held 10, FAILED
2) on evidence recorded before the change existed.

The first rule this ledger is built on is that a prediction is never judged
against the round that prompted it. A date-only stamp walks straight around it,
and nothing said so. It is now stamped `2026-08-22 23:00` and reads
`no round yet`, which is the truth.

`stakedWithoutATime` warns on the next one. Only when `afterBuild` is absent
from the archive: a build that IS there pins the position exactly and the date
is never consulted, so warning then would be noise on most of the ledger.

Worth noting what caught it: **the population judge built four hours earlier.**
The old single-round judge would have printed one verdict off one round and
nothing would have looked wrong. `held 10 · FAILED 2` is what made the
back-dating visible.

### The remaining in-place decline now carries its evidence

`the chart has no parts list and no readable group members` has fired exactly
once per round since round 151 — deterministic, twelve rounds — and every trace
carried the sentence and nothing else. `groupMembersInOrder` returns `[]` for
three different reasons (no collection queued, items unreadable, fewer than two
members) and nothing could tell them apart. It now records `collection`,
`members`, `nodes` and the target ids. The next round says which.

## Rounds 166 and 167 — the arm was never missing; I was calling the flag the arm

    166  36dc5ec  fresh=false  pane 69s  grouped 9/9
    167  36dc5ec  fresh=true   pane 69s  grouped 8/8

The first pair where both halves record how they were launched. It refuted the
instrument that recorded it.

### The flag is not the variable

**166 ran WITHOUT `--fresh` and started on a 69-second pane anyway**, because a
merge preceded it and recovery reloads a stale pane. So the two rounds differ in
the FLAG and agree on the PANE — and the whole PAIR POSITION argument is about
the pane. `--fresh` is one way to get a fresh one, not the thing itself.

Splitting the arms on `driverRun.fresh` therefore filed a fresh-pane round under
"aged" and made both arms mean nothing. The house defect, in an instrument built
four hours earlier to fix the same defect: **measuring the proxy instead of the
thing.**

Split on pane age instead, over all 18 rounds carrying `driverRun`:

    fresh pane 15/16 round(s) refused no group · aged pane 0/2
      --fresh disagreed with the pane in 1 of them — the flag is not the variable

### And the arm had been in the file all along

`paneAgeAtStartSeconds` reads the smallest `ms` in the trace. Every round has
one. **The arm was never missing** — #685 added `driverRun.fresh` to fix "the
archive cannot say which arm a round was in", and the archive could say, because
the arm is the pane and the pane was always recorded. What was wrong was my
description of the experiment, not the instrumentation.

The flag is kept: it is free, and `flagDisagreed` is now the standing
measurement of how often the proxy and the variable part company. Once in
eighteen so far.

### The prediction: undetermined, and the ledger said held

`a-fresh-second-round-stops-refusing-groups` came out `held 2 · FAILED 0`. Both
halves of that are wrong.

**The claim could not select its own population.** It is about a `--fresh`
SECOND round; `trace-line-present` knows nothing about arms and judged every
eligible round. Round 166 — not `--fresh`, first of its pair — was counted as
evidence for a claim about `--fresh`. A first round on a fresh pane trivially
refuses no group.

**And the absolute form is already refuted by the archive.** Round 150 started
on a 67-second pane and refused a group. `a fresh round refuses no group` is
false as a law; what the archive supports is a rate — 15 of 16 against 0 of 2.
Staking an absolute where the data supports a rate produces an entry that will
one day FAIL for the wrong reason.

Recorded `undetermined` and **not re-staked**. The finding is a rate with a
denominator and the gate line is its right home. Forcing a rate into an
absence-shaped claim kind is what produced the entry.

### A count guard is not an alignment guard, again

The first mutation of the new test — putting `dr.fresh` back — **passed**. The
assertion checked arm SIZES, and the sizes are 1 and 1 under either split. Only
asserting which round landed in which arm catches it. That rule is written down
in this repo and it still took a mutation to notice.

### Round 166 on the scratch fix

`scratch: all | 83 of 83 scratch slide(s) deleted`, and its trace says
`left: 0` — so this round had no unnamed adds and never reached the branch #685
repaired. A regression check on the clean path, not a confirmation of the fix.
The negative case fires in about 18% of rounds; it will show up on its own.

## Rounds 168 and 169 — 13 of 13 has been hiding loose shapes, including tonight

    168  059b4b2  fresh=false  pane 64s  grouped 9/9  13/13
    169  059b4b2  fresh=true   pane 64s  grouped 9/9  13/13

A clean pair. The finding is about every round before it.

### A slide holding 48 shapes, and the battery said 13 of 13

**Eight of the last thirty rounds ended with a slide holding between 11 and 48
shapes. Every one of them reported 13 of 13 scenarios passed.**

    140  [2,4,25,46,24,24,24]   ok
    141  [2,3,48,29,24,24,24]   ok
    142  [2,4,2,11,1,1,24]      ok
    147  [2,4,2,11,1,1,1]       ok
    150  [0,4,2,11,1,1,1]       ok
    159  [0,4,1,2,17,1,1]       ok
    161  [0,4,1,2,17,1,1]       ok
    167  [0,4,1,2,11,1,1]       ok

A clean round ends `0,4,1,2,5,1,1` — nothing above five. A chart that fails to
group is left as its loose shapes, and **no scenario verdict looks at the deck.**

`does a rasterise poison the next draw` is the clearest case. It draws four
charts and reports `all four draws landed`, which is literally true and
narrower than any reader takes it: it asks whether the CALL came back, not
whether a chart survived. It passes with eight loose shapes sitting where a
chart should be.

**Round 167 is in that list, and I called it clean in this journal a few hours
ago.** The gate printed `the deck ended holding 0,4,1,2,11,1,1` on the same
screen I read the verdict from. Nothing compared that line to anything, and
neither did I — which is the whole difference between a number being on screen
and a number being read.

    fullest slide per round, last 8: [5,5,5,5,5,11,5,5]  <- 1 above 6, so a chart was left as loose shapes

Printed as a sequence and never as a failure. The gate's own rule is that it
does not cry wolf on a host whose mood swings 4-of-5 to 1-of-5, and this is a
reason to read a round rather than a verdict on a build.

### Where it came from

A five-lens sweep over the archive, 25 agents, each finding checked by a
verifier told to refute it by default and to re-derive every number rather than
trust the one quoted. It returned 19 confirmed findings and 1 refuted, which is
a suspiciously high rate for adversarial verification and worth saying out loud.

Two have now been checked by hand against the archive. The negative-`left`
finding reproduced exactly and its verifier found the root cause the reporter
had missed. **This one did not reproduce as stated** — the claim was "2 of its 4
charts are left as loose shapes", and round 169 has all four grouped. What is
real is the weaker, older, and more useful version: the verdict CANNOT SEE the
failure, so it passes whenever grouping fails, which is what rounds 159, 161 and
167 did.

So the verifier's REAL was right about the mechanism and wrong about the
present tense. An agent's finding is a lead, not a result — the same rule this
project applies to its own instruments.

## Round 174 — the instrument found the hole in the instrument beside it

`poolFullestSlide` landed in round 173's build. Round 174 is the first round it
judged, and it immediately contradicted the story it was built on.

    GROUPING: 8 of 8 attempt(s) grouped, 0 refused
    fullest slide per round, last 8: [11,5,5,5,5,5,5,11]

**Zero refusals and an 11-shape slide.** The account written a round earlier —
"a chart that fails to group is left as its loose shapes" — named the wrong
mechanism. Nothing was refused. Something threw.

### Grouping has three outcomes and the counter knew two

    grouped the chart's shapes      success
    not grouping: …                 declined by rule
    grouping the chart's shapes     THREW — carries an `error`, usually InvalidArgument

Only the first two were counted. So a round where a group threw printed
`8 of 8 attempt(s) grouped, 0 refused` — **the missing attempt was the defect,
and the count written to expose it hid it.** Both halves of the ratio moved
together, which is the one way a ratio can lie while every number in it is true.

Pooled: **183 throws across 65 of 150 rounds.** They match the loose-shape slides
exactly, with no exceptions in the recent window:

    round  159 160 161 162 163 164 165 166 167 168 169 170 171 172 173 174
    threw    1   0   1   0   0   0   0   0   1   0   0   0   0   0   0   1
    fullest 17   5  17   5   5   5   5   5  11   5   5   5   5   5   5  11

Four for four and twelve for twelve. A chart whose group throws is left as its
shapes; `poolFullestSlide` was seeing that from the deck side while
`poolGroupingOutcome` reported a clean sweep from the trace side.

`attempts per round` now reads a flat `[9,9,9,9,9,9,9,9]`. It was moving before
— 9,9,8,9,8 — and that movement was the throws leaving the denominator. **A
denominator that shrinks by exactly the number of failures cannot show a failure
rate.**

### What it cost, and what it did not

Nothing shipped is wrong: the chart is drawn, and its shapes are on the slide.
What is lost is that it is not a chart any more — no group, so no config tag
through the group handle, and the user sees eight rectangles they cannot edit.
That is the failure `13 of 13` has been passing over, and it happens about once
every four rounds.

### The lesson, which is not the one from an hour ago

I wrote yesterday that a summary can outlive the regime it summarised. This is
narrower and worse: **a count can be wrong at the moment it is written, because
the thing it does not count also leaves the denominator.** No staleness
involved. `poolGroupingOutcome` has counted two of three outcomes since the day
it was written, and every round it ever reported read as perfect.

The sweep flagged this — "the grouping instrument drops a chart whose addGroup
threw from BOTH the numerator and the denominator" — and I had it in the queue
behind four other findings. Round 174 produced it live before I got there, which
is the better provenance: the archive said it, not an agent.

### CORRECTION to the entry above: the correspondence is ONE-WAY

The entry above reads "four for four and twelve for twelve" from rounds 159-174,
and the commit message for #693 repeats it. That window is real and the general
claim it invites is not. Over all 150 archived rounds:

    threw AND ended over-full        66
    threw, NOT over-full              0
    over-full, did NOT throw         27
    neither                          59

**A throw always leaves an over-full slide — 66 for 66, no exceptions.** That
half is exact and it is the half the fix rests on.

**An over-full slide does not imply a throw.** 27 rounds ended over-full having
thrown nothing: 028, 066, 067, 069, 070, 075, 076, 083, 095, 099, 109, 114, 116,
118 and more, at 11, 24, 28, 32 and 38 shapes. Those are the REFUSAL era — the
fresh-slide charts that answered `not grouping: no member handle this host will
accept`, which stopped after round 142.

So there are two mechanisms that leave a chart as loose shapes, and only one of
them is live:

    refused (`not grouping`)   the old fresh-slide failure, ~0 since round 142
    thrown  (`addGroup`)        the current one, ~1 round in 4

The recent window looks like a perfect correspondence because only one mechanism
is still firing in it. Reading a law off sixteen rounds when the archive holds
150 is the same mistake as reading a rate off a window that spans a regime — the
thing this journal has spent two days documenting — committed while documenting
it, in the sentence announcing the fix.

Caught by testing my own claim against the whole archive rather than the window
I had already looked at. One query.

## Rounds 177 and 178 — the cleanest pair yet, and it disagrees with itself

    177  0a5dd16  pane 65s  grouped 8  threw 1  fullest 11  13/13
    178  0a5dd16  pane 64s  grouped 9  threw 0  fullest  5  13/13

Same build, nothing merged between them, **panes one second apart**. Pane age is
fully controlled here and one round still threw. So the addGroup throw is not a
property of a build, and **no single clean round can ever show it fixed** — that
is what the pair bought.

### The precursor is the whole story

    177  the re-read named none of the chart's shapes: listed:9  → positional fallback → THREW
    178  the re-read named none of the chart's shapes: (none at all)     → normal path → clean

The difference is not the guess going wrong. It is **whether the pre-grouping
re-read fails at all.** That reframes the chain:

1. the re-read matches none of the chart's own ids — the root, ~1 round in 4
2. the positional fallback fires, taking the TAIL of the host's listing
3. whether that throws depends on whether the tail names dead or loose shapes

Everything written about the throw so far has been about step 3. **Step 1 is the
one to chase**, and nothing yet says why a re-read that had `withOwnId: 7` — all
seven of our shapes carrying ids — matches nothing in a listing of nine.

### TWO CORRECTIONS to what this journal said about the throw

**"183 throws across 65 of 150 rounds" was two different failures conflated.**

    ShapeCollection.getItem    155 events, rounds 023..065   code 5010
    ShapeCollection.addGroup    29 events, rounds 068..175   InvalidArgument

Zero overlap, a clean boundary at ~066. The getItem population is a different
defect that stopped at round 065 and has not recurred. **The real addGroup
population is 29 throws across 26 rounds**, and #693's commit message carries the
wrong number.

Both were pooled under one trace message, which is why one query answered "how
often does the group throw" with a number spanning two eras and two mechanisms.
The message name was the same; the `errorLocation` was not, and nothing read it.

**And `fullest = 5 + 6 × (throws + declines)` reproduces 22 for 22, no
residual** — each failed grouping converts one top-level group into seven loose
shapes. So the over-full slide is fully derived, not a separate symptom.

### What is inferred and NOT observed, said plainly

The stale tail explains the throw when it names shapes already absorbed into a
group. **When it names shapes that are still LOOSE, the same guess succeeds — on
the wrong chart's shapes — and feeds them to the parts tag as this chart's own.**
It throws nothing and traces nothing.

That is a prediction of the mechanism, not a finding. The success line carries
`{charts, partial, by}` and no member ids, so the archive cannot distinguish a
correct group from a mis-group. Round 161 is suggestive — three consecutive arms
declined, threw, and "succeeded" on `listed` values of 1, 15 and 16 — and
suggestive is all it is.

So it becomes an instrument rather than a claim: the fallback now records
`mine` and `chose`, the ids rather than a verdict, and triage splits them into
own / another chart's / a mixture. A partial pick counts as a mixture and never
as "own" — rounding that toward the reassuring bucket is how it stayed invisible.

### Provenance

A four-angle workflow, each account checked by a verifier told to refute it and
specifically to test the SUCCESS-case counterfactual. That check is what makes
the account honest: `re-read named none` precedes all 29 throws, and it also
appears 95 times in total, splitting 29 throw / 39 grouped-OK / 27 declined.
**Necessary, not sufficient** — and an account of a rare failure that never
checked its signal against the successes is not established however good it
sounds.

Everything load-bearing above was re-derived by hand before being written down.
The 155/29 split, the 22-for-22 formula, and the code path were all checked
directly; the wrong-chart success was checked and found NOT observable, which is
why it is an instrument and not a finding.

## Rounds 179 and 180 — the instrument answered on its first round

Round 179, the first round carrying the traced ids:

    mine  ["35","36","37","38","39","40","41"]
    chose ["27","28","29","30","31","32","33"]   listed: 8

**Zero overlap**, with `withOwnId: 7` — every one of our own ids in hand, and the
guess took a different chart's seven shapes. The host's listing was one
grouping-event stale and its tail named the chart drawn before this one.

That round threw, because 27-33 had already been absorbed into their own group.
**The same guess on shapes still loose would have SUCCEEDED** — grouping another
chart's shapes under this chart's name and writing them into its parts tag, with
nothing thrown and nothing traced. One draw's timing away from the 29 archived
throws.

Yesterday this was an inference the archive could not reach. It took one round to
become an observation, which is what an instrument is for.

### The fix, and what it does not fix

`positionalGuessNamesOurs` refuses the guess when we have ids of our own and the
pick contains none of them. Refusing is strictly better than guessing wrong: an
ungrouped chart keeps its own shapes and loses its group, which is measured and
recoverable; a wrongly grouped one takes another chart apart and reports success.

A partial overlap is allowed through on purpose — a mixture means the listing is
only partly stale, and the one-for-one checks downstream judge it. When the host
reads no id back at all there is nothing to compare and the positional rule stays
the last resort it was added for.

**The root is untouched.** The chain is:

    1. the pre-grouping re-read matches NONE of our own ids   <- unexplained
    2. the positional fallback fires
    3. it may pick another chart's shapes                     <- fixed here

### The root's rate, and the test for the fix

Over rounds 157-180:

    re-read named none fired in   7 of 24 rounds
    every one of those 7 threw
    the other 17 threw nothing

Per EVENT the precursor is necessary and not sufficient — 95 events across the
archive split 29 throw / 39 grouped-OK / 27 declined. Per ROUND in this window it
is 7 for 7. Both are true and the second is what makes the fix testable:

> **A round where `the re-read named none of the chart's shapes` fires should now
> record `not grouping: the positional guess named no shape of ours` INSTEAD of
> an `addGroup` throw.**

**Round 180 does not test it.** It is clean — 9 of 9, no throw, fullest 5 — and
it carries no positional-guess line at all, so the guard was never exercised.
177/178 established that the throw varies within a build, so a clean round is not
evidence either way. The condition fires about once in three and a half rounds;
the next one that fires is the test.

**NOT STAKED IN THE LEDGER, deliberately.** The claim is conditional on the
precursor firing, and `trace-line-present` cannot express "when X happens, Y
rather than Z" — an absolute would read FAILED on every round where the fallback
simply never ran. Forcing a conditional into an absence-shaped claim kind is
exactly what produced #684's false `held`. The criterion is written here instead,
where it can be checked by eye against one grep.

## Round 181 — STOPPED: the tab is on a Microsoft sign-in error

The night ends here, on the credential boundary. **This is the owner's to clear
and nothing was touched.**

    heading "Sign in"
    "Sorry, but we're having trouble signing you in."
    AADSTS900561: The endpoint only accepts POST requests. Received a GET request.
    Timestamp: 2026-08-23T03:09:02Z

`--fresh` closes the browser and reopens it at `https://onedrive.live.com/`.
That navigation landed on the sign-in flow and the flow errored, so the tab holds
an error page instead of a deck. Everything the driver then reported follows from
that: `deck ? slide(s)` because there is no deck, `the pane is closed` because
there is no document, and seven attempts because none of it is recoverable by
reloading.

### The diagnosis took four wrong turns, and three of them were mine

**1. "The widen is too narrow."** REAL, and fixed in #698 — `ensureRibbonRoom`
asked for 1900px against a `MIN_RIBBON_WIDTH` of 2375, so it could not reach the
bar it exists to clear. It had been wrong for weeks and the device scale factor
hid it: a request of 1900 usually ARRIVES as 3800 CSS px, so every round printed
"widened to 3800px". `--fresh` reopened at 1:1 and 1900 arrived as 1900.

**Not the cause of this failure.** The retry widened to 3088px and failed
identically. A real defect found while chasing a different one — worth having,
and it is not the answer.

**2. "The session key is wrong."** WRONG. `list --all` showed the browser under
workspace `/` and the driver keys by cwd string, which is the trap that cost an
afternoon earlier in this archive. But running `list` from
`C:\devtools\PowerChart\.pw-session` shows it fine. The key was never wrong.

**3. "The browser is named `ms`, so session-keyed calls miss it."** WRONG, and
this one is instructive: `snapshot` from my shell answered *"The browser
'default' is not open"*, which reads exactly like a driver defect. The driver
passes `-s=ms` on every call. **My probe was mis-specified and I nearly filed its
output as a finding about the driver.** The rule this project applies to agents
applies to hand queries too: an instrument that disagrees with a working system
is usually the instrument.

**4. The actual answer** took one command — `-s=ms --raw snapshot` — and was
sitting on the page the whole time.

### What is landed

Rounds 154-180 are archived and every fix is on main at `e47fe6b`. Round 181 has
no archive because it never ran; that is correct, not a gap.

### What the next session needs to know

The signature of a sign-in wall, so it is not chased as an add-in or a pane
problem:

    deck ? slide(s)   AND   pane ?   AND   host not asked
    on EVERY attempt, from the first

A lost add-in looks different: the slide list READS and `button "Insert chart"`
is absent. A stale pane reads a build stamp. **Only a dead document reads
nothing at all from the first attempt**, and the one command that settles it is
`-s=ms --raw snapshot` from `.pw-session`.

## CORRECTION — `shape.id` is not volatile on this host, and the chain has a different root

Several entries above, and the commit messages for #696 and #701, rest on a
premise that is false: *"the product identifies shapes by `shape.id`, and this
host reassigns them."* It does not.

**The proof is in the round those entries used as their own evidence.** Round 179
recorded the positional guess taking `chose [27..33]` while the chart had drawn
`mine [35..41]`, and that was read as the ids having moved. The same round's
end-of-round deck inventory, slide `260#1666125241`:

    {"id":"35","name":"title"}          {"id":"38","name":"seg-0-0"}
    {"id":"36","name":"category-0"}     {"id":"39","name":"seg-0-1"}
    {"id":"37","name":"category-1"}     {"id":"40","name":"baseline"}
                                        {"id":"41","name":"series-label-0"}

Seven ids, their real names, still there when the round finished. **They never
moved.** Two more, independently:

- 228 in-place updates across 30 rounds resolved **4,992 tag-stored shape ids**
  written in an earlier context.
- `addGroup` keyed on id grouped **2,108 charts** against 36 throws.

### The actual root, which the host has been answering every round

The re-read returned `[27..33]` because **the LISTING is stale** — it held the
previous chart's shapes, not ours. Matching failed for want of the right listing,
not the right key. The probe sheet says so and has since the beginning:

    shapes-items-count-honest    unreadable 140   short-0 16     (156 rounds)
    tag-through-refetched-shape  no-id      149 of 149

**The shape collection has never once honestly reported freshly added shapes.**
That answer has been on every sheet, read as a fact about a probe rather than as
the cause of the product's only live defect.

### What it costs the creationId plan

Nothing survives of it as a fix for this chain. Matching on creationId against
the same stale listing matches zero, exactly as id does. And there is no
`getItemByCreationId`: `getItem`, `getItemOrNullObject` and `addGroup` all take
ids, so creationId can LABEL a shape and never ADDRESS one.

The three probes are kept, for two reasons that are not this one. The BACKLOG
item they were half-aimed at — "Retire the positional group-member mapping" — is
about node-to-shape ordering INSIDE a group and is untouched by any of the above.
And the contract is worth having on the sheet before the next person reaches for
it on the strength of the name, which is exactly what happened here.

### How the error was made, because the shape of it matters

The evidence was read once and not cross-checked against the deck. `mine` and
`chose` disagreeing is equally consistent with "our ids changed" and "the listing
is of something else", and the second was never tested — although the round file
carried the inventory that settles it, in the same JSON, two fields away.

It is the house defect one more time: **a single reading promoted to a cause.**
The rule this archive already carries — cross-check every number against a second
source in the same file — was written down twice tonight and not applied to the
claim everything else was being built on.

### What the next work is

Upstream of the match, not at it: **make the pre-grouping re-read return the
shapes just drawn, or stop depending on a re-read at all.** The guard shipped in
#696 stays — it converts a certain double loss into a loose-but-tagged chart —
but it is a seatbelt, not the fix.

## The product sweep, verified by hand — eight claims, five wrong

A four-lens workflow asked what 180 rounds teach about the PRODUCT rather than
the harness. Every headline was then re-derived here. **Five of eight did not
survive contact with the archive**, and every one of those five had something
real underneath.

| claim | verdict |
| --- | --- |
| Set Same Scale costs ~2½ minutes | **confirmed** — 158s median, 3x the next scenario |
| the parts list has never been produced | **confirmed** — 0 across 872 charts |
| `Shape.group` throws every update round | **confirmed** — exactly 2/round, deterministic |
| `setSelectedShapes` never wedges this host | **confirmed** — 1,092 measurements, 0 wedges |
| every chart group is named `PowerChart` | **confirmed** — 1,782 in the inventories |
| `shape.id` is volatile | **refuted** — see the correction above |
| the repair pass never runs | **refuted** — it runs every time and FAILS 137 of 172 |
| a 100% tag route the code ignores | **refuted and inverted** |

### The three that matter most

**`Shape.group` poisons a batch, twice a round, forever.** Microsoft's reference
says `.group` returns GeneralException when the shape is not a group. The
archive has it 76 times across 38 rounds and **exactly 2 in every one of the last
20**, always `statement: var group = itemOrNullObject1.group;` — and the next
trace line is `a by-id lookup refused the whole resolve — re-reading the slides
instead`. The comment above it claimed "it costs no extra round trip". It costs
a whole extra resolve.

It fires on every chart because it is guarded by `parts.length`, and the parts
list has never been produced. **So the parts repair may remove this by itself** —
a chart arriving with parts never queues `.group`. Left for a round to measure
rather than restructured on the theory; that batch has three shipped-broken
fixes on record behind it.

**Set Same Scale costs 158 seconds.** Median over the last 20 rounds, against 49s
for the next-slowest scenario. And it is not setup: `sameScaleAcrossDeck`
DISCOVERS its charts with `probeCharts` rather than drawing them, so that is the
real cost of the user pressing the button on eight charts — about 20s each.

**The wedge this project designed around does not happen.** 1,092 selection
measurements, zero wedges, and `which selection call wedges the host` has passed
156 of 156 with "the host answered all 7 rung(s)". office-js#3698 reports
`setSelectedShapes([])` never resolving; on this host, across the whole archive,
it always has.

### The tag-route inversion, because acting on it would have hurt

The claim was that a route measured at 100% is ignored while the code prefers one
that fails. Failures by route, all 156 rounds: `created` 235, `undefined` 61,
`group` 51, `by-id` 26. **The code prefers `by-id` — the fewest failures.** The
probe the claim rests on is measured on a scratch shape, and production
disagrees with it 235 times.

But neither number is a rate: the tag write recorded FAILURES by route and
successes not at all. Both the claim and its opposite were unsupported, and the
missing denominator now ships.

### The rule this is now the second demonstration of

Across both sweeps: **eight headlines hand-checked, five overstated or inverted,
all eight with something real underneath.** Sub-agents locate unread signals
reliably and interpret them badly. Take the LOCATION as the deliverable and treat
the CONCLUSION as a hypothesis — and re-derive every number from the archive
before repeating it to anyone.

## Round 181 — the parts-list fix is INERT, and the reason is the whole story

    181  1e8455e  ok 13/13 | grouped 9 | refused 0 | threw 0 | fullest 5

A clean round, and the first to run tonight's product changes. **One of them
does nothing.**

### What was claimed and what happened

#702 changed `ungroupedFallback` to take the ids already in hand rather than
re-read them, on the strength of `withOwnId: 7` in the re-read trace. Round 181:

    PARTS where= the id read-back threw   charts=1  gotPartsList=0  grouped=0

The new `the ids were already loaded` exit was never taken. `needLoading` was
full, the 5010 fired exactly as before, and the loose chart got no parts list.

**The premise was wrong in a way worth naming.** `withOwnId` is measured inside
`chooseGroupMembers`, and that only runs when a refresh was asked for. Round 181
carried no `re-read named none` at all, so there was no such reading — and in
the ordinary path the created proxies carry no loaded ids at fallback time.
Which is precisely why the code loads them.

I read one field, saw 7 of 7, and generalised it to a path where it is not
measured. Same shape as the `shape.id` error four hours earlier: **a reading
taken in one context, asserted about another.**

The change is KEPT because it cannot cost anything — it only skips a call whose
answer is in hand — but it is not a fix and the parts list is still 0.

### The obvious next step is a trap already paid for

"Load the ids in the drawing batch" is what `tagAnchorIndex` spent four rounds
and a host probe on. From this file's own history:

> the tag write goes through the handle that DREW the shape, and **this host
> refuses such a handle once a `load()` has resolved it into
> `shapes.getItem(id)`**

That is the mechanism behind `from: created` — **235 tagging failures, the
largest bucket in the archive.** Loading ids earlier buys a parts list and spends
the tag write that makes the chart re-editable at all. It is a trade, not a fix,
and it is a bad one.

### One root, three symptoms

The parts list is blocked behind the same thing as the grouping and the re-read:

    shapes-items-count-honest    unreadable 140   short-0 16    (156 rounds)

**The shape collection will not honestly report freshly added shapes.** That one
fact produces the failed re-read, the positional guess, the addGroup throw, and
the missing parts list. Every workaround attempted at the symptom has cost more
than it bought — the tag-anchor move (reverted), the id load (implicated in 235
failures), and now this.

Fix the collection read and three symptoms go with it. Work around it downstream
and the workaround is the next defect.

### What is confirmed working

The widened window: 2020px → 6176px on the live host, against a 2375px
threshold. Round 181 started first time on it after the sign-in scare.

The positional guard has still not been exercised — no guess fired in 181.
`Shape.group` still threw its deterministic 2, unchanged, because the parts
list is still 0 and every chart therefore still reaches it.

## CORRECTION — the collection read works 95.7% of the time

Two entries above say **"the shape collection has never once honestly reported
freshly added shapes"**, and one calls it the single root of every live defect.
It is a SCRATCH-SLIDE answer quoted as a production fact.

    charts grouped by id-match   2135
    charts grouped by created      49
    re-read named NONE             97   ->  4.3% of charts

**The re-read matches ids for 2,135 charts.** The claim came from
`shapes-items-count-honest` — `unreadable` 142, `short-0` 16 of 158 — which runs
on the scratch slide, and this repo already documents that the scratch slide is
strictly worse at collection reads than a real one. It says so about three other
questions on the same sheet.

### The third time tonight, and the second after writing the rule down

    the ids are reassigned        a mine/chose mismatch read as volatility
    the ids are already loaded    withOwnId: 7 from the re-read path,
                                  asserted about the fallback path
    the collection never answers  a scratch-slide probe, asserted about
                                  production

Every one is the same move: **a reading taken in one context asserted about
another.** Each was caught by one query against the archive, and each had
already been written into a commit message before it was caught.

The pattern is not carelessness about numbers — every number quoted was real.
It is carelessness about SCOPE. The number is measured somewhere, and the claim
is made everywhere.

### What is actually true, which is narrower and more useful

The collection read usually works. It fails about **one chart in twenty-three**,
roughly one round in four, and when it fails everything downstream fails with
it — the re-read matches nothing, the positional guess takes another chart's
shapes, addGroup throws, and no parts list is written.

That does not weaken the case for `shapes-by-index-vs-items`; it sharpens the
question. Not "can an index walk replace a read that never works" but **"does an
index walk survive the 4% where items does not"** — which is answerable, and is
what the probe now asks.

### Rounds 181-183, and a fourth that never ran

    181  1e8455e  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5
    182  1e8455e  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5
    183  7eaa1ae  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5

Three identical rounds across two builds. No guess fired in any of them, so the
positional guard is still unexercised — and at 4.3% per chart on nine charts a
round, three clean rounds is unremarkable rather than evidence of a fix.

**Round 184 refused before it started, and it was my doing.** I checked out a
branch to write the correction above while 184 was already running. The driver
re-reads git HEAD on every retry, so it compared the branch commit against the
deployed main and refused with site-behind:

    HEAD c1fab2d · site 7eaa1ae

**That is the third time**, and it is written down in this project's own notes:
do branch work before or after a round, or in a worktree. The refusal is the
driver working — it caught a comparison that would have been meaningless — but
the round is lost and the pair with it.


## Round 184 — two leads closed in one round, both by measurement

    184  a133cf7  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5

The fourth identical round in a row, and the first to carry the two probes
written to settle the product's open questions. Both answered, both negative,
both worth more than the plans they killed.

### `Shape.creationId` does not exist on this host

    creationid-on-fresh-shape     => absent | the property is not on the shape at all
    creationid-survives-a-sync    => no-creation-id | before=absent
    creationid-survives-grouping  => no-creation-id | before=[absent,absent]

Not null. Not equal to `id`. **Absent** — the property is not on the object,
though the host advertises requirement set 1.10, which is where Microsoft
documents it.

So there is nothing to record at draw time, nothing to match on, and nothing to
migrate to. The route is closed.

**This is the clearest vindication of the night's discipline.** The backlog
called creationId "a durable per-shape identifier" and cited no source; I read
the same word off the API name and proposed a migration on it. The documentation
turned out to be one sentence that promises nothing and permits null. Three
probes and one round settled what a migration would have discovered after
however many rounds of building on a property that is not there.

### The index walk is closed too, and its premise was a fourth scope error

    shapes-by-index-vs-items => no-count | getCount did not populate — the premise is gone

The probe was built on `getcount-populates-same-sync` answering `yes` 158 of 158.
**That probe measures `ctx.slides.getCount()` on the DECK.** It carries
`noSlideNeeded: true` and the comment "No slide of ours involved". It never
said anything about `slide.shapes.getCount()`, which is what an index walk needs.

On the real host `shapes.getCount()` does not populate either. There is no length
to walk and no index route around the collection read.

That is the fourth scope error in a night — a number measured in one place,
asserted about another — and the first one caught BEFORE it became a design
rather than after it became a commit message. The probe cost one scratch slide.

### Where that leaves the product

Two cheap ideas are gone and the picture is simpler for it:

- the collection read fails about **4.3% of charts**, roughly one round in four
- when it fails, the re-read matches nothing, the positional guess takes another
  chart's shapes, addGroup throws, and no parts list is written
- there is **no alternative identifier** (creationId is absent) and **no
  alternative enumeration** (getCount does not populate on a slide's shapes)

So the next work has to be about the 4% itself — why the read misses, and
whether anything the product already holds can tell a stale listing from a good
one before it acts on it. Not about replacing the mechanism.


## STOPPED — a real sign-in prompt, and this time it is the genuine article

    0: (current) [Sign in to your account](https://login.microsoft.com/consumers/fido/get?...)

One tab, `login.microsoft.com`, a FIDO flow. **The page was not touched and no
host work continues.** Signing in is the owner's alone.

### It is distinguishable from the false alarm eight hours earlier, by the test
### written down after that one

    earlier   title 'Home - OneDrive', account active, AADSTS900561 on a redirect
              -> the session was ALIVE and stopping was a misread
    now       title 'Sign in to your account', login.microsoft.com, sole tab
              -> the session is gone

The rule from ROUND-LOOP-BRIEF, applied and working: *a dead browser is not a
lost sign-in — only a redirect to a login host needs the owner.* Checking the
TITLE is what separates them, and it took one command.

### How it got here

Round 185 ran with `--fresh`, which closes the browser and reopens it. The
persistent profile has carried the sign-in across every previous `--fresh`; this
time it did not. Seven attempts, each reporting `deck ? slide(s)` and a closed
pane, because the browser was sitting on a sign-in page the whole time.

Round 185 has no archive. That is correct rather than a gap: nothing ran.

### State at the stop

Rounds 001-184 archived. Everything landed and green. The last four rounds are
identical on every counter:

    181  1e8455e  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5
    182  1e8455e  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5
    183  7eaa1ae  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5
    184  a133cf7  13/13 | grouped 9 | threw 0 | GenEx 2 | gotParts 0 | fullest 5

### To resume

Sign in, open `Presentation64.pptx`, and confirm **Insert chart** is in the Home
ribbon. Then `node scripts/round.mjs --check --dir .pw-session` confirms the
pane, and the loop restarts at 185.

If the add-in is gone from the ribbon after signing in, that is the agent's to
fix — `sideloadAddIn` uploads a manifest and never asks for a password.


## The settle retry is load-bearing, and a second ask costs a second a round

Pooled over 161 archived rounds:

    settle retries fired                 1057
    failures that SURVIVED the retry       97
    rescue rate                          90.8%

**The first read of the shape collection fails on roughly half the charts this
loop sees** — 1,057 retry passes against 2,135 successful groups — and a pause
plus a fresh slide handle fixes nine in ten of them. The 4.3% quoted everywhere
above is what is LEFT after the retry, not the failure rate of the read.

### The constant was set on an argument the code contradicts

`REREAD_ATTEMPTS` was 1, with this reasoning:

> a third would only be waiting on a host that has already answered the same
> way twice, at a second and a half a go, on the path a deck-wide update runs
> for every chart

**Both halves are wrong.**

The host does NOT answer the same way twice — it changes its answer 91% of the
time, which is the entire reason the first retry exists and works.

And the pause is NOT per chart. The `await` sits under `if (attempt > 0)`, once
per PASS, covering every chart still pending, and `pending` shrinks to `retry`
each time round. A further attempt costs one pause only while charts remain —
97 events across 161 rounds, **about one second per round**.

### What it targets

Those 97 are exactly the charts whose re-read named nothing. They then take the
positional guess, throw addGroup, and lose both the group and the config tag —
the one live user-visible defect left in this product.

### The criterion, in advance

**HELD** means `the re-read named none of the chart's shapes` falls against its
own recent rate, judged as a RATE across every eligible round. The event fires
about 0.6 times a round, so a single clean round proves nothing — that is the
n=1 mistake that made #684 read `held` on evidence predating the change.

**FAILED** means the survivors are not slow but structurally wrong — a listing
one grouping-event stale, which no pause can cure — and that is the more
valuable outcome: it closes the last cheap idea and points the next work at
re-fetching rather than waiting. Two leads have already closed that way today.

**REVERT** if the failure rate does not fall, if a round gets materially slower,
if any scenario fails, or if any counter moves the wrong way. Three
shipped-broken fixes in this repo came from changing this path on a theory, so
this one names its refutation before it runs.

### Not staked in the ledger, deliberately

The claim is that a RATE falls. `trace-line-present` can only assert a line is
ABSENT, and staking an absolute where the data supports a rate is exactly what
made #684 come out `held` on rounds recorded before the change existed. The
ledger's claim kinds cannot express this one, so it lives here where a grep
settles it.


## The survivors stopped EIGHT ROUNDS BEFORE the fix aimed at them

    rounds 150-179   12 survivors over 30 rounds   =  0.40 per round
    rounds 180+       0 survivors over 12 rounds   =  0.00 per round

    167:1  174:1  175:1  177:1  179:1   <- the last one
    180..191: 0 0 0 0 0 0 0 0 0 0 0 0

#709 raised `REREAD_ATTEMPTS` to 2 to rescue exactly this population, and it
landed at round **187**. The last survivor was round **179**.

**So the five clean rounds since the change are not evidence for it.** Reading
them as a win would be the regime error this journal has spent two days
cataloguing: a number that moved for a reason, credited to a change that came
afterwards.

### What this means for the change

It may be permanently unexercised. If `the re-read named none of the chart's`
`shapes` has genuinely stopped firing, the second ask has nothing to rescue and
will sit there costing nothing and doing nothing — which is a perfectly fine
state for a guard, and a useless one to keep claiming credit for.

What it is NOT is refuted. The mechanism is proven: with the fault injected,
reverting the constant leaves the chart ungrouped and the test fails with
`expected [ 1 ] to include 2`. The second ask works. There is simply nothing
for it to work on at the moment.

### What actually changed at round 180 is not known

Round 180 was the first on a5e6ab5 — #696, the positional guard. But that
guard fires AFTER `the re-read named none` is traced, so it cannot suppress the
line. Candidates left: the host, the deck's state, or coincidence. 12 rounds at
zero against a 0.40 rate is P(0) of about 1%, so coincidence is unlikely but
not excluded, and nothing here identifies a cause.

**Recorded as unexplained rather than attributed.** The temptation is to credit
#696 because it is nearby and it is mine; the trace order says it cannot be the
reason, and 'a change I made happens to sit near the inflection' is not an
argument.

### The rule this is the fifth instance of

Before crediting a change, check when the number actually moved. Four times
today a reading was asserted outside the scope it was measured in; this is the
same failure across TIME rather than across context — and it would have been
invisible without laying the per-round series side by side, which took one
query.


### Following it up: #696 exonerated, and the variable nobody measures

The entry above left three candidates for the round-180 inflection and leaned
on trace ORDER to clear #696. Order is weak evidence, so here is the diff.

```
985c547..a5e6ab5  =  one commit, one file, and all of it downstream:

  + if (!positionalGuessNamesOurs(mine, chosenIds)) {
  +   trace("group", "not grouping: the positional guess named no shape of ours", ...)
  + } else {
      freshMembers.set(i, chose);
      wholeMatch.add(i);
  + }
```

It decides whether to ACCEPT a guess. The survivor line is traced at
powerpoint.ts:8989, the guard runs at :9063 — 74 lines later, in the same
pass. Nothing in the commit can reduce a count that was already emitted.
**#696 is not the cause.** Not by position: by content.

### The denominator held perfectly, which makes it a real change

    175..184   passes 5   charts [1,1,1,1,1]   every single round

The retried population is not merely similar across the inflection, it is
IDENTICAL — five passes, one chart each, ten rounds running. So this is not a
case of the path going quiet and taking its failures with it. The same work
happens and stopped failing.

### And the variable that could explain it is not recorded at all

    distinct host strings across 167 rounds: 1
      "PowerPoint · OfficeOnline · 0.0.0.0"

Office Online does not version itself to an add-in. Every round agrees about
the host because every round is reading the same placeholder — **a constant
that is constant because nothing measures it, which is indistinguishable on
the page from a controlled variable.** A host-side change cannot be ruled in
or out here, and the archive has been quietly presenting that as though the
host were pinned.

That is now printed under the finding rather than left for a reader to
assume. It is the same house defect one more time — an unmeasured thing
reading as a measured one — and it took a chase for it to surface at all.

**The honest state: the number moved for a reason outside our source, and
this archive cannot name it.** The instrument that would (a real host
fingerprint off the page, not `Office.context.diagnostics`) needs the driver idle
to build, so it is the next thing to try between pairs — not mid-round.


### The host did not change either, and my own p-value was inflated

Chasing the round-180 inflection into the probe sheet. Splitting all 40
questions at the inflection and comparing their answer SETS reported seven
changes — of which four were probes added later (`creationid-*`,
`shapes-by-index-vs-items`), i.e. not-asked-before, not changed. The three
that looked real:

    picture-then-shape-read                       unreadable  ->  unreadable|yes
    how-many-collection-reads-a-context-survives  unreadable-at-1 -> short-at-1|…
    shape-add-held-slide-proxy                    threw|yes   ->  threw

The first two point the same way — collection reads becoming MORE readable —
which is exactly the mechanism a survivor is made of. It fit beautifully.

**It is not true.** The series:

    picture-then-shape-read     168..183 unreadable, 184 YES, 185..191 unreadable
    how-many-collection-reads   168..189 unreadable-at-1, 190 SHORT, 191 unreadable

One sample each, at rounds 184 and 190. Neither is at 180. Nothing changed at
the inflection; two rare faces turned up afterwards and my comparison could
not tell them from a regime change.

**A SET COMPARISON SCORES 1-OF-12 THE SAME AS 12-OF-12.** That is the defect in
the instrument I had just written to check the last one. Comparing which
answers OCCUR discards how OFTEN, and how often was the entire question. It is
the same shape as pooling survivors into a total and losing the timing — the
thing this morning was spent on. Twice in one session, one layer apart.

### And the 1% was not 1%

The earlier entry said 12 rounds at zero against a 0.40 rate is P(0) of about
1%, which is arithmetically right and rhetorically wrong: **the split point was
chosen AFTER seeing where the zeros started.** The honest question is not "how
unlikely is a 12-run at this spot" but "how unlikely is the longest zero-run
ANYWHERE in a 30-round series", and that is a much weaker claim. Quoting the
first as though it were the second is how a post-hoc split becomes a finding.

### Where that leaves the inflection

    not our source    #696 is downstream of the trace line, by diff, not by order
    not the probes    every question settled across 180; the two flips are singletons
    not measurable    the host reports 0.0.0.0 in all 167 rounds
    not significant   at least, not at the strength first claimed

**So: unexplained, and weaker evidence than it looked.** The one thing that did
hold up all the way through is the denominator — five passes, one chart each,
every round across the boundary — so whatever moved, it did not move by the
work going away. That is the only load-bearing fact here, and it is enough to
keep watching without being enough to conclude anything.


## The biggest user-facing number in the project, and the model under it is unmeasured

Chasing the product question rather than the harness for once. Scenario
medians over 169 rounds:

    same scale across the deck                166s   <- 38% of the round
    does a rasterise poison the next draw      50s
    insert onto a slide that already has …     48s
    …
    13 scenarios, medians summing to 431s per round

166s looks like a Set-Same-Scale problem and is not one. It reports
"8 of 8 charts carry the shared scale" — eight redraws, ~20s each. And the
single-chart scenarios cost the same: 23s to edit the selected chart, 18s to
edit one on the visible slide. **Nothing is wrong with same-scale. One chart
redraw costs about twenty seconds, and that scenario does eight of them.**

That is the number a user feels, on every edit.

### The instrument existed and had never been pooled

`batch issued` carries `prevBatchMs`. 2,913 measured batches:

    LIVE (visible slide)   n=  336   median 12283ms per batch
    OFF-SCREEN slide       n= 2577   median  5330ms per batch

A batch on the slide the user is looking at costs **2.3x** one that is not,
which is the repaint the code already suspects. ~140s of every round is spent
inside these batches.

### A wrong answer I published to myself for ten minutes

The first cut of this said **10 shapes and 20 shapes cost the same, so the
payload is free and the batch size should go up.** It was wrong, and wrong in
the way that is hardest to see: `upTo` is a RUNNING TOTAL, not a batch size.
The `"20 shapes"` bucket was the second batch of ten. Every batch in the
archive is ten. There is no size variation in 169 rounds at all, so **the
archive cannot compare batch sizes** — the comparison I thought I had was two
samples of the same size, one later in a sequence than the other.

Retracted before it reached a commit, by reading the raw records instead of
the aggregate. An aggregate over a misread field is confident and clean and
says nothing.

### What the constant rests on

Two docstrings justify the batch sizes, and both price a round-trip the same way:

> Ten is comfortably under the last known-good (18) and still coarse enough
> that **the round-trips (~0.1s each)** disappear next to the drawing.

> a stacked chart at 10 shapes takes 4 syncs, at 40 it takes 1; **the extra
> round-trips (~0.1s each) dominate** a 60-slide run's wall clock.

**These cannot both be true.** Three extra trips at 0.1s is 0.3s; 0.3s does not
dominate a 60-slide run, and it does not need justifying against drawing
either. One sentence uses the figure to argue the trips are negligible and the
other uses the same figure to argue they are decisive.

And nothing here settles it. A batch's `prevBatchMs` is **trip plus draw**, so
12.3s live is consistent with a 0.1s trip and 12.2s of drawing, and equally
consistent with the reverse. The measurement does not refute the model; it
shows the model has never been measured.

**So: the constant governing 38% of a round's wall clock, and every second a
user waits on an edit, rests on a cost split nobody has ever separated.**

### What NOT to do, written down because it is the obvious move

Raise `SHAPES_PER_SYNC`. Ten sits under a measured ceiling of about eighteen —
"a 30-shape butterfly never commits in 90 seconds" — and this project has
already raised a threshold on a 4% measured win and crashed PowerPoint 6 runs
out of 6. A batch that is too big does not get slower, it stops answering, and
nothing lands at all. The headroom is real and it is not free.

The next step is the instrument, not the constant: time `context.sync()` apart
from the statements queued before it. The adds are local and the sync is the
trip, so the two ARE separable — which means the 0.1s claim is testable, and
until it is tested the batch size is not a decision anyone can defend.


### Settled the same hour, by a number already printed on every round

The entry above said the two docstrings could not both be true and that
nothing here separated trip from draw. The second half was wrong — the
separation was already being measured, and printed on the console line above
every round the loop has ever run.

The driver's readiness ping is:

    PowerPoint.run(async (c) => { c.presentation.slides.getCount(); await c.sync(); })

A real Office.js round-trip carrying essentially nothing. It answers in
**4-7ms** — `host answered in 5ms`, every round, for months.

    round-trip, empty payload          ~5ms
    batch of 10, off-screen slide     5330ms
    batch of 10, visible slide       12283ms

**The round-trip is about a tenth of one percent of a batch.** Drawing is the
whole cost, and the ~0.1s figure in both docstrings is itself 20x too high.

So the two sentences resolve, and not evenly:

- The LIVE one is right, and understates its own case. Round-trips do not
  merely disappear next to the drawing; they are invisible.
- **The OFF-SCREEN one is wrong.** "The extra round-trips (~0.1s each) dominate
  a 60-slide run's wall clock" — at 5ms, going from 4 syncs to 1 saves about
  15ms on a chart that takes five seconds to draw. Three hundredths of one
  percent. It is the stated justification for `SHAPES_PER_SYNC_OFFSCREEN = 40`,
  a value chosen to buy a saving that does not exist, sitting nearer a
  documented wedge ceiling than the live constant is.

### The optimization I was circling is dead, and that is the useful part

Raising `SHAPES_PER_SYNC` cannot speed anything up. The cost is per-shape
server-side drawing, so the same shapes cost the same whether they cross in
one sync or four. Batching is not a latency lever on this host **at all** —
and the crash history says a bigger batch's downside is not slowness but
silence. All risk, no return, now measured rather than suspected.

### Where the lever actually is

    off-screen    533ms per shape
    visible      1228ms per shape      2.3x

Drawing onto the slide the user is looking at costs **2.3x** drawing onto one
they are not, per shape, at identical batch size. That is the repaint the live
docstring already suspected — and it is the only lever these 2,913 batches
point at: draw off-screen, then reveal. A 20s chart edit becomes about 9s.

**Caveat, stated rather than buried:** the two buckets are different scenarios
(n=336 visible, n=2577 not), so the shapes inside those batches may differ in
complexity. Batch size is pinned at ten in both, so the comparison is fair on
payload count and NOT proven fair on payload content. It is a strong lead, not
a measured 2.3x for one identical chart, and the way to settle it is to draw
the same chart both ways.

### And the reason this took an hour

Three separate instruments in this file measure something and reach nobody:
`unmatchedReReads` (fixed in cc6d7b7), `prevBatchMs` (2,913 samples, never
pooled until today), and the ping's own latency (printed on every round for
months, never once compared against a batch). The number that settled a
question about the product's biggest latency cost was on screen the entire
time, one line above the round it was needed for.


## Four product claims, checked at last. None of them held.

Four claims had been made from round data and never verified. Thirteen agents
located each one's ORIGINAL wording in the session transcript, re-derived it
from the archive, and had two adversarial lenses attack the derivation. Every
verdict survived its attackers. Every claim failed.

    readers or writers, never both (8s a round)      OVERSTATED
    one deterministic 24-shape redraw per round      OVERSTATED
    the second chart is ~7x more dangerous           REFUTED
    file path beats shape path 15x / 300x            OVERSTATED

**Three of the four paraphrases I supplied were themselves wrong**, which is
its own lesson: the claims had decayed in my own retelling between being made
and being checked. The 15x/300x one was never "two multipliers for one
comparison" — it was 15x on SPEED and 300x on RELIABILITY, two metrics
deliberately. The 7x one was about a second chart ON A SLIDE, not at insert
position 2. Checking a paraphrase would have produced four confident answers
to questions nobody asked.

### The one that mattered: the 8s is real and the user never pays it

Re-derived here, not taken on report:

    settle sleeps, last 40 rounds   199  =  4.97 per round  =  7,462ms
    as a share of scenario time                                2.0%
    rescue rate   1,103 sleeps, 97 survivors                  91.2%

The arithmetic behind "8s a round" reproduces. What does not survive is the
sentence around it — "roughly eight seconds of every render is the add-in
sleeping". Attributing every sleep to its enclosing scenario:

    4.00 /round   does a rasterise poison the next draw
    0.97 /round   explode a degraded picture
    0.00 /round   insert onto a slide that already has content
    0.00 /round   insert on top of an earlier run
    0.00 /round   edit the chart the user selected
    0.00 /round   edit a chart on the visible slide
    0.00 /round   same scale across the deck        (+ 7 more at zero)

**Two of thirteen scenarios pay it, and every mainline user path pays nothing.**
A user inserting or editing a chart waits zero of that 8s. The proposed fix was
to split every draw into two contexts and delete the settle-retry family —
a large, risky rewrite, to recover a cost users do not pay, by removing a
mechanism that rescues 91% of what it catches.

The STRUCTURAL half of the same claim confirms and is stronger than stated:
pure-read deck scans are 3,833 of 3,833 complete, while the probe that writes
then reads in one context is 0-for-465 honest, and the read fails on the FIRST
attempt 394 of 394 times — not degradation over time. Contexts here really are
readers or writers. That was right; the price tag on it was not.

### And one number the check itself under-reported

The file-vs-shape speed half was pooled as 15.8x — 47,911ms against 3,025ms
across the two isolating arms. Their details:

    insert on top of an earlier run          3s   deck grew by 4   (4 charts)
    insert onto a slide that already has …  48s   2 charts drawn onto the visible slide

**Four charts against two.** Per chart that is 0.75s against 24s — about 32x,
twice the pooled figure, because the pooled figure divided two totals without
normalising for how much each arm did.

That correction does not make it a clean 32x measurement of the PATH, and the
reason matters more than the number: the fast arm grows the deck, so it draws
onto NEW slides. The slow arm draws onto the slide already on screen. **You
cannot hand the host a file to add shapes to an existing slide** — the file
path is structurally unavailable exactly where the user spends their time.
So the arms differ in path AND in placement AND in count, and the honest read
is a strong lead about the combination rather than an isolated path result.

It also lines up with the batch measurement from earlier today: visible-slide
drawing costs 2.3x off-screen drawing per shape. Placement alone explains part
of the gap; it does not explain 32x.

### What this session's product answer actually is

Every fast thing this add-in does, it does onto a slide the user is not looking
at. Every slow thing it does, it does onto one they are. That is the shape of
all of it — the 2.3x per shape, the 24s chart on the visible slide against
0.75s on a new one, and a file path that is only available where the user is
not. **Draw off-screen, then reveal** is the one direction the evidence keeps
pointing at, and nothing in 169 rounds has tested it.


## Retraction: I committed a 2.3x, and the label it rests on does not mean that

Commit c5ecd94 said: **"drawing where the user is looking costs 2.3x per shape.
That is the only latency lever here."** It is wrong, it was pushed, and this is
the correction.

The split was on `onSlideKey === "(visible)"`. That is not a statement about the
screen. `slideKeyFor` returns the sentinel when the slide's id has not been
loaded yet:

    if (opts.slideId) return opts.slideId;
    const id = loadedValue(() => getSlide().id);
    return typeof id === "string" && id ? id : "(visible)";

The comment above it even records the sentinel being wrong before — 260 shapes
keyed to it across eight different slides. I read the word and not the
function.

### Three checks, and the finding fails all three

    sentinel-keyed batches   median onSlide  0     (emptier targets)
    id-keyed batches         median onSlide 10

    FIRST batch of a draw    5802ms  n=1835
    LATER batches            5591ms  n=1083

If visibility were the cause, the sentinel batches would be drawing onto the
busy slide the user is on; they draw onto emptier ones and are still slower. If
setup were the cause, first batches would stand out; they are within 4% of
later ones. **Two candidate mechanisms, both excluded, and the effect I named
is left with nothing behind it.**

### What actually holds, and it was already written down

    shapes this run already drew here    median batch
      0                                     3886ms
      1-20                                  5490ms
      21-50                                13995ms
      51-100                               18074ms

**A run slows ITSELF down 4.7x by piling onto one slide.** Which is exactly what
`shapesDrawnOn`'s own docstring says — "about +0.44s per shape present, measured
… quadratically" — and it is exported precisely so a caller can avoid the slide
it has been filling. My pooling rediscovered a documented effect and then
attached it to the wrong cause.

Note what `onSlide` counts: shapes THIS RUN drew there, not the slide's
occupancy. A slide that arrives with content reads as 0. So even this curve is
not "drawing onto a busy slide is slow" — it is "a run that keeps choosing the
same target slows itself down". Stated in the reader, so the next person does
not make my mistake one level down.

### And "draw off-screen, then reveal" is withdrawn

That was the closing line of the previous entry and the headline of the report
above it. It rested entirely on the 2.3x. There is no measurement here that
drawing off-screen is cheaper, because nothing in this archive distinguishes
on-screen from off-screen at all. The idea may still be right; **this data has
never spoken to it.**

### The pattern, now five deep in one day

Four scope errors this morning, then the era error, then this. Every one is a
number that is real, measured, and reproducible, carrying a sentence about a
population it never covered. The number is never the problem. The clause
attached to it is.

What caught this one: going to read what `(visible)` meant before building the
next thing on top of it. What did NOT catch it: a green gate, a passing test, a
clean mutation run, and a commit message that explained the reasoning at
length. The test pinned the arithmetic perfectly and the arithmetic was never
in doubt.

### A third vacuous test, same session

The replacement test fixture had two batches per draw. The only batch that gets
timed with two is the FIRST — whose size is 10 whether `upTo` is read as a delta
or as a running total, because a running total starts at zero. So it passed
with the bug restored. It needed a THIRD batch to reach the place the two
readings disagree. Caught by mutation, like the other two today; there is no
other way I have found to catch these.


### The round-180 inflection: the deck is eliminated too

Three candidates were named for the survivors stopping at round 180 — our
source, the host, the deck. The first went to a diff. The deck goes here.

**The archive does not record which deck a round ran against.** The driver
prints it on every line (`deck 1 slide(s) · Presentation64.pptx`) and it reaches
no round file — the same shape of blind spot as the unversioned host, found
the same way: by going to look for it and finding nothing there.

But it is answerable by proxy. Slide ids are deck-specific, and the deck's
first slide is the one the sweep never removes:

    023: 256#109857222
    029: 287#62081387
    …
    111: 256#109857222
    149: 256#2587447327     <- last change in the archive

Eight distinct values across 170 rounds, and **the last change was round 149**.
The id has been the same one through 180 and every round since. The deck did
not change at the inflection.

    not our source     #696 is downstream of the trace line, by diff
    not the probes     all settled across 180; the two flips are singletons at 184 and 190
    NOT THE DECK       first-slide id unchanged since round 149
    the host           unversioned in all 170 rounds — cannot be ruled in or out
    chance             possible, and weaker than first claimed (post-hoc split)

**Two candidates left from five, and one of them is unmeasurable here.** That is
a real narrowing rather than a conclusion, and it is worth the distinction: the
honest state of this question is now "the host or nothing", not "unexplained".

No instrument added for the deck name. The proxy answered it in one query, and
the technique is written down here so the next person does not have to find it
again — which is the cheap half of an instrument without the standing cost.


### The survivors came from ONE scenario, which did not change — and the gap is marginal

Splitting the survivors by the scenario they occurred in, which is the one
dimension this question had never been cut on:

    94  does a rasterise poison the next draw
     1  the chart is actually visible
     1  an update follows a moved chart
     1  explode a degraded picture

**97 survivors, and 94 of them are one scenario.** The same scenario carries 4.00
of the 4.97 settle sleeps a round. Whatever this is, it lives almost entirely
in one code path, not across the product.

And that scenario did not change at the inflection:

    174   51s ok  all four draws landed — …
    …
    180   48s ok  all four draws landed — …
    …
    186   47s ok  all four draws landed — …

Same duration, same detail string, same verdict, every round either side.
**Identical work, different outcome.**

### Then the check that matters, applied to my own claim

Before calling a quiet stretch an inflection: was a quiet stretch ever normal?

    rounds with at least one survivor      43 of 172
    largest gap between survivor rounds     14 rounds
    current gap (179 -> 196)                17 rounds
    historical gaps of 17 or more            0

So it is the longest quiet stretch on record — and a 14-round gap has happened
before, so 17 is **not out of family**. This is "the longest gap observed", not
"impossible under the old regime", and the difference between those two
sentences is the whole finding.

**That is materially weaker than the ~1% this journal first quoted**, which came
from a split point chosen after seeing where the zeros began. Third time today
a claim of mine has had to be walked back in strength rather than in substance.
The number moved; how much it means is smaller each time I check it properly.

### Where the question now stands

    not our source     by diff
    not the probes     all settled across 180
    not the deck       first-slide id unchanged since round 149
    not the scenario   identical duration, detail and verdict either side
    the host           unversioned — cannot be ruled in or out
    chance             still live, and stronger than I first allowed

Four eliminations and two survivors, one of which is unmeasurable here. The
honest reading is that this may simply be a long tail on a sporadic event, and
the way to tell is more rounds — which the loop is producing anyway.


## The second ask fired on the real host, and it worked

This morning's entry said the change might be permanently inert: the
population it targets had gone quiet eight rounds before it shipped, and
"a round with no survivors cannot tell a working change from an idle one".

Round 194 told them apart. Verbatim from the trace, in order:

    the cold re-read fell short — asking again after the settle
        {kind: "zero-match", index: 0, drew: 7, listed: 8}
    re-reading the slide's shapes again after a settle delay
        {attempt: 1, charts: 1, waitedMs: 1500}
    the cold re-read fell short — asking again after the settle
        {kind: "zero-match", index: 0, drew: 7, listed: 8}      <- THE FIRST ASK FAILED
    re-reading the slide's shapes again after a settle delay
        {attempt: 2, charts: 1, waitedMs: 1500}                 <- #709
    grouped the chart's shapes
        {charts: 1, partial: 0, by: "ids"}                      <- RESCUED, in full

The host listed 8 shapes for a chart that drew 7 and matched none of them.
The first ask, after a 1.5s settle, got the same answer. The second ask — the
one that did not exist before #709 — got a complete match, and the chart
grouped by ids with `partial: 0`.

**Under `REREAD_ATTEMPTS = 1` that chart is a survivor**, which on this host means
an ungrouped chart, which means a chart that loses its config and stops being
re-editable. A real one, on the real host, not the fake.

It is exactly the sequence the fault-injected test predicted, which is worth
saying plainly: the test set `faults.unmatchedIdReads = 2` and asserted the
second ask recovers what the first cannot, and reverting the constant made it
fail with `expected [ 1 ] to include 2`. The host has now produced the same shape
unprompted.

### What this does and does not change

It does NOT rehabilitate the eight clean rounds. The population still dried up
before the change shipped; those rounds are still not evidence for it, and the
report still says so. Nothing about today's finding makes yesterday's
reasoning better than it was.

It DOES close the question the morning entry left open. "Either the first
always succeeds now, or the change is inert" — the answer is neither. The
first ask still fails sometimes, and when it does the second one catches it.

**n = 1.** One rescue is not a rate and this journal has been wrong about
exactly that distinction more than once today. What it is: the difference
between a change that works and a change nothing has ever exercised.

### And the report now counts the outcome, not the event

    of the 1 second ask(s): 1 rescued the chart, 0 still lost it.

Counting that an ask FIRED says nothing about whether it HELPED — the same
gap between "it happened" and "it worked" that this section has been about all
day. The reader pairs each second ask with the next decisive line after it: a
group means rescued, a survivor line means lost, and anything else is left
uncounted rather than assumed in either direction.


## Round 197 died on a busy machine, and six retries watched it happen

    NOT READY — a round now would not prove anything:
      - playwright-cli could not be run — nothing below was actually measured.
        Install it (`npm i -g @playwright/cli`), or set PLAYWRIGHT_CLI_JS …
        the call that could not be spawned: `eval () => String(window.innerWidth)`
          — spawnSync …\node.exe ETIMEDOUT
    [exited with code 1]

The tool was installed. It had driven six rounds that day. **I caused this**: the
full 124-file vitest suite was running in the same session when the driver
tried its first spawn, and `spawnSync` timed out waiting for a process slot.

### Two faults in one branch

**The message named a cause it could not know.** "Install it" is the first thing
it says, on a condition that is only ever "the spawn did not return". The
comment beneath it already conceded the point — "without them this message
points at the install every time, and the install is almost never it" — and the
fix at the time was to APPEND the errno rather than to stop guessing in the
headline. So the evidence was there, underneath a sentence contradicting it.

**And the stop was unclassifiable, therefore fatal.** The unreachable-CLI branch
returns before any code is assigned — deliberately, and documented:
"the sign-in and unreachable-CLI states never get here because they return
before the codes do". That is right for a missing binary, which no number of
retries will conjure. It is wrong for `ETIMEDOUT`, where retrying is the entire
cure. **One label over two conditions that want opposite treatment** — the same
shape as every other finding today, this time costing a round instead of a
conclusion.

### The fix

`spawnFailureIsTransient` splits them on the errno. `ETIMEDOUT / EAGAIN / EBUSY /`
`EMFILE / ENFILE / ENOMEM / SIGTERM / SIGKILL` are "could not start it just now"
and return the new recoverable code `cli-busy`. `ENOENT` stays fatal. So does
anything unrecognised, and that default is deliberate: this loop runs
unattended for hours, so a wrong `transient` spins forever on something no retry
can fix, while a wrong `fatal` costs one round and prints why. The asymmetry
picks the default.

The headline now names the condition — "could not be STARTED — the machine was
busy, not the tool missing" — and only says "install it" when the errno
actually supports it.

### Both halves are pinned, because either alone does nothing

A classifier that nothing consults is decoration, and a recoverable code with
no classifier never fires. Reverting the regex fails with `expected false to be`
`true`; removing `cli-busy` from `RECOVERABLE_STOPS` fails with "cli-busy is not
recoverable, so --retry still cannot see it".

### The operational lesson, which is mine

**Do not run the gate while a round is starting.** The driver spawns ~8
subprocesses in its first sweep and the full suite runs 124 files; on this box
they contend. Rounds and the test suite are both cheap to schedule apart and
expensive to interleave. The round is a measurement of the HOST, and a
measurement taken while the measuring machine is saturated is a measurement of
the machine.


### A prediction that held, and half of what it promised never arrived

Today's host blind spot has a four-day-old entry in the ledger, and they should
be read together.

`the-round-file-can-finally-say-which-host-it-ran-on` (#599, judged on round
096) is marked **held**, correctly: the claim was `trace-line-present` for
`environment`, the line appears, and it replicated across the pair. Nothing about
that judgement is wrong.

But the sentence that motivated it says more than the claim tested:

> HELD means the next round is the first in the project's history to record
> **which host, platform and Office version it ran on**, and the deck-style
> replay finally has a route into the archive.

The line records host and platform. The version is `0.0.0.0` — in all 172
rounds, one distinct value archive-wide, because Office Online does not version
itself to an add-in. **A third of what the prediction promised is a placeholder,
and the ledger reads as though all three arrived.**

This is not a mis-judgement. The claim kind was `trace-line-present` and the
line is present; a claim that tested the CONTENT would have needed a different
kind and a value to test against. It is the gap between what a prediction
asserts and what its prose says it will buy — and the prose is what a reader
remembers four days later.

Worth noticing because it cost time today. "Did the host change at round 180?"
was a reasonable question to put to an archive advertised as recording the
Office version, and the answer is that it cannot be asked here at all.

**The judged entry is left exactly as it is.** Predictions are the record of what
was believed and when, and editing one after the fact to reflect what was
learned later is the one thing that would make the ledger worthless. The note
belongs here instead.


## The update path finally has a clock, and it accounts for 94% of the worst scenario

`ms` landed on `updated only the shapes that changed` in 05a27fd. Round 198 is
the first to carry it:

    {"chart":"1/8","changed":18,"of":24,"ms":39355}
    {"chart":"2/8","changed":9, "of":16,"ms":12583}
    {"chart":"3/8","changed":9, "of":16,"ms":12565}
    {"chart":"4/8","changed":18,"of":24,"ms":18333}
    {"chart":"5/8","changed":18,"of":24,"ms":17992}
    {"chart":"6/8","changed":18,"of":24,"ms":18999}
    {"chart":"7/8","changed":18,"of":24,"ms":17928}
    {"chart":"8/8","changed":18,"of":24,"ms":18768}

**156.5s of the scenario's 166s — 94%.** The measurement accounts for nearly the
whole of the product's largest cost, which the batch pooler could not see at
all. Within one chart size: **~1589ms fixed + ~930ms per changed shape**.

### And chart 1/8 costs 2.1x its own siblings

39355ms against ~18000ms, at identical `changed: 18` and `of: 24`. The fit
predicts 18407ms, so about **21 seconds of excess on the first chart of a run**.

One observation. The reader prints its own caveat — "n=1 first-chart sample(s)
— a LEAD, not a rate" — and stops at n=5. The rounds will decide it.

### One explanation is already dead, killed by the same round

The obvious story is a cold start: chart 1/8 begins 1352ms after a deck scan.
But the round contains three single-chart updates elsewhere, and they begin
**nearer** a scan than chart 1/8 did:

    (solo)  changed=1  of=24  ms=2727   1020ms after a scan
    (solo)  changed=1  of=24  ms=2290    581ms after a scan
    (solo)  changed=1  of=24  ms=2519   3243ms after a scan

The fit predicts 2519ms for one changed shape. All three land on it. **No
excess anywhere, at closer range to a scan than the expensive chart.** "The
first update after a deck scan is cold" is refuted, on data already in hand,
before a single round was spent on it.

### What is left, and the field that had to be added to tell them apart

Two explanations survive:

    position in the run     the `chart` field already carries it
    the slide itself        a chart on a fuller slide costs more to touch

The second is not idle speculation — this file measures a 4.7x curve in exactly
that variable on the draw path. And the update line **did not record which
slide**, so those two were indistinguishable no matter how long the loop ran.
A hundred rounds of a two-way ambiguity is still a two-way ambiguity.

`slideId` now rides on the line. It is one field, it is free, and it is the
difference between a question that resolves and one that accumulates evidence
forever without ever deciding.


### A warning about the field I just added: it may not break the tie

Round 199 replicated the first-chart cost to within 2% — 38512ms against
39355ms, siblings flat at ~18s both times — and `slideId` went onto the update
line to separate the two surviving explanations: **position in the run** versus
**the slide itself**.

**It probably cannot, and it is better to say so now than to discover it in
five rounds and call it a result.**

The charts come from `scanned the deck for charts`, which returns them in deck
order. Deck order is stable across rounds — the archive's first-slide id has
not changed since round 149. So the chart that is `"1/8"` is very likely the
same chart, on the same slide, every round. **Position and slide are then
perfectly confounded**, and no number of rounds separates two variables that
always move together. `slideId` will record which slide it was; it will not say
whether that mattered.

What the field DOES buy, and why it still belongs there:

- It makes the confound VISIBLE. If `"1/8"` carries the same id every round, the
  tie is proven rather than suspected, and the report can say so instead of
  accumulating samples that cannot decide.
- If the order ever does vary — a deck change, a scan that returns differently
  — the data to exploit it is already being recorded rather than starting from
  zero on the day it happens.

**What would actually settle it is a design change, not more rounds:** update
the charts in reverse, or from a shuffled order, once. If the cost follows the
POSITION the first chart stays expensive whichever chart it is; if it follows
the SLIDE the expense moves with the chart. One round of that is worth fifty of
the current shape.

That is a change to what the self-test does, on the scenario that is 38% of a
round, so it is not a thing to slip in unattended — it is written down here as
the experiment that exists, ready for whoever wants the answer.

Recording it because the alternative is the failure this session has made six
times already: watching a number replicate, feeling the confidence rise, and
not noticing that the design could never have answered the question.


## I shipped the house defect, and round 202 printed it back at me

The previous commit put `syncsOf(context)` on the update's trace line to split the
first-chart cost two ways — more syncs, or slower syncs. Round 202:

    chart  changed/of   ms       syncs  contextSyncs
     1/8    18/24        37140       0             0
     2/8     9/16        11705       0             0
     4/8    18/24        17077       0             0
     …            all eight, all zero

**A 37-second update that issues no syncs is not a finding, it is a broken
gauge.** Zero there meant NOT MEASURED — the exact defect this journal has been
cataloguing all day, written by the person cataloguing it, and caught only
because the number was absurd on its face.

### The cause is bigger than the gauge

`tryInPlaceUpdate` syncs through `step(…, () => context.sync())` directly, and only
`boundedSync` increments the WeakMap. Counted across the file:

    27  boundedSync call sites
    74  direct context.sync() calls

**About three quarters of this file's syncs increment nothing.** And the
docstring on `syncsPerContext` said, in as many words:

> Every sync in this file goes through `boundedSync`, which is what makes one
> counter enough.

It is false, and it is false in the worst possible place: that counter exists
to test whether `same scale across the deck` degrades inside one context, and
`tryInPlaceUpdate` IS that scenario. **The decay hypothesis has never been
measurable on the path it was written about**, and the comment asserting
otherwise is what stopped anyone noticing.

### The fix, and what was deliberately not fixed

The update path counts its own syncs now — a local counter incremented beside
the calls it counts, where it cannot go blind. `contextSyncs` is GONE rather
than shipped reading 0, because a field that is always zero gets read as a
measurement of zero.

Routing the other 74 sites through `boundedSync` was NOT done. It would put a
`BATCH_TIMEOUT_MS` on every one of them — a behaviour change to the whole file,
on a path that currently has no deadline at all, unattended, at the end of a
long session. That wants its own rounds. What is fixed is the CLAIM: a counter
that admits its blind spot can be reasoned about, one that denies it cannot.

### And the fork is nearly answered anyway

`IN_PLACE_WRITES_PER_SYNC` is 6, so a chart changing 18 shapes issues 3 write
syncs plus 1 for the tags — **four, first chart or not**. Charts 1/8 and 4/8
both change 18 of 24. If the count is the same and the first costs 37s against
17s, then its syncs are **slower**, not more numerous — roughly 9.3s each
against 4.3s.

That is derived from a constant rather than measured, which is precisely the
kind of reasoning this session keeps having to retract. Round 203 will carry
the counted number, and the derivation is written down here so it can be
checked against one rather than quietly replaced by it.


## The fork is answered, and my reason for widening the fix was wrong

### Round 203: same syncs, slower syncs

    chart  changed/of   ms      syncs  ms/sync
     1/8    18/24        43837      4    10959
     4/8    18/24        20548      4     5137
     5/8    18/24        19920      4     4980
     8/8    18/24        20967      4     5242

**The first chart of a run issues exactly the same number of syncs as every
other chart, and each one costs about 2.1x as much.** So the ~20s excess is
host-side, not our batching. Six rounds, four mechanisms eliminated: not the
slide, not a post-scan cold start, not a per-pane warmup, and now not our
sync count.

The derivation held exactly. `IN_PLACE_WRITES_PER_SYNC` is 6, so `changed: 18`
predicts 3 write syncs plus 1 for tags and `changed: 9` predicts 3 in total. Both
measured on the nose, and both were written down BEFORE the round that
measured them.

**What the number still cannot say** is whether one sync is slow or four are:
10959 is a mean over four, and one sync of ~28.5s with three normal ones gives
the same mean as four uniformly slow ones. Those want opposite fixes — a cheap
warm-up sync absorbs the first, and buys nothing against the second. `syncMs`
now records each sync in order, so round 204 decides it. Proposing a warm-up
from the mean would be this journal's own recurring defect.

## 20 of 30 "safe" verdicts were wrong

Nine agents classified all 65 bare `context.sync()` calls for whether a deadline
could be added, then an adversarial pass attacked ONLY the SAFE ones — a wrong
SAFE breaks production, a wrong UNSAFE costs an improvement.

    30  proposed SAFE
    20  overturned
    10  surviving
    16  NEEDS-BIGGER-BUDGET
    19  UNSAFE

Three objections I had not thought of, each disqualifying on its own:

1. **`withTimeout` does not cancel the promise** — deliberately, so it can "keep
   listening after giving up". A fired deadline therefore leaves an in-flight
   `context.sync()` on a context that then gets a second batch queued and synced
   on top of it. That is a new concurrency state, on the exact object whose
   degradation `syncsPerContext` exists to investigate.

2. **An inner deadline is often a BUDGET CUT.** Many of these sites sit inside
   `boundedRun(…, READBACK_TIMEOUT_MS)` at 90s. Adding a 45s inner deadline
   halves the budget on the pane's most-used read — and the archive contains
   `did not respond while reading the deck's style (90s)` four times, so those
   reads demonstrably do run past 45s. Others sit inside a 4s selection run,
   where a 45s inner deadline is unreachable dead code.

3. **`boundedSync` has side effects on a cross-round baseline.** It routes
   through `step()`, which on a rejected sync increments `hostFriction.errors`
   and `generalExceptions` — numbers stamped into every round record and
   differenced across rounds. Adding counting sites mid-investigation moves the
   baseline and reads as a host regression. That is this project's own
   no-history-is-not-a-spike rule, quoted back at me.

### And the premise for the whole exercise was false

I said, here and in a commit message: **74 of 101 syncs have no timeout, so a
hang there is unbounded.** The count is right. The clause is not.

`boundedRun` exists, with 21 call sites and three budgets — 45s batch, 4s
selection, 90s readback. **A bare sync inside a bounded run is bounded, by the
run.** What those sites actually lack is attribution, counting and per-sync
granularity: real, useful, and nothing like a wedge that never returns.

Seventh time today that a correct number carried a sentence about something it
did not describe. This one would have cost 65 edits to a hot path.

**Nothing in Class B is being converted.** Even the 10 survivors add `step()`
sites to a friction baseline that a live investigation is reading. The five
ALREADY-LABELLED sites converted earlier are unaffected by that objection —
they went through `step()` before and after, so no counting site was added.


## A push reported green, and CI failed on the thing the report was about

    [warn] crashes/2026-08-23T16-01-49-crashed-run.json
    [warn] Code style issues found in the above file.
    ##[error]Process completed with exit code 1.

The push was announced as "gate green: typecheck 0, lint 0, format clean, 3310
tests". The local check behind "format clean" was:

    npx prettier --check src test scripts docs

CI runs `prettier --check .` — the whole repo. Round 203's browser had died,
the driver archived the crash dump, `git add -A` swept it in, and a check scoped
to four directories never looked at it.

**The claim was "format clean". The check was "format clean in four
directories."** Eighth time today a claim and its evidence had different
scopes — and this one had crept into the thing that verifies everything else,
and had been repeated in every gate report for hours.

### The repo already knew this class

`rounds/*.json` is in `.prettierignore` with: "Machine-written round evidence —
reformatting a 250KB trace is churn, and every archived round failed the gate
until this line existed." A crashed run is the same artifact by another name,
written by the same driver into a sibling directory, and nobody had added it.
It is there now, with what its absence cost.

### The bigger gap underneath

Checking what CI's `test` job actually runs turned up **nine steps**, and the
hand-assembled local gate was running four — and not the same four:

    lint            local: yes
    format:check    local: NO — a scoped approximation
    typecheck       local: yes
    coverage        local: NO — `npm test` does not enforce CI's thresholds
    vite build      local: NO
    build:manifest  local: NO
    skill           local: NO
    verify-deck     local: NO
    validate-ooxml  local: NO

Six CI steps never ran locally. The format scope was the symptom that happened
to fire first.

### `npm run gate`

A gate that runs CI's own npm scripts, by name, in CI's order — not an
approximation of them. `npm run format:check` and not a scoped prettier call, so
if CI's scope changes the gate changes with it, because it is the same script.

**And it prints what it does not run.** Coverage thresholds, the build, the
manifest check, the skill build, deck verification, the e2e. Skipping those is
a choice about speed; hiding that they were skipped would be a lie about scope,
which is the thing the gate exists to stop. The list is exported and asserted
non-empty, so a future version that claims to cover everything fails a test
reading "a gate that claims to cover everything is the original defect".

Three properties are mutation-proven: dropping `format:check` fails, carrying on
past a failing step fails, and emptying the not-covered list fails.


## Not one slow sync — three. The warm-up fix is dead before it was built.

Round 204, each sync timed in order:

    chart  changed/of   ms      syncs  syncMs
     1/8    18/24        36189      4   [11801, 11421, 12274, 690]
     2/8     9/16        12122      3   [ 6156,  5238,        726]
     3/8     9/16        11615      3   [ 6076,  4872,        664]
     4/8    18/24        16515      4   [ 5412,  5130,  5318, 654]
     5/8    18/24        16855      4   [ 5594,  5176,  5390, 694]
     8/8    18/24        16643      4   [ 5283,  5418,  5340, 601]

**All three of the first chart's write syncs are slow, uniformly, at about
2.2x.** Not the first one with three normal ones behind it — every one of
them. And the tag sync, the fourth, is ~650ms on every chart including the
first: identical whether the chart is dear or cheap.

### What that kills

The obvious product fix was a cheap warm-up sync before the run: if the first
sync of a run pays a one-time host initialisation, a 5ms round-trip could pay
it instead and save ~20s on every multi-chart update. **That fix cannot work.**
There is no single expensive sync to absorb — the expense is a STATE that lasts
the whole of the first chart and is gone by the fourth.

The previous entry recorded both worlds before the measurement and said they
wanted opposite fixes. This is why that was worth writing down: the mean of
10959ms was equally consistent with one 28.5s sync and with three 11.8s ones,
and the fix I would have reached for is the one that is now excluded.

### What it leaves

The tag sync being flat across all charts is the sharpest clue. It is the same
call, on the same context, in the same run, and it costs the same everywhere —
so whatever is expensive is specific to the SHAPE WRITES, not to the context,
the connection, or the run as a whole.

Charts 2 and 3 hint at a decay rather than a cliff: their first write syncs are
6156 and 6076 against second syncs of 5238 and 4872. Slightly elevated, on
smaller charts, right after the expensive one. That is a shape worth more
rounds before anyone names it.

**A hypothesis, held as one:** the deck scan runs 1383ms before chart 1 and
returns, but the host may keep working afterwards, and chart 1's writes may be
competing with it. That predicts a pause after the scan would normalise chart
1 — and it is only a WIN if the pause is shorter than the ~19s it saves, which
is a real possibility and an unmeasured one. Scanning earlier, while the user
is idle, would be the version that costs nothing. Neither is worth building
before a round tests the first.


## A fresh clone cannot run a round, and every message about it named the wrong cause

A session opened on a second machine, cloning from scratch for the first time
since the loop began. Three separate stops before a round could start, and **not
one of them said what was actually wrong.**

### 1. The gate failed on all 248 files

`format:check` rejected the entire repo on a tree that was byte-clean against
HEAD. The repo has no `.gitattributes`; git's global `core.autocrlf=true` on a
Windows machine checks out CRLF; prettier's `endOfLine` defaults to `lf`.

The obvious repair is `prettier --write .`, which rewrites all 248 files and
buries whatever was actually being worked on in the diff. The correct repair is
to fix the CHECKOUT — `core.autocrlf=false`, re-checkout — and then to make the
next clone self-correcting, which is what `* text=auto eol=lf` now does.

### 2. A test that could only pass where the driver could not run

    FAIL  finds the CLI entry beside node, or says it cannot
    expected 'C:\Users\...\playwright-cli.js' to be null

`cliEntry` returns `process.env.PLAYWRIGHT_CLI_JS` on its first line, before the
injected `exists` predicate is consulted. The variable is set on any machine
that has ever run a round. So `cliEntry(execPath, () => false)` never reached
the branch the assertion names, and the test really asserted *"nobody exported
that variable"* — **green on CI, red on the only machine that runs rounds.**

The test thirty lines below it carries a comment warning about exactly this
hazard for the injected `entry`, ending "a test about spawn results must not
depend on whether a tool is installed." The lesson was learned once and not
carried up the file.

### 3. The one that cost the round: a missing directory reported as a missing tool

    NOT READY
      - playwright-cli could not be run. Install it (`npm i -g @playwright/cli`)
        the call that could not be spawned: `find --regex /[0-9a-f]{7} ·/`
        — spawnSync C:\Program Files\nodejs\node.exe ENOENT

The tool was installed and answered `0.1.18`. Node was on PATH. **`.pw-session/`
is gitignored, so a fresh clone has no session directory** — and Node reports a
missing `cwd` as ENOENT *against the executable*:

    missing cwd -> spawnSync <node.exe> ENOENT
    valid cwd   -> error: undefined | out: 0.1.18

One errno, two causes, and the message asserted the one that was false. Only
`scripts/pw.sh` created that directory, with `mkdir -p`; the driver assumed
somebody had sourced it, which is true of every machine that has run a round
before and of no machine setting one up.

**This is the fourth appearance of the same defect in this file.** The code
directly above the message already said it: *"'playwright-cli could not be run'
sent two rounds' worth of debugging at an install that was fine, because the
message named the tool and the tool was never the problem."* The fix made then
— surface the subcommand and the errno — is what made this diagnosable in one
step, and it earned its keep. **The remedy text was left naming the install
anyway.** A comment that admits "the install has never once been the problem"
sat directly above a sentence beginning "Install it".

### What changed

`ensureSessionDir()`, called from `cli()` **before** `sessionDir()`. The
ordering carries a second fix: `realpathSync.native` throws on a path that does
not exist, and the fallback returns the un-normalised string — the 8.3
short-name bug `sessionDir` exists to prevent, reached silently whenever the
directory was absent.

Spawn failures now record `entryMissing` and `cwdMissing` as CHECKED facts, and
`spawnFailureRemedy()` names the input that was missing. Both paths are
checkable, so neither is guessed. When both exist it says so, because naming
either would be the same guess in a new coat.

Seven mutants, each killed by the test that names it — including the one the
original `cliEntry` test could not have caught.

### What this session did NOT do

**No round ran. The archive is unchanged at 182.** Threads 1, 2 and 3 —
`bindings.add`, the first chart's 2.2x write syncs, the 27 untagged charts —
have no new data. Everything above is setup infrastructure, and none of it is
product.

One thing is left open and unexplained: after recovery opened the browser, the
deck loaded and sat on screen, signed in, ribbon fully rendered — and the driver
spent nine minutes on one attempt without reaching the pane. Attempt 1's
`browser-gone` was CORRECT, because nothing was open. What attempt 2 was waiting
on is not known, and it is the first thing the next round will hit.

**Answered by the next round, and it was ours:** `the window was 1280px and hides
ribbon commands — widened to 3088px`. The cramped ribbon, which `ensureRibbonRoom`
clears by itself once it is given a browser it opened. Nothing was wrong with the
host.


## The pair says deterministic — and corrects the clue the last entry called sharpest

Rounds 207 and 208, one build, second leg `--fresh`.

    chart  changed/of   207 ms   208 ms      207 syncMs              208 syncMs
     1/8    18/24        36252    36650   [11565,11189,12511,982] [11642,11832,12472,702]
     2/8     9/16        11494    11621   [ 5900, 4819,       773] [ 5938, 5085,      598]
     3/8     9/16        11560    11963   [ 5821, 4935,       804] [ 5902, 5255,      805]
     4/8    18/24        16785    16572   [ 5396, 5449, 5350, 588] [ 5264, 5201, 5535,570]
     8/8    18/24        17292    16924   [ 5869, 5230, 5469, 722] [ 5482, 5134, 5492,816]

The first chart differs by **1.1%** across the pair. Tags 11 of 11 written in
both. The settle retry: 6 first asks and 1 second ask that rescued its chart, in
both. The binding probe answered `commit-threw` 3 of 3 in both. The gap from the
deck scan to chart 1: 1237ms and 1231ms.

This is rounds 066/067 again — deterministic, not mood — and it is what makes
everything below sayable at all.

### The first chart is a rate now, not a lead

Every 18-of-24 update in the archive, by position:

    first chart of the run    n=11   median 37140   range 36189-43837
    later, same slide         n=10   median 18233   range 17395-18999
    later, other slide        n=45   median 17033   range 16273-20967

**The distributions do not overlap.** The slowest later chart ever recorded is
15 seconds faster than the fastest first chart ever recorded. The previous entry
called this "a shape worth more rounds before anyone names it". It is named.

**And it is not the slide**, which was previously unscorable because every 9-of-16
sample shared the first chart's slide. At 18 of 24 the comparison is clean:
18233 against 17033, with ranges that overlap almost entirely. The slide is worth
about a second and possibly nothing; being first is worth twenty.

**Nor is the host cold.** An ordinary 1-of-24 update on chart 1's OWN slide, 87
seconds earlier in round 207, cost 2330ms against a fit predicting ~2491 — dead
on it. The context was warm, the slide was warm, and chart 1 still paid 36s.
Whatever this is, it is a property of the RUN.

### CORRECTION — the tag sync is not flat, and that sentence came from one round

The previous entry says:

> the tag sync, the fourth, is ~650ms on every chart including the first:
> identical whether the chart is dear or cheap.

and builds on it: *"The tag sync being flat across all charts is the sharpest
clue."* **It was written from round 204 alone.** The five first-chart tag syncs
now on record:

    204: 690    205: 904    206: 841    207: 982    208: 702

Median 841 against a later-chart median of 644 — **1.31x, not 1.00x.** Only round
204's value is ~650, and it was read as the population.

What survives is the inference, and the pair strengthens it: the three write
syncs are 2.11x, 2.23x and 2.28x with complete separation, and the tag sync is
plainly not that. The expense really is concentrated in the shape writes.

What does not survive is the word "identical", and the number. Three of the five
first-chart tag syncs sit INSIDE the later-chart range of 503-846, and two do
not, at n=5 — which is not a measured elevation and not a measured absence
either. The honest form: **the tag sync does not share the writes' 2.2x, and
whether it is mildly elevated is unsettled.**

Ninth time in this file that a correct number carried a sentence about a
population it did not describe. This one was load-bearing: it was the evidence
for "nothing run-wide is elevated", and a 1.31x tag sync is a run-wide thing.

### What the pair still cannot separate

Every run's chart 1 is preceded by a deck scan, and no later chart is —
1237ms and 1231ms ahead of it in the two legs. The confound is total, and no
amount of mining separates a variable that never varies. The instrument is the
experiment the last entry already framed: a settling pause between the scan and
chart 1, traced beside chart 1's `syncMs`. With complete separation at n=11, one
round decides it — chart 1 normalises or it does not.

### Two instruments whose populations died

`a chart's tag could not even be queued` last fired in round **065**, 143 rounds
ago. `settle pass: could not repair any config tag` last fired in round **078**.
The trace that says which half of the queue key was missing was built to be read
"once rounds accumulate"; the rounds accumulated and the fault stopped happening
first. `tagging failed` is the live counter — 374 events, last fired round 206,
and zero in both legs of this pair.

### And a narrowing on bindings, from reading the code that writes the field

`bindTagTarget` is already wrapped in its own try/catch and returns `false`; the
caller clears the binding id and the `target.tags.add` on the next line still
runs. **A synchronous refusal cannot cost the chart its tag today.** The only
surviving mechanism is a queued `bindings.add` that fails at `context.sync()`,
taking the tag writes batched into the same sync with it — which means isolation
would have to be a separate SYNC, not a separate guard. Thread 1's proposal was
right, and now for a reason rather than by analogy.


## The blind-gauge detector was itself reporting from the wrong population

Reading round 207 printed this, and I repeated it in a report as a finding:

    emptyReReads has NEVER been non-zero in any scenario in any round — it
    measures nothing.

**It is non-zero in 72 of 184 archived rounds, across 136 events.**

The sentence is derived from whatever logs `poolScenarioFriction` is handed. Over
the archive it is nearly true and harmless — `npm run rounds` does not print it
at all, because pooled over 183 rounds the counter is not dead. Run on ONE round,
which is how every fresh round is read, **it asserted a fact about 184 rounds
from n=1.**

The counter is gated behind the settle retry running out. Rounds 207 and 208 each
took 6 first asks and 1 second ask that rescued its chart, so neither reached the
counting site. That is a TRUE ZERO — precisely the reading this file's own
blind-gauge warning exists to separate from a broken gauge, printed as though the
separation had been done.

Tenth time a correct number has carried a sentence about a population it did not
describe, and this one was inside the instrument that reports that class. It now
says what it measured over: below two rounds it reports silence and names both
readings rather than picking one.

**And the archive already knew.** `emptyReReads` last fired recently enough to be
live, while `a chart's tag could not even be queued` (round 065) and `settle
pass: could not repair any config tag` (round 078) are the genuinely dead ones —
three counters, one dead-looking output, and the difference is only visible by
reading the code that writes each.

## The pause experiment is built, and off

`scanSettleMs()` in `src/taskpane/selftest.ts`, between the deck scan and the
first chart of the rescale run — the only gap where a pause separates them.
`localStorage.setItem("powerchart-scan-settle-ms", "3000")` turns it on; nothing
else does. **The value is traced every run including zero**, so a round always
says which arm it ran in; an untraced default is how an experiment contaminates
its own control.

It does not measure whether a pause is worth shipping — it only measures whether
sitting still moves chart 1 at all. With the distributions not overlapping at
n=11, one round in each arm decides it. If chart 1 stays at ~37s the deck scan is
excluded and the search moves to what else is true of a run's first chart and of
nothing else in it.


## The pause changed nothing. The deck scan is excluded.

Rounds 211 and 212, one build, both `--fresh`, differing only in the pause:

    arm                 chart 1/8      chart 4-8        round    redraws
    settleMs = 0          35727ms   16229-16427ms       612s          3
    settleMs = 10000      38552ms   16257-17413ms       627s          3

**Ten seconds of sitting still did not move chart 1.** If scan contention were
the cause it should have fallen toward the ~16.5s the same-sized later charts
cost. It did not move toward it at all — it came out 2825ms SLOWER, inside the
ordinary first-chart spread. The later charts are indistinguishable across the
arms.

That was the last hypothesis the archive could not reach by mining, and it is now
dead by measurement rather than by argument. Four explanations are excluded: not
the slide, not a cold host, not our sync count, not the deck scan. Whatever this
is, it is true of a run's FIRST chart and of nothing else in that run.

    18-of-24 updates, whole archive
      first chart  n=14  median 36971  range 35727-43837
      later chart  n=70  median 17021  range 16229-20967
      separation still complete

### ROUND 210 DOES NOT TEST THE PAUSE — do not read it as the null result

The first attempt at the treated arm is archived and it is invalid. Both arms
were run without `--fresh`, to keep them comparable. **That matched the flag and
not the condition:** the control went through a recovery, which reloads the pane,
and round 210 was the only round of six that started on a pane nobody had
reloaded.

    round  driverRun                                          length
    207    attempts 3, recovered browser-gone, pane-closed      616s
    208    attempts 2, recovered pane-closed, fresh              627s
    209    attempts 2, recovered host-silent+pane-stale          618s
    210    attempts 1, recovered []                             1657s

It ran 2.7x slower overall and was already 3.5x behind before the rescale
scenario started — 996642ms against 282740ms — so the difference predates the
pause entirely. It then redrew all eight charts rather than updating them (12
`no parts list and no readable group members` against 2), which means it carries
no chart-1 update time to compare at all.

This file already had the rule: *"The pane's age when a round STARTS is the best
predictor of what that round will report"* — 0.43 post-retry fresh against 4.57
reused. It was read earlier the same session and then not applied. **A control
arm is a STATE, not a flag**, and the cheapest way to hold that state is to force
it on both arms rather than to reason about which one the driver happened to give
it to.

Kept rather than discarded: it is the cleanest demonstration of the pane-age
effect in the archive, and a 2.7x whole-round slowdown from pane age alone is
worth having on record.

### Where this leaves the first chart

No fix, and no instrument left to build for the scan. What is left is the shape
of the thing: three write syncs at 2.11x, 2.23x and 2.28x, a tag sync that does
not share it, four syncs either way, and a cost that appears on the first chart
of a run whether or not the slide, the context, the host or the scan is warm.

The next question worth a round is whether "first of a run" survives being made
NOT first — an eighth chart updated alone, in its own run, on a warm deck. If it
pays the 37s, the run is the unit and the cost is a set-up the first chart
happens to absorb. If it pays 17s, the cost belongs to whatever the run does
before its first chart that it does not do again — and the scan is already
excluded from that list.


## It was never the first chart. It was the slide it was standing on.

Rounds 213, 214 and 215, the lone-chart arm:

    round   first (slide 257)   alone     alone's slide   later (median)
    213           35894ms       35033ms   257  SAME            15328ms
    214           40328ms       39570ms   257  SAME            19564ms
    215           36350ms       15235ms   262  different       16542ms

**The arm that answered "the run is the unit" was picking `charts[0]` — the same
chart the deck-wide rescale updates FIRST, and therefore the same slide.** It
reproduced the confound it was built to break, and I reported its answer here
before checking which slide it had landed on.

Moved off that slide, the lone chart costs **15235ms** — the cheapest 18-of-24
update in the round, below every later chart in it. Being first in a run does not
cost anything.

### What the slide was

`onSlide` readings from round 215's own batch lines: slide 257 carried **42**
shapes drawn by the run against 20-21 on every other slide. Slide 257 is the
VISIBLE slide, which `insert onto a slide that already has content` and `edit a
chart on the visible slide` have both already drawn onto by the time the rescale
starts. **In this harness the first chart of a deck-wide rescale is always the
chart on the busiest slide, by construction.** Position and load were perfectly
confounded for 14 samples across 10 rounds, and the label on the more obvious of
the two stuck.

This file already measured the effect on the draw path — *"A run slows ITSELF
down 4.6x by piling onto one slide"*, 5486ms at 1-20 shapes against 14044ms at
21-50. The in-place update pays the same tax and nobody had joined the two.

**Visibility is excluded for free.** The arm calls `showSlide` on its own target,
so round 215's cheap chart was the visible one. Visible and clear is cheap;
visible and loaded is dear.

### Six exclusions, and the answer was in the archive the whole time

Not the slide *identity*, not a cold host, not our sync count, not the deck scan,
not queue depth, not a fixed timeout — and the ~20.5s excess that looked so
suspiciously like a 20s constant is just what 22 extra shapes on a slide cost
this host. The three write syncs each carrying ~+6.9s is the same story told per
sync: every write into that slide is dearer, so all three inflate together, which
is exactly what a timeout would NOT do.

### The instrument, and a mistake inside the fix for the mistake

`priorDrawsOnSlide` joins each update row to what the run had already drawn on
that slide, off `batch issued` lines that every round already carries — no new
host call, so no counting site added to a timed path and no baseline moved.

The first version called it `slideOccupancy` and printed "0 shapes" for slides
the run had drawn ten shapes onto. `onSlide` counts what THIS RUN drew and reads
0 on a slide's first batch — and the warning saying exactly that is a comment in
this repo, which I had read earlier in the same session and used as occupancy
anyway. Renamed to what it measures. **Eleventh time, and the second inside a
tool built to catch the tenth.**

### What is left of thread 2

Nothing to fix. A chart on a slide holding twice as many shapes costs about twice
as much to update, on a host that has always behaved that way, and the open
question was an artifact of a self-test deck that piles three charts onto its
visible slide. The thread closes as **explained, not fixed** — and what closes it
is a control arm moved by one slide.

### CORRECTION, same day: it is an interaction, and "load" alone is as wrong as "position"

The paragraph above replaced *"position is the cause"* with *"the slide is the
cause"*, and `triage.mjs` was edited to say so. Both are wrong, and the evidence
against the second was sitting in a comment beside the code being edited.

Round 200 put charts 1, 2 and 3 **all on crowded slide 257**. Chart 1 came in at
**+19631** above fit; charts 2 and 3 at **+2768 and +3207**. Same slide, same
load, later position — and the cost collapses.

                     crowded slide        uncrowded slide
    first of run     +19631  (dear)       ~0   (round 215: 15235ms)
    later in run     ~+3000  (cheap)      ~0

**Only first-position-on-a-loaded-slide is expensive.** Neither factor alone
reproduces it: a later chart on the same crowded slide is cheap, and a first
chart on a clear slide is cheap. It is an interaction.

Three single-cause claims have now been printed as conclusions here — position,
then load — and the trap each time was the same: a control that varied ONE factor
while the other sat at a value that happened to matter. Round 215 varied the
slide and held position at 1; round 200 varied position and held the slide
crowded. Each looked decisive alone. Only the two together say anything.

### AND THE INTERACTION CLAIM DIED TOO, an hour later, on the proxy underneath it

Splitting every 18-of-24 update by position and load produced this:

    first / clear    n=  3   median 48330   ← should be CHEAP under the interaction
    first / loaded   n= 27   median 38552
    later / clear    n=111   median 18628

A first chart on a "clear" slide came out dearer than one on a loaded slide,
which the interaction reading forbids. Then the three rows in that bucket turned
out to be **slide 257 — the crowded one — reading load 16, 26 and 10.**

`priorDrawsOnSlide` counts what THIS RUN's batches happened to record. The same
slide reads 42 in one round and 10 in another, and **18 of 50 first-position
charts have no reading at all.** It is not an occupancy measure. It never was:
this file renamed it from `slideOccupancy` for exactly that reason, and the
comment explaining the rename is four hundred lines up.

**Then I used it as occupancy. Twice.** Once to conclude load was the cause, once
to conclude it was an interaction. Both rest on a number that reports one slide
as 10 and as 42.

### What is actually known

- The lone arm cost **15235ms on slide 262** and **35033/39570ms on slide 257**.
  Direct observations, 2.3x apart, no proxy involved.
- The two slides differ. **Why** they differ — occupancy, or something else about
  257 — is NOT established, because nothing in a round measures a slide's
  occupancy.
- Round 200's charts 1/2/3 on one slide, +19631 against +2768/+3207, remains a
  real position effect with the slide held constant.

Both facts stand. The bridge between them does not, and the missing piece is
**what a slide actually holds when an update starts.**

### That measurement DOES exist — and its coverage is the problem, not its absence

Written above as "a measurement nobody has". Wrong again, in the same hour:
`countSlideShapesOnce` reads a real per-slide shape count from the host, and the
orphan check writes it into every round as `before` on `shapes left on the slide
after an in-place update`. Its values across the archive run 1, 3, 5, 12, 24, 47
— a genuine reading, not a proxy.

Joining it to the 18-of-24 updates:

    18-of-24 updates: 200 · joined to a host occupancy reading: 53
      first / 1-2    n=  1   median 37527
      later / 1-2    n= 52   median 18810

**Every joined reading lands in the 1-2 bucket.** The orphan check runs on slides
holding a single grouped chart, which is exactly where occupancy is lowest and
least interesting. The crowded slide is never read at the moment its first chart
updates, so the loaded arm has no samples at all and the first-position arm has
one.

So the honest specification, finally precise enough to build:

> An occupancy reading taken on the FIRST CHART'S SLIDE, immediately before its
> update. One `getCount()` in a sync that already exists. The reader already
> exists too — it is `countSlideShapesOnce`, and it simply is not called there.

That is a small change to a hot path, which is why it is written down rather than
made at hour five of a frozen-HEAD block. But it is no longer "we cannot measure
this" — it is "we measure it in the one place it does not matter".

Note also that groups count as ONE shape to the host, so even this reading
under-represents what a slide holds. Whether the cost tracks top-level shapes or
underlying ones is a second open question, and the same instrument would answer
it.

Fifteenth instance today, and the only one that is purely self-inflicted: the
flaw was found, the function was renamed to record the flaw, and the flaw was
then relied on anyway. Renaming a thing is not the same as remembering it.


## Rounds get slower the longer you run them, and no round records when it ran

Ten rounds, one build, every one `--fresh`, run back to back for four hours to
measure this project's noise floor properly for the first time. It measured
something else instead.

    round  finished   min-in    len   laterMed   scenarios
    216    07:57         0     699s     19321      14/14
    217    08:23        26     684s     18810      14/14
    218    08:40        44     747s     22689      14/14
    219    09:02        65     935s     24423      14/14
    220    09:31        94     889s     31089      13/14
    221    09:52       116     957s     40469      12/14
    222    10:10       133     724s      (none)    12/14
    223    10:43       166    1085s     36712      12/14
    224    11:03       186     847s     29023      14/14
    ---- 15 minutes idle ----
    225    11:40       224    1011s     40067      11/14

`laterMed` is the median in-place update of an 18-of-24 chart on its own slide —
the most repeated measurement this harness makes. **It roughly doubles over two
hours**, and past about ninety minutes scenarios start being SKIPPED because the
host stops answering mid-scenario. Nothing failed; the host simply stopped
replying inside the budget.

### It is not this machine

`--fresh` closes the browser and opens a new one before every round, so the
client is destroyed and rebuilt ten times across that table and the degradation
walks straight through it. The machine had **8.6 GB of 15.7 GB free** at the
bottom of the run with Chrome at 1.9 GB. There is no local resource curve here.

### A 15-minute rest did not clear it

That was the experiment worth the last half hour, staked in advance: recovery
toward ~19000 would mean transient load that decays, and the rule would be to
space rounds out. **It came back 40067 — worse than the 29023 immediately before
the rest.**

Held honestly, this is ONE rested round in a regime that is bouncing between
29023 and 40469, so it does not *prove* the effect is cumulative. What it does is
remove the cheap fix: fifteen idle minutes is not a remedy, and anyone reaching
for "just wait a bit" should know it was tried.

What is left un-separated: document growth, account-level throttling, and plain
time-of-day server load all remain live, and this block cannot tell them apart.

### The part that reaches back through the whole archive

**No round records when it ran.** The round file carries the BUILD stamp
(`77fc64f · 2026-08-24 05:33Z`) and nothing about the round's own wall clock. So
session position is unrecoverable from 190 archived rounds, and every cross-round
comparison ever made here is confounded by a variable that moves the headline
measurement by 2x.

The noise-floor note already had its finger on this without knowing it:

> NOT symmetric: see PAIR POSITION below — the second run is usually the worse
> one, so a floor measured this way includes an effect as well as noise.

That asymmetry is real, it was never explained, and this is the explanation. A
pair run back to back does not have two equivalent legs — **the second leg is
systematically disadvantaged**, and how much depends on how deep into a session
it lands.

### The instrument, and it is the deliverable of this block

Stamp every round with its own start time and its position in the session, so the
archive can be split on both. Until that exists, "this build is slower" and "this
round ran later in the day" are the same sentence, and 190 rounds cannot say
which one they mean.

Not built here: HEAD was frozen for the whole block, deliberately, because a
commit moves HEAD ahead of the deployed site and the driver then refuses every
subsequent round as `site-behind`. It is the first thing to build next.

### What the block did NOT produce

**A noise floor.** That was the stated deliverable and it is not in these ten
rounds: within-session drift is larger than whatever run-to-run noise exists
underneath it, so pooling the ten gives a spread that is mostly the trend. A real
floor needs rounds at a CONSTANT session position — the first round of ten
separate sessions, not ten rounds of one.

### CORRECTION, from mining the block after reporting it

The entry above says *"Nothing failed; the host simply stopped replying inside
the budget."* **That is wrong. The host CRASHED three times during the block**,
and the crash dumps were sitting in `crashes/` unread while the summary above was
being written off the round files alone.

    2026-08-24T06:12:20Z   crash 362s into a round    round 218
    2026-08-24T07:17:15Z   crash 364s into a round    round 220/221
    2026-08-24T08:26:07Z   crash 406s into a round    round 223

Roughly hourly, and at a strikingly consistent **~6 minutes into the round**:
362s, 364s, 406s. The driver recovered all three — which is why the rounds
completed at all, and why `attempts` reads 2 or 3 on exactly those rounds — so
nothing in the round files says "crash" and the summary read the recoveries as
ordinary retries.

`crash-forensics.mjs` earns its keep here for the second time. It is also the
only artefact in this project that records a WALL CLOCK: `startedAt`, plus 400+
steps of what the host was doing. The rounds have no such field, which is the
instrument this block is asking for — and the crash dump shows the shape it
should take.

**What this changes.** "Slower under sustained use" and "crashes hourly under
sustained use" are different claims with different fixes, and the second is the
one the archive should carry. Whether the slow-down and the crashes are one
mechanism or two is not settled here: the crashes land inside the degraded
stretch, which is suggestive and not evidence.

Twelfth time in this file, and the plainest: the headline was written from one
source when a second, richer one had been produced by the same run and was
already on disk.

### CORRECTION TO THE CORRECTION — the round files DID record the crashes

The entry above says *"nothing in the round files says 'crash'"*. That is wrong,
and it was wrong in the same way as the thing it was correcting: written without
opening the field it was making a claim about.

    round  attempts  recovered
    216       2      ["not-ready:pane-closed"]
    217       3      ["not-ready:pane-closed", "crashed"]
    220       3      ["not-ready:deck-dirty+pane-closed", "crashed"]
    223       3      ["not-ready:pane-closed", "crashed"]

`driverRun.recovered` says **`crashed`**, in plain text, on exactly the three
rounds concerned. The attribution above is also wrong: the crashes belong to
rounds **217, 220 and 223**, not 218 — the first crash landed at 08:03 during
round 217's first attempt, which is why 217 reads `attempts: 3`.

Thirteenth time, and the first one that is a correction of a correction. The
lesson is not "be more careful"; it is that **the field existed and nothing ever
printed it.**

## The gap under all three mistakes: archived and never rendered

`driverRun` has been written into every round since it was added and **triage
never printed a character of it.** A round rescued from a PowerPoint crash
rendered identically to one that walked straight through. That is what made the
same block readable wrongly twice in a row by someone with all the evidence.

Now on the round header:

    build 77fc64f · 2026-08-24 05:33Z · PowerPoint · OfficeOnline · 0.0.0.0
    3 attempt(s) · fresh browser
    recovered from: not-ready:deck-dirty+pane-closed, crashed

Named rather than counted — *"recovered from 2"* invites a skim, `crashed` stops
one.

### Skipped is a third outcome and the headline had two

`SELF-TEST 12 of 14 scenarios passed` reads as two failures. Every missing
scenario in the block was a SKIP: the host stopped answering and **nothing was
checked**. A degrading host looked like a regressing product. The per-line marks
have always said `skip`; the line people quote did not:

    SELF-TEST 12 of 14 scenarios passed · 2 skipped (nothing was checked)

### And the round now records when it ran

`sessionPosition` chains through `.round-outcome.json`, because each round is its
own process and there is nowhere else that outlives one. A gap over 45 minutes
starts a new session; a receipt from the FUTURE is a moved clock and is refused
rather than trusted into a negative gap.

`driverRun` gains `startedAt`, `sessionIndex` and `sincePrevRoundMs`, and any
round at position 5 or later prints:

    DEEP IN A SESSION — the in-place update cost roughly doubles by round 6-8 and
    scenarios start being skipped; compare against rounds at a similar position,
    not against round 1

It says nothing at all for the 190 rounds archived before the field existed. A
row of question marks on every one of those is noise, not information.

Ten mutants across the four changes, each killed by the test that names it.


## The upstream audit: one stale citation, and an answer nobody had asked for

`issue-status.mjs` flags **8 of 23 cited issues as closed/completed** and says
"re-read what cites it". Read properly, most of those citations had already been
re-read and deliberately kept — but not all.

### office-js#5022 — the citation was current, and the probe had an answer

`host-probe.ts` already carries: *"CLOSED AS COMPLETED on 2024-11-18 — checked
against the GitHub API on 2026-08-14, not read off a stale note. It is kept and
still asked, because a fix upstream is a claim about the service and this probe
is the only thing here that can check it."*

That is exactly right, and **nobody had ever pooled the probe's answers.** Across
201 rounds `picture-then-shape-read` says:

    520  unreadable
     19  yes
      9  threw
     55  no-scratch-slide

So of the 28 DECIDED answers, 19 say the symptom did not reproduce, and `yes`
appears as recently as rounds 222 and 225. The probe was built to answer whether
the upstream fix reached this host; it has been answering for a year and the
answer was never read. It is not clean — 9 throws — but it is not the 2024 bug
either.

### office-js#3014 — genuinely stale, and the repo contradicted itself

`reconcile.ts` said a grouped slide's children are "unreachable on every
PowerPoint (office-js#3014)". `powerpoint.ts` says of the same issue: *"that note
is from 2022 and is now out of date"*. **Two statements about one issue in one
repo, one of them treating a bug closed in March 2025 as a current fact** — and
`docs/ROUNDS.md` already names this exact citation as its worked example of
reasoning from something Microsoft fixed eighteen months ago.

The behaviour it justifies is right anyway, and the archive says so far more
strongly than the ticket ever did. Across 201 rounds:

    group-reports-its-children      548 threw,       0 answered
    group-children-via-getcount     544 unreadable,  0 answered
    addgroup-returns-usable         541 unreadable,  0 answered

**1633 observations that the fix has not reached this host.** So the claim is
kept and re-grounded in what this host does, measured, rather than in an upstream
ticket. That is the same doctrine the #5022 comment states, applied in the
direction nobody had applied it: an upstream fix is a claim about the service,
and here the archive refutes it locally.

### What the audit did NOT find

The other six citations are either already self-audited (`#5022`, `#225`,
`#1650`) or sit in docs and watch-lists where a closed issue is the point.
`powerpoint.ts`'s `#3014` block is one of the best-evidenced comments in the
repo — it refutes its own earlier claim with 76 archived occurrences. Nothing
there needed touching.

### And one thread-1 crumb, from the same pass

`does-a-failed-group-poison-the-tag` answered **`tags-gone` in rounds 209 and
221** — twice in the last twenty. A failed group taking the tag with it is not
historical; it is still happening, and it is the mechanism thread 1 is about.
39 `tags-gone` against 476 `no-refusal` across the archive.


## Thread 3, re-scoped: the trace it waits on has been dead for 161 rounds

The standing item reads *"27 charts failed tagging and got no repair attempt. A
new trace now says which half of the queue key was missing — read it once rounds
accumulate."*

The rounds accumulated. The trace did not fire.

    a chart's tag could not even be queued   last round 065   dead 161 rounds
    tagging failed                           last round 206   73 of 202 rounds

**The instrument was built for a fault that stopped happening before it shipped.**
Waiting for more rounds cannot help; nothing in 161 of them produced the line.

What is still alive is `tagging failed`, and it is SPARSE rather than gone: 73 of
202 rounds carry at least one, but only rounds 179 and 206 in the last
twenty-five, almost always exactly one chart. That is the fault worth carrying
forward, and it wants a different question from the one thread 3 asks — not
"which half of the queue key was missing" (that path is unreachable) but "what
distinguishes the one chart in a round that fails to tag from the ten that do
not".

Thread 3 as written is **closed**. Its successor is a sparse single-chart tagging
failure, and at roughly one round in twelve it needs either a long series or a
narrower trigger before it can be studied at all.

The dormancy report now prints this class automatically, so the next instrument
built for a dead population should announce itself before anyone waits on it.


## Every crash announced itself a minute in advance, in a field already archived

The three crash dumps from the 2026-08-24 block had never been opened. They carry
400+ steps each, and all three say the same thing.

    crash        chart 1/8    chart 2/8    died at
    06:12:20      64928ms      18542ms     chart 3/8
    07:17:15      60572ms      20085ms     chart 2/8
    08:26:07      65961ms      23239ms     chart 2/8
    healthy      36189-43837   ~11500      —

**All three died inside `same scale across the deck`, at chart 2 or 3.** Not
spread across the battery — the same scenario, the same few seconds into it,
three times.

And every one was already failing when it got there. The first chart of that run
cost 60-66 seconds against a healthy range of 36-44 across fourteen samples:
**no overlap.** The crash follows about a minute later, at the second or third
chart.

### What this joins together

The session drift and the crashes are not two findings. `laterMed` doubling over
a session is the same curve, and the crash is its end: the host slows, keeps
slowing, and dies in the heaviest scenario in the battery. The block reported
them separately because the round files showed the slowdown and only the
unopened dumps showed the crash.

### The early warning is already in every round

Chart 1/8's `ms` is written to the trace the moment it finishes, roughly a minute
before the host gives up. A run that sees it above ~45s — clear of every healthy
sample on record — is a run that is going to lose the rescale.

Not built into the driver here, and not because it is hard: three observations
is a lead, the threshold is drawn from fourteen healthy samples and three sick
ones, and a driver that aborts on a number this new would start throwing away
rounds that would have finished. It is written down so the next crash can be
checked against it rather than being the fourth one nobody compared.

**What would settle it:** the next crash dump, read the same way. If its chart
1/8 is also above 45s, that is four for four and the threshold is worth acting
on. If it crashes from a healthy first chart, the signal is gone and this entry
is why.

### It was settled within the hour, against me

Round 232, archived about forty minutes after the paragraph above was written:
**chart 1/8 at 53138ms, and no crash.** No dump, no skipped scenario, the round
finished normally.

So a first chart above 45s does NOT predict a crash. The threshold was drawn
from three sick samples at 60-66s and fourteen healthy ones at 36-44s, and the
gap between those clusters is exactly where the line got put — with nothing in
it. 232 landed in that gap and survived.

The available move is to redraw the line at 60s, above 232 and below the
crashes. That is fitting a threshold to four observations by moving it past the
one that refutes it, and it is not worth having. **What survives is weaker and
true: all three crashes were preceded by a badly elevated first chart, and an
elevated first chart on its own means very little.** Necessary, not sufficient,
on n=4.

Written before the refutation and left standing with it, because the prediction
being wrong is the useful part — and because a threshold quietly moved to fit is
how a number becomes doctrine without ever being tested.

### The fourth crash: pattern broken, threshold untestable

Round 239, 2026-08-25. It crashed in **`insert onto a slide that already has
content`** — not the rescale — with exactly one update recorded, and never
reached a chart 1/8 at all.

Two things follow, and neither is comfortable:

- **"All three died inside `same scale across the deck`" is now 3 of 4.** That
  sentence is three paragraphs above this one and was written as a pattern. One
  more observation and it is a majority, not a rule.
- **The threshold could not be tested.** Its settle condition was "read the next
  crash dump's chart 1/8" and there was no chart 1/8 — the host died before the
  measurement the warning depends on exists.

That is a limitation of the signal, not a neutral outcome: **an early warning
that lives in the rescale can only warn about crashes that happen after the
rescale starts.** A quarter of the crashes on record now happen before it.

Standing at three supporting cases, one false positive (round 232) and one
inapplicable. Not a rule, and the honest next step is still to read the crash
after this one rather than to adjust the number.

### The fifth crash: threshold up to 4 of 5, scenario pattern down to 3 of 5

Round 240's first attempt, 2026-08-25. It reached `chart 1/9` at **49997ms** and
died in **`explode a degraded picture`** — a THIRD scenario.

Both tallies move, in opposite directions:

- **The threshold gains a case.** 49997ms is above the 45s line and the round
  crashed. Four supporting, one false positive (round 232 at 53138ms, no crash),
  one inapplicable. Still not a rule — a line with a known false positive on
  either side of it is a hint — but it survived a real test rather than a
  convenient one.
- **The scenario pattern loses one.** "All crashes happen in the rescale" is now
  **3 of 5**: three in `same scale across the deck`, one in `insert onto a slide
  that already has content`, one in `explode a degraded picture`. Written as a
  rule at n=3, a majority at n=4, and barely that at n=5.

The pattern is dissolving as the sample grows, which is what a pattern read from
three observations usually does. The threshold is holding. Neither is settled and
both are now written where the next crash can be checked against them.


## The host is unversioned — but the probes are a fingerprint, and it has not moved

`host` reads `PowerPoint · OfficeOnline · 0.0.0.0` in **all 207 rounds**: one
distinct value, no version. So the standing caveat — *"a host-side change cannot
be ruled in OR out from this archive"* — has been true of the only field that
could have settled it.

That is a bigger hole than it looks, because this session found three fault
populations stopping at specific rounds: `no-queue` at 065, `cfg5010` at 069, and
every tag fault after 206. If the host changed under us, all three are explained
without any of our fixes working.

**The probe answers are a behavioural fingerprint, and they can be read for
exactly this.** Four probes, binned by round:

    picture-then-shape-read          unreadable  90% / 79% / 93% / 88%
    does-a-failed-group-poison…      no-refusal  94% / 80% / 94% / 87%
    group-of-existing-shape-readable no-group-id 87% / 82% / 88% / 91%
                                     (bands: 001-069 · 070-149 · 150-205 · 206-231)

**No step change at any of the boundaries.** The host answers the same questions
the same way, within ordinary variation, across the whole archive.

One probe DID invert, and it is instructive that it is not the host:
`binding-names-shape-later` goes `no-scratch-slide` 82% → 12% and `commit-threw`
14% → 83% at the 070 band. That is the harness learning to hand the probe a
scratch slide — the question started being ASKED, not answered differently. A
fingerprint has to be read for what moved it.

### What that buys

The dead populations are far more likely to be **our fixes working** than the
host moving. `cfg5010` stopping at 069 sits right behind the settled retry
shipping at 064/065, which is the explanation the archive already offers — and
this is the first check that the alternative explanation was ruled out rather
than assumed.

It also means the unversioned-host caveat can be stated more usefully: the
version string says nothing, but the archive is not blind to host change. It has
a fingerprint and the fingerprint is flat.


## The parts list: written on every chart, consumed on none, in 206 rounds

`CHART_PARTS_TAG` exists so an in-place update can find and delete the rest of an
ungrouped chart. Pooled over the whole archive:

    parts list outcome, by `where`
      1748x  gotPartsList=0   no loose chart had siblings   — by design, none needed
       102x  gotPartsList=2   the id read-back THREW        — wanted, failed
         3x  gotPartsList=1   read the ids back             — wanted, succeeded

    in-place update readings: 1005 · charts seen: 1005 · WITH a parts list: 0

So of roughly 105 occasions where a parts list was actually WANTED, about three
produced one — and **zero of 1005 charts have ever reached an update carrying
one.** Meanwhile `not updating in place` has fired **1601 times across all 206
rounds**, 1197 of them specifically for want of a parts list, each redrawing a
whole chart instead of writing only the shapes that changed.

**The read-back failure has a name.** 438 archived occurrences of
`InvalidParam passed to GetItem(id)` at `reading back an ungrouped chart's shape
ids` — the 5010 stale shape proxy, office-js#2903, closed as NOT PLANNED.

### Why this is recorded rather than acted on

The repo half-knew: *"the parts list has never once been produced on this host —
`withParts` is 0 across 872 charts"* sits inside a comment about a different
problem, as an aside. What was missing is the consumption side and the
denominator — 105 wanted, 3 produced, 0 consumed — which is what turns an aside
into a fact about a feature.

It is the same shape as `Retire settleByBinding — 0 rescues in 180 rounds`, two
commits before this one. It is NOT the same case:

- `settleByBinding` failed for no reason anyone could name. This fails for a
  named upstream bug that Microsoft has closed as not planned.
- Retiring it would also remove the only thing that MEASURES the block. A
  mechanism that is blocked rather than useless still earns its trace.
- And the cost is one tag write per chart, which is small beside the 1197
  redraws its absence causes — the expensive half is the fallback, not the
  attempt.

So: on the list, with the numbers attached, for the owner to call. What would
change the answer is a way to re-resolve the shapes by id after the 5010 —
which is the same repair `does-a-failed-group-poison-the-tag` points at from the
other direction, and both are blocked on the same stale-proxy behaviour.

### AND THAT REPAIR ALREADY EXISTS, ten thousand lines away, at 100%

    a by-id lookup refused the whole resolve — re-reading the slides instead
      86 refusals across 86 rounds
      re-read fallback: asked 86 · recovered 86  (100%)

The in-place update path meets the SAME 5010 refusal, falls back to re-reading
the slide's shapes, and has **recovered every single time in 86 rounds.** Once a
round, silently, for months.

The parts-list read-back meets the same refusal and does something weaker: it
SALVAGES ids the drawing batch had already answered (`loadedValue(() => s.id)`)
and does not re-read. That salvage is right on its own terms — a sync failing
does not un-populate what was already answered — but it yields a whole list
almost never: 102 throws, about 2 lists.

I wrote here, an hour before finishing the paragraph, that **"the repair is
proven in this codebase and two sites do not use it."** That was wrong, and the
way it was wrong is worth more than the claim was.

### The repair cannot transfer, for exactly the reason it works

`reReadRefusedShapes` recovers by re-reading the slide's shape COLLECTION and
matching each entry against `e.it.target.shapeId` — **an id the chart already
holds, off its own tag.** The handle went stale; the identity never did. That
asymmetry is the whole mechanism: *"a by-id lookup is the one thing PowerPoint on
the web reliably refuses, and a collection read is the one thing it reliably
honours."*

The parts-list read-back fails at the opposite thing. It is trying to LEARN the
sibling ids — that is what the read is for. When it throws there is nothing to
match a collection read against, because the identities are precisely what was
lost. A collection read there returns shapes nobody can name.

**Same errno, same host bug, incompatible remedies.** The similarity is at the
level of the error text and nowhere else.

That is `tagAnchorIndex` again — *"aimed one level too low"* — reached by pattern
matching `InvalidParam passed to GetItem(id)` in two places and assuming one fix
covered both. It survived about forty minutes, and only because reading
`reReadRefusedShapes` was the next thing rather than writing the patch.

### What is actually true

- The re-read fallback is excellent WHERE IT APPLIES: 86 of 86, once a round.
- The parts-list read-back is blocked by the same host bug and has no comparable
  remedy available, because it has no id to hold on to.
- `.group` throwing twice a round is downstream of the same emptiness: the guard
  is `parts.length`, and with no parts list every chart reaches `.group`.

The repo already framed the one honest direction, on the batch itself: *"the
parts-list repair may remove this by itself: a chart that arrives carrying parts
never queues `.group` at all. Left to a round to measure rather than
restructured on that theory — the batch shape here has three shipped-broken
fixes on record behind it."*

So the question is not how to recover a failed read-back. It is whether the ids
can be obtained WITHOUT a by-id read at all — and this file has no evidence
either way, which is a better place to stop than a fix that cannot work.

What would still prove any of it: the parts list appearing on a chart that
reaches an update. `withParts` has been 0 across 1005 charts, so ONE is a signal
and there is no noise floor to argue about. That much stands.


## Thread 1 cannot be tested right now, and building it would prove nothing

The proposal is to isolate `bindings.add` in its own sync, so a refused binding
stops costing the chart's config tag. The code reading holds up: `bindTagTarget`
is already wrapped in its own try/catch and returns `false`, the caller clears
the id, and `target.tags.add` on the next line still runs — so a SYNCHRONOUS
refusal cannot cost the tag today. The only surviving mechanism is a queued
`bindings.add` failing at `context.sync()` and taking the batched tag writes with
it, which is why isolation would have to be a separate SYNC rather than a
separate guard.

**And the fault has stopped happening.**

    build     n   tags-undefined  cfg-tag-5010  group-5010  no-queue  tagging-failed
    77fc64f  10         0              0            0          0            0
    07c25ae   1         0              0            0            0          0

Last tag fault of ANY kind: build `beec9a6`, round 206 — about twenty rounds ago.
`cfg5010` last fired in round **069**, 157 rounds back. `tags-undefined` appears
in **zero of 202 rounds**. A failed config-tag route: zero of 202.

So both arms of the experiment would score zero. The change would merge, the next
pair would show no difference, and "no difference" would be read as "the
isolation is safe" when it actually means **the fault did not occur in either
arm**. That is precisely the shape this file already recorded for the settle
retry — *"THE POPULATION DRIED UP BEFORE THE CHANGE: 8 clean rounds came FIRST,
so the quiet rounds since are NOT evidence for it"* — and for `tagAnchorIndex`,
which merged, produced no measured effect across five rounds and four builds, and
was reverted.

**Not built.** Not deferred either: what it needs is stated, so the next session
does not re-derive it.

- A **trigger that reproduces tag loss on demand** — a scenario that forces the
  batch to be refused — turns this from a waiting game into a measurement. That
  is the instrument, and it is the whole deliverable for thread 1.
- Failing that, a **long series** to catch the sparse survivor: `tagging failed`
  runs at roughly one round in twelve lately, so a pair cannot see it and even
  ten rounds is a coin flip.

The counter-argument in the standing notes — that the binding's value is an
upstream side effect of being in the drawing batch, since `cfg5010` went 8/round
to 0 when it arrived — is now **unfalsifiable from the archive for the same
reason**. `cfg5010` has been 0 for 157 rounds. Whether that is the binding doing
its job or the fault having left on its own cannot be told apart by removing the
binding and watching a zero stay zero.


## The slowdown is entirely the host's — our idle time is 1ms in every round

`idleMs` is `issued - lastAnswered`: the gap between the host answering and our
code issuing the next call. It is OUR time, and it has never been read by
anything. Across 32 rounds spanning a 2.6x range of speed:

    laterMed 15328ms (fastest round)   median idleMs 1ms
    laterMed 40469ms (slowest round)   median idleMs 2ms

    fastest half: 1ms      slowest half: 1ms      max ever seen: 17ms

**Flat.** A round that takes two and a half times as long to update the same
chart spends the same one millisecond between calls. Every additional second is
inside a sync, waiting for the host.

That closes a whole class of explanation for the drift, the crashes and the
first-chart cost alike: it is not our batching, not our scheduling, not our
sleep, not garbage collection between calls. Combined with what was already
known — 8.6 GB free at the bottom of the worst block, the browser destroyed and
rebuilt every round by `--fresh`, four syncs issued either way — there is nothing
left on this side of the wire.

Worth stating because the opposite was never tested. "The host is slow" has been
the assumption behind every timing entry in this file, and an assumption held
that long deserves one measurement. It now has one, from a field that was being
written 119 times a round and read zero times.

### How it degrades per operation: three claims, all of them noise

This section first said the slowdown "scales with the work" — light calls barely
moving while heavy shape-writing syncs took the whole hit. **That was an
arithmetic error**: the light operations were measured as fast-half against
slow-half medians, and the update was quoted as min-to-max across every round.
+11% and +164% were never comparable numbers.

Everything on one basis — median over the fastest half of rounds against the
slowest half:

    draw batch                        5474ms →  9516ms   +74%
    counting the deck's slides          13ms →    18ms   +38%
    UPDATE, 18 of 24                 16855ms → 22689ms   +35%
    the visibility CONTROL render      345ms →   454ms   +32%
    deleting the chart being replaced   41ms →    49ms   +20%
    writing the chart's origin tag     582ms →   655ms   +13%
    the rasterise-poisons arm          653ms →   728ms   +11%
    selecting a shape                  840ms →   835ms    -1%

### …and then the table itself failed, on its own error bars

Printed with sample sizes and interquartile spread, every per-operation row
collapses:

    operation                        nFast nSlow   fast  slow  change   IQR(fast)
    counting the deck's slides          48    51     13    16    +23%       11
    selecting a shape                   16    17    808   835     +3%      182
    writing the chart's origin tag      16    16    582   655    +13%      105
    deleting the chart being replaced   24    26     41    49    +20%       50
    the rasterise-poisons arm           32    34    653   728    +11%      105

**Every difference is smaller than the spread inside its own bucket.** "Selecting
a shape is untouched" — the sharpest thing in the previous paragraph — is a 27ms
gap against a 182ms IQR at n=16. It says nothing.

The tell was there before the IQRs: adding round 233 moved the slide count from
+38% to +23%. A figure that swings fifteen points on one round is not measuring
an operation.

**Retracted.** The per-operation comparison cannot support any conclusion at
these sample sizes, in either direction. What is left:

- The UPDATE and DRAW measurements have hundreds of samples each and both
  degrade in slow rounds. That much holds.
- Whether the degradation is flat, proportional, or selective is **unknown**, and
  three successive claims about it in this section were all built on noise.
- The per-sync asymmetry — three write syncs at 2.2x against a tag sync at
  1.19x — still stands, because it is a within-call comparison over n=26/130 and
  does not come from this table.

Two corrections in one section, forty minutes apart. The first was mixed
statistics; the second was reading a difference without ever asking how wide the
distribution underneath it was. The IQR column took one line to add and would
have prevented both.

### The same test, turned on the finding this file has been leaning on

The per-sync 2.2x has been quoted here all day. It had never been shown its own
quartiles either, so:

    sync      first n=27  q1     med     q3      later n=135  q1     med    q3    overlap?
    write 1        11585  12763  16102               5489   6030   8398      no
    write 2        11718  12988  14397               5198   5727   7456      no
    write 3        12446  13858  16491               5471   6106   7508      no
    tag              702    794    899                588    666    769     YES

**The three write syncs do not overlap at the quartiles.** The slowest quarter of
later charts is still faster than the fastest quarter of first charts, on all
three, at n=27 against n=135. That is a real separation and it survives the test
that just killed the per-operation table.

**The tag sync overlaps.** Its 1.19x is not established — which is what this file
already said from a smaller sample ("whether it is mildly elevated is
unsettled"), now with the reason attached rather than the hedge.

So the write/tag asymmetry stands: the writes separate, the tag does not, in the
same call and the same context. It remains the sharpest structural fact about
this cost, and it is now the only one in this section with evidence behind it.

Worth doing deliberately: after two claims died on their error bars, the honest
next move is to point the same test at the claim you would rather keep.


## The host's speed is a property of the ROUND, not of the chart

Every chart in a round runs at about the same speed; different rounds run at very
different speeds. Across 32 rounds carrying four or more timed later charts:

    WITHIN a round   median spread   8%   (range 1-51%)
    BETWEEN rounds   15328-40469ms = 164%
    between RESTED rounds only              ~47%

    230   19468 19695 19237 19311 19552    within  2%
    231   17424 17928 18628 17544 16887    within 10%
    232   25707 23080 28568 25822 25845    within 24%

Round 230's five charts agree to 2% and round 232's to 24%, while their medians
differ by 33%. The round has a speed; the charts inside it inherit it.

### Why this matters more than the numbers

**Comparing charts inside one round understates the real variance by roughly six
times** (against rested rounds) **to twenty-one** (against the archive as a
whole). Any conclusion drawn from "chart A was faster than chart B in this round"
carries an error bar six times wider than it looks.

That is not hypothetical — it is how the first-chart cost was read for ten
rounds. Chart 1 against charts 4-8 IN THE SAME ROUND is exactly this comparison,
and it survived only because the effect was 2.2x, far outside even the
between-round spread. A smaller effect read the same way would have been noise
wearing a finding's clothes, and nothing in the report would have said so.

The same applies to the `alone` arm added earlier today: it is compared against
later charts in its own round, which is the RIGHT design precisely because
between-round variance is the larger term. Holding the round constant is what
makes that comparison work — and it is worth naming, because the reason it is
correct is the reason the other reading was fragile.

### And it tightens the within-round spread as a signal

Within-round spread grows with the round's slowness: 2% at 19.4s, 10% at 17.5s,
24% at 25.8s, and 51% at the archive's worst. A round that is slow is also
INCONSISTENT, which is a second, independent symptom of the same sickness and one
that needs no baseline to read — it is computed entirely inside the round.


## Rested rounds do not drift — they vary, and the floor is wider than hoped

Three samples in, every one taken at `sessionIndex 1` after a 47-minute rest, on
one build:

    round   chart 1/8   laterMed   length
     230      42987      19468      691s
     231      35491      17544      625s
     232      53138      25822      847s

**232 is degraded, and it should not be.** It had the same rest as the other two
and sits at position 1 by construction, verified rather than assumed — and it is
47% slower than 231 on `laterMed` and 50% slower on the first chart.

### What that does to the protocol

The floor protocol assumes a 45-minute rest returns the host to a common
starting state, so that samples at position 1 differ only by noise. **If they
drift anyway, the spread I am measuring contains drift as well as noise, which
is exactly the defect the protocol was written to escape** — the same defect as
pooling ten back-to-back rounds, one level up.

It also fits the earlier result nobody could explain: a 15-minute rest did not
restore performance either. Two rest lengths, neither sufficient.

### And that reading was wrong too — checked before it hardened

The obvious explanation was cumulative drift over the DAY rather than the
session. Laid out against the clock, it does not survive:

    morning, back-to-back   06:17→09:52   35894 → 40328 → 54578 → 56783 → 58301
    evening, rested         18:48→22:50   46556 → 38782 → 42987 → 35491 → 53138

**The morning block climbs monotonically. The evening rested rounds do not climb
at all.** 231, at 21:43, is the lowest first chart of the entire evening; 232, an
hour later, is the highest. If the day were accumulating, 231 could not sit where
it does.

So the honest statement is narrower than either of my first two:

- Back-to-back rounds DRIFT — a trend, five samples, one direction.
- Rested rounds VARY — 35491 to 53138 on the first chart, no trend.
- **232 is not drift. It is the spread.**

Which means the protocol is working and the floor is simply wider than hoped.
That is a worse headline and a better measurement: a floor of "±50% on the first
chart" is still the first honest answer this project has had to "how far apart
can two identical rounds be", and it is dramatically better than the ±5 on one
counter it has been using.

Three readings in ninety minutes — the rest fails, the day accumulates, the floor
is wide — and only the third survived contact with the clock. Written down in
that order deliberately: the first two were reasonable, evidenced, and wrong, and
the thing that killed both was plotting the samples against time instead of
against each other.


## The occupancy reading, first round: POSITION is real, with load held constant

Round 239, the first round to carry a host reading of what each slide held before
the rescale:

    chart      size    ms      slide  shapes   vs fit
    draw-1     1/24    2431    257      3      on fit
    1/8        18/24   37130   257      3      +21800
    2/8        9/16    12157   257      3      +3700
    3/8        9/16    11767   257      3      +3700
    4/8-8/8    18/24   ~17000  258-262  1      on fit
    1/1 alone  18/24   15699   262      1      on fit

**Charts 1, 2 and 3 are on the same slide, holding the same three shapes** —
measured by `getCount()`, not inferred from what a run's batches happened to
record. Chart 1 is +21800 above the fit; charts 2 and 3 are +3700. Load is
constant and the cost is not.

And the slide is not the whole story on its own: a 1-of-24 update on that SAME
loaded slide, moments earlier, cost 2431ms — dead on the fit. Slide 257 is not
inherently slow.

**So position is real.** That is the half of the question that four corrections
could not nail down, and it took one honest measurement rather than another
argument.

### What is still open, stated precisely so the next round can close it

Load's INDEPENDENT effect is not settled by this round. The comparison that would
settle it is the lone arm on a LOADED slide against the lone arm on a clear one,
which holds run-length constant at one and varies only the slide:

    alone, clear slide (measured)     15235ms (round 215) · 15699ms (round 239)
    alone, loaded slide (inferred)    35033ms (round 213) · 39570ms (round 214)

The second row is INFERRED, not measured — rounds 213 and 214 predate the
occupancy reading, so "slide 257 held three shapes then" is an assumption from
the slide number. Plausible, and not evidence.

**The experiment:** make the lone arm pick a chart on the FIRST chart's slide —
the exact opposite of what `pickLoneChart` does today, which was written to avoid
that confound. One round in each arm, with occupancy recorded in both, and the
question is finished.

Worth noting what the instrument cost: one `getCount()` in the scenario, taken
before the loop, adding nothing to the timed path. The question had been open
across four corrections and roughly two hundred rounds.

### THE OTHER ARM, round 240: load is the factor, and "position is real" was overstated

    arm                     ms      slide  shapes   vs that round's later charts
    alone, CLEAR slide    15699      262      1           0.93x   (round 239)
    alone, LOADED slide   41851      257      3           1.66x   (round 240)

**Run-length held at one. Size held at 18 of 24. Occupancy measured on both
sides. 2.7x raw, 1.8x once each arm is normalised against its own round** — and
that normalisation matters, because round 240 ran about 30% slower throughout,
which is exactly what the 14% IQR and 73% range say to expect.

**Load has a large independent effect.** That is the half that was open, and it
is now closed with a measurement on both arms rather than an inference from a
slide number.

### And it takes "position is real" down with it

The entry above says position was demonstrated in round 239 with load held
constant: chart 1 at **+21800** above the fit against charts 2 and 3 at **+3700**,
all three on the same slide. That comparison is between an **18-of-24** chart and
two **9-of-16** charts, reconciled through the per-size fit — and this file
already records that the 16-node fit is built from a single change-count and is
unreliable for exactly this purpose.

**So position's independent effect is NOT established.** What round 240 adds is
the opposite: on the loaded slide, a lone chart (41851) and the first of eight
(47214) cost about the same — both position 1, 12% apart, inside the floor. The
run length does not matter; the slide does.

Corrected ninety minutes after being written, by the round designed to test the
other half. The lesson is the same one this thread has taught four times: a
control that holds one factor constant while comparing across a DIFFERENT
uncontrolled one — here chart size, through a fit the archive distrusts — is not
a control.

**Where the thread actually stands:**

- **Load: measured, large, both arms.** A chart on a slide holding three shapes
  costs ~1.8x one on a slide holding one, at the same size and run length.
- **Position: unestablished.** Every comparison supporting it crosses chart
  sizes, and the size-matched evidence (lone versus first-of-eight, both loaded)
  shows no difference at all.
- **What would test position properly:** two charts of the SAME size on the same
  slide, one first and one later. The harness does not currently produce that.


## READ THIS FIRST is fixed: 36 of 36, against a baseline of 1 of 74

`docs/BACKLOG.md` opens its biggest item with **READ THIS FIRST: a chart on a
FRESHLY ADDED SLIDE cannot be grouped, and that is why it loses its config**,
measured over the whole archive on 2026-08-15 and called *"a switch rather than a
tendency"*:

    slide already had shapes   82 chart(s), 81 grouped = 99%
    freshly added, empty       74 chart(s),  1 grouped =  1%

The current rate, over the last twenty rounds:

    slide already had shapes   36 chart(s), 36 grouped = 100%
    freshly added, empty       36 chart(s), 36 grouped = 100%
    per round: [4/4 4/4 4/4 4/4 4/4 4/4 4/4 4/4]

**Eight consecutive rounds, four of four every time, no exceptions.** 1% to 100%,
on a population of 36 — past the twenty this report requires before it will state
a rate at all.

The chain that made it a switch is broken at its first link: a chart on a freshly
added slide no longer gets a short or empty pre-grouping re-read, so it groups, so
its tag goes through the group handle rather than a `created` one, so it keeps its
config. Everything the backlog item predicted would follow from fixing that link
has followed.

### It was fixed some time ago and nobody could see it

This is not new behaviour. It is newly VISIBLE. The metric reads `0/0` for the
twenty rounds before 228 — not "fixed", not "broken", nothing at all — because
`poolFreshVsEstablished` keys on a per-chart label that the draw path stopped
emitting when in-place updates started succeeding. Restoring that label this
morning is the only reason there is a number here.

So the archive's flagship open problem has been closed for an unknown number of
rounds, in silence, while the instrument that would have said so read zero over
zero. **The fix is not the finding. The finding is that nothing noticed.**

`BACKLOG.md`'s READ THIS FIRST needs a line saying so, and the 1% figure needs
its date carried beside it wherever it is quoted — it is a true measurement of a
host that no longer behaves that way.


## THE NOISE FLOOR, MEASURED — 66%, and it retires the pair as a detector

Five sessions on one build, every sample the FIRST round of its session after a
47-minute rest, position verified from the round file rather than trusted:

    eba1c4d  5 session(s): 230 231 232 233 234
      in-place update, 18 of 24: 15538-25822ms, median 17663 — SPREAD 66% of the min

**Two identical rounds, nothing changed between them, differ by 66% on the most
repeated measurement this harness makes.**

### CORRECTED at eight samples: 21% typical, 73% worst — and 66% was the wrong statistic

The 66% above is the observed RANGE, min to max. A range only ever grows with
sample size, so it is a lower bound wearing an estimate's clothes — and it proved
it within three sessions:

    5 sessions   range 66%
    8 sessions   range 73%     (one faster round arrived; the host did not change)

The interquartile spread does not drift that way:

    in-place update, 18 of 24: 14965-25822ms, median 17544 — RANGE 73% of the min
    middle half 16045-19468ms — IQR 21% of q1

So the floor is better stated in two numbers:

- **Under ~21%: noise.** The middle half of identical rounds spans that much.
- **21% to 73%: ambiguous on a pair.** Inside the observed range, outside the
  typical spread — one round of each proves nothing either way.
- **Above ~73%: larger than anything two identical rounds have produced.**

That is a tighter and more useful bar than 66% was, and it moves the earlier
verdict: a pair CAN see a 2x effect comfortably, and cannot see a 30% one. The
first-chart 2.2x and the lone arm at 130% both clear it with room. The 7%
same-slide comparison is not merely below the floor, it is below a third of the
IQR.

Reported both ways from here, because the range says how bad it has ever been and
the IQR says what to expect — and quoting the first as if it were the second is
the mistake this correction is made of.

### FINAL, nine sessions

    eba1c4d  9 session(s): 230 231 232 233 234 235 236 237 238
      in-place update, 18 of 24: 14965-25822ms, median 16792 — RANGE 73% of the min
      middle half 15538-17663ms — IQR 14% of q1
      scenarios skipped per round: [0 0 0 0 0 0 0 0 0]   failed: [0 0 0 0 0 0 0 0 0]

The IQR TIGHTENED as the sample grew — 21% at eight sessions, **14% at nine** —
while the range sat unmoved at 73%. That is the two statistics behaving exactly
as the correction above predicted: one converging, one pinned by a single outlier
(round 232, the only sample above 20s).

**The floor, as it should be quoted:**

> On one build, first round of a session, nothing changed: the middle half of
> rounds fall within **14%** of each other, and the worst pair seen differs by
> **73%**. Under 14% is noise. Over 73% has never happened by chance.

Nine rounds, nine zeroes for skipped and failed. The completeness result is not a
tendency either.

### The audit the floor demands, started here

A floor is only worth having if the claims already in this file are held to it.
Under 14% is noise; between 14% and 73% a pair cannot tell.

**Failed, and it was quoted as causal evidence:**

- *"Later charts on the first chart's own slide cost 18233 against 17033
  elsewhere"* — **7%**. Half the IQR. It was used to argue the slide was not the
  cause, and it cannot support that or its opposite. Already retracted above on
  other grounds; it fails on this one too.

**Failed, and it was mine, written today:**

- Five statements in this very entry still said "the 66% floor" after the
  correction that replaced 66% with 14%/73% had been written directly beneath
  them. A correction that does not sweep its own section is half a correction.
  Reconciled.

**Clears the bar comfortably:**

- The first chart at **2.2x** and the lone arm at **130%** — both far above 73%.
- Fresh-slide grouping, **1 of 74 against 36 of 36** — a proportion, not a
  timing, and not subject to this floor at all.
- The per-sync write/tag split — a WITHIN-call comparison whose quartiles do not
  overlap; it does not pay the between-round floor.

**Not swept:** `docs/BACKLOG.md`, `docs/ROUNDS.md` and the source comments carry
their own timing claims and have not been checked one by one. The rule to apply
is simple and mechanical — any timing difference under 14% quoted as evidence is
noise, and between 14% and 73% needs more than a pair — and the sweep is worth a
session of its own.

The floor this replaces was one line: *"cabb357 scored 1 and 5 for tags-undefined
with NOTHING changed between them."* Two observations, on a counter, from one
build run twice. It has carried every "is this a change?" judgement in this
archive for two hundred rounds.

### What it licenses, and what it does not

It is a statement about **one round against one round**. It is not the error bar
on a pooled median: forty-five rounds averaged do not carry that spread, and reading it
that way would retire half this file for no reason.

Where it bites is the **pair** — and the pair is this project's unit of evidence:

> **A pair cannot detect an effect smaller than the round-to-round spread.**
> See the correction below for the numbers that replaced the first estimate here:
> under 14% is noise, and anything up to 73% has been produced by two identical
> rounds.

That is not a small correction to practice. `tagAnchorIndex` was judged across
five rounds and four builds and found to have "no measured effect" — with a floor
this wide, that verdict was sound only because the effect was absent, not because
five rounds could have seen a modest one. Any future change judged on a pair
needs to clear that spread or be judged another way.

### What still survives it

- The first chart at **2.2x** a later chart — 120% above, clear of the floor.
- The lone arm at **15235ms against 35033/39570ms** — 130%, clear.
- The per-sync write/tag split, which is a WITHIN-call comparison and does not
  pay this floor at all.

### What does not

- *"Later charts on the first chart's own slide cost 18233 against 17033
  elsewhere"* — a **7%** difference, quoted in this file as evidence that the
  slide was not the cause. Below the floor by an order of magnitude. It was
  pooled across rounds, so it is not strictly a pair comparison, but nothing in
  the report said which it was and nobody could have told.

Five samples is the minimum this report will call a floor and the number will
move as more arrive. It is a first measurement, not a constant.

## Rest does not buy speed. It buys COMPLETENESS, and that is the better prize

The floor above says rested rounds still vary on timing. So does resting do
anything at all? Yes, and not where anyone was looking:

    BACK-TO-BACK (216-225)   10 rounds   10 scenarios skipped   [0 0 0 0 1 2 2 2 0 3]
    RESTED (230-235)          6 rounds    0 scenarios skipped   [0 0 0 0 0 0]

**Zero against ten.** And the shape of the back-to-back row is the whole story:
the first FOUR rounds of that block skipped nothing either. Skips begin at the
fifth round and never really stop.

A skip is not a slow scenario. It is `the host stopped answering during this
scenario, so nothing was checked` — the question was asked and no answer exists.
A round with three skips is a round that measured eleven things instead of
fourteen, and the three it lost are not random: they are the heaviest.

So the operating rule, evidenced rather than assumed:

> **Resting does not make a round faster — the floor is there either way.
> It makes a round COMPLETE.** Run back to back and from about the fifth round
> you start buying rounds that answer fewer questions than they claim to.

That reframes the earlier failed rest experiments. A 15-minute rest did not
restore SPEED and neither does a 47-minute one; the entry that recorded that was
measuring the wrong thing. What the rest protects is whether the battery finishes.

And it explains the crashes without needing the timing at all: all three landed
in `same scale across the deck`, the heaviest scenario, in rounds deep inside a
back-to-back block. The skips and the crashes are the same failure at different
severities — the host declining to finish the heavy work — and rest prevents
both.

### "Skips NOTHING" was refuted by its own checker, minutes after being written

This finding became `rested-rounds-skip-nothing` in `scripts/claims.mjs`. The
first run of that file reported it **STALE**:

    STALE rested-rounds-skip-nothing   1 skip(s) across 12 rested round(s)

**Round 226 is rested, did not crash, and skipped the rescale anyway.**

The claim was true of rounds 230-238 and false of the population. I had measured
a window I chose and stated the result as a property of rested rounds; the
checker applied it to every round that qualified and found the counterexample
immediately. That is the entire point of the file working on its first day, and
it worked against the person who wrote it.

Restated as `rested-rounds-rarely-skip` — a comparison of RATES, which is what
the data supports. That is not a goalpost moved past its counterexample: the
finding was always about the gap between two populations, and "zero" was an
artifact of the window.

**And the restated claim now reports `?`, which is also honest.** The 10-in-10
back-to-back evidence comes from rounds 216-225, which predate `sessionIndex` —
so the machine cannot see the comparison the human can. The strongest half of
this finding rests on rounds recorded before the instrument that would let it be
checked, and it will stay `?` until enough deeper rounds carry the field.

Ten rounds on one build produced within-session drift, not a floor: `laterMed`
doubles across a session, which is larger than whatever run-to-run noise sits
underneath it. Pooling them measures the trend.

A floor needs rounds at a **constant session position**. Concretely:

1. **The first round of N separate sessions**, `--fresh`, on ONE build, with at
   least 45 minutes of idle before each so `sessionIndex` reads 1 — which the
   round now records, so the constraint is checkable after the fact instead of
   trusted.
2. **N of at least 8.** The existing floor came from one build run twice
   (`cabb357` scored 1 and 5 for tags-undefined), which is a range of two
   observations and has been carrying every "is this a change?" judgement since.
3. **Compare `laterMed`, tag faults and scenario skips**, and report the SPREAD,
   not the mean. The question a floor answers is "how far apart can two identical
   rounds be", and a mean cannot answer it.

That is calendar days, not machine hours — eight sessions with an hour between
them is a working day of mostly waiting. It is written down here so it can be
run rather than re-derived, and so nobody pools ten back-to-back rounds again and
calls the spread a floor.


---

## 2026-08-26 — the sheet was measuring a slide the product never uses

Four conclusions were reversed in one day, and they had one cause between them.

The probe sheet asks 33 of its 39 questions on a SCRATCH slide it adds and takes
back. That slide does not behave like the slides a user's charts live on, and the
sheet is this project's main source of "what this host does". So a whole class of
belief here was measured somewhere the product never runs.

It took a new tool to see it. A round costs fourteen minutes and answers thirty
questions at once, which is the right shape for watching the product and the
wrong shape for settling one fact. `scripts/experiment.mjs` puts a single
question to a real slide in about two seconds and cleans up after itself. That
changed what was worth doing: being wrong stopped costing a round.

**What it overturned:**

| the sheet said | the real slide said |
| --- | --- |
| `tags-add-same-key-twice: other`, 224 rounds | `overwrites` — a slide tag lands, a second write replaces it |
| `group-children-via-getcount: unreadable`, 34/40 | `yes` — a group lists and names its children, fresh or aged |
| `grouped-child-by-id`: unanswered in 125 rounds | `no-such-shape` — decisive, and it killed a week of planned work |
| `shapes-items-count-honest: unreadable`, 92% | production reads collections all day — 2,135 charts grouped by id-match |

The third is the only one where the two agree.

### What it cost, concretely

`groupReadRefused` latched for the SESSION on the first `Shape.group` refusal.
Its stated premise was "the feature has never once worked here", citing two of
the scratch-slide answers above. Since the parts list is never consumed,
`queueGroupMembers` is the ONLY route an in-place update has — so one refusal
turned every remaining chart in the round into a full redraw:

    in-place updates after the first refusal: 0, across rounds 254-261
    redraw rate: 21.4%, seven identical rounds

Resetting it per batch: 14.3%, three identical rounds, and 1 in-place update
after the refusal in each. `a-group-refusal-no-longer-ends-the-round` watches it.

### The reasoning error, stated plainly

I concluded that morning that a grouped chart MUST redraw and wrote it into the
backlog and twice into `WHAT-WE-KNOW.md`. The evidence was real and the inference
was not: I had proven that a grouped child is not addressable BY ID, and treated
that as proving the group could not be read AT ALL. Those are different questions
and only the first had been asked.

    by id            slide.shapes.getItemOrNullObject(childId)   does not work
    by enumeration   shape.group.shapes loaded for items/id      works

The same shape of error appeared four times in the day, each time as "one
observation, one mechanism, no check". It is cheap to make and, with the
experiment tool, cheap to catch.

### What the self-test already knew

`selectionWedged` latches the same way when the selection ladder finds a host
that hangs — and it is reset PER RUN, with a comment that reads as though written
for this bug: *"a stale 'this host wedges' from the last round would make the
next one skip a scenario on evidence it no longer has."* The renderer latched for
the session, on the same kind of evidence, three files away.

### The rule that follows

**Do not quote a scratch-slide answer as a fact about the host without checking
it on a real slide** — and never where it is load-bearing for a latch, a
fallback, or a decision not to build something. It costs two seconds now.

The sheet is not wrong and its questions are not wasted. What it measures is
real: the scratch slide behaves this way, and both the probe and the self-test
use one. What it is not is a statement about the slides the product runs on, and
it has been read as one — by this project's documentation, and by its code.

---

## 2026-08-26 — three rounds spent measuring dead code

A retry stood in `updateChartsInSlides` for three rounds and never once produced
an in-place update. Deleting it took less evidence than keeping it had.

### What it was for

When a by-id lookup refuses the resolve, the batch's `groupMembers` proxies die
with the sync they were queued in. The shape itself is rescued a moment later by
`reReadRefusedShapes` — the archive shows `recovered: 1 of 1` in every round on
file — and then the update redraws a chart it is holding a live handle to. The
retry was meant to re-queue that group read and save the redraw.

### Why it never worked, three answers deep

    round 265   got: 0   the in-batch latch made `queueGroupMembers` return
                         undefined — the retry was inert from the day it was
                         written
    round 266   got: 0   latch cleared, still nothing. Diagnosed as "the
                         recovered proxy does not expose `.group`"
    round 267   got: 0   with `viaCollection: 1` — which DISPROVES round 266's
                         diagnosis. The property answered fine; the SYNC threw

Each round bought one bit because `got` could not say why it was zero. That is
the whole cost, and it was avoidable from the first one.

### The instrument was the bug

`got` was `[...requeued.values()].filter(Boolean).length`, and
`queueGroupMembers` returns a PROXY — truthy the instant it is queued, whether or
not the host will honour it. The number reported *asked* and read as *got*.

The test agreed with it:

    expect(again!.data.got, "the retry asked nothing").toBeGreaterThan(0)

Trace, test and code all agreed, and all three were satisfied by the act of
asking. Nothing in the loop could see the feature doing nothing, because every
instrument measured the attempt rather than the result.

**A count taken before the sync cannot report what the sync returned.**

### Round 267 could not test what it was built to test

The by-id route was guarded by `if (!fromRecovered)`. `fromRecovered` is a proxy,
so it was never falsy, so the branch was unreachable. The third round in a row
measured dead code — and the guard was written in the same commit that added the
counters meant to tell the two routes apart.

### The finding, which is smaller than the effort

Both routes into a refused group are shut, for different reasons:

    the recovered proxy   `.group.shapes` answers, then the sync throws
    a fresh by-id proxy   the LOOKUP is what poisons the sync it joins — and a
                          by-id refusal is what sent the chart here

There is no third route, so the redraw is CORRECT. What rescues the group is the
next batch, via the per-batch reset — still measured at 21.4% -> 14.3%.

### Two beliefs corrected on the way

**"The fake cannot model a sync-time refusal."** It can, and does:
`faults.refuseShapeById` sets `refuseThisSync` at `office-host.ts:1722`, and
`pendingHostError` throws ONCE and clears rather than poisoning everything after
it. Two of the three rounds were spent believing the suite could not be asked —
and the suite answers in two seconds.

**"The recovered proxy does not give up its group."** Round 266's guess, written
into a comment as though established. Round 267 disproved it with one number.

### The rule that follows

**Instrument the result, not the attempt.** A counter incremented where a call is
made says only that the call was made. Where the answer arrives at a later sync,
the count belongs after that sync, over values the sync resolved — and a field
that cannot distinguish its own failure modes buys one round of information and
then costs one per question after it.

The same day, the same shape turned up twice more in the triage reader: a
NEVER-VARY section that printed a field's name and sample count but not the value
it never varied from, and a POSITIONAL GUESS section that counted foreign picks
without counting the guard that refused them. Both cost a trip through the
archive to answer a question the tool was already holding the answer to. Both now
print it.

## Window 2026-09-16 — 24 autonomous hours — the one window where a round means what everyone assumed it meant

**Preconditions, recorded because this window's whole value rests on them.**

    tree clean                                    yes
    HEAD                                          1f186b1
    https://ssf-chart.../build.json               1f186b1 · 2026-09-13 14:18Z
    git diff --name-only v0.6.1..HEAD -- src manifest.xml manifest-prod.xml vite.config.ts
                                                  (empty)

**deployed == HEAD == the bytes inside v0.6.1.** Six commits separate the tag
from HEAD and not one of them touches the shipped bundle: they are tests, docs,
CI and Stryker config. So a round run in this window is evidence about **what a
user actually installs**, which no round in the 454-round archive has ever been.
The archive stops at 454 / `454303d` / 2026-09-10; v0.6.1 has zero rounds.

**THE FREEZE, for the length of the window.** Nothing under `src/`,
`manifest*.xml` or `vite.config.ts` is committed. Commits to `scripts/`,
`test/` and `docs/` are batched BETWEEN cycles, never during one. The reason is
mechanical rather than tidy: a push moves HEAD, Pages redeploys, and
`round.mjs` refuses with `site-behind` until the stamp catches up — so a commit
at hour three does not merely risk the property above, it stops the loop. An
agent that "fixes something small" mid-cycle spends the rest of the window
measuring a build nobody can ship.

**A CORRECTION, made before any round ran rather than after.** The plan for this
window claimed the 1.10 floor gave rounds a new invariant to confirm — that
`canRotate()` is now always true, so the picture fallback can never fire. That
is not something a round can see. `supports()` is
`Office.context.requirements.isSetSupported(...)` (`powerpoint.ts:9734`): it
reads what the HOST advertises, and PowerPoint on the web has reported
`[1.1 … 1.10]` since long before the floor moved. The manifest floor governs
INSTALLABILITY and is invisible to this instrument. It can only be desk-checked
against Microsoft's version tables, which is a separate block of this window and
not a round.

**Six rounds, capped, and the cap is the point.** The archive says novelty has
plateaued — "nothing new in the trace — 140 known signature(s)" three legs
running, and nine identical rounds on `eba1c4d` returned all-zero skipped and
all-zero failed. Six rounds mined through the five-part gate beat thirty
archived and unread. Rounds 7-30 would produce p-values nobody reads, on a
question the archive already answers.

**Staged manifest refreshed first.** `C:\OfficeAddins\ssf-charts-manifest-prod.xml`
held `Version 1.0.0.0` / floor `1.4`, dated 2026-09-13 09:24 — before both floor
raises and before the version bump. Now byte-identical to the repo's
`manifest-prod.xml`: `1.0.1.0` / floor `1.10`. SSF Charts is staged in that
catalog and has never been ADDED in PowerPoint, so nothing was consuming the
stale copy; this only means the add-in is the real one whenever it is added.
## Rounds 455 and 456 — b82048a — 19 of 19, twice — THE FIRST ROUNDS ON A RELEASED BUILD

`deployed == HEAD == the bytes inside v0.6.1`, verified before the first round
and unbroken through both. Six commits separate the tag from HEAD and none
touches the shipped bundle, so these are the first rounds in 432 that are
evidence about **what a user installs** rather than about a working tree.

    455   19/19   0 skipped   895 trace entries, 0 dropped   3 attempts
    456   19/19   0 skipped   905 trace entries, 0 dropped   5 attempts
    both  112 of 112 probe questions answered · host PowerPoint · OfficeOnline
          requirement sets 1.1 … 1.10

**The pair is the point and the asymmetry is the finding.** 456 needed five
attempts and four recoveries — `deck-dirty+pane-closed`, `browser-gone`,
`pane-closed` twice — to produce a verdict identical to 455's three-attempt run.
Read alone, either round says "the host is calm". It was not: three browser
deaths in one session (mid-round at 827s and 1074s, plus a gone browser at the
start) and one crash whose file never arrived. The driver absorbed all of it.

### 1. Mine — the deck, and a heuristic that cries wolf

Round 455's fullest slide holds 8 shapes and the gate printed its standing line,
`<- 8 above 6, so a chart was left as loose shapes`. **It had not been.** All
eight are named `PowerChart`, which is the GROUP name — eight grouped charts on
one slide, exactly what `same scale across the deck` is supposed to leave.

Measured across the last twelve rounds whose fullest slide exceeded six:

    all `PowerChart`                    8 rounds   benign (grouped charts)
    `Title 1`, `Subtitle 2`, `title`    3 rounds   benign (layout placeholders)
    `PowerChart` + `category-0`, …      1 round    REAL — round 449

Eleven of twelve were benign: a false-positive rate of 92% on a line that states
a conclusion rather than raising a question. Round 449 is what the real thing
looks like — four `PowerChart` groups beside `title, category-0, category-1,
seg-0-0, seg-0-1, baseline, series-label-0`, one chart left as its parts.

**So the count is the wrong test.** A slide of N `PowerChart` groups is clean at
any N; a slide is dirty when loose CHART-PART names appear at top level. That
test is also strictly more sensitive — a five-shape slide hiding one stray
`category-0` passes the count test and fails the name test.

### 2. Research — nothing upstream to ask

No unexplained host behaviour in either round. `shape-proxy-survives-one-sync`
still answers `threw` and `shapes-items-count-honest` still `short-3`; both are
long-declared divergences, not news.

### 3. Instrument — `sideloadAddIn` gives up on a dialog that merely loads slowly

**The loop cannot survive a browser death unattended, and that is one line.**
A web sideload does not survive the browser process, so every mid-round death
ends with `addin-missing`, which is deliberately outside `RECOVERABLE_STOPS`.
`sideloadAddIn` exists to put it back and walks Add-ins ▸ More Add-ins ▸
MY ADD-INS ▸ Manage My Add-ins ▸ Upload My Add-in. It failed here at
`no 'Manage My Add-ins' control`, twice, and the cycle stopped.

It is not a rename. **The Office Add-ins dialog's content panel intermittently
fails to load**, and when it does none of its controls exist — Microsoft ships a
`Retry` button in that dialog for exactly this. Established three ways on
2026-09-16: it failed on an empty panel; clicking `Retry` by hand made
`Manage My Add-ins` appear immediately, and the walk completed from there; and
later in the same session the walk succeeded unaided, which is what
intermittent means.

The walk needs a retry rung: absent control + present `Retry` → click, look
again. Without it a night ends on the first flaky dialog after any browser
death, and there were three deaths in three hours.

### 4. Fix — batched, deliberately, and this is mechanical rather than tidy

Both fixes are `scripts/` and neither is committed while a cycle runs.
`cycle.mjs` invokes `rounds-gate.mjs` after every round, so an edit mid-cycle is
picked up half-written by the next round; and any push moves HEAD, Pages
redeploys, and `round.mjs` refuses `site-behind` until the stamp catches up. A
commit at hour three does not risk the window, it stops the loop.

### 5. Doctrine — what a round can and cannot see

**A round cannot test the manifest floor, and the plan for this window said it
could.** `supports()` is `Office.context.requirements.isSetSupported(...)`
(`powerpoint.ts:9734`): it reads what the HOST advertises, and PowerPoint on the
web has reported `1.1 … 1.10` since long before the floor moved to 1.10. The
floor governs INSTALLABILITY and is invisible here. Both these rounds report the
full set, and would have done so at a floor of 1.4.

It is a desk-check instead, and it was re-done rather than restated on
2026-09-16: the lowest SUPPORTED Windows channel is 2606 against the 2601 that
1.10 needs, and Mac is 16.112.4 against 16.105. The floor holds on today's
numbers.

**And 19/19 still says nothing about the one thing that changed.** The only
Office.js-reaching behaviour difference between the last round-validated build
(`454303d`) and v0.6.1 is `desc` in `src/core/elements.ts` — eleven additions
flowing `scene.desc → opts.altText → Shape.altTextDescription`, a 1.10 write, at
three call sites. The manifest ships five Elements deep-link buttons. **No
scenario touches Elements at all**, so the green verdict and the changed code do
not overlap anywhere.

### Round 457 — the 4:3 leg, and what it taught about the heuristic above

19/19 as well, at 720x540 from `pageSetup`, on five attempts with four
recoveries. The 4:3 arm is the one the archive puts at **31.1% crashes per
attempt** against 16:9's 3.6%, and it behaved like it: a browser death at
1074s, a deck still holding seven slides from the last round (swept, not
refused), and a crashed run whose file never arrived.

**It also refuted the fix that was about to be written for §1.** The plan was to
replace the count test with a NAME test — loose chart-part names mean an
ungrouped chart. Round 457's title slide holds `Title 1, Subtitle 2` beside
`title, category-0..3, seg-0-0..seg-1-3, baseline, series-label-0/1`: a complete
set of loose parts. That is not a failure. It is `explode a degraded picture`
doing its job, and it passed.

So an exploded chart and a chart that failed to group leave slides that are
identical in count AND in naming. The heuristic cannot be repaired by looking
harder at the deck, because the deck does not hold the answer. What does is the
grouping counter the gate already prints — 20 of 20 attempts grouped, 0 refused.
The line now reports the shape of the deck and points there, and claims nothing.
## Cycle 2 — 0fde152 — a degraded round, and the environment got worse

`deployed == HEAD == 0fde152`, and `git diff v0.6.1..HEAD -- src manifest*.xml
vite.config.ts` is still EMPTY: eight commits past the tag and the shipped
bundle has not moved, so these rounds remain evidence about what a user
installs.

### The driver refused a stale pane before it refused anything else

Attempt 1 opened on `Presentation70` — the 4:3 deck left fronted by cycle 1's
last leg — carrying `pane b82048a`, the PREVIOUS build. It refused twice over:
`deck-missing`, because it was told to run against `Presentation64`, and a stale
pane serving a build that is not the one under test.

Worth recording because the whole three-way-identity argument this window rests
on depends on that check existing. A round that measured `b82048a` while
claiming `0fde152` would corrupt the archive in the way it is least recoverable
from. The two builds happen to be byte-identical in what ships, so such a round
would have been harmless in substance and still wrong in provenance — and the
driver does not know the first part and correctly does not weigh it. It refuses
on the stamp.

### Round 458 — 18 pass, 1 SKIPPED BLIND

    the host stopped answering during `insert onto a slide that already has
    content` — PowerPoint did not respond while drawing shapes 1-10 of 16 (45s);
    the last thing it answered was "writing the chart's origin tag", 0s earlier

`blind: true`, so nothing was checked. A skipped scenario is not a passing one.

### The grouping counter earned its keep one round after the change

The same round reported `14 of 18 attempt(s) grouped, 4 refused`. **Its deck
held 0,3,1,1,2,5,3 shapes per slide — every slide under the ceiling of 6 — so
the fullest-slide heuristic would have been SILENT.**

That is the same argument as yesterday's from the other side. Yesterday the
count test fired on eleven benign decks; today it would have missed a real
refusal entirely. Wrong in both directions, and the counter the gate was
redirected to is the instrument that answers the question.

**AND THE REFUSALS ARE NOT AN ANOMALY, which the first reading of them said.**
The gate prints `usually 0`, that is the MEDIAN, and it was read here as "never
before". Counted properly over the archive — trace entries reading
`not grouping: …`, the same signal the gate uses:

    rounds with grouping activity        431
    rounds with at least one refusal     110   (25.5%)
    recent examples                      407 refused 13 · 440 and 441 refused 7
                                         · 434 refused 3 · 445 refused 2

Four sits well inside that. The first extraction attempted here searched
`trace.summary.steps` for the wrong text and returned a confident zero; the
method above agrees with the gate's own count of 4 on round 458, which is what
makes it trustworthy. Most likely the refusals are part of the same degraded
episode that stalled the host for 45s, and not a finding of their own.

### THE ENVIRONMENT CHANGED, and this is the finding of the cycle

Browser deaths, from `driverRun.recovered`:

    baseline    34 of 307 rounds recording recoveries had one   11.1%
                37 events total                                 0.12 per round
    recent      452:1 453:1 454:0 455:1 456:1 457:1 + 459        5 in the last 6

At an 11.1% per-round rate, five of six rounds affected is about a 1-in-11,000
event. The environment has changed by roughly sevenfold, and it is not
small-sample noise.

Three things follow. It explains this window's friction, which was being
reported round by round as incidents rather than as a pattern: 456 needing five
attempts and four recoveries is the current normal, not bad luck. It makes
**`sideloadAddIn`'s retry rung load-bearing rather than tidy** — every death
destroys the web sideload, `addin-missing` is outside `RECOVERABLE_STOPS`, and
so before that fix roughly five rounds in six would have ended the loop at the
first flaky Add-ins dialog. That was landed for a different reason and turns out
to be the thing holding the loop up.

And the cause is unknown, here as everywhere else: `ROUNDS.md` already says the
losses come mid-round and leave no entry in the Windows Application log. The
rate change is recorded; the mechanism is not claimed.

One caveat on the baseline: `driverRun.recovered` exists in 307 of 431 rounds,
so 11.1% is over the rounds that record it rather than the whole archive.

### And then the question five green rounds never asked

`docs/evidence/elements-alt-text-2026-09-17.json`. Four of the five Elements
write a correct, element-specific `altTextDescription` on the live web host —
harvey, check, flow and kpi, read straight off the slide. The 1.10 write
reaches PowerPoint and the description a user gets is the one the engine built.

`table` is neither confirmed nor a defect: its insert was attempted at 23 shapes
and the host went silent, and the pane said so —
`Failed: PowerPoint did not respond while drawing shapes 1-10 of 23 (45s)`.
That is the same 45s stall that blind-skipped a scenario in round 458, on the
same unwell host. The slide held 52 shapes before the click and 52 after.

**THE PROBE COST THREE SELF-INFLICTED WOUNDS AND THE THIRD IS THE ONE TO
REMEMBER.** It anchored on `tab "Chart"`, which this pane does not have. It
passed `--regex` to a `find` that takes plain text, and took the first ref in
the output rather than the one on the matching line — both traps that
`refFor`'s own comment documents, walked into by copying the SHAPE of that call
without reading it. And its parser could not read the CLI's escaped quotes, so
`JSON.parse` threw and all five elements reported "the host would not list the
slide" while the host was answering perfectly.

The first two refused to measure. The third reported a FAILURE against a
product that was working, which is worse, and is the same shape as every other
instrument fault this archive records: **a parser that cannot read the answer
says exactly what a feature that was never written would say.**

## No round 460 — 2026-09-17 — three attempts, zero archived rounds

**The gap in the archive between 459 and whatever comes next is the HOST, and
this entry exists so nobody re-derives that.** The build under test was
`6cc74c9`; site, pane and HEAD all agreed on it at the start of the third
attempt, on a deck cleaned to one slide.

    attempt   how far it got                               what ended it
    1         refused at the pre-flight                    site-behind + pane shut + deck 4
    2         ready, host answered in 2ms, ran 1322s       the browser process died
    3         ready, host answered in 4ms                  the pane stopped answering

None of the three is a product verdict and none is archived, which is correct —
a round that did not finish must not leave a file that looks like one.

**Four distinct failures in one afternoon, on one host.** `PowerPoint.run(...)`
returning nothing while `Office` and `PowerPoint` are both loaded objects and
the pane answers DOM questions; grouping refused on three of five Elements
inserts at a time, repeatedly, leaving sixty loose parts on a slide; a browser
death 22 minutes into a round; and a pane left holding a crash-recovery offer
whose run button could not be found. The first two ran all afternoon, through
every Elements probe run.

**WHAT WAS WAITING ON A ROUND, AND STILL IS.** `an update follows a moved
chart` now asserts BOTH axes (office-js#6183 — PowerPoint Online updates `left`
and not `top`). If this host does what that issue describes, that scenario goes
red on the next finished round, and **that would be a true finding rather than a
regression** — the scenario compared x alone from the day it was written, so a
chart travelling half the distance the user asked for has been reported as 19 of
19 for the whole archive. It is proven at the unit level: `faults.ignoresTopWrites`
models the host, and without the new y clause that fault passes with "the origin
round trip held". What is missing is the host's own answer.

**One correction, made here rather than left standing.** Mid-recovery I reported
that the browser death had taken the web sideload with it. It had not. The check
ran while the window was 1929px, and at that width PowerPoint hides `Insert
chart` in the ribbon overflow — which is precisely what `MIN_RIBBON_WIDTH`
exists for, and what the driver's own comment calls indistinguishable from an
absent command. After the driver widened to 4632px the command was there. The
recovery was widen, reopen the pane, clean the deck; no sideload was needed or
performed.

**A driver gap found while reading, and deliberately NOT changed tonight.** The
"present but stale" sideload path in `needsSideload` requires `sightings >= 2`,
and `commandWithoutPane` is a module-level counter that resets with the process.
`recover` polls readiness inside one process and can reach it; a hand-run
`node scripts/round.mjs` gets exactly one sighting and returns, so that path is
unreachable from a single invocation. That is defensible for a pre-flight check
— but it means a stale pane will never self-heal for someone running rounds by
hand, and the refusal they see is `could not read the pane's build stamp`, which
does not hint that a re-sideload is what they need. Written down rather than
fixed at the end of a long session on a host that has failed four ways.
