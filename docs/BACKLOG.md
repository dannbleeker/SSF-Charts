# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).

**This is the only backlog document.** Items graduate from here into PRs and are
deleted when they ship — what has shipped is recorded by the README feature
table and by git, not here. Rejected ideas stay in §2 so they aren't
re-proposed, and §3 keeps what the round archive against the live host has
established — `ls rounds/` is the count, and writing it here only ever made it
stale —
because a finding outlives the fix that answered it. **§1 opens with the whole
open list; if a thing is not on that list it is not open.**

Feasibility is judged against the live-add-in constraint: rects, lines, text,
ellipses, and any of PowerPoint's 177 preset geometries (all of them
PowerPointApi 1.4), plus polygon *outlines* — no freeform curves, and no
images. The SVG and skill-pptx renderers additionally have filled polygons and
patterns.

## 0a. A chart is still editable tomorrow — MEASURED 2026-09-06, first time ever

The most basic property a chart tool has, and **no scenario tests it**. Every
archived round runs inside a single session on a deck the driver sweeps
immediately afterwards — that is a property of the loop, not a count that goes
stale — so nothing had ever asked whether a chart survives the document being
closed and reopened.

`Presentation72.pptx` answered it by accident of being a throwaway nobody swept.
Its charts were drawn around 18:00 on 2026-09-05; this was measured the next
morning, across several browser restarts including a `--fresh` that closes the
browser outright, and across four deployed builds.

    the add-in's sideload survived the reopen     yes, `Insert chart` returned
    slides / shapes read back                     51 / 122, unread 0
    shapes still carrying a POWERCHART tag        33, on 32 slides
    a config tag parsed                           334 bytes -> "stacked / Stacked"
    the pane recognised the selection             "An SSF chart is selected on
                                                   the slide. Edit it"
    clicking through                              "Chart loaded — edits will
                                                   update it in place."

Every link holds: save, close, reopen, scan, read the tag, parse it, recognise
the selection, load it for editing. **Zero pages came back short** on a 51-slide
deck, which is also the deepest clean scan on record here.

WHAT IT DOES NOT COVER, so it is not over-read: one chart was loaded, not
edited-and-redrawn; the deck never left OneDrive, so nothing here speaks to
downloading it and reopening in desktop PowerPoint; and it is one document, not
a population. It is a strong single observation of a property that previously
had none at all.

WHY IT IS NOT A SCENARIO: a scenario cannot close and reopen a document. This
stays a hand measurement unless the driver learns to, which is the honest
blocker and the reason it sits here as evidence rather than in the open list as
a task.

## 0. Product health, measured 2026-09-06

The one summary worth reading first, split on whether a build contains the
two-master fix (`6dfaa4b`, 2026-09-04). **As of round 413** — the post-fix rows
grow with every round, so they are anchored rather than left to rot:

    era / arm        rounds   all-green   scenario pass rate
    PRE-fix  16:9       250         194                96.9%
    PRE-fix  4:3         48          28                94.3%
    post-fix 16:9         9           8                99.4%
    post-fix 4:3         28          24                97.1%

    crash records:  94 on PRE-fix builds,  3 on post-fix,  97 total
                    ^ NOT RE-DERIVABLE — see below

The crash line resisted three attempts to reproduce it on 2026-09-06 and none
of them landed on 97: counting `.md` reports by filename date gives 108 / 0,
counting every file in `crashes/` gives 202 / 4, and pairing them gives 103.
The reason is that the directory holds TWO capture mechanisms — a written
report and a downloaded run — which arrive at different moments and did not
both exist for the whole period, so "a crash record" is not one thing. Left
exactly as it was rather than replaced with whichever number looked best.
Whoever fixes it should say what a record IS beside the count.

(An earlier draft of this paragraph added "`crashes/` is gitignored, so this can
only be checked on the machine that runs rounds", which is wrong and worth
correcting rather than deleting: only `crashes/*.md` is ignored. The 98
`-crashed-run.json` files are committed, and are the half a stranger can check.
See that directory's README.)

Re-derived 2026-09-06 evening: the 4:3 row read `26 / 22 / 96.8%` and rounds
412 and 413 had landed since. The other three rows reproduced to the digit,
which is what makes the fifth one a stale count rather than a different method.
That is the FIFTH stale number in three days — `test/backlog-open-count.test.ts`
now guards the one of them that recurred twice, and this row is the argument for
widening it.

**AND THE PRE-FIX ARM IS MEASURED ON A POOL ITS OWN FAILURES WERE REMOVED
FROM.** A round whose host crashed before it could be filed is not in `rounds/`
at all. `rounds-salvaged/` holds 69 of them — every one PRE-fix, because
post-fix barely crashes — and putting them back moves one arm a long way:

    PRE 4:3    archived   28 / 48   58.3%      + salvaged   57 / 81   70.4%
    PRE 16:9   archived  194 / 250  77.6%      + salvaged  199 / 257  77.4%

A salvage exists only when the host died in `collectDeckEvidence`, i.e. after
every verdict was already in, so salvages are selected for having COMPLETED.
Rounds that crashed mid-scenario are in neither pool and are presumably the
worst of all.

**"ZERO POST-FIX SALVAGES, WHICH IS A CONSEQUENCE OF THE FIX" WAS CIRCULAR and
is withdrawn.** Salvaging had already been designed out before the fix landed:
`rounds-salvaged/README.md` records that `app.ts` "now banks the log _before_
the scan, so this should stop". Every salvaged build is dated 2026-08-24 to
2026-09-02; `6dfaa4b` is 2026-09-04. There was no window in which a post-fix
salvage could have existed, so their absence credits the two-master fix with
something a different change caused. **A check with no failing branch is not a
check** — the same shape as a test that passes with and without the code it
guards.

**AND "70% TO 86%" IS ALSO WRONG, corrected 2026-09-07.** An adversarial pass
confirmed every figure above to the digit and then found the thing that
matters, verified here independently: **THE POOLS SAT DIFFERENT EXAMS.**

    archived PRE 4:3   18 distinct scenario names
    salvaged PRE 4:3   15
    archived-only      where a rotated shape lands, stop a run mid-draw,
                       a big chart on a slide of its own

Two of those three are among the most failure-prone in the suite, so the
archived pool was sitting a harder exam than the salvaged one — and post-fix
sits a harder one still, at 19. **All-green is exactly the statistic that
punishes extra chances to fail, so it was never comparable across eras at all.**

**HOW BADLY, inside ONE era and ONE arm, with no build change involved.**
Archived PRE 4:3, split by how many scenarios the round ran:

    13 scenarios    9 / 15   60%
    14 scenarios   17 / 18   94%
    15 scenarios    1 /  5   20%
    16 scenarios    1 /  6   17%
    17 scenarios    0 /  4    0%

From 94% to 0% on the length of the exam alone. (The 13-scenario row breaks the
run and is 15 rounds, so read the trend, not the first cell.) **This undercuts
every row of the table at the top of this section, not just the 4:3 one** — and
it is the one bias in this whole entry that runs AGAINST the fix rather than
for it, because the post-fix arm sat the longest exam of all and still scored
96.6% on the common set.

Scored on the 15 scenarios all three 4:3 pools ran:

    PRE archived                 36 / 48   75.0%
    PRE salvaged, all            28 / 32   87.5%
    PRE salvaged, COMPLETE only  24 / 28   85.7%
    PRE pooled, complete only    60 / 76   78.9%
    post archived                28 / 29   96.6%

**So the honest headline is 79% to 97% on a fixed exam, pooling only complete
salvages** — not 58 to 86, and not 70 to 86. The salvage selection effect is
still real and still large: 85.7% against 75.0% ON THE SAME EXAM is a pool
selected for having finished.

**AND THE POOLING BROKE A RULE THIS REPO HAD ALREADY WRITTEN DOWN.**
`scripts/salvage-crashed.mjs` says, of a round that reached 6 of 16 scenarios:
"has not PASSED the other ten, and a naive rate reads ten silences as ten
successes. So the reach is written down and **anything pooling these must
condition on it**." My first two poolings did not. 18 of the 68 salvaged runs
are partial, and their reach is brutal — `4/14` eleven times, `3/14` once —
each counting as all-green beside rounds that ran eighteen. The rows above
exclude them.

**AND `rounds-salvaged/` HOLDS 69 FILES BUT 68 RUNS.** `041-4275306.json` and
`043-4275306.json` are the same round salvaged twice — identical build, host,
slide size, `salvagedPartial`, missing list and per-scenario milliseconds, from
two crash records five hours apart on 2026-08-29. It is the only duplicate in
the pool, checked by comparing every file's fingerprint against every other's.
Both are 4:3, so every 4:3 salvage figure above is deduplicated; the whole-pool
"69" quoted earlier in this section is a file count, not a run count.

**AND ONLY 39 OF THE 68 CAN ENTER AN ARM AT ALL.** 29 carry no `slideSize`,
because `salvage-crashed.mjs` refuses to CLAIM a profile it could not verify and
parks the reading in `slideSizeUnverified` instead. So 32 reach the 4:3 arm, 7
the 16:9 arm, and 29 reach neither. "69 salvages" describes a directory, not a
population that can be pooled.

**THE SAME HOLE EXISTS IN THE MAIN ARCHIVE and the §0 table does not mention
it.** 54 archived rounds, 023 through 078, carry no `slideSize` either. PRE 16:9
(250) plus PRE 4:3 (48) is 298 of 352 pre-fix archived rounds; the other 54 are
in no row of that table.

**AND THEY ARE RECOVERABLE, so it is a fixable silent exclusion rather than
lost data.** `deck.inventory` is present in all 54 and carries every shape's
`left`. The widest one separates the profiles with a gap nothing falls into:

    16:9   n=262   min 898   max 927.3
    4:3    n= 74   min  90   max 690.9

A threshold at 794 is **336 of 336 correct on every round that does carry a
size**. Applied to the 54 it places 53 as 16:9 and 1 as 4:3, and folding them in
would move the PRE 16:9 row from `250 / 194 / 96.9%` to `303 / 199 / 95.0%` —
all-green from 77.6% to **65.7%**.

**THAT CORRECTION RUNS IN THE PRODUCT'S FAVOUR, WHICH IS WHY IT GETS MORE
SCRUTINY, NOT LESS.** Excluding those rounds flatters the PRE baseline and so
understates the improvement. Two things make it more than a convenient story:
the rule is separable by a factor with no overlap at all, and the excluded
rounds ran the SHORTEST exams in the archive — 12 or 13 scenarios — while
scoring 10.9% all-green. This file's own finding says a shorter exam should
FLATTER a round. These are short and terrible anyway.

**THE TABLE IS NOT RESTATED ON the strength of it.** The rule is a heuristic
that could misfire on a nearly-empty deck, rounds 023-078 are a different era of
instrument maturity as well as of build, and `test/backlog-health-table.test.ts`
now pins those rows to their anchor. Folding them in changes a published
comparison, so it is a decision rather than a repair.

The `scenario pass rate` column in the table above was the comparable one all
along, and it says 94.3% to 97.1%.

**All-green rounds at 4:3 went from 28 of 48 to 24 of 28** — figures that
compare two different exams, kept here because the sentence below quotes them —
and
the crash archive is the blunter number: ninety-four records on pre-fix builds
against two on post-fix — which says "two" where the table above says "three",
and neither could be re-derived; see the flag on the crash line. The direction
is not in doubt and the count is. The rate ratchet has come down with it and nobody
edited it: `a big chart on a slide of its own` reads 224.5 deaths per 1000
against a ceiling of 460, `stop a run mid-draw` 140.4 against 330. Both were
above 320 in early September.

**A CHART-LEVEL NUMBER, pooled for the first time on 2026-09-06, and NARROWER
THAN IT LOOKS.** `insert on top of an earlier run` reports "N of 4 charts
re-editable" every round, and nobody had added them up. Across 386 rounds:

    era     rounds   charts re-editable        lost
    PRE        347    1,379 of 1,388  99.35%      9
    post        39      156 of   156 100.00%      0

**"NINE LOSSES IN 1,535" WAS WRONG TWICE, corrected 2026-09-07 after an
adversarial pass and checked here before being believed.**

First, the denominator was the NUMERATOR: 1,535 is the count that SURVIVED.
Charts placed is 1,544 at that pool, 1,556 as of round 418. Dividing losses by
survivors, in a section about denominators.

Second and much larger: **eight of the nine are not losses.** All nine come from
three rounds, and they are:

    243   3 of 4 charts re-editable                       1 genuine
    360   deck grew by NaN (want 4); 0 of 4               4 instrument
    361   deck grew by NaN (want 4); 0 of 4               4 instrument

`selftest.ts` already says of that NaN verdict, in these words, that it "reads
as measured data loss and was nothing of the kind" — `deckGrowth` could not read
the deck and the scenario reported total loss. I quoted a nine without reading
the three rounds it came from.

**So the number is ONE genuine loss in 1,556 charts placed — 0.064%.** That is
better than what was published, which is its own lesson: the error ran against
the product, and nobody checks a number that flatters the other side.

**AND THAT CORRECTION WAS ITSELF TOO NARROW — fourth revision, 2026-09-07.**
Nobody checked it either, for the same reason: it flattered the product. It
counted ONE scenario, because `insert on top of an earlier run` is the only one
that phrases the result as "N of M charts re-editable". **At least four others
measure the same thing in different words**, and the biggest of them is
`same scale across the deck`, which reports "N still re-editable" against the
deck's eight charts. Pooling every scenario that reports it:

    rounds 023-078    649 charts    88 not re-editable   13.6%
    rounds 079+     4,018 charts    10 not re-editable   0.25%
    TOTAL           4,667 charts    98                    2.1%

**The archive reports 98 charts not re-editable, not one.** 88 of them fall in
the first 56 rounds and stop dead at 078. Of the 10 since, 8 are the NaN
artifact above — so **2 genuine losses in 4,018 charts since round 079, 0.05%**.

The product's reliability still reads well, and now it rests on 4,667 charts
across five scenarios instead of 1,556 from one, and it shows the early era
honestly instead of hiding it behind a regex.

WHAT IT IS NOT, checked rather than assumed. The first draft of this paragraph
said "every scenario that draws charts", which is false: exactly ONE scenario
prints that phrase, and 1,388 / 4 = 347 rounds confirms it arithmetically. That
scenario goes through `insertSlidesFromPptx`, so these charts arrived as
generated slides with their tags PRE-BAKED and were only read back. **This says
nothing about the tag-WRITE path** — the one Draft A is about. It is a real
reliability number for one real route, not a headline for the product.

READ THE 100% WITH ITS SAMPLE. Zero failures in 156 is consistent with a true
rate as bad as ~1.9%: "no loss observed since the fix", not "loss cannot
happen". Re-editable also means the tag read back IN THE SAME SESSION, which
§0a is the only measurement on the other side of. A scan that could not see the
deck is skipped rather than counted as a loss, so a blind host cannot inflate
it.

**THE THIN ARM WAS FILLED THE SAME NIGHT.** Post-fix 16:9 was three rounds when
this table was first written. Rounds 404-407 on Presentation64 are all 19 of 19,
and 408-411 then alternated the two decks — 4:3, 16:9, 4:3, 16:9 — and are also
all 19 of 19. Eight consecutive clean rounds across both arms, taking 16:9 to
nine rounds at 99.4%.

That alternation interleaves the arms in TIME; it does not cross the confound
the rounds gate names. Presentation70 is the 4:3 deck and Presentation64 the
16:9 one, so deck and aspect ratio still move together, exactly as
"cyclePlan has never crossed them" says. Only round 373 has ever crossed it, by
holding Presentation70 at 960pt on purpose.

**And the two crashes that preceded them are worth recording without a
mechanism attached.** Both landed on 16:9 within thirteen minutes (23:34 and
23:47 on 2026-09-05), two in about six attempts against that arm's 4.4%
historical rate, both mid-draw on an in-place update. The gate calls no
regression. They came at the end of a browser session open for five hours
driving a 39-slide deck and a seventeen-minute self-test, and four clean rounds
followed a `--fresh` restart.

**The obvious story is that this is pane age, and this file already refutes it
twice — which is why no mechanism is claimed here.** `sessionDepthWarning`'s own
docstring rules the machine out with "`--fresh` rebuilds the browser every round
and 8.6 GB of 15.7 GB was free at the bottom of the run", i.e. a run that used
`--fresh` on EVERY round degraded anyway. And the rounds gate prints, next to
the aged-pane split it does measure, "--fresh disagreed with the pane in 119 of
them — the flag is not the variable". So a restart is not a freshening, and the
degradation this driver rests for is not something a new browser cures.

What stands is the ordering and nothing more: crashes late in a long session,
clean rounds after a restart, n=6. The honest caveat is the same either way —
any run of clean rounds that follows a restart is measuring the state of a
session as much as the state of a build.

The pooled "4:3 crashes 36.1% of attempts" the rounds gate prints is a
whole-history figure and should not be quoted as current: it is dominated by the
155 pre-fix attempts.

## 1. Open

**Everything actually open, as of 2026-09-06.** The sections below carry the
evidence; this list carries the state. Anything not on it is either shipped,
refused, or a finding rather than a task.

**Six of the eight closed on 2026-08-29; 11 and 12 closed on 2026-08-31; 18
closed on 2026-09-01, 16 the same evening. 19 closed on 2026-09-03, by a round
rather than by work. 20 was added and closed on 2026-09-05, the same day — it
was also a third instance of this file's own trap, written up as a section and
left off this list until someone asked what was open. 17 closed the same
evening, on a deck created for it rather than a harness one. 3 closed on
2026-09-06: its remainder was answered NO, and the budget it argued about was
raised 90 -> 105 on the first evidence that ever bore on it. ONE is open, and
it waits on the owner's GitHub identity:**

> The count here read "six" against a list of four, and then "five" against a
> list of three. Stale both times, and this file states its own tie-breaker:
> "if a thing is not on that list it is not open". Counted from the list, and
> the list is what to trust.


    5  filing this project's host measurements to the office-js tracker
       — docs/OFFICE-JS-DRAFTS.md, nothing filed. Drafts A and B are now DEAD,
       killed on 2026-09-06 by the experiment they were held for: a slide this
       add-in made in a PREVIOUS session tags fine, so neither "add-in
       introduced" nor "a held proxy" is the variable. DRAFT D replaces them
       and is better: `shapes.load("items/id")` returns an EMPTY collection for
       a shape a screenshot shows on the slide, for a while after that slide
       was added, and `getItem` refuses the slide's add-time id with 5010 in
       the same window. Recency is the axis.
       TWO OF THE THREE PRE-FILING TASKS ARE DONE (2026-09-07). The window is
       bounded and there isn't one: on a clean deck the drawn shape is not
       counted at all inside two minutes, 0 of 6 trials, 124 reads, control
       healthy throughout — so #2903's guessed 2-second workaround is not
       enough, and an earlier "1.9 second" figure of mine was an artifact and
       is retracted. The tracker search says THIS IS NOT ORIGINAL: it is
       #2903, closed by an inactivity sweep, and the contribution is the
       measurement its reporter guessed at. What remains is a Script Lab repro
       with no part of this add-in in it — which needs Script Lab installed
       under the owner's account, so it is his either way.
       Submission is the owner's identity and his alone


**The 4:3 arm is no longer on this list, and was never on it as a numbered
item.** It closed 2026-09-05 on fifteen post-fix rounds against four pre-fix
ones, p = 0.0010; twenty-six post-fix rounds now stand against the same four,
one failure among them, **p = 0.00018**. See the two-master entry at the end of
this file.

> NOT RE-DERIVED, and flagged rather than quietly corrected on 2026-09-06. The
> round count has moved on (post-fix 4:3 is 28 as of round 413), but "one
> failure among them" does not match the all-green split in §0 — 24 of 28 — so
> the two sentences are counting different things and this one does not say
> which. Recomputing it would mean guessing at the test that produced it, and a
> guessed p-value is worse than a stale one. **Whoever recomputes this should
> write down the arms and the statistic beside the number**, which is the whole
> reason it cannot be checked now.

**3 — DECIDED, AND TWO THIRDS SHIPPED, 2026-09-04.** The question was "picture
or native shapes for a crowded slide". The answer is "ask, and keep the picture
as the floor":

- `0596b70` — the insert used to promise `"Explode to native shapes" turns it
  back` while `doExplode` refused on the identical predicate. It now says what
  is true: the picture keeps its config tag, so the pane can still reload and
  restyle it, and only the conversion to shapes needs the desktop.
- `3613be8` — a too-dense chart is now OFFERED its own slide as native shapes,
  delivered as a one-slide .pptx in a single call: ~280ms against 46.4s of
  shape-by-shape drawing, whose per-shape cost climbs from 161ms to 784ms as
  the slide fills. The picture is what a refusal means and what every failure
  falls back to, measured from a before-and-after slide count.
- `cf2d01b` — the risk that would have killed it is measured and gone. Four
  inserts on a ONE-master deck reported `mastersAfter: 1` every time, so a
  generated slide does not drag its own master into a customer's template.

WHY IT STAYS ON THE LIST: **the 90-shape budget is in the wrong unit.** It
prices a cost that varies 4-5x with whose slide it is — round 374 measured
`onSlide 38` at 18,054ms per batch against `onSlide 40` at 3,500ms.
`src/core/insert-cost.ts` already holds the honest model, and `SLOW_INSERT_MS`
already disagrees with the shape threshold. Re-express the gate as a TIME
estimate. Do not raise the number on current evidence.

**RESEARCHED 2026-09-06, AND THE ANSWER IS NO: DO NOT BUILD THE TIME GATE.**
The remaining task asked to re-express the 90-shape budget as a time estimate.
Four things kill it, and the first is the one that matters.

**1. The 90 is a CRASH gate, and time is not what crashes it.** Its own
docstring (`powerpoint.ts:4530`) gives two purposes: a chart above it "will not
finish inside the batch timeout" — time — and it "loads the host toward the 'we
ran into a problem' crash". `wantsAutoPicture` is blunter: on the web "nothing
throttles an add-in that asks too much: the tab dies." A time threshold
addresses the first purpose and is silent on the second. Re-expressing a crash
gate in seconds is a category error, not a unit fix.

**2. The crash rationale is FALSIFIED up to 103 shapes.** The archive has drawn
charts above the budget 104 times. Split on the two-master fix:

    era     rounds   103-shape draws   a big chart on a slide of its own
    PRE        352                35                  2 ok / 10 failed
    post        35                69                 34 ok /  1 failed

Same chart, same shape count, opposite outcome. What changed was the SLIDE ADD
(`6dfaa4b`), not the density. Of the post-fix draws, 35 queued every shape they
meant to; the 34 that stopped short are `stop a run mid-draw` doing exactly what
its name says. Density was never the killer here — a documented API misuse was.

**3. The replacement model is worse than the thing it replaces**, which is the
measurement immediately below: never validated out of sample, and scored on the
103 rounds it was not fitted to it reads +99 / +24 / +23 / -61%, non-monotonic.

**4. There is no external number to converge on.** Microsoft's resource-limits
page says the CPU and memory thresholds "apply to add-ins running in Office
clients on Windows and Mac, but not on mobile apps or in a browser" — so this
repo's claim is verified, and the web host gives no warning before it dies. No
maximum shapes per slide is published anywhere. And there is no prior art to
copy: think-cell is a desktop COM add-in and never meets this constraint.

WHAT THIS LEAVES is a smaller and better question than the item asked: **the 90
is now known to be CONSERVATIVE.** A user inserting a 103-shape chart gets a
picture today, and the archive says that chart draws. Raising the number to ~105
would make waffle charts native and re-editable. That changes what a user
RECEIVES, so it is the owner's call; the evidence for it is 35 completed draws
of one chart on two decks — narrow but real.

WHAT WOULD CHANGE THE PICTURE: drawing a 176-shape area chart or a 253-shape
violin on a host. Neither has ever been attempted. The gate exists for those and
they are entirely unmeasured, which is the honest reason not to raise it far.

**DEEPER WEB RESEARCH, 2026-09-06 — it strengthens the answer rather than
complicating it.** Four things, and the third is new to this file.

**a. The platform statement is verbatim.** Microsoft's resource-limits page:
"The following runtime resource limits apply to add-ins running in Office
clients on Windows and Mac, but not on mobile apps or in a browser." The desktop
limits it lists are specific — 90% CPU sampled three times in five seconds, a
dynamic memory threshold, a **crash tolerance of four crashes per document
session**, and a five-second unresponsiveness threshold that restarts the
add-in. On the web an add-in gets none of it: no warning, no restart, no budget.
The tab is the only thing that fails.

**b. Microsoft DOES publish a web payload ceiling — for Excel, not PowerPoint.**
"Excel on the web has a payload size limit for requests and responses of 5MB.
RichAPI.Error will be thrown if that limit is exceeded", plus 5,000,000 cells per
read. Where a ceiling exists, the platform's pattern is one that THROWS at a
payload size. Nothing equivalent is published for PowerPoint. Worth noting this
repo's own MAX_PICTURE_BASE64 is 4 MB, just under Excel's number.

**c. untrack() DOES NOT EXIST IN THE POWERPOINT API, and that matters here.**
Microsoft's performance guidance names it as the remedy for exactly this
workload: "Large batch operations may generate a lot of proxy objects that are
only needed once by the add-in and can be released from memory before the batch
executes... Calling untrack() after your add-in is done with the object should
yield a noticeable performance benefit when using large numbers of proxy
objects." Counted in @types/office-js: **Word 174 declarations, OneNote 35,
Excel 3, PowerPoint 0.** This project had already measured it empirically and
files it as untrack-available-on-shape = no, stable, **383 rounds**, asked of a
real created shape proxy rather than a null object.

So the one documented lever for "too many proxies" is unavailable to a
PowerPoint add-in. renderShapesChunked holds one proxy per shape for a whole
draw and can never release them.

**d. Nobody has reported a shape-count crash upstream.** Searches of
OfficeDev/office-js return zero issues for "too many shapes", "PowerPoint
shapes.add slow" or "PowerPoint web timeout sync shapes". #6363 is PowerPoint on
the web, 41 comments, Status: under investigation, regression — but it is about
properties not being available after a sync, not about density. Weak evidence
either way, since this repo may be close to alone in drawing hundreds of native
shapes from an add-in on that host.

**WHAT (c) DOES TO THE RECOMMENDATION.** It sharpens the case for keeping a
SHAPE COUNT rather than a time estimate. On this host an add-in has no resource
budget, no untrack and no published ceiling, so the only quantity it can
actually control is how many shapes it sends. Time is an OUTPUT of the host's
mood; shape count is an INPUT this code chooses. A gate belongs on the input.

**The measurement that got here, kept below.**

**MEASURED 2026-09-05: the replacement model
is in the wrong unit too.** 4,137 timed batches across 360 rounds, every one on
a build containing the `onSlide` retag (`ee1741e`), so all readings share a
definition.

**Occupancy is real, and here is the measurement that cannot be argued with.**
PAIRED, inside a single draw, where the only thing that changes between two
consecutive batches is that ten more shapes are on the slide. Batch 2 is slower
than batch 1 in **1,270 of 1,330** 24-shape draws — 95% — median 3,693 → 5,548,
a 1.5x climb for ten shapes. The 103-shape chart runs the series out to ten
positions, n=16 at every one, without a single reversal: 1608 2107 2882 3664
4310 5141 6081 6658 7394 8293. Chart, path and scenario are all held constant;
only the slide fills.

**But the four coded constants are NOT a clean measure of it, and "the curve
has held up" was circular.** I wrote that twice before testing it. The
constants were fitted at `cc59d4d` on the 261 rounds then archived, and 103
rounds have been added since. On the rounds it was FITTED ON, the curve matches
to -4.2 / +2.1 / -1.8 / -0.3% — which is what fitting means, not what holding
up means. Held out on the 103 rounds it never saw:

    bucket    coded    held-out (n=722)    off by
      0        3886      7717 (n=236)       +99%
      1-20     5490      6813 (n=230)       +24%
     21-50    13995     17236 (n=189)       +23%
     51-100   18074      7097 (n=66)        -61%

Non-monotonic, and the anchor the pane quotes for a blank slide is 2x low.

The reason is that **the occupancy buckets are chart-mix buckets.** In the
fitted set, bucket 0 is 89% 24-shape charts — the cheap kind — and bucket
51-100 is 98% 16-shape charts, the dear kind. Much of the climb across buckets
is the mix changing rather than the slide filling. In the held-out set bucket
51-100 is 98% 103-shape charts and reads 7,097ms against 18,016ms in the fitted
set: same occupancy, different charts, 2.5x apart.

So there are two real effects and one axis carrying both. Occupancy is the
first, and the paired numbers above are what it actually costs. The second is
what the chart IS, and the model has no term for it at all.

`estimateInsertMs` is `ceil(shapes / 10) × batchMs(present)`, which says every
ten-shape batch costs the same. They do not. At a chart's first batch, held to
a slide this run had genuinely drawn nothing on:

     16 shapes    10,098ms   n=258
     24 shapes     3,617ms   n=1,202
    103 shapes     1,584ms   n=15

**2.8x between the first two, backwards to shape count, and 6.4x across the
table.** Three confounds were ruled out. *Batch position*: a 24-shape chart
contributes cheap later batches a 16-shape one never reaches, so this is batch
1 only. *Draw path*: size and target slide are nearly collinear here, but
16-shape charts are slow on both. *In-place updates*: `updated only the shapes
that changed` writes a median of 8 of 24 shapes and would have poisoned the
cheap arm — but it emits no `batch issued` line at all, and 0 of 2,714 batch-1
readings sit on a chart that emitted one.

**THE FIRST VERSION OF THIS TABLE WAS WRONG BY 13x AND SAID "identical
occupancy" WHEN IT HAD NONE.** It read 21,094 / 13,688 / 3,780 / 1,608 for 11
rotated / 16 / 24 / 103 shapes. A batch-1 trace line is written BEFORE the host
answers `slide.load("id")`, so it keys on the `(visible)` sentinel — and the
retag empties that sentinel after every draw. All 775 such lines report
`onSlide 0`, and **516 of them, 66.6%, were on a slide this run had already put
a median of 16 and up to 88 shapes on.** The rows were not at one occupancy;
they were at 0 / 0 / 32 / 48, in exactly the order of their costs. The true
prior is recoverable only from the NEXT batch — `batch2.onSlide` minus what
batch 1 drew — which is how the corrected table is built. The 11-shape rotated
chart that headed the old table has NO true-zero readings at all (45 of 45 on
occupied slides), so it cannot appear here; its 20,662ms is an occupancy
number.

Found by an adversarial pass over the committed text, then reproduced
independently before anything was edited. The correction cuts the effect from
13x to 2.8x and it is still the wrong unit — but "wrong by 13x" was itself the
kind of claim this file exists to stop.

**~~Half the batches are never timed, and not at random.~~ FIXED IN `6b0b2a4`,
AND THE FIRST READINGS AGREE WITH THE HOLD-OUT.** `prevBatchMs` was only
readable from the following batch, so every draw's last batch was unmeasured
(4,583 of 8,771) and 1,869 single-batch draws contributed nothing. All 4,188
readings drew exactly ten shapes; not one was a partial tail. The sizes drawn
are 7, 9, 10, 11, 14, 16, 24, 103 — and all 1,799 draws at 7, 9 or 10 shapes
were untimed. The model priced single-batch charts from a sample containing
none of them.

`last batch settled` now carries them. Eight rounds (388-395) produced 96 new
readings, 40 of them single-batch draws that were previously invisible. Held to
those drawn at **true zero occupancy**, which is the only like-for-like:

    7 shapes    n=8    med 5,683ms     model says 3,886    1.46x low
    9 shapes    n=8    med 8,547ms     model says 3,886    2.20x low

**A second, independent line of evidence for the hold-out's answer.** The
hold-out put the blank-slide anchor at +99% using data this line could not see;
these are charts the hold-out could not see either. Two methods, no shared
readings, the same direction and roughly the same size.

It also extends the size inversion further down than the archive could reach: a
**9-shape** chart costs 8,547ms where a **24-shape** chart's first batch costs
3,617ms — and a 7-shape one costs 5,683ms. Neither ordering nor magnitude
follows the count.

**MY FIRST READING OF THIS QUOTED 2.1x, AND IT WAS THE OCCUPANCY MISTAKE AGAIN
IN A NEW PLACE.** It pooled all 40 single-batch draws, and those draws sit at
occupancies 0, 7, 14 and 21 — four to a slide. Conditioning on a true prior of
zero is what gives 1.46x and 2.20x above. Third time occupancy has hidden
inside a pooled median in one day; the standing rule now is that no median from
this archive is quoted without saying what occupancy it is at.

Remaining caveat: a last batch is not a first batch in general — it sits at the
draw's maximum occupancy and draws a partial tail rather than ten. The
single-batch subset is the clean one, because there the last batch IS the first
batch, and it is the only subset quoted.

**A RASTERISE DOES NOT SLOW THE NEXT DRAW — the lead this line produced was
killed by the experiment's own control, 2026-09-05.** Worth the whole entry,
because the false version was convincing.

Those single-batch draws are the rasterise scenario's, which is already a
counterbalanced A/B — "after a rasterise" against "after a cheap read" — and
until `6b0b2a4` it could only be read for STALLS, where it says 0.0% against
0.1% over 359 rounds. Read for DURATION over eight rounds it looked like a
find: rasterise arm median 8,888ms against 7,509ms, and paired per round the
rasterise arm slower in **8 of 8**, p = 0.0039 on a sign test.

It is an artefact, and the scenario was built to catch exactly this. The design
is ABBA — rasterise at positions 0 and 3, cheap read at 1 and 2 — so splitting
by half is the control:

    rasterise slower in the EARLY half   0 of 8      p = 1.0
    rasterise slower in the LATE  half   8 of 8      p = 0.0039

Perfectly symmetric, which is what "no effect plus a position trend" looks like.
The trend is the whole story: in all eight rounds the four draws climb
monotonically by POSITION, and pooling an arm's two positions left a residue
that happened to lean one way every time.

**And the trend is not position either — it is occupancy, measured more cleanly
here than anywhere else in the archive.** All four draws land on the SAME slide
and each draws 7 shapes, so `onSlideAfter` runs 7, 14, 21, 28 and the prior
runs 0, 7, 14, 21:

    prior 0    n=8    p25 5,660   med 5,683   p75 5,958
    prior 7    n=8    p25 6,721   med 7,128   p75 7,446
    prior 14   n=8    p25 7,810   med 7,847   p75 8,319
    prior 21   n=8    p25 9,072   med 9,282   p75 9,384

Identical chart, identical size, one slide, one scenario, monotonic in every
round — **+63% for 21 shapes, about 260ms per shape already present.** That is
the second independent within-slide series, alongside the 103-shape chart's ten
batches, and the two agree that occupancy is real and roughly linear at this
scale.

The lesson is filed with the finding: this scenario's ABBA was designed for the
stall question years before anyone read it for time, and it paid off on the
first duration reading taken from it. A pooled per-round comparison would have
published the opposite.

WHAT those 16-shape charts contain that costs 2.8x is the open question, and it
is the term the model needs.

**AND THIS ARCHIVE CANNOT ANSWER IT. NO OBSERVATIONAL SPLIT OF IT CAN.**
Established 2026-09-05, after the last candidate explanation died.

The candidate was context: `batch issued` records what the host last answered
before each draw, and sorting batch-1 readings at zero occupancy by that field
separates them beautifully — everything after a per-slide operation costs
~3,400-3,700ms, everything after a deck-wide scan costs 7,500-10,300ms. It
looks like a finding, and it even suggests a fix (do not scan the deck
immediately before drawing).

It is dead, killed by the one cell that breaks the pattern: **103-shape charts
are preceded by `listing the deck's slides` in 23 of 23 readings and are the
CHEAPEST batch in the archive at 1,611ms.** A deck scan does not slow the next
draw.

What that leaves is the structural problem. Every candidate variable is
perfectly collinear with WHICH SCENARIO a draw belongs to, because each
scenario draws one chart in one context:

    16-shape charts    273 of 274 after a deck scan      med 10,046ms
    24-shape charts      0 of 1,210 after a deck scan    med  3,615ms
    103-shape charts    23 of 23 after a deck listing    med  1,611ms

Chart size, preceding call, scenario and slot are one variable wearing four
names. Occupancy, batch position, draw path and in-place updates have all been
ruled out; what is left cannot be separated by any grouping of these rounds.

**THE EXPERIMENT THAT WOULD ANSWER IT — BUILT 2026-09-05 (`0461a3e`), AND ITS
FIRST RUN BROKE IT (`afd5f5d`).** `what a chart kind costs` draws four kinds
back to back on one slide, forwards then backwards so each appears once early
and once late, deleting each specimen before the next. The verdict is
deliberately about whether specimens LANDED, not about cost: a verdict that
failed on a timing would go red on host weather, which
`rasteriseArmVerdict` records the cost of.

Rounds 398 and 399 ran it and came back **19 of 19, green both times** — and
the cost readings said the design was wrong:

    round  pos  kind        shapes  prior      ms
      398    0  clustered        7      0    5519
      398    1  line            10      7    9926
      398    2  area             7     17    8093
      398    3  pie             37     60    3688
      398    5  area             7     74   23588
      398    7  clustered        7     81   29235

`pie` draws THIRTY-SEVEN shapes at this size. So it was multi-batch — making
`last batch settled` its tail rather than its cost — and it drove the prior
occupancy to 60 before the second half began, reaching 91, where a draw costs
three to five times what it does at zero. **The palindrome balances POSITION.
It only balances OCCUPANCY when the specimens are the same size**, and that was
not checked. The confound this scenario exists to escape had been rebuilt
inside it, invisibly, behind two green rounds.

Corrected by measuring first. Shapes per kind across four candidate boxes:

    clustered   7  7  9  9        line       10 10 10 10
    funnel      8  8  8  8        waterfall   9  9  9  9
    area        8 11 15 23        pie        37 37 37 37

`area` is out as well, for a reason the first pick missed: its count varies
THREE-FOLD with the box, so it cannot hold occupancy equal even against itself.
The four kept are flat everywhere and within three shapes of one another, and
still span four ways of making a shape — plain rectangles, filled trapezoids,
rectangles mixed with connector lines, and open strokes. Mean prior per kind is
now 30.5 / 30 / 29.5 / 29.

The constraint is a TEST now, not prose: specimen counts within 4 of each other
and none above 10, at every box a specimen plausibly gets. It fails on the exact
set that shipped and names it.

**FIRST READING FROM THE CORRECTED DESIGN — rounds 400, 401, 402, and it points
the OTHER WAY.** All three came back 19 of 19 with 8 of 8 specimens landing, and
the priors fell exactly where the design says they should: 0, 7, 15, 24, 34, 44,
53, 61.

Each kind's own early/late pair gives its own occupancy slope, so each is
corrected to a genuinely empty slide by its OWN climb rather than a pooled one:

    kind         shapes   cost at prior 0   ms per shape   slope ms/prior-shape
    clustered         7             6,011            859                    181
    funnel            8             8,360          1,045                    243
    waterfall         9             8,464            940                    201
    line             10             9,234            923                    305

**Cost rises monotonically with shape count — 6,011 to 9,234 across 7 to 10
shapes — and per-shape cost is nearly flat, 859 to 1,045, a 1.2x spread.** So
when kind is compared at equal occupancy with position balanced, the count
explains the cost and the kind adds at most about a fifth. That is the opposite
of the 13x and 2.8x inversions read off the archive, and it is what those
inversions being occupancy and context artefacts looks like.

WHAT IT DOES NOT SETTLE, and the distinction matters. It does not explain the
16-shape against 24-shape gap (10,098ms against 3,617ms, both at a true prior of
zero). Those are two different scenarios in two different contexts, and this
experiment only shows that GEOMETRY — rectangles against trapezoids against
mixed rects-and-lines against open strokes — is not what separates them. The
gap is still open; it is now known not to be the shape kind.

Held deliberately loose, because three rounds is three rounds: the correction
uses a two-point slope per kind and those slopes range 175 to 305, every
specimen is 7 to 10 shapes so nothing here speaks to a 103-shape chart, and
"per shape" is not constant across scales — these tiny charts run about 900ms a
shape against about 360 for a 24-shape chart's first batch. More rounds are
free from here.

So the gate would move from a number that is wrong in a KNOWN direction (90
shapes ignores occupancy) to one wrong in an unknown one — time, priced by a
model whose blank-slide anchor is 2x low on every round it was not fitted to.
That is not an improvement, and the honest order this file already argues for —
measure first, decide second, edit numbers third — says stop at the
measurement.

WHAT WOULD UNBLOCK IT, and it is now two things rather than one:

1. **A term for the chart**, not just its shape count. Rotation is the dearest
   property seen; beyond that the archive cannot say, because it only ever
   draws eight distinct chart sizes.
2. **A refit validated on rounds it was not fitted to.** No version of these
   constants has ever been, which is why the failure above went unnoticed for
   nine days. Any replacement must be held out before it ships.

Until then the 90 stays, and it stays understood rather than defended. Nothing
should be tightened against `estimateInsertMs` in the meantime; treat its
output as a floor.

### 20 CLOSED 2026-09-05: the own-slide offer keeps its advice and loses its numbers

I first wrote that the model error was conservative, because a low estimate
makes `isSlowInsert` fire less often. That is true of the blank-slide anchor and
false of the thing that actually gates the offer, so it is corrected here.

`worthOwnSlide` fires only when moving the chart at least HALVES the wait:

    estimateInsertMs(shapes, 0) <= estimateInsertMs(shapes, present) / 2

That is a test on the RATIO of the curve's ends, and the curve's ends are the
two points the hold-out moved furthest and in opposite directions. Coded, the
ratio is 3,886 to 18,074 — 4.65x, so the halving test passes almost everywhere.
On the 103 rounds the curve was never fitted to, the same two points read 7,717
and 7,097: a ratio of 0.9, where a fresh slide saves nothing at all.

Swept across chart sizes 16-103 and occupancies 0-90, **18 of 36 cells flip.**
Every flip is the same direction — the offer fires today where the out-of-sample
reading says it should not. And the sentence it puts up quotes both numbers: for
a 40-shape chart on a slide holding 90, the pane says about 72 seconds here
against about 16 on a new slide, where the held-out reading is about 28 either
way.

This is exactly the failure `worthOwnSlide`'s own docstring exists to prevent —
"an offer that cannot deliver is worse than silence: it spends the user's
attention and their trust in the next warning."

**What is NOT established, and why this is a decision rather than a fix.** The
held-out set is 722 readings over 103 rounds, and its 51-100 bucket is 98%
103-shape charts (n=66) — the cheapest per batch of anything measured. So the
hold-out is composition-skewed in its own way and is not proof the offer is
wrong. What it is: the only out-of-sample evidence that exists, and it inverts
the premise. The honest statement is that the offer rests on a ratio nothing has
ever validated.

    THE OWNER'S CALL, three options and my recommendation is the second:
      a) leave it — the offer is useful more often than not, and the evidence
         against it is thin and skewed
      b) keep the offer, drop the two quoted SECONDS from `offerSentence` and
         say only that this slide is crowded — the advice survives, the number
         nobody can defend goes away
      c) gate the offer off on the web until a validated refit exists

    (b) costs a catalogue key and no behaviour. It is reversible the moment the
    refit lands, and it stops the pane asserting a figure that the one honest
    check available says is 2-3x out.

**THE OWNER CHOSE (b), AND IT SHIPPED THE SAME DAY.** The sentence was

> This slide already holds 35 shapes, so adding here takes about 40 seconds. On
> a new slide, about 15 seconds.

and is now

> This slide already holds 35 shapes, which is what makes adding here slow. A
> new slide would be quicker.

`offerSentence` takes only the count; `offerOwnSlide` no longer takes the two
durations; the pane no longer computes them. `describeMs` went with them — it
had four tests of its own and `offerSentence` was its only caller, and a tested
export the product does not call reads as coverage without being any. Its tests
were deleted rather than left green over nothing.

WHAT SURVIVES IS WHAT IS MEASURED. The shape count is read off the slide by
`getSlideShapeBounds`, not modelled. And the DIRECTION is supported three
independent ways: paired inside single draws, where batch 2 beat batch 1 in
1,270 of 1,330; one chart's own ten batches climbing 1,608ms to 8,293ms; and a
7-shape chart at priors 0/7/14/21 costing 5,683/7,128/7,847/9,282. "This slide
is full, which is what makes it slow, and a new one is quicker" is measured.
"By forty seconds" was not.

The gate is unchanged — `worthOwnSlide` still decides WHEN to ask, on the same
ratio. That is deliberate and it is the part still resting on the unvalidated
curve: this change stops the pane ASSERTING a number it cannot defend, and does
not pretend to fix when the offer fires. Narrowing the gate needs the refit.

Three mutants killed: a duration smuggled back into the sentence, the advice
dropped so only the fact remains, and an estimate returning to the call site.

**And `present` means two different things either side of the model.** The
curve is indexed by the renderer's `onSlide`, which counts only the shapes THIS
RUN drew on that slide. The pane calls `estimateInsertMs` with
`occupied.length` — `getSlideShapeBounds`, the shapes actually there, whoever
put them there (`app.ts:2461`). So a user's slide holding forty shapes we never
touched is priced from readings taken where the renderer had just drawn forty
itself, into a context still warm from drawing them. Whether those cost the
same is untested in 360 rounds.

A first pass at this said the archive contained one crowded-slide reading, from
counting `onSlideKey === "(visible)"` batches. That was wrong and is recorded
because the mistake is reusable: all 763 of them are FIRST batches, where the
host has not yet answered the slide id, so the sentinel and the zero are both
artefacts of timing rather than facts about the deck. `onSlide` cannot answer
the question at all — it is the wrong counter, not a thin one. The missing
measurement is a round that inserts onto a genuinely pre-loaded slide, the same
deck-preparation problem as item 17.

Also unverified: that the densest shipped chart — the 401-shape hex tile map,
~152KB of base64 — actually lands through this path on a host. It has only been
measured offline.

**19 CLOSED 2026-09-03, and the answer is yes.** Round 374 drew a 103-shape
chart in ELEVEN batches onto a slide the add-in had just added, on the deck
that had never managed one: `upTo` 10, 20 ... 103 with `onSlide` climbing 0 to
100, no crash, first attempt. The offer survives a multi-batch chart.

It could not have been answered before because the add itself was broken — the
slide was rejected by the server and rolled back before a shape reached it. See
the two-master entry at the end of this file. What is left on that scenario is
the CLEANUP, which is item 3028's add-time-id defect and is unrelated.

**17 CLOSED 2026-09-05 — the second page ran, and it works. What fails is the
LAST page.** Evidence: `docs/evidence/17-deck-scan-second-page.json`.

Not on a harness deck and not by preparing one: `Presentation72.pptx` was
created blank for the purpose, sideloaded, and given the demo deck by the file
path — 38 items, 0 lost, 14.5s — for 39 slides. The self-test then ran on it and
grew it to 45. Both are things the driver cannot do, because `cleanDeckScript`
deletes every slide but index 0 on a positional rule with no name, tag or
session test, `readiness` refuses `deck-dirty` above one slide, and the heal
runs BEFORE `--check` returns. Nothing about item 17 was reachable through
`round.mjs`; it needed a deck the driver never sees.

**Three pages, 87 attempts, and the second page is clean.**

    from  attempts  slides  charts  unread  short  tagsUnread   failed
       0        29      20     464       0      0           0        0
      20        29      20     539       0      0           0        0
      40        15     5-6      68       0      0           0       15

The middle column is the one this item existed for: **29 pages at `from: 20`,
every one complete, against 379 archived rounds where `from` was 0 every time.**
`chartsSoFar` climbs 16 → 32 → 35 across the pages, which is the field added to
prove the MERGE happens rather than merely the page returning. Median complete
scan 4,392ms over 43-46 slides.

**And the last page failed 15 times out of 15, always the same way:**

    InvalidParam passed to GetItem(id) | at=reading chart tags on slides 40-44 | code=5010

Same error, same code, same call as item 5's Draft A. That is a second,
independent manifestation of the same defect at a different call site, and it
strengthens the filing rather than being a separate problem.

**SEPARATED, AND IT IS NEITHER.** The tail page is both partial and made of
recently added slides, so both were candidates. Ordered in time, the outcomes
are a step function:

    tail page, in order:  11111111111111000000000000000

Fourteen came back, then fifteen failed, and it never recovered — while pages 0
and 20 went 29 for 29 over the same span. Deck size was flat at 45 across the
transition, so size is out; the failures begin BEFORE the two traced slide adds,
so recency is out; and the page had already succeeded fourteen times, so being
partial is out.

What is at the transition is the mechanism itself. The scenario running is
`explode a degraded picture`, working slide `296#19242001` — which sits in the
tail range — and in that window the trace holds `reading back an ungrouped
chart's shape ids → InvalidParam passed to GetItem(id)` twice, a cold re-read
listing **0 of 24 shapes drawn**, and `not grouping: no member handle this host
will accept`. That is `collection-read-poisons-the-creation-handle` happening
live at a call site, and from that moment the page containing that slide could
not have its tags read.

**It is not permanent.** Three same-scale runs on the settled deck twenty
minutes later all scanned completely. So the state belongs to the run, not to
the slide or the document.

**And the product is honest about it, which is the part worth keeping.**
`doSameScale` checks `scanIsComplete(scan)` before doing anything and refuses
with "Same scale needs to see the whole deck first — {gap}. Try again; nothing
was changed." A deck-wide action never silently operates on a page it could not
read. The scan marks itself `complete: false` and 15 of 29 scans in that run
did so. This was checked because the failure mode it would otherwise be — a
deck-wide rescale quietly skipping five slides — is exactly the kind this
project has shipped before.

**AND THE OBVIOUS IMPROVEMENT IS DEAD ON ARRIVAL — checked 2026-09-05 before
building it.** The tempting change is to retry a short page in a FRESH context
before telling the user to try again: the answer sheet says a context degrades
after one collection re-read (`how-many-collection-reads-a-context-survives →
short-at-1`), so a new context should be the cure, and the state did clear on
its own twenty minutes later.

It would do nothing. `readChartsPage` already runs each page inside its own
`ppRun`, and `ppRun`'s own docstring records that "real Office.js hands out a
fresh context per run". So the 15 consecutive failures were 15 consecutive FRESH
contexts, across 15 separate scans spanning about seven minutes. A retry would
add a page-read of latency and refuse anyway.

What cured it was time on the order of twenty minutes, or the run ending — not
a new context. So the refusal stays as it is, and any future attempt at this
should start by explaining what the retry would do differently from the fourteen
that already happened.

**The premise this item was written on was also wrong**, and it is worth saying
because it inflated the risk: the page size of 20 was justified by "the web
>50-item load ceiling (office-js#4272)". #4272 was read directly on 2026-09-05 —
its threshold is items in ONE LOAD, driven by shapes and tags, not slides in a
deck, and it is open since 2024-03-19 with zero comments and no Microsoft reply.
`READBACK_PAGE` bounds slides and never bounded items. See its docstring.

**Originally:** 17 is still open, and now has a number. Across 350 archived rounds carrying
a trace, `deck scan — settling a page of slides` has NEVER been emitted with
`from > 0`. Not once. The second page is unexercised on a live host, exactly as
this item says, and no round can exercise it as things stand: the driver sweeps
the deck between rounds, so the scan always runs against 7-9 slides against a
page size of 20. Closing it needs a deck deliberately left holding 21+ slides
at scan time, which is a deck-preparation step and not a code change.

**AND THE OBVIOUS SHORTCUT IS A TRAP — considered and rejected 2026-09-05.**
The tempting move is to make `READBACK_PAGE` settable, shrink it to 3 in a
scenario, and let an ordinary 7-9 slide deck page twice on a live host. It
would go green, it would look like this item closing, and it would prove almost
nothing. The loop is ALREADY covered against the fake at 25 slides. What is
untested is not the arithmetic of `start += READBACK_PAGE`; it is what a real
host does on a second page — and the page is 20 precisely to stay clear of the
web >50-item load ceiling (office-js#4272). A page of 3 never approaches it. So
the shortcut exercises the half that is covered and skips the half that is the
risk, at the cost of making a shipped constant mutable in the hot read path.

The honest options are a deck genuinely holding 21+ slides, or leaving this
open. It stays open.

**13, 14 and 15 closed on 2026-09-03.** Each was closed by an answer rather
than by a decision, and two of the answers were the opposite of what the item
assumed.

**13 — "the 4:3 leg crashes on half its attempts".** It is not the profile.
Presentation70 held fixed and run at 960pt crashed 2 of 3 attempts, against
52 of 73 at 720pt on the same deck; rounds 362 (4:3) and 363 (16:9) are
near-identical, 15 passed / 1 failed / 0 skipped each. The archive's
52-of-73-against-0-of-51 contrast was confounded: every 4:3 round ran
Presentation70 and every 16:9 round ran Presentation64, and no file named
either until `driverDeck` landed. Width and deck are near-collinear across
the whole archive. What is wrong with that DOCUMENT is now the open question,
and it is answerable because rounds name theirs.

> **ANSWERED 2026-09-04, and nothing was wrong with the document.** It has TWO
> slide masters where Presentation64 has one, and that was enough: this add-in
> sent `slides.add` a `layoutId` with no `slideMasterId`, which Office.js
> documents as an error unless the layout resolves against the default master.
> One master hid it for months. The deck was only ever the messenger — any
> customer deck with two masters would have done the same. Fixed by sending the
> pair; round 377 then swept 17 of 17 on that document at 4:3. See the last two
> entries in this file.

**14 — "40% of crashes leave no round file; may a stub be archived?"** The
question dissolved: nothing needs fabricating, because the evidence was
already there. `recordCrashFinding("selftest:<name>", r)` has been banking
each verdict as it completes all along — 69 of 77 crash records carry them,
50 of those a complete set, 837 verdicts in total, and `triage` referenced
`crashes/` once. So no stub, and nothing enters `rounds/`: the number IS the
chronology and the filename carries the build SHA. Two things instead —
the driver now waits out the pane's own 45-second deck-evidence budget
before giving a crashed round up (that timeout exists to save exactly these
rounds and had fired in 0 of 77 records, because the driver returned the
instant it saw the dialog), and `salvage-crashed.mjs` went from 22 rounds to
69 once its two soft bars recorded their uncertainty instead of discarding
the round.

**15 — "the mid-draw stop scenario needs one round against a real host".**
It got seven. It failed all seven, and the cause is a change of mine:
`ca138f8` moved it onto a freshly-added slide and it has thrown
`GeneralException` at `SlideCollection.getItem` ever since, after passing
twice on the build before. That is a clean build boundary with one `src/`
commit in it, and the elapsed times overlap through it (a pass at 488s
against a failure at 400s), so it is not the host tiring. It also costs
rounds: 7 of the 76 sound crash records carry a `stop requested` in their
final ten steps and every one is dated 2026-09-02 or later. What it cannot
say is WHICH of its two changes did it — the slide or the stop — which is
why item 19 exists.

**Closed 2026-09-01:** 18 — no ratchet varied `decorations`, and now one does,
swept at the frames where the text has room. The answer is that the engine adds
no overlap there (the total is unchanged at 785) and no shape-budget risk (+18
shapes at worst; the heaviest decorated chart is `area` at 195 against the ~300
where this host dies). The blind spot was real and what it hid was nothing.

**Closed 2026-08-31:** 11 (a European paste read 1000x too small — the owner
called it, and the convention is now inferred per paste and written into the
grid) and 12 (a second action cleared the user's Stop).

Items 3, 5, 13 and 14 are the owner's. Two are decisions about what a user
RECEIVES or about how this project is validated — how a crowded slide is drawn,
and whether the append-only archive may hold a file no pane ever produced; one
spends a cycle leg on an experiment; one goes out under his GitHub identity.
Items 15 to 18 are engineering and NONE of them is a defect: three are places
where an instrument has never been exercised, and one is the queue of things the
next round has to confirm. Items 9, 10, 11 and 12 closed on 2026-08-31 — the
violin costed and thinned, the 1.10 question decided WARN AND PICTURE, the
European paste inferred per block and written into the grid, and the Stop no
longer thrown away by the next click.

Items 1, 2, 4, 6, 7 and 8 were closed by shipping, by measuring the remedy and
declining it, or by the host refusing the mechanism outright — each says which,
in its own row.

**Item 1 did not need the decision it had been waiting for**, which is the
lesson worth keeping from it. It was framed as "where does a unit belong when
the band cannot hold the title, the unit and the top tick" — a question about
placement, with three ways out written up and none of them good. The answer was
that the band's OTHER occupants had already settled it years earlier: on a chart
that cannot hold everything, chrome yields and the title stays. The unit was
simply the one thing in the band that had never been told. See §3.

| | What | Where | State |
|---|---|---|---|
| 1 | ~~**`valueAxisTitle` runs over the legend and the totals row**~~ | §3, *Text drawn over text* | **CLOSED 2026-08-29, and it never was a placement decision.** Two changes: the unit is CLIPPED to 40% of the chart (1,327 → 1,014), and it now YIELDS to the title instead of printing through it (1,014 → **785**). The second killed `title / value-axis-title` — 205 pairs, the only member of the family no width remedy could touch, because both nodes are `align:"left"` at `x:0` so their ink always shares the x range and only the `y` decides. **Measured:** of 176 charts drawing both, 22 overlapped, all at 80x60 and 300x60 at 18pt, with a clean 5.9pt dead zone between the worst clear gap (+3.31pt) and the best overlapping one (−2.56pt). 154 of 176 units survive; 100% at 480x300 and above. Dropping author text is normally refused here — it is right in this one case because the old behaviour did not KEEP the unit, it printed it through the title so the reader lost both, and because on those exact charts the engine already drops the category names, the axis strip and the legend. **And the headline number was never what it looked like**: run the same sweep at the documented `€m` and the family is 118 rather than 687. See the warning now at the top of `test/overlap-budget.test.ts`. (Was "287 of 317" — a figure from the uncrossed sweep, not comparable.) |
| 2 | ~~**The slow-insert offer is not wired**~~ | §1, *The slow-insert offer* | **SHIPPED 2026-08-29.** All five steps, five pane tests, four mutants dead. The blocker this row named was a stale finding: a fresh slide's id IS usable once the slide settles, and the probe answer flipped at round 254 because of our own commit. Kept in §1 for the lesson, not the task. |
| 3 | **Adding to an occupied slide costs ~24s; a fresh one ~0.75s** | §1, *Adding a chart to an existing slide* | Measured over 2,917 batches. Item 2 (shipped) is the user-facing half. **A fourth option is now costed, 2026-08-29**: draw the chart as ONE picture — ~43s becomes rasterise 705ms plus a single-shape insert, a 7–20× saving, and every part of the mechanism already ships. **Still the owner's call**, and a bigger one than the offer: it changes what the user RECEIVES, not just what they are told. |
| 4 | ~~**The picture fast-path**~~ | §1, *The largest product cost was in the FAST path* | **CLOSED 2026-08-29 on evidence, both halves.** The cost it was filed under had moved: redraws are 14.5% and exactly 2 a round, and the real cost was inside the FAST path, which took a median of 15.7s because it wrote every property of every changed node. It now writes only what differs — `same scale across the deck` 146–156s → **103s**, below the IQR of 275 rounds. And the picture feature itself is **not worth building**: a picture redraw is ~40ms of delete plus ~770ms of picture against 15–50s to draw the same shapes natively, so a fast path could save about five percent of an operation that is already twenty to sixty times cheaper than its alternative. |
| 5 | **File this project's host measurements to the office-js tracker** | §1, *Report what this project has measured* | **Owner-gated** — it goes out under his GitHub identity, so nothing is filed without his word on that specific issue. Three are written and ready. |
| 6 | ~~**The dual-axis gutter**~~ — 14 pairs, and the remedy costs 1,394 axis readings | §3, *Text drawn over text* | **MEASURED AND DECLINED 2026-08-29.** Not "30 pairs": 14, of 1,327, and none of them a tick number. Applying the merge to the secondary case fixes 8 and DELETES 1,394 secondary tick numbers — 174 readings lost per pair gained, where the original decision recorded about ten. The trade got seventeen times worse as the rest of the engine improved. The design is written and stays written; on these numbers nobody should build it. |
| 7 | ~~**Positional group-member mapping → `Shape.creationId`**~~ | §3 | **CLOSED 2026-08-29 — the host refuses it.** This host reports PowerPointApi **1.10** and does not populate `Shape.creationId`: `absent` / `no-creation-id` on all three probe questions, 25 of 25 rounds. So it was never a matter of gating on 1.10. The positional mapping stays, still inferred, still guarded by the node-0 anchor test — which is now the permanent answer rather than a stopgap. |
| 8 | ~~**`slideSize()` rung 1 times out in EVERY round that reads a size**~~ | §3 | **DONE 2026-08-29**, and measured on both sides. The bound is now its own constant at 1,500ms rather than the shared selection budget, and rounds 303 and 305 are a clean A/B on the same fallback path: `exportedSlide ms=4404` → `ms=2108`, **2.3 seconds back** per cold insert, with rung 1 still answering in 259ms when it answers. Remaining, and much smaller: whether a warm-up call removes the stall altogether. The measurement that got here: Re-measured 2026-08-29 over 270 rounds: 157 of 157, always the full 4000ms — not "about twice as often as it answers". **And now timed**: a SUCCESSFUL rung-1 read costs 246–270ms (rounds 297–299), so the stall and the answer are different calls and the bound is not buying the answer. **Round 300 then broke the "157 of 157" entirely** — no stall, 138ms, on a round that followed a crash recovery and a full tab reload; round 301, on a tab left idle 35 minutes, stalled again. **And 301 measured rung 2 at ~280ms**, which is what this was blocked on — the export costs what the read costs, so a 500ms bound would take a stalled run from 4,280ms to under 800. No longer blocked on a measurement; only on choosing between a lower bound and a warm-up. |
| 9 | **The violin is the heaviest chart we ship — 259 shapes — and was budgeted at 16** | §1, *A gate is only as true as the quantity it reads* | **CLOSED 2026-08-31 — costed, and it was 246 shapes of tail nobody could see.** Three KDE bodies of 82 points, one line per edge on the Office.js sink. Thinned to a quarter-point tolerance (Douglas-Peucker, so the tolerance is a guarantee): 246 polygon edges to 72, the chart 259 shapes to 97, the deck 6,873 to 6,375. The violin is an ordinary chart again. Originally: **OPEN 2026-08-30, a measurement rather than a defect.** `test/shape-budget.test.ts` read `scene.nodes.length` for its whole life while its title, its 300-shape crash line and its 767ms-per-shape opening were all about SHAPES. Now reads `estimateOfficeShapes`. The violin's 16 nodes are **259 shapes** — a 16× under-count, and the largest single chart in the deck now the hex tile map is 146. Nothing here says 259 is too many; it says nobody has ever looked at it. Cost what it spends before deciding whether it needs anything. |
| 10 | ~~**What a host below PowerPointApi 1.10 loses — 18 of 123 charts, 8 of them wholly**~~ | §1, *Below 1.10* | **DECIDED AND SHIPPED 2026-08-31 — the owner chose WARN AND PICTURE.** A chart carrying marks this host cannot draw now inserts as a PICTURE, complete, with a message naming what would have been missing ('This PowerPoint cannot draw 4 pie slices...'). Where no picture is possible it still draws and still says what is absent. And Explode refuses on the same test: exploding such a chart is the door back to the version that cannot be drawn, which is the mistake the density guard's own comment records the product making once before. Evidence below is what the decision was made on. The renderer cannot rotate a shape below 1.10, so `addWedgeFan` and the arrowhead case both trace and return nothing. **Measured, not estimated:** 18 of the 123 shipped charts lose ink — pie 4/4, doughnut 2/2, sunburst 2/2 (those **8 keep only their labels**, the wedge IS the chart), then gantt 4/7, stacked 3/14, waterfall 1/6, scatter 1/8, radar 1/5, which lose annotation arrows and keep their marks. This supersedes the "9 of 123" figure carried until now, which was an undercount by half. **And the elements are not in that count**: a Harvey ball draws an empty ring at every fraction between 1% and 99% — 0% and 100% are ellipses and survive, so exactly the informative range is lost — and a table `[up]`/`[down]`/`[flat]` cell keeps its text and loses its arrow. A diagonal SEGMENT is fine: it falls back to a real line. The choice is refuse / warn / fall back to a picture, and the pane already has the machinery — the `stalled` path tells a user which charts came out blank. **Not urgent for the web**, which reports 1.10; this is desktop, Mac and every volume-licensed build. |
| 11 | **A European paste is read 1000× too small, silently, in two cells out of three** | §1, *Below 1.10* (neighbouring note) | **DONE 2026-08-31. The call was the owner's — it changes how a user's DATA is read, the highest-stakes call in the product — and the owner made it in one line that day: "continue with the european paste fix."** `numericValue` (src/taskpane/datasheet.ts:72) refuses a comma it cannot place — the file's own rule, stated there, is "a visible gap beats a wrong number" — and applies nothing of the kind to the DOT. Measured on one paste out of a Danish Excel: `1.234` → **1.234** (means 1234), `2.500` → **2.5** (means 2500), `987,5` → `null` (correctly refused). So the same paste silently mis-reads two cells by a factor of a thousand and visibly refuses the third, and a revenue chart reads 1.234 where the sheet said 1,234. A US paste of the same table is read correctly throughout. **Per cell it is genuinely ambiguous** — `1.234` is a legitimate en-US decimal — which is why this is not a bug with an obvious fix. **Across the SHEET it usually is not**: a paste containing any comma-decimal cell is European, and that evidence is already in the same clipboard block, so the proposal is to infer the convention per paste and read dot-groups as thousands when it holds. Not done unilaterally: interpreting numbers differently is what the user RECEIVES. Danish, German, Nordic and most continental Excels are affected, i.e. this project's own author. **What shipped:** `looksEuropean` + `fromEuropeanNumber` (src/taskpane/datasheet.ts), inferring once per paste and rewriting the cells IN the grid rather than at read time — so the user watches `1.234` become `1234` and can correct it if the guess is ever wrong. The inference is deliberately asymmetric: European evidence counts only when NOTHING contradicts it, so a US grouping anywhere, a dot decimal that cannot be a group, or a block with no comma at all all leave the paste untouched. The eager direction would corrupt US data, which is worse than leaving European data as it was. The pane says so through a new `onNote` channel — a neutral hint, not "ok" (a green tick over a reinterpretation reads as confirmation) and not "busy" (which un-hides the progress bar for work already finished) — naming the count so the claim is checkable against what is on screen. Six mutants killed; the whole-paste rules have their own file, `test/european-numbers.test.ts`. |
| 12 | **A second action clears the Stop the user pressed, and the cancelled insert completes** | §1, *The pane's action state* | **DONE 2026-08-31 in `bf92e5a` — this row was left saying OPEN for a day after the fix landed, which is its own small lesson: a backlog that lags the code stops being readable as a statement of what is true. The fix: `guard()` counts actions in flight (`actionsInFlight`) and only the OUTERMOST one may reset the stop flag or take down the Stop button, so a second click can no longer clear the first click's decision. The count is read AFTER the decrement, not captured on entry — the first version captured ownership on entry and got it wrong whenever the outer action finished first.** `guard()` calls `resetStop()` on entry (app.ts:3920) and again on exit, and it leaves ~13 controls live during a slow insert. So: start an insert the code itself prices at "the better part of a minute", press **Stop** — the button disables and reads "Stopping…" — then touch any other live control, and the insert resumes with `isStopRequested()` false and **writes its chart onto the slide**. Reproduced against the real `src/taskpane/app.ts` under jsdom with the REAL stop functions (only host calls doubled), with a control that shows the stop honoured when nothing else is clicked: `stopSeen: [false] landed: 1` against `stopSeen: [true] landed: 0`. **The same `finally` also lies about the outcome**: it posts "Done." in green, hides Stop and clears the elapsed interval while the first action is still in flight — and that interval is the sole carrier of the `SILENT_RUN_MS` watchdog, so the 60s "PowerPoint died and the pane did not" check is disarmed for the rest of the action. Measured: elapsed reads `""` 2.3s in versus `"2s"` in the control, i.e. the timer is gone, not blanked. The fix is per-action stop/elapsed state rather than one module-level flag; it wants care, not speed. |
| 13 | **The 4:3 validation leg crashes PowerPoint on half its attempts — and it is the PROFILE, not the deck** | §3, *The crash, counted across the whole archive* | **THE CONFOUND IS BROKEN, 2026-09-02.** Aspect ratio and deck file had never been crossed: every 4:3 round in the archive ran against `Presentation70`, so "4:3 is dangerous" and "that deck is worn out after two dozen rounds" fitted the evidence equally. The owner supplied a 4:3 deck; the file could not be uploaded into the session (a file chooser does not survive between two `playwright-cli` processes), so a blank one was CREATED in OneDrive instead — `Presentation71` — and set to 4:3 with the driver's own `PW_SET_SIZE`, which exists for exactly that. Four standalone 4:3 rounds followed. **A brand-new, empty deck crashes at the same rate as the battered one: 4 crashes in 11 attempts (36.4%) against Presentation70's 47 in 94 (50.0%), and 16:9's 9 in 380 (2.4%).** Eleven attempts is a small n and the two 4:3 figures should not be read as distinguishable from each other — but against 2.4% they are overwhelming: four crashes in eleven attempts at the 16:9 rate is a one-in-a-million event. **So the deck is exonerated and the 720x540 profile is not.** **THREE MORE CONFOUNDS FELL WITH IT.** Deck CONTENT: the new deck was empty, so accumulated slides are not it. LEG POSITION: every previous 4:3 round was leg 3 of a cycle, and these four ran as their own session at indices 1-4, so "the third leg of a tired session" is not it either. DURATION: 4:3 rounds average 454s of scenario time against 16:9's 442s — a 3% difference carrying a 20x difference in crashes, so "the 4:3 leg simply runs longer and meets more chances" is dead. What remains unseparated is 4:3 from anything else that travels with a 720x540 slide on this host; the claim here is only that the profile, not the deck instance, the content, the position or the length, is what the rate follows. Next step is a mechanism, not another rate: the crash reports carry no ErrorName in 67 of 77 cases, so instrumenting WHERE in a round the 4:3 leg dies is the question worth asking now. Originally: **OPEN 2026-08-31, and the next step is the owner's because it spends rounds.** Measured over all 322 archived rounds, with ATTEMPTS as the denominator — the figure this file itself said nobody had computed: **4:3 produced 40 crash events in 81 attempts (49.4%); 16:9 produced 9 in 365 (2.5%).** 17 of 21 4:3 rounds hit at least one crash, against 9 of 173 at 16:9. Attempts is the right denominator and rounds is not: a round that crashed three times before landing is one round and four attempts, and counting rounds hides exactly the arm that is failing. **Session position does not explain it.** Recomputed from `driverRun.startedAt` with the driver's own 45-minute rule — rather than from the archived `sessionIndex`, which was reset by the cycle for every round before `fd970c9` — 4:3 runs 47-51% at every session position while 16:9 runs 6.5% at the first and **0% at every later one**. **What it does NOT say is which variable owns the number.** `cyclePlan` (scripts/cycle.mjs:63-65) pairs 4:3 with the tall deck and 16:9 with the wide one and **has never crossed them**, so aspect ratio and deck file are perfectly confounded across the entire archive. The mechanism is also unestablished: `errorLocalChangeLostSingleUser` appears in a minority of crash reports and most carry no ErrorName at all. **The experiment**, one round long: run a leg with the WIDE deck resized to 4:3 in place. If it crashes, the aspect ratio owns it; if it does not, the deck file does. Owner's call because it edits the validation plan and burns a cycle leg, and because a `wrong-size` mismatch stops the whole cycle — the comment above `cyclePlan` records that this has already happened once. **Done meanwhile:** the rate is now printed by `rounds-gate.mjs` on every run, worst arm first, with the confound named on the line beneath it, so nobody has to rediscover the division. |
| 14 | **40% of this project's crashes left no round file, so every crash rate in the archive is a floor** | §3, *The crash, counted across the whole archive* | **OPEN 2026-09-01, and the remaining step is the owner's because it writes into the append-only archive.** A round the driver never recovers archives nothing — `scripts/round.mjs` states the mechanism outright: *"A crashed round archives nothing: it never reaches the download button."* So `rounds/` holds only the crashes that were SURVIVED. Measured 2026-09-01: **77 crash reports in `crashes/` against 46 crash events in archived `driverRun.recovered` — 40% left no round file at all**, worst on 2026-08-29 (22 reports, 6 archived). Every crash rate computed from the round files, including item 13's 49.4% for the 4:3 leg, is therefore conditional on recovery succeeding and is a FLOOR rather than a rate. It is probably a floor that understates 4:3 specifically: `rounds-salvaged/` holds 22 hand-recovered rounds and all of them are 4:3, which is the same gap being closed by hand. **Done meanwhile:** `rounds-gate.mjs` now prints the two counts side by side and labels the rate a floor, so the number is never read as complete. **What is left, and why it is not mine:** the driver could write a stub round file (or `NNN-<sha>.aborted.json`) when `--retry` is exhausted, carrying `driverRun` and the final reason. That would let the archive tell *"this build did not crash"* from *"this build crashed so hard nothing was filed"* — the verified-versus-not-attempted distinction this project keeps rediscovering. But `rounds/` is append-only and every triage and gate tool reads it as the record of what happened; putting a file in there that no pane ever produced changes what the archive MEANS, and every count computed from it. That is a decision about the instrument, not a bug fix. |
| 15 | **Seven of the fifteen scenarios have never failed in 322 rounds, and the one that never fails hardest passes on a path it does not claim** | §1, *The pane's action state* | **THE STOP HALF IS BUILT 2026-09-01 and needs one round to confirm it on a real host.** `stop a run mid-draw` requests the stop from the first `onPhase("commit", …)`, so a batch of shapes has landed when the abort happens — the case a user actually meets, and the one item 12's fix was written for, exercised ZERO times in 322 rounds. A NEW scenario rather than a change to `stop a run part-way`, for the reason that one's own comment gives: a mid-draw abort leaves a partly drawn slide, and `same scale across the deck` discovers its chart population from the same deck. So it cleans up after itself — records the slide's shape ids before the insert, deletes what the aborted draw added, and FAILS if it cannot. What it asserts is the contract the pane already states in words: partial shapes are kept, but half a chart must not carry a config tag and claim to be whole. **It has a ceiling on what it may delete**, because this host is documented to renumber shape ids: an id diff larger than the chart itself means the ids moved, not that the draw added them, and deleting on that reading would take the probe chart with it. **What is left:** it has never run against PowerPoint. Until it has, this row stays open — a scenario that draws and deletes on a live deck is exactly the kind of thing that must be watched once before it is trusted. **The other half of this row is already answered elsewhere:** a 14-of-15 scenario block says almost nothing, and the discriminating content moved into `hostAnswers`, which `rounds-gate.mjs` now reads. Originally: **OPEN 2026-09-01, engineering, and the larger half is already done.** Lifetime pass/fail/skip over 322 rounds: `stop a run part-way` 322/0/0, `which selection call wedges the host` 318/0/4, `edit the chart the user selected` 321/0/1, `does a rasterise poison the next draw` 321/0/1, `a selected shape survives an insert` 309/0/13, `the chart is actually visible` 303/0/19, `one chart alone on a warm deck` 133/0/1. In the last 24 hours scenarios 0-13 passed in all 13 rounds and only the rotation scenario varied at all. **A 14-of-14 is no longer evidence that a round measured anything**, which matters because the loop reads it as though it were. That half is addressed: the discriminating content has moved into `hostAnswers`, and `rounds-gate.mjs` now reads it — 163 of 2,542 probe slots differ between two runs of the SAME commit, which is where a round's signal actually lives. **What is still open** is `stop a run part-way`, the sharpest of the seven because it passes on a path it does not claim: its detail reads *"stopped before the first batch — no shape was ever queued"* in **251 of 322 rounds** and *"stopped at a batch boundary"* in the other 71. The mid-flight stop — the case a user actually hits, and the one item 12's fix exists for — has been exercised **zero** times on a real host. The scenario's own comment says so ("the mid-flight promise has never been exercised") and names the seam: `onPhase("commit", …)` in the same function. It wants a stop timed from that seam rather than from the wall clock. Not urgent; it is simply the scenario whose green says least. |
| 16 | ~~**What the next round has to confirm, and the fixture refresh that is waiting on it**~~ | §1, *The answer sheet* | **CLOSED 2026-09-01 after six rounds (347-352) across two cycles.** The fixture is refreshed from round 351 — a POST-fix capture, which is what this row said to wait for — and **PENDING_QUESTIONS is now EMPTY**, which is the register working rather than being abandoned. **The refresh got cheaper because of the week's own work:** from round 346 it would have brought SIX undeclared divergences; from 351 it brought THREE. `binding-names-shape-later`, `shape-add-fresh-getitem-slide` and `which-end-a-short-read-drops` stopped diverging on their own once `unreadable`, `silent` and a probe's own `all` stopped locking rows against the answers underneath them. The three that remain are declared with their evidence. **`named-preset-resolves` retired** the way the register says to — `draws` in 14 of 14 rounds, 42 of 42 samples — and lost its `resample` mark, because asking a settled question three times a round buys certainty already bought. **`rotation-keeps-the-unrotated-box` did NOT get its reading** — `unreadable` on all 18 passes across six rounds — so rather than leave the register waiting on a 1-in-9 event indefinitely, it moved to KNOWN_DIVERGENCES *and* UNSTABLE_ANSWERS, keeping its `resample` mark. That is the honest filing: the host answers no geometry for a just-added shape, the four readings it has ever managed all agree with the fake, and nothing has ever come back saying the box is the post-rotation one. If a future capture reads `rotated-box`, that entry is the first thing to revisit. **Also fixed along the way:** the "declares nothing that has stopped diverging" gate could not tell a fixed fake from a coin landing heads, and demanded the deletion of `how-many-syncs-a-creation-handle-survives` — declared unstable at 23-of-25 — because this capture caught its rare face. An id in UNSTABLE_ANSWERS is exempt now. Originally: **THREE OF FOUR CONFIRMED by rounds 347-349 on 2026-09-01; one still needs a round.** **(b) CONFIRMED** — the host answers `items/width,items/height`: 20, 14 and 15 shapes across the three rounds, every one sized. That was the week's only change to a shipped code path, and `deckGeometryFaults` now reads it. **(c) CONFIRMED** — `which-end-a-short-read-drops` reads `not-a-short-read`. **(e) CONFIRMED** — `named-preset-resolves` still `draws`, now in 14 of 14 rounds. **(a) NOT YET, and not refuted either**: `rotation-keeps-the-unrotated-box` answered `unreadable` on all NINE passes across the three rounds. At its historical read rate of 4 in 33 that is unremarkable — roughly a one-in-three chance of nine blanks — and there was simply nothing for the fix to promote. The recipe stands: expect `unrotated-box` in any round where a pass reads. **(d) and (f) still wait on (a)**, since refreshing the fixture before the rotation probe has been seen post-fix would bake in the answer this week's change exists to correct. Originally: **OPEN 2026-09-01, engineering, and deliberately not done on 2026-09-01.** Five changes shipped that week alter what a probe sheet SAYS, and none of them has yet been seen on a real host. Until a round runs on a build at or after `30d28f9` this is a queue, not a task list. **(a)** `rotation-keeps-the-unrotated-box` must report **`unrotated-box`** in any round where a pass reads at all, instead of `unreadable`; `stable: true` for a 339-shaped round (two strong samples agreeing), and `stable: undefined` — NOT false — for a 344/345-shaped one. If that does not happen, the `UNINFORMATIVE` fix did not land and the whole answer-ranking pass needs re-reading. **(b)** the deck inventory now asks for `items/width,items/height`; confirm this host answers them at all. If it refuses, the inventory degrades to exactly what it recorded before, which is why each is read through `loadedValue` — but the refusal is itself the finding and belongs in `KNOWN_DIVERGENCES`. **(c)** `which-end-a-short-read-drops` should read `not-a-short-read`, and `RENAMED_ANSWERS` can be dropped once no committed or archived sheet still carries `all` for it. **(d)** then, and ONLY then, refresh `test/fixtures/host-answers-web.json`. It is 53 rounds stale (build `8d8267f`, 2026-08-28) and the ritual at the top of `test/host-contract.test.ts` is "replace the fixture, run this, and deal with what goes red". Refreshing from round 346 was tried and rejected: it imports **six** undeclared divergences — `binding-names-shape-later`, `shape-add-fresh-getitem-slide`, `shape-proxy-survives-one-sync`, `shapes-by-index-vs-items`, `rotation-keeps-the-unrotated-box`, `which-end-a-short-read-drops` — and five of those are answers this week's fixes change, so declaring them against a pre-fix capture would be wrong within one round. **(e)** with the fixture refreshed, delete `named-preset-resolves` from `PENDING_QUESTIONS`: it has answered `draws` in 11 of 11 rounds, 33 of 33 samples, across five builds. `rounds-gate.mjs` prints this every run now, so it cannot be forgotten, but it cannot be closed before (d). **(f)** re-derive `UNSTABLE_ANSWERS` from `probeFlipsWithinBuild` rather than by hand — but only after a few post-fix rounds, because some of today's 21 flipping probes were flipping BECAUSE a weak answer locked a row, and those should stop. |
| 17 | **The deck scan's paging loop has never taken a second iteration on a real host** | §3, *The instruments were the problem* | **HALF DONE 2026-09-01; what remains is the live host, and this row overstated the gap when it was written.** Our own logic was already covered for the OTHER paged readers — `snapshotAddedSlides` and `readAddedSlides` are tested at `READBACK_PAGE + 5` slides including a refused second page. The deck scan was not: `listChartsInDeck`, which Same Scale, the repair pass and five scenarios all read through, appears in 44 tests in `office-render.test.ts` and every one installed a handful of slides, so `start += READBACK_PAGE` never came round a second time even under the fake. Now tested at 25 slides, asserting a shape on the SECOND page reaches the inventory, and that a lost page is reported as `unread` rather than silently returning page one — because a scan that loses a page and does not say so has Same Scale rescale a subset of the deck and report success. Two mutants killed: reading only the first page, and not counting a lost page as unread. **Still open:** no REAL host has been asked to list more than nine slides. Originally: **OPEN 2026-09-01, engineering.** `READBACK_PAGE` is 20 and the deepest deck ever scanned in 322 archived rounds is **9 slides** — 305 of them scanned exactly 7. So `for (let start = 0; start < total; start += READBACK_PAGE)` has executed its second pass **zero** times against PowerPoint, and the page boundary is where a scan silently returns a prefix. Same Scale, the repair pass and five self-test scenarios all read the deck through it. A user pressing **Same scale** on a 40-slide deck — an ordinary corporate deck — is the first person to run that code on a real host. The fix is not to the loop, which looks right; it is a scenario, or a probe, that puts more than 20 slides in front of it once. Cheap to state and awkward to stage, which is presumably why it has waited. |
| 18 | ~~**No ratchet varies `decorations`, and decorations are what generate the text the overlap ratchet measures**~~ | §3, *Text drawn over text* | **CLOSED 2026-09-01, the same day it was filed, and the answer is that there was nothing behind it.** Four decoration bundles now sweep alongside every other option — merged onto each kind's own defaults rather than replacing them, which a shallow spread would have done while appearing to widen the sweep. Swept at the frames where the text has room, and there the total does not move: **785, unchanged**. The variants are not idle: each builds 100 charts with zero refusals and adds up to +232 text nodes. Nor is there a shape-budget risk, the other thing more labels could have cost — across 25 kinds at a full-slide frame full decoration is **+18 shapes at worst**, and the heaviest decorated chart is `area` at 195 against the ~300 where this host starts dying. Zero kinds cross it. **Why the frame restriction, since it is the one variant family that has one:** turning everything on at every frame counts 2,748 against 785, and the split by frame says what that is — 80x60 97, 60x300 95, 120x90 86, 160x120 61, 200x150 42, 300x60 20, 480x300 **2**, 960x540 **0**. An 80x60pt box asked to label twenty-four categories cannot place them anywhere; recording ~2,000 impossible-frame collisions would also blunt the ratchet, hiding a real regression of ten inside a tolerated bucket of 1,243. Originally: **OPEN 2026-09-01, engineering.** `overlap-budget`, `shape-budget` and `showcase-overlap` contain **zero** references to `decorations` between them. The sweep varies chart kind, frame size, font size, and its cross-product of options and shapes — and never the titles, segment labels, series labels, column totals, grand totals, category labels, value axis, gridlines, connector lines or the `100% =` note. Those are precisely the nodes that collide: every overlap defect this project has fixed was a label against another label. So the budget that guards text-over-text is measured on charts drawn with their text mostly turned off, and the number it defends is a number for the quiet case. Worth its own measured pass rather than being bolted on: turning every decoration on across the existing sweep will move all three ratchets at once, and the honest order is to measure first, decide what is a defect second, and edit the numbers down third. |

Two standing costs that are not tasks: the host crashes during evidence
collection after a long round (§3 — it costs the evidence, not the run), and
`test/overlap-budget.test.ts` holds **1,327** overlapping pairs whose per-shape
table is the live list. Three figures in one day, and only the last step was the
engine: **537** was the uncrossed sweep, **4,010** the same engine measured by a
sweep that finally crossed option variants with data shapes, and **1,327** what
was left once `buildMultiples` stopped building grids with no room. Do not
compare any of them with an earlier number without naming the sweep.

**Found and fixed 2026-08-29: `buildMultiples` built grids with no room.** A
panel height of −0.4 was handed on, `clampDim` rewrote it to
`DEFAULT_SIZE.height`, and ten full 300-point charts were stacked 9.6 points
apart inside a box 60 points tall. 2,683 of the sweep's pairs; all of them gone,
total now **1,327**. **Validated**: rounds 298 and 299 at 16:9 and 300 at 4:3,
all 14 of 14, plus a byte-identical showcase across 123 configs — see §3.

**Newly open, 2026-08-29: `duplicateSlot` judges a window it cannot fill.** The
scenario sizes its reconcile from a slide count this host reports LATE, so a
round where an insert answers "nothing landed" for slides that did land gets a
hard FAILED verdict about two slides out of four — and the cycle's fatal check
stops the night on it. Round 297 is the worked example; the fix is a few lines
and is deliberately not taken at 2am, because it changes a verdict rule with a
273-round series behind it. See §3.

**Newly reopened, 2026-08-29: PowerPointApi 1.8 bindings.** §2 retires them as
"unanswerable on this host", on the strength of the host rejecting the batch
that carries a binding. Over the 25 rounds to 2026-08-28 that face appears
ONCE, against `yes` 15 times and `silent` 8. It is the last untried way out of
the id refusals and the evidence that closed it has weakened. Not put on the
list above because nothing has been designed — the `silent` third is its own
problem — but it is no longer refuted, and `binding-names-shape-later` is back
on the probe's resample shortlist so the next rounds keep counting.

### A doomed round trip on every grouped-chart update — bisected 2026-08-30

**Open, small, and NOT a regression — which took a second look to establish.**

Three errors appear in every recent round, exactly once each. All three come from
one operation, `explode a degraded picture`, which is also the only scenario
carrying real friction (4.0 events a round against 0.0 everywhere else) and was
the worst in the archive before it was repaired:

    reading back an ungrouped chart's shape ids  :: InvalidParam passed to GetItem(id)
    resolving the charts' shapes                 :: GeneralException
    a by-id lookup refused the whole resolve — re-reading the slides instead

The code around the fallback (`powerpoint.ts:2604-2640`) says "Costs nothing on a
healthy host: no refusal, nothing unanswered, no read." On this path that premise
is false — the refusal is not the exceptional case, it is the only case — so the
engine spends a guaranteed-failed sync and then does the collection read it was
always going to need.

**BISECTED FROM THE ARCHIVE, which is what it was kept for.** Refusal rate by
round, and the flip is one build wide:

    rounds 0-99      0%        round 142  bcfc126   ok
    rounds 100-149  14%        round 143  408d00b   REFUSES
    rounds 150-349  92-100%    ...and never recovers

`408d00b` is *"Teach the in-place update to read a group's members (#642)"*. Its
own comment says that before it, the update **refused every grouped chart** for
want of a parts list — "18 of 21 in round 142, and the feature had never once run
in 117 rounds."

So nothing broke. A feature that never ran started running, and its by-id lookup
refuses every time on this host. The 123 rounds that show no refusal are the
rounds where the code never tried. Reading the 0% as "it used to work" would have
been exactly backwards, and the raw rate invites that reading.

**What is worth doing**, when there is a cycle to validate it against: on the
grouped-member path only, skip the by-id resolve and go straight to the re-read
that already works. Never `resolveCharts` generally — the other 11.8 updates a
round never hit this and by-id is fine for them.

**Why it was not done on 2026-08-30 despite being planned.** It touches the
hottest path in the product — the in-place update is 108 seconds and a quarter of
every round — and the verification that justified it also showed the change could
not be checked against a live host in the time left. A saved round trip is not
worth an unvalidated edit to that function.

**RE-DERIVED AND STILL LIVE, 2026-09-07.** Counted independently from the whole
archive: **274 of the 282 rounds since 143 carry it, and a re-read followed all
274 times.** The eight that do not are rounds that ended before the path ran. So
the cost is unchanged and the repair has never once failed — which is why this
has stayed cheap to ignore.

> **"EXACTLY TWICE EACH" STOOD HERE AND WAS MY OWN MISCOUNT**, corrected the
> same evening. Two trace ENTRIES, one event: the error line `resolving the
> charts' shapes` and its recovery line `a by-id lookup refused the whole
> resolve`. Rounds 426 and 427 carry one refusal apiece. Anyone re-deriving this
> should count the recovery line, which fires once per refusal.

The same pass killed a lookalike. `drawing the chart's shapes` also throws
GeneralException, 15 rounds' worth, `errorLocation` `SlideCollection.getItem`
or `getItemAt` — but **every one of them is in rounds 356-373 and there has not
been one since**. Anyone grepping the archive for GeneralException meets both
and should not merge them: one is a live per-round cost, the other is a closed
era.

**THE MEASUREMENT THIS ENTRY WAS WAITING FOR NOW EXISTS**, and no restructuring
came with it. `powerpoint.ts` traces `parts lists at the update` once per batch,
before the sync that may refuse it: `charts`, `withParts`, `queuedGroup`. The
question is the one the code comment poses — a chart that arrives carrying parts
never queues `.group`, so if the parts-list repair lands, this cost may remove
itself and the edit to the hot path is never needed.

Why a new counter rather than the one that was already there: `churn.withParts`
is incremented in the REDRAW loop, so it only ever sees charts whose in-place
update was **already refused** — a pool that excludes its own successes. Its 0
was read as "no chart ever arrives with a list" until 2026-09-07, when
`gotPartsList` turned out to be 35. `queuedGroup` is recorded rather than
subtracted because `queueGroupMembers` also declines below PowerPointApi 1.8 and
when the refusal is already latched for the batch, so it is not
`charts - withParts` and a version that computed it would be wrong on both.

**ANSWERED BY ROUNDS 426 AND 427, both all-green on `630e4b3`.** The question
was whether a chart ever reaches the update carrying a parts list, because one
that does never queues `.group`:

    charts at the update          28   (14 a round, one chart per batch)
    withParts                      0   every one
    queuedGroup                   28   every one

**So the parts-list repair will NOT remove this by itself.** The list is absent
at this point, every chart therefore asks for `.group`, and the hypothesis this
entry was holding the fix for is dead. Note what the number is not: it says the
list is absent HERE, not that it is never written — `gotPartsList` is 35 on the
production side, and the two can both be true if the list does not survive onto
the target.

**AND THE COST IS AN ORDER OF MAGNITUDE SMALLER THAN THIS SECTION'S TITLE.**
"A doomed round trip on every grouped-chart update" is wrong. Charted per chart
across rounds 426, 427, 428 and 429 — the same in all four:

    13 of 14   updated IN PLACE, no redraw, `.group` answered
     1 of 14   refused — `draw-6`, every round, deterministic

`draw-6` is inside `explode a degraded picture`, and **its old shape is a
picture**: `removed: 1`, one shape, not a group. So the mechanism is the one the
code comment always named — `.group` on a shape that is not a group returns
GeneralException and poisons the sync — and it fires on exactly the one chart in
the suite whose old shape is not a group. The other thirteen are groups, their
`.group` reads are honoured, and they never reach a redraw at all.

**And the fix would not save the redraw.** That is the part worth knowing before
anyone spends the hot path on it. After the refusal, draw-6 redraws 24 shapes
because "the chart has no parts list and no readable group members". A picture
has no group members whether or not the sync was poisoned, so it would redraw
anyway — exploding a picture to native shapes IS a redraw. What the fix saves is
one poisoned sync and one recovery re-read, once a round. Real, and small.

The shape of it: not "skip the by-id resolve on the grouped-member path", but
"do not ask a shape that is not a group for its group". `EditTarget` carries no
flag for that, and `old.load("left,top,type")` cannot answer in time — the sync
that would tell us is the one `.group` has already poisoned. The explode path,
though, KNOWS it is exploding a picture. That is where a flag would come from,
and it is a caller-side change rather than a batch-shape one.

### The largest product cost was in the FAST path, not the redraws — 2026-08-29

**Looking for the picture feature below, I measured the redraws and found they
had stopped being the cost.** Over the last 30 rounds:

    in-place 341 · redrawn 58 = 14.5% redrawn   (was 20.9%)
    29  this update draws a picture          — legitimate, cannot be diffed
    29  no parts list and no group members   — the addressable half

Exactly two redraws per round, every round. But the same query printed the
number nobody had looked at:

    in-place update ms: n=341 median 15,661 · max 59,000

**The fast path was taking sixteen seconds.** `same scale across the deck` runs
entirely through it and costs 166s, 39% of the battery, which is what "the
largest product cost" was actually made of.

**What it was spending it on.** `applyNodeInPlace` wrote every property of every
changed node unconditionally. Counted by a recording proxy, a whole text node is
exactly **20 statements** and a rect **7** — 20 being the figure `powerpoint.ts`
already quoted from the host's own statement list, which is how the count is
known to be measuring the right thing. So a retitle, the single edit this path
exists for, sent twenty statements to change one string.

`planSceneUpdate` holds BOTH scenes and therefore already knew which of the
twenty were stale. It just never said. It now returns `parts` beside `changed`
— box, fill, line, text, font, align — and the applier writes only those.

**Three measurements, all agreeing.** Statements first:

    retitle                 20 ->   2   (90% fewer)
    recolour               152 ->  44   (71% fewer)
    rescale, same scale    272 -> 180   (34% fewer)

Then per-update milliseconds on the live host, like-for-like by node count,
against the five preceding builds:

    changed=1     ~2,100-2,500ms  ->   1,117ms
    changed=9    ~11,100-12,100ms ->   4,825ms
    changed=18   ~15,700-16,600ms ->  12,042ms

Then the scenario itself:

    299 4275306  147s      302 aa459dd  156s
    301 e0b82ef  146s      303 81ab4a4  103s

**103s against a pooled median of 166s and an IQR of 151-211s over 275 rounds**
— below the whole interquartile range of the archive.

Time falls by less than statements do, consistently, and that is the expected
shape: `syncMs` shows a fixed floor of roughly a second on every sync, so a
retitle that sends a tenth of the statements still pays two syncs.

**One round for the scenario figure**, n=12 for the per-update figures. Three
independent size classes all moved the same way by a lot, and the baseline is
eight builds deep, but it is one round and should be read as one.

**What is left of the original entry** is the picture feature — and rounds 304
and 305 close it. **It is not worth building.**

The two halves of a picture redraw were timed on 2026-08-29 and the split is
lopsided:

    deleting the chart being replaced    13-38ms    (removed=1)
    drawing the chart as one picture     742-799ms  (nodes=24)

About **810ms in total**, against 15 to 50 seconds for the same twenty-four
shapes drawn one at a time — three batches at the 5,013 to 16,579ms this archive
measures per batch. So the picture path is already twenty to sixty times cheaper
than the thing it replaces.

A "picture fast path" could at best reuse one existing shape as the host and
save the delete: **40ms of 810, about five percent**, for a feature that has to
teach a differ about a fill it cannot see. The argument for expecting that was
written down before the measurement and the measurement agrees with it, which is
the only reason to trust either.

The `removed=1` is why the delete is so cheap and is worth keeping in mind: a
grouped chart is one shape to delete however many it holds. A chart that failed
to group would cost twenty-four deletes, and this host groups reliably now.

**So item 4 is closed on evidence rather than deferred.** The redraws that
remain are correct, cheap, and not worth removing.

### Two of every three redraws are a grouped chart the update cannot map — 2026-08-26

**CLOSED 2026-08-27. The redraw rate is at its design floor — both remaining
redraws are correct, and neither is a missed optimisation.**

Rounds 275 and 276 each redraw exactly twice, and the two reasons are these:

    this update draws a picture, which is not in the scene the differ compares
    the chart has no parts list and no readable group members

The FIRST is a deliberate refusal and removing it would reintroduce a silent
wrong answer. `render: "image"` does not produce a picture — the renderer takes
that path only when handed `pictureBase64` — so collapsing a chart to a picture
builds the SAME scene it already has. The differ compared 24 nodes to 24
identical nodes, said "nothing changed", wrote nothing and reported success,
while the slide kept its 24 native shapes. Worse, the auto-picture fallback is
what the add-in reaches for when the host has ALREADY failed to draw shapes, so
the one path that rescues a struggling host was the one being skipped. Teaching
the fast path to handle it is a real feature — it writes a closed set of `rect`
and `text` properties and a picture fill is neither — not a small fix.

The SECOND is settled: see "ANSWERED 2026-08-27" below. Both routes into a
refused group are shut inside the batch, the redraw is the right answer, and what
rescues the group is the next batch.

So the remaining work here is not "stop redrawing". It is either the picture
feature above, or nothing. Leaving the original analysis below because the
reasoning that got here is worth more than the conclusion.

**The largest product cost still on the table, and the obvious fix is wrong.**

An update either writes only the shapes that changed, or redraws the chart
whole. A redraw is far more expensive: it is why `same scale across the deck`
costs 167s, 38% of the battery. Measured over the last 30 rounds:

    in-place 325 · redrawn 86 = 20.9% redrawn

    56  the chart has no parts list AND no readable group members
    30  this update draws a picture, which is not in the scene the differ compares

The picture case is legitimate — a picture cannot be diffed against a scene. The
other 56 are the whole of the addressable cost, and they are all the same shape:
a chart that GROUPED, whose group the update could not read — not because the
host refuses, as this said until 2026-08-26, but because a session-wide latch had
turned the read off after one refusal.

**How it got here is worth stating, because each step was right.** Grouping used
to fail, so charts were loose, so they carried a parts list and the update mapped
nodes to shapes through it. Grouping now succeeds — `fresh-slides-group` reads
96/98 — and a grouped chart is not "loose", so it collects no parts list. The
update was given a different route for those: `shape.group.shapes` (#643),
written through in #646. That route is what fails:
`group-children-via-getcount` answers `unreadable` in 34 of 40 rounds.

So a grouped chart has neither route, and redraws.

**`grouped-child-by-id-from-slide` is no longer moot.** It was retired
2026-08-21 with the reasoning that the by-id route "is not used and does not need
an answer". The route that made it moot is the route that is failing. A question
is only moot while the thing that made it moot works.

**The naive fix was tried on 2026-08-26 and REVERTED.** Collecting a parts list
for grouped charts as well as loose ones looks free — the ids are already
populated by the drawing batch, so it reads a property rather than making a call
— and the suite killed it in one run:

    reports no growth for a GROUPED chart — "a group goes in one delete": expected 14 to be 1

The parts tag is not only the in-place MAPPING, it is also the DELETE list. Hand
one to a grouped chart and a single group delete becomes fourteen shape deletes,
which is the "chart grows by a whole chart on every edit" bug the tag exists to
prevent. `loose` gated it for a reason that is not written down anywhere.

### ANSWERED 2026-08-26, and the fix is dead

`grouped-child-by-id-from-slide` finally has an answer. It starved for 125
rounds, was retired as moot, and was re-opened above because the route that made
it moot is the route that fails. Asked directly, three times, in under two
seconds each:

    grouped-child-by-id — no-such-shape: child 2 of group 5 read back as undefined
    grouped-child-by-id — no-such-shape: child 5 of group 9 read back as undefined
    grouped-child-by-id — no-such-shape: child 9 of group 13 read back as undefined

**A grouped child is not addressable by id off the slide on this host.** Not
"unreadable", not refused — the slide's shape collection does not contain it at
all. Grouping takes the children out of the slide's collection, and the way back
is `shape.group.shapes` — which DOES enumerate on a real slide, `yes` both fresh
and re-resolved. Only the by-id route is closed.

So the mapping-tag idea is DEAD, and so is anything else built on reaching a
grouped chart's shapes BY ID. Capturing child ids at draw time would have
produced a tag full of ids that resolve to nothing.

### CORRECTED THE SAME DAY — the redraw was NOT unavoidable

The paragraph that stood here said "the redraw is not a missed optimisation, it
is the only correct behaviour available for a grouped chart on this host". That
was wrong, and it was wrong for the reason everything else on this page was
wrong: it rested on a probe answer measured on the scratch slide.

Two routes to a grouped chart's shapes exist and they are not the same question:

    by id            `slide.shapes.getItemOrNullObject(childId)`   — DOES NOT WORK
    by enumeration   `shape.group.shapes` loaded for `items/id`    — WORKS

The first is settled above, three runs, `no-such-shape`. The second was asked on
a real slide the same afternoon and answered `yes` both fresh and re-resolved by
id — where the probe's scratch-slide version, `group-children-via-getcount`,
answers `unreadable` in 34 of 40 rounds.

So the update path had a working route all along. What was taking it away was a
LATCH: `groupReadRefused` held for the whole SESSION, and one refusal ended
in-place updates for the rest of the round. Across rounds 254-261, in-place
updates after the first refusal: zero, eight rounds running. Resetting it per
batch (`55630e7`) gives:

    baseline 255-261   11 in-place / 3 redrawn = 21.4%   seven identical rounds
    with fix 262-264   12 in-place / 2 redrawn = 14.3%   three identical rounds

`a-group-refusal-no-longer-ends-the-round` watches it at 3/3.

**ANSWERED 2026-08-26, and the answer is no.** The latch is deliberately kept
WITHIN a batch, because the refusal arrives at the sync rather than at the
property access and poisons the batch it lands in. So charts after the refusal in
the SAME batch still redraw, which is why the gain is one chart rather than
three.

That is not a limitation to be relaxed. Rounds 265-267 spent three attempts
re-asking inside the batch and produced no in-place update at any of them. Both
routes are shut: the recovered proxy's `.group.shapes` answers and then its sync
throws (round 267 in production), and a fresh by-id proxy fails because the by-id
LOOKUP is what poisons these syncs — the same refusal that sent the chart down
this path. The retry has been deleted.

One premise of the paragraph this replaces was also wrong: **the fake CAN model a
sync-time refusal**, and does — `faults.refuseShapeById` sets `refuseThisSync` at
`test/helpers/office-host.ts:1722`, and `pendingHostError` throws once and clears
rather than poisoning everything after it. Two of the three rounds were spent
believing the suite could not be asked.


**How it was answered matters as much as the answer.** Not by a round — by
`scripts/experiment.mjs`, in under two seconds, four iterations apart. The first
three attempts were wrong in ways this archive has documented for months: asking
a creation proxy for its id, reading it without a load, and asking on a SCRATCH
slide whose shape collection this host answers `unreadable` for in 92% of
rounds. Each mistake cost a minute instead of a round.

### Adding a chart to an existing slide costs ~24s; making a new one costs ~0.75s — MEASURED 2026-08-23, not fixed

Pooled over 2,917 timed batches in 169 rounds. Every batch in the archive is
ten shapes, so this is a curve in slide load, not in batch size:

    shapes this run already drew on the target    median batch
      0                                              3886ms
      1-20                                           5490ms
      21-50                                         13995ms
      51-100                                        18074ms

**A run slows itself 4.7x by piling onto one slide.** This is the effect
`shapesDrawnOn` already documents ("about +0.44s per shape present … quadratically")
and it is now measured across the whole archive rather than in one round.

What a user sees, from the two self-test arms:

    insert on top of an earlier run           3s   4 charts onto NEW slides    ~0.75s each
    insert onto a slide that already has …   48s   2 charts onto an EXISTING   ~24s each

**About 32x, and it is structural rather than a bug.** The two arms separate
cleanly on the path they take — counted from one round's trace:

    insert on top of an earlier run          0 batches   2 file handovers
    insert onto a slide that already has …   4 batches   0 file handovers

The fast arm draws **no shapes at all**: it hands the host a generated file. That
is also why 4 charts in 3s is not a contradiction of the curve above — a single
ten-shape batch has a floor of 3886ms, so 4 charts drawn shape-by-shape could
not finish in three seconds and never do.

You cannot hand the host a file to add shapes to a slide that already exists, so
that path draws shape by shape and pays the curve. **SSF Charts is fast exactly
when it CREATES slides and slow exactly when it ADDS to one.**

**Why this is a product decision and not a fix.** The side-by-side layout — two
or three small charts on one slide — is the everyday think-cell workflow, and
it is the slow case. Nothing in the pane tells the user that, and nothing
should be changed unilaterally: the cost is the host's, the mitigation
(`leastLoadedChart`) is correctly scoped to the self-test, which picks a chart
arbitrarily and so should pick wisely. A user picks the chart they mean, so
there is no choice for the product to make on their behalf.

**A FOURTH OPTION, COSTED 2026-08-29: draw it as ONE PICTURE.** The three below
all decide what to SAY about the cost. This one removes most of it, and it is
the only route that can — the cost is per shape drawn, so the way out is to draw
fewer shapes, and a chart can be one.

Re-derived from `batch issued`'s own `prevBatchMs` and `onSlide`, over the whole
archive rather than the 169 rounds above:

    shapes already on the slide      median ten-shape batch
      1-20        n=2,405                    5,013ms
      21-50       n=  904                   13,971ms
      51-100      n=  241                   16,579ms

    rasterising one slide            n=2,093    705ms

So a 24-node chart onto a crowded slide is three batches at ~14.3s — **about
43 seconds** — against a rasterise of 705ms plus a single-shape insert. Even
pessimistically that is under ten seconds, so the saving is roughly **7x to
20x**, and every part of the mechanism already ships: `pictureBase64`, the
auto-picture fallback for a host that has failed to draw shapes, and the
`explode a degraded picture` scenario that turns one back into native shapes.

**It stays the owner's call, and it is a bigger call than the offer was.** A
picture is not a chart PowerPoint can edit — no nudging a bar, no restyling
through the ribbon — and the add-in's own explode path is the only way back.
That is a change to what the user RECEIVES, where the three options below only
change what they are TOLD. Costed here so the decision is informed; not taken.

The options are all product calls: say nothing; show progress honestly on the
slow path; or offer "put this on its own slide" when the target is already
loaded. Worth a decision, not worth guessing at.

**What this is NOT.** Not a batching problem — an empty Office.js round-trip is
**5ms** (the driver's readiness ping, `slides.getCount()` plus one sync), so
round-trips are ~0.1% of a batch and raising `SHAPES_PER_SYNC` cannot help; the
cost is per-shape drawing on the host. And not a visibility problem: an earlier
reading of this same data claimed 2.3x for "the slide the user is looking at"
and was retracted in `0638b6a` — it split on a sentinel meaning "slide id not
loaded yet". Nothing in this archive distinguishes on-screen from off-screen.

### The slow-insert offer — SHIPPED 2026-08-29, and the blocker was a stale finding

**Owner's decision, taken 2026-08-27:** show progress on the slow path, and when
the estimate exceeds **14 seconds**, quietly offer to put the chart on its own
slide instead. Not a modal — it interrupts the press the user just made — but it
does gate this insert, because it is a question about THIS chart.

All five steps are done. `insert-cost.ts` prices the insert from the archive's
four measured medians; `worthOwnSlide` gates the offer on the slowness being the
SLIDE'S fault rather than the chart's; `offerSentence` quotes both costs;
`addSlideForChart` adds the slide through the existing retry path; and the pane
recomputes the placement for the blank slide instead of carrying over a cascade
that dodged obstacles which are not there. Five pane tests cover it end to end,
four mutants dead and the fifth recorded as equivalent.

**THE THING WORTH KEEPING IS WHY IT LOOKED BLOCKED.** This entry said the
wiring could not be finished because a fresh slide's id is useless on this host,
citing `insertSceneIntoSlide`'s own comment, which cited the 2026-08-08 probe
sheet: `shape-add-fresh-getitem-slide` answers `threw`, so a freshly-added slide
"does not work, by any route".

That was true when written and had stopped being true. Over 269 archived rounds
the answer is `threw` 227 times and `yes` 41 — and **every `threw` is round 253
or earlier**. From 254 on it is `yes` in 38 of 40.

**The flip is OURS, not the host's.** The commit on the boundary is 77f9ca4,
which stopped the probe HOLDING the id `slides.add()` hands back and made it
re-read the slide's id positionally once the add had settled. The two are
different id spaces, not near-misses — `4123571114#123571113` at add time
against `256#2587447327` a moment later, for the same slide. So the old answer
was never a fact about `getItem`. It was a fact about a stale id, recorded under
a question whose name says `getitem`.

The corrected rule, now in `host-baseline.mjs` and at the call site:

> A new slide's id is not durable until the slide settles, and is durable
> afterwards. Re-read it positionally after the add and `slides.getItem(id)`
> resolves it.

`addSlideForChart` was already built that way — it goes through `addSlides`,
which does not return until a FRESH context has confirmed the deck grew, then
reads the id off a positional handle. So the wiring worked, and only the
documentation said it could not.

**Two lessons, and the second is the expensive one.**

- A probe question is named for the CALL it makes, and answers about whatever is
  actually broken. `shape-add-fresh-getitem-slide` spent three weeks reading as
  a verdict on `getItem` while measuring a stale id, and the fix for the stale
  id silently rewrote the verdict.
- **A finding copied into a comment does not get re-measured.** The sheet is
  regenerated every round; the comment quoting it is not, and it was the comment
  that stopped this work. Where a comment leans on an archived answer it should
  say which question, so the next reader can re-run it in one command instead of
  believing it.

Still open here: making an insert onto a crowded slide actually faster. The
offer routes around the cost; it does not remove it. See the entry above.


### Text that overlaps text on data the samples do not carry — MEASURED 2026-08-19, not fixed

Superseded 2026-08-28. Its 75-pair figure was measured with a cruder ink rule and was never comparable; re-measuring gave 2,148. See **Text drawn over text** in §3, and `test/overlap-budget.test.ts` for the live per-shape table.

### The crash is in the deck SCAN, not the screenshots — located 2026-08-27

Superseded by **The crash, counted across the whole archive** in §3, which counts the whole archive by build rather than by record and finds no dominant phase.

### Report what this project has measured to the office-js tracker

**Researched:** 2026-08-06. **Owner-gated — nothing may be filed without the
owner's word**, because it goes out under his GitHub identity.

Reading the tracker turned up five open defects sitting under code this add-in
ships, and the weekly sweep (`scripts/office-js-watch.mjs`) now keeps that
current. The traffic has been one-way. This project holds measurements of
PowerPoint on the web that are **not in the tracker at all**, and a fixed host
retires a workaround permanently where a guard only routes around it.

**REWRITTEN 2026-09-05. THREE FINDINGS BECAME TWO, AND BOTH NAMED THE WRONG
VARIABLE.** Read back against the answer sheet the item itself says to attach
(`test/fixtures/host-answers-web.json`, build `8643e2d`) and against the round
archive. Ready-to-paste drafts are in `docs/OFFICE-JS-DRAFTS.md`; nothing has
been filed.

**~~2. A non-empty `setSelectedShapes([id])` wedges the selection subsystem.~~
STRUCK — it does not, on this host, and has not for 371 rounds.** The item said
"measured, twice: `getSelectedShapes` ran out a 90-second budget".
`selectionLadder` was built to prove exactly that, and has since run in **378
archived rounds: 371 report "the host answered all 7 rung(s) — nothing wedged",
2,597 rungs, zero silences.** The seven non-passes are skips — no probe chart,
or a deck scan that could not see the deck — not wedges. Whatever was seen twice
has not recurred in 371 consecutive opportunities, and filing it would report a
behaviour this project's own instrument contradicts. Those same rounds are also
counter-evidence to the "never resolves" half of #3698 and #4225: the EMPTY call
returns in about a second, every time.

**1. The variable is the SLIDE, not the proxy's age.** The item said
`shape-add-held-slide-proxy` threw, and concluded "it is the holding that
fails". The sheet says otherwise, and the real split is cleaner:

    shape-add-held-slide-proxy         yes    3/3   scratch=reused-slide
    shape-add-held-slide-proxy-again   threw  3/3   scratch=fresh-slide

Same question, same host, moments apart. Holding a slide proxy across a sync is
fine on a slide that already existed and throws `GeneralException` on one added
this session. Two controls survive and narrow it: resolve-and-use inside ONE
sync works on both, and `getItemAt(index)` works on both. So it is not the id,
not the slide and not `getItem` — it is holding a proxy to a NEWLY ADDED slide
across a sync.

**3. The variable is a COLLECTION RE-READ, not the proxy's age either.** The
item blamed "a shape proxy several syncs old". The sheet says age is not what
kills it:

    tag-the-creation-proxy-a-sync-later            yes             3/3
    how-many-syncs-a-creation-handle-survives      survives-8      (healthy)
    how-many-collection-reads-a-context-survives   short-at-1      3/3  (0 of 3)
    collection-read-poisons-the-creation-handle    refused         3/3  (5010)

A creation handle survives EIGHT syncs untouched. Re-read the slide's shapes
once — a read that itself comes back short, 0 of 3, with no error — and the
handle is refused with `InvalidParam passed to GetItem(id)`. That is the
mechanism behind the 46 failures in one 38-item run, and it is one finding
rather than two.

**The corroboration target moved.** #2903 is not the dead end `KNOWN_ISSUES`
treats it as — its closure was an automated inactivity sweep, not a decision —
and #4204 is an open re-file of it with zero comments. Better: **#6237 is open,
labelled `Type: product bug`, and had Microsoft activity on 2026-08-31**,
carrying the identical 5010 / `ShapeCollection.getItem` on a tag read. Its
reporter blames a date placeholder, which is a DIFFERENT trigger from the
collection re-read measured here — so that is a corroborating comment with a
second mechanism, and the cheapest thing on this list.

**What makes this a day of work rather than an afternoon.** Microsoft's issue
template wants a Script Lab repro, and an issue without one is triaged slowly or
not at all. Each finding has to be reduced to a minimal snippet a stranger can
paste — the probe is the wrong artifact to hand over, being a whole add-in.
That reduction is the work; the findings themselves are already written down.

**Attach the answer sheet.** `test/fixtures/host-answers-web.json` is a
reproducible record with the build stamp and requirement sets in it, which is
more than most reports carry. Check it before sending: every chart in this repo
is invented dummy data, but a sheet carries slide ids and raw host error text
from a real deck, and the owner should read what goes out under his name.

**Do not open duplicates.** Search the tracker first — the sweep's
`KNOWN_ISSUES` table is the list of what has already been read, and #2903,
#3083 and #3698 are corroboration targets, not new issues.

**Priority:** medium. High leverage and very slow: the issues this project
depends on have sat open for one to nine years, so nothing here should be
planned around a fix arriving. File it because a fixed host helps everyone
writing a PowerPoint add-in, not because it unblocks us.

### Drive PowerPoint on the web from Playwright, so the manual run is a command

Shipped. It is the round loop: `npm run cycle`, `scripts/round.mjs`, `scripts/cyclePlan.mjs`, and 290-odd archived rounds under `rounds/`. See docs/ROUNDS.md.

### ~~Take more than two draws per arm~~ — MOOT, the question it served is closed

Struck through by its own author: the question it served closed. See git.

### ~~READ THIS FIRST: a chart on a FRESHLY ADDED SLIDE cannot be grouped~~ — FIXED, measured 2026-08-25

36 of 36 charts on freshly added slides now group. Struck through by its own author; see git for the eight-commit route.

### THE WORK QUEUE that comes out of all this — 2026-08-16

Stage 0 of the matcher fix was done first, deliberately, because it decides the
shape of everything under it. **It dissolved the blocker it was meant to
confirm**, and split one item into two with very different prospects.

**What Stage 0 asked:** the fix wants to treat a short re-read on a slide we own
as a host lie rather than an instruction. Who can honestly promise the slide is
ours? `Grouping.refreshShapes` documents itself as exactly that guarantee — "the
caller can guarantee the target N shapes are the last N on the slide" — but every
call site sets it from `spansBatches()`, i.e. "this chart spanned sync batches".
**The field's contract and its use have diverged** (see the separate item below).
And the obvious substitute is unsafe: `onSlide === 0` means only that THIS RUN has
not drawn there, which is equally true of the user's own slide full of content.

**What Stage 0 found instead, and it is better news.** The ownership guarantee is
not needed for the short-read case, because **the matcher already proves
ownership by our own ids**. It matches `created[k].id` against the re-read's
collection; the 20 of 24 that matched on chart 4 are provably ours, on any slide,
without any new flag. Nothing has to be promised.

So the two failures separate:

| | chart 4 — short read | chart 5 — empty read |
| --- | --- | --- |
| what came back | 20 of 24, matched BY OUR IDS | nothing at all |
| ownership | proven, no flag needed | nothing to prove it with |
| what blocks it | a TRADE-OFF, not a signal | genuinely needs the guarantee |

#### 1. ~~Group a partial match rather than declining~~ — CALLED AND SHIPPED 2026-08-19

**The owner took it: group the 20.** The comparison it was decided on:

    group the 20    chart re-editable, 4 shapes stranded in its box
    group nothing   chart whole, and loses its config about 7 times in 10

**What shipped**, so a round is read against the right thing:

- The matcher groups a partial match instead of discarding it, and the
  short-read trace carries `grouping:` saying which — a round archive spans the
  change, and the same line meant "so it was not grouped" before it.
- **A bound the call did not name.** The subset is taken only when it is a strict
  MAJORITY of the chart. 20 of 24 is what every round produces; 1 of 24 would be
  a group holding one label with twenty-three shapes around it — the config saved
  and the object destroyed. Majority rather than a tuned share, because "more of
  the chart is inside the group than outside it" is a statement about the object.
  **Worth the owner's eye**: it is the one part of the change he did not specify.
- **`wholeMatch`**, because `freshMembers` now carries subsets and two readers
  need whole lists. `ungroupedFallback` builds the parts tag off that map, and a
  SHORT parts tag is worse than none — the next update deletes what it can name,
  redraws everything and leaves the rest, so the chart grows on every edit. That
  case is reachable: grouping can throw after a subset was chosen. The
  single-member tag-target swap is the other reader.
- **The stranded remainder is NOT recorded for the update to delete.** Tried
  first, and it is how this trade would become data loss: the only ids we hold
  for those shapes are the ones creation returned, and creation ids are what this
  host has been seen not to answer to (`withOwnId 7 of 7 … matched 0`, rounds
  068/069). The parts tag is a list the update path deletes BY.
- `chooseGroupMembers` unchanged, deliberately: it sees the re-read's members and
  not `created.length`, so it cannot tell two of two from two of twenty-four and
  cannot judge whether the group would still be the chart.

Three guards, each mutation-proven against the line it protects: the subset
groups and says what it left behind, a minority does not group, and a partial
group that then throws still writes a whole parts tag.

**WHAT A ROUND SHOULD SAY.** Staked before it runs, per the method below.
`same scale across the deck` fails at chart 4 in every round on record, and
chart 4 is the 20-of-24 one — so it should now group, keep its config, and the
scenario should reach chart 5 (the EMPTY read, item 2, which this does not
touch). `grouped the chart's shapes` should carry `partial=1 left=N:4` for it,
which is the first time that field reports an intended outcome. If chart 4 still
loses its config after grouping, the tag write is refused THROUGH THE GROUP —
which is what rounds 064/065 found on a freshly-added slide, and a different
problem from this one.

#### 2. The empty re-read (chart 5) — BLOCKED, and honestly so

Nothing came back, so nothing is proven ours and the positional rule would be a
guess on a user's slide. The only remaining route is `chooseGroupMembers`'
`use: "created"`, which hands `addGroup` the drawing proxies — and this host
throws on those, which takes the whole batch's tagging with it. So it needs
per-chart isolation of the grouping loop FIRST (today one try/catch wraps every
chart), and even then it is a gamble on a host that refuses those handles.

**And the paths that matter cannot make the promise anyway.** Stage 0 checked the
callers: only the self-test and the demo path add their own blank slides. The
pane's Insert draws on the slide the user is looking at, and Same Scale
(`updateChartsInSlides`) redraws onto the user's existing slides. **A fix gated on
"we added this slide" would improve the self-test and not the product** — which is
optimising the measurement rather than the thing measured, and this project has
made that mistake before.

#### 3. Retry a short or empty re-read once, after a delay — BUILT 2026-08-16, AWAITING ITS ROUND

**This did not come out of the archive. It came out of the office-js tracker on
2026-08-16, and it reframes items 1 and 2 above.**

**IT IS BUILT AND MERGED. What it has NOT had is a real host** — every claim
below is a prediction until a round tests it, and the prediction is staked here
on purpose so the round can refute it.

PowerPoint Online has a **known settling delay on a slide that has just been
materialised**: the shapes collection is not populated immediately after
`slides.add()`, and the community workaround — in
[#2903](https://github.com/OfficeDev/office-js/issues/2903), the very issue this
project already cites, and echoed in
[#5022](https://github.com/OfficeDev/office-js/issues/5022) — is to **wait one to
two seconds before reading it.**

This repo had read #2903 and recorded "upstream has nothing, `sleep(2000)` only".
That dismissal was reasonable when the failure looked like a tag problem. It is
wrong now: the failure has been isolated to *exactly* the state the workaround
addresses — a slide this run has just added. **We did not know that was our state
when we read the issue.**

**The shape of it: retry, do not delay.** Not a blanket wait before every
re-read — that taxes the 99% of charts whose read is already complete. Re-read
once more, after ~1.5s, only when the first read came back short or empty. Then:

    complete first read   costs nothing at all, which is charts 1-3
    short or empty        costs ~1.5s, on a chart that is otherwise losing its config
    still short           falls through to today's exact behaviour, nothing lost

**Why it beats both items above.** It needs no ownership guarantee (item 2's
blocker) and strands no shapes (item 1's trade). It is additive: every existing
path is reachable and unchanged, and the worst case is the current behaviour plus
a second and a half.

**What would prove it:** one same-build pair. `WHICH SLIDE THE CHART LANDED ON`
should move `freshly added, empty` off 1% grouped. If it does not move, the
settling delay is not our mechanism and the tracker lead is spent for the price
of one round.

> **ANSWERED, 2026-08-25 — it moved to 100%.** `freshly added, empty` reads 36 of
> 36 across eight consecutive rounds. The prediction staked above was correct and
> the mechanism is ours.
>
> **Read the delay before celebrating it.** The metric read `0/0` for the twenty
> rounds before that, so this had been true for an unknown number of rounds with
> nothing able to say so — see the entry in the journal. The prediction was
> right; the loop simply could not hear the answer.

**Stake the prediction before the round** (the method that earned itself
overnight): charts 4 and 5 of `same scale across the deck` group, and the
scenario stops failing at 34 of 34.

### ROUNDS 064 + 065 ANSWERED IT — half right, and the half that failed is the useful half

**Run as a PAIR on `bcd5773`, no merge between them, and the two are structurally
identical chart for chart.** `the settled retry repaired 2` in both; charts 4 and
5 take the retry, group, and are refused through `from: group×1` in both; the
slide ids differ between rounds, so what repeats is the position and the
freshness rather than an id. The noise floor is about counts and none of this is
a count — no third round is needed. What DID move is the downstream score, 4 of 8
then 3 of 8, which was always the noisy half.

**Charts 4 and 5 grouped**, both after the retry, both for the first time in five
rounds of failing identically. `the settled retry repaired 2`. No chart in the
round went ungrouped at all.

**The scenario still failed**, because the tag write is refused *through the
group*:

    4/8   tagging failed   from: group×1   slide 258#4111159134   5010
    5/8   tagging failed   from: group×1   slide 259#3844610554   5010

**This refutes "grouping is what saves a config."** That 2%-vs-79% split was
measured on a population that excluded freshly-added slides by construction —
they could not group, so they could never be in the grouped column. Round 064's
own split is `grouped 5, 2 lost = 40%`.

**The next target is the SLIDE handle, and it is named rather than guessed.** A
shape proxy carries its parent's object path; the group is made fresh in the
grouping batch but hangs off a slide handle Office has rewritten to
`slides.getItem(id)`, and a freshly-added slide's id does not round-trip here
(`shape-add-held-slide-proxy: threw`). The round says it outright: *the host
grouped through a slide handle two syncs old.* The members were never too old —
the parent was.

Keep the retry regardless: it is what turned an invisible failure into a named
one, and it costs nothing on a healthy read.

**What shipped**, so the round is read against the right thing:

- `REREAD_RETRY_MS = 1500`, `REREAD_ATTEMPTS = 1` in `powerpoint.ts`. The
  pre-grouping re-read runs in an attempt loop over a `pending` list; a chart
  whose answer was complete leaves the loop after the first pass and pays
  nothing. A fresh slide handle is taken per attempt, because Office.js rewrites
  the previous one to `slides.getItem(id)` and this host refuses shapes hanging
  off that.
- Two new trace readings. `re-reading the slide's shapes again after a settle
  delay` says the retry ran, with the slides it ran for; and both failure lines
  now carry `afterRetry: true`, so a round log distinguishes "short" from "short
  even after a pause" — the second is a much stronger claim about the host and
  the logs could not make it before.
- `hostFriction.emptyReReads` counts only reads the retry failed to repair, so a
  round is not told the host failed on a chart that came out fine.
- `faults.readsMissingFirst` in the fake — short once, then honest. Neither
  existing fault could express a slide that settles (`readsMissing` is permanent,
  `hollowReads` heals but only from empty), so without it the suite could show
  the retry costing a delay and never show it saving anything.

Guarded by two tests that assert **the group**, not the absence of a complaint,
and mutation-proven: `REREAD_ATTEMPTS = 0` turns both red.

**If the round refutes it**, the thing to check first is whether 1500ms is simply
too short — the tracker's reports say "one to two seconds" and this sits at the
bottom of that range deliberately, because the cost is paid on the live insert
path too.

#### DONE — cleared 2026-08-16

- **Revert `tagAnchorIndex`.** No measured effect across five rounds and four
  builds. The anchor is `created[0]` again, the draw loop loads every id, the
  matcher's contiguity deduction is gone, and `parts()` is `.slice(1)`. Kept, as
  planned: the `origin tag lost` trace line, the fake's `handleResolved` split
  (the host confirmed it), and the `taggedShape` test helper. The two assertions
  that flip back are annotated with what they asserted in between and why, and
  the reasoning behind the anchor move is preserved in full beside
  `CHART_ORIGIN_TAG` rather than deleted — it was good reasoning that this host
  does not reward.
- **`Grouping.refreshShapes` says one thing and is set from another.** Fixed the
  comment rather than introducing the guarantee, because the guarantee has no
  honest caller. Two sibling comments in the same file carried the same stale
  "last N on the slide" claim and were swept with it.

### Grouping is what saves a config, not the tag handle

**Measured 2026-08-15 over the whole archive, and it had been sitting there for
eleven rounds unqueried:**

    grouped      64 chart(s),  1 lost the tag =  2%
    NOT grouped  62 chart(s), 41 lost the tag = 66%

Per round it is almost mechanical — three grouped and none lost, two or three
ungrouped and two lost, round after round after round. `npm run rounds` prints it
now, under **DOES GROUPING SAVE THE CONFIG**.

**Why it is that stark.** When grouping succeeds the tag target is the GROUP —
a handle made in the grouping batch, never loaded, never resolved — and the write
lands. When grouping is skipped the target falls back to a `created` handle, and
that is the path that loses two charts in three.

**What it means for everything below.** The three sections that follow are about
WHICH handle the fallback should use, and four rounds plus a merged renderer
change (`tagAnchorIndex`) went into them. They are a question about the losing
path. A chart that never has to take that path does not care what the answer is.

`not grouping: no member handle this host will accept` carries `refreshed: 0`
every time, so what actually decides a chart's config is **whether the
pre-grouping re-read returned anything**. That is one level up from where this
work has been aimed, and it is where the next attempt belongs — the re-read, not
the tag.

**AND THE RE-READ FAILS THE SAME TWO WAYS EVERY ROUND.** Per-chart, identical in
rounds 042 through 046 without exception:

    1/8, 2/8, 3/8   GROUPED by=ids                          config kept
    4/8             re-read PARTIAL, matched 20 of 24       no group → config lost
    5/8             re-read EMPTY, drew 24                  no group → config lost

Two different faults wearing one outcome, and both are deterministic rather than
moody — the same chart, the same numbers, five rounds running.

- **Chart 4 matches 20 of 24, every time.** A partial match is thrown away on
  purpose (grouping a subset strands the rest), so this chart is ungrouped by our
  own rule and then loses its config to the fallback. Twenty of twenty-four is
  not noise; four specific shapes are not being named. Find those four and the
  chart groups.
- **Chart 5's re-read comes back EMPTY** with 24 drawn — the host listing nothing
  for a slide it has just been drawn on.

**The comment in the partial branch has been corrected** (`powerpoint.ts`): it
claimed an ungrouped chart "is still tagged, still re-editable", which the 66%
above refutes.

**BOTH BULLETS ABOVE ARE HISTORY AS OF 2026-08-19, AND THIS SECTION ASSERTED THEM
AS CURRENT UNTIL THEN.** They describe rounds 042-046. Two things have happened
since, and the file said the opposite of each in another section while still
saying this here:

1. **The partial branch is no longer left alone.** #586 groups the majority the
   host names. "The branch itself is left alone deliberately" was the rule for
   eight days and is not the rule now.
2. **The re-read WAS fixed, which is why neither harm is being chosen.** The
   settled retry (`REREAD_RETRY_MS`) repairs it: rounds 079-087 report
   `repaired=5, re-editable=8` every time, `same scale across the deck` passes,
   and no round records a short re-read since the retry shipped (42 in the whole
   archive, the `afterRetry` field ABSENT on every one — it postdates them all).
   "Fix the re-read and neither harm has to be chosen" was the right instruction
   and it has been carried out.

   **CORRECTED 2026-08-20 — that zero is a floor, not a count.** Until then the
   COLD read's outcome was never traced: attempt 0 pushed the entry onto the
   retry list and returned in silence. So "0 since the retry" compared a
   post-settle read against 42 COLD ones — different units, and the archive
   could not distinguish "the fault stopped happening" from "the retry hides
   it". The cold read is traced now; a few rounds will say which.

   **ANSWERED 2026-08-20, by rounds 111 and 112 as a pair: THE RETRY HIDES IT.**
   The cold read fails 8-11 times a round, every round, and that replicates.
   111: 11 failures (3 short, 4 empty, 4 zero), 11 retries, 0 survivors.
   112: 8 failures (2 short, 3 empty, 3 zero), 8 retries, 2 survivors (1 empty,
   1 zero). So the retry does not repair everything — but it repairs `short`
   FIVE TIMES OUT OF FIVE across the pair, and `short` is precisely what the
   subset branch below needs. The branch is starved by the one case the retry
   is best at, while the two cases that do survive are ones it cannot use.
   That is a sharper answer than either round gave alone.

So chart 4 does NOT match 20 of 24 any more — 084-087 group every chart with
`partial:0`. #586's subset branch is presently unreachable on this host, which is
not an argument against it: it is the guard for a regime this host has left, and
the round that would exercise it has not happened.

**THE `NNN#0` SLIDE LEAD IS DEAD — refuted 2026-08-15 from the archive, before
the instrument built for it had even deployed.** The draw trace already carried
both `chart` and `onSlideKey`, so the join could be done on rounds already
archived:

    1/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    2/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    3/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    4/8  slide 257#0            no group  CONFIG LOST    <- a #0 slide
    5/8  slide 288#1168146411   no group  CONFIG LOST    <- an ordinary id

Identical in rounds 043-046. A `#0` slide carries the three best-behaved charts
AND one failure, and an ordinary slide fails too. **The id shape predicts
nothing.**

**WHAT DOES PREDICT IT: POSITION IN THE DECK-WIDE UPDATE.** Those five charts are
`same scale across the deck`, and its own summary says `the host flipped at chart
4 of 8, so the last 3 were not attempted` — in every round it has run (34 of 34
on 2026-08-16; `npm run rounds` has the live count).
Read the re-read outcomes in order and it is a decay curve, not a coin:

    charts 1-3   re-read matches all 24     grouped, config kept
    chart 4      re-read matches 20 of 24   partial, thrown away, config lost
    chart 5      re-read returns NOTHING    config lost
    charts 6-8   never attempted

**The hypothesis, and it points at our own perf work.** `updateChartsInSlides`
was deliberately made **one context, four syncs, flat in N** — the fix for
`doSameScale` spending 80 syncs across 20 contexts. That is the right shape for a
host that can hold a context, and this one degrades as a context is used: the
re-read is the first casualty, at chart four, every time. `#112` already made the
opposite call for the demo deck, one `PowerPoint.run` per slide, because a
context that has done too much stops answering.

**The experiment: chunk the deck-wide update into a fresh context every ~3
charts** — not per chart, which is what the perf work correctly removed, but at
the boundary the data actually shows. Cost is two extra contexts for an 8-chart
deck against 20 for the old shape. **Prediction: `same scale` moves off 17-of-17
failures, and charts 4-8 start keeping their configs.** If it does not move, the
context is not the limit and the re-read fails for a reason that survives a fresh
one — which is worth knowing for the price of one round.

Two further leads on the re-read, both cheap and neither yet followed:

- **The `NNN#0` slide ids.** Two of every round's added slides come back with an
  id whose second half is `0` — not the shape this host gives a slide it has
  finished adding — and the one line that names a slide in a tag failure named
  `257#0` in rounds 043, 044 and 045. The failure lines now carry the slide
  (`not grouping`, `a chart's tag could not even be queued`, `tagging failed`),
  so the next round says outright whether the charts that lose their config are
  the ones sitting on those slides. If they are, this is one fault and not
  three, and it is a SLIDE fault.
- **The re-read comes back empty rather than short.** `the re-read before
  grouping came back empty` appears beside the losses, and
  `shapes-items-count-honest` answers `unreadable` with `items` undefined. A
  slide whose shape collection will not enumerate cannot be grouped and cannot
  be tagged; both symptoms may be the same slide problem.

### The settle pass cannot repair what this host loses

MEASURED, 2026-08-15, rounds 029-034. `same scale across the deck` has failed six
rounds running at 4-5 of 8, and every loss ends the same way:

    settle pass: could not repair any config tag the drawing context lost
      charts=1  settled=0  lost=1  withId=0

`withId: 0`, every time. The settle pass is the second chance that exists to put
a config tag back on a chart the drawing context could not tag — and it works by
resolving the shape by id. This host does not give ids back for shapes it has
just made (`tag-through-refetched-shape: no-id`, three rounds, nine samples), so
the repair has nothing to try. It is not failing; it is unreachable.

The host's own error names the same wall from the other side:

    errorLocation: ShapeCollection.getItem
    statement: var shape = shapes.getItem(...) /* originally addTextBox(...) */;

Office.js rewrites the CREATION proxy's object path into `getItem(id)`, so even
the code path that never mentions an id has become an id lookup by the time the
host sees it.

What makes this actionable rather than another dead end: round 033 and 034 both
answered `tag-the-creation-proxy-a-sync-later: yes`. **A write through the
creating handle goes through on this host, a sync later, when neither the id nor
a read does.** So the settle pass has a route it is not using.

**The settle pass cannot be where this is fixed.** It opens a FRESH
`PowerPoint.run`, so the creation handle — the one thing this host still accepts
writes through — does not exist by the time it runs. Its two routes are an id and
a collection read, and this host refuses both (`no-id`; `shapes-items-count-honest:
unreadable`). `withId: 0` is that fact showing up as a number.

So the retry belongs in the DRAWING context, before the handle is thrown away:
when the tag write is refused there, write it again through the proxy that
created the shape rather than handing the chart to a repair pass that has nothing
to work with.

This is the same goal `binding-names-shape-later` was written for — "whether the
repair pass can be given a handle that does not go through
`ShapeCollection.getItem(id)`" — and that question answers `unreadable` on this
host, four rounds running. `tag-the-creation-proxy-a-sync-later` is a second
route to it that does answer `yes`, and it needs no 1.8 surface.

**ATTEMPTED 2026-08-15, and the one-line version does not work.** Worth recording
so nobody spends the hour again.

The rule was measured first: it is RESOLUTION, not age. `faults.strictTags`
models "refuse anything older than one sync", and this host does not do that —
`tag-the-creation-proxy-a-sync-later` answers `yes` four rounds running, so a
handle that made a shape keeps taking writes however old it is. What it refuses
is a handle a `load()` has resolved, which Office.js rewrites into
`shapes.getItem(id)`. `faults.refuseTagWritesOnResolvedProxy` now models exactly
that, and it is the first thing in the fake that can.

The obvious fix follows and is wrong: `finishCharts` replaces the tag target with
the pre-grouping re-read handle (`tagTargets[i] = fresh[0]`), so keep the
creation handle instead. Tried, and the reproduction still failed — because the
pass loads the CREATED shapes too before the tag is written, to read their ids
for the parts tag. By the time the write happens there is no unresolved handle
left to use.

So the fix is not a swap, it is an ordering change: one handle has to stay
unresolved all the way to the tag write. That is a change to the shape of the
whole pass and wants a session of its own.

**THE BUDGET IS MEASURED NOW, and it is not a constraint (round 039, 2026-08-15,
replicated from the recovered round-42 log).**
`how-many-syncs-a-creation-handle-survives` answers **`survives-8`** — five
samples across two builds, `stable: true`, and every sample taken while the host
was in the `collection-refused` regime. An unresolved creation handle accepted a
tag write at each of eight successive syncs.

That removes the gamble the probe was written to remove. The ordering change does
not have to move the tag write EARLIER in time; it only has to keep one handle
unresolved. Eight syncs is past anything a chart does, so the constraint is
purely `load()` — and the resolver has a name and a line number:
`powerpoint.ts:6956`, `created[k].load("id")` on each batch's own sync, which
resolves every created shape including `created[0]`, the tag target. Round 039's
host dump shows the consequence with the ids in it: `shapes.getItem("27")
/* originally addTextBox(...) */`, where 27 is the first batch's title box.

`ungroupedFallback` is NOT that resolver, so swapping it changes nothing:
`parts()` slices index 0 off deliberately, and the anchor is never in its
`load("id")` list.

**AND THE LAST ALTERNATIVE IS CLOSED (round 040, 2026-08-15).**
`tag-through-refetched-shape` — "can a fresh shape be tagged through a handle
re-fetched by its own id?" — answers **`no-id`**, stable over three passes: *the
fresh shape would not report an id*. Not "the id is refused"; there is no id to
re-fetch it by. Route by route on this host:

    created    works, and does not age out (survives-8) — the pass throws it away
    refreshed  resolved by a load, rewritten to getItem(id), refused
    by-id      cannot be constructed at all — no id
    group      never loaded, but only exists where grouping worked, and here it does not

So the ordering change is not the best remaining option, it is the only one.

One caveat the probe cannot cover, stated so the fix is not built on it: it asks
through `ctx.scratch()`, a slide resolved in the same batch, while production's
parent is a slide handle Office has rewritten to `slides.getItem(id)`. A refusal
originating in the PARENT path would be invisible to this question. The host's
`errorLocation` is `ShapeCollection.getItem` rather than the slide, which is
evidence against that reading but not proof.

### Why it is a redesign and not a reordering — the three-way knot

Attempted properly on 2026-08-15 and stopped at a constraint rather than at
effort. Recorded because every route below LOOKS like the fix for about twenty
minutes, and the thing that kills it is somewhere else in the file.

Three requirements, each independently non-negotiable, and no two-line change
satisfies all three:

1. **The tag target must be unresolved when the tag is written.** Measured:
   `survives-8` and `tag-through-refetched-shape: no-id`. A `load()` is what
   makes Office rewrite the handle into `shapes.getItem(id)`, which this host
   refuses.
2. **The pre-grouping re-read needs every created id.** `powerpoint.ts:7347`
   matches created shapes to the re-read BY ID, and a PARTIAL match is thrown
   away on purpose — the comment above it is explicit that keeping one strands
   shapes inside the chart's box, with `grouped … partial=1 left=0:4` from a real
   slide as the evidence. Skip loading the anchor and every chart becomes an N-1
   of N partial match, so **grouping stops working everywhere, desktop included**.
   The positional fallback is not available either: it is deliberately reachable
   only when NOTHING matched, because a slide holding the user's own shapes can
   satisfy `items.length >= created.length` and "the last N" would then group the
   user's content in, to be deleted with the chart on the next update.
3. **The tag must not land before the chart is complete.** Tagging `created[0]`
   in the first draw batch is the obvious way to get requirement 1, and it makes
   a stalled or stopped draw leave a tagged partial chart behind.

   **DO NOT CITE `stop a run part-way` AS EVIDENCE FOR THIS — corrected
   2026-08-19.** This entry rested on that scenario's pass ("nothing left
   claiming to be a chart"), and the scenario has never drawn a shape: it calls
   `requestStop()` BEFORE the insert, `throwIfStopped()` sits at the top of the
   batch loop, and across all 69 archived rounds it has committed **zero**
   batches. It tests that a stop asked for before a draw prevents it — a real
   promise, but not this one. **Requirement 3 is currently untested**, and this
   design constraint rests on reasoning alone.

   The seam for a real test is `onPhase("commit", …)`, which fires once per
   batch: requesting the stop from the first commit aborts a genuinely
   half-drawn chart. It is not a change to slip in — a mid-draw abort leaves a
   PARTIALLY DRAWN SLIDE, and `same scale across the deck` discovers its chart
   population from that same deck, so it would contaminate every scenario after
   it in the round. It wants its own pair.

The two routes the earlier note proposed both die on requirement 3 or on cost:

- **Tag early, drop the anchor's tag once the group's lands.** `TagCollection.delete(key)`
  does exist, so the mechanism is real, and on this host grouping never succeeds
  so the delete would never need to run. It still leaves requirement 3 broken.
- **A second key only the recovery path reads.** For the chart to be re-editable
  the READERS have to accept that key, and `chartTagsOf` does not read the scene
  tag today — so deduping "the group and its anchor are one chart" costs another
  per-shape tag lookup deck-wide, on a host already observed reading collections
  short. That is a real cost against the scan that Same Scale depends on.

**The shape of the actual fix, now that the knot is stated:** the tag target has
to be a shape created in the LAST draw batch and never loaded — then it is
unresolved (1), every OTHER created shape keeps its id for the re-read (2), and
nothing is tagged until the chart is finished (3). That means the anchor stops
being `created[0]`, which `ungroupedFallback`'s "everything after index 0 is a
part" and `CHART_ORIGIN_TAG`'s frame both assume.

### It was BUILT, and the fake refused it — 2026-08-15

Written in full and reverted the same hour, because the thing that refuses it is
worth more than the change was.

What was built: `tagAnchorIndex` moving the anchor to the last shape drawn, the
draw loop's `load("id")` running one shape behind so that shape is never
resolved, `ungroupedFallback`'s parts list taking everything but the anchor, and
the pre-grouping matcher learning that exactly one shape is deliberately
unmatchable — deducing it as the contiguous neighbour of the last matched
sibling, and falling through to today's don't-group behaviour when that does not
land cleanly. `CHART_ORIGIN_TAG` needs no change: it records the TAGGED shape's
own corner and shifts by how far that shape has moved, so it is self-consistent
whichever shape carries it. 211 office-render tests green, including the eleven
grouping ones.

**Then the reproduction still failed, and the trace said why.** The config tag
write was refused, `from: created`, in the drawing context — with the anchor's
own `load` held back. Instrumenting the fake's `load()` with a stack per call
found the resolver: not the draw loop, but `groupAndTagAll`'s
`shapes.load("items/id")`, the pre-grouping re-read. It resolved eight shapes;
the draw loop had loaded seven.

**The fake makes that read poison every handle onto the shape, and that is the
open question.** `freshHandle` is a `Proxy` over the same object: it deliberately
gives the fresh handle its own `syncCreated` and its own tag writer, and shares
everything else "because it is one shape". `loadedProps` is on the shared side —
so resolving a collection item resolves the creation handle too. Office.js gives
each proxy its own object path, which says they should be independent; the fake's
own treatment of `syncCreated` says the same. Either `loadedProps` is handle
state modelled as shape state, or it is the one place the fake is right.

**No round can be argued into answering it, so it is now asked directly:**
`collection-read-poisons-the-creation-handle` ships with this note.

- `yes` → the read is innocent, the fake is wrong, and the fix above works as
  written. Rebuild it from this description; it is not large.
- `refused` → no arrangement of loads saves the drawing context's write while
  grouping needs a re-read, and the second-key route is the only one left.

### ROUND 042 SAID `yes`, AND THE FIX IS IN — 2026-08-15

Three passes, `stable: true`, every one taken while the host was refusing
collection reads. The pre-grouping re-read does not touch the creation handle.
The fake was modelling handle state as shape state — `loadedProps` shared across
every handle onto a shape, while `freshHandle` had always given each handle its
own `syncCreated` and its own tag writer. Split, and the collection read now
loads through a fresh handle, which is what it hands back anyway.

Shipped exactly as described above: `tagAnchorIndex` (the last shape drawn), the
draw loop's `load("id")` one shape behind, the parts list taking everything but
the anchor, and the matcher deducing the one deliberately-unmatchable shape as
the contiguous neighbour of the last matched sibling. The reproduction under
`refuseTagWritesOnResolvedProxy` now shows the drawing context's own write
landing with no repair, and asserts the absence of a `tagging failed` line rather
than the presence of one.

**Two things it changed that were not planned, both kept:**

- **`origin tag lost` is its own trace line.** The config tag commits one sync
  before the origin tag, so the good case on this host is config-lands-origin-
  fails, and the old shared catch called that "charts are not re-editable until
  repaired" — false, and about to be the common case.
- **A short deck scan can no longer see a tail-anchored tag, and that is the
  price.** `faults.readsMissing` drops from the tail, so a scan blinded that way
  misses the shape carrying the config; `web-host.test.ts` asserts the new
  behaviour and, in the next line, that an unblinded scan finds the chart — so it
  is a visibility cost, not a lost tag. Taken deliberately: the write failed
  CERTAINLY, on every chart big enough to span batches, four rounds running,
  while this scan fails intermittently.

  **AND THE COST IS MOOT ON THIS HOST — rounds 043 and 044.**
  `which-end-a-short-read-drops` answers `unreadable` both times: *the collection
  would not list its items*. `shapes-items-count-honest` says the same on the
  same sheet, with `items` undefined rather than short. **A read that returns NO
  list cannot drop one end rather than the other**, so where the anchor sits in
  the collection cannot decide whether a scan sees it. The trade was real to
  take and turns out to cost nothing here. The question stays for a host that
  lists items — it costs one scratch slide — but nothing should wait on it.

**And the trace now says which handle was used**, which is what the six rounds
before it could not. `tagging failed` carries `from: created×N, refreshed×N`, one
of four routes:

    created    the proxy that drew the shape, never loaded — writes go through
    refreshed  the pre-grouping re-read, RESOLVED by a load and therefore
               rewritten to shapes.getItem(id)
    group      the group made in the grouping batch, also never loaded
    by-id      an explicit getItemOrNullObject(id)

The field earned itself on its first run. The suspect going in was `refreshed`;
the trace answered `created×1` — the target WAS the drawing handle and the write
was refused anyway, because the pass loads the created shapes before it writes.
That is the evidence for "ordering, not swap", and it took one test run rather
than a round to get.

Also measured while trying: arming `refuseShapeIdLoads` alongside — to starve
the settle the way the real host does — models something HARSHER than the host.
It makes the insert throw outright, where a real round carries on and merely
loses the tag. The reproduction in `test/office-render.test.ts` therefore pins
the first half (the drawing context's write is refused) and names the second (the
settle repairs it here and cannot there).

NOT attempted further, deliberately. Two things have to be true first and only
one of them is: the write-through-the-handle answer needs a third round, and the update
path needs the same treatment `insertSceneIntoSlide` already has — a test in
`test/office-render.test.ts` shows the INSERT path survives a refused
`load("id,left,top")` unharmed, so whatever costs the tag in the update path is
not the shared batch, and guessing which line it is has already produced one
wrong theory tonight.

### Stop the probe carrying seventy scratch slides through a round

**MOSTLY WRONG — THE DECK NEVER GREW. Corrected 2026-08-14 from round 29.**

The premise below was that the deck fills with abandoned scratch slides and that
this is what makes `listing the deck's slides` slow enough to time out. Round 29
measured it directly and says otherwise:

    gave the scratch slides back  returned=0 swept=0 left=73
      deckBefore=1 deckAfter=1 shrankBy=0 stillListed=0
      heldIds=["4123571114#123571113", …]  deckIds=["287#62081387"]

Seventy-three slides taken, and the deck is **one slide before and one after**.
None of the held ids is in the deck; `stillListed` is zero. The deck evidence in
the same round agrees — 7 slides scanned, 6 added, all by the self-test. The
scratch slides never landed at all.

Look at the two id namespaces. The scratch slides come back as
`4123571114#123571113` and the deck lists `287#62081387` — that is the
`getItemAt` → `getItem(id)` rewrite from `RESEARCH.md` seen from the other end: a
freshly added slide reports an id the deck will not answer to, so the probe can
neither find nor delete it, and its counter climbs while the deck stands still.

So the cost is NOT paid in an O(deck) slide listing, because the deck stays at
one; and the wedge these numbers were used to explain has since been identified
as the host's editing session dying (`docs/ROUNDS.md`, "The wedge"), which the
probe does not cause. **The counter is the bug, not the deck.** What is worth
building is not a cheaper clean-up but an honest count: `left=73` should say
`never landed=73`, and the round after that can ask why the host accepts an add
whose slide never appears.

The rest of this entry is kept because the two reverted fixes below are still
reverted for the reasons given, and both would still break the same guards.

Two fixes were built and both reverted on 2026-08-14:

- **Delete the abandoned slide, then add its replacement.** Safe-looking, because
  the abandoned slide is the deck's last at that moment so a positional delete
  can take it. It leaves the run with NO slide when the add then fails, and
  `asks every question even when the host keeps losing the scratch slide` caught
  it — `shape-add-fresh-slide-proxy` stopped being put at all.
- **Add the replacement, then delete the old one by its index.**
  `deleteTrailingSlides` deletes any range despite its name, so this works
  mechanically. It broke four guards, including
  `left its replacement slides in the deck` (11 where 2 was expected) and
  `shape-resolve-held-slide-proxy` changing its answer. Deleting mid-question
  does not merely tidy the deck — it changes what the probe measures.

So this is not a tidy-up, it is a change to the instrument, and it needs its own
round with a control rather than a green suite.

**The pass-boundary variant was then tried too, and it is blocked by something
else — the accounting, not the timing.** Sweeping between passes does leave the
questions alone, exactly as predicted: none of the four guards that killed the
per-replacement attempts fired. What fires instead is `leaves the deck exactly as
it found it`, and it keeps firing through every reconciliation:

- the end-of-run clean-up derives what it owes from `scratchIds.length` and from
  the deck's growth since the run began (`positionalSweepPlan`), and BOTH are
  invalidated by a slide going back early;
- scoping the clean-up to what is still outstanding fixes the double-count and
  then under-sweeps, because the positional plan is capped by `deckNow -
  deckAtStart`, which the early handbacks have already shrunk;
- feeding the handbacks back into the plan as `alreadyDeleted` changes nothing —
  the deck still ends with slides the run cannot account for.

Three spellings, one cause: **the clean-up assumes every scratch slide is still
there when it runs.** That assumption is load-bearing in `slidesActuallyReturned`,
`positionalSweepPlan` and the `scratch-slides-returned` answer, and no amount of
arithmetic at the call sites removes it.

The clean-up rewrite was then done — `outstandingScratch`, the ledger, tests that
exercise an early return — and **the pass-boundary sweep still fails, so the
accounting was not the blocker after all.** The instrumented numbers say why, and
they are worth more than the three theories that preceded them:

    taken 10 · handed back early 5 · deckAtStart 2 · deckBefore 12

`deckBefore` is read when the clean-up starts, after the handbacks. 2 + 10 = 12.
**The deck at clean-up time is as if nothing went back at all** — even though
`deleteTrailingSlides` reported five successful deletions and the deck visibly
shrank at the moment each ran.

So the handback deletes A slide and not necessarily THE slide it named. Which
makes sense of everything else: this host does not honour delete-by-id, and the
only reason the end-of-run sweep works is that by then every remaining scratch
slide is a **contiguous trailing block** — a positional RANGE needs no slide to
be identified individually. During a run that block does not exist, because the
live scratch slide sits at the end of it.

That looked like the real prerequisite — "there is no way to delete one known
slide mid-run on this host" — and **it is wrong.** Asked directly rather than
inferred from a misbehaving run, the answer is yes: `deleteTrailingSlides(i, 1)`
removes exactly the slide at index `i`, confirmed by id either side, and it keeps
doing so under `renumbersOnAdd` — the very behaviour that makes delete-by-id
useless here. `test/delete-one-slide.test.ts` pins it, and a mutation that
deletes a fixed index instead turns it red.

The run's bookkeeping is not at fault either: on clean main the ids it records
match the deck's growth exactly, 13 for 13.

So both things the last four attempts blamed are sound.

**The pass boundary was then instrumented — deck length either side of every
handback, and again at the next one — and the deletes are innocent too:**

    boundary 1   deck 7 -> 3   reported 4   actual shrink 4   MATCH
    boundary 2   deck 8 -> 7   reported 1   actual shrink 1   MATCH

Every slide the handback claimed to remove did leave, immediately, verified by a
count either side. The leak is in the SEGMENTS BETWEEN, and it is an ADD problem
rather than a delete problem:

    segment 1 -> 2         deck 3 -> 8  (+5)   run noted +4
    segment 2 -> clean-up  deck 7 -> 12 (+5)   run noted +2

The deck grows faster than the run records, and only while the handback is
active. Four candidates for the un-noted adds have been read and eliminated:
`addScratchSlide` reports a slide it could not remove (`onAdded` fires on exactly
that branch), `deleteSlideById` verifies with `slideIsGone` before returning
true, `deleteSlideByPosition` re-reads the deck and returns false rather than
guessing when the id is absent, and on clean main the run's own count matches the
deck's growth exactly.

So the question is now precise and small: **which call adds a slide between one
pass boundary and the next without the run recording it, and why only once slides
have been handed back?** Everything upstream and downstream of it has been
measured and is sound.

The next attempt logs the deck length either side of every single
`addScratchSlide` in one run — not the probe's behaviour, not the accounting,
just which add fails to conserve the count. That is one number and it ends this.

The ledger from the clean-up rewrite stays: it is right, it is tested, and it
removes a genuine assumption. It simply was not the blocker.

Four implementations have been reverted here, and every theory each was abandoned
on has since been disproved by a one-line experiment: the primitive works, the
bookkeeping is accurate, the deletes land where aimed. The pattern is worth more
than any of them — reasoning from a full probe run was wrong every time, and
asking a single component directly gave a clean answer in under a minute, every
time.

### Improvements the 2026-08-15/16 overnight run was owed — ALL CLEARED 2026-08-16

All cleared 2026-08-16. See git.

### Method that earned itself overnight, and should be kept

- **Run the same build TWICE.** Nineteen archived rounds and almost every one on
  its own build, so no two could be compared. The one exception, `cabb357` run
  twice, is the only noise measurement this project owns — 1 and 5 for the same
  fault with nothing changed. Pairs are what let `tagAnchorIndex` be called
  no-effect rather than unproven.
- **Judge a change on a count that did NOT move, or on a line that appears where
  none did.** Those are the two readings that survived; every count that moved
  moved with the host's mood.
- **Ask the archive before spending a round.** Grouping-vs-config, the `#0` ids
  and the fresh-slide split were all answered from rounds already filed. Three
  hypotheses died for the price of a query.
- **A question that cannot be asked is an answer about the harness.** Three
  probes died on the scratch slide refusing to enumerate a collection; the
  measurement moved into production and worked first time.
- **Stake a prediction before the round that tests it**, especially against your
  own work. The one staked on 2026-08-15 failed, and the failure is what
  redirected the effort.

## 2. Rejected or already covered (do not re-propose)

- **A PowerPointApi 1.8 binding as a durable handle to a drawn chart** — the
  last untried route out of `same scale across the deck`, and it cannot be
  ASKED on this host. Not refuted: unanswerable, which is a different and more
  annoying thing.

  The idea is sound and stays written down in case the host changes.
  `ShapeCollection.getItem(id)` is where every 5010 lands, and it is what leaves
  a chart drawn and nameless — no group, no config tag, nothing to settle one
  onto. `bindings.add` takes the live Shape proxy inside the batch that CREATED
  it, so it needs neither an id round trip nor a collection read, and the
  document persists it. If it worked, the repair pass would have a handle it
  does not have today.

  The probe `binding-names-shape-later` asked eight times across eight rounds
  and never once reached its own question:

      13  no-scratch-slide   15  no-scratch-slide   18  no-scratch-slide
      14  no-scratch-shape   16  no-scratch-slide   19  no-scratch-slide
      17  no-scratch-shape   21  no-scratch-shape

  Six distinct causes were found and five of them fixed — the end-of-run second
  pass (#353/#356), the poisoned slot after `shape-add-held-slide-proxy` (#360),
  an unclaimable add when the host renumbers a slide (#364), a slide-acquisition
  budget of ninety seconds against a question budget of eight (#364), and the
  probe's position in the order (#359). Each fix got it one step further. The
  sixth is not a defect and is not going away.

  **A shape question needs the scratch slide resolved once per batch, and this
  host hands out about one usable resolution per slide** — round 21's trace
  shows nearly every question taking a replacement before it can answer. The
  binding question needs more resolutions than any other shape probe, because
  an honest answer needs a CONTROL: the same batch without the binding, proving
  the host would have taken a plain shape. Without that control the probe claims
  `commit-threw` on a host refusing every shape add, which
  `test/host-probe.test.ts` catches. Running the control second does not help —
  a failed sync poisons its own context, so the control cannot resolve the slide
  afterwards. Both were tried and reverted, in that order.

  So: honest and unanswerable, or answerable and dishonest. Recorded rather than
  re-litigated. `same scale across the deck` therefore has **no route found**,
  and the failure under it (a chart that draws and is then unreachable) stands
  as a host limit. Anything that revisits this needs a host that resolves a
  scratch slide more than once, and the probe will say so the day one appears.

- **Settling the count `addSlides` verifies with** — swept 2026-09-07, and not
  done. The reasoning that gets you here is sound: the same late count that made
  `insertSlidesFromPptx` report `landed: 0` is read raw at `powerpoint.ts:4184`,
  and there it does not just misreport — a false deficit RETRIES the adds, which
  on a deck that was never short means duplicate slides.

  The archive says it has never happened. **128 `slides added` events across 401
  rounds, 0 shortfalls, 0 `slide add(s) never landed`, 0 `slide count disagreed
  between contexts`.** The retry loop has never run on this host. The drop it
  guards against is real but predates the archive — 10 slides for 20 issued adds
  in `Presentation_3.pptx` — so there is nothing here to fix and no round that
  could judge a change to it.

  Note the thin denominator honestly: 128 events is 0.3 a round, so this is
  "never seen in 128 tries", not "cannot happen". If a shortfall ever appears,
  the settle is the first thing to reach for and this is why.

  `replaceSlideWithDeck` was the other site the same sweep found, and that one
  WAS changed — its raw count returns "failed" before the delete, leaving the
  user both slides. It has never run in the archive either; what carried it is
  that the defect is proven on the same API call on the same host.

- **A golden-image gate on the generated deck** — rejected as a hash comparison,
  and largely covered as a structural one. The cheap half of this DID ship:
  `validate-ooxml.mjs` checks the deck against the OOXML grammar and
  `verify-deck.mjs` gained a duplicate-`cNvPr`-id check that neither tool had,
  both gating CI on `examples/showcase.pptx`.

  The measurements that made the visual half look feasible are real and still
  true (2026-08-02): LibreOffice headless renders the deck to PDF in 7s and to
  122 PNGs in 17.6s, byte-identically across fresh profiles, and its usual PPTX
  weak spots do not apply — this generator emits no gradients, no pattern fills,
  no `normAutofit`, no effects. What kills it is the **pinned container**. The
  deck asks for Segoe UI and Calibri, neither present in CI, so every render is a
  substituted render: self-consistent, different from a laptop with the real
  fonts, and different again after any LibreOffice minor bump. The baseline
  churns, and a visual gate that cries wolf is one that gets switched off — the
  failure this entry predicted about itself.

  **The structural half is thin rather than missing.** All three renderers
  consume the SAME scene graph, and `scripts/visible-charts.mjs` already proves
  that scene rasterises to something visible (ink present, more than one colour,
  inside the frame, at least half its usual coverage). A pptx-only invisibility
  bug would have to originate in `pptx-paint.mjs`'s own mapping, which is
  asserted per node at 100% statements / 88% branches, on top of `verify-deck`
  auditing the bytes and `validate-ooxml` checking the grammar.

  **Do not read this as "the pptx output is pixel-checked" — it is not.** The
  headless pptx renderer is the one of the three with no look at its pixels, and
  the residual gap is a composition bug that survives every node-level assertion.
  That is a narrow target, not an empty one. If it is ever revisited, copy the
  shape of `visible-charts.mjs` (assert properties that hold on any renderer, no
  committed images) and reuse its `judge()` wholesale; the part that wants
  pinning is LibreOffice, not the method.

  Two traps worth keeping whatever is built. Frame any such gate as "did our own
  output change", never "does this match PowerPoint" — no FOSS renderer is close
  enough for the second claim, and a gate that overclaims gets ignored. And
  **never** use a LibreOffice round-trip as a PowerPoint proxy: converting the
  showcase deck back to `.pptx` silently deleted all 122 `ppt/tags/*.xml` parts,
  leaving 121 charts non-re-editable. Anything learned that way is a fact about
  LibreOffice.

- **An image / icon node** — not reachable in the live add-in, so nothing can
  be built on it. PowerPoint's `ShapeCollection` exposes exactly
  `addGeometricShape`, `addGroup`, `addLine`, `addTable`, `addTextBox`;
  `addImage` exists only in the Excel namespace. The one route to pixels is
  `ShapeFill.setImage`, which is **PowerPointApi 1.8** against a manifest
  pinned at **1.4** — it would vanish silently on any older host. It also
  bloats twice over: `ChartConfig` round-trips through a shape tag verbatim
  (`tagData: JSON.stringify(cfg)`), and PptxgenJS does not dedupe media, so an
  identical icon is re-embedded once per point. The only thing presets cannot
  do is a **logo** (a competitor-positioning map). If that is ever wanted, the
  shape is: user-supplied data URI (no asset library, so no licensing call),
  gated on `supports("1.8")` like grouping already is, and drawn *additively
  over* a marker so an old host degrades to a plain point rather than a
  missing one.
- **Scatter/bubble point icons** — the primitive above is unavailable, and the
  real itch was not bitmaps: shape as a categorical channel is what Excel and
  Highcharts mean by point markers, and PowerPoint's presets give it filled and
  native at 1.4. Shipped as `scatter.markers`.
- **Heatmap per-cell icon overlays** — a heatmap's rows ARE its data series, so
  a "Glyph" row yields one row of glyphs, not a per-cell matrix; carrying a
  genuine second dimension would mean inventing a datasheet convention for
  every heatmap, for a want nobody has stated. The demonstrable gap was
  narrower and is fixed: a diverging scale states direction in hue alone, and
  its strongest + and − sit at 1.12:1 in greyscale — the same tone — with no
  label to fall back on under `sizeEncode`. Shipped as `heatmap.symbols`.
- **An X / cross marker, and a five-point star** — a marker shape has to
  reproduce its OOXML preset *exactly*, because the SVG renderer draws the
  points while the PowerPoint renderers name the preset, and `markerScale`
  measures the area off those same points. A shape that only approximates its
  preset therefore breaks the bubble's "area ∝ size" claim in the deck while
  keeping it in the preview — the worst kind of divergence, since the preview
  looks right. `mathMultiply` fails on angled arms (and is redundant with
  `plus` at 3-4pt). `star5` fails on its `hf`/`vf` stretch: the preset widens
  itself 1.05146x and heightens itself 1.10557x so the star fills its box
  (a 5-point star spans 1.902R by 1.809R, and those factors are exactly
  2/1.902 and 2/1.809), which makes PowerPoint's star **16.2% larger in area**
  than an inscribed SVG one. Reproducing that faithfully is possible on paper
  but not verifiable here without a PowerPoint rasteriser, so the set stops at
  the shapes whose geometry is exact: circle, square, diamond, triangle, plus.

- **Waterfall connector re-routing** (drag to skip columns) — the drag is out of
  Office.js reach, and the rendering feature underneath is ill-posed: the
  connector is not an object with authored endpoints, it is a derived assertion
  that "this level carries into the next bar". Re-routing it is truthful only
  where the skipped columns leave the level unchanged — and there it already
  works, via `spacerIndices` and `totalIndices`. Anywhere else the line would
  end in mid-air pointing at a bar it does not touch. What was actually wanted
  shipped instead as `waterfall.detailGroups`.
- **Scatter-on-combo** — a continuous-x scatter over categorical columns needs a
  second x scale that means nothing beside category slots. The coherent reading
  — unconnected marks at category positions — shipped as `type: "marker"`
  (overlay) and already existed as `decorations.barStyle: "dot"` (clustered).
- **Free 2D bubble repulsion** — both axes carry data, so moving a marker in 2D
  corrupts two readings at once with nothing to bound them. The honest version
  shipped as `scatter.spread`: one named axis, hard-capped, cap printed in the
  footnote.
- **Gantt resource capacity vs load** — a per-resource per-week histogram against
  a capacity line. That is a stacked column chart with a `Target` row, both
  shipped; it needs a value axis, which a timeline does not have. Recipe in
  docs/MANUAL.md. (Lane grouping itself shipped as `gantt.lanes`.)
- **Dial / needle gauge** — bullet charts replace it deliberately (Few); low
  deck demand. Note the *semi-circle scorecard* gauge did ship (`doughnut` +
  `pie.semi`); what stays rejected is the dial-with-needle and its threshold
  bands.
- **Sankey / chord / arc** — need curved ribbons; infographic genre.
- **Ridgeline** — stacked density curves; academic register. (The single-column
  `violin` kind shipped at owner request; ridgeline stacking not pursued.)
- **Stream graph** — feasible but editorial aesthetic; no deck demand.
- **Pictogram with icon libraries** — needs the image node above, which is not
  reachable in the live add-in at all (see the first entry), plus an asset
  library. Waffle is the deliberate substitute: it covers the part-to-whole
  genre with square cells, it does not render icons.
- **Histogram as a kind** — the look is `clustered` + `gapWidth: 0`, both
  shipped. If this is ever revisited, auto-binning raw samples into categories
  is the only real gap; the bar geometry is not. (`histogramBins` in
  src/core/format.ts bins over a fixed domain, but nothing derives categories.)
- **Choropleth maps, 3D, drill-down interactivity** — out of scope by design
  (see CLAUDE.md). Tilemap proportional-area cartograms and tilemap drill-down
  fall here too: hard/infeasible.
- **Population pyramid, plain dot chart** — already covered by `butterfly`
  (+ `butterfly.split`) and `decorations.barStyle: "dot"`.
- **Radar vertex markers** — already there: radar emits `marker-*` ellipse
  nodes, which the Office.js renderer draws, so they appear in the live add-in
  too.
- ~~Candlestick / OHLC~~ — shipped as the `candlestick` kind at owner request,
  despite the thin consulting-demand signal.
- ~~Alt text in the headless pptx renderer~~ — **shipped 2026-08-17.** The
  rejection was right about pptxgenjs and stopped one step short: this project
  already hand-patches the generated OOXML (`injectGroupsAndTags`), and it
  writes the very element alt text lives on — the group's `<p:cNvPr>`. One
  attribute, and `scene.desc` now reaches the deck. Image-mode charts take
  pptxgenjs's own `altText`, which it does expose on a picture. Kept here as the
  shape of the mistake rather than as a candidate: a limit of a DEPENDENCY is
  not automatically a limit of the product when the product already
  post-processes that dependency's output.

## 3. Findings from the round archive

What ~290 rounds against the live host have established. Kept because the
finding outlives the fix: each one says what was measured and how, so nobody
re-derives it. Open questions among them are marked as such.

### The first thing the unblocked gate reported is a divergence that is not one — 2026-09-08

Signing round 428's death let `rounds:gate` finish for the first time since that
round, and the first thing it surfaced was:

    1 scenario(s) DIVERGED between slide sizes on the same build:
      stop a run mid-draw — passed at 16:9, failed at 4:3 (31d358f)

The 4:3 detail is *"the slide this drew on could not be deleted, so the aborted
draw is still in the deck"* — a slide-level version of the shape-delete refusal
that has occupied this file all week, which makes it look like a lead.

**It is not a profile difference.** Counted across every round that ran the
scenario and did not skip it:

              ran   any failure   this failure ("could not be deleted")
    16:9       39     10 (26%)                  9 (23%)
    4:3        38     12 (32%)                 10 (26%)

Three points apart on the specific failure, six on any failure, over ~38 runs
each — and it has not happened since **round 383**, 49 rounds ago, on both arms.
So this is one flaky failure landing on the 4:3 side of one build pair, which is
the case `docs/ROUNDS.md` names outright:
*"Sending someone to investigate 4:3 for a scenario that is merely flaky is how
a useful report teaches people to ignore it."*

**What that says about the check rather than the scenario, and it is FIXED.**
`profileDivergence` compared one pair of rounds on one build and knew nothing
about the scenario's own history, so it could not tell "fails only at 4:3" from
"fails at both, and this time the coin landed 4:3". It now carries each
profile's lifetime record with the finding, and the gate prints it underneath:

    stop a run mid-draw — passed at 16:9, failed at 4:3 (31d358f)
      lifetime, this scenario: 16:9 10/39 (26%) · 4:3 12/38 (32%)

Same rule as the check above it — a skipped scenario does not enter the
denominator, because it measured nothing. A reader who sees those two rates
closes the question in a second.

> **THIS SHIPPED INSIDE A COMMIT THAT DOES NOT MENTION IT.** `c9c7c00` is
> titled for two driver fixes — the pane-closed reclassification and the salvage
> clearing its landing spot — and `git add -A` also swept in this divergence
> work, the deck-accumulation correction above, rounds 433 and 434, and the
> duplicate crash record that the salvage bug produced. The message describes
> about half of what the commit contains, which in a repo whose commit messages
> are the record is a real defect and not a tidy-up note. Said here because the
> commit is pushed and cannot be amended.
>
> The duplicate `crashes/2026-09-08T13-42-37-crashed-run.json` is left in place.
> `loadCrashRecords` already drops it as an exact duplicate, and removing a file
> from an append-only archive is the owner's call, not a cleanup.

### The sweep deleted nothing seven times, and that is what killed the tab — 2026-09-08

**NOT FIXED — the fix was built, measured against a real host, and removed.**
Three of my own claims about it are retracted further down, because two of them
were in a report to the owner.

Round 428 tripped the host-death gate: `what a chart kind costs`, first death in
35 runs, ceiling 0. What actually happened is in the crash record, and it is not
about time or density thresholds:

    554.8s  wreckage the host would not resolve  unresolved=1 swept=0
    566.8s  ...  unresolved=1 swept=0          (seven times, one per specimen)
    onSlide: 0, 7, 15, 24, 34, 44, 53, 61

`deleteShapesById` resolves each stray by id through a slide handle a sync old —
the one lookup this host reliably refuses — and gave up when refused, though its
own comment already said a refusal is not "the stray is gone". So
`kindCostSpread` deleted none of its eight specimens, against a docstring
promising "the slide this scenario leaves is the slide it found", occupancy
climbed to 61, and PowerPoint died on the eighth.

**THE REMEDY THAT LOOKED OBVIOUS DOES NOT WORK, AND ROUND 431 SAYS SO.** A
collection-read fallback was added on the asymmetry the rest of the file is
built on — by-id refused, collection honoured — and shipped to the host. It
fired seven times, threw never, recovered **zero**:

    asked  1  listed  8   recovered 0
    asked  1  listed 12   recovered 0
    asked 10  listed 14   recovered 0
    asked  1  listed 31   recovered 0

The shapes are there and the host lists them. It lists them under ids that are
not the ones it returned at creation, so nothing we hold matches. **That is the
defect this file already names two sections down** — "THE RE-READ NEVER MATCHES
OUR IDS", `matched 0` against `listed 9, 10, 16, 17`, with its own note that a
1.5-second settle does not change it. I did not read that section before
building against it.

`reReadRefusedShapes` works because the charts it re-reads were drawn in an
EARLIER round and have settled. A shape drawn seconds ago is a different
population, and my commit message generalised across it. The fallback is
reverted: it cost one extra sync per refused sweep — eight a round, on a host
where #6329 says every sync forces a save — and repaired nothing. Matching by
name or by "the last N on the slide" is how the grouping fallback works and is
not legal for a DELETE.

**AND THE ROUTE THIS FILE PARKS AS "WORTH REOPENING" IS ALSO CLOSED.** The
obvious next move is a binding: `bindings.add` takes the live Shape proxy inside
the batch that created it, with no id round trip and no collection read — the
two things that fail above — and the plumbing exists (`bindTagTarget`,
`bindingShape`, and a `bindingId` already minted per chart). It does not work
here, and the archive is unambiguous. Config tags written, by route, across all
430 rounds:

    group    2696
    created   284
    binding     0   — the key never appears

`settleByBinding` was removed for the same reason after rescuing 0 in 180
rounds, and it used the identical `bindings.getItem(id).getShape()` that a
delete-by-binding would need. What `bindings.add` is kept for is a SIDE EFFECT
at draw time — binding a shape in the drawing batch stabilises its identity, and
`cfg5010` went 8 per round to 0 — not for naming anything afterwards.

**So every route is now measured, and all four are closed:**

    by-id through an aged slide handle   refused; 239 of 412 sweeps, ~93% of the
                                         fresh-shape ones
    collection read, matched by id       lists them under other ids; 0 of 7 in
                                         round 431, and a 1.5s settle does not
                                         change it
    binding retrieval                    0 successes in 430 rounds
    name or "the last N on the slide"    legal for GROUPING, never for a delete —
                                         it would take shapes the caller never
                                         named

There is no fifth idea in this repo. What that leaves is not a fix but a choice
about the scenario: `kindCostSpread` could stop needing the sweep — draw each
specimen on a scratch slide it deletes whole, or accept the litter and say so in
its docstring instead of promising the deck is left as found. Both change what
the instrument does, so they are the owner's.

**SIZED ACROSS THE ARCHIVE, and scoped honestly.** Over 406 rounds:

    sweeps that RESOLVED by id   173 event(s), 286 shape(s) taken
    sweeps the host REFUSED      239 event(s), 248 shape(s) left, 0 taken alongside
    rounds carrying a refusal     33, from round 353 to round 430

So the sweep fails 58% of the time by event, and a failure has never once taken
even a partial bite: 248 shapes stayed on slides the code reported clean.

It is a coin rather than a wall, which is worth knowing before reading the fix
as "the host always refuses". Per round, `what a chart kind costs`: 28 of its 33
rounds refused all eight sweeps, four took some, and **round 407 took all
eight**. Other scenarios sweep by id successfully all the time — `a chart of
rotated shapes` 82 times, `a big chart on a slide of its own` 64. What the
fallback buys is not doing the impossible; it is not depending on the coin.

**But
238 of the 239 refusals belong to `what a chart kind costs`** and one to `stop a
run mid-draw`, so this is not a broad user-facing win. It is by far the heaviest
caller — eight sweeps a round, each on a shape drawn seconds earlier, which is
the recency family Draft D describes: the host will not name what it has just
created.

The two user-facing callers (`app.ts:2694` and `:3655`) both sweep debris from a
redraw that STALLED, which is equally fresh — so they are exposed to the same
refusal, and the single non-kind-cost refusal in the archive did come from
`stop a run mid-draw`. Rare in the archive, and the case where it matters most
when it happens, because the user is looking at the debris.

What a working fix would be worth, then: the scenario keeps its own promise,
strays stop accumulating in the owner's deck, and the occupancy climb that
preceded a tab death goes away. Not "the sweep is broken for users" — and, as
of round 431, not yet worth anything, because there is no working fix.

**AND WHERE THE STRAYS GO — I got this wrong twice before plotting it, so the
correction is below.** The 16:9 deck as each round found it, median top-level shapes:

    rounds   1-200   14        rounds 353-400   16
    rounds 201-300   14        rounds 401-430   22
    rounds 301-352   14

The step is 8, which is one per specimen, and it arrives exactly where the
failed sweeps do. Round 430's inventory names them: slide index 6 (the seventh)
holds **nine** `PowerChart` shapes — one probe chart and eight specimens that
were reported swept and were not.

> **"AND NOTHING ELSE WILL REMOVE THEM" STOOD HERE AND IS WRONG**, corrected
> 2026-09-08 after I had told the owner twice that eight strays were permanently
> in his deck. **The deck is not accumulating.** As each round found it:
>
>     420  45   421  19   422  17   423  22   424  21   425  22   426  22
>     427  22   428  22   429  22   430  22   431  37   432  20   433  21
>
> It sits at about 21, and the fullest slide at 7-9, for thirty rounds. If the
> strays were permanent it would be climbing by eight a round and would now be
> in the thousands. What the deck carries is roughly ONE round's worth at a
> time.
>
> **WHAT CLEARS THEM IS NOT ESTABLISHED, and the tempting answer is already in
> trouble.** The obvious story is the recency defect — unresolvable by id while
> fresh, resolvable once settled, so a later sweep takes them. Round 434's
> entire trace carries three deletion events (`deleted the chart being replaced`
> twice, `swept` once), nowhere near the eight that would need removing. So
> either something outside the trace does it — the driver's `deck-dirty`
> recovery is the candidate, and it has fired 43 times — or the specimens are
> not what the inventory is counting. Named and left open rather than guessed,
> because guessing it is how the sentence above got written in the first place.
>
> The litter is real and transient, not real and permanent. Nothing about the
> owner's deck needs doing, and the "this needs the owner" heading above was the
> wrong call on a number I never plotted.

**RETRACTED — a kill band that was the length of the script.** I reported that
75% of deaths land in 420-900s and read it as a hazard window. It has no control
arm. Rounds that SURVIVED end there just as often:

    survivors (rounds that filed)   322 of 406 = 79.3%   in 420-900s
    crashes                          80 of 106 = 75.5%   in 420-900s

The band is when the battery finishes. Nothing about dying.

And the clock I read it off is not the one I thought. A crash record's step
prefix is `ms since setTracing(true)` — when the PANE was wired, not when the
round started — so it spans attempts. The driver's own header disagrees loudly
on the one pair that can be compared: `crashes/2026-09-07T21-39-03.md` says
**"90s into a round"** while the record beside it ends at **838.7s**.
`src/core/trace.ts` already documents this family of mistake, for probe samples
against trace lines: *"Two time series in one file, on two origins, with nothing
anywhere saying so."* I made the same one a level up. **Use the driver's
`Ns into a round` header, or `last − roundStart` within a record; never the raw
step prefix.**

The distribution is also right-censored and any percentile off it inherits that:
`scripts/round.mjs:2248` stops polling 30 minutes after the Run click, so no
death can be recorded past that however long the host would have survived.

**RETRACTED — "death credit tracks when a scenario runs".** Scenario order is
fixed across every round, so elapsed time and scenario identity are one variable
and no partition of the names is evidence about either. The universal I stated —
"every entry in `FATAL_SCENARIO_RATE` runs after 490s" — is false on three of
eight: `two slides claiming one slot` runs SECOND at 184s, `insert onto a slide
that already has content` fourth at 215s, `same scale across the deck` at 300s.

**AND THE ATTRIBUTION RULE IS NOT UNSTABLE — do not re-open it.** An adversarial
pass reported that death counts swing wildly with the rule (0 vs 41 for `does a
rasterise poison the next draw`) and called pinning it the owner's most urgent
decision. It is already pinned. `fatalScenarios` credits a death only to a
scenario whose `scenario starting` was never closed; the alternative — crediting
whatever ran last — is refuted in that function's own comment, which records
that it once invented three deaths and put two innocent scenarios into the
budget table. Under the shipped rule the archive reads:

    99 records kept · 36 attributed · 63 credited to nothing
    12  a big chart on a slide of its own      9  same scale across the deck
     8  stop a run mid-draw                    2  explode a degraded picture
     1  each: insert onto a slide that already has content, one chart alone on a
        warm deck, two slides claiming one slot, edit the chart the user
        selected, what a chart kind costs

**What does survive, and is worth a person's attention:** `what a chart kind
costs` has run 33 times, reported green 33 times, and **six of those rounds say
"the host stopped answering" in their own detail.** `kindCostVerdict` returns ok
for a stall deliberately and well — a kind is only dead if the host REFUSED it,
not if the host went quiet — but the consequence is that the verdict channel
cannot see this scenario stressing the host. The crash-rate check was the only
thing that could, and it is the one that fired. The gate was right.

### office-js#5022 was WITHDRAWN, and a 1-second cost was resting on it — 2026-09-07

A 100-agent web pass over this project's upstream citations. Most of it
confirmed what the repo already had — #6363 is the live hollow-read bug, #6237
is the date placeholder, #6329's save-per-sync is accepted and unfixed — and one
citation came back materially different.

**`SETTLE_MS = 1_000` cited #5022 as "under investigation", quoting the reporter's
1-2 second timer as "the cheapest known answer".** Checked against the GitHub
API rather than the summary:

- The quote stops one sentence early. The next words are *"But sometimes it
  still struggling and never finish."*
- The issue is **closed as completed, 2024-11-18**, five hours after the
  reporter self-diagnosed: an effect of his own, firing on every selection,
  opened a second `PowerPoint.run` alongside the one in flight. Nothing was
  fixed. `issue-status.mjs` reads that state as "FIXED UPSTREAM — re-read what
  cites it", which is right to flag and wrong about why.
- **This add-in already prevents the cause.** `onSelectionChanged` returns early
  while `hostBusy()` and replays the last event when the work finishes.

And the probe built to watch for the symptom has never seen it.
`picture-then-shape-read` answers `silent` on a hang; across 401 rounds it is
**327 unreadable · 47 yes · 23 threw · 4 no-scratch-slide · 0 silent.** The 327
are #6363's signature, not #5022's. So the second is paid against a failure this
host has never produced.

**NOT REMOVED.** Nothing here measures what removing it would do, and a cost
dropped on a corrected citation is exactly as unevidenced as one kept on it. The
open task is a pair of rounds at `SETTLE_MS = 0` — cheap, and the only thing
that can answer it.

**What the research did NOT find**, recorded so the gap is not mistaken for a
clean bill: nothing at all on `Shape.group` / `ShapeGroup.shapes` on the web
host, and nothing on `untrack()` for PowerPoint proxies — zero surviving claims
on that whole sub-question. And the id-space transition (a fresh slide's
`4123571130#123571113` becoming `323#2528698050`) is undocumented anywhere; the
one candidate source was refuted. Both remain ours to characterise.

### The file insert has two failure modes and they are disjoint — 2026-09-07

**FIXED THE SAME DAY, and the interesting half is the change it killed.**

`insertSlidesFromPptx` measured its result with a raw `slideCount()`. Counted
over the whole archive: **1,641 traced inserts, 14 landing short, 4 landing
OVER.** The four overs are the finding.

Every one of them — rounds **148, 315, 375, 422** — is the same three lines in
the same order:

    handed the host a generated deck    expectedSlides 2, landed 0
    scenario FAILED                     insert on top of an earlier run: GeneralException
    handed the host a generated deck    expectedSlides 2, landed 4

The second insert's `before` count includes two slides the first insert had just
reported as not having arrived. They arrived; the count was read before the deck
caught up. `settledSlideCount` was written for exactly this and its docstring
cites round 297 — an insert on this path returning `landed: 0` for slides that
landed — and the call site was left reading the unsettled count anyway.
`selftest.ts` settles its own count, which is why the scenario judged the right
population while the product function returned zero to everyone else.

Scale of it: `insert on top of an earlier run` ran in **401 of 401 rounds** and
failed 9 times; **4 of those 9 are GeneralExceptions, and all 4 are these.**

**THE RETRY THAT WAS STAGED FOR THIS PATH IS DROPPED**, and the same four rows
are why. It was read as "the insert failed and the next attempt passed". The
insert did not fail. A retry would have handed the host a second copy of a deck
already on the slide — and `landed: 4` is what that looks like from here. Do not
re-propose it from these rows.

**The 45-second timeout is a DIFFERENT event, not this one.** 19 timeouts across
16 rounds (076, 081, 099, 107, 139, 140, 225, 302, 369, 395, 397, 401, 402, 406,
407), every one at exactly the budget, every one on a 2-slide insert. **Not one
is in a round that landed short.** `withTimeoutOrVerify` swallows them and lets
the deck's own count answer, so a timeout here is a slow insert rather than a
failed one. Merging the two counts would produce a 33-event "insert failure
rate" out of two unrelated things.

### GeneralException in the archive is two things, and one of them is closed

A count of the string is a count of nothing. Per round it looks like a constant
~12, which is a property of the harness: four trace strings, some repeating the
word. What they actually are:

- **`resolving the charts' shapes`** — the live one. 274 of the 282 rounds since
  143, exactly twice each, repaired by a re-read all 274 times. See the doomed
  round trip, above.
- **`drawing the chart's shapes`** — 15 rounds, `errorLocation`
  `SlideCollection.getItem` or `getItemAt`, and **every one is in rounds 356-373
  with none since**. A closed era, and every round in it that failed did so on
  `stop a run mid-draw` or `a big chart on a slide of its own`.

### The archive draws on a slide it did not itself add TWICE — 2026-09-06

The denominator that decides what 385 rounds are evidence *of*, and it went
unmeasured until the 5010 question needed it.

`batch issued` names the slide a draw went to. Split against each round's own
`deck.newSlides`, **as of round 413** — every number here grows with the
archive, so it is anchored to a round rather than left to decay:

    on a slide the add-in added that round      8,221 batches
    on a slide the document already had             1 batch
    the `(visible)` sentinel                      883 batches
      of which recoverable from the next batch    866  (863 added, 1 own, 2 ?)
    a slide id absent from that round's final
      inventory — swept before the deck was read  536 batches

**DO NOT QUOTE THE ABSOLUTES.** Two skeptics refuted this entry on 2026-09-07
and both were right about the same thing: these counts drift every round — each
round adds roughly +22 added, +4 sentinel, +12 unplaceable — so three figures
written on three different evenings were sold as one measurement, and the
fourth bucket was left out entirely, which is why the first three did not sum.
Run the script for current values.

**WHAT DOES NOT DRIFT IS THE ONE THAT MATTERS: the second row.** It was 1 at
round 385, 1 at 413, and 1 at 418.

Re-derive with `node scripts/slide-provenance.mjs`, which prints the failures
and this denominator in one output on purpose: the first table is meaningless
without the second, and separating them is how the wrong conclusion was nearly
drawn.

So the sentinel hides nothing: **there is no second arm.** Every scenario adds
a slide and draws on it, because the sweep afterwards has to leave the deck as
it found it. The code says the same thing as the data, which is what makes this
a fact about the harness rather than a quirk of one archive: the two scenarios
that draw without naming a slide — `insert onto a slide that already has
content` and `edit a chart on the visible slide` — both reach their slide by
`probeCharts` first, and the probe chart is on a slide the add-in added earlier
in the same round. "Already has content" means content THIS ROUND put there. That is a reasonable design and it has a consequence nobody had
written down: the round loop validates *"add a slide, draw on it"* and says
almost nothing about *"draw on the slide I already had"* — which is what a user
does.

**WHAT IT COSTS IMMEDIATELY.** 1,369 archived events carried code 5010 at round
413, and 1,361 of them sat on an added slide, none on the document's own. That
looks like confirmation of the 2026-09-06 controlled experiment and is not: with a
denominator of one, every 5010 lands on an added slide because there is nothing
else to land on. The archive cannot separate a slide's PROVENANCE from the
created-proxy rewrite `powerpoint.ts` blames elsewhere — here the two are never
apart. The controlled experiment stays the only evidence that does.

The slide id was on file the whole time, inside `debugInfo.fullStatements`
(`slides.getItem("262#1236456497")`), so this needed no new rounds. The `slides`
field added to `parts list outcome` the same day was recorded with "one field
closes that" beside it. It does not, for this reason, and that docstring has
been corrected rather than left to be quoted later.

**Item 0a is the only measurement on the other side of this line**, and it is
one document.

**THE FIRST VERSION OF THIS PARAGRAPH GAVE A BLOCKER THAT DOES NOT EXIST.** It
said such a scenario would have to skip the sweep, and the sweep is why the
archive is clean. That is wrong: the sweep removes ADDED SLIDES, and a scenario
drawing on the document's own slide would clean up after itself with
`deleteShapesById` — the primitive the wreckage sweep already uses on exactly
this path. The deck would end as it started.

So it is feasible, and cheap:

    read slide 0's shape ids, draw a chart on it, assert it drew and is
    re-editable, then `deleteShapesById` on the ids that appeared.

`slideShapeList` gives the before-and-after lists, so the ids to remove are a
set difference rather than something the draw has to be trusted to report.

**AND THE CLEAN-UP IS NOT OPTIONAL.** `sweepDeck` deletes SLIDES down to one; it
never touches the shapes on the survivor. A scenario that drew on slide 0 and
left them there would hand every later round a deck that starts dirty, and
`deck-dirty` is a refusal — so the scenario would poison the arm it was added
to measure.

WHAT IS ACTUALLY THE OWNER'S CALL is narrower: a twentieth scenario changes the
verdict set, so every cross-build comparison gains a column and the all-green
rate is no longer measuring the same thing either side of it. This file has
gone 12 -> 13 -> 14 -> 15 -> 16 -> 17 -> 18 -> 19 before, so it is a normal decision rather than a
forbidden one.

**AND THE OBVIOUS ALTERNATIVE IS CLOSED**, so nobody spends an hour on it. A
probe looks like the cheaper home — probes run every round and do not touch the
verdict set — but `host-probe.ts` opens with "Nothing here touches the user's
deck. Every probe works on one scratch slide... A probe that would damage a real
slide does not belong here." Drawing on slide 0 is exactly what that rule
forbids. Experiments are out for the same reason. It is the scenario route or
nothing.

**Recommendation: do it.** Inserting a chart onto a slide the user already had
is the single most common thing this add-in exists to do, it is exercised once
in ~9,600 draws, and it is the arm that would let the archive speak to the 5010
question at all.

### The two-master fix improved OUTCOMES without reducing friction — 2026-09-06

§0 shows all-green 4:3 rounds going from 58% to 86% and crashes collapsing. The
natural next sentence is "so the host is refusing less". It is not.

`friction` is a per-scenario DELTA — a snapshot after each scenario minus the
one before it — so it sums and divides cleanly. The conclusion below survived an
adversarial pass; the table under it did not, and was rebuilt on 2026-09-07.

Like-for-like — the 17 scenario names present in BOTH eras — over rounds >= 205,
where every column exists:

    era    scenarios   errors  idRefusals  generalExceptions | repaired
    PRE         2472    0.240       0.074              0.078 |    0.363
    post         714    0.246       0.070              0.057 |    0.403

Flat — **and "flat" does not survive controlling for MIX, which is the third
revision this entry has needed.** Restricting to shared NAMES does not equalise
how often each is run. Post runs all 17 exactly 45 times; PRE runs 13 of them
172 times and the other four 164, 36, 24 and 12 times, because they were added
mid-archive. Weighting every scenario equally instead of pooling instances:

    era    errors  idRefusals  generalExceptions | repaired
    PRE     0.294       0.063              0.134 |    0.325
    post    0.247       0.071              0.058 |    0.403

Two of the four columns change direction. Errors FELL 16% once mix is
controlled, where the pooled table says they rose slightly; `generalExceptions`
fell 57%. (The post row is identical either way, which is the check: post is
already balanced at 45 apiece.)

**So the honest conclusion is weaker than the one first written here.** It is
not "friction is flat"; it is **"the answer depends on the weighting, and a
comparison that flips under a reasonable reweighting was never robust enough to
carry a claim"**. What survives both readings: nothing in this table shows
friction getting materially WORSE after the fix, and the product's outcomes
improved regardless. **The product absorbs at least as much host misbehaviour as
before and now survives it** — which is the durable half.

**`repaired` IS ON THE OTHER SIDE OF THE BAR, and putting it in a friction
table was a category error.** `hostFriction.reReadsRepaired += pending.length -
retry.length` counts the charts that RESOLVED on a retry pass, and its own
docstring calls it "charts whose re-read was short or empty FIRST and complete
after a pause ... the number that says whether `REREAD_RETRY_MS` is buying
anything". Up is GOOD. An earlier version of this entry read 0.36 -> 0.56 as
"the repair path carrying more" risk; it is the repair path working more often.

**THE PUBLISHED VERSION OF THIS TABLE WAS WRONG THREE WAYS**, each found by a
skeptic and each checked here before it was believed. They are kept because the
shapes recur.

1. **The window's stated reason was false for most of the table.** It said
   rounds < 205 were excluded "because the field did not exist". Measured first
   appearance: `errors` 023, `idRefusals` 023, `generalExceptions` 023,
   `emptyReReads` 023 — only `reReadsRepaired` and its two neighbours start at
   205. Three of four columns discarded 2,292 scenario instances for a reason
   that did not apply to them.
2. **`repaired` 0.36 -> 0.56 was a scenario-MIX artifact.** `what a chart kind
   costs` is post-only and supplies 162 of 450 post repairs — 36% of them — from
   21 of 773 instances, 2.7%. Like-for-like it is 0.363 -> 0.403.
3. **Per-scenario fixes the COUNT and not the MIX.** Normalising by scenario
   count still compares two different sets of scenarios.

AND THE WINDOW HIDES THE REAL IMPROVEMENT. Same 17 scenarios, window removed:

    era    scenarios   errors  idRefusals
    PRE         4764    0.404       0.265
    post         714    0.246       0.070

Friction fell a long way across this archive's history — and almost all of it
happened BEFORE the two-master fix, between rounds 023 and 204. Across the fix
itself it is flat. Both statements are true and only the second is about the fix.

**`generalExceptions` STILL SHOULD NOT BE READ AS PRODUCT.** `selftest.ts`
records that `explode a degraded picture` throws exactly one per round — "a
constant, not a signal" — so that column is largely a constant over a growing
scenario list.

TWO EARLIER READINGS, BOTH WRONG, kept for the shape: per round over the whole
archive `idRefusals` appeared to FALL 58%, an artifact of 180 rounds counting as
zero; per round over rounds that carried it, the same number appeared to RISE,
an artifact of the scenario list growing. Four opposite readings of one dataset
before the denominator was right.

### Half the parts-list events are no-ops, and 41% of host deaths are one scan — 2026-09-07

Two instrument findings from the same afternoon, both fixable, both found while
attacking something else.

**About HALF of all `parts list outcome` events carried `charts: 0`** — 3,012 of
6,275 through round 421, and 3,000 of 6,179 when this was written, which is the
same ratio and a reminder that the integers here drift every round. Note the
denominator now includes the rounds cited below as evidence, which the first
version of this paragraph excluded while quoting them. An outcome
traced for an empty item list. Nearly half the population is noise, and it
inflates every denominator taken off this event, including the "640 carry the
id read-back threw" that started an evening's enquiry. A regression with a
boundary: rounds 142-145 have none, round 146 (build `2521d23`) has three, and
the round total stayed at 21, so three events CONVERTED rather than appeared.
Guarded in `tracePartsOutcome`; no test kills that mutant and its docstring says
why, so a round is what confirms it.

**A SATISFYING CAUSAL STORY DIED HERE.** The empty-collection defect measured
on 2026-09-07 looked like it explained this repo's oldest gap — `withParts` 0 of
1,370 — because the parts list needs a COLLECTION read while tagging uses
creation handles. It fits everything and it is wrong. Split the dominant exit
("no loose chart had siblings", 5,855 events): 3,000 have no charts at all, and
the other 2,855 are single charts that were GROUPED, which `powerpoint.ts` says
in as many words have no parts tag by design. Grouping succeeding is the good
outcome.

**"AND IT IS THE WHOLE EXPLANATION" WAS TOO STRONG, corrected within hours by
attacking it.** That exit is 2,855 of 6,179 parts-outcome events — 46%, not all
of them — and the 3,000 charts:0 events explain nothing whatever, being the
no-op regression guarded two paragraphs up. **324 events took the two exits that
DID want a list** (`the id read-back threw` 309, `read the ids back` 15), and
grouping says nothing about those: those charts were loose. So grouping explains
the dominant EXIT, not `withParts` being 0, which stays open. The refutation of
the empty-collection story survives; its replacement over-reached in the same
breath.

**AND 41% OF HOST DEATHS HAPPEN IN ONE SCAN — WITHDRAWN THE SAME DAY. Read the
correction below the table, not the table.** Every crash record carries its own
`steps`; the last one before the record was written:

    41  pane    collecting deck evidence — scanning
    15  error   drawing the chart's shapes
     6  draw    parts list outcome
     4  probe   asking
     4  draw    batch issued

    last-step time: p25 468s   p50 529s   p90 761s

`crashes/README.md` already says that scan runs AFTER every verdict is in. So
the largest single cause of lost rounds is a read taken at maximum session
fatigue, after the round's measurements are complete.

**AND IT IS CHEAP, WHICH IS WHAT MAKES THE NUMBER DAMNING.** Measured over the
last 60 rounds where the scan completed:

    scan duration            p25 6.8s   p50 7.6s   p90 11.6s
    scan as a share of the
      round's own work        p50 1.4%   p90 1.9%

**1.4% of the time, 41% of the deaths** — around a thirty-fold
over-representation. This is not a heavy operation that sometimes fails; it is a
seven-second read that is the last thing in flight for four host deaths in ten.

(The duration is measured on scans that SURVIVED — a scan that killed the host
leaves no round file. So this is the cost when it works, and the true cost is
worse.)

**THE RECOMMENDATION THAT STOOD HERE IS WITHDRAWN.** It said making the scan
cheaper would cut lost rounds a long way. Four things kill it, all found by
attacking my own correction a few hours after writing it and all verified
independently.

**1. The number was already in this repo TWICE, and already acted on.**
`scripts/round.mjs` says "41 of the 75 crash records on file — **55%** — end at
the same trace line", names it "the POST-ROUND inventory pass, which runs after
every verdict is already banked", and records that the pane **already bounds it
with `DECK_EVIDENCE_TIMEOUT_MS`, 45 seconds, "exactly so a host dying in there
cannot take the round with it"**. `powerpoint.ts` carries the same 41 against a
denominator of 60. Mine restated that same frozen 41 against 99 — publishing the
repo's own 55% as 41% by growing the denominator, which makes the crash profile
look more diffuse than this project's own reading of it, and then recommending a
mitigation that was built weeks ago.

**2. It is a frozen numerator over a growing archive.** Every one of the 41 is
dated 2026-08-27 to 2026-09-01. **33 crash records since contain the line zero
times**, while the trace still exists in `app.ts`. The count was 36 at 60
records and has been 41 ever since.

**3. The 08-27 start is when the LINE began, not the deaths.** `b5c534a` warns
of exactly this trap for the cross-arm comparison and I walked into the
time-series version of it.

**4. The timing was the wrong interval.** "p50 7.6s, 1.4% of a round" measured
scan-start to the round's LAST trace — everything after the scan, not the
awaited call that kills. The call at risk is the gap to the NEXT trace:

    rounds <= 340, the scan-death era    p50 1390ms   p90 1612ms
    rounds  > 340, what I measured       p50    0ms   p90 1636ms

The transition is sharp — round 359 has a 2285ms gap, 360 onward reads 0-1ms.
**The scan's first awaited call stopped costing anything, which is the most
plausible reason deaths stopped landing in it.** What made it free is NOT
established: the three commits between those builds are about a probe returning
undefined, a slide-profile fix and an instrumentation field, and none of them
obviously touches the scan. Saying "the scan was fixed on 2026-09-02" would be a
fifth guess dressed as a finding.

**5. And the step is last BY CONSTRUCTION.** It is logged one statement before
the phase's first host call, so any hang in that call leaves it as the final
line whatever the cause. The 41 then sat silent a median of 173 seconds after
it — a host hang, not a death sampled inside a short scan. Attributing risk to a
step chosen because it precedes the hang is the trigger-that-cannot-fail shape,
one level up from the test I threw away this morning.

**What survives is worth keeping**: within the window where the line exists,
every crash record that contains it ENDS at it. That concentration is real and
this project found it first. It is not current, the mitigation already exists,
and nothing should be changed on the strength of it.

**THE HONEST SUMMARY OF THIS WHOLE ENTRY**: I rediscovered a known finding,
restated it with a worse denominator, measured the wrong interval by a factor of
thousands, read a frozen count as a live rate, and recommended a fix that had
already shipped. It took a second adversarial pass to notice, and the reason it
survived the first one is that it sounded like a discovery.

A WRONG TURN AVOIDED, recorded because it was nearly taken: the first attempt
attributed crashes to scenarios by their median start time in RECENT rounds,
which would have blamed three scenarios that mostly did not exist when those
crashes happened. The records name what was in flight themselves.

### A round's verdicts are not independent, and all-green is a session statistic — 2026-09-07

The general form of the pair below, and it changes how every rate in §0 should
be read.

    rounds                                              398
    rounds with at least one failure                    131
    total scenario failures                             263
    failures per round      0:267  1:56  2:62  3:8  4:1  12:1  13:1  15:2

    of the 131 failing rounds, 75 had TWO OR MORE          57%
    expected under independence, using each scenario's
      own measured failure rate                            28%

**EVERYTHING ABOVE COUNTS SKIPS AS FAILURES, AND 43% OF THEM ARE SKIPS.** A
selftest entry carries a `skipped` flag beside `ok`, and of the 263 not-ok
entries **113 are skips** — a precondition that was not there, not a product
that failed. The detail text never says "skip", which is how both an adversarial
pass and I missed it.

**AND CALLING THAT A DISTORTION WAS ITSELF TOO STRONG**, corrected within the
hour by reading the instrument instead of only the archive. `rounds-gate.mjs`
already separates them — `const ran = s.filter((x) => !x.skipped).length` — and
states the stance in as many words: **"A skipped scenario is not a passing
one."** So counting a skipped round as not-all-green is a deliberate,
conservative choice this project already made, not an accident.

What is genuinely wrong is only the LABEL. "All-green" reads as "nothing
failed", and it means "everything ran AND passed". Both numbers are worth
having and they differ by eleven points:

    all-green counting skips as failures      269 of 400   67.3%
    all-green ignoring skips                  315 of 400   78.8%
    scenario pass rate, skips as failures                  95.35%
    scenario pass rate, skips ignored                      97.35%

**Eleven and a half points of the all-green rate is skipped scenarios**, and
`test/backlog-health-table.test.ts` pins the stricter definition, so the rows
are internally consistent — it is the word "all-green" that carries more than it
should. **Quote the number with the definition attached, or quote both.**

WITH SKIPS SEPARATED THE STRUCTURE IS DIFFERENT AND MUCH SHARPER:

    genuine failures per round   0:315   1:23   2:60   3:1   4:1

Two failures is nearly THREE TIMES commoner than one. Genuine failures arrive in
pairs, and four pairs account for all sixty rounds:

    43   explode a degraded picture  +  same scale across the deck
     9   a big chart on a slide of its own  +  stop a run mid-draw
     6   insert on top of an earlier run  +  two slides claiming one slot
     2   an update follows a moved chart  +  same scale across the deck

**One pair is 43 of those 60 rounds — 86 of the archive's 150 genuine failures,
57% of everything that has ever actually failed here.** This file already knew
those two are the scenarios that fail; it did not know they fail *together*.

AND THE FOUR "CATASTROPHIC" ROUNDS ARE CASCADES OF SKIPS, NOT FAILURES —
which is the correction to what this section said an hour ago:

    round 287   13 of 14 not-ok    9 say "no probe chart"   ALL 40 ARE SKIPS
    round 360   15 of 16 not-ok   11 say "no probe chart"
    round 361   15 of 16 not-ok   11 say "no probe chart"
    round 384   12 of 18 not-ok    9 say "no probe chart"

Ten scenarios need a probe chart to work on — `edit the chart the user
selected`, `does a rasterise poison the next draw`, `edit a chart on the visible
slide`, `an update follows a moved chart`, `one chart alone on a warm deck` and
the rest — and they get it from the run's opening, the same file-insert path as
the pair below. **When the opening fails to place probe charts, every one of
them SKIPS**, correctly and by design, and the archive records ten not-ok
entries for one event. Every one of the 40 "no probe chart" entries carries
`skipped: true`; not one is a failure.

    those 4 rounds are 1.0% of rounds and 21% of all NOT-OK entries
    "no probe chart" is 40 of 263 not-ok entries — 15%, and 0 of 150 failures

    all not-ok as failures   399 rounds  5,639 runs  263  pass 95.34%
    minus those 4 rounds     395 rounds  5,575 runs  208  pass 96.27%

Four rounds move the archive's apparent pass rate by nearly a point **without
containing a single product failure between them.** That is what counting a
skipped precondition as evidence does to a mean — and it is why the two
readings above differ by eleven points.

The pairwise view says the same thing. Ranked by excess over independence, the
top pairs run 39x to 100x, and they are not random — `edit the chart the user
selected`, `does a rasterise poison the next draw`, `edit a chart on the visible
slide`, `which selection call wedges the host`, `an update follows a moved
chart` and `one chart alone on a warm deck` co-fail with each other repeatedly.

**SO ALL-GREEN IS AS MUCH A READING OF THE SESSION AS OF THE BUILD.** A round is
not 19 independent trials; when it goes wrong it tends to go wrong in several
places at once. Two consequences worth carrying:

- Any confidence interval computed as 19 independent Bernoulli trials is too
  narrow, including anything anyone derives from the §0 table.
- The all-green rate and the "231 of 263 rounds needed a repair" figure earlier
  in this file are measuring overlapping things: session health shows up in
  both.

Nothing here says the product is better or worse than the table reports. It says
the table's denominator is softer than it looks, in the same family as the four
denominator traps above.

### Two "flaky scenarios" are one flaky dependency — 2026-09-07, p = 1e-13

Round 422 came in 17 of 19 and the rounds gate stopped the cycle on its one
fatal check, a scenario that had been passing. Reading it produced something
better than the regression it was reporting.

**The two that fell are not independent. They almost never fail apart.**

    rounds running both                                        398
    `insert on top of an earlier run` failed                     9   2.26%
    `two slides claiming one slot` failed                       11   2.76%
    BOTH failed in the same round                                8
    expected under independence                               0.25

    2x2: both 8 | only insertTwice 1 | only duplicateSlot 3 | neither 386
    Fisher's exact, one-tailed                          p = 1.0e-13

Of the nine rounds where the first failed, **eight** had the second fail too.
The co-failing rounds are 148, 287, 297, 315, 360, 361, 375 and 422 — and 360
and 361 are the pair whose `deck grew by NaN` verdicts account for eight of the
nine "chart losses" corrected elsewhere in this file.

**WHAT THEY SHARE IS `insertSlidesFromBase64`.** Both scenarios build a probe
deck and insert it through `insertSlidesFromPptx`; neither does anything else in
common. Round 422's failure is that call throwing `GeneralException` at
`errorLocation: Microsoft.Office.PrivateApiService`, with the base64 payload in
the statement.

So the archive has been carrying this as two scenarios that each fail about 2%
of the time. It is **one shared dependency taking both dependents down at once**
— a different fact, with a different fix, and a much better place to look.

**"ONE DEPENDENCY THAT FAILS ~2%" WAS TOO TIDY, narrowed an hour later by
reading the eight rather than the aggregate.** They are not one exception:

    148  GeneralException          the insert throwing
    315  GeneralException
    375  GeneralException
    422  GeneralException
    287  Failed to fetch dynamically imported module   a stale deploy, not the host
    297  the deck scan could not see the whole deck    a blind scan
    360  deck grew by NaN                              instrument artifact
    361  deck grew by NaN                              instrument artifact

**Four of the eight are `insertSlidesFromBase64` throwing.** The other four are
three unrelated causes. What makes all eight co-fail is not one call but the
pair's position: they are the first two scenarios in the run and share the whole
early deck path — build a probe deck, insert it, scan it back. Anything that
breaks that breaks both.

The correlation stands at p = 1e-13 and the diagnosis narrows: **the first two
scenarios are not two independent readings of the product, they are one reading
of whether the round's opening worked.** Which also means the archive has ~2%
fewer independent scenario samples than its verdict count implies.

A caution against the obvious next step: co-failing rounds average 3.8 slides
against 7.0, and 1,325 syncs against 2,046. That is the round not getting far —
an effect of failing early, not a cause of it, and it would read as a tidy
"smaller decks fail more" if taken at face value.

NOT MY CHANGES, and the gate said so before I could wonder: "THE SHIPPED BUNDLE
IS UNCHANGED since the previous round at this profile — nothing under `src/`
differs between the two builds." Verified independently — `git diff bc88a7c
1334194 -- src/` is empty; only docs, rounds and crash records moved.

### `withParts` is 0 on a pool that excludes its own successes — 2026-09-07

The oldest unexplained zero in this repo, and the reason it may not mean what it
has been quoted to mean. Found by a skeptic attacking a correction, verified
here.

`powerpoint.ts` increments `withParts` inside the churn block — and that block
sits AFTER the `continue` taken whenever `tryInPlaceUpdate` succeeds. So the
counter only ever sees charts whose in-place update was REFUSED and which fell
through to delete-and-redraw. Measured across the archive:

    in-place updates that succeeded, skipping the counter   3,048
    declined, falling through to where it counts            2,002
    churn events carrying `withParts`                       1,406
    their sum                                                   0

**AND IT IS WORSE THAN "EXCLUDES ITS SUCCESSES" — THE POOL IS SELECTED ON THE
VERY PROPERTY BEING MEASURED.** `tryInPlaceUpdate` traces why it refuses, and
those refusals are exactly the population that reaches the counter. Across the
archive:

    1,197  the chart has no parts list, so its nodes cannot be mapped to shapes
      335  the chart has no parts list and no readable group members
      392  this update draws a picture, not in the scene the differ compares
       61  too much of the chart changed to be worth writing shape by shape
       12  the chart carries no scene fingerprint — drawn by an older build
        5  the parts list does not match the scene one for one
    -----
    2,002  refusals, of which 1,532 (77%) are refused BECAUSE there is no
           parts list

So `withParts` asks "did this chart have a parts list?" of a population three
quarters of which is there precisely because it did not. **The zero is very
largely circular**, and cannot be read as evidence about how often lists exist.

ONE LOOSE END, NAMED RATHER THAN TIDIED: five refusals say "the parts list does
not match the scene one for one" — those charts HAD a list. If they reach the
churn block they should make `withParts` at least 5, and it is 0. Either the
`watching` guard excluded them or `it.target.partIds` is empty where that
refusal reads a list from elsewhere. Not chased down; recorded because a strict
zero with five known counterexamples upstream is a thread someone should pull.

WHAT THIS DOES AND DOES NOT SHOW. It does not show that parts lists are read
back — nothing here does. It shows that **0 of 1,406 is not evidence that they
are not**, because two thirds of the population never reaches the counter. The
sibling counter on the production side, `gotPartsList`, is NOT zero — 35 events
— so lists do get written; they have simply never been observed on a path that
excludes most of its own traffic.

Not fixed here. Moving the counter above the `continue` changes what the field
means across 1,406 archived events, and that is a decision about the archive.

### Round duration is a fourth denominator trap — 2026-09-07

    per ROUND (selftest total)   PRE p50 411s  ->  post p50 604s   (+47%)
    per SCENARIO                 PRE p50 17.5s ->  post p50 15.1s  (-14%)
                                 PRE p90 62.9s ->  post p90 87.8s

The +47% is the scenario list growing 12 -> 19 (12 at round 023, 19 at 398). Per scenario the typical case
got slightly faster and the tail got heavier. Nobody has quoted the +47% yet;
this exists so nobody starts.

### 231 of 263 rounds needed a repair before they could start — 2026-09-06

`driverRun.recovered` has been recorded since round ~150 and had never been
totalled. **As of round 414** — 267 rounds by 418, and this ratio is the point
rather than the integers — of the 263 rounds carrying the field, **231 needed at
least one recovery** before the driver would call the setup ready:

    not-ready:pane-closed                            109
    crashed                                           71
    not-ready:pane-stale                              51
    not-ready:deck-dirty+pane-closed                  35
    not-ready:host-silent+pane-stale                  27
    not-ready:deck-missing+host-silent+pane-stale     12
    not-ready:browser-gone                         8 + 8

Attempts per round say the same thing more sharply — **32 of 263 rounds started
on the first try**:

    attempts   1    2    3   4   5   6   7
    rounds    32  163   49   9   5   1   4

READ IT AS THE LOOP WORKING, not as the host being sick. The pane closes
between rounds as a matter of course, and repairing that is what `recover` is
for. What the number changes is how a clean round should be read: **88% of the
archive is a measurement of a session the driver had just repaired**, which is
the same caveat §0 already states for restarts, now with a denominator.

It also says tonight's browser trouble is not new — `browser-gone` is in the
archive sixteen times. What was new on 2026-09-06 is that the profile stayed
locked afterwards, which the driver could not see and now can.

### The probe has been blind on GROUPS for the whole archive — found 2026-08-16

Found and fixed 2026-08-16; the probe reads groups now. See git.

### A SINGLE-BATCH CHART CANNOT GROUP ON THIS HOST, AND THAT IS OUR BUG — 2026-08-16

`npm run rounds` prints it under **DID THE CHART SPAN BATCHES**. Pooled over 41
rounds and 537 draws:

    spanned batches   452 draw(s),  353 grouped = 78%
    one batch only    214 draw(s),   49 grouped = 23%

**353 of 452 against 49 of 214.** (First reported here as 333 of 333 — 100% —
which was a bug in the pooling function and not a fact: it searched a fixed few
entries for the verdict and did not count `not grouping` at all, so every honest
decline was dropped from both arms. The separation survived the correction; the
absolutes did not. On builds carrying the settled retry the multi arm really is
10 of 10.) The sharpest separation in the archive, and
unlike the fresh-slide split it is not a fact about the host — it is a
consequence of one line of ours.

**The mechanism, end to end, every link measured:**

1. `refreshShapes` is set from `spansBatches(created, opts)`, so **only a
   multi-batch chart gets the pre-grouping re-read.**
2. A single-batch chart therefore hands `addGroup` the raw `created` proxies.
3. This host refuses those: `InvalidParam passed to GetItem(id)`, 5010, at
   `grouping the chart's shapes`.
4. The failed group takes the tag with it — `target.tags` comes back
   **undefined**, and it is *always* this: **155 of 155 across the whole archive,
   every one immediately after a 5010 group, zero exceptions.**

So `tags-undefined` is not an independent fault and never was. It is the second
half of a failed group, and the TAG FAULTS table's columns are not independent —
one refused group produces `group-5010`, `tags-undefined`, `no-queue` and
`tagging-failed`, four counts for one event.

**THE RASTERISE WAS A RED HERRING.** It first looked like draws after a rasterise
group 22% of the time against 93% for everything else. Splitting the arms by
batch count collapses it to **22% against 27% — nothing**. The scenarios that
rasterise simply draw small charts. This is the same confound `poolEveryDraw`
already warns about, met from the other direction, and it is written down because
it took a deliberate control to kill and would otherwise be re-found.

**THE FIX, and why it is not in this commit.** Refresh before grouping for every
groupable chart rather than only multi-batch ones. One line — `refreshShapes`
stops being gated on `spansBatches()`.

The original reason for the gate is in the code at the `needsRefresh` call site:
asking for a re-read a chart does not need was "a way to LOSE a group, not gain
one", because this host answers a re-read short or empty. **That objection is
much weaker now the settled retry exists** — a short or empty answer is asked
again after 1.5s before anything is decided.

**Stake the prediction before the round that tests it:** single-batch grouping
moves off 24% toward the multi-batch 100%, `tags-undefined` falls with it because
it has no other cause, and `the chart is actually visible` starts reporting a
chart that carries its config. **Run it as a pair**, and watch for the cost: an
extra sync per single-batch chart on the live insert path, which is the path a
user waits on.

**Owner's call, because it changes every insert.** The evidence is as strong as
this project gets, and the risk is a round-trip on the interactive path.


### THE RE-READ NEVER MATCHES OUR IDS — named precisely by rounds 068/069

**The clearest statement of the grouping defect this project has, and it took one
instrumented field to get.**

    withOwnId 7 of 7, 9 of 9   our handles are fine — the ids we hold are real
    listed    9, 10, 16, 17    the host named plenty of shapes...
              1, 1, 1, 1, 31   ...and in the next round, ONE
    matched   0                none of them ours, in either round

So the host lists a freshly drawn slide's shapes **under ids that are not the
ones it returned at creation**, and how many it lists at all is moody. Every one
of those traces carries `afterRetry: true`: it read that way *after* a
1.5-second settle, so waiting longer is not the fix.

**Consequence.** Grouping for these charts now rests entirely on the positional
fallback — "the last N on the slide" — which is a guess, is only legal when
nothing matched, and needs `listed >= drawn`. That is why the same build grouped
4 of 5 in one round and 1 of 5 in the next.

**The internal parallel says what kind of bug this is.** Scratch slides came back
as `4123571114#123571113` while the deck listed `287#62081387` — a freshly added
SLIDE reporting an id in a namespace the deck will not answer to. This is that,
one level down, on shapes.

#### The route worth reopening: a BINDING, measured in production

`bindings.add` takes the live Shape proxy inside the batch that created it — **no
id round trip and no collection read**, which are precisely the two things
failing above. Microsoft's own shape-binding documentation surfaced unprompted in
two of three web searches on this symptom.

This repo parked bindings as **unanswerable, not refuted**: `binding-names-shape-
later` was asked eight times in eight rounds and never once reached its own
question, because a shape probe needs the scratch slide resolved more times than
this host allows. **That is a fact about the probe, not about bindings** — and it
is the same shape as the six questions found 0-for-41.

The settled retry proved the alternative: ask in PRODUCTION. `repaired N` on a
verdict line settled in one evening what twelve rounds of probing could not.

**Plan it before building it** (the four points): the defect is that no durable
handle survives from creation to grouping; the seam is `groupAndTagAll`'s member
choice, beside `chooseGroupMembers`; what proves it is a `by: "binding"` reading
on the grouping trace against today's `ids`/`created`/positional; and it must not
touch the id path that already works for multi-batch charts, which group at 100%
on current builds and must stay that way.

**Owner-gated on cost, not on doubt**: bindings are PowerPointApi 1.8 against a
manifest pinned at 1.4, so it needs a `supports("1.8")` gate like grouping
already has, and it must degrade to today's behaviour on an older host.

### `same scale` PASSES — and the next work is where the failures moved to

**Rounds 070/071/072 on `01f3607`: 12 of 12 scenarios, `same scale` 8 of 8, three
times.** After 35 rounds of failing. Full numbers in the journal.

**The cause is not the code that was written for it.** The binding was added to
give the SETTLE a durable handle; the settle route has never executed (`bind 0/0`
in all three). What it did instead was upstream: `bindings.add` on the live proxy
in the drawing batch stabilises the shape's identity, so the pre-grouping re-read
stops listing shapes under ids we cannot match, and the config tag write lands.
`cfg5010` is 0 in all three rounds against 8 before.

**Reframe the item accordingly.** This is not "the settle got a handle", it is
**"binding a shape makes its identity resolvable"** — and if that is the real
mechanism, the settle-by-binding route is dead weight riding along with a
one-line side effect. Worth deciding deliberately rather than leaving both.

#### 1. ~~THE ORIGIN TAG IS NOW THE TOP FAILURE, and nothing can check it~~ — SHIPPED

Both halves are done and this entry outlived them by two days. The scenario is
`an update follows a moved chart` (`dragThenUpdate` in `src/taskpane/selftest.ts`,
in the battery's list): it moves a chart PROGRAMMATICALLY, confirms the move
landed, updates it, and asserts the redraw follows the delta rather than snapping
back to where the chart was inserted. No selection call anywhere in it.

And the defect it was written for is closed — see **THE ORIGIN TAG IS FIXED**
below: rounds 075/076, 13 of 13 scenarios twice, `origin5010` 0 both times, after
routing that one tag write through the chart's binding.

What is still owed is the part a scenario cannot do: a REAL mouse drag, by a
human. That is in the owner-only list at the end of this file and nowhere else.

#### 2. ~~NOTHING ENFORCES A ROUND RESULT~~ — SHIPPED

`scripts/rounds-gate.mjs` (`npm run rounds:gate`) is that gate: it reads the
archive, and exits 1 when a scenario that passed the previous three rounds of the
SAME profile fails in the newest one. `scenarioRegressions` in `scripts/triage.mjs`
is the decision, kept pure; a skipped scenario counts as "did not measure" rather
than as a fall, and exit 2 means the gate could not do its job at all.

Two limits worth knowing rather than re-proposing. It cannot be a CI check — CI
has no rounds — so it runs after archiving, and `cycle.mjs` acts on its exit
code. And it reads the pass/fail FLAG, so a scenario that goes on passing while
its numbers drift is not what it watches; `npm run rounds` is still where a
person reads the pooled counts.

This entry stood for a day after the gate landed, claiming the result had no
guard. Read the scripts before quoting a gap here.

#### 3. THE LOOP STILL NEEDS SIX MANUAL BROWSER STEPS A ROUND

Download the log, clean the deck, reopen the pane after a reload, verify the
toggles. This is throughput, but it is also correctness: hand-driven cleanup is
what took a deck to **0 slides** on 2026-08-16 and produced `slide 1 REFUSED`.
`cleanDeckScript` already exists — readiness REFUSES on a dirty deck instead of
using it.

#### 4. OWNER-ONLY, and neither has ever happened

- **Desktop PowerPoint.** Every round in this archive is the web host.
  `PUBLISHING.md` asks for desktop at least once.
- **A 4:3 deck**, same.
- **A real mouse drag.** The native-editability premise the product rests on has
  never been exercised by a human hand — and it is the thing item 1 can only
  partially substitute for.

### HALF THE CHARTS CANNOT FOLLOW A DRAG, and a passing scenario was hiding it

`npm run rounds` prints it now, under **CHARTS THAT CANNOT FOLLOW A DRAG**.

    rounds 073/074:  origin tag lost on 9 of 19 and 8 of 17 charts
    the scenario:    `an update follows a moved chart` PASSED in both

The scenario tests ONE chart and picks the least loaded. Roughly half the
population loses its origin tag in the same round, and every one of those would
snap back to where it was inserted rather than following the user's drag.

**The same shape this project keeps finding.** `does a rasterise poison the next
draw` counted only its own four draws while the round held 195; the fresh-slide
split sat unqueried for eleven rounds. **A scenario samples; a pooled count does
not**, and a green verdict over a sample is the most expensive kind of quiet.

**What it costs the user, stated precisely.** The chart is re-editable and its
config is intact — `cfg5010` is 0 across every post-binding round. What is lost
is only the drag delta. So this is a real defect and NOT a data-loss one, which
is why it can sit behind an item that is.

**Where the failure came from.** It moved here. Before the binding change the
config tag took the 5010s (8 in round 069) and the origin tag none; now the
config tag takes none and the origin tag takes 8-9 a round. The second write is
the one that fails, and it fails on the handle the first write just succeeded
through — which is worth a probe question of its own, since nothing currently
asks whether a shape accepts a SECOND tag in a later sync.

**No upstream twin.** Searched 2026-08-16: no office-js issue reports a second
`tags.add` on the same shape failing where the first succeeded. Two adjacent
ones are worth knowing and neither bites us — #6079 (tags forced uppercase, and
case-sensitive on web only; every tag here is already uppercase) and #3784
(shape tags lost on cut/paste on web). Recorded so the absence is not
rediscovered.

**No rate is printed, deliberately.** Only failures are traced — a successful
origin write says nothing — so there is no honest denominator, and inventing one
would be the kind of number this file has already had to correct once.

### THE ORIGIN TAG IS FIXED — rounds 075/076, and the binding is proven

Fixed and proven by rounds 075/076. See git.

### PUTTING 4:3 INTO THE NIGHTLY RUNS — four blockers — ALL CLEARED 2026-08-16

All four blockers cleared 2026-08-16. The 4:3 leg is `cyclePlan`'s third; see docs/ROUNDS.md.

### THE INSTRUMENTS WERE THE PROBLEM — rounds 082-087, 2026-08-16/17

Six consecutive 13/13 rounds, zero `UnexpectedError`. **Two product defects were
found across the whole stretch. Nine were in the reporting**, five of them
introduced the same day they were found.

That ratio is the finding. The product has been stable for six rounds; nearly
everything that went wrong was a number about the product rather than the product
itself. **A report that lies is worse than a crash**, because a crash stops the
night and a wrong number redirects a week.

#### The four ways one instrument was wrong

The orphan reading — "does an update strand the rest of an ungrouped chart" —
produced four confident wrong answers in two days:

| | what it reported | what was true |
| --- | --- | --- |
| mismatched units | 283 shapes stranded | its own line said `before: 3, after: 3` |
| an identity | 0, always | the two terms summed to zero by construction |
| one stale read | 92 shapes grew | the deck showed one grouped chart per slide |
| two agreeing stale reads | growth 23, `settled` | both reads fell inside one lag |

**Every one was caught by `deck.inventory`** — a second source already in the
round file, taken at end of round, long after any host lag. It has never been
wrong. It is now cross-checked on every reading and disagreements are COUNTED,
because an instrument's own error rate belongs in its own report.

**The rule that comes out of this**: never quote a number until a second,
independent measurement in the same data agrees. Not a re-run of the same
instrument — two reads of one lagging source agree perfectly.

#### What is left open here

- **The stranding question is at three observations, all zero growth.** Three is
  not five. It needs rounds that fail to group something; 082-086 grouped
  essentially everything and could not answer either way.
- **`reading back an ungrouped chart's shape ids`** is still the one live
  `GetItem(id)` refusal site, in 56 of 63 rounds. What it costs is now measured
  rather than assumed: it denies a chart its parts list, and only a chart in that
  state can strand anything.

#### Two product defects, for the record

- **A blank slide of ours shipped in the finished deck.** One `slides.add()` can
  land two; the branch reporting un-nameable landings counted the event, not the
  slides. Nine archived rounds affected; five of the last thirteen began on a
  deck already dirty by one slide, invisible because the scenarios measure
  GROWTH rather than absolute size.
- **An update died on a refused id.** The resolve sync was unguarded, so one
  by-id lookup this host would not honour took the whole update down. Guarded,
  then re-asked through a collection read. **Still unexercised on a real host** —
  it has not fired once in six rounds.

#### Mechanics that earned themselves

- Every `playwright-cli` call is bounded; the round's own 30-minute deadline
  could never fire because it is checked between calls.
- An unexpected throw is a `threw` reason rather than a dead process.
- The cycle stops when a leg finished but archived nothing — the state that
  otherwise lets the next leg overwrite the only copy of the evidence.
- `sweepDeck` reads its own answer instead of always claiming success.
- The gate's novelty bucket uses a five-round window and names the build a
  signature FIRST appeared in. It had spent fifteen rounds blaming nine innocent
  builds for one 064-era signature.

#### Owner-only, unchanged

Sign-in, a real mouse drag, and desktop PowerPoint. Re-sideloading an add-in into
a document is drivable and now automated, but a **browser death does not always
lose the sideload** — once it did, once it did not — so that path is written and
tested and has not yet met the state it exists for.

The browser deaths themselves are **connected standby**, not crashes: Windows
Event 507 four times in a morning, `ERR_NETWORK_IO_SUSPENDED` in the console
tail. Power settings on AC now prevent it; on battery the display still sleeps at
four minutes, which on a Modern Standby machine is what triggers it.

### Retire the positional group-member mapping with `Shape.creationId` — CLOSED 2026-08-29, REFUSED BY THE HOST

**The remedy does not exist here.** This host reports PowerPointApi **1.10** in
`requirementSets` and does not populate `Shape.creationId`. All three probe
questions say so, unanimously, over the 25 rounds to 2026-08-28:

    creationid-on-fresh-shape       absent          25 of 25
    creationid-survives-a-sync      no-creation-id  25 of 25
    creationid-survives-grouping    no-creation-id  25 of 25

So the blocker recorded below — "the paths that reach this code are gated at
1.8, so it needs a 1.10 branch" — was never the blocker. A 1.10 branch would
compile, ship, and find nothing to read. The positional mapping stays, and the
test asserting the TARGET of a node-0 write is its permanent guard rather than
a stopgap.

Worth noting how this stayed open: the three answers had been on every sheet
for months, and the questions sat in `PENDING_QUESTIONS` marked "the committed
sheet predates this question" because the fixture beside the gate was a
2026-08-12 capture. The probe was even spending a scarce scratch slide
re-asking each of them. The analysis below is kept because the reasoning about
what the mapping rests on is still true.

The in-place update maps scene nodes to a grouped chart's shapes by POSITION:
member 0 is the anchor, the rest line up in drawing order. Round 145 showed what
a mistake in that mapping costs — node 0's properties were written onto the
group itself, and only the host caught it.

**The order is not documented.** Checked 2026-08-21: the `PowerPoint.ShapeGroup`
reference says `shapes` is "the collection of Shape objects in the group" and
nothing more. The assumption rests on the classic Office object model, where a
shape's index in a Shapes collection is its z-order position and the anchor —
drawn first — is at the back. Consistent with everything observed, and inferred
rather than promised.

**`Shape.creationId` (PowerPointApi 1.10) is a durable per-shape identifier.**
Recording creation ids at draw time would make the node-to-shape mapping
explicit and ordering irrelevant, which is the only real fix for the assumption
above. Blocked on the requirement set: the paths that reach this code are gated
at 1.8, so it needs a 1.10 branch with the positional mapping kept as fallback.

Until then the guard is the test that asserts the TARGET of a node-0 write
(`writes node 0 to the group's ANCHOR, not to the group itself`) — a count check
cannot catch a mapping error that preserves the count.

### Two questions production answered before the probe sheet could — RETIRED 2026-08-21

#### What actually answered them

**`tag-on-group-survives` — "Does a tag written on a GROUP read back?" YES, from
production, over 149 rounds.** It is the mechanism the whole product runs on:
every chart is discovered by reading `POWERCHART_CONFIG` off its group, and
since #645 the in-place update reads `POWERCHART_SCENE` off the same group
thirteen times a round. Rounds 146, 147 and 149 each completed three in-place
updates, which is not possible unless a tag on a group reads back. A probe
cannot beat that evidence; it can only agree with it more slowly.

**`grouped-child-by-id-from-slide` — "Can a shape INSIDE a group still be
resolved by id off the slide?" MOOT.** It was described as the question that
decides whether the in-place update can ever work here. It is not, because the
update no longer takes that route: #643 reads a grouped chart's members from
`shape.group.shapes` (a `ShapeScopedCollection` at PowerPointApi 1.8), and #646
writes through those member proxies. The by-id-off-the-slide path is not used
and does not need an answer.

Note what that means for the older entry above, "The probe has been blind on
GROUPS for the whole archive": the blindness was real, and the product routed
around it without the probe ever seeing. **A question worth asking is not the
same as a question worth waiting for.**

#### The branch that tried the other remedy

`claude/ask-the-decisive-probes-early` (`f09041c`, 2026-08-13) proposed moving
both to positions 5 and 6. It is pushed and preserved, and it was NOT merged,
for three independent reasons: it ships a deliberately failing test ("KNOWN RED,
AND THE REASON IS THE POINT"); 23 commits have touched `host-probe.ts` since,
and it no longer rebases cleanly; and its premise — that these two questions
gate the in-place update — is now false.

Its diagnosis was still worth keeping, and is recorded above: these questions
never get put. The conclusion inverted once production answered them.

#### What is left open

The `no-scratch-*` split in the older entry still stands for the four remaining
starved questions. Retiring two of the six is not a fix for the harness; it
removes two questions that could never pay, and gives their slide back to the
ones that might.

### The slide-size ladder's first rung hangs about twice as often as it answers — RE-MEASURED, AND IT IS WORSE THAN THAT

**RE-MEASURED 2026-08-29 over 270 rounds, and the headline below is wrong in
the direction that matters. It is not "about twice as often as it answers".
Every round that reads a slide size times out on rung 1 first:**

    rounds that hang on rung 1, then record pageSetup        82
    rounds that hang on rung 1, then record exportedSlide    75
    rounds that record a size WITHOUT hanging on rung 1       0
    rounds that hang and never get a size                     0

**157 of 157, always the full 4000ms — one hang per round, never two.** The four
seconds is universal, not occasional. And rung 1 is not the loser it reads as:
it supplies the final answer in 82 of those 157, more often than rung 2 does.

The 2026-08-23 figures below (22 · 22 · 44) were the same questions asked with
a pattern that matched a fraction of the archive. Both counts were low and the
ratio between them survived, which is what made them look reliable.

**So neither obvious action follows.** "Lower the bound" now saves four seconds
in EVERY round rather than a third of them — a much better prize — but rung 1
is also the rung that wins half the time, and a lower bound might buy the time
by giving those 82 answers away to an export whose cost is still unmeasured.

**What decides it is one number nobody has: how long a SUCCESSFUL rung-1 read
takes.** Under 4000ms and the timeout is being paid by a different call than the
one that answers, so the bound comes down for free. Over it, and the bound is
exactly what makes those 82 answers possible.

`slideSize()` now traces `ms` on all four rungs (2026-08-29), so the next rounds
answer it. Four tests, four mutants — three killed by the runner and the fourth
confirmed by hand after the runner failed to apply it.

**ANSWERED THE SAME NIGHT, rounds 297-299 — and it is the cheap answer.**

    297   STALL afterMs=4000   READ pageSetup ms=270
    298   STALL afterMs=4000   READ pageSetup ms=246
    299   STALL afterMs=4000   READ pageSetup ms=263

A successful rung-1 read costs **246-270ms**, against a stall of a flat 4000.
`ms` is counted from `slideSize()`'s own entry, so a read that traced 270 cannot
be the call that waited 4000 — **the stall and the answer are different calls**,
and the second one succeeds cheaply without the first's bound helping it at all.
The 4000ms is not buying those 82 answers. It is being paid by a first call the
host is not yet ready for, and a later call gets the same answer in a quarter of
a second.

So the bound can come down, and by a lot: 1000ms leaves a **4x margin** over
every successful read yet observed and returns ~3 seconds per round. What it
must not do is come down to where the export on rung 2 starts winning races it
used to lose, since nothing has measured what THAT costs — the reason this was
not simply lowered in the first place, and still the reason to measure rung 2
before choosing the number.

**Three samples.** Enough to answer the question that was asked (is a successful
read cheaper than the bound? yes, by 15x) and not enough to pick the new bound
from. Every round from 47db58a on carries the field, so the sample grows on its
own; re-read before choosing.

**AND ROUND 300 CONTRADICTS THE "157 OF 157" ABOVE, which is the most useful
thing here.** It carries no stall at all and a rung-1 read of **138ms** — the
first round in the archive to read a slide size without paying the four seconds
first.

What was different is not the code, which is the same build as 298 and 299. It
is that round 300 ran immediately after a crash recovery: PowerPoint's own
Refresh, a full tab reload, and a pane reopened seconds before. **So the stall
looks like a property of a host that has been sitting, not of the call.** A
freshly loaded document answers `pageSetup` first time, in 138ms.

That is one round and it is offered as one round. But it points the remedy
somewhere different from "lower the bound": if the four seconds is the first
call meeting a cold host, then a shorter bound simply reaches rung 2 sooner on
exactly the runs that are already unwell, and the thing worth measuring is
whether a cheap warm-up call before the ladder removes it entirely.

**ROUND 301 SUPPLIES THE MISSING NUMBER, and it unblocks the decision.** It is
the first round in the archive to show both halves in ONE call:

    301   STALL afterMs=4000   READ source=exportedSlide ms=4280

`ms` counts from `slideSize()`'s entry and the bound is 4000, so **rung 2's
export cost about 280ms** — the same as a successful rung-1 read (246-270ms).
That was the one thing this entry said it was blocked on: "nothing has measured
what the export costs".

It is measured now, and it is cheap. **So the bound can come down and the
fallback is nearly free.** At a 500ms bound a cold host would spend 500 on rung
1 and ~280 on rung 2 — under 800ms against today's 4,280, a saving of about
three and a half seconds, with the same answer at the end of it. The worry that
made this wait ("a lower bound might trade those 82 answers for an export whose
cost nobody knows") is retired: the export costs what the read costs.

301 also supports the cold-host reading. Its tab had been reloaded 35 minutes
earlier and then sat idle, and the stall came back — where round 300, running
seconds after a reload, had none.

**Round 302 then narrowed it, which is more useful than a third confirmation.**
Its PANE was closed and reopened three minutes before the round, and it stalled
anyway (`READ pageSetup ms=504`, after the usual 4000). So "recently touched"
is not the variable. What round 300 had that 301 and 302 did not is a full TAB
RELOAD after PowerPoint's own crash-Refresh — a new editing session, not a new
pane. Reopening the pane does not buy it.

Three samples, one stall-free, and the thing that distinguishes it is the
heaviest possible reset. That is a narrower and more testable claim than the one
it replaces, and it is still three samples.

**What remains before changing the number:** decide between the two remedies
rather than doing both blindly. Lowering the bound saves ~3.5s on every stalled
run; a warm-up call might remove the stall altogether and save the rung-2 hop as
well. They are not exclusive, and neither is now blocked on a measurement.

One defect found writing those tests, worth recording because it is this
project's recurring one: the rung-3 assertion used `traceLog().entries.find`,
and that log accumulates across the file, so it read a line an EARLIER test had
written and passed against a deliberately broken rung. Measuring something
adjacent to what you meant to measure — the same shape as the gauge total
measured unbold and drawn bold. It searches from a captured index now.

**Original entry, 2026-08-23, kept because the reasoning about WHY not to act
blindly is still right:**

`slideSize()` rung 1 reads `presentation.pageSetup` directly, bounded by
`SELECTION_TIMEOUT_MS` (4000ms). Across the archive:

    slide size read, by source:  pageSetup 22 · exportedSlide 22
    gave up waiting on "reading the slide size":  44 rounds

So rung 1 **is not a dead rung** — it answers 22 times — and it hangs roughly
twice as often, first seen in round 117 and in 26 of the last 30 rounds. Every
hang costs a flat four seconds before rung 2 exports a slide and answers.

The trace already names the shape precisely, and it is worth quoting because it
rules out "the host was busy":

    "THIS CALL ALONE is stuck — the host answered something else 3845ms AFTER
     this call"

**Why this is recorded rather than fixed.** The obvious change is to lower the
bound, and the obvious change is not obviously right:

- rung 2 exports a slide, and nothing has measured what THAT costs. Trading a
  4-second wait for an unmeasured export is not known to be a saving.
- 4000ms is `SELECTION_TIMEOUT_MS`, shared with the selection path. Lowering it
  here means either a separate constant or a change to a budget something else
  depends on — see the 0.5→0.8 `UPDATE_SHARE_LIMIT` episode for what happens
  when a shared budget is raised on a good measurement of the wrong thing.
- the cost is ~4s in a ~490s round, under one percent. It is worth knowing and
  it is not worth a blind change.

**What would settle it:** time rung 2 on the rounds where rung 1 hangs. Both
already trace; neither carries a duration. One field on each, and the next round
answers whether a lower bound saves anything at all.

Found by the archive sweep of 2026-08-23. The sweep reported it as "every round
eats a 4-second timeout — 38 of 38, never zero", which is the recent era and not
the archive: it is 44 of 147 overall, and universal only since round 117.

### The four dependabot advisories — ALL FOUR NOW FIXED, 2026-08-27

All four fixed 2026-08-27. See `package.json`'s `_overrides_why` and git.

### The original analysis — 2026-08-26

Every push prints "GitHub found 3 vulnerabilities on the default branch (2 high,
1 moderate)". Checked once so it does not have to be re-checked every push.

    image-size  *        HIGH x2   ICNS and JXL/HEIF parsers loop forever on a
                                   malformed image (GHSA-w3rx-r6r6-pgpr,
                                   GHSA-5p2g-fcmc-qvqq)
    qs  6.11.1-6.15.1    MODERATE  qs.stringify crashes on null entries in
                                   comma-format arrays (GHSA-q8mj-m7cp-5q26)

**`qs` is dev-only, and is now FIXED** — `@stryker-mutator/core` ->
`typed-rest-client` -> `qs`, never shipped, never runs outside mutation testing.

Two things I wrote here first were wrong, and both were measured rather than
assumed the second time:

`npm audit fix` does NOT clear it. npm reports `fixAvailable: true` and then
changes nothing — zero lockfile delta, four vulnerabilities before and after —
because it cannot move a transitive that `typed-rest-client` pins. A fix that is
"available" and inert is worse than one that is absent, because the report reads
as actionable.

What does clear it is an `overrides` pin to `qs@^6.15.3`. The advisory range is
`6.11.1 - 6.15.1` and 6.15.3 is outside it, so this is a PATCH bump on the same
minor rather than a version gamble. `npm audit` goes from 4 vulnerabilities (2
moderate, 2 high) to 2 high. The `_overrides_why` key in package.json says when
to remove it.

**`image-size` is NOT REACHABLE.** It arrives under `pptxgenjs`, which IS a
runtime dependency — `src/render/pptx-deck.ts:74` imports it dynamically — so the
tree alone makes it look shipped and live. It is not: pptxgenjs calls
`image-size` to size images handed to `addImage`, and **SSF Charts never calls
`addImage`**. There is no image path in `pptx-deck.ts`, no PNG or JPEG or data
URI anywhere in it. The deck is shapes and text. Both parsers named in the
advisories need an image to parse and are handed none.

**There is also nothing to upgrade TO.** Both advisories give their vulnerable
range as `<=2.0.2`, and 2.0.2 is the latest published `image-size`. Every version
that exists is flagged, so an `overrides` pin — which is what fixed `qs` above —
has no target here.

And npm's proposed remedy is not a bump but a THREE-MAJOR DOWNGRADE: `fixAvailable`
names `pptxgenjs@1.1.5` against the 4.0.1 in use. `npm audit fix --force` would
gut the deck writer to patch a code path this product never executes. Not taken,
and the direction is worth stating — I first recorded this as "a major bump",
which sounds like progress. It is the opposite.

**What would change this.** Any feature that puts a bitmap into a generated deck:
a logo, a screenshot, a rasterised chart fallback, an image placeholder. If
`addImage` ever appears in this repo, `image-size` becomes live and this entry is
wrong — upgrade pptxgenjs before shipping that feature, not after.

### The arm that would settle the first-chart penalty has been run ONCE — 2026-08-26

The largest single cost in a round is the first chart of the deck-wide rescale:
~44.1s against ~17.6s for the same chart size, same changed count, and the same
slide occupancy (n=26 each — see `WHAT THE SLIDE ACTUALLY HELD` in triage). Every
one of the three WRITE syncs carries the penalty at 2.1-2.35x while the tag sync
carries only 1.16x, and `our-idle-is-negligible` puts our own waiting at 1ms
median. So the time is going to the host, during writes, on the first chart.

The leading explanation is contention with the DECK SCAN that runs immediately
before the rescale. `scanSettleMs` exists precisely to test it — a pause inserted
between the scan and the first update, "the only gap where a pause can tell the
two apart", traced at zero as well so every archived round records its arm.

**It has one nonzero sample in the whole archive.**

    settleMs = 0        58 round(s)
    settleMs = 10000     2 round(s)   -- one with a usable first-chart timing
    absent              60 round(s)   -- predates the trace

n=1 against a first-chart distribution spanning 13.3s to 58.3s says nothing at
all, and the single sample (38552ms) sits in the middle of it.

**Why it is never run: it cannot be set from the driver.** `scanSettleMs()` reads
the pane's `localStorage`, so arming it means typing

    localStorage.setItem("powerchart-scan-settle-ms", "3000")

into the pane console by hand, before a round, every time. `scripts/round.mjs`
has no way to do it — it drives entirely through playwright-cli refs and never
evaluates JavaScript. So the arm is available in principle and unreachable in
practice, which is exactly the shape of a control that quietly never runs.

**What would fix it.** `playwright-cli` DOES have `eval <func> [target]`. The
pane is a cross-origin iframe, so a page-level eval cannot touch its
`localStorage` — but Playwright evaluates in the frame that OWNS the target
element, so passing a ref from inside the pane should reach it:

    eval "() => localStorage.setItem('powerchart-scan-settle-ms','3000')" <pane-ref>

That is untested. It needs a `--scan-settle <ms>` flag on the driver, defaulting
to off so no existing round changes what it measures, and then enough rounds in
each arm to clear a noise floor the archive puts at IQR 14%.

**THE DECK CANNOT SEPARATE POSITION FROM OCCUPANCY, and this is the concrete
blocker.** The rescale's charts 1, 2 and 3 all sit on the same busy slide, but 1
is a 24-node chart and 2 and 3 are 16-node. Charts 4-8 are 24-node and sit on
1-shape slides. So the archive offers `24/18 on a 3-shape slide, FIRST` and
`24/18 on a 1-shape slide, LATER` — and no `24/18 on a 3-shape slide, LATER` cell
at all. Position and occupancy move together by construction, in every round.

One round makes the point without any pooling. From the crashed run of
2026-08-26 (`crashes/2026-08-26T19-56-29-crashed-run.json`), one run, one deck:

    chart 1/8   24 nodes, 18 changed   38174ms   syncMs [12635,11803,13078,...]
    chart 4/8   24 nodes, 18 changed   17999ms   syncMs [ 5975, 5692, 5736,...]

Same size, same changed count, same run — 2.1x, and EVERY write sync carries it
rather than one slow step. Charts 2 and 3, on the SAME slide as chart 1, cost
13.6s and 13.3s, so the slide is not what makes chart 1 expensive. What the deck
cannot tell us is whether a 24-node chart LATER in the run on that same busy
slide would also be cheap. Building one such slide would answer it in one round.

**THE SCAN'S OWN DURATION DOES NOT PREDICT THE PENALTY — checked 2026-08-26.**
`scanned the deck for charts` carries an `ms` that no tool reads. Split 69 rounds
at the median scan time and compare each round's first chart against ITS OWN
later charts, so a generally slow host cannot manufacture the result:

    short scans   n=34   median scan 1139ms   median first/later ratio 2.22
    long  scans   n=35   median scan 1383ms   median first/later ratio 2.15

No effect, and what there is runs the wrong way.

**This does NOT clear contention**, and the limit is worth stating rather than
overreading a null. The natural range is narrow — 1.1s against 1.4s, 21% apart —
and a weak contrast cannot detect a real effect. A scan may also leave the host
busy for a period unrelated to how long the scan itself took. What it does rule
out is a simple dose-response, where a longer scan costs a slower first chart.

That is an argument FOR running the settle arm rather than against it: a 3-10s
pause is a far stronger manipulation than the 250ms of natural variation above,
and it is the only one that separates the two.

**What NOT to do.** Do not make the pane counterbalance this on its own. Half of
every future round would then carry a pause, which changes what a round measures
for every claim already resting on the archive. The flag is opt-in for that
reason.

### Text drawn over text — 2,148 to 317, and what is left

**The measurement.** Every kind under 24 option variants and 10 data shapes, at
eight frame sizes, two fonts and both orientations — about 24,000 charts — with
overlap measured by the frame gate's own ink rule. It ran for the first time on
2026-08-27 at **2,148 overlapping pairs** and stands at **317**.

**AND THE GATE HAD THE SAME BLIND SPOT IT WAS BUILT TO CLOSE — 2026-08-29.**

`frame-fit.test.ts` swept kinds, frames and fonts and missed the option and
data-shape variants; that is why this family lived nine days behind a green
gate. `overlap-budget.test.ts` was written to close that hole, and it swept
every option and every data shape **and never one of each** — the two tables
were concatenated, not crossed. A chart with a secondary axis AND ten series
was not among its 24,000. Neither was a footnote on twenty-four categories.

Crossing a slice of them (six layout-changing options against four
label-stressing data shapes) took the count from **537 to 4,010** and produced
**twenty-two shapes this engine has never been seen to draw**. Nothing about the
engine changed that day. The number moved because the sweep did, and a figure
from before it is not comparable with one after — the same trap as the
"75 pairs" of 2026-08-19, sprung a second time in ten days.

    2,683   `p#-*`, small-multiples panels — ALL of it new
    1,058   `value-axis-title` — the open decision below, unchanged in kind
      269   everything else, including three new shapes

**THE FIND — AND THE FIRST ANSWER WAS WRONG, so read the correction first.**
This entry said, from the count alone, that "small multiples do not thin their
labels when the data gets dense". Tested on 2026-08-29 and refuted. The category
axis thins perfectly well — `catFs` shrinks until every name fits its slot and
`clipToWidth` clips what no floor can fit — and **zero** of the 910 overlaps are
between two names in the SAME panel:

    panel category-name overlaps:  WITHIN a panel 0 · ACROSS panels 910
    by frame:                      300x60  910   (every other frame: none)

All of it is one frame, and these labels are not merely touching. On a 300x60
chart with ten series in two columns, `p0-category-0` is drawn at **y = 307 on a
chart 60 points tall**, on top of the next row's names at 316.

**A DEFENSIVE CLAMP TURNS AN IMPOSSIBLE GRID INTO A PLAUSIBLE WRONG ANSWER**,
and that is the whole finding. It took two wrong explanations to get here, both
written from arithmetic rather than measurement; the third is measured end to
end.

Ten series in two columns is five rows, and

    panelH = (cfg.height − titleH − footH − gap × (rows − 1)) / rows
           = (60 − 22 − 0 − 40) / 5
           = −0.4

`buildMultiples` builds the grid anyway and hands each panel `height: −0.4`.
Then `normalizeConfig` runs, and `clampDim` says:

    if (typeof v !== "number" || !Number.isFinite(v) || v <= 0) return fallback;

**−0.4 is `<= 0`, so the panel is given `DEFAULT_SIZE.height`, 300.** Every
panel is therefore laid out as a full 300-point chart, and the composition
places those at a 9.6-point pitch — ten 300-point charts stacked 9.6 points
apart inside a box 60 points tall. 129 of the scene's 139 text nodes end up
below the bottom of the chart.

Measured rather than reasoned, and this is the number that settles it:

    a standalone 155x300 chart puts category-0 at   y = 285.0
    the panel puts category-0 at, in its own frame  y = 285.0

Identical. The panel is not a squeezed panel; it is a full-size chart wearing a
panel's offset.

`clampDim` is not at fault and should not be changed. It exists because
`width: NaN` arrives from pasted configs, saved templates and shape tags written
in other decks, and repairing that beats a stack trace in a headless render. The
fault is that an INTERNAL arithmetic error reached it. A repair meant for
malformed input from outside absorbed a computation that should never have been
attempted, and converted "this grid cannot exist" into "here are ten charts on
top of each other" — a loud failure made quiet.

So the fix belongs in `buildMultiples`: check that a panel has a usable height
BEFORE building the grid, and if it has not, do what this engine does everywhere
else with a reservation it cannot pay for — decline it, and render the chart
whole. This is the house pattern named in the memory *a fit and a clamp fight,
and the clamp wins*, in its most expensive form yet: here the clamp does not
merely undo the fit, it manufactures a chart nobody asked for.

The counts below stand as measured; only the explanation changed. Crossed with
twenty-four categories or ten series:

    910  p#-category# / p#-category#
    408  p#-title / p#-title
    378  p#-value-axis / p#-value-axis
    292  p#-label# / p#-label#
    261  p#-label## / p#-label##
    198  p#-total# / p#-total#

**Every family traced, 2026-08-29 — it is one defect, not six.** The
inference above ("the same grid seen through different labels") was checked
rather than left standing. Over all twenty-three `p#-` shapes:

    shape                          within  across  off-chart  frames
    p#-category# / p#-category#         0     910        910  300x60
    p#-value-axis / p#-value-axis       0     420        420  300x60, 80x60
    p#-title / p#-title                 0     408        312  300x60, 80x60
    p#-label# / p#-label#               0     252        252  300x60, 80x60
    p#-total# / p#-total#               0     140        140  80x60, 300x60
    …and eighteen more, all the same shape

Three things hold across every one of them. **`within` is zero** — no two
labels in the same panel ever collide. **`off-chart` equals the overlap count**,
family by family: essentially every pair involves a label drawn outside the
chart's own bounds. And **only two frames appear at all**, 300x60 and 80x60,
the two shortest in the sweep. That is the signature of the grid, seen twenty-
three ways.

(Those counts are from a `multiples 2 columns` × four-data-shape probe, so they
do not match the ratchet's totals line for line — it crosses six options. The
structure is what is being claimed, not the arithmetic.)

**~~One shape is NOT the grid, and it is the interesting leftover.~~ WRONG, and
the ratchet said so within the hour.** This claimed `p#-total# / p#-cagr-label`
was "a real label collision at a size somebody would present" — 9 of its 26
within-panel, and the only family appearing at 480x300 and 960x540 — and that
it "will still be there after the grid is fixed".

It is not there. After the fix the budget table holds **no `p#-` entry at all**,
and the ratchet's "draws no shape it has never drawn before" test passes, which
it could not if any `p#-` pair still occurred. The grid took the whole family,
that one included.

**Why the probe and the ratchet disagreed, because that is the reusable part.**
They were not measuring the same charts. The probe built its `10 series` shape
from scratch — ten fresh series with generated values — where `DATA_SHAPES`
spreads the sample's existing series and keeps their colours and roles. Same
name, different data, different decorations, different overlaps. A diagnostic
written beside a gate has to reuse the gate's own variant table or its numbers
answer a different question, and the two agreeing everywhere else is exactly
what made this one look trustworthy.

**The lesson, and it is the same one twice in a night.** The first explanation
here was written from a count and read plausibly — this engine really does thin
labels everywhere else, so "not inside a panel" fitted. It took one diagnostic,
which cost minutes, to find that the true answer was in a different file
entirely and an order of magnitude worse than the story. A count tells you where
to look and never why.

**FIXED the same night, once it was diagnosed rather than guessed.** The whole
2,683 goes to zero — every one of the twenty-three `p#-` shapes — and the
sweep's total falls from 4,010 to **1,327**. The change is one line in
`buildMultiples`:

    if (!(panelW > 0) || !(panelH > 0)) return null;

`> 0` and no more, deliberately. Every positive size passes `clampDim`
untouched, so this cannot alter a chart that renders today; it refuses only the
case that was being silently rewritten. Declining the grid renders the chart
whole, which is what this engine does with every other reservation it cannot
pay for — and the test asserts BOTH halves, because declining must not cost the
user their chart.

Two mutants, both dead, and the second is the point: removing the guard fails
the test, and so does weakening `> 0` to `>= 0`. That second one survived the
first version of the test, because the case it used has `panelH = −0.4` and
fails both operators. `clampDim`'s own test is `v <= 0`, so **exactly zero** is
rewritten too — there is now a case pinning it (four series, two columns, height
10, no title → `panelH = (10 − 10) / 2`).

This was written up as "not fixed tonight, it wants the sweep, the ratchet and a
round". It got the sweep and the ratchet; the round is what it still owes, and
it is a RENDERING change, so it must not be called done until one has run.

**Still not crossed**, so nobody has to re-derive it: every option outside
`CROSS_OPTIONS` against every data shape, and any product of three or more.
Widen that list before calling a family closed.

**It is a gate now, not a script somebody remembers.**
`test/overlap-budget.test.ts` runs the sweep in eight seconds on every build and
holds a budget PER SHAPE. Three of its four tests catch a regression; the fourth
catches a budget left ABOVE the real figure, so improving the engine is supposed
to fail the file and the numbers get edited down. A shape absent from its table
is a regression by definition — it is text this engine has never been seen to
draw over text. **The per-shape table there is the live list; this section does
not duplicate it.**

The fixes themselves are in git, with a commit each explaining what the defect
was and what the fix cost. What is worth keeping here is what remains, and what
the exercise taught.

---

**CLOSED 2026-08-29: `valueAxisTitle`. Two changes, and the second one was not
the decision this entry spent four days framing.**

The entry below is kept in full because the reasoning is the point. It asks
"where does a unit belong when the band above the plot cannot hold the title, the
unit and the topmost tick", writes up three ways out, and measures four remedies
into the ground. The question was wrong. **The band's other occupants had settled
it long before**: on a chart that cannot hold everything, chrome yields and the
title stays — `docs/MANUAL.md`, "Chrome yields to the title", whose worked
example is a 300x60 banner. The category names go. The axis strip goes. The
legend goes. The unit was the one thing in that band nobody had ever told, so it
went on being painted over the title.

    clip to 40% of the chart      1,327  ->  1,014
    yield to the title            1,014  ->    785      title/VAT 205 -> 0

**Why dropping author text is right in this one case**, when this entry spends
three paragraphs below explaining why it is normally not: the old behaviour did
not KEEP the unit. It printed it through the title, so the reader lost both. A
chart that shows neither its unit nor its heading has not preserved anything.

**And the number that framed the whole thing was measuring the wrong input.**
`OPTIONS.valueAxisTitle` is a 27-character sentence for an option documented as
`€m`. At `€m` the family is 118, not 687. The "largest defect family in the
engine" corresponds to ONE overlapping pair in the 123-chart deck we ship. That
warning now sits at the top of `test/overlap-budget.test.ts`, because this entry
is the third decision the figure has steered and the second it steered wrongly.

Everything from here down is the record of getting there.

The count is bigger than the "287 of 317" this said before and the problem is
the same size — the sweep widened underneath it, and the panels that briefly
outnumbered it have been fixed. It is now the whole of what remains bar a
119-pair tail, which makes it the only overlap decision left worth the name.

**It said 1,058 and a 269-pair tail for half a day, and both were arithmetic
rather than measurement.** Those were the 4,010 sweep's split, carried forward
by subtracting the 2,683 that the small-multiples fix removed. Re-summing the
table gives 1,208 and 119 — the family GREW by about 150 while the tail shrank
by the same, because a chart that used to render as an impossible grid now
renders as an ordinary one, and an ordinary chart with a unit label contributes
`value-axis-title` overlaps of its own. **A subtraction is not a
re-measurement**, and this file has now made that mistake with two different
figures in two days.

Four collision shapes, and they are not one problem. Two were this label's own
`y` clamped into the title and onto the topmost tick — fixed. The other two are
the WIDTH, and they are the question:

    value-axis-title / legend#    ~100
    value-axis-title / total#      ~43

Its width is `Math.max(frame.x - 4, textWidth(…))` — a floor that RAISES a width
— so a long unit grows right across the totals row and the legend. Fitting it to
the axis gutter was tried in 2026-08-19 and reverted: `frame.x` is the value
axis's own column and a chart drawn WITHOUT a value axis has none, so the fit
dropped the unit from ordinary charts.

**Three remedies were implemented and measured on 2026-08-28. All three were
reverted, and the numbers are why:**

    drop on box overlap     100 of 352 units survive — seven charts in ten lose
                            the unit, most to collisions only the empty part of
                            an over-wide box was having
    drop on ink overlap     249 of 352 — both clamp shapes go to zero, but a
                            clustered chart at 480x300 in 18pt loses its unit,
                            and that is a size people present at
    move below the title    keeps more, and lands it on the tick numbers — the
                            collision it was moved away from

**Why it cannot be settled the way everything else was.** Every other remedy in
this engine drops or shrinks a label the layout generated. This one is TEXT THE
AUTHOR TYPED. Dropping a tick number costs a reading the gridline still carries;
dropping `valueAxisTitle` deletes something a person wrote and cannot see is
gone. The band above the plot holds the chart title, the unit and the topmost
tick number, and there is not always room for three.

**A FOURTH WAY WAS TRIED ON 2026-08-29 AND DOES NOT WORK, which is worth
knowing because it is the one this engine would reach for by reflex.**

The unit is the only label here drawn at a FIXED size and fitted to nothing —
every other label shrinks to its room and drops past `MIN_LABEL_FS`. Shrinking
is not dropping, so it does not raise the author's-text objection below, and
none of the three remedies above tried it. Measured over every kind, frame, font
and orientation, shrinking the unit until it fits the chart's own width:

    charts where the unit overlaps something    174  ->  153

    what it collides with       before   after
      legend#                      100     100
      total#                        68      39
      title                         41      53
      value-axis                    33      18

It TRADES. The totals row and the tick numbers get most of their space back and
the title loses some of its own, because a smaller unit sits lower in its box
and moves up into the title's ink. Twenty-one charts of 352 is not worth a
change to how every unit is drawn, and `legend#` — the largest partner — does
not move at all, because the legend is nowhere near the unit's width.

So the fit is not the missing piece, and the entry's judgement stands: there is
no bound available here that is not a coupling to something drawn later. It is a
decision about WHERE A UNIT BELONGS, and the three ways out below are still the
three ways out.

**A FIFTH WAY SHIPPED ON 2026-08-29, and it is the one that worked: clip the
unit, because the option is a unit and not a subtitle.**

The four remedies above all argue about WHERE to put a long unit. None of them
asked whether a long unit is a thing this option supports. It is not.
`ChartConfig.valueAxisTitle` is documented as "units label shown at the top of
the value axis (e.g. `€m`)", the pane's input is 56px wide with the placeholder
`e.g. €m`, and the two uses in the 123-chart shipped deck are `€m` and
`$m (log)`. The sweep tests it with **"Revenue in millions of euro, restated"** —
twenty-seven characters. Most of the largest overlap family in this engine was
the cost of a string the feature does not offer.

So: `clipToWidth` at 40% of the chart's width. Measured on the ratchet, not on a
scratch harness:

    total across the sweep      1,327  ->  1,014      and no new shapes
    value-axis-title family     1,208  ->    916
    the tail                      119  ->     98      (unchanged in kind)

It keeps author text where the three 2026-08-28 remedies deleted it — `€m` is
untouched, a sentence is truncated with the start kept — and truncating what
does not fit is what this engine already does to gantt and category names.

**AND THE OTHER HALF OF THE SAME RECOMMENDATION WAS REFUSED BY THE
MEASUREMENT.** Flooring the unit's `y` at the title's ink — `title /
value-axis-title` is 205 and length-INDEPENDENT, which is exactly the clamp
signature this file has recorded five times — was written and swept in the same
hour:

    clip alone                1,014, no new shapes
    clip + floor at the ink   1,156, and `value-axis-title / category#` at 310,
                              a family that did not exist before

The floor buys the title collisions by moving the unit into the category names
on short charts. **A clamp moves a label whether or not the destination is
free** — which the CAGR caption's own note already said, about the same move.
Pinned by `test/value-axis-title.test.ts`, which says re-measure before
"fixing" it.

**~~What is still open~~ — ANSWERED 2026-08-29, and not by any of these.** The
question was where a unit belongs when the band cannot hold the title, the unit
and the topmost tick all three. Three ways out were written up:

- Put the unit at the END of the axis rather than above it — Excel's rotated
  axis title, without the rotation.
- Fold it into the tick numbers themselves: `€m` on the top tick only.
- Accept it as the one label allowed to displace the plot, which is the gutter
  idea the 2026-08-19 attempt reverted.

All three answer "where does it GO". The answer was that on the charts where the
question bites, it does not go anywhere — it yields, exactly as the category
names, the axis strip and the legend already do on the same charts. Every one of
the three would have been real work, and the shipped change is one conditional.

**The tell was in the numbers and went unread for four days.** `title /
value-axis-title` was LENGTH-INDEPENDENT — identical at two characters and at
twenty-seven — which says plainly that the label's size is not the variable. Four
remedies in a row adjusted its size or its position anyway. A family that does
not respond to the thing you keep changing is telling you which question you are
answering.

---

**DIAGNOSED AND DELIBERATELY NOT PATCHED: the dual-axis gutter.**

**MEASURED PROPERLY 2026-08-29, and the answer is: leave it alone.** The "30
pairs" below is not what this costs, and the trade for fixing it is far worse
than when the decision was taken. Both numbers, over every kind, frame, font and
orientation, with `secondaryAxis` and with `pareto`, across four data shapes:

                              scoping kept    merge applied to secondary too
    gutter overlaps                     14                                 6
    secondary ticks DRAWN            4,864                             3,470

**Eight overlaps fixed for 1,394 tick numbers deleted.** That is 174 readings
lost per pair gained. The entry below records the same trade at 34 fixed for 335
deleted — about ten — so the ratio has got seventeen times worse as the rest of
the engine improved, and the original decision to scope the merge off is not
merely still right, it is far more clearly right than when it was made.

The remaining family is 14 pairs of **1,014** (it was 1,327 before the unit-label
clip, which touches nothing here), or 1.4 percent, and **not one of
them is a tick number** — they are `title / series-label#` (8),
`combo-series-label#` on itself (5) and `series-label#` on itself (1). Zero
overlapping pairs anywhere in the sweep involve a `secondary-axis` label, out of
4,864 of them drawn across 2,228 charts.

So this is no longer a design waiting to be built. It is eight overlapping pairs
whose obvious remedy costs 1,394 axis readings, and the design below stays
written for the day somebody wants those eight badly enough to pay for a real
allocator. Nobody should, on these numbers.

A combo's column names, its line names and its secondary-axis tick numbers all
occupy the same two points of x. Laying the first two out in one pass fixed the
ordinary combo (see `seriesLabelNodes`' `extra` parameter), and doing the same
where the secondary axis is present is worse, not better: measured, it fixed 34
overlaps and DELETED 335 tick numbers, leaving a comfortable 480x300 pareto with
two of them. So that case keeps its 30 pairs.

Three families in one strip wants a design — a gutter that is allocated once,
with each family's claim on it stated — not a third pass over the same nodes.

**THE DESIGN, written 2026-08-29. Not implemented: it changes how labels are
placed on every combo chart, so it wants the sweep, the ratchet and a round,
not a late-night patch.**

The failed attempt shared the band by YIELD — merge everything, then let the
secondary strip drop whatever lands on a name. That is why it deleted 335 tick
numbers: `seriesLabelNodes` spreads names over the WHOLE band from the title's
bottom to the frame's bottom whenever they collide, so after the merge there is
almost nowhere a tick can land that is not on a name.

The asymmetry the fix should turn on is not importance, it is **whether a label
may move at all**:

- A secondary tick's `y` is DETERMINED — it is that value's position on the
  secondary scale. Move it and it is a lie. There are typically four to six.
- A series name's `y` is a PREFERENCE — the last segment's midpoint — and
  `seriesLabelNodes` already overrides it freely: it pushes names apart by
  `lineH`, clamps the overflow, spreads them evenly over the band when the gap
  cannot be honoured, shrinks to fit, and drops them all past `MIN_LABEL_FS *
  1.25`.

So the anchored family should be placed FIRST and become geometry, and the
movable family should be fitted around it. Today it is exactly backwards: the
ticks are emitted first and then yield to whatever was already drawn, while the
names spread as though the strip were theirs alone.

**The change, concretely.** Give `seriesLabelNodes` a fourth input — the y
intervals the secondary strip has claimed — and have the spread place names in
the gaps between them rather than across the band. Everything else it does
already applies: too little room and it shrinks, too little for that and it
drops, which is the answer this engine gives everywhere else a reservation
cannot be paid for. No new pass, no new yield rule, and the secondary strip
stops needing one at all.

**What it costs, stated in advance so the measurement is honest.** Names lose
band, so charts with many series in a dual-axis combo will shrink or drop names
that are drawn today. That is the trade this file has already made twice, and
it is the right way round only if the count moves: the 30 pairs should go to
roughly zero without the 335 tick numbers being spent. Anything else and the
answer is that a dual-axis combo with many series has no room for both, which
is a legitimate finding and should be recorded rather than forced.

**Order of work:** intervals in first with the parameter unused and the sweep
re-run to prove it is inert; then the spread; then re-measure. Two numbers
decide it — pairs, and tick numbers drawn.

---

**FIVE INSTANCES OF ONE DEFECT, worth naming because it will recur.** A strip is
fitted to the room it has, and a CLAMP a few lines below — there to keep a label
on the canvas — moves it back toward its neighbour and takes the room away
again. Found in the secondary value axis, the scatter's x tick numbers, the
gantt's date strip, the value-axis title, and the gauge's slice labels. The
symptom is labels that are SIZED correctly and still touch, beside a
`Math.max(0, …)` or `Math.min(width - w, …)` on the label's own x or y.

Three remedies, in order of preference: one shift for the WHOLE strip so every
gap is preserved; pay in SIZE, shrinking until the clamped layout clears; or
test the fit on the CLAMPED position rather than the raw one.

**And its sibling: a fit that measures something ADJACENT to what is drawn.**
The gauge's total measured unbold and drawn bold. The gantt's nudge measured at
`fs * 0.9` and drawn at `headFs`. A test measuring node boxes where the gate
measures ink. A label's `y` read as its ink when it is the box top and the ink
sits lower inside it — that one was mine, and the post-condition I wrote against
it never fired once, which is how it was caught.

**THE BLIND SPOT IS THE LESSON ABOVE ALL OF THEM.**
`test/frame-fit.test.ts` sweeps every kind at eight frames and seven fonts and
allows no overlap at all. It was green for the nine days that this entire family
sat in the option and data-shape variants it does not sweep. A gate is only as
wide as its sweep, and a green one is a statement about its sweep rather than
about the product.

A related caution, learned the same way: the figure that stood in this file from
2026-08-19 — "75 overlapping pairs" — was measured with a cruder ink rule and
was never comparable to anything. Re-measuring gave 2,148. A number in prose
decays; the sweep is four lines of script and the archive is on disk, so re-run
it rather than quote it.

### A reconcile window sized from a slide count the host reports late — 2026-08-29

**Round 297 stopped a whole cycle on a verdict that was arithmetically correct
about the wrong population.** `two slides claiming one slot` FAILED with
`4 slides inserted, 4 kept, 4 of 2 still re-editable; 0 queued as duplicates`,
and the fatal check did its job: a scenario that was passing had stopped.

It is not the flake this scenario already has. Five earlier failures in 273
rounds: four of them — 060, 148, 253, 273 — read `2 queued as duplicates`, so
the dedup RAN and left slides behind, and the fifth (287) is the stale-chunk
`Failed to fetch dynamically imported module` that `src/render/lazy.ts` was
written for. None of the five found zero duplicates. This one did, which is a
different thing.

**The trace names the cause in three lines:**

    handed the host a generated deck   expectedSlides 2, landed 0
    handed the host a generated deck   expectedSlides 2, landed 2
    read the deck back                 range [3,5], read 2, unread 0

The first insert reported **nothing landed**. Its slides did land — the deck
went 3 to 7, all four of them — but not by the time anyone counted. So
`afterInsert = await slideCount()` answered 5, `reconcileDeck` was handed the
window `[3,5]`, and it examined two slides out of the four that existed. Two
distinct titles in the window, no duplicates, verdict `0`. Every step correct;
the window was wrong.

`unread` is **0** on that read, which matters: the host was not refusing. It was
BEHIND. Those are different failures and only one of them the guards already
cover — `duplicateSlot` checks `blind` on its readback and there was nothing
blind to catch.

**This is `addSlides`' problem seen from the other side.** That function exists
because "PowerPoint on the web silently drops some `slides.add()` calls", and it
verifies growth from a fresh context rather than trusting the batch. Here the
host does the mirror image: it reports slides as NOT landed that did land. And
`slideCount()` is already the fresh-context read, so a fresh context is not the
remedy — the host is simply late, the same lateness `slideShapeCounts` pays
`COUNT_SETTLE_MS` for on the shape counts.

**The fix, and why it is not in tonight.** `duplicateSlot` knows it asked for
four slides. When `afterInsert - before` is not four, the reconcile is about to
be given a window that cannot contain the duplicates, and the honest answer is
the one `blindSkip` already gives for the other kind of unreadable host: say the
run could not be judged, rather than judge it. That is a few lines.

It is not in tonight because it changes what a round MEASURES, and this
scenario's pass/fail series is 273 rounds long and is itself evidence. A change
to a verdict rule belongs in daylight, with the series re-read afterwards, not
at 2am between cycles. Recorded here so it is not re-derived.

**Read alongside** the entry above about the ratchet's own blind spot: three
times this week a gate has been correct about the thing it was measuring and
wrong about the thing it was believed to measure.

### The crash, counted across the whole archive — 2026-08-28

Twenty-four crash records over fourteen distinct builds, taking the LAST step
each one wrote and counting each phase once per build so one bad night cannot
vote three times:

    4 build(s)   draw  — parts list outcome
    4 build(s)   pane  — collecting deck evidence — scanning
    2 build(s)   probe — asking
    2 build(s)   draw  — batch issued
    2 build(s)   probe — re-asked what the empty deck could not answer
    1 build each update / group / probe answered / probe second pass

**Counting RECORDS rather than builds would have said `parts list outcome` five
times to the scan's four** — and three of those five are the same build on the
same day. One build that crashed three times is one piece of evidence, not
three, and the difference decides which phase looks dominant.

**2026-08-29 adds one build, and it breaks the tie toward the scan.** Build
`4275306` crashed five times in thirty-six minutes on the 4:3 leg, and four of
the five ended at the SAME step:

    01-49   513.4s   pane  — collecting deck evidence — scanning
    02-00   533.5s   pane  — collecting deck evidence — scanning
    02-10   523.6s   pane  — collecting deck evidence — scanning
    02-21   538.9s   pane  — collecting deck evidence — scanning
    02-25   160.8s   probe — did not re-acquire, the slide resolved

By the rule above that is **one** build, taking the scan to 5 against draw's 4 —
a lead, not a verdict, on fifteen builds. What is new is not the count but the
**consistency**: four crashes on one build, one phase, and all four between 513
and 539 seconds. Round 290's 524s crash sits inside that band too. So this
failure has a time signature — roughly nine minutes into a round, during deck
evidence collection — where the archive as a whole looked scattered.

Worth one more build's worth of records before anyone acts on it. Four of five
is a pattern; it is not yet the answer to "why nine minutes".

### FOUR MORE BUILDS ARRIVED, AND THE SCAN WINS — 2026-08-29

The entry above asked for one more build. That evening supplied four, and they
are unanimous. Counted the same way — one vote per build, so a bad night cannot
vote six times:

    e9222a8   4 of 4 crashes   pane — collecting deck evidence — scanning
    3495fb8   5 of 5           same
    557b10e   1 of 1           same
    4275306   4 of 6           same (the other two: one probe, one selftest)

**That is 9 builds on the scan against 4 on `draw`** — 14 of the day's 17
records, in a band of 441-572s. The sentence below ("no phase dominates") was
true of fifteen builds and is not true of nineteen. `collectDeckEvidence`'s own
comment blamed `slideShots`, the screenshot loop, as "the prime suspect"; the
per-phase tracing it added in the same breath has now acquitted it. Every one of
these died in `listChartsInDeck({ withInventory: true })`, the FIRST phase,
before a single screenshot was attempted.

**And what it costs is not the round — it is the FILING of the round.** Every
verdict is in before that scan. 33 of 49 crash records hold a complete
fourteen-scenario result, over 13 builds; on 2026-08-29 one 4:3 leg produced a
full result five times, 14/14 in four of them, and archived none of the five.
The pane now assembles and banks its run log BEFORE the scan, and
`scripts/salvage-crashed.mjs` recovered 22 of the historical ones into
`rounds-salvaged/` — all 22 of them 4:3, taking that arm from 26 rounds to 48.

**A rate nobody had tracked**, and it is a reason to watch rather than a finding.
Crashes per archived round, by day:

    08-24  0.14      08-27  0.75
    08-25  0.12      08-28  1.33
    08-26  0.23      08-29  1.24

Roughly tenfold, breaking on 08-27. The cause is unknown and the denominator is
sensitive: a leg that burns seven attempts contributes seven crashes and zero
rounds, so a single bad leg moves the ratio a long way. Round duration is not it
— 12 scenarios/364s then, 14/390s now. What else changed that day is the domain
move and v0.4.0. Nobody should conclude anything from this yet; it is here so the
next reading has something to be compared against.

**THE NEXT READING — 2026-08-30, and it is 0.29** (4 crashes, 14 rounds, five
cycles). Against 1.24 the night before, and the obvious reading of that is wrong.

Nothing shipped that night prevents a crash. PowerPoint still falls over at
425-427s — four times, in the same band it has occupied all week. What changed is
that a crash stopped costing the leg: the 4:3 arm burned seven attempts for zero
rounds on 08-29 and two attempts for a filed round on 08-30. **The numerator
barely moved; the denominator did.** That is precisely the sensitivity the
paragraph above warned about, working in the useful direction for once, and it
means this ratio measures round SURVIVAL rather than host health.

Two readings that would be wrong: "crashes fell 4x" (they did not — the rate
did), and "the tail bound fixed it" (the bound has never fired in 14 rounds; it
covers a hang, and these are crashes). What actually did it: `deck-missing`
recovery reaching the file list, `wrong-size` stopping poisoning it, and the
download button no longer being flipped mid-tail.

If a per-round crash figure is wanted for HOST health rather than loop health, it
needs attempts as the denominator, not archived rounds. Nobody has counted
attempts.

So — of the sentence that follows — no phase dominated across FIFTEEN builds.
The per-phase traces did their job — before them, six of these would have been
filed under the probe re-ask — and what they showed then was a host that falls
over in several places late in a long round, not one broken call.

**The host says the channel was healthy.** `crashes/2026-08-28T01-51-11.md` is
PowerPoint's own account of the 524s crash in round 290: the document channel
answered `200` on every one of its last twelve `GetUpdates` calls, and every
console error in the window belongs to Microsoft's own infrastructure. Nothing
in it points at us. One number worth keeping: **87,331 document-channel requests**
by the time it fell over. That is what a ten-minute round costs a live
PowerPoint session, and it is the scale at which this happens.

**Severity, stated plainly:** the scenarios pass and the round archives 14/14.
This costs the EVIDENCE COLLECTION after a run, not the run. That is why it has
been survivable for a month and why it has never been urgent — and it is also
what stopped the 4:3 leg on 2026-08-28.

### A slide added in this session cannot be deleted by the id it was added under — round 369, 2026-09-03

The by-index insert fix (`6438dd1`) landed and works: round 369 drew all **11
batches** onto a freshly-added slide where every prior round died on the first,
and the trace carries a signature this archive had never produced —
`8x insert|resolved the target slide to an index`.

What it uncovered is the SAME defect one call further on. Both own-slide
scenarios still fail, and neither now mentions `GeneralException`:

    a big chart on a slide of its own   the scenario could not delete the slide it added
    stop a run mid-draw                 the slide this drew on could not be deleted

**The cause, in the host's own words.** A slide is added under one id and
re-keyed once it settles. Round 369 recorded the swap twenty times:

    re-acquired the scratch slide by position instead of adding one
      {"was":"4123571139#123571113","now":"282#3064692630"}

and the probe layer's positional sweep states the conclusion outright:

    "delete-by-id left slides behind and this host does not list the ids
     they were added under"

So `deleteSlideById` gets an add-time id, `slideIds()` no longer lists it,
`deleteSlideByPosition` finds `indexOf === -1`, and it returns false.

**That refusal is CORRECT and must stay.** The comment above it was bought with
round 124's receipt — 62 scratch slides reported swept and zero actually gone.
An index whose id does not match is not the slide we added, and deleting it
would remove a slide from someone's presentation with no undo. Do not "fix"
this by relaxing the check.

**The blast radius is the HARNESS, not the pane — measured, not assumed.** The
first draft of this entry warned that the pane's Tidy button was affected. It is
not. Tidy walks `tidyable`, and `tidyable` is `deck.newSlides`, which
`app.ts:3826` computes as `idsAfter.filter(...)` — a diff over the deck listing
read AFTER the round. Those are SETTLED ids, so they are listed, so
`deleteSlideByPosition` finds them. Round 369 shows one working:
`gave the scratch slides back {"returned":1,"shrankBy":1}`.

What breaks is narrower: a caller holding the id `addSlideForChart` returned,
which is add-time. That is `selftest.ts:1918` and `selftest.ts:2058` — the two
scenarios that fail — and nothing a user presses.

**The shape of the fix, not the fix.** The insert path survived by resolving the
id to an INDEX once, early, while the add-time id was still listed, and then
carrying the index. Deletion has no such early moment — the caller holds a
string across an arbitrary gap. Options, none costed yet:

1. `addSlideForChart` returns a handle carrying both the id and the index it was
   added at, and the caller re-verifies the slide at that index before deleting.
2. Tag the slide at add time and delete by tag lookup, which survives re-keying.
3. Leave deletion refusing, and make the CALLERS say so — the scenarios would
   then report "the host will not let this be cleaned up", which is true, rather
   than failing.

(3) is honest and cheap and does not touch a destructive path; (2) is the real
fix. This wants an owner decision, because option (3) changes what two
long-running scenarios report and the archive compares scenarios by name.

**Not urgent, and not a regression.** These two scenarios were already failing —
on `GeneralException`, which was worse. The 10 loose shapes the grouping check
flagged this round are the deliberately-aborted draw that this defect then could
not clear, not a grouping fault: 11 of 11 attempts grouped, 0 refused.

### The death ratchet ranks EXPOSURE, not danger — read it accordingly, 2026-09-03

`FATAL_SCENARIO_BUDGET` works: it is a "no new deaths" gate and it fires
correctly. But its ORDERING invites a wrong reading, and the wrong reading was
mine until I measured it.

Raw deaths say `same scale across the deck` (10) is the most dangerous thing the
harness does. It is not. Against round 369's per-scenario durations:

    scenario                                dur    deaths
    same scale across the deck              146s      10
    a big chart on a slide of its own       107s       4
    does a rasterise poison the next draw    83s       0
    stop a run mid-draw                       9s       8

`stop a run mid-draw` takes NINE SECONDS and has collected eight deaths, while an
83-second scenario has collected none. Per second of exposure it is roughly 30x
more lethal than `same scale` — which is what you would expect from the one
scenario that deliberately abandons a draw mid-flight and leaves the host holding
a half-issued batch.

`same scale` looks worst for a duller reason: it is the longest scenario, it runs
late, and every one of its ten deaths landed in the 420-600s window inside
`updated only the shapes that changed`, with single syncs taking 9-20s. That is
the session-age crash this archive has documented all month, arriving while the
longest scenario happens to be open. Attribution credits whoever holds the floor.

**What this does NOT mean.** It does not mean the budget is wrong or should be
normalised. A per-second rate cannot gate anything — the honest gate is "no new
deaths", and that is what is shipped. This is a READING note: do not rank the
budget's entries by size and call the top one the worst offender.

**What it suggests next.** `stop a run mid-draw` is the cheapest place to buy
host stability, because it is nine seconds long and kills the host eight times.
Whether the abort path leaves a batch half-issued is answerable from the trace
and has never been asked. Note it also just changed: `6438dd1` altered the slide
targeting under it, so its next deaths are not comparable to its old ones.

### The by-index fix works at 16:9 and cannot work at 4:3, because the slide is GONE — round 370, 2026-09-03

> **SUPERSEDED, and the title is wrong in the way that matters: it is not 4:3.**
> Round 373 ran the SAME deck at 16:9 and lost the slide there too, and the
> crash separates by deck shape rather than aspect ratio — 5% on a one-master
> deck, 100% on a two-master one. The observations below are sound; every
> inference from them to "4:3" is not. See "FIXED: the '4:3 crash' was a
> two-master crash, and a documented API misuse" at the end of this file.

I verified `6438dd1` on Presentation64 (round 369, 16:9), wrote that the fix was
verified on a host, and generalised. Round 370 on Presentation70 (4:3) refutes
the generalisation, and the reason is more interesting than the fix.

At 4:3 both own-slide scenarios throw `GeneralException` again — but at a new
place:

    errorLocation: "SlideCollection.getItemAt"
    fullStatements: ["var itemAt = slides.getItemAt(7);", "itemAt.load([\"id\"])", ...]

That is MY call, so the first reading is "the fix is wrong". It is not. The
resolution was correct and the trace proves it:

    host    slides added                          requested=1 landed=1 from=7
    insert  resolved the target slide to an index chart=draw-11 index=7
    draw    batch issued                          chart=draw-11 total=103
    error   drawing the chart's shapes            GeneralException @ getItemAt

Deck of 7, one added, so the new slide is index 7 of 8. Correct. Then
`getItemAt(7)` throws — and the NEXT scenario opens with `deckSlides=7`.

**The slide does not persist.** It is listed, `landed=1` confirms it against a
before/after diff of the deck's own listing, and moments later the deck is back
to seven. This is the failure this repo already has receipts for — `slides.add()`
acknowledged and absent under load, whole decks taken and never landed — and it
is the one thing no targeting strategy can survive. By id or by index is a
question about how to NAME a slide. There is no slide.

So the two results are not in conflict:

    16:9  the slide persists; targeting WAS the bug; by-index fixed it
          (round 369: 103 of 103 shapes, 11 batches, grouped and tagged)
    4:3   the slide evaporates; targeting is downstream of that
          (round 370: same scenarios, same exception, new location)

**What is proven and what is not.** Proven: the index resolved correctly; the
add reported landed; `getItemAt` on that index threw; the deck was back to 7
afterwards. NOT proven: that the slide vanished as opposed to `getItemAt`
refusing a slide that is listed but not yet settled. Both fit the evidence. The
distinguishing measurement is cheap and has never been run — after the add,
read `slides.getCount()` in the SAME batch as the `getItemAt`, and record both.
One number separates "the slide is gone" from "the host will not hand it over
yet", and they need completely different fixes.

**Do not revert `6438dd1`.** At 16:9 it turned 1 batch into 11 and it is the only
reason a chart has ever drawn onto a slide this product added. At 4:3 it fails
where the old code also failed. It is strictly better and it is not the cause.

**Guard worth costing, not yet built.** `getItemAt` throwing takes down the
whole `PowerPoint.run`, and it kills the host often enough to be one of the
budget's top entries. A bounds check against `getCount()` in the same batch,
falling back to the by-id path when the index is past the end, would convert a
`GeneralException` into the older, survivable failure. That is a change to a
draw path that four previous unit-green fixes have failed on, so it wants the
measurement above FIRST — not a fifth guess.

### The four owner decisions, resolved — 2026-09-03

All four were escalated as needing a person. Researched, decided and built.

**1. The ratchet was red. It is now a RATE, and the count could never have
worked.** Deaths only accumulate, so a budget seeded at today's total breaches
on the very next death — and for a scenario dying all week, that is Tuesday. It
fired on the episode continuing, not on a regression. The count also RANKED
WRONG, because deaths arrive in episodes rather than at a steady drip:

    same scale across the deck    all 9 deaths 08-24..08-29, none since
    stop a run mid-draw           all 8 on 09-02 and 09-03
    a big chart on a slide of its own   all 5 today, on a scenario written today

Against exposure the order inverts, and the two young scenarios are the killers:

    a big chart on a slide of its own    5 /  11 runs = 455 per 1000
    stop a run mid-draw                  8 /  25 runs = 320
    same scale across the deck           9 / 418 runs =  21
    every other listed scenario          1 / ~420     =   2.4

455 and 320 per 1000 is a host killed on nearly half, and on a third, of the
runs. The scenario that led on raw count kills on 2% of its runs.

The gate now compares deaths against `p*n + 2*sqrt(p*(1-p)*n)` — a rate seeded
at today's, with a binomial noise bound. Without the bound a rate is still a
hair trigger: at 8-of-25 a ninth death reads as 346 against a ceiling of 330,
red for a scenario doing exactly what it always has. THE COST IS STATED AND
REAL: at eleven runs this cannot tell 45% from 80%, so `a big chart on a slide
of its own` needs nine deaths in eleven to trip. That is a property of eleven
samples, not a choice. The printed table is the surveillance; the gate is for
what the evidence can carry.

And a rate CAN FALL. Runs accumulate, so a scenario that stops dying sinks on
its own and a landed fix protects itself with nobody editing a number — the
property this instrument needed, which a cumulative count can never have, and
which an earlier docstring promised and had to retract.

**2. `6438dd1` stays.** At 16:9 it turned one batch into eleven and is the only
reason a chart has ever drawn onto a slide this product added. At 4:3 it fails
where the old code also failed, so it is strictly better and is not the cause.
The rate table is the strongest evidence for the fix mattering: the two worst
scenarios in the suite are exactly the two own-slide ones.

**3. The discriminator is built, and it is not a probe.** On a draw failure the
insert path now takes one fresh-context deck read and traces which of the three
things is true — the index is past the end (the slide is GONE), the slide is
listed under a new name, or it is listed under the same id and was refused
anyway. Zero cost unless a draw has already failed.

A probe was the obvious vehicle and is the wrong one: the probe sheet asks early,
on a scratch slide, in a healthy session, and this failure arrives 670 seconds
in, on a 4:3 deck, behind a 103-shape batch. A probe would not reproduce it and
might never sample it. This measures in the real conditions or not at all.

**4. The delete reporting was inventing wreckage.** Both own-slide scenarios
read `deleteSlideById` returning false as "a slide is stuck in the deck". At 4:3
it means the opposite: the slide was already gone, so the delete correctly
refused to remove an index whose id did not match. The scenarios announced "the
aborted draw is still in the deck" about a deck that had never been cleaner.

Wrong in the one direction a self-test must never be wrong. A missed problem
costs a round; an invented one sends someone looking for wreckage that was never
there. The DECK now decides, not the return value, and the pass detail says
which of the two happened.

The fake could not previously express "refused, AND gone" — it offered a refusal
that leaves the slide or a success that removes it, and nothing in between. That
is why no test caught this. `faults.slideVanishesInsteadOfDeleting` now
reproduces round 370's pairing.

**Still open, and now the sharpest question in the project:** whether the slide
is lost or merely withheld at 4:3. The next 4:3 round answers it in one trace
line. `a big chart on a slide of its own` kills PowerPoint on 45% of its runs
and the path it walks is one a user walks — accepting "put it on a slide of its
own" — so this is the product's worst defect, not a harness problem.

### MEASURED: the slide is added, listed, and gone before the first batch — round 371, 2026-09-03

> **The measurement stands; one conclusion in it does not.** This entry says the
> vanish is "NOT the documented server-side lost edit ... zero steps of round
> 371". That search was of the add-in's own step stream, which structurally
> cannot carry PowerPoint's ULS. The ULS is in `crashes/*.md`, and
> `errorLocalChangeLostSingleUser` is there in 17 of them — the only ErrorName
> in all 109. It IS the lost edit, and the cause was sending a layout without
> its slide master. Fixed; see the last entry in this file.

The fork round 370 could not settle is settled. The diagnostic fired twice in
round 371 and both samples agree:

    index 7 · deckSlides 7 · indexInRange false · stillListedUnderTheSameId false
    reading: "gone — the index is past the end of the deck"

Not withheld. Not re-keyed. NOT LISTED AT ALL, under any name, seconds after
`addSlideForChart` confirmed it landed by diffing the deck's own listing.

**AND THE TIMING IS TIGHTER THAN THE LAST ENTRY SAID.** That entry left open
whether the slide vanished or the host merely refused a slide it still held.
The trace closes it, and closes it earlier than expected:

    host    slides added                          requested=1 landed=1 from=7
    insert  resolved the target slide to an index index=7
    draw    batch issued                          total=103 onSlide=0
    error   drawing the chart's shapes            GeneralException @ getItemAt
    insert  the draw failed — asking the deck     deckSlides=7 ... "gone"

`findIndex` returned 7, which it can only do against a listing of EIGHT. So the
slide was there at the resolution sync. `getItemAt(7)` then threw, which it can
only do against a deck of SEVEN. Both syncs belong to the same `PowerPoint.run`,
milliseconds apart.

So the slide is lost BETWEEN TWO SYNCS OF ONE CONTEXT, before a single shape is
issued. **The `GeneralException` is the consequence, not the cause** — the draw
never had a slide to fail on.

**What this rules out.** Every targeting strategy, permanently. By id, by index,
by tag, by creation handle: there is no slide to name. The four ruled-out fixes
and the by-index fix that replaced them were all answering the wrong question at
4:3. `6438dd1` is still right — at 16:9 the slide persists and it is the only
reason a chart has ever drawn onto one this product added — but it cannot be
extended to cover this.

**What it does NOT explain, and I checked rather than assumed.** Not the
documented server-side lost edit: `errorLocalChangeLostSingleUser` appears in
eight older crash records and in ZERO steps of round 371 or its crash record. So
the obvious mechanism is not this one. Unknown remains unknown.

**Where the fix has to go.** At the ADD, not the draw. Three shapes, uncosted:

1. Verify the slide SURVIVES a round trip before drawing on it — one extra
   listing after the add, and treat a slide that does not persist as a slide
   that was never added. `addSlideForChart` already returns null on other
   failures and `app.ts` already falls back to the slide the user was on.
2. Retry the add. Cheap, and the archive has no evidence either way about
   whether a second add sticks.
3. Do not offer at all on a deck where adds do not persist, which needs a way to
   know that in advance and probably cannot be answered.

(1) is the honest one: the pane's offer is "put it on a slide of its own", and a
slide that evaporates makes that a promise the product cannot keep. Turning a
`GeneralException` — on the path that kills PowerPoint on HALF its runs, 500 per
1000 as of round 371 — into a quiet fallback to the visible slide is strictly
better for a user than what ships today.

IT IS A PRODUCT DECISION, because it changes what someone who accepted the offer
receives: a chart on their current slide instead of an error and a dead host.
That is the owner's call, and it is the last thing standing between this defect
and a fix.

### FIXED: the "4:3 crash" was a two-master crash, and a documented API misuse — 2026-09-03

**The slide stopped vanishing.** Round 374, Presentation70 at 4:3, first
attempt, no crash:

    slides added  {requested:1, landed:1, from:7}
    slides added  {requested:1, landed:1, from:8}

The deck went 7 -> 8 -> 9 and KEPT both. Every prior 4:3 round went 7 -> 8 ->
back to 7 with the slide gone under any name, twice a round. Zero "gone"
readings in 374; the one draw that still failed reports `deckSlides:9
index:8 indexInRange:true stillListedUnderTheSameId:true` — the slide is
there.

**The cause.** `AddSlideOptions` is explicit: a `layoutId` sent without a
`slideMasterId` "needs to be available for the default Slide Master ...
Otherwise, an error will be thrown", and that default is THE PREVIOUS SLIDE'S
master. `blankLayoutId` took the blank layout of the FIRST master and
`addSlides` sent it alone. One master, always legal — which is why it ran for
months and why the fake, which had one master, could never catch it.

**The evidence that named it**, over 220 archived rounds bucketed by the
`layouts-readable` probe on whether the round crashed:

    1 master  /  1 layout   @ 16:9     9 of 182  =   5%
    2 masters / 12 layouts  @ 16:9     4 of   4  = 100%
    2 masters / 12 layouts  @ 4:3     27 of  30  =  90%

Hold the deck and the aspect ratio changes nothing; hold the ratio at 16:9 and
the deck moves it 5% -> 100%. Round 373 confirmed it by experiment rather than
correlation: Presentation70 at 16:9 lost the slide twice, exactly as at 4:3.

**Two weeks of "4:3" in this file is mis-attributed.** 4:3 was along for the
ride because `cyclePlan` has pinned the tall arm to one deck since 2026-08-26.
Entries above that reason about aspect ratio should be read as reasoning about
`Presentation70`.

**And it IS the documented lost edit.** PowerPoint's ULS carries
`errorLocalChangeLostSingleUser` in 17 crash logs and it is the ONLY ErrorName
in the corpus of 109 — it discards the rejected revision, resets undo history
and rolls the deck back, which is precisely "listed, listed again, gone". An
earlier entry here said it was NOT that, on the grounds that the string appears
in zero steps of round 371. That search was of the add-in's own step stream,
which structurally cannot carry the host's ULS. Wrong scope, wrong conclusion.

**What is fixed and what is not.**

Fixed: the add. `blankLayoutTarget` returns `{layoutId, slideMasterId}` and
every add — including the retry inside `addSlides` — sends both.

NOT fixed, and still open above: the two own-slide scenarios still fail, now on
DELETING the slide they added rather than drawing on it. That is the add-time
vs settled id problem, unchanged and unrelated.

NOT established: that the crash rate is fixed. One clean 4:3 round against a
~49%-per-attempt history is encouraging and is not a rate. The vanish is
settled by mechanism — slides persist, the deck grows, the readings say
"there" — but the crash needs rounds, and the rate ratchet will show it
falling on its own if it is real.

### 17 of 17 at 4:3, and a correction to why the master fix works — round 377, 2026-09-04

**Round 377 swept the deck that had never swept.** Presentation70 at 4:3, build
`6dfaa4b`: 17 passed, 0 failed, 0 skipped, no crash. That deck crashed 27 of
its 30 archived rounds and both own-slide scenarios had been red for weeks.

Its 16:9 control the round before — 376, Presentation64 — also went 17 of 17,
so nothing regressed on the main arm.

**AND IT CORRECTS THE MECHANISM I COMMITTED.** `6dfaa4b` argued that the deck's
own master could not be the first one, because the add had been rejected for a
layout "not available for the default Slide Master" and the default is the
previous slide's. Round 377 read both and they are the same master:

    slideMasterId (from slideMasters)      2147483660#2460954070
    deckMasterId  (from Slide.slideMaster) 2147483660
    matchedTheDeck                         false     <- WRONG

One master, two renderings, and a `===` between them that reported a mismatch
about a deck that matched perfectly. The same id-space split this project
already documents for slides, one level up, and I walked into it while writing
the check that was meant to detect it.

So the corrected reading: sending `slideMasterId` did NOT fix a wrong master.
It disambiguated a layout id the host would not resolve against the default on
its own. The fix is unchanged and the measurements are unchanged — slides
persist, the deck grows and keeps them, three 4:3 rounds in a row without a
crash — but the mechanism is "the pair resolves where the layout alone did
not", not "we were naming another master".

The comparison now keys on the part before the `#`, and the fake renders the id
both ways so the normalisation is testable at all: with it removed, the test
falls back to the wrong master and fails.

**What is now established.** The add is fixed and verified on both decks and
both aspect ratios. The recovery for a slide that does vanish is verified
(round 373, fired twice). The cleanup by set difference is verified (rounds
376 and 377, "slide swept", no refusal note).

**~~What is still not established.~~ THE RATE IS NOW ESTABLISHED, 2026-09-05.**
Every 4:3 round that carries the own-slide scenario, split on whether its build
contains the fix (`6dfaa4b`):

    PRE-fix    370  371  372  375                        4 of 4 FAILED
    post-fix   377  378  382  383  385  386  387         15 of 15 passed
               388  389  390  391  392  393  394  395
               384                                       failed elsewhere

Fisher's exact, one-sided, counting 384 as a failure so no judgement is
required of the reader: **p = 0.0010**. Excluding it on its own evidence — its
detail says the CONTROL insert onto the visible slide hung and that "nothing
here says anything about a freshly-added slide" — p = 0.0003. The conservative
number is the one to quote.

Three of the four pre-fix failures are `GeneralException at=drawing the chart's
shapes`; the fourth could not delete the slide it added. None of that has
recurred in fifteen attempts, and **385 through 395 are eleven consecutive
rounds at 18 of 18** — the longest clean run at 4:3 in the archive.

**Round 374 is absent from both columns, and the rounds gate is right to flag
it — but the disagreement is now resolved.** The gate reports it as "filed
under a profile the driver did not measure": the archive says 960x540 from
`exportedSlide`, `driverSlideSize` says 4:3, `driverDeck` says
Presentation70.pptx. The archive is the correct reading. Round 373 deliberately
held Presentation70 at 960pt to prove the crash was not about aspect ratio, and
that resize was still in place when 374 ran an hour later; 375 is back at
720x540, so the deck was reset between them. `driverSlideSize` is echoing what
the round was ASKED for, `exportedSlide` is measuring what it got, and on this
round they honestly differ. 374 is a 16:9 round on the 4:3 deck. It is
pre-fix (`40dfee0`, 2026-09-03 23:01, against the fix at 2026-09-04 08:26) and
it failed the own-slide scenario — consistent with the pre-fix column it is not
counted in.

Originally written as: "Three clean 4:3 rounds against a 27-of-30 history is a
strong signal and is not yet a rate." It took seven, and fifteen now stand. The
rate ratchet moved on its own throughout, which is the property it was rebuilt
for — `stop a run mid-draw` fell from 320 to 222 per 1000 with nobody editing a
number.

### The set-difference delete is PARTIAL — it works only after the slide settles, 2026-09-04

`deleteSlideAddedSince` (89604bc) finds the added slide by diffing the deck
listing at delete time, so it is immune to the id having CHANGED. It is not
immune to the id not having changed YET.

Round 378 caught it. `stop a run mid-draw` aborts early by design, so its
cleanup runs seconds after the add — before the host re-keys. The diff then
returns the ADD-TIME id, which is exactly the name `deleteSlideById` cannot
resolve, and the scenario failed on the old sentence:

    deck 7 -> 8, "stop requested", reading "there" (deckSlides 8, in range)
    then: "the slide this drew on could not be deleted"

Rounds 376 and 377 passed the same scenario because their cleanup ran later,
after the re-key, so the diff handed over the settled id. Same code, different
timing, opposite result — which is why one green round was not evidence.

**Where it actually breaks.** `deleteSlideByPosition` looks the id up in a fresh
listing, finds it, then re-reads the id through a by-index handle and compares:

    if (loadedValue(() => slide.id) !== slideId) return;

The listing says `4123571151#123571113` and the by-index handle answers the
settled form, so the comparison fails and it declines. Normalising on the `#`
does NOT help here as it did for masters — an add-time slide id and its settled
form share no prefix (`4123571151#123571113` against `256#2587447327`).

**What would fix it, and why I have not built it.** The diff already knows the
INDEX of the fresh id in the listing it just read. Deleting at that index and
verifying both that the deck shrank by one AND that the fresh id is no longer
listed would be sound. But that is a destructive path guarded by a check bought
with round 124's receipt — 62 scratch slides reported swept and none deleted —
and it deserves its own careful pass rather than being bolted onto a session
that is already several fixes deep. Filed rather than rushed.

**Severity: harness only.** The product's Tidy walks `deck.newSlides`, which is
a diff over the POST-round listing, so it holds settled ids and works. This
costs one intermittent scenario failure, not a user anything.

**RESOLVED BY THE ROUNDS, NOT BY A TEST — 2026-09-05.** `bc9a386` made the diff
re-read and re-take the difference over three passes 400ms apart, on the theory
that the fault was the id not having settled YET rather than having changed.
That theory could not be tested: the fake settles an id after N LOOKUPS and one
delete attempt makes several, so it always settles within a single attempt,
where the host's settle is a matter of TIME; and the call site is
`deleteSlideAddedSince(...) || deleteSlideById(slideId)`, so a failed diff falls
through to the old route and the verdict is identical either way. Cutting the
loop to one pass changed no test, checked twice after two failed fixture
attempts. The docstring said so rather than letting a green suite imply
coverage, and named the rounds that would adjudicate it.

They have. `stop a run mid-draw` failed this cleanup in 378, 382 and 383, and
passed it at 4:3 in **385, 386 and 387** — three in a row, all 18 of 18. The
retry is doing the work. It stays, and it stays labelled as verified by the
harness rather than by the suite.

The positional delete described below was NOT built, and should still not be:
the diff knows the index and it would have worked first time, but deleting an
unconfirmed index is what round 124 punished. The cheaper fix was enough.

### ITEM 3, MEASUREMENT 1: an inserted .pptx does NOT bring its own master — round 379, 2026-09-04

The kill switch is cleared. Four generated-deck inserts on Presentation64, the
ONE-master deck, every one with `KeepSourceFormatting`:

    {"expectedSlides":2,"landed":2,"base64Bytes":112852,
     "formatting":"KeepSourceFormatting","mastersAfter":1}   x4

The deck stayed at one master. PowerPoint maps the incoming master onto the
destination rather than importing it, so a generated slide cannot push a
customer's deck from one master to two — which was the risk that would have
killed this whole line of work, because a two-master deck is the state that
crashed 90-100% of this archive until 2026-09-04.

It matters that the file DOES carry one. Unzipping a generated .pptx shows
`ppt/slideMasters/slideMaster1.xml` and `ppt/theme/theme1.xml` of its own. The
host declines to bring them in; the file is not innocent, the host is careful.

The two-master deck agreed earlier by accident: masters stayed at 2 across
rounds 377 and 378 (n=3 after the last insert against 35-37 before). This is
the reading that was missing — 1 staying 1 — and it is the case a customer with
a clean template would meet.

**So option B is unblocked, and it is now a product decision rather than a
technical one.** What remains is not "can we", it is "should an over-budget
chart take a slide of its own on the web, and does it ask first".

**The measured case for it**, all from this archive:

    native, 103 shapes, own slide     46.4s over 11 batches, and superlinear:
                                      161ms per shape at the start, 784ms at
                                      the end (rounds 374-378: 44.6-49.1s)
    the same chart as a .pptx         280ms to build, ~152KB base64, ONE call
    reliability of that one call      352 of 353 archived rounds already use it
    fidelity                          HIGHER: the pptx sink draws filled
                                      polygons, exact arcs and text alpha that
                                      Office.js cannot (src/core/scene.ts)
    editability                       kept — ooxml.ts writes POWERCHART_CONFIG
                                      on the group, the same tag the pane
                                      re-edits a drawn chart by

**The case against, which is not technical.** The user asked for a chart on
THIS slide and would get one on a new slide. `insertSlidesFromBase64` inserts
SLIDES after a target slide — it can never place shapes on a slide that already
exists — so there is no version of this that honours the original request. That
is the whole cost, and it is a product judgement.

**What the industry does, for calibration.** Every web-capable competitor ships
a picture on purpose: Datawrapper (the closest analogue, a real Office.js
add-in) inserts a static image by published choice; Power BI's own add-in
snapshots to an image; Mekko Graphics' founder says "our charts are just
pictures within PowerPoint"; UpSlide ships vector pictures. The native-shapes
champions — think-cell, Deckary, Power-user — are all DESKTOP ONLY and never
face this constraint. Doing native shapes on the web would be unusual, and
would be the product's differentiator rather than its convention.

### The death ratchet is falling on its own, which is what a landed fix looks like — 2026-09-04

The rate instrument was rebuilt on 2026-09-03 with one argument for choosing a
rate over a count: **a rate can fall**, so a fix that works protects itself
without anyone editing a number. That claim can now be checked against a day of
rounds, and it holds.

Tracked across today, the two own-slide scenarios:

    a big chart on a slide of its own   455 (5/11) -> 500 (7/14)
                                        -> 524 (11/21) -> 393 (11/28)
    stop a run mid-draw                 320 (8/25) -> 296 (8/27)
                                        -> 276 (8/29) -> 222 (8/36)

The death COUNTS are flat while the run counts climb. Nothing was edited; the
rate fell because clean runs accumulated under it. Both are now well below the
ceilings they were seeded at.

**And the deaths stopped where the fix landed.** Attributed across sound crash
records, split at `40dfee0`:

    deaths on PRE-fix builds    33
    deaths on POST-fix builds    0

Zero, across every arm and every scenario.

**What that does and does not establish.** It establishes that the own-slide
scenarios have stopped killing the host: `a big chart on a slide of its own`
has run seven more times and `stop a run mid-draw` eleven more, both without a
death, against pre-fix rates of ~500 and ~320 per 1000. At those rates the
chance of seeing none is roughly 1 in 100 and 1 in 60 respectively.

It does NOT establish the 4:3 crash rate, and the difference is worth keeping
straight. Post-fix the 4:3 arm has FIVE archived rounds with no crash — 374,
375, 377, 378 and 382 — and ONE crash record, which died at 128.9s in the PROBE
phase on `creationid-survives-grouping`, with no scenario open. That is a
different population from the own-slide defect, and there is no reason the
master fix would have touched it.

One crash in six attempts against 55 in 113 is suggestive and not yet
conclusive: at the old rate, seeing at most one in six happens about 11% of the
time by chance. Around ten clean attempts is where this stops being a hopeful
reading, and each one is a ten-minute round.

So: the mechanism is fixed and the deaths it caused have stopped, measured. The
ARM is not yet demonstrably healthy, and saying otherwise would be reading a
scenario-level result as a deck-level one.
