# Evidence kept beside the claim that rests on it

Not rounds, and not tests. These are the one-off measurements that decided
something a document asserts, kept so the assertion can be re-run rather than
re-argued.

## `17-deck-scan-second-page.json`

The run log behind BACKLOG item 17 — the first deck scan on record to return a
complete SECOND page (29 slides at `from:20`), after 379 rounds where `from` was
always 0.

## `run-scope-arms.mjs`, `run-unsettled2.mjs`, `run-split.mjs`

The three arms that killed office-js drafts A and B on 2026-09-06 and produced
DRAFT D in their place. Run in order against `Presentation72` with the pane
open; each evaluates plain Office.js inside the pane's frame through
`pw eval <script> <pane-ref>`, the same route `slideResolveScript` uses in
`scripts/round.mjs`.

What they establish, in order:

1. **`run-scope-arms.mjs`** — the document's own slide, a slide added this
   session, and a slide the add-in made in a PREVIOUS session all tag fine.
   That is what killed Draft A: provenance is not the variable.
2. **`run-unsettled2.mjs`** — on slides added moments ago the shape collection
   reads 0, twice, while the same run reads 6 and 7 on the two older slides. Its
   first version crashed on its own assumption (`items[listed - 1]` with
   `listed` 0) and reported six THREWs that were this script's bug, not the
   host's; separating `empty-read` from `THREW` is why the second version
   exists, and is the whole reason the conclusion changed.
3. **`run-split.mjs`** — asks by index and by id in SEPARATE syncs, because the
   first attempt queued both and the 5010 from the id killed the sync before
   the index read could answer. It is what produced the add-time id
   (`4123571130#123571113`), the 5010 on `getItem` with it, and the 0-shape
   reads by index and by settled id, 3 of 3.

**They are kept as scripts rather than as output** because the output is three
lines of JSON and the reasoning is in what each arm controls for. Read the
headers; they say what a passing arm would have meant.

**They mutate the deck they run on** — shapes, and a slide per trial. That is
why they name `Presentation72`, a throwaway, and why nothing here should be
pointed at a deck anyone cares about.

## `windows-desktop-first-run.mjs`, `windows-desktop-2026-09-13.json`

**The first reading this project has taken on any host but the web.** All 430
archived rounds report `platform: OfficeOnline`; Microsoft's validator says a
submission is tested on Windows and Mac too, so the two things the store listing
claims there had never been measured — that the host supports what the product
needs, and that a chart arrives as native shapes rather than a picture.

Both, on AITEST, Office 16.0.20326.20144, Current Channel:

    platform                      PC
    PowerPointApi 1.1 … 1.10      all true
    the default chart inserted    1 shape, msoGroup, `PowerChart`, 40 items
    pictures on the slide         0

The chart is the pane's own default, and that is the point rather than
convenience: `stacked` ships `decorations.cagr`, a CAGR arrow is an `arrowhead`
node, and an arrowhead is exactly what `marksThisHostWillDrop` drops when
`canRotate()` is false. So the default chart is the one that would have exposed
a picture fallback, and it did not.

**Verified with COM, not from the pane's message.** "Scaled to fit the space
left on the slide" is what a SUCCESSFUL picture insert says too; only counting
the shapes tells them apart. The script says so and prints the one-liner.

WHAT IT DOES NOT COVER, so it is not over-read: Mac, which the manifest also
claims and which remains unmeasured; volume-licensed/LTSC Windows, where
Microsoft lists 1.6 through 1.10 as **Not available** and which this machine
cannot represent because no update reaches it; any build below 19610.20002, the
1.10 minimum — this is one host in the good tier; and anything beyond a single
insert of one kind. **n = 1.**

**It needs a person once, and that is structural rather than an oversight.**
Adding an add-in from a Shared Folder catalog is UI-only — no COM, no command
line — and the CDP route cannot bootstrap itself, because the WebView only
spawns when the ribbon button is clicked. The script's header carries the four
clicks and the launch flag.

## `bound-empty-read.mjs`, `empty-read-6-trials.jsonl`

Draft D's measurement: how long a shape drawn on a freshly added slide stays
invisible to `shapes.load("items/id")`. Six trials on `Presentation73`, a clean
six-slide deck, polling every ~6s for two minutes with a control read on a
pre-existing slide beside every reading.

The answer is that it does not become visible at all inside two minutes: `seen`
is 2 on all 124 reads, the control is 3 on all 124, and a screenshot shows the
drawn rectangle on the slide next to the two placeholders the API does report.

TWO SCARS, both worth more than the result:

- The first version stopped as soon as the collection was non-empty. A slide
  from `slides.add()` arrives carrying two layout placeholders, so every trial
  "filled" on the first read and never looked at the rectangle. It produced
  "1.87 - 1.93 seconds", which matched a workaround guessed in someone else's
  2023 bug report, and the match was manufactured. **A trigger that cannot fail
  is not a measurement.**
- The archived jsonl did not record `before`, the threshold `seen` is compared
  against — so the "two placeholders" figure was inferable from the file but not
  re-derivable from it. Fixed; older lines in the jsonl predate the field.

**A THIRD SCAR, found by a skeptic and larger than the other two.** Every read
in that harness called `shapes.load("items/id")` on both slides and then kept
only `.length`. So the archived jsonl carries no shape id at all, and the
draft's central sentence — "the one it omits is the one the add-in just drew" —
was **not re-derivable from the archive**. Two placeholders with the rectangle
missing reads identically to one placeholder plus the rectangle with a
placeholder missing. Only the SCREENSHOT distinguished them, and a screenshot is
not something a query can join against.

The harness now records `beforeIds` and `seenIds`, so the id sets can be
differenced. Lines already in `empty-read-6-trials.jsonl` predate both fields.
**The claim in DRAFT D stands on the screenshot until a re-run replaces it with
id sets** — which is a weaker footing than the prose implied, and is why this is
written down rather than quietly fixed.
