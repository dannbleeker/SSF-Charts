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
