# Changelog

Notable changes, newest first. Dates are the release date.

This file starts at 0.4.0, the release that renamed the project. Earlier
releases are in the git history and on the [releases page][releases]; there was
no changelog before this one, and inventing entries for 451 commits after the
fact would produce a document nobody could trust.

## Unreleased

## 0.6.0 — 2026-09-12

The release that stops PowerPoint falling over on a deck with more than one
slide master. If you work from a corporate template, that is almost certainly
your deck, and this is the entry to read.

### Adding a slide crashed PowerPoint, and it was ours

Everything that adds a slide went through one call: the own-slide offer below,
**Agenda slides**, and **the demo deck**. That call asked for a layout without
saying which master it belonged to — and
`AddSlideOptions` is explicit that a `layoutId` sent without a `slideMasterId`
has to exist on the **default** master, which is whatever the previous slide
uses. On a deck with one master that is always true, which is why this worked
for months. On a deck with two it is an error, and PowerPoint's answer to a
rejected add is not a message: it discards the revision, **resets your undo
history and rolls the deck back**.

Bucketed over 220 archived runs by how many masters the deck had:

    1 master  /  1 layout   @ 16:9     9 of 182  =   5%
    2 masters / 12 layouts  @ 16:9     4 of   4  = 100%
    2 masters / 12 layouts  @ 4:3     27 of  30  =  90%

Hold the deck fixed and the aspect ratio changes nothing; hold the ratio at
16:9 and the deck moves it from 5% to 100%. This was filed as "the 4:3 crash"
for two weeks and 4:3 was along for the ride, because that arm had only ever
run on the two-master deck.

The layout and its master are now sent together, and the added slide is built
from **the deck's own master** rather than whichever one came first — so a
chart added to a template's second master no longer arrives styled from the
first.

### A too-dense chart is now offered a real choice

A chart with more shapes than the web host will draw onto a busy slide used to
be rasterised **before anyone was asked** — the offer then priced a picture,
and so never fired. You are now asked, and the chart can go on a slide of its
own as editable shapes rather than as a flat image. It stays re-editable:
select it and the pane reloads its data.

Two further faults in the same path, both of which made the offer useless
rather than dangerous, are fixed with it — it handed the host a slide id the
host refuses, and then targeted the new slide by that id instead of by its
position.

- **A third button, "Insert as a picture"**, sits alongside "Put it on its own
  slide" and "Add here anyway" — offered only where this PowerPoint can
  rasterise, and never the default, because a picture is not editable.
- **The "Insert as picture" tick is remembered** between charts instead of
  resetting every time.

### Messages that dropped what you needed to know

Four, all the same shape: the pane has one message slot and the last write
wins.

- **A setback mid-insert was destroyed** by the busy note that follows it, so
  when your chosen route failed you were told only the outcome.
- **Same Scale said one of the three things it had to say.** A run that both
  rescued charts as pictures and degraded others showed only the degraded
  clause — losing the rescued count, and with it the pointer to "Explode to
  native shapes", which is the control that turns them back.
- **The insert promised a door that Explode refused.** A chart inserted as a
  picture said *"Explode to native shapes turns it back"* while Explode
  declined on the very same test.
- **The own-slide offer quoted the wrong number** when it explained itself.

### Smaller, and one you may notice

- **The shape budget rises from 90 to 105**, on the first evidence that ever
  bore on it — so charts that previously tipped into a picture now insert as
  shapes.
- **The Harvey ball, checkbox, process flow, KPI tile and table previews** are
  no longer unnamed images to a screen reader; each now carries a text
  alternative describing what it shows.
- **A frozen-host warning stopped pointing at a button you do not have.** The
  60-second "PowerPoint has not answered" message ended by naming a download
  that only exists after a test run.
- **The add-in's high-resolution icon** now points at the 64×64 asset Office
  asks for, in all four manifests, rather than an 80×80 one. Reinstall from
  this release's `manifest-prod.xml` to pick it up.

## 0.5.0 — 2026-09-02

The release the European paste fix has been waiting in. Eleven user-facing
changes since 0.4.0; the two worth reading first are that a paste out of a
Danish, German or Nordic Excel is finally read correctly, and that a chart this
PowerPoint cannot draw now arrives as a picture rather than as a gap.

### Charts stop drawing text over text

A sweep of every chart kind under twenty-four options and ten data shapes — at
eight frame sizes, two fonts and both orientations — counted **2,148 places
where this engine drew one piece of text through another**. What is left is one
open question about where a unit label belongs, plus a small tail, and the sweep
runs on every build so the number can only go down.

**Three figures appeared while this was in progress and only one of them was a
change to the engine**, so none of them should be read against another without
saying which sweep produced it: 2,148 → 467 was the fixes below; 467 → 4,010 was
the sweep itself widening to cross options with data shapes, which it had never
done; and 4,010 → 1,327 was the small-multiples bug in the next entry. A fourth
and fifth, 1,327 → 1,014 → **785**, are the unit label:

- **The value-axis unit label is clipped** to at most 40% of the chart's width.
  It is a unit — "€m", "%", "$m (log)" — and it was drawn with no width bound at
  all, so a long one ran clear across the totals row and the legend. A short
  unit is untouched; a long one is truncated with its start kept, the way
  category and gantt names already are. Put the explanation in the title or the
  footnote.
- **And it yields to the chart title** instead of being drawn through it. On a
  chart too short to carry both, the unit used to overprint the title and you
  lost the reading of both. It now steps aside on exactly the charts where that
  would happen — measured, that is a 300x60 banner at 18pt and nothing anyone
  presents at; every chart 480x300 and larger keeps its unit at every font size.
  This is the same rule the category names, the axis strip and the legend have
  always followed: on a chart that cannot hold everything, the title stays.

What that looked like on a real chart:

- **A pie or doughnut with many slices** drew its outside labels through each
  other. They now shrink to the room the neighbour leaves and are dropped when
  even that is too small — the wedge is still drawn, so a dropped label loses a
  name, not a number.
- **A radar** did the same with its spoke names. On a web this matters more than
  anywhere else: the chart *is* the mapping from name to axis, so a name nudged
  onto its neighbour's spoke does not look untidy, it lies. Nothing is moved.
- **A combo with a label on every point** ran the numbers together at twenty-odd
  categories, and a combo with several line series stacked their labels on one
  category. Both now yield to their neighbours.
- **A scatter or bubble** printed point labels across the axis numbers. The
  engine had already decided that a point's label is data and an axis number is
  chrome; it just never acted on it. The chrome yields now, and the placer tries
  to dodge a number before overwriting it, so the axis keeps most of its scale.
- **A dual-axis or pareto chart's second axis** had no fit at all: five tick
  numbers, each placed at its own tick and never measured against the next. On a
  short plot they were simply drawn on top of one another.
- **A gantt's date row** thinned by a fixed gap that could not know how wide a
  date is, so "December 2024" ran into the next one.
- **A gauge** crowded its slice labels and printed them through the big total in
  the middle — and drew that total wider than the arc it sits in.
- **A bubble chart's size key** centred each reference number in a box the width
  of its own circle, so the smallest number spilled out over its neighbours.

### Small multiples no longer scatter their labels off the slide

Asking for a grid of panels that could not fit produced a chart whose text ran
hundreds of points below the bottom of it — on the slide, under whatever came
next, and nowhere near the chart it belonged to. Ten series in two columns on a
short chart is the case: five rows, and after the title and the gaps there was
nothing left per panel.

The engine was computing a negative panel height and handing it on, where a
guard meant for malformed input from outside — a width of `NaN` pasted from
somewhere — quietly replaced it with a default. Each panel was then laid out as
a full-size chart, and ten full-size charts were stacked nine points apart
inside a box sixty points tall.

A grid whose panels have no room is no longer drawn: the chart renders whole
instead. Nothing that fits today changes, because only a size of zero or less
was ever being rewritten.

### Editing a chart is two to four times faster

An edit that changes one thing used to send the host everything about every
shape it touched — twenty separate instructions to change a single word of a
title. It now sends only what actually differs.

Measured on the real thing, not in a simulator: retitling sends two instructions
where it sent twenty, recolouring a series 44 where it sent 152, and a deck-wide
rescale 180 where it sent 272. On PowerPoint on the web that is a rescale across
eight charts falling from around 150 seconds to 103, and the smaller edits
falling further in proportion.

### Inserting a chart no longer pays a four-second pause first

Every insert asked PowerPoint for the slide's size, and on a document that had
been sitting a moment that question went unanswered for a flat four seconds
before the pane gave up and got the answer another way. The wait was not buying
anything: when the question is answered at all it is answered in about a quarter
of a second, and the fallback costs about the same again.

The pane now waits a second and a half rather than four, so a cold document
costs roughly two and a half seconds less per insert. Nothing else changed about
how the size is found.

### The pane says when an insert will be slow

Adding a chart to a slide that already holds content costs several times what
the same chart costs on an empty one, and the pane used to say nothing — a
loaded slide simply looked like it had hung. It now estimates the wait from
2,917 timed inserts, and where a slide of its own would at least halve it,
offers one. Both real waits are quoted; neither is called instant.

### It tells you when a release lands mid-session

If a new version is published while your pane is open, the pane is holding an
older page and asks the server for files that have been replaced. That used to
surface as a browser error naming a URL. It now says SSF Charts has been
updated, that closing and reopening the pane fixes it, and that your slides are
untouched.

### A paste from a Danish, German or Nordic Excel is read correctly

Continental Excel writes `1.234` for a thousand two hundred and thirty-four, and
`987,5` for nine hundred and eighty-seven and a half. Pasting that gave you
`1.234` read as one-and-a-bit, `2.500` read as two-and-a-half, and `987,5`
refused outright — two cells silently wrong by a factor of a thousand and one
visibly empty, out of a single ordinary paste.

The convention is now inferred once per paste and written **into the sheet**, so
you watch `1.234` become `1234` in the cell and can correct it if the guess is
ever wrong. It fires only when nothing in the block contradicts it: a US-style
`1,234` anywhere, or a decimal that cannot be a thousands group, and the paste is
left exactly as it came. The pane says how many cells it rewrote.

### A chart this PowerPoint cannot draw arrives as a picture, not as a gap

Below PowerPoint API 1.10 — most often desktop Windows and Mac — a wedge cannot
be built, so pies, doughnuts and sunbursts inserted with no slices at all: a
legend and a title around an empty space. Such a chart now arrives as a complete
picture, with a message naming what would otherwise have been missing. Where no
picture is possible it still draws and still says what is absent, and Explode
refuses rather than handing back the version that cannot be drawn.

### An empty cell is no longer reported as a zero

A blank in a funnel, waffle or cascade was drawn and labelled `0`, and cascade
went further and computed a 100% drop from it — a slide asserting something
untrue about your business from a cell you had simply not filled in. A blank is
now left out, the way it always was in clustered and line charts.

### Stop means stop

Pressing **Stop** during a slow insert and then touching any other control used
to clear the stop, letting the cancelled insert finish and write its chart onto
the slide. The stop now survives whatever you click next.

### Labels in Chinese, Japanese and Korean are measured at their real width

Every CJK character was measured at roughly half its true width, so labels in
those scripts were laid out as though they needed half the room they do — which
is how text ends up drawn over other text.

## 0.4.0 — 2026-08-27

### Renamed to SSF Charts — **this breaks existing installs**

The add-in is now **SSF Charts**, hosted at
`ssf-chart.struktureretsundfornuft.dk`.

**Everyone must re-install it.** An Office add-in manifest pins the pane to a
host, GitHub Pages serves one custom domain per repository, and the old address
no longer answers. An add-in sideloaded from a previous release will open a pane
that cannot load. There is no way to migrate this from our side — the manifest
lives in your PowerPoint, not in ours.

**What to do:** remove the old add-in, then sideload the new
[`manifest-prod.xml`][manifest]. Nothing in your decks changes: every chart you
have already inserted stays editable, because the tags and shape names written
into your slides were deliberately left on their old values. See "Not renamed"
below.

Also renamed: the repository (`dannbleeker/SSF-Charts`, old links redirect), the
npm package, and the Claude Agent Skill — which uploads as a **new** skill
rather than replacing the old one, so delete the `powerchart-charts` entry after
installing `ssf-charts.zip`.

### New visual identity

The pane wears the SSF design system: navy header and headings, blue for chrome,
and a single orange accent — the tick above the content, which is the system's
signature. Ribbon icons are redrawn to match. Both light and dark themes were
checked for contrast; every foreground/background pair clears WCAG AA and most
clear AAA.

### Fixes

- **Fewer full redraws when updating a chart in place.** A single refused group
  read used to disable in-place updates for the rest of a run, so every
  remaining chart was redrawn whole — the expensive path. The refusal is now
  scoped to the batch that saw it. Measured across ten rounds: redraw rate
  **21.4% → 14.3%**, and in-place updates after a refusal went from 0 to 1 per
  round.
- **Deck style is stored in the deck**, so a shared deck keeps its branding for
  whoever opens it.
- **Chart output**: fourteen fixes across options, data shapes and decorations
  where a chart could be left wrong; the text-overlap gate now runs against
  every font, and the value axis no longer climbs into the title.
- **Stability**: the per-sync shape budget is measured rather than guessed —
  earlier values crashed PowerPoint on the web.

### Security

`npm audit` reports **0 vulnerabilities**, down from four. `qs` was pinned past
its advisory range. `image-size` has no patched release — every published
version is covered by two HIGH advisories — but nothing imports it, so it is
replaced by a stub that throws if anything ever does. See
[`vendor/image-size-stub/README.md`][stub].

### Not renamed, deliberately

These are written **into your decks** and read back to recognise a chart as
ours. Renaming them would orphan every chart inserted by an earlier build:

    POWERCHART_CONFIG / _PARTS / _ORIGIN / _SCENE / _DEMO_SLOT   shape tags
    PowerChart                                                   group/shape name
    PowerChart:not-complete                                      banner name
    <powerchartStyle>                                            deck style element

The manifest's `<Id>` is also unchanged: a new GUID would make this a different
add-in rather than an update. Two archive format markers
(`powerchart-host-answers`, `powerchart-crash-log`) are read under both
spellings so 257 archived test rounds stay readable.

[releases]: https://github.com/dannbleeker/SSF-Charts/releases
[manifest]: https://github.com/dannbleeker/SSF-Charts/blob/main/manifest-prod.xml
[stub]: https://github.com/dannbleeker/SSF-Charts/blob/main/vendor/image-size-stub/README.md
