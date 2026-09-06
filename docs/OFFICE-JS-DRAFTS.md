# Drafts for the office-js tracker — NOTHING HERE HAS BEEN FILED

**Owner-gated.** These go out under Dann's GitHub identity, so they are written
here and submitted by him, or not at all. Nothing in this file has been posted.
Prepared 2026-09-05 against the answer sheet `test/fixtures/host-answers-web.json`
(build `8643e2d · 2026-09-01 14:38Z`).

## What changed since BACKLOG item 5 was written

Item 5 listed three findings. Read back against the answer sheet and the round
archive, **one is dead and the other two named the wrong variable.** That is the
whole reason this file exists rather than a copy of the backlog.

**Finding 2 is struck: the selection subsystem does not wedge on this host.**
The backlog says a non-empty `setSelectedShapes([id])` wedges everything after
it, "measured, twice". `selectionLadder` was built to prove exactly that, and it
has since run in **378 archived rounds. 371 of them report "the host answered
all 7 rung(s) — nothing wedged", 2,597 rungs, zero silences**; the seven
non-passes are skips (no probe chart, or a deck scan that could not see the
deck), not wedges. Whatever was measured twice has not recurred in 371
consecutive opportunities. Filing it would report a behaviour this project's own
instrument contradicts.

As a bonus, those 371 rounds are counter-evidence to the "never resolves" half
of #3698 and #4225: the EMPTY call returns in about a second, every time.

**Finding 1 named proxy age. The variable is the SLIDE.**

    shape-add-held-slide-proxy         yes    3/3   scratch=reused-slide
    shape-add-held-slide-proxy-again   threw  3/3   scratch=fresh-slide

Same question, same host, moments apart. Holding a slide proxy across a sync is
fine on a slide that already existed and throws `GeneralException` on one added
this session. The backlog's "it is the holding that fails" is false as worded.

**Finding 3 named proxy age too. The variable is a COLLECTION RE-READ.**

    tag-the-creation-proxy-a-sync-later            yes             3/3
    how-many-syncs-a-creation-handle-survives      survives-8      (healthy)
    how-many-collection-reads-a-context-survives   short-at-1      3/3
    collection-read-poisons-the-creation-handle    refused         3/3

A creation handle survives eight syncs untouched. Re-read the slide's shapes
once and it is refused with `InvalidParam passed to GetItem(id)`. Age is not
what kills it.

## Duplicate search, 2026-09-05

Searched the tracker before writing. Both findings below appear to be original;
the corroboration target has moved.

- No issue in `OfficeDev/office-js` reports a fresh-vs-reused slide split on a
  held slide proxy. The only issue containing `SlideCollection.getItem` is
  #2427, a closed feature request.
- **#6237 is live and is the right home for the second finding.** Open, labelled
  `Type: product bug`, `Area: PowerPoint`, and it carries the identical
  `InvalidParam passed to GetItem(id)` / 5010 / `ShapeCollection.getItem` on a
  tag read. Its reporter attributes it to a **date placeholder on the slide**,
  which is a different trigger from the one measured here — so this is a
  corroborating comment with a second mechanism, not a duplicate.
- #4204 is an open re-file of #2903 with zero comments. Worth knowing #2903 is
  not the dead end the repo's `KNOWN_ISSUES` treats it as: its closure was an
  automated inactivity sweep, not an engineering decision.

## On the Script Lab repro

The issue template asks for one inside an HTML comment; it is requested, not
mechanically required. The line that is load-bearing sits under *Steps to
reproduce*: **"If we cannot reproduce the issue, we cannot triage."** Both
drafts below are therefore written so a stranger can paste the snippet into
Script Lab and watch it fail, without any part of this add-in.

---

## DRAFT A — a collection re-read poisons every handle in the context

**Title:** Re-reading a slide's shapes poisons the context: the re-read returns
short, and the handle that created a shape is then refused with `InvalidParam
passed to GetItem(id)`

> ### Your Environment
>
> - Platform: Office on the web
> - Host: PowerPoint
> - Browser: Chrome
>
> ### Expected behavior
>
> Re-reading `slide.shapes` inside a `PowerPoint.run` should return the slide's
> shapes, and should not affect a `Shape` handle returned by `shapes.add*()`
> earlier in the same context.
>
> ### Current behavior
>
> Two failures, in order, both reproducible:
>
> 1. The **first** re-read of a slide's shape collection in a context comes back
>    short. Reading a slide holding 3 shapes lists 0 of 3. No error is raised —
>    the call resolves with a short collection.
> 2. **After** that re-read, a handle returned by `shapes.addGeometricShape()`
>    earlier in the same context is refused when used, with
>    `RichApi.Error: InvalidParam passed to GetItem(id)`, code `5010`.
>
> The handle is fine until the re-read happens. Carried through **eight**
> consecutive `context.sync()` calls with no collection re-read in between, the
> same handle still accepts a tag write. So the trigger is the collection read,
> not the age of the handle or the number of syncs.
>
> ### Steps to reproduce
>
> ```js
> await PowerPoint.run(async (context) => {
>   const slide = context.presentation.slides.getItemAt(0);
>
>   // 1. Create a shape and KEEP the handle.
>   const created = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
>   created.left = 50; created.top = 50; created.width = 80; created.height = 60;
>   await context.sync();
>
>   // 2. Re-read the slide's shapes. Put at least three shapes on slide 1 first.
>   slide.shapes.load("items/id");
>   await context.sync();
>   console.log("re-read listed", slide.shapes.items.length, "shape(s)");  // 0 of 3
>
>   // 3. Use the handle from step 1. This throws.
>   created.tags.add("MYKEY", "value");
>   await context.sync();                                                  // 5010
> });
> ```
>
> ### Useful logs
>
> ```
> RichApi.Error: InvalidParam passed to GetItem(id)
> code: 5010
> ```
>
> ### Context
>
> An add-in that draws a chart as native shapes has to write a config tag onto
> what it drew. The natural implementation — create, then read back, then tag —
> hits both failures at once: the read-back is short, so the fallback is to tag
> through the creation handle, and that is exactly the handle the read-back has
> just poisoned. Measured at **46 failures in one 38-item run**, leaving shapes
> on the slide carrying no tag.
>
> Possibly related: #6237 reports the same error and code on a tag read, but
> attributes it to a date placeholder on the slide. If those are one bug, the
> placeholder may be a second way to reach the same poisoned state.

---

## DRAFT B — a held slide proxy throws only on a freshly added slide

**Title:** Adding a shape through a slide proxy resolved one sync earlier works
on an existing slide and throws `GeneralException` on a slide added in the same
session

> ### Your Environment
>
> - Platform: Office on the web
> - Host: PowerPoint
> - Browser: Chrome
>
> ### Expected behavior
>
> A `Slide` proxy resolved in an earlier sync should behave the same whether the
> slide existed before the add-in ran or was added by `slides.add()` during it.
>
> ### Current behavior
>
> It does not, and the split is clean:
>
> | slide the proxy points at | `shapes.add*()` through a proxy resolved one sync earlier |
> | --- | --- |
> | already existed | **works**, 3 of 3 attempts |
> | added by `slides.add()` this session | **throws `GeneralException`**, 3 of 3 attempts |
>
> Both arms were run moments apart on the same host in the same session, so this
> is not the host having a bad minute. Two controls narrow it further:
>
> - resolving the slide and adding **inside one sync** works on both;
> - reaching the slide by `slides.getItemAt(index)` instead of by id works on
>   both.
>
> So it is not the id, not the slide, and not `getItem` — it is holding a proxy
> to a *newly added* slide across a sync boundary.
>
> `errorLocation` on the throw is `SlideCollection.getItem`.
>
> ### Steps to reproduce
>
> ```js
> // ARM 1 — an existing slide. Works.
> await PowerPoint.run(async (context) => {
>   const slide = context.presentation.slides.getItemAt(0);
>   slide.load("id");
>   await context.sync();                       // the proxy is now one sync old
>   slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
>   await context.sync();                       // OK
> });
>
> // ARM 2 — a slide added in this session. Throws.
> await PowerPoint.run(async (context) => {
>   context.presentation.slides.add();
>   await context.sync();
>   const slides = context.presentation.slides;
>   slides.load("items/id");
>   await context.sync();
>   const fresh = slides.items[slides.items.length - 1];
>   await context.sync();                       // the proxy is now one sync old
>   fresh.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
>   await context.sync();                       // GeneralException
> });
> ```
>
> ### Context
>
> Any add-in that adds a slide and then draws onto it has to hold that slide
> across at least one sync — the drawing is batched, and a chart of any size does
> not fit in one. The workaround is to re-resolve the slide by index before every
> batch, which costs a round trip per batch.
>
> Possibly related: #2903 describes a freshly added slide not being usable
> immediately on Online, and was closed by an inactivity sweep rather than a
> decision. #4204 is an open re-file of it.

---

## DRAFT C — a comment on #6237, not a new issue

The cheapest of the three, and the one most likely to be read: #6237 is open,
labelled a product bug, and had Microsoft activity on 2026-08-31.

> We hit the identical error — `InvalidParam passed to GetItem(id)`, code
> `5010`, `errorLocation: ShapeCollection.getItem` — on PowerPoint for the web,
> but reached by a different route, so this may be a second trigger for whatever
> is underneath.
>
> No date placeholder is involved. In our case a plain **re-read of a slide's
> shape collection** inside a `PowerPoint.run` is enough:
>
> 1. `shapes.add*()` a shape and keep the returned handle;
> 2. `slide.shapes.load("items/id")` and sync — this comes back **short**, 0 of
>    3 shapes on a slide holding 3, with no error;
> 3. use the handle from step 1 — refused with 5010.
>
> Reproducible 3 of 3, and it is the collection read that matters rather than
> elapsed syncs: the same handle survives eight consecutive syncs when no re-read
> happens in between.
>
> Happy to supply a Script Lab snippet if that would help.

---

## STOP — DRAFT A DID NOT REPRODUCE WHEN I RAN IT, 2026-09-06

The checklist below says to paste each snippet into Script Lab once before
filing, and that a snippet which does not reproduce "is a finding about the
draft, not about the host". I ran Draft A against the live host on
`Presentation72`. It is a finding about the draft.

**Draft A, exactly as written above, on slide 0 of a real deck:**

    before=2 | created | reread=3 | tag=OK-NO-REPRO

The re-read listed **all three** shapes — not short — and the tag through the
creation handle **succeeded**. Both of the draft's two failures are absent.

**The condition is not "any slide", and the draft never said which.** Two more
runs, same session, minutes apart:

| where the shapes are drawn | what happened |
| --- | --- |
| slide 0, an ordinary document slide | no shortfall, tag succeeds — **no repro** |
| a slide added in the SAME `PowerPoint.run` | the shape add itself throws `GeneralException` — that is Draft B's defect, reached before Draft A's question |
| a slide added in an EARLIER run, then reused | the add never resolved — 60s `TIMEOUT`, which is office-js#1650 |

The answer sheet's own sample tags say why: every sample of
`collection-read-poisons-the-creation-handle` is marked `scratch=reused-slide`.
That is a slide **the add-in created earlier in the session and kept**, which is
neither of the two states a stranger would naturally try, and the draft asks
them to use `getItemAt(0)`.

**What this means for filing.** Do not send Draft A as written. Microsoft's
template is explicit — "If we cannot reproduce the issue, we cannot triage" —
and a repro that does not repro is worse than no issue: it spends the one
credibility this project has upstream. Draft C, the comment on #6237, is
unaffected and remains the cheapest thing on the list.

**What would make Draft A filable:** a snippet that reaches the poisoned state
from a cold start, with the slide's provenance stated. The probe does it 3 of 3,
so the state is real and reachable; what is missing is the shortest path to it
that does not depend on this add-in's own scratch-slide bookkeeping. That is an
hour with a live host and a settled deck, and it was not available tonight — by
the end of these three runs the host had stopped resolving slide adds at all.

**Draft B was not run.** Its arm did fire incidentally — the same-run case above
threw `GeneralException` exactly where Draft B predicts — but that was a
side-effect of testing A, on a deck already 51 slides deep, not a clean run of
B's own snippet. Treat B as unverified rather than as supported by this.

## Before submitting — a checklist for the owner

1. **Read the snippets.** They are written from measurements, not run as
   snippets. Paste each into Script Lab once before filing; if one does not
   reproduce, that is a finding about the draft, not about the host.
2. **Attachment.** The backlog suggests attaching
   `test/fixtures/host-answers-web.json`. It carries probe ids, host error text,
   and slide ids of the form `257#4103259385`. No name, no email, no file name,
   no URL. Slide ids are meaningless outside that document, but they are from a
   real deck — attach only if you want to.
3. **File at most two.** Draft C is a comment and costs nothing. If only one new
   issue is filed, make it **Draft A**: it is stable 3/3, it has a live
   neighbour, and it is the one that costs this project 46 failures a run.
4. **The recurring cost is the real one.** Only the issue author can clear
   `Needs: author feedback`, and the tracker's own bot closes on silence. A
   filed issue wants a one-line reply every few days or it lapses.
5. **Do not plan around a fix.** The issues this project depends on have sat open
   for one to nine years. File it because a fixed host helps everyone writing a
   PowerPoint add-in.
