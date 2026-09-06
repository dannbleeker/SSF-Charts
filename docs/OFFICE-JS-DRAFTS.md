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

**Finding 3 named proxy age too — and then I named the wrong thing as well.**
On 2026-09-05 I read the answer sheet as saying a COLLECTION RE-READ poisons the
handle, from the probe called `collection-read-poisons-the-creation-handle`
(refused 3/3). Running it on 2026-09-06 refuted that: with the re-read removed
entirely, the tag still throws 5010.

**Findings 1 and 3 are ONE defect, and it is about the SLIDE.** A slide added by
`slides.add()` is unusable from any later `PowerPoint.run` — shapes drawn on it
cannot be tagged through a creation handle OR through a proxy re-fetched by id,
and its collection reads short. The document's own slides are unaffected, and
stay unaffected after five such adds. That is Draft A below, and it now carries
its own controls rather than a probe id.

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

## DRAFT A — a shape drawn on an add-in-introduced slide cannot be tagged

> ## HOLD — DO NOT FILE THIS YET. The product contradicts its own title.
>
> Found 2026-09-06 while measuring the archive, after the draft was written.
>
> **This add-in tags shapes on slides it added itself, constantly, and it
> works.** `addSlideForChart` adds the slide in its own `PowerPoint.run` and
> returns an id. `insertSceneIntoSlide` then opens a NEW run, draws onto that
> id, and the tag is written by `rescueGroupAndTag` — or, for a chart that
> could not be tagged inline, by `settleAndTagChart` / `settleByCollectionRead`
> in a later run still. Every one of those is `tags.add(CHART_TAG, …)` on a
> shape sitting on a slide the add-in introduced in an EARLIER run. That is
> Draft A's failing configuration exactly.
>
> Be exact about what the archive does and does not say here. It says 8,221 of
> ~9,600 draw batches went to a slide added that round — that is where the
> DRAWING happens, not a count of successful tag writes, and quoting it as one
> would be the same overreach this note exists to catch. What carries the claim
> is that those scenarios PASS: `a big chart on a slide of its own` runs this
> exact sequence and reads 34 ok / 1 failed on post-fix builds. A defect as
> worded below predicts none of them can.
>
> (The first version of this note cited `insert on top of an earlier run —
> 4 of 4 charts re-editable`, and that citation was wrong: `insertTwice` goes
> through `insertSlidesFromPptx`, so its tags arrive pre-baked in the file and
> are only READ back. It proves nothing about writing one. The claim survives
> the correction; the evidence for it changed.)
>
> So one of two things is true, and filing before knowing which is how a report
> gets closed as not reproducible:
>
> **1. The repro's slide is UNSETTLED, not merely earlier.** This repo already
> established the rule, over 269 rounds: *a new slide's id is not durable until
> the slide settles, and is durable afterwards.* The probe answered `threw` 227
> times and `yes` 41, and every `threw` predates our own `77f9ca4` — which
> stopped holding the id `slides.add()` hands back and re-read it positionally
> after the add had settled. `addSlides` does not return until a FRESH context
> has confirmed the deck grew. **Run 1 of the repro below never does that.** It
> reads `slides.items.length - 1` from the same context that added the slide.
>
> **2. The variable is the SESSION, not the run.** Against (1): Presentation72
> slide 5 was inserted from a file in a previous session, was settled by an
> entire browser restart, and still threw. Against the title: within one
> session the product does this thousands of times without a failure.
>
> Both are explained by a narrower defect than the one written below — *a slide
> the add-in introduced in an EARLIER SESSION cannot take a new tag* — and the
> round archive is structurally blind to it, because every round is one session
> on a deck swept immediately afterwards.
>
> **THE EXPERIMENT THAT DECIDES IT**, on a deck reopened from a previous day
> (Presentation72 has both kinds of slide). Draw a shape and tag it on:
>
>     a  the document's own slide
>     b  a slide the add-in added in THIS session
>     c  a slide the add-in added in a PREVIOUS session
>
> If b passes and c fails, the title below is wrong and the real finding is
> both narrower and more interesting — and it is a daily-use bug, because it is
> exactly "open yesterday's deck and add a second chart".
>
> Until that is run, everything below is measured but its SCOPE is unverified,
> and the scope is what a tracker issue is judged on.

**REWRITTEN 2026-09-06 after the first version failed to reproduce.** The
original blamed a collection re-read for poisoning the context. Run against a
live host it did not reproduce at all on an ordinary slide, and a control
showed the re-read is not the trigger. What follows is what actually happens,
measured 3 of 3 on a fresh two-slide deck with a control on each claim.

**Title:** A shape drawn onto an add-in-introduced slide cannot be tagged from
a later `PowerPoint.run` — `InvalidParam passed to GetItem(id)`, 5010 — however
that slide arrived

> ### Your Environment
> - Platform: Office on the web
> - Host: PowerPoint
> - Browser: Chrome
>
> ### Expected behavior
>
> A slide created with `slides.add()` should behave like any other slide once
> it has been created and synced. Shapes added to it should be taggable, and
> `slide.shapes.load("items/id")` should list them.
>
> ### Current behavior
>
> It does not. On a slide added by the add-in in an EARLIER `PowerPoint.run`,
> two things fail, and they fail independently of each other:
>
> 1. `slide.shapes.load("items/id")` comes back SHORT — 2 of 3 shapes just
>    drawn and synced. No error is raised.
> 2. `shape.tags.add(...)` throws `RichApi.Error: InvalidParam passed to
>    GetItem(id)`, code `5010`.
>
> The same code on a slide that was already in the document works.
>
> ### What I checked, so you do not have to
>
> | slide the shapes are drawn on | tag |
> | --- | --- | 
> | already in the document when the add-in started | **OK** |
> | added by `slides.add()` in an EARLIER run | **5010** |
> | inserted by `insertSlidesFromBase64()` | **5010** |
> | added by `slides.add()` in the SAME run | the `shapes.add*()` call itself throws `GeneralException` |
>
> The last two rows are the point: it does not matter HOW the slide arrived.
> A slide inserted from a generated .pptx fails exactly like one from
> `slides.add()`. Only slides the document already had are unaffected.
>
> The clearest single reading is two runs on ONE deck, seconds apart, same code:
>
> ```
> slide 0  (the document's own first slide)  ->  tag OK
> slide 5  (inserted from a generated file)  ->  tag 5010
> ```
>
> Three further controls:
>
> - **It is not the re-read.** Removing the `load`/`sync` entirely and tagging
>   straight after the draw still throws 5010. The re-read is a second symptom,
>   not the cause.
> - **It is not the handle.** Re-fetching the shape by id with
>   `slide.shapes.getItem(id)` and tagging THAT also throws 5010. So it is not
>   a stale creation proxy — it is the slide.
> - **It is not the session, and it is not the document.** On the very same deck
>   and in the same minute, slide 0 tags fine while slide 5 throws. After five
>   slide adds, slide 0 still succeeds.
> - **The tag failure is the reliable symptom.** The short collection read
>   accompanies it sometimes (2 of 3 on one run, complete on another), so treat
>   the 5010 as the bug and the short read as a second, intermittent one.
> - **Existing tags stay readable.** A deck reopened the next day still returned
>   33 tagged shapes across 32 slides with their values intact. This is about
>   WRITING a tag to a newly drawn shape, not about the slide being unreadable.
>
> Reproduced 3 of 3 on a freshly created presentation, and again on a second,
> heavily-used one.
>
> ### Steps to reproduce
>
> ```js
> // Run 1 — add a slide and remember where it landed.
> let index;
> await PowerPoint.run(async (context) => {
>   const slides = context.presentation.slides;
>   slides.add();
>   await context.sync();
>   slides.load("items/id");
>   await context.sync();
>   index = slides.items.length - 1;
> });
>
> // Run 2 — a NEW context. Draw three shapes on that slide and tag one.
> await PowerPoint.run(async (context) => {
>   const slide = context.presentation.slides.getItemAt(index);
>   const rect = PowerPoint.GeometricShapeType.rectangle;
>   for (const left of [10, 60, 110]) {
>     const s = slide.shapes.addGeometricShape(rect);
>     s.left = left; s.top = 10; s.width = 40; s.height = 30;
>   }
>   await context.sync();
>
>   slide.shapes.load("items/id");
>   await context.sync();
>   console.log("listed", slide.shapes.items.length, "of 3");   // 2 of 3
>
>   slide.shapes.items[0].tags.add("MYKEY", "value");
>   await context.sync();                                        // 5010
> });
> ```
>
> Change `getItemAt(index)` to `getItemAt(0)` — the document's own first slide
> — and both failures disappear.
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
> An add-in that creates a slide and then draws on it cannot do so in more than
> one batch, because everything it draws becomes untaggable and only partly
> readable the moment a new `PowerPoint.run` begins. A chart of any size needs
> several batches, so this affects every non-trivial insert onto a new slide.
> Measured at 46 tagging failures in one 38-item run before the shape of the
> bug was understood.

## DRAFT B — a held slide proxy throws only on a freshly added slide

> ## HOLD — DO NOT FILE. This repo may already own this bug.
>
> Found 2026-09-06 by reading `powerpoint.ts` against the draft, not by running
> anything. Draft B is unverified either way; what follows is why verifying it
> matters more than filing it.
>
> **The archive says this behaviour has our fingerprints on it.** The probe
> `shape-add-fresh-getitem-slide` asks almost exactly Draft B's question. Over
> 269 rounds it answered `threw` **227** times and `yes` **41** — and every
> `threw` is round 253 or earlier. The boundary is `77f9ca4`, **our** commit,
> which stopped the probe holding the id `slides.add()` hands back and made it
> re-read the id positionally once the add had settled. The two are different
> id spaces rather than near-misses: `4123571114#123571113` at add time,
> `256#2587447327` a moment later for the same slide.
>
> The rule that came out of it is in `powerpoint.ts` in these words: *a new
> slide's id is not durable until the slide settles, and is durable
> afterwards.* `addSlides` does not return until a FRESH context has confirmed
> the deck grew.
>
> **ARM 2 BELOW NEVER SETTLES THE SLIDE.** It adds, syncs, loads and takes
> `slides.items[len - 1]` inside the same context. That is the pre-`77f9ca4`
> pattern exactly — and `errorLocation` on the throw is
> `SlideCollection.getItem`, which is what a stale id looks like, not what
> holding a proxy looks like.
>
> Draft B's own controls point the same way: it works via `getItemAt(index)`
> and works inside a single sync. Both of those avoid the add-time id. The one
> arm that fails is the one that uses it.
>
> **THE CONTROL THAT DECIDES IT.** Repeat arm 2, but obtain the slide the way
> `addSlides` does — a fresh `PowerPoint.run` confirming the deck grew, then
> read the id positionally — and hold THAT proxy across a sync before adding.
>
>     works  ->  the draft names proxy age; the variable is the stale id, and
>                the bug was ours. Do not file.
>     throws ->  proxy age survives a settled slide, and the draft stands.
>
> This is the third time a draft in this file has named the wrong variable
> (finding 2 was struck, findings 1 and 3 were renamed). The pattern is not
> carelessness; it is that every arm here changes two things at once.

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

## RESOLVED — Draft A now reproduces, 3 of 3. The history is kept below.

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

**RESOLVED THE SAME DAY.** A fresh two-slide deck (`Presentation73`) gave the
settled host the hunt needed. The missing condition was the slide's PROVENANCE:
it must have been added by `slides.add()` in an EARLIER `PowerPoint.run`. With
that stated, the snippet reproduces both failures 3 of 3, and three controls
narrow it further — it is not the re-read, not the handle, and not the session.
Draft A above is rewritten around that and is filable as it stands.

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
