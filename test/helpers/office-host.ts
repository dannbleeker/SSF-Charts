/**
 * A fake PowerPoint host: recording doubles for the Office.js proxy-object API.
 *
 * Extracted from `office-render.test.ts`, which owned it for the whole of this
 * project's history and could not share it. Every host bug found against a
 * real PowerPoint has been paid for with a deploy, a run someone had to sit
 * through, and an uploaded .pptx — and each fix then added one more knob here
 * so the same bug could never come back. Those knobs are worth more than one
 * test file: pinned together they describe a host far nastier than any suite
 * had been running against, which is why the next bug class should be findable
 * without leaving CI.
 */
import { vi } from "vitest";
// @ts-expect-error — the .mjs auditor has no types. Deliberately the SAME
// decoder `npm run verify-deck` uses: a fake that parsed the generated deck
// its own way could agree with the renderer while both disagreed with the file.
import { readDeckBytes } from "../../scripts/verify-deck.mjs";

interface DeckRow {
  slot: number | null;
  title: string | null;
  run: string | null;
  configJson: string | null;
  chartShapes: number;
}

const DEMO_SLOT_TAG_KEY = "POWERCHART_DEMO_SLOT";
const CHART_TAG_KEY = "POWERCHART_CONFIG";

/**
 * The host's misbehaviours, in one mutable object.
 *
 * These were module-scope `let`s inside the test file that owned this harness,
 * which is exactly why the harness could not move: an importing module cannot
 * assign to another module's binding. Gathering them here is what lets any
 * suite drive the fake host — and lets a profile turn a set of them on at once
 * (see `WEB_PROFILE`).
 *
 * Every one of them was added AFTER a real PowerPoint taught us the behaviour
 * it models. That is the pattern this file exists to break: the next such
 * lesson should be learnable in CI.
 */
export const faults = {
  /**
   * Report `left`/`top`/`width`/`height` as the bounding box AFTER rotation,
   * rather than before it.
   *
   * The renderer assumes the opposite everywhere: `addSegment` draws a diagonal
   * as a rectangle of the segment's LENGTH placed at its midpoint and then
   * rotated, and `arrowheadBox` offsets its box on the same premise. Nobody has
   * ever checked which this host means — `Shape.rotation` has never once been
   * written on a real PowerPoint in 333 rounds, because the battery draws only
   * `clustered`, whose single line node is the horizontal baseline.
   *
   * This models the answer we do NOT expect, so the scenario that asks the
   * question can be shown to fail when the answer is bad.
   */
  reportRotatedBounds: false,
  /**
   * A preset name this host's `GeometricShapeType` does not carry — reported
   * the way a real host reports it, as `undefined`.
   *
   * The stub's enum is a throwing Proxy so that a preset it has not been told
   * about cannot come back undefined and get drawn as a shape with no geometry.
   * That is right for catching OUR mistakes and wrong for modelling the HOST's:
   * a real Office.js enum simply lacks the key, `addGeometricShape(undefined,
   * …)` is accepted, and what lands on the slide is invisible. This fault is
   * the only way to drive that branch.
   */
  presetMissing: "" as string,
  /**
   * Accept a write to `rotation` and leave the shape exactly where it was.
   *
   * A host below 1.10 has no `rotation` at all — a different state, already
   * covered by `requirementSets`. This models the nastier one: the property is
   * there, the assignment does not throw, and nothing turns. A shape that never
   * turned reports its original width, which is the SAME reading a well-behaved
   * host gives, so without this the rotation probe could not be shown to tell
   * "this host means the unrotated box" apart from "no rotation happened" — and
   * would confirm the product's own assumption from no evidence at all.
   */
  ignoreRotationWrites: false,
  /**
   * Reject a write to `rotation` outright, the way a read-only property does.
   *
   * The sibling of `ignoreRotationWrites` and a different host: that one accepts
   * the write and does nothing, this one refuses it. They need telling apart
   * because they call for different words in the sheet — `rotation-did-not-take`
   * is a measurement that failed, `rotation-not-writable` is a host that has the
   * property and will not let go of it.
   */
  rotationWriteThrows: false,
  /**
   * Answer no geometry for a shape whose rotation was written in this batch,
   * while an untouched shape in the same batch reads normally.
   *
   * The behaviour rounds 336-338 point at: the rotation probe came back
   * `unreadable` three times while `named-preset-resolves` read a width in the
   * same shape of batch and got 24x24 every time. One flat "unreadable" cannot
   * tell "this host answers no geometry" from "this host answers no geometry
   * AFTER a rotation write", and those are different facts — so the probe now
   * carries an unrotated control, and this is the host that makes the control
   * mean something.
   */
  rotationBlindsTheRead: false,
  /**
   * Accept a preset by name and draw it with no extent — a shape that is there
   * and is not visible.
   *
   * The half of the preset question that `presetMissing` cannot ask. A name
   * present in the enum is not the same as a name this host can DRAW, and the
   * two fail differently: one hands back `undefined` before the call, the other
   * takes the call and puts nothing on the slide. Only the second is invisible
   * from the code, which is why the probe measures the shape rather than
   * trusting the lookup.
   */
  presetDrawsNothing: "" as string,
  /**
   * List the GROUP on the slide and not its members, the way a real slide does.
   *
   * This fake leaves grouped children in the slide's own collection, so a
   * `slide.shapes` read here answers with every segment of every chart. A real
   * PowerPoint answers with one shape called `PowerChart` per chart and nothing
   * else — rounds 334 and 335 show six of them and not one part.
   *
   * It is the reason `where a rotated shape lands` passed in CI and SKIPPED on
   * the host twice: the scenario looked for its segments at slide level, found
   * them here, and found nothing there. Off by default because turning it on
   * globally would rewrite what several hundred tests are looking at; armed by
   * the tests that are about reaching INTO a group.
   */
  groupHidesChildren: false,
  failSyncOn: 0,
  /**
   * Refuse to store a custom XML part — the deck-style write path.
   *
   * A host below PowerPointApi 1.7 is modelled by `supported()`, which is a
   * different thing: there the collection is never reached. This is the host
   * that HAS the API and says no, which is the case `writeDeckStyle` reports as
   * "the deck was not updated" rather than pretending it wrote.
   */
  refuseCustomXmlWrites: false,
  /**
   * Refuse to ANSWER a custom XML read — the deck-style read path.
   *
   * The sibling of the switch above, and it was missing: a host that stores the
   * part fine and will not hand it back. Round 089 is the case, and the defect
   * it exposed was that `readDeckStyle` returned the same `null` for that as for
   * a deck carrying no style at all.
   */
  refuseCustomXmlReads: false,
  /**
   * Refuse the NEXT custom-XML read only, then behave.
   *
   * The shape the real host has: six rounds show the FIRST custom-XML call after
   * a pane loads failing and the second answering, which is what the deck-style
   * read now spends a retry on. A boolean cannot express "once", and without it
   * the retry cannot be tested at all — every attempt would fail and the test
   * could not tell a retry from a read that simply gave up.
   */
  refuseCustomXmlReadsOnce: false,
  /**
   * Make the deck's style namespace hold TWO parts.
   *
   * `writeDeckStyle` deletes before it adds precisely so this cannot arise, and
   * that left the state its design exists to prevent with no way to reach it in
   * a test. The real read answers a null object for it — the same answer as an
   * unbranded deck — so the interesting assertion is that a doubly-branded deck
   * reads as carrying NO style, silently.
   */
  duplicateCustomXmlPart: false,
  swallowAdds: 0,
  /**
   * Report the deck SHORT by this many slides, for the next
   * `slideCountShortReads` resolutions of `getCount`.
   *
   * The fake already reports `getCount` one sync behind, which a real host does
   * and which `committedCount` models. Round 297 was worse, and this pair
   * reproduces it: two inserts of two slides each, the first reporting
   * `landed: 0` for slides that did land, and `slideCount()` still answering
   * **5 when the deck held 7**. A caller sizing a range from that examined two
   * slides of four and drew a confident conclusion from the wrong population.
   *
   * A CEILING, not a subtraction, and it took two wrong shapes to get here —
   * both caught by mutation rather than by the tests passing:
   *
   * - Counting SYNCS armed and expired before the read that mattered, because
   *   the inserts' own syncs ate the lag. The bug's own mutant stayed alive.
   * - Subtracting a constant from every read cancelled in the DELTA the caller
   *   actually uses: `before` came back short too, so `after - before` was
   *   right and nothing reproduced.
   *
   * A ceiling reproduces the round exactly. `before` is read when the deck is
   * below it and is therefore accurate; `after` is read when the deck has grown
   * past it and is capped — 5 when the deck held 7.
   *
   * `slideCountCeilingBinds` is how many reads it may CAP before lifting, and
   * "cap" is the operative word — a read the ceiling does not actually reduce
   * spends nothing. Three shapes were tried before this one, and each was
   * caught by the bug's own mutant staying alive rather than by a test failing:
   *
   * - counting SYNCS armed and expired before the read that mattered, because
   *   the inserts' own syncs ate the lag;
   * - subtracting a constant from every read cancelled in the DELTA the caller
   *   uses, since `before` came back short too;
   * - counting every READ spent the allowance on `before` and on the several
   *   `slideCount()` calls `insertSlidesFromPptx` makes internally;
   * - lifting on a CLOCK expired during the two scenarios `runSelfTest` always
   *   runs ahead of the named one.
   *
   * Counting binds is the one that survives all four, and it is deterministic
   * where a clock is not.
   *
   * The distinction that matters, and why this is not `hollowReads`: nothing
   * here REFUSES. Every read is a real past value and `unread` stays 0. The
   * host is behind, not blind, and only one of those had a guard.
   */
  slideCountCeiling: null as number | null,
  slideCountCeilingBinds: 0,
  faultShapeGetCount: false,
  strictGroup: false,
  strictTags: false,
  /**
   * Refuse this many `tags.add` calls outright, however fresh the proxy.
   *
   * `strictTags` models the stale-proxy refusal, which `settleAndTagChart`
   * exists to survive: it opens a fresh context, resolves the shape by id in
   * the batch that writes, and the fake honours that. A real PowerPoint on the
   * web did not. A round on 2026-08-06 came back with four charts redrawn to a
   * new scale, ungrouped and carrying no config at all — after the settle pass
   * had had its turn — and `same scale across the deck` reported "4 of 8" with
   * no cause attached, because nothing modelled a host on which the second
   * chance also fails.
   *
   * Counted down per call so a test can refuse the drawing context's write and
   * the settle's, and nothing after.
   */
  refuseTagWrites: 0,
  /**
   * Refuse a tag write on a proxy a `load()` has RESOLVED, and only then.
   *
   * `strictTags` models an AGE rule — refuse anything older than one sync — and
   * four real rounds say that is not this host's rule. It answers
   * `tag-the-creation-proxy-a-sync-later: yes`: a shape created, synced, and
   * then tagged through the handle that made it takes the write, however old the
   * handle is. What it refuses is a handle Office.js has RESOLVED, and the host
   * says so in the statement it echoes back:
   *
   *   errorLocation: ShapeCollection.getItem
   *   statement: var shape = shapes.getItem(...) 'originally addTextBox(...)'
   *
   * "originally addTextBox" is the tell. A `load()` makes Office.js rewrite the
   * creation proxy's object path into `shapes.getItem(id)`, and this host will
   * not resolve that for a shape it has just made — so the write fails for a
   * reason that has nothing to do with when it was queued.
   *
   * The distinction is the whole fix: `finishCharts` may keep re-read handles
   * for GROUPING, where they are needed, and must not tag through them.
   */
  refuseTagWritesOnResolvedProxy: false,
  /**
   * Take a write to `shape.top` and keep the old value, while `shape.left`
   * lands normally.
   *
   * office-js#6183, "PowerPoint Online: OfficeJS API does not update left and
   * top properties of selected shape simultaneously" — on the web host, setting
   * both in one sync moves the shape horizontally only. That is the host every
   * round in this archive runs against.
   *
   * Modelled because the scenario that should catch it, `an update follows a
   * moved chart`, asked about x alone until 2026-09-17: it moved a chart 60pt
   * across and 40pt down, then checked one of the two. A fault nothing fails
   * under is a fault nothing is watching for.
   */
  ignoresTopWrites: false,
  /**
   * Refuse this many `shape.load("id,left,top")` calls outright.
   *
   * The one load that asks WHERE a chart landed. See the shape's `load` for
   * the host transcript this models and why nothing else in the fake could
   * reach the state it produces.
   */
  refuseIdLeftTopLoads: 0,
  /**
   * Refuse this many `shape.load("id")` calls outright, however fresh.
   *
   * `strictIdLoads` models the age rule; this one refuses regardless, which is
   * what the probe's own setup meets. `scratchShapes` adds two shapes and reads
   * their ids back on the very next sync — young enough that the age rule lets
   * it through, and exactly the read PowerPoint on the web refuses
   * (`shape-proxy-survives-one-sync: unreadable`).
   *
   * Without it the fake cannot reach the state that made three probe questions
   * unanswerable for weeks: they died in their own setup and reported
   * `no-scratch-slide`, a statement about the probe rather than the host.
   */
  refuseShapeIdLoads: 0,
  /**
   * `load("id")` is refused on a shape proxy older than one sync.
   *
   * The read-side twin of `strictTags`, and the same age rule, because it is the
   * same host behaviour: Office.js rewrites a resolved proxy's object path to
   * `shapes.getItem(id)`, and this host refuses that call whoever makes it. The
   * tag writer had a fault for it; the id reader did not.
   *
   * What it reaches that nothing else could: `ungroupedFallback` reads ids off
   * `it.created`, the proxies `addGeometricShape` handed back, which by then span
   * several batches. `reading back an ungrouped chart's shape ids` failed three
   * times that way in the 2026-08-07 run, and each failure cost a chart its
   * CHART_PARTS_TAG — on a host that ungroups every chart, so on the web that is
   * every chart, and a chart with no parts list grows by a whole chart per edit.
   *
   * Members of a collection read are exempt, and that is the point rather than a
   * convenience: `freshHandle` re-stamps their age, which is exactly the
   * distinction the host draws and the one the fix turns on.
   */
  strictIdLoads: false,
  /**
   * Any sync that resolved a shape BY ID fails; a collection read still works.
   *
   * The divergence the 2026-08-07 round is made of. Sixty-six errors in one run
   * log, every one of them `InvalidParam passed to GetItem(id)`, code 5010, at
   * `errorLocation: ShapeCollection.getItem` — the slide answered, the shape did
   * not. It refused the drawing context's tag write, the readback, the ungrouped
   * chart's id read, AND `settleAndTagChart`'s own fresh-context write, which is
   * what "the update reported 5×no-config" was.
   *
   * The fake had no way to be this host. `strictTags` refuses a tag write on a
   * proxy that is too old and `refuseTagWrites` refuses one regardless — both
   * are about the WRITE. This is about the LOOKUP, and it is the difference
   * that matters, because the recovery it leaves open is a collection read:
   * `shapes.load("items/id")` answers on this host, which is how the repair
   * pass landed 23 retags in the same run that lost 46 tag writes.
   *
   * Thrown from the sync, never from the call that queued it, because that is
   * where Office.js raises it.
   */
  refuseShapeById: false,
  /** How many `getCount()` calls answer as though a just-committed group is not there — see the getCount body. */
  /** Extra slides ONE `add()` lands beyond the one asked for — see the add body. */
  addLandsExtra: 0,
  shapeCountLag: 0,
  /** How far off those stale counts are. A group of N inner shapes reads N-1 too high. */
  shapeCountLagBy: 0,
  hollowReads: 0,
  /** Answer HONESTLY for this many `items/id` reads, then short forever. */
  hollowReadsAfter: null as number | null,
  /**
   * Leave this many shapes OUT of every `items/id` read, rather than all of them.
   *
   * `hollowReads`' own comment describes a host that "asked about 19 shapes and
   * got 3 back", and then returns `[]` — so for two years the fake modelled the
   * limit case and never the one it documents. The difference is not academic:
   * an empty read makes `chooseGroupMembers` say "group nothing", while a SHORT
   * one is deliberately kept and grouped (see the partial-match comment in
   * `powerpoint.ts`), so the two faults drive opposite branches.
   *
   * The real host produced the short case on 2026-08-11: one slide came back
   * carrying a `PowerChart` group plus four loose shapes — `label-1-3`,
   * `baseline`, `series-label-0`, `series-label-1` — all inside the chart's own
   * box and all with lower ids than the group. That is a partial group, and a
   * partial group is a chart the user drags away from its own baseline.
   */
  readsMissing: 0,
  /**
   * Leave this many shapes out of the NEXT `items/id` read only, then answer
   * honestly — a slide that is still settling rather than a host that is short
   * for good.
   *
   * `readsMissing` above is permanent and `hollowReads` heals but only from
   * EMPTY. Neither can express the case the office-js tracker describes and the
   * round archive shows: PowerPoint Online does not populate a freshly
   * materialised slide's shape collection straight away, so the first answer is
   * partial and a later one is complete. That is the case `REREAD_RETRY_MS`
   * exists for, and without this it could not be reproduced — a suite that can
   * only model permanent shortness would show the retry costing a delay and
   * never show it saving anything.
   *
   * The VALUE is how many shapes to drop, and it applies to the next
   * `items/id` read only — so `4` means "the first answer is four short, every
   * one after it is right".
   */
  readsMissingFirst: 0,
  /**
   * Answer this many `items/id` reads with the slide's shapes listed under ids
   * that match nothing, then answer honestly.
   *
   * The third way the pre-grouping re-read fails and the one no fault could
   * express: not empty (`hollowReads`), not short (`readsMissing`), but a full
   * list the join cannot use. Rounds 066 and 067 showed five of five
   * single-batch charts declining this way, and the only visible sign was the
   * ABSENCE of a settle-delay line beside the decline.
   *
   * A single-batch chart loads its shape ids in the very sync that creates them
   * and re-reads on the next one, so this models a host that has listed the
   * slide before those ids settled.
   *
   * A COUNT, like `hollowReads`: `1` is a slide that settles by the second look,
   * a large number is a host that never joins up. Both are worth testing and
   * they want opposite outcomes.
   */
  unmatchedIdReads: 0,
  /**
   * The id a SELECTION reports for its slide, when that is not the deck's id.
   *
   * office-js#2474: a `SlideRange`'s id lacks the `#XYZ` suffix the same slide
   * carries in `presentation.slides`, so an id taken from the selection cannot
   * be handed to `getItem`. The fake returned the very same object for both
   * reads, so the divergence the issue is about could not exist here — and the
   * add-in stores exactly that id as an edit target's `slideId`.
   */
  selectedSlideIdAs: null as string | null,
  /**
   * A slide whose shape collection AND its count both answer zero once a
   * picture has landed, while the slide really holds shapes.
   *
   * `slideShapeList` corroborates `items` against `getCount()` and answers null
   * when they disagree — which is why `hollowReads` and `hollowNameReads` are
   * both caught. This is the case that defeats it: the two signals agree, at
   * zero, and are both wrong.
   *
   * Observed on 2026-08-12 (`957aca0`). `explode a degraded picture` read the
   * slide it had just drawn on, was told it held nothing, and reported `the
   * slide went from 1 to 0` — while the deck inventory taken at the end of the
   * same run shows that slide holding one shape named `PowerChart`. A verdict
   * claiming data loss the add-in did not cause, which is the third time this
   * scenario has produced one by counting a slide.
   *
   * Two arming times, because WHEN it starts decides whether a test proves
   * anything. `"now"` is for a unit test that calls the read directly.
   * `"after-a-picture"` is for anything that goes through the battery: a fault
   * that is on from the first read starves `probeCharts`, the scenario skips
   * for want of a chart, and the guard then passes against the unfixed build —
   * which is the shape of a test that proves nothing, and was the first version
   * of one here. `fill.setImage` is the turn because that is where the real
   * host turned: the reads before the collapse answered, the read after it did
   * not.
   */
  slideReadsEmpty: null as null | "now" | "after-a-picture",
  /**
   * A slide whose `id` is not readable until this run's first sync has
   * answered — which is what a real host does, and what the fake does not.
   *
   * `slide.load("id")` populates nothing until a `context.sync()` runs, so code
   * that resolves a slide and reads its id in the same breath gets
   * `PropertyNotLoaded` on the web and a plain string here. The fake's slide
   * object doubles as the fixture tests assert against (`deck.map(s => s.id)`),
   * so this cannot be the default for the same reason `strictShapeReads`
   * cannot: it would fail hundreds of tests on their own reads rather than on
   * the code under test.
   *
   * Keyed on the run's first sync rather than on per-property load tracking.
   * That is coarser than the real rule and enough for the thing it exists to
   * prove: the window where a draw has begun and the host has not yet named its
   * slide, which is where `renderShapesChunked`'s first batch lives and where
   * its shape count would otherwise be stranded on the `(visible)` sentinel.
   */
  slideIdUnreadableBeforeFirstSync: false,
  /**
   * The id NEVER reads, however many syncs the load has had.
   *
   * The sibling above models the ordinary PropertyNotLoaded — a read that came
   * too early and works once the sync lands. This models the host that takes the
   * `load("id")`, takes the sync, and still will not say: the slide exists, and
   * its name is unobtainable. The archive's `idRefusals` counter is this.
   *
   * It matters because the alternative to "no id" is not "wait longer", it is
   * "make one up". `addSlideForChart` ends in `?? null` for exactly this, and
   * without a fault that reaches it, a mutant returning a fabricated id passed
   * every test — an insert addressed to a slide that does not answer to that
   * name, reported as success.
   */
  slideIdNeverReadable: false,
  /**
   * The same short answer, on an `items/name` load.
   *
   * A real host does not care which properties were asked for — a collection
   * that answers short answers short. The split exists only because this fake
   * needs to blind ONE reader without blinding the count it is checked against:
   * the tag pass asks for `items/id`, while the slide-swap gate and the
   * group-child count both ask for `items/name`. Arming the wrong one is how a
   * test comes to exercise nothing.
   */
  hollowNameReads: 0,
  refusePictureFill: false,
  refuseGroups: 0,
  /**
   * Whole decks the host accepts and never lands — the file path's version of
   * `swallowAdds`, and observed for real: a 12-item file insert came back
   * "11 of 12 complete · 1 lost" with nothing thrown. Counted down per call.
   */
  swallowDecks: 0,
  /**
   * `shapes.load(...)` throws a TypeError instead of queueing.
   *
   * Observed on the web as "e.load is not a function": the host handed back
   * something without the method, and because the QUEUE step sat outside the
   * try/catch that was written to cover it, an update whose shapes had already
   * committed reported total failure. Charts on the slide, one TypeError, no
   * result — which is what Same Scale across the deck did.
   */
  faultShapeCollectionLoad: false,
  /**
   * A shape's `id`/`left`/`top` are unreadable until a `load()` on it has been
   * ANSWERED by a sync — which is simply how Office.js works, and how this fake
   * has never worked.
   *
   * Off by default, and only because the fake's shape objects double as the
   * surface tests assert geometry against: turning it on globally would fail
   * hundreds of tests on their own reads rather than on the code under test.
   * Where it IS on, it catches the class of bug that a lenient shape hides
   * completely — reading a position off a proxy whose load went down with a
   * failed sync, which is `PropertyNotLoaded` at `Shape.left` on a real host
   * and a plain number here. That is how a deck-wide rescale came to report a
   * crash for charts that were, on the slide, correctly redrawn.
   *
   * A test may switch it back off before inspecting the fake's own state: the
   * gate models a HOST read, not the bookkeeping a test does afterwards.
   */
  strictShapeReads: false,
  /** The host refusing `Shape.group` — see the accessor. One-shot via `refuseGroupReadTimes`. */
  refuseGroupRead: false,
  /** Refuse the group read N times, then allow it — the case where a retry can pay. */
  refuseGroupReadTimes: 0,
  /** See `settleAddedSlideId`. Null is off; N settles the id after N lookups. */
  newSlideIdSettlesAfter: null as number | null,
  /**
   * A shape resolves by id and then will not say what its id is.
   *
   * office-js #2903 itself, and the state this fake could not produce: every
   * shape it finds reads back, so the two halves of an `unreadable` —
   * "no such shape on that slide" and "found it, would not read it back" —
   * were indistinguishable in any test. `shape-resolve-held-slide-proxy` exists
   * to tell a real host's answer apart from exactly that, so the fake needs to
   * be able to be the host on this point.
   *
   * NOT the same as `strictShapeReads`, which THROWS PropertyNotLoaded. Here
   * the sync succeeds and the proxy is not null; the property is simply never
   * populated, which is what a held slide handle does on PowerPoint web and why
   * the failure is silent rather than caught.
   */
  shapeIdNeverPopulates: false,
  /**
   * A shape cannot be NAMED in the batch that created it.
   *
   * `shape.load("id")` queued in the same sync as `addGeometricShape` is taken
   * and never answered, so the id stays unreadable. Implied by the 2026-08-05
   * answer sheet, where the five questions that read an id back that way were
   * the only five never put — and no other property of those five explains it.
   *
   * Not proven directly: no question asks it yet. Armed by the tests that need
   * a host which behaves this way, and off elsewhere, so nothing in the tree
   * depends on it being true.
   */
  noIdInCreatingSync: false,
  /**
   * How long a sync takes, as a function of what has happened before it.
   *
   * A fake that is always instant cannot be used to check a measurement, and
   * the degradation experiment is nothing but a measurement: its whole job is
   * to tell "the context is what slows down" from "the host is what slows
   * down", and a fake with no clock says neither. Both arguments are needed
   * because they are exactly the two hypotheses —
   *
   * - `syncsInContext` grows within one `PowerPoint.run` and resets at the next.
   *   A cost that follows it is proxy accumulation, and only the long-context
   *   arm feels it.
   * - `syncsTotal` never resets. A cost that follows it is the deck growing or
   *   the tab ageing, and both arms feel it equally.
   *
   * Stands for a real observation without claiming to model its mechanism: a
   * degrading host is the one thing every real-host artefact this project owns
   * has in common, and none of them can say which of the two it was. Off by
   * default — an ordinary test must not pay milliseconds for it.
   */
  syncCostMs: null as null | ((load: { syncsInContext: number; syncsTotal: number }) => number),
  /**
   * A sync that does not come back while a shape is SELECTED.
   *
   * What the real host does to `a selected shape survives an insert`: that
   * scenario's draw stalled its first batch in four of the last five rounds
   * (`957aca0`, `ee1741e`, `89675b6`, `47a80c8`) while every other draw in the
   * same rounds landed — including draws that also follow `selecting a shape`,
   * which is why the preceding call is not the variable.
   *
   * Modelled as a delay rather than a hang so a test cannot wedge: pair it with
   * a short `_setBatchTimeoutForTest` and the draw gives up exactly as it does
   * on the web, while the fake's own promise still settles.
   */
  stallDrawAfterSelect: false,
  /**
   * Tag loads the host takes and never answers — the next N `load()`s on a tag
   * proxy leave it unreadable after their sync.
   *
   * The tag twin of `unansweredShapeReads`, and the one with the widest blast
   * radius: `isNullObject`/`value` on a config tag is read on every path that
   * asks "is this shape a chart" — the selection readers, the deck scan, the
   * repair pass. Each of those used to read both properties raw, so a tag the
   * host stayed quiet about did not mean "not a chart", it meant the call
   * threw. Counted down per load, so a test can silence one read and watch the
   * rest of a scan carry on.
   */
  unansweredTagLoads: 0,
  /**
   * `setSelectedShapes([])` does not clear the selection — PowerPoint on the
   * web, office-js#3083. Desktop clears it.
   *
   * Worth modelling because of what it costs downstream rather than what it
   * costs at the call: on the web a picture cannot be inserted while a shape is
   * selected (office-js#3698), so a battery whose scenario leaves a chart
   * selected fails the NEXT scenario instead of its own. That is the hardest
   * kind of failure to read in a run log, and it is only reachable here if the
   * fake refuses the clear the way the web host does.
   */
  webIgnoresDeselect: false,
  /**
   * `setSelectedShapes(ids)` selects SOMETHING, but not what was asked for.
   *
   * A host that takes the call and gets the shape wrong is worse than one that
   * refuses: the pane then edits a chart the user did not click. The scenario
   * that covers the selection round trip has to be able to see that, and it
   * can only see it if the fake is capable of doing it.
   */
  selectionIgnoresIds: false,
  /**
   * A programmatic `setSelectedShapes(ids)` **wedges the selection subsystem**.
   *
   * Measured on PowerPoint on the web: the call itself is taken — no error, no
   * refusal, the sync resolves — and then every selection call after it goes
   * silent. `getSelectedShapes` ran out a full 90-second budget, and so did the
   * `setSelectedSlides` behind it. Nothing throws. The host simply stops
   * answering, which is the one failure shape a fake that only ever throws or
   * only ever lies cannot produce.
   *
   * The third of the web host's selection bugs, after #3083 and #3698, and the
   * reason the battery's selection scenario reports *skipped* rather than red
   * there. That gate is only worth having if CI can watch it work, and it can
   * only watch it work against a host that behaves this way.
   *
   * Deliberately NOT in `applyWebProfile()`, though it belongs to the same real
   * host: it is the one fault whose cost is measured in seconds rather than in
   * assertions. Every test using the profile would pay two full budgets to
   * reach a verdict it was not asking about.
   */
  selectionWedgesHost: false,
  /**
   * `Slide.getImageAsBase64` answers with the same bytes whatever is on the
   * slide — a host that hands back a blank render, or a cached one.
   *
   * The visibility scenario exists to catch a chart that is structurally
   * perfect and invisible. Against a host like this its comparison is
   * meaningless, and it must report that rather than pass.
   */
  constantSlideImage: false,
  /** The host takes the rasterise and hands back an empty image. */
  emptySlideImage: false,
  /**
   * `Slide.getImageAsBase64` does not answer until this promise settles.
   *
   * Rasterising a slide is the heaviest single call the add-in makes, and it is
   * the one a real host wedged on — 1819 seconds, ended by closing the tab. A
   * fake that can only make it lie or refuse cannot show what a log says while
   * a call is still outstanding, and that is the only thing a run which never
   * ends leaves behind.
   */
  slideImageGate: null as null | Promise<void>,
  /**
   * `Slide.delete()` is taken and does nothing.
   *
   * The self-test's visibility scenario is the only one that borrows a slide
   * and gives it back, so it is the only one that can leave litter. A host that
   * refuses the delete without saying so leaves a tagged chart in the deck under
   * a verdict that reads clean — reporting success while leaving a mess, which
   * is the pattern this project keeps finding. The scenario has to notice.
   */
  refuseSlideDelete: false,
  /**
   * The deck has TWO slide masters, and its slides belong to the second.
   *
   * The shape of every deck that has ever killed this host. `Presentation70`
   * reads `2 master(s), 12 layout(s)` and crashed 31 of 34 archived rounds;
   * `Presentation64` reads `1 master(s), 1 layout(s)` and crashed 9 of 182.
   * This fake had one master, so the whole class was unreachable and the
   * add-in shipped a documented API misuse for months.
   */
  twoMasterDeck: false,
  /**
   * The slide IS GONE and the delete still reports failure — round 370's 4:3
   * host, reproduced.
   *
   * There the add reports `landed=1`, the listing puts the new slide at index 7
   * of 8, and moments later the deck is back to seven. `deleteSlideById` is
   * then holding an id the deck no longer lists, correctly refuses to remove an
   * index whose id does not match, and returns false.
   *
   * Nothing in this fake could produce that pairing. It offered a refusal that
   * LEAVES the slide, or a success that REMOVES it, but never "refused, and it
   * was never there" — so the scenarios read `false` as "a slide is stuck in
   * the deck", announced wreckage that did not exist, and no test could catch
   * it, because the double could not fail the way the host fails.
   */
  slideVanishesInsteadOfDeleting: false,
  /**
   * The next deck insert LANDS, and its `context.sync()` never answers.
   *
   * office-js#1650, verbatim: "the first time `context.sync()` is called the
   * promise resolves, but in subsequent calls the promise doesn't resolve,
   * although **the slide still gets added successfully**." Marked fixed
   * upstream; the SHAPE of it is the point, because it is the proof that on
   * this platform silence says nothing about whether work happened.
   *
   * A fake that could only ever throw or lie cannot produce it, and without it
   * "verify from the document instead of the promise" is a claim no test can
   * check. One-shot, so a retry can still be observed succeeding.
   */
  deckInsertNeverAnswers: false,
  /**
   * `getSelectedShapes().load(...)` throws instead of queueing.
   *
   * The fake could make a SLIDE's shape collection fail
   * (`faultShapeCollectionLoad`) but never the SELECTION's — which is the
   * pane's most-used read, the one behind "Edit it", and therefore the one
   * whose failure a user actually meets. Observed on the web as "e.load is not
   * a function": the host hands back something without the method.
   *
   * Modelled because a click the host will not describe and a click on
   * something that is not a chart are different findings, and a fake that can
   * only produce the second cannot show that they are told apart.
   */
  selectionReadThrows: false,
  /**
   * The next N shapes to be tagged answer `.tags` as **undefined**.
   *
   * Observed on a real host, four times in one run, each on a chart whose
   * grouping had just been refused with InvalidParam 5010. Reading `.add` off
   * it throws SYNCHRONOUSLY, which is what made it so expensive: it escaped
   * the tagging loop rather than failing one chart, so every chart after it in
   * the batch lost its config without being attempted. Counted down per read.
   */
  tagsUndefinedOn: 0,
  /**
   * A slide this deck ADDED resolves by id this many times, then reports gone.
   * `null` (the default) is every host anyone has met: ids stay good.
   *
   * Straight from a real PowerPoint on the web. The host-probe run added one
   * scratch slide, resolved its id for the first question, and refused it for
   * the other thirteen — which arrived as thirteen identical "the host would
   * not resolve the scratch slide" answers where thirteen host behaviours
   * should have been. Counted per slide, because the interesting question is
   * what the run does NEXT: a replacement slide has to get its own fresh
   * lease, or "make another one" is not a recovery.
   *
   * `null` rather than `0` for "off" so that `0` can mean what it says — this
   * slide never resolves, starting now. A test that arms the fault AFTER
   * getting an id needs that, and needs it not to depend on how many lookups
   * the code under test happened to spend on the way: a count-based version of
   * that test passed against the very code it was written to falsify, because
   * the fix spends one lookup more than the bug did and the lease outlasted
   * both.
   */
  newSlideResolvesTimes: null as number | null,
  /**
   * Refuse a held slide handle the first N times, then start honouring it.
   *
   * Models the one behaviour a real host has shown that no fake here could:
   * `shape-add-held-slide-proxy` answers `threw` early in a run and `yes`
   * later, in the SAME run, on the same build — seventeen `threw` on pass 1
   * against three `yes` on later passes, across two rounds. Every fake host
   * until now refused a stale handle consistently, so the probe pair built to
   * decide whether that flip is a host state or a coin had nothing to fire
   * against and its trigger condition could not be tested at all.
   *
   * Counted in REFUSALS rather than syncs or seconds, because what the guard
   * needs is "the answer changes partway through a run" and a refusal count is
   * the only clock the fake and the probe agree on.
   */
  heldSlideProxyRelentsAfter: null as number | null,
  /**
   * The first N slides this deck adds never resolve; every one after them does.
   *
   * The sibling above models a host that loses a slide and keeps losing it.
   * This one models what two real rounds actually showed: the ability to resolve
   * a freshly added slide comes and goes, in windows of roughly fifteen seconds,
   * and the SAME RUN recovers on the other side. The 2026-08-09 round answered
   * questions at positions 17, 18, 22, 24 and 26 after losing 10 through 16.
   *
   * Counted in slides rather than seconds because nothing the probe does
   * consults a clock, and a window measured in slides is the one the code can
   * actually be driven through. Without it the fake could express "the host
   * never resolves a scratch slide" and "the host always does", and the interval
   * between them — the state every real sheet has been taken in — not at all.
   *
   * Off by default (`0`). Set it and the probe's end-of-run second pass has
   * something to recover from.
   */
  newSlideRefusedForFirst: 0,
  /**
   * How this host refuses a PowerPointApi 1.8 shape binding, if it does.
   *
   * `"call"` throws from `bindings.add` itself. `"sync"` takes the call and
   * rejects the batch carrying it — which is the shape the 2026-08-09 evening
   * round showed: `UnexpectedError` back from the commit in 1.3 seconds, on a
   * run where `shape-add-fresh-slide-proxy` answered `yes`.
   *
   * Two different facts about the same API, and `binding-names-shape-later` has
   * a distinct answer word for each, so the fake has to be able to produce both
   * or neither word is reachable in a test.
   *
   * `null` (the default) is a host that binds — the happy path the rest of this
   * file models. Nothing has established which of the three a real PowerPoint
   * is; closing that is what the probe is for.
   */
  refuseBindings: null as null | "call" | "sync",
  /**
   * Appending a slide RENUMBERS an existing one, so a before/after id diff shows
   * two new ids for one added slide.
   *
   * Straight off the wire, and common rather than exotic: seven observations
   * across four real rounds, every one the same arithmetic — the deck grew by
   * exactly one and two ids read as new (`before=20 after=21 fresh=2`). One id
   * has to leave the list for that to add up.
   *
   * It is not a curiosity. `addScratchSlide` refused to claim either candidate,
   * which on 2026-08-10 cost the probe's entire second pass: three attempts,
   * three `fresh=2`, five questions never re-asked. The fake could not express
   * it, so nothing in CI could have caught the refusal.
   *
   * The renumbered slide is always one that was ALREADY there and keeps its
   * position; the appended slide is untouched and stays last. That is exactly
   * the property the fix relies on, so a fake that scrambled the new slide too
   * would be asserting something no round has shown.
   */
  renumbersOnAdd: false,
  /**
   * An added slide does NOT land at the end of the deck.
   *
   * The counter-case to `renumbersOnAdd`. `addScratchSlide` may claim the last
   * slide only when that slide is genuinely new, and claiming by position is
   * safe precisely because `add()` appends — so a fake that ALWAYS appends
   * cannot exercise the check that makes it safe, and it could be deleted with
   * every test still green. It was.
   *
   * Not invented for the test either: this deck is documented as scrambling
   * under load — `DemoReport.blankSlides` is reported by POSITION exactly
   * because "the host reorders/merges/loses slides, which breaks any positional
   * item mapping". If that can happen to a demo run's slides it can happen to a
   * scratch one, and the answer there must be to give up rather than claim
   * somebody else's slide and delete it afterwards.
   */
  addsAtFront: false,
  /**
   * A freshly-added slide's `slides.getItem(id)` handle is single-sync too.
   *
   * A FAULT, not the default, and the distinction is the point. Every
   * held-handle failure PowerPoint web has reported names `errorLocation:
   * SlideCollection.getItem`, so making this unconditional is tempting — and it
   * would be the fake asserting something nobody has asked. All seventeen probe
   * questions resolve slides through `getItemOrNullObject`; not one asks what
   * `getItem` does to a new slide, in either direction. Turning it on
   * unconditionally asserts TWO unasked things at once: harsher on holding, and
   * kinder on a fresh use, because a spent handle's `load()` is a no-op here and
   * its `.id` is a plain value, so a fresh `getItem` read is never refused.
   *
   * `shape-add-fresh-getitem-slide` asks the first half. Until a sheet answers
   * it, this stays a knob a test can turn to demonstrate the hypothesis, and
   * nothing in the tree depends on it being true.
   */
  newSlideGetItemExpires: false,
  /**
   * `slides.getItem(id)` REFUSES a slide this run added, on the spot.
   *
   * Distinct from `newSlideGetItemExpires`, which hands back a handle that dies
   * at the next sync: this one never gives a usable handle at all. Every
   * held-handle failure a real PowerPoint has reported names
   * `errorLocation: SlideCollection.getItem`, which reads as though the call
   * itself refuses new ids — but every one of those was a handle resolved a
   * sync earlier, so the plain claim has never been tested either way.
   *
   * Off by default and armed only where a test needs a host that behaves this
   * way, because nothing has established that any host does. That is the whole
   * reason `shape-add-fresh-getitem-slide` and its follow-up exist.
   */
  refuseGetItemOnNewSlide: false,
  /**
   * Adding a text box deletes the selected shape — office-js#2775, web only.
   *
   * See the call site for why it is worth modelling and why it is off by
   * default. `dropShapeSelection` is the guard; this is what proves the guard
   * is load-bearing rather than decorative.
   */
  textBoxDeletesSelection: false,
  /**
   * The host takes a shape add and fails the sync that would land it.
   *
   * The other branch of the fork `expiringSlideHandle` documents: a host that
   * will not put a shape on a scratch slide AT ALL, however the slide is named.
   * The probe cannot tell the two apart from a failure alone, so it must not
   * try — a probe that cannot get its shapes answers `no-scratch-shape`, which
   * no probe can produce as an answer to its own question.
   */
  refuseShapeAdds: false,
  /** See the counted branch beside `refuseShapeAdds`: refuse N adds, then allow. */
  refuseShapeAddsTimes: 0,
  /**
   * After this many syncs, EVERY sync stops settling — neither resolving nor
   * rejecting, forever.
   *
   * The one failure shape no `catch` can see, and the one this host actually
   * produces: office-js#3698 and the wedge measured on this project's own build
   * are both "the promise simply never comes back". A fault that throws models a
   * host saying no; this models a host saying nothing, which is what a deadline
   * exists for. `null` (the default) never wedges.
   *
   * Broader than `selectionWedgesHost`, which arms only after a programmatic
   * `setSelectedShapes` — production never makes that call, so it could not be
   * used to test the paths production actually takes.
   */
  wedgeAfterSyncs: null as number | null,
};

/**
 * PowerPoint on the web, at its worst — every observed misbehaviour at once.
 *
 * Not the default: applied to every existing test it would fail hundreds of
 * them for the wrong reason, asserting nothing about the code under test. It
 * is a named profile, and the demo path is pinned to it (`web-host.test.ts`)
 * because the demo path is where all of this was found.
 *
 * `swallowAdds` and `failSyncOn` are deliberately NOT here: they are one-shot
 * counters aimed at a particular sync, and a blanket value would decide which
 * call fails rather than letting the test say.
 *
 * Call AFTER `installHost` — installing a host resets every fault to its
 * well-behaved default, so applying the profile first would be undone.
 */
export function applyWebProfile(): void {
  // A shape proxy older than one sync is rejected — the web host rewrites
  // proxies as `shapes.getItem(id)` and answers InvalidParam for a stale one.
  faults.strictGroup = true;
  faults.strictTags = true;
  // Shape collections read back shorter than they are, without throwing.
  faults.hollowReads = 3;
  // The first grouping attempt is refused outright.
  faults.refuseGroups = 1;
  // setSelectedShapes([]) is ignored — office-js#3083.
  faults.webIgnoresDeselect = true;
  // NOT refusePictureFill. The web host supports PowerPointApi 1.8 and takes
  // pictures — that support is the whole reason the degraded-picture path
  // exists there. Refusing it models a DIFFERENT host, and it silently
  // switches off the one path where the worst of the stale-proxy bugs lived:
  // a degraded chart is a single shape, and a single shape is what the tag
  // write had no fresh proxy for.
}

/**
 * Recording doubles for the PowerPoint JS proxy-object API: every shape the
 * renderer creates is captured with the geometry/format calls made on it, so
 * the whole scene→native-shapes mapping is testable without an Office host.
 */

let idSeq = 0;

/**
 * The slide size this fake deck declares, in EMU. Defaults to PowerPoint's
 * canonical 16:9; a test sets it to 9144000×6858000 for 4:3.
 *
 * Drives BOTH reads the renderer can make — `pageSetup` and the `<p:sldSz>` in
 * an exported slide — so a test cannot accidentally assert against a size only
 * one of the two rungs would report.
 */
export const hostSlideSize = { cx: 12192000, cy: 6858000 };

/** Exports whose base64 is built during the sync that was asked for it. */
const pendingExports: { result: { value: string }; build: () => Promise<string> }[] = [];

/**
 * The host's current SHAPE selection, shared by everything that touches it.
 *
 * One array, mutated in place, because `Slide.setSelectedShapes` (the writer)
 * and `Presentation.getSelectedShapes` (the reader) are on different classes
 * and a test needs the round trip between them to be real. `installHost` seeds
 * it; the slide method writes it; the presentation getter reads it.
 */
const selectionRef: FakeShape[] = [];

/**
 * Set once a programmatic select has wedged the host — see
 * `faults.selectionWedgesHost`. Sticky for the rest of the run, because on the
 * real host it was: nothing un-wedged it short of a reload.
 */
let selectionWedged = false;
/** A selection call joined the sync now being built, and it will never land. */
let wedgeThisSync = false;
/** A by-id shape lookup joined the sync now being built — see faults.refuseShapeById. */
let refuseThisSync = false;

/** Called by every selection entry point once the subsystem is wedged. */
function noteSelectionCall(): void {
  if (selectionWedged) wedgeThisSync = true;
}

/**
 * A minimal .pptx carrying nothing but a `ppt/presentation.xml` with the
 * deck's `<p:sldSz>` — which is the only part `slideSize` reads.
 */
async function exportedDeckBase64(): Promise<string> {
  const { default: JSZip } = await import("jszip");
  const zip = new JSZip();
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:sldSz cx="${hostSlideSize.cx}" cy="${hostSlideSize.cy}"/>` +
      `</p:presentation>`,
  );
  return zip.generateAsync({ type: "base64" });
}

/**
 * Tag proxies whose `load()` has been queued but whose sync has not yet run.
 *
 * Flushed by a successful sync, dropped by a failed one — a load whose sync
 * rejects never populates anything, same as the real host.
 */
const pendingTagLoads: (() => void)[] = [];

/**
 * Shape `load()`s awaiting their sync — the same contract as `pendingTagLoads`,
 * for the properties a chart's tagged shape is identified by. Only read when
 * `faults.strictShapeReads` is on; queued always, so arming the fault
 * mid-render cannot grant a load that never landed.
 */
const pendingShapeLoads: (() => void)[] = [];

/**
 * A tag proxy that answers like Office.js does: resolving one gives you an
 * object whose properties are NOT readable until `load()` AND a sync.
 *
 * This fake used to populate `isNullObject` and `value` the instant
 * `getItemOrNullObject` was called, which made a whole class of bug invisible
 * here. Reading an unloaded proxy is `PropertyNotLoaded` on a real host, and
 * that is precisely how editing an ungrouped chart came to be impossible on
 * PowerPoint web while every test in this repo stayed green: `partIds` proxies
 * were resolved and never loaded, so the fake answered and the host threw.
 *
 * One honest proxy found the next one immediately (`deleteSlide`'s slot-tag
 * read). Being kinder than the host you model is not a neutral simplification —
 * it is a promise the tests cannot keep.
 */
function makeTagProxy(store: Map<string, string>, key: string, onUntrack?: () => void) {
  let readable = false;
  const need = (prop: string) => {
    if (!readable) {
      throw new Error(
        `PropertyNotLoaded: The property '${prop}' is not available. ` +
          "Before reading the property's value, call the load method on the containing object " +
          "and call context.sync() on the associated request context.",
      );
    }
  };
  return {
    load(_p?: string) {
      // Taken and never answered — see `faults.unansweredTagLoads`. The load
      // is still queued, so the sync carries it; nothing comes back.
      if (faults.unansweredTagLoads > 0) {
        faults.unansweredTagLoads--;
        return;
      }
      pendingTagLoads.push(() => {
        readable = true;
      });
    },
    untrack() {
      onUntrack?.();
    },
    get isNullObject() {
      need("isNullObject");
      return !store.has(key);
    },
    get value() {
      need("value");
      return store.get(key) ?? "";
    },
  };
}

export function makeShape(
  type: string,
  geo: string | undefined,
  box: { left: number; top: number; width: number; height: number },
) {
  const tagStore = new Map<string, string>();
  let ownId = `shape-${++idSeq}`;
  let ownLeft = box.left;
  let ownTop = box.top;
  /** Whether an answered `load()` has made this shape's properties readable. */
  let loadedProps = false;
  /**
   * Whether THIS handle — the one `addGeometricShape` returned — has been
   * resolved by a load of its own.
   *
   * Split from `loadedProps` on 2026-08-15, and the host is what split it.
   * `refuseTagWritesOnResolvedProxy` models "this host refuses a handle Office
   * has rewritten into `shapes.getItem(id)`", which is a fact about a HANDLE;
   * it was reading `loadedProps`, which is a fact about the SHAPE. So a
   * collection re-read — `shapes.load("items/id")`, which hands back fresh
   * handles onto the same shapes — marked the creation handle resolved too, and
   * the drawing context's tag write was refused however carefully its own load
   * was held back.
   *
   * That mattered because it refused a real fix. The question went to the host
   * rather than being argued: `collection-read-poisons-the-creation-handle`
   * answers **`yes`** — three passes, stable, every one taken while the host was
   * refusing collection reads. The creation handle still takes a tag after the
   * re-read. The fake was modelling handle state as shape state, exactly as
   * `freshHandle` already suspected by giving each handle its own `syncCreated`.
   *
   * Data readability stays SHARED, because that is genuinely shape state: once
   * any handle has loaded `id`, the value is on the object.
   */
  let handleResolved = false;
  const needLoaded = (prop: string) => {
    if (!faults.strictShapeReads || loadedProps) return;
    throw new Error(
      `The property '${prop}' is not available. Before reading the property's value, call the load ` +
        'method on the containing object and call "context.sync()" on the associated request context. ' +
        `| code=PropertyNotLoaded | errorLocation=Shape.${prop}`,
    );
  };
  const shape = {
    type,
    geo,
    box,
    fillColor: null as string | null,
    fillCleared: false,
    text: undefined as string | undefined,
    name: undefined as string | undefined,
    // A GETTER/SETTER PAIR, not a data property — see `faults.ignoreRotationWrites`,
    // which needs a write that is accepted and does nothing. The backing field
    // is what `width`'s own getter reads through `this.rotation` below.
    _rotation: undefined as number | undefined,
    get rotation(): number | undefined {
      return this._rotation;
    },
    set rotation(v: number | undefined) {
      if (faults.rotationWriteThrows) throw new Error("rotation is read-only on this host");
      if (faults.ignoreRotationWrites) return;
      this._rotation = v;
    },
    deleted: false,
    // The sync count at proxy creation — used to model the web host's
    // getItem(id) rewrite: a shape proxy is valid within the sync that queued
    // it plus the immediately following commit sync, and stale beyond that.
    // A hardened addGroup checks each member's age.
    syncCreated: trips.syncs,
    // A created shape's id exists only on the host: the renderer must load()
    // it back before it can write one down (see the parts tag).
    load(props?: string) {
      // The host refusing to say WHERE a chart landed — `groupAndTagAll`'s
      // last-ditch "reading back where the charts landed", and only that,
      // which is why the fault is keyed on the property set it asks for.
      //
      // Recorded on PowerPoint on the web, 2026-08-06, in the run that failed
      // `same scale across the deck`:
      //
      //   error  reading back where the charts landed
      //   InvalidParam passed to GetItem(id) | code=5010
      //   errorLocation: ShapeCollection.getItem
      //   statement: var shape = shapes.getItem(...); shape.load(["id","left","top"]);
      //
      // It matters because it is what leaves a chart DRAWN and nameless: no
      // group, no tag, and no id to settle one onto later. Nothing in the fake
      // could produce that state, so the hole it falls through in
      // `updateChartsInSlides` was unreachable from CI.
      if (faults.refuseIdLeftTopLoads > 0 && props === "id,left,top") {
        faults.refuseIdLeftTopLoads--;
        throw new Error("InvalidParam passed to GetItem(id) | code=5010");
      }
      // See faults.strictIdLoads. The age comes off `this` rather than the
      // closure, and it has to: `freshHandle` re-stamps `syncCreated` on the
      // proxy it hands back, so a closure over the original shape would report a
      // re-read member as being as old as the batch that drew it — the handle
      // would look fresh and behave stale, which is the exact trap the `tags`
      // rebinding below already documents.
      if (faults.refuseShapeIdLoads > 0 && props === "id") {
        faults.refuseShapeIdLoads--;
        throw new Error("InvalidParam passed to GetItem(id) | code=5010 | errorLocation: ShapeCollection.getItem");
      }
      const age = (this as unknown as { syncCreated?: number } | undefined)?.syncCreated ?? shape.syncCreated;
      if (faults.strictIdLoads && props === "id" && trips.syncs > age + 1) {
        throw new Error("InvalidParam passed to GetItem(id) | code=5010 | errorLocation: ShapeCollection.getItem");
      }
      // Taken and never answered when the shape was created in THIS batch —
      // see `faults.noIdInCreatingSync`. The load is queued like any other;
      // nothing comes back, so the property stays unreadable.
      if (faults.noIdInCreatingSync && shape.syncCreated === trips.syncs) return;
      // Queued, not granted. A load only takes effect on the sync that carries
      // it, and a sync that rejects carries nothing — same contract as
      // `pendingTagLoads`, and the reason `faults.strictShapeReads` can tell
      // "loaded" from "asked for".
      pendingShapeLoads.push(() => {
        loadedProps = true;
        // Through the shape's OWN handle, so this one is resolved. A load
        // arriving through a fresh collection handle takes `loadDataOnly`
        // instead and leaves this alone — see `handleResolved`.
        if (!(this as unknown as { viaFreshHandle?: boolean } | undefined)?.viaFreshHandle) handleResolved = true;
      });
    },
    get id() {
      needLoaded("id");
      return ownId;
    },
    set id(v: string) {
      ownId = v;
    },
    get left() {
      needLoaded("left");
      return ownLeft;
    },
    set left(v: number) {
      ownLeft = v;
    },
    get top() {
      needLoaded("top");
      return ownTop;
    },
    set top(v: number) {
      // See faults.ignoresTopWrites — office-js#6183. The write is accepted and
      // discarded, exactly as the web host does it: no error, and a later read
      // answers the OLD value rather than the one just set.
      if (faults.ignoresTopWrites) return;
      ownTop = v;
    },
    get width() {
      // See faults.presetDrawsNothing: the host took the preset and put nothing
      // on the slide. Reported as no extent, which is what "not visible" looks
      // like from the API side.
      if (faults.presetDrawsNothing && geo === faults.presetDrawsNothing) return 0;
      // See faults.rotationBlindsTheRead: this shape was rotated, so the host
      // answers nothing about its geometry — while its unrotated neighbour in
      // the same batch answers normally.
      if (faults.rotationBlindsTheRead && typeof this._rotation === "number") return undefined as unknown as number;
      // See faults.reportRotatedBounds. A rotated rectangle's bounding box is
      // |w·cos| + |h·sin| across, which for a long thin diagonal is far shorter
      // than the rectangle itself — that gap is what the scenario measures.
      if (faults.reportRotatedBounds && typeof this.rotation === "number" && this.rotation % 180 !== 0) {
        const rad = (this.rotation * Math.PI) / 180;
        return Math.abs(box.width * Math.cos(rad)) + Math.abs(box.height * Math.sin(rad));
      }
      return box.width;
    },
    set width(v: number) {
      box.width = v;
    },
    height: box.height,
    tagStore,
    get tags() {
      // See faults.tagsUndefinedOn — a host that hands back a shape with no
      // tags collection at all.
      if (faults.tagsUndefinedOn > 0) {
        faults.tagsUndefinedOn--;
        return undefined as unknown as FakeShape["tagsImpl"];
      }
      return shape.tagsImpl;
    },
    tagsImpl: {
      add: (k: string, v: string) => {
        // The web host rewrites a proxy as `shapes.getItem(id)` and rejects it
        // once stale — the SAME trap addGroup falls into, seen in the wild as
        // `InvalidParam passed to GetItem(id)`, code 5010, at
        // `ShapeCollection.getItem`, 28 times in one 38-item run. Tagging
        // crosses a sync boundary just like grouping does.
        if (faults.strictTags && trips.syncs > shape.syncCreated + 1) {
          throw new Error("InvalidParam passed to GetItem(id) | code=5010");
        }
        // See `faults.refuseTagWritesOnResolvedProxy`. Resolution, not age.
        if (faults.refuseTagWritesOnResolvedProxy && handleResolved) {
          throw new Error(
            "InvalidParam passed to GetItem(id) | code=5010 | errorLocation=ShapeCollection.getItem | " +
              "statement=var shape = shapes.getItem(...) /* originally addTextBox(...) */",
          );
        }
        if (faults.refuseTagWrites > 0) {
          faults.refuseTagWrites--;
          throw new Error("InvalidParam passed to GetItem(id) | code=5010");
        }
        tagStore.set(k, v);
      },
      getItemOrNullObject: (k: string) =>
        makeTagProxy(tagStore, k, () => {
          untracked.tags++;
        }),
    },
    // Set by fill.setImage — the bare base64 the host received. Records what the
    // renderer actually handed over, so a test can prove the data:/image-prefix
    // stripping happened rather than just that setImage was reached.
    imageBase64: undefined as string | undefined,
    fill: {
      setSolidColor(c: string) {
        shape.fillColor = c;
      },
      setImage(b64: string) {
        // Model the real API's one visible side effect: "This changes the fill
        // type to PictureAndTexture" (@types/office-js). `faults.refusePictureFill`
        // makes the property access itself throw, which is what a pre-1.8 host
        // does — the method simply is not there.
        if (faults.refusePictureFill) throw new Error("setImage is not available on this host");
        shape.imageBase64 = b64;
        shape.fillType = "PictureAndTexture";
        // The moment `slideReadsEmpty: "after-a-picture"` is waiting for.
        pictureLanded = true;
      },
      clear() {
        shape.fillCleared = true;
      },
    },
    fillType: undefined as string | undefined,
    lineFormat: {} as Record<string, unknown>,
    textFrame: {
      // `textRange.text` is a WRITE-THROUGH to the shape's own string, not a
      // property that happens to sit on a bag. `addTextBox` takes the string as
      // an argument, so this setter is the only way to change one afterwards —
      // and it is the whole of what an in-place retitle does. Modelled as an
      // inert field, the fake would have let a fast path that changed nothing
      // at all pass every assertion about which shapes it touched.
      textRange: {
        font: {} as Record<string, unknown>,
        paragraphFormat: {} as Record<string, unknown>,
        get text(): string | undefined {
          return shape.text;
        },
        set text(v: string | undefined) {
          shape.text = v;
        },
      },
    } as Record<string, unknown> & {
      textRange: { font: Record<string, unknown>; paragraphFormat: Record<string, unknown>; text?: string };
    },
    grouped: undefined as unknown[] | undefined,
    // PowerPointApi 1.8's `Shape.group` — present only on an actual group, and
    // a throw otherwise, because that is what a repair pass has to survive: on
    // a host without 1.8 the property access poisons the sync it was queued in.
    get group() {
      // See `faults.refuseGroupRead`. THE ERROR SHAPE MATTERS AS MUCH AS THE
      // THROW: production latches on `errorLocation: "Shape.group"` and nothing
      // else, so a generic error here exercises none of that path — which is why
      // a mutant that skipped clearing the latch survived until this existed.
      // AT THE SYNC, NOT AT THE ACCESS. Production says so in as many words:
      // "the access only QUEUES a proxy read, and the GeneralException arrives
      // later at the sync — under a different label, outside that try, in a batch
      // it can poison". A fault that throws synchronously here is a DIFFERENT
      // bug: it takes down the caller's try block instead of the sync, and it
      // poisoned the re-read before the retry under test was ever reached.
      //
      // Counted, so a test can arm "refuse once, then answer" — the only shape
      // in which a retry can be shown to PAY rather than merely to run.
      const refuseGroup = faults.refuseGroupReadTimes > 0 || faults.refuseGroupRead;
      if (faults.refuseGroupReadTimes > 0) faults.refuseGroupReadTimes--;
      if (refuseGroup) {
        pendingHostError = new Error(
          'GeneralException | code=GeneralException | debugInfo={"code":"GeneralException","message":"GeneralException","errorLocation":"Shape.group"}',
        );
      }
      if (!shape.grouped) throw new Error("shape is not a group");
      const live = (): { name?: string; deleted: boolean }[] =>
        (shape.grouped as { name?: string; deleted: boolean }[]).filter((c) => !c.deleted);
      return {
        shapes: {
          getCount: () => ({ value: live().length }),
          get items() {
            return live();
          },
          load() {},
        },
      };
    },
    delete() {
      shape.deleted = true;
      // Deleting a group deletes what it contains — otherwise an in-place update
      // leaves every child on the slide and the fake shows orphans the host
      // would never have.
      for (const child of (shape.grouped ?? []) as { deleted: boolean }[]) child.deleted = true;
    },
    untrack() {
      untracked.shapes++;
    },
  };
  return shape;
}

export type FakeShape = ReturnType<typeof makeShape>;

/** Layout ids passed to slides.add() since the last installHost(). */
export const addedWithLayout: (string | undefined)[] = [];

/** Slide-master ids passed to slides.add() since the last installHost(). */
export const addedWithMaster: (string | undefined)[] = [];

/**
 * The layouts a new slide may name WITHOUT also naming a master.
 *
 * They belong to the deck's DEFAULT master — the previous slide's — which is
 * the second one under `faults.twoMasterDeck`. A layout from the other master
 * is only legal when the add names its master too, which is the rule
 * `slides.add` models above.
 */
export const LAYOUTS_VALID_WITHOUT_A_MASTER = new Set(["m2-layout-blank"]);

/**
 * The queued-command failure a stalled getItemAt handle produces, surfaced at
 * the NEXT sync — because Office.js reports queued-command errors only at sync,
 * never at the call that queued them. `installHost`'s sync() throws it. See
 * `freshWindowedHandle` below for why a reused fresh-slide handle poisons.
 */
let pendingHostError: Error | null = null;

/**
 * The slide count at the start of the current `PowerPoint.run` — every slide at
 * an index >= this was `add()`ed during this context and so is "fresh". A fresh
 * slide's getItemAt handle is only good within the sync it was acquired in (see
 * `freshWindowedHandle`); a pre-existing slide's handle is durable.
 */
let contextBaseCount = 0;

/**
 * Syncs issued inside the current `PowerPoint.run` — reset at every new context.
 *
 * The counterpart to `trips.syncs`, which never resets. Together they are the
 * two hypotheses `faults.syncCostMs` lets a test choose between: a cost that
 * follows this one is the context accumulating, a cost that follows the total
 * is the host itself slowing down.
 */
let syncsInContext = 0;
/**
 * True only while `run` is making the closing auto-sync that carries a batch's
 * return value out. See both call sites: it is exercised so a wrapper around
 * `sync` is caught, and not charged so the suite's fault indices stay put.
 */
let inClosingAutoSync = false;

/** Set by installHost so a slide can splice itself out of the live deck. */
let deckRemove: ((s: { id: string }) => void) | null = null;

/**
 * Every live shape in the deck, by id — set by `installHost`.
 *
 * Only `userClicksShape` needs it: a click names a shape, and the deck lives
 * in `installHost`'s closure where no helper can reach it.
 */
let findShape: ((id: string) => FakeShape | undefined) | null = null;

/**
 * A fresh proxy onto a shape, as `shapes.items` hands back after a load+sync.
 *
 * Everything reads and writes through to the one real shape — a tag added via
 * a fresh handle is on the shape, because it is the same shape. What the
 * handle does NOT share is `syncCreated`: this one was minted now, so it is
 * current, while whatever handle the caller was holding before stays as old as
 * it was. That is the whole distinction, and modelling it is what lets a test
 * tell "re-fetched the collection and used the new handle" apart from
 * "re-fetched the collection and went on using the old one".
 */
function freshHandle(shape: FakeShape, override?: { id: string }): FakeShape {
  let own = trips.syncs;
  return new Proxy(shape, {
    get(target, prop, recv) {
      if (prop === "syncCreated") return own;
      // An id the caller supplies instead of the shape's own — for a host that
      // lists a slide under ids that have not settled yet. See
      // `faults.unmatchedIdsFirst`. Deliberately only the id: it is still the
      // same shape, and everything else about it must stay true.
      if (prop === "id" && override) return override.id;
      // A LOAD THROUGH THIS HANDLE RESOLVES THIS HANDLE, and not the one that
      // drew the shape. `load` reads the flag off `this`, and `this` is the
      // proxy — the same trick `syncCreated` above already relies on. Without
      // it, `shapes.load("items/id")` marked the creation handle resolved and a
      // tag written through it was refused, which is the fake refusing a fix the
      // host allows (`collection-read-poisons-the-creation-handle: yes`).
      if (prop === "viaFreshHandle") return true;
      // `tags` has to be rebound too. The shape's own tags object closes over
      // the shape's age, so a fresh handle reaching for `.tags` would get a
      // writer that still consults the STALE number and refuses — the handle
      // would look fresh and behave stale. Everything else (the tag store, the
      // geometry, the delete) is deliberately shared: it is one shape.
      if (prop === "tags") {
        return {
          add: (k: string, v: string) => {
            if (faults.strictTags && trips.syncs > own + 1) {
              throw new Error("InvalidParam passed to GetItem(id) | code=5010");
            }
            // Here too, and the reason is the warning this file already gives
            // about `hollowNameReads`: a fault armed on one of two identical
            // writers is a fault that half-fires. This is the writer the
            // SETTLE pass uses — it re-resolves the shape to get a fresh
            // handle — so a fault that skipped it could only ever refuse the
            // drawing context's write, which `settleAndTagChart` then repairs.
            // A host on which the second chance also fails is the one the
            // 2026-08-06 round found, and it was unreachable from here.
            if (faults.refuseTagWrites > 0) {
              faults.refuseTagWrites--;
              throw new Error("InvalidParam passed to GetItem(id) | code=5010");
            }
            target.tagStore.set(k, v);
          },
          getItemOrNullObject: target.tags.getItemOrNullObject,
        };
      }
      return Reflect.get(target, prop, prop === "syncCreated" ? recv : target);
    },
    set(target, prop, value) {
      if (prop === "syncCreated") {
        own = value as number;
        return true;
      }
      return Reflect.set(target, prop, value);
    },
  });
}

/**
 * Wrap a `getItemOrNullObject` result the way Office.js hands one over.
 *
 * `isNullObject` is not a property you can just read: it is populated on the
 * sync the proxy participated in, and a proxy nobody called `load()` on never
 * participates. Reading it then throws PropertyNotLoaded. The fake answered it
 * unconditionally, so a whole class of bug — resolve a shape, never load it,
 * read `isNullObject` — was invisible here and fatal on a real host.
 *
 * The load must name a REAL property. `load("isNullObject")` looks like the
 * obvious way to ask and is not one: `isNullObject` is not a property the host
 * holds, it is a flag Office.js sets from the response to a load of real
 * properties. Selecting it alone asks the host for nothing, the proxy takes no
 * part in the sync, and the flag stays unreadable — `PropertyNotLoaded`, with
 * no `errorLocation`, because the getter is on Office.js's base class.
 *
 * The fake accepted it, so five resolves in `powerpoint.ts` were written that
 * way and every test agreed with them. On PowerPoint on the web the one on the
 * in-place update path meant editing a chart threw before it did anything —
 * "edit a chart on the visible slide", the first self-test scenario to fail on
 * a real host.
 */
function nullObjectProxy<T extends object>(found: T | undefined, kind: "slide" | "shape" = "slide") {
  let loaded = false;
  const base = found ?? ({ delete() {} } as unknown as T);
  // Read once, off the raw object, before any strictness gate can refuse it —
  // the fault is addressed by id and must not depend on the fault it models.
  const ownId = (() => {
    try {
      return (found as { id?: string } | undefined)?.id;
    } catch {
      return undefined;
    }
  })();
  return new Proxy(base, {
    get(target, prop, recv) {
      if (prop === "load")
        return (p?: string | string[]) => {
          const asked = (Array.isArray(p) ? p : (p ?? "").split(",")).map((s) => s.trim()).filter(Boolean);
          if (asked.length && asked.every((a) => a === "isNullObject")) return;
          // The host took the load and answered nothing — see
          // `unansweredNullChecks`. One-shot, so a retry can find it.
          if (ownId && unansweredNullChecks.has(ownId)) {
            unansweredNullChecks.delete(ownId);
            return;
          }
          // ON THE SYNC, not on the call — the same queue every other load in
          // this fake goes through. This assignment used to be immediate, which
          // made `isNullObject` readable the instant someone ASKED for it, and
          // that is a state Office.js never produces: a load is a request, and
          // only the sync answers it.
          //
          // The fiction hid a live production defect. A batch that PowerPoint
          // refuses populates nothing, so the real host leaves this unreadable
          // and the caller cannot tell a refused lookup from a deleted shape;
          // the fake answered `isNullObject = false` on a proxy whose sync had
          // just thrown, so the update sailed on and the recovery for it could
          // not be tested at all. `discard()` clears this queue on exactly the
          // paths where the host would have answered nothing.
          pendingShapeLoads.push(() => {
            loaded = true;
          });
        };
      if (prop === "isNullObject") {
        if (!loaded)
          throw new Error(
            "The property 'isNullObject' is not available. Before reading the property's value, call the load " +
              'method on the containing object and call "context.sync()" on the associated request context.',
          );
        return !found;
      }
      // See `faults.shapeIdNeverPopulates`. AFTER the isNullObject branch on
      // purpose: the proxy is not null, which is the whole point — the host
      // found the shape and answered nothing about it.
      if (prop === "id" && kind === "shape" && found && faults.shapeIdNeverPopulates) return undefined;
      return Reflect.get(target, prop, recv);
    },
  });
}

export function makeSlide(id: string) {
  // Reassignable: the `id` accessor installed at the end of this function reads
  // and writes it, so the property keeps behaving like the plain field it was.
  const created: FakeShape[] = [];
  // Shapes queued since the last successful sync. On success the pending list
  // is cleared (they're committed). On a failed sync the fake removes them
  // from `created` — the real host discards queued commands whose sync
  // rejected, and modeling that faithfully is what lets a mid-render
  // failure-mode test distinguish "40 batches committed, 41st batch dropped"
  // from "all 47 shapes on the slide" (the PR-8 partial-landed threshold).
  const pending: FakeShape[] = [];
  const slideTagStore = new Map<string, string>();
  const slide = {
    // Plain data by default; a throwing getter only when a test arms
    // `slideIdUnreadableBeforeFirstSync`. Defined below with
    // `Object.defineProperty` so the default stays an own enumerable value and
    // nothing about spreading or JSON-ing a slide changes.
    id,
    created,
    pending,
    isNullObject: false,
    /**
     * Which master this slide is built from — `Slide.slideMaster`, API 1.3.
     *
     * The deck's own slides sit on the SECOND master under
     * `faults.twoMasterDeck`. That is what makes the first master "foreign",
     * and it is the whole reason the real host refuses a layout taken from it.
     * Without this the fake cannot express a deck whose master differs from the
     * one a naive walk of `slideMasters` picks — so it cannot tell a chart
     * slide that matches its neighbours from one wearing a different theme.
     */
    slideMaster: {
      /**
       * SUFFIXED, because the host renders one master's id two ways.
       *
       * Round 377: the master walked out of `slideMasters` answered
       * `2147483660#2460954070`; the SAME master read through
       * `Slide.slideMaster` answered `2147483660`. A plain `===` between the
       * two called them different and reported a mismatch about a deck that
       * matched perfectly. This fake used to hand back an identical string
       * from both, so the comparison looked sound and no test could say
       * otherwise.
       */
      get id() {
        return faults.twoMasterDeck ? "master-2#7788" : "master-1#7788";
      },
      load() {},
    },
    /**
     * The slide's layout, with a shape collection that answers empty.
     *
     * Modelled because office-js#3826 says a NEW slide's
     * `load("layout/shapes/items")` fails the sync with GeneralException on the
     * web, and `new-slide-layout-readable` asks it. Without a layout here the
     * fake answered "unreadable" — which is not "we do not model this", it is a
     * claim that the host refuses, and a real desktop PowerPoint does not. A
     * fake that lies on the happy path turns every real answer into a
     * divergence.
     *
     * Empty on purpose: this add-in creates slides on the BLANK layout, so zero
     * layout shapes is the honest neutral answer, and nothing in the repo reads
     * layout shapes for their content.
     */
    layout: { shapes: { items: [] as unknown[], load() {} }, load() {} },
    load() {},
    // Slide.delete() — the repair pass removes duplicate slides, and a fake
    // that kept them would make every deletion assertion vacuous. installHost
    // owns the deck array, so removal goes through the hook it registers.
    delete() {
      // Taken and not performed — see faults.refuseSlideDelete.
      if (faults.refuseSlideDelete) return;
      deckRemove?.(slide);
    },
    /**
     * PowerPointApi 1.8's single-slide export.
     *
     * Returns a real .pptx (base64) carrying the deck's `<p:sldSz>` — the part
     * `slideSize`'s middle rung parses. A ClientResult, so `.value` is empty
     * until the sync it was queued in lands, exactly like the real one.
     */
    exportAsBase64() {
      const result = { value: "" };
      pendingExports.push({ result, build: exportedDeckBase64 });
      return result;
    },
    /**
     * PowerPointApi 1.5's shape selection — on **Slide**, which is the point.
     *
     * `Presentation` has `getSelectedShapes` and `setSelectedSlides` but no
     * `setSelectedShapes`; the shape half is one class down. Believing
     * otherwise is why the in-host battery went months without covering the
     * pane's selection-driven paths at all.
     *
     * Writes through to the same array `getSelectedShapes()` reads, so a
     * select-then-read round trip means something here.
     */
    setSelectedShapes(ids: string[]) {
      noteSelectionCall();
      if (!ids.length) trips.emptyDeselects++;
      // Taken, and poisonous. The select itself still happens below — the web
      // host does perform it — and only what comes AFTER stops answering.
      if (faults.selectionWedgesHost && ids.length) selectionWedged = true;
      // Tracked for `stallDrawAfterSelect`: a selection is STANDING until it is
      // cleared, which is the window the real host refuses a draw in.
      selectionStanding = ids.length > 0;
      if (!ids.length) {
        // Cleared on desktop, IGNORED on PowerPoint on the web
        // (office-js#3083). Modelled as the web does it when the fault is
        // armed, because a scenario that relies on the clear working is one
        // that will pass here and strand the next scenario in the field.
        if (!faults.webIgnoresDeselect) selectionRef.length = 0;
        return;
      }
      const found = faults.selectionIgnoresIds
        ? created.filter((s) => !s.deleted).slice(0, 1)
        : created.filter((s) => !s.deleted && ids.includes(s.id));
      selectionRef.length = 0;
      selectionRef.push(...found);
    },
    /**
     * PowerPointApi 1.8's slide raster. A ClientResult like `exportAsBase64`.
     *
     * The fake cannot draw, so it answers with a PNG whose payload encodes what
     * is ON the slide — live shape count and total ink area. That is enough for
     * the only assertion worth making here: an EMPTY slide and a slide with a
     * chart on it must not produce the same image. A fake that returned a fixed
     * string would let a scenario claiming "the chart is visible" pass over a
     * blank slide, which is precisely the failure the scenario exists to catch.
     */
    getImageAsBase64(_options?: { width?: number }) {
      const result = { value: "" };
      pendingExports.push({
        result,
        build: async () => {
          if (faults.slideImageGate) await faults.slideImageGate;
          if (faults.constantSlideImage) return btoa("PNG:blank");
          // Taken, raised nothing, produced nothing — see
          // `faults.emptySlideImage`. The quietest failure a rasteriser has,
          // and the one a real host produced with no line in the log to say so.
          if (faults.emptySlideImage) return "";
          const live = created.filter((s) => !s.deleted);
          const ink = live.reduce((n, s) => n + Math.max(0, s.width) * Math.max(0, s.height), 0);
          const payload = `PNG:${slide.id}:shapes=${live.length}:ink=${ink}`;
          rasterised.push(payload);
          return btoa(payload);
        },
      });
      return result;
    },
    tags: {
      add: (k: string, v: string) => void slideTagStore.set(k, v),
      getItemOrNullObject: (k: string) => makeTagProxy(slideTagStore, k),
    },
    tagStore: slideTagStore,
    shapes: {
      // A deleted shape is gone from the host's collection; keeping it in `items`
      // let readbacks (listChartsInDeck) count a chart that no longer exists.
      //
      // Reading the collection hands back FRESH HANDLES, and leaves every
      // handle taken earlier exactly as stale as it was. This is the one place
      // the fake used to be kinder than the host it models: it refreshed
      // `syncCreated` on the shape objects themselves, so a re-fetch anywhere
      // healed a stale proxy held anywhere else. Real Office.js returns new
      // proxy objects — your old variable still points at the old proxy, and
      // the old proxy is still refused.
      //
      // That difference is not academic. It is precisely why a whole class of
      // stale-proxy bug was invisible here and had to be found by a human
      // running the add-in in a real PowerPoint: the code could re-fetch a
      // collection, tag the WRONG (old) handle, and this fake would let it
      // through. Restoring the distinction is what makes `web-host.test.ts`
      // able to catch it.
      get items() {
        // The host never answered the load queued for this collection — see
        // `unansweredShapeReads`. Nothing to hand back, and nothing to say
        // about what is on the slide.
        if (unansweredShapeReads.has(id)) {
          unansweredShapeReads.delete(id);
          throw new Error(
            "The property 'items' is not available. Before reading the property's value, call the load " +
              'method on the containing object and call "context.sync()" on the associated request context. ' +
              "| code=PropertyNotLoaded | errorLocation=ShapeCollection.items",
          );
        }
        // An arrow, not a bare `.map(freshHandle)`: `map` passes the INDEX as the
        // second argument, which `freshHandle`'s id override would take as a
        // config object. Typecheck catches it today; it would be a silent bug the
        // moment that parameter took something more forgiving than an object.
        // See `faults.groupHidesChildren`: a real slide lists the GROUP and not
        // its members, and this fake lists both.
        // BY ID, not by identity: `addGroup` is handed shape HANDLES, and a
        // handle is not the FakeShape it stands for — a Set of the members
        // matched nothing at all, and the mutant that removed the descent under
        // test walked straight through.
        const grouped = faults.groupHidesChildren
          ? new Set(
              created.flatMap((s) =>
                ((s.grouped ?? []) as { id?: string }[]).map((c) => String(c?.id ?? "")).filter(Boolean),
              ),
            )
          : null;
        const live = created
          .filter((s) => !s.deleted && !(grouped && grouped.has(String(s.id))))
          .map((s) => freshHandle(s));
        // The web host has been observed answering a shape-collection read
        // with FAR fewer shapes than it holds: one readback page asked about
        // 19 slides carrying 19 shapes and got 3 back. `faults.hollowReads` models
        // that — the collection is short, nothing throws, and the caller has
        // no way to know unless it compares against a count it took earlier.
        // `startsWith`, not equality: the deck scan asks for
        // `items/id,items/left,items/top`, which is the same collection read
        // under a wider projection and which a real host has no reason to treat
        // differently. Matching only the exact string meant `hollowReads` could
        // not blind `listChartsInDeck` at all — so the tests that thought they
        // were exercising a short deck scan were exercising a healthy one.
        // A host that answers honestly and then stops, rather than one that is
        // short from the first read. `hollowReads` blinds the FIRST N reads,
        // which is the opposite of how a real host fails: the 2026-08-05 run
        // answered fine for two minutes and was reading short by the time it
        // reached the fifth scenario. A scenario that checks its first scan for
        // blindness and trusts the ones after it can only be caught this way.
        if (faults.hollowReadsAfter !== null && lastShapeLoad.startsWith("items/id")) {
          if (faults.hollowReadsAfter > 0) faults.hollowReadsAfter--;
          else return [];
        }
        if (slideReadsEmptyNow()) return [];
        if (faults.hollowReads > 0 && lastShapeLoad.startsWith("items/id")) {
          faults.hollowReads--;
          return [];
        }
        // Short, not empty — the case `hollowReads` describes and does not do.
        if (faults.readsMissing > 0 && lastShapeLoad.startsWith("items/id")) {
          return live.slice(0, Math.max(0, live.length - faults.readsMissing));
        }
        // Short ONCE, then honest — a slide still settling. See
        // `readsMissingFirst`; checked after `readsMissing` so a test arming
        // both gets the permanent answer, which is the stricter one.
        if (faults.readsMissingFirst > 0 && lastShapeLoad.startsWith("items/id")) {
          const drop = faults.readsMissingFirst;
          faults.readsMissingFirst = 0;
          return live.slice(0, Math.max(0, live.length - drop));
        }
        // Lists everything, under ids nothing can be joined to — the THIRD way
        // this re-read fails, and the one the fake could not express. See
        // `unmatchedIdsFirst`.
        if (faults.unmatchedIdReads > 0 && lastShapeLoad.startsWith("items/id")) {
          faults.unmatchedIdReads--;
          return live.map((sh) => freshHandle(sh, { id: `unsettled-${sh.id}` }));
        }
        // Any projection that asks for names, not the exact string `items/name`.
        // `slideShapeList` asks for `items/id,items/name` — it needs ids to tell
        // one shape from another — and exact matching silently stopped blinding
        // it the moment that projection changed, which turns a guard green
        // without anyone touching it. Same lesson as `hollowReads`' own
        // `startsWith` a few lines up, learned the same way.
        if (faults.hollowNameReads > 0 && lastShapeLoad.includes("name")) {
          faults.hollowNameReads--;
          return [];
        }
        return live;
      },
      // The tag pass is the only reader that asks for `items/id`; pass A and
      // the group-child count both ask for `items/name`. Recording which lets
      // a test make the TAG read hollow without also blinding the count it is
      // supposed to be checked against.
      load(p?: string) {
        if (faults.faultShapeCollectionLoad) throw new TypeError("e.load is not a function");
        if (p) lastShapeLoad = p;
        // `items/id,items/left,items/top` loads the SHAPES, not just the
        // collection — that is what the selector means, and it is how the real
        // host populates them. Without this the fake would be stricter than
        // Office.js rather than equal to it, and `strictShapeReads` would fail
        // correct code, which is the mirror image of the sin it exists to
        // prevent.
        // THROUGH A FRESH HANDLE, which is what this read actually hands back
        // (see the `items` getter below, and `freshHandle`). Calling the shape's
        // own `load` marked the CREATION handle resolved as a side effect, so a
        // tag written through it was refused under
        // `refuseTagWritesOnResolvedProxy` — the fake refusing a fix this host
        // allows. `collection-read-poisons-the-creation-handle` answers `yes`,
        // three passes, stable: after this read the creating handle still takes
        // a tag. Data readability is shared, because that is shape state; the
        // resolution is not, because that is handle state.
        if (p?.includes("items/")) for (const s of created) freshHandle(s).load();
      },
      addGeometricShape(geo: string, box: FakeShape["box"]) {
        const s = makeShape("geometric", geo, box);
        // Taken and refused at the sync — see `faults.refuseShapeAdds`. The
        // shape is still handed back, because a queued-command failure is what
        // a real host answers with: the caller holds a proxy for something
        // that never lands.
        // The COUNTED form, for the case the flag cannot express: a refusal
        // that lifts. Buying a replacement slide only earns its cost when the
        // question can succeed on the new one, and with the flag on it never
        // can — so the flag alone can only ever show the replacement failing.
        // `swallowAdds` is the same idea one layer up.
        if (faults.refuseShapeAddsTimes > 0) faults.refuseShapeAddsTimes--;
        else if (!faults.refuseShapeAdds) {
          created.push(s);
          pending.push(s);
          return s;
        }
        {
          pendingHostError = new Error(
            'GeneralException | code=GeneralException | debugInfo={"code":"GeneralException","message":"GeneralException","errorLocation":"ShapeCollection.addGeometricShape"}',
          );
          return s;
        }
        created.push(s);
        pending.push(s);
        return s;
      },
      addLine(_kind: string, box: FakeShape["box"]) {
        const s = makeShape("line", undefined, box);
        created.push(s);
        pending.push(s);
        return s;
      },
      addTextBox(text: string, box: FakeShape["box"]) {
        // office-js#2775: on PowerPoint on the web, adding a text box DELETES
        // whatever shape was selected. Reported against ScriptLab's own image
        // sample, open, and explicitly "works fine on desktop".
        //
        // Modelled because every chart this add-in draws contains text boxes,
        // and the insert path deliberately leaves the user's shape selected —
        // so a real host behaving this way destroys the user's own content on
        // the everyday path. Off by default: nothing has established that the
        // build the owner runs still does it, and a fake must not assert host
        // behaviour nobody has asked about.
        if (faults.textBoxDeletesSelection) {
          for (const victim of selectionRef.splice(0)) {
            victim.deleted = true;
            const i = created.indexOf(victim);
            if (i >= 0) created.splice(i, 1);
          }
        }
        const s = makeShape("text", undefined, box);
        s.text = text;
        created.push(s);
        pending.push(s);
        return s;
      },
      addGroup(items: FakeShape[]) {
        // A host that refuses to group at all — the call throws where the API
        // exists but the host declines. Consumed per call, so a test can refuse
        // the render's own grouping and still let the fresh-context rescue's
        // addGroup work, which is the whole point of the rescue.
        if (faults.refuseGroups > 0) {
          faults.refuseGroups--;
          throw new Error("host refused addGroup");
        }
        // Model the web-host stale-proxy trap: a Shape proxy is valid within
        // the sync that queued it plus its immediately following commit sync,
        // and stale (getItem(id) rewrite, non-round-trippable id) beyond that.
        // addGroup(theseStaleProxies) silently loses grouping on real Office —
        // no group appears on the slide. Enable via faults.strictGroup.
        // A member whose PARENT handle has expired, however young the member
        // itself is. Unconditional, like `expiringSlideHandle`: the host named
        // this one in its own error, and the members' ages were not the reason.
        const orphan = items.find((s) => {
          const ok = (s as FakeShape & { parentWindowOk?: () => boolean }).parentWindowOk;
          return typeof ok === "function" && !ok();
        });
        if (orphan) {
          throw new Error(
            'InvalidParam passed to GetItem(id) | code=5010 | debugInfo={"code":"5010","message":"InvalidParam passed to GetItem(id)","errorLocation":"ShapeCollection.getItem"}',
          );
        }
        if (faults.strictGroup && items.some((s) => trips.syncs > s.syncCreated + 1)) {
          const g = makeShape("group", undefined, { left: 0, top: 0, width: 0, height: 0 });
          // Not pushed to `created` — no group lands on the slide, exactly
          // what the web host does for a stale-proxy addGroup call.
          return g;
        }
        // A real group's frame is the BOUNDING BOX of its children, not the
        // origin the caller drew at. Returning {0,0,0,0} made a group's position
        // meaningless and hid the update-drift bug entirely: the renderer reads
        // this back as an EditTarget and (before CHART_ORIGIN_TAG) fed it in as
        // the next render's frame origin, walking the chart down the slide.
        const box = items.length
          ? {
              left: Math.min(...items.map((s) => s.left)),
              top: Math.min(...items.map((s) => s.top)),
              width: Math.max(...items.map((s) => s.left + s.width)) - Math.min(...items.map((s) => s.left)),
              height: Math.max(...items.map((s) => s.top + s.height)) - Math.min(...items.map((s) => s.top)),
            }
          : { left: 0, top: 0, width: 0, height: 0 };
        const g = makeShape("group", undefined, box);
        g.grouped = items;
        created.push(g);
        pending.push(g);
        return g;
      },
      getItemOrNullObject(id: string) {
        // A null-object proxy still accepts load() in real Office.js — queuing a
        // load on it is the normal pattern, and only isNullObject tells you it
        // resolved to nothing. Without load() here the fake threw where the host
        // would not.
        //
        // And `isNullObject` is only READABLE once the proxy has been loaded.
        // This fake used to hand it over unconditionally, which is a fiction:
        // Office.js populates it on the sync the proxy took part in, and a
        // proxy nobody loaded takes part in nothing. Reading it then throws
        // PropertyNotLoaded — which is exactly what editing an ungrouped chart
        // did on a real host while every test here passed, because the fake
        // answered a question the host refuses. Model the refusal.
        // See faults.refuseShapeById. Armed, the LOOKUP is what poisons the
        // batch — the proxy is still handed back so the caller's synchronous
        // `.tags.add` and `.load` queue exactly as they do against the real
        // host, and the refusal lands where Office.js puts it: the sync.
        if (faults.refuseShapeById) refuseThisSync = true;
        const found = created.find((s) => s.id === id && !s.deleted);
        // A FRESH handle, for the same reason `items` hands back fresh ones:
        // resolving by id makes a new proxy object, and a new proxy has not
        // been rewritten to anything. The staleness this file models is what
        // happens to a proxy AFTER the sync that resolved it — hand back the
        // original shape here and its creation-time age comes with it, so a
        // handle resolved in the very batch that uses it is refused as though
        // it were minutes old.
        //
        // That is not what the host does, and the repair pass is the proof: it
        // resolves a chart and writes its tag in the next batch, and landed 23
        // of them in the same real run that lost 46 tag writes through genuinely
        // stale proxies. The fake refusing this pattern made the ordinary
        // path's recovery untestable — it could only be written by making the
        // fake lie in the other direction.
        return nullObjectProxy(found ? freshHandle(found) : undefined, "shape");
      },
      // Top-level shape count the host reports on readback — non-deleted shapes.
      getCount: () => {
        if (faults.faultShapeGetCount) throw new Error("readback getCount faulted");
        // Agrees with the empty `items` above — see `slideReadsEmpty`. Both
        // signals saying zero is the whole point: one of them telling the truth
        // is the case `slideShapeList` already survives.
        if (slideReadsEmptyNow()) return { value: 0 };
        const live = created.filter((s) => !s.deleted).length;
        // THE HOST LAGGING ITS OWN COMMIT — see `faults.shapeCountLag`. Round 084
        // reported four slides growing by 23 shapes each while the deck they
        // belonged to ended holding one grouped chart apiece; every count was
        // taken 1.3s after an `addGroup` that had already synced, and answered
        // as though the group were not there. One read of this number is not a
        // measurement, and the fake could not say so until now.
        if (faults.shapeCountLag > 0) {
          faults.shapeCountLag--;
          return { value: live + faults.shapeCountLagBy };
        }
        return { value: live };
      },
      /**
       * `ShapeCollection.getItemAt(index)` — the API surface the fake was
       * missing, not a behaviour it was modelling.
       *
       * Added 2026-08-23 for `shapes-by-index-vs-items`, which asks whether this
       * host will walk a collection by INDEX where it refuses to read it as a
       * LIST. Without this the probe answered `threw` — "not a function" — which
       * would have been recorded as a fact about PowerPoint rather than about
       * the double. The real host demonstrably HAS it: `getitemat-past-end`
       * answers `threw` in 158 of 158 rounds, which is the correct answer to an
       * out-of-range index and impossible from a method that does not exist.
       *
       * Modelled honestly and no more: live shapes, in order, and a throw past
       * the end. It does NOT model the host's deafness — the fake's collection
       * answers, so the probe reads `both-answer` here and only the real host
       * can produce the interesting one. A double that reproduced the deafness
       * would be modelling a bug instead of a contract.
       */
      getItemAt(index: number) {
        const live = created.filter((s) => !s.deleted);
        if (!Number.isInteger(index) || index < 0 || index >= live.length)
          throw new Error(
            `InvalidArgument | code=InvalidArgument | debugInfo={"code":"InvalidArgument","errorLocation":"ShapeCollection.getItemAt"}`,
          );
        return live[index];
      },
    },
  };
  // See `faults.slideIdUnreadableBeforeFirstSync`. Installed unconditionally
  // and gated inside, so a test can arm the fault after the deck is built; the
  // getter is a plain read-through in every other case.
  Object.defineProperty(slide, "id", {
    enumerable: true,
    configurable: true,
    get() {
      if (faults.slideIdNeverReadable || (faults.slideIdUnreadableBeforeFirstSync && syncsInContext === 0))
        throw new Error(
          "The property 'id' is not available. Before reading the property's value, call the load " +
            'method on the containing object and call "context.sync()" on the associated request context. ' +
            "| code=PropertyNotLoaded | errorLocation=Slide.id",
        );
      return id;
    },
    set(v: string) {
      id = v;
    },
  });
  // The fake's OWN bookkeeping reads through this, never through `id`.
  //
  // `slides.add` files each new slide under its id, and doing that through the
  // throwing accessor made `faults.slideIdNeverReadable` break the fake instead
  // of the code under test: the add itself raised, no slide landed, and a test
  // aimed at "the slide exists and has no name" quietly became another
  // dropped-add test. Non-enumerable so nothing that spreads or JSONs a slide
  // sees a second id.
  Object.defineProperty(slide, "rawId", {
    enumerable: false,
    configurable: true,
    get: () => id,
  });
  // Cast rather than a literal field: `defineProperty` is invisible to
  // inference, and a plain `get rawId()` in the literal would be enumerable —
  // which is the one thing the note above says it must not be.
  return slide as typeof slide & { readonly rawId: string };
}

export type FakeSlide = ReturnType<typeof makeSlide>;

/**
 * The reference `slides.getItemAt(i)` hands back for a FRESHLY-ADDED slide.
 *
 * It draws fine within the sync it was acquired in, but reusing it to draw AFTER
 * a later sync is the object-path rewrite trap: Office.js has by then rewritten
 * its path to `getItem(<web-non-round-trippable id>)`, so the next shape throws
 * "InvalidParam passed to GetItem(id)" (code 5010) at the following sync. That is
 * exactly why the fix re-acquires a brand-new getItemAt proxy every batch — a
 * fresh handle is always inside its own window — and why HOLDING one across
 * batches fails. When poisoned, nothing lands: the queued shapes are detached,
 * not pushed to the real slide's `created`.
 *
 * The old fake could not express this at all: it returned the live slide from
 * getItemAt, so a held handle was as good as a fresh one.
 */
function freshWindowedHandle(real: FakeSlide) {
  return windowedHandle(
    real,
    () =>
      new Error(
        'InvalidParam passed to GetItem(id) | code=5010 | debugInfo={"errorLocation":"SlideCollection.getItem"}',
      ),
  );
}

/**
 * The same window, on a handle taken by ID.
 *
 * Unconditional, like the `getItemAt` version, because the host was asked
 * directly and answered. It began as a fault — the evidence then fitted two
 * explanations, and the fake asserting one of them would have been the mistake
 * this file exists to stop. So the probe asked all three ways of naming a new
 * slide apart, and PowerPoint on the web (build a609c9c) said:
 * `shape-add-fresh-slide-proxy` **yes**, `shape-add-positional-slide-proxy`
 * **yes**, `shape-add-held-slide-proxy` **threw**. It is the holding that
 * fails, not the id and not the slide's newness.
 *
 * PowerPoint on the web on 2026-08-04, twice over. The host probe resolved the
 * scratch slide by id, synced, and every one of the eight questions that then
 * wrote through that handle failed; the six that resolved a handle of their own
 * were all answered. The self-test's visibility check failed the same way, with
 * the host naming the call: `GeneralException`, `errorLocation:
 * SlideCollection.getItem`, `statement: var slide = slides.getItem(...);
 * slide.getImageAsBase64(...)` — on a slide whose liveness check had passed
 * through that very proxy one sync earlier.
 *
 * Which is `SlideThunk`'s rule exactly, and the fake could only express it for
 * `getItemAt`: `getItemOrNullObject` handed back the live slide, durable, so
 * holding one across a sync was free here and fatal there.
 */
function expiringSlideHandle(real: FakeSlide) {
  return windowedHandle(
    real,
    () =>
      new Error(
        'GeneralException | code=GeneralException | debugInfo={"code":"GeneralException","message":"GeneralException","errorLocation":"SlideCollection.getItem"}',
      ),
  );
}

/** A slide handle that is only good inside the sync it was acquired in. */
/** Held-handle refusals so far — see `faults.heldSlideProxyRelentsAfter`. */
let heldProxyRefusals = 0;

function windowedHandle(real: FakeSlide, makeError: () => Error) {
  const acquiredSync = trips.syncs;
  // Valid only until the next sync moves past the window it was acquired in.
  const ok = () => {
    if (trips.syncs <= acquiredSync) return true;
    // A host that stops refusing stale handles partway through a run — see
    // `faults.heldSlideProxyRelentsAfter`. Off by default, so every existing
    // test keeps the consistent host it was written against.
    if (faults.heldSlideProxyRelentsAfter !== null) {
      if (heldProxyRefusals >= faults.heldSlideProxyRelentsAfter) return true;
      heldProxyRefusals += 1;
    }
    pendingHostError = makeError();
    return false;
  };
  return {
    id: real.id,
    isNullObject: false,
    load() {},
    delete: () => real.delete(),
    tags: real.tags,
    // Forwarded, and gated like every other member: `new-slide-layout-readable`
    // reads this through exactly such a handle, and a handle that simply did
    // not have a layout would answer "unreadable" — a claim that the host
    // refuses, which is a different thing from a fake that models nothing.
    get layout() {
      return ok() ? real.layout : undefined;
    },
    // The host's own rasteriser is reached through a slide handle like any
    // other call, so a spent window refuses that too — which is exactly how a
    // real host answered the self-test's visibility check.
    getImageAsBase64: (options?: { width?: number }) => (ok() ? real.getImageAsBase64(options) : { value: "" }),
    shapes: {
      // Lazy — evaluating this at handle-creation would fire the .items
      // getter's syncCreated refresh, keeping every shape perpetually fresh
      // and hiding the stale-proxy trap the fake exists to model.
      //
      // Each item remembers WHICH HANDLE produced it. A shape proxy carries its
      // parent's object path, so a member read off a handle whose window has
      // since closed is refused for its parent's sake however young it is —
      // which is what a real host spelled out, listing the statements it would
      // not run: `var slide = slides.getItem(...); var shapes1 = slide.shapes;
      // var shape = shapes1.getItem(...)` ← 5010, in the same batch as a
      // perfectly good `getItemOrNullObject` handle. The members were never too
      // old; their parent was.
      get items() {
        return real.shapes.items.map((sh: FakeShape) => {
          (sh as FakeShape & { parentWindowOk?: () => boolean }).parentWindowOk = ok;
          return sh;
        });
      },
      load() {},
      addGeometricShape: (geo: string, box: FakeShape["box"]) =>
        ok() ? real.shapes.addGeometricShape(geo, box) : makeShape("geometric", geo, box),
      addLine: (kind: string, box: FakeShape["box"]) =>
        ok() ? real.shapes.addLine(kind, box) : makeShape("line", undefined, box),
      addTextBox: (text: string, box: FakeShape["box"]) => {
        if (ok()) return real.shapes.addTextBox(text, box);
        const s = makeShape("text", undefined, box);
        s.text = text;
        return s;
      },
      addGroup: (items: FakeShape[]) =>
        ok() ? real.shapes.addGroup(items) : makeShape("group", undefined, { left: 0, top: 0, width: 0, height: 0 }),
      getItemOrNullObject: (id: string) => real.shapes.getItemOrNullObject(id),
      // FORWARDED LIKE `getItemOrNullObject`, and added for the same reason it
      // was: the windowed proxy is a SUBSET of the collection's surface, and a
      // method missing from the subset throws "not a function" — which a probe
      // records as the HOST refusing, not the double being incomplete.
      //
      // `shapes-by-index-vs-items` hit exactly that: it answered `threw |
      // shapes.getCount is not a function`, which would have gone into the
      // archive as a fact about PowerPoint. The real host answers
      // `getcount-populates-same-sync` yes in 158 of 158 rounds.
      getCount: () => real.shapes.getCount(),
      getItemAt: (index: number) => real.shapes.getItemAt(index),
    },
  };
}

/**
 * Office round-trips since the last installHost(). Every context.sync() is a
 * trip to PowerPoint and dominates insert latency, so the count is a behaviour
 * worth asserting — see "round-trips do not scale with the chart count".
 */
export const trips = {
  syncs: 0,
  contexts: 0,
  /**
   * Calls to `setSelectedShapes([])` — the empty-array deselect.
   *
   * Counted because the invariant is now that it is NEVER made:
   * office-js#3698 reports it does not clear the selection on the web AND
   * leaves the `PowerPoint.run` promise unresolved. A test can only hold that
   * line by seeing the call, not by inspecting what the call would have done.
   */
  emptyDeselects: 0,
};

/** Proxy objects the renderer released via untrack(), by kind. */
export const untracked = { shapes: 0, tags: 0 };

/**
 * Every slide raster the run asked for, in order, as the payload the fake
 * built: `PNG:<slideId>:shapes=<n>:ink=<n>`.
 *
 * Recorded because WHICH slide a caller rasterises is the whole question on
 * PowerPoint web. `getImageAsBase64` on a slide added moments earlier has now
 * failed there five distinct ways, the fifth of them killing the tab, and no
 * count of slides can see the difference — the caller that did it also deleted
 * the slide afterwards, so the deck ended the size it started. The payload can:
 * a slide that already carries a chart reports `shapes` above zero.
 */
export const rasterised: string[] = [];

/**
 * Make the Nth context.sync() of the next run throw. Office.js queues commands
 * and only reports their errors at sync — so this, not a throwing addGroup, is
 * how a host actually refuses something. 0 = never.
 */

/** Like faults.failSyncOn but a SET of sync indices — models a stall that persists across
 * the retry (fail both attempts' syncs), so a slide is truly lost, not recovered. */
export const failSyncsOn = new Set<number>();

/** Make the next N slides.add() calls no-ops — models PowerPoint web silently
 * dropping a slide-add under load, the corruption the self-check must catch. */

/** Deck indices whose shapes.getCount() reports 0 on readback — models a slide
 * that committed but came back BLANK (its shapes detached), the silent partial
 * the on-slide readback must catch. */
export const blankReadbackAt = new Set<number>();

/** When true, shapes.getCount() throws — models the blank readback faulting, so
 * the report must come back blanksRead:false rather than an empty "no blanks". */

/** Sync indices that STALL — the promise resolves after `stallSyncDelayMs`, but
 * withTimeout (shortened via _setBatchTimeoutForTest) fires first and abandons
 * it. Models the real-host bug: sync at 60s, shapes on the slide, no error to
 * blame. When the abandoned promise later settles, withTimeout records
 * `lastLateSync = "…SUCCEEDED after Ns"` — the signal PR 1 gates on. */
export const stallSyncOn = new Set<number>();
const stallSyncDelayMs = 40;

/** When true, addGroup silently drops the group if any member proxy is more
 * than one sync stale — the exact web-host behavior that made multi-batch
 * charts land ungrouped. See PR 3's re-fetch fix.  */

/** When true, `tags.add` rejects a proxy more than one sync old — see makeShape. */

/** Make the next N addGroup calls THROW — a host that declines to group even
 * though the API is there. Decremented per call, so refusing 1 leaves a chart
 * rendered-but-ungrouped and still lets the fresh-context rescue group it. */

/**
 * True once any shape on this host has taken a picture fill.
 *
 * The later arming time for `faults.slideReadsEmpty`, kept as state rather than
 * as a count because the interesting thing about the real host's refusal is
 * WHEN it started, not how many reads it swallowed.
 */
let pictureLanded = false;

/** Whether a shape selection is standing — see `faults.stallDrawAfterSelect`. */
let selectionStanding = false;

/** Whether `faults.slideReadsEmpty` is armed AND its moment has arrived. */
function slideReadsEmptyNow(): boolean {
  if (faults.slideReadsEmpty === "now") return true;
  return faults.slideReadsEmpty === "after-a-picture" && pictureLanded;
}

/** Shape-collection reads that come back EMPTY without throwing — see `items`. */
let lastShapeLoad = "";

/**
 * The property list of the most recent `shapes.load(...)`.
 *
 * Exposed because a request's SHAPE is a fact about cost, and cost is the only
 * reason the deck scan's shape inventory is opt-in: `items/name` is a per-shape
 * string deck-wide, and the callers that sweep every slide on a live web host
 * must keep paying for exactly what they paid for before. Nothing else can check
 * that — a fake shape hands back its name whether or not anyone asked.
 */
export const lastShapeLoadSpec = (): string => lastShapeLoad;

/**
 * Slide ids whose next shape-collection read finds the load UNANSWERED.
 *
 * One step past `hollowReads`, and observed as its limit: the web host answers
 * a collection read short, and sometimes not at all. Office.js then leaves the
 * collection unpopulated, and `.items` throws `PropertyNotLoaded` at
 * `ShapeCollection.items` rather than coming back empty — which is a very
 * different thing for a caller to be told, and the reason a deck-wide chart
 * scan used to die on one silent slide out of thirty-eight.
 *
 * Per slide id, and consumed on use, so a test can silence exactly one slide
 * without touching the fake's own internal reads of the same collection.
 */
export const unansweredShapeReads = new Set<string>();

/**
 * Shape/slide ids whose `getItemOrNullObject` proxy is never populated, however
 * it is loaded — so `isNullObject` stays unreadable.
 *
 * The single-object twin of `unansweredShapeReads`, and the state a caller has
 * the hardest time reasoning about: the object was asked for, the sync landed,
 * and the host said nothing either way. It is neither "there" nor "gone", and
 * code that treats it as one of the two either deletes something it cannot see
 * or refuses to touch something that is plainly there. Only the first of those
 * is unrecoverable, which is why `isLive` answers false.
 */
export const unansweredNullChecks = new Set<string>();

/** When true, `fill.setImage` throws on access — a pre-1.8 host, where the
 * method does not exist at all. Drives the picture-insert fall-through. */

/** Install a fake PowerPoint global whose run() drives the mocked context.
 * `supported(version)` models the host's requirement-set support (default: all)
 * — pass a predicate to simulate e.g. PowerPoint on the web lacking grouping. */
export function installHost(
  slides: FakeSlide[],
  selectedShapes: FakeShape[] = [],
  selectedSlideArg = slides[0],
  supported: (version: string) => boolean = () => true,
) {
  // The slide count as of the last COMMITTED sync. getCount() reports THIS, not
  // the live array — so an add() queued in the current batch is invisible to a
  // getCount() in the SAME batch, exactly as PowerPoint web behaves. A getCount
  // result resolves at the next sync to the count from before that sync's adds.
  let selectedSlide = selectedSlideArg;
  let committedCount = slides.length;
  /**
   * Advance the committed count. One helper rather than three assignments so
   * the shortfall fault below has a single place to reason about.
   */
  const advanceCommitted = () => {
    committedCount = slides.length;
  };
  /**
   * What a resolving `getCount` should report — see `faults.slideCountShortBy`.
   *
   * Floored at zero: a fake answering a negative count would fail callers for a
   * reason no host has ever produced.
   */
  const reportedCount = () => {
    const cap = faults.slideCountCeiling;
    // Only a read the ceiling actually REDUCES spends an allowance.
    if (cap !== null && faults.slideCountCeilingBinds > 0 && committedCount > cap) {
      faults.slideCountCeilingBinds--;
      return cap;
    }
    return committedCount;
  };
  /** Slides `slides.add()` created here — see `faults.newSlideResolvesTimes`. */
  const addedSlideIds = new Set<string>();
  /** How many times each of those has been asked for by id. */
  const addedSlideLookups = new Map<string, number>();
  /**
   * Which ADD each live slide id came from — see `faults.newSlideRefusedForFirst`.
   *
   * Keyed off a monotonic counter rather than the set's size, because ids here
   * are minted from `slides.length` and a scratch slide that is added, refused
   * and deleted frees its id for the next add to mint again. `addScratchSlide`
   * does exactly that on every refusal, so the same id can be the first add and
   * the twentieth — and an order taken from the set's size stayed 0 forever,
   * which made a window that closes look like one that never does.
   */
  let slideAddSeq = 0;
  const addedSlideOrder = new Map<string, number>();
  /**
   * Whether an added slide's lease on being resolvable has run out.
   *
   * Only added slides, and only while the fault is armed: a deck's original
   * slides are durable on every host anyone has met, and it is the freshly
   * made ones a real host lost.
   */
  /**
   * An added slide's id SETTLES to a different one after it has been used.
   *
   * The fake could not be this host on the point that costs the most. Round 253
   * measured a probe run holding `4123571114#123571113` while the deck listed
   * the same slide as `256#2587447327` — not a renumbered neighbour, a
   * different id space — so `getItemOrNullObject(<held id>)` answers
   * `isNullObject: true` while the slide sits in the deck. That is 61 of 64
   * scratch replacements in a round, and a clean-up that swept 107 slides out of
   * a deck of one because delete-by-id recovered none of them.
   *
   * `renumbersOnAdd` is the neighbouring fault and models the OTHER half — an
   * existing slide renumbered when a new one appends. Neither could produce a
   * held id going stale on the slide it names, which is the state the positional
   * re-acquire exists for.
   *
   * Counted in resolutions rather than time so a test can be exact: the id works
   * for the first N by-id lookups and is a different id afterwards.
   */
  const settleAddedSlideId = (id: string): void => {
    if (faults.newSlideIdSettlesAfter === null) return;
    if (!addedSlideIds.has(id)) return;
    const used = addedSlideSettleLookups.get(id) ?? 0;
    addedSlideSettleLookups.set(id, used + 1);
    if (used + 1 < faults.newSlideIdSettlesAfter) return;
    const slide = slides.find((sl) => sl.id === id);
    if (!slide || settledSlideIds.has(id)) return;
    settledSlideIds.add(id);
    // A DIFFERENT SPACE, not an adjacent number, because that is what the host
    // does and a near-miss id would let a buggy caller succeed by accident.
    slide.id = `settled-${settleSeq++}`;
    /**
     * THE SETTLED ID IS NOT AN "ADDED" ID ANY MORE, and that is the whole
     * point of settling.
     *
     * It used to go straight back into `addedSlideIds`, so every fault keyed on
     * that set — `refuseGetItemOnNewSlide` above all — went on refusing the
     * slide under its NEW name too. That models a slide no caller can ever
     * reach by any id, which is not a host: PowerPoint refuses the id it handed
     * you at add time and resolves the one the deck settles on. A correct fix
     * failed against it, which is the wrong direction for a fake to be wrong in.
     */
    /**
     * AND IT IS MARKED SETTLED, or the slide is renamed on every lookup for
     * ever.
     *
     * `settledSlideIds` was keyed on the id being replaced, while the
     * REPLACEMENT went into `addedSlideIds` — so the next by-id lookup found it
     * eligible again and renamed it to `settled-2`, then `settled-3`. A caller
     * doing the right thing (re-read the deck, test whether the id resolves,
     * try again) could never catch up, because the id moved every time it
     * looked. That is not what the host does: a slide settles ONCE.
     *
     * Unused until 2026-09-03, which is why it went unnoticed — the fault was
     * built and never armed. The first test to arm it was the one for
     * `addSlideForChart`, and it failed against a correct fix.
     */
    settledSlideIds.add(slide.id);
    addedSlideOrder.set(slide.id, addedSlideOrder.get(id) ?? 0);
  };

  const addedSlideSettleLookups = new Map<string, number>();
  const settledSlideIds = new Set<string>();
  let settleSeq = 1;
  /**
   * A test's own replacement for `slideMasters`, when it installs one.
   *
   * `slideMasters` became a getter so `faults.twoMasterDeck` could vary how
   * many masters a deck has, and that silently broke assignment — the test
   * that swaps in a throwing collection to check the no-masters fallback got
   * "Cannot set property ... which has only a getter". The setter keeps both
   * working: the fault chooses the shape, an explicit assignment overrides it.
   */
  let mastersOverride: unknown;

  const newSlideLeaseSpent = (id: string): boolean => {
    if (!addedSlideIds.has(id)) return false;
    // The window: the first N added slides are refused outright, whatever their
    // lookup count, and everything added after the window resolves normally.
    // Checked before the lease so the two faults compose rather than one hiding
    // the other — a run can be inside the window AND spending a lease.
    if (faults.newSlideRefusedForFirst > 0) {
      const order = addedSlideOrder.get(id) ?? 0;
      if (order < faults.newSlideRefusedForFirst) return true;
    }
    const lease = faults.newSlideResolvesTimes;
    if (lease === null) return false;
    const used = (addedSlideLookups.get(id) ?? 0) + 1;
    addedSlideLookups.set(id, used);
    return used > lease;
  };
  /** Decks handed to insertSlidesFromBase64 and not yet resolved by a sync. */
  const pendingDecks: string[] = [];
  /** Reasons the NEXT sync must reject, queued by the call that poisoned it. */
  const pendingSyncFailures: string[] = [];
  findShape = (id) => {
    for (const sl of slides) {
      const hit = sl.created.find((sh: FakeShape) => !sh.deleted && sh.id === id);
      if (hit) return hit;
    }
    return undefined;
  };
  deckRemove = (s) => {
    const i = slides.findIndex((x) => x.id === s.id);
    if (i >= 0) slides.splice(i, 1);
  };
  const pendingCounts: { value: number }[] = [];
  const context = {
    presentation: {
      slides: {
        items: slides,
        load() {},
        /**
         * By id — and single-sync when the slide was added in this session.
         *
         * This handed back the live slide, durable, and the comment beside it
         * said "by id, the reference is always durable". That is the kindest
         * the fake was at the one call a real host is harshest about: every
         * held-handle failure PowerPoint web has reported names
         * `errorLocation: SlideCollection.getItem`, and the last one listed the
         * statement it refused — `var slide = slides.getItem(...)` — in the same
         * batch as a `getItemOrNullObject` handle it was perfectly happy with.
         *
         * `getItemOrNullObject` and `getItemAt` have both been windowed for
         * added slides for a while. `getItem` being the exception is backwards,
         * and it is why `insertSceneIntoSlide` could hold one proxy for a whole
         * multi-batch draw with nothing here to notice.
         */
        getItem: (id: string) => {
          const found = slides.find((s) => s.id === id)!;
          /**
           * SETTLING IS COUNTED HERE TOO, because a caller probing with
           * `getItem` is asking the same by-id question `getItemOrNullObject`
           * asks and the host does not care which name the caller used.
           * Counting it only on the lenient call meant a caller that correctly
           * probes with the HARSH one could never advance the settle, so the id
           * never moved and the fake modelled a host that does not exist.
           */
          settleAddedSlideId(id);
          // Windowed only when ARMED — see `faults.newSlideGetItemExpires`.
          if (found && faults.refuseGetItemOnNewSlide && addedSlideIds.has(id)) {
            throw new Error(
              "GeneralException | errorLocation: SlideCollection.getItem | " +
                "the host will not name a slide this run added",
            );
          }
          return found && faults.newSlideGetItemExpires && addedSlideIds.has(id)
            ? (expiringSlideHandle(found) as unknown as FakeSlide)
            : found;
        },
        // Real Office.js hands back a null OBJECT for an unknown id — it does
        // not throw and does not return undefined. A fake that returns
        // undefined would make `slide.isNullObject` a TypeError instead of the
        // false it should be, and hide the very case this models. By id, the
        // reference is always durable.
        //
        // Through `nullObjectProxy` for the same reason the shape collection
        // goes through it: `isNullObject` is populated by the sync a `load()`
        // enrolled the proxy in, so reading it off an unloaded proxy throws.
        // Answering it unconditionally is what let the tag version of this
        // mistake ship — every caller here happens to load first today, and
        // this is what keeps that true.
        getItemOrNullObject: (id: string) => {
          // Settle FIRST, so the lookup that spends the last use is the one that
          // still works and the next one meets the new id. A caller that never
          // looks a slide up never sees it settle, which is the host's own shape.
          settleAddedSlideId(id);
          /**
           * THE HOST HAS ALREADY LOST IT — see
           * `faults.slideVanishesInsteadOfDeleting`. Dropped here rather than at
           * `delete()`, because at 4:3 the slide is gone BEFORE anyone asks to
           * remove it: the lookup finds nothing, `isLive` is false, the
           * positional fallback cannot match an unlisted id, and
           * `deleteSlideById` returns false about a slide that is not there.
           * Only an ADDED slide vanishes; a deck's original slides do not.
           */
          if (faults.slideVanishesInsteadOfDeleting && addedSlideIds.has(id)) {
            const at = slides.findIndex((s) => s.id === id);
            if (at >= 0) slides.splice(at, 1);
          }
          const found = slides.find((s) => s.id === id);
          const live = found && !newSlideLeaseSpent(id) ? found : undefined;
          // A freshly-added slide's by-id handle is single-sync when the fault
          // is armed — see `expiringSlideHandle`. Only added slides: a deck's
          // original ids round-trip through `getItem` on every host anyone has
          // met, which is why editing a chart in place always worked.
          return nullObjectProxy(
            live && addedSlideIds.has(id) ? (expiringSlideHandle(live) as unknown as FakeSlide) : live,
          );
        },
        // A pre-existing slide's handle is durable; a freshly-added one's is only
        // good within the sync it was acquired in (see freshWindowedHandle), so
        // HOLDING one across the render's batches is the bug the fix avoids by
        // re-acquiring fresh each batch.
        getItemAt: (i: number) => {
          const s = slides[i];
          if (!s) return s;
          if (i >= contextBaseCount) return freshWindowedHandle(s);
          // A durable slide the host reports empty on readback (see blankReadbackAt).
          if (blankReadbackAt.has(i)) return { ...s, shapes: { ...s.shapes, getCount: () => ({ value: 0 }) } };
          return s;
        },
        // Resolves at the NEXT sync to the committed count (from before that
        // sync's adds), never to slides.length now — see committedCount.
        getCount: () => {
          const result = { value: committedCount };
          pendingCounts.push(result);
          return result;
        },
        add: (options?: { layoutId?: string; slideMasterId?: string }) => {
          addedWithLayout.push(options?.layoutId);
          addedWithMaster.push(options?.slideMasterId);
          /**
           * THE DOCUMENTED RULE, AND THE ONE THIS FAKE COULD NOT BREAK.
           *
           * `AddSlideOptions` says a `layoutId` sent without a `slideMasterId`
           * "needs to be available for the default Slide Master ... Otherwise,
           * an error will be thrown", and that default is THE PREVIOUS SLIDE'S
           * master. On a one-master deck it can never mismatch — and this fake
           * had exactly one master, so it accepted the pairing forever and
           * every test passed.
           *
           * The archive paid for that. Bucketed by the `layouts-readable`
           * probe over 220 rounds: a one-master deck crashed 9 of 182 rounds
           * (5%); a two-master deck 4 of 4 at 16:9 and 27 of 30 at 4:3. Same
           * add-in, same builds. PowerPoint's own ULS names it
           * `errorLocalChangeLostSingleUser`: it discards the revision and
           * rolls the deck back, which is why a slide is listed twice and then
           * gone, and why `getItemAt` then lands past the end.
           *
           * Modelled rather than assumed away, and armed only on a deck with
           * more than one master, because that is the only state in which the
           * real host can refuse.
           */
          if (
            faults.twoMasterDeck &&
            options?.layoutId &&
            !options.slideMasterId &&
            !LAYOUTS_VALID_WITHOUT_A_MASTER.has(options.layoutId)
          ) {
            throw new Error(
              "GeneralException | code=GeneralException | the layout is not available for the default slide master",
            );
          }
          if (faults.swallowAdds > 0) {
            faults.swallowAdds--; // the host dropped this add — no slide appears
            return;
          }
          const made = makeSlide(`slide-${slides.length + 1}`);
          addedSlideOrder.set(made.rawId, slideAddSeq++);
          addedSlideIds.add(made.rawId);
          // Appending renumbers an EXISTING slide — see `faults.renumbersOnAdd`.
          // Done before the push so the new slide is untouched and stays last,
          // which is the whole basis on which `addScratchSlide` claims it.
          if (faults.renumbersOnAdd && slides.length) {
            const victim = slides[0];
            victim.id = `${victim.id}-renumbered-${slideAddSeq}`;
          }
          if (faults.addsAtFront) slides.unshift(made);
          else slides.push(made);
          // ONE add, TWO slides — see `faults.addLandsExtra`. The real host does
          // this: across the archive, every `scratch slide landed but could not
          // be named` event has `after - before === 2`, ten for ten. The fake
          // could only ever land one, so the branch that counts them was tested
          // against a host that never produced the state it exists for.
          for (let extra = 0; extra < faults.addLandsExtra; extra++) {
            const twin = makeSlide(`slide-${slides.length + 1}`);
            addedSlideOrder.set(twin.rawId, slideAddSeq++);
            addedSlideIds.add(twin.rawId);
            slides.push(twin);
          }
        },
      },
      /**
       * PowerPointApi 1.10's direct slide-size read.
       *
       * Load-gated like everything else here: `slideWidth` off an unloaded
       * PageSetup is PropertyNotLoaded on a real host, and a fake that answered
       * anyway would accept a `slideSize()` that forgot to load it — which is
       * the one mistake this whole class of proxy keeps producing.
       */
      pageSetup: (() => {
        let readable = false;
        const need = (prop: string) => {
          if (!readable) throw new Error(`PropertyNotLoaded: The property '${prop}' is not available.`);
        };
        return {
          load(_p?: string) {
            pendingTagLoads.push(() => {
              readable = true;
            });
          },
          get slideWidth() {
            need("slideWidth");
            return hostSlideSize.cx / 12700;
          },
          get slideHeight() {
            need("slideHeight");
            return hostSlideSize.cy / 12700;
          },
        };
      })(),
      /**
       * PowerPointApi 1.8 shape bindings — the happy path, and only that.
       *
       * A binding is the one handle to a shape that never goes through
       * `ShapeCollection.getItem(id)`: it is made from the live proxy in the
       * batch that created the shape, and asked for later by a key the caller
       * chose. That is exactly the route around the 5010 refusals this fake
       * models everywhere else, which is why `binding-names-shape-later` asks
       * about it and why the answer here is `yes`.
       *
       * `yes` is a claim about a host that BEHAVES, in keeping with the rest of
       * this happy path — not a claim that PowerPoint on the web behaves. It has
       * never been asked there, so the entry sits in `PENDING_QUESTIONS` and the
       * first real sheet to answer it will either agree or land as a divergence.
       * Nothing in `src/` uses bindings yet; do not build on this until it has.
       *
       * Deliberately NOT wired into the stale-proxy faults. Whether a binding
       * outlives the refusals is the open question, and a fake that answered it
       * either way would be inventing the finding the probe exists to get.
       */
      bindings: (() => {
        const store = new Map<string, FakeShape>();
        /** A binding that names nothing — the shape it held is gone. */
        const noShape = () => ({ load(_p?: string) {}, id: undefined as unknown as string });
        return {
          add: (shape: FakeShape, _bindingType: string, id: string) => {
            if (faults.refuseBindings === "call")
              throw new Error("UnexpectedError | errorLocation: BindingCollection.add");
            // Taken, and the BATCH is what fails — see `faults.refuseBindings`.
            // Queued like any other command, so the rejection lands on the sync
            // rather than here, which is the whole distinction the two words
            // `add-threw` and `commit-threw` are there to keep apart.
            if (faults.refuseBindings === "sync") pendingSyncFailures.push("UnexpectedError");
            // "If the provided ID is already being used by a binding, the
            // existing binding will be overwritten" — so a repeat key replaces
            // rather than accumulates, and a probe can use a fixed one.
            store.set(id, shape);
            return { id };
          },
          // THROUGH A FRESH HANDLE, which is what "deliberately NOT wired into
          // the stale-proxy faults" above requires and what the store alone did
          // not deliver. The stored object is the shape's ORIGINAL handle, and
          // its `tags` writer closes over that handle's age — so handing it back
          // put every binding write straight into `strictTags` and answered the
          // open question `no` by accident, which is precisely the finding this
          // fake is not allowed to invent.
          //
          // Found on 2026-08-16, the first time `src/` used bindings at all —
          // the comment above says not to build on this until it did.
          getItemOrNullObject: (id: string) => ({
            getShape: () => {
              const sh = store.get(id);
              return sh ? freshHandle(sh) : noShape();
            },
          }),
          getItem: (id: string) => ({
            getShape: () => {
              const sh = store.get(id);
              return sh ? freshHandle(sh) : noShape();
            },
          }),
          getCount: () => ({ value: store.size }),
          load() {},
          get items() {
            return [...store.keys()].map((id) => ({ id }));
          },
        };
      })(),
      /**
       * The presentation-scoped custom XML parts, which is where a deck keeps
       * its STYLE (`src/core/deck-style.ts`).
       *
       * Modelled with the two behaviours the readers actually depend on, both
       * of which the typings state and neither of which is obvious:
       * `getOnlyItemOrNullObject` answers a NULL OBJECT when the namespace holds
       * nothing, and it returns a NULL OBJECT when the namespace holds more than
       * one — it does not pick, and it does not refuse either. `writeDeckStyle`
       * deletes every part in the namespace before adding, and that second rule
       * is why.
       *
       * This said "it REFUSES ... it does not pick" and the fake threw to match.
       * `@types/office-js`: `getOnlyItem` raises `GeneralException` on "no items
       * or more than one"; `getOnlyItemOrNullObject`, which is the one we call,
       * "returns null" otherwise. Modelling the throw made a two-part deck look
       * loud when the real thing is silent.
       */
      customXmlParts: (() => {
        const parts: { id: string; namespaceUri: string; xml: string }[] = [];
        let nextId = 1;
        const partHandle = (p: { id: string; namespaceUri: string; xml: string }) => ({
          id: p.id,
          namespaceUri: p.namespaceUri,
          isNullObject: false,
          load() {},
          getXml: () => ({
            get value() {
              return p.xml;
            },
          }),
          setXml(xml: string) {
            p.xml = xml;
          },
          delete() {
            const i = parts.indexOf(p);
            if (i >= 0) parts.splice(i, 1);
          },
        });
        /** The namespace a part's XML declares, which is how a real host files it. */
        const nsOf = (xml: string) => /xmlns="([^"]*)"/.exec(xml)?.[1] ?? "";
        return {
          add(xml: string) {
            if (faults.refuseCustomXmlWrites) throw new Error("GeneralException | the host refused a custom XML part");
            const p = { id: `xml-${nextId++}`, namespaceUri: nsOf(xml), xml };
            parts.push(p);
            return partHandle(p);
          },
          getByNamespace(ns: string) {
            const inNs = () => {
              const hits = parts.filter((p) => p.namespaceUri === ns);
              // A second part nobody asked for, to reach the state the write
              // path is built to prevent.
              return faults.duplicateCustomXmlPart && hits.length === 1 ? [hits[0], { ...hits[0], id: "dup" }] : hits;
            };
            return {
              load() {},
              get items() {
                return inNs().map(partHandle);
              },
              getCount: () => {
                // THE FIRST CUSTOM-XML CALL, WHATEVER IT IS. The real fault this
                // models does not care which method it lands on: seven rounds
                // show the first call after a pane loads failing and the second
                // working, and it has been a `getOnlyItemOrNullObject` hang and
                // a `getCount` that resolved unloaded on different builds. A
                // once-fault that fired only on the item read could not express
                // that, so the warm-up — whose whole job is to BE the first call
                // — sailed through it and the test failed for the wrong reason.
                if (faults.refuseCustomXmlReadsOnce) {
                  faults.refuseCustomXmlReadsOnce = false;
                  throw new Error("GeneralException | the host did not answer the custom XML read");
                }
                return { value: inNs().length };
              },
              getOnlyItemOrNullObject() {
                // THE HOST THAT HAS THE API AND WILL NOT ANSWER THE READ. Round
                // 089 recorded exactly that on the day #583 merged — `reading
                // the deck's style` hung for its whole 90s budget — and nothing
                // here could express it, so the one caller that AWAITS the read
                // could not be tested against a failure at all. `failSyncOn`
                // reaches it only by guessing a global sync index, which the
                // comment on that fault already calls fragile.
                if (faults.refuseCustomXmlReadsOnce) {
                  faults.refuseCustomXmlReadsOnce = false;
                  throw new Error("GeneralException | the host did not answer the custom XML read");
                }
                if (faults.refuseCustomXmlReads)
                  throw new Error("GeneralException | the host did not answer the custom XML read");
                const hits = inNs();
                // MORE THAN ONE IS A NULL OBJECT, NOT A THROW — checked against
                // `@types/office-js`, which is as close to the contract as this
                // repo can get without a deck that carries two:
                //
                //   getOnlyItem()             no items OR more than one -> GeneralException
                //   getOnlyItemOrNullObject() exactly one -> it; "Otherwise ... returns null"
                //
                // This fake threw `InvalidArgument` for the two-part case and
                // said in a comment that the real API "refuses rather than
                // choosing". It refuses in `getOnlyItem`; the OrNullObject
                // variant we actually call does not. The difference is not
                // cosmetic: on a real host a deck holding two style parts reads
                // as CARRYING NO STYLE rather than as a failed read, which is
                // the quieter and worse of the two outcomes.
                //
                // `writeDeckStyle`'s delete-then-add is still right, and for a
                // better reason than the one written down: a second part does
                // not announce itself at all, it silently unbrands the deck.
                //
                // NOT VERIFIED ON THE REAL HOST — no round has ever put two
                // parts in this namespace, and the typings are documentation.
                // Recorded as the documented contract, not as an observation.
                return hits.length === 1
                  ? partHandle(hits[0])
                  : { isNullObject: true, load() {}, getXml: () => ({ value: undefined }) };
              },
            };
          },
          /** The raw store, for a test that wants to see what the deck now holds. */
          get items() {
            return parts.map(partHandle);
          },
          load() {},
        };
      })(),
      // A real deck's master carries several layouts; only one is blank, and
      // its NAME is localised — which is why the renderer matches on type.
      set slideMasters(v: unknown) {
        mastersOverride = v;
      },
      get slideMasters() {
        if (mastersOverride) return mastersOverride as { items: unknown[]; load(): void };
        const first = {
          id: "master-1",
          layouts: {
            items: [
              { id: "layout-title", name: "Titeldias", type: "titleSlide" },
              { id: "layout-blank", name: "Tom", type: "blank" },
              { id: "layout-content", name: "Titel og indhold", type: "object" },
            ],
          },
        };
        /**
         * A SECOND MASTER, AND THE DECK'S SLIDES BELONG TO IT — see
         * `faults.twoMasterDeck`. It is second in `items` on purpose: code that
         * walks the masters looking for a blank layout finds the FIRST one's,
         * which is exactly the layout that is not legal for this deck's default
         * master unless the add names its master too.
         */
        const items = faults.twoMasterDeck
          ? [first, { id: "master-2", layouts: { items: [{ id: "m2-layout-blank", name: "Tom", type: "blank" }] } }]
          : [first];
        return { items, load() {} };
      },
      getSelectedSlides: () => ({
        getItemAt: () =>
          faults.selectedSlideIdAs === null
            ? selectedSlide
            : // Same slide, different id — which is the whole of #2474.
              new Proxy(selectedSlide as object, {
                get: (t, k) => (k === "id" ? faults.selectedSlideIdAs : Reflect.get(t, k)),
              }),
      }),
      getSelectedShapes: () => ({
        // The LIVE selection, not the array installHost was handed — otherwise
        // `setSelectedShapes` would write somewhere nothing reads.
        get items() {
          return selectionRef;
        },
        // Same contract as a slide's own collection: an `items/…` selector
        // populates the shapes it names. See ShapeCollection.load.
        load(p?: string) {
          noteSelectionCall();
          if (faults.selectionReadThrows) throw new TypeError("e.load is not a function");
          if (p?.includes("items/")) for (const s of selectionRef) s.load();
        },
      }),
      setSelectedSlides: (ids: string[]) => {
        noteSelectionCall();
        const found = slides.find((sl) => ids.includes(sl.id));
        if (found) selectedSlide = found;
      },
      /**
       * The whole-deck insert — and the path a real run actually takes.
       *
       * The fake could not model it at all, so every test of the default path
       * either skipped or ran against the shape-by-shape one instead. Here the
       * bytes are really decoded (by the same auditor `npm run verify-deck`
       * uses, so the fake cannot disagree with the tool) and each slide in the
       * file becomes a slide in the deck, carrying its slot tag and a single
       * `PowerChart` shape holding the config — which is what the generator
       * genuinely writes.
       *
       * Queued, not immediate: Office.js resolves this at the next sync, and a
       * fake that appended straight away would hide every ordering bug the
       * queue creates.
       */
      insertSlidesFromBase64: (b64: string) => {
        if (faults.swallowDecks > 0) {
          faults.swallowDecks--; // taken, acknowledged, and never landed
          return;
        }
        pendingDecks.push(b64);
      },
    },
    /**
     * `passThroughValue` is part of the contract, not decoration:
     * `sync<T>(passThroughValue?: T): Promise<T>` resolves with what it was
     * handed, and `PowerPoint.run` relies on that to carry a batch's result
     * out. This fake took no argument and returned undefined, so nothing in
     * the suite could see a wrapper that dropped it — and one did, on
     * 2026-09-02, breaking every host read in the build while 3,831 tests
     * stayed green. See the `run` stub below.
     */
    sync: async <T>(passThroughValue?: T): Promise<T | undefined> => {
      /**
       * THE BATCH'S CLOSING AUTO-SYNC CARRIES THE VALUE AND NOTHING ELSE.
       *
       * Office.js ends every `run` with `.then(r => ctx.sync(r))`, and modelling
       * that is what lets this fake catch a wrapper that drops the argument —
       * the defect that let a completely broken add-in pass 3,831 tests.
       *
       * But it is also a real sync, and counting it renumbered every
       * fault-injection index in the suite: ten tests that pin wedge and stall
       * behaviour to "the Nth sync" broke at once. Those indices are load-bearing
       * and re-numbering them by hand would silently move what each test asserts.
       *
       * So this call is exercised but not CHARGED. Anything wrapping `sync` still
       * runs — that is the whole point, and a wrapper that swallows
       * `passThroughValue` still fails here — while the fault machinery, the trip
       * counters and the commit bookkeeping stay on the syncs the code under test
       * actually asked for. The fake stays faithful in the dimension the bug
       * lives in and stable in the dimension the suite indexes by.
       */
      if (inClosingAutoSync) return passThroughValue;
      trips.syncs++;
      syncsInContext++;
      // Charged BEFORE anything else this sync does, and unconditionally, so
      // the cost lands on every sync including the ones that go on to fail.
      // A host that only slows down when it succeeds is not a host anyone has
      // seen.
      // Longer than any batch budget a test sets, so the draw gives up — see
      // `faults.stallDrawAfterSelect`.
      if (faults.stallDrawAfterSelect && selectionStanding) await new Promise((r) => setTimeout(r, 200));
      if (faults.syncCostMs) {
        const ms = faults.syncCostMs({ syncsInContext, syncsTotal: trips.syncs });
        if (ms > 0) await new Promise((r) => setTimeout(r, ms));
      }
      const decks = pendingDecks.splice(0);
      for (const b64 of decks) {
        const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
        const { rows } = (await readDeckBytes(bytes)) as { rows: DeckRow[] };
        for (const row of rows) {
          const added = makeSlide(`slide-${slides.length + 1}`);
          if (row.slot !== null) {
            added.tagStore.set(DEMO_SLOT_TAG_KEY, JSON.stringify({ i: row.slot, title: row.title, run: row.run }));
          }
          // One group named PowerChart, tagged, holding as many children as
          // the file really holds — what the generator writes.
          //
          // The child COUNT is the load-bearing part. A readback measures a
          // chart by how many shapes are inside its group, and a fake that put
          // one shape there made every generated chart read as wreckage: the
          // repair pass then had no healthy copy to prefer and queued no
          // duplicates, so the duplicate-slot scenario reported a failure that
          // was the fake's, not the code's.
          const shape = added.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 10, height: 10 });
          shape.name = "PowerChart";
          shape.grouped = Array.from({ length: Math.max(1, row.chartShapes) }, (_, k) =>
            makeShape("geometric", "rectangle", { left: k, top: 0, width: 1, height: 1 }),
          );
          if (row.configJson) shape.tagStore.set(CHART_TAG_KEY, row.configJson);
          added.pending.length = 0;
          slides.push(added);
        }
      }
      // The slides are on the deck by now and the caller will never be told.
      // See `faults.deckInsertNeverAnswers`.
      //
      // `committedCount` is advanced FIRST, and that is the whole fidelity of
      // this fault. It normally moves at the end of a successful sync, which
      // models shapes landing when their batch commits — but a deck insert is
      // not a batch of shapes, and office-js#1650 is explicit that "the slide
      // still gets added successfully" while the promise hangs. A fake that
      // left the count behind would report the slides as absent to the very
      // re-read that exists to find them, and the test would be measuring the
      // fake's bookkeeping rather than the host behaviour it stands for.
      if (decks.length && faults.deckInsertNeverAnswers) {
        faults.deckInsertNeverAnswers = false;
        faults.selectionReadThrows = false;
        committedCount = slides.length;
        await new Promise(() => {});
      }
      // Commit or discard the shapes each slide has queued since the last
      // sync. On success, mark them "committed" (leave in `created`, drop
      // from `pending`). On failure, remove them from `created` — the real
      // host discards a batch of queued commands whose sync rejects, and the
      // PR-8 partial-landed gate reads back only the committed count.
      const commit = () => {
        for (const s of slides) s.pending.length = 0;
        // Queued tag loads become readable here and nowhere else — see
        // makeTagProxy. A load only takes effect when its sync lands.
        for (const apply of pendingTagLoads) apply();
        pendingTagLoads.length = 0;
        for (const apply of pendingShapeLoads) apply();
        pendingShapeLoads.length = 0;
      };
      // Exports resolve on the sync that asked for them, same as a real
      // ClientResult. Awaited before `commit()` so the value is there the
      // moment the caller's `await context.sync()` returns.
      if (pendingExports.length) {
        for (const e of pendingExports) e.result.value = await e.build();
        pendingExports.length = 0;
      }
      const discard = () => {
        for (const s of slides) {
          for (const p of s.pending) {
            const i = s.created.indexOf(p);
            if (i >= 0) s.created.splice(i, 1);
          }
          s.pending.length = 0;
        }
        // A sync that rejects populates nothing: the loads it carried are lost,
        // and reading those proxies still throws.
        pendingTagLoads.length = 0;
        pendingShapeLoads.length = 0;
      };
      if (trips.syncs === faults.failSyncOn || failSyncsOn.has(trips.syncs)) {
        discard();
        throw new Error("host refused a queued command");
      }
      // A batch poisoned by one of the commands IN it, rather than by its
      // number. `faults.failSyncOn` needs a sync index, which is the fragile
      // magic number this repo has been bitten by; this fires on the sync that
      // actually carries the offending call, wherever that lands.
      if (pendingSyncFailures.length) {
        const why = pendingSyncFailures.shift()!;
        pendingSyncFailures.length = 0;
        discard();
        throw new Error(why);
      }
      // A sync carrying a selection call on a wedged host. Never settles —
      // neither resolve nor reject, which is what the web host did and is why
      // the wait had to be bounded rather than caught. The caller's
      // `withTimeout` is the only thing that ends it.
      if (wedgeThisSync) {
        wedgeThisSync = false;
        await new Promise(() => {});
      }
      // Verbatim from the 2026-08-07 run log, because the code that recovers
      // from this reads `errorLocation` nowhere and a future reader will.
      if (refuseThisSync) {
        refuseThisSync = false;
        // DISCARD, like every other rejecting path above. This one did not, and
        // the omission made a live production defect untestable: the fake
        // populated the loads the refused batch carried, so `isNullObject` read
        // back cleanly on a proxy the host had just refused to resolve, and the
        // update sailed on as though nothing had happened. A test written
        // against the real failure passed with the fix for it deleted.
        //
        // "A sync that rejects populates nothing" is stated three branches above
        // as the rule. It is the rule here too.
        discard();
        throw Object.assign(new Error("InvalidParam passed to GetItem(id)"), {
          code: "5010",
          debugInfo: {
            code: "5010",
            errorLocation: "ShapeCollection.getItem",
            statement: "var shape = shapes.getItem(...);",
          },
        });
      }
      if (faults.wedgeAfterSyncs !== null && trips.syncs > faults.wedgeAfterSyncs) await new Promise(() => {});
      if (stallSyncOn.has(trips.syncs)) {
        // Sleep past withTimeout's deadline, then settle successfully. The
        // queued shapes commit at settle time — same as real Office.js where
        // a slow sync eventually reports success and the shapes are on the
        // slide by then.
        await new Promise((r) => setTimeout(r, stallSyncDelayMs));
        for (const r of pendingCounts) r.value = reportedCount();
        pendingCounts.length = 0;
        advanceCommitted();
        commit();
        return passThroughValue;
      }
      // A queued-command failure (e.g. drawing on a poisoned getItemAt handle)
      // surfaces here, at the sync, exactly as Office.js reports it.
      if (pendingHostError) {
        const err = pendingHostError;
        pendingHostError = null;
        discard();
        throw err;
      }
      // Each getCount from this batch resolves to the PRE-batch committed count,
      // then this batch's adds become visible to the NEXT sync's getCount.
      for (const r of pendingCounts) r.value = reportedCount();
      pendingCounts.length = 0;
      advanceCommitted();
      commit();
      return passThroughValue;
    },
  };
  trips.syncs = 0;
  trips.emptyDeselects = 0;
  trips.contexts = 0;
  untracked.shapes = 0;
  untracked.tags = 0;
  rasterised.length = 0;
  // Proxies from a previous test's host must not become readable inside this
  // one's first sync.
  pendingTagLoads.length = 0;
  pendingShapeLoads.length = 0;
  pendingExports.length = 0;
  // Back to 16:9 unless a test says otherwise — a leaked 4:3 would silently
  // change placement for every test after it.
  hostSlideSize.cx = 12192000;
  hostSlideSize.cy = 6858000;
  pendingHostError = null;
  // failSyncOn was the one fault installHost did not reset, so a test that set
  // it leaked into every later test in the file and every test in every file
  // that ran after. It only ever looked safe because each user reset it by
  // hand — a convention, one `return` away from being forgotten.
  faults.failSyncOn = 0;
  faults.swallowAdds = 0;
  failSyncsOn.clear();
  stallSyncOn.clear();
  faults.strictGroup = false;
  faults.strictTags = false;
  faults.refuseCustomXmlWrites = false;
  faults.refuseCustomXmlReads = false;
  faults.refuseCustomXmlReadsOnce = false;
  faults.duplicateCustomXmlPart = false;
  faults.refuseTagWritesOnResolvedProxy = false;
  faults.refuseTagWrites = 0;
  faults.refuseIdLeftTopLoads = 0;
  faults.refuseShapeById = false;
  faults.addLandsExtra = 0;
  faults.shapeCountLag = 0;
  faults.shapeCountLagBy = 0;
  faults.strictIdLoads = false;
  faults.refuseShapeIdLoads = 0;
  refuseThisSync = false;
  faults.refuseGroups = 0;
  faults.hollowReads = 0;
  faults.hollowReadsAfter = null;
  faults.readsMissing = 0;
  faults.slideReadsEmpty = null;
  faults.slideIdUnreadableBeforeFirstSync = false;
  faults.slideIdNeverReadable = false;
  faults.stallDrawAfterSelect = false;
  selectionStanding = false;
  pictureLanded = false;
  faults.selectedSlideIdAs = null;
  faults.hollowNameReads = 0;
  lastShapeLoad = "";
  faults.refusePictureFill = false;
  blankReadbackAt.clear();
  faults.faultShapeGetCount = false;
  faults.faultShapeCollectionLoad = false;
  faults.strictShapeReads = false;
  faults.noIdInCreatingSync = false;
  faults.unansweredTagLoads = 0;
  faults.webIgnoresDeselect = false;
  faults.selectionIgnoresIds = false;
  faults.selectionWedgesHost = false;
  selectionWedged = false;
  wedgeThisSync = false;
  faults.constantSlideImage = false;
  faults.emptySlideImage = false;
  faults.slideImageGate = null;
  faults.refuseSlideDelete = false;
  faults.slideVanishesInsteadOfDeleting = false;
  faults.twoMasterDeck = false;
  faults.deckInsertNeverAnswers = false;
  faults.selectionReadThrows = false;
  faults.tagsUndefinedOn = 0;
  faults.newSlideResolvesTimes = null;
  faults.heldSlideProxyRelentsAfter = null;
  heldProxyRefusals = 0;
  faults.newSlideRefusedForFirst = 0;
  faults.refuseBindings = null;
  faults.renumbersOnAdd = false;
  faults.addsAtFront = false;
  faults.newSlideGetItemExpires = false;
  faults.refuseGetItemOnNewSlide = false;
  faults.textBoxDeletesSelection = false;
  faults.refuseShapeAdds = false;
  faults.wedgeAfterSyncs = null;
  faults.syncCostMs = null;
  syncsInContext = 0;
  // The live shape selection starts as installHost was told, and is mutated
  // from there by Slide.setSelectedShapes.
  selectionRef.length = 0;
  selectionRef.push(...selectedShapes);
  unansweredShapeReads.clear();
  unansweredNullChecks.clear();
  faults.swallowDecks = 0;
  addedWithLayout.length = 0;
  addedWithMaster.length = 0;
  vi.stubGlobal("PowerPoint", {
    run: async <T>(cb: (ctx: typeof context) => Promise<T>) => {
      trips.contexts++;
      // Slides present at the start of this context are "existing"; anything
      // add()ed past here is fresh, and its getItemAt handle is window-limited.
      contextBaseCount = slides.length;
      syncsInContext = 0;
      /**
       * THE FINAL AUTO-SYNC, which this fake did not have and which is how the
       * real host returns a value at all.
       *
       * Office.js's `_runCommon` ends every batch with
       * `.then(runBodyResult => ctx.sync(runBodyResult))`, and
       * `ClientRequestContext.sync(passThroughValue)` resolves with exactly
       * what it was handed. So the batch's result rides out through one last
       * sync rather than being returned directly.
       *
       * This fake used to `return cb(context)` — correct-looking, and blind to
       * the entire class of bug where something wraps or replaces `sync`. On
       * 2026-09-02 a sync-counting wrapper did exactly that, took no
       * parameters, and made EVERY `PowerPoint.run` in the build resolve
       * undefined. The whole suite stayed green: 3,831 tests, every one of them
       * passing against a fake that could not express the defect. It took two
       * real-host rounds, an abandoned experiment and a wrong diagnosis of a
       * "damaged deck" to find it.
       *
       * A fake that cannot fail the way the real thing fails is not a test
       * double, it is a second implementation that agrees with you.
       *
       * MEASURED, NOT ASSUMED, and not yet adopted. Routing the return value
       * through `context.sync(await cb(context))` makes 123 tests in
       * `office-render` go red the moment the argument is dropped again — so
       * the faithful version earns its place. It also adds one sync to every
       * run, which renumbers every fault index in the suite and breaks 11
       * tests that pin wedge and stall behaviour to a specific sync. Those
       * encode real host behaviour and deserve a careful pass rather than a
       * hurried one at the end of a long day, so the upgrade is filed and the
       * contract is guarded directly instead — see the pass-through test in
       * `test/office-render.test.ts`.
       *
       * ADOPTED 2026-09-03. The deferral above was the right call at the end of
       * a long day and the wrong one to leave standing: a fake that cannot fail
       * the way the real host fails is a second implementation that agrees with
       * you, and this one let a completely broken add-in pass 3,831 tests.
       *
       * The lookup happens INSIDE this callback, after `cb` resolves, because
       * that is where Office.js does it — `.then(runBodyResult => ctx.sync(...))`.
       * Written as `context.sync(await cb(context))` it would resolve the member
       * reference BEFORE the argument, capture the unpatched `sync`, and be
       * blind to exactly the wrapper class it exists to catch. That is not
       * hypothetical: the first draft of the guard test made that mistake and
       * passed against the bug.
       */
      const runBodyResult = await cb(context);
      inClosingAutoSync = true;
      try {
        return await context.sync(runBodyResult);
      } finally {
        inClosingAutoSync = false;
      }
    },
    // Real Office.js exposes all 177 presets. A plain object listing only the
    // ones in use today hands back `undefined` for any other name, and the
    // renderer then records a shape with no geometry while this suite still
    // passes green — a test that asserts nothing about the shape it drew.
    // The Proxy makes that impossible: reaching for a preset this stub has not
    // been told about throws instead of returning undefined.
    GeometricShapeType: new Proxy(
      {
        rectangle: "rectangle",
        ellipse: "ellipse",
        triangle: "triangle",
        chevron: "chevron",
        homePlate: "homePlate",
        lineInverse: "lineInverse",
        diamond: "diamond",
        plus: "plus",
        // The hex tile map's tile. The Proxy below did exactly what it was
        // built for on the way in: NOTHING rendered a hex tilemap through this
        // renderer, so moving its tiles from polygons to symbols shipped green
        // while the very first `addGeometricShape` for one would have thrown.
        // `draws every hex tile as a fillable hexagon` closes that.
        hexagon: "hexagon",
      } as Record<string, string>,
      {
        get(target, prop: string) {
          // See faults.presetMissing: a real host's enum answers `undefined`
          // for a name it lacks rather than throwing, and that is the branch
          // the renderer's fallback exists for.
          if (faults.presetMissing && prop === faults.presetMissing) return undefined;
          if (!(prop in target)) throw new Error(`office stub: unknown GeometricShapeType "${String(prop)}"`);
          return target[prop];
        },
      },
    ),
    SlideLayoutType: { blank: "blank", titleSlide: "titleSlide", object: "object" },
    ConnectorType: { straight: "straight" },
    ShapeLineDashStyle: { dash: "dash", roundDot: "roundDot" },
    ShapeAutoSize: { autoSizeNone: "none" },
    TextVerticalAlignment: { top: "top", middle: "middle", bottom: "bottom" },
    ParagraphHorizontalAlignment: { left: "left", center: "center", right: "right" },
  });
  selectionHandlers.length = 0;
  vi.stubGlobal("Office", {
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    context: {
      host: "PowerPoint",
      requirements: { isSetSupported: (_set: string, version: string) => supported(version) },
      /**
       * The Common API selection event — the one the pane's own banner uses,
       * and the only route to the selection that does NOT go through the
       * subsystem a programmatic `setSelectedShapes` wedges.
       *
       * Modelled because there is now a scenario that waits for a real click.
       * Without a way to raise the event from a test, that scenario could only
       * ever be observed timing out, which proves nothing about the half that
       * matters.
       */
      document: {
        addHandlerAsync: (type: string, handler: () => void) => {
          if (type === "documentSelectionChanged") selectionHandlers.push(handler);
        },
        removeHandlerAsync: (type: string, opts?: { handler?: () => void }) => {
          if (type !== "documentSelectionChanged") return;
          if (!opts?.handler) selectionHandlers.length = 0;
          else {
            const i = selectionHandlers.indexOf(opts.handler);
            if (i >= 0) selectionHandlers.splice(i, 1);
          }
        },
      },
    },
  });
  return context;
}

/** Handlers registered for the Common API selection event. */
const selectionHandlers: (() => void)[] = [];

/**
 * A user clicking `shape` on the slide — the event, not a programmatic select.
 *
 * Writes the selection the way the host would and raises the event, so code
 * that LISTENS sees exactly what a real click produces. Deliberately does not
 * go near `setSelectedShapes`: that is the call under suspicion, and a helper
 * that used it would make every listener test depend on it.
 */
export function userClicksShape(shapeId: string | null): void {
  selectionRef.length = 0;
  const shape = shapeId ? findShape?.(shapeId) : undefined;
  if (shape) selectionRef.push(shape);
  for (const h of [...selectionHandlers]) h();
}

/** How many selection handlers are currently registered — leak check. */
export const selectionHandlerCount = (): number => selectionHandlers.length;
