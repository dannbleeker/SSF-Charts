/**
 * An answer sheet: what THIS PowerPoint actually does, question by question.
 *
 * `CLAUDE.md` states the problem plainly — "every Office.js assertion in this
 * repo is against a fake". The fake has grown twenty-odd faults, each added
 * after a real host taught us something, and it is the reason bugs can be found
 * in CI at all. But nobody has ever checked whether it is FAITHFUL. Its faults
 * are things we learned; its happy path is a set of things we assumed. Where an
 * assumption is wrong, every test resting on it is confidently wrong with it,
 * and no amount of green tells us so.
 *
 * This is not a call log. A recording of what the add-in did is what the trace
 * already gives, and joining two of them proves nothing. This is a fixed list
 * of QUESTIONS — each one a behaviour `powerpoint.ts` actually depends on —
 * asked identically of a real host and of the fake, so the two answer sheets
 * can be diffed. Every divergence is exactly one of two things:
 *
 * - the fake lies, and the tests built on it are worth less than they look, or
 * - the host does something we did not know, which is a finding in itself.
 *
 * Both are worth a round. Neither is visible any other way.
 *
 * **Answers are small and comparable on purpose.** `"yes"`, `"no"`, `"threw"`,
 * `"silent"`, a count — never free text, because the whole point is a machine
 * diff between two hosts. Timings and error messages ride along as `detail`
 * and are deliberately NOT compared: they differ between any two runs and
 * would bury the signal.
 *
 * **Nothing here touches the user's deck.** Every probe works on one scratch
 * slide, added at the start and removed at the end. A probe that would damage a
 * real slide does not belong here.
 */
import {
  addScratchSlide,
  scratchSlideIdByPosition,
  deadlinesFired,
  deckSlideIds,
  deleteTrailingSlides,
  deleteSlideById,
  isTimeout,
  requirementSets,
  ScratchSlideUnavailable,
  withProbeContext,
  type ProbeContext,
} from "./powerpoint";
import { trace, traceElapsed } from "../core/trace";
import { SYMBOL_PRESET } from "../core/geometry";
// @ts-expect-error — a plain .mjs table with no types, and deliberately the
// SAME one the CLI and the CI gate read. Two copies of it is how a claim
// quietly stops matching its check.
import { FAKE_BASELINE, KNOWN_DIVERGENCES, UNSTABLE_ANSWERS, diffAnswers } from "../../scripts/host-baseline.mjs";

/** One question and what this host said. */
export interface HostAnswer {
  /** Stable key. Renaming one breaks the diff against older sheets — don't. */
  id: string;
  /** What was asked, in words, for whoever reads the sheet. */
  question: string;
  /**
   * The comparable fact, from a small vocabulary. This is the ONLY field the
   * diff looks at.
   */
  answer: string;
  /** Wall clock. Context for a human; never diffed. */
  ms: number;
  /** Error text, counts, whatever helps read a surprising answer. Never diffed. */
  detail?: string;
  /**
   * Every answer this run got for this question, in the order it got them.
   *
   * A question asked ONCE has been sampled, not answered — this project has
   * written that sentence in three places and still had to learn each unstable
   * answer the expensive way, across ten rounds, because one round could only
   * ever produce one sample. `UNSTABLE_ANSWERS` in `scripts/host-baseline.mjs`
   * is the artefact of that: a table maintained by hand from whatever rounds
   * happened to be open.
   *
   * A run asks each question up to `PROBE_PASSES` times, spread across the run
   * rather than back to back, so the samples land in different minutes of the
   * host's life. `stable` is then a fact one sheet can state on its own.
   *
   * The ROW stays one per id, and `answer` keeps its meaning — the diff, the
   * contract gate, `host-diff` and `host-baseline` all read that field and none
   * of them needed to change.
   */
  samples?: ProbeSample[];
  /**
   * Whether every real sample agreed. Undefined when the run got fewer than two.
   *
   * `false` here is the finding `UNSTABLE_ANSWERS` exists to record, said by
   * the sheet itself instead of by a human comparing rounds.
   */
  stable?: boolean;
  /**
   * How many attempts reached the question, when some did not.
   *
   * Present ONLY when `asked < of`, so its presence is the finding: this answer
   * came from fewer attempts than were made, and the rest never got to ask.
   * `{ asked: 1, of: 4 }` is a fact the host said once and refused three times —
   * still the best answer available, and not the same thing as one four attempts
   * agreed on. See `supportOf`.
   */
  support?: { asked: number; of: number };
}

/** One ask of one question, and the state the host was in when it answered. */
export interface ProbeSample {
  answer: string;
  /** Which pass over the question list this came from — 1, 2, 3. */
  pass: number;
  /**
   * Milliseconds on the RUN LOG's clock, so a sample can be lined up against
   * the trace lines around it.
   *
   * It used to be milliseconds into the probe, which is a different origin —
   * the probe starts several seconds after tracing does (7.9s in the round of
   * 2026-08-11). Nothing in the file said so, so joining the two series, which
   * is how every one of these rounds gets analysed, was silently off by that
   * constant: the trace put pass 2 at 41.6s and the samples put its first
   * answer at 34.4s, i.e. a sample apparently arriving before the pass that
   * produced it. Falls back to probe-relative only when tracing is off, in
   * which case the file carries no trace to be misjoined with.
   */
  atMs: number;
  /**
   * What the run had already seen the host doing when this was asked.
   *
   * Derived from what the run has ALREADY observed — no extra call, so asking
   * cannot itself disturb the thing being measured. Recorded on every sample
   * and not only on the odd ones: a value written down only when something goes
   * wrong cannot be compared against anything, which is a mistake this project
   * has now made four times with four different fields.
   */
  regime: HostRegime;
  /**
   * The scratch slide this question ran on. See `ScratchState` for why it is
   * four words rather than an age or a count, and for the round that made it
   * the live candidate once `regime` was eliminated.
   */
  scratch: ScratchState;
  /**
   * WHICH ROUTE got the slide this sample was taken on. See
   * `ScratchAcquisition` — the state above says what the slide IS, this says
   * where it came from, and round 254 is the round where those two answers
   * stopped agreeing.
   *
   * Recorded on EVERY sample, like its two neighbours, and for the reason their
   * docstrings each give: a value written down only when something looks wrong
   * cannot be compared against anything.
   */
  acquired: ScratchAcquisition;
}

/**
 * How the host was behaving, in the only terms this run can honestly claim.
 *
 * `collection-refused` is the documented regime flip — the shape collection
 * stops answering, and everything downstream of it degrades. `slide-trouble` is
 * the other known window: freshly added slides stop resolving for ~15 seconds
 * at a time. Neither is a guess about the host's insides; both are things this
 * run watched happen.
 */
export type HostRegime = "healthy" | "slide-trouble" | "collection-refused" | "unknown";

/**
 * Which scratch slide a question was asked on — the state `regime` is not.
 *
 * Round 17 (`cd3b60c`) established that `shape-add-held-slide-proxy` and its
 * partner agree perfectly at every instant (seven pairs, seven agreements) and
 * yet answer both ways inside ONE regime. So there is a definite state, it
 * changes during a run, and the regime stamp does not name it. That round also
 * killed the two easy candidates from its own numbers — `threw` at 16.3s, `yes`
 * at 33.9s, `threw` at 55.6s is non-monotonic, so neither elapsed time nor the
 * pass number is the variable.
 *
 * The candidate left is the one `UNSTABLE_ANSWERS` already fingers: this
 * question WRECKS its own scratch slide every time it is asked, so pass 1 meets
 * a deck with no scratch history and later passes do not.
 *
 * CATEGORICAL on purpose. A duration or a counter would give every sample its
 * own value, and "every value maps to one answer" is then true for any data at
 * all — the unfalsifiable shape `host-regimes.mjs` reports as `untested`. Four
 * closed values can actually disagree with themselves.
 *
 * Derived from what the run has already watched itself do — no extra host call,
 * so asking cannot disturb the thing being measured — and recorded on EVERY
 * sample, not only the odd ones, which is the mistake this project has now made
 * four times with four different fields.
 */
export type ScratchState = "first-slide" | "fresh-slide" | "reused-slide" | "no-slide";

/**
 * HOW the run came to hold this sample's scratch slide — the ROUTE, where
 * `ScratchState` above is the resulting STATE. They are different facts and
 * round 254 is what proves it.
 *
 * `77f9ca4` (2026-08-25, "Re-acquire the scratch slide instead of buying
 * another one") changed acquisition from "buy a new one" to "ask the deck for
 * the one we already have", and nothing in a round file said which route a
 * sample took. The only record was `trace.entries`: 5,122 "re-acquired the
 * scratch slide by position" lines, every one in a round numbered 254 or above
 * and none in the 229 below it. Reconstructing the split from those lines is
 * how the round-254 wall was found, weeks late, and it cost an entire analysis
 * to do what this field now does for free. `docs/BACKLOG.md` §"ROUND 254 IS OUR
 * OWN COMMIT" is the write-up.
 *
 * SIX SITES TAKE A SLIDE and `takeScratch`'s `sameSlide` flag names only one of
 * them. It is a parameter read once inside an `if` and never stored, so it
 * cannot be the record: it separates the re-acquire from everything else and
 * says nothing about WHICH of the other five, nor about there being no slide at
 * all. Hence a value per route rather than a boolean.
 *
 * `unknown` IS NOT A ROUTE, it is the absence of one, and the distinction is
 * this file's own rule — see `archivableDeck`'s `undisclosed` in
 * `scripts/round.mjs` for the same argument. A sample archived before this
 * field existed carries no `acquired` at all; a sample that carries `unknown`
 * was taken by a path that reached `record` without going through
 * `takeScratch`. A reader can tell those apart, and collapsing them is exactly
 * the mistake the `regime` and `ScratchState` docstrings above are each
 * counting instances of. Do not add to that tally here; point at it.
 */
export type ScratchAcquisition = "first-add" | "re-acquired" | "recovered" | "replaced" | "no-slide" | "unknown";

/**
 * How long an observation still describes the host.
 *
 * This host's states come and go in windows of roughly fifteen seconds — a
 * freshly added slide stops resolving, then starts again — so a refusal from a
 * minute ago is history, not a description of now. Twenty seconds is one such
 * window plus a little, and it is the number that turns this field from a latch
 * into a measurement.
 *
 * The fifteen is folklore that has never been measured, and the first attempt
 * to measure it came out SHORTER. Grouping every probe attempt in rounds 23 and
 * 26 into runs of consecutive starvation: 17 runs and 14 runs, longest 7.6s and
 * 8.9s, most of them spanning 0s. Clustering is real but mild — 14 runs observed
 * against ~17 expected if each attempt failed independently — so the twenty here
 * is comfortably above anything yet seen, and is left alone.
 *
 * That measurement does NOT cover the whole claim: it counts starvation among
 * probe attempts, not every host state, and the four slots that starve on every
 * attempt in both rounds are burnt slots rather than windows (see
 * `Probe.burnsTheSlide`) — removing them is what any honest window figure would
 * have to do first. Treat the fifteen as unverified rather than wrong.
 */
export const REGIME_WINDOW_MS = 20_000;

/**
 * What the host was doing, from the most RECENT thing this run watched it do.
 *
 * The first version of this was two sticky booleans: once any shape question
 * answered a refusal, every later sample was stamped `collection-refused` for
 * the rest of the run. The first real round showed what that is worth — the
 * flag latched 8.9 seconds into a 110-second probe and 55 of 65 samples came
 * back carrying the same value. A field that nearly every sample shares cannot
 * separate anything, which is the exact mistake this project has now recorded
 * against `idleMs`, `afterAnswering`, the settle's shared label and
 * `listChartsInDeck` — four fields, four times, and this was the fifth.
 *
 * So: timestamps, a window, and an explicit `unknown` for "nothing recent
 * enough to say". Refusal outranks slide trouble because a host that will not
 * list shapes is in the deeper hole, and both outrank `healthy`, which is only
 * claimed on the strength of something that actually worked.
 */
export function regimeFrom(
  o: {
    at: number;
    lastRefusalAt?: number;
    lastSlideTroubleAt?: number;
    lastGoodAt?: number;
    /**
     * The last time a question that ASKS THE SHAPE COLLECTION got a real answer
     * out of it — not merely the last time any question answered, which is what
     * `lastGoodAt` is and which nearly every question sets.
     */
    lastCollectionGoodAt?: number;
  },
  windowMs = REGIME_WINDOW_MS,
): HostRegime {
  const recent = (t?: number) => t !== undefined && o.at - t <= windowMs;
  // A refusal inside the window used to win outright, and the docstring above
  // says the opposite — "from the most RECENT thing this run watched it do".
  // The code was priority-ordered, so one refusal painted every sample for the
  // next twenty seconds however well the collection answered in between.
  //
  // Round 17 is what that costs: the collection ANSWERED 14 of the 28 times it
  // was asked, interleaved with the refusals throughout, and 88% of that
  // round's samples still read `collection-refused` — worse than the 85% sticky
  // flag this function was written to replace, and past the failure criterion
  // its own docstring sets. Simulated over that round, clearing the refusal
  // when the collection has answered SINCE takes it to 72%.
  //
  // Deliberately NOT most-recent-wins across all three: `lastGoodAt` is set by
  // almost every question, so recency alone would saturate on `healthy` instead
  // and this comment would be describing the same bug in the other direction.
  // Only a later COLLECTION answer clears a collection refusal.
  const answeredSince =
    o.lastCollectionGoodAt !== undefined && o.lastRefusalAt !== undefined && o.lastCollectionGoodAt > o.lastRefusalAt;
  if (recent(o.lastRefusalAt) && !answeredSince) return "collection-refused";
  if (recent(o.lastSlideTroubleAt)) return "slide-trouble";
  if (recent(o.lastGoodAt)) return "healthy";
  return "unknown";
}

/**
 * The sheet's format marker, written into every archived round.
 *
 * BOTH SPELLINGS ARE VALID ON READ, and that is not tidiness — 39 of the last
 * 40 archived rounds carry `powerchart-host-answers`, and `host-baseline.mjs`
 * and `host-diff.mjs` select on it. Renaming the write without widening the read
 * would make every tool reject 257 rounds of evidence, which is the archive this
 * whole project is built on.
 *
 * New sheets are stamped `ssf-charts-host-answers`. Old ones keep working
 * forever; nothing rewrites them, because rewriting an archive to match a rename
 * is how a record stops being a record.
 */
export type HostAnswerSheetKind = "ssf-charts-host-answers" | "powerchart-host-answers";

/** The marker written on sheets this build produces. */
export const HOST_ANSWERS_KIND: HostAnswerSheetKind = "ssf-charts-host-answers";

/** Does this look like an answer sheet, under either name it has had? */
export function isHostAnswersKind(kind: unknown): boolean {
  return kind === "ssf-charts-host-answers" || kind === "powerchart-host-answers";
}

/** A complete sheet, plus what produced it. */
export interface HostAnswerSheet {
  kind: HostAnswerSheetKind;
  /**
   * The host's answer to the cheapest possible call, before any question.
   *
   * Rounds 24, 25 and 29 wedged, and each was read as the probe walking into
   * trouble around question six. Round 29 disproved that: the question that has
   * answered `yes` in every round on record went SILENT at question four, so the
   * host was already unwell before the run asked it anything. Nothing in the
   * sheet could say so, which is why every hypothesis about the wedge — time of
   * day, deck size, build, session age — was fitted to the wrong moment.
   *
   * A number, deliberately, not a verdict: three healthy rounds and three wedged
   * ones are not enough to set a threshold, and a guessed one would be quoted as
   * fact by the next reader.
   */
  opened?: { ms: number; answered: boolean };
  /** Which host answered — a real PowerPoint, or the fake in CI. */
  source: string;
  build: string;
  /** Requirement sets this host admits to, since half the answers depend on it. */
  requirementSets: string[];
  answers: HostAnswer[];
}

/**
 * How long a single question may take before it counts as unanswered.
 *
 * Short. A wedged host must cost one question, not the sheet — the sheet's
 * value is that it comes back complete even from a host that is misbehaving,
 * which is exactly when it is most worth having.
 */
let PROBE_BUDGET_MS = 8_000;

/**
 * How long the probe will wait for a scratch slide to land.
 *
 * The per-question budget above bounds the QUESTION and nothing else, and the
 * slide a question needs is acquired outside it — so on 2026-08-10
 * `shape-resolve-held-slide-proxy` took 95.6 seconds against an eight-second
 * budget, because `addScratchSlide` defaults to `READBACK_TIMEOUT_MS`: ninety
 * seconds, sized for a twenty-slide repair page. A bound that something else
 * can walk around is not a bound, and the sheet's whole promise is that a
 * wedged host costs one question rather than the run.
 *
 * Fifteen because the two outcomes are nowhere near each other. Every scratch
 * add that worked in that run took between 0.21 and 4.0 seconds; every one that
 * failed took the full ninety. Fifteen is nearly four times the slowest success
 * and a sixth of the failure, so it cannot turn a slow add into a lost question
 * — and a probe run makes twenty-odd of these.
 */
let SCRATCH_ADD_BUDGET_MS = 15_000;

/**
 * Test-only: shorten the scratch-add budget.
 *
 * The same reason `_setProbeBudgetForTest` exists one screen up — a bound is
 * only reachable by letting an add actually miss it, and fifteen seconds of
 * that per case is longer than any suite should sit still for.
 */
export function _setScratchAddBudgetForTest(ms: number): void {
  SCRATCH_ADD_BUDGET_MS = ms;
}

/**
 * Test-only: shorten the per-question budget.
 *
 * A wedged host is only reachable in a test by letting a question actually miss
 * its deadline, and at eight seconds a full sheet of those outlasts any sane
 * test timeout. Production never touches this; the default is the eight seconds
 * a real host is given. Restore by passing 8_000 back in.
 */
export function _setProbeBudgetForTest(ms: number): void {
  PROBE_BUDGET_MS = ms;
}

/** What a probe does, given a scratch slide it is free to wreck. */
type Probe = {
  id: string;
  question: string;
  /** Answers, or throws — a throw is recorded as an answer, not a failure. */
  ask: (ctx: ProbeContext) => Promise<{ answer: string; detail?: string }>;
  /**
   * This question does not need a LIVE scratch slide, so do not spend one on it.
   *
   * `withProbeContext` checks the scratch slide's liveness before handing a
   * context to any probe, and on this host that check is the single most
   * likely thing to fail — a fresh slide's id resolves once and then stops. So
   * a question that never touches the slide inherits a failure that has nothing
   * to do with it, and the sheet records `no-scratch-slide` where a real answer
   * was available all along.
   *
   * That is not hypothetical: the 2026-08-08 `a546897` round lost
   * `untrack-available` this way, at 43 seconds, AFTER the host had recovered
   * and answered four questions in a row. `untrack` is a `typeof` on a proxy —
   * `getItemOrNullObject` builds one without a round trip — so the question was
   * answerable and the sheet said otherwise. A statement about the probe wearing
   * the clothes of a statement about the host is the error this whole file is
   * built to avoid.
   */
  noSlideNeeded?: true;
  /**
   * Asking this question leaves the scratch slide unusable, so the run takes a
   * fresh one before the NEXT question rather than handing on the wreckage.
   *
   * Earned by seven rounds of evidence and one accidental control. The question
   * after `shape-add-held-slide-proxy` has never been answered — not once in
   * seven sheets — and for six of them that was `shape-resolve-held-slide-proxy`,
   * whose 0/6 everything here (this file's own notes included) put down to its
   * reading ids back off setup shapes. Then a reorder for unrelated reasons put
   * `binding-names-shape-later` in that slot instead, and the slot failed
   * exactly the same way, in 397ms, at the liveness check, before the probe's
   * own code ran at all:
   *
   *   round 12-17  #5 shape-add-held-slide-proxy threw | #6 shape-resolve-…  never put
   *   round 18     #5 shape-add-held-slide-proxy threw | #6 binding-names-…  never put
   *
   * Two different questions, one slot, seven for seven. It is the SLOT, and the
   * cause is what the probe above it is for: `shape-add-held-slide-proxy`
   * deliberately writes through a slide proxy the host has stopped honouring,
   * and this host does not forgive that on the slide afterwards.
   *
   * It ANSWERS, so none of the not-asked machinery notices — the replacement
   * paths all trigger on a question that failed, and this one succeeds and
   * poisons the well on its way out.
   *
   * WHAT THE FLAG BUYS, MEASURED — AND THE PREDICTION IT FAILED. Round 27 was
   * run to judge the three carriers added for #8, #16 and #22, against a stated
   * prediction that all three would answer for the first time in eleven rounds.
   * **They did not.** All three came back never-put again. The round is still a
   * clean control for what the flag DOES do:
   *
   *   flagged   shape-resolve-held-slide-proxy   no-scratch-slide -> no-scratch-shape
   *   flagged   tags-add-same-key-twice          no-scratch-slide -> no-scratch-shape
   *   flagged   grouped-child-by-id-from-slide   no-scratch-slide -> no-scratch-shape
   *   UNflagged tag-on-group-survives            no-scratch-slide -> no-scratch-slide
   *
   * (`grouped-child-by-id-from-slide` and `tag-on-group-survives` were RETIRED
   * on 2026-08-21 — production had answered one and routed around the other.
   * The reading above is left as it was taken; it is a record of that round.)
   *
   * Every flagged question changed failure mode; the unflagged control did not;
   * and the trace shows the flag firing 11 times on exactly its four carriers.
   * The slide half is fixed — these questions now GET a live slide. They fail
   * one layer down, and the detail names it: "the host would not read back the
   * shapes' ids", and for the grouped one "the host would not name the members
   * before grouping, so there was no child id to look up".
   *
   * That is not a starvation problem and no amount of slide management moves it.
   * It is a fact this sheet ALREADY records twice — `shape-proxy-survives-one-sync`
   * and `shapes-items-count-honest` both answer `unreadable` — so a question that
   * must name a freshly added shape on a fresh slide is unanswerable AS WRITTEN
   * on this host. Fixing these three means taking this file's own long-standing
   * advice and not reading ids back off their own setup shapes: a child id has to
   * come from a shape this host has already consented to name.
   */
  burnsTheSlide?: true;
  /**
   * A second question, asked in the same run, when this one's answer admits
   * two readings.
   *
   * The repo's rule is "when two explanations fit the evidence, ask — do not
   * reason". Following it has meant me writing the partner question, the owner
   * running the probe again, and a session going by — twice, at a cost of two
   * full sheets. The reasoning was never the expensive part; the ROUND TRIP
   * was. A pair that can be decided in one run should be.
   *
   * `when` is what keeps this honest. An unconditional partner is just another
   * probe and belongs in the list; a follow-up earns its place by being worth
   * asking only in the light of a particular answer, and by SAYING which answer
   * that was — `because` goes into the sheet, so a reader sees the pair rather
   * than two unrelated rows.
   */
  follow?: {
    when: (answer: string) => boolean;
    because: string;
    probe: Probe;
  };
  /**
   * Ask this one again even when the run is short of scratch slides.
   *
   * The shortlist the later passes fall back to under pressure: the questions
   * whose answers this project does NOT yet trust — everything in
   * `UNSTABLE_ANSWERS` (answered differently on different runs of one build)
   * and everything in `PENDING_QUESTIONS` (never answered at all). Marked on
   * the question rather than kept as a list somewhere else, so it cannot drift
   * from the question it describes; `test/host-probe.test.ts` reads those two
   * tables out of `scripts/host-baseline.mjs` and fails if the marks and the
   * tables disagree, which is what keeps them honest in the other direction.
   */
  resample?: true;
};

/**
 * Answers that mean "this question was never put", not "the host said so".
 *
 * None of them is in any probe's own vocabulary, which is the point: a diff
 * that mistook one for an answer would report a host divergence that nobody
 * ever asked about. They were earned the same way, rounds apart — see
 * `ProbeSetupFailed`.
 *
 * EXPORTED because `NEVER_ASKED` in `scripts/host-baseline.mjs` is the same
 * vocabulary, read by the diff tool — and a comment promising the two are kept
 * in step is not a mechanism. `not-asked` was added here and not there, so every
 * question the mute breaker abandons was being compared against the fake and
 * reported as a real-host divergence. `test/host-probe.test.ts` asserts the two
 * sets are equal now.
 */
export const NOT_ASKED = new Set(["no-scratch-slide", "no-scratch-shape", "no-named-slide", "not-asked"]);

/**
 * Why a scratch slide had to be bought, in the vocabulary the archive pools on.
 *
 * ONE FUNCTION BECAUSE TWO COPIES DRIFTED. This expression lived inline at both
 * places that buy a slide, and only the main loop was ever given the
 * `no-scratch-shape` case — so every partner replacement was filed as
 * `unrecorded`: 68 of the last ten rounds' 248, 27% of them, saying nothing
 * about why a slide was bought. Duplication is why they could differ at all.
 *
 * `unrecorded` MEANS UNKNOWN, never "known and unsaid". It is the bucket for the
 * ~17,000 events archived before any cause was carried out of the catch, and
 * putting a cause we DO have into it makes those rounds unreadable — the reading
 * that motivated the main loop's version of this line.
 */
export function scratchReplacementWhy(result: { why?: string; answer: string }): string {
  return result.why ?? (result.answer === "no-scratch-shape" ? "shape-refused" : "unrecorded");
}

/**
 * Answers that mean "the question was put and produced nothing to name".
 *
 * `other` is every probe's catch-all: it is what a question returns when the
 * outcome was none of the ones it knows how to recognise. This file says so
 * itself about the tag-key question — "this host answered `other —
 * value=undefined`: not an opinion about tag keys at all, just the stale handle
 * refusing both writes."
 *
 * IT IS NOT IN `NOT_ASKED`, AND THAT COST A REAL ANSWER TWICE. The row keeps its
 * FIRST real answer and `other` counts as real, so a pass-1 `other` locked the
 * row and a later named answer could not replace it. Across 86 archived rounds
 * `tags-add-same-key-twice` reads `other` in EIGHTY-THREE of them — while the
 * samples of rounds 074 and 091, on two different builds, both carry
 * `overwrites`. The host answered the question twice and the sheet said UNKNOWN.
 *
 * So a NAMED answer now outranks `other`, and nothing outranks a named answer.
 * The original rule's intent is kept — a sheet means today what it meant
 * yesterday, and disagreement is `stable`'s job to report — because `other` was
 * never a meaning to be stable about.
 *
 * ── AND THEN IT HAPPENED AGAIN, IN THE SAME FUNCTION, WITH A DIFFERENT WORD ──
 *
 * `unreadable` is the other half of the failure vocabulary: a probe says it when
 * a value did not come back, or came back the wrong shape, or the read threw
 * inside the probe's own catch. It was never added here, so it locked rows
 * exactly as `other` used to.
 *
 * Measured over all 322 archived rounds and 12,346 rows: 87 rows say
 * `unreadable` while their OWN samples carry a real answer, across eight
 * different questions. `which-end-a-short-read-drops` answered `none` in 11 of
 * them; `shapes-items-count-honest` answered `short-0` in 8; `picture-then-
 * shape-read` and `binding-names-shape-later` each answered a plain `yes`.
 *
 * The sharpest one is `rotation-keeps-the-unrotated-box`. Across its 33 passes
 * the host said `unrotated-box` FOUR times and the contradicting answer NOT
 * ONCE — the rest could not read. That is the question the round harness gained
 * a whole scenario to ask, because if this host placed rotated shapes by their
 * post-rotation box then every diagonal in every line, scatter, radar and violin
 * chart would land wrong. The host answered. The sheet said `unreadable` in all
 * eleven rounds.
 *
 * `unreadable` is not a meaning to be stable about either. It is the absence of
 * a reading, and a reading beats it.
 *
 * ── AND A THIRD TIME, WHICH IS WHEN IT STOPPED BEING A COINCIDENCE ──
 *
 * `silent` is what a probe records when its own sync misses the deadline (see
 * the `isTimeout(err)` branch that emits it). Unlike `threw`, which is the host
 * actively refusing a call and therefore a fact worth keeping, `silent` says
 * only that no reply arrived in time — the host may well answer the same
 * question on the next pass, and repeatedly does.
 *
 * Archive-wide it locked 52 rows whose own samples held a real answer:
 * `binding-names-shape-later` 34 of them, `shape-add-held-slide-proxy-again` 12.
 * SIXTEEN of the 52 are a plain `yes` reported as a timeout — including
 * `shape-add-held-slide-proxy` in round 346, inside the window this was mined
 * from.
 *
 * Three words now, found one at a time, each costing rounds of real answers
 * before anyone noticed. So the vocabulary is pinned by a test rather than by
 * whoever next remembers: see "the failure vocabulary is classified on purpose"
 * in `test/host-probe.test.ts`, which fails when a new non-answer word is
 * emitted without being classified here.
 *
 * ── AND TWO THAT ARE ONLY NON-ANSWERS FOR THE QUESTION THAT ASKS THEM ──
 *
 * `not-a-short-read` and `none-of-ours` belong to `which-end-a-short-read-drops`
 * and mean its precondition never occurred. They were `all` and `none` until
 * 2026-09-01, which could not be listed here: `scratch-slides-returned` answers
 * `all`, `some` and `none` and means every one of them. The SAME WORD IS A REAL
 * ANSWER TO ONE QUESTION AND A SHRUG AT ANOTHER, so the words were made specific
 * before they were made weak. A probe that needs a non-answer of its own should
 * name it for its own question rather than reach for a generic one.
 */
export const UNINFORMATIVE = new Set([
  "other",
  "unreadable",
  "silent",
  "not-a-short-read",
  "none-of-ours",
  // `no-shapes-to-list` belongs to `durable-slide-lists-its-shapes` and means
  // the slide it was asked about was empty — the host answered the call, there
  // was simply nothing to count. Named for its own question exactly as the
  // paragraph above demands, and weak for the reason that paragraph gives: on
  // the 16:9 arm the durable slide is empty in nearly every round, so a strong
  // answer here would lock the row with a non-answer in the arm where the
  // question cannot inform, and the 4:3 arm — where it CAN — would never get to
  // speak. A trigger that cannot fail is not a measurement.
  "no-shapes-to-list",
]);

/** Weak enough to be replaced by a named answer: never asked, or asked and unnameable. */
export function weakAnswer(a: string): boolean {
  return NOT_ASKED.has(a) || UNINFORMATIVE.has(a);
}

/**
 * How much an answer is worth when two of them compete for one row.
 *
 * THREE TIERS, BECAUSE A BOOLEAN COLLAPSED TWO REAL THINGS. Promotion used to
 * ask `weakAnswer(seen) && !weakAnswer(arriving)`, which works only while "weak"
 * has exactly one meaning. The moment `unreadable` joined `UNINFORMATIVE` that
 * stopped being true and a genuine distinction went with it: `no-scratch-shape`
 * means the probe never REACHED its question, `unreadable` means it reached it
 * and could not read the result. The second is strictly more informative than
 * the first, and a re-ask that gets as far as reading must still displace a
 * deferral — `test/host-probe.test.ts` says so, and it is right.
 *
 * So rank rather than flag, and promote only on a strict increase:
 *
 *   0  never asked        the question was not reached
 *   1  asked, unnameable  reached it; `other` or `unreadable` came back
 *   2  a named answer     the host actually said something
 *
 * The original invariant survives at the top: two named answers do not displace
 * each other, so the FIRST named answer is still the row's answer and a sheet
 * means today what it meant yesterday. Disagreement remains `stable`'s job.
 */
export function answerRank(a: string): 0 | 1 | 2 {
  if (NOT_ASKED.has(a)) return 0;
  if (UNINFORMATIVE.has(a)) return 1;
  return 2;
}

/**
 * How many times a run asks each question.
 *
 * Three, and the number is the whole point of the exercise: one sample cannot
 * say whether an answer is stable, two cannot say which of a disagreeing pair is
 * the odd one, and three is where a round starts producing the fact that
 * `UNSTABLE_ANSWERS` was built by hand across ten rounds to produce.
 *
 * Spread across the run, not asked back to back — three asks inside four
 * seconds sample one minute of the host's life three times, which is one
 * sample wearing a hat. Pass 1 is the whole list, exactly as before, so a host
 * that dies early still yields today's sheet and nothing regresses.
 */
export const PROBE_PASSES = 3;

/**
 * Whether a question's samples agree — `undefined` when it has fewer than two
 * real ones to compare.
 *
 * Only REAL answers count. A question put once and refused twice is not
 * unstable; it is a question this run mostly could not ask, and calling that
 * instability would manufacture exactly the noise `UNSTABLE_ANSWERS` warns
 * against. Pure, and exported, so the rule can be checked without a host.
 *
 * TWO TIERS, AND THE SECOND ONE WAS MEASURED BEFORE IT WAS WRITTEN. Once
 * `unreadable` became weak (see `UNINFORMATIVE`), the obvious move was to filter
 * it here too. That is WRONG, and the archive says so: filtering every weak
 * answer changes 1,765 rows, and the overwhelming direction is `true` →
 * `undefined` — rows where all three passes said `unreadable` and the sheet
 * currently reports, correctly, that the host is CONSISTENTLY unreadable.
 * Consistency of refusal is a real fact and worth keeping.
 *
 * So: judge among the strong answers when there are two or more of them, and
 * otherwise fall back to the original rule. Measured over the same 12,346 rows
 * that is 52 changes, every one `false` → `true` — rows that claimed the host
 * contradicted itself when it had not: it answered the same thing every time it
 * answered at all, and merely failed to read on the other passes. A pass that
 * could not read is not a dissenting vote.
 */
export function stabilityOf(samples: ProbeSample[] | undefined): boolean | undefined {
  const answers = (samples ?? []).map((x) => x.answer);
  const strong = answers.filter((a) => !weakAnswer(a));
  if (strong.length >= 2) return new Set(strong).size === 1;
  const real = answers.filter((a) => !NOT_ASKED.has(a));
  return real.length >= 2 ? new Set(real).size === 1 : undefined;
}

/**
 * How many attempts actually reached the question, out of how many were made.
 *
 * `record` promotes the first REAL answer over a never-asked, which is right —
 * two passes that could not ask and a third that did should report what the
 * third found. But the row it produces is then indistinguishable from one every
 * pass agreed on: `answer` says `threw` whether four attempts said it or one
 * did, and `stable` is `undefined` for BOTH "answered once, never repeated" and
 * "never answered at all", because it needs two real samples to say anything.
 *
 * So a single sample becomes an unqualified fact about the host. That matters
 * here more than it would elsewhere, because a `threw` is exactly what a
 * misclassified SETUP failure looks like — the probe never reached its
 * question, the error surfaced anyway, and the never-asked vocabulary that
 * exists to catch that gets overwritten by the one attempt that got it wrong.
 * `KNOWN_DIVERGENCES`, the CI fixture and `host-diff` all read `answer` and none
 * of them can tell the two apart.
 *
 * This does not fix the misclassification — that is a separate and harder
 * problem, and guessing at it has already cost this project a reverted branch.
 * It makes the weight of an answer visible, so a one-of-four does not get
 * quoted as though it were a four-of-four.
 */
export function supportOf(samples: ProbeSample[] | undefined): { asked: number; of: number } {
  const all = samples ?? [];
  return { asked: all.filter((x) => !NOT_ASKED.has(x.answer)).length, of: all.length };
}

/**
 * The support worth RECORDING, or nothing when there is nothing to say.
 *
 * Only when the answer rests on fewer attempts than were made. A row every
 * attempt reached needs no qualifier, and stamping one on all of them puts the
 * interesting case back where it started — in a field nobody reads because it is
 * always present. A row NO attempt reached already says `no-scratch-slide` in
 * `answer`, and `0 of 3` beside it adds nothing.
 *
 * Split out of the loop that used it because a condition written inline there is
 * a condition no test can reach: deleting it left the entire suite green.
 */
export function thinSupport(samples: ProbeSample[] | undefined): { asked: number; of: number } | undefined {
  const { asked, of } = supportOf(samples);
  return of > 1 && asked > 0 && asked < of ? { asked, of } : undefined;
}

/**
 * The share of a pass that may come back never-asked before the later passes
 * stop being worth their scratch slides.
 *
 * Scratch slides are the scarcest thing on this host and the repeats spend
 * three times as many of them. When a pass is already losing a third of its
 * questions to `no-scratch-slide`, asking the settled ones twice more is
 * bidding for slides against the questions that actually need a second sample —
 * so the later passes fall back to the resample shortlist, and the run says so
 * in its trace rather than quietly doing less.
 */
export const PASS_PRESSURE_LIMIT = 1 / 3;

/**
 * Consecutive unanswered questions before the probe stops asking.
 *
 * Three, for the reason the self-test's SICK_LIMIT is three: one is often the
 * finding, and three in a row is a finding about the host rather than about a
 * question.
 */
const PROBE_MUTE_LIMIT = 3;

/**
 * Thrown when a probe could not get the shapes its question is about.
 *
 * The same lesson as `ScratchSlideUnavailable`, one layer down, and it cost a
 * second real answer sheet to learn. That guard covers a scratch slide the host
 * will not resolve; it says nothing about a slide that resolves perfectly and
 * then refuses to take a shape. PowerPoint on the web did exactly that on
 * 2026-08-04: six questions that only read were answered, and all eight that
 * needed a shape failed in setup — `GeneralException` at the add, or a sync
 * that never came back. Every one of the eight was recorded as `"threw"` or
 * `"silent"`, both legitimate answers to those questions, so `host-diff`
 * reported eight host divergences from a sheet that had asked six questions.
 *
 * A probe that never reached its question must not answer it.
 */
class ProbeSetupFailed extends Error {
  constructor(readonly why: string) {
    super(`the probe never got as far as its question: ${why}`);
    this.name = "ProbeSetupFailed";
  }
}

/**
 * Whether the sync that was to deliver a probe's setup shapes is still out.
 *
 * A setup failure that THROWS is caught where it happens; a setup failure that
 * WEDGES is caught by the probe budget, which fires in `ask` — outside any
 * knowledge of what the probe was in the middle of. Three of that real sheet's
 * eight failures were the wedging kind, and `"silent"` is a comparable answer,
 * so they read as host divergences too.
 *
 * Module-level because probes run strictly one at a time; `ask` clears it
 * before every question.
 */
let awaitingSetupShapes = false;

/**
 * A 1x1 transparent PNG — the smallest thing that is genuinely a picture.
 *
 * The probe that asks about office-js#5022 needs a real image and cares nothing
 * about what it looks like. Bare base64, no `data:` prefix, because that is what
 * `fill.setImage` takes.
 */
const ONE_PIXEL_PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

/** A shape's box, in points on the scratch slide. */
type ProbeBox = { left: number; top: number; width: number; height: number };

/**
 * Put shapes on the scratch slide, or admit that this host would not.
 *
 * Through a slide proxy resolved in THIS sync and used before the next one —
 * `ProbeContext.scratch` is a thunk for that reason, and holding what it
 * returns re-opens the trap it exists to close.
 *
 * `load` names a real property to read back off the shapes, and it costs a
 * SECOND sync — add, commit, then ask.
 *
 * It used to be queued in the same batch as the add, to save a round trip. That
 * saving is why five questions on the 2026-08-05 sheet were never put: the five
 * that call `idsOf`, and no others. Perfect correlation, and no other property
 * of those five explains it — batch count does not (`shapes-items-count-honest`
 * and `delete-then-lookup` each span two batches and were answered, while three
 * of the five span two and failed). The implied host fact is narrow and
 * plausible: **this PowerPoint will not name a shape in the batch that created
 * it.**
 *
 * Asking after the commit costs one round trip per id-needing question and
 * makes them answerable at all. It ages the returned proxies by a sync, which
 * is free here: every probe that asks for ids goes on to work through
 * `probeShape(ctx, id)`, so the proxies are never used again — only the ids.
 * The probes that do NOT ask for ids are untouched, and still measure exactly
 * what they measured before.
 */
async function scratchShapes(ctx: ProbeContext, boxes: ProbeBox[], load?: string): Promise<PowerPoint.Shape[]> {
  // WHICH of the two syncs failed, in the reason the sheet carries.
  //
  // Both were wrapped in one `catch` and both produced the same
  // `no-scratch-shape`, which is two different diagnoses wearing one string —
  // the mistake `ProbeSetupFailed` itself exists to stop, made one level down.
  // Adding shapes and NAMING them are separate abilities and this host has
  // them separately: `shape-add-fresh-slide-proxy` says yes to the add, and
  // `shape-proxy-survives-one-sync` says `unreadable` to reading anything back
  // through a proxy a sync later.
  //
  // It matters because six of the seven questions this host has never answered
  // are exactly the six that pass a `load` here, and every probe that calls
  // this WITHOUT one gets an answer. That correlation is the whole lead, and
  // until now it could only be read across probes rather than recorded in one.
  let phase = "adding the setup shapes";
  try {
    const shapes = ctx.scratch().shapes;
    const made = boxes.map((box) => shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box));
    awaitingSetupShapes = true;
    await ctx.sync();
    awaitingSetupShapes = false;
    if (load) {
      phase = `reading "${load}" back off the setup shapes`;
      for (const s of made) s.load(load);
      awaitingSetupShapes = true;
      await ctx.sync();
      awaitingSetupShapes = false;
    }
    return made;
  } catch (err) {
    awaitingSetupShapes = false;
    throw new ProbeSetupFailed(`${phase}: ${short(err)}`);
  }
}

/**
 * A shape on the scratch slide, resolved in the batch that is about to use it.
 *
 * The rule that makes the answers below mean anything: **only an id crosses a
 * sync, never a handle.** Four questions on the 2026-08-04 sheet came back with
 * answers that were about the probe rather than the host — a tag written
 * through a group proxy a sync old read back undefined, and a group whose
 * children were counted a sync after it was made reported none. Taken at face
 * value those say no chart in any deck is re-editable, which the same run
 * disproves: its repair pass landed 23 retags on grouped charts.
 *
 * A question whose setup trips the one rule this host is strictest about
 * measures the trip, not the question.
 */
/**
 * The scratch slide's shape collection, resolved in the batch about to use it.
 *
 * Same guard as `probeShape`, for the probes that reach for the collection
 * rather than one shape: a slide that has stopped answering hands back a null
 * object, and `.shapes` off one of those fails in whatever way the host
 * chooses. None of those ways is an answer to the question being asked.
 */
function probeShapes(ctx: ProbeContext): PowerPoint.ShapeCollection {
  let shapes: PowerPoint.ShapeCollection | undefined;
  try {
    shapes = ctx.scratch().shapes;
  } catch (err) {
    throw new ProbeSetupFailed(`the scratch slide stopped answering: ${short(err)}`);
  }
  // Absent, not throwing — the shape of it on a slide the host has stopped
  // resolving, and the more dangerous half: the failure then surfaces as
  // whatever the NEXT line does to `undefined`, which every probe's catch
  // faithfully records as the host having thrown.
  if (!shapes) throw new ProbeSetupFailed("the scratch slide has no shape collection");
  return shapes;
}

function probeShape(ctx: ProbeContext, id: string): PowerPoint.Shape {
  try {
    return ctx.scratch().shapes.getItemOrNullObject(id);
  } catch (err) {
    // The slide stopped answering PART WAY THROUGH a question. Not this
    // probe's answer either — `ScratchSlideUnavailable` only guards the start
    // of a probe's context, and a question that loses its slide at its third
    // batch has asked no more than one that lost it at its first.
    throw new ProbeSetupFailed(`the scratch slide stopped answering: ${short(err)}`);
  }
}

/** The 1.8 binding collection, as much of it as one probe needs. */
type BindingsLike = {
  add(shape: PowerPoint.Shape, bindingType: string, id: string): unknown;
  getItemOrNullObject(id: string): { getShape(): PowerPoint.Shape };
};

/**
 * `presentation.bindings`, or undefined on a host that does not have it.
 *
 * Feature-detected rather than read off `isSetSupported`. The requirement set a
 * host ADMITS to and the API surface it actually exposes are two facts, this
 * file exists because they have come apart before, and the sheet already
 * records the admitted list beside every answer — so a reader can tell "1.8 is
 * missing" from "1.8 is claimed and the object is not there".
 */
function bindingsOf(ctx: ProbeContext): BindingsLike | undefined {
  try {
    const b = (ctx.presentation as unknown as { bindings?: BindingsLike }).bindings;
    return b && typeof b.add === "function" && typeof b.getItemOrNullObject === "function" ? b : undefined;
  } catch {
    return undefined;
  }
}

/**
 * A probe's own catch, for the answers that are real answers.
 *
 * Every probe below turns a throw into `"threw"` or `"unreadable"`, which are
 * things this host genuinely says. A setup failure must not be dressed in
 * either — it is the whole reason `no-scratch-shape` exists — so it goes back
 * up to `ask`, where the run replaces the scratch slide and puts the question
 * again.
 */
function threw(err: unknown): { answer: string; detail?: string } {
  if (err instanceof ProbeSetupFailed) throw err;
  return { answer: "threw", detail: short(err) };
}

/**
 * A shape's creationId, or a word for why there is none.
 *
 * FEATURE-DETECTED, never gated on `isSetSupported`, which is this file's rule:
 * a requirement set a host CLAIMS is a claim, and PowerPointApi 1.10 advertising
 * `creationId` does not mean this host fills it in. The documentation is one
 * sentence and says as much — "Returns null if the shape has no creation ID".
 *
 * Three outcomes that must not be folded together: `absent` (the property is not
 * on the object at all — the set is not really there), `null` (the host has the
 * property and this shape has no id), and a string.
 */
function readCreationId(shape: unknown): string | null | "absent" {
  try {
    const v = (shape as { creationId?: string | null }).creationId;
    if (v === undefined) return "absent";
    return v === null ? null : String(v);
  } catch {
    return "absent";
  }
}

/**
 * A number the host loaded, or NaN if it did not answer.
 *
 * NaN rather than 0 or undefined, deliberately: these readings get compared
 * against an expected geometry, and a width that never arrived but reads as 0
 * compares as a real measurement of zero. `Number.isFinite` then separates "the
 * host said nothing" from "the host said something surprising", which is the
 * distinction a probe exists to report.
 */
function readSize(read: () => unknown): number {
  try {
    return Number(read());
  } catch {
    return NaN;
  }
}

/** A shape's id, or undefined if the host did not answer the load. */
function readShapeId(shape: { id?: string }): string | undefined {
  try {
    const v = shape.id;
    return typeof v === "string" && v ? v : undefined;
  } catch {
    return undefined;
  }
}

/** A queued `getCount()` result, or undefined if the sync did not populate it. */
function loadedCount(result: { value: number }): number | undefined {
  try {
    const v = result.value;
    return typeof v === "number" ? v : undefined;
  } catch {
    return undefined;
  }
}

/** A slide's id, or undefined if the host did not answer the load. */
function readId(slide: PowerPoint.Slide): string | undefined {
  try {
    const v = (slide as unknown as { id?: string }).id;
    return typeof v === "string" && v ? v : undefined;
  } catch {
    return undefined;
  }
}

/** The same, for probes whose failure vocabulary is `"unreadable"`. */
function unreadable(err: unknown): { answer: string; detail?: string } {
  if (err instanceof ProbeSetupFailed) throw err;
  return { answer: "unreadable", detail: short(err) };
}

/**
 * The ids of shapes just made — the only thing allowed to outlive their sync.
 *
 * A host that will not read an id back has not answered the question either,
 * so this is a setup failure rather than a finding: `no-scratch-shape`, in the
 * vocabulary no probe can produce as its own answer.
 */
/**
 * Two shapes to group, and HOW they were obtained.
 *
 * The three group questions were unanswerable on the one host that matters, and
 * `scratchShapes`' own comment had already named the lead: the six questions
 * this host never answers are exactly the six that pass a `load`. They add
 * shapes, sync, read the ids back, and it is that read the host refuses — the
 * same refusal behind `shape-proxy-survives-one-sync: unreadable` and the empty
 * collection reads everywhere else. The questions were never reached. They died
 * in their own setup and reported `no-scratch-slide`, which is a statement about
 * the probe, not the host.
 *
 * So: try the strict route first, and fall back to grouping the proxies in the
 * batch that created them — no id, no sync in between, nothing crossing.
 *
 * The fallback is a DIFFERENT question and must never be allowed to look like
 * the same one. Grouping members resolved by id is what production does;
 * grouping same-batch proxies is the friendliest case a host could be given. An
 * answer from the second says nothing about the first. Hence `via`, which every
 * caller puts in its `detail` — a sheet that says `two (members via same-batch)`
 * and one that says `two (members via ids)` record two different facts, and the
 * diff can tell them apart.
 */
async function groupMembers(
  ctx: ProbeContext,
  boxes: ProbeBox[],
): Promise<{ members: PowerPoint.Shape[]; via: "ids" | "same-batch" }> {
  try {
    const ids = idsOf(await scratchShapes(ctx, boxes, "id"));
    if (ids.length === boxes.length) return { members: ids.map((id) => probeShape(ctx, id)), via: "ids" };
  } catch (err) {
    // Only a SETUP failure earns the fallback. Anything else is the host
    // answering the question and must travel on unchanged.
    if (!(err instanceof ProbeSetupFailed)) throw err;
  }
  // No `load`, so nothing is read back and nothing crosses a sync: the members
  // are the proxies `addGeometricShape` just returned, used in their own batch.
  return { members: await scratchShapes(ctx, boxes), via: "same-batch" };
}

function idsOf(made: PowerPoint.Shape[]): string[] {
  const ids = made.map((s) => {
    try {
      const v = (s as unknown as { id?: string }).id;
      return typeof v === "string" && v ? v : undefined;
    } catch {
      return undefined;
    }
  });
  if (ids.some((id) => !id)) throw new ProbeSetupFailed("the host would not read back the shapes' ids");
  return ids as string[];
}

/**
 * The questions.
 *
 * Each one corresponds to something `powerpoint.ts` relies on. That is the bar
 * for adding one: not "this is interesting about Office.js" but "if the answer
 * were different, code in this repo would be wrong". A probe that no code
 * depends on is a fact nobody will act on.
 */
const PROBES: Probe[] = [
  {
    id: "load-isnullobject-populates",
    question: "Does load('isNullObject') make isNullObject readable after a sync?",
    // The assumption behind `queueNullCheck`, and it is a NEGATIVE one: this
    // file loads "id" instead precisely because loading the flag by name does
    // not populate it. If a host answers yes, `queueNullCheck`'s whole comment
    // is wrong for that host.
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject(ctx.scratchId);
      slide.load("isNullObject");
      await ctx.sync();
      try {
        const v = (slide as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: typeof v === "boolean" ? "yes" : "unreadable", detail: `read ${String(v)}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "load-id-populates-isnullobject",
    question: "Does loading a REAL property make isNullObject readable?",
    // The positive half, and what `queueNullCheck` actually does.
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject(ctx.scratchId);
      slide.load("id");
      await ctx.sync();
      try {
        const v = (slide as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: typeof v === "boolean" ? "yes" : "unreadable", detail: `read ${String(v)}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "getitemornullobject-missing",
    // Resolves a deliberately bogus id. Needing a real slide to ask about a missing one is a contradiction.
    noSlideNeeded: true,
    question: "What does getItemOrNullObject give for an id that does not exist?",
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject("ssf-charts-no-such-slide");
      slide.load("id");
      await ctx.sync();
      try {
        const nul = (slide as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: nul === true ? "null-object" : nul === false ? "claims-live" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-add-fresh-slide-proxy",
    question: "Can a shape be added through a slide proxy resolved in THIS sync?",
    // The three questions below exist because a real host refused eight others
    // and the sheet could not say which of two things it meant: that this host
    // will not take a shape on a freshly-added slide at all, or that it will
    // not take one through a slide proxy it resolved a sync earlier. Both
    // explain every failure in that sheet; they call for different code. So ask
    // the three ways apart, and let the next sheet say which.
    ask: async (ctx) => {
      try {
        probeShapes(ctx).addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 10,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-add-held-slide-proxy",
    resample: true,
    question: "Can a shape be added through a slide proxy resolved a sync ago?",
    // Writing through a handle the host has stopped honouring is this question,
    // and it costs the slide. Seven rounds say whoever follows never gets an
    // answer — see `Probe.burnsTheSlide`.
    burnsTheSlide: true,
    // What `withProbeContext` used to hand every probe, and what the whole
    // add-in avoids on freshly-added slides through `SlideThunk`: Office.js
    // rewrites a resolved proxy's object path to `getItem(id)`, and a new
    // slide's id does not round-trip through `getItem` on the web.
    ask: async (ctx) => {
      const held = ctx.scratch();
      held.load("id"); // a REAL property: this is the sync that resolves it
      await ctx.sync();
      try {
        held.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 40,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
    // THE PARTNER: the same question again, immediately, on a second fresh slide.
    //
    // This question is the one open lead on this host. Sorted by which pass
    // asked it, the observations separate almost perfectly — pass 1 gave
    // `threw` seventeen times out of eighteen, every later pass gave `yes`
    // three times out of three — and that second row did not exist until
    // `PROBE_PASSES` shipped, so sixteen earlier rounds were sixteen samples of
    // ONE condition rather than sixteen tosses of a coin.
    //
    // Two readings fit that, and they want opposite things from this project:
    //
    //   the MOMENT   the host's answer is a property of when it is asked, so
    //                there is a state to find and `yes` is reachable on purpose.
    //   a COIN       it lands differently every time, the pass split is luck,
    //                and no amount of re-running whole rounds will ever settle it.
    //
    // A third reading — the scratch slide's own age — is already dead, and from
    // data rather than argument: the 2026-08-11 run log shows the recorded
    // answer came from a brand-new REPLACEMENT slide in all three passes (first
    // attempt `no-scratch-slide`, replacement added, retry answered), so slide
    // age was identical while the answer changed.
    //
    // Asking again at the same instant is what separates the two that are left.
    // Same second of the run, same host state, same brand-new slide, one
    // difference: this is a second toss. Agreement says the host has a definite
    // behaviour right now and the variable is the moment; disagreement says it
    // is a coin and the populations above are an artefact of when we happened to
    // look. Either way one round decides it, which is the whole point of
    // `Probe.follow` — the reasoning was never the expensive part, the round
    // trip was.
    //
    // Fires on BOTH real answers on purpose. A partner that only followed `yes`
    // would sample the coin exactly where the old design already sampled it.
    follow: {
      when: (answer) => answer === "yes" || answer === "threw",
      because: "asking the identical question a second time at the same instant is what tells a host state from a coin",
      probe: {
        id: "shape-add-held-slide-proxy-again",
        question: "Asked a second time moments later on another fresh slide, does it answer the same way?",
        // No `resample` mark, and it does not need one: a follow-up is never
        // scheduled on its own, it rides its trigger, and the trigger carries
        // the mark — which is how this pair got three paired samples in one
        // round. It carried the mark while it sat in `PENDING_QUESTIONS`, and
        // came off it on 2026-08-11 (`756682e`) as stable across three passes.
        //
        // THAT STABILITY CLAIM IS DEAD. Round `89675b6` (2026-08-12) flipped it
        // inside a single round — `threw` on pass 1 while the host was healthy,
        // `yes` on pass 2 in slide-trouble — with its TRIGGER answering `threw`
        // both times. So the pair has now been seen the other way round from
        // the way this file describes it: the partner is the coin and the
        // trigger held. It is in `UNSTABLE_ANSWERS` for that, which a follow-up
        // could not be until the shortlist invariant learned that a partner
        // rides its trigger rather than carrying its own mark.
        // Same damage as its trigger — it writes through a proxy this host has
        // stopped honouring — so it gives up its slide too rather than handing
        // the wreckage to whatever runs next. See `Probe.burnsTheSlide`.
        burnsTheSlide: true,
        // Deliberately byte-for-byte the trigger's question. A partner that
        // differed in any detail would be a second question, and the answer
        // would no longer be about the host's consistency.
        ask: async (ctx) => {
          const held = ctx.scratch();
          held.load("id");
          await ctx.sync();
          try {
            held.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
              left: 40,
              top: 100,
              width: 20,
              height: 20,
            });
            await ctx.sync();
            return { answer: "yes" };
          } catch (err) {
            return threw(err);
          }
        },
      },
    },
  },
  {
    id: "binding-names-shape-later",
    // RESAMPLED SINCE 2026-08-29, and this one is worth its slide. The idea was
    // retired on the strength of `commit-threw` — the host rejecting the batch
    // that carries a binding. Over the 25 rounds to 2026-08-28 that face appears
    // ONCE, against `yes` 15 times and `silent` 8. The evidence that closed the
    // last untried way out of the id refusals has weakened, so keep asking.
    resample: true,
    // Burns the slide, found the same way the first two were — by the question
    // BELOW it never being answered. `shape-resolve-held-slide-proxy` at the
    // next slot starved on 4 of 4 attempts in round 23 and 4 of 4 in round 26,
    // every one of them "the host says the scratch slide is gone", at elapsed
    // times from 9s to 112s. A question that fails at every time in the round
    // is not losing a race for slides; the slide it is handed is already dead.
    // A binding is exactly the kind of thing this host does not forgive on the
    // slide afterwards — the same shape as `shape-add-held-slide-proxy` above.
    burnsTheSlide: true,
    // No `resample` mark since 2026-08-12: the question is ANSWERED
    // (`commit-threw`, with its control arm landing the same batch without a
    // binding seconds earlier), so it no longer needs the scarce-slide
    // shortlist — that slot goes to a question still without an answer. The
    // reading is in `KNOWN_DIVERGENCES`; a round that contradicts it shows up
    // as a changed answer at the next fixture swap.
    question: "Can a binding made in a shape's CREATING sync still name that shape afterwards?",
    // ASKED SIXTH ON PURPOSE, beside the questions its control arm duplicates.
    //
    // It sat at position 20 for five rounds and was answered in none of them.
    // That is not bad luck, and six sheets say what it is. Sorting every
    // question by whether it needs a SHAPE and by where it runs:
    //
    //   needs no shape, any position  43/48  90%
    //   needs a shape, positions 1-8  23/30  77%
    //   needs a shape, positions 9+   36/77  47%
    //
    // **THE POSITION HALF OF THAT DOES NOT REPRODUCE, AND SHOULD NOT BE PLANNED
    // AROUND.** Recomputed per-attempt over two later rounds it is flat, and in
    // one of them it inverts:
    //
    //   round 23   positions 1-8  20/25  80%   positions 9+  48/69  70%
    //   round 26   positions 1-8  11/20  55%   positions 9+  33/53  62%
    //
    // Round 26 answered #31 — the LAST question on the sheet — while #8, #16,
    // #22 and #23 all starved. A run does not exhaust its slides late; the
    // slots that starve are the same slots every time. Needing a SHAPE still
    // costs you, and that half stands. What the 77-vs-47 split was really
    // measuring is the burnt slots, which cluster after the shape-heavy
    // questions and therefore later in the list — position was the correlate,
    // never the cause.
    //
    // A branch was built on the position reading (moving two questions from
    // 22 and 23 to 5 and 6) and reverted once the per-attempt data came in.
    // Read `Probe.burnsTheSlide` instead: five slots are now known burnt, and
    // three of those five were identified from exactly this recomputation.
    //
    // Position alone was not the whole story, and the round after this moved
    // found out why in the sharpest possible way: it landed in the one slot on
    // the sheet that has never produced an answer — the one directly after
    // `shape-add-held-slide-proxy`, which wrecks the scratch slide on its way
    // out. `shape-resolve-held-slide-proxy`'s 0/6 was blamed here on its reading
    // ids back off setup shapes; it was the neighbour. See `Probe.burnsTheSlide`,
    // which this move is what uncovered, and which now takes a fresh slide in
    // between so the slot is worth having.
    //
    // The only untried route out of the failure that costs this project the most.
    //
    // `same scale across the deck` fails the same way every round: the chart
    // DRAWS, and then the host refuses every handle to it —
    // `InvalidParam passed to GetItem(id)`, code 5010, at
    // `ShapeCollection.getItem` — so there is no group, no config tag, and no id
    // to settle one onto later. On 2026-08-09 that was five charts of eight, each
    // leaving 24 shapes on a slide that is no longer a chart. Both handles the
    // settle pass has are refused: `shapes-items-count-honest` says the
    // collection reads back empty, and `shapes-items-via-positional-slide` says
    // renaming the parent does not help.
    //
    // A binding is the one reference that never crosses either. `bindings.add`
    // takes the live Shape proxy in the batch that CREATED it — no id round trip,
    // no collection read — and the document persists it, so a later, healthier
    // context can ask for the shape back by a key we chose ourselves. If that
    // works here, the repair pass gets the handle it currently lacks and a lost
    // tag becomes a repairable one rather than a chart the user cannot edit.
    //
    // PowerPointApi 1.8 (GA), and this host reports 1.10 — but feature-detected
    // rather than gated on `isSetSupported`, which is a claim about the host
    // rather than a look at it.
    //
    // Deliberately asked across a sync boundary: resolving the binding in the
    // batch that made it would answer a question nobody has, since the live
    // proxy is right there. The later sync IS the question.
    ask: async (ctx) => {
      const bindings = bindingsOf(ctx);
      if (!bindings) return { answer: "no-binding-api", detail: "presentation.bindings is absent (needs 1.8)" };
      // A fixed key on purpose: the docs say a repeat id overwrites its binding,
      // so a run cannot accumulate them, and deleting the scratch slide's shape
      // takes the binding with it.
      const key = "POWERCHART_PROBE_BINDING";
      const box = (left: number) => ({ left, top: 10, width: 20, height: 20 });
      // The CONTROL arm, and the whole reason this question is answerable.
      //
      // The binding has to be made in the batch that CREATES the shape — a
      // proxy one sync old is refused here (`shape-proxy-survives-one-sync`), so
      // binding a committed shape would measure staleness instead. That leaves
      // one batch carrying two things, and its refusal attributable to neither.
      //
      // The 2026-08-09 evening round is exactly that dead end: the commit came
      // back `UnexpectedError` in 1.3 seconds and this probe honestly reported
      // "never asked". `shape-add-fresh-slide-proxy` answered `yes` in the same
      // sheet, which points hard at the binding — but that is a different
      // question asked at a different moment on a host that flaps between
      // minutes, and inferring across two of those has cost this project two
      // full sheets already.
      //
      // So the probe carries its own control: the identical batch WITHOUT the
      // binding, on the same slide, seconds earlier. If that commits and the
      // bound one does not, the binding is the only difference and the answer
      // is about bindings. If the control fails, this host is not taking shapes
      // right now and the question was never put — which is the truth, said by
      // the probe rather than assembled by a reader.
      try {
        probeShapes(ctx).addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box(290));
        await ctx.sync();
      } catch (err) {
        throw new ProbeSetupFailed(`the control shape would not commit: ${short(err)}`);
      }
      let shape: PowerPoint.Shape;
      try {
        shape = probeShapes(ctx).addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box(320));
      } catch (err) {
        throw new ProbeSetupFailed(`adding the shape to bind: ${short(err)}`);
      }
      let made: unknown;
      try {
        made = bindings.add(shape, "Shape", key);
      } catch (err) {
        // Synchronous, before any round trip, so this is `bindings.add` itself
        // objecting to the call — a genuine answer about the binding API.
        return { answer: "add-threw", detail: short(err) };
      }
      try {
        await ctx.sync();
      } catch (err) {
        // An ANSWER now, not a setup failure: the same batch minus the binding
        // committed on this slide moments ago.
        return {
          answer: "commit-threw",
          detail: `the same batch without a binding committed seconds earlier: ${short(err)}`,
        };
      }
      if (!made) return { answer: "add-returned-nothing" };
      try {
        const got = bindingsOf(ctx)!.getItemOrNullObject(key).getShape();
        got.load("id");
        await ctx.sync();
        const id = readShapeId(got as unknown as { id?: string });
        return id ? { answer: "yes", detail: `shape id=${id}` } : { answer: "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-resolve-held-slide-proxy",
    question: "Can a shape be RESOLVED (not added) through a slide proxy resolved a sync ago?",
    // The half of the held-handle rule nobody has asked about. It has never
    // once been answered — 133 rounds, 133 `no-scratch-shape` — and the reason
    // is in the body below.
    //
    // ONE PRODUCTION SITE RESTS ON IT, NOT THREE. This comment used to name
    // `deleteShapesById`, `setShapeSelection` and the selection path, and two of
    // those three no longer do the thing (corrected 2026-08-22):
    //
    //   - the in-place update path takes a DELIBERATELY FRESH handle for its
    //     lookups and says so at the call site — see `powerpoint.ts`, "A FRESH
    //     handle for the lookups, not the one the liveness check above just
    //     resolved";
    //   - `setShapeSelection` never resolves a SHAPE through the aged handle at
    //     all. It calls `setSelectedShapes(ids)` ON the slide, which is a
    //     different question — a method on an aged slide, not a by-id lookup
    //     through one.
    //
    // Which leaves `deleteShapesById`, and only when there is wreckage to sweep.
    // Its trace has not fired once in 133 rounds, because a healthy round leaves
    // no strays — so that site cannot answer this either, and now says so
    // positively rather than by silence.
    //
    // AND PRODUCTION HAS ALREADY ANSWERED THE OTHER HALF, unread for 133 rounds.
    // `setShapeSelection` resolves a slide, syncs, checks liveness, and then
    // calls `setSelectedShapes` on that one-sync-old handle — and `which
    // selection call wedges the host` reports "the host answered all 7 rung(s)"
    // in every round of the archive. So "this host refuses an aged slide handle"
    // is too broad a rule to carry around: it refuses `shapes.add*` through one
    // (`shape-add-held-slide-proxy`, threw in 128 of 133) and accepts
    // `setSelectedShapes` through one, every round, without exception.
    //
    // `shape-add-held-slide-proxy`'s error named `errorLocation:
    // SlideCollection.getItem`, which points at the slide lookup rather than at
    // the add and would mean reads fail too. Against 133 rounds of selection
    // working through the same kind of handle, that reading is now the weaker
    // one. The question below is what would settle it.
    ask: async (ctx) => {
      // REACHED FOR POSITIONALLY, BECAUSE THE ID ROUTE IS CLOSED ON THIS HOST.
      //
      // NO ROUTE FOUND, and the reason is worth more than a workaround.
      //
      // This needs an id, because what it exists to check needs an id: all three
      // production sites — `deleteShapesById`, `setShapeSelection` and the
      // selection path — reach through an aged slide handle for
      // `shapes.getItemOrNullObject(<id>)`. And an id for a fresh shape is
      // exactly what this host will not give. It will not name a shape in the
      // batch that created it (why `scratchShapes` takes a second sync at all),
      // and `shape-proxy-survives-one-sync` answers `unreadable` for reading the
      // proxy back a sync later. Both doors, closed.
      //
      // Reaching by `shapes.getItemAt(0)` instead was tried and REVERTED: it
      // makes the question answerable and worthless. `getItemAt` appears nowhere
      // in `src/` — every production site is `getItemOrNullObject(id)` — so a
      // `yes` from the index route would say nothing about the three call sites
      // this probe is named after. An answerable question about the wrong thing
      // is worse than an honest hole, and this file has paid for that lesson
      // before.
      //
      // It becomes askable the moment anything on this host consents to name a
      // shape: an id from a chart the self-test has already drawn and tagged
      // would do, since those shapes demonstrably carry ids the host honours.
      // That needs the probe run to see the self-test's output, which it
      // currently cannot.
      // BUILT 2026-08-25, from the paragraph directly above this one.
      //
      // `ctx.namedShape` is a chart the self-test has drawn AND tagged — the tag
      // writing is the proof this host honours that id. It closes the hole the
      // comment described: the question needs an id, the host will not name a
      // fresh scratch shape, and a tagged chart's id is one it already named.
      //
      // The QUESTION IS UNCHANGED. It still resolves a shape by
      // `getItemOrNullObject(<id>)` through a slide handle that is a sync old,
      // which is what `deleteShapesById` does. Reaching by `getItemAt(0)` was
      // tried and reverted for making the question answerable and worthless;
      // this does the opposite, keeping the production shape and supplying the
      // one input the harness could not manufacture.
      //
      // The scratch route stays as the fallback, so a first pass — before any
      // chart exists — still reports `no-scratch-shape` rather than pretending.
      const known = ctx.namedShape;
      const id = known
        ? known.shapeId
        : idsOf(await scratchShapes(ctx, [{ left: 10, top: 140, width: 20, height: 20 }], "id"))[0];
      // The chart's OWN slide when using a chart, because that is where its
      // shape lives; the scratch slide otherwise.
      const held = known ? ctx.slides.getItemOrNullObject(known.slideId) : ctx.scratch();
      held.load("id"); // a REAL property: this is the sync that resolves it
      // `isNullObject` is NOT in that load list, deliberately. It is populated by
      // the sync the proxy took part in rather than by being asked for, so
      // loading one real property is what makes it readable — which is why the
      // documented Office.js pattern loads a property and then reads the flag.
      // Asking for both was a distinction with no difference: the mutant that
      // removed it could not be killed, on the fake or on a real host.
      await ctx.sync();
      // WHETHER THE SLIDE CAME BACK AT ALL, asked before the shape.
      //
      // Round 241 asked this question for the first time in 217 rounds and
      // answered `unreadable`, which would have been read as the host losing a
      // shape proxy across a slide switch — office-js #2903, the thing the probe
      // is named for. It cannot mean that yet, because a null SLIDE handle
      // produces a null shape whose id reads undefined by exactly the same path.
      // Two very different facts arriving as one string is what this file calls
      // "two diagnoses wearing one string", and it has cost rounds before.
      //
      // The id spaces on this host do not obviously agree — the scratch ids read
      // `4123571115#123571113` while the deck listed `256#109857222` — so a
      // slide id minted by the draw path failing to resolve here is not a remote
      // possibility. If that is what happened, it is a SETUP failure of ours and
      // not a finding about the host.
      //
      // Safe to trust: this same sheet answers `load-isnullobject-populates:
      // yes` on this host, so the flag is real rather than left undefined.
      if (known) {
        const slide = held as unknown as { isNullObject?: boolean; id?: string };
        if (slide.isNullObject === true)
          return {
            answer: "no-named-slide",
            detail: `slide ${known.slideId} came back a null object`,
          };
      }
      try {
        const shape = held.shapes.getItemOrNullObject(id);
        shape.load("id");
        await ctx.sync();
        const back = (shape as unknown as { id: string }).id;
        // WHICH KIND OF NOTHING, when the id does not come back.
        //
        // The same distinction this probe now draws for the slide, one level
        // down, and it changes what the answer means. A null shape proxy says
        // the host did not find that id on that slide — it lost the shape, or
        // never agreed the id belonged there. A proxy that is NOT null but whose
        // id never populated says the host found it and would not read it back,
        // which is the stale-proxy failure of office-js #2903 proper.
        //
        // Recorded in the detail rather than as a new answer: `unreadable` is
        // the right verdict for both — the id route through an aged slide handle
        // does not work, which is what the three production sites need to know —
        // and splitting the ANSWER would break every comparison with the 217
        // rounds behind it.
        // WHETHER THE SLIDE STILL HOLDS ANYTHING, which is what separates the two
        // readings of "no such shape".
        //
        // Round 243 answered `no such shape on that slide` and that is two
        // stories again. Either the host refuses a by-id lookup for a shape
        // sitting right there — the finding, and the same wall the draw path hit
        // twice in that round's own trace ("a by-id lookup refused the whole
        // resolve") — or the chart we recorded was redrawn away before the
        // re-ask, in which case the id is stale and the fault is ours.
        //
        // A COUNT, not a lookup, so the question itself is untouched: the verdict
        // still comes from `getItemOrNullObject(<id>)` exactly as the three
        // production sites use it. `getItemAt` was once tried AS the question and
        // reverted for making it answerable and worthless; using a count as
        // CONTEXT is the opposite move.
        //
        // Best-effort: a host that will not count must not cost the answer we
        // already have.
        let holds = "count unread";
        try {
          const n = held.shapes.getCount();
          await ctx.sync();
          holds = `slide holds ${String((n as unknown as { value?: number }).value)} shape(s)`;
        } catch {
          holds = "slide would not count its shapes";
        }
        // IS THE ID ACTUALLY THERE? The count says the slide is not empty; this
        // says whether the shape we asked for is one of the ones on it, and that
        // is what finally separates the two readings of "no such shape".
        //
        // Present in the list and unresolvable by id: the host refuses a by-id
        // lookup for a shape sitting right there, which is the finding and what
        // `deleteShapesById` and `setShapeSelection` need to know. Absent: the
        // chart is genuinely gone and the record that named it is still stale,
        // which is ours to fix and not a fact about PowerPoint.
        //
        // Enumeration as CONTEXT, never as the question — the same line the
        // count walks. `getItemAt` was tried AS this question and reverted for
        // making it answerable and worthless; the verdict above still comes from
        // `getItemOrNullObject(<id>)` exactly as the production sites use it.
        //
        // Best-effort last: this host has answered `unreadable` and `not-listed`
        // for the shape collection in other rounds, and a collection that will
        // not list must not cost the answer already in hand.
        let listed = "shapes not listed";
        try {
          held.shapes.load("items/id");
          await ctx.sync();
          const ids = held.shapes.items.map((x) => (x as unknown as { id?: string }).id);
          listed = ids.includes(id)
            ? `the id IS among the slide's ${ids.length} listed shapes`
            : `the id is NOT among the slide's ${ids.length} listed shapes`;
        } catch {
          listed = "the slide would not list its shapes";
        }
        const matched = back === id;
        // ONLY WHEN THE ID DID NOT COME BACK. Round 244 answered `yes` and
        // carried "the shape resolved but would not read back" in the same
        // string, because this was computed whatever the verdict was: the shape
        // is not a null object on the success path either, so the else-branch
        // fired and flatly contradicted the answer beside it. A detail that
        // argues with its own verdict is worse than no detail — it is the kind
        // of line someone quotes months later against the answer it belongs to.
        const lost = matched
          ? ""
          : (() => {
              try {
                return (shape as unknown as { isNullObject?: boolean }).isNullObject === true
                  ? "no such shape on that slide"
                  : "the shape resolved but would not read back";
              } catch {
                // `isNullObject` unreadable is itself an answer: the host
                // populated nothing for this proxy, so the sync it took part in
                // told us nothing.
                return "the host populated nothing for the proxy";
              }
            })();
        return {
          answer: matched ? "yes" : "unreadable",
          detail: `read ${String(back)}${
            known
              ? ` (through a tagged chart's id ${id}, on slide ${String(
                  (held as unknown as { id?: string }).id,
                )}) — ${lost ? `${lost}; ` : ""}${holds}; ${listed}`
              : ""
          }`,
        };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-add-fresh-getitem-slide",
    question: "Can a shape be added through slides.getItem(id) on a freshly-added slide?",
    // The fourth way of naming the same slide, and the one every other question
    // here quietly avoids: all seventeen resolve through `getItemOrNullObject`,
    // so `getItem` has never been asked about in either direction.
    //
    // It is not academic. `getTargetSlide` — the only `slides.getItem(id)` in
    // `powerpoint.ts` — is how `insertSceneIntoSlide` finds the slide it draws
    // on, and it holds that handle for the whole draw. Every held-handle
    // failure this host has reported names `errorLocation:
    // SlideCollection.getItem`, which reads as though this must fail; but the
    // failures were all through handles resolved a sync earlier, and a fresh
    // `getItemOrNullObject` on the same slide works. Those are different
    // claims, and the difference decides whether that insert path is broken.
    ask: async (ctx) => {
      const slides = ctx.slides as unknown as { getItem(id: string): PowerPoint.Slide };
      try {
        const byGetItem = slides.getItem(ctx.scratchId);
        byGetItem.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 100,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
    // The partner, and the one pair this project has already paid for twice.
    //
    // A refusal above has two readings and they lead opposite ways: `getItem`
    // cannot name a FRESH slide (so `insertSceneIntoSlide`'s held handle is
    // fine on the slides users actually edit, and the risk is confined to the
    // one caller that passes a new id), or `getItem` is broken on this host
    // full stop (so the everyday insert path is broken for everyone). Nothing
    // in a sheet can tell those apart, and reasoning about which cost two
    // sessions.
    //
    // Read-only, and that is not a detail. The obvious control — repeat the
    // shape ADD on a pre-existing slide — would draw on a slide in the owner's
    // own presentation, and a diagnostic that litters someone's deck is one
    // they stop clicking. Resolving the slide and reading its id back proves
    // the same thing about `getItem` and changes nothing.
    follow: {
      when: (answer) => answer !== "yes",
      because: "getItem refused a freshly-added slide, which could be about getItem or about the slide's newness",
      probe: {
        id: "getitem-durable-slide",
        question: "Does slides.getItem(id) resolve a slide that was in the deck before this run?",
        ask: async (ctx) => {
          const durable = ctx.durableSlideId;
          if (!durable) return { answer: "no-durable-slide", detail: "the deck has no slide this run did not add" };
          const slides = ctx.slides as unknown as { getItem(id: string): PowerPoint.Slide };
          try {
            const slide = slides.getItem(durable);
            slide.load("id");
            await ctx.sync();
            const v = readId(slide);
            return v === durable
              ? { answer: "yes" }
              : { answer: "unreadable", detail: `read back ${String(v)} for ${durable}` };
          } catch (err) {
            return threw(err);
          }
        },
      },
    },
  },
  {
    id: "shape-add-positional-slide-proxy",
    resample: true,
    question: "Can a shape be added through slides.getItemAt(index) rather than by id?",
    // The other half of the same fork. If by-index works where by-id does not,
    // the id is what this host will not take, and every write path that names a
    // freshly-added slide by id needs an index instead.
    ask: async (ctx) => {
      ctx.slides.load("items/id");
      await ctx.sync();
      let index: number;
      try {
        index = ctx.slides.items.findIndex((s) => s.id === ctx.scratchId);
      } catch (err) {
        return unreadable(err);
      }
      // A real fact about this host, not a failure to ask: it listed its slides
      // and the scratch slide was not among them.
      if (index < 0) return { answer: "not-listed" };
      try {
        ctx.slides.getItemAt(index).shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 70,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-proxy-survives-one-sync",
    question: "Is a shape proxy still usable one sync after it was created?",
    // office-js#2903 — the stale-proxy bug the whole `targetRef` design exists
    // for. If a host keeps proxies alive, a lot of re-fetching here is waste.
    //
    // RESAMPLED SINCE 2026-08-29: it is a coin, not the flat refusal this
    // project described. `threw` 18 of the 25 rounds to 2026-08-28 and `yes` 7 —
    // and the entry that read it as settled was comparing against a fixture that
    // happened to carry `threw`. See UNSTABLE_ANSWERS.
    resample: true,
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 10, top: 10, width: 20, height: 20 }]);
      await ctx.sync(); // a second round trip: this is what ages the proxy
      try {
        shape.load("id");
        await ctx.sync();
        const id = (shape as unknown as { id: string }).id;
        return { answer: typeof id === "string" && id ? "yes" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shapes-items-count-honest",
    resample: true,
    question: "After adding 5 shapes, how many does the collection report?",
    // `faults.hollowReads` models a host that answers SHORT without throwing —
    // observed asking about 19 shapes and being told 3. If a host is honest
    // here, the readback paging is cheaper than it needs to be.
    ask: async (ctx) => {
      await scratchShapes(
        ctx,
        Array.from({ length: 5 }, (_, i) => ({ left: i * 5, top: 40, width: 4, height: 4 })),
      );
      const shapes = probeShapes(ctx);
      shapes.load("items/id");
      await ctx.sync();
      try {
        const n = shapes.items.length;
        // Reported as "at least 5" rather than an exact count: the slide
        // carries whatever earlier probes left on it, and the question is
        // whether the host UNDER-reports, not what the total happens to be.
        return { answer: n >= 5 ? "at-least-5" : `short-${n}`, detail: `items=${n}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "shapes-items-via-positional-slide",
    // Burns the slide. `tags-add-same-key-twice` two slots below starved on 4
    // of 4 attempts in round 23 and 3 of 3 in round 26, always "the host says
    // the scratch slide is gone", from 27s to 113s.
    //
    // TWO slots below, not one, and that is the trap: the question directly
    // after this is `getcount-populates-same-sync`, which is `noSlideNeeded`
    // and therefore cannot be the burner and cannot be starved by one either.
    // A slide-free question is transparent to this failure, so the blame for a
    // dead slide has to be traced back to the last question that actually TOOK
    // one. Reading the neighbour is not enough; read the neighbour that used a
    // slide.
    burnsTheSlide: true,
    resample: true,
    question: "Same collection read, but through slides.getItemAt(index) — any different?",
    // The one contaminated answer that could NOT be cleaned up by re-resolving,
    // and the reason it needs a pair instead.
    //
    // A collection read cannot avoid crossing a sync: the load is queued in one
    // batch and the answer read in the next, by definition. So when this host
    // answered the question above with `items` UNDEFINED, there was no way to
    // tell "this host under-reports collections" from "the by-id slide handle
    // the collection hangs off was spent by the very sync that answered it".
    //
    // The three shape-add questions already proved a positional handle works
    // where a held by-id one does not. If this reads back and the one above
    // does not, the collection is fine and the parent handle was the whole
    // story — which is a fact about every readback in `powerpoint.ts`, since
    // the repair pass reads by index and the probe reads by id.
    ask: async (ctx) => {
      // Its own five, not the previous question's. Probes share a scratch
      // slide, so leaning on what an earlier one left would make this answer
      // depend on whether THAT one's setup worked — and the run replaces a
      // scratch slide it loses, so it may not even be the same slide.
      await scratchShapes(
        ctx,
        Array.from({ length: 5 }, (_, i) => ({ left: i * 5, top: 120, width: 4, height: 4 })),
      );
      ctx.slides.load("items/id");
      await ctx.sync();
      let index: number;
      try {
        index = ctx.slides.items.findIndex((s) => s.id === ctx.scratchId);
      } catch (err) {
        return unreadable(err);
      }
      if (index < 0) return { answer: "not-listed" };
      const shapes = ctx.slides.getItemAt(index).shapes;
      shapes.load("items/id");
      await ctx.sync();
      try {
        const n = shapes.items.length;
        return { answer: n >= 5 ? "at-least-5" : `short-${n}`, detail: `items=${n}` };
      } catch (err) {
        return unreadable(err);
      }
    },
    // THE ROUND-254 CONTROL, and the only half of this question that can be put
    // read-only.
    //
    // This question is the worst of the crossings at round 254. Measured over
    // the whole archive at SAMPLE level: `not-listed` 606 times in the 229
    // rounds up to 253 and ZERO in the 200 since; what is there instead is
    // `short-0` 397 and `at-least-5` 173. The boundary is `77f9ca4`, OUR commit,
    // which stopped the probe buying a fresh scratch slide per question and made
    // it re-acquire one by position. So the ask above now reads a slide this run
    // added, lost and RE-ACQUIRED, and nothing in the sheet says so.
    //
    // Two readings fit that and they want opposite things:
    //
    //   the COLLECTION   this host under-reports a shape collection, and
    //                    `slideShapeList`'s corroborated count is load-bearing.
    //   the SLIDE        it under-reports OUR slide, and says nothing about the
    //                    slides a user's charts actually live on.
    //
    // `experiments.ts` found the SECOND for four sibling questions on
    // 2026-08-26, two of them reversing a shipped decision. That is the reason
    // to ask rather than to assume.
    //
    // READ-ONLY, and that is why this is a probe follow rather than an
    // experiment. The write half — "after adding five shapes" — cannot be asked
    // here: `durableSlideId`'s own docstring says nothing may ever write to it,
    // because a diagnostic that litters the owner's deck on every round forever
    // is one they stop running. What is left is the half the crossing is about:
    // whether a collection on a slide we did NOT create enumerates at all.
    follow: {
      // NOT on `at-least-5`. That face is the read WORKING — there is nothing to
      // disambiguate, and firing on it would spend a slide-free question on the
      // one answer that already means what it says.
      when: (answer) => answer !== "at-least-5",
      because:
        "the collection under-reported a slide this run added and re-acquired by position, which could be about the collection or about that slide",
      probe: {
        id: "durable-slide-lists-its-shapes",
        // DELIBERATELY NOT NAMED `shapes-…`. `record` treats any id matching
        // /^shapes?-/ as evidence about the scratch collection and feeds it into
        // `lastRefusalAt`, which stamps the `regime` of every sample taken in the
        // next twenty seconds. A durable-slide observation entering that signal
        // would silently re-classify every other question in the run.
        question: "Same positional collection read, on a slide that was in the deck before this run — any different?",
        ask: async (ctx) => {
          const durable = ctx.durableSlideId;
          if (!durable) return { answer: "no-durable-slide", detail: "the deck has no slide this run did not add" };
          ctx.slides.load("items/id");
          await ctx.sync();
          let index: number;
          try {
            index = ctx.slides.items.findIndex((s) => s.id === durable);
          } catch (err) {
            return unreadable(err);
          }
          // The SAME positional route as the trigger, so the only thing that
          // differs between the two answers is which slide was asked about.
          if (index < 0) return { answer: "not-listed" };
          const shapes = ctx.slides.getItemAt(index).shapes;
          shapes.load("items/id");
          await ctx.sync();
          try {
            const n = shapes.items.length;
            // An empty durable slide cannot answer this and must not pretend to
            // — see `no-shapes-to-list` in `UNINFORMATIVE`.
            return n > 0
              ? { answer: "lists-shapes", detail: `items=${n}` }
              : {
                  answer: "no-shapes-to-list",
                  detail: "the durable slide is empty, so there is nothing to under-report",
                };
          } catch (err) {
            return unreadable(err);
          }
        },
      },
    },
  },
  {
    id: "getcount-populates-same-sync",
    // `slides.getCount()` on the deck. No slide of ours involved.
    noSlideNeeded: true,
    question: "Does getCount()'s value arrive on the sync that queued it?",
    ask: async (ctx) => {
      const count = ctx.slides.getCount();
      await ctx.sync();
      try {
        const v = count.value;
        return { answer: typeof v === "number" ? "yes" : "unreadable", detail: `value=${String(v)}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "tags-add-same-key-twice",
    question: "Does writing the same tag key twice overwrite, or duplicate?",
    // Re-editing a chart rewrites POWERCHART_CONFIG on the same shape every
    // time. If a host appended instead of overwriting, a chart edited ten times
    // would carry ten configs and the reader would pick one arbitrarily.
    // Every write goes through a shape resolved in its own batch. The first
    // version held one proxy across all four syncs, and this host answered
    // "other — value=undefined": not an opinion about tag keys at all, just
    // the stale handle refusing both writes.
    ask: async (ctx) => {
      // ASKED ON THE SLIDE, NOT ON A SHAPE, AND THAT IS WHAT MAKES IT ASKABLE.
      //
      // It used to add a shape and tag that, which meant naming a shape it had
      // just created — and this host will not do that by either route. It will
      // not name a shape in the batch that made it (the reason `scratchShapes`
      // takes a second sync at all), and it will not read a proxy back a sync
      // later either (`shape-proxy-survives-one-sync` answers `unreadable`).
      // Between the two there is no way to get an id for a fresh shape here, so
      // this question went unasked for eleven rounds while being about tag
      // semantics rather than about shapes at all.
      //
      // The scratch slide is a taggable thing whose id this host DOES honour —
      // the whole probe context is built on resolving it fresh every batch — so
      // asking there costs nothing and answers the same question. It is also
      // closer to production than the old form: `Slide.tags` is a real
      // dependency here, carrying `DEMO_SLOT_TAG` in `powerpoint.ts`.
      //
      // "ANSWERS THE SAME QUESTION" IS FALSE, MEASURED 2026-08-26. It does not.
      //
      // This question has answered `other` — UNINFORMATIVE — in 224 rounds, and
      // the detail added this morning says why: `no such tag on the slide`. Put
      // to a REAL slide by `scripts/experiment.mjs`, three times in under three
      // seconds each, the same question answers:
      //
      //     slide-tag-round-trip — overwrites: second write replaced the first
      //
      // So a slide tag lands and a second write overwrites. `DEMO_SLOT_TAG` is
      // fine, and 224 rounds of `other` were measuring the scratch slide rather
      // than tag semantics — the same trap that stopped the grouped-child
      // question being answerable there, and that `shapes-items-count-honest`
      // (`unreadable` in 92% of rounds) has been reporting from the same slide
      // all along.
      //
      // KEPT, NOT MOVED. Moving it to slide 0 would write a probe tag onto the
      // user's deck on every round, and the answer is already in hand from a
      // tool that cleans up after itself. What this question is still good for
      // is watching whether the SCRATCH slide's behaviour changes.
      //
      // `ctx.scratch()` per batch, never a held handle: same rule as everywhere.
      // SAY WHICH SLIDE WAS WRITTEN AND WHICH WAS READ.
      //
      // Round 28 was the first time this question had ever been answered, and it
      // came back `other`, `value=undefined`, stably on all three passes in
      // 411-614ms. Two readings fit that and they are not close in consequence:
      // either this host takes a tag write on a Slide and does not give it back
      // — which would matter, because `powerpoint.ts` puts `DEMO_SLOT_TAG` on
      // slides and reads it back — or the scratch slide was REPLACED between the
      // write batch and the read batch, in which case the undefined is the
      // probe reading a different slide and says nothing about the host.
      //
      // Nothing in the sheet could tell those apart, so the probe now carries
      // the evidence: the slide's own id at each of the three batches. Same id
      // throughout and the answer is about tags; a changed id and the answer is
      // about slide replacement, and is not the host's fault.
      const wrote = ctx.scratch();
      wrote.load("id");
      wrote.tags.add("POWERCHART_PROBE", "first");
      await ctx.sync();
      const again = ctx.scratch();
      again.load("id");
      again.tags.add("POWERCHART_PROBE", "second");
      await ctx.sync();
      const readFrom = ctx.scratch();
      readFrom.load("id");
      const tag = readFrom.tags.getItemOrNullObject("POWERCHART_PROBE");
      tag.load("value");
      await ctx.sync();
      try {
        const v = (tag as unknown as { value: string }).value;
        // IS THE TAG THERE AT ALL? `value=undefined` has been the whole answer
        // for 15 rounds running — `other`, which this file classes as
        // UNINFORMATIVE — and it covers two unrelated facts.
        //
        // A null-object tag means the key is NOT ON THE SLIDE: the write did not
        // land, which is a real finding, because `powerpoint.ts` puts
        // `DEMO_SLOT_TAG` on slides and reads it back. A tag that is NOT null
        // with an unreadable value means the host took the write and will not
        // give it back, which is a different bug with a different owner.
        //
        // Read defensively: an unloaded `isNullObject` throws `PropertyNotLoaded`
        // and that is itself informative — the host answered nothing for this
        // proxy — so it must not take the answer down with it.
        const present = (() => {
          try {
            const n = (tag as unknown as { isNullObject?: boolean }).isNullObject;
            return n === true
              ? "no such tag on the slide"
              : n === false
                ? "the tag IS on the slide"
                : "presence unread";
          } catch {
            return "the host populated nothing for the tag proxy";
          }
        })();
        const ids = [wrote, again, readFrom].map((s) => readShapeId(s as unknown as { id?: string }));
        // THREE UNKNOWNS COMPARE EQUAL, and for seven rounds this said so out
        // loud: `slide stable (?)`, on a host that had refused every one of the
        // three id reads. The instrument was built (#472) to settle whether the
        // undefined tag value came from the slide being replaced mid-question,
        // and it answered `stable` from no evidence at all — the shape of
        // mistake this project calls a floor read as a count.
        //
        // So the ids have to be READ before they can agree. Unknown is its own
        // word, and the prediction that rests on this is undetermined rather
        // than held while it shows.
        const readable = ids.every((id) => typeof id === "string" && id);
        const sameSlide = readable && ids[0] === ids[1] && ids[1] === ids[2];
        return {
          answer: v === "second" ? "overwrites" : v === "first" ? "keeps-first" : "other",
          detail:
            `value=${v}; ${present}; slide ` +
            // THIS CANNOT DETECT A REPLACED SLIDE, and it was built to. All three
            // handles resolve `getItemOrNullObject(scratchId)` from ONE captured
            // string — `scratchId` is fixed for the whole question and only
            // `handle` is re-made per batch — so the three ids are equal by
            // construction and the CHANGED branch is unreachable from here. A run
            // does replace scratch slides, 65 times in a round, but between
            // questions rather than inside one.
            //
            // Kept, narrowed to what it can actually say: whether this host will
            // NAME the slide it is working on. That is a real fact and a familiar
            // one — the same refusal that left `shape-resolve-held-slide-proxy`
            // unasked for 216 rounds.
            (!readable
              ? `unnameable (${ids.map((id) => id ?? "?").join(", ")}) — the host would not name it`
              : sameSlide
                ? `named (${ids[0]})`
                : `ids differ, which should be impossible: ${ids.join(" -> ")}`),
        };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "tags-on-fresh-shape",
    question: "Does a shape added THIS sync already have a .tags collection?",
    // `faults.tagsUndefinedOn` models a host handing back a shape whose `.tags`
    // is undefined — reading `.add` off it throws SYNCHRONOUSLY, which is what
    // made it so expensive: it escaped the tagging loop rather than failing one
    // chart.
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 90, top: 10, width: 20, height: 20 }]);
      const tags = (shape as unknown as { tags?: unknown }).tags;
      return { answer: tags ? "yes" : "undefined" };
    },
  },
  {
    id: "tag-through-refetched-shape",
    // On the shortlist because it is new and untrusted — every question in
    // PENDING_QUESTIONS is asked on every pass until it has said the same thing
    // enough times to mean it.
    resample: true,
    question: "Can a fresh shape be tagged through a handle re-fetched by its own id?",
    // THE PRODUCTION PATH, asked directly. `finishCharts` writes the config tag
    // through `slide.shapes.getItemOrNullObject(id)`, where the id was read off
    // a shape the run created a sync earlier — the rule that only an id crosses
    // a sync, never a handle. Rounds 29 and 30 then failed `same scale across
    // the deck` in the same way: `InvalidParam passed to GetItem(id)` (5010) at
    // `writing the chart's config tag`, charts left with no config, the scenario
    // stopping after the second consecutive loss.
    //
    // `tags-on-fresh-shape` above already says a shape created THIS sync has a
    // usable `.tags`, and it answers yes every round. The id round trip is the
    // untested half of that path, and it is the half production uses.
    //
    // Deliberately a question and not a fix. The tagging path has a history of
    // changes reverted on a theory, and every investigation that reasoned about
    // this host instead of asking it has been wrong. `threw` here means stop
    // re-fetching; `yes` means the 5010 is about something else and a rewrite
    // would have been wasted work.
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 150, top: 10, width: 20, height: 20 }], "id");
      const id = (shape as unknown as { id: string }).id;
      if (typeof id !== "string" || !id) return { answer: "no-id", detail: "the fresh shape would not report an id" };
      const refetched = probeShapes(ctx).getItemOrNullObject(id);
      try {
        refetched.tags.add("POWERCHART_PROBE", "through-a-refetched-handle");
        await ctx.sync();
      } catch (err) {
        return threw(err);
      }
      const back = probeShapes(ctx).getItemOrNullObject(id).tags.getItemOrNullObject("POWERCHART_PROBE");
      back.load("value");
      await ctx.sync();
      try {
        const value = (back as unknown as { value?: string }).value;
        return {
          answer: value === "through-a-refetched-handle" ? "yes" : value === undefined ? "unreadable" : "other",
        };
      } catch (err) {
        return threw(err);
      }
    },
    // `no-id` closes one door and leaves the other one untried. Production does
    // not give up when the id is unreadable — `finishCharts` falls back to the
    // creation proxy — so a round that stops at `no-id` has said nothing about
    // the path the add-in actually takes in that case, which is the case this
    // host produces. Ask the fallback too, in the same breath.
    follow: {
      when: (answer) => answer === "no-id",
      because: "an unreadable id is what production falls back FROM, so the fallback is the half still unmeasured",
      probe: {
        id: "tag-the-creation-proxy-a-sync-later",
        question: "With no id to be had, can the shape still be tagged through the handle that created it?",
        ask: async (ctx) => {
          const [shape] = await scratchShapes(ctx, [{ left: 180, top: 10, width: 20, height: 20 }]);
          try {
            // The sync that ages the handle. Without it this asks
            // `tags-on-fresh-shape`, which already answers yes every round.
            await ctx.sync();
            (shape as unknown as { tags: { add(k: string, v: string): void } }).tags.add(
              "POWERCHART_PROBE",
              "through-the-creation-handle",
            );
            await ctx.sync();
            return { answer: "yes" };
          } catch (err) {
            return threw(err);
          }
        },
      },
    },
  },
  {
    id: "how-many-syncs-a-creation-handle-survives",
    resample: true,
    question: "How many syncs can a creation handle be carried before .tags goes?",
    // THE BUDGET THE FIX NEEDS. `tag-the-creation-proxy-a-sync-later` says a
    // creation handle takes a write ONE sync later, four rounds running. Round
    // 037 then showed production failing through that same handle — every tag
    // failure in the round reported `from: created`, not one `refreshed` — and
    // four of them failed a different way again:
    //
    //   Cannot read properties of undefined (reading 'add')
    //
    // `.tags` GONE, not refused. `tags-on-fresh-shape` answers `yes` every round
    // and asks in the creating batch; the partner asks one sync later and also
    // answers `yes`. Production's handles are older than that — the renderer
    // chunks a chart across batches, so `created[0]` is several syncs old by the
    // time the tag is written.
    //
    // So the question is not whether a creation handle survives but HOW LONG,
    // and the answer is a number the fix can be built against: keep the write
    // inside that many syncs of the draw, or restructure so it never has to be
    // carried further. Guessing the budget is what makes an ordering change a
    // gamble.
    //
    // COVERAGE, stated rather than implied: the `refused-after-N` branch is
    // proven by the `strictTags` sheet, which asserts this question answers
    // `refused-after-1` where the healthy fake answers `survives-8`. The
    // `tags-gone-after-N` branch is real-host only — `faults.tagsUndefinedOn`
    // models a shape whose `.tags` is missing from birth, not one that loses it
    // partway through a run, and only round 037 has shown the latter.
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 210, top: 10, width: 20, height: 20 }]);
      const tagsOf = (s: unknown) => (s as { tags?: { add(k: string, v: string): void } }).tags;
      // Eight is past anything production does to one chart and cheap to ask:
      // each turn is one empty sync, which this host answers in single-digit ms
      // when it is answering at all.
      for (let aged = 1; aged <= 8; aged++) {
        try {
          await ctx.sync();
          if (!tagsOf(shape))
            return { answer: `tags-gone-after-${aged}`, detail: `.tags was undefined at sync ${aged}` };
          tagsOf(shape)!.add("POWERCHART_PROBE", `aged-${aged}`);
          await ctx.sync();
        } catch (err) {
          return { answer: `refused-after-${aged}`, detail: short(err) };
        }
      }
      return { answer: "survives-8" };
    },
  },
  {
    id: "how-many-collection-reads-a-context-survives",
    resample: true,
    question: "How many times can one context re-read a slide's shapes before the answer goes short?",
    // THE STRONGEST LEAD THIS PROJECT HAS, asked instead of acted on.
    //
    // `same scale across the deck` has failed in every round it has run, and its
    // per-chart trace is a decay curve rather than a coin — identical in rounds
    // 043 through 046:
    //
    //     charts 1-3   re-read matches all 24     grouped, config kept
    //     chart 4      re-read matches 20 of 24   partial, thrown away, lost
    //     chart 5      re-read returns NOTHING    lost
    //     charts 6-8   never attempted
    //
    // And grouping is what saves a config (64 grouped, 1 lost; 62 ungrouped, 41
    // lost), so the re-read going short IS the failure, one level above the tag
    // handle everything has been aimed at.
    //
    // The suspect is our own perf work. `updateChartsInSlides` was deliberately
    // made ONE context, four syncs, flat in N — the right shape for a host that
    // can hold a context, and this one appears to degrade as a context is used.
    // `#112` already made the opposite call for the demo deck (one
    // `PowerPoint.run` per slide) for exactly this reason.
    //
    // Asked rather than fixed because the fix is a 390-line restructure of the
    // live update path, and this project has three shipped-broken fixes on
    // record from changing that path on a theory. A number here decides it: if
    // the reads go short at a small fixed count, chunk the update at that
    // boundary; if a context survives twenty, the context is not the limit and
    // the re-read fails for a reason a fresh one would not fix.
    ask: async (ctx) => {
      const want = 3;
      await scratchShapes(ctx, [
        { left: 330, top: 10, width: 20, height: 20 },
        { left: 360, top: 10, width: 20, height: 20 },
        { left: 390, top: 10, width: 20, height: 20 },
      ]);
      // Re-read the SAME collection, in the SAME context, the way the update
      // path does once per chart. A fresh slide handle each time, because that
      // is what the update takes per chart — the question is about the context,
      // not about holding a stale handle.
      for (let read = 1; read <= 12; read++) {
        try {
          const shapes = ctx.scratch().shapes;
          shapes.load("items/id");
          await ctx.sync();
          const items = (shapes as unknown as { items?: unknown[] }).items;
          if (!Array.isArray(items))
            return { answer: `unreadable-at-${read}`, detail: `read ${read} would not list its items` };
          if (items.length < want)
            return {
              answer: `short-at-${read}`,
              detail: `read ${read} listed ${items.length} of ${want} — the context degraded`,
            };
        } catch (err) {
          return { answer: `threw-at-${read}`, detail: short(err) };
        }
      }
      return { answer: "survives-12" };
    },
  },
  {
    id: "collection-read-poisons-the-creation-handle",
    question: "After re-reading the slide's shapes, does the handle that CREATED a shape still take a tag?",
    // THE QUESTION THAT DECIDES THE ORDERING FIX, and the reason that fix is not
    // in this commit.
    //
    // `survives-8` says a creation handle keeps taking writes for at least eight
    // syncs. `tag-through-refetched-shape: no-id` says there is no id to re-fetch
    // one by. Between them the fix looked settled: make the tag target a shape
    // the draw loop never `load()`s, and the write goes through.
    //
    // It was built on 2026-08-15 and the Office.js fake refused it. Production
    // does one thing the two probes above do not: `groupAndTagAll` re-reads the
    // whole slide's shape collection before grouping (`shapes.load("items/id")`),
    // because grouping needs fresh handles. In the fake that read marks the
    // shape resolved for EVERY handle onto it, the creation handle included, so
    // holding the anchor's own load back changes nothing and the write is refused
    // anyway.
    //
    // Whether that is true of the real host is exactly what nobody knows.
    // Office.js gives each proxy its own object path, so a collection item and a
    // creation handle should be independent — and the fake itself takes that
    // view everywhere else, giving a fresh handle its own `syncCreated` and its
    // own tag writer while sharing only the shape's state. `loadedProps` is
    // handle state modelled as shape state, which is either a bug in the fake or
    // the one place it is right and the rest is wrong.
    //
    // A round settles it, and nothing else can:
    //   `yes`      the collection read is innocent, the fake is wrong, and the
    //              ordering fix works — build it.
    //   `refused`  the fake is right, and no arrangement of loads saves the
    //              drawing context's write while grouping needs a re-read. The
    //              fix then has to be a second tag key or nothing.
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 240, top: 10, width: 20, height: 20 }]);
      const tagsOf = (s: unknown) => (s as { tags?: { add(k: string, v: string): void } }).tags;
      try {
        // The re-read production does, on the slide this shape is on — the whole
        // point of the question. Its RESULT is deliberately unused: what is being
        // asked is what the read did to the handle we already hold.
        const shapes = ctx.scratch().shapes;
        shapes.load("items/id");
        await ctx.sync();
        if (!tagsOf(shape)) return { answer: "tags-gone", detail: ".tags was undefined after the collection read" };
        tagsOf(shape)!.add("POWERCHART_PROBE", "after-a-collection-read");
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return { answer: "refused", detail: short(err) };
      }
    },
  },
  {
    id: "does-a-failed-group-poison-the-tag",
    question: "After addGroup is refused, can a tag still be written in that same context?",
    // THE HYPOTHESIS ROUND 043 PRODUCED, and it may retire the whole ordering
    // effort — which is why it is asked before another line of that effort is
    // written.
    //
    // The tag anchor was moved onto a handle nothing resolves, and round 043
    // then scored EXACTLY what the round before it scored: `cfg-tag-5010` six
    // times, `origin tag lost` zero times. If the handle were the lever, both
    // numbers should have moved.
    //
    // What sits between the draw and the tag in production is a grouping
    // attempt, and on this host it is refused — 5010, five times in that same
    // round. A failed sync poisons its own context; this project already knows
    // that and rebuilds a fresh context elsewhere for exactly that reason. And
    // `no chart's tag could be queued` fired five times, which is a CONTEXT-level
    // symptom: the queue refused before any handle was exercised.
    //
    // So the question is not which handle the write goes through but whether the
    // context it goes through is already dead. `tag-the-creation-proxy-a-sync-later`
    // answers `yes` four rounds running and is the control: same handle, same
    // age, no grouping attempt in between.
    //
    //   refused-after-group  the context is the lever, not the handle. Tag in a
    //                        context that has not tried to group, and the anchor
    //                        move was aimed one level too low.
    //   yes                  the context survives a refused group, and the tag
    //                        failure needs another explanation.
    //   tags-gone            THE ONLY ANSWER THIS HOST HAS EVER GIVEN, and it
    //                        was missing from this table for as long as it has
    //                        been firing. 39 occurrences across 36 rounds, last
    //                        in round 221; `yes` and `refused-after-group` have
    //                        NEVER occurred. So of the times the refusal was
    //                        actually provoked, the shape lost `.tags`
    //                        every time — 39 of 39.
    //
    //                        That is narrower than "the context is dead". The
    //                        context is never asked: `.tags` is undefined before
    //                        a write is attempted, so the PROXY is stripped by
    //                        the refusal rather than the write being refused
    //                        through it. The production counterpart is already
    //                        on record — "the failed group took the tag with it
    //                        (target.tags undefined, 155 of 155)".
    //
    //                        It points at a different fix from the one this
    //                        probe was written to test. A fresh context does not
    //                        help a proxy that no longer has the accessor; what
    //                        would is re-resolving the shape by id after a
    //                        refused group, before tagging. Untested — and it
    //                        needs a population, which production has not
    //                        produced in ~20 rounds.
    //   no-refusal           the host grouped today, so the question was never
    //                        put. Not an answer.
    ask: async (ctx) => {
      const shapes = await scratchShapes(ctx, [
        { left: 270, top: 10, width: 20, height: 20 },
        { left: 300, top: 10, width: 20, height: 20 },
      ]);
      const tagsOf = (s: unknown) => (s as { tags?: { add(k: string, v: string): void } }).tags;
      let grouped = false;
      // AGED FIRST, and the first version of this question did not do it — so
      // the host grouped happily three times and the real question was never
      // put (`no-refusal`, round 044). Production never groups shapes this
      // young: the renderer chunks a chart across batches, and by the time
      // `addGroup` is called the members and the slide handle behind them are
      // several syncs old. That is the state this host refuses — five 5010s at
      // `grouping the chart's shapes` in round 043 — and two empty syncs are
      // what it costs to reach it here.
      //
      // The slide handle is taken ONCE, before the ageing, and the members are
      // reached through it afterwards. A real host named the parent as the thing
      // it refused, not the members, so re-taking it would age the wrong half.
      const aged = ctx.scratch().shapes as unknown as { addGroup(items: unknown[]): unknown };
      await ctx.sync();
      await ctx.sync();
      try {
        aged.addGroup(shapes);
        await ctx.sync();
        grouped = true;
      } catch {
        /* expected on this host — the refusal IS the setup */
      }
      if (grouped)
        return {
          answer: "no-refusal",
          // Worth saying at length, because this answer changed meaning once. It
          // used to mean "the shapes were too fresh to be refused"; with the
          // ageing above it means the host grouped handles as old as production's
          // and refused production's anyway — which would make the difference
          // something other than age, and that is a finding rather than a miss.
          detail: "the host grouped through a slide handle two syncs old, so the refusal was never provoked",
        };
      try {
        // SAME context, deliberately. A fresh one answers a different question
        // and is the fix this may end up recommending, not the test of it.
        if (!tagsOf(shapes[0])) return { answer: "tags-gone", detail: ".tags was undefined after the refused group" };
        tagsOf(shapes[0])!.add("POWERCHART_PROBE", "after-a-refused-group");
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return { answer: "refused-after-group", detail: short(err) };
      }
    },
  },
  {
    id: "which-end-a-short-read-drops",
    question: "When a shape collection reads short, which end of it survives?",
    // ASKED FOR AN ORDERING FIX THAT HAS SINCE BEEN REVERTED, AND KEPT ANYWAY.
    //
    // It was written when the chart's config tag moved onto the LAST shape
    // drawn — the one handle no `load()` resolves, and therefore the one this
    // host was expected to accept a tag through. That move measured no effect
    // across five rounds and four builds and is gone; the anchor is `created[0]`
    // again.
    //
    // The question outlives it because it is about the HOST, not about the fix:
    // a deck scan that reads a collection SHORT sees some prefix or some suffix
    // of it, and which one decides whether a tag is findable at all. Same Scale
    // cannot rescale a chart its scan cannot see, wherever the tag sits. Any
    // future attempt to move the anchor needs this answered first, which is
    // exactly the position the reverted one was in.
    //
    // Nobody knows whether it does. The fake truncates the tail, which is a
    // modelling choice and not evidence, and this host's own
    // `shapes-items-count-honest` answers `short-0` — it returns NOTHING, so it
    // says nothing about which end. A `head` answer makes the trade real and
    // worth mitigating; `tail` makes it free; `none` says the question does not
    // arise on this host and the fake is the only place it bites.
    //
    // ASKED BY POSITION, not by id, and that is deliberate. Every probe here
    // that needs an id off a scratch shape spends its life answering
    // `no-scratch-shape` on this host. `left` is set by us at creation and comes
    // back in the same collection read, so this question survives a host that
    // will not name a shape.
    ask: async (ctx) => {
      const lefts = [10, 40, 70, 100, 130, 160];
      await scratchShapes(
        ctx,
        lefts.map((left) => ({ left, top: 40, width: 20, height: 20 })),
      );
      try {
        const shapes = ctx.scratch().shapes;
        shapes.load("items/left");
        await ctx.sync();
        const items = (shapes as unknown as { items?: { left?: number }[] }).items;
        if (!Array.isArray(items)) return { answer: "unreadable", detail: "the collection would not list its items" };
        const seen = items.map((s) => s.left).filter((l) => typeof l === "number");
        const mine = seen.filter((l) => lefts.includes(l));
        /**
         * NEITHER OF THESE IS AN ANSWER, AND BOTH USED TO READ LIKE ONE.
         *
         * The question is which END a short read drops. Only `keeps-head`,
         * `keeps-tail` and `scattered` answer it. These two say the question
         * did not arise: nothing was dropped, or the collection listed shapes
         * that were not ours at all.
         *
         * They were `all` and `none`, which are perfectly good words —
         * `scratch-slides-returned` uses both as genuine answers, which is why
         * neither could simply join `UNINFORMATIVE` under those names. Here they
         * ranked as named answers and LOCKED the row, so a later pass that did
         * see a short read could never displace them.
         *
         * Measured over 304 archived rounds and 908 samples: `keeps-head`,
         * `keeps-tail` and `scattered` appear ZERO times, while 87 rounds lock
         * on `all`. `FAKE_BASELINE` holds `all` too, and `host-diff` compares
         * only the answer — so the sheet has recorded 87 rounds of agreement
         * with the fake on a question this host has never once been in a
         * position to answer.
         *
         * Renamed so the word cannot be misread, and so it can be ranked weak
         * without dragging another probe's real answers down with it.
         */
        if (!mine.length)
          return { answer: "none-of-ours", detail: `${seen.length} shape(s) listed, none of them ours` };
        if (mine.length === lefts.length)
          return { answer: "not-a-short-read", detail: "nothing was dropped — the question did not arise" };
        // Which of OUR shapes came back, in the order we drew them.
        const kept = lefts.filter((l) => mine.includes(l));
        const head = lefts.slice(0, kept.length);
        const tail = lefts.slice(lefts.length - kept.length);
        const same = (a: number[], b: number[]) => a.length === b.length && a.every((v, i) => v === b[i]);
        const detail = `kept ${kept.length} of ${lefts.length}: ${kept.join(",")}`;
        if (same(kept, head)) return { answer: "keeps-head", detail };
        if (same(kept, tail)) return { answer: "keeps-tail", detail };
        return { answer: "scattered", detail };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "delete-then-lookup",
    question: "Right after delete()+sync, does getItemOrNullObject report it gone?",
    // `deleteSlideById` verifies from a FRESH context because of this. If a
    // host answers honestly in the same context, that re-check is unnecessary.
    ask: async (ctx) => {
      // The id is loaded in the SAME sync as the add — a second round trip for
      // it would age the proxy, and this question is not about proxy age.
      const [shape] = await scratchShapes(ctx, [{ left: 120, top: 10, width: 20, height: 20 }], "id");
      const id = (shape as unknown as { id: string }).id;
      (shape as unknown as { delete(): void }).delete();
      await ctx.sync();
      const gone = probeShapes(ctx).getItemOrNullObject(id);
      gone.load("id");
      await ctx.sync();
      try {
        const nul = (gone as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: nul === true ? "reports-gone" : nul === false ? "still-there" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "addgroup-returns-usable",
    question: "Is a group proxy usable in the same sync that created it?",
    // Members resolved in the grouping batch, and the id asked for in that same
    // batch. Both were a sync late before, and this host answered "unreadable"
    // — which said only that a one-sync-old group proxy is refused, a fact
    // three other questions already establish about proxies in general.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 150 + i * 25, top: 10, width: 20, height: 20 })),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as { addGroup(shapes: unknown[]): { load(p: string): void; id: string } }
        ).addGroup(members);
        group.load("id");
        await ctx.sync();
        return {
          answer: typeof group.id === "string" && group.id ? "yes" : "unreadable",
          detail: `members via ${via}`,
        };
      } catch (err) {
        return { ...threw(err), detail: `members via ${via}` };
      }
    },
  },
  {
    id: "group-children-via-getcount",
    resample: true,
    question: "Does a group report its child COUNT, where it will not list them?",
    // The variant that separates two things `group-reports-its-children` runs
    // together, and the 2026-08-08 sheet is why it exists. That question asked
    // through `group/shapes/items/id` and the host answered `threw` — "The
    // property 'items' is not available", office-js#6363's exact signature —
    // even with the load queued in the sync that made the group.
    //
    // But the same sheet says `getcount-populates-same-sync: yes, value=9`. This
    // host COUNTS a shape collection it will not LIST. Nobody has asked whether
    // that also holds for a GROUP's collection, and the answer decides something
    // concrete: `contentShapes` returns UNKNOWN_CONTENT for every grouped slide,
    // which is what makes the reconcile report a slide complete without counting
    // it (`SlotVerdict.measured`). A count is all it needs. If this answers with
    // a number, those verdicts become measurements.
    //
    // Asked in the group's own sync for the same reason its sibling is: a proxy
    // one sync old is refused here, and that would answer a different question.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 250 + i * 25, top: 110, width: 20, height: 20 })),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as {
            addGroup(shapes: unknown[]): { group: { shapes: { getCount(): { value: number } } } };
          }
        ).addGroup(members);
        const count = group.group.shapes.getCount();
        await ctx.sync();
        const n = loadedCount(count);
        return typeof n === "number"
          ? { answer: n === 2 ? "two" : `reports-${n}`, detail: `members via ${via}` }
          : { answer: "unreadable", detail: `members via ${via}` };
      } catch (err) {
        return { ...threw(err), detail: `members via ${via}` };
      }
    },
  },
  {
    id: "group-reports-its-children",
    // Burns the slide, and this is the one that matters most.
    //
    // TWO QUESTIONS USED TO SIT DIRECTLY BELOW IT AND STARVED 125 TIMES OUT OF
    // 125 — `grouped-child-by-id-from-slide` and `tag-on-group-survives`, never
    // once answered in the whole archive. They were retired on 2026-08-21; see
    // `docs/BACKLOG.md`, "Two questions production answered before the probe
    // sheet could".
    //
    // The reason they starved is this probe: a question placed under a
    // slide-burner finds the scratch slide gone. A branch was once built on the
    // rival theory that they starved for sitting LATE in the list, and the data
    // refutes it — round 26 answered #31, the last question of all, while #8,
    // #16, #22 and #23 starved, and the positional split came out 55% for
    // questions 1-8 against 62% for 9-and-later. Starvation here is a property
    // of the SLOT, not of how late the slot is. Anything moved in under this
    // one inherits the same fate.
    burnsTheSlide: true,
    question: "After grouping two shapes, does the group report two children?",
    // The single most load-bearing answer here. A chart IS a group, and the
    // readback measures whether a chart survived by counting what is inside
    // it — so a host that groups successfully and then reports no children
    // makes every chart read back as wreckage, and the repair pass then
    // "fixes" charts that were never broken. The fake's own notes say a
    // version of it that put one shape there did exactly that.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 200 + i * 25, top: 60, width: 20, height: 20 })),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as {
            addGroup(shapes: unknown[]): { load(p: string): void; group: { shapes: { items: unknown[] } } };
          }
        ).addGroup(members);
        // Queued in the batch that MAKES the group, so the answer arrives on
        // the group's own sync. Asked one sync later — the first version — this
        // host answered PropertyNotLoaded, which is a statement about proxy age
        // and not about what the group contains.
        group.load("group/shapes/items/id");
        await ctx.sync();
        const n = group.group?.shapes?.items?.length;
        return {
          answer: n === 2 ? "two" : typeof n === "number" ? `reports-${n}` : "unreadable",
          detail: `n=${n}; members via ${via}`,
        };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "getitemat-past-end",
    // Asks `slides.getItemAt(9999)` of the DECK collection and never touches the scratch slide. Round 15 lost it to `no-scratch-slide` — a free answer thrown away, which is the `untrack-available` incident repeating.
    noSlideNeeded: true,
    question: "What does slides.getItemAt() past the end of the deck do?",
    ask: async (ctx) => {
      try {
        const slide = ctx.slides.getItemAt(9999);
        slide.load("id");
        await ctx.sync();
        const id = (slide as unknown as { id: string }).id;
        return { answer: typeof id === "string" && id ? "returns-something" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "picture-then-shape-read",
    // Resampled because it is a COIN: `yes` then `unreadable` twice inside one
    // round on 2026-08-12 (`1789749`), the host moving rather than two builds
    // disagreeing. See `UNSTABLE_ANSWERS` for why the other value is the
    // dangerous one to be handed.
    resample: true,
    question: "After a picture is added, does re-reading the slide's shapes still answer?",
    // office-js#5022, CLOSED AS COMPLETED on 2024-11-18 — checked against the
    // GitHub API on 2026-08-14, not read off a stale note. It is kept and still
    // asked, because a fix upstream is a claim about the service and this probe
    // is the only thing here that can check it: get shapes, insert an image, delete the
    // old shapes, sync — and `context.sync()` "can run indefinitely" when the
    // shapes are read back to rename the new image. The reporter's only
    // workaround is a 1-2 second pause, and it still recurs.
    //
    // WHY IT CLOSED, added 2026-09-07, and it changes what this probe is
    // evidence about: the reporter WITHDREW it. *"I forgot an effect that occur
    // everytime an element is selected and that use PowerPoint.run &&
    // context.sync. So everytime my code created a Shape, Powerpoint select it
    // and a parallel sync occured. By removing this effect everything work
    // well."* Nothing was fixed upstream, so this cannot be asking whether a fix
    // arrived — it is asking whether a host that was never shown to have the
    // defect has it. Kept: the question is still worth an answer, this is the
    // only thing here that can give one, and its `silent` arm is the repo's only
    // watch on a sync that never returns. In 401 rounds that arm has never
    // fired — 327 unreadable, 47 yes, 23 threw, 4 no-scratch-slide, 0 silent.
    //
    // `drawDemoItem` does exactly this shape. A chart too dense to draw becomes
    // ONE picture, and `needsRefresh` is true whenever `pictureBase64` is set —
    // so the picture is added and the shape collection is re-read a sync later,
    // in the same context. Every unexplained hang this project has recorded is
    // consistent with it, and none of them can be pinned on it without asking.
    //
    // A hang here answers `silent` on the probe budget rather than taking the
    // sheet down, which is the whole reason each question gets its own bound.
    ask: async (ctx) => {
      // `requirementSets()` rather than a private `supports` — same source, and
      // the sheet already carries the list, so a reader can check the gate.
      if (!requirementSets().includes("1.8"))
        return { answer: "no-api", detail: "picture fills need PowerPointApi 1.8" };
      try {
        const [shape] = await scratchShapes(ctx, [{ left: 260, top: 120, width: 40, height: 40 }]);
        (shape.fill as unknown as { setImage(b64: string): void }).setImage(ONE_PIXEL_PNG);
        await ctx.sync();
        // The read that is said to hang — a fresh collection off a fresh slide
        // handle, one sync after the picture landed.
        const shapes = probeShapes(ctx);
        shapes.load("items/id");
        await ctx.sync();
        const items = (shapes as unknown as { items?: unknown[] }).items;
        return items ? { answer: "yes", detail: `${items.length} shape(s)` } : { answer: "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "group-of-existing-shape-readable",
    question: "Can a group found in a later batch report its children?",
    // office-js#5849: reading `shape.group` throws GeneralException. Distinct
    // from `group-reports-its-children`, which asks in the batch that MADE the
    // group — this asks the way `countGroupChildrenPage` asks, about a group
    // resolved from the deck afterwards.
    //
    // That pass is what decides whether a chart read back is intact or
    // wreckage, and every failure in it is swallowed per shape. So a host that
    // refuses this produces no error and no measurement, and the repair pass
    // then has nothing to go on for the charts it was meant to check.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 300 + i * 25, top: 160, width: 20, height: 20 })),
      );
      try {
        const made = (
          probeShapes(ctx) as unknown as { addGroup(shapes: unknown[]): { load(p: string): void; id: string } }
        ).addGroup(members);
        // The id in its OWN sync, never in the one that made the group.
        // That saving is what cost the 2026-08-05 sheet five of its questions:
        // this host will not name a shape in the batch that created it, and
        // `scratchShapes` was rewritten for exactly the same reason.
        await ctx.sync();
        made.load("id");
        await ctx.sync();
        const groupId = readShapeId(made);
        // An ANSWER, not a setup failure. A host that will not name a group it
        // just made cannot be asked the later-batch question at all — there is
        // no id to resolve one with — and that is a fact about the host, so it
        // belongs in the sheet. Reported as `no-scratch-slide` it read as the
        // probe having lost its slide, which is what kept this question out of
        // four consecutive rounds while the actual finding sat one line away.
        if (!groupId) return { answer: "no-group-id", detail: "the host would not name the group it just made" };
        // A LATER batch, through a fresh handle — the thing that separates this
        // question from the one above it.
        const found = probeShape(ctx, groupId) as unknown as {
          group: { shapes: { getCount(): { value: number } } };
        };
        const count = found.group.shapes.getCount();
        await ctx.sync();
        const n = loadedCount(count);
        return typeof n === "number"
          ? { answer: String(n), detail: `members via ${via}` }
          : { answer: "unreadable", detail: `members via ${via}` };
      } catch (err) {
        return { ...threw(err), detail: `members via ${via}` };
      }
    },
  },
  {
    id: "slide-layout-readable",
    // Asked of `ctx.durableSlideId` on purpose — a slide that predates the run — so demanding a live scratch slide refuses it for the absence of something it deliberately avoids.
    noSlideNeeded: true,
    question: "Can a slide's layout shapes be read?",
    // office-js#3826, open and marked a product bug: on Office on the web
    // `slide.load("layout/shapes/items")` fails the sync with GeneralException.
    // The per-slide companion to `layouts-readable`, which asks the same of the
    // deck's masters — #4906 and #2328 report the master form, #3826 the slide
    // form, and nothing says whether they are one defect or three.
    //
    // Asked of a slide that was in the deck BEFORE this run, not the scratch
    // slide, and that is the whole care taken here. A freshly-added slide's
    // handle is good for exactly one sync on this host — established, and
    // unconditional in the fake — so a load queued on it and read after its own
    // sync answers "unreadable" whatever the layout API does. That would be an
    // answer about the probe's own plumbing dressed as a fact about layouts,
    // which is the failure mode this file exists to prevent.
    //
    // Read-only. It resolves someone's own slide and asks what its layout
    // holds; nothing is created, moved or drawn.
    ask: async (ctx) => {
      const durable = ctx.durableSlideId;
      if (!durable) return { answer: "no-durable-slide", detail: "the deck has no slide this run did not add" };
      try {
        const slide = ctx.slides.getItemOrNullObject(durable) as unknown as {
          load(p: string): void;
          layout?: { shapes?: { items?: unknown[] } };
        };
        slide.load("layout/shapes/items/id");
        await ctx.sync();
        const items = slide.layout?.shapes?.items;
        return items ? { answer: "yes", detail: `${items.length} layout shape(s)` } : { answer: "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "layouts-readable",
    // Reads `presentation.slideMasters`. Nothing to do with any slide this run added.
    noSlideNeeded: true,
    question: "Can the deck's slide masters and their layouts be read?",
    // office-js#4906 and office-js#2328: loading shapes off `SlideLayout` or
    // `SlideMaster` throws GeneralException in PowerPoint Online — and #4906
    // reports it happening ONLY on presentations built from a custom template,
    // with `errorLocation: SlideMasterCollection.getItem`. Both are open.
    //
    // `blankLayoutId` reads exactly this, to give every slide the add-in creates
    // a BLANK layout. It is try/caught, so a host that refuses degrades quietly
    // to the inherited layout — which means the demo deck and every scratch
    // slide land on top of the previous slide's placeholders, "Click to add
    // title" showing through the chart. That is a visible defect with no error
    // anywhere, and nobody would know which of the two it was.
    //
    // The owner's decks come from a corporate template, which is precisely the
    // case #4906 singles out. Read-only, and no slide is created or touched.
    ask: async (ctx) => {
      try {
        const masters = ctx.presentation.slideMasters;
        masters.load("items/id,items/layouts/items/id,items/layouts/items/type");
        await ctx.sync();
        const items = (masters as unknown as { items?: { layouts?: { items?: unknown[] } }[] }).items;
        if (!items) return { answer: "unreadable", detail: "the host answered the sync but not the collection" };
        const layouts = items.reduce((n, m) => n + (m.layouts?.items?.length ?? 0), 0);
        // A master with no layouts is not the same as no masters, and neither is
        // the same as a throw. `blankLayoutId` returns undefined for all three
        // and the caller cannot tell them apart; the sheet can.
        return {
          answer: layouts > 0 ? "yes" : "no-layouts",
          detail: `${items.length} master(s), ${layouts} layout(s)`,
        };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "untrack-available",
    question: "Do proxies expose untrack()?",
    // Nothing here syncs: `getItemOrNullObject` returns a proxy without a round
    // trip, and the question is whether that object carries a method. A live
    // slide is not required and must not be demanded — see `noSlideNeeded`.
    noSlideNeeded: true,
    // `untrack` is best-effort here precisely because a null-object proxy may
    // not expose it. Worth knowing which hosts do.
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject(ctx.scratchId);
      const has = typeof (slide as unknown as { untrack?: unknown }).untrack === "function";
      return { answer: has ? "yes" : "no" };
    },
    // …and the `no` above is exactly the shape of answer this project has a
    // rule about: it could be a fact about the host, or a fact about the
    // NULL-OBJECT proxy it was asked of, which the comment two lines up already
    // suspects and which nobody had followed up. The partner separates them,
    // and it asks the proxy type the answer would actually be spent on.
    //
    // What rests on it: Microsoft's own performance guidance names this as the
    // remedy for our exact symptom — "large batch operations may generate a lot
    // of proxy objects… Calling untrack() after your add-in is done with the
    // object should yield a noticeable performance benefit when using large
    // numbers of proxy objects". `renderShapesChunked` creates one proxy per
    // shape and holds every one of them for the whole draw, hundreds per run,
    // and untracks none of them; the read paths untrack and the hot path does
    // not. A `yes` here makes that an omission worth fixing. A `no` closes it,
    // and closes it on the right evidence rather than on a null object's.
    follow: {
      when: (answer) => answer === "no",
      because: "a null-object proxy is the one kind most likely to lack the method, so the no may be about the probe",
      probe: {
        id: "untrack-available-on-shape",
        question: "Does a SHAPE proxy — the kind a draw makes hundreds of — expose untrack()?",
        ask: async (ctx) => {
          // A REAL created proxy, which is what the draw loop holds — not a
          // null object, which is the confound this partner exists to remove.
          // Nothing syncs: `addGeometricShape` hands the proxy back before any
          // round trip, and the question is whether that object carries a
          // method. The queued add rides whatever sync comes next and lands on
          // the scratch slide, which the run deletes.
          const shape = probeShapes(ctx).addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
            left: 10,
            top: 10,
            width: 20,
            height: 20,
          });
          const has = typeof (shape as unknown as { untrack?: unknown }).untrack === "function";
          return { answer: has ? "yes" : "no" };
        },
      },
    },
  },
  {
    id: "shapes-by-index-vs-items",
    resample: true,
    question: "Can the shapes be enumerated by getCount + getItemAt where `items` will not answer?",
    /**
     * THE ONE QUESTION UPSTREAM OF EVERY LIVE DEFECT.
     *
     * **CORRECTED BEFORE IT WAS EVER RUN.** This said "the shape collection will
     * not honestly report freshly added shapes", citing this sheet's own
     * `shapes-items-count-honest` at `unreadable` 142 / `short-0` 16 of 158.
     * That is a SCRATCH-SLIDE answer, and production disagrees with it 2,135
     * times:
     *
     *     charts grouped by id-match   2135
     *     charts grouped by `created`    49
     *     re-read named NONE             97   -> 4.3% of charts
     *
     * The scratch slide is strictly worse at collection reads than a real one —
     * this file says so about three other questions — so a probe answer about
     * collection reads is not a production fact. Quoting one as if it were is
     * how "never works" got written about a call that works 95.7% of the time.
     *
     * THE REAL QUESTION IS NARROWER AND BETTER. The read usually works and
     * fails about one chart in twenty-three, roughly one round in four; when it
     * fails, everything downstream fails with it — the re-read that matches
     * nothing, the positional guess at a stale listing, the `addGroup` throw on
     * another chart's shapes, and the parts list that has never once been
     * written (0 across 872 charts).
     *
     * So this asks whether an INDEX walk survives the 4% where `items` does
     * not. Not whether it can replace a read that never works.
     *
     * Every workaround tried DOWNSTREAM has cost more than it bought: the
     * tag-anchor move (reverted after five rounds), loading ids earlier
     * (implicated in `from: created`, 235 tagging failures), and the fast path
     * in `ungroupedFallback` (shipped, measured inert in round 181).
     *
     * BUT THE HOST IS NOT UNIFORMLY DEAF, and that is the lead. On the same
     * sheets, in the same rounds:
     *
     *     getcount-populates-same-sync   yes    158 of 158
     *     shapes-items-count-honest      unreadable 142 / short-0 16
     *     getitemat-past-end             threw  158 of 158
     *
     * `getCount()` is answered every single time. `items` never is. And
     * `getItemAt` is demonstrably CALLABLE — a past-end index throws, which is
     * the correct answer to a wrong question rather than a refusal to answer.
     *
     * So: can the collection be walked by INDEX where it will not be read as a
     * LIST? If yes, the re-read has a route that does not depend on `items`, and
     * the root is addressable rather than only worked around. If no, that closes
     * the last cheap idea and the answer is worth as much — this project has
     * spent four rounds on a downstream fix that a probe would have refused.
     *
     * DELIBERATELY NOT WIRED INTO PRODUCTION. `getItemAt` appears nowhere in
     * `src/`, on purpose: an earlier attempt to reach a shape that way was
     * reverted for making a question answerable and worthless. This asks the
     * host without changing anything, which is the whole point of a probe.
     *
     * The comparison is the deliverable — both routes, same slide, same sync —
     * because "index worked" means nothing without "and list did not, here,
     * now".
     */
    ask: async (ctx) => {
      const want = 3;
      await scratchShapes(
        ctx,
        Array.from({ length: want }, (_, i) => ({ left: 60 + i * 25, top: 200, width: 20, height: 20 })),
      );
      const shapes = probeShapes(ctx);
      // BOTH ROUTES QUEUED IN ONE SYNC, so neither can be blamed on the host's
      // mood a moment later. `getCount` first: it is the one this host has never
      // refused, and if it fails the index walk has no length to walk.
      const count = shapes.getCount();
      shapes.load("items/id");
      await ctx.sync();
      const n = loadedCount(count);
      if (n === undefined) return { answer: "no-count", detail: "getCount did not populate — the premise is gone" };
      let listed: number | undefined;
      try {
        listed = shapes.items.length;
      } catch {
        listed = undefined;
      }
      // Walk by index in a SECOND sync, because `getItemAt` needs the count that
      // the first one produced.
      try {
        const picked = Array.from({ length: Math.min(n, want) }, (_, i) => shapes.getItemAt(i));
        for (const sh of picked) sh.load("id");
        await ctx.sync();
        const ids = picked.map((sh) => readShapeId(sh as unknown as { id?: string })).filter(Boolean);
        const detail = `count=${n} listed=${listed === undefined ? "unreadable" : listed} byIndex=${ids.length}`;
        if (!ids.length) return { answer: "index-unreadable", detail };
        // The finding is the DIFFERENCE. Index answering where the list did not
        // is the only outcome that changes what the product can do.
        if (listed === undefined || listed < want) return { answer: "index-beats-items", detail };
        return { answer: "both-answer", detail };
      } catch (err) {
        return { ...threw(err), detail: `count=${n} listed=${listed === undefined ? "unreadable" : listed}` };
      }
    },
  },
  {
    id: "creationid-on-fresh-shape",
    question: "Does a shape this run just added report a creationId, or null?",
    /**
     * THE FIRST OF THREE, and the one that decides whether the other two are
     * worth asking.
     *
     * Every user-visible defect left in this product traces to one thing: the
     * pre-grouping re-read does not return the chart's own shapes. The chain is
     * a match that finds none of our ids in the listing, a positional guess at
     * the tail of it, another chart's shapes, `addGroup` throwing, and a chart
     * left as loose rectangles that cannot be re-edited — 6 of the last 20
     * rounds.
     *
     * **THE IDS ARE NOT THE PROBLEM, and this comment said they were.** Round
     * 179's chart reported `mine [35..41]`, and those seven ids are in that
     * round's own end-of-round deck inventory under their real names — `title`,
     * `category-0`, `seg-0-0`, `baseline`. They never moved. 228 in-place
     * updates across 30 rounds resolved 4,992 tag-stored ids written in an
     * earlier context, and `addGroup` keyed on id grouped 2,108 charts against
     * 36 throws. `shape.id` is stable on this host.
     *
     * What is wrong is the LISTING: the re-read returned `[27..33]`, the
     * previous chart's shapes, because the collection read is stale. The host
     * has said so in every round — `shapes-items-count-honest` answers
     * `unreadable` 140 times and `short-0` 16 times across 156 rounds, and
     * `tag-through-refetched-shape` answers `no-id` 149 of 149 — **on the SCRATCH
     * SLIDE, which is strictly worse at collection reads than a real one.** In
     * production the same read matches ids for 2,135 charts and names none for
     * 97, a 4.3% failure share. "Never works" was a scratch answer quoted as a
     * production fact; the read usually works and fails about one chart in 23.
     *
     * So creationId cannot fix that chain — matching on it against the same
     * stale listing matches zero, exactly as id does, and there is no
     * `getItemByCreationId` for it to address a shape with. These probes are
     * kept for the SEPARATE question in BACKLOG ("Retire the positional
     * group-member mapping"), which is about node-to-shape ordering inside a
     * group and is untouched by the above.
     *
     * `Shape.creationId` (PowerPointApi 1.10, which this host advertises) is the
     * obvious answer, and **the documentation does not say what this project
     * has been assuming.** Checked 2026-08-23, the whole of it:
     *
     *     readonly creationId: string | null;
     *     "Gets the creation ID of the shape. Returns null if the shape has no
     *      creation ID."
     *
     * That is all. Nothing about when it is assigned, nothing about surviving a
     * save or a session, and an explicit `null` for shapes that have none.
     * "Durable" is an inference from the NAME — this repo's own backlog calls it
     * "a durable per-shape identifier" and cites no source for it.
     *
     * So it is measured before anything is built on it. `null` here and the
     * whole plan is dead on this host, which is worth one probe slot to learn
     * before it is worth a migration to find out.
     *
     * **ANSWERED, ROUND 184: `absent` — the property is not on the shape at
     * all.** Not null, not equal to `id`: absent. `Shape.creationId` does not
     * exist on this host's shapes, though the host advertises requirement set
     * 1.10, where it is documented. Both siblings answered `no-creation-id` on
     * the same sheet, `before=[absent,absent]`.
     *
     * The route is closed — nothing to record at draw time, nothing to match
     * on, nothing to migrate to — and this is exactly what the probe was for.
     * The backlog called it "a durable per-shape identifier" and cited no
     * source; a migration built on that phrase would have been built on a
     * property that is not there. Three probes, one round, question settled.
     */
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 120, top: 10, width: 20, height: 20 }], "creationId,id");
      const creationId = readCreationId(shape);
      const id = readShapeId(shape as unknown as { id?: string });
      if (creationId === "absent") return { answer: "absent", detail: "the property is not on the shape at all" };
      if (creationId === null) return { answer: "null", detail: `a fresh shape has no creation id; id=${String(id)}` };
      // BOTH, because the point is whether they are different things. If the
      // host simply returns `id` here the property is decoration and the plan
      // is no better than what it replaces.
      return {
        answer: String(creationId) === String(id) ? "same-as-id" : "yes",
        detail: `creationId=${String(creationId)} id=${String(id)}`,
      };
    },
  },
  {
    id: "creationid-survives-a-sync",
    question: "Does a shape's creationId still read the same a sync later, when its id may not?",
    /**
     * THE PROPERTY THE PRODUCT ACTUALLY NEEDS. `shape.id` is the thing that goes
     * stale — `shape-proxy-survives-one-sync` answers `unreadable` in 133 of 133
     * rounds, and the whole "only an id crosses a sync, never a handle" rule
     * exists because of it.
     *
     * This asks whether creationId is any better ACROSS THE SAME BOUNDARY. If it
     * changes too, the migration buys nothing; if it holds while `id` moves, it
     * is exactly the identifier the grouping path has been missing.
     *
     * Re-fetched by POSITION rather than by id on purpose: fetching by id would
     * make the answer depend on the very thing under test.
     */
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 150, top: 10, width: 20, height: 20 }], "creationId,id");
      const before = readCreationId(shape);
      const idBefore = readShapeId(shape as unknown as { id?: string });
      if (before === "absent" || before === null)
        return { answer: "no-creation-id", detail: `before=${String(before)}` };
      try {
        const again = ctx.scratch().shapes.getItemAt(0);
        again.load("creationId,id");
        await ctx.sync();
        const after = readCreationId(again);
        const idAfter = readShapeId(again as unknown as { id?: string });
        if (after === "absent") return { answer: "unreadable", detail: "the host would not load it a sync later" };
        return {
          answer: String(after) === String(before) ? "stable" : "changed",
          detail: `creationId ${String(before)} -> ${String(after)}; id ${String(idBefore)} -> ${String(idAfter)}`,
        };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "creationid-survives-grouping",
    question: "Does a shape keep its creationId after it is absorbed into a group?",
    /**
     * THE STEP THE DEFECT DIES OR SURVIVES ON.
     *
     * The chain that costs a chart is: the pre-grouping re-read matches none of
     * our ids, so the code guesses positionally, and the guess takes shapes that
     * have ALREADY BEEN GROUPED — round 179 recorded `mine [35..41]` against
     * `chose [27..33]`, zero overlap. Grouping is precisely where identity gets
     * lost today.
     *
     * If creationId survives being grouped, the re-read can match on it and the
     * guess never runs. If it does not, the migration fixes the in-place
     * mapping (worth having on its own — see BACKLOG, "Retire the positional
     * group-member mapping") but NOT the defect that is costing charts, and the
     * two must not be conflated in whatever gets built next.
     *
     * `burnsTheSlide` because grouping does: a question placed under a
     * slide-burner finds the scratch slide gone, which starved two probes for
     * 125 rounds before anyone noticed.
     */
    burnsTheSlide: true,
    ask: async (ctx) => {
      const made = await scratchShapes(
        ctx,
        [0, 1].map((i) => ({ left: 180 + i * 25, top: 10, width: 20, height: 20 })),
        "creationId,id",
      );
      const before = made.map((s) => readCreationId(s));
      if (before.some((c) => c === "absent" || c === null))
        return { answer: "no-creation-id", detail: `before=${JSON.stringify(before)}` };
      try {
        const group = ctx.scratch().shapes.addGroup(made);
        group.load("id");
        await ctx.sync();
        const members = (group as unknown as { group?: { shapes?: PowerPoint.ShapeScopedCollection } }).group?.shapes;
        if (!members) return { answer: "no-members", detail: "the group would not name a shapes collection" };
        members.load("items/creationId");
        await ctx.sync();
        let items: PowerPoint.Shape[] | undefined;
        try {
          items = members.items;
        } catch {
          items = undefined;
        }
        if (!Array.isArray(items)) return { answer: "unreadable", detail: "the group would not list its members" };
        const after = items.map((s) => readCreationId(s));
        const kept = before.filter((c) => after.some((a) => String(a) === String(c))).length;
        return {
          answer: kept === before.length ? "kept" : kept ? "partial" : "lost",
          detail: `${kept} of ${before.length} kept; before=${JSON.stringify(before)} after=${JSON.stringify(after)}`,
        };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "rotation-keeps-the-unrotated-box",
    question: "After setting `rotation`, do left/top/width/height still describe the box BEFORE the turn?",
    /**
     * THE ASSUMPTION UNDER EVERY DIAGONAL THIS PRODUCT DRAWS, and until this
     * probe nothing had ever asked it.
     *
     * `addSegment` draws a solid diagonal as a rectangle of the segment's
     * LENGTH, placed at its midpoint, and then rotated. `arrowheadBox` offsets
     * its box on the same premise. Both are only correct if `width` keeps
     * meaning the unrotated side after the turn. If this host instead reports
     * the post-rotation BOUNDING BOX, every diagonal in every line, scatter,
     * radar and violin chart is drawn at the wrong size and in the wrong place.
     *
     * ASKED AS A PROBE RATHER THAN THROUGH A CHART, which is the lesson from
     * round 334. The scenario that asked this by drawing a line chart and
     * reading its segments back SKIPPED: the insert path groups a chart, so
     * `slide.shapes` lists one shape called `PowerChart` and never the segments.
     * The question is about the host, not about our pipeline, so measuring it
     * through the pipeline only added a way to fail.
     *
     * 80 x 10 at 45 degrees: the unrotated width is 80 and the rotated extent is
     * 80·cos45 + 10·sin45 = 63.6. Far enough apart that no rounding can
     * confuse them, and both are checked so a host answering some third thing
     * is reported rather than bucketed.
     */
    resample: true,
    ask: async (ctx) => {
      // `requirementSets()` rather than powerpoint.ts's private `supports`, for
      // the reason the 1.8 probe already gives: same source, and this module
      // answers from what the host declared rather than from a helper the host
      // cannot see.
      if (!requirementSets().includes("1.10"))
        return { answer: "no-rotation-api", detail: "PowerPointApi 1.10 is not declared" };
      const W = 80;
      const H = 10;
      const DEG = 45;
      // THE ADD IS SETUP, NOT THE QUESTION. Through `scratchShapes`, so a host
      // that refuses every add answers `no-scratch-shape` — "never asked" —
      // instead of `threw`, which would read as an opinion about rotation from
      // a run in which nothing was ever rotated.
      await scratchShapes(ctx, [{ left: 100, top: 100, width: 8, height: 8 }]);
      try {
        /**
         * THE MEASURED SHAPE IS ADDED HERE, not returned by `scratchShapes`.
         *
         * Round 336 answered `unreadable` — "width did not come back a number" —
         * and the reason is in the same sheet: `scratchShapes` syncs, so its
         * handle is a sync old by the time anything is read through it, and
         * `shape-proxy-survives-one-sync` says this host refuses exactly that.
         * The probe was measuring proxy staleness, which is the mistake this
         * file's four contaminated 2026-08-04 answers were.
         *
         * `named-preset-resolves`, two probes down, reads its shape without
         * trouble because it ADDS, sets and loads inside ONE batch. Same shape
         * here. The `scratchShapes` call above is kept for the other half: a
         * host that refuses every add must answer `no-scratch-shape` — never
         * asked — rather than offering an opinion about rotation.
         */
        const shapes = probeShapes(ctx);
        /**
         * A CONTROL SHAPE, ADDED AND READ THE SAME WAY, AND NEVER ROTATED.
         *
         * Rounds 336-338 all answered `unreadable` — "width did not come back a
         * number" — while `named-preset-resolves`, two probes down, reads a
         * width in exactly this shape of batch and gets 24x24 every time. So
         * the geometry read is not what this host refuses, and reporting the
         * refusal as one flat "unreadable" hid which of the two it was.
         *
         * The control answers that. If it reads and the rotated one does not,
         * the finding is `no-read-after-rotation` — this host will not answer a
         * load in a batch that also WRITES rotation — which is a fact about the
         * host worth having, and a different one from "we cannot measure here".
         *
         * AND ITS LIMIT, which round 341 reached: the control rides in the SAME
         * batch as the rotation write, so on a host where that write poisons
         * the whole sync BOTH widths come back empty and the answer is
         * `unreadable` — a reading that still cannot be told apart from "this
         * host answers no geometry for a probe-added shape at all". Separating
         * those two needs the control in a batch of its own, which is a third
         * sync and a question nobody has yet needed answered. What the control
         * does settle is that the refusal is not specific to the ROTATED shape,
         * and that is what stops this being read as an opinion about rotation.
         */
        const control = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 200,
          top: 100,
          width: W,
          height: H,
        });
        control.load("width");
        const made = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 120,
          top: 100,
          width: W,
          height: H,
        });
        try {
          (made as unknown as { rotation: number }).rotation = DEG;
        } catch (err) {
          return { answer: "rotation-not-writable", detail: short(err) };
        }
        made.load("left,top,width,height,rotation");
        await ctx.sync();
        const w = readSize(() => made.width);
        const h = readSize(() => made.height);
        const rot = readSize(() => (made as unknown as { rotation: number }).rotation);
        const controlW = readSize(() => control.width);
        if (!Number.isFinite(w))
          return Number.isFinite(controlW)
            ? {
                answer: "no-read-after-rotation",
                detail: `an unrotated shape in the same batch read back ${controlW.toFixed(1)}pt wide; the rotated one answered nothing`,
              }
            : { answer: "unreadable", detail: "neither shape's width came back — this host answers no geometry here" };
        const detail = `set ${W}x${H} at ${DEG}deg; host reports ${w.toFixed(1)}x${h.toFixed(1)}, rotation ${String(rot)}`;
        /**
         * AND IT MUST NOT CONCLUDE FROM A ROTATION THAT DID NOT TAKE.
         *
         * If the set never landed, the shape is still 80 wide and unrotated —
         * which is indistinguishable from the answer we hope for. A probe that
         * cannot tell "the host means the unrotated box" from "nothing was
         * rotated" would confirm the product's assumption from no evidence, and
         * that is the worse of the two failures by a distance.
         */
        if (!Number.isFinite(rot) || Math.abs(rot - DEG) > 0.5) return { answer: "rotation-did-not-take", detail };
        const rad = (DEG * Math.PI) / 180;
        const extent = Math.abs(W * Math.cos(rad)) + Math.abs(H * Math.sin(rad));
        const full = `${detail} (unrotated ${W}, rotated extent ${extent.toFixed(1)})`;
        if (Math.abs(w - W) < 0.5) return { answer: "unrotated-box", detail: full };
        if (Math.abs(w - extent) < 0.5) return { answer: "POST-ROTATION-BOX", detail: full };
        return { answer: "neither", detail: full };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "named-preset-resolves",
    question: "Does a GeometricShapeType named from our own table actually draw?",
    /**
     * `SYMBOL_PRESET` names are handed straight to the host as
     * `PowerPoint.GeometricShapeType[name]`, and a name this host's enum does
     * not carry resolves to `undefined` — which `addGeometricShape` accepts and
     * draws as a shape with NO GEOMETRY. Invisible, and impossible to explain
     * from the file afterwards. `symbolPreset` guards names missing from OUR
     * table; nothing guards a name missing from the HOST's.
     *
     * It stopped being hypothetical when the hex tile map moved onto `hexagon`:
     * a shipped chart now depends on one of these names resolving, on a node
     * kind (`symbol`) that no round has ever drawn. Asked for every name in the
     * table, not just that one, because the answer is about the enum rather
     * than about any single shape.
     *
     * NO LONGER RESAMPLED, as of 2026-09-01. It carried `resample: true` while
     * it was a new question nobody had an answer for. It has one now, and it is
     * as settled as anything in the archive: `draws` in 14 of 14 rounds, 42 of
     * 42 samples, stable in every one, across six builds — and that answer is in
     * the committed sheet now rather than in `PENDING_QUESTIONS`. Asking it three
     * times a round buys certainty that has already been bought.
     *
     * Put the mark back if `SYMBOL_PRESET` gains a name, which is the change
     * that could make this false again.
     */
    ask: async (ctx) => {
      const wanted = [...new Set(Object.values(SYMBOL_PRESET))].sort();
      const enumeration = PowerPoint.GeometricShapeType as unknown as Record<string, unknown>;
      // READING THE ENUM NEEDS NO SLIDE, so it is asked first and answered even
      // on a host that would refuse the shape below. The names being absent is
      // the whole finding; drawing one is the corroboration.
      const missing = wanted.filter((n) => enumeration[n] === undefined);
      if (missing.length)
        return {
          answer: "MISSING-FROM-ENUM",
          detail: `this host's GeometricShapeType has no ${missing.join(", ")} (asked for ${wanted.join(", ")})`,
        };
      // THE ADD IS SETUP. A host refusing every add must answer
      // `no-scratch-shape` rather than `threw` — see the rotation probe above.
      // The rectangle proves adds work at all; the hexagon is the question.
      await scratchShapes(ctx, [{ left: 140, top: 140, width: 8, height: 8 }]);
      try {
        const shape = probeShapes(ctx).addGeometricShape(enumeration.hexagon as PowerPoint.GeometricShapeType, {
          left: 160,
          top: 140,
          width: 24,
          height: 24,
        });
        // One batch, for the reason the rotation probe gives at length: a read
        // through a sync-old handle is a question about proxy staleness.
        shape.load("width,height");
        await ctx.sync();
        const w = readSize(() => shape.width);
        const h = readSize(() => shape.height);
        return {
          answer: Number.isFinite(w) && w > 0 && Number.isFinite(h) && h > 0 ? "draws" : "drew-nothing",
          detail: `every name in the table is in the enum (${wanted.join(", ")}); hexagon drew ${String(w)}x${String(h)}`,
        };
      } catch (err) {
        return threw(err);
      }
    },
  },
];

/** An error, short enough to sit in a sheet without swamping it. */
function short(err: unknown): string {
  const text = err instanceof Error ? err.message : String(err);
  return text.slice(0, 140);
}

/**
 * Ask this host every question, and come back with a complete sheet.
 *
 * Complete is the operative word. A probe that throws, times out, or wedges
 * contributes its answer and nothing more — the sheet from a misbehaving host
 * is the one worth having, so it must survive misbehaviour. Each question gets
 * its own context and its own budget for that reason.
 */
/**
 * Narrow a probe run to some questions, for a caller that has changed the world.
 *
 * A round asks the whole list ONCE, at the top, before the battery has drawn
 * anything — so any question needing something the battery produces is asked at
 * the only moment it cannot be answered. `shape-resolve-held-slide-proxy` is the
 * case that forced this: it needs an id this host has agreed to name, the
 * self-test's tagged charts are the only such ids on this host, and they do not
 * exist yet when the sheet is built.
 *
 * The answer is NOT to move the sheet after the battery. It is written early on
 * purpose — a battery that takes the tab down must not also lose the probe's
 * answers, which are complete, cheap, and the half more likely to be worth
 * reading. So the sheet stays where it is and the few deferred questions get
 * re-put afterwards, their samples merged in by `mergeHostSheets`.
 */
export interface ProbeRunOptions {
  /** Ask only these question ids. Omit for the whole list. */
  only?: readonly string[];
  /**
   * How many passes to make. Omit for `PROBE_PASSES`.
   *
   * A re-ask wants one: the sheet already holds this question's samples from the
   * three passes it made at the top of the round, and the thing being tested is
   * whether an id changes the answer — not whether the answer is stable, which
   * those earlier samples already speak to.
   */
  passes?: number;
}

export async function runHostProbes(
  source: string,
  build: string,
  opts: ProbeRunOptions = {},
): Promise<HostAnswerSheet> {
  const answers: HostAnswer[] = [];
  /**
   * Every scratch slide this run has made, so all of them come back out.
   *
   * Plural because the run replaces one it has lost. A diagnostic that leaves
   * blank slides in someone's deck is one they stop clicking, and the whole
   * point of replacing a lost slide is to keep going — which means keeping a
   * list rather than overwriting the id of the thing still to be cleaned up.
   */
  /**
   * Every slide this run put in the deck, usable or not.
   *
   * The sweep clamps at "no more than this run added", so this set has to be
   * everything the run LEFT IN THE DECK — including the slides
   * `addScratchSlide` landed, refused, and then could not take back out. It
   * swallowed those, because its return value means "an id you may use" and
   * these are precisely the ids you may not. The 2026-08-11 round left exactly
   * two behind for that reason: the deck grew by 70 while this list held 68, so
   * the clamp correctly refused to touch the other two.
   *
   * Fed by `addScratchSlide`'s callback for that case, and de-duplicating,
   * because a slide it successfully removes must NOT appear here — counting it
   * would make the clean-up owe a delete for a slide that is already gone, and
   * report "some" where the truth is "all".
   */
  const scratchIds: string[] = [];
  const noteScratch = (id: string): void => {
    if (!scratchIds.includes(id)) scratchIds.push(id);
  };
  /**
   * Scratch slides already handed back, before the clean-up runs.
   *
   * EMPTY ON EVERY PATH TODAY — this is the seam a mid-run return needs, and
   * putting it in first is what makes that change small. See
   * `outstandingScratch` for why the clean-up cannot simply count `scratchIds`
   * once anything returns a slide early.
   */
  const returnedEarly = new Set<string>();
  /**
   * Slides that landed and could not be named, so they have no id to record.
   *
   * They are still in the deck and still this run's. The sweep clamps at "no
   * more than this run added", so leaving them out of that total is the clamp
   * refusing to remove slides the run put there — measured at two per run.
   *
   * COUNTED PER SLIDE, not per event, and it was not always. The caller used to
   * call this once however many slides had appeared, and one of those adds has
   * never in the archive produced fewer than two: ten events across nine rounds,
   * `after - before` equal to 2 every time. So the clamp ran one short per
   * event and left a blank slide of ours in the finished deck — round 085's own
   * inventory carries it, `257#3837665135`, zero shapes, listed in `newSlides`.
   */
  let unnamedLeftBehind = 0;
  const noteUnnamed = (): void => {
    unnamedLeftBehind++;
  };
  /**
   * The scratch slide's own history, for `ScratchState`.
   *
   * `generation` counts the slides this run has taken; `used` says whether a
   * question has already been recorded against the current one. Both are
   * bookkeeping over things the run does anyway — nothing here calls the host.
   *
   * `takeScratch` is the ONE place either moves, and every assignment to
   * `scratchId` goes through it. Four sites take a slide (the first one, the
   * recovery at the top of the loop, the replacement after a never-asked, and
   * the replacement for a partner question) and a fifth drops it; a stamp that
   * silently missed one of them would report `reused-slide` for a slide that
   * was in fact brand new, which is the answer this question turns on.
   */
  let scratchGeneration = 0;
  let scratchUsed = false;
  /**
   * The route the slide currently in hand arrived by. `unknown` until the first
   * take, which is honest rather than tidy: it means no take has happened yet,
   * and `record` only ever writes it out alongside a non-null `scratchId`.
   */
  let scratchAcquired: ScratchAcquisition = "unknown";
  const takeScratch = (
    id: string | null,
    { sameSlide = false, via }: { sameSlide?: boolean; via: ScratchAcquisition },
  ): string | null => {
    // THE ROUTE IS RECORDED EVEN WHEN THE TAKE FAILED TO PRODUCE A SLIDE, and
    // then `record` overrides it with `no-slide` for the sample. Two different
    // questions: this variable answers "how did we last try", the sample answers
    // "what was this measured on". Keeping them separate is why a failed replace
    // does not silently label the next sample with the route before it.
    scratchAcquired = via;
    // `sameSlide` is the re-acquire: the deck listing handed back the id the
    // scratch slide is CURRENTLY listed under, and it is the slide this run has
    // been using all along. Bumping the generation there would report the next
    // sample as `fresh-slide` on a slide that is nothing of the kind, and
    // `scratch` is one of the few columns a reader uses to tell "the host
    // refused a new slide" from "the host refused a slide it has had for a
    // while". Everything else about the bookkeeping is unchanged, which is why
    // this still goes through here rather than around it.
    if (id && !sameSlide) {
      scratchGeneration++;
      scratchUsed = false;
    }
    return id;
  };
  // Read BEFORE the first scratch slide is added, so it cannot pick one up.
  // Follow-up questions use it as the control for "was that about the call, or
  // about the slide being new" — the one distinction this project has paid for
  // twice — and nothing may ever write to it.
  //
  // Wrapped, and the test that made me wrap it is the one this whole file is
  // about: a host that goes silent made this throw before a single question had
  // been asked, and the run produced NO SHEET. A control nobody can name costs
  // the follow-up. It must never cost the sheet.
  let durableSlideId: string | undefined;
  try {
    durableSlideId = (await deckSlideIds())?.[0];
  } catch {
    trace("probe", "no durable slide to use as a control — the deck would not list itself", {});
  }
  // The deck's size before this run added anything — the floor the positional
  // sweep may never reach past. Read here rather than in the clean-up because
  // by then the run's own slides are already in the count.
  // IS THE HOST AWAKE, BEFORE A SINGLE QUESTION IS PUT?
  //
  // Three rounds have wedged — 24, 25 and 29 — and each was read as the probe
  // walking into trouble somewhere around question six. Round 29 showed that is
  // not what happens: `shape-add-fresh-slide-proxy`, which has answered `yes` in
  // every round on record, came back SILENT at question four. The host was
  // already not answering before the run had asked it anything worth asking, and
  // nothing in the sheet could say so.
  //
  // The cheapest call there is, timed, before anything else touches the host.
  // `deckSlideIds` is what the run was about to do anyway, so this costs one
  // extra round trip and turns "it wedged" into "it was already slow when we
  // opened", which are different facts and want different responses: a bad host
  // is worth waiting out, a provoked one is worth investigating.
  //
  // The number, not a verdict. A threshold here would be a guess with no rounds
  // behind it; three healthy rounds and three wedged ones are what will set one,
  // and until then the sheet carries the milliseconds and the reader decides.
  const openedAt = Date.now();
  const opened = await deckSlideIds().catch(() => undefined);
  const openedMs = Date.now() - openedAt;
  trace("host", opened ? "answered before the first question" : "did not answer before the first question", {
    ms: openedMs,
    slides: opened?.length,
  });
  const deckAtStart = opened?.length;
  let scratchId = takeScratch(await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch, noteUnnamed), {
    via: "first-add",
  });
  if (scratchId) scratchIds.push(scratchId);
  /**
   * Consecutive questions that could not get an answer out of the host at all.
   *
   * Counted on DEADLINE MISSES, not on `no-scratch-slide`, and the distinction
   * is the whole design. A round on 2026-08-08 answered `no-scratch-slide` four
   * times running — the host had stopped resolving fresh slide ids for about
   * fifteen seconds — and then recovered and answered five more questions,
   * including the two this project cares most about. A breaker keyed on that
   * signal would have thrown those away. Those questions were still being
   * ANSWERED, in two to six seconds; they just could not be set up.
   *
   * A missed deadline is different: the host took the question and never came
   * back inside its budget. Three of those in a row is not a finding about a
   * question, it is a finding about the host — the same argument, and the same
   * number, as the self-test's SICK_LIMIT.
   */
  let mute = 0;
  let abandoned: string | null = null;
  const runStarted = Date.now();
  /**
   * A sample's timestamp, on the run log's clock so it can be read beside the
   * trace lines around it. See `ProbeSample.atMs` for what that cost.
   */
  const stampNow = (): number => traceElapsed() ?? Date.now() - runStarted;
  /**
   * What the run has already watched the host do. Never a guess, never a call.
   */
  /** When this run last SAW each thing. Timestamps, not flags — see `regimeFrom`. */
  let lastRefusalAt: number | undefined;
  let lastSlideTroubleAt: number | undefined;
  let lastGoodAt: number | undefined;
  /** Only the shape collection answering — see `regimeFrom` for why it is separate. */
  let lastCollectionGoodAt: number | undefined;
  const regimeNow = (at: number): HostRegime =>
    regimeFrom({ at, lastRefusalAt, lastSlideTroubleAt, lastGoodAt, lastCollectionGoodAt });
  /** Answers that mean the shape collection has stopped answering. */
  const REFUSAL = new Set(["unreadable", "short-0", "not-listed"]);
  /**
   * One row per question, however many times it is asked.
   *
   * `answer` keeps exactly the meaning it has always had — the first REAL
   * answer, with a never-asked sentinel standing only until a later pass
   * replaces it. That is what the second pass has always done to a lost row;
   * this generalises it to three passes without the diff, the contract gate or
   * any reader of a sheet noticing. The repeats go to `samples`.
   */
  const rows = new Map<string, HostAnswer>();
  const record = (row: HostAnswer, pass: number): HostAnswer => {
    const atMs = stampNow();
    // Each observation stamped with WHEN, so a later reading can tell a host
    // that is refusing now from one that refused a minute ago.
    const aboutTheCollection = /^shapes?-/.test(row.id);
    if (aboutTheCollection && REFUSAL.has(row.answer)) lastRefusalAt = atMs;
    else if (aboutTheCollection && !NOT_ASKED.has(row.answer)) {
      lastGoodAt = atMs;
      lastCollectionGoodAt = atMs;
    }
    if (row.answer === "no-scratch-slide") lastSlideTroubleAt = atMs;
    else if (!NOT_ASKED.has(row.answer)) lastGoodAt = atMs;
    // Read BEFORE `scratchUsed` is set below, so this sample describes the
    // slide as it was when this question ran, not as the next one will find it.
    const scratch: ScratchState = !scratchId
      ? "no-slide"
      : scratchUsed
        ? "reused-slide"
        : scratchGeneration <= 1
          ? "first-slide"
          : "fresh-slide";
    // THE ROUTE, alongside the state. `scratchAcquired` is set by `takeScratch`
    // at whichever of its six call sites actually produced the slide this run is
    // holding; with no slide the route is `no-slide` regardless of what the last
    // successful take recorded, because the sample was not taken on that slide.
    const acquired: ScratchAcquisition = scratchId ? scratchAcquired : "no-slide";
    const sample: ProbeSample = { answer: row.answer, pass, atMs, regime: regimeNow(atMs), scratch, acquired };
    // Marked here rather than at the ask, because `record` is the one path every
    // question's outcome goes through — a retry, a partner and a second-pass
    // rescue all land here, and each of them genuinely has used the slide.
    if (scratchId) scratchUsed = true;
    const seen = rows.get(row.id);
    if (!seen) {
      const created: HostAnswer = { ...row, samples: [sample] };
      rows.set(row.id, created);
      answers.push(created);
      return created;
    }
    seen.samples = [...(seen.samples ?? []), sample];
    // A NAMED answer replaces a never-asked or an unnameable one, and nothing
    // replaces a named one: the first NAMED answer is the row's answer, so a
    // sheet means today what it meant yesterday. Disagreement between samples is
    // `stable`'s job to report, not `answer`'s to hide.
    //
    // `other` used to count as real and lock the row — see `UNINFORMATIVE`. That
    // is how `tags-add-same-key-twice` read `other` in 83 of 86 rounds while two
    // of them had `overwrites` sitting in their samples.
    if (answerRank(row.answer) > answerRank(seen.answer)) {
      seen.answer = row.answer;
      seen.ms = row.ms;
      seen.detail = row.detail;
    }
    return seen;
  };
  try {
    /** Questions the current pass will put. Pass 1 is always the whole list. */
    /**
     * The questions this run is allowed to put, before any per-pass shortening.
     *
     * An `only` naming nothing real would silently ask the entire list, which is
     * the opposite of what a caller asking for two questions wants — so an empty
     * selection stays empty and the run comes back with an empty sheet.
     */
    const asking = opts.only ? PROBES.filter((pr) => opts.only!.includes(pr.id)) : PROBES;
    const passes = opts.passes ?? PROBE_PASSES;
    let list: Probe[] = [...asking];
    for (let pass = 1; pass <= passes; pass++) {
      if (pass > 1) {
        // Decided from the pass just finished, not from the run's whole
        // history: pressure comes and goes here in fifteen-second windows, and
        // a run that struggled early and recovered should get its repeats.
        const put = [...rows.values()].filter((r) => r.samples?.some((x) => x.pass === pass - 1));
        const lostLast = put.filter((r) => NOT_ASKED.has(r.samples!.find((x) => x.pass === pass - 1)!.answer)).length;
        const pressured = put.length > 0 && lostLast / put.length > PASS_PRESSURE_LIMIT;
        list = pressured ? asking.filter((pr) => pr.resample) : [...asking];
        trace("probe", "starting another pass over the questions", {
          pass,
          asking: list.length,
          of: asking.length,
          why: pressured ? `${lostLast} of ${put.length} were never put last pass — shortlist only` : "full list",
        });
        if (!list.length) break;
      }
      for (const probe of list) {
        const started = Date.now();
        // Stop asking a host that has stopped answering.
        //
        // The probe had no rung for this at all while the self-test has had one
        // for weeks, and PowerPoint's own "Sorry, we ran into a problem" dialog is
        // what exposed the gap: the host was gone, the pane's timer kept counting,
        // and the run went on putting questions to a dead document — every one of
        // them spending its full budget to record an answer about nothing.
        if (!abandoned && mute >= PROBE_MUTE_LIMIT) {
          abandoned = `${mute} questions in a row got no answer out of the host`;
          trace("probe", "giving up on the host", { after: answers.length, why: abandoned });
        }
        if (abandoned) {
          record(
            {
              id: probe.id,
              question: probe.question,
              answer: "not-asked",
              ms: 0,
              detail: `not reached — ${abandoned}`,
            },
            pass,
          );
          continue;
        }
        const deadlinesBefore = deadlinesFired;
        trace("probe", "asking", { id: probe.id });
        let result: { answer: string; detail?: string; why?: "gone" | "silent" };
        // One more attempt at a slide, every question, rather than writing the
        // rest of the sheet off.
        //
        // `scratchId` goes null when a replacement also failed to resolve, and
        // nothing used to set it back — so a single bad moment cost every
        // question after it, all of them recorded `no-scratch-slide` as though
        // the host had been asked. That is the same shape as the failure this
        // whole rung exists to prevent, one level up: an answer that describes
        // the run's own state rather than the host's.
        //
        // A host that genuinely will not keep a slide pays one `addScratchSlide`
        // per question for the honest answer, which is what the budget is for.
        if (!scratchId) {
          const recovered = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch, noteUnnamed);
          if (recovered) {
            noteScratch(recovered);
            scratchId = takeScratch(recovered, { via: "recovered" });
            trace("probe", "took another scratch slide after giving up on the last", { id: probe.id });
          }
        }
        if (!scratchId && !probe.noSlideNeeded) {
          result = { answer: "no-scratch-slide", detail: "the host would not add a slide to work on" };
        } else if (probe.noSlideNeeded) {
          // Not gated on a scratch slide at any level — not on having one, and
          // not on its liveness. The id is passed through because `ProbeContext`
          // carries one, but nothing here resolves it. See `Probe.noSlideNeeded`.
          result = await ask(probe, scratchId ?? "", durableSlideId);
        } else {
          result = await ask(probe, scratchId as string, durableSlideId);
          // A slide that stopped resolving costs one question, not the sheet.
          //
          // It cost thirteen once: a real host lost the scratch slide after the
          // first probe and the remaining thirteen questions each reported the
          // same failure as if it were their own answer. Thirteen apparent
          // divergences, one cause, and none of the questions actually asked.
          // Replacing the slide is the difference between a sheet that reports a
          // finding and a sheet that IS the finding.
          //
          // Either kind of never-asked gets the replacement: a slide the host
          // will not resolve, and a slide that resolves and will not take a
          // shape. The second is a weaker reason to suspect the slide, but the
          // cost is one add and one question, and the alternative is a sheet
          // that gives up on eight questions because one slide went bad.
          if (NOT_ASKED.has(result.answer)) {
            // ASK THE DECK BEFORE BUYING A SLIDE, when the held id is what went
            // bad. See `scratchSlideIdByPosition`: the id this run holds stops
            // resolving while the slide is still in the deck, and round 253
            // measured the bill — 61 of 64 replacements were `gone`, and the
            // clean-up swept 107 slides out of a deck that started with one,
            // recovering none of them by id.
            //
            // ONLY on `gone`. A `no-scratch-shape` means the slide resolved
            // perfectly well and refused a shape, so re-acquiring the same slide
            // asks the same question of the same slide and learns nothing.
            if (result.why !== "gone") {
              // SAID OUT LOUD, because the alternative is an untestable
              // optimisation. Skipping the listing here is the whole value of
              // gating on the cause — a round refuses shape adds often, and
              // spending a deck listing on each would trade one cheap call for
              // another. Without this line nothing distinguishes "we decided not
              // to look" from "we looked and the id was the same".
              trace("probe", "did not re-acquire — the slide resolved and the shape is what failed", {
                id: probe.id,
                answer: result.answer,
              });
            } else if (typeof deckAtStart === "number") {
              const settled = await scratchSlideIdByPosition(deckAtStart);
              if (settled && settled !== scratchId) {
                const onSettled = await ask(probe, settled, durableSlideId);
                trace("probe", "re-acquired the scratch slide by position instead of adding one", {
                  id: probe.id,
                  was: scratchId,
                  now: settled,
                  answer: onSettled.answer,
                  worked: !NOT_ASKED.has(onSettled.answer),
                });
                if (!NOT_ASKED.has(onSettled.answer)) {
                  // Recorded under the id the deck actually lists, which is also
                  // the id the clean-up would need to delete it by.
                  noteScratch(settled);
                  scratchId = takeScratch(settled, { sameSlide: true, via: "re-acquired" });
                  result = onSettled;
                }
              }
            }
          }
          // Re-checked, because the re-acquire above may have answered it.
          if (NOT_ASKED.has(result.answer)) {
            const replacement = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch, noteUnnamed);
            if (replacement) {
              noteScratch(replacement);
              scratchId = takeScratch(replacement, { via: "replaced" });
              trace("probe", "replaced the scratch slide", {
                id: probe.id,
                scratchId: replacement,
                after: result.answer,
                // WHICH KIND of unavailable, so the archive can say whether this
                // is a slide being deleted or a proxy going unanswered. Only the
                // second is the case the update path already knows how to
                // recover from without adding a slide at all.
                // `unrecorded` must mean UNKNOWN, not "known and unsaid".
                // A `no-scratch-shape` has no `why` because it never reached
                // `ScratchSlideUnavailable` — the slide was fine and the shape
                // was refused — and calling that unrecorded put 12 events a
                // round into the bucket reserved for rounds that predate the
                // instrument.
                why: scratchReplacementWhy(result),
              });
              const retry = await ask(probe, replacement, durableSlideId);
              // DID BUYING THE SLIDE HELP? Counted, because the decision to buy
              // one rests on it and nothing has ever measured it.
              //
              // Round 255 bought 18 slides and every one of them followed a
              // SHAPE refusal, not a slide one — the code buys on either, on the
              // theory that a slide which will not take a shape may itself be
              // bad. That theory may be right; it has simply never been checked,
              // and it is the same shape as the re-read whose successes went
              // uncounted for 227 rounds while its failures were tallied.
              trace("probe", "the replacement slide answered", {
                id: probe.id,
                after: result.answer,
                answer: retry.answer,
                rescued: !NOT_ASKED.has(retry.answer),
              });
              // Only adopt a retry that actually got somewhere. A second failure
              // on a brand-new slide is a stronger statement than the first, and
              // overwriting it with a never-asked would hide that.
              if (!NOT_ASKED.has(retry.answer)) result = retry;
              // Give up on the SLIDE only when the slide is what failed. A host
              // that resolves slides and refuses shapes has nothing wrong with
              // its slides, and writing off the scratch slide there would turn
              // one refusal into "no-scratch-slide" for every question after it —
              // the very noise this whole rung exists to prevent.
              else if (retry.answer === "no-scratch-slide") scratchId = null;
            } else if (result.answer === "no-scratch-slide") {
              scratchId = null;
            }
          }
        }
        const askedMs = Date.now() - started;
        record({ id: probe.id, question: probe.question, ms: askedMs, ...result }, pass);
        // THIS PASS's answer, not the row's.
        //
        // `record` returns the accumulated ROW, whose `answer` is deliberately
        // the FIRST real answer so a sheet means today what it meant yesterday.
        // Every line below wants the opposite — what the host said just now —
        // and for three passes they all read the row instead. The 2026-08-11
        // round is the proof: the samples say `threw, yes, yes` for
        // `shape-add-held-slide-proxy` while the run log says `answered: threw`
        // three times, with an identical `ms: 637` on asks forty-five seconds
        // apart. A reader of that log concludes the opposite of what the host
        // did. Same family as `batch committed`: a line asserting something it
        // does not know.
        trace("probe", "answered", { id: probe.id, answer: result.answer, ms: askedMs, pass });
        // A question that wrecks the slide on its way out gives up the slide, so
        // the next one starts on a fresh handle instead of inheriting the damage.
        //
        // Dropped AFTER the answer is recorded, and only on `burnsTheSlide`: this
        // question succeeded, so nothing else here would have replaced anything.
        // The `if (!scratchId)` at the top of the loop takes the new one, which is
        // the same path a lost slide already uses — no second mechanism.
        if (probe.burnsTheSlide && scratchId) {
          trace("probe", "giving up the scratch slide this question wrecked", { id: probe.id, scratchId });
          scratchId = null;
        }
        // Reset on ANY question that came back inside its budget, including one
        // that could not be set up: a host still refusing questions promptly is a
        // host still talking to us.
        mute = deadlinesFired > deadlinesBefore ? mute + 1 : 0;
        // The partner question, in the same run, when the answer admits two
        // readings. Never asked when the question was never PUT: a follow-up to
        // `no-scratch-shape` would be a second question about the probe's own
        // setup, dressed as a fact about the host.
        const follow = probe.follow;
        // Gated on THIS pass's answer for the same reason. Reading the row here
        // meant a partner conditioned on a later pass's answer could never fire
        // on it, and one conditioned on the first answer fired on every pass
        // regardless of what the host had just said — a follow-up asked
        // "because X answered yes" when X had, that time, answered no.
        if (follow && !NOT_ASKED.has(result.answer) && follow.when(result.answer)) {
          const at = Date.now();
          trace("probe", "asking the partner question", { after: probe.id, id: follow.probe.id, was: result.answer });
          let r = scratchId
            ? await ask(follow.probe, scratchId, durableSlideId)
            : { answer: "no-scratch-slide", detail: "the host would not add a slide to work on" };
          // The same replacement the main loop gets, and for the same reason.
          //
          // A follow-up went unasked in three consecutive rounds:
          // `getitem-durable-slide` reads the DURABLE slide and never touches the
          // scratch one, yet `ask` liveness-checks the scratch slide first, so a
          // question with no use for that slide was refused for its absence. On a
          // host that loses a scratch slide after almost every question — this one
          // replaced seventeen in one run — that is not an edge case, it is the
          // normal path.
          if (NOT_ASKED.has(r.answer)) {
            const replacement = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch, noteUnnamed);
            if (replacement) {
              noteScratch(replacement);
              scratchId = takeScratch(replacement, { via: "replaced" });
              trace("probe", "replaced the scratch slide for a partner question", {
                id: follow.probe.id,
                scratchId: replacement,
                after: r.answer,
                // THE SAME FORMULA THE MAIN LOOP USES, and it was missing here.
                // `WHAT THE PROBE SPENDS ON SLIDES IT COULD NOT KEEP` pools both
                // traces, so every partner replacement landed in `unrecorded` —
                // 68 of the last ten rounds' 248 replacements, 27%, saying
                // nothing about why a slide was bought. The main line has
                // carried this since the `unrecorded`-means-UNKNOWN fix and this
                // one was simply never given it.
                why: scratchReplacementWhy(r),
              });
              const retry = await ask(follow.probe, replacement, durableSlideId);
              // DID BUYING IT HELP, asked here too. The main loop's own note
              // says this measurement is what the decision to buy a slide rests
              // on — "the same shape as the re-read whose successes went
              // uncounted for 227 rounds while its failures were tallied" — and
              // the partner path buys about seven slides a round without it.
              trace("probe", "the replacement slide answered", {
                id: follow.probe.id,
                after: r.answer,
                answer: retry.answer,
                rescued: !NOT_ASKED.has(retry.answer),
                partner: true,
              });
              // Same rule as the main loop: only adopt a retry that got somewhere.
              if (!NOT_ASKED.has(retry.answer)) r = retry;
            }
          }
          record(
            {
              id: follow.probe.id,
              question: follow.probe.question,
              ms: Date.now() - at,
              ...r,
              // Which answer triggered it, in the sheet itself. Two rows that do
              // not say they are a pair are two unrelated facts to whoever reads
              // them a week later.
              detail: `asked because ${probe.id} answered "${result.answer}" — ${follow.because}${r.detail ? `; ${r.detail}` : ""}`,
            },
            pass,
          );
          trace("probe", "partner answered", { id: follow.probe.id, answer: r.answer });
        }
      }
      // The breaker is a fact about the host, not about a pass: once the run has
      // given up there is nothing for a later pass to sample.
      if (abandoned) break;
    }

    // A LAST pass over the questions this run never managed to put, at the
    // END of the run rather than beside the failure.
    //
    // The retry above already takes a fresh slide and asks again — and it is not
    // enough. Two consecutive rounds lost fourteen of twenty-seven and thirteen
    // of twenty-eight questions to `no-scratch-slide`, with twenty-one slide
    // replacements between them: the retries fired, and failed the same way.
    //
    // Time is what they were missing. This host's ability to resolve a freshly
    // added slide comes and goes in windows of roughly fifteen seconds, so a
    // retry issued immediately lands INSIDE the window that just refused. A pass
    // at the end is a different sample, and both rounds show the recovery
    // happening within the same run: the 2026-08-09 round answered questions at
    // positions 17, 18, 22, 24 and 26 after losing 10 through 16.
    //
    // Which questions get lost is close to random — the two rounds disagree on
    // six of them — so this is not a fix for one question. It is the difference
    // between a sheet that answers half its questions and one that answers as
    // many as the host will take.
    //
    // Skipped entirely when the breaker fired: `not-asked` means the run gave up
    // on a host that had stopped answering, and asking it thirteen more times is
    // the behaviour the breaker exists to stop.
    const lost = abandoned ? [] : answers.filter((a) => NOT_ASKED.has(a.answer));
    if (lost.length) {
      trace("probe", "second pass over the questions that were never put", { count: lost.length });
      let rescued = 0;
      let noSlide = 0;
      for (const entry of lost) {
        const probe = PROBE_BY_ID.get(entry.id);
        if (!probe) continue;
        if (mute >= PROBE_MUTE_LIMIT) {
          trace("probe", "stopping the second pass — the host is not answering", { at: entry.id });
          break;
        }
        const replacement = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch, noteUnnamed);
        // One refused slide is not a dead host, and treating it as one made
        // this whole pass theatre.
        //
        // It used to `break` here, silently. Round 15 is what that looks like
        // from outside: `second pass over the questions that were never put
        // {count: 10}`, one slide attempt two seconds later, `scratch slide
        // landed but its id will not resolve`, and then nothing — no rescues,
        // and no line saying why. Ten questions decided by one coin flip, on a
        // host whose slide resolution is known to flap in fifteen-second
        // windows, and a log that could not tell "tried them all and failed"
        // from "gave up on the first".
        //
        // The main loop has always taken one add attempt PER QUESTION, which is
        // the right shape against that host. The sweep took one for the whole
        // list. They match now, bounded by CONSECUTIVE failures so a genuinely
        // dead host still costs a few adds rather than one per lost question.
        if (!replacement) {
          if (++noSlide >= PROBE_MUTE_LIMIT) {
            trace("probe", "stopping the second pass — no scratch slide to be had", { at: entry.id, tried: noSlide });
            break;
          }
          continue;
        }
        noSlide = 0;
        noteScratch(replacement);
        scratchId = takeScratch(replacement, { via: "replaced" });
        const deadlinesBefore = deadlinesFired;
        const started = Date.now();
        const retry = await ask(probe, replacement, durableSlideId);
        mute = deadlinesFired > deadlinesBefore ? mute + 1 : 0;
        // Only an answer replaces a never-asked. A second failure is the same
        // fact the first one already recorded, and overwriting the row with it
        // would lose the original `ms` for no gain.
        // Recorded as a SAMPLE like every other ask, so the rescue is visible
        // as the late sample it is rather than as an answer with no provenance.
        // `record` applies the same rule this block always did — only a real
        // answer replaces a never-asked — so the row reads exactly as before.
        const retriedAt = stampNow();
        entry.samples = [
          ...(entry.samples ?? []),
          {
            answer: retry.answer,
            pass: PROBE_PASSES + 1,
            atMs: retriedAt,
            regime: regimeNow(retriedAt),
            // The second pass takes a REPLACEMENT slide for every question it
            // retries, so this sample is by construction on a slide nothing
            // else has used — the one place the stamp is known rather than
            // derived. `scratchGeneration` is past 1 by now in any run that
            // reached here, so `fresh-slide` rather than `first-slide`.
            scratch: "fresh-slide",
            // Same construction, same certainty: the slide under this sample was
            // taken by the replacement path a few lines above, so the route is
            // known outright rather than read off `scratchAcquired`.
            acquired: "replaced",
          },
        ];
        if (!NOT_ASKED.has(retry.answer)) {
          entry.answer = retry.answer;
          entry.ms = Date.now() - started;
          entry.detail = `${retry.detail ? `${retry.detail}; ` : ""}answered on a second pass at the end of the run`;
          trace("probe", "second pass answered", { id: probe.id, answer: retry.answer });
          rescued++;
        }
      }
      // What the pass was WORTH, every time, including nothing. A sweep that
      // rescued none and a sweep that never ran read identically in round 15's
      // log, and the difference is the whole question of whether to keep it.
      trace("probe", "second pass finished", { of: lost.length, rescued });
    }
  } finally {
    // The scratch slides go back whatever happened. A diagnostic that litters
    // the user's deck is one they will stop running.
    //
    // And it SAYS how that went, because for as long as it did not, this
    // failing was something the owner discovered by opening a deck. One run
    // left 21 blank slides behind — a previous one, fourteen — and the sheet
    // that run produced recorded neither. `deleteSlideById` already answers
    // honestly (it re-reads the deck rather than trusting a queued delete);
    // the boolean was simply thrown away here.
    //
    // Filed as an answer rather than a trace line so it travels the same road
    // as every other fact about this host: into the sheet, through the diff,
    // and against the fake — which returns every slide it is given, so a host
    // that does not diverges and is reported without anyone remembering to
    // look.
    const cleanupStarted = Date.now();
    // COUNT the deck either side, and let the count outrank the per-slide
    // booleans — see `slidesActuallyReturned` for what that is worth.
    const idsBefore = await deckSlideIds().catch(() => undefined);
    const deckBefore = idsBefore?.length;
    // Does the deck still LIST the ids we are about to delete by?
    //
    // The one number that separates the two readings of a clean-up that
    // reports 45 deletes into a deck that does not shrink. Either the ids are
    // stable and the deletes simply fail, or the id we captured at add time is
    // not the id the deck answers to any more — in which case
    // `deleteSlideByPosition`'s `indexOf(id) < 0` reads "already gone" for a
    // slide sitting right there, and delete-by-id is not merely failing, it is
    // structurally impossible.
    //
    // 2026-08-11 made this worth asking rather than assuming: the scratch ids
    // read `4123571115#123571113` while the deck listed `256#109857222`
    // through `314#195537992`, and both come from the SAME `slideIds()`
    // projection minutes apart. Reasoning cannot choose between "the host
    // renumbers" and "two readers disagree"; one count can.
    const outstanding = outstandingScratch(scratchIds, returnedEarly);
    const stillListed = idsBefore ? outstanding.filter((id) => idsBefore.includes(id)).length : undefined;
    let returned = 0;
    for (const id of outstanding) if (await deleteSlideById(id).catch(() => false)) returned++;
    let deckAfter = (await deckSlideIds().catch(() => undefined))?.length;
    const byId = slidesActuallyReturned({
      claimed: returned,
      added: outstanding.length,
      deckBefore,
      deckAfter,
    });
    // EVERY SLIDE THIS RUN PUT IN THE DECK, named once and shared by everything
    // downstream that needs a denominator.
    //
    // It was written out twice, 24 lines apart, and the two copies disagreed:
    // the sweep was licensed with `outstanding.length + unnamedLeftBehind` and
    // the verdict was scored against `outstanding.length`. The numerator then
    // counted slides the denominator did not know existed, `left` came out
    // NEGATIVE, and 25 of 140 archived rounds graded a clean handback as `some`
    // and printed sentences like "98 of 94 scratch slide(s) deleted; -4 never
    // landed".
    //
    // ONE BINDING RATHER THAN A TEST THAT THE TWO MATCH. A guard comparing two
    // expressions can only fail after someone has already written the second
    // one; there is nothing to keep in step if there is only one. `byId` above
    // keeps its own `outstanding.length` deliberately — its question is whether
    // delete-by-id returned the slides it could NAME, and the unnamed ones have
    // no id for it to try.
    const added = outstanding.length + unnamedLeftBehind;
    // The sweep, and ONLY when delete-by-id left something behind. On a host
    // where the ids work this never runs and nothing about the old path
    // changes; on this one it is the only thing that can work, because the ids
    // it would need are not ids the deck lists.
    let swept = 0;
    if (byId.left > 0) {
      const plan = positionalSweepPlan({
        deckAtStart,
        deckNow: deckAfter,
        // The unnamed ones have no id to delete by, so they are not in
        // `outstanding` — but they ARE in the deck, and the clamp is what
        // decides whether the sweep may reach them. Same binding the verdict
        // below is scored against; see `added`.
        added,
        alreadyDeleted: byId.actually,
      });
      if (plan) {
        trace("probe", "sweeping the run's own slides by position", {
          ...plan,
          deckAtStart,
          deckNow: deckAfter,
          why: "delete-by-id left slides behind and this host does not list the ids they were added under",
        });
        swept = await deleteTrailingSlides(plan.from, plan.count);
        deckAfter = (await deckSlideIds().catch(() => undefined))?.length;
        trace("probe", "swept", { swept, deckNow: deckAfter });
      } else {
        trace("probe", "no positional sweep — the deck does not support one safely", {
          deckAtStart,
          deckNow: deckAfter,
          added: outstanding.length,
        });
      }
    }
    // Re-derived from the deck AFTER the sweep, so the number reported is what
    // the deck lost overall rather than what either mechanism claimed.
    const { actually, left, shrankBy } = slidesActuallyReturned({
      claimed: returned + swept,
      added,
      deckBefore,
      deckAfter,
    });
    answers.push({
      id: SCRATCH_CLEANUP_ID,
      question: "Does the host give back the scratch slides this probe added?",
      answer: scratchCleanupAnswer({ addedAny: scratchIds.length > 0, left, actually }),
      ms: Date.now() - cleanupStarted,
      detail:
        // `added`, not `scratchIds.length` — the same correction as `left`
        // above, and the half a reader sees first. "98 of 94" is what an
        // unnamed add looks like in a denominator that does not count it.
        `${actually + returnedEarly.size} of ${added} scratch slide(s) deleted` +
        // "left in the deck" is a claim about the deck, and round 029 caught it
        // being false: `left=73` beside `deckBefore=1 deckAfter=1
        // stillListed=0`. Seventy-three slides were counted as abandoned while
        // the deck stood at one the whole time and listed none of their ids —
        // they never landed at all. Five attempts were reverted trying to stop a
        // leak that was this sentence.
        //
        // So the deck decides which sentence gets used. `stillListed === 0` with
        // a deck no bigger than it started is the host having accepted an add
        // whose slide never appeared, which is a different bug from a clean-up
        // that missed something, and wants a different fix.
        //
        // NOT COVERED BY THE SUITE, and said out loud rather than left to be
        // assumed. `left in the deck` has a test (`reports the scratch slides it
        // could NOT give back`, via `refuseSlideDelete`) and this branch does
        // not: reaching it needs a host that ACCEPTS an add and then never lists
        // the slide, and `swallowAdds` does not get there — a test written
        // against it asserted nothing and was deleted rather than kept as
        // decoration. The evidence for this branch is round 029 itself.
        // EXTRACTED so the branch above can finally be tested. It was inline in
        // a 200-line probe reachable only through a host that accepts an add and
        // then never lists the slide, which is why it went 25 rounds printing a
        // negative number inside a sentence that means the opposite.
        scratchLeftSentence({ left, stillListed, deckBefore, deckAfter }) +
        // Only when some came back before the clean-up ran. Silent otherwise, so
        // the sentence a reader has seen a hundred times does not change shape
        // for a number that is zero.
        (returnedEarly.size ? `; ${returnedEarly.size} of them handed back during the run` : "") +
        // Only when the deck lost LESS than the deletes claimed. That clause was
        // written when the by-id path was the only one, where any disagreement
        // meant an over-claim — and the sweep made it read backwards on its
        // first outing: `the deletes reported 0 but the deck only shrank by 68`,
        // which is the sweep working, described as a shortfall. A line that says
        // the opposite of what happened is the `batch committed` mistake in a
        // new place, so it is conditioned on the direction it describes.
        (shrankBy !== undefined && shrankBy < returned
          ? ` (the deletes reported ${returned} but the deck only shrank by ${shrankBy})`
          : "") +
        // "RECOGNISED", NOT "STILL LISTS". This number is measured BEFORE the
        // deletes run — deliberately, and the comment at its assignment says
        // why: it is what separates "the deletes fail" from "the id we captured
        // at add time is not the id the deck answers to any more". The word
        // `still` made it read as a leftover count taken afterwards, and printed
        // beside a deck that ends at one slide it said the deck was holding
        // twenty of them.
        (stillListed !== undefined && stillListed < outstanding.length
          ? `; the deck recognised ${stillListed} of ${outstanding.length} of these ids before the sweep`
          : "") +
        (swept ? `; ${swept} removed by a positional sweep after delete-by-id took none` : ""),
    });
    trace("probe", "gave the scratch slides back", {
      returned,
      swept,
      left,
      deckBefore,
      deckAfter,
      shrankBy,
      stillListed,
      // A SAMPLE of both id lists, side by side. Reconstructing the id-space
      // mismatch by hand cost two rounds — the scratch ids read
      // `4123571115#123571113` while the deck listed `256#109857222`, and
      // nothing in the file put those two facts next to each other. Three of
      // each is enough to see it and small enough never to crowd the buffer.
      heldIds: scratchIds.slice(0, 3),
      deckIds: (idsBefore ?? []).slice(0, 3),
    });
  }
  // Said by the sheet, from its own samples — the fact `UNSTABLE_ANSWERS` was
  // assembled by hand across ten rounds to say. Only REAL answers count: a
  // question that was put once and refused twice is not unstable, it is a
  // question this run mostly could not ask.
  for (const row of answers) {
    const verdict = stabilityOf(row.samples);
    if (verdict !== undefined) row.stable = verdict;
    const support = thinSupport(row.samples);
    if (support) row.support = support;
  }
  return {
    kind: HOST_ANSWERS_KIND,
    source,
    build,
    requirementSets: requirementSets(),
    // How the host was BEFORE the first question — see the call that measures it.
    opened: { ms: openedMs, answered: opened !== undefined },
    answers,
  };
}

/**
 * Questions this sheet lost for want of a shape it was allowed to name.
 *
 * `no-scratch-shape` is a SETUP failure, not a finding: it means the question
 * never ran. This host declines to name the scratch shapes the probe mints, so
 * on it the string means "ask me again when you have an id I will admit to".
 * After the battery, the round has one — see `namedShape`.
 */
export function deferredForLackOfShape(sheet: HostAnswerSheet): string[] {
  return sheet.answers
    .filter(
      (r) =>
        // NEVER REALLY ANSWERED. A question that got a real answer on any pass
        // has its answer, and a re-ask could only add samples to something
        // already settled.
        NOT_ASKED.has(r.answer) &&
        // AND blocked for want of a SHAPE on at least one pass — read from the
        // samples, not from `answer`.
        //
        // `answer` holds the FIRST weak answer and nothing weak displaces it, so
        // a question that failed once on the slide and three times on the shape
        // reads `no-scratch-slide` forever. Round 245 is exactly that:
        // `[no-scratch-slide, no-scratch-shape, no-scratch-shape,
        // no-scratch-shape]`, and this function looked only at the row, found no
        // match, and made no re-ask at all — so the round produced four samples
        // where every other round since 241 produced five.
        (r.answer === "no-scratch-shape" || (r.samples ?? []).some((x) => x.answer === "no-scratch-shape")),
    )
    .map((r) => r.id);
}

/**
 * Whether the round can put its deferred questions yet, and which.
 *
 * Pure, because it is the part of the re-ask that was got WRONG the first time.
 * The original fix gave the probe a way to use a named id and left the ask at
 * the top of the round, where no such id exists — so it changed nothing and the
 * next round would have been read as the probe failing rather than the caller
 * asking too early. A decision that subtle belongs somewhere a test can reach.
 *
 * `skipped` rather than silence: a round that draws nothing this host will name
 * still deferred real questions, and that is worth a trace line. It is not a
 * failure — `no-scratch-shape` is then simply true.
 */
export function reaskPlan(sheet: HostAnswerSheet, named: unknown): { ask: string[]; skipped: string[] } {
  const deferred = deferredForLackOfShape(sheet);
  return named ? { ask: deferred, skipped: [] } : { ask: [], skipped: deferred };
}

/**
 * Fold a re-ask back into the sheet that deferred the question.
 *
 * Pure, and it returns a new sheet: the caller has already written the original
 * into the crash store, and a merge that mutated it would rewrite a record whose
 * entire purpose is to survive whatever happens next.
 *
 * The promotion rule is the one `record` uses within a run and for the same
 * reason — the first NAMED answer is the row's answer, so a sheet means today
 * what it meant yesterday. Here that rule is what makes the re-ask worth doing
 * at all: `no-scratch-shape` is weak, so a real answer displaces it, and a
 * question already answered at the top of the round keeps the answer it gave.
 */
export function mergeHostSheets(base: HostAnswerSheet, extra: HostAnswerSheet): HostAnswerSheet {
  // Extra's passes are numbered from 1 again, because it was a separate run.
  // Left alone, a merged row would read `pass 1, 2, 3, 1` and every reader —
  // including this file's own per-pass shortlist arithmetic — would take the
  // last sample for the first. Offsetting keeps the samples in the order they
  // were actually taken, which is the one thing a pass number is for.
  const after = Math.max(0, ...base.answers.flatMap((r) => (r.samples ?? []).map((x) => x.pass)));
  const shifted = new Map(
    extra.answers.map((r) => [r.id, { ...r, samples: (r.samples ?? []).map((x) => ({ ...x, pass: x.pass + after })) }]),
  );
  const merged = base.answers.map((row) => {
    const again = shifted.get(row.id);
    if (!again) return row;
    shifted.delete(row.id);
    const samples = [...(row.samples ?? []), ...(again.samples ?? [])];
    const next: HostAnswer =
      answerRank(again.answer) > answerRank(row.answer)
        ? { ...row, answer: again.answer, ms: again.ms, detail: again.detail, samples }
        : { ...row, samples };
    const verdict = stabilityOf(samples);
    if (verdict !== undefined) next.stable = verdict;
    else delete next.stable;
    const support = thinSupport(samples);
    if (support) next.support = support;
    else delete next.support;
    return next;
  });
  // A re-ask that produced a row the sheet has never seen is still an answer
  // this host gave, and dropping it would lose it silently.
  return { ...base, answers: [...merged, ...shifted.values()] };
}

/**
 * One question, and whatever came back — including "I never got to ask".
 *
 * A probe that could not reach its question must not answer as though it had.
 * `"threw"` is a legitimate answer to several of these questions and the diff
 * compares it against the fake's; a scratch slide that vanished underneath a
 * probe reported as `"threw"` therefore reads as a genuine host divergence, and
 * that is precisely how thirteen of one real sheet's fourteen answers came to
 * be noise. `"no-scratch-slide"` is not in any probe's own vocabulary, so it
 * can never be mistaken for one.
 */
async function ask(
  probe: Probe,
  scratchId: string,
  durableSlideId?: string,
): Promise<{ answer: string; detail?: string; why?: "gone" | "silent" }> {
  awaitingSetupShapes = false;
  try {
    const answered = await withProbeContext(scratchId, PROBE_BUDGET_MS, probe.ask, durableSlideId, probe.noSlideNeeded);
    /**
     * AN ANSWER, ALWAYS. The return type promises one; on 2026-09-02 it did
     * not deliver, and the round it killed cost half an experiment.
     *
     * WHERE IT THREW, established from the record rather than guessed: the
     * caller does `record({ …, ...result })` and then
     * `trace("probe", "answered", { answer: result.answer })`. The spread
     * tolerates `undefined` silently — `{...undefined}` is `{}` — so `record`
     * survives and the TRACE is what dies. Round 360's attempt 2 shows exactly
     * that shape: `probe asking id=getitemornullobject-missing` with no
     * matching `answered` line, then `Cannot read properties of undefined
     * (reading 'answer')`. `result` was undefined, and only this function can
     * have made it so.
     *
     * WHY IT WAS UNDEFINED IS STILL UNKNOWN, and this comment will not pretend
     * otherwise. `withTimeout` rejects rather than resolving empty,
     * `withProbeContext` returns `fn(...)` or throws, and the probe in flight
     * returns on every path — the three obvious explanations are all refuted by
     * reading them. Something below still resolves empty on a host sick enough,
     * and this is the first host in 161 crash records that was.
     *
     * SO THE BOUNDARY IS DEFENDED RATHER THAN THE CAUSE FIXED. That is worth
     * doing on its own terms: one unanswerable question should cost one
     * question, which is the whole reason the catch below exists. It caught
     * throwing and not silence, and this file's oldest lesson is that those are
     * different. `action failed` now records a stack, so the next occurrence
     * names the frame this one could not.
     *
     * `unreadable`, not a new word — it is already in `UNINFORMATIVE`, so the
     * sheet ranks it below any named answer and the vocabulary ratchet holds.
     */
    return answered ?? { answer: "unreadable", detail: "the probe returned no answer at all" };
  } catch (err) {
    // WHY the slide was unavailable, carried out rather than flattened into
    // prose. The two causes want opposite fixes and the archive cannot tell them
    // apart today: `gone` means the host answered `isNullObject: true` and the
    // slide really is deleted, so something is removing it; `silent` means the
    // host populated nothing, which is the SAME unanswered-proxy state the
    // update path recovers from by re-reading the collection — 105 of 105 on
    // this host, per `the-re-read-always-rescues-a-refused-lookup`.
    //
    // Worth knowing because it is not rare. The run replaces its scratch slide
    // about 78 times per round, 1495 of 1548 of them after a `no-scratch-slide`,
    // and every one costs a slide add on a host where that takes 0.2-4.0s. The
    // retry then usually succeeds, so the SHEET shows one such answer in ten
    // rounds and the cost is invisible outside the trace.
    if (err instanceof ScratchSlideUnavailable) return { answer: "no-scratch-slide", detail: short(err), why: err.why };
    if (err instanceof ProbeSetupFailed) return { answer: "no-scratch-shape", detail: short(err) };
    // A budget that fired while the setup shapes were still out is a setup
    // failure too, and `"silent"` is a comparable answer that would read as a
    // host divergence. Which sync the deadline caught is the whole difference
    // between "this host says nothing about tag overwrites" and "this host
    // never gave the probe a shape to write a tag on".
    if (isTimeout(err) && awaitingSetupShapes)
      return { answer: "no-scratch-shape", detail: `the sync that was to add them never came back: ${short(err)}` };
    return { answer: isTimeout(err) ? "silent" : "threw", detail: short(err) };
  } finally {
    awaitingSetupShapes = false;
  }
}

/**
 * Every question this build knows how to ask — for the fake's baseline test.
 *
 * Follow-ups included, and flattened rather than listed apart: a conditional
 * question is still a question the sheet can carry, and a baseline that did not
 * know about it would report its answer as an id nobody recognises.
 */
const withFollows = (p: Probe): string[] => [p.id, ...(p.follow ? withFollows(p.follow.probe) : [])];

/**
 * Every probe by id, follow-ups included — what the end-of-run second pass
 * re-asks from. A follow-up only has a row at all when its trigger fired, so a
 * row that is there and never put is a fair thing to put again.
 */
const flatten = (p: Probe): Probe[] => [p, ...(p.follow ? flatten(p.follow.probe) : [])];
const PROBE_BY_ID = new Map(PROBES.flatMap(flatten).map((p) => [p.id, p]));

/**
 * The cleanup's own row in the sheet.
 *
 * Not a `Probe` — it has no question to put and no scratch slide to put it on,
 * because it IS what happens to the scratch slides. It is still a fact about
 * this host, it still has a fake answer to diverge from, and the sheet is the
 * only place anyone will read it. So it is listed here by hand rather than
 * derived, and the two lists below are where it joins the invariant.
 */
/**
 * Which slides a positional clean-up sweep may remove — and it is a SAFETY
 * rule before it is a clean-up one.
 *
 * Delete-by-id cannot work here: 2026-08-11 (`756682e`) measured `the deck
 * still lists 0 of 62 of these ids`, so every by-id delete takes the "already
 * gone" branch and removes nothing. What survives that finding is position: the
 * probe's slides are APPENDED by `slides.add()`, so they are the last N in the
 * deck and need no id at all.
 *
 * Position is also how an add-in destroys someone's work, so this returns a
 * plan only when every one of these holds, and each is a separate way of saying
 * "never touch a slide this run did not add":
 *
 *  - Both deck counts are known. A host that will not count its slides does not
 *    get to have slides deleted from it by arithmetic.
 *  - The deck GREW. If it did not, the run added nothing that is still there.
 *  - Never more than this run added, minus whatever already went back.
 *  - Never more than the deck grew — if something else removed slides while the
 *    probe ran, the sum no longer describes the deck and the smaller number is
 *    the honest one.
 *  - The first index to delete is at or after the deck's size when the run
 *    started. This is implied by the two rules above and asserted anyway,
 *    because it is THE property: everything before that index was the user's
 *    before this run began.
 *
 * The caller deletes from the highest index down, so removing one cannot shift
 * the index of another still to go.
 */
export function positionalSweepPlan(o: {
  deckAtStart?: number;
  deckNow?: number;
  added: number;
  alreadyDeleted: number;
}): { from: number; count: number } | null {
  if (o.deckAtStart === undefined || o.deckNow === undefined) return null;
  const grew = o.deckNow - o.deckAtStart;
  if (grew <= 0) return null;
  const count = Math.min(o.added - o.alreadyDeleted, grew);
  if (count <= 0) return null;
  const from = o.deckNow - count;
  // Belt and braces: the arithmetic above already guarantees this, and a plan
  // that reached into the user's own slides would be the one bug here that
  // cannot be apologised for.
  if (from < o.deckAtStart) return null;
  return { from, count };
}

/**
 * How many scratch slides ACTUALLY went back, from the deck rather than from
 * the deletes' own opinion of themselves.
 *
 * Extracted so it can be checked without a PowerPoint, which is the only way
 * this particular bug is checkable at all: reproducing it through the fake
 * needs three knobs turned together (a delete that no-ops, a host that refuses
 * the id, and a deck listing that id under another string), and a test built
 * out of three coincidences tests the coincidences.
 *
 * `claimed` is what `deleteSlideById` reported. The deck's own count outranks
 * it, because on 2026-08-11 the two disagreed completely: `42 of 42 deleted,
 * left: 0` into a deck that ended the round holding 56 slides it had added, 47
 * of them blank, with the owner's screenshot showing one carrying the 20x20
 * rectangle `shape-add-held-slide-proxy` draws. `deleteSlideByPosition` reads
 * `indexOf(id) < 0` as "already gone", which is sound only while the id we hold
 * and the ids the deck lists are the same strings — and there they were not:
 * scratch ids read `4123571115#123571113`, the deck listed `257#2599158489`.
 *
 * The MINIMUM of the two, never the maximum: a delete that cannot be
 * corroborated is unknown, and reporting an unknown as a success is how a run
 * comes to leave forty blank slides in someone's deck and say it left none.
 * When the deck will not give a count at all the claim is all there is and
 * stands — with `shrankBy` undefined, so nothing pretends otherwise.
 */
/**
 * Of the slides this run took, which is the clean-up still responsible for?
 *
 * Every number in the clean-up used to derive from "how many scratch slides did
 * this run take", which equals "how many are still in the deck" only because
 * nothing ever gave one back early. That assumption is load-bearing in three
 * places — `slidesActuallyReturned`, `positionalSweepPlan` and the
 * `scratch-slides-returned` answer — and it is exactly what blocked the
 * pass-boundary sweep: three reconciliations at the call sites produced three
 * different wrong totals, because the assumption sits upstream of all of them.
 *
 * `deckBefore` is read when the clean-up starts, so a slide returned before that
 * is already gone from the deck's own count. Counting it again makes the run
 * report slides it never failed to return; leaving it in the total makes the
 * clean-up owe a slide that is not there. Both were observed.
 *
 * NOTHING RETURNS A SLIDE EARLY YET. `returnedEarly` is empty on every path
 * today, which is what makes this a pure refactor. The tests exercise it
 * non-empty, so the accounting is proven for the case it exists to allow before
 * anything depends on it.
 */
export function outstandingScratch(taken: readonly string[], returnedEarly: ReadonlySet<string>): string[] {
  return taken.filter((id) => !returnedEarly.has(id));
}

/**
 * The scratch clean-up's verdict.
 *
 * `left <= 0` IS A COMPLETE HANDBACK, not just `left === 0`. Grading on `!left`
 * meant a clean-up that removed every slide it added could never say "all" the
 * moment the count went negative — which it did in **25 of 140 archived
 * rounds**, every one of them graded `some` while the deck came back to the size
 * it started at. 25 of the 30 `some` verdicts in the whole archive are that
 * inversion rather than a shortfall.
 *
 * The cause was two denominators 24 lines apart: the positional sweep was
 * licensed with `outstanding.length + unnamedLeftBehind` and the verdict was
 * scored against `outstanding.length`, so the numerator counted slides the
 * denominator did not know existed. That is fixed at the call site. This keeps
 * the sharp edge off, so a future off-by-one costs the wording and not the
 * meaning.
 */
export function scratchCleanupAnswer(o: { addedAny: boolean; left: number; actually: number }): string {
  if (!o.addedAny) return "none-added";
  if (o.left <= 0) return "all";
  return o.actually ? "some" : "none";
}

/**
 * The clause describing what the clean-up did not get back.
 *
 * A NEGATIVE `left` GETS ITS OWN SENTENCE and never one of the other two. Both
 * of those describe slides that are MISSING, so a negative number in either says
 * the opposite of what happened: the archive holds `-4 never landed — the host
 * took the add and the deck never listed them` (round 088, beside a deck that
 * ended exactly where it started) and `-1 left in the deck` printed directly
 * beside `the deck still lists 1` (round 073).
 *
 * The denominator fix should make the negative branch unreachable. It is kept
 * because the condition is real — a sweep that reached past this run's own
 * slides would produce it — and named rather than clamped, because clamping a
 * number into a sentence that cannot hold it is exactly how the last one
 * survived 25 rounds. **Name the condition, never the cause.**
 *
 * WHICH SENTENCE THE DECK CHOOSES, for the positive case: round 029 recorded
 * `left=73` beside `deckBefore=1 deckAfter=1 stillListed=0`. Seventy-three
 * slides were counted as abandoned while the deck stood at one the whole time
 * and listed none of their ids — they never landed at all, which is a different
 * bug from a clean-up that missed something and wants a different fix. Five
 * attempts were reverted trying to stop a leak that was this sentence.
 */
export function scratchLeftSentence(o: {
  left: number;
  stillListed?: number;
  deckBefore?: number;
  deckAfter?: number;
}): string {
  if (o.left < 0)
    return `; the deck lost ${-o.left} MORE slide(s) than this run can account for — read the sweep plan, not this line`;
  if (!o.left) return "";
  const neverLanded =
    o.stillListed === 0 && o.deckAfter !== undefined && o.deckBefore !== undefined && o.deckAfter <= o.deckBefore;
  if (neverLanded) return `; ${o.left} never landed — the host took the add and the deck never listed them`;
  /**
   * A THIRD CONDITION, BECAUSE THE DECK CAN CONTRADICT THE ARITHMETIC OUTRIGHT.
   *
   * `left` is `added - actually`, and those are not the same KIND of number.
   * `added` is every scratch slide this run is still holding, accumulated across
   * the whole round; `actually` is bounded by `shrankBy`, a delta between two
   * readings of the deck. Subtracting a point-in-time delta from a cumulative
   * count is a category error, and it shows: round 346 recorded `left=20` beside
   * `deckAfter=1`. Twenty slides cannot be left in a deck that holds one. Across
   * the archive 106 of 422 of these records — 25% — claim more left behind than
   * the deck contains.
   *
   * NOT CLAMPED, for the reason the docstring above already gives: "clamping a
   * number into a sentence that cannot hold it is exactly how the last one
   * survived 25 rounds." Both numbers are printed and the contradiction is
   * NAMED, which is the one thing a reader cannot do for themselves from a
   * single figure.
   */
  if (o.deckAfter !== undefined && o.left > o.deckAfter)
    return (
      `; ${o.left} unaccounted for, and the deck ends at ${o.deckAfter} slide(s)` +
      ` — they cannot all be in it, so this is a counting question and not a leak`
    );
  return `; ${o.left} left in the deck`;
}

export function slidesActuallyReturned(o: {
  claimed: number;
  added: number;
  deckBefore?: number;
  deckAfter?: number;
}): { actually: number; left: number; shrankBy?: number } {
  const shrankBy = o.deckBefore !== undefined && o.deckAfter !== undefined ? o.deckBefore - o.deckAfter : undefined;
  const actually = shrankBy === undefined ? o.claimed : Math.min(o.claimed, Math.max(0, shrankBy));
  return { actually, left: o.added - actually, shrankBy };
}

export const SCRATCH_CLEANUP_ID = "scratch-slides-returned";

export const PROBE_IDS: readonly string[] = [...PROBES.flatMap(withFollows), SCRATCH_CLEANUP_ID];

/**
 * The questions that answer with no live scratch slide at all.
 *
 * Exported so the tests can ask the probes rather than carry a copy of this
 * list. The copy is what drifts: a hardcoded `excused` set is why marking five
 * genuinely slide-free questions turned two passing tests red for the right
 * reason, and a list that has to be edited by hand is a list that will one day
 * be edited wrong in the quiet direction.
 */
/**
 * The shortlist the later passes fall back to when scratch slides are scarce.
 *
 * Derived from the marks rather than listed again, for the same reason
 * `NO_SLIDE_NEEDED_IDS` is: a hand-kept copy is a copy that will one day be
 * edited wrong in the quiet direction.
 *
 * SEVEN MARKS CAME OFF ON 2026-08-29, and the reason is worth keeping because
 * it is a cost this list quietly carried. Every id here is a scarce scratch
 * slide spent under pressure, and these seven were spending it on questions the
 * host had settled:
 *
 *     shape-resolve-held-slide-proxy              yes             24 of 24
 *     collection-read-poisons-the-creation-handle refused         25 of 25
 *     does-a-failed-group-poison-the-tag          tags-gone       25 of 25
 *     which-end-a-short-read-drops                all             25 of 25
 *     creationid-on-fresh-shape                   absent          25 of 25
 *     creationid-survives-a-sync                  no-creation-id  25 of 25
 *     creationid-survives-grouping                no-creation-id  25 of 25
 *
 * They were marked because they sat in `PENDING_QUESTIONS`, and they sat there
 * because the committed fixture predated them — a 2026-08-12 capture still in
 * place on 2026-08-28. So the probe re-asked seven settled questions on every
 * pressured run for a fortnight, to keep a table current that nothing was
 * making current. The invariant in `test/host-probe.test.ts` was doing its job
 * perfectly; it keeps this list equal to those tables, and the tables were
 * wrong.
 *
 * The lesson for whoever adds the next mark: this list is downstream of the
 * FIXTURE. Swap the fixture when a newer sheet answers at least as much
 * (`scripts/host-history.mjs --fixture` counts it for you) and the marks
 * correct themselves.
 */
export const RESAMPLE_IDS: readonly string[] = PROBES.flatMap(flatten)
  .filter((p) => p.resample)
  .map((p) => p.id);

export const NO_SLIDE_NEEDED_IDS: readonly string[] = PROBES.flatMap(flatten)
  .filter((p) => p.noSlideNeeded)
  .map((p) => p.id);

/**
 * The questions EVERY run puts, whatever the host says.
 *
 * A follow-up is conditional by definition, so "one answer per known id" stopped
 * being the sheet's invariant the moment the first one existed. The invariant
 * that replaced it is two-sided and both halves matter: every id in this list
 * appears in every sheet, and no sheet ever carries an id outside `PROBE_IDS`.
 * A single count could not express either.
 */
export const ALWAYS_ASKED_IDS: readonly string[] = [...PROBES.map((p) => p.id), SCRATCH_CLEANUP_ID];

/**
 * Ids that exist ONLY as a partner question.
 *
 * A follow-up is conditional by definition — it rides its trigger and is never
 * scheduled on its own — so it is neither always asked nor a candidate for the
 * scarce-slide shortlist, and the two invariants that read those lists have to
 * know the difference. Derived rather than listed, for the same reason
 * `RESAMPLE_IDS` is: a hand-kept copy is the one that will drift.
 */
export const FOLLOW_UP_IDS: readonly string[] = PROBES.flatMap((p) =>
  p.follow ? flatten(p.follow.probe).map((f) => f.id) : [],
);

/**
 * What a probe run FOUND, in one line, on screen.
 *
 * Every probe run so far has cost a round trip to establish nothing: download,
 * send, wait for someone to run `host-diff`, hear that the answers are the same
 * as last time. The comparison table is plain data, so the pane can do the diff
 * itself and say whether this run is worth sending at all.
 *
 * Divergences that are already DECLARED are counted separately from new ones.
 * A run that reproduces seven known disagreements is not news; a run with one
 * undeclared answer is the only kind worth anyone's attention, and it should
 * not have to be found by eye in a seventeen-row JSON file.
 */
/**
 * Whether a question's disagreement is explained by the run falling over.
 *
 * `stable: false` reads as "the host changed its answer mid-round", and drives
 * both the `CHANGED ITS ANSWER MID-ROUND` line and whether a sheet is worth
 * sending on at all. For the three collection questions it is mostly reporting
 * OUR run's decay instead.
 *
 * Every sample carries the `regime` the host was in when it was taken. Measured
 * across the archive, of the rounds where these report `stable: false`, the
 * share whose every dissenting sample was taken under `collection-refused`:
 *
 *     which-end-a-short-read-drops        112 of 127   (88%)
 *     shapes-items-count-honest           114 of 118   (97%)
 *     shapes-items-via-positional-slide   118 of 122   (97%)
 *
 * Pass 1 answers while the host is healthy; passes 2 and 3 answer once it has
 * started refusing collections, and answer differently. That is not the host
 * being non-deterministic, it is the host being in a state we already named.
 *
 * NOT DISCARDED, SEPARATED. The disagreement is real and stays on the sheet —
 * it just stops counting as the unexplained kind, which is the kind that asks
 * for someone's attention and the kind `UNSTABLE_ANSWERS` is curated from.
 */
export function dissentIsDegradation(samples: ProbeSample[] | undefined): boolean {
  const real = (samples ?? []).filter((s) => !NOT_ASKED.has(s.answer));
  if (real.length < 2) return false;
  const first = real[0].answer;
  const dissent = real.filter((s) => s.answer !== first);
  return dissent.length > 0 && dissent.every((s) => s.regime === "collection-refused");
}

export function summariseHostSheet(sheet: HostAnswerSheet): {
  asked: number;
  neverPut: string[];
  known: string[];
  fresh: string[];
  /**
   * Questions that answered one way and then another INSIDE this run.
   *
   * The fact `UNSTABLE_ANSWERS` was assembled by hand across ten rounds to
   * state, now stated by the round that saw it. A sheet that knows it disagreed
   * with itself and does not say so has only moved the finding somewhere nobody
   * looks.
   */
  unstable: string[];
  /**
   * The same disagreement, but every dissenting sample was taken while the host
   * had already started refusing collections.
   *
   * Kept apart from `unstable` rather than dropped: it is a real difference and
   * it stays on the sheet, but it is OUR run decaying rather than the host being
   * non-deterministic, and it does not merit anyone's attention. See
   * `dissentIsDegradation` for the numbers — 88-97% of the instability on the
   * three collection questions is this.
   */
  whileDegraded: string[];
  /**
   * Answers the host gave FEWER times than it was asked.
   *
   * `unstable` catches a question that disagreed with itself, which needs two
   * real samples. This catches the other half: a question answered once and
   * refused the rest, where there is no disagreement to find because there is
   * only one answer. Round 26 carried `group-reports-its-children threw (1/2)`
   * and `shape-add-fresh-getitem-slide threw (1/2)` — a `threw` is what a
   * misclassified SETUP failure looks like, and one attempt is exactly how many
   * it takes for that to become the row's answer.
   */
  thin: string[];
} {
  const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
  const neverPut = sheet.answers.filter((a) => NOT_ASKED.has(a.answer)).map((a) => a.id);
  const { differ } = diffAnswers(answers, FAKE_BASELINE) as { differ: { id: string }[] };
  // `UNSTABLE_ANSWERS` counts as declared too, and leaving it out made the pane
  // cry wolf on a schedule. `shape-add-positional-slide-proxy` is a coin: it has
  // answered `yes, yes, threw, yes, yes, threw` across six rounds, and the fake
  // says `yes`. So every round the coin lands `threw` the pane announced
  // "NEW: shape-add-positional-slide-proxy" — news the sixth time it was seen,
  // about the one question this repo has documented at greatest length as
  // varying run to run. A declaration is a declaration wherever it lives.
  const declared = (id: string) => id in KNOWN_DIVERGENCES || id in UNSTABLE_ANSWERS;
  const known = differ.filter((d) => declared(d.id)).map((d) => d.id);
  const fresh = differ.filter((d) => !declared(d.id)).map((d) => d.id);
  const disagreed = sheet.answers.filter((a) => a.stable === false);
  const unstable = disagreed.filter((a) => !dissentIsDegradation(a.samples)).map((a) => a.id);
  const whileDegraded = disagreed.filter((a) => dissentIsDegradation(a.samples)).map((a) => a.id);
  const thin = sheet.answers.filter((a) => a.support && a.support.asked === 1).map((a) => a.id);
  return { asked: sheet.answers.length - neverPut.length, neverPut, known, fresh, unstable, whileDegraded, thin };
}

/** Whether a probe run is worth sending on — anything unexplained in it. */
export function sheetNeedsAttention(sheet: HostAnswerSheet): boolean {
  const s = summariseHostSheet(sheet);
  // Instability counts. A question that answered two ways inside one round is
  // the finding this project used to need ten rounds and a hand-kept table to
  // reach, and the verdict line already says so — a round that says "worth
  // sending" in words and "nothing to see" in its return value is worse than
  // either.
  return s.fresh.length > 0 || s.neverPut.length > 0 || s.unstable.length > 0;
}

/** The same summary, in the words the pane shows. */
export function describeHostSheet(sheet: HostAnswerSheet): string {
  const { asked, neverPut, known, fresh, unstable, whileDegraded } = summariseHostSheet(sheet);
  const parts = [`${asked} of ${sheet.answers.length} questions answered`];
  if (neverPut.length) parts.push(`${neverPut.length} never put (${neverPut.join(", ")})`);
  if (known.length) parts.push(`${known.length} known divergence${known.length === 1 ? "" : "s"}`);
  if (fresh.length) parts.push(`NEW: ${fresh.join(", ")}`);
  // A question that changed its answer mid-round is worth sending whatever the
  // diff says: it is the one thing a single sheet could never report before.
  if (unstable.length) parts.push(`CHANGED ITS ANSWER MID-ROUND: ${unstable.join(", ")}`);
  // SAID, BUT NOT AS THE SAME THING. Every dissenting sample here was taken
  // after the host had started refusing collections, so the disagreement is this
  // run decaying rather than the host changing its mind. Reported so it is not
  // invisible, worded so it does not read as a finding, and deliberately absent
  // from the send/don't-send verdict below.
  if (whileDegraded.length)
    parts.push(`differed only once the host was refusing collections: ${whileDegraded.join(", ")}`);
  // The verdict the owner actually needs: send it, or don't.
  parts.push(
    fresh.length || neverPut.length || unstable.length
      ? "Saved — worth sending."
      : "Saved. Nothing new — no need to send it.",
  );
  return parts.join(" · ");
}
