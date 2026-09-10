// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, applyWebProfile, faults } from "./helpers/office-host";
import { addScratchSlide } from "../src/render/powerpoint";
import {
  runHostProbes,
  PROBE_IDS,
  ALWAYS_ASKED_IDS,
  FOLLOW_UP_IDS,
  SCRATCH_CLEANUP_ID,
  NO_SLIDE_NEEDED_IDS,
  describeHostSheet,
  sheetNeedsAttention,
  summariseHostSheet,
  _setProbeBudgetForTest,
  NOT_ASKED,
  UNINFORMATIVE,
  weakAnswer,
  PROBE_PASSES,
  RESAMPLE_IDS,
  stabilityOf,
  answerRank,
  outstandingScratch,
  supportOf,
  thinSupport,
  regimeFrom,
  slidesActuallyReturned,
  positionalSweepPlan,
  type HostAnswerSheet,
} from "../src/render/host-probe";
import { readFileSync } from "node:fs";
// @ts-expect-error — a plain .mjs tool with no types. Imported so the shortlist
// and the tables that define it are pinned to each other by a test.
import { UNSTABLE_ANSWERS, PENDING_QUESTIONS } from "../scripts/host-baseline.mjs";
// @ts-expect-error — a plain .mjs tool with no types. The baseline lives THERE
// rather than here, so the diff tool and this test cannot drift apart: two
// copies of the same table is how a claim quietly stops matching its check.
// prettier-ignore
import { FAKE_BASELINE, diffAnswers, answersOf, sheetOf, NEVER_ASKED, UNINFORMATIVE_ANSWERS } from "../scripts/host-diff.mjs";
import { setTracing, trace, traceLog } from "../src/core/trace";

/**
 * The fake's own answer sheet, frozen.
 *
 * Two things this holds down, and the second is the point.
 *
 * 1. The sheet comes back COMPLETE. A probe that throws, times out or wedges
 *    contributes its answer and nothing else — the sheet from a misbehaving
 *    host is the one worth having, so surviving misbehaviour is the feature.
 *
 * 2. What the fake CLAIMS about itself is written down. Every assertion in this
 *    repo about Office.js rests on this fake, and nobody has ever checked it
 *    against a real PowerPoint. The baseline below is one half of that check;
 *    the other half arrives the first time someone runs the probe in a real
 *    host and `npm run host-diff` compares the two.
 *
 * So a change to this baseline is never "just update it". It means the fake now
 * claims something different about the host it stands for, and the question is
 * whether the real host agrees.
 */

afterEach(() => vi.unstubAllGlobals());

/** The probe's own "this was never put" vocabulary — never a host answer. */
const NOT_ASKED_WORDS = ["no-scratch-slide", "no-scratch-shape"];

describe("a named answer outranks the catch-all", () => {
  it("treats `other` as weak, and a named answer as strong", () => {
    // `other` is every probe's catch-all — the outcome was none of the ones the
    // question knows how to name. host-probe.ts says so about the tag-key
    // question in as many words: "not an opinion about tag keys at all, just the
    // stale handle refusing both writes."
    expect(weakAnswer("other"), "the catch-all locked the row").toBe(true);
    expect(weakAnswer("no-scratch-slide"), "never-asked was already weak").toBe(true);
    expect(weakAnswer("overwrites"), "a named answer must be strong").toBe(false);
    expect(weakAnswer("keeps-first")).toBe(false);
  });

  it("is what `tags-add-same-key-twice` needed: 83 of 86 rounds said UNKNOWN", () => {
    // The row keeps its FIRST real answer, and `other` counted as real — so a
    // pass-1 `other` locked it and a later named answer could not replace it.
    // Rounds 074 and 091, on two different builds, both carry `overwrites` in
    // their samples while the rolled-up answer reads `other`.
    //
    // Modelled as the promotion rule sees it: a weak answer already on the row,
    // then a named one arriving from a later pass.
    const promote = (seen: string, arriving: string) => (weakAnswer(seen) && !weakAnswer(arriving) ? arriving : seen);
    expect(promote("other", "overwrites"), "the host answered and the sheet hid it").toBe("overwrites");
    expect(promote("no-scratch-slide", "overwrites")).toBe("overwrites");
    // AND NOTHING OUTRANKS A NAMED ANSWER. That half is the original rule and it
    // stays: a sheet must mean today what it meant yesterday, and disagreement
    // between samples is `stable`'s job to report rather than `answer`'s to hide.
    expect(promote("overwrites", "keeps-first"), "a named answer was overwritten").toBe("overwrites");
    expect(promote("overwrites", "other"), "the catch-all displaced a real answer").toBe("overwrites");
  });
});

describe("the failure vocabulary is classified on purpose", () => {
  /**
   * THREE WORDS, FOUND ONE AT A TIME, EACH COSTING ROUNDS OF REAL ANSWERS.
   *
   *   `other`       locked `tags-add-same-key-twice` in 83 of 86 rounds
   *   `unreadable`  locked 87 rows across 8 questions, archive-wide
   *   `silent`      locked 52 more, 16 of them a plain `yes`
   *
   * Each was fixed the same way and each was found by accident, months apart, by
   * someone reading samples for an unrelated reason. The defect is not really
   * any of the three words: it is that a probe can invent a new way of saying
   * "no answer" and nothing forces anyone to decide which kind of word it is.
   *
   * So the vocabulary is pinned here. Add a word to `host-probe.ts` and this
   * test fails until you say what it is — a REAL answer the host gave, or one
   * more way of not answering. That is a thirty-second decision at the moment it
   * is made, and it has three times now cost far more than that afterwards.
   */
  const readVocabulary = () => {
    const src = readFileSync("src/render/host-probe.ts", "utf8");
    const words = new Set<string>();
    for (const m of src.matchAll(/answer:\s*"([a-z0-9-]+)"/g)) words.add(m[1]);
    // `answer: isTimeout(err) ? "silent" : "threw"` — the form that hid `silent`
    // from a first, simpler version of this scan.
    for (const m of src.matchAll(/answer:\s*[^,;}\n]*?\?[^,;}\n]*/g))
      for (const q of m[0].matchAll(/"([a-z0-9-]+)"/g)) words.add(q[1]);
    return words;
  };

  /**
   * Words that are a real thing the host said. `threw` and `commit-threw` belong
   * here and NOT with the non-answers: a host actively raising an error on a
   * call is a fact about that call, reproducible and worth keeping. A deadline
   * expiring is not.
   */
  const REAL_ANSWERS = new Set([
    "yes",
    "no",
    "null",
    "undefined",
    "boolean",
    "number",
    "string",
    "absent",
    "none",
    "all",
    "neither",
    "both-answer",
    "partial",
    "some",
    "first",
    "second",
    "two",
    "at-least-5",
    "kept",
    "lost",
    "changed",
    "stable",
    "still-there",
    "draws",
    "drew-nothing",
    "refused",
    "refused-after-group",
    "no-refusal",
    "threw",
    "commit-threw",
    "add-threw",
    "add-returned-nothing",
    "claims-live",
    "null-object",
    "not-listed",
    "reports-gone",
    "tags-gone",
    "overwrites",
    "keeps-first",
    "keeps-head",
    "keeps-tail",
    "scattered",
    "index-beats-items",
    "index-unreadable",
    // `durable-slide-lists-its-shapes` — the round-254 control. A REAL answer:
    // the host enumerated a collection on a slide this run did not create, which
    // is precisely the thing its trigger cannot establish about its own slide.
    "lists-shapes",
    "survives-8",
    "survives-12",
    "same-as-id",
    "through-a-refetched-handle",
    "returns-something",
    "unrotated-box",
    "rotation-did-not-take",
    "rotation-not-writable",
    "no-read-after-rotation",
    "no-rotation-api",
    "no-api",
    "no-binding-api",
    "no-count",
    "no-creation-id",
    "no-durable-slide",
    "no-group-id",
    "no-id",
    "no-members",
    "no-layouts",
  ]);

  it("classifies every answer word the probe can emit", () => {
    const words = readVocabulary();
    // The scan is itself an instrument, so it gets checked before it is trusted.
    // `silent` is only ever emitted from a ternary; an extractor that misses it
    // would pass this whole test while covering nothing.
    expect(words.has("silent"), "the extractor missed the ternary form and proves nothing").toBe(true);
    expect(words.has("unreadable")).toBe(true);
    expect(words.size, "the extractor found suspiciously few words").toBeGreaterThan(40);

    const unclassified = [...words].filter((w) => !NOT_ASKED.has(w) && !UNINFORMATIVE.has(w) && !REAL_ANSWERS.has(w));
    expect(
      unclassified,
      `Unclassified probe answer(s): ${unclassified.join(", ")}. Decide what each one IS. ` +
        `A real thing the host said goes in REAL_ANSWERS; one more way of not answering goes in ` +
        `UNINFORMATIVE, where it can be displaced by a real answer. Getting this wrong is what made ` +
        `the sheet report "unreadable" for eleven rounds while the host was answering.`,
    ).toEqual([]);
  });

  it("keeps the three known non-answers weak, and a throw strong", () => {
    for (const w of ["other", "unreadable", "silent"]) {
      expect(answerRank(w), `${w} must not lock a row`).toBe(1);
    }
    // The line the classification turns on: an error the host RAISED is a fact
    // about the host; a deadline that expired is a fact about the network.
    expect(answerRank("threw"), "a throw is the host speaking").toBe(2);
    expect(answerRank("commit-threw")).toBe(2);
  });
});

describe("`unreadable` is the same trap as `other`, and cost eight questions their answers", () => {
  const sheetWith = (rows: { id: string; answer: string; passes?: number[] }[]): HostAnswerSheet =>
    ({
      kind: "powerchart-host-answers",
      source: "fake",
      build: "test",
      requirementSets: [],
      answers: rows.map((r) => ({
        ...r,
        question: r.id,
        ms: 1,
        samples: (r.passes ?? [1]).map((pass) => ({
          answer: r.answer,
          pass,
          atMs: 0,
          regime: "healthy" as const,
          scratch: "first-slide" as const,
          acquired: "first-add" as const,
        })),
      })),
    }) as HostAnswerSheet;

  it("ranks never-asked below unnameable below a real answer", () => {
    // A BOOLEAN COLLAPSED TWO REAL THINGS. `no-scratch-shape` means the probe
    // never reached its question; `unreadable` means it reached it and could not
    // read the result. The second is strictly more informative, so a re-ask that
    // gets as far as reading must still displace a deferral — while a real
    // answer displaces both.
    expect(answerRank("no-scratch-shape"), "never-asked is the floor").toBe(0);
    expect(answerRank("not-asked")).toBe(0);
    expect(answerRank("unreadable"), "reached the question, read nothing").toBe(1);
    expect(answerRank("other")).toBe(1);
    expect(answerRank("unrotated-box"), "a real answer is the ceiling").toBe(2);
    expect(answerRank("threw"), "a throw is a fact about the host, not a non-answer").toBe(2);
  });

  it("lets a real answer displace `unreadable` — through mergeHostSheets, not a model of it", async () => {
    const { mergeHostSheets } = await import("../src/render/host-probe");
    /**
     * Measured over all 322 archived rounds and 12,346 rows: 87 rows report
     * `unreadable` while their own samples carry a real answer, across eight
     * questions. `which-end-a-short-read-drops` answered `none` in 11 of them,
     * `shapes-items-count-honest` answered `short-0` in 8, and
     * `picture-then-shape-read` and `binding-names-shape-later` each answered a
     * plain `yes`.
     *
     * The sharpest is `rotation-keeps-the-unrotated-box`, used here. Over its 33
     * passes the host said `unrotated-box` four times and the contradicting
     * answer NOT ONCE; the rest could not read. That question exists because if
     * this host placed rotated shapes by their post-rotation box, every diagonal
     * in every line, scatter, radar and violin chart would land wrong. The host
     * answered it. Eleven sheets in a row said `unreadable`.
     *
     * Driven through `mergeHostSheets` on purpose: the rule has TWO call sites
     * and a test that models the comparison locally proves nothing about either.
     */
    const base = sheetWith([
      { id: "rotation-keeps-the-unrotated-box", answer: "unreadable" },
      { id: "untouched", answer: "unreadable" },
    ]);
    const merged = mergeHostSheets(
      base,
      sheetWith([{ id: "rotation-keeps-the-unrotated-box", answer: "unrotated-box", passes: [2] }]),
    );
    const row = merged.answers.find((r) => r.id === "rotation-keeps-the-unrotated-box")!;
    expect(row.answer, "the host answered and the sheet still hid it").toBe("unrotated-box");
    // The re-ask asked ONE question; every other row keeps what it said.
    expect(merged.answers.find((r) => r.id === "untouched")!.answer).toBe("unreadable");
  });

  it("still refuses to let a non-answer displace a real one", async () => {
    const { mergeHostSheets } = await import("../src/render/host-probe");
    // The invariant that must survive the change, in the other direction: a
    // later `unreadable` must not overwrite an answer the host already gave.
    const base = sheetWith([{ id: "q", answer: "unrotated-box" }]);
    const merged = mergeHostSheets(base, sheetWith([{ id: "q", answer: "unreadable", passes: [2] }]));
    expect(merged.answers.find((r) => r.id === "q")!.answer, "a failed read erased a real answer").toBe(
      "unrotated-box",
    );
  });
});

describe("a host run that resolves nothing at all", () => {
  it("costs one question, not the whole round", async () => {
    /**
     * THE ROUND THIS COST. On 2026-09-02 the crossed-deck experiment's first
     * two attempts each died at ~83s with `Cannot read properties of undefined
     * (reading 'answer')`, and the driver — whose completion signal is a button
     * the dead round never enables — then waited out its full 30-minute wedge
     * budget on each. Two hours of host time for a fault in this file, and the
     * only "30 minutes with no finish" records in an archive of 161.
     *
     * WHERE IT THREW. `ask`'s caller does `record({ …, ...result })` and then
     * `trace("probe", "answered", { answer: result.answer })`. The spread
     * swallows `undefined` — `{...undefined}` is `{}` — so `record` survives
     * and the trace is what dies. The crash record shows exactly that: an
     * `asking` line for `getitemornullobject-missing` with no `answered` line
     * after it.
     *
     * WHY `ask` returned undefined on the real host is still unknown, and this
     * test does not claim to reproduce that cause — it reproduces the SHAPE, a
     * host whose `run` resolves with nothing, which is the only way the value
     * can get through `withProbeContext` and `withTimeout` without either
     * throwing. What is asserted is the property that matters either way: one
     * unanswerable question costs one question.
     */
    installHost([makeSlide("s1")]);
    const real = (globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint;
    // Everything else about the host is left exactly as it was — only the one
    // thing that hands back a value is emptied.
    vi.stubGlobal("PowerPoint", { ...real, run: async () => undefined });

    const sheet = await runHostProbes("fake", "test");

    // IT CAME BACK AT ALL. Before the guard this rejected, and the round with
    // it — the whole point is that a sheet arrives.
    expect(sheet.answers.length, "the run threw instead of returning a sheet").toBeGreaterThan(0);
    // And every question is ANSWERED, with a word the sheet already knows.
    // `unreadable` is in `UNINFORMATIVE`, so it ranks below any named answer
    // and cannot lock a row against a later pass that succeeds.
    const empty = sheet.answers.filter((a) => a.answer === "unreadable");
    expect(empty.length, "a host answering nothing produced no `unreadable` rows").toBeGreaterThan(0);
    for (const a of sheet.answers) {
      expect(typeof a.answer, `${a.id} came back with no answer at all`).toBe("string");
    }
  });
});

const probeSheet = async () => {
  const answers = await runHostProbes("fake", "test");
  return Object.fromEntries(answers.answers.map((a) => [a.id, a.answer]));
};

describe("the fake host's answer sheet", () => {
  it("answers every question, and says what it claims", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await probeSheet();
    // Two-sided, because a follow-up is conditional and a single list cannot
    // say both things. Every question a run ALWAYS puts is here, and nothing
    // here is a question this build does not know how to ask — an id outside
    // `PROBE_IDS` would be an answer the baseline could never account for.
    //
    // Containment rather than equality on the first half, because a follow-up
    // whose trigger the FAKE happens to satisfy is always put on the fake and
    // never unconditional in general — `shape-add-held-slide-proxy-again`
    // follows an answer the fake always gives. Exact membership is not lost:
    // `toEqual(FAKE_BASELINE)` below pins it, and pins the answers too.
    for (const id of ALWAYS_ASKED_IDS) expect(Object.keys(sheet), `${id} was not asked`).toContain(id);
    for (const id of Object.keys(sheet)) expect(PROBE_IDS, `${id} is not a question this build asks`).toContain(id);
    // Compared against the table the DIFF tool carries, so the two cannot
    // disagree about what the fake claims. The commentary on each answer lives
    // beside it in `scripts/host-diff.mjs`, together with what would be wrong
    // if a real host answered differently.
    expect(sheet).toEqual(FAKE_BASELINE);
  });

  it("agrees with itself, so a diff of a sheet against the baseline is empty", () => {
    // The diff's own sanity: identical sheets must produce no divergence, or
    // every real comparison is noise.
    const d = diffAnswers(FAKE_BASELINE, FAKE_BASELINE);
    expect(d.differ).toEqual([]);
    expect(d.onlyReal).toEqual([]);
    expect(d.onlyFake).toEqual([]);
    expect(d.agree.length).toBe(Object.keys(FAKE_BASELINE).length);
  });

  it("reports a disagreement, and a question one side was never asked", () => {
    // A question one side has no answer for is a HOLE in the comparison, not a
    // match. Counting it as agreement is how a diff comes to mean nothing.
    const real = { ...FAKE_BASELINE, "untrack-available": "yes", "extra-question": "hm" };
    delete real["getitemat-past-end"];
    const d = diffAnswers(real, FAKE_BASELINE);
    expect(d.differ.map((x: { id: string }) => x.id)).toEqual(["untrack-available"]);
    expect(d.differ[0].real).toBe("yes");
    expect(d.differ[0].fake).toBe("no");
    expect(d.differ[0].means, "a divergence with nothing said about what rests on it").toBeTruthy();
    expect(d.onlyReal).toEqual(["extra-question"]);
    expect(d.onlyFake).toEqual(["getitemat-past-end"]);
  });

  it("shares its never-asked vocabulary with the diff tool, word for word", () => {
    // A comment saying the two are "kept in step" is not a mechanism. `not-asked`
    // — what the mute breaker records when it abandons the rest of a sheet — was
    // added on the probe side and not on the tool's, so every question that
    // breaker gives up on was compared against the fake and reported as a real
    // host DIVERGENCE. That is the exact failure diffAnswers' `notAsked` branch
    // exists for, wearing a word that did not exist when it was written.
    expect([...NEVER_ASKED].sort()).toEqual([...NOT_ASKED].sort());
    // And no probe may ever answer with one of these words, or the diff would
    // read a genuine answer as a question nobody put.
    for (const w of NOT_ASKED) expect(FAKE_BASELINE, `the fake answers "${w}"`).not.toHaveProperty(w);
  });

  it("reads a real sheet's wrapper, not just a bare map", () => {
    const sheet = {
      kind: "powerchart-host-answers",
      source: "PowerPoint web",
      build: "abc",
      requirementSets: ["1.1"],
      answers: [{ id: "untrack-available", question: "?", answer: "yes", ms: 1 }],
    };
    expect(answersOf(sheet)).toEqual({ "untrack-available": "yes" });
    expect(answersOf({ nonsense: true })).toEqual({ nonsense: true });
    expect(answersOf(null)).toBeNull();
    // A bare sheet is already the sheet; a bare map passes through, so the
    // committed baseline still reads.
    expect(sheetOf(sheet)).toBe(sheet);
    expect(sheetOf({ nonsense: true })).toEqual({ nonsense: true });
    expect(sheetOf(null)).toBeNull();
  });

  it("finds the sheet's HEADER inside a whole round's file, not only its answers", () => {
    // `host-diff.mjs` unwrapped the answers and then read `source` and
    // `requirementSets` off the OUTER object, so the round file the pane
    // actually writes reported `REAL HOST ?` and `requirement sets: unknown`
    // above a page of real answers. Which host it was, and whether it offers
    // PowerPointApi 1.8, is what several open questions turn on.
    const round = {
      build: "3d17165 · 2026-08-10 12:43Z",
      hostAnswers: {
        kind: "powerchart-host-answers",
        source: "PowerPoint · OfficeOnline · 0.0.0.0",
        build: "3d17165 · 2026-08-10 12:43Z",
        requirementSets: ["1.1", "1.8"],
        answers: [{ id: "untrack-available", question: "?", answer: "yes", ms: 1 }],
      },
      selftest: { scenarios: [] },
    };
    expect(sheetOf(round).source).toBe("PowerPoint · OfficeOnline · 0.0.0.0");
    expect(sheetOf(round).requirementSets).toContain("1.8");
    // And the answers still come out of the same shape, through the same seam.
    expect(answersOf(round)).toEqual({ "untrack-available": "yes" });
  });

  it("still comes back complete from a host at its worst", async () => {
    // The sheet's whole value is that a misbehaving host still produces one.
    // A probe run that gave up half way would be worthless exactly when it
    // was needed.
    installHost([makeSlide("s1")]);
    applyWebProfile();
    const sheet = await runHostProbes("fake-web-profile", "test");
    // Every unconditional question is present. Not an exact COUNT: a follow-up
    // fires whenever its trigger's answer matches, so the total is a property
    // of what the host said, while "nothing was dropped" is the invariant this
    // test is named for.
    const ids = sheet.answers.map((a) => a.id);
    for (const id of ALWAYS_ASKED_IDS) expect(ids, `${id} was dropped`).toContain(id);
    for (const a of sheet.answers) {
      expect(a.answer, `${a.id} gave no answer`).toBeTruthy();
      expect(a.question, `${a.id} lost its question`).toBeTruthy();
      expect(a.ms).toBeGreaterThanOrEqual(0);
    }
  });

  /**
   * The property that makes these answers worth reading: no probe holds a
   * proxy across a sync, so a host that refuses stale ones changes nothing.
   *
   * `strictTags` and `strictGroup` are the two faults that model exactly what
   * PowerPoint on the web does to a proxy older than its sync — and four
   * questions on the 2026-08-04 sheet were answering those faults rather than
   * their own questions. `tag-on-group-survives` said NO, which read as "no
   * chart in any deck is re-editable"; `group-reports-its-children` threw
   * PropertyNotLoaded; `tags-add-same-key-twice` and `addgroup-returns-usable`
   * came back undefined. Every one of them wrote or read through a handle it
   * had been given a sync earlier.
   *
   * Deliberately NOT the whole web profile: `refuseGroups` and `hollowReads`
   * change the answers legitimately, and a guard that tolerated those could not
   * tell them from a probe that had gone stale again.
   *
   * What it does NOT prove, said plainly: the fake refuses a proxy older than
   * TWO syncs, and three of those four probes were exactly one sync late, so
   * only `tags-add-same-key-twice` goes red here without the rewrite. The
   * others rest on the same rule and on the real host's own answer to
   * `shape-add-held-slide-proxy` — one sync was enough there. Tightening the
   * fake to one sync for SHAPES would make this guard catch all four, and
   * nobody has asked a host whether that is true; `shape-proxy-survives-one-
   * sync` only establishes it at two. That question is the next sheet's.
   */
  it("answers the same under a host that refuses every stale proxy", async () => {
    installHost([makeSlide("s1")]);
    faults.strictTags = true;
    faults.strictGroup = true;
    try {
      const sheet = await runHostProbes("fake-strict-proxies", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      // One question is SUPPOSED to move here, and it is the one whose whole
      // job is to measure how long a handle lasts: under `strictTags` a proxy
      // is refused the moment it is more than one sync old, so
      // `how-many-syncs-a-creation-handle-survives` answers `refused-after-1`
      // where the healthy fake answers `survives-8`. Folding it into the
      // baseline would assert that the age fault changes nothing, which is the
      // opposite of what the fault is for.
      expect(answers["how-many-syncs-a-creation-handle-survives"], "the age fault did not reach the age question").toBe(
        "refused-after-1",
      );
      // TWO questions move, for one reason. `collection-read-poisons-the-creation-handle`
      // spends a sync on the collection read it exists to make, so under
      // `strictTags` its handle is more than one sync old by the time it writes
      // and the write is refused — by the AGE rule, not by the read. Asserting
      // the move keeps that visible; folding it into the baseline would claim
      // the age fault leaves it alone, which is false and would hide the day
      // this probe stops spending that sync.
      expect(
        answers["collection-read-poisons-the-creation-handle"],
        "the age fault did not reach the question that spends a sync",
      ).toBe("refused");
      // THREE, and this one is a warning about how its answer reads. Under
      // `strictTags` `does-a-failed-group-poison-the-tag` says
      // `refused-after-group` — and the group had nothing to do with it: the
      // question ages its handles two syncs on purpose, so the age rule refuses
      // the write on its own. On a real host the same words could mean either
      // thing, and what separates them is the CONTROL:
      // `tag-the-creation-proxy-a-sync-later` answers `yes` there, so an aged
      // creation handle demonstrably takes a write and age is excluded. Read
      // them together or not at all.
      expect(
        answers["does-a-failed-group-poison-the-tag"],
        "the age fault did not reach the question that ages its own handles",
      ).toBe("refused-after-group");
      const moves = [
        "how-many-syncs-a-creation-handle-survives",
        "collection-read-poisons-the-creation-handle",
        "does-a-failed-group-poison-the-tag",
      ];
      const without = (o: Record<string, unknown>) =>
        Object.fromEntries(Object.entries(o).filter(([k]) => !moves.includes(k)));
      expect(without(answers)).toEqual(without(FAKE_BASELINE));
    } finally {
      faults.strictTags = false;
      faults.strictGroup = false;
    }
  });

  it("carries the requirement sets, without which no answer can be read", async () => {
    // The same verdict means different things on 1.4 and on 1.10. A sheet that
    // did not say which would be uninterpretable a week later.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    expect(sheet.kind).toBe("ssf-charts-host-answers");
    expect(sheet.requirementSets.length).toBeGreaterThan(0);
    expect(sheet.requirementSets).toContain("1.5");
  });

  it("leaves the deck exactly as it found it", async () => {
    // A diagnostic that litters is one people stop running. Probes add shapes,
    // groups and tags; all of it belongs to a scratch slide that goes back.
    const deck = [makeSlide("s1"), makeSlide("s2")];
    installHost(deck);
    const { slideCount } = await import("../src/render/powerpoint");
    const before = await slideCount();
    await runHostProbes("fake", "test");
    expect(await slideCount(), "the probe run left a slide behind").toBe(before);
  });

  /**
   * The failure a real host actually produced, and the reason this run now
   * replaces a scratch slide it has lost.
   *
   * PowerPoint on the web resolved the scratch slide's id for the first
   * question and refused it for the other thirteen. Every one of those thirteen
   * was recorded as `"threw"` — a legitimate answer to several of these
   * questions, and one the diff tool compares against the fake's — so a sheet
   * that had asked ONE question came back looking like thirteen host
   * divergences. The questions were never put.
   */
  it("asks every question even when the host keeps losing the scratch slide", async () => {
    installHost([makeSlide("s1")]);
    // Each new slide answers to its id six times and then denies existing.
    //
    // Was two, when every probe shared the one slide handle `withProbeContext`
    // resolved; then four; now six. The number tracks one thing — the most
    // slide lookups a single question can need — and it grows every time a
    // probe stops holding a handle across a sync, because a handle per batch
    // is a lookup per batch. The longest are `tags-add-same-key-twice` and
    // `tag-on-group-survives` at four batches each, plus the liveness check,
    // plus the verify inside `addScratchSlide` for a replacement: six.
    //
    // It has to be the SMALLEST lease that lets a replacement slide carry a
    // whole question, and no larger. Too tight and the sheet comes back
    // incomplete for a reason that has nothing to do with recovery; too loose
    // and the fault never bites — which is how a count-based version of this
    // once passed against the very code it was written to falsify.
    faults.newSlideResolvesTimes = 6;
    try {
      const sheet = await runHostProbes("fake-loses-slides", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      // COMPLETENESS is the invariant, and it is the one this test is named for:
      // every question got put, on a host that takes the slide away between
      // them. Nothing here may come back in the never-asked vocabulary.
      for (const id of ALWAYS_ASKED_IDS) expect(Object.keys(answers), `${id} was dropped`).toContain(id);
      for (const [id, answer] of Object.entries(answers))
        expect(NOT_ASKED_WORDS, `${id} was never put — the replacement path did not carry it`).not.toContain(answer);
      // This USED to assert `answers` equalled `FAKE_BASELINE` outright, and
      // that assertion is gone deliberately rather than by accident.
      //
      // The lease is per SLIDE and questions share one until it fails, so how
      // much budget a question inherits depends on what ran before it. Moving
      // `binding-names-shape-later` from position 20 to 6 therefore left
      // `shape-resolve-held-slide-proxy` a partly-spent slide, and it answered
      // `threw` instead of `yes` — a correct answer to "can you resolve a shape
      // through a handle this host has stopped honouring", and a real one, not
      // a never-asked. Value-equality here was quietly pinning probe ORDER.
      //
      // Raising the lease until it went green is the one thing not done: the
      // comment above sets it as the SMALLEST lease that carries the longest
      // question, and tuning it to accommodate a reorder is how a count-based
      // version of this test once passed against the code it was written to
      // falsify. The values are still pinned, on the healthy host, by "answers
      // every question, and says what it claims" above.
      expect(answers["shape-add-fresh-slide-proxy"], "a question early enough to be unaffected drifted").toBe(
        FAKE_BASELINE["shape-add-fresh-slide-proxy"],
      );
      // KNOWN HOLE, and it predates this edit: nothing here proves the
      // REPLACEMENT path ran, though the test is named for it.
      //
      // Measured, not suspected. Disabling the post-not-asked replacement
      // outright — `if (recovered)` never taken, and `addScratchSlide` forced to
      // null — leaves this test green, and left it green against the
      // `toEqual(FAKE_BASELINE)` assertion that used to stand here too. At a
      // lease of six the original slide carries almost every question, so the
      // recovery is never needed and never exercised.
      //
      // A `took > 1` check on the cleanup row was tried and thrown away: it
      // passes under both mutations, because the second pass and the top-of-loop
      // re-acquire take slides of their own. A guard that survives the mutation
      // it is named for is decoration.
      //
      // Closing it needs a fault that forces a not-asked and then relents —
      // `newSlideRefusedForFirst` is the shape of it — which is its own change,
      // not a rider on a probe reorder.
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("takes every replacement scratch slide back out again", async () => {
    // A run that recovers by adding slides is a run that can litter several. It
    // has to clean up all of them, not the last one it happened to be holding.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    const { slideCount } = await import("../src/render/powerpoint");
    const before = await slideCount();
    faults.newSlideResolvesTimes = 2;
    try {
      await runHostProbes("fake-loses-slides", "test");
      expect(await slideCount(), "left its replacement slides in the deck").toBe(before);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  /**
   * The sheet's whole promise is that it comes back from a misbehaving host.
   *
   * Every QUESTION was bounded for that reason — `PROBE_BUDGET_MS`, "so the
   * sheet from a misbehaving host is the one worth having". The scratch-slide
   * lifecycle around them was not: `addScratchSlide`, the mid-run replacement,
   * and the cleanup all reached `slideIds`/`slideCount`/`deleteSlideById`,
   * whose syncs carried a label and no deadline. One silent deck-listing read
   * and `runHostProbes` never resolved — with `answers` already complete in
   * memory and never returned.
   */
  it("comes back with its sheet even when the host goes silent at cleanup", async () => {
    installHost([makeSlide("s1")]);
    const { _setReadbackTimeoutForTest } = await import("../src/render/powerpoint");
    _setReadbackTimeoutForTest(20);
    faults.wedgeAfterSyncs = 0;
    _setProbeBudgetForTest(20);
    try {
      const sheet = await Promise.race([
        runHostProbes("fake-goes-silent", "test"),
        new Promise<"never came back">((r) => setTimeout(() => r("never came back"), 800)),
      ]);
      expect(sheet, "the probe run never returned its answers").not.toBe("never came back");
      expect(typeof sheet === "object" && sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
    }
  });

  /**
   * When the host will not keep ANY slide, the sheet has to say that once per
   * question in a word no probe can produce — never `"threw"`, which is a real
   * answer to several of them and would read as a host divergence.
   */
  /**
   * The second half of the same lesson, and it cost a second real sheet.
   *
   * `no-scratch-slide` covers a slide the host will not resolve. It says
   * nothing about a slide that resolves perfectly and then refuses to take a
   * shape — which is what PowerPoint on the web did on 2026-08-04, because
   * `withProbeContext` handed every probe the slide proxy IT had resolved, and
   * a freshly-added slide's by-id handle is good for exactly one sync there.
   * The six questions that resolved a handle of their own were answered; all
   * eight that wrote through the held one failed, and every one of those was
   * recorded as `"threw"` or `"silent"` — both real answers to those questions.
   * `npm run host-diff` duly reported eight host divergences from a sheet that
   * had asked six questions.
   */
  it("keeps asking when a held slide handle is what the host refuses", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
    // Every question answered, despite the fake now refusing a held by-id
    // handle exactly as the host does — because no probe holds one any more.
    // The single `threw` in the baseline is the question that is ABOUT holding
    // one, which is asked on purpose.
    expect(answers).toEqual(FAKE_BASELINE);
    expect(answers["shape-add-held-slide-proxy"], "the question stopped being asked").toBe("threw");
  });

  it("never reports a setup the host refused as an answer to the question", async () => {
    // A probe that could not get its shapes must say so in a word no probe can
    // produce. `"threw"` and `"silent"` are real answers here, and a diff
    // compares them against the fake's — which is how eight questions nobody
    // asked came back looking like eight host divergences.
    installHost([makeSlide("s1")]);
    faults.refuseShapeAdds = true;
    try {
      const sheet = await runHostProbes("fake-refuses-shapes", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      const needShapes = [
        "group-children-via-getcount",
        "shape-proxy-survives-one-sync",
        "shape-resolve-held-slide-proxy",
        "shapes-items-count-honest",
        "shapes-items-via-positional-slide",
        // `tags-add-same-key-twice` is deliberately NOT here any more. It asks
        // about tag semantics, not about shapes, and it now asks on the scratch
        // SLIDE — so a host that refuses every shape add still answers it, which
        // is the whole point of the change. Leaving it in this list would assert
        // that a question needs a shape when it no longer does, and would fail
        // the moment the question started working.
        "tags-on-fresh-shape",
        // Needs a shape to create, read an id off, and then re-fetch by that id.
        "tag-through-refetched-shape",
        // The creationId questions need a shape to read it off. The FAKE DOES
        // NOT MODEL `creationId` and deliberately still does not: a double that
        // handed back a plausible id would make them read `yes`, `stable` and
        // `kept` in the suite while saying nothing whatever about the real host,
        // which is the one thing they exist to ask.
        //
        // The FIRST question to run after a refused add is not in this list, and
        // the reason is ordering rather than the probe: it absorbs the
        // `pendingHostError` that add left queued, so the throw reaches it
        // outside `scratchShapes` and it answers `threw` — truthfully — rather
        // than `no-scratch-shape`. That is `shapes-by-index-vs-items` today and
        // was `creationid-on-fresh-shape` before it was inserted ahead of it.
        // Everything after the error is consumed demonstrates the contract this
        // list is about.
        "creationid-on-fresh-shape",
        "creationid-survives-a-sync",
        "creationid-survives-grouping",
        // Needs a shape to age.
        "how-many-syncs-a-creation-handle-survives",
        // Needs a shape to hold a handle onto across the collection read.
        "collection-read-poisons-the-creation-handle",
        "delete-then-lookup",
        "addgroup-returns-usable",
        "group-reports-its-children",
        "group-of-existing-shape-readable",
        "picture-then-shape-read",
        // Needs six shapes to see which of them a short read keeps.
        "which-end-a-short-read-drops",
        // Needs two shapes to offer addGroup something to refuse.
        "does-a-failed-group-poison-the-tag",
        // Needs three shapes for a short read to be short OF something.
        "how-many-collection-reads-a-context-survives",
        // The binding question needs a shape to bind, so a host that refuses
        // every add has told it nothing. Its first version answered `add-threw`
        // here — a statement about `bindings.add` from a run in which
        // `bindings.add` was never reached — which is why the shape add, the
        // binding call and the commit are now three separate failure points.
        "binding-names-shape-later",
        // Needs a rectangle to rotate, and a hexagon to try to draw. Both call
        // `probeShapes` inside their own try/catch, which is only safe because
        // `threw` re-throws `ProbeSetupFailed`; without that they would report a
        // refused add as `threw` — an opinion about rotation from a run in which
        // nothing was ever rotated. That contract is what this list asserts.
        "rotation-keeps-the-unrotated-box",
        "named-preset-resolves",
      ];
      for (const id of needShapes) expect(answers[id], `${id} claimed a host answer`).toBe("no-scratch-shape");
      // And the questions that need no shape are still answered — one refusal
      // must not cost the sheet.
      expect(answers["getcount-populates-same-sync"]).toBe("yes");
      expect(answers["untrack-available"]).toBe("no");
      // A never-asked question is not a divergence, and this is the tool that
      // used to call it one. The shape-add questions ARE divergences — asking
      // whether this host takes a shape is their whole job, and "no" is an
      // answer. Two of the three: the held-handle question already expects a
      // refusal, so a host that refuses every add agrees with the fake there.
      const d = diffAnswers(answers, FAKE_BASELINE);
      expect(d.differ.map((x: { id: string }) => x.id).sort()).toEqual([
        "shape-add-fresh-getitem-slide",
        "shape-add-fresh-slide-proxy",
        "shape-add-positional-slide-proxy",
        // WHICHEVER QUESTION RUNS FIRST AFTER A REFUSED ADD absorbs the
        // `pendingHostError` that add left queued, and answers `threw` where the
        // baseline says otherwise. It was `creationid-on-fresh-shape`; adding
        // `shapes-by-index-vs-items` ahead of it moved the artifact here.
        //
        // Worth keeping as a divergence rather than smoothing away: it is a
        // truthful answer that differs from the fake's, which is what a
        // divergence IS. But it is a property of ORDERING, not of either
        // question — so if a probe is ever inserted before this one, expect the
        // name in this list to move again rather than a new fault to have
        // appeared.
        "shapes-by-index-vs-items",
      ]);
      expect(d.notAsked.map((n: { id: string }) => n.id).sort()).toEqual([...needShapes].sort());
    } finally {
      faults.refuseShapeAdds = false;
    }
  });

  it("says a question was never put, rather than inventing a host answer", async () => {
    installHost([makeSlide("s1")]);
    faults.newSlideResolvesTimes = 1; // one lookup each: addScratchSlide verifies, the probe context fails
    try {
      const sheet = await runHostProbes("fake-loses-slides", "test");
      // Every always-asked question has a row, and anything EXTRA is a partner
      // that fired. Length alone said the same thing until a follow-up existed
      // whose trigger answers on a slideless host — then the count grew for a
      // reason this property does not care about, and the assertion failed on
      // an unrelated probe being added.
      const ids = sheet.answers.map((a) => a.id);
      expect(
        ALWAYS_ASKED_IDS.filter((id) => !ids.includes(id)),
        "a question the run always puts has no row at all",
      ).toEqual([]);
      expect(
        ids.filter((id) => !ALWAYS_ASKED_IDS.includes(id) && !FOLLOW_UP_IDS.includes(id)),
        "the sheet carries a row for something that is neither always asked nor a partner",
      ).toEqual([]);
      // Every question that NEEDS a slide. Two rows are not those:
      //
      // - the cleanup row, which reports what became of the slides the run
      //   borrowed — a fact about the host whether or not a single question got
      //   put, and reading it as an answer would make this assertion demand
      //   that the cleanup fail too;
      // - a `noSlideNeeded` question, which is answerable with no slide at all
      //   and must not inherit this failure. Charging it one is how the
      //   2026-08-08 `a546897` sheet lost `untrack-available`, at 43 seconds,
      //   after the host had recovered and answered four questions in a row.
      //
      // The excused set is DERIVED from the probes rather than listed here, so
      // it cannot drift from what the code actually declares — and the loop
      // below is what stops that from being a way to silence this test. An
      // excusal has to be earned: a question marked `noSlideNeeded` must come
      // back with a real answer on a host that has no slide to give, and one
      // that answers `no-scratch-slide` anyway is mismarked and says so here.
      const excused = new Set([SCRATCH_CLEANUP_ID, ...NO_SLIDE_NEEDED_IDS]);
      for (const a of sheet.answers.filter((x) => !excused.has(x.id)))
        expect(a.answer, `${a.id} claimed a host answer`).toBe("no-scratch-slide");
      expect(NO_SLIDE_NEEDED_IDS.length, "no question claims to be answerable without a slide").toBeGreaterThan(0);
      for (const id of NO_SLIDE_NEEDED_IDS) {
        const a = sheet.answers.find((x) => x.id === id);
        expect(a, `${id} is marked noSlideNeeded and produced no row at all`).toBeTruthy();
        expect(NOT_ASKED_WORDS, `${id} is marked noSlideNeeded but was charged for a slide anyway`).not.toContain(
          a!.answer,
        );
      }
      // And it has to be there. A run that could not ask anything is precisely
      // the run that churns through scratch slides, so a sheet silent about
      // them is silent about the only thing that run left behind.
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID);
      expect(cleanup, "the sheet says nothing about the slides the run borrowed").toBeTruthy();
      expect(cleanup!.detail).toMatch(/scratch slide/);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("puts the questions a dead window swallowed, once the host is answering again", async () => {
    // Two consecutive real rounds lost roughly HALF their questions this way —
    // 14 of 27, then 13 of 28 — every one of them `no-scratch-slide`, and the
    // immediate retry beside each failure fired 21 times between them and
    // changed nothing. It could not: this host's ability to resolve a freshly
    // added slide comes and goes in windows of about fifteen seconds, so a retry
    // issued straight away lands inside the window that just refused.
    //
    // What both rounds also show is the recovery, in the same run: 2026-08-09
    // answered positions 17, 18, 22, 24 and 26 after losing 10 through 16. So
    // the questions were answerable; the run simply asked them all at the wrong
    // moment and never went back.
    //
    // `newSlideRefusedForFirst` is that window, counted in slides because
    // nothing the probe does consults a clock. Six is comfortably wider than the
    // handful of slides the opening questions spend, so questions genuinely fall
    // into it — and it closes, which is the half the fake could not express
    // before.
    installHost([makeSlide("s1")]);
    faults.newSlideRefusedForFirst = 12;
    try {
      const sheet = await runHostProbes("fake-loses-then-recovers", "test");
      const lost = sheet.answers.filter((a) => NOT_ASKED_WORDS.includes(a.answer));
      // Asked on the PROPERTY, not on which rung delivered it. A run makes three
      // passes now, so a question swallowed by an early window is usually put
      // again long before the end-of-run sweep — and this test exists to say the
      // window does not cost the question, not to say which pass rescued it.
      const recovered = sheet.answers.filter(
        (a) => !NOT_ASKED_WORDS.includes(a.answer) && (a.samples ?? []).some((x) => NOT_ASKED_WORDS.includes(x.answer)),
      );
      expect(recovered.length, "nothing was re-asked after the window closed").toBeGreaterThan(5);
      expect(
        lost.map((a) => a.id),
        "a question the host would have answered was still filed unanswerable",
      ).toEqual([]);
      // A rescued row carries a real answer, not a relabelled failure.
      for (const a of recovered)
        expect(NOT_ASKED_WORDS, `${a.id} "recovered" into ${a.answer}`).not.toContain(a.answer);
    } finally {
      faults.newSlideRefusedForFirst = 0;
    }
  });

  /**
   * The binding question's two ways of being refused, and why they are two.
   *
   * A binding has to be made in the batch that CREATES its shape — a proxy one
   * sync old is refused on this host — so the batch carries two things and its
   * failure is attributable to neither. That is not a hypothetical: the
   * 2026-08-09 evening round came back `UnexpectedError` from the commit in 1.3
   * seconds, and the probe honestly reported "never asked" while
   * `shape-add-fresh-slide-proxy` answered `yes` two rows above it.
   *
   * The control arm is what makes the difference readable. The same batch minus
   * the binding runs first, on the same slide; if it commits and the bound one
   * does not, the binding is the only variable left.
   */
  it("blames the binding only when the same batch without one just worked", async () => {
    installHost([makeSlide("s1")]);
    faults.refuseBindings = "sync";
    try {
      const sheet = await runHostProbes("fake-refuses-bindings-at-sync", "test");
      const row = sheet.answers.find((a) => a.id === "binding-names-shape-later");
      expect(row?.answer, "a refusal the control had already ruled out is an ANSWER, not a missing question").toBe(
        "commit-threw",
      );
      expect(row?.detail).toMatch(/without a binding committed seconds earlier/);
    } finally {
      faults.refuseBindings = null;
    }
  });

  it("tells a binding refused at the call apart from one refused at the commit", async () => {
    // Same API, two different facts about it: `bindings.add` objecting on the
    // spot is not the host rejecting the batch that carried it, and a single
    // word for both would put them in one bucket in the diff.
    installHost([makeSlide("s1")]);
    faults.refuseBindings = "call";
    try {
      const sheet = await runHostProbes("fake-refuses-bindings-at-call", "test");
      const row = sheet.answers.find((a) => a.id === "binding-names-shape-later");
      expect(row?.answer).toBe("add-threw");
      expect(row?.detail).toMatch(/BindingCollection\.add/);
    } finally {
      faults.refuseBindings = null;
    }
  });

  it("claims the slide it appended even when the host renumbers another one", async () => {
    // The failure that cost a whole second pass on 2026-08-10: three attempts,
    // three `scratch slide did not land ... fresh=2`, five questions never
    // re-asked. And it is not rare — seven observations across four rounds,
    // every one the same arithmetic:
    //
    //   before=20 after=21 fresh=2      before=3  after=4  fresh=2
    //   before=37 after=38 fresh=2      before=18 after=19 fresh=2
    //   before=21 after=22 fresh=2      before=19 after=20 fresh=2
    //   before=20 after=21 fresh=2
    //
    // The deck grows by ONE and TWO ids read as new, so an existing id has to
    // have changed. `addScratchSlide` refused both candidates — right about the
    // danger (claiming a slide that was already there would later delete the
    // user's work) and wrong about this host, where nothing else is adding
    // slides at all.
    //
    // What makes it claimable is position: `add()` appends, so ours is last and
    // a renumbered slide keeps the place it always had.
    installHost([makeSlide("s1")]);
    faults.renumbersOnAdd = true;
    try {
      const sheet = await runHostProbes("fake-renumbers-on-add", "test");
      const lost = sheet.answers.filter((a) => NOT_ASKED_WORDS.includes(a.answer));
      expect(
        lost.map((a) => a.id),
        "the churned id list cost questions their slide",
      ).toEqual([]);
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID)!;
      expect(cleanup.answer, `slides were claimed but not given back — "${cleanup.detail}"`).toBe("all");
    } finally {
      faults.renumbersOnAdd = false;
    }
  });

  it("refuses to claim a slide that was already in the deck", async () => {
    // The safety half of the rule above, and it needed its own fault to exist:
    // with a fake that always appends, the "is the last slide actually new"
    // check could be deleted and every test stayed green.
    //
    // Claiming by position is only safe because `add()` appends. When the added
    // slide lands somewhere else — and this deck is documented as reordering
    // under load, which is why `blankSlides` is reported by position at all —
    // the last slide belongs to the USER. Giving up costs a question. Claiming
    // it costs them a slide, later, silently, when the probe tidies up.
    // BOTH faults, and it takes both to build the dangerous shape. Front-insert
    // alone still leaves exactly one fresh id, so the positional branch is never
    // reached and the check under test never runs — the first version of this
    // armed only that one and passed against the mutation it names.
    // Renumbering is what forces the fallback; front-insert is what puts a
    // slide that is not ours in the last position once we get there.
    const deck = [makeSlide("theirs-1"), makeSlide("theirs-2")];
    installHost(deck);
    faults.renumbersOnAdd = true;
    faults.addsAtFront = true;
    try {
      // By IDENTITY, not by id. `renumbersOnAdd` renames these slides as it
      // goes, so an id lookup fails whether or not they survived — the first
      // version of this asserted on ids and went red against correct code.
      const [one, two] = deck;
      await runHostProbes("fake-adds-at-front", "test");
      expect(deck.includes(one), "the probe deleted a slide it never added").toBe(true);
      expect(deck.includes(two), "the probe deleted a slide it never added").toBe(true);
    } finally {
      faults.renumbersOnAdd = false;
      faults.addsAtFront = false;
    }
  });

  it("takes a fresh slide after a question that wrecks the one it used", async () => {
    // The finding seven rounds took to see, and only because a reorder ran the
    // control by accident.
    //
    // The slot after `shape-add-held-slide-proxy` has never produced an answer.
    // For six rounds that slot held `shape-resolve-held-slide-proxy`, and its
    // 0/6 was put down to its own nature — it reads ids back off setup shapes,
    // which is a known-bad thing to do here. Then `binding-names-shape-later`
    // was moved into the slot for unrelated reasons and failed identically, in
    // 397ms, at the liveness check, before its own code ran:
    //
    //   12-17  #5 shape-add-held-slide-proxy threw | #6 shape-resolve-…  never put
    //   18     #5 shape-add-held-slide-proxy threw | #6 binding-names-…  never put
    //
    // Two questions, one slot, seven for seven: it is the slot. #5 writes
    // through a slide proxy the host has stopped honouring — that IS its
    // question — and the slide does not survive it. Because it ANSWERS, none of
    // the not-asked replacement paths ever noticed.
    //
    // Asserted structurally rather than through a fault, because the fake does
    // not model the poisoning and inventing it would be asserting a mechanism
    // nobody has measured. What IS measured is the rule: a question that
    // declares it wrecks the slide must not hand that slide to the next one.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID)!;
    const took = Number(/of (\d+) scratch/.exec(cleanup.detail ?? "")?.[1] ?? 0);
    // A healthy fake loses nothing, so without the rule ONE slide serves the
    // whole run. Any second slide here is the burnt one being given up.
    expect(took, `the run never replaced the wrecked slide — "${cleanup.detail}"`).toBeGreaterThan(1);
  });

  it("keeps sweeping after the host refuses it a slide, instead of bailing on the first", async () => {
    // Round 15 is this bug, in the owner's own deck: `second pass over the
    // questions that were never put {count: 10}`, ONE slide attempt two seconds
    // later, `scratch slide landed but its id will not resolve`, and then
    // nothing. Ten questions decided by a single coin flip on a host whose
    // slide resolution is known to flap — and the `break` was silent, so the
    // log could not tell "tried them all and failed" from "gave up at once".
    //
    // A window of 28 is the discriminator, measured rather than picked: it
    // closes DURING the sweep, so the first attempts are refused and later ones
    // succeed. With the loop continuing past a refusal that rescues 19 of 20;
    // bailing on the first rescues none of them.
    installHost([makeSlide("s1")]);
    faults.newSlideRefusedForFirst = 28;
    try {
      const sheet = await runHostProbes("fake-refuses-then-relents-mid-sweep", "test");
      // The property, as above: refused once, asked again, answered in the end.
      const rescued = sheet.answers.filter(
        (a) => !NOT_ASKED_WORDS.includes(a.answer) && (a.samples ?? []).some((x) => NOT_ASKED_WORDS.includes(x.answer)),
      );
      expect(rescued.length, "one refused slide ended the whole sweep").toBeGreaterThan(10);
      const lost = sheet.answers.filter((a) => NOT_ASKED_WORDS.includes(a.answer));
      expect(lost.length, `still lost: ${lost.map((a) => a.id).join(", ")}`).toBeLessThan(3);
    } finally {
      faults.newSlideRefusedForFirst = 0;
    }
  });

  it("gives up on a second pass the host is never going to answer", async () => {
    // The other side of the rung, and the one that keeps it from being a way to
    // spend three minutes on a dead host. A window wider than the whole run
    // never closes, so every question stays unasked and the sweep must stop
    // asking rather than work through the list. The sheet still comes back
    // complete — that is the invariant every other case here defends too.
    installHost([makeSlide("s1")]);
    faults.newSlideRefusedForFirst = 500;
    try {
      const sheet = await runHostProbes("fake-never-recovers", "test");
      const missing = ALWAYS_ASKED_IDS.filter((id) => !sheet.answers.some((a) => a.id === id));
      expect(missing, "the sheet came back short — a question the run always puts has no row").toEqual([]);
      expect(sheet.answers.filter((a) => a.detail?.includes("second pass"))).toEqual([]);
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID);
      expect(cleanup, "a run that asked nothing still owes an account of the slides it borrowed").toBeTruthy();
    } finally {
      faults.newSlideRefusedForFirst = 0;
    }
  });

  it("reports the scratch slides it could NOT give back", async () => {
    // The failure this row exists for, and it is not hypothetical: a real round
    // on 2026-08-06 left 21 blank slides in the owner's deck, an earlier one
    // left 14, and both sheets came back saying nothing about it. The cleanup
    // loop already asked `deleteSlideById`, which re-reads the deck rather than
    // trusting a queued delete — it simply threw the boolean away, so the only
    // way to learn the probe litters was to open the deck afterwards.
    installHost([makeSlide("s1")]);
    faults.refuseSlideDelete = true;
    try {
      const sheet = await runHostProbes("fake-refuses-deletes", "test");
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID);
      expect(cleanup, "the sheet says nothing about the slides the run borrowed").toBeTruthy();
      expect(cleanup!.answer, `the host kept every slide and the sheet said "${cleanup!.answer}"`).toBe("none");
      expect(cleanup!.detail).toMatch(/left in the deck/);
    } finally {
      faults.refuseSlideDelete = false;
    }
  });
});

describe("what the pane says about a probe run", () => {
  it("tells the owner not to bother sending a run that found nothing", async () => {
    // The round trip this removes: download, send, wait for someone to run
    // `host-diff`, hear that the answers are the same as last time. Most probe
    // runs are that.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    expect(sheetNeedsAttention(sheet), "asked for a file that says nothing new").toBe(false);
    expect(describeHostSheet(sheet)).toContain("Nothing new");
  });

  it("names an answer nobody has declared, and asks for the file", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    // A host that answers one question differently from the fake, where the
    // difference is not in KNOWN_DIVERGENCES: the only kind of run worth
    // anyone's attention, and it should not have to be found by eye in a
    // seventeen-row JSON file.
    const changed = {
      ...sheet,
      answers: sheet.answers.map((a) => (a.id === "delete-then-lookup" ? { ...a, answer: "still-there" } : a)),
    };
    expect(sheetNeedsAttention(changed)).toBe(true);
    expect(describeHostSheet(changed)).toContain("NEW: delete-then-lookup");
    expect(describeHostSheet(changed)).toContain("worth sending");
  });

  it("counts a declared divergence as known, not as news", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    // `group-reports-its-children` is declared divergent, so a sheet reproducing
    // it is not a finding. It stands in for `tag-on-group-survives`, which this
    // test used until that question was retired — what is under test is the
    // COUNTING, so any declared id will do.
    const known = {
      ...sheet,
      answers: sheet.answers.map((a) => (a.id === "group-reports-its-children" ? { ...a, answer: "reports-1" } : a)),
    };
    expect(sheetNeedsAttention(known), "treated a declared divergence as news").toBe(false);
    expect(describeHostSheet(known)).toContain("1 known divergence");
  });

  it("asks for the file when a question was never put, however few diverge", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const incomplete = {
      ...sheet,
      answers: sheet.answers.map((a) => (a.id === "tags-on-fresh-shape" ? { ...a, answer: "no-scratch-shape" } : a)),
    };
    expect(sheetNeedsAttention(incomplete), "an incomplete sheet looked fine").toBe(true);
    expect(describeHostSheet(incomplete)).toContain("never put");
  });
});

describe("a host that will not name a shape in the batch that made it", () => {
  /**
   * The five questions the 2026-08-05 sheet never put, and why.
   *
   * They are exactly the five that read an id back — `shape-resolve-held-slide-
   * proxy`, `tags-add-same-key-twice`, `addgroup-returns-usable`,
   * `group-reports-its-children`, `tag-on-group-survives` — and no others.
   * Perfect correlation with `idsOf`, and no competing property explains it:
   * batch count does not, since `shapes-items-count-honest` and
   * `delete-then-lookup` span two batches each and were answered while three of
   * the five span two and failed.
   *
   * The id load used to ride in the same sync as the add, to save a round trip.
   * It now costs its own sync, and that is what makes these five answerable on
   * a host like this.
   */
  it("still asks every question, and answers them as a healthy host would", async () => {
    installHost([makeSlide("s1")]);
    faults.noIdInCreatingSync = true;
    // Both, because they are two halves of one behaviour: the load is not
    // answered, AND an unanswered property is unreadable rather than quietly
    // available. Without `strictShapeReads` the fake hands the id over anyway
    // and the first half is inert — which is exactly how a guard comes to pass
    // against the code it was written to falsify.
    faults.strictShapeReads = true;
    try {
      const sheet = await runHostProbes("fake-cannot-name-new-shapes", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      // The property, not the whole sheet: every question was PUT. Demanding
      // the healthy sheet would be wrong here — `strictShapeReads` legitimately
      // changes what some questions answer, and a guard that could not tell a
      // legitimate change from an unasked question is the confusion this file
      // exists to prevent.
      const neverPut = sheet.answers.filter((a) => a.answer === "no-scratch-shape").map((a) => a.id);
      expect(neverPut, "these questions were never put — the ids came back unreadable").toEqual([]);
      // And the five that read an id back reached their own vocabulary.
      for (const id of [
        "shape-resolve-held-slide-proxy",
        "tags-add-same-key-twice",
        "addgroup-returns-usable",
        "group-reports-its-children",
        "tag-on-group-survives",
      ]) {
        expect(NOT_ASKED_WORDS, `${id} did not get asked`).not.toContain(answers[id]);
      }
    } finally {
      faults.noIdInCreatingSync = false;
      faults.strictShapeReads = false;
    }
  });
});

describe("a probe run that has lost its scratch slide", () => {
  /**
   * Losing a slide must cost the questions it happened during, not the sheet.
   *
   * `scratchId` went null when a replacement also failed to resolve, and nothing
   * set it back — so one bad moment recorded `no-scratch-slide` for every
   * question after it, as though the host had been asked each one. That is the
   * same failure this file's whole design is against, one level up: an answer
   * that describes the run's own state rather than the host's.
   */
  it("takes another slide rather than writing off the rest of the sheet", async () => {
    installHost([makeSlide("s1")]);
    // Every new slide answers to its id exactly once: `addScratchSlide`'s own
    // verify spends it, so every probe context then fails its liveness check
    // and every replacement is written off immediately. The run should keep
    // trying, and keep saying honestly that it could not get a slide.
    faults.newSlideResolvesTimes = 1;
    try {
      const sheet = await runHostProbes("fake-loses-every-slide", "test");
      const missing = ALWAYS_ASKED_IDS.filter((id) => !sheet.answers.some((a) => a.id === id));
      expect(missing, "the sheet came back short — a question the run always puts has no row").toEqual([]);
      // The cleanup row and the slideless question aside — see the sibling case
      // above for both. One is not a question; the other does not need what
      // this fault withholds.
      const excused = new Set([SCRATCH_CLEANUP_ID, ...NO_SLIDE_NEEDED_IDS]);
      for (const a of sheet.answers.filter((x) => !excused.has(x.id))) {
        expect(a.answer, `${a.id} claimed a host answer`).toBe("no-scratch-slide");
      }
      expect(sheet.answers.find((a) => a.id === "untrack-available")?.answer).toBe("no");
      // This is the run that makes one scratch slide per question and throws
      // each away. How many came back is the whole finding, and until this row
      // existed the sheet did not carry it.
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID)!;
      expect(cleanup.detail, "the sheet counted no scratch slides at all").toMatch(/\d+ of \d+ scratch slide/);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("comes back at all when the host never answers the slide add", async () => {
    // `addScratchSlide` used to wrap `slides.add()` in a raw `await
    // context.sync()` — the only slide-add in the file without a deadline.
    // office-js#1650 is explicit that the promise can hang while the slide
    // lands, and a probe run that hangs there produces no sheet at all, which
    // is the one thing this diagnostic must never do.
    installHost([makeSlide("s1")]);
    const { _setReadbackTimeoutForTest } = await import("../src/render/powerpoint");
    _setReadbackTimeoutForTest(20);
    faults.wedgeAfterSyncs = 1;
    try {
      const sheet = await Promise.race([
        runHostProbes("fake-hangs-on-add", "test"),
        new Promise<"never came back">((r) => setTimeout(() => r("never came back"), 2000)),
      ]);
      expect(sheet, "the run hung on the slide add and produced no sheet").not.toBe("never came back");
      expect(typeof sheet === "object" && sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
    }
  });
});

/**
 * The partner question, asked in the same run.
 *
 * The repo's rule is "when two explanations fit the evidence, ask — do not
 * reason". Following it has meant writing the partner question, waiting for the
 * owner to run the probe again, and losing a session — twice, at a cost of two
 * full sheets. The reasoning was never the expensive part; the round trip was.
 */
describe("questions that ask their own follow-up", () => {
  const sheetRows = async () => (await runHostProbes("fake", "test")).answers;

  it("asks it when the answer admits two readings, and says which answer caused it", async () => {
    // `getItem` refusing a freshly-added slide has two readings that lead
    // opposite ways: getItem cannot name a NEW slide (so the everyday insert
    // path is fine and only one caller is at risk), or getItem is broken here
    // (so the insert path is broken for everyone). A sheet cannot tell them
    // apart; the pre-existing slide can.
    installHost([makeSlide("s1")]);
    faults.refuseGetItemOnNewSlide = true;
    try {
      const rows = await sheetRows();
      const asked = rows.find((r) => r.id === "shape-add-fresh-getitem-slide")!;
      const partner = rows.find((r) => r.id === "getitem-durable-slide");
      expect(asked.answer, "the fault did not provoke the answer this pair is about").not.toBe("yes");
      expect(partner, "the follow-up was never asked").toBeTruthy();
      // The discriminating result: getItem works perfectly on a slide that was
      // already in the deck, so the refusal is about the slide's newness.
      expect(partner!.answer).toBe("yes");
      expect(partner!.detail, "the sheet does not say the two rows are a pair").toContain(
        'asked because shape-add-fresh-getitem-slide answered "threw"',
      );
    } finally {
      faults.refuseGetItemOnNewSlide = false;
    }
  });

  it("asks the durable slide when the positional collection read comes back short", async () => {
    // THE ROUND-254 CONTROL. `shapes-items-via-positional-slide` reads a
    // collection on a slide this run added and re-acquired by position, and its
    // answer changed at round 254 because OUR commit `77f9ca4` changed how that
    // slide is obtained — not because PowerPoint changed. A short read has two
    // readings that lead opposite ways: the host under-reports COLLECTIONS, or
    // it under-reports OUR SLIDE and says nothing about the slides a user's
    // charts live on. Only a slide the run did not create can tell them apart.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    faults.hollowReads = 99;
    try {
      const rows = await sheetRows();
      const asked = rows.find((r) => r.id === "shapes-items-via-positional-slide")!;
      const partner = rows.find((r) => r.id === "durable-slide-lists-its-shapes");
      expect(asked.answer, "the fault did not provoke the answer this pair is about").not.toBe("at-least-5");
      expect(partner, "the follow-up was never asked").toBeTruthy();
      expect(partner!.detail, "the sheet does not say the two rows are a pair").toContain(
        "asked because shapes-items-via-positional-slide answered",
      );
    } finally {
      faults.hollowReads = 0;
    }
  });

  it("does not ask the durable slide when the collection read worked", async () => {
    // The other half, and the one that keeps this from being a question that is
    // always asked: on a host that enumerates its own slide there is nothing to
    // disambiguate, and the follow costs a deck read for no answer.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    const rows = await sheetRows();
    expect(rows.find((r) => r.id === "shapes-items-via-positional-slide")!.answer).toBe("at-least-5");
    expect(
      rows.find((r) => r.id === "durable-slide-lists-its-shapes"),
      "asked a follow-up nobody needed",
    ).toBeFalsy();
  });

  it("does not ask it when there is nothing to disambiguate", async () => {
    // An unconditional partner is just another probe and belongs in the list.
    // A follow-up earns its place by being worth asking only in the light of a
    // particular answer — so on a host that answers plainly it must cost
    // nothing at all.
    installHost([makeSlide("s1")]);
    const rows = await sheetRows();
    expect(rows.find((r) => r.id === "shape-add-fresh-getitem-slide")!.answer).toBe("yes");
    expect(
      rows.some((r) => r.id === "getitem-durable-slide"),
      "asked a follow-up nobody needed",
    ).toBe(false);
  });

  it("does not follow up a question that was never put", async () => {
    // `no-scratch-slide` is the probe's own vocabulary for "this never reached
    // its question". A partner to that would be a second question about the
    // probe's own setup, dressed as a fact about the host — precisely the
    // confusion `NOT_ASKED` exists to prevent, and worse here because the
    // follow-up's answer looks like a real finding in its own right.
    installHost([makeSlide("s1")]);
    faults.swallowAdds = 500;
    try {
      const rows = await sheetRows();
      expect(rows.find((r) => r.id === "shape-add-fresh-getitem-slide")!.answer).toBe("no-scratch-slide");
      expect(
        rows.some((r) => r.id === "getitem-durable-slide"),
        "followed up a question the host was never asked",
      ).toBe(false);
    } finally {
      faults.swallowAdds = 0;
    }
  });

  it("says it had no control rather than measuring a slide this run added", async () => {
    // On a deck the run has built entirely there is no pre-existing slide, and
    // the honest answer is that the control was unavailable. Falling back to a
    // scratch slide would be the control measuring the very thing it exists to
    // be compared against — the exact confusion the pair removes.
    installHost([]);
    faults.refuseGetItemOnNewSlide = true;
    try {
      const rows = await sheetRows();
      expect(rows.find((r) => r.id === "getitem-durable-slide")?.answer).toBe("no-durable-slide");
    } finally {
      faults.refuseGetItemOnNewSlide = false;
    }
  });

  it("still produces a sheet when the deck will not list itself at all", async () => {
    // The control read happens before the first question, so an unbounded one
    // took the whole run down with it — no sheet at all, which is the single
    // failure mode this file exists to prevent. A control nobody can name costs
    // the follow-up and nothing else.
    installHost([makeSlide("s1")]);
    const { _setReadbackTimeoutForTest } = await import("../src/render/powerpoint");
    faults.wedgeAfterSyncs = 0;
    _setReadbackTimeoutForTest(20);
    try {
      const ids = new Set((await sheetRows()).map((r) => r.id));
      for (const id of ALWAYS_ASKED_IDS) {
        expect(ids, `lost ${id} because the deck would not list itself`).toContain(id);
      }
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
    }
  });
});

describe("the questions added from the office-js tracker", () => {
  const sheetOfRows = async () => (await runHostProbes("fake", "test")).answers;

  it("asks about layouts on a slide the run did not add", async () => {
    // A freshly-added slide's handle is good for exactly one sync on this host,
    // so a load queued on the scratch slide and read after its own sync answers
    // "unreadable" whatever the layout API does. That is an answer about the
    // probe's plumbing wearing the clothes of a fact about layouts — the
    // failure this whole file exists to prevent, and the first draft of this
    // question had it.
    installHost([]);
    const rows = await sheetOfRows();
    expect(rows.find((r) => r.id === "slide-layout-readable")?.answer).toBe("no-durable-slide");
  });

  it("still answers the layout question when the deck has a settled slide", async () => {
    installHost([makeSlide("s1")]);
    const rows = await sheetOfRows();
    expect(rows.find((r) => r.id === "slide-layout-readable")?.answer).toBe("yes");
  });

  it("names a group's children through a handle resolved later, not the one that made it", async () => {
    // office-js#5849 is about `Shape.group` on a group found afterwards, which
    // is how `countGroupChildrenPage` finds one. `group-reports-its-children`
    // already asks in the batch that MADE the group, and the two answers are
    // different claims — one about the API, one about proxy age.
    installHost([makeSlide("s1")]);
    const rows = await sheetOfRows();
    expect(rows.find((r) => r.id === "group-of-existing-shape-readable")?.answer).toBe("2");
  });
});

describe("a host that answers every call and reads back nothing", () => {
  /**
   * The third kind of bad host, and the one with no test until now.
   *
   * Two are covered already: a host that THROWS (`refuseShapeAdds`) and one that
   * goes SILENT (`wedgeAfterSyncs`). This is the one in between — every call is
   * taken, every sync resolves, and then the properties those syncs were
   * supposed to populate are unreadable. It is the shape PowerPoint on the web
   * shows most often, and the shape a probe is most likely to mis-report:
   * "the value is not there" and "the value is false" are one line apart in
   * every one of these questions.
   *
   * The property, not the sheet. What each question answers here is allowed to
   * differ — an unreadable host is a legitimate finding — but a question that
   * silently vanishes, or an answer that is not a word, is the probe failing
   * rather than the host.
   */
  it("still puts every question, and answers each from its own vocabulary", async () => {
    installHost([makeSlide("s1")]);
    faults.strictShapeReads = true;
    faults.unansweredTagLoads = 500;
    faults.hollowNameReads = 500;
    try {
      const sheet = await runHostProbes("fake-unreadable", "test");
      const ids = sheet.answers.map((a) => a.id);
      for (const id of ALWAYS_ASKED_IDS) {
        expect(ids, `${id} vanished from the sheet on an unreadable host`).toContain(id);
      }
      for (const a of sheet.answers) {
        expect(typeof a.answer, `${a.id} answered with a ${typeof a.answer}`).toBe("string");
        expect(a.answer.length, `${a.id} answered with an empty string`).toBeGreaterThan(0);
        // A JS `undefined` reaching the sheet as a word is the specific bug this
        // guards. Every one of these questions reads a property that may not be
        // there, and `String(undefined)` is a perfectly good-looking answer that
        // the diff would compare against a real host forever.
        expect(a.answer, `${a.id} put a JS undefined in the sheet`).not.toMatch(/^undefined$/);
      }
    } finally {
      faults.strictShapeReads = false;
      faults.unansweredTagLoads = 0;
      faults.hollowNameReads = 0;
    }
  });
});

/**
 * The probe stops asking a host that has stopped answering.
 *
 * The self-test has had this rung for weeks; the probe had none, and
 * PowerPoint's own "Sorry, we ran into a problem" dialog is what exposed the
 * gap — the host was gone, the pane's elapsed timer kept counting, and the run
 * went on putting questions to a dead document, each one spending its full
 * budget to record an answer about nothing.
 *
 * Only the NO-FALSE-POSITIVE half is guarded here, and that is worth stating
 * rather than leaving to be discovered. Reaching the positive case needs a host
 * that takes a question and never answers, and the fake's wedge hangs
 * `addScratchSlide` first — the run stops before any question can miss its
 * deadline, so the test times out instead of exercising the breaker. A fault
 * that wedges question syncs while leaving slide adds alone is what it would
 * take, and that is a change to the fake with its own justification owed.
 */
describe("a host that has stopped answering", () => {
  /**
   * Why the breaker counts deadline misses and not `no-scratch-slide`.
   *
   * A round on 2026-08-08 answered `no-scratch-slide` FOUR times running — the
   * host had stopped resolving fresh slide ids for about fifteen seconds — then
   * recovered and answered five more questions, including two of the three this
   * project cares most about. Those questions were still being answered, in two
   * to six seconds; they just could not be set up. A breaker keyed on that
   * signal throws the recovery away.
   */
  it("keeps going through a host that refuses slides but still answers promptly", async () => {
    installHost([makeSlide("s1")]);
    faults.newSlideResolvesTimes = 1;
    try {
      const sheet = await runHostProbes("flaky-slides", "web");
      const notReached = sheet.answers.filter((a) => a.answer === "not-asked");
      expect(
        notReached,
        `gave up on a host that was still answering — ${notReached.length} question(s) abandoned`,
      ).toHaveLength(0);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  }, 60_000);
});

/**
 * The three group questions, on the host that could never answer them.
 *
 * They were unanswerable on the one host that matters, and `scratchShapes`' own
 * comment had already named the lead: the six questions this host never answers
 * are exactly the six that pass a `load`. They add two shapes, sync, read the
 * ids back — and it is that read the host refuses, the same refusal behind
 * `shape-proxy-survives-one-sync: unreadable` and the empty collection reads
 * everywhere else. The questions were never reached. They died in their own
 * setup and reported `no-scratch-slide`, which describes the probe rather than
 * the host.
 *
 * They matter more than their position in the list suggests. `contentShapes`
 * returns UNKNOWN_CONTENT for every grouped slide, which is what makes the
 * reconcile report a slide complete without counting it — so whether the web
 * host can count a group's children decides whether those verdicts can ever be
 * measurements. office-js#5849 is closed for inactivity and is a DESKTOP report;
 * nobody has established the web answer.
 */
describe("the group questions, when the host will not read an id back", () => {
  const GROUP_QUESTIONS = ["addgroup-returns-usable", "group-reports-its-children", "group-of-existing-shape-readable"];

  it("still asks them, and says which route the members came by", async () => {
    installHost([makeSlide("s1")]);
    // The host from the run logs: it will not read an id back off a shape this
    // run just added, however young the proxy — which is exactly what the strict
    // setup depends on. A large count, because every group question tries it.
    faults.refuseShapeIdLoads = 500;
    try {
      const sheet = await runHostProbes("no-id-readback", "web");
      const byId = Object.fromEntries(sheet.answers.map((a) => [a.id, a]));
      for (const id of GROUP_QUESTIONS) {
        const a = byId[id];
        expect(a, `${id} is missing from the sheet entirely`).toBeTruthy();
        expect(
          ["no-scratch-slide", "no-scratch-shape", "not-asked"],
          `${id} still died in its own setup: ${a.answer} — ${a.detail ?? "no detail"}`,
        ).not.toContain(a.answer);
        // A fallback that does not SAY it is a fallback is the trap this whole
        // file is about: grouping same-batch proxies is the friendliest case a
        // host can be given, and an answer from it says nothing about grouping
        // by id. The sheet has to keep them apart.
        expect(a.detail ?? "", `${id} did not record how its members were obtained`).toMatch(
          /members via (ids|same-batch)/,
        );
      }
    } finally {
      faults.refuseShapeIdLoads = 0;
    }
  }, 60_000);

  it("prefers the strict route when the host allows it", async () => {
    // The negative control. If the fallback ran unconditionally these questions
    // would quietly stop testing what production does — grouping by id — and
    // nothing would say so.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("healthy", "web");
    const byId = Object.fromEntries(sheet.answers.map((a) => [a.id, a]));
    for (const id of GROUP_QUESTIONS) {
      expect(byId[id]?.detail ?? "", `${id} took the fallback on a host that never needed it`).toContain(
        "members via ids",
      );
    }
  }, 60_000);
});

describe("a question that needs no slide", () => {
  /**
   * The liveness check `withProbeContext` runs before every question is, on
   * PowerPoint web, the single most likely thing to fail: a fresh slide's id
   * resolves once and then stops. Charging it to a question that never touches
   * the slide turns a real answer into `no-scratch-slide`, which is a statement
   * about the probe wearing the clothes of a statement about the host.
   *
   * That is not hypothetical. The 2026-08-08 `a546897` round lost
   * `untrack-available` this way at 43 seconds — AFTER the host had recovered
   * and answered four questions in a row. `untrack` is a `typeof` on a proxy
   * that `getItemOrNullObject` builds without a round trip, so the answer was
   * there for the taking.
   */
  it("answers even when the scratch slide has stopped resolving", async () => {
    installHost([makeSlide("s1")]);
    // Every slide this run adds is dead on arrival, so the liveness check can
    // only fail. Anything that still answers did so without one.
    faults.newSlideResolvesTimes = 0;
    try {
      const sheet = await runHostProbes("fake-dead-scratch", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      expect(answers["untrack-available"], "a question that needs no slide was charged for one").toBe("no");
      // Non-vacuity: the scratch slide really is unusable, so this is not a run
      // where the check simply passed.
      expect(answers["shape-add-fresh-slide-proxy"]).toBe("no-scratch-slide");
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  }, 60_000);
});

describe("the pane's own read of a sheet", () => {
  /**
   * A declared answer must not be announced as news, wherever it is declared.
   *
   * `shape-add-positional-slide-proxy` is a coin — `yes, yes, threw, yes, yes,
   * threw` across six rounds — and the fake says `yes`. The summary counted
   * only `KNOWN_DIVERGENCES`, so every round the coin landed `threw` the pane
   * said "NEW: shape-add-positional-slide-proxy", about the single question
   * this repo documents at greatest length as varying run to run. A gate that
   * cries wolf on a schedule is one people stop reading.
   */
  it("does not call a documented coin flip a new divergence", () => {
    const sheet = {
      kind: "powerchart-host-answers" as const,
      source: "test",
      build: "test",
      requirementSets: [],
      answers: [
        // The coin, landed on the side the fake does not take.
        { id: "shape-add-positional-slide-proxy", question: "?", answer: "threw", ms: 1 },
        // Non-vacuity: something genuinely undeclared, so this is not a sheet
        // in which nothing could have been reported either way.
        { id: "getitemat-past-end", question: "?", answer: "surprise", ms: 1 },
      ],
    };
    const s = summariseHostSheet(sheet);
    expect(s.fresh, "a documented unstable answer was announced as news").toEqual(["getitemat-past-end"]);
    expect(s.known).toContain("shape-add-positional-slide-proxy");
  });
});

/**
 * A question asked ONCE has been sampled, not answered — this project has
 * written that sentence in three places and still had to learn each unstable
 * answer the expensive way, across ten rounds, because one round could only
 * ever produce one sample. A run asks each question up to `PROBE_PASSES` times
 * now, spread across the run, so the sheet can state stability on its own.
 */
describe("a run samples each question more than once", () => {
  it("asks a healthy host every question three times", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const asked = sheet.answers.filter((a) => a.id !== SCRATCH_CLEANUP_ID);
    expect(asked.length).toBeGreaterThan(10);
    for (const a of asked) {
      expect(a.samples?.length, `${a.id} was asked ${a.samples?.length ?? 0} time(s)`).toBe(PROBE_PASSES);
    }
  });

  it("still files ONE row per question, which is what every reader of a sheet needs", () => {
    // The contract gate, `host-diff`, `host-baseline` and `host-history` all key
    // off one row per id and read `answer`. The repeats had to go somewhere that
    // none of them would notice.
    installHost([makeSlide("s1")]);
    return runHostProbes("fake", "test").then((sheet) => {
      const ids = sheet.answers.map((a) => a.id);
      expect(new Set(ids).size, `duplicate rows: ${ids.filter((x, i) => ids.indexOf(x) !== i).join(", ")}`).toBe(
        ids.length,
      );
    });
  });

  it("says a healthy host's answers are stable, and stamps each sample", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const answered = sheet.answers.filter((a) => a.samples?.length && !NOT_ASKED_WORDS.includes(a.answer));
    expect(answered.length).toBeGreaterThan(10);
    for (const a of answered) expect(a.stable, `${a.id} disagreed with itself on a healthy host`).toBe(true);
    // Every sample carries its pass, its time and the regime — recorded always,
    // not only when something goes wrong, because a value written down only on
    // failures cannot be compared against anything.
    for (const s of answered[0].samples!) {
      expect(s.pass).toBeGreaterThan(0);
      expect(typeof s.atMs).toBe("number");
      expect(["healthy", "slide-trouble", "collection-refused", "unknown"]).toContain(s.regime);
      expect(["first-slide", "fresh-slide", "reused-slide", "no-slide"]).toContain(s.scratch);
      // THE ROUTE, on every sample and for the same reason as the three above.
      // `unknown` is in the list because it is a real state — `record` reached
      // without a take — and NOT because a missing stamp may fall back to it:
      // a sample with no `acquired` at all fails `toContain(undefined)` here,
      // which is the point. Round 254 cost an entire analysis to reconstruct
      // this from `trace.entries` after the fact.
      expect(["first-add", "re-acquired", "recovered", "replaced", "no-slide", "unknown"]).toContain(s.acquired);
    }
  });

  it("stamps the route a slide arrived by, not merely that one did", async () => {
    // A STAMP THAT CANNOT DISAGREE WITH ITSELF IS A CONSTANT WEARING A FIELD'S
    // NAME — the same argument the scratch-stamp test below makes. `takeScratch`
    // has SIX call sites and its `sameSlide` flag distinguishes exactly one of
    // them, which is why this is a value per route rather than a boolean.
    //
    // DROPPING `via` FROM A CALL SITE IS A COMPILE ERROR, not a silent gap:
    // `takeScratch`'s options parameter is required and `via` has no default,
    // so the type system catches the mutation this field exists to prevent.
    // What THIS test adds is the half types cannot see — that the value which
    // reaches the archive is a real route and not the initial `unknown`.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test", { passes: 1 });
    const samples = sheet.answers.flatMap((r) => r.samples ?? []);
    expect(samples.length, "the fake host answers enough to judge").toBeGreaterThan(3);

    const onASlide = samples.filter((s) => s.scratch !== "no-slide");
    expect(onASlide.length, "the fake host does hand out a scratch slide").toBeGreaterThan(0);
    expect(
      onASlide.map((s) => s.acquired).filter((a) => a === "unknown"),
      "a sample taken ON a slide always knows which route produced it",
    ).toEqual([]);
    // And a sample taken on NO slide says so through this field too, rather
    // than inheriting the route of the last slide the run happened to hold.
    for (const s of samples.filter((x) => x.scratch === "no-slide")) {
      expect(s.acquired, "no slide means no route, not the previous route").toBe("no-slide");
    }
    // The two fields answer DIFFERENT questions, which is the whole reason this
    // one exists. If they ever moved together they would be one fact twice.
    expect([...new Set(onASlide.map((s) => s.scratch))].length, "the state varies across a run").toBeGreaterThan(1);
  });

  /**
   * The scratch stamp has to VARY, or it is a constant wearing a field's name.
   *
   * Round 17 eliminated `regime` as the state behind the held-slide-proxy flip:
   * the question agrees with its partner at every instant and still answers two
   * ways inside one regime. `scratch` is the candidate that replaced it, and a
   * candidate that reads the same word on every sample can never explain
   * anything — it would make `explainBy` answer `untested` forever, which looks
   * like caution and is actually a dead field.
   *
   * This is the check the four earlier failure-only fields needed and did not
   * have. It asserts the stamp moves across a whole run on the FAKE, where the
   * host is healthy and the probe still burns and replaces slides — so it is
   * about the bookkeeping, not about a degraded host.
   */
  /**
   * Every acquisition goes through `takeScratch` — held by a SOURCE SCAN,
   * because the fake cannot tell the five sites apart.
   *
   * Five places take a scratch slide: the first one, the recovery after a
   * question burned the last, the replacement after a never-asked, the
   * replacement for a partner question, and the second pass at the end of the
   * run. Bypassing any ONE of them leaves the generation counter and the `used`
   * flag stale, so the next question is stamped `reused-slide` for a slide that
   * is in fact brand new — the exact distinction the field was added to make.
   *
   * A behavioural test cannot hold this: mutating the recovery site to assign
   * `scratchId` directly leaves the whole suite green, because `fresh-slide`
   * still arrives from the replacement path and the run-level assertion is
   * satisfied. Proven by doing it. So the invariant is stated over the source,
   * the way `web-host.test.ts` pins its collection loads and
   * `office-render.test.ts` pins the in-place property list.
   */
  it("routes every scratch-slide acquisition through takeScratch", () => {
    // Repo-relative, the way every other source scan here reads its file:
    // `import.meta.url` is not a file: URL under vitest.
    const src = readFileSync("src/render/host-probe.ts", "utf8");
    const assignments = [...src.matchAll(/scratchId = (.+?);/g)].map((m) => m[1].trim());
    expect(assignments.length, "no assignments found — the scan matched nothing").toBeGreaterThan(3);
    for (const rhs of assignments) {
      // `= null` DROPS the slide and must not touch the counter; everything that
      // takes one has to go through the bookkeeping.
      if (rhs === "null") continue;
      expect(rhs, `an acquisition bypasses takeScratch: scratchId = ${rhs}`).toMatch(/^takeScratch\(/);
    }
  });

  it("stamps more than one scratch state across a run", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const stamps = new Set(sheet.answers.flatMap((a) => (a.samples ?? []).map((s) => s.scratch)));
    // All three of the slide-bearing states, not merely "more than one". A run
    // that only ever said `first-slide` then `reused-slide` would satisfy a
    // size-greater-than-one check and would mean `takeScratch` had stopped
    // resetting `used` — the stamp would be "has any question run yet", not
    // "which slide is this". Measured against the fake before being asserted:
    // 1 x first-slide, 3 x fresh-slide, 83 x reused-slide.
    expect([...stamps].sort()).toEqual(["first-slide", "fresh-slide", "reused-slide"]);
    // `first-slide` is the one the hypothesis turns on: pass 1 meets a deck with
    // no scratch history. If the run never records it, the question that made
    // this field exist cannot be asked of the data.
    expect([...stamps]).toContain("first-slide");
  });

  /**
   * One round file, one clock.
   *
   * A sample's time and a trace line's time are the two series every analysis
   * of these rounds joins, and they were on different origins: the trace starts
   * at `setTracing(true)`, the probe's own stamp started when `runHostProbes`
   * was called — 7.9 seconds later in the round of 2026-08-11. Nothing in the
   * file said so. The trace put pass 2 at 41.6s and the samples put its first
   * answer at 34.4s, which reads as a sample arriving before the pass that
   * produced it, and that is exactly how far a reader gets before drawing a
   * wrong conclusion about which host state an answer came from.
   *
   * Asserted as the property — a sample is timed on the same clock as the trace
   * lines it sits between — rather than against the offset, which is whatever
   * the caller did before the probe.
   */
  it("stamps samples on the same clock the run log uses", async () => {
    installHost([makeSlide("s1")]);
    setTracing(true);
    // Burn a little time on the trace clock BEFORE the probe starts, which is
    // what a real round does (the battery traces its start, then probes).
    trace("selftest", "round starting");
    const spin = Date.now();
    while (Date.now() - spin < 25) {
      /* make the two origins provably different */
    }
    const sheet = await runHostProbes("fake", "test");
    const entries = traceLog().entries.filter((e) => e.scope === "probe");
    expect(entries.length, "the probe traced nothing, so there is no clock to compare against").toBeGreaterThan(0);
    const samples = sheet.answers.flatMap((a) => a.samples ?? []);
    expect(samples.length).toBeGreaterThan(0);
    // Every sample must sit inside the window the probe's own trace lines span.
    // On the old origin the earliest samples read as ~0ms while the probe's
    // first trace line was already tens of ms in, so they fell outside it.
    const first = Math.min(...entries.map((e) => e.ms));
    const last = Math.max(...entries.map((e) => e.ms));
    const early = samples.filter((s) => s.atMs < first);
    expect(
      early.length,
      `${early.length} sample(s) are stamped before the probe's first trace line — the two are on different clocks`,
    ).toBe(0);
    expect(Math.min(...samples.map((s) => s.atMs))).toBeLessThanOrEqual(last);
  });
});

describe("supportOf — how many attempts reached the question", () => {
  const at = (answer: string) => ({ answer, pass: 1, atMs: 0, regime: "healthy", scratch: "fresh-slide" }) as never;

  it("counts real answers apart from attempts that never got to ask", () => {
    expect(supportOf([at("no-scratch-slide"), at("no-scratch-slide"), at("reports-gone")])).toEqual({
      asked: 1,
      of: 3,
    });
  });

  it("says an answer every attempt reached is backed by every attempt", () => {
    expect(supportOf([at("yes"), at("yes")])).toEqual({ asked: 2, of: 2 });
  });

  it("is the half `stabilityOf` cannot report", () => {
    // `stable` needs TWO real samples to say anything, so it is `undefined` for
    // "answered once, never repeated" AND for "never answered at all" — the two
    // cases a reader most needs told apart. One of them is a fact about the
    // host; the other is the absence of one.
    const answeredOnce = [at("no-scratch-slide"), at("threw")];
    const neverAnswered = [at("no-scratch-slide"), at("no-scratch-slide")];
    expect(stabilityOf(answeredOnce)).toBeUndefined();
    expect(stabilityOf(neverAnswered)).toBeUndefined();
    expect(supportOf(answeredOnce)).toEqual({ asked: 1, of: 2 });
    expect(supportOf(neverAnswered)).toEqual({ asked: 0, of: 2 });
  });
});

describe("thinSupport — what gets recorded on the row", () => {
  const at = (answer: string) => ({ answer, pass: 1, atMs: 0, regime: "healthy", scratch: "fresh-slide" }) as never;

  it("records the shortfall when the answer rests on fewer attempts than were made", () => {
    expect(thinSupport([at("no-scratch-slide"), at("threw")])).toEqual({ asked: 1, of: 2 });
  });

  it("says NOTHING when every attempt reached the question", () => {
    // Stamping support on every row puts the interesting case back where it
    // started: in a field nobody reads, because it is always there.
    expect(thinSupport([at("yes"), at("yes")])).toBeUndefined();
  });

  it("says nothing when NO attempt reached it — `answer` already says so", () => {
    expect(thinSupport([at("no-scratch-slide"), at("no-scratch-slide")])).toBeUndefined();
  });

  it("says nothing about a question asked only once", () => {
    expect(thinSupport([at("yes")])).toBeUndefined();
  });
});

describe("a clean-up that can count slides returned before it ran", () => {
  // Nothing hands a slide back mid-run yet. These prove the accounting for the
  // case that exists to allow it, BEFORE anything depends on it — which is the
  // opposite order from the three attempts that were reverted, where the return
  // was written first and the accounting was reconciled at the call sites
  // afterwards, three times, wrongly.
  it("owes only what is still outstanding", () => {
    expect(outstandingScratch(["a", "b", "c"], new Set(["b"]))).toEqual(["a", "c"]);
  });

  it("owes everything when nothing came back early — today's every path", () => {
    expect(outstandingScratch(["a", "b"], new Set())).toEqual(["a", "b"]);
  });

  it("owes nothing when the run gave them all back itself", () => {
    expect(outstandingScratch(["a", "b"], new Set(["a", "b"]))).toEqual([]);
  });

  it("composes with the sweep plan so an early return does not double-count", () => {
    // The failure that killed the pass-boundary attempt, as arithmetic. A run
    // takes 5, hands 3 back mid-run, and the clean-up starts with a deck of
    // 1 + 2. Counting all 5 asks the sweep for slides that are already gone;
    // counting the outstanding 2 asks for exactly what is there.
    const taken = ["a", "b", "c", "d", "e"];
    const early = new Set(["a", "b", "c"]);
    const outstanding = outstandingScratch(taken, early);
    expect(outstanding).toHaveLength(2);
    const plan = positionalSweepPlan({
      deckAtStart: 1,
      deckNow: 3,
      added: outstanding.length,
      alreadyDeleted: 0,
    });
    expect(plan, "the sweep should take exactly the two still in the deck").toEqual({ from: 1, count: 2 });
    const wrong = positionalSweepPlan({ deckAtStart: 1, deckNow: 3, added: taken.length, alreadyDeleted: 0 });
    expect(wrong, "counting all five caps at the growth and reaches for a slide that left").toEqual({
      from: 1,
      count: 2,
    });
  });

  it("reports a shrink measured from the clean-up's own start, not the run's", () => {
    // `deckBefore` is read when the clean-up begins, so slides returned earlier
    // are already out of it. Feeding them back in as `claimed` reports slides
    // the run never failed to return.
    const outstanding = outstandingScratch(["a", "b", "c"], new Set(["a"]));
    expect(slidesActuallyReturned({ claimed: 2, added: outstanding.length, deckBefore: 3, deckAfter: 1 })).toEqual({
      actually: 2,
      left: 0,
      shrankBy: 2,
    });
  });
});

describe("stabilityOf", () => {
  const at = (answer: string) => ({
    answer,
    pass: 1,
    atMs: 0,
    regime: "healthy" as const,
    scratch: "first-slide" as const,
    acquired: "first-add" as const,
  });

  it("needs two REAL samples before it will say anything", () => {
    expect(stabilityOf(undefined)).toBeUndefined();
    expect(stabilityOf([at("yes")])).toBeUndefined();
    // Put once and refused twice is not instability, it is a run that could not
    // ask — calling it unstable would manufacture the noise this exists to find.
    expect(stabilityOf([at("yes"), at("no-scratch-slide"), at("not-asked")])).toBeUndefined();
  });

  it("reports agreement and disagreement", () => {
    expect(stabilityOf([at("yes"), at("yes"), at("yes")])).toBe(true);
    expect(stabilityOf([at("yes"), at("threw")])).toBe(false);
    // A never-asked between two agreeing answers does not make them disagree.
    expect(stabilityOf([at("yes"), at("no-scratch-slide"), at("yes")])).toBe(true);
  });

  it("does not let a failed READ vote against the answers that succeeded", () => {
    /**
     * `rotation-keeps-the-unrotated-box`, round 339, exactly as archived: pass 1
     * could not read, passes 2 and 3 both said `unrotated-box`. The sheet called
     * that UNSTABLE — the host contradicting itself — when the host had said one
     * thing every time it said anything at all.
     *
     * Across the 12,346 archived rows that mislabelling hit 52. A pass that
     * could not read is not a dissenting vote.
     */
    expect(stabilityOf([at("unreadable"), at("unrotated-box"), at("unrotated-box")])).toBe(true);
    expect(stabilityOf([at("unreadable"), at("yes"), at("no")])).toBe(false);
    // `other` is the same kind of non-answer and gets the same treatment.
    expect(stabilityOf([at("other"), at("overwrites"), at("overwrites")])).toBe(true);
  });

  it("tells a host that changed its mind from a run that fell over", async () => {
    /**
     * `stable: false` reads as "the host changed its answer mid-round" and
     * drives both that line and whether a sheet is worth sending at all. On the
     * three collection questions it is mostly reporting OUR decay.
     *
     * Every sample carries the regime it was taken in. Of the archived rounds
     * where these report `stable: false`, the share whose every dissenting
     * sample came under `collection-refused`:
     *
     *     which-end-a-short-read-drops        112 of 127   (88%)
     *     shapes-items-count-honest           114 of 118   (97%)
     *     shapes-items-via-positional-slide   118 of 122   (97%)
     *
     * Pass 1 answers while the host is healthy; passes 2 and 3 answer once it
     * has started refusing collections, and answer differently. That is the host
     * in a state we already named, not the host being unpredictable.
     */
    const { dissentIsDegradation } = await import("../src/render/host-probe");
    const s = (answer: string, regime: string) => ({ ...at(answer), regime: regime as "healthy" });

    expect(
      dissentIsDegradation([s("all", "healthy"), s("none", "collection-refused"), s("none", "collection-refused")]),
      "the archived shape of these three questions was not recognised",
    ).toBe(true);

    // A HOST THAT REALLY DID CHANGE ITS MIND still counts. One dissenting sample
    // taken while healthy is enough to make it the host's business, whatever the
    // others were doing.
    expect(
      dissentIsDegradation([s("all", "healthy"), s("none", "healthy"), s("none", "collection-refused")]),
      "a disagreement from a HEALTHY pass was excused as degradation",
    ).toBe(false);
    // Agreement is not dissent, however degraded the run was.
    expect(dissentIsDegradation([s("all", "healthy"), s("all", "collection-refused")])).toBe(false);
    // And one sample cannot disagree with anything.
    expect(dissentIsDegradation([s("all", "collection-refused")])).toBe(false);
  });

  it("still reports a host that is CONSISTENTLY unreadable as consistent", () => {
    /**
     * The half that must not move, and the reason this rule has two tiers rather
     * than one. Filtering every weak answer outright changes 1,765 archived
     * rows, nearly all `true` → `undefined`: three passes that all said
     * `unreadable` are a host reliably refusing, which is a real and useful
     * fact. Only with two or more REAL answers to compare does the strict
     * reading take over.
     */
    expect(stabilityOf([at("unreadable"), at("unreadable"), at("unreadable")])).toBe(true);
    expect(stabilityOf([at("unreadable"), at("other")])).toBe(false);
    // One real answer is not two, so the fallback still governs here.
    expect(stabilityOf([at("unreadable"), at("unreadable"), at("yes")])).toBe(false);
  });
});

/**
 * The shortlist the later passes fall back to under slide pressure is marked on
 * the questions themselves, and the two tables that define it live in
 * `scripts/host-baseline.mjs`. Pinned to each other here, because a list in one
 * file and a set of marks in another drift the moment either changes — which is
 * the failure mode this repo's own memory warns about by name.
 */
/**
 * The partner that decides whether `shape-add-held-slide-proxy` is a host state
 * or a coin — and the bug in the mechanism it is built on.
 *
 * `record` returns the accumulated ROW, whose `answer` is deliberately the
 * FIRST real answer so a sheet means today what it meant yesterday. The pass
 * loop read that row for three things that all mean "this pass": the `answered`
 * trace line, the follow-up's trigger condition, and the `asked because …`
 * detail. The 2026-08-11 round is the proof — its samples say `threw, yes, yes`
 * while its run log says `answered: threw` three times with an identical
 * `ms: 637` on asks forty-five seconds apart.
 *
 * For the log that is a line asserting what it does not know, the same family
 * as `batch committed`. For the follow-up it is worse: a partner conditioned on
 * a later pass's answer could never fire on it, so the mechanism built to
 * settle a question in ONE round would have needed the answer it was looking
 * for to arrive on pass 1.
 */
describe("the partner question for the held-slide-proxy flip", () => {
  it("asks the identical question again and records it as its own row", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const trigger = sheet.answers.find((a) => a.id === "shape-add-held-slide-proxy");
    const partner = sheet.answers.find((a) => a.id === "shape-add-held-slide-proxy-again");
    expect(trigger, "the trigger question is gone").toBeTruthy();
    expect(partner, "the partner never fired, so the pair cannot decide anything").toBeTruthy();
    // Same vocabulary as its trigger, or the two rows cannot be compared — which
    // is the entire question.
    expect([trigger!.answer, partner!.answer].every((a) => a === "yes" || a === "threw")).toBe(true);
    // The sheet says they are a pair, so a reader a week later sees one finding
    // rather than two unrelated rows.
    expect(partner!.detail).toContain("shape-add-held-slide-proxy");
    expect(partner!.detail).toContain(`answered "${trigger!.answer}"`);
  });

  it("fires on whichever answer THIS pass produced, not the row's first", async () => {
    // Against a host that FLIPS, which is the only host this matters on and the
    // one no fake modelled until `faults.heldSlideProxyRelentsAfter`. Tuned to
    // reproduce the real round exactly: threw, then yes, then yes.
    installHost([makeSlide("s1")]);
    setTracing(true);
    faults.heldSlideProxyRelentsAfter = 2;
    try {
      const sheet = await runHostProbes("fake-relents", "test");
      const trigger = sheet.answers.find((a) => a.id === "shape-add-held-slide-proxy")!;
      const samples = (trigger.samples ?? []).map((x) => x.answer);
      expect(samples, "the fake did not flip, so nothing about per-pass answers is tested").toEqual([
        "threw",
        "yes",
        "yes",
      ]);
      // The row keeps the FIRST real answer — that contract is unchanged.
      expect(trigger.answer).toBe("threw");
      // And the partner is told what the host said THAT pass, not what the row
      // remembers. On the unfixed code every firing reports "threw", so the
      // mechanism built to catch a mid-run flip would be blind to the flip.
      const was = traceLog()
        .entries.filter(
          (e) => e.message === "asking the partner question" && e.data?.after === "shape-add-held-slide-proxy",
        )
        .map((e) => e.data?.was);
      expect(was, "the partner was told the row's answer instead of this pass's").toEqual(["threw", "yes", "yes"]);
    } finally {
      faults.heldSlideProxyRelentsAfter = null;
    }
  });

  it("records the per-pass answer in the run log", async () => {
    // The bug, stated as the property. A follow-up exists to be asked in the
    // light of a particular answer; gating it on the row meant gating it on
    // whatever pass 1 happened to say, so a flip on pass 2 — the exact event
    // this pair was built for — could never trigger it.
    const seen: { was: unknown; pass: number }[] = [];
    installHost([makeSlide("s1")]);
    setTracing(true);
    await runHostProbes("fake", "test");
    for (const e of traceLog().entries) {
      // Scoped to the pair this test is about. Asserting over EVERY partner in
      // the run made it a gate on the whole probe list: adding
      // `untrack-available-on-shape`, which legitimately fires on `no`, broke a
      // test that has nothing to do with it.
      if (e.message === "asking the partner question" && e.data?.after === "shape-add-held-slide-proxy")
        seen.push({ was: e.data?.was, pass: 0 });
      if (e.message === "answered" && e.data?.id === "shape-add-held-slide-proxy") {
        // The per-pass answer is on the line now, so a reader can see the flip.
        expect(e.data, "the answered line does not say which pass it is reporting").toHaveProperty("pass");
      }
    }
    expect(seen.length, "the partner was never asked, so nothing about triggering is tested").toBeGreaterThan(0);
    // Every firing names the answer that caused it, and it is a real one.
    for (const s of seen) expect(["yes", "threw"]).toContain(s.was);
  });
});

describe("the resample shortlist matches the tables that define it", () => {
  it("marks every question this project does not yet trust", () => {
    // A follow-up is exempt from the mark, and not as a loophole: it is never
    // scheduled on its own — it rides its trigger — so marking it would put an
    // id in the scarce-slide shortlist that the shortlist can never ask. The
    // way to keep a partner asked under pressure is to mark its TRIGGER. The
    // codebase already said this in prose on `shape-add-held-slide-proxy-again`
    // and the invariant did not know it, so `untrack-available-on-shape` could
    // not be added to `PENDING_QUESTIONS` without either failing this test or
    // carrying a mark that does nothing.
    const rides = new Set(FOLLOW_UP_IDS);
    const shouldResample = new Set(
      [...Object.keys(UNSTABLE_ANSWERS), ...Object.keys(PENDING_QUESTIONS)].filter((id) => !rides.has(id)),
    );
    const marked = new Set(RESAMPLE_IDS);
    for (const id of shouldResample) {
      expect(marked.has(id), `${id} is in UNSTABLE_ANSWERS/PENDING_QUESTIONS but is not marked resample`).toBe(true);
    }
    for (const id of marked) {
      expect(shouldResample.has(id), `${id} is marked resample but neither table asks for it`).toBe(true);
    }
    expect(marked.size, "the shortlist is empty, so the fallback asks nothing").toBeGreaterThan(0);
  });
});

/**
 * Scratch slides are the scarcest thing on this host, and three passes spend
 * three times as many. When a pass is already losing a third of its questions
 * to `no-scratch-slide`, asking the settled ones twice more bids for slides
 * against the questions that actually need a second sample.
 */
describe("under slide pressure the later passes ask only the shortlist", () => {
  it("keeps re-asking the untrusted questions and stops re-asking the settled ones", async () => {
    installHost([makeSlide("s1")]);
    // Wide enough that pass 1 loses well over a third of its questions, and
    // measured rather than picked: the assertion below reads the ratio the run
    // actually saw, so a fake that stops refusing early fails loudly instead of
    // passing for the wrong reason.
    // DERIVED, because 40 was silently a function of how many probes existed.
    // The fault refuses the FIRST N new-slide requests, so a hardcoded N stops
    // spanning both passes the moment a probe is added: three new questions in
    // 2026-08 pushed pass 1 past 40 on its own, the fault was spent before pass
    // 2 began, and every settled question got a second slide — the exact
    // opposite of what this test asserts, from a change that had nothing to do
    // with scheduling.
    //
    // Scaled to the question count so the pressure lasts into pass 2 whatever
    // the sheet holds. The precondition below still measures what the run
    // actually saw rather than trusting this number.
    faults.newSlideRefusedForFirst = Math.round(PROBE_IDS.length * 1.5);
    try {
      const sheet = await runHostProbes("fake-under-slide-pressure", "test");
      const asked = sheet.answers.filter((a) => a.id !== SCRATCH_CLEANUP_ID);
      const passesFor = (id: string) =>
        new Set((asked.find((a) => a.id === id)?.samples ?? []).map((s) => s.pass).filter((p) => p <= PROBE_PASSES))
          .size;
      const shortlisted = RESAMPLE_IDS.filter((id) => asked.some((a) => a.id === id));
      const settled = asked.filter((a) => !RESAMPLE_IDS.includes(a.id)).map((a) => a.id);
      // The precondition, asserted rather than assumed.
      const firstPass = asked.map((a) => a.samples?.find((x) => x.pass === 1)?.answer).filter(Boolean) as string[];
      const lostFirst = firstPass.filter((x) => NOT_ASKED_WORDS.includes(x)).length;
      expect(lostFirst / firstPass.length, "pass 1 was not under pressure, so this proves nothing").toBeGreaterThan(
        1 / 3,
      );
      expect(shortlisted.length, "no shortlisted question was asked at all").toBeGreaterThan(0);
      // The shortlist still gets its repeats…
      expect(Math.max(...shortlisted.map(passesFor))).toBeGreaterThan(1);
      // …and the settled questions stop after the first pass, which is the
      // whole point: the slides go to the questions that need them.
      expect(Math.max(...settled.map(passesFor)), "settled questions kept spending slides").toBe(1);
    } finally {
      faults.newSlideRefusedForFirst = 0;
    }
  });

  it("gives every question its repeats when there is no pressure", async () => {
    // The negative control: the fallback must be a response to pressure, not
    // the normal path wearing a justification.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const settled = sheet.answers.filter((a) => a.id !== SCRATCH_CLEANUP_ID && !RESAMPLE_IDS.includes(a.id));
    expect(settled.length).toBeGreaterThan(5);
    for (const a of settled) expect(a.samples?.length, `${a.id}`).toBe(PROBE_PASSES);
  });
});

/**
 * A sheet that knows it disagreed with itself has to say so, or the finding has
 * only moved somewhere nobody looks.
 */
describe("the summary reports a question that changed its answer mid-round", () => {
  const sheetWith = (rows: { id: string; answer: string; stable?: boolean }[]) =>
    ({
      kind: "powerchart-host-answers",
      source: "fake",
      build: "test",
      requirementSets: [],
      answers: rows.map((r) => ({ ...r, question: r.id, ms: 1 })),
    }) as HostAnswerSheet;

  it("names it, and calls the round worth sending", () => {
    const sheet = sheetWith([{ id: "shape-add-held-slide-proxy", answer: "threw", stable: false }]);
    expect(summariseHostSheet(sheet).unstable).toEqual(["shape-add-held-slide-proxy"]);
    expect(describeHostSheet(sheet)).toContain("CHANGED ITS ANSWER MID-ROUND");
    expect(sheetNeedsAttention(sheet)).toBe(true);
  });

  it("stays quiet when every question agreed with itself", () => {
    // The negative control: a round of stable answers must still be able to say
    // "no need to send it".
    const sheet = sheetWith([{ id: "getitem-durable-slide", answer: "yes", stable: true }]);
    expect(summariseHostSheet(sheet).unstable).toEqual([]);
    expect(describeHostSheet(sheet)).not.toContain("CHANGED ITS ANSWER");
  });
});

/**
 * The first version of this was two sticky booleans, and the first real round
 * showed what that is worth: the flag latched 8.9 seconds into a 110-second
 * probe and 55 of 65 samples came back carrying the same value. A field nearly
 * every sample shares cannot separate anything — the same mistake this project
 * has recorded against `idleMs`, `afterAnswering`, the settle's shared label and
 * `listChartsInDeck`.
 */
describe("regimeFrom describes the host NOW, not the host at any point", () => {
  /**
   * A refusal used to win outright inside the window, so ONE refusal painted
   * every sample `collection-refused` for the next twenty seconds however well
   * the collection answered in between — priority order, where the docstring
   * promised "the most RECENT thing this run watched it do".
   *
   * Round 17 measured the cost: the collection answered 14 of the 28 times it
   * was asked, interleaved with the refusals, and 88% of that round's samples
   * still read `collection-refused` — worse than the 85% sticky flag this
   * function replaced, and past the failure criterion its own docstring sets
   * ("a field that nearly every sample shares cannot separate anything").
   * Simulated over the same round, this rule takes it to 72%.
   */
  it("lets a later COLLECTION answer clear an earlier refusal", () => {
    // Refused at 900, answered at 950, asked at 1000: the collection is talking.
    expect(regimeFrom({ at: 1000, lastRefusalAt: 900, lastCollectionGoodAt: 950, lastGoodAt: 950 })).not.toBe(
      "collection-refused",
    );
    // The other order still reports the refusal — this clears a refusal, it does
    // not suppress one.
    expect(regimeFrom({ at: 1000, lastRefusalAt: 950, lastCollectionGoodAt: 900, lastGoodAt: 950 })).toBe(
      "collection-refused",
    );
  });

  it("is not cleared by an ordinary question answering", () => {
    // `lastGoodAt` is set by almost every question, so recency across ALL three
    // signals would saturate on `healthy` instead — the same defect mirrored.
    // Only a later COLLECTION answer counts.
    expect(regimeFrom({ at: 1000, lastRefusalAt: 900, lastGoodAt: 990 })).toBe("collection-refused");
  });

  it("reports the most recent thing seen", () => {
    expect(regimeFrom({ at: 1000, lastRefusalAt: 900 })).toBe("collection-refused");
    expect(regimeFrom({ at: 1000, lastSlideTroubleAt: 900 })).toBe("slide-trouble");
    expect(regimeFrom({ at: 1000, lastGoodAt: 900 })).toBe("healthy");
  });

  it("lets a refusal go stale instead of latching for the rest of the run", () => {
    // The bug, stated as a test: a refusal a minute old is history.
    const stale = { at: 90_000, lastRefusalAt: 8_900, lastGoodAt: 89_000 };
    expect(regimeFrom(stale)).toBe("healthy");
    // …and while it IS recent it still outranks a good read, because a host
    // that will not list shapes is in the deeper hole.
    expect(regimeFrom({ at: 10_000, lastRefusalAt: 9_000, lastGoodAt: 9_500 })).toBe("collection-refused");
  });

  it("says unknown rather than guessing when nothing recent is known", () => {
    expect(regimeFrom({ at: 90_000 })).toBe("unknown");
    expect(regimeFrom({ at: 90_000, lastRefusalAt: 100, lastSlideTroubleAt: 200, lastGoodAt: 300 })).toBe("unknown");
  });

  it("takes the window as a parameter, so the boundary is testable", () => {
    expect(regimeFrom({ at: 100, lastRefusalAt: 0 }, 100)).toBe("collection-refused");
    expect(regimeFrom({ at: 101, lastRefusalAt: 0 }, 100)).toBe("unknown");
  });
});

/**
 * A clean-up that says it happened, into a deck that says otherwise.
 *
 * The 2026-08-11 round reported `all — 42 of 42 scratch slide(s) deleted,
 * left: 0` and ended holding 56 slides it had added, 47 of them blank; the
 * owner's screenshot showed one carrying the 20x20 rectangle
 * `shape-add-held-slide-proxy` draws. `deleteSlideByPosition` reads
 * `indexOf(id) < 0` as "already gone", which is sound only while the id we hold
 * and the ids the deck lists are the same strings — and in that round the
 * scratch ids read `4123571115#123571113` while the deck listed
 * `257#2599158489`. An id nobody can find is UNKNOWN, not absent.
 *
 * This is the second time this project has shipped a clean-up that reported
 * work it had not done (`deleteSlideById` itself was the first), and both times
 * the deck was the thing that knew.
 */
describe("what the scratch clean-up may claim", () => {
  it("believes the deck over the deletes when they disagree", () => {
    // The real round, in numbers: every delete said yes, the deck did not move.
    const r = slidesActuallyReturned({ claimed: 42, added: 42, deckBefore: 57, deckAfter: 57 });
    expect(r.actually, "a clean-up nothing corroborates was reported as complete").toBe(0);
    expect(r.left).toBe(42);
    expect(r.shrankBy).toBe(0);
  });

  it("grades a clean handback ALL even when the count runs past its own denominator", async () => {
    // 25 OF 140 ARCHIVED ROUNDS, and 25 of the 30 `some` verdicts in the whole
    // archive. The positional sweep was licensed with
    // `outstanding.length + unnamedLeftBehind` and the verdict was scored
    // against `outstanding.length` — two denominators 24 lines apart — so the
    // numerator counted slides the denominator did not know existed and `left`
    // came out NEGATIVE.
    //
    // `!left` is false for -4, so a clean handback was graded `some`. Round 088,
    // verbatim: "98 of 94 scratch slide(s) deleted; -4 never landed — the host
    // took the add and the deck never listed them" — 104% of its own
    // denominator, reported as a shortfall, beside a deck that came back to
    // exactly the size it started at.
    const { scratchCleanupAnswer, scratchLeftSentence } = await import("../src/render/host-probe");
    expect(scratchCleanupAnswer({ addedAny: true, left: -4, actually: 98 })).toBe("all");
    expect(scratchCleanupAnswer({ addedAny: true, left: 0, actually: 42 })).toBe("all");
    // AND THE POSITIVE CASES ARE UNTOUCHED — a real shortfall must still grade
    // as one, or this "fix" would be a clamp that hides the defect it was
    // written for.
    expect(scratchCleanupAnswer({ addedAny: true, left: 4, actually: 90 })).toBe("some");
    expect(scratchCleanupAnswer({ addedAny: true, left: 4, actually: 0 })).toBe("none");
    expect(scratchCleanupAnswer({ addedAny: false, left: 0, actually: 0 })).toBe("none-added");

    // A NEGATIVE `left` GETS ITS OWN SENTENCE and never one of the other two.
    // Both of those describe slides that are MISSING, so a negative number in
    // either says the opposite of what happened. Round 073 printed
    // "-1 left in the deck" directly beside "the deck still lists 1".
    const negative = scratchLeftSentence({ left: -4, stillListed: 0, deckBefore: 99, deckAfter: 1 });
    expect(negative).toMatch(/MORE slide\(s\) than this run can account for/);
    expect(negative, "a negative count rode inside a sentence that means the opposite").not.toMatch(/never landed/);
    expect(negative).not.toMatch(/left in the deck/);
    expect(negative, "the magnitude is reported unsigned, so the sentence reads").toMatch(/lost 4 MORE/);

    // The two positive sentences still split on what the DECK says, which is
    // what round 029 bought: 73 slides counted as abandoned while the deck stood
    // at one the whole time and listed none of their ids.
    expect(scratchLeftSentence({ left: 73, stillListed: 0, deckBefore: 1, deckAfter: 1 })).toMatch(/never landed/);
    expect(scratchLeftSentence({ left: 1, stillListed: 1, deckBefore: 1, deckAfter: 2 })).toMatch(/left in the deck/);
    expect(scratchLeftSentence({ left: 0 }), "a complete handback says nothing at all").toBe("");
  });

  it("will not claim more slides are left in the deck than the deck holds", async () => {
    const { scratchLeftSentence } = await import("../src/render/host-probe");
    /**
     * Round 346, exactly as archived: `left=20` beside `deckAfter=1`. Twenty
     * slides cannot be left in a deck that holds one, and the sentence said they
     * were. Across the archive 106 of 422 of these records — 25% — claim more
     * left behind than the deck contains.
     *
     * `left` is `added - actually`, and those are different KINDS of number:
     * `added` accumulates across the whole round, `actually` is bounded by a
     * delta between two readings of the deck. Subtracting one from the other is
     * a category error.
     *
     * NAMED, NOT CLAMPED — this function's own docstring says clamping a number
     * into a sentence that cannot hold it "is exactly how the last one survived
     * 25 rounds". Both figures stay on the page and the contradiction is stated,
     * because that is the part a reader cannot reconstruct from one number.
     */
    const s = scratchLeftSentence({ left: 20, stillListed: 20, deckBefore: 38, deckAfter: 1 });
    expect(s, "claimed 20 slides sit in a 1-slide deck").not.toMatch(/left in the deck/);
    expect(s, "dropped the count instead of naming the contradiction").toContain("20");
    expect(s, "dropped the deck's own figure, which is the half that refutes it").toContain("1 slide(s)");
    expect(s).toMatch(/cannot all be in it/);

    // The boundary: exactly as many left as the deck holds is not a
    // contradiction, and must keep the plain sentence.
    expect(scratchLeftSentence({ left: 2, stillListed: 2, deckBefore: 9, deckAfter: 2 })).toMatch(/left in the deck/);
    // And a round that never learned the deck size cannot be contradicted by it.
    expect(scratchLeftSentence({ left: 20, stillListed: 20 })).toMatch(/left in the deck/);
  });

  it("still reports success when the deck agrees", () => {
    const r = slidesActuallyReturned({ claimed: 42, added: 42, deckBefore: 99, deckAfter: 57 });
    expect(r.actually).toBe(42);
    expect(r.left).toBe(0);
  });

  it("takes the smaller number, never the larger", () => {
    // A deck that shrank MORE than we deleted is not 50 successes — something
    // else removed slides, and claiming them would be claiming the user's work.
    expect(slidesActuallyReturned({ claimed: 10, added: 42, deckBefore: 99, deckAfter: 49 }).actually).toBe(10);
    // And a deck that GREW during clean-up is not negative progress.
    expect(slidesActuallyReturned({ claimed: 10, added: 42, deckBefore: 50, deckAfter: 57 }).actually).toBe(0);
  });

  it("falls back to the claim when the deck will not give a count", () => {
    // Not a licence to believe it — there is simply nothing else, and a host
    // that will not list its slides has already said the interesting thing.
    const r = slidesActuallyReturned({ claimed: 42, added: 42 });
    expect(r.actually).toBe(42);
    expect(r.shrankBy).toBeUndefined();
  });
});

/**
 * Which slides a positional sweep may remove.
 *
 * This is the safety, not the clean-up. Delete-by-id cannot work on this host —
 * 2026-08-11 measured `the deck still lists 0 of 62 of these ids` — and what
 * survives is position, because `slides.add()` appends and a run's own slides
 * are therefore the last N. Position is also how an add-in destroys someone's
 * work, so every case below is a way of asking the same question: can this plan
 * ever reach a slide that was in the deck before the run started?
 */
describe("what a positional sweep is allowed to delete", () => {
  const plan = positionalSweepPlan;

  it("removes exactly the run's own slides, off the end", () => {
    // The real shape: deck started at 3, the run added 62 and got none back.
    expect(plan({ deckAtStart: 3, deckNow: 65, added: 62, alreadyDeleted: 0 })).toEqual({ from: 3, count: 62 });
  });

  it("never reaches below the deck's size when the run started", () => {
    // The property everything else exists to protect. Whatever the arithmetic,
    // `from` is the floor and the user's slides live below it.
    for (const added of [1, 5, 62, 500]) {
      for (const deckNow of [3, 4, 10, 65]) {
        const p = plan({ deckAtStart: 3, deckNow, added, alreadyDeleted: 0 });
        if (p) expect(p.from, `a plan reached into the deck's original ${3} slides`).toBeGreaterThanOrEqual(3);
      }
    }
  });

  it("never removes more than this run added", () => {
    // A deck that grew by more than we added grew for some other reason, and
    // the extra is not ours to take.
    expect(plan({ deckAtStart: 3, deckNow: 90, added: 10, alreadyDeleted: 0 })).toEqual({ from: 80, count: 10 });
  });

  it("never removes more than the deck actually grew", () => {
    // Something else removed slides while the probe ran, so the count we added
    // no longer describes the deck. The smaller number is the honest one.
    expect(plan({ deckAtStart: 3, deckNow: 8, added: 62, alreadyDeleted: 0 })).toEqual({ from: 3, count: 5 });
  });

  it("subtracts whatever delete-by-id already got back", () => {
    expect(plan({ deckAtStart: 3, deckNow: 65, added: 62, alreadyDeleted: 60 })).toEqual({ from: 63, count: 2 });
  });

  it("refuses when the deck did not grow, or shrank", () => {
    expect(plan({ deckAtStart: 10, deckNow: 10, added: 5, alreadyDeleted: 0 })).toBeNull();
    expect(plan({ deckAtStart: 10, deckNow: 4, added: 5, alreadyDeleted: 0 })).toBeNull();
  });

  it("refuses when nothing is left to take back", () => {
    expect(plan({ deckAtStart: 3, deckNow: 65, added: 62, alreadyDeleted: 62 })).toBeNull();
  });

  it("refuses when the deck will not say how big it is or was", () => {
    // A host that will not count its slides does not get to have slides
    // deleted from it by arithmetic.
    expect(plan({ deckAtStart: undefined, deckNow: 65, added: 62, alreadyDeleted: 0 })).toBeNull();
    expect(plan({ deckAtStart: 3, deckNow: undefined, added: 62, alreadyDeleted: 0 })).toBeNull();
    expect(plan({ added: 62, alreadyDeleted: 0 })).toBeNull();
  });
});

/**
 * The clean-up's own sentence, in both directions.
 *
 * The shortfall clause was written when delete-by-id was the only mechanism, so
 * any disagreement between the claim and the deck meant an over-claim. The
 * positional sweep made the deck lose MORE than the deletes claimed, and the
 * line then read `the deletes reported 0 but the deck only shrank by 68` — the
 * sweep working, described as a shortfall. A line that says the opposite of
 * what happened is the `batch committed` mistake in a new place.
 */
describe("what the clean-up's report may say about a disagreement", () => {
  it("is a shortfall only when the deck lost LESS than the deletes claimed", () => {
    // The real round: by-id took none, the sweep took 68, the deck lost 68.
    expect(slidesActuallyReturned({ claimed: 68, added: 68, deckBefore: 71, deckAfter: 3 })).toEqual({
      actually: 68,
      left: 0,
      shrankBy: 68,
    });
    // and the other direction, which IS a shortfall and must still be caught.
    expect(slidesActuallyReturned({ claimed: 45, added: 45, deckBefore: 57, deckAfter: 57 })).toEqual({
      actually: 0,
      left: 45,
      shrankBy: 0,
    });
  });
});

/**
 * A probe that could not set itself up must say so in a word the gates know.
 *
 * `NEVER_ASKED` is the vocabulary meaning "the run never put this question",
 * and both readers of a sheet depend on it: `compareSheets` counts anything
 * else as a fact about the host, and `history` starts a streak on it. A probe
 * that invents its own word for a setup failure turns "we could not ask" into
 * "the host said so" — the mistake this repo already paid for when
 * `no-scratch-slide` counted as agreement.
 *
 * It reached a real round. `grouped-child-by-id-from-slide` — the question that
 * decides whether the in-place update has a future on this host — answered
 * `no-child-id` three times out of three on `275a76a`. That is its own early
 * return for "the host would not name the members before grouping", and no gate
 * knows the word: the contract diff would have called it a host divergence and
 * `history` would have reported a three-round streak on it.
 *
 * Asserted through the fake rather than by reading names. `no-binding-api`,
 * `no-group-id` and `no-durable-slide` are genuine ANSWERS about a host that
 * lacks an API or will not name a group, and no naming rule tells those apart
 * from a setup failure — only driving the failure does.
 */
describe("what a probe says when it could not set itself up", () => {
  it("uses the gates' vocabulary when the host will not name a shape", async () => {
    installHost([makeSlide("s1")]);
    // The state the real round was in: shapes are added, and their ids will not
    // read back — so a question that has to re-fetch a shape BY its id was
    // never put. `tag-through-refetched-shape` stands in for
    // `grouped-child-by-id-from-slide`, which this test used until that
    // question was retired; both need an id read back, which is the setup the
    // fault refuses.
    faults.refuseShapeIdLoads = 9999;
    try {
      const sheet = await runHostProbes("fake-refuses-ids", "test");
      const row = sheet.answers.find((a) => a.id === "tag-through-refetched-shape");
      expect(row, "the question vanished from the sheet").toBeTruthy();
      expect(
        NEVER_ASKED.has(String(row!.answer)),
        `answered "${row!.answer}", which no gate reads as "never asked" — a setup failure would count as a fact about the host`,
      ).toBe(true);
    } finally {
      faults.refuseShapeIdLoads = 0;
    }
  }, 30_000);

  it("keeps NEVER_ASKED and the probe's own vocabulary in step", () => {
    const src = readFileSync("src/render/host-probe.ts", "utf8");
    for (const word of NEVER_ASKED)
      expect(src, `the probe never emits "${word}", so the gate's vocabulary has drifted`).toContain(`"${word}"`);
  });

  it("keeps the .mjs tools' second tier in step with UNINFORMATIVE", async () => {
    // The same arrangement as NEVER_ASKED above and for the same reason: the
    // .mjs tools import nothing from the TypeScript, so the words are copied,
    // and the copies are pinned here rather than to anyone's memory. Each of
    // these cost rounds of real answers before it was classified at all.
    expect([...UNINFORMATIVE_ANSWERS].sort(), "the tools and the probe disagree about what is not an answer").toEqual(
      [...UNINFORMATIVE].sort(),
    );
  });
});

/**
 * Counting the slides a single add leaves behind.
 *
 * `addScratchSlide` reports slides that landed and could not be named, and the
 * sweep clamps at "no more than this run added" — so an undercount here is the
 * clamp declining to remove a slide the run put in the deck.
 */
describe("a slide add that lands more than it was asked for", () => {
  afterEach(() => {
    faults.addLandsExtra = 0;
    faults.newSlideResolvesTimes = 0;
  });

  it("reports every slide that landed, not one per add", async () => {
    // TEN EVENTS ACROSS NINE ARCHIVED ROUNDS, and `after - before` is 2 in every
    // one of them — never 1. This branch counted events, so the clamp ran one
    // short each time and round 085 shipped the remainder: its own inventory
    // carries `257#3837665135` with zero shapes, listed in `newSlides`.
    //
    // The fake could only ever land one slide per add, so the branch was tested
    // against a host that never produced the state it exists for.
    installHost([makeSlide("s1")]);
    faults.addLandsExtra = 1; // one add, two slides — what the real host does
    faults.newSlideResolvesTimes = 0; // and the new slide cannot be named
    let unnamed = 0;
    const id = await addScratchSlide(undefined, undefined, () => {
      unnamed++;
    });
    expect(id, "this test needs the un-nameable branch").toBeNull();
    expect(unnamed, "counted the add, not the slides it landed").toBe(2);
  });

  it("reports nothing when the slide it added could be named", async () => {
    // The over-count guard. Trading an undercount for an over-count would have
    // the sweep clamp permit deleting a slide the run did NOT add, which is the
    // user's own deck.
    //
    // Note there is no "one landed and could not be named" case to test: with a
    // single fresh slide the code names it, and the ambiguity that defeats
    // naming is exactly what a second one creates. That is why the archive's
    // ten events are all two-slide events.
    installHost([makeSlide("s1")]);
    let unnamed = 0;
    const id = await addScratchSlide(undefined, undefined, () => {
      unnamed++;
    });
    expect(id, "the ordinary path should name its slide").toBeTruthy();
    expect(unnamed, "counted a slide it had just named").toBe(0);
  });
});

/**
 * Asking again, once the round has something the host will name.
 *
 * A round builds its sheet before the battery draws anything, so a question
 * needing a named id is put at the one moment none exists. These cover the two
 * halves of the repair: narrowing a run to the deferred questions, and folding
 * what comes back into the sheet that deferred them.
 */
describe("re-asking what an empty deck could not answer", () => {
  const sheetWith = (rows: { id: string; answer: string; passes?: number[] }[]): HostAnswerSheet => ({
    kind: "powerchart-host-answers",
    source: "fake",
    build: "test",
    requirementSets: ["1.5"],
    answers: rows.map((r) => ({
      id: r.id,
      question: r.id,
      answer: r.answer,
      ms: 1,
      samples: (r.passes ?? [1, 2, 3]).map((pass) => ({
        answer: r.answer,
        pass,
        atMs: pass,
        regime: "healthy" as const,
        scratch: "first-slide" as const,
        acquired: "first-add" as const,
      })),
    })),
  });

  it("names only the questions lost for want of a shape", async () => {
    const { deferredForLackOfShape } = await import("../src/render/host-probe");
    // `no-scratch-slide` is a DIFFERENT setup failure and a re-ask cannot help
    // it — the deck refused a slide, and drawing a chart does not change that.
    // Pulling it in would spend a re-ask on a question certain to fail again.
    const picked = deferredForLackOfShape(
      sheetWith([
        { id: "a", answer: "no-scratch-shape" },
        { id: "b", answer: "no-scratch-slide" },
        { id: "c", answer: "yes" },
        { id: "d", answer: "no-scratch-shape" },
      ]),
    );
    expect(picked).toEqual(["a", "d"]);
  });

  it("catches a question the slide blocked first and the shape blocked after", async () => {
    const { deferredForLackOfShape } = await import("../src/render/host-probe");
    // ROUND 245, exactly. `answer` holds the FIRST weak answer and nothing weak
    // displaces it, so a question that failed once on the slide and three times
    // on the shape reads `no-scratch-slide` forever. Reading only the row found
    // no match, made no re-ask, and the round produced four samples where every
    // round since 241 produced five.
    const sheet = sheetWith([{ id: "a", answer: "no-scratch-slide" }]);
    sheet.answers[0].samples = ["no-scratch-slide", "no-scratch-shape", "no-scratch-shape"].map((answer, i) => ({
      answer,
      pass: i + 1,
      atMs: i,
      regime: "healthy" as const,
      scratch: "first-slide" as const,
      acquired: "first-add" as const,
    }));
    expect(deferredForLackOfShape(sheet)).toEqual(["a"]);
  });

  it("leaves alone a question that was actually answered", async () => {
    const { deferredForLackOfShape } = await import("../src/render/host-probe");
    // A real answer on a later pass settles it. Re-asking could only add samples
    // to something already decided, and it costs a scratch slide to do it.
    const sheet = sheetWith([{ id: "a", answer: "yes" }]);
    sheet.answers[0].samples = [
      {
        answer: "no-scratch-shape",
        pass: 1,
        atMs: 1,
        regime: "healthy",
        scratch: "first-slide",
        acquired: "first-add",
      },
      { answer: "yes", pass: 2, atMs: 2, regime: "healthy", scratch: "first-slide", acquired: "first-add" },
    ];
    expect(deferredForLackOfShape(sheet)).toEqual([]);
  });

  it("puts only the questions it was given", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test", {
      only: ["shape-resolve-held-slide-proxy"],
      passes: 1,
    });
    // The cleanup row is bookkeeping about the slide the run borrowed, not a
    // question anyone asked — it is appended to every sheet regardless.
    const asked = sheet.answers.filter((r) => r.id !== SCRATCH_CLEANUP_ID);
    expect(asked.map((r) => r.id)).toEqual(["shape-resolve-held-slide-proxy"]);
    expect(asked[0].samples).toHaveLength(1);
  });

  it("asks nothing when the selection names nothing real", async () => {
    installHost([makeSlide("s1")]);
    // A caller asking for two questions and silently getting twenty-odd is the
    // opposite of what it wanted, and on a real host it is minutes of deck churn
    // nobody asked for.
    const sheet = await runHostProbes("fake", "test", { only: ["not-a-question"], passes: 1 });
    expect(sheet.answers.filter((r) => r.id !== SCRATCH_CLEANUP_ID)).toEqual([]);
  });

  it("lets a real answer displace the deferral it replaces", async () => {
    const { mergeHostSheets } = await import("../src/render/host-probe");
    const base = sheetWith([
      { id: "a", answer: "no-scratch-shape" },
      { id: "b", answer: "yes" },
    ]);
    const merged = mergeHostSheets(base, sheetWith([{ id: "a", answer: "unreadable", passes: [1] }]));
    expect(merged.answers.find((r) => r.id === "a")!.answer).toBe("unreadable");
    // The re-ask asked ONE question. Everything else keeps what it said.
    expect(merged.answers.find((r) => r.id === "b")!.answer).toBe("yes");
    expect(merged.answers).toHaveLength(2);
  });

  it("keeps an answer the round already had", async () => {
    const { mergeHostSheets } = await import("../src/render/host-probe");
    // The rule `record` uses inside a run, and for the same reason: the FIRST
    // real answer is the row's answer, so a sheet means today what it meant
    // yesterday. A re-ask is extra evidence, not a correction.
    const merged = mergeHostSheets(
      sheetWith([{ id: "a", answer: "yes" }]),
      sheetWith([{ id: "a", answer: "no", passes: [1] }]),
    );
    expect(merged.answers[0].answer).toBe("yes");
    // …and the disagreement is still visible rather than hidden.
    expect(merged.answers[0].samples!.map((x) => x.answer)).toContain("no");
    expect(merged.answers[0].stable).toBe(false);
  });

  it("numbers the re-ask's passes after the ones already taken", async () => {
    const { mergeHostSheets } = await import("../src/render/host-probe");
    // Both runs number their passes from 1. Left alone the row would read
    // `1,2,3,1` and this file's own per-pass arithmetic — which finds a sample
    // BY pass number — would take the last sample for the first.
    const merged = mergeHostSheets(
      sheetWith([{ id: "a", answer: "no-scratch-shape", passes: [1, 2, 3] }]),
      sheetWith([{ id: "a", answer: "yes", passes: [1] }]),
    );
    expect(merged.answers[0].samples!.map((x) => x.pass)).toEqual([1, 2, 3, 4]);
  });

  it("does not rewrite the sheet the round already banked", async () => {
    const { mergeHostSheets } = await import("../src/render/host-probe");
    // The original is in the crash store precisely so a tab that dies next does
    // not take the probe's answers with it. A merge that mutated it would edit
    // the record whose whole job is to be already written.
    const base = sheetWith([{ id: "a", answer: "no-scratch-shape" }]);
    const before = JSON.stringify(base);
    mergeHostSheets(base, sheetWith([{ id: "a", answer: "yes", passes: [1] }]));
    expect(JSON.stringify(base)).toBe(before);
  });

  it("keeps a row the sheet has never seen", async () => {
    const { mergeHostSheets } = await import("../src/render/host-probe");
    // An answer this host gave. Dropping it because the base had no slot for it
    // would lose it with nothing saying so.
    const merged = mergeHostSheets(
      sheetWith([{ id: "a", answer: "yes" }]),
      sheetWith([{ id: "z", answer: "threw", passes: [1] }]),
    );
    expect(merged.answers.map((r) => r.id)).toEqual(["a", "z"]);
  });
});

describe("deciding whether the round can ask yet", () => {
  const deferredSheet = (): HostAnswerSheet => ({
    kind: "powerchart-host-answers",
    source: "fake",
    build: "test",
    requirementSets: ["1.5"],
    answers: [
      { id: "a", question: "a", answer: "no-scratch-shape", ms: 1 },
      { id: "b", question: "b", answer: "yes", ms: 1 },
    ],
  });

  it("asks once the battery has left a chart this host named", async () => {
    const { reaskPlan } = await import("../src/render/host-probe");
    expect(reaskPlan(deferredSheet(), { slideId: "s1", shapeId: "sh1" })).toEqual({ ask: ["a"], skipped: [] });
  });

  it("asks nothing when nothing got tagged", async () => {
    const { reaskPlan } = await import("../src/render/host-probe");
    // A re-ask without an id is the SAME question that already failed, and it
    // would spend a scratch slide and a minute to fail identically. Worse, it
    // would make `no-scratch-shape` look like a deferral when here it is true.
    expect(reaskPlan(deferredSheet(), null)).toEqual({ ask: [], skipped: ["a"] });
  });

  it("asks nothing when nothing was deferred", async () => {
    const { reaskPlan } = await import("../src/render/host-probe");
    const answered: HostAnswerSheet = {
      ...deferredSheet(),
      answers: [{ id: "b", question: "b", answer: "yes", ms: 1 }],
    };
    expect(reaskPlan(answered, { slideId: "s1", shapeId: "sh1" })).toEqual({ ask: [], skipped: [] });
  });
});

/**
 * Which of two very different facts an `unreadable` was.
 *
 * A null SLIDE handle produces a null shape whose id reads undefined by exactly
 * the same path as the host losing a shape proxy across a slide switch. One is
 * a setup failure of ours; the other is office-js #2903, the thing this probe is
 * named for. Round 241 asked the question for the first time in 217 rounds and
 * these two were still indistinguishable in its answer.
 */
describe("resolving a shape through a tagged chart's own slide", () => {
  afterEach(async () => {
    // A planted id changes what EVERY probe run in this file does, so it must
    // not outlive the test that planted it.
    const { _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
  });

  it("says the slide never resolved rather than blaming the host", async () => {
    const { _rememberNamedShapeForTest } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    _rememberNamedShapeForTest("no-such-slide", "no-such-shape");
    const sheet = await runHostProbes("fake", "test", {
      only: ["shape-resolve-held-slide-proxy"],
      passes: 1,
    });
    const row = sheet.answers.find((r) => r.id === "shape-resolve-held-slide-proxy")!;
    // NOT `unreadable`. Reporting a lost proxy here would put a host bug in the
    // sheet on the strength of our own bad slide id.
    expect(row.answer).toBe("no-named-slide");
    expect(row.detail).toContain("no-such-slide");
  });

  it("says which kind of nothing came back when the shape is missing", async () => {
    const { _rememberNamedShapeForTest } = await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    // A REAL slide, so the slide check passes and the shape is what fails. The
    // two produce the same `unreadable` and mean different things: no such
    // shape says the host never agreed the id belonged there, while a shape
    // that resolves and will not read back is office-js #2903 proper.
    _rememberNamedShapeForTest("s1", "no-such-shape");
    const sheet = await runHostProbes("fake", "test", {
      only: ["shape-resolve-held-slide-proxy"],
      passes: 1,
    });
    const row = sheet.answers.find((r) => r.id === "shape-resolve-held-slide-proxy")!;
    expect(row.answer).toBe("unreadable");
    expect(row.detail).toContain("no such shape on that slide");
    // The verdict itself must NOT split — 217 rounds of comparison rest on it.
    expect(NOT_ASKED.has(row.answer)).toBe(false);
  });

  it("names the stale-proxy failure when the shape resolves and stays silent", async () => {
    const { _rememberNamedShapeForTest } = await import("../src/render/powerpoint");
    const deck = [makeSlide("s1")];
    installHost(deck);
    const real = deck[0].shapes.addGeometricShape("Rectangle", { left: 1, top: 1, width: 9, height: 9 });
    _rememberNamedShapeForTest("s1", real.id);
    faults.shapeIdNeverPopulates = true;
    try {
      const sheet = await runHostProbes("fake", "test", {
        only: ["shape-resolve-held-slide-proxy"],
        passes: 1,
      });
      const row = sheet.answers.find((r) => r.id === "shape-resolve-held-slide-proxy")!;
      expect(row.answer).toBe("unreadable");
      // office-js #2903 proper: the host FOUND it and would not read it back.
      // The other `unreadable` — no such shape — means the host never agreed the
      // id belonged to that slide, which is a different bug with a different fix.
      expect(row.detail).toContain("the shape resolved but would not read back");
      expect(row.detail).not.toContain("no such shape");
      // Here the shape IS on the slide and the listing says so, which is what
      // makes this reading the host's fault rather than a stale id of ours.
      expect(row.detail).toContain("the id IS among the slide's");
    } finally {
      faults.shapeIdNeverPopulates = false;
    }
  });

  it("says whether the slide still holds anything, to tell a stale id from a refusal", async () => {
    const { _rememberNamedShapeForTest } = await import("../src/render/powerpoint");
    const deck = [makeSlide("s1")];
    installHost(deck);
    deck[0].shapes.addGeometricShape("Rectangle", { left: 1, top: 1, width: 9, height: 9 });
    // An id that is NOT on the slide, while the slide plainly holds a shape.
    // Without the count, this reading and "the chart was redrawn away before we
    // asked" produce the identical `no such shape on that slide` — one is a
    // finding about the host, the other is our own id gone stale.
    _rememberNamedShapeForTest("s1", "no-such-shape");
    const sheet = await runHostProbes("fake", "test", {
      only: ["shape-resolve-held-slide-proxy"],
      passes: 1,
    });
    const row = sheet.answers.find((r) => r.id === "shape-resolve-held-slide-proxy")!;
    expect(row.detail).toContain("no such shape on that slide");
    expect(row.detail).toMatch(/slide holds [1-9]\d* shape\(s\)/);
    // …and the id it actually tried, so the pairing can be checked against the
    // draw trace rather than assumed.
    expect(row.detail).toContain("no-such-shape");
    // THE DISCRIMINATOR. An id absent from the slide's own listing is a shape
    // that is gone — ours to fix. An id PRESENT in the listing and still
    // unresolvable by `getItemOrNullObject` is the host refusing a lookup for a
    // shape sitting right there, which is the finding. Without this the two
    // arrive as the same `no such shape on that slide`.
    expect(row.detail).toContain("is NOT among the slide's");
  });

  it("counts a slide that never resolved as a question never put", async () => {
    // Same standing as `no-scratch-shape`: setup failed, so the question was not
    // asked. If it counted as an answer the diff would compare it against the
    // fake and report a divergence nobody observed.
    expect(NOT_ASKED.has("no-named-slide")).toBe(true);
  });
});

describe("a detail that agrees with its own verdict", () => {
  afterEach(async () => {
    const { _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
  });

  it("says nothing about failing when the id came back", async () => {
    const { _rememberNamedShapeForTest } = await import("../src/render/powerpoint");
    const deck = [makeSlide("s1")];
    installHost(deck);
    const real = deck[0].shapes.addGeometricShape("Rectangle", { left: 1, top: 1, width: 9, height: 9 });
    _rememberNamedShapeForTest("s1", real.id);
    const sheet = await runHostProbes("fake", "test", {
      only: ["shape-resolve-held-slide-proxy"],
      passes: 1,
    });
    const row = sheet.answers.find((r) => r.id === "shape-resolve-held-slide-proxy")!;
    expect(row.answer).toBe("yes");
    // Round 244 shipped `yes` carrying "would not read back" in the same string.
    // The failure clause is computed from a shape that is not a null object,
    // which is also true on the success path, so it fired against its own answer.
    expect(row.detail).not.toContain("would not read back");
    expect(row.detail).not.toContain("no such shape");
    // The context that is true either way stays.
    expect(row.detail).toMatch(/slide holds [1-9]\d* shape\(s\)/);
  });
});

/**
 * Which of two facts `value=undefined` was.
 *
 * `tags-add-same-key-twice` has answered `other` — the word this file classes as
 * UNINFORMATIVE — in 15 of the last 15 rounds, always as `value=undefined`. That
 * covers a tag key that is NOT ON THE SLIDE, meaning the write did not land, and
 * a tag the host took and will not read back. `powerpoint.ts` puts
 * `DEMO_SLOT_TAG` on slides and reads it back, so the difference has an owner.
 */
describe("writing the same tag key twice", () => {
  const askIt = async () => {
    const sheet = await runHostProbes("fake", "test", { only: ["tags-add-same-key-twice"], passes: 1 });
    return sheet.answers.find((r) => r.id === "tags-add-same-key-twice")!;
  };

  it("says the tag is there when the host honours the write", async () => {
    installHost([makeSlide("s1")]);
    const row = await askIt();
    expect(row.answer).toBe("overwrites");
    expect(row.detail).toContain("the tag IS on the slide");
  });

  it("says the slide was named, and never claims the id changed", async () => {
    installHost([makeSlide("s1")]);
    const row = await askIt();
    // All three handles come from ONE captured `scratchId`, so the ids are equal
    // by construction. A report of "CHANGED under the probe" would be impossible
    // rather than alarming, and for seven rounds the same instrument reported
    // `stable` off three unread ids.
    expect(row.detail).toContain("slide named");
    expect(row.detail).not.toContain("should be impossible");
  });
});

/**
 * A held scratch id that stops naming the slide it named.
 *
 * Round 253 measured the bill: 61 of 64 scratch replacements happened because
 * `getItemOrNullObject(<held id>)` answered `isNullObject: true` while the slide
 * sat in the deck, and the clean-up swept 107 slides out of a deck of one
 * because delete-by-id recovered none. The run held `4123571114#123571113`; the
 * deck listed `256#2587447327`.
 */
describe("re-acquiring the scratch slide instead of buying another", () => {
  afterEach(() => {
    faults.newSlideIdSettlesAfter = null;
  });

  it("asks the deck for the settled id before adding a slide", async () => {
    installHost([makeSlide("s1")]);
    // TWO, not one. At one the id settles inside `addScratchSlide`'s own
    // verification, so the slide never becomes usable and the run never reaches
    // the state under test — the add fails instead. Two lets the add succeed and
    // the NEXT lookup meet a stale id, which is the sequence a round sees.
    faults.newSlideIdSettlesAfter = 2;
    setTracing(true);
    try {
      await runHostProbes("fake", "test", { only: ["load-id-populates-isnullobject"], passes: 1 });
      const reacquired = traceLog().entries.filter((e) =>
        /re-acquired the scratch slide by position instead of adding one/.test(e.message),
      );
      expect(reacquired.length, "a stale held id must be re-acquired, not replaced").toBeGreaterThan(0);
      // AND IT HAS TO WORK. A re-acquire that answers nothing is a listing spent
      // for no reason, and the add still happens behind it.
      expect(reacquired.some((e) => (e.data as { worked?: boolean }).worked)).toBe(true);
    } finally {
      setTracing(false);
    }
  });

  it("buys no slide when the re-acquire answers", async () => {
    installHost([makeSlide("s1")]);
    faults.newSlideIdSettlesAfter = 2;
    setTracing(true);
    try {
      await runHostProbes("fake", "test", { only: ["load-id-populates-isnullobject"], passes: 1 });
      // The POINT of the whole change. Deck size proves nothing here — the sweep
      // gives the slides back either way — so what has to be asserted is that no
      // replacement was bought at all.
      const replaced = traceLog().entries.filter((e) => /^replaced the scratch slide/.test(e.message));
      expect(replaced, "a re-acquire that works must not be followed by an add").toHaveLength(0);
    } finally {
      setTracing(false);
    }
  });
});

describe("what the re-acquire must NOT do", () => {
  afterEach(() => {
    faults.newSlideIdSettlesAfter = null;
    faults.refuseShapeAdds = false;
    setTracing(false);
  });

  it("does not re-acquire when the slide was fine and the SHAPE failed", async () => {
    installHost([makeSlide("s1")]);
    // `no-scratch-shape` means the slide resolved perfectly well and then
    // refused a shape. Re-acquiring asks the same question of the same slide and
    // learns nothing — it just spends a deck listing on every refused add.
    faults.refuseShapeAdds = true;
    setTracing(true);
    const sheet = await runHostProbes("fake", "test", { only: ["tags-on-fresh-shape"], passes: 1 });
    const row = sheet.answers.find((r) => r.id === "tags-on-fresh-shape")!;
    expect(row.answer, "the fixture needs a shape refusal, not a slide one").toBe("no-scratch-shape");
    expect(traceLog().entries.filter((e) => /re-acquired the scratch slide/.test(e.message))).toHaveLength(0);
    // And it SAID it decided not to, which is what separates "we chose not to
    // look" from "we looked and the id had not moved". The saving is a deck
    // listing per refused shape add, and without this line it is untestable.
    expect(
      traceLog().entries.some((e) =>
        /did not re-acquire — the slide resolved and the shape is what failed/.test(e.message),
      ),
    ).toBe(true);
  });

  it("does not adopt a re-acquire that answered nothing", async () => {
    installHost([makeSlide("s1")]);
    // The id settles AND the shape adds are refused, so the re-acquired slide
    // resolves and the question still cannot be put. Adopting that would record
    // a never-asked as the question's answer and leave `scratchId` pointing at a
    // slide the run never proved it could use.
    faults.newSlideIdSettlesAfter = 2;
    faults.refuseShapeAdds = true;
    setTracing(true);
    await runHostProbes("fake", "test", { only: ["tags-on-fresh-shape"], passes: 1 });
    const tried = traceLog().entries.filter((e) => /re-acquired the scratch slide/.test(e.message));
    expect(tried.length, "the fixture needs the re-acquire to actually be attempted").toBeGreaterThan(0);
    expect((tried[0].data as { worked?: boolean }).worked, "the shape refusal must defeat it").toBe(false);
    const replaced = traceLog().entries.filter((e) => /^replaced the scratch slide/.test(e.message));
    expect(replaced.length, "a failed re-acquire must still be followed by the add").toBeGreaterThan(0);
    // THE DISCRIMINATOR. If a failed re-acquire were adopted, `result` would
    // become its `no-scratch-shape` and the replacement would be recorded as
    // following a SHAPE failure — losing the fact that the slide is what went
    // wrong, which is the only thing that would have explained the round.
    expect((replaced[0].data as { after?: string }).after).toBe("no-scratch-slide");
  });
});

describe("whether buying a slide actually helped", () => {
  afterEach(() => {
    faults.refuseShapeAdds = false;
    faults.newSlideIdSettlesAfter = null;
    setTracing(false);
  });

  it("records that the replacement rescued the question", async () => {
    installHost([makeSlide("s1")]);
    // One refused add, then the fault lifts — so the replacement slide is the
    // one that answers, which is the case the buying exists for.
    faults.refuseShapeAddsTimes = 1;
    setTracing(true);
    await runHostProbes("fake", "test", { only: ["tags-on-fresh-shape"], passes: 1 });
    const said = traceLog().entries.filter((e) => /^the replacement slide answered/.test(e.message));
    expect(said.length, "buying a slide must say whether it helped").toBeGreaterThan(0);
    expect(said.some((e) => (e.data as { rescued?: boolean }).rescued)).toBe(true);
  });

  it("records that it did not, when the shape is refused throughout", async () => {
    installHost([makeSlide("s1")]);
    faults.refuseShapeAdds = true;
    setTracing(true);
    await runHostProbes("fake", "test", { only: ["tags-on-fresh-shape"], passes: 1 });
    const said = traceLog().entries.filter((e) => /^the replacement slide answered/.test(e.message));
    expect(said.length).toBeGreaterThan(0);
    // A slide bought against a shape refusal that a new slide cannot fix. Until
    // this line existed the cost was paid and the outcome never written down.
    expect(said.every((e) => (e.data as { rescued?: boolean }).rescued === false)).toBe(true);
  });

  it("calls a shape refusal by its name, not `unrecorded`", async () => {
    installHost([makeSlide("s1")]);
    faults.refuseShapeAdds = true;
    setTracing(true);
    await runHostProbes("fake", "test", { only: ["tags-on-fresh-shape"], passes: 1 });
    const replaced = traceLog().entries.filter((e) => /^replaced the scratch slide$/.test(e.message));
    expect(replaced.length).toBeGreaterThan(0);
    // `unrecorded` is reserved for the 16548 events archived before the cause
    // was carried out of the catch. Putting a known cause in that bucket makes
    // the pre-instrument rounds unreadable.
    expect((replaced[0].data as { why?: string }).why).toBe("shape-refused");
  });
});

describe("why a scratch slide was bought", () => {
  it("names a shape refusal instead of filing it as unrecorded", async () => {
    const { scratchReplacementWhy } = await import("../src/render/host-probe");
    // `unrecorded` is reserved for the ~17,000 events archived before any cause
    // was carried out of the catch. Putting a cause we DO have into that bucket
    // makes those rounds unreadable, which is the reading that motivated this
    // line existing at all.
    expect(scratchReplacementWhy({ answer: "no-scratch-shape" })).toBe("shape-refused");
  });

  it("prefers a cause the probe already carried over any inference", async () => {
    const { scratchReplacementWhy } = await import("../src/render/host-probe");
    // A `why` on the result is the host's own account. Inferring one from the
    // answer when a real one is present would overwrite evidence with a guess.
    expect(scratchReplacementWhy({ answer: "no-scratch-shape", why: "gone" })).toBe("gone");
  });

  it("says unrecorded ONLY when it genuinely does not know", async () => {
    const { scratchReplacementWhy } = await import("../src/render/host-probe");
    expect(scratchReplacementWhy({ answer: "no-scratch-slide" })).toBe("unrecorded");
  });

  it("is ONE function, because two inline copies drifted", async () => {
    // THE REASON THIS IS EXTRACTED. The expression lived inline at both places
    // that buy a slide and only the main loop was ever given the
    // `no-scratch-shape` case, so every partner replacement filed as
    // `unrecorded` — 68 of the last ten rounds' 248 replacements, 27%.
    const src = await import("fs").then((fs) => fs.readFileSync("src/render/host-probe.ts", "utf8"));
    const inlineCopies = src.match(/no-scratch-shape" \? "shape-refused" : "unrecorded"/g) ?? [];
    expect(inlineCopies.length, "the formula was inlined again instead of calling the helper").toBe(1);
    // Both call sites go through it.
    expect((src.match(/scratchReplacementWhy\(/g) ?? []).length).toBeGreaterThanOrEqual(3);
  });
});

describe("the two questions a chart's ink rests on", () => {
  /**
   * These probes exist to be answered by a REAL host, so what matters in the
   * suite is that each of their answers is REACHABLE — a probe that can only
   * ever say the comfortable thing is not evidence. The happy answers are
   * pinned in FAKE_BASELINE; these are the refusals.
   */
  const ask = async (id: string) => {
    const sheet = await runHostProbes("fake", "test");
    return sheet.answers.find((a) => a.id === id);
  };

  it("reports a host that places a rotated shape by its POST-rotation box", async () => {
    // The answer that would invalidate every diagonal the product draws. If
    // this were unreachable the probe would confirm our own assumption and
    // never be able to deny it, which is worse than not asking.
    installHost([makeSlide("s1")]);
    faults.reportRotatedBounds = true;
    try {
      const a = await ask("rotation-keeps-the-unrotated-box");
      expect(a?.answer, `did not notice the bounding-box host: ${a?.detail}`).toBe("POST-ROTATION-BOX");
      expect(a?.detail, "reported the finding without the measurement behind it").toMatch(/rotated extent/);
    } finally {
      faults.reportRotatedBounds = false;
    }
  });

  it("refuses to conclude when the rotation did not take", async () => {
    // A shape that never turned is still 80 wide, which is the same reading a
    // well-behaved host gives — so "no rotation happened" must not be reported
    // as "the host means the unrotated box". This is the guard against a
    // confirmation drawn from no evidence.
    installHost([makeSlide("s1")]);
    faults.ignoreRotationWrites = true;
    try {
      const a = await ask("rotation-keeps-the-unrotated-box");
      expect(a?.answer, `concluded from a shape that never turned: ${a?.detail}`).toBe("rotation-did-not-take");
    } finally {
      faults.ignoreRotationWrites = false;
    }
  });

  it("says so plainly on a host with no rotation API at all", async () => {
    /**
     * Not a hypothetical: PowerPointApi 1.10 is missing from most desktop Office
     * builds and from every volume-licensed one, permanently. Such a host cannot
     * answer this question, and must not be recorded as having answered it —
     * `no-rotation-api` is a "never asked" word, not an opinion about geometry.
     */
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.10");
    const a = await ask("rotation-keeps-the-unrotated-box");
    expect(a?.answer).toBe("no-rotation-api");
  });

  it("separates a host that refuses the write from one that swallows it", async () => {
    // Two different hosts and two different words. `rotation-did-not-take` is a
    // measurement that failed; this is a host that HAS the property and will
    // not let go of it, which is a fact about the host worth carrying.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.rotationWriteThrows = true;
    try {
      const a = await ask("rotation-keeps-the-unrotated-box");
      expect(a?.answer).toBe("rotation-not-writable");
    } finally {
      faults.rotationWriteThrows = false;
    }
  });

  it("notices a preset the host accepts and then does not draw", async () => {
    // The half the enum lookup cannot see. A name present in
    // `GeometricShapeType` is not a promise that the host will draw it, and a
    // shape that is there with no extent is invisible on the slide and looks
    // exactly like success from the calling code.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.presetDrawsNothing = "hexagon";
    try {
      const a = await ask("named-preset-resolves");
      expect(a?.answer).toBe("drew-nothing");
    } finally {
      faults.presetDrawsNothing = "";
    }
  });

  it("tells a host that answers nothing from one that answers nothing AFTER a rotation", async () => {
    /**
     * Rounds 336, 337 and 338 all came back `unreadable` — "width did not come
     * back a number" — while `named-preset-resolves` read a width in the same
     * shape of batch and got 24x24 in every one of them. One flat word could
     * not tell "this host answers no geometry" from "this host answers no
     * geometry AFTER a rotation write", and those call for different next
     * moves. The probe carries an unrotated control now; this is the host that
     * makes the control mean something.
     */
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.rotationBlindsTheRead = true;
    try {
      const a = await ask("rotation-keeps-the-unrotated-box");
      expect(a?.answer).toBe("no-read-after-rotation");
      expect(a?.detail, "did not report the control that makes this a finding").toMatch(/unrotated shape/);
    } finally {
      faults.rotationBlindsTheRead = false;
    }
  });

  it("names a preset this host's enum does not carry", async () => {
    // The failure that is otherwise invisible: `addGeometricShape(undefined)` is
    // accepted and draws a shape with no geometry.
    installHost([makeSlide("s1")]);
    faults.presetMissing = "hexagon";
    try {
      const a = await ask("named-preset-resolves");
      expect(a?.answer).toBe("MISSING-FROM-ENUM");
      expect(a?.detail, "did not say WHICH name was missing").toMatch(/hexagon/);
    } finally {
      faults.presetMissing = "";
    }
  });
});
