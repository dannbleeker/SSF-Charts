import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
// @ts-expect-error - .mjs script, deliberately outside tsconfig's typed surface
import * as regimes from "../scripts/host-regimes.mjs";

// Destructured after the import so the directive above stays attached to a
// single short line — prettier wraps a five-name named import, which pushes the
// `@ts-expect-error` off the statement it is suppressing and it silently stops
// applying.
const {
  explainBy,
  explainByRegime,
  flippedTogether,
  neverPutByRegime,
  verdictLine,
  collectionTimeline,
  COLLECTION_REFUSALS,
  stampSpread,
} = regimes;

/**
 * `scripts/host-regimes.mjs` — does the host's STATE account for a question that
 * changed its answer mid-round?
 *
 * The probe stamps every sample with the regime the host was in. Nothing read
 * those stamps mechanically before this: they were read by hand into
 * `UNSTABLE_ANSWERS`, which is how that table came to call a question "a coin"
 * that had answered the same way ten rounds running.
 *
 * The load-bearing case is `untested`. With three passes in three different
 * regimes, "every regime maps to one answer" is true by construction — a claim
 * that cannot fail is not evidence, and reporting it as `explained` would
 * manufacture a finding every time a question flipped. That is the same defect
 * as the rasterise control that always blamed the call because its arms ran in a
 * fixed order.
 */

type S = { answer: string; pass: number; atMs: number; regime: string };
const s = (answer: string, pass: number, regime: string): S => ({ answer, pass, atMs: pass * 1000, regime });

describe("does host state account for a mid-round flip", () => {
  it("says steady when the question never moved", () => {
    const r = explainByRegime([s("yes", 1, "healthy"), s("yes", 2, "collection-refused")]);
    expect(r.verdict).toBe("steady");
  });

  it("calls it EXPLAINED only when the mapping could have failed", () => {
    // healthy sampled twice and agreed both times, so "healthy means threw" was
    // a claim with a way to come out wrong. That is what makes it evidence.
    const r = explainByRegime([s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")]);
    expect(r.verdict).toBe("explained");
    expect(r.repeated).toEqual(["healthy"]);
  });

  it("calls it UNTESTED when every regime was sampled exactly once", () => {
    // THE GUARD THIS TOOL EXISTS FOR. Three passes, three regimes, three
    // answers: "every regime maps to one answer" is unfalsifiable here, so it
    // must NOT be reported as explained.
    const r = explainByRegime([
      s("threw", 1, "healthy"),
      s("yes", 2, "slide-trouble"),
      s("short-0", 3, "collection-refused"),
    ]);
    expect(r.verdict).toBe("untested");
  });

  it("calls it a COIN when one regime gave two different answers", () => {
    const r = explainByRegime([s("threw", 1, "healthy"), s("yes", 2, "healthy"), s("yes", 3, "collection-refused")]);
    expect(r.verdict).toBe("coin");
    expect(r.split[0].regime).toBe("healthy");
  });

  it("ignores passes that never put the question, rather than counting them as a face", () => {
    // A `no-scratch-slide` pass says nothing about the host. Counting it would
    // report a steady question as a flipper on a bad night — the same rule
    // `history()` applies across rounds.
    const r = explainByRegime([
      s("yes", 1, "healthy"),
      s("no-scratch-slide", 2, "slide-trouble"),
      s("yes", 3, "healthy"),
    ]);
    expect(r.verdict).toBe("steady");
    expect(r.real).toBe(2);
  });

  /**
   * THE OTHER HALF OF THAT RULE, which was missing — found 2026-09-28.
   *
   * There are TWO tiers of non-answer, both in `host-baseline.mjs`:
   * `NEVER_ASKED` ("the question was never put", covered by the test above) and
   * `UNINFORMATIVE_ANSWERS` ("it was put and produced nothing to name" — other,
   * unreadable, silent, not-a-short-read, none-of-ours, no-shapes-to-list).
   * `explainBy` filtered only the first, so the second counted as a FACE: a
   * question answered `yes` twice and `unreadable` once looked like it had
   * changed its answer, and the machinery then explained a flip that never
   * happened.
   *
   * Measured against the committed archive: `host-regimes.mjs
   * rounds/302-aa459dd.json` reported `picture-then-shape-read` as EXPLAINED by
   * regime, and another question as COIN because "healthy gave threw and
   * unreadable". Both verdicts were manufactured out of a non-answer.
   * `host-baseline.mjs` applies both tiers at its own line 1211 — the rule was
   * right and one of its halves had not reached here.
   */
  it("does not count an uninformative answer as a face", () => {
    const r = explainByRegime([
      s("yes", 1, "healthy"),
      s("yes", 2, "healthy"),
      s("unreadable", 3, "collection-refused"),
    ]);
    expect(
      r.verdict,
      `"unreadable" means the question was put and named nothing. Counting it as a second answer turns ` +
        `a steady question into a flip, and then sets about explaining the flip.`,
    ).toBe("steady");
    expect(r.real).toBe(2);
  });

  it.each(["other", "unreadable", "silent", "not-a-short-read", "none-of-ours", "no-shapes-to-list"])(
    "ignores %s in the faces and in the regime mapping alike",
    (word) => {
      const r = explainByRegime([s("yes", 1, "healthy"), s(word, 2, "healthy")]);
      expect(r.verdict).toBe("steady");
      expect(r.faces).toEqual(["yes"]);
      // And it must not survive into the per-regime mapping, or a regime would
      // look like it carried two answers while the host said one thing.
      for (const m of r.mapping as { answers: string[] }[]) expect(m.answers).not.toContain(word);
    },
  );

  it("still calls a genuine two-answer flip a coin", () => {
    // THE DIRECTION THAT MATTERS: the fix must not silence real disagreement.
    // Two REAL answers, in the shape a coin actually needs — two buckets with
    // one of them split. (Two answers in a SINGLE bucket is `blind`, which the
    // renderer test above records this file getting wrong once already.)
    const r = explainByRegime([
      s("short-2", 1, "healthy"),
      s("short-1", 2, "healthy"),
      s("short-1", 3, "slide-trouble"),
    ]);
    expect(r.verdict).toBe("coin");
    expect(r.faces).toEqual(["short-2", "short-1"]);
  });
});

describe("a trigger and its partner: one mechanism or two", () => {
  const trigger = [s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")];

  it("reads a shared boundary as lockstep", () => {
    const partner = [s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")];
    expect(flippedTogether(trigger, partner).verdict).toBe("lockstep");
  });

  it("reads different boundaries as independent", () => {
    // The partner moved between passes 1 and 2; the trigger between 2 and 3.
    const partner = [s("threw", 1, "healthy"), s("yes", 2, "healthy"), s("yes", 3, "collection-refused")];
    expect(flippedTogether(trigger, partner).verdict).toBe("independent");
  });

  it("refuses to compare when they share fewer than two answered passes", () => {
    // Nothing can be concluded from one shared pass, and saying "lockstep"
    // there would be the unfalsifiable shape again.
    const partner = [
      s("no-scratch-slide", 1, "slide-trouble"),
      s("no-scratch-slide", 2, "slide-trouble"),
      s("yes", 3, "collection-refused"),
    ];
    expect(flippedTogether(trigger, partner).verdict).toBe("not-comparable");
  });

  it("says neither moved when both held still across the shared passes", () => {
    const steady = [s("yes", 1, "healthy"), s("yes", 2, "healthy"), s("yes", 3, "collection-refused")];
    expect(flippedTogether(steady, steady).verdict).toBe("neither-moved");
  });
});

describe("what state the host was in for the questions it never put", () => {
  it("tallies the never-put attempts by regime, and names the questions", () => {
    const answers = [
      {
        id: "binding-names-shape-later",
        answer: "no-scratch-slide",
        samples: [s("no-scratch-slide", 1, "slide-trouble"), s("no-scratch-slide", 2, "slide-trouble")],
      },
      {
        id: "tag-on-group-survives",
        answer: "no-scratch-shape",
        samples: [s("no-scratch-shape", 1, "collection-refused")],
      },
      { id: "answered-one", answer: "yes", samples: [s("yes", 1, "healthy")] },
    ];
    const np = neverPutByRegime(answers);
    expect(np.ids).toEqual(["binding-names-shape-later", "tag-on-group-survives"]);
    expect(np.byRegime).toEqual([
      ["slide-trouble", 2],
      ["collection-refused", 1],
    ]);
  });
});

describe("rendering a verdict as prose", () => {
  /**
   * The first version of this built all four strings in one object literal and
   * picked one, so `split.map` ran on an `explained` result — where `split` does
   * not exist. It threw on the FIRST real round it was pointed at, and every
   * unit test above passed, because they call the decision functions and never
   * render a line. A tool that cannot print its own answer is not a tool.
   */
  const cases = [
    [[s("yes", 1, "healthy"), s("yes", 2, "healthy")], "steady"],
    [[s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")], "EXPLAINED"],
    [[s("threw", 1, "healthy"), s("yes", 2, "slide-trouble")], "UNTESTED"],
    // A genuine coin needs TWO buckets with one of them split. Both answers in
    // a single bucket is `blind`, which this list used to assert as COIN.
    [[s("threw", 1, "healthy"), s("yes", 2, "healthy"), s("yes", 3, "collection-refused")], "COIN"],
    [[s("threw", 1, "healthy"), s("yes", 2, "healthy")], "BLIND"],
  ] as const;

  it("renders every verdict without touching a field that verdict does not carry", () => {
    for (const [samples, expected] of cases) {
      const r = explainByRegime(samples as unknown as S[]);
      expect(() => verdictLine(r), `${expected} threw while rendering`).not.toThrow();
      expect(verdictLine(r)).toContain(expected);
    }
  });
});

describe("reading a stamp other than the regime", () => {
  /**
   * Round 17 eliminated `regime` as the state behind the held-slide-proxy flip,
   * and `scratch` is the candidate that replaced it. The decision is the same
   * either way — group the samples by a stamp and ask whether the mapping could
   * have failed — so the field is a parameter rather than a second copy.
   *
   * The case that matters is DISAGREEMENT between the two stamps: the same
   * samples reading COIN by regime and EXPLAINED by scratch is precisely the
   * result that would answer the open question, and the tool has to be able to
   * express it.
   */
  const samples = [
    { answer: "threw", pass: 1, atMs: 16300, regime: "healthy", scratch: "first-slide" },
    { answer: "yes", pass: 2, atMs: 33900, regime: "collection-refused", scratch: "reused-slide" },
    { answer: "threw", pass: 3, atMs: 55600, regime: "collection-refused", scratch: "first-slide" },
  ];

  it("reads the same samples two ways, and they can disagree", () => {
    // Round 17's actual shape: one regime produced both faces, so by regime it
    // is a coin. Grouped by scratch state each stamp maps to one answer, and
    // `first-slide` was sampled twice and agreed — so it could have failed.
    expect(explainByRegime(samples).verdict).toBe("coin");
    const byScratch = explainBy(samples, "scratch");
    expect(byScratch.verdict).toBe("explained");
    expect(byScratch.repeated).toEqual(["first-slide"]);
  });

  it("says a stamp the round never recorded is BLIND, not a coin", () => {
    // Rounds saved before the stamp existed carry no `scratch` at all, so every
    // sample lands in one `unknown` bucket. That must NOT read as a coin: a coin
    // is a claim about the host, and the only thing being observed here is that
    // the field is missing.
    const older = samples.map(({ scratch: _drop, ...rest }) => rest);
    expect(explainBy(older, "scratch").verdict).toBe("blind");
    expect(explainBy(older, "scratch").mapping.map((m: { regime: string }) => m.regime)).toEqual(["unknown"]);
  });

  it("calls a stamp that never moved for this question BLIND, however many samples", () => {
    // THE CASE THAT MADE THIS VERDICT EXIST. Tested retrospectively against
    // round 17, `shape-add-held-slide-proxy` ran on a slide replaced 0.2-0.5s
    // before every single ask, so the scratch stamp reads `fresh-slide` on all
    // three samples — and the answers still disagree. Without `blind` the tool
    // answers `coin`, which reads as a fact about the host when it is a fact
    // about the stamp being constant here.
    const constant = [
      { answer: "threw", pass: 1, atMs: 16300, regime: "healthy", scratch: "fresh-slide" },
      { answer: "yes", pass: 2, atMs: 33900, regime: "collection-refused", scratch: "fresh-slide" },
      { answer: "threw", pass: 3, atMs: 55600, regime: "collection-refused", scratch: "fresh-slide" },
    ];
    expect(explainBy(constant, "scratch").verdict).toBe("blind");
    // The same samples by REGIME are a real coin — two buckets, one split. So
    // the two verdicts are genuinely different readings of one round, which is
    // the whole reason both stamps are printed.
    expect(explainByRegime(constant).verdict).toBe("coin");
  });

  it("still defaults to the regime when no field is named", () => {
    expect(explainBy(samples).verdict).toBe(explainByRegime(samples).verdict);
  });
});

describe("what the shape collection did, and whether that is a time effect", () => {
  /**
   * The check that stopped a wrong conclusion being shipped.
   *
   * Read in time order, round 17 looks like a collection that refuses, comes
   * back, and refuses again — three clean bursts. Grouped by QUESTION it is
   * nothing of the kind: seven of its eight collection questions gave the same
   * verdict on every pass, and the interleaving is the fixed question order
   * cycling through three passes.
   *
   * A summary that only said "an answer came after a refusal" reports RECOVERY
   * on a host that never changed, which is what the first version of this did.
   */
  const q = (id: string, verdicts: string[]) => ({
    id,
    answer: verdicts[0],
    samples: verdicts.map((answer, i) => ({ answer, pass: i + 1, atMs: (i + 1) * 10_000, regime: "healthy" })),
  });

  it("does not call it a time effect when every question is constant", () => {
    // Answers and refusals interleave in TIME — the refusing question is asked
    // after the answering one on every pass — and nothing varied.
    const ct = collectionTimeline([
      q("shape-add-fresh-slide-proxy", ["yes", "yes", "yes"]),
      q("shapes-items-count-honest", ["unreadable", "unreadable", "unreadable"]),
    ]);
    expect(ct.answered).toBe(3);
    expect(ct.refused).toBe(3);
    expect(ct.variesOverTime).toBe(false);
    expect(ct.alwaysAnswered).toEqual(["shape-add-fresh-slide-proxy"]);
    expect(ct.alwaysRefused).toEqual(["shapes-items-count-honest"]);
  });

  it("names only the questions that actually changed verdict", () => {
    const ct = collectionTimeline([
      q("shapes-items-count-honest", ["unreadable", "unreadable", "unreadable"]),
      q("shape-add-positional-slide-proxy", ["not-listed", "yes", "yes"]),
    ]);
    expect(ct.variesOverTime).toBe(true);
    expect(ct.varied).toEqual(["shape-add-positional-slide-proxy"]);
  });

  it("counts only questions that ask the collection, and never a never-put pass", () => {
    const ct = collectionTimeline([
      q("shapes-items-count-honest", ["unreadable", "no-scratch-slide"]),
      // Not a collection question — must not appear at all.
      q("tags-on-fresh-shape", ["yes", "yes"]),
    ]);
    expect(ct.events.filter((e: { verdict: string }) => e.verdict !== "never-put")).toHaveLength(1);
    expect(ct.answered).toBe(0);
    expect(ct.alwaysAnswered).toEqual([]);
  });

  it("keeps the refusal vocabulary in step with the probe's own", () => {
    // Duplicated because a .mjs script cannot import from the TypeScript source
    // — the same unavoidable copy `NEVER_ASKED` carries, held the same way.
    const src = readFileSync("src/render/host-probe.ts", "utf8");
    for (const word of COLLECTION_REFUSALS)
      expect(src, `the probe never emits "${word}", so this vocabulary has drifted`).toContain(`"${word}"`);
  });
});

describe("a stamp that cannot separate anything", () => {
  it("reports saturation, and the share it is saturated at", () => {
    // `regime` sat at 88% one value in round 17 — worse than the 85% sticky flag
    // `regimeFrom` was written to replace. A saturated stamp emits `untested`
    // and `blind`, both of which read as caution rather than a broken gauge.
    const answers = [
      {
        id: "a",
        answer: "yes",
        samples: Array.from({ length: 9 }, (_, i) => ({
          answer: "yes",
          pass: 1,
          atMs: i,
          regime: "collection-refused",
        })).concat([{ answer: "yes", pass: 1, atMs: 9, regime: "healthy" }]),
      },
    ];
    const sp = stampSpread(answers, "regime");
    expect(sp.total).toBe(10);
    expect(sp.saturated).toBe(true);
    expect(sp.byValue[0]).toEqual(["collection-refused", 9]);
  });

  it("does not cry saturation on a stamp that is spread", () => {
    const answers = [
      {
        id: "a",
        answer: "yes",
        samples: [
          { answer: "yes", pass: 1, atMs: 0, regime: "healthy" },
          { answer: "yes", pass: 2, atMs: 1, regime: "collection-refused" },
          { answer: "yes", pass: 3, atMs: 2, regime: "slide-trouble" },
        ],
      },
    ];
    expect(stampSpread(answers, "regime").saturated).toBe(false);
  });
});
