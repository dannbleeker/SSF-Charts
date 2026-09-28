import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types.
import { countOf, judgeCount } from "../scripts/test-count.mjs";

/**
 * The guard against a suite that silently shrinks.
 *
 * `CLAUDE.md` records the incident: a reorg deleted 43 tests and the suite still
 * went green, because nothing compared before against after. What is asserted
 * here is the decision, not the number — the number lives in
 * `test/fixtures/test-count.json` and rises on its own.
 */
describe("the test-count high-water mark", () => {
  it("fails when the suite has shrunk, and says by how much", () => {
    const v = judgeCount(2035, 2078);
    expect(v.ok).toBe(false);
    expect(v.message).toContain("43 test(s) went missing");
    // And tells the reader what to do about a deliberate drop, in the same
    // breath — a gate that only says no gets switched off.
    expect(v.message).toContain("--update");
  });

  it("passes when the suite grew, and when it stood still", () => {
    // Growth must cost nothing. An exact-count gate taxes every PR that adds a
    // test with a second edit and a merge conflict, which is how a gate earns
    // the resentment that gets it deleted.
    expect(judgeCount(2079, 2078).ok).toBe(true);
    expect(judgeCount(2078, 2078).ok).toBe(true);
  });

  it("refuses a report that is not a vitest run", () => {
    // The failure mode that would make this gate silently useless: point it at
    // the wrong file, read `undefined` as a count, and compare nothing forever.
    expect(() => countOf({ someOtherTool: true })).toThrow(/not a vitest JSON report/);
    expect(() => countOf(null)).toThrow();
    expect(countOf({ numTotalTests: 7 })).toBe(7);
  });

  /**
   * THE ONE EDIT THIS GATE COULD NOT SEE — found 2026-09-28.
   *
   * `numTotalTests` counts tests that never executed, so adding `.skip` to a
   * `describe` keeps the number identical while the suite shrinks in exactly
   * the sense this gate exists to catch. The likeliest way to lose coverage
   * was the one way that stayed invisible — and this file's own docstring says
   * the point is that a deletion becomes "something a reviewer sees rather
   * than something the suite absorbs". A skip absorbed it.
   *
   * `rounds-gate.mjs` states the same rule one layer up: "A skipped scenario
   * is not a passing one."
   */
  it("does not count a skipped test as one the suite ran", () => {
    expect(
      countOf({ numTotalTests: 100, numPendingTests: 7, numTodoTests: 3 }),
      "a `.skip` keeps numTotalTests whole, so the gate saw no shrink while ten tests stopped running",
    ).toBe(90);
  });

  it("degrades to the total when a report omits the skip counts", () => {
    // An older or partial report carries no `numPendingTests`. Treating that
    // as zero is the previous behaviour exactly; throwing would turn a
    // cosmetic change in vitest's report into a broken gate.
    expect(countOf({ numTotalTests: 42 })).toBe(42);
    expect(countOf({ numTotalTests: 42, numPendingTests: 2 })).toBe(40);
  });
});
