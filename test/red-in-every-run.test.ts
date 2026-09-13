import { describe, expect, it } from "vitest";
// @ts-expect-error — a .mjs tool script with no types, imported for its pure helpers.
import { redInEveryRun, repoFrame } from "../scripts/red-in-every-run.mjs";

/**
 * The quality sweep runs the suite three times under deliberate CPU load. When
 * all three go red the SAME way, `flaky.mjs` correctly reports no flake — and
 * the job used to exit 1 saying so while naming nothing. This is the script
 * that names them.
 *
 * THE FIXTURES BELOW ARE SHAPED FROM A REAL REPORT, not from the docs. Vitest
 * 4's json reporter writes `Error: STACK_TRACE_ERROR` as `failureMessages[0]`
 * and opens the stack with several `@vitest/runner` frames — so the reason is
 * absent and the first stack line is always a dependency. A first draft printed
 * that first line as the failure reason and matched it against /timed out/;
 * it would have printed a placeholder and never matched once.
 */

const STACK = [
  "Error: STACK_TRACE_ERROR",
  "    at task (file:///C:/repo/node_modules/@vitest/runner/dist/chunk-artifact.js:1784:27)",
  "    at Object.<anonymous> (file:///C:/repo/node_modules/@vitest/runner/dist/chunk-artifact.js:1817:16)",
  "    at C:/devtools/SSF-Charts/test/rounds.test.ts:43:3",
].join("\n");

const result = (name: string, status: "passed" | "failed", stack = STACK) => ({
  fullName: name,
  title: name,
  status,
  failureMessages: status === "failed" ? [stack] : [],
});

const report = (suites: { name: string; tests: ReturnType<typeof result>[] }[]) => ({
  testResults: suites.map((s) => ({ name: s.name, assertionResults: s.tests })),
});

describe("naming the tests that were red in every run", () => {
  it("skips dependency frames and reports the repo's own file:line", () => {
    expect(repoFrame(STACK)).toBe("test/rounds.test.ts:43:3");
  });

  it("returns nothing when a frame is only ever in node_modules", () => {
    expect(repoFrame("Error: x\n    at foo (/a/node_modules/b/c.js:1:1)")).toBe("");
    expect(repoFrame(undefined)).toBe("");
  });

  it("reports a test that failed in every run", () => {
    const r = report([{ name: "test/rounds.test.ts", tests: [result("archive reads", "failed")] }]);
    expect(redInEveryRun([r, r, r])).toEqual([
      { name: "test/rounds.test.ts > archive reads", where: "test/rounds.test.ts:43:3" },
    ]);
  });

  it("does NOT report a flake — that is flaky.mjs's job, and blurring them defeats both", () => {
    // The whole point of the split. A test red once and green twice disagreed
    // with itself; reporting it here would make "broken build" and "flake"
    // read identically in the log, which is the distinction the sweep exists on.
    const red = report([{ name: "t.ts", tests: [result("sometimes", "failed")] }]);
    const green = report([{ name: "t.ts", tests: [result("sometimes", "passed")] }]);
    expect(redInEveryRun([red, green, green])).toEqual([]);
  });

  it("reports a test red in every run that EXECUTED it, even if a run died early", () => {
    // A run that crashes before reaching a file lists nothing for it. Requiring
    // failures === total-runs would then hide the failure; requiring
    // failures === runs-that-ran-it does not.
    const red = report([{ name: "t.ts", tests: [result("slow", "failed")] }]);
    const partial = report([]);
    expect(redInEveryRun([red, red, partial])).toHaveLength(1);
  });

  it("survives an empty, missing or malformed report without inventing a finding", () => {
    expect(redInEveryRun([])).toEqual([]);
    expect(redInEveryRun([{}, null, undefined])).toEqual([]);
    expect(redInEveryRun([{ testResults: [{ name: "t.ts" }] }])).toEqual([]);
  });

  it("keeps suites apart when two files share a test name", () => {
    const r = report([
      { name: "a.ts", tests: [result("same name", "failed")] },
      { name: "b.ts", tests: [result("same name", "failed")] },
    ]);
    // Annotated because the helper comes from an untyped .mjs — the import
    // carries a @ts-expect-error, so nothing infers this callback's parameter.
    const names = redInEveryRun([r]).map((f: { name: string }) => f.name);
    expect(names).toEqual(["a.ts > same name", "b.ts > same name"]);
  });
});
