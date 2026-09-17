import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tools, no types.
import { mergeRanges, onlyMutated, rangesFromDiff, toMutateArg } from "../scripts/mutation-scope.mjs";
// @ts-expect-error — as above.
import { KNOWN_SURVIVORS, reportBody, scoreByFile, survivorsOf } from "../scripts/mutation-triage.mjs";

/**
 * The mutation job has never produced a report. Four runs were killed at
 * GitHub's 6h ceiling and five failed outright; Monday's left the arithmetic
 * behind — 3,254 of 18,724 mutants in 5h49m, so ~33 hours at that rate against
 * a 6 hour limit, with 671 survivors already.
 *
 * So the run is scoped to the week's diff, and these are the two pure halves of
 * that: which lines to mutate, and which survivors nobody has ruled on. Both
 * are tested without Stryker, because neither can be exercised by running the
 * job — it does not finish.
 */

describe("choosing what to mutate from a diff", () => {
  it("takes the + side of a hunk, which is the only side a mutation range can mean", () => {
    // `-` counts describe the file as it WAS. Mutating those line numbers would
    // point at whatever now happens to sit there.
    const diff = ["+++ b/src/core/chart.ts", "@@ -10,8 +12,3 @@", "+ a", "@@ -40 +50,2 @@", "+ b"].join("\n");
    expect(rangesFromDiff(diff)).toEqual([
      { file: "src/core/chart.ts", start: 12, end: 14 },
      { file: "src/core/chart.ts", start: 50, end: 51 },
    ]);
  });

  it("reads a missing count as one line, not zero", () => {
    // `@@ -5 +7 @@` is a single-line hunk. Defaulting to 0 would drop it.
    expect(rangesFromDiff("+++ b/src/core/a.ts\n@@ -5 +7 @@")).toEqual([{ file: "src/core/a.ts", start: 7, end: 7 }]);
  });

  it("contributes nothing for a pure deletion", () => {
    // `+0` means the new file gained no lines here, so there is nothing to mutate.
    expect(rangesFromDiff("+++ b/src/core/a.ts\n@@ -5,4 +7,0 @@")).toEqual([]);
  });

  it("ignores a deleted file", () => {
    expect(rangesFromDiff("+++ b/dev/null\n@@ -1,5 +0,0 @@")).toEqual([]);
  });

  it("keeps only what stryker.config.json actually mutates", () => {
    const rs = [
      { file: "src/core/chart.ts", start: 1, end: 2 },
      { file: "src/core/types.ts", start: 1, end: 2 },
      { file: "src/core/samples.ts", start: 1, end: 2 },
      { file: "src/render/powerpoint.ts", start: 1, end: 2 },
      { file: "test/chart.test.ts", start: 1, end: 2 },
    ];
    // types.ts and samples.ts carry `!` negations in `mutate`; render/ and test/
    // are outside it. Mutating any of them would instrument a file Stryker was
    // never told to touch, and the ranges would be silently ignored.
    expect(onlyMutated(rs).map((r: { file: string }) => r.file)).toEqual(["src/core/chart.ts"]);
  });

  it("merges hunks that are close, so one edited function is one range", () => {
    const rs = [
      { file: "a.ts", start: 10, end: 12 },
      { file: "a.ts", start: 15, end: 16 },
      { file: "a.ts", start: 90, end: 90 },
    ];
    expect(mergeRanges(rs)).toEqual([
      { file: "a.ts", start: 10, end: 16 },
      { file: "a.ts", start: 90, end: 90 },
    ]);
  });

  it("formats the spelling Stryker's mutate accepts", () => {
    expect(toMutateArg([{ file: "src/core/chart.ts", start: 120, end: 160 }])).toBe("src/core/chart.ts:120-160");
  });

  it("produces nothing at all when no mutated file changed", () => {
    // The workflow reads empty as "quiet week, exit after the dry run". An
    // accidental "" entry would become `--mutate ""` and mutate everything.
    expect(toMutateArg(mergeRanges(onlyMutated(rangesFromDiff("+++ b/docs/README.md\n@@ -1 +1,3 @@"))))).toBe("");
  });
});

describe("deciding which survivors are news", () => {
  const report = {
    files: {
      "src/core/chart.ts": {
        mutants: [
          { status: "Survived", mutatorName: "EqualityOperator", location: { start: { line: 12 } }, replacement: ">" },
          { status: "Killed", mutatorName: "BooleanLiteral", location: { start: { line: 20 } } },
          { status: "Timeout", mutatorName: "ArithmeticOperator", location: { start: { line: 30 } } },
          { status: "Ignored", mutatorName: "StringLiteral", location: { start: { line: 40 } } },
        ],
      },
    },
  };

  it("finds only the survivors, keyed file:line:mutator", () => {
    expect(survivorsOf(report)).toEqual([
      {
        key: "src/core/chart.ts:12:EqualityOperator",
        file: "src/core/chart.ts",
        line: 12,
        mutator: "EqualityOperator",
        replacement: ">",
      },
    ]);
  });

  it("counts a Timeout as killed and drops Ignored from the denominator", () => {
    // A timed-out mutant WAS caught — the suite hung on it rather than passing.
    // An ignored one was never run, so scoring it would move the percentage
    // without anyone changing a test.
    expect(scoreByFile(report)).toEqual([{ file: "src/core/chart.ts", killed: 2, total: 3, pct: 66.7 }]);
  });

  it("starts with an EMPTY register, and that is deliberate", () => {
    // Seeding it from guesses would silence findings nobody has looked at,
    // which is the exact failure the table exists to prevent. There has never
    // been a real report to seed it from.
    expect(KNOWN_SURVIVORS).toEqual({});
  });

  it("says how many survived as well as how many are new", () => {
    const body = reportBody([{ key: "a:1:X", replacement: "y" }], [{ key: "a:1:X" }, { key: "b:2:Y" }], []);
    expect(body).toContain("1 surviving mutant(s) nobody has ruled on, of 2");
    expect(body).toContain("no exposure");
  });

  it("warns that a single overall percentage is not the finding", () => {
    // The run mutates only the week's diff, so the headline score moves with
    // the size of the diff and means nothing across weeks.
    const body = reportBody([], [], [{ file: "a.ts", killed: 1, total: 4, pct: 25 }]);
    expect(body).toMatch(/single overall percentage is not/);
  });
});
