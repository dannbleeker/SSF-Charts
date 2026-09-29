import { describe, expect, it } from "vitest";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { functionSizes, keyOf, largeFunctions, toMark, type SizeMark } from "./helpers/function-sizes";

/**
 * NO FUNCTION IN `src/` GETS BIGGER THAN IT IS TODAY.
 *
 * WHY THIS EXISTS. On 2026-09-29 this repo had six functions over 800 lines and
 * 38 over 200, and the four files holding most of them appeared in 49% of all
 * 1,363 commits. Nothing noticed them growing, because nothing was looking:
 * there is no `max-lines-per-function` here and adding one would have gone red
 * 38 times on arrival and been switched off the same week.
 *
 * A ratchet is the shape that works. It records what is true today and fails
 * only on growth, so the debt is frozen rather than demanded — and every entry
 * that shrinks is a diff a reviewer can see.
 *
 * ── MEASURED IN CODE LINES, WHICH IS THE WHOLE DESIGN ───────────────────────
 * This source is 48% comment. Counting raw lines would fire every time somebody
 * wrote down WHY — the single habit here most worth keeping — and a guard that
 * punishes good practice is one that gets deleted. Comment-only and blank lines
 * are dropped, and a function's doc block is not part of it at all. See
 * `test/helpers/function-sizes.ts`.
 *
 * ── THE FLOOR ───────────────────────────────────────────────────────────────
 * 150 code lines. Below that the noise of ordinary editing would dominate and
 * the fixture would churn on every commit. Above it, a function is something a
 * reader has to hold in their head.
 *
 * ── WHEN IT GOES RED ────────────────────────────────────────────────────────
 * Either split the function, or re-record deliberately:
 *
 *     UPDATE_FUNCTION_SIZES=1 node ./node_modules/vitest/vitest.mjs run test/function-size.test.ts
 *
 * Re-recording is a legitimate answer — code grows for real reasons — but it
 * shows up in the diff, which is the entire point. Absorbing it silently is
 * what this exists to prevent.
 */
const FIXTURE = "test/fixtures/function-sizes.json";
const FLOOR = 150;

/**
 * Slack, in code lines.
 *
 * Not zero, and not because zero is too strict: `functionSizes` counts a line
 * as code when anything survives comment-stripping, so a multi-line call
 * reformatted by Prettier can move a count by one or two without a statement
 * being added. Three absorbs that and nothing else — a real addition to a
 * 300-line function is never three lines.
 */
const SLACK = 3;

describe("function sizes in src/", () => {
  const measured = largeFunctions(FLOOR);
  const mark: SizeMark = existsSync(FIXTURE) ? (JSON.parse(readFileSync(FIXTURE, "utf8")) as SizeMark) : {};

  if (process.env.UPDATE_FUNCTION_SIZES) {
    writeFileSync(FIXTURE, `${JSON.stringify(toMark(measured), null, 2)}\n`);
  }

  it("measured something, or this is asserting over an empty list", () => {
    // The vacuous pass this suite has now been swept for twice. If the AST walk
    // or the file walk breaks, every assertion below passes silently.
    expect(measured.length, "no functions at all over the floor — the measurement broke").toBeGreaterThan(10);
    expect(Object.keys(mark).length, `${FIXTURE} is empty — re-record it`).toBeGreaterThan(10);
  });

  it("has no function that grew", () => {
    const grew = measured
      .filter((f) => mark[keyOf(f)] !== undefined && f.code > mark[keyOf(f)] + SLACK)
      .map((f) => `${keyOf(f)}  ${mark[keyOf(f)]} -> ${f.code} code lines`);
    expect(
      grew,
      `these functions grew past their recorded size:\n${grew.join("\n")}\n\n` +
        `Split it, or re-record with UPDATE_FUNCTION_SIZES=1 so the growth shows in the diff.`,
    ).toEqual([]);
  });

  it("has no NEW function over the floor that nobody decided to add", () => {
    // A fresh 400-line function is the thing this is really for. Growth of an
    // existing one is visible in review; a new one arrives looking normal.
    const fresh = measured.filter((f) => mark[keyOf(f)] === undefined).map((f) => `${keyOf(f)}  ${f.code} code lines`);
    expect(
      fresh,
      `new function(s) over ${FLOOR} code lines:\n${fresh.join("\n")}\n\n` +
        `That may be right — but it is a decision, so record it with UPDATE_FUNCTION_SIZES=1.`,
    ).toEqual([]);
  });

  it("counts code rather than prose, so documenting a function cannot trip it", () => {
    // Proven against a synthetic pair rather than against the tree being quiet.
    // This is the design claim of the whole file: if it were false, the ratchet
    // would fight the comment density that makes this codebase legible.
    // In the OS temp dir, not the repo: a test that writes into `test/fixtures/`
    // leaves a file behind for the next reader to wonder about, and `git status`
    // dirty after a test run is how a real change gets committed by accident.
    const tmp = join(mkdtempSync(join(tmpdir(), "ssf-fnsize-")), "probe.ts");
    const bare = ["function sample(): number {", "  const a = 1;", "  const b = 2;", "  return a + b;", "}"].join("\n");
    const documented = [
      "function sample(): number {",
      "  /**",
      "   * A long explanation that adds no complexity whatsoever,",
      "   * spread over several lines, as this repo does everywhere.",
      "   */",
      "  const a = 1;",
      "  // and a line comment",
      "",
      "  const b = 2;",
      "  return a + b;",
      "}",
    ].join("\n");
    try {
      writeFileSync(tmp, bare);
      const before = functionSizes(tmp).find((f) => f.name === "sample")?.code;
      writeFileSync(tmp, documented);
      const after = functionSizes(tmp).find((f) => f.name === "sample")?.code;
      expect(before, "the probe function was not measured at all").toBe(5);
      expect(after, "prose counted as code — the ratchet would punish documenting").toBe(before);
    } finally {
      rmSync(join(tmp, ".."), { recursive: true, force: true });
    }
  });
});
