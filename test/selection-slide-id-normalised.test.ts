import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { deckIdForSelectedSlide } from "../src/render/powerpoint";

/**
 * EVERY SELECTION READER NORMALISES THE SLIDE ID, not just the one that was
 * fixed.
 *
 * A SlideRange's id is not roundtrippable — office-js#2474, reported on Windows
 * desktop and closed `not planned`. The deck's id for a slide is the range's id
 * plus a `#suffix`, and `slides.getItem(rangeId)` answers InvalidArgument where
 * `getItemAt(index)` works.
 *
 * The failure is the silent kind, and `deckIdForSelectedSlide`'s own docstring
 * spells it out: an update resolves the un-normalised id with
 * `getItemOrNullObject`, gets a null object, and the chart is filtered out as
 * "the slide is gone, nothing to do". The user clicks their chart, edits it,
 * and nothing happens — no error anywhere.
 *
 * `loadChartFromSelection` was fixed. `listChartsInSelection` — which builds the
 * very same EditTargets, for the multi-select path — kept the raw range id, and
 * was found on 2026-09-28. Same shape as "the fix was written and one call site
 * was missed", which this repo has now paid for several times.
 *
 * The ids happen to round-trip on this web host today. That is what makes it a
 * latent defect rather than a visible one, and why the guard is a SWEEP over
 * every reader rather than a test of one.
 */
const SRC = readFileSync("src/render/powerpoint.ts", "utf8");

/**
 * The CODE of a named exported function — comments stripped.
 *
 * Stripping them is load-bearing, and the first version of this file proved it.
 * Both functions carry a comment EXPLAINING `deckIdForSelectedSlide`, so a
 * sweep matching the bare identifier matched the explanation: a mutation that
 * removed the actual call left the comment behind and the test went green. A
 * detector that reads its own postmortem as compliance is worse than none —
 * the same trap `is-main.test.ts` strips comments for, arriving from the
 * opposite side.
 */
function bodyOf(name: string): string {
  const at = SRC.indexOf(`export async function ${name}`);
  expect(at, `${name} is gone from powerpoint.ts — re-point this test`).toBeGreaterThan(-1);
  const rest = SRC.slice(at);
  const end = rest.indexOf("\nexport ", 1);
  const body = end > 0 ? rest.slice(0, end) : rest;
  return body
    .split(/\/\*[\s\S]*?\*\//)
    .join("")
    .split("\n")
    .map((l) => l.replace(/\/\/.*$/, ""))
    .join("\n");
}

describe("a slide id taken from the selection", () => {
  it.each(["loadChartFromSelection", "listChartsInSelection"])(
    "%s matches the range id against the deck's own list",
    (fn) => {
      const body = bodyOf(fn);
      // The CALL, not the name: a comment mentioning the normaliser is not a
      // use of it, and matching the bare identifier is how the first version
      // of this test passed against code that had stopped calling it.
      expect(
        body,
        `${fn} builds an EditTarget from the RAW SlideRange id. That id is not roundtrippable ` +
          `(office-js#2474), so a later edit resolves it to a null object and silently does nothing.`,
      ).toMatch(/deckIdForSelectedSlide\(\s*rangeId\s*,\s*deckIds\s*\)/);
      // …and it must actually load the deck's ids, or the normaliser has
      // nothing to match against and quietly returns the range id unchanged.
      expect(body, `${fn} calls the normaliser without loading the deck's ids`).toMatch(/deck\.load\("items\/id"\)/);
    },
  );

  /**
   * The normaliser's own contract, so the sweep above cannot pass against a
   * function that has been gutted.
   */
  it("prefers the deck's suffixed id, and refuses rather than guessing", () => {
    expect(deckIdForSelectedSlide("256", ["256#2587447327", "257#3204983374"])).toBe("256#2587447327");
    // Already the deck's own id: unchanged.
    expect(deckIdForSelectedSlide("256#2587447327", ["256#2587447327"])).toBe("256#2587447327");
    expect(deckIdForSelectedSlide(undefined, ["256#a"])).toBeUndefined();
    /**
     * UNDEFINED, NOT THE RANGE ID. The deck answered and nothing matched, so
     * the prefix assumption is wrong on this host — and two slides answering to
     * one prefix would mean guessing which of them the user is looking at. A
     * wrong guess puts their edit on the wrong slide, which is worse than the
     * silent no-op this whole repair exists to prevent.
     *
     * The fall back to the range id belongs at the CALL SITE, and only for a
     * host that would not list the deck at all — `deckIds.length ? … : rangeId`.
     */
    expect(deckIdForSelectedSlide("999", ["256#a", "257#b"])).toBeUndefined();
    expect(deckIdForSelectedSlide("256", ["256#a", "256#b"])).toBeUndefined();
  });

  it("falls back to the range id only when the deck would not answer at all", () => {
    for (const fn of ["loadChartFromSelection", "listChartsInSelection"]) {
      expect(
        bodyOf(fn),
        `${fn} does not keep the range id for a host that will not list the deck — every round on this ` +
          `host has relied on that path`,
      ).toMatch(/deckIds\.length \? deckIdForSelectedSlide\(rangeId, deckIds\) : rangeId/);
    }
  });
});
