import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describeOffenders, expectSweptSomething, sweep } from "./helpers/module-source";

/**
 * NO WRITE PATH FAILS SILENTLY.
 *
 * Found 2026-09-28. Every host action in the pane runs inside `guard()`, whose
 * catch is the one thing that turns a thrown update into
 * `note("Failed: {error}", "err")`. The auto-update timer is the only write
 * path that does not go through it, and it read:
 *
 *     void doInsert(false).catch(() => {});
 *
 * — the rejection discarded, nothing reported, nothing else going to report it.
 *
 * It matters more there than anywhere else because `updateChartResilient` only
 * throws AFTER its destructive first act: layer 1 deletes the old chart and
 * `deleteShapesById` sweeps the wreckage. If the slide swap and the picture
 * fallback also fail, it rethrows — and the visible outcome was a chart
 * vanishing from the slide with the pane saying nothing, to a user who had
 * pressed no button and had no reason to be looking.
 *
 * ── WHY THIS IS A SOURCE SWEEP ──────────────────────────────────────────────
 * The behaviour needs a live Office host and a host that fails in a specific
 * way, which no unit test here can stage — `a-fake-that-cannot-fail-the-way-
 * the-real-thing-fails` is this repo's name for pretending otherwise. What CAN
 * be checked without a host is the shape: an empty catch on a promise is a
 * decision to discard a failure, and in a file whose whole job is telling the
 * user what the host did, it is never the right one.
 */
const app = readFileSync(fileURLToPath(new URL("../src/taskpane/app.ts", import.meta.url)), "utf8");

/** Source with block comments and strings blanked, so prose cannot match. */
const code = app
  .split(/\/\*[\s\S]*?\*\//)
  .join("")
  .split("\n")
  .map((l) => l.replace(/\/\/.*$/, ""))
  .join("\n");

/** `.catch(() => {})` and its spellings, including a named-but-unused argument. */
const EMPTY_CATCH = /\.catch\(\s*\(\s*\w*\s*\)\s*=>\s*\{\s*\}\s*\)/;

describe("the pane never discards a failure", () => {
  it("has no empty promise catch anywhere in shipped source", () => {
    /**
     * An empty body is the whole tell: a catch that logs, notes, traces or
     * re-throws is a decision; a catch with nothing in it is the absence of one.
     *
     * ── WIDENED FROM `app.ts` TO ALL OF `src/`, 2026-09-29 ────────────────────
     * It read `app.ts` alone, and that is how a source sweep dies quietly. The
     * rule is about discarding a host failure, and the renderer is where host
     * failures COME FROM — so the narrower file was the less important half. The
     * first run over the whole tree found five sites nobody had ever seen:
     * four in `powerpoint.ts` and one in `excel.ts` where the user presses Copy,
     * the browser refuses the clipboard, and nothing at all is said.
     *
     * Two of the five were real and are fixed; three were genuinely fire-and-
     * forget and now say so in the catch body, which is what this message has
     * always asked for and what the narrow version never made anyone do.
     *
     * The sweep has to be a DIRECTORY walk rather than a file read, because a
     * ban pinned to one path is defeated by moving the code — the guard does not
     * go red, it goes quiet. See `test/helpers/module-source.ts`.
     */
    expectSweptSomething("src");
    const empty = sweep(EMPTY_CATCH, "src");
    expect(
      empty,
      `an empty catch discards a host failure. Every write path in the pane reports through ` +
        `note(..., "err") — see guard(). If a rejection really is not worth reporting, say so in the ` +
        `catch body rather than leaving it blank:\n${describeOffenders(empty)}`,
    ).toEqual([]);
  });

  it("still reports the auto-update path by name", () => {
    // The specific instance, pinned so a refactor cannot quietly drop it back
    // to silence. Auto-update has no clicked control to attribute a failure to,
    // which is why it names itself rather than borrowing "Failed: {error}".
    expect(code, "the auto-update timer no longer reports its failures").toMatch(
      /doInsert\(false\)\.catch\([\s\S]{0,400}?Auto-update failed/,
    );
  });

  it("keeps that message in the catalogue, or it renders as a raw key", () => {
    const i18n = readFileSync(fileURLToPath(new URL("../src/taskpane/i18n.ts", import.meta.url)), "utf8");
    expect(i18n).toContain('"Auto-update failed: {error}"');
  });
});
