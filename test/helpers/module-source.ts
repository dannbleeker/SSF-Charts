import { readdirSync, readFileSync } from "fs";
import { join } from "path";
import { expect } from "vitest";

/**
 * Read shipped source as TEXT, across a whole directory rather than one file.
 *
 * WHY THIS EXISTS. Counted 2026-09-29: **47 sites** in this suite call
 * `readFileSync` on a hard-coded `src/` path and regex over the result, across
 * 20 test files — 33 of them on files a refactor was about to move. They are
 * some of the best guards here: they catch things a unit test cannot, like "no
 * call site bypasses the sync counter" or "no empty catch discards a host
 * failure".
 *
 * They also have one failure mode, and it is silent. A **ban** — `not.toMatch`,
 * or a match list asserted empty — keeps passing when the code it guards moves
 * to a file the sweep does not read. The guard does not go red. It goes quiet,
 * and stays quiet forever, which is worse than never having existed because
 * somebody has stopped thinking about the rule.
 *
 * That is not hypothetical here. The empty-catch ban in
 * `no-silent-write-failure.test.ts` reads `src/taskpane/app.ts` alone, and the
 * first time it was pointed at the whole tree it found five violations nobody
 * had ever seen: four in `powerpoint.ts` and one in `excel.ts`, including a
 * clipboard write that fails in silence where the user clicked Copy.
 *
 * So: sweep the DIRECTORY, not the file. A rule worth having in one file is
 * almost always worth having in its neighbours, and a sweep that cannot be
 * narrowed by someone moving code is the only kind that keeps working.
 *
 * `no-native-dialogs.test.ts` already does exactly this, correctly, in private.
 * This is that helper lifted out so the other sweeps can have it too — three
 * tests were carrying near-identical copies of the comment-blanking below.
 */

/** Extensions a source sweep should look at. `.html` carries markup the pane owns. */
const SOURCE_FILE = /\.(ts|html|css)$/;

/**
 * Every source file under `dir`, recursively, in a stable order.
 *
 * Sorted, because `readdirSync` order is not guaranteed across platforms and a
 * sweep that reports offenders in a different order on CI than locally reads as
 * flakiness. This repo has already paid for an unsorted `readdirSync` once —
 * `round-driver.test.ts:2103` bans it as a default for exactly that reason.
 */
export function sourceFiles(dir = "src", match: RegExp = SOURCE_FILE): string[] {
  const out: string[] = [];
  for (const e of readdirSync(dir, { withFileTypes: true }).sort((a, b) => (a.name < b.name ? -1 : 1))) {
    const p = join(dir, e.name);
    if (e.isDirectory()) out.push(...sourceFiles(p, match));
    else if (match.test(e.name)) out.push(p);
  }
  return out;
}

/**
 * A file's code with its PROSE blanked and its line numbers intact.
 *
 * Block comments first, then HTML comments, then line comments — and each is
 * replaced by spaces rather than removed, so `split("\n")[i]` still names the
 * line a human would open the file to.
 *
 * COMMENTS MUST GO, and this repo has been bitten both ways. A colour check
 * once flagged a comment ABOUT a deleted colour; a slide-id sweep matched the
 * word inside a comment I had just written and would have passed against code
 * that no longer called the function. Half this source is prose, so a sweep
 * that reads prose is not a sweep, it is a coin.
 */
export function codeOf(file: string): string {
  const blank = (m: string): string => m.replace(/[^\n]/g, " ");
  return readFileSync(file, "utf8")
    .replace(/<!--[\s\S]*?-->/g, blank)
    .replace(/\/\*[\s\S]*?\*\//g, blank)
    .split("\n")
    .map((l) => l.replace(/\/\/.*$/, ""))
    .join("\n");
}

/** One place a banned pattern was found: the file, the 1-based line, and the line itself. */
export interface Offender {
  file: string;
  line: number;
  text: string;
}

/**
 * Every line under `dir` whose CODE matches `pattern`.
 *
 * The returned array is what a ban asserts empty, and its entries print as
 * `file:line  code` so a failure names the place rather than the rule.
 */
export function sweep(pattern: RegExp, dir = "src", match: RegExp = SOURCE_FILE): Offender[] {
  // `/g` STRIPPED ONCE, HERE. A global regex carries `lastIndex` between `.test`
  // calls, so reusing the caller's would match every OTHER line — a bug that
  // only appears on the second offender, which is exactly when a sweep matters.
  // Callers write `/x/g` out of habit and must not be punished for it.
  const per = new RegExp(pattern.source, pattern.flags.replace("g", ""));
  const found: Offender[] = [];
  for (const file of sourceFiles(dir, match)) {
    codeOf(file)
      .split("\n")
      .forEach((text, i) => {
        if (per.test(text)) found.push({ file, line: i + 1, text: text.trim().slice(0, 120) });
      });
  }
  return found;
}

/** `file:line  code`, one per line — the message a failed ban should carry. */
export function describeOffenders(found: Offender[]): string {
  return found.map((o) => `${o.file}:${o.line}  ${o.text}`).join("\n");
}

/**
 * Assert a sweep actually read something, before believing it found nothing.
 *
 * THE GUARD ON THE GUARD, and the reason this helper exists at all. A widened
 * sweep that resolves to zero files passes every ban in the suite, silently and
 * forever — the same vacuous-pass shape as `web-host.test.ts`'s tag-key loop,
 * which iterates a match list and does nothing at all when the list is empty.
 *
 * `privacy-terms-pages.test.ts` floors at 10 files and `no-control-bytes.test.ts`
 * at 100; both are the same idea. Call this beside every ban that uses `sweep`.
 */
export function expectSweptSomething(dir = "src", floor = 40): string[] {
  const files = sourceFiles(dir);
  expect(
    files.length,
    `the sweep over ${dir}/ found ${files.length} source files — under ${floor} means the walk broke, ` +
      `and a ban asserted over nothing passes forever`,
  ).toBeGreaterThanOrEqual(floor);
  return files;
}

/**
 * The text of the one file that declares `decl`, searched across `dir`.
 *
 * For the OTHER kind of sweep — the ones that slice a named function's body and
 * assert on what is inside it. Those go loudly red when their target moves,
 * which is honest but leaves somebody hunting for the new home. This finds it.
 *
 * Throws on zero matches and on more than one, because both are real answers:
 * zero means the declaration is gone (the guard is looking at nothing, which is
 * the state `office-render.test.ts:4894` already asserts against by hand), and
 * two means the name is ambiguous and a slice would pick arbitrarily.
 */
export function sourceDeclaring(decl: RegExp, dir = "src", match: RegExp = SOURCE_FILE): string {
  const hits = sourceFiles(dir, match).filter((f) => decl.test(readFileSync(f, "utf8")));
  expect(hits, `no file under ${dir}/ declares ${String(decl)} — this guard is looking at nothing`).not.toEqual([]);
  expect(
    hits,
    `${String(decl)} matches ${hits.length} files (${hits.join(", ")}) — a body slice would pick one arbitrarily`,
  ).toHaveLength(1);
  return readFileSync(hits[0], "utf8");
}

// NOTHING ELSE IS EXPORTED HERE ON PURPOSE. An earlier draft of this file also
// carried `sourceUnder` (every file concatenated into one string) and `assertDir`.
// Neither had a caller — `assertDir` did not even have a test — and this repo's
// own lesson is that an unused export is a comment that compiles. A sweep that
// needs one string can have `sourceUnder` back on the day something calls it.
