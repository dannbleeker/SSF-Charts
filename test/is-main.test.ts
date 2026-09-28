import { describe, expect, it } from "vitest";
import { existsSync, mkdtempSync, readdirSync, readFileSync, symlinkSync, writeFileSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { pathToFileURL } from "url";
// @ts-expect-error — a plain .mjs tool with no types.
import { isMain } from "../scripts/is-main.mjs";

/**
 * The predicate that decides whether a script is the CLI or a test's import.
 *
 * Every case here is a real invocation that used to answer wrong. Four scripts
 * carried four different answers to this question and three of them were false
 * on Windows for the same reason — `argv[1].split("/")` does not split a
 * backslashed path — so `node scripts/test-count.mjs results.json` printed
 * nothing and exited 0. That does not read as "the tool did not run". It reads
 * as a pass, from the guard against a silently shrinking suite, from the flake
 * sweep, and from the visual gate.
 *
 * The platform is injected on purpose. CI runs on ubuntu, where the broken form
 * works perfectly; a suite that could only exercise its own platform would go
 * green against the bug these cases exist to catch.
 */
describe("isMain", () => {
  const WIN_URL = "file:///C:/repo/scripts/flaky.mjs";
  const WIN_ARGV = "C:\\repo\\scripts\\flaky.mjs";
  const POSIX_URL = "file:///repo/scripts/flaky.mjs";
  const POSIX_ARGV = "/repo/scripts/flaky.mjs";

  it("says yes when Windows ran the script directly", () => {
    // The regression. `import.meta.url.endsWith(argv[1].split("/").pop())`
    // returns false here, which is how three CLIs came to be no-ops.
    expect(isMain(WIN_URL, WIN_ARGV, "win32")).toBe(true);
  });

  it("says yes when the drive letter's case differs", () => {
    // cmd, PowerShell and git-bash do not agree on the case of the drive
    // letter, and node keeps whichever one it was handed. Same file.
    expect(isMain("file:///c:/repo/scripts/flaky.mjs", "C:\\repo\\scripts\\flaky.mjs", "win32")).toBe(true);
    expect(isMain(WIN_URL, "c:\\REPO\\scripts\\flaky.mjs", "win32")).toBe(true);
  });

  it("says yes when posix ran the script directly", () => {
    expect(isMain(POSIX_URL, POSIX_ARGV, "linux")).toBe(true);
  });

  it("says no when something else was invoked — a test importing the module", () => {
    // The half the old guard got right, and the only half its tests could see:
    // vitest is argv[1] when a spec imports one of these tools, and a CLI that
    // ran on import would call process.exit in the middle of the suite.
    expect(isMain(POSIX_URL, "/repo/node_modules/vitest/vitest.mjs", "linux")).toBe(false);
    expect(isMain(WIN_URL, "C:\\repo\\node_modules\\vitest\\vitest.mjs", "win32")).toBe(false);
  });

  it("says no for a sibling script that merely shares a basename", () => {
    // A basename test — the other tempting shortcut — matches any script called
    // flaky.mjs anywhere on the machine.
    expect(isMain(POSIX_URL, "/somewhere/else/flaky.mjs", "linux")).toBe(false);
  });

  it("says no when there is no argv[1] and no usable URL", () => {
    expect(isMain(POSIX_URL, undefined, "linux")).toBe(false);
    expect(isMain(undefined, POSIX_ARGV, "linux")).toBe(false);
    // A data: or http: module URL is not a path. Not an error — under a bundler
    // there is simply no CLI to be.
    expect(isMain("data:text/javascript,0", POSIX_ARGV, "linux")).toBe(false);
    expect(isMain("https://example.com/flaky.mjs", POSIX_ARGV, "linux")).toBe(false);
  });

  it("resolves a relative argv[1], which is how npm scripts invoke these", () => {
    // `node scripts/flaky.mjs` hands node a relative path.
    const cwd = process.cwd();
    const url = new URL(`file://${cwd.replace(/\\/g, "/").replace(/^([A-Za-z]:)/, "/$1")}/scripts/flaky.mjs`).href;
    expect(isMain(url, "scripts/flaky.mjs", process.platform)).toBe(true);
  });
});

/**
 * And every script uses THAT predicate, not a fifth spelling of its own.
 *
 * Three scripts were fixed in the morning by grepping for the wording they
 * shared — `import.meta.url.endsWith(argv[1].split("/").pop())`. A fourth,
 * `validate-ooxml.mjs`, wrote the same idea a different way,
 * `import.meta.url === \`file://${process.argv[1]}\``, and the grep sailed past
 * it: `file://C:\repo\…` never equals `file:///C:/repo/…`, so the OOXML grammar
 * gate exited 0 without opening the file, for as long as it has existed.
 *
 * Four spellings, four bugs, one of them found only after the "sweep". So this
 * does not look for a wording. It asserts the SHAPE: a script that decides
 * whether it is the CLI must ask `is-main.mjs`.
 */
describe("every tool script asks the same question the same way", () => {
  const dirs = ["scripts", "skill/scripts"];
  const files = dirs.flatMap((d) =>
    readdirSync(d)
      .filter((f) => f.endsWith(".mjs"))
      .map((f) => `${d}/${f}`),
  );

  it("uses isMain(), never a hand-rolled comparison against import.meta.url", () => {
    const offenders: string[] = [];
    let guards = 0;
    for (const file of files) {
      if (file.endsWith("is-main.mjs")) continue; // the definition, and its docstring quotes the bad forms
      const src = readFileSync(file, "utf8");
      // Comments are stripped first: three of these files explain the old bug in
      // prose, and a detector that reads its own postmortem as a violation is
      // one somebody deletes.
      const code = src.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
      /**
       * THE POPULATION IS "ASKS WHETHER IT IS THE CLI", not "mentions this URL".
       *
       * Widened 2026-09-28. The filter admitted only files containing
       * `import.meta.url` — one of the two ways to write the check, and the one
       * the CORRECT form happens to use. A script hand-rolling the decision the
       * OTHER way, reaching for `process.argv[1]` alone, was never in the
       * population, so the sweep that exists to catch hand-rolled CLI guards
       * could not see that hand-rolled CLI guard.
       *
       * A sweep is only as wide as its filter, and the filter was the part
       * nobody re-read. The verdict is unchanged: a correct
       * `isMain(import.meta.url, process.argv[1])` contains both tokens, so
       * every correct caller still passes.
       */
      const withoutUrlBuilders = code.replace(/new URL\([^)]*import\.meta\.url[^)]*\)/g, "");
      const asksIfItIsTheCli = /import\.meta\.url/.test(withoutUrlBuilders) || /process\.argv\[1\]/.test(code);
      if (!asksIfItIsTheCli) continue;
      guards++;
      if (!/isMain\(\s*import\.meta\.url\s*,\s*process\.argv\[1\]\s*\)/.test(code)) {
        offenders.push(file);
      }
    }
    expect(offenders, "these hand-roll the CLI check; three of four such spellings were wrong on Windows").toEqual([]);
    // A scan that matched nothing reports a clean sweep. Several scripts are
    // CLIs, so finding none means the detector stopped working.
    expect(guards, "no entry guard found at all — the scan matched nothing").toBeGreaterThanOrEqual(4);
  });
});

/**
 * The second route to the same silent pass. `import.meta.url` is always the
 * RESOLVED real path; `process.argv[1]` is whatever the shell handed over. Run
 * a script through a symlink — a workspace bin shim, a checkout under a linked
 * parent, `/tmp` on a mac — and the two strings differ, so `main()` never ran
 * and the tool printed nothing and exited 0. That is exactly what this file
 * exists to stop, arriving by a different door.
 */
describe("a script reached through a symlink is still the script that was run", () => {
  const dir = mkdtempSync(join(tmpdir(), "is-main-"));
  const real = join(dir, "tool.mjs");
  const link = join(dir, "link.mjs");
  writeFileSync(real, "export default 1;\n");
  try {
    symlinkSync(real, link);
  } catch {
    /* a filesystem without symlinks — the assertions below are skipped by the guard */
  }
  const canLink = existsSync(link);

  it("answers true when argv[1] is a link to this module", () => {
    if (!canLink) return;
    expect(isMain(pathToFileURL(real).href, link)).toBe(true);
  });

  it("still answers false for a DIFFERENT file", () => {
    // The negative control: resolving links must not turn every invocation into
    // a match.
    const other = join(dir, "other.mjs");
    writeFileSync(other, "export default 2;\n");
    expect(isMain(pathToFileURL(real).href, other)).toBe(false);
    expect(isMain(pathToFileURL(real).href, join(dir, "missing.mjs"))).toBe(false);
  });

  it("keeps answering from the strings alone when the platform is injected", () => {
    // The Windows case has no filesystem to resolve against on an ubuntu
    // runner, and it is the case this file was written for — it must not start
    // depending on realpath.
    expect(isMain("file:///C:/repo/scripts/flaky.mjs", "C:\\repo\\scripts\\flaky.mjs", "win32")).toBe(true);
    expect(isMain("file:///C:/repo/scripts/flaky.mjs", "C:\\repo\\scripts\\other.mjs", "win32")).toBe(false);
  });
});
