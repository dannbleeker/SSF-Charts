import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";

/**
 * THE ONE PROPERTY THAT KEEPS `scripts/round-pools.mjs` MUTATABLE.
 *
 * Every function in that module takes parsed round logs and returns numbers. It
 * imports nothing, reads no file, spawns nothing and never asks the clock. That
 * is not tidiness — it is the whole reason the module exists, and the reason
 * `test/triage.test.ts` could be taken OUT of `vitest.mutation.config.ts`'s
 * exclusion list on 2026-09-29.
 *
 * ── WHY A TEST AND NOT A COMMENT ────────────────────────────────────────────
 * One `readFileSync` in there and the tests that reach it stop being runnable in
 * Stryker's sandbox, which puts the pooled readers back outside `mutate` —
 * SILENTLY, because everything would still pass here and the weekly job would
 * still be green. `stryker.config.json` records what that cost the last time:
 * four tests written for these readers passed against the code they were written
 * to catch, and every one was found by stashing the source by hand.
 *
 * So the absence is asserted rather than trusted.
 *
 * ── AND THIS FILE ITSELF CANNOT RUN UNDER MUTATION ──────────────────────────
 * It greps a source Stryker rewrites. The instrumentation injected into
 * `round-pools.mjs` reads `process.env` to find the active mutant, so the first
 * scoped run died in its dry run reporting that the module "reaches outside
 * itself via process" — true of the rewritten bytes, false of the source. The
 * guard was right about what it was looking at and wrong about which file that
 * was. It is excluded in `vitest.mutation.config.ts` for the same reason
 * `secondary-axis-ticks.test.ts` is, and it does its job in `npm test`, where
 * the source is the source.
 *
 * ── THE READERS' OWN TESTS ARE NOT HERE ─────────────────────────────────────
 * They are in `test/triage.test.ts`, where they were written, and they stayed
 * there deliberately. Moving them would have meant sorting forty interleaved
 * `it`s out of one 1,300-line describe that tests pooled readers and the tool
 * around them from the same brace — and a botched test split is a worse outcome
 * than a file whose name is broader than one module. What mattered was that
 * those tests RUN under mutation, and taking two blocks out of that file
 * achieved it; see `test/triage-repo.test.ts`.
 */
describe("the pooled readers stay pure, which is what makes them mutatable", () => {
  const src = readFileSync(new URL("../scripts/round-pools.mjs", import.meta.url), "utf8");
  /**
   * Comments stripped first. Half this repo is prose and the module's own header
   * discusses `readFileSync` and `traceSource` by name — a detector that reads
   * its own postmortem as a violation is one somebody deletes. The same strip,
   * for the same reason, is in `test/helpers/module-source.ts`,
   * `test/pane-state.test.ts` and `test/manifest.test.ts`.
   */
  const code = src
    .split(/\/\*[\s\S]*?\*\//)
    .join("")
    .split("\n")
    .map((l) => l.replace(/\/\/.*$/, ""))
    .join("\n");

  it("imports nothing at all, so it cannot be in a cycle or pull in the filesystem", () => {
    const imports = [...code.matchAll(/^\s*import\s/gm)];
    expect(
      imports.length,
      "round-pools.mjs has grown an import. A module with none cannot be part of a cycle and " +
        "cannot reach disk by accident — check what was added before widening this.",
    ).toBe(0);
    expect(code, "a `require` is an import wearing a hat").not.toMatch(/\brequire\s*\(/);
    expect(code, "a dynamic import reaches the same places a static one does").not.toMatch(/\bimport\s*\(/);
  });

  it("touches no filesystem, no clock, no process and no child process", () => {
    // Named individually rather than by a `node:` prefix sweep, because the
    // module has no imports: an unimported `readFileSync` would be a runtime
    // error, and what this actually guards against is somebody ADDING the import
    // and the call together.
    const banned = [
      "readFileSync",
      "writeFileSync",
      "readdirSync",
      "existsSync",
      "statSync",
      "mkdirSync",
      "rmSync",
      "spawnSync",
      "execSync",
      "execFileSync",
      "process",
      "__dirname",
      "import.meta",
      "Date.now",
      "new Date",
      "Math.random",
    ];
    const found = banned.filter((b) => code.includes(b));
    expect(
      found,
      `round-pools.mjs reaches outside itself via ${found.join(", ")}. That makes its tests unrunnable in ` +
        "Stryker's sandbox, which silently puts the pooled readers back outside `mutate`.",
    ).toEqual([]);
  });

  it("is the module stryker.config.json and mutation-scope.mjs both name", async () => {
    // The three places that have to agree, checked from the side that would
    // otherwise only be checked by a weekly job nobody watches finish.
    const stryker = JSON.parse(readFileSync(new URL("../stryker.config.json", import.meta.url), "utf8"));
    expect(stryker.mutate, "round-pools.mjs is not in stryker.config.json's `mutate`").toContain(
      "scripts/round-pools.mjs",
    );
    // @ts-expect-error — plain .mjs tool, no types.
    const { MUTATED, MUTATED_ROOTS } = await import("../scripts/mutation-scope.mjs");
    expect(
      MUTATED.test("scripts/round-pools.mjs"),
      "the weekly job's `--mutate` OVERRIDES the config, and its scope comes from here",
    ).toBe(true);
    expect(MUTATED_ROOTS, "and the git diff that feeds that scope would never show this file").toContain(
      "scripts/round-pools.mjs",
    );
  });

  it("exports every pooled reader that left triage.mjs, and triage.mjs exports none", async () => {
    // A re-export left behind would mean two names for one function and a
    // `mutate` entry that only covers one of them.
    // @ts-expect-error — plain .mjs tool, no types.
    const pools = await import("../scripts/round-pools.mjs");
    const names = Object.keys(pools).filter((n) => n.startsWith("pool"));
    expect(names.length, "the pooled readers are not all here").toBe(33);
    // @ts-expect-error — as above.
    const triage = await import("../scripts/triage.mjs");
    expect(
      Object.keys(triage).filter((n) => n.startsWith("pool")),
      "triage.mjs still exports a pooled reader, so there are two names for one function",
    ).toEqual([]);
  });
});
