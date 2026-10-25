#!/usr/bin/env node
/**
 * The fast local gate — CI's own commands, read from package.json, and an
 * honest list of what it does NOT run.
 *
 * WHY THIS EXISTS. On 2026-08-23 a push went out reported as "gate green:
 * typecheck 0, lint 0, format clean, 3310 tests" and CI failed on
 * `prettier --check .`. The local check had been `prettier --check src test
 * scripts docs` — a hand-assembled subset that never looked at `crashes/`,
 * where the driver had just archived a dump from a round whose browser died.
 * The claim was "format clean"; the check was "format clean in four
 * directories". A gate whose scope does not match the claim made about it is
 * the defect this repo spends most of its comments on, and it had reached the
 * thing that verifies everything else.
 *
 * WHY IT READS package.json INSTEAD OF SHELLING OUT TO `npm run`. Two reasons,
 * and the second is the interesting one.
 *
 * The point of the gate is that its commands cannot drift from CI's. Hardcoding
 * `prettier --check .` here would reintroduce exactly the drift it exists to
 * stop, so the command TEXT is read from the same `scripts` block CI invokes by
 * name. Change `format:check` in package.json and this changes with it.
 *
 * And on this box it could not shell out anyway: AppLocker blocks an npm script
 * that nests `npm run` inside itself — "This program is blocked by group
 * policy" — which is what the first version of this file did, and it died on
 * its first real use. Running the script BODIES with `node_modules/.bin` on
 * PATH is what `npm run` does, minus the nesting.
 *
 * WHAT IT DELIBERATELY SKIPS, printed on success so it can never be mistaken
 * for the whole of CI: coverage thresholds, the suite-shrink floor, the vite
 * build, the manifest check, the skill build, deck currency, verification and
 * OOXML validation, the visual gate that rasterises every sample chart in a
 * browser, and Microsoft's manifest validator (a job of its own). Skipping them
 * is a choice about speed; hiding that they were skipped would be a lie about
 * scope, which is the thing this file exists to stop.
 *
 * THE LIST IS CHECKED AGAINST `.github/workflows/ci.yml` (`test/gate.test.ts`).
 * It named "the Playwright e2e job", which no workflow has ever had, and left
 * out three steps CI does run: two of them in the required `test` job, so a
 * pusher reading the list could see a green gate and a red required check
 * (r3-78). A list of what CI runs that is written from memory drifts from CI.
 */
import { spawnSync } from "child_process";
import { readFileSync } from "fs";
import path from "path";
import { isMain } from "./is-main.mjs";

/** The fast subset of CI's `test` job, in CI's order. Names, not commands. */
export const GATE_STEPS = ["lint", "format:check", "typecheck", "test"];

/** CI steps this gate does not run, so success can never read as "CI will pass". */
export const NOT_COVERED = [
  "coverage (CI enforces thresholds; `npm test` does not)",
  "npx vite build",
  "build:manifest -- --check",
  "skill",
  "verify-deck / validate-ooxml on examples/showcase.pptx",
  // NAMED SEPARATELY from verify-deck, because it is a different question and
  // missing it cost a red CI on 2026-08-28. verify-deck asks whether the
  // committed deck is SOUND; this asks whether it is CURRENT — CI rebuilds it
  // and diffs the slide parts against the committed file. Any layout change
  // rewrites it, so any layout change has to `npm run showcase` and commit the
  // result. A green local gate says nothing about it.
  "`npm run showcase` — CI fails if examples/showcase.pptx is stale after a layout change",
  // THE SUITE-SHRINK FLOOR. It reads the coverage run's results, so it cannot
  // run here at all, and the failure it exists for is invisible to this gate:
  // delete or `.skip` tests and everything above stays green while CI goes red.
  "scripts/test-count.mjs — CI fails if the suite has shrunk below its recorded floor",
  // THE VISUAL GATE, plus the browser install it needs. It is in the required
  // `test` job. This entry used to read "the Playwright e2e job"; the only
  // Playwright in CI is the browser this step drives.
  "scripts/visible-charts.mjs — CI rasterises every sample chart in a browser and fails on one with no ink",
  // A separate job, not a required check (the validator is a call to a Microsoft
  // service, and an outage there must not block a merge), so a red here is a
  // thing to look at rather than a wall. Still run by CI, still not run here.
  "office-addin-manifest validate — the separate `manifest` job (not a required check)",
];

/** The command CI runs for a step, straight from package.json. */
export function commandFor(step, pkgJson = readFileSync("package.json", "utf8")) {
  const scripts = JSON.parse(pkgJson).scripts ?? {};
  const cmd = scripts[step];
  if (!cmd) throw new Error('package.json has no script "' + step + '" — the gate names a step CI does not have');
  return cmd;
}

/** `node_modules/.bin` ahead of PATH, which is the part of `npm run` still needed. */
function binPath(env = process.env) {
  const bin = path.resolve("node_modules/.bin");
  return { ...env, PATH: bin + path.delimiter + (env.PATH ?? "") };
}

export function runGate({
  run = (cmd) => spawnSync(cmd, { stdio: "inherit", shell: true, env: binPath() }).status,
  log = console.log,
  read = () => readFileSync("package.json", "utf8"),
} = {}) {
  for (const step of GATE_STEPS) {
    const cmd = commandFor(step, read());
    log("\n=== " + step + ": " + cmd + " ===");
    const status = run(cmd, step);
    if (status !== 0) {
      log("\nFAILED at `" + step + "` — stopping here rather than reporting a partial pass.");
      return { ok: false, failedAt: step };
    }
  }
  log("\nGate green: " + GATE_STEPS.join(", "));
  log("NOT run by this gate, and CI does run them:");
  for (const s of NOT_COVERED) log("  - " + s);
  return { ok: true, failedAt: null };
}

if (isMain(import.meta.url, process.argv[1])) process.exit(runGate().ok ? 0 : 1);
