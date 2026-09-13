import { defineConfig } from "vitest/config";

/**
 * The suite Stryker runs, which is not the whole suite.
 *
 * Mutation testing copies the project into a sandbox and runs the tests there,
 * and a good third of this suite does not survive that: some read files by a
 * path relative to the working directory (the manifests, the tracker table, the
 * docs gates), some need `dist-lib/` to have been built, and one needs a
 * browser. None of them covers `src/core` — the only thing being mutated — so
 * excluding them costs no signal and is what makes the run possible at all.
 *
 * Keep this list honest. A test excluded because it is inconvenient rather than
 * because it cannot run in a sandbox is signal thrown away, and a mutation score
 * computed over a suite nobody can name is a number rather than a finding.
 */
export default defineConfig({
  test: {
    // THE SAME 20s AS `vitest.config.ts`, AND IT HAS TO BE REPEATED HERE.
    //
    // This file does not extend that one. Stryker is pointed straight at it by
    // `stryker.config.json`'s `vitest.configFile`, so it REPLACES the base
    // config rather than merging with it — and the 20s default added on
    // 2026-09-13 therefore did not reach the mutation run. That day's sweep
    // failed in Stryker's INITIAL DRY RUN with "Test timed out in 5000ms" on
    // `24 categories stays inside every frame`: the exact class of failure the
    // base config had just been changed to stop, in the one place the change
    // could not reach.
    //
    // Contention here is not incidental. Stryker spawns a test-runner process
    // per core — four on the runner — and they share the box, which is the
    // condition that pushes a multi-second test past a 5s ceiling.
    //
    // Stryker's own `timeoutMS: 30000` is a DIFFERENT dial: it bounds how long
    // one mutant may run before being called timed-out. It does not govern
    // vitest's per-test timeout, which is what fired here.
    //
    // AND 20s WAS NOT ENOUGH EITHER — measured, not guessed. With the dry-run
    // cap lifted so the run could finish and say so, it reported:
    //
    //     One or more tests failed in the initial test run:
    //       no top-level option draws outside the chart's frame
    //         title stays inside every frame
    //       Test timed out in 20000ms.
    //
    // i.e. this file IS being read (it said 20000, not 5000) and the test still
    // overran. `test/frame-fit.test.ts` is 143 tests and ~10.5s of test time
    // idle — the heaviest file in the suite — and it hammers precisely the
    // `src/core` code Stryker has instrumented. Per-test coverage analysis
    // rewrites every function in the mutated set, so the multiplier lands
    // hardest exactly here.
    //
    // 120s is proportionate to a ~5-10x instrumentation penalty on a file whose
    // slowest tests are already seconds, not a surrender: nothing in `src/core`
    // is I/O-bound, so a test still running after two minutes is hung rather
    // than slow.
    //
    // WHY NOT EXCLUDE IT, which is what this file did for `agg.test.ts` under
    // the same symptom. Those are two 180k-value stress grids — pathological
    // INPUT, whose extent/spread code other tests also reach, so dropping them
    // costs no signal. `frame-fit.test.ts` is the opposite: 143 ordinary tests
    // that are the main guard on placement and frame fitting. Excluding it
    // would leave those mutants unkilled and quietly lower the score this job
    // exists to produce — buying a green run by deleting the finding.
    testTimeout: 120_000,
    exclude: [
      "**/node_modules/**",
      "**/dist/**",
      "**/.claude/**",
      "**/.stryker-tmp/**",
      // THEY ASSERT ON THE TEXT OF A MUTATED SOURCE FILE, which Stryker
      // rewrites by design. Both read `src/core/layout/*.ts` and match a regex
      // against it; under instrumentation that source begins
      // `// @ts-nocheck\nfunction stryNS_9fa48…` and no longer resembles what
      // was written. This is an incompatibility, not a defect in either test:
      // a source-grepping assertion and a source-rewriting tool cannot both be
      // right about the same bytes.
      //
      // THE TWO FAIL DIFFERENTLY, and the quieter one is the worse.
      // `secondary-axis-ticks` uses `.toMatch`, so it goes RED — that is what
      // broke the dry run on 2026-09-13 ("the sideways strip stopped measuring
      // its labels' width"). `axis-ticks-are-distinct` uses `.not.toMatch`, so
      // instrumentation makes it pass VACUOUSLY: the pattern it forbids is
      // absent because the whole file was rewritten, not because the code is
      // right. It would have reported a clean bill forever and nothing would
      // ever have said otherwise.
      // THE COMPLETE SET, enumerated rather than discovered one red run at a
      // time — the first two were added on 2026-09-13 and the third surfaced
      // eight minutes later in the next dry run, which is this repo's
      // most-repeated defect showing up again. The scan that settles it:
      //
      //   grep -rnE 'readFileSync\([^)]*src/core/[A-Za-z0-9/_-]+\.ts' test/
      //   grep -rnE 'readFileSync\([^)]*src/core/[^)]*\$\{'            test/
      //
      // Six call sites in five files. Three read `src/core/layout/*.ts`, which
      // IS mutated, and are listed here. The other two read `src/core/types.ts`
      // — excluded from `mutate` by the `!src/core/types.ts` negation in
      // stryker.config.json, so never instrumented and safe to leave in.
      // Re-run those two greps if a source-reading test is ever added.
      "test/secondary-axis-ticks.test.ts",
      "test/axis-ticks-are-distinct.test.ts",
      "test/strip-thinning.test.ts",
      // Needs `dist-lib/`, which the sandbox has no build step to produce.
      "test/skill-scripts.test.ts",
      "test/showcase.test.ts",
      "test/visible-charts.test.ts",
      // Reads repo files by a cwd-relative path.
      "test/manifest.test.ts",
      "test/office-js-watch.test.ts",
      "test/skill-docs.test.ts",
      "test/manual.test.ts",
      "test/test-count.test.ts",
      "test/host-contract.test.ts",
      "test/triage.test.ts",
      "test/verify-deck.test.ts",
      "test/ooxml-validate.test.ts",
      "test/crashlog.test.ts",
      // The Office.js and task-pane layers. They exercise `src/render` and
      // `src/taskpane`, not the engine, and they are the slowest third of the
      // suite — mutating core and running these would measure the fake.
      "test/office-render.test.ts",
      "test/web-host.test.ts",
      "test/selftest.test.ts",
      "test/host-probe.test.ts",
      "test/pane-*.test.ts",
      "test/dom-pane.test.ts",
      "test/demo.test.ts",
      "test/excel-*.test.ts",
      "test/templates.test.ts",
      "test/preflight-smoke.test.ts",
      // Two 180k-value stress grids. They pass in seconds normally and do not
      // survive per-test coverage instrumentation, which is a fact about
      // Stryker rather than about them — and the extent/spread code they guard
      // is still mutated, just scored by the other tests that reach it.
      "test/agg.test.ts",
    ],
  },
});
