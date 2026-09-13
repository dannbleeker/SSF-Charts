import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Exclude in-repo agent worktrees (.claude/worktrees/**) so vitest doesn't
    // discover duplicate test files copied into each worktree.
    // `.stryker-tmp` is a full copy of the repo, tests and all. A leftover
    // sandbox would make `npm test` run every file twice — once real, once from
    // a stale copy — and report failures against source nobody edited.
    exclude: ["**/node_modules/**", "**/dist/**", "**/.claude/**", "**/.stryker-tmp/**"],
    // See test/setup.ts: a fake host has no lag to settle, so no suite pays the
    // real host's settle delay.
    setupFiles: ["test/setup.ts"],
    // 20s rather than vitest's 5s default, and this REVERSES A RULE this repo
    // wrote down, so here is why.
    //
    // The rule, stated in `test/rounds.test.ts` when one test first outgrew the
    // default: name the timeout on the slow test, never raise it globally,
    // because a suite-wide bump hides the next test that is slow for a REAL
    // reason. That was right when it was one test.
    //
    // WHAT CHANGED. `quality-sweep.yml` runs the suite three times under
    // deliberate CPU load and went red on 2026-08-31 and 2026-09-07 naming
    // nothing — the job wrote failures to a json in /tmp on a runner that is
    // then destroyed. Reproduced by hand on 2026-09-13 on a 4-core box, the
    // same core count as the runner: every failure was a plain TIMEOUT, not one
    // assertion failure. Nothing about the product was broken.
    //
    // AND THE AFFECTED SET IS NOT FIXED. One run produced seven tests, the next
    // six, and a third pulled in `backlog-health-table` which neither had. They
    // are simply the tests whose idle cost sits closest to 5s — measured
    // 2026-09-13, eleven of them are over 700ms and the archive walkers grow by
    // one round file per round, forever. Patching the ones a given run happened
    // to catch is fitting to a sample; the next run picks a different subset.
    //
    // So the premise of the old rule — "a handful, individually nameable" — is
    // gone, and its cost is now a permanently red watchdog, which is a watchdog
    // nobody reads. 20s is ~8x the slowest ordinary test here (2.2s) so a real
    // regression of several times still trips, while 2x contention no longer
    // does. The one test that genuinely needs more says so at its own site.
    testTimeout: 20_000,
    coverage: {
      provider: "v8",
      // src plus the one skill script that is pure and importable: pptx-paint.mjs
      // (the paint/node helpers of the headless pptx renderer). render-pptx.mjs
      // itself is a subprocess CLI — v8 can't measure it — so only its extracted,
      // in-process-tested core is gated here.
      include: ["src/**", "skill/scripts/pptx-paint.mjs"],
      // Type-only modules and the browser/Office-only entry files. app.ts is
      // NOT excluded any more: the pane-state suite boots and drives the real
      // module, so the file most likely to regress is measured rather than
      // hidden — a floor below its current level guards against silent drops.
      //
      // The *.html entries are the pane/excel HTML shells. v8 lists them as
      // uncovered source and then tries to PARSE them as JS, dumping a
      // RolldownError ("Unexpected JSX expression") stack per run that buries
      // real failures in the coverage log — exclude them so the log stays clean.
      exclude: ["src/demo/**", "src/core/types.ts", "src/index.ts", "**/*.html"],
      reporter: ["text", "html"],
      thresholds: {
        // A global backstop: glob-keyed thresholds only gate files they match, so
        // without this a NEW directory outside every glob below would be measured
        // but never asserted. Set well under the current aggregate.
        statements: 85,
        branches: 75,
        // The pure engine is the product — hold it to a high bar.
        "src/core/**": { statements: 95, branches: 88 },
        // 80 is currently 0.06 points under the measured 80.06, which is a
        // trap rather than a guard: the next PR to add an uncovered branch
        // fails here for a reason unrelated to its change, and the obvious fix
        // under deadline is to lower the number.
        //
        // Do not lower it. The slack is all in ONE file — `host-probe.ts` sits
        // at 65% branches while `powerpoint.ts` is at 80.16 — and it is there
        // for a legible reason: every probe carries a defensive branch for a
        // host that answers oddly, and reaching each one needs its own bespoke
        // fault. Writing twenty tests to hit twenty catch blocks would be
        // coverage-chasing, which `test/README.md` warns against by name. Buy
        // headroom with a test worth having or leave it; do not buy it by
        // moving the floor.
        "src/render/**": { statements: 85, branches: 80 },
        // The task pane is driven end-to-end (pane-state.test.ts + the host
        // command handlers in pane-host-actions.test.ts, which drive Insert /
        // Same-scale / Load against a mocked host) — a regression floor, not the
        // engine's bar. Raised from 75/58 once the Office-backed command layer
        // was covered; still a few points under today's numbers for headroom.
        "src/taskpane/**": { statements: 80, branches: 65 },
        // The Excel data bridge — a regression floor under today's numbers (it was
        // live code changed by #141 but matched no glob, so it was ungated).
        "src/excel/**": { statements: 80, branches: 70 },
        // The headless pptx renderer's pure paint/node core (pptx-paint.mjs),
        // unit-tested in-process by pptx-paint.test.ts. A floor under today's
        // 100/88 — the CLI wrapper stays out of coverage as a subprocess.
        "skill/scripts/**": { statements: 95, branches: 80 },
      },
    },
  },
});
