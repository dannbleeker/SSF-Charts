/**
 * NAME THE TESTS THAT WERE RED IN EVERY RUN.
 *
 * `scripts/flaky.mjs` answers one question — did a test DISAGREE with itself
 * across runs — and by design it is silent when three runs are red in exactly
 * the same way, because that is agreement. The quality sweep then exits 1 with
 * "every run was red in the same way — that is a broken build, not a flake",
 * which is a correct verdict carrying no evidence.
 *
 * IT CARRIED NO EVIDENCE TWICE. The sweep was red on 2026-08-31 and 2026-09-07.
 * Neither log names a single failing test, because the job ran vitest with
 * `--reporter=json --outputFile=...` and nothing else: every failure went into
 * a file in /tmp on a runner that is then destroyed, and the console got
 * "run N: RED".
 *
 * THE FIRST RUN AFTER THIS EXISTED NAMED IT IN ONE LINE, and it was not what
 * anyone had guessed. `backlog-health-table.test.ts > reproduces every row at
 * the round it claims`, failing 3 of 3 on `fatal: ambiguous argument
 * '6dfaa4b..HEAD'` — the sweep's checkout was shallow and that workflow had
 * never been given `fetch-depth: 0`. Deterministic, nothing to do with load.
 *
 * That is worth dwelling on, because the day was first spent on a different
 * answer: reproducing the sweep's CPU load by hand on a 4-core box turned up
 * seven timeouts, and they looked exactly like a cause. They were a real
 * fragility and they were not the outage. A shallow-clone failure CANNOT be
 * reproduced locally — a working clone has the history — so the local
 * experiment could only ever find something else and offer it as the answer.
 *
 * WHAT THIS FILE CANNOT TELL YOU, AND WHY — measured, not assumed.
 *
 * Vitest 4's json reporter does NOT carry the failure reason. Every
 * `failureMessages[0]` reads:
 *
 *     Error: STACK_TRACE_ERROR
 *         at task (.../@vitest/runner/dist/chunk-artifact.js:1784:27)
 *         ...
 *         at C:/devtools/SSF-Charts/test/pane-state.test.ts:257:3
 *
 * The literal string `STACK_TRACE_ERROR` is a placeholder; "Test timed out in
 * 5000ms" appears ONLY in the default reporter's console output. A first draft
 * of this script printed that first line as the reason and tested it against
 * /timed out/ — which would have printed a placeholder and never once matched.
 * It was caught by generating a real red report instead of trusting the shape.
 *
 * So this gives the two things the json genuinely holds — WHICH tests, and the
 * file:line inside the repo that the stack points at — and the job runs
 * `--reporter=default` alongside for the reason, and uploads the json for the
 * rest. Do not reintroduce a reason here without re-reading a real report.
 *
 * Prints to STDERR, because its caller is already failing and this is part of
 * that failure's explanation.
 *
 * Usage:  node scripts/red-in-every-run.mjs /tmp/run-1.json /tmp/run-2.json ...
 * Exit:   always 0 — it explains a failure, it does not add one.
 */

import { readFileSync } from "node:fs";
import { isMain } from "./is-main.mjs";

/**
 * The first stack frame that lands in this repo rather than in a dependency.
 * Every vitest failure stack opens with several `@vitest/runner` frames, so the
 * naive "first line after the message" is always node_modules and never the
 * test. Returns "" when no frame qualifies rather than guessing.
 */
export function repoFrame(stack) {
  for (const line of String(stack ?? "").split("\n")) {
    if (line.includes("node_modules")) continue;
    const m = line.match(/([A-Za-z]:\/|\/)?[\w./\\-]*\/((?:test|src|scripts)\/[\w./\\-]+:\d+:\d+)/);
    if (m) return m[2];
  }
  return "";
}

/**
 * A test is reported when it failed in EVERY run that executed it — which is
 * what "red in every run" means. A test that failed once and passed twice is a
 * flake and belongs to `flaky.mjs`, not here; printing it would blur the one
 * distinction this file exists to keep sharp.
 */
export function redInEveryRun(reports) {
  const failed = new Map();
  const ran = new Map();
  for (const report of reports) {
    for (const suite of report?.testResults ?? []) {
      for (const t of suite?.assertionResults ?? []) {
        const key = `${suite.name ?? "?"} > ${t.fullName ?? t.title ?? "?"}`;
        ran.set(key, (ran.get(key) ?? 0) + 1);
        if (t.status === "failed") {
          const prev = failed.get(key) ?? { count: 0, where: "" };
          failed.set(key, {
            count: prev.count + 1,
            where: prev.where || repoFrame((t.failureMessages ?? [])[0]),
          });
        }
      }
    }
  }
  return [...failed.entries()]
    .filter(([key, v]) => v.count === ran.get(key))
    .map(([name, v]) => ({ name, where: v.where }));
}

function main(argv) {
  const reports = [];
  for (const path of argv) {
    try {
      reports.push(JSON.parse(readFileSync(path, "utf8")));
    } catch {
      // A missing or truncated report is not this script's failure to report:
      // the run that should have written it already failed loudly.
    }
  }
  const red = redInEveryRun(reports);
  if (!red.length) {
    console.error("  (no per-test detail in the json reports)");
    return;
  }
  console.error(`  ${red.length} test(s) red in EVERY run:`);
  for (const r of red) console.error(`    ${r.name}${r.where ? `\n        ${r.where}` : ""}`);
  console.error("");
  console.error("  The json reporter does not carry WHY — vitest writes a STACK_TRACE_ERROR");
  console.error("  placeholder. Scroll up to the default reporter's output in the same log, or");
  console.error("  download the `flake-run-reports` artifact. If they are all timeouts, this is");
  console.error("  a suite starved of CPU on a saturated runner, not a broken product.");
}

if (isMain(import.meta.url, process.argv[1])) {
  main(process.argv.slice(2));
}
