/**
 * THE TWO TRIAGE TESTS THAT CANNOT RUN UNDER MUTATION TESTING.
 *
 * They are here so that `test/triage.test.ts` can be, which is the whole point:
 * that file holds every pooled reader's tests, and while it was excluded from
 * `vitest.mutation.config.ts` those tests could not kill a single mutant. The
 * pooled readers are now `scripts/round-pools.mjs` and inside `mutate` — see
 * that file's header — and a mutated module whose tests do not run is a green
 * run measuring nothing.
 *
 * ── WHY THESE TWO, AND ONLY THESE TWO ───────────────────────────────────────
 * The exclusion was written as "reads repo files by a cwd-relative path". Read
 * file by file, that turned out to be broader than the actual obstruction:
 * Stryker copies the whole repo into its sandbox and runs with cwd at the
 * sandbox root, so an ordinary read of `docs/WHAT-WE-KNOW.md` works there, and
 * `new URL("../scripts/rounds-gate.mjs", import.meta.url)` was never
 * cwd-relative at all. Four blocks were flagged; two of them were fine.
 *
 * **The stale-detector sweep genuinely cannot run.** It reads EVERY file under
 * `src/` and matches the trace messages `triage.mjs` looks for against the
 * source that emits them. Stryker rewrites `src/core/**` by design — the
 * instrumented file opens `// @ts-nocheck\nfunction stryNS_9fa48…` — so this is
 * the same incompatibility `vitest.mutation.config.ts` records for
 * `secondary-axis-ticks.test.ts`: a source-grepping assertion and a
 * source-rewriting tool cannot both be right about the same bytes. Worse, the
 * failure would be a mutant KILLED for a reason that is nothing to do with the
 * mutant, which inflates a score rather than reporting one.
 *
 * **The CLI invocation test cannot pay for itself.** It spawns
 * `node scripts/triage.mjs` twice with a 120-second budget. A child process does
 * not inherit Stryker's active-mutant global, so it always runs the ORIGINAL
 * code and can never kill anything — while costing two minutes per mutant run.
 *
 * Keep that division. A test moved here because it is slow, rather than because
 * it cannot run, is signal thrown away — the rule
 * `vitest.mutation.config.ts` states about its own list.
 */
import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { spawnSync } from "child_process";
import { mkdtempSync, writeFileSync, rmSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
// @ts-expect-error — plain .mjs tool, no types.
import { FALLBACK_SIGNALS } from "../scripts/round-pools.mjs";

/**
 * The tool's own documented invocation must not print LESS than the degraded one.
 *
 * `reportTrace` was reachable only from the no-deck branch, so
 * `triage.mjs <deck.pptx> <run-log.json>` — the form in the usage line and in
 * CLAUDE.md, the one you use when you actually have a deck — dropped the entry
 * histogram, "phases an error escaped", the problems tally and every
 * `known host bug: office-js#…` annotation, and said nothing about it. A round
 * with a trace and no self-test went further and reported "this log holds no
 * runs and no self-test" over 186 entries, exit 0.
 *
 * Driven through the CLI rather than the exported functions on purpose: what
 * was wrong was the WIRING, and every function involved was already correct.
 */
describe("triage's two invocations", () => {
  const run = (args: string[]) =>
    spawnSync(process.execPath, ["scripts/triage.mjs", ...args], { encoding: "utf8", timeout: 60_000 });

  it("reports the trace whether or not a deck was passed", async () => {
    const log = {
      build: "test-build",
      host: "test-host",
      runs: [],
      trace: {
        summary: {
          steps: [{ scope: "draw", message: "batch committed", n: 2 }],
          problems: [
            {
              text: "PowerPoint did not respond while drawing shapes 1-10 of 24 (45s) | at=drawing the chart's shapes",
              n: 1,
            },
          ],
        },
        entries: [
          { ms: 1, scope: "draw", message: "batch committed" },
          { ms: 2, scope: "draw", message: "batch committed" },
        ],
      },
      selftest: [],
    };
    const dir = mkdtempSync(join(tmpdir(), "pc-triage-"));
    const logPath = join(dir, "round.json");
    writeFileSync(logPath, JSON.stringify(log));

    const withoutDeck = run([logPath]);
    const withDeck = run(["examples/showcase.pptx", logPath]);

    // The trace is a property of the FILE, like the structural faults, so both
    // forms must show it — and in particular the problem line, which is the
    // most locating thing in any round.
    expect(withoutDeck.stdout).toMatch(/TRACE 2 entries/);
    expect(withDeck.stdout, "the deck path dropped the whole trace section").toMatch(/TRACE 2 entries/);
    expect(withDeck.stdout).toMatch(/did not respond while drawing/);
    // And a round that carries a trace is never described as holding nothing.
    expect(withDeck.stdout).not.toMatch(/holds no runs and no self-test/);
    rmSync(dir, { recursive: true, force: true });
  }, 120_000);
});

describe("a detector keyed to a message that no longer exists", () => {
  it("matches only trace messages the source still emits", () => {
    // THE FAILURE MODE THIS PREVENTS IS SILENCE. Every one of these tools finds
    // its evidence by comparing a trace message to a string literal. Rename the
    // message in src/ and the detector does not break — it reports ZERO, every
    // round, forever, and zero is exactly what a healthy round looks like.
    //
    // Nothing in the repo checked this. All the current literals happen to be
    // live, which is the point: the guard is for the rename that has not
    // happened yet, and the archive would carry months of false calm first.
    const tool = readFileSync("scripts/triage.mjs", "utf8");
    const app = readdirSync("src", { recursive: true })
      .map(String)
      .filter((f) => f.endsWith(".ts"))
      .map((f) => readFileSync(`src/${f}`, "utf8"))
      .join("\n");

    const matched = new Set<string>(Object.keys(FALLBACK_SIGNALS));
    // Deliberately backslash-free: a regex written through a heredoc has been
    // corrupted here before, and a mangled one matches nothing while looking fine.
    for (const m of tool.matchAll(/(?:e[.])?message *=== *"([^"]+)"/g)) matched.add(m[1]);
    for (const m of tool.matchAll(/[^A-Za-z]m === "([^"]+)"/g)) matched.add(m[1]);

    // THE REGEXES THEMSELVES ARE INSTRUMENTS. If one stops matching, the loop
    // below passes vacuously and this test becomes decoration — the exact shape
    // it exists to catch. Claim a positive count, not the absence of failures.
    expect(matched.size, "the extractors matched almost nothing — they have stopped working").toBeGreaterThanOrEqual(8);

    for (const message of matched)
      expect(app, `no source file emits "${message}" any more — its detector now reports zero forever`).toContain(
        message,
      );
  });
});

/**
 * THE "READ BY NOTHING" DETECTOR HAS TO KNOW WHICH FILES THE TOOL IS.
 *
 * `rounds-gate.mjs` ends its report with the busiest trace messages that no tool
 * matches on — the section that would have found `poolFallbackRates` and
 * `poolInPlaceUpdates` months before anybody did. It answers by reading the
 * tool's own source and asking whether the message appears in it verbatim.
 *
 * So the answer is only as good as the file list. When the pooled readers moved
 * to `scripts/round-pools.mjs` on 2026-09-29 the list still named two files, and
 * two entries in that section changed places immediately — not because anything
 * about the rounds had changed, but because half the matching code had moved
 * somewhere the concatenation did not look. It was caught by diffing the gate's
 * whole output against the pre-split build over 462 archived rounds. Nothing
 * else would have said so: being wrong here looks exactly like a finding.
 *
 * This is the guard for the next time a file is added. It is a source sweep, so
 * it lives in this file, which Stryker never instruments.
 */
describe("the unread-signal detector knows every file the tool is made of", () => {
  it("reads the source of every tool module that could match a message", () => {
    const gate = readFileSync(new URL("../scripts/rounds-gate.mjs", import.meta.url), "utf8");
    /**
     * THE ARRAY ITSELF, NOT THE FILE.
     *
     * The first draft of this asserted `gate.toContain('"./round-pools.mjs"')`
     * and was vacuous: that string is also the module's own `import` statement
     * at the top of the file, so deleting it from the source list left this
     * green. Proven by deleting it and watching nothing happen. What has to be
     * read is the expression the detector actually passes.
     */
    const list = /const src = \[([^\]]*)\]/.exec(gate)?.[1];
    expect(list, "the unread-signal source list is no longer an array literal this can read").toBeDefined();
    const reads = new Set([...(list ?? "").matchAll(/"\.\/([a-z-]+\.mjs)"/g)].map((m) => m[1]));

    /**
     * Modules that hold no trace-message literal and are deliberately left out.
     * Each one is a judgement, so each one is named: `is-main.mjs` is a
     * twenty-line entry-point check, `host-baseline.mjs` is the probe's answer
     * sheet, `claims.mjs` is a table of claims and `verify-deck.mjs` reads OOXML.
     * None of them matches on a trace message. Add a file here only after
     * checking that, not to make this pass.
     */
    const noMessages = new Set(["is-main.mjs", "host-baseline.mjs", "claims.mjs", "verify-deck.mjs"]);

    // Every `./*.mjs` the two tools import, plus the gate itself. Read out of
    // their imports rather than copied here, so a new tool module is caught the
    // first time it is imported and not the first time somebody remembers.
    const triage = readFileSync(new URL("../scripts/triage.mjs", import.meta.url), "utf8");
    const imported = new Set<string>(["rounds-gate.mjs"]);
    for (const src of [gate, triage])
      for (const m of src.matchAll(/^import[\s\S]*?from "\.\/([a-z-]+\.mjs)";/gm)) imported.add(m[1]);

    const missing = [...imported].filter((f) => !noMessages.has(f) && !reads.has(f));
    expect(
      missing,
      `rounds-gate.mjs matches trace messages against its own source, and ${missing.join(", ")} is not in the ` +
        "list it reads. Every message those modules look for will be reported as read by nothing, which looks " +
        "exactly like a finding.",
    ).toEqual([]);
  });
});
