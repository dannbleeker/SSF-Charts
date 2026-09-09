import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "fs";
import { spawnSync } from "child_process";
// @ts-expect-error — plain .mjs tool, no types.
import { poolEveryDraw } from "../scripts/triage.mjs";
// Its own line: the grouped-import + `@ts-expect-error` trap is documented at
// the top of `triage.test.ts` and has now bitten twice.
// @ts-expect-error — as above.
import { poolRasteriseArms } from "../scripts/triage.mjs";
// Its own line, same trap as every other single import in this file.
// @ts-expect-error — as above.
import { loadRounds } from "../scripts/rounds-gate.mjs";
// And its own line again, for the same reason.
// @ts-expect-error — as above.
import { countCrashReports, coverageOf, COVERAGE_FLOOR } from "../scripts/rounds-gate.mjs";

/**
 * The archive is evidence, so it has to stay readable and honestly labelled.
 *
 * Not a snapshot of the findings — those live in the round files themselves and
 * in the commits that acted on them. This checks the two things that would make
 * the archive quietly wrong: a file that no longer parses as a round, and a file
 * whose name claims a build it does not carry. The second is not hypothetical.
 * Round 27 ran on `fef1c2a` while `162f80a` was the last commit merged, and the
 * first attempt at this directory named it `027-162f80a.json`. An archive that
 * misattributes a round to a build is worse than no archive: every later
 * comparison inherits the error and nothing on the machine disagrees.
 */
describe("the round archive", () => {
  const dir = new URL("../rounds/", import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1");
  // `NNN-<build>.json` only, exactly as `triage.mjs` expands the directory.
  // "Every .json in rounds/" is the obvious spelling and it is wrong, because
  // `predictions.json` lives here too — it bit triage's directory expansion and
  // then bit this test within the hour, both times by counting the ledger as a
  // round. One shape, two places; if a third reader of this directory appears it
  // needs the same filter.
  const files = readdirSync(dir).filter((f) => /^\d{3}-.*\.json$/.test(f));

  it("has rounds in it", () => {
    expect(files.length, "the archive is empty — see rounds/README.md").toBeGreaterThan(0);
  });

  it("names every file for the build that round actually reported", () => {
    for (const f of files) {
      const round = JSON.parse(readFileSync(dir + f, "utf8"));
      const build = String(round.build ?? "").split(" ")[0];
      expect(build, `${f} carries no build stamp`).toMatch(/^[0-9a-f]{7}$/);
      expect(f, `${f} is named for a build it does not carry (${build})`).toContain(build);
      expect(f, `${f} does not start with a round number`).toMatch(/^\d{3}-/);
    }
  });

  it("keeps every round readable by the tools that pool them", () => {
    const logs = files.map((f) => JSON.parse(readFileSync(dir + f, "utf8")));
    for (const [i, log] of logs.entries()) {
      expect(Array.isArray(log?.trace?.entries), `${files[i]} has no trace entries`).toBe(true);
      expect(Array.isArray(log?.hostAnswers?.answers), `${files[i]} has no answer sheet`).toBe(true);
    }
    // The whole point of keeping them: the pooled view must see more than one
    // round's worth. A single round contributes 4 arm draws and ~40 total, so
    // anything at or below those numbers means the pooling silently read one
    // file — which is the state this directory exists to end.
    const arms = poolRasteriseArms(logs);
    const every = poolEveryDraw(logs);
    const armDraws = Object.values(arms.arms as Record<string, { ok: number; stall: number }>).reduce(
      (n, a) => n + a.ok + a.stall,
      0,
    );
    const allDraws = Object.values(every.after as Record<string, { ok: number; stall: number }>).reduce(
      (n, a) => n + a.ok + a.stall,
      0,
    );
    expect(arms.rounds, "pooling read fewer rounds than the archive holds").toBe(files.length);
    expect(armDraws, "the counterbalanced arms pooled no further than one round").toBeGreaterThan(4);
    expect(allDraws, "every-draw pooling saw no more than a single round").toBeGreaterThan(40);
  });

  it("reports on the NEWEST round when handed the directory", () => {
    // `npm run rounds` is the one command a loop glances at between rounds, and
    // it expanded the directory sorted then read `[0]` — the OLDEST round. So
    // the deck evidence and self-test at the top came from two days before the
    // grid underneath, with nothing saying so. Nineteen rounds went by without
    // it being noticed, because the grid was right.
    //
    // Asserted through the CLI rather than a unit, because the defect lived in
    // argument handling and a unit test of the reporter would have passed
    // throughout.
    const sorted = [...files].sort();
    const newestBuild = /^\d{3}-(.*)\.json$/.exec(sorted[sorted.length - 1])![1];
    const out = spawnSync(process.execPath, ["scripts/triage.mjs", "rounds"], { encoding: "utf8" }).stdout ?? "";
    const header = out.split("\n").find((l) => l.trimStart().startsWith("build ")) ?? "";
    expect(header, "reported some round other than the newest").toContain(newestBuild);
    // An explicitly named file still means THAT file: two rounds named on the
    // command line means the first of them, and only the directory case flips.
    const firstBuild = /^\d{3}-(.*)\.json$/.exec(sorted[0])![1];
    const named =
      spawnSync(process.execPath, ["scripts/triage.mjs", dir + sorted[0], dir + sorted[sorted.length - 1]], {
        encoding: "utf8",
      }).stdout ?? "";
    expect(named.split("\n").find((l) => l.trimStart().startsWith("build ")) ?? "").toContain(firstBuild);
    // A BUDGET THAT SHRINKS EVERY ROUND. This test spawns triage TWICE, and the
    // directory case parses the WHOLE archive — 116 rounds on 2026-08-20 and one
    // more after every run. It took ~6.5s against vitest's 5s default and began
    // failing in the full suite while passing alone, which reads as a code
    // regression and is not one: the test is doing exactly what it should, and
    // its allowance was set when the archive was a fraction of this size.
    //
    // Named explicitly rather than raised globally, because a suite-wide bump
    // would hide the next test that is slow for a REAL reason.
  }, 30_000);

  it("does not carry the slide images, which are half the bytes", () => {
    for (const f of files) {
      const raw = readFileSync(dir + f, "utf8");
      // A base64 slide picture is thousands of characters. Anything that long
      // and base64-shaped means an unstripped round got committed.
      expect(/"[A-Za-z0-9+/=]{2000,}"/.test(raw), `${f} still carries an embedded image — strip it`).toBe(false);
    }
  });
});

/**
 * The gate reads this directory, and `archive` writes straight to the final
 * path rather than writing-then-renaming — so an interrupted write leaves a
 * truncated round behind, and the gate has to survive meeting one.
 */
describe("reading an archive with a bad file in it", () => {
  const list = () => ["080-aaaaaaa.json", "081-bbbbbbb.json", "082-ccccccc.json"];
  const read = (p: string) => (p.includes("081") ? '{"build":"bbbbbbb","selftest":[' : `{"build":"x","selftest":[]}`);

  it("reads past a round that will not parse instead of dying on it", () => {
    // Unguarded this threw a SyntaxError, node exited 1, and `cycle.mjs` reads
    // ANY non-zero gate as a regression — so a corrupt file stopped the night
    // reporting a fall that never happened.
    const rounds = loadRounds("rounds", list as never, read as never);
    expect(rounds).toHaveLength(2);
  });

  it("names the file it could not read, because a skipped round hides its own regression", () => {
    // Silently reading past it would be the other half of the same mistake: a
    // round left out of the comparison is a round whose fall cannot be seen.
    const rounds = loadRounds("rounds", list as never, read as never) as unknown as { unreadable: string[] };
    expect(rounds.unreadable).toEqual(["081-bbbbbbb.json"]);
  });

  it("counts the crashes that left no round file, so a rate is not read as complete", () => {
    /**
     * A round the driver never recovered archives NOTHING. `round.mjs` states
     * the mechanism outright: "A crashed round archives nothing: it never
     * reaches the download button." So `rounds/` holds only the crashes that
     * were survived, and every crash rate computed from it is a FLOOR.
     *
     * Measured 2026-09-01: 77 reports in `crashes/` against 46 crash events in
     * archived `driverRun.recovered` — 40% of the crashes this project has seen
     * left no round file to be counted in. That caveat belongs beside the
     * 4:3-versus-16:9 rate, which is computed from exactly that denominator.
     */

    const dir = ["2026-08-29-a.md", "2026-08-29-b.md", "notes.txt", "README"];
    expect(countCrashReports("crashes", (() => dir) as never)).toBe(2);
    // A missing directory is not a claim of zero crashes, but there is nothing
    // to report from it either — and it must not take the gate down.
    expect(
      countCrashReports("nope", (() => {
        throw new Error("ENOENT");
      }) as never),
      "an absent crashes/ directory took the gate down with it",
    ).toBe(0);
  });
});

describe("whether the shipped bundle could possibly be responsible", () => {
  // Round 273 is why this exists. `two slides claiming one slot` stopped passing
  // and the cycle halted on its one fatal check, correctly — but the build under
  // judgement had changed `scripts/round.mjs`, a test and an archive file, and
  // NOTHING under `src/`. It was the same bundle four earlier rounds had passed
  // on, so the product could not have caused it. Establishing that took a git
  // diff anyone could have run and nobody was prompted to.
  it("says the bundle is unchanged when only non-shipping files moved", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { bundleChanged } = await import("../scripts/rounds-gate.mjs");
    // `git diff --name-only … -- src/` answering empty is the whole signal.
    const noSrc = () => "";
    expect(bundleChanged("e97699e · 2026-08-26 21:28Z", "10f8c60 · 2026-08-27 00:10Z", noSrc)).toBe(false);
  });

  it("says it changed when a line that RUNS moved", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { bundleChanged } = await import("../scripts/rounds-gate.mjs");
    // Two calls now — the name list, then the diff. The stub answers by which
    // one it was handed, because the second is what decides.
    const withCode = (_cmd: string, args: string[]) =>
      args.includes("--name-only")
        ? "src/render/powerpoint.ts\n"
        : "--- a/x\n+++ b/x\n-  const a = 1;\n+  const a = 2;\n";
    expect(bundleChanged("aaaaaaa · x", "bbbbbbb · y", withCode)).toBe(true);
  });

  it("says COMMENTS when src/ moved but no line that runs did", async () => {
    /**
     * Round 440's 4:3 leg failed and the gate could not say "read this as the
     * host", because 27 lines of COMMENT had been added to `app.ts` and
     * `powerpoint.ts` an hour earlier. `--name-only` sees a file, so the reader
     * was sent to run by hand the diff this check exists to spare them.
     *
     * A THIRD VALUE, never `false`. The caller prints it as a fact rather than
     * as the "the product did not change" verdict — saying that wrongly excuses
     * a regression, while saying "changed" wrongly costs one diff.
     */
    // @ts-expect-error — plain .mjs tool, no types.
    const { bundleChanged } = await import("../scripts/rounds-gate.mjs");
    const commentsOnly = (_cmd: string, args: string[]) =>
      args.includes("--name-only")
        ? "src/taskpane/app.ts\n"
        : "--- a/x\n+++ b/x\n+  // a note about why\n+   * and its continuation\n+\n";
    expect(bundleChanged("aaaaaaa · x", "bbbbbbb · y", commentsOnly)).toBe("comments");

    // AND THE ASYMMETRY IS PINNED. One real line among the comments makes it a
    // change: anything this cannot plainly see as a comment counts as code,
    // which is the safe direction.
    const mixed = (_cmd: string, args: string[]) =>
      args.includes("--name-only")
        ? "src/taskpane/app.ts\n"
        : '--- a/x\n+++ b/x\n+  // a note\n+  const s = "// not a comment";\n';
    expect(bundleChanged("aaaaaaa · x", "bbbbbbb · y", mixed), "a code line hid behind a comment").toBe(true);
  });

  it("answers NULL rather than guessing when it cannot tell", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { bundleChanged } = await import("../scripts/rounds-gate.mjs");
    // THE THIRD VALUE IS THE POINT. "The product did not change" is a strong
    // claim, and a wrong one would excuse a real regression — so a shallow
    // clone, an unknown sha, or no git at all must say nothing at all.
    const throws = () => {
      throw new Error("fatal: bad object");
    };
    expect(bundleChanged("aaaaaaa · x", "bbbbbbb · y", throws), "a git failure was read as 'unchanged'").toBe(null);
    expect(bundleChanged(undefined, "bbbbbbb · y", throws), "a missing build was read as an answer").toBe(null);
    expect(bundleChanged("not-a-sha", "bbbbbbb · y", throws), "an unparseable stamp was read as an answer").toBe(null);
  });

  it("does not shell out when both rounds name the same build", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { bundleChanged } = await import("../scripts/rounds-gate.mjs");
    // A pair runs twice on ONE build, which is the common case for this check.
    // Spawning git to compare a sha with itself is waste, and the answer is
    // known: the same build ships the same bundle.
    let calls = 0;
    const counting = () => {
      calls++;
      return "src/x.ts";
    };
    expect(bundleChanged("abc1234 · x", "abc1234 · y", counting)).toBe(false);
    expect(calls, "it shelled out to compare a build with itself").toBe(0);
  });
});

describe("a round that measured nothing", () => {
  const round = (ran: number, total: number) => ({
    build: "abc1234 · x",
    selftest: Array.from({ length: total }, (_, i) => ({
      name: `s${i}`,
      ok: true,
      ...(i >= ran ? { skipped: true } : {}),
    })),
  });

  it("is not read as a clean one", () => {
    /**
     * ROUNDS 360 AND 361. A wrapper bug made every `PowerPoint.run` resolve
     * undefined; both rounds archived with exit 0 and a green gate while
     * reading nothing at all — 13 of 16 scenarios skipped for want of a chart
     * that could never be inserted. Nothing complained, because a SKIPPED
     * scenario is not a FAILING one, and every check in the gate is about
     * failures. I read them as evidence for two rounds and concluded from them
     * that a healthy document was damaged.
     */
    expect(coverageOf(round(3, 16)).collapsed, "a round that ran 3 of 16 passed for evidence").toBe(true);
    expect(coverageOf(round(3, 14)).collapsed).toBe(true);
  });

  it("does not fire on any round this archive has ever produced", () => {
    /**
     * THE THRESHOLD IS MEASURED. Across 339 archived rounds the share of
     * scenarios that ran is 284 at 100%, 43 at 90-99%, 6 at 80-89%, 3 at
     * 70-79% — then nothing until 21%, 19%, 19%. A fifty-point dead zone.
     * These are the real edges of it, so a future cut that wanders into the
     * gap fails here rather than on a night.
     */
    expect(coverageOf(round(11, 14)).collapsed, "79% — the worst legitimate round on file").toBe(false);
    expect(coverageOf(round(12, 14)).collapsed).toBe(false);
    expect(coverageOf(round(16, 16)).collapsed).toBe(false);
    // And the top of the collapsed group, 3 of 14 = 21%.
    expect(coverageOf(round(3, 14)).collapsed).toBe(true);
    // The floor sits inside the gap rather than on either edge of it.
    expect(COVERAGE_FLOOR).toBeGreaterThan(0.21);
    expect(COVERAGE_FLOOR).toBeLessThan(0.79);
  });

  it("says nothing about a round that carries no self-test at all", () => {
    /**
     * Several early rounds hold only host answers. Reporting an ABSENT
     * self-test as a collapsed one would be the same defect this guard exists
     * to stop, pointed the other way: an absence read as a measurement.
     *
     * EQUIVALENT MUTANT, recorded rather than chased: deleting the `total > 0`
     * guard changes nothing this test can see, because `0 / 0` is `NaN` and
     * `NaN < COVERAGE_FLOOR` is already false. Two mechanisms deliver the same
     * answer. The guard stays because relying on NaN's comparison semantics to
     * express "there was nothing to measure" is a thing a reader has to work
     * out, and this file is read at 2am — but the assertion below pins the
     * BEHAVIOUR, which is what matters, and would survive either spelling.
     */
    expect(coverageOf({ build: "x", selftest: [] }).collapsed).toBe(false);
    expect(coverageOf({ build: "x" }).collapsed).toBe(false);
    expect(coverageOf(undefined).collapsed).toBe(false);
  });

  it("counts a skip as not-run whatever its ok flag says", () => {
    // A skip carries `ok: false` in some scenarios and `ok: true` in others —
    // `blindSkip` and `countlessSkip` both set `ok: false`, while an early-out
    // may not. Coverage is about whether the question was PUT, so only
    // `skipped` decides it.
    const mixed = {
      build: "x",
      selftest: [
        { name: "a", ok: false, skipped: true },
        { name: "b", ok: true, skipped: true },
        { name: "c", ok: true },
      ],
    };
    expect(coverageOf(mixed).ran, "an ok:true skip was counted as having run").toBe(1);
    expect(coverageOf(mixed).collapsed).toBe(true);
  });
});
