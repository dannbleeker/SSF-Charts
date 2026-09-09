#!/usr/bin/env node
/**
 * Fail when a scenario that WAS passing has stopped.
 *
 * The only automatic check this project has on a round's own result. Every other
 * number a round produces is read by a person and then filed, so rounds 070-072
 * — which took `same scale across the deck` from 35 consecutive failures to
 * three consecutive passes — bought a result that nothing guarded. A later build
 * could take it back and no gate would notice.
 *
 *     npm run rounds:gate
 *
 * NOT part of CI, and it cannot be: CI has no rounds. It runs after archiving,
 * against `rounds/`, and its exit code is the point.
 *
 * It is deliberately quiet about everything else. A gate that reports on a round
 * being merely worse is a gate that cries wolf on a host whose mood swings 4-of-5
 * to 1-of-5 with nothing changed — and `docs/BACKLOG.md` records what happens to
 * a gate like that: it gets switched off.
 */
import { readFileSync, readdirSync } from "fs";
import { execFileSync } from "child_process";
import { isMain } from "./is-main.mjs";
import {
  scenarioRegressions,
  profileDivergence,
  roundProfile,
  traceNovelty,
  poolScenarioPopulations,
  poolGroupingOutcome,
  poolProfileDisagreements,
  poolPairPosition,
  poolFallbackRates,
  poolFullestSlide,
  CLEAN_SLIDE_CEILING,
  poolDriverRuns,
  unreadSignals,
  poolInPlaceUpdates,
  roundSpanSeconds,
  paneAgeAtStartSeconds,
  probeFlipsWithinBuild,
  deckGeometryFaults,
  fatalScenarios,
  fatalRateBreaches,
  fatalDeathsAllowed,
  deathsAcknowledged,
  poolCrashLastSteps,
  crashStepKey,
  scenarioRuns,
} from "./triage.mjs";
import { pendingAlreadyAnswered, UNSTABLE_ANSWERS, FATAL_SCENARIO_RATE } from "./host-baseline.mjs";

/**
 * Did the SHIPPED BUNDLE change between two archived rounds?
 *
 * `true` changed · `false` identical · `null` CANNOT TELL, and the third value
 * is the point. A shallow clone, a build stamp that names no commit, or no git
 * at all must leave the gate exactly as it was rather than volunteer a guess —
 * "the product did not change" is a strong claim and a wrong one would excuse a
 * real regression. Everything unknown returns null and prints nothing.
 *
 * `src/` ONLY. `scripts/` and `test/` do not ship; a round runs against the
 * deployed bundle, and that is built from `src/`. A commit that rewrites the
 * driver cannot change what the pane executes.
 */
export function bundleChanged(fromBuild, toBuild, run = execFileSync) {
  const sha = (b) => /^([0-9a-f]{7,40})/.exec(String(b ?? ""))?.[1];
  const from = sha(fromBuild);
  const to = sha(toBuild);
  if (!from || !to) return null;
  if (from === to) return false;
  try {
    if (!run("git", ["diff", "--name-only", `${from}..${to}`, "--", "src/"], { encoding: "utf8" }).trim()) return false;
    /**
     * A FILE UNDER `src/` CHANGED — but did any LINE that runs?
     *
     * Round 440 is why this is here. Its 4:3 leg failed, the gate could not say
     * "read this as the host", and the reason was 27 lines of COMMENT I had
     * added to `app.ts` and `powerpoint.ts` an hour earlier. `--name-only` sees
     * a file, so the guard reported the product as changed when the bundle was
     * behaviourally identical, and the one check that exists to spare a reader
     * that hunt sent them on it.
     *
     * DELIBERATELY ASYMMETRIC, because the two errors are not equal. Saying
     * "changed" when only comments moved costs a manual `git diff`. Saying
     * "unchanged" when something real moved EXCUSES A REGRESSION, which is the
     * failure this whole check exists to prevent. So anything that is not
     * plainly a comment or a blank line counts as a change: a `//` inside a
     * string literal reads as code here, and that is the safe direction.
     *
     * `"comments"` is a THIRD value, not `false`. The caller prints it as a
     * fact, never as the "read this as the host" verdict — a reader who wants
     * that conclusion can have it from the fact, and a heuristic must not hand
     * out the strong claim on its own.
     */
    const diff = run("git", ["diff", "-U0", `${from}..${to}`, "--", "src/"], { encoding: "utf8" });
    const moved = String(diff)
      .split("\n")
      .filter((l) => /^[+-]/.test(l) && !/^(\+\+\+|---)/.test(l))
      .map((l) => l.slice(1).trim())
      .filter((l) => l.length > 0);
    const code = moved.filter((l) => !(l.startsWith("//") || l.startsWith("*") || l.startsWith("/*") || l === "*/"));
    return code.length > 0 ? true : "comments";
  } catch {
    return null;
  }
}

/**
 * Every archived round, oldest first — the order `scenarioRegressions` expects.
 *
 * A FILE THAT WILL NOT PARSE IS NAMED, NOT THROWN OVER, and the difference
 * decides what a night does next. `archive` writes straight to the final path
 * rather than writing-then-renaming, so an interrupted write leaves a truncated
 * round behind; unguarded, one of those took the whole gate down with a
 * SyntaxError, node exited 1, and `cycle.mjs` reads any non-zero gate as a
 * REGRESSION. A corrupt download would have stopped the night reporting a fall
 * that never happened.
 *
 * Skipping it silently would be the other half of the same mistake: a round
 * missing from the comparison is a round whose regression cannot be seen, so
 * the caller is told which files were dropped and decides what that is worth.
 * `triage.mjs` already takes exactly this line for the same reason.
 */
/**
 * How many crash reports the driver wrote, whether or not a round survived.
 *
 * THE HALF OF THE CRASH RECORD THAT IS NOT IN `rounds/`. A round the driver
 * never recovered archives nothing — `round.mjs` says so in as many words: "A
 * crashed round archives nothing: it never reaches the download button" — so
 * every crash rate computed from the round files is conditional on recovery
 * having succeeded, and is a floor rather than a rate.
 *
 * Measured 2026-09-01: 77 reports in `crashes/` against 46 crash events in
 * archived `driverRun.recovered`. Forty per cent of the crashes this project has
 * seen left no round file to be counted in.
 *
 * Counted rather than parsed: the reports are prose for a person, and all this
 * needs from them is how many there are.
 */
export function countCrashReports(dir = "crashes", list = readdirSync) {
  try {
    return list(dir).filter((f) => f.endsWith(".md")).length;
  } catch {
    // No crashes directory is not zero crashes, but it is nothing to report.
    return 0;
  }
}

/**
 * Builds whose crash records are INSTRUMENT ARTEFACTS rather than host
 * behaviour, and must not be counted as either.
 *
 * On 2026-09-02 a wrapper made every `PowerPoint.run` resolve undefined, so
 * every host read on these four builds returned nothing. Their records describe
 * a broken add-in, not a broken host, and pooling them would put deaths against
 * scenarios that never got to run. Kept as a named list rather than a date
 * range because a build is a fact and a date is an inference.
 */
export const POISONED_BUILDS = new Set(["b5c534a", "3eaab20", "2934204", "6421ba2"]);

/**
 * Host deaths a person has read, one entry per CRASH RECORD.
 *
 * SHIPS EMPTY, AND THAT IS NOT AN OVERSIGHT. The mechanism is the owner's to
 * use; writing the first entry is his signature and nobody else's.
 *
 * WHY IT EXISTS. A scenario absent from `FATAL_SCENARIO_RATE` has a ceiling of
 * 0, and at p=0 the allowance is 0 for every denominator — so its first host
 * death breaches, and the crash record being permanent means it breaches for
 * ever. The nightly cycle stopped after one round from round 428 onward, which
 * cost the 4:3 validation arm entirely, because it is leg three.
 *
 * WHY NOT JUST RAISE THE CEILING. Because that is a different statement.
 * A ceiling says *every future death of this shape is expected*; a receipt says
 * *I read this one*. The table's own docstring forbids the first — "Do not add
 * a name here to quiet a gate" — and this is deliberately placed beside
 * `POISONED_BUILDS`, the other read-time exclusion, rather than anywhere near
 * the ceilings.
 *
 * WHAT A RECEIPT CANNOT DO, which is what stops it becoming a silencer:
 *
 *   - It clears the EXIT, never the count. The breach still prints, the death
 *     still counts for ever, no rate moves.
 *   - It cannot touch a scenario with a real ceiling. Those already have a
 *     green path — the rate falls as clean runs accumulate — and signing one
 *     away would be the ceiling edit under another name.
 *   - It cannot cover a SECOND death. Two deaths is the archive's own signal
 *     that the first was not a one-off: across the nine scenarios that have
 *     ever killed the host, every repeat death landed within 25 rounds of its
 *     predecessor, median 1. Five of the nine never died again — which is why
 *     one death is signable and two are not.
 *   - It cannot pre-authorise anything. The key is a specific file in
 *     `crashes/`. A death that has not happened has no filename.
 *   - It goes STALE and fatal if the record stops crediting that scenario, so
 *     an entry nobody re-reads fails the gate rather than quietly holding.
 *
 * Each entry: `record` (the filename in `crashes/`), `scenario` (the name it
 * credited, which must still match), `seen` (the date a person read it), and
 * `why` — a reason of the kind `KNOWN_DIVERGENCES` demands, naming the section
 * of `docs/BACKLOG.md` that carries the evidence. "We have not looked into it"
 * is not a reason.
 *
 * @type {{record: string, scenario: string, seen: string, why: string}[]}
 */
export const DEATHS_ACKNOWLEDGED = [
  // EMPTY AGAIN, AND THE ONE ENTRY IT HELD WAS RETIRED RATHER THAN DROPPED IN
  // PASSING. `2026-09-07T22-13-42-crashed-run.json` signed the FIRST death of
  // `what a chart kind costs` on 2026-09-08. That scenario took a second death
  // the same night, which no receipt may cover, and on 2026-09-09 the owner
  // seeded it at 50 per 1000 — so it now has a green path of its own, and a
  // receipt beside it would be a second lock on an open door. The gate says so
  // itself: a receipt for a scenario carrying a ceiling is STALE, and stale is
  // exit 2.
  //
  // That is the mechanism's whole life so far, and it behaved: signed once,
  // refused to sign the second, retired when the ceiling replaced it.
];

/**
 * The last step of one crash record, by filename — the line the host died on.
 *
 * Printed beside a signable death so the person signing sees WHAT they are
 * signing rather than a scenario name. Returns "" when the record is not in the
 * pool, which happens for a hand-built record in a test and for one dropped by
 * `POISONED_BUILDS`.
 */
function lastStepOf(crashes, file) {
  const rec = (crashes ?? []).find((c) => c?._file === file);
  const steps = Array.isArray(rec?.steps) ? rec.steps : [];
  return steps.length ? String(steps[steps.length - 1]) : "";
}

/**
 * Every crash record, parsed — the half of the evidence `loadRounds` cannot see.
 *
 * A round that dies files nothing, so `rounds/` holds only the runs that lived.
 * These are the rest, and they carry the one thing no verdict list can express:
 * which scenario was in flight when the host went. See `fatalScenarios`.
 *
 * A record that will not parse is named and skipped, exactly as `loadRounds`
 * treats a truncated round — one bad file must not refuse the other eighty.
 *
 * ONE EVENT FILED TWICE IS ONE EVENT. `2026-08-29T03-32-07-crashed-run.json`
 * and `2026-08-29T08-27-00-crashed-run.json` carry the same build, the same
 * `startedAt` to the millisecond, and 392 byte-identical steps: one crash
 * re-archived five hours later under a second name. Counted twice it put
 * `same scale across the deck` at 10 deaths when it has 9.
 *
 * Deduped HERE, at read time, rather than by deleting a file. The archive is
 * append-only and what it means is the owner's call; a reader that counts an
 * event once is not a reader that edits history. The key is deliberately
 * `build + startedAt + steps`, all three: two genuine crashes of the same
 * build seconds apart have different `startedAt`, and a record that shares a
 * start but diverges in steps is a different run of the same session and must
 * still count. Exactly one pair in 86 records matches.
 */
export function loadCrashRecords(dir = "crashes", list = readdirSync, read = readFileSync) {
  const unreadable = [];
  const duplicates = [];
  let records;
  try {
    const seen = new Set();
    records = list(dir)
      .filter((f) => f.endsWith("-crashed-run.json"))
      .sort()
      .map((f) => {
        try {
          const parsed = JSON.parse(read(`${dir}/${f}`, "utf8"));
          if (parsed && typeof parsed === "object") parsed._file = f;
          return parsed;
        } catch {
          unreadable.push(f);
          return null;
        }
      })
      .filter(Boolean)
      .filter((c) => !POISONED_BUILDS.has(String(c?.build ?? "").split(" ")[0]))
      .filter((c) => {
        const key = `${c?.build ?? ""}|${c?.startedAt ?? ""}|${JSON.stringify(c?.steps ?? [])}`;
        // A record carrying NEITHER a build nor a start is not identifiable, so
        // it is kept rather than folded into the first other anonymous one.
        if (!c?.build && !c?.startedAt) return true;
        if (seen.has(key)) {
          duplicates.push(c._file);
          return false;
        }
        seen.add(key);
        return true;
      });
  } catch {
    // No crashes directory at all is a fresh checkout, not a finding.
    return Object.assign([], { unreadable, duplicates });
  }
  return Object.assign(records, { unreadable, duplicates });
}

export function loadRounds(dir = "rounds", list = readdirSync, read = readFileSync) {
  const unreadable = [];
  const rounds = list(dir)
    .filter((f) => /^\d{3}-.*\.json$/.test(f))
    .sort()
    .map((f) => {
      try {
        return JSON.parse(read(`${dir}/${f}`, "utf8"));
      } catch {
        unreadable.push(f);
        return null;
      }
    })
    .filter(Boolean);
  rounds.unreadable = unreadable;
  return rounds;
}

/**
 * The share of a round's scenarios that actually RAN, and whether so few did
 * that the round is evidence of nothing.
 *
 * WHY THIS EXISTS. On 2026-09-02 a wrapper bug made every `PowerPoint.run`
 * resolve undefined. Rounds 360 and 361 archived as ordinary rounds — exit 0,
 * gate green — while reading nothing at all: 13 of 16 scenarios skipped for want
 * of a chart that could never be inserted, `deckSlides` unreadable, deck growth
 * `NaN`. Nothing complained. I read them as evidence for two rounds and
 * concluded from them that a healthy deck was damaged. The instrument was
 * broken and every check in this gate said the product was fine, because a
 * skipped scenario is not a failing one.
 *
 * THE THRESHOLD IS MEASURED, NOT CHOSEN. Across 339 archived rounds the
 * proportion of scenarios that ran is: 284 at 100%, 43 at 90-99%, 6 at 80-89%,
 * 3 at 70-79% — and then nothing at all until 21%, 19%, 19%. A fifty-point dead
 * zone with three rounds below it and 336 above. Any cut inside that gap
 * separates them; 50% sits in the middle of it and cannot fire on a round this
 * archive has ever produced.
 *
 * The third round below the gap is 287 (`2bf766b`, 3 of 14), which predates the
 * wrapper bug entirely and has been sitting in the archive unremarked.
 *
 * NOT A REGRESSION, so never exit 1. A collapsed round means the instrument
 * failed, not the product — the same distinction `unreadable` draws everywhere
 * else in this project. It is exit 2's case: the gate could not do its job.
 */
export const COVERAGE_FLOOR = 0.5;

export function coverageOf(round) {
  const s = round?.selftest ?? [];
  const total = s.length;
  const ran = s.filter((x) => !x.skipped).length;
  // A round with no self-test at all is not collapsed — it is a different kind
  // of round, and several early ones carry only host answers. Judging those
  // here would report an absence as a failure, which is the mistake this whole
  // function exists to stop.
  return { ran, total, collapsed: total > 0 && ran / total < COVERAGE_FLOOR };
}

if (isMain(import.meta.url, process.argv[1])) {
  // EXIT 2 MEANS "I COULD NOT DO MY JOB", and it has to be its own code. This
  // gate's whole value is that exit 1 means a scenario stopped passing — so a
  // gate that fell over on its own reading, exiting 1 the way node does for any
  // uncaught throw, is indistinguishable from the finding it exists to report.
  // `cycle.mjs` acts on that difference: 1 stops the night saying a scenario
  // regressed, 2 stops it saying the gate needs looking at. Same convention
  // `triage.mjs` already uses for a file it cannot read.
  let rounds;
  try {
    rounds = loadRounds();
  } catch (err) {
    console.error(`  the gate could not read rounds/: ${err?.message ?? err}`);
    console.error("  Nothing was judged. This is not a regression — see docs/ROUNDS.md.");
    process.exit(2);
  }
  if (rounds.unreadable?.length) {
    // Loud, because a round missing from the comparison is a round whose
    // regression cannot be seen. Not fatal: the rest of the archive still
    // answers, and refusing to judge 57 good rounds over one bad file would be
    // the worse trade.
    console.error(`  ${rounds.unreadable.length} archived round(s) WOULD NOT PARSE and were left out:`);
    for (const f of rounds.unreadable) console.error(`    rounds/${f}`);
    console.error("  A regression inside one of those cannot be seen from here.");
  }
  if (!rounds.length) {
    console.error("  no readable rounds to judge — nothing was checked");
    process.exit(2);
  }
  /**
   * A ROUND THAT MEASURED NOTHING IS NOT EVIDENCE, and until now it read as a
   * clean one. See `coverageOf` for what that cost.
   *
   * Two separate consequences, because they answer different questions:
   *
   *   - a collapsed round anywhere in the archive is dropped from the
   *     comparison population. Leaving it in lets an instrument failure set a
   *     baseline, so a scenario that "passed" in a round which never ran it
   *     silently lowers the bar for every round after.
   *   - a collapsed round AT THE HEAD stops the gate. The build under judgement
   *     has no evidence either way, and saying so is the entire job. Exit 2,
   *     not 1: nothing regressed, the instrument failed.
   */
  const atHead = rounds[rounds.length - 1];
  const collapsed = rounds.filter((r) => coverageOf(r).collapsed);
  if (collapsed.length) {
    console.error(`  ${collapsed.length} round(s) RAN ALMOST NOTHING and are excluded from the comparison:`);
    for (const r of collapsed) {
      const c = coverageOf(r);
      console.error(`    ${r.build ?? "?"} — ${c.ran} of ${c.total} scenarios ran; the rest were skipped, not passed`);
    }
    console.error("  A skipped scenario is not a passing one. Read these as the instrument, not the product.");
  }
  if (coverageOf(atHead).collapsed) {
    const c = coverageOf(atHead);
    console.error(
      `\n  THE NEWEST ROUND MEASURED NOTHING — ${c.ran} of ${c.total} scenarios ran. This build has no\n` +
        "  evidence for or against it, so nothing has been judged. Suspect the instrument before the\n" +
        "  product: rounds 360 and 361 looked exactly like this and the cause was a wrapper in the\n" +
        "  renderer that made every host call return undefined. See docs/ROUNDS.md.",
    );
    process.exit(2);
  }
  // JUDGED ON WHAT WAS ACTUALLY MEASURED, from here down.
  rounds = rounds.filter((r) => !coverageOf(r).collapsed);
  /**
   * THE FOURTH OUTCOME CLASS, and the reason this gate has been half-blind.
   *
   * Everything above judges VERDICTS, and a scenario that kills the host leaves
   * none: no verdict, no round file, nothing in `rounds/` at all. So the worst
   * thing a scenario can do has been the one thing this gate could not see, and
   * it has been hiding the sharpest finding in the archive — `same scale across
   * the deck`, 0 failures in 282 recorded verdicts, joint-safest in the suite,
   * and the scenario the host died inside TEN times.
   *
   * Read from the crash records' own steps, because nothing else knows: a
   * salvaged round carries verdicts, and a killed scenario has none.
   */
  const crashes = loadCrashRecords();
  if (crashes.unreadable?.length)
    console.error(`  ${crashes.unreadable.length} crash record(s) would not parse: ${crashes.unreadable.join(", ")}`);
  const fatal = fatalScenarios(crashes);
  const runs = scenarioRuns(rounds, crashes);
  if (fatal.attributed) {
    console.log(
      `\n  ${fatal.attributed} crash(es) died INSIDE a scenario; ${fatal.unattributed} died elsewhere ` +
        "(probe phase, deck scan) and are credited to nothing.",
    );
    // SORTED BY RATE, not by count. The count order is upside down — it ranks
    // how long a scenario has been exposed, and the two that kill the host on a
    // third to a half of their runs sit below one that kills on 2% of them.
    const byRate = Object.entries(fatal.deaths)
      .map(([name, count]) => ({ name, count, ran: runs[name] ?? 0 }))
      .map((r) => ({ ...r, rate: r.ran ? (1000 * r.count) / r.ran : null }))
      .sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1));
    console.log("  deaths per 1000 runs, and the count that would trip the gate:");
    for (const r of byRate) {
      const allowed = FATAL_SCENARIO_RATE[r.name] ?? 0;
      const rate = r.rate === null ? "  no runs" : r.rate.toFixed(1).padStart(8);
      // The trip point printed alongside, because a ceiling in per-1000 is not
      // something a reader can compare to "5 of 11" in their head — and the
      // gap between them is the headroom the noise bound buys.
      const trips = r.ran ? Math.floor(fatalDeathsAllowed(allowed, r.ran)) + 1 : 0;
      console.log(
        `    ${rate}   ${String(r.count).padStart(2)} of ${String(r.ran).padStart(4)}  ` +
          `(ceiling ${String(allowed).padStart(3)}, trips at ${String(trips).padStart(2)})  ${r.name}`,
      );
    }
  }
  const breaches = fatalRateBreaches(fatal.deaths, runs, FATAL_SCENARIO_RATE);
  const receipts = deathsAcknowledged(breaches, fatal.credits, DEATHS_ACKNOWLEDGED, FATAL_SCENARIO_RATE);
  const signedFor = new Map(receipts.cleared.map((c) => [c.name, c.receipt]));
  /**
   * A STALE RECEIPT IS CHECKED WHETHER OR NOT ANYTHING IS BREACHING, and it was
   * not — which the mechanism's very first real use exposed.
   *
   * The check lived inside the breach block. On 2026-09-09 the owner seeded
   * `what a chart kind costs` at 50 per 1000, which is exactly the event that
   * makes its receipt meaningless: a scenario with a ceiling has a green path of
   * its own and must not also be signed for. The ceiling removed the breach, the
   * breach block was skipped, and the now-pointless entry sat in
   * `DEATHS_ACKNOWLEDGED` undetected — while its own docstring promised "It goes
   * STALE and fatal if the record stops crediting that scenario".
   *
   * So the guard that keeps this ledger from becoming a silencer did not run in
   * the one case that had ever occurred. It runs first now, before anything can
   * return: an entry nobody re-read is a reason to refuse to judge, not a
   * footnote underneath a breach that may not be there.
   */
  if (receipts.stale.length) {
    console.error("\n  A RECEIPT NO LONGER MATCHES WHAT IT SIGNED FOR — this gate cannot judge that:");
    for (const s of receipts.stale) console.error(`    ${s.record} → ${s.why}`);
    console.error("  An entry nobody re-read is not a receipt. Fix or remove it; until then nothing here is judged.");
    process.exit(2);
  }
  /**
   * The regression verdict, COMPUTED BEFORE the breach block can exit.
   *
   * It used to be computed after, so a round that tripped the host-death check
   * never had its verdicts judged at all — and since round 428 that check has
   * been permanently red, so `scenarioRegressions` has not run once. Two fatal
   * questions, and the one that fires first was hiding the other.
   *
   * ONLY THIS CALL IS HOISTED, and only because it is pure: it reads `rounds`,
   * which is already loaded and validated, and touches nothing else. The REPORT
   * is deliberately left where it is — sixteen unguarded triage calls sit
   * between here and the exit, and a throw in any of them would make node exit
   * 1, which `cycle.mjs` reads as "a scenario that WAS passing has stopped".
   * That is `d12dadb`'s bug, and moving the report would reintroduce it.
   */
  const gone = scenarioRegressions(rounds);
  if (breaches.length) {
    console.error("\n  A SCENARIO IS KILLING THE HOST MORE OFTEN THAN IT DID:");
    for (const b of breaches)
      console.error(
        `    ${b.name} — ${b.rate.toFixed(1)} per 1000 (${b.count} of ${b.runs}), ceiling ${b.allowed}` +
          (b.allowed === 0 ? " (it had never killed the host before)" : "") +
          (signedFor.has(b.name) ? `   [acknowledged ${signedFor.get(b.name).seen}]` : ""),
      );
    console.error(
      "  A scenario that stops passing is exit 1 here; one that starts taking PowerPoint down is\n" +
        "  at least as serious and produces no verdict to notice it by, and exits 3.\n" +
        "  A RATE, not a count, and the difference matters when you act on this: the count version\n" +
        "  went red on the next death after it was seeded, because deaths only accumulate. This one\n" +
        "  falls on its own as runs pile up without deaths, so a landed fix protects itself and\n" +
        "  NOTHING NEEDS EDITING to make it green. If a ceiling in `FATAL_SCENARIO_RATE` is genuinely\n" +
        "  meant to be higher, that is a person deciding a scenario may kill PowerPoint more often\n" +
        "  than it used to. See docs/ROUNDS.md.",
    );
    // EXCEPT AT A CEILING OF ZERO, WHERE THE PARAGRAPH ABOVE IS NOT TRUE. It
    // was written for the seeded ceilings, and for those it holds: the allowance
    // is `p*n + 2*sqrt(p(1-p)n)`, which grows with runs. At `p = 0` that
    // expression is 0 for every n — `fatalDeathsAllowed` returns 0 before it
    // computes anything — so a scenario absent from the table breaches on its
    // first death and CANNOT fall back under, however many clean rounds follow.
    // The crash record is in `crashes/` permanently.
    //
    // So a first-ever death makes this gate red until a person changes
    // something, while the table's own docstring says "Do not add a name here to
    // quiet a gate". Both instructions are right on their own and together they
    // leave no green path. Said out loud here rather than resolved, because what
    // a first death should DO — stop the night once, or stop it until
    // acknowledged — is a decision about the instrument, not a bug in it.
    const unsignable = receipts.standing.filter((b) => b.allowed === 0);
    if (unsignable.length)
      console.error(
        "\n  A CEILING OF ZERO CANNOT GO GREEN ON ITS OWN. The allowance grows with runs only when the\n" +
          "  ceiling is above zero; at zero it is zero for every denominator, and the crash stays in\n" +
          "  `crashes/`. Clean rounds will not clear these.",
      );
    for (const b of unsignable) {
      const record = (fatal.credits?.[b.name] ?? []).filter(Boolean);
      if (b.count === 1 && record.length === 1) {
        // ONE DEATH: signable, and the evidence to sign it by is printed here
        // rather than left to be dug out of `crashes/`. What a reader needs is
        // whether this death has a SHAPE the archive has seen before.
        const last = lastStepOf(crashes, record[0]);
        const key = crashStepKey(last);
        // MAPPED TO THE SHAPE THAT FUNCTION TAKES, not handed the raw records.
        // `poolCrashLastSteps` sorts ties on `at[0].localeCompare`, so it needs
        // a `name` on every record; `loadCrashRecords` stamps `_file` and never
        // `name`, and passing its output straight in throws on the sort. Its
        // only other caller builds `{name, steps, build}` by hand, which is the
        // contract — met here rather than loosened there.
        const pooled = poolCrashLastSteps(crashes.map((c) => ({ name: c._file ?? "", steps: c.steps ?? [] }))).find(
          (g) => g.key === key,
        );
        console.error(
          `\n    \`${b.name}\` has ONE death and can be acknowledged. Read it, then add to\n` +
            `    DEATHS_ACKNOWLEDGED in this file:\n` +
            `      { record: "${record[0]}", scenario: "${b.name}", seen: "<date>", why: "<backlog section>" }\n` +
            `    last step: ${String(last).trim().slice(0, 96)}\n` +
            `    ${pooled ? `${pooled.n} of ${crashes.length} kept records end on \`${key}\`` : "no pooling for that step"}`,
        );
      } else {
        // TWO OR MORE: not signable, deliberately. The only green path is a
        // ceiling, and the smallest one that holds is printed so nobody has to
        // solve `p*n + 2*sqrt(p(1-p)n) >= count` in their head.
        //
        // "THE POINT ESTIMATE IS NOT HIGH ENOUGH" WAS PRINTED HERE FOR HALF A
        // DAY AND IS NOT A TRUE SENTENCE. Setting the ceiling to the current
        // rate makes the allowance `count + 2*sqrt(p(1-p)n)`, which exceeds
        // `count` by the whole noise bound — so the point estimate always holds
        // today. Its first live outing said 2 of 45 needed more than 44.4 when
        // the floor is 12 and 44.4 allows 4.73.
        //
        // The only case where the floor really does exceed the rate is integer
        // rounding at the bottom of the scale: 1 death in 4,000 runs is 0.25
        // per 1000 and the smallest integer ceiling is 1. That is arithmetic
        // about the units, not a statistical warning, so it is not worth a
        // branch.
        //
        // What IS true is forward-looking, and is the table's whole design: a
        // ceiling seeded at today's rate stays silent while the scenario stays
        // exactly as bad as it is, and fires when it gets worse.
        let c = 1;
        while (c < 1000 && fatalDeathsAllowed(c, b.runs) < b.count) c++;
        console.error(
          `\n    \`${b.name}\` has ${b.count} deaths and CANNOT be acknowledged — two deaths is the archive's\n` +
            `    own signal that the first was not a one-off. Its only green path is a ceiling: the smallest\n` +
            `    integer that holds ${b.count} of ${b.runs} is ${c} per 1000, and the current rate of ` +
            `${((1000 * b.count) / b.runs).toFixed(1)} holds too.\n` +
            `    A ceiling seeded at the current rate goes quiet while this stays as bad as it is, and fires\n` +
            `    when it gets worse — which is what this table is for.`,
        );
      }
    }
    for (const c of receipts.cleared)
      console.error(
        `\n  ACKNOWLEDGED, AND STILL COUNTED ABOVE — a death this gate has been told a person read:\n` +
          `    ${c.name} · ${c.receipt.record} · seen ${c.receipt.seen}\n` +
          `    why: ${c.receipt.why}`,
      );
    // EXIT 3, NOT 1, AND THE CODE IS THE POINT. This gate has two fatal checks
    // and they shared one exit code, so `cycle.mjs` — its only consumer — printed
    // "a scenario that WAS passing has stopped" for both. Round 428 tripped THIS
    // one: 19 of 19 scenarios green, one host death, and a stop message that
    // sent its reader looking for a failed verdict that did not exist.
    //
    // `docs/ROUNDS.md` said of this gate that it "answers several different
    // questions and they must not be confused. Exactly one of them is fatal."
    // Two are, and that doc is corrected alongside this.
    //
    // GUARDED ON WHAT IS STILL STANDING, not on `breaches.length`. A breach a
    // person has signed for still prints — the count never moves — but it no
    // longer stops the night. If everything here is signed, control falls
    // through to the regression check below, which is the whole point of the
    // receipt.
    if (receipts.standing.length) {
      // AND THE OTHER FATAL QUESTION IS ANSWERED BEFORE LEAVING, because until
      // now this exit hid it: a round that killed the host never had its
      // verdicts judged, and the host-death check has been red since round 428.
      if (gone.length)
        console.error(
          `\n  AND ${gone.length} scenario(s) STOPPED PASSING in the same round — read that too:\n` +
            gone.map((g) => `    ${g.name} — failed ${g.failed} of ${g.ran} at this profile`).join("\n"),
        );
      process.exit(3);
    }
  }
  // A SECOND, DIFFERENT QUESTION. The gate above asks whether a scenario fell
  // against its OWN history; this asks whether one slide size failed what
  // another passed on the same build. Round 077 was exactly that — 10 of 13 at
  // 4:3 against 13 of 13 at 16:9 — and nothing said so automatically.
  //
  // Reported, never fatal. A nightly cycle runs 16:9 twice and 4:3 once as
  // VALIDATION, and the agreed response to divergence is to run 4:3 again or on
  // its own, not to fail the build. Exiting non-zero here would turn a signal
  // that means "look closer" into one that means "stop", which is how a useful
  // report becomes an ignored one.
  const diverged = profileDivergence(rounds);
  const real = diverged.filter((d) => !d.flaky);
  const flaky = diverged.filter((d) => d.flaky);
  if (real.length) {
    console.log(`  ${real.length} scenario(s) DIVERGED between slide sizes on the same build:`);
    for (const d of real) {
      console.log(
        `    ${d.name} — passed at ${d.passedIn.join(", ")}, failed at ${d.failedIn.join(", ")} (${d.build})`,
      );
      // THE SCENARIO'S OWN RECORD AT EACH PROFILE, beside the split. One build
      // cannot tell "fails only here" from "fails at both, and this time the
      // coin landed here" — and the first divergence this gate reported after
      // being unblocked was the second kind, four points apart and fifty rounds
      // stale, wearing a 4:3-flavoured detail.
      const rates = Object.entries(d.history ?? {})
        .map(
          ([prof, h]) => `${prof} ${h.failed}/${h.ran}${h.ran ? ` (${((100 * h.failed) / h.ran).toFixed(0)}%)` : ""}`,
        )
        .join(" · ");
      if (rates) console.log(`      lifetime, this scenario: ${rates}`);
    }
    console.log(
      "  Run that profile again, or as a pair, before treating it as a property of the slide size —\n" +
        "  and read the lifetime line first: rates that match mean the split is what chance does.",
    );
  }
  // NAMED APART, because the response is different. A profile that disagrees
  // with ITSELF has said nothing about its slide size, and sending someone to
  // investigate an aspect ratio for a scenario that is simply flaky is how a
  // useful report teaches people to ignore it.
  //
  // This is the shape the check produced on its first live outing: `explode a
  // degraded picture` passed at 4:3, then passed once and failed once at 16:9
  // on build 17a8204. "Diverged between slide sizes" was true of the worst
  // reading and wrong about the cause.
  if (flaky.length) {
    console.log(`  ${flaky.length} scenario(s) were UNSTABLE WITHIN a slide size, which is not divergence:`);
    for (const d of flaky)
      console.log(`    ${d.name} — passed and failed at ${d.unstableIn.join(", ")} on the same build (${d.build})`);
    console.log("  Treat that as a flaky scenario, not a property of the slide size.");
  }
  // A SCENARIO CAN PASS ON LESS THAN IT USED TO, and none of the three questions
  // around this one can see it. `scenarioRegressions` compares PASS to PASS;
  // divergence compares slide sizes; novelty reads the trace. But `same scale
  // across the deck` scores itself `scaled === charts.length` against a
  // population it DISCOVERS — `probeCharts` returns whatever the deck scan finds
  // — so round 088's `6 of 6` and every earlier round's `8 of 8` are both a pass
  // and the gate said "no scenario regressed" between them.
  //
  // Reported, never fatal, for the same reason as the two above: round 088's six
  // is downstream of a host stall that skipped the scenario seeding the probe
  // charts, which is weather rather than a fault. It is a reason to read the
  // round, and a reason not to quote the pass without its denominator.
  // WHAT THE VERDICT CANNOT SEE. Rounds 092 and 093, one build run twice: 20
  // charts grouped and none refused, then 15 grouped and 4 refused with three
  // slides ending on 24 shapes each — and both reported 13/13 with the identical
  // verdict line. Reported every round, because a number only printed when it
  // looks bad is a number nobody has a baseline for.
  //
  // NEVER A REGRESSION, and the pair above is exactly why: 0 and 4 on the same
  // build is inside this project's own noise floor (1 vs 5, nothing changed). It
  // is a reason to read the round, which is all this line claims.
  // TWO READINGS OF THE SAME FACT, COMPARED. Everything below groups rounds by
  // profile, so a round filed under the wrong one silently contaminates every
  // comparison it appears in — and rounds 115 and 116 did exactly that while
  // `PW_EXPECT_SIZE` reported a match, because the guard read the live host and
  // the archive recorded the pane. Loud, because a wrongly-filed round is worse
  // than a missing one: it answers.
  // THE PAIR IS NOT TWO SAMPLES OF ONE CONDITION. Printed above everything that
  // compares rounds, because every such comparison assumes it is.
  // HOW LONG THIS ROUND TOOK, printed before anything that reads its counters.
  // A slow round is a degraded round in this archive, and the reader has never
  // been able to see which kind they were looking at.
  const newest = rounds[rounds.length - 1];
  const span = roundSpanSeconds(newest);
  const priorSpans = rounds
    .slice(0, -1)
    .map(roundSpanSeconds)
    .filter((n) => typeof n === "number");
  if (span !== null && priorSpans.length >= 3) {
    const sorted = [...priorSpans].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const ratio = median ? span / median : 1;
    const age = paneAgeAtStartSeconds(newest);
    console.log(`
  THIS ROUND TOOK ${span}s (median of ${priorSpans.length} prior round(s): ${median}s)`);
    // THE READING THAT PREDICTS THE COUNTERS. See `paneAgeAtStartSeconds`.
    if (age !== null)
      console.log(
        age < 200
          ? `    Pane was FRESH at the start (${age}s old). Fresh-pane rounds average 0.4 post-retry failures.`
          : `    PANE WAS REUSED — ${age}s old at the start. Reused-pane rounds average 4.6 post-retry
` +
              `    failures against 0.4, and leave 60+ deck shapes against 16. Read the counters below as a
` +
              `    degraded sample; reload the pane between rounds to avoid it.`,
      );
    if (ratio >= 1.5)
      console.log(
        `    ${ratio.toFixed(1)}x the usual. Slow rounds in this archive average roughly twice the post-retry
` + `    failures of fast ones — read this round's counters as a degraded sample, not as a change.`,
      );
  }

  // THE PATHS THE CODE TOOK BECAUSE ITS FIRST CHOICE FAILED. Recorded thousands
  // of times and read by nothing until now — see `poolFallbackRates` for the
  // thirtyfold win and the 40% drift that both went unnoticed.
  // A FEATURE THAT HAS NEVER RUN. See `poolInPlaceUpdates`.
  const ip = poolInPlaceUpdates(rounds);
  // fell + threw: a host-side refusal is a fallback too, and counting only the
  // rule-based declines hid three of them for two rounds.
  const ipDown = ip.fell + ip.threw;
  if (ipDown > 0 && ip.ok === 0)
    console.log(
      `
  IN-PLACE UPDATE HAS NEVER SUCCEEDED — 0 successes against ${ipDown} fallbacks over ${ip.rounds} round(s).
` +
        "    #406 was titled 'The in-place update fired zero times and would not say why' and added the trace" +
        "\n    that answers it. The answer has been sitting in every round file since. See FALLBACKS below for why.",
    );
  else if (ip.ok > 0) {
    console.log(`
  in-place update: ${ip.ok} succeeded, ${ip.fell} declined, ${ip.threw} refused by the host over ${ip.rounds} round(s)`);
    // THIS ROUND ON ITS OWN, WITH ITS DENOMINATOR. A pooled total hides the
    // thing that makes two rounds incomparable: round 148 scored the same 3
    // successes as 147 out of 11 attempts rather than 13, because two scenarios
    // failed and never reached the update. "3 again" is not the same evidence.
    const now = poolInPlaceUpdates(rounds.slice(-1));
    const attempts = now.ok + now.fell + now.threw;
    console.log(
      `    this round: ${now.ok} succeeded of ${attempts} attempt(s)` +
        (now.threw ? ` — ${now.threw} refused BY THE HOST, which is a defect, not a decline` : ""),
    );
  }
  // WHY, not just how many. A decline the differ made on purpose and a write the
  // host threw out read identically in a total.
  if (ip.reasons.length) {
    console.log("    why it fell back:");
    for (const r of ip.reasons.slice(0, 6)) console.log(`      ${String(r.n).padStart(4)}x  ${r.why}`);
  }
  if (ip.unexplained.length)
    console.log(
      `    ${ip.unexplained.length} carried NO reason at all — open these first, they are the ones no category fits`,
    );

  // WHAT IT TOOK TO GET THIS ROUND AT ALL. A successful recovery erases its own
  // evidence — the round that follows it looks like any other — so a round run
  // against a host that was already unwell read as clean. Round 148 took three
  // attempts and then failed two scenarios that had not failed in 109 rounds;
  // round 149's browser died 245 seconds in. Neither fact was archived.
  const dr = rounds[rounds.length - 1]?.driverRun;
  if (dr && Number(dr.attempts) > 1)
    console.log(
      `
  THIS ROUND NEEDED ${dr.attempts} ATTEMPTS — recovered from: ${(dr.recovered ?? []).join(", ") || "unrecorded"}
    A round the driver had to rescue is evidence taken from a host that was already unwell.
    Read a scenario failure here against that, not against a clean round.`,
    );

  // WHAT IT TOOK TO GET THE ROUNDS AT ALL, pooled. One round's `driverRun` says
  // whether THAT round was rescued; this says whether rescuing is normal, and
  // names what from — `pane-stale` after a deploy is a property of how rounds
  // are run, `host-silent` is host health, and they were one word until now.
  // WHAT THIS ROUND RECORDED THAT NOTHING READS. Twice now a signal has sat in
  // every round file for months before someone noticed it by hand —
  // `poolFallbackRates` and `poolInPlaceUpdates` are both that story. Most of
  // this list is noise and should stay unread; the point is that the next one
  // does not have to be found by scrolling.
  try {
    const src =
      readFileSync(new URL("./triage.mjs", import.meta.url), "utf8") +
      readFileSync(new URL("./rounds-gate.mjs", import.meta.url), "utf8");
    const unread = unreadSignals(rounds[rounds.length - 1], src);
    if (unread.length) {
      console.log(`
  RECORDED, AND READ BY NOTHING — the busiest signals no tool matches on`);
      for (const u of unread) console.log(`      ${String(u.n).padStart(4)}x  ${u.message}`);
      console.log("    A FLOOR, not a count: a matcher built by concatenation reads as unread here.");
    }
  } catch {
    /* the gate is not worth failing over its own footnote */
  }

  // DID ANY CHART LAND OFF THE SLIDE, OR ON ANOTHER ONE? Unanswerable from the
  // archive until 2026-09-01, because the inventory carried origins and no
  // extent. Reported on the NEWEST round rather than pooled: it is a question
  // about the deck this round left behind, and older rounds cannot answer it at
  // all — which the unmeasured count says out loud rather than passing quietly.
  const geom = newest ? deckGeometryFaults(newest) : null;
  if (geom && (geom.measured || geom.offSlide.length || geom.collisions.length)) {
    console.log(`
  WHERE THE CHARTS LANDED — ${geom.measured} shape(s) measured on the newest round`);
    for (const o of geom.offSlide.slice(0, 6))
      console.log(`      OFF THE SLIDE  ${o.off.padEnd(12)} shape ${o.id} at ${o.box} (slide ${o.slideId})`);
    for (const c of geom.collisions.slice(0, 6))
      console.log(`      TWO CHARTS OVERLAP  ${c.a} and ${c.b} share ${c.area}pt² (slide ${c.slideId})`);
    if (!geom.offSlide.length && !geom.collisions.length)
      console.log("      every chart group inside the slide, none overlapping another");
    if (geom.unmeasured)
      console.log(`      ${geom.unmeasured} shape(s) carried no size — not measured, and not counted as clean`);
  }

  // THE HALF NOTHING GATED. This file compares scenario verdicts and slide-size
  // divergence and, until 2026-09-01, contained no reference to `hostAnswers` at
  // all — while 14 of 15 scenarios pass in every round and seven have never
  // failed in 322. Two runs of the SAME COMMIT disagree on 10.4% of probe slots,
  // and nothing looked.
  const flips = probeFlipsWithinBuild(rounds);
  if (flips.differing) {
    console.log(`
  THE SAME BUILD ANSWERED DIFFERENTLY — ${flips.differing} of ${flips.slots} probe slot(s) across ${flips.builds} build(s) with a pair`);
    for (const f of flips.flips.slice(0, 8))
      console.log(`      ${f.id.padEnd(44)} ${f.answers.join(" / ")}   (${f.builds} build(s))`);
    if (flips.flips.length > 8) console.log(`      … and ${flips.flips.length - 8} more`);
    // DERIVED AGAINST DECLARED. `UNSTABLE_ANSWERS` is hand-written, which is
    // exactly why most of these are missing from it.
    const undeclared = flips.flips.filter((f) => !(f.id in UNSTABLE_ANSWERS));
    if (undeclared.length)
      console.log(
        `    ${undeclared.length} of ${flips.flips.length} are not in UNSTABLE_ANSWERS: ${undeclared
          .slice(0, 6)
          .map((f) => f.id)
          .join(", ")}${undeclared.length > 6 ? ", …" : ""}`,
      );
    console.log("    Some of these were answer-ranking defects, not the host — re-derive, do not just declare.");
  }

  // THE REGISTER CANNOT AUDIT ITSELF AGAINST THE FIXTURE, so it is audited here
  // against the archive. `PENDING_QUESTIONS` says to delete an entry once the
  // host answers, and its own gate compares to the committed sheet — where the
  // id is legitimately absent precisely BECAUSE the fixture predates it. Green
  // whether the question is unanswered or answered fifty times. The archive can
  // tell.
  const answered = pendingAlreadyAnswered(rounds);
  if (answered.length) {
    console.log(`
  DECLARED UNANSWERED, BUT THE ARCHIVE ANSWERS IT — ${answered.length} question(s)`);
    for (const q of answered)
      console.log(
        `      ${q.id} — ${q.answers.map((a) => `${a.answer} in ${a.n}`).join(", ")} of ${q.rounds} round(s) that asked`,
      );
    console.log("    Refresh test/fixtures/host-answers-web.json from a recent round, then delete the entry.");
  }

  const starts = poolDriverRuns(rounds);
  if (starts.rounds) {
    console.log(`
  WHAT IT TOOK TO START — ${starts.clean} of ${starts.rounds} round(s) started first time`);
    for (const c of starts.causes.slice(0, 6))
      console.log(`      ${String(c.n).padStart(3)}x  recovered from ${c.cause}`);
    console.log("    Counts, not a rate: driverRun is newer than most of the archive.");
    // CRASHES PER ATTEMPT, BY SLIDE SIZE. Every input to this has been archived
    // for months and nothing divided one by the other; `docs/BACKLOG.md` says as
    // much in as many words. The gap it prints is sixteen-fold, and invisible
    // until something printed it.
    const s0Crashes = (starts.bySize ?? []).reduce((n, s) => n + s.crashes, 0);
    if (starts.bySize?.length > 1) {
      console.log("    crashes per ATTEMPT, by slide size — the arm that is failing, not the round count:");
      for (const s of starts.bySize)
        console.log(
          `      ${s.size.padEnd(11)} ${String(s.crashes).padStart(3)} crash(es) in ${String(s.attempts).padStart(4)} attempt(s)` +
            ` = ${((100 * s.crashes) / Math.max(1, s.attempts)).toFixed(1).padStart(5)}%` +
            `   (${s.roundsWithCrash} of ${s.rounds} round(s) hit one)`,
        );
      console.log("      Aspect ratio and deck file are CONFOUNDED — cyclePlan has never crossed them.");
      // AND THE DENOMINATOR IS CONDITIONAL ON RECOVERY, which the rate above
      // cannot say for itself. A round the driver never recovered writes no
      // round file at all — `round.mjs` states the mechanism outright: "A
      // crashed round archives nothing: it never reaches the download button."
      // So `rounds/` holds only the crashes that were survived, and every crash
      // rate computed from it is a FLOOR. The crash reports are the other half
      // of the record, and they are on disk already.
      const reports = countCrashReports();
      if (reports > s0Crashes)
        console.log(
          `      FLOOR, not a rate: crashes/ holds ${reports} report(s) against ${s0Crashes} archived event(s)` +
            ` — ${Math.round((100 * (reports - s0Crashes)) / reports)}% of crashes left no round file to count.`,
        );
    }
    // THE ARM, SPLIT ON THE PANE RATHER THAN ON THE FLAG. Round 166 ran without
    // `--fresh` and started on a 69-second pane anyway, because a merge preceded
    // it. The flag is one way to get a fresh pane, not the variable itself, and
    // splitting on it filed a fresh-pane round under "aged".
    const a = starts.arms;
    if (a && a.fresh + a.aged > 0) {
      console.log(
        `    fresh pane ${a.freshRefusedNone}/${a.fresh} round(s) refused no group · ` +
          `aged pane ${a.agedRefusedNone}/${a.aged}`,
      );
      if (a.flagDisagreed)
        console.log(
          `      --fresh disagreed with the pane in ${a.flagDisagreed} of them — the flag is not the variable`,
        );
    }
  }

  const fb = poolFallbackRates(rounds);
  if (fb.length) {
    console.log(`
  FALLBACKS TAKEN — this round against the median of ${fb[0].rounds} prior round(s)`);
    for (const r of fb) {
      // DRIFT FIRST, because it is the reading a median cannot give. A signal
      // that climbs steadily looks NORMAL against its own history the whole way
      // up: `in-place update fell back to a redraw` went from 9 to 13 per round
      // across sixty rounds, and by the time anyone looked, "now" and "usually"
      // were both 13. The oldest third against the newest third sees the shape
      // a median absorbs.
      const rising = r.newest > r.oldest * 1.3 && r.newest - r.oldest >= 2;
      const falling = r.oldest > r.newest * 1.3 && r.oldest - r.newest >= 2;
      // AND WHEN THE LAST FEW ROUNDS SAY OTHERWISE, SAY SO. A thirds reading is
      // still a summary over 44 rounds, so a step inside the newest third is
      // invisible to it — and it goes on asserting a direction after that
      // direction has reversed. This row printed `now 2` beside `RISING, 8 to
      // 13` on a signal that had read 2 for five rounds running: a conclusion
      // that outlived its evidence, which is the defect this whole gate exists
      // to catch elsewhere.
      const med = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
      const last = r.recent?.length ? med(r.recent) : null;
      const stale =
        last !== null && ((rising && last < r.oldest) || (falling && last > r.oldest))
          ? ` — but the last ${r.recent.length} rounds median ${last}, so the thirds are describing an era that ended`
          : "";
      const drift = rising
        ? `  <- RISING, ${r.oldest} to ${r.newest} across ${r.span}-round thirds${stale}`
        : falling
          ? `  <- falling, ${r.oldest} to ${r.newest}${stale}`
          : "";
      // THE SEQUENCE BESIDE THE SUMMARIES, because a median cannot see a step
      // and this row proved it: `now 2` printed next to `usually 12` and
      // `RISING, 8 to 13`, on a signal that has read 2 for five rounds running.
      // All three summaries were describing an era that had ended.
      const seq = r.recent?.length ? `  [${r.recent.join(",")}]` : "";
      console.log(`    ${r.label.padEnd(38)} ${String(r.now).padStart(4)}  (usually ${r.median})${drift}${seq}`);
    }
  }

  const pos = poolPairPosition(rounds);
  if (pos.pairs >= 4) {
    const moved = pos.worse + pos.better;
    console.log(
      `
  PAIR POSITION — over ${pos.pairs} build(s) run twice: the SECOND round was worse ` +
        `${pos.worse}x, better ${pos.better}x, unchanged ${pos.tied}x`,
    );
    if (moved > 0 && pos.worse > pos.better * 2)
      console.log(
        `    ${pos.worse} of the ${moved} pairs that moved went the same way. That is a direction, not a mood.
` +
          `    THE CAUSE IS THE PANE, not the position: a second round INHERITS the first round's
` +
          `    pane, and pane age separates post-retry 0.4 from 4.6. ${pos.secondFresh} of ${pos.pairs} pairs had a
` +
          `    fresh second round — and NOT because the driver freshens one. It does not: reloading the
` +
          `    pane raises a beforeunload prompt over unsaved work, and accepting it has cost the
` +
          `    SIDELOAD twice (rounds 124 and 132). What freshens the pane is a MERGE, which makes it
` +
          `    stale so recovery reloads it. So a pair run properly — nothing merged between the two
` +
          `    rounds, which is the whole point of a pair — is exactly the case that inherits an aged
` +
          `    pane. Measured: rounds 148-158, each after a merge, started at 62-76s; rounds 159 and 161,
` +
          `    the only true second-rounds in that stretch, started at 696s and 666s, and both were the
` +
          `    worse half of their pair. Run the second round with --fresh: it closes the BROWSER, which
` +
          `    the persistent profile survives, instead of reloading the page that holds unsaved work.`,
      );
  }

  const disagreed = poolProfileDisagreements(rounds);
  if (disagreed.length) {
    console.log(`
  SLIDE SIZE DISAGREES — ${disagreed.length} round(s) filed under a profile the driver did not measure`);
    for (const d of disagreed)
      console.log(`    ${d.build}  archive says ${d.pane} (from ${d.source}), the driver measured ${d.driver}`);
    console.log("    Every profile comparison below groups by the ARCHIVE's value. Treat these rounds as unfiled.");
  }

  const grouping = poolGroupingOutcome(rounds);
  if (grouping) {
    const { now, refusedMedian, rounds: priorRounds, attempts, recent } = grouping;
    console.log(
      `  GROUPING, which no scenario verdict reports: ${now.grouped} of ${attempts} attempt(s) grouped, ` +
        // NO BASELINE IS NOT A BASELINE OF ZERO. This used to print `usually 0`
        // when there was no history at all, which reads as "clean until now".
        (refusedMedian === null
          ? `${now.refused} refused (no baseline — ${priorRounds} prior round(s) is too few to say what is usual)`
          : `${now.refused} refused (usually ${refusedMedian} over ${priorRounds} prior round(s))`),
    );
    // THE THROW, NAMED. Grouping has three outcomes and this line counted two,
    // so a round where a group threw printed `8 of 8 grouped, 0 refused` and
    // read as perfect — the missing attempt WAS the defect. 183 throws across
    // 65 of 150 rounds, and they match the loose-shape slides exactly: 159,
    // 161, 167 and 174 threw once each and ended holding 17, 17, 11 and 11
    // shapes, while the twelve rounds between them threw none and ended at 5.
    if (now.threw)
      console.log(
        `    ${now.threw} group(s) THREW — a chart whose group throws is left as its shapes; see the fullest-slide line`,
      );
    // THE POPULATION, BESIDE THE COUNT. Attempts per round ran 15-20 for the
    // whole archive and halved to 9 at round 153 — benign (the in-place update
    // started working, and a chart that is not redrawn is never regrouped) and
    // completely invisible, while silently rebasing every grouping figure in
    // triage. "0 refused (usually 2)" reads as an improvement when half the
    // attempts stopped happening.
    if (recent?.length) console.log(`    attempts per round, last ${recent.length}: [${recent.join(",")}]`);
    console.log(`    the deck ended holding ${now.deck.join(",")} shape(s) per slide`);
    // AND WHETHER THAT IS NORMAL, which nothing has ever said. A clean round's
    // fullest slide holds five; eight of the last thirty rounds ended with one
    // holding 11 to 48, and every one of them reported 13 of 13. A chart that
    // fails to group is left as its loose shapes, and no scenario verdict looks
    // at the deck — `does a rasterise poison the next draw` asks whether the
    // CALL came back, not whether a chart survived, so it passes with eight
    // loose shapes sitting where a chart should be.
    //
    // Printed as a sequence and never as a failure: the gate's own rule is that
    // it does not cry wolf on a host whose mood swings, and this is a reason to
    // read the round rather than a verdict on the build.
    const fullest = poolFullestSlide(rounds);
    if (fullest.length) {
      const over = fullest.filter((n) => n > CLEAN_SLIDE_CEILING).length;
      console.log(
        `    fullest slide per round, last ${fullest.length}: [${fullest.join(",")}]` +
          (over ? `  <- ${over} above ${CLEAN_SLIDE_CEILING}, so a chart was left as loose shapes` : ""),
      );
    }
    if (now.refused > 0)
      console.log(
        [
          "    A refused chart is left as loose shapes in its own box — it keeps its config",
          "    and looks identical to the scenario, which is why 13/13 can hide it. Inside the",
          "    noise floor unless a PAIR on one build agrees; read the deck line above.",
        ].join("\n"),
      );
  }
  const shrunk = poolScenarioPopulations(rounds);
  if (shrunk.length) {
    console.log(`  ${shrunk.length} scenario(s) PASSED ON A SMALLER POPULATION than they usually run:`);
    for (const p of shrunk)
      console.log(
        `    ${p.name} — ${p.now} this round, usually ${p.usual} over ${p.rounds} prior round(s)` +
          `${p.ok ? " (and it still reports PASS)" : ""}`,
      );
    console.log("  A ratio whose bottom half moved is not the same evidence. Read why before comparing it.");
  }
  // A THIRD QUESTION, and the cheapest of the three to answer wrongly. The two
  // above ask about scenarios — thirteen named outcomes a person already reads.
  // This asks about the TRACE, which is 95K characters nobody can count by eye,
  // and its whole job is to say which parts of it are worth the reading.
  //
  // Never fatal, for the same reason divergence is not: it reports difference,
  // and a build is not broken for being different. Exiting non-zero here would
  // make every landed fix fail the gate on the night it starts working.
  const nov = traceNovelty(rounds);
  if (nov.novel.length) {
    console.log(`  ${nov.novel.length} trace signature(s) NEVER SEEN in ${nov.priors} prior round(s):`);
    for (const s of nov.novel.slice(0, 8)) console.log(`    ${String(s.n).padStart(4)}x  ${s.sig}`);
    console.log("  Read the trace. A shape this archive has never produced is the reason to.");
  }
  if (nov.sinceBuild.length) {
    console.log(
      `  ${nov.sinceBuild.length} signature(s) are NEW BEHAVIOUR, not a spike (absent recently, common now):`,
    );
    // THE BUILD IT STARTED IN, per signature. This printed the build being
    // judged, for every entry — so one 064-era signature was blamed on nine
    // consecutive innocent commits, the newest of them a slide-counter fix.
    for (const s of nov.sinceBuild.slice(0, 8))
      console.log(`    ${String(s.n).padStart(4)}x  ${s.sig}${s.startedIn ? `  (first seen in ${s.startedIn})` : ""}`);
    console.log("  The shape a mechanism makes when it starts working — check it is one that build widened.");
  }
  if (nov.spikes.length) {
    console.log(`  ${nov.spikes.length} signature(s) SPIKED against their own history:`);
    for (const s of nov.spikes.slice(0, 8))
      console.log(`    ${String(s.n).padStart(4)}x (usually ${s.median})  ${s.sig}`);
    console.log("  These had a baseline and left it. Most likely to be the round's real story.");
  }
  if (!nov.novel.length && !nov.sinceBuild.length && !nov.spikes.length)
    console.log(
      `  nothing new in the trace — ${nov.vocabulary} known signature(s) across ${nov.priors} prior round(s).` +
        " The trace is still the evidence; this only says where to start.",
    );
  if (!gone.length) {
    console.log(
      `  no scenario regressed — checked the newest of ${rounds.length} archived round(s)` +
        ` at ${roundProfile(rounds[rounds.length - 1])}`,
    );
    process.exit(0);
  }
  console.error(`  ${gone.length} scenario(s) STOPPED PASSING in the newest round:`);
  // WHETHER THE PRODUCT COULD POSSIBLY BE RESPONSIBLE, asked of git rather than
  // left to the reader. A round whose SHIPPED BUNDLE is byte-identical to the
  // rounds it is judged against cannot have regressed the product — a verdict
  // that moved across identical `src/` is the host, definitionally.
  //
  // Round 273 is why this exists. `two slides claiming one slot` stopped passing
  // and the cycle halted on its one fatal check, correctly. The build under
  // judgement changed `scripts/round.mjs`, a test, and an archive file, and
  // nothing under `src/` at all — so the bundle was the same one four previous
  // rounds had passed on. Establishing that took a git diff anyone could have
  // run and nobody was prompted to.
  const sameProfileRounds = rounds.filter((r) => roundProfile(r) === roundProfile(rounds[rounds.length - 1]));
  const bundleMoved = bundleChanged(
    sameProfileRounds[sameProfileRounds.length - 2]?.build,
    rounds[rounds.length - 1]?.build,
  );
  for (const g of gone) {
    console.error(
      `    ${g.name} — ${g.failed === 1 ? "FIRST failure" : `failed ${g.failed} times`} in ${g.ran} round(s) at this profile`,
    );
    // IS TONIGHT'S FAILURE ONE OF THE FAMILIAR ONES, or a new shape under an
    // old name? Those want opposite responses, and the count above cannot tell
    // them apart. Round 435 said character-for-character what rounds 060, 253
    // and 273 had said — and 273 is the round the bundle guard below exists
    // for. Everything needed to recognise it was already in the archive.
    const seen = g.sameDetailIn ?? [];
    console.error(
      seen.length
        ? `      this exact failure text has appeared ${seen.length} time(s) before — ${seen.slice(0, 6).join(", ")}`
        : "      this failure text is NEW — no earlier round at this profile failed with these words",
    );
  }
  if (bundleMoved === false)
    console.error(
      "  THE SHIPPED BUNDLE IS UNCHANGED since the previous round at this profile — nothing under\n" +
        "  `src/` differs between the two builds. Whatever moved, the product did not: read this as\n" +
        "  the host. Still exit 1, because a scenario falling is worth a person's eyes either way.",
    );
  // THE THIRD VALUE, printed as a fact rather than as the verdict above it.
  // `src/` moved but no line that runs did — round 440's 4:3 leg failed against
  // 27 lines of comment. A reader can draw the same conclusion from this; the
  // heuristic is not allowed to draw it for them, because "unchanged" said
  // wrongly excuses a regression while "changed" said wrongly costs a diff.
  else if (bundleMoved === "comments")
    console.error(
      "  `src/` changed since the previous round at this profile, but ONLY IN COMMENTS AND BLANK\n" +
        "  LINES — no line that runs moved. Read that as a fact, not as a verdict: anything this\n" +
        "  check cannot plainly see as a comment is counted as code, on purpose.",
    );
  console.error("  A round is evidence; this is the only thing that holds a build to it. See docs/ROUNDS.md.");
  process.exit(1);
}
