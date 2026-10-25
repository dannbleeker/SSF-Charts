#!/usr/bin/env node
/**
 * WAIT FOR *THIS* COMMIT TO BE LIVE — CI green, then Pages deployed.
 *
 * Written because I got it wrong four times in one afternoon, the same way each
 * time: `gh run list ... | select(.name == "Deploy Pages") | first` picks the
 * NEWEST run, which right after a push is still the PREVIOUS commit's. Watching
 * that one finish reports success for a deploy that never happened, and the
 * round then runs against stale code — or refuses, which is how it was finally
 * caught: `the site is serving 55630e7 but HEAD is acb5a9e`.
 *
 * That is worse than not checking at all. A missing check leaves you unsure; a
 * check keyed to the wrong row manufactures evidence.
 *
 * THE SAME MISTAKE, ONE LEVEL DOWN, WAS STILL HERE (r3-79). The fix for the one
 * above matched the run on `headSha`, and that column is right for CI and WRONG
 * for Deploy Pages. Since 2026-09-11 `pages.yml` is triggered by `workflow_run`,
 * and for that event GitHub sets a run's `head_sha` to the last commit on main,
 * not to the commit the run was started for; the build job checks out
 * `workflow_run.head_sha` of the CI run instead and so deploys a different
 * commit. On 2026-09-27 a second push landed before the first one's deploy, CI
 * for the second finished first, and the two Deploy Pages runs both carried the
 * second commit's `headSha` while the later one published the first. A tool
 * keyed to that column printed "is live" for the second commit while the site
 * served the first (`docs/PUBLISHING.md`, "AND THE WAIT CAN NEVER END").
 *
 * So the run is identified by the commit it CHECKED OUT, which the build log
 * states (`HEAD is now at <sha>`), and not by any column the run list offers. A
 * hand-started run (`workflow_dispatch`) checks out the ref it was started on,
 * and there `headSha` is that commit. "Live" then means the most recent
 * successful deploy built this commit and none is still in flight: a later one
 * lands on top of an earlier one, and a tool that stops at the first success
 * says "live" about a site that is about to move.
 *
 * CI IS A PRECONDITION OF THE DEPLOY, and this file used to say the opposite.
 * Its header read "Deploys and CI are independent" and its exit message "It may
 * well be LIVE — deploys do not wait for CI", true until 2026-09-11 (`63f46f6`)
 * and false since: `pages.yml` builds only when the CI run it follows CONCLUDED
 * `success`, so a commit that fails CI is never deployed and the site keeps the
 * previous one. CI is still waited for first, now because nothing can be live
 * before it passes, and a round against a build the suite rejects produces
 * numbers everyone will trust and a suite nobody has satisfied.
 *
 * Usage:  node scripts/await-deploy.mjs [--timeout-min 30]
 * Exit:   0 live · 1 CI failed · 2 deploy failed or replaced · 3 timed out
 */
import { execFileSync } from "child_process";
import { isMain } from "./is-main.mjs";

const POLL_MS = 15_000;

/**
 * Clock slack on "created after CI finished", not a window. GitHub makes the
 * Deploy Pages run about two seconds after the CI run concludes, and a run
 * created earlier than that cannot have been started by it.
 */
const SKEW_MS = 10_000;

const FIELDS = "databaseId,name,headSha,event,status,conclusion,createdAt,updatedAt";

/** The workflow run for THIS commit, or null while GitHub has not made one yet. */
export function ciRunFor(rows, sha) {
  return rows.find((r) => r.name === "CI" && r.headSha === sha) ?? null;
}

/** Where CI stands: still going, passed, or ended some other way. */
export function judgeCi(row) {
  // Absent is NOT failed. GitHub takes a few seconds to create the run, and
  // treating "no row yet" as a failure would make this flap on every push.
  if (!row) return { state: "waiting", note: "(not created yet)" };
  if (row.status !== "completed") return { state: "waiting", note: `${row.status} ${row.conclusion ?? ""}`.trim() };
  return row.conclusion === "success"
    ? { state: "success", note: "success" }
    : { state: "failed", note: `failed: ${row.conclusion}` };
}

/**
 * The commit a Deploy Pages run BUILT, as a (possibly abbreviated) sha, or null
 * when that cannot be told.
 *
 * `readLog` is only asked for a `workflow_run` run, where `headSha` is main's tip
 * and says nothing about what was checked out. The build job's log carries it:
 * `HEAD is now at bbfcf5b Merge: ...` is what `git checkout <sha>` prints.
 */
export function builtShaOf(row, readLog) {
  if (row.event !== "workflow_run") return row.headSha ?? null;
  return /HEAD is now at ([0-9a-f]{7,40})\b/.exec(readLog(row.databaseId) ?? "")?.[1] ?? null;
}

/**
 * Has `sha` been deployed? One of waiting, live, failed, or replaced.
 *
 * Only runs created after CI for `sha` concluded are considered (`since`): the
 * deploy for a commit cannot exist before the run that triggers it. Of those,
 * a run that was skipped (CI red, or a fork's pull request) or cancelled
 * published nothing and is left out; a run still in flight might yet publish
 * over this commit, so it holds the answer at "waiting".
 */
export function judgeDeploy({ rows, sha, since, builtOf }) {
  const mine = rows.filter(
    (r) =>
      r.name === "Deploy Pages" &&
      (r.event === "workflow_run" || r.event === "workflow_dispatch") &&
      Date.parse(r.createdAt) >= since,
  );
  const inFlight = mine.filter((r) => r.status !== "completed");
  if (inFlight.length)
    return { state: "waiting", note: `Deploy Pages run ${inFlight.map((r) => r.databaseId).join(", ")} in flight` };

  const built = mine
    .filter((r) => r.conclusion === "success" || r.conclusion === "failure")
    .map((r) => ({ run: r, sha: builtOf(r) }));
  // A run whose commit cannot be read is not ours and not somebody else's, so
  // nothing can be said yet. Saying "live" on it is the manufactured evidence.
  const unread = built.filter((b) => b.sha === null);
  if (unread.length)
    return {
      state: "waiting",
      note: `could not read which commit run ${unread.map((b) => b.run.databaseId).join(", ")} built`,
    };

  const ours = built.filter((b) => sha.startsWith(b.sha));
  if (!ours.length) return { state: "waiting", note: "no Deploy Pages run has built it yet" };

  const published = built
    .filter((b) => b.run.conclusion === "success")
    .sort((a, b) => Date.parse(a.run.updatedAt) - Date.parse(b.run.updatedAt));
  const latest = published.at(-1);
  if (latest && sha.startsWith(latest.sha)) return { state: "live", note: `run ${latest.run.databaseId}` };
  if (ours.some((b) => b.run.conclusion === "success"))
    return { state: "replaced", note: `run ${latest.run.databaseId} then published ${latest.sha} on top of it` };
  return { state: "failed", note: `failed: run ${ours.at(-1).run.databaseId} ${ours.at(-1).run.conclusion}` };
}

const ghDefault = (args) => execFileSync("gh", args, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });

/**
 * Wait for `sha`. Returns the exit code; prints as it goes.
 *
 * Everything that touches the world is a parameter (`gh`, `sleep`, `now`, the
 * two writers) so the whole wait can be driven by a test without a network.
 */
export async function awaitDeploy({
  sha,
  minutes = 30,
  gh = ghDefault,
  sleep = (ms) => new Promise((r) => setTimeout(r, ms)),
  now = Date.now,
  say = console.log,
  warn = console.error,
  tick = (s) => process.stdout.write(s),
}) {
  const short = sha.slice(0, 7);
  const deadline = now() + minutes * 60_000;
  const rowsNow = () => JSON.parse(gh(["run", "list", "--limit", "50", "--json", FIELDS]));

  say(`  waiting for ${short} to go live`);
  let ci = null;
  let lastSeen = "(not created yet)";
  while (now() < deadline) {
    const row = ciRunFor(rowsNow(), sha);
    const verdict = judgeCi(row);
    lastSeen = verdict.note;
    if (verdict.state !== "waiting") {
      ci = row;
      break;
    }
    tick(`  CI: ${lastSeen}\r`);
    await sleep(POLL_MS);
  }
  if (!ci || judgeCi(ci).state !== "success") {
    // A TIMEOUT IS NOT A FAILURE, and saying so was this tool's own version of the
    // conflation it exists to prevent. The exit codes already told them apart (3
    // against 1) while the message called both "did not pass CI" — and the first
    // time it fired, CI was still running: GitHub had taken seventeen minutes to
    // even create the run, against a three-to-four minute norm.
    if (!ci) {
      say(`  CI: timed out (last seen: ${lastSeen})`.padEnd(60));
      warn(`  ${short}: CI has not finished yet — timed out (last seen: ${lastSeen}).`);
      warn("  Nothing is wrong with the build; give it longer with --timeout-min.");
      return 3;
    }
    say(`  CI: ${judgeCi(ci).note}`.padEnd(60));
    warn(`  ${short} did not pass CI, so it was never deployed: pages.yml builds only after the CI run it`);
    warn("  follows concludes success, and the site still serves an earlier commit. Fix CI and push again;");
    warn("  waiting longer changes nothing.");
    return 1;
  }
  say("  CI: success".padEnd(60));

  // Unknown completion time means every recent run is a candidate, which the
  // log check below still sorts out; it must not mean none is.
  const finished = Date.parse(ci.updatedAt);
  const since = Number.isFinite(finished) ? finished - SKEW_MS : 0;
  const readLog = (id) => {
    try {
      return gh(["run", "view", String(id), "--log"]);
    } catch {
      return null;
    }
  };
  // A completed run never changes what it built, so each log is read once. A
  // log that cannot be read is not cached: it may be there next poll.
  const builtCache = new Map();
  const builtOf = (row) => {
    const known = builtCache.get(row.databaseId);
    if (known) return known;
    const found = builtShaOf(row, readLog);
    if (found) builtCache.set(row.databaseId, found);
    return found;
  };

  lastSeen = "(not created yet)";
  while (now() < deadline) {
    const verdict = judgeDeploy({ rows: rowsNow(), sha, since, builtOf });
    lastSeen = verdict.note;
    if (verdict.state === "live") {
      say(`  Deploy Pages: success (${verdict.note})`.padEnd(60));
      say(`  ${short} is live`);
      return 0;
    }
    if (verdict.state === "failed" || verdict.state === "replaced") {
      say(`  Deploy Pages: ${verdict.note}`.padEnd(60));
      if (verdict.state === "replaced") {
        warn(`  ${short} was deployed and then replaced: ${verdict.note}.`);
        warn("  A deploy of a commit whose CI finished later lands on top of this one. Re-run the deploy by hand:");
        warn('  `gh workflow run "Deploy Pages" --ref main` (docs/PUBLISHING.md).');
      }
      return 2;
    }
    tick(`  Deploy Pages: ${lastSeen}\r`);
    await sleep(POLL_MS);
  }
  say(`  Deploy Pages: timed out (last seen: ${lastSeen})`.padEnd(60));
  return 3;
}

if (isMain(import.meta.url, process.argv[1])) {
  const argv = process.argv.slice(2);
  const at = argv.indexOf("--timeout-min");
  // THIRTY, not twenty. CI itself runs in three to four minutes, but GitHub can
  // leave a run uncreated for far longer — seventeen minutes on 2026-08-26 — and a
  // timeout tuned to the job rather than to the queue reports trouble that is not
  // there.
  const minutes = at === -1 ? 30 : Number(argv[at + 1]) || 30;
  const sha = execFileSync("git", ["rev-parse", "HEAD"], { encoding: "utf8" }).trim();
  process.exit(await awaitDeploy({ sha, minutes }));
}
