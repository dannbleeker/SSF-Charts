/**
 * WHICH LINES TO MUTATE THIS WEEK, so the run can finish at all.
 *
 * THE WHOLE-ENGINE RUN CANNOT COMPLETE, and that is measured rather than
 * feared. GitHub's hard job ceiling is 6 hours and the mutation job has been
 * killed at it four times — 2026-08-10, 08-17, 08-24 and 2026-09-14 at
 * 6h00m29s — producing ZERO reports in its entire life. Monday's run left the
 * numbers behind:
 *
 *     elapsed: ~5h 49m, remaining: ~144h 26m
 *     3254/18724 tested (671 survived, 21 timed out)
 *
 * 3,254 mutants in 5h49m is **559 per hour**, so the full set needs about
 * **33 hours** at that rate and Stryker's own adaptive estimate said 144. It is
 * not a run that needs one more timeout raised; it is 5-25x over budget. (The
 * four stacked timeouts fixed on 2026-09-13 got the DRY RUN to complete in
 * 12m07s. Necessary, and nowhere near sufficient.)
 *
 * AND THE OUTPUT WOULD BE UNREADABLE EVEN IF IT FINISHED. 671 of 3,254 tested
 * survived — 20.6% — which extrapolates to roughly **3,900 survivors**. This
 * repo already knows what happens to a report that size: `office-js-watch.mjs`
 * says "Without it the sweep reports the same twenty issues forever and is
 * ignored within a month."
 *
 * So the run is scoped to the lines that CHANGED, which is the same rule both
 * sibling jobs in `quality-sweep.yml` already follow — look every week, and say
 * only what is new. A quiet week mutates nothing and the job exits after the
 * dry run; a busy week mutates a few hundred.
 *
 * Stryker supports this directly: `mutate` entries may carry a mutation range,
 * `path:startLine[:startColumn]-endLine[:endColumn]`, documented in its own
 * schema. There is no `--since` flag in Stryker 9; the range syntax is the
 * mechanism.
 *
 * Usage:  node scripts/mutation-scope.mjs [--days 7] [--since <ref>]
 * Output: one comma-separated `--mutate` value on stdout, or NOTHING when no
 *         mutated file changed. Exit is always 0 — "nothing changed" is an
 *         answer, not a failure.
 */

import { execFileSync } from "node:child_process";
import { isMain } from "./is-main.mjs";

/** Only these are mutated — kept in step with `stryker.config.json`'s `mutate`. */
export const MUTATED = /^src\/core\/.*\.ts$/;
export const NOT_MUTATED = [/^src\/core\/types\.ts$/, /^src\/core\/samples\.ts$/];

/**
 * Parse `git diff --unified=0` into `path:start-end` ranges.
 *
 * ONLY THE `+` SIDE OF EACH HUNK. A mutation range says which lines of the
 * CURRENT file to mutate, so the `-` counts (what the old file had) are the
 * wrong number and a deletion-only hunk contributes nothing to mutate. The
 * header is `@@ -oldStart,oldCount +newStart,newCount @@`, and a missing count
 * means 1 — `@@ -5 +7 @@` is one line, not zero.
 */
export function rangesFromDiff(diff) {
  const out = [];
  let file = null;
  for (const line of String(diff ?? "").split("\n")) {
    const f = /^\+\+\+ b\/(.+)$/.exec(line);
    if (f) {
      file = f[1] === "dev/null" ? null : f[1];
      continue;
    }
    if (!file) continue;
    const h = /^@@ -\d+(?:,\d+)? \+(\d+)(?:,(\d+))? @@/.exec(line);
    if (!h) continue;
    const start = Number(h[1]);
    const count = h[2] === undefined ? 1 : Number(h[2]);
    if (!count) continue; // a pure deletion changes no current line
    out.push({ file, start, end: start + count - 1 });
  }
  return out;
}

/** Keep only what Stryker is configured to mutate. */
export function onlyMutated(ranges) {
  return ranges.filter((r) => MUTATED.test(r.file) && !NOT_MUTATED.some((x) => x.test(r.file)));
}

/**
 * Merge touching or overlapping ranges in the same file.
 *
 * Two hunks three lines apart would otherwise become two `mutate` entries over
 * nearly the same code, and Stryker would instrument the file twice. Merging
 * with a small gap also catches the case a diff splits one edited function into
 * several hunks, where mutating the gap is what you wanted anyway.
 */
export function mergeRanges(ranges, gap = 5) {
  const byFile = new Map();
  for (const r of ranges) {
    if (!byFile.has(r.file)) byFile.set(r.file, []);
    byFile.get(r.file).push({ ...r });
  }
  const out = [];
  for (const [file, list] of [...byFile.entries()].sort()) {
    list.sort((a, b) => a.start - b.start);
    let cur = null;
    for (const r of list) {
      if (cur && r.start <= cur.end + gap) cur.end = Math.max(cur.end, r.end);
      else {
        if (cur) out.push(cur);
        cur = { file, start: r.start, end: r.end };
      }
    }
    if (cur) out.push(cur);
  }
  return out;
}

/** `src/core/chart.ts:120-160` — the spelling Stryker's `mutate` accepts. */
export const toMutateArg = (ranges) => ranges.map((r) => `${r.file}:${r.start}-${r.end}`).join(",");

function main(argv) {
  const days = Number(argv[argv.indexOf("--days") + 1]) || 7;
  const sinceArg = argv.includes("--since") ? argv[argv.indexOf("--since") + 1] : null;
  const git = (...a) => execFileSync("git", a, { encoding: "utf8", maxBuffer: 64e6 });

  // A COMMIT, NOT `@{7.days.ago}`. The reflog form resolves against a LOCAL
  // reflog that a fresh CI clone does not have, and silently yields HEAD — so
  // the diff would be empty and the job would report "nothing changed" every
  // single week while looking perfectly healthy.
  let since = sinceArg;
  if (!since) {
    since = git("rev-list", "-1", `--before=${days} days ago`, "HEAD").trim();
  }
  if (!since) {
    // No commit is older than the window: the whole history is inside it.
    since = git("rev-list", "--max-parents=0", "HEAD").trim().split("\n")[0];
  }
  const diff = git("diff", "--unified=0", `${since}..HEAD`, "--", "src/core");
  const ranges = mergeRanges(onlyMutated(rangesFromDiff(diff)));
  if (ranges.length) process.stdout.write(toMutateArg(ranges));
}

if (isMain(import.meta.url, process.argv[1])) {
  main(process.argv.slice(2));
}
