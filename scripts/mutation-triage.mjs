/**
 * WHICH SURVIVING MUTANTS NOBODY HAS RULED ON YET.
 *
 * A surviving mutant is a line this repo changed and no test noticed — which is
 * the question `stryker.config.json` says the job exists to ask: "Which
 * assertions are decorative? ... Three guards in one session turned out to pass
 * against the code they were written to catch. Nobody has ever checked the other
 * two thousand."
 *
 * **`KNOWN_SURVIVORS` IS THE LOAD-BEARING HALF**, exactly as `KNOWN_ISSUES` is
 * for the office-js sweep, and for the reason that file states outright:
 * without it "the sweep reports the same twenty issues forever and is ignored
 * within a month". With it the report is precisely "here is something nobody
 * here has looked at".
 *
 * A surviving mutant is not automatically a defect. Three honest reasons to
 * keep one, and recording WHICH is the point:
 *
 *   equivalent   the mutant cannot change observable behaviour. `>=` to `>` on
 *                a loop bound that a later clamp makes unreachable is the
 *                classic, and `src/core` is full of shrink-to-fit loops.
 *   unasserted   the behaviour is real but deliberately not asserted — a
 *                cosmetic default nothing depends on.
 *   accepted     a genuine gap somebody decided not to close, with the reason.
 *
 * "No exposure" is a legitimate entry and an important one: it records that
 * somebody checked, which is otherwise indistinguishable from nobody looking.
 *
 * Usage:  node scripts/mutation-triage.mjs reports/mutation/mutation.json
 * Exit:   0 nothing new · 3 at least one survivor nobody has ruled on
 *         · 2 the report could not be read
 *
 * 3 IS THE FINDING, matching `flaky.mjs` and `check-published-install.mjs`, so
 * the workflow must wrap the call in `set +e` or the step dies on the one
 * outcome the job exists to produce.
 */

import { readFileSync } from "node:fs";
import { isMain } from "./is-main.mjs";

/**
 * Survivors already ruled on, keyed `file:line:mutatorName`.
 *
 * The value is WHAT WAS DECIDED and why — never a bare "ok". A key met in a
 * future report can then be traced to a decision without re-deriving it.
 *
 * EMPTY ON PURPOSE, 2026-09-17. The job has produced zero reports in its life —
 * four runs killed at GitHub's 6h ceiling and five outright failures — so there
 * is not one real survivor to seed this with. Filling it from guesses would be
 * worse than leaving it bare: every guessed entry silences a finding nobody
 * looked at, which is the exact failure the table exists to prevent.
 */
export const KNOWN_SURVIVORS = {};

/** Stryker's json report -> the surviving mutants, flattened and keyed. */
export function survivorsOf(report) {
  const files = report?.files ?? {};
  const out = [];
  for (const [path, entry] of Object.entries(files)) {
    for (const m of entry?.mutants ?? []) {
      if (m?.status !== "Survived") continue;
      const line = m?.location?.start?.line ?? 0;
      out.push({
        key: `${path}:${line}:${m.mutatorName ?? "?"}`,
        file: path,
        line,
        mutator: m.mutatorName ?? "?",
        replacement: String(m.replacement ?? "").slice(0, 80),
      });
    }
  }
  return out.sort((a, b) => a.key.localeCompare(b.key));
}

/** Per-file kill rate, worst first — the 43-row table a person can act on. */
export function scoreByFile(report) {
  const rows = [];
  for (const [path, entry] of Object.entries(report?.files ?? {})) {
    const mutants = entry?.mutants ?? [];
    const scored = mutants.filter((m) => m?.status !== "Ignored" && m?.status !== "CompileError");
    if (!scored.length) continue;
    const killed = scored.filter((m) => m.status === "Killed" || m.status === "Timeout").length;
    rows.push({ file: path, killed, total: scored.length, pct: Math.round((1000 * killed) / scored.length) / 10 });
  }
  return rows.sort((a, b) => a.pct - b.pct);
}

/** The markdown body for the issue the workflow files or updates. */
export function reportBody(fresh, all, rows) {
  const lines = [];
  lines.push(`${fresh.length} surviving mutant(s) nobody has ruled on, of ${all.length} that survived this run.`);
  lines.push("");
  lines.push(
    "A survivor is a line the engine changed and **no test noticed**. Not all are defects — a mutant can be " +
      "equivalent, or cover behaviour deliberately left unasserted. Decide which, then add it to " +
      "`KNOWN_SURVIVORS` in `scripts/mutation-triage.mjs` **with what was decided**, including " +
      '"no exposure", which records that somebody checked. A survivor left out of that table comes back next week.',
  );
  if (fresh.length) {
    lines.push("");
    lines.push("### Not yet ruled on");
    lines.push("");
    for (const s of fresh.slice(0, 50)) lines.push(`- \`${s.key}\` → \`${s.replacement}\``);
    if (fresh.length > 50) lines.push(`- …and ${fresh.length - 50} more (see the artifact)`);
  }
  if (rows.length) {
    lines.push("");
    lines.push("### Kill rate by file, worst first");
    lines.push("");
    lines.push("| file | killed | of | % |");
    lines.push("| --- | ---: | ---: | ---: |");
    for (const r of rows.slice(0, 15)) lines.push(`| \`${r.file}\` | ${r.killed} | ${r.total} | ${r.pct} |`);
    lines.push("");
    lines.push(
      "A file low in that table is the finding. A single overall percentage is not — it moves with the " +
        "scope of the run, and this run mutates only the week's diff.",
    );
  }
  return lines.join("\n");
}

function main(argv) {
  const path = argv[0];
  if (!path) {
    console.error("usage: node scripts/mutation-triage.mjs <mutation.json>");
    process.exit(2);
  }
  let report;
  try {
    report = JSON.parse(readFileSync(path, "utf8"));
  } catch (err) {
    // NOT exit 3. A report that cannot be read is "not attempted", and reporting
    // it as the finding would file a missing file as a clean bill or a defect.
    console.error(`could not read ${path}: ${err.message}`);
    process.exit(2);
  }
  const all = survivorsOf(report);
  const fresh = all.filter((s) => !(s.key in KNOWN_SURVIVORS));
  console.log(reportBody(fresh, all, scoreByFile(report)));
  process.exit(fresh.length ? 3 : 0);
}

if (isMain(import.meta.url, process.argv[1])) {
  main(process.argv.slice(2));
}
