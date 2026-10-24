#!/usr/bin/env node
/**
 * Re-read every office-js issue this repo cites, and say which have moved.
 *
 * This repo reasons from upstream bugs constantly — a comment says "the host
 * refuses X, see office-js#NNNN" and the code around it is shaped by that. Those
 * statements decay silently. Nobody is told when Microsoft closes one, and a
 * closed-as-completed issue means the opposite of what the comment says: the
 * behaviour may be FIXED, and the code may be working around nothing.
 *
 * The first time this was checked by hand, on 2026-08-14, four of the ten cited
 * issues had been closed as completed and two of the stale statuses were
 * load-bearing:
 *
 *   #5022  called "open and assigned" in host-probe.ts, and printed as "(open)"
 *          by triage on every did-not-respond problem. Closed 2024-11-18.
 *   #3014  the stated reason a `no` from `grouped-child-by-id-from-slide` was
 *          "expected". Closed 2025-03-03, so neither answer is expected now.
 *
 * Eighteen months of planning around a fixed bug, in a repo that documents
 * everything. Hence a script rather than a good intention.
 *
 *   node scripts/issue-status.mjs            # report
 *   node scripts/issue-status.mjs --check    # exit 1 if any cited issue is closed
 *
 * Needs `gh` on PATH and authenticated; without it this says so and exits 0
 * rather than failing a gate for a reason that is not about the code.
 */
import { readFileSync, readdirSync, statSync } from "fs";
import { join, extname } from "path";
import { spawnSync } from "child_process";
import { isMain } from "./is-main.mjs";

const ROOTS = ["src", "scripts", "docs", "test"];
const FILES = [".ts", ".mjs", ".js", ".md"];

/**
 * How this repo cites an office-js issue: `office-js#NNNN`, or the issue's own
 * URL (`github.com/OfficeDev/office-js/issues/NNNN`, which is what a markdown
 * link `[#NNNN](…)` carries).
 *
 * A BARE `#NNNN` is deliberately not matched. It is far too common in prose here
 * to be a reliable signal, and a wrong number in this report is worse than a
 * missing one: it sends someone to an unrelated issue. That reason does not reach
 * a full URL, which cannot be wrong about WHICH issue it means — and skipping it
 * hid #2881 and #3309 (the reasons `docs/RESEARCH.md` gives for rejecting SVG
 * coercion), #5896 and #6266 from a report whose header promises "every
 * office-js issue this repo cites". `test/office-js-watch.test.ts` already counts
 * the URL form as a citation; the two tools disagreed about what "cited" means.
 *
 * ONE PATTERN FOR BOTH SCANS below (the tree, and the two root files), so a third
 * spelling is added in one place.
 */
const CITATION = /office-js(?:\s*#|\/issues\/)(\d{3,6})/g;

/** Every office-js issue this repo cites, and where. */
export function citedIssues(root = ".") {
  const found = new Map();
  const walk = (dir) => {
    let entries;
    try {
      entries = readdirSync(dir);
    } catch {
      return;
    }
    for (const e of entries) {
      if (e === "node_modules" || e === ".git") continue;
      const p = join(dir, e);
      let s;
      try {
        s = statSync(p);
      } catch {
        continue;
      }
      if (s.isDirectory()) walk(p);
      else if (FILES.includes(extname(e))) {
        let text;
        try {
          text = readFileSync(p, "utf8");
        } catch {
          continue;
        }
        // See `CITATION`: `office-js#NNNN` or the issue's URL, never a bare `#NNNN`.
        for (const m of text.matchAll(CITATION)) {
          const n = m[1];
          if (!found.has(n)) found.set(n, new Set());
          found.get(n).add(p.replace(/\\/g, "/"));
        }
      }
    }
  };
  for (const r of ROOTS) walk(join(root, r));
  for (const f of ["CLAUDE.md", "README.md"]) {
    try {
      for (const m of readFileSync(join(root, f), "utf8").matchAll(CITATION)) {
        const n = m[1];
        if (!found.has(n)) found.set(n, new Set());
        found.get(n).add(f);
      }
    } catch {
      /* absent is fine */
    }
  }
  return found;
}

/** One issue's live state, or null when `gh` could not answer. */
export function issueState(n, run = spawnSync) {
  const r = run("gh", ["api", `repos/OfficeDev/office-js/issues/${n}`, "--jq", ".state,.state_reason,.title"], {
    encoding: "utf8",
  });
  if (r.status !== 0) return null;
  const [state, reason, ...rest] = String(r.stdout).trim().split("\n");
  return { state, reason: reason === "null" ? null : reason, title: rest.join(" ") };
}

function main(argv) {
  const check = argv.includes("--check");
  const cited = citedIssues();
  if (!cited.size) {
    console.log("no office-js issues cited anywhere — nothing to check");
    return 0;
  }
  const probe = issueState([...cited.keys()][0]);
  if (!probe) {
    console.log("gh unavailable or unauthenticated — skipping the upstream check");
    return 0;
  }
  let closed = 0;
  console.log(`\n  UPSTREAM STATUS — ${cited.size} issue(s) cited by this repo\n`);
  for (const n of [...cited.keys()].sort((a, b) => Number(a) - Number(b))) {
    const s = issueState(n);
    if (!s) {
      console.log(`  #${n}  ?        could not be read`);
      continue;
    }
    const shut = s.state === "closed";
    // COMPLETED is the one that matters. `not_planned` means upstream declined
    // it, so the behaviour stands and the comment citing it is still true.
    const loud = shut && s.reason === "completed";
    if (loud) closed++;
    console.log(
      `  #${n}  ${(s.state + (s.reason ? `/${s.reason}` : "")).padEnd(18)}${loud ? "FIXED UPSTREAM — re-read what cites it" : ""}`,
    );
    if (loud) for (const f of [...cited.get(n)].sort()) console.log(`        ${f}`);
  }
  if (closed) {
    console.log(
      `\n  ${closed} cited issue(s) are closed as COMPLETED. A comment that plans around one is\n` +
        `  describing a bug that may no longer exist. Check before trusting it — and remember a\n` +
        `  fix upstream is a claim about the service, not proof it reached this host.\n`,
    );
  }
  return check && closed ? 1 : 0;
}

// `process.argv[1]` is NOT optional here. Called with one argument `isMain`
// returns false every time, `main()` never runs, and the script prints nothing
// and exits 0 — which does not read as "did not run", it reads as a pass. That
// is the precise bug `is-main.mjs` exists to end, and it caught this file on its
// first run.
if (isMain(import.meta.url, process.argv[1])) process.exit(main(process.argv.slice(2)));
