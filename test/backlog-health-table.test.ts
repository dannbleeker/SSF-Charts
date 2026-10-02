import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "fs";
import { execFileSync } from "child_process";

/**
 * §0's health table must still reproduce at the round it is anchored to.
 *
 * THE NUMBER THAT KEEPS GOING STALE. That table has been wrong twice in three
 * days — the 4:3 row read `26 / 22 / 96.8%` after two more rounds landed, and
 * before that the whole thing drifted silently. It now carries "As of round
 * NNN", and an anchor nobody checks is just a date on a wrong number.
 *
 * WHY THIS ONE CAN BE GUARDED AND A GENERAL CHECKER CANNOT. A test that
 * re-derived every figure in that file would go red on every round and be
 * disabled inside a week. This one is stable by construction: it re-derives at
 * the STATED anchor, so it only moves when a person edits the table or the
 * anchor. That is exactly when it has been wrong.
 *
 * It also catches the subtler failure, which is the one that actually happened:
 * committing a table anchored to 413 when the archive already held 417. The
 * anchor was honest about the arithmetic and misleading about the currency, and
 * the assertion at the bottom is what says so.
 *
 * `execFileSync` and NOT `execSync` for git: on Windows `execSync` goes through
 * cmd.exe, where `^` is the escape character, and a revision like `abc123^{commit}`
 * is silently mangled. An adversarial pass hit exactly that this week and
 * misclassified 421 of 462 builds before catching it.
 */
const FIX = "6dfaa4b";

/**
 * Builds that contain the two-master fix, by ancestry rather than by date.
 *
 * NEEDS FULL GIT HISTORY, and now says so when it does not have it. A shallow
 * clone — the default for `actions/checkout` — has no `6dfaa4b`, so git answers
 * "ambiguous argument '6dfaa4b..HEAD': unknown revision or path not in the
 * working tree", which names neither the cause nor the cure. This test passed
 * on every developer machine and could never pass in CI: it went red the hour
 * it was added and stayed red for 26 hours and twenty commits, while the local
 * gate — which does not run coverage, and prints a list saying so — kept
 * reporting green.
 *
 * Fixed at the source: `ci.yml` checks out with `fetch-depth: 0`. Translated
 * here as well, because the next environment without history should be told
 * what this needs rather than what git said.
 */
function postFixBuilds(): Set<string> {
  let out: string;
  try {
    out = execFileSync("git", ["log", "--format=%h", `${FIX}..HEAD`], { encoding: "utf8" });
  } catch (err) {
    let reachable = true;
    try {
      execFileSync("git", ["cat-file", "-e", `${FIX}^{commit}`], { stdio: "ignore" });
    } catch {
      reachable = false;
    }
    // The original is carried in the MESSAGE rather than as `cause`, which is
    // what `preserve-caught-error` would rather see. `Error.cause` is ES2022 and
    // this project compiles to ES2020, so the typed form does not exist here —
    // and raising the whole repo's target to satisfy one throw is a bigger
    // change than the thing it fixes. Nothing is lost: git's own words are in
    // the text.
    // eslint-disable-next-line preserve-caught-error -- ES2020 target has no Error.cause; the original is in the message
    throw new Error(
      reachable
        ? `git could not list ${FIX}..HEAD although ${FIX} is present — ${String(err)}`
        : `this test needs full git history — ${FIX} (the two-master fix) is not in this clone, ` +
            `so nothing can say which builds contain it. A shallow checkout cannot answer it; ` +
            `CI uses fetch-depth: 0 for exactly this reason. git said: ${String(err)}`,
    );
  }
  const set = new Set(
    out
      .split("\n")
      .map((s) => s.trim())
      .filter(Boolean),
  );
  set.add(FIX);
  return set;
}

interface Row {
  era: "PRE" | "post";
  width: number;
  rounds: number;
  green: number;
  pass: string;
}

function statedRows(md: string): { anchor: number; rows: Row[] } {
  const anchor = Number(/\*\*As of round (\d+)\*\*/.exec(md)?.[1]);
  const rows: Row[] = [];
  const re = /^\s{4}(PRE|post)-fix\s+(16:9|4:3)\s+(\d+)\s+(\d+)\s+([\d.]+)%$/gm;
  for (const m of md.matchAll(re))
    rows.push({
      era: m[1] as "PRE" | "post",
      width: m[2] === "16:9" ? 960 : 720,
      rounds: Number(m[3]),
      green: Number(m[4]),
      pass: m[5],
    });
  return { anchor, rows };
}

function derive(anchor: number, post: Set<string>) {
  const acc: Record<string, { rounds: number; green: number; ok: number; total: number }> = {};
  for (const f of readdirSync("rounds").filter((x) => /^\d+-/.test(x))) {
    if (Number(f.slice(0, 3)) > anchor) continue;
    let j;
    try {
      j = JSON.parse(readFileSync(`rounds/${f}`, "utf8"));
    } catch {
      continue;
    }
    const st = Array.isArray(j.selftest) ? j.selftest : [];
    if (!st.length) continue;
    const short = String(j.build?.commit ?? f.split("-")[1]?.replace(".json", "")).slice(0, 7);
    const key = `${post.has(short) ? "post" : "PRE"}:${j.slideSize?.width}`;
    acc[key] ??= { rounds: 0, green: 0, ok: 0, total: 0 };
    const a = acc[key];
    a.rounds++;
    a.total += st.length;
    const ok = st.filter((s: { ok: boolean }) => s.ok).length;
    a.ok += ok;
    if (ok === st.length) a.green++;
  }
  return acc;
}

describe("§0's health table", () => {
  const md = readFileSync("docs/BACKLOG.md", "utf8");
  const { anchor, rows } = statedRows(md);

  it("states an anchor and four rows", () => {
    // A parse that found nothing would make every assertion below vacuous —
    // the shape this repo has now been bitten by twice.
    expect(anchor, "§0 no longer says which round it is anchored to").toBeGreaterThan(0);
    expect(rows, "the four era/arm rows did not parse").toHaveLength(4);
  });

  it("reproduces every row at the round it claims", () => {
    const acc = derive(anchor, postFixBuilds());
    for (const r of rows) {
      const a = acc[`${r.era}:${r.width}`];
      const where = `${r.era}-fix ${r.width === 960 ? "16:9" : "4:3"} at round ${anchor}`;
      expect(a, `${where}: no rounds matched at all`).toBeTruthy();
      expect(a.rounds, `${where}: round count`).toBe(r.rounds);
      expect(a.green, `${where}: all-green count`).toBe(r.green);
      expect(((100 * a.ok) / a.total).toFixed(1), `${where}: scenario pass rate`).toBe(r.pass);
    }
  });

  it("is anchored to a round the archive actually reached", () => {
    // The failure that actually happened: a table anchored to 413 was committed
    // when the archive already held 417. Honest arithmetic, misleading
    // currency. This does not demand the anchor be the newest round — rounds
    // land faster than tables are rewritten — only that it exists.
    const highest = Math.max(
      ...readdirSync("rounds")
        .filter((x) => /^\d+-/.test(x))
        .map((f) => Number(f.slice(0, 3))),
    );
    expect(anchor, "§0 is anchored to a round that has not happened").toBeLessThanOrEqual(highest);
  });
});

/**
 * §0's CRASH LINE must reproduce too, and for four weeks it could not.
 *
 * The line above it has been guarded since 2026-09-08. This one sat beside it
 * carrying `94 / 3 / 97` and the words "NOT RE-DERIVABLE", with three failed
 * attempts recorded underneath — 108/0 counting `.md` by filename date, 202/4
 * counting every file, 103 pairing them. The reason given was right: the
 * directory holds TWO capture mechanisms, so "a crash record" was not one thing.
 *
 * It is one thing now, stated in §0 and asserted here:
 *
 *     a death = one distinct pane trace buffer,
 *               keyed by (build sha, `startedAt` to the millisecond)
 *
 * ── WHY A FILE IS NOT A DEATH, which is the whole finding ───────────────────
 * The pane's buffer survives in `localStorage` until some later round downloads
 * it, so a round that meets an orphaned buffer downloads it AGAIN. Six of the 115
 * files are re-downloads of five deaths already recorded — one pair is whole-file
 * byte-identical, one death has three copies — and FIVE of the six are post-fix,
 * so a file count overstates the post-fix arm by 31%. That arm is the one the era
 * comparison rests on.
 *
 * ── WHY THE KEY IS (build, startedAt) AND NOT THE FILENAME ──────────────────
 * The filename is the moment of DOWNLOAD, which is what differs between copies.
 * `startedAt` is the moment the run that died began, which is what does not.
 * Cross-checked when this was written: grouping instead on a SHA-1 of the raw
 * `steps` array returns the same 109, and two groupings that disagree would mean
 * neither could be trusted.
 *
 * ── AND IT IS A FLOOR ───────────────────────────────────────────────────────
 * A crash that closes the pane destroys the buffer before anything can download
 * it. Round 492 met PowerPoint's own crash dialog at readiness and produced no
 * record at all. This asserts what was CAPTURED, which is a lower bound on what
 * happened — the same framing `rounds-gate.mjs` prints for its own crash rates.
 */
describe("§0's crash line", () => {
  /** `93 deaths on PRE-fix builds, 16 on post-fix, 109 total` — read off §0. */
  function statedCrashes(md: string) {
    const m = /crash records:\s+(\d+) deaths on PRE-fix builds,\s*(\d+) on post-fix,\s*(\d+) total/.exec(md);
    const files = /from (\d+) files/.exec(md);
    return m
      ? { pre: Number(m[1]), post: Number(m[2]), total: Number(m[3]), files: files ? Number(files[1]) : null }
      : null;
  }

  /** One entry per DISTINCT death, with the build that died. */
  function deaths(): { build: string }[] {
    const seen = new Map<string, { build: string }>();
    for (const f of readdirSync("crashes").filter((x) => x.endsWith("-crashed-run.json"))) {
      let j;
      try {
        j = JSON.parse(readFileSync(`crashes/${f}`, "utf8"));
      } catch {
        continue;
      }
      const build = String(j.build ?? "").split(" ")[0];
      // A death is the run that died, not the download that saved it.
      const key = `${build}|${String(j.startedAt ?? f)}`;
      if (!seen.has(key)) seen.set(key, { build });
    }
    return [...seen.values()];
  }

  const md = readFileSync("docs/BACKLOG.md", "utf8");
  const stated = statedCrashes(md);

  it("states a crash count this can read", () => {
    // A parse that found nothing makes every assertion below vacuous — the shape
    // the sibling test above was bitten by twice.
    expect(stated, "§0's crash line no longer parses — re-read it before editing this").toBeTruthy();
  });

  it("counts DEATHS, not files, and says so", () => {
    const files = readdirSync("crashes").filter((x) => x.endsWith("-crashed-run.json")).length;
    const d = deaths();
    expect(d.length, "§0's crash total is not the number of distinct deaths").toBe(stated!.total);
    expect(files, "§0's file count is wrong").toBe(stated!.files);
    // THE NON-VACUITY HALF. If files and deaths were ever equal this test would
    // pass while proving nothing about de-duplication — which is the state §0 was
    // in. Re-downloads are a property of how the buffer is rescued, so they will
    // keep happening; if this ever fails, check whether `round.mjs` stopped
    // leaving orphaned buffers before loosening it.
    expect(files, "no duplicate crash files at all — has the rescue changed?").toBeGreaterThan(d.length);
  });

  it("splits the deaths by ancestry of the two-master fix", () => {
    const post = postFixBuilds();
    const d = deaths();
    const pre = d.filter((x) => !post.has(x.build.slice(0, 7))).length;
    const after = d.filter((x) => post.has(x.build.slice(0, 7))).length;
    expect(pre, "§0's PRE-fix death count").toBe(stated!.pre);
    expect(after, "§0's post-fix death count").toBe(stated!.post);
    expect(pre + after, "some death's build is in neither era — unresolvable sha?").toBe(d.length);
  });
});
