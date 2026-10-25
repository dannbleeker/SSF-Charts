import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "fs";
import { execFileSync } from "child_process";
// @ts-expect-error — plain .mjs tool, no types.
import { roundNumberOf } from "../scripts/round-pools.mjs";

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
 *
 * ── SEVEN CHARACTERS, CUT HERE AND NOT BY GIT ──────────────────────────────
 * Every lookup below takes `.slice(0, 7)` of a build stamp, because that is the
 * width the project stamps (`rounds.test.ts` pins `^[0-9a-f]{7}$`; `round.mjs`
 * pins `--short=7`). This used to ask git for `%h`, which is git's AUTO length:
 * seven characters until the packed-object count reaches 2^14 = 16,384, eight
 * after that, and eight sooner for any one commit whose seven-character prefix
 * is shared with another object. This repo had 12,940 packed objects when this
 * was measured (2026-10-02), growing by a round archive and some commits a day.
 * On the day it crossed, every post-fix build would have missed a set of
 * eight-character strings and been counted PRE-fix — `core.abbrev=8` reproduces
 * it: two of the six tests here fail on unchanged data. `%H` cut to seven is
 * the same set without the dependence.
 *
 * ── TWO SETS, BECAUSE "NOT POST-FIX" IS NOT "PRE-FIX" ──────────────────────
 * The era used to be `post.has(sha) ? "post" : "PRE"`, so a build git could not
 * place at all — `nogit` or `dev` (what the product stamps when it cannot ask
 * git), an empty stamp, a sha rebased out of history — was counted PRE-fix, and
 * the assertion whose message names that case (`pre + after === deaths`) was a
 * tautology: its two terms filter one array with a predicate and its negation.
 * The only thing left to go red was §0's PRE number, and whoever refreshed §0
 * then folded the misplaced death into PRE-fix with no signal. PRE-fix is now
 * what descends into the fix, and the rest is `undefined`.
 */
interface Eras {
  /** The fix and everything after it. */
  post: Set<string>;
  /** Everything the fix descends from. */
  pre: Set<string>;
}

function fixEras(): Eras {
  let out: string;
  try {
    out = execFileSync("git", ["log", "--format=%H", `${FIX}..HEAD`], { encoding: "utf8" });
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
  const sevens = (log: string) =>
    new Set(
      log
        .split("\n")
        .map((s) => s.trim())
        .filter(Boolean)
        .map((h) => h.slice(0, 7)),
    );
  const post = sevens(out);
  post.add(FIX);
  // FIX is reachable (the call above would have thrown), so this cannot fail on
  // a clone that got that far. The fix's own entry is already in `post`.
  const pre = sevens(execFileSync("git", ["log", "--format=%H", FIX], { encoding: "utf8" }));
  for (const s of post) pre.delete(s);
  return { post, pre };
}

type Era = "PRE" | "post";

/** Which side of the fix a build is on — undefined when git history cannot place it. */
function eraOf(build: string, eras: Eras): Era | undefined {
  const short = build.slice(0, 7);
  return eras.post.has(short) ? "post" : eras.pre.has(short) ? "PRE" : undefined;
}

/** Count builds by era, keeping the ones git cannot place apart from both. */
function splitByEra(builds: string[], eras: Eras): { pre: number; post: number; unplaced: string[] } {
  const split = { pre: 0, post: 0, unplaced: [] as string[] };
  for (const build of builds) {
    const era = eraOf(build, eras);
    if (era === "PRE") split.pre++;
    else if (era === "post") split.post++;
    else split.unplaced.push(build);
  }
  return split;
}

interface Row {
  era: Era;
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
      era: m[1] as Era,
      width: m[2] === "16:9" ? 960 : 720,
      rounds: Number(m[3]),
      green: Number(m[4]),
      pass: m[5],
    });
  return { anchor, rows };
}

/**
 * Is this round newer than the one §0 is anchored to?
 *
 * THE WHOLE NUMBER IN THE NAME, not its first three characters: `1004-x.json` is
 * round 1004, and `Number(f.slice(0, 3))` counted it as round 100, so a table
 * anchored at 1003 took in a round newer than itself (r3-48, the same cut r1-61
 * removed from the pooled readers). Latent until the archive reaches round 1000.
 */
function newerThan(f: string, anchor: number): boolean {
  return roundNumberOf(f) > anchor;
}

/** The newest round in a listing of the rounds directory. */
function highestRound(listing: string[]): number {
  return Math.max(...listing.filter((x) => /^\d+-/.test(x)).map((f) => roundNumberOf(f)));
}

function derive(anchor: number, eras: Eras) {
  const acc: Record<string, { rounds: number; green: number; ok: number; total: number }> = {};
  // Rounds whose build git cannot place, which belong to neither row of the
  // table and used to be counted in PRE-fix — see `fixEras`.
  const unplaced: string[] = [];
  for (const f of readdirSync("rounds").filter((x) => /^\d+-/.test(x))) {
    if (newerThan(f, anchor)) continue;
    let j;
    try {
      j = JSON.parse(readFileSync(`rounds/${f}`, "utf8"));
    } catch {
      continue;
    }
    const st = Array.isArray(j.selftest) ? j.selftest : [];
    if (!st.length) continue;
    const short = String(j.build?.commit ?? f.split("-")[1]?.replace(".json", "")).slice(0, 7);
    const era = eraOf(short, eras);
    if (!era) {
      unplaced.push(`${f} (build ${short})`);
      continue;
    }
    const key = `${era}:${j.slideSize?.width}`;
    acc[key] ??= { rounds: 0, green: 0, ok: 0, total: 0 };
    const a = acc[key];
    a.rounds++;
    a.total += st.length;
    const ok = st.filter((s: { ok: boolean }) => s.ok).length;
    a.ok += ok;
    if (ok === st.length) a.green++;
  }
  return { acc, unplaced };
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
    const { acc, unplaced } = derive(anchor, fixEras());
    // BEFORE THE COUNTS, because a round in neither era is the cause and a
    // mismatched count is only what it looks like from §0.
    expect(
      unplaced,
      "a round's build is in neither era — git cannot place it (rebased out of history, or a local 'dev' build)",
    ).toEqual([]);
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
    const highest = highestRound(readdirSync("rounds"));
    expect(anchor, "§0 is anchored to a round that has not happened").toBeLessThanOrEqual(highest);
  });

  it("reads a round number past 999 as the number it is", () => {
    // NO ROUND 1000 EXISTS YET, so the archive cannot show this; a listing can.
    // The archive is at about round 505 and the driver pads to three digits and
    // never truncates, so `1000-x.json` is a few hundred rounds away. With the
    // three-character cut the newest file of this listing read as 100, the
    // highest round as 999, and a table anchored at 1003 failed "is anchored to
    // a round the archive actually reached" for a table that was right.
    const listing = ["998-a.json", "999-b.json", "1000-c.json", "1003-d.json", "1004-e.json", "predictions.json"];
    expect(highestRound(listing)).toBe(1004);
    expect(
      listing.filter((f) => /^\d+-/.test(f) && !newerThan(f, 1003)),
      "a round newer than the anchor was counted into the table",
    ).toEqual(["998-a.json", "999-b.json", "1000-c.json", "1003-d.json"]);
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
    const {
      pre,
      post: after,
      unplaced,
    } = splitByEra(
      deaths().map((x) => x.build),
      fixEras(),
    );
    // FIRST, because the others would blame §0 for it: a death git cannot place
    // used to fall into PRE-fix, and refreshing §0 to match then made it so.
    expect(
      unplaced,
      "some death's build is in neither era — 'nogit', 'dev', an empty stamp, or a sha rebased out of history",
    ).toEqual([]);
    expect(pre, "§0's PRE-fix death count").toBe(stated!.pre);
    expect(after, "§0's post-fix death count").toBe(stated!.post);
  });
});

/**
 * THE ERA SPLIT HAS TO BE ABLE TO FAIL, AND TO NOT DEPEND ON HOW GIT PRINTS.
 *
 * Both halves are properties of the helpers above rather than of today's
 * archive, which classifies cleanly (93 + 16 deaths, 352 + 121 rounds) and so
 * proves neither: an unplaceable build has to come back as unplaced, and the
 * sets have to hold the same seven characters whatever width git abbreviates to.
 */
describe("the era split", () => {
  const rev = (what: string) => execFileSync("git", ["rev-parse", what], { encoding: "utf8" }).trim();

  it("does not count a build git cannot place as PRE-fix", () => {
    const eras = fixEras();
    const real = { post: rev("HEAD").slice(0, 7), pre: rev(`${FIX}^`).slice(0, 7) };
    // What the product stamps without git (`nogit`, `dev`), a stamp with nothing
    // in it, and a well-formed sha that is in no history.
    const unplaceable = ["nogit", "dev", "", "deadbee"];
    expect(unplaceable.filter((b) => eras.post.has(b) || eras.pre.has(b))).toEqual([]);
    expect(splitByEra([real.pre, real.post, ...unplaceable], eras)).toEqual({ pre: 1, post: 1, unplaced: unplaceable });
  });

  it("is the same set at any abbreviation width git chooses", () => {
    // `core.abbrev` stands in for the repo growing past 16,384 packed objects,
    // where git's own default moves from seven characters to eight. Set through
    // the environment so `fixEras` is run exactly as it is everywhere else.
    const env = { GIT_CONFIG_COUNT: "1", GIT_CONFIG_KEY_0: "core.abbrev", GIT_CONFIG_VALUE_0: "12" };
    const saved = Object.fromEntries(Object.keys(env).map((k) => [k, process.env[k]]));
    Object.assign(process.env, env);
    try {
      // The override has to take effect, or this proves nothing about it: an old
      // git ignores these variables and would pass with the dependence intact.
      const printed = execFileSync("git", ["log", "-1", "--format=%h"], { encoding: "utf8" }).trim();
      expect(printed.length, "this git ignores GIT_CONFIG_COUNT, so the check below would be vacuous").toBe(12);
      const wide = fixEras();
      expect(eraOf(rev("HEAD"), wide)).toBe("post");
      expect(eraOf(rev(`${FIX}^`), wide)).toBe("PRE");
    } finally {
      for (const [k, v] of Object.entries(saved)) {
        if (v === undefined) delete process.env[k];
        else process.env[k] = v;
      }
    }
  });
});
