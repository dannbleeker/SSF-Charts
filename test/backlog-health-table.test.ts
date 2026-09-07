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

/** Builds that contain the two-master fix, by ancestry rather than by date. */
function postFixBuilds(): Set<string> {
  const out = execFileSync("git", ["log", "--format=%h", `${FIX}..HEAD`], { encoding: "utf8" });
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
