import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";

/**
 * The backlog's stated count of open items must match the list it prints.
 *
 * THE SAME MISTAKE FOUR TIMES IN TWO DAYS. `docs/BACKLOG.md` said "six" above a
 * list of four, then "five" above a list of three. Nobody was careless: the
 * count sits in a sentence and the list sits below it, and closing an item
 * touches the list. The file even carries its own tie-breaker —
 *
 *     "if a thing is not on that list it is not open"
 *
 * — which is the right rule and cannot enforce itself.
 *
 * NARROW ON PURPOSE. A checker that re-derived every number in that file would
 * go red on every round and be disabled within a week. This asserts one thing
 * that only changes when a person edits the list, which is exactly when it has
 * been wrong.
 */
const WORDS: Record<string, number> = {
  none: 0,
  one: 1,
  two: 2,
  three: 3,
  four: 4,
  five: 5,
  six: 6,
  seven: 7,
  eight: 8,
  nine: 9,
  ten: 10,
};

/** §1, which the file says "opens with the whole open list". */
function openSection(md: string): string[] {
  const from = md.indexOf("## 1. Open");
  const to = md.indexOf("## 2.");
  expect(from, "docs/BACKLOG.md has no `## 1. Open`").toBeGreaterThan(-1);
  expect(to, "docs/BACKLOG.md has no `## 2.` to bound section 1").toBeGreaterThan(from);
  return md.slice(from, to).split("\n");
}

/**
 * The ids in the open list, and only those.
 *
 * `§1` also carries data tables whose rows are indented digits — `29`, `56`,
 * `30` — and a naive scan reads those as open items, which is how the first
 * version of this test claimed five. So it takes the FIRST contiguous indented
 * block that looks like the list and stops at the prose underneath it.
 */
function openIds(lines: string[]): string[] {
  const entry = /^ {4}(\d+) {2}[a-z]/;
  const start = lines.findIndex((l) => entry.test(l));
  expect(start, "section 1 prints no open list at all").toBeGreaterThan(-1);
  const ids: string[] = [];
  for (let i = start; i < lines.length; i++) {
    const line = lines[i];
    // A blank line inside the block is a gap between entries; a blank followed
    // by anything unindented ends it.
    if (line.trim() === "") {
      if (!/^ {4,}\S/.test(lines[i + 1] ?? "")) break;
      continue;
    }
    if (!/^ {4,}/.test(line)) break;
    const m = entry.exec(line);
    if (m && !ids.includes(m[1])) ids.push(m[1]);
  }
  return ids;
}

/**
 * The sentence that says how many are open.
 *
 * `#6237 is open` is a tracker issue in the same section and matches the same
 * shape, so a leading `#` disqualifies a number — the count is a quantity, not
 * an identifier.
 */
function claimedCount(section: string): number | null {
  const m = /(^|[^#\w])(none|one|two|three|four|five|six|seven|eight|nine|ten|\d+)\s+(?:is|are)\s+open\b/i.exec(
    section,
  );
  if (!m) return null;
  const token = m[2].toLowerCase();
  return token in WORDS ? WORDS[token] : Number(token);
}

describe("the backlog's open count", () => {
  const md = readFileSync("docs/BACKLOG.md", "utf8");
  const lines = openSection(md);

  it("matches the list printed under it", () => {
    const ids = openIds(lines);
    const claimed = claimedCount(lines.join("\n"));
    expect(claimed, "section 1 states no count of open items — one of them was always wrong before").not.toBeNull();
    expect(claimed, `section 1 says ${claimed} open, and lists ${ids.length}: ${ids.join(", ")}`).toBe(ids.length);
  });

  it("stops at the end of the list instead of sweeping the tables below it", () => {
    // The guard on the guard, and the first version of it asserted the wrong
    // rule — that a table is recognisable as a table. It is not: the same
    // shape means "open item" here and "row" three screens down. What actually
    // separates them is POSITION, so that is what this pins.
    const ids = openIds(lines);
    expect(ids, "section 1's real list").toEqual(["5"]);
    expect(
      openIds([
        "## 1. Open",
        "",
        "    7  a real open item",
        "       with a continuation line",
        "",
        "Prose that ends the list.",
        "",
        "    29  this update draws a picture          — legitimate",
        "    56  the chart has no parts list",
      ]),
      "prose ends the list; the table under it is not part of it",
    ).toEqual(["7"]);
  });
});
