import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "fs";

/**
 * `crashes/*.json` is committed to a PUBLIC repo. Nothing in it may identify
 * the session it came from.
 *
 * The split in `.gitignore` is load-bearing and was undocumented: `crashes/*.md`
 * is excluded because those reports carry request URLs, session ids and the
 * document's name, while `crashes/*-crashed-run.json` is committed — 98 of them
 * — because it is the only surviving copy of what PowerPoint said as it died,
 * and the pane assembles it from its own trace buffer rather than from the
 * page.
 *
 * That second half was an assumption until 2026-09-06, when all 98 were scanned
 * and none carried a URL, a `resid`/`cid`, a personal-drive path, an address or
 * a token. This test is that scan, kept: the pane's payload could grow a field
 * at any time, and the first person to notice would otherwise be a stranger
 * reading the repo.
 *
 * IT DOES NOT LOOK AT THE `.md` FILES. Those legitimately contain all of it and
 * are excluded for exactly that reason — asserting over them would fail
 * immediately and teach the next reader the opposite of the rule.
 */
const FORBIDDEN: Array<[string, RegExp]> = [
  ["a host URL", /onedrive\.live\.com|officeapps\.live\.com|sharepoint\.com/],
  ["a document id in a query string", /\b(resid|cid)=/],
  ["a personal-drive path", /personal\/[0-9a-f]{8,}/],
  ["an email address", /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/],
  ["a bearer token", /access_token|Bearer /],
];

describe("committed crash records", () => {
  const files = readdirSync("crashes").filter((f) => f.endsWith(".json"));

  it("has records to check at all", () => {
    // A scan of nothing passes every assertion below, and "verified" would be
    // indistinguishable from "never ran". This repo has paid for that shape
    // four times in one session.
    expect(files.length, "no crashes/*.json found — this guard checked nothing").toBeGreaterThan(0);
  });

  it("carries no URL, session id, personal path, address or token", () => {
    const offenders: string[] = [];
    for (const f of files) {
      let text: string;
      try {
        text = readFileSync(`crashes/${f}`, "utf8");
      } catch {
        continue;
      }
      for (const [what, re] of FORBIDDEN) if (re.test(text)) offenders.push(`${f}: ${what}`);
    }
    // Named, not counted: the point of failing is to say which file to look at.
    expect(offenders, `these would publish session data — see crashes/README.md`).toEqual([]);
  });

  it("would actually catch one", () => {
    // WITHOUT THIS, EMPTYING `FORBIDDEN` IS A GREEN TEST. A guard whose whole
    // output is "nothing matched" cannot tell "clean" from "looking for
    // nothing", which is the same failure as a scan over an empty directory —
    // so the patterns are exercised against a record that should trip every
    // one of them.
    const planted = [
      "https://onedrive.live.com/personal/911d2b9d99de05ee/_layouts/15/doc.aspx?resid=abc&cid=def",
      "someone@example.com",
      "Authorization: Bearer eyJhbGciOi",
    ].join("\n");
    const caught = FORBIDDEN.filter(([, re]) => re.test(planted)).map(([what]) => what);
    expect(caught, "the patterns no longer recognise what they exist to find").toHaveLength(FORBIDDEN.length);
  });
});
