#!/usr/bin/env node
/**
 * Does the install path in the README still work?
 *
 * Everything in this repo gates the file in the working tree. Nothing has ever
 * looked at the file a USER downloads, and the two have now diverged twice:
 *
 * - v0.1.0 shipped the DEV manifests while the README said to download
 *   `manifest-prod.xml`, which was not in the release at all. Broken for twelve
 *   days, while `release.yml` sat correct and un-run.
 * - v0.3.0 (2026-07-31) shipped `<Version>0.1.0</Version>`, which Microsoft's
 *   validator rejects. CI found it on 2026-08-06 and #289 fixed the repo the
 *   same day. No release has been cut since, so the published manifest is still
 *   the rejected one and the fix has reached nobody.
 *
 * The second is the case a byte-comparison finds and a rules check alone does
 * not: the release was fine when it was cut, and went stale when main moved. So
 * this asks three questions, and the third is the one nothing else asks.
 *
 *   1. Does every asset the README names exist in the latest release?
 *   2. Do the published manifests satisfy the rules in `manifest-rules.mjs`?
 *   3. Do they still MATCH the committed ones — i.e. does the latest release
 *      carry the current fixes?
 *
 *   node scripts/check-published-install.mjs [--repo owner/name]
 *
 * Exit 0 when the published install path is sound, 3 when it is not (the same
 * "there is something to report" code the other sweeps use), 2 when the check
 * itself could not run — a network failure must never read as a clean bill, and
 * must never read as a finding either: that goes for the URL sweep at the end as
 * much as for the release fetch at the start.
 */
import { readFileSync } from "node:fs";
import { checkManifest, urlsIn } from "./manifest-rules.mjs";
import { isMain } from "./is-main.mjs";

const REPO = "dannbleeker/SSF-Charts";

/** The files the README's install steps tell a user to download. */
export const PUBLISHED_MANIFESTS = ["manifest-prod.xml", "manifest-excel-prod.xml"];

/**
 * Compare what is published against what is committed.
 *
 * Line endings are normalised before the comparison: a release asset that has
 * been through a checkout on another platform can differ by \r alone, and
 * reporting that as "your users have the wrong manifest" would train the reader
 * to ignore this check within two weeks.
 *
 * @param {{name: string, published: string|null, committed: string}[]} pairs
 * @returns {string[]} problems, empty when the path is sound
 */
export function judgePublished(pairs) {
  const problems = [];
  for (const { name, published, committed } of pairs) {
    if (published == null) {
      problems.push(`${name} is not in the latest release — the README's step 1 downloads a 404`);
      continue;
    }
    problems.push(...checkManifest(published, name, { expectProduction: true }));
    const norm = (s) => s.replace(/\r\n/g, "\n").trimEnd();
    if (norm(published) !== norm(committed)) {
      problems.push(
        `${name} in the latest release is NOT the committed one — main has moved and no release was cut, ` +
          `so every fix since then has reached nobody. Cut a release.`,
      );
    }
  }
  return problems;
}

/** The report a sweep posts. */
export function reportBody(problems, tag) {
  if (!problems.length) {
    return `The published install path is sound: release \`${tag}\` carries the committed manifests and they pass every rule.\n`;
  }
  return (
    `The install path the README gives users is broken or stale, against release \`${tag}\`.\n\n` +
    `This is the one failure nothing else in the repo can see. Every gate here checks the file in the\n` +
    `working tree; a user downloads the one attached to the latest release, and those have diverged twice.\n\n` +
    problems.map((p) => `- ${p}`).join("\n") +
    `\n\nCut a release with \`gh workflow run release.yml -f version=vX.Y.Z\` from a green main.\n`
  );
}

/**
 * HEAD every URL the published manifests ask a host to fetch.
 *
 * TWO KINDS OF NOT-OK, AND ONLY ONE IS A FINDING. A URL that ANSWERS 404 is a
 * reading about the manifest: it goes in `problems`. A fetch that REJECTS — a
 * DNS blip, a reset, a timeout — is not an answer about the URL at all; it is
 * this run failing to look, so it goes in `unreachable` and never in `problems`.
 *
 * It used to be one list. The rejection was caught and turned into
 * `{ ok: false, status: e.message }`, so a single dropped connection among the
 * ~23 requests of a weekly run read as "<url> answers fetch failed", exited 3,
 * and `quality-sweep.yml` filed or reopened the "install watch" issue — whose
 * body ends "Cut a release", which cannot fix a network failure. The same
 * failure on the release fetch above already exited 2, as the header promises.
 *
 * @param {{name: string, published: string}[]} pairs  manifests that are present — a missing one is `judgePublished`'s finding
 * @param {(url: string) => Promise<{ok: boolean, status: number}>} head  injectable so a test can fail it
 * @returns {Promise<{problems: string[], unreachable: string[]}>}
 */
export async function sweepUrls(pairs, head = (url) => fetch(url, { method: "HEAD", redirect: "follow" })) {
  const problems = [];
  const unreachable = [];
  for (const { name, published } of pairs) {
    for (const url of urlsIn(published)) {
      let res;
      try {
        res = await head(url);
      } catch (err) {
        unreachable.push(`${name} → ${url}: ${err?.message ?? err}`);
        continue;
      }
      if (!res.ok) problems.push(`${name} points at ${url}, which answers ${res.status}`);
    }
  }
  return { problems, unreachable };
}

/**
 * The whole check, returning the exit code instead of setting it — so a test can
 * hand it a `fetch` that fails and read what the workflow would have seen.
 *
 * @param {{ repo?: string, get?: typeof fetch, read?: (p: string, enc: string) => string,
 *           out?: (s: string) => void, warn?: (s: string) => void }} [deps]
 * @returns {Promise<0 | 2 | 3>}
 */
export async function main({
  repo = REPO,
  get = fetch,
  read = readFileSync,
  out = (s) => process.stdout.write(s),
  warn = (s) => console.error(s),
} = {}) {
  let tag;
  const pairs = [];
  try {
    const rel = await get(`https://api.github.com/repos/${repo}/releases/latest`, {
      headers: { accept: "application/vnd.github+json" },
    });
    if (!rel.ok) throw new Error(`the releases API answered ${rel.status}`);
    const latest = await rel.json();
    tag = latest.tag_name;
    for (const name of PUBLISHED_MANIFESTS) {
      const asset = (latest.assets ?? []).find((a) => a.name === name);
      let published = null;
      if (asset) {
        const got = await get(asset.browser_download_url, { redirect: "follow" });
        if (!got.ok) throw new Error(`${name} answered ${got.status}`);
        published = await got.text();
      }
      pairs.push({ name, published, committed: read(name, "utf8") });
    }
  } catch (err) {
    // Exit 2, never 0. A check that cannot reach the network and says nothing is
    // indistinguishable from a check that looked and found nothing wrong.
    warn(`could not read the published install path: ${err.message}`);
    return 2;
  }

  const problems = judgePublished(pairs);
  let unreachable = [];
  // Only worth reporting when the manifests themselves are sound — a URL sweep
  // on a manifest nobody can install is noise on top of the real finding.
  if (!problems.length) {
    const swept = await sweepUrls(pairs, (url) => get(url, { method: "HEAD", redirect: "follow" }));
    problems.push(...swept.problems);
    unreachable = swept.unreachable;
  }

  // A SWEEP THAT COULD NOT FINISH IS NOT A CLEAN ONE, AND NOT A FINDING EITHER.
  // With nothing else wrong it is exit 2, the same "the check itself could not
  // run" the release fetch gives: the job fails visibly and files nothing.
  // With a real finding already in hand the finding still stands — a URL that
  // answered 404 is broken whatever the others did — and the ones that could not
  // be checked are said on stderr, which the issue body (stdout) does not carry.
  if (unreachable.length) {
    warn(
      `could not check ${unreachable.length} URL(s) — the request failed rather than being answered:\n` +
        unreachable.map((u) => `  ${u}`).join("\n"),
    );
    if (!problems.length) return 2;
  }

  out(reportBody(problems, tag));
  return problems.length ? 3 : 0;
}

if (isMain(import.meta.url, process.argv[1])) {
  const argRepo = process.argv.indexOf("--repo");
  process.exitCode = await main({ repo: argRepo > -1 ? process.argv[argRepo + 1] : REPO });
}
