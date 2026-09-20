import { describe, expect, it } from "vitest";
import { readdirSync, readFileSync } from "fs";
import { fileURLToPath } from "url";

/**
 * THE TWO PAGES A CERTIFICATION REVIEWER READS INSTEAD OF THE CODE.
 *
 * `docs/STORE-LISTING.md` submits `/privacy.html` and `/terms.html` as the
 * listing's Privacy and Terms URLs, and they are the only description of this
 * product's data handling that anyone outside the repo will ever see. Nothing
 * checked them, and they drifted the way `manifest.test.ts` and
 * `support-url-page.test.ts` describe for the README: by the product growing
 * underneath a page that nobody re-read.
 *
 * WHAT DRIFTED, found 2026-09-20. Both pages said "PowerPoint add-in" and
 * nothing else, written when that was the whole product. This repo now
 * store-preps a SECOND manifest — `manifest-excel.xml` / `manifest-excel-prod.xml`,
 * the same `<Version>`, validated in the same sweep, pointed at
 * `https://ssf-chart.struktureretsundfornuft.dk/src/excel/excel.html`, which
 * `vite.config.ts` builds and Pages deploys. That companion reads the values of
 * the user's selected range (`Excel.run` → `range.load("values,address")` in
 * `src/excel/excel.ts`) and puts the result on their clipboard, and the privacy
 * policy covering it did not know Excel existed.
 *
 * So the rule these tests pin is the one that failed: **the pages describe the
 * hosts the repo ships, all of them and only them.** Both directions matter —
 * the listing's own words, "a policy that claims handling the code does not do
 * is as much a problem as one that omits what it does".
 *
 * The pages live in `public/`, which is `.prettierignore`d and hand-formatted;
 * these assertions are on content, never on layout.
 */
const repoText = (name: string) => readFileSync(fileURLToPath(new URL(`../${name}`, import.meta.url)), "utf8");

const PAGES = ["public/privacy.html", "public/terms.html"];

/** The dev manifests are the source: `build-manifest.mjs` derives the prod pair from them. */
const MANIFESTS = ["manifest.xml", "manifest-excel.xml"];

/**
 * `<Host Name="...">` is the manifest's word for an application; the pages use
 * the application's name. Only the hosts a task-pane add-in can claim are
 * listed — an unmapped one throws below rather than passing quietly, because a
 * host this table has never heard of is exactly the case that just bit.
 */
const HOST_APP: Record<string, string> = {
  Presentation: "PowerPoint",
  Workbook: "Excel",
  Document: "Word",
  Notebook: "OneNote",
  Project: "Project",
  Mailbox: "Outlook",
};

/** Every application either page could plausibly name. Used for the reverse check. */
const ALL_APPS = [...new Set(Object.values(HOST_APP))];

function hostsShipped(): string[] {
  const apps = new Set<string>();
  for (const name of MANIFESTS) {
    for (const [, host] of repoText(name).matchAll(/<Host\s+Name="([^"]+)"/g)) {
      const app = HOST_APP[host];
      if (!app) throw new Error(`${name} claims host "${host}", which this test has no application name for`);
      apps.add(app);
    }
  }
  return [...apps].sort();
}

describe("the privacy and terms pages the store listing submits", () => {
  /**
   * The premise. If the manifests stop naming a host — or the store listing
   * stops submitting these two files — every assertion below is auditing the
   * wrong thing, and should say so rather than keep passing.
   */
  it("still ships manifests for at least PowerPoint and Excel", () => {
    expect(hostsShipped()).toEqual(["Excel", "PowerPoint"]);
  });

  it.each(PAGES)("%s is still the page the listing submits", (page) => {
    const listing = repoText("docs/STORE-LISTING.md");
    const url = `https://ssf-chart.struktureretsundfornuft.dk/${page.replace("public/", "")}`;
    expect(listing, `docs/STORE-LISTING.md no longer submits ${url} — re-point these checks`).toContain(url);
  });

  /**
   * THE OMISSION. A host whose manifest ships but whose name appears on neither
   * page is data handling the user was never told about — the defect this file
   * was written for.
   */
  it.each(PAGES)("%s names every Office application the repo ships a manifest for", (page) => {
    const text = repoText(page);
    for (const app of hostsShipped()) {
      expect(
        text.includes(app),
        `${page} never mentions ${app}, but this repo store-preps an add-in for it. A user installing ` +
          `the ${app} add-in is handed a policy that describes a different product.`,
      ).toBe(true);
    }
  });

  /**
   * THE OVERCLAIM, the same defect pointing the other way. Describing handling
   * in an application this repo does not ship is a promise nothing keeps, and
   * it is the failure mode a copy-paste of a bigger vendor's policy produces.
   */
  it.each(PAGES)("%s claims no Office application the repo does not ship", (page) => {
    const text = repoText(page);
    const shipped = hostsShipped();
    for (const app of ALL_APPS) {
      if (shipped.includes(app)) continue;
      expect(
        new RegExp(`\\b${app}\\b`).test(text),
        `${page} mentions ${app}, but no manifest in this repo targets it. The pages must describe what ships.`,
      ).toBe(false);
    }
  });

  /**
   * THE ONE SENTENCE IN THE POLICY THAT CODE CAN FALSIFY.
   *
   * `/privacy.html` tells the reader that loading the add-in's own files, one
   * `build.json`, and Microsoft's Office.js are "the only network requests
   * either add-in makes". `docs/PUBLISHING.md` files that claim under "verified
   * fine, so nobody re-checks them" — which is precisely how a claim goes stale.
   *
   * `src/taskpane/app.ts` already records that the `build.json` GET is "the
   * first outbound request this add-in has ever made". The second one will
   * arrive in an ordinary feature PR, and the policy will not be part of that
   * diff. This is what notices.
   */
  it("keeps the policy's 'only network requests' sentence true of the shipped source", () => {
    const files = readdirSync("src", { recursive: true })
      .map(String)
      .filter((f) => f.endsWith(".ts") || f.endsWith(".html"))
      // `split`/`join`, not `replaceAll`: this repo's tsconfig `lib` predates
      // es2021, so `replaceAll` does not typecheck here.
      .map((f) => `src/${f.split("\\").join("/")}`);
    expect(files.length, "no sources found — is the walk pointed at the right place?").toBeGreaterThan(10);

    const fetches: string[] = [];
    const others: string[] = [];
    for (const file of files) {
      const text = repoText(file);
      // Only the CALL matters; a comment discussing fetch is not a request. The
      // pane's one live call is `fetch(\`/build.json?t=${Date.now()}\`, …)`.
      for (const [, target] of text.matchAll(/(?<!\/[/*].*)\bfetch\(\s*[`"']([^`"']*)/g)) {
        fetches.push(`${file}: fetch(${target})`);
      }
      for (const [, api] of text.matchAll(/\b(XMLHttpRequest|sendBeacon|EventSource|new WebSocket)\b/g)) {
        others.push(`${file}: ${api}`);
      }
    }

    expect(
      others,
      `a transport the privacy policy does not mention appears in src/. Either drop it or amend ` +
        `public/privacy.html — the policy says these are the only network requests the add-ins make.`,
    ).toEqual([]);
    for (const site of fetches) {
      expect(
        site,
        `a new outbound request: ${site}. public/privacy.html discloses exactly one (\`/build.json\`, ` +
          `same-origin, sending nothing). Amend the policy in the same change that adds the call.`,
      ).toMatch(/fetch\(\/build\.json/);
    }
    // A guard that stops finding its own subject has stopped guarding: if the
    // stale-build check is ever removed, the disclosure should go with it.
    expect(fetches, "the disclosed /build.json request is gone from src/ — the policy still describes it").toHaveLength(
      1,
    );
  });
});
