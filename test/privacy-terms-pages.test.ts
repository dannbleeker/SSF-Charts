import { describe, expect, it } from "vitest";
import { readdirSync, readFileSync } from "fs";
import { fileURLToPath } from "url";
import ts from "typescript";

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

/**
 * EVERY WAY A SOURCE FILE REACHES FOR `fetch`, read off the syntax tree.
 *
 * This used to be a regex, `fetch(` followed by a quote or a backtick, with a
 * look-behind that skipped anything after a `//` on the same line. It saw only a
 * call whose first argument is a string literal, which is the one shape the
 * pane's own request happens to have and the LEAST common way to write a second
 * one: `fetch(ENDPOINT, …)`, `fetch(new URL(…))`, `window.fetch(url)` and
 * `const f = fetch` all slipped past it, and so did a template literal on a line
 * that already held an `https://` string, whose `//` the look-behind took for a
 * comment. The test header says the second request "will arrive in an ordinary
 * feature PR" and that this "is what notices"; every one of those left it green.
 *
 * The tree has no comments or strings in it to confuse, so a call is a call:
 *
 *   - `fetch(…)`, `x.fetch(…)` and `x["fetch"](…)` are calls, reported with what
 *     their first argument is — a literal's text, a template's leading text, or
 *     the expression's own source for anything that is not a literal;
 *   - any OTHER mention of `fetch` as a value (`const f = fetch`, `fetch.call`,
 *     passing it on) is a reference, which is a call this cannot see the target
 *     of and so cannot clear.
 *
 * `inline` is the markup case: the `<script>` bodies and `on…=` handlers of an
 * `.html` file, parsed as script, because the old regex read those too.
 */
export function fetchSites(file: string, source: string, inline = false): string[] {
  const bodies = inline
    ? [
        ...[...source.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map((m) => m[1]),
        ...[...source.matchAll(/\bon[a-z]+="([^"]*)"/g)].map((m) => m[1]),
      ]
    : [source];
  const sites: string[] = [];
  for (const body of bodies) {
    const sf = ts.createSourceFile(file, body, ts.ScriptTarget.ES2020, true);
    const isFetchName = (n: ts.Node): boolean =>
      (ts.isIdentifier(n) && n.text === "fetch") ||
      (ts.isPropertyAccessExpression(n) && n.name.text === "fetch") ||
      (ts.isElementAccessExpression(n) &&
        ts.isStringLiteralLike(n.argumentExpression) &&
        n.argumentExpression.text === "fetch");
    const describeArg = (arg: ts.Expression | undefined): string => {
      if (!arg) return "<no argument>";
      if (ts.isStringLiteralLike(arg)) return arg.text;
      if (ts.isTemplateExpression(arg)) return `${arg.head.text}\${…}`;
      return `<${arg.getText(sf).slice(0, 60)}>`;
    };
    const visit = (node: ts.Node): void => {
      if (ts.isCallExpression(node) && isFetchName(node.expression)) {
        const callee = node.expression;
        sites.push(`${file}: fetch(${describeArg(node.arguments[0])})`);
        // The callee is accounted for; its arguments, and the `window` in
        // `window.fetch(…)`, may hold more — but not the `.fetch` just reported.
        node.arguments.forEach(visit);
        if (ts.isPropertyAccessExpression(callee) || ts.isElementAccessExpression(callee)) visit(callee.expression);
        return;
      }
      if (ts.isIdentifier(node) && node.text === "fetch" && isReference(node))
        sites.push(`${file}: fetch referenced, not called`);
      ts.forEachChild(node, visit);
    };
    visit(sf);
  }
  return sites;
}

/** Is this `fetch` identifier a VALUE being used, rather than a name being declared or a member name? */
function isReference(id: ts.Identifier): boolean {
  const p = id.parent;
  // `typeof globalThis.fetch` is erased before anything runs.
  for (let up: ts.Node | undefined = p; up && !ts.isStatement(up); up = up.parent) if (ts.isTypeNode(up)) return false;
  if (ts.isPropertyAccessExpression(p) && p.name === id) return true;
  const declares =
    (ts.isParameter(p) ||
      ts.isVariableDeclaration(p) ||
      ts.isPropertyAssignment(p) ||
      ts.isPropertyDeclaration(p) ||
      ts.isPropertySignature(p) ||
      ts.isMethodDeclaration(p) ||
      ts.isMethodSignature(p) ||
      ts.isFunctionDeclaration(p) ||
      ts.isBindingElement(p) ||
      ts.isImportSpecifier(p) ||
      ts.isExportSpecifier(p)) &&
    p.name === id;
  return !declares;
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
      fetches.push(...fetchSites(file, text, file.endsWith(".html")));
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

  /**
   * THE GUARD ABOVE HAS TO BE ABLE TO FAIL ON A SECOND REQUEST IN ANY SPELLING.
   * Each of these is a call the old regex let through; they are run through the
   * same function the real test uses, so a regression in how it reads a call
   * fails here by name instead of by the policy quietly going false.
   */
  it.each([
    ["an identifier as the URL", 'const T = "https://c.example.com/e"; await fetch(T, { method: "POST" });'],
    [
      "a template on a line that already holds https://",
      'const base = "https://c.example.com"; await fetch(`${base}/e`, { method: "POST" });',
    ],
    ["a URL object", 'await fetch(new URL("/e", "https://c.example.com"));'],
    ["window.fetch", "await window.fetch(endpoint, { method: 'POST' });"],
    ["globalThis.fetch", "await globalThis.fetch(endpoint);"],
    ["a computed member", 'await self["fetch"](endpoint);'],
    ["fetch with no argument", "await fetch();"],
    ["a literal URL that is not build.json", 'await fetch("https://c.example.com/e");'],
    ["fetch aliased", "const f = fetch; await f(endpoint);"],
    ["fetch handed on", "send(fetch);"],
    ["fetch.call", "await fetch.call(globalThis, endpoint);"],
  ])("sees %s", (_name, code) => {
    const sites = fetchSites("src/x.ts", code);
    expect(sites, `${code} passed the network guard`).not.toEqual([]);
    expect(
      sites.some((s) => !/fetch\(\/build\.json/.test(s)),
      `${code} was read as the disclosed request`,
    ).toBe(true);
  });

  it("reads the disclosed request as exactly that, and prose and declarations as nothing", () => {
    expect(fetchSites("src/x.ts", "await fetch(`/build.json?t=${Date.now()}`, { cache: 'no-store' });")).toEqual([
      "src/x.ts: fetch(/build.json?t=${…})",
    ]);
    expect(
      fetchSites(
        "src/x.ts",
        [
          "// await fetch(url) is how the old version did it",
          "/* fetch(`/elsewhere`) */",
          "const msg = 'failed to fetch dynamically imported module';",
          "const re = /failed to fetch/i;",
          "function wrap(fetch: typeof globalThis.fetch) { return { fetch: 1 }; }",
        ].join("\n"),
      ),
    ).toEqual([]);
  });

  it("reads inline scripts and handlers in markup", () => {
    const html = '<script src="./app.ts"></script><button onclick="fetch(url)">x</button><script>fetch(`/x`)</script>';
    expect(fetchSites("src/p.html", html, true).sort()).toEqual(["src/p.html: fetch(/x)", "src/p.html: fetch(<url>)"]);
  });
});
