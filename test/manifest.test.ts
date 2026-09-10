import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
// @ts-expect-error — plain .mjs tools with no types. The rules live THERE so the
// weekly published-install sweep and this test cannot drift apart.
import { checkManifest, urlsIn } from "../scripts/manifest-rules.mjs";
// @ts-expect-error — as above.
import { judgePublished, reportBody } from "../scripts/check-published-install.mjs";

/**
 * The manifests, checked offline against the rules that matter.
 *
 * CI runs Microsoft's own `office-addin-manifest validate` in a job of its own,
 * and that is the authority. This is not a second copy of it: it pins the small
 * number of rules whose violation this project has actually met, so they cannot
 * come back when the validator is unreachable — which is not hypothetical. The
 * validator calls a Microsoft SERVICE, so it cannot run in a sandbox with no
 * route to it, and for the whole life of this repo that meant the manifests were
 * never validated at all.
 *
 * The first thing the CI job found on its first run: `<Version>0.1.0</Version>`.
 * "Manifest Version Too Low: The manifest has unsupported version number less
 * than 1.0" — an error, on all four manifests, since the day they were written.
 * It passed every test in this repo, because nothing here had ever looked.
 */
const MANIFESTS = ["manifest.xml", "manifest-excel.xml", "manifest-prod.xml", "manifest-excel-prod.xml"];

const read = (name: string) => readFileSync(fileURLToPath(new URL(`../${name}`, import.meta.url)), "utf8");

describe("the add-in manifests", () => {
  // The rules live in `scripts/manifest-rules.mjs` because a SECOND caller needs
  // them: the file a user downloads is not this file, and the two have diverged
  // twice. See `scripts/check-published-install.mjs`.
  it.each(MANIFESTS)("%s satisfies every rule this project has been bitten by", (name) => {
    expect(checkManifest(read(name), name)).toEqual([]);
  });

  it("would catch each of those rules being broken", () => {
    // The rules are only worth importing if they can still fail. Every branch
    // below is a thing this repo has actually shipped or nearly shipped.
    const good = read("manifest-prod.xml");
    // "Manifest Version Too Low" — an error on all four manifests since the day
    // they were written, found by Microsoft's validator on its first CI run.
    expect(
      checkManifest(good.replace(/<Version>[^<]+<\/Version>/, "<Version>0.1.0</Version>"), "manifest-prod.xml"),
    ).toEqual([expect.stringContaining("below 1.0")]);
    // A changed GUID is a different add-in: every sideload orphaned.
    expect(
      checkManifest(
        good.replace(/<Id>[^<]+<\/Id>/, "<Id>00000000-0000-0000-0000-000000000000</Id>"),
        "manifest-prod.xml",
      ),
    ).toEqual([expect.stringContaining("orphaned")]);
    // A prod manifest can be perfectly CURRENT and full of localhost, if the
    // rewrite in build-manifest.mjs ever stops matching.
    expect(
      checkManifest(good.replace("ssf-chart.struktureretsundfornuft.dk", "localhost:3000"), "manifest-prod.xml"),
    ).toEqual([expect.stringContaining("localhost")]);
    expect(checkManifest("<html>not a manifest</html>", "manifest-prod.xml")).toEqual([
      expect.stringContaining("not an Office add-in manifest"),
    ]);
  });
});

/**
 * The install path the README hands a user, which nothing else here can see.
 *
 * Every other gate in this repo reads the working tree. A user downloads the
 * asset attached to the latest RELEASE, and those two have now diverged twice —
 * v0.1.0 shipped the dev manifests, and v0.3.0 shipped the `<Version>0.1.0</Version>`
 * that Microsoft's validator rejects, a fix that landed in the repo on
 * 2026-08-06 and has still reached nobody.
 *
 * The network half runs weekly in `quality-sweep.yml`. This is the half that
 * decides what it found, and it is the half that was missing: a byte comparison,
 * not just a rules check, because the release was VALID when it was cut and went
 * stale when main moved past it.
 */
describe("the published install path", () => {
  const committed = read("manifest-prod.xml");

  it("passes when the release carries exactly what is committed", () => {
    expect(judgePublished([{ name: "manifest-prod.xml", published: committed, committed }])).toEqual([]);
    // Line endings alone are not a finding — a release asset that has been
    // through a checkout elsewhere can differ by \r, and crying wolf about that
    // trains the reader to ignore this check inside a fortnight.
    const crlf = committed.replace(/\n/g, "\r\n");
    expect(judgePublished([{ name: "manifest-prod.xml", published: crlf, committed }])).toEqual([]);
  });

  it("says so when the release is missing the file the README names", () => {
    // v0.1.0, verbatim: the only documented install path was a 404 for twelve
    // days while release.yml sat correct and un-run.
    expect(judgePublished([{ name: "manifest-prod.xml", published: null, committed }])).toEqual([
      expect.stringContaining("not in the latest release"),
    ]);
  });

  it("says so when main has moved past the release", () => {
    // The live case, and the one a rules check alone cannot see: the published
    // manifest is a perfectly well-formed manifest. It is just the OLD one.
    const stale = committed.replace(/<Version>[^<]+<\/Version>/, "<Version>1.0.0.1</Version>");
    const problems = judgePublished([{ name: "manifest-prod.xml", published: stale, committed }]);
    expect(problems).toEqual([expect.stringContaining("NOT the committed one")]);
    expect(problems[0]).toContain("Cut a release");
  });

  it("reports a published manifest Office would reject, whatever main says", () => {
    const rejected = committed.replace(/<Version>[^<]+<\/Version>/, "<Version>0.1.0</Version>");
    const problems = judgePublished([{ name: "manifest-prod.xml", published: rejected, committed }]);
    expect(problems.some((p: string) => p.includes("below 1.0"))).toBe(true);
  });

  it("never reads a clean sweep as a broken one", () => {
    expect(reportBody([], "v9.9.9")).toContain("sound");
    expect(reportBody(["something"], "v0.3.0")).toContain("v0.3.0");
    expect(reportBody(["something"], "v0.3.0")).toContain("Cut a release");
  });
});

/**
 * Which URLs the sweep will actually go and fetch.
 *
 * The first version filtered namespaces with `u.includes("schemas.microsoft.com")`,
 * and CodeQL failed the PR for it — `js/incomplete-url-substring-sanitization`,
 * high. The consequence here is not an injection: it is that any real URL
 * carrying that string anywhere would be silently excused from being checked,
 * which is the sanitiser deciding what the checker gets to look at.
 */
describe("the URLs a manifest asks the host to fetch", () => {
  it("drops namespaces by hostname, and is not fooled by one in a query string", () => {
    // Asserted on HOSTNAMES, not substrings — CodeQL flags a substring test
    // against a URL wherever it appears, and it is right to: the assertion would
    // pass for a host that merely carries the string. Same rule as the code.
    const hosts = urlsIn(read("manifest-prod.xml")).map((u: string) => new URL(u).hostname);
    expect(hosts).not.toContain("schemas.microsoft.com");
    expect(hosts).toContain("ssf-chart.struktureretsundfornuft.dk");
    // The substring test excused this one. It is a real host and must be checked.
    expect(urlsIn('"https://evil.example/x?ref=schemas.microsoft.com"')).toEqual([
      "https://evil.example/x?ref=schemas.microsoft.com",
    ]);
    // And a subdomain is not the namespace host either.
    expect(urlsIn('"https://schemas.microsoft.com.evil.example/x"')).toEqual([
      "https://schemas.microsoft.com.evil.example/x",
    ]);
  });
});

/**
 * THE TESTING SECTION AND THE MANIFESTS HAVE TO MOVE TOGETHER.
 *
 * The Automation tab's first half is the round-loop harness, and every published
 * build has shipped it to every user. `TESTING_UI_NEEDS_OPT_IN` in
 * `src/taskpane/app.ts` can hide it — but the SAME section is how `round.mjs`
 * drives a round: it clicks `Probe, then self-test` and waits on
 * `Download run log`, both by name, in the accessibility tree that `hidden`
 * removes them from.
 *
 * So flipping that constant without giving the driver a manifest that opts back
 * in does not fail loudly. It ends the round loop, quietly, and the next person
 * to notice is whoever wonders why no round has been archived in a week. That is
 * this repo's most-repeated defect — the fix that reached all but one call site
 * — and this is the test that refuses to let it happen again.
 */
describe("hiding the test harness from a published add-in", () => {
  const appSrc = read("src/taskpane/app.ts");
  const paneHtml = read("src/taskpane/taskpane.html");
  const optIn = /const TESTING_UI_NEEDS_OPT_IN = (true|false);/.exec(appSrc)?.[1];

  it("keeps the switch where a reader can find it", () => {
    expect(optIn, "`TESTING_UI_NEEDS_OPT_IN` is gone or no longer a plain boolean literal").toBeDefined();
  });

  it("targets a section that actually exists", () => {
    // The gate hides `#testing-section`. Rename the id in the HTML and the gate
    // silently stops hiding anything — a published add-in would then ship the
    // harness while this file claimed otherwise.
    expect(appSrc, "the gate no longer names the section it hides").toContain('getElementById("testing-section")');
    expect(paneHtml, "the section the gate hides is not in the pane").toContain('id="testing-section"');
    // And it must be the block that holds the harness, not some other section.
    const at = paneHtml.indexOf('id="testing-section"');
    expect(paneHtml.slice(at, at + 400), "`#testing-section` is not the Testing block").toContain("<h2>Testing</h2>");
  });

  it("turns the verbose-trace toggle off, not merely out of sight", () => {
    // THE TRAP THE FIRST DRAFT OF THIS GATE SHIPPED. `#demo-trace` lives inside
    // `#testing-section`, ships `checked`, and `wireInsert` reads it at boot to
    // call `setTracing(true)` — which reaches `enableExtendedErrorLogging`.
    // `hidden` takes an element out of the accessibility tree; it does not
    // uncheck a checkbox. So hiding the section on its own would leave verbose
    // tracing running for every user with its only switch invisible.
    expect(paneHtml, "the toggle is no longer inside the gated section — re-check this gate").toMatch(
      /id="testing-section"[\s\S]*id="demo-trace"/,
    );
    const gate = /if \(TESTING_UI_NEEDS_OPT_IN[\s\S]*?\n\}/.exec(appSrc)?.[0] ?? "";
    expect(gate, "the gate does not uncheck the trace toggle").toContain('getElementById("demo-trace")');
    expect(gate, "the gate does not stop tracing that has already started").toContain("setTracing(false)");
  });

  it("requires a harness manifest the moment the switch is flipped", () => {
    if (optIn !== "true") {
      // Not flipped: the harness is visible to everyone, which is today's
      // behaviour and the driver works. Nothing to require.
      expect(optIn).toBe("false");
      return;
    }
    // Flipped: at least one manifest must ask for the pane WITH the opt-in, or
    // no round can be driven against the deployment ever again.
    const withParam = MANIFESTS.filter((m) => /taskpane\.html\?[^"<]*harness=1/.test(read(m)));
    expect(
      withParam,
      "TESTING_UI_NEEDS_OPT_IN is true but no manifest opens the pane with `?harness=1` — " +
        "the round loop cannot reach `Probe, then self-test` and will stop without saying so",
    ).not.toEqual([]);
    // And the manifest a USER installs must NOT carry it, or nothing was gained.
    expect(
      /taskpane\.html\?[^"<]*harness=1/.test(read("manifest-prod.xml")),
      "the production manifest opts INTO the harness, so users still get it",
    ).toBe(false);
  });
});
