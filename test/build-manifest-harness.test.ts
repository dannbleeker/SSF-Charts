import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import {
  harnessManifest,
  withHarness,
  // @ts-expect-error — a .mjs tool script with no types, imported for its pure helpers.
} from "../scripts/build-manifest.mjs";

/**
 * The harness manifest: production, plus `?harness=1` on every task-pane URL.
 *
 * WHY IT MATTERS ENOUGH TO TEST. `app.ts` can hide the Automation ▸ Testing
 * section — demo deck, self-test, host probe, download run log, clean up the
 * last round — unless the pane is opened with `harness=1`, and its own comment
 * says why that matters for a store listing: "A stranger cannot parse it, and
 * two of those buttons change their document."
 *
 * The gate ships OFF, because flipping it alone would hide the section from the
 * round driver too and stop the loop. This manifest is what makes the flip a
 * one-line change — so if it silently stopped carrying the parameter, the flip
 * would look safe and would end the round loop instead.
 */

describe("the harness manifest", () => {
  it("adds the parameter with the right separator for each URL shape", () => {
    // Both shapes occur in the real manifest: bare task-pane URLs and ribbon
    // deep links that already carry a query.
    expect(withHarness("https://x/src/taskpane/taskpane.html")).toBe("https://x/src/taskpane/taskpane.html?harness=1");
    expect(withHarness("https://x/src/taskpane/taskpane.html?kind=stacked")).toBe(
      "https://x/src/taskpane/taskpane.html?kind=stacked&amp;harness=1",
    );
    // `&amp;` and not `&` — this is XML, and the deep links already read
    // `?tab=elements&amp;el=harvey`. A bare ampersand would not parse.
    expect(withHarness("https://x/src/taskpane/taskpane.html?tab=elements&amp;el=harvey")).toContain("&amp;harness=1");
  });

  it("touches ONLY task-pane URLs", () => {
    // Icons and the support URL are not pane loads. Rewriting them would be a
    // change with no meaning attached, and one of them points at GitHub.
    expect(withHarness("https://x/assets/icon-32.png")).toBe("https://x/assets/icon-32.png");
    expect(withHarness("https://github.com/dannbleeker/SSF-Charts")).toBe("https://github.com/dannbleeker/SSF-Charts");
    expect(withHarness("https://x/src/excel/excel.html")).toBe("https://x/src/excel/excel.html");
  });

  it("is idempotent, so a re-run cannot double the parameter", () => {
    const once = withHarness("https://x/src/taskpane/taskpane.html?kind=pie");
    expect(withHarness(once)).toBe(once);
    expect(once.match(/harness=1/g)).toHaveLength(1);
  });

  it("is the production manifest in every respect but that parameter", () => {
    // THE PROPERTY THAT KEEPS IT HONEST. The round loop's whole claim is that it
    // validates the bytes users get; a harness manifest that drifted in GUID,
    // requirement set or ribbon layout would quietly break that.
    const prod = readFileSync("manifest-prod.xml", "utf8");
    const harness = readFileSync("manifest-harness.xml", "utf8");
    expect(harnessManifest(prod), "manifest-harness.xml is stale").toBe(harness);
    // Strip the parameter back out and the two must be identical, byte for byte.
    expect(harness.split("&amp;harness=1").join("").split("?harness=1").join("")).toBe(prod);
  });

  it("carries the parameter on every task-pane URL, not just the first", () => {
    const harness = readFileSync("manifest-harness.xml", "utf8");
    const paneUrls = [...harness.matchAll(/DefaultValue="([^"]*taskpane\.html[^"]*)"/g)].map((m) => m[1]);
    expect(paneUrls.length, "no task-pane URLs found at all").toBeGreaterThan(5);
    for (const u of paneUrls) expect(u, u).toMatch(/[?&](amp;)?harness=1/);
  });

  it("leaves the PRODUCTION manifests free of it", () => {
    // The one that would actually hurt: shipping `harness=1` to users would
    // expose the testing UI to precisely the reviewer it is meant to be hidden
    // from, which is the opposite of the point.
    for (const f of ["manifest-prod.xml", "manifest-excel-prod.xml"]) {
      expect(readFileSync(f, "utf8"), f).not.toContain("harness");
    }
  });
});
