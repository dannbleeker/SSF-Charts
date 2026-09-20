import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { estimateOfficeShapes } from "../src/core/scene";
import { DEMO_SHAPE_BUDGET, wantsAutoPicture } from "../src/render/powerpoint";
import type { ChartConfig } from "../src/core/types";

/**
 * WHAT ARRIVES AS A PICTURE ON POWERPOINT ON THE WEB — kept honest by
 * re-derivation, and the sibling of `below-1-10-census.test.ts`.
 *
 * That file measures what a host BELOW PowerPointApi 1.10 loses. Since the
 * floor became 1.10 on 2026-09-13 no such host can install the add-in, so that
 * census now describes a population with no users. THIS one describes the
 * population that is every certification reviewer: PowerPoint on the web, where
 * a chart denser than `DEMO_SHAPE_BUDGET` goes onto the slide as a picture
 * rather than as native shapes (`wantsAutoPicture`, gated on `isWebHost()`).
 *
 * WHY IT IS A TEST AND NOT A COMMENT. The store listing and the manifest's
 * `<Description>` both make a native-shapes claim, and "functionality does not
 * match the offer description" is the ordinary AppSource rejection. Until
 * 2026-09-20 both promised native shapes with no qualifier, and the one
 * qualifier that had been written keyed off the build's AGE — a thing the 1.10
 * floor makes irrelevant, because a build that cannot draw the chart cannot
 * install the add-in at all.
 *
 * TWO POPULATIONS, AND AN AUDIT PRODUCED ONE ANSWER FOR EACH — "Area and Tile
 * map" and "9 of 123". Both are right; neither alone is the answer, and that is
 * exactly why both are asserted here:
 *
 * - `sampleConfig` is what the PICKER loads, so it is what a user or a reviewer
 *   meets. Two of the 25 kinds cross the budget at the default frame.
 * - `examples/showcase.json` is the shipped deck and the denominator the 1.10
 *   census already publishes. It holds denser instances than the samples do, in
 *   kinds whose samples sit well under.
 *
 * The gap between them IS the finding: the gate is a shape count, not a list of
 * kinds, so a user's own data can push a kind over that its sample never
 * approaches. Copy that names kinds has to be written knowing that.
 *
 * AND THE TWO NINES ARE DIFFERENT NINES. Nine charts lose their subject below
 * 1.10; nine charts arrive as a picture on the web. They share not one chart —
 * asserted below, because a document that conflates them would read as
 * consistent.
 *
 * The predicate and the budget are IMPORTED rather than restated. A test that
 * hard-codes 105 goes on passing when the budget moves and the listing goes
 * quietly wrong, which is the whole failure this file exists to prevent.
 */

const ROOT = resolve(__dirname, "..");
const showcase: ChartConfig[] = JSON.parse(readFileSync(resolve(ROOT, "examples/showcase.json"), "utf8"));

/** The insert path's own question, with the arguments PowerPoint on the web gives it. */
const picturedOnWeb = (cfg: ChartConfig): boolean =>
  wantsAutoPicture(estimateOfficeShapes(buildChart(cfg)), { web: true, canPicture: true, alreadyPicture: false });

/** Same question with the arguments a desktop host gives it. */
const picturedOnDesktop = (cfg: ChartConfig): boolean =>
  wantsAutoPicture(estimateOfficeShapes(buildChart(cfg)), { web: false, canPicture: true, alreadyPicture: false });

describe("what arrives as a picture on PowerPoint on the web", () => {
  it("names the 2 of 25 picker samples that cross the budget", () => {
    const over = CHART_KINDS.filter(({ kind }) => picturedOnWeb(sampleConfig(kind))).map(({ label }) => label);
    expect(CHART_KINDS.length, "the listing publishes 25 chart kinds").toBe(25);
    // Labels, not kind ids, because these are the words the listing and the
    // notes-for-certification draft put in front of a reviewer.
    expect(over).toEqual(["Area", "Tile map"]);
  });

  it("counts 9 of the 123 shipped charts, across five kinds", () => {
    const over = showcase.filter(picturedOnWeb);
    expect(showcase.length).toBe(123);
    expect(over.length).toBe(9);
    const byKind: Record<string, number> = {};
    for (const c of over) byKind[c.kind] = (byKind[c.kind] ?? 0) + 1;
    // Asserting the breakdown and not just the total, for the reason
    // `below-1-10-census.test.ts` gives: a compensating pair of errors keeps a
    // total right and publishes a false list of kinds.
    expect(byKind).toEqual({ tilemap: 4, area: 2, waffle: 1, line: 1, combo: 1 });
  });

  it("is a web-only rescue — the desktop apps picture nothing", () => {
    // The old copy qualified the native-shapes claim with "on the web", which is
    // backwards: the web is the ONLY host that rasterises for density.
    expect(showcase.some(picturedOnDesktop)).toBe(false);
    expect(CHART_KINDS.some(({ kind }) => picturedOnDesktop(sampleConfig(kind)))).toBe(false);
  });

  it("the web nine and the below-1.10 nine share no chart", () => {
    // The collision this file was written to stop: two different measurements
    // that both answer "nine of 123".
    const dense = new Set(showcase.map((c, i) => (picturedOnWeb(c) ? i : -1)).filter((i) => i >= 0));
    const losesSubject = new Set(
      showcase.map((c, i) => (buildChart(c).nodes.some((n) => n.kind === "wedge") ? i : -1)).filter((i) => i >= 0),
    );
    expect(losesSubject.size, "below-1-10-census.test.ts measures nine of these").toBe(9);
    expect([...dense].filter((i) => losesSubject.has(i))).toEqual([]);
  });
});

/**
 * The listing as one long line.
 *
 * Markdown hard-wraps at 80 columns and prefixes the note block with `> `, so a
 * sentence this file asserts can be split anywhere — and a raw `toContain`
 * would then fail on a REFLOW rather than on a changed claim, which trains the
 * next reader to loosen the assertion instead of fixing the document.
 */
const flat = (md: string): string => md.replace(/\n\s*>?\s*/g, " ").replace(/\s+/g, " ");

describe("the documents publish what this file measures", () => {
  const listing = flat(readFileSync(resolve(ROOT, "docs/STORE-LISTING.md"), "utf8"));
  const manifest = readFileSync(resolve(ROOT, "manifest.xml"), "utf8");

  it("the store listing states both measured figures and the budget they rest on", () => {
    const over = showcase.filter(picturedOnWeb).length;
    expect(
      listing,
      "STORE-LISTING.md no longer states the shipped-deck figure — update it or the claim is unsourced",
    ).toContain(`${over} of ${showcase.length} shipped charts insert as a picture`);
    expect(listing, "STORE-LISTING.md no longer states the picker figure").toContain("2 of the 25 kinds");
    // The budget is the whole reason either figure is what it is. A raise that
    // left the prose alone would change both numbers silently.
    expect(listing, "STORE-LISTING.md quotes a budget the code no longer uses").toContain(
      `DEMO_SHAPE_BUDGET\` (${DEMO_SHAPE_BUDGET})`,
    );
  });

  it("the listing no longer promises anything to a build that cannot install", () => {
    // The floor is 1.10 and it governs installability, so a sentence about what
    // "older builds get" describes an experience nobody can have. Checked as a
    // substring because that is the exact phrasing that shipped.
    expect(listing.split("MUST NOT COME BACK")[0], "the live copy promises older builds a degraded run").not.toMatch(
      /older builds get a complete chart/,
    );
    expect(manifest).toContain('<Set Name="PowerPointApi" MinVersion="1.10" />');
  });

  it("the manifest description is qualified and fits Microsoft's 250 characters", () => {
    const description = /<Description DefaultValue="([^"]*)"/.exec(manifest)?.[1] ?? "";
    expect(description, "manifest.xml has no <Description>").not.toBe("");
    // The limit is Microsoft's, not ours; this is the field AppSource shows.
    expect(description.length).toBeLessThanOrEqual(250);
    expect(
      description,
      "the manifest description promises native shapes with no qualifier — the web pictures the densest charts",
    ).toMatch(/picture/i);
  });
});
