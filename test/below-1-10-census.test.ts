import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * WHAT A HOST BELOW PowerPointApi 1.10 LOSES — kept honest by re-derivation.
 *
 * This census is not an internal curiosity: it is printed in `STORE-LISTING.md`
 * as the reason the words "never flat pictures" were struck from the store copy,
 * and in `BACKLOG.md` item 10 as the evidence the owner decided WARN AND PICTURE
 * on. It is therefore a number a reviewer could check against the product.
 *
 * IT WAS PUBLISHED WRONG FOR A FORTNIGHT. From 2026-08-31 to 2026-09-13 both
 * documents said EIGHT charts lose their subject. It is nine. Showcase #107 is
 * a `radar` with `bars: true`, and its eight radial bars are eight `wedge`
 * nodes — so it was filed with gantt and waterfall under the charts that "lose
 * annotation arrows and keep their marks", and it has no arrows at all. Nothing
 * caught it because nothing re-computed it; the figure was copied forward.
 *
 * So this file asserts the census AND asserts the documents still say it. A
 * chart added, removed or re-decorated fails here and names the document to fix,
 * rather than silently making a published claim false.
 *
 * The gate below 1.10 is `canRotate()` in `src/render/powerpoint.ts`, which
 * makes `addWedgeFan` and the arrowhead case trace and return nothing. The two
 * node kinds counted here are exactly what `marksThisHostWillDrop` counts.
 *
 * WHY IT SPLITS IN TWO. A chart that loses `wedge` nodes loses its SUBJECT —
 * the wedge IS the pie, the doughnut, the sunburst ring, the radial bar — and
 * keeps only its labels. A chart that loses only `arrowhead` nodes keeps every
 * mark and loses an annotation. Those are different promises to a user, and
 * collapsing them into one "18" is what let the radar sit on the wrong side.
 */

const ROOT = resolve(__dirname, "..");
const showcase: ChartConfig[] = JSON.parse(readFileSync(resolve(ROOT, "examples/showcase.json"), "utf8"));

type Census = {
  total: number;
  loseInk: number;
  subject: Record<string, number>;
  annotation: Record<string, number>;
  subjectCount: number;
  annotationCount: number;
};

function census(): Census {
  const subject: Record<string, number> = {};
  const annotation: Record<string, number> = {};
  for (const cfg of showcase) {
    const scene = buildChart(cfg);
    const wedges = scene.nodes.filter((n) => n.kind === "wedge").length;
    const arrows = scene.nodes.filter((n) => n.kind === "arrowhead").length;
    if (wedges) subject[cfg.kind] = (subject[cfg.kind] ?? 0) + 1;
    else if (arrows) annotation[cfg.kind] = (annotation[cfg.kind] ?? 0) + 1;
  }
  const sum = (o: Record<string, number>) => Object.values(o).reduce((a, b) => a + b, 0);
  return {
    total: showcase.length,
    subject,
    annotation,
    subjectCount: sum(subject),
    annotationCount: sum(annotation),
    loseInk: sum(subject) + sum(annotation),
  };
}

describe("what a host below PowerPointApi 1.10 loses", () => {
  it("counts 18 of 123 charts losing ink, split 9 subject and 9 annotation", () => {
    const c = census();
    expect(c.total).toBe(123);
    expect(c.loseInk).toBe(18);
    expect(c.subjectCount).toBe(9);
    expect(c.annotationCount).toBe(9);
  });

  it("names which kinds lose their subject, radar among them", () => {
    const c = census();
    // radar 1/5 is the entry that was mis-filed. Asserting the whole map rather
    // than just the total is deliberate: a compensating pair of errors — one
    // chart moving each way — would keep the total at 9 and still publish a
    // false breakdown, which is the exact failure this file exists for.
    expect(c.subject).toEqual({ pie: 4, doughnut: 2, sunburst: 2, radar: 1 });
    expect(c.annotation).toEqual({ gantt: 4, stacked: 3, waterfall: 1, scatter: 1 });
  });

  it("the radial-bar radar loses wedges and has no arrows at all", () => {
    // The specific claim the corrected documents now make. If `bars: true`
    // ever stops drawing wedges this must be re-read, not re-pinned.
    const radial = showcase.findIndex(
      (c) => c.kind === "radar" && (c as { radar?: { bars?: boolean } }).radar?.bars === true,
    );
    expect(radial, "no radial-bar radar in the showcase any more").toBeGreaterThanOrEqual(0);
    const scene = buildChart(showcase[radial]);
    expect(scene.nodes.filter((n) => n.kind === "wedge").length).toBe(8);
    expect(scene.nodes.filter((n) => n.kind === "arrowhead").length).toBe(0);
  });

  it("the store listing and the backlog publish the number this file measures", () => {
    const c = census();
    const listing = readFileSync(resolve(ROOT, "docs/STORE-LISTING.md"), "utf8");
    const backlog = readFileSync(resolve(ROOT, "docs/BACKLOG.md"), "utf8");

    expect(
      listing,
      "STORE-LISTING.md no longer states the measured loss — update it or this claim is unsourced",
    ).toContain(`${c.loseInk} of ${c.total} shipped charts lose ink`);
    expect(listing, "STORE-LISTING.md publishes a different subject count").toContain(`and ${c.subjectCount} lose`);
    expect(backlog, "BACKLOG.md item 10 publishes a different subject count").toContain(
      `those **${c.subjectCount} keep only their labels**`,
    );
    expect(backlog, "BACKLOG.md item 10's heading publishes a different subject count").toContain(
      `${c.loseInk} of ${c.total} charts, ${c.subjectCount} of them wholly`,
    );
  });
});
