import { describe, expect, it } from "vitest";
import {
  demoItems,
  buildResultsScene,
  buildResultsScenes,
  type ResultRow,
  type ResultsSummary,
} from "../src/core/demo";
import { CHART_KINDS } from "../src/core/samples";
import { sceneToSvg } from "../src/render/svg";
import { estimateOfficeShapes, shiftNodeX } from "../src/core/scene";
import type { SceneNode } from "../src/core/scene";

describe("demo deck", () => {
  const items = demoItems();

  it("covers every chart kind plus feature and element slides", () => {
    // One slide per kind, plus the extra feature/element highlights.
    expect(items.length).toBeGreaterThan(CHART_KINDS.length);
    const titles = new Set(items.map((i) => i.title));
    expect(titles.has("Small multiples")).toBe(true);
    expect(titles.has("Agenda")).toBe(true);
    expect(titles.has("Forecast split")).toBe(true);
  });

  it("builds a non-empty, renderable scene for every slide", () => {
    for (const item of items) {
      expect(item.scene.nodes.length, `${item.title} has nodes`).toBeGreaterThan(0);
      // Every scene must render to SVG without producing NaN geometry.
      const svg = sceneToSvg(item.scene);
      expect(svg, `${item.title} renders clean`).not.toMatch(/NaN/);
    }
  });

  it("opens with a title slide and a contents/manifest table", () => {
    expect(items[0].title).toBe("Title");
    expect(items[1].title).toMatch(/^Contents/);
    // Title slide names the deck. Renamed with the product on 2026-08-27:
    // "PowerChart chart gallery" read as a stutter once the name changed.
    expect(items[0].scene.nodes.some((n) => n.kind === "text" && /ssf charts gallery/i.test(n.text))).toBe(true);
    // Every contents page stays within the harness page budget. Not the ~90
    // web budget it used to fit: at 79 shapes the full deck's single contents
    // page crashed PowerPoint on the web, while the 27-shape one did not.
    const contents = items.filter((i) => /^Contents/.test(i.title));
    for (const page of contents) expect(estimateOfficeShapes(page.scene)).toBeLessThanOrEqual(45);
    const idxText = contents.flatMap((p) => p.scene.nodes.filter((n) => n.kind === "text").map((n) => n.text));
    expect(idxText).toContain("Shapes");
    // Every chart is listed somewhere across the pages — pagination drops none.
    expect(idxText.some((t) => /Doughnut/.test(t))).toBe(true);
    expect(idxText.some((t) => /Stacked/.test(t))).toBe(true);
    expect(idxText.some((t) => /Table/.test(t))).toBe(true);
    // Neither structural slide is a re-editable chart.
    expect(items[0].configJson).toBeUndefined();
    expect(items[1].configJson).toBeUndefined();
  });

  it("stamps the running build AND host onto the title slide so a test PDF is self-identifying", () => {
    const stamped = demoItems({
      buildStamp: "abc1234 · 2026-07-17 20:00Z",
      host: "PowerPoint · OfficeOnline · 16.0.1",
    });
    const titleTexts = stamped[0].scene.nodes.filter((n) => n.kind === "text").map((n) => n.text);
    expect(titleTexts.some((t) => t.includes("abc1234 · 2026-07-17 20:00Z"))).toBe(true);
    expect(titleTexts.some((t) => t.includes("PowerPoint · OfficeOnline · 16.0.1"))).toBe(true);
    // Defaults (nothing passed) still render, with placeholders.
    const def = demoItems()[0].scene.nodes;
    expect(def.some((n) => n.kind === "text" && /Build local build/.test(n.text))).toBe(true);
    expect(def.some((n) => n.kind === "text" && /unknown host/.test(n.text))).toBe(true);
  });

  it("estimates the EXPANDED office shape count so wedge/polygon charts are budgeted honestly", () => {
    const scene = (t: string) => items.find((i) => i.title === t)!.scene;
    // The bug the self-check exposed: node count under-counts the render. A wedge
    // fans out and a polygon draws one line per edge, so these explode.
    // ~79 from 10 nodes. It was ~250 until the violin's KDE outlines were
    // thinned to a quarter-point tolerance — 246 polygon edges down to 72 — and
    // the PROPERTY this test is about, that a polygon costs far more than its
    // node, is untouched. The line below states it without a magnitude at all.
    expect(estimateOfficeShapes(scene("Violin"))).toBeGreaterThan(60);
    expect(estimateOfficeShapes(scene("Violin"))).toBeGreaterThan(scene("Violin").nodes.length * 5);
    // Sunburst is ~101. That was over the old 90 budget and therefore skipped;
    // since the budget rose to 105 on 2026-09-06 it is UNDER, and draws as
    // shapes. What this line is about — a wedge fan expanding one node into
    // many shapes — is unchanged. Only the consequence moved, so the comment
    // did too, and the second assertion pins the side of the line it is on.
    expect(estimateOfficeShapes(scene("Sunburst"))).toBeGreaterThan(90);
    expect(estimateOfficeShapes(scene("Sunburst")), "sunburst crossed 105 and is gated again").toBeLessThan(105);
    expect(estimateOfficeShapes(scene("Pie"))).toBeGreaterThan(scene("Pie").nodes.length); // wedge fan expands
    // A plain bar chart is one shape per node — no expansion, no over-count.
    expect(estimateOfficeShapes(scene("Stacked"))).toBe(scene("Stacked").nodes.length);
  });

  it("tags real charts with re-editable config JSON, and leaves elements untagged", () => {
    const charts = items.filter((i) => i.configJson);
    // Every chart kind is a tagged, re-editable config.
    expect(charts.length).toBeGreaterThanOrEqual(CHART_KINDS.length);
    for (const c of charts) expect(() => JSON.parse(c.configJson!)).not.toThrow();
    // The static element slides (agenda, KPI, …) are not chart configs.
    expect(items.some((i) => i.title === "Agenda" && !i.configJson)).toBe(true);
  });
});

describe("results slide", () => {
  const summary = (over: Partial<ResultsSummary> = {}): ResultsSummary => ({
    buildStamp: "abc1234 · 2026-07-17 20:00Z",
    items: 37,
    rendered: 28,
    skipped: 4,
    failed: 5,
    lost: 4,
    totalMs: 92_400,
    ...over,
  });

  it("shows a summary line, the build stamp, and total seconds", () => {
    const scene = buildResultsScene([], summary());
    const texts = scene.nodes.filter((n) => n.kind === "text").map((n) => n.text);
    expect(texts.some((t) => /37 items · 28 rendered · 4 skipped · 5 failed · 4 lost/.test(t))).toBe(true);
    expect(texts.some((t) => /total 92\.4s/.test(t))).toBe(true);
    expect(texts.some((t) => t.includes("abc1234 · 2026-07-17 20:00Z"))).toBe(true);
    expect(texts.some((t) => /Regression results/.test(t))).toBe(true);
  });

  it("tables ONLY the skipped/failed items, not the whole 37-row deck", () => {
    const rows: ResultRow[] = [
      { title: "Stacked", status: "rendered", shapes: 10, ms: 120 },
      { title: "Pie", status: "failed", shapes: 54, ms: 45012 },
      { title: "Area", status: "skipped", shapes: 120, ms: 2 },
    ];
    const scene = buildResultsScene(rows, summary({ items: 3, rendered: 1, skipped: 1, failed: 1, lost: 0 }));
    const texts = scene.nodes.filter((n) => n.kind === "text").map((n) => n.text);
    // The two problem rows are listed…
    expect(texts).toContain("Pie");
    expect(texts).toContain("Area");
    // …the clean one is not (it lives on the contents slide already).
    expect(texts).not.toContain("Stacked");
  });

  it("says so plainly on a clean run and stays under the ~90 web shape budget", () => {
    const rows: ResultRow[] = [{ title: "Stacked", status: "rendered", shapes: 10, ms: 120 }];
    const scene = buildResultsScene(rows, summary({ items: 1, rendered: 1, skipped: 0, failed: 0, lost: 0 }));
    const texts = scene.nodes.filter((n) => n.kind === "text").map((n) => n.text);
    expect(texts.some((t) => /All slides rendered cleanly/.test(t))).toBe(true);
    expect(estimateOfficeShapes(scene)).toBeLessThan(90);
    expect(sceneToSvg(scene)).not.toMatch(/NaN/);
  });

  it("stays under budget even with a full slate of failures", () => {
    const rows: ResultRow[] = Array.from({ length: 12 }, (_, i) => ({
      title: `Chart ${i}`,
      status: "failed" as const,
      shapes: 100 + i,
      ms: 45000,
    }));
    const scene = buildResultsScene(rows, summary({ items: 12, rendered: 0, skipped: 0, failed: 12, lost: 0 }));
    expect(estimateOfficeShapes(scene)).toBeLessThan(90);
  });

  it("paginates when the failure count would push one page past the budget", () => {
    // Presentation_2.pptx: 32 failures pushed the single-page scene to ~135
    // shapes; the run's own summary was the first casualty. Pagination splits
    // into ~20-row pages, each safely under the ~90 web shape budget.
    const rows: ResultRow[] = Array.from({ length: 32 }, (_, i) => ({
      title: `Chart ${i}`,
      status: "failed" as const,
      shapes: 50 + i,
      ms: 100 + i,
    }));
    const pages = buildResultsScenes(rows, summary({ items: 32, rendered: 0, skipped: 0, failed: 32, lost: 0 }));
    expect(pages.length).toBeGreaterThan(1);
    // Against the harness page budget, not the ~90 web one — see
    // HARNESS_PAGE_SHAPES for why a text-only slide needs the tighter number.
    for (const p of pages) expect(estimateOfficeShapes(p)).toBeLessThanOrEqual(45);
    // Every failure appears on some page — no rows dropped by the split.
    const titles = pages.flatMap((p) => p.nodes.filter((n) => n.kind === "text").map((n) => n.text));
    for (let i = 0; i < 32; i++) expect(titles).toContain(`Chart ${i}`);
    // Each page's title flags which page of the run it is.
    const first = pages[0].nodes.find((n) => n.kind === "text" && /Regression results/.test(n.text));
    expect(first && "text" in first ? first.text : "").toMatch(/page 1 of \d+/);
  });

  it("a clean or lightly-failing run still returns exactly one page", () => {
    const rows: ResultRow[] = Array.from({ length: 3 }, (_, i) => ({
      title: `Chart ${i}`,
      status: "failed" as const,
      shapes: 50 + i,
      ms: 100 + i,
    }));
    const pages = buildResultsScenes(rows, summary({ items: 3, rendered: 0, skipped: 0, failed: 3, lost: 0 }));
    expect(pages).toHaveLength(1);
    // No page-of-N marker on the single-page path.
    const title = pages[0].nodes.find((n) => n.kind === "text" && /Regression results/.test(n.text));
    expect(title && "text" in title ? title.text : "").toBe("Regression results");
  });

  it("notes recovered-on-retry items in the summary, and omits the note when none", () => {
    const withRetries = buildResultsScene([], summary({ retried: 2 }))
      .nodes.filter((n) => n.kind === "text")
      .map((n) => n.text);
    expect(withRetries.some((t) => /2 recovered/.test(t))).toBe(true);
    const noRetries = buildResultsScene([], summary())
      .nodes.filter((n) => n.kind === "text")
      .map((n) => n.text);
    expect(noRetries.some((t) => /recovered/.test(t))).toBe(false);
  });
});

/**
 * Moving a node horizontally, including the one coordinate that is not a scalar.
 *
 * `shiftNodeX` began as a duck-typed loop over four scalar names inside the
 * demo gallery — a shape that can never fail to COMPILE when a node kind gains
 * a coordinate — and `points` was already missing from it, so a polygon would
 * have stayed where it was while everything around it moved.
 *
 * Latent rather than live: `combineRow` composes only Harvey balls and
 * checkboxes, and `elements.ts` emits no polygon. It was also unprovable where
 * it lived, because `src/demo/demo.ts` touches the DOM at import time — which
 * is why it now sits in `scene.ts`, next to the node contract it has to keep up
 * with, and can be checked here.
 */
describe("shifting a scene node horizontally", () => {
  it("moves a polygon's points, not just the scalar coordinates", () => {
    const poly = {
      kind: "polygon",
      name: "p",
      points: [
        { x: 1, y: 5 },
        { x: 3, y: 7 },
      ],
      fill: "#000",
    } as unknown as SceneNode;
    shiftNodeX(poly, 10);
    const pts = (poly as unknown as { points: { x: number; y: number }[] }).points;
    expect(
      pts.map((p) => p.x),
      "a polygon did not move with the scene it belongs to",
    ).toEqual([11, 13]);
    expect(
      pts.map((p) => p.y),
      "a horizontal shift moved a polygon vertically",
    ).toEqual([5, 7]);
  });

  it("moves every scalar horizontal coordinate a node can carry", () => {
    // The other half, or the rule above could be satisfied by handling polygons
    // alone. One node per horizontal field named in `scene.ts`.
    const rect = { kind: "rect", x: 1, y: 2, w: 3, h: 4 } as unknown as SceneNode;
    const line = { kind: "line", x1: 1, y1: 2, x2: 5, y2: 6 } as unknown as SceneNode;
    const ell = { kind: "ellipse", cx: 4, cy: 2, rx: 1, ry: 1 } as unknown as SceneNode;
    for (const n of [rect, line, ell]) shiftNodeX(n, 10);
    expect((rect as unknown as { x: number }).x).toBe(11);
    expect([(line as unknown as { x1: number }).x1, (line as unknown as { x2: number }).x2]).toEqual([11, 15]);
    expect((ell as unknown as { cx: number }).cx).toBe(14);
    // …and leaves the vertical ones alone.
    expect((rect as unknown as { y: number }).y).toBe(2);
  });

  /**
   * The seam, made loud.
   *
   * `shiftNodeX` is a duck-typed loop over a fixed list of key names, so a node
   * kind that arrives with a coordinate the list does not know is left where it
   * was, silently and without a compile error — which is exactly what happened
   * to `points`. CLAUDE.md lists this among the seams that "do not fail loudly"
   * when a `SceneNode` kind is added.
   *
   * `Record<SceneNode["kind"], …>` is what makes it fail loudly: a new kind
   * cannot be added to the union without this map going red, and filling the
   * entry in means naming that kind's horizontal coordinates — at which point
   * whether `shiftNodeX` handles them is a question the author has to answer
   * rather than one they can walk past.
   */
  it("moves every horizontal coordinate of every node kind, and nothing else", () => {
    /** One node of each kind, with the names of the coordinates a shift must move. */
    const EACH: Record<SceneNode["kind"], { node: SceneNode; horizontal: string[] }> = {
      rect: { node: { kind: "rect", x: 1, y: 2, w: 3, h: 4, fill: "#000" }, horizontal: ["x"] },
      line: { node: { kind: "line", x1: 1, y1: 2, x2: 5, y2: 6, stroke: "#000" }, horizontal: ["x1", "x2"] },
      text: {
        node: {
          kind: "text",
          x: 1,
          y: 2,
          w: 3,
          h: 4,
          text: "t",
          fontSize: 10,
          color: "#000",
          align: "left",
          valign: "top",
        },
        horizontal: ["x"],
      },
      ellipse: { node: { kind: "ellipse", cx: 1, cy: 2, rx: 3, ry: 4, fill: "#000" }, horizontal: ["cx"] },
      wedge: {
        node: { kind: "wedge", cx: 1, cy: 2, r: 3, innerR: 1, startAngle: 0, endAngle: 90, fill: "#000" },
        horizontal: ["cx"],
      },
      chevron: { node: { kind: "chevron", x: 1, y: 2, w: 3, h: 4, fill: "#000" }, horizontal: ["x"] },
      arrowhead: { node: { kind: "arrowhead", x: 1, y: 2, angle: 90, size: 3, fill: "#000" }, horizontal: ["x"] },
      symbol: {
        node: { kind: "symbol", shape: "diamond", cx: 1, cy: 2, size: 3, fill: "#000" },
        horizontal: ["cx"],
      },
      polygon: {
        node: {
          kind: "polygon",
          points: [
            { x: 1, y: 5 },
            { x: 3, y: 7 },
          ],
          fill: "#000",
        },
        horizontal: ["points"],
      },
    };

    const DX = 10;
    for (const [kind, { node, horizontal }] of Object.entries(EACH)) {
      const before = structuredClone(node) as unknown as Record<string, unknown>;
      shiftNodeX(node, DX);
      const after = node as unknown as Record<string, unknown>;
      for (const [key, was] of Object.entries(before)) {
        const now = after[key];
        const shouldMove = horizontal.includes(key);
        if (key === "points") {
          const a = was as { x: number; y: number }[];
          const b = now as { x: number; y: number }[];
          expect(
            b.map((p) => p.x),
            `${kind}.points did not move with the scene`,
          ).toEqual(a.map((p) => p.x + DX));
          expect(
            b.map((p) => p.y),
            `${kind}.points moved vertically`,
          ).toEqual(a.map((p) => p.y));
          continue;
        }
        if (typeof was !== "number") continue;
        expect(now, `${kind}.${key} ${shouldMove ? "did not move" : "moved and should not have"}`).toBe(
          shouldMove ? was + DX : was,
        );
      }
    }
  });
});
