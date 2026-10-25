import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { buildChart } from "../src/core/chart";
import { estimateOfficeShapes } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * HOW MANY SHAPES A CHART IS ALLOWED TO BE, and nothing has ever capped it.
 *
 * Shapes are the unit of cost on PowerPoint on the web, and the archive is
 * unambiguous about the price. An in-place update runs about **767ms per shape**
 * — measured over 178 updates — which makes `updated only the shapes that
 * changed` 11.8 calls and 108 seconds a round, a quarter of a round, with a
 * worst single chart at 26.4s. It is not batchable: within one update the chunks
 * cost 4514, 4366 and 4657ms, flat, so the cost is linear in shapes written and
 * server-side.
 *
 * And past a point it stops being slow and starts being fatal. Round 150
 * recorded **400-500 shapes in one context crashing PowerPoint on all seven
 * attempts**, and before the slab rule took the rise into account a chart in the
 * shipped deck sat at 300.
 *
 * So this is the third ratchet in the house style — `overlap-budget` for the
 * variant sweep, `showcase-overlap` for the shipped deck, and this for density.
 * Same rules: a kind over budget is a regression, an unbudgeted kind is a
 * regression by definition, and a budget left ABOVE the real figure fails, so
 * making a chart lighter forces the number down instead of banking slack.
 *
 * WHAT THIS IS NOT. It is not a claim that any of these figures is too high; it
 * is a claim that they should not RISE without somebody saying so.
 *
 * ── IT COUNTED THE WRONG THING FOR ITS WHOLE LIFE ───────────────────────────
 * Every number above is about SHAPES, and until 2026-08-30 this file measured
 * `scene.nodes.length` — the graph the layouts emit, not what the renderer
 * writes. They are not the same count and `estimateOfficeShapes` exists to say
 * so: a polygon is one node and one line per EDGE, a wedge one node and a whole
 * fan. So the gate was blind in exactly the places density comes from.
 *
 * What it was hiding, measured the day it was fixed:
 *   - The hex tile map: 146 nodes, **401 shapes**. Sitting inside the 400-500
 *     window that took PowerPoint down on all seven attempts in round 150,
 *     while `keeps every shipped chart clear of the count that crashed the
 *     host` passed on 146. That test could not have failed: the number it
 *     compared against 300 was not the number in its own docstring.
 *   - The violin: 16 nodes, **259 shapes**. Budgeted as the second-LIGHTEST
 *     chart we ship when it is now the heaviest — a 16x under-count.
 *   - Sunburst 21 -> 101, pie 17 -> 56, doughnut 15 -> 55, radar 36 -> 79.
 *   - The deck total: 5560 nodes, 7128 shapes.
 *
 * The tile map is 146 shapes now because its tiles became fillable symbols, and
 * the deck 6873. The violin is the one that would repay attention next.
 *
 * A gate is only as true as the quantity it reads, and this one named its
 * quantity correctly in prose while measuring its neighbour in code.
 * ────────────────────────────────────────────────────────────────────────────
 */

/**
 * Worst chart of each kind in `examples/showcase.json`, in OFFICE SHAPES, as of
 * 2026-08-30. Nine of these rose when the metric was corrected; none rose
 * because a chart changed.
 */
const BUDGET: Record<string, number> = {
  // 97, from 259, when the KDE outlines were thinned to a quarter-point
  // tolerance: 246 polygon edges became 72, and the violin stopped being the
  // heaviest chart in the deck. It opened this file as the one that would
  // repay attention; this is what attending to it was worth.
  violin: 97,
  line: 226,
  // 165, from 164, on 2026-10-25: "Revenue mix (stacked area) + margin %" now
  // names its margin line. An upright combo with an area base never drew the
  // line's name at all (a bug-hunt finding), so this is the one text shape
  // that was missing. Five other showcase combos and the Pareto rose by the
  // same one shape for the same reason; none is near the 400-500 that crashed
  // the host.
  combo: 165,
  tilemap: 146,
  area: 123,
  waffle: 107,
  sunburst: 101,
  heatmap: 100,
  radar: 79,
  scatter: 75,
  gantt: 68,
  boxplot: 62,
  pie: 56,
  // 54, from 55, on 2026-10-25: a doughnut's slices are real ring segments now
  // instead of pie wedges with an opaque disc laid over the middle, so the
  // disc ("hole") is gone. The only budget in this table that went DOWN, and it
  // had to: a budget above the real figure fails.
  doughnut: 54,
  butterfly: 53,
  stacked: 47,
  bubble: 46,
  // 41, from 40, on 2026-10-25: "Pareto chart (defect causes)" now names its
  // cumulative-% line, the one shape its combo rewrite never drew.
  clustered: 41,
  candlestick: 35,
  stacked100: 33,
  mekko: 32,
  waterfall: 31,
  treemap: 31,
  cascade: 30,
  funnel: 20,
};

/**
 * The whole deck, so a rise spread thinly across many charts still shows.
 *
 * 6385, from 6375, on 2026-10-25, and the +10 was attributed chart by chart
 * against the engine at 949e27c rather than absorbed: six charts +1 each (the
 * series name that was never drawn), the owner-lane Gantt +5 (unowned tasks
 * now sit under a header of their own instead of inside the last owner's
 * lane), and the doughnut -1 (no covering disc).
 */
const TOTAL_BUDGET = 6385;

function measure(): { byKind: Map<string, number>; total: number; worst: Map<string, string> } {
  const items = JSON.parse(readFileSync("examples/showcase.json", "utf8")) as ChartConfig[];
  const byKind = new Map<string, number>();
  const worst = new Map<string, string>();
  let total = 0;
  for (const cfg of items) {
    let n: number;
    try {
      // What the RENDERER writes, not what the layout emitted. `nodes.length`
      // is the tempting one and it is the wrong one — see the header.
      n = estimateOfficeShapes(buildChart(cfg));
    } catch {
      // A config the engine refuses is not a density problem, and
      // `showcase.test.ts` is what holds the deck to building at all.
      continue;
    }
    total += n;
    const kind = String(cfg.kind ?? "?");
    if (n > (byKind.get(kind) ?? 0)) {
      byKind.set(kind, n);
      worst.set(kind, String(cfg.title ?? kind));
    }
  }
  return { byKind, total, worst };
}

describe("no chart quietly becomes too heavy to draw", () => {
  const { byKind, total, worst } = measure();

  it("draws no kind it has never drawn before", () => {
    const unbudgeted = [...byKind.keys()].filter((k) => BUDGET[k] === undefined);
    expect(unbudgeted, "a new chart kind reached the shipped deck — budget it here and say what it is").toEqual([]);
  });

  it("keeps every kind inside its budget", () => {
    const over = [...byKind.entries()]
      .filter(([k, n]) => n > (BUDGET[k] ?? 0))
      .map(([k, n]) => `${k}: ${n} shapes (budget ${BUDGET[k] ?? 0}) — ${worst.get(k)}`);
    expect(over, "a chart kind got denser, and shapes are what this host charges for").toEqual([]);
  });

  it("holds no budget ABOVE the real figure", () => {
    // The half that makes it a ratchet rather than a ceiling. Making a chart
    // lighter is supposed to fail this file so the number gets edited down —
    // which is exactly what the slab rule did to `area`, 300 -> 123.
    const slack = Object.entries(BUDGET)
      .filter(([k, n]) => (byKind.get(k) ?? 0) < n)
      .map(([k, n]) => `${k}: budget ${n}, actual ${byKind.get(k) ?? 0}`);
    expect(slack, "a budget here is above the real figure — edit it down").toEqual([]);
  });

  it("has not grown across the deck as a whole", () => {
    // Per-kind ceilings cannot see a rise spread thinly over many charts, which
    // is the shape a data change makes.
    expect(total, `${total} shapes across the shipped deck, budget ${TOTAL_BUDGET}`).toBeLessThanOrEqual(TOTAL_BUDGET);
  });

  it("keeps every shipped chart clear of the count that crashed the host", () => {
    /**
     * The hard one, and the reason this file exists rather than a comment.
     * Round 150: 400-500 shapes in ONE context, PowerPoint down on all seven
     * attempts. A single chart is not a whole context — a slide can hold
     * several — so the per-chart line sits well below it.
     */
    const heavy = [...byKind.entries()].filter(([, n]) => n >= 300).map(([k, n]) => `${k}: ${n} — ${worst.get(k)}`);
    expect(heavy, "a shipped chart is within range of the count that crashed PowerPoint").toEqual([]);
  });
});
