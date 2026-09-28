import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * A MEASURED ZERO IS A POINT. A BLANK IS STILL NOTHING.
 *
 * `blank-is-not-zero.test.ts` guards one direction — a cell nobody filled in
 * must not be printed as a measured figure. This guards the other, which the
 * same code got wrong in the opposite way.
 *
 * Found 2026-09-28. `column.ts` gated every mark on `raw != null && v !== 0`.
 * The `v !== 0` half is right for a RECTANGLE: a bar of no height is invisible
 * whether or not it is emitted. It is wrong for a DOT or a LOLLIPOP, where the
 * mark is a point AT the value and zero is an ordinary place on the axis.
 *
 * Measured, `[10, 0, 30]` against `[10, null, 30]` at `barStyle: "dot"` and
 * `"lollipop"`: both produced the identical scene for category 1 — the axis
 * heading and nothing else. A zero somebody measured and a cell nobody filled
 * in were indistinguishable on the chart. The harm in the sibling file is
 * inventing data; the harm here is silently discarding it.
 */

const cfg = (barStyle: string, values: (number | null)[], extra: unknown[] = []): ChartConfig =>
  ({
    kind: "clustered",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values }, ...extra] },
    decorations: { barStyle, segmentLabels: true },
  }) as unknown as ChartConfig;

/** Every node this chart draws that speaks about category 1. */
const at1 = (c: ChartConfig) =>
  buildChart(c)
    .nodes.map((n) => String((n as { name?: string }).name ?? ""))
    .filter((n) => /-1$|-0-1$/.test(n))
    // The axis heading is drawn whatever the data says and is not a mark.
    .filter((n) => n !== "category-1");

describe("an exact zero on a point-marked chart", () => {
  it.each(["dot", "lollipop"])("draws a mark for a measured zero on %s", (barStyle) => {
    const drawn = at1(cfg(barStyle, [10, 0, 30]));
    expect(
      drawn.length,
      `a measured 0 drew nothing at all on a ${barStyle} chart — it is a point on the axis, not an ` +
        `absent value, and a reader cannot tell it from a blank cell`,
    ).toBeGreaterThan(0);
    expect(
      drawn.some((n) => n.startsWith("seg-")),
      `no point mark for the zero on ${barStyle}`,
    ).toBe(true);
  });

  it.each(["dot", "lollipop"])("still draws nothing for a blank on %s", (barStyle) => {
    expect(
      at1(cfg(barStyle, [10, null, 30])),
      "a cell nobody filled in was given a mark — that is the sibling defect in blank-is-not-zero.test.ts",
    ).toEqual([]);
  });

  it("tells the two apart, which is the whole point", () => {
    for (const barStyle of ["dot", "lollipop"]) {
      expect(
        at1(cfg(barStyle, [10, 0, 30])),
        `a measured zero and a blank produce the identical scene on ${barStyle}`,
      ).not.toEqual(at1(cfg(barStyle, [10, null, 30])));
    }
  });

  /**
   * The scope guard. A zero-height RECTANGLE is invisible either way, and
   * `range` pairs two series — neither was part of this defect, and a fix that
   * changed them would be a behaviour change nobody asked for.
   */
  it("leaves bar and range exactly as they were", () => {
    expect(at1(cfg("bar", [10, 0, 30])), "a bar chart started emitting zero-height rects").toEqual([]);
    const range = (values: (number | null)[]) => at1(cfg("range", values, [{ name: "T", values: [20, 5, 40] }]));
    expect(range([10, 0, 30])).toContain("range-1");
    expect(range([10, null, 30])).not.toContain("range-1");
  });
});
