import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * A DIFFERENCE ARROW ANCHORED TO A MEAN VALUE LINE MUST START ON THAT LINE.
 *
 * Found 2026-09-28. `decor.ts` computed the mean twice. The value line filtered
 * blank categories out — every layout builds `columnValue` with `?? 0`, so a
 * blank is both a spurious zero in the sum and a spurious entry in the
 * denominator — and the `fromValueLine` anchor kept an unfiltered copy.
 *
 * On `[100, null, 200]` the chart therefore drew TWO means: the value line at
 * 150 labelled "Ø 150", the arrow's baseline 65.75pt away at 100, and the arrow
 * captioned "+100%" where the truth is "+33%".
 *
 * WHAT THIS PINS IS THE AGREEMENT, NOT THE NUMBERS. Asserting "y1 === 84.75"
 * would pass just as well if both computations were wrong in the same way, and
 * would break on any innocent change to frame padding. The invariant is that
 * the arrow starts ON the line it names — so the test reads both out of one
 * scene and compares them, which is false for exactly one reason.
 *
 * The original fix (`35840f6`) closed with "the second time this week the fix
 * was already written and one call site was missed". This is that class again.
 */

const cfg = (values: (number | null)[]): ChartConfig =>
  ({
    kind: "clustered",
    width: 480,
    height: 300,
    data: { categories: ["A", "B", "C"], series: [{ name: "Revenue", values }] },
    decorations: { valueLines: [{ mode: "mean" }], difference: { from: 0, to: 2, fromValueLine: 0 } },
  }) as unknown as ChartConfig;

function flatten(scene: unknown): Record<string, unknown>[] {
  const out: Record<string, unknown>[] = [];
  const seen = new Set<unknown>();
  const walk = (n: unknown) => {
    if (!n || typeof n !== "object" || seen.has(n)) return;
    seen.add(n);
    if (Array.isArray(n)) return n.forEach(walk);
    out.push(n as Record<string, unknown>);
    for (const v of Object.values(n as Record<string, unknown>)) if (v && typeof v === "object") walk(v);
  };
  walk(scene);
  return out;
}

const first = (flat: Record<string, unknown>[], prefix: string) =>
  flat.find((n) => String(n.name ?? "").startsWith(prefix));

describe("a difference arrow anchored to a mean value line", () => {
  it.each([
    ["a blank category", [100, null, 200] as (number | null)[]],
    ["no blanks — the control", [100, 150, 200] as (number | null)[]],
  ])("starts on the line it is anchored to, with %s", (_what, values) => {
    const flat = flatten(buildChart(cfg(values)));
    const line = first(flat, "value-line-0") as { y1?: number } | undefined;
    const level = first(flat, "diff-level") as { y1?: number } | undefined;
    expect(line?.y1, "no mean value line was drawn — re-point this test").toBeTypeOf("number");
    expect(level?.y1, "no difference-arrow baseline was drawn — re-point this test").toBeTypeOf("number");
    expect(
      Math.abs((level!.y1 as number) - (line!.y1 as number)),
      `the arrow's baseline is ${level!.y1} and the mean line it names is at ${line!.y1}. They must be ` +
        `the same line: the arrow claims to measure FROM that mean, and a reader sees both.`,
    ).toBeLessThan(0.5);
  });

  /**
   * And the caption has to match the geometry. A blank category used to move
   * the printed growth from +33% to +100% — the arrow was right about its own
   * (wrong) baseline, which is why only comparing against the value line
   * catches it.
   */
  it("prints the growth from the measured mean, not from a blank-diluted one", () => {
    const withBlank = flatten(buildChart(cfg([100, null, 200])));
    const filled = flatten(buildChart(cfg([100, 150, 200])));
    const label = (flat: Record<string, unknown>[]) =>
      (first(flat, "diff-label") as { text?: string } | undefined)?.text;
    expect(label(withBlank), "no difference label was drawn — re-point this test").toBeTypeOf("string");
    expect(
      label(withBlank),
      `a blank category changed the arrow's caption to ${JSON.stringify(label(withBlank))}. A cell nobody ` +
        `filled in must not change a stated percentage — it is not a zero.`,
    ).toBe(label(filled));
  });
});
