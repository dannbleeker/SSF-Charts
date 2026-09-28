import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { authoredDecimals } from "../src/core/format";
import type { TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * Four ways a STATED FIGURE came out at a precision that did not belong to it.
 *
 * This repo's worst defect is a wrong number in ink, and every case below is
 * one: the geometry was right and the caption beside it was not. They share a
 * shape — a number was formatted against the wrong set of values — so they
 * share a file.
 */

const text = (cfg: ChartConfig, prefix: string): string[] =>
  (buildChart(cfg).nodes as TextNode[])
    .filter((n) => n.kind === "text" && n.name?.startsWith(prefix))
    .map((n) => n.text);

describe("a line on its own axis is labelled in its own precision", () => {
  /**
   * `combo.lineAxes: "independent"` zooms every line to its own range and draws
   * NO numeric axis, so the point labels are the only numbers that series has.
   * They were formatted against a precision resolved over EVERY line at once,
   * which comes from the largest — so the small series lost every digit it had.
   */
  const two = (lineAxes?: "independent"): ChartConfig => ({
    kind: "combo",
    width: 640,
    height: 400,
    ...(lineAxes ? { combo: { lineAxes } } : { secondaryAxis: true }),
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "Revenue", type: "line", values: [1200, 1300, 1400, 1500] },
        { name: "Margin", type: "line", values: [0.31, 0.34, 0.38, 0.45] },
      ],
    },
  });

  it("keeps a sub-1 series' digits beside a thousands-scale one", () => {
    // Every one of these four read "0" — the exact pairing independent axes exist for.
    expect(text(two("independent"), "combo-label-1-")).toEqual(["0.31", "0.34", "0.38", "0.45"]);
    // …and the large series is untouched.
    expect(text(two("independent"), "combo-label-0-").slice(0, 3)).toEqual(["1,200", "1,300", "1,400"]);
  });

  it("leaves a SHARED axis on one precision, which is what a shared tick strip means", () => {
    // Not a bug: these points are read against one labelled strip, so they must
    // round together. The fix must not reach this branch.
    expect(new Set(text(two(), "combo-label-1-"))).toEqual(new Set(["0"]));
  });
});

describe("a derived label rounds at the finer of two precisions", () => {
  const withMean = (values: number[], numberFormat?: ChartConfig["numberFormat"]): ChartConfig => ({
    kind: "clustered",
    width: 640,
    height: 420,
    ...(numberFormat ? { numberFormat } : {}),
    data: { categories: values.map((_, i) => `c${i}`), series: [{ name: "v", values }] },
    decorations: { segmentLabels: true, valueLines: [{ mode: "mean" }] },
  });

  it("does not coarsen a mean below what the chart's own labels show", () => {
    // `formatNumber`'s auto ladder stops at 2 decimals under 1 where
    // `resolveFormat` widens to 6, so this read "Ø 0.01" — twice the mean, and
    // above every bar on the chart.
    const cfg = withMean([0.004, 0.005, 0.006]);
    expect(text(cfg, "label-0-")).toEqual(["0.0040", "0.0050", "0.0060"]);
    expect(text(cfg, "value-line-label-")).toEqual(["Ø 0.0050"]);
  });

  it("does not coarsen a fractional mean to its whole-numbered columns", () => {
    // The mirror error: resolving ONLY against the columns turned the true mean
    // of -2.5 into "Ø -3", a number nothing on the chart says.
    expect(text(withMean([-40, -20, -10, 60]), "value-line-label-")).toEqual(["Ø -2.5"]);
  });

  it("still lets an authored decimals count win outright", () => {
    expect(text(withMean([-40, -20, -10, 60], { decimals: 3 }), "value-line-label-")).toEqual(["Ø -2.500"]);
    expect(text(withMean([0.004, 0.005, 0.006], { decimals: 0 }), "value-line-label-")).toEqual(["Ø 0"]);
  });

  it("applies the same rule to the difference arrow's absolute fallback", () => {
    // A negative base has no percentage (see decor.ts), so this falls back to
    // the absolute difference — which was formatted at 2 decimals as "+0.01"
    // beside bars labelled to 4.
    const cfg: ChartConfig = {
      kind: "clustered",
      width: 640,
      height: 420,
      data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [-0.004, 0.005, 0.006] }] },
      decorations: { segmentLabels: true, difference: { from: 0, to: 2 } },
    };
    expect(text(cfg, "diff-label")).toEqual(["+0.0100"]);
  });
});

describe("a suffix belongs to the axis that declared it", () => {
  const bubble = (numberFormat: ChartConfig["numberFormat"]): ChartConfig => ({
    kind: "bubble",
    width: 720,
    height: 460,
    numberFormat,
    data: {
      categories: ["Alpha", "Beta", "Gamma", "Delta"],
      series: [
        { name: "X", values: [10, 20, 30, 40] },
        { name: "Y", values: [12, 26, 31, 44] },
        { name: "Size", values: [120, 340, 90, 500] },
        { name: "Color", values: [1, 2, 3, 4] },
      ],
    },
  });

  it("does not stamp the value axis's unit on the Size and Color keys", () => {
    // These read "600 €m" / "300 €m" and "1.0 €m" … "4.0 €m": a unit belonging
    // to Y, printed over two other variables. Same class as the 100% axis that
    // once read "25 m%".
    const cfg = bubble({ decimals: "auto", suffix: " €m" });
    expect(text(cfg, "size-legend-label-")).toEqual(["600", "300"]);
    expect(text(cfg, "color-legend-min")).toEqual(["1.0"]);
    expect(text(cfg, "color-legend-max")).toEqual(["4.0"]);
    // …and the axes keep it, because it is theirs.
    expect(text(cfg, "y-axis").every((t) => t.endsWith(" €m"))).toBe(true);
  });

  it("does not sign a size key either", () => {
    expect(text(bubble({ decimals: "auto", forceSign: true }), "size-legend-label-")).toEqual(["600", "300"]);
  });

  it("still carries the metric-neutral halves", () => {
    expect(text(bubble({ decimals: 2, locale: "de-DE" }), "size-legend-label-")).toEqual(["600,00", "300,00"]);
  });
});

describe("an authored decimals count reaches the percent labels too", () => {
  const hundred = (numberFormat?: ChartConfig["numberFormat"]): ChartConfig => ({
    kind: "stacked100",
    width: 720,
    height: 460,
    ...(numberFormat ? { numberFormat } : {}),
    data: {
      categories: ["Q1"],
      series: [
        { name: "A", values: [1] },
        { name: "B", values: [1] },
        { name: "C", values: [1] },
      ],
    },
    decorations: { segmentLabels: true, valueAxis: true },
  });

  it("stops a 100% chart contradicting its own axis", () => {
    // `resolveAxisFormat` honours an explicit count and `segmentLabel` hardcoded
    // zero, so the axis read 0.0% … 100.0% over three segments each "33%":
    // 99% of a column drawn at full height.
    const cfg = hundred({ decimals: 1 });
    expect(text(cfg, "value-axis")).toContain("100.0%");
    expect(text(cfg, "label-")).toEqual(["33.3%", "33.3%", "33.3%"]);
  });

  it("leaves an auto chart at whole percents", () => {
    // The count must come from what the AUTHOR wrote, never from the resolved
    // format — whose decimals are always a number, derived from the values'
    // magnitude, and have nothing to do with a share.
    expect(text(hundred(), "label-")).toEqual(["33%", "33%", "33%"]);
    expect(authoredDecimals(undefined)).toBeUndefined();
    expect(authoredDecimals({ decimals: "auto" })).toBeUndefined();
    expect(authoredDecimals({ decimals: 2 })).toBe(2);
    // Out-of-range counts are repaired here as everywhere else.
    expect(authoredDecimals({ decimals: -1 })).toBe(0);
  });

  it("carries to a pie's share labels as well", () => {
    const pie: ChartConfig = {
      kind: "pie",
      width: 720,
      height: 460,
      numberFormat: { decimals: 1 },
      data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [1, 1, 1] }] },
    };
    expect(text(pie, "label-")).toEqual(["A 33.3%", "B 33.3%", "C 33.3%"]);
    expect(text({ ...pie, numberFormat: undefined }, "label-")).toEqual(["A 33%", "B 33%", "C 33%"]);
  });
});
