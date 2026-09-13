import { describe, expect, it } from "vitest";
import { maxOf, minOf } from "../src/core/agg";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * `Math.max(seed, ...arr)` throws RangeError once the array passes V8's argument
 * limit (~10⁵). The extent/format passes flatten the whole data grid (up to the
 * ~10⁶-cell cap) before taking a min/max, so a valid large chart used to crash.
 * maxOf/minOf fold instead of spreading — same result, no overflow.
 */
describe("maxOf / minOf", () => {
  it("match Math.max/Math.min including the seed", () => {
    expect(maxOf([3, 1, 4, 1, 5], 0)).toBe(Math.max(0, 3, 1, 4, 1, 5));
    expect(minOf([3, 1, 4, 1, 5], 0)).toBe(Math.min(0, 3, 1, 4, 1, 5));
    expect(maxOf([-3, -1, -4])).toBe(Math.max(-3, -1, -4)); // seed -Infinity
    expect(minOf([3, 1, 4])).toBe(Math.min(3, 1, 4)); // seed +Infinity
  });

  it("propagate NaN exactly like the spread", () => {
    expect(Number.isNaN(maxOf([1, NaN, 2], 0))).toBe(true);
    expect(Number.isNaN(minOf([1, NaN, 2], 0))).toBe(true);
  });

  it("return the seed for an empty array", () => {
    expect(maxOf([], 0)).toBe(0);
    expect(minOf([], 0)).toBe(0);
    expect(maxOf([])).toBe(-Infinity);
    expect(minOf([])).toBe(Infinity);
  });

  it("do not overflow on an array far past the argument limit", () => {
    const big = Array.from({ length: 200_000 }, (_, i) => i - 100_000);
    expect(() => maxOf(big, 0)).not.toThrow();
    expect(maxOf(big, 0)).toBe(99_999);
    expect(minOf(big, 0)).toBe(-100_000);
  });
});

describe("buildChart survives a large within-cap grid (no RangeError)", () => {
  it("renders a chart whose flattened values exceed the spread limit", () => {
    // ~180k values — over V8's arg limit, so the old `Math.max(0, ...all)` in the
    // extent pass threw `Maximum call stack size exceeded` on this valid input.
    const nCats = 720;
    const nSeries = 256;
    const cfg: ChartConfig = {
      kind: "clustered",
      width: 960,
      height: 600,
      data: {
        categories: Array.from({ length: nCats }, (_, i) => `C${i}`),
        series: Array.from({ length: nSeries }, (_, s) => ({
          name: `S${s}`,
          values: Array.from({ length: nCats }, (_, c) => (s * c) % 100),
        })),
      },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
  it("clustered-stacked: a large categories x stacks grid does not blow the argument list", () => {
    // The stacked && nStacks > 1 branch flattens categories x DISTINCT STACKS,
    // which #150 left as a spread while converting its siblings — so a valid
    // in-cap grid still threw RangeError from layoutColumns.
    const nCats = 1024;
    const nSeries = 128;
    const cfg: ChartConfig = {
      kind: "stacked",
      width: 960,
      height: 600,
      data: {
        categories: Array.from({ length: nCats }, (_, i) => `C${i}`),
        series: Array.from({ length: nSeries }, (_, s) => ({
          name: `S${s}`,
          stack: s, // distinct stack per series => nStacks > 1
          values: Array.from({ length: nCats }, (_, c) => (s * c) % 100),
        })),
      },
    };
    expect(() => buildChart(cfg)).not.toThrow();
    // 0.7 s idle, measured 2026-09-13 — a 1024 x 128 grid is large on purpose,
    // that being the whole point. It has no timeout of its own: it is covered
    // by the 20s default in `vitest.config.ts`, which that file explains.
  });

  it("radar: a large cells-scaled grid does not blow the argument list", () => {
    // radar's tickMax spread was skipped by #150 entirely.
    const nCats = 512;
    const nSeries = 256;
    const cfg: ChartConfig = {
      kind: "radar",
      width: 600,
      height: 600,
      data: {
        categories: Array.from({ length: nCats }, (_, i) => `C${i}`),
        series: Array.from({ length: nSeries }, (_, s) => ({
          name: `S${s}`,
          values: Array.from({ length: nCats }, (_, c) => (s + c) % 50),
        })),
      },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
});
