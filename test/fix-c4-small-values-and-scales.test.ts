import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { textWidth } from "../src/core/scene";
import { sceneToSvg } from "../src/render/svg";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, LineNode, PolygonNode, RectNode, SceneNode, TextNode, WedgeNode } from "../src/core/scene";

/**
 * Layouts that flatten or clamp small or negative values.
 *
 * One family: a layout that decides its scale from the data, and quietly lets
 * something that is NOT the data decide it instead — a floor of 1 in data units,
 * a lower bound of 0 whatever the data minimum is, a fallback of 1 where the
 * unit should have set it. The chart then depends on the unit the numbers were
 * typed in: 0.08 and 8 are the same chart, so they must draw the same picture.
 * Every test below states that as a comparison of the SAME data at two units,
 * which is the invariant, rather than as a pinned number that would only be the
 * current output.
 */

/** The same data at another unit. Nulls stay blank. */
const scaled = (cfg: ChartConfig, k: number): ChartConfig => ({
  ...cfg,
  data: {
    ...cfg.data,
    series: cfg.data.series.map((s) => ({ ...s, values: s.values.map((v) => (v == null ? v : v * k)) })),
  },
});

const named = <T extends SceneNode>(nodes: SceneNode[], kind: T["kind"], prefix: string): T[] =>
  nodes.filter((n): n is T => n.kind === kind && !!n.name?.startsWith(prefix));

const rects = (cfg: ChartConfig, prefix: string) => named<RectNode>(buildChart(cfg).nodes, "rect", prefix);

const texts = (cfg: ChartConfig, prefix: string) =>
  named<TextNode>(buildChart(cfg).nodes, "text", prefix).map((n) => n.text);

/** Two lists of numbers agree to a micro-point. */
const sameNumbers = (a: number[], b: number[]) => {
  expect(a.length).toBe(b.length);
  a.forEach((v, i) => expect(v).toBeCloseTo(b[i], 6));
};

/** How far each radar vertex marker sits from the web's centre, spoke by spoke. */
const markerDistances = (cfg: ChartConfig, series = 0) => {
  const nodes = buildChart(cfg).nodes;
  const spoke = named<LineNode>(nodes, "line", "spoke-0")[0];
  return named<EllipseNode>(nodes, "ellipse", `marker-${series}-`)
    .sort((a, b) => Number(a.name!.split("-")[2]) - Number(b.name!.split("-")[2]))
    .map((m) => Math.hypot(m.cx - spoke.x1, m.cy - spoke.y1));
};

describe("a floor of 1 in data units does not decide the scale (r1-9)", () => {
  const butterfly: ChartConfig = {
    kind: "butterfly",
    ...DEFAULT_SIZE,
    decorations: { valueAxis: true },
    data: {
      categories: ["0-17", "18-34", "35-54", "55-74", "75+"],
      series: [
        { name: "Male", values: [0.08, 0.1, 0.11, 0.1, 0.09] },
        { name: "Female", values: [0.08, 0.09, 0.11, 0.11, 0.12] },
      ],
    },
  };

  it("butterfly bars are as long for 0.08 as for 8", () => {
    // 22.4pt against 149.2pt before: `niceTicks(0, Math.max(1, ...sums))` forced the
    // axis to at least 0..1, so the longest bar of a sub-1 chart took 12% of its
    // half. The floor guarded nothing — niceTicks already answers [0, 1] for 0..0.
    const small = rects(butterfly, "seg-").map((r) => r.w);
    const big = rects(scaled(butterfly, 100), "seg-").map((r) => r.w);
    sameNumbers(small, big);
  });

  it("butterfly ticks follow the data: a 0..0.12 chart is ticked in hundredths", () => {
    const ticks = texts(butterfly, "tick-").map(Number);
    expect(Math.max(...ticks)).toBeLessThan(0.5);
    expect(Math.max(...ticks)).toBeGreaterThanOrEqual(0.12);
  });

  it("butterfly with nothing to measure still draws a finite zero-length chart", () => {
    // The guard the floor was thought to be. All-zero data must not divide by zero.
    const zero = scaled(butterfly, 0);
    const lens = rects(zero, "seg-").map((r) => r.w);
    expect(lens.length).toBeGreaterThan(0);
    for (const w of lens) expect(w).toBe(0);
    for (const n of buildChart(zero).nodes)
      for (const v of Object.values(n)) if (typeof v === "number") expect(Number.isFinite(v)).toBe(true);
  });

  const radar: ChartConfig = {
    kind: "radar",
    ...DEFAULT_SIZE,
    scale: undefined,
    data: {
      categories: ["Speed", "Power", "Range", "Cost", "Style"],
      series: [{ name: "Unit", values: [0.2, 0.4, 0.3, 0.25, 0.35] }],
    },
  };

  const polyPoints = (cfg: ChartConfig, name: string) =>
    named<PolygonNode>(buildChart(cfg).nodes, "polygon", name).flatMap((p) => p.points.flatMap((q) => [q.x, q.y]));

  it("radar draws 0.2-0.4 as the same web as 20-40", () => {
    // Rings 0.5 / 1.0 before: the same `1` floor on the auto domain left the
    // largest value at 40% of the way out.
    sameNumbers(polyPoints(radar, "series-0"), polyPoints(scaled(radar, 100), "series-0"));
  });

  it("radar ring labels come from the data, not from a unit", () => {
    const ticks = texts(radar, "tick-").map(Number);
    expect(Math.max(...ticks)).toBeLessThan(0.5);
  });

  it("radial bars are as long for 0.2-0.4 as for 20-40", () => {
    const bars: ChartConfig = { ...radar, radar: { bars: true } };
    const radii = (cfg: ChartConfig) => named<WedgeNode>(buildChart(cfg).nodes, "wedge", "bar-").map((w) => w.r);
    sameNumbers(radii(bars), radii(scaled(bars, 100)));
  });

  it("radar with nothing to measure still draws a finite web", () => {
    for (const bars of [false, true]) {
      const scene = buildChart({ ...scaled(radar, 0), radar: { bars } });
      for (const n of scene.nodes)
        for (const v of Object.values(n)) if (typeof v === "number") expect(Number.isFinite(v)).toBe(true);
      expect(named<TextNode>(scene.nodes, "text", "category-").length).toBeGreaterThan(0);
    }
  });

  const cascade: ChartConfig = {
    kind: "cascade",
    ...DEFAULT_SIZE,
    data: { categories: ["Total", "Answered", "Solved"], series: [{ name: "Calls", values: [0.09, 0.06, 0.03] }] },
  };

  it("cascade's first stage fills the plot whatever the unit", () => {
    // 25.6pt of a 284pt plot for [0.09, 0.06, 0.03]: the divisor was
    // `Math.max(...values, 1)`, so the first stage — which is the chart's whole
    // reference — drew at 9% of its height.
    const heights = (cfg: ChartConfig) => rects(cfg, "stage-").map((r) => r.h);
    sameNumbers(heights(cascade), heights(scaled(cascade, 100)));
  });

  it("cascade with nothing to measure still draws finite stages", () => {
    const zero = scaled(cascade, 0);
    for (const n of buildChart(zero).nodes)
      for (const v of Object.values(n)) if (typeof v === "number") expect(Number.isFinite(v)).toBe(true);
  });
});

describe("a radar's auto scale reaches below zero when the data does (r1-11)", () => {
  const radar = (values: number[], scale?: ChartConfig["scale"]): ChartConfig => ({
    kind: "radar",
    ...DEFAULT_SIZE,
    scale,
    data: { categories: ["A", "B", "C", "D", "E"], series: [{ name: "Growth", values }] },
  });

  it("draws -20, -5 and 0 at three different radii", () => {
    // 0.00 / 0.00 / 0.00 before: the auto lower bound was 0 whatever the data held,
    // and `toR` clamps everything below it onto the centre, so three different
    // values were one point and no ring showed a negative one.
    const d = markerDistances(radar([-20, -5, 0, 10, 30]));
    expect(d[1]).toBeGreaterThan(d[0] + 1);
    expect(d[2]).toBeGreaterThan(d[1] + 1);
    expect(d[3]).toBeGreaterThan(d[2] + 1);
    expect(d[4]).toBeGreaterThan(d[3] + 1);
  });

  it("puts a ring on zero when the web reaches below it", () => {
    const labels = texts(radar([-20, -5, 0, 10, 30]), "tick-");
    expect(labels).toContain("0");
  });

  it("draws an all-negative series as a web, not as a point under invented rings", () => {
    // Every vertex on the centre, with rings labelled 0.5 and 1.0 — the `1` floor
    // on the upper bound, nothing to do with the data.
    const cfg = radar([-20, -5, -1, -10, -30]);
    const d = markerDistances(cfg);
    expect(new Set(d.map((x) => x.toFixed(2))).size).toBeGreaterThan(3);
    const ticks = texts(cfg, "tick-").map(Number);
    expect(ticks).not.toContain(0.5);
    expect(Math.max(...ticks)).toBeLessThanOrEqual(0);
  });

  it("still clamps to a scale the author pinned", () => {
    // The clamp is deliberate when the author fixed the domain (the sample ships
    // {min:0,max:5}); only the AUTO domain was wrong to leave the data out.
    const d = markerDistances(radar([-3, 1, 2, 3, 4], { min: 0, max: 5 }));
    expect(d[0]).toBe(0);
    expect(d[1]).toBeGreaterThan(0);
  });

  it("leaves a non-negative web where it was: the centre is zero", () => {
    const d = markerDistances(radar([0, 10, 20, 30, 40]));
    expect(d[0]).toBe(0);
  });
});

describe("a funnel is scaled by its own widest stage, not by 1 (r1-15)", () => {
  const funnel = (values: number[]): ChartConfig => ({
    kind: "funnel",
    ...DEFAULT_SIZE,
    data: { categories: ["Visits", "Leads", "Won"].slice(0, values.length), series: [{ name: "Stage", values }] },
  });
  const widths = (values: number[]) => rects(funnel(values), "stage-").map((r) => r.w);

  it("draws 0.62 / 0.41 / 0.18 as wide as 62 / 41 / 18", () => {
    // 265.5 / 175.6 / 77.1 before, against 428.2 / 283.2 / 124.3: the widest stage
    // took 62% of the plot because the divisor was floored at 1.
    sameNumbers(widths([0.62, 0.41, 0.18]), widths([62, 41, 18]));
  });

  it("lets the widest stage fill the plot at any unit", () => {
    for (const k of [0.001, 0.05, 1, 1000]) {
      const [top] = widths([0.9 * k, 0.5 * k, 0.2 * k]);
      expect(top).toBeCloseTo(widths([90, 50, 20])[0], 6);
    }
  });

  it("keeps the stages in proportion below the 2pt band minimum's old threshold", () => {
    // 2.1 / 2.0 / 2.0 before: v x plot.w fell under the 2pt minimum for v < ~0.005,
    // so three stages of 5, 3 and 1 thousandths were one width.
    const w = widths([0.005, 0.003, 0.001]);
    expect(w[1] / w[0]).toBeCloseTo(0.6, 6);
    expect(w[2] / w[0]).toBeCloseTo(0.2, 6);
  });

  it("draws an all-zero funnel as finite minimum-width bands", () => {
    // The guard the floor was thought to be.
    const w = widths([0, 0, 0]);
    expect(w).toEqual([2, 2, 2]);
    for (const n of buildChart(funnel([0, 0, 0])).nodes)
      for (const v of Object.values(n)) if (typeof v === "number") expect(Number.isFinite(v)).toBe(true);
  });

  describe("and a variable-radius pie, which carried the same floor", () => {
    const pie = (radius: number[]): ChartConfig => ({
      kind: "pie",
      ...DEFAULT_SIZE,
      data: {
        categories: ["North", "South"],
        series: [
          { name: "Share", values: [60, 40] },
          { name: "Radius", values: radius },
        ],
      },
    });
    const radii = (radius: number[]) =>
      named<WedgeNode>(buildChart(pie(radius)).nodes, "wedge", "slice-").map((w) => w.r);

    it("draws a Radius of 0.4 / 0.8 as 40 / 80", () => {
      // 0.70r / 0.90r before, against 0.75r / 1.00r.
      sameNumbers(radii([0.4, 0.8]), radii([40, 80]));
    });

    it("keeps the all-zero Radius row finite", () => {
      for (const r of radii([0, 0])) expect(Number.isFinite(r)).toBe(true);
    });
  });
});

describe("a violin's constant category is a spike at any unit (r2-12)", () => {
  // Every row is an observation; the first category never varies.
  const violin: ChartConfig = {
    kind: "violin",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Fixed", "Var B", "Var C"],
      series: [
        { name: "r1", values: [0.5, 0.41, 0.52] },
        { name: "r2", values: [0.5, 0.43, 0.55] },
        { name: "r3", values: [0.5, 0.45, 0.57] },
        { name: "r4", values: [0.5, 0.47, 0.58] },
        { name: "r5", values: [0.5, 0.44, 0.6] },
      ],
    },
  };
  const bodyOf = (cfg: ChartConfig, name: string) => {
    const scene = buildChart(cfg);
    const body = named<PolygonNode>(scene.nodes, "polygon", name)[0];
    const ys = body.points.map((p) => p.y);
    return { top: Math.min(...ys), bottom: Math.max(...ys) };
  };

  it("does not stretch the constant category across the whole value axis", () => {
    // y 8.0..282.0 of a 300pt chart with 6 points — a full-width rectangle — on a
    // 0.40-0.60 axis. The zero-spread fallback was 1 IN DATA UNITS, so the
    // bandwidth was ~0.8 against an axis 0.2 tall and the density came out flat.
    const { top, bottom } = bodyOf(violin, "violin-0");
    expect(bottom - top).toBeLessThan(DEFAULT_SIZE.height / 10);
  });

  it("draws the same spike for 0.5 as for 50 and for 500", () => {
    const base = bodyOf(violin, "violin-0");
    for (const k of [100, 1000]) {
      const s = bodyOf(scaled(violin, k), "violin-0");
      expect(s.top).toBeCloseTo(base.top, 4);
      expect(s.bottom).toBeCloseTo(base.bottom, 4);
    }
  });

  it("leaves the categories that do vary alone", () => {
    const small = bodyOf(violin, "violin-1");
    const big = bodyOf(scaled(violin, 100), "violin-1");
    expect(small.top).toBeCloseTo(big.top, 4);
    expect(small.bottom).toBeCloseTo(big.bottom, 4);
  });
});

describe("a bubble's negative Size does not draw as a positive one (r1-42)", () => {
  const bubble = (size: number[], kind: "bubble" | "scatter" = "bubble"): ChartConfig => ({
    kind,
    ...DEFAULT_SIZE,
    data: {
      categories: ["Gain", "Loss", "Other"].slice(0, size.length),
      series: [
        { name: "X", values: [10, 20, 30].slice(0, size.length) },
        { name: "Y", values: [10, 20, 15].slice(0, size.length) },
        { name: "Size", values: size },
      ],
    },
  });
  const point = (cfg: ChartConfig, i: number) => named<EllipseNode>(buildChart(cfg).nodes, "ellipse", `point-${i}`)[0];

  it("draws +500 and -500 differently, at the same size", () => {
    // Both were rx/ry 25.02, fill #2a78d6, stroke #ffffff, strokeWidth 1.
    const cfg = bubble([500, -500]);
    const gain = point(cfg, 0);
    const loss = point(cfg, 1);
    expect(loss.rx).toBeCloseTo(gain.rx, 6);
    expect(loss.fill).not.toBe(gain.fill);
    expect(loss.stroke).not.toBe(gain.stroke);
  });

  it("rings the loss in the colour of the bubble it would have been", () => {
    const cfg = bubble([500, -500]);
    expect(point(cfg, 1).stroke).toBe(point(cfg, 0).fill);
    expect(point(cfg, 1).strokeWidth).toBeGreaterThan(point(cfg, 0).strokeWidth!);
  });

  it("marks the largest negative on a mixed-sign chart, where it is the biggest bubble", () => {
    // [100, -900, 400]: the -900 point is the largest bubble, painted at the back,
    // under a key that reads 1,000 / 500 — nothing said it was a loss.
    const cfg = bubble([100, -900, 400]);
    const radii = [0, 1, 2].map((i) => point(cfg, i).rx);
    expect(radii[1]).toBeGreaterThan(radii[2]);
    expect(radii[2]).toBeGreaterThan(radii[0]);
    // No Group row, so every point shares one colour: the ring is that colour.
    expect(point(cfg, 1).stroke).toBe(point(cfg, 0).fill);
    expect(point(cfg, 1).fill).not.toBe(point(cfg, 0).fill);
  });

  it("leaves a positive bubble exactly as it was: white halo, 1pt, solid colour", () => {
    const cfg = bubble([500, 300]);
    for (const i of [0, 1]) {
      const p = point(cfg, i);
      expect(p.stroke).toBe("#ffffff");
      expect(p.strokeWidth).toBe(1);
    }
  });

  it("does not read a Size row on a plain scatter", () => {
    const cfg = bubble([500, -500], "scatter");
    expect(point(cfg, 1).stroke).toBe("#ffffff");
    expect(point(cfg, 1).fill).toBe(point(cfg, 0).fill);
  });
});

describe("a sideways waterfall's outside label gets the room its row has (r1-8)", () => {
  const bridge = (values: (number | null)[], width = 480, scale?: ChartConfig["scale"]): ChartConfig => ({
    kind: "waterfall",
    width,
    height: DEFAULT_SIZE.height,
    horizontal: true,
    scale,
    waterfall: { totalIndices: [4] },
    data: {
      categories: ["Start", "Churn", "Upsell", "Costs", "End"],
      series: [{ name: "EUR", values }],
    },
  });
  const label = (cfg: ChartConfig, c: number) => {
    const t = named<TextNode>(buildChart(cfg).nodes, "text", `label-${c}`)[0];
    return { text: t.text, x: t.x, w: t.w, fs: t.fontSize };
  };

  it("prints +125,000 whole on a 480pt chart", () => {
    // "+125,0…" before: the box was `fs * 4` = 40pt for a 45pt number, with the
    // rest of the row empty to its right.
    const l = label(bridge([1250000, -45000, 125000, -8000, null]), 2);
    expect(l.text).toBe("+125,000");
    expect(l.fs).toBe(10);
    expect(l.x + l.w).toBeLessThanOrEqual(480);
  });

  it("does not need a wider chart to stop truncating: +1,250,000 is whole at 960pt", () => {
    // "+1,250…" at 960 as well — the box stayed 40pt at every width, so the
    // number only came out whole once the BAR grew big enough to hold it inside.
    const l = label(bridge([125000000, -450000, 1250000, -80000, null], 960), 2);
    expect(l.text).toBe("+1,250,000");
    expect(l.x + l.w).toBeLessThanOrEqual(960);
  });

  it("shrinks a label to the room the bar leaves before it clips it", () => {
    // Pinned so the delta bars end ~40pt short of the chart's edge: the number is
    // wider than that at 10pt, and a smaller whole number beats "+1,234…" — which
    // reads as 1,234.
    const cfg = bridge([1000000000, 1234567, 12345678, -80000, null], 480, { min: 0, max: 1100000000 });
    for (const c of [1, 2]) {
      const l = label(cfg, c);
      expect(l.text).toBe(c === 1 ? "+1,234,567" : "+12,345,678");
      expect(l.fs).toBeLessThan(10);
      expect(l.x + l.w).toBeLessThanOrEqual(480);
    }
  });

  it("leaves a label that already fits its box exactly where it was", () => {
    // The fs * 4 box is the floor: a short number keeps the 40pt box it always had.
    const l = label(bridge([1250000, -45000, 125000, -8000, null]), 3);
    expect(l.text).toBe("-8,000");
    expect(l.w).toBe(40);
    expect(l.fs).toBe(10);
  });

  it("leaves a bar that ends at the frame's edge with the box and font it had", () => {
    // The row holds less than 4em there; there is no wider box to give, and the
    // chart's own frame clip cuts the label at the edge as it always did.
    const cfg = bridge([1000000000, 1234567, 12345678, -80000, null], 480, { min: 0, max: 1013000000 });
    const l = label(cfg, 1);
    expect(l.w).toBe(40);
    expect(l.fs).toBe(10);
  });
});

describe("a doughnut's hole is empty, not painted over (r2-35)", () => {
  /** Whether a node paints an opaque pixel at (x, y). Only the shapes a doughnut can hold. */
  const paints = (n: SceneNode, x: number, y: number): boolean => {
    if (n.kind === "ellipse") return n.fill !== "none" && ((x - n.cx) / n.rx) ** 2 + ((y - n.cy) / n.ry) ** 2 <= 1;
    if (n.kind === "rect") return n.fill !== "none" && x >= n.x && x <= n.x + n.w && y >= n.y && y <= n.y + n.h;
    if (n.kind === "wedge") {
      const d = Math.hypot(x - n.cx, y - n.cy);
      return d >= n.innerR && d <= n.r;
    }
    return false;
  };
  const doughnut = (style?: ChartConfig["style"]): ChartConfig => ({ ...sampleConfig("doughnut"), style });

  it("paints nothing over the centre of the ring, in any slide colour", () => {
    // An opaque ellipse filled with `style.background` — #ffffff by default, never
    // derived from the slide — sat over the middle: a 129pt white patch on a cream,
    // grey or photo slide, exactly where the chart is meant to blend in. The pane's
    // canvas is white, so it could not show the problem and no control reached it.
    for (const style of [undefined, { background: "#202020" }]) {
      const scene = buildChart(doughnut(style));
      const ring = named<WedgeNode>(scene.nodes, "wedge", "slice-")[0];
      expect(scene.nodes.filter((n) => paints(n, ring.cx, ring.cy))).toEqual([]);
    }
  });

  it("draws every slice as a ring segment, the hole being their shared inner radius", () => {
    const wedges = named<WedgeNode>(buildChart(doughnut()).nodes, "wedge", "slice-");
    expect(wedges.length).toBeGreaterThan(1);
    for (const w of wedges) {
      expect(w.innerR).toBeGreaterThan(0);
      expect(w.innerR).toBeCloseTo(w.r * 0.55, 6);
    }
  });

  it("keeps the total in the middle, still fitted to the hole", () => {
    const scene = buildChart(doughnut());
    const wedge = named<WedgeNode>(scene.nodes, "wedge", "slice-")[0];
    const label = named<TextNode>(scene.nodes, "text", "hole-label")[0];
    expect(label).toBeTruthy();
    expect(label.x + label.w / 2).toBeCloseTo(wedge.cx, 6);
    expect(textWidth(label.text, label.fontSize, true)).toBeLessThanOrEqual(wedge.innerR * 2);
  });

  it("is a ring in the picture too: the SVG has no hole shape and each slice carries an inner arc", () => {
    const svg = sceneToSvg(buildChart(doughnut()));
    expect(svg).not.toContain('data-name="hole"');
    const slices = [...svg.matchAll(/<path d="([^"]*)"[^>]*data-name="slice-\d+"/g)];
    expect(slices.length).toBeGreaterThan(1);
    // Outer arc and inner arc: a pie wedge has one.
    for (const [, d] of slices) expect((d.match(/A /g) || []).length).toBe(2);
  });

  it("moves an exploded slice as a whole ring segment", () => {
    const base = doughnut();
    const exploded = buildChart({ ...base, pie: { explode: [0] } });
    const flat = buildChart(base);
    const a = named<WedgeNode>(exploded.nodes, "wedge", "slice-")[0];
    const b = named<WedgeNode>(flat.nodes, "wedge", "slice-")[0];
    expect(Math.hypot(a.cx - b.cx, a.cy - b.cy)).toBeGreaterThan(0);
    expect(a.innerR).toBeGreaterThan(0);
    expect(a.innerR / a.r).toBeCloseTo(0.55, 6);
  });
});
