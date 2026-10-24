import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { layoutPie } from "../src/core/layout/pie";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import { sceneToSvg } from "../src/render/svg";
import { textWidth } from "../src/core/scene";
import type { EllipseNode, LineNode, RectNode, TextNode, WedgeNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Pie / doughnut — breakouts, all-zero totals, narrow-frame radius, full-circle slice. */

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

describe("pie breakout (bar-of-pie)", () => {
  const cfg: ChartConfig = {
    kind: "pie",
    ...DEFAULT_SIZE,
    data: {
      categories: ["EMEA", "Americas", "APAC", "Nordics", "Benelux", "DACH"],
      series: [{ name: "Revenue", values: [80, 100, 60, 20, 15, 25] }],
    },
    pie: { breakout: [3, 4, 5] },
  };
  const s = buildChart(cfg);

  it("collapses breakout categories into one muted Other slice facing the bar", () => {
    expect(s.nodes.some((n) => n.name === "slice-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "slice-3")).toBe(false);
    const other = s.nodes.find((n) => n.name === "slice-other") as WedgeNode;
    expect(other.fill).toBe("#898781");
    // Other (60 of 300 = 72°) is centered at 3 o'clock: spans 90° ± 36°.
    const mid = ((other.startAngle + other.endAngle) / 2) % 360;
    expect(mid).toBeCloseTo(90, 1);
  });

  it("details the breakout in a stacked bar with connectors and grand-total %", () => {
    const segs = [3, 4, 5].map((c) => s.nodes.find((n) => n.name === `breakout-seg-${c}`) as RectNode);
    expect(segs.every(Boolean)).toBe(true);
    // Stacked contiguously, heights ∝ values (20/15/25 of 60).
    expect(segs[0].y + segs[0].h).toBeCloseTo(segs[1].y, 5);
    expect(segs[1].y + segs[1].h).toBeCloseTo(segs[2].y, 5);
    expect(segs[2].h / segs[0].h).toBeCloseTo(25 / 20, 5);
    // Bar sits right of the pie.
    const other = s.nodes.find((n) => n.name === "slice-other") as WedgeNode;
    expect(segs[0].x).toBeGreaterThan(other.cx);
    // Labels carry the share of the GRAND total (20/300 ≈ 7%).
    expect((s.nodes.find((n) => n.name === "breakout-label-3") as TextNode).text).toContain("7%");
    const conns = s.nodes.filter((n): n is LineNode => !!n.name?.startsWith("breakout-conn"));
    expect(conns).toHaveLength(2);
    // Connectors join the bar's top and bottom.
    const ends = conns.map((c) => c.y2).sort((a, b) => a - b);
    expect(ends[0]).toBeCloseTo(segs[0].y, 5);
    expect(ends[1]).toBeCloseTo(segs[2].y + segs[2].h, 5);
  });

  it("plain pies and doughnuts are unaffected", () => {
    const plain = buildChart({ ...cfg, pie: {} });
    expect(plain.nodes.some((n) => n.name === "slice-3")).toBe(true);
    expect(plain.nodes.some((n) => n.name === "slice-other")).toBe(false);
    const dough = buildChart({ ...cfg, kind: "doughnut" });
    expect(dough.nodes.some((n) => n.name === "slice-other")).toBe(false);
  });
});

describe("variable-radius pie", () => {
  const cfg: ChartConfig = {
    kind: "pie",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "Share", values: [50, 30, 20] }, // angle
        { name: "Radius", values: [20, 80, 50] }, // radius metric
      ],
    },
    pie: { variableRadius: true },
  };
  const s = buildChart(cfg);
  const slice = (c: number) => s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === `slice-${c}`)!;

  it("keeps angle ∝ the first series but sets radius from the Radius row", () => {
    // Angle still encodes Share: A (50) sweeps the widest.
    const spanA = slice(0).endAngle - slice(0).startAngle;
    const spanB = slice(1).endAngle - slice(1).startAngle;
    expect(spanA).toBeGreaterThan(spanB);
    // Radius encodes the Radius row: B (80) longest, A (20) shortest — even
    // though A has the biggest share.
    expect(slice(1).r).toBeGreaterThan(slice(2).r);
    expect(slice(2).r).toBeGreaterThan(slice(0).r);
  });

  it("is inert without a Radius row or the flag", () => {
    const plain = buildChart({
      ...cfg,
      pie: {},
      data: { categories: ["A", "B"], series: [{ name: "Share", values: [50, 50] }] },
    });
    const a = plain.nodes.find((n): n is WedgeNode => n.name === "slice-0")!;
    const b = plain.nodes.find((n): n is WedgeNode => n.name === "slice-1")!;
    expect(a.r).toBe(b.r); // uniform radius
  });
});

/** Regression tests for defects found during a codebase bug-hunt pass. */
describe("full-circle pie / doughnut (single slice)", () => {
  it("renders a visible circle path for a 360° pie", () => {
    const cfg: ChartConfig = {
      kind: "pie",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [100] }] },
    };
    const scene = buildChart(cfg);
    const wedge = scene.nodes.find((n) => n.kind === "wedge");
    expect(wedge).toBeTruthy();
    // The lone slice spans the whole circle.
    expect((wedge as any).endAngle - (wedge as any).startAngle).toBeCloseTo(360);
    const svg = sceneToSvg(scene);
    // Must emit a real path with area — the old code drew a degenerate arc
    // between coincident points, producing nothing.
    expect(svg).toMatch(/fill-rule="evenodd"/);
    expect(svg).toMatch(/<path[^>]*A /);
  });

  it("renders a full ring, with a real hole, for a 360° doughnut", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [100] }] },
    };
    const scene = buildChart(cfg);
    // The lone slice is a full-circle ANNULAR wedge. The hole used to be faked
    // with an overpainted background ellipse "for Office.js compatibility", and
    // that stopped being true a day after it was written (4714c27 added annular
    // sectors to all three renderers; the gauge and the sunburst use them), so
    // the ring is drawn as a ring and the slide shows through the middle of it.
    const wedge = scene.nodes.find((n) => n.kind === "wedge") as any;
    expect(wedge.endAngle - wedge.startAngle).toBeCloseTo(360);
    expect(wedge.innerR).toBeGreaterThan(0);
    expect(scene.nodes.some((n) => n.kind === "ellipse")).toBe(false);
    const svg = sceneToSvg(scene);
    expect(svg).toMatch(/fill-rule="evenodd"/);
    // A real ring, not a degenerate zero-area arc: the outer circle and the inner
    // one, two arcs each, filled evenodd so the inner one is the hole.
    const path = svg.match(/<path d="([^"]*)"[^>]*fill-rule="evenodd"/)!;
    expect(path[1]).not.toMatch(/NaN/);
    expect((path[1].match(/A /g) || []).length).toBeGreaterThanOrEqual(4);
  });
});

describe("pie radius never goes negative on a narrow frame", () => {
  it("floors the radius so wedges and the hole keep valid geometry", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 80, // width*0.5 - fs*7 = -30 before the floor
      height: 320,
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [3, 4, 5] }] },
    };
    const nodes = buildChart(cfg).nodes;
    for (const n of nodes) {
      if (n.kind === "wedge") {
        expect((n as WedgeNode).r).toBeGreaterThanOrEqual(0);
        expect((n as WedgeNode).innerR).toBeGreaterThanOrEqual(0);
      }
      if (n.kind === "ellipse") {
        expect((n as EllipseNode).rx).toBeGreaterThanOrEqual(0);
        expect((n as EllipseNode).ry).toBeGreaterThanOrEqual(0);
      }
    }
  });
});

describe("pie & doughnut", () => {
  const c = cfg({
    kind: "pie",
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [50, 30, 20] }] },
  });

  it("slices sum to 360° in data order", () => {
    const { nodes } = layoutPie(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const wedges = nodes.filter((n): n is WedgeNode => n.kind === "wedge");
    expect(wedges).toHaveLength(3);
    expect(wedges[0].endAngle - wedges[0].startAngle).toBeCloseTo(180, 5);
    expect(wedges[2].endAngle).toBeCloseTo(360, 5);
  });

  it("doughnut leaves a real hole with the total in it", () => {
    const scene = buildChart({ ...c, kind: "doughnut" });
    const wedges = scene.nodes.filter((n): n is WedgeNode => n.kind === "wedge");
    expect(wedges).toHaveLength(3);
    // Every slice is a ring segment with one inner radius — nothing is painted over
    // the middle, so a slide that is not white shows through it.
    for (const w of wedges) expect(w.innerR).toBeCloseTo(w.r * 0.55, 6);
    expect(scene.nodes.find((n) => n.name === "hole")).toBeUndefined();
    const label = scene.nodes.find((n) => n.name === "hole-label") as TextNode;
    expect(label.text).toBe("100");
  });
});

describe("pie / doughnut all-zero total", () => {
  it("shows the true total (0), not the divisor fallback of 1", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 400,
      height: 300,
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [0, 0, 0] }] },
    };
    const nodes = buildChart(cfg).nodes;
    const texts = nodes.filter((n): n is TextNode => n.kind === "text");
    expect(nodes.some((n) => n.name === "hole-label")).toBe(true);
    // The centre shows the honest total (0, at the data's 2-decimal precision),
    // never the divisor fallback the old `|| 1` displayed as "1".
    expect(texts.some((t) => /^0(\.0+)?$/.test(t.text))).toBe(true);
    expect(texts.some((t) => /^1(\.0+)?$/.test(t.text))).toBe(false);
  });

  it("still renders a normal doughnut total unchanged", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 400,
      height: 300,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [30, 70] }] },
    };
    const texts = buildChart(cfg).nodes.filter((n): n is TextNode => n.kind === "text");
    expect(texts.some((t) => t.text === "100")).toBe(true);
  });
});

describe("pie fallbacks", () => {
  it("tolerates missing values and empty series names", () => {
    const s = buildChart({
      kind: "pie",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B", "C"], series: [{ name: "", values: [3, 2] }] },
      decorations: { segmentLabels: true },
    } as ChartConfig);
    expect(s.nodes.some((n) => n.kind === "wedge")).toBe(true);
  });
});

describe("a slice label stays on the chart", () => {
  /**
   * The outside label ran away from the slice edge with nothing to stop it. The
   * radius reserves a FIXED `fs * 7` for labels — a guess, where the breakout
   * path in the same file MEASURES the widest label it actually has — so any
   * category name wider than that guess put ink off the chart.
   *
   * Measured before the fix: `label-0` reached x = 548 on a 480pt frame, 68pt
   * past the right edge. Neither PowerPoint renderer wraps or clips a text box,
   * so in a deck that is a label lying across whatever sits beside the chart on
   * the slide — and off a picture-mode render entirely.
   */
  const inkRight = (t: TextNode) => {
    const w = Math.min(t.w, textWidth(t.text, t.fontSize, t.bold));
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    return x + w;
  };
  const labels = (cfg: ChartConfig) =>
    buildChart(cfg).nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("label-"));

  const longCfg = (kind: "pie" | "doughnut"): ChartConfig =>
    ({
      kind,
      ...DEFAULT_SIZE,
      data: {
        categories: ["A very long category label indeed", "B", "Another rather long one here", "D"],
        series: [{ name: "S", values: [4, 3, 2, 1] }],
      },
      decorations: { segmentLabels: true },
    }) as ChartConfig;

  it("keeps a long category label inside the frame", () => {
    for (const kind of ["pie", "doughnut"] as const) {
      const cfg = longCfg(kind);
      const found = labels(cfg);
      expect(found.length, `${kind} drew no slice labels`).toBeGreaterThan(0);
      for (const t of found) {
        expect(inkRight(t), `${kind} ${t.name} ran past the right edge`).toBeLessThanOrEqual(DEFAULT_SIZE.width);
        expect(t.x, `${kind} ${t.name} started left of the frame`).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("still shows as much of the label as fits, and a short one whole", () => {
    // The rule must not be satisfiable by drawing nothing, or by ellipsizing
    // every label regardless of the room available.
    const long = labels(longCfg("doughnut"));
    expect(
      long.some((t) => t.text.length > 8),
      "every label was clipped to nothing",
    ).toBe(true);

    const shortCfg = {
      kind: "doughnut",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 1] }] },
      decorations: { segmentLabels: true },
    } as ChartConfig;
    for (const t of labels(shortCfg)) {
      expect(t.text.endsWith("…"), `a short label was clipped: ${JSON.stringify(t.text)}`).toBe(false);
    }
  });
});

/**
 * The slices TILE the circle: each one starts where the last ended, and together
 * they cover 360 degrees exactly once.
 *
 * This is the invariant a pie is FOR, and it is the one that no frame or ink
 * check can see — every wedge stays inside the chart whether it is placed
 * correctly or stacked on top of its neighbour. It broke here for real: an early
 * `return` added to skip an outer label also skipped the `angle += span` at the
 * end of the same callback, so every slice after the first started at zero. The
 * frame sweep passed, the snapshots passed (they are taken at one size), and the
 * doughnut simply showed the wrong data at every small size.
 *
 * Swept over the sizes rather than one, because the bug was reachable only where
 * the label branch it hid in was.
 */
describe("a pie's slices tile the circle", () => {
  const SIZES: [number, number][] = [
    [120, 90],
    [160, 120],
    [200, 150],
    [300, 60],
    [480, 300],
  ];
  const data = {
    categories: ["EMEA", "Americas", "APAC", "Other"],
    series: [{ name: "Revenue", values: [29, 38, 24, 8] }],
  };

  /** A slice's radius, which is the same for every slice. */
  const arcR = (kind: string, width: number, height: number) =>
    buildChart({ kind, width, height, data, decorations: { segmentLabels: true } } as ChartConfig).nodes.find(
      (n): n is WedgeNode => n.kind === "wedge",
    )!.r;

  it("gives up the outer label margin rather than collapse to a dot", () => {
    // That margin is a flat `fs * 7` — 70pt either side of a 120pt-wide frame —
    // so a pie under ~140pt wide had nothing left and fell to the 1pt floor: a
    // 2pt dot, 0.1% of a thumbnail in ink, with four labels drawn around it as
    // though there were a chart there. The margin yields, and on the axis where
    // yielding is not enough the ring comes off, the way the radar's web and the
    // sunburst's ring already do.
    expect(arcR("pie", 120, 90)).toBeGreaterThan(20);
    expect(arcR("doughnut", 120, 90)).toBeGreaterThan(20);
    // A frame far too short for a label ring is a separate reservation and
    // yields separately: the labels shrink until they hit their floor, so the
    // arc here is smaller than a thumbnail's but is still an arc.
    expect(arcR("pie", 300, 60)).toBeGreaterThan(12);
  });

  it("takes the FULL margin once the frame can pay for it", () => {
    // Above about 280pt wide the flat `fs * 7` fits inside the share below, so
    // it is the margin — not the share, and not the rescue — that sets the
    // radius, exactly as it always did.
    expect(arcR("pie", 300, 400)).toBe(300 * 0.5 - 10 * 7);
    expect(arcR("pie", 400, 400)).toBe(400 * 0.5 - 10 * 7);
    // At 480x300 the HEIGHT term binds instead, and it is untouched either way.
    expect(arcR("pie", 480, 300)).toBeLessThan(480 * 0.5 - 10 * 7);
  });

  /**
   * The arc never shrinks when the frame grows.
   *
   * This is the property the flat margins broke, and they broke it hard: taken
   * in full the moment the frame could afford them, they collapsed the arc at
   * the threshold. Growing a chart from 160 to 170 points wide took its radius
   * from 75 to 15, and a 280pt-wide pie was no bigger than a 160pt one — a chart
   * that gets SMALLER as its frame grows, which is a bug whatever the arithmetic
   * says.
   *
   * A residual step of a few points survives where the outer labels hit their
   * size floor and the ring is dropped outright; `TOLERATED` is that floor
   * crossing and nothing else. It is checked rather than waved at: one step, of
   * a handful of points, on a frame too small to read.
   */
  it("never shrinks when the frame grows", () => {
    const TOLERATED = 6;
    for (const [axis, fixed] of [
      ["width", 300],
      ["height", 480],
    ] as const) {
      const steps: string[] = [];
      let prev = 0;
      let worst = 0;
      for (let v = 60; v <= 480; v += 5) {
        const r = axis === "width" ? arcR("pie", v, fixed) : arcR("pie", fixed, v);
        if (r < prev - 0.001) {
          steps.push(`${axis}=${v}: ${prev.toFixed(1)} -> ${r.toFixed(1)}`);
          worst = Math.max(worst, prev - r);
        }
        prev = r;
      }
      expect(steps.length, `${axis}: ${steps.join(", ")}`).toBeLessThanOrEqual(1);
      expect(worst, `${axis}: worst drop`).toBeLessThanOrEqual(TOLERATED);
    }
  });

  for (const kind of ["pie", "doughnut"] as const) {
    for (const [width, height] of SIZES) {
      it(`${kind} at ${width}x${height}`, () => {
        const wedges = buildChart({
          kind,
          width,
          height,
          data,
          decorations: { segmentLabels: true },
        } as ChartConfig).nodes.filter((n): n is WedgeNode => n.kind === "wedge");

        expect(wedges).toHaveLength(4);
        // Contiguous: each slice begins where its predecessor ended.
        for (let i = 1; i < wedges.length; i++) {
          expect(wedges[i].startAngle, `slice ${i} does not start where slice ${i - 1} ends`).toBeCloseTo(
            wedges[i - 1].endAngle % 360,
            5,
          );
        }
        // And they cover the circle once, so no slice is drawn over another.
        const swept = wedges.reduce((t, w) => t + (w.endAngle - w.startAngle), 0);
        expect(swept, "the slices do not cover the circle exactly once").toBeCloseTo(360, 5);
      });
    }
  }
});

/**
 * An INSIDE label fits the slice it sits in.
 *
 * These were drawn at the chart font and fitted to nothing, so a name wider than
 * its own wedge was drawn across the neighbouring slices — and on a small frame
 * past the edge of the chart, where the frame clip cut it to an ellipsis. Both
 * the preview and the deck showed "mericas 38…" on a 200x150 pie.
 *
 * An OUTSIDE label is bounded by the frame instead and has its own rule, so it
 * is identified here by the leader line it carries and skipped.
 */
describe("a pie label drawn inside its slice", () => {
  /**
   * The room the label has: the chord of its slice at the radius the label sits
   * on. Deliberately the same arithmetic the layout uses — the point is that the
   * label is fitted to its own wedge, and a second, subtly different chord here
   * would be measuring something the layout never promised.
   */
  const chord = (spanDeg: number, r: number) => 2 * (r * 0.62) * Math.sin((Math.min(spanDeg, 180) * Math.PI) / 360);

  const SIZES: [number, number][] = [
    [120, 90],
    [160, 120],
    [200, 150],
    [300, 60],
    [300, 200],
    [480, 300],
  ];

  for (const [width, height] of SIZES) {
    it(`fits at ${width}x${height}`, () => {
      const nodes = buildChart({ ...sampleConfig("pie"), width, height } as ChartConfig).nodes;
      const wedges = nodes.filter((n): n is WedgeNode => n.kind === "wedge");
      expect(wedges.length).toBeGreaterThan(1);

      let checked = 0;
      for (const w of wedges) {
        const c = w.name?.replace("slice-", "");
        const label = nodes.find((n): n is TextNode => n.kind === "text" && n.name === `label-${c}`);
        // A slice too thin for an inside label is labelled outside, with a
        // leader line: a different rule, checked by the frame sweep instead.
        if (!label || nodes.some((n) => n.name === `leader-${c}`)) continue;
        checked++;
        const ink = textWidth(label.text, label.fontSize, label.bold);
        expect(ink, `${JSON.stringify(label.text)} is wider than slice ${c}`).toBeLessThanOrEqual(
          chord(w.endAngle - w.startAngle, w.r) + 0.5,
        );
      }
      expect(checked, "no inside label was actually checked").toBeGreaterThan(0);
    });
  }

  it("leaves a chart with room alone, at the chart's own font", () => {
    // Last resort, like every other shrink in this engine: at a size where the
    // labels already fit, nothing moves.
    const nodes = buildChart({ ...sampleConfig("pie"), width: 480, height: 300 } as ChartConfig).nodes;
    const inside = nodes.filter(
      (n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("label-") && n.align === "center",
    );
    expect(inside.length).toBeGreaterThan(0);
    for (const t of inside) {
      expect(t.fontSize, `${JSON.stringify(t.text)} was shrunk on a chart with room`).toBe(10);
      expect(t.text.endsWith("…"), `${JSON.stringify(t.text)} was clipped on a chart with room`).toBe(false);
    }
  });
});
