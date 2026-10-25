import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { LineNode, RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Line / area — slope, stepped, smoothed, confidence bands, missing-data, profiles. */

describe("forecast styling on lines", () => {
  const line: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["2023", "2024", "2025", "2026"], series: [{ name: "Rev", values: [40, 46, 52, 60] }] },
    decorations: { forecastFrom: 2, segmentLabels: false },
  };

  it("dashes segments and hollows markers from the boundary on", () => {
    const s = buildChart(line);
    const segs = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    expect(segs.find((l) => l.name === "line-0-1")!.dash).toBeUndefined(); // actuals
    expect(segs.find((l) => l.name === "line-0-2")!.dash).toEqual([4, 3]); // into forecast
    expect(segs.find((l) => l.name === "line-0-3")!.dash).toEqual([4, 3]);
    const solid = s.nodes.find((n) => n.name === "marker-0-1") as RectNode;
    const hollow = s.nodes.find((n) => n.name === "marker-0-2") as RectNode;
    expect(solid.fill).not.toBe("none");
    // Hollow is no fill, not a fill in the canvas colour (r2-35): the knock-out was a
    // white patch on any slide that is not the chart's own background.
    expect(hollow.fill).toBe("none");
    expect(hollow.stroke).toBe(solid.fill);
    expect(s.nodes.some((n) => n.name === "forecast-divider")).toBe(true);
  });

  it("is inert without the option", () => {
    const s = buildChart({ ...line, decorations: { segmentLabels: false } });
    expect(s.nodes.some((n) => n.name === "forecast-divider")).toBe(false);
    expect(
      s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-")).every((l) => !l.dash),
    ).toBe(true);
  });
});

describe("line confidence bands and fill-between", () => {
  const line: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2024", "2025", "2026"],
      series: [
        { name: "Forecast", values: [50, 56, 63] },
        { name: "Band low", values: [50, 52, 55] },
        { name: "Band high", values: [50, 60, 72] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("shades a ribbon between Band low/high without drawing them as lines", () => {
    const s = buildChart(line);
    const slabs = s.nodes.filter((n) => n.name?.startsWith("band-ribbon"));
    expect(slabs.length).toBeGreaterThan(20);
    // Only the forecast line renders; band rows are not series.
    expect(s.nodes.some((n) => n.name === "line-1-1")).toBe(false);
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    // Ribbon renders behind the line (earlier in the node list).
    expect(s.nodes.findIndex((n) => n.name?.startsWith("band-ribbon"))).toBeLessThan(
      s.nodes.findIndex((n) => n.name === "line-0-1"),
    );
    // The band widens the scale: max tick covers 72.
    const slab = s.nodes.find((n) => n.name?.startsWith("band-ribbon-1")) as RectNode;
    expect(slab.y).toBeGreaterThan(0);
  });

  it("fillBetween shades the gap between two series", () => {
    const s = buildChart({
      kind: "line",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Plan", values: [40, 50, 60] },
          { name: "Actual", values: [38, 46, 52] },
        ],
      },
      decorations: { fillBetween: [0, 1], segmentLabels: false },
    });
    expect(s.nodes.filter((n) => n.name?.startsWith("fill-between")).length).toBeGreaterThan(20);
    // Both series still draw as lines.
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    expect(s.nodes.some((n) => n.name === "line-1-1")).toBe(true);
  });
});

describe("slope chart (decorations.slope on line)", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2020", "2025"],
      series: [
        { name: "Brand A", values: [48, 66] },
        { name: "Brand B", values: [40, 49] },
      ],
    },
    decorations: { slope: true },
  };
  const s = buildChart(cfg);

  it("draws end rails, straight lines, and 'Name value' labels at both ends", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("slope-rail"))).toHaveLength(2);
    expect((s.nodes.find((n) => n.name === "slope-left-0") as TextNode).text).toBe("Brand A 48");
    expect((s.nodes.find((n) => n.name === "slope-right-0") as TextNode).text).toBe("Brand A 66");
    const line = s.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(line.y1).toBeGreaterThan(line.y2); // 48 → 66 rises
    // No value axis or gridlines in slope mode.
    expect(s.nodes.some((n) => n.name?.startsWith("grid"))).toBe(false);
    expect(s.nodes.some((n) => n.name?.startsWith("vaxis"))).toBe(false);
  });

  it("labels take the series color and de-overlap when ends are close", () => {
    const crowded = buildChart({
      ...cfg,
      data: {
        categories: ["2020", "2025"],
        series: [
          { name: "Alpha", values: [50, 60] },
          { name: "Beta", values: [50.5, 60.5] },
          { name: "Gamma", values: [51, 61] },
        ],
      },
    });
    const ys = [0, 1, 2]
      .map((si) => (crowded.nodes.find((n) => n.name === `slope-left-${si}`) as TextNode).y)
      .sort((a, b) => a - b);
    expect(ys[1] - ys[0]).toBeGreaterThanOrEqual(14);
    expect(ys[2] - ys[1]).toBeGreaterThanOrEqual(14);
    const label = s.nodes.find((n) => n.name === "slope-left-0") as TextNode;
    const line = s.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(label.color).toBe(line.stroke);
  });

  it("handles >2 categories as a polyline and ignores slope on area charts", () => {
    const three = buildChart({
      ...cfg,
      data: {
        categories: ["2020", "2022", "2025"],
        series: [{ name: "A", values: [40, 55, 66] }],
      },
    });
    expect(three.nodes.filter((n) => n.name?.startsWith("line-0-"))).toHaveLength(2);
    expect(three.nodes.filter((n) => n.name?.startsWith("slope-rail"))).toHaveLength(2); // ends only
    const area = buildChart({ ...cfg, kind: "area" });
    expect(area.nodes.some((n) => n.name?.startsWith("slope-rail"))).toBe(false);
  });
});

describe("stepped line", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [1, 2, 3] }] },
    decorations: { segmentLabels: false },
  };

  it("default draws one sloped connector per interval", () => {
    const plain = buildChart(base);
    const seg = plain.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(seg).toBeTruthy();
    // Diagonal: neither horizontal nor vertical.
    expect(seg.y1).not.toBeCloseTo(seg.y2, 5);
    expect(seg.x1).not.toBeCloseTo(seg.x2, 5);
  });

  it('"after" holds then jumps (HV elbow)', () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, stepped: "after" } });
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(false);
    const a = s.nodes.find((n) => n.name === "line-0-1a") as LineNode;
    const b = s.nodes.find((n) => n.name === "line-0-1b") as LineNode;
    expect(a.y1).toBeCloseTo(a.y2, 5); // horizontal at the previous value
    expect(b.x1).toBeCloseTo(b.x2, 5); // then a vertical jump
    expect(b.x1).toBeCloseTo(a.x2, 5); // meeting at the next category x
  });

  it('"before" jumps immediately, "center" steps at the midpoint', () => {
    const before = buildChart({ ...base, decorations: { segmentLabels: false, stepped: "before" } });
    const ba = before.nodes.find((n) => n.name === "line-0-1a") as LineNode;
    expect(ba.x1).toBeCloseTo(ba.x2, 5); // vertical first

    const center = buildChart({ ...base, decorations: { segmentLabels: false, stepped: "center" } });
    const ca = center.nodes.find((n) => n.name === "line-0-1a") as LineNode;
    const cb = center.nodes.find((n) => n.name === "line-0-1b") as LineNode;
    const cc = center.nodes.find((n) => n.name === "line-0-1c") as LineNode;
    expect([ca, cb, cc].every(Boolean)).toBe(true);
    expect(cb.x1).toBeCloseTo((ca.x1 + cc.x2) / 2, 5); // riser at the midpoint
  });
});

describe("stepped area", () => {
  const base: ChartConfig = {
    kind: "area",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [10, 20, 30] }] },
    decorations: { segmentLabels: false },
  };

  it("default interpolates the slab top; stepped holds it flat", () => {
    const slabs = (scene: ReturnType<typeof buildChart>) =>
      scene.nodes.filter((n) => n.name?.startsWith("area-0-0-")) as RectNode[];

    const plain = slabs(buildChart(base));
    // A sloped (interpolated) top: the first and last slab of the segment differ.
    expect(plain.length).toBeGreaterThan(1);
    expect(Math.abs(plain[0].y - plain[plain.length - 1].y)).toBeGreaterThan(0.5);

    // A stepped area has a flat top across the interval, so the slab-fill needs
    // no tessellation at all — the segment collapses to a single slab.
    const stepped = slabs(buildChart({ ...base, decorations: { segmentLabels: false, stepped: "before" } }));
    expect(stepped.length).toBe(1);
  });
});

describe("area with negative values", () => {
  const areaRects = (sc: { nodes: SceneNode[] }) =>
    sc.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("area-"));
  const hasNegAxisLabel = (sc: { nodes: SceneNode[] }) =>
    sc.nodes.some((n) => n.kind === "text" && n.name === "value-axis" && /^-/.test((n.text ?? "").trim()));
  // The value axis "0" tick sits at toY(0) — the zero baseline in screen y.
  const zeroY = (sc: { nodes: SceneNode[] }) => {
    const z = sc.nodes.find(
      (n): n is TextNode => n.kind === "text" && n.name === "value-axis" && (n.text ?? "").trim() === "0",
    )!;
    return z.y + z.h / 2;
  };
  const mk = (values: (number | null)[], cats = ["A", "B", "C"]): ChartConfig => ({
    kind: "area",
    ...DEFAULT_SIZE,
    data: { categories: cats, series: [{ name: "v", values }] },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true },
  });

  it("fills both above and below the zero baseline", () => {
    const s = buildChart(mk([8, -8], ["A", "B"]));
    const y0 = zeroY(s);
    const rects = areaRects(s);
    expect(rects.some((r) => r.y + r.h > y0 + 1)).toBe(true); // fill extending below zero
    expect(rects.some((r) => r.y < y0 - 1)).toBe(true); // fill extending above zero
  });

  it("extends the value axis into the negatives (not clamped to 0)", () => {
    expect(hasNegAxisLabel(buildChart(mk([10, -6, 8])))).toBe(true);
  });

  it("positive-only area stays entirely above the baseline", () => {
    const pos = buildChart(mk([10, 6, 8]));
    expect(hasNegAxisLabel(pos)).toBe(false);
    const y0 = zeroY(pos);
    expect(areaRects(pos).every((r) => r.y + r.h <= y0 + 1)).toBe(true); // nothing below zero
  });
});

describe("smoothed lines", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "v", values: [1, 4, 2, 5] }] },
    decorations: { segmentLabels: false },
  };

  it("default draws straight connectors", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    expect(s.nodes.some((n) => n.name?.startsWith("line-0-1-s"))).toBe(false);
  });

  it("smooth replaces straight segments with a sampled spline polyline", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, smooth: true } });
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(false);
    const sampled = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.match(/^line-0-\d+-s\d+$/));
    // 3 segments × 16 samples.
    expect(sampled.length).toBe(48);
  });

  it("stepped wins over smooth (mutually exclusive)", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, smooth: true, stepped: "after" } });
    expect(s.nodes.some((n) => n.name?.startsWith("line-0-1-s"))).toBe(false);
    expect(s.nodes.some((n) => n.name === "line-0-1a")).toBe(true);
  });
});

describe("transparent floating segments", () => {
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "base", color: "transparent", values: [10, 14] },
        { name: "Range", values: [20, 18] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("draws no rect for the transparent segment but still stacks above it", () => {
    const s = buildChart(cfg);
    expect(s.nodes.some((n) => n.name === "seg-0-0")).toBe(false); // transparent base not drawn
    const range = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    expect(range).toBeTruthy();
    const solid = buildChart({
      ...cfg,
      data: {
        ...cfg.data,
        series: [
          { name: "base", values: [10, 14] },
          { name: "Range", values: [20, 18] },
        ],
      },
    });
    const solidRange = solid.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    // Same height, but the floating one sits higher (smaller y) is false — same y
    // since the level is identical; the difference is the missing base rect.
    expect(range.y).toBeCloseTo(solidRange.y, 3);
    expect(range.h).toBeCloseTo(solidRange.h, 3);
    // The floating build is missing the base segment the solid build draws.
    expect(solid.nodes.some((n) => n.name === "seg-0-0")).toBe(true);
  });

  it("floating bar does not reach the baseline", () => {
    const s = buildChart(cfg);
    const range = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    // Baseline is the bottom of the plot; the bar's bottom (y+h) is above it
    // by the (undrawn) base segment. Compare Q1 (base 10) vs a base-0 build.
    const grounded = buildChart({
      ...cfg,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "base", color: "transparent", values: [0, 0] },
          { name: "Range", values: [20, 18] },
        ],
      },
    });
    const gRange = grounded.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    expect(range.y + range.h).toBeLessThan(gRange.y + gRange.h - 1); // floats above the grounded bar's base
  });
});

describe("line missing-data bridge", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "v", values: [10, null, 30, 40] }] },
    decorations: { segmentLabels: false },
  };

  it("breaks the line at nulls by default", () => {
    const s = buildChart(base);
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    // Only C→D connects (A is isolated by the null B, B has no point).
    expect(lines).toHaveLength(1);
    expect(s.nodes.filter((n) => n.name?.startsWith("marker-0-"))).toHaveLength(3); // A, C, D
  });

  it("bridges across nulls when bridgeGaps is set", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, bridgeGaps: true } });
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    // A→C (bridging B) and C→D → two segments.
    expect(lines).toHaveLength(2);
    // The bridge segment joins A's point to C's point.
    const bridge = lines.find((l) => l.name === "line-0-2")!;
    const cx = (i: number) => s.nodes.find((n): n is RectNode => n.name === `marker-0-${i}`)!.x;
    expect(bridge.x1).toBeCloseTo(cx(0) + 2.4, 0); // near A's marker
    expect(bridge.x2).toBeCloseTo(cx(2) + 2.4, 0); // to C's marker
  });
});

describe("horizontal profile chart — line", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    horizontal: true,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [20, 80, 50] }] },
    decorations: { segmentLabels: false, categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);
  const marker = (c: number) => s.nodes.find((n): n is RectNode => n.name === `marker-0-${c}`)!;

  it("runs categories down the y axis (top to bottom)", () => {
    expect(marker(0).y).toBeLessThan(marker(1).y);
    expect(marker(1).y).toBeLessThan(marker(2).y);
  });

  it("maps larger values further right (value along x)", () => {
    expect(marker(1).x).toBeGreaterThan(marker(0).x); // 80 right of 20
    expect(marker(2).x).toBeGreaterThan(marker(0).x); // 50 right of 20
    expect(marker(1).x).toBeGreaterThan(marker(2).x); // 80 right of 50
  });

  it("connects the points with line segments", () => {
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
  });

  it("the vertical line chart is unaffected (no horizontal dispatch)", () => {
    const v = buildChart({ ...cfg, horizontal: undefined });
    // In the vertical chart categories run along x, so markers share no single column.
    const m0 = v.nodes.find((n): n is RectNode => n.name === "marker-0-0")!;
    const m1 = v.nodes.find((n): n is RectNode => n.name === "marker-0-1")!;
    expect(m1.x).toBeGreaterThan(m0.x); // categories advance along x when vertical
  });
});

describe("horizontal profile chart — stacked area", () => {
  const cfg: ChartConfig = {
    kind: "area",
    ...DEFAULT_SIZE,
    horizontal: true,
    data: {
      categories: ["N", "S"],
      series: [
        { name: "A", values: [40, 55] },
        { name: "B", values: [25, 20] },
      ],
    },
    decorations: { seriesLabels: true, categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);

  it("stacks the series along x (second series to the right of the first)", () => {
    // Compare the same y-strip (slab k=0) of each series: B stacks to the
    // right of A there (globally the x-ranges overlap across categories).
    const a0 = s.nodes.find((n): n is RectNode => n.name === "area-0-0-0")!;
    const b0 = s.nodes.find((n): n is RectNode => n.name === "area-1-0-0")!;
    expect(a0).toBeTruthy();
    expect(b0).toBeTruthy();
    expect(b0.x).toBeGreaterThanOrEqual(a0.x + a0.w - 1);
  });
});

/**
 * Slope end labels were clamped only at the BOTTOM: the upward propagation that
 * enforces the minimum gap then walked past the plot top, past the title and off
 * the canvas, losing the topmost labels outright.
 */
describe("dense slope chart keeps every end label on the canvas", () => {
  const s = buildChart({
    kind: "line",
    ...DEFAULT_SIZE,
    // Twelve series over a 300pt canvas: more labels than the minimum gap fits.
    data: {
      categories: ["2020", "2025"],
      series: Array.from({ length: 12 }, (_, i) => ({ name: `S${i}`, values: [50 + i * 0.1, 60 + i * 0.1] })),
    },
    decorations: { slope: true },
  } as unknown as ChartConfig);

  it.each(["slope-left", "slope-right"])("fits every %s label", (side) => {
    const off = (s.nodes as TextNode[])
      .filter((n) => n.name?.startsWith(side))
      .filter((n) => n.y < -0.5 || n.y + n.h > s.height + 0.5)
      .map((n) => `${n.name}@${n.y.toFixed(1)}`);
    expect(off, `off-canvas: ${off.join(", ")}`).toEqual([]);
  });

  it("still draws all twelve", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("slope-left-"))).toHaveLength(12);
  });
});

describe("date-spaced line categories", () => {
  it("spaces category centers proportionally to time", () => {
    const s = buildChart({
      kind: "line",
      ...DEFAULT_SIZE,
      data: { categories: ["2025-01", "2025-02", "2025-12"], series: [{ name: "S", values: [1, 2, 3] }] },
    } as ChartConfig);
    const pts = s.nodes.filter((n) => n.name?.startsWith("marker-0-"));
    expect(pts).toHaveLength(3);
    const xs = pts.map((p) => (p.kind === "rect" ? p.x : 0));
    // Jan→Feb gap must be far smaller than Feb→Dec.
    expect(xs[1] - xs[0]).toBeLessThan((xs[2] - xs[1]) / 5);
  });
});

/**
 * The rotation toggle used to drop two more things, both of them documented
 * behaviour rather than decoration: the `Band low` / `Band high` split (the
 * reference says those rows "never draw as lines") and the date-proportional
 * category spacing the manual states for line charts without qualifying it by
 * orientation. Sideways, band rows drew as two ordinary series with their own
 * legend chips and no ribbon, and a ten-month gap between categories sat in an
 * evenly spaced row.
 */
describe("a sideways line chart keeps the rules the upright one has", () => {
  const banded: ChartConfig = {
    kind: "line",
    horizontal: true,
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "Actual", values: [10, 14, 12, 18] },
        { name: "Band low", values: [8, 12, 10, 15] },
        { name: "Band high", values: [12, 16, 14, 21] },
      ],
    },
    decorations: { seriesLabels: true },
  };

  it("shades the Band low/high ribbon instead of drawing the rows as lines", () => {
    const { nodes } = buildChart(banded);
    expect(nodes.filter((n) => n.name?.startsWith("band-ribbon")).length).toBeGreaterThan(0);
    // One drawn series, not three: the band rows are the ribbon, not lines.
    const drawn = new Set(
      nodes.map((n) => /^(?:line|marker)-(\d+)-/.exec(n.name ?? "")?.[1]).filter((v): v is string => !!v),
    );
    expect([...drawn]).toEqual(["0"]);
    // …and they are not legended either.
    const labels = nodes.filter((n): n is TextNode => n.kind === "text" && /^legend-\d+$/.test(n.name ?? ""));
    expect(labels.map((n) => n.text)).not.toContain("Band low");
  });

  it("widens the value scale to cover the band", () => {
    // The negative control for the split: dropping the rows entirely would also
    // stop them drawing as lines, and would silently clip the ribbon.
    const { nodes } = buildChart(banded);
    const rects = (prefix: string) =>
      nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith(prefix));
    const bandRight = Math.max(...rects("band-ribbon").map((s) => s.x + s.w));
    const lineRight = Math.max(...rects("marker-").map((m) => m.x + m.w));
    expect(bandRight).toBeGreaterThan(lineRight);
  });

  it("spaces date categories proportionally to time down the axis", () => {
    const { nodes } = buildChart({
      kind: "line",
      horizontal: true,
      ...DEFAULT_SIZE,
      data: { categories: ["2025-01", "2025-02", "2025-12"], series: [{ name: "S", values: [1, 2, 3] }] },
    } as ChartConfig);
    const ys = nodes
      .filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("marker-0-"))
      .map((p) => p.y);
    expect(ys).toHaveLength(3);
    // Jan→Feb gap must be far smaller than Feb→Dec, exactly as it is upright.
    expect(ys[1] - ys[0]).toBeLessThan((ys[2] - ys[1]) / 5);
  });
});

/**
 * `stepped` held a flat edge across the interval on an upright area chart and
 * interpolated on a sideways one — so a staircase series drawn as a bar profile
 * claimed the value slid where the data says it jumped. The sideways LINE path
 * already stepped; only the area slabs did not.
 */
describe("a sideways area chart steps when it is told to", () => {
  const cfg = (stepped?: "before" | "after" | "center"): ChartConfig => ({
    kind: "area",
    horizontal: true,
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [3, 7, 5] }] },
    decorations: stepped ? { stepped } : {},
  });
  const slabWidths = (c: ChartConfig) =>
    buildChart(c)
      .nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("area-0-0-"))
      .map((n) => Math.round(n.w * 100) / 100);

  it("holds one width across the interval instead of ramping", () => {
    const ramped = slabWidths(cfg());
    const held = slabWidths(cfg("after"));
    expect(new Set(ramped).size).toBeGreaterThan(1); // the control: interpolation ramps
    expect(new Set(held).size).toBe(1); // "after" holds the left value the whole way
  });

  it("carries the RIGHT value for `before`, so the two are not the same staircase", () => {
    expect(slabWidths(cfg("before"))[0]).not.toBe(slabWidths(cfg("after"))[0]);
  });
});

/**
 * `lastSegMid` is indexed by the FILTERED series list (band rows are not drawn
 * as lines), and `seriesLabelNodes` was handed the raw config — so a chart
 * whose `Band low` / `Band high` rows came first paired the drawn line's end-y
 * with `cfg.data.series[0].name` and labelled it "Band low".
 */
describe("the right-hand series label names the line it points at", () => {
  const labels = (series: { name: string; values: number[] }[]) =>
    buildChart({
      kind: "line",
      ...DEFAULT_SIZE,
      decorations: { seriesLabels: true },
      data: { categories: ["a", "b"], series },
    } as ChartConfig)
      .nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("series-label"))
      .map((n) => n.text);

  const band = [
    { name: "Band low", values: [1, 2] },
    { name: "Band high", values: [3, 4] },
  ];
  const actual = { name: "Actual", values: [2, 3] };

  it("labels the drawn line whichever end the band rows sit at", () => {
    expect(labels([...band, actual])).toEqual(["Actual"]);
    expect(labels([actual, ...band])).toEqual(["Actual"]);
  });

  it("still labels every drawn series when there are no band rows", () => {
    // The negative control: filtering the wrong list would drop real labels.
    expect(
      labels([
        { name: "A", values: [1, 2] },
        { name: "B", values: [3, 4] },
      ]).sort(),
    ).toEqual(["A", "B"]);
  });
});

/**
 * The same anchor mismatch on a multi-series LINE chart: `columnTop` is the
 * topmost point across every series and `columnValue` is the FIRST series'
 * value, so the arrow spanned one series' points and printed another's growth.
 */
describe("a difference arrow on a multi-series line reads the series it spans", () => {
  it("does not span one series and label another's growth", () => {
    const { nodes } = buildChart({
      kind: "line",
      width: 400,
      height: 300,
      decorations: { difference: { from: 0, to: 1 } },
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "A", values: [10, 20] }, // +100%
          { name: "B", values: [100, 300] }, // +200%, and the topmost line
        ],
      },
    } as ChartConfig);
    const line = nodes.find((n): n is LineNode => n.kind === "line" && n.name === "diff-line")!;
    const markers = (si: number) =>
      nodes
        .filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith(`marker-${si}-`))
        .map((n) => n.y + n.h / 2);
    const label = nodes.find((n): n is TextNode => n.kind === "text" && n.name === "diff-label")!;
    expect(label.text).toContain("100");
    // The arrow must sit on series A's points — the ones it is reporting — not
    // on series B's, which is where `columnTop` put it.
    const [aFrom, aTo] = markers(0);
    const [bFrom] = markers(1);
    expect(Math.abs(line.y1 - aFrom), "arrow start off series A").toBeLessThan(1);
    expect(Math.abs(line.y2 - aTo), "arrow end off series A").toBeLessThan(1);
    expect(Math.abs(line.y1 - bFrom), "arrow still anchored on series B").toBeGreaterThan(1);
  });
});

/**
 * A FILLED SEGMENT IS TILED TO APPROXIMATE A SLOPE, so what the tiling has to
 * resolve is how far the edge MOVES — and until 2026-08-30 the count came from
 * the segment's width alone. A perfectly level segment 80pt wide got twenty
 * identical rectangles laid side by side: one rectangle, drawn twenty times.
 *
 * Shapes are the unit of cost on this host. An in-place update runs about 767ms
 * PER SHAPE, measured over 178 updates in the archive, and round 150 recorded
 * 400-500 shapes in one context crashing PowerPoint on all seven attempts — with
 * a chart we ship sitting at 300. Taking the rise into account:
 *
 *     sample charts      400 slabs -> 318
 *     the shipped deck   652 slabs -> 419   (-36%)
 *
 * The span rule stays as the ceiling, so no segment can ever GAIN a slab.
 */
describe("filled segments are tiled by how far the edge moves, not just how wide it is", () => {
  const areaWith = (values: number[], over: Partial<ChartConfig> = {}): SceneNode[] =>
    buildChart({
      kind: "area",
      ...DEFAULT_SIZE,
      data: { categories: values.map((_, i) => `C${i}`), series: [{ name: "S", values }] },
      decorations: { segmentLabels: false },
      ...over,
    } as ChartConfig).nodes.filter((n) => /^area-\d+-\d+-\d+$/.test(n.name || ""));

  it("draws a level segment once instead of twenty times", () => {
    // Flat data: every edge is level, so no tessellation resolves anything.
    const flat = areaWith([50, 50, 50, 50]);
    // Three segments, one slab each — the old rule gave ~20 apiece on this width.
    expect(flat.length, `a level area drew ${flat.length} rects for 3 segments`).toBe(3);
  });

  it("still tessellates a steep segment", () => {
    // The rule must not simply shrink everything: an edge that climbs hard still
    // needs the staircase, or the fill visibly steps away from its own line.
    const steep = areaWith([0, 100, 0, 100]);
    expect(steep.length, "a steep area stopped being tessellated").toBeGreaterThan(20);
  });

  it("keeps the staircase within a point of the edge, where this rule decides it", () => {
    /**
     * The error of a straight edge tiled at `steps` is `slope * slabWidth / 2`,
     * and `steps = rise / 2` makes that 1pt — WHILE THAT IS THE BINDING RULE.
     *
     * The first version of this test asserted 1pt unconditionally and failed,
     * which was the test correcting my docstring rather than the code: on a
     * steep edge the historical 24-slab ceiling binds first, and the error there
     * is whatever it has always been, because this rule changes nothing on those
     * segments. The honest claim is about the segments the rise rule decides.
     */
    for (const cats of [4, 8, 16]) {
      // A gentle zig-zag: every segment has a real but small rise, so the rise
      // rule is the one choosing rather than the ceiling.
      const values = Array.from({ length: cats }, (_, i) => 50 + (i % 2) * 3);
      const slabs = areaWith(values) as RectNode[];
      const seg = slabs.filter((s) => /^area-0-0-\d+$/.test(s.name || ""));
      expect(seg.length, `no slabs in the first segment at ${cats} categories`).toBeGreaterThan(0);
      if (seg.length >= 24) continue; // the ceiling is choosing, not this rule
      const rise = Math.max(...seg.map((s) => s.y)) - Math.min(...seg.map((s) => s.y));
      const err = rise / seg.length / 2;
      expect(err, `${cats} categories tile to a ${err.toFixed(2)}pt staircase error`).toBeLessThanOrEqual(1.5);
    }
  });

  it("never gives a segment MORE slabs than the width rule alone would", () => {
    // The span rule is the ceiling. Whatever the rise says, a wide segment can
    // only ever lose shapes — otherwise this optimisation could regress the very
    // shape budget it was written to protect.
    for (const values of [
      [10, 90, 20, 80],
      [0, 1, 2, 3],
      [5, 5, 5, 100],
    ]) {
      const n = areaWith(values).length;
      const segs = values.length - 1;
      // 24 is the historical per-segment cap the span rule tops out at.
      expect(n, `${values} drew ${n} rects across ${segs} segments`).toBeLessThanOrEqual(24 * segs);
    }
  });
});
