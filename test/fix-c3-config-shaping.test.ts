import { describe, expect, it } from "vitest";
import { buildAgendaScene, SLIDE } from "../src/core/agenda";
import { buildChart, normalizeConfig, valueExtent } from "../src/core/chart";
import { planReconcile, type ExpectedItem, type SlideSnapshot } from "../src/core/reconcile";
import { spreadAlongAxis } from "../src/core/spread";
import { textWidth } from "../src/core/scene";
import type { RectNode, Scene, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * Config shaping — the findings of stream c3: otherBucket, sorting, lanes,
 * multiples, scatter inputs, the reconcile keeper and the agenda slide.
 *
 * Each block names the finding it guards and the input that reached it. Every
 * test that guards a fix was shown to FAIL against the code it replaced, on the
 * assertion that targets the bug; the few that pass either way are controls
 * pinning what must not move, and say so in their comments.
 */

const cfg = (c: unknown) => c as ChartConfig;

/** Every text node's text, by name. */
function texts(scene: Scene): Record<string, string> {
  const out: Record<string, string> = {};
  for (const n of scene.nodes) if (n.kind === "text" && n.name) out[n.name] = String(n.text);
  return out;
}

/** The category labels in the order they are drawn left to right / top to bottom. */
function categoryLabels(scene: Scene): string[] {
  return scene.nodes
    .filter((n): n is TextNode => n.kind === "text" && /^category-\d+$/.test(n.name ?? ""))
    .sort((a, b) => +a.name!.split("-")[1] - +b.name!.split("-")[1])
    .map((n) => n.text);
}

// ───────────────────────────────────────────────────────────────────────────
// r1-18 — otherBucket turned a category with no data into a printed "0"
// ───────────────────────────────────────────────────────────────────────────
describe("r1-18: a category blank in every collapsed series gets no Other value", () => {
  const cats = ["Q1", "Q2", "Q3", "Q4"];
  // Q3 has nothing in ANY series — a quarter not reported yet, pasted as an
  // empty column, which the datasheet turns into nulls.
  const allBlankQ3 = [
    { name: "A", values: [40, 46, null, 50] },
    { name: "B", values: [22, 26, null, 30] },
    { name: "C", values: [18, 17, null, 19] },
    { name: "D", values: [6, 7, null, 8] },
    { name: "E", values: [4, 5, null, 5] },
  ];
  // Only the collapsed tail is blank at Q3; the kept series have data there.
  const tailBlankQ3 = [
    { name: "A", values: [40, 46, 48, 50] },
    { name: "B", values: [22, 26, 27, 30] },
    { name: "C", values: [18, 17, null, 19] },
    { name: "D", values: [6, 7, null, 8] },
    { name: "E", values: [4, 5, null, 5] },
  ];
  const build = (kind: string, series: unknown, other: boolean, decorations: object = { totals: true }) =>
    buildChart(
      cfg({
        kind,
        width: 480,
        height: 300,
        data: { categories: cats, series },
        decorations,
        ...(other ? { otherBucket: { max: 3 } } : {}),
      }),
    );

  for (const kind of ["stacked", "clustered"]) {
    it(`${kind}: prints no total over a column nobody filled in`, () => {
      const t = texts(build(kind, allBlankQ3, true));
      expect(t["category-2"], "the column keeps its heading").toBe("Q3");
      expect(t["total-2"] ?? "", `Q3 is blank in every series and total-2 says ${JSON.stringify(t["total-2"])}`).toBe(
        "",
      );
      // …and the columns that DO have data still carry theirs, or the fix
      // deleted the feature instead of gating it.
      expect(t["total-0"]).toMatch(/\d/);
      expect(t["total-3"]).toMatch(/\d/);
    });
  }

  it("a mean line averages the columns that have data, not a fabricated zero", () => {
    const mean = { totals: true, valueLine: { mode: "mean" } };
    const plain = texts(build("stacked", allBlankQ3, false, mean))["value-line-label-0"];
    const bucketed = texts(build("stacked", allBlankQ3, true, mean))["value-line-label-0"];
    expect(plain).toMatch(/Ø/);
    expect(bucketed, "the bucket must not change what the mean of the real columns is").toBe(plain);
  });

  it("grand total does not gain a total for the blank column either", () => {
    const t = texts(build("stacked", allBlankQ3, true, { totals: true, grandTotal: true }));
    expect(t["total-2"] ?? "").toBe("");
  });

  it("control: a blank TAIL under kept data still totals what the kept series hold", () => {
    // 48 + 27 = 75 whether or not the tail is folded into Other — the zero the
    // old code made was invisible here, which is why only the all-blank column
    // is asserted above. This pins that the fix did not blank a real total.
    for (const other of [false, true]) expect(texts(build("stacked", tailBlankQ3, other))["total-2"]).toBe("75");
  });

  it("control: an Other bucket over real data still sums the tail", () => {
    // C+D+E at Q1 = 18 + 6 + 4 = 28; the stack total is unchanged by folding.
    const withOther = texts(build("stacked", allBlankQ3, true));
    const without = texts(build("stacked", allBlankQ3, false));
    expect(withOther["total-0"]).toBe(without["total-0"]);
    expect(withOther["total-3"]).toBe(without["total-3"]);
  });

  it("pareto: a blank bar stays blank instead of printing a measured 0", () => {
    // The sibling that wrote the same `?? 0` into a series it synthesizes.
    const t = texts(
      buildChart(
        cfg({
          kind: "clustered",
          width: 480,
          height: 300,
          pareto: true,
          decorations: { totals: true },
          data: { categories: ["A", "B", "C", "D"], series: [{ name: "V", values: [40, null, 25, 10] }] },
        }),
      ),
    );
    // Sorted by value, the blank category sorts last: C, D... then B.
    const blank = Object.entries(t)
      .find(([, v]) => v === "B")![0]
      .replace("category-", "");
    expect(t[`total-${blank}`] ?? "", "the blank category was given a measured total").toBe("");
    expect(t["total-0"]).toBe("40");
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r1-46 — a non-numeric otherBucket.max collapsed every series into one Other
// ───────────────────────────────────────────────────────────────────────────
describe("r1-46: otherBucket.max that is not a number falls back to the default", () => {
  const six = Array.from({ length: 6 }, (_, i) => ({ name: `S${i + 1}`, values: [60 - i * 8, 50 - i * 7] }));
  const build = (kind: string, otherBucket: unknown) =>
    buildChart(cfg({ kind, width: 480, height: 300, otherBucket, data: { categories: ["A", "B"], series: six } }));
  const seriesCount = (s: Scene) => /(\d+) data series/.exec(s.desc ?? "")![1];

  for (const kind of ["stacked", "clustered", "stacked100"]) {
    for (const [label, max] of [
      ["abc", "abc"],
      ["x", "x"],
      ["{}", {}],
      ["[1, 2]", [1, 2]],
      ["NaN", NaN],
    ] as [string, unknown][]) {
      it(`${kind}: max ${label} keeps the default five series`, () => {
        expect(seriesCount(build(kind, { max }))).toBe("5");
      });
    }
  }

  it("repairs the value once, at the boundary", () => {
    expect(
      normalizeConfig(cfg({ kind: "stacked", data: { categories: [], series: [] }, otherBucket: { max: "abc" } }))
        .otherBucket,
    ).toEqual({});
  });

  it("Same scale measures the bars that are drawn", () => {
    const junk = valueExtent(
      cfg({ kind: "clustered", data: { categories: ["A", "B"], series: six }, otherBucket: { max: "abc" } }),
    );
    const dflt = valueExtent(cfg({ kind: "clustered", data: { categories: ["A", "B"], series: six } }));
    expect(junk).toEqual(dflt);
    expect(junk!.max).toBeLessThan(100);
  });

  it("control: a numeric string and a number are still read as the number they are", () => {
    expect(seriesCount(build("stacked", { max: "3" }))).toBe("3");
    expect(seriesCount(build("stacked", { max: 4 }))).toBe("4");
    // Two is the floor — the largest series and Other.
    expect(seriesCount(build("stacked", { max: 1 }))).toBe("2");
  });

  it("an empty string or a bare true is not a count either", () => {
    // Math.floor("") is 0 and Math.floor(true) is 1: both clamped to the floor
    // of two and quietly folded a six-series chart into S6 + Other.
    expect(seriesCount(build("stacked", { max: "" }))).toBe("5");
    expect(seriesCount(build("stacked", { max: true }))).toBe("5");
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r1-19 — Gantt owner lanes drew unowned tasks inside the last owner's lane
// ───────────────────────────────────────────────────────────────────────────
describe("r1-19: unowned tasks under gantt.lanes get a header of their own", () => {
  const day = (iso: string) => Math.round(Date.parse(iso) / 86400000);
  const plan = (over: Partial<ChartConfig> = {}, categories?: string[]) =>
    cfg({
      kind: "gantt",
      width: 560,
      height: 320,
      gantt: { lanes: "owner" },
      data: {
        categories: categories ?? ["Spec | Ana", "Build | Ben", "Review | Ana", "Ship | Ben", "Handover"],
        series: [
          {
            name: "Start",
            values: [day("2026-01-05"), day("2026-01-12"), day("2026-01-19"), day("2026-01-26"), day("2026-02-02")],
          },
          {
            name: "End",
            values: [day("2026-01-12"), day("2026-01-19"), day("2026-01-26"), day("2026-02-02"), day("2026-02-05")],
          },
        ],
        dates: true,
      },
      ...over,
    });
  const rect = (scene: Scene, name: string) => scene.nodes.find((n) => n.name === name) as RectNode;
  /** Row index of a drawn label. */
  const rowOf = (scene: Scene, text: string) => categoryLabels(scene).indexOf(text);

  it("puts a synthesized header between the last lane and the unowned tasks", () => {
    expect(categoryLabels(buildChart(plan()))).toEqual([
      "Ana",
      "Spec",
      "Review",
      "Ben",
      "Build",
      "Ship",
      "Other",
      "Handover",
    ]);
  });

  it("draws the unowned header as a section band, like every owner's", () => {
    const scene = buildChart(plan());
    expect(rect(scene, `section-${rowOf(scene, "Other")}`), "no band for the unowned header").toBeTruthy();
  });

  it("a lane's summary bar ends where that owner's work ends", () => {
    const scene = buildChart(plan({ decorations: { summaryBars: true } }));
    const ben = rect(scene, `summary-${rowOf(scene, "Ben")}`);
    const ship = rect(scene, `bar-${rowOf(scene, "Ship")}`);
    const handover = rect(scene, `bar-${rowOf(scene, "Handover")}`);
    // Ben's last task is Ship (ends 2 Feb); Handover (to 5 Feb) is nobody's.
    expect(ben.x + ben.w).toBeCloseTo(ship.x + ship.w, 3);
    expect(ben.x + ben.w).toBeLessThan(handover.x + handover.w - 1);
  });

  it("the unowned tasks get a summary of their own", () => {
    const scene = buildChart(plan({ decorations: { summaryBars: true } }));
    const other = rect(scene, `summary-${rowOf(scene, "Other")}`);
    const handover = rect(scene, `bar-${rowOf(scene, "Handover")}`);
    expect(other).toBeTruthy();
    expect(other.x + other.w).toBeCloseTo(handover.x + handover.w, 3);
  });

  it("names the header with the caller's translated word for Other", () => {
    expect(categoryLabels(buildChart(plan({ labels: { other: "Øvrige" } }))).at(-2)).toBe("Øvrige");
  });

  it("control: a plan where every task has an owner gets no extra header", () => {
    const owned = plan({}, ["Spec | Ana", "Build | Ben", "Review | Ana", "Ship | Ben", "Handover | Ben"]);
    expect(categoryLabels(buildChart(owned))).toEqual(["Ana", "Spec", "Review", "Ben", "Build", "Ship", "Handover"]);
  });

  it("control: with lanes off the rows are exactly as written", () => {
    expect(categoryLabels(buildChart(plan({ gantt: {} })))).toEqual(["Spec", "Build", "Review", "Ship", "Handover"]);
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r1-41 — pie/doughnut "Sort descending" ranked by every row, drew only the first
// ───────────────────────────────────────────────────────────────────────────
describe("r1-41: a pie sorts by the series it draws", () => {
  const order = (c: unknown) => {
    const t = texts(buildChart(cfg(c)));
    return Object.keys(t)
      .filter((k) => /^label-\d+$/.test(k))
      .sort((a, b) => +a.split("-")[1] - +b.split("-")[1])
      .map((k) => t[k].split(" ")[0]);
  };
  const pie = (kind: string, series: unknown, over: object = {}) => ({
    kind,
    width: 480,
    height: 300,
    categorySort: "descending",
    data: { categories: ["A", "B", "C"], series },
    ...over,
  });
  const slices = { name: "2024", values: [10, 50, 30] };

  it("control: a single-series pie sorts by its slices", () => {
    expect(order(pie("pie", [slices]))).toEqual(["B", "C", "A"]);
    expect(order(pie("pie", [slices], { categorySort: "ascending" }))).toEqual(["A", "C", "B"]);
  });

  for (const kind of ["pie", "doughnut"]) {
    it(`${kind}: a second row left over from another chart does not decide the order`, () => {
      expect(order(pie(kind, [slices, { name: "2025", values: [100, 0, 0] }]))).toEqual(["B", "C", "A"]);
    });
  }

  it("a Radius row is a size, not a share", () => {
    expect(order(pie("pie", [slices, { name: "Radius", values: [100, 1, 1] }]))).toEqual(["B", "C", "A"]);
  });

  it("a variable-radius pie sorts by angle whatever the second series says", () => {
    expect(
      order(pie("pie", [slices, { name: "Margin", values: [100, 1, 1] }], { pie: { variableRadius: true } })),
    ).toEqual(["B", "C", "A"]);
  });

  it("ascending is read off the same series", () => {
    expect(order(pie("pie", [slices, { name: "2025", values: [100, 0, 0] }], { categorySort: "ascending" }))).toEqual([
      "A",
      "C",
      "B",
    ]);
  });

  it("control: a stacked column still ranks by the whole stack", () => {
    // The other SORTABLE kinds draw every row, so every row is the right total.
    const s = buildChart(
      cfg({
        kind: "stacked",
        width: 480,
        height: 300,
        categorySort: "descending",
        data: {
          categories: ["A", "B", "C"],
          series: [slices, { name: "2025", values: [100, 0, 0] }],
        },
      }),
    );
    expect(categoryLabels(s)).toEqual(["A", "B", "C"]); // 110, 50, 30
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r1-44 — small-multiples alt text described the undrawn config
// ───────────────────────────────────────────────────────────────────────────
describe("r1-44: a small-multiples chart is described as it is drawn", () => {
  const multi = (extraSeries: unknown[] = [], over: object = {}) => ({
    kind: "clustered",
    width: 480,
    height: 300,
    multiples: {},
    data: {
      categories: ["Small", "Big", "Mid"],
      series: [{ name: "North", values: [5, 30, 15] }, { name: "South", values: [4, 20, 10] }, ...extraSeries],
    },
    ...over,
  });

  it("lists the categories in the order the panels plot them", () => {
    const c = multi([], { categorySort: "descending" });
    const scene = buildChart(cfg(c));
    expect(scene.desc).toBe("clustered column chart. 2 data series: North, South. 3 categories: Big, Mid, Small.");
    // …which is the order the panels actually run in.
    expect(texts(scene)["p0-category-0"]).toBe("Big");
  });

  it("does not count a carried Error row as a data series", () => {
    const scene = buildChart(cfg(multi([{ name: "Error", values: [1, 2, 1] }])));
    expect(scene.desc).toContain("2 data series: North, South.");
    expect(scene.desc).not.toContain("Error");
  });

  it("does not count a Target row as a data series", () => {
    const scene = buildChart(cfg(multi([{ name: "Target", values: [8, 25, 12] }])));
    expect(scene.desc).toContain("2 data series: North, South.");
  });

  it("says the same thing the same chart says when it is not split", () => {
    const withRows = multi([{ name: "Error", values: [1, 2, 1] }], { categorySort: "descending" });
    const split = buildChart(cfg(withRows)).desc;
    const whole = buildChart(cfg({ ...withRows, multiples: undefined })).desc;
    expect(split).toBe(whole);
  });

  it("control: a plain multiples chart is described as before", () => {
    expect(buildChart(cfg(multi())).desc).toBe(
      "clustered column chart. 2 data series: North, South. 3 categories: Small, Big, Mid.",
    );
  });

  it("control: a grid that cannot be drawn falls back to the whole chart's description", () => {
    // Too short for two rows of panels: buildMultiples declines and the single
    // chart path describes what it draws.
    const scene = buildChart(cfg(multi([], { height: 12, multiples: { columns: 1 } })));
    expect(scene.desc).toContain("2 data series");
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r2-40 — a fractional multiples.columns placed panels at half-column offsets
// ───────────────────────────────────────────────────────────────────────────
describe("r2-40: multiples.columns is a whole number of columns", () => {
  const grid = (columns?: unknown) =>
    buildChart(
      cfg({
        kind: "clustered",
        width: 480,
        height: 300,
        multiples: { columns },
        data: {
          categories: ["A", "B", "C"],
          series: [1, 2, 3, 4, 5].map((i) => ({ name: `S${i}`, values: [i * 3, i * 4, i * 5] })),
        },
      }),
    );
  const pastEdge = (scene: Scene) =>
    scene.nodes.filter((n) => n.kind === "rect" && n.x + n.w > scene.width + 0.01).map((n) => n.name);

  for (const [columns, whole] of [
    [2.5, 2],
    [2.2, 2],
    [1.5, 1],
    ["2.5", 2],
  ] as [unknown, number][]) {
    it(`columns ${JSON.stringify(columns)} lays out exactly as ${whole}`, () => {
      expect(grid(columns).nodes).toEqual(grid(whole).nodes);
    });
    it(`columns ${JSON.stringify(columns)} keeps every bar on the chart`, () => {
      expect(pastEdge(grid(columns))).toEqual([]);
    });
  }

  it("control: whole numbers and the default grid are untouched", () => {
    expect(pastEdge(grid(2))).toEqual([]);
    expect(pastEdge(grid(3))).toEqual([]);
    expect(pastEdge(grid(undefined))).toEqual([]);
    // 0.5 and below were already clamped to one column.
    expect(grid(0.5).nodes).toEqual(grid(1).nodes);
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r1-34 / r1-35 — a non-numeric scatter.spreadLimit / trendDegree
// ───────────────────────────────────────────────────────────────────────────
describe("r1-34: a non-numeric scatter.spreadLimit", () => {
  const scatter = (spreadLimit?: unknown) =>
    buildChart(
      cfg({
        kind: "scatter",
        width: 480,
        height: 300,
        scatter: { spread: "x", ...(spreadLimit === undefined ? {} : { spreadLimit }) },
        data: {
          categories: ["a", "b", "c", "d"],
          series: [
            { name: "X", values: [1, 1, 1.01, 5] },
            { name: "Y", values: [1, 1, 1, 5] },
          ],
        },
      }),
    );
  const points = (s: Scene) => s.nodes.filter((n) => /^point-\d+$/.test(n.name ?? "")).length;
  const note = (s: Scene) => String(texts(s)["footnote"] ?? "");

  it("spreadAlongAxis moves nothing when its cap is not a number", () => {
    // The function's own contract: no item ever moves further than `limit`.
    // NaN passed `limit <= 0` and turned every displacement into NaN, including
    // for a marker that overlaps nothing.
    const items = [
      { m: 100, c: 50, r: 5 },
      { m: 101, c: 50, r: 5 },
      { m: 400, c: 50, r: 5 },
    ];
    expect(spreadAlongAxis(items, { limit: NaN, min: 0, max: 480 })).toEqual([0, 0, 0]);
    // control: a real cap still spreads the overlapping pair and leaves the third.
    const moved = spreadAlongAxis(items, { limit: 5, min: 0, max: 480 });
    expect(moved[0]).toBeLessThan(0);
    expect(moved[1]).toBeGreaterThan(0);
    expect(moved[2]).toBe(0);
  });

  for (const [label, value] of [
    ["auto", "auto"],
    ["{}", {}],
    ["[]", []],
    ["NaN", NaN],
  ] as [string, unknown][]) {
    it(`${label} draws every point with the default cap and names a number in the footnote`, () => {
      const s = scatter(value);
      expect(points(s)).toBe(4);
      expect(note(s)).toMatch(/±\d/);
      expect(note(s)).toBe(note(scatter()));
    });
  }

  it("control: a numeric string is still read as the number it is", () => {
    const s = scatter("0.05");
    expect(points(s)).toBe(4);
    expect(note(s)).toContain("±0.050");
  });
});

describe("r1-35: a non-numeric scatter.trendDegree", () => {
  const trend = (trendDegree?: unknown, n = 6) =>
    cfg({
      kind: "scatter",
      width: 480,
      height: 300,
      scatter: trendDegree === undefined ? {} : { trendDegree },
      data: {
        categories: Array.from({ length: n }, (_, i) => `p${i}`),
        series: [
          { name: "X", values: Array.from({ length: n }, (_, i) => i + 1) },
          { name: "Y", values: Array.from({ length: n }, (_, i) => (i + 1) ** 2) },
          { name: "Trend", values: [1, ...Array.from({ length: n - 1 }, () => null)] },
        ],
      },
    });

  for (const [label, value] of [
    ["two", "two"],
    ["quadratic", "quadratic"],
    ["{}", {}],
    ["[1, 2]", [1, 2]],
    ["NaN", NaN],
  ] as [string, unknown][]) {
    for (const n of [4, 6]) {
      it(`${label} with ${n} points falls back to the linear fit instead of throwing`, () => {
        let scene: Scene | undefined;
        expect(() => (scene = buildChart(trend(value, n)))).not.toThrow();
        expect(scene!.nodes).toEqual(buildChart(trend(undefined, n)).nodes);
      });
    }
  }

  it("control: a numeric string and a number are still read as the degree they are", () => {
    expect(buildChart(trend("2")).nodes).toEqual(buildChart(trend(2)).nodes);
    expect(texts(buildChart(trend(2)))["trend-stats"]).toContain("quadratic");
  });
});

// ───────────────────────────────────────────────────────────────────────────
// The same root cause in the siblings of r1-46 / r1-34 / r1-35: an authored
// number that is not one reaches a clamp that passes NaN through.
// ───────────────────────────────────────────────────────────────────────────
describe("siblings: a non-numeric authored count or limit falls back to its default", () => {
  const segs = (s: Scene) => s.nodes.filter((n) => /^seg-\d+-\d+$/.test(n.name ?? "")).length;
  const columns = (over: object) =>
    buildChart(
      cfg({
        kind: "clustered",
        width: 480,
        height: 300,
        data: {
          categories: ["A", "B", "C", "D"],
          series: [
            { name: "S1", values: [10, 20, 30, 40] },
            { name: "S2", values: [15, 25, 35, 45] },
          ],
        },
        ...over,
      }),
    );

  for (const key of ["gapWidth", "overlap"]) {
    for (const bad of ["abc", "50%", {}, NaN]) {
      it(`${key} ${JSON.stringify(bad)} still draws the columns, at the default`, () => {
        expect(segs(columns({ [key]: bad })), "a column chart with no columns").toBe(8);
        expect(columns({ [key]: bad }).nodes).toEqual(columns({}).nodes);
      });
    }
    it(`control: ${key} as a number, and as a numeric string, is the number it spells`, () => {
      expect(columns({ [key]: "30" }).nodes).toEqual(columns({ [key]: 30 }).nodes);
      expect(columns({ [key]: 30 }).nodes).not.toEqual(columns({}).nodes);
    });
  }

  it("butterfly.split that is no number keeps the classic one-series-a-side layout", () => {
    // `slice(0, NaN)` is empty: no left flank, every series piled onto the right.
    const fly = (split?: unknown) =>
      buildChart(
        cfg({
          kind: "butterfly",
          width: 480,
          height: 300,
          butterfly: split === undefined ? {} : { split },
          data: {
            categories: ["A", "B", "C"],
            series: [
              { name: "Left", values: [10, 20, 30] },
              { name: "Right", values: [12, 18, 33] },
            ],
          },
        }),
      );
    for (const bad of ["abc", {}, NaN]) expect(fly(bad).nodes, JSON.stringify(bad)).toEqual(fly().nodes);
    // control: a real split still moves the flanks.
    expect(fly(2).nodes).not.toEqual(fly(1).nodes);
  });

  it("boxplot.iqrMultiplier that is no number still reports the outliers", () => {
    // NaN fences make every `v < loFence || v > hiFence` false: the outlier is
    // not drawn, and the whisker swallows it — data the chart then fails to show.
    const box = (iqrMultiplier?: unknown) =>
      buildChart(
        cfg({
          kind: "boxplot",
          width: 480,
          height: 300,
          boxplot: iqrMultiplier === undefined ? {} : { iqrMultiplier },
          data: {
            categories: ["A", "B"],
            series: [1, 2, 3, 4, 5, 6, 7, 40].map((v, i) => ({ name: `r${i}`, values: [v, v + 1] })),
          },
        }),
      );
    const outliers = (s: Scene) => s.nodes.filter((n) => /outlier/.test(n.name ?? "")).length;
    expect(outliers(box()), "the control must draw outliers for this to mean anything").toBeGreaterThan(0);
    for (const bad of ["abc", {}, NaN]) {
      expect(outliers(box(bad)), JSON.stringify(bad)).toBe(outliers(box()));
      expect(box(bad).nodes, JSON.stringify(bad)).toEqual(box().nodes);
    }
    // control: a wide fence still moves them, so the option is not being ignored.
    expect(outliers(box(1000))).toBe(0);
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r1-45 — reconcile kept the loose untagged twin and deleted the grouped one
// ───────────────────────────────────────────────────────────────────────────
describe("r1-45: on equal content the keeper is the re-editable twin", () => {
  const slide = (index: number, over: Partial<SlideSnapshot> = {}): SlideSnapshot => ({
    index,
    slot: 3,
    title: "Line",
    run: "r",
    shapes: 36,
    stamped: false,
    tagged: false,
    ...over,
  });
  const item = (over: Partial<ExpectedItem> = {}): ExpectedItem => ({
    slot: 3,
    title: "Line",
    shapes: 36,
    chart: true,
    ...over,
  });
  const opts = { run: "r", dropOrphanBlanks: true };
  const deletes = (plan: ReturnType<typeof planReconcile>) =>
    plan.actions.filter((a) => a.kind === "delete").map((a) => a.index);

  // Attempt 1 timed out before grouping and left 36 loose shapes; attempt 2
  // grouped and tagged on a new, later slide. Both ended up in the deck.
  const loose = (over: Partial<SlideSnapshot> = {}) => slide(4, over);
  const grouped = (over: Partial<SlideSnapshot> = {}) =>
    slide(5, { shapes: 1, groupChildren: 36, grouped: true, tagged: true, ...over });

  it("keeps the grouped, tagged copy and deletes the loose one", () => {
    const plan = planReconcile([loose(), grouped()], [item()], opts);
    expect(plan.verdicts[0].index).toBe(5);
    expect(plan.verdicts[0].tagged).toBe(true);
    expect(deletes(plan)).toEqual([4]);
  });

  it("does not leave the sole survivor un-editable when the loose copy's tags are unreadable", () => {
    // The harm case: the loose keeper's tag read failed, the run did not record
    // a failed write, so no regroup or retag was planned for it — and the one
    // copy that was re-editable was deleted as its twin.
    const plan = planReconcile([loose({ tagRead: false }), grouped()], [item()], opts);
    expect(plan.verdicts[0].index).toBe(5);
    expect(plan.verdicts[0].tagged).toBe(true);
    expect(plan.summary.undetermined).toBe(0);
    expect(deletes(plan)).toEqual([4]);
  });

  it("needs no regroup once the keeper is the grouped copy", () => {
    const plan = planReconcile([loose(), grouped()], [item()], opts);
    expect(plan.actions.filter((a) => a.kind === "regroup" || a.kind === "retag")).toEqual([]);
  });

  it("prefers a grouped copy that merely lacks its tag to a loose pile", () => {
    // One retag is cheaper and safer than regrouping 36 loose shapes.
    const plan = planReconcile([loose(), grouped({ tagged: false })], [item()], opts);
    expect(plan.verdicts[0].index).toBe(5);
    expect(plan.actions.filter((a) => a.kind === "retag").map((a) => a.index)).toEqual([5]);
    expect(deletes(plan)).toEqual([4]);
  });

  it("control: the grouped copy first is still the one kept", () => {
    const plan = planReconcile([grouped({ index: 4 }), loose({ index: 5 })], [item()], opts);
    expect(plan.verdicts[0].index).toBe(4);
    expect(deletes(plan)).toEqual([5]);
  });

  it("control: more content still beats re-editability", () => {
    // A 36-shape loose pile against a 30-child group: the group is the short
    // one, so it is the one kept as evidence and the pile is the keeper.
    const plan = planReconcile([loose(), grouped({ groupChildren: 30 })], [item()], opts);
    expect(plan.verdicts[0].index).toBe(4);
  });

  it("control: two loose twins keep the earlier one", () => {
    const plan = planReconcile([loose(), slide(5)], [item()], opts);
    expect(plan.verdicts[0].index).toBe(4);
    expect(deletes(plan)).toEqual([5]);
  });

  it("control: an unstamped twin still outranks a stamped one", () => {
    const plan = planReconcile([loose({ stamped: true, shapes: 37 }), slide(5)], [item()], opts);
    expect(plan.verdicts[0].index).toBe(5);
  });
});

// ───────────────────────────────────────────────────────────────────────────
// r1-43 / r2-43 — the agenda slide title and the row font
// ───────────────────────────────────────────────────────────────────────────
describe("r1-43: the agenda slide title is fitted to its box", () => {
  const titleOf = (title: string, size: { width?: number } = {}) =>
    buildAgendaScene(["One", "Two"], { title, ...size }).nodes.find((n): n is TextNode => n.name === "agenda-title")!;
  const inkRight = (t: TextNode) => t.x + textWidth(t.text, t.fontSize, t.bold);

  it("a long title stays inside its box instead of running off the slide", () => {
    const t = titleOf("Quarterly business review: strategy, operations and finance");
    expect(inkRight(t)).toBeLessThanOrEqual(t.x + t.w + 0.01);
    expect(inkRight(t)).toBeLessThanOrEqual(SLIDE.width);
  });

  it("shrinks before it cuts: a title that fits at a smaller size keeps every word", () => {
    const title = "Quarterly business review: strategy, operations and finance";
    const t = titleOf(title);
    expect(t.text).toBe(title);
    expect(t.fontSize).toBeLessThan(28);
    expect(t.fontSize).toBeGreaterThanOrEqual(14);
  });

  it("a title too long for the smallest size is ellipsized, still inside the box", () => {
    const t = titleOf("Quarterly business review ".repeat(12));
    expect(t.text.endsWith("…")).toBe(true);
    expect(inkRight(t)).toBeLessThanOrEqual(t.x + t.w + 0.01);
  });

  it("fits a title on a narrower slide too", () => {
    const t = titleOf("Quarterly business review: strategy and operations", { width: 640 });
    expect(inkRight(t)).toBeLessThanOrEqual(t.x + t.w + 0.01);
  });

  it("control: a title that fits is untouched, at the full 28pt", () => {
    const t = titleOf("Agenda for the quarter");
    expect(t.text).toBe("Agenda for the quarter");
    expect(t.fontSize).toBe(28);
    expect(buildAgendaScene(["One"]).nodes.find((n): n is TextNode => n.name === "agenda-title")!.text).toBe("Agenda");
  });
});

describe("r2-43: every agenda slide in a set draws its list at one size", () => {
  const chapters = (len: number) => ["Intro", "X".repeat(len), "End"];
  const rows = (cs: string[], highlight?: number) => {
    const scene = buildAgendaScene(cs, highlight === undefined ? {} : { highlight });
    return scene.nodes.filter((n): n is TextNode => /^agenda-item-\d+$/.test(n.name ?? ""));
  };

  // 71–76 characters is where only the highlighted slide used to shrink; 91 is
  // the verifier's own example, where every slide shrank and the highlighted
  // one shrank further; 120 is deep in the same range.
  for (const len of [72, 75, 91, 120]) {
    it(`a ${len}-character chapter: the pane's insert makes three slides that agree`, () => {
      const cs = chapters(len);
      const sizes = [0, 1, 2].map((i) => rows(cs, i)[0].fontSize);
      expect(new Set(sizes).size, `font sizes per slide: ${sizes.join(", ")}`).toBe(1);
      const xs = [0, 1, 2].map((i) => rows(cs, i)[0].x);
      expect(new Set(xs).size, `chapter text x per slide: ${xs.join(", ")}`).toBe(1);
      // …and so does the overview slide with nothing highlighted.
      expect(rows(cs)[0].fontSize).toBe(sizes[0]);
    });

    it(`a ${len}-character chapter still fits on the slide when it is the highlighted one`, () => {
      const cs = chapters(len);
      const hl = rows(cs, 1)[1];
      expect(hl.bold).toBe(true);
      expect(hl.x + textWidth(hl.text, hl.fontSize, hl.bold)).toBeLessThanOrEqual(hl.x + hl.w + 0.01);
    });
  }

  it("control: chapters that already fit keep the full-size row font on every slide", () => {
    for (const i of [0, 1, 2]) expect(rows(["Intro", "Market", "Strategy"], i)[0].fontSize).toBe(18);
  });
});
