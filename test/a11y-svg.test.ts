import { describe, expect, it } from "vitest";
import { buildChart, describeChart } from "../src/core/chart";
import { buildCheckbox, buildHarveyBall, buildKpiTile, buildProcessFlow, buildTableScene } from "../src/core/elements";
import type { TextNode } from "../src/core/scene";
import { sceneToSvg } from "../src/render/svg";
import type { ChartConfig } from "../src/core/types";

const base: ChartConfig = {
  kind: "clustered",
  width: 480,
  height: 300,
  title: "Revenue by region",
  data: {
    categories: ["North", "South", "East"],
    series: [
      { name: "2024", values: [10, 20, 30] },
      { name: "2025", values: [12, 18, 33] },
    ],
  },
};

describe("SVG accessibility", () => {
  it("marks the root as an image and gives it a title + description", () => {
    const svg = sceneToSvg(buildChart(base));
    expect(svg).toContain('role="img"');
    expect(svg).toContain("<title>Revenue by region</title>");
    expect(svg).toContain("<desc>");
    // title/desc must be the FIRST children of the root for the SVG a11y mapping
    expect(svg.indexOf("<title>")).toBeLessThan(svg.indexOf("<rect"));
    expect(svg.indexOf("<title>")).toBeLessThan(svg.indexOf("<desc>"));
  });

  it("escapes markup in the title", () => {
    const svg = sceneToSvg(buildChart({ ...base, title: "A & B <chart>" }));
    expect(svg).toContain("<title>A &amp; B &lt;chart&gt;</title>");
  });

  it("names an untitled chart with its generated summary", () => {
    // role="img" with a <desc> but no <title> has a description and no NAME —
    // the axe-core role-img-alt failure. The summary becomes the name instead,
    // and is not then repeated as the description.
    const cfg = { ...base, title: undefined };
    const svg = sceneToSvg(buildChart(cfg));
    expect(svg).toContain('role="img"');
    expect(svg).toContain(`<title>${describeChart(cfg)}</title>`);
    expect(svg).not.toContain("<desc>");
    expect(svg.indexOf("<title>")).toBeLessThan(svg.indexOf("<rect"));
  });

  it("describeChart names the kind, series and categories", () => {
    const d = describeChart(base);
    expect(d).toContain("clustered column chart");
    expect(d).toContain("2 data series");
    expect(d).toContain("2024, 2025");
    expect(d).toContain("3 categories");
    expect(d).toContain("North, South, East");
  });

  it("describeChart notes horizontal orientation and truncates long lists", () => {
    expect(describeChart({ ...base, horizontal: true })).toContain("(horizontal)");
    const many = describeChart({
      ...base,
      data: { categories: ["a", "b", "c", "d", "e", "f"], series: [{ name: "S", values: [1, 2, 3, 4, 5, 6] }] },
    });
    expect(many).toContain("and 2 more");
  });
});

/**
 * The description has to be of the chart that is DRAWN.
 *
 * `buildChart` described the config as it arrived, before the transforms that
 * decide what a reader sees: `sortCategories` and `applyPareto` reorder the
 * categories, `applyPareto` also turns a clustered chart into a combo and adds
 * the cumulative line, and `collapseOther` buckets the tail into "Other".
 *
 * So a pareto announced itself as a "clustered column chart. 1 data series: V.
 * 4 categories: C0, C1, C2, C3" while drawing a combo of two series with its
 * categories in the order C1, C2, C0, C3. The one reader who cannot check the
 * description against the picture got the wrong chart, the wrong series count
 * and the wrong order — and this text is not only the SVG's `<desc>`, it is the
 * alt text written onto every shape in the .pptx.
 */
describe("the accessible description matches the chart that was drawn", () => {
  const cats = ["C0", "C1", "C2", "C3"];
  const base = {
    width: 480,
    height: 300,
    data: { categories: cats, series: [{ name: "V", values: [10, 40, 25, 5] }] },
  };
  const drawnCategories = (scene: { nodes: { name?: string; kind: string }[] }) =>
    scene.nodes
      .filter((n) => n.kind === "text" && /^category-\d+$/.test(n.name ?? ""))
      .map((n) => (n as TextNode).text);

  it("lists the categories in the order they are plotted", () => {
    for (const extra of [{ categorySort: "descending" }, { pareto: true }] as Partial<ChartConfig>[]) {
      const scene = buildChart({ ...base, kind: "clustered", ...extra } as ChartConfig);
      const order = drawnCategories(scene);
      expect(order.length, "no category labels drawn, so this proves nothing").toBe(4);
      expect(scene.desc, JSON.stringify(extra)).toContain(order.join(", "));
    }
  });

  it("names the kind and the series a pareto actually becomes", () => {
    const scene = buildChart({ ...base, kind: "clustered", pareto: true } as ChartConfig);
    expect(scene.desc).toMatch(/^combination chart/);
    expect(scene.desc).toContain("2 data series");
  });

  it("leaves an untransformed chart's description exactly as it was", () => {
    // The negative control: no sort, no pareto, no bucket — nothing may move.
    const scene = buildChart({ ...base, kind: "clustered" } as ChartConfig);
    expect(scene.desc).toBe("clustered column chart. 1 data series: V. 4 categories: C0, C1, C2, C3.");
  });
});

/**
 * THE SAME RULE, FOR THE ELEMENTS — written for charts, and the element family
 * never got it.
 *
 * `sceneToSvg` stamps `role="img"` on EVERY scene, and its own comment says what
 * that costs when there is no name: "role=img with no name is exactly what
 * axe-core's role-img-alt reports as a 1.1.1 failure". `buildChart` always set a
 * `desc`; the five builders in `src/core/elements.ts` set neither field, so the
 * four Elements-tab previews shipped to every user as unnamed images. Found by
 * running axe-core against the LIVE pane on 2026-09-10, not by reading the code
 * — every test above passed the whole time it was true, because every one of
 * them asks about a chart.
 *
 * The repo's most-repeated defect wearing a new hat: the fix reached all but one
 * call site. So this iterates the family rather than naming a favourite.
 */
describe("every element scene carries a text alternative", () => {
  const scenes: [string, ReturnType<typeof buildHarveyBall>][] = [
    ["harvey ball", buildHarveyBall(0.75)],
    ["checkbox", buildCheckbox("yes")],
    ["process flow", buildProcessFlow(["Scope", "Design", "Build"], 1)],
    ["kpi tile", buildKpiTile({ label: "Revenue", value: "€4.2m", delta: "+12% vs LY" })],
    [
      "table",
      buildTableScene([
        ["a", "b"],
        ["1", "2"],
      ]),
    ],
    ["empty table", buildTableScene([])],
  ];

  it.each(scenes)("%s is named, so its role=img is not anonymous", (_name, scene) => {
    expect(scene.desc, "no text alternative — this is the axe-core role-img-alt failure").toBeTruthy();
    const svg = sceneToSvg(scene);
    expect(svg).toContain('role="img"');
    // `sceneToSvg` promotes a desc-only scene's description to the NAME, which
    // is the point: a description without a name still fails 1.1.1.
    expect(svg, "the description did not become the accessible name").toContain(`<title>${scene.desc}</title>`);
  });

  it("describes what was drawn rather than what was asked for", () => {
    // Each builder sanitises its input, and the text alternative has to agree
    // with the sanitised value or it describes a graphic nobody drew.
    expect(buildCheckbox("__proto__" as never).desc, "a bogus state draws the neutral glyph").toBe(
      "Checkbox, partial.",
    );
    expect(buildProcessFlow(null as never).desc, "a non-array draws an empty flow").toBe("Process flow with no steps.");
    expect(buildProcessFlow(["A", "B"], 9).desc, "a highlight past the end names no step").toBe("Process flow: A, B.");
    expect(buildHarveyBall(2).desc, "a fraction above 1 is clamped before it is drawn").toBe(
      "Harvey ball, 100% filled.",
    );
    expect(buildTableScene([["a", "b", "c"], ["1"]]).desc, "a ragged paste is drawn to its widest row").toBe(
      "Table, 2 rows by 3 columns.",
    );
  });
});
