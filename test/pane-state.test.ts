// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import type { ChartConfig } from "../src/core/types";
import { buildChart, clampDim } from "../src/core/chart";
import { placeChart } from "../src/core/placement";

/**
 * Task-pane state tests. app.ts is a side-effecting entry module: it wires
 * itself to taskpane.html on import and takes the non-Office branch when the
 * `Office` global is absent, which is exactly the case under jsdom. Boot it
 * against the real markup and drive it the way a user does — through the JSON
 * import/export box — so "load a chart" is covered end to end.
 */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

async function bootPane(search = "") {
  // A CLEAN browser, not just a clean DOM.
  //
  // The pane reads `localStorage` on boot to find a run that never reported
  // finishing, and says so. Earlier tests in this file click "Insert demo
  // deck", which starts recording one — so without this, whether a later test
  // sees an empty status strip depends on how far the previous test's stubbed
  // host got before the next boot. It passed locally and failed in CI, which is
  // exactly what an ordering dependency looks like. A test that asserts the
  // pane has "nothing to say" has to own everything that could speak.
  try {
    window.localStorage.clear();
  } catch {
    /* a jsdom without storage — nothing to carry over either */
  }
  // app.ts reads ?lang= at import time, so the URL has to be set up front.
  window.history.replaceState({}, "", `/taskpane.html${search}`);
  // Parse rather than regex out the <script> tags: the office.js tag has no
  // business loading here, and a regex that thinks it can find "</script>"
  // misses "</script >" (CodeQL js/bad-tag-filter).
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  // app.ts holds module-level state and element handles, so it has to re-run
  // against each fresh DOM. Without this the cached module keeps listening to
  // the previous test's detached nodes and every click silently does nothing.
  vi.resetModules();
  await import("../src/taskpane/app");
}

/** Load `cfg` through the pane's JSON import box, as a user would. */
const importConfig = (cfg: Partial<ChartConfig>) => {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(cfg);
  $("json-import").click();
};

/** Read the pane's current config back out of the JSON export box. */
const exportConfig = (): ChartConfig => {
  $("json-export").click();
  return JSON.parse(($("json-io") as HTMLTextAreaElement).value);
};

const cell = (r: number, c: number) =>
  document.querySelector<HTMLInputElement>(`#datasheet input[data-row="${r}"][data-col="${c}"]`)!;

const type = (r: number, c: number, value: string) => {
  const input = cell(r, c);
  input.value = value;
  input.dispatchEvent(new Event("input"));
};

const baseData = { categories: ["A", "B"], series: [{ name: "S1", values: [1, 2] }] };

describe("task pane — loading a chart config", () => {
  beforeEach(async () => {
    await bootPane();
  });

  it("syncs the size fields, so the chart keeps its own dimensions", () => {
    importConfig({ kind: "clustered", width: 720, height: 400, data: baseData });
    // currentConfig() reads the size straight off these inputs, so a stale
    // field silently resized every loaded chart back to 480x300.
    expect(($("chart-w") as HTMLInputElement).value).toBe("720");
    expect(($("chart-h") as HTMLInputElement).value).toBe("400");
    const out = exportConfig();
    expect([out.width, out.height]).toEqual([720, 400]);
  });

  it("writes edited size inputs back into the exported config", () => {
    importConfig({ kind: "clustered", width: 480, height: 300, data: baseData });
    const w = $("chart-w") as HTMLInputElement;
    w.value = "640";
    w.dispatchEvent(new Event("input"));
    expect(exportConfig().width).toBe(640);
    // A sub-usable value is ignored — the last good size holds.
    w.value = "5";
    w.dispatchEvent(new Event("input"));
    expect(exportConfig().width).toBe(640);
  });

  it("keeps a loaded chart's style rather than falling back to pane defaults", () => {
    const style = {
      fontFamily: "Georgia",
      fontSize: 14,
      negative: "#b00020",
      neutral: "#8a8a8a",
      palette: ["#111111", "#222222", "#333333"],
    };
    importConfig({ kind: "clustered", data: baseData, style });
    expect(exportConfig().style).toMatchObject(style);
  });

  it("preserves series fields the datasheet can't carry (type / pattern / colors)", () => {
    // The pane rebuilds ChartConfig from the sheet on every export/insert, and the
    // sheet only holds name+values. Without a side-channel these three fields are
    // silently dropped on import and destroyed on re-save.
    importConfig({
      kind: "combo",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Rev", values: [100, 120] },
          {
            name: "Margin",
            values: [30, 40],
            type: "line",
            pattern: "diagonal",
            colors: ["#ff0000", null],
            scenario: "FC",
          },
        ],
      },
    });
    const margin = exportConfig().data.series.find((s) => s.name === "Margin")!;
    expect(margin.type).toBe("line");
    expect(margin.pattern).toBe("diagonal");
    expect(margin.colors).toEqual(["#ff0000", null]);
    expect(margin.scenario).toBe("FC"); // IBCS scenario survives the sheet round-trip too
  });

  it("keeps those series fields through a datasheet edit (round-trip on rebuild)", () => {
    importConfig({
      kind: "combo",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Rev", values: [100, 120] },
          { name: "Margin", values: [30, 40], type: "line" },
        ],
      },
    });
    type(1, 1, "150"); // edit a cell → currentConfig rebuilds from the sheet
    expect(exportConfig().data.series.find((s) => s.name === "Margin")!.type).toBe("line");
  });

  it("preserves top-level chart features the pane has no control for (round-trip)", () => {
    // The pane rebuilds ChartConfig from a handful of state fields on every
    // export/insert. Keys with no matching control — radar/scatter/gantt modes,
    // secondaryAxis, categorySort, otherBucket, labelOffsets, multiples/butterfly/
    // tilemap options — used to be dropped on import and destroyed on re-save.
    importConfig({
      kind: "radar",
      data: baseData,
      radar: { perSpoke: true },
      secondaryAxis: true,
      categorySort: "descending",
      otherBucket: { max: 3 },
      labelOffsets: { "S1@0": { dx: 4, dy: -2 } },
    });
    const out = exportConfig();
    expect(out.radar).toEqual({ perSpoke: true });
    expect(out.secondaryAxis).toBe(true);
    expect(out.categorySort).toBe("descending");
    expect(out.otherBucket).toEqual({ max: 3 });
    expect(out.labelOffsets).toEqual({ "S1@0": { dx: 4, dy: -2 } });
  });

  /**
   * The same question asked of EVERY key, from the interface rather than from a
   * list somebody remembered to extend.
   *
   * CLAUDE.md states this seam and its failure mode: "new decoration keys
   * round-trip automatically; new top-level config keys need a state field or
   * the `state.extras` passthrough". Miss that and the key is dropped on import
   * and destroyed on the next re-save — silently, in a pane that otherwise looks
   * like it loaded the chart fine. The case above pins five keys and was written
   * the day five keys were found missing; it cannot say anything about the
   * sixth.
   *
   * `TOP_LEVEL` is parsed out of `ChartConfig` for the reason `DECOR_KEYS` in
   * `chart-hostile-input.test.ts` is: a hand-written list of names is a list
   * that goes stale the first time somebody adds a field and does not think of
   * this file.
   */
  const TOP_LEVEL = (() => {
    const src = readFileSync("src/core/types.ts", "utf8");
    const body = /export interface ChartConfig \{(.*?)\n\}/s.exec(src)?.[1] ?? "";
    return [...body.matchAll(/^ {2}([a-zA-Z_$][\w$]*)\??\s*:/gm)].map((m) => m[1]);
  })();

  /**
   * A plausible non-default value per key. Written out rather than generated,
   * because "survives the round trip" is only meaningful for a value the pane
   * would actually accept — a generated one would prove the key is echoed, not
   * that it is understood.
   *
   * `kind`, `data` and `style` are excluded and NAMED: they are what the pane's
   * own controls are made of, they are asserted all over this file already, and
   * a generic round-trip on them would be asserting that a select and a grid
   * echo themselves.
   */
  const OWNED_BY_CONTROLS = ["kind", "data", "style"];
  const SAMPLES: Record<string, unknown> = {
    horizontal: true,
    scale: { min: -5, max: 55 },
    segmentOrder: "descending",
    categorySort: "descending",
    secondaryAxis: true,
    axisBreak: { from: 10, to: 20 },
    valueAxisTitle: "EUR m",
    labelOffsets: { "S1@0": { dx: 4, dy: -2 } },
    logScale: true,
    gapWidth: 0.42,
    overlap: 0.25,
    footnote: "Source: invented",
    pie: { donut: 0.4 },
    pareto: { cumulative: true },
    multiples: { columns: 3 },
    boxplot: { whiskers: "minmax" },
    map: { region: "world" },
    heatmap: { scheme: "blues" },
    otherBucket: { max: 3 },
    tilemap: { shape: "hex" },
    butterfly: { gap: 40 },
    scatter: { markers: ["diamond"] },
    gantt: { today: 5 },
    radar: { perSpoke: true },
    combo: { types: ["column", "line"] },
    width: 720,
    height: 400,
    title: "A title",
    decorations: { totals: true },
    waterfall: { totalIndices: [1] },
    numberFormat: { decimals: 2, suffix: "m" },
    labels: { series: ["X"] },
    render: "image",
  };

  it("every key in ChartConfig has a sample here — a new one is not silently skipped", () => {
    // The half that makes the sweep below exhaustive rather than merely long.
    const unsampled = TOP_LEVEL.filter((k) => !OWNED_BY_CONTROLS.includes(k) && !(k in SAMPLES));
    expect(
      unsampled,
      "a ChartConfig key has no sample, so nothing checks whether the pane keeps it — add one to SAMPLES, " +
        "or to OWNED_BY_CONTROLS if a pane control owns it outright",
    ).toEqual([]);
    expect(TOP_LEVEL.length, "no ChartConfig keys parsed — the sweep would test nothing").toBeGreaterThan(30);
  });

  it("carries every top-level key through import → export", () => {
    const lost: string[] = [];
    const changed: string[] = [];
    for (const key of TOP_LEVEL) {
      if (OWNED_BY_CONTROLS.includes(key)) continue;
      const value = SAMPLES[key];
      // `waterfall` is meaningful only on a waterfall — its total columns are
      // rebuilt from the sheet, so asking a clustered chart to keep them is
      // asking the wrong question.
      const kind = key === "waterfall" ? "waterfall" : "clustered";
      importConfig({ kind, data: baseData, [key]: value } as Partial<ChartConfig>);
      const got = (exportConfig() as unknown as Record<string, unknown>)[key];
      if (got === undefined) lost.push(key);
      else if (JSON.stringify(got) !== JSON.stringify(value))
        changed.push(`${key}: sent ${JSON.stringify(value)}, got ${JSON.stringify(got)}`);
    }
    expect(lost, "dropped on import — the next re-save destroys them").toEqual([]);
    expect(changed, "altered on the way through").toEqual([]);
    // 2.2 s idle, measured 2026-09-13, and the slowest ordinary test in the
    // suite — it boots the real pane module and walks every top-level key. It
    // is the number the 20s default in `vitest.config.ts` was sized against.
  });

  /**
   * The sweep above asks `waterfall` only of a WATERFALL, on the reasoning that
   * its total columns are rebuilt from the sheet — and that reasoning has a
   * hole. `combo` with `combo.columns: "waterfall"` reads the same field for its
   * bridge's column bases, and the pane wrote `totalIndices` from the sheet
   * unconditionally: the "e" tokens that fill it are only seeded for the
   * waterfall kind, so every other kind had its list replaced with an EMPTY one.
   *
   * Slide 90 of the committed showcase is exactly that chart, and it was the
   * only one of the 123 that did not round-trip through the pane. Opening it —
   * from its POWERCHART_CONFIG tag, the JSON box or a template — and inserting
   * again dropped the closing total column and the label on it.
   */
  it("keeps a combo's waterfall totals, which are not rebuilt from its sheet", () => {
    const cfg = {
      kind: "combo",
      combo: { columns: "waterfall" },
      waterfall: { totalIndices: [5] },
      data: {
        categories: ["Start", "Q1", "Q2", "Q3", "Q4", "End"],
        series: [
          { name: "Delta", values: [100, 20, -15, 25, -10, 0] },
          { name: "Margin %", type: "line", values: [40, 42, 38, 44, 41, 43] },
        ],
      },
    } as unknown as Partial<ChartConfig>;
    importConfig(cfg);
    const out = exportConfig();
    expect(out.waterfall?.totalIndices).toEqual([5]);
    // …and the bar it pays for is still drawn.
    const bars = buildChart(out).nodes.filter((n) => /^bar-/.test(n.name ?? ""));
    expect(bars.map((n) => n.name)).toContain("bar-5");
  });

  it('preserves render: "image" through import → export and the shape-tag re-save', () => {
    // `render` selects the OUTPUT format (native shapes vs one raster picture).
    // It began life in state.extras (no control) and now has #render-image, but
    // the contract this pins is the same either way: importing an
    // image-mode config and exporting it — or letting the pane re-save it into
    // POWERCHART_CONFIG on an update — silently downgraded the chart to shapes.
    // A silent format change is worse than a refusal: the skill CLI honours the
    // key, so a config that round-tripped through the pane stopped rasterising
    // with nothing to indicate why.
    importConfig({ kind: "stacked", data: baseData, render: "image" });
    expect(exportConfig().render).toBe("image");

    // The default stays absent rather than becoming an explicit "shapes" — the
    // key is optional and the engine treats undefined as shapes, so writing it
    // out would add noise to every exported config.
    importConfig({ kind: "stacked", data: baseData });
    expect(exportConfig().render).toBeUndefined();

    // Survives a datasheet edit, which rebuilds the config from sheet state.
    importConfig({ kind: "stacked", data: baseData, render: "image" });
    type(1, 1, "150");
    expect(exportConfig().render).toBe("image");
  });

  it("syncs the #render-image checkbox both ways with the config", () => {
    // The checkbox now OWNS `render` (it left state.extras when it got a
    // control), so the two directions are separate failure modes: a loaded
    // image config that leaves the box unticked lies to the user about what the
    // next insert will do, and a ticked box that doesn't reach the config makes
    // the control inert.
    const box = () => $("render-image") as HTMLInputElement;

    importConfig({ kind: "stacked", data: baseData, render: "image" });
    expect(box().checked, "loading an image config ticks the box").toBe(true);

    importConfig({ kind: "stacked", data: baseData });
    expect(box().checked, "loading a shapes config unticks it").toBe(false);

    // Ticking it by hand reaches the exported config.
    box().checked = true;
    box().dispatchEvent(new Event("change"));
    expect(exportConfig().render).toBe("image");

    box().checked = false;
    box().dispatchEvent(new Event("change"));
    expect(exportConfig().render).toBeUndefined();
  });

  it("merges pane-owned fields with import-only ones for pie/waterfall/numberFormat", () => {
    // These three are split: the control owns explode / total "e" tokens /
    // decimals+suffix+locale, but semi/breakout/variableRadius, detailGroups/
    // spacerIndices and forceSign live only in the imported config. A naive
    // rebuild let the control's half clobber the whole object.
    importConfig({
      kind: "pie",
      data: baseData,
      pie: { semi: true, breakout: [1], explode: [0] },
      numberFormat: { decimals: 1, forceSign: true },
    });
    const pieOut = exportConfig();
    // The pane's explode survives AND the import-only pie fields do too.
    expect(pieOut.pie).toMatchObject({ semi: true, breakout: [1], explode: [0] });
    expect(pieOut.numberFormat).toMatchObject({ decimals: 1, forceSign: true });

    importConfig({
      kind: "waterfall",
      data: { categories: ["A", "B", "C"], series: [{ name: "S1", values: [10, -4, 6] }] },
      waterfall: { detailGroups: [{ of: 0, indices: [1, 2] }], spacerIndices: [1] },
    });
    const wfOut = exportConfig();
    expect(wfOut.waterfall).toMatchObject({
      detailGroups: [{ of: 0, indices: [1, 2] }],
      spacerIndices: [1],
    });
    expect(Array.isArray(wfOut.waterfall!.totalIndices)).toBe(true);
  });

  it("keeps import-only features through a datasheet edit (rebuild from sheet)", () => {
    importConfig({ kind: "radar", data: baseData, radar: { perSpoke: true }, secondaryAxis: true });
    type(1, 1, "42"); // edit a cell → currentConfig rebuilds
    const out = exportConfig();
    expect(out.radar).toEqual({ perSpoke: true });
    expect(out.secondaryAxis).toBe(true);
  });

  it("does not undo into the previous chart's data", () => {
    importConfig({ kind: "clustered", data: baseData });
    type(1, 1, "99"); // an edit, so there is history to undo
    importConfig({
      kind: "clustered",
      data: { categories: ["X", "Y"], series: [{ name: "T1", values: [7, 8] }] },
    });
    const loaded = exportConfig();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true }));
    // Ctrl+Z used to replay the *previous* chart's cells into the new sheet,
    // leaving a chart whose data belonged to neither.
    expect(exportConfig().data).toEqual(loaded.data);
  });

  it("undoes the colours and combo types with the cells, not just the cells", () => {
    // `seriesColors`/`seriesMeta` are positional side-channels kept in step only
    // by the structure handler, which fires on the EDIT and has no inverse. Undo
    // rewound the grid and left them spliced, so deleting the middle row of a
    // three-series combo and pressing Ctrl+Z brought the row back wearing the
    // colour and type of the series that had followed it — a chart that now says
    // something different, from the one keystroke whose whole promise is that it
    // changes nothing.
    const combo = {
      kind: "clustered",
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "Revenue", values: [10, 20], color: "#ff0000" },
          { name: "Margin", values: [3, 4], color: "#00ff00", type: "line" },
          { name: "Costs", values: [7, 8], color: "#0000ff", pattern: "hatch" },
        ],
      },
    };
    importConfig(combo as never);
    const before = exportConfig().data.series;
    // Delete the middle row through the grid's own button, so the structure
    // handler runs exactly as it does for a user.
    const rows = document.querySelectorAll<HTMLElement>("#datasheet tr");
    (rows[2].querySelector("td, th") as HTMLElement).click();
    [...document.querySelectorAll<HTMLButtonElement>("#datasheet button")]
      .find((b) => b.textContent?.includes("− Row"))!
      .click();
    expect(exportConfig().data.series, "the row was not removed, so this proves nothing").toHaveLength(2);
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true }));
    expect(exportConfig().data.series, "undo did not restore the row").toHaveLength(3);
    expect(exportConfig().data.series).toEqual(before);
  });

  it("keeps a renamed series' colour, combo type, pattern and scenario", () => {
    // Renaming a row is the datasheet's core edit. The side-channel used to be
    // keyed by series NAME, so the new name matched nothing and the overlay line
    // collapsed back into a plain column — silently, on one keystroke.
    importConfig({
      kind: "combo",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Rev", values: [10, 12] },
          { name: "Margin %", values: [3, 4], type: "line", color: "#ff0000", pattern: "dots", scenario: "FC" },
        ],
      },
    });
    type(2, 0, "Margin"); // the NAME cell of the second series row
    expect(exportConfig().data.series[1]).toMatchObject({
      name: "Margin",
      values: [3, 4],
      color: "#ff0000",
      type: "line",
      pattern: "dots",
      scenario: "FC",
    });
  });

  /**
   * seriesColors / seriesMeta are POSITIONAL. The grid's row buttons spliced
   * rows in and out of the sheet without touching them, so every series below
   * the edit inherited its neighbour's colour and combo type — a routine
   * datasheet edit silently changed what the chart drew.
   */
  describe("a row or column edit carries the per-series side-channel with it", () => {
    const sheetButton = (label: string) =>
      [...document.querySelectorAll<HTMLButtonElement>("#datasheet .sheet-controls button")]
        .find((b) => b.textContent === label)!
        .click();

    const threeSeries = () =>
      importConfig({
        kind: "combo",
        data: {
          categories: ["A", "B"],
          series: [
            { name: "One", values: [1, 2], color: "#111111" },
            { name: "Two", values: [3, 4], color: "#222222" },
            { name: "Three", values: [5, 6], color: "#333333", type: "line" },
          ],
        },
      });

    it("keeps each colour and combo type on its own series after a row delete", () => {
      threeSeries();
      cell(1, 1).focus(); // the "One" row
      sheetButton("− Row");
      const series = exportConfig().data.series;
      expect(series.map((s) => s.name)).toEqual(["Two", "Three"]);
      expect(series.map((s) => s.color)).toEqual(["#222222", "#333333"]);
      expect(series[1].type).toBe("line");
    });

    it("keeps them after a row insert", () => {
      threeSeries();
      cell(1, 1).focus();
      sheetButton("+ Row"); // a new empty series between One and Two
      const series = exportConfig().data.series;
      expect(series.map((s) => s.color)).toEqual(["#111111", undefined, "#222222", "#333333"]);
      // The overlay line is still the LAST series, not the new empty one.
      expect(series.filter((s) => s.type === "line").map((s) => s.name)).toEqual(["Three"]);
    });

    it("drops the mapping on a transpose, where nothing positional survives", () => {
      threeSeries();
      sheetButton("⇄ Transpose");
      expect(exportConfig().data.series.some((s) => s.type === "line")).toBe(false);
    });
  });

  it("keeps two same-named series apart", () => {
    // The same name key collapsed both rows onto one entry, so they rendered in
    // a single colour.
    importConfig({
      kind: "clustered",
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [1], color: "#ff0000" },
          { name: "S", values: [2], color: "#00ff00" },
        ],
      },
    });
    expect(exportConfig().data.series.map((s) => s.color)).toEqual(["#ff0000", "#00ff00"]);
  });

  it("keeps the CAGR arrow's series anchor when the from/to spinner moves", () => {
    // pairControl only knows from/to, so touching a spinner dropped
    // decorations.cagr.series — and core/decor.ts then measured the column
    // TOTALS instead of the anchored series, printing a different growth rate.
    importConfig({
      kind: "clustered",
      data: {
        categories: ["FY21", "FY22", "FY23"],
        series: [
          { name: "A", values: [10, 20, 40] },
          { name: "B", values: [100, 100, 100] },
        ],
      },
      decorations: { cagr: { from: 0, to: 2, series: 0 } },
    });
    const label = [...document.querySelectorAll<HTMLElement>("#options label")].find((l) =>
      l.textContent?.includes("CAGR arrow"),
    )!;
    const to = label.querySelectorAll<HTMLInputElement>("input[type=number]")[1];
    to.value = "2"; // 1-based in the UI → to: 1
    to.dispatchEvent(new Event("input"));
    expect(exportConfig().decorations!.cagr).toEqual({ from: 0, to: 1, series: 0 });
  });

  it("keeps a decimals value the select has no option for when the suffix is edited", () => {
    // The <select> offers auto/0/1/2 only. A loaded decimals:3 left it at value
    // "", and emitNf — which writes BOTH controls — turned that into 0 the
    // moment the suffix box was touched, so every label lost its decimals.
    importConfig({ kind: "clustered", data: baseData, numberFormat: { decimals: 3, forceSign: true } });
    const suffix = [...document.querySelectorAll<HTMLInputElement>("#options input[type=text]")].find(
      (i) => i.placeholder === "e.g. €m",
    )!;
    suffix.value = "%";
    suffix.dispatchEvent(new Event("input"));
    expect(exportConfig().numberFormat).toMatchObject({ decimals: 3, suffix: "%", forceSign: true });
  });
});

describe("task pane — action label survives re-renders", () => {
  it("keeps the insert button label correct after a chart loads", async () => {
    await bootPane();
    const insert = () => $("insert").textContent;
    expect(insert()).toBe("Insert into slide");

    // renderActionState rewrites this label whenever the edit target changes
    // (then re-applies the active language via localizeTree). Loading a fresh,
    // non-edit chart must leave it as the insert label — English-only today.
    importConfig({ kind: "clustered", data: baseData });
    expect(insert()).toBe("Insert into slide");
  });
});

describe("task pane — status colour and headings", () => {
  beforeEach(async () => {
    await bootPane();
  });

  it("shows a failure in the error colour, not the previous success's green", () => {
    const note = () => $("host-note");
    // Stand in for a prior successful action having stamped the green class.
    note().className = "hint status-ok";
    ($("json-io") as HTMLTextAreaElement).value = "{ not json";
    $("json-import").click();
    expect(note().textContent).toMatch(/^Invalid JSON:/);
    expect(note().className).toBe("hint status-err");
  });

  it("does not call valid JSON a syntax error", () => {
    /**
     * One `try` used to cover both `JSON.parse` and `applyConfig`, so a config
     * that PARSES and then breaks ingest was reported as "Invalid JSON:
     * data.series.some is not a function" — sending the user to hunt a syntax
     * error that is not there, in text any validator calls fine.
     *
     * `stateFromConfig`'s own comment records this symptom, and `asArray` made
     * the common shapes survivable without making the MESSAGE honest. These two
     * are shapes it does not cover.
     */
    for (const valid of [
      { kind: "clustered", data: null },
      { kind: "clustered", data: { categories: "A", series: "x" } },
    ]) {
      ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(valid);
      $("json-import").click();
      const said = String($("host-note").textContent);
      expect(said, `called valid JSON a syntax error: ${said}`).not.toMatch(/^Invalid JSON:/);
      expect(said, `said nothing useful instead: ${said}`).toMatch(/valid JSON, but not a chart/);
      expect($("host-note").className, "reported a failure as a success").toBe("hint status-err");
    }
  });

  it("still calls a real syntax error one", () => {
    // The other half: separating the two must not stop the pane naming an
    // actual syntax error, which is what that message was always right about.
    ($("json-io") as HTMLTextAreaElement).value = "{ not json";
    $("json-import").click();
    expect(String($("host-note").textContent)).toMatch(/^Invalid JSON:/);
  });

  it("labels a successful load as such", () => {
    importConfig({ kind: "clustered", data: baseData });
    expect($("host-note").className).toBe("hint status-ok");
  });
});

describe("task pane — accordion step headings", () => {
  it("renders the numbered step headings", async () => {
    await bootPane();
    const titles = [...document.querySelectorAll(".acc-title")].map((e) => e.textContent);
    // These live in a <span> inside the <summary> — the pane renders them; a
    // future language re-skins them via the .acc-title selector (see i18n.ts).
    expect(titles).toContain("1 · Chart type");
    expect(titles).toContain("3 · Decorations");
    expect(titles).toContain("Preview & size");
  });
});

/**
 * Boot the pane down its HOST branch, with a PowerPoint.run we control: it
 * parks until we release it, so we can look at the buttons mid-flight — which
 * is the only moment the bug existed.
 */
async function bootHost() {
  let release!: () => void;
  const parked = new Promise<void>((r) => (release = r));
  let runs = 0;
  vi.stubGlobal("Office", {
    context: { host: "PowerPoint", requirements: { isSetSupported: () => false } },
  });
  vi.stubGlobal("PowerPoint", {
    run: async (cb: (ctx: unknown) => Promise<unknown>) => {
      runs++;
      await parked;
      return cb({
        presentation: {
          slides: {
            getCount: () => ({ value: 0 }),
            add() {},
            getItemAt: () => ({ shapes: { addGeometricShape: stubShape, addLine: stubShape, addTextBox: stubShape } }),
          },
        },
        sync: async () => {},
      });
    },
    GeometricShapeType: new Proxy({}, { get: (_t, p) => String(p) }),
    ConnectorType: { straight: "straight" },
    ShapeLineDashStyle: { dash: "dash" },
    ShapeAutoSize: { autoSizeNone: "none" },
    TextVerticalAlignment: { top: "t", middle: "m", bottom: "b" },
    ParagraphHorizontalAlignment: { left: "l", center: "c", right: "r" },
  });
  const stubShape = () => ({
    fill: { setSolidColor() {}, clear() {} },
    lineFormat: {},
    textFrame: { textRange: { font: {}, paragraphFormat: {} } },
    tags: { add() {} },
  });
  await bootPane();
  return { release, runs: () => runs };
}

describe("busy-guard on host actions", () => {
  it("disables the clicked button AND the primary Insert while an action runs", async () => {
    // The bug, seen in real PowerPoint: guard disabled only the primary button,
    // so "Insert demo deck" stayed live through a multi-minute run (one more
    // click = another 35 slides), while Insert went dead WITHOUT looking dead —
    // so a stuck action read as "Insert is broken".
    const { release } = await bootHost();
    const demo = $<HTMLButtonElement>("demo-insert");
    const insert = $<HTMLButtonElement>("insert");
    expect(demo.disabled).toBe(false);
    expect(insert.disabled).toBe(false);

    demo.click();
    await Promise.resolve();
    expect(demo.disabled, "clicked button must not accept a second click").toBe(true);
    expect(insert.disabled, "primary acts on the same deck").toBe(true);

    release();
    await vi.waitFor(() => expect(demo.disabled).toBe(false));
    expect(insert.disabled).toBe(false);
  });

  it("re-enables both when the action fails", async () => {
    vi.stubGlobal("Office", { context: { host: "PowerPoint", requirements: { isSetSupported: () => false } } });
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("host refused");
      },
    });
    await bootPane();
    const demo = $<HTMLButtonElement>("demo-insert");
    demo.click();
    await vi.waitFor(() => expect(document.getElementById("host-note")!.textContent).toMatch(/^Failed:/));
    // A failed action must not leave the pane permanently dead.
    expect(demo.disabled).toBe(false);
    expect($<HTMLButtonElement>("insert").disabled).toBe(false);
  });
});

describe("the action bar belongs to the Chart tab", () => {
  const bar = () => document.querySelector<HTMLElement>(".action-bar")!;
  const clickTab = (name: string) =>
    document.querySelector<HTMLButtonElement>(`.tabs .tab[data-tab="${name}"]`)!.click();

  it("hides itself on every tab that is not Chart", async () => {
    // Every action in this bar reads the CHART's state: "Insert into slide"
    // inserts currentConfig(). On Elements that put a big primary button that
    // inserts a stacked column chart directly under the small "Insert" that
    // inserts the Harvey ball you are looking at — the prominent button was
    // the wrong one.
    await bootPane();
    expect(bar().hasAttribute("hidden"), "Chart tab").toBe(false);
    for (const tab of ["elements", "agenda", "automation"]) {
      clickTab(tab);
      expect(bar().hasAttribute("hidden"), tab).toBe(true);
    }
    clickTab("chart");
    expect(bar().hasAttribute("hidden"), "back on Chart").toBe(false);
  });

  it("honours the tab a deep link opens on", async () => {
    // ?tab=elements clicks the tab after wiring, so the bar must follow the
    // link rather than the markup's default active tab.
    await bootPane("?tab=elements");
    expect(bar().hasAttribute("hidden")).toBe(true);
  });

  it("actually disappears — [hidden] must beat the bar's own display:flex", async () => {
    // The trap: `.action-bar { display: flex }` is an author rule and outranks
    // the UA stylesheet's `[hidden] { display: none }`, so the attribute would
    // be set and the bar would stay on screen. Assert the CSS, not the DOM.
    const css = readFileSync("src/taskpane/taskpane.css", "utf8");
    expect(css).toMatch(/\.action-bar\[hidden\]\s*\{[^}]*display:\s*none/);
  });
});

describe("status is pane-wide, and only claims what it knows", () => {
  const strip = () => document.getElementById("status-strip")!;
  const bar = () => document.getElementById("status-bar")!;
  const noteEl = () => document.getElementById("host-note")!;

  it("lives OUTSIDE the action bar, so hiding that bar cannot silence it", async () => {
    // The regression this exists for: host-note used to live inside
    // <footer class="action-bar">, which is Chart-only. Hiding the bar took
    // every "Working…", "Failed:" and progress count on Elements / Agenda /
    // Automation down with it — including the demo deck's own counter, which
    // is on Automation. Inserting is slow enough here that silence reads as
    // broken.
    await bootPane();
    expect(document.querySelector(".action-bar #host-note"), "note must not be inside the action bar").toBeNull();
    expect(strip().contains(noteEl())).toBe(true);
    // And the strip must not be a child of the thing that gets hidden.
    expect(document.querySelector(".action-bar")!.contains(strip())).toBe(false);
  });

  it("collapses when there is nothing to say", async () => {
    await bootPane();
    expect(strip().hasAttribute("hidden")).toBe(true);
    expect(bar().hasAttribute("hidden")).toBe(true);
  });

  it("shows an INDETERMINATE bar for work whose progress we cannot know", async () => {
    // A single insert is one context.sync(); Office.js reports nothing until it
    // lands. Any percentage would be invented, and a bar stuck at 99% is a
    // worse lie than no bar.
    const { release } = await bootHost();
    $<HTMLButtonElement>("demo-insert").click();
    await Promise.resolve();
    expect(strip().hasAttribute("hidden")).toBe(false);
    expect(bar().hasAttribute("hidden")).toBe(false);
    expect(bar().classList.contains("indeterminate")).toBe(true);
    // Indeterminate means NO width claim.
    expect(bar().querySelector("i")!.style.width).toBe("");
    release();
    await vi.waitFor(() => expect(bar().hasAttribute("hidden")).toBe(true));
  });

  it("counts the seconds while the host works, and stops when it is done", async () => {
    // The only number we can honestly report mid-sync — and on a host that
    // takes 20s to draw a chart, a number that moves is the whole difference
    // between "working" and "dead".
    vi.useFakeTimers();
    try {
      const { release } = await bootHost();
      const elapsed = () => document.getElementById("status-elapsed")!.textContent;
      $<HTMLButtonElement>("demo-insert").click();
      await Promise.resolve();
      expect(elapsed()).toBe("0s");
      await vi.advanceTimersByTimeAsync(3_000);
      expect(elapsed()).toBe("3s");
      release();
      await vi.advanceTimersByTimeAsync(10);
      await vi.waitFor(() => expect(elapsed()).toBe(""));
      // The ticker must not outlive the work.
      await vi.advanceTimersByTimeAsync(5_000);
      expect(elapsed()).toBe("");
    } finally {
      vi.useRealTimers();
    }
  });

  it("clears the bar when the note turns into an error", async () => {
    vi.stubGlobal("Office", { context: { host: "PowerPoint", requirements: { isSetSupported: () => false } } });
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("host refused");
      },
    });
    await bootPane();
    $<HTMLButtonElement>("demo-insert").click();
    await vi.waitFor(() => expect(noteEl().textContent).toMatch(/^Failed:/));
    // The message stays; the "still working" signal must not.
    expect(strip().hasAttribute("hidden")).toBe(false);
    expect(bar().hasAttribute("hidden")).toBe(true);
  });
});

describe("task pane — tab accessibility (ARIA tabs pattern)", () => {
  beforeEach(async () => {
    await bootPane();
  });
  const tabs = () => Array.from(document.querySelectorAll<HTMLButtonElement>(".tabs .tab"));

  it("marks the active tab aria-selected and roves tabindex", () => {
    const t = tabs();
    const active = t.find((x) => x.classList.contains("active"))!;
    expect(active.getAttribute("aria-selected")).toBe("true");
    expect(active.tabIndex).toBe(0);
    const inactive = t.find((x) => !x.classList.contains("active"))!;
    expect(inactive.getAttribute("aria-selected")).toBe("false");
    expect(inactive.tabIndex).toBe(-1);
  });

  it("moves the selection with the arrow keys and follows focus", () => {
    const t = tabs();
    t[0].focus();
    t[0].dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    expect(t[1].getAttribute("aria-selected")).toBe("true");
    expect(document.activeElement).toBe(t[1]);
    // Home jumps back to the first.
    t[1].dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true }));
    expect(t[0].getAttribute("aria-selected")).toBe("true");
  });
});

describe("task pane — overflow menu accessibility (ARIA menu pattern)", () => {
  beforeEach(async () => {
    await bootPane();
  });

  it("exposes menuitems, opens into the menu, and returns focus to the trigger on Escape", () => {
    const btn = $("more-actions") as HTMLButtonElement;
    const menu = document.getElementById("actions-menu")!;
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button"));
    expect(items.length).toBeGreaterThan(0);
    expect(items.every((i) => i.getAttribute("role") === "menuitem")).toBe(true);

    btn.click();
    expect(menu.hidden).toBe(false);
    expect(btn.getAttribute("aria-expanded")).toBe("true");
    expect(document.activeElement).toBe(items[0]);

    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(menu.hidden).toBe(true);
    expect(document.activeElement).toBe(btn);
  });

  it("announces status through a polite live region", () => {
    const n = $("host-note");
    expect(n.getAttribute("role")).toBe("status");
    expect(n.getAttribute("aria-live")).toBe("polite");
  });
});

describe("task pane — PNG export", () => {
  it("exposes a Download PNG menuitem beside Download SVG", async () => {
    await bootPane();
    const btn = $("download-png");
    expect(btn).toBeTruthy();
    expect(btn.getAttribute("role")).toBe("menuitem");
  });

  it("rasterizes the preview to a ssf-charts.png download", async () => {
    await bootPane();
    // jsdom decodes no SVG image and has no 2D canvas, so stand both in.
    class FakeImage {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      set src(_v: string) {
        Promise.resolve().then(() => this.onload?.());
      }
    }
    vi.stubGlobal("Image", FakeImage);
    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue({
      scale() {},
      drawImage() {},
    } as unknown as CanvasRenderingContext2D);
    vi.spyOn(HTMLCanvasElement.prototype, "toBlob").mockImplementation(function (cb: BlobCallback) {
      cb(new Blob(["png"], { type: "image/png" }));
    });
    const clicks: string[] = [];
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
      clicks.push(this.download);
    });
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:x");
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    try {
      $("download-png").click();
      await new Promise((r) => setTimeout(r, 5));
      expect(clicks).toContain("ssf-charts.png");
    } finally {
      vi.unstubAllGlobals();
      vi.restoreAllMocks();
    }
  });

  it("surfaces an error note when the browser can't decode the SVG", async () => {
    await bootPane();
    class FailImage {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      set src(_v: string) {
        Promise.resolve().then(() => this.onerror?.());
      }
    }
    vi.stubGlobal("Image", FailImage);
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:x");
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    try {
      $("download-png").click();
      await new Promise((r) => setTimeout(r, 5));
      expect($("host-note").textContent).toMatch(/PNG/);
      expect($("host-note").className).toContain("status-err");
    } finally {
      vi.unstubAllGlobals();
      vi.restoreAllMocks();
    }
  });
});

describe("task pane — the canvas is the chart's own colour", () => {
  it("previews and downloads a dark-theme chart on its own background, not white", async () => {
    // A forced white canvas put the dark theme's light text at 1.13:1 contrast
    // in the preview and in both downloads, while insertSceneIntoSlide drops the
    // same shapes onto the real (dark) slide with no background at all — the
    // preview contradicted the deck it is a preview of.
    await bootPane();
    importConfig({ kind: "clustered", data: baseData, style: { background: "#1b1b1a", text: "#f2f1ec" } });
    await new Promise((r) => setTimeout(r, 150)); // renderPreview is debounced
    expect($("preview").innerHTML).toContain('fill="#1b1b1a"');

    let blob: Blob | undefined;
    vi.spyOn(URL, "createObjectURL").mockImplementation((b: Blob | MediaSource) => {
      blob = b as Blob;
      return "blob:x";
    });
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => {});
    try {
      $("download").click();
      expect(await blob!.text()).toContain('fill="#1b1b1a"');
    } finally {
      vi.restoreAllMocks();
    }
  });

  it("still paints a default chart white", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: baseData });
    await new Promise((r) => setTimeout(r, 150));
    expect($("preview").innerHTML).toContain('fill="#ffffff"');
  });
});

describe("task pane — shareable chart link", () => {
  it("reopens the exact chart from a #c= share link on boot", async () => {
    const cfg = {
      kind: "radar",
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [1, 2, 3] }] },
      radar: { perSpoke: true },
    };
    const hash = "#c=" + btoa(encodeURIComponent(JSON.stringify(cfg)));
    await bootPane(hash);
    const out = exportConfig();
    expect(out.kind).toBe("radar");
    expect(out.radar!.perSpoke).toBe(true);
  });

  it("ignores a malformed share link and boots the default chart", async () => {
    await bootPane("#c=not-valid-base64!!");
    expect(() => exportConfig()).not.toThrow();
    expect($("host-note").className).not.toContain("status-err");
  });

  it("ignores a malformed ?tab deep link instead of aborting the rest of boot", async () => {
    // ?tab is a shipped public deep link (the ribbon uses it). Interpolating it
    // into a selector made querySelector THROW at module top level on a quote,
    // and everything after that line was skipped: no build stamp, no size
    // inputs, and no wireInsert() — leaving Insert enabled with no handler.
    await bootPane("?tab=" + encodeURIComponent('chart"]'));
    expect($("build-stamp").textContent).not.toBe("");
    expect(($("chart-w") as HTMLInputElement).value).not.toBe("");
    // wireInsert() ran: outside PowerPoint it disables Insert and says so.
    expect(($("insert") as HTMLButtonElement).disabled).toBe(true);
    expect($("host-note").textContent).toMatch(/PowerPoint/);
  });

  it("copies a link that round-trips the current chart through the clipboard", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: baseData, decorations: { grandTotal: true } });
    let copied = "";
    Object.defineProperty(navigator, "clipboard", {
      configurable: true,
      value: {
        writeText: async (t: string) => {
          copied = t;
        },
      },
    });
    try {
      $("copy-link").click();
      await new Promise((r) => setTimeout(r, 5));
      expect(copied).toContain("#c=");
      const encoded = copied.slice(copied.indexOf("#c=") + 3);
      const cfg = JSON.parse(decodeURIComponent(atob(encoded)));
      expect(cfg.kind).toBe("clustered");
      expect(cfg.decorations.grandTotal).toBe(true);
    } finally {
      // @ts-expect-error — remove the stub
      delete navigator.clipboard;
    }
  });
});

describe("element previews are sized for their own shape", () => {
  it("does not stretch the KPI tile like the process flow", async () => {
    // The flow is 480x44 and built to shrink, so it wants width:100%. The KPI
    // is 160x90 — the same rule blows it up to ~2.5x its natural height, which
    // pushed the KPI card's own Insert button out of view and left "Table from
    // datasheet"'s Insert as the nearest one. The owner clicked it and got a
    // table, exactly as the layout invited.
    await bootPane();
    expect(document.getElementById("kpi-preview")!.className).toBe("element-tile-preview");
    expect(document.getElementById("flow-preview")!.className).toBe("element-flow-preview");
    const css = readFileSync("src/taskpane/taskpane.css", "utf8");
    // A tile caps at its natural size; only the flow may stretch.
    expect(css).toMatch(/\.element-tile-preview svg \{[^}]*max-width:\s*100%/);
    expect(css).not.toMatch(/\.element-tile-preview svg \{[^}]*[^-]width:\s*100%/);
    expect(css).toMatch(/\.element-flow-preview svg \{[^}]*[^-]width:\s*100%/);
  });

  it("renders the KPI preview at the tile's real aspect, not the card's width", async () => {
    await bootPane();
    const svg = document.querySelector("#kpi-preview svg")!;
    // sceneToSvg states the scene's own size; the CSS must not override it.
    expect(svg.getAttribute("width")).toBe("160");
    expect(svg.getAttribute("height")).toBe("90");
  });
});

/**
 * PowerPoint dies; the task pane does not.
 *
 * The pane is a separate frame, and the `PowerPoint.run` promise it is waiting
 * on simply never settles when the host process goes — there is no rejection to
 * catch and no timeout that helps, because nothing ever answers. So the pane
 * counted upward under the word "Working…" forever, and the owner watched a
 * number climb on three separate crashes with no way to tell it from a chart
 * that was merely slow.
 *
 * Silence is the only evidence available, so silence is what the readout
 * carries. Checked as a rule rather than through a fake clock: what was wrong
 * was the decision, and the decision is one function.
 */
describe("the elapsed readout", () => {
  it("says only the elapsed seconds while a run is talking", async () => {
    await bootPane();
    const { elapsedLabel } = await import("../src/taskpane/app");
    expect(elapsedLabel(17_000, 0)).toBe("17s");
    // A draw batch on PowerPoint web has been measured at ~17s and a stalled
    // sync at 45, so neither may be called silent. A warning that fires on a
    // merely slow host is one nobody believes.
    expect(elapsedLabel(60_000, 45_000)).toBe("60s");
  });

  it("says how long the silence has lasted once it is not credible", async () => {
    await bootPane();
    const { elapsedLabel } = await import("../src/taskpane/app");
    expect(elapsedLabel(200_000, 190_000)).toBe("200s · silent for 190s");
  });
});

/**
 * The pane's ingest boundary, handed the configs it writes itself.
 *
 * `stateFromConfig` trusted its own types. The ENGINE does not —
 * `test/chart-hostile-input.test.ts` pins that `{ palette: "red" }` renders
 * fine — so a config that draws correctly in the preview, in the deck and in
 * the skill's .pptx could not be OPENED in the pane that has to edit it. The
 * JSON box reported the user's perfectly valid JSON as
 * `Invalid JSON: palette.join is not a function`, sending them to hunt a syntax
 * error that is not there.
 *
 * It is more than a hostile-input case because the pane WRITES these: a style
 * file imported through `style-import` is persisted with no shape check, and
 * from then on every export, every saved template and every POWERCHART_CONFIG
 * shape tag carries it.
 */
describe("a config whose arrays are not arrays", () => {
  const note = () => $("host-note").textContent ?? "";

  it("opens a config the pane itself could have written", () => {
    importConfig({
      kind: "stacked",
      style: { palette: "red" },
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
    } as unknown as Partial<ChartConfig>);
    expect(note(), "the pane called its own valid JSON invalid").not.toMatch(/Invalid JSON/);
    // And it round-trips, rather than merely not throwing.
    expect(() => exportConfig()).not.toThrow();
  });

  it("survives every field that used to throw, and says nothing about invalid JSON", () => {
    const cases: [string, Partial<ChartConfig>][] = [
      ["palette as a string", { style: { palette: "red" } } as unknown as Partial<ChartConfig>],
      ["palette as an object", { style: { palette: { a: 1 } } } as unknown as Partial<ChartConfig>],
      ["palette as a number", { style: { palette: 5 } } as unknown as Partial<ChartConfig>],
      ["labelContent as a string", { decorations: { labelContent: "value" } } as unknown as Partial<ChartConfig>],
      ["pie.explode as a number", { pie: { explode: 1 } } as unknown as Partial<ChartConfig>],
      [
        "waterfall.totalIndices as a number",
        { kind: "waterfall", waterfall: { totalIndices: 1 } } as unknown as Partial<ChartConfig>,
      ],
    ];
    for (const [what, extra] of cases) {
      importConfig({
        kind: "stacked",
        data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
        ...extra,
      } as unknown as Partial<ChartConfig>);
      expect(note(), `${what}: the pane reported valid JSON as invalid`).not.toMatch(/Invalid JSON/);
    }
  });

  it("drops an out-of-range waterfall total instead of holing the datasheet grid", () => {
    // The write had no bound on `i`, so it wrote past the row's end and left
    // `cells[1]` SPARSE. `mountDatasheet`'s `row.forEach` skips holes, so the
    // grid rendered one stray cell under no column header while every other row
    // was short — visibly misaligned, and editing nothing the chart reads.
    importConfig({
      kind: "waterfall",
      data: { categories: ["a", "b", "c"], series: [{ name: "Delta", values: [1, 2, 3] }] },
      waterfall: { totalIndices: [8] },
    } as unknown as Partial<ChartConfig>);
    expect(note()).not.toMatch(/Invalid JSON/);
    const rows = [...document.querySelectorAll("#datasheet tr")];
    const widths = new Set(rows.map((r) => r.querySelectorAll("td,th").length));
    expect(widths.size, `the grid has ragged rows: ${[...widths].join(", ")}`).toBe(1);
  });

  // NO in-range negative control here, deliberately. One was written and
  // removed: it failed against a build with ONLY the bound removed, which the
  // arithmetic says it should have passed, so it was not testing what it
  // claimed. A guard nobody can explain is worse than no guard — the trap this
  // repo already records as "a guard can fail for the wrong reason and still
  // look proven". The in-range path stays covered by the ragged-row assertion
  // above, which requires the grid to stay rectangular whatever is written.
});

/**
 * `??` catches null and undefined. A config's `width` can be anything.
 *
 * A config arrives from the JSON box, a saved template, a shape tag written in
 * another deck and the skill's caller, so `width: number` in the types is a
 * promise nothing enforces. The pane built the size it hands `placeChart` as
 * `cfg.width ?? DEFAULT_SIZE.width` — which passes NaN, Infinity, 0 and
 * negatives straight through, because none of them is nullish.
 *
 * The engine's own `clampDim` handles exactly those, and it runs INSIDE
 * `buildChart` — after the placement. So `placeChart` shrank a NaN width to fit
 * the slide and returned **882 points**: wider than the 720pt slide, finite,
 * and therefore waved through by every check after it. The user gets a chart
 * running off the edge with nothing said.
 *
 * One rule in one place. Two clamps agree until somebody changes one.
 */
describe("a chart size the pane will act on", () => {
  it("refuses the values `??` lets through", () => {
    for (const v of [NaN, Infinity, -Infinity, 0, -5]) {
      expect(clampDim(v, 480), `clampDim(${v}) kept a size no chart can have`).toBe(480);
    }
    expect(clampDim(undefined, 480)).toBe(480);
    // And a real width is untouched, or this "fix" would be a regression.
    expect(clampDim(300, 480)).toBe(300);
    // Above the engine's ceiling it clamps rather than falls back: a too-wide
    // chart is a user's intent expressed badly, not a broken value.
    expect(clampDim(1e6, 480)).toBeLessThanOrEqual(7199);
  });

  it("never hands the placer a size that becomes a chart wider than a slide", () => {
    // The number the bug actually produced, reached through the real call.
    const occupied = [{ left: 0, top: 0, width: 100, height: 100 }];
    for (const v of [NaN, Infinity, 0, -5]) {
      const size = { width: clampDim(v, 480), height: clampDim(v, 300) };
      const p = placeChart(occupied, size, { left: 60, top: 90 }, { left: 60, top: 90 });
      expect(Number.isFinite(p.width), `width ${v} → ${p.width}`).toBe(true);
      expect(Number.isFinite(p.height), `height ${v} → ${p.height}`).toBe(true);
      expect(p.width, `width ${v} produced a chart wider than the slide`).toBeLessThanOrEqual(720);
    }
  });
});

/**
 * The Same Scale outcome sentence.
 *
 * Pure, and tested directly for the reason `elapsedLabel` above it is: what was
 * wrong was the decision, and the decision is one function. Three notes used to
 * go out back to back into a channel that holds one, so the last erased the
 * others and a run with both qualifiers reported only the degraded clause.
 */
describe("what Same Scale says it did", () => {
  it("keeps the qualifiers after the outcome, in that order", async () => {
    await bootPane();
    const { sameScaleNote } = await import("../src/taskpane/app");
    expect(sameScaleNote({ base: "Applied to 6.", rescued: "2 as pictures.", degraded: "1 fell back." })).toEqual({
      text: "Applied to 6. 2 as pictures. 1 fell back.",
      status: "err",
    });
  });

  it("is just the outcome when nothing needed qualifying", async () => {
    await bootPane();
    const { sameScaleNote } = await import("../src/taskpane/app");
    expect(sameScaleNote({ base: "Applied to 6." })).toEqual({ text: "Applied to 6.", status: "ok" });
  });

  it("stays GREEN for a rescue, because the guard working is not a failure", async () => {
    /**
     * `chartPicture` returns a warn WITH a png on its success path — the chart
     * was too dense and was rasterised rather than drawn. Reporting that in red
     * claimed the dangerous thing had happened when the guard against it had
     * just worked, and steered the user away from "Explode to native shapes".
     */
    await bootPane();
    const { sameScaleNote } = await import("../src/taskpane/app");
    expect(sameScaleNote({ base: "Applied to 6.", rescued: "2 as pictures." }).status).toBe("ok");
  });

  it("turns RED as soon as something fell back", async () => {
    await bootPane();
    const { sameScaleNote } = await import("../src/taskpane/app");
    expect(sameScaleNote({ base: "Applied to 6.", degraded: "1 fell back." }).status).toBe("err");
  });
});
