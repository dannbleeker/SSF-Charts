// @vitest-environment jsdom
import { afterAll, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { buildChart, normalizeConfig } from "../src/core/chart";
import { formatNumber, localeTag, resolveAxisFormat, resolveFormat } from "../src/core/format";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { finiteNodes } from "../src/core/scene";
import type { SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { EN } from "../src/taskpane/i18n";
import {
  decimalMark,
  fromEuropeanNumber,
  looksEuropean,
  sheetToData,
  spellNumber,
  splitValueList,
  typedNumber,
} from "../src/taskpane/datasheet";
import { sceneToSvg } from "../src/render/svg";

/**
 * THE THIRD BUG HUNT, STREAM m2: what the third fix run could not finish in its
 * own files. One describe per finding.
 *
 * The pane tests boot the real `app.ts` against `taskpane.html` the way
 * `pane-state.test.ts` does; the rest are pure functions and need no DOM.
 */

/** See `pane-state.test.ts`: the pane's last debounced render must not outlive jsdom. */
afterAll(async () => {
  await new Promise((r) => setTimeout(r, 250));
});

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

async function bootPane() {
  try {
    window.localStorage.clear();
  } catch {
    /* a jsdom without storage */
  }
  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  const app = await import("../src/taskpane/app");
  await app.harnessReady;
}

const importConfig = (cfg: Partial<ChartConfig>) => {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(cfg);
  $("json-import").click();
};
const exportConfig = (): ChartConfig => {
  $("json-export").click();
  return JSON.parse(($("json-io") as HTMLTextAreaElement).value);
};
const typeInto = (el: HTMLInputElement, value: string) => {
  el.value = value;
  el.dispatchEvent(new Event("input", { bubbles: true }));
};
const tick = (box: HTMLInputElement, on: boolean) => {
  box.checked = on;
  box.dispatchEvent(new Event("change", { bubbles: true }));
};
/** The checkbox in the options label that starts with `prefix`. */
const checkbox = (prefix: string): HTMLInputElement => {
  const row = [...document.querySelectorAll<HTMLElement>("#options label")].find((l) =>
    (l.textContent ?? "").startsWith(prefix),
  );
  if (!row) throw new Error(`no options row starting "${prefix}"`);
  return row.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
};
/** The two text boxes of the options row that starts with `prefix` ("Axis scale min"). */
const boxes = (prefix: string): HTMLInputElement[] => {
  const row = [...document.querySelectorAll<HTMLElement>("#options label")].find((l) =>
    (l.textContent ?? "").startsWith(prefix),
  );
  if (!row) throw new Error(`no options row starting "${prefix}"`);
  return [...row.querySelectorAll<HTMLInputElement>('input[type="text"]')];
};

const texts = (cfg: ChartConfig): TextNode[] => buildChart(cfg).nodes.filter((n): n is TextNode => n.kind === "text");
const textsNamed = (cfg: ChartConfig, name: string | RegExp): string[] =>
  texts(cfg)
    .filter((n) => (typeof name === "string" ? n.name === name : name.test(n.name ?? "")))
    .map((n) => n.text);

/* -------------------------------------------------------------------------- */

describe("M2a: numberFormat.grouping switches the thousands separator off", () => {
  const YEARS = [2015, 2016, 2017, 2018];
  const scatter = (numberFormat?: unknown, decorations: object = {}): ChartConfig =>
    ({
      kind: "scatter",
      width: 480,
      height: 300,
      data: {
        categories: ["a", "b", "c", "d"],
        series: [
          { name: "X", values: YEARS },
          { name: "Y", values: [1200, 2300, 3400, 4500] },
        ],
      },
      decorations,
      ...(numberFormat === undefined ? {} : { numberFormat }),
    }) as unknown as ChartConfig;

  it("formatNumber: 2015 is a year when the author says so, and 2,015 otherwise", () => {
    expect(formatNumber(2015)).toBe("2,015");
    expect(formatNumber(2015, { grouping: true })).toBe("2,015");
    expect(formatNumber(2015, { grouping: false })).toBe("2015");
    expect(formatNumber(1234567.5, { grouping: false, decimals: 1 })).toBe("1234567.5");
  });

  it("formatNumber: the other halves of the format are untouched, in the chart's own locale", () => {
    expect(formatNumber(1234.5, { grouping: false, decimals: 1, locale: "de-DE" })).toBe("1234,5");
    expect(formatNumber(1234.5, { decimals: 1, locale: "de-DE" })).toBe("1.234,5");
    expect(formatNumber(2015, { grouping: false, suffix: " €m" })).toBe("2015 €m");
    expect(formatNumber(2015, { grouping: false, forceSign: true })).toBe("+2015");
    expect(formatNumber(-2015, { grouping: false })).toBe("-2015");
  });

  it("formatNumber: grouping is read strictly, so only the documented false changes anything", () => {
    for (const junk of ["false", "no", 0, null, undefined, "", {}, []]) {
      expect(formatNumber(2015, { grouping: junk } as never), JSON.stringify(junk)).toBe("2,015");
    }
  });

  it("formatNumber: a locale that does not group four digits by default still does not", () => {
    // `useGrouping: true` is NOT what the default means, and passing it would
    // have changed these: es-ES leaves a four-digit number whole.
    const es = new Intl.NumberFormat("es-ES", { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(2015);
    expect(formatNumber(2015, { locale: "es-ES" })).toBe(es);
    expect(formatNumber(2015, { locale: "es-ES", grouping: false })).toBe("2015");
  });

  it("resolveFormat and resolveAxisFormat carry the switch to whoever formats with their answer", () => {
    expect(resolveFormat(YEARS, { grouping: false }).grouping).toBe(false);
    expect(resolveAxisFormat(YEARS, { grouping: false }).grouping).toBe(false);
    expect(resolveFormat(YEARS, {}).grouping).toBeUndefined();
  });

  it("a scatter of years prints 2015 … 2018 on the x axis and beside its points", () => {
    const on = scatter({ grouping: false }, { labelContent: ["category", "value"] });
    const axis = textsNamed(on, "x-axis");
    expect(axis, "the x axis").toEqual(["2015", "2016", "2017", "2018"]);
    expect(textsNamed(on, "y-axis"), "the y axis follows: one switch for the chart").toEqual([
      "1000",
      "2000",
      "3000",
      "4000",
    ]);
    const labels = textsNamed(on, /^label-\d+$/);
    expect(labels).toEqual(["a (2015, 1200)", "b (2016, 2300)", "c (2017, 3400)", "d (2018, 4500)"]);
    expect(sceneToSvg(buildChart(on))).not.toMatch(/\d,\d{3}/);
  });

  it("the default still prints 2,015 (the control)", () => {
    const off = scatter(undefined, { labelContent: ["category", "value"] });
    expect(textsNamed(off, "x-axis")).toEqual(["2,015", "2,016", "2,017", "2,018"]);
    expect(textsNamed(off, /^label-\d+$/)[0]).toBe("a (2,015, 1,200)");
  });

  it("every kind: no number the chart prints keeps a separator once it is off", () => {
    // The sweep that finds a site that forgot. Values are scaled into the
    // thousands so every kind has a number to group; `Group` is a category index.
    // Bubble is swept without its Size key (see the todo below).
    const grouped = (t: string) => /\d,\d{3}/.test(t);
    const kindsWithAGroupedNumber: string[] = [];
    for (const { kind } of CHART_KINDS) {
      const base = sampleConfig(kind);
      const cfg = {
        ...base,
        data: {
          ...base.data,
          series: base.data.series.map((s) =>
            s.name === "Group" ? s : { ...s, values: s.values.map((v) => (v == null ? v : v * 1000)) },
          ),
        },
        decorations: {
          ...base.decorations,
          segmentLabels: true,
          totals: true,
          grandTotal: true,
          valueAxis: true,
          ...(kind === "scatter" || kind === "bubble" ? { labelContent: ["value"] } : {}),
        },
      } as ChartConfig;
      const withoutKey = (ts: TextNode[]) =>
        kind === "bubble" ? ts.filter((t) => !/^size-legend-label-/.test(t.name ?? "")) : ts;
      if (texts(cfg).some((t) => grouped(t.text))) kindsWithAGroupedNumber.push(kind);
      const left = withoutKey(texts({ ...cfg, numberFormat: { grouping: false } })).filter((t) => grouped(t.text));
      expect(
        left.map((t) => `${t.name}: ${t.text}`),
        `${kind} still groups`,
      ).toEqual([]);
    }
    // The sweep means nothing if it only ever looked at numbers with no thousands.
    expect(kindsWithAGroupedNumber.length, "kinds whose default has a grouped number").toBeGreaterThan(15);
  });

  it.todo(
    "bubble Size key, scatter Color legend and gantt `Column` cells pick `{ decimals, locale }` by hand " +
      "(scatter.ts, gantt.ts) and so keep their separator: they need `grouping` carried with the locale",
  );

  it("the default is byte-identical: grouping true, or anything but false, is the chart without the key", () => {
    for (const { kind } of CHART_KINDS) {
      const base = sampleConfig(kind);
      const plain = JSON.stringify(buildChart(base));
      for (const grouping of [true, "false", 0, null, "no"]) {
        const cfg = { ...base, numberFormat: { ...base.numberFormat, grouping } } as unknown as ChartConfig;
        expect(JSON.stringify(buildChart(cfg)), `${kind} grouping ${JSON.stringify(grouping)}`).toBe(plain);
      }
    }
  });

  it("switching it off changes nothing on a chart that has nothing to group", () => {
    // Under a thousand there is no separator to remove, so the scene is the same.
    for (const kind of ["stacked", "line", "pie"] as const) {
      const base = sampleConfig(kind);
      expect(JSON.stringify(buildChart({ ...base, numberFormat: { grouping: false } })), kind).toBe(
        JSON.stringify(buildChart(base)),
      );
    }
  });

  describe("in the pane", () => {
    it("is a ticked 'Thousands separator' box in the Labels group, and a catalogued string", async () => {
      expect(Object.keys(EN)).toContain("Thousands separator");
      await bootPane();
      const box = checkbox("Thousands separator");
      expect(box.checked).toBe(true);
      expect(box.closest("details")?.dataset.group, "it lives with the other label options").toBe("labels");
    });

    it("leaves a chart nobody touched exactly as it was", async () => {
      await bootPane();
      importConfig({ kind: "clustered", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } });
      expect(exportConfig().numberFormat).toBeUndefined();
      importConfig({
        kind: "clustered",
        data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
        numberFormat: { decimals: 2 },
      });
      expect(exportConfig().numberFormat, "no grouping key is written for the default").toEqual({ decimals: 2 });
    });

    it("unticking it writes grouping:false, and the preview follows", async () => {
      await bootPane();
      importConfig(scatter(undefined, { valueAxis: true }));
      tick(checkbox("Thousands separator"), false);
      expect(exportConfig().numberFormat?.grouping).toBe(false);
      await vi.waitFor(() => expect($("preview").innerHTML).toContain(">2015<"), { timeout: 15_000 });
      expect($("preview").innerHTML).not.toContain("2,015");
    });

    it("ticking it again takes the key out, whatever the chart arrived with", async () => {
      await bootPane();
      importConfig(scatter({ grouping: false, decimals: 1, forceSign: true }));
      const box = checkbox("Thousands separator");
      expect(box.checked, "a chart that says false shows an unticked box").toBe(false);
      tick(box, true);
      const out = exportConfig().numberFormat;
      expect(out?.grouping).toBeUndefined();
      expect(out, "the neighbours survive").toMatchObject({ decimals: 1, forceSign: true });
    });

    it("a loaded grouping:false survives an edit that does not touch the box", async () => {
      await bootPane();
      importConfig(scatter({ grouping: false, suffix: " km" }));
      typeInto(document.querySelector<HTMLInputElement>('input[placeholder="e.g. €m"]')!, " mi");
      expect(exportConfig().numberFormat).toMatchObject({ grouping: false });
    });

    it("a hand-written non-boolean is a ticked box, and is repaired on the way out", async () => {
      await bootPane();
      importConfig(scatter({ grouping: "no", decimals: 0 }));
      expect(checkbox("Thousands separator").checked).toBe(true);
      expect(exportConfig().numberFormat?.grouping).toBeUndefined();
    });

    // The Table element's computed cells are tested where the Table element is:
    // `test/fix-p2-batch-table-roundtrip.test.ts`, which has the fake host it needs.
  });
});

/* -------------------------------------------------------------------------- */

describe("M2b: the pane reads typed numbers in the locale the chart prints", () => {
  it("localeTag is the engine's reading of a locale string, exported", () => {
    expect(localeTag("de_DE")).toBe("de-DE");
    expect(localeTag("zh_Hans_CN")).toBe("zh-Hans-CN");
    expect(localeTag("da-DK")).toBe("da-DK");
  });

  it("decimalMark: de_DE is German, like de-DE", () => {
    for (const spelled of ["de_DE", "da_DK", "fr_FR"]) {
      expect(decimalMark(spelled), spelled).toBe(",");
      expect(decimalMark(spelled.replace("_", "-")), spelled).toBe(",");
    }
    for (const spelled of ["en_US", "en_GB", "ja_JP"]) expect(decimalMark(spelled), spelled).toBe(".");
  });

  it("decimalMark: a locale the runtime cannot parse is still a dot, not an error", () => {
    expect(decimalMark("not a locale")).toBe(".");
    expect(decimalMark("")).toBe(".");
  });

  it("typedNumber reads 1,500 and 3.000 the German way under de_DE", () => {
    // The two shapes that depend on the locale: one mark, then exactly three digits.
    expect(typedNumber("1,5", "de_DE")).toBe(1.5);
    expect(typedNumber("1,500", "de_DE")).toBe(1.5);
    expect(typedNumber("3.000", "de_DE")).toBe(3000);
    for (const text of ["1,5", "1,500", "3.000", "0,5", "1.234,5"]) {
      expect(typedNumber(text, "de_DE"), text).toBe(typedNumber(text, "de-DE"));
    }
    // And the English reading is still what en_US gives.
    expect(typedNumber("1,500", "en_US")).toBe(1500);
    expect(typedNumber("3.000", "en_US")).toBe(3);
  });

  it("spellNumber and splitValueList agree", () => {
    expect(spellNumber(1.5, "de_DE")).toBe("1,5");
    expect(splitValueList("1,5", "de_DE")).toEqual(["1,5"]);
    expect(splitValueList("1,5", "en_US")).toEqual(["1", "5"]);
  });

  it("a config that says de_DE has the pane read its Axis scale boxes in German", async () => {
    await bootPane();
    importConfig({
      kind: "clustered",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
      numberFormat: { locale: "de_DE" },
    });
    const [min, max] = boxes("Axis scale min");
    typeInto(min, "1,500");
    typeInto(max, "3.000");
    expect(exportConfig().scale).toEqual({ min: 1.5, max: 3000 });
  });
});

/* -------------------------------------------------------------------------- */

describe("M2c: the finite-number guard sees a missing coordinate", () => {
  /**
   * The smallest node of each kind: ONLY what it cannot be drawn without. Typed
   * per kind, so a kind added to the union is a compile error here until it is
   * answered, and the required fields are exactly the numbers in these literals.
   */
  const MINIMAL: { [K in SceneNode["kind"]]: Extract<SceneNode, { kind: K }> } = {
    rect: { kind: "rect", x: 1, y: 2, w: 3, h: 4, fill: "#000000" },
    line: { kind: "line", x1: 1, y1: 2, x2: 3, y2: 4, stroke: "#000000" },
    text: {
      kind: "text",
      x: 1,
      y: 2,
      w: 3,
      h: 4,
      text: "t",
      fontSize: 9,
      color: "#000000",
      align: "left",
      valign: "top",
    },
    ellipse: { kind: "ellipse", cx: 1, cy: 2, rx: 3, ry: 4, fill: "#000000" },
    wedge: {
      kind: "wedge",
      cx: 1,
      cy: 2,
      r: 3,
      innerR: 0,
      startAngle: 0,
      endAngle: 90,
      fill: "#000000",
    },
    chevron: { kind: "chevron", x: 1, y: 2, w: 3, h: 4, fill: "#000000" },
    symbol: { kind: "symbol", shape: "diamond", cx: 1, cy: 2, size: 3, fill: "#000000" },
    arrowhead: { kind: "arrowhead", x: 1, y: 2, angle: 90, size: 3, fill: "#000000" },
    polygon: {
      kind: "polygon",
      points: [
        { x: 0, y: 0 },
        { x: 5, y: 5 },
        { x: 0, y: 5 },
      ],
    },
  };
  const numberFields = (n: SceneNode): string[] =>
    Object.entries(n)
      .filter(([, v]) => typeof v === "number")
      .map(([k]) => k);

  it("keeps the minimal node of every kind: a field a kind can do without is not required", () => {
    for (const kind of Object.keys(MINIMAL) as SceneNode["kind"][]) {
      expect(finiteNodes([MINIMAL[kind]]), kind).toHaveLength(1);
    }
  });

  it("drops a node lacking ANY of the numbers its kind is drawn from, however it lacks it", () => {
    const survivors: string[] = [];
    for (const kind of Object.keys(MINIMAL) as SceneNode["kind"][]) {
      for (const field of numberFields(MINIMAL[kind])) {
        const without = { ...MINIMAL[kind] } as Record<string, unknown>;
        delete without[field];
        // Absent, `undefined` (the shape a bad index produced), null, and text.
        const spoiled: unknown[] = [without, { ...without, [field]: undefined }, { ...without, [field]: null }];
        spoiled.push({ ...without, [field]: "12" });
        for (const node of spoiled) if (finiteNodes([node as SceneNode]).length) survivors.push(`${kind}.${field}`);
      }
    }
    expect([...new Set(survivors)], "a node with a missing field got through").toEqual([]);
  });

  it("drops a polygon whose point is missing a coordinate, and keeps the rest of the scene", () => {
    const bad = (p: object): SceneNode => ({ kind: "polygon", points: [{ x: 0, y: 0 }, { x: 5, y: 5 }, p as never] });
    const kept = finiteNodes([
      MINIMAL.rect,
      bad({ x: 1 }),
      bad({ y: 1 }),
      bad({ x: undefined, y: 1 }),
      bad(null as never),
      MINIMAL.polygon,
    ]);
    expect(kept.map((n) => n.kind)).toEqual(["rect", "polygon"]);
  });

  it("still drops a non-finite number, and a degenerate polygon (the two it always did)", () => {
    expect(
      finiteNodes([{ ...MINIMAL.rect, x: Number.NaN } as SceneNode, { ...MINIMAL.line, y2: Infinity } as SceneNode]),
    ).toEqual([]);
    expect(finiteNodes([{ kind: "polygon", points: [{ x: 1, y: 1 }] }])).toEqual([]);
  });

  it("keeps a node that leaves out its OPTIONAL fields, and checks one that has them", () => {
    const optional: SceneNode[] = [
      { ...MINIMAL.rect, strokeWidth: 1, stroke: "#111111", pattern: "dots", name: "r" },
      { ...MINIMAL.line, strokeWidth: 2, dash: [2, 2], name: "l" },
      { ...MINIMAL.polygon, fillOpacity: 0.5, strokeWidth: 1 } as SceneNode,
      { ...MINIMAL.chevron, flatLeft: true },
    ];
    expect(finiteNodes(optional)).toHaveLength(optional.length);
    expect(finiteNodes([{ ...MINIMAL.rect, strokeWidth: Number.NaN }])).toEqual([]);
  });

  it('the hole it closes: an arrow with y1 undefined reached the SVG as y1="NaN"', () => {
    const arrow = { ...MINIMAL.line, y1: undefined, name: "difference-arrow" } as unknown as SceneNode;
    const svg = (nodes: SceneNode[]) => sceneToSvg({ width: 100, height: 100, nodes });
    expect(svg([arrow, MINIMAL.rect]), "the SVG renderer is not the guard").toContain("NaN");
    expect(svg(finiteNodes([arrow, MINIMAL.rect]))).not.toContain("NaN");
  });

  it("a kind this gate has never heard of is not its to judge", () => {
    const stranger = { kind: "starburst", x: 1 } as unknown as SceneNode;
    expect(finiteNodes([stranger])).toEqual([stranger]);
  });

  it("the engine's own scenes still lose nothing to it: every sample kind, both ways round", () => {
    for (const { kind } of CHART_KINDS) {
      for (const horizontal of [false, true]) {
        const scene = buildChart({ ...sampleConfig(kind), horizontal });
        expect(finiteNodes(scene.nodes), `${kind} horizontal=${horizontal}`).toHaveLength(scene.nodes.length);
      }
    }
  });
});

/* -------------------------------------------------------------------------- */

describe("M2d: a combo's line labels honour the label content", () => {
  const pareto = (horizontal: boolean, labelContent?: unknown): ChartConfig =>
    ({
      kind: "clustered",
      width: 480,
      height: 300,
      pareto: true,
      horizontal,
      data: { categories: ["A", "B", "C", "D"], series: [{ name: "Defects", values: [30, 50, 5, 15] }] },
      decorations: labelContent === undefined ? {} : { labelContent },
    }) as unknown as ChartConfig;
  const combo = (horizontal: boolean, labelContent?: unknown, extra: Partial<ChartConfig> = {}): ChartConfig =>
    ({
      kind: "combo",
      width: 480,
      height: 300,
      horizontal,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Revenue", values: [10, 20, 15] },
          { name: "Margin", type: "line", values: [12, 30, 21] },
        ],
      },
      decorations: labelContent === undefined ? {} : { labelContent },
      ...extra,
    }) as unknown as ChartConfig;
  const lineLabels = (cfg: ChartConfig) => textsNamed(cfg, /^combo-label-0-\d+$/);

  it("a horizontal Pareto asked for categories labels its points with them", () => {
    // The cumulative line runs over the ranked bars: B, A, D, C.
    expect(lineLabels(pareto(true, ["category"]))).toEqual(["B", "A", "D", "C"]);
  });

  it("upright too: rotating a chart does not change what its labels say", () => {
    expect(lineLabels(pareto(false, ["category"]))).toEqual(["B", "A", "D", "C"]);
    for (const parts of [["category"], ["series"], ["category", "value"], ["series", "value"]]) {
      expect(lineLabels(combo(true, parts)), parts.join("+")).toEqual(lineLabels(combo(false, parts)));
    }
  });

  it("every part reads as it does on a line chart", () => {
    expect(lineLabels(combo(true, ["category", "value"]))).toEqual(["Q1 12", "Q2 30", "Q3 21"]);
    expect(lineLabels(combo(true, ["series"]))).toEqual(["Margin", "Margin", "Margin"]);
    expect(lineLabels(combo(true, ["series", "value"]))).toEqual(["Margin 12", "Margin 30", "Margin 21"]);
  });

  it("value + percent prints the value: a point has no share of anything", () => {
    expect(lineLabels(combo(true, ["value", "percent"]))).toEqual(["12", "30", "21"]);
    expect(lineLabels(pareto(true, ["value", "percent"]))).toEqual(["50", "80", "95", "100"]);
  });

  it("a list with nothing a point can print keeps the value the line always carried", () => {
    for (const parts of [["percent"], []]) {
      expect(lineLabels(combo(true, parts)), JSON.stringify(parts)).toEqual(["12", "30", "21"]);
      expect(lineLabels(combo(false, parts)), JSON.stringify(parts)).toEqual(["12", "30", "21"]);
    }
  });

  it("no label content is the chart it always was", () => {
    expect(lineLabels(pareto(true))).toEqual(["50", "80", "95", "100"]);
    expect(lineLabels(combo(false))).toEqual(["12", "30", "21"]);
  });

  it("the label at the plot's right edge still flips left with its new text", () => {
    const cfg = combo(true, ["category", "value"], {
      data: {
        categories: ["Quarter one", "Quarter two"],
        series: [
          { name: "Revenue", values: [10, 20] },
          { name: "Margin", type: "line", values: [12, 30] },
        ],
      },
    } as never);
    const labels = texts(cfg).filter((t) => /^combo-label-0-\d+$/.test(t.name ?? ""));
    expect(labels.map((t) => t.text)).toEqual(["Quarter one 12", "Quarter two 30"]);
    for (const t of labels) {
      expect(t.x, t.text).toBeGreaterThanOrEqual(-0.5);
      expect(t.x + t.w, t.text).toBeLessThanOrEqual(480.5);
    }
  });
});

/* -------------------------------------------------------------------------- */

describe("M2e: a thousands group cannot start with a zero", () => {
  const block = (...cells: string[]) => [["Region", "N"], ...cells.map((c) => ["x", c])];

  it("0.500 is the decimal 0.5, not five hundred", () => {
    expect(fromEuropeanNumber("0.500")).toBe("0.500");
    expect(fromEuropeanNumber("-0.500")).toBe("-0.500");
    expect(fromEuropeanNumber("+0.500")).toBe("+0.500");
    expect(fromEuropeanNumber("0.500 €")).toBe("0.500 €");
    expect(fromEuropeanNumber("0.000")).toBe("0.000");
    expect(
      sheetToData({
        cells: [
          ["", "A"],
          ["S", fromEuropeanNumber("0.500")],
        ],
      }).series[0].values,
    ).toEqual([0.5]);
  });

  it("1.500 is still fifteen hundred, and the other European shapes are as they were", () => {
    expect(fromEuropeanNumber("1.500")).toBe("1500");
    expect(fromEuropeanNumber("10.500")).toBe("10500");
    expect(fromEuropeanNumber("100.500")).toBe("100500");
    expect(fromEuropeanNumber("1.234.567,89")).toBe("1234567.89");
    expect(fromEuropeanNumber("-1.234,5")).toBe("-1234.5");
    expect(fromEuropeanNumber("0.5")).toBe("0.5");
    expect(fromEuropeanNumber("0,500")).toBe("0.500");
    expect(fromEuropeanNumber("(1.234,50)")).toBe("(1234.50)");
  });

  it("a number with a leading-zero first group and a comma decimal is no number, not a wrong one", () => {
    expect(fromEuropeanNumber("0.500,25")).toBe("0.500,25");
    expect(fromEuropeanNumber("00.500")).toBe("00.500");
    expect(fromEuropeanNumber("01.234")).toBe("01.234");
  });

  it("a European block that carries 0.500 is still judged European, and keeps its other cells", () => {
    const cells = block("1.234,5", "0.500", "2.250,25");
    expect(looksEuropean(cells)).toBe(true);
    expect(cells.slice(1).map(([, v]) => fromEuropeanNumber(v))).toEqual(["1234.5", "0.500", "2250.25"]);
  });

  it("0,500 is the decimal a Danish Excel prints at three decimals: 0.5, and no vote for an American sheet", () => {
    // It matched the US thousands pattern: read as 500, and counted as American
    // evidence, which kept the rest of a Danish block from being converted at all.
    expect(looksEuropean(block("0,500", "1.234,5"))).toBe(true);
    expect(looksEuropean(block("0,500", "0,750"))).toBe(true);
    expect(
      sheetToData({
        cells: [
          ["", "A"],
          ["S", "0,500"],
        ],
      }).series[0].values,
      "no European evidence: a gap",
    ).toEqual([null]);
    expect(
      sheetToData({
        cells: [
          ["", "A"],
          ["S", fromEuropeanNumber("0,500")],
        ],
      }).series[0].values,
    ).toEqual([0.5]);
  });

  it("the US shapes still read as US: 1,500 and 12,345.6 are groups", () => {
    expect(
      sheetToData({
        cells: [
          ["", "A", "B"],
          ["S", "1,500", "12,345.6"],
        ],
      }).series[0].values,
    ).toEqual([1500, 12345.6]);
    expect(looksEuropean(block("1,500", "2,500", "987.5"))).toBe(false);
    expect(looksEuropean(block("1,5", "0.5"))).toBe(false);
  });
});

/* A config the engine is handed whole still normalises with the field in place. */
describe("normalizeConfig leaves numberFormat.grouping as written", () => {
  it("a boolean is kept, and the engine reads anything else as the default", () => {
    const cfg = { ...sampleConfig("stacked"), numberFormat: { grouping: false } } as ChartConfig;
    expect(normalizeConfig(cfg).numberFormat?.grouping).toBe(false);
    const junk = { ...sampleConfig("stacked"), numberFormat: { grouping: "no" } } as unknown as ChartConfig;
    expect(JSON.stringify(buildChart(junk))).toBe(JSON.stringify(buildChart(sampleConfig("stacked"))));
  });
});
