// @vitest-environment jsdom
import { afterAll, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { evaluateFormula, mountDatasheet, sheetToData, type SheetModel } from "../src/taskpane/datasheet";
import { waterfallChain } from "../src/core/layout/waterfall";
import { buildChart, normalizeConfig } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * Datasheet paste, formulas and keyboard — the findings of stream p3.
 *
 * Each block names the finding it guards and the input that reached it. The
 * tests that guard a fix were each shown to FAIL against the code it replaced,
 * on the assertion that targets the bug; the few that pass either way are
 * controls pinning what must not move (a cell that is no number, a plain value,
 * a text box left alone) and say so in their comments.
 */

/**
 * Let the pane's last debounced render fire before jsdom goes away — the same
 * wait, and for the same reason, as `test/pane-state.test.ts` (an armed 80 ms
 * timer reads `document` after the environment is torn down).
 */
afterAll(async () => {
  await new Promise((r) => setTimeout(r, 250));
});

const input = (host: HTMLElement, r: number, c: number) =>
  host.querySelector<HTMLInputElement>(`input[data-row="${r}"][data-col="${c}"]`)!;

/** A cancelable paste event carrying `text`, as a browser dispatches it. */
function pasteEvent(text: string): ClipboardEvent {
  const e = new Event("paste", { cancelable: true, bubbles: true }) as ClipboardEvent;
  Object.defineProperty(e, "clipboardData", { value: { getData: () => text } });
  return e;
}

/** Mount a grid over `model` and record everything it tells its owner. */
function mount(model: SheetModel) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const notes: { message: string; params?: Record<string, string | number> }[] = [];
  const structure: unknown[] = [];
  let changed = 0;
  mountDatasheet(
    host,
    model,
    () => changed++,
    (change) => structure.push(change),
    (message, params) => notes.push({ message, params }),
  );
  return { host, notes, structure, changed: () => changed };
}

/** Fire a cancelable paste at one cell and report what the sheet did. */
function paste(model: SheetModel, text: string, at: { r: number; c: number }) {
  const m = mount(model);
  const e = pasteEvent(text);
  input(m.host, at.r, at.c).dispatchEvent(e);
  return { ...m, prevented: e.defaultPrevented };
}

/** Type `value` into one cell, as a keystroke does: set it, then fire `input`. */
function typeInto(host: HTMLElement, r: number, c: number, value: string) {
  const el = input(host, r, c);
  el.value = value;
  el.dispatchEvent(new Event("input", { bubbles: true }));
}

// --- the real pane, booted against the real markup (see test/pane-state.test.ts) ----------------

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

const paneCell = (r: number, c: number) =>
  document.querySelector<HTMLInputElement>(`#datasheet input[data-row="${r}"][data-col="${c}"]`)!;

const paneType = (r: number, c: number, value: string) => typeInto($("datasheet"), r, c, value);

/** The series the pane would draw, reduced to what the colour/type side-channels decide. */
const seriesLook = () =>
  exportConfig().data.series.map((s) => ({ name: s.name, color: s.color, type: s.type, stack: s.stack }));

// Spelled as code points: all three look like a space or a hyphen in an editor.
const NBSP = "\u{a0}";
const NNBSP = "\u{202f}";
const MINUS = "\u{2212}";
const RETURN = "\u{21b5}";

/** The first value a single cell reads as once it is in a series row. */
const readCell = (text: string): number | null =>
  sheetToData({
    cells: [
      ["", "C"],
      ["S", text],
    ],
  }).series[0].values[0];

describe("r1-2: numbers written the way Excel DISPLAYS them are numbers, not gaps", () => {
  it("reads an accounting negative in parentheses", () => {
    expect(readCell("(20)")).toBe(-20);
    expect(readCell("(1,234)")).toBe(-1234);
    expect(readCell("(1,234.50)")).toBe(-1234.5);
    // Excel's accounting format pads the currency sign away from the parenthesis.
    expect(readCell("$ (1,234.00)")).toBe(-1234);
    expect(readCell("($1,234)")).toBe(-1234);
  });

  it("reads a currency sign on either side of the number", () => {
    expect(readCell("$1,234")).toBe(1234);
    expect(readCell("$ 1,234.50")).toBe(1234.5);
    expect(readCell("£99")).toBe(99);
    expect(readCell("€1,234")).toBe(1234);
    expect(readCell("1,234 €")).toBe(1234);
    expect(readCell("99€")).toBe(99);
    expect(readCell("-$5")).toBe(-5);
    expect(readCell("$-5")).toBe(-5);
    // Danish Excel's currency format.
    expect(readCell("kr. 1,234")).toBe(1234);
    expect(readCell("1,234 kr.")).toBe(1234);
  });

  it("reads the Unicode minus Excel writes in some locales", () => {
    // U+2212, not the hyphen-minus: spelled out so the two cannot be confused.
    expect(readCell(`${MINUS}20`)).toBe(-20);
    expect(readCell(`${MINUS}1,234`)).toBe(-1234);
  });

  it("reads a space, a no-break space or a narrow no-break space as a thousands group", () => {
    expect(readCell("1 234")).toBe(1234);
    expect(readCell(`1${NBSP}234`)).toBe(1234);
    expect(readCell(`1${NNBSP}234`)).toBe(1234);
    expect(readCell("12 345 678.5")).toBe(12345678.5);
    expect(readCell("(1 234)")).toBe(-1234);
  });

  it("carries a percent sign inside the parentheses", () => {
    expect(readCell("(20%)")).toBe(-20);
  });

  it("still leaves a cell that is not a number as a gap", () => {
    for (const s of [
      "n/a",
      "(20",
      "20)",
      "()",
      "$",
      "(-5)",
      "$5%",
      "$1,234 €",
      "abc 5",
      "5 abc",
      "Q1 234",
      // The refusals the file already pinned: a comma that is not a US group.
      "1.234,5",
      "1,5",
      "1 234,5",
    ]) {
      expect(readCell(s), s).toBeNull();
    }
    // And the plain forms that always read.
    expect(readCell("35%")).toBe(35);
    expect(readCell("1,234")).toBe(1234);
    expect(readCell("-20")).toBe(-20);
    expect(readCell("1e3")).toBe(1000);
  });

  it("is read by a formula over the cell, not just by the cell itself", () => {
    const cells = [
      ["", "A", "B"],
      ["S", "100", "(20)"],
    ];
    expect(evaluateFormula(cells, "SUM(B2:C2)")).toBe(80);
    expect(evaluateFormula(cells, "C2*2")).toBe(-40);
  });

  it("closes a waterfall at the right total when a step is an accounting negative", () => {
    // The finder's case: 100, then a cost shown as (20), then an `e` total. The
    // step used to read as a gap, drop out of the running sum, and the total
    // printed 100 -- a wrong number with nothing said.
    const totals = new Set<number>();
    const data = sheetToData(
      {
        cells: [
          ["", "Start", "Cost", "End"],
          ["S", "100", "(20)", "e"],
        ],
      },
      totals,
    );
    const cfg: ChartConfig = {
      kind: "waterfall",
      width: 480,
      height: 300,
      data,
      waterfall: { totalIndices: [...totals] },
    };
    expect(data.series[0].values).toEqual([100, -20, 0]);
    expect(waterfallChain(cfg)[2].value).toBe(80);
  });

  it("reads a pasted Excel block of currency and accounting cells, and says nothing", () => {
    const model: SheetModel = {
      cells: [
        ["", "Start", "Cost", "End"],
        ["S", "1", "2", "3"],
      ],
    };
    const { notes } = paste(model, ["\tStart\tCost\tEnd", "S\t$100\t(20)\te"].join("\r\n") + "\r\n", { r: 0, c: 0 });
    expect(model.cells[1]).toEqual(["S", "$100", "(20)", "e"]);
    expect(sheetToData(model, new Set()).series[0].values).toEqual([100, -20, 0]);
    expect(notes).toEqual([]);
  });

  it("converts a European block that carries currency, keeping the currency in the cell", () => {
    // `1.234,50 EUR` passed `looksEuropean` but `fromEuropeanNumber` left it
    // unconverted, so it read as a gap while its neighbours were rewritten.
    const model: SheetModel = {
      cells: [
        ["", "Q1"],
        ["S", ""],
        ["T", ""],
        ["U", ""],
        ["V", ""],
      ],
    };
    const { notes } = paste(model, ["€1.234,50", "1.234,50 €", "(1.234,50)", "1 234,5"].join("\r\n") + "\r\n", {
      r: 1,
      c: 1,
    });
    expect(model.cells.slice(1).map((r) => r[1])).toEqual(["€1234.50", "1234.50 €", "(1234.50)", "1234.5"]);
    expect(sheetToData(model).series.map((s) => s.values[0])).toEqual([1234.5, 1234.5, -1234.5, 1234.5]);
    expect(notes).toHaveLength(1);
    expect(notes[0].params).toEqual({ n: 4 });
  });

  it("does not read a US block with currency as European", () => {
    const model: SheetModel = {
      cells: [
        ["", "Q1"],
        ["S", ""],
        ["T", ""],
      ],
    };
    const { notes } = paste(model, ["$1,234.50", "($987.5)"].join("\r\n") + "\r\n", { r: 1, c: 1 });
    expect(model.cells.slice(1).map((r) => r[1])).toEqual(["$1,234.50", "($987.5)"]);
    expect(sheetToData(model).series.map((s) => s.values[0])).toEqual([1234.5, -987.5]);
    expect(notes).toEqual([]);
  });
});

describe("r1-28: a row that becomes (or stops being) a series is reported, so colours stay put", () => {
  // Plan | (blank stack separator) | Actual | Margin -- what `dataToSheet` writes for a
  // clustered-stacked chart, and so what any such chart opens as in the pane.
  const stacked = (): SheetModel => ({
    cells: [
      ["", "Q1", "Q2"],
      ["Plan", "10", "20"],
      ["", "", ""],
      ["Actual", "11", "21"],
      ["Margin", "3", "4"],
    ],
  });

  it("reports typing into a blank separator row as a new series at that position", () => {
    const m = mount(stacked());
    typeInto(m.host, 2, 0, "F");
    expect(m.structure).toEqual([{ kind: "series-insert", index: 1 }]);
    // Typing more into a row that is already a series is not a structure change.
    typeInto(m.host, 2, 0, "Fo");
    typeInto(m.host, 2, 1, "7");
    expect(m.structure).toHaveLength(1);
    // And emptying it again turns it back into the separator it was.
    typeInto(m.host, 2, 0, "");
    typeInto(m.host, 2, 1, "");
    expect(m.structure).toEqual([
      { kind: "series-insert", index: 1 },
      { kind: "series-remove", index: 1 },
    ]);
  });

  it("reports a series row whose every cell was cleared as removed", () => {
    const m = mount(stacked());
    typeInto(m.host, 3, 0, "");
    typeInto(m.host, 3, 1, "");
    expect(m.structure, "a row with a value left in it is still a series").toEqual([]);
    typeInto(m.host, 3, 2, "");
    expect(m.structure).toEqual([{ kind: "series-remove", index: 1 }]);
  });

  it("says nothing for an edit that leaves every row what it was", () => {
    const m = mount(stacked());
    typeInto(m.host, 1, 1, "99");
    typeInto(m.host, 3, 0, "Renamed");
    typeInto(m.host, 0, 1, "H1");
    expect(m.structure).toEqual([]);
  });

  it("reports a paste that blanks a series row, and one that fills the separator", () => {
    const blanked = paste(stacked(), "\t\t\r\n", { r: 3, c: 0 });
    expect(blanked.structure).toEqual([{ kind: "series-remove", index: 1 }]);
    const filled = paste(stacked(), "Z\t5\t6\r\n", { r: 2, c: 0 });
    expect(filled.structure).toEqual([{ kind: "series-insert", index: 1 }]);
    // Values pasted over existing series rows move nothing.
    const values = paste(stacked(), "1\t2\r\n3\t4\r\n", { r: 3, c: 1 });
    expect(values.structure).toEqual([]);
  });

  it("reports the series rows a paste appends, at the end of the series list", () => {
    const m = paste(
      {
        cells: [
          ["", "A"],
          ["S1", "1"],
          ["S2", "2"],
        ],
      },
      ["1", "2", "3", "4"].join("\r\n") + "\r\n",
      { r: 1, c: 1 },
    );
    expect(m.structure).toEqual([
      { kind: "series-insert", index: 2 },
      { kind: "series-insert", index: 3 },
    ]);
  });

  describe("in the pane, where colour and combo type hang off the series by position", () => {
    const combo = (): Partial<ChartConfig> => ({
      kind: "combo",
      width: 480,
      height: 300,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "Plan", values: [10, 20], color: "#111111", stack: 0 },
          { name: "Actual", values: [11, 21], color: "#222222", stack: 1 },
          { name: "Margin", values: [3, 4], color: "#ff0000", type: "line", stack: 1 },
        ],
      },
    });

    it("keeps every colour and the line type on their own series when the separator gets a name", async () => {
      await bootPane();
      importConfig(combo());
      expect(paneCell(2, 0).value, "the separator row is not where this test expects it").toBe("");
      paneType(2, 0, "F");
      // Before the fix: Actual wore Margin's red and line type, and Margin had neither.
      const look = seriesLook();
      expect(look.map((s) => s.name)).toEqual(["Plan", "F", "Actual", "Margin"]);
      expect(look[0].color).toBe("#111111");
      expect(look[1].color, "the new series took its neighbour's colour").toBeUndefined();
      expect(look[2]).toMatchObject({ name: "Actual", color: "#222222", type: undefined });
      expect(look[3]).toMatchObject({ name: "Margin", color: "#ff0000", type: "line" });
    });

    it("keeps them when a cleared row turns a series into a separator", async () => {
      await bootPane();
      importConfig({
        kind: "combo",
        width: 480,
        height: 300,
        data: {
          categories: ["Q1", "Q2"],
          series: [
            { name: "Rev", values: [10, 20], color: "#111111" },
            { name: "Cost", values: [5, 6], color: "#222222" },
            { name: "Margin", values: [3, 4], color: "#ff0000", type: "line" },
          ],
        },
      });
      for (const c of [0, 1, 2]) paneType(2, c, "");
      const look = seriesLook();
      expect(look.map((s) => s.name)).toEqual(["Rev", "Margin"]);
      expect(look[0].color).toBe("#111111");
      // Before the fix Margin lost its line type and its red, and took Cost's colour.
      expect(look[1]).toMatchObject({ color: "#ff0000", type: "line" });
    });

    it("keeps them when a pasted blank row overwrites a series row", async () => {
      await bootPane();
      importConfig(combo());
      const e = pasteEvent("\t\t\r\n");
      paneCell(3, 0).dispatchEvent(e);
      const look = seriesLook();
      expect(look.map((s) => s.name)).toEqual(["Plan", "Margin"]);
      expect(look[1]).toMatchObject({ color: "#ff0000", type: "line" });
    });

    it("gives the colours back with the row on undo", async () => {
      await bootPane();
      importConfig(combo());
      paneType(2, 0, "F");
      document.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true }));
      const look = seriesLook();
      expect(look.map((s) => s.name)).toEqual(["Plan", "Actual", "Margin"]);
      expect(look[1]).toMatchObject({ color: "#222222" });
      expect(look[2]).toMatchObject({ color: "#ff0000", type: "line" });
    });
  });
});

describe("r1-12: a Gantt '% Complete' row has ONE unit, decided by the row", () => {
  /** What fraction of each bar the progress fill covers, drawn from the grid as typed. */
  const drawn = (cellsOfRow: string[]) => {
    const sheet: SheetModel = {
      cells: [
        ["", "A", "B", "C"],
        ["Start", "1", "4", "9"],
        ["End", "4", "10", "13"],
        ["% Complete", ...cellsOfRow],
      ],
    };
    const data = sheetToData(sheet);
    const scene = buildChart(normalizeConfig({ kind: "gantt", width: 600, height: 240, data }));
    const share = (c: number) => {
      const bar = scene.nodes.find((n) => n.name === `bar-${c}`) as RectNode;
      const fill = scene.nodes.find((n) => n.name === `progress-${c}`) as RectNode | undefined;
      return fill ? fill.w / bar.w : 0;
    };
    return { parsed: data.series[2].values, shares: [0, 1, 2].map(share) };
  };

  it("draws a typed 1% as 1%, not as complete, when the row is written in percent", () => {
    const { parsed, shares } = drawn(["100%", "40%", "1%"]);
    expect(parsed).toEqual([100, 40, 1]);
    // Before: [1, 0.4, 1] -- the task that had just started drew as finished.
    expect(shares[0]).toBeCloseTo(1, 5);
    expect(shares[1]).toBeCloseTo(0.4, 5);
    expect(shares[2]).toBeCloseTo(0.01, 5);
  });

  it("draws a fractional percent as that percent", () => {
    expect(drawn(["100%", "40%", "0.5%"]).shares[2]).toBeCloseTo(0.005, 5);
  });

  it("draws 0%, 1%, 2% in increasing order", () => {
    const { shares } = drawn(["0%", "1%", "2%"]);
    expect(shares[0]).toBe(0);
    expect(shares[1]).toBeCloseTo(0.01, 5);
    expect(shares[2]).toBeCloseTo(0.02, 5);
    expect(shares[1]).toBeLessThan(shares[2]);
  });

  it("reads a row whose values are all within 0-1 as fractions", () => {
    const { shares } = drawn(["1", "0.4", "0.01"]);
    expect(shares[0]).toBeCloseTo(1, 5);
    expect(shares[1]).toBeCloseTo(0.4, 5);
    expect(shares[2]).toBeCloseTo(0.01, 5);
  });

  it("reads a lone '1%' as a fraction -- the one reading the row cannot settle", () => {
    // The '%' is stripped before layout, so "1%" and "1" arrive as the same
    // number and a row of only such values cannot say which it meant. Fractions
    // is the documented answer (docs/MANUAL.md, Gantt), and this pins it.
    expect(drawn(["1%", "", ""]).shares[0]).toBeCloseTo(1, 5);
  });

  it("clamps once the row is in percent, whatever passes 100", () => {
    const { shares } = drawn(["0.5", "150", ""]);
    expect(shares[0]).toBeCloseTo(0.005, 5);
    expect(shares[1]).toBeCloseTo(1, 5);
  });
});

describe("r1-53: MIN and MAX fail on a text cell or a cycle wherever it sits, like SUM and AVG", () => {
  /** Row 3 holds `values` (B3..), and B2 holds the formula under test. */
  const over = (formula: string, values: string[]): number | null =>
    sheetToData({
      cells: [
        ["", "R"],
        ["S", formula],
        ["row", ...values],
      ],
    }).series[0].values[0];

  const orders = [
    ["n/a", "10", "20"],
    ["10", "n/a", "20"],
    ["10", "20", "n/a"],
  ];

  it.each(["MIN", "MAX", "SUM", "AVG"])("%s over a range with a text cell is an error in every position", (fn) => {
    for (const values of orders) {
      // Before the fix MIN and MAX answered 10 / 20 whenever the text was not first.
      expect(over(`=${fn}(B3:D3)`, values), `${fn} ${values.join(",")}`).toBeNull();
    }
  });

  it.each(["MIN", "MAX"])("%s over bare references fails the same way", (fn) => {
    for (const values of orders) {
      expect(over(`=${fn}(B3,C3,D3)`, values), `${fn} ${values.join(",")}`).toBeNull();
    }
  });

  it("still ignores a blank, and still answers when every cell is a number", () => {
    expect(over("=MIN(B3:D3)", ["10", "", "20"])).toBe(10);
    expect(over("=MAX(B3:D3)", ["-5", "", "-8"])).toBe(-5);
    expect(over("=MIN(B3:D3)", ["10", "5", "20"])).toBe(5);
    expect(over("=MAX(B3:D3)", ["10", "5", "20"])).toBe(20);
    expect(over("=MAX(B3:D3,3)", ["(10)", "5", "1"])).toBe(5);
  });

  it("is a cycle, not a number, when the loop runs through MIN or MAX", () => {
    const series = (b: string, c: string) =>
      sheetToData({
        cells: [
          ["", "R", "S"],
          ["A", b, c],
        ],
      }).series[0].values;
    // B2 -> C2 -> B2. MIN(5, NaN) used to be 5, and C2 = B2 followed it.
    expect(series("=MIN(5,C2)", "=B2")).toEqual([null, null]);
    expect(series("=MAX(5,C2)", "=B2")).toEqual([null, null]);
    expect(series("=MIN(C2:C2,5)", "=B2")).toEqual([null, null]);
  });

  it("is a cycle for a self-reference in either argument order", () => {
    const self = (formula: string) =>
      sheetToData({
        cells: [
          ["", "R"],
          ["A", formula],
        ],
      }).series[0].values[0];
    // `MIN(5,B2)` was 5 and `MIN(B2,5)` was null: the answer depended on which side the cycle stood.
    expect(self("=MIN(5,B2)")).toBeNull();
    expect(self("=MIN(B2,5)")).toBeNull();
    expect(self("=MAX(5,B2)")).toBeNull();
    expect(self("=MAX(B2,5)")).toBeNull();
  });

  it("still folds a range wider than the engine's argument limit", () => {
    // 399 x 399 = 159k numbers, past V8's ~125k call arguments: why MIN and MAX
    // are folded (`minOf`/`maxOf`) and not spread. Taking NaN seriously must not
    // bring `Math.max(...nums)` back.
    const side = 400;
    const cells = Array.from({ length: side }, (_, r) =>
      Array.from({ length: side }, (_, c) =>
        r === 0 ? (c === 0 ? "" : `H${c}`) : c === 0 ? `S${r}` : String((r * c) % 97),
      ),
    );
    expect(evaluateFormula(cells, "MAX(B2:OJ400)")).toBe(96);
    expect(evaluateFormula(cells, "MIN(B2:OJ400)")).toBe(0);
  });
});

describe("r1-54: the paste gate asks what the clipboard holds, not which control characters it contains", () => {
  const grid = (): SheetModel => ({
    cells: [
      ["", "A", "B"],
      ["Revenue ", "1", "2"],
      ["S2", "3", "4"],
    ],
  });

  it("leaves one copied Excel cell to the browser, so it lands at the caret", () => {
    // Desktop Excel writes a single copied cell as "value\r\n". That contains "\n",
    // so the old gate took it over and replaced the cell's WHOLE content -- the
    // "Revenue " the user had already typed included -- while the same text copied
    // from anywhere else (no terminator) went to the browser and was inserted.
    for (const text of ["(EURm)\r\n", "(EURm)\n", "(EURm)\r", "(EURm)"]) {
      const model = grid();
      const m = paste(model, text, { r: 1, c: 0 });
      expect(m.prevented, JSON.stringify(text)).toBe(false);
      expect(m.changed(), JSON.stringify(text)).toBe(0);
      expect(model.cells[1][0]).toBe("Revenue ");
    }
  });

  it("still converts one copied cell that is a European number", () => {
    // Not a casualty of the gate: a Danish Excel's "987,5" copied as ONE cell
    // carried a terminator, was taken over and rewritten, and the browser's
    // default would have left it as a gap.
    const model = grid();
    const m = paste(model, "987,5\r\n", { r: 2, c: 1 });
    expect(m.prevented).toBe(true);
    expect(model.cells[2][1]).toBe("987.5");
    expect(m.notes).toHaveLength(1);
  });

  it("leaves a cell with a stray quote in it alone too", () => {
    // `5" pipe` is data: a quote only opens a cell at its start.
    const m = paste(grid(), '5" pipe\r\n', { r: 1, c: 0 });
    expect(m.prevented).toBe(false);
  });

  it("still takes over a single cell the grid parser has to interpret", () => {
    // A quoted cell with an Alt+Enter break in it is ONE cell, but the browser
    // would paste the quotes and turn the break into a space.
    const model = grid();
    const m = paste(model, '"Revenue\n(EURm)"\r\n', { r: 1, c: 0 });
    expect(m.prevented).toBe(true);
    expect(model.cells[1][0]).toBe("Revenue\n(EURm)");
  });

  it("still takes over a row of cells, a column of cells and a block", () => {
    const a = grid();
    expect(paste(a, "9\t8\r\n", { r: 1, c: 1 }).prevented).toBe(true);
    expect(a.cells[1]).toEqual(["Revenue ", "9", "8"]);
    const b = grid();
    expect(paste(b, "7\r\n6\r\n", { r: 1, c: 1 }).prevented).toBe(true);
    expect([b.cells[1][1], b.cells[2][1]]).toEqual(["7", "6"]);
  });

  it("splits a column written with lone carriage returns, as older Mac Excel writes it", () => {
    // `parseClipboardGrid` has always read a lone CR as a row break, but the old
    // gate looked for "\t" and "\n" and never let such a column reach it. The
    // browser then joined the three values into one cell as "10 20 30", which
    // reads as a gap while 20 and 30 never reached their rows.
    const model = grid();
    const m = paste(model, "10\r20\r30\r", { r: 1, c: 1 });
    expect(m.prevented).toBe(true);
    expect(model.cells.slice(1).map((r) => r[1])).toEqual(["10", "20", "30"]);
  });

  it("splits a CR-only block with tabs, as it always did", () => {
    const model = grid();
    paste(model, "10\t11\r20\t21\r", { r: 1, c: 1 });
    expect(model.cells.slice(1, 3).map((r) => r.slice(1))).toEqual([
      ["10", "11"],
      ["20", "21"],
    ]);
  });

  it("leaves plain text with no terminator to the browser", () => {
    expect(paste(grid(), "just text", { r: 1, c: 1 }).prevented).toBe(false);
  });
});

describe("r1-56: a line break inside a cell is shown, and survives an edit of that cell", () => {
  // The Alt+Enter header a finance table is full of, as Excel's clipboard writes it.
  const wrapped = '"Revenue\n(EURm)"\t2024\r\n';

  it("shows the break, instead of running the two lines together", () => {
    const model: SheetModel = {
      cells: [
        ["", "A"],
        ["S", "1"],
      ],
    };
    const m = paste(model, wrapped, { r: 0, c: 1 });
    expect(m.prevented).toBe(true);
    expect(model.cells[0][1]).toBe("Revenue\n(EURm)");
    // An <input type=text> strips a line break from its value, so the grid used
    // to show "Revenue(EURm)": the words glued together over a label that holds two lines.
    expect(input(m.host, 0, 1).value).toBe(`Revenue${RETURN}(EURm)`);
  });

  it("keeps the break when the cell is edited", () => {
    const model: SheetModel = {
      cells: [
        ["", "A"],
        ["S", "1"],
      ],
    };
    const m = paste(model, wrapped, { r: 0, c: 1 });
    typeInto(m.host, 0, 1, `${input(m.host, 0, 1).value}x`);
    // Before: "Revenue(EURm)x" -- the break gone for good, the words glued.
    expect(model.cells[0][1]).toBe("Revenue\n(EURm)x");
    expect(sheetToData(model).categories[0]).toBe("Revenue\n(EURm)x");
  });

  it("leaves the break alone when a different cell is edited", () => {
    const model: SheetModel = {
      cells: [
        ["", "A"],
        ["S", "1"],
      ],
    };
    const m = paste(model, wrapped, { r: 0, c: 1 });
    typeInto(m.host, 1, 1, "5");
    expect(model.cells[0][1]).toBe("Revenue\n(EURm)");
  });

  it("lets the user take the break out by deleting what stands for it", () => {
    const model: SheetModel = {
      cells: [
        ["", "A"],
        ["S", "1"],
      ],
    };
    const m = paste(model, wrapped, { r: 0, c: 1 });
    typeInto(m.host, 0, 1, "Revenue (EURm)");
    expect(model.cells[0][1]).toBe("Revenue (EURm)");
  });

  it("shows a break that arrived from a loaded chart, and a CRLF as one", () => {
    const model: SheetModel = {
      cells: [
        ["", "Q1\n2024", "Q2\r\n2024"],
        ["S", "1", "2"],
      ],
    };
    const { host } = mount(model);
    expect(input(host, 0, 1).value).toBe(`Q1${RETURN}2024`);
    expect(input(host, 0, 2).value).toBe(`Q2${RETURN}2024`);
    typeInto(host, 0, 1, `${input(host, 0, 1).value}x`);
    expect(model.cells[0][1]).toBe("Q1\n2024x");
  });

  it("still draws a grid whose model holds a null cell, as an undo round trip can leave", () => {
    // JSON turns an array hole into null, and `input.value = null` was an empty
    // field; swapping the assignment must not turn it into a TypeError.
    const model = {
      cells: [
        ["", null],
        ["S", "1"],
      ],
    } as unknown as SheetModel;
    const { host } = mount(model);
    expect(input(host, 0, 1).value).toBe("");
  });

  it("changes nothing for a cell that holds no break", () => {
    const model: SheetModel = {
      cells: [
        ["", "Q1"],
        ["Plain", "1"],
      ],
    };
    const { host } = mount(model);
    expect(input(host, 0, 1).value).toBe("Q1");
    typeInto(host, 1, 0, "Plain 2");
    expect(model.cells[1][0]).toBe("Plain 2");
  });

  it("carries a loaded label's break through the pane's edit", async () => {
    await bootPane();
    importConfig({
      kind: "clustered",
      width: 480,
      height: 300,
      data: { categories: ["Q1\n2024", "Q2"], series: [{ name: "S", values: [1, 2] }] },
    });
    expect(paneCell(0, 1).value).toBe(`Q1${RETURN}2024`);
    paneType(0, 1, `${paneCell(0, 1).value}x`);
    expect(exportConfig().data.categories[0]).toBe("Q1\n2024x");
  });
});

describe("r1-29: Ctrl+Z / Ctrl+Y in a text box elsewhere in the pane undoes that box, not the datasheet", () => {
  const keydown = (target: EventTarget, init: KeyboardEventInit = {}) => {
    const e = new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true, cancelable: true, ...init });
    target.dispatchEvent(e);
    return e;
  };
  const firstValue = () => exportConfig().data.series[0].values[0];

  /** A loaded chart whose datasheet has been edited once, so there is something to undo. */
  async function editedPane() {
    await bootPane();
    importConfig({
      kind: "clustered",
      width: 480,
      height: 300,
      data: { categories: ["A", "B"], series: [{ name: "S1", values: [1, 2] }] },
    });
    paneType(1, 1, "42");
    expect(firstValue(), "the edit did not land, so there is nothing to undo").toBe(42);
  }

  // The agenda box, the chart title, the JSON box, a KPI field, the type search
  // and a number spinner: none of them is the datasheet, and the first three are
  // on tabs where the datasheet is not even in view.
  it.each(["agenda-chapters", "chart-title", "json-io", "kpi-label", "type-search-input", "flow-highlight"])(
    "leaves the keystroke to #%s",
    async (id) => {
      await editedPane();
      const field = $(id);
      field.focus();
      const e = keydown(field);
      // Before the fix the datasheet went back to [1, 2] and the field's own undo was blocked.
      expect(e.defaultPrevented, "the pane cancelled the field's own undo").toBe(false);
      expect(firstValue(), "the pane undid the datasheet instead").toBe(42);
    },
  );

  it("does the same for Cmd+Z, and for redo", async () => {
    await editedPane();
    const field = $("agenda-chapters");
    expect(keydown(field, { ctrlKey: false, metaKey: true }).defaultPrevented).toBe(false);
    // Undo from the grid to make a redo available, then ask for it from the text box.
    expect(keydown(paneCell(1, 1)).defaultPrevented).toBe(true);
    expect(firstValue()).toBe(1);
    expect(keydown(field, { key: "y" }).defaultPrevented).toBe(false);
    expect(firstValue(), "redo ran from a text box").toBe(1);
    // And it is still there for the grid.
    expect(keydown(paneCell(1, 1), { key: "y" }).defaultPrevented).toBe(true);
    expect(firstValue()).toBe(42);
  });

  it("still undoes the datasheet from a cell, from a checkbox and from nowhere in particular", async () => {
    await editedPane();
    expect(keydown(paneCell(1, 1)).defaultPrevented).toBe(true);
    expect(firstValue()).toBe(1);
    keydown(paneCell(1, 1), { key: "y" });
    expect(firstValue()).toBe(42);
    expect(keydown($("auto-update")).defaultPrevented).toBe(true);
    expect(firstValue()).toBe(1);
    keydown(document.body, { key: "y" });
    expect(firstValue()).toBe(42);
    expect(keydown(document.body).defaultPrevented).toBe(true);
    expect(firstValue()).toBe(1);
  });
});

describe("r1-55: 'Datamark axis' and 'Value axis' are two boxes over one field, and say the same thing", () => {
  const boxFor = (label: string) =>
    [...document.querySelectorAll<HTMLLabelElement>("#options label")]
      .filter((l) => l.textContent?.trim() === label)
      .map((l) => l.querySelector<HTMLInputElement>("input[type=checkbox]"))[0]!;
  const DATAMARK = "Datamark axis (ticks only)";
  const valueAxisNodes = () => buildChart(exportConfig()).nodes.filter((n) => n.name?.startsWith("value-axis")).length;

  it("gives a plain value axis back when 'Datamark axis' is ticked and unticked again", async () => {
    await bootPane();
    importConfig(sampleConfig("line"));
    expect(exportConfig().decorations?.valueAxis, "the sample no longer sets valueAxis").toBe(true);
    const before = valueAxisNodes();
    expect(before, "the sample draws no value axis, so this proves nothing").toBeGreaterThan(0);

    boxFor(DATAMARK).click();
    expect(exportConfig().decorations?.valueAxis).toBe("datamarks");
    boxFor(DATAMARK).click();
    // Before the fix: false -- the axis the chart arrived with was gone, and the
    // "Value axis" box read unticked.
    expect(exportConfig().decorations?.valueAxis).toBe(true);
    expect(boxFor("Value axis").checked).toBe(true);
    expect(boxFor(DATAMARK).checked).toBe(false);
    expect(valueAxisNodes()).toBe(before);
  });

  it("leaves a chart that had no value axis without one", async () => {
    await bootPane();
    importConfig({
      kind: "stacked",
      data: { categories: ["A", "B"], series: [{ name: "S1", values: [10, 20] }] },
    });
    boxFor(DATAMARK).click();
    boxFor(DATAMARK).click();
    expect(exportConfig().decorations?.valueAxis).toBe(false);
    expect(boxFor("Value axis").checked).toBe(false);
  });

  it("does not carry one chart's axis into the next chart it loads", async () => {
    await bootPane();
    importConfig(sampleConfig("line"));
    boxFor(DATAMARK).click();
    // A different chart, which arrives already in datamark style: there is no
    // "before" for it to go back to.
    importConfig({
      kind: "line",
      data: { categories: ["A", "B"], series: [{ name: "S1", values: [10, 20] }] },
      decorations: { valueAxis: "datamarks" },
    });
    expect(boxFor(DATAMARK).checked).toBe(true);
    boxFor(DATAMARK).click();
    expect(exportConfig().decorations?.valueAxis).toBe(false);
  });

  it("unticks 'Datamark axis' when 'Value axis' is unticked over it", async () => {
    await bootPane();
    importConfig({
      kind: "line",
      data: { categories: ["A", "B"], series: [{ name: "S1", values: [10, 20] }] },
      decorations: { valueAxis: "datamarks" },
    });
    expect(boxFor("Value axis").checked).toBe(true);
    expect(boxFor(DATAMARK).checked).toBe(true);
    boxFor("Value axis").click();
    expect(exportConfig().decorations?.valueAxis).toBe(false);
    // Before the fix this box stayed ticked over a chart with no value axis at all.
    expect(boxFor(DATAMARK).checked).toBe(false);
  });
});
