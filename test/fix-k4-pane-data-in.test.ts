// @vitest-environment jsdom
import { afterAll, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { buildChart, normalizeConfig } from "../src/core/chart";
import { parseDateToken } from "../src/core/format";
import type { SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { cellText, trimBlankEdges } from "../src/excel/cell-text";
import { excelSerialToISO } from "../src/excel/serial-dates";
import {
  dataToSheet,
  evaluateFormula,
  fromEuropeanNumber,
  looksEuropean,
  parseClipboardGrid,
  sheetToData,
  splitValueList,
  typedNumber,
} from "../src/taskpane/datasheet";

/**
 * THE THIRD BUG HUNT, STREAM k4: what the datasheet, the Excel bridge and the
 * pane's own boxes make of text and numbers. One describe per finding.
 *
 * The pane tests boot the real `app.ts` against `taskpane.html` the way
 * `pane-state.test.ts` does; the rest are pure functions and need no DOM.
 */

/** See `pane-state.test.ts`: the pane's last debounced render must not outlive jsdom. */
afterAll(async () => {
  await new Promise((r) => setTimeout(r, 250));
});

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Boot the pane on a clean browser, with `stored` already in localStorage. */
async function bootPane(stored: Record<string, string> = {}) {
  try {
    window.localStorage.clear();
    for (const [k, v] of Object.entries(stored)) window.localStorage.setItem(k, v);
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
const note = () => $("host-note").textContent ?? "";
const cell = (r: number, c: number) =>
  document.querySelector<HTMLInputElement>(`#datasheet input[data-row="${r}"][data-col="${c}"]`)!;
const typeInto = (el: HTMLInputElement, value: string) => {
  el.value = value;
  el.dispatchEvent(new Event("input", { bubbles: true }));
};
/** The two text boxes of the options row that starts with `prefix` ("Axis scale min", "Axis break from"). */
const boxes = (prefix: string): HTMLInputElement[] => {
  const row = [...document.querySelectorAll<HTMLElement>("#options label")].find((l) =>
    (l.textContent ?? "").startsWith(prefix),
  );
  if (!row) throw new Error(`no options row starting "${prefix}"`);
  return [...row.querySelectorAll<HTMLInputElement>('input[type="text"]')];
};
const ctrl = (key: string) =>
  document.dispatchEvent(new KeyboardEvent("keydown", { key, ctrlKey: true, bubbles: true }));

const base = { categories: ["A", "B", "C"], series: [{ name: "S1", values: [1, 2, 3] }] };

/** An Excel serial for a calendar day (1900 date system), and the epoch day the pane's date reader gives it. */
const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
const epochDay = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d) / 86400000;
const header = (value: number, format: string, shown: string) =>
  cellText(value, format, shown, { label: true, date1904: false });

/* ───────────────────────────────── r3-2 ───────────────────────────────────── */

describe("a date header the workbook DISPLAYS day-first is delivered as the day the cell holds (r3-2)", () => {
  it("replaces a display the date reader would read as another day with the serial's ISO date", () => {
    // Danish short date. 5 January reads as 1 May month-first, 10 March as 3 October.
    expect(header(serial(2026, 1, 5), "dd-mm-yyyy", "05-01-2026")).toBe("2026-01-05");
    expect(header(serial(2026, 3, 10), "dd-mm-yyyy", "10-03-2026")).toBe("2026-03-10");
    // The built-in short date, as the file stores it and a Danish Excel prints it.
    expect(header(serial(2026, 4, 7), "m/d/yyyy", "07-04-2026")).toBe("2026-04-07");
    expect(header(serial(2026, 1, 5), "dd-mm-yy", "05-01-26")).toBe("2026-01-05");
    // 5-1-2026 (nl-NL): one-digit day and month.
    expect(header(serial(2026, 1, 5), "d-m-yyyy", "5-1-2026")).toBe("2026-01-05");
  });

  it("whatever it delivers, the date reader reads the day the cell holds", () => {
    // 15 January cannot be a month, so the parser has no month-first reading to
    // get wrong; it either refuses the text or (once it reads day-first) reads it
    // right. Either way what reaches the chart must be the 15th.
    for (const [y, m, d] of [
      [2026, 1, 15],
      [2026, 12, 31],
      [2026, 1, 5],
      [2026, 10, 3],
      [2026, 7, 4],
    ] as const) {
      const shown = `${String(d).padStart(2, "0")}-${String(m).padStart(2, "0")}-${y}`;
      expect(parseDateToken(header(serial(y, m, d), "dd-mm-yyyy", shown)), shown).toBe(epochDay(y, m, d));
    }
  });

  it("keeps the displayed text when it names the day, at the resolution the format shows", () => {
    // `mmm-yy` prints "Jan-26" for every day of January, and the reader takes it as
    // the 1st: a month-END header (EOMONTH) must not become "2026-01-31".
    expect(header(serial(2026, 1, 1), "mmm-yy", "Jan-26")).toBe("Jan-26");
    expect(header(serial(2026, 1, 31), "mmm-yy", "Jan-26")).toBe("Jan-26");
    expect(header(serial(2026, 2, 28), "mmmm yyyy", "February 2026")).toBe("February 2026");
    // No year in the format, so none is demanded of the text.
    expect(header(serial(2026, 3, 15), "mmmm", "March")).toBe("March");
    // The day IS shown, so it must agree.
    expect(header(serial(2026, 1, 5), "d-mmm-yy", "5-Jan-26")).toBe("5-Jan-26");
    expect(header(serial(2026, 1, 5), "yyyy-mm-dd", "2026-01-05")).toBe("2026-01-05");
    expect(header(serial(2026, 1, 15), "dd.mm.yyyy", "15.01.2026")).toBe("15.01.2026");
  });

  it("keeps text that is not a date reading at all", () => {
    // `"Q"1 yyyy`-style formats print a label; a year-only format prints a number.
    expect(header(serial(2026, 1, 1), '"Q1 "yyyy', "Q1 2026")).toBe("Q1 2026");
    expect(header(serial(2026, 1, 1), "yyyy", "2026")).toBe("2026");
    expect(header(serial(2026, 1, 5), "dddd", "Monday")).toBe("Monday");
  });

  it("replaces a display the reader misreads outright (mm/yy is read as a day of 2001)", () => {
    expect(header(serial(2026, 1, 1), "mm/yy", "01/26")).toBe("2026-01-01");
  });

  it("is unchanged for body cells, which were always ISO", () => {
    expect(cellText(serial(2026, 1, 5), "dd-mm-yyyy", "05-01-2026", { label: false, date1904: false })).toBe(
      "2026-01-05",
    );
  });

  it("charts three irregular Danish dates on a proportional axis, as the ISO header did before c668a1d", () => {
    const heads = [
      [serial(2026, 1, 5), "05-01-2026"],
      [serial(2026, 3, 10), "10-03-2026"],
      [serial(2026, 4, 7), "07-04-2026"],
    ] as const;
    const grid: unknown[][] = [
      ["", ...heads.map(([v]) => v)],
      ["Sales", 10, 20, 30],
    ];
    const fmt = [
      ["General", ...heads.map(() => "dd-mm-yyyy")],
      ["General", "General", "General", "General"],
    ];
    const shown = [
      ["", ...heads.map(([, s]) => s)],
      ["Sales", "10", "20", "30"],
    ];
    const cells = trimBlankEdges(
      grid.map((row, r) =>
        row.map((v, c) => cellText(v, fmt[r][c], shown[r][c], { label: r === 0 || c === 0, date1904: false })),
      ),
    );
    const data = sheetToData({ cells });
    expect(data.categories).toEqual(["2026-01-05", "2026-03-10", "2026-04-07"]);
    const texts = buildChart(normalizeConfig({ kind: "line", data } as ChartConfig)).nodes.filter(
      (n: SceneNode): n is TextNode => n.kind === "text" && data.categories.includes(n.text),
    );
    const mid = (t: string) => {
      const n = texts.find((x) => x.text === t)!;
      return n.x + n.w / 2;
    };
    const gap1 = mid("2026-03-10") - mid("2026-01-05");
    const gap2 = mid("2026-04-07") - mid("2026-03-10");
    // 64 days against 28: the second gap is well under half the first.
    expect(gap1 / gap2).toBeCloseTo(64 / 28, 1);
  });
});

/* ───────────────────────────────── r3-1 ───────────────────────────────────── */

describe("number-shaped text the number reader refuses is a gap, never a date (r3-1)", () => {
  const valuesOf = (...cellsOfRow: string[]) => {
    const data = sheetToData({
      cells: [
        ["", ...cellsOfRow.map((_, i) => `C${i}`)],
        ["Row", ...cellsOfRow],
      ],
    });
    return { values: data.series[0].values, dates: data.dates };
  };

  it("reads a trailing minus as the negative it is (SAP and ERP list exports)", () => {
    expect(valuesOf("120", "45-", "80")).toEqual({ values: [120, -45, 80], dates: undefined });
    expect(valuesOf("100-", "1.5-", "1,234-", "−7-")).toEqual({ values: [-100, -1.5, -1234, null], dates: undefined });
    // Beside the other accounting spellings, and with the dressing the reader knows.
    expect(valuesOf("(45)", "-45", "45-").values).toEqual([-45, -45, -45]);
    expect(valuesOf("45-%", "$45-", "45-€").values).toEqual([-45, -45, -45]);
  });

  it("keeps a trailing minus honest: one sign, adjoining the digits", () => {
    for (const t of ["-45-", "(45)-", "(45-)", "45 -", "45-45-", "45--", "+45-"]) {
      const { values, dates } = valuesOf(t);
      expect(values, `${t} was read as a number`).toEqual([null]);
      expect(dates, `${t} was read as a date`).toBeUndefined();
    }
  });

  it("does not turn fractions, ratios, pairs or an unplaceable European comma into dates", () => {
    // Each of these used to chart as a day in 2001 or year 100 AND flip the whole
    // ChartData into date mode.
    for (const t of ["1/2", "3/4", "1:100", "5 12", "1,5-", "1.5.", "12:30", "1/2 3", "1-2-"]) {
      const { values, dates } = valuesOf(t);
      expect(values, `${t} became a value`).toEqual([null]);
      expect(dates, `${t} flipped the chart into date mode`).toBeUndefined();
    }
  });

  it("still reads every spelling of a real date", () => {
    for (const t of [
      "2026-01-15",
      "2026-01",
      "15.01.2026",
      "01/15/2026",
      "1/1/2026",
      "2026/01/15",
      "2026-01-15 10:30",
      "2026-01-15T10:30:00Z",
      "Jan 2026",
      "15 Jan 2026",
      "Mon, 05 Jan 2026",
    ]) {
      const { values, dates } = valuesOf(t);
      expect(values[0], `${t} stopped being a date`).not.toBeNull();
      expect(dates, t).toBe(true);
    }
  });

  it("reads the same through a formula, and through a European paste", () => {
    const cells = [
      ["", "A", "B", "C"],
      ["Row", "45-", "=B2+30", "=B2*2"],
    ];
    expect(sheetToData({ cells }).series[0].values).toEqual([-45, -15, -90]);
    // "1.234,5-" in a block that is European: the comma is a decimal, the minus survives.
    const grid = parseClipboardGrid("\tQ1\tQ2\r\nNet\t1.234,5-\t987,5\r\n");
    expect(looksEuropean(grid)).toBe(true);
    const fixed = grid.map((row) => row.map(fromEuropeanNumber));
    expect(sheetToData({ cells: fixed }).series[0].values).toEqual([-1234.5, 987.5]);
  });
});

/* ───────────────────────────────── r3-22 ───────────────────────────────────── */

describe("the number reader takes the groupings and currencies Excel prints in more places (r3-22)", () => {
  const read = (t: string) =>
    sheetToData({
      cells: [
        ["", "X"],
        ["Row", t],
      ],
    }).series[0].values[0];

  it("reads the Swiss apostrophe grouping (de-CH, fr-CH, it-CH)", () => {
    expect(read("1'234.50")).toBe(1234.5);
    expect(read("1’234")).toBe(1234);
    expect(read("-1'000")).toBe(-1000);
    expect(read("2'500.50")).toBe(2500.5);
    expect(read("12'345'678")).toBe(12345678);
    // A word in front is still a word.
    expect(read("Fr. 1'234.50")).toBeNull();
  });

  it("reads the Indian lakh grouping (en-IN, hi-IN)", () => {
    expect(read("12,34,567.89")).toBe(1234567.89);
    expect(read("1,00,000")).toBe(100000);
    expect(read("1,23,456")).toBe(123456);
    expect(read("1,23,45,678")).toBe(12345678);
    // The US-shaped Indian numbers always worked, and a plain decimal comma is still refused.
    expect(read("12,345")).toBe(12345);
    expect(read("12,34")).toBeNull();
    expect(read("123,45,678")).toBeNull();
  });

  it("reads a currency sign that carries letters (R$, US$, NZ$)", () => {
    expect(read("US$ 5")).toBe(5);
    expect(read("US$ 5")).toBe(5);
    expect(read("NZ$1,234.50")).toBe(1234.5);
    expect(read("(R$ 20)")).toBe(-20);
    expect(read("12 R$")).toBe(12);
  });

  it("does not make a word, a bare code or an incoherent dressing into a number", () => {
    for (const t of [
      "USD 5",
      "abc 5",
      "R 5",
      "R$$ 5",
      "US$ 5%",
      "R$ 5 US$",
      "1'23",
      "1,2,3",
      "12'3456",
      "(a)",
      "2024-05x",
    ]) {
      expect(read(t), t).toBeNull();
    }
  });

  it("reads Arabic-Indic and Persian digits and their separators", () => {
    expect(read("٢٥")).toBe(25); // ٢٥
    expect(read("٣٫٥")).toBe(3.5); // ٣٫٥
    expect(read("١٬٢٣٤")).toBe(1234); // ١٬٢٣٤
    expect(read("(٢٠)")).toBe(-20);
    expect(read("۱۲")).toBe(12); // ۱۲
    expect(read("‎-٥")).toBe(-5);
    expect(read("٢٥٪؜")).toBe(25); // ٢٥٪؜ — a share, as Intl prints it
  });

  it("lets a pt-BR paste be read as European when its currency cells are the evidence", () => {
    // The block from the report: nothing but the "R$" cells carries a decimal comma.
    // A reader that does not understand "R$" casts no vote, and the dot-grouped
    // integers beside it were read 1000x too small with no note.
    const grid = parseClipboardGrid(
      "\tJan\tFeb\tMar\r\nReceita\tR$ 1.234,50\tR$ 2.500,00\tR$ 99,90\r\nUnidades\t1.234\t2.500\t987\r\n",
    );
    expect(looksEuropean(grid)).toBe(true);
    const data = sheetToData({ cells: grid.map((row) => row.map(fromEuropeanNumber)) });
    expect(data.series.map((s) => s.values)).toEqual([
      [1234.5, 2500, 99.9],
      [1234, 2500, 987],
    ]);
  });

  it("counts an Indian grouping as American evidence, so a mixed block is left alone", () => {
    expect(looksEuropean([["12,34,567.89", "987,5"]])).toBe(false);
    expect(looksEuropean([["987,5", "1.234"]])).toBe(true);
  });
});

/* ───────────────────────────────── r3-62 ───────────────────────────────────── */

describe("the pane's number boxes read the chart's locale (r3-62)", () => {
  it("typedNumber: only a single mark and exactly three digits is left to the locale", () => {
    // Where the locale writes its decimal with a comma:
    expect(typedNumber("0,5", "da-DK")).toBe(0.5);
    expect(typedNumber("3.000", "da-DK")).toBe(3000);
    expect(typedNumber("1,500", "da-DK")).toBe(1.5);
    expect(typedNumber("1.234,5", "de-DE")).toBe(1234.5);
    expect(typedNumber("1.234.567", "da-DK")).toBe(1234567);
    expect(typedNumber("1 234,5", "fr-FR")).toBe(1234.5);
    // Where it writes it with a dot:
    expect(typedNumber("3.000", "en-US")).toBe(3);
    expect(typedNumber("1,500", "en-US")).toBe(1500);
    expect(typedNumber("1,234.5", "en-US")).toBe(1234.5);
    expect(typedNumber("1,234,567", "en-US")).toBe(1234567);
    // One reading only, so any locale:
    expect(typedNumber("0,5", "en-US")).toBe(0.5);
    expect(typedNumber("1.25", "da-DK")).toBe(1.25);
    expect(typedNumber("0.500", "da-DK")).toBe(0.5);
    expect(typedNumber("-0,5", "da-DK")).toBe(-0.5);
    expect(typedNumber("−0.5", "en-US")).toBe(-0.5);
    expect(typedNumber(".5", "en-US")).toBe(0.5);
    expect(typedNumber("1e3", "en-US")).toBe(1000);
    expect(typedNumber("1.1e2", "en-US")).toBe(110);
    // Not numbers:
    for (const t of ["", "abc", "1,2,3", "1.2.3", "1,5.2,3", ",", "1 23", "0x10", "Infinity", "5-"]) {
      expect(typedNumber(t, "en-US"), t).toBeNaN();
    }
  });

  it("splitValueList: a comma that is the locale's decimal does not split", () => {
    expect(splitValueList("50, 100", "en-US")).toEqual(["50", " 100"]);
    expect(splitValueList("1,5", "en-US")).toEqual(["1", "5"]);
    expect(splitValueList("1,5", "da-DK")).toEqual(["1,5"]);
    expect(splitValueList("50, 100", "da-DK")).toEqual(["50", "100"]);
    expect(splitValueList("1,5; 2,5", "da-DK")).toEqual(["1,5", " 2,5"]);
    expect(splitValueList("1.5; 2.5", "en-US")).toEqual(["1.5", " 2.5"]);
  });

  it("Axis scale takes a decimal comma and a Danish thousand in a da-DK chart", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: base, numberFormat: { locale: "da-DK" } });
    const [min, max] = boxes("Axis scale min");
    typeInto(min, "0,5");
    typeInto(max, "3.000");
    expect(exportConfig().scale).toEqual({ min: 0.5, max: 3000 });
    // The box keeps what was typed; it is not rewritten under the cursor.
    expect([min.value, max.value]).toEqual(["0,5", "3.000"]);
  });

  it("Axis scale still reads 3.000 as 3 in an en-US chart, and 0,5 as a half", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: base });
    const [min, max] = boxes("Axis scale min");
    typeInto(min, "0,5");
    typeInto(max, "3.000");
    expect(exportConfig().scale).toEqual({ min: 0.5, max: 3 });
    typeInto(max, "1,500");
    expect(exportConfig().scale).toEqual({ min: 0.5, max: 1500 });
  });

  it("Axis break takes decimal commas", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: base, numberFormat: { locale: "de-DE" } });
    const [from, to] = boxes("Axis break from");
    typeInto(from, "1,5");
    typeInto(to, "2,5");
    expect(exportConfig().axisBreak).toEqual({ from: 1.5, to: 2.5 });
  });

  it("changing the locale keeps the numbers and re-spells the boxes", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: base, scale: { min: 0.5, max: 1.234 } });
    const [min, max] = boxes("Axis scale min");
    expect([min.value, max.value]).toEqual(["0.5", "1.234"]);
    const locale = document.querySelector<HTMLSelectElement>('select[aria-label="Label locale"]')!;
    locale.value = "da-DK";
    locale.dispatchEvent(new Event("change", { bubbles: true }));
    // 1.234 must not become 1234 because the same characters now mean something else.
    expect([min.value, max.value]).toEqual(["0,5", "1,234"]);
    expect(exportConfig().scale).toEqual({ min: 0.5, max: 1.234 });
    locale.value = "en-US";
    locale.dispatchEvent(new Event("change", { bubbles: true }));
    expect([min.value, max.value]).toEqual(["0.5", "1.234"]);
    expect(exportConfig().scale).toEqual({ min: 0.5, max: 1.234 });
  });

  it("Value line: 1,5 is one line in da-DK and stays a comma list in en-US", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: base, numberFormat: { locale: "da-DK" } });
    const [values] = boxes("Value line");
    typeInto(values, "1,5");
    expect(exportConfig().decorations?.valueLines).toEqual([{ mode: "value", value: 1.5 }]);
    typeInto(values, "1,5; 2,5");
    expect(exportConfig().decorations?.valueLines).toEqual([
      { mode: "value", value: 1.5 },
      { mode: "value", value: 2.5 },
    ]);
    expect(values.placeholder).toContain(";");

    importConfig({ kind: "clustered", data: base });
    const [en] = boxes("Value line");
    typeInto(en, "1,5");
    expect(exportConfig().decorations?.valueLines).toEqual([
      { mode: "value", value: 1 },
      { mode: "value", value: 5 },
    ]);
    typeInto(en, "50, 100.5");
    expect(exportConfig().decorations?.valueLines).toEqual([
      { mode: "value", value: 50 },
      { mode: "value", value: 100.5 },
    ]);
  });
});

/* ───────────────────────────────── r3-11 ───────────────────────────────────── */

/** A one-series sheet whose row counts how many times a cell is read. */
function countingSheet(first: string, n: number, formula: (prev: string) => string) {
  let reads = 0;
  const letter = (c: number) => String.fromCharCode(65 + c);
  const row = ["Chain", first];
  for (let c = 2; c < n + 1; c++) row.push(`=${formula(`${letter(c - 1)}2`)}`);
  const counted = new Proxy(row, {
    get(target, prop, receiver) {
      if (typeof prop === "string" && /^\d+$/.test(prop)) reads++;
      return Reflect.get(target, prop, receiver);
    },
  });
  const cells = [["", ...row.slice(1).map((_, i) => `C${i}`)], counted as unknown as string[]];
  return { cells, reads: () => reads };
}

describe("a text or circular leaf does not bring back the 2^depth freeze (r3-11)", () => {
  const double = (prev: string) => `${prev}+${prev}*0.05`;

  it("reads each cell of a double-referencing chain a bounded number of times when the root is text", () => {
    for (const root of ["TBD", "n/a", "#DIV/0!"]) {
      const { cells, reads } = countingSheet(root, 16, double);
      const values = sheetToData({ cells }).series[0].values;
      expect(
        values.every((v) => v === null),
        root,
      ).toBe(true);
      // 16 cells, each evaluated afresh with its own memo: a few reads per cell
      // per step, hundreds in all. Unmemoised it is 2^16 per cell.
      expect(reads(), `${root}: ${reads()} reads`).toBeLessThan(3000);
    }
  });

  it("and when the root is a cycle", () => {
    // B2 and C2 name each other; D2 onward double up on C2.
    const { cells, reads } = countingSheet("=C2", 16, double);
    cells[1][2] = "=B2";
    const values = sheetToData({ cells }).series[0].values;
    expect(values.every((v) => v === null)).toBe(true);
    expect(reads(), `${reads()} reads`).toBeLessThan(3000);
  });

  it("is unchanged for a numeric root, and agrees with the arithmetic", () => {
    const { cells, reads } = countingSheet("100", 16, double);
    const values = sheetToData({ cells }).series[0].values;
    expect(values[0]).toBe(100);
    expect(values[15]).toBeCloseTo(100 * 1.05 ** 15, 6);
    expect(reads()).toBeLessThan(3000);
  });

  it("still tells a cycle from a cell that merely reads one", () => {
    const cells = [
      ["", "A", "B", "C", "D"],
      ["Row", "=C2+1", "=B2+1", "=B2", "7"],
      ["Two", "=E2", "=B3", "5", "=C3*2"],
    ];
    const data = sheetToData({ cells });
    // B2 and C2 name each other and D2 reads the cycle; E2 is a plain 7. Row 3
    // reaches E2 and nothing circular: its answers are the same on every path.
    expect(data.series[0].values).toEqual([null, null, null, 7]);
    expect(data.series[1].values).toEqual([7, 7, 5, 14]);
  });
});

/* ───────────────────────────────── r3-51 ───────────────────────────────────── */

describe("a deeply nested formula is refused, never thrown (r3-51)", () => {
  const nested = (n: number) => "(".repeat(n) + "1" + ")".repeat(n);

  it("returns null for parentheses, unary minus and nested functions far past the stack", () => {
    for (const expr of [nested(6000), "-".repeat(100_000) + "1", "SUM(".repeat(4000) + "1" + ")".repeat(4000)]) {
      expect(() => evaluateFormula([["", "A"]], expr)).not.toThrow();
      expect(evaluateFormula([["", "A"]], expr)).toBeNull();
    }
  });

  it("keeps sheetToData and the cells beside it working", () => {
    const cells = [
      ["", "A", "B", "C"],
      ["Row", "=" + nested(6000), "=((1+2))*3", "5"],
    ];
    let data!: ReturnType<typeof sheetToData>;
    expect(() => (data = sheetToData({ cells }))).not.toThrow();
    expect(data.series[0].values).toEqual([null, 9, 5]);
  });

  /** A row whose cell at column `i` adds one to the cell before it: B2 = 1, C2 = B2+1, … */
  const chain = (length: number) => {
    const col = (c: number) => {
      let s = "";
      for (let i = c; i >= 0; i = Math.floor(i / 26) - 1) s = String.fromCharCode(65 + (i % 26)) + s;
      return s;
    };
    const row = ["Chain", "1"];
    for (let c = 2; c <= length; c++) row.push(`=${col(c - 1)}2+1`);
    return { cells: [["", ...row.slice(1).map((_, i) => `C${i}`)], row], last: `${col(length)}2` };
  };

  it("allows what Excel allows, and a long chain of cells", () => {
    expect(evaluateFormula([["", "A"]], nested(64))).toBe(1);
    expect(evaluateFormula([["", "A"]], "-".repeat(60) + "1")).toBe(1);
    const { cells, last } = chain(300);
    expect(evaluateFormula(cells, last)).toBe(300);
  });

  it("refuses a chain of cells deeper than the evaluator will follow, without throwing", () => {
    const { cells, last } = chain(900);
    expect(() => evaluateFormula(cells, last)).not.toThrow();
    expect(evaluateFormula(cells, last)).toBeNull();
    // A cell the chain reaches sooner is still answered: the refusal is per formula.
    expect(evaluateFormula(cells, "B2+C2")).toBe(3);
  });

  it("answers a cell that names itself five hundred times without throwing", () => {
    const cells = [
      ["", "A"],
      ["Row", "=" + Array.from({ length: 500 }, () => "B2").join("+")],
    ];
    expect(() => sheetToData({ cells })).not.toThrow();
    expect(sheetToData({ cells }).series[0].values).toEqual([null]);
  });
});

/* ──────────────────────────── r3-37 and r3-52 ────────────────────────────── */

describe("a date format over a number that is no date does not throw (r3-37, r3-52)", () => {
  const ISO = /^\d{4}-\d{2}-\d{2}$/;

  it("excelSerialToISO answers a four-digit-year date or nothing, for any number", () => {
    const numbers = [
      NaN,
      Infinity,
      -Infinity,
      1e100,
      -1e100,
      1767571200000,
      1e9,
      100_000_001,
      -100_000_001,
      1e8,
      -1e8,
      20260105,
      3_500_000,
      2_958_466,
      -693_960,
      -1e6,
      -700_000,
      0,
      1,
      60,
      46027,
      2_958_465,
      0.5,
      Number.MAX_VALUE,
      Number.MIN_VALUE,
    ];
    for (const n of numbers) {
      for (const flag of [false, true]) {
        let out = "";
        expect(() => (out = excelSerialToISO(n, flag)), `${n} threw`).not.toThrow();
        expect(out === "" || ISO.test(out), `${n} gave ${out}`).toBe(true);
      }
    }
    expect(excelSerialToISO(1e9)).toBe("");
    expect(excelSerialToISO(1767571200000)).toBe("");
    expect(excelSerialToISO(20260105)).toBe("");
    expect(excelSerialToISO(2_958_466)).toBe("");
    expect(excelSerialToISO(NaN)).toBe("");
    expect(excelSerialToISO(2_958_465)).toBe("9999-12-31");
    expect(excelSerialToISO(46027)).toBe("2026-01-05");
    expect(excelSerialToISO(44565, true)).toBe("2026-01-05");
  });

  it("a body cell keeps the number it holds instead of failing the whole selection", () => {
    const how = { label: false, date1904: false };
    expect(() => cellText(1767571200000, "yyyy-mm-dd", "########", how)).not.toThrow();
    expect(cellText(1767571200000, "yyyy-mm-dd", "########", how)).toBe("1767571200000");
    // Not "+057370-04", which sheetToData then read as the number 20234525.
    expect(cellText(20260105, "yyyy-mm-dd", "########", how)).toBe("20260105");
    expect(cellText(46027, "yyyy-mm-dd", "2026-01-05", how)).toBe("2026-01-05");
    const sheet = sheetToData({
      cells: [
        ["", "X"],
        ["Row", cellText(20260105, "yyyy-mm-dd", "########", how)],
      ],
    });
    expect(sheet.series[0].values).toEqual([20260105]);
    expect(sheet.dates).toBeUndefined();
  });

  it("a label whose displayed text did not come back falls through the same guard", () => {
    expect(() => cellText(1767571200000, "yyyy-mm-dd", undefined, { label: true, date1904: false })).not.toThrow();
    expect(cellText(1767571200000, "yyyy-mm-dd", undefined, { label: true, date1904: false })).toBe("1767571200000");
    // And a label that is out of range but DID come back keeps what Excel showed.
    expect(cellText(1767571200000, "yyyy-mm-dd", "########", { label: true, date1904: false })).toBe("########");
  });

  it("the datasheet's own twin (isoDay, through a Gantt row) tolerates the same numbers", () => {
    // Already total at HEAD — this pins it, so the two cannot drift apart again.
    const values = [1e308, -1e308, 1e9, 1e8, 20260105, epochDay(2026, 1, 5), -1e8 - 1, 1e8 + 1, 0];
    const data = {
      categories: values.map((_, i) => `C${i}`),
      series: [{ name: "Start", values }],
      dates: true,
    };
    let sheet!: ReturnType<typeof dataToSheet>;
    expect(() => (sheet = dataToSheet(data))).not.toThrow();
    const row = sheet.cells[1].slice(1);
    expect(row[5]).toBe("2026-01-05");
    expect(row[8]).toBe("1970-01-01");
    // Everything that is no day stays the number it was, which round-trips as one.
    for (const i of [0, 1, 2, 3, 4, 6, 7]) expect(row[i], `value ${values[i]}`).toBe(String(values[i]));
  });
});

/* ───────────────────────────────── r3-49 ───────────────────────────────────── */

describe("a style file with a null palette is refused, not persisted (r3-49)", () => {
  const importStyle = (text: string) => {
    ($("json-io") as HTMLTextAreaElement).value = text;
    $("style-import").click();
  };
  const preview = () => $("preview").textContent ?? "";

  it("says why and keeps the pane drawing", async () => {
    await bootPane();
    importStyle('{"palette": null}');
    expect(note()).toMatch(/Style import failed/);
    expect(note()).toMatch(/palette/);
    expect($("host-note").className).toMatch(/err/);
    expect(window.localStorage.getItem("ssf-charts-style")).toBeNull();
    // The drawing is untouched, and not "Could not render: Cannot read properties of null".
    importConfig({ kind: "clustered", data: base });
    await vi.waitFor(() => expect($("preview").querySelector("svg")).not.toBeNull(), { timeout: 5000 });
    expect(preview()).not.toMatch(/Could not render/);
  });

  it("does not replace a good style that was already imported", async () => {
    await bootPane();
    importStyle('{"palette": ["#112233", "#445566"]}');
    expect(window.localStorage.getItem("ssf-charts-style")).toBe('{"palette":["#112233","#445566"]}');
    importStyle('{"palette": null, "fontSize": 14}');
    expect(note()).toMatch(/Style import failed/);
    expect(window.localStorage.getItem("ssf-charts-style")).toBe('{"palette":["#112233","#445566"]}');
  });

  it("still accepts a style without a palette and one with a list", async () => {
    await bootPane();
    importStyle("{}");
    expect(note()).toMatch(/Style imported/);
    importStyle('{"palette": ["#112233"], "fontSize": 12}');
    expect(note()).toMatch(/Style imported/);
    expect(window.localStorage.getItem("ssf-charts-style")).toContain("#112233");
  });
});

/* ───────────────────────────────── r3-13 ───────────────────────────────────── */

describe("a config whose valueLines the engine repairs opens in the pane too (r3-13)", () => {
  const withLines = (valueLines: unknown, title = "T") =>
    ({ kind: "clustered", title, data: base, decorations: { valueLines } }) as unknown as Partial<ChartConfig>;
  const groups = () => document.querySelectorAll("#options details").length;
  const controls = () => document.querySelectorAll("#options input, #options select").length;

  it("opens every shape the engine draws, with the options panel intact", async () => {
    await bootPane();
    importConfig({ kind: "clustered", data: base });
    const healthy = { groups: groups(), controls: controls() };
    expect(healthy.groups).toBe(5);
    for (const lines of ["mean", 5, true, false, 0, "", {}, [null], [null, { mode: "value", value: 3 }], [[1]]]) {
      importConfig(withLines(lines, "NEW TITLE"));
      expect(note(), JSON.stringify(lines)).toBe("Chart config loaded.");
      expect(groups(), `${JSON.stringify(lines)} left the options panel empty`).toBe(healthy.groups);
      expect(controls()).toBe(healthy.controls);
      expect(($("chart-title") as HTMLInputElement).value).toBe("NEW TITLE");
    }
  });

  it("keeps the lines the engine keeps, and repairs only what it repairs", async () => {
    await bootPane();
    importConfig(withLines([null, { mode: "mean" }, { mode: "value", value: 3 }]));
    const [values] = boxes("Value line");
    expect(values.value).toBe("3");
    expect(exportConfig().decorations?.valueLines).toEqual([{ mode: "mean" }, { mode: "value", value: 3 }]);
    // A valid list, and an empty one, are exactly as written.
    importConfig(withLines([{ mode: "value", value: 50 }]));
    expect(exportConfig().decorations?.valueLines).toEqual([{ mode: "value", value: 50 }]);
    importConfig(withLines([]));
    expect(exportConfig().decorations?.valueLines).toEqual([]);
    // Nothing readable left: dropped, as the engine drops it.
    importConfig(withLines("mean"));
    expect(exportConfig().decorations?.valueLines).toBeUndefined();
  });

  it("falls back to the legacy single valueLine only when valueLines says nothing", async () => {
    await bootPane();
    importConfig({
      kind: "clustered",
      data: base,
      decorations: { valueLines: "mean", valueLine: { mode: "value", value: 7 } },
    } as unknown as Partial<ChartConfig>);
    expect(note()).toBe("Chart config loaded.");
    expect(boxes("Value line")[0].value).toBe("7");
  });

  it("the engine draws the same config the pane now opens", () => {
    for (const lines of ["mean", [null]]) {
      expect(() => buildChart(normalizeConfig(withLines(lines) as ChartConfig))).not.toThrow();
    }
  });
});

/* ───────────────────────────────── r3-81 ───────────────────────────────────── */

describe("the remembered 'Insert as picture' default seeds every NEW chart (r3-81)", () => {
  const PREF = { "ssf-charts-render-image": "1" };
  const box = () => $("render-image") as HTMLInputElement;
  const pickKind = (kind: string) => document.querySelector<HTMLButtonElement>(`.thumb[data-kind="${kind}"]`)!.click();

  it("survives picking a chart type", async () => {
    await bootPane(PREF);
    expect(box().checked).toBe(true);
    pickKind("line");
    expect(box().checked, "the tick was dropped by picking a chart type").toBe(true);
    expect(exportConfig().render).toBe("image");
    expect(window.localStorage.getItem("ssf-charts-render-image")).toBe("1");
  });

  it("survives picking a starter template", async () => {
    await bootPane(PREF);
    const list = $("template-list") as HTMLSelectElement;
    const starter = [...list.options].find((o) => o.value.startsWith("builtin:"))!;
    list.value = starter.value;
    list.dispatchEvent(new Event("change", { bubbles: true }));
    expect(box().checked).toBe(true);
    expect(exportConfig().render).toBe("image");
  });

  it("survives a tick made in this session, even when storage is blocked", async () => {
    await bootPane();
    expect(box().checked).toBe(false);
    const setItem = vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => {
      throw new Error("blocked");
    });
    try {
      box().checked = true;
      box().dispatchEvent(new Event("change", { bubbles: true }));
    } finally {
      setItem.mockRestore();
    }
    pickKind("pie");
    expect(box().checked).toBe(true);
  });

  it("does not turn the default on for a pane that never had it, and honours an unticked box", async () => {
    await bootPane();
    pickKind("line");
    expect(box().checked).toBe(false);
    box().checked = true;
    box().dispatchEvent(new Event("change", { bubbles: true }));
    box().checked = false;
    box().dispatchEvent(new Event("change", { bubbles: true }));
    pickKind("pie");
    expect(box().checked).toBe(false);
  });

  it("is still a default and not an override: a loaded chart keeps its own render", async () => {
    await bootPane(PREF);
    importConfig({ kind: "clustered", data: base });
    expect(box().checked, "a JSON import is an existing chart").toBe(false);
    importConfig({ kind: "clustered", data: base, render: "image" });
    expect(box().checked).toBe(true);
    // A new chart after a loaded one starts from the PREFERENCE, not from the loaded chart.
    importConfig({ kind: "clustered", data: base });
    expect(box().checked).toBe(false);
    pickKind("line");
    expect(box().checked).toBe(true);
  });
});

/* ───────────────────────────────── r3-82 ───────────────────────────────────── */

describe("a series colour pick is an undo step (r3-82)", () => {
  const swatches = () => [...document.querySelectorAll<HTMLInputElement>(".series-colors input[type=color]")];
  const pick = (i: number, colour: string) => {
    const el = swatches()[i];
    el.value = colour;
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const seriesColor = (i: number) => exportConfig().data.series[i].color;
  const cellValue = (r: number, c: number) => cell(r, c).value;

  async function twoSeries() {
    await bootPane();
    importConfig({
      kind: "clustered",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [1, 2] },
          { name: "S2", values: [3, 4] },
        ],
      },
    });
  }

  it("Ctrl+Z takes the colour back and leaves the cell edit, Ctrl+Y puts the colour back", async () => {
    await twoSeries();
    typeInto(cell(1, 1), "10");
    pick(0, "#ff0000");
    expect(seriesColor(0)).toBe("#ff0000");

    ctrl("z");
    expect(seriesColor(0), "the pick was not an undo step").toBeUndefined();
    expect(cellValue(1, 1), "undoing the pick took the cell edit with it").toBe("10");
    ctrl("z");
    expect(cellValue(1, 1)).toBe("1");

    ctrl("y");
    expect(cellValue(1, 1)).toBe("10");
    ctrl("y");
    expect(seriesColor(0), "redo could not bring the colour back").toBe("#ff0000");
  });

  it("the swatch follows the undo, not just the chart", async () => {
    await twoSeries();
    const before = swatches()[0].value;
    typeInto(cell(1, 1), "10");
    pick(0, "#ff0000");
    expect(swatches()[0].value).toBe("#ff0000");
    ctrl("z");
    expect(swatches()[0].value).toBe(before);
    ctrl("y");
    expect(swatches()[0].value).toBe("#ff0000");
  });

  it("a pick between two edits no longer rides on the next edit's snapshot", async () => {
    await twoSeries();
    typeInto(cell(1, 1), "10");
    pick(0, "#ff0000");
    typeInto(cell(1, 2), "20");
    ctrl("z"); // the second edit
    expect(cellValue(1, 2)).toBe("2");
    expect(seriesColor(0), "undoing edit 2 dropped the colour").toBe("#ff0000");
    ctrl("z"); // the pick
    expect(seriesColor(0)).toBeUndefined();
    expect(cellValue(1, 1)).toBe("10");
  });

  it("one drag of the picker is one step, and does not push the datasheet's history out", async () => {
    await twoSeries();
    typeInto(cell(1, 1), "10");
    for (let i = 0; i < 150; i++) pick(0, `#${(0x100000 + i * 997).toString(16).padStart(6, "0")}`);
    swatches()[0].dispatchEvent(new Event("change", { bubbles: true }));
    expect(seriesColor(0)).toBeDefined();
    ctrl("z"); // the whole drag
    expect(seriesColor(0)).toBeUndefined();
    expect(cellValue(1, 1)).toBe("10");
    ctrl("z"); // the edit that came before it is still there
    expect(cellValue(1, 1)).toBe("1");
  });

  it("a second pick after the picker was let go is a step of its own", async () => {
    await twoSeries();
    pick(0, "#ff0000");
    swatches()[0].dispatchEvent(new Event("change", { bubbles: true }));
    pick(0, "#00ff00");
    swatches()[0].dispatchEvent(new Event("change", { bubbles: true }));
    expect(seriesColor(0)).toBe("#00ff00");
    ctrl("z");
    expect(seriesColor(0)).toBe("#ff0000");
    ctrl("z");
    expect(seriesColor(0)).toBeUndefined();
  });
});
