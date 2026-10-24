// @vitest-environment jsdom
import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { buildChart, normalizeConfig } from "../src/core/chart";
import type { SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { cellText, isPercentFormat, keepVisible, trimBlankEdges } from "../src/excel/cell-text";
import { excelSerialToISO, isDate1904, isDateFormat } from "../src/excel/serial-dates";
import { parseClipboardGrid, sheetToData } from "../src/taskpane/datasheet";

/**
 * THE EXCEL BRIDGE READ A RANGE AS IF EVERY CELL WERE ITS STORED NUMBER.
 *
 * Six findings from the 2026-10 sweep, one root: `values` is what a cell
 * STORES, and the chart wants what the sheet SHOWS. A 3.5% cell is stored as
 * 0.035, a `[h]:mm` duration as a fraction of a day, a date as a serial counted
 * from one of two epochs, a hidden row as an ordinary row, and a whole column
 * as nothing at all.
 *
 * THE FAKE BELOW IS THE POINT OF THE FILE. `test/excel-bridge.test.ts` stubs a
 * range with a `values` getter and a `load` that does nothing, which cannot
 * fail the way Excel fails: it hands back data that was never loaded, accepts a
 * property the host does not have, and has no notion of an unbounded range.
 * This one is queued like the real object model — `load` asks, `sync` answers,
 * and reading a property that was not asked for throws — it refuses the ExcelApi
 * 1.2 members on a 1.1 host, returns null cell data for a whole-column selection
 * as Microsoft documents, and reports `rowHidden` the documented three ways
 * (true for all, false for none, null for some).
 */

interface Book {
  /** The cells as Excel STORES them: a date is its serial, a percent a fraction. */
  values: unknown[][];
  numberFormat?: string[][];
  /** What the cells SHOW. Defaults to the value as text. */
  text?: string[][];
  address?: string;
  /** Zero-based rows and columns of `values` that are hidden or filtered out. */
  hiddenRows?: number[];
  hiddenCols?: number[];
  date1904?: boolean;
  /** Minor version of ExcelApi 1.x the host reports. Default 2. */
  minor?: 1 | 2;
  /** The selection is whole columns: cell-level data comes back null. */
  unbounded?: boolean;
  /** Something other than "nothing there" goes wrong reading the used range. */
  usedRangeFault?: Error;
}

interface Queued {
  flush(): void;
  missing?: boolean;
  fault?: Error;
  pending(): boolean;
}

const a1 = (r0: number, c0: number, nr: number, nc: number): string =>
  `Sheet1!${String.fromCharCode(65 + c0)}${r0 + 1}:${String.fromCharCode(65 + c0 + nc - 1)}${r0 + nr}`;

class Host {
  readonly counts = { getRow: 0, getColumn: 0, usedRange: 0, dateCalls: 0 };
  readonly loaded: string[] = [];
  private readonly objects: Queued[] = [];
  constructor(readonly book: Book) {}

  get minor(): number {
    return this.book.minor ?? 2;
  }
  supports(name: string, version: string): boolean {
    return name === "ExcelApi" && Number(version.split(".")[1]) <= this.minor;
  }
  register(o: Queued): void {
    this.objects.push(o);
  }
  async sync(): Promise<void> {
    for (const o of this.objects) {
      if (o.missing && o.pending()) {
        throw Object.assign(new Error("The requested resource doesn't exist."), { code: "ItemNotFound" });
      }
      if (o.fault && o.pending()) throw o.fault;
    }
    for (const o of this.objects) o.flush();
  }
}

/** A queued client object: `load` asks, `sync` answers, an unasked property throws. */
abstract class ClientObject implements Queued {
  missing = false;
  fault?: Error;
  private readonly ready = new Set<string>();
  private queued: string[] = [];
  constructor(protected readonly host: Host) {
    host.register(this);
  }
  load(props: string): void {
    for (const p of props.split(",")) {
      if (/Hidden$/.test(p) && this.host.minor < 2) throw new Error(`ApiNotFound: ${p} needs ExcelApi 1.2`);
      this.queued.push(p);
      this.host.loaded.push(p);
    }
  }
  pending(): boolean {
    return this.queued.length > 0;
  }
  flush(): void {
    for (const p of this.queued) this.ready.add(p);
    this.queued = [];
  }
  protected need(p: string): void {
    if (!this.ready.has(p)) throw new Error(`PropertyNotLoaded: ${p} was read before it was loaded and synced`);
  }
}

class FunctionResult extends ClientObject {
  constructor(
    host: Host,
    private readonly serial: number,
  ) {
    super(host);
  }
  get value(): number {
    this.need("value");
    return this.serial;
  }
}

class FakeRange extends ClientObject {
  constructor(
    host: Host,
    private readonly r0: number,
    private readonly nr: number,
    private readonly c0: number,
    private readonly nc: number,
    private readonly bounded: boolean,
    private readonly addr: string,
  ) {
    super(host);
  }
  private grid<T>(all: T[][]): T[][] | null {
    return this.bounded
      ? all.slice(this.r0, this.r0 + this.nr).map((row) => row.slice(this.c0, this.c0 + this.nc))
      : null;
  }
  get values(): unknown[][] | null {
    this.need("values");
    return this.grid(this.host.book.values);
  }
  get numberFormat(): string[][] | null {
    this.need("numberFormat");
    return this.grid(this.host.book.numberFormat ?? this.host.book.values.map((r) => r.map(() => "General")));
  }
  get text(): string[][] | null {
    this.need("text");
    return this.grid(
      this.host.book.text ?? this.host.book.values.map((r) => r.map((v) => (v == null ? "" : String(v)))),
    );
  }
  get address(): string {
    this.need("address");
    return this.addr;
  }
  private tri(hidden: number[], from: number, n: number): boolean | null {
    const k = hidden.filter((i) => i >= from && i < from + n).length;
    return k === 0 ? false : k === n ? true : null;
  }
  get rowHidden(): boolean | null {
    this.need("rowHidden");
    return this.tri(this.host.book.hiddenRows ?? [], this.r0, this.nr);
  }
  get columnHidden(): boolean | null {
    this.need("columnHidden");
    return this.tri(this.host.book.hiddenCols ?? [], this.c0, this.nc);
  }
  getRow(i: number): FakeRange {
    this.host.counts.getRow++;
    return new FakeRange(this.host, this.r0 + i, 1, this.c0, this.nc, true, this.addr);
  }
  getColumn(j: number): FakeRange {
    this.host.counts.getColumn++;
    return new FakeRange(this.host, this.r0, this.nr, this.c0 + j, 1, true, this.addr);
  }
  getUsedRange(valuesOnly?: boolean): FakeRange {
    // `getUsedRange` is 1.1 but its `valuesOnly` argument is 1.2 (@types/office-js).
    if (valuesOnly !== undefined && this.host.minor < 2) throw new Error("ApiNotFound: valuesOnly needs ExcelApi 1.2");
    this.host.counts.usedRange++;
    const cells = this.host.book.values;
    let r1 = -1;
    let c1 = -1;
    let r0 = Infinity;
    let c0 = Infinity;
    cells.forEach((row, r) =>
      row.forEach((v, c) => {
        if (v == null || v === "") return;
        r0 = Math.min(r0, r);
        c0 = Math.min(c0, c);
        r1 = Math.max(r1, r);
        c1 = Math.max(c1, c);
      }),
    );
    const none = r1 < 0;
    const used = new FakeRange(
      this.host,
      none ? 0 : r0,
      none ? 1 : r1 - r0 + 1,
      none ? 0 : c0,
      none ? 1 : c1 - c0 + 1,
      true,
      none ? this.addr : a1(r0, c0, r1 - r0 + 1, c1 - c0 + 1),
    );
    used.missing = none;
    used.fault = this.host.book.usedRangeFault;
    return used;
  }
}

/** Put a fake workbook behind the globals the pane reads at call time. */
function install(book: Book): Host {
  const host = new Host(book);
  const rows = book.values.length;
  const cols = book.values[0]?.length ?? 0;
  const selected = new FakeRange(host, 0, rows, 0, cols, !book.unbounded, book.address ?? a1(0, 0, rows, cols));
  // Serial Excel gives DATE(y, m, d): counted from 1899-12-30 on the 1900 system
  // and from 1904-01-01 on the 1904 one.
  const dateFn = (y: number, m: number, d: number): FunctionResult => {
    host.counts.dateCalls++;
    const epoch = book.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    return new FunctionResult(host, Math.round((Date.UTC(y, m - 1, d) - epoch) / 86400000));
  };
  const workbook: Record<string, unknown> = { getSelectedRange: () => selected };
  if (host.minor >= 2) workbook.functions = { date: dateFn };
  vi.stubGlobal("Office", {
    onReady: (cb: () => void) => cb(),
    context: { host: "Excel", requirements: { isSetSupported: (n: string, v: string) => host.supports(n, v) } },
  });
  vi.stubGlobal("Excel", {
    run: async (cb: (ctx: unknown) => Promise<void>) => cb({ workbook, sync: () => host.sync() }),
  });
  return host;
}

const el = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const note = (): string => el("note").textContent ?? "";
const output = (): ChartConfig => JSON.parse(el<HTMLTextAreaElement>("output").value);

async function generate(kind = "stacked", transpose = false): Promise<void> {
  el<HTMLSelectElement>("kind").value = kind;
  el<HTMLInputElement>("transpose").checked = transpose;
  el("generate").click();
  // The click handler is async and every sync is a microtask: one turn flushes it.
  await new Promise((r) => setTimeout(r, 0));
}

beforeAll(async () => {
  document.body.innerHTML = `
    <select id="kind">${["stacked", "clustered", "line", "gantt", "waterfall"].map((k) => `<option value="${k}">${k}</option>`).join("")}</select>
    <input id="title" value="">
    <input type="checkbox" id="transpose">
    <button id="generate"></button>
    <button id="copy"></button>
    <textarea id="output"></textarea>
    <p id="note"></p>`;
  install({ values: [[""]] });
  await import("../src/excel/excel");
});

beforeEach(() => {
  el<HTMLTextAreaElement>("output").value = "";
  el("note").textContent = "";
});

/** Text drawn on a chart, in scene order. */
const chartTexts = (cfg: ChartConfig): string[] =>
  buildChart(normalizeConfig(cfg))
    .nodes.filter((n: SceneNode): n is TextNode => n.kind === "text")
    .map((n) => n.text);

const epochDay = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d) / 86400000;

/* ───────────────────────────── r2-27: percents ───────────────────────────── */

describe("a percent-formatted cell is the number the paste route reads (r2-27)", () => {
  const growth = (fmt: string, v: number): Book => ({
    values: [
      ["", "2024"],
      ["Growth", v],
    ],
    numberFormat: [
      ["General", "General"],
      ["General", fmt],
    ],
    text: [
      ["", "2024"],
      ["Growth", "shown"],
    ],
  });
  const firstValue = () => output().data.series[0].values[0];

  it("brings 3.5% in as 3.5, exactly what pasting the displayed text gives", async () => {
    install({
      values: [
        ["", "2024", "2025"],
        ["Growth", 0.035, 0.042],
      ],
      numberFormat: [
        ["General", "General", "General"],
        ["General", "0.0%", "0.0%"],
      ],
      text: [
        ["", "2024", "2025"],
        ["Growth", "3.5%", "4.2%"],
      ],
    });
    await generate();
    const pasted = sheetToData({ cells: parseClipboardGrid("\t2024\t2025\r\nGrowth\t3.5%\t4.2%\r\n") });
    expect(output().data.series).toEqual(pasted.series);
    expect(output().data.series[0].values).toEqual([3.5, 4.2]);
  });

  it("so a '%' suffix labels the bars 3.5% and not 0.035%", async () => {
    install({
      values: [
        ["", "2024", "2025"],
        ["Growth", 0.035, 0.042],
      ],
      numberFormat: [
        ["General", "General", "General"],
        ["General", "0.0%", "0.0%"],
      ],
    });
    await generate("clustered");
    const texts = chartTexts({ ...output(), numberFormat: { suffix: "%" } });
    expect(texts).toContain("3.5%");
    expect(texts).toContain("4.2%");
    expect(texts, "a hundredth of the sheet's figure reached the label").not.toContain("0.035%");
  });

  it.each([
    ["0%", 0.35, 35],
    ["0.00%", -0.125, -12.5],
    ["#,##0.0%", 12.5, 1250],
    ["0.0%;[Red]-0.0%", -0.02, -2],
    ["[$-409]0.0%", 0.5, 50],
    ["0.0%_);(0.0%)", 0.25, 25],
    // 0.07 * 100 is 7.000000000000001; it must not reach the JSON box as that.
    ["0%", 0.07, 7],
    ["0.0%", 0.035, 3.5],
  ])("reads %s on %d as %d", async (fmt, v, expected) => {
    install(growth(fmt, v));
    await generate();
    expect(firstValue()).toBe(expected);
  });

  it.each([
    ["a quoted percent sign", '0.0"%"'],
    ["an escaped percent sign", "0.0\\%"],
    ["a skipped-width percent sign", "0.0_%"],
    ["a plain number", "0.0"],
    ["General", "General"],
  ])("leaves a cell alone when its format only PRINTS a percent sign: %s", async (_label, fmt) => {
    // 3.5 under 0.0"%" shows as "3.5%": already percentage points. Multiplying
    // it would be the mistake in the other direction.
    install(growth(fmt, 3.5));
    await generate();
    expect(firstValue()).toBe(3.5);
  });

  it("answers per code, through the same reader as the date check", () => {
    expect(isPercentFormat("0.0%")).toBe(true);
    expect(isPercentFormat("[Red]0%")).toBe(true);
    expect(isPercentFormat('0.0" %"')).toBe(false);
    expect(isPercentFormat(undefined)).toBe(false);
    expect(isPercentFormat(0.5)).toBe(false);
  });
});

/* ──────────────────────── r1-30: elapsed-time formats ─────────────────────── */

describe("an elapsed-time format is not a date (r1-30)", () => {
  it.each(["[h]:mm", "[hh]:mm", "[H]:MM", '[h]"h" mm"m"', "[h]:mm;@", "[m]", "[mm]:ss", "[s]", "[ss].00", "[h]:mm:ss"])(
    "does not read %s as a date",
    (code) => {
      expect(isDateFormat(code), `"${code}" was read as a date format`).toBe(false);
    },
  );

  it.each(["[$-409]d mmmm yyyy", "[Red]mm/dd/yyyy", "yyyy-mm-dd", "[$-F800]dddd, mmmm dd, yyyy", "mmm-yy", "mm"])(
    "still reads %s as a date",
    (code) => {
      expect(isDateFormat(code)).toBe(true);
    },
  );

  it("charts 36 hours as the number it is, not as 1900-01-01 minus 25,568 days", async () => {
    install({
      values: [
        ["", "Proj A", "Proj B", "Proj C"],
        ["Hours", 1.5, 0.5208333333, 2.0833333333],
      ],
      numberFormat: [
        ["General", "General", "General", "General"],
        ["General", "[h]:mm", "[h]:mm", "[h]:mm"],
      ],
    });
    await generate();
    const data = output().data;
    expect(data.dates, "the durations were read as calendar days").toBeUndefined();
    expect(data.series[0].values).toEqual([1.5, 0.5208333333, 2.0833333333]);
  });
});

/* ──────────────────────── r2-45: the 1904 date system ─────────────────────── */

describe("the workbook's date system decides which day a serial is (r2-45)", () => {
  it("counts a 1904 serial from 1904, a day-for-day match for the 1900 one", () => {
    // The same two days, as the two systems store them (they differ by 1,462).
    expect(excelSerialToISO(46027)).toBe("2026-01-05");
    expect(excelSerialToISO(44565, true)).toBe("2026-01-05");
    expect(excelSerialToISO(44575, true)).toBe("2026-01-15");
    expect(excelSerialToISO(0, true)).toBe("1904-01-01");
  });

  it("tells the systems apart by what DATE(2000,1,1) answers", () => {
    expect(isDate1904(36526)).toBe(false);
    expect(isDate1904(35064)).toBe(true);
    // An answer that fits neither is the 1900 system, as before the bridge could ask.
    expect(isDate1904(undefined)).toBe(false);
    expect(isDate1904(Number.NaN)).toBe(false);
    expect(isDate1904("#NAME?")).toBe(false);
  });

  const schedule = (start: number, end: number): Book => ({
    values: [
      ["", "Start", "End"],
      ["Task", start, end],
    ],
    numberFormat: [
      ["General", "General", "General"],
      ["General", "yyyy-mm-dd", "yyyy-mm-dd"],
    ],
    text: [
      ["", "Start", "End"],
      ["Task", "2026-01-05", "2026-01-15"],
    ],
  });
  const days = [epochDay(2026, 1, 5), epochDay(2026, 1, 15)];

  it("reads a 1904 workbook's Gantt dates as the days they are, not four years early", async () => {
    install({ ...schedule(44565, 44575), date1904: true });
    await generate("gantt");
    expect(output().data.series[0].values).toEqual(days);
  });

  it("is unchanged for the default 1900 workbook", async () => {
    install(schedule(46027, 46037));
    await generate("gantt");
    expect(output().data.series[0].values).toEqual(days);
  });

  it("asks the workbook only where the host can answer, and keeps reading 1900 where it cannot", async () => {
    // ExcelApi 1.1 has no `workbook.functions`. Asking anyway would end the run
    // in a TypeError, so the bridge must not.
    const host = install({ ...schedule(46027, 46037), date1904: true, minor: 1 });
    await generate("gantt");
    expect(host.counts.dateCalls).toBe(0);
    expect(note()).not.toMatch(/^Failed/);
    expect(output().data.series[0].values).toEqual(days);
  });
});

/* ──────────────────────── r2-47: header and label cells ───────────────────── */

describe("the row and column that label the chart read as Excel shows them (r2-47)", () => {
  const months = {
    values: [
      ["", 46023, 46054, 46082],
      ["North", 10, 20, 30],
    ],
    numberFormat: [
      ["General", "mmm-yy", "mmm-yy", "mmm-yy"],
      ["General", "General", "General", "General"],
    ],
    text: [
      ["", "Jan-26", "Feb-26", "Mar-26"],
      ["North", "10", "20", "30"],
    ],
  };

  it("keeps a month header as Jan-26, which is what the paste route gives", async () => {
    install(months);
    await generate("line");
    const pasted = sheetToData({ cells: parseClipboardGrid("\tJan-26\tFeb-26\tMar-26\r\nNorth\t10\t20\t30\r\n") });
    expect(output().data.categories).toEqual(["Jan-26", "Feb-26", "Mar-26"]);
    expect(output().data.categories).toEqual(pasted.categories);
    expect(chartTexts(output())).not.toContain("2026-01-01");
  });

  it("keeps a percent header as 10%, and a numeric header as the sheet shows it", async () => {
    install({
      values: [
        ["", 0.1, 2024],
        ["North", 5, 6],
      ],
      numberFormat: [
        ["General", "0%", "General"],
        ["General", "General", "General"],
      ],
      text: [
        ["", "10%", "2024"],
        ["North", "5", "6"],
      ],
    });
    await generate();
    expect(output().data.categories).toEqual(["10%", "2024"]);
  });

  it("does the same to the dates down column A when the series run across the top", async () => {
    install({
      values: [
        ["", "North"],
        [46023, 10],
        [46054, 20],
      ],
      numberFormat: [
        ["General", "General"],
        ["mmm-yy", "General"],
        ["mmm-yy", "General"],
      ],
      text: [
        ["", "North"],
        ["Jan-26", "10"],
        ["Feb-26", "20"],
      ],
    });
    await generate("line", true);
    expect(output().data.categories).toEqual(["Jan-26", "Feb-26"]);
    expect(output().data.series[0].values).toEqual([10, 20]);
  });

  it("still turns a date VALUE into a day, which is what a Gantt needs", async () => {
    install({
      values: [
        ["", "Start", "End"],
        ["Task", 46027, 46037],
      ],
      numberFormat: [
        ["General", "General", "General"],
        ["General", "d-mmm-yy", "d-mmm-yy"],
      ],
      text: [
        ["", "Start", "End"],
        ["Task", "5-Jan-26", "15-Jan-26"],
      ],
    });
    await generate("gantt");
    expect(output().data.dates).toBe(true);
    expect(output().data.series[0].values).toEqual([epochDay(2026, 1, 5), epochDay(2026, 1, 15)]);
  });

  it("falls back to reading the cell when the host sent no displayed text", () => {
    expect(cellText(46023, "mmm-yy", undefined, { label: true, date1904: false })).toBe("2026-01-01");
    expect(cellText(46023, "mmm-yy", "Jan-26", { label: true, date1904: false })).toBe("Jan-26");
    expect(cellText(46023, "mmm-yy", "Jan-26", { label: false, date1904: false })).toBe("2026-01-01");
    expect(cellText(null, "mmm-yy", "", { label: true, date1904: false })).toBe("");
  });
});

/* ──────────────────── r2-28: hidden and filtered rows and columns ─────────── */

describe("a hidden or filtered-out row or column is not in the chart (r2-28)", () => {
  /** A filtered table: 20 regions, the header row, and only 5 left showing. */
  const table = (): Book => {
    const values: unknown[][] = [["", "Q1", "Q2", "Q3"]];
    for (let i = 1; i <= 20; i++) values.push([`Region${i}`, i, i * 2, i * 3]);
    return { values, address: "Sheet1!A1:D21", hiddenRows: [] };
  };
  const names = () => output().data.series.map((s) => s.name);

  it("charts the 5 rows that show, not the 20 the filter hid 15 of", async () => {
    const book = table();
    book.hiddenRows = Array.from({ length: 21 }, (_, i) => i).filter((i) => i > 0 && ![3, 8, 12, 15, 20].includes(i));
    install(book);
    await generate();
    expect(names()).toEqual(["Region3", "Region8", "Region12", "Region15", "Region20"]);
    expect(note(), "the user was not told rows were left out").toContain("left out: 15 hidden or filtered-out rows");
    expect(note()).toContain("Sheet1!A1:D21");
  });

  it("leaves a hidden column out too", async () => {
    const book = table();
    book.hiddenCols = [2];
    install(book);
    await generate();
    expect(output().data.categories).toEqual(["Q1", "Q3"]);
    expect(output().data.series[0].values).toEqual([1, 3]);
    expect(note()).toContain("left out: 1 hidden column)");
  });

  it("says one row in the singular, and rows and columns together", async () => {
    const book = table();
    book.hiddenRows = [4];
    book.hiddenCols = [1, 3];
    install(book);
    await generate();
    expect(note()).toContain("left out: 1 hidden or filtered-out row and 2 hidden columns)");
  });

  it("asks nothing per row when nothing is hidden, which is nearly every selection", async () => {
    const host = install(table());
    await generate();
    expect(output().data.series).toHaveLength(20);
    expect(host.counts.getRow).toBe(0);
    expect(host.counts.getColumn).toBe(0);
    expect(note()).not.toContain("left out");
    expect(note()).not.toContain("cannot report");
  });

  it("keeps the header row when it is the only one showing", async () => {
    const book = table();
    book.hiddenRows = Array.from({ length: 20 }, (_, i) => i + 1);
    install(book);
    await generate();
    expect(output().data.series).toEqual([]);
    expect(output().data.categories).toEqual(["Q1", "Q2", "Q3"]);
  });

  it("says so, rather than charting nothing silently, when every row is hidden", async () => {
    const book = table();
    book.hiddenRows = Array.from({ length: 21 }, (_, i) => i);
    install(book);
    await generate();
    expect(note()).toMatch(/^Failed: Every row or column in Sheet1!A1:D21 is hidden or filtered out/);
    expect(el<HTMLTextAreaElement>("output").value).toBe("");
  });

  it("on a host below ExcelApi 1.2 includes every row and SAYS it cannot tell", async () => {
    // The members that report hidden rows are 1.2. Asking a 1.1 host for them
    // throws, so they are asked about first — and the older host gets the old
    // reading with a sentence, not a silence.
    const book = table();
    book.hiddenRows = [1, 2, 3];
    book.minor = 1;
    const host = install(book);
    await generate();
    expect(note()).not.toMatch(/^Failed/);
    expect(output().data.series).toHaveLength(20);
    expect(note()).toContain("cannot report hidden rows or columns");
    expect(host.loaded.some((p) => /Hidden$/.test(p))).toBe(false);
  });

  it("drops the same cells from every grid, so a date still lines up with its own format", () => {
    const kept = keepVisible(
      [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
      ],
      [true, false, true],
      [false, true, true],
    );
    expect(kept).toEqual([
      [2, 3],
      [8, 9],
    ]);
  });
});

/* ─────────────── r2-44: oversized and unbounded selections ────────────────── */

describe("a selection that is not the shape of its data (r2-44)", () => {
  const pad = (n: number, row: unknown[]) => [...row, ...Array.from({ length: n - row.length }, () => "")];

  it("does not turn the blank columns around a small table into empty categories", async () => {
    // A1:Z20 around a 3 x 4 table: 22 empty categories used to squeeze the real
    // ones into the left tenth of the plot.
    const cells = [
      pad(26, ["", "Q1", "Q2", "Q3"]),
      pad(26, ["North", 10, 20, 30]),
      pad(26, ["South", 5, 8, 13]),
      ...Array.from({ length: 17 }, () => pad(26, [])),
    ];
    install({ values: cells, address: "Sheet1!A1:Z20" });
    await generate();
    expect(output().data.categories).toEqual(["Q1", "Q2", "Q3"]);
    expect(output().data.series).toEqual([
      { name: "North", values: [10, 20, 30] },
      { name: "South", values: [5, 8, 13] },
    ]);
    // 17 trailing blank rows are no stack separators either.
    expect(output().data.series.every((s) => s.stack === undefined)).toBe(true);
  });

  it("leaves a gap the user left in the MIDDLE alone", async () => {
    install({
      values: [
        ["", "Q1", "", "Q3"],
        ["North", 10, 20, 30],
      ],
    });
    await generate();
    expect(output().data.categories).toEqual(["Q1", "", "Q3"]);
  });

  it("keeps a trailing column that has a heading and no figures yet", async () => {
    install({
      values: [
        ["", "Q1", "Q2"],
        ["North", 10, ""],
      ],
    });
    await generate();
    expect(output().data.categories).toEqual(["Q1", "Q2"]);
    expect(output().data.series[0].values).toEqual([10, null]);
  });

  it("charts the data inside a whole-column selection instead of dying on null", async () => {
    // Microsoft: cell-level properties "come back as null for unbounded reads".
    const host = install({
      values: [
        ["", "Q1", "Q2"],
        ["North", 10, 20],
        ["South", 5, 8],
      ],
      address: "Sheet1!A:C",
      unbounded: true,
    });
    await generate();
    expect(note()).not.toMatch(/^Failed/);
    expect(note()).toContain("Sheet1!A1:C3");
    expect(host.counts.usedRange).toBe(1);
    expect(output().data.categories).toEqual(["Q1", "Q2"]);
    expect(output().data.series).toEqual([
      { name: "North", values: [10, 20] },
      { name: "South", values: [5, 8] },
    ]);
  });

  it("does the same on a host that predates the values-only argument of getUsedRange", async () => {
    const host = install({
      values: [
        ["", "Q1"],
        ["North", 10],
      ],
      address: "Sheet1!A:B",
      unbounded: true,
      minor: 1,
    });
    await generate();
    expect(note()).not.toMatch(/^Failed/);
    expect(host.counts.usedRange).toBe(1);
    expect(output().data.series).toEqual([{ name: "North", values: [10] }]);
  });

  it("still finds the hidden rows of a whole-column selection, in the used range", async () => {
    install({
      values: [
        ["", "Q1"],
        ["North", 10],
        ["South", 5],
      ],
      address: "Sheet1!A:B",
      unbounded: true,
      hiddenRows: [2],
    });
    await generate();
    expect(output().data.series.map((s) => s.name)).toEqual(["North"]);
  });

  it("says there is nothing to chart when the selected columns are empty", async () => {
    install({ values: [[""]], address: "Sheet1!D:F", unbounded: true });
    await generate();
    expect(note()).toBe("Failed: There is nothing in Sheet1!D:F to chart — select the cells that hold the data.");
    expect(note()).not.toMatch(/TypeError|null/);
  });

  it("passes any other failure reading the used range through as it is", async () => {
    // Only "nothing there" is the user's to fix; a protected sheet is not.
    install({
      values: [["a"]],
      address: "Sheet1!A:A",
      unbounded: true,
      usedRangeFault: Object.assign(new Error("The workbook is protected."), { code: "AccessDenied" }),
    });
    await generate();
    expect(note()).toBe("Failed: The workbook is protected.");
  });

  it("trims in the grid helper: whitespace is blank, the middle is kept", () => {
    expect(
      trimBlankEdges([
        ["", "a", "", "b", " ", ""],
        ["x", "", "", "", "", ""],
        ["", "", "", "", "", ""],
        [" ", "", "", "", "", ""],
      ]),
    ).toEqual([
      ["", "a", "", "b"],
      ["x", "", "", ""],
    ]);
    expect(trimBlankEdges([["", ""]])).toEqual([]);
    expect(trimBlankEdges([])).toEqual([]);
  });
});
