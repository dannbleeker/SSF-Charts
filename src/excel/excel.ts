/**
 * Excel companion pane: turns the selected range into an SSF chart JSON
 * config (same datasheet convention: row 1 = categories, column A = series).
 * Users paste it into SSF Charts' Automation box in PowerPoint and re-run
 * whenever the data changes — the feasible substitute for live data links.
 */
import type { ChartConfig, ChartKind } from "../core/types";
import { sheetToData, transposeSheet } from "../taskpane/datasheet";
import { DEFAULT_SIZE } from "../core/chart";
import { cellText, keepVisible, trimBlankEdges } from "./cell-text";
import { isDate1904 } from "./serial-dates";

/* global Excel, Office */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/**
 * The selected range as the bridge reads it: what the cells hold, what they
 * show, and what the workbook says about how to read them. Hidden and filtered
 * rows and columns are already gone from every grid.
 */
interface SelectedRange {
  values: unknown[][];
  numberFormat?: unknown[][];
  text?: unknown[][];
  address: string;
  date1904: boolean;
  /** How many rows and columns of the selection were left out for being hidden. */
  hiddenRows: number;
  hiddenCols: number;
  /** False where the host cannot say which are hidden, so none were left out. */
  hiddenKnown: boolean;
}

function rangeToConfig(sel: SelectedRange, kind: ChartKind, title: string, transpose: boolean): ChartConfig {
  // Each cell's text is decided in `cellText` (dates, percents, labels); what
  // is left to settle here is the shape. A selection larger than its data
  // leaves blank edges that would become empty categories, so they are trimmed.
  const cells = trimBlankEdges(
    sel.values.map((row, r) =>
      row.map((v, c) =>
        cellText(v, sel.numberFormat?.[r]?.[c], sel.text?.[r]?.[c], {
          label: r === 0 || c === 0,
          date1904: sel.date1904,
        }),
      ),
    ),
  );
  // The datasheet convention is row 1 = categories, column A = series. A user
  // whose sheet is laid out the other way (series across the top) would silently
  // get a transposed chart — the transpose toggle swaps axes before parsing.
  const sheet = transpose ? transposeSheet({ cells }) : { cells };
  const totals = new Set<number>();
  const data = sheetToData(sheet, kind === "waterfall" ? totals : undefined);
  // Name a row whose column-A cell was blank. `sheetToData` used to do this and
  // no longer does — an invented name is wrong where its answer is written back
  // into a chart the user already authored (see the note at that call) and right
  // here, where the input is an arbitrary spreadsheet selection and a nameless
  // row is a gap in the user's sheet rather than a decision they made.
  data.series.forEach((s, i) => {
    if (!s.name) s.name = `Series ${i + 1}`;
  });
  return {
    kind,
    data,
    ...DEFAULT_SIZE,
    title: title || undefined,
    waterfall: kind === "waterfall" ? { totalIndices: [...totals] } : undefined,
  };
}

/**
 * READ THE SELECTION, and everything the host can say about how to read it.
 *
 * `numberFormat` is what separates a date from a quantity and a percent from a
 * fraction, which Excel's `values` cannot; `text` is what the cell shows.
 *
 * WHAT NEEDS EXCELAPI 1.2 IS ASKED ABOUT, NOT REQUIRED: whether a row or column
 * is hidden, the workbook's date system, and the values-only argument of
 * `getUsedRange`. Each makes a result right without being what the pane is
 * for, and the manifest's floor stays at 1.1 (see `manifest-excel.xml` on why
 * that number cannot be hedged). A host that lacks them gets the older reading
 * and, where that can show, says so in the note rather than in silence.
 */
async function readSelection(context: Excel.RequestContext): Promise<SelectedRange> {
  const modern = !!Office.context.requirements?.isSetSupported("ExcelApi", "1.2");
  let range = context.workbook.getSelectedRange();
  range.load("values,numberFormat,text,address");
  // DATE(2000,1,1) is the one thing that comes out differently on the 1904 date
  // system, and it rides the sync that was already happening.
  const jan2000 = modern ? context.workbook.functions.date(2000, 1, 1) : undefined;
  jan2000?.load("value");
  await context.sync();

  /**
   * A WHOLE COLUMN OR ROW HAS NO CELL DATA TO READ. Microsoft's unbounded-range
   * rule: `values`, `numberFormat` and `text` come back null, so selecting A:D
   * by its column headers — the natural way to say "all of this" — died on
   * `null.map` with a raw TypeError. The used range inside the selection is
   * bounded and holds the data, so the read is made again from that.
   */
  if (range.values == null) {
    const asked = range.address;
    range = modern ? range.getUsedRange(true) : range.getUsedRange();
    range.load("values,numberFormat,text,address");
    let empty = false;
    try {
      await context.sync();
    } catch (err) {
      // `getUsedRange` throws ItemNotFound when nothing in the selection has a value.
      if ((err as { code?: string } | null)?.code !== "ItemNotFound") throw err;
      empty = true;
    }
    if (empty) throw new Error(`There is nothing in ${asked} to chart — select the cells that hold the data.`);
  }

  const values = range.values as unknown[][];
  const width = values[0]?.length ?? 0;
  let keepRow = values.map(() => true);
  let keepCol = Array.from({ length: width }, () => true);
  if (modern) {
    // Hidden is asked of the range first: `rowHidden` is true for all, false for
    // none and null for some, so a selection with nothing hidden — nearly all
    // of them — never pays for a request per row.
    range.load("rowHidden,columnHidden");
    await context.sync();
    const someRows = range.rowHidden !== false;
    const someCols = range.columnHidden !== false;
    if (someRows || someCols) {
      const rows = someRows ? values.map((_, i) => range.getRow(i)) : [];
      const cols = someCols ? keepCol.map((_, j) => range.getColumn(j)) : [];
      rows.forEach((row) => row.load("rowHidden"));
      cols.forEach((col) => col.load("columnHidden"));
      await context.sync();
      if (someRows) keepRow = rows.map((row) => !row.rowHidden);
      if (someCols) keepCol = cols.map((col) => !col.columnHidden);
    }
  }

  const kept = keepVisible(values, keepRow, keepCol);
  if (!kept.length || !kept[0]?.length) {
    throw new Error(`Every row or column in ${range.address} is hidden or filtered out, so there is nothing to chart.`);
  }
  const numberFormat = range.numberFormat as unknown[][] | undefined;
  const text = range.text as unknown[][] | undefined;
  return {
    values: kept,
    numberFormat: numberFormat && keepVisible(numberFormat, keepRow, keepCol),
    text: text && keepVisible(text, keepRow, keepCol),
    address: range.address,
    date1904: isDate1904(jan2000?.value),
    hiddenRows: keepRow.filter((k) => !k).length,
    hiddenCols: keepCol.filter((k) => !k).length,
    hiddenKnown: modern,
  };
}

/** What the note says about the rows and columns the selection did not bring in. */
function leftOut(sel: SelectedRange): string {
  if (!sel.hiddenKnown)
    return " (this Excel cannot report hidden rows or columns, so any in the selection are included)";
  const count = (n: number, what: string) => (n ? `${n} ${what}${n === 1 ? "" : "s"}` : "");
  const parts = [count(sel.hiddenRows, "hidden or filtered-out row"), count(sel.hiddenCols, "hidden column")];
  const named = parts.filter(Boolean).join(" and ");
  return named ? ` (left out: ${named})` : "";
}

async function generate() {
  const note = $("note");
  try {
    await Excel.run(async (context) => {
      const sel = await readSelection(context);
      const cfg = rangeToConfig(
        sel,
        ($("kind") as HTMLSelectElement).value as ChartKind,
        ($("title") as HTMLInputElement).value,
        ($("transpose") as HTMLInputElement | null)?.checked ?? false,
      );
      ($("output") as HTMLTextAreaElement).value = JSON.stringify(cfg, null, 2);
      note.textContent = `Generated from ${sel.address}${leftOut(sel)}. Paste into SSF Charts → Automation → Import.`;
    });
  } catch (err) {
    note.textContent = `Failed: ${err instanceof Error ? err.message : String(err)}`;
  }
}

function wire() {
  const inExcel = typeof Excel !== "undefined" && !!Office.context?.host;
  ($("generate") as HTMLButtonElement).disabled = !inExcel;
  if (!inExcel) {
    $("note").textContent = "Not running inside Excel — sideload manifest-excel.xml to use the data bridge.";
  }
  $("generate").addEventListener("click", () => void generate());
  $("copy").addEventListener("click", () => {
    const out = ($("output") as HTMLTextAreaElement).value;
    if (!out) return;
    /**
     * SAY WHETHER IT COPIED, and this used to say nothing either way.
     *
     * It read `void navigator.clipboard?.writeText(out).catch(() => {})`. Two
     * silences in one line. The clipboard API rejects routinely — a document
     * that is not focused, a permission the browser withholds, an add-in frame
     * without clipboard-write — and the user, who pressed Copy and has a
     * textarea full of JSON they now believe is on their clipboard, is told
     * nothing and pastes whatever was there before.
     *
     * The `?.` hid a second one: where `navigator.clipboard` does not exist at
     * all, optional chaining short-circuits the whole expression, so the catch
     * never ran and a browser with no clipboard support also failed quietly.
     *
     * This file already had the right pattern six lines up — `generate()` puts
     * both of its outcomes in `#note`. The copy button now does the same, and
     * the failure names the way out, because selecting the textarea by hand is
     * something the user can actually do.
     */
    const note = $("note");
    const clipboard = navigator.clipboard;
    if (!clipboard) {
      note.textContent = "This browser will not let an add-in write the clipboard — select the JSON above and copy it.";
      return;
    }
    void clipboard.writeText(out).then(
      () => {
        note.textContent = "Copied. Paste into SSF Charts → Automation → Import.";
      },
      (err: unknown) => {
        note.textContent =
          "Could not copy — select the JSON above and copy it by hand. " +
          `(${err instanceof Error ? err.message : String(err)})`;
      },
    );
  });
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => wire());
} else {
  wire();
}
