/**
 * Excel companion pane: turns the selected range into an SSF chart JSON
 * config (same datasheet convention: row 1 = categories, column A = series).
 * Users paste it into SSF Charts' Automation box in PowerPoint and re-run
 * whenever the data changes — the feasible substitute for live data links.
 */
import type { ChartConfig, ChartKind } from "../core/types";
import { sheetToData, transposeSheet } from "../taskpane/datasheet";
import { DEFAULT_SIZE } from "../core/chart";
import { excelSerialToISO, isDateFormat } from "./serial-dates";

/* global Excel, Office */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

function rangeToConfig(
  values: unknown[][],
  kind: ChartKind,
  title: string,
  transpose: boolean,
  numberFormat?: unknown[][],
): ChartConfig {
  /**
   * A DATE CELL BECOMES AN ISO STRING, NOT ITS SERIAL NUMBER.
   *
   * `String(v)` over Excel's `values` turned every date into the number Excel
   * stores it as, so a Gantt built from a real schedule column arrived as
   * "46027" and drew a numeric axis — nothing downstream can tell a serial
   * number from a quantity. `normalizeData` reads ISO strings in Gantt
   * Start/End rows and marks the data as a calendar; this is what lets it.
   * Every other cell is untouched.
   */
  const cells = values.map((row, r) =>
    row.map((v, c) => {
      if (v == null) return "";
      if (typeof v === "number" && Number.isFinite(v) && isDateFormat(numberFormat?.[r]?.[c])) {
        return excelSerialToISO(v);
      }
      return String(v);
    }),
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

async function generate() {
  const note = $("note");
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      // `numberFormat` rides the sync that was already happening — it is what
      // separates a date from a quantity, and Excel's `values` cannot.
      range.load("values,numberFormat,address");
      await context.sync();
      const cfg = rangeToConfig(
        range.values as unknown[][],
        ($("kind") as HTMLSelectElement).value as ChartKind,
        ($("title") as HTMLInputElement).value,
        ($("transpose") as HTMLInputElement | null)?.checked ?? false,
        range.numberFormat as unknown[][],
      );
      ($("output") as HTMLTextAreaElement).value = JSON.stringify(cfg, null, 2);
      note.textContent = `Generated from ${range.address}. Paste into SSF Charts → Automation → Import.`;
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
