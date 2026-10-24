/**
 * What one cell of the selection says to the datasheet, and which cells of the
 * selection are in the chart at all.
 *
 * SEPARATE FROM `excel.ts` for the reason `serial-dates.ts` is: that module
 * wires the pane on import, so nothing in it can be unit-tested without a DOM,
 * and every decision here is a pure function of what Excel handed back.
 */
import { excelSerialToISO, formatPattern, isDateFormat } from "./serial-dates";

/**
 * IS THIS NUMBER FORMAT A PERCENT — does the cell show its value times 100?
 *
 * Only an unquoted, unescaped `%` does. `0.0"%"` and `0.0\%` print a percent
 * sign beside a number that was already in percent, which is how a column of
 * plain 3.5s is made to read as shares, and multiplying those would be the
 * mistake this exists to prevent in the other direction. The code is read the
 * way `isDateFormat` reads it, through `formatPattern`, and across ALL of its
 * sections: a format whose sections disagree about the `%` is not in common
 * use, and a conditional section (`[<1]0%;0.0`) would need the value tested
 * against its condition to settle, which nothing here has called for.
 */
export function isPercentFormat(code: unknown): boolean {
  return formatPattern(code).includes("%");
}

/**
 * A fraction as the percentage points Excel shows for it: 0.035 -> 3.5.
 *
 * Rounded to 15 significant digits because a double holds no more than that
 * exactly, and 0.035 * 100 is 3.5000000000000004, which would reach the JSON
 * box, the grid and the label as written.
 */
function percentPoints(fraction: number): number {
  return Number((fraction * 100).toPrecision(15));
}

/**
 * ONE CELL AS THE TEXT THE DATASHEET READS.
 *
 * `values` hands back what a cell STORES and the sheet reads what a cell SHOWS,
 * and the number format is the only thing that bridges them:
 *
 *  - A LABEL (the first row and the first column, which `sheetToData` takes as
 *    categories and series names) is the text Excel displays. `Range.text` is
 *    the same string the clipboard carries, so pasting a range into the pane
 *    and generating from it give the same labels: "Jan-26" stays "Jan-26", a
 *    percent header stays "10%", a year stays "2024". The date conversion below
 *    used to run on these too and printed "2026-01-01" under every month. A
 *    label whose displayed text did not come back is read like any other cell.
 *  - A DATE becomes an ISO string, not its serial number, so a Gantt built from
 *    a real schedule column draws a calendar instead of a numeric axis (nothing
 *    downstream can tell 46027 from a revenue figure).
 *  - A PERCENT is multiplied back up, because Excel stores 3.5% as 0.035 and
 *    the pane's own paste reads the displayed "3.5%" as 3.5 and stores a share
 *    as 35, never 0.35. `numberFormat.suffix` only APPENDS a "%", so a chart
 *    fed 0.035 printed "0.035%", a hundredth of what the sheet showed.
 *  - Every other cell is untouched.
 */
export function cellText(
  value: unknown,
  format: unknown,
  shown: unknown,
  how: { label: boolean; date1904: boolean },
): string {
  if (value == null) return "";
  if (how.label && typeof shown === "string") return shown;
  if (typeof value === "number" && Number.isFinite(value)) {
    if (isDateFormat(format)) return excelSerialToISO(value, how.date1904);
    if (isPercentFormat(format)) return String(percentPoints(value));
  }
  return String(value);
}

/**
 * THE SELECTION, LESS THE ROWS AND COLUMNS THAT ARE NOT IN THE CHART.
 *
 * `keepRow` and `keepCol` say which survive. A hidden or filtered-out row is
 * still a row of `values`: Excel's Insert Chart leaves it out by default, the
 * clipboard skips filtered rows, and the bridge used to put every one back.
 */
export function keepVisible<T>(grid: T[][], keepRow: boolean[], keepCol: boolean[]): T[][] {
  return grid.filter((_, r) => keepRow[r]).map((row) => row.filter((_, c) => keepCol[c]));
}

/**
 * THE GRID WITHOUT ITS TRAILING BLANK ROWS AND COLUMNS.
 *
 * A selection is a rectangle and the data in it is not always. A trailing blank
 * ROW is already harmless — `sheetToData` reads it as a stack separator with
 * nothing after it — but it still stamps `stack: 0` on every series. A trailing
 * blank COLUMN is a category with an empty label and nothing in it, and
 * `A1:Z20` around a 3 x 4 table drew 25 categories, the three real ones
 * squeezed into the left tenth of the plot. A blank column in the MIDDLE is the
 * user's own gap and is left alone, as is a column with a header and no values.
 *
 * "Blank" is what `sheetToData` calls blank: nothing but whitespace.
 */
export function trimBlankEdges(cells: string[][]): string[][] {
  const blank = (cell: string | undefined) => (cell ?? "").trim() === "";
  let rows = cells.length;
  while (rows > 0 && cells[rows - 1].every(blank)) rows--;
  const used = cells.slice(0, rows);
  const width = (row: string[]) => {
    let c = row.length;
    while (c > 0 && blank(row[c - 1])) c--;
    return c;
  };
  const cols = used.reduce((n, row) => Math.max(n, width(row)), 0);
  return used.map((row) => row.slice(0, cols));
}
