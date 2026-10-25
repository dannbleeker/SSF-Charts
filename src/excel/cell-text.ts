/**
 * What one cell of the selection says to the datasheet, and which cells of the
 * selection are in the chart at all.
 *
 * SEPARATE FROM `excel.ts` for the reason `serial-dates.ts` is: that module
 * wires the pane on import, so nothing in it can be unit-tested without a DOM,
 * and every decision here is a pure function of what Excel handed back.
 */
import { parseDateToken } from "../core/format";
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
 * THE DAY A DATE LABEL'S DISPLAYED TEXT NAMES — does it agree with the day the
 * cell holds, as far as the number format shows it?
 *
 * The pane's date reader is month-first for a dashed date, so a Danish workbook
 * whose headers show "05-01-2026" (5 January) handed over that spelling and the
 * line chart read 1 May: dates out of order, the time axis abandoned or wrong.
 * The cell also holds the serial, which has no locale, so the serial is the
 * truth and the text only wins when it agrees.
 *
 * AGREES AT THE RESOLUTION THE FORMAT SHOWS, which is why this does not compare
 * whole days. "Jan-26" is what `mmm-yy` prints for 1 January and for 31 January
 * alike, and the reader takes it as the 1st; holding it to the day would turn
 * every month-end header (EOMONTH is how finance writes them) into
 * "2026-01-31" — the very thing keeping the displayed text exists to prevent.
 * So only the parts the format has a token for are compared: year, month and day
 * of the month. A format with no year ("mmmm") is not held to one.
 *
 * Text the reader does not read as a date at all ("Q1 2026", "2026", "Monday") is
 * the user's label and is kept — except when it is spelled like a numeric date, three
 * numbers round separators, which the reader refuses when it cannot tell day
 * from month ("05-01-2026"). That the CELL is a date is already settled by its
 * format, so there the serial speaks.
 */
const NUMERIC_DAY = /^\d{1,4}[-/. ]\d{1,2}[-/. ]\d{1,4}(?: \d{1,2}:\d{2}(?::\d{2})?)?$/;

function displayNamesTheDay(shown: string, iso: string, format: unknown): boolean {
  const day = parseDateToken(shown);
  if (day == null) return !NUMERIC_DAY.test(shown.trim());
  const [year, month, dayOfMonth] = iso.split("-").map(Number);
  const read = new Date(day * 86400000);
  const pattern = formatPattern(format).toLowerCase();
  if (/y/.test(pattern) && read.getUTCFullYear() !== year) return false;
  if (/m/.test(pattern) && read.getUTCMonth() + 1 !== month) return false;
  // `d` and `dd` are the day of the month; `ddd` and `dddd` are weekday names.
  if (/(?<!d)d{1,2}(?!d)/.test(pattern) && read.getUTCDate() !== dayOfMonth) return false;
  return true;
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
 *    The one exception is a DATE label whose text the pane would read as a
 *    different day than the cell holds (`displayNamesTheDay`): there the ISO date
 *    of the serial is delivered instead, because a header that DISPLAYS as
 *    "05-01-2026" is 5 January in Danish Excel and 1 May to the date reader.
 *  - A DATE becomes an ISO string, not its serial number, so a Gantt built from
 *    a real schedule column draws a calendar instead of a numeric axis (nothing
 *    downstream can tell 46027 from a revenue figure). A number that is no
 *    calendar day at all (`excelSerialToISO` answers "") stays the number it is.
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
  const date = typeof value === "number" && Number.isFinite(value) && isDateFormat(format);
  if (how.label && typeof shown === "string") {
    if (!date) return shown;
    const iso = excelSerialToISO(value, how.date1904);
    return iso && !displayNamesTheDay(shown, iso, format) ? iso : shown;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    if (date) return excelSerialToISO(value, how.date1904) || String(value);
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
