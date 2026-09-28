import type { ChartData } from "../core/types";
import { isTotalToken } from "../core/layout/waterfall";
import { GANTT_DATE_ROW, parseDateToken } from "../core/format";

export interface SheetModel {
  /** Raw cell text; row 0 = category names, column 0 = series names. */
  cells: string[][];
}

// --- Formulas ----------------------------------------------------------------

/** Column letters → 0-based index (A=0, Z=25, AA=26 …). */
function colIndex(letters: string): number {
  let v = 0;
  for (const ch of letters.toUpperCase()) v = v * 26 + (ch.charCodeAt(0) - 64);
  return v - 1;
}

/** 0-based index → column letters (0=A, 25=Z, 26=AA …), colIndex's inverse. */
function colLetters(index: number): string {
  let s = "";
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) s = String.fromCharCode(65 + (i % 26)) + s;
  return s;
}

/**
 * The grid's rows and columns must address the SAME number of cells, or a
 * transpose cannot be total — a row with no column to become is a reference the
 * sheet is unable to write down.
 *
 * They did not. Rows go to 999 (`[0-9]{1,3}` in every reference pattern) and
 * columns stopped at ZZ (`[A-Za-z]{1,2}`, index 701), so `=SUM(B2:B999)` — the
 * whole-column sum an Excel user types without thinking, and one this engine
 * evaluates correctly — came back from Transpose as `=SUM(B2:ALK2)`, three
 * letters, which no pattern here matches. The formula silently evaluated to
 * nothing and the total column went BLANK. Exactly the failure transposeFormula
 * was written to prevent, arriving from the other side: it is not enough for the
 * reference to MOVE correctly, the destination has to be sayable.
 *
 * So columns read three letters, and are capped at the index a 3-digit row can
 * become. Both spaces are now 0..998, which makes the swap total: every
 * reference this parser accepts transposes to one it also accepts, and
 * transposing twice is the identity.
 */
const LAST_INDEX = 998;
/**
 * One A1-style reference — two capture groups, shared by the evaluator's three
 * patterns and by `transposeFormula`, so the two cannot drift apart. That they
 * match the SAME set of references is what the totality above rests on.
 *
 * The trailing `(?![0-9])` makes it a whole token rather than a prefix of one.
 * `A1000` used to match its first four characters: the evaluator read "A100",
 * choked on the leftover "0" and answered null, while the transposer rewrote
 * that prefix and left the digit behind — so `=A1000` came back as `=CV10`, a
 * real cell holding a real number where there had been a blank. Wrong in the
 * worse direction than blank.
 */
const REF = "([A-Za-z]{1,3})([0-9]{1,3})(?![0-9])";
/** A reference's cell, or null when it is outside the address space above. */
function refCell(letters: string, digits: string): { row: number; col: number } | null {
  const col = colIndex(letters);
  const row = Number(digits) - 1;
  return col > LAST_INDEX || row > LAST_INDEX ? null : { row, col };
}

/**
 * A cell's number, accepting the US thousands grouping Excel copies ("1,234").
 * A comma ANYWHERE ELSE is not a separator we can read: stripping every comma
 * turned a European "1.234,5" into 1.2345 (a silent 1000× error) and "1,5" into
 * 15. Refuse those instead — a visible gap beats a wrong number.
 */
function numericValue(text: string): number {
  if (US_GROUPED.test(text)) return Number(text.replace(/,/g, ""));
  return text.includes(",") ? NaN : Number(text);
}

/** "1,234" / "1,234.5" — the US thousands grouping Excel copies. */
const US_GROUPED = /^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/;

/** "1.234" / "1.234.567,89" — dot groups, optionally with a comma decimal. */
const EU_GROUPED = /^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/;

/** "987,5" — a plain comma decimal, no grouping. */
const EU_DECIMAL = /^[+-]?\d+,\d+$/;

/**
 * A TRAILING PERCENT SIGN, SPLIT OFF SO THE NUMBER UNDERNEATH CAN BE READ.
 *
 * All three patterns above are anchored, so `"12,5%"` matches none of them —
 * and a share table is the single most likely thing to be pasted in here. The
 * consequence, measured 2026-09-28 on a sheet whose every number carried a
 * percent sign: `looksEuropean` returned FALSE where the identical sheet
 * without the signs returned true, and `fromEuropeanNumber("12,5%")` handed
 * back `"12,5%"` untouched. The block was then read as American, so the decimal
 * comma was never converted.
 *
 * `rawCellValue` already strips a trailing "%" before parsing — the two places
 * that care about percent cells disagreed about whether they exist. This is the
 * same shape as the CHAR(11)/CHAR(10) lesson: a rule written for one spelling
 * of a cell silently skips the others.
 *
 * Applied to BOTH sides of the evidence, not just the European one. Counting
 * "12,5%" as European while ignoring "1,234%" as American would bias the vote
 * rather than fix it.
 */
const percentBody = (t: string): string => t.replace(/\s*%$/, "");

/**
 * A CLIPBOARD BLOCK READ AS QUOTED TSV, which is what Excel actually writes.
 *
 * This used to be `text.replace(/\r/g, "").split("\n")` and then `split("\t")`
 * per row — unquoted TSV. Excel's `text/plain` clipboard quotes any cell whose
 * content holds a line break, a tab or a double quote, and keeps the break
 * inside the quotes, exactly as it does for CSV. Nothing here knew that.
 *
 * So a wrapped header — the most ordinary thing in a finance table, Alt+Enter
 * inside a cell — arrived as `"Revenue\n(EURm)"\t2024\t2025` and the break was
 * read as the end of the row. One pasted row became two, every later row landed
 * on the wrong series, and the sheet said nothing. Same family as the
 * CHAR(11)/CHAR(10) lesson: the character Alt+Enter really produces is the one
 * nothing was looking for.
 *
 * AND THE TRAILING BLANK, which is the second defect in the same three lines.
 * Excel appends exactly ONE row terminator, but the old code popped empty lines
 * in a `while`. In a single-column copy a blank CELL is an empty line — the
 * comment it sat under said precisely that, and used it to justify keeping
 * INTERIOR blanks before eating the trailing ones. Copying `10, 20, 30, ,` (two
 * deliberate blanks at the end) wrote three cells and left the sheet's old
 * fourth and fifth values showing underneath. Exactly one terminator is
 * dropped now, because exactly one is what is appended.
 */
export function parseClipboardGrid(text: string): string[][] {
  const s = String(text ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (quoted) {
      // `""` inside a quoted cell is one literal quote — Excel's own escape.
      if (ch === '"' && s[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') quoted = false;
      else cell += ch;
      continue;
    }
    // A quote only OPENS a cell at its start. A stray quote mid-cell is data,
    // which is how a cell like `5" pipe` survives.
    if (ch === '"' && cell === "") quoted = true;
    else if (ch === "\t") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else cell += ch;
  }
  row.push(cell);
  rows.push(row);
  const last = rows[rows.length - 1];
  if (rows.length > 1 && last.length === 1 && last[0] === "") rows.pop();
  return rows;
}

/**
 * IS THIS PASTE WRITTEN IN THE EUROPEAN CONVENTION — dot for thousands, comma
 * for the decimal?
 *
 * The comma half was already refused, deliberately, by `numericValue`: it reads
 * a US grouping and turns anything else with a comma into a visible gap, on the
 * stated rule that a gap beats a wrong number. The DOT half had no such care,
 * and there `Number()` is only too willing — so one paste out of a Danish Excel
 * came apart three different ways:
 *
 *     1.234   ->    1.234    a thousand two hundred, read as one and a bit
 *     2.500   ->    2.5      likewise
 *     987,5   ->    null     correctly refused
 *
 * Two cells silently wrong by a factor of a thousand, one visibly refused, on a
 * sheet whose only difference from a US one is the punctuation.
 *
 * PER CELL IT CANNOT BE DECIDED. "1.234" is a perfectly good en-US decimal and
 * always will be. Across a PASTE it usually can, because the evidence arrives in
 * the same clipboard block: a comma that is not a thousands group is a decimal
 * comma, and a decimal comma means the dots are groups.
 *
 * So this asks for evidence in BOTH directions and answers only when one side is
 * unopposed. A block carrying both — a mixed paste, or a category named "1,5" —
 * gets today's reading, which is the honest outcome when the sheet contradicts
 * itself.
 */
export function looksEuropean(cells: string[][]): boolean {
  let european = false;
  let american = false;
  for (const row of cells) {
    for (const cell of row) {
      // A trailing "%" is stripped before ANY of these are asked, because every
      // pattern here is anchored and a share table is the likeliest paste of
      // all. See `percentBody`.
      const t = percentBody(String(cell ?? "").trim());
      if (!t) continue;
      if (EU_GROUPED.test(t) && t.includes(",")) european = true;
      else if (EU_DECIMAL.test(t) && !US_GROUPED.test(t)) european = true;
      // A US grouping, or a dot decimal that CANNOT be a group because it does
      // not carry exactly three digits ("987.5", "1.2"). A bare "1.234" is not
      // evidence either way and is deliberately not counted.
      if (US_GROUPED.test(t)) american = true;
      else if (/^[+-]?\d+\.\d+$/.test(t) && !/^[+-]?\d{1,3}(\.\d{3})+$/.test(t)) american = true;
    }
  }
  return european && !american;
}

/**
 * One cell rewritten from the European convention into the one this sheet
 * stores — which is the plain machine form `Number()` reads.
 *
 * REWRITTEN AT PASTE TIME, INTO THE GRID, and that is the point rather than an
 * implementation detail. The alternative was to teach `numericValue` the
 * convention and convert on every read, which would make the same characters
 * mean different things depending on cells elsewhere, and would hide the
 * decision where nobody could see or correct it. Here the user watches 1.234
 * become 1234 in the cell, and if the guess is ever wrong the value is right
 * there to fix.
 *
 * Anything that is not a number in that convention — text, a date, a formula —
 * comes back untouched.
 */
export function fromEuropeanNumber(text: string): string {
  const raw = String(text ?? "").trim();
  // The percent sign is carried through rather than dropped: `rawCellValue`
  // strips it later, and a share table must still READ as a share in the grid.
  // Without this split "12,5%" matched nothing and came back unconverted, so
  // the comma survived into a cell the rest of the sheet reads as a number.
  const t = percentBody(raw);
  const pct = raw.slice(t.length);
  if (EU_GROUPED.test(t)) return t.replace(/\./g, "").replace(",", ".") + pct;
  if (EU_DECIMAL.test(t)) return t.replace(",", ".") + pct;
  return text;
}

/**
 * The number a raw cell holds, and whether it got there as a calendar date.
 *
 * The single place a typed cell becomes a number. `parseRow` and `cellNumeric`
 * each carried their own version and they had drifted: only parseRow stripped a
 * trailing "%" (how Excel copies a share table, the canonical source for a 100%
 * chart) and only parseRow read a date, so "35%" was 35 to the chart and an
 * error to `=SUM` over it, and a Gantt duration `=B3-B2` over two ISO dates came
 * back blank.
 */
function rawCellValue(raw: string): { n: number; date: boolean } {
  const num = numericValue(raw.replace(/\s*%$/, ""));
  if (Number.isFinite(num)) return { n: num, date: false };
  const day = parseDateToken(raw);
  return day != null ? { n: day, date: true } : { n: NaN, date: false };
}

/** Shared state for one formula evaluation pass: cycle guard + value cache. */
interface EvalContext {
  visiting: Set<string>;
  memo: Map<string, number>;
}

/** Numeric value of a cell, following "=" formulas with cycle protection. */
function cellNumeric(cells: string[][], row: number, col: number, ctx: EvalContext): number {
  const key = `${row},${col}`;
  // Memoised: the cycle guard is added before a descent and removed after, so a
  // formula naming the same cell twice re-walked its whole subtree twice, and a
  // chain of such cells cost 2^depth evaluations — a ~25-row sheet froze the
  // pane for minutes, on the UI thread, on every debounced keystroke. Only
  // FINITE results are cached: a NaN can come from a cycle, and a cycle's answer
  // depends on the path that reached it.
  const hit = ctx.memo.get(key);
  if (hit !== undefined) return hit;
  if (ctx.visiting.has(key)) return NaN; // circular reference
  const raw = (cells[row]?.[col] ?? "").trim();
  let v: number;
  if (raw === "") {
    v = 0;
  } else if (raw.startsWith("=")) {
    ctx.visiting.add(key);
    v = evaluateFormula(cells, raw.slice(1), ctx.visiting, ctx.memo) ?? NaN;
    ctx.visiting.delete(key);
  } else {
    // A non-numeric NON-blank cell (text, an error token) is not a value — NaN
    // propagates as an error, the same stance parseRow takes when it returns
    // null for the same cell in place. (A BLANK cell stays 0 above, so SUM still
    // treats gaps as zero — Excel's convention.) The old silent 0 meant a stray
    // "n/a" in a referenced cell vanished into a computed total.
    v = rawCellValue(raw).n;
  }
  if (Number.isFinite(v)) ctx.memo.set(key, v);
  return v;
}

/**
 * Evaluate a spreadsheet formula against the datasheet: A1-style refs
 * (row 1 = the category header row), + - * / ( ), and SUM/AVG/MIN/MAX
 * over ranges like B2:E2. Returns null on parse errors or cycles.
 */
export function evaluateFormula(
  cells: string[][],
  expr: string,
  visiting: Set<string> = new Set(),
  memo: Map<string, number> = new Map(),
): number | null {
  const ctx: EvalContext = { visiting, memo };
  const s = expr.replace(/\s+/g, "");
  let i = 0;

  const ref = () => {
    const m = new RegExp(`^${REF}`).exec(s.slice(i));
    if (!m) return null;
    // Not advanced past a reference outside the address space: `factor` then
    // takes its unparseable-token branch, which is what such a token got when
    // the pattern was two letters wide and could not match it at all.
    const at = refCell(m[1], m[2]);
    if (!at) return null;
    i += m[0].length;
    return cellNumeric(cells, at.row, at.col, ctx);
  };
  // A blank cell in a range is `null`, not 0: Excel's MIN/MAX/AVG ignore empty
  // cells (only SUM treats them as 0, which it still does below). Distinguish a
  // blank from a real 0 by looking at the raw cell before cellNumeric coerces it.
  // Rows and columns the SHEET actually has. A reference past the end reads as
  // blank whatever we do, so clamping to the grid changes no answer — it only
  // stops the walk being as large as the reference someone typed. `=SUM(A1:ZZ999)`
  // is 702 x 999 cells, and it did not merely take a while: it threw
  // "Maximum call stack size exceeded" out of `sheetToData`, through
  // `currentConfig`, and into the pane's live preview. One typed cell.
  const lastRow = () => cells.length - 1;
  const lastCol = () => cells.reduce((n, row) => Math.max(n, row.length - 1), -1);
  const rangeValues = (m: RegExpExecArray): (number | null)[] => {
    const c1 = colIndex(m[1]);
    const r1 = Number(m[2]) - 1;
    const c2 = colIndex(m[3]);
    const r2 = Number(m[4]) - 1;
    const out: (number | null)[] = [];
    const rEnd = Math.min(Math.max(r1, r2), lastRow());
    const cEnd = Math.min(Math.max(c1, c2), lastCol());
    for (let r = Math.max(0, Math.min(r1, r2)); r <= rEnd; r++)
      for (let c = Math.max(0, Math.min(c1, c2)); c <= cEnd; c++)
        out.push((cells[r]?.[c] ?? "").trim() === "" ? null : cellNumeric(cells, r, c, ctx));
    return out;
  };

  function expr0(): number {
    let v = term();
    while (s[i] === "+" || s[i] === "-") {
      const op = s[i++];
      const r = term();
      v = op === "+" ? v + r : v - r;
    }
    return v;
  }
  function term(): number {
    let v = factor();
    while (s[i] === "*" || s[i] === "/") {
      const op = s[i++];
      const r = factor();
      v = op === "*" ? v * r : v / r;
    }
    return v;
  }
  function factor(): number {
    if (s[i] === "-") {
      i++;
      return -factor();
    }
    if (s[i] === "(") {
      i++;
      const v = expr0();
      if (s[i] === ")") i++;
      return v;
    }
    const fn = /^(SUM|AVG|MIN|MAX)\(/i.exec(s.slice(i));
    if (fn) {
      i += fn[0].length;
      const args: (number | null)[] = [];
      while (i < s.length && s[i] !== ")") {
        const range = new RegExp(`^${REF}:${REF}`).exec(s.slice(i));
        if (range) {
          i += range[0].length;
          // Appended, never spread. `push(...values)` passes one argument per
          // value, so a range wider than the engine's argument limit overflows
          // the stack — the crash above arrived here, not in the walk.
          for (const v of rangeValues(range)) args.push(v);
        } else {
          // A BARE cell reference must contribute the cell's EMPTINESS, not
          // expr0()'s 0 — otherwise MIN/MAX/AVG over comma-separated args counted a
          // blank as a real 0 while the range form correctly ignores it, so
          // =MIN(B2,C2,D2) returned 0 where =MIN(B2:D2) returned 10.
          const one = new RegExp(`^${REF}\\s*(?=[,)])`).exec(s.slice(i));
          const at = one && refCell(one[1], one[2]);
          if (one && at) {
            i += one[0].length;
            const { row: rr, col: cc } = at;
            args.push((cells[rr]?.[cc] ?? "").trim() === "" ? null : cellNumeric(cells, rr, cc, ctx));
          } else {
            args.push(expr0());
          }
        }
        if (s[i] === ",") i++;
      }
      i++;
      if (!args.length) return NaN;
      const name = fn[1].toUpperCase();
      // SUM counts a blank range cell as 0 (Excel convention); MIN/MAX/AVG ignore
      // it — an all-negative range with a gap must not report a max of 0.
      if (name === "SUM") return args.reduce((a: number, b) => a + (b ?? 0), 0);
      const nums = args.filter((v): v is number => v != null);
      if (!nums.length) return NaN;
      if (name === "AVG") return nums.reduce((a, b) => a + b, 0) / nums.length;
      // Reduced, not spread, for the same reason as the append above: a big
      // enough range makes `Math.min(...nums)` an argument list.
      if (name === "MIN") return nums.reduce((a, b) => (b < a ? b : a), nums[0]);
      return nums.reduce((a, b) => (b > a ? b : a), nums[0]);
    }
    const num = /^[0-9]*\.?[0-9]+/.exec(s.slice(i));
    if (num) {
      i += num[0].length;
      return Number(num[0]);
    }
    const r = ref();
    if (r != null) return r;
    i++;
    return NaN;
  }

  const result = expr0();
  return i >= s.length && Number.isFinite(result) ? result : null;
}

/** think-cell's special datasheet rows, matched by row name. */
const HUNDRED_ROW = /^100\s*%\s*=?$/i;
// "extent" is REQUIRED: scatter/bubble use a series literally named "X"
// (layout/scatter.ts matches /^x$/i), and a lenient pattern here silently ate it
// into the Mekko-only xExtent field on every pane round-trip. dataToSheet only
// ever writes the full "X extent", so nothing legitimate depends on the short form.
const XEXTENT_ROW = /^x\s*extent$/i;

// Every Gantt row whose values are calendar positions — the same set layoutGantt
// feeds through its time scale (core/layout/gantt.ts). With only start/end/
// milestone here, re-opening your own calendar Gantt showed the raw epoch day
// ("20494") where you had typed "2026-02-10", so the row could only be edited in
// epoch days. "% complete" and "After" are NOT dates and stay out.
// Declared in `core/format.ts` beside `parseDateToken`, because the engine now
// needs the same answer: a config authored anywhere but this grid — the skill, a
// hand-written JSON, a POWERCHART_CONFIG tag — was losing its dates in silence.

/**
 * An epoch-day value as "YYYY-MM-DD", or null when it is not a calendar day.
 *
 * `new Date(v * 86400000).toISOString()` THROWS RangeError past the Date range,
 * and a Gantt row can hold anything a cell can: an epoch-second or -millisecond
 * timestamp pasted from an export, or a typo. The exception escaped through
 * applyConfig into the template/selection change listeners, so the chart never
 * loaded and nothing was reported. Values that survive Date but leave its
 * four-digit-year window ("20260105" → "+057440-04-…") are equally unusable, so
 * the check is on the rendered string, not on the number.
 */
function isoDay(v: number): string | null {
  if (!Number.isFinite(v) || Math.abs(v) > 1e8) return null;
  const d = new Date(v * 86400000);
  if (Number.isNaN(d.getTime())) return null;
  const s = d.toISOString().slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : null;
}

/**
 * Reorder series so every stack group is contiguous, keeping the chart the same.
 *
 * The grid can only say which series stack together by putting them NEXT TO each
 * other and separating groups with a blank row — so a group interrupted in the
 * series list, `stack: [0, 1, 0]`, cannot be written down at all. `dataToSheet`
 * emitted a separator whenever the stack changed from the previous series and
 * `sheetToData` then renumbered by separator count, so [0,1,0] came back as
 * [0,1,2]: two columns became three, and merely loading such a chart into the
 * pane restructured it on the next update.
 *
 * Grouping by first appearance draws the identical chart — within a group the
 * paint order is unchanged, and groups keep the column order they already had —
 * so this loses nothing except an ordering the sheet was never able to express.
 * `undefined` counts as group 0, which is what `sheetToData` assigns it anyway.
 *
 * Idempotent, and returns the SAME object when nothing needs moving, so a caller
 * can lean on identity.
 */
export function contiguousStacks(data: ChartData): ChartData {
  if (!data.series.some((s) => s.stack != null)) return data;
  const groups = new Map<number, ChartData["series"]>();
  for (const s of data.series) {
    const key = s.stack ?? 0;
    const bucket = groups.get(key);
    if (bucket) bucket.push(s);
    else groups.set(key, [s]);
  }
  const series = [...groups.values()].flat();
  return series.every((s, i) => s === data.series[i]) ? data : { ...data, series };
}

export function dataToSheet(input: ChartData): SheetModel {
  const data = contiguousStacks(input);
  const cells: string[][] = [["", ...data.categories]];
  const numRow = (name: string, values: (number | null)[]) => [
    name,
    ...data.categories.map((_, i) => (values[i] == null ? "" : String(values[i]))),
  ];
  if (data.hundredPercent) cells.push(numRow("100%=", data.hundredPercent));
  if (data.xExtent) cells.push(numRow("X extent", data.xExtent));
  let prevStack: number | undefined;
  for (const s of data.series) {
    // Blank separator row between stack groups (clustered-stacked round trip).
    //
    // Normalised with `?? 0`, which is what `contiguousStacks` above and
    // `sheetToData` below both do. Comparing the RAW values instead meant a
    // boundary next to a series with no `stack` was never written: a config of
    // `[{name:"Plan"}, {name:"Actual", stack:1}]` — the ordinary way to write a
    // two-group clustered-stacked chart, since the first group needs no marker
    // — came back out of the grid with both stacks gone, and drew ONE stacked
    // column where the author had two side by side.
    const stack = s.stack ?? 0;
    if (prevStack != null && stack !== prevStack) {
      cells.push(Array.from({ length: data.categories.length + 1 }, () => ""));
    }
    prevStack = stack;
    // Calendar Gantt round trip: show epoch-day values as ISO dates again.
    const asDate = data.dates && GANTT_DATE_ROW.test(s.name.trim());
    cells.push([
      s.name,
      // One cell per CATEGORY, as `numRow` above already does. Walking the
      // values instead produced a row shorter than the header whenever a series
      // carried fewer values than there are categories — and `mountDatasheet`
      // renders one input per cell, so those categories had no cell at all:
      // nothing to type in, and `handleNav`'s ArrowRight with nowhere to go.
      // The data a user could not reach was the data they most likely opened
      // the sheet to fill in.
      //
      // A date row cell that is not a representable day falls back to its raw
      // number, which round-trips as a number rather than taking the load down.
      ...data.categories.map((_, i) => {
        const v = s.values[i];
        return v == null ? "" : ((asDate ? isoDay(v) : null) ?? String(v));
      }),
    ]);
  }
  return { cells };
}

export function sheetToData(sheet: SheetModel, waterfallTotals?: Set<number>): ChartData {
  const [head = [], ...rows] = sheet.cells;
  const nCats = Math.max(0, head.length - 1);
  const categories = Array.from({ length: nCats }, (_, i) => head[i + 1] ?? `Cat ${i + 1}`);
  let sawDate = false;
  const parseRow = (r: string[], catIdxTotals?: Set<number>) =>
    Array.from({ length: nCats }, (_, i) => {
      const raw = (r[i + 1] ?? "").trim();
      if (raw === "") return null;
      if (catIdxTotals && isTotalToken(raw)) {
        catIdxTotals.add(i);
        return 0;
      }
      if (raw.startsWith("=") && raw.length > 1) {
        return evaluateFormula(sheet.cells, raw.slice(1));
      }
      // A trailing percent sign is stripped the same way the thousands separator
      // is: Excel copies a percent-formatted cell as its DISPLAYED text ("35%"),
      // so a pasted share table — the canonical source for a 100%/stacked chart
      // — arrives percent-suffixed. Calendar dates (Gantt timelines) become days
      // since the epoch. Both live in rawCellValue, which formulas read too.
      const { n, date } = rawCellValue(raw);
      if (!Number.isFinite(n)) return null;
      if (date) sawDate = true;
      return n;
    });

  let hundredPercent: (number | null)[] | undefined;
  let xExtent: (number | null)[] | undefined;
  const series: ChartData["series"] = [];
  // Blank rows split stacks (think-cell's clustered-stacked convention).
  let stack = 0;
  let usedStacks = false;
  rows.forEach((r) => {
    const blank = r.every((c) => c.trim() === "");
    if (blank) {
      if (series.length) {
        stack++;
        usedStacks = true;
      }
      return;
    }
    const name = (r[0] ?? "").trim();
    if (HUNDRED_ROW.test(name)) {
      hundredPercent = parseRow(r);
    } else if (XEXTENT_ROW.test(name)) {
      xExtent = parseRow(r);
    } else {
      series.push({ name: name || `Series ${series.length + 1}`, values: parseRow(r, waterfallTotals), stack });
    }
  });
  if (!usedStacks) for (const s of series) delete s.stack;
  return { categories, series, hundredPercent, xExtent, dates: sawDate || undefined };
}

/**
 * Swap the row and column of every A1 reference in a formula. A cell at (r,c)
 * lands at (c,r) after a transpose, so each of its references has to make the
 * same move — carrying "=SUM(B2:B3)" over verbatim left it summing a different
 * set of cells, silently changing the very numbers the transpose preserves.
 */
function transposeFormula(text: string): string {
  return text.replace(new RegExp(`(?<![A-Za-z0-9])${REF}`, "g"), (m, letters: string, digits: string) => {
    // Left verbatim when the reference is outside the shared address space (see
    // LAST_INDEX): the parser does not accept it either, so moving it would turn
    // one unreadable reference into a different unreadable reference. Inside the
    // space the swap is total — a transposed reference always parses, and
    // transposing twice gives the text back.
    const at = refCell(letters, digits);
    return at ? `${colLetters(at.row)}${at.col + 1}` : m;
  });
}

/** Swap rows and columns (think-cell's Transpose): series become categories. */
export function transposeSheet(sheet: SheetModel): SheetModel {
  const rows = sheet.cells.length;
  const cols = Math.max(0, ...sheet.cells.map((r) => r.length));
  const cells = Array.from({ length: cols }, (_, c) =>
    Array.from({ length: rows }, (_, r) => {
      const v = sheet.cells[r][c] ?? "";
      return v.startsWith("=") ? transposeFormula(v) : v;
    }),
  );
  return { cells };
}

/**
 * A row/column edit that MOVES data, reported so the caller can carry its own
 * per-series and per-category state along.
 *
 * `state.seriesColors` / `state.seriesMeta` in the pane are positional, and the
 * grid buttons spliced rows in and out without telling anyone: every series
 * below the edit picked up its neighbour's colour and combo type, so a routine
 * datasheet edit silently changed what the chart drew. "reset" is a transpose,
 * after which no positional mapping survives.
 */
export type SheetStructureChange =
  | { kind: "series-insert" | "series-remove" | "category-insert" | "category-remove"; index: number }
  | { kind: "reset" };

/**
 * How many SERIES rows precede sheet row `row` — i.e. the series index a row
 * edit at that row lands on. Blank separator rows and the "100%=" / "X extent"
 * rows are not series, so a raw row number is not a series number.
 */
export function seriesIndexOfRow(cells: string[][], row: number): number {
  let n = 0;
  for (let r = 1; r < Math.min(row, cells.length); r++) {
    const cs = cells[r] ?? [];
    if (cs.every((c) => (c ?? "").trim() === "")) continue;
    const name = (cs[0] ?? "").trim();
    if (HUNDRED_ROW.test(name) || XEXTENT_ROW.test(name)) continue;
    n++;
  }
  return n;
}

/** Whether sheet row `row` is itself a series row (see seriesIndexOfRow). */
function isSeriesRow(cells: string[][], row: number): boolean {
  const cs = cells[row];
  if (!cs || row < 1) return false;
  if (cs.every((c) => (c ?? "").trim() === "")) return false;
  const name = (cs[0] ?? "").trim();
  return !HUNDRED_ROW.test(name) && !XEXTENT_ROW.test(name);
}

/**
 * Editable datasheet grid: an HTML table of inputs with Excel-style TSV paste.
 * Calls onChange with the raw sheet on every edit.
 */
export function mountDatasheet(
  host: HTMLElement,
  sheet: SheetModel,
  onChange: (sheet: SheetModel) => void,
  onStructure: (change: SheetStructureChange) => void = () => {},
  /**
   * Something the sheet needs to SAY, not just do.
   *
   * Added for the European paste: rewriting "1.234" to "1234" in the user's grid
   * is a judgement about their data, and a judgement made in silence is the
   * thing this fix exists to stop. Optional so no existing caller changes, and
   * the pane passes its own `note`.
   *
   * `message` is an EN catalogue KEY carrying {placeholders}, filled from
   * `params` after translation — the same shape every other pane note uses. The
   * count must not be interpolated here: baking it in would make "2 cells" and
   * "3 cells" two different keys, and neither would ever be translated.
   */
  onNote: (message: string, params?: Record<string, string | number>) => void = () => {},
): { setSheet(next: SheetModel): void } {
  let model = sheet;
  /** Last focused cell, so row/column operations act at the cursor. */
  let cursor = { row: 1, col: 1 };

  function focusCell(row: number, col: number) {
    const el = host.querySelector<HTMLInputElement>(`input[data-row="${row}"][data-col="${col}"]`);
    if (el) {
      el.focus();
      el.select();
    }
  }

  /** Excel-style navigation: Enter ↓, arrows move at the text boundaries. */
  function handleNav(e: KeyboardEvent, ri: number, ci: number, input: HTMLInputElement) {
    const atStart = input.selectionStart === 0 && input.selectionEnd === 0;
    const atEnd = input.selectionStart === input.value.length && input.selectionEnd === input.value.length;
    const move = (r: number, c: number) => {
      e.preventDefault();
      focusCell(r, c);
    };
    if (e.key === "Enter" || e.key === "ArrowDown") {
      if (ri < model.cells.length - 1) move(ri + 1, ci);
    } else if (e.key === "ArrowUp") {
      if (ri > 0) move(ri - 1, ci);
    } else if (e.key === "ArrowRight" && atEnd) {
      if (ci < model.cells[0].length - 1) move(ri, ci + 1);
    } else if (e.key === "ArrowLeft" && atStart) {
      if (ci > 0) move(ri, ci - 1);
    }
  }

  function render() {
    host.innerHTML = "";
    const table = document.createElement("table");
    table.className = "datasheet";
    model.cells.forEach((row, ri) => {
      const tr = document.createElement("tr");
      row.forEach((cell, ci) => {
        const td = document.createElement("td");
        const input = document.createElement("input");
        input.value = cell;
        input.dataset.row = String(ri);
        input.dataset.col = String(ci);
        if (ri === 0 || ci === 0) input.classList.add("header");
        if (ri === 0 && ci === 0) input.disabled = true;
        /**
         * EVERY CELL SAYS WHERE IT IS, because a screen reader has nothing else
         * to go on. These are bare `<input>`s in bare `<td>`s — no label, no
         * `<th>`, no caption — so the pane's primary data-entry surface
         * announced "edit, blank" twenty times over, and a keyboard user had no
         * way to tell the 2023 column from the 2024 one.
         *
         * Named from the sheet's OWN headers, which is what a sighted user reads
         * off the first row and column: "Enterprise, 2024". The header cells name
         * themselves by axis instead — their own text is already the label, and
         * repeating it would announce "2024, 2024".
         *
         * Recomputed on every render, which is why it is here and not set once:
         * renaming a series has to rename the cells under it.
         */
        const rowName = model.cells[ri]?.[0]?.trim();
        const colName = model.cells[0]?.[ci]?.trim();
        input.setAttribute(
          "aria-label",
          ri === 0 && ci === 0
            ? "Corner cell, not editable"
            : ri === 0
              ? `Column ${ci} heading`
              : ci === 0
                ? `Row ${ri} heading`
                : `${rowName || `Row ${ri}`}, ${colName || `column ${ci}`}`,
        );
        input.addEventListener("input", () => {
          model.cells[ri][ci] = input.value;
          onChange(model);
        });
        input.addEventListener("focus", () => {
          cursor = { row: ri, col: ci };
        });
        input.addEventListener("keydown", (e) => handleNav(e, ri, ci, input));
        input.addEventListener("paste", (e) => handlePaste(e, ri, ci));
        td.appendChild(input);
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    host.appendChild(table);

    const controls = document.createElement("div");
    controls.className = "sheet-controls";
    // Row/column operations act at the cursor (insert after / delete at).
    controls.append(
      button("+ Row", () => {
        const at = Math.min(cursor.row + 1, model.cells.length);
        // Seed the name cell: a FULLY blank row is sheetToData's stack separator,
        // so "+ Row" used to split a stacked chart into two stacks before the
        // user had typed anything. A placeholder name (unique, so it cannot
        // collide with an existing series) reads as a new empty series instead.
        const used = new Set(model.cells.map((r) => (r[0] ?? "").trim()));
        let n = at;
        while (used.has(`Series ${n}`)) n++;
        const si = seriesIndexOfRow(model.cells, at);
        model.cells.splice(
          at,
          0,
          model.cells[0].map((_, i) => (i === 0 ? `Series ${n}` : "")),
        );
        onStructure({ kind: "series-insert", index: si });
        render();
        onChange(model);
      }),
      button("+ Column", () => {
        const at = Math.min(cursor.col + 1, model.cells[0].length);
        model.cells.forEach((r) => r.splice(at, 0, ""));
        onStructure({ kind: "category-insert", index: at - 1 });
        render();
        onChange(model);
      }),
      button("− Row", () => {
        if (model.cells.length > 2 && cursor.row > 0) {
          const at = Math.min(cursor.row, model.cells.length - 1);
          const si = isSeriesRow(model.cells, at) ? seriesIndexOfRow(model.cells, at) : null;
          model.cells.splice(at, 1);
          if (si != null) onStructure({ kind: "series-remove", index: si });
          render();
          onChange(model);
        }
      }),
      button("− Column", () => {
        if (model.cells[0].length > 2 && cursor.col > 0) {
          const at = Math.min(cursor.col, model.cells[0].length - 1);
          model.cells.forEach((r) => r.splice(at, 1));
          onStructure({ kind: "category-remove", index: at - 1 });
          render();
          onChange(model);
        }
      }),
      button("⇄ Transpose", () => {
        model = transposeSheet(model);
        // Series become categories: nothing positional survives.
        onStructure({ kind: "reset" });
        render();
        onChange(model);
      }),
    );
    host.appendChild(controls);
  }

  function handlePaste(e: ClipboardEvent, ri: number, ci: number) {
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (!text.includes("\t") && !text.includes("\n")) return; // single cell — default behavior
    e.preventDefault();
    // Quoted TSV, and exactly one trailing terminator — see `parseClipboardGrid`.
    // Interior blanks are kept, because in a single-column copy a blank cell IS
    // an empty line and dropping it pulls every later value onto the wrong
    // category. (A multi-column blank row arrives as "\t\t", which is why that
    // only ever bit single-column pastes.)
    const grid = parseClipboardGrid(text);
    /**
     * DECIDED ONCE, FOR THE WHOLE BLOCK, because that is the only place the
     * evidence exists. A single cell cannot say whether "1.234" is a thousand
     * or a fraction; a paste that also carries "987,5" can. See `looksEuropean`.
     */
    const european = looksEuropean(grid);
    let converted = 0;
    const grewFrom = model.cells.length;
    grid.forEach((row, dr) => {
      row.forEach((val, dc) => {
        const r = ri + dr;
        const c = ci + dc;
        while (model.cells.length <= r) model.cells.push(model.cells[0].map(() => ""));
        model.cells.forEach((mr) => {
          while (mr.length <= c) mr.push("");
        });
        // Rewritten into the GRID, so the reading is visible and correctable —
        // see `fromEuropeanNumber`. A cell that is not a European number in the
        // first place comes back identical, so `converted` counts only what
        // actually moved.
        const stored = european ? fromEuropeanNumber(val) : val;
        if (stored !== val) converted++;
        model.cells[r][c] = stored;
      });
    });
    // A row this paste APPENDED and left FULLY blank is `sheetToData`'s stack
    // separator, so pasting a column that runs past the end of the sheet with a
    // gap in it — an ordinary Excel spacer row — split one stacked column into
    // two side-by-side sub-stacks. No error, no cue. "+ Row" is guarded against
    // creating exactly this row and says so in its own comment; paste grows
    // rows too and never was.
    //
    // Seeded rather than dropped: the blank row is what keeps every later value
    // on the category it was pasted against, which is a separate bug this file
    // already fixed once. A row with a name and no numbers is an empty series,
    // which is what the user pasted.
    const used = new Set(model.cells.map((r) => (r[0] ?? "").trim()));
    for (let r = grewFrom; r < model.cells.length; r++) {
      if (model.cells[r].some((v) => v.trim() !== "")) continue;
      let n = r;
      while (used.has(`Series ${n}`)) n++;
      used.add(`Series ${n}`);
      model.cells[r][0] = `Series ${n}`;
    }
    render();
    onChange(model);
    // SAID, not just done. The pane read the user's numbers in a convention it
    // inferred, and rewrote their cells to match — that is a judgement about
    // their data and it has to be visible. The count is what makes it checkable:
    // it names how many cells moved, and they are all on screen.
    if (converted)
      onNote(
        "Read as European numbers — 1.234 as 1234, 987,5 as 987.5. Rewrote {n} cell(s); edit any that are wrong.",
        {
          n: converted,
        },
      );
  }

  function button(label: string, onClick: () => void): HTMLButtonElement {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.addEventListener("click", onClick);
    return b;
  }

  render();
  return {
    setSheet(next: SheetModel) {
      model = next;
      render();
    },
  };
}
