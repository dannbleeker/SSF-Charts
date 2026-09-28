import { describe, expect, it } from "vitest";
import { parseClipboardGrid } from "../src/taskpane/datasheet";

/**
 * THE CLIPBOARD IS QUOTED TSV, AND THIS USED TO SPLIT IT ON NEWLINES.
 *
 * Two defects in three lines, both found 2026-09-28.
 *
 * QUOTES. Excel's `text/plain` clipboard quotes any cell holding a line break,
 * a tab or a double quote, and keeps the break inside the quotes — the same
 * convention it uses for CSV. The paste handler did
 * `text.replace(/\r/g,"").split("\n")` and then `split("\t")` per row, so a
 * wrapped header (Alt+Enter inside a cell — the most ordinary thing in a
 * finance table) was read as the end of a row. One pasted row became two and
 * every later row landed on the wrong series, silently. Same family as the
 * CHAR(11)/CHAR(10) lesson: the character Alt+Enter actually produces is the
 * one nothing was looking for.
 *
 * TRAILING BLANKS. Excel appends exactly ONE row terminator; the old code
 * popped empty lines in a `while`. In a single-column copy a blank CELL is an
 * empty line — the comment it sat under said exactly that — so copying
 * `10, 20, 30, <blank>, <blank>` wrote three cells and left the sheet's old
 * fourth and fifth values showing underneath.
 */
describe("reading a clipboard block", () => {
  it("keeps a line break that lives inside a quoted cell", () => {
    // A1 = "Revenue" + Alt+Enter + "(EURm)", then two year columns.
    const clip = '"Revenue\n(EURm)"\t2024\t2025\r\nEnterprise\t34\t40\r\nSMB\t22\t26\r\n';
    const grid = parseClipboardGrid(clip);
    expect(
      grid.length,
      `the in-cell line break was read as a row terminator, so ${grid.length} rows came out of a ` +
        `3-row paste and every later row landed on the wrong series`,
    ).toBe(3);
    expect(grid[0]).toEqual(["Revenue\n(EURm)", "2024", "2025"]);
    expect(grid[1]).toEqual(["Enterprise", "34", "40"]);
    expect(grid[2]).toEqual(["SMB", "22", "26"]);
  });

  it("keeps a tab that lives inside a quoted cell", () => {
    expect(parseClipboardGrid('"a\tb"\tc\r\n')).toEqual([["a\tb", "c"]]);
  });

  it("reads Excel's doubled quote as one literal quote", () => {
    expect(parseClipboardGrid('"say ""hi"""\tx\r\n')).toEqual([['say "hi"', "x"]]);
  });

  it("treats a quote that does not open a cell as data", () => {
    // `5" pipe` is a measurement, not a quoted cell.
    expect(parseClipboardGrid('5" pipe\t12\r\n')).toEqual([['5" pipe', "12"]]);
  });

  it("drops exactly one trailing terminator, not every trailing blank", () => {
    // Five cells down a column, the last two deliberately empty.
    const grid = parseClipboardGrid("10\r\n20\r\n30\r\n\r\n\r\n");
    expect(
      grid.length,
      `a single-column paste ending in blank cells came back ${grid.length} rows long. The blanks are ` +
        `the user's data; dropping them leaves the sheet's OLD values showing underneath.`,
    ).toBe(5);
    expect(grid).toEqual([["10"], ["20"], ["30"], [""], [""]]);
  });

  it("keeps interior blanks, which is what the old code got right", () => {
    // The regression guard for the fix that came before this one.
    expect(parseClipboardGrid("10\r\n\r\n30\r\n")).toEqual([["10"], [""], ["30"]]);
  });

  it("still reads an ordinary rectangular block", () => {
    expect(parseClipboardGrid("a\tb\r\nc\td\r\n")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
    // A multi-column blank row arrives as tabs, and must survive as one.
    expect(parseClipboardGrid("a\tb\r\n\t\r\nc\td\r\n")).toEqual([
      ["a", "b"],
      ["", ""],
      ["c", "d"],
    ]);
  });

  it("survives a block with no trailing terminator at all", () => {
    // Not every source appends one; dropping a real row would be worse than
    // keeping a blank.
    expect(parseClipboardGrid("a\tb\r\nc\td")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("handles a lone \\r as a row break, as the old code did", () => {
    expect(parseClipboardGrid("a\rb")).toEqual([["a"], ["b"]]);
  });
});
