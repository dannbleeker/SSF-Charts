import { describe, expect, it } from "vitest";
import { excelSerialToISO, isDateFormat } from "../src/excel/serial-dates";

/**
 * A DATE IN EXCEL IS A NUMBER, AND THE BRIDGE WAS STRINGIFYING IT.
 *
 * Found 2026-09-28. `src/excel/excel.ts` loaded `values` and did `String(v)`
 * over every cell. Excel returns a date cell's value as its SERIAL NUMBER, so
 * 2026-01-05 arrived as "46027" — and nothing downstream can tell a serial
 * number from a quantity, because 46027 is a perfectly ordinary revenue
 * figure. A Gantt built from a real schedule column came out with a numeric
 * axis.
 *
 * It pairs with the `normalizeData` fix of the same day: ISO strings in a
 * Gantt Start/End row are parsed and the data is marked as a calendar. Emitting
 * ISO here is what lets that chain work end to end.
 *
 * THE DECISION IS MADE FROM THE FORMAT CODE, never the value — there is no
 * other way to tell. And the code must be taken apart first: `[$-409]` and
 * `[Red]` are sections, a quoted `"y"` is a literal the user typed, `\m` is an
 * escape. Reading those as pattern is how `[$-409]0.00` becomes a date.
 */
describe("telling an Excel date from a number", () => {
  it.each([
    ["yyyy-mm-dd", true],
    ["d/m/yyyy", true],
    ["mmm-yy", true],
    ["m/d/yy h:mm", true],
    ["[$-409]d mmmm yyyy", true],
    ["mm", true],
  ])("reads %s as a date: %s", (code, expected) => {
    expect(isDateFormat(code)).toBe(expected);
  });

  it.each([
    ["General", false],
    ["0.00", false],
    ["$#,##0.00", false],
    ["0.0%", false],
    ["#,##0;[Red]-#,##0", false],
    // A quoted literal is text the user typed, not a pattern.
    ['0.0" days"', false],
    ['#,##0" m"', false],
    // Times carry `m` as MINUTES, so `m` alone must not decide it.
    ["h:mm:ss", false],
  ])("does not read %s as a date", (code, expected) => {
    expect(isDateFormat(code), `"${code}" was read as a date format`).toBe(expected);
  });

  it("refuses anything that is not a format string", () => {
    expect(isDateFormat(undefined)).toBe(false);
    expect(isDateFormat(null)).toBe(false);
    expect(isDateFormat(46027)).toBe(false);
    expect(isDateFormat("")).toBe(false);
  });

  /**
   * The epoch is 1899-12-30, NOT 1900-01-01. Excel deliberately reproduces
   * Lotus 1-2-3's belief that 1900 was a leap year, so serial 60 is a day that
   * never existed and everything after it sits one further out than the
   * arithmetic suggests. Anchoring two days early absorbs it.
   */
  it("converts a serial number to the date Excel shows", () => {
    expect(excelSerialToISO(46027)).toBe("2026-01-05");
    expect(excelSerialToISO(45658)).toBe("2025-01-01");
    // 61 is 1900-03-01, the first serial PAST the phantom 29 February, and the
    // point from which the 1899-12-30 anchor is exact.
    expect(excelSerialToISO(61)).toBe("1900-03-01");
  });

  /**
   * NO SINGLE OFFSET IS RIGHT ON BOTH SIDES OF THE PHANTOM DAY, and pretending
   * otherwise is how this gets "fixed" wrongly later.
   *
   * Excel numbers 1900-02-29, a day that did not exist. So serials 1..59 are
   * one day ahead of the arithmetic and 61 onwards are exact, against a single
   * 1899-12-30 anchor. Serial 1 comes back 1899-12-31 here where Excel shows
   * 1900-01-01 — a real one-day discrepancy, confined entirely to the first
   * two months of 1900.
   *
   * Left alone deliberately. Every date a chart is built from is far past
   * serial 61, and special-casing the phantom window would add a branch whose
   * only exercise is a test. Written down so the next reader knows it was
   * measured rather than missed.
   */
  it("is one day out below serial 61, which is the phantom-day window and is accepted", () => {
    expect(excelSerialToISO(1)).toBe("1899-12-31");
    expect(excelSerialToISO(59)).toBe("1900-02-27");
  });

  it("uses UTC, so the day does not move with the reader's timezone", () => {
    // A local-time Date would hand back the previous or next day west or east
    // of the meridian — wrong for half the world, and invisible to whoever
    // wrote it.
    expect(excelSerialToISO(45658)).toBe(new Date(Date.UTC(2025, 0, 1)).toISOString().slice(0, 10));
  });

  it("rounds a serial that carries a time of day to its date", () => {
    // 45658.75 is 18:00 on that day. A Gantt row wants the day, not the hour.
    expect(excelSerialToISO(45658.25)).toBe("2025-01-01");
  });
});
