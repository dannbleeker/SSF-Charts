/**
 * Excel's date representation, and how to tell it from a quantity.
 *
 * SEPARATE FROM `excel.ts` because that module wires the pane on import — it
 * reaches for `document` at load — so nothing there can be unit-tested without
 * a DOM. These two functions are pure and carry the whole decision, which is
 * the part worth testing: the rest of the bridge is plumbing.
 */

/**
 * IS THIS NUMBER FORMAT A DATE?
 *
 * Decided from the FORMAT CODE, never from the value. Excel returns a date
 * cell's `values` entry as a serial number — 2026-01-05 arrives as 46027 — and
 * a serial number is indistinguishable from a quantity by inspection. 46027 is
 * a perfectly ordinary revenue figure.
 *
 * The code is taken apart before it is read (`formatPattern`), because the
 * parts that are not the pattern can hold anything. Only once those are gone
 * does a y/m/d token mean what it looks like — otherwise `"Day"` or
 * `[$-409]0.00` would read as dates.
 */
export function isDateFormat(code: unknown): boolean {
  const bare = formatPattern(code);
  // `m` is MINUTES when it sits beside h or s; any y or d settles it outright,
  // and a run of m with no time parts around it is a month.
  return /[yd]/i.test(bare) || (/m/i.test(bare) && !/[hs]/i.test(bare));
}

/**
 * A FORMAT CODE WITH EVERYTHING THAT IS NOT PATTERN TAKEN OUT, leaving the
 * tokens Excel reads: `[$-409]` and `[Red]` are locale and colour sections,
 * `"y"` inside quotes is a literal the user typed, and `\m` and `_m` are
 * escapes. "" for a code that is not one, or is General.
 *
 * ELAPSED-TIME TOKENS ARE BRACKETED TOO, and they are the one bracket that is
 * pattern. `[h]`, `[m]` and `[s]` (and their doubled forms) run past 24 hours,
 * 60 minutes and 60 seconds, which is how a timesheet is written. Stripping
 * them with the sections took `[h]:mm` down to `:mm` — an `m` with no `h` or
 * `s` beside it, which `isDateFormat` reads as a MONTH — so every duration
 * arrived as a date in 1899 and charted as -25,568. The only question the `m`
 * rule asks of a code is whether a time part is present, and any of the three
 * is one, so each becomes an `h`. `[mm]:ss` and a lone `[m]` came out right
 * before by luck of what was left over; this makes them right on purpose.
 */
export function formatPattern(code: unknown): string {
  const raw = typeof code === "string" ? code : "";
  if (!raw || raw === "General") return "";
  return raw
    .split(/"[^"]*"/)
    .join("")
    .replace(/\[(?:h+|m+|s+)\]/gi, "h")
    .split(/\[[^\]]*\]/)
    .join("")
    .replace(/[\\_].?/g, "");
}

/**
 * An Excel serial number as an ISO date.
 *
 * The epoch is 1899-12-30, not 1900-01-01: Excel deliberately reproduces Lotus
 * 1-2-3's belief that 1900 was a leap year, so serial 60 is a day that never
 * existed and every later serial sits one further out than the arithmetic
 * suggests. Anchoring two days early is what absorbs that.
 *
 * A WORKBOOK ON THE 1904 DATE SYSTEM counts from 1904-01-01 instead (Excel's
 * "Use 1904 date system", and the default of Mac Excel before 2011). There is
 * no phantom day to absorb, so the anchor is the date itself, and the same day
 * is 1,462 serials smaller: 2026-01-05 is 46027 in one and 44565 in the other.
 * `values` carries only the number, so reading a 1904 serial against the 1900
 * anchor put every date exactly four years and a day early, with nothing on
 * screen to say so.
 *
 * UTC throughout — a local-time Date would shift the day across a timezone and
 * hand back the wrong date for half the world, invisibly to whoever wrote it.
 */
export function excelSerialToISO(serial: number, date1904 = false): string {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return new Date(epoch + Math.round(serial * 86400000)).toISOString().slice(0, 10);
}

/**
 * DOES THE WORKBOOK COUNT FROM 1904, given what it answered for DATE(2000,1,1)?
 *
 * Nothing in a cell says which system it was stored in, so the workbook is
 * asked to do the one thing that depends on it: build a date. DATE(2000,1,1) is
 * 36526 on the 1900 system and 35064 on the 1904 one. Decoding the answer as a
 * 1904 serial lands on 2000-01-01 only if that is what it is, so no table of
 * magic numbers is kept here; any other answer (a host that returned an error
 * or nothing) reads as the 1900 system, which is what the bridge assumed
 * before it could ask.
 */
export function isDate1904(serialOf2000: unknown): boolean {
  return (
    typeof serialOf2000 === "number" &&
    Number.isFinite(serialOf2000) &&
    excelSerialToISO(serialOf2000, true) === "2000-01-01"
  );
}
