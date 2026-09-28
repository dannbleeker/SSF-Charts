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
 * The code is taken apart before it is read, because the parts that are not the
 * pattern can hold anything: `[$-409]` and `[Red]` are locale and colour
 * sections, `"y"` inside quotes is a literal the user typed, and `\m` and `_m`
 * are escapes. Only once those are gone does a y/m/d token mean what it looks
 * like — otherwise `"Day"` or `[$-409]0.00` would read as dates.
 */
export function isDateFormat(code: unknown): boolean {
  const raw = typeof code === "string" ? code : "";
  if (!raw || raw === "General") return false;
  const bare = raw
    .split(/"[^"]*"/)
    .join("")
    .split(/\[[^\]]*\]/)
    .join("")
    .replace(/[\\_].?/g, "");
  // `m` is MINUTES when it sits beside h or s; any y or d settles it outright,
  // and a run of m with no time parts around it is a month.
  return /[yd]/i.test(bare) || (/m/i.test(bare) && !/[hs]/i.test(bare));
}

/**
 * An Excel serial number as an ISO date.
 *
 * The epoch is 1899-12-30, not 1900-01-01: Excel deliberately reproduces Lotus
 * 1-2-3's belief that 1900 was a leap year, so serial 60 is a day that never
 * existed and every later serial sits one further out than the arithmetic
 * suggests. Anchoring two days early is what absorbs that.
 *
 * UTC throughout — a local-time Date would shift the day across a timezone and
 * hand back the wrong date for half the world, invisibly to whoever wrote it.
 */
export function excelSerialToISO(serial: number): string {
  return new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000)).toISOString().slice(0, 10);
}
