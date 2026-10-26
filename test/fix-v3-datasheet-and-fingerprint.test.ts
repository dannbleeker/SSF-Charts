// @vitest-environment jsdom
import { afterAll, afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { setTracing } from "../src/core/trace";
import type { ChartConfig } from "../src/core/types";
import { fromEuropeanNumber, looksEuropean, sheetToData, typedNumber } from "../src/taskpane/datasheet";

/**
 * THE THIRD BUG HUNT, STREAM v3: two leftovers the m2 review named.
 *
 * V3a: a thousands group cannot start with a zero. M2e (9d59b76) fixed
 * `EU_GROUPED` and `US_GROUPED`; every OTHER place the datasheet reads a thousands
 * group kept the head `\d{1,3}` or `\d{1,2}`, and BACKLOG.md called them
 * unreachable. They are reachable by any text a cell or a box can hold.
 *
 * V3b (the skill stamping a scene fingerprint) was implemented in 774205f and is HELD BACK:
 * its tests went with it. See docs/BACKLOG.md.
 */

afterEach(() => {
  vi.unstubAllGlobals();
  setTracing(false);
});

/** See `pane-state.test.ts`: the pane's last debounced render must not outlive jsdom. */
afterAll(async () => {
  await new Promise((r) => setTimeout(r, 250));
});

/** What the datasheet makes of one typed or pasted cell: the number, or null for a gap. */
const cell = (text: string): number | null =>
  sheetToData({
    cells: [
      ["", "A"],
      ["S", text],
    ],
  }).series[0].values[0];

/** A pasted block with one column of data under a header, the way `looksEuropean` is handed one. */
const block = (...cells: string[]) => [["Region", "N"], ...cells.map((c) => ["x", c])];

describe("V3a: the space and apostrophe group in a datasheet cell cannot open with a zero", () => {
  // Every spelling of the group gap `GROUP_GAP` names: plain, no-break, narrow
  // no-break (fr-FR), thin space, the ASCII apostrophe and U+2019 (de-CH, fr-CH).
  const GAPS = [" ", " ", " ", " ", "'", "’"];

  it("reads 0 500 as no number, not five hundred, whichever mark groups it", () => {
    for (const gap of GAPS) {
      expect(cell(`0${gap}500`), JSON.stringify(`0${gap}500`)).toBeNull();
      expect(cell(`00${gap}500`), "two zeros").toBeNull();
      expect(cell(`007${gap}500`), "a zero-led three-digit head").toBeNull();
      expect(cell(`0${gap}500${gap}000`), "two groups").toBeNull();
      expect(cell(`0${gap}000`), "a zero is not written as a group of zeros").toBeNull();
    }
  });

  it("refuses it with its dressing on too: a sign, parentheses, a currency, a percent", () => {
    expect(cell("-0 500")).toBeNull();
    expect(cell("(0 500)")).toBeNull();
    expect(cell("€ 0 500")).toBeNull();
    expect(cell("0 500 kr.")).toBeNull();
    expect(cell("0 500%")).toBeNull();
    expect(cell("0 500-")).toBeNull();
  });

  it("refuses a zero-led space group that carries a decimal, in either mark", () => {
    expect(cell("0 500.5")).toBeNull();
    expect(cell("0 500,5")).toBeNull();
  });

  it("still reads every group that opens with 1 to 9", () => {
    for (const gap of GAPS) {
      expect(cell(`1${gap}500`), `1${gap}500`).toBe(1500);
      expect(cell(`10${gap}500`)).toBe(10500);
      expect(cell(`999${gap}999`)).toBe(999999);
      expect(cell(`1${gap}234${gap}567`)).toBe(1234567);
      expect(cell(`-1${gap}500`)).toBe(-1500);
      expect(cell(`(1${gap}500)`)).toBe(-1500);
      expect(cell(`€ 1${gap}500`)).toBe(1500);
      expect(cell(`1${gap}500.5`)).toBe(1500.5);
    }
  });

  it("leaves a plain zero, a zero decimal and an unspaced leading zero as they were", () => {
    expect(cell("0")).toBe(0);
    expect(cell("0.5")).toBe(0.5);
    expect(cell("-0.5")).toBe(-0.5);
    expect(cell("0.500")).toBe(0.5);
    expect(cell(".5")).toBe(0.5);
    // No separator, so no group: a leading zero here is only a number written with a
    // zero in front, which `Number` has always read.
    expect(cell("0500")).toBe(500);
    expect(cell("007")).toBe(7);
  });

  it("casts no European vote and is not converted: the rest of a Danish block still is", () => {
    // `0 500,5` read as the core `0500,5`, which `EU_DECIMAL` took for a comma
    // decimal: it voted European and `fromEuropeanNumber` handed back 0500.5.
    expect(fromEuropeanNumber("0 500,5")).toBe("0 500,5");
    expect(fromEuropeanNumber("0'500,5")).toBe("0'500,5");
    expect(looksEuropean(block("0 500,5"))).toBe(false);
    // And a real European paste around it is unaffected.
    expect(looksEuropean(block("0 500,5", "1.234,5"))).toBe(true);
    expect(fromEuropeanNumber("1 234,5")).toBe("1234.5");
  });
});

describe("V3a: Indian grouping cannot open with a zero", () => {
  it("reads 0,12,345 as no number, not twelve thousand", () => {
    expect(cell("0,12,345")).toBeNull();
    expect(cell("01,23,456")).toBeNull();
    expect(cell("0,00,500")).toBeNull();
    expect(cell("00,12,345.5")).toBeNull();
    expect(cell("-0,12,345")).toBeNull();
    expect(cell("(0,12,345)")).toBeNull();
  });

  it("is no evidence for an American sheet, so a Danish block keeps its conversion", () => {
    // It matched `LAKH_GROUPED`, which `looksEuropean` counts as a comma that is a
    // thousands mark, and one American cell takes the conversion from the whole block.
    expect(looksEuropean(block("0,12,345", "1.234,5"))).toBe(true);
    expect(looksEuropean(block("01,23,456", "987,5"))).toBe(true);
  });

  it("still reads every Indian grouping that opens with 1 to 9", () => {
    expect(cell("1,23,456")).toBe(123456);
    expect(cell("12,34,567.89")).toBe(1234567.89);
    expect(cell("1,00,000")).toBe(100000);
    expect(cell("99,99,999")).toBe(9999999);
    expect(cell("-1,23,456")).toBe(-123456);
    expect(cell("+12,34,567")).toBe(1234567);
    expect(cell("(1,23,456)")).toBe(-123456);
    expect(looksEuropean(block("1,23,456", "987.5"))).toBe(false);
    expect(looksEuropean(block("12,34,567.89", "1.234,5"))).toBe(false);
  });

  it("leaves the US and the European groups as M2e left them", () => {
    expect(cell("1,500")).toBe(1500);
    expect(cell("12,345.6")).toBe(12345.6);
    expect(cell("0,500")).toBeNull();
    expect(fromEuropeanNumber("1.500")).toBe("1500");
    expect(fromEuropeanNumber("10.500")).toBe("10500");
    expect(fromEuropeanNumber("1.234.567")).toBe("1234567");
    expect(fromEuropeanNumber("0.500")).toBe("0.500");
  });
});

describe("V3a: typedNumber reads a zero-led group as a decimal or as nothing, never as a group", () => {
  const LOCALES = ["en-US", "da-DK", "de-DE", "fr-FR"];

  it("is NaN for a repeated mark that opens with a zero", () => {
    for (const locale of LOCALES) {
      for (const text of ["0.500.500", "0,500,500", "00.500.500", "01,234,567", "-0.500.500", "+0,500,500"]) {
        expect(typedNumber(text, locale), `${text} in ${locale}`).toBeNaN();
      }
    }
  });

  it("is NaN for both marks with a zero-led group: 0.500,5 and 0,500.5", () => {
    for (const locale of LOCALES) {
      for (const text of ["0.500,5", "0,500.5", "00.500,25", "0.500.500,5", "-0,500.5"]) {
        expect(typedNumber(text, locale), `${text} in ${locale}`).toBeNaN();
      }
    }
  });

  it("is NaN for a zero-led space or apostrophe group, the box's copy of the cell rule", () => {
    for (const locale of LOCALES) {
      for (const gap of [" ", " ", " ", " ", "'", "’"]) {
        expect(typedNumber(`0${gap}500`, locale), `0${gap}500 in ${locale}`).toBeNaN();
        expect(typedNumber(`00${gap}500`, locale)).toBeNaN();
        expect(typedNumber(`0${gap}500,5`, locale)).toBeNaN();
        expect(typedNumber(`0${gap}500.5`, locale)).toBeNaN();
        expect(typedNumber(`-0${gap}500`, locale)).toBeNaN();
      }
    }
  });

  it("still reads every group that opens with 1 to 9, in every locale", () => {
    for (const locale of LOCALES) {
      expect(typedNumber("1.234.567", locale), locale).toBe(1234567);
      expect(typedNumber("1,234,567", locale), locale).toBe(1234567);
      expect(typedNumber("10.500.250", locale)).toBe(10500250);
      expect(typedNumber("999,999,999", locale)).toBe(999999999);
      expect(typedNumber("1.234,5", locale)).toBe(1234.5);
      expect(typedNumber("1,234.5", locale)).toBe(1234.5);
      expect(typedNumber("1.234.567,89", locale)).toBe(1234567.89);
      expect(typedNumber("1,234,567.89", locale)).toBe(1234567.89);
      expect(typedNumber("-1.234.567", locale)).toBe(-1234567);
      expect(typedNumber("1 500", locale)).toBe(1500);
      expect(typedNumber("1'500", locale)).toBe(1500);
      expect(typedNumber("1’500", locale)).toBe(1500);
      expect(typedNumber("1 234 567", locale)).toBe(1234567);
      expect(typedNumber("-1 500", locale)).toBe(-1500);
    }
    expect(typedNumber("1 234,5", "da-DK")).toBe(1234.5);
    expect(typedNumber("1 234.5", "en-US")).toBe(1234.5);
  });

  it("keeps the single-mark decimals and the leading-zero decimals as they were", () => {
    for (const locale of LOCALES) {
      expect(typedNumber("0.500", locale), `0.500 in ${locale}`).toBe(0.5);
      expect(typedNumber("-0.5", locale)).toBe(-0.5);
      expect(typedNumber("0,5", locale)).toBe(0.5);
      expect(typedNumber("-0,5", locale)).toBe(-0.5);
      expect(typedNumber("0", locale)).toBe(0);
      expect(typedNumber("0.0", locale)).toBe(0);
      expect(typedNumber("10.500", "da-DK")).toBe(10500);
      expect(typedNumber("10.500", "en-US")).toBe(10.5);
    }
    // The ones the locale decides, as k4 pinned them.
    expect(typedNumber("1,500", "da-DK")).toBe(1.5);
    expect(typedNumber("1,500", "en-US")).toBe(1500);
    expect(typedNumber("3.000", "da-DK")).toBe(3000);
    expect(typedNumber("3.000", "en-US")).toBe(3);
  });
});

describe("V3a: the pane's own boxes refuse them, which is how a person can reach typedNumber", () => {
  const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const typeInto = (el: HTMLInputElement, value: string) => {
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const boxes = (prefix: string): HTMLInputElement[] => {
    const row = [...document.querySelectorAll<HTMLElement>("#options label")].find((l) =>
      (l.textContent ?? "").startsWith(prefix),
    );
    if (!row) throw new Error(`no options row starting "${prefix}"`);
    return [...row.querySelectorAll<HTMLInputElement>('input[type="text"]')];
  };

  async function bootPane() {
    try {
      window.localStorage.clear();
    } catch {
      /* a jsdom without storage */
    }
    window.history.replaceState({}, "", "/taskpane.html");
    const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
    parsed.querySelectorAll("script").forEach((s) => s.remove());
    document.body.innerHTML = parsed.body.innerHTML;
    vi.resetModules();
    const app = await import("../src/taskpane/app");
    await app.harnessReady;
  }

  it("Axis scale min and max: 0 500 and 0.500.500 set nothing, 1 500 and 1.500.500 set a number", async () => {
    await bootPane();
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({
      kind: "clustered",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
      numberFormat: { locale: "da-DK" },
    } satisfies Partial<ChartConfig>);
    $("json-import").click();
    const exported = (): ChartConfig => {
      $("json-export").click();
      return JSON.parse(($("json-io") as HTMLTextAreaElement).value);
    };
    const [min, max] = boxes("Axis scale min");
    typeInto(min, "0 500");
    typeInto(max, "0.500.500");
    expect(exported().scale, "a zero-led group set an axis bound").toBeUndefined();
    typeInto(min, "1 500");
    typeInto(max, "1.500.500");
    expect(exported().scale).toEqual({ min: 1500, max: 1500500 });
  });
});

/* -------------------------------------------------------------------------- */
