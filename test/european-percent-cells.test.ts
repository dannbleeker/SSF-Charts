import { describe, expect, it } from "vitest";
import { fromEuropeanNumber, looksEuropean } from "../src/taskpane/datasheet";

/**
 * A EUROPEAN SHARE TABLE IS STILL EUROPEAN.
 *
 * Found 2026-09-28. `US_GROUPED`, `EU_GROUPED` and `EU_DECIMAL` are all
 * anchored, so `"12,5%"` matched none of them — and a share table is the single
 * most likely thing anyone pastes into this grid. Measured: a sheet whose every
 * number carried a percent sign returned `looksEuropean === false`, where the
 * identical sheet without the signs returned true, and
 * `fromEuropeanNumber("12,5%")` handed back `"12,5%"` untouched. The block was
 * then read as American and the decimal comma never converted.
 *
 * `rawCellValue` in the same file already strips a trailing "%" before parsing,
 * so the two places that care about percent cells disagreed about whether they
 * exist. Same shape as the CHAR(11)/CHAR(10) lesson: a rule written for one
 * spelling of a cell silently skips the others.
 */
describe("percent cells in a European paste", () => {
  const sheet = (share: string[]) => [["Region", "Share"], ...share.map((s, i) => [`R${i}`, s])];

  it("counts a comma decimal as European even wearing a percent sign", () => {
    expect(
      looksEuropean(sheet(["12,5%", "37,5%", "50,0%"])),
      "a sheet of European percentages read as American, so every decimal comma survived into the grid",
    ).toBe(true);
    // The control the defect was measured against.
    expect(looksEuropean(sheet(["12,5", "37,5", "50,0"]))).toBe(true);
  });

  it("still refuses to guess when the block contradicts itself", () => {
    // Percent signs must not tip a mixed block — the function's whole contract
    // is that it answers only when one side is unopposed.
    expect(looksEuropean(sheet(["12,5%", "1,234.5%"]))).toBe(false);
  });

  it("does not read an American percentage as European", () => {
    expect(looksEuropean(sheet(["12.5%", "37.5%"]))).toBe(false);
    expect(looksEuropean(sheet(["1,234%", "5,678%"]))).toBe(false);
  });

  it("converts the number and keeps the sign", () => {
    expect(fromEuropeanNumber("12,5%")).toBe("12.5%");
    expect(fromEuropeanNumber("1.234,5%")).toBe("1234.5%");
    // Unchanged behaviour for the bare forms, or the fix moved something else.
    expect(fromEuropeanNumber("12,5")).toBe("12.5");
    expect(fromEuropeanNumber("1.234,5")).toBe("1234.5");
    // And a cell that is not European is returned verbatim, sign and all.
    expect(fromEuropeanNumber("12.5%")).toBe("12.5%");
    expect(fromEuropeanNumber("Region")).toBe("Region");
  });

  it("leaves a percent sign separated by a space alone in the number", () => {
    expect(fromEuropeanNumber("12,5 %")).toBe("12.5 %");
  });
});
