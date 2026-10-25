import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { CHART_KINDS } from "../src/core/samples";
import { DEFAULT_STYLE } from "../src/core/style";

/**
 * Keeps the Agent Skill honest: whenever the feature set grows, the skill's
 * reference must grow with it — otherwise CI fails and the zip can't ship
 * stale. (The skill zip itself is rebuilt from these sources by CI on every
 * push, so passing docs ⇒ current skill.)
 */
describe("skill documentation coverage", () => {
  const reference = readFileSync("skill/reference.md", "utf8");
  const skillMd = readFileSync("skill/SKILL.md", "utf8");

  for (const { kind } of CHART_KINDS) {
    it(`reference.md documents chart kind "${kind}"`, () => {
      expect(reference).toContain(`"${kind}"`);
    });
  }

  it("SKILL.md names every chart kind at least once", () => {
    for (const { kind } of CHART_KINDS) {
      expect(skillMd.includes(kind)).toBe(true);
    }
  });

  it("reference.md documents the combo marker series", () => {
    // The bare word "marker" is already all over this file (forecast markers,
    // milestone markers, the boxplot mean marker), so assert the quoted
    // Series.type value — the thing that actually has to be documented.
    expect(reference).toContain('"marker"');
  });

  it("reference.md's style schema names every ink field, and no dead one", () => {
    // The schema line drifted both ways: it promised `fontFamily` (no renderer
    // reads ChartStyle.fontFamily — the deck font is fixed) while omitting the
    // five ink fields that DO change the output, so an author could not
    // discover them. Derived from DEFAULT_STYLE so a new field must be documented.
    const schema = /style\?: \{[^}]*\}/.exec(reference)?.[0] ?? "";
    for (const key of Object.keys(DEFAULT_STYLE)) {
      if (key === "fontFamily") continue;
      expect(schema).toContain(key);
    }
    expect(schema).not.toContain("fontFamily");
  });

  it("reference.md documents the special datasheet rows", () => {
    for (const row of [
      "100%=",
      "X extent",
      "Start",
      "After",
      "Today",
      "Holiday",
      "Bracket",
      "Trend",
      "X line",
      "Color",
      "Min",
      "Q1",
      "Median",
      "Q3",
      "Max",
      "Mean",
      "Outlier",
      "Error",
      "Target",
      "% Complete",
      "Baseline start",
      "Band low",
      "Band high",
      "Column <label>",
    ]) {
      expect(reference).toContain(row);
    }
  });

  it("reference.md documents the decoration keys", () => {
    for (const key of [
      "cagr",
      "difference",
      "valueLines",
      "labelContent",
      "segmentOrder",
      "axisBreak",
      "categorySort",
      "secondaryAxis",
      "labelOffsets",
      "connectors",
      "callouts",
      "bands",
      "hundredPercentNote",
      "footnote",
      "explode",
      "colors",
      "pattern",
      "scenario",
      "datamarks",
      "tickMode",
      "tightLabelPriority",
      "gridShape",
      "fillOpacity",
      "whiskers",
      "quartileMethod",
      "iqrMultiplier",
      "jitter",
      "notch",
      "map",
      "negativeColor",
      "forecastFrom",
      "quadrants",
      "columns",
      "barStyle",
      "fillBetween",
      "totals",
      "grandTotal",
      "variance",
      "slope",
      "breakout",
      "multiples",
      "stepped",
      "gapWidth",
      "overlap",
      "trajectory",
      "smooth",
      "spacerIndices",
      "summaryBars",
      "radarBand",
      "otherBucket",
      "calendar",
      "bridgeGaps",
      "perSpoke",
      "lineAxes",
      "tilemap",
      "glyph",
      "semi",
      "pareto",
      "bump",
      "criticalPath",
      "meanSd",
      "sparkline",
      "workdays",
      "marginals",
      "spreadLimit",
      "detailGroups",
      "lanes",
      "markers",
      "trendDegree",
      "symbols",
      "variableRadius",
      "bars",
      "stacked",
      "sizeEncode",
      "cluster",
      // The field, not the word: "grouping" also turns up in two unrelated sentences.
      "grouping?: boolean",
    ]) {
      expect(reference).toContain(key);
    }
  });
});
