import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { CHART_KINDS } from "../src/core/samples";

/**
 * Keeps the user manual honest, the same way skill-docs.test.ts keeps the
 * Agent Skill honest: when a chart kind, datasheet row, pane control, or
 * decoration ships, docs/MANUAL.md must mention it — otherwise CI fails.
 * Extend the manual (not this test's exception lists) when adding features.
 */
describe("user manual coverage", () => {
  const manual = readFileSync("docs/MANUAL.md", "utf8");
  const lower = manual.toLowerCase();

  for (const { kind, label } of CHART_KINDS) {
    it(`documents chart kind "${kind}"`, () => {
      // Kinds appear by their gallery label (the user-facing name).
      expect(lower).toContain(label.toLowerCase());
    });
  }

  it("documents every special datasheet row", () => {
    for (const row of [
      "100%=",
      "X extent",
      "X line",
      "Y line",
      "Trend",
      "Error+",
      "Error-",
      "Min",
      "Q1",
      "Median",
      "Q3",
      "Max",
      "Mean",
      "Outlier",
      "Start",
      "End",
      "Milestone",
      "After",
      "Today",
      "Holiday",
      "Bracket",
      "Column <label>",
    ]) {
      expect(manual, `datasheet row ${row}`).toContain(row);
    }
  });

  it("documents the decorations and chart options", () => {
    for (const feature of [
      "Segment labels",
      "Series labels",
      "totals",
      "Grand total",
      "IBCS scenario",
      "IBCS variance",
      "polynomial",
      "Value axis",
      "Gridlines",
      "Connector lines",
      "100% = note",
      "Datamark axis",
      "CAGR",
      "Difference arrow",
      "Value line",
      "Segment order",
      "Axis break",
      "label content",
      "Thousands separator",
      "Footnote",
      "Explode slices",
      "Same scale",
      "callouts",
      "background bands",
      "pattern fills",
      "slope",
      "breakout",
      "Small multiples",
      "Stepped line/area",
      "Gap width",
      "Area with negatives",
      "Scatter trajectory",
      "Boxplot jitter",
      "Continuous color",
      "Smoothed lines",
      "Waterfall spacers",
      "Gantt summary bars",
      "Notched boxplots",
      "Radar band",
      "Other bucket",
      "Calendar heatmap",
      "Butterfly stacked flanks",
      "Radar per-spoke scales",
      "Missing-data bridge",
      "Floating segments",
      "Combo base modes",
      "Combo independent line axes",
      "Hex tile maps",
      "Tilemap mini-glyphs",
      "100% charts with negatives",
      "Semi-circle gauge",
      "Pareto",
      "Bump chart",
      "Horizontal profile chart",
      "Critical path",
      "Mean±SD box",
      "Sparklines",
      "Radial bar",
      "Stacked radar",
      "Variable-radius pie",
      "Cell-size heatmap",
      "Heatmap clustering",
      "Combo area base",
    ]) {
      expect(manual, `option ${feature}`).toContain(feature);
    }
  });

  it("documents the pane workflows", () => {
    for (const flow of [
      "Edit selected chart",
      "Update chart",
      "Auto-update",
      "Transpose",
      "Use deck theme",
      "Export style",
      "Save style to deck",
      "Use deck style",
      "Save as template",
      "Export current",
      "Import",
      "Insert batch",
      "Horizontal (bar)",
      "Insert as picture",
      "Explode to native shapes",
      "formulas".toLowerCase(),
    ]) {
      expect(lower, `workflow ${flow}`).toContain(flow.toLowerCase());
    }
  });

  it("documents the elements and their in-cell tokens", () => {
    for (const el of [
      "Harvey ball",
      "Checkbox",
      "Process flow",
      "KPI tile",
      "Total row",
      "[hb:",
      "[up]",
      "[good]",
      "Agenda",
    ]) {
      expect(manual, `element ${el}`).toContain(el);
    }
  });

  it("documents the Excel companion and the Claude skill", () => {
    expect(manual).toContain("manifest-excel.xml");
    expect(manual).toContain("ssf-charts.zip");
  });
});
