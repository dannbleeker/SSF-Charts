import { it, expect } from "vitest";
import { writeFileSync } from "fs";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { estimateOfficeShapes } from "../src/core/scene";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";

it("TMPSHAPES", () => {
  const rows = CHART_KINDS.map((k) => {
    const cfg = { ...sampleConfig(k.kind), ...DEFAULT_SIZE };
    const n = estimateOfficeShapes(buildChart(cfg));
    return { kind: k.kind, label: k.label, shapes: n, overWebBudget: n > 105 };
  }).sort((a, b) => b.shapes - a.shapes);
  writeFileSync("C:/Users/dann_/AppData/Local/Temp/claude/shapes.json", JSON.stringify(rows, null, 1));
  expect(rows.length).toBe(25);
});
