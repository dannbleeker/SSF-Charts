import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * A GANTT WRITTEN WITH ISO DATES GETS A CALENDAR AXIS.
 *
 * Found 2026-09-28. `normalizeData` turns an ISO string in a Start/End row into
 * days-since-epoch — that half was already fixed, and there is a comment above
 * it explaining why. What it never did was record that the row HAD been a
 * calendar, and the number that comes out cannot say so.
 *
 * `layoutGantt` gates everything calendar on `data.dates`: week and month
 * ticks, the tick-label format, the today marker. Only `sheetToData` ever set
 * that flag, so only a PASTE produced a calendar. The skill's own documented
 * config — ISO strings in Start/End rows, stated in SKILL.md as a hard rule —
 * parsed its dates correctly and then drew a numeric axis labelled in raw epoch
 * days. The bars land in the right relative places, which is exactly what made
 * it look deliberate.
 *
 * This asserts the OUTCOME a reader sees — that no axis label is a bare epoch
 * number — rather than the `dates` flag, so it survives the flag being renamed
 * and fails if the flag is set but the rendering still does not follow.
 */
const iso = (): ChartConfig =>
  ({
    kind: "gantt",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Design", "Build", "Test"],
      series: [
        { name: "Start", values: ["2026-01-05", "2026-02-02", "2026-03-09"] },
        { name: "End", values: ["2026-01-30", "2026-03-06", "2026-03-27"] },
      ],
    },
  }) as unknown as ChartConfig;

const numeric = (): ChartConfig =>
  ({
    kind: "gantt",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Design", "Build"],
      series: [
        { name: "Start", values: [0, 10] },
        { name: "End", values: [8, 20] },
      ],
    },
  }) as unknown as ChartConfig;

const texts = (cfg: ChartConfig) =>
  buildChart(cfg)
    .nodes.filter((n) => n.kind === "text")
    .map((n) => String((n as { text?: string }).text ?? ""))
    .filter(Boolean);

describe("a Gantt whose Start/End rows are ISO strings", () => {
  it("labels its timeline as dates, not as epoch day numbers", () => {
    const labels = texts(iso());
    /**
     * A number around 20,000 anywhere in the label, separators and ranges
     * included.
     *
     * The first version of this matched `/^\d{5}$/` and the defect walked
     * straight through it: `formatNumber` groups thousands, so the axis reads
     * "20,440" and the task bars read "20,458–20,483". The test passed with the
     * fix reverted, which is the only reason it was caught — a mutation check,
     * not a green run. Matching the DIGITS rather than a whole-string shape is
     * what makes it see both the axis ticks and the range captions.
     */
    const epochish = labels.filter((t) => /\b2[01],?\d{3}\b/.test(t));
    expect(
      epochish,
      `the timeline is labelled in raw days-since-epoch (${epochish.join(" ")}). The dates parsed, but ` +
        `nothing recorded that they WERE dates, so layoutGantt drew a numeric axis.`,
    ).toEqual([]);
  });

  it("draws a bar for every task, which is what made it look deliberate", () => {
    // The guard on the guard: if the ISO rows stopped parsing at all, the axis
    // assertion above would pass on a chart with no bars.
    const bars = buildChart(iso()).nodes.filter((n) =>
      /^task-|^bar-/.test(String((n as { name?: string }).name ?? "")),
    );
    expect(
      bars.length,
      "no task bars at all — the ISO rows are not parsing, so this test is not asking anything",
    ).toBeGreaterThan(0);
  });

  it("leaves a numeric Gantt numeric", () => {
    // The scope guard. A plan measured in days is a legitimate Gantt and must
    // not acquire a calendar because this fix exists.
    const labels = texts(numeric());
    expect(
      labels.some((t) => /\d{4}-\d{2}-\d{2}|Jan|Feb|Mar|W\d/.test(t)),
      "a numeric Gantt was given calendar labels",
    ).toBe(false);
  });
});
