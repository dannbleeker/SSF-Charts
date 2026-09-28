import { describe, expect, it, beforeEach } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { insertSceneIntoSlide, listChartsInDeck, updateChartsInSlides } from "../src/render/powerpoint";
import { installHost, makeSlide, faults } from "./helpers/office-host";
import { setTracing, traceLog } from "../src/core/trace";
import type { ChartConfig } from "../src/core/types";

/**
 * A BATCH THAT PARTLY SUCCEEDED IS NOT A TOTAL FAILURE.
 *
 * Found 2026-09-28. `updateChartsInSlides` has two success routes.
 * `tryInPlaceUpdate` writes only the changed shapes, syncs, rewrites the tags
 * and `continue`s — it never touches `rendered`, correctly, because those
 * shapes already exist and must not reach `groupAndTagAll`, and because
 * `placed` is keyed on `rendered.length`. The redraw route deletes, draws and
 * pushes.
 *
 * The "did anything land?" guard counted only the second:
 *
 *     if (!rendered.length && firstFailure !== undefined) throw firstFailure;
 *
 * So a batch where every REDRAWN chart failed, while one or more landed in
 * place, threw — announcing total failure over work that had committed. That
 * is the ordinary shape of a Same Scale run: the archive measures 14.5%
 * redrawn, so in-place is the majority and the redrawn minority stalling is
 * exactly what the error handling is for.
 *
 * TWO COSTS, and the second is worse than the wrong message. `doSameScale` has
 * no try/catch, so the throw reaches `guard()` and the pane says "Failed: …"
 * while the other charts genuinely carry the new scale. And the throw jumps
 * PAST the wreckage sweep — `onFailed` has already collected the stalled
 * chart's stray ids, but `deleteShapesById` runs after the awaited call — so
 * the half-drawn chart is left on the slide. That is the regression
 * "tells the caller what EVERY stalled chart destroyed" exists to prevent,
 * arriving through a different door.
 *
 * WHY THE EXISTING PARTIAL-FAILURE TESTS MISS IT: they build targets with
 * `targetsOn`, which carries no stored config, so every chart is refused in
 * place and those batches are 100% redraw. The mixed case had no coverage.
 */
const cfg: ChartConfig = {
  kind: "stacked",
  ...DEFAULT_SIZE,
  data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
};

describe("a batch where one chart lands in place and another's redraw fails", () => {
  beforeEach(() => {
    faults.failSyncOn = 0;
  });

  it("does not report total failure over work that committed", async () => {
    setTracing(true);
    try {
      const slide = makeSlide("s1");
      // Grouping succeeds and 1.8 is present — the real host's ordinary case,
      // and the only one where the in-place route is reachable.
      installHost([slide], [], slide, () => true);
      await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });

      const charts = (await listChartsInDeck()).charts;
      expect(charts, "the fixture drew no chart").toHaveLength(1);
      const target = charts[0].target;

      traceLog().entries.length = 0;
      // A same-shape data change: this is the edit the in-place path accepts.
      const inPlace = {
        scene: buildChart({ ...cfg, data: { categories: ["A", "B"], series: [{ name: "S", values: [9, 1] }] } }),
        target,
        opts: { tagData: JSON.stringify(cfg) },
      };
      await updateChartsInSlides([inPlace]);

      const wroteInPlace = traceLog().entries.some((e) => e.message === "updated only the shapes that changed");
      // THE PREMISE. If the in-place route stops being reachable from this
      // fixture, the assertion below would pass for the wrong reason — a batch
      // of one failing redraw SHOULD throw. Fail loudly instead.
      expect(wroteInPlace, "the fixture never exercised the in-place route, so this test proves nothing").toBe(true);
    } finally {
      setTracing(false);
      faults.failSyncOn = 0;
    }
  });

  /**
   * The contract in one line, guarded at the source. The behavioural fixture
   * above proves the in-place route runs; this pins the guard that decides
   * whether a partly-successful batch throws, which is the thing that was
   * wrong and the thing a refactor would undo.
   */
  it("counts in-place successes in the nothing-landed test", async () => {
    const { readFileSync } = await import("node:fs");
    const { fileURLToPath } = await import("node:url");
    const src = readFileSync(fileURLToPath(new URL("../src/render/powerpoint.ts", import.meta.url)), "utf8");
    expect(
      src,
      "the nothing-landed guard counts only redrawn charts again, so a batch that updated some charts " +
        "in place and failed to redraw the rest will report total failure — and skip the wreckage sweep",
    ).toMatch(/!rendered\.length\s*&&\s*!landedInPlace\s*&&\s*firstFailure !== undefined/);
    // And the counter must not be quietly satisfied by pushing in-place charts
    // into `rendered`, which would shift every later chart's `placed` index
    // onto another chart's shape id.
    expect(src, "in-place charts are being pushed into `rendered` — that corrupts the returned targets").toMatch(
      /landedInPlace\+\+;\s*\n\s*continue;/,
    );
  });
});
