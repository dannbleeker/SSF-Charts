import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";

import type { ChartConfig, ChartKind } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

/**
 * A BLANK CELL MUST NOT READ AS A MEASURED ZERO.
 *
 * The worst thing this engine can do is not to fail — it is to state something
 * false about the user's business. A missing bar reads as missing. A printed
 * "0" reads as measured, and a percentage computed from it reads as measured
 * twice. Funnel, waffle and cascade all did this until 2026-08-28: they shared
 * `Math.max(0, values[c] ?? 0)` and formatted their LABELS from the clamped
 * array, so one empty spreadsheet cell — the most ordinary input there is —
 * printed "0", and cascade went on to compute a 100% drop from it.
 *
 * Those three were fixed one at a time, by hand, from a list. This is the
 * ratchet that makes the list unnecessary: it asks every kind the engine ships,
 * and a kind that starts confusing blank for zero fails here whether or not
 * anybody thought to look at it.
 *
 * ── WHAT IT ASKS, AND THE VERSION THAT WAS WRONG ────────────────────────────
 * The first cut compared a blank run against a zero run and called them being
 * IDENTICAL the defect. That over-reports badly, and measuring it said so: it
 * named eight kinds, and most were innocent. `clustered` drops the label for a
 * blank category and drops it for a zero-height bar too — identical, and
 * correct both times. `pie` omits the slice and RECOMPUTES the other three
 * percentages off a denominator that excludes it, which is exactly right. Only
 * `stacked` was actually doing the bad thing, and a test reporting seven false
 * positives beside it is one nobody will keep believing.
 *
 * The harm is narrower than "cannot tell them apart". It is: **a label that
 * asserts a NUMBER about a category we have no data for.** So each kind is
 * built twice, changing one category's values only — every series `null`, then
 * every series `7` — and the labels named for that category are compared:
 *
 *   the 7 run must label it     — otherwise the kind is not-applicable, and
 *                                 without this the check passes on kinds that
 *                                 label nothing, which is how a ratchet quietly
 *                                 stops ratcheting
 *   the blank run must not      — a label there may be absent, or say "n/a", or
 *     print a DIGIT for it        say anything without a number in it. What it
 *                                 may not do is print a figure nobody measured.
 *
 * Labels are matched by name suffix, which is this engine's own convention:
 * `label-0-3`, `total-3` and `label-3` all speak about category 3, and
 * `label-0` speaks about category 0 and is none of this test's business.
 */

/** One category's worth of values replaced, everything else left alone. */
function withCategory(cfg: ChartConfig, index: number, value: number | null): ChartConfig {
  return {
    ...cfg,
    ...DEFAULT_SIZE,
    data: {
      ...cfg.data,
      series: cfg.data.series.map((s) => ({
        ...s,
        values: s.values.map((v, i) => (i === index ? value : v)),
      })),
    },
  };
}

/**
 * The labels this chart draws whose NAME says they speak about category `at`.
 *
 * `-${at}` as a suffix is the engine's own convention — `label-0-3`, `total-3`
 * and `col-total-3` all speak about category 3. `row-` is excluded because it
 * indexes the OTHER axis: this sweep blanks a category (a column), and
 * `row-total-3` is the fourth SERIES' total, which legitimately changes when a
 * column empties and legitimately still has a figure from the columns that
 * remain. Counting it reported `heatmap: row-total-3="55"` as a fabricated
 * number when 55 was the honest sum of the cells left.
 */
function labelsFor(cfg: ChartConfig, at: number): { name: string; text: string }[] {
  return buildChart(cfg)
    .nodes.filter((n): n is TextNode => n.kind === "text")
    .map((n) => ({ name: String(n.name ?? ""), text: String(n.text ?? "") }))
    .filter((n) => n.name.endsWith(`-${at}`) && !n.name.startsWith("row-"));
}

type Verdict = {
  kind: ChartKind;
  state: "asserts-a-number" | "says-nothing-false" | "does-not-label" | "refused";
  saying?: string;
};

function judge(kind: ChartKind): Verdict {
  let base: ChartConfig;
  try {
    /**
     * THE LABELS TURNED ON, not whichever ones the sample happens to use.
     *
     * The sweep read each kind's sample verbatim at first, and a sample with
     * totals switched off has no total to get wrong — so the check passed the
     * kind without ever asking it the question. `heatmap` came back
     * "distinguishes" that way while `sum()` in its own file counts a null as
     * zero and a blank row prints `row-total = "0"`; it was distinguishing on
     * its CELL labels, and the marginals were never built.
     *
     * This is the same blind spot the archive already names — no ratchet varies
     * `decorations` — narrowed here to the decorations that make a chart SPEAK.
     */
    const sample = sampleConfig(kind) as ChartConfig;
    base = {
      ...sample,
      ...DEFAULT_SIZE,
      decorations: { ...(sample.decorations ?? {}), totals: true, segmentLabels: true },
      /**
       * AND THE TOTALS THAT ARE NOT `decorations.totals`.
       *
       * Turning the shared decoration on was not enough: the heatmap's
       * marginals are `heatmap: { totals }`, so the sweep built a heatmap with
       * no marginals and reported it clean while `sum()` in its own file
       * counted a null as zero — a blank row printed `row-total = "0"` and a
       * blank column `col-total = "0"`. Per-kind switches have to be named
       * one at a time; that is the cost of the engine having them, and a kind
       * whose labels hide behind an option nobody sets here is a kind this
       * ratchet is not asking.
       */
      ...(kind === "heatmap" ? { heatmap: { ...(sample.heatmap ?? {}), totals: "both" as const } } : {}),
    };
  } catch {
    return { kind, state: "refused" };
  }
  /**
   * The last category that actually READS its cell.
   *
   * Last, because several kinds treat the first specially — a cascade's opening
   * bar, a waterfall's base — and a defect hiding behind that exception would
   * be reported as the exception.
   *
   * But NOT a declared total. `waterfall: { totalIndices: [5] }` says index 5's
   * bar is computed from the deltas before it, so blanking that cell removes
   * nothing: the sample's FY24 total is 86+14+9-12-4 = 93 whatever is in the
   * cell, and the first version of this test reported that 93 as a fabricated
   * figure. A derived bar has no "no data" state to get wrong.
   */
  const totals = new Set<number>(
    ((base as { waterfall?: { totalIndices?: number[] } }).waterfall?.totalIndices ?? []).map(Number),
  );
  let at = (base.data?.categories?.length ?? 0) - 1;
  while (at >= 0 && totals.has(at)) at--;
  if (at < 1 || !base.data.series.length) return { kind, state: "refused" };
  try {
    /**
     * THE CATEGORY'S OWN NAME IS NOT AN ASSERTION ABOUT ITS VALUE.
     *
     * Half the first sharpened run's report was `category-3="2025"`,
     * `col-3="Q4"`, `category-5="FY24"` — axis headings that carry digits
     * because businesses name their categories after years and quarters. An
     * axis that still says "2025" over an empty column is telling the truth,
     * and a chart that dropped the heading would be worse. Matched on the TEXT
     * rather than by excluding `category-`/`col-`/`row-` by name, because the
     * principle is "it is repeating the label, not reporting a figure" and that
     * survives a renaming the name list would not.
     */
    const heading = String(base.data.categories?.[at] ?? "");
    const speaks = (l: { text: string }) => /\d/.test(l.text) && (!heading || !l.text.includes(heading));
    const labelled = labelsFor(withCategory(base, at, 7), at).filter(speaks);
    if (!labelled.length) return { kind, state: "does-not-label" };
    const asserting = labelsFor(withCategory(base, at, null), at).filter(speaks);
    return asserting.length
      ? { kind, state: "asserts-a-number", saying: asserting.map((l) => `${l.name}="${l.text}"`).join(", ") }
      : { kind, state: "says-nothing-false" };
  } catch {
    return { kind, state: "refused" };
  }
}

describe("a blank cell is not a measured zero", () => {
  const verdicts = CHART_KINDS.map((k) => judge(k.kind));

  it("asks a question of enough kinds to be worth having", () => {
    // The guard on the guard. If sample data or a refusal quietly took most
    // kinds out of the sweep, everything below would pass on a handful — which
    // is how a ratchet stops ratcheting without anyone noticing.
    const asked = verdicts.filter((v) => v.state !== "refused" && v.state !== "does-not-label");
    expect(asked.length, `only ${asked.length} of ${verdicts.length} kinds could be asked`).toBeGreaterThan(8);
  });

  it("a heatmap row with no data at all gets no total either", () => {
    /**
     * The OTHER axis, which the sweep above cannot reach: it blanks a category,
     * so it can only ever exercise `col-total`. `row-total` has the identical
     * defect from the identical cause — `sum()` counting a null as zero — and
     * would have gone on printing "0" for an empty row with the column half
     * fixed and green beside it.
     */
    const sample = sampleConfig("heatmap") as ChartConfig;
    const cfg = {
      ...sample,
      ...DEFAULT_SIZE,
      heatmap: { ...(sample.heatmap ?? {}), totals: "both" as const },
      data: {
        ...sample.data,
        series: sample.data.series.map((s, i) => (i === 0 ? { ...s, values: s.values.map(() => null) } : s)),
      },
    } as ChartConfig;
    const totals = buildChart(cfg)
      .nodes.filter((n): n is TextNode => n.kind === "text" && String(n.name ?? "").startsWith("row-total-"))
      .map((n) => ({ name: String(n.name), text: String(n.text) }));
    // ABSENT OR EMPTY, either is right. Withholding the figure leaves a text
    // node with no text, and the scene prunes those, so the node is simply not
    // there — asserting on the empty string would have pinned the mechanism
    // instead of the harm, and broken the day pruning changed.
    const emptied = totals.find((t) => t.name === "row-total-0");
    expect(emptied?.text ?? "", "an empty row was given a measured total").toBe("");
    // …and the rows that DO have data still say so, or the fix has just
    // deleted the feature.
    expect(totals.filter((t) => /\d/.test(t.text)).length).toBeGreaterThan(1);
  });

  it("no kind prints a number for a category it has no data for", () => {
    const guilty = verdicts.filter((v) => v.state === "asserts-a-number").map((v) => `${v.kind}: ${v.saying}`);
    expect(
      guilty,
      "a blank cell is being printed as a measured figure — the slide asserts something the data does not say",
    ).toEqual([]);
  });

  /**
   * THE SITE THIS RATCHET COULD NOT REACH, found 2026-09-28.
   *
   * The sweep above asks every KIND, via `sampleConfig(kind)`. Clustered-
   * stacked is not a kind — it is a config variant, reached by giving series a
   * `stack` id — so no sample carries one and `column.ts`'s `nStacks > 1`
   * branch was never entered. And `labelsFor` matches by category suffix
   * (`-3`), which the sub-stack's `total-3-s0` does not end with. Missed twice
   * over, for two independent reasons.
   *
   * What that hid: the sub-stack branch sits ABOVE the `else if` carrying the
   * `columnHasData` gate, whose own comment says it "gates BOTH branches below
   * ... so a third one cannot be added past it". This one was added above it.
   * A wholly blank category printed `total-1-s0 = "0"` and `total-1-s1 = "0"`,
   * while the single-stack control over the same blank column correctly
   * printed no total at all.
   */
  const stackedCfg = (a: (number | null)[], b: (number | null)[], c: (number | null)[], d: (number | null)[]) =>
    ({
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Prod A", stack: 0, values: a },
          { name: "Prod B", stack: 0, values: b },
          { name: "Serv A", stack: 1, values: c },
          { name: "Serv B", stack: 1, values: d },
        ],
      },
      decorations: { totals: true },
    }) as unknown as ChartConfig;

  const subTotals = (cfg: ChartConfig): Record<string, string> => {
    const out: Record<string, string> = {};
    for (const n of buildChart(cfg).nodes) {
      const name = String((n as TextNode).name ?? "");
      if (n.kind === "text" && /^total-\d+-s\d+$/.test(name)) out[name] = String((n as TextNode).text ?? "");
    }
    return out;
  };

  it("a clustered-stacked sub-column with nothing in it gets no total", () => {
    const got = subTotals(stackedCfg([10, null, 30], [5, null, 15], [8, null, 12], [4, null, 6]));
    expect(Object.keys(got).length, "no sub-stack totals were drawn at all — re-point this test").toBeGreaterThan(0);
    for (const slot of ["total-1-s0", "total-1-s1"]) {
      expect(
        got[slot] ?? "",
        `Q2 is blank in every series and ${slot} says ${JSON.stringify(got[slot])}. A sum of nothing is 0, ` +
          `which is right for the stack's reach and wrong for a label: it tells the reader the business ` +
          `measured zero where nobody filled the cells in.`,
      ).toBe("");
    }
    // …and the categories that DO carry data still print, or the fix deleted
    // the feature instead of gating it.
    expect(got["total-0-s0"] ?? "", "a populated sub-stack lost its total").toMatch(/\d/);
    expect(got["total-2-s1"] ?? "", "a populated sub-stack lost its total").toMatch(/\d/);
  });

  it("gates per sub-stack, not per category", () => {
    // One stack blank in Q2, the other not. Only the blank one may go
    // unlabelled — a category-wide gate would wrongly silence both, which is
    // the over-correction this class invites.
    const got = subTotals(stackedCfg([10, 20, 30], [5, 10, 15], [8, null, 12], [4, null, 6]));
    expect(got["total-1-s0"] ?? "", "the populated stack in a half-blank category lost its total").toMatch(/\d/);
    expect(got["total-1-s1"] ?? "", "the blank stack was still labelled a measured zero").toBe("");
  });
});
