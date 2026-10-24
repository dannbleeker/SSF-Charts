import { buildChart } from "../../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../../src/core/samples";
import type { ChartConfig } from "../../src/core/types";

/**
 * WHAT A SWEEP'S OPTION TABLE CANNOT SAY BY ITSELF: THE DATA IT NEEDS.
 *
 * `frame-fit.test.ts` and `overlap-budget.test.ts` each hold an `OPTIONS` table
 * and lay one entry over `sampleConfig(kind)` for every kind with a shallow
 * spread. That makes a table entry only as good as the sample it lands on, and
 * six of the entries meant nothing on theirs: `tilemap.glyph` draws glyphs only
 * from a SECOND series and the tilemap sample has one; `otherBucket: {max: 3}`
 * folds series and no sample has more than three; `pie.variableRadius` takes its
 * radius from series 1; `butterfly.split: 1` is the layout the sample already
 * has; and the `boxplot` options (notch, jitter, the mean) all need raw
 * observations where the sample carries a five-number summary. Each was swept
 * with the option on and the scene byte-identical to the one with it off — the
 * `tilemap.glyph` entry was green over 32 of 32 cells that drew no glyph, and
 * would have stayed green over glyphs thrown a chart-width off the frame.
 *
 * Nothing said so. A sweep entry that changes nothing is indistinguishable from
 * one that passes, which is the failure `docs/BACKLOG.md` names for every gate
 * here: a gate is only as wide as its sweep. So this module does two things.
 * `withOptionData` gives an option the data it needs, for the kind it is about
 * and for no other; `optionsThatChangeNothing` is the check that the table's
 * entries are doing something, for whichever file owns the table.
 */

/** `n` series from the sample's own, the extra ones named "Series i" and shifted so they differ. */
function withSeries(c: Record<string, any>, n: number): Record<string, any> {
  const own = c.data.series.length;
  return {
    ...c,
    data: {
      ...c.data,
      series: Array.from({ length: n }, (_, i) => ({
        ...c.data.series[i % own],
        name: i < own ? c.data.series[i].name : `Series ${i + 1}`,
        values: c.data.series[i % own].values.map((v: number | null, k: number) =>
          v == null || i < own ? v : v + i * 3 + k,
        ),
      })),
    },
  };
}

/**
 * Raw observations for the boxplot: one row per OBSERVATION, one column per
 * category, which is what turns the five-number-summary rows of the sample into
 * something a notch, a jitter overlay or a mean can be computed from. The last
 * row is an outlier on purpose, so the whisker and outlier paths run too.
 */
function withObservations(c: Record<string, any>): Record<string, any> {
  return {
    ...c,
    data: {
      ...c.data,
      series: Array.from({ length: 9 }, (_, i) => ({
        name: `Run ${i + 1}`,
        values: c.data.categories.map((_: string, k: number) => 4 + k + ((i * 7 + k * 3) % 9) + (i === 8 ? 12 : 0)),
      })),
    },
  };
}

/** The kinds `otherBucket` folds; see `OTHER_KINDS` in `src/core/chart.ts`. */
const COLUMN_FAMILY = ["stacked", "clustered", "stacked100"];

/**
 * How to give each option the data it needs. A function of the whole config that
 * returns it unchanged for every kind the option is not about, because the table
 * is applied to all of them.
 */
const OPTION_DATA: Record<string, (c: Record<string, any>) => Record<string, any>> = {
  "tilemap.glyph": (c) => (c.kind === "tilemap" ? withSeries(c, 3) : c),
  // `max: 3` folds anything past three series into "Other".
  otherBucket: (c) => (COLUMN_FAMILY.includes(c.kind) ? withSeries(c, 6) : c),
  "pie.variableRadius": (c) => (c.kind === "pie" || c.kind === "doughnut" ? withSeries(c, 2) : c),
  // `split: 1` puts series 0 on the left and the REST on the right, stacked.
  "butterfly.split": (c) => (c.kind === "butterfly" ? withSeries(c, 3) : c),
  boxplot: (c) => (c.kind === "boxplot" ? withObservations(c) : c),
};

/** `cfg` with whatever data the sweep option `key` needs to take effect; `cfg` itself if it needs none. */
export function withOptionData<T extends object>(key: string, cfg: T): T {
  return (OPTION_DATA[key]?.(cfg as Record<string, any>) ?? cfg) as T;
}

/**
 * Options that are scene-neutral BY NATURE, each with the reason, so the check
 * below does not report them. `render: "image"` chooses how the PowerPoint
 * writer places the chart, a decision made after `buildChart` has finished, so
 * the scene a layout sweep measures is the same with it on or off.
 */
export const SCENE_NEUTRAL_OPTIONS: Record<string, string> = {
  "render image": "chooses the write path in the pane, after the scene is built",
};

/**
 * The entries of `options` that change no chart's scene, at the two frames
 * roomy enough for every option to show (a cramped frame can only hide an
 * effect, so it cannot be what makes an option look inert).
 *
 * `prepare` is how the caller gives an entry its data, so a caller that chooses
 * not to prepare one gets it back reported. A kind that refuses a config throws,
 * and that is not a difference.
 */
export function optionsThatChangeNothing(
  options: Record<string, Record<string, unknown>>,
  prepare: (key: string, cfg: Record<string, any>) => Record<string, any> = withOptionData,
): string[] {
  const scene = (cfg: unknown): string => {
    try {
      return JSON.stringify(buildChart(cfg as ChartConfig).nodes);
    } catch {
      return "refused";
    }
  };
  return Object.entries(options)
    .filter(([key, value]) => {
      for (const { kind } of CHART_KINDS)
        for (const [width, height] of [
          [480, 300],
          [960, 540],
        ]) {
          const plain = prepare(key, { ...sampleConfig(kind), width, height });
          if (scene(plain) !== scene({ ...plain, ...value })) return false;
        }
      return true;
    })
    .map(([key]) => key);
}
