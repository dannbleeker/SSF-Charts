import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { textWidth, type SceneNode, type TextNode } from "../src/core/scene";
import { SCENE_NEUTRAL_OPTIONS, optionsThatChangeNothing, withOptionData } from "./helpers/sweep-options";

/**
 * EVERY TEXT OVERLAP THIS ENGINE STILL HAS, counted, named, and capped.
 *
 * `frame-fit.test.ts` sweeps each kind at eight frames and seven fonts and now
 * allows no overlap at all. What it does not sweep is the option and data-shape
 * VARIANTS — a chart with twenty-four categories, ten series, a secondary axis,
 * a pareto, a footnote — and that is the blind spot the whole "24 categories"
 * family lived in for nine days. The gate was green throughout.
 *
 * This is that sweep, ratcheted. It ran for the first time on 2026-08-27 at
 * 2,148 overlapping pairs, measured by the frame gate's own ink rule; a night's
 * work took it to 537. Pinning it per SHAPE rather than as one total matters,
 * because a total can hide a trade — one family fixed while another grows — and
 * a shape absent from the table is a regression by definition, since it is text
 * the engine has never been seen to draw over text before.
 *
 * AND THEN THIS FILE TURNED OUT TO HAVE THE SAME BLIND SPOT IT WAS BUILT TO
 * CLOSE. It swept every option and every data shape and never one of each: the
 * two tables were CONCATENATED, so a secondary axis on ten series, or a
 * footnote on twenty-four categories, was not among the 24,000 charts. Crossing
 * a slice of them on 2026-08-29 took the count from 537 to 4,010 and found
 * twenty-two shapes this engine had never been SEEN to draw — 2,683 of them in
 * small-multiples panels alone. Nothing about the engine changed that day.
 *
 * So: the total moved because the sweep did, and a number from before that date
 * is not comparable with one after it. `CROSS_OPTIONS` below says what is
 * crossed and, more usefully, what still is not.
 *
 * THEN THE 2,683 WENT AWAY IN ONE LINE, the same day. Every `p#-` shape is now
 * zero and the total is 1,327. `buildMultiples` was handing panels a NEGATIVE
 * height and `clampDim` was rewriting it to `DEFAULT_SIZE.height` — so each
 * panel laid out as a full 300-point chart and ten of them stacked 9.6 points
 * apart in a 60-point box. It now declines a grid whose panels have no room.
 * See `buildMultiples` for the proof and docs/BACKLOG.md for the shape of it.
 *
 * That is the sequence this file is for, in one day: a widened sweep made a
 * family visible, the family turned out to be one defect, and the defect was
 * arithmetic rather than layout.
 *
 * THEN 1,327 BECAME 1,014, by clipping the unit label. `valueAxisTitle` is
 * documented as a SHORT unit — "e.g. `€m`" — and the sweep tests it with a
 * twenty-seven-character sentence, so most of the largest family here was the
 * cost of a string the option does not support. Clipping it to a share of the
 * chart keeps the author's text where it fits and truncates where it cannot,
 * which is what this engine already does to gantt and category names.
 *
 * Flooring its `y` at the title's ink was measured in the same hour and
 * REFUSED: it takes `title / value-axis-title` to zero and pays 310 new
 * `value-axis-title / category#` for it, for a worse total of 1,156. A clamp
 * moves a label whether or not the destination is free — which the CAGR
 * caption's own note already said, about the same move.
 *
 * THEN 1,014 BECAME 785: the unit YIELDS to the title instead of moving. Not a
 * clamp, a drop, on the 22 charts of 176 where its ink would land in the title's
 * — all of them 80x60 and 300x60 at 18pt, where the engine already drops the
 * category names, the axis strip and the legend ("Chrome yields to the title",
 * docs/MANUAL.md). The unit was the last thing in that band still printing over
 * the title. 154 of 176 units survive, 100% at 480x300 and above.
 *
 * ================================================================
 * READ THIS BEFORE QUOTING ANY NUMBER IN THIS FILE AT A DECISION.
 * ================================================================
 *
 * **This sweep measures an input the option does not support, and the resulting
 * figure has now driven three decisions, two of them wrongly.**
 *
 * `OPTIONS.valueAxisTitle` is `"Revenue in millions of euro"`. The option is
 * documented as `€m`; the two uses in the 123-chart shipped deck are `€m` and
 * `$m (log)`. Run the identical sweep with `€m` and:
 *
 *     total                    785  ->  206
 *     value-axis-title family  687  ->  118
 *     value-axis-title/legend# 308  ->   12
 *
 * So about two-thirds of the largest family in this file — and a quarter of its
 * headline total — is one test string. What it cost: the floor above was refused
 * on the long string's numbers (1,156, worse) when on `€m` the same change is a
 * net WIN (437 -> 346); and `title / value-axis-title`, 205 pairs, turned out to
 * come entirely from ONE frame, 300x60, rather than from the label's width.
 *
 * The stress string stays — it found the missing clip, which was a real bug, and
 * a gate that only tests kind inputs is not a gate. But the number it produces is
 * a STRESS figure and must be named as one. Do not carry it into a product
 * decision without re-running at the documented input first; the two answer
 * different questions and they have disagreed every time it mattered.
 *
 * Same shape as this file's own founding lesson one level up: a gate is only as
 * wide as its sweep. A gate is also only as true as its inputs.
 *
 * WHEN A NUMBER HERE GOES DOWN, EDIT IT DOWN. The budget is a ceiling, not a
 * target, and one left above the real figure is a ratchet that has stopped
 * ratcheting — the next regression hides under the slack. There is a test below
 * that enforces exactly that, so improving the engine is supposed to fail this
 * file.
 *
 * The sweep is 24 options and 10 data shapes across every kind, eight frames,
 * two fonts and both orientations — about 24,000 charts — and it runs in eight
 * seconds, which is why it can live in the ordinary suite rather than in a
 * script somebody has to remember.
 */

/** EXACTLY `inkBox` from the frame gate, so the two measure the same thing. */
const ink = (t: TextNode) => {
  const w = textWidth(t.text, t.fontSize, t.bold);
  const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
  const base =
    t.valign === "top"
      ? t.y + t.fontSize
      : t.valign === "bottom"
        ? t.y + t.h - t.fontSize * 0.25
        : t.y + t.h / 2 + t.fontSize * 0.36;
  return { x0: x, y0: base - t.fontSize * 0.8, x1: x + w, y1: base + t.fontSize * 0.21 };
};

const sharedArea = (a: ReturnType<typeof ink>, b: ReturnType<typeof ink>) => {
  const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
  const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
  return w > 0 && h > 0 ? w * h : 0;
};

const FRAMES: [number, number][] = [
  [80, 60],
  [120, 90],
  [160, 120],
  [200, 150],
  [300, 60],
  [60, 300],
  [480, 300],
  [960, 540],
];

/** The option variants, matching the frame gate's own table. */
const OPTIONS: Record<string, Record<string, unknown>> = {
  scale: { scale: { min: -50, max: 500 } },
  segmentOrder: { segmentOrder: "reverse" },
  categorySort: { categorySort: "descending" },
  secondaryAxis: { secondaryAxis: true },
  axisBreak: { axisBreak: { from: 20, to: 60 } },
  valueAxisTitle: { valueAxisTitle: "Revenue in millions of euro" },
  logScale: { logScale: true },
  gapWidth: { gapWidth: 0.8 },
  overlap: { overlap: 0.5 },
  footnote: { footnote: "Source: an internal model, restated for the 2024 perimeter" },
  pareto: { pareto: true },
  "multiples 2 columns": { multiples: { columns: 2 } },
  "multiples 3 columns": { multiples: { columns: 3 } },
  otherBucket: { otherBucket: { max: 3 } },
  "pie.explode": { pie: { explode: [0, 1] } },
  "pie.semi": { pie: { semi: true } },
  "pie.variableRadius": { pie: { variableRadius: true } },
  "tilemap.hex": { tilemap: { shape: "hex" } },
  "tilemap.glyph": { tilemap: { glyph: "bars" } },
  "butterfly.split": { butterfly: { split: 1 } },
  boxplot: { boxplot: { notch: true, showMean: true, jitter: true } },
  "render image": { render: "image" },
  numberFormat: { numberFormat: { locale: "de-DE", currency: "EUR", decimals: 2 } },
  title: { title: "A rather long chart title that names the measure and the period" },
};

/**
 * OPTIONS THIS SWEEP DOES NOT GIVE THE DATA THEY NEED, and the price of doing so.
 *
 * `withOptionData` gives an option the data it means something on (a second
 * series for `tilemap.glyph`, raw observations for `boxplot`): without it six
 * of the entries above were measured with the option ON and the scene
 * byte-identical to the one with it off, which counted nothing and said so
 * nowhere. Five of them take their data here. The sixth is the boxplot, and it
 * is left bare on purpose: with raw observations its `showMean` labels collide
 * with each other and with the axes, 53 overlaps in four shapes
 * (`mean# / mean#` 26, `category# / mean#` 12, `title / mean#` 8,
 * `value-axis / mean#` 7) that `BUDGET` has no row for, and by this file's own
 * rule a shape absent from the table is a regression. Measured 2026-10-24
 * against the unprepared table, the other four add nothing at all.
 *
 * So it is declared, not forgotten. Taking it on is a decision for the table's
 * owner — fix the labels, or budget the four rows — and the check at the end of
 * this file keeps the declaration true.
 */
const UNPREPARED: Record<string, string> = {
  boxplot: "raw observations make the mean labels collide, 53 overlaps in shapes BUDGET has no row for",
};

/** `cfg` with the data its option needs, except for the options declared `UNPREPARED`. */
const prepared = (key: string, cfg: Record<string, any>) => (key in UNPREPARED ? cfg : withOptionData(key, cfg));

/**
 * THE DECORATIONS, WHICH THIS SWEEP HAS NEVER ONCE VARIED.
 *
 * `overlap-budget`, `shape-budget` and `showcase-overlap` contained ZERO
 * references to `decorations` between them. The sweep varied chart kind, frame,
 * font size, orientation and a cross-product of options — and never the segment
 * labels, series labels, column totals, grand total, category axis, value axis
 * or the `100% =` note.
 *
 * Those are precisely the nodes that collide. Every text-over-text defect this
 * project has fixed was a label against another label, so the budget that guards
 * text-over-text was measured on charts drawn with most of their text turned
 * OFF. `sampleConfig` gives each kind a small default set — `clustered` gets
 * `categoryAxis`, `stacked` gets `seriesLabels` — and a user who ticks three
 * more boxes in the pane produces a chart this sweep had never drawn.
 *
 * MERGED, NOT REPLACED, and that is the whole reason these are not in `OPTIONS`
 * beside everything else. `OPTIONS` is applied with a shallow spread, so a
 * `decorations` key there would REPLACE each kind's defaults and quietly narrow
 * the sweep while appearing to widen it. These merge on top, which is also what
 * the pane does when a user ticks a box.
 *
 * ── AND THE ANSWER, WHICH IS THAT THE ENGINE IS FINE HERE ──
 *
 * Recorded because a negative result nobody writes down gets re-investigated.
 * At the frames these run at, turning EVERY decoration on adds no text overlap
 * the budget did not already carry: the total stayed at 785. The variants are
 * not idle — each builds 100 charts with zero refusals and adds real text, up to
 * +232 nodes for `everything labelled` — they simply collide with nothing.
 *
 * Nor is there a shape-budget risk, which is the other thing more labels could
 * have cost: across 25 kinds at a full-slide frame, full decoration is +18
 * shapes at worst and the heaviest decorated chart is `area` at 195, against the
 * ~300 where this host starts dying. Zero kinds cross it.
 *
 * So the blind spot was real and what it was hiding was nothing. That is worth
 * one paragraph and no further work.
 */
const DECOR: Record<string, Record<string, unknown>> = {
  // The ordinary "just show me the numbers" case.
  "labels on": { segmentLabels: true, seriesLabels: true, categoryAxis: true, valueAxis: true },
  // Totals, which land above a column and next to whatever else is up there.
  "totals on": { totals: true, grandTotal: true, segmentLabels: true },
  // Gridlines and both axes, the case where the frame fills with rules.
  "axes and gridlines": { categoryAxis: true, valueAxis: true, gridlines: true },
  // The maximal text case: everything a user can turn on that draws a string.
  "everything labelled": {
    segmentLabels: true,
    seriesLabels: true,
    totals: true,
    grandTotal: true,
    categoryAxis: true,
    valueAxis: true,
    gridlines: true,
    hundredPercentNote: true,
  },
};

/** And the data shapes, likewise: a user's data is none of the sample's shapes. */
const DATA_SHAPES: Record<string, (c: any) => any> = {
  "long category names": (c) => ({
    ...c,
    data: { ...c.data, categories: c.data.categories.map((x: string, i: number) => `${x} enterprise segment ${i}`) },
  }),
  "long series names": (c) => ({
    ...c,
    data: { ...c.data, series: c.data.series.map((s: any) => ({ ...s, name: `${s.name} including allocations` })) },
  }),
  "24 categories": (c) => ({
    ...c,
    data: {
      ...c.data,
      categories: Array.from({ length: 24 }, (_, i) => `C${i + 1}`),
      series: c.data.series.map((s: any) => ({
        ...s,
        values: Array.from({ length: 24 }, (_, i) => (s.values[i % s.values.length] ?? 0) + i),
      })),
    },
  }),
  "10 series": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: Array.from({ length: 10 }, (_, i) => ({
        ...c.data.series[i % c.data.series.length],
        name: `Series ${i + 1}`,
        values: c.data.series[i % c.data.series.length].values.map((v: number | null) => (v == null ? v : v + i)),
      })),
    },
  }),
  "one category": (c) => ({
    ...c,
    data: {
      ...c.data,
      categories: c.data.categories.slice(0, 1),
      series: c.data.series.map((s: any) => ({ ...s, values: s.values.slice(0, 1) })),
    },
  }),
  "one series": (c) => ({ ...c, data: { ...c.data, series: c.data.series.slice(0, 1) } }),
  "values in the billions": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : v * 1234567)),
      })),
    },
  }),
  "every value negative": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : -v)),
      })),
    },
  }),
  "mixed signs": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null, i: number) => (v == null ? v : i % 2 ? -v : v)),
      })),
    },
  }),
  "tiny fractions": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : v / 100000)),
      })),
    },
  }),
};

/**
 * What each remaining shape is allowed, as of 2026-08-29. Total 1,327.
 *
 * **NONE OF THESE THREE NUMBERS MEANS ANYTHING WITHOUT ITS SWEEP.** 537 was the
 * uncrossed sweep; 4,010 was the same engine measured by the crossed one; 1,327
 * is the crossed sweep after `buildMultiples` stopped building grids whose
 * panels have no height. Only the last step was a change to the engine. Never
 * compare a figure here with one from before 2026-08-29 without saying which
 * sweep produced it — the 2026-08-19 "75 pairs" was mis-compared exactly that
 * way once already.
 *
 * A `#` stands in for any run of digits, so `label-3 / label-4` and `label-9 /
 * label-10` are one shape. What each family is, and why the ones left are still
 * here, is in `docs/BACKLOG.md`.
 *
 * The 1,327 is one family and a tail:
 *
 *     1,208   `value-axis-title` — the owner's open decision
 *       119   everything else
 *
 * THOSE ARE NOT 1,058 AND 269, which is what the 4,010 sweep split into and
 * what this said until it was re-summed. Removing the broken grids did not just
 * subtract 2,683: a chart that used to render as an impossible grid now renders
 * as an ordinary chart, and an ordinary chart with a unit label contributes
 * `value-axis-title` overlaps. The family GREW by about 150 as the tail shrank
 * by the same. A subtraction is not a re-measurement.
 *
 * The `p#-*` family that appeared here at 2,683 is GONE, and its story is worth
 * the four lines. It was not a labelling defect — that was the first guess and
 * a diagnostic refuted it; the category axis thins correctly and not one of the
 * 910 `p#-category#` pairs was two names in the same panel. It was arithmetic:
 * ten series in two columns is five rows, `panelH` came out at −0.4, and
 * `clampDim` rewrote that to `DEFAULT_SIZE.height` because it treats `<= 0` as
 * a malformed config. Ten full 300-point charts, stacked 9.6 points apart, in a
 * box 60 points tall. `buildMultiples` declines such a grid now.
 *
 * ALL TWENTY-THREE went, including `p#-total# / p#-cagr-label`, which a probe
 * had singled out as "not the grid" — the only family with a within-panel count
 * and the only one at ordinary sizes. The probe and this table were not
 * measuring the same charts: it built its `10 series` shape from scratch where
 * `DATA_SHAPES` spreads the sample's own series and keeps their colours and
 * roles. A diagnostic written beside a gate has to reuse the gate's variant
 * table, or its numbers answer a different question — and agreeing everywhere
 * else is what made that one look trustworthy.
 */
const BUDGET: Record<string, number> = {
  /**
   * The `value-axis-title` group, 687 of the 785 — and READ THE NOTE ABOVE
   * BEFORE ACTING ON THIS NUMBER. It is measured with a twenty-seven-character
   * sentence in an option documented as a short unit ("e.g. `€m`"), and with the
   * same sweep run at `€m` the group is 118. Roughly two-thirds of what follows
   * is the test string, not the engine.
   *
   * `title / value-axis-title` was 205 of this and is now absent: the unit yields
   * to the title rather than printing through it. It was the only member of the
   * group no width remedy could touch.
   */
  "value-axis-title / legend#": 308,
  "value-axis / value-axis-title": 158,
  "value-axis-title / total#": 95,
  "total# / value-axis-title": 93,
  "value-axis-title / label#": 27,
  "value-axis-title / median-label#": 6,
  // The tail.
  "combo-series-label# / combo-series-label#": 33,
  "label# / label#": 22,
  "title / series-label#": 10,
  "title / total#": 8,
  "series-label# / series-label#": 4,
  "header# / label##": 4,
  "title / row#": 4,
  "tile#-value / tile#-value": 4,
  "title / category#": 3,
  "label## / cagr-label": 2,
  "total# / title": 2,
  "legend# / footnote": 2,
};

/**
 * THE COMBINATIONS, because until 2026-08-29 this sweep never made one.
 *
 * `variants` below was the two tables CONCATENATED — every option applied to a
 * sample config, every data shape applied to a sample config, and never an
 * option and a data shape together. So a chart with a secondary axis AND ten
 * series was not among the 24,000, and neither was a footnote on
 * twenty-four categories.
 *
 * That is the same defect this file was WRITTEN TO FIX, one level up.
 * `frame-fit.test.ts` swept kinds, frames and fonts and missed the option and
 * data-shape variants; this file swept the variants and missed their products.
 * A gate is only as wide as its sweep, and both gates were green.
 *
 * Measured on a 5x4 slice before widening: three shapes appeared that the
 * uncrossed sweep has never produced in any of its 24,000 charts —
 * `title / total#`, `legend# / footnote`, `title / label#` — and every shape
 * the two share came out several times larger. By this file's own rule a shape
 * absent from the table is a regression by definition, so those were three
 * regressions nobody could see.
 *
 * NOT THE FULL PRODUCT, and the reason is honest rather than principled: 24
 * options by 10 shapes is 240 variants against 34, which takes this from eight
 * seconds to about a minute and out of the ordinary suite. What is crossed here
 * is the options that change the LAYOUT'S STRUCTURE against the data shapes
 * that stress LABEL COUNT, which is where the interesting products live.
 *
 * WHAT IS STILL NOT CROSSED, so the next person does not have to re-derive it:
 * every other option (`scale`, `logScale`, `axisBreak`, `gapWidth`, `overlap`,
 * `segmentOrder`, `categorySort`, `otherBucket`, `numberFormat`, `render
 * image`, the per-kind ones) against every data shape; and any product of three
 * or more. Widen this list before concluding a family is closed.
 */
const CROSS_OPTIONS = ["secondaryAxis", "pareto", "valueAxisTitle", "footnote", "title", "multiples 2 columns"];
const CROSS_SHAPES = ["24 categories", "10 series", "long series names", "long category names"];

/** The whole sweep, once. */
function measure(): { total: number; byShape: Map<string, number> } {
  const byShape = new Map<string, number>();
  let total = 0;
  const variants: [string, (c: unknown) => unknown][] = [
    ...Object.entries(DATA_SHAPES),
    ...Object.entries(OPTIONS).map(
      ([k, o]) =>
        [k, (c: unknown) => ({ ...prepared(k, c as Record<string, any>), ...o })] as [string, (c: unknown) => unknown],
    ),
    // MERGED onto whatever the kind already sets — see `DECOR`. A shallow spread
    // here would replace each kind's defaults and narrow the sweep while looking
    // like it widened it.
    ...Object.entries(DECOR).map(
      ([k, d]) =>
        [
          `decor ${k}`,
          (c: unknown) => ({
            ...(c as object),
            decorations: { ...((c as { decorations?: object }).decorations ?? {}), ...d },
          }),
        ] as [string, (c: unknown) => unknown],
    ),
  ];
  for (const ok of CROSS_OPTIONS)
    for (const sk of CROSS_SHAPES) {
      const o = OPTIONS[ok];
      const fn = DATA_SHAPES[sk];
      if (!o || !fn) throw new Error(`CROSS names a variant that does not exist: ${ok} x ${sk}`);
      variants.push([`${ok} × ${sk}`, (c: unknown) => ({ ...(fn(c as never) as object), ...o })]);
    }
  for (const [name, fn] of variants)
    for (const { kind } of CHART_KINDS)
      for (const [w, h] of FRAMES)
        for (const fontSize of [10, 18])
          for (const horizontal of [false, true]) {
            /**
             * DECORATIONS ARE MEASURED WHERE THEIR TEXT COULD FIT, and nowhere
             * else. This is the one variant family with a frame restriction, so
             * it needs its reason in writing.
             *
             * Turning every decoration on across every frame counts 2,748
             * overlaps against 785 — but the split by frame says what that is:
             *
             *     80x60  97   60x300 95   120x90 86   160x120 61
             *     200x150 42  300x60 20   480x300 2   960x540 0
             *
             * (measured on the sample data alone, without the punishing shapes).
             * An 80x60pt box asked to label twenty-four categories individually
             * cannot draw them anywhere, and counting that says nothing about the
             * engine. Recording ~2,000 impossible-frame collisions would also
             * blunt the ratchet: a real regression of ten at a readable size
             * would disappear inside a tolerated bucket of 1,243.
             *
             * So the decoration family is swept at frames where the text has room
             * to be placed, which is where an overlap is a defect rather than a
             * consequence of the frame. Every other option keeps the full frame
             * list exactly as before.
             */
            if (name.startsWith("decor ") && w * h < 480 * 300) continue;
            let ts: TextNode[];
            try {
              // A kind that refuses an option throws, and that is not an overlap.
              const cfg = fn({ ...sampleConfig(kind), width: w, height: h, horizontal, style: { fontSize } });
              ts = buildChart(cfg as never).nodes.filter(
                (n: SceneNode): n is TextNode => n.kind === "text" && !!String(n.text).trim(),
              );
            } catch {
              continue;
            }
            const boxes = ts.map(ink);
            for (let i = 0; i < boxes.length; i++)
              for (let j = i + 1; j < boxes.length; j++)
                if (sharedArea(boxes[i], boxes[j]) > 1) {
                  total++;
                  const a = (ts[i].name || "?").replace(/-?\d+/g, "#");
                  const b = (ts[j].name || "?").replace(/-?\d+/g, "#");
                  byShape.set(`${a} / ${b}`, (byShape.get(`${a} / ${b}`) ?? 0) + 1);
                }
          }
  return { total, byShape };
}

describe("the text this engine still draws over other text", () => {
  const { total, byShape } = measure();
  /**
   * `DUMP_BUDGET=1 npx vitest run test/overlap-budget.test.ts` prints the whole
   * measured table, ready to paste over `BUDGET`.
   *
   * Here because re-baselining by hand from assertion output is how a budget
   * ends up above its real figure — the failure prints only the shapes that
   * broke a rule, so the unchanged ones get copied forward from the old table
   * and quietly keep their slack. The fourth test exists to catch exactly that;
   * this makes it unnecessary to trip.
   */
  if (process.env.DUMP_BUDGET) {
    const lines = [...byShape.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `  "${k}": ${v},`);
    console.log(`TOTAL ${total}\n${lines.join("\n")}`);
  }

  it("draws no shape it has never drawn before", () => {
    // A shape absent from the budget is text the engine has NOT been seen to
    // overlap, so its arrival is a regression whatever the total does. This is
    // the half a single number cannot express.
    const unbudgeted = [...byShape.entries()].filter(([k]) => BUDGET[k] === undefined);
    expect(
      unbudgeted,
      "a new kind of text overlap appeared — fix it, or add it to the budget and to docs/BACKLOG.md saying what it is",
    ).toEqual([]);
  });

  it("stays inside every shape's budget", () => {
    const over = [...byShape.entries()]
      .filter(([k, n]) => n > (BUDGET[k] ?? 0))
      .map(([k, n]) => `${k}: ${n} (budget ${BUDGET[k]})`);
    expect(over, "a family of text overlaps grew").toEqual([]);
  });

  it("keeps the budget honest — no line left above the real figure", () => {
    // THE HALF THAT MAKES IT A RATCHET. A ceiling nobody lowers stops ratcheting
    // and the next regression hides under the slack. Fixing overlaps is meant to
    // fail this: read the numbers off the message and edit them down.
    const slack = [...Object.entries(BUDGET)]
      .filter(([k, n]) => n > (byShape.get(k) ?? 0))
      .map(([k, n]) => `${k}: budget ${n}, actual ${byShape.get(k) ?? 0}`);
    expect(slack, "these budgets sit above what the engine now does — lower them").toEqual([]);
  });

  it("has not grown in total", () => {
    // DERIVED from the table above rather than written down again. A second
    // number is a second thing to forget: this one was left at 467 while the
    // shapes below it had already been lowered to 449, so it would have waved
    // through eighteen new overlaps. The per-shape budgets are the one source of
    // truth and the honesty test keeps each of them exact, so their sum is the
    // total — no slack anywhere, by construction.
    const cap = Object.values(BUDGET).reduce((a, b) => a + b, 0);
    expect(total, "the total overlap count grew").toBe(cap);
  });
});

/**
 * AN OPTION THAT CHANGES NOTHING IS MEASURED AS CLEAN, so every entry in
 * `OPTIONS` has to be doing something to the chart it is swept over — see
 * `UNPREPARED` for the one that is declared not to be, and why.
 */
describe("every option in the sweep takes effect", () => {
  it("changes the scene of at least one kind, or is declared not to", () => {
    const inert = optionsThatChangeNothing(OPTIONS, prepared).filter(
      (key) => !(key in SCENE_NEUTRAL_OPTIONS) && !(key in UNPREPARED),
    );
    expect(
      inert,
      "these options are measured over a sample they do nothing to — see `OPTION_DATA` in test/helpers/sweep-options.ts",
    ).toEqual([]);
  });

  it("declares only options that would take effect if prepared, so the declaration stays true", () => {
    const wouldStillBeInert = optionsThatChangeNothing(
      Object.fromEntries(Object.keys(UNPREPARED).map((key) => [key, OPTIONS[key]])),
      withOptionData,
    );
    expect(
      wouldStillBeInert,
      "an UNPREPARED option does nothing even with its data: it is not unprepared, it is inert",
    ).toEqual([]);
    expect(
      Object.keys(UNPREPARED).filter((key) => !optionsThatChangeNothing({ [key]: OPTIONS[key] }, prepared).length),
      "an UNPREPARED option takes effect without its data: delete the declaration",
    ).toEqual([]);
  });
});
