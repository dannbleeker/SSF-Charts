import type { ChartConfig, ChartData, ChartKind, ChartStyle, Decorations } from "./types";
import type { Scene } from "./scene";
import { maxOf, minOf } from "./agg";
import { DEFAULT_DECOR, DEFAULT_STYLE, seriesColor } from "./style";
import { layoutColumns, layoutCombo } from "./layout/column";
import { layoutWaterfall, waterfallExtent } from "./layout/waterfall";
import { layoutMekko } from "./layout/mekko";
import { layoutLine } from "./layout/line";
import { layoutButterfly } from "./layout/butterfly";
import { layoutScatter, spreadCap } from "./layout/scatter";
import { layoutGantt } from "./layout/gantt";
import { layoutPie } from "./layout/pie";
import { boxplotExtent, layoutBoxplot } from "./layout/boxplot";
import { layoutRadar } from "./layout/radar";
import { layoutHeatmap } from "./layout/heatmap";
import { layoutTilemap } from "./layout/tilemap";
import { layoutCascade } from "./layout/cascade";
import { layoutFunnel } from "./layout/funnel";
import { layoutWaffle } from "./layout/waffle";
import { layoutTreemap } from "./layout/treemap";
import { layoutSunburst } from "./layout/sunburst";
import { layoutViolin } from "./layout/violin";
import { layoutCandlestick } from "./layout/candlestick";
import { footnoteNode, hundredPercentTotal, titleHeight, titleNode } from "./layout/frame";
import { bandNodes, decorationNodes } from "./decor";
import { resolveLabelCollisions, unplaceableComboLabels, tickLabelsUnderPointLabels } from "./collide";
import { formatNumber, niceTicks, parseDateToken, resolveFormat, GANTT_DATE_ROW } from "./format";
import type { SceneNode } from "./scene";
import { clipTextToFrame, finiteNodes } from "./scene";
import type { LayoutResult } from "./layout/column";

export const DEFAULT_SIZE = { width: 480, height: 300 };

/**
 * The kinds whose layout actually reads `horizontal` and rotates. Every other
 * kind draws the same chart either way — see `normalizeConfig`, which drops the
 * flag where it cannot apply so that nothing downstream acts on a rotation that
 * never happened.
 *
 * Kept as a list rather than derived, because there is no way to ask a layout
 * whether it consulted the flag. `test/layouts.test.ts` builds every kind both
 * ways and fails if the list and the layouts disagree, in either direction.
 */
export const ROTATABLE_KINDS: ReadonlySet<ChartKind> = new Set<ChartKind>([
  "stacked",
  "clustered",
  "stacked100",
  "combo",
  "waterfall",
  "mekko",
  "line",
  "area",
  "boxplot",
]);

/**
 * A slide is 13.33 inches wide. Nothing on one is a hundred.
 *
 * pptxgenjs has a documented rule — "any number >= 100 sure isn't inches,
 * assume it is EMU already" — so the moment a dimension crosses 100 INCHES
 * (7200pt) it writes the value through unrounded: `<a:ext cx="101.388…"/>`.
 * `ST_PositiveCoordinate` is an xsd:long, so that part is schema-invalid and
 * what the user meets is PowerPoint's repair dialog. The renderer reported
 * success, and two of this repo's three gates disagreed about the file.
 *
 * 7200pt is already fifteen times the widest slide, so the ceiling costs
 * nothing real and turns an unopenable deck into a very large chart.
 *
 * 7199, NOT 7200, and the single point matters: pptxgenjs treats any number
 * `>= 100` as EMU already, and 7200pt is EXACTLY 100 inches — `7200 / 72 === 100`
 * with no float slack at all. So the old ceiling did not avoid the trap, it
 * landed precisely on it, and `Math.min` meant every oversize request was
 * clamped TO the one broken value. A chart at `width: 10000` shipped its title
 * box as `<a:ext cx="100"/>` — 100 EMU, a ten-thousandth of an inch — beside a
 * sibling frame at `cx="91440000"`, with the CLI reporting success and
 * `validate-ooxml` reporting schema-valid.
 */
const MAX_DIM = 7199;

/** A positive, finite dimension within the ceiling, or the fallback. */
/**
 * A chart dimension the rest of the engine can rely on.
 *
 * Exported because the PANE needs the same rule before `buildChart` ever sees
 * the config. `app.ts` builds a size for `placeChart` as
 * `cfg.width ?? DEFAULT_SIZE.width`, and `??` catches only null and undefined —
 * so `width: NaN` from a pasted config, a saved template or a shape tag written
 * in another deck went straight to the placer, which shrank it to fit and
 * returned **882** points. Wider than the 720pt slide, finite, and therefore
 * waved through by this very function on the way back in: the user gets a chart
 * running off the edge of the slide with nothing said.
 *
 * One rule, one place. The alternative is two clamps that agree until somebody
 * changes one.
 */
export function clampDim(v: number | undefined, fallback: number): number {
  if (typeof v !== "number" || !Number.isFinite(v) || v <= 0) return fallback;
  return Math.min(v, MAX_DIM);
}

/**
 * The point size a chart's text may be AUTHORED at.
 *
 * This is `clampDim`'s missing twin. The guard below it refused a fontSize that
 * was not a positive finite number and stopped there — no ceiling, no floor —
 * where the dimension rule has `MAX_DIM` and says why. So `fontSize: 1e6` and
 * `fontSize: 0.0001` sailed through, and the pptx renderer wrote them into
 * OOXML: `sz="120000000"` and `sz="0"`, against a `ST_TextFontSize` of
 * 100..400000 (1pt..4000pt in hundredths). Both are files PowerPoint offers to
 * repair, produced by a CLI that reported success.
 *
 * The authored bound has to sit inside the schema's, because the layout scales
 * it before the renderer writes it — by at most 1.2 (a title) and at least 0.85
 * (a small label), measured across every kind. 2..2000 becomes 1.7..2400 by the
 * time it is written, comfortably inside 1..4000, and it is two orders of
 * magnitude either side of the 10-12pt a real chart uses.
 *
 * Clamped rather than replaced with the default, unlike the unreadable case: a
 * caller asking for 5000pt has said something about intent, and the largest
 * size the format can carry honours it better than 12 does.
 */
const MIN_FONT = 2;
const MAX_FONT = 2000;
export function clampFontSize(v: unknown): number {
  if (typeof v !== "number" || !Number.isFinite(v) || v <= 0) return DEFAULT_STYLE.fontSize;
  return Math.min(MAX_FONT, Math.max(MIN_FONT, v));
}

/** Coerce a data block to trustworthy arrays: every series padded to the
 *  category count, non-finite cells nulled, per-category arrays aligned. */
// Hard ceilings on grid size. buildChart is fed arbitrary authored JSON (the
// skill) and configs round-tripped from shape tags, so a config claiming
// millions of categories × thousands of series would otherwise make normalizeData
// allocate ~10⁹ cells and OOM the headless renderer. These caps sit far above any
// real chart (a multi-year daily calendar heatmap is ~1–4k days, business charts
// rarely exceed a few dozen series), so truncating past them is a no-op for valid
// input and a bound on abusive input.
const MAX_CATEGORIES = 4096;
const MAX_SERIES = 256;

/**
 * A label as text, for whatever turned up where a label belongs.
 *
 * Numeric categories are not a hostile input, they are an ordinary one — years,
 * quarters, ids — and `categories: [2023, 2024]` used to throw
 * `raw.trim is not a function` and `c.split is not a function` out of the
 * layout. `String()` gives back the label the author meant; null and undefined
 * become blank, which is what a gap in a category strip should look like.
 */
const labelText = (v: unknown): string => (v == null ? "" : typeof v === "string" ? v : String(v));

function normalizeData(raw: ChartData): ChartData {
  // A config with no data at all is REFUSED, deliberately and by name.
  //
  // It used to be refused by accident: `data?.` on two of four reads meant the
  // other two threw `Cannot read properties of null`, and two tests quietly
  // relied on that TypeError to prove a batch isolates a bad chart. Removing
  // the crash removed the refusal with it, and a chart with nothing in it began
  // rendering as a silent empty frame — which for the batch renderer is worse
  // than an error, because an agent gets back a blank chart instead of a
  // sentence telling it what it got wrong.
  //
  // So: the same outcome, on purpose and with something readable in it. An
  // EMPTY `{}` is still fine — "no series" and "no categories" are legitimate
  // states with their own tests — this is only about there being no data object
  // to speak of.
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    throw new Error("chart config has no `data` object (needs at least `{ categories: [], series: [] }`)");
  }
  const data = raw;
  const categories = (Array.isArray(data.categories) ? data.categories : []).slice(0, MAX_CATEGORIES).map(labelText);
  const n = categories.length;
  const cell = (v: number | null | undefined): number | null => (v == null ? null : Number.isFinite(v) ? v : null);
  /**
   * DID A GANTT ROW ACTUALLY ARRIVE AS A CALENDAR?
   *
   * The conversion below turns an ISO string into days-since-epoch, and the
   * number that comes out cannot say where it came from. `layoutGantt` gates
   * EVERYTHING calendar on `data.dates` — week and month ticks, the tick label
   * format, the today marker — so without this the axis of a date Gantt was
   * drawn by `niceTicks` and labelled in raw epoch days: "20458", "20486".
   *
   * Only `datasheet.ts` set the flag, so only a PASTE produced a calendar. The
   * skill's own documented config — ISO strings in Start/End rows, which
   * SKILL.md states as a hard rule — came through here, had its dates parsed
   * correctly, and then rendered as a numeric axis. Bars land in the right
   * relative places, which is what made it look deliberate.
   *
   * `|| undefined` rather than `false`, matching `sheetToData`'s own idiom, so
   * a config that never carried dates is byte-identical to before.
   */
  let sawDate = false;
  const series = (Array.isArray(data.series) ? data.series : [])
    .slice(0, MAX_SERIES)
    // A null entry survived as `{ values: [...] }` with no name, and every
    // consumer that reached for `s.name.trim()` died on it. JSON arrays hold
    // nulls; a series that is not an object is not a series.
    .filter((s): s is NonNullable<typeof s> => !!s && typeof s === "object")
    .map((s) => {
      const rawValues = Array.isArray(s.values) ? s.values : [];
      // A Gantt Start/End row may arrive as ISO strings, because that is what
      // the skill tells an agent to write: SKILL.md says "ISO dates supported"
      // and, as a hard Rule, "dates as ISO strings only in Gantt rows". Only the
      // datasheet ever parsed them. Through any other door — the skill, a
      // hand-written JSON, a POWERCHART_CONFIG tag — every string fell to `null`
      // here, so `layoutGantt` saw tasks with no Start and no End, read each as
      // a section header, and drew a plan slide with the task names on grey
      // bands: no bars, no date labels, no timeline. Silently, exit 0, looking
      // deliberate. That is the skill's primary use case.
      const dateRow = GANTT_DATE_ROW.test(labelText(s.name).trim());
      const values = Array.from({ length: n }, (_, c) => {
        const raw = rawValues[c];
        if (!dateRow || typeof raw !== "string") return cell(raw);
        const day = parseDateToken(raw);
        if (day != null) sawDate = true;
        return day;
      });
      // ALWAYS a string, including when there was no name at all.
      //
      // This used to read `s.name == null ? s : …`, on the reasoning that an
      // unnamed series is legitimate and coercing `undefined` to `""` would
      // give it a name it never had. The first half is true and the second was
      // wrong twice over. `Series.name` is `string`, not `string | undefined`,
      // and five places in this file act on it unguarded —
      // `CARRIED_ROW.test(s.name.trim())` and its four siblings — so a series
      // written as `{ values: [1, 2, 3] }`, which is the obvious way to write a
      // single-series chart, crashed SEVENTEEN of the twenty-five kinds on
      // `Cannot read properties of undefined (reading 'trim')`. Not a hostile
      // input; an ordinary one.
      //
      // And `""` does not give it a name: the only consumer that lists names
      // does `series.map((s) => s.name).filter(Boolean)`, which drops an empty
      // name exactly as it dropped an absent one.
      const named = { ...s, name: labelText(s.name) };
      if (named.colors) {
        return { ...named, values, colors: Array.from({ length: n }, (_, c) => named.colors![c] ?? null) };
      }
      return { ...named, values };
    });
  // `Array.isArray`, not truthiness. The guard used to be `arr ? … : arr`,
  // which hands a FALSY non-array straight back: `xExtent: NaN` survived
  // untouched and `data.xExtent?.some(...)` in the mekko layout then died on
  // `some is not a function` — the optional chain does not help, because NaN is
  // not nullish. Truthy non-arrays were already repaired by accident, since
  // `Array.from` reads them by index; this makes the two cases agree.
  const pad = <T>(arr: (T | null)[] | undefined): (T | null)[] | undefined =>
    Array.isArray(arr) ? Array.from({ length: n }, (_, c) => arr[c] ?? null) : undefined;
  return {
    ...data,
    categories,
    series,
    hundredPercent: pad(data.hundredPercent),
    xExtent: pad(data.xExtent),
    // An author who wrote `dates` keeps what they wrote; otherwise it is set
    // only when a Gantt row really did parse as a calendar. See `sawDate`.
    dates: data.dates ?? (sawDate || undefined),
  };
}

/**
 * Decoration keys whose value is a list of OBJECTS the layout then iterates.
 *
 * Each one is read as `decor.<key>.forEach(...)` and each entry as a record
 * with named fields, with nothing in between. A value of the wrong shape
 * therefore does not mis-render, it CRASHES: `valueLines: "mean"` gives
 * `valueLines.forEach is not a function` on nine of the twenty-five kinds,
 * `callouts` on thirteen and `bands` on fifteen, and a null entry dies one
 * field later on `Cannot read properties of null`.
 *
 * All of them are ordinary things to write. `valueLines` mirrors a checkbox in
 * the pane labelled "mean", so `"valueLines": "mean"` is the obvious guess for
 * anyone hand-editing the JSON box — and a config also arrives from a saved
 * template, from a POWERCHART_CONFIG tag authored in another deck, and from the
 * skill's caller.
 *
 * `labelContent` is deliberately NOT here. It is a list of strings read with
 * `.includes`, which a bare string answers plausibly rather than fatally, and
 * changing that would change a rendering rather than prevent a crash.
 */
const OBJECT_LIST_DECORATIONS = ["valueLines", "callouts", "bands"] as const;

/**
 * Keep only the entries of those lists that the layout can actually read: an
 * array of non-null objects, or nothing at all.
 *
 * Dropping beats repairing here. There is no sensible reading of `"mean"` as a
 * value line — the entry it stands for is `{ mode: "mean" }`, and guessing that
 * would make a typo silently mean something — so a decoration that cannot be
 * understood is simply not drawn, which is what omitting it already does.
 */
function cleanDecorations(decor: Partial<Decorations>): Partial<Decorations> {
  let out: Record<string, unknown> | undefined;
  // `fillBetween` is a pair of series indices, and the line layout DESTRUCTURES
  // it — `const [ai, bi] = decor.fillBetween` — which throws on anything that
  // is not iterable rather than reading undefined. Same class as the lists
  // below, different shape, so it is checked as the pair it is.
  const fb: unknown = decor.fillBetween;
  if (fb !== undefined && !(Array.isArray(fb) && fb.length === 2 && fb.every((v) => Number.isFinite(v)))) {
    out = { ...decor } as Record<string, unknown>;
    delete out.fillBetween;
  }
  for (const key of OBJECT_LIST_DECORATIONS) {
    const v: unknown = decor[key];
    if (v === undefined) continue;
    const kept = Array.isArray(v) ? v.filter((e) => !!e && typeof e === "object" && !Array.isArray(e)) : [];
    // A no-op unless something was actually wrong, so a valid config keeps its
    // own array object and renders byte-identically.
    if (Array.isArray(v) && kept.length === v.length) continue;
    out ??= { ...decor } as Record<string, unknown>;
    if (kept.length) out[key] = kept;
    else delete out[key];
  }
  return (out as Partial<Decorations> | undefined) ?? decor;
}

/** Drop label nudges carrying a non-finite delta (they corrupt a node's coords). */
function cleanOffsets(offs: Record<string, { dx: number; dy: number }>): Record<string, { dx: number; dy: number }> {
  const out: Record<string, { dx: number; dy: number }> = {};
  for (const [k, o] of Object.entries(offs)) {
    if (o && Number.isFinite(o.dx) && Number.isFinite(o.dy)) out[k] = o;
  }
  return out;
}

/**
 * Coerce a possibly-malformed ChartConfig into one the layout engine can trust.
 *
 * buildChart is now fed by the Claude skill (arbitrary authored JSON) and by
 * configs round-tripped out of slide shape tags, so it can no longer assume its
 * input is well-formed. This REPAIRS rather than throws — a chart with a clamped
 * size beats a stack trace in a headless render — and every branch is a no-op
 * unless the field is actually out of range, so valid configs render identically.
 */
export function normalizeConfig(cfg: ChartConfig): ChartConfig {
  const width = clampDim(cfg.width, DEFAULT_SIZE.width);
  const height = clampDim(cfg.height, DEFAULT_SIZE.height);

  // `horizontal` is a request to rotate, and only these kinds have a layout
  // that can honour it. On every other kind the layout ignores the flag — but
  // the decoration stage did not: `skipDecor` and the value-axis map keyed off
  // the raw flag, so setting it on a treemap, sunburst, violin or candlestick
  // drew a byte-identical chart with its CAGR arrow, difference arrow, value
  // lines, callouts and Error/Target marks silently removed. The pane's
  // "Horizontal (bar)" toggle is offered for every kind and its state survives
  // a kind change, so that is one click away.
  //
  // Dropped here rather than at each reader, so the layouts, the decoration
  // gate and the accessible description all see the same world. `test/
  // layouts.test.ts` pins the set against what the layouts actually do.
  const horizontal = cfg.horizontal && ROTATABLE_KINDS.has(cfg.kind) ? cfg.horizontal : undefined;

  // Order a reversed manual scale and drop non-finite ends — a NaN or inverted
  // bound reaches niceTicks and blanks the whole value axis.
  let scale = cfg.scale;
  if (scale) {
    let min = typeof scale.min === "number" && Number.isFinite(scale.min) ? scale.min : undefined;
    let max = typeof scale.max === "number" && Number.isFinite(scale.max) ? scale.max : undefined;
    if (min != null && max != null && min > max) [min, max] = [max, min];
    scale = { ...(min != null ? { min } : {}), ...(max != null ? { max } : {}) };
  }

  // `style.fontSize` is arithmetic, not decoration: every label box, every
  // decoration lift and half the axis geometry is derived from it. A string
  // there — and a config arrives from JSON import, from a template, from a
  // POWERCHART_CONFIG tag authored in another deck — turns those into NaN and
  // poisons the LAYOUT, not just the label. The renderers can each defend
  // their own attributes (SVG does); none of them can put the chart back
  // where it belonged. Repaired here, once, on the way in, exactly as the
  // dimensions and the manual scale already are.
  const style =
    cfg.style && clampFontSize(cfg.style.fontSize) !== cfg.style.fontSize
      ? { ...cfg.style, fontSize: clampFontSize(cfg.style.fontSize) }
      : cfg.style;

  return {
    ...cfg,
    width,
    height,
    scale,
    style,
    horizontal,
    data: normalizeData(cfg.data),
    decorations:
      cfg.decorations && typeof cfg.decorations === "object" && !Array.isArray(cfg.decorations)
        ? cleanDecorations(cfg.decorations)
        : cfg.decorations,
    labelOffsets: cfg.labelOffsets ? cleanOffsets(cfg.labelOffsets) : cfg.labelOffsets,
  };
}

const SORTABLE: ChartKind[] = ["stacked", "clustered", "stacked100", "mekko", "pie", "doughnut", "butterfly"];

/**
 * Move every per-category decoration through a category permutation.
 *
 * `order[newPosition] = oldIndex`. Sorting and Pareto already carry `colors`
 * along for the stated reason that a highlight belongs to a data point, not to
 * a screen position — but the decorations were copied through untouched, so a
 * callout, a CAGR arrow or a difference arrow stayed on the slot the category
 * used to occupy. That is worse than misplaced ink: `decorationNodes` reads
 * `columnValue[from]`/`[to]` at those indices, so the arrows also PRINTED a
 * growth rate for a pair of categories the author never named.
 *
 * An x-band spans a range, and an arbitrary permutation does not keep a range
 * contiguous; its endpoints are moved and re-ordered, which is the best a
 * span can do.
 */
/** Old category index → new one, after `order` has been applied. */
function positionMap(order: number[]): (i: number) => number {
  const pos = new Map(order.map((oldIdx, newIdx) => [oldIdx, newIdx]));
  return (i: number) => pos.get(i) ?? i;
}

function remapDecorations<T extends Partial<Decorations> | undefined>(decor: T, order: number[]): T {
  if (!decor) return decor;
  const at = positionMap(order);
  const pair = <T extends { from: number; to: number }>(p: T): T => {
    const [from, to] = [at(p.from), at(p.to)].sort((a, b) => a - b);
    return { ...p, from, to };
  };
  return {
    ...decor,
    ...(decor.callouts ? { callouts: decor.callouts.map((c) => ({ ...c, category: at(c.category) })) } : {}),
    ...(decor.cagr ? { cagr: pair(decor.cagr) } : {}),
    ...(decor.difference ? { difference: pair(decor.difference) } : {}),
    ...(decor.bands ? { bands: decor.bands.map((b) => (b.axis === "x" ? pair(b) : b)) } : {}),
  };
}

/** Reorder categories (and every per-category array) by column total. */
function sortCategories(cfg: ChartConfig): ChartConfig {
  if (!cfg.categorySort || !SORTABLE.includes(cfg.kind)) return cfg;
  const { data } = cfg;
  // Sort runs before error/target rows are extracted; those carried rows are
  // not real stack contributors, so exclude them from the ranking totals.
  const ranked = data.series.filter((s) => !CARRIED_ROW.test(s.name.trim()));
  const totals = data.categories.map((_, c) => ranked.reduce((a, s) => a + (s.values[c] ?? 0), 0));
  const sign = cfg.categorySort === "ascending" ? 1 : -1;
  const order = data.categories.map((_, c) => c).sort((a, b) => sign * (totals[a] - totals[b]));
  const pick = <T>(arr: T[] | undefined) => (arr ? order.map((c) => arr[c]) : undefined);
  return {
    ...cfg,
    data: {
      ...data,
      categories: order.map((c) => data.categories[c]),
      // colors is per-category too: leaving it in pre-sort order pins a
      // highlight to a screen position instead of to its data point.
      series: data.series.map((s) => ({
        ...s,
        values: order.map((c) => s.values[c]),
        ...(s.colors ? { colors: pick(s.colors) } : {}),
      })),
      hundredPercent: pick(data.hundredPercent),
      xExtent: pick(data.xExtent),
    },
    decorations: remapDecorations(cfg.decorations, order),
    // `pie.explode` and `pie.breakout` are category indices too — they just
    // live on the top-level config rather than under `decorations`, so
    // `remapDecorations` never saw them. Unpermuted, `explode: [0]` offset
    // whichever category the SORT put in slot 0 and `breakout: [0, 2]`
    // collapsed whichever two landed there: a different set of data points
    // than the author named. That is the exact failure `remapDecorations`'
    // docstring exists for — a highlight belongs to a data point, not to a
    // screen position — and pie/doughnut are in `SORTABLE`.
    ...(cfg.pie ? { pie: remapPieIndices(cfg.pie, order) } : {}),
  };
}

function remapPieIndices(pie: NonNullable<ChartConfig["pie"]>, order: number[]): NonNullable<ChartConfig["pie"]> {
  const at = positionMap(order);
  return {
    ...pie,
    ...(pie.explode ? { explode: pie.explode.map(at) } : {}),
    ...(pie.breakout ? { breakout: pie.breakout.map(at) } : {}),
  };
}

/** Column-family kinds whose series can be collapsed into an "Other" bucket. */
const OTHER_KINDS: ChartKind[] = ["stacked", "clustered", "stacked100"];

/**
 * Collapse the long tail of series into one "Other" segment: keep the
 * (max − 1) largest by absolute total and sum the rest into "Other", so the
 * result has at most `max` series. No-op when already within budget.
 */
function collapseOther(cfg: ChartConfig): ChartConfig {
  const ob = cfg.otherBucket;
  if (!ob || !OTHER_KINDS.includes(cfg.kind)) return cfg;
  const max = Math.max(2, Math.floor(ob.max ?? 5));
  const series = cfg.data.series;
  if (series.length <= max) return cfg;
  const totals = series.map((s) => s.values.reduce((a: number, v) => a + Math.abs(v ?? 0), 0));
  const rank = series.map((_, i) => i).sort((a, b) => totals[b] - totals[a]);
  const keep = new Set(rank.slice(0, max - 1));
  const kept = series.filter((_, i) => keep.has(i)); // original order preserved
  const otherVals = cfg.data.categories.map((_, c) =>
    series.reduce((a, s, i) => (keep.has(i) ? a : a + (s.values[c] ?? 0)), 0),
  );
  return {
    ...cfg,
    data: { ...cfg.data, series: [...kept, { name: cfg.labels?.other ?? "Other", values: otherVals }] },
  };
}

/**
 * Pareto helper: sort categories by the first (non-line) series descending and
 * overlay a computed cumulative-% line on a secondary axis — the classic 80/20
 * view. Rewrites the config into a combo.
 */
function applyPareto(cfg: ChartConfig): ChartConfig {
  if (!cfg.pareto) return cfg;
  // The bars are the first series that is neither overlay kind. Testing only
  // against "line" would rank the chart by a marker series when one is written
  // first — sorting the columns by their own benchmark.
  const bar = cfg.data.series.find((s) => s.type !== "line" && s.type !== "marker");
  if (!bar) return cfg;
  const order = cfg.data.categories.map((_, c) => c).sort((a, b) => (bar.values[b] ?? 0) - (bar.values[a] ?? 0));
  const pick = <T>(arr: T[] | undefined) => (arr ? order.map((c) => arr[c]) : undefined);
  const barVals = order.map((c) => bar.values[c] ?? 0);
  const total = barVals.reduce((a, v) => a + Math.max(0, v), 0) || 1;
  let run = 0;
  const cum = barVals.map((v) => {
    run += Math.max(0, v);
    return Math.round((run / total) * 1000) / 10;
  });
  return {
    ...cfg,
    kind: "combo",
    secondaryAxis: true,
    data: {
      ...cfg.data,
      categories: order.map((c) => cfg.data.categories[c]),
      series: [
        // colors follows the permutation, like values — Pareto's whole point is
        // that the ranked order changes, and a highlight has to travel with it.
        { ...bar, values: barVals, ...(bar.colors ? { colors: pick(bar.colors) } : {}) },
        { name: "Cumulative %", type: "line", values: cum },
      ],
      hundredPercent: pick(cfg.data.hundredPercent),
    },
    decorations: remapDecorations(cfg.decorations, order),
  };
}

/** "After" / "Dep" row: a 1-based predecessor index, so it moves with the rows. */
const AFTER_ROW = /^(after|dep(endency)?)$/i;

/**
 * Gantt lanes: regroup task rows under one synthesized header per owner.
 *
 * A stable partition, not a sort — inside a lane the rows keep the order they
 * were written in, because that order is the plan. (categorySort refuses to
 * touch a Gantt for exactly that reason; this reorders only on request.) Tasks
 * with no owner keep to themselves at the end.
 *
 * The subtle part is `After`: its values are 1-based ROW INDICES, so moving
 * rows without renumbering them would silently point every dependency arrow at
 * the wrong task.
 */
function applyGanttLanes(cfg: ChartConfig): ChartConfig {
  if (cfg.kind !== "gantt" || cfg.gantt?.lanes !== "owner") return cfg;
  const { data } = cfg;
  const part = (c: string, i: number) => (c.split("|")[i] ?? "").trim();
  const ownerOf = (c: string) => part(c, 1);
  const lanes: string[] = [];
  for (const c of data.categories) {
    const o = ownerOf(c);
    if (o && !lanes.includes(o)) lanes.push(o);
  }
  if (!lanes.length) return cfg;

  /** New row order: -1 marks a synthesized header. */
  const perm: number[] = [];
  const labels: string[] = [];
  const push = (i: number) => {
    const c = data.categories[i];
    const act = part(c, 0).replace(/^>+\s*/, "");
    const remark = part(c, 2);
    // Drop the owner from the row — the lane header carries it now.
    labels.push(remark ? `> ${act} |  | ${remark}` : `> ${act}`);
    perm.push(i);
  };
  for (const lane of lanes) {
    labels.push(lane);
    perm.push(-1);
    data.categories.forEach((c, i) => {
      if (ownerOf(c) === lane) push(i);
    });
  }
  // Unassigned rows keep their own order, after the lanes.
  data.categories.forEach((c, i) => {
    if (!ownerOf(c)) push(i);
  });

  const oldToNew = new Map<number, number>();
  perm.forEach((old, next) => {
    if (old >= 0) oldToNew.set(old, next);
  });
  return {
    ...cfg,
    data: {
      ...data,
      categories: labels,
      series: data.series.map((s) => ({
        ...s,
        values: perm.map((old, next) => {
          if (old < 0) return null; // synthesized header: no bar
          const v = s.values[old];
          if (v == null || !AFTER_ROW.test(s.name.trim())) return v;
          // Renumber the dependency onto its row's new position.
          const target = oldToNew.get(Math.round(v) - 1);
          return target == null || target === next ? null : target + 1;
        }),
      })),
    },
  };
}

/** Datasheet rows carrying error-bar deltas: Error (±), Error+ / Error−. */
const ERROR_ROW = /^error\s*([+\-−])?$/i;
/** Bullet-chart target row: a bold tick across each column at the value. */
const TARGET_ROW = /^target$/i;
const ERROR_KINDS: ChartKind[] = ["stacked", "clustered", "line", "area"];
/** Kinds that accept a Target row (error bars stay column/line-family). */
const TARGET_KINDS: ChartKind[] = [...ERROR_KINDS, "waterfall"];

/**
 * Pull Error rows out of the data (so they don't render as segments) and
 * return per-category plus/minus deltas. Bars anchor at the column total
 * (single-series charts: the value) or the first line series.
 */
function extractErrorRows(cfg: ChartConfig): {
  cfg: ChartConfig;
  errors: { plus: (number | null)[]; minus: (number | null)[] } | null;
  targets: (number | null)[] | null;
} {
  if (!TARGET_KINDS.includes(cfg.kind)) return { cfg, errors: null, targets: null };
  const rows = ERROR_KINDS.includes(cfg.kind) ? cfg.data.series.filter((s) => ERROR_ROW.test(s.name.trim())) : [];
  const targetRow = cfg.data.series.find((s) => TARGET_ROW.test(s.name.trim()));
  if (!rows.length && !targetRow) return { cfg, errors: null, targets: null };
  const pick = (sign: "+" | "-") =>
    cfg.data.categories.map((_, c) => {
      for (const r of rows) {
        const m = r.name.trim().match(ERROR_ROW)!;
        const dir = m[1] === "−" ? "-" : (m[1] ?? "both");
        if ((dir === "both" || dir === sign) && r.values[c] != null) return Math.abs(r.values[c]!);
      }
      return null;
    });
  return {
    cfg: {
      ...cfg,
      data: {
        ...cfg.data,
        series: cfg.data.series.filter((s) => !ERROR_ROW.test(s.name.trim()) && !TARGET_ROW.test(s.name.trim())),
      },
    },
    errors: rows.length ? { plus: pick("+"), minus: pick("-") } : null,
    targets: targetRow ? cfg.data.categories.map((_, c) => targetRow.values[c] ?? null) : null,
  };
}

/** Kinds whose multi-series data splits cleanly into per-series panels. */
const MULTIPLES_KINDS: ChartKind[] = ["stacked", "clustered", "line", "area", "waterfall", "radar"];
/** Special rows carried into every small-multiples panel (not split). */
const CARRIED_ROW = /^(error\s*[+\-−]?|target|band\s*(low|high))$/i;

/** Shift a scene's nodes by (dx, dy) and prefix their names. */
function translateNodes(nodes: SceneNode[], dx: number, dy: number, prefix: string): SceneNode[] {
  return nodes.map((node) => {
    const n = { ...node, name: node.name ? `${prefix}${node.name}` : undefined };
    switch (n.kind) {
      case "rect":
      case "text":
      case "chevron":
      case "arrowhead":
        n.x += dx;
        n.y += dy;
        break;
      case "line":
        n.x1 += dx;
        n.y1 += dy;
        n.x2 += dx;
        n.y2 += dy;
        break;
      case "ellipse":
      case "wedge":
      case "symbol":
        n.cx += dx;
        n.cy += dy;
        break;
      case "polygon":
        n.points = n.points.map((p) => ({ x: p.x + dx, y: p.y + dy }));
        break;
      default: {
        // A new SceneNode kind that reaches here would be silently left at the
        // origin panel's coordinates — a misplaced shape, with nothing failing.
        // This turns that into a compile error at the point of adding the kind.
        const unreached: never = n;
        return unreached;
      }
    }
    return n;
  });
}

/**
 * Small multiples: one single-series panel per data series, titled by the
 * series name, laid out in a grid and pinned to one shared value scale so
 * panels compare honestly. Null when the config doesn't call for it.
 */
function buildMultiples(rawCfg: ChartConfig): Scene | null {
  const multiples = rawCfg.multiples;
  if (!multiples || !MULTIPLES_KINDS.includes(rawCfg.kind)) return null;
  // Sort categories ONCE, ranked across the full data, so every panel shares one
  // x-axis. (Letting each panel sort itself — it inherits categorySort — ranked
  // each panel by its own single series and gave the panels contradictory axes.)
  const cfg = sortCategories(rawCfg);
  const carried = cfg.data.series.filter((s) => CARRIED_ROW.test(s.name.trim()));
  const dataSeries = cfg.data.series.filter((s) => !CARRIED_ROW.test(s.name.trim()));
  if (dataSeries.length < 2) return null;

  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const fs = style.fontSize;
  const n = dataSeries.length;
  const cols = Math.max(1, Math.min(n, multiples.columns ?? (n <= 3 ? n : Math.ceil(Math.sqrt(n)))));
  const rows = Math.ceil(n / cols);
  const gap = 10;
  const titleH = titleHeight(cfg, style);
  const footH = cfg.footnote ? fs * 1.3 : 0;
  const panelW = (cfg.width - gap * (cols - 1)) / cols;
  const panelH = (cfg.height - titleH - footH - gap * (rows - 1)) / rows;
  /**
   * A GRID THAT CANNOT EXIST IS NOT DRAWN, and the alternative was worse than
   * "a bit cramped".
   *
   * The subtraction above goes NEGATIVE easily: ten series in two columns is
   * five rows, and on a 300x60 chart with a title and four 10-point gaps it
   * comes out at −0.4. Handing that on as a panel's `height` looks harmless
   * because something downstream will surely clamp it — and something does.
   * `normalizeConfig` calls `clampDim`, which treats any size `<= 0` as a
   * malformed config and substitutes `DEFAULT_SIZE.height`, 300. So each panel
   * was laid out as a FULL 300-POINT CHART and the composition stacked ten of
   * them 9.6 points apart inside a box 60 points tall: 129 of the scene's 139
   * text nodes below the bottom of the chart, and 2,683 overlapping text pairs
   * in the sweep.
   *
   * Proof it was that and not a squeezed panel: a standalone 155x300 chart and
   * one of these panels both place `category-0` at y = 285.0, to the point.
   *
   * `clampDim` is right and stays as it is — `width: NaN` really does arrive
   * from pasted configs and shape tags written in other decks, and repairing it
   * beats a stack trace in a headless render. What was wrong is that an
   * INTERNAL arithmetic error reached a repair meant for external input, which
   * turned "this cannot be drawn" into a plausible-looking wrong answer.
   *
   * `> 0` and no more. Every positive size passes `clampDim` untouched, so this
   * cannot alter a chart that renders today; it only refuses the case that was
   * being silently rewritten. Declining the grid renders the chart whole, which
   * is what this engine does with every other reservation it cannot pay for.
   */
  if (!(panelW > 0) || !(panelH > 0)) return null;

  const panelCfg = (s: (typeof dataSeries)[number], si: number): ChartConfig => ({
    ...cfg,
    multiples: undefined,
    categorySort: undefined, // already applied above, on the full data
    // Pareto reorders categories too, and each panel holds ONE series — so left
    // on, every panel sorted itself by its own values and the small multiples
    // ended up with different category orders (A,B,C became B,C,A beside A,C,B),
    // which is exactly the contradictory-axes failure categorySort is cleared to
    // avoid. Each panel also silently grew its own cumulative-% line.
    pareto: undefined,
    title: s.name,
    footnote: undefined,
    width: panelW,
    height: panelH,
    data: { ...cfg.data, series: [{ ...s, color: seriesColor(style, si, s.color) }, ...carried] },
    decorations: { ...cfg.decorations, seriesLabels: false },
  });

  // Per-panel extent that also covers the carried Error whiskers (base ± delta)
  // and Target rows — `valueExtent` alone treats an error row as a standalone
  // series, so a whisker or high target could render past the panel edge.
  const panelExtent = (s: (typeof dataSeries)[number], si: number): { min: number; max: number } | null => {
    const pcfg = panelCfg(s, si);
    const { cfg: baseCfg, errors, targets } = extractErrorRows(pcfg);
    const base = valueExtent(baseCfg);
    if (!base) return null;
    let { min, max } = base;
    const vals = baseCfg.data.series[0]?.values ?? [];
    if (errors) {
      vals.forEach((v, c) => {
        if (v == null) return;
        if (errors.plus[c] != null) max = Math.max(max, v + errors.plus[c]!);
        if (errors.minus[c] != null) min = Math.min(min, v - errors.minus[c]!);
      });
    }
    if (targets) {
      for (const t of targets)
        if (t != null) {
          max = Math.max(max, t);
          min = Math.min(min, t);
        }
    }
    return { min, max };
  };

  // Shared scale across panels; radar (no valueExtent) pins 0..global max.
  let scale = cfg.scale;
  if (scale?.min == null || scale?.max == null) {
    const exts = dataSeries.map((s, si) => panelExtent(s, si));
    const global = exts.every((e) => e != null)
      ? { min: Math.min(...exts.map((e) => e!.min)), max: Math.max(...exts.map((e) => e!.max)) }
      : {
          min: 0,
          max: maxOf(
            dataSeries.flatMap((s) => s.values.filter((v): v is number => v != null)),
            1,
          ),
        };
    scale = { min: scale?.min ?? global.min, max: scale?.max ?? global.max };
  }

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  dataSeries.forEach((s, si) => {
    const panel = buildChart({ ...panelCfg(s, si), scale });
    const dx = (si % cols) * (panelW + gap);
    const dy = titleH + Math.floor(si / cols) * (panelH + gap);
    nodes.push(...translateNodes(panel.nodes, dx, dy, `p${si}-`));
  });
  const foot = footnoteNode(cfg, style, String(cfg.footnote ?? ""));
  if (foot) nodes.push(foot);
  return { width: cfg.width, height: cfg.height, nodes: clipTextToFrame(nodes, cfg.width) };
}

/** Human-readable chart-kind names for the accessible description. */
const KIND_LABEL: Record<ChartKind, string> = {
  stacked: "stacked column chart",
  clustered: "clustered column chart",
  stacked100: "100% stacked column chart",
  waterfall: "waterfall chart",
  mekko: "Mekko chart",
  line: "line chart",
  area: "area chart",
  butterfly: "butterfly chart",
  scatter: "scatter plot",
  bubble: "bubble chart",
  gantt: "Gantt chart",
  combo: "combination chart",
  pie: "pie chart",
  doughnut: "doughnut chart",
  boxplot: "box-and-whisker plot",
  radar: "radar chart",
  heatmap: "heatmap",
  tilemap: "tile-grid map",
  cascade: "cascade chart",
  funnel: "funnel chart",
  waffle: "waffle chart",
  treemap: "treemap",
  sunburst: "sunburst chart",
  violin: "violin plot",
  candlestick: "candlestick chart",
};

/**
 * A one-line text alternative for the chart — the accessible description a
 * screen reader reads after the title. Names the kind, the series and the
 * categories (scatter/bubble rows and points read as series/categories, which
 * is what the data model calls them). Deliberately concise; it is a summary,
 * not a data table.
 */
export function describeChart(cfg: ChartConfig): string {
  // hasOwnProperty, not a bare lookup: `kind: "constructor"` otherwise reaches
  // Object.prototype and `??` does not catch it, because a function is not
  // nullish. The chart's accessible description then opened with
  // "function Object() { [native code] }" — which is what a screen reader
  // announces, and what the .pptx carries as the shape's alt text.
  const label = (Object.prototype.hasOwnProperty.call(KIND_LABEL, cfg.kind) && KIND_LABEL[cfg.kind]) || "chart";
  const cats = (cfg.data?.categories ?? []).filter(Boolean);
  const series = cfg.data?.series ?? [];
  const seriesNames = series.map((s) => s.name).filter(Boolean);
  const list = (xs: string[], n = 4) =>
    xs.length <= n ? xs.join(", ") : `${xs.slice(0, n).join(", ")} and ${xs.length - n} more`;
  const bits: string[] = [`${label}${cfg.horizontal ? " (horizontal)" : ""}`];
  if (seriesNames.length)
    bits.push(`${seriesNames.length} data ${seriesNames.length === 1 ? "series" : "series"}: ${list(seriesNames)}`);
  if (cats.length) bits.push(`${cats.length} ${cats.length === 1 ? "category" : "categories"}: ${list(cats)}`);
  return bits.join(". ") + ".";
}

/** Build a renderer-agnostic scene from a chart config. Pure and synchronous. */
export function buildChart(rawCfg: ChartConfig): Scene {
  const cfg0 = normalizeConfig(rawCfg);
  const multiples = buildMultiples(cfg0);
  if (multiples) return { ...multiples, title: cfg0.title, desc: describeChart(cfg0) };
  const extracted = extractErrorRows(sortCategories(applyPareto(applyGanttLanes(cfg0))));
  let cfg = collapseOther(extracted.cfg);
  // Described from the config the chart is DRAWN from, not the one it arrived
  // as. Every transform above changes what a reader sees: `sortCategories` and
  // `applyPareto` reorder the categories, `applyPareto` also turns a clustered
  // chart into a combo and adds the cumulative line, and `collapseOther` buckets
  // the tail into "Other". Described from `cfg0`, a pareto announced itself as a
  // "clustered column chart. 1 data series: V. 4 categories: C0, C1, C2, C3"
  // while drawing a combo with two series and the categories in the order
  // C1, C2, C0, C3 — so the one reader who cannot check it against the picture
  // got the wrong chart, the wrong series count and the wrong order.
  const a11y = { title: cfg.title, desc: describeChart(cfg) };
  const errors = extracted.errors;
  const targets = extracted.targets;
  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const decor: Decorations = { ...DEFAULT_DECOR, ...cfg.decorations };

  // Widen the auto scale so error bars and target ticks stay inside the plot.
  // valueExtent already folds the whiskers and ticks in — it has to, since Same
  // Scale reads it — so this only has to round the result to nice ticks.
  //
  // The `!cfg.horizontal` this carried was the THIRD place that assumption was
  // written down, and the last one to bite: with the marks now drawn on a bar
  // chart, an unwidened axis put a whisker for 34 at x = 530.9 on a 480-wide
  // canvas. The widening was never orientation-specific — it is arithmetic on
  // the value domain — it was only unreachable while nothing drew there.
  if ((errors || targets) && (cfg.scale?.max == null || cfg.scale?.min == null)) {
    const ext = drawnExtent(cfg, errors, targets);
    if (ext) {
      const ticks = niceTicks(ext.min, ext.max, 5);
      cfg = {
        ...cfg,
        scale: { ...cfg.scale, min: cfg.scale?.min ?? ticks[0], max: cfg.scale?.max ?? ticks[ticks.length - 1] },
      };
    }
  }

  let result: LayoutResult;
  switch (cfg.kind) {
    case "waterfall":
      result = layoutWaterfall(cfg, style, decor);
      break;
    case "mekko":
      result = layoutMekko(cfg, style, decor);
      break;
    case "line":
    case "area":
      result = layoutLine(cfg, style, decor);
      break;
    case "butterfly":
      result = layoutButterfly(cfg, style, decor);
      break;
    case "scatter":
    case "bubble":
      result = layoutScatter(cfg, style, decor);
      break;
    case "gantt":
      result = layoutGantt(cfg, style, decor);
      break;
    case "combo":
      result = layoutCombo(cfg, style, decor);
      break;
    case "pie":
    case "doughnut":
      result = layoutPie(cfg, style, decor);
      break;
    case "boxplot":
      result = layoutBoxplot(cfg, style, decor);
      break;
    case "radar":
      result = layoutRadar(cfg, style, decor);
      break;
    case "heatmap":
      result = layoutHeatmap(cfg, style, decor);
      break;
    case "tilemap":
      result = layoutTilemap(cfg, style, decor);
      break;
    case "cascade":
      result = layoutCascade(cfg, style, decor);
      break;
    case "funnel":
      result = layoutFunnel(cfg, style, decor);
      break;
    case "treemap":
      result = layoutTreemap(cfg, style, decor);
      break;
    case "sunburst":
      result = layoutSunburst(cfg, style, decor);
      break;
    case "violin":
      result = layoutViolin(cfg, style, decor);
      break;
    case "candlestick":
      result = layoutCandlestick(cfg, style, decor);
      break;
    case "waffle":
      result = layoutWaffle(cfg, style, decor);
      break;
    default:
      result = layoutColumns(cfg, style, decor);
  }

  // Two different reasons to skip, and they were one flag.
  //
  // Most decorations really do assume a vertical value axis. The Error and
  // Target marks do not — they follow whichever axis carries the value, and the
  // anchors already publish both maps. Folding "this kind has no decorations"
  // together with "this chart is on its side" meant a `Target` row on a bar
  // chart was extracted from the data and then never drawn back: the numbers
  // the user typed vanished from the chart entirely, which is worse than
  // undecorated, because the row would at least have rendered as a bar. That is
  // the documented bullet-chart recipe, and a bullet graph is conventionally
  // horizontal.
  const decorlessKind = [
    "butterfly",
    "scatter",
    "bubble",
    "gantt",
    "pie",
    "doughnut",
    "radar",
    "heatmap",
    "tilemap",
    "cascade",
    "funnel",
    "waffle",
  ].includes(cfg.kind);
  const skipDecor = cfg.horizontal || decorlessKind;
  /** The map from a value to its position, whichever way the value axis runs. */
  const valueToPos = cfg.horizontal ? result.anchors.valueToX : result.anchors.valueToY;
  /** Where a value mark may be drawn: any kind that has them, either orientation. */
  const markValues = !decorlessKind && !!valueToPos;
  // Background bands go BEFORE the layout's nodes so they render behind the
  // data (scatter/bubble draw their own, in value units).
  // Bands are the other half of the bullet recipe and follow the same rule:
  // `bandNodes` reads the orientation itself now, so only the kind gates them.
  const bands = !decorlessKind && decor.bands?.length ? bandNodes(cfg, style, decor, result.anchors) : [];
  const nodes = skipDecor
    ? [...bands, ...result.nodes]
    : [...bands, ...result.nodes, ...decorationNodes(cfg, style, decor, result.anchors)];

  // Error bars from Error / Error+ / Error− rows: a whisker with caps at the
  // column total (or line point), on the shared value scale.
  if (errors && markValues) {
    const a = result.anchors;
    const H = !!cfg.horizontal;
    cfg.data.categories.forEach((_, c) => {
      const plus = errors.plus[c];
      const minus = errors.minus[c];
      if (plus == null && minus == null) return;
      const base = a.columnValue[c];
      // `categoryX` is the category's centre on the CROSS axis — its x on a
      // column chart, its y on a bar. `types.ts` says so where it is declared.
      const mid = a.categoryX[c];
      const capW = Math.min(a.categoryWidth[c] * 0.35, 10);
      const hi = valueToPos!(base + (plus ?? 0));
      const lo = valueToPos!(base - (minus ?? 0));
      /** A line along the VALUE axis, at `mid` on the cross axis. */
      const along = (v1: number, v2: number) =>
        H ? { x1: v1, y1: mid, x2: v2, y2: mid } : { x1: mid, y1: v1, x2: mid, y2: v2 };
      /** A cap ACROSS the value axis, at value `v`. */
      const across = (v: number) =>
        H
          ? { x1: v, y1: mid - capW / 2, x2: v, y2: mid + capW / 2 }
          : { x1: mid - capW / 2, y1: v, x2: mid + capW / 2, y2: v };
      nodes.push({ kind: "line", ...along(hi, lo), stroke: style.axis, strokeWidth: 1, name: `error-${c}` });
      if (plus != null)
        nodes.push({ kind: "line", ...across(hi), stroke: style.axis, strokeWidth: 1, name: `error-cap-hi-${c}` });
      if (minus != null)
        nodes.push({ kind: "line", ...across(lo), stroke: style.axis, strokeWidth: 1, name: `error-cap-lo-${c}` });
    });
  }

  // Bullet-chart target ticks: a bold marker across each column at the
  // target value (combine with decorations.bands for the range zones).
  if (targets && markValues) {
    const a = result.anchors;
    const H = !!cfg.horizontal;
    cfg.data.categories.forEach((_, c) => {
      const t = targets[c];
      if (t == null) return;
      const mid = a.categoryX[c];
      const half = Math.min(a.categoryWidth[c] * 0.62, 26);
      const v = valueToPos!(t);
      // ACROSS the bar at the target value: a horizontal rule on a column
      // chart, a vertical one on a bar. Same mark, same meaning, turned.
      const tick = H
        ? { x1: v, y1: mid - half, x2: v, y2: mid + half }
        : { x1: mid - half, y1: v, x2: mid + half, y2: v };
      nodes.push({ kind: "line", ...tick, stroke: style.text, strokeWidth: 2.25, name: `target-${c}` });
      // Budget-vs-actual bridge: on a waterfall, a hatched gap segment shows
      // the distance from the achieved level to the target.
      if (cfg.kind === "waterfall") {
        const actual = a.columnValue[c];
        const gap = t - actual;
        if (Math.abs(gap) > 1e-9) {
          const vA = valueToPos!(actual);
          const thick = a.categoryWidth[c];
          const lo = Math.min(v, vA);
          const span = Math.abs(vA - v);
          // The hatched gap runs ALONG the value axis and is as thick as the
          // bar, whichever way round that is.
          const box = H
            ? { x: lo, y: mid - thick / 2, w: span, h: thick }
            : { x: mid - thick / 2, y: lo, w: thick, h: span };
          nodes.push({
            kind: "rect",
            ...box,
            fill: style.neutral,
            pattern: "diagonal",
            stroke: style.mutedText,
            strokeWidth: 0.75,
            name: `target-gap-${c}`,
          });
          const fs = style.fontSize;
          // Just clear of the bar on the cross axis, centred on the gap along
          // the value axis — the same relationship in both orientations.
          const labelBox = H
            ? { x: lo - 4, y: mid - thick / 2 - fs * 1.5, w: span + 8, h: fs * 1.4 }
            : { x: mid - thick / 2 - 4, y: lo - fs * 1.5, w: thick + 8, h: fs * 1.4 };
          nodes.push({
            kind: "text",
            ...labelBox,
            text: `Gap ${formatNumber(gap, { ...resolveFormat([t, actual], cfg.numberFormat), forceSign: true })}`,
            fontSize: fs * 0.95,
            bold: true,
            color: style.text,
            align: "center",
            valign: "bottom",
            name: `target-gap-label-${c}`,
          });
        }
      }
    });
  }

  // Footnote line: source citation and/or the "100% = N" note, bottom-left.
  const footParts: string[] = [];
  if (decor.hundredPercentNote) {
    const total = hundredPercentTotal(cfg);
    if (total != null) footParts.push(`100% = ${formatNumber(total, resolveFormat([total], cfg.numberFormat))}`);
  }
  if (cfg.footnote) footParts.push(cfg.footnote);
  // Overlap relief is an approximation, so the chart discloses it — quoting
  // the cap the layout actually enforces, not a restatement of it.
  const sc = spreadCap(cfg);
  if (sc) {
    footParts.push(
      `${sc.axis.toUpperCase()} positions approximate: markers spread by up to ±${formatNumber(sc.limit, resolveFormat([sc.limit], cfg.numberFormat))} to reduce overlap`,
    );
  }
  if (footParts.length) {
    const foot = footnoteNode(cfg, style, footParts.join("   ·   "));
    if (foot) nodes.push(foot);
  }

  // Global de-collision for outside labels (vertical cartesian charts).
  if (!skipDecor) resolveLabelCollisions(nodes, cfg.height);

  // A combo's totals and its line's point labels can both want one band on a
  // short frame, where neither can move (the flip below the point lands off the
  // bottom of the canvas). Whichever `tightLabelPriority` does not favour is
  // dropped — but only where it is STILL colliding after every legal move above,
  // so a roomy chart keeps both.
  // Sideways too, which `skipDecor` used to exclude along with the rest of the
  // decoration stage. This pass adds nothing to a scene — it only removes labels
  // that are still unreadable after every legal move — and a horizontal combo
  // needs it as much as an upright one: its line name sits at the line's last
  // point, in the same few points of canvas the last point labels want.
  {
    const drop = unplaceableComboLabels(nodes, decor.tightLabelPriority ?? "columns");
    // Spliced rather than reassigned: `nodes` is the scene array the whole
    // builder has been appending to, and later passes hold the same reference.
    if (drop.size) for (let i = nodes.length - 1; i >= 0; i--) if (drop.has(nodes[i])) nodes.splice(i, 1);
  }

  // The other half of a trade this engine had already decided but not carried
  // out: a scatter or bubble point label is allowed into the strip where the
  // tick numbers live, because a point's label is data and a tick number is
  // chrome. Both were drawn anyway, through each other. The chrome yields now.
  //
  // After de-collision and after the combo drop, for the same reason those run
  // in that order: only a tick number still covered once every legal move has
  // been made is worth removing.
  {
    const drop = tickLabelsUnderPointLabels(nodes);
    if (drop.size) for (let i = nodes.length - 1; i >= 0; i--) if (drop.has(nodes[i])) nodes.splice(i, 1);
  }

  // Manual label nudges (think-cell's label dragging, config-driven).
  if (cfg.labelOffsets) {
    for (const n of nodes) {
      const off = n.name && cfg.labelOffsets[n.name];
      if (off && n.kind === "text") {
        n.x += off.dx;
        n.y += off.dy;
      }
    }
  }
  // Last gate before the scene leaves the engine — see `finiteNodes`. The
  // PowerPoint renderers write a file, and a NaN coordinate makes it one
  // PowerPoint may refuse to open.
  return { width: cfg.width, height: cfg.height, nodes: clipTextToFrame(finiteNodes(nodes), cfg.width), ...a11y };
}

/** The data's own range, ignoring the anatomy drawn on top of it. */
function dataExtent(cfg: ChartConfig): { min: number; max: number } | null {
  const { data, kind } = cfg;
  const cats = data.categories.map((_, c) => c);
  const vals = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  if (!vals.length) return null;
  switch (kind) {
    case "stacked": {
      const pos = cats.map((c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
      const neg = cats.map((c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
      return { min: minOf(neg, 0), max: maxOf(pos, 0) };
    }
    case "clustered":
    case "line":
      return { min: minOf(vals, 0), max: maxOf(vals, 0) };
    case "boxplot":
      return boxplotExtent(cfg);
    case "area": {
      // Areas stack under the baseline too (layout/line.ts dips negatives below
      // zero), so the extent must mirror "stacked" — a hard 0 floor here clipped
      // every negative area on the shared-scale paths (small multiples, Same Scale).
      const pos = cats.map((c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
      const neg = cats.map((c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
      return { min: minOf(neg, 0), max: maxOf(pos, 0) };
    }
    case "waterfall": {
      // One chain, shared with the layout. This used to be a second walk that
      // added only series[0] and skipped spacerIndices, so a stacked bridge
      // reported the first series' total as the whole chart's — 63 instead of
      // 114 — and Same scale then clipped the bars off the shape.
      return waterfallExtent(cfg);
    }
    default:
      return null;
  }
}

/**
 * Value-axis extent of a chart (for think-cell's Same Scale): the range the
 * auto scale would cover. Null for charts without a value axis (100%, Mekko,
 * butterfly, scatter, gantt).
 *
 * This reports what the layout DRAWS, not what the datasheet holds. Error and
 * Target rows are chart anatomy rather than data — counting an Error row's
 * magnitude as a data point gave a range that was neither — so they are pulled
 * out first and then folded back in as the whiskers and ticks they become.
 * Same Scale writes this straight back as a hard `scale` override, which
 * suppresses the auto-widen below, so anything this under-reports renders off
 * the shape.
 */
export function valueExtent(cfg: ChartConfig): { min: number; max: number } | null {
  // The SAME normalisation `buildChart` runs before it lays anything out, in the
  // same order. Anything this skips is measured on data the chart does not draw.
  //
  // `collapseOther` is the one that changes VALUES rather than order, and it was
  // missing: on a clustered chart the synthesized "Other" series is the SUM of
  // the collapsed tail, which is taller than any single series `dataExtent` can
  // see. Same Scale writes this answer straight back as a hard `cfg.scale`,
  // which suppresses every auto-widen downstream — precisely the failure this
  // function's own docstring names. Measured: an `otherBucket:{max:3}` chart
  // whose Other bar reaches 120 reported max 50, and Same Scale then drew that
  // bar at y = -375.6 with h = 657.6 on a 300pt canvas.
  //
  // `normalizeConfig` was the OTHER one missing, and it is the first thing
  // `buildChart` does. It is what guarantees `Series.name` is a string and drops
  // a null series entry, and the three passes below all call `s.name.trim()`
  // unguarded — so a series written as `{ values: [...] }` with no name, which
  // the skill will happily write into a POWERCHART_CONFIG tag, rendered fine
  // through `buildChart` and threw a TypeError here. `doSameScale` maps this
  // over the deck without a guard, so one such chart aborted the whole
  // Same-Scale operation and nothing was rescaled.
  const extracted = extractErrorRows(sortCategories(applyPareto(applyGanttLanes(normalizeConfig(cfg)))));
  return drawnExtent(collapseOther(extracted.cfg), extracted.errors, extracted.targets);
}

/**
 * valueExtent for a config whose carried rows are already extracted — the state
 * buildChart is in by the time it needs the scale. Re-extracting there would
 * find nothing left to extract and silently widen by zero.
 */
function drawnExtent(
  clean: ChartConfig,
  errors: { plus: (number | null)[]; minus: (number | null)[] } | null,
  targets: (number | null)[] | null,
): { min: number; max: number } | null {
  const base = dataExtent(clean);
  return base ? widenForAnatomy(base, clean, errors, targets) : null;
}

/**
 * Grow an extent to cover the decorations drawn beyond the data: error whiskers,
 * target ticks and threshold lines. Deliberately conservative — whiskers are
 * anchored per category, but taking the largest delta against the extreme value
 * can only over-reach, and over-reaching costs a little white space where
 * under-reaching pushes ink off the shape.
 */
function widenForAnatomy(
  base: { min: number; max: number },
  cfg: ChartConfig,
  errors: { plus: (number | null)[]; minus: (number | null)[] } | null,
  targets: (number | null)[] | null,
): { min: number; max: number } {
  const nums = (xs: (number | null)[] | null | undefined) => (xs ?? []).filter((v): v is number => v != null);
  const maxPlus = Math.max(0, ...nums(errors?.plus));
  const maxMinus = Math.max(0, ...nums(errors?.minus));
  const targetVals = nums(targets);
  // Same normalization decor.ts applies to the legacy single valueLine. Only an
  // explicit threshold can sit outside the data — a "mean" line never can.
  const decor = cfg.decorations;
  const lineVals = (decor?.valueLines ?? (decor?.valueLine ? [decor.valueLine] : []))
    .filter((l): l is { mode: "value"; value: number } => l.mode === "value")
    .map((l) => l.value);
  return {
    min: Math.min(base.min - maxMinus, 0, ...targetVals, ...lineVals),
    max: Math.max(base.max + maxPlus, 0, ...targetVals, ...lineVals),
  };
}

export { layoutColumns, layoutWaterfall, layoutMekko, layoutLine };
