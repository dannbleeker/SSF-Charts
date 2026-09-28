import type { NumberFormat } from "./types";
import { maxOf } from "./agg";

export const DEFAULT_FORMAT: NumberFormat = { decimals: "auto" };

/**
 * Coerce a decimals count into the range `toFixed`/`Intl.NumberFormat` accept
 * (0–100). An authored or hand-edited `numberFormat.decimals` of -1 or 500
 * would otherwise throw a RangeError out of `toFixed`, aborting the whole
 * render — the same class of bad input the locale try/catch already repairs.
 */
const safeDecimals = (d: number): number => (Number.isFinite(d) ? Math.min(100, Math.max(0, Math.trunc(d))) : 0);

/**
 * Intl.NumberFormat instances are expensive to construct but immutable and
 * reusable, and a chart formats hundreds of labels sharing a handful of
 * (locale, decimals, style) triples. `Number.prototype.toLocaleString(locale,
 * opts)` is specified to construct a fresh NumberFormat on every call, so caching
 * by that triple and reusing `.format()` is byte-identical output at a fraction
 * of the cost.
 */
const NUMBER_FORMATTERS = new Map<string, Intl.NumberFormat>();
function numberFormatter(locale: string, decimals: number, style?: "percent"): Intl.NumberFormat {
  const key = `${locale} ${decimals} ${style ?? ""}`;
  let nf = NUMBER_FORMATTERS.get(key);
  if (!nf) {
    const opts: Intl.NumberFormatOptions = {
      style,
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    };
    // A malformed BCP-47 tag (an authored config, a hand-edited shape tag)
    // makes the Intl constructor throw a RangeError — fall back to en-US
    // rather than let one bad locale abort the whole render.
    try {
      nf = new Intl.NumberFormat(locale, opts);
    } catch {
      nf = new Intl.NumberFormat("en-US", opts);
    }
    NUMBER_FORMATTERS.set(key, nf);
  }
  return nf;
}

/** Format a value the way think-cell's default label format does: compact, thousands-separated. */
export function formatNumber(v: number, fmt: Partial<NumberFormat> = {}): string {
  // A non-finite value (a divide-by-zero, an empty average, a stray Infinity in
  // authored data) would otherwise print the literal "NaN"/"Infinity" as chart
  // text — suppress the label instead of drawing a broken one.
  if (!Number.isFinite(v)) return "";
  const f = { ...DEFAULT_FORMAT, ...fmt };
  const abs = Math.abs(v);
  // Exact zero takes no fractional digits — "0", not "0.0". (The old test
  // `abs !== 0 && abs < 1` fell through to the `< 10 → 1 decimal` branch for 0.)
  const decimals = f.decimals === "auto" ? (abs === 0 ? 0 : abs < 1 ? 2 : abs < 10 ? 1 : 0) : safeDecimals(f.decimals);
  // A small negative that rounds toward zero would print as "-0". Normalise the
  // VALUE, not the formatted string: Intl renders the sign as U+2212 in some
  // locales, prefixes an invisible directional mark in RTL ones, and uses
  // non-Latin digits in others — no pattern match on the output survives all
  // three, and -0 leaked through in every locale that does any of them.
  if (Number(v.toFixed(decimals)) === 0) v = 0;
  let s = numberFormatter(f.locale ?? "en-US", decimals).format(v);
  if (f.forceSign && v > 0) s = "+" + s;
  if (f.suffix) s += f.suffix;
  return s;
}

/**
 * Resolve "auto" decimals once per chart from the data's magnitude, so all
 * labels in one chart share the same precision (as think-cell does).
 */
export function resolveFormat(values: number[], format?: Partial<NumberFormat> | null): NumberFormat {
  // A default parameter only fires for `undefined`. `numberFormat: null` is
  // ordinary JSON — it is what a serialiser writes for an absent field — and it
  // sailed past the default and threw on `.decimals`. Coerced here rather than
  // at each of the call sites, which is where it would be forgotten.
  const fmt = asFormat(format);
  if (fmt.decimals != null && fmt.decimals !== "auto") {
    return { ...DEFAULT_FORMAT, ...fmt, decimals: fmt.decimals };
  }
  const maxAbs = maxOf(values.filter((v) => Number.isFinite(v)).map(Math.abs), 0);
  // Below 1, keep widening until the largest value shows two significant digits.
  //
  // The ladder used to stop at 2 decimals however small the data got, so a chart
  // of rates, yields, defect fractions or probabilities — everything that lives
  // under 1 — printed "0.00" on every label for values that plainly differ. The
  // chart then contradicted itself out loud: `resolveAxisFormat` widens on the
  // TICK STEP, so the same chart's axis read 0.000 / 0.001 / 0.002 beside bars
  // all labelled "0.00".
  //
  // Capped at six: six resolves to a millionth, and past that the label is
  // longer than the bar it sits on.
  let decimals = maxAbs >= 10 ? 0 : maxAbs >= 1 ? 1 : 2;
  if (maxAbs > 0 && maxAbs < 1) {
    while (decimals < 6 && maxAbs * Math.pow(10, decimals) < 10) decimals++;
  }
  return { ...DEFAULT_FORMAT, ...fmt, decimals };
}

/** Whatever arrived where a number format belongs, as something spreadable. */
function asFormat(fmt: Partial<NumberFormat> | null | undefined): Partial<NumberFormat> {
  return fmt && typeof fmt === "object" ? fmt : {};
}

/**
 * The decimals count the AUTHOR wrote, or undefined when the chart is on "auto".
 *
 * `resolveFormat`'s answer cannot tell those two apart — it always returns a
 * number — so anything that must obey an explicit setting without inventing one
 * asks here. `segmentLabel`'s percent part is the case that needed it.
 */
export function authoredDecimals(fmt?: Partial<NumberFormat> | null): number | undefined {
  const d = asFormat(fmt).decimals;
  return typeof d === "number" ? safeDecimals(d) : undefined;
}

/**
 * Same, for an AXIS's tick labels, where magnitude alone is not enough.
 *
 * A tick strip must let the reader tell one tick from the next: an axis over
 * 7.444–7.471 has ticks 0.01 apart, but its magnitude (≥1) buys one decimal, so
 * five gridlines printed ["7.4","7.5","7.5","7.5","7.5"] — two distinct labels,
 * and a top tick named as a value outside the scale. Take the finer of the
 * magnitude precision and the precision the tick STEP needs, so no two ticks
 * share a label and every label names its own tick. Data labels keep the plain
 * magnitude rule: they are read one at a time, not against each other.
 */
export function resolveAxisFormat(ticks: number[], format?: Partial<NumberFormat> | null): NumberFormat {
  const fmt = asFormat(format);
  const resolved = resolveFormat(ticks, fmt);
  // An explicit `decimals` is the author's call — never widen it.
  if (fmt.decimals != null && fmt.decimals !== "auto") return resolved;
  const sorted = ticks.filter((t) => Number.isFinite(t)).sort((a, b) => a - b);
  let step = Infinity;
  for (let i = 1; i < sorted.length; i++) {
    const gap = sorted[i] - sorted[i - 1];
    if (gap > 0) step = Math.min(step, gap);
  }
  if (!Number.isFinite(step)) return resolved;
  // The fewest decimals that render EVERY tick exactly — the contract above is
  // "every label names its own tick", and only the ticks themselves can settle
  // that. Deriving it from log10(step) assumed the 1/2/5×10^k steps niceTicks
  // emits; a hand-built tick list (the 100% axis's 0/.25/.5/.75/1) breaks that
  // assumption, and 0.25 came back as "0.3" — a label naming no tick on the axis.
  // "Exactly" is only meaningful against the STEP: ticks 1e-7 apart need seven
  // decimals to stay distinct, ticks 1/3 apart need four. A fixed absolute
  // epsilon answered neither — it ran to the ceiling on a small-magnitude axis
  // (which then printed one identical label on every gridline, the exact
  // contract this function exists to keep) and spent seventeen decimals on a
  // non-terminating tick list.
  const tol = step * 1e-3;
  let stepDecimals = 0;
  while (stepDecimals < 20 && !sorted.every((t) => Math.abs(t - Number(t.toFixed(stepDecimals))) < tol)) {
    stepDecimals++;
  }
  const decimals = typeof resolved.decimals === "number" ? resolved.decimals : 0;
  return { ...resolved, decimals: Math.max(decimals, stepDecimals) };
}

/**
 * The label function for a value axis's tick strip.
 *
 * The one place a value-axis tick becomes text, shared by the vertical chrome
 * and the horizontal (bar) chrome — they had drifted twice over. Only the
 * vertical one used `resolveAxisFormat`, so a narrow horizontal axis printed
 * the duplicate labels that function exists to prevent; and the share branch's
 * suffix fix had to be made in both places.
 *
 * A share axis is labelled in percent: the ticks are fractions and the segments
 * beside them already read "60%", so precision comes from the SCALED ticks and
 * each label still names its own tick. A share is also unitless — `formatNumber`
 * appends `numberFormat.suffix` (the documented way to say "millions"), which
 * labelled a 100% axis "25 m%".
 */
export function axisTickLabel(
  ticks: number[],
  percent: boolean | undefined,
  fmt: Partial<NumberFormat> = {},
): (t: number) => string {
  const axisFmt = resolveAxisFormat(percent ? ticks.map((t) => t * 100) : ticks, fmt);
  const shareFmt = { ...axisFmt, suffix: undefined };
  return (t: number) => (percent ? `${formatNumber(t * 100, shareFmt)}%` : formatNumber(t, axisFmt));
}

/**
 * A share (0.358 → "35.8%") in the chart's own locale.
 *
 * Percent labels sit next to `formatNumber` ones — the funnel's conversion rate
 * beside its stage value, the CAGR arrow beside the value axis — so they must
 * share the locale, or one chart prints two number systems ("12.000" beside
 * "35.8%"). Routing through the same cached Intl formatter also grouped the
 * thousands a hand-rolled `toFixed` never did ("3550%" → "3,550%").
 */
export function formatPercent(v: number, decimals = 0, forceSign = false, locale?: string): string {
  if (!Number.isFinite(v)) return "";
  decimals = safeDecimals(decimals);
  // A small negative that rounds toward zero would print as "-0%". Normalise the
  // VALUE, not the formatted string — the same reason formatNumber does: the
  // sign glyph, digits and percent spacing are all locale-dependent, so no
  // pattern match on the output survives every locale.
  if (Number((v * 100).toFixed(decimals)) === 0) v = 0;
  const s = numberFormatter(locale ?? "en-US", decimals, "percent").format(v);
  return forceSign && v > 0 ? "+" + s : s;
}

/**
 * "Nice" axis ticks covering [min, max] with roughly `count` steps.
 * Returns the tick values including the padded ends.
 */
/** More gridlines than any axis can show. A ceiling, not a target. */
const MAX_TICKS = 1000;

export function niceTicks(min: number, max: number, count = 5): number[] {
  // A non-finite bound (from a NaN in the data extent) or an inverted range
  // (a reversed manual scale.min>max) makes `Math.log10(span)` NaN and turns
  // every tick into NaN — invisible geometry for the whole value axis. Repair
  // to a usable range instead of propagating the poison.
  if (!Number.isFinite(min) || !Number.isFinite(max)) return [0, 1];
  if (min > max) [min, max] = [max, min];
  if (min === max) {
    if (min === 0) return [0, 1];
    min = Math.min(0, min);
    max = Math.max(0, max);
    if (min === max) max = min + 1;
  }
  const span = max - min;
  const rawStep = span / Math.max(1, count);
  const mag = Math.pow(10, Math.floor(Math.log10(rawStep)));
  const norm = rawStep / mag;
  const step = (norm >= 5 ? 10 : norm >= 2 ? 5 : norm >= 1 ? 2 : 1) * mag;
  const lo = Math.floor(min / step) * step;
  const hi = Math.ceil(max / step) * step;
  // Rounding OUTWARDS can overflow where the raw bound did not, and a span
  // small enough to be subnormal collapses `step` to 0. Both end the same way:
  // `(hi - lo) / step` is Infinity or NaN, and the loop below has no end while
  // the array it fills has no limit. `niceTicks(0, 1.7e308)` never returned —
  // it filled memory until the tab died, from a number someone can type.
  //
  // The entry guard checks the INPUTS are finite. These are outputs, and a
  // finite input does not promise a finite bound: ceil(1.7e308 / 5e307) is 4,
  // and 4 x 5e307 is Infinity.
  const span_ = Math.round((hi - lo) / step);
  if (!Number.isFinite(lo) || !Number.isFinite(hi) || !Number.isFinite(span_) || span_ < 0) return [min, max];
  const ticks: number[] = [];
  // Clean the accumulated FP drift at the STEP's own precision, not at a fixed
  // 12 significant digits: a step of 1 around 1e13 needs 14 of them, and
  // `toPrecision(12)` collapsed every tick of such an axis onto one value —
  // gridlines stacked on top of each other and a top tick below the data max.
  // (Steps are always 1/2/5 × 10ᵏ, so this rounds to the tick grid exactly.)
  const stepDecimals = Math.min(100, Math.max(0, -Math.floor(Math.log10(step) + 1e-9)));
  // Guard against FP drift producing an extra/short tick. Capped as well: a
  // caller asking for a huge `count` is asking for a chart nobody can read,
  // and an axis is not the place to find that out by running out of memory.
  const last = Math.min(span_, MAX_TICKS);
  for (let i = 0; i <= last; i++) {
    ticks.push(+(lo + i * step).toFixed(stepDecimals));
  }
  return ticks;
}

/**
 * Segment label text from think-cell's label-content dropdown: any ordered
 * combination of value, percent (of column), series and category names.
 */
export function segmentLabel(
  parts: ("value" | "percent" | "series" | "category")[],
  ctx: {
    value: number;
    fraction: number | null;
    series: string;
    category: string;
    fmt: Partial<NumberFormat>;
    /**
     * Decimals for the PERCENT part, when the author named one.
     *
     * It cannot be read off `ctx.fmt`: that is a RESOLVED format, so its
     * `decimals` is always a number and always derived from the VALUES'
     * magnitude — a pie of 0-to-10 numbers would silently start printing
     * "29.0%" because its values happen to want one decimal. So the caller
     * passes the authored count (see `authoredDecimals`) and nothing else.
     *
     * Why it matters: a 100% chart with `numberFormat.decimals: 1` labelled its
     * axis 0.0% / 25.0% / … / 100.0% — `resolveAxisFormat` honours an explicit
     * count — and its three equal segments "33%", "33%", "33%". The author
     * asked for the precision that makes the parts add up, one of the two
     * places that consume it obeyed, and the chart contradicted itself: 99% of
     * a column drawn full height under an axis that reads 100.0%.
     */
    percentDecimals?: number;
  },
): string {
  // `decorations.labelContent` is a LIST, and a config that wrote a bare
  // `"value"` instead of `["value"]` — an easy thing to write by hand or to
  // generate — threw `parts.map is not a function`. A single part is what was
  // meant, so treat it as a list of one; anything else contributes nothing,
  // which is what an empty list already does.
  const list = Array.isArray(parts) ? parts : typeof parts === "string" ? [parts] : [];
  return list
    .map((p) => {
      switch (p) {
        case "value":
          return formatNumber(ctx.value, ctx.fmt);
        case "percent":
          return ctx.fraction == null
            ? null
            : formatPercent(ctx.fraction, ctx.percentDecimals ?? 0, false, ctx.fmt.locale);
        case "series":
          return ctx.series;
        case "category":
          return ctx.category;
      }
    })
    .filter(Boolean)
    .join(p2sep(parts));
}

/** Multi-part labels read best on one line for two parts, else spaced. */
function p2sep(parts: string[]): string {
  return parts.length > 1 ? " " : "";
}

const DAY_MS = 86400000;

/** The only words a date token may contain — see `parseDateToken`. */
const DATE_WORDS = new Set([
  ..."jan feb mar apr may jun jul aug sep oct nov dec".split(" "),
  ..."january february march april june july august september october november december".split(" "),
  ..."mon tue tues wed thu thur thurs fri sat sun".split(" "),
  ..."monday tuesday wednesday thursday friday saturday sunday".split(" "),
  // Date-time markers Date.parse understands: the ISO separator/zone letters,
  // named zones, and the meridiem.
  ..."t z utc gmt am pm".split(" "),
]);

/**
 * Parse a calendar-date cell ("2026-01-15", "15.01.2026", "Jan 2026", …)
 * into days since the Unix epoch. Returns null for non-dates.
 */
/**
 * Gantt rows whose values are calendar days rather than plain numbers.
 *
 * Lives here, beside `parseDateToken`, because two callers need the same answer:
 * the datasheet, which renders epoch days back as ISO on the way out, and
 * `normalizeConfig`, which reads ISO in on the way IN. It was only in the
 * datasheet, so a config authored anywhere else — the skill, a hand-written
 * JSON, a POWERCHART_CONFIG tag — lost its dates in silence.
 */
export const GANTT_DATE_ROW = /^(?:start|end|milestone|today|holidays?|baseline\s*(?:start|end))$|^bracket\b/i;

export function parseDateToken(raw: string): number | null {
  // Coerced, like every other text boundary in this engine. `raw` is a CELL —
  // it comes from a pasted block, a JSON config or the skill's caller — and
  // `null`/`undefined` threw `Cannot read properties of null (reading 'trim')`
  // while a number threw `raw.trim is not a function`. A number is the case
  // that matters: this function's whole job is to decide whether a cell is a
  // date, a bare number is explicitly NOT one (see the guard below), and it
  // could not reach that answer without crashing first.
  const t = String(raw ?? "").trim();
  if (!t) return null;
  const dmy = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!dmy && /^[-+]?[\d,.]+$/.test(t)) return null; // plain numbers are not dates
  // A percentage is never a date. Without this `Date.parse("50% UTC")` yields a
  // finite garbage instant, so a perfectly ordinary "50%" cell became epoch day
  // -7305 AND flipped the whole chart into date mode.
  if (/%/.test(t)) return null;
  // Numeric ranges ("3-5", "10–20") are category labels, not dates — Date.parse
  // would otherwise misread them as partial ISO dates.
  if (/^\d{1,3}\s*[-–]\s*\d{1,3}$/.test(t)) return null;
  // A slash date whose first two numbers could BOTH be a month is two different
  // dates and this parser cannot know which. `Date.parse` picks the American
  // one silently, and the dotted form ten lines up picks the European one — so
  // the same digits, in one datasheet, meant dates two months apart:
  //
  //     03/01/2026  ->  1 March        (Date.parse, month first)
  //     03.01.2026  ->  3 January      (the `dmy` form above, day first)
  //
  // On a Gantt that is not a cosmetic difference. A European programme plan of
  // `01/02, 05/02, 10/03` came back as 2 January, 2 May and 3 October: three
  // months spread over ten, with nothing said.
  //
  // Refused rather than guessed, which is the rule `numericValue` in the
  // datasheet already applies to the mirror-image case — a European "1,5"
  // against an American "1,500" — in the same words: a visible gap beats a
  // wrong number. An unambiguous slash date is untouched, so `01/15/2026` still
  // reads as 15 January (no other reading exists) and `03/03/2026` still reads
  // as 3 March (both readings agree).
  //
  // CONFIRMED BY THE OWNER (2026-08-12) with the cost stated: a US author who
  // writes `03/01/2026` for 1 March gets an empty cell rather than a chart.
  // That is the trade, and it was taken deliberately — this is not a default
  // waiting to be tuned. The two alternatives were offered and declined: read
  // slash dates day-first to match the dotted form, or disambiguate from
  // `style.locale` when a config declares one. Both pick a winner where the
  // data does not. Don't reopen it without him.
  const slash = /^(\d{1,2})\/(\d{1,2})\/\d{2,4}$/.exec(t);
  if (slash && Number(slash[1]) <= 12 && Number(slash[2]) <= 12 && slash[1] !== slash[2]) return null;
  // Everything below reaches `Date.parse`, which is far more lenient than a date
  // cell has any right to be: "#DIV/0! UTC" comes back as 2000-01-01, "Store 5"
  // as 2001-05-01, "<0.1" as 2000-01-01. So an Excel error cell, a shop name or
  // a threshold became epoch day 10957 AND flipped the whole ChartData into date
  // mode. Two shape rules first — a date is spelled with digits and date
  // punctuation only, and every word in it names a month, a weekday, or a
  // timezone/meridiem marker.
  if (/[^A-Za-z0-9 ,./:+\-–]/.test(t)) return null;
  for (const w of t.match(/[A-Za-z]+/g) ?? []) {
    if (!DATE_WORDS.has(w.toLowerCase())) return null;
  }
  // `Jan-24` is Excel's `mmm-yy`, the commonest monthly category label there
  // is — and `Date.parse` reads the 24 as a DAY, in whatever year it defaults
  // to (2001 in V8). Relative spacing survives that inside one year, so it
  // looked fine; across a year boundary it does not. `Oct-23, Nov-23, Dec-23,
  // Jan-24, Feb-24` gave January an epoch day 333 LOWER than December's, and
  // the line chart drew its last two months at the far left of the plot, before
  // October — data silently plotted in the wrong place, which is worse than
  // refusing the label.
  //
  // Only the hyphen/slash forms are claimed here. `Jan 24` with a space is
  // genuinely ambiguous with "January 24" and is left to `Date.parse`; nobody
  // writes a day-of-month as "Jan-24".
  //
  // Two-digit years pivot at 30, which is Excel's own rule for the cells these
  // labels are formatted from.
  const mmmYY = /^([A-Za-z]{3,9})[-/](\d{2})$/.exec(t);
  if (mmmYY) {
    const m = MONTH_INDEX.get(mmmYY[1].slice(0, 3).toLowerCase());
    if (m != null) {
      const yy = Number(mmmYY[2]);
      return Math.floor(Date.UTC(yy < 30 ? 2000 + yy : 1900 + yy, m, 1) / DAY_MS);
    }
  }
  /**
   * An ISO token — a bare date, or a date-time with any zone or none — read for
   * the CALENDAR DATE it names, rather than handed to `Date.parse`.
   *
   * `Date.parse` reads a bare ISO date as UTC and an offset-less ISO date-TIME
   * as LOCAL, which is the one inconsistency this parser was still passing
   * through. So `2026-01-15T20:00` became 2026-01-16T01:00Z for a reader in New
   * York and `2026-01-15T02:00` became 2026-01-14T17:00Z for one in Tokyo — and
   * the day check below then compared the 15 the token NAMES against the 16 or
   * the 14 the instant landed on, decided the day did not exist, and returned
   * null. The cell was not merely shifted: it silently stopped being a date,
   * taking its Gantt row or its category spacing with it, and WHICH cells did
   * that depended on the reader's timezone and the time of day in the cell. A
   * pasted export carrying local timestamps is the ordinary way in.
   *
   * An explicit offset had the same ending by a different route: a legal
   * `2026-01-15T20:00-05:00` is 16 January in UTC, so the day check refused it
   * in every timezone on earth.
   *
   * The date a cell NAMES is the answer to both, and it is the honest one: this
   * function returns whole days, a plan written as "15 Jan, 20:00" means the
   * 15th to whoever wrote it, and the same config has to render the same chart
   * on every machine. The time of day was already discarded by the floor at the
   * end; only the day it was allowed to move was ever in question.
   */
  const iso = /^(\d{4})-(\d{2})(?:-(\d{2}))?(?:[T ][\d:.]+(?:[Zz]|[+-]\d{2}:?\d{2})?)?$/.exec(t);
  // A month outside 1..12 is not a date, and `Date.UTC` normalises rather than
  // refuses — so `2026-13-01` becomes 1 January 2027 and `15.13.2026` becomes
  // 15 January 2027: a year out, silently, on a Gantt row. This is the month's
  // half of the day-existence rule below, which has always refused `Apr 31`;
  // the dotted form has been able to roll a year since it was written, and
  // reading the ISO form ourselves would newly let it do the same (`Date.parse`
  // simply returns NaN for a 13th month).
  const month = dmy ? Number(dmy[2]) : iso ? Number(iso[2]) : null;
  if (month != null && (month < 1 || month > 12)) return null;
  const ms = dmy
    ? Date.UTC(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]))
    : iso
      ? Date.UTC(Number(iso[1]), Number(iso[2]) - 1, iso[3] ? Number(iso[3]) : 1)
      : // Everything else ("Jan 2026", "Mon 5 Jan") still needs the UTC anchor
        // to avoid the same local-timezone drift.
        Date.parse(`${t} UTC`);
  if (!Number.isFinite(ms)) return null;
  // A day that does not EXIST in its month is not a date, and until now it
  // silently became one in the next month.
  //
  // `Date.UTC` and `Date.parse` both normalise rather than refuse, so
  // `Feb 29 2023` — a plausible mistake about a leap year — came back as
  // 1 March 2023, and `Apr 31` as 1 May. On a Gantt row that is a task
  // silently starting a month late, with nothing to say so. The parser was
  // already inconsistent about it: `Jan 32 2024` is refused, because 32 is not
  // a day number at all, while `Apr 31` was accepted and moved.
  //
  // Checked only where the token names a day UNAMBIGUOUSLY — the ISO and
  // dotted forms, and a month-word form carrying exactly one one-or-two-digit
  // number. Anything looser is left alone rather than guessed at, which is the
  // same rule the shape gates above follow.
  const asked = dmy
    ? Number(dmy[1])
    : (/^\d{4}-\d{2}-(\d{2})(?:[T ]|$)/.exec(t)?.[1] ??
      // Only a token that NAMES its month in words can have its single
      // one-or-two-digit number read as a day. Without that condition
      // `2025-12` — a year-month category label, which the line chart spaces
      // proportionally — had its `12` taken for a day, compared against the
      // 1st that `Date.parse` returns for it, and was refused as a date; the
      // chart then fell back to even spacing. Caught by `line.test.ts`, which
      // is why the whole suite runs before anything is believed.
      (/[A-Za-z]/.test(t)
        ? ((m) => (m && m.length === 1 ? m[0] : undefined))(t.match(/(?<!\d)\d{1,2}(?!\d)/g))
        : undefined));
  if (asked != null && new Date(ms).getUTCDate() !== Number(asked)) return null;
  // Floor, not round: a token carrying a time of day at/after 12:00 would round UP
  // to the next calendar day. Bare dates are exact midnights, so this is a no-op
  // for them.
  return Math.floor(ms / DAY_MS);
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** Three-letter month prefix → month index, for the `mmm-yy` form. */
const MONTH_INDEX = new Map(MONTHS.map((m, i) => [m.toLowerCase(), i]));

/** Short label for an epoch-day value: "5 Jan" or "Jan 26" on month starts. */
export function formatDay(days: number, withYear = false): string {
  const d = new Date(days * DAY_MS);
  // A day number JS cannot make a Date of prints as literal text, and text is
  // not what `finiteNodes` filters — that net catches non-finite NUMBERS, so a
  // string sails through it into all three renderers and into the .pptx the
  // skill hands back. A Gantt whose dates arrived as epoch SECONDS (an ordinary
  // thing for an export or an agent to produce, and `data.dates` is a plain
  // passthrough flag) put "NaN undefined" in 18 of its 21 text nodes.
  //
  // Empty rather than a guess, which is exactly what `formatNumber` does for a
  // non-finite value four hundred lines up: a label that cannot be computed has
  // no right answer, and a blank tick is readable where "NaN undefined" is not.
  if (!Number.isFinite(days) || Number.isNaN(d.getTime())) return "";
  const m = MONTHS[d.getUTCMonth()];
  if (withYear) return `${m} ${String(d.getUTCFullYear()).slice(2)}`;
  return d.getUTCDate() === 1 ? m : `${d.getUTCDate()} ${m}`;
}

/**
 * The two ends of a calendar SPAN, as one label.
 *
 * `formatDay` collapses the first of a month to the bare month name, which is
 * right for the tick strip — a month gridline reads `Mar`, and the day number
 * would be noise on every one of them. It is wrong for a span, where the day
 * number IS the information: a task running 1 Jan to 31 Mar was labelled
 * `Jan–31 Mar`, and one running 1 Jan to 1 Apr — the ordinary shape of a
 * quarterly plan — was labelled `Jan–Apr`, carrying no dates at all. A 30-year
 * roadmap whose phases began and ended on the 1st read `Jan–Jan`.
 *
 * The year is added when the two ends fall in different years, because without
 * it `1 Jan–1 Jan` is the label for any span of whole years.
 */
export function formatDayRange(from: number, to: number): string {
  const day = (v: number) => {
    const d = new Date(v * DAY_MS);
    if (!Number.isFinite(v) || Number.isNaN(d.getTime())) return "";
    return `${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]}`;
  };
  const year = (v: number) => new Date(v * DAY_MS).getUTCFullYear();
  const crossesYears =
    Number.isFinite(from) &&
    Number.isFinite(to) &&
    !Number.isNaN(new Date(from * DAY_MS).getTime()) &&
    !Number.isNaN(new Date(to * DAY_MS).getTime()) &&
    year(from) !== year(to);
  const yy = (v: number) => ` ${String(year(v)).slice(2)}`;
  const a = day(from);
  const b = day(to);
  return `${a}${a && crossesYears ? yy(from) : ""}–${b}${b && crossesYears ? yy(to) : ""}`;
}

/** Epoch-day values of every Monday covering [minDay, maxDay]. */
export function weekStarts(minDay: number, maxDay: number): number[] {
  // Day 0 (1970-01-01) was a Thursday; Monday ≡ 4 (mod 7).
  const first = minDay + ((4 - (minDay % 7) + 7) % 7);
  const out: number[] = [];
  for (let d = first; d <= maxDay && out.length < 120; d += 7) out.push(d);
  return out;
}

/**
 * Epoch-day values of every `stepMonths`-th month start covering [minDay, maxDay].
 *
 * The step is what keeps a long span COVERED. The caller used to take every
 * month and then filter the result down to quarters, and the guard below bounds
 * the WALK — so on a span longer than the guard the walk stopped part-way and
 * the filter thinned what was left of it. A 40-year Gantt drew its last
 * gridline halfway across the plot and labelled it `Q4 18`, with bars running
 * on past it to 2040 and no time reference under them at all: not a missing
 * tick, an axis that stops and does not say so.
 *
 * Stepping instead of filtering makes the walk cost what the OUTPUT costs, so
 * the same guard covers 20 years of months, 60 of quarters and 240 of years.
 *
 * The first tick is aligned DOWN to a multiple of the step within its year, so
 * quarters land on Jan/Apr/Jul/Oct and years on January whatever month the data
 * happens to start in — which is what the filter it replaces did, and why a
 * quarterly span renders byte-identically.
 */
export function monthStarts(minDay: number, maxDay: number, stepMonths = 1): number[] {
  const step = Math.max(1, Math.round(stepMonths));
  const start = new Date(minDay * DAY_MS);
  let y = start.getUTCFullYear();
  let m = Math.floor(start.getUTCMonth() / step) * step;
  const out: number[] = [];
  for (let guard = 0; guard < 240; guard++) {
    const day = Date.UTC(y, m, 1) / DAY_MS;
    if (day > maxDay) break;
    if (day >= minDay) out.push(day);
    m += step;
    while (m >= 12) {
      m -= 12;
      y++;
    }
  }
  return out;
}

/** Compound annual growth rate between two values over `periods` steps. */
export function cagr(from: number, to: number, periods: number): number | null {
  if (periods <= 0 || from <= 0 || to <= 0) return null;
  return Math.pow(to / from, 1 / periods) - 1;
}

/**
 * OLS trend statistics: R² and the two-tailed p-value of the slope
 * (Student's t via the regularized incomplete beta function). Good charts
 * always state fit and significance next to a trend line.
 */
export function trendStats(pts: { x: number; y: number }[]): { r2: number; p: number | null } | null {
  const n = pts.length;
  if (n < 2) return null;
  const mx = pts.reduce((s, p) => s + p.x, 0) / n;
  const my = pts.reduce((s, p) => s + p.y, 0) / n;
  const sxx = pts.reduce((s, p) => s + (p.x - mx) ** 2, 0);
  const syy = pts.reduce((s, p) => s + (p.y - my) ** 2, 0);
  const sxy = pts.reduce((s, p) => s + (p.x - mx) * (p.y - my), 0);
  if (sxx <= 0 || syy <= 0) return null;
  const r2 = (sxy * sxy) / (sxx * syy);
  const df = n - 2;
  if (df < 1) return { r2, p: null };
  if (r2 >= 1) return { r2: 1, p: 0 };
  const t2 = (r2 * df) / (1 - r2);
  // Two-tailed p for Student's t: p = I_{df/(df+t²)}(df/2, 1/2).
  return { r2, p: betaI(df / 2, 0.5, df / (df + t2)) };
}

/**
 * Least-squares polynomial fit of degree `degree` (2–4 for scatter's higher-order
 * trend lines). Solves the normal equations of the Vandermonde system in a
 * centered AND unit-scaled variable u = (x − mx) / sx, and returns an evaluator
 * plus the fit's R² and the degree actually used — clamped to points − 2, so at
 * least one residual degree of freedom remains and the fit never interpolates
 * noise (a degree n−1 fit through n points has a meaningless R² of 1).
 */
export function polyTrend(
  pts: { x: number; y: number }[],
  degree: number,
): { at: (x: number) => number; r2: number; degree: number } | null {
  const n = pts.length;
  if (n < 2) return null;
  const d = Math.max(1, Math.min(Math.floor(degree), n - 2));
  const mx = pts.reduce((s, p) => s + p.x, 0) / n;
  const my = pts.reduce((s, p) => s + p.y, 0) / n;
  // Unit-scale the centered abscissa. Centering fixes LARGE-x conditioning, but
  // the power sums S[k] scale as span^k, so a SMALL x-span drives S[2d] below the
  // solver's 1e-12 pivot floor and the whole trendline silently vanished. In u the
  // sums are O(1) regardless of span; `at` divides by the same sx, so the fit is
  // identical, only better conditioned.
  const sx = pts.reduce((s, p) => Math.max(s, Math.abs(p.x - mx)), 0) || 1;
  const m = d + 1;
  // Power sums in u = (x − mx) / sx: S[k] = Σ uᵏ (k=0..2d), T[k] = Σ uᵏ·y (k=0..d).
  const S = new Array(2 * d + 1).fill(0);
  const T = new Array(m).fill(0);
  for (const p of pts) {
    const u = (p.x - mx) / sx;
    let up = 1;
    for (let k = 0; k <= 2 * d; k++) {
      S[k] += up;
      if (k < m) T[k] += up * p.y;
      up *= u;
    }
  }
  const A = Array.from({ length: m }, (_, i) => S.slice(i, i + m));
  const c = gaussSolve(A, T.slice());
  if (!c) return null;
  const at = (x: number) => {
    const u = (x - mx) / sx;
    let up = 1;
    let y = 0;
    for (let k = 0; k < m; k++) {
      y += c[k] * up;
      up *= u;
    }
    return y;
  };
  let ssRes = 0;
  let ssTot = 0;
  for (const p of pts) {
    ssRes += (p.y - at(p.x)) ** 2;
    ssTot += (p.y - my) ** 2;
  }
  const r2 = ssTot > 0 ? Math.max(0, 1 - ssRes / ssTot) : 1;
  return { at, r2, degree: d };
}

/** Gaussian elimination with partial pivoting; null if the matrix is singular. */
function gaussSolve(A: number[][], b: number[]): number[] | null {
  const n = b.length;
  for (let col = 0; col < n; col++) {
    let piv = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(A[r][col]) > Math.abs(A[piv][col])) piv = r;
    if (Math.abs(A[piv][col]) < 1e-12) return null;
    [A[col], A[piv]] = [A[piv], A[col]];
    [b[col], b[piv]] = [b[piv], b[col]];
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const f = A[r][col] / A[col][col];
      for (let k = col; k < n; k++) A[r][k] -= f * A[col][k];
      b[r] -= f * b[col];
    }
  }
  return b.map((v, i) => v / A[i][i]);
}

/** Human p-value: "< 0.001", "< 0.01", "< 0.05", or "= 0.31". */
export function formatP(p: number): string {
  for (const cut of [0.001, 0.01, 0.05]) if (p < cut) return `< ${cut}`;
  return `= ${p.toFixed(2)}`;
}

/** Regularized incomplete beta I_x(a, b) via continued fraction (NR-style). */
function betaI(a: number, b: number, x: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const lnBeta = lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x);
  const front = Math.exp(lnBeta);
  const symmetric = x >= (a + 1) / (a + b + 2);
  const [aa, bb, xx] = symmetric ? [b, a, 1 - x] : [a, b, x];
  // Lentz's continued fraction.
  let c = 1;
  let d = 1 - ((aa + bb) * xx) / (aa + 1);
  if (Math.abs(d) < 1e-30) d = 1e-30;
  d = 1 / d;
  let f = d;
  for (let m = 1; m <= 200; m++) {
    let num = (m * (bb - m) * xx) / ((aa + 2 * m - 1) * (aa + 2 * m));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) d = 1e-30;
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) c = 1e-30;
    d = 1 / d;
    f *= d * c;
    num = (-(aa + m) * (aa + bb + m) * xx) / ((aa + 2 * m) * (aa + 2 * m + 1));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) d = 1e-30;
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) c = 1e-30;
    d = 1 / d;
    const delta = d * c;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-10) break;
  }
  // `f` is Numerical Recipes' betacf continued fraction, so I_x(a,b) = front·f/a.
  // (An earlier `front·(f-1)/a` mixed in a different formulation's offset and made
  // every p-value wrong — e.g. p=0.014 where the true value is 0.205.)
  const result = (front * f) / aa;
  return symmetric ? 1 - result : result;
}

/** Lanczos log-gamma. */
function lnGamma(z: number): number {
  const g = [
    676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905,
    -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ];
  if (z < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * z)) - lnGamma(1 - z);
  z -= 1;
  let a = 0.99999999999980993;
  for (let i = 0; i < g.length; i++) a += g[i] / (z + i + 1);
  const t = z + g.length - 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}

/**
 * Count `values` into `k` equal bins tiling [lo, hi]. The caller supplies k
 * rather than a rule, because the only k worth using here is one derived from
 * the axis's own tick grid — see layoutScatter's marginals, where k is a
 * multiple of the tick-interval count so every tick is a bin edge. A rule
 * keyed off the sample size alone (Sturges, Freedman-Diaconis) produces edges
 * that land between the ticks, which is exactly what a marginal histogram
 * must not do: it is read against the axis beside it.
 *
 * A value on an interior edge counts to the upper bin; hi counts to the last.
 * Values outside [lo, hi] are ignored — the caller's domain is the axis, and
 * the axis already covers the data.
 */
export function histogramBins(values: number[], lo: number, hi: number, k: number): number[] {
  const counts = new Array(Math.max(1, k)).fill(0);
  const span = hi - lo;
  if (!(span > 0)) return counts;
  for (const v of values) {
    if (v < lo || v > hi) continue;
    const i = Math.min(counts.length - 1, Math.floor(((v - lo) / span) * counts.length));
    counts[i]++;
  }
  return counts;
}
