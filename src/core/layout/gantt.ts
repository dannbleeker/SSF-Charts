import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import {
  formatDay,
  formatDayRange,
  formatNumber,
  monthStarts,
  niceTicks,
  resolveFormat,
  resolveAxisFormat,
  weekStarts,
} from "../format";
import { seriesColor } from "../style";
import type { LayoutResult } from "./column";
import { bandFontSize, fitPlot, footnoteH, titleHeight, titleNode, MIN_LABEL_FS } from "./frame";
import { lerpColor, zoneFill } from "../color";

/** A plot rectangle in points — the shape `fitPlot` returns. */
interface PlotBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Bracket annotations above the timeline header.
 *
 * `spanLabel` is passed in rather than rebuilt: it closes over the timeline's own
 * format resolution, and the row loop uses the same one. Two copies of a label
 * formatter is how two parts of one chart start disagreeing about how a date is
 * written.
 */
function bracketNodes(
  brackets: readonly { label: string; from: number; to: number }[],
  toX: (v: number) => number,
  titleH: number,
  fs: number,
  style: ChartStyle,
  spanLabel: (s: number, e: number) => string,
): SceneNode[] {
  const out: SceneNode[] = [];
  brackets.forEach((b, i) => {
    const x1 = toX(b.from);
    const x2 = toX(b.to);
    const y = titleH + fs * 1.5;
    out.push(
      { kind: "line", x1, y1: y, x2, y2: y, stroke: style.text, strokeWidth: 1, name: `bracket-${i}` },
      { kind: "line", x1, y1: y, x2: x1, y2: y + 3.5, stroke: style.text, strokeWidth: 1, name: `bracket-tick-a-${i}` },
      { kind: "line", x1: x2, y1: y, x2, y2: y + 3.5, stroke: style.text, strokeWidth: 1, name: `bracket-tick-b-${i}` },
      {
        kind: "text",
        x: x1,
        y: y - fs * 1.35,
        w: x2 - x1,
        h: fs * 1.3,
        text: b.label || spanLabel(b.from, b.to),
        fontSize: fs * 0.9,
        bold: true,
        color: style.text,
        align: "center",
        valign: "middle",
        name: `bracket-label-${i}`,
      },
    );
  });
  return out;
}

/**
 * Holiday shading (any granularity).
 *
 * THE `workdays` GATE IS AT THE CALL SITE, not in here. It used to read
 * `for (const h of workdays ? [] : holidays)`, and iterating an empty array is
 * the same nothing as not being called — the condition reads better beside the
 * other two shading passes than buried in a loop header. Pointless under a
 * working-day scale: a holiday has no width left to shade.
 */
function holidayNodes(
  holidays: readonly number[],
  toX: (v: number) => number,
  plot: PlotBox,
  style: ChartStyle,
): SceneNode[] {
  const out: SceneNode[] = [];
  for (const h of holidays) {
    const x1 = Math.max(plot.x, toX(h));
    const x2 = Math.min(plot.x + plot.w, toX(h + 1));
    if (x2 > x1) {
      out.push({
        kind: "rect",
        x: x1,
        y: plot.y,
        w: x2 - x1,
        h: plot.h,
        fill: zoneFill(style.background, "#efe7e7"),
        name: `holiday-${h}`,
      });
    }
  }
  return out;
}

/**
 * Weekend shading in week granularity.
 *
 * GATED ON `!workdays` AT THE CALL SITE, explicitly rather than left to the
 * `x2 > x1` guard below: that only collapses for a Mon–Fri week. Under a custom
 * workweek (say Sun–Thu) Saturday has width again, so the guard passes and the
 * block shades a Sunday — a working day there — while leaving the real
 * non-working Friday unshaded.
 */
function weekendNodes(
  span: { lo: number; hi: number },
  toX: (v: number) => number,
  plot: PlotBox,
  style: ChartStyle,
): SceneNode[] {
  const out: SceneNode[] = [];
  for (let d = span.lo - 7; d <= span.hi + 7; d++) {
    if (d % 7 === 2) {
      // Day ≡ 2 (mod 7) is Saturday (day 0 = Thursday); shade Sat+Sun.
      const x1 = Math.max(plot.x, toX(d));
      const x2 = Math.min(plot.x + plot.w, toX(d + 2));
      if (x2 > x1) {
        out.push({
          kind: "rect",
          x: x1,
          y: plot.y,
          w: x2 - x1,
          h: plot.h,
          fill: zoneFill(style.background, "#f4f3f0"),
          name: `weekend-${d}`,
        });
      }
    }
  }
  return out;
}

/**
 * Simplified Gantt / timeline: categories are activities; rows named
 * Start and End give each activity's span on a numeric timeline (week,
 * month index, year — any number). A row named Milestone adds a diamond
 * marker at that position. think-cell's calendar-based Gantt is richer;
 * this covers the project-on-a-slide case.
 */
export function layoutGantt(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()));
  const starts = find(/^start$/i)?.values ?? [];
  const ends = find(/^end$/i)?.values ?? [];
  const milestones = find(/^milestone$/i)?.values ?? [];
  /** "After" row: 1-based predecessor index → dependency arrow. */
  const after = find(/^(after|dep(endency)?)$/i)?.values ?? [];
  /** "Today" row: a single date/number → today line. */
  const today = (find(/^today$/i)?.values ?? []).find((v): v is number => v != null);
  /** "Holiday(s)" row: dates shaded like weekends. */
  const holidays = (find(/^holidays?$/i)?.values ?? []).filter((v): v is number => v != null);
  // Progress (0-100 or 0-1) and plan-vs-actual baseline rows.
  const completes = find(/^%?\s*complete$/i)?.values ?? [];
  const baseStarts = find(/^baseline\s*start$/i)?.values ?? [];
  const baseEnds = find(/^baseline\s*end$/i)?.values ?? [];
  /** "Bracket <label>" rows: first/last non-null values span an annotation. */
  const brackets = data.series
    .filter((s) => /^bracket\b/i.test(s.name.trim()))
    .map((s) => {
      const vals = s.values.filter((v): v is number => v != null);
      // A bracket needs two endpoints to span anything; fewer and it is dropped
      // below. Guard the min/max so an empty row can't spread to ±Infinity here.
      const ok = vals.length >= 2;
      return {
        label: s.name.replace(/^bracket\s*:?\s*/i, "").trim(),
        from: ok ? Math.min(...vals) : 0,
        to: ok ? Math.max(...vals) : 0,
        ok,
      };
    })
    .filter((b) => b.ok);
  // "Column <label>" rows: a numeric gutter column beside the task labels, the
  // MS-Project table look. Values are per-category data, so this is a datasheet
  // row like Start/End rather than config. Each column resolves its own format,
  // so a Cost column and an FTE column keep their own precision — the chart-wide
  // format is resolved over the timeline's epoch-day ticks and means nothing
  // here. suffix/forceSign are deliberately not picked up: they are value-axis
  // concerns, and units belong in the label ("Column Cost €k").
  const columns = data.series
    .filter((s) => /^column\b/i.test(s.name.trim()))
    .map((s) => {
      const nums = s.values.filter((v): v is number => v != null);
      const fmt = resolveFormat(nums, {
        decimals: cfg.numberFormat?.decimals,
        locale: cfg.numberFormat?.locale,
      });
      const label = s.name
        .trim()
        .replace(/^column\s*:?\s*/i, "")
        .trim();
      const cells = s.values.map((v) => (v == null ? "" : formatNumber(v, fmt)));
      const w = Math.min(
        cfg.width * 0.12,
        Math.max(textWidth(label, fs * 0.9), ...cells.map((t) => textWidth(t, fs))) + 10,
      );
      return { label, cells, w };
    });

  // "Activity | Owner | Remark" category convention; ">" prefix indents.
  const parts = data.categories.map((c) => c.split("|").map((p) => p.trim()));
  const indents = parts.map((p) => (p[0].startsWith(">") ? 1 : 0));
  const acts = parts.map((p) => p[0].replace(/^>+\s*/, ""));
  const owners = parts.map((p) => p[1] ?? "");
  const remarks = parts.map((p) => p[2] ?? "");
  const hasOwners = owners.some(Boolean);
  const hasRemarks = remarks.some(Boolean);
  // A row with no bar data at all is a section header.
  const isHeader = data.categories.map((_, c) => starts[c] == null && ends[c] == null && milestones[c] == null);

  // Critical path: over the "After" dependency edges, find the chain with the
  // greatest cumulative duration and flag its activities + connecting arrows.
  const predOf = (c: number) => {
    const pred = after[c];
    if (pred == null) return -1;
    const p = Math.round(pred) - 1;
    return p >= 0 && p < data.categories.length && p !== c ? p : -1;
  };
  const critical = new Set<number>();
  if (decor.criticalPath && after.some((v) => v != null)) {
    const dur = (c: number) => (starts[c] != null && ends[c] != null ? Math.max(0, ends[c]! - starts[c]!) : 0);
    // Longest cumulative duration ending at each activity (memoized; cycle-safe).
    const cum: number[] = data.categories.map(() => -1);
    const seen = new Set<number>();
    const longest = (c: number): number => {
      if (cum[c] >= 0) return cum[c];
      if (seen.has(c)) return dur(c); // break any accidental cycle
      seen.add(c);
      const p = predOf(c);
      const v = dur(c) + (p >= 0 ? longest(p) : 0);
      seen.delete(c);
      return (cum[c] = v);
    };
    let end = -1;
    let best = -1;
    data.categories.forEach((_, c) => {
      if (!isHeader[c] && longest(c) > best) {
        best = longest(c);
        end = c;
      }
    });
    for (let c = end; c >= 0; c = predOf(c)) {
      if (critical.has(c)) break;
      critical.add(c);
    }
  }

  const titleH = titleHeight(cfg, style);
  const bracketH = brackets.length ? fs * 1.9 : 0;
  const headerH = fs * 1.6;
  /**
   * The task-name gutter, reserved for the longest name AT ITS OWN INDENT.
   *
   * It used to reserve `longest name + 10` and then draw each row into
   * `catW - 6 - indent * 10`, so every indented row got a box ten points
   * narrower than the reservation had allowed for. With a lane header above them
   * — `gantt.lanes`, which indents every task — that is every task on the chart:
   * a 560x280 plan reserved 53 points for "Handover" and drew it into 37.
   *
   * Invisible until the names were clipped to their boxes, because an unclipped
   * name simply ran out of the gutter and across the plot. Two defects that
   * cancelled in appearance and compounded in fact.
   */
  const catW0 = Math.min(cfg.width * 0.32, Math.max(0, ...acts.map((c, i) => textWidth(c, fs) + indents[i] * 10)) + 10);
  const ownerW0 = hasOwners ? Math.max(0, ...owners.map((o) => textWidth(o, fs))) + 12 : 0;
  const remarkW0 = hasRemarks ? Math.max(0, ...remarks.map((o) => textWidth(o, fs * 0.9))) + 12 : 0;
  // Nothing capped what the gutters take TOGETHER: a "Column" row is capped at
  // 12% of the width but their COUNT is not, and the owner/remark gutters grow
  // with their longest string, so enough of them drove plot.w NEGATIVE — which
  // inverts toX, and bars came out zero-wide and left of the plot. Keep a fifth
  // of the width for the timeline: drop whole gutter columns once they no longer
  // fit (a sliver column shows nothing anyway), then shrink the text gutters if
  // the labels alone still overflow.
  const gutterBudget = Math.max(0, cfg.width * 0.8 - 6);
  let room = Math.max(0, gutterBudget - catW0 - ownerW0 - remarkW0);
  const fitted = columns.filter((c) => {
    if (c.w > room) return false;
    room -= c.w;
    return true;
  });
  const colsW = fitted.reduce((a, c) => a + c.w, 0);
  const textGutters = catW0 + ownerW0 + remarkW0;
  const gutterScale = textGutters > gutterBudget ? gutterBudget / textGutters : 1;
  const catW = catW0 * gutterScale;
  const ownerW = ownerW0 * gutterScale;
  const remarkW = remarkW0 * gutterScale;
  const bottomH = (today != null ? fs * 1.6 : 6) + footnoteH(cfg, style, decor);
  // Fitted before the row geometry is derived from it: `slotH` and `barH` read
  // plotH, so a frame too short for the chrome would give every row a negative
  // height and hang the bars above the plot rather than merely squashing them.
  const plotBox = fitPlot(cfg, {
    x: 0,
    y: titleH + bracketH + headerH,
    w: cfg.width,
    h: cfg.height - titleH - bracketH - headerH - bottomH,
  });
  const plotH = plotBox.h;
  /**
   * The header strip above the plot — the timeline's tick labels and the gutter
   * column headings — measured against the TITLE rather than taken as a flat
   * `fs * 1.6`.
   *
   * `fitPlot` grows the plot up from its bottom edge on a frame that cannot pay
   * for its chrome, so `plot.y - headerH` walked back into the title's band and
   * the tick labels were drawn across the chart's own title at 18pt on a 300x60
   * frame. Same rule the totals in `column.ts` take: fit the band that is
   * actually there, and where it cannot carry a readable label, do not draw one.
   */
  const headBand = Math.max(0, Math.min(headerH, plotBox.y - titleH));
  const headFs = bandFontSize(fs * 0.9, headBand, 1.2);
  // Row geometry is derived here rather than after the plot because the RIGHT
  // MARGIN depends on it (see `msR`). It reads `plotH` and the row count only —
  // neither depends on the plot's width — so hoisting it changes no value.
  const slotH = plotH / Math.max(1, data.categories.length);
  // The three row texts — activity, owner, remark — are each centred on their
  // row in a box `fs * 1.5` tall, so once the font outgrows the row pitch they
  // overlap the rows above and below and the last one leaves the plot (8.7pt
  // past a 200x150 frame at a 32pt font). Bound by the row they label, and
  // shrunk together so one row reads at one size. At any font that already fits
  // its row this is `fs` and nothing moves.
  // Bound by the row, floored by legibility: a Gantt on a sliver of a frame
  // gave every activity name a ONE-POINT font, which is ink rather than text and
  // is below what OOXML will even accept. Zero means the row cannot carry a
  // name, and the name is dropped — the bar still shows the span.
  const rowFs = bandFontSize(fs, slotH, 1.5);
  /** The in-bar span label is drawn at 0.9 of the row font, so it floors later. */
  const barLabelFs = rowFs * 0.9 >= MIN_LABEL_FS ? rowFs * 0.9 : 0;
  const barH = Math.min(slotH * 0.55, fs * 1.4);
  // A milestone marker is a circle CENTRED on its date, so half of it sits to
  // the RIGHT of the last position the timeline can reach — and the right margin
  // was a flat 6pt regardless of how big that circle is. A milestone on the last
  // date therefore had its right half cut off by the frame, from a 12pt font up:
  // +1.6pt at 12, +4.1 at 16, +4.8 at 22.
  //
  // That is a data-bearing marker, not a label. Its centre IS the date, so it
  // cannot be nudged left to fit the way a label can — the timeline has to end
  // far enough from the edge for the whole marker to sit inside. Reserved only
  // when the data HAS a milestone, so a gantt without one keeps the geometry it
  // always had.
  const msR = milestones.some((v) => v != null) ? barH * 0.45 : 0;
  const plot = fitPlot(cfg, {
    x: catW + colsW,
    y: plotBox.y,
    // Math.max: at an absurdly small cfg.width even zero gutters overrun the 6pt
    // right margin, and a negative width would invert the timeline again.
    w: Math.max(0, cfg.width - catW - colsW - ownerW - remarkW - 6 - msR),
    h: plotH,
  });

  const dates = !!data.dates;
  const all = [
    ...starts,
    ...ends,
    ...milestones,
    // A plan that slipped has its baseline OUTSIDE the actual span; leaving the
    // ghost bars out of the extent drew them off the plot (and off the canvas).
    ...baseStarts,
    ...baseEnds,
    ...(today != null ? [today] : []),
    ...holidays,
    ...brackets.flatMap((b) => [b.from, b.to]),
  ].filter((v): v is number => v != null);
  const lo = Math.min(...(all.length ? all : [0]));
  const hi = Math.max(...(all.length ? all : [1]));
  // Calendar granularity by span: weeks → months → quarters → years.
  //
  // The years tier is what keeps a long plan's axis under it. Quarters used to
  // run to any span, taken by filtering a month walk that is itself bounded —
  // so a 40-year roadmap got 80 gridlines covering its first 18 years and
  // nothing at all across the remaining half of the plot. Even without the
  // bound, 160 quarter ticks on one axis is not a timeline anybody reads.
  //
  // `monthStarts` takes the step now rather than being filtered afterwards, so
  // the walk it does is the walk the ticks need. A quarterly span is unchanged:
  // the step aligns to Jan/Apr/Jul/Oct exactly as the filter selected.
  const weeks = dates && hi - lo <= 130;
  const quarters = dates && hi - lo > 550;
  // ~6 years, the point at which quarters pass two dozen ticks.
  const years = dates && hi - lo > 2200;
  const ticks = dates
    ? weeks
      ? weekStarts(lo - 7, hi + 7)
      : monthStarts(lo - 31, hi + 31, years ? 12 : quarters ? 3 : 1)
    : niceTicks(lo, hi, 6);
  const t0 = dates ? Math.min(lo, ticks[0] ?? lo) : ticks[0];
  const t1 = dates ? Math.max(hi, ticks[ticks.length - 1] ?? hi) : ticks[ticks.length - 1];

  // Working-day timeline: give non-working days zero width so a bar's length
  // reads as the working days it contains. Every x in this file goes through
  // toX, so bars, milestones, dependencies, brackets, summary bars, baselines,
  // progress fills, the today line and the gridlines all follow from here.
  // Calendar only — "working day" means nothing on a numeric timeline.
  const wdCfg = cfg.gantt?.workdays;
  const workSet =
    wdCfg === true ? new Set([1, 2, 3, 4, 5]) : Array.isArray(wdCfg) ? new Set(wdCfg.map((n) => Math.round(n))) : null;
  // ~55 years. A mistyped date must not allocate a giant array.
  const SPAN_CAP = 20000;
  const workdays = dates && !!workSet && workSet.size > 0 && t1 - t0 <= SPAN_CAP;
  // ISO weekday, day 0 = Thursday. The (d%7+7)%7 fold also handles pre-1970
  // days, which the weekend-shading loop below never did.
  const isoDow = (d: number) => (((((d % 7) + 7) % 7) + 3) % 7) + 1;
  const offDays = new Set(holidays.map((h) => Math.round(h)));
  const pre: number[] = [];
  if (workdays) {
    let n = 0;
    for (let d = Math.floor(t0); d <= Math.ceil(t1); d++) {
      // Working days STRICTLY BEFORE d, so a span's width is the working days
      // it contains: Mon→Fri stays 4 units (as on the elapsed scale, so the
      // "End is an instant" convention is unchanged), Mon→Mon becomes 5, not 7.
      pre.push(n);
      if (workSet!.has(isoDow(d)) && !offDays.has(d)) n++;
    }
  }
  // The last prefix entry, not the running counter: it is by construction the
  // working-day count of [t0, t1), which is what makes toX(t1) land exactly on
  // the right edge, as the linear branch does.
  const workTotal = pre.length ? pre[pre.length - 1] : 0;
  const workIndex = (v: number) => pre[Math.max(0, Math.min(pre.length - 1, Math.round(v) - Math.floor(t0)))];
  const toX = (v: number) =>
    workdays
      ? plot.x + (workIndex(v) / Math.max(1, workTotal)) * plot.w
      : plot.x + ((v - t0) / (t1 - t0 || 1)) * plot.w;
  /**
   * Width for a span that collapses to nothing: a weekend-only task on a
   * working-day scale, or a single-day activity (Start === End) on any scale.
   * A constant, so the result stays deterministic.
   */
  const HAIRLINE = 1.5;
  /**
   * Only a working-day scale can collapse a span that covers real days, so this
   * stays 0 otherwise and no existing output moves.
   */
  const minW = workdays ? HAIRLINE : 0;
  const fmt = resolveAxisFormat(ticks, cfg.numberFormat);
  const tickLabel = (t: number, i: number) => {
    if (!dates) return formatNumber(t, fmt);
    const d = new Date(t * 86400000);
    // A yearly tick is a year, not "Q1" of one — every tick would carry the
    // same quarter, which reads as a quarterly axis that has lost its other
    // three quarters. Full four digits, because a 40-year span crossing a
    // century is exactly the case this tier exists for and `00` does not say
    // which one.
    if (years) return String(d.getUTCFullYear());
    if (quarters) return `Q${Math.floor(d.getUTCMonth() / 3) + 1} ${String(d.getUTCFullYear()).slice(2)}`;
    if (weeks) return formatDay(t);
    return formatDay(t, i === 0 || d.getUTCMonth() === 0);
  };
  // `formatDayRange`, not two `formatDay`s: a span's ends are two specific days,
  // and `formatDay`'s month-start shorthand belongs to the tick strip, where a
  // lone month name IS the tick. Through it, a task running 1 Jan to 1 Apr was
  // labelled `Jan–Apr` and one running whole years `Jan–Jan`.
  const spanLabel = (s: number, e: number) =>
    dates ? formatDayRange(s, e) : `${formatNumber(s, fmt)}–${formatNumber(e, fmt)}`;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // PUSH ORDER IS PAINT ORDER. Brackets sit above the header; both shadings go
  // behind everything the plot draws. Each gate is here rather than inside its
  // helper so the three passes read as one decision about the calendar.
  nodes.push(...bracketNodes(brackets, toX, titleH, fs, style, spanLabel));
  if (!workdays) nodes.push(...holidayNodes(holidays, toX, plot, style));
  if (weeks && !workdays) nodes.push(...weekendNodes({ lo, hi }, toX, plot, style));

  // Timeline header on top + vertical gridlines (think-cell's calendar strip).
  /**
   * How far apart two date labels must be before both are drawn.
   *
   * `fs * 2.6` — a constant, and a constant cannot know how wide a date is.
   * "Dec" clears it easily; "December 2024" is more than twice it at the same
   * font, so the thinning below let two of them through side by side and they
   * were drawn into each other. 36 pairs of `timeline` on `timeline` in the
   * variant sweep, the third-largest shape left in it.
   *
   * The widest label the strip will draw, plus a two-point gutter, is what the
   * gap actually has to be. Kept as a floor against the old constant so a strip
   * of short labels thins exactly as it always did and no ordinary gantt moves.
   */
  const minLabelGap = Math.max(
    fs * 2.6,
    ticks.reduce((m, t, i) => Math.max(m, textWidth(tickLabel(t, i), headFs)), 0) + 2,
  );
  let lastLabelX = -1e9;
  ticks.forEach((t, i) => {
    const x = toX(t);
    nodes.push({
      kind: "line",
      x1: x,
      y1: plot.y,
      x2: x,
      y2: plot.y + plot.h,
      stroke: style.gridline,
      strokeWidth: 0.75,
      name: "gridline",
    });
    // Thin out header labels when months are dense.
    //
    // THINNED ON THE NUDGED POSITION, not the raw tick. The nudge below moves
    // the first label right and the last one left to keep them on the chart,
    // and either move closes the gap this test had just verified — the same
    // clamp-undoes-the-fit defect as the secondary axis and the scatter's x
    // strip. Testing where the label actually lands costs nothing and needs no
    // second rule.
    const text = tickLabel(t, i);
    const half = textWidth(text, headFs) / 2;
    const at = Math.min(Math.max(x, half), Math.max(half, cfg.width - half));
    if (at - lastLabelX >= minLabelGap) {
      // A tick label is CENTRED on its tick, so the last one puts half its width
      // past the end of the timeline and off the chart — from a 14pt font, and
      // separately from the milestone above: reserving the marker radius moved
      // the last tick left far enough to hide this, so a gantt with no
      // Milestone row still lost the right-hand end of its axis (+8.6pt at 30).
      //
      // Nudged by exactly the overflow rather than bounded by the 48pt box: the
      // ink is what leaves the chart, the box is wider than the ink here, and a
      // tick label that has moved further than it had to no longer reads as
      // belonging to its tick. Every label that already fits is untouched.
      if (headFs > 0)
        nodes.push({
          kind: "text",
          x: at - 24,
          y: plot.y - headBand,
          w: 48,
          h: headBand,
          text,
          fontSize: headFs,
          color: style.mutedText,
          align: "center",
          valign: "middle",
          name: "timeline",
        });
      lastLabelX = at;
    }
  });

  const columnTop: number[] = [];

  data.categories.forEach((_, c) => {
    const cy = plot.y + slotH * (c + 0.5);
    columnTop.push(cy - barH / 2);
    // Section header rows: bold label, light band, no bar.
    if (isHeader[c]) {
      nodes.push({
        kind: "rect",
        x: 0,
        y: cy - slotH / 2 + 1,
        w: cfg.width,
        h: slotH - 2,
        fill: zoneFill(style.background, "#f0efec"),
        name: `section-${c}`,
      });
      if (rowFs > 0)
        nodes.push({
          kind: "text",
          x: 0,
          y: cy - rowFs * 0.75,
          w: cfg.width,
          h: rowFs * 1.5,
          text: acts[c],
          fontSize: rowFs,
          bold: true,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `category-${c}`,
        });
      // Auto-summary bar: span min(start)→max(end) of the child activities
      // (the rows below this header up to the next header), with end caps.
      if (decor.summaryBars) {
        let s = Infinity;
        let e = -Infinity;
        for (let k = c + 1; k < data.categories.length && !isHeader[k]; k++) {
          for (const v of [starts[k], ends[k], milestones[k]]) {
            if (v != null) {
              s = Math.min(s, v);
              e = Math.max(e, v);
            }
          }
        }
        if (e > s) {
          const x1 = toX(s);
          const x2 = toX(e);
          const sbH = barH * 0.4;
          nodes.push(
            {
              kind: "rect",
              x: x1,
              y: cy - sbH / 2,
              w: Math.max(x2 - x1, minW),
              h: sbH,
              fill: style.text,
              name: `summary-${c}`,
            },
            {
              kind: "line",
              x1,
              y1: cy - sbH / 2,
              x2: x1,
              y2: cy + sbH * 1.4,
              stroke: style.text,
              strokeWidth: 1.25,
              name: `summary-cap-a-${c}`,
            },
            {
              kind: "line",
              x1: x2,
              y1: cy - sbH / 2,
              x2,
              y2: cy + sbH * 1.4,
              stroke: style.text,
              strokeWidth: 1.25,
              name: `summary-cap-b-${c}`,
            },
          );
        }
      }
      return;
    }
    if (rowFs > 0) {
      // CLIPPED TO THE GUTTER IT WAS GIVEN, which it never was.
      //
      // `catW` is capped at 32% of the chart and then scaled down again when the
      // gutters together want more than 80% of it, so a long task name is
      // routinely wider than the box it is drawn in. The box said 32 points and
      // the ink was 86: left-aligned, the name ran straight out of the gutter,
      // across the plot and over the bar labels inside the bars — 18 pairs in
      // the variant sweep, and the numbers on those bars are what the chart is
      // for.
      //
      // Every other label in this engine that is given a box goes through
      // `clipToWidth`; this one was reading `acts[c]` raw. An ellipsis says the
      // name is longer than the room, which is true and which nothing else was
      // saying.
      const catBoxW = catW - 6 - indents[c] * 10;
      nodes.push({
        kind: "text",
        x: indents[c] * 10,
        y: cy - rowFs * 0.75,
        w: catBoxW,
        h: rowFs * 1.5,
        text: clipToWidth(acts[c], rowFs, catBoxW),
        fontSize: rowFs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `category-${c}`,
      });
    }
    // Responsible + remark columns right of the timeline.
    if (hasOwners && owners[c]) {
      nodes.push({
        kind: "text",
        x: plot.x + plot.w + 6,
        y: cy - rowFs * 0.75,
        w: ownerW - 6,
        h: rowFs * 1.5,
        // Clipped like the task name beside it, and for the same reason: `ownerW`
        // is scaled down whenever the gutters together want more than 80% of the
        // chart, so this box is routinely narrower than the name in it.
        text: clipToWidth(owners[c], rowFs, ownerW - 6),
        fontSize: rowFs,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: `owner-${c}`,
      });
    }
    if (hasRemarks && remarks[c]) {
      nodes.push({
        kind: "text",
        x: plot.x + plot.w + ownerW + 4,
        y: cy - rowFs * 0.7,
        w: remarkW - 4,
        h: rowFs * 1.4,
        text: clipToWidth(remarks[c], rowFs * 0.9, remarkW - 4),
        fontSize: rowFs * 0.9,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: `remark-${c}`,
      });
    }
    // Faint row separator.
    if (c > 0) {
      nodes.push({
        kind: "line",
        x1: plot.x,
        y1: cy - slotH / 2,
        x2: plot.x + plot.w,
        y2: cy - slotH / 2,
        stroke: style.gridline,
        strokeWidth: 0.5,
        name: `row-${c}`,
      });
    }
    // A pair typed end-before-start is a data-entry slip, not a reason to draw
    // nothing: take the span in the order the two dates actually run.
    const s0 = starts[c];
    const e0 = ends[c];
    const s = s0 != null && e0 != null ? Math.min(s0, e0) : s0;
    const e = s0 != null && e0 != null ? Math.max(s0, e0) : e0;
    // Baseline ghost bar (the original plan), thin, beneath the actual bar.
    const bs = baseStarts[c];
    const be = baseEnds[c];
    if (bs != null && be != null && be > bs) {
      nodes.push({
        kind: "rect",
        x: toX(bs),
        y: cy + barH * 0.55,
        w: Math.max(toX(be) - toX(bs), minW),
        h: barH * 0.4,
        // Like every other background tint in this file: unchanged on a light
        // canvas, a faint lift off a dark one instead of a light-grey glare.
        fill: zoneFill(style.background, "#cfcdc5"),
        name: `gantt-baseline-${c}`,
      });
    }
    if (s != null && e != null && e >= s) {
      const isCrit = critical.has(c);
      const barFill = seriesColor(style, 0);
      const bx = toX(s);
      // A task living entirely inside non-working days has zero working length,
      // and a one-day activity (a cutover typed with the same date twice) has no
      // span at all. Both would vanish; keep a hairline so the row still shows
      // its activity instead of reading as empty.
      const bw = Math.max(toX(e) - bx, e === s ? HAIRLINE : minW);
      nodes.push({
        kind: "rect",
        x: bx,
        y: cy - barH / 2,
        w: bw,
        h: barH,
        fill: barFill,
        name: `bar-${c}`,
        ...(isCrit ? { stroke: style.negative, strokeWidth: 1.75 } : {}),
      });
      // Percent-complete fill: a denser shade OF the bar over the elapsed share.
      // A fixed blue read as a foreign series on any other palette, and on a
      // dark canvas it was BRIGHTER than the bar it sat in — the opposite of
      // what it promises. Blending toward the text ink darkens it on a light
      // canvas and brightens it on a dark one, always in the bar's own hue.
      const rawPct = completes[c];
      if (rawPct != null) {
        const pct = Math.max(0, Math.min(1, rawPct > 1 ? rawPct / 100 : rawPct));
        if (pct > 0) {
          nodes.push({
            kind: "rect",
            x: bx,
            y: cy - barH / 2,
            w: bw * pct,
            h: barH,
            fill: lerpColor(barFill, style.text, 0.35),
            name: `progress-${c}`,
          });
        }
      }
      if (decor.segmentLabels) {
        const label = spanLabel(s, e);
        if (bw >= textWidth(label, fs * 0.9) + 4 && barLabelFs > 0) {
          nodes.push({
            kind: "text",
            x: bx,
            y: cy - rowFs * 0.7,
            w: bw,
            h: rowFs * 1.4,
            text: label,
            fontSize: barLabelFs,
            // Ink chosen for the bar it sits on, like every other in-shape
            // label: white vanished on a light palette colour.
            color: contrastInk(barFill),
            align: "center",
            valign: "middle",
            name: `bar-label-${c}`,
          });
        }
      }
    }
    const m = milestones[c];
    if (m != null) {
      const r = barH * 0.45;
      // Diamond milestone marker.
      nodes.push({
        kind: "ellipse",
        cx: toX(m),
        cy,
        rx: r,
        ry: r,
        fill: style.text,
        name: `milestone-${c}`,
      });
    }
  });

  // Gutter columns, in their own pass. This has to come after the row loop:
  // the section-header band is a full-width rect (`section-*`, x: 0,
  // w: cfg.width) and would paint straight over cells emitted alongside it.
  // A header row shows whatever value it carries — nothing is auto-summed here
  // (summaryBars sums spans, not money).
  fitted.forEach((col, i) => {
    let cx = catW;
    for (let k = 0; k < i; k++) cx += fitted[k].w;
    if (headFs > 0)
      nodes.push({
        kind: "text",
        x: cx,
        y: plot.y - headBand,
        w: col.w - 6,
        h: headBand,
        text: col.label,
        fontSize: headFs,
        bold: true,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: `col-head-${i}`,
      });
    data.categories.forEach((_, c) => {
      if (!col.cells[c]) return;
      const cy = plot.y + slotH * (c + 0.5);
      nodes.push({
        kind: "text",
        x: cx,
        y: cy - rowFs * 0.75,
        w: col.w - 6,
        h: rowFs * 1.5,
        text: col.cells[c],
        fontSize: rowFs,
        color: isHeader[c] ? style.text : style.mutedText,
        bold: isHeader[c],
        align: "right",
        valign: "middle",
        name: `col-${i}-${c}`,
      });
    });
  });

  // Dependency arrows ("After" row): elbow from the predecessor's end down
  // to the successor's start.
  data.categories.forEach((_, c) => {
    const pred = after[c];
    if (pred == null) return;
    const p = Math.round(pred) - 1;
    if (p < 0 || p >= data.categories.length || p === c) return;
    // A gate row carries only a Milestone, and it is the most natural thing to
    // depend on ("Build starts after the Kickoff gate"); anchoring on End alone
    // dropped the arrow the author asked for. Same on the successor side.
    const predEnd = ends[p] ?? milestones[p] ?? starts[p];
    const succStart = starts[c] ?? milestones[c];
    if (predEnd == null || succStart == null) return;
    const x1 = toX(predEnd);
    const yPred = plot.y + slotH * (p + 0.5);
    const ySucc = plot.y + slotH * (c + 0.5);
    const x2 = toX(succStart);
    // Edges on the critical path are drawn thicker and red.
    const critEdge = critical.has(c) && critical.has(p);
    const dcolor = critEdge ? style.negative : style.mutedText;
    const dw = critEdge ? 1.75 : 1;
    nodes.push(
      { kind: "line", x1, y1: yPred, x2: x1, y2: ySucc, stroke: dcolor, strokeWidth: dw, name: `dep-v-${c}` },
      { kind: "line", x1, y1: ySucc, x2: x2 - 2, y2: ySucc, stroke: dcolor, strokeWidth: dw, name: `dep-h-${c}` },
      {
        kind: "arrowhead",
        x: x2 - 1,
        y: ySucc,
        angle: x2 >= x1 ? 0 : 180,
        size: critEdge ? 4.2 : 3.5,
        fill: dcolor,
        name: `dep-head-${c}`,
      },
    );
  });

  // Today line.
  if (today != null && today >= t0 && today <= t1) {
    const x = toX(today);
    nodes.push(
      {
        kind: "line",
        x1: x,
        y1: plot.y,
        x2: x,
        y2: plot.y + plot.h,
        stroke: style.negative,
        strokeWidth: 1.25,
        dash: [3, 2],
        name: "today-line",
      },
      {
        kind: "text",
        x: x - 24,
        y: plot.y + plot.h + 1,
        w: 48,
        h: fs * 1.3,
        text: "Today",
        fontSize: fs * 0.85,
        color: style.negative,
        align: "center",
        valign: "top",
        name: "today-label",
      },
    );
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.y + slotH * (c + 0.5)),
      categoryWidth: data.categories.map(() => barH),
      columnTop,
      columnValue: data.categories.map((_, c) => ends[c] ?? 0),
      baselineY: plot.x,
      plot,
    },
  };
}
