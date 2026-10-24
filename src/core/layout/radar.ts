import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { polar, textWidth, type SceneNode } from "../scene";
import { formatNumber, niceTicks, resolveAxisFormat } from "../format";
import { seriesColor } from "../style";
import { bandFontSize, fitPlot, footnoteH, legendRowCount, titleHeight, titleNode } from "./frame";
import { legendRow, type LayoutResult, type LegendEntry } from "./column";
import { columnPositiveTotal } from "./totals";
import { maxOf, minOf } from "../agg";

/**
 * Smallest a perimeter name may be drawn before it is dropped instead.
 *
 * The same 5pt the pie's labels use, and for the same reason: below it the text
 * is present without being readable, which costs the reader a glance and gives
 * them nothing. Kept as a named constant so the two rings cannot drift apart.
 */
const PERIM_FLOOR_FS = 5;

/**
 * Radar (spider) chart: categories = spokes (first at 12 o'clock,
 * clockwise), series = polygons with translucent fills. Gridlines are
 * straight polygons by default (business style); scale is shared across
 * all spokes and ticked on the 12 o'clock spoke only.
 */
export function layoutRadar(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  if (cfg.radar?.bars) return layoutRadialBars(cfg, style, decor);
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  // Stacked radar: series stack cumulatively along each spoke, so the scale
  // must reach the per-spoke sums, not the largest single value.
  const stacked = !!cfg.radar?.stacked && data.series.length >= 2;
  // Min–max band: shade the per-spoke envelope of the peer series (all but the
  // last) and draw the last series ("us") on top — the "peer range + us" view.
  const band = !!decor.radarBand && !stacked && data.series.length >= 2;
  // Legend entries are needed up here because the wrap walk decides how many
  // rows to reserve, and band mode legends two swatches, not one per series.
  const legendEntries: LegendEntry[] = band
    ? [
        { label: "Peer range", color: style.mutedText, name: "legend-band" },
        {
          label: data.series[data.series.length - 1].name,
          color: seriesColor(style, data.series.length - 1, data.series[data.series.length - 1].color),
          name: "legend-us",
        },
      ]
    : data.series.map((s, si) => ({ label: s.name, color: seriesColor(style, si, s.color), name: `legend-${si}` }));
  const legendRows =
    decor.seriesLabels && data.series.length > 1
      ? legendRowCount(
          legendEntries.map((e) => e.label),
          fs,
          0,
          cfg.width - 4,
        )
      : 0;
  // Zero when the band would not be ON the canvas. `legendH` both reserves the
  // strip and gates the draw below (`if (legendH) drawLegend()`), so one number
  // keeps the two in step — and it had no bound: two rows at a 32pt font want
  // 102 points, so on a 300x60 chart the entries were drawn at y 68 and y 119,
  // wholly below the foot of the chart. Chrome that cannot be paid for is not
  // drawn, and the web itself is the chart.
  const legendH = titleH + legendRows * fs * 1.6 <= cfg.height ? legendRows * fs * 1.6 : 0;
  // Fitted so the web's CENTRE and RADIUS are both derived from a positive box:
  // on a frame too short for title + legend + footnote the raw height goes
  // negative, which put the centre below the bottom of the chart and the
  // perimeter label ring past it.
  const box = fitPlot(cfg, { x: 0, y: titleH + legendH, w: cfg.width, h: cfg.height - titleH - legendH - footH });
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  // Perimeter labels need a margin around the web.
  const labelW = Math.max(0, ...data.categories.map((c) => textWidth(c, fs)));
  // The vertical term already reserves `fs * 1.9` for the perimeter label ring,
  // so the web fits its frame — unless the 10pt floor overrides it, which on a
  // small frame is exactly when it does: at 120x90 the fit asks for ~8 and the
  // floor insists on 10, putting the bottom category label through the frame.
  // The floor yields to the frame, like the tilemap's tile floor: a small web is
  // still a web, one drawn past the edge is a label nobody can read. Unreachable
  // on any chart big enough to want it, so nothing of an ordinary size moves.
  const rWant = Math.min(box.w / 2 - labelW - fs, box.h / 2 - fs * 1.9);
  const r = Math.max(1, rWant);
  // The perimeter label ring is what `rWant` reserves room for on both axes. When
  // the floor has to override it the reservation is gone, so the ring no longer
  // has anywhere to be drawn — the bottom label sat 12pt past a 70pt frame. Drop
  // the ring rather than draw it off the chart: a web this small is a dot, and a
  // label nobody can read is worse than no label. Unreachable at any size that
  // wants a radar.
  const ringFits = rWant >= 1;
  /**
   * The size the perimeter labels are drawn at.
   *
   * They sit on a ring, one per spoke, so the room each has is the SEPARATION
   * between adjacent spokes at the radius they sit on — and none was fitted to
   * it, so on a web whose font is large relative to its size adjacent names were
   * drawn over each other. Bound by that separation, the same way the ring ticks
   * are bound by their ring gap.
   *
   * Last resort: on any web whose names already clear each other this is `fs`,
   * which is every radar at an ordinary size.
   */
  // The room a perimeter label has is the chord between two neighbouring spokes
  // — but with a SINGLE category there is no neighbour, and `sin(PI / 1)` is
  // zero, so the one formula that should have been unconstrained produced a
  // font of zero. One spoke gets the chart font; two or more get the chord.
  const perimFs = n <= 1 ? fs : bandFontSize(fs, 2 * (r + fs * 0.6) * Math.sin(Math.PI / Math.max(1, n)), 1.4);

  const spokeSum = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  // `all` is cells-scaled (series x categories) so it must be FOLDED, not spread —
  // a large grid overflows the argument list. `spokeSum` is category-scaled and safe.
  //
  // The auto domain is the data's own, anchored at 0 the way every column layout's
  // is — not floored at 1, and not pinned at 0 from below. `niceTicks(0, 0)` already
  // answers [0, 1], so an all-zero web has its guard, and the `1` floors that stood
  // here as well put any radar of fractions on a 0..1 web (rings 0.5 / 1.0 for data
  // of 0.2-0.4, against 20 / 40 for the same numbers x100). The lower bound was `0`
  // whatever the data held, so with no `scale` a NEGATIVE value was clamped by `toR`
  // onto the centre: -20, -5 and 0 drew at one point, and an all-negative series put
  // every vertex there under rings labelled 0.5 and 1.0. A stacked web only ever
  // stacks the positive parts (`columnPositiveTotal`), so its floor stays 0.
  const tickMax = stacked ? Math.max(0, ...spokeSum) : maxOf(all, 0);
  const tickMin = stacked ? 0 : minOf(all, 0);
  const rawTicks = niceTicks(Math.min(0, cfg.scale?.min ?? tickMin), cfg.scale?.max ?? tickMax, 4);
  const min = cfg.scale?.min ?? rawTicks[0];
  const max = cfg.scale?.max ?? rawTicks[rawTicks.length - 1];
  // Clamp to the RESOLVED domain, as valueScale does. niceTicks rounds outward
  // from the auto range while min/max come from cfg.scale, so a pinned scale
  // left rings mapped past the outer radius — a grid circle 25% outside the web
  // and its tick label off the top of the canvas.
  const ticks = rawTicks.filter((t) => t >= min - 1e-9 && t <= max + 1e-9);
  if (!ticks.length) ticks.push(max);
  const fmt = resolveAxisFormat(ticks, cfg.numberFormat);
  // Clamped at BOTH ends. It held values below `scale.min` on the centre and let
  // anything above `scale.max` run past the outer radius unbounded — and
  // `sampleConfig("radar")` ships `{min:0,max:5}`, so no bad scale has to be
  // typed: edit one datasheet cell from 4 to 8, as anyone would for a 1-10
  // maturity scale, and the vertex leaves the web and the canvas with no ring
  // anywhere near it. A web is a picture of a scale; a point outside it is not
  // on the scale.
  const toR = (v: number) => ((Math.min(max, Math.max(min, v)) - min) / (max - min || 1)) * r;
  const angle = (c: number) => (360 / Math.max(1, n)) * c;

  /**
   * Where a perimeter label lands, at a given size. One source for the box the
   * crowding check measures and the box the node is drawn with, because two
   * copies of this arithmetic would drift and the check would then be measuring
   * a label the chart does not draw.
   *
   * The ANCHOR does not move with `size` — it is pinned at `r + fs * 0.6` off
   * the chart font — so shrinking a label pulls it in around its anchor rather
   * than sliding it along the ring. That is what makes shrinking safe here: a
   * label cannot travel to another spoke.
   */
  const perimBox = (c: number, size: number) => {
    const p = polar(cx, cy, r + fs * 0.6, angle(c));
    const a = angle(c) % 360;
    const align: "center" | "left" | "right" =
      a < 10 || a > 350 || Math.abs(a - 180) < 10 ? "center" : a < 180 ? "left" : "right";
    const w = textWidth(data.categories[c] ?? "", size) + 4;
    return {
      x: align === "center" ? p.x - w / 2 : align === "left" ? p.x : p.x - w,
      y: p.y - (a < 10 || a > 350 ? size * 1.4 : Math.abs(a - 180) < 10 ? 0 : size * 0.7),
      w,
      h: size * 1.4,
      align,
    };
  };

  /**
   * How big each perimeter name may be, given the ones beside it. `null` drops it.
   *
   * `perimFs` above bounds every name by the CHORD between two spokes, which is
   * the room a label has along the ring. It is not the room it has where the
   * ring runs horizontally: near the top and the bottom two adjacent names sit
   * almost side by side, and what limits them there is their WIDTH, which the
   * chord never looked at. A 24-spoke web drew `category-9` over `category-10`
   * for exactly that reason — the last of the "24 categories" family, after the
   * pie's outside labels were fixed the same way. Across the frame sweep: 96
   * overlapping pairs before, none after, and 236 of 288 names still drawn.
   *
   * SHRINK, THEN DROP, and nothing moves. The alternative — nudging labels apart
   * — puts a name beside the wrong spoke, which on a web is not a cosmetic loss:
   * the whole chart is read by which axis a point sits on. A name that leaves is
   * a name the reader looks up in the legend; a name that lies is worse than no
   * name at all.
   *
   * ONLY THE TWO ADJACENT SPOKES, and that was tested rather than assumed. An
   * all-pairs version was written first, on the worry that a wide name near the
   * top reaches past its neighbour to the one beyond. It changed no answer: a
   * search over 3,024 combinations — spoke counts 8 to 60, name lengths 1 to 40
   * characters, eight frames, six fonts — found not one overlap that the
   * neighbour check missed. Yielding to the neighbour already costs enough size
   * to clear everything past it. The general check was kept for one commit, cost
   * a spoke-count cap to bound its quadratic, and justified neither.
   *
   * Neighbours are measured at the SAME size, because an even ring shrinks
   * uniformly where it shrinks at all.
   */
  const perimSizes: (number | null)[] = (() => {
    if (!ringFits || perimFs <= 0) return data.categories.map(() => null);
    // OVERLAPPING AREA, not overlapping extent on both axes. The first version
    // asked for more than a point of intersection in x AND in y, which let a
    // sliver a third of a point wide and eight tall — plainly two names touching
    // — pass as clear. Four such pairs survived the whole pass. Area is also the
    // rule the frame gate measures overlaps by, and a layout that fits itself by
    // a different rule than the one it is checked against will always leave a
    // residue nobody can explain.
    const hits = (i: number, j: number, size: number) => {
      const a = perimBox(i, size);
      const b = perimBox(j, size);
      const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
      const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
      return w > 0 && h > 0 && w * h > 1;
    };
    return data.categories.map((_, c) => {
      if (n < 2) return perimFs;
      const before = (c - 1 + n) % n;
      const after = (c + 1) % n;
      let size = perimFs;
      while (size >= PERIM_FLOOR_FS && (hits(c, before, size) || hits(c, after, size))) size -= 0.5;
      return size >= PERIM_FLOOR_FS ? size : null;
    });
  })();

  // Per-spoke scales: normalise each spoke to its own maximum so spokes in
  // different KPI units become comparable in shape (numeric ticks dropped).
  const perSpoke = !!cfg.radar?.perSpoke && !stacked && data.series.length >= 1;
  const spokeMax = data.categories.map((_, c) => {
    if (!perSpoke) return max;
    // Each spoke normalises against its OWN maximum — that is the whole point
    // of perSpoke (a 0–1 rate beside a count). The floor only guards a
    // non-positive divisor; clamping at 1 kept any sub-1 spoke off its rim.
    const m = Math.max(0, ...data.series.map((s) => s.values[c] ?? 0));
    return m > 0 ? m : 1;
  });
  const toRc = (v: number, c: number) => (perSpoke ? (Math.max(0, v) / spokeMax[c]) * r : toR(Math.max(min, v)));

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Grid: concentric polygons (or circles). Per-spoke mode uses fraction
  // rings with no numeric labels (each spoke has its own scale); otherwise
  // rings sit at the value ticks and the 12 o'clock spoke is labelled.
  const gridShape = decor.gridShape ?? "polygon";
  const rings = perSpoke
    ? [0.25, 0.5, 0.75, 1].map((f) => ({ rr: f * r, t: f }))
    : ticks.filter((t) => t > min).map((t) => ({ rr: toR(t), t }));
  /**
   * A ring's tick label is centred on that ring in a box `fs * 1.2` tall, so once
   * the font outgrows the gap BETWEEN rings the labels sit on top of each other
   * and the innermost one runs past the middle of the web and out of the chart —
   * 11.1pt past a 200x150 frame at a 32pt font. Bound by the ring gap, which is
   * the space each label actually has, and dropped below the same 5pt floor the
   * perimeter ring uses: a label nobody can read is not worth the web it covers.
   *
   * Last resort, like every other shrink here: at any font that already fits its
   * ring gap this is `fs * 0.85` and nothing moves.
   */
  //
  // Measured from the rings THEMSELVES, not as `r / rings.length`. That average
  // is not the gap: the rings drawn are the ticks above the minimum, so the
  // outermost radius is not one of them and the division overstated the space
  // by a quarter — 7.0 where the rings are 5.6 apart on a 120x90 web. The
  // labels were then fitted to a gap they did not have and drew through each
  // other at 8pt and at 14pt.
  const ringGap = rings.length > 1 ? Math.min(...rings.slice(1).map((x, i) => x.rr - rings[i].rr)) : r;
  const tickGap = Math.max(0, ringGap);
  const tickFs = Math.min(fs * 0.85, tickGap / 1.2);
  // The box is `fs * 1.2` for a font of `fs * 0.85` — deliberately taller than
  // the text — so the two must shrink TOGETHER or the label moves even when
  // nothing needed shrinking. Expressed as the ratio for that reason: it is
  // exactly 1 whenever the tick font is untouched, which is every radar at an
  // ordinary size, and the geometry is then byte-identical. Collapsing both onto
  // `tickFs` instead shifted every tick by 0.9pt and moved three showcase
  // slides, which is how this was caught.
  const tickScale = tickFs / (fs * 0.85);
  for (const { rr, t } of rings) {
    if (gridShape === "circle") {
      nodes.push({
        kind: "ellipse",
        cx,
        cy,
        rx: rr,
        ry: rr,
        fill: "none",
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: `grid-${t}`,
      });
    } else {
      nodes.push({
        kind: "polygon",
        points: data.categories.map((_, c) => polar(cx, cy, rr, angle(c))),
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: `grid-${t}`,
      });
    }
    if (!perSpoke && tickFs >= 5) {
      // Tick label on the 12 o'clock spoke only.
      nodes.push({
        kind: "text",
        x: cx + 3,
        y: cy - rr - fs * 0.6 * tickScale,
        w: fs * 3.4,
        h: fs * 1.2 * tickScale,
        text: formatNumber(t, fmt),
        fontSize: tickFs,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: `tick-${t}`,
      });
    }
  }
  data.categories.forEach((cat, c) => {
    const end = polar(cx, cy, r, angle(c));
    nodes.push({
      kind: "line",
      x1: cx,
      y1: cy,
      x2: end.x,
      y2: end.y,
      stroke: style.gridline,
      strokeWidth: 0.75,
      name: `spoke-${c}`,
    });
    // Perimeter category label, anchored by which side of the web it sits on.
    // Dropped when the ring will not fit, and equally when the chord between
    // two spokes cannot carry a legible label — the same answer, from the same
    // reasoning, for the two ways the room can run out.
    if (!ringFits) return;
    const size = perimSizes[c];
    if (size == null) return;
    const b = perimBox(c, size);
    nodes.push({
      kind: "text",
      x: b.x,
      y: b.y,
      w: b.w,
      h: b.h,
      text: cat,
      fontSize: size,
      color: style.text,
      align: b.align,
      valign: "middle",
      name: `category-${c}`,
    });
  });

  // Min–max band (declared above with the legend entries it changes): shade the
  // per-spoke envelope of the peer series as an annulus of per-sector quads,
  // then draw the last series ("us") prominently on top.
  if (band) {
    const peers = data.series.slice(0, -1);
    const peerMin: number[] = [];
    const peerMax: number[] = [];
    for (let c = 0; c < n; c++) {
      const vals = peers.map((s) => s.values[c]).filter((v): v is number => v != null);
      peerMin[c] = vals.length ? Math.max(min, Math.min(...vals)) : min;
      peerMax[c] = vals.length ? Math.max(min, Math.max(...vals)) : min;
    }
    const minPts = data.categories.map((_, c) => polar(cx, cy, toRc(peerMin[c], c), angle(c)));
    const maxPts = data.categories.map((_, c) => polar(cx, cy, toRc(peerMax[c], c), angle(c)));
    for (let c = 0; c < n; c++) {
      const c2 = (c + 1) % n;
      nodes.push({
        kind: "polygon",
        points: [minPts[c], minPts[c2], maxPts[c2], maxPts[c]],
        fill: style.mutedText,
        fillOpacity: 0.16,
        name: `band-${c}`,
      });
    }
    nodes.push(
      { kind: "polygon", points: maxPts, stroke: style.mutedText, strokeWidth: 1, name: "band-max" },
      { kind: "polygon", points: minPts, stroke: style.mutedText, strokeWidth: 1, name: "band-min" },
    );
  }

  // Stacked radar: draw each series as an annular band between the cumulative
  // level below it and its own cumulative level (part-to-whole across spokes).
  if (stacked) {
    const cum = data.categories.map(() => 0);
    data.series.forEach((s, si) => {
      const innerPts = data.categories.map((_, c) => polar(cx, cy, toR(cum[c]), angle(c)));
      for (let c = 0; c < n; c++) cum[c] += Math.max(0, s.values[c] ?? 0);
      const outerPts = data.categories.map((_, c) => polar(cx, cy, toR(cum[c]), angle(c)));
      const color = seriesColor(style, si, s.color);
      nodes.push({
        kind: "polygon",
        points: [...outerPts, ...innerPts.slice().reverse()],
        fill: color,
        fillOpacity: decor.fillOpacity ?? 0.55,
        stroke: color,
        strokeWidth: 1.2,
        name: `series-${si}`,
      });
    });
    if (legendH) drawLegend();
    return {
      nodes,
      anchors: {
        categoryX: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).x),
        categoryWidth: data.categories.map(() => r / 2),
        columnTop: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).y),
        columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
        baselineY: cy,
        plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
      },
    };
  }

  // Series polygons: translucent fill (SVG), full-opacity outline + markers.
  const defaultOpacity = data.series.length === 1 ? 0.25 : 0.18;
  const drawIdx = band ? [data.series.length - 1] : data.series.map((_, i) => i);
  drawIdx.forEach((si) => {
    const s = data.series[si];
    const color = seriesColor(style, si, s.color);
    const pts = data.categories.map((_, c) => polar(cx, cy, toRc(s.values[c] ?? min, c), angle(c)));
    nodes.push({
      kind: "polygon",
      points: pts,
      fill: color,
      fillOpacity: decor.fillOpacity ?? defaultOpacity,
      stroke: color,
      strokeWidth: 1.6,
      name: `series-${si}`,
    });
    pts.forEach((p, c) => {
      if (s.values[c] == null) return;
      nodes.push({
        kind: "ellipse",
        cx: p.x,
        cy: p.y,
        rx: 2.4,
        ry: 2.4,
        fill: color,
        stroke: style.background,
        strokeWidth: 1,
        name: `marker-${si}-${c}`,
      });
    });
  });

  // Legend row under the title when there are multiple series. In band mode
  // it collapses the peers into one "Peer range" swatch plus the "us" series.
  if (legendH) drawLegend();

  // Hoisted so the stacked-radar early return can reuse it. Routed through the
  // shared wrapping legend rather than a forked single-row walk, so the chips
  // wrap inside cfg.width and land on exactly the rows legendRows reserved.
  function drawLegend() {
    nodes.push(...legendRow(cfg, style, 0, titleH, { maxX: cfg.width - 4, entries: legendEntries }));
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).x),
      categoryWidth: data.categories.map(() => r / 2),
      columnTop: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}

/**
 * Radial (polar) bar chart / coxcomb: each category is an equal-angle sector
 * whose bar radius encodes its value, drawn from a small inner hole so it reads
 * as bars bent around a circle rather than a pie. A single series colours bars
 * by category (Nightingale rose); multiple series stack outward within each
 * sector. Concentric value rings give the scale.
 */
function layoutRadialBars(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const multi = data.series.length > 1;

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  const legendEntries: LegendEntry[] = data.series.map((s, si) => ({
    label: s.name,
    color: seriesColor(style, si, s.color),
    name: `legend-${si}`,
  }));
  const legendRows =
    decor.seriesLabels && multi
      ? legendRowCount(
          legendEntries.map((e) => e.label),
          fs,
          0,
          cfg.width - 4,
        )
      : 0;
  // Zero when the band would not be ON the canvas. `legendH` both reserves the
  // strip and gates the draw below (`if (legendH) drawLegend()`), so one number
  // keeps the two in step — and it had no bound: two rows at a 32pt font want
  // 102 points, so on a 300x60 chart the entries were drawn at y 68 and y 119,
  // wholly below the foot of the chart. Chrome that cannot be paid for is not
  // drawn, and the web itself is the chart.
  const legendH = titleH + legendRows * fs * 1.6 <= cfg.height ? legendRows * fs * 1.6 : 0;
  // Fitted so the web's CENTRE and RADIUS are both derived from a positive box:
  // on a frame too short for title + legend + footnote the raw height goes
  // negative, which put the centre below the bottom of the chart and the
  // perimeter label ring past it.
  const box = fitPlot(cfg, { x: 0, y: titleH + legendH, w: cfg.width, h: cfg.height - titleH - legendH - footH });
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const labelW = Math.max(0, ...data.categories.map((c) => textWidth(c, fs)));
  // The vertical term already reserves `fs * 1.9` for the perimeter label ring,
  // so the web fits its frame — unless the 10pt floor overrides it, which on a
  // small frame is exactly when it does: at 120x90 the fit asks for ~8 and the
  // floor insists on 10, putting the bottom category label through the frame.
  // The floor yields to the frame, like the tilemap's tile floor: a small web is
  // still a web, one drawn past the edge is a label nobody can read. Unreachable
  // on any chart big enough to want it, so nothing of an ordinary size moves.
  const rWant = Math.min(box.w / 2 - labelW - fs, box.h / 2 - fs * 1.9);
  const r = Math.max(1, rWant);
  // The perimeter label ring is what `rWant` reserves room for on both axes. When
  // the floor has to override it the reservation is gone, so the ring no longer
  // has anywhere to be drawn — the bottom label sat 12pt past a 70pt frame. Drop
  // the ring rather than draw it off the chart: a web this small is a dot, and a
  // label nobody can read is worse than no label. Unreachable at any size that
  // wants a radar.
  const ringFits = rWant >= 1;
  const innerR = r * 0.18;

  // Scale reaches the per-category stack sums (a single series is its own sum).
  const catSum = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  // No floor of 1 — see the radar web above: it only clamped data below 1.
  const rawTicks = niceTicks(0, cfg.scale?.max ?? Math.max(0, ...catSum), 4);
  const max = cfg.scale?.max ?? rawTicks[rawTicks.length - 1];
  // Same clamp as the radar web above: a pinned scale.max must not leave rings
  // outside the outer radius.
  const ticks = rawTicks.filter((t) => t <= max + 1e-9);
  if (!ticks.length) ticks.push(max);
  const fmt = resolveAxisFormat(ticks, cfg.numberFormat);
  // Same one-sided clamp, same fix — see the radar web above.
  const toR = (v: number) => innerR + (Math.min(max, Math.max(0, v)) / (max || 1)) * (r - innerR);
  const sector = 360 / Math.max(1, n);
  const angle = (c: number) => sector * c;
  const pad = sector * 0.12;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Concentric value rings + tick labels on the 12 o'clock line.
  for (const t of ticks.filter((t) => t > 0)) {
    nodes.push({
      kind: "ellipse",
      cx,
      cy,
      rx: toR(t),
      ry: toR(t),
      fill: "none",
      stroke: style.gridline,
      strokeWidth: 0.75,
      name: `grid-${t}`,
    });
    nodes.push({
      kind: "text",
      x: cx + 3,
      y: cy - toR(t) - fs * 0.6,
      w: fs * 3.4,
      h: fs * 1.2,
      text: formatNumber(t, fmt),
      fontSize: fs * 0.85,
      color: style.mutedText,
      align: "left",
      valign: "middle",
      name: `tick-${t}`,
    });
  }

  data.categories.forEach((cat, c) => {
    const a0 = (((angle(c) + pad / 2) % 360) + 360) % 360;
    const aSpan = sector - pad;
    let base = 0;
    data.series.forEach((s, si) => {
      const v = Math.max(0, s.values[c] ?? 0);
      if (v <= 0) return;
      const rin = toR(base);
      const rout = toR(base + v);
      base += v;
      const color = multi
        ? seriesColor(style, si, s.color)
        : (s.colors?.[c] ?? style.palette[c % style.palette.length]);
      nodes.push({
        kind: "wedge",
        cx,
        cy,
        r: rout,
        innerR: rin,
        startAngle: a0,
        endAngle: a0 + aSpan,
        fill: color,
        stroke: style.background,
        strokeWidth: 1,
        name: multi ? `bar-${c}-${si}` : `bar-${c}`,
      });
    });
    // Perimeter category label.
    if (!ringFits) return;
    const mid = angle(c) + sector / 2;
    const p = polar(cx, cy, r + fs * 0.6, mid);
    const am = ((mid % 360) + 360) % 360;
    const align = am < 10 || am > 350 || Math.abs(am - 180) < 10 ? "center" : am < 180 ? "left" : "right";
    const w = textWidth(cat, fs) + 4;
    nodes.push({
      kind: "text",
      x: align === "center" ? p.x - w / 2 : align === "left" ? p.x : p.x - w,
      y: p.y - (am < 10 || am > 350 ? fs * 1.4 : Math.abs(am - 180) < 10 ? 0 : fs * 0.7),
      w,
      h: fs * 1.4,
      text: cat,
      fontSize: fs,
      color: style.text,
      align,
      valign: "middle",
      name: `category-${c}`,
    });
  });

  // Series legend (multi-series stacks) — shared wrapping row, same as above.
  if (legendH) {
    nodes.push(...legendRow(cfg, style, 0, titleH, { maxX: cfg.width - 4, entries: legendEntries }));
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => polar(cx, cy, r, angle(c) + sector / 2).x),
      categoryWidth: data.categories.map(() => (r - innerR) / 2),
      columnTop: data.categories.map((_, c) => polar(cx, cy, r, angle(c) + sector / 2).y),
      columnValue: catSum,
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}
