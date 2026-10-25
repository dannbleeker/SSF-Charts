import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, resolveFormat } from "../format";
import { maxOf, minOf } from "../agg";
import { lerpColor, noDataFill, sequentialScale } from "../color";
import { seriesColor } from "../style";
import { resolveTileLayout } from "./tilemap-layouts";
import {
  titleInkBottom,
  bandFontSize,
  fitPlot,
  footnoteH,
  legendRowCount,
  legendWrapWalk,
  MIN_LABEL_FS,
  titleHeight,
  titleNode,
} from "./frame";
import type { LayoutResult } from "./column";

/**
 * What the legend below the grid actually occupies, in font sizes — not a round
 * number close to it. It starts `fs * 0.5` under the grid, the swatch strip is
 * `fs * 0.9`, and the min/max labels sit at `fs * 0.95` below the strip's top in
 * a box `fs * 1.2` tall, so the last ink is `fs * 2.65` under the grid.
 *
 * Both the reservation above the legend and the clamp that keeps it inside the
 * frame are stated in terms of this, so they cannot drift apart: the reservation
 * carries an extra gutter of slack, and a clamp written against the reservation
 * would spend that slack and move every ordinary tilemap by half a point.
 */
const LEGEND_INK = 2.65;

/** Where a glyph legend that cannot wrap from the grid's left edge starts instead (the footnote's own margin). */
const LEGEND_MARGIN = 2;

/**
 * The smallest tile, in font sizes, that still draws its mini-bars. Below it the
 * glyph loop draws only the code, and in glyph mode the tile's fill is a faint
 * backdrop that carries no value — so the grid shows no data. Stated once
 * because the draw and the legend's row reservation both ask it.
 */
const GLYPH_MIN_TILE = 2;

/**
 * The glyph-mode series key: one colour chip and one name per series, laid out
 * by `legendWrapWalk` — the walk that already reserved its rows, so the drawer
 * and the reserver count the same ones. Its own function, not a block in
 * `layoutTilemap`, because that function is held to its recorded size
 * (`test/function-size.test.ts`) and a legend that wraps is a self-contained
 * thing to draw: it needs only the series, where it starts and where it stops.
 */
function glyphLegendNodes(
  series: { name: string; color?: string }[],
  style: ChartStyle,
  fs: number,
  startX: number,
  maxX: number,
  top: number,
  rowPitch: number,
): SceneNode[] {
  const chip = fs * 0.7;
  const slots = legendWrapWalk(
    series.map((s) => s.name),
    fs,
    startX,
    maxX,
  );
  const out: SceneNode[] = [];
  series.forEach((s, si) => {
    const { x: lx, row } = slots[si];
    const y = top + row * rowPitch;
    out.push(
      {
        kind: "rect",
        x: lx,
        y: y + fs * 0.6,
        w: chip,
        h: chip,
        fill: seriesColor(style, si, s.color),
        name: `legend-chip-${si}`,
      },
      {
        kind: "text",
        x: lx + chip + 3,
        y: y + fs * 0.3,
        w: textWidth(s.name, fs) + 6,
        h: fs * 1.4,
        text: s.name,
        fontSize: fs * 0.85,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `legend-${si}`,
      },
    );
  });
  return out;
}

/**
 * A hex tile's width and height as a fraction of its grid cell.
 *
 * √3/2, and it is inherited rather than chosen: the tile used to be a pointy-top
 * regular hexagon of circumradius `tile/2`, which spans `tile` tall by
 * `0.866·tile` wide. The whole hex grid — the `cols + 0.5` width, the 0.87·tile
 * row step, the half-cell odd-row shift, and the "tiles never touch" property
 * all of it is fitted around — was derived from THAT width, so the flat-top
 * preset that replaced it takes the same 0.866·tile as its box side and none of
 * the fit math moves.
 *
 * At the cell's full `tile` the flat top and bottom reach the neighbouring row:
 * rows step 0.87·tile, and two hexes an odd row apart sit only half a cell over
 * from each other. The pointy-top hexagon got away with a full-cell HEIGHT
 * because its narrow axis faced the shift, which is horizontal. Turning the
 * shape a quarter turn swaps which axis is narrow, so the size has to follow —
 * `no two hex tiles overlap` in test/tilemap.test.ts is the guard, and it fails
 * on ME/NH the moment this becomes 1.
 */
const HEX_W = Math.sqrt(3) / 2;

/**
 * Tile-grid cartogram ("map chart"): every region is a uniform square, so
 * geography stays recognizable without area distortion — and it renders as
 * native shapes everywhere (no freeform paths needed). Categories are region
 * codes (US postal / ISO-2 / world macro-regions); the single series' values
 * color the tiles on a sequential scale. Regions in the layout without data
 * render as gray "no data" tiles.
 */
export function layoutTilemap(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const layout = resolveTileLayout(cfg.map, data.categories);

  const nodes: SceneNode[] = [];
  const titleH = titleHeight(cfg, style);
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  const empty: LayoutResult = {
    nodes,
    anchors: {
      categoryX: data.categories.map(() => cfg.width / 2),
      categoryWidth: data.categories.map(() => 10),
      columnTop: data.categories.map(() => titleH),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: cfg.height,
      // Fitted because the decorations read this rect: a frame shorter than its
      // own title would otherwise hand them a negative height to map bands
      // through.
      plot: fitPlot(cfg, { x: 0, y: titleH, w: cfg.width, h: cfg.height - titleH }),
    },
  };
  if (!layout) {
    nodes.push({
      kind: "text",
      x: 0,
      y: titleH + fs,
      w: cfg.width,
      h: fs * 1.5,
      text: 'No recognized region codes — set map: "us" | "eu" | "europe" | "world"',
      fontSize: fs,
      color: style.mutedText,
      align: "center",
      valign: "middle",
      name: "tilemap-error",
    });
    return empty;
  }

  // Value per region code (uppercased), from the first series.
  const values = new Map<string, number>();
  data.categories.forEach((code, c) => {
    const v = data.series[0]?.values[c];
    if (v != null && code.trim().toUpperCase() in layout) values.set(code.trim().toUpperCase(), v);
  });
  const vals = [...values.values()];
  const min = minOf(vals);
  const max = maxOf(vals);
  const base = data.series[0]?.color ?? style.palette[0];
  // Constant data: a flat mid-tone beats a zero-width ramp.
  const fill =
    min === max ? () => lerpColor(style.background, base, 0.5) : sequentialScale(min, max, base, style.background);
  const fmt = resolveFormat(vals, cfg.numberFormat);

  const topts = cfg.tilemap ?? {};
  const hex = topts.shape === "hex";
  const glyph = topts.glyph === "bars" && data.series.length > 1;
  // Mini-glyph mode: each region carries a series of values (one bar each).
  const seriesVals = new Map<string, (number | null)[]>();
  if (glyph) {
    data.categories.forEach((code, c) => {
      const key = code.trim().toUpperCase();
      if (key in layout)
        seriesVals.set(
          key,
          data.series.map((s) => s.values[c] ?? null),
        );
    });
  }
  const glyphVals = glyph ? [...seriesVals.values()].flat().filter((v): v is number => v != null) : [];
  // Floored at 0, not 1: a floor of 1 in data units made a map of shares typed as
  // 0.05-0.4 draw its tallest mini-bar at 40% of the room the same numbers x100 get
  // (the unit set the picture; butterfly, radar and cascade lost the same floor in
  // 47aa582). The all-zero guard is the span's `|| 1` below, where it divides.
  const glyphMax = maxOf(glyphVals, 0);
  // The mini-bar scale spans zero: a negative value draws BELOW the tile's zero
  // line instead of clamping to zero height, which was pixel-identical to "no
  // change". All-positive data keeps glyphMin = 0, i.e. the original scale.
  const glyphMin = Math.min(0, minOf(glyphVals, 0));
  const glyphSpan = glyphMax - glyphMin || 1;

  // Fit uniform square tiles into the plot area.
  const cols = Math.max(...Object.values(layout).map(([c]) => c)) + 1;
  const rows = Math.max(...Object.values(layout).map(([, r]) => r)) + 1;
  const gutter = 2.5;
  // The reservation is the legend's own ink (see LEGEND_INK) plus one gutter of
  // slack, because `rowsBottom` adds a gutter per row where the height budget
  // pays for the gaps BETWEEN rows — one gutter more than it is given. Those two
  // together used to predict `fs * 0.0785 + 0.5` of overflow, and did: the min
  // and max labels had their descenders cut by the frame on EVERY tilemap at
  // EVERY size, 1.3pt at a 10pt font and 1.9 at 18, the measurement to the tenth.
  const oneRowLegendH = vals.length ? LEGEND_INK * fs + gutter : fs * 0.5;
  const availW = cfg.width - 4;
  // The tile that fits once `legendH` is taken off the height.
  //
  // Hex tiles nest: rows step ~0.87·tile and odd rows shift half a column, so
  // the footprint needs an extra half column of width and less height.
  // The 6pt floor keeps a tile visible on an ordinary chart, and it is the last
  // place in this engine that refused to yield to its frame: on a frame too
  // small to pay for it, honouring the floor made the grid taller than the space
  // budgeted for it and pushed the legend under it off the bottom — 27pt on a
  // 120x90 thumbnail. A small tile is still legible as a SHAPE, which is what a
  // cartogram is read for; a tile drawn outside the frame is not there at all.
  //
  // So the floor applies only when the frame can pay for it. It is unreachable
  // on any chart big enough to want it, so nothing of an ordinary size moves.
  const fitTile = (legendH: number) => {
    const availH = cfg.height - titleH - legendH - footnoteH(cfg, style, decor) - 4;
    const want = hex
      ? Math.min((availW - (cols - 1) * gutter) / (cols + 0.5), availH / ((rows - 1) * 0.87 + 1))
      : Math.min((availW - (cols - 1) * gutter) / cols, (availH - (rows - 1) * gutter) / rows);
    return Math.max(1, want);
  };
  /**
   * The glyph legend WRAPS, and the rows it needs are reserved before the grid is
   * fitted — the same walk (`legendWrapWalk`) the boxplot, radar, scatter and
   * butterfly legends take, so the reserver and the drawer count the same rows.
   *
   * It advanced a chip, a gap and a measured name per series along ONE row from
   * the grid's left edge and never asked whether the next fit: twelve monthly
   * series on the default Europe map put "Dec" at x=493 on a 480-wide chart, and
   * six long names put the last chip at x=651. `clipTextToFrame` only reads
   * text, so the names were clipped or dropped while the colour chips — rects —
   * stayed where they were, across whatever sits beside the chart on the slide.
   *
   * Counted from the left MARGIN, which is the most room the legend can have,
   * because the grid's own left edge depends on the tile and the tile depends on
   * what is reserved here. The draw hangs it from the grid when that costs no
   * more rows than this, so a legend that already fit keeps its place.
   *
   * Each row past the first costs one `legendRow` pitch (`fs * 1.6`), on top of
   * the band every tilemap already reserves, which holds one row of it. Those
   * rows are taken only while the tile still reaches `GLYPH_MIN_TILE`: below it
   * the bars are not drawn at all (see the glyph loop), so a key that costs the
   * grid its bars is naming nothing. Past that the key is DROPPED — "where there
   * is no room for both, the key is the one that goes", as the title clearance
   * below has it — rather than squeezing the grid to the tile floor for a legend
   * that has nothing left to explain.
   */
  const legendNames = glyph ? data.series.map((s) => s.name) : [];
  const wantRows = legendRowCount(legendNames, fs, LEGEND_MARGIN, cfg.width - 4);
  const rowPitch = fs * 1.6;
  const glyphLegendRows =
    wantRows <= 1 || fitTile(oneRowLegendH + (wantRows - 1) * rowPitch) >= fs * GLYPH_MIN_TILE ? wantRows : 0;
  const extraLegendH = Math.max(0, glyphLegendRows - 1) * rowPitch;
  const legendH = oneRowLegendH + extraLegendH;
  const tile = fitTile(legendH);
  const gridW = (hex ? cols + 0.5 : cols) * tile + (cols - 1) * gutter;
  const x0 = (cfg.width - gridW) / 2;
  // The grid's own height, and a top edge that keeps it on the canvas. The tile
  // floor above is deliberately allowed to overrun the height budget, and on a
  // frame whose title and legend leave nothing it overran the CANVAS too: the
  // bottom row was drawn up to 24.7pt below the foot of a 300x60 chart. Pushed
  // up rather than shrunk further — the same call the comment above makes, since
  // a tile over the title is still a tile and one off the chart is not there.
  const gridH = hex ? (rows - 1) * tile * 0.87 + tile : rows * (tile + gutter);
  const y0 = Math.max(0, Math.min(titleH + 2, cfg.height - gridH));
  const rowsBottom = y0 + gridH;
  /**
   * Where the legend hangs from. Normally the bottom of the grid — but the tile
   * floor above is explicitly allowed to overrun the height budget, and when it does the
   * grid drags the legend off the frame with it: 17pt below a 300x60 chart. So
   * the legend takes the grid's bottom or the last position its own reserved
   * band still fits at, whichever is higher. On any chart whose grid fits its
   * budget the first is always the smaller, so nothing of an ordinary size moves.
   */
  const legendTop = Math.max(
    0,
    Math.min(rowsBottom, cfg.height - LEGEND_INK * fs - extraLegendH - footnoteH(cfg, style, decor)),
  );
  /**
   * …and only where that position is clear of the TITLE.
   *
   * The clamp above keeps the legend on the canvas, and on a 300x60 frame at
   * 18pt the only position left on the canvas is inside the title's band — so
   * the "no data" caption was drawn across the chart's own title. A colour key
   * explains the tiles; the title says what the tiles are. Where there is no
   * room for both, the key is the one that goes, which is what every other
   * reservation in this engine does when it cannot be met.
   */
  const legendClearOfTitle = legendTop >= titleInkBottom(cfg, style);

  for (const [code, [col, row]] of Object.entries(layout)) {
    const v = values.get(code);
    // In glyph mode the tile is a faint backdrop for the bars; otherwise it
    // carries the value color.
    const tileFill = glyph
      ? lerpColor(style.background, base, 0.1)
      : v == null || !vals.length
        ? noDataFill(style.background)
        : fill(v);
    const x = x0 + col * (tile + gutter) + (hex && row % 2 === 1 ? (tile + gutter) / 2 : 0);
    const y = hex ? y0 + row * tile * 0.87 : y0 + row * (tile + gutter);
    if (hex) {
      // A SYMBOL, not a polygon, and that is the whole fix for this chart.
      //
      // Office.js has no freeform fill, so a PolygonNode degrades to its stroked
      // outline (scene.ts's parity contract). A choropleth says what it says
      // through FILL — 51 tiles carrying 16 distinct colours — so the add-in
      // drew the one chart whose entire message is fill as 51 hollow rings while
      // the SVG preview beside it drew them solid. The preview did not
      // approximate the slide; it disagreed with it. The old workaround here set
      // stroke = fill so the outline at least carried the colour, which stopped
      // the cartogram being white-on-white and left it hollow.
      //
      // SymbolNode exists for exactly this: it names a native preset the host
      // fills itself. The cost is the tile's POINTY top — PowerPoint's `hexagon`
      // preset points left and right, and turning it needs `Shape.rotation`, API
      // 1.10, which no round has ever exercised and which most desktop builds do
      // not have at all. A flat-top hexagon is still a hexagon and still reads as
      // a cartogram; a hollow one is not a tile. It also drops the chart from 6
      // shapes a tile to 1 — 401 shapes to 146 — which is most of why this was
      // the heaviest chart we ship, and inside the 400-500 band that took
      // PowerPoint down on all seven attempts in round 150.
      nodes.push({
        kind: "symbol",
        shape: "hexagon",
        cx: x + tile / 2,
        cy: y + tile / 2,
        // HALF-WIDTH, not half-cell: see HEX_W. `size` is half a SQUARE box side
        // and the preset spans it corner to corner.
        size: (tile * HEX_W) / 2,
        fill: tileFill,
        name: `tile-${code}`,
      });
    } else {
      nodes.push({ kind: "rect", x, y, w: tile, h: tile, fill: tileFill, name: `tile-${code}` });
    }
    const ink = contrastInk(tileFill);
    // Mini bar glyph: one bar per series, from the region's row of values.
    if (glyph) {
      const svals = seriesVals.get(code);
      if (svals && tile >= fs * GLYPH_MIN_TILE) {
        const nb = svals.length;
        const bw = (tile * 0.78) / nb;
        const bx0 = x + tile * 0.11;
        const bBase = y + tile * 0.86;
        const bMax = tile * 0.5;
        const zeroY = bBase + (glyphMin / glyphSpan) * bMax;
        svals.forEach((sv, si) => {
          if (sv == null) return;
          const vy = bBase - ((sv - glyphMin) / glyphSpan) * bMax;
          nodes.push({
            kind: "rect",
            x: bx0 + si * bw,
            y: Math.min(vy, zeroY),
            w: Math.max(1, bw - 0.5),
            h: Math.abs(vy - zeroY),
            fill: seriesColor(style, si),
            name: `glyph-${code}-${si}`,
          });
        });
      }
      // A tile's code is bounded by the TILE, and a grid squeezed onto a sliver
      // of a frame gives tiles of a point or two — so this answered fonts below
      // anything readable, and below anything OOXML accepts. Dropped instead:
      // the tile still carries the value through its fill.
      const glyphCodeFs = bandFontSize(fs * 0.85, tile, 1 / 0.3);
      if (glyphCodeFs > 0)
        nodes.push({
          kind: "text",
          x,
          y: y + tile * 0.06,
          w: tile,
          h: fs * 1.2,
          text: code,
          fontSize: glyphCodeFs,
          bold: true,
          color: ink,
          align: "center",
          valign: "middle",
          name: `tile-code-${code}`,
        });
      continue;
    }
    const showValue =
      v != null && decor.segmentLabels && tile >= fs * 2.6 && textWidth(formatNumber(v, fmt), fs * 0.8) <= tile - 2;
    const codeFs = bandFontSize(fs, tile, 1 / 0.34);
    if (codeFs > 0)
      nodes.push({
        kind: "text",
        x,
        y: showValue ? y + tile / 2 - fs * 1.25 : y,
        w: tile,
        h: showValue ? fs * 1.3 : tile,
        text: code,
        fontSize: codeFs,
        bold: true,
        color: ink,
        align: "center",
        valign: "middle",
        name: `tile-code-${code}`,
      });
    if (showValue) {
      nodes.push({
        kind: "text",
        x,
        y: y + tile / 2,
        w: tile,
        h: fs * 1.2,
        text: formatNumber(v, fmt),
        fontSize: fs * 0.8,
        color: ink,
        align: "center",
        valign: "top",
        name: `tile-value-${code}`,
      });
    }
  }

  // Glyph mode: a series legend instead of the value gradient.
  if (glyph && glyphLegendRows > 0 && legendClearOfTitle) {
    // Hung from the grid's left edge, where it has always been, whenever that
    // costs no more rows than were reserved; from the margin otherwise.
    const fromGrid = legendRowCount(legendNames, fs, x0, cfg.width - 4) <= glyphLegendRows;
    nodes.push(
      ...glyphLegendNodes(data.series, style, fs, fromGrid ? x0 : LEGEND_MARGIN, cfg.width - 4, legendTop, rowPitch),
    );
  }
  // Gradient legend + "no data" swatch.
  if (!glyph && vals.length && min !== max && legendClearOfTitle) {
    const ly = legendTop + fs * 0.5;
    const lw = Math.min(gridW * 0.5, fs * 12);
    const steps = 24;
    for (let i = 0; i < steps; i++) {
      const v = min + ((max - min) * i) / (steps - 1);
      nodes.push({
        kind: "rect",
        x: x0 + (lw / steps) * i,
        y: ly,
        w: lw / steps + 0.5,
        h: fs * 0.9,
        fill: fill(v),
        name: `legend-step-${i}`,
      });
    }
    // The two ends of the scale each own HALF the bar and are anchored to its
    // outer edges, so their ink meets in the middle the moment a number is
    // wider than `lw / 2`. The bar is `min(gridW * 0.5, fs * 12)`, so on a
    // small frame it is a few points wide and the two numbers are drawn on top
    // of each other — inside the frame, so no overflow gate could see it.
    //
    // Same answer the rest of this engine gives: shrink both together to the
    // room they actually have, and drop the PAIR when that would be illegible.
    // Both, never one: a gradient bar labelled at one end says the wrong thing,
    // where an unlabelled bar just says less.
    const endText = [formatNumber(min, fmt), formatNumber(max, fmt)];
    const endFs = (() => {
      const base = fs * 0.85;
      const widest = Math.max(...endText.map((t) => textWidth(t, base)));
      if (!(widest > 0)) return base;
      const f = Math.min(base, base * (lw / 2 / widest));
      return f >= MIN_LABEL_FS ? f : 0;
    })();
    if (endFs > 0)
      nodes.push(
        {
          kind: "text",
          x: x0,
          y: ly + fs * 0.95,
          w: lw / 2,
          h: fs * 1.2,
          text: endText[0],
          fontSize: endFs,
          color: style.mutedText,
          align: "left",
          valign: "top",
          name: "legend-min",
        },
        {
          kind: "text",
          x: x0 + lw / 2,
          y: ly + fs * 0.95,
          w: lw / 2,
          h: fs * 1.2,
          text: endText[1],
          fontSize: endFs,
          color: style.mutedText,
          align: "right",
          valign: "top",
          name: "legend-max",
        },
      );
    if (values.size < Object.keys(layout).length) {
      // The no-data key is a swatch and a caption, and BOTH have to fit to the
      // right of the gradient bar. Bounding the caption's box was not enough:
      // its box was clamped to a point wide and clipped to the empty string,
      // and an empty text box still has an ORIGIN — 27pt past the right edge of
      // an 80x60 chart at 24pt, with the swatch itself hanging 5pt over. A node
      // that draws nothing is still a shape on the slide, and the overflow gate
      // reads it as one.
      //
      // So this is the reservation answer the radar, sunburst, pie and the
      // scale ends above already give: shrink the caption to the room it has,
      // and where that room cannot carry a readable one, DROP THE PAIR. A
      // swatch with no caption says nothing; the tiles it explains are still
      // gray, and the reader loses a legend entry rather than gaining ink
      // outside the chart.
      const swatchX = x0 + lw + fs;
      const textX = x0 + lw + fs * 2.1;
      const room = cfg.width - textX - 2;
      const capFs = (() => {
        const base = fs * 0.85;
        const w = textWidth("no data", base);
        if (!(w > 0)) return base;
        const f = Math.min(base, base * (room / w));
        return f >= MIN_LABEL_FS ? f : 0;
      })();
      if (capFs > 0 && swatchX + fs * 0.9 <= cfg.width)
        nodes.push(
          {
            kind: "rect",
            x: swatchX,
            y: ly,
            w: fs * 0.9,
            h: fs * 0.9,
            fill: noDataFill(style.background),
            name: "legend-nodata",
          },
          {
            kind: "text",
            x: textX,
            y: ly - fs * 0.2,
            // Bounded by the room actually left to the right of the swatch. The
            // flat `fs * 6` is 192 points wide at a 32pt font and starts 67
            // points in, so on a narrow chart the caption ran 37pt off the right
            // edge. Last resort in both directions: where the flat box fits it
            // is still the flat box, so no chart of an ordinary size moves.
            w: Math.min(fs * 6, room),
            h: fs * 1.3,
            text: clipToWidth("no data", capFs, Math.min(fs * 6, room)),
            fontSize: capFs,
            color: style.mutedText,
            align: "left",
            valign: "middle",
            name: "legend-nodata-label",
          },
        );
    }
  }

  return {
    ...empty,
    nodes,
  };
}
