/**
 * think-cell's non-chart elements: harvey balls, checkboxes, process flows,
 * and a simple table — all as scenes for the same renderers as charts.
 */
import type { Scene, SceneNode } from "./scene";
import { contrastInk, ellipsize, finiteNodes, textWidth } from "./scene";
import { DEFAULT_STYLE, PALETTE } from "./style";

const S = DEFAULT_STYLE;

/** The smallest a table cell is drawn at before its text is ellipsized instead. */
const MIN_TABLE_FS = 6;

/**
 * Shorten `text` with an ellipsis until it fits `maxW`. Neither PowerPoint
 * renderer wraps or clips a text box (wrap/wordWrap are off), so a label wider
 * than its shape is drawn straight over its neighbours — the last resort once
 * shrinking the font has hit its floor. Shared with the agenda slide.
 */
export function clipToWidth(text: string, fs: number, maxW: number, bold = false): string {
  // Coerced for the reason `textWidth` gives just below it: the type says
  // `string` and the value came out of a file someone pasted. A NUMBER made
  // this quietly stop clipping — `(2024).length` is `undefined`, `undefined > 0`
  // is false, so the walk never ran and the function returned its input with an
  // ellipsis STUCK ON THE END, wider than what it was asked to fit. A KPI tile's
  // value and a table cell are both routinely numeric.
  const t = String(text ?? "");
  if (textWidth(t, fs, bold) <= maxW) return t;
  // Through the SHARED walk — see `ellipsize` in scene.ts. This was its own
  // copy of the same loop, and the two drifted the moment one of them learned
  // not to cut a surrogate pair in half.
  return ellipsize(t, fs, bold, maxW);
}

/**
 * An arrowhead node's (x, y) is its TIP, and its body extends 1.8*size back
 * along `angle` (see the parity contract in scene.ts). Passing the text's
 * centre line therefore hangs the whole triangle to one side of it — an [up]
 * row and a [down] row end up 1.8*size apart. Convert the glyph's intended
 * visual centre into the tip the node wants.
 */
function arrowheadTip(cx: number, cy: number, angle: number, size: number): { x: number; y: number } {
  const rad = (angle * Math.PI) / 180;
  return { x: cx + Math.cos(rad) * size * 0.9, y: cy + Math.sin(rad) * size * 0.9 };
}

/** Harvey ball: fraction filled clockwise from 12 o'clock (0..1). */
export function buildHarveyBall(fraction: number, size = 24): Scene {
  const f = Math.max(0, Math.min(1, fraction));
  const r = size / 2 - 1.5;
  const cx = size / 2;
  const cy = size / 2;
  const nodes: SceneNode[] = [
    { kind: "ellipse", cx, cy, rx: r, ry: r, fill: "#ffffff", stroke: S.text, strokeWidth: 1.25, name: "harvey-ring" },
  ];
  if (f >= 1) {
    nodes.push({ kind: "ellipse", cx, cy, rx: r - 1.5, ry: r - 1.5, fill: S.text, name: "harvey-fill" });
  } else if (f > 0) {
    nodes.push({
      kind: "wedge",
      cx,
      cy,
      r: r - 1.5,
      innerR: 0,
      startAngle: 0,
      endAngle: f * 360,
      fill: S.text,
      name: "harvey-fill",
    });
  }
  // A TEXT ALTERNATIVE, because `sceneToSvg` marks every scene `role="img"`.
  // Its own comment says why that matters: "role=img with no name is exactly
  // what axe-core's role-img-alt reports as a 1.1.1 failure". `buildChart` sets
  // `desc` on every chart and `test/a11y-svg.test.ts` pins it — the element
  // builders were simply never given the same treatment, so the four Elements
  // previews shipped unnamed. Measured against the live pane 2026-09-10.
  return { width: size, height: size, desc: `Harvey ball, ${Math.round(f * 100)}% filled.`, nodes: finiteNodes(nodes) };
}

export type CheckState = "yes" | "no" | "partial";

/** Checkbox / status mark: ✓ (good), ✗ (critical), − (neutral). */
export function buildCheckbox(state: CheckState, size = 20): Scene {
  const colors: Record<CheckState, string> = { yes: "#0ca30c", no: "#d03b3b", partial: "#898781" };
  const glyph: Record<CheckState, string> = { yes: "✓", no: "✗", partial: "–" };
  // A state that is not a state must not reach `Object.prototype`.
  //
  // `CheckState` is a union in the types and this function is exported from
  // `src/index.ts`, so the argument is whatever a caller passed —
  // `buildCheckbox("__proto__")` put Object.prototype into BOTH the glyph and
  // the stroke, and `"constructor"` put a FUNCTION there; both reached the
  // markup. This is the seventh table in this repo to need the same guard, and
  // CLAUDE.md's list of the first six did not have this file on it.
  //
  // Falls back to `partial`, the neutral mark, rather than throwing: an element
  // builder that refuses is worse for a caller than one that draws the
  // undecided glyph.
  const known: CheckState = Object.prototype.hasOwnProperty.call(glyph, state) ? state : "partial";
  // See `buildHarveyBall` for why every element scene carries a `desc`. Named
  // from the RESOLVED state, not the argument: a caller passing nonsense gets
  // the neutral glyph above, and the text alternative has to agree with what
  // was actually drawn rather than with what was asked for.
  const said: Record<CheckState, string> = { yes: "yes", no: "no", partial: "partial" };
  return {
    width: size,
    height: size,
    desc: `Checkbox, ${said[known]}.`,
    nodes: [
      {
        kind: "rect",
        x: 0.5,
        y: 0.5,
        w: size - 1,
        h: size - 1,
        fill: "#ffffff",
        stroke: colors[known],
        strokeWidth: 1.5,
        name: "check-box",
      },
      {
        kind: "text",
        x: 0,
        y: 0,
        w: size,
        h: size,
        text: glyph[known],
        fontSize: size * 0.62,
        bold: true,
        color: colors[known],
        align: "center",
        valign: "middle",
        name: "check-glyph",
      },
    ],
  };
}

/** Process flow: a row of chevrons with the active step highlighted. */
export function buildProcessFlow(steps: string[], highlight = -1, width = 480, height = 40): Scene {
  // A `string[]` in the types is not an array in the object someone passed.
  // These builders are public API (`src/index.ts` exports all five), so a
  // caller with no steps, or with a config field that came back undefined, gets
  // an empty flow rather than "Cannot read properties of undefined (reading
  // 'length')" out of a render. Same rule the config boundary already follows.
  const list: string[] = Array.isArray(steps) ? steps.map((s) => String(s ?? "")) : [];
  const n = Math.max(1, list.length);
  const overlap = height * 0.28; // chevron notch overlaps the previous step
  const stepW = (width + overlap * (n - 1)) / n;
  const labelW = stepW - overlap * 1.6; // the flat part of the chevron
  // Shrink the labels to fit their own chevron (as buildKpiTile does for its
  // number): at a fixed 11pt a crowded flow drew each label over its neighbour
  // and the first one off the left edge of the scene.
  let fontSize = Math.min(11, height * 0.3);
  const overflows = (f: number) => list.some((s, i) => textWidth(s, f, i === highlight) > labelW);
  while (fontSize > 6 && overflows(fontSize)) fontSize -= 0.5;
  const nodes: SceneNode[] = [];
  list.forEach((label, i) => {
    const x = i * (stepW - overlap);
    const active = i === highlight;
    const fill = active ? PALETTE[0] : "#dcdbd2";
    nodes.push(
      { kind: "chevron", x, y: 0, w: stepW, h: height, fill, flatLeft: i === 0, name: `step-${i}` },
      {
        kind: "text",
        x: x + overlap * 0.8,
        y: 0,
        w: labelW,
        h: height,
        text: clipToWidth(label, fontSize, labelW, active),
        fontSize,
        bold: active,
        color: contrastInk(fill),
        align: "center",
        valign: "middle",
        name: `step-label-${i}`,
      },
    );
  });
  // See `buildHarveyBall`. Built from `list`, the SANITISED steps, so a caller
  // that passed a non-array gets "Process flow with no steps." rather than a
  // description of something that was not drawn — and the highlight is named
  // only when it points at a step that exists.
  const flowDesc = list.length
    ? `Process flow: ${list.join(", ")}.` +
      (highlight >= 0 && highlight < list.length ? ` ${list[highlight]} highlighted.` : "")
    : "Process flow with no steps.";
  return { width, height, desc: flowDesc, nodes: finiteNodes(nodes) };
}

export interface KpiTileOptions {
  /** Small muted caption above the number, e.g. "Revenue". */
  label?: string;
  /** The big number, preformatted ("€4.2m", "87 NPS"). */
  value: string;
  /** Delta line, e.g. "+12% vs LY". A leading +/− picks the arrow. */
  delta?: string;
  /** Arrow direction override; default inferred from the delta's sign. */
  direction?: "up" | "down" | "flat";
  /** Whether up is the good direction (colors the delta). Default true. */
  goodIsUp?: boolean;
}

/**
 * KPI / number tile: big number + delta arrow + caption — the dashboard
 * scorecard element. The delta arrow and text take semantic colors
 * (good/bad) from the direction and `goodIsUp`.
 */
export function buildKpiTile(opts: KpiTileOptions, width = 160, height = 90): Scene {
  // Coerced on the way in, like every other text boundary here: `delta` is typed
  // `string` and arrives as JSON, and a bare `delta: 12` reached `.replace` and
  // threw. `RegExp.test` coerces silently, so the direction was inferred fine
  // and only the render step fell over — which is why it needs fixing once here
  // rather than at the one call that happened to be strict.
  const delta = opts.delta == null || opts.delta === "" ? undefined : String(opts.delta);
  const dir =
    opts.direction ?? (delta ? (/^\s*[-−▼]/.test(delta) ? "down" : /^\s*[+▲]/.test(delta) ? "up" : "flat") : undefined);
  const goodIsUp = opts.goodIsUp ?? true;
  const deltaColor = dir === "flat" || dir == null ? S.mutedText : (dir === "up") === goodIsUp ? "#0ca30c" : "#d03b3b";

  const pad = 10;
  const labelFs = 10;
  // Shrink the big number to fit the tile width.
  let valueFs = Math.min(26, height * 0.34);
  while (valueFs > 11 && textWidth(opts.value, valueFs) > width - pad * 2) valueFs -= 1;

  const nodes: SceneNode[] = [
    {
      kind: "rect",
      x: 0.5,
      y: 0.5,
      w: width - 1,
      h: height - 1,
      fill: "#ffffff",
      stroke: "#e1e0d9",
      strokeWidth: 1,
      name: "kpi-box",
    },
  ];
  let y = pad;
  if (opts.label) {
    nodes.push({
      kind: "text",
      x: pad,
      y,
      w: width - pad * 2,
      h: labelFs * 1.3,
      text: clipToWidth(opts.label, labelFs, width - pad * 2),
      fontSize: labelFs,
      color: S.mutedText,
      align: "left",
      valign: "top",
      name: "kpi-label",
    });
    y += labelFs * 1.5;
  }
  nodes.push({
    kind: "text",
    x: pad,
    y,
    w: width - pad * 2,
    h: valueFs * 1.25,
    // Shrunk above, then ellipsized here: the shrink stops at 11pt, because a
    // KPI number smaller than that is not the thing this tile exists to show —
    // so a long value ran straight off the tile and over whatever sat beside it
    // on the slide (47pt past a 160pt tile), neither renderer clipping a text
    // box. Same shrink-then-clip order as the process flow and the table.
    text: clipToWidth(opts.value, valueFs, width - pad * 2, true),
    fontSize: valueFs,
    bold: true,
    color: S.text,
    align: "left",
    valign: "top",
    name: "kpi-value",
  });
  if (delta) {
    const dy = height - pad - labelFs * 0.75;
    let dx = pad;
    if (dir && dir !== "flat") {
      const size = labelFs * 0.45;
      const angle = dir === "up" ? -90 : 90;
      // dy is the delta text's centre line; centre the glyph on it too.
      const tip = arrowheadTip(dx + size, dy, angle, size);
      nodes.push({
        kind: "arrowhead",
        x: tip.x,
        y: tip.y,
        angle,
        size,
        fill: deltaColor,
        name: "kpi-arrow",
      });
      dx += size * 2 + 4;
    }
    // Strip a leading arrow glyph — the arrowhead node already shows it.
    const text = delta.replace(/^\s*[▲▼]\s*/, "");
    nodes.push({
      kind: "text",
      x: dx,
      y: dy - labelFs * 0.75,
      w: width - dx - pad,
      h: labelFs * 1.5,
      text: clipToWidth(text, labelFs, width - dx - pad, true),
      fontSize: labelFs,
      bold: true,
      color: deltaColor,
      align: "left",
      valign: "middle",
      name: "kpi-delta",
    });
  }
  // See `buildHarveyBall`. Assembled from the COERCED locals above rather than
  // from `opts`, so a `delta` that arrived as a number is described the way it
  // was actually drawn. The value carries its own units ("€4.2m"), so it is
  // read out as-is.
  const kpiDesc = [
    opts.label ? `${String(opts.label)}:` : "KPI:",
    String(opts.value ?? ""),
    delta ? `${delta}${dir && dir !== "flat" ? ` (${dir})` : ""}` : "",
  ]
    .filter(Boolean)
    .join(" ");
  return { width, height, desc: `${kpiDesc}.`, nodes: finiteNodes(nodes) };
}

export interface TableOptions {
  /**
   * "rules" (default): horizontal rules only — top, under the header, and
   * bottom; never side borders, no lines between data rows. "grid": legacy
   * full grid with accent header and zebra rows.
   */
  style?: "rules" | "grid";
  /** Treat the last row as a totals row: bold, with a rule above and below. */
  totalRow?: boolean;
  /**
   * Insert a small separator gap after every N body rows so long tables stay
   * scannable (default 5; 0 disables).
   */
  groupRows?: number;
}

/** Semantic colors for in-cell effects. */
const GOOD = "#0ca30c";
const BAD = "#d03b3b";

/**
 * In-cell effects, conditional-formatting style: leading tokens turn into
 * mini harvey balls, trend arrows, or semantic font colors.
 *   "[hb:0.75] Progress"  → ◕ Progress
 *   "[up] +14%" / "[down] -3%" / "[flat] 0%"
 *   "[good] on track" / "[bad] at risk"
 */
function parseCell(raw: string): { text: string; harvey?: number; trend?: "up" | "down" | "flat"; color?: string } {
  // Coerced for the same reason `xmlText` and `textWidth` are: `cells` is typed
  // `string[][]` and arrives as JSON — from the skill's caller, from a template,
  // from the pane's datasheet. A cell holding the NUMBER 2024 is not exotic, it
  // is a year in a table, and `(2024).match` threw `t.match is not a function`
  // and took the whole element down. The `?? ""` at the call site caught null
  // and undefined and nothing else.
  const text0 = String(raw ?? "");
  let text = text0;
  const out: ReturnType<typeof parseCell> = { text: text0 };
  let m: RegExpMatchArray | null;
  while ((m = text.match(/^\s*\[(hb:([\d.]+%?)|up|down|flat|good|bad)\]\s*/i))) {
    const tok = m[1].toLowerCase();
    if (tok.startsWith("hb:")) {
      const v = parseFloat(m[2]);
      // A bare "." parses to NaN; clamping NaN leaks it into the harvey fraction
      // (an empty ring with reserved space). Ignore a non-finite value instead.
      if (Number.isFinite(v)) out.harvey = Math.max(0, Math.min(1, m[2].includes("%") ? v / 100 : v));
    } else if (tok === "up" || tok === "down" || tok === "flat") out.trend = tok;
    else out.color = tok === "good" ? GOOD : BAD;
    text = text.slice(m[0].length);
  }
  out.text = text;
  return out;
}

/** Mini harvey ball + trend arrow glyph nodes for a cell, left of the text. */
function cellEffectNodes(
  cell: ReturnType<typeof parseCell>,
  x: number,
  cy: number,
  fs: number,
  ri: number,
  c: number,
): SceneNode[] {
  const nodes: SceneNode[] = [];
  let gx = x;
  if (cell.harvey != null) {
    const r = fs * 0.5;
    nodes.push({
      kind: "ellipse",
      cx: gx + r,
      cy,
      rx: r,
      ry: r,
      fill: "#ffffff",
      stroke: S.text,
      strokeWidth: 1,
      name: `cell-hb-ring-${ri}-${c}`,
    });
    if (cell.harvey >= 1)
      nodes.push({ kind: "ellipse", cx: gx + r, cy, rx: r - 1, ry: r - 1, fill: S.text, name: `cell-hb-${ri}-${c}` });
    else if (cell.harvey > 0)
      nodes.push({
        kind: "wedge",
        cx: gx + r,
        cy,
        r: r - 1,
        innerR: 0,
        startAngle: 0,
        endAngle: cell.harvey * 360,
        fill: S.text,
        name: `cell-hb-${ri}-${c}`,
      });
    gx += r * 2 + 3;
  }
  if (cell.trend) {
    const size = fs * 0.42;
    const angle = cell.trend === "up" ? -90 : cell.trend === "down" ? 90 : 0;
    const fill = cell.trend === "up" ? GOOD : cell.trend === "down" ? BAD : S.mutedText;
    // Centred on the text line and in its reserved slot, like the harvey ball above.
    const tip = arrowheadTip(gx + size, cy, angle, size);
    nodes.push({ kind: "arrowhead", x: tip.x, y: tip.y, angle, size, fill, name: `cell-trend-${ri}-${c}` });
  }
  return nodes;
}

/** Effect glyph width reserved before the cell text. */
const effectW = (cell: ReturnType<typeof parseCell>, fs: number) =>
  (cell.harvey != null ? fs + 3 : 0) + (cell.trend ? fs * 0.84 + 3 : 0);

/**
 * Table element. Default style follows chart-design practice: rules at the
 * top, under the header, and at the bottom — never side borders or lines
 * between data rows — with a separator gap every 5 rows and an optional
 * bold totals row. Cells accept in-cell effect tokens (see parseCell).
 */
export function buildTableScene(cellsIn: string[][], width = 480, opts: TableOptions = {}): Scene {
  const styleMode = opts.style ?? "rules";
  const groupEvery = opts.groupRows ?? 5;
  // A ROW is as untrusted as a cell. `string[][]` is the type; the value is
  // JSON, and a row that came through as `null` (a blank line in a pasted
  // block) or as a bare string (one row written without its inner array) threw
  // out of `.map`/`.length` before any of the cell-level care below could run.
  // A non-array row reads as an empty row, which is what it looks like.
  const cells = (Array.isArray(cellsIn) ? cellsIn : []).map((r) => (Array.isArray(r) ? r : []));
  const rows = cells.length;
  // No rows: the closing rule would be drawn at y = -0.5, i.e. above the top of
  // a zero-height scene (and at a negative offset from the insertion point in
  // the PowerPoint renderers). Nothing to draw.
  // Still named, even with nothing in it: `sceneToSvg` stamps `role="img"` on
  // every scene it renders, including this one, and an empty img with no
  // accessible name is the same 1.1.1 failure as a full one.
  if (rows === 0) return { width, height: 0, desc: "Empty table.", nodes: [] };
  const cols = Math.max(1, ...cells.map((r) => r.length));
  const parsed = cells.map((row) => row.map((c) => parseCell(c)));
  const colWidths = (f: number) =>
    Array.from({ length: cols }, (_, c) =>
      Math.max(f * 3, ...parsed.map((r) => textWidth(r[c]?.text ?? "", f) + effectW(r[c] ?? { text: "" }, f) + 12)),
    );
  /**
   * The font every cell is drawn at, shrunk until the row of them fits.
   *
   * Column widths are proportional to their longest content and then SCALED to
   * the table's width, so a table whose content is wider than it is gets every
   * column squeezed while the text in them stayed at 10pt: the cells ran into
   * each other and the last one ran off the end — 685pt past a 480pt table for
   * one long sentence, drawn over whatever sits beside it on the slide, because
   * neither PowerPoint renderer clips a text box.
   *
   * Shrinking is the remedy that fits this layout rather than a generic one:
   * the widths are computed FROM the font, so a smaller font narrows the text
   * without narrowing the column it sits in. Bold costs more than regular, and
   * the header and total rows are bold, so the fit asks at the weight each cell
   * is actually drawn at.
   *
   * Floored at MIN_TABLE_FS and then ellipsized, which is the order every fit in
   * this engine uses — shrink while shrinking still buys readable text, clip
   * what is left. A table that already fits never enters the loop and is
   * byte-identical.
   */
  const fits = (f: number): boolean => {
    const ws = colWidths(f);
    const sc = width / (ws.reduce((a, b) => a + b, 0) || 1);
    return parsed.every((row, ri) =>
      row.every((cell, c) => {
        const bold = ri === 0 || (!!opts.totalRow && ri === rows - 1);
        return textWidth(cell.text, f, bold) + effectW(cell, f) <= ws[c] * sc - 10;
      }),
    );
  };
  let fs = 10;
  while (fs > MIN_TABLE_FS && !fits(fs)) fs -= 0.5;
  const rowH = fs * 2.1;
  const gapH = rowH * 0.4;
  // Column widths proportional to their longest content (incl. effect glyphs).
  const widths = colWidths(fs);
  const totalW = widths.reduce((a, b) => a + b, 0) || 1;
  const scale = width / totalW;
  const rules = styleMode === "rules";
  const rule = (y: number, weight: number, name: string): SceneNode => ({
    kind: "line",
    x1: 0,
    y1: y,
    x2: width,
    y2: y,
    stroke: S.text,
    strokeWidth: weight,
    name,
  });

  const nodes: SceneNode[] = [];
  let y = 0;
  cells.forEach((row, ri) => {
    const header = ri === 0;
    const total = !!opts.totalRow && ri === rows - 1;
    // Separator gap after every N body rows (rules style).
    if (rules && groupEvery > 0 && ri > 1 && !total && (ri - 1) % groupEvery === 0) y += gapH;
    if (rules && total) {
      y += gapH * 0.5;
      nodes.push(rule(y, 0.75, "rule-total"));
    }
    const fill = header ? PALETTE[0] : ri % 2 === 0 ? "#f4f3f0" : "#ffffff";
    let x = 0;
    for (let c = 0; c < cols; c++) {
      const w = widths[c] * scale;
      const cell = parsed[ri][c] ?? { text: "" };
      if (!rules) {
        nodes.push({
          kind: "rect",
          x,
          y,
          w,
          h: rowH,
          fill,
          stroke: "#e1e0d9",
          strokeWidth: 0.75,
          name: `cell-${ri}-${c}`,
        });
      }
      const ew = effectW(cell, fs);
      // Every column but the first is right-aligned, so its text sits at the
      // cell's right edge: anchoring the glyph at the left edge orphaned it
      // from its own value and parked it next to the previous column's number.
      const alignRight = c !== 0;
      const bold = header || total;
      const gx = alignRight ? Math.max(x + 5, x + w - 5 - textWidth(cell.text, fs, bold) - ew) : x + 5;
      nodes.push(...cellEffectNodes(cell, gx, y + rowH / 2, fs, ri, c));
      nodes.push({
        kind: "text",
        x: x + 5 + ew,
        y,
        w: w - 10 - ew,
        h: rowH,
        text: clipToWidth(cell.text, fs, w - 10 - ew, bold),
        fontSize: fs,
        bold,
        color: cell.color ?? (!rules && header ? contrastInk(fill) : S.text),
        align: alignRight ? "right" : "left",
        valign: "middle",
        name: `cell-text-${ri}-${c}`,
      });
      x += w;
    }
    y += rowH;
    if (rules && header) nodes.push(rule(y, 0.75, "rule-header"));
  });
  if (rules) {
    nodes.unshift(rule(0.5, 1.25, "rule-top"));
    nodes.push(rule(y - 0.5, 1.25, "rule-bottom"));
  }
  // See `buildHarveyBall`. Reuses `cols` from above rather than recounting, so
  // the description cannot disagree with the grid that was actually drawn — it
  // is already the widest row, which is what a ragged paste is drawn to.
  const tableDesc = `Table, ${rows} row${rows === 1 ? "" : "s"} by ${cols} column${cols === 1 ? "" : "s"}.`;
  return { width, height: y, desc: tableDesc, nodes: finiteNodes(nodes) };
}
