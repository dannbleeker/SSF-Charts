/**
 * Renderer-agnostic scene graph. Layouts emit these nodes; the SVG renderer
 * (preview/tests) and the Office.js renderer (native PowerPoint shapes)
 * consume them. Coordinates are in points, origin top-left.
 *
 * ── Renderer parity contract ────────────────────────────────────────────────
 * Three renderers consume this graph: SVG (src/render/svg.ts), Office.js
 * (src/render/powerpoint.ts, the live add-in), and PptxgenJS
 * (skill/scripts/render-pptx.mjs, the headless skill). SVG is the reference —
 * it can draw anything — so where the two PowerPoint renderers differ it is
 * because Office.js and OOXML presets cannot express what SVG can. The
 * divergences are intentional; each is noted on the field or kind it affects so
 * a later change does not "fix" an approximation into a regression:
 *
 *  - Pattern fills (rect.pattern): SVG only; solid elsewhere.
 *  - Polygon fills (polygon.fill/fillOpacity): SVG + pptx custGeom; Office.js
 *    has no freeform fill and degrades to the stroked outline.
 *  - Wedge geometry: SVG + pptx draw the exact arc; Office.js approximates with
 *    a triangle/rectangle fan (no adjustable pie geometry).
 *  - Dash arrays (line.dash): SVG honours the exact array; the PowerPoint
 *    renderers expose enums, so they map to the nearest native style via
 *    `dashKind` (dotted → roundDot/sysDot, else dash) rather than the exact rhythm.
 *    WHETHER a line is dashed at all is not approximate, and all three sinks must
 *    answer it the same way — they ask `dashKind`, which returns `none` for an
 *    array carrying no positive finite length. Do not re-guard on `dash` being
 *    truthy at a call site: `[]` is truthy, and that divergence drew a solid line
 *    in the preview and a dotted one in both decks.
 *  - TEXT alpha (text.color carrying one, e.g. `#0b0b0b80` or an `rgba()`):
 *    SVG and pptx honour it — pptxgenjs takes the same 0-100 transparency on a
 *    text run as on a shape — and Office.js cannot: `font.color` is a hex
 *    string with nowhere to put an alpha, so the live add-in draws such a label
 *    OPAQUE. Fills and strokes honour their alpha in all three, so this is the
 *    one paint channel that does not, and it is a host limit rather than a
 *    choice. Muted ink is how a chart de-emphasises a label, so a chart using
 *    it reads as flatter in the add-in than in the preview or the skill's deck.
 *  - Chevron point depth and arrowhead proportions: SVG draws its own geometry;
 *    the PowerPoint renderers name a native preset whose default proportions
 *    differ slightly (see the notes on those kinds). Reproducing the preset
 *    geometry exactly is not verifiable without a PowerPoint rasteriser, so the
 *    preview approximates a shape the deck draws natively — deliberately, the
 *    same call made for the rejected star5 marker.
 * ────────────────────────────────────────────────────────────────────────────
 */

export interface RectNode {
  kind: "rect";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  /** Hatch/dot pattern overlaid on the fill (SVG renderer; solid elsewhere). */
  pattern?: "diagonal" | "crosshatch" | "dots" | "horizontal";
  name?: string;
}

export interface PolygonNode {
  kind: "polygon";
  points: { x: number; y: number }[];
  /** Fill color; rendered translucent via fillOpacity in SVG. PowerPoint
   * renderers degrade to the stroked outline only (no freeform fills). */
  fill?: string;
  fillOpacity?: number;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

export interface LineNode {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  stroke: string;
  strokeWidth?: number;
  /**
   * Dash pattern in points, e.g. [2, 2]. SVG renders the exact array; the
   * PowerPoint renderers have only an enum of named styles, so they collapse it
   * to the nearest one via `dashKind` (a dotted [1.5,1.5] stays dotted;
   * everything else is a dash). The rhythm is approximate in the deck by design.
   *
   * An array with no positive finite length (`[]`, `[0, 0]`, `[-5, -5]`, `[NaN]`)
   * means NOT DASHED in every sink — ask `dashKind`, never `if (n.dash)`.
   */
  dash?: number[];
  name?: string;
}

export interface TextNode {
  kind: "text";
  /** Bounding box; alignment applies within it. */
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  fontSize: number;
  color: string;
  bold?: boolean;
  align: "left" | "center" | "right";
  valign: "top" | "middle" | "bottom";
  fontFamily?: string;
  name?: string;
}

export interface EllipseNode {
  kind: "ellipse";
  cx: number;
  cy: number;
  rx: number;
  ry: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

/**
 * Pie/doughnut wedge. Angles in degrees, 0 = 12 o'clock, clockwise.
 * SVG renders an exact path; PowerPoint approximates with a triangle fan
 * (Office.js exposes no adjustable pie geometry).
 */
export interface WedgeNode {
  kind: "wedge";
  cx: number;
  cy: number;
  r: number;
  /** Inner radius for doughnuts; 0 for pies. */
  innerR: number;
  startAngle: number;
  endAngle: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

/**
 * Process-flow chevron / pentagon-arrow (PowerPoint's chevron & homePlate).
 * SVG draws the arrow with its notch at a fixed fraction of the height; the
 * PowerPoint renderers name the native chevron/homePlate preset, whose own
 * default point depth differs slightly — so the arrow's point is a touch
 * deeper/shallower in the deck than the preview. Intentional (see the parity
 * contract at the top): the preset can't be matched pixel-for-pixel here.
 */
export interface ChevronNode {
  kind: "chevron";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  /** First step in a flow has a flat left edge (homePlate), the rest are chevrons. */
  flatLeft?: boolean;
  name?: string;
}

/**
 * Filled marker symbol centred on (cx, cy), inscribed in a `2*size` square.
 *
 * Shape is an encoding channel that survives what color does not: greyscale
 * printing and red-green color blindness both flatten a palette, and a deck
 * gets printed. A PolygonNode would render the same outline in SVG but
 * degrades to an unfilled outline in PowerPoint (no freeform fills there),
 * so a symbol is its own kind: each shape maps to a native preset geometry
 * and stays filled in all three renderers. See `symbolPoints` / `SYMBOL_PRESET`.
 */
export interface SymbolNode {
  kind: "symbol";
  shape: SymbolShape;
  cx: number;
  cy: number;
  /** Half the box side, so it reads like an ellipse's radius. */
  size: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

/**
 * Filled triangle with tip at (x, y), pointing along `angle` (degrees, 0 = east,
 * clockwise). SVG draws a narrow isosceles triangle; the PowerPoint renderers
 * name the native `triangle` preset in a `2*size` square (see `arrowheadBox`),
 * which is a touch broader. The tip anchor and angle match across all three;
 * only the triangle's proportions differ, intentionally (see the parity contract).
 */
export interface ArrowheadNode {
  kind: "arrowhead";
  x: number;
  y: number;
  angle: number;
  size: number;
  fill: string;
  name?: string;
}

export type SceneNode =
  RectNode | LineNode | TextNode | EllipseNode | WedgeNode | ChevronNode | ArrowheadNode | PolygonNode | SymbolNode;

// Circle/wedge math lives in ./geometry (shared with the renderers); re-exported
// here so scene consumers (layouts) keep importing `polar` from the scene module.
export { polar } from "./geometry";
import { wedgeFanSteps, type SymbolShape } from "./geometry";
import { toRgb } from "./color";
export type { SymbolShape };

/**
 * Drop any node whose geometry is not a finite number.
 *
 * The scene is the contract between the chart engine and three renderers, and
 * "every coordinate is a real number" was a property it happened to have
 * rather than one it promised. The SVG renderer defends itself — every numeric
 * goes through `num()` — but the two PowerPoint renderers do not, and they are
 * the ones that write a file. A `NaN` there lands in `addGeometricShape({left:
 * NaN, …})` and in OOXML as an EMU value, so the produced .pptx is one
 * PowerPoint may simply refuse to open. Nine chart kinds could do it, from
 * values a datasheet cell can hold (`1e308`, `5e-324`).
 *
 * DROPPED rather than zeroed. A node whose position could not be computed has
 * no right position to fall back to, and zeroing puts a stray bar in the
 * corner of the chart — wrong in a way that looks deliberate. Leaving it out
 * loses that one node and keeps the rest of the chart, which is what a reader
 * can actually interpret.
 *
 * For a valid config this drops nothing, so it is a floor and not a filter.
 */
export function finiteNodes(nodes: SceneNode[]): SceneNode[] {
  return nodes.filter((n) => allNumbersFinite(n) && !degeneratePolygon(n));
}

/**
 * Text clipped to what the chart can actually hold — the backstop under every
 * layout's own fitting.
 *
 * A label is drawn at whatever size its layout chose, in a box that layout sized
 * from the frame, and neither PowerPoint renderer wraps or clips a text box. So
 * any label wider than the room in front of it draws straight off the chart:
 * invisible in a picture-mode render, and lying across whatever sits beside the
 * chart on a slide.
 *
 * At a thumbnail frame that was not a corner case but the normal outcome — 18 of
 * the 25 kinds put ink outside their own frame at 120x90, by as much as 124pt on
 * a 120pt-wide chart. Most of it came from a handful of SHARED nodes (the title,
 * the footnote, the series labels) and each is now fitted where it is built,
 * which is better than clipping because shrinking keeps the whole word. This
 * catches what those did not, once, instead of in twenty-five layouts — the
 * per-site fixes stop being a list somebody has to finish.
 *
 * Only the horizontal axis: a label too TALL for its frame cannot be rescued by
 * shortening it, and the layouts that had that problem now reserve for it.
 *
 * Clipped from the anchor the node was placed by, so alignment is preserved: a
 * left-aligned label keeps its left edge, a right-aligned one its right, a
 * centred one its centre. A label already inside the frame is returned
 * untouched and byte-identical, which is why no snapshot moves.
 */
export function clipTextToFrame<T extends SceneNode>(nodes: T[], width: number): T[] {
  /**
   * A label the clip empties is DROPPED, not left in place carrying "".
   *
   * An empty text node draws no ink and is still a shape: it keeps the origin
   * the layout gave it, and the only labels this walk can empty are the ones
   * whose anchor is already outside the frame — 5pt left of a 60x300 doughnut,
   * 27pt right of an 80x60 tilemap. So the pass that exists to keep ink inside
   * the chart was itself the last thing putting a node outside it, and the
   * overflow gate reads that node exactly as it reads any other.
   */
  const dropped = new Set<T>();
  for (const n of nodes) {
    if (n.kind !== "text") continue;
    const t = n as unknown as TextNode;
    // A label with no text at all goes the same way, whoever emptied it. The
    // layouts clip to their own marks before this runs, so a category name on a
    // narrow chart arrives here already reduced to "" — five of them on an
    // 80x60 line chart, 203 across the kind/frame/font sweep. Each is a text
    // box PowerPoint has to create, on a host where the cost of a draw grows
    // with the shapes already on the slide.
    if (!t.text) {
      dropped.add(n);
      continue;
    }
    const ink = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - ink : t.align === "center" ? t.x + (t.w - ink) / 2 : t.x;
    if (x >= -0.5 && x + ink <= width + 0.5) continue;
    // The room in front of the anchor this node was positioned by. A centred
    // label may only grow to twice its distance from the nearer edge before one
    // side leaves the frame.
    const centre = t.x + t.w / 2;
    const room =
      t.align === "right" ? t.x + t.w : t.align === "center" ? 2 * Math.min(centre, width - centre) : width - t.x;
    // The same ellipsis walk as `clipToWidth` in `elements.ts`, inlined rather
    // than imported: that module imports `textWidth` and `finiteNodes` from
    // here, so reaching back for it would put this file in an import cycle with
    // it for four lines of loop.
    if (room <= 0) {
      dropped.add(n);
      continue;
    }
    // Through the shared walk — see `ellipsize`. This was its own copy of the
    // same loop, and it kept cutting emoji in half after the copy in
    // `clipToWidth` was fixed.
    const cut = ellipsize(t.text, t.fontSize, !!t.bold, room);
    if (cut) t.text = cut;
    else dropped.add(n);
  }
  return dropped.size ? nodes.filter((n) => !dropped.has(n)) : nodes;
}

/**
 * A polygon with nothing to draw — and the hole this gate had.
 *
 * `allNumbersFinite` asks whether every number in a node is finite, and an
 * EMPTY point list satisfies that trivially: there are no numbers to fail. So a
 * polygon carrying `points: []` sailed through a filter whose whole job is to
 * keep un-openable files from being written, and then broke exactly the
 * renderer the filter exists to protect. `pptx-paint.mjs` takes the polygon's
 * bounding box with `Math.min(...xs)`, which is `Infinity` for no points, and
 * writes `x="Infinity"` into the OOXML — not a number, not an Int64, and
 * Microsoft's own validator rejects the deck.
 *
 * Found by rendering 3033 hostile configs through the skill's headless
 * renderer: `{kind: "radar", data: {}}` lays out its grid rings before it knows
 * it has no axes, and emits two of them empty.
 *
 * Two points, not one, because a polygon is a closed path: one point has no
 * edges and cannot be a shape either. The rest of the chart is kept, which is
 * the same trade the filter above already makes.
 */
function degeneratePolygon(n: SceneNode): boolean {
  return n.kind === "polygon" && (n.points?.length ?? 0) < 2;
}

/** Every number anywhere in the node — including a polygon's point list. */
function allNumbersFinite(value: unknown, depth = 0): boolean {
  // Scene nodes are shallow; the bound is a cycle guard, not a shape claim.
  if (depth > 6) return true;
  if (typeof value === "number") return Number.isFinite(value);
  if (Array.isArray(value)) return value.every((v) => allNumbersFinite(v, depth + 1));
  if (value && typeof value === "object") {
    for (const v of Object.values(value)) if (!allNumbersFinite(v, depth + 1)) return false;
  }
  return true;
}

export interface Scene {
  width: number;
  height: number;
  nodes: SceneNode[];
  /**
   * Accessible name (the chart title) and a one-line text alternative
   * summarising the data. Emitted by the SVG renderer as `<title>`/`<desc>`
   * under `role="img"`, so a screen reader announces the chart instead of
   * silence. Set by buildChart; optional so hand-built scenes stay valid.
   */
  title?: string;
  desc?: string;
}

/**
 * How many NATIVE shapes a scene becomes on the Office.js host — NOT the node
 * count. A wedge fans out into `wedgeFanSteps` shapes (+2 stroke edges); a polygon
 * draws one line per edge. So a 10-node pie is ~50 shapes and a 10-node violin
 * ~250. This is the number the web shape budget cares about (counting nodes waved
 * both past it and the host choked) and the number the demo's contents table shows.
 */
export function estimateOfficeShapes(scene: Scene): number {
  let total = 0;
  for (const n of scene.nodes) {
    if (n.kind === "wedge") {
      const span = n.endAngle - n.startAngle;
      total += wedgeFanSteps(n.r, span).steps + (n.stroke && span < 359.9 ? 2 : 0);
    } else if (n.kind === "polygon") {
      total += n.points.length; // one line per edge, closed
    } else {
      total += 1;
    }
  }
  return total;
}

/**
 * Approximate rendered text width in points (average glyph ≈ 0.54 em for UI sans).
 *
 * Coerces, for the same reason `xmlText` and `paintText` do: the type says
 * `string` and the value came out of a file someone pasted. Sixty-odd call
 * sites across every layout ask this question about a title, a category, a
 * series name or a table cell, and a non-string used to answer two different
 * ways, both silent:
 *
 * - `null`/`undefined` THREW `Cannot read properties of null (reading
 *   'length')`, taking the whole chart down. `buildKpiTile({})` — a tile with
 *   no value yet — and `buildProcessFlow([null])` did exactly that, and both
 *   are exported from `src/index.ts` as the skill's public API.
 * - a NUMBER returned `NaN`, because `(2024).length` is `undefined`. That is
 *   the worse one. Every fit-to-width test here is a comparison, and each of
 *   them is FALSE against NaN — so shrink-to-fit silently stopped shrinking,
 *   and a width built as `Math.max(w, textWidth(...))` became NaN and had its
 *   whole node dropped by `finiteNodes`. `valueAxisTitle: 99` — a units label
 *   of `99`, or any year — vanished from the chart in all 25 kinds, with no
 *   error anywhere: the safety net turned a crash into a disappearance.
 *
 * `String()` rather than a rejection, matching `xmlText`: the renderers will
 * draw `2024`, so measuring it as four characters is the honest answer, and
 * measuring a missing string as zero is what every caller already means by it.
 */
/**
 * A code point that takes a FULL em rather than the Latin average.
 *
 * CJK ideographs, kana, Hangul and the full-width forms are one em wide in any
 * font that has them. Charged at 0.54 they came out 46% narrow — "売上高" measured
 * 19.4pt at 12pt where the glyphs take 36 — and this feeds every clip, every
 * shrink-to-fit and every de-collision test in the engine. So a Japanese,
 * Chinese or Korean chart had its labels judged to fit where they do not, went
 * unclipped where it should have been cut, and carried overlaps no fit could
 * see. Half-width katakana (U+FF61-FF9F) is deliberately excluded: it is narrow,
 * which is the entire point of that block.
 *
 * THE EMOJI LINE USED TO START AT U+1F300 AND THE BMP ONES FELL OFF THE BOTTOM.
 * "emoji and pictographs" is the rule, and about forty emoji do not live in the
 * astral planes at all — they sit in Miscellaneous Symbols, Dingbats,
 * Miscellaneous Technical and Geometric Shapes, and they have the same
 * East_Asian_Width=Wide and Emoji_Presentation=Yes as their astral twins. So
 * the engine measured one em for 🌟 and 0.54 for ⭐, 🟢 and ✅, 🔴 and ❌ — the same
 * glyph size, charged two different ways, and a status label like "✅ On track"
 * came out 5.5pt narrow at 12pt: the same 46%-per-glyph error the CJK note above
 * describes, in the same direction (judged to fit where it does not).
 *
 * The ranges below are exactly UAX #11's Wide entries in the BMP outside the CJK
 * blocks already listed. AMBIGUOUS-WIDTH SYMBOLS ARE NOT HERE and must not be
 * added: ✓ (U+2713), ✗ (U+2717), ▴/▾ (U+25B4/U+25BE) and → really do render at a
 * Latin advance, and this engine draws all four itself — `buildCheckbox`'s
 * glyphs and the funnel's conversion markers would start measuring double.
 */
function fullWidth(cp: number): boolean {
  return (
    (cp >= 0x231a && cp <= 0x231b) || // ⌚⌛ — BMP emoji, East_Asian_Width=Wide
    (cp >= 0x23e9 && cp <= 0x23ec) ||
    cp === 0x23f0 ||
    cp === 0x23f3 ||
    (cp >= 0x25fd && cp <= 0x25fe) ||
    (cp >= 0x2614 && cp <= 0x2615) ||
    (cp >= 0x2648 && cp <= 0x2653) ||
    cp === 0x267f ||
    cp === 0x2693 ||
    cp === 0x26a1 ||
    (cp >= 0x26aa && cp <= 0x26ab) ||
    (cp >= 0x26bd && cp <= 0x26be) ||
    (cp >= 0x26c4 && cp <= 0x26c5) ||
    cp === 0x26ce ||
    cp === 0x26d4 ||
    cp === 0x26ea ||
    (cp >= 0x26f2 && cp <= 0x26f3) ||
    cp === 0x26f5 ||
    cp === 0x26fa ||
    cp === 0x26fd ||
    cp === 0x2705 ||
    (cp >= 0x270a && cp <= 0x270b) ||
    cp === 0x2728 ||
    cp === 0x274c ||
    cp === 0x274e ||
    (cp >= 0x2753 && cp <= 0x2755) ||
    cp === 0x2757 ||
    (cp >= 0x2795 && cp <= 0x2797) ||
    cp === 0x27b0 ||
    cp === 0x27bf ||
    (cp >= 0x2b1b && cp <= 0x2b1c) ||
    cp === 0x2b50 ||
    cp === 0x2b55 ||
    (cp >= 0x1100 && cp <= 0x115f) || // Hangul Jamo
    (cp >= 0x2e80 && cp <= 0x303e) || // CJK radicals, Kangxi, CJK punctuation
    (cp >= 0x3041 && cp <= 0x33ff) || // kana, Hangul compat, CJK compat
    (cp >= 0x3400 && cp <= 0x4dbf) || // CJK ext A
    (cp >= 0x4e00 && cp <= 0x9fff) || // CJK unified
    (cp >= 0xa000 && cp <= 0xa4cf) || // Yi
    (cp >= 0xac00 && cp <= 0xd7a3) || // Hangul syllables
    (cp >= 0xf900 && cp <= 0xfaff) || // CJK compat ideographs
    (cp >= 0xfe30 && cp <= 0xfe6f) || // CJK compat forms, small forms
    (cp >= 0xff00 && cp <= 0xff60) || // FULL-width forms — NOT ff61+, which are half
    (cp >= 0xffe0 && cp <= 0xffe6) ||
    (cp >= 0x1f300 && cp <= 0x1faff) || // emoji and pictographs
    (cp >= 0x20000 && cp <= 0x3fffd) // CJK ext B and beyond
  );
}

/**
 * A code point that takes NO advance width at all — a combining mark.
 *
 * A mark is drawn ON its base letter, not beside it, so "e" and "é" are
 * the same number of points wide and render as the same glyph. Charged the
 * Latin average they came out DOUBLE, and which of the two a label is made of
 * is not something the author chose: `"é"` is one code point in NFC and two in
 * NFD, and **macOS hands out NFD** — a category name pasted from a Mac
 * datasheet arrives decomposed, the same text a Windows paste delivers
 * composed.
 *
 * So the engine measured the same visible string two different ways depending
 * on where it was copied from, and every fit downstream followed: at 12pt on a
 * 120x90 frame the NFC title was drawn at 9.5pt and the NFD one at 9.0, a
 * treemap drew 4 labels composed and 2 decomposed, and a butterfly lost one.
 * 252 charts out of a 25-kind × 5-frame × 3-font sweep laid out differently
 * with no visible difference in their text. Over-measuring is the quieter
 * direction — it drops labels rather than spilling ink off the chart — which is
 * why it survived the CJK pass sitting right next to it.
 *
 * The blocks here are the Mn/Me (non-spacing and enclosing) marks a chart
 * actually meets: the Latin/Greek/Cyrillic diacritics NFD produces, the Hebrew
 * points, the Arabic harakat, the Thai vowel signs and tone marks, and the
 * combining blocks for symbols and half-marks. It is not the whole Unicode mark
 * table, the same way `fullWidth` above is not the whole East-Asian-width
 * table — both are lists of the ranges that reach a business chart.
 *
 * VARIATION SELECTORS (U+FE00-FE0F) ARE DELIBERATELY NOT HERE, though they are
 * Mn. `❤️` is U+2764 plus U+FE0F, and the selector is the only reason that pair
 * measures 1.08em — which is about right for the emoji it actually renders as,
 * where the base alone would be charged 0.54. Zeroing the selector would make
 * an emoji-presentation sequence measure half its glyph, which is the same
 * silent halving `charWidth` refuses for surrogate pairs just below.
 */
function zeroWidth(cp: number): boolean {
  return (
    (cp >= 0x0300 && cp <= 0x036f) || // Combining Diacritical Marks — what NFD makes
    (cp >= 0x0483 && cp <= 0x0489) || // Cyrillic combining
    (cp >= 0x0591 && cp <= 0x05bd) || // Hebrew points
    cp === 0x05bf ||
    cp === 0x05c1 ||
    cp === 0x05c2 ||
    cp === 0x05c4 ||
    cp === 0x05c5 ||
    cp === 0x05c7 ||
    (cp >= 0x0610 && cp <= 0x061a) || // Arabic marks
    (cp >= 0x064b && cp <= 0x065f) ||
    cp === 0x0670 ||
    (cp >= 0x06d6 && cp <= 0x06dc) ||
    (cp >= 0x06df && cp <= 0x06e4) ||
    cp === 0x06e7 ||
    cp === 0x06e8 ||
    (cp >= 0x06ea && cp <= 0x06ed) ||
    cp === 0x0e31 || // Thai vowel signs and tone marks
    (cp >= 0x0e34 && cp <= 0x0e3a) ||
    (cp >= 0x0e47 && cp <= 0x0e4e) ||
    (cp >= 0x1ab0 && cp <= 0x1aff) || // Combining Diacritical Marks Extended
    (cp >= 0x1dc0 && cp <= 0x1dff) || // Combining Diacritical Marks Supplement
    (cp >= 0x20d0 && cp <= 0x20f0) || // Combining marks for symbols
    (cp >= 0xfe20 && cp <= 0xfe2f) // Combining half marks
  );
}

/**
 * One code point's width, the unit `textWidth` sums.
 *
 * An ASTRAL code point that is NOT in the full-width set — a flag's two regional
 * indicators (U+1F1E6-1F1FF), a playing card, a mahjong tile, an enclosed
 * ideograph, a mathematical alphanumeric — is charged TWO narrow units, which is
 * the 1.08em it was worth back when this counted UTF-16 code units. That number
 * is kept deliberately: `textWidth`'s own note says a surrogate pair measured
 * ~1.08em "about right for a pictograph by accident", and halving it is the
 * silent regression that note exists to refuse.
 *
 * The rule has to live in ONE place because two accountings that disagree by a
 * factor of two is exactly what `ellipsize` was tripping over: it seeds a
 * running width from `textWidth` and subtracts this per dropped code point, so
 * the moment the two charge a code point differently the clip returns text
 * wider than the width it was asked to fit.
 */
function charWidth(cp: number, fontSize: number, bold: boolean): number {
  if (zeroWidth(cp)) return 0;
  if (fullWidth(cp)) return fontSize;
  return fontSize * (bold ? 0.58 : 0.54) * (cp > 0xffff ? 2 : 1);
}

/**
 * `text` cut to fit `maxW`, with an ellipsis — the ONE implementation.
 *
 * There were two, identical and thirty lines apart: `clipToWidth` in elements.ts
 * and the de-collision walk here. Both did `while (…textWidth(cut) > room) cut =
 * cut.slice(0, -1)`, and fixing the surrogate bug in one of them left the other
 * cutting emoji in half — which is precisely the "the fix was written and one
 * call site was missed" shape this repo keeps finding, made once more while
 * fixing an instance of it. One function, two callers, no third answer.
 *
 * LINEAR, where both of those were quadratic. `textWidth` is O(n) now that it
 * has to look at the characters, and calling it once per dropped character made
 * the walk O(n²): a 200-character title cost 40,000 character inspections, and a
 * table of long cells pays it per cell. This keeps a running width and subtracts
 * each dropped code point's own contribution, so the whole walk is one pass.
 */
export function ellipsize(text: string, fontSize: number, bold: boolean, maxW: number): string {
  const s = String(text ?? "");
  let w = textWidth(s, fontSize, bold);
  if (w <= maxW) return s;
  const ell = textWidth("…", fontSize, bold);
  let end = s.length;
  while (end > 0 && w + ell > maxW) {
    const last = s.charCodeAt(end - 1);
    const prev = end > 1 ? s.charCodeAt(end - 2) : 0;
    // A pair comes off together — half a pair is a lone surrogate, which is
    // invalid XML and reaches a .pptx as the repair dialog.
    const pair = last >= 0xdc00 && last <= 0xdfff && prev >= 0xd800 && prev <= 0xdbff;
    w -= charWidth(pair ? s.codePointAt(end - 2)! : last, fontSize, bold);
    end -= pair ? 2 : 1;
  }
  return end > 0 ? `${s.slice(0, end)}…` : "";
}

export function textWidth(text: string, fontSize: number, bold = false): number {
  const s = String(text ?? "");
  const narrow = bold ? 0.58 : 0.54;
  /**
   * THE LATIN PATH IS BYTE-IDENTICAL, deliberately.
   *
   * Sixty-odd call sites depend on this number and every shipped chart's fits
   * were measured against it, so a string with no full-width code point takes
   * exactly the `length * fontSize * ratio` it always did and nothing in the
   * deck moves. Only text that was being measured wrongly changes.
   *
   * EMOJI KEEP THEIR OLD WIDTH BY A DIFFERENT ROUTE. A surrogate pair used to
   * be counted as two Latin characters — 1.08em, which is about right for a
   * pictograph by accident. Counting code points naively would have HALVED
   * them: a silent regression on the way to fixing something else. A pair whose
   * code point is full-width is charged a full em; one that is NOT — a flag's
   * regional indicators, a playing card, a mathematical alphanumeric — keeps the
   * two narrow units it always had. See `charWidth`, which is now the only place
   * that rule is written.
   *
   * THAT ASTRAL RULE USED TO DEPEND ON THE REST OF THE STRING, which is the
   * defect this loop shape carried. The sum was written two ways — a code-point
   * count when the string held a full-width glyph, `s.length` when it did not —
   * so "🇩🇰" measured 2.16em alone and 1.08em next to a kanji. `ellipsize` seeds
   * its running width from here and subtracts `charWidth`, so a clip that
   * removed the last full-width character flipped the accounting underneath
   * itself and returned a string up to TWICE the width it was asked to fit:
   * `clipToWidth("🇩🇰🇸🇪🇳🇴 北欧の売上", 12, 40)` came back 71.3pt wide. Downstream
   * that is a row label 29pt off the left edge of a 160pt heatmap — drawn onto
   * whatever the chart sits beside, since neither PowerPoint renderer clips.
   *
   * One accounting now, summed per code point, and it is ADDITIVE: a code
   * point's width no longer depends on its neighbours. Latin is still
   * bit-identical (no surrogate pairs, so the unit count IS `s.length`), pure
   * CJK is unchanged, and a lone 📊 is still one em.
   */
  let wide = 0;
  let narrowUnits = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    const next = i + 1 < s.length ? s.charCodeAt(i + 1) : 0;
    // A HIGH surrogate is only half a pair if a LOW one follows it. Without the
    // second half of this test a lone high surrogate swallowed the character
    // after it, and that character then cost nothing at all.
    const pair = c >= 0xd800 && c <= 0xdbff && next >= 0xdc00 && next <= 0xdfff;
    const cp = pair ? s.codePointAt(i)! : c;
    if (pair) i++;
    // A combining mark is drawn on its base and adds nothing — see `zeroWidth`.
    if (zeroWidth(cp)) continue;
    if (fullWidth(cp)) wide++;
    else narrowUnits += pair ? 2 : 1;
  }
  return (wide + narrowUnits * narrow) * fontSize;
}

/**
 * Pick black or white ink for a given fill so segment labels stay readable.
 * Parses via the shared `toRgb` — this used to carry its own hex-only copy, which
 * read every rgb()/hsl() fill as pure black and so chose WHITE ink for a
 * near-white segment.
 */
export function contrastInk(fill: string): string {
  const [r, g, b] = toRgb(fill).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum > 0.35 ? "#0b0b0b" : "#ffffff";
}

/**
 * Move a node horizontally by `dx`.
 *
 * Every horizontal coordinate a `SceneNode` can carry, which is four scalars
 * (`x`, `x1`, `x2`, `cx`) and one array — `points`, the polygon's own geometry.
 * The array is the whole reason this lives here rather than beside its caller.
 *
 * It began as a duck-typed loop over the four scalar names in the demo
 * gallery, a shape that can never fail to COMPILE when a node kind gains a
 * coordinate, and `points` was already missing from it: a polygon stayed where
 * it was while everything around it moved. That was latent — the gallery only
 * composes Harvey balls and checkboxes, and `src/core/elements.ts` emits no
 * polygon — but the helper's contract is "shift this node" and it quietly did
 * not, for one kind, with nothing to say so.
 *
 * Here it sits next to the node contract it has to keep up with, and it is
 * reachable by a test, which the gallery module is not: that file touches the
 * DOM at import time.
 */
export function shiftNodeX<T extends SceneNode>(n: T, dx: number): T {
  const node = n as unknown as Record<string, number>;
  for (const k of ["x", "x1", "x2", "cx"]) if (typeof node[k] === "number") node[k] += dx;
  const pts = (n as unknown as { points?: { x: number; y: number }[] }).points;
  if (Array.isArray(pts)) for (const p of pts) p.x += dx;
  return n;
}
