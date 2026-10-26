/**
 * What actually changed between two renderings of the same chart.
 *
 * The add-in redraws a chart by deleting every shape and adding every shape
 * back. On PowerPoint for the web that is the single most expensive thing it
 * does: the 2026-08-11 rounds put a 24-shape chart at ~50 seconds — three
 * batches at ~17s each — and the cost grows with what is already on the slide,
 * so it gets worse as a deck fills up.
 *
 * Almost none of that work is needed. Measured against the clustered sample:
 *
 *     retitle                 1 of 24 nodes changed
 *     one data point edited   2 of 24
 *     rescale (scale.max)    18 of 24
 *
 * A retitle costs 24 deletes and 24 adds to move one text box's string.
 *
 * This module decides, from the two scenes alone, whether the difference can be
 * applied to the shapes already on the slide — and says no whenever anything is
 * less than obvious. It is pure so that the decision can be checked without a
 * PowerPoint, which is the same reason `positionalSweepPlan` and
 * `chooseGroupMembers` are their own functions: on this host the decision is
 * where the safety lives, and the host calls are just hands.
 */

import type { Scene, SceneNode } from "./scene";
import { arrowheadBox, dashKind } from "./geometry";

/**
 * Node kinds whose every drawn property can be written back onto an existing
 * shape, exactly as `addNode` would have set it.
 *
 * `rect` and `text` only, and the list is short on purpose rather than by
 * accident. Both are `addGeometricShape`/`addTextBox` calls with a closed set
 * of property writes — geometry, fill, line, name; plus the string, font and
 * alignment for text — so an in-place applier can mirror them line for line and
 * be checked against the adder.
 *
 * Everything else is geometry BAKED AT CREATION and unreachable afterwards.
 * A wedge is a fan of rotated triangles, an arrowhead is a triangle rotated
 * about a computed box, a polygon degrades to an outline: changing the numbers
 * that produced them means producing them again. Office.js has no freeform
 * paths to edit, so there is nothing to write to.
 *
 * These two cover the cases that matter — every `seg-*`, `label-*`, `title`,
 * `category-*`, `series-label-*` and `baseline` in a bar, column, line or area
 * chart is one or the other.
 */
const UPDATABLE = new Set<SceneNode["kind"]>(["rect", "text"]);

/** Whether every changed property of this kind can be written to a live shape. */
export function isUpdatableKind(kind: SceneNode["kind"]): boolean {
  return UPDATABLE.has(kind);
}

/**
 * A stable fingerprint of a scene's nodes.
 *
 * This is what closes the version-skew hole, and without it the whole idea is
 * unsound. An update rebuilds the OLD scene from the config stored on the chart
 * and diffs it against the new one — but "what the stored config renders to
 * today" and "what was actually drawn on that slide" are the same thing only
 * while the engine has not changed. Ship a new default colour, a nudged label
 * offset, an extra gridline, and every chart already in the deck is one the
 * rebuilt scene does not describe. The diff would then report those nodes as
 * unchanged, skip them, and leave the old rendering on the slide FOREVER —
 * where today's redraw-everything quietly repairs them on the next edit.
 *
 * So the fingerprint is written when the chart is drawn and checked when it is
 * updated. Matching means the rebuilt scene is the one on the slide; anything
 * else — a different engine, a hand-edited tag, a chart from another deck —
 * fails the check and takes the full redraw, which is the behaviour that has
 * always been correct.
 *
 * FNV-1a over the same JSON the diff compares, so the hash cannot disagree with
 * the comparison it guards. Not a security boundary — a chart's tags are
 * editable in the host, and the worst a forged fingerprint buys is the redraw
 * path this function exists to avoid.
 */
export function sceneFingerprint(scene: Scene): string {
  const text = JSON.stringify([scene.width, scene.height, scene.nodes]);
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    // 16777619, as shifts — `Math.imul` keeps it in 32 bits on every host.
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(36);
}

/** Which shapes an in-place update would have to write to. */
/**
 * The groups of shape properties an in-place write can set independently.
 *
 * These mirror `applyNodeInPlace`'s own structure so the two can be read
 * side by side, which is the same reason `UPDATABLE` is a closed set: an
 * applier that writes something this diff cannot describe would silently write
 * stale values.
 *
 *     box    left, top, width, height   (and, for text, the frame constants
 *                                        that only matter when geometry moves)
 *     fill   the interior               (rect only)
 *     line   the outline                (rect only)
 *     text   the string itself
 *     font   size, colour, bold, family (text only)
 *     align  horizontal and vertical    (text only)
 */
export type NodeChange = "box" | "fill" | "line" | "text" | "font" | "align";

export interface SceneUpdatePlan {
  /** Indices into `next.nodes`, ascending. Empty means nothing changed at all. */
  changed: number[];
  /**
   * For each entry of `changed`, in the same order, which groups actually
   * differ.
   *
   * WHY THIS EXISTS. `applyNodeInPlace` wrote every property of a changed node
   * unconditionally — the host's own statement list put ONE text node at
   * roughly twenty statements (left, top, width, height, fill, fill.clear,
   * lineFormat, textFrame, textRange, text, wordWrap, four font properties,
   * paragraphFormat, name, tags). A retitle changes one string and was sending
   * twenty; a rescale moves eighteen bars and was sending some three hundred
   * and sixty statements to change seventy-two numbers.
   *
   * The diff already holds both scenes, so it already knows which of those
   * twenty are stale. Saying so costs nothing here and is the only place the
   * answer exists.
   */
  parts: NodeChange[][];
}

/**
 * The plan for turning `prev` into `next` by writing to shapes that already
 * exist, or `null` when that cannot be done safely.
 *
 * Null means "redraw the chart" and is the answer to every question this
 * function is not certain about. The bar is deliberately high, because the
 * fallback costs seconds and a wrong plan costs a chart:
 *
 * - **The frame must be identical.** A chart that changed size is re-laid-out;
 *   every node moves, and the shapes' own box no longer matches the scene's.
 * - **Same node count, same order, same kind and same name at every index.**
 *   This is not a similarity metric. The shape a plan writes to is found
 *   POSITIONALLY — the chart's anchor is node 0 and its parts tag lists the
 *   rest in drawing order — so an index that means a different node in the two
 *   scenes writes a bar's geometry onto a label. Structural change of any sort
 *   is a redraw.
 * - **Every CHANGED node must be an updatable kind.** An unchanged wedge is
 *   fine and stays where it is; a changed one has no writable geometry, so its
 *   chart is redrawn whole.
 *
 * An empty `changed` is a real answer and not the same as null: the config
 * round-tripped without altering the picture, and the right response is to
 * write the new config tag and touch no shapes at all.
 */
export function planSceneUpdate(prev: Scene, next: Scene): SceneUpdatePlan | null {
  if (prev.width !== next.width || prev.height !== next.height) return null;
  if (prev.nodes.length !== next.nodes.length) return null;
  const changed: number[] = [];
  const parts: NodeChange[][] = [];
  for (let i = 0; i < next.nodes.length; i++) {
    const a = prev.nodes[i];
    const b = next.nodes[i];
    if (a.kind !== b.kind) return null;
    // A node with no name cannot be found on the slide by any route, so a
    // change to one is not applicable — but an unchanged one is harmless, and
    // several kinds legitimately go unnamed. Compared rather than required.
    if (a.name !== b.name) return null;
    if (JSON.stringify(a) === JSON.stringify(b)) continue;
    if (!isUpdatableKind(b.kind)) return null;
    // A changed node has to be nameable to be written to.
    if (!b.name) return null;
    const groups = changedGroups(a, b);
    // A node that differs by JSON but names no group is a node whose change is
    // in a field the applier cannot write. Refusing the whole plan is the same
    // answer this function gives to every other "less than obvious" case, and
    // it is what keeps `parts` honest: an empty list would be read as "write
    // nothing" and would silently drop the edit.
    if (!groups.length) return null;
    changed.push(i);
    parts.push(groups);
  }
  return { changed, parts };
}

/**
 * Which property groups differ between two versions of one node.
 *
 * Every field either belongs to a group or is compared by `planSceneUpdate`
 * before this is reached (`kind`, `name`). The `default` on the exhaustiveness
 * check below is what makes a NEW node field fail loudly: add `n.opacity` to a
 * rect and this returns no group for a node the differ says changed, which
 * `planSceneUpdate` turns into a redraw rather than a silent stale write.
 */
function changedGroups(a: SceneNode, b: SceneNode): NodeChange[] {
  const out: NodeChange[] = [];
  const box = (x: SceneNode, y: SceneNode) =>
    (x as { x: number }).x !== (y as { x: number }).x ||
    (x as { y: number }).y !== (y as { y: number }).y ||
    (x as { w: number }).w !== (y as { w: number }).w ||
    (x as { h: number }).h !== (y as { h: number }).h;
  if (box(a, b)) out.push("box");
  if (b.kind === "rect" && a.kind === "rect") {
    if (a.fill !== b.fill) out.push("fill");
    if (a.stroke !== b.stroke || (a.strokeWidth ?? 0) !== (b.strokeWidth ?? 0)) out.push("line");
    return out;
  }
  if (b.kind === "text" && a.kind === "text") {
    if (a.text !== b.text) out.push("text");
    if (a.fontSize !== b.fontSize || a.color !== b.color || !!a.bold !== !!b.bold || a.fontFamily !== b.fontFamily)
      out.push("font");
    if (a.align !== b.align || a.valign !== b.valign) out.push("align");
  }
  return out;
}

/**
 * What a host read back from one shape. A field is absent when the host did not
 * answer for it, which is an ordinary outcome here and never evidence of
 * anything — see `checkShapesAgainstScene`.
 */
export interface LiveShape {
  name?: string;
  width?: number;
  height?: number;
}

/** The first place the shapes on the slide and the scene they should be drawing part company. */
export interface ShapeMismatch {
  what: "name" | "size";
  /** Index into the scene's nodes, which is also the shape's place in the one-for-one mapping. */
  index: number;
  node: string | undefined;
  expected: string | [number, number];
  found: string | [number, number];
}

export interface ShapeCheck {
  /** Shapes whose NAME the host read back and that matched their node's. */
  names: number;
  /** Rect shapes whose SIZE the host read back and that matched. */
  sizes: number;
  /** The first disagreement, or null. */
  off: ShapeMismatch | null;
}

/**
 * How far a rect's live side may sit from the scene's before it counts as
 * changed. The host stores EMU and rounds, so a bar drawn 35.7pt wide reads back
 * as 35.7 to a couple of decimals; a point is far above that and far below any
 * resize a person would make on purpose.
 */
const SIZE_TOLERANCE = 1;

/**
 * The floor `addNode` and `applyNodeInPlace` both put under a rect's sides, so
 * a node that lays out to nothing is compared with what was actually drawn.
 */
const MIN_RECT_SIDE = 0.2;

/**
 * Do the shapes on the slide look like the scene the plan was diffed against?
 *
 * `planSceneUpdate` compares two scenes with each other, and
 * `sceneFingerprint` compares the stored config with the stored fingerprint.
 * NEITHER LOOKS AT A SHAPE. "The rebuilt scene is what is on the slide" was a
 * premise, and the slide is something a person can change: select one member of
 * a group and Bring it to Front, rename a bar in the Selection Pane, drag the
 * group's handle. The count of members survives all three, so the one-for-one
 * guard passed and the writer put a bar's geometry on a label.
 *
 * Two facts the host can be asked, both already known to this module:
 *
 * - **name.** `applyNodeInPlace` and the live adders write `shape.name =
 *   node.name`, and the fingerprint covers the names, so a matching fingerprint
 *   means the shape at index `i` was DRAWN under `nodes[i].name`. If it now
 *   carries ANOTHER NODE'S name, two shapes have traded places and the
 *   positional mapping is no longer the mapping. That is the only name that is
 *   evidence: a name this scene has no node for says nothing about order. The
 *   deck builder (`skill/scripts/pptx-paint.mjs`) never names its shapes, so a
 *   chart inserted that way carries `Shape 3` and `Text 12` until its first
 *   redraw, and its first edit is allowed to be fast; a bar a person renamed in
 *   the Selection Pane is still where it was. Kinds that draw several shapes
 *   under suffixed names (`-f0`, `-edge`) are not compared at all.
 * - **size** (`sizes`), for rects only. A user resize scales every member, so
 *   writing the changed bars at the scene's size leaves them beside neighbours
 *   at the stretched one — and the edited bar reads the wrong value against
 *   its neighbours as well as sitting off the axis. A text box is left out
 *   because its host-side height depends on an autosize setting that is
 *   written in a try and may not have landed. The caller asks for sizes only
 *   when it is about to write geometry: a retitle writes none, is right on a
 *   stretched chart, and should not cost the person their resize.
 *
 * NARROW ON PURPOSE, the way `positionalGuessNamesOurs` is: a mismatch needs
 * POSITIVE evidence — another node's name, or a size the host read back and that
 * is wrong. A field the host did not answer for, or answered with a name this
 * scene does not know, is skipped, so a host that reads none of them behaves
 * exactly as before. `names` and `sizes` count what WAS compared, so the caller
 * can tell "checked 24 and all agree" from "checked nothing", which return the
 * same `off: null`.
 */
export function checkShapesAgainstScene(prev: Scene, live: readonly LiveShape[], opts: { sizes: boolean }): ShapeCheck {
  const out: ShapeCheck = { names: 0, sizes: 0, off: null };
  const ours = new Set(prev.nodes.map((n) => n.name).filter(Boolean));
  const count = Math.min(prev.nodes.length, live.length);
  for (let i = 0; i < count; i++) {
    const node = prev.nodes[i];
    const shape = live[i];
    if (node.name && node.kind !== "wedge" && node.kind !== "polygon" && shape.name && ours.has(shape.name)) {
      if (shape.name !== node.name) {
        out.off = { what: "name", index: i, node: node.name, expected: node.name, found: shape.name };
        return out;
      }
      out.names++;
    }
    if (opts.sizes && node.kind === "rect" && Number.isFinite(shape.width) && Number.isFinite(shape.height)) {
      const want: [number, number] = [Math.max(MIN_RECT_SIDE, node.w), Math.max(MIN_RECT_SIDE, node.h)];
      const got: [number, number] = [shape.width as number, shape.height as number];
      if (Math.abs(got[0] - want[0]) > SIZE_TOLERANCE || Math.abs(got[1] - want[1]) > SIZE_TOLERANCE) {
        out.off = { what: "size", index: i, node: node.name, expected: want, found: got };
        return out;
      }
      out.sizes++;
    }
  }
  return out;
}

/**
 * The factors a person has stretched a chart by since it was drawn: how many
 * times as wide and as tall its rects are now, against the scene they were drawn
 * from. `1` on an axis whose rects all still agree with the scene to within
 * `SIZE_TOLERANCE`, so a chart stretched in one direction reads `{ x: 1, y: 2 }`.
 */
export interface LiveScale {
  x: number;
  y: number;
  /** Rects whose size the host read back and that the factors were fitted to. */
  samples: number;
  /**
   * Whose size the factors were read off. Absent means the members', which is
   * what `liveScaleOfScene` returns; `"group"` is `groupScaleOfScene`, read off
   * the group's own box for a host whose members said nothing about size.
   */
  via?: "group";
}

/**
 * Fewer rects than this say nothing about the chart as a whole: one bar a person
 * dragged by hand is a bar they changed, not a chart they resized.
 */
const MIN_SCALE_SAMPLES = 2;

/**
 * The factors beyond which a reading is a host that answered nonsense, not a
 * person: a chart a tenth of its size or ten times it. A reading outside them
 * is no evidence, so the chart keeps the size its config gives it as before.
 */
const MIN_LIVE_SCALE = 0.1;
const MAX_LIVE_SCALE = 10;

/**
 * How much a person has stretched a chart, read off the shapes the host reported
 * on — or `null` when they have not, or when the shapes cannot say.
 *
 * `checkShapesAgainstScene` answers "is a rect off by more than a point", and
 * stops at the first that is. That is enough to refuse writing a changed bar
 * beside stretched neighbours; it is not enough to draw the chart again at the
 * size the person made it, which needs the FACTOR. Resizing a group in
 * PowerPoint scales every member about one corner, so the factor is one number
 * per axis, and the rects the host already reported on carry it: no host call
 * of its own.
 *
 * The same discipline as the check beside it, POSITIVE EVIDENCE ONLY, and in
 * three ways that matter when the answer is used to change what gets drawn:
 *
 * - **The pairing has to be the real one.** Rects are paired with nodes by
 *   position, so a shape that now carries another node's name (a child brought
 *   to the front) is no mapping at all and the answer is `null`.
 * - **One factor per axis, and every rect has to agree with it** to within the
 *   same point. A person who dragged a single bar, or two bars by different
 *   amounts, has not resized the chart; fitting a factor to that would redraw
 *   the whole chart at a size nobody chose. Such a chart gets `null` and keeps
 *   the old behaviour.
 * - **Silence is not evidence.** A rect whose size the host did not read is
 *   skipped, fewer than `MIN_SCALE_SAMPLES` read means `null`, and a chart whose
 *   rects all agree with the scene means `null` too, so a chart nobody touched
 *   never gets a factor and is redrawn exactly as it always was.
 *
 * The factor is the ratio of the SUMS of the sides, which weights each rect by
 * its size: a 0.2pt bar (the floor under a bar with no value) that the host
 * rounds is outvoted by the 300pt bar beside it. The per-rect agreement test is
 * what stops that from hiding an outlier.
 */
export function liveScaleOfScene(prev: Scene, live: readonly LiveShape[]): LiveScale | null {
  if (checkShapesAgainstScene(prev, live, { sizes: false }).off) return null;
  const want: [number, number][] = [];
  const got: [number, number][] = [];
  for (let i = 0; i < Math.min(prev.nodes.length, live.length); i++) {
    const node = prev.nodes[i];
    const { width, height } = live[i];
    if (node.kind !== "rect" || !Number.isFinite(width) || !Number.isFinite(height)) continue;
    want.push([Math.max(MIN_RECT_SIDE, node.w), Math.max(MIN_RECT_SIDE, node.h)]);
    got.push([width as number, height as number]);
  }
  if (want.length < MIN_SCALE_SAMPLES) return null;
  const axis = (k: 0 | 1): number | null => {
    if (want.every((w, i) => Math.abs(got[i][k] - w[k]) <= SIZE_TOLERANCE)) return 1;
    const sum = (rows: [number, number][]) => rows.reduce((total, r) => total + r[k], 0);
    const factor = sum(got) / sum(want);
    if (!(factor >= MIN_LIVE_SCALE && factor <= MAX_LIVE_SCALE)) return null;
    return want.every((w, i) => Math.abs(got[i][k] - w[k] * factor) <= SIZE_TOLERANCE) ? factor : null;
  };
  const x = axis(0);
  const y = axis(1);
  if (x === null || y === null || (x === 1 && y === 1)) return null;
  return { x, y, samples: want.length };
}

/**
 * How many members the host reported BOTH sides of, whatever each is drawn as.
 * Zero is a host whose members said nothing about size, which is the case
 * `groupScaleOfScene` exists for, and it is a different fact from "every rect
 * agreed with the scene" — `null` from `liveScaleOfScene` and `off: null` from
 * `checkShapesAgainstScene` are the same for both.
 *
 * EVERY MEMBER COUNTS, A TEXT BOX OR AN ELLIPSE AS MUCH AS A RECT. This used to
 * count rects only (`rectSizesRead`), which asked "did a RECT say how big it is"
 * when the question is "does this host say how big its members are". A scene with
 * no rect in it (showcase scatter #41 and #117, bubble #116, clustered #56: four
 * of 123 charts) read zero on a host that answered every one of its members, so
 * the group's box was consulted and a hand resize kept by it on the live host —
 * with no live evidence that those charts' boxes match what the scene derives
 * (the four matches on record are rect-bearing kinds). Whether a host answers a
 * size is a property of the host and not of the node kind; `liveScaleOfScene` and
 * `checkShapesAgainstScene` read the answers they get, and having none to read is
 * what lets the box stand in.
 */
export function memberSizesRead(live: readonly LiveShape[]): number {
  return live.filter((shape) => Number.isFinite(shape.width) && Number.isFinite(shape.height)).length;
}

/** A shape's size as the host gave it; absent when it did not. */
export interface LiveBox {
  width?: number;
  height?: number;
}

/** How big a drawn chart's shapes are, taken together. */
export interface DrawnExtent {
  width: number;
  height: number;
}

/**
 * The floors `addNode`, `addText` and `addSegment` put under a side, so a node
 * that lays out to nothing is framed as what was drawn. `MIN_RECT_SIDE` is the
 * rect's, ellipse's, chevron's and symbol's.
 */
const MIN_TEXT_SIDE = 4;
const MIN_LINE_SIDE = 0.5;

/**
 * The frame one node becomes on the slide, as [left, top, right, bottom] in scene
 * points, or `null` for a node drawn as something the host ROTATES.
 *
 * A diagonal line is a rectangle of the segment's length turned about its midpoint,
 * a wedge is a fan of turned triangles and an arrowhead a turned triangle: each
 * has a box whose `left`/`top`/`width`/`height` is the one before the turn, and
 * the group the host builds around them is framed on the box after it — or not;
 * `rotation-keeps-the-unrotated-box` has never been answered. A frame that depends
 * on which is a frame this module will not guess, so those nodes make the whole
 * scene's extent unknown. What stays: rects, text, ellipses, chevrons, symbols,
 * and a line that is horizontal, vertical or dashed (`addSegment` draws those as
 * `addLine`, unrotated, in the line's own bounding box).
 *
 * KEPT IN STEP WITH THE ADDERS BY A MEASUREMENT, NOT BY A COMMENT. The constants
 * above are copied from them; `test/scene-drawn-extent.test.ts` draws every sample
 * kind and every chart in the showcase on the fake host and requires this to
 * equal the box the group it builds reports.
 */
function drawnFrame(n: SceneNode): [number, number, number, number] | null {
  const at = nodeCorner(n);
  if (!at) return null;
  const box = (w: number, h: number, floor = MIN_RECT_SIDE): [number, number, number, number] => [
    at[0],
    at[1],
    at[0] + Math.max(floor, w),
    at[1] + Math.max(floor, h),
  ];
  switch (n.kind) {
    case "rect":
    case "chevron":
      return box(n.w, n.h);
    case "ellipse":
      return box(n.rx * 2, n.ry * 2);
    case "symbol":
      return box(n.size * 2, n.size * 2);
    case "text":
      return box(n.w, n.h, MIN_TEXT_SIDE);
    case "line": {
      const w = Math.abs(n.x2 - n.x1);
      const h = Math.abs(n.y2 - n.y1);
      if (w >= MIN_LINE_SIDE && h >= MIN_LINE_SIDE && dashKind(n.dash) === "none") return null;
      return box(w, h, MIN_LINE_SIDE);
    }
    default:
      return null;
  }
}

/**
 * Where a node's frame starts, [left, top] in scene points, for EVERY kind the
 * deck builder draws, or `null` for one it draws nothing for.
 *
 * The one place a node's corner is decided: `drawnFrame` takes the left and top of
 * its box from here, and `sceneInkOffset` is the least of these over a scene. A
 * node the host TURNS has a corner here although it has no frame in `drawnFrame`,
 * because the deck builder frames it unturned and `boundsOf` in `ooxml.ts` reads
 * only that frame: a wedge on its circle's box (a solid pie wedge, a ring and a
 * full turn alike), a polygon on its points, an arrowhead on `arrowheadBox`, and a
 * diagonal line on its bounding box. What the LIVE host frames a turned node on is
 * the `rotation-keeps-the-unrotated-box` question, and nothing here answers it.
 */
function nodeCorner(n: SceneNode): [number, number] | null {
  switch (n.kind) {
    case "rect":
    case "chevron":
    case "text":
      return [n.x, n.y];
    case "ellipse":
      return [n.cx - n.rx, n.cy - n.ry];
    case "symbol":
      return [n.cx - n.size, n.cy - n.size];
    case "line":
      return [Math.min(n.x1, n.x2), Math.min(n.y1, n.y2)];
    case "polygon":
      // Fewer than two points is a polygon the deck builder does not write.
      if (!Array.isArray(n.points) || n.points.length < 2) return null;
      return [Math.min(...n.points.map((p) => p.x)), Math.min(...n.points.map((p) => p.y))];
    case "wedge":
      return [n.cx - n.r, n.cy - n.r];
    case "arrowhead": {
      const box = arrowheadBox(n.x, n.y, n.size, n.angle);
      return [box.left, box.top];
    }
    default:
      return null;
  }
}

/**
 * Where the group the DECK BUILDER frames on this scene begins, measured from the
 * frame origin the scene is drawn at: the first ink, `null` for a scene it draws
 * nothing for.
 *
 * THE INK OFFSET. A titled chart has none as a rule (the title sits at 0,0), not
 * always: a titled mekko starts 2pt BEFORE the frame at every sample size, a titled
 * combo 10.65pt before it at 200x150, and a titled horizontal stacked chart at
 * 160x100 with its category axis off and a value axis on 22pt before it, where the
 * first tick label overhangs. An untitled clustered chart starts 42pt down and a
 * pie 112pt to the right, so it can be positive or negative on either axis.
 * `POWERCHART_ORIGIN` records the frame origin and the group's own corner; a reader
 * that has only the first (the tag the old deck generator wrote stored it twice)
 * adds this to it for the second. That only holds for the engine that drew the
 * group: see `withTrueAnchor`, which asks the scene's fingerprint first.
 *
 * AN ANSWER FOR EVERY KIND, where `sceneDrawnExtent` abstains on any scene with a
 * node the host turns. The two ask different questions. That one is the size of
 * the group the LIVE host makes, and for a turned node the host's frame is not the
 * scene's to guess. This is the corner of the group a FILE holds, which is the
 * deck builder's and nobody else's, and the deck builder frames every node
 * unturned (see `nodeCorner`). Only a node's own corner reaches it, never a floor
 * on its far side, so the floors that decide the extent of a hairline cannot move it.
 *
 * KEPT IN STEP BY A MEASUREMENT: `test/fix-v4-legacy-origin-anchor.test.ts` builds
 * every sample kind, titled and not, and the showcase, through `buildDeckBase64`,
 * whose tag records the corner of the group it really made, and requires the two
 * to agree.
 */
export function sceneInkOffset(scene: Scene): { x: number; y: number } | null {
  let x = Infinity;
  let y = Infinity;
  for (const node of scene.nodes) {
    const at = nodeCorner(node);
    if (!at) continue;
    x = Math.min(x, at[0]);
    y = Math.min(y, at[1]);
  }
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
}

/**
 * The size of the box the host's group takes around this scene's shapes, or
 * `null` when it cannot be known from the scene alone.
 *
 * THE SAME COMPUTATION THE GROUP FOLLOWS, which is the union of its members'
 * frames: `boundsOf` in `src/render/ooxml.ts` frames the deck builder's group on
 * it and measures the origin anchor from it (0e9500f), and the fake host's
 * `addGroup` does the same over the shapes it was handed. This is that union over
 * the nodes' frames instead of over XML, because the update path has a scene and
 * must not load a zip library to ask a size.
 *
 * AND IT AGREES WITH THE LIVE HOST WHERE IT HAS BEEN COMPARED, which is what
 * makes the factor below worth trusting. The archive's deck inventories carry
 * the real group boxes of the charts the battery drew, but not the config behind
 * each, so this is a match of SIZES and not a join on identity. The configs the
 * battery draws derive these, and each is among the `PowerChart` boxes the
 * inventories record, to the hundredth of a point:
 *
 *     clustered 480x300   this 480 x 299     host 480 x 299     885 groups
 *     clustered 780x420   this 780 x 419     host 780 x 419      10
 *     clustered 168x80    this 168 x 79      host 168 x 79      361
 *     waffle    320x220   this 320 x 213.34  host 320 x 213.34    4
 *
 * Four coincidences like that are unlikely, and they say that for those two kinds
 * the host neither auto-sizes a text box nor widens the group by a stroke. Every
 * other kind rests on the fake host alone, which is why `groupScaleOfScene` stays
 * quiet inside a margin.
 */
export function sceneDrawnExtent(scene: Scene): DrawnExtent | null {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const node of scene.nodes) {
    const frame = drawnFrame(node);
    if (!frame) return null;
    left = Math.min(left, frame[0]);
    top = Math.min(top, frame[1]);
    right = Math.max(right, frame[2]);
    bottom = Math.max(bottom, frame[3]);
  }
  return right > left && bottom > top ? { width: right - left, height: bottom - top } : null;
}

/**
 * The group's size against the size the scene would draw it at.
 *
 * `ratio` is live over drawn, per axis, and is present whenever both are known —
 * including when the chart was never touched and it reads 1. That is what a trace
 * needs to say how far the instrument is from a perfect reading on a healthy
 * chart, and `groupScaleOfScene` is what decides whether a ratio is a person.
 */
export function measureGroup(
  prev: Scene,
  group: LiveBox,
): { drawn: DrawnExtent | null; ratio: { x: number; y: number } | null } {
  const drawn = sceneDrawnExtent(prev);
  const ok = Number.isFinite(group.width) && Number.isFinite(group.height);
  if (!drawn || !ok) return { drawn, ratio: null };
  return { drawn, ratio: { x: (group.width as number) / drawn.width, y: (group.height as number) / drawn.height } };
}

/**
 * How far the live group may sit from the drawn one before it is a person and not
 * the host: the larger of 5pt and 5% of that side. The host's box agreed to the
 * digit on every live comparison above, so this is not sized to a noise that has
 * been seen; it is sized to what a wrong answer COSTS. A reading outside it makes
 * the redraw lay the whole chart out at a new size, and that size is written into
 * the chart's config — a bias in this instrument would move every chart a little
 * further on every edit. A hand resize under 5% is left to the half-stretched
 * write it would get without this, which is under 5% off.
 */
const GROUP_NOISE_RATIO = 0.05;
const GROUP_NOISE_POINTS = 5;

/** A side shorter than this says nothing: 5pt of 20 is a quarter. */
const MIN_GROUP_SIDE = 40;

/** The factors a group box can be believed to have been dragged to: 4x either way. */
const MIN_GROUP_SCALE = 0.25;
const MAX_GROUP_SCALE = 4;

/**
 * How much a person has stretched a chart, read off the GROUP's own box — for the
 * host that will not say how big the group's members are.
 *
 * Everything `liveScaleOfScene` reads is a member, and a host can answer a
 * member's id and nothing else (`faults.groupMembersSilentOn`). The group is the
 * one shape the update path already holds and has already loaded the size of, and
 * resizing a group scales its box by the same factor it scales every member, so a
 * box that no longer matches what the scene draws is a resize.
 *
 * WEAKER EVIDENCE, AND TREATED AS WEAKER:
 *
 * - **It cannot tell a resize from a dragged member.** Moving one bar inside the
 *   group moves the box too, and the per-rect agreement test that rejects that
 *   case is exactly what a silent host denies us. So the caller consults this
 *   ONLY when no member of any kind gave a size (`memberSizesRead` is zero); once
 *   a member has answered, `liveScaleOfScene` and `checkShapesAgainstScene` are
 *   the evidence and the group's box is only traced.
 * - **A scene with a rotated node has no known box** (`sceneDrawnExtent`), and
 *   answers `null`: a line chart or a pie is never read this way.
 * - **Inside the noise band an axis reads exactly 1**, so a chart stretched in one
 *   direction reads `{ x: 1.5, y: 1 }`, and a chart inside it on both reads
 *   `null` — as does a factor beyond 4x either way.
 *
 * `samples` is 0 because no member size was read; `via` says whose box it was.
 */
export function groupScaleOfScene(prev: Scene, group: LiveBox): LiveScale | null {
  const { drawn, ratio } = measureGroup(prev, group);
  if (!drawn || !ratio) return null;
  const axis = (live: number, want: number): number | null => {
    if (want < MIN_GROUP_SIDE || Math.abs(live - want) <= Math.max(GROUP_NOISE_POINTS, GROUP_NOISE_RATIO * want))
      return 1;
    const factor = live / want;
    return factor >= MIN_GROUP_SCALE && factor <= MAX_GROUP_SCALE ? factor : null;
  };
  const x = axis(group.width as number, drawn.width);
  const y = axis(group.height as number, drawn.height);
  if (x === null || y === null || (x === 1 && y === 1)) return null;
  return { x, y, samples: 0, via: "group" };
}

/**
 * Whether a plan is worth taking, given what the redraw would have cost.
 *
 * Kept separate from `planSceneUpdate` because it is a judgement about COST,
 * where the plan is a judgement about SAFETY.
 *
 * THE ORIGINAL ARITHMETIC WAS WRONG, AND 0.5 IS STILL THE RIGHT NUMBER. Both
 * halves of that are measured, and they are measured separately.
 *
 * The old reasoning read: "a redraw is one delete plus one add per node, so an
 * update touching more than half is already doing more host calls per shape
 * saved." That counts an in-place write as ONE call per node. It is about
 * twenty, as the host's own statement list showed in round 145.
 *
 * AND THE FAST PATH IS MUCH FASTER, at a small share. Rounds 143-145 redrew;
 * 146-149 took the fast path, same scenarios, same deck, same host:
 *
 *     edit a chart on the visible slide   21218 22471 25352  ->   5176  5445  6108  5409
 *     edit the chart the user selected    27459 30275 29341  ->   9494  9207  9130  9294
 *     an update follows a moved chart     27364 31637 28914  ->   8886 10491  9120  9109
 *
 * Three to four times faster, three scenarios, no overlap. Every one of those
 * samples is `changed 1 of 24` — a 4% share.
 *
 * SO 0.8 WAS TRIED, AND IT CRASHED POWERPOINT — SIX TIMES OUT OF SIX. Round 150
 * (`6359d83`) raised the limit to admit the 18-of-24 and 9-of-16 edits that
 * `same scale across the deck` makes. Every one of its seven attempts died:
 *
 *     crashed at  255s  284s  282s  278s  257s  282s
 *
 * A 29-second window, six for six, on the first round after the change — and
 * `same scale across the deck` starts at 280s (round 149) and 302s (round 147).
 * The crash sits exactly where the one scenario the limit governs begins. No
 * round before it had crashed more than once, and none had failed to archive.
 *
 * The likely mechanism is the one this project already documents elsewhere: the
 * web host's per-slide budget, and proxy-memory exhaustion on a large batch.
 * Eight charts writing eighteen shapes apiece in place is a far bigger batch
 * than eight redraws, because a redraw spreads its work across deletes and adds
 * that the host paces itself.
 *
 * **So the cost model is refuted and a limit still stands.** The fast path wins
 * at 4% and kills the host at 75%; the crossover is somewhere between.
 *
 * 0.6 IS THAT NEXT STEP, and it is deliberately a small one. It admits the
 * 9-of-16 edits and still declines the 18-of-24 ones, so where 0.8 handed the
 * host eight charts writing eighteen shapes apiece — 144 extra shape-writes in
 * a round — this hands it two charts writing nine, which is 18. If the crash
 * was batch size, that is an eighth of the dose that caused it.
 *
 * Round 150 at 0.5 is the control and the host is healthy: 13/13 scenarios,
 * `same scale across the deck` at 195928 ms against a 196-237s band.
 *
 * ROUND 151 MEASURED IT AND 0.6 IS THE CEILING. It completed, 13/13, with
 * declines 8 to 6 and successes 3 to 5, and that scenario ran 164518 ms — 31
 * seconds below the MINIMUM of its five prior observations, on the same eight
 * charts and the same thirteen attempts.
 *
 * There is no rung above it. This battery produces exactly two shares, 9-of-16
 * and 18-of-24, so 0.7 would admit nothing new; and any limit that admits
 * 18-of-24 also admits 9-of-16, which is precisely the dose 0.8 died on six
 * times. 0.75 is not untested — it is tested, and it kills the host.
 *
 *     0.5    0 extra shape-writes    safe
 *     0.6    18 extra                safe, and 31s faster
 *     0.75+  162 extra               fatal, six times over
 *
 * THE SHARE IS A PROXY FOR THE WRONG THING. What the host refuses is shapes per
 * BATCH, and this update used to write every changed shape of every chart into
 * one batch and sync once. The demo deck hit the same wall and was fixed by
 * chunking — one `PowerPoint.run` per slide (#112) — not by a threshold.
 *
 * SO THE BATCH IS BOUNDED NOW, and this is 0.8 again on purpose.
 * `IN_PLACE_WRITES_PER_SYNC` caps a write at six shapes per sync, and round 152
 * measured that cap as free: 164626 ms against 164518 ms unchunked, 0.07% on a
 * 164-second scenario, with successes and declines unmoved.
 *
 * 0.8 is the value that crashed PowerPoint six times at 255-284s, and it is
 * deliberately the value used again, because it is the only one that
 * discriminates between the two explanations that round left standing:
 *
 * ROUND 153 ANSWERED IT: THE CRASH WAS PER-SYNC SIZE. It completed, 13/13, at
 * the same 0.8 that died six times unbounded.
 *
 *     150  0.5  unchunked   ok  3 | declined 10 | same scale 195928 ms
 *     151  0.6  unchunked   ok  5 | declined  8 | same scale 164518 ms
 *     152  0.6  CHUNKED     ok  5 | declined  8 | same scale 164626 ms
 *     153  0.8  CHUNKED     ok 11 | declined  2 | same scale 157403 ms
 *
 * `too much of the chart changed` went from six a round to NONE. The two
 * declines left are the correct ones: a picture, which is not in the scene, and
 * a chart with no parts list and no readable group members. Eleven of thirteen
 * is every attempt this battery can produce.
 *
 * SO THIS CONSTANT IS NO LONGER WHAT HOLDS THE HOST UP — `IN_PLACE_WRITES_PER_SYNC`
 * is. What survives here is a risk cap, not a cost model: where nearly every
 * node changes there is little left to save, and a redraw is the better-tested
 * way to a clean chart. Move it only on a round, and never without the batch
 * bound moving with it.
 */
export const UPDATE_SHARE_LIMIT = 0.8;

// `Pick`, not the whole plan: this is a judgement about how MANY nodes changed
// and it has never read anything else. Taking the full type made every caller
// that only has a count — the tests, and any future cost model — construct a
// `parts` it does not look at.
export function worthUpdating(plan: Pick<SceneUpdatePlan, "changed">, total: number): boolean {
  if (!total) return false;
  return plan.changed.length / total <= UPDATE_SHARE_LIMIT;
}
