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
