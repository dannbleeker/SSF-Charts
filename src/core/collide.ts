import type { SceneNode, TextNode } from "./scene";
import { textWidth } from "./scene";
import { BoxHash, gridCellFor } from "./grid";

/**
 * Global label de-collision pass (a lightweight take on think-cell's
 * guaranteed non-overlapping labels): after layout + decorations, outside
 * labels (totals, CAGR/difference/value-line labels, series labels) are
 * nudged upward until they clear every other label's tight bounding box.
 * Inside-segment and axis labels stay fixed — they own their space.
 */

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Movable label name prefixes, in nudge-priority order (later = moves). */
const MOVABLE = [
  /^total-/,
  /^value-line-label/,
  // The column series' names and the combo LINE's name share one gutter, so
  // they share one rank. Separate tiers put the line's name in the later one,
  // where it settled against names already fixed and could only move UP —
  // so on a 160x120 combo at the default font it climbed from 58.5 to 31.0,
  // past "Services" at 48.6, and the gutter read Margin % / Services / Product
  // top to bottom against lines running the other way. A label that names
  // someone else's line is the failure the settle order below exists to
  // prevent, and it is prevented WITHIN a rank only: same rank, sorted
  // bottom-up, each nudge moves a label away from the one beneath it, so two of
  // them can never cross.
  /^(?:combo-)?series-label-/,
  /^diff-label$/,
  /^cagr-label$/,
  // The combo line's point labels, LAST so they are the ones that move: the
  // column's value label is anchored to the bar it sits on, where the line's
  // label is a floater above a point. They were absent, so on an ordinary
  // 480x300 combo at the default font a line label and a column label were drawn
  // through each other — this pass exists for exactly that and was simply never
  // told about them.
  /^combo-label-/,
];

/**
 * Labels that may be flipped BELOW their mark when there is no room above.
 *
 * The nudge only goes up, for a good reason — of two labels naming different
 * things, pushing one down crosses the one beneath it and they end up naming
 * each other's marks. That reason does not apply to a label anchored to a
 * single POINT: below its point is as true as above it, which is where
 * think-cell puts one when the room above is spent.
 *
 * Without the flip these labels had nowhere to go. On a 480x300 combo at a 26pt
 * font a line's point label sat on the column total beneath it, was nudged 143
 * points up over ten tries, and came to rest inside the TITLE — every nudge
 * legal, the budget exhausted, and the collision merely relocated onto a worse
 * partner. Restricted to this list rather than opened to every movable label,
 * because a series label really can be reordered by moving down.
 *
 * The CAGR caption is here for the same reason and is worth its own note,
 * because the OTHER way of getting it off the title was tried and reverted:
 * flooring it at the title's bottom turned five `title x cagr-label` overlaps
 * into eight against the column totals. That was an unconditional move, and
 * this is not — the flip only takes a position already clear of everything
 * settled, and the totals rank ahead of it and are settled first. So it cannot
 * buy its way off the title by landing on a total; it stays put instead. Worth
 * 13 of the sweep's overlapping pairs against the earlier attempt's −3.
 */
const FLIPPABLE = [/^combo-label-/, /^cagr-label$/];

/**
 * The furthest a label anchored to a single mark may be nudged, in multiples of
 * its own font size. Past this it is restored to where the layout put it, and
 * `unplaceableComboLabels` decides whether it survives at all.
 */
const MAX_ANCHORED_TRAVEL = 2;

/**
 * Flippable labels whose bound BELOW is measured from their mark: a combo's point
 * labels, which hang just above a marker they name.
 *
 * Not the CAGR caption, deliberately. Its mark is an ARROW, sloped, and the
 * caption legitimately flips past the whole of it — 3.3 em of travel in the
 * report that raised this, ending about one em under the arrow — while all this
 * pass can see of the arrow is the caption's own box. A bound read off the
 * caption would either refuse that flip or be no bound at all, and the flip is
 * worth its 13 pairs. The report's sweep over 6000 random combo and stacked
 * charts found it far from its arrow in 3 of ~230 flips, 45pt at the worst.
 */
const POINT_LABEL = /^combo-label-/;

const canFlip = (name: string | undefined) => !!name && FLIPPABLE.some((re) => re.test(name));

const movableRank = (name: string | undefined): number => {
  if (!name) return -1;
  return MOVABLE.findIndex((re) => re.test(name));
};

/** Actual painted extent of a text node, given its alignment. */
export function tightBox(n: TextNode): Box {
  const w = Math.min(n.w, textWidth(n.text, n.fontSize, n.bold));
  const h = Math.min(n.h, n.fontSize * 1.25);
  const x = n.align === "left" ? n.x : n.align === "right" ? n.x + n.w - w : n.x + (n.w - w) / 2;
  const y = n.valign === "top" ? n.y : n.valign === "bottom" ? n.y + n.h - h : n.y + (n.h - h) / 2;
  return { x, y, w, h };
}

const overlaps = (a: Box, b: Box) => a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;

export function resolveLabelCollisions(nodes: SceneNode[], canvasH?: number): void {
  const texts = nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text);
  const fixed: Box[] = [];
  const movable: { node: TextNode; rank: number }[] = [];
  for (const t of texts) {
    const rank = movableRank(t.name);
    if (rank >= 0) movable.push({ node: t, rank });
    else fixed.push(tightBox(t));
  }
  // Lower rank settles first and becomes an obstacle for the rest; within a
  // rank, the LOWEST label settles first.
  //
  // The order within a rank used to be the order the layout emitted, and the
  // nudge only ever goes up — so of two overlapping labels the one that
  // happened to come first stayed put and the other was pushed past it. On a
  // chart dense enough that its series labels cannot all fit at their minimum
  // gap, that reordered them: every label still on the canvas, each naming the
  // wrong line. Settling bottom-up makes the nudge move a label AWAY from the
  // one below it, which can never cross.
  movable.sort((a, b) => a.rank - b.rank || b.node.y - a.node.y);

  // Spatial hash over the settled boxes so each movable label tests only its
  // neighbourhood, not the whole settled set on every one of its ≤10 nudges. The
  // exact `overlaps` test still decides, so the nudged positions are identical.
  const cell = gridCellFor([...fixed, ...movable.map((m) => tightBox(m.node))]);
  const settled = new BoxHash<Box>(cell);
  for (const b of fixed) settled.insert(b, b);
  for (const { node } of movable) {
    const startY = node.y;
    let box = tightBox(node);
    let tries = 0;
    let clear = !settled.some(box, (s) => overlaps(box, s));
    // How far a label anchored to a SINGLE MARK may travel before it stops being
    // that mark's label. The budget below is ten steps of 0.55 em, so without a
    // bound a combo point label could climb 5.5 times its own font size: measured
    // over 25 kinds x 8 frames x 7 fonts, 68 of them moved and the worst went
    // 123pt — five em — ending as a number floating in the title band with
    // nothing under it, while every other family stayed within 1.7em. A label
    // that far from its point is not labelling it, and the reader has no way to
    // tell which point it came from.
    //
    // Restored rather than left where the climb ran out, which is what the flip
    // below already does and for the same reason: a label that moved and did not
    // clear collides somewhere its author did not choose.
    //
    // Only the anchored ones (`FLIPPABLE` — a combo's point labels and the CAGR
    // caption). A series label names a LINE and reads correctly anywhere along
    // it, and a total names the column beneath it however high it sits.
    const cap = canFlip(node.name) ? node.fontSize * MAX_ANCHORED_TRAVEL : Infinity;
    while (!clear && tries < 10) {
      if (startY - node.y + node.fontSize * 0.55 > cap) {
        node.y = startY;
        box = tightBox(node);
        break;
      }
      node.y -= node.fontSize * 0.55; // nudge upward
      box = tightBox(node);
      tries++;
      if (box.y < 0) {
        // Nudging would push the label off the top of the canvas (y=0 for every
        // chart). An overlapping label still reads; an off-canvas one is lost — so
        // give up and restore. Without this, a column total sharing the totals row
        // with the fixed grand-total label was nudged clean off the top.
        node.y = startY;
        box = tightBox(node);
        break;
      }
      clear = !settled.some(box, (s) => overlaps(box, s));
    }
    // Up was blocked or spent, so try BELOW the mark — from the label's own
    // starting point, never from wherever the failed climb left it. Only for a
    // label anchored to a single point (see FLIPPABLE), and only while the
    // canvas height is known: without it there is no bottom edge to refuse at,
    // and this would push labels off the foot of the chart the way the guard
    // above stops it pushing them off the head.
    if (!clear && canvasH != null && canFlip(node.name)) {
      node.y = startY;
      /**
       * HOW FAR BELOW ITS MARK A POINT LABEL MAY LAND, which the climb above has a
       * cap for and this loop had none.
       *
       * Ten steps of 0.55 em is 5.5 em, and the first text-clear spot was taken
       * wherever it fell. A combo's "66" went from y 103 to 158 — 3.8 em under
       * its own marker, inside the column beneath it and under that column's "68"
       * — because the CAGR caption had flipped into the space beneath the marker
       * first and the only clear spot left was past it. That is the failure the
       * cap exists to prevent ("a label that far from its point is not labelling
       * it"), reached by the other direction. About 1% of random default-font
       * combos end with a point label two em or more under its marker.
       *
       * Measured from the MARK, not from where the label started: the label
       * starts above its marker, so getting just under it already costs 1.7-2.2
       * em, and the climb's own bound (travel from the start) would refuse the
       * ordinary flip. The mark is read off the label's own box — `valign` says
       * which edge faces it (a point label hangs above its marker, so its bottom).
       * Refused, the label stays home like every other label that found no place,
       * and `unplaceableComboLabels` decides whether it survives.
       */
      const home = tightBox(node);
      const markY = node.valign === "bottom" ? home.y + home.h : node.valign === "top" ? home.y : home.y + home.h / 2;
      const reach = POINT_LABEL.test(node.name ?? "") ? markY + node.fontSize * MAX_ANCHORED_TRAVEL : Infinity;
      for (let down = 0; down < 10; down++) {
        node.y += node.fontSize * 0.55;
        const cand = tightBox(node);
        if (cand.y + cand.h > canvasH) break;
        if (cand.y > reach) break;
        if (!settled.some(cand, (s) => overlaps(cand, s))) {
          box = cand;
          clear = true;
          break;
        }
      }
      // Neither direction cleared: leave it where the layout put it. A label
      // moved and still colliding is strictly worse than one that never moved —
      // it collides somewhere its author did not choose.
      if (!clear) node.y = startY;
      box = tightBox(node);
    }
    settled.insert(box, box);
  }
}

/**
 * A combo chart's two families of value labels — the column totals and the
 * line's point labels — can both want the same band on a short frame, and on a
 * short frame neither can move.
 *
 * `resolveLabelCollisions` above resolves every one of these if the canvas is
 * unbounded: the point label is FLIPPABLE, so it goes below its point, landing
 * at y 57-75 on a 60pt-tall chart. That destination is off the bottom, so the
 * flip is correctly refused and the label stays where the layout put it —
 * drawn through the total. Shrinking does not help either, because the two are
 * centred at the same y whenever the column total and the line value coincide,
 * which is a property of the DATA.
 *
 * So one of them has to go, and `decor.tightLabelPriority` says which. This runs
 * AFTER de-collision on purpose: only labels that are still colliding once every
 * legal move has been tried are dropped, so a roomy chart keeps both and so does
 * a short chart whose data happens to separate them.
 *
 * Returns the nodes to drop rather than mutating, so the caller owns the scene
 * array and this stays a pure decision — the same split as `positionalSweepPlan`
 * and `chooseGroupMembers`.
 */
export function unplaceableComboLabels(nodes: SceneNode[], priority: "columns" | "line" = "columns"): Set<SceneNode> {
  // `!!n.text`, never `n.text.trim()`: `text` is `string` in the types and a
  // number in a config someone pasted, so calling a string method on it throws.
  // `resolveLabelCollisions` above tests truthiness for the same reason and
  // `textWidth` coerces with `String(text ?? "")`. The hostile-input sweep
  // caught this within one run of adding it.
  const texts = nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text);
  const totals = texts.filter((t) => /^total-/.test(t.name ?? ""));
  const points = texts.filter((t) => /^combo-label-/.test(t.name ?? ""));
  const drop = new Set<SceneNode>();
  // The loser is dropped, the winner is left exactly as de-collision settled it.
  const [losers, keepers] = priority === "line" ? [totals, points] : [points, totals];
  // Only "is there anything to drop". This used to also require the KEEPER
  // family to exist, which is why a HORIZONTAL combo — which draws no column
  // totals at all — got no drop pass whatever it collided with, and its point
  // labels sat on the line's own name with nothing able to separate them.
  if (!losers.length) return drop;
  // The keeper family, plus every label that will not move for anyone: the
  // title, the category names, the segment labels. A point label capped to two
  // em of travel can no longer climb out from under those, and one left sitting
  // on a category name is as unreadable as one sitting on a total. Movable
  // labels are deliberately not in this list — they may still settle elsewhere,
  // and dropping a label because of something that has not stopped moving is
  // how a pass like this starts eating its own output.
  const others = [
    ...keepers,
    ...texts.filter((t) => movableRank(t.name) < 0),
    // The line's own NAME, which settles a rank ahead of these labels and so is
    // fixed by the time they are placed — movable in principle, immovable in
    // practice for them. Sideways it sits at the line's last point, in the same
    // few points of canvas the last point labels want, and one label naming the
    // series is worth more than one more number on a chart that already carries
    // its column totals.
    ...texts.filter((t) => /^combo-series-label-/.test(t.name ?? "")),
  ];
  for (const l of losers) {
    const lb = tightBox(l);
    if (others.some((k) => overlaps(lb, tightBox(k)))) drop.add(l);
  }
  return drop;
}

/**
 * Axis tick numbers a point label has been drawn over. The tick numbers go.
 *
 * SCATTER AND BUBBLE ONLY, by the names involved. The placer that positions
 * point labels is given a band a line and a half TALLER than the plot, and that
 * extra strip is where the x tick numbers live. That was decided and measured:
 * confining the band to the plot takes the overlapping-text count for these two
 * kinds from 889 to 599 and drops 56 of 301 point labels on a chart as
 * comfortable as 480x300. The verdict was that a point's label is DATA and a
 * tick number is chrome, so the chrome yields.
 *
 * IT DID NOT ACTUALLY YIELD. Both were drawn, on top of each other — the one
 * overlap the frame gate allows by name, and 140 pairs across the sweep. A
 * number printed through another number is not chrome giving way; it is two
 * unreadable numbers where the decision called for one readable one. This is the
 * pass that carries out the verdict already reached.
 *
 * The tick numbers are dropped, never moved. Moving one puts it beside the wrong
 * gridline, which on an axis is the same class of lie a pie label beside the
 * wrong wedge would be.
 *
 * NOT bounded by "keep at least N ticks". A bound like that would have to keep a
 * number the reader cannot see, which is the state this pass exists to end — and
 * the gridlines stay whatever happens, so the axis keeps its structure. Where
 * every tick on an axis is covered, the chart was already telling the reader
 * nothing on that axis, and now it says so honestly.
 */
export function tickLabelsUnderPointLabels(nodes: SceneNode[]): Set<SceneNode> {
  const drop = new Set<SceneNode>();
  // `!!n.text` rather than `.trim()`, for the reason above: `text` is typed as a
  // string and arrives as a number from a pasted config.
  const texts = nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text);
  const points = texts.filter((t) => /^label-\d+$/.test(t.name ?? ""));
  if (!points.length) return drop;
  const boxes = points.map(tightBox);
  for (const t of texts) {
    if (t.name !== "x-axis" && t.name !== "y-axis") continue;
    const tb = tightBox(t);
    if (boxes.some((b) => overlaps(tb, b))) drop.add(t);
  }
  return drop;
}
