/**
 * Beeswarm-constrained overlap relief: separate circles along ONE axis, leaving
 * the cross axis exact.
 *
 * Free 2D repulsion is deliberately not offered. In a scatter/bubble BOTH axes
 * encode data and a reader takes a value off each, so moving a marker in 2D
 * corrupts two readings at once with nothing to bound them. Along one named
 * axis there is exactly one approximation, it is capped, and the caller can
 * print the cap — the reader is told the error bar instead of being misled.
 *
 * Deterministic by construction: a fixed pass count, Jacobi updates (every pair
 * is measured against the same state, so the result cannot depend on visit
 * order), array iteration only, and no randomness.
 */
export interface SpreadItem {
  /** Position on the spread axis, px. */
  m: number;
  /** Position on the fixed cross axis, px. */
  c: number;
  /** Radius, px. */
  r: number;
}

const PASSES = 12;
const RELAX = 0.5;

/**
 * Signed displacement along the spread axis, index-aligned with `items`.
 *
 * `limit` is a hard cap: no item ever moves further than it, and keeping items
 * inside [min, max] can only ever reduce motion, never force it past the cap.
 * That ordering is the whole contract — the caller discloses `limit`.
 */
export function spreadAlongAxis(
  items: SpreadItem[],
  opts: { limit: number; min: number; max: number; pad?: number },
): number[] {
  const n = items.length;
  const disp = new Array<number>(n).fill(0);
  // `!(limit > 0)`, not `limit <= 0`: a NaN cap fails BOTH comparisons, so the
  // old guard let it through and `Math.max(-NaN, …)` made EVERY displacement NaN
  // — for a marker that overlaps nothing as much as for one that does, which is
  // the contract above broken at its root. No usable cap, no motion.
  if (n < 2 || !(opts.limit > 0)) return disp;
  const pad = opts.pad ?? 1;
  // Biggest first, ties by index: the big bubble is the visually dominant one
  // and should barely move. Mirrors the label pass, which already lets the
  // biggest bubbles win.
  const order = items.map((_, i) => i).sort((a, b) => items[b].r - items[a].r || a - b);

  for (let pass = 0; pass < PASSES; pass++) {
    const delta = new Array<number>(n).fill(0);
    const contacts = new Array<number>(n).fill(0);
    for (let ai = 0; ai < order.length; ai++) {
      for (let bi = ai + 1; bi < order.length; bi++) {
        const a = order[ai];
        const b = order[bi];
        const A = items[a];
        const B = items[b];
        const need = A.r + B.r + pad;
        const dc = Math.abs(A.c - B.c);
        if (dc >= need) continue; // the cross axis already separates them
        // How far apart they must be ALONG the spread axis, given the gap the
        // cross axis already provides.
        const minSep = Math.sqrt(need * need - dc * dc);
        const cur = items[b].m + disp[b] - (items[a].m + disp[a]);
        const gap = Math.abs(cur);
        if (gap >= minSep) continue;
        // cur === 0 (exactly co-located): break the tie by `order`, which is
        // itself deterministic.
        const sign = cur > 0 ? 1 : cur < 0 ? -1 : 1;
        const push = ((minSep - gap) / 2) * RELAX;
        // Weight by area, so the large bubble yields far less than the small one.
        const wa = (B.r * B.r) / (A.r * A.r + B.r * B.r || 1);
        delta[a] -= sign * push * 2 * wa;
        delta[b] += sign * push * 2 * (1 - wa);
        contacts[a]++;
        contacts[b]++;
      }
    }
    for (let i = 0; i < n; i++) {
      // AVERAGE the pair corrections rather than summing them. Summing applies
      // every pair's full correction to each member, so a circle in a cluster of
      // k takes k-1 of them in one pass — several times its own share, and the
      // dense cluster is exactly what this exists for. (Measured: the two settle
      // within a few px of each other over the passes below, and summing is not
      // unstable — averaging is simply the conservative one, and it is the
      // motion here that has to be justified to the reader.) Whichever is used,
      // the cap below is what bounds the result.
      const step = contacts[i] > 0 ? delta[i] / contacts[i] : 0;
      // The cap wins. Bounding the CENTRE to the plot can only pull the allowed
      // range toward 0 (the centre is inside by construction), so it composes
      // as an intersection with [-limit, +limit] instead of overriding it.
      // Bounding the circle's EXTENT would not: a bubble at an axis extreme
      // would be shoved out by up to its radius — ~4x the cap, and even when it
      // overlaps nothing. Markers already overhang the plot edge today.
      const lo = Math.max(-opts.limit, opts.min - items[i].m);
      const hi = Math.min(opts.limit, opts.max - items[i].m);
      disp[i] = Math.max(lo, Math.min(hi, disp[i] + step));
    }
  }
  return disp;
}
