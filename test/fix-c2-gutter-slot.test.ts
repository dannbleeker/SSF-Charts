import { describe, expect, it } from "vitest";
import { nearestClearCentre, tightBox } from "../src/core/collide";
import { uprightNameSlot } from "../src/core/layout/column";
import type { TextNode } from "../src/core/scene";

/**
 * The two pure pieces behind r2-11's follow-up: where a label may stand in a
 * strip that already has text in it (`nearestClearCentre`), and where an upright
 * combo's line name may stand in the gutter it shares with the column names and
 * the secondary tick numbers (`uprightNameSlot`).
 *
 * The scene-level proof — that the 90 pairs of the overlap sweep are gone and the
 * names are still drawn — lives in `fix-c2-line-family.test.ts`. These pin the
 * rules those scenes depend on, one at a time, so a change to any of them fails
 * here with its own name rather than as a pair of nodes somewhere in a chart.
 */

describe("nearestClearCentre: the closest place in a strip that touches nothing", () => {
  const box = (y: number, h: number, x = 0, w = 100) => ({ x, y, w, h });

  it("stays where it was asked to stand when that is clear", () => {
    expect(nearestClearCentre(50, 10, 0, 40, 0, 200, [box(100, 10)], 99)).toBe(50);
  });

  it("goes to the nearer end of the obstacle it would cross", () => {
    // An obstacle 44..52 and a label 10 tall centred on 50: it clears by centring
    // at 57 (below, 7 away) or 39 (above, 11 away).
    const c = nearestClearCentre(50, 10, 0, 40, 0, 200, [box(44, 8)], 99) as number;
    expect(c).toBeGreaterThan(57);
    expect(c).toBeLessThan(57.5);
    const up = nearestClearCentre(46, 10, 0, 40, 0, 200, [box(44, 8)], 99) as number;
    expect(up).toBeLessThan(39);
    expect(up).toBeGreaterThan(38.5);
  });

  it("keeps a little air off what it clears, because touching exactly is overlapping to the settle pass", () => {
    const c = nearestClearCentre(50, 10, 0, 40, 0, 200, [box(44, 8)], 99) as number;
    // The label's top edge is c - 5; the obstacle's bottom edge is 52.
    expect(c - 5).toBeGreaterThan(52);
  });

  it("ignores an obstacle that is not in the label's column", () => {
    expect(nearestClearCentre(50, 10, 0, 40, 0, 200, [box(44, 8, 60, 40)], 99)).toBe(50);
    expect(nearestClearCentre(50, 10, 0, 40, 0, 200, [box(44, 8, -100, 100)], 99)).toBe(50);
  });

  it("threads a gap between two obstacles when it is wide enough, and not when it is not", () => {
    // Obstacles end at 40 and start at 62: 22 of room.
    expect(nearestClearCentre(51, 10, 0, 40, 0, 200, [box(30, 10), box(62, 10)], 99)).toBe(51);
    // 12 of room for a label 10 tall + air each side fits; 9 does not.
    const tight = nearestClearCentre(50, 10, 0, 40, 0, 200, [box(30, 10), box(49, 10)], 5);
    expect(tight).toBeNull();
  });

  it("will not travel past the reach it is given", () => {
    const wall = [box(0, 100)];
    expect(nearestClearCentre(50, 10, 0, 40, 0, 200, wall, 40)).toBeNull();
    expect(nearestClearCentre(50, 10, 0, 40, 0, 200, wall, 60)).not.toBeNull();
  });

  it("keeps the label's own extent inside the limits", () => {
    // A 10-tall label cannot be centred above 5 or below 195 in 0..200.
    expect(nearestClearCentre(2, 10, 0, 40, 0, 200, [], 99)).toBe(5);
    expect(nearestClearCentre(199, 10, 0, 40, 0, 200, [], 99)).toBe(195);
    // …and there is no answer at all when the band is shorter than the label.
    expect(nearestClearCentre(5, 10, 0, 40, 0, 8, [], 99)).toBeNull();
  });
});

describe("uprightNameSlot: where a combo line's name stands in the gutter", () => {
  const FS = 10;
  const text = (over: Partial<TextNode>): TextNode => ({
    kind: "text",
    x: 100,
    y: 0,
    w: 40,
    h: 14,
    text: "Series 1",
    fontSize: FS,
    color: "#000",
    align: "left",
    valign: "middle",
    ...over,
  });
  /** A tick number: drawn at 0.9 of the chart font, left-aligned at the gutter's edge. */
  const tick = (y: number) => text({ y, w: 20, text: "60", fontSize: FS * 0.9, name: "secondary-axis" });
  /** Column names stacked edge to edge from `from` to `to`, each as tall as a painted name. */
  const wall = (from: number, to: number) =>
    Array.from({ length: Math.round((to - from) / 12.5) }, (_, i) =>
      text({ y: from + i * 12.5 - 0.75, h: 14, name: `series-label-${i}` }),
    );
  const NAME_X0 = 100;
  const NAME_X1 = 140;
  const limits = { top: 0, bottom: 120 };
  const drawnAt = (y: number) => y + FS * 0.7; // the centre of a name whose box top is y

  it("stays on its natural spot when nothing is there", () => {
    expect(uprightNameSlot([tick(10)], [], 50, FS, FS, NAME_X0, NAME_X1, limits)).toBe(50);
  });

  it("moves off a tick number, to the room beside it, and the tick is not what gives way", () => {
    // A tick centred under where the name hangs, as the pareto's top tick is.
    const t = tick(50);
    const y = uprightNameSlot([t], [], 50, FS, FS, NAME_X0, NAME_X1, limits) as number;
    expect(y).not.toBe(50);
    const name = { y: drawnAt(y) - 6.25, h: 12.5 };
    const tb = tightBox(t);
    expect(name.y + name.h <= tb.y || name.y >= tb.y + tb.h, "the name clears the tick").toBe(true);
  });

  it("stands ON a tick number only when there is no room anywhere else near", () => {
    // Names edge to edge above and below, and the one gap between them is where the
    // tick is: the tick is what gives way (the caller drops it).
    const t = tick(52);
    const around = [...wall(0, 48), ...wall(66, 120)];
    const y = uprightNameSlot([t, ...around], [], 50, FS, FS, NAME_X0, NAME_X1, limits) as number;
    expect(y).not.toBeNull();
    const c = drawnAt(y);
    expect(c - 6.25).toBeGreaterThanOrEqual(48);
    expect(c + 6.25).toBeLessThanOrEqual(66);
    const tb = tightBox(t);
    expect(c - 6.25 < tb.y + tb.h && tb.y < c + 6.25, "it overlaps the tick").toBe(true);
  });

  it("is homeless when the strip is full", () => {
    expect(uprightNameSlot(wall(0, 120), [], 50, FS, FS, NAME_X0, NAME_X1, limits)).toBeNull();
  });

  it("keeps the order the lines end in: it will not climb past a name that belongs above it", () => {
    // The only room is above a name that has already been drawn for a line that
    // ends above this one, so the limit shuts it out.
    const aboveIt = { top: 40, bottom: 120 };
    expect(uprightNameSlot([...wall(40, 120)], [], 60, FS, FS, NAME_X0, NAME_X1, aboveIt)).toBeNull();
    // Without the limit the same strip would have taken it just above the wall.
    expect(
      uprightNameSlot([...wall(40, 120)], [], 60, FS, FS, NAME_X0, NAME_X1, { top: 0, bottom: 120 }),
    ).not.toBeNull();
  });

  it("does not stand a name further from its line than the settle pass would carry one", () => {
    // A wall from 20 to 120 leaves 0..20 free, room for a name. 5.5 em (55pt at
    // 10pt) is the reach: that room is 43pt from a natural spot 50 down the strip,
    // and out of reach from one at 100.
    expect(uprightNameSlot(wall(20, 120), [], 50, FS, FS, NAME_X0, NAME_X1, limits)).not.toBeNull();
    expect(uprightNameSlot(wall(20, 120), [], 100, FS, FS, NAME_X0, NAME_X1, limits)).toBeNull();
  });

  it("measures a name that was shrunk by what it paints, not by its box", () => {
    // The gap between the walls is 11.5 tall. 6pt text paints 7.5 and fits; 10pt
    // text paints 12.5 and does not, though both have the same 14pt box.
    const gap = [...wall(0, 40), ...wall(49, 120)];
    expect(uprightNameSlot(gap, [], 40, FS, 6, NAME_X0, NAME_X1, limits)).not.toBeNull();
    expect(uprightNameSlot(gap, [], 40, FS, FS, NAME_X0, NAME_X1, limits)).toBeNull();
  });
});
