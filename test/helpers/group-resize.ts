import { expect } from "vitest";
import { buildChart } from "../../src/core/chart";
import { sampleConfig } from "../../src/core/samples";
import { insertSceneIntoSlide } from "../../src/render/powerpoint";
import type { ChartConfig } from "../../src/core/types";
import { installHost, makeSlide, type FakeShape, type FakeSlide } from "./office-host";

/**
 * A person resizing a chart on the slide, and the charts to do it to.
 *
 * Shared by the suites about what an edit does AFTER a hand resize, so that each
 * of them stages the resize the same way: the way PowerPoint's own handles do,
 * every member scaling about the group's corner by one factor per axis.
 */

/** The frame every chart here is drawn at. */
export const FRAME = { left: 60, top: 90 };

/** The clustered sample, as a fresh object each time so a test can edit its copy. */
export const clustered = (): ChartConfig => JSON.parse(JSON.stringify(sampleConfig("clustered"))) as ChartConfig;

/** A copy of `cfg` with `change` applied. */
export const edited = (cfg: ChartConfig, change: (c: ChartConfig) => void): ChartConfig => {
  const next = JSON.parse(JSON.stringify(cfg)) as ChartConfig;
  change(next);
  return next;
};

/** One bar's value moved, which writes geometry and nothing else. */
export const aValueEdit = (cfg: ChartConfig): ChartConfig => edited(cfg, (c) => (c.data.series[0].values[1] = 30));

/** Draw `cfg` the way the pane does, and hand back the group the host made of it. */
export async function drawGrouped(
  cfg: ChartConfig,
): Promise<{ slide: FakeSlide; group: FakeShape; members: FakeShape[] }> {
  const slide = makeSlide("s1");
  installHost([slide], [], slide, () => true);
  await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), ...FRAME });
  const group = slide.created.find((s) => s.type === "group" && !s.deleted)!;
  expect(group, "the fixture did not group").toBeTruthy();
  return { slide, group, members: group.grouped as FakeShape[] };
}

/** A real group's box is the bounding box of its members, whenever one moves. */
export function groupBoxFollowsMembers(group: FakeShape): void {
  const kids = () => (group.grouped as FakeShape[]).filter((k) => !k.deleted);
  const edge = (pick: (k: FakeShape) => number, fold: (...n: number[]) => number) => fold(...kids().map(pick));
  const box = {
    left: () => edge((k) => k.left, Math.min),
    top: () => edge((k) => k.top, Math.min),
    width: () => edge((k) => k.left + k.width, Math.max) - box.left(),
    height: () => edge((k) => k.top + k.height, Math.max) - box.top(),
  };
  for (const [key, read] of Object.entries(box))
    Object.defineProperty(group, key, { get: read, set: () => {}, configurable: true });
}

/** A user resizing the group with PowerPoint's handles: every member scales about the group's corner. */
export function userResizes(group: FakeShape, sx: number, sy: number): void {
  const gl = group.left;
  const gt = group.top;
  for (const k of group.grouped as FakeShape[]) {
    k.left = gl + (k.left - gl) * sx;
    k.top = gt + (k.top - gt) * sy;
    k.width = k.width * sx;
    k.height = k.height * sy;
  }
  groupBoxFollowsMembers(group);
}

/** Every live shape on the slide that is not a group. */
export const liveShapes = (slide: FakeSlide): FakeShape[] =>
  slide.created.filter((s) => !s.deleted && s.type !== "group");

/** Every live group on the slide. */
export const liveGroups = (slide: FakeSlide): FakeShape[] =>
  slide.created.filter((s) => !s.deleted && s.type === "group");

/** The one live shape named `name`. */
export const byName = (slide: FakeSlide, name: string): FakeShape => {
  const found = liveShapes(slide).filter((s) => s.name === name);
  expect(found, `${name} is on the slide ${found.length} times`).toHaveLength(1);
  return found[0];
};
