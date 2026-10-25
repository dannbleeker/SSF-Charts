// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { listChartsInDeck, updateChartInSlide } from "../src/render/powerpoint";
import { setTracing, traceLog } from "../src/core/trace";
import { buildChart } from "../src/core/chart";
import { faults, installHost, makeSlide } from "./helpers/office-host";
import {
  aValueEdit,
  byName,
  clustered,
  drawGrouped,
  groupBoxFollowsMembers,
  liveShapes,
  userResizes,
} from "./helpers/group-resize";

/**
 * An edit after the person resized a chart, on a host that withholds sizes.
 *
 * `checkShapesAgainstScene` and `liveScaleOfScene` are verified against shapes
 * that ANSWER. Both are deliberately "positive evidence only": a width or a name
 * the host did not give is skipped, so a host that gives none behaves exactly as
 * it did before either existed — and a resized chart then gets the half-stretched
 * write they were written to stop. Nothing could show that until the fake could
 * fail the way a withholding host does (`faults.groupMembersSilentOn`).
 *
 * What the live host actually does is NOT what this file establishes. It states
 * what the product does when the host is the withholding one, and holds the fake
 * to doing what its switches say.
 */

afterEach(() => {
  vi.unstubAllGlobals();
  setTracing(false);
});

const sayings = () => traceLog().entries.filter((e) => e.message === "updated only the shapes that changed");

describe("the fake host that withholds what a group's members would say", () => {
  it("answers nothing by default, so no other suite changes", async () => {
    const { slide } = await drawGrouped(clustered());
    const group = slide.created.find((s) => s.type === "group")!;
    const shown = (group.group.shapes.items as { name?: string; width?: number; height?: number }[])[0];
    expect(typeof shown.name).toBe("string");
    expect(shown.width).toBeGreaterThan(0);
    expect(shown.height).toBeGreaterThan(0);
    expect(faults.groupMembersSilentOn).toEqual([]);
    expect(faults.groupSilentOn).toEqual([]);
  });

  it("withholds a member's name, width and height as an unpopulated property, and keeps its id", async () => {
    const { group, members } = await drawGrouped(clustered());
    faults.groupMembersSilentOn = ["name", "width", "height"];
    const first = (group.group.shapes.items as typeof members)[0];
    expect(first.id, "the id is what the host does answer").toBe(members[0].id);
    for (const prop of ["name", "width", "height"] as const)
      expect(() => first[prop], `${prop} was answered`).toThrow(/code=PropertyNotLoaded/);
    // Only what is named: the position is read by the origin tag's arithmetic.
    expect(typeof first.left).toBe("number");
    expect(typeof first.top).toBe("number");
  });

  it("can withhold as undefined instead, which is the other way a host fails to answer", async () => {
    const { group } = await drawGrouped(clustered());
    faults.groupMembersSilentOn = ["width"];
    faults.silentReadsAs = "undefined";
    const first = (group.group.shapes.items as { width?: number; name?: string }[])[0];
    expect(first.width).toBeUndefined();
    expect(typeof first.name).toBe("string");
  });

  it("reads the same member off the slide's own collection in full", async () => {
    const { slide, members } = await drawGrouped(clustered());
    faults.groupMembersSilentOn = ["name", "width", "height"];
    const listed = slide.shapes.items.find((s) => s.id === members[0].id)!;
    expect(listed.name).toBe(members[0].name);
    expect(listed.width).toBe(members[0].width);
  });

  it("writes through a withheld member to the one real shape", async () => {
    const { group, members } = await drawGrouped(clustered());
    faults.groupMembersSilentOn = ["name", "width", "height"];
    const first = (group.group.shapes.items as typeof members)[0];
    first.width = 123;
    first.left = 7;
    expect(members[0].left).toBe(7);
    faults.groupMembersSilentOn = [];
    expect(members[0].width).toBe(123);
  });

  it("withholds the group's own size and not its position", async () => {
    const { slide, group } = await drawGrouped(clustered());
    faults.groupSilentOn = ["width", "height"];
    const listed = slide.shapes.items.find((s) => s.id === group.id)!;
    expect(() => listed.width).toThrow(/code=PropertyNotLoaded/);
    expect(() => listed.height).toThrow(/code=PropertyNotLoaded/);
    expect(typeof listed.left).toBe("number");
    expect(typeof listed.top).toBe("number");
    faults.groupSilentOn = ["height"];
    expect(listed.width).toBeGreaterThan(0);
  });

  it("is switched off again by the next host", async () => {
    faults.groupMembersSilentOn = ["name"];
    faults.groupSilentOn = ["width"];
    faults.silentReadsAs = "undefined";
    installHost([makeSlide("s1")]);
    expect(faults.groupMembersSilentOn).toEqual([]);
    expect(faults.groupSilentOn).toEqual([]);
    expect(faults.silentReadsAs).toBe("throws");
  });
});

describe("an edit after a hand resize, when neither the members nor the group will say how big they are", () => {
  it("is the same flow that redraws at one scale when the members DO answer", async () => {
    // The control. Without it the test below could be red for a reason that has
    // nothing to do with what the host withholds.
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = aValueEdit(cfg);
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(sayings(), "an answering host took the in-place route on a resized chart").toHaveLength(0);
    const edit = byName(slide, "seg-0-1");
    const neighbour = byName(slide, "seg-0-0");
    expect(edit.height / 30 / (neighbour.height / 42)).toBeCloseTo(1, 1);
  });

  /**
   * TODAY'S BEHAVIOUR, and it is the residual, not a goal. Nothing the host
   * answers says the chart was stretched, so the edit takes the in-place route
   * and writes the changed bar at the stored config's size beside neighbours at
   * the person's. It is pinned because "if neither answers, behaviour is exactly
   * as today" is a promise, and a promise with no test is a hope.
   */
  it("writes the edited bar at the old scale beside bars at the new one, and says it checked nothing", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    faults.groupMembersSilentOn = ["name", "width", "height"];
    faults.groupSilentOn = ["width", "height"];
    setTracing(true);
    const next = aValueEdit(cfg);
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    const said = sayings();
    expect(
      said,
      "the edit did not take the in-place route, so this is not the blindness it is meant to be",
    ).toHaveLength(1);
    expect(said[0].data).toMatchObject({ checked: { names: 0, sizes: 0 } });
    // Per unit of value, the edited bar against its untouched neighbour: one scale
    // would make these equal, and a bar written at the stored size is half of it.
    const edit = byName(slide, "seg-0-1");
    const neighbour = byName(slide, "seg-0-0");
    const perUnit = edit.height / 30 / (neighbour.height / 42);
    expect(perUnit, "the edited bar is on the same scale as its neighbour").toBeCloseTo(0.5, 1);
    expect(liveShapes(slide).length).toBeGreaterThan(0);
  });
});
