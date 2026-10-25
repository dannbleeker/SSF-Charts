// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  CHART_TAG,
  insertSceneIntoSlide,
  listChartsInDeck,
  updateChartInSlide,
  updateChartsInSlides,
} from "../src/render/powerpoint";
import { setTracing, traceLog } from "../src/core/trace";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import { faults, installHost, makeSlide } from "./helpers/office-host";
import {
  aValueEdit,
  byName,
  clustered,
  drawGrouped,
  edited,
  groupBoxFollowsMembers,
  liveGroups,
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

const routeTaken = () => (sayings().length ? "in place" : "redraw");
const keptTrace = () => traceLog().entries.filter((e) => e.message === "redrawing at the size the user gave the chart");
const declines = () => traceLog().entries.filter((e) => e.message === "not updating in place — redrawing instead");
const storedConfig = (shape: { tagStore: Map<string, string> }): ChartConfig =>
  JSON.parse(shape.tagStore.get(CHART_TAG)!) as ChartConfig;
const sizedTo = (cfg: ChartConfig, width: number, height: number): ChartConfig =>
  edited(cfg, (c) => {
    c.width = width;
    c.height = height;
  });
/** The host that gives a member's id and nothing else, with the group still saying how big it is. */
const silentMembers = () => {
  faults.groupMembersSilentOn = ["name", "width", "height"];
};

describe("an edit after a hand resize, when the members are silent and the group says how big it is", () => {
  /**
   * The case `groupScaleOfScene` exists for. Nothing a member says tells the edit
   * that the chart was stretched, so the group's own box does: it is the one
   * shape the update path already holds and has already loaded the size of.
   */
  it("redraws at one scale and keeps the person's size", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const drawn = { width: group.width, height: group.height };
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    silentMembers();
    setTracing(true);
    const next = aValueEdit(cfg);
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    expect(routeTaken(), "the half-stretched write was made").toBe("redraw");
    const edit = byName(slide, "seg-0-1");
    const neighbour = byName(slide, "seg-0-0");
    expect(edit.height / 30 / (neighbour.height / 42), "the edited bar is on another scale").toBeCloseTo(1, 1);
    // One group, at the size the person made it, within the couple of points the
    // ink margin does not scale by.
    const groups = liveGroups(slide);
    expect(groups).toHaveLength(1);
    expect(Math.abs(groups[0].width - 2 * drawn.width)).toBeLessThan(2);
    expect(Math.abs(groups[0].height - 2 * drawn.height)).toBeLessThan(2);
    expect(back?.size).toEqual({ width: 960, height: 600 });
    expect(storedConfig(groups[0])).toMatchObject({ width: 960, height: 600 });
    // And the trace says whose size it was, and why the in-place route was refused.
    expect(keptTrace()).toHaveLength(1);
    expect(keptTrace()[0].data).toMatchObject({ via: "group", samples: 0 });
    expect(declines().map((d) => d.data?.why)).toEqual([expect.stringContaining("the group's own box")]);
  });

  it("lets the next edit work in place, once the redraw is at the size the config now says", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const first = aValueEdit(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    silentMembers();
    const back = await updateChartInSlide(buildChart(first), target, { tagData: JSON.stringify(first) });
    expect(back?.size).toEqual({ width: 960, height: 600 });

    setTracing(true);
    const second = edited(sizedTo(first, 960, 600), (c) => (c.data.series[1].values[2] = 11));
    await updateChartInSlide(buildChart(second), back!, { tagData: JSON.stringify(second) });
    expect(routeTaken(), "a chart redrawn at its own size was read as resized again").toBe("in place");
    expect(keptTrace()).toEqual([]);
    expect(liveGroups(slide)).toHaveLength(1);
  });

  it("keeps the size through an edit the in-place path declines before it looks at any size", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    silentMembers();
    const next = edited(cfg, (c) => {
      c.data.categories.push("Q5");
      for (const s of c.data.series) s.values.push(12);
    });
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(back?.size).toEqual({ width: 960, height: 600 });
    expect(liveGroups(slide)).toHaveLength(1);
  });

  it("keeps a stretch in one direction to that direction", async () => {
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 1.5, 1);
    const target = (await listChartsInDeck()).charts[0].target;
    silentMembers();
    const back = await updateChartInSlide(buildChart(aValueEdit(cfg)), target, {
      tagData: JSON.stringify(aValueEdit(cfg)),
    });
    // 480 x 1.5 is exact; the other axis read 1, so it is the config's 300.
    expect(back?.size).toEqual({ width: 720, height: 300 });
  });

  it("keeps the stretched chart for an edit that writes no geometry, as r2-2 decided", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 1.5, 1.5);
    const wide = byName(slide, "seg-1-2").width;
    const target = (await listChartsInDeck()).charts[0].target;
    silentMembers();
    setTracing(true);
    const next = edited(cfg, (c) => (c.title = "Renamed"));
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(routeTaken(), "a retitle was redrawn, throwing the person's resize away for nothing").toBe("in place");
    expect(byName(slide, "seg-1-2").width).toBeCloseTo(wide, 6);
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
  });

  it("applies to each chart of a batch at its own size", async () => {
    const a = clustered();
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => true);
    await insertSceneIntoSlide(buildChart(a), { tagData: JSON.stringify(a), left: 40, top: 40 });
    await insertSceneIntoSlide(buildChart(a), { tagData: JSON.stringify(a), left: 40, top: 360 });
    const groups = liveGroups(slide);
    expect(groups).toHaveLength(2);
    for (const g of groups) groupBoxFollowsMembers(g);
    userResizes(groups[0], 1.5, 1.5);
    const charts = (await listChartsInDeck()).charts;
    silentMembers();
    const next = edited(a, (c) => (c.scale = { max: 200 }));
    const back = await updateChartsInSlides(
      charts.map((c) => ({ scene: buildChart(next), target: c.target, opts: { tagData: JSON.stringify(next) } })),
    );
    expect(back.map((t) => t.size).filter(Boolean), "only the resized chart should report a kept size").toEqual([
      { width: 720, height: 450 },
    ]);
    expect(liveGroups(slide)).toHaveLength(2);
  });

  it("is quiet for a chart nobody resized, and takes the fast path as it always did", async () => {
    // The control that matters most: a detector that fires on an untouched chart
    // would redraw every edit on a silent host and write a made-up size into the
    // config of each.
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const ids = (group.grouped as { id: string }[]).map((m) => m.id);
    const target = (await listChartsInDeck()).charts[0].target;
    silentMembers();
    setTracing(true);
    const back = await updateChartInSlide(buildChart(aValueEdit(cfg)), target, {
      tagData: JSON.stringify(aValueEdit(cfg)),
    });
    expect(routeTaken()).toBe("in place");
    expect("size" in back!).toBe(false);
    expect(keptTrace()).toEqual([]);
    expect((liveGroups(slide)[0].grouped as { id: string }[]).map((m) => m.id)).toEqual(ids);
  });

  it("leaves the members to say it once they answer, so their own sizes stay the evidence", async () => {
    // Names withheld, sizes given: a member HAS answered, so the group's box is not
    // consulted and the size is the members' (`samples` counts rects).
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    faults.groupMembersSilentOn = ["name"];
    setTracing(true);
    const back = await updateChartInSlide(buildChart(aValueEdit(cfg)), target, {
      tagData: JSON.stringify(aValueEdit(cfg)),
    });
    expect(back?.size).toEqual({ width: 960, height: 600 });
    expect(keptTrace()[0].data).toMatchObject({ via: "members" });
    expect(keptTrace()[0].data?.samples).toBeGreaterThan(1);
  });

  it("treats a member that gave one side as silent, since a size needs both", async () => {
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    faults.groupMembersSilentOn = ["height"];
    faults.silentReadsAs = "undefined";
    setTracing(true);
    const back = await updateChartInSlide(buildChart(aValueEdit(cfg)), target, {
      tagData: JSON.stringify(aValueEdit(cfg)),
    });
    expect(back?.size).toEqual({ width: 960, height: 600 });
    expect(keptTrace()[0].data).toMatchObject({ via: "group" });
  });

  /**
   * THE WEAKNESS OF A BOX, pinned so nobody reads this as handled. The group's
   * outline moves when ONE bar is dragged out of it exactly as it does when the
   * whole chart is stretched, and the per-rect agreement that tells those apart
   * is what a silent host denies us. With the members answering, the same drag is
   * not a resize; without them it reads as one.
   */
  it("cannot tell one dragged bar from a resize when the members are silent, which is why it waits for them to be", async () => {
    const dragOneBar = async () => {
      const cfg = clustered();
      const { group, members } = await drawGrouped(cfg);
      groupBoxFollowsMembers(group);
      const bar = members.find((m) => m.name === "seg-1-3")!;
      bar.left += 150;
      return { cfg, target: (await listChartsInDeck()).charts[0].target };
    };

    const answering = await dragOneBar();
    setTracing(true);
    const a = await updateChartInSlide(buildChart(aValueEdit(answering.cfg)), answering.target, {
      tagData: JSON.stringify(aValueEdit(answering.cfg)),
    });
    expect("size" in a!, "a dragged bar was taken for a resize by the members").toBe(false);
    setTracing(false);

    const silent = await dragOneBar();
    silentMembers();
    setTracing(true);
    const b = await updateChartInSlide(buildChart(aValueEdit(silent.cfg)), silent.target, {
      tagData: JSON.stringify(aValueEdit(silent.cfg)),
    });
    expect(b?.size, "the group's box no longer reads a dragged bar as a resize — update the note").toBeTruthy();
  });
});

type Host = {
  first: Record<string, unknown>;
  group?: Record<string, unknown>;
  answered: { of: number; names: number; widths: number; heights: number };
  ids: string;
  drawn?: number[] | null;
  live?: number[] | null;
  ratio?: number[] | null;
};
const hostOf = (line: { data?: Record<string, unknown> } | undefined) => line?.data?.host as Host;

describe("what the update's trace lines say the host returned", () => {
  /**
   * `checked {names: 0, sizes: 0}` is what three different hosts produce, and the
   * line could not choose between them. These hold down what now lets a round
   * choose: the raw reading of the first member and the group, how many shapes
   * answered each field, and the group-size detector's own measurement.
   */
  it("reports a healthy chart's readings, and a ratio of 1 for its group", async () => {
    const cfg = clustered();
    const { slide, group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    await updateChartInSlide(buildChart(aValueEdit(cfg)), target, { tagData: JSON.stringify(aValueEdit(cfg)) });
    const host = hostOf(sayings()[0]);
    const nodes = buildChart(cfg).nodes.length;
    expect(host.first).toMatchObject({ name: expect.stringMatching(/^string:/), width: expect.any(Number) });
    expect(host.group).toMatchObject({ name: "string:PowerChart", width: 480, height: 299 });
    expect(host.answered).toMatchObject({ of: nodes, widths: nodes, heights: nodes });
    expect(host.drawn).toEqual([480, 299]);
    expect(host.live).toEqual([480, 299]);
    expect(host.ratio, "the detector is not quiet on a chart nobody touched").toEqual([1, 1]);
    expect(host.ids).toBe("ascending");
    expect(liveGroups(slide)).toHaveLength(1);
  });

  it("says how a withheld read failed, so a throw and an undefined are not the same line", async () => {
    for (const [mode, shown] of [
      ["throws", "threw PropertyNotLoaded"],
      ["undefined", "undefined"],
    ] as const) {
      const cfg = clustered();
      const { group } = await drawGrouped(cfg);
      groupBoxFollowsMembers(group);
      const target = (await listChartsInDeck()).charts[0].target;
      silentMembers();
      faults.groupSilentOn = ["width", "height"];
      faults.silentReadsAs = mode;
      setTracing(true);
      await updateChartInSlide(buildChart(aValueEdit(cfg)), target, { tagData: JSON.stringify(aValueEdit(cfg)) });
      const host = hostOf(sayings()[0]);
      expect(host.first, mode).toEqual({ name: shown, width: shown, height: shown });
      expect(host.group, mode).toMatchObject({ width: shown, height: shown });
      expect(host.answered, mode).toMatchObject({ names: 0, widths: 0, heights: 0 });
      // The group said nothing, so there is a drawn box and no ratio.
      expect(host.drawn).toEqual([480, 299]);
      expect(host.live).toBeNull();
      expect(host.ratio).toBeNull();
      setTracing(false);
      vi.unstubAllGlobals();
    }
  });

  it("is on the decline line too, with the ratio of a chart the members said was resized", async () => {
    // Both sources answer here, which is the comparison a round most wants: the
    // members decline the write, and the group's own ratio is beside it.
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    userResizes(group, 2, 2);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    await updateChartInSlide(buildChart(aValueEdit(cfg)), target, { tagData: JSON.stringify(aValueEdit(cfg)) });
    const line = declines()[0];
    expect(line.data?.why).toMatch(/resized since it was drawn/);
    const host = hostOf(line);
    expect(host.ratio?.[0]).toBeCloseTo(2, 2);
    expect(host.ratio?.[1]).toBeCloseTo(2, 2);
    expect(host.answered.widths).toBe(host.answered.of);
  });

  it("is absent from a decline made before the shapes were read", async () => {
    const cfg = clustered();
    await drawGrouped(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    // A picture is declined first thing, before any shape is looked at.
    await updateChartInSlide(buildChart(cfg), target, { tagData: JSON.stringify(cfg), pictureBase64: "AAAA" });
    const line = declines().find((d) => /draws a picture/.test(String(d.data?.why)));
    expect(line, "the picture decline was not traced").toBeTruthy();
    expect(line!.data).not.toHaveProperty("host");
  });

  it("notices a reordered member by the order of the ids, and records it without acting on it", async () => {
    // The premise — ids ascend in the order the host lists a group — has no
    // evidence from the live host yet, so this is recorded and never a reason to
    // decline: the update below still takes the in-place route.
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const members = group.grouped as { name?: string }[];
    members.push(members.splice(3, 1)[0]);
    const target = (await listChartsInDeck()).charts[0].target;
    silentMembers();
    setTracing(true);
    await updateChartInSlide(buildChart(aValueEdit(cfg)), target, { tagData: JSON.stringify(aValueEdit(cfg)) });
    expect(hostOf(sayings()[0])?.ids).toBe("not ascending");
    expect(routeTaken(), "a reorder the ids showed was acted on without evidence that ids mean order").toBe("in place");
  });

  it("is a few hundred characters, not a dump", async () => {
    const cfg = clustered();
    const { group } = await drawGrouped(cfg);
    groupBoxFollowsMembers(group);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    await updateChartInSlide(buildChart(aValueEdit(cfg)), target, { tagData: JSON.stringify(aValueEdit(cfg)) });
    expect(JSON.stringify(hostOf(sayings()[0])).length).toBeLessThan(700);
  });
});
