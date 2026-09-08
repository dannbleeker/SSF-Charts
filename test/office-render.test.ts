// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  readDeckStyle,
  readDeckStyleWithReason,
  slideImageBase64,
  warmCustomXmlSurface,
  replayDeckStyleVerdict,
  _resetDeckStyleVerdictForTest,
  stallShape,
  traceEnvironment,
  writeDeckStyle,
  _setBatchTimeoutForTest,
  _setBlankReReadDelayForTest,
  CHART_PARTS_TAG,
  CHART_TAG,
  CHART_ORIGIN_TAG,
  getSelectionBounds,
  insertAgendaSlides,
  insertDemoDeck,
  insertSceneIntoSlide,
  shapeGeometryByName,
  marksThisHostWillDrop,
  isPowerPointHost,
  listChartsInDeck,
  scanIsComplete,
  listChartsInSelection,
  loadChartFromSelection,
  onLateSync,
  READBACK_PAGE,
  wantsAutoPicture,
  DEMO_SLOT_TAG,
  DEMO_SHAPE_BUDGET,
  applyReconcilePlan,
  readAddedSlides,
  _setReadbackTimeoutForTest,
  _setDeckInsertPerSlideForTest,
  insertSlidesFromPptx,
  lastLateSyncOwner,
  lastLateSyncSeq,
  waitForLateSync,
  reconcileDeck,
  snapshotAddedSlides,
  updateChartInSlide,
  updateChartsInSlides,
  withSlideDeselected,
  slideHoldsOnlyChart,
  shapesDrawnOn,
  replacedShapeCount,
  getSlideShapeBounds,
  _setSelectionTimeoutForTest,
  _setCountSettleDelayForTest,
  replaceSlideWithDeck,
  deleteShapesById,
  addScratchSlide,
  addSlideForChart,
  deleteSlideById,
  clearShapeSelection,
  MAX_ADD_RETRY_ROUNDS,
  wreckageOf,
  requestStop,
  resetStop,
  slideShots,
  isStopped,
  slideSize,
  _resetSlideSizeCache,
  _setSlideSizeTimeoutForTest,
  _setReReadRetryDelayForTest,
} from "../src/render/powerpoint";

/**
 * Dropped `slides.add()` calls needed to defeat ONE `addSlides` call outright:
 * the original plus every retry round it is allowed.
 *
 * Derived, never hardcoded. These tests used to spell it `2`, which silently
 * stopped meaning "the add and all its retries" the moment the retry bound
 * moved — the assertions still passed for the wrong reason, because a
 * recovered add and a lost one differ only in numbers the test did not check.
 */
const ADDS_TO_DEFEAT_ONE_SLIDE = 1 + MAX_ADD_RETRY_ROUNDS;
import { readFileSync } from "fs";
import { syncsSoFar, resetSyncCount, slideCount } from "../src/render/powerpoint";
import { onTrace, setTracing, traceAbout, traceLog } from "../src/core/trace";
import { planReconcile } from "../src/core/reconcile";
import { planSceneUpdate, worthUpdating } from "../src/core/scene-diff";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { estimateOfficeShapes } from "../src/core/scene";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildAgendaScene } from "../src/core/agenda";
import type { ChartConfig, MarkerSymbol } from "../src/core/types";
import type { DemoReport } from "../src/render/powerpoint";

/** The indices a demo run did not render as a real chart (skipped or failed). */
const failedIndices = (r: DemoReport) =>
  r.results.map((x, i) => (x.status !== "rendered" ? i : -1)).filter((i) => i >= 0);

import {
  addedWithLayout,
  addedWithMaster,
  blankReadbackAt,
  failSyncsOn,
  faults,
  hostSlideSize,
  installHost,
  makeShape,
  makeSlide,
  stallSyncOn,
  trips,
  unansweredNullChecks,
  untracked,
  lastShapeLoadSpec,
  type FakeShape,
  type FakeSlide,
} from "./helpers/office-host";

/**
 * The shape a chart's tags actually landed on — asked, never assumed.
 *
 * These assertions used to read `slide.created[0]` and were rewritten when the
 * anchor moved to the last shape drawn. That move has since been reverted, and
 * this helper deliberately did NOT go back with it: a hardcoded index had to be
 * edited in six places to follow one experiment, and would need editing again
 * for the next. Asking which shape carries the tag is what every one of these
 * tests meant in the first place, and it is true wherever the anchor sits.
 */
const taggedShape = (slide: FakeSlide, key = CHART_TAG) => slide.created.find((s) => s.tagStore.get(key));

const config: ChartConfig = {
  kind: "stacked",
  ...DEFAULT_SIZE,
  data: {
    categories: ["A", "B"],
    series: [
      { name: "S1", values: [3, 4] },
      { name: "S2", values: [1, 2] },
    ],
  },
};

afterEach(() => vi.unstubAllGlobals());

describe("insertSceneIntoSlide", () => {
  it("creates native shapes at the requested offset, groups, and tags", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { left: 100, top: 50, tagData: JSON.stringify(config) });

    const rects = slide.created.filter((s) => s.geo === "rectangle");
    expect(rects.length).toBeGreaterThanOrEqual(4); // one per stacked segment
    for (const r of rects) expect(r.box.left).toBeGreaterThanOrEqual(100);
    const group = slide.created.find((s) => s.type === "group")!;
    expect(group.name).toBe("PowerChart");
    expect(group.grouped).toHaveLength(slide.created.length - 1);
    expect(group.tagStore.get(CHART_TAG)).toBe(JSON.stringify(config));
  });

  it("draws every hex tile as a fillable hexagon, not an outline", async () => {
    /**
     * THE ONE CHART WHOSE ENTIRE MESSAGE IS FILL, and until this test nothing
     * had ever put it through THIS renderer.
     *
     * Its tiles were PolygonNodes, which Office.js cannot fill — it draws a
     * polygon as one line per edge — so the add-in rendered a 51-region
     * choropleth as 51 hollow rings while the SVG preview beside it drew them
     * solid. They are SymbolNodes naming the native `hexagon` preset now.
     *
     * That change shipped green with the tile map covered only at SCENE level,
     * and the fake caught it the moment a test finally rendered one: its
     * `GeometricShapeType` Proxy throws on a preset it has not been told about,
     * exactly so an unknown name cannot come back `undefined` and be drawn as a
     * shape with no geometry at all. This asserts the consequence instead — a
     * count of FILLED hexagons carrying more than one colour.
     */
    const cfg = {
      kind: "tilemap",
      width: 480,
      height: 300,
      map: "us",
      tilemap: { shape: "hex" },
      data: { categories: ["CA", "TX", "NY"], series: [{ name: "S", values: [100, 80, 60] }] },
    } as unknown as typeof config;
    const slide = makeSlide("s-hex");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(cfg), { left: 0, top: 0 });

    const hexes = slide.created.filter((s) => s.geo === "hexagon");
    // Every region in the layout gets a tile, with or without data.
    expect(hexes.length).toBeGreaterThan(40);
    // `fillColor`, NOT `fill`. `fill` is the API object — `setSolidColor` and
    // friends — and there is a fresh one per shape, so a Set of them has one
    // entry per tile and "more than one colour" passes on a slide painted
    // entirely in white. `fillColor` is where the colour actually lands.
    expect(hexes.filter((h) => !h.fillColor).length).toBe(0);
    expect(new Set(hexes.map((h) => String(h.fillColor))).size).toBeGreaterThan(1);
  });

  it("still draws something when the host's enum has no such preset", async () => {
    /**
     * A real host's `GeometricShapeType` answers `undefined` for a name it does
     * not carry, and `addGeometricShape(undefined, …)` is ACCEPTED — it draws a
     * shape with no geometry, which is invisible on the slide and unexplainable
     * from the file. Our table is not the host's enum, and nothing checked that.
     *
     * An ellipse in the wrong place is a bug someone can see and report. Fifty
     * invisible tiles are a chart that silently lost its data, which is the
     * failure mode this repo keeps deciding is the worse one.
     */
    faults.presetMissing = "hexagon";
    try {
      const cfg = {
        kind: "tilemap",
        width: 480,
        height: 300,
        map: "us",
        tilemap: { shape: "hex" },
        data: { categories: ["CA", "TX"], series: [{ name: "S", values: [10, 20] }] },
      } as unknown as typeof config;
      const slide = makeSlide("s-nohex");
      installHost([slide]);
      await insertSceneIntoSlide(buildChart(cfg), { left: 0, top: 0 });
      // BY NAME, not by "anything without a geo" — a group and a text box have
      // no preset either, and a test that counted those would pass on a slide
      // where every tile was invisible.
      // COUNTS AND STRINGS ONLY, never a FakeShape handed to `expect`. A failing
      // assertion makes vitest serialize the actual value, serializing walks
      // every getter, and `group` on a non-group throws "shape is not a group"
      // — which then REPLACES the assertion error and reports the fake instead
      // of the defect. It cost twenty minutes here; do not hand these to expect.
      // `tile-XX`, NOT `startsWith("tile-")` — that also catches the
      // `tile-code-XX` text labels, which have no preset and are supposed to
      // have none. Counting them made this read "52 tiles drew no geometry" on
      // a slide where every tile was fine.
      const tiles = slide.created.filter((s) => /^tile-[A-Z]{2}$/.test(String(s.name ?? "")));
      expect(tiles.length).toBeGreaterThan(40);
      // Not one drawn with no geometry, and all of them on the preset every
      // host has.
      expect(tiles.filter((t) => t.geo === undefined).length).toBe(0);
      expect(tiles.filter((t) => t.geo !== "ellipse").length).toBe(0);
      // …and the choropleth still carries its value.
      expect(new Set(tiles.map((t) => String(t.fillColor))).size).toBeGreaterThan(1);
    } finally {
      faults.presetMissing = "";
    }
  });

  describe("what this host cannot draw", () => {
    /**
     * `pane-host-actions.test.ts` mocks this whole module, so the pane tests
     * prove the pane REACTS to the answer and can say nothing about whether the
     * answer is right. This asks the real function.
     *
     * The list is exactly the node kinds that draw NOTHING without
     * `Shape.rotation`: a wedge and an arrowhead. A diagonal `line` must never
     * appear — `addSegment` falls back to a real line and still draws, so
     * counting it would rasterise every line chart on every desktop Office for
     * no reason at all.
     */
    const scene = (kind: string) =>
      buildChart({
        kind,
        width: 480,
        height: 300,
        data: {
          categories: ["A", "B", "C"],
          series: [{ name: "S", values: [40, 35, 25] }],
        },
      } as unknown as typeof config);

    it("names the pie's slices on a host without rotation", () => {
      const slide = makeSlide("s-caps");
      installHost([slide], [], slide, (v) => v !== "1.10");
      const got = marksThisHostWillDrop(scene("pie"));
      expect(got.nodes, "a pie lost nothing on a host that cannot draw a wedge").toBeGreaterThan(0);
      expect(got.what.join(" "), "did not say what the user would lose").toMatch(/pie slice/);
    });

    it("says a line chart loses nothing, because its diagonals fall back", () => {
      // The one that would be expensive to get wrong: `addSegment` routes a
      // diagonal to a real line without rotation. Reporting a loss here would
      // turn every line chart on every desktop Office into a picture.
      const slide = makeSlide("s-line");
      installHost([slide], [], slide, (v) => v !== "1.10");
      expect(marksThisHostWillDrop(scene("line")).nodes).toBe(0);
      expect(marksThisHostWillDrop(scene("clustered")).nodes).toBe(0);
    });

    it("counts in words a person would use, singular and plural", () => {
      // The message is the whole product of this function, and "1 pie slices"
      // is the kind of thing that makes a warning look automated and skippable.
      // Built as a scene directly: no chart kind emits exactly one wedge beside
      // exactly one arrow, and the wording is what is under test.
      const slide = makeSlide("s-words");
      installHost([slide], [], slide, (v) => v !== "1.10");
      const sceneOf = (nodes: unknown[]) =>
        marksThisHostWillDrop({ width: 100, height: 100, nodes } as unknown as ReturnType<typeof buildChart>);
      const wedge = { kind: "wedge", cx: 50, cy: 50, r: 20, innerR: 0, startAngle: 0, endAngle: 90, fill: "#111" };
      const arrow = { kind: "arrowhead", x: 10, y: 10, angle: 0, size: 4, fill: "#111" };
      const one = sceneOf([wedge, arrow]);
      expect(one.what).toEqual(["a pie slice", "an arrow"]);
      expect(one.nodes).toBe(2);
      // …and the slices lead, because where both are present the slices are the
      // subject and the arrows are the annotation.
      expect(sceneOf([wedge, wedge, arrow, arrow]).what).toEqual(["2 pie slices", "2 arrows"]);
    });

    it("loses nothing at all on a host that has rotation", () => {
      const slide = makeSlide("s-full");
      installHost([slide]);
      for (const k of ["pie", "doughnut", "line", "clustered"]) {
        expect(marksThisHostWillDrop(scene(k)), `${k} reported a loss on a 1.10 host`).toEqual({ what: [], nodes: 0 });
      }
    });
  });

  describe("reading a chart's parts back off a real slide", () => {
    /**
     * `shapeGeometryByName` is how a scenario measures what actually landed, and
     * on a real slide the parts are NOT on the slide: `insertSceneIntoSlide`
     * groups them, so `slide.shapes` answers with one `PowerChart` per chart.
     * Rounds 334 and 335 both skipped `where a rotated shape lands` on that.
     *
     * The refusals below matter as much as the happy path. Reaching into a
     * group costs a `Shape.group` access, and on a shape that is NOT a group
     * that throws GeneralException and POISONS the sync it rides in — the
     * archive carries it 76 times across 38 rounds. So the descent has to be
     * able to give up in three places without taking the reading down with it.
     */
    const drawLineChart = async (slideId: string) => {
      const slide = makeSlide(slideId);
      installHost([slide]);
      await insertSceneIntoSlide(
        buildChart({
          kind: "line",
          width: 160,
          height: 220,
          data: { categories: ["a", "b", "c"], series: [{ name: "S", values: [1, 9, 2] }] },
        } as unknown as typeof config),
        { left: 0, top: 0 },
      );
      return slide;
    };

    it("finds the segments inside the group when the slide lists only the group", async () => {
      const slide = await drawLineChart("s-grp");
      faults.groupHidesChildren = true;
      try {
        const got = await shapeGeometryByName(slide.id, (n) => /^line-\d+-\d+$/.test(n));
        expect(got, "the reader gave up entirely").toBeTruthy();
        expect(got!.length, "found no segments inside the group").toBeGreaterThan(0);
        expect(got!.filter((g) => !Number.isFinite(g.width)).length).toBe(0);
      } finally {
        faults.groupHidesChildren = false;
      }
    });

    it("needs no group API when nothing grouped the parts", async () => {
      /**
       * Written expecting an empty reading, and the host corrected it: below
       * 1.8 the renderer does not GROUP either, so the segments stay on the
       * slide and are found without ever reaching for `Shape.group`. The
       * descent is for hosts that group, which are the hosts that can be asked.
       *
       * Kept rather than deleted, because the pairing is the point: the two
       * abilities travel together, and a reading that came back empty here
       * would mean the renderer had started grouping on a host that cannot open
       * a group again.
       */
      const slide = makeSlide("s-no18");
      installHost([slide], [], slide, (v) => v !== "1.8");
      await insertSceneIntoSlide(
        buildChart({
          kind: "line",
          width: 160,
          height: 220,
          data: { categories: ["a", "b"], series: [{ name: "S", values: [1, 9] }] },
        } as unknown as typeof config),
        { left: 0, top: 0 },
      );
      const got = await shapeGeometryByName(slide.id, (n) => /^line-\d+-\d+$/.test(n));
      expect(got!.length, "a host that does not group still has its parts on the slide").toBeGreaterThan(0);
    });

    it("survives a group that refuses to open", async () => {
      const slide = await drawLineChart("s-refuse");
      faults.groupHidesChildren = true;
      faults.refuseGroupRead = true;
      try {
        const got = await shapeGeometryByName(slide.id, (n) => /^line-\d+-\d+$/.test(n));
        // Nothing found is the right answer; taking the whole read down is not.
        expect(got ?? []).toEqual([]);
      } finally {
        faults.refuseGroupRead = false;
        faults.groupHidesChildren = false;
      }
    });

    it("matches nothing when no name matches, without reaching into a group", async () => {
      const slide = await drawLineChart("s-nomatch");
      const got = await shapeGeometryByName(slide.id, () => false);
      expect(got).toEqual([]);
    });
  });

  it("keeps the config tag when the host refuses to say where the chart landed", async () => {
    // The failure five real rounds reported and none could reproduce. `same
    // scale across the deck` scores 4 of 8 every time, and every loss is the
    // same 5010 — with the host naming the culprit itself:
    //
    //   errorLocation: ShapeCollection.getItem
    //   statement: var shape = shapes.getItem(...) /* originally addTextBox(...) */;
    //
    // "originally addTextBox" — Office.js rewrote the CREATION proxy's object
    // path into `getItem(id)`, and this host will not resolve that for a shape
    // it has just made. What forces the rewrite is the `load("id,left,top")`
    // asking where the chart ended up, and that load sits in the same batch as
    // the tag writes. One refused positional read therefore takes the config
    // tag with it, and the chart is left drawn, nameless and not re-editable.
    //
    // BEHAVIOUR, NOT A GUARD — and said out loud because this test was written
    // expecting to fail. The theory was that the positional load and the tag
    // writes share a batch, so one refusal costs both; the insert path turns out
    // to already survive it, and the test passed the moment it was written. It
    // is kept because it pins something real that CI could not reach before, and
    // because the next person to have this idea should find it already answered.
    //
    // Where the loss actually happens is the UPDATE path, and the round says so
    // in one field: every settle-pass failure reports `withId: 0`. The repair
    // that exists to rescue a lost config tag is id-based, and this host never
    // yields an id — so the second chance is unavailable by construction rather
    // than failing. See `docs/BACKLOG.md`.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseIdLeftTopLoads = 1;
    try {
      await insertSceneIntoSlide(buildChart(config), { left: 100, top: 50, tagData: JSON.stringify(config) });
    } finally {
      faults.refuseIdLeftTopLoads = 0;
    }
    const tagged = slide.created.find((s) => s.tagStore.get(CHART_TAG));
    expect(tagged, "drawn but carrying no config — the user cannot edit it again").toBeTruthy();
    expect(tagged!.tagStore.get(CHART_TAG)).toBe(JSON.stringify(config));
  });

  it("says WHICH handle the failed tag write went through", async () => {
    // The instrument the last six rounds needed and did not have. Every one of
    // them recorded `tagging failed` with a count and an error and nothing about
    // the handle, and there are four routes to a tag target — two of which this
    // host resolves and refuses (`refreshed`, `by-id`) and two of which it
    // accepts (`created`, `group`). Without this field the trace says a write
    // failed; with it, the trace says it failed THROUGH A RESOLVED HANDLE, which
    // is the difference between a fix and another round.
    setTracing(true);
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseTagWritesOnResolvedProxy = true;
    try {
      await insertSceneIntoSlide(buildChart(config), {
        tagData: JSON.stringify(config),
        group: false,
        shapesPerSync: 1,
      });
      const failed = traceLog().entries.find((e) => e.message.startsWith("tagging failed"));
      expect(failed, "the write was refused and the trace did not say so").toBeDefined();
      expect(failed!.data!.from, "named no handle — the whole point of the field").toBeTruthy();
      // `created×1`, and that answer is the field earning its place on its first
      // outing. The guess going in was `refreshed` — the pre-grouping re-read
      // being the obvious suspect — and the trace says otherwise: the target was
      // the handle that DREW the shape, and the write was still refused, because
      // the pass loads the created shapes too before it writes (the parts tag
      // reads their ids). That is why swapping the target does not fix this and
      // the ordering has to change, and it took one run to see instead of a
      // round. See `docs/BACKLOG.md`.
      expect(String(failed!.data!.from)).toMatch(/^(created|refreshed|group|by-id)×\d+/);
      // AND WHICH SLIDE, which a lead in the archive cannot be tested without.
      // Rounds 043-045 each lost a chart's config, and the one line that names a
      // slide named `257#0` every time — an id whose second half is `0`, which
      // is not the shape this host gives a slide it has finished adding. Two
      // such ids turn up among every round's added slides. With the slide on
      // this line, "are the charts that lose their tag the ones sitting on those
      // slides" becomes a join instead of an inference off a settle-pass line.
      expect(failed!.data!.slides, "the failure names no slide, so the #0 lead stays untestable").toBeTruthy();
    } finally {
      faults.refuseTagWritesOnResolvedProxy = false;
      setTracing(false);
    }
  });

  it("survives a refused resolved handle here, the way the real host does not", async () => {
    // Six real rounds, one failure, and the host naming its own cause:
    //
    //   errorLocation: ShapeCollection.getItem
    //   statement: var shape = shapes.getItem(...) /* originally addTextBox(...) */
    //
    // `same scale across the deck` loses half the deck's config tags to that,
    // and the settle pass cannot repair them because it resolves by id and this
    // host yields none (`withId: 0`, on every failure).
    //
    // The rule is RESOLUTION, not age. `tag-the-creation-proxy-a-sync-later`
    // answers `yes` four rounds running, so the handle that made a shape keeps
    // working however old it gets; a `load()` is what makes Office.js rewrite it
    // into `shapes.getItem(id)`, which this host refuses for a shape it has just
    // created.
    //
    // `finishCharts` re-reads shapes before grouping and then tags through the
    // re-read handle. The re-read is needed FOR GROUPING. Tagging through it is
    // what costs the user their chart.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseTagWritesOnResolvedProxy = true;
    // ON, or the `tagging failed` assertion below is vacuous — an empty trace
    // holds no failure line for the same reason it holds nothing at all.
    setTracing(true);
    try {
      // `shapesPerSync: 1` is what makes this the real case: a chart that spans
      // sync batches sets `refreshShapes`, which is the only thing that triggers
      // the pre-grouping re-read — and the re-read handle is the poisoned one.
      // Every chart in a real round spans batches; none of the small fixtures
      // here did, which is why CI has never seen this.
      await insertSceneIntoSlide(buildChart(config), {
        tagData: JSON.stringify(config),
        group: false,
        shapesPerSync: 1,
      });
    } finally {
      faults.refuseTagWritesOnResolvedProxy = false;
      setTracing(false);
    }
    // The tag SURVIVES here, and the gap between this and the real host is the
    // point. The drawing context's write is refused exactly as it is in a real
    // round; the settle pass then repairs it, because the fake can resolve a
    // shape by id. The real host cannot — `withId: 0` on every failure across
    // six rounds — so there the chart stays nameless.
    //
    // So this pins the first half of the reproduction and names the second.
    // Arming `refuseShapeIdLoads` alongside was tried and models something
    // harsher than the host: it makes the insert throw outright, where a real
    // round carries on and merely loses the tag.
    //
    // IT ASSERTED THE OPPOSITE BETWEEN 2026-08-15 AND 2026-08-16 — that the
    // drawing context's own write landed with no repair — because the tag anchor
    // had been moved to a shape no `load()` resolved, which left this fault
    // nothing to refuse. That change measured no effect on the real host across
    // five rounds and four builds and has been reverted, so the weaker claim is
    // the true one again. The trace assertion below is what says which half is
    // being pinned, and it is deliberately the pessimistic one.
    expect(traceLog().entries.length, "tracing was off, so the assertion below proves nothing").toBeGreaterThan(0);
    const tagged = slide.created.find((s) => s.tagStore.get(CHART_TAG));
    expect(tagged, "the settle pass should have repaired what the drawing context lost").toBeTruthy();
    expect(tagged!.tagStore.get(CHART_TAG)).toBe(JSON.stringify(config));
    // REPAIRED, not landed first time — and the failure line is the evidence
    // that the drawing context's write was refused at all. Without it this test
    // would pass on a host that never refused anything, which is the one thing
    // it is not for.
    expect(
      traceLog().entries.some((e) => e.message.startsWith("tagging failed")),
      "the drawing context's write was supposed to be refused here, and the trace does not say it was",
    ).toBe(true);
  });

  it("describes the chart group with accessible alt text", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(config) });
    const group = slide.created.find((s) => s.type === "group") as FakeShape & {
      altTextDescription?: string;
      altTextTitle?: string;
    };
    expect(group.altTextDescription).toBe(scene.desc);
    if (scene.title) expect(group.altTextTitle).toBe(scene.title);
  });

  it("describes the chart even when it is NOT grouped", async () => {
    // The alt text used to be assigned only on the group object, so every
    // ungrouped chart (group:false, a refused addGroup, a one-shape chart)
    // silently lost its text alternative. It belongs on whatever shape stands
    // for the chart — the same one the config tag lands on.
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg", group: false });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    const anchor = taggedShape(slide) as FakeShape & { altTextDescription?: string; altTextTitle?: string };
    expect(anchor.altTextDescription).toBe(scene.desc);
    expect(anchor.tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("skips alt text on a host below 1.10 instead of losing the group with it", async () => {
    // Shape.altTextDescription is PowerPointApi 1.10. Assigning it on an older
    // host is a queued command rejected at the next sync — the same sync that
    // carries the grouping, so an ungated assignment costs the group.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.10");
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group") as FakeShape & { altTextDescription?: string };
    expect(group).toBeDefined();
    expect(group.altTextDescription).toBeUndefined();
  });

  it("renders a pie as a rotated triangle fan", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart({
      ...config,
      kind: "pie",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
    });
    await insertSceneIntoSlide(scene, {});
    const tris = slide.created.filter((s) => s.geo === "triangle" && s.name?.includes("-f"));
    expect(tris.length).toBeGreaterThan(10);
    for (const t of tris) expect(typeof t.rotation).toBe("number");
  });

  it("maps title font and alignment onto text boxes", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart({ ...config, title: "Hello" }), { fontFamily: "Arial" });
    const title = slide.created.find((s) => s.text === "Hello")!;
    expect(title.fillCleared).toBe(true);
    expect(title.textFrame.textRange.font).toMatchObject({ name: "Arial", bold: true });
  });

  it("draws value lines as dashed native connectors", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(
      buildChart({ ...config, decorations: { valueLines: [{ mode: "mean" }], segmentLabels: true } }),
      {},
    );
    const dashed = slide.created.filter((s) => s.type === "line" && s.lineFormat.dashStyle === "dash");
    expect(dashed.length).toBeGreaterThanOrEqual(1);
  });
});

describe("scene node mapping", () => {
  const insert = async (nodes: object[], opts = {}) => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide({ width: 200, height: 100, nodes } as never, opts);
    return slide;
  };

  it("maps ellipses with stroke or hidden outline", async () => {
    const slide = await insert([
      {
        kind: "ellipse",
        cx: 50,
        cy: 50,
        rx: 20,
        ry: 10,
        fill: "#ff0000",
        stroke: "#000000",
        strokeWidth: 2,
        name: "dot",
      },
      { kind: "ellipse", cx: 10, cy: 10, rx: 5, ry: 5, fill: "#00ff00" },
    ]);
    const [a, b] = slide.created.filter((s) => s.geo === "ellipse");
    // center − radius, plus the default 60/90pt insert offset
    expect(a.box).toEqual({ left: 90, top: 130, width: 40, height: 20 });
    expect(a.lineFormat).toMatchObject({ color: "#000000", weight: 2 });
    expect(b.lineFormat.visible).toBe(false);
  });

  it("never hands Office a font size it rejects", async () => {
    // The THIRD sink, and the worst place to be missing this guard. `font.size`
    // was assigned straight from the scene node, where the skill's pptx sink
    // clamps to OOXML's 1..4000pt and the SVG sink substitutes for a non-finite
    // value. Office.js REJECTS an out-of-range size, and a rejected property
    // throws — which on this host takes the rest of the batch with it, config
    // tag included, so the user gets a chart that is not re-editable or no
    // chart at all. A clamped label is merely small.
    const slide = await insert([
      { kind: "text", x: 0, y: 0, w: 50, h: 12, text: "zero", fontSize: 0, color: "#000", name: "a" },
      { kind: "text", x: 0, y: 20, w: 50, h: 12, text: "negative", fontSize: -5, color: "#000", name: "b" },
      { kind: "text", x: 0, y: 40, w: 50, h: 12, text: "enormous", fontSize: 1e6, color: "#000", name: "c" },
      { kind: "text", x: 0, y: 60, w: 50, h: 12, text: "nan", fontSize: Number.NaN, color: "#000", name: "d" },
      { kind: "text", x: 0, y: 80, w: 50, h: 12, text: "ordinary", fontSize: 11.5, color: "#000", name: "e" },
    ]);
    const sizes = slide.created
      .filter((sh) => sh.textFrame?.textRange?.font?.size !== undefined)
      .map((sh) => sh.textFrame.textRange.font.size as number);
    expect(sizes).toHaveLength(5);
    for (const size of sizes) {
      expect(Number.isFinite(size)).toBe(true);
      expect(size).toBeGreaterThanOrEqual(1);
      expect(size).toBeLessThanOrEqual(4000);
    }
    // The negative control: a size Office accepts is passed through untouched,
    // so this is a clamp and not a rewrite of every label in the deck.
    expect(sizes).toContain(11.5);
  });

  it("honours an 8-digit #RRGGBBAA fill: 6-digit hue + transparency, never mis-parsed", async () => {
    // #RRGGBBAA is a valid hand-authored colour that the SVG preview and the
    // skill's pptx render translucent. Office.js setSolidColor validates 6-digit
    // hex only, so the alpha byte has to move to fill.transparency (1.4) or the
    // live add-in would mis-parse the value and lose the colour.
    const slide = await insert([{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#4e79a780", name: "band" }]);
    const rect = slide.created.find((s) => s.geo === "rectangle") as (typeof slide.created)[number] & {
      fill: { transparency?: number };
    };
    expect(rect.fillColor).toBe("#4e79a7"); // hue survives, no 8-digit string reaches Office.js
    expect(rect.fill.transparency).toBeCloseTo(1 - 128 / 255, 3);
  });

  it("leaves an opaque 6-digit fill untouched (no transparency set)", async () => {
    const slide = await insert([{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#4e79a7", name: "bar" }]);
    const rect = slide.created.find((s) => s.geo === "rectangle") as (typeof slide.created)[number] & {
      fill: { transparency?: number };
    };
    expect(rect.fillColor).toBe("#4e79a7");
    expect(rect.fill.transparency).toBeUndefined();
  });

  it("renders a fill:'none' rect as an outlined/hollow shape (IBCS plan columns)", async () => {
    const slide = await insert([
      { kind: "rect", x: 0, y: 0, w: 20, h: 40, fill: "none", stroke: "#3b6ea5", strokeWidth: 1.5, name: "pl" },
    ]);
    const rect = slide.created.find((s) => s.geo === "rectangle")!;
    expect(rect.fillCleared).toBe(true); // no fill
    expect(rect.fillColor).toBeNull(); // never mis-parsed as a colour
    expect(rect.lineFormat).toMatchObject({ color: "#3b6ea5", weight: 1.5 });
  });

  it("maps chevrons to chevron/homePlate geometry", async () => {
    const slide = await insert([
      { kind: "chevron", x: 0, y: 0, w: 40, h: 20, fill: "#123456", flatLeft: true },
      { kind: "chevron", x: 50, y: 0, w: 40, h: 20, fill: "#123456" },
    ]);
    expect(slide.created.filter((s) => s.type !== "group").map((s) => s.geo)).toEqual(["homePlate", "chevron"]);
  });

  it("draws axis-aligned lines with a clamped non-zero box (never a degenerate diagonal)", async () => {
    const slide = await insert([
      {
        kind: "line",
        x1: 10,
        y1: 50,
        x2: 200,
        y2: 50,
        stroke: "#333333",
        strokeWidth: 1,
        dash: [3, 2],
        name: "connector",
      },
    ]);
    const line = slide.created.find((s) => s.type === "line")!;
    // Horizontal line: width spans, height is clamped up from 0 so the web host
    // can't blow a zero-thickness box into a giant diagonal.
    expect(line.box.width).toBeGreaterThan(180);
    expect(line.box.height).toBeGreaterThanOrEqual(0.5);
    expect(line.lineFormat.dashStyle).toBe("dash");
  });

  it("honours a translucent stroke on BOTH the line and the rotated-rect branch", async () => {
    // The two branches of addSegment diverged: the axis-aligned/dashed line
    // dropped the alpha byte while the diagonal rotated-rect folded it into
    // transparency, so one series colour rendered at two opacities. Both must now
    // agree — a 50% alpha (#…80) → transparency ≈ 0.498, colour a bare 6-digit hex.
    const alpha = "#33333380";
    const horiz = await insert([{ kind: "line", x1: 10, y1: 50, x2: 200, y2: 50, stroke: alpha, name: "h" }]);
    const line = horiz.created.find((s) => s.type === "line")!;
    expect(line.lineFormat.color).toBe("#333333"); // alpha byte stripped off the colour
    expect(line.lineFormat.transparency).toBeCloseTo(1 - 0x80 / 255, 4); // …but carried here

    const diag = await insert([{ kind: "line", x1: 0, y1: 0, x2: 100, y2: 100, stroke: alpha, name: "d" }]);
    const rect = diag.created.find((s) => s.geo === "rectangle")!;
    expect(rect.fillColor).toBe("#333333");
    expect((rect.fill as unknown as { transparency?: number }).transparency).toBeCloseTo(1 - 0x80 / 255, 4);
  });

  it("draws diagonal lines as thin rotated rectangles (direction-correct on every host)", async () => {
    // Up-right and down-right diagonals a bounding box alone can't distinguish.
    const down = await insert([
      { kind: "line", x1: 0, y1: 0, x2: 100, y2: 100, stroke: "#a00000", strokeWidth: 2, name: "d" },
    ]);
    const dr = down.created.find((s) => s.geo === "rectangle")!;
    expect(dr).toBeTruthy();
    expect(dr.fillColor).toBe("#a00000");
    expect(dr.rotation).toBeCloseTo(45, 0); // down-right
    expect(down.created.some((s) => s.type === "line")).toBe(false);

    const up = await insert([
      { kind: "line", x1: 0, y1: 100, x2: 100, y2: 0, stroke: "#00a000", strokeWidth: 2, name: "u" },
    ]);
    const ur = up.created.find((s) => s.geo === "rectangle")!;
    expect(ur.rotation).toBeCloseTo(-45, 0); // up-right — the case a box would mirror
  });

  it("draws dashed diagonals as real line shapes, picking the geometry per direction", async () => {
    // A rotated rectangle carries its colour in its fill, which can't be
    // dashed — scatter trend lines and forecast segments came out solid.
    const down = await insert([
      {
        kind: "line",
        x1: 0,
        y1: 0,
        x2: 100,
        y2: 60,
        stroke: "#a00000",
        strokeWidth: 1.25,
        dash: [4, 2],
        name: "trend",
      },
    ]);
    const dl = down.created.find((s) => s.name === "trend")!;
    expect(dl.type).toBe("line"); // not a filled rectangle
    expect(dl.lineFormat.dashStyle).toBe("dash");
    expect(dl.box).toMatchObject({ width: 100, height: 60 });

    // Up-right: addLine only ever draws the box's top-left→bottom-right
    // diagonal, so this direction needs the lineInverse geometry.
    const up = await insert([
      {
        kind: "line",
        x1: 0,
        y1: 60,
        x2: 100,
        y2: 0,
        stroke: "#a00000",
        strokeWidth: 1.25,
        dash: [4, 2],
        name: "trend",
      },
    ]);
    const ul = up.created.find((s) => s.name === "trend")!;
    expect(ul.geo).toBe("lineInverse");
    expect(ul.lineFormat.dashStyle).toBe("dash");
  });

  it("renders a dotted array as roundDot, not a generic dash", async () => {
    // [1.5,1.5] is the dotted waterfall carry connector. It used to flatten to
    // the same dash enum as every other pattern, so the deck lost the dotted
    // look the SVG preview shows.
    const slide = await insert([
      {
        kind: "line",
        x1: 10,
        y1: 50,
        x2: 200,
        y2: 50,
        stroke: "#333",
        strokeWidth: 1,
        dash: [1.5, 1.5],
        name: "carry",
      },
    ]);
    const line = slide.created.find((s) => s.name === "carry")!;
    expect(line.lineFormat.dashStyle).toBe("roundDot");
  });

  /**
   * The live sink's half of the three-sink dash sweep (the two offline sinks are
   * swept against the same corpus in `geometry.test.ts`). This renderer guarded
   * on `s.dash` being TRUTHY, and `[]` is truthy — so an empty dash array set
   * `roundDot` here while the SVG preview drew a solid line. `dashKind` answers
   * `none` for it now, and no dash style is written at all.
   */
  it("leaves a line undashed when its dash array specifies no dash", async () => {
    const slide = await insert(
      [[], [0, 0], [-5, -5], [NaN], [-1, 4]].map((dash, i) => ({
        kind: "line" as const,
        x1: 10,
        y1: 20 + i * 10,
        x2: 200,
        y2: 20 + i * 10,
        stroke: "#333",
        strokeWidth: 1,
        dash,
        name: `undashed-${i}`,
      })),
    );
    for (let i = 0; i < 5; i++) {
      const l = slide.created.find((s) => s.name === `undashed-${i}`)!;
      expect(l.lineFormat.dashStyle, `line ${i} was given a dash style`).toBeUndefined();
    }
  });

  it("draws polygon edges direction-correct, with no zero-thickness boxes", async () => {
    // A violin body: an up-right edge, a horizontal edge and a down-right edge.
    const slide = await insert([
      {
        kind: "polygon",
        points: [
          { x: 0, y: 40 },
          { x: 50, y: 0 },
          { x: 100, y: 40 },
          { x: 100, y: 40 },
        ],
        fill: "#eeeeee",
        stroke: "#3366cc",
        strokeWidth: 1,
        name: "violin-0",
      },
    ]);
    const edges = slide.created.filter((s) => s.name?.startsWith("violin-0-e"));
    expect(edges).toHaveLength(4);
    for (const e of edges) {
      // Every edge is a real segment: a bounding box collapsed to zero on one
      // axis let the web host blow it up into a giant diagonal.
      expect(e.box.width).toBeGreaterThanOrEqual(0.5);
      expect(e.box.height).toBeGreaterThanOrEqual(0.5);
    }
    // Edge 0 (0,40)->(50,0) rises to the right; edge 1 (50,0)->(100,40) falls.
    // Passing both bounding boxes to addLine drew them as the same diagonal.
    const [e0, e1] = edges;
    expect(e0.rotation).toBeLessThan(0);
    expect(e1.rotation).toBeGreaterThan(0);
  });

  it("maps arrowheads to rotated triangles anchored at the tip", async () => {
    const slide = await insert([{ kind: "arrowhead", x: 10, y: 10, size: 4, angle: 45, fill: "#000000", name: "ah" }]);
    const tri = slide.created[0];
    expect(tri.geo).toBe("triangle");
    expect(tri.rotation).toBe(135); // scene angle + 90
    // The triangle's tip (box top-centre, rotated θ about the box centre) must
    // land on the scene point (10,10) + the default 60/90pt insert offset.
    const s = 8; // size * 2
    const theta = (tri.rotation! * Math.PI) / 180;
    const cx = tri.box.left + s / 2;
    const cy = tri.box.top + s / 2;
    const tipX = cx + (s / 2) * Math.sin(theta);
    const tipY = cy - (s / 2) * Math.cos(theta);
    expect(tipX).toBeCloseTo(70, 4); // 60 + 10
    expect(tipY).toBeCloseTo(100, 4); // 90 + 10
  });

  it("renders an annular wedge (sunburst ring / gauge) as a rotated rectangle band", async () => {
    const slide = await insert([
      {
        kind: "wedge",
        cx: 50,
        cy: 50,
        r: 30,
        innerR: 15,
        startAngle: 0,
        endAngle: 90,
        fill: "#333333",
        stroke: "#ffffff",
        strokeWidth: 1,
        name: "ring",
      },
    ]);
    const band = slide.created.filter((s) => s.geo === "rectangle" && s.name?.includes("-f"));
    expect(band.length).toBeGreaterThan(2); // the annular band, not a triangle fan
    for (const b of band) {
      expect(b.fillColor).toBe("#333333");
      expect(typeof b.rotation).toBe("number");
    }
    // No triangles for an annular wedge (a triangle can't leave a hole).
    expect(slide.created.some((s) => s.geo === "triangle")).toBe(false);
    // Two radial separators in the stroke colour.
    const edges = slide.created.filter((s) => s.name === "ring-edge");
    expect(edges.length).toBe(2);
    for (const e of edges) expect(e.fillColor).toBe("#ffffff");
  });

  it("skips grouping when group:false or only one shape", async () => {
    const slide = await insert(
      [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#111111" },
        { kind: "rect", x: 20, y: 0, w: 10, h: 10, fill: "#222222" },
      ],
      { group: false, tagData: "cfg" },
    );
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    // The tag falls back onto a plain shape rather than a group.
    expect(taggedShape(slide)!.tagStore.get(CHART_TAG)).toBe("cfg");
  });

  // Shape.rotation is PowerPointApi 1.10 and the manifests admit hosts from 1.4.
  // A try/catch around the assignment catches NOTHING on a real host: Office.js
  // proxy setters do not throw synchronously — the host rejects the queued
  // command at the next context.sync(), which carries the whole batch. So the
  // fake here rejects at SYNC, the way PowerPoint does, not at the setter.
  it.each(["pie", "doughnut", "sunburst"])(
    "%s still inserts on a pre-1.10 host (rotation gated, not wrapped)",
    async (kind) => {
      const slide = makeSlide("s-old");
      const ctx = installHost([slide], [], slide, (v) => v !== "1.10");
      const realSync = ctx.sync;
      ctx.sync = async <T>(passThroughValue?: T) => {
        // Any rotation assigned on a host without 1.10 poisons the whole batch.
        if (slide.created.some((sh) => sh.rotation !== undefined)) {
          throw new Error("PropertyNotSupported: Shape.rotation requires PowerPointApi 1.10");
        }
        // Forwarded, like the real thing — see the `sync` stub in
        // `test/helpers/office-host.ts` for what dropping it cost.
        return realSync(passThroughValue);
      };
      const scene = buildChart(sampleConfig(kind as never));
      // Must not reject: the chart degrades (no wedges) instead of failing.
      let failure: unknown = null;
      try {
        await insertSceneIntoSlide(scene, { tagData: "{}" });
      } catch (e) {
        failure = e;
      }
      expect(failure, `insert rejected on a pre-1.10 host: ${String(failure)}`).toBeNull();
      // The rest of the chart — labels, leader lines — still lands on the slide.
      expect(slide.created.length).toBeGreaterThan(0);
      // …and nothing carries a rotation the host cannot accept.
      expect(slide.created.every((sh) => sh.rotation === undefined)).toBe(true);
    },
  );

  it("degrades gracefully when the host lacks grouping and rotation", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    // Break addGroup and rotation assignment the way an old host would.
    slide.shapes.addGroup = () => {
      throw new Error("addGroup requires PowerPointApi 1.8");
    };
    const scene = {
      width: 200,
      height: 100,
      nodes: [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#111111" },
        { kind: "wedge", cx: 50, cy: 50, r: 30, innerR: 0, startAngle: 0, endAngle: 90, fill: "#333333", name: "w" },
      ],
    };
    const realAdd = slide.shapes.addGeometricShape.bind(slide.shapes);
    slide.shapes.addGeometricShape = (geo, box) => {
      const s = realAdd(geo, box);
      if (geo === "triangle") {
        Object.defineProperty(s, "rotation", {
          set() {
            throw new Error("rotation requires PowerPointApi 1.10");
          },
        });
      }
      return s;
    };
    await insertSceneIntoSlide(scene as never, { tagData: "cfg" });
    // No group, no fan triangles survive — but the rect is inserted and tagged.
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(taggedShape(slide)!.tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("still inserts (ungrouped) when the host lacks grouping — the web case", async () => {
    // PowerPoint on the web: grouping (1.8) unsupported, tags (1.3) supported.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.8");
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    // The shapes are committed and no grouping was attempted…
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(slide.created.filter((s) => s.geo === "rectangle").length).toBeGreaterThanOrEqual(4);
    // …and the config tag lands on one of them, so the chart is re-editable.
    expect(taggedShape(slide)!.tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("skips tagging when the host lacks tags", async () => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => false); // nothing supported
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(taggedShape(slide), "a host without tags must leave every shape untagged").toBeUndefined();
    expect(slide.created.length).toBeGreaterThan(0);
  });

  it("falls back to the first slide when nothing is selected", async () => {
    const slide = makeSlide("s1");
    const ctx = installHost([slide]);
    ctx.presentation.getSelectedSlides = () => {
      throw new Error("no selection");
    };
    await insertSceneIntoSlide(
      { width: 10, height: 10, nodes: [{ kind: "rect", x: 0, y: 0, w: 5, h: 5, fill: "#111111" }] } as never,
      {},
    );
    expect(slide.created).toHaveLength(1);
  });
});

describe("looking away while a chart redraws", () => {
  /** Wire a controllable selection onto the fake host. */
  function withSelection(ctx: ReturnType<typeof installHost>, initial: string[]) {
    const state = { selected: [...initial], sets: [] as string[][] };
    const p = ctx.presentation as unknown as Record<string, unknown>;
    p.getSelectedSlides = () => ({ items: state.selected.map((id) => ({ id })), load() {} });
    p.setSelectedSlides = (ids: string[]) => {
      state.sets.push([...ids]);
      state.selected = [...ids];
    };
    return state;
  }

  it("puts the user back where they were", async () => {
    const ctx = installHost([makeSlide("s1"), makeSlide("s2")]);
    const sel = withSelection(ctx, ["s1"]);
    const saw = await withSlideDeselected(["s1"], async (deselected) => deselected);
    expect(saw).toBe(true);
    // Parked elsewhere for the redraw, then restored.
    expect(sel.sets).toEqual([["s2"], ["s1"]]);
    expect(sel.selected).toEqual(["s1"]);
  });

  it("leaves the selection alone when the user moved during the redraw", async () => {
    // An off-screen redraw runs for tens of seconds and the user is free to
    // click through the deck while it does. Restoring unconditionally snapped
    // them back to wherever they happened to be standing when it started,
    // throwing away their navigation with no notice.
    const ctx = installHost([makeSlide("s1"), makeSlide("s2"), makeSlide("s3")]);
    const sel = withSelection(ctx, ["s1"]);
    await withSlideDeselected(["s1"], async () => {
      sel.selected = ["s3"]; // the user clicks slide 3 mid-redraw
    });
    expect(sel.selected).toEqual(["s3"]);
    // Only the park was written; nothing restored over the user's own move.
    expect(sel.sets).toEqual([["s2"]]);
  });

  it("makes somewhere to look when every slide is one it is about to draw on", async () => {
    // The one-slide deck. It used to run `fn(false)` — the live canvas, batch
    // 10 — which is the exact configuration a real run died in, and it is the
    // deck a user building their first chart actually has. So make a slide to
    // look at, and take it away again.
    const deck = [makeSlide("s1")];
    const ctx = installHost(deck);
    const sel = withSelection(ctx, ["s1"]);
    const saw = await withSlideDeselected(["s1"], async (deselected) => {
      // MID-redraw: the scratch slide exists and the view is on it, which is
      // the whole point — asserting only the end state would pass against a
      // function that added and removed a slide without ever looking at it.
      expect(deck).toHaveLength(2);
      expect(sel.selected).toEqual([deck[1].id]);
      return deselected;
    });
    // True: the caller may legitimately spend the off-screen batch budget.
    expect(saw).toBe(true);
    // And the deck is exactly as the user left it.
    expect(deck).toHaveLength(1);
    expect(deck[0].id).toBe("s1");
    expect(sel.selected).toEqual(["s1"]);
  });

  it("draws on the live canvas when the scratch slide will not land", async () => {
    // The host swallows slides.add() under load — the behaviour addSlides
    // exists to survive. A scratch slide that never landed must not be
    // reported as parked: `fn` would then use the off-screen batch size on a
    // slide the user is still looking at, which is worse than not trying.
    const deck = [makeSlide("s1")];
    const ctx = installHost(deck);
    const sel = withSelection(ctx, ["s1"]);
    faults.swallowAdds = 1;
    try {
      const saw = await withSlideDeselected(["s1"], async (deselected) => deselected);
      expect(saw).toBe(false);
      expect(deck).toHaveLength(1);
      // Nothing was selected or restored — there was nowhere to go.
      expect(sel.sets).toEqual([]);
    } finally {
      faults.swallowAdds = 0;
    }
  });

  /**
   * The failure no `catch` can see: a sync that neither resolves nor rejects.
   *
   * `withSlideDeselected` was two raw `PowerPoint.run`s sandwiching draw work
   * that IS bounded per batch — so the one part of an in-place update that could
   * hang forever was the part before anything was drawn. And the second one sits
   * in a `finally`, so a sync that never settles there stops the `finally` from
   * completing, which means the CALLER's `finally` never runs either: the pane's
   * busy counter stays up for the rest of the session, the selection banner goes
   * dead, the status strip freezes, and the auto-update timer re-arms forever.
   * Stop cannot break in — it is checked at batch boundaries this never reaches.
   *
   * Raced against a timer rather than simply awaited, so the guard fails on its
   * own assertion instead of on a suite timeout.
   */
  const returnedWithin = async <T>(ms: number, work: Promise<T>): Promise<T | "never came back"> =>
    Promise.race([work, new Promise<"never came back">((r) => setTimeout(() => r("never came back"), ms))]);

  it("comes back from a redraw even when the host stops answering selection calls", async () => {
    // The selection API has to be wired, or the park throws before it ever
    // syncs and the wedge is never reached — a version of this that skipped
    // `withSelection` passed against an unbounded park, which is the whole
    // thing it was written to catch.
    const ctx = installHost([makeSlide("s1"), makeSlide("s2")]);
    withSelection(ctx, ["s1"]);
    _setSelectionTimeoutForTest(20);
    // From the first sync on, the host answers nothing at all — ever.
    faults.wedgeAfterSyncs = 0;
    try {
      const saw = await returnedWithin(
        200,
        withSlideDeselected(["s1"], async (deselected) => deselected),
      );
      expect(saw, "an in-place update never returned on a host that went quiet").not.toBe("never came back");
      // And it degraded honestly: nothing was parked, so the caller must NOT be
      // told it may use the off-screen batch size on a live canvas.
      expect(saw).toBe(false);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setSelectionTimeoutForTest(4_000);
    }
  });

  it("comes back from the RESTORE too, which sits in a finally", async () => {
    // The second round trip is the worse one. A sync that never settles inside
    // a `finally` stops that `finally` from completing, so the caller's own
    // `finally` never runs either — and the pane decrements its busy counter in
    // one of those. The view not being restored is cheap; the update never
    // returning is what kills the session.
    const ctx = installHost([makeSlide("s1"), makeSlide("s2")]);
    withSelection(ctx, ["s1"]);
    _setSelectionTimeoutForTest(20);
    try {
      const saw = await returnedWithin(
        200,
        withSlideDeselected(["s1"], async (deselected) => {
          // The park has happened; go quiet from here, so only the restore hits it.
          faults.wedgeAfterSyncs = 0;
          return deselected;
        }),
      );
      expect(saw, "the redraw never returned because the restore hung").not.toBe("never came back");
      expect(saw, "did not park at all").toBe(true);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setSelectionTimeoutForTest(4_000);
    }
  });

  /**
   * The draw was bounded per batch and everything AFTER the last batch was not.
   *
   * `groupAndTagAll`'s five syncs and `ungroupedFallback`'s bare one all went
   * through `step()`, which labels and adds no deadline — so a host that went
   * quiet once the shapes were on the slide left the insert with no timer, no
   * `gave up waiting`, no phase note past "grouping…", and no way for Stop to
   * break in (it is checked at batch boundaries this never reaches). That is the
   * 1819-second wedge shape, one phase further along than the one that got
   * bounded. `ungroupedFallback` matters most: grouping is refused on the web,
   * so every web-host insert goes down it.
   */
  it("comes back from grouping and tagging when the host goes quiet after the draw", async () => {
    installHost([makeSlide("s1")]);
    _setBatchTimeoutForTest(20);
    // Past the context and the first draw batches; the shapes are committed and
    // the host stops answering from here on.
    faults.wedgeAfterSyncs = 3;
    try {
      const got = await returnedWithin(400, insertSceneIntoSlide(buildChart(config), { tagData: "{}" }));
      expect(got, "the insert never returned once the host went quiet after drawing").not.toBe("never came back");
    } finally {
      faults.wedgeAfterSyncs = null;
      _setBatchTimeoutForTest(45_000);
    }
  });

  it("comes back from the Insert click's reads when the host stops answering", async () => {
    // `getSelectionBounds` was the only selection read in the file not on
    // `boundedRun`, and it is the FIRST host call the Insert button makes —
    // so a quiet host took the whole insert with it: buttons disabled,
    // "Working…" counting up, nothing drawn and nothing said. `guard()` has no
    // deadline on the action either, so there was nothing else to stop it.
    installHost([makeSlide("s1")]);
    _setSelectionTimeoutForTest(20);
    faults.wedgeAfterSyncs = 0;
    try {
      expect(await returnedWithin(200, getSelectionBounds()), "getSelectionBounds never returned").toBe(null);
      expect(await returnedWithin(200, getSlideShapeBounds()), "getSlideShapeBounds never returned").toBe(null);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setSelectionTimeoutForTest(4_000);
    }
  });

  /**
   * A slide the host will not hand back by id is not necessarily gone, and
   * treating it as gone is how an add-in litters someone's deck.
   *
   * PowerPoint on the web resolved a freshly-added slide's id once and then
   * refused it — while still listing that same id in `slides.load("items/id")`.
   * `deleteSlideById` read the refusal as "already gone, nothing to do" and
   * reported success, so a host-probe run left fourteen blank slides behind and
   * said it had cleaned up. The same call cleans up after every off-screen
   * redraw, on the user's own deck.
   */
  it("takes out a slide the host will not resolve by id", async () => {
    const deck = [makeSlide("s1")];
    installHost(deck);
    const id = await addScratchSlide();
    expect(id, "the scratch slide did not land").toBeTruthy();
    expect(deck).toHaveLength(2);
    // From here on the host denies that slide exists whenever it is asked for
    // by id — while still listing it among the deck's slides.
    faults.newSlideResolvesTimes = 0;
    try {
      expect(await deleteSlideById(id!), "reported a clean-up it had not done").toBe(true);
      expect(
        deck.map((s) => s.id),
        "the slide is still in the deck",
      ).toEqual(["s1"]);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("deletes nothing when the id is not in the deck at all, and does not claim it did", async () => {
    // The other half of the same question. A positional delete driven by an id
    // nobody can find is how an add-in destroys work, so an id the deck does
    // not list must end the search — nothing to remove, and nothing removed.
    // That half is unchanged and is the one that protects the user's deck.
    //
    // What CHANGED is the verdict it reports. "Not in the deck's list" used to
    // return true, i.e. "already gone", and 2026-08-11 (`756682e`) measured
    // that reading false on this host: of the 62 scratch slides a probe run was
    // deleting, `the deck still lists 0 of 62 of these ids` — zero — while the
    // deck stayed at 65 slides and the run reported a clean sweep. Both id
    // lists come from the same `slideIds()` projection minutes apart, so "the
    // id is not there" cannot mean the slide is not there.
    //
    // Unfindable is UNKNOWN. For an id that genuinely never existed this is now
    // a shade pessimistic, and that is the right way to be wrong: an
    // under-count costs a line in a report, while an over-count leaves sixty
    // blank slides in someone's deck and says it left none. The caller's deck
    // count (`slidesActuallyReturned`) is what turns this honest false back
    // into an honest number.
    const deck = [makeSlide("s1"), makeSlide("s2")];
    installHost(deck);
    expect(await deleteSlideById("no-such-slide"), "an id nobody can find was reported as confirmed gone").toBe(false);
    expect(
      deck.map((s) => s.id),
      "a delete driven by an unfindable id touched the deck",
    ).toEqual(["s1", "s2"]);
  });

  /**
   * The slide-swap gate authorises DELETING the user's slide. It has to be sure.
   *
   * `slideHoldsOnlyChart` read `slide.shapes.items` and answered yes for an
   * empty list — and it is consulted only AFTER the host has already stalled,
   * which is exactly the state in which this host answers shape collections
   * short (`shapesExpected=19 shapesSeen=15`). A hollow read therefore looked
   * identical to a bare slide, and the swap replaced a slide holding the user's
   * logo, title and footnote with a generated one carrying none of it and no
   * speaker notes either. Recoverable only by Ctrl-Z, and the note the user
   * gets does not mention shapes.
   */
  it("refuses the slide swap when it cannot corroborate what is on the slide", async () => {
    const slide = makeSlide("s1");
    // Two real shapes the user put there...
    slide.created.push(makeShape("geometric", "rectangle", { left: 0, top: 0, width: 10, height: 10 }));
    slide.created.push(makeShape("geometric", "rectangle", { left: 20, top: 0, width: 10, height: 10 }));
    installHost([slide]);
    // ...and a host that answers the collection empty while still counting two.
    faults.hollowNameReads = 1;
    try {
      expect(await slideHoldsOnlyChart("s1"), "believed a hollow read and offered the swap").toBe(false);
    } finally {
      faults.hollowNameReads = 0;
    }
  });

  it("still allows the swap on a slide it CAN corroborate as empty", async () => {
    // The gate has to stay usable, or the fallback it guards is dead code. A
    // bare slide this run never drew on, honestly read, is still a yes.
    //
    // A slide id nothing else in this file uses, for the reason the counter
    // tests at the bottom give: `shapesDrawnOnSlide` is a per-RUN total and is
    // not reset between tests, so a shared `s1` carries every earlier draw into
    // this one — and this gate now reads that counter. It passed under `s1`
    // only while unnamed draws banked under the `(visible)` sentinel, which is
    // the bug `slideKeyFor` fixes.
    installHost([makeSlide("swap-bare")]);
    expect(await slideHoldsOnlyChart("swap-bare")).toBe(true);
  });

  /**
   * The case the corroboration above cannot see: both signals agreeing at ZERO.
   *
   * The test before this one arms `hollowNameReads`, where the collection reads
   * empty and `getCount()` still says two — a disagreement, which
   * `slideShapeNames` catches. Round `957aca0` produced the other shape: the
   * collection AND the count both answered zero for a slide holding a shape
   * named `PowerChart`. Nothing in the answer distinguishes that from a bare
   * slide, so the gate has to decide on something that is not the answer.
   *
   * `shapesDrawnOn` is that something, and it splits along the pathology's own
   * line — this host does not list the shapes a run just added. A slide this
   * run drew on cannot credibly read empty, and this gate deletes the slide it
   * says yes to.
   */
  it("refuses the swap when a slide reads empty AFTER this run drew on it", async () => {
    // A slide id nothing else in this file uses: `shapesDrawnOnSlide` is a
    // per-RUN total and is not reset between tests, so a shared `s1` would
    // carry every earlier draw into this one. Same reason as the counter tests
    // at the bottom of this file.
    installHost([makeSlide("swap-drawn-on")]);
    // Draw a chart onto the slide, so the run's own bookkeeping says shapes
    // went there — via the ordinary insert path, not by reaching into the
    // counter, or the test would be asserting against its own fixture.
    const scene = buildChart(sampleConfig("clustered"));
    await insertSceneIntoSlide(scene, { slideId: "swap-drawn-on" });
    expect(
      shapesDrawnOn("swap-drawn-on"),
      "the insert drew nothing, so the premise of this test is missing",
    ).toBeGreaterThan(0);
    // Now the host stops listing that slide — both signals, both zero.
    faults.slideReadsEmpty = "now";
    try {
      expect(
        await slideHoldsOnlyChart("swap-drawn-on"),
        "believed an empty read on a slide this run had drawn on, and offered to delete it",
      ).toBe(false);
    } finally {
      faults.slideReadsEmpty = null;
    }
  });

  /**
   * …and the swap comes BACK once the run has taken its shapes off again.
   *
   * The refusal above must not cost the case the empty branch exists for: a
   * redraw that stalled, whose litter the caller sweeps immediately before
   * asking this question, on a slide that held nothing but the chart. That run
   * has a net contribution of zero, there are no run-added shapes for the host
   * to be blind about, and its empty read is an answer again.
   *
   * Both halves come from one subtraction — `deleteShapesById` gives the count
   * back what it sweeps — so this is the guard that would catch the decrement
   * being dropped, which the test above cannot see.
   */
  it("allows the swap again once the run has swept the shapes it drew", async () => {
    const slide = makeSlide("swap-swept");
    installHost([slide]);
    const scene = buildChart(sampleConfig("clustered"));
    const target = await insertSceneIntoSlide(scene, { slideId: "swap-swept" });
    const ids = slide.created.filter((s) => !s.deleted).map((s) => s.id);
    expect(target, "the insert produced no target").toBeTruthy();
    expect(ids.length, "nothing was drawn, so the sweep below would prove nothing").toBeGreaterThan(0);
    const swept = await deleteShapesById("swap-swept", ids);
    expect(swept, "the sweep removed nothing").toBe(ids.length);
    expect(shapesDrawnOn("swap-swept"), "the sweep did not give the slide's count back").toBe(0);
    // The host is still refusing to list the slide — and now that is the truth.
    faults.slideReadsEmpty = "now";
    try {
      expect(
        await slideHoldsOnlyChart("swap-swept"),
        "refused a swept slide, which is the case the fallback is for",
      ).toBe(true);
    } finally {
      faults.slideReadsEmpty = null;
    }
  });

  /**
   * A refused lookup is not "the original slide is gone".
   *
   * The swap used to open its own context, ask for the target by id, and throw
   * `original slide is gone` the moment the null flag was not `false` — on the
   * very id `insertSlidesFromBase64` had just accepted as `targetSlideId`, with
   * the count check already proving the deck had grown by one. So the original
   * was demonstrably still there, and the user was told to sort out two
   * identical slides by hand, permanently. This host does exactly that: it
   * resolves a slide id once and refuses it after, while still listing it.
   */
  it("removes the slide it replaced even when the host will not resolve its id", async () => {
    const built = await buildDeckBase64(
      [{ scene: buildChart(sampleConfig("clustered")), title: "A", configJson: "{}", slot: 0, run: "r1" }],
      { width: 720, height: 405 },
    );
    const deck = [makeSlide("s1"), makeSlide("s2")];
    installHost(deck);
    // The host takes the null check for s1 and answers nothing — neither "gone"
    // nor "there", which is the state `isLive` is deliberately pessimistic about.
    unansweredNullChecks.add("s1");
    try {
      expect(await replaceSlideWithDeck("s1", built.base64)).toBe("swapped");
      expect(
        deck.map((s) => s.id),
        "left the replaced slide in the deck",
      ).not.toContain("s1");
    } finally {
      unansweredNullChecks.clear();
    }
  });

  /**
   * A deck this host took, counted before it caught up, reported as a failure.
   *
   * The swap decides on `slideCount() !== before + 1`, and the raw read is the
   * one `insertSlidesFromPptx` was returning `landed: 0` from — same
   * `insertSlidesFromBase64` call, same host, four times in the archive (rounds
   * 148, 315, 375, 422). Here it is worse than a wrong number: the return is
   * BEFORE the delete, so a user told "failed" is left with the new slide and
   * the original both on the deck.
   *
   * A ceiling of two, one bind. Two is the deck's size when `before` is read,
   * so that read is accurate — the ceiling only caps a deck that has grown past
   * it — and the read after the insert is pinned back to exactly `before`,
   * which is the reading that produces the false "failed".
   */
  it("does not call a swap failed because the deck had not caught up", async () => {
    const built = await buildDeckBase64(
      [{ scene: buildChart(sampleConfig("clustered")), title: "A", configJson: "{}", slot: 0, run: "r1" }],
      { width: 720, height: 405 },
    );
    const deck = [makeSlide("s1"), makeSlide("s2")];
    installHost(deck);
    faults.slideCountCeiling = 2;
    faults.slideCountCeilingBinds = 1;
    _setCountSettleDelayForTest(20);
    try {
      expect(await replaceSlideWithDeck("s1", built.base64)).toBe("swapped");
      // The allowance is spent by the raw read too, so this proves the lag
      // happened whichever version is running — it cannot be the assertion that
      // does the work.
      expect(faults.slideCountCeilingBinds, "the ceiling never bound, so no lag was reproduced").toBe(0);
    } finally {
      faults.slideCountCeiling = null;
      faults.slideCountCeilingBinds = 0;
      _setCountSettleDelayForTest(4_000);
    }
  });

  /**
   * A stray the host will not resolve has not been swept, and saying so matters
   * to three separate callers.
   *
   * Every id handed to this sweep names a shape that COMMITTED, so it is on the
   * slide by construction. The old code read `s.isNullObject` raw — which on an
   * unanswered proxy throws `PropertyNotLoaded`, took the whole sweep down, and
   * left the catch reporting 0 for the addressable strays as well. The picture
   * fallback then drew over the debris, the slide-swap gate saw a slide that was
   * not clean, and Stop's promise that nothing is left behind rested on a count
   * that could not tell "nothing to do" from "could not touch any of it".
   */
  it("sweeps the strays it can reach even when one of them will not answer", async () => {
    /**
     * THE QUIET ONE STAYS, and this assertion was briefly inverted and then put
     * back — 2026-09-08. A collection-read fallback was added on the reasoning
     * that a shape listed in its own slide's collection is positively
     * confirmed, which is true and is not the problem. Round 431 measured the
     * fallback in production: seven fires, zero recovered, while the collection
     * listed 8 to 31 shapes. This host lists a freshly drawn shape under an id
     * that is not the one it returned at creation, so nothing we hold matches.
     * The fallback is gone; `powerpoint.ts` carries the numbers.
     */
    const slide = makeSlide("s1");
    installHost([slide]);
    const a = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 5, height: 5 });
    const b = slide.shapes.addGeometricShape("rectangle", { left: 6, top: 0, width: 5, height: 5 });
    unansweredNullChecks.add(a.id); // this one the host will not describe
    try {
      const swept = await deleteShapesById("s1", [a.id, b.id]);
      expect(swept, "one quiet stray took the whole sweep down").toBe(1);
      expect(b.deleted, "the reachable stray was left on the slide").toBe(true);
      expect(a.deleted, "deleted a shape the host would not confirm").toBe(false);
    } finally {
      unansweredNullChecks.clear();
    }
  });

  it("records that a shape resolved by id through a slide handle a sync old, not only that one did not", async () => {
    // THE ONLY WITNESS LEFT. This sweep is the last place in production that
    // resolves a shape by id through a slide handle resolved a SYNC AGO, and the
    // probe that asks whether this host permits that —
    // `shape-resolve-held-slide-proxy` — has answered `no-scratch-shape` in all
    // 133 archived rounds and structurally cannot do better: it needs an id for
    // a freshly added shape, which this host will not give.
    //
    // Until 2026-08-22 the sweep testified only AGAINST: a line when the host
    // would not resolve, and nothing at all when it did. A healthy round sweeps
    // no wreckage and produces exactly the same silence, so the archive held 133
    // rounds of nothing and no way to tell the two apart.
    const slide = makeSlide("s1");
    installHost([slide]);
    const a = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 5, height: 5 });
    setTracing(true);
    try {
      expect(await deleteShapesById("s1", [a.id])).toBe(1);
      const line = traceLog().entries.find((e) =>
        /resolved a shape by id through a slide handle a sync old/.test(e.message),
      );
      expect(line, "the sweep resolved a shape through an aged handle and said nothing about it").toBeTruthy();
      expect((line as unknown as { data?: { resolved?: number } }).data?.resolved).toBe(1);
    } finally {
      setTracing(false);
    }
  });

  it("takes back a scratch slide whose id it cannot verify", async () => {
    // `addScratchSlide` refuses to hand out an id it could not resolve — but
    // the slide landed, so refusing without also removing it would leave a
    // blank slide in the deck on every attempt.
    const deck = [makeSlide("s1")];
    installHost(deck);
    // The fake names an added slide after the deck's length; the null-check on
    // it goes unanswered, which is the state a caller can least reason about.
    unansweredNullChecks.add("slide-2");
    try {
      expect(await addScratchSlide(), "handed out an id it could not resolve").toBeNull();
      expect(
        deck.map((s) => s.id),
        "left the unusable scratch slide behind",
      ).toEqual(["s1"]);
    } finally {
      unansweredNullChecks.clear();
    }
  });

  it("removes the scratch slide even when the user navigates away mid-redraw", async () => {
    // The restore is deliberately skipped when the user has moved themselves.
    // The CLEANUP is not conditional on it: a blank slide the add-in left at
    // the end of someone's deck is litter whether or not they are looking at
    // it, and tying the delete to the restore would leak one every time.
    const deck = [makeSlide("s1")];
    const ctx = installHost(deck);
    const sel = withSelection(ctx, ["s1"]);
    await withSlideDeselected(["s1"], async () => {
      sel.selected = ["s1"]; // the user clicks back to their own slide mid-redraw
    });
    expect(deck).toHaveLength(1);
    expect(sel.selected).toEqual(["s1"]);
  });
});

describe("updateChartInSlide", () => {
  it("deletes the old group and re-renders at the same position", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "x" });
    const oldGroup = slide.created.find((s) => s.type === "group")!;
    const before = slide.created.length;
    await updateChartInSlide(buildChart(config), { slideId: "s1", shapeId: oldGroup.id, left: 33, top: 44 });
    expect(oldGroup.deleted).toBe(true);
    const fresh = slide.created.slice(before).filter((s) => s.type !== "group");
    expect(fresh.length).toBeGreaterThan(0);
    expect(Math.min(...fresh.map((s) => s.box.left))).toBeGreaterThanOrEqual(33);
  });

  it("deletes the WHOLE chart, not just its tagged shape, when it is ungrouped", async () => {
    // PowerPoint on the web: no grouping, so the config tag can only sit on ONE
    // of the chart's shapes. The update deleted exactly that shape and redrew
    // the chart, leaving the other twelve underneath — 13 shapes became 25, then
    // 37, on successive edits, as stacked misaligned duplicates.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.8");
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg" });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    const drawn = slide.created.length;
    expect(drawn).toBeGreaterThan(1);
    // The tagged shape carries the rest of the chart with it.
    expect(JSON.parse(taggedShape(slide)!.tagStore.get(CHART_PARTS_TAG)!)).toHaveLength(drawn - 1);

    const live = () => slide.created.filter((s) => !s.deleted);
    for (const edit of [1, 2]) {
      // Same Scale: read the deck back, re-render every chart it finds.
      const found = (await listChartsInDeck()).charts.filter((c) => live().some((s) => s.id === c.target.shapeId));
      expect(found, `edit ${edit}`).toHaveLength(1);
      await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: "cfg" } }]);
      expect(live(), `edit ${edit}`).toHaveLength(drawn);
    }
  });

  it("redraws a chart whose id the host refused, instead of treating it as deleted", async () => {
    // THE ARCHIVE'S OLDEST LIVE FAULT. 46 of the 47 recorded `explode a
    // degraded picture` failures carry `idRefusals > 0`, and the two that could
    // report a verdict at all both said the same thing: "the host would not
    // work on the chart again, but it is STILL ON THE SLIDE — nothing was
    // lost". A by-id lookup poisons the sync it was queued in, the whole
    // resolve died, and the caller got a null target — a silent no-op on a
    // chart the user is looking at.
    //
    // The recovery is the asymmetry the rest of this file already leans on: the
    // host refuses `shapes.getItem(id)` and answers `shapes.load("items/id")`.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group")!;
    const before = slide.created.length;
    faults.refuseShapeById = true;
    try {
      const next = await updateChartInSlide(buildChart(config), {
        slideId: "s1",
        shapeId: group.id,
        left: 33,
        top: 44,
      });
      expect(next, "the update gave up on a chart that is still on the slide").toBeTruthy();
      expect(group.deleted, "the old chart was left behind under the redraw").toBe(true);
      expect(slide.created.slice(before).length, "nothing was redrawn").toBeGreaterThan(0);
    } finally {
      faults.refuseShapeById = false;
    }
  });

  it("still does not resurrect a deleted chart when the host is refusing ids", async () => {
    // THE GUARD THE RECOVERY ABOVE MUST NOT COST. A re-read that finds the
    // chart is the whole point; a re-read that resurrects one the user deleted
    // would be strictly worse than the silent no-op it replaces, because an
    // in-place update that inserts is not an update.
    //
    // Safe by construction rather than by care: a deleted shape is absent from
    // the slide's collection too, so the re-read cannot find it. This is what
    // holds that claim to it.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group")!;
    for (const s of slide.created) s.delete();
    faults.refuseShapeById = true;
    try {
      await updateChartInSlide(buildChart(config), { slideId: "s1", shapeId: group.id, left: 10, top: 20 });
      expect(
        slide.created.filter((s) => !s.deleted),
        "the update resurrected a chart the user deleted",
      ).toHaveLength(0);
    } finally {
      faults.refuseShapeById = false;
    }
  });

  /**
   * The instrument for the last live `GetItem(id)` refusal.
   *
   * `reading back an ungrouped chart's shape ids` fires in 56 of 57 archived
   * rounds, and each failure costs a chart its parts list — 9 to 12 per round in
   * the six newest. The consequence written beside that read is that the update
   * deletes the one shape it can name and redraws all of them, so the chart
   * grows by a whole chart. Nobody has ever seen it happen, and no round records
   * whether one of those charts was ever the one an update touched.
   */
  describe("counting what an update leaves behind", () => {
    const orphanLine = () => traceLog().entries.find((e) => /shapes left on the slide/.test(e.message));

    it("reports no orphans when the chart's parts list is intact", async () => {
      // The control, and it has to come first: an instrument that cries orphan
      // on a healthy update would say nothing about the sick one.
      const slide = makeSlide("s1");
      installHost([slide], [], slide, (v) => v !== "1.8");
      const scene = buildChart(config);
      await insertSceneIntoSlide(scene, { tagData: "cfg" });
      expect(
        slide.created.some((s) => s.type === "group"),
        "this test needs an UNGROUPED chart",
      ).toBe(false);
      const found = (await listChartsInDeck()).charts;
      expect(found).toHaveLength(1);
      expect(found[0].target.partIds?.length, "the parts list is the thing being controlled for").toBeGreaterThan(0);
      setTracing(true);
      try {
        await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: "cfg" } }]);
        const line = orphanLine();
        expect(line, "the instrument did not run").toBeTruthy();
        const d = line?.data as { growth: number; withParts: number; atRisk: number };
        expect(d.withParts).toBe(1);
        expect(d.growth, "a faithful replacement leaves the slide the size it found it").toBe(0);
        // Ungrouped, but its parts list is intact, so the update can name every
        // shape it owns — not exposure either.
        expect(d.atRisk, "a chart that can name its own parts is not at risk").toBe(0);
      } finally {
        setTracing(false);
      }
    });

    it("does NOT count a chart that drew one shape — it can strand nothing", async () => {
      // EVERY NON-ZERO `atRisk` IN THE ARCHIVE WAS THIS. All nine across 69
      // rounds came from `explode a degraded picture`, always `atRisk=1,
      // charts=1`: its second update replaces `pictured`, the single picture
      // shape the collapse just made. One shape, deleted by its own id, nothing
      // behind it. Not a group and with no parts list, so it scored — and
      // `docs/ROUNDS.md` read those nine as three real exposures sampled clean.
      //
      // A one-shape chart is the control this counter never had.
      const slide = makeSlide("s1");
      installHost([slide], [], slide, (v) => v !== "1.8");
      // A scene with literally one node. `buildChart` never makes one — the
      // smallest config still renders five shapes, which the assertion below
      // caught on the first draft of this test.
      const full = buildChart(config);
      const oneShape = { ...full, nodes: full.nodes.slice(0, 1) };
      await insertSceneIntoSlide(oneShape, { tagData: "cfg" });
      const found = (await listChartsInDeck()).charts;
      const blind = { ...found[0].target, partIds: undefined };
      setTracing(true);
      try {
        await updateChartsInSlides([{ scene: oneShape, target: blind, opts: { tagData: "cfg" } }]);
        const d = orphanLine()?.data as { atRisk: number; charts: number };
        // The scene must really be a single shape, or this test proves nothing
        // about the guard and everything about the fixture.
        expect(estimateOfficeShapes(oneShape), "fixture is not a one-shape chart").toBe(1);
        expect(d?.atRisk, "a chart that drew one shape cannot strand a second").toBe(0);
      } finally {
        setTracing(false);
      }
    });

    it("counts the shapes stranded when the chart has no parts list", async () => {
      // THE READING THIS EXISTS FOR. A target with no `partIds` is exactly what
      // a refused `reading back an ungrouped chart's shape ids` produces: the
      // chart is whole on the slide and the update can name only its anchor.
      const slide = makeSlide("s1");
      installHost([slide], [], slide, (v) => v !== "1.8");
      const scene = buildChart(config);
      await insertSceneIntoSlide(scene, { tagData: "cfg" });
      const drawn = slide.created.filter((s) => !s.deleted).length;
      expect(drawn).toBeGreaterThan(1);
      const found = (await listChartsInDeck()).charts;
      // Strip it — the shapes stay exactly where they are, which is the point.
      const blind = { ...found[0].target, partIds: undefined };
      setTracing(true);
      try {
        await updateChartsInSlides([{ scene, target: blind, opts: { tagData: "cfg" } }]);
        const d = orphanLine()?.data as { growth: number; withParts: number; removedCalls: number; atRisk: number };
        expect(d, "the instrument did not run").toBeTruthy();
        expect(d.withParts, "no chart in this update had a parts list").toBe(0);
        expect(d.removedCalls, "only the anchor could be named").toBe(1);
        // The number is the finding: every shape but the anchor is still there,
        // underneath a full redraw.
        expect(d.growth, "the growth this instrument exists to measure").toBe(drawn - 1);
        // Ungrouped AND with no parts list — the one population that can strand.
        expect(d.atRisk, "this is exactly the chart the question is about").toBe(1);
        // And it is real, not just arithmetic — the slide really did grow.
        expect(slide.created.filter((s) => !s.deleted).length).toBe(drawn - 1 + drawn);
      } finally {
        setTracing(false);
      }
    });

    it("reports no growth for a GROUPED chart, whatever it holds inside", async () => {
      // THE CASE THAT LET A UNIT BUG THROUGH TWO DRAFTS AND A REAL ROUND. Every
      // other test here uses the ungrouped host, where each of a chart's shapes
      // is top-level and "shapes inside the chart" happens to equal "shapes on
      // the slide" — so subtracting one from the other looked right.
      //
      // A group is ONE top-level shape however many it contains. Deleting it is
      // one call, redrawing it is one shape, and the slide never changes size.
      // Round 082 reported `shortfall: 23` on exactly this, beside its own
      // `before: 3, after: 3`.
      const slide = makeSlide("s1");
      installHost([slide]);
      const scene = buildChart(config);
      await insertSceneIntoSlide(scene, { tagData: "cfg" });
      expect(
        slide.created.some((s) => s.type === "group"),
        "this test needs a GROUPED chart",
      ).toBe(true);
      const found = (await listChartsInDeck()).charts;
      setTracing(true);
      try {
        await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: "cfg" } }]);
        const d = orphanLine()?.data as { growth: number; removedCalls: number; drewInner: number; atRisk: number };
        expect(d, "the instrument did not run").toBeTruthy();
        expect(d.growth, "a group replaced by a group leaves the slide exactly as it was").toBe(0);
        // AND IT WAS NEVER AT RISK. A group is deleted whole, so it cannot
        // strand anything — which is why zero growth here proves nothing about
        // the question. `partIds` alone cannot say this: a grouped chart has
        // none either, so it looks identical to the dangerous case without the
        // host's own `type`.
        expect(d.atRisk, "a group can strand nothing and must not be counted as exposure").toBe(0);
        // The two numbers whose difference used to be reported as stranding.
        // Kept as context, in their own units, and never subtracted.
        expect(d.drewInner, "a chart holds more than one shape").toBeGreaterThan(1);
        expect(d.removedCalls, "a group goes in one delete").toBe(1);
      } finally {
        setTracing(false);
      }
    });

    it("keeps the second read and marks it unsettled, rather than dropping both", async () => {
      // ROUND 084, AND THE ONLY NON-ZERO NUMBER THIS INSTRUMENT HAS EVER
      // PRODUCED. Four slides reported growing by 23 shapes each. The deck
      // inventory taken at the end of that same round showed each of them
      // holding ONE shape named `PowerChart` — a group. Every one had logged
      // `grouped the chart's shapes {partial:0}` 1.3 seconds earlier, from a
      // sync that had already resolved.
      //
      // The count was exactly `drewInner`, and 24 rather than 25: the host had
      // caught up with the delete and all three draw batches, and not with the
      // addGroup. A single read inside that window invents a chart-sized leak.
      //
      // A phantom is worse than a gap. The whole point of the reading is to say
      // whether an update strands shapes, and a false positive sends someone
      // hunting a bug that is not there.
      const slide = makeSlide("s1");
      installHost([slide]);
      const scene = buildChart(config);
      await insertSceneIntoSlide(scene, { tagData: "cfg" });
      const found = (await listChartsInDeck()).charts;
      _setReReadRetryDelayForTest(1);
      setTracing(true);
      try {
        // Three counts happen for one slide: the before-reading, then the
        // after-reading twice. Lagging the first TWO leaves the two
        // after-reads disagreeing, which is the round-084 shape — the earlier
        // one answering as though the group were still loose shapes.
        faults.shapeCountLag = 2;
        faults.shapeCountLagBy = 23;
        await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: "cfg" } }]);
        // KEPT AND FLAGGED, not dropped. This test used to assert the reading
        // vanished — and dropping it threw away the only half that is ever
        // right. Measured across the archive: of 76 discarded readings the deck
        // can adjudicate all 76, the SECOND read matched its final count 48
        // times, and the FIRST matched it ZERO times. The first read is the
        // known-stale one, taken before the host has caught up with an addGroup
        // it has already committed — which is why it reads 24 where the slide
        // holds 1.
        const line = orphanLine();
        expect(line, "threw away the second read, which is the one the deck backs").toBeTruthy();
        const d = line?.data as { settled: boolean };
        // The flag has to be a MEASUREMENT. It was hardcoded `true`, which was
        // only ever correct because a disagreeing slide never got this far.
        expect(d.settled, "asserted agreement about a reading that had none").toBe(false);
        const unsettled = traceLog().entries.find((e) => /would not settle/.test(e.message));
        expect(unsettled, "took the second read silently").toBeTruthy();
      } finally {
        setTracing(false);
        faults.shapeCountLag = 0;
        faults.shapeCountLagBy = 0;
        _setReReadRetryDelayForTest(1500);
      }
    });

    it("stays silent when nobody is tracing", async () => {
      // It costs an extra context per update. That is a price a round pays and
      // an ordinary edit must not.
      const slide = makeSlide("s1");
      installHost([slide], [], slide, (v) => v !== "1.8");
      const scene = buildChart(config);
      await insertSceneIntoSlide(scene, { tagData: "cfg" });
      const found = (await listChartsInDeck()).charts;
      await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: "cfg" } }]);
      setTracing(true);
      try {
        expect(orphanLine(), "the instrument ran with tracing off").toBeFalsy();
      } finally {
        setTracing(false);
      }
    });
  });

  it("does not resurrect a chart whose shape the user deleted", async () => {
    // The pane still holds an editTarget for a chart the user has since removed
    // from the slide. A stale SLIDE id is already treated as nothing to do; a
    // stale SHAPE id redrew the chart at its old position instead.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group")!;
    for (const s of slide.created) s.delete(); // the user deletes the chart
    await updateChartInSlide(buildChart(config), { slideId: "s1", shapeId: group.id, left: 10, top: 20 });
    expect(slide.created.filter((s) => !s.deleted)).toHaveLength(0);
  });

  /**
   * Grouping and tagging are best-effort by design — the shapes are on the
   * slide before either runs, and the catches around them say exactly that. But
   * the statements that QUEUE the work sat outside those catches, so a host
   * that threw while being asked to re-read a shape collection ("e.load is not
   * a function", seen on the web) rejected the whole request context and failed
   * an update that had, on the slide, already succeeded. Same Scale across the
   * deck reported one TypeError for four redrawn charts.
   */
  it("keeps a redraw that landed, when the host faults on the re-read after it", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group")!;
    const before = slide.created.filter((s) => !s.deleted).length;

    faults.faultShapeCollectionLoad = true;
    // Small batches, so the redraw spans several and asks for the re-read that
    // faults — the exact condition on a live canvas.
    await expect(
      updateChartInSlide(scene, { slideId: "s1", shapeId: group.id, left: 10, top: 20 }, { shapesPerSync: 2 }),
    ).resolves.not.toThrow();
    faults.faultShapeCollectionLoad = false;

    // The chart is on the slide: the old one gone, a new one drawn. What the
    // fault costs is the group and the tag, not the chart.
    expect(group.deleted).toBe(true);
    expect(slide.created.filter((s) => !s.deleted).length).toBeGreaterThanOrEqual(before - 1);
  });
});

describe("selection readers", () => {
  it("loadChartFromSelection returns the tagged config and target", async () => {
    const slide = makeSlide("s1");
    const chart = makeShape("group", undefined, { left: 10, top: 20, width: 300, height: 200 });
    chart.tagStore.set(CHART_TAG, '{"kind":"pie"}');
    const other = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 5, height: 5 });
    installHost([slide], [other, chart]);
    const res = await loadChartFromSelection();
    expect(res?.configJson).toBe('{"kind":"pie"}');
    expect(res?.target).toMatchObject({ slideId: "s1", shapeId: chart.id, left: 10, top: 20 });
  });

  it("loadChartFromSelection returns null for untagged selections", async () => {
    const slide = makeSlide("s1");
    installHost([slide], [makeShape("geometric", "rectangle", { left: 0, top: 0, width: 5, height: 5 })]);
    expect(await loadChartFromSelection()).toBeNull();
  });

  it("getSelectionBounds returns plain shape bounds but skips charts and multi-selects", async () => {
    const slide = makeSlide("s1");
    const box = makeShape("geometric", "rectangle", { left: 7, top: 8, width: 100, height: 60 });
    installHost([slide], [box]);
    expect(await getSelectionBounds()).toEqual({ left: 7, top: 8, width: 100, height: 60 });

    box.tagStore.set(CHART_TAG, "{}");
    installHost([slide], [box]);
    expect(await getSelectionBounds()).toBeNull();

    installHost([slide], [box, makeShape("geometric", "rectangle", { left: 0, top: 0, width: 1, height: 1 })]);
    expect(await getSelectionBounds()).toBeNull();
  });

  it("getSelectionBounds swallows host errors", async () => {
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("no selection");
      },
    });
    expect(await getSelectionBounds()).toBeNull();
  });

  it("listChartsInSelection filters to tagged shapes", async () => {
    const slide = makeSlide("s1");
    const a = makeShape("group", undefined, { left: 1, top: 1, width: 1, height: 1 });
    a.tagStore.set(CHART_TAG, "{}");
    const b = makeShape("geometric", "rectangle", { left: 2, top: 2, width: 1, height: 1 });
    installHost([slide], [a, b]);
    const res = await listChartsInSelection();
    expect(res).toHaveLength(1);
    expect(res[0].target.shapeId).toBe(a.id);
  });
});

describe("listChartsInDeck", () => {
  /**
   * The shapes that are NOT charts.
   *
   * The scan has always held them and always dropped them, which is the right
   * trade for rescaling and repair and the wrong one for a diagnostic: "41
   * shapes became 79", "the chart landed ungrouped", "the slide still holds what
   * was there before" are all questions about the shapes a chart scan discards,
   * and answering them has meant asking the owner to save the deck and upload
   * it.
   */
  it("hands back every shape on every slide, when asked", async () => {
    const s1 = makeSlide("s1");
    const s2 = makeSlide("s2");
    installHost([s1, s2]);
    const loose = s2.shapes.addGeometricShape("rectangle", { left: 3, top: 4, width: 1, height: 1 });
    loose.name = "not a chart";

    const scan = await listChartsInDeck({ withInventory: true });
    const second = scan.inventory?.find((s) => s.slideId === "s2");
    expect(second?.index, "the inventory has to say where in the deck a slide is").toBe(1);
    expect(second?.shapes.map((s) => s.name)).toContain("not a chart");
    expect(second?.shapes[0].left).toBe(3);
    /**
     * AND ITS EXTENT, which is what turns the inventory from a list of origins
     * into something that can witness a geometry fault.
     *
     * Without a size, a chart that landed 200pt off the right edge is recorded
     * identically to one that fits, and two charts drawn over each other
     * identically to two side by side. Across all 322 archived rounds 10,362
     * shapes carry a numeric left/top and NOT ONE carries a size — so the only
     * geometry fault the archive can express is an origin outside the slide.
     * The off-slide chart found on 2026-09-01 was visible solely because that
     * failure happened to move the origin rather than the extent.
     */
    expect(second?.shapes[0].width, "the inventory cannot see how big anything is").toBe(1);
    expect(second?.shapes[0].height).toBe(1);
    // ASKED FOR, not merely read. The fake answers every property whether it was
    // loaded or not, so the assertions above pass even if the load spec never
    // requests a size — while a real host throws PropertyNotLoaded on exactly
    // that. The request is the half only this can check.
    expect(lastShapeLoadSpec(), "the inventory scan never asked the host for a size").toContain("items/width");
    expect(lastShapeLoadSpec()).toContain("items/height");
    // Charts are unaffected — the inventory rides along, it does not replace.
    expect(scan.charts).toHaveLength(0);
  });

  it("still describes a shape whose size the host will not answer", async () => {
    /**
     * Every inventory property goes through `loadedValue` because a host that
     * answered the collection has not necessarily answered every property on it
     * — "a diagnostic that throws while describing the deck is worse than one
     * that reports a shape with no name". Size is no different, and this host
     * refuses geometry often enough that it is the likely case, not the corner.
     *
     * The shape must still be listed with the size simply ABSENT, so a consumer
     * reads "not measured" rather than a zero it would treat as a fact.
     */
    const s1 = makeSlide("s1");
    installHost([s1]);
    const shape = s1.shapes.addGeometricShape("rectangle", { left: 5, top: 6, width: 7, height: 8 });
    shape.name = "sizeless";
    for (const prop of ["width", "height"])
      Object.defineProperty(shape, prop, {
        get() {
          throw new Error("PropertyNotLoaded");
        },
      });

    const scan = await listChartsInDeck({ withInventory: true });
    const only = scan.inventory?.[0]?.shapes.find((s) => s.name === "sizeless");
    expect(only, "a shape vanished from the inventory because its size would not read").toBeTruthy();
    expect(only?.left, "the properties that DID answer were thrown away too").toBe(5);
    expect(only?.width, "an unreadable size was invented rather than left absent").toBeUndefined();
    expect(only?.height).toBeUndefined();
  });

  it("costs nothing on the paths that did not ask for it", async () => {
    // `items/name` is a per-shape string deck-wide, and the callers that scan
    // every slide on a live web host — Same Scale, the repair pass, five
    // self-test scenarios — are exactly the ones that must not pay for a
    // diagnostic. The default has to be the old request, unchanged.
    const s1 = makeSlide("s1");
    installHost([s1]);
    s1.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    const scan = await listChartsInDeck();
    expect(scan.inventory, "the inventory came back unasked").toBeUndefined();
    expect(lastShapeLoadSpec(), "the default scan started asking for shape names").toBe(
      "items/id,items/left,items/top",
    );
  });

  it("finds tagged charts across all slides", async () => {
    const s1 = makeSlide("s1");
    const s2 = makeSlide("s2");
    installHost([s1, s2]);
    await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
    const g = s2.shapes.addGroup([]);
    g.tagStore.set(CHART_TAG, '{"b":2}');
    s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });

    const found = await listChartsInDeck();
    expect(found.charts).toHaveLength(2);
    expect(found.charts.map((f) => f.target.slideId).sort()).toEqual(["s1", "s2"]);
    // The deck-wide sweep releases its proxy objects (one tag + one shape per
    // shape on every slide) once their values are read.
    expect(untracked.tags).toBeGreaterThan(0);
    expect(untracked.shapes).toBeGreaterThan(0);
  });
});

describe("proxy lifecycle", () => {
  it("listChartsInSelection untracks the shape and tag proxies it scans", async () => {
    const slide = makeSlide("s1");
    const a = makeShape("group", undefined, { left: 1, top: 1, width: 1, height: 1 });
    a.tagStore.set(CHART_TAG, "{}");
    const b = makeShape("geometric", "rectangle", { left: 2, top: 2, width: 1, height: 1 });
    installHost([slide], [a, b]);
    await listChartsInSelection();
    expect(untracked.tags).toBe(6); // both selected shapes' config, parts AND origin tags
    expect(untracked.shapes).toBe(2);
  });
});

describe("a grouped chart can be updated in place", () => {
  // THE REASON THE IN-PLACE UPDATE HAD NEVER RUN. A grouped chart carries no
  // CHART_PARTS_TAG — correctly, because its shapes live inside the group and
  // are deleted with it — and the update path refused every chart without one.
  // On this host grouping nearly always succeeds: round 142 grouped 18 of 21
  // charts, wrote a parts list for none of them, and the feature had gone 117
  // archived rounds without a single success.
  //
  // office-js#3014 says grouped sub-shapes "cannot be reached". That note is
  // from 2022 and is out of date: `ShapeGroup.shapes` is reachable through
  // `shape.group` at PowerPointApi 1.8, the set this project already requires
  // for `getImageAsBase64`.
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
  };

  it("reads the group's members instead of refusing for want of a parts tag", async () => {
    setTracing(true);
    try {
      const slide = makeSlide("s1");
      // Grouping SUCCEEDS and 1.8 is present — the real host's ordinary case,
      // and the one that was never able to update in place.
      installHost([slide], [], slide, () => true);
      await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });

      const charts = (await listChartsInDeck()).charts;
      expect(charts, "the fixture drew no chart").toHaveLength(1);
      const target = charts[0].target;
      // The premise: this chart genuinely has no parts list. If the fixture ever
      // starts writing one, this test stops testing what it says it does.
      expect(target.partIds?.length ?? 0, "a grouped chart should carry no parts tag").toBe(0);

      traceLog().entries.length = 0;
      await updateChartsInSlides([
        {
          scene: buildChart({ ...cfg, data: { categories: ["A", "B"], series: [{ name: "S", values: [9, 1] }] } }),
          target,
          opts: { tagData: JSON.stringify(cfg) },
        },
      ]);

      const refused = traceLog().entries.find((e) => e.message === "not updating in place — redrawing instead");
      const why = refused ? String(refused.data?.why ?? "") : "";

      // THE STRONGEST FORM, available again now the share limit admits this
      // edit. Under 0.5 and 0.6 this small chart changed too much and fell
      // back, so every assertion here could do was check WHICH reason it gave —
      // and the note below says why that is weak: the refusal just moves one
      // reason along. Claim the positive outcome, and keep the two negatives as
      // guards in case a refusal ever returns.
      expect(
        traceLog().entries.some((e) => e.message === "updated only the shapes that changed"),
        "a grouped chart was not updated in place at all",
      ).toBe(true);
      expect(why, "still refused a grouped chart — the group members were not read").not.toMatch(/no parts list/);

      // AND THE COUNT MUST LINE UP, which is what the first version got wrong
      // and this suite did not catch. It removed the anchor by matching the
      // TAGGED SHAPE'S id — but for a grouped chart the tagged shape IS THE
      // GROUP, and a group is never among its own members, so nothing was
      // removed and every group came back one member too long.
      //
      // Round 143 found it in five of five cases (parts 25 against nodes 24,
      // parts 17 against nodes 16) while 3209 tests stayed green. Asserting
      // "not refused for a missing parts list" was too weak: the refusal simply
      // moved to the next reason along. This asserts the number.
      expect(why, "the group came back the wrong length — the anchor was not removed").not.toMatch(/one for one/);
    } finally {
      setTracing(false);
    }
  });

  it("still refuses when the group's members do not line up with the scene", async () => {
    // THE GUARD THAT MAKES THE ABOVE SAFE. The mapping is positional — node 0 is
    // the anchor, the rest line up in drawing order — so a group holding a
    // different number of shapes than the scene has nodes must be refused, not
    // guessed at. Writing into a mismatched mapping would corrupt a chart the
    // user is looking at, which is worse than the redraw it falls back to.
    setTracing(true);
    try {
      const slide = makeSlide("s1");
      installHost([slide], [], slide, () => true);
      await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
      const charts = (await listChartsInDeck()).charts;
      const target = charts[0].target;

      traceLog().entries.length = 0;
      // A DIFFERENT SHAPE OF CHART: more categories means more nodes, so the
      // group's members cannot line up one for one.
      const wider: ChartConfig = {
        ...cfg,
        data: { categories: ["A", "B", "C", "D"], series: [{ name: "S", values: [1, 2, 3, 4] }] },
      };
      await updateChartsInSlides([{ scene: buildChart(wider), target, opts: { tagData: JSON.stringify(wider) } }]);

      const refused = traceLog().entries.find((e) => e.message === "not updating in place — redrawing instead");
      expect(refused, "wrote into a mapping that does not line up").toBeDefined();
      // ANY of the refusals is correct here, and which one fires is not the
      // point — a chart whose shape changed is caught by the node-compatibility
      // check BEFORE the one-for-one count is even reached. The assertion is
      // that it refuses rather than guesses, because a wrong mapping writes a
      // scrambled chart instead of falling back to the redraw.
      expect(String(refused!.data?.why ?? "")).toMatch(/one for one|no parts list|no readable group|node-compatible/);
    } finally {
      setTracing(false);
    }
  });
});

describe("in-place update keeps the chart where it is", () => {
  // The tagged shape's left/top is NOT the frame origin: grouped it is the
  // group's bounding box, ungrouped it is whatever created[0] happens to be.
  // Feeding it back as the next render's origin shifted the chart by the scene's
  // ink offset — and compounded, so Same Scale walked charts off the slide.
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
  };

  const cycle = async (grouped: boolean) => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, grouped ? () => true : (v) => v !== "1.8");
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
    const seen: string[] = [];
    for (let i = 0; i < 3; i++) {
      const charts = (await listChartsInDeck()).charts;
      expect(charts).toHaveLength(1);
      const t = charts[0].target;
      seen.push(`${Math.round(t.left)},${Math.round(t.top)}`);
      await updateChartsInSlides([{ scene: buildChart(cfg), target: t, opts: { tagData: JSON.stringify(cfg) } }]);
    }
    return seen;
  };

  it("does not drift when the host groups (desktop)", async () => {
    const seen = await cycle(true);
    expect(new Set(seen).size, `chart moved across update cycles: ${seen.join(" -> ")}`).toBe(1);
  });

  it("follows a chart the user has DRAGGED, instead of teleporting it back", async () => {
    // Re-rendering at the tagged shape's corner drifts; re-rendering at the
    // recorded origin teleports a moved chart back to where it was first
    // inserted, silently undoing the user's drag. The origin tag therefore also
    // records the ANCHOR (where the tagged shape landed), so an update shifts the
    // origin by exactly how far the shape has moved since.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });

    const before = (await listChartsInDeck()).charts[0].target;
    // The user drags the whole chart across the slide.
    const [dx, dy] = [240, 110];
    for (const sh of slide.created) {
      sh.left += dx;
      sh.top += dy;
    }
    const moved = (await listChartsInDeck()).charts[0].target;
    expect(moved.left).toBeCloseTo(before.left + dx, 5);

    await updateChartsInSlides([{ scene: buildChart(cfg), target: moved, opts: { tagData: JSON.stringify(cfg) } }]);
    const after = (await listChartsInDeck()).charts[0].target;
    expect(after.left, "update dragged the chart back to its insert position").toBeCloseTo(moved.left, 5);
    expect(after.top, "update dragged the chart back to its insert position").toBeCloseTo(moved.top, 5);
  });

  it("follows a drag even when the caller reuses a CACHED target (what the pane does)", async () => {
    // The pane captures state.editTarget once, when the chart is loaded, and
    // re-uses it for every subsequent "Update chart". Measuring the drag against
    // that snapshot reports no movement, so the update put the chart back where
    // it was — the same teleport, just reached through the pane's real flow
    // rather than a fresh deck read.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
    const held = (await listChartsInDeck()).charts[0].target; // captured once, then kept

    for (const sh of slide.created.filter((s) => !s.deleted)) {
      sh.left += 200;
      sh.top += 60;
    }
    const moved = (await listChartsInDeck()).charts[0].target;

    // The pane updates with the STALE target it has been holding all along.
    await updateChartInSlide(buildChart(cfg), held, { tagData: JSON.stringify(cfg) });
    const after = (await listChartsInDeck()).charts[0].target;
    expect(after.left, "update teleported the chart back").toBeCloseTo(moved.left, 5);
    expect(after.top, "update teleported the chart back").toBeCloseTo(moved.top, 5);
  });

  it("does not drift when the host cannot group (web)", async () => {
    const seen = await cycle(false);
    expect(new Set(seen).size, `chart moved across update cycles: ${seen.join(" -> ")}`).toBe(1);
  });
});

describe("repeated in-place updates keep landing", () => {
  // An update replaces every shape, so the caller's EditTarget is dead the
  // moment it returns. Before updateChartsInSlides handed back a fresh one, the
  // SECOND update named a shape that no longer existed, was filtered out as
  // "the user deleted this chart", and did nothing — with no error. Auto-update
  // (which fires 900ms after every control change) died the same way after one
  // push, so the pane looked like "Update stopped working".
  it("a second update against the returned target still lands", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: '{"v":0}' });
    let target = (await listChartsInDeck()).charts[0].target;

    for (const v of [1, 2, 3]) {
      const next = await updateChartInSlide(buildChart(cfg), target, { tagData: `{"v":${v}}` });
      expect(next, `update ${v} returned no target`).toBeTruthy();
      target = next!;
      // The edit actually reached the slide, every time.
      const live = (await listChartsInDeck()).charts;
      expect(live, `update ${v} lost the chart`).toHaveLength(1);
      expect(JSON.parse(live[0].configJson).v, `update ${v} silently did nothing`).toBe(v);
    }
  });
});

describe("insertAgendaSlides", () => {
  it("appends one slide per chapter and renders ungrouped", async () => {
    const s1 = makeSlide("s1");
    const slides = [s1];
    installHost(slides);
    const chapters = ["Intro", "Findings", "Next steps"];
    const scenes = chapters.map((_, i) => buildAgendaScene(chapters, { highlight: i }));
    await insertAgendaSlides(scenes);
    expect(slides).toHaveLength(4);
    for (let i = 1; i < 4; i++) {
      expect(slides[i].created.length).toBeGreaterThan(0);
      expect(slides[i].created.some((s) => s.type === "group")).toBe(false);
    }
  });
});

describe("insertDemoDeck", () => {
  it("appends one slide per item and tags the charts with their config", async () => {
    const s1 = makeSlide("s1");
    const slides = [s1];
    installHost(slides);
    const items = [
      {
        scene: buildChart({
          ...DEFAULT_SIZE,
          kind: "pie" as const,
          data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
        }),
        tagData: '{"kind":"pie"}',
      },
      {
        scene: buildChart({
          ...DEFAULT_SIZE,
          kind: "clustered" as const,
          data: { categories: ["A"], series: [{ name: "S", values: [5] }] },
        }),
        tagData: '{"kind":"clustered"}',
      },
      {
        scene: {
          width: 100,
          height: 40,
          nodes: [{ kind: "rect" as const, x: 0, y: 0, w: 10, h: 10, fill: "#111111" }],
        },
      }, // untagged element
    ];
    await insertDemoDeck(items);
    // Three slides appended after the original.
    expect(slides).toHaveLength(4);
    for (let i = 1; i < 4; i++) expect(slides[i].created.length).toBeGreaterThan(0);
    // The two chart slides carry their config tag; the element slide does not.
    expect(slides[1].created.some((s) => s.tagStore.get(CHART_TAG) === '{"kind":"pie"}')).toBe(true);
    expect(slides[3].created.every((s) => !s.tagStore.has(CHART_TAG))).toBe(true);
  });
});

describe("isPowerPointHost", () => {
  it("is false outside an Office host and true inside", () => {
    expect(isPowerPointHost()).toBe(false);
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", { context: { host: "PowerPoint" } });
    expect(isPowerPointHost()).toBe(true);
  });
});

describe("marker symbols in the live add-in", () => {
  const markerScene = (markers: MarkerSymbol[]) =>
    buildChart({
      kind: "scatter",
      width: 480,
      height: 300,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 2, 3] },
          { name: "Y", values: [2, 4, 3] },
          { name: "Group", values: [1, 2, 3] },
        ],
      },
      scatter: { markers },
    });

  it("draws each symbol as native preset geometry, filled", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(markerScene(["diamond", "plus", "triangle"]), { left: 0, top: 0 });

    // Filled preset geometry is the whole reason a symbol is not a polygon:
    // PowerPoint can only outline a freeform, so a polygon marker would be
    // hollow here while the SVG preview showed it solid.
    for (const preset of ["diamond", "plus", "triangle"]) {
      const shapes = slide.created.filter((s) => s.geo === preset);
      expect(shapes.length, preset).toBeGreaterThan(0);
      for (const s of shapes) {
        expect(s.fillColor, preset).toMatch(/^#[0-9a-f]{6}$/i);
        expect(s.fillCleared, preset).toBe(false);
        expect(s.box.width, preset).toBeGreaterThan(0);
        expect(s.box.width, preset).toBeCloseTo(s.box.height, 9);
      }
    }
  });

  it("needs no rotation, so it works on a bare 1.4 host", async () => {
    // Arrowheads and pie fans need Shape.rotation (1.10+) and degrade without
    // it. The marker set is deliberately rotation-free: nothing here may set
    // rotation, so a 1.4 host draws the same shapes as a current one.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => false);
    await insertSceneIntoSlide(markerScene(["diamond", "triangle", "plus"]), { left: 0, top: 0 });
    const presets = slide.created.filter((s) => ["diamond", "triangle", "plus"].includes(s.geo!));
    expect(presets.length).toBeGreaterThan(0);
    for (const s of presets) expect(s.rotation).toBeUndefined();
  });

  it("places the symbol's box centred on the point", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = markerScene(["diamond", "diamond", "diamond"]);
    await insertSceneIntoSlide(scene, { left: 100, top: 50 });
    const nodes = scene.nodes.filter((n) => n.kind === "symbol");
    expect(nodes.length).toBeGreaterThan(0);
    for (const n of nodes) {
      if (n.kind !== "symbol") continue;
      const s = slide.created.find((c) => c.geo === "diamond" && Math.abs(c.box.left - (100 + n.cx - n.size)) < 1e-6);
      expect(s, `no shape at cx=${n.cx}`).toBeTruthy();
      expect(s!.box.top).toBeCloseTo(50 + n.cy - n.size, 9);
      expect(s!.box.width).toBeCloseTo(n.size * 2, 9);
    }
  });
});

describe("Office round-trips do not scale with the chart count", () => {
  const cfgFor = (v: number): ChartConfig => ({
    ...config,
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [v, v + 1] }] },
  });
  const targetsOn = (slide: FakeSlide, n: number) =>
    Array.from({ length: n }, (_, i) => {
      // A real target names a shape that exists on the slide; make one per chart.
      const s = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
      return {
        scene: buildChart(cfgFor(i)),
        target: { slideId: slide.id, shapeId: s.id, left: 10, top: 20 },
        opts: { tagData: `{"i":${i}}` },
      };
    });

  it("re-renders N charts in ONE context, whatever N is", async () => {
    // The defect this guards: doSameScale awaited the single-chart update in a
    // loop, so each chart opened its own PowerPoint.run — 20 contexts across a
    // 20-chart deck. That is the property worth pinning. The SYNC count is no
    // longer flat and must not be: shapes commit in batches, because a live
    // canvas will not take a whole chart at once (SHAPES_PER_SYNC).
    for (const n of [1, 2, 10, 20]) {
      const slide = makeSlide("s1");
      installHost([slide]);
      await updateChartsInSlides(targetsOn(slide, n));
      expect(trips.contexts, `${n} charts`).toBe(1);
    }
  });

  it("costs syncs per BATCH OF SHAPES, not a fixed toll per chart", async () => {
    // The two failure modes this sits between: a per-chart context (the old
    // N+1, 80 round-trips for 20 charts), and a per-chart mega-batch that the
    // host silently refuses. Syncs must track the shapes, and nothing else.
    const slide = makeSlide("s1");
    installHost([slide]);
    await updateChartsInSlides(targetsOn(slide, 1));
    const one = trips.syncs;
    installHost([makeSlide("s2")]);
    const slide2 = makeSlide("s2");
    installHost([slide2]);
    await updateChartsInSlides(targetsOn(slide2, 2));
    const two = trips.syncs;
    // Doubling the charts doubles the drawing, not a fixed per-chart overhead:
    // the growth is the extra shapes' batches, so it stays well under 2x.
    expect(two).toBeGreaterThan(one);
    expect(two).toBeLessThan(one * 2 + 2);
  });

  it("still draws every chart it batches, tagged and grouped", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const items = targetsOn(slide, 3);
    await updateChartsInSlides(items);
    const groups = slide.created.filter((s) => s.type === "group");
    expect(groups).toHaveLength(3);
    // Each group carries its OWN config, not the last one written.
    expect(groups.map((g) => g.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":1}', '{"i":2}']);
    // The old shape each target named is gone.
    for (const it of items) expect(slide.created.find((s) => s.id === it.target.shapeId)!.deleted).toBe(true);
    // Charts land at their target's position, not the default offset.
    for (const r of slide.created.filter((s) => s.geo === "rectangle" && s.box.width > 1)) {
      expect(r.box.left).toBeGreaterThanOrEqual(10);
    }
  });

  it("keeps the single-chart paths to ONE context each", async () => {
    // updateChartInSlide is now updateChartsInSlides([one]); the Insert button
    // opens its own. Neither may open more than one, however many shapes the
    // chart has.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "{}" });
    expect(trips.contexts).toBe(1);

    const slide2 = makeSlide("s2");
    installHost([slide2]);
    const s = slide2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    await updateChartInSlide(buildChart(config), { slideId: "s2", shapeId: s.id, left: 0, top: 0 }, { tagData: "{}" });
    expect(trips.contexts).toBe(1);
  });

  it("does nothing, and opens no context, for an empty batch", async () => {
    installHost([makeSlide("s1")]);
    await updateChartsInSlides([]);
    expect([trips.syncs, trips.contexts]).toEqual([0, 0]);
  });

  it("keeps every chart re-editable when the grouping sync is refused", async () => {
    // Batching costs granularity: a refused grouping now loses grouping for the
    // whole batch, not just one chart. What must NOT be lost is the config tag —
    // the charts are already on the slide (their shapes committed a phase
    // earlier), so each must fall back to tagging its own first shape or it
    // silently stops being re-editable.
    //
    // The failure has to come from the SYNC, not from addGroup: Office.js only
    // reports queued commands there, which means every tag target has already
    // been pointed at a group that turned out not to exist. A test that throws
    // from addGroup instead never overwrites them and proves nothing.
    const slide = makeSlide("s1");
    installHost([slide]);
    // The group sync is no longer a fixed number: the shapes commit in batches
    // first, so its index depends on the chart's size. Find it rather than
    // hardcode it — a wrong number here silently tests nothing.
    const batches = Math.ceil(buildChart(cfgFor(0)).nodes.length / 10);
    // 1 resolve slides, 1 resolve old shapes, then PER CHART one delete sync
    // plus its render batches, then GROUP. The delete is per chart because a
    // shared one commits every chart's removal before any redraw runs — see
    // updateChartsInSlides.
    // ...plus the PRE-GROUPING RE-READ, which is new since 2026-08-16: every
    // groupable chart asks for one now, not just those that span batches
    // (`needsPreGroupRefresh`), and all of them refresh in a single shared sync.
    // Missing it here does not fail loudly — it fails the WRONG sync and the
    // test quietly stops exercising the refused group, which is the trap the
    // comment above is about.
    faults.failSyncOn = 2 + 3 /* charts */ * (1 + batches) + 1 /* the re-read */ + 1; /* the group */
    try {
      const items = targetsOn(slide, 3);
      // One refreshed target per chart — the caller needs them to stay live.
      await expect(updateChartsInSlides(items)).resolves.toHaveLength(items.length);
      // Each chart's OWN config, back on each chart's OWN first shape.
      const tagged = slide.created.filter((s) => s.tagStore.has(CHART_TAG));
      expect(tagged.map((s) => s.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":1}', '{"i":2}']);
      expect(tagged.every((s) => s.type !== "group")).toBe(true);
    } finally {
      faults.failSyncOn = 0;
    }
  });

  it("does not blank the rest of the deck when one chart's redraw stalls", async () => {
    // The defect: every chart's old shapes were deleted in ONE shared sync,
    // and only then were the charts redrawn one at a time. Those deletes are
    // committed, so a single stalled redraw rejected the whole PowerPoint.run
    // and left every chart AFTER it blank — old shapes gone, new ones never
    // queued. Same Scale runs across the whole deck and necessarily includes
    // the chart on the visible slide, which is the one condition documented
    // here as reliably stalling a redraw. So one slow chart emptied the deck.
    const slide = makeSlide("s1");
    installHost([slide]);
    const batches = Math.ceil(buildChart(cfgFor(0)).nodes.length / 10);
    // 2 resolve syncs, then per chart 1 delete + `batches` renders. Fail the
    // FIRST render batch of chart 2 (0-based index 1).
    faults.failSyncOn = 2 + (1 + batches) + 1 + 1;
    const failed: string[] = [];
    try {
      const items = targetsOn(slide, 3);
      const next = await updateChartsInSlides(items, (it) => failed.push(it.opts!.tagData!));
      // Chart 2 is reported, once, by name.
      expect(failed).toEqual(['{"i":1}']);
      // Charts 1 and 3 are re-editable — they redrew and carry their OWN tags.
      const tagged = slide.created.filter((s) => s.tagStore.has(CHART_TAG));
      expect(tagged.map((s) => s.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":2}']);
      // And the caller's targets are not shifted onto each other: chart 3 must
      // NOT be handed chart 1's new shape id. Indexing `tagged` by position
      // would do exactly that once a chart drops out of the batch.
      expect(next).toHaveLength(3);
      expect(new Set(next.map((t) => t.shapeId)).size).toBe(3);
      expect(next[1].shapeId).toBe(items[1].target.shapeId); // unchanged: it failed
    } finally {
      faults.failSyncOn = 0;
    }
  });

  it("tells the caller what EVERY stalled chart destroyed, not just the first", async () => {
    // `onFailed` is the only channel a deck-wide caller has for finding out
    // that a chart went blank — and, through the wreckage on the error, what
    // was left on its slide. Same Scale sweeps that debris; it cannot sweep
    // what it is not told about.
    //
    // The bug was subtle because failure #1 looked fine: `wreck()` mutates the
    // host's error object in place, so the raw `err` handed to `onFailed`
    // happened to carry the wreckage anyway. But `wreck()` ran INSIDE the
    // `firstFailure === undefined` test, so failures #2..n were never
    // annotated at all — and a deck-wide update is exactly the caller that
    // gets more than one. Hence two stalls here: one would pass either way.
    const slide = makeSlide("s1");
    installHost([slide]);
    const batches = Math.ceil(buildChart(cfgFor(0)).nodes.length / 10);
    // Sync map: 2 resolves, then per chart 1 delete + up to `batches` renders.
    // A chart that fails on its first render batch consumes exactly 2. So
    // chart 1's first render is sync 5+batches, and — chart 1 having stopped
    // there — chart 2's is 7+batches.
    failSyncsOn.add(5 + batches);
    failSyncsOn.add(7 + batches);
    const seen: { tag: string; wreckage: ReturnType<typeof wreckageOf> }[] = [];
    try {
      const items = targetsOn(slide, 3);
      await updateChartsInSlides(items, (it, err) => seen.push({ tag: it.opts!.tagData!, wreckage: wreckageOf(err) }));
      // Charts 2 and 3 stalled; chart 1 is untouched and still redrew.
      expect(seen.map((s) => s.tag)).toEqual(['{"i":1}', '{"i":2}']);
      // BOTH carry their wreckage. Pre-fix the second was undefined, so Same
      // Scale swept the first chart's debris and left the rest on the deck.
      for (const s of seen) {
        expect(s.wreckage, `wreckage for ${s.tag}`).toBeDefined();
        expect(s.wreckage!.slideId).toBe("s1");
        expect(Array.isArray(s.wreckage!.strayIds)).toBe(true);
      }
    } finally {
      failSyncsOn.clear();
    }
  });

  it("still throws when the ONE chart it was given fails", async () => {
    // updateChartResilient catches this throw to reach its slide-swap and
    // picture fallbacks. Swallowing a total failure would strand it on layer 1
    // with a chart whose old shapes are already deleted.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.failSyncOn = 4; // 2 resolves, 1 delete, then the first render batch
    try {
      await expect(updateChartsInSlides(targetsOn(slide, 1))).rejects.toThrow();
    } finally {
      faults.failSyncOn = 0;
    }
  });

  it("renders one slide per context and reports progress per slide", async () => {
    // One PowerPoint.run per slide isolates a chart the host can't finish and
    // keeps each context light (one chart's shapes, not a chunk's four). Progress
    // is per slide, so a slow host shows slides landing instead of freezing.
    for (const n of [2, 12, 35] as const) {
      installHost([makeSlide("s1")]);
      const seen: string[] = [];
      const report = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
        (done, total) => seen.push(`${done}/${total}`),
      );
      expect(failedIndices(report), `${n} slides`).toEqual([]);
      // Self-check: the deck grew by exactly one slide per item, nothing lost.
      expect(report.slidesAdded, `${n} slides`).toBe(n);
      // One context per SLIDE, plus one more per slide for addSlides' settled
      // fresh-context verify (nothing lost, so no retry context), the two
      // settled slideCount reads (before/after), and the paged on-slide
      // readback (nothing lost, so it runs).
      expect(trips.contexts, `${n} slides`).toBe(2 * n + 2 + Math.ceil(n / READBACK_PAGE));
      expect(seen, `${n} slides`).toHaveLength(n);
      expect(seen.at(-1)).toBe(`${n}/${n}`);
      // Monotonic, never over-counting.
      expect(seen.map((x) => Number(x.split("/")[0]))).toEqual(
        [...seen.map((x) => Number(x.split("/")[0]))].sort((a, b) => a - b),
      );
    }
  });

  it("appends every demo slide, each tagged with its own config", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 35;
    const report = await insertDemoDeck(
      Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
    );
    expect(failedIndices(report)).toEqual([]);
    expect(report.slidesAdded).toBe(n);
    // The fake appends a slide per add(); the original + n new ones.
    expect(deck.length).toBe(1 + n);
    const tags = deck.slice(1).map((s) => s.created.map((c) => c.tagStore.get(CHART_TAG)).find(Boolean));
    expect(tags).toEqual(Array.from({ length: n }, (_, i) => `{"i":${i}}`));
  });

  it("records per-item and total wall-clock so a run's duration is on the record", async () => {
    installHost([makeSlide("s1")]);
    const n = 5;
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // Every item is timed; the total is present and never negative.
    expect(report.results.every((r) => typeof r.ms === "number" && r.ms >= 0)).toBe(true);
    expect(typeof report.totalMs).toBe("number");
    expect(report.totalMs).toBeGreaterThanOrEqual(0);
    // The whole run is at least as long as its slowest single item.
    expect(report.totalMs).toBeGreaterThanOrEqual(Math.max(...report.results.map((r) => r.ms)));
  });

  it("finds no blank slots and completes the readback on a clean run", async () => {
    installHost([makeSlide("s1")]);
    const n = 4;
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    expect(report.blankSlides).toEqual([]); // every added slide read back with shapes
    expect(report.blanksRead).toBe(true); // the readback finished
    expect(report.addsIssued).toBe(n); // one add per item, no retries/fails
  });

  it("reports a host-blanked slide by DECK POSITION, not by item name", async () => {
    const deck = [makeSlide("s1")]; // pre-existing at index 0; added slides take indices 1..n
    installHost(deck);
    const n = 4;
    blankReadbackAt.add(2); // the added slide at deck index 2 reads back empty on readback
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // A blank slide has no content/tag to name it — reported as the 1-based deck position (index 2 → slide 3).
    expect(report.blankSlides).toEqual([3]);
    expect(report.blanksRead).toBe(true);
  });

  it("names a blank slide from its slot tag when the item carries a title", async () => {
    // A blank readback used to say only "slide 3". Every demo slide now gets a
    // POWERCHART_DEMO_SLOT tag on creation, so the readback can name the missing
    // chart by title. blankReadbackAt makes index 2 report 0 shapes; the slot
    // tag survives (we never emptied the tag store) and gives us the item name.
    _setBlankReReadDelayForTest(0); // no wall-clock sleep in the test
    try {
      const deck = [makeSlide("s1")];
      installHost(deck);
      const n = 3;
      blankReadbackAt.add(2);
      const report = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({
          scene: buildChart(cfgFor(i)),
          title: `chart-${i}`,
        })),
      );
      expect(report.blankSlides).toEqual([3]);
      // Index 2 corresponds to item 1 (item 0 is index 1, item 1 is index 2, ...).
      expect(report.blankItems).toEqual([{ position: 3, title: "chart-1" }]);
    } finally {
      _setBlankReReadDelayForTest(200);
    }
  });

  it("marks the blank readback incomplete when it faults, not falsely clean", async () => {
    installHost([makeSlide("s1")]);
    faults.faultShapeGetCount = true; // every readback getCount throws
    const report = await insertDemoDeck(Array.from({ length: 3 }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // An empty list must NOT read as "no blanks" when we could not measure.
    expect(report.blanksRead).toBe(false);
    expect(report.blankSlides).toEqual([]);
    // The run itself still succeeded — a readback fault is not a render failure.
    expect(report.results.every((r) => r.status === "rendered")).toBe(true);
  });

  it("bypassBudget lets a text-heavy scene render even when its shape count is over the budget", async () => {
    // The results/contents slide bug: 32 failures pushed the results scene to
    // 135 shapes — over DEMO_SHAPE_BUDGET (90) — and the run's own summary
    // came back as a red "NOT COMPLETE" stamp. Text-only scenes don't hit the
    // wedge/polygon flood the budget guards against; they should render.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const denseTextScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "text" as const,
        x: k,
        y: 0,
        w: 40,
        h: 20,
        text: `row ${k}`,
        fontSize: 12,
        color: "#000000",
        align: "left" as const,
        valign: "top" as const,
      })),
    };
    // With bypassBudget: the scene renders as a real chart, no stamp.
    const withBypass = await insertDemoDeck([{ scene: denseTextScene, title: "Results", bypassBudget: true }]);
    expect(withBypass.results[0].status).toBe("rendered");
    expect(deck[1].created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    expect(deck[1].created.filter((s) => s.type === "text").length).toBeGreaterThanOrEqual(120);
    // Without bypassBudget: the scene is stamped instead of drawn.
    installHost([makeSlide("s1")]);
    const withoutBypass = await insertDemoDeck([{ scene: denseTextScene, title: "Results" }]);
    expect(withoutBypass.results[0].status).toBe("skipped");
  });

  it("skips a chart too dense for the host, keeps the slide, and stamps it NOT COMPLETE", async () => {
    // The heavy charts (area ~208 shapes) will not land on web and burn the
    // timeout trying. They are skipped up-front — the slide is kept (coverage),
    // its chart NOT drawn, and a stamp makes the placeholder unmistakable.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 1,
        h: 1,
        fill: "#111111",
      })),
    };
    const light = () => buildChart(cfgFor(0)); // a handful of shapes, well under budget
    const report = await insertDemoDeck([{ scene: light() }, { scene: dense }, { scene: light() }]);
    // Only the dense one is reported (as "skipped"); the deck still has all three.
    expect(failedIndices(report)).toEqual([1]);
    expect(report.results[1].status).toBe("skipped");
    expect(report.slidesAdded).toBe(3);
    expect(deck.length).toBe(1 + 3);
    // The dense slide (deck[2]) carries a stamp, NOT the 120 chart shapes.
    const denseSlide = deck[2];
    const stamp = denseSlide.created.find((s) => s.name === "PowerChart:not-complete");
    expect(stamp, "dense slide is stamped").toBeTruthy();
    expect(stamp!.text).toContain("NOT COMPLETE");
    // A top strip, not a slab over the middle — a mis-targeted stamp must not
    // obliterate a real chart under it (a 540pt-tall slide).
    expect(stamp!.top, "stamp sits at the top").toBeLessThan(80);
    expect(stamp!.height, "stamp is a strip").toBeLessThan(120);
    expect(denseSlide.created.length, "chart not drawn").toBeLessThan(120);
    // The light neighbours rendered as real charts (no stamp).
    expect(deck[1].created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    expect(deck[3].created.length).toBeGreaterThan(1);
  });

  it("does not report a phantom lost slide for a too-dense item whose stamp was refused", async () => {
    // `addsIssued` used to be inferred as "one per item, plus one more for every
    // retried or failed item". A too-dense item never re-renders — there is
    // nothing to re-render, only a placeholder to stamp — so when its stamp sync
    // is refused it ends "failed" having issued exactly ONE add. The inference
    // then charged it a second, and `addsIssued − slidesAdded` accused the host
    // of losing a slide it had actually kept. This harness exists to stop
    // inventing failures, so a false ⚠ is the bug.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 1,
        h: 1,
        fill: "#111111",
      })),
    };
    failSyncsOn.add(6); // the too-dense item's stampSlide — see the test above
    const report = await insertDemoDeck([
      { scene: dense, title: "too dense" },
      { scene: buildChart(cfgFor(1)), tagData: '{"i":1}', title: "fine" },
    ]);
    expect(report.results[0].status).toBe("failed"); // the premise: stamped, refused
    expect(report.results[0].attempts, "one add issued, not two").toBe(1);
    // Both slides are on the deck, so nothing was lost — and the report agrees.
    expect(report.slidesAdded).toBe(2);
    expect(report.addsIssued).toBe(2);
    expect(report.addsIssued - report.slidesAdded, "no phantom loss").toBe(0);
  });

  it("self-check catches a slide the host silently drops (deck grew by less than asked)", async () => {
    // The corruption a visual scan misses and today cost us 3 lost slides: an
    // add() that never lands leaves the deck one slide short with no error the
    // user sees. The report's slidesAdded is read back from the host, so the
    // shortfall is caught — and the dropped item shows up as not rendered.
    //
    // addSlides self-heals dropped adds with its own fresh-context retries, so
    // losing a slide for good takes the add AND every retry round. There is no
    // longer an outer per-item retry on top of that — it was the source of
    // every duplicate slide, and the end-of-run reconcile pass reports what
    // landed with better evidence than a mid-flight guess.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.swallowAdds = ADDS_TO_DEFEAT_ONE_SLIDE; // the add and all its retries → gone for good
    try {
      const report = await insertDemoDeck(
        Array.from({ length: 4 }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
      );
      // 4 items asked for, but the deck only grew by 3 — the lost-slide signal.
      expect(report.slidesAdded).toBe(3);
      expect(report.results).toHaveLength(4); // every item is still accounted for
      expect(failedIndices(report).length).toBeGreaterThanOrEqual(1); // the dropped one is flagged
    } finally {
      faults.swallowAdds = 0;
    }
  });

  it("addSlides self-heals one dropped add via its own fresh-context retry, and surfaces it when every retry also fails", async () => {
    // addSlides now verifies its own adds landed (a settled getCount() in a
    // FRESH context, after the existing 2 syncs) and gets MAX_ADD_RETRY_ROUNDS
    // retry rounds before giving up — the fix for the Presentation_3.pptx bug where
    // PowerPoint web silently dropped ~half of 20 issued add()s. A single
    // dropped add should never even reach insertDemoDeck's own item-level
    // retry: it should be invisible, recovered inside addSlides itself.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.swallowAdds = 1; // the FIRST add() call anywhere is dropped, none after
    try {
      const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)), tagData: '{"i":0}' }]);
      // The retry landed: the deck grew by exactly the one slide asked for.
      expect(report.slidesAdded).toBe(1);
      expect(report.results[0].status).toBe("rendered");
      // Nothing was lost AT COMMIT: the one drop was fully recovered.
      expect(report.addsLostAtCommit).toBe(0);
    } finally {
      faults.swallowAdds = 0;
    }

    // Second sub-case: the drop persists through every retry round too. One
    // addSlides call issues the original add plus one per retry round, so
    // defeating item 0 entirely takes exactly that many dropped adds. A second
    // item follows so the run is not a TOTAL loss (which insertDemoDeck itself
    // would throw on, per "A whole deck lost to HOST errors" below) — item 1
    // renders once faults.swallowAdds is exhausted.
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const deck2: FakeSlide[] = [makeSlide("s1")];
    installHost(deck2);
    faults.swallowAdds = ADDS_TO_DEFEAT_ONE_SLIDE;
    try {
      const report = await insertDemoDeck([
        { scene: buildChart(cfgFor(0)), tagData: '{"i":0}' },
        { scene: buildChart(cfgFor(1)), tagData: '{"i":1}' },
      ]);
      // Item 0 is a genuine total loss; item 1 landed once faults.swallowAdds ran out.
      expect(report.slidesAdded).toBe(1);
      expect(report.results[0].status).toBe("failed");
      expect(report.results[1].status).toBe("rendered");
      // addSlides confirmed the loss — its own retry did not recover it.
      expect(report.addsLostAtCommit).toBeGreaterThanOrEqual(1);
      expect(warnSpy).toHaveBeenCalled();
    } finally {
      faults.swallowAdds = 0;
      warnSpy.mockRestore();
    }
  });

  it("recovers a run of consecutive dropped adds that a single retry round could not", async () => {
    // The reason MAX_ADD_RETRY_ROUNDS is no longer 1. The host that motivates
    // all of this dropped ~half of 20 adds in ONE burst — drops arrive in runs,
    // not singly, and a retry issued under the same load is dropped by the same
    // load. At one round, two consecutive drops cost a slide outright; the deck
    // came back short, the item was reported failed, and the user lost a chart
    // to a condition that was transient the whole time.
    //
    // Two drops is the smallest case that separates the bounds: recoverable
    // now, a total loss before. Guarded against the constant rather than the
    // literal 2 so the case stays "one more drop than a single round can take".
    expect(MAX_ADD_RETRY_ROUNDS).toBeGreaterThan(1);
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.swallowAdds = 2;
    try {
      const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)), tagData: '{"i":0}' }]);
      // Fully recovered: the deck grew by the one slide asked for, the chart
      // rendered, and nothing was written off at commit.
      expect(report.slidesAdded).toBe(1);
      expect(report.results[0].status).toBe("rendered");
      expect(report.addsLostAtCommit).toBe(0);
    } finally {
      faults.swallowAdds = 0;
    }
  });

  it('render:"image" draws ONE picture shape instead of the scene nodes', async () => {
    // The whole point of image mode: a dense chart becomes one shape, so the
    // PowerPoint-web dense-shape wall (office-js #4272 / #5022 / #6498) never
    // gets hit. A violin-sized scene is ~250 native shapes; here it must be 1.
    const NODES = 25;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA", left: 60, top: 90 });
    // Exactly one shape, geometry rectangle, sized to the FRAME (not the nodes).
    const live = slide.created.filter((s) => !s.deleted);
    expect(live).toHaveLength(1);
    expect(live[0].geo).toBe("rectangle");
    expect(live[0].box).toEqual({ left: 60, top: 90, width: 480, height: 300 });
    // The fill became a picture, carrying the payload, and the outline is off.
    expect(live[0].imageBase64).toBe("AAAA");
    expect(live[0].fillType).toBe("PictureAndTexture");
    expect(live[0].lineFormat.visible).toBe(false);
    // Named like a chart group so the Selection Pane reads the same either way,
    // and it carries the config tag — the picture IS the re-editable chart.
    expect(live[0].name).toBe("PowerChart");
    expect(live[0].tagStore.get(CHART_TAG)).toBe("{}");
  });

  it("strips every base64 spelling down to the bare payload the host wants", async () => {
    // Three forms circulate: a browser toDataURL (`data:image/png;base64,…`),
    // what render-pptx.mjs hands pptxgen (`image/png;base64,…` — NO data:
    // scheme), and already-bare. A `startsWith("data:")` guard would pass the
    // middle one through with `image/png;base64,` still glued on, and the host
    // would get a corrupt payload. Splitting on the last comma handles all three.
    const scene = { width: 100, height: 100, nodes: [{ kind: "rect" as const, x: 0, y: 0, w: 4, h: 4, fill: "#111" }] };
    for (const [input, expected] of [
      ["data:image/png;base64,PAYLOAD", "PAYLOAD"],
      ["image/png;base64,PAYLOAD", "PAYLOAD"],
      ["PAYLOAD", "PAYLOAD"],
    ] as const) {
      const slide = makeSlide("s1");
      installHost([slide]);
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: input });
      expect(slide.created.filter((s) => !s.deleted)[0].imageBase64, input).toBe(expected);
    }
  });

  it("falls back to native shapes on a host without PowerPointApi 1.8", async () => {
    // setImage is 1.8 and the manifests admit hosts from 1.4, so this must be
    // GATED, not attempted-and-caught: a queued command the host rejects takes
    // the whole sync with it. The chart must still land, as shapes.
    const NODES = 12;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    // Everything except 1.8 — so grouping is off too, exactly like the web host.
    installHost([slide], [], slide, (v) => v !== "1.8");
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.length).toBeGreaterThanOrEqual(NODES); // the nodes, not a picture
    expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
    // Still re-editable: the config tag landed on the first shape (no group at 1.4).
    expect(live.some((s) => s.tagStore.get(CHART_TAG) === "{}")).toBe(true);
  });

  it("falls back to native shapes when the host refuses the picture fill", async () => {
    // A 1.8-advertising host that still rejects setImage (wrong payload format,
    // host quirk). renderPictureShape catches, best-effort deletes the rect, and
    // renderShapesChunked draws the nodes in the SAME request context.
    const NODES = 12;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refusePictureFill = true;
    try {
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
      const live = slide.created.filter((s) => !s.deleted);
      // The nodes landed, and no picture-filled shape survived.
      expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
      expect(live.filter((s) => s.geo === "rectangle").length).toBeGreaterThanOrEqual(NODES);
      expect(live.some((s) => s.tagStore.get(CHART_TAG) === "{}")).toBe(true);
    } finally {
      faults.refusePictureFill = false;
    }
  });

  it("refuses an over-budget payload rather than burning the batch timeout", async () => {
    // MAX_PICTURE_BASE64 is a guard against a pathological custom frame size,
    // not a measured host limit — real payloads are 20-133 KB. Crossing it
    // degrades to shapes; the code that says so is PC-IMG-TOOBIG on the console,
    // because the chart still appears and the fallback is otherwise invisible.
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: 12 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "A".repeat(4_000_001) });
      const live = slide.created.filter((s) => !s.deleted);
      expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
      expect(live.length).toBeGreaterThanOrEqual(12);
      // A specific, greppable code — so a real bug report teaches us the number.
      expect(warn.mock.calls.flat().join(" ")).toContain("PC-IMG-TOOBIG");
    } finally {
      warn.mockRestore();
    }
  });

  it("costs ONE sync for the drawing, however dense the scene", async () => {
    // The reliability claim: image mode replaces N/batchSize render syncs with
    // exactly one, so a chart that could never commit as shapes lands in a
    // single round trip. Compare a 60-node scene both ways.
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: 60 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    installHost([makeSlide("shapes")]);
    await insertSceneIntoSlide(scene, { tagData: "{}" });
    const shapesSyncs = trips.syncs;

    installHost([makeSlide("picture")]);
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
    const pictureSyncs = trips.syncs;

    // 60 nodes at SHAPES_PER_SYNC=10 is 6 render syncs; the picture is 1. Both
    // pay the same tag/group syncs afterwards, so the picture must be strictly
    // and substantially cheaper.
    expect(pictureSyncs).toBeLessThan(shapesSyncs);
    expect(shapesSyncs - pictureSyncs).toBeGreaterThanOrEqual(4);
  });

  it("re-fetches the slide's shape collection before addGroup on a multi-batch chart", async () => {
    // The real-host bug this guards: a >10-shape chart commits in multiple
    // batches, and the Shape proxies returned by earlier batches have their
    // object paths rewritten to getItem(id) by the time the group sync runs.
    // The web host silently drops addGroup(theseStaleProxies), leaving the
    // chart loose and unable to carry its POWERCHART_CONFIG tag. In the run
    // behind this fix, agenda / KPI+flow / table all landed ungrouped and
    // therefore un-re-editable. Fix: re-load slide.shapes.items right before
    // addGroup and pass those fresh proxies to addGroup.
    const NODES = 25; // 3 batches at SHAPES_PER_SYNC=10
    const bigScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.strictGroup = true; // web-host stale-proxy semantics — without the re-fetch, no group appears
    try {
      const report = await insertDemoDeck([{ scene: bigScene, tagData: '{"i":0}', title: "multi-batch" }]);
      expect(report.results[0].status).toBe("rendered");
      // Every appended chart is grouped — the new field flowed through.
      expect(report.results[0].grouped).toBe(true);
      // The slide holds ONE native group carrying the CHART_TAG — proving the
      // group survived and the tag landed on it (not on a stray first shape).
      const groups = deck[1].created.filter((s) => s.type === "group");
      expect(groups).toHaveLength(1);
      expect(groups[0].tagStore.get(CHART_TAG)).toBe('{"i":0}');
    } finally {
      faults.strictGroup = false;
    }
  });

  it("re-acquires each freshly-added slide per batch, so a rewritten getItemAt cannot 5010 mid-deck", async () => {
    // The real regression: HOLD one getItemAt handle to a new slide and reuse it
    // across the render's batched syncs, and once Office.js rewrites its path to
    // getItem(<web-non-round-trippable id>) the next shape throws "InvalidParam
    // passed to GetItem(id)", code 5010 — the deck dies partway through, as it did
    // on the real host. The fix re-acquires a fresh proxy each batch; the fake
    // window-limits a held one.
    //
    // Load-bearing: each slide must span MORE than one batch, because a held
    // handle only goes stale on the batch AFTER a sync. SHAPES_PER_SYNC is 10, so
    // a 25-node scene is 3 batches — a single-batch chart (e.g. cfgFor) can hold
    // its handle and never notice, which is exactly how a weaker version of this
    // test passed against the very bug it meant to guard.
    const NODES = 25;
    const bigScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 6;
    const report = await insertDemoDeck(Array.from({ length: n }, () => ({ scene: bigScene })));
    expect(failedIndices(report)).toEqual([]);
    // Every appended slide got all its shapes (plus the group) — nothing stranded
    // by a mid-batch 5010.
    expect(deck.length).toBe(1 + n);
    for (let i = 1; i <= n; i++) expect(deck[i].created.length, `slide ${i}`).toBeGreaterThanOrEqual(NODES);
  });
});

describe("a stalled host is legible, and does not hang the pane", () => {
  it("reports every phase, in order, with the shape count", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const seen: string[] = [];
    await insertSceneIntoSlide(buildChart(config), { tagData: "{}" }, (p, d) => seen.push(d ? `${p}:${d}` : p));
    expect(seen[0]).toBe("context");
    expect(seen.at(-1)).toBe("done");
    // "commit" now repeats — once per batch — because shapes land in batches.
    expect(seen.filter((s) => s.startsWith("commit:")).length).toBeGreaterThan(1);
    expect([...new Set(seen.map((s) => s.split(":")[0]))]).toEqual(["context", "queue", "commit", "group", "done"]);
    expect(seen.find((s) => s.startsWith("queue:"))).toMatch(/^queue:\d+ nodes$/);
    // Real progress: "10 of 40 shapes", ending at the total.
    const commits = seen.filter((s) => s.startsWith("commit:"));
    expect(commits[0]).toMatch(/^commit:\d+ of \d+ shapes$/);
    const [done, total] = commits
      .at(-1)!
      .match(/(\d+) of (\d+)/)!
      .slice(1);
    expect(done).toBe(total);
  });

  it("gives up on a host that never answers, naming the phase it died in", async () => {
    // The real failure mode: Office.js does not throw when the host stops
    // answering — the sync promise simply never settles, so the pane spins for
    // ever with nothing to report. This is the only way out.
    vi.useFakeTimers();
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      // A sync that never settles, exactly like a stalled PowerPoint.ashx.
      (slide as unknown as { id: string }).id = "s1";
      const ctxSync = () => new Promise<void>(() => {});
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: ctxSync,
          }),
      });
      const seen: string[] = [];
      const p = insertSceneIntoSlide(buildChart(config), {}, (ph) => seen.push(ph));
      const assertion = expect(p).rejects.toThrow(/did not respond while drawing shapes \d+-\d+ of \d+/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      // And it says where it stopped — "commit" is the last thing reached.
      expect(seen.at(-1)?.startsWith("commit")).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("a host that answers late still gets heard", () => {
  it("reports the real Office error when an abandoned sync finally rejects", async () => {
    // The evidence problem: racing a timeout throws the answer away. The
    // abandoned sync keeps running, and Office.js reports queued-command
    // failures THERE and nowhere else — so whatever it says next is the only
    // description of the bug we will ever get. Without this it is lost.
    vi.useFakeTimers();
    const heard: string[] = [];
    onLateSync((m) => heard.push(m));
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      let rejectSync!: (e: unknown) => void;
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: () => new Promise<void>((_, rej) => (rejectSync = rej)),
          }),
      });
      const p = insertSceneIntoSlide(buildChart(config), {});
      const assertion = expect(p).rejects.toThrow(/did not respond/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      expect(heard, "nothing heard before the host answers").toHaveLength(0);

      // Now the host finally answers — with a real RichApi-shaped error.
      rejectSync({
        message: "An internal error has occurred.",
        code: "GeneralException",
        debugInfo: { errorLocation: "Shape.name" },
      });
      await vi.advanceTimersByTimeAsync(1);
      expect(heard).toHaveLength(1);
      // The generic message alone is useless; code + debugInfo name the bug.
      expect(heard[0]).toContain("the host eventually FAILED");
      expect(heard[0]).toContain("code=GeneralException");
      expect(heard[0]).toContain("Shape.name");
    } finally {
      vi.useRealTimers();
    }
  });

  it("says so when the host was merely slow, not broken", async () => {
    vi.useFakeTimers();
    const heard: string[] = [];
    onLateSync((m) => heard.push(m));
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      let finish!: () => void;
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: () => new Promise<void>((res) => (finish = res)),
          }),
      });
      const p = insertSceneIntoSlide(buildChart(config), {});
      const assertion = expect(p).rejects.toThrow(/did not respond/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      finish();
      await vi.advanceTimersByTimeAsync(1);
      // "SUCCEEDED late" means the timeout is too short — a different bug from
      // a host that is actually broken, and the note has to distinguish them.
      expect(heard[0]).toContain("the host eventually SUCCEEDED");
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("added slides use the blank layout", () => {
  it("asks for the blank layout by TYPE, not by its localised name", async () => {
    // A slide added with no layout inherits the PREVIOUS slide's — on a fresh
    // deck that is the title slide, so an agenda lands on top of "Click to add
    // title" with the placeholder showing through. We draw everything
    // ourselves and want no placeholders.
    // The master here is Danish ("Tom"), which is the point: matching the name
    // "Blank" would silently do nothing for most of the world.
    installHost([makeSlide("s1")]);
    await insertAgendaSlides([
      buildAgendaScene(["Intro", "Body"], { highlight: 0 }),
      buildAgendaScene(["Intro", "Body"], { highlight: 1 }),
    ]);
    expect(addedWithLayout).toEqual(["layout-blank", "layout-blank"]);
  });

  it("names the layout's MASTER too, which is what a two-master deck requires", async () => {
    /**
     * THE WORST DEFECT THIS ADD-IN HAS SHIPPED, and it was a documented API
     * misuse sitting in plain sight.
     *
     * `AddSlideOptions`: a `layoutId` sent without a `slideMasterId` "needs to
     * be available for the default Slide Master ... Otherwise, an error will be
     * thrown" — and the default is THE PREVIOUS SLIDE'S master. We took the
     * blank layout of the FIRST master and sent it alone. On a one-master deck
     * that is always legal, which is why it worked for months and why this
     * fake — which had one master — could never catch it.
     *
     * On a two-master deck it is a server-side error. PowerPoint discards the
     * revision (`errorLocalChangeLostSingleUser`, the only ErrorName in 109
     * crash logs), resets undo history and rolls the deck back — so the slide
     * is listed, listed again, and then gone, and the draw's `getItemAt(index)`
     * lands past the end.
     *
     * Measured over 220 archived rounds, bucketed by the `layouts-readable`
     * probe on whether the round crashed:
     *
     *     1 master  /  1 layout   @ 16:9     9 of 182  =   5%
     *     2 masters / 12 layouts  @ 16:9     4 of   4  = 100%
     *     2 masters / 12 layouts  @ 4:3     27 of  30  =  90%
     *
     * Hold the deck and the aspect ratio changes nothing. Hold the ratio and
     * the deck moves it from 5% to 100%. It had been filed as "the 4:3 crash"
     * for two weeks; 4:3 was along for the ride because that arm has only ever
     * run on the two-master deck.
     */
    /**
     * installHost FIRST — it resets every fault, so one armed before it is
     * wiped. The first draft of this test armed `twoMasterDeck` beforehand and
     * therefore ran against ONE master, where the misuse is legal and the
     * assertion `["master-1"]` was true for the wrong reason. It passed while
     * testing nothing about two masters at all — the exact shape of vacuous
     * test this file keeps being caught by.
     */
    installHost([makeSlide("s1")]);
    faults.twoMasterDeck = true;
    try {
      const slideId = await addSlideForChart();
      expect(slideId, "the add was refused — the layout went without its master").toBeTruthy();
      expect(addedWithMaster, "the layout was sent without the master it belongs to").toEqual(["master-2"]);
      expect(addedWithLayout, "the layout did not come from the deck's own master").toEqual(["m2-layout-blank"]);
    } finally {
      faults.twoMasterDeck = false;
    }
    /**
     * A SURVIVING MUTANT, RECORDED. Disarming the fake's own enforcement of the
     * rule — the `throw` in `slides.add` — breaks no test here, because these
     * assertions read `addedWithMaster` directly rather than relying on the
     * host refusing. That is the right way round: the assertion IS the
     * contract, and it kills the mutants that matter (sending no master at all,
     * from either the first add or the retry).
     *
     * The fake's throw is there for every OTHER test in this file, so that any
     * future path which adds a slide on a two-master deck fails loudly instead
     * of quietly reproducing the defect. Nothing asserts that today.
     */
  });

  it("names the master on the RETRY add too, not just the first one", async () => {
    /**
     * `addSlides` issues its own `add()` again when the host swallows the first
     * — a real path, because PowerPoint web does drop adds under load. That is
     * what `faults.swallowAdds` models and what `slide add(s) never landed`
     * counts.
     *
     * The retry is a SEPARATE call with its own options object, and a mutation
     * run showed it surviving: dropping the master from the retry alone broke
     * no test. On a two-master deck that would make every recovery from a
     * swallowed add fail exactly the way the original defect did — rarer, and
     * so even harder to diagnose a second time.
     */
    // installHost FIRST. It resets fault state, so faults armed before it are
    // wiped — and arming them the other way round made this test exercise no
    // retry at all while still passing in isolation.
    installHost([makeSlide("s1")]);
    faults.twoMasterDeck = true;
    faults.swallowAdds = 1;
    try {
      const slideId = await addSlideForChart();
      expect(slideId, "the retry add was refused — it went without its master").toBeTruthy();
      expect(addedWithLayout.length, "the add was never retried, so this tested nothing").toBeGreaterThan(1);
      // `master-2` is the DECK's master — see the test above for why the layout
      // is taken from there rather than from whichever master comes first.
      expect(addedWithMaster.filter((m) => m === "master-2").length, "one of the adds went without its master").toBe(
        addedWithMaster.length,
      );
    } finally {
      faults.twoMasterDeck = false;
      faults.swallowAdds = 0;
    }
  });

  it("uses it for the demo deck too", async () => {
    installHost([makeSlide("s1")]);
    await insertDemoDeck([{ scene: buildChart(config), tagData: "{}" }, { scene: buildChart(config) }]);
    expect(addedWithLayout).toEqual(["layout-blank", "layout-blank"]);
  });

  it("still adds slides on a host that exposes no masters", async () => {
    // Layout choice is a nicety; inserting is not. If the host will not tell us
    // its layouts, fall back to the inherited one rather than failing.
    const ctx = installHost([makeSlide("s1")]);
    (ctx.presentation as unknown as { slideMasters: unknown }).slideMasters = {
      load() {},
      get items(): never {
        throw new Error("masters unavailable on this host");
      },
    };
    await insertAgendaSlides([buildAgendaScene(["Intro"], { highlight: 0 })]);
    expect(addedWithLayout).toEqual([undefined]);
  });
});

describe("the wait budget scales with the work", () => {
  /** Park the sync so we can watch the clock without the host ever answering. */
  const parkedHost = (slide: FakeSlide) =>
    vi.stubGlobal("PowerPoint", {
      ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
      run: async (cb: (ctx: unknown) => Promise<unknown>) =>
        cb({
          presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
          sync: () => new Promise<void>(() => {}),
        }),
    });

  /** A scene of `n` trivial shapes — the budget is a function of the count. */
  const sceneOf = (n: number) => ({
    width: 400,
    height: 300,
    nodes: Array.from({ length: n }, (_, i) => ({ kind: "rect" as const, x: i, y: 0, w: 4, h: 4, fill: "#111111" })),
  });

  it("never hands the host more than a batch at once — THE bug", async () => {
    // Measured against real PowerPoint on the web: ~10 shapes insert instantly,
    // the 18-shape table element works, a 30-shape butterfly NEVER commits —
    // the sync simply stops answering and nothing lands. The same shapes go
    // onto off-screen slides by the hundred, because those are not painted.
    // So the fix was never a bigger timeout; it was a smaller batch.
    const slide = makeSlide("s1");
    installHost([slide]);
    const perSync: number[] = [];
    let last = 0;
    const ctx = installHost([slide]);
    ctx.sync = async <T>(passThroughValue?: T) => {
      trips.syncs++;
      perSync.push(slide.created.length - last);
      last = slide.created.length;
      // Forwarded, like the real thing: PowerPoint.run carries a batch result
      // out through one last sync, and a stub that swallows it makes every run
      // in the file resolve undefined.
      return passThroughValue;
    };
    const scene = buildChart(config);
    expect(scene.nodes.length).toBeGreaterThan(10); // must actually span batches
    await insertSceneIntoSlide(scene, { tagData: "{}" });
    expect(Math.max(...perSync), `handed over at once: ${perSync.join(",")}`).toBeLessThanOrEqual(10);
  });

  // The batching above counts SHAPES. A node is not a shape: a wedge fans into
  // triangles and a polygon becomes one line per edge, so the kinds that flood
  // the host are exactly the ones the all-rect `stacked` config cannot exercise.
  it.each(["pie", "doughnut", "sunburst", "radar", "violin"])(
    "batches %s by shapes, not nodes — the wedge/polygon flood",
    async (kind) => {
      const slide = makeSlide(`s-${kind}`);
      const perSync: number[] = [];
      let last = 0;
      const ctx = installHost([slide]);
      ctx.sync = async <T>(passThroughValue?: T) => {
        trips.syncs++;
        perSync.push(slide.created.length - last);
        last = slide.created.length;
        return passThroughValue;
      };
      const scene = buildChart(sampleConfig(kind as never));
      await insertSceneIntoSlide(scene, { tagData: "{}" });
      // A single indivisible node (a wedge fan) may exceed the budget on its own;
      // nothing may exceed the host's measured breaking point of ~18.
      expect(Math.max(...perSync), `${kind} handed over at once: ${perSync.join(",")}`).toBeLessThanOrEqual(18);
    },
  );

  it("still bounds a trivial insert — the floor, not zero", async () => {
    vi.useFakeTimers();
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      parkedHost(slide);
      let settled = false;
      // 1 shape: the per-shape budget is tiny, so the 45s floor is what holds —
      // and 30s is past the old flat 20s, which is the thing that broke.
      const p = insertSceneIntoSlide(sceneOf(1), {}).catch(() => void (settled = true));
      await vi.advanceTimersByTimeAsync(30_000);
      expect(settled, "the floor keeps a small insert waiting past 30s").toBe(false);
      await vi.advanceTimersByTimeAsync(200_000);
      await p;
      expect(settled).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("EVERY insert path batches its shapes", () => {
  /** Max shapes handed to the host in any single sync of `run`. */
  async function maxPerSync(run: () => Promise<unknown>, slides: FakeSlide[]) {
    const worst = { n: 0 };
    let last = 0;
    const ctx = installHost(slides);
    // Wrap, don't replace: addSlides now verifies its adds via a settled
    // getCount() in a fresh context (see the addSlides retry/verify test),
    // which needs the real sync's committedCount/pendingCounts bookkeeping.
    // A bare replacement (as this used to be) freezes getCount() at its
    // initial value forever, which reads as "every add() was lost" and
    // starves insertAgendaSlides/insertDemoDeck of slides to render onto.
    const realSync = ctx.sync;
    const count = () => slides.reduce((a, s) => a + s.created.length, 0);
    ctx.sync = async <T>(passThroughValue?: T) => {
      worst.n = Math.max(worst.n, count() - last);
      last = count();
      // RETURNED, not merely awaited — the real `sync` resolves with what it
      // was handed, and `PowerPoint.run` carries the batch's result out that
      // way. A spy that awaits and returns nothing breaks every run it wraps.
      return await realSync(passThroughValue);
    };
    await run();
    return worst.n;
  }

  it("insert, update, agenda AND demo deck — none may send a whole scene", async () => {
    // The omission this exists for: I chunked insertSceneIntoSlide and
    // updateChartsInSlides and forgot insertAgendaSlides and insertDemoDeck.
    // The demo deck kept handing over ~200 shapes (4 slides at once) and sat at
    // "Working… 845s" having added nothing — and reported no progress, because
    // progress only fires when a chunk COMPLETES and the first never did.
    //
    // Live-canvas paths stay at ≤10 (repaints choke past that). Off-screen
    // append paths use a larger batch — the host tolerates far more when it
    // isn't repainting — but still bounded; they must NOT hand over the whole
    // scene, so a value at or under SHAPES_PER_SYNC_OFFSCREEN (40) is the
    // invariant, not the old flat 10.
    const scene = () => buildChart(config);
    expect(scene().nodes.length).toBeGreaterThan(10); // must span batches

    const s1 = makeSlide("s1");
    expect(
      await maxPerSync(() => insertSceneIntoSlide(scene(), { tagData: "{}" }), [s1]),
      "insert",
    ).toBeLessThanOrEqual(10);

    const s2 = makeSlide("s2");
    const old = s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    expect(
      await maxPerSync(
        () => updateChartInSlide(scene(), { slideId: "s2", shapeId: old.id, left: 0, top: 0 }, {}),
        [s2],
      ),
      "update",
    ).toBeLessThanOrEqual(11); // +1: the pre-existing shape this test planted

    const s3 = makeSlide("s3");
    expect(await maxPerSync(() => insertAgendaSlides([scene(), scene()]), [s3]), "agenda").toBeLessThanOrEqual(40);

    const s4 = makeSlide("s4");
    expect(
      await maxPerSync(
        () =>
          insertDemoDeck([
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
          ]),
        [s4],
      ),
      "demo deck",
    ).toBeLessThanOrEqual(40);
  });

  it("off-screen demo/agenda batches larger than the live canvas — cuts syncs per chart", async () => {
    // The live canvas caps at 10 shapes per batch (repaint mid-render kills the
    // host past that). Off-screen slides don't repaint, so demo/agenda push a
    // larger batch — 40 — cutting ~4x round-trips per chart. This asserts the
    // demo path ACTUALLY sends more than the live-canvas ceiling for a scene
    // that could fit in one 40-batch.
    const scene = () => ({
      width: 100,
      height: 100,
      nodes: Array.from({ length: 25 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    });
    installHost([makeSlide("s-live")]);
    // Live canvas: caps at 10.
    const live = await maxPerSync(() => insertSceneIntoSlide(scene(), { tagData: "{}" }), [makeSlide("s-live")]);
    expect(live, "live canvas ≤10").toBeLessThanOrEqual(10);
    // Off-screen demo: uses ≥15 in some batch — proves the flat 10 cap is gone.
    installHost([makeSlide("s-off")]);
    const off = await maxPerSync(() => insertDemoDeck([{ scene: scene() }]), [makeSlide("s-off")]);
    expect(off, "off-screen batches larger than live").toBeGreaterThan(10);
  });
});

describe("a target whose slide is gone is nothing to do, not a crash", () => {
  it("skips a stale slideId instead of throwing InvalidParam", async () => {
    // The real error, from the real host:
    //   InvalidParam passed to GetItem(id) | code=5010
    //   errorLocation: SlideCollection.getItem
    // An EditTarget outlives the slide it names — delete the slide, undo, or
    // reopen the deck and the id is stale. getItem THROWS on that; it is a
    // normal condition wearing a crash's clothes. Same Scale over a deck would
    // take one deleted chart and lose every OTHER chart's rescale with it.
    const live = makeSlide("s-live");
    installHost([live]);
    const s = live.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    await expect(
      updateChartsInSlides([
        {
          scene: buildChart(config),
          target: { slideId: "s-deleted", shapeId: "gone", left: 0, top: 0 },
          opts: { tagData: "{}" },
        },
        {
          scene: buildChart(config),
          target: { slideId: "s-live", shapeId: s.id, left: 10, top: 20 },
          opts: { tagData: '{"ok":1}' },
        },
      ]),
      // One refreshed target back: the live chart's. The dead one contributes
      // nothing, and must not take the live one down with it.
    ).resolves.toHaveLength(1);
    // The live chart still got drawn and tagged — one dead target must not take
    // the others down.
    const group = live.created.find((c) => c.type === "group");
    expect(group, "the live chart was skipped too").toBeTruthy();
    expect(group!.tagStore.get(CHART_TAG)).toBe('{"ok":1}');
  });

  it("does nothing at all when every target is stale", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const before = slide.created.length;
    await expect(
      updateChartsInSlides([
        { scene: buildChart(config), target: { slideId: "nope", shapeId: "nope", left: 0, top: 0 }, opts: {} },
      ]),
    ).resolves.toEqual([]);
    expect(slide.created.length).toBe(before);
  });
});

/**
 * The repair pass, against the host fake — `src/core/reconcile.ts` decides
 * what to do, this is the half that talks to PowerPoint and has to survive it.
 */
describe("charts too dense for the web to draw", () => {
  const web = { web: true, canPicture: true, alreadyPicture: false };

  it("rasterises a chart past the budget on the web", () => {
    // Violin is 253 native shapes; area 176; tile map 122. On the host with no
    // resource limits at all, those are the charts the budget is still for.
    //
    // The lower bound moved on 2026-09-06: the budget went 90 -> 105 because a
    // WAFFLE at 103 drew 35 times post-fix without taking the tab down. So 91
    // is deliberately no longer over the line, and 106 is what tests the rule.
    expect(wantsAutoPicture(253, web)).toBe(true);
    expect(wantsAutoPicture(122, web)).toBe(true);
    expect(wantsAutoPicture(106, web)).toBe(true);
    expect(wantsAutoPicture(103, web), "a waffle is measured to draw and must not be pictured").toBe(false);
  });

  it("leaves an ordinary chart alone", () => {
    // Gantt at 31 and Heatmap at 67 have drawn as shapes on the web all along.
    expect(wantsAutoPicture(31, web)).toBe(false);
    expect(wantsAutoPicture(67, web)).toBe(false);
    expect(wantsAutoPicture(90, web)).toBe(false);
  });

  it("never overrides the user's own choice of picture mode", () => {
    expect(wantsAutoPicture(253, { ...web, alreadyPicture: true })).toBe(false);
  });

  it("draws shapes on desktop however dense the chart is", () => {
    // Desktop has the resource limits the web lacks: Office throttles or
    // restarts the ADD-IN there rather than letting the client die, so the
    // native shapes the user actually wants stay the right answer.
    expect(wantsAutoPicture(253, { ...web, web: false })).toBe(false);
  });

  it("does not promise a picture a host cannot insert", () => {
    // Below PowerPointApi 1.8 there is no setImage to call.
    expect(wantsAutoPicture(253, { ...web, canPicture: false })).toBe(false);
  });
});

describe("reading a demo deck back and repairing it", () => {
  /** A slide as a damaged run leaves it: some shapes, maybe a banner, maybe a group. */
  function demoSlide(
    id: string,
    opts: {
      slot?: { i: number; title: string };
      shapes?: number;
      stamped?: boolean;
      tagged?: boolean;
      grouped?: boolean;
      /**
       * A degraded picture: ONE shape named PowerChart that is NOT a group.
       * The readback calls it `grouped` (it matches the name) but cannot count
       * children, so the slide comes back unmeasured — exactly the state a run
       * that fell back to pictures leaves behind.
       */
      picture?: boolean;
    },
  ): FakeSlide {
    const slide = makeSlide(id);
    if (opts.slot) slide.tags.add(DEMO_SLOT_TAG, JSON.stringify(opts.slot));
    if (opts.picture) {
      const pic = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 100, height: 100 });
      pic.name = "PowerChart";
      if (opts.tagged) pic.tags.add(CHART_TAG, `{"kind":"line"}`);
      return slide;
    }
    const parts: FakeShape[] = [];
    for (let i = 0; i < (opts.shapes ?? 0); i++) {
      const shape = slide.shapes.addTextBox(`n${i}`, { left: 0, top: 0, width: 10, height: 10 });
      shape.name = `part-${i}`;
      parts.push(shape);
    }
    if (opts.grouped && parts.length) {
      const group = slide.shapes.addGroup(parts);
      group.name = "PowerChart";
      if (opts.tagged) group.tags.add(CHART_TAG, `{"kind":"line"}`);
    } else if (opts.tagged && parts.length) {
      parts[parts.length - 1].tags.add(CHART_TAG, `{"kind":"line"}`);
    }
    if (opts.stamped) {
      const banner = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 100, height: 20 });
      banner.name = "PowerChart:not-complete";
    }
    return slide;
  }

  const expect3 = (slot: number, title: string, chart = true) => ({ slot, title, shapes: 3, chart });

  /** A two-point line chart — small enough that one batch draws the whole thing. */
  const tinyChart = (): ChartConfig => ({
    ...sampleConfig("line"),
    title: "Line",
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [1, 2] }] },
  });

  it("reads shape counts, banners, slot tags and group children off the deck", async () => {
    const deck = [
      demoSlide("s0", { shapes: 1 }),
      demoSlide("s1", { slot: { i: 0, title: "Line" }, shapes: 3, tagged: true }),
      demoSlide("s2", { slot: { i: 1, title: "Gantt" }, shapes: 3, stamped: true }),
      demoSlide("s3", { slot: { i: 2, title: "Pie" }, shapes: 3, grouped: true, tagged: true }),
    ];
    installHost(deck);
    const snaps = await snapshotAddedSlides(0, 4);
    expect(snaps.map((s) => s.slot)).toEqual([null, 0, 1, 2]);
    expect(snaps.map((s) => s.title)).toEqual([null, "Line", "Gantt", "Pie"]);
    expect(snaps[1]).toMatchObject({ shapes: 3, stamped: false, tagged: true });
    expect(snaps[2]).toMatchObject({ shapes: 4, stamped: true, tagged: false });
    // The group is one top-level shape; without its child count a complete
    // chart would read as 1 of 3 and be condemned as wreckage.
    expect(snaps[3]).toMatchObject({ grouped: true, groupChildren: 3, tagged: true });
  });

  /**
   * Pass C was the one pass that was not paged.
   *
   * It opened its own context with two syncs for EVERY grouped slide, which
   * made `READBACK_TIMEOUT_MS` meaningless in aggregate — the budget is per
   * slide, so sixty grouped slides is ninety minutes of rope with a full ninety
   * seconds left at every step. On a real 38-slide run it accounted for a
   * 49-second gap in the log, 43% of the post-insert wall clock, immediately
   * before the tab died.
   *
   * Counted in CONTEXTS rather than timed: the cost is round trips, and a
   * duration assertion on a fake host measures nothing.
   */
  it("counts group children a page at a time, not a context per slide", async () => {
    const n = READBACK_PAGE + 5;
    const deck = Array.from({ length: n }, (_, i) =>
      demoSlide(`s${i}`, { slot: { i, title: "Line" }, shapes: 3, grouped: true, tagged: true }),
    );
    installHost(deck);
    const before = trips.contexts;
    const snaps = await snapshotAddedSlides(0, n);
    // The measurement still happens — a pass that stopped answering would make
    // the repair treat every chart as unmeasured, which is worse than slow.
    expect(
      snaps.every((s) => s.groupChildren === 3),
      "lost the group-child counts",
    ).toBe(true);
    // Pass A + pass B + pass C, each paged. One context per grouped SLIDE would
    // be n on its own.
    expect(trips.contexts - before, "opened a context per grouped slide").toBeLessThan(n);
  });

  it("deletes a duplicate slide, clears a stale banner, and re-groups a loose chart", async () => {
    // The shape of Presentation_4.pptx: a clean chart, the same chart again
    // under a NOT COMPLETE banner, and an empty slide the host left behind.
    const deck = [
      demoSlide("title", { slot: { i: 0, title: "Title" }, shapes: 3, grouped: true }),
      demoSlide("line", { slot: { i: 1, title: "Line" }, shapes: 3 }),
      demoSlide("line-dup", { slot: { i: 1, title: "Line" }, shapes: 3, stamped: true }),
      demoSlide("stray", {}),
    ];
    installHost(deck);
    const outcome = await reconcileDeck(
      [expect3(0, "Title", false), expect3(1, "Line")],
      { before: 0, after: 4 },
      () => `{"kind":"line"}`,
      { dropOrphanBlanks: true },
    );
    expect(outcome.applied).toEqual({ unstamped: 0, regrouped: 1, deleted: 2 });
    expect(outcome.refused).toBe(0);
    expect(deck.map((s) => s.id)).toEqual(["title", "line"]);
    // The surviving chart is now a group carrying the config — re-editable.
    const group = deck[1].created.filter((s) => !s.deleted).find((s) => s.name === "PowerChart");
    expect(group?.tagStore.get(CHART_TAG)).toBe(`{"kind":"line"}`);
  });

  it("stops the readback when the user asks, and counts what it never looked at", async () => {
    // Stop used to reach exactly three places: a drawing batch boundary,
    // between demo items, and between charts in an update. The repair pass —
    // readback, tag pass, group count, the deletes — checked nothing. So a run
    // that got past drawing could not be cancelled at all: the pane switched
    // its button to "Stopping…" and the counter kept climbing. Observed at
    // 1819 seconds on a real host, and the run had to be abandoned by closing
    // the tab, which is also how its log was lost.
    //
    // What the pages it never reached are called matters as much as that it
    // stops. UNREAD, not clean: an unseen slide is never deleted by the pass
    // that follows, and "we did not look" must not read as "nothing there".
    const deck = Array.from({ length: 40 }, (_, i) =>
      demoSlide(`s${i}`, { slot: { i, title: `S${i}` }, shapes: 3, tagged: true }),
    );
    installHost(deck);
    requestStop();
    try {
      const { snapshots, unread } = await readAddedSlides(0, 40);
      expect(snapshots, "kept reading after a stop").toHaveLength(0);
      expect(unread, "a stopped readback reported slides as read").toBe(40);
    } finally {
      resetStop();
    }
  });

  it("gives up on a repair-pass page the host never answers", async () => {
    // 75 of this file's 79 syncs had no timeout. The four that did are all in
    // the insert path, so the drawing phase was bounded and everything after it
    // — group, tag, readback, repair — could wait forever on one unanswered
    // sync. That is the other half of the 1819-second run: even without a stop,
    // nothing would ever have broken the wait.
    const deck = [demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true })];
    installHost(deck);
    _setReadbackTimeoutForTest(5);
    stallSyncOn.add(1);
    try {
      const { snapshots, unread } = await readAddedSlides(0, 1);
      // Abandoned, and reported as unread — the same honest answer a page that
      // threw already gets. The call RETURNS, which is the whole point.
      expect(snapshots).toHaveLength(0);
      expect(unread).toBe(1);
    } finally {
      _setReadbackTimeoutForTest(90_000);
      stallSyncOn.clear();
    }
  });

  it("counts the slides a silent deck insert landed, instead of calling it a failure", async () => {
    // office-js#1650, verbatim: "the first time `context.sync()` is called the
    // promise resolves, but in subsequent calls the promise doesn't resolve,
    // although **the slide still gets added successfully**."
    //
    // Every timeout in this file used to throw. For a READ that is right — an
    // unread page is unread. For a WRITE it is wrong twice: it discards work
    // that landed, and it sends the caller off to do the work again, which is
    // how one stalled insert becomes two copies of a chart. This function
    // already measured the deck before and after; it simply never reached the
    // measurement on the runs that needed it.
    const built = await buildDeckBase64(
      [{ scene: buildChart(sampleConfig("clustered")), title: "A", configJson: "{}", slot: 0, run: "r1" }],
      { width: 720, height: 405 },
    );
    installHost([makeSlide("s1")]);
    faults.deckInsertNeverAnswers = true;
    // Generous on purpose, and this number is a FLOOR rather than a tuned
    // value. The insert's deadline fires either way — `deckInsertNeverAnswers`
    // hangs the sync forever — so nothing is weakened by giving it room. What a
    // tight budget did instead was fire during the round trip BEFORE the
    // insert, the one that reads the last slide's id to append rather than
    // front the deck: then the insert never ran, nothing landed, and the test
    // failed claiming work had been thrown away. At 40ms that lost a CI run and
    // reproduced locally 2 times in 5 while the machine was busy, and not at
    // all when it was idle — the signature of a race, not of a regression.
    _setBatchTimeoutForTest(500);
    _setDeckInsertPerSlideForTest(500);
    try {
      const landed = await insertSlidesFromPptx(built.base64, 1);
      expect(landed, "threw away a slide the host had already added").toBe(1);
    } finally {
      faults.deckInsertNeverAnswers = false;
      _setBatchTimeoutForTest(30_000);
      _setDeckInsertPerSlideForTest(5_000);
    }
  });

  it("never asks the host to clear the selection with an empty array", async () => {
    // office-js#3698, verbatim: `slide.setSelectedShapes([])` on the web "does
    // not clear the selection, causes the `PowerPoint.run` promise to never
    // resolve, and produces no error messages."
    //
    // So the call cannot succeed on the host it was written for, and is a
    // documented candidate for the silence this project measured. It is gone.
    // Re-selecting the SLIDE drops the shape selection on every host observed,
    // and was already there as the fallback — what is left is the half that
    // works.
    //
    // Asserted by counting the call rather than by inspecting its effect: the
    // effect of a call that never resolves is precisely what nobody can observe.
    installHost([makeSlide("s1")]);
    const before = trips.emptyDeselects;
    await clearShapeSelection("s1");
    expect(trips.emptyDeselects - before, "made the call office-js#3698 says never returns").toBe(0);
  });

  it("does not call a slide untagged when the host answered with no shapes", async () => {
    // The confirmed mechanism. A real 38-slide run had a readback page ask
    // about 19 slides carrying 19 shapes and see 3 — nothing threw, the
    // collection was simply short. Read naively that is "no shapes, so no
    // tags, so not re-editable", and the repair then rewrote 14 charts whose
    // config was already correct.
    const deck = [demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true })];
    installHost(deck);
    // Both the first pass and its re-read come back hollow.
    faults.hollowReads = 2;
    const { snapshots } = await readAddedSlides(0, 1);
    expect(snapshots[0].tagged).toBe(false); // could not see it…
    expect(snapshots[0].tagRead).toBe(false); // …and says so, which is the point
    faults.hollowReads = 0;
  });

  it("gets the right answer on the re-read when only the first look was hollow", async () => {
    const deck = [demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true })];
    installHost(deck);
    faults.hollowReads = 1; // first pass short, second fine
    const { snapshots } = await readAddedSlides(0, 1);
    expect(snapshots[0].tagged).toBe(true);
    expect(snapshots[0].tagRead).not.toBe(false);
    faults.hollowReads = 0;
  });

  it("traces what the tag pass asked for against what it got back", async () => {
    // The open question this exists to answer: a 39-slide run reported 20
    // tagged charts where the file provably carried 31, every miss in the
    // second page. Two candidates — collections coming back short, or
    // collections full but tag lookups resolving null — need different fixes
    // and are indistinguishable from the log as it stood. These numbers
    // separate them.
    setTracing(true);
    try {
      const deck = [
        demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true }),
        demoSlide("b", { slot: { i: 1, title: "B" }, shapes: 3 }),
      ];
      installHost(deck);
      await readAddedSlides(0, 2);
      const page = traceLog().entries.find((e) => e.message === "tag pass over a page");
      expect(page).toBeDefined();
      // Asked about both slides, saw every shape pass A counted, found the one
      // tag that is really there. A hollow read would show shapesSeen short of
      // shapesExpected — which is the signal the next real run has to produce.
      expect(page!.data).toMatchObject({ slides: 2, shapesExpected: 6, shapesSeen: 6, tagsFound: 1 });
    } finally {
      setTracing(false);
    }
  });

  it("counts the slides it could not read, instead of calling them lost", async () => {
    // A page whose read throws is skipped — safe, since an unseen slide is
    // never deleted. But its items then came back "lost", indistinguishable in
    // the run summary from slides the host really dropped. "Gantt: lost" when
    // Gantt rendered perfectly is exactly the report that makes someone insert
    // the deck a second time.
    // Two pages' worth, so one can be read and the other refused — a partial
    // read is the case that misreports; a total one is obvious.
    const deck = Array.from({ length: READBACK_PAGE + 5 }, (_, i) =>
      demoSlide(`s${i}`, { slot: { i, title: `T${i}` }, shapes: 3 }),
    );
    installHost(deck);
    failSyncsOn.add(trips.syncs + 2); // the SECOND page's read
    try {
      const { snapshots, unread } = await readAddedSlides(0, deck.length);
      expect(snapshots).toHaveLength(READBACK_PAGE);
      expect(unread).toBe(5);
    } finally {
      failSyncsOn.clear();
    }
  });

  it("writes the config tag onto a chart that is whole but untagged", async () => {
    // A degraded picture is one shape named PowerChart carrying no config.
    // Nothing to group — only the tag is missing — and until `retag` existed
    // there was no repair that could reach it.
    const deck = [demoSlide("pic", { slot: { i: 0, title: "Line" }, picture: true })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Line")], { before: 0, after: 1 }, () => `{"kind":"line"}`, {});
    expect(outcome.plan.actions.map((a) => a.kind)).toEqual(["retag"]);
    expect(outcome.applied.regrouped).toBe(1);
    expect(outcome.refused).toBe(0);
    // The chart object itself now carries the config — re-editable again.
    const chart = deck[0].created.filter((s) => !s.deleted).find((s) => s.name === "PowerChart");
    expect(chart?.tagStore.get(CHART_TAG)).toBe(`{"kind":"line"}`);
    // And nothing was grouped, because there was nothing to group.
    expect(deck[0].created.filter((s) => !s.deleted && s.type === "group")).toHaveLength(0);
  });

  it("never writes an origin tag when retagging — it cannot know where the chart was drawn", async () => {
    // The repair used to write `[caller.left, caller.top, shape.left,
    // shape.top]` from its DEFAULT origin of (60, 90). A generated deck
    // centres its charts, so on a real 38-slide run it rewrote 14 correct
    // origins of `[239.988, 120, 239.988, 120]` to `[60, 90, 239.988, 120]`.
    // Since an update renders at `origin + (live - anchor)`, every one of
    // those charts would have jumped ~180pt left on its first edit.
    const deck = [demoSlide("pic", { slot: { i: 0, title: "Line" }, picture: true })];
    installHost(deck);
    const chart = deck[0].created.find((s) => s.name === "PowerChart")!;
    chart.tagStore.set(CHART_ORIGIN_TAG, "[239.988,120,239.988,120]");
    const outcome = await reconcileDeck([expect3(0, "Line")], { before: 0, after: 1 }, () => `{"kind":"line"}`, {});
    expect(outcome.applied.regrouped).toBe(1);
    expect(chart.tagStore.get(CHART_TAG)).toBe(`{"kind":"line"}`); // the tag IS written
    expect(chart.tagStore.get(CHART_ORIGIN_TAG)).toBe("[239.988,120,239.988,120]"); // the origin is NOT
  });

  it("refuses the retag when there is no SSF Charts object to put it on", async () => {
    const deck = [demoSlide("bare", { slot: { i: 0, title: "Line" }, shapes: 1 })];
    installHost(deck);
    // Shapes present but loose and unnamed: this is a regroup, not a retag.
    const outcome = await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 1, chart: true }],
      { before: 0, after: 1 },
      () => `{"kind":"line"}`,
      {},
    );
    expect(outcome.plan.actions.map((a) => a.kind)).toEqual(["regroup"]);
  });

  it("refuses to delete a slide that is no longer the one it profiled", async () => {
    // A plan is a list of POSITIONS, decided from a readback taken several
    // round-trips ago. Nothing locks the deck in between: the pane leaves most
    // of its buttons live during a run, and PowerPoint's own UI is always
    // there. Deleting position 1 because position 1 used to be a duplicate is
    // how a repair pass destroys whatever occupies it now.
    const deck = [
      demoSlide("keep", { slot: { i: 0, title: "Line" }, shapes: 3 }),
      demoSlide("theirs", { shapes: 9 }), // NOT what the snapshot describes
    ];
    installHost(deck);
    // What the pass believed the deck held when it decided to delete index 1.
    const snapshots = [
      { index: 0, slot: 0, title: "Line", run: null, shapes: 3, stamped: false, tagged: false },
      { index: 1, slot: 0, title: "Line", run: null, shapes: 3, stamped: false, tagged: false },
    ];
    const plan = planReconcile(snapshots, [expect3(0, "Line", false)]);
    expect(plan.actions.filter((a) => a.kind === "delete").map((a) => a.index)).toEqual([1]);
    const outcome = await applyReconcilePlan(plan, () => undefined, { left: 0, top: 0 }, snapshots);
    expect(outcome.applied.deleted).toBe(0);
    expect(outcome.refused).toBe(1);
    // Both slides still there — the stranger above all.
    expect(deck.map((s) => s.id)).toEqual(["keep", "theirs"]);
  });

  it("pulls the banner off a chart that is in fact complete", async () => {
    const deck = [demoSlide("agenda", { slot: { i: 0, title: "Agenda" }, shapes: 3, stamped: true })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Agenda", false)], { before: 0, after: 1 }, () => undefined, {});
    expect(outcome.applied.unstamped).toBe(1);
    expect(deck[0].created.filter((s) => !s.deleted).some((s) => s.name === "PowerChart:not-complete")).toBe(false);
  });

  it("deletes from the end, so an earlier delete cannot renumber a later one", async () => {
    // Ascending deletes would remove index 1, shifting B's duplicate into the
    // slot the plan meant for something else — and take a good slide with it.
    const deck = [
      demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3 }),
      demoSlide("a-dup", { slot: { i: 0, title: "A" }, shapes: 3, stamped: true }),
      demoSlide("b", { slot: { i: 1, title: "B" }, shapes: 3 }),
      demoSlide("b-dup", { slot: { i: 1, title: "B" }, shapes: 3, stamped: true }),
    ];
    installHost(deck);
    const outcome = await reconcileDeck(
      [expect3(0, "A"), expect3(1, "B")],
      { before: 0, after: 4 },
      () => undefined,
      {},
    );
    expect(outcome.applied.deleted).toBe(2);
    expect(deck.map((s) => s.id)).toEqual(["a", "b"]);
  });

  it("leaves a slide from an unrelated deck untouched", async () => {
    const deck = [demoSlide("theirs", { shapes: 2 }), demoSlide("ours", { slot: { i: 0, title: "Line" }, shapes: 3 })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Line")], { before: 0, after: 2 }, () => undefined, {});
    expect(outcome.plan.orphans.map((o) => o.index)).toEqual([0]);
    expect(deck.map((s) => s.id)).toEqual(["theirs", "ours"]);
  });

  it("never groups the NOT COMPLETE banner in with the chart", async () => {
    // A real run shipped a Line chart whose group held 37 shapes: 36 of them
    // the chart, one a red NOT COMPLETE stripe. Once inside, the banner is
    // invisible to every later repair — a snapshot reads top-level names — so
    // it rides along with the chart forever.
    // 18 of 20 shapes: enough to re-group, not enough to count as complete,
    // so the plan asks for a regroup with the banner still on the slide.
    const deck = [demoSlide("partial", { slot: { i: 0, title: "Line" }, shapes: 18, stamped: true })];
    installHost(deck);
    await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 20, chart: true }],
      { before: 0, after: 1 },
      () => undefined,
      {},
    );
    const live = deck[0].created.filter((s) => !s.deleted);
    const group = live.find((s) => s.name === "PowerChart")!;
    expect(group.grouped).toHaveLength(18);
    expect((group.grouped as { name?: string }[]).some((c) => c.name === "PowerChart:not-complete")).toBe(false);
  });

  it("clears the banner on a slide whose chart is already grouped", async () => {
    // NOTE ON WHAT THIS DOES AND DOES NOT PROVE: in this fake a grouped shape
    // stays visible at the top level, so the repair finds the banner there.
    // On a real host a group hides its children, and the deck from the
    // 2026-07-31 run has a banner buried INSIDE a chart's group — `deleteStamp`
    // reaches into the group for exactly that case, and that reach is NOT
    // exercised here. `test/reconcile.test.ts` covers the planning half
    // (a grouped, stamped slide still gets an unstamp action).
    const deck = [demoSlide("buried", { slot: { i: 0, title: "Line" }, shapes: 3, stamped: true })];
    const shapes = deck[0].created;
    const group = deck[0].shapes.addGroup(shapes.slice());
    group.name = "PowerChart";
    installHost(deck);
    const outcome = await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 3, chart: true }],
      { before: 0, after: 1 },
      () => undefined,
      {},
    );
    expect(outcome.applied.unstamped).toBe(1);
    expect(deck[0].created.filter((s) => !s.deleted).some((s) => s.name === "PowerChart:not-complete")).toBe(false);
  });

  it("stamps nothing when the item's own slide never landed", async () => {
    // `stampLastSlide` used to brand whatever was last in the deck. When the
    // host swallowed the add, that was the PREVIOUS item's slide — a real run
    // defaced a KPI tile that had rendered perfectly, because a results page
    // that never landed stamped it.
    const existing = makeSlide("theirs");
    existing.shapes.addTextBox("someone else's work", { left: 0, top: 0, width: 10, height: 10 });
    const deck: FakeSlide[] = [existing];
    installHost(deck);
    faults.swallowAdds = ADDS_TO_DEFEAT_ONE_SLIDE; // the add and every retry vanish
    _setBatchTimeoutForTest(5); // no slide to draw on — do not wait 45s for it
    try {
      // Nothing rendered, so insertDemoDeck rethrows — and it rethrows the
      // REASON. With every add dropped, `addSlides` hands back no thunk at all;
      // the destructured `getSlide` was then called unchecked and the run died
      // with "getSlide is not a function", a TypeError from renderer internals
      // standing in for a diagnosis the code had already made. Assert the
      // message, not merely that something threw — a bare toThrow() passes just
      // as happily on the TypeError.
      await expect(insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }])).rejects.toThrow(
        /did not add a slide/i,
      );
      expect(existing.created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    } finally {
      _setBatchTimeoutForTest(45_000);
    }
    // Generous: an earlier test in this file can leave an abandoned sync
    // outstanding, and the failure path waits up to 5s for it to report.
  }, 20_000);

  it("stops drawing shapes and inserts pictures once the host falls behind", async () => {
    // We cannot catch the crash — the tab dies, there is no rejected promise —
    // so the run watches what comes BEFORE it. A budget of zero shapes stands
    // in for a host that has already spent its allowance.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const scene = buildChart(tinyChart());
    const report = await insertDemoDeck(
      [
        { scene, title: "One", tagData: `{"i":0}` },
        { scene, title: "Two", tagData: `{"i":1}` },
        { scene, title: "Three", tagData: `{"i":2}` },
      ],
      undefined,
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 0 },
    );
    expect(report.degradedAt).toBe(1);
    expect(report.degradeReason).toMatch(/shapes drawn/);
    // Item 1 drew its shapes; items 2 and 3 are one picture each.
    expect(report.results[0].created).toBeGreaterThan(1);
    expect(report.results[1].created).toBe(1);
    expect(report.results[2].created).toBe(1);
    // And the picture carries the config tag, so the chart is still editable.
    const last = deck[deck.length - 1].created.filter((s) => !s.deleted);
    expect(last.some((s) => s.imageBase64 === "iVBORw0KGgo=")).toBe(true);
  }, 20_000);

  it("reports per item whether the config tag actually committed", async () => {
    // Intent, recorded next to the settled readback's answer to the same
    // question. A real 38-slide run reported 20 tagged charts where the file
    // provably carried 31, and establishing that took unzipping the .pptx.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([
      { scene: buildChart(tinyChart()), title: "Tagged", tagData: `{"i":0}` },
      // No tagData: nothing to write, so `tagged` must be false — and that is
      // correct, not a fault. A reader has to be able to tell the two apart.
      { scene: buildChart(tinyChart()), title: "Untagged" },
    ]);
    expect(report.results.map((r) => !!r.tagged)).toEqual([true, false]);
  }, 20_000);

  it("reports tagged false when the host cannot write tags at all", async () => {
    // The case that produced 19 silently-untagged charts. `created` still
    // counts the shapes — they are on the slide — so only this field
    // distinguishes "drawn and editable" from "drawn and lost".
    const scene = buildChart(tinyChart());
    const item = { scene, title: "One", tagData: `{"i":0}` };

    // Control: a host that can tag reports true.
    installHost([makeSlide("ok")]);
    expect((await insertDemoDeck([item])).results[0].tagged).toBe(true);

    // A host below PowerPointApi 1.3 has no shape tags at all, so the chart is
    // drawn and is NOT re-editable. `created` still counts the shapes — they
    // are on the slide — so this field is the only thing that tells the two
    // states apart, which is exactly what a run of 19 silently-untagged charts
    // had no way to say.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck, [], undefined, (v) => v !== "1.3");
    const report = await insertDemoDeck([item]);
    expect(report.results[0].created).toBeGreaterThan(1);
    expect(report.results[0].tagged).toBe(false);
  }, 20_000);

  it("does not degrade a healthy run because somebody ELSE's stalled call answered", async () => {
    // The degrade signal was a global counter with no owner, so ANY abandoned
    // promise settling during a run bumped it. A chart the user edited before
    // the run started, whose stalled sync happened to answer during item 0,
    // degraded a perfectly healthy deck to rasters from item 1 — and reported
    // "the host answered after we gave up waiting" about an operation the run
    // never made. Ownership is captured when a call is ISSUED, which is the
    // only moment that distinguishes the two.
    const slide = makeSlide("s1");
    installHost([slide]);
    const realPowerPoint = (globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint;
    _setBatchTimeoutForTest(5);
    let finishForeign!: () => void;
    // A call issued OUTSIDE any run, on a host that never answers it.
    vi.stubGlobal("PowerPoint", {
      ...realPowerPoint,
      run: async (cb: (ctx: unknown) => Promise<unknown>) =>
        cb({
          presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
          sync: () => new Promise<void>((res) => (finishForeign = res)),
        }),
    });
    await expect(insertSceneIntoSlide(buildChart(tinyChart()), {})).rejects.toThrow(/did not respond/);
    const seqBefore = lastLateSyncSeq;
    // Back to a healthy host for the run itself.
    installHost([makeSlide("s2")]);
    _setBatchTimeoutForTest(45_000);
    let released = false;
    const report = await insertDemoDeck(
      Array.from({ length: 4 }, (_, i) => ({ scene: buildChart(tinyChart()), title: `C${i}`, tagData: `{"i":${i}}` })),
      () => {
        // Mid-run, exactly as it happened: the foreign call finally answers.
        if (!released) {
          released = true;
          finishForeign();
        }
      },
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 10_000, runId: "the-run" },
    );
    // Non-vacuity: the foreign call really did answer late, during the run.
    expect(lastLateSyncSeq).toBeGreaterThan(seqBefore);
    expect(lastLateSyncOwner).toBeNull();
    // And the run, which never stalled, drew shapes throughout.
    expect(report.degradeReason).toBeUndefined();
    expect(report.degradedAt).toBeUndefined();
    // Same ownership bug, one level down and unfixed for longer: `lateOutcome`
    // read the global late-sync string with no owner check at all, so the item
    // that happened to be drawing when the foreign call answered was recorded
    // as having stalled. Not one of these four gave up on anything.
    expect(report.results.map((r) => r.lateOutcome)).toEqual([undefined, undefined, undefined, undefined]);
    expect(report.results.some((r) => r.abandoned)).toBe(false);
  }, 20_000);

  it("attributes a stall to the run that issued it", async () => {
    // The other half: ownership must not cost the signal it exists to sharpen.
    // A call THIS run made, answering after it gave up, is exactly the
    // host-drowning evidence the degrade path was built for — and it must
    // still be recognised as the run's own.
    installHost([makeSlide("s1")]);
    _setBatchTimeoutForTest(5);
    try {
      // Stall whatever sync the run reaches next — which one is wrapped in
      // withTimeout is an implementation detail, and pinning an index here
      // would silently stop testing anything the day it moves.
      for (let k = 1; k <= 40; k++) stallSyncOn.add(trips.syncs + k);
      await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` }], undefined, {
        runId: "the-run",
      }).catch(() => {});
      stallSyncOn.clear();
      await waitForLateSync(500);
      expect(lastLateSyncOwner).toBe("the-run");
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncOn.clear();
    }
  }, 20_000);

  it("says which item gave up on a call, and pairs the late answer with it", async () => {
    // The other half of the ownership fix above. `abandoned` is the part that
    // is knowable the moment an item ends — a deadline fired inside it — and
    // it is the part a run log can always carry, because the host's eventual
    // answer routinely arrives minutes later, long after any wait a run can
    // afford. `lateOutcome` is then only ever set on an item that abandoned
    // something, so a healthy item can no longer inherit one.
    installHost([makeSlide("s1")]);
    _setBatchTimeoutForTest(20);
    try {
      // Stall everything item 0 reaches. The fake settles a stalled sync 40ms
      // in, well past the 20ms deadline, so item 0 is guaranteed to give up
      // and then be told how it went.
      for (let k = 1; k <= 40; k++) stallSyncOn.add(trips.syncs + k);
      const report = await insertDemoDeck(
        [
          { scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` },
          { scene: buildChart(tinyChart()), title: "Two", tagData: `{"i":1}` },
        ],
        (done) => {
          // Item 0 is over: let the host be healthy again, on a deadline no
          // ordinary sync can trip. Item 1 must show a clean sheet.
          if (done === 1) {
            stallSyncOn.clear();
            _setBatchTimeoutForTest(45_000);
          }
        },
        { runId: "the-run" },
      );
      // The bug, first: the late answer belongs to the item that gave up.
      expect(report.results[0].lateOutcome, "item 0 abandoned a call and was never told how it ended").toMatch(
        /SUCCEEDED/,
      );
      // …and it is not handed on to the next item, which stalled on nothing.
      expect(report.results[1].lateOutcome, "item 1 was credited with item 0's late answer").toBeUndefined();
      // Which item stalled, readable on its own — the half that is knowable
      // the moment the item ends, whether or not the host ever answers.
      expect(report.results[0].abandoned).toBe(true);
      expect(report.results[1].abandoned).toBe(false);
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncOn.clear();
    }
  }, 20_000);

  it("keeps a degraded picture re-editable on a host that rejects stale proxies", async () => {
    // The root cause of the shape path's lost tags, from a real 38-item run:
    //
    //   InvalidParam passed to GetItem(id) | code=5010
    //   errorLocation: ShapeCollection.getItem
    //   statement: var shape = shapes.getItem(...); var tags = shape.tags;
    //
    // 28 times. A degraded picture is ONE shape, so it was not "groupable",
    // so the refresh that exists for exactly this trap never ran on it — and
    // its proxy was created in one sync and tagged in another. The trap was
    // never about grouping; it is about using a proxy across a sync boundary,
    // and tagging does that too.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.strictTags = true;
    try {
      const report = await insertDemoDeck(
        [
          { scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` },
          { scene: buildChart(tinyChart()), title: "Two", tagData: `{"i":1}` },
        ],
        undefined,
        // Budget 0 degrades after the first item, so item 2 arrives as a picture.
        { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 0 },
      );
      expect(report.degradedAt).toBe(1);
      expect(report.results[1].created).toBe(1); // it IS the picture
      // …and it is re-editable, which is the whole claim the degrade makes.
      expect(report.results[1].tagged).toBe(true);
      const last = deck[deck.length - 1].created.filter((sh) => !sh.deleted);
      expect(last.some((sh) => sh.tagStore.get(CHART_TAG) === `{"i":1}`)).toBe(true);
    } finally {
      faults.strictTags = false;
    }
  }, 20_000);

  it("draws a too-dense chart as a picture once the run has degraded", async () => {
    // The density budget exists to stop a wedge/polygon flood timing the host
    // out. A picture is ONE shape — not a flood — so a run already drawing
    // pictures must not still be skipping its densest charts.
    //
    // It did. A real 38-item web run degraded at item 2 and then skipped and
    // stamped Area (176), Tile map (122), Waffle (103), Sunburst (101),
    // Violin (253) and Smoothed line (101), while the other thirty went on as
    // one-shape pictures in about a second each. Those six are exactly the
    // charts picture mode exists for, and they were the six it refused.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    // TILE MAP, NOT WAFFLE. This fixture used a waffle until 2026-09-06, when the
    // budget rose to 105 and a 103-shape waffle stopped being too dense — which
    // is the raise working, not the test breaking. Tile map is 122 at its sample
    // size and is one of the kinds the gate is still for.
    const dense = buildChart({ ...sampleConfig("tilemap"), title: "Tile map" });
    expect(estimateOfficeShapes(dense)).toBeGreaterThan(DEMO_SHAPE_BUDGET);
    const report = await insertDemoDeck(
      [
        { scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` },
        { scene: dense, title: "Tile map", tagData: `{"i":1}` },
      ],
      undefined,
      // Budget 0 degrades the run after the first item, so the dense chart is
      // reached with a picture available.
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 0 },
    );
    expect(report.degradedAt).toBe(1);
    // Drawn, not skipped — and as exactly one shape.
    expect(report.results[1].status).toBe("rendered");
    expect(report.results[1].created).toBe(1);
    const last = deck[deck.length - 1].created.filter((sh) => !sh.deleted);
    expect(last.some((sh) => sh.imageBase64 === "iVBORw0KGgo=")).toBe(true);
    expect(last.some((sh) => sh.name === "PowerChart:not-complete")).toBe(false);
  }, 20_000);

  it("still skips a too-dense chart when there is no picture to fall back on", async () => {
    // The budget must keep working on a host that offers no rasterizer —
    // otherwise the flood it exists to prevent goes straight through.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    // TILE MAP, NOT WAFFLE. This fixture used a waffle until 2026-09-06, when the
    // budget rose to 105 and a 103-shape waffle stopped being too dense — which
    // is the raise working, not the test breaking. Tile map is 122 at its sample
    // size and is one of the kinds the gate is still for.
    const dense = buildChart({ ...sampleConfig("tilemap"), title: "Tile map" });
    const report = await insertDemoDeck([{ scene: dense, title: "Tile map", tagData: `{"i":0}` }], undefined, {});
    expect(report.results[0].status).toBe("skipped");
  }, 20_000);

  it("keeps drawing shapes when the host is keeping up", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      [{ scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` }],
      undefined,
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 10_000 },
    );
    expect(report.degradedAt).toBeUndefined();
    expect(report.results[0].created).toBeGreaterThan(1);
  }, 20_000);

  it("never degrades when the caller offers no picture to fall back to", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "One" }], undefined, {
      shapeBudget: 0,
    });
    expect(report.degradedAt).toBeUndefined();
  }, 20_000);

  it("traces the run in enough detail to diagnose it afterwards", async () => {
    // Every hard thing in this project was diagnosed after the fact, from a
    // deck and a one-line summary. This is the record that was missing: which
    // item, what landed, what the host refused, and when it was given up on.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    setTracing(true);
    try {
      await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line", tagData: `{"i":0}` }]);
      const { entries } = traceLog();
      const byScope = (scope: string) => entries.filter((e) => e.scope === scope);
      // The per-item verdict, with the numbers a reader needs to judge it.
      const item = byScope("demo").find((e) => e.message === "item finished");
      expect(item?.data).toMatchObject({ i: 0, title: "Line", status: "rendered" });
      expect(item?.data?.created).toBeGreaterThan(1);
      // And the drawing itself, batch by batch — "died at batch 1 of 4" was
      // the entire diagnosis of the update stall.
      expect(byScope("draw").length).toBeGreaterThan(0);
      expect(byScope("draw")[0].data).toHaveProperty("total");
    } finally {
      setTracing(false);
    }
  }, 20_000);

  it("costs a run nothing while it is off", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    setTracing(false);
    // Switching off deliberately KEEPS the log readable, so measure growth
    // rather than emptiness — the claim is that a disabled trace records
    // nothing new, not that it forgets what it already had.
    const before = traceLog().entries.length;
    await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }]);
    expect(traceLog().entries).toHaveLength(before);
  }, 20_000);

  it("closes a demo run with the settled truth when asked for it", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      [{ scene: buildChart(tinyChart()), title: "Line", tagData: `{"i":0}` }],
      undefined,
      { reconcile: true },
    );
    expect(report.reconcile).toBeDefined();
    expect(report.reconcile?.plan.verdicts[0]).toMatchObject({ title: "Line", status: "rendered" });
    // A clean run needs no repair, and says so instead of inventing work.
    expect(report.reconcile?.plan.actions).toEqual([]);
    expect(report.blankSlides).toEqual([]);
  });

  it("does not reconcile unless the caller opts in", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }]);
    expect(report.reconcile).toBeUndefined();
  });
});

describe("stopping work in flight", () => {
  // The stop flag is module state, so a test that set it and threw would arm
  // every test after it. Cleared unconditionally.
  afterEach(() => resetStop());

  const cfgFor = (v: number): ChartConfig => ({
    ...config,
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [v, v + 1] }] },
  });
  const targetsOn = (slide: FakeSlide, n: number) =>
    Array.from({ length: n }, (_, i) => {
      const s = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
      return {
        scene: buildChart(cfgFor(i)),
        target: { slideId: slide.id, shapeId: s.id, left: 10, top: 20 },
        opts: { tagData: `{"i":${i}}` },
      };
    });

  it("stops at the next batch and keeps what already committed", async () => {
    // Office.js has no abort: a sync already handed to PowerPoint runs to
    // completion whatever we want. So "stop" means "queue nothing further",
    // and the batches that already landed stay on the slide — which is why
    // this throws rather than returning, so the caller cleans up rather than
    // grouping and tagging a half-drawn chart as a finished one.
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(cfgFor(0));
    let commits = 0;
    let thrown: unknown;
    try {
      await insertSceneIntoSlide(scene, { shapesPerSync: 5 }, (phase) => {
        if (phase === "commit" && ++commits === 1) requestStop();
      });
    } catch (err) {
      thrown = err;
    }
    expect(isStopped(thrown), "did not report a stop").toBe(true);
    // The first batch is on the slide; the rest was never queued.
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.length).toBeGreaterThan(0);
    expect(live.length).toBeLessThan(estimateOfficeShapes(scene));
    // And nothing half-finished was passed off as a chart.
    expect(live.some((s) => s.tagStore.has(CHART_TAG))).toBe(false);
  });

  /**
   * A chart with no config tag is on the slide and is not an SSF chart.
   *
   * `groupAndTagAll` has always answered this honestly — it returns `tagged` —
   * and the answer had nowhere to go: `EditTarget` carried no such field, so
   * the demo path (which has a repair pass) consumed it and the everyday insert
   * and in-place update (which have none) did not. The user gets "Done." in
   * green, clicking the chart says "the selection is not an SSF chart", and
   * reopening the deck loses the config for good. A real host produced this
   * four times in one run: *"a chart's tag could not even be queued"* followed
   * by *"tagging failed — charts are not re-editable until repaired"*.
   *
   * Two undefined `.tags` now, not one. A failure inside the drawing context is
   * no longer the end of the story — `settleAndTagChart` opens a fresh one and
   * writes the tag there, which is what makes that same real-host failure
   * survivable. So one is the case where the chart ends up re-editable after
   * all, and this test is about the case where it does not: the host refuses
   * the settled write too, and the caller has to be told the truth.
   */
  it("says so when the chart landed but its config did not", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.strictTags = true;
    // `.tags` undefined for the drawing context, the settle's BINDING write, its
    // by-id write, AND its collection-read fallback. Four writers, so refusing
    // "every tag write" now costs four — and the count going UP is the point of
    // the case below it.
    //
    // The binding joined them on 2026-08-16 and goes FIRST, because it is the
    // one route that needs neither an id nor a collection read — the two things
    // this host has been measured refusing. A chart that ends with no config
    // now has to survive all four being refused, which is what this asserts.
    faults.tagsUndefinedOn = 4;
    try {
      const target = await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
      // The chart IS there — this is not a failed insert, and reporting it as
      // one would send the user to draw a second copy.
      expect(slide.created.filter((s) => !s.deleted).length, "nothing was drawn").toBeGreaterThan(0);
      expect(target, "lost the target for a chart that is on the slide").toBeTruthy();
      expect(target!.lost, "reported a chart with no config as fully editable").toBe("no-config");
    } finally {
      faults.strictTags = false;
      faults.tagsUndefinedOn = 0;
    }
  });

  it("says nothing of the sort when the config DID land", async () => {
    // The negative control: a flag that is always set is not a signal.
    installHost([makeSlide("s1")]);
    const target = await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
    expect(target?.lost, "marked a perfectly good chart as lost").toBeUndefined();
  });

  /**
   * The host from 2026-08-07: it will not name a shape by id, and it will read
   * the collection.
   *
   * Sixty-six errors in that run log, every one `InvalidParam passed to
   * GetItem(id)` at `errorLocation: ShapeCollection.getItem`. Among them was
   * `settleAndTagChart`'s own write — a slide and a shape both resolved fresh
   * inside a first sync of their own, refused anyway. The settle then gave up,
   * on the reasoning that a collection search "would only find a DIFFERENT
   * shape to put this chart's config on", and five charts shipped with no
   * config: `same scale across the deck` reported *"3 of 8 charts carry the
   * shared scale ... the update reported 5×no-config"*.
   *
   * The reasoning is sound with no id and wrong with one. The read loads
   * `items/id`, so the caller's id picks its own shape out of the answer — no
   * guess is involved, and a chart that is not in the answer is simply not
   * tagged. And a collection read is what this host DOES honour: the repair
   * pass landed 23 retags that way in the same run that lost 46 tag writes.
   */
  it("settles the config through a collection read when the host will not name the shape by id", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    // Both halves of the real transcript, and both are needed.
    //
    // `refuseTagWrites = 1` is the run log's `tagging failed — charts are not
    // re-editable until repaired`: the drawing context's write goes, and the
    // settle is what has to save the chart. `refuseShapeById` is then the host
    // refusing the settle's by-id write too, which is the line after it.
    //
    // Arming only `refuseShapeById` is NOT enough, and it is worth saying why:
    // the drawing context tags an ungrouped chart through the proxy it created,
    // never by id, so the write lands, the settle never runs, and the case
    // passes against the unfixed file — a guard that proves nothing.
    faults.refuseTagWrites = 1;
    faults.refuseShapeById = true;
    try {
      const target = await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
      expect(slide.created.filter((s) => !s.deleted).length, "nothing was drawn").toBeGreaterThan(0);
      // The whole point: re-editable anyway. `lost` set here is the 5×no-config
      // verdict, reproduced.
      expect(target?.lost, "gave up on a chart the collection read could still have tagged").toBeUndefined();
      const tagged = slide.created.filter((s) => !s.deleted && s.tagStore.has(CHART_TAG));
      expect(tagged.length, "no shape on the slide carries the config").toBeGreaterThan(0);
    } finally {
      faults.refuseShapeById = false;
      faults.refuseTagWrites = 0;
    }
  });

  it("gives the two settle writes different names in the source", () => {
    // `settleAndTagChart` writes the tag twice by two different routes — once
    // by shape id, and once through a member of a collection re-read — and both
    // used to pass the SAME label to `boundedSync`. A refusal in a round log
    // then named a write without naming which write.
    //
    // Round 8 was decodable anyway, but only by reasoning across two traces:
    // the absence of `the host refused a settle by id` said the by-id branch
    // had been skipped, so the refusal had to be the collection read's. That is
    // an inference from a missing line, which is the reading this project gets
    // wrong most often. Distinct labels make it direct.
    //
    // Checked against the SOURCE because the labels only reach a log when a
    // write is refused, and arming a fake to refuse both halves in one run
    // exercises the fake's error plumbing rather than this property.
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    // The whole settle path: the by-id write, the collection re-read, and the
    // write through a member of that read. Sliced by function rather than
    // matched on wording, so renaming a label cannot quietly stop this looking.
    const from = src.indexOf("async function settleAndTagChart");
    const to = src.indexOf("async function settleUntaggedCharts");
    expect(from, "settleAndTagChart is gone — this guard is looking at nothing").toBeGreaterThan(0);
    expect(to, "settleUntaggedCharts is gone — the slice below has no end").toBeGreaterThan(from);
    const labels = [...src.slice(from, to).matchAll(/boundedSync\([^,]+,\s*[`"]([^`"]+)[`"]/g)].map((m) => m[1]);
    // FOUR since the name branch started checking that the chart it matched is
    // not already carrying someone else's config: the by-id write, the
    // collection re-read, that check, and the write through a member of the
    // read. The number is a tripwire for a sync added without thinking — raise
    // it deliberately, as here, or find out why one appeared.
    expect(labels.length, `expected four bounded syncs in the settle path, saw ${JSON.stringify(labels)}`).toBe(4);
    expect(new Set(labels).size, `two calls in the settle path share a label: ${JSON.stringify(labels)}`).toBe(
      labels.length,
    );
  });

  /** The settle pass's summary line, identified by its payload rather than its text. */
  const settleSummary = () =>
    traceLog().entries.find(
      (e) => e.scope === "group" && !!e.data && "charts" in e.data && "settled" in e.data && "lost" in e.data,
    );

  it("does not report a repair the settle did not make", async () => {
    // The 2026-08-08 round printed `settled the config tag the drawing context
    // could not write` five times, each carrying `settled: 0, lost: 1`. The
    // message named an outcome and the numbers underneath contradicted it, so
    // a reader scanning messages — which is how a 190-entry log is read — saw
    // five repairs that never happened.
    //
    // Every tag write refused, including the settle's own, so nothing can be
    // repaired and the line has to say so.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseTagWrites = 9999;
    faults.refuseShapeById = true;
    setTracing(true);
    try {
      await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' }).catch(() => {});
      // Found by the shape of its DATA, not by its wording. Several lines in
      // this scope mention the settle; only its summary carries these three
      // keys — and picking it by wording would mean the guard could not see
      // the message it exists to judge.
      const line = settleSummary();
      expect(line, "the settle pass left no line at all").toBeTruthy();
      // Non-vacuity: this really is the nothing-was-repaired case. Without it
      // the assertion below would pass on a run where the settle succeeded.
      const data = line?.data as { charts: number; settled: number; lost: number };
      expect(data.charts, "no chart reached the settle").toBeGreaterThan(0);
      expect(data.settled, "the settle was supposed to fail here").toBe(0);
      expect(line?.message, `the log claims a repair it did not make: "${line?.message}"`).toMatch(/could not repair/);
    } finally {
      setTracing(false);
      faults.refuseShapeById = false;
      faults.refuseTagWrites = 0;
    }
  });

  it("says the settle repaired the tag when it did", async () => {
    // The other side of the same mapping, so the message cannot simply be
    // pessimistic and pass the guard above forever.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseTagWrites = 1;
    faults.refuseShapeById = true;
    setTracing(true);
    try {
      await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
      const line = settleSummary();
      const data = line?.data as { settled: number; lost: number };
      expect(data.settled, "the settle was supposed to succeed here").toBeGreaterThan(0);
      expect(data.lost).toBe(0);
      expect(line?.message).toMatch(/repaired every/);
    } finally {
      setTracing(false);
      faults.refuseShapeById = false;
      faults.refuseTagWrites = 0;
    }
  });

  it("leaves every chart untouched when the stop lands before the first one", async () => {
    // updateChartsInSlides checks BEFORE each chart's delete, so a stop taken
    // there costs nothing: no shapes removed, nothing queued. The charts keep
    // their existing targets, which is what the caller needs to stay editable.
    const slide = makeSlide("s1");
    installHost([slide]);
    const items = targetsOn(slide, 3);
    const before = slide.created.filter((s) => !s.deleted).length;
    requestStop();
    const next = await updateChartsInSlides(items);
    // Every target comes back as it went in — none of them redrawn.
    expect(next).toEqual(items.map((it) => it.target));
    expect(slide.created.filter((s) => !s.deleted).length).toBe(before);
    expect(slide.created.some((s) => s.deleted)).toBe(false);
  });

  it("stops a deck run between items, keeping the slides already finished", async () => {
    // The longest thing the add-in does, and the reason a stop is worth having.
    // An item boundary costs nothing to stop at: every slide so far is complete,
    // grouped and tagged, and the next one has not been added.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      Array.from({ length: 4 }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
      (done) => {
        if (done === 1) requestStop();
      },
    );
    // One item ran; the run then stopped instead of drawing the other three.
    expect(report.results).toHaveLength(1);
    expect(report.results[0].status).toBe("rendered");
    // And the deck grew by exactly that one slide.
    expect(report.slidesAdded).toBe(1);
  });
});

describe("saying why a chart got no parts list", () => {
  it("reports an outcome at every exit, not just the one that succeeds", () => {
    // THE OBSERVABILITY GAP THIS CLOSES. The in-place chart update has never
    // once succeeded — 0 times against 1301 fallbacks across 117 archived
    // rounds — and the fallback's own reason is "the chart has no parts list",
    // 12 times out of 13, every round. But nothing recorded whether a parts
    // list was ever BUILT, so the archive could not tell apart:
    //
    //   never built    the chart was grouped, so it is not `loose` and no
    //                  siblings are collected for it at all
    //   built and lost the id read-back sync threw and the catch returns an
    //                  all-undefined list
    //   built, not found by the update path
    //
    // Three faults, three different fixes, and a day of archive mining could
    // not separate them.
    //
    // `ungroupedFallback` is not exported and needs a live PowerPoint context,
    // so this reads the source: a weaker check, and labelled as one. It catches
    // an exit losing its report — which is the specific way this gap reopens,
    // because two of the three exits are early returns that are easy to add to
    // and easy to forget.
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    const fn = src.slice(src.indexOf("async function ungroupedFallback"));
    const body = fn.slice(0, fn.indexOf("function tracePartsOutcome"));
    const returns = (body.match(/return partsJson;/g) ?? []).length;
    const reports = (body.match(/tracePartsOutcome\(/g) ?? []).length;
    expect(returns, "the function stopped returning a parts list at all").toBeGreaterThan(1);
    expect(reports, `${returns} exits but only ${reports} report — an exit lost its trace`).toBe(returns);
  });
});

describe("reading the presentation's slide size", () => {
  // Cached per deck, so a value from one test would answer for the next.
  beforeEach(() => _resetSlideSizeCache());
  afterEach(() => _resetSlideSizeCache());

  it("reads pageSetup directly when the host has 1.10", async () => {
    installHost([makeSlide("s1")]);
    hostSlideSize.cx = 9144000; // 4:3 — 720pt
    const size = await slideSize();
    expect(size).toEqual({ width: 720, height: 540, source: "pageSetup" });
  });

  it("falls back to exporting a slide when the host is below 1.10", async () => {
    // PageSetup arrived in 1.10. A 1.8 host still exports a slide as its own
    // .pptx, and that file declares the SOURCE deck's <p:sldSz> — exact, no
    // guessing, one rung down.
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.10");
    hostSlideSize.cx = 9144000;
    const size = await slideSize();
    expect(size).toEqual({ width: 720, height: 540, source: "exportedSlide" });
  });

  /**
   * RUNG 1 IS BOUNDED BY ITS OWN BUDGET, not by the selection one it borrowed.
   *
   * It timed out in 162 of the 163 archived rounds that read a slide size,
   * always for the full four seconds, and timing the rungs showed the bound was
   * not buying the answer: a SUCCESSFUL rung-1 read costs 246-504ms and rung 2's
   * export ~280ms, so the stall and the answer are different calls.
   *
   * The bound therefore came down — but into a constant of its own, because
   * `SELECTION_TIMEOUT_MS` also bounds five real selection and view calls that
   * nothing here has measured. This test is what makes that separation real
   * rather than a comment: it sets the two budgets far apart and checks which
   * one the ladder actually obeys, so wiring the call back to the selection
   * budget fails rather than passing quietly.
   *
   * The host is 1.10-but-not-1.8 on purpose. Rung 1 needs 1.10 and rung 2 needs
   * 1.8, so this leaves rung 1 as the ONLY sync in the path — a global wedge
   * then isolates it, and the ladder still returns (via the assumed floor)
   * instead of hanging on a rung this test is not about.
   */
  it("bounds rung 1 by the slide-size budget, not the selection one", async () => {
    const elapsed = async (): Promise<number> => {
      _resetSlideSizeCache();
      installHost([makeSlide("s1")], [], undefined, (v) => v === "1.10");
      faults.wedgeAfterSyncs = 0;
      const at = Date.now();
      try {
        const size = await slideSize({ refresh: true });
        // It must still ANSWER — the floor is the point of a ladder.
        expect(size.source, "a wedged rung 1 did not fall through to the floor").toBe("assumed");
      } finally {
        faults.wedgeAfterSyncs = null;
      }
      return Date.now() - at;
    };

    _setSelectionTimeoutForTest(600);
    _setSlideSizeTimeoutForTest(30);
    try {
      const short = await elapsed();
      _setSelectionTimeoutForTest(30);
      _setSlideSizeTimeoutForTest(600);
      const long = await elapsed();
      // Wired to the selection budget these two would come out the same way
      // round. The gap is the assertion.
      expect(short, `rung 1 took ${short}ms on a 30ms slide-size budget`).toBeLessThan(400);
      expect(long, `rung 1 took ${long}ms on a 600ms slide-size budget`).toBeGreaterThanOrEqual(500);
    } finally {
      _setSelectionTimeoutForTest(4_000);
      _setSlideSizeTimeoutForTest(1_500);
      _resetSlideSizeCache();
    }
  });

  it("assumes 16:9 — and says so — when no rung can answer", async () => {
    // The floor. A wrong-but-LABELLED width degrades placement to what it did
    // before; throwing here would take down an insert over a layout hint.
    installHost([makeSlide("s1")], [], undefined, () => false);
    const size = await slideSize();
    expect(size).toEqual({ width: 960, height: 540, source: "assumed" });
  });

  /**
   * The ladder must SAY how long it took, on whichever rung answers.
   *
   * The archive can count these traces and cannot time them, and that is the
   * whole reason the slide-size rung sat open. Over 270 rounds, every round
   * that reads a size times out on rung 1 first — 157 of 157, always the full
   * 4000ms — and rung 1 still supplies the answer in 82 of them. Whether the
   * bound can come down turns entirely on how long a SUCCESSFUL read takes,
   * and no trace carried it.
   *
   * Asserted on every rung, not just the first: the comparison the next sweep
   * makes is rung 1 against rung 2, so a duration on one of them measures
   * nothing.
   */
  it("says how long the read took, on whichever rung answers", async () => {
    const ladder: [string, () => void][] = [
      ["pageSetup", () => installHost([makeSlide("s1")])],
      ["exportedSlide", () => installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.10")],
      ["assumed", () => installHost([makeSlide("s1")], [], undefined, () => false)],
    ];
    for (const [rung, install] of ladder) {
      _resetSlideSizeCache();
      install();
      setTracing(true);
      try {
        const size = await slideSize();
        expect(size.source, `the ${rung} case did not reach its rung`).toBe(rung);
        const line = traceLog().entries.find((e) => /^slide size (read|unavailable)/.test(e.message));
        expect(line, `${rung} answered without tracing anything`).toBeTruthy();
        const ms = (line!.data as { ms?: unknown }).ms;
        expect(typeof ms, `${rung} traced no duration — the archive cannot time this rung`).toBe("number");
        expect(ms as number, `${rung} traced a negative duration`).toBeGreaterThanOrEqual(0);
      } finally {
        setTracing(false);
      }
    }
  });

  it("caches the answer, and re-reads on request", async () => {
    installHost([makeSlide("s1")]);
    expect((await slideSize()).width).toBe(960);
    // The user changes slide size from PowerPoint's own Design tab.
    hostSlideSize.cx = 9144000;
    expect((await slideSize()).width, "went back to the host for a cached value").toBe(960);
    expect((await slideSize({ refresh: true })).width).toBe(720);
  });

  it("does not let a fallback reading become the deck's permanent answer", async () => {
    // ROUND 115, REPRODUCED. The driver had just set the deck to 16:9 and
    // confirmed it twice against live `PageSetup`. The round then filed itself
    // as 720x540 from `documentFile` — the wrong profile, in the field every
    // later comparison groups by.
    //
    // The sequence: the first `slideSize()` after the resize caught the host
    // still busy, so rung 1 threw and rung 2 stalled; rung 3 read the SAVED
    // file, which PowerPoint had not yet written the new size into. That
    // fallback was cached, and one unlucky moment became the whole round's
    // answer even though the good rungs recovered seconds later.
    //
    // A fallback is what you take when the measurements are unavailable. It
    // must never outrank one that becomes available.
    installHost([makeSlide("s1")], [], undefined, () => false);
    hostSlideSize.cx = 12192000; // the deck really is 16:9 — 960pt
    const stale = await slideSize();
    expect(stale.source, "no rung above the floor could answer yet").toBe("assumed");

    // The host comes back. No `refresh: true`, because nothing in production
    // knows to pass it — that is precisely why round 115 filed wrongly.
    installHost([makeSlide("s1")]);
    const good = await slideSize();
    expect(good.source, "a fallback outranked a live measurement").toBe("pageSetup");
    expect(good.width, "kept the stale width").toBe(960);
  });

  it("keeps using a fallback rather than re-reading the whole deck for it", async () => {
    // The other half, and without it the fix above is a performance fault:
    // rung 3 copies the entire presentation in 4MB slices. Once the cheap rungs
    // have failed AGAIN, the answer already in hand is the right one to use.
    let fileReads = 0;
    installHost([makeSlide("s1")], [], undefined, () => false);
    const { default: JSZip } = await import("jszip");
    const zip = new JSZip();
    zip.file(
      "ppt/presentation.xml",
      `<p:presentation xmlns:p="x"><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    vi.stubGlobal("Office", {
      context: {
        host: "PowerPoint",
        requirements: { isSetSupported: () => false },
        document: {
          getFileAsync: (_t: unknown, _o: unknown, cb: (r: unknown) => void) => {
            fileReads++;
            cb({
              status: "succeeded",
              value: {
                size: bytes.length,
                sliceCount: 1,
                getSliceAsync: (_i: number, scb: (r: unknown) => void) =>
                  scb({ status: "succeeded", value: { data: bytes } }),
                closeAsync: () => {},
              },
            });
          },
        },
      },
      FileType: { Compressed: "compressed" },
    });
    expect((await slideSize()).source).toBe("documentFile");
    expect(fileReads, "the expensive rung ran once").toBe(1);
    expect((await slideSize()).width, "lost the fallback it already had").toBe(720);
    expect(fileReads, "re-read the entire deck for an answer it already held").toBe(1);
    vi.unstubAllGlobals();
  });

  it("loads pageSetup before reading it", async () => {
    // The bug class this repo keeps finding: a proxy resolved and never loaded
    // answers on the fake and throws PropertyNotLoaded on the host. The fake is
    // honest now, so forgetting the load falls through to the next rung — and
    // the source is how the test can tell that happened.
    installHost([makeSlide("s1")]);
    expect((await slideSize()).source).toBe("pageSetup");
  });

  it("pulls the whole document through the Common API when PowerPointApi cannot answer", async () => {
    // The only rung a 1.4-1.7 host has. `getFileAsync` predates the
    // PowerPointApi requirement sets entirely, so it answers where nothing else
    // does — at the cost of copying the entire deck in 4MB slices to reach two
    // numbers in its first part. Hence last, and hence still worth having.
    installHost([makeSlide("s1")], [], undefined, () => false);
    const { default: JSZip } = await import("jszip");
    const zip = new JSZip();
    zip.file(
      "ppt/presentation.xml",
      `<p:presentation xmlns:p="x"><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    // Two slices, so the reassembly is actually exercised rather than assumed.
    const cut = Math.floor(bytes.length / 2);
    const slices = [bytes.slice(0, cut), bytes.slice(cut)];
    let closed = false;
    vi.stubGlobal("Office", {
      context: {
        host: "PowerPoint",
        requirements: { isSetSupported: () => false },
        document: {
          getFileAsync: (_t: unknown, _o: unknown, cb: (r: unknown) => void) =>
            cb({
              status: "succeeded",
              value: {
                size: bytes.length,
                sliceCount: slices.length,
                getSliceAsync: (i: number, scb: (r: unknown) => void) =>
                  scb({ status: "succeeded", value: { data: slices[i] } }),
                closeAsync: () => {
                  closed = true;
                },
              },
            }),
        },
      },
      FileType: { Compressed: "compressed" },
    });
    setTracing(true);
    // THE INDEX BEFORE, and the search starts from it. `traceLog()` accumulates
    // across this file, so `find` returned the FIRST "slide size read" in the
    // run — a line another test wrote, with a duration this rung never traced.
    // The assertion passed against a deliberately broken rung 3 because of it,
    // which is how it was caught: three of four mutants died and this one would
    // not, however the trace site was mutated.
    const before = traceLog().entries.length;
    let size;
    try {
      size = await slideSize();
      // The fourth rung's duration, asserted HERE rather than in the block
      // above, because reaching this rung costs a two-slice zip and a stubbed
      // Common API and this test already builds both.
      const line = traceLog()
        .entries.slice(before)
        .find((e) => e.message === "slide size read");
      expect(line, "the deepest rung answered without tracing at all").toBeTruthy();
      expect(typeof (line!.data as { ms?: unknown }).ms, "the deepest rung traced no duration").toBe("number");
    } finally {
      setTracing(false);
    }
    expect(size).toEqual({ width: 720, height: 540, source: "documentFile" });
    // The handle MUST be released: a leaked one holds the host's copy of the
    // document alive and can block later getFileAsync calls outright.
    expect(closed, "leaked the document file handle").toBe(true);
  });
});

/**
 * The LIVE renderer against a config nobody validated.
 *
 * Every hostile-input sweep this project has ran against the SVG renderer and
 * the pptx one. Neither is the one that runs in a real PowerPoint, and the
 * Office renderer turned out to hold the third independent copy of the same
 * hole: `officeHex` did `color.trim()`, so `style.palette: [1, 2, 3]` threw
 * `color.trim is not a function` and took down a live insert — on the path a
 * user is actually standing on, for a config that came out of the JSON box or
 * a `POWERCHART_CONFIG` tag written in another deck.
 *
 * Three sinks, three separate holes, found one at a time because each sweep
 * only knew about the renderer it was written for.
 */
describe("a hostile config cannot take down a live insert", () => {
  const HOSTILE_CONFIGS: [string, (c: ChartConfig) => unknown][] = [
    ["numeric title", (c) => ({ ...c, title: 2024 })],
    ["numeric categories", (c) => ({ ...c, data: { ...c.data, categories: [1, 2, 3] } })],
    [
      "numeric series name",
      (c) => ({ ...c, data: { ...c.data, series: c.data.series.map((s) => ({ ...s, name: 7 })) } }),
    ],
    ["numeric palette", (c) => ({ ...c, style: { palette: [1, 2, 3] } })],
    // The fourth copy of the same hole, found the same way and in all three
    // sinks at once. `officeHex` hands anything non-alphabetic to `toHex6`, and
    // `toHex6` used to CRASH on a colour whose numbers are not numbers: the
    // regex that finds them matches a bare ".", `parseFloat(".")` is NaN, and
    // the hue sector table has no NaN entry. A malformed colour is exactly what
    // a hand-edited config or a template written in another deck arrives with.
    ["a palette colour whose numbers are not numbers", (c) => ({ ...c, style: { palette: ["hsl(., 50%, 50%)"] } })],
    ["labelContent a bare string", (c) => ({ ...c, decorations: { ...c.decorations, labelContent: "value" } })],
    ["numberFormat null", (c) => ({ ...c, numberFormat: null })],
    ["numeric valueAxisTitle", (c) => ({ ...c, valueAxisTitle: 5 })],
  ];

  for (const [name, mutate] of HOSTILE_CONFIGS) {
    it(
      name,
      async () => {
        const bad: string[] = [];
        for (const { kind } of CHART_KINDS.slice(0, 12)) {
          installHost([makeSlide("s1")]);
          try {
            await insertSceneIntoSlide(buildChart(mutate(sampleConfig(kind)) as ChartConfig), { tagData: "{}" });
          } catch (e) {
            bad.push(`${kind}: ${e instanceof Error ? e.message : String(e)}`);
          }
          vi.unstubAllGlobals();
        }
        expect(bad.slice(0, 4)).toEqual([]);
      },
      30_000,
    );
  }
});

describe("what a deck scan says about itself", () => {
  /**
   * It used to say nothing at all unless it came back short.
   *
   * The battery scans the deck about a dozen times a round — once per scenario,
   * through `probeCharts` — and every one of those was invisible. Round 10 is
   * what that cost: `stop a run part-way` reported 39.4 seconds against 2.6-3.2s
   * in the eight rounds before it, and the log had a 39-second HOLE where the
   * scan was. The stop itself was instant; the verification after it was not,
   * and nothing in the file said so.
   *
   * It is also the one operation the quadratic per-slide cost predicts should
   * grow worst — it reads every slide's shapes, on a deck the battery keeps
   * adding to — and it was the one operation never measured.
   */
  it("scans a deck DEEPER than one page, and says so when a page is lost", async () => {
    /**
     * THE PAGED LOOP IN THE DECK SCAN HAD NEVER RUN TWICE — NOT EVEN HERE.
     *
     * `listChartsInDeck` reads the deck `READBACK_PAGE` slides at a time and is
     * what Same Scale, the repair pass and five self-test scenarios all go
     * through. It appears in 44 tests in this file and every one of them
     * installed a deck of a handful of slides, so `start += READBACK_PAGE` never
     * came round a second time. On a real host it never has either: the deepest
     * deck in 322 archived rounds is NINE slides, and 305 of them scanned seven.
     *
     * The other paged readers — `snapshotAddedSlides`, `readAddedSlides` — are
     * covered past the boundary, including a refused second page. The one a USER
     * reaches was not. Anyone pressing Same scale on an ordinary 40-slide deck
     * is the first to run it.
     *
     * Two halves, and the second is the one that matters. A scan that loses a
     * page must SAY it lost one: if it silently returns page 1, Same Scale
     * rescales a subset of the deck and reports success, which is the exact
     * shape of "no charts here" being indistinguishable from "I could not look".
     */
    const n = READBACK_PAGE + 5;
    const deck = Array.from({ length: n }, (_, i) => makeSlide(`s${i}`));
    // A named shape on the LAST slide — the one only a second page can reach.
    const far = deck[n - 1].shapes.addGeometricShape("rectangle", { left: 7, top: 8, width: 9, height: 10 });
    far.name = "beyond the first page";
    /**
     * CHARTS ON BOTH PAGES, with DIFFERENT counts per page.
     *
     * Without them every page contributes zero, and a per-page count is
     * indistinguishable from the running total — a mutation swapping one for
     * the other survived, because 0 === 0. Three here and two beyond the
     * boundary make the two numbers differ on the second page, which is the
     * whole point of reporting both.
     */
    for (const i of [0, 1, 2, READBACK_PAGE, READBACK_PAGE + 1]) {
      const c = deck[i].shapes.addGeometricShape("rectangle", { left: 1, top: 2, width: 3, height: 4 });
      c.tagStore.set(CHART_TAG, '{"kind":"pie"}');
    }
    installHost(deck);

    setTracing(true);
    const whole = await listChartsInDeck({ withInventory: true });
    expect(whole.unread, "a deck of 25 did not scan cleanly").toBe(0);
    expect(scanIsComplete(whole), "a complete scan did not report itself complete").toBe(true);
    expect(whole.inventory, "the inventory stopped at the page boundary").toHaveLength(n);
    // The slides BEYOND the first page are the ones this test exists for.
    expect(whole.inventory?.[n - 1]?.slideId, "the last slide of the second page went missing").toBe(`s${n - 1}`);
    expect(
      whole.inventory?.[n - 1]?.shapes.map((s) => s.name),
      "a shape on the second page never reached the inventory",
    ).toContain("beyond the first page");

    /**
     * AND EACH PAGE SAYS WHAT IT CONTRIBUTED.
     *
     * The loop merged `readChartsPage`'s findings in silence, so a page that
     * ran and contributed nothing looked exactly like a page of slides holding
     * no charts — which is the failure a paging loop actually has, and the one
     * thing the instrument could not see. `chartsSoFar` is the running total on
     * purpose: it shows the MERGE happening rather than merely the return.
     */
    const pages = traceLog().entries.filter((e) => e.message === "deck scan — a page came back");
    expect(pages.length, "the second page never reported back").toBe(Math.ceil(n / READBACK_PAGE));
    expect(pages[0]?.data).toMatchObject({ from: 0, to: READBACK_PAGE - 1, charts: 3, chartsSoFar: 3 });
    // THE TWO NUMBERS DIFFER HERE, and that is the assertion. `charts` is what
    // this page found; `chartsSoFar` is what has survived the merge. Reporting
    // the running total in both places would hide a page whose findings were
    // dropped, which is the failure this instrument exists for.
    expect(pages[1]?.data).toMatchObject({ from: READBACK_PAGE, to: n - 1, charts: 2, chartsSoFar: 5 });
    setTracing(false);

    // Now lose the second page. The count comes back, the charts on it do not,
    // and the scan has to be honest about which.
    setTracing(true);
    const before = traceLog().entries.length;
    failSyncsOn.add(trips.syncs + 2);
    try {
      const partial = await listChartsInDeck({ withInventory: true });
      expect(partial.unread, "a lost page was reported as nothing to read").toBeGreaterThan(0);
      expect(scanIsComplete(partial), "a scan that lost a page called itself complete").toBe(false);
      /**
       * AND IT NAMES THE PAGE IT LOST, AND WHY.
       *
       * This catch swallowed the error and added the page to `unread`, so a
       * failed page reached the reader as a bigger number and nothing else —
       * not which page, not why, not even that one had failed rather than the
       * deck being unreadable. On the multi-page scan this loop exists for,
       * that was the whole diagnosis.
       *
       * Sliced from `before` so this reads only THIS scan's entries: the log
       * accumulates across the two halves of this test, and asserting over all
       * of it would let the clean half satisfy the broken half's assertions.
       */
      const lost = traceLog()
        .entries.slice(before)
        .filter((e) => e.message === "deck scan — a page did not come back");
      expect(lost.length, "a page was lost and nothing said so").toBeGreaterThan(0);
      expect(typeof lost[0]?.data?.from, "the lost page was not identified").toBe("number");
      // The host's own words, not a boolean. Measured: "host refused a queued
      // command | at=reading slides 0-19 for charts" — which names the sync as
      // well as the failure, and is what a crashed round would carry.
      expect(String(lost[0]?.data?.error ?? ""), "the reason was discarded").toMatch(/refused|at=reading slides/);
    } finally {
      failSyncsOn.clear();
      setTracing(false);
    }
  });

  it("reports a CLEAN scan, not only a short one", async () => {
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    try {
      const scan = await listChartsInDeck();
      expect(scan.unread, "this fixture was supposed to scan cleanly").toBe(0);
      const line = traceLog().entries.find((e) => e.message === "scanned the deck for charts");
      expect(line, "a clean scan left no trace, so a slow one is a hole in the log").toBeDefined();
      expect(line?.data).toMatchObject({ slides: 2, unread: 0, complete: true });
      expect(typeof line?.data?.ms, "no duration, which is the number the hole was hiding").toBe("number");
    } finally {
      setTracing(false);
    }
  });

  it("still says a short scan was short, in the same line", async () => {
    // The negative control. Moving the trace out of the failure branch must not
    // cost the failure branch its report — a scan that read nothing and one
    // that read everything have to stay distinguishable.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    faults.failSyncOn = 2;
    try {
      await listChartsInDeck();
      const line = traceLog().entries.find((e) => e.message === "scanned the deck for charts");
      expect(line?.data?.complete, "a scan that could not read a page reported itself complete").toBe(false);
      expect(line?.data?.unread, "the unread count went missing with the old message").toBeGreaterThan(0);
    } finally {
      faults.failSyncOn = 0;
      setTracing(false);
    }
  });
});

describe("the pictures a diagnostic round sends back", () => {
  /**
   * Three different reasons a slide comes back without a picture, and for a
   * long time one message for all of them.
   *
   * Every real round said `slides the host would not draw {asked: 22, drew: 12,
   * max: 12}` — which reads as a host refusing ten slides, and the host had
   * refused nothing at all: ten slides were over OUR cap and never asked about.
   * That is the same "never asked looks like answered no" mistake the contract
   * gate used to make, in the line a reader reaches for first when a deck comes
   * back short.
   */
  it("blames the cap for slides it never asked about, not the host", async () => {
    installHost([makeSlide("s1"), makeSlide("s2"), makeSlide("s3")]);
    setTracing(true);
    try {
      const shots = await slideShots(["s1", "s2", "s3"], { max: 1 });
      // Every id still comes back — a capped run that showed the first one only
      // would read as a deck of one.
      expect(shots.map((s) => s.slideId)).toEqual(["s1", "s2", "s3"]);
      expect(shots.filter((s) => s.png)).toHaveLength(1);
      const said = traceLog().entries.filter((e) => e.scope === "host");
      const refusal = said.find((e) => e.message === "slides the host would not draw");
      expect(
        refusal,
        `the host refused nothing and was blamed anyway: ${JSON.stringify(refusal?.data)}`,
      ).toBeUndefined();
      const capped = said.find((e) => e.message === "slides never asked about");
      expect(capped?.data).toMatchObject({ asked: 3, drew: 1, overCap: 2, stopped: 0, max: 1 });
    } finally {
      setTracing(false);
    }
  });

  it("still blames the host when the host is the one that would not draw", async () => {
    // The negative control. A message that never fires is not an improvement on
    // one that fires wrongly — a real refusal has to keep reaching the log.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    faults.emptySlideImage = true;
    try {
      const shots = await slideShots(["s1", "s2"], { max: 12 });
      expect(shots.filter((s) => s.png)).toHaveLength(0);
      const said = traceLog().entries.filter((e) => e.scope === "host");
      expect(said.find((e) => e.message === "slides the host would not draw")?.data).toMatchObject({
        asked: 2,
        drew: 0,
        refused: 2,
      });
      expect(
        said.find((e) => e.message === "slides never asked about"),
        "nothing was capped or stopped",
      ).toBeUndefined();
    } finally {
      faults.emptySlideImage = false;
      setTracing(false);
    }
  });

  it("counts a stop apart from the cap, so a short deck says which", async () => {
    // The third reason, and the one most likely to be misread: a run the user
    // stopped comes back with the same shape of gap as a capped one.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    requestStop();
    try {
      const shots = await slideShots(["s1", "s2"], { max: 12 });
      expect(shots.filter((s) => s.png)).toHaveLength(0);
      const capped = traceLog().entries.find((e) => e.message === "slides never asked about");
      expect(capped?.data).toMatchObject({ asked: 2, drew: 0, overCap: 0, stopped: 2 });
    } finally {
      resetStop();
      setTracing(false);
    }
  });
});

/**
 * `officeHex` handed the host any run of LETTERS verbatim, on the reasoning
 * that Office.js knows the CSS names. It knows the names; it does not know
 * every word. `style.palette: ["banana"]` — or a series colour of
 * `constructor`, which the pane's own saved-template store made reachable — is
 * a run of letters and is not a colour, and `setSolidColor` on a name the host
 * does not know is rejected inside the draw batch. One bad word did not degrade
 * one shape; it took the batch, and with it the chart.
 */
describe("only colour names the host actually knows reach the host", () => {
  const drawWith = async (color: string) => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide({
      width: 100,
      height: 60,
      nodes: [{ kind: "rect", x: 0, y: 0, w: 40, h: 20, fill: color, name: "seg-0-0" }],
    });
    return slide.created.map((s) => s.fillColor).filter(Boolean) as string[];
  };

  it("passes a real CSS name through, because Office knows it", async () => {
    expect(await drawWith("steelblue")).toContain("steelblue");
  });

  it("normalises a word that is not a colour instead of handing it over", async () => {
    for (const word of ["banana", "constructor", "notacolour"]) {
      const fills = await drawWith(word);
      expect(fills, `${word} was handed to the host verbatim`).not.toContain(word);
      for (const f of fills) expect(f, `${word} produced ${f}`).toMatch(/^#[0-9a-f]{6}$/i);
    }
  });
});

/**
 * A deck-wide rescale redraws each chart on its own slide, and the trace's
 * per-slide shape counter is the input to this project's only performance
 * claim — that drawing cost grows with what is already on the slide.
 *
 * It pooled them. `updateChartsInSlides` never named the slide it was aimed at,
 * so every chart in the `4feb5be` round keyed on the `(visible)` sentinel and
 * the counter climbed to 260 on a deck whose fullest slide held 24. The number
 * described nothing, and it looked perfectly healthy — a curve, rising, on
 * every line.
 *
 * `onSlideKey` is the only reason it was caught rather than plotted, which is
 * the argument for emitting a key beside any pooled total.
 */
describe("what the per-slide shape counter counts", () => {
  it("keys on the slide each chart is redrawn on, not on one sentinel", async () => {
    // Slide ids nothing else in this file uses. `shapesDrawnOnSlide` is a
    // per-RUN total by design — it answers "how much has this run already put
    // here" — so it is not reset between operations, and a shared `s1` would
    // carry every earlier test's draws into this one's first reading.
    const slides = [makeSlide("counter-a"), makeSlide("counter-b")];
    installHost(slides);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), slideId: "counter-a" });
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), slideId: "counter-b" });

    const found = (await listChartsInDeck()).charts;
    expect(found.length, "the two charts did not both land").toBe(2);
    setTracing(true);
    try {
      // A DIFFERENTLY-SHAPED SCENE, so this is genuinely a REDRAW.
      //
      // Until 2026-08-21 an identical scene redrew too, because a grouped chart
      // could not be updated in place and always fell back. It can now — the
      // update reads the group's members — so re-rendering the same config
      // takes the fast path, issues no batches, and this test had nothing to
      // count. The subject here is what a redraw keys on, so the fixture has to
      // provoke one rather than rely on the fast path being broken.
      const reshaped = buildChart({
        ...cfg,
        data: { categories: ["A", "B", "C", "D", "E"], series: [{ name: "S", values: [1, 2, 3, 4, 5] }] },
      });
      await updateChartsInSlides(
        found.map((c) => ({ scene: reshaped, target: c.target, opts: { tagData: JSON.stringify(cfg) } })),
      );
      const batches = traceLog().entries.filter((e) => e.message === "batch issued");
      expect(batches.length, "the redraw issued no batches to check").toBeGreaterThan(1);
      const keys = new Set(batches.map((b) => String(b.data?.onSlideKey)));
      expect(
        keys.has("(visible)"),
        "the redraw pooled its charts under the unnamed-slide sentinel, so the count spans slides",
      ).toBe(false);
      expect(keys.size, `every chart keyed the same: ${[...keys].join(", ")}`).toBe(2);
      // Independent, not merely differently labelled. Both slides took exactly
      // one identical insert before this redraw, so their first batches must
      // start from the same count — if the totals were still pooled, whichever
      // slide was redrawn second would start where the first one finished.
      const firsts = [...keys].map((key) => Number(batches.find((b) => b.data?.onSlideKey === key)!.data?.onSlide));
      expect(
        firsts[0],
        `the two slides' counters started at ${firsts.join(" and ")} after identical work, so one is carrying the other's total`,
      ).toBe(firsts[1]);
    } finally {
      setTracing(false);
    }
  });

  /**
   * EVERY DRAW'S LAST BATCH WENT UNTIMED, and the hole was half the data.
   *
   * `prevBatchMs` is only ever read by the NEXT batch's `issued` line, so a
   * draw's final batch was never recorded — 4,583 of 8,771 batches across the
   * archive, measured 2026-09-05 — and a chart small enough to fit in ONE batch
   * was never timed at all. `estimateInsertMs` prices exactly those, so the
   * curve in `insert-cost.ts` was fitted on a sample containing none of them.
   *
   * The single-batch case is what this draws, because it is the one the hole
   * swallowed whole: without the `last batch settled` line such a draw produces
   * NO timing at all, and every assertion below has nothing to find.
   */
  it("times the last batch of a draw, including a draw that is only one batch", async () => {
    installHost([makeSlide("settled-a")]);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    setTracing(true);
    try {
      await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), slideId: "settled-a" });
      const entries = traceLog().entries;
      const issued = entries.filter((e) => e.message === "batch issued");
      const settled = entries.filter((e) => e.message === "last batch settled");
      expect(issued.length, "nothing was drawn, so there is no last batch to time").toBeGreaterThan(0);
      expect(settled.length, "the draw ended without timing its final batch").toBe(1);
      const d = settled[0].data ?? {};
      expect(typeof d.ms, `the settled line carried no duration: ${JSON.stringify(d)}`).toBe("number");
      expect(Number(d.batchNo), "the settled line disagrees with how many batches were issued").toBe(issued.length);
      expect(Number(d.drew), "the settled line says the last batch drew nothing").toBeGreaterThan(0);
      // AFTER, and the name is load-bearing. On a one-batch draw the count is
      // still under the sentinel, so a "before" here would be the fabricated
      // zero that made the first reading of this archive wrong by 13x.
      expect(Number(d.onSlideAfter)).toBeGreaterThanOrEqual(Number(d.drew));
    } finally {
      setTracing(false);
    }
  });

  /**
   * A redraw is not an addition, and for a round this counter said it was.
   *
   * The field answers "how much has this run already put here" and feeds two
   * readers that both take it as the slide's size: `leastLoadedChart`, which
   * spreads the battery's draws, and every hand reading of the per-slide cost
   * curve. A chart redrawn in place deletes twenty-four shapes and adds
   * twenty-four — the slide does not change size — and the counter went up by
   * twenty-four each time.
   *
   * Round `957aca0` is the measurement. Slide 257's counter reached **92**
   * across three redraws of the deck-wide rescale, while the deck inventory
   * taken at the end of that run shows the slide holding **3** shapes. The
   * numbers looked perfectly healthy, which is the same way `onSlideKey`'s
   * pooling bug looked healthy, and they were describing no slide in the deck.
   */
  /**
   * An UNNAMED draw must still land on the slide it drew on.
   *
   * `opts.slideId` is absent on the ordinary insert — the pane's Insert button
   * names no slide — so every chart a user adds banked its shapes under the
   * `(visible)` sentinel and `shapesDrawnOn(realId)` answered zero for a slide
   * this run had just filled. That is not merely untidy: `slideHoldsOnlyChart`
   * reads this counter to decide whether an empty read of a slide is
   * believable, and it authorises DELETING the user's slide. The guard was
   * defeated on the commonest path in the add-in.
   *
   * Round `393e6e4` is where it showed: every batch of `insert onto a slide
   * that already has content` keyed on `(visible)` while every other scenario
   * named its slide.
   */
  it("counts an unnamed insert against the slide it actually drew on", async () => {
    installHost([makeSlide("counter-unnamed")]);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    // No slideId — exactly what the pane's Insert button passes.
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg) });
    expect(
      shapesDrawnOn("counter-unnamed"),
      "an unnamed insert did not count against the slide it drew on",
    ).toBeGreaterThan(0);
    // And the first batch is not stranded under the sentinel: the whole chart
    // is accounted for on the real slide, not just the batches after the host
    // answered the id.
    expect(shapesDrawnOn("counter-unnamed"), "some of the chart is still banked under the sentinel").toBe(
      estimateOfficeShapes(scene),
    );
    expect(shapesDrawnOn("(visible)"), "shapes were left behind on the sentinel").toBe(0);
  });

  /**
   * The half of that which the fake could not otherwise reach.
   *
   * On a real host `slide.load("id")` populates nothing until a sync runs, and
   * the first batch's key is chosen BEFORE that sync — so an unnamed draw's
   * first ten shapes are banked under `(visible)` and only the rest can name
   * the slide. `renderShapesChunked` moves them across the moment the host
   * answers; without that, `onSlide` under-reports by a batch forever and the
   * sentinel keeps a growing total that describes no slide.
   *
   * The fake answers `slide.id` from the first read, so nothing here could
   * exercise it until `slideIdUnreadableBeforeFirstSync` existed. That is the
   * point of the fault: a carry-over nothing can make fail is decoration, and
   * this repo has shipped that before.
   */
  it("moves the first batch's shapes across once the host names the slide", async () => {
    installHost([makeSlide("counter-late-id")]);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    faults.slideIdUnreadableBeforeFirstSync = true;
    try {
      await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg) });
    } finally {
      faults.slideIdUnreadableBeforeFirstSync = false;
    }
    expect(
      shapesDrawnOn("counter-late-id"),
      "the batches drawn before the host named the slide were left on the sentinel",
    ).toBe(estimateOfficeShapes(scene));
    expect(shapesDrawnOn("(visible)"), "the sentinel kept a total that describes no slide").toBe(0);
  });

  /**
   * The decision behind that, checked on both branches without a PowerPoint.
   *
   * The behavioural test above only ever exercises the second one: the fake
   * groups, so its redraw makes a single `delete()` call and the parts list is
   * empty. An ungrouped chart — the case this host actually produces — takes
   * the first branch, and nothing else here would notice it inverting.
   */
  it("reads a redraw's delete calls as shapes only when they enumerate the chart", () => {
    // The parts tag listed the chart: 24 calls went out, 24 shapes came off.
    expect(replacedShapeCount(24, 24), "an enumerated delete was not believed").toBe(24);
    expect(replacedShapeCount(24, 9), "an enumerated delete was overridden by the redraw's size").toBe(24);
    // One call: a group went, and its children are not knowable from here. The
    // chart going back is the best estimate of the chart that left.
    expect(replacedShapeCount(1, 24), "a group's children were counted as one shape").toBe(24);
    // …and one call for a chart that really was one shape still nets to zero
    // against a same-size redraw, which is the point.
    expect(replacedShapeCount(1, 1)).toBe(1);
  });

  it("does not count a redraw as if it had added the shapes again", async () => {
    installHost([makeSlide("counter-redraw")]);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), slideId: "counter-redraw" });
    const afterInsert = shapesDrawnOn("counter-redraw");
    expect(afterInsert, "the insert drew nothing").toBeGreaterThan(0);

    const found = (await listChartsInDeck()).charts;
    expect(found.length, "the chart did not land").toBe(1);
    await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: JSON.stringify(cfg) } }]);

    // The slide holds one chart before and one chart after, so the run's net
    // contribution has not moved. Exact, not "less than double": a decrement
    // that gave back the wrong number would still beat a doubling.
    expect(
      shapesDrawnOn("counter-redraw"),
      `a redraw of the same chart moved the slide's count from ${afterInsert}, so it is counting deleted shapes`,
    ).toBe(afterInsert);
  });
});

/**
 * Redrawing a chart to change one string.
 *
 * The add-in's update deletes every shape and adds every shape back, and on
 * PowerPoint for the web that is ~50 seconds for a 24-shape chart. A retitle
 * changes one node of twenty-four; a single edited data point changes two.
 *
 * The fast path applies to UNGROUPED charts, which is not a limitation so much
 * as the case that matters: a grouped chart's shapes are inside the group and
 * its parts tag does not list them, so there is no mapping from node to shape —
 * and the web host, which is where the 50 seconds live, ungroups every chart it
 * cannot group. A healthy host keeps its groups and keeps redrawing, which it
 * can afford.
 */
describe("updating only what changed", () => {
  const clustered = () => ({ ...sampleConfig("clustered"), ...DEFAULT_SIZE }) as ChartConfig;
  const liveIds = (slide: ReturnType<typeof makeSlide>) => slide.created.filter((s) => !s.deleted).map((s) => s.id);

  /** An ungrouped chart, which is what the web host produces for everything. */
  const drawLoose = async (cfg: ChartConfig) => {
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseGroups = 99;
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
    return slide;
  };

  afterEach(() => {
    faults.refuseGroups = 0;
    setTracing(false);
  });

  /**
   * A picture is not in the scene, so the scene may not decide the update.
   *
   * `render: "image"` on a config does NOT make a picture — the renderer takes
   * that path only when handed `pictureBase64` — so collapsing a chart to a
   * picture builds the SAME scene the chart already has. The differ compared 24
   * nodes to 24 identical ones, said nothing changed, wrote nothing, and
   * reported success; the slide kept its native shapes and the caller was told
   * the update landed.
   *
   * Round `89675b6` caught it: `updated only the shapes that changed {changed:
   * 0, of: 24}` on the collapse, and `explode a degraded picture` then reported
   * `the collapse added 0 shapes — the slide went from 24 to 24`.
   *
   * It matters well beyond the self-test. The auto-picture fallback is what the
   * add-in reaches for when this host has ALREADY failed to draw shapes, so the
   * one path that exists to rescue a struggling host was the path being
   * skipped.
   */
  it("does not take the fast path for an update that draws a picture", async () => {
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    const before = liveIds(slide).length;
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    // The same config and therefore the SAME scene — only `pictureBase64`
    // differs, which is exactly the input the differ cannot see.
    await updateChartInSlide(buildChart(cfg), target, {
      tagData: JSON.stringify(cfg),
      pictureBase64: "data:image/png;base64,UE5H",
    });
    const said = traceLog().entries.filter((e) => e.message === "updated only the shapes that changed");
    expect(said, "the differ decided a picture update, which it cannot see").toEqual([]);
    // …and the picture actually landed: a picture is ONE shape where the chart
    // was many, so the slide's live shape count must have collapsed.
    expect(liveIds(slide).length, `the slide still holds ${before} shapes, so no picture was drawn`).toBeLessThan(
      before,
    );

    /**
     * AND BOTH HALVES OF THE REDRAW ARE TIMED, which decides whether the
     * picture fast-path is a feature worth building.
     *
     * ~29 updates a round take this path. A fast path could at best reuse one
     * existing shape as the picture host, saving one create and one delete out
     * of ~25 operations — so whether it is worth 8% or 80% turns on what the
     * DELETE costs against the draw, and neither was timed. `batch issued`
     * covers the shape-by-shape draw and reported a picture as `1 of 1` with no
     * elapsed time at all.
     */
    const del = traceLog().entries.find((e) => e.message === "deleted the chart being replaced");
    expect(del, "a redraw's delete is still untimed").toBeTruthy();
    expect(typeof (del!.data as { ms?: unknown }).ms, "the delete traced no duration").toBe("number");
    // `removed` is what the host confirmed and took, and reading the ms without
    // it averages a one-shape delete together with a twenty-four-shape one.
    expect(typeof (del!.data as { removed?: unknown }).removed, "the delete did not say how many went").toBe("number");

    const pic = traceLog().entries.find((e) => e.message === "drew the chart as one picture");
    expect(pic, "the picture insert is still untimed").toBeTruthy();
    expect(typeof (pic!.data as { ms?: unknown }).ms, "the picture insert traced no duration").toBe("number");
    // What the picture stands in FOR: the comparison is against the same chart
    // drawn shape by shape, and the node count is what sets that price.
    expect((pic!.data as { nodes?: number }).nodes, "the picture did not say what it replaced").toBeGreaterThan(1);
  });

  it("writes one shape for a retitle and leaves the other 23 alone", async () => {
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    const before = liveIds(slide);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = { ...cfg, title: "Renamed" };
    const back = await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    const said = traceLog().entries.filter((e) => e.message === "updated only the shapes that changed");
    expect(said.length, "the chart was redrawn rather than updated").toBe(1);
    expect(said[0].data).toMatchObject({ changed: 1, of: 24 });

    // The strongest evidence available: an update that redraws cannot possibly
    // hand back the same shape ids, because every shape is new.
    expect(liveIds(slide), "the shapes were replaced, so this was a redraw").toEqual(before);
    expect(back?.lost, "an in-place update reported a loss").toBeFalsy();

    // And the chart is actually correct afterwards — both the picture and the
    // config, which is the pair a wrong fast path would split.
    const found = (await listChartsInDeck()).charts;
    expect(found).toHaveLength(1);
    expect((JSON.parse(found[0].configJson) as ChartConfig).title).toBe("Renamed");
    expect(
      slide.created.some((s) => !s.deleted && s.text === "Renamed"),
      "the new title never reached the slide",
    ).toBe(true);
  });

  it("falls back to a redraw when the engine has moved on", async () => {
    // The fingerprint is the whole reason this is sound. A chart drawn by a
    // different engine renders from its stored config to a scene that is NOT
    // what is on the slide, and diffing against it would skip nodes that really
    // did change and leave them stale forever.
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    // Exactly the state an upgrade produces: the config is intact, the
    // fingerprint names a scene this build does not produce.
    const anchor = slide.created.find((s) => s.id === target.shapeId)!;
    anchor.tagStore.set("POWERCHART_SCENE", "notthisone");
    const before = liveIds(slide);
    setTracing(true);
    const next = { ...cfg, title: "Renamed" };
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    expect(
      traceLog().entries.some((e) => e.message === "updated only the shapes that changed"),
      "a chart whose fingerprint did not match was updated in place anyway",
    ).toBe(false);
    expect(liveIds(slide), "the fallback did not actually redraw").not.toEqual(before);
    const found = (await listChartsInDeck()).charts;
    expect((JSON.parse(found[0].configJson) as ChartConfig).title, "the fallback lost the edit").toBe("Renamed");
  });

  it("falls back when the chart was resized, because every node moved", async () => {
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    const before = liveIds(slide);
    setTracing(true);
    const next = { ...cfg, width: cfg.width! + 40 };
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(
      traceLog().entries.some((e) => e.message === "updated only the shapes that changed"),
      "a resized chart took the in-place path",
    ).toBe(false);
    expect(liveIds(slide)).not.toEqual(before);
  });

  it("says WHY it declined, every time it declines", async () => {
    // The first real round on a build carrying the fast path produced not one
    // line — no success, no refusal — which is indistinguishable from the code
    // not being there, and left the reason to be reasoned out of grouping
    // traces and a deck inventory. The reasons are not interchangeable: a
    // grouped chart has no mapping by design, a missing fingerprint means an
    // older build drew it, and a refused id readback is a fact about the host.
    // Three different next steps, one silent `return false`.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg = clustered();
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    const next = { ...cfg, title: "Renamed" };
    await updateChartInSlide(buildChart(next), target, {
      tagData: JSON.stringify(next),
      // Forces the one decline that is still correct — see below.
      pictureBase64: "iVBORw0KGgo=",
    });
    const said = traceLog().entries.filter((e) => e.message === "not updating in place — redrawing instead");
    expect(said.length, "a declined in-place update passed without a word").toBe(1);
    // THE PICTURE CASE, named as itself rather than as a generic refusal.
    //
    // This fixture used a GROUPED chart until 2026-08-21, because a grouped
    // chart was refused for want of a parts list. It is not any more — the
    // update reads the group's members — so the fixture had to change to keep
    // the test testing what it says: that a decline is never silent.
    //
    // The picture path is the decline that remains legitimate: a picture is not
    // in the scene, so the differ has nothing to compare and refusing is the
    // correct answer rather than a limitation.
    expect(String(said[0].data?.why)).toMatch(/picture/);
  });

  it("names a missing fingerprint separately from a missing mapping", async () => {
    // A chart with no fingerprint and a chart with no mapping both decline, and
    // they want different responses: the first fixes itself on this very
    // redraw, the second never will.
    //
    // The assertion anchors on the CONDITION, not on why it arose. It used to
    // match /older build/, from a message that told a story it could not know —
    // for six rounds it said "drawn by an older build" about charts the current
    // build had just drawn through a writer that never stamped the tag.
    const cfg = clustered();
    const slide = await drawLoose(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    slide.created.find((s) => s.id === target.shapeId)!.tagStore.delete("POWERCHART_SCENE");
    setTracing(true);
    const next = { ...cfg, title: "Renamed" };
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    const said = traceLog().entries.filter((e) => e.message === "not updating in place — redrawing instead");
    expect(said.length).toBe(1);
    expect(String(said[0].data?.why), "a missing fingerprint was reported as having no parts list").toMatch(
      /fingerprint/,
    );
  });

  it("updates a GROUPED chart through its group members", async () => {
    // THIS TEST USED TO ASSERT THE OPPOSITE, and the reversal is the point.
    //
    // It read: "redraws a GROUPED chart, which has no node-to-shape mapping —
    // its shapes are inside the group and the parts tag does not list them, so
    // there is nothing to write to." That was true of the code and false of the
    // host: `ShapeGroup.shapes` reaches the members at PowerPointApi 1.8, and
    // office-js#3014's "sub-shapes cannot be reached" is a 2022 note that has
    // gone stale.
    //
    // The cost of believing it was the whole feature. The in-place update had
    // never run once in 117 archived rounds, because this host groups nearly
    // every chart (18 of 21 in round 142) and every grouped chart was refused
    // here.
    //
    // WHAT THE OLD TEST WAS REALLY PROTECTING SURVIVES, in the assertion below
    // and in the sibling test above: a mapping that does not line up must be
    // REFUSED, never guessed at, because writing a bar's geometry onto whatever
    // shape sits at that index is worse than any redraw.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg = clustered();
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
    const target = (await listChartsInDeck()).charts[0].target;
    // The premise, pinned: this chart really is grouped and really has no parts
    // tag. If either changes, this stops testing what it says it does.
    expect(target.partIds?.length ?? 0, "a grouped chart should carry no parts tag").toBe(0);
    setTracing(true);
    const next = { ...cfg, title: "Renamed" };
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    expect(
      traceLog().entries.some((e) => e.message === "updated only the shapes that changed"),
      "a grouped chart was NOT updated in place — the group members were not read",
    ).toBe(true);
    // And the update has to have actually landed, not merely been attempted.
    const found = (await listChartsInDeck()).charts;
    expect((JSON.parse(found[0].configJson) as ChartConfig).title).toBe("Renamed");
  });

  it("records WHAT IT SAW when a grouped chart has no readable members, not only the conclusion", async () => {
    // TWELVE ROUNDS OF A DEFECT WITH NO EVIDENCE ATTACHED TO ANY OF THEM. This
    // decline has fired exactly once per round since round 151 — deterministic
    // and reproducible — and every one of those traces carried the sentence and
    // nothing else. There was no way to tell which of three things had happened.
    //
    // `groupMembersInOrder` returns `[]` for three different reasons and only
    // the trace can separate them: no member collection was queued at all (the
    // chart was not read as grouped), the host would not load the collection's
    // items, or it loaded and held fewer than two — which cannot be a chart.
    // Three different next steps, one silent sentence. The house defect, in the
    // product's own instrument.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg = clustered();
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
    const target = (await listChartsInDeck()).charts[0].target;
    expect(target.partIds?.length ?? 0, "this fixture needs a GROUPED chart").toBe(0);
    // The collection is queued and answers — with nothing in it. That is the
    // case a bare sentence cannot distinguish from the host refusing to answer.
    const group = slide.created.find((s) => s.id === target.shapeId)!;
    expect(group.grouped?.length ?? 0, "the premise: this really is a group with members").toBeGreaterThan(1);
    group.grouped = [];
    setTracing(true);
    try {
      const next = { ...cfg, title: "Renamed" };
      await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
      const said = traceLog().entries.filter((e) => e.message === "not updating in place — redrawing instead");
      expect(said.length).toBe(1);
      const d = said[0].data as { why: string; collection: string; members: number | null; nodes: number };
      expect(String(d.why)).toMatch(/no readable group members/);
      expect(d.collection, "could not say whether a collection was even asked for").toBe("queued");
      expect(d.members, "an empty collection and an unreadable one read identically").toBe(0);
      expect(d.nodes, "the scene's size, so a mismatch can be reasoned about").toBeGreaterThan(0);
    } finally {
      setTracing(false);
    }
  });

  it("writes node 0 to the group's ANCHOR, not to the group itself", async () => {
    // THE HOST CAUGHT THIS AND THE SUITE DID NOT. Round 145 ran the in-place
    // update on the real host for the first time and refused three charts with
    // `InvalidArgument | errorLocation=Shape.textFrame`, every one of them
    // `changed 1 of 24`.
    //
    // `shapes = [old, ...parts]` is right for a parts-list chart, where the
    // tagged shape IS node 0. For a GROUPED chart the tagged shape is the group
    // and node 0 is the group's first member, so node 0's properties went onto
    // the group — which has `fill` and `lineFormat` (both navigations
    // succeeded) but no `textFrame`, which is precisely where the host stopped.
    //
    // Nodes 1..n were mapped correctly the whole time, so ONLY an edit that
    // touches node 0 can show it. A title edit is exactly that, and it is why
    // the test above — which also edits the title — passed anyway: the fake
    // host lets a group accept a name and a text frame, so the wrong target was
    // invisible here while the real host refused it every time.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg = clustered();
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
    const target = (await listChartsInDeck()).charts[0].target;
    expect(target.partIds?.length ?? 0, "premise: a grouped chart carries no parts tag").toBe(0);
    setTracing(true);
    const next = { ...cfg, title: "Renamed" };
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });

    const group = slide.created.find((s) => s.type === "group");
    expect(group, "premise: the fixture did not produce a group").toBeDefined();
    // The group must still be the group. If node 0 was written onto it, it is
    // wearing the title node's name.
    expect(group!.name, "node 0 was written onto the GROUP — the host refuses that at Shape.textFrame").toBe(
      "PowerChart",
    );
    // …and the anchor, which is what node 0 should have gone to, carries it.
    expect(
      slide.created.find((s) => s.name === "title"),
      "no shape carries the title node",
    ).toBeDefined();
  });

  it("still lands the chart when the host refuses the in-place batch", async () => {
    // Nothing has been deleted when the writes go out, so a refusal costs the
    // chart nothing but time — the redraw does the whole job. The alternative,
    // a chart whose picture is new and whose config is old, would silently
    // revert the user's edit the next time they opened it.
    const cfg = clustered();
    await drawLoose(cfg);
    const target = (await listChartsInDeck()).charts[0].target;
    setTracing(true);
    faults.refuseTagWrites = 1;
    try {
      const next = { ...cfg, title: "Renamed" };
      await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    } finally {
      faults.refuseTagWrites = 0;
    }
    const said = traceLog().entries.map((e) => e.message);
    expect(said, "a refused in-place update passed without a word").toContain(
      "in-place update refused — redrawing instead",
    );
    const found = (await listChartsInDeck()).charts;
    expect(found, "the chart was lost when the fast path was refused").toHaveLength(1);
    expect((JSON.parse(found[0].configJson) as ChartConfig).title).toBe("Renamed");
  });
});

/**
 * `applyNodeInPlace` and the adders must write the same properties.
 *
 * A property the adder sets and the applier does not is a chart that looks
 * right when first drawn and wrong after an edit — a bar keeping its old
 * colour, a label keeping its old font — with nothing in any log to say so,
 * because from the host's point of view the update succeeded. Neither the type
 * checker nor any behavioural test can see that: the applier compiles fine
 * without the line, and a test that drew a chart and edited it would only catch
 * the properties it happened to assert on.
 *
 * So the source is the check. Crude by design, and it fails with the NAME of
 * whichever property was forgotten — which is exactly what a reader adding a
 * line to `addText` needs to be told.
 *
 * WHAT IT CANNOT SEE, since 2026-08-29: the applier's writes are now GATED on
 * which property groups the diff says changed, and a source scan cannot tell a
 * gated write from an unreachable one. This file still answers "is the line
 * present"; `test/in-place-writes.test.ts` answers "does it fire, and only when
 * it should" by counting what a real call sends. Neither replaces the other.
 */
describe("the in-place applier and the adders it mirrors", () => {
  const source = readFileSync("src/render/powerpoint.ts", "utf8");
  const body = (start: RegExp): string => {
    const at = source.search(start);
    expect(at, `could not find ${start} in powerpoint.ts`).toBeGreaterThan(-1);
    const rest = source.slice(at + 1);
    const end = rest.search(/\n(?:export )?(?:async )?function /);
    return rest.slice(0, end === -1 ? undefined : end);
  };
  /** Property writes, as `receiver.property`, ignoring the local's name. */
  const writes = (text: string): Set<string> => {
    const found = new Set<string>();
    for (const m of text.matchAll(/\b(?:shape|tf|font|lf)\.([A-Za-z]+(?:\.[A-Za-z]+)*)\s*=/g)) found.add(m[1]);
    for (const m of text.matchAll(/\b(?:solidFill|strokeColor)\(\s*\w+\.([A-Za-z]+)/g)) found.add(`${m[1]}:filled`);
    for (const m of text.matchAll(/\b\w+\.([A-Za-z]+)\.clear\(\)/g)) found.add(`${m[1]}:cleared`);
    return found;
  };
  // `(?:export )?` because the applier is exported for `in-place-writes.test.ts`
  // to count what it sends. The end pattern below already allowed for it; the
  // start did not, and the whole suite failed to LOAD rather than to assert.
  const applier = writes(body(/\n(?:export )?function applyNodeInPlace\(/));
  const rectCase = writes(body(/\nfunction addNode\(/).split('case "line"')[0]);
  const textAdder = writes(body(/\nfunction addText\(/));

  it("writes everything the rect adder writes", () => {
    const missing = [...rectCase].filter((p) => !applier.has(p));
    expect(missing, `applyNodeInPlace never sets: ${missing.join(", ")}`).toEqual([]);
  });

  it("writes everything the text adder writes", () => {
    const missing = [...textAdder].filter((p) => !applier.has(p));
    expect(missing, `applyNodeInPlace never sets: ${missing.join(", ")}`).toEqual([]);
  });

  it("also writes the string, which the adder passes as an argument", () => {
    expect(applier, "the applier cannot change a chart's text").toContain("textRange.text");
  });

  it("is comparing a real set of properties, not an empty one", () => {
    // A regex that stopped matching would make both tests above pass forever.
    expect(rectCase.size, "the rect scan found nothing to compare").toBeGreaterThan(3);
    expect(textAdder.size, "the text scan found nothing to compare").toBeGreaterThan(6);
  });
});

describe("telling a wedged host from a wedged call", () => {
  // `idleMs` is `issued - lastAnswered`, and `lastAnswered` is sampled when the
  // DEADLINE fires. So a negative value means the host answered something else
  // AFTER this call went out, while it was still waiting — the host is alive and
  // one call is stuck. Rounds 089 and 090 both reported `idleMs: -89057`-ish and
  // nobody could read it; that number was the answer to the question the round
  // was being run to settle.
  it("says THIS CALL ALONE when the host answered during the wait", () => {
    const said = stallShape(-89_057);
    expect(said).toMatch(/THIS CALL ALONE/);
    expect(said, "the reader has to be told how long after, not just that it happened").toMatch(/89057ms AFTER/);
  });

  it("says the host had gone quiet when the gap is real and positive", () => {
    expect(stallShape(12_000)).toMatch(/already been quiet for 12000ms/);
  });

  it("refuses to read anything into a sequential gap", () => {
    // A 1ms gap is true of every call in a healthy round — the trap this
    // project already fell into once with `idleMs: 1` on a rasterise.
    expect(stallShape(1)).toMatch(/says little/);
  });

  it("says so when the host has never answered at all", () => {
    expect(stallShape(Infinity)).toMatch(/nothing at all yet/);
  });
});

describe("how long a rasterise actually takes", () => {
  it("names its call site in the STALL, which is the half that matters", async () => {
    // The success line was easy; the stall is the point. All 34 rasterise stalls
    // on record say `what: "rasterising a slide"` — one label for five call
    // sites — so "every stall is the visibility control" could only ever be
    // inferred from where the line sat in the trace.
    //
    // The first version of this guard tested only the success path, and a
    // mutation that put the generic label back on the STALL left it green.
    const slide = makeSlide("s1");
    installHost([slide]);
    _setReadbackTimeoutForTest(30);
    // The SECOND sync only. `slideImageBase64` syncs once for the null check —
    // which is unbounded — and again for the rasterise, which is the bounded one
    // this test is about. Wedging from zero hangs the first and proves nothing.
    faults.wedgeAfterSyncs = 1;
    const seen: { message: string; data?: Record<string, unknown> }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    const png = await slideImageBase64("s1", 640, "the visibility CONTROL render");
    onTrace(undefined);
    setTracing(false);
    faults.wedgeAfterSyncs = null;
    _setReadbackTimeoutForTest(90_000);

    expect(png, "the wedge did not take").toBeUndefined();
    const gaveUp = seen.filter((e) => e.message === "gave up waiting");
    expect(gaveUp.length, "no stall was recorded at all").toBeGreaterThanOrEqual(1);
    expect(gaveUp[0].data?.what, "the stall could not say which rasterise it was").toBe(
      "the visibility CONTROL render",
    );
  });

  it("says WHICH rasterise it was, so a stall names its own call site", () => {
    // All 34 rasterise stalls across 86 rounds trace `what: "rasterising a
    // slide"` — one generic label for five different call sites. The claim that
    // every stall lands on the visibility control is therefore an INFERENCE from
    // where the line sits in the trace, and no round file can be asked to
    // confirm it. The diagnosis rested on something unreadable.
    const slide = makeSlide("s1");
    installHost([slide]);
    const seen: { message: string; data?: Record<string, unknown> }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    return slideImageBase64("s1", 640, "the visibility CONTROL render").then(() => {
      onTrace(undefined);
      setTracing(false);
      const timed = seen.filter((e) => e.message === "rasterised a slide");
      expect(timed.length).toBe(1);
      expect(timed[0].data?.label, "the rasterise did not name its call site").toBe("the visibility CONTROL render");
    });
  });

  it("times a SUCCESSFUL rasterise — 86 rounds recorded only the failures", () => {
    // 35 rasterise stalls on record, every one burning the full 20s budget, and
    // not one successful duration. So the budget can only be sized from harm,
    // which is exactly where the deck-style timeout started before it was
    // measured. The visibility gate is blind in more than half of all rounds
    // waiting on this call; whether a good one costs 200ms or 8s decides whether
    // the budget should move at all.
    const slide = makeSlide("s1");
    installHost([slide]);
    const seen: { message: string; data?: Record<string, unknown> }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    return slideImageBase64("s1", 640).then((png) => {
      onTrace(undefined);
      setTracing(false);
      expect(png, "the fake did not rasterise at all").toBeTruthy();
      const timed = seen.filter((e) => e.message === "rasterised a slide");
      expect(timed.length, "a successful rasterise left no duration behind").toBe(1);
      expect(typeof timed[0].data?.ms, "the duration is not a number").toBe("number");
      expect(timed[0].data?.bytes, "the size is what says a rasterise produced anything").toBeGreaterThan(0);
    });
  });
});

describe("the style a deck carries", () => {
  it("round-trips a style through the deck's own custom XML part", async () => {
    installHost([makeSlide("s1")]);
    expect(await readDeckStyle()).toBeNull(); // an unbranded deck says nothing
    expect(await writeDeckStyle({ palette: ["#2a78d6"], fontSize: 11 })).toBe(true);
    expect(await readDeckStyle()).toEqual({ palette: ["#2a78d6"], fontSize: 11 });
  });

  it("REPLACES the style rather than adding a second part", async () => {
    // `getOnlyItemOrNullObject` — how the read finds it — returns a NULL OBJECT
    // for a namespace holding two, so a deck that quietly accumulated three
    // styles would have no style at all. The write deletes every part in the
    // namespace first, and this is the test that says so: without it the second
    // read comes back empty on a deck that has just been branded twice.
    //
    // This comment said the read "refuses" and that the second read "throws".
    // `@types/office-js` says otherwise — `getOnlyItem` raises on more than one,
    // the OrNullObject variant we call returns null — and the fake had been
    // throwing to match the comment rather than the contract. The corrected
    // behaviour is WORSE for a user and better to model: a second part does not
    // announce itself, it silently unbrands the deck.
    installHost([makeSlide("s1")]);
    await writeDeckStyle({ palette: ["#111111"] });
    await writeDeckStyle({ palette: ["#222222"] });
    expect(await readDeckStyle()).toEqual({ palette: ["#222222"] });
  });

  it("clears the deck's style when handed nothing", async () => {
    installHost([makeSlide("s1")]);
    await writeDeckStyle({ fontFamily: "Segoe UI" });
    expect(await writeDeckStyle(null)).toBe(true);
    expect(await readDeckStyle()).toBeNull();
  });

  it("says the write did not happen when the host refuses it", async () => {
    // A style nobody stored and a style stored somewhere invisible fail the
    // same way from the outside, and only one of them is worth retrying — so
    // the pane is told, rather than shown a success it cannot verify.
    installHost([makeSlide("s1")]);
    faults.refuseCustomXmlWrites = true;
    expect(await writeDeckStyle({ fontSize: 12 })).toBe(false);
    faults.refuseCustomXmlWrites = false;
    expect(await readDeckStyle()).toBeNull();
  });

  it("reads a DOUBLY-branded deck as carrying no style, silently — the contract, not the fake", async () => {
    // The state `writeDeckStyle`'s delete-then-add exists to prevent, and which
    // nothing could reach until now. Per `@types/office-js`,
    // `getOnlyItemOrNullObject` returns null for anything but exactly one item;
    // it is `getOnlyItem` that raises on "no items or more than one". So this is
    // NOT a failed read and must not be reported as one — it is a branded deck
    // that answers as unbranded, which is the quiet, worse outcome.
    //
    // NOT VERIFIED ON A REAL HOST: no round has put two parts in the namespace.
    // This pins the documented contract so a future change cannot drift from it
    // silently, and the fake no longer models a throw nobody has observed.
    installHost([makeSlide("s1")]);
    expect(await writeDeckStyle({ palette: ["#2a78d6"] })).toBe(true);
    faults.duplicateCustomXmlPart = true;
    expect(await readDeckStyleWithReason()).toEqual({ style: null, unreadable: false });
    faults.duplicateCustomXmlPart = false;
    // And with the duplicate gone the deck is branded again, so the fault is
    // modelling a second PART and not a corrupted store.
    expect(await readDeckStyle()).toEqual({ palette: ["#2a78d6"] });
  });

  it("asks ONE cheaper question when the read fails, to say which half is broken", async () => {
    // Rounds 089 and 090 both show the read using its whole budget and never
    // answering, and the round file cannot say WHICH part hung — the namespace
    // lookup, `getOnlyItemOrNullObject`, the `load` and `getXml` all go into one
    // batch, so one sync covers all four.
    //
    // `getCount()` needs the namespace lookup and nothing else, so its answer
    // splits the two explanations. Here the read is refused and the count is
    // not, which is the "reachable, fault is further in" branch.
    installHost([makeSlide("s1")]);
    // A deck that HAS a part, so the read reaches `getOnlyItemOrNullObject` at
    // all: the count-first guard returns early on an empty namespace, which is
    // the whole point of it — the empty case can no longer hang.
    await writeDeckStyle({ palette: ["#2a78d6"] });
    const seen: { message: string }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    faults.refuseCustomXmlReads = true; // getOnlyItemOrNullObject throws, getCount does not
    const r = await readDeckStyleWithReason();
    faults.refuseCustomXmlReads = false;
    onTrace(undefined);
    setTracing(false);
    expect(r).toEqual({ style: null, unreadable: true });
    const said = seen.filter((e) => /namespace IS reachable/.test(e.message));
    // Once per failed attempt, and the read attempts twice now — see the retry
    // test below. What matters is that a failure is never silent.
    expect(said.length, "failed without saying which half of the read broke").toBeGreaterThanOrEqual(1);
    // WHICH CALL, MEASURED. This field was a hardcoded sentence naming
    // getOnlyItemOrNullObject, and it kept printing that after the read gained a
    // `getCount` of its own — round 098 shows it asserting the old answer on a
    // build where the count runs first. It now carries the operation the bounded
    // sync names, so it cannot describe a call that did not fail.
    const failedAt = (said[0] as { data?: { failedAt?: string } }).data?.failedAt ?? "";
    // Derived from the actual failure, whatever shape it takes. A synchronous
    // refusal (this fake) carries the thrown message; a real-host TIMEOUT comes
    // through `boundedSync`, whose text carries `at=<operation>` and is parsed
    // to the operation name. Either way it describes what happened.
    expect(failedAt, "the probe said nothing about which call failed").toBeTruthy();
    expect(failedAt, "still asserting the old hardcoded conclusion").not.toMatch(
      /getOnlyItemOrNullObject\/load\/getXml/,
    );
    expect(failedAt, "did not come from the failure").toMatch(/custom XML read/);
  });

  it("replays what the read concluded, because it happens BEFORE the round does", async () => {
    // Round 091: not one deck-style line in 529 entries, on a build where the
    // probe exists — because the read fires from `Office.onReady` and its 10s
    // budget expires during the driver's setup, before tracing is on. The
    // staked prediction came back `undetermined  neither line appeared`.
    //
    // A fix that hides its own subject is worse than the cost it removed, so the
    // conclusion is remembered and re-emitted verbatim when a round begins.
    installHost([makeSlide("s1")]);
    await writeDeckStyle({ palette: ["#2a78d6"] }); // count 1, so the read gets past the guard
    _resetDeckStyleVerdictForTest();
    faults.refuseCustomXmlReads = true;
    await readDeckStyleWithReason(); // happens with tracing OFF, as on a real pane
    faults.refuseCustomXmlReads = false;

    const seen: { message: string }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    replayDeckStyleVerdict();
    onTrace(undefined);
    setTracing(false);
    expect(
      seen.filter((e) => /namespace IS reachable/.test(e.message)).length,
      "the round could not see what the read concluded",
    ).toBe(1);
  });

  it("is emitted BY the round-start line, not only by a direct call", async () => {
    // The first draft tested `replayDeckStyleVerdict()` on its own, and deleting
    // the call from `traceEnvironment` — the actual wiring, and the actual bug —
    // left the suite green. Mutation caught it. The round only ever reaches this
    // through `traceEnvironment`, so that is what has to be driven.
    installHost([makeSlide("s1")]);
    await writeDeckStyle({ palette: ["#2a78d6"] }); // count 1, so the read gets past the guard
    _resetDeckStyleVerdictForTest();
    faults.refuseCustomXmlReads = true;
    await readDeckStyleWithReason();
    faults.refuseCustomXmlReads = false;

    const seen: { message: string }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    traceEnvironment("deadbee");
    onTrace(undefined);
    setTracing(false);
    expect(
      seen.filter((e) => /namespace IS reachable/.test(e.message)).length,
      "round start did not carry what the read concluded",
    ).toBe(1);
  });

  it("says nothing at round start when the read never failed", async () => {
    // A replay that fires on a healthy pane would put a deck-style line in every
    // round file and make the ledger judge a question nobody asked.
    _resetDeckStyleVerdictForTest();
    const seen: { message: string }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    replayDeckStyleVerdict();
    onTrace(undefined);
    setTracing(false);
    expect(seen.length).toBe(0);
  });

  it("gives the deck-style read its OWN budget, not the 90s readback one", async () => {
    // ROUNDS 089 AND 090, BOTH: this read consumed its entire 90-second budget
    // and never answered, ~50ms after the host had answered a slide listing. The
    // pane-load caller is fire-and-forget, so that held a `PowerPoint.run`
    // context open for a minute and a half on every load, for a call that was
    // not going to answer.
    //
    // The assertion is that the read is bounded by `deckStyleTimeoutMs` — set
    // absurdly low here — rather than by `readbackTimeoutMs`, which is left at
    // its default. A wall-clock ceiling well under the readback budget is what
    // separates the two; asserting a duration is not the point, asserting WHICH
    // budget applies is.
    installHost([makeSlide("s1")]);
    process.env.PW_DECK_STYLE_TIMEOUT_MS = "30";
    faults.wedgeAfterSyncs = 0; // every sync from here on never answers
    const began = Date.now();
    const r = await readDeckStyleWithReason();
    const took = Date.now() - began;
    faults.wedgeAfterSyncs = null;
    delete process.env.PW_DECK_STYLE_TIMEOUT_MS;
    expect(r).toEqual({ style: null, unreadable: true });
    expect(took, "waited on the readback budget instead of its own").toBeLessThan(5_000);
  });

  it("SAYS when the read answered — silence is not evidence of success", async () => {
    // ROUND 106 IS WHY. It carried zero `deck-style` entries and the staked
    // claim came back `held` — but the probe only writes when the read FAILS, so
    // an empty scope could equally mean the read never ran. A cure and a silence
    // are the same shape, which is this project's house defect pointed the other
    // way: absence of a failure line read as evidence of success.
    installHost([makeSlide("s1")]);
    _resetDeckStyleVerdictForTest();
    await readDeckStyleWithReason();

    const seen: { message: string }[] = [];
    setTracing(true);
    onTrace((e) => seen.push(e));
    replayDeckStyleVerdict();
    onTrace(undefined);
    setTracing(false);
    expect(
      seen.filter((e) => /the deck-style read answered/.test(e.message)).length,
      "a successful read left the round file silent, which cannot be told from never running",
    ).toBe(1);
  });

  it("warms the surface with a call whose failure is the point", async () => {
    // Seven rounds say the FIRST custom-XML call after a pane loads fails and the
    // second works. Retrying inside the read was tried (#602) and reverted (#603)
    // because the read's retry spent the good second call and starved the probe.
    // So the bad call is spent here, where nothing waits on it.
    installHost([makeSlide("s1")]);
    await writeDeckStyle({ palette: ["#2a78d6"] });
    faults.refuseCustomXmlReadsOnce = true;
    // The warm-up eats the refusal...
    await expect(warmCustomXmlSurface()).resolves.toBeUndefined();
    // ...so the read that follows is the second call, and it answers.
    expect(await readDeckStyleWithReason()).toEqual({ style: { palette: ["#2a78d6"] }, unreadable: false });
    faults.refuseCustomXmlReadsOnce = false;
  });

  it("never throws out of the warm-up — failing is what it is for", async () => {
    // A warm-up that rejects would take the pane-load chain down with it, and
    // the chain it sits in is `void warm().then(read)`. The whole design is that
    // this call is allowed to fail.
    installHost([makeSlide("s1")]);
    faults.refuseCustomXmlReads = true;
    await expect(warmCustomXmlSurface()).resolves.toBeUndefined();
    faults.refuseCustomXmlReads = false;
  });

  it("attempts ONCE — the retry made this worse and the rounds took it back", async () => {
    // #602 shipped a retry here and the very next pair reverted it. Two rounds
    // on each side of that single commit, nothing else changed:
    //
    //   096-099   pre-retry              probe: the namespace IS reachable
    //   100, 101  count-first, no retry   probe: the namespace IS reachable
    //   102, 103  WITH the retry          probe: the namespace is UNREACHABLE too
    //
    // The probe had answered on every round it ever ran — that WAS the evidence
    // for "the second call works" — and it stopped the moment the read began
    // making two calls of its own ahead of it. The read's retry spends the good
    // second call and the probe, now third, gets the same nothing the first did.
    //
    // So one refused read is a failed read, and says so.
    installHost([makeSlide("s1")]);
    await writeDeckStyle({ palette: ["#2a78d6"] });
    faults.refuseCustomXmlReadsOnce = true;
    const r = await readDeckStyleWithReason();
    faults.refuseCustomXmlReadsOnce = false;
    expect(r, "retried, and the rounds say the retry costs the probe its answer").toEqual({
      style: null,
      unreadable: true,
    });
  });

  it("reports unreadable rather than a silent absence when the read fails", async () => {
    // The `unreadable` flag exists so a failed read is never reported as "this
    // deck carries no style". That was true when the read attempted twice and
    // stays true now that it attempts once.
    installHost([makeSlide("s1")]);
    await writeDeckStyle({ palette: ["#2a78d6"] });
    faults.refuseCustomXmlReads = true;
    const r = await readDeckStyleWithReason();
    faults.refuseCustomXmlReads = false;
    expect(r).toEqual({ style: null, unreadable: true });
  });

  it("never asks for the only item of an EMPTY namespace — the call that hangs on this host", async () => {
    // ROUNDS 096 AND 097, a pair on one build, both recorded the same thing
    // through the failure probe: `getCount()` ANSWERS and reports `parts: 0`,
    // and it is `getOnlyItemOrNullObject` that then never returns. So the
    // hanging call is the one asking for the only item of an EMPTY collection —
    // which is the state every unbranded deck is in, i.e. every pane load.
    //
    // The fault here refuses exactly that call. An unbranded deck must now come
    // back clean WITHOUT touching it, so this passes only if the count-first
    // guard is doing its job.
    installHost([makeSlide("s1")]);
    faults.refuseCustomXmlReads = true;
    const r = await readDeckStyleWithReason();
    faults.refuseCustomXmlReads = false;
    expect(r, "an unbranded deck reached the call that hangs").toEqual({ style: null, unreadable: false });
  });

  it("says the read FAILED rather than reporting an absence", async () => {
    // ROUND 089. `reading the deck's style` hung for its full 90s budget on the
    // real host and the catch turned that into `null` — the same value a deck
    // carrying no style returns. `style-from-deck` AWAITS this, so the pane told
    // the user their deck was unbranded and switched them to the browser's style
    // on the strength of a read that never happened.
    installHost([makeSlide("s1")]);
    expect(await writeDeckStyle({ palette: ["#2a78d6"] })).toBe(true); // the deck DOES carry one
    faults.refuseCustomXmlReads = true;
    expect(await readDeckStyleWithReason()).toEqual({ style: null, unreadable: true });
    faults.refuseCustomXmlReads = false;
    // And the distinction is real: the same shape of answer, from a deck that
    // genuinely has none, is NOT unreadable.
    await writeDeckStyle(null);
    expect(await readDeckStyleWithReason()).toEqual({ style: null, unreadable: false });
  });

  it("keeps the plain reader's contract, because one caller must not care", async () => {
    // The pane-load caller is fired and forgotten and wants exactly this: a
    // style or nothing, no branch. Widening its return would have made a
    // failed read something that path had to handle, on a path whose whole
    // design is that it does not.
    installHost([makeSlide("s1")]);
    faults.refuseCustomXmlReads = true;
    expect(await readDeckStyle()).toBeNull();
    faults.refuseCustomXmlReads = false;
  });

  it("neither reads nor writes on a host below PowerPointApi 1.7", async () => {
    installHost([makeSlide("s1")], [], undefined, (v) => Number(v) < 1.7);
    expect(await writeDeckStyle({ fontSize: 12 })).toBe(false);
    expect(await readDeckStyle()).toBeNull();
  });
});

describe("an in-place update writes in chunks", () => {
  const CATS = ["a", "b", "c", "d", "e", "f", "g", "h"];
  const base: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    title: "A",
    data: { categories: CATS, series: [{ name: "S", values: [10, 20, 30, 40, 50, 60, 70, 80] }] },
  };
  const withValues = (v: number[]): ChartConfig => ({
    ...base,
    data: { categories: CATS, series: [{ name: "S", values: v }] },
  });

  /** Syncs spent updating one chart with `next`, from a freshly installed host. */
  const syncsToUpdate = async (next: ChartConfig) => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => true);
    await insertSceneIntoSlide(buildChart(base), { tagData: JSON.stringify(base), left: 60, top: 90 });
    const target = (await listChartsInDeck()).charts[0].target;
    const before = trips.syncs;
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    return trips.syncs - before;
  };

  it("splits a write too big for one batch, and does not split a small one", async () => {
    // WHAT THE HOST REFUSES IS SIZE, NOT SHARE. Round 150 raised
    // UPDATE_SHARE_LIMIT to 0.8 so eight charts would write eighteen shapes
    // apiece, and PowerPoint crashed on all seven attempts at 255-284s. At 0.6
    // the same scenario writes nine apiece, completes 13/13, and runs 31s
    // faster. The limit is a proxy for a cost the host measures in shapes per
    // batch, and this bounds that cost directly.
    //
    // THE PREMISE, PINNED. These two edits are chosen for their changed-node
    // counts, not their labels: if a layout change alters either, this test
    // stops testing what it says it does.
    const b = buildChart(base);
    const small = planSceneUpdate(b, buildChart({ ...base, title: "Renamed" }));
    const big = planSceneUpdate(b, buildChart(withValues([11, 21, 31, 41, 50, 60, 70, 80])));
    expect(small?.changed.length, "the small edit is no longer one node").toBe(1);
    expect(big?.changed.length, "the big edit no longer spans more than one chunk").toBe(8);
    // Both must still be TAKEN, or the comparison is between two redraws.
    expect(worthUpdating(small!, b.nodes.length) && worthUpdating(big!, b.nodes.length)).toBe(true);

    const oneChunk = await syncsToUpdate({ ...base, title: "Renamed" });
    const twoChunks = await syncsToUpdate(withValues([11, 21, 31, 41, 50, 60, 70, 80]));
    // A relative claim, deliberately: the resolve and tag syncs around the
    // write are the same for both, so the DIFFERENCE is the extra chunk and
    // nothing else. An absolute count would break every time the surrounding
    // path gained or lost a sync.
    expect(twoChunks, "eight shapes went out in one batch — the write is not chunked").toBeGreaterThan(oneChunk);
  });
});

describe("every draw carries a chart label, so its slide readings can be attributed", () => {
  it("mints one when the caller named none", async () => {
    // `traceAbout({ item })` wraps `runDemoDeck`, which the self-test never
    // calls — `.item` appears in 0 of 203 archived rounds. So every draw a round
    // makes arrived unlabelled, and the `onSlide` readings those draws emit
    // could not be attributed to a chart. That blinded WHICH SLIDE THE CHART
    // LANDED ON, whose pooling keys on `chart`, and the report ended up printing
    // "freshly added, empty — 5 chart(s), 5 grouped = 100%" against a documented
    // 1% baseline.
    setTracing(true);
    installHost([makeSlide("s1")]);
    try {
      await insertSceneIntoSlide(buildChart(config), { tagData: JSON.stringify(config), shapesPerSync: 1 });
      const labelled = traceLog().entries.filter((e) => e.data && e.data.chart !== undefined);
      expect(labelled.length, "no draw line carried a chart label").toBeGreaterThan(0);
      expect(String(labelled[0].data!.chart)).toMatch(/^draw-\d+$/);
    } finally {
      setTracing(false);
    }
  });

  it("lets a caller-supplied name win", async () => {
    // The rescale wraps each chart in `traceAbout({ chart: "i/n" })`, and that
    // label is what makes a deck run readable in order. `traceAbout` merges
    // inner over outer, so minting unconditionally would clobber it — fixing one
    // label by destroying another.
    setTracing(true);
    installHost([makeSlide("s1")]);
    try {
      await traceAbout({ chart: "3/8" }, () =>
        insertSceneIntoSlide(buildChart(config), { tagData: JSON.stringify(config), shapesPerSync: 1 }),
      );
      const labels = new Set(
        traceLog()
          .entries.filter((e) => e.data && e.data.chart !== undefined)
          .map((e) => String(e.data!.chart)),
      );
      expect(labels.has("3/8"), "the caller label was clobbered").toBe(true);
      expect(
        [...labels].some((l) => l.startsWith("draw-")),
        "minted one anyway",
      ).toBe(false);
    } finally {
      setTracing(false);
    }
  });
});

describe("an update that redraws is labelled too", () => {
  it("mints a key when the caller named none", async () => {
    // Round 228 left six `onSlide` readings unattributed: an update that cannot
    // map its nodes to shapes REDRAWS, and those redraws issue batches carrying
    // slide readings exactly like a fresh draw. `explode a degraded picture` and
    // `edit the chart the user selected` both reach that path without naming a
    // chart.
    //
    // THE SETUP RUNS WITH TRACING OFF, and that is load-bearing. The first
    // version cleared the log with `traceLog().entries.length = 0` — but
    // `traceLog` returns a SLICE, so the clear was a no-op and the setup
    // insert's own `draw-1` stayed in the log. The test then passed on the
    // insert's label while claiming to measure the update's. `setTracing(true)`
    // empties the buffer, so drawing first and recording second is the isolation.
    installHost([makeSlide("s1")]);
    const target = await insertSceneIntoSlide(buildChart(config), {
      tagData: JSON.stringify(config),
      shapesPerSync: 1,
    });
    expect(target, "the fixture needs a chart to update").toBeTruthy();
    setTracing(true);
    try {
      await updateChartInSlide(buildChart(config), target!, { tagData: JSON.stringify(config) });
      const labelled = traceLog().entries.filter((e) => e.data && e.data.chart !== undefined);
      expect(labelled.length, "no update line carried a chart label").toBeGreaterThan(0);
      expect(String(labelled[0].data!.chart)).toMatch(/^draw-\d+$/);
    } finally {
      setTracing(false);
    }
  });

  it("lets the rescale keep its own readable label", async () => {
    // The deck-wide rescale wraps each chart in traceAbout({ chart: "i/n" }),
    // and that is what makes a deck run readable in order.
    installHost([makeSlide("s1")]);
    const target = await insertSceneIntoSlide(buildChart(config), {
      tagData: JSON.stringify(config),
      shapesPerSync: 1,
    });
    setTracing(true);
    try {
      await traceAbout({ chart: "5/8" }, () =>
        updateChartInSlide(buildChart(config), target!, { tagData: JSON.stringify(config) }),
      );
      const labels = new Set(
        traceLog()
          .entries.filter((e) => e.data && e.data.chart !== undefined)
          .map((e) => String(e.data!.chart)),
      );
      expect(labels.has("5/8"), "the rescale label was clobbered").toBe(true);
      expect(
        [...labels].some((l) => l.startsWith("draw-")),
        "minted one anyway",
      ).toBe(false);
    } finally {
      setTracing(false);
    }
  });
});

describe("an id this host has agreed to name", () => {
  it("has nothing to offer before a chart exists", async () => {
    const { namedShape, _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    // The first probe pass of a round runs before the battery draws anything.
    // That is a real state and the probe must report it, not invent an id.
    expect(namedShape()).toBe(null);
  });

  it("remembers a chart whose tag actually wrote", async () => {
    const { namedShape, _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    installHost([makeSlide("s1")]);
    const target = await insertSceneIntoSlide(buildChart(config), {
      tagData: JSON.stringify(config),
      shapesPerSync: 1,
    });
    expect(target, "the fixture needs a drawn chart").toBeTruthy();
    const named = namedShape();
    expect(named, "a tagged chart is an id the host named").toBeTruthy();
    // It must be the SHAPE the tag went on, and its own slide — the probe
    // resolves the shape through that slide's handle.
    expect(named!.shapeId).toBe(target!.shapeId);
    expect(named!.slideId).toBe(target!.slideId);
  });

  it("refuses to remember a chart that lost its config", async () => {
    const { namedShape, _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    installHost([makeSlide("s1")]);
    faults.refuseTagWritesOnResolvedProxy = true;
    try {
      const target = await insertSceneIntoSlide(buildChart(config), {
        tagData: JSON.stringify(config),
        group: false,
        shapesPerSync: 1,
      });
      // THE WRITE IS THE EVIDENCE. A chart whose tag was refused proves nothing
      // about whether this host honours its id — which is the entire property
      // the probe needs — so recording it would hand the probe a guess.
      expect(target?.lost ?? "no-config", "the fixture must actually lose the tag").toBeTruthy();
      expect(namedShape(), "an untagged chart is not evidence of a named id").toBe(null);
    } finally {
      faults.refuseTagWritesOnResolvedProxy = false;
    }
  });
});

describe("the empty-id guard", () => {
  it("refuses an id the host declined to give", async () => {
    const { namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest } =
      await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    // An empty id is WORSE than none: getItemOrNullObject("") throws or answers
    // nothing, so the probe reports `threw` or `unreadable` — a real-looking
    // answer about a bogus input, which is the one thing a diagnostic must not
    // produce. This host has been seen to decline to name a shape.
    _rememberNamedShapeForTest("s1", "");
    expect(namedShape()).toBe(null);
    _rememberNamedShapeForTest("", "shape-1");
    expect(namedShape()).toBe(null);
    _rememberNamedShapeForTest("s1", "shape-1");
    expect(namedShape()).toEqual({ slideId: "s1", shapeId: "shape-1" });
  });
});

describe("keeping the named id fresh", () => {
  it("records a grouped chart, whose tag the deferred pass writes", async () => {
    const { namedShape, _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    installHost([makeSlide("s1")]);
    // A grouped chart's tag is written by the deferred pass, not inline — but
    // the inline site records on `tagged.tagged`, which IS that pass's verdict.
    // So one record site covers both, and the grouped case is the one that
    // proves it rather than an assumption about which path ran.
    const target = await insertSceneIntoSlide(buildChart(config), {
      tagData: JSON.stringify(config),
      group: true,
    });
    expect(target, "the fixture needs a drawn chart").toBeTruthy();
    const named = namedShape();
    expect(named, "a tag written by the deferred pass is still a named id").toBeTruthy();
    expect(named!.shapeId).toBe(target!.shapeId);
    _resetNamedShapeForTest();
  });

  it("never records a slide the host would not name", async () => {
    const { namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest, VISIBLE_SLIDE_KEY } =
      await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    // `slideKeyFor` returns this when the host declined to name the slide. It is
    // the one string that looks like a slide key and cannot be resolved as one:
    // handed to `getItemOrNullObject` it yields a null object, which the probe
    // would report as a slide that never resolved — a setup failure of ours
    // wearing the costume of a finding.
    _rememberNamedShapeForTest(VISIBLE_SLIDE_KEY, "shape-1");
    expect(namedShape()).toBe(null);
  });

  it("forgets a shape a redraw has deleted", async () => {
    const { namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest, forgetNamedShape } =
      await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    _rememberNamedShapeForTest("s1", "shape-1");
    // A deleted shape's id is no longer one the host will vouch for: it answers
    // "no such shape on that slide", which is indistinguishable from refusing a
    // lookup for a shape that IS there. Rounds 242 and 243 read `unreadable` for
    // this reason and looked, twice, like office-js #2903 confirmed.
    forgetNamedShape(["shape-1"]);
    expect(namedShape()).toBe(null);
  });

  it("keeps a shape someone else's delete did not touch", async () => {
    const { namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest, forgetNamedShape } =
      await import("../src/render/powerpoint");
    _resetNamedShapeForTest();
    _rememberNamedShapeForTest("s1", "shape-1");
    // Deletes are constant during a round and almost none of them are this
    // shape. Forgetting on every delete would leave the probe permanently
    // without an id, which is the old broken state wearing a new cause.
    forgetNamedShape(["shape-2", "shape-3"]);
    expect(namedShape()).toEqual({ slideId: "s1", shapeId: "shape-1" });
    _resetNamedShapeForTest();
  });
});

describe("the named id across an update", () => {
  it("is recorded by the tag pass alone, where the inline path never runs", async () => {
    const { namedShape, _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "x" });
    const oldGroup = slide.created.find((s) => s.type === "group")!;
    // CLEARED, so only the update can put something back. `updateChartInSlide`
    // does not go through the inline recording site at all — the deferred tag
    // pass is the only place an update's write is confirmed, and for thirty-odd
    // rounds that pass recorded nothing.
    _resetNamedShapeForTest();
    // WITH `tagData`, or the update has no config to write and comes back
    // `lost: "no-config"` — which is correct behaviour and would make this test
    // pass for a reason that has nothing to do with what it is checking.
    await updateChartInSlide(
      buildChart(config),
      { slideId: "s1", shapeId: oldGroup.id, left: 33, top: 44 },
      { tagData: JSON.stringify(config) },
    );
    const named = namedShape();
    expect(named, "the update wrote a tag and nothing recorded it").toBeTruthy();
    expect(named!.slideId).toBe("s1");
    // The REDRAWN chart, not the one this update deleted. Recording the old id
    // here would hand the probe a shape that is provably gone — which is the
    // whole failure this change exists to end.
    expect(oldGroup.deleted, "this fixture is the redraw route").toBe(true);
    expect(named!.shapeId).not.toBe(oldGroup.id);
    _resetNamedShapeForTest();
  });

  it("is dropped when the redraw deletes it and the new tag is refused", async () => {
    const { namedShape, _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "x" });
    const oldGroup = slide.created.find((s) => s.type === "group")!;
    expect(namedShape()?.shapeId, "the fixture needs the old chart recorded").toBe(oldGroup.id);
    // An update with NO config to write: the redraw deletes the recorded chart
    // and comes back `lost`, so nothing re-records. Without the forget, the id
    // of a shape that is provably gone would survive — and the probe would ask
    // the host about it and be told "no such shape on that slide", which is what
    // rounds 242 and 243 reported and what read, twice, as office-js #2903.
    const next = await updateChartInSlide(buildChart(config), {
      slideId: "s1",
      shapeId: oldGroup.id,
      left: 33,
      top: 44,
    });
    expect(oldGroup.deleted, "the fixture needs the old chart actually deleted").toBe(true);
    expect(next?.lost, "the fixture needs the update to end with no config").toBeTruthy();
    expect(namedShape(), "a deleted shape is not an id this host will vouch for").toBe(null);
    _resetNamedShapeForTest();
  });
});

describe("the named id on the route a round actually takes", () => {
  it("records an in-place update, where the grouping pass never runs", async () => {
    const { namedShape, _resetNamedShapeForTest } = await import("../src/render/powerpoint");
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
    };
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => true);
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
    const t = (await listChartsInDeck()).charts[0].target;
    _resetNamedShapeForTest();
    setTracing(true);
    try {
      await updateChartsInSlides([{ scene: buildChart(cfg), target: t, opts: { tagData: JSON.stringify(cfg) } }]);
      // THE IN-PLACE ROUTE, and it is the common one: round 243 took it eleven
      // times. It writes its tag without going near the grouping pass, so a
      // record hung on that pass alone leaves the most frequent write in a round
      // unrecorded — which is how the probe came to be handed stale ids.
      expect(
        traceLog().entries.some((e) => e.message === "updated only the shapes that changed"),
        "this fixture is meant to update in place",
      ).toBe(true);
      const named = namedShape();
      expect(named, "an in-place update wrote a tag and nothing recorded it").toBeTruthy();
      expect(named!.shapeId).toBe(t.shapeId);
    } finally {
      setTracing(false);
      _resetNamedShapeForTest();
    }
  });
});

describe("the wreckage sweep and the named id", () => {
  it("forgets a shape it is sweeping away", async () => {
    const { namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest } =
      await import("../src/render/powerpoint");
    const slide = makeSlide("s1");
    installHost([slide]);
    const a = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 5, height: 5 });
    _resetNamedShapeForTest();
    _rememberNamedShapeForTest("s1", a.id);
    // These ids are wreckage from a redraw that stalled. Forgetting BEFORE the
    // attempt is deliberate: a shape this run is trying to delete is no longer
    // one the host will vouch for whether or not the sweep gets it, and the
    // asymmetry decides it — a needless forget costs the probe one honest
    // deferral, while keeping a gone shape costs a confident wrong answer.
    expect(await deleteShapesById("s1", [a.id])).toBe(1);
    expect(namedShape()).toBe(null);
    _resetNamedShapeForTest();
  });

  it("keeps one the sweep is not touching", async () => {
    const { namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest } =
      await import("../src/render/powerpoint");
    const slide = makeSlide("s1");
    installHost([slide]);
    const a = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 5, height: 5 });
    _resetNamedShapeForTest();
    _rememberNamedShapeForTest("s1", "kept-shape");
    await deleteShapesById("s1", [a.id]);
    expect(namedShape()).toEqual({ slideId: "s1", shapeId: "kept-shape" });
    _resetNamedShapeForTest();
  });
});

describe("taking the named id from a chart the deck has now", () => {
  it("takes a chart the scan actually found", async () => {
    const { refreshNamedShapeFromDeck, namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest } =
      await import("../src/render/powerpoint");
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: JSON.stringify(config) });
    // A stale record standing before the scan — the exact state rounds 242-247
    // were in. The scan must replace it, not be skipped because something was
    // already remembered.
    _rememberNamedShapeForTest("gone-slide", "gone-shape");
    const found = await refreshNamedShapeFromDeck();
    expect(found, "the deck has a chart and the scan found it").toBeTruthy();
    expect(found!.shapeId).not.toBe("gone-shape");
    expect(namedShape()).toEqual(found);
    _resetNamedShapeForTest();
  });

  it("clears a remembered id when the deck holds no chart", async () => {
    const { refreshNamedShapeFromDeck, namedShape, _resetNamedShapeForTest, _rememberNamedShapeForTest } =
      await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    _rememberNamedShapeForTest("s1", "shape-1");
    // Nothing on this deck is a chart this host has named. Saying so costs the
    // probe one honest deferral; keeping the id costs a confident wrong answer,
    // which is what four rounds reported as office-js #2903.
    expect(await refreshNamedShapeFromDeck()).toBe(null);
    expect(namedShape()).toBe(null);
    _resetNamedShapeForTest();
  });
});

describe("the group-read refusal does not outlive its batch", () => {
  it("forgets a refusal from an earlier batch", async () => {
    const { groupReadHasBeenRefused, _setGroupReadRefusedForTest } = await import("../src/render/powerpoint");
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: JSON.stringify(config) });
    const group = slide.created.find((s) => s.type === "group")!;
    // LATCHED FIRST, as a refusal in some earlier batch would leave it. Without
    // this the assertion below passes on a flag that was never set.
    _setGroupReadRefusedForTest();
    expect(groupReadHasBeenRefused()).toBe(true);
    await updateChartsInSlides([
      {
        scene: buildChart(config),
        target: { slideId: "s1", shapeId: group.id, left: 0, top: 0 },
        opts: { tagData: "x" },
      },
    ]);
    // THE POINT: a refusal from a previous batch must not decide this one. Across
    // the last 25 rounds every redraw blaming "no readable group members" fell
    // AFTER the latch fired and none before, and no in-place update happened
    // after it at all.
    expect(groupReadHasBeenRefused(), "the batch must start by forgetting").toBe(false);
  });

  it("keeps a refusal WITHIN the batch that saw it", async () => {
    const { isGroupReadRefusal } = await import("../src/render/powerpoint");
    // The refusal arrives at the SYNC, not at the property access, so it poisons
    // the batch it lands in. Re-asking for every chart in that same batch spends
    // the same exception again with nothing new to learn — which is the half of
    // the original latch that was right.
    expect(isGroupReadRefusal('{"errorLocation":"Shape.group"}')).toBe(true);
    expect(isGroupReadRefusal('{"errorLocation":"ShapeCollection.getItem"}')).toBe(false);
  });
});

describe("a group refused in the batch redraws, and says so", () => {
  it("rescues the shape, cannot rescue its group, and redraws", async () => {
    // THE SEQUENCE THE ARCHIVE SHOWS IN EVERY ROUND ON FILE, asserted as the
    // CORRECT outcome rather than as a bug to be retried away.
    //
    // A retry stood here for three rounds and never once produced an in-place
    // update, because both routes into the group are shut inside this batch:
    // the recovered proxy's `.group.shapes` answers and then its sync throws
    // (round 267, in production), and a fresh by-id lookup is the very thing
    // poisoning these syncs. What rescues the group is the NEXT batch, via the
    // reset asserted in the describe above.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: JSON.stringify(config) });
    const group = slide.created.find((s) => s.type === "group")!;
    setTracing(true);
    faults.refuseShapeById = true;
    try {
      await updateChartInSlide(
        buildChart(config),
        { slideId: "s1", shapeId: group.id, left: 0, top: 0 },
        { tagData: JSON.stringify(config) },
      );
      const said = traceLog().entries;
      // NON-VACUITY FIRST. Without this the redraw assertion below would pass on
      // a run where the shape was never recovered at all — a different failure
      // wearing the same trace.
      expect(
        said.some((e) => /^re-read recovered shapes a by-id lookup had refused/.test(e.message)),
        "the fixture needs the re-read to actually rescue the shape",
      ).toBe(true);
      const redraw = said.find((e) => /^not updating in place — redrawing instead/.test(e.message));
      expect(redraw, "the chart should have redrawn").toBeTruthy();
      // THE REASON IS NOT PINNED, deliberately. In production this redraw reads
      // "no readable group members"; under this fixture it reads "no stored
      // config to diff against", because `refuseShapeById` refuses the stored
      // config read too and the chart never reaches the group check.
      //
      // Asserting the production wording here would be asserting an artifact of
      // the fake. What this test is for is the BEHAVIOUR both paths share: a
      // chart whose shape was rescued but whose group was not still redraws,
      // rather than silently doing nothing to a chart the user is looking at.
      expect(String((redraw!.data as { why?: string }).why ?? ""), "the redraw gave no reason").not.toBe("");
    } finally {
      faults.refuseShapeById = false;
      setTracing(false);
    }
  });

  it("does not claim it re-asked for the group", async () => {
    // The retry's trace is gone, and this is what keeps it gone. Re-adding one
    // without re-reading the note in `powerpoint.ts` costs another three rounds:
    // its `got` counted queued PROXIES, which are truthy before the sync has
    // said anything, so the number read as "obtained" while meaning "asked".
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: JSON.stringify(config) });
    const group = slide.created.find((s) => s.type === "group")!;
    setTracing(true);
    faults.refuseShapeById = true;
    try {
      await updateChartInSlide(
        buildChart(config),
        { slideId: "s1", shapeId: group.id, left: 0, top: 0 },
        { tagData: JSON.stringify(config) },
      );
      const said = traceLog().entries;
      expect(said.length, "nothing was traced, so this asserts nothing").toBeGreaterThan(0);
      expect(said.some((e) => /asked a recovered chart for its group again/.test(e.message))).toBe(false);
    } finally {
      faults.refuseShapeById = false;
      setTracing(false);
    }
  });
});

/**
 * `addSlideForChart` — the slide the slow-insert offer puts a chart on.
 *
 * Small, and worth its own block because BOTH its answers are load-bearing. The
 * id sends the insert to a blank slide; `null` sends it back to the slide the
 * user picked. This host drops `slides.add()` calls, so `null` is not a
 * defensive branch nobody reaches — it is a Tuesday.
 */
describe("adding the slide a slow insert is offered", () => {
  it("answers the new slide's id, and grows the deck by exactly one", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const before = deck.length;
    const id = await addSlideForChart();
    expect(id, "no id came back for a slide that was added").toBeTruthy();
    expect(deck.length - before, "the deck did not grow by one").toBe(1);
    // The id has to name a slide that EXISTS, or the insert addresses nothing.
    expect(deck.some((s) => s.id === id)).toBe(true);
  });

  it("answers null when every add is dropped, rather than an id for no slide", async () => {
    // The failure this host really has. `addSlides` retries; when the retries
    // are defeated too, the honest answer is "there is no slide" — the pane then
    // inserts where the user asked. An id for a slide that was never created
    // would send the chart nowhere and report success.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const before = deck.length;
    faults.swallowAdds = ADDS_TO_DEFEAT_ONE_SLIDE;
    try {
      expect(await addSlideForChart()).toBeNull();
      expect(deck.length, "a slide appeared despite every add being dropped").toBe(before);
    } finally {
      faults.swallowAdds = 0;
    }
  });

  it("draws onto a freshly-added slide on a host that refuses getItem for one", async () => {
    /**
     * THE SHIPPED BUG, AND THE FIX, IN ONE TEST.
     *
     * `offerOwnSlide` ships. It adds a slide and passes that id into
     * `insertSceneIntoSlide`, which called `slides.getItem(id)` — and this host
     * refuses that for a slide the run just added. The user accepting "put it
     * on a slide of its own" got GeneralException at `SlideCollection.getItem`
     * on the FIRST batch: an error and an empty slide.
     *
     * FOUR ATTEMPTS TRIED TO VALIDATE THE ID and every one was ruled out by a
     * real round. None asked whether we should be using one. The archive had
     * already answered that, twice over, for 342 rounds — over the 318 rounds
     * where both probes replied cleanly:
     *
     *     by-id THREW while by-index WORKED   198   (62%)
     *     by-index threw while by-id worked     3   (0.9%)
     *
     * and `shape-add-positional-slide-proxy` says the conclusion outright:
     * "every write path that names a freshly-added slide by id needs an index
     * instead."
     *
     * So the insert resolves the id to an INDEX through the deck listing — the
     * one place the id does resolve — and targets `getItemAt`. Arming
     * `refuseGetItemOnNewSlide` makes this test fail against every version of
     * the code that reached a user.
     */
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.refuseGetItemOnNewSlide = true;
    try {
      const slideId = await addSlideForChart();
      expect(slideId, "no slide id came back").toBeTruthy();
      const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
      const target = await insertSceneIntoSlide(buildChart(cfg), {
        slideId: slideId!,
        tagData: JSON.stringify(cfg),
      });
      expect(target, "the insert produced no chart — the own-slide offer is still broken").toBeTruthy();
      /**
       * AND THE SHAPES ARE ON THAT SLIDE — asserted by COUNTING them, not by
       * reading the id back.
       *
       * The id assertion alone is vacuous here and the mutation run proved it:
       * reverting the thunk to `getTargetSlide(context, opts.slideId)` — the
       * shipped bug — still passed, because `getTargetSlide` CATCHES the
       * refusal and quietly falls back to the selected slide. The draw went to
       * the wrong slide while the reported id, read from a different handle,
       * stayed right. A user would get a chart on the slide they were already
       * looking at and no error at all.
       *
       * Where the ink landed is the only thing that cannot lie.
       */
      expect(target?.slideId, "the chart reported a different slide than the one added").toBe(deck[1].id);
      expect(deck[1].created.length, "nothing was drawn on the slide that was added for it").toBeGreaterThan(0);
      expect(
        deck[0].created.length,
        "the draw fell back to the slide the user was already on — the offer silently did nothing",
      ).toBe(0);
    } finally {
      faults.refuseGetItemOnNewSlide = false;
    }
  });

  it("asks the deck what it still holds when the draw onto an added slide fails", async () => {
    /**
     * THE FORK ROUND 370 COULD NOT SETTLE. At 4:3 the add reports `landed=1`,
     * the listing puts the slide at index 7 of 8, `getItemAt(7)` throws, and
     * the next scenario opens on a deck of SEVEN. Two readings fit equally
     * well: the slide is GONE — an acknowledged-and-absent `slides.add()`,
     * which this archive has receipts for — or it is THERE and the host will
     * not hand it over yet. Those want completely different fixes.
     *
     * Deck length at the moment of failure separates them, so the failure path
     * now asks. BOTH branches are exercised here: a classifier that cannot
     * produce both of its answers has not been tested, only executed.
     */
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const seen: { message: string; data?: Record<string, unknown> }[] = [];
    const found = () => seen.find((e) => e.message.startsWith("the draw failed — asking the deck"));
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    setTracing(true);
    onTrace((e) => seen.push({ message: e.message, data: e.data }));
    try {
      const slideId = await addSlideForChart();
      expect(slideId).toBeTruthy();
      // The index resolution is the next sync and the first render batch
      // follows it. COMPUTED, not hardcoded — the tests one screen up warn that
      // a wrong index here silently stops testing anything.
      faults.failSyncOn = trips.syncs + 2;
      let threw = false;
      await insertSceneIntoSlide(buildChart(cfg), { slideId: slideId!, tagData: JSON.stringify(cfg) }).catch(() => {
        threw = true;
      });
      expect(threw, "the draw did not fail, so nothing was diagnosed").toBe(true);

      // CASE ONE — the slide is still there under the id we hold.
      const asked = found();
      expect(asked, "a draw onto an added slide failed and the deck was never asked about it").toBeTruthy();
      expect(asked!.data!.index).toBe(1);
      expect(asked!.data!.deckSlides).toBe(2);
      expect(asked!.data!.stillListedUnderTheSameId).toBe(true);
      expect(asked!.data!.reading).toBe("there, listed under the same id, and refused anyway");
      /**
       * AND IT DOES NOT FALL BACK HERE. The slide is present, so this is an
       * ordinary host failure and must surface. Recovering from ANY failure
       * would quietly redraw the chart elsewhere and swallow a real error.
       *
       * ASSERTED ON THE TRACE, not on where the ink landed, and the difference
       * is why the mutation run needed two attempts. With the guard relaxed the
       * retry happens and then fails too, so the deck looks identical either
       * way — `threw`, `deck[0]` and `deck[1]` are all unchanged. The fallback
       * announces itself before retrying, so the announcement is the only thing
       * that separates "did not fall back" from "fell back and got nowhere".
       */
      expect(
        seen.some((e) => e.message.includes("drawing on the visible slide instead")),
        "fell back on a failure that had nothing to do with a missing slide",
      ).toBe(false);
      /**
       * A SURVIVING MUTANT, RECORDED RATHER THAN HIDDEN. Dropping
       * `reading === "gone"` from the guard in `insertSceneIntoSlideInner`
       * survives this whole file, and it took two wrong explanations to find
       * out why: the fault used here rejects with a non-object, so the
       * `typeof err === "object"` arm filters it out and the flag is never set
       * whatever the reading says. The assertion above is therefore true for a
       * reason other than the one it names.
       *
       * Killing it needs a fault that fails a draw with an OBJECT error while
       * the slide is still listed and nothing has committed, and no fault in
       * `office-host.ts` produces that combination today. Left as a gap on the
       * record instead of a green tick that means less than it looks like: the
       * guard is right, and it is only partly tested.
       */
    } finally {
      faults.failSyncOn = 0;
      onTrace(undefined);
      setTracing(false);
    }

    // CASE TWO — the same failure with the slide no longer in the deck, which
    // is the reading that would mean the add was acknowledged and lost.
    seen.length = 0;
    setTracing(true);
    onTrace((e) => seen.push({ message: e.message, data: e.data }));
    try {
      const slideId = await addSlideForChart();
      /**
       * NO ARMED SYNC FAILURE HERE, and that is not an oversight. A slide that
       * is gone makes `getItemAt(index)` fail by itself, which is the whole
       * point — arming one as well put the armed index on the DIAGNOSTIC's own
       * listing sync, so it threw, the catch swallowed it, and the measurement
       * never happened. The first draft did exactly that and reported nothing.
       *
       * IT HAS TO VANISH IN THE WINDOW, not before it. Removing the slide up
       * front makes the index resolution miss it, `index` stays -1, and the
       * diagnostic is never reached — which is what the first draft of this
       * test did, and it proved nothing.
       *
       * `onPhase("queue")` fires after the index has resolved and before the
       * first batch is issued, which is precisely the gap round 370 describes:
       * listed at index 7 of 8 one moment, a deck of 7 the next.
       */
      let removed: FakeSlide | undefined;
      await insertSceneIntoSlide(buildChart(cfg), { slideId: slideId!, tagData: JSON.stringify(cfg) }, (phase) => {
        if (phase === "queue" && !removed) removed = deck.pop();
      }).catch(() => {});
      expect(removed?.id, "the wrong slide was removed, or none was").toBe(slideId);
      const asked = found();
      expect(asked, "the deck was not asked").toBeTruthy();
      expect(asked!.data!.indexInRange, "an out-of-range index was reported as in range").toBe(false);
      expect(asked!.data!.reading).toBe("gone — the index is past the end of the deck");
    } finally {
      faults.failSyncOn = 0;
      onTrace(undefined);
      setTracing(false);
    }
  });

  it("draws on the visible slide when the slide added for the chart has vanished", async () => {
    /**
     * WHAT A USER GETS FOR ACCEPTING THE OFFER. `offerOwnSlide` adds a slide
     * and hands its id to the insert. At 4:3 that slide is added, listed,
     * listed AGAIN by the index resolution, and gone by the next sync of the
     * same context — before one shape is issued. Round 371 measured it twice.
     *
     * Until now that was `GeneralException` and, on about half of those runs, a
     * dead PowerPoint: `a big chart on a slide of its own` sits at 500 deaths
     * per 1000 runs, the worst number in the suite. Now the chart lands on the
     * slide the user was already looking at.
     *
     * THE INK IS THE ASSERTION, not the absence of a throw. A version that
     * swallowed the error and returned null would pass a "does not reject"
     * test while drawing nothing — the vacuous shape this file has been caught
     * by twice already.
     */
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const slideId = await addSlideForChart();
    expect(slideId, "no slide id came back").toBeTruthy();
    const added = deck.length - 1;
    const target = await insertSceneIntoSlide(
      buildChart(cfg),
      { slideId: slideId!, tagData: JSON.stringify(cfg) },
      (p) => {
        // Gone between the index resolution and the first batch — the window
        // round 371 puts it in.
        if (p === "queue") deck.splice(added, 1);
      },
    );
    expect(target, "the offer threw instead of falling back — this is what killed the host").toBeTruthy();
    expect(target!.slideId, "the target names a slide that is not in the deck").toBe(deck[0].id);
    expect(deck[0].created.length, "the fallback reported success and drew nothing").toBeGreaterThan(0);
  });

  it("falls back even when the vanished slide had already taken shapes with it", async () => {
    /**
     * THE GUARD THAT WAS REMOVED, AND WHY ITS ABSENCE IS SAFE.
     *
     * An earlier version refused to fall back once any shape had been
     * committed, to avoid re-running a half-drawn scene and drawing it twice.
     * That guard was broken — it counted `onBatch`'s `sending` argument, which
     * is reported before the sync, so on the host it was never zero and the
     * fallback never fired at all (round 372). It was also unnecessary, which
     * is the more interesting half.
     *
     * Shapes live on a slide. This path is only reached when that slide is
     * GONE, so anything already drawn went with it — there is nothing left to
     * draw twice. The double-draw risk exists only while the slide SURVIVES,
     * and that case never reaches here because the reading is not "gone".
     *
     * So a chart that failed half-drawn onto a vanishing slide SHOULD be
     * re-run, and this asserts the outcome a user cares about: exactly one
     * chart, on the slide that still exists.
     */
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const slideId = await addSlideForChart();
    const added = deck.length - 1;
    let vanished: FakeSlide | undefined;
    const target = await insertSceneIntoSlide(
      buildChart(cfg),
      { slideId: slideId!, tagData: JSON.stringify(cfg) },
      (p, detail) => {
        // Vanish only AFTER a batch has gone out, so the scene is part-drawn.
        if (p === "commit" && Number(String(detail).split(" ")[0]) > 0 && !vanished)
          vanished = deck.splice(added, 1)[0];
      },
    );
    expect(vanished?.id, "the slide never vanished, so this tested nothing").toBe(slideId);
    expect(target, "a part-drawn chart on a vanished slide was not recovered").toBeTruthy();
    expect(target!.slideId).toBe(deck[0].id);
    // ONE chart on the surviving slide. The shapes that went out before the
    // slide disappeared went with it; the fake keeps them on the detached
    // object, which is why the assertion is about `deck[0]` and not a total.
    expect(deck[0].created.length, "the recovered chart drew nothing").toBeGreaterThan(0);
    expect(deck.length, "the vanished slide came back").toBe(1);
  });

  it("adds nothing at all when the deck will not say what is already in it", async () => {
    /**
     * THE ANSWER IS STILL NULL AND THE SLIDE IS NO LONGER LEFT BEHIND.
     *
     * This used to assert the opposite of its second half: a slide landed, the
     * host refused its id, and `null` came back — correct, because a fabricated
     * id would send the insert to a slide that does not answer to it and report
     * success. That half is unchanged and is what matters.
     *
     * What changed on 2026-09-03 is WHEN the deck is read. The id now comes
     * from diffing the deck's own listing before and after the add, because the
     * old positional read returned an ADD-TIME id that `getItem` refuses — see
     * `addSlideForChart`. A deck that cannot be listed BEFORE the add is a deck
     * whose new slide could never be named afterwards, so nothing is added at
     * all rather than a blank slide being left for the user to find.
     *
     * Strictly better on the same evidence: same null, no orphan.
     */
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const before = deck.length;
    faults.slideIdNeverReadable = true;
    try {
      expect(await addSlideForChart(), "an id came back from a host that gave none").toBeNull();
    } finally {
      faults.slideIdNeverReadable = false;
    }
    expect(deck.length, "a blank slide was left behind on a deck that could not be listed").toBe(before);
  });

  it("recovers when ONE sync fails, because addSlides retries", async () => {
    // Written expecting null; the code was right and the expectation was wrong.
    // A single failed sync is recoverable and `addSlides` recovers it. Kept as
    // the boundary against the case below — the difference between a host having
    // a bad moment and a host that will not do this at all.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.failSyncOn = 1;
    try {
      expect(await addSlideForChart(), "a recoverable stall lost the slide").toBeTruthy();
    } finally {
      faults.failSyncOn = 0;
    }
  });

  it("answers null when every sync fails, rather than taking the pane down", async () => {
    // This runs inside the Insert click. A throw here would abandon a chart the
    // user has already pressed for, so the catch is the point, not decoration.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    for (let i = 1; i <= 40; i++) failSyncsOn.add(i);
    try {
      expect(await addSlideForChart()).toBeNull();
    } finally {
      failSyncsOn.clear();
    }
  });

  it("answers null when there is no host at all", async () => {
    // The pane boots in browsers with no Office. `PowerPoint.run` throws on the
    // global lookup itself, before any of the above.
    const saved = (globalThis as { PowerPoint?: unknown }).PowerPoint;
    delete (globalThis as { PowerPoint?: unknown }).PowerPoint;
    try {
      expect(await addSlideForChart()).toBeNull();
    } finally {
      (globalThis as { PowerPoint?: unknown }).PowerPoint = saved;
    }
  });
});

/**
 * WHAT AN OLD HOST ACTUALLY DRAWS, which until 2026-08-30 nothing asserted.
 *
 * `Shape.rotation` is PowerPointApi 1.10 — recent, so most desktop Office is
 * below it and volume-licensed builds are below it permanently. Both of the
 * things that need rotation failed there, and neither failed loudly:
 *
 *   - a diagonal line was drawn as a rectangle at the right length and midpoint
 *     and then left LYING FLAT, so line charts, scatter trend lines and radar
 *     outlines came out as stacks of horizontal bars;
 *   - a pie wedge was skipped entirely, so pie, doughnut, sunburst and gauge
 *     inserted with a title, a legend, labels and no slices.
 *
 * `docs/MANUAL.md` said "missing capabilities degrade gracefully — charts still
 * insert". The charts inserted; they were wrong.
 */
describe("a host that cannot rotate", () => {
  const noRotation = (v: string) => v !== "1.10";

  it("draws a diagonal as a real diagonal, not a rectangle lying flat", async () => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, noRotation);
    await insertSceneIntoSlide(
      buildChart({
        ...config,
        kind: "line",
        data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [1, 9, 3] }] },
      }),
      {},
    );
    // A line-chart segment on this host must be a LINE shape (`addLine`, or the
    // `lineInverse` geometry for an up-right slope) — both draw a true diagonal
    // without rotating anything, which is what this branch already did for
    // dashed lines and simply never did for solid ones.
    const flatBars = slide.created.filter(
      (s) => s.name?.startsWith("line-") && s.geo === "rectangle" && s.rotation === undefined,
    );
    expect(flatBars, `${flatBars.length} line segments drawn as unrotated rectangles`).toEqual([]);
    const segments = slide.created.filter((s) => s.name?.startsWith("line-"));
    expect(segments.length, "no line segments at all").toBeGreaterThan(0);
  });

  it("still draws a proper rotated diagonal where the host can rotate", async () => {
    // The fix must not cost the good host its precise geometry: with 1.10 the
    // solid diagonal stays a rotated rectangle, which is what carries an exact
    // stroke weight.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(
      buildChart({
        ...config,
        kind: "line",
        data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [1, 9, 3] }] },
      }),
      {},
    );
    const rotated = slide.created.filter((s) => s.name?.startsWith("line-") && typeof s.rotation === "number");
    expect(rotated.length, "the rotating host stopped rotating its diagonals").toBeGreaterThan(0);
  });

  it("omits an arrowhead it cannot aim, rather than pointing it the wrong way", async () => {
    /**
     * The old code applied `arrowheadBox`'s rotation-derived OFFSET and then
     * skipped the rotation, so the triangle pointed up AND sat displaced by
     * `size` in two axes. Its comment said "stays axis-aligned", which is the
     * half of it that sounds harmless.
     *
     * A difference arrow pointing up whatever it measured is a bridge claiming
     * growth where the number fell — wrong output, stated confidently, in
     * silence. Skipping is the honest outcome and it now traces.
     */
    setTracing(true);
    const slide = makeSlide("s1");
    installHost([slide], [], slide, noRotation);
    const from = traceLog().entries.length;
    await insertSceneIntoSlide(
      buildChart({
        ...config,
        decorations: { cagr: { from: 0, to: 2 }, segmentLabels: true },
      }),
      {},
    );
    const heads = slide.created.filter((s) => s.geo === "triangle");
    expect(heads, `drew ${heads.length} arrowhead(s) it could not aim`).toEqual([]);
    const said = traceLog()
      .entries.slice(from)
      .some((e) => e.message === "cannot draw an arrowhead on this host");
    expect(said, "omitted the arrow and said nothing anywhere").toBe(true);
  });

  it("says so when it cannot draw a pie, instead of inserting an empty one", async () => {
    /**
     * A wedge cannot be built from axis-aligned shapes — that is why the fan of
     * rotated triangles exists — so unlike the diagonal there is no rotation-free
     * remedy here, and choosing one (refuse the insert, or fall back to the
     * picture path) is a product decision recorded in docs/BACKLOG.md.
     *
     * What is NOT a decision is whether the failure is silent. It was.
     */
    setTracing(true);
    const slide = makeSlide("s1");
    installHost([slide], [], slide, noRotation);
    const from = traceLog().entries.length;
    await insertSceneIntoSlide(
      buildChart({
        ...config,
        kind: "pie",
        data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
      }),
      {},
    );
    const said = traceLog()
      .entries.slice(from)
      .some((e) => e.message === "cannot draw a pie wedge on this host");
    expect(said, "a pie inserted with no slices and nothing anywhere said why").toBe(true);
  });
});

/**
 * The sync counter, and the one property that makes it worth reading.
 *
 * office-js#6329 reports that on PowerPoint for the web every `context.sync()`
 * forces a full presentation save, read-only syncs included. If that holds
 * here, the load this add-in puts on the save channel is proportional to sync
 * count — and every crash figure this repo owns is indexed by shapes instead,
 * because nothing has ever counted syncs.
 */
describe("counting what the round asks of the host", () => {
  const source = readFileSync("src/render/powerpoint.ts", "utf8");

  it("lets the value a batch returns ride out through one last sync", async () => {
    /**
     * THE BUG THIS WHOLE BLOCK NEARLY SHIPPED, and the most expensive mistake
     * in this file's history.
     *
     * `sync` is declared `sync<T>(passThroughValue?: T): Promise<T>`, and that
     * parameter is how `PowerPoint.run` returns anything at all: Office.js
     * ends every batch with `.then(r => ctx.sync(r))` and
     * `ClientRequestContext.sync` resolves with exactly what it was handed.
     *
     * The first `ppRun` took no parameters and called `sync()`. So that final
     * auto-sync resolved undefined, and EVERY host run in the build resolved
     * undefined with it: `slideCount()` gave nothing, `deckSlides` read
     * `unreadable`, the self-test's arithmetic became `deck grew by NaN`,
     * `slideSize()`'s live rung threw and fell through to the saved file so the
     * pane reported the deck's old profile, and the probe harness aborted a
     * round on `Cannot read properties of undefined (reading 'answer')`.
     * Rounds 360 and 361 measured none of what they claimed; the crossed-deck
     * experiment was written up as void and a healthy document was blamed.
     *
     * NOTHING IN THIS SUITE COULD SEE IT. All 197 `.sync()` calls in this repo
     * are argless — the only caller that passes a value lives inside Office.js
     * — and the shared fake's `run` returns the callback's value directly with
     * no final sync at all. 3,831 tests passed with the add-in fully broken.
     *
     * So this test brings its own host, modelled on the real one in the single
     * respect that matters: the batch's result comes back THROUGH sync.
     */
    const seen: unknown[] = [];
    const ctx = {
      sync: async <T>(passThroughValue?: T) => {
        seen.push(passThroughValue);
        return passThroughValue;
      },
    };
    vi.stubGlobal("PowerPoint", {
      /**
       * Office.js's `_runCommon`, reduced to its value-carrying role — and the
       * LOOKUP ORDER is load-bearing, which cost this test its first draft.
       *
       * Written as `ctx.sync(await cb(ctx))` it passes against the bug. In
       * `obj.method(arg)` JavaScript resolves the member reference BEFORE
       * evaluating the argument, so that form captures the sync that existed
       * before the batch ran — the unpatched one — and the wrapper under test
       * is never invoked at all. The real host resolves it inside a `.then`
       * callback, after the batch and after any patch. So must this.
       */
      run: async <T>(cb: (c: typeof ctx) => Promise<T>) => {
        const runBodyResult = await cb(ctx);
        return ctx.sync(runBodyResult);
      },
    });

    // Exercised through the real wrapper rather than a copy of it: `ppRun` is
    // module-private, and a test that reimplemented it would pass while the
    // shipped one stayed broken — which is precisely how this got out.
    const { _ppRunForTest } = (await import("../src/render/powerpoint")) as unknown as {
      _ppRunForTest: <T>(fn: (c: unknown) => Promise<T>) => Promise<T>;
    };
    const got = await _ppRunForTest(async (c) => {
      await (c as typeof ctx).sync();
      return 42;
    });

    expect(got, "the batch's return value did not survive the run — see the docstring").toBe(42);
    // And the LAST sync is the one carrying it, which is the mechanism rather
    // than the symptom: an inner argless sync must not be mistaken for the
    // auto-sync that matters.
    expect(seen[seen.length - 1], "the final auto-sync was handed nothing to carry").toBe(42);
  });

  it("routes every host run through the counting wrapper", () => {
    /**
     * THE ASSERTION THAT MAKES THE INSTRUMENT TRUSTWORTHY, and the reason it is
     * a source scan rather than a behavioural test.
     *
     * A sync counter that misses a call site does not read LOW — it reads
     * WRONG, and it would be used to exonerate the very path it cannot see. The
     * 37 call sites were converted mechanically, and this repo has shipped a
     * fix-with-one-call-site-left-behind four times.
     *
     * So completeness is asserted directly: exactly ONE bare `PowerPoint.run(`
     * may exist in this file, the one inside `ppRun` itself. A new call site
     * added by hand fails here on the day it is written, which is the only day
     * it is cheap to fix.
     */
    const bare = source.match(/PowerPoint\.run\(/g) ?? [];
    expect(bare, "a host run bypasses the sync counter — see ppRun").toHaveLength(1);
    const inWrapper = /function ppRun<T>\([\s\S]{0,400}?PowerPoint\.run\(/.test(source);
    expect(inWrapper, "the one remaining PowerPoint.run is not the wrapper's own").toBe(true);
    // And the wrapper must actually be used: 37 sites were converted, so a
    // number far below that means the replacement was reverted wholesale.
    expect((source.match(/\bppRun\(/g) ?? []).length).toBeGreaterThanOrEqual(37);
  });

  it("counts on the context the callback is handed, not on a copy of it", () => {
    /**
     * All 90 syncs in this file are `context.sync()` or `retryContext.sync()`,
     * both handed to the callback by `ppRun`. Patching the object BEFORE the
     * callback runs is what makes the count complete by construction — a
     * counter incremented at the call sites instead would need 90 edits and
     * would drift on the next one written.
     *
     * Two things would silently break it, and both are cheap to pin: taking a
     * reference to `sync` before the patch, and patching after `fn` is invoked.
     */
    const wrapper = source.slice(source.indexOf("function ppRun<T>"));
    const body = wrapper.slice(0, wrapper.indexOf("\n}\n"));
    expect(body, "the context is handed to the callback before its sync is patched").toMatch(
      /context\.sync = [\s\S]*?return fn\(context\)/,
    );
    expect(body, "the counter stopped incrementing").toMatch(/SYNCS\+\+/);
    // Nothing may destructure `sync` off a context anywhere, or that path
    // escapes the patch entirely.
    expect(source, "a destructured sync would bypass the counter").not.toMatch(/const\s*\{[^}]*\bsync\b[^}]*\}\s*=/);
  });

  it("actually counts, and counts each sync exactly once", async () => {
    /**
     * THE HALF A SOURCE SCAN CANNOT REACH. Everything above proves the wiring
     * is present; only this proves the number moves, and a counter that reads
     * zero forever would pass every assertion in this file but one.
     *
     * Two failure modes are pinned rather than one, because the second is the
     * one that would go unnoticed. Under-counting reads as a quiet round;
     * DOUBLE-counting reads as a busy one, and this suite's host hands the same
     * context object to every run — so a wrapper that wraps itself would show
     * sync load climbing across a session that was flat. `ppRun` marks the
     * function it installs and refuses to wrap it twice; this is that guard.
     */
    installHost([makeSlide("s1")]);
    resetSyncCount();
    expect(syncsSoFar(), "the counter did not start at zero").toBe(0);

    await slideCount();
    const afterOne = syncsSoFar();
    expect(afterOne, "a host call made no sync the counter could see").toBeGreaterThan(0);

    await slideCount();
    const afterTwo = syncsSoFar();
    // The SAME call again must cost the same again — not more, which is what a
    // wrapper wrapping a wrapper produces on the second run through a reused
    // context.
    expect(afterTwo - afterOne, "the second identical call counted a different number of syncs").toBe(afterOne);

    resetSyncCount();
    expect(syncsSoFar(), "the round-start reset does not reset").toBe(0);
  });

  it("names each stage of the deck scan before the sync that could end it", async () => {
    /**
     * THE HOLE THE 41 CRASH RECORDS ARE. Every one ends at the pane's
     * `collecting deck evidence — scanning` line, emitted one statement BEFORE
     * `listChartsInDeck` is called — and nothing from inside the scan has ever
     * appeared after it. `step` traces only when something THROWS, so a host
     * that simply stops answering leaves no trace, and all those records can
     * say is "it died somewhere in the scan".
     *
     * "Somewhere" spans a slide count, a page of queued loads, and two syncs
     * with very different payloads. Each stage announces itself first now.
     *
     * DRIVEN BY A FAILING SYNC rather than asserted from source, because
     * ordering is the entire property: a line traced AFTER its sync is exactly
     * as useless as no line, and reads identically in a healthy round.
     */
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    const from = traceLog().entries.length;
    failSyncsOn.add(2);
    await listChartsInDeck({ withInventory: true });
    failSyncsOn.clear();

    const said = traceLog()
      .entries.slice(from)
      .map((e) => e.message);
    expect(said, "a scan that died named no stage — the crash record is still a hole").toContain(
      "deck scan — asking how many slides",
    );
    expect(said, "the page settle was not announced before the sync that swallowed it").toContain(
      "deck scan — settling a page of slides",
    );
    // Ordering is the property under test: the announcement must PRECEDE the
    // stage it announces, not merely exist somewhere in the log.
    const count = said.indexOf("deck scan — asking how many slides");
    const page = said.indexOf("deck scan — settling a page of slides");
    expect(count, "the slide count was announced after the page it precedes").toBeLessThan(page);
  });

  it("zeroes the count per round and banks it with the outcome", () => {
    /**
     * A count that carries over from the pane's earlier life belongs to no
     * round, and the driver reuses a pane across rounds far more often than it
     * reloads one. The reset therefore sits with the trace mark, which exists
     * for exactly the same reason.
     *
     * And the count has to reach BOTH files. `lastRunLog` is assembled after
     * the deck scan, so a crashed round never writes `syncs` — but 41 of the
     * last 60 4:3 crashes end inside that scan, which is precisely the
     * population the number is for. Tracing it on the scan line puts it in the
     * crash record too, where `traceLog` is a local array read and survives a
     * host that has stopped answering.
     */
    const pane = readFileSync("src/taskpane/app.ts", "utf8");
    expect(pane, "the counter is no longer zeroed per round").toMatch(/resetSyncCount\(\);\s*\n\s*const traceFrom/);
    expect(pane, "the round file stopped carrying its sync count").toMatch(/syncs: syncsSoFar\(\)/);
    expect(pane, "a crashed round can no longer report a sync count").toMatch(
      /collecting deck evidence — scanning",\s*\{[\s\S]{0,200}?syncs: syncsSoFar\(\)/,
    );
  });
});
