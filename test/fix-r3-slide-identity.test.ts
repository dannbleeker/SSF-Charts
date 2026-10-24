// @vitest-environment jsdom
/**
 * Slide and chart identity across the deck — the regressions behind stream
 * r3-slide-identity. Each block names the finding it pins and why the older
 * tests could not see it.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  CHART_PARTS_TAG,
  CHART_TAG,
  _rememberNamedShapeForTest,
  _resetNamedShapeForTest,
  _resetSlideSizeCache,
  _setSlideSizeTimeoutForTest,
  addSlideForChart,
  insertSceneIntoSlide,
  listChartsInDeck,
  loadChartFromSelection,
  namedShape,
  refreshNamedShapeFromDeck,
  scanGap,
  scanIsComplete,
  slideSize,
  updateChartInSlide,
} from "../src/render/powerpoint";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { placeChart } from "../src/core/placement";
import { onTrace, setTracing } from "../src/core/trace";
import {
  failSyncsOn,
  faults,
  hostSlideSize,
  installHost,
  makeShape,
  makeSlide,
  trips,
  userClicksShape,
  type FakeSlide,
} from "./helpers/office-host";

afterEach(() => vi.unstubAllGlobals());

/** A trace listener that collects what a call said, and puts the tracer back. */
function listen(): { lines: { message: string; data?: Record<string, unknown> }[]; stop: () => void } {
  const lines: { message: string; data?: Record<string, unknown> }[] = [];
  setTracing(true);
  onTrace((e) => lines.push({ message: e.message, data: e.data as Record<string, unknown> | undefined }));
  return {
    lines,
    stop: () => {
      onTrace(undefined);
      setTracing(false);
    },
  };
}

describe("r1-27 — a slide-size change made while the pane is open", () => {
  beforeEach(() => _resetSlideSizeCache());
  afterEach(() => {
    _resetSlideSizeCache();
    _setSlideSizeTimeoutForTest(1_500);
    faults.wedgeAfterSyncs = null;
  });

  it("is noticed by the next call, with no refresh requested", async () => {
    // Design > Slide Size, 16:9 to 4:3, with the pane left open. Nothing in the
    // product passes `refresh: true` (placement, the slide-swap deck and the
    // own-slide deck all call `slideSize()` bare), so a cache that only
    // re-reads on request keeps answering 960 until the pane reloads.
    installHost([makeSlide("s1")]);
    expect((await slideSize()).width).toBe(960);
    hostSlideSize.cx = 9144000; // 720pt
    const size = await slideSize();
    expect(size, "answered from the cache after the deck was resized").toEqual({
      width: 720,
      height: 540,
      source: "pageSetup",
    });
  });

  it("puts the next chart on the new slide rather than 177pt past its edge", async () => {
    // The consequence, not the number. A 480x300 chart sits at (60, 90) on a
    // 4:3 slide that PowerPoint has just rescaled to 720pt wide; the next
    // insert is placed from whatever width `slideSize()` reports.
    installHost([makeSlide("s1")]);
    await slideSize();
    hostSlideSize.cx = 9144000;
    const slide = await slideSize();
    const placed = placeChart(
      [{ left: 45, top: 135, width: 360, height: 225 }],
      { width: 480, height: 300 },
      { left: 60, top: 90 },
      { left: 74, top: 104 },
      slide,
    );
    // Against the deck's TRUE width. Comparing with `slide.width` would pass on
    // the stale answer too: 897 is inside the 960 it still believes in.
    const truth = hostSlideSize.cx / 12700;
    expect(truth).toBe(720);
    expect(placed.left + placed.width, "the new chart runs off a 720pt slide").toBeLessThanOrEqual(truth);
  });

  it("says so when the size moved, and says nothing when it did not", async () => {
    installHost([makeSlide("s1")]);
    await slideSize();
    const heard = listen();
    try {
      await slideSize();
      expect(
        heard.lines.filter((l) => /^slide size/.test(l.message)),
        "traced a steady size on every call",
      ).toEqual([]);
      hostSlideSize.cx = 9144000;
      await slideSize();
      const moved = heard.lines.find((l) => l.message === "slide size changed since it was read");
      expect(moved, "a resize went by unrecorded").toBeTruthy();
      expect((moved!.data!.was as { width: number }).width).toBe(960);
      expect((moved!.data!.now as { width: number }).width).toBe(720);
    } finally {
      heard.stop();
    }
  });

  it("keeps the last measurement when the cheap rung cannot answer this time", async () => {
    // The re-ask is only worth its cost if a host that has gone quiet cannot
    // take the answer away: the bound expires and the cache stands.
    installHost([makeSlide("s1")]);
    hostSlideSize.cx = 9144000;
    expect((await slideSize()).width).toBe(720);
    _setSlideSizeTimeoutForTest(30);
    faults.wedgeAfterSyncs = 0;
    const at = Date.now();
    const kept = await slideSize();
    expect(kept).toEqual({ width: 720, height: 540, source: "pageSetup" });
    expect(Date.now() - at, "waited past the slide-size budget").toBeLessThan(1_000);
  });

  it("does not re-export the slide for a measurement it already holds", async () => {
    // Falling through to rung 2 on every call would charge a whole-file round
    // trip per insert. With rung 1 unavailable the exported reading stands.
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.10");
    hostSlideSize.cx = 9144000;
    expect((await slideSize()).source).toBe("exportedSlide");
    hostSlideSize.cx = 12192000;
    const before = trips.syncs;
    const again = await slideSize();
    expect(again.width, "re-exported the slide instead of keeping what it held").toBe(720);
    expect(trips.syncs - before, "synced although rung 1 cannot answer on this host").toBe(0);
  });

  it("upgrades an exported reading as soon as the direct read can answer", async () => {
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.10");
    hostSlideSize.cx = 9144000;
    expect((await slideSize()).source).toBe("exportedSlide");
    // The host recovers (or a later call finds rung 1 answering): a direct read
    // outranks a file the host exported, and reflects a resize made since.
    installHost([makeSlide("s1")]);
    hostSlideSize.cx = 12192000;
    expect(await slideSize()).toEqual({ width: 960, height: 540, source: "pageSetup" });
  });
});

const deckOf = (n: number): FakeSlide[] => Array.from({ length: n }, (_, i) => makeSlide(`s${i + 1}`));
const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };

describe("r1-49 — one add that lands more slides than it was asked for", () => {
  afterEach(() => {
    faults.addLandsExtra = 0;
    faults.renumbersOnAdd = false;
    faults.addsAtFront = false;
    faults.failSyncOn = 0;
  });

  it("uses the last of the slides that landed and takes the twin back out", async () => {
    // The state `addScratchSlide`'s branch for un-nameable landings exists for,
    // and which this path had never been given: `addSlides` counts have = 2,
    // hands out one thunk and says `added`, the id diff then names TWO ids, and
    // the function returned null with both blank slides still in the deck.
    const deck = deckOf(2);
    installHost(deck);
    faults.addLandsExtra = 1;
    const id = await addSlideForChart();
    expect(id, "two slides landed and no id came back").toBeTruthy();
    expect(deck.length, "the twin the host added was left in the deck").toBe(3);
    expect(deck[deck.length - 1].id, "the id names a slide other than the one kept").toBe(id);
  });

  it("takes back every extra when the host lands more than two", async () => {
    const deck = deckOf(1);
    installHost(deck);
    faults.addLandsExtra = 2;
    const id = await addSlideForChart();
    expect(id).toBeTruthy();
    expect(deck.length, "blank slides were left behind").toBe(2);
    expect(deck[1].id).toBe(id);
  });

  it("draws the chart on the kept slide, and on no other", async () => {
    // The ink, not the id: an id that names a slide the draw never reaches
    // would pass a bare `toBeTruthy`.
    const deck = deckOf(2);
    installHost(deck);
    faults.addLandsExtra = 1;
    const id = await addSlideForChart();
    const target = await insertSceneIntoSlide(buildChart(cfg), { slideId: id!, tagData: JSON.stringify(cfg) });
    expect(target?.slideId).toBe(id);
    expect(deck[deck.length - 1].created.length, "nothing was drawn on the kept slide").toBeGreaterThan(0);
    expect(deck[0].created.length + deck[1].created.length, "the chart landed on the user's own slides").toBe(0);
  });

  it("names the appended slide when appending renumbered an existing one", async () => {
    // The other way the diff goes wrong: one slide landed, but the host renamed
    // an old one as it appended, so the deck grew by one and TWO ids are new.
    // Nothing extra landed, so nothing is deleted — the user's slides stay.
    const deck = deckOf(3);
    installHost(deck);
    faults.renumbersOnAdd = true;
    const id = await addSlideForChart();
    expect(id, "a renumbered slide made the add unnameable").toBeTruthy();
    expect(deck.length, "a slide was deleted for an add that landed once").toBe(4);
    expect(deck[deck.length - 1].id).toBe(id);
  });

  it("refuses, and deletes nothing, when the newest slides are not provably new", async () => {
    // The safety half. A host that adds at the FRONT and renumbers an old slide
    // leaves a tail that is one of the user's own: claiming it would put the
    // chart on their slide, and cleaning up after it would delete their work.
    const deck = deckOf(3);
    installHost(deck);
    faults.addsAtFront = true;
    faults.renumbersOnAdd = true;
    expect(await addSlideForChart()).toBeNull();
    expect(deck.length, "a slide was removed on an unattributable diff").toBe(4);
    const kept = deck.map((s) => s.id);
    expect(kept, "one of the user's own slides was deleted").toEqual(expect.arrayContaining(["s2", "s3"]));
  });

  it("asks for the listing after the add a second time before giving up on the slide", async () => {
    // `listOnce` was retried before the add and not after it, and the asymmetry
    // costs more after: a failed listing there is not "no id", it is a slide
    // the add has already put in the deck and nobody can now name.
    const probe = deckOf(2);
    installHost(probe);
    const start = trips.syncs;
    await addSlideForChart();
    const syncsInACall = trips.syncs - start;

    const deck = deckOf(2);
    installHost(deck);
    // The last sync of a call is the listing after the add.
    faults.failSyncOn = trips.syncs + syncsInACall;
    const id = await addSlideForChart();
    expect(id, "one failed listing after the add lost the slide it had just added").toBeTruthy();
    expect(deck.length).toBe(3);
    expect(deck[2].id).toBe(id);
  });
});

describe("r1-50 — an earlier slide leaves the deck while the own-slide chart is drawing", () => {
  afterEach(() => {
    onTrace(undefined);
    setTracing(false);
  });

  /**
   * Deck [s1, s2], the user looking at s2, an own slide appended as slide-3 at
   * index 2. `leaves` is deleted from the deck on the SECOND commit, after the
   * first batch has synced — the own slide then sits one place lower than the
   * index the draw resolved, so `getItemAt(2)` runs off the end.
   */
  async function drawWhileADeckSlideLeaves(leaves: (deck: FakeSlide[]) => void) {
    const deck = deckOf(2);
    installHost(deck, [], deck[1]);
    const heard = listen();
    const slideId = (await addSlideForChart())!;
    expect(slideId, "the fixture needs an own slide").toBeTruthy();
    let commits = 0;
    let error: unknown;
    let target: Awaited<ReturnType<typeof insertSceneIntoSlide>> | undefined;
    try {
      target = await insertSceneIntoSlide(buildChart(cfg), { slideId, tagData: JSON.stringify(cfg) }, (phase) => {
        if (phase === "commit" && ++commits === 2) leaves(deck);
      });
    } catch (e) {
      error = e;
    } finally {
      heard.stop();
    }
    return { deck, slideId, target, error, heard, left: commits >= 2 };
  }

  it("does not read the surviving slide as gone, so the chart is not drawn a second time", async () => {
    const run = await drawWhileADeckSlideLeaves((deck) => deck.splice(0, 1));
    expect(run.left, "the earlier slide never left, so this tested nothing").toBe(true);
    const own = run.deck.find((s) => s.id === run.slideId)!;
    const visible = run.deck.find((s) => s.id === "s2")!;
    expect(own, "the own slide is no longer in the deck").toBeTruthy();
    // The ink. A fallback onto the visible slide redraws the WHOLE chart there
    // while the batches that had committed stay on the own slide, so the chart
    // exists twice and the pane — which never compares the slide it asked for
    // with the one it got — reports plain success.
    expect(visible.created.length, "the chart was redrawn on the visible slide").toBe(0);
    expect(run.target, "a half-drawn chart came back as a success").toBeUndefined();
    expect(run.error, "the draw failed and nothing said so").toBeTruthy();
    const asked = run.heard.lines.find((l) => l.message.startsWith("the draw failed — asking the deck"));
    expect(asked!.data!.stillListedUnderTheSameId).toBe(true);
    expect(asked!.data!.reading).not.toMatch(/^gone/);
  });

  it("still falls back when the own slide itself is what left", async () => {
    // The case the retry was built for (rounds 370/371): the id is missing from
    // the listing as well as the index being past the end. Unchanged.
    const run = await drawWhileADeckSlideLeaves((deck) => deck.pop());
    expect(run.target, "the vanished own slide was not recovered").toBeTruthy();
    expect(run.target!.slideId).toBe("s2");
    expect(run.deck.find((s) => s.id === "s2")!.created.length).toBeGreaterThan(0);
    const asked = run.heard.lines.find((l) => l.message.startsWith("the draw failed — asking the deck"));
    expect(asked!.data!.reading).toBe("gone — the index is past the end of the deck");
  });
});

describe("r1-51 — a slide id the deck does not list", () => {
  /**
   * Real Office.js semantics for `SlideCollection.getItem(<id not in the
   * deck>)`: a proxy comes back synchronously and any statement queued through
   * it rejects the NEXT sync. The shared fake throws at the call instead, which
   * is the only reason the old fallback ever appeared to work.
   */
  function deadSlideProxy(): unknown {
    const poison = () => failSyncsOn.add(trips.syncs + 1);
    const handler: ProxyHandler<() => void> = {
      get(_t, prop) {
        if (prop === "then" || typeof prop === "symbol") return undefined;
        poison();
        return new Proxy(function () {}, handler);
      },
      set() {
        poison();
        return true;
      },
      apply() {
        poison();
        return new Proxy(function () {}, handler);
      },
    };
    return new Proxy(function () {}, handler);
  }

  async function installHostWhereDeadIdsAreRefusedAtTheSync(deck: FakeSlide[]) {
    installHost(deck);
    const ctx = (await (
      globalThis as unknown as { PowerPoint: { run: (cb: unknown) => Promise<unknown> } }
    ).PowerPoint.run(async (c: unknown) => c)) as {
      presentation: { slides: { items: FakeSlide[]; getItem: (id: string) => unknown } };
    };
    const slides = ctx.presentation.slides;
    const real = slides.getItem;
    slides.getItem = (id: string) => (slides.items.some((s) => s.id === id) ? real(id) : deadSlideProxy());
  }

  afterEach(() => {
    failSyncsOn.clear();
    onTrace(undefined);
    setTracing(false);
  });

  it("draws on the selected slide, as InsertOptions.slideId promises", async () => {
    const deck = deckOf(2);
    await installHostWhereDeadIdsAreRefusedAtTheSync(deck);
    const heard = listen();
    let target: Awaited<ReturnType<typeof insertSceneIntoSlide>>;
    try {
      target = await insertSceneIntoSlide(buildChart(cfg), {
        slideId: "a-slide-nobody-lists",
        tagData: JSON.stringify(cfg),
      });
    } finally {
      heard.stop();
    }
    expect(target, "an unlisted id threw where the contract says it falls back").toBeTruthy();
    expect(target!.slideId).toBe("s1");
    expect(deck[0].created.length, "nothing was drawn on the selected slide").toBeGreaterThan(0);
    expect(deck[1].created.length).toBe(0);
    expect(
      heard.lines.some((l) => /not in the deck — drawing on the selected slide/.test(l.message)),
      "the fallback went unrecorded",
    ).toBe(true);
  });

  it("is the same outcome the picture recovery depends on when the user deleted the slide mid-stall", async () => {
    // `updateChartResilient` re-inserts onto `wreckage.slideId` after a stalled
    // redraw. If the user deleted that slide while it stalled, the chart goes on
    // the slide they are looking at instead of the original stall error.
    const deck = deckOf(3);
    await installHostWhereDeadIdsAreRefusedAtTheSync(deck);
    const gone = deck.splice(2, 1)[0];
    const target = await insertSceneIntoSlide(buildChart(cfg), {
      slideId: gone.id,
      tagData: JSON.stringify(cfg),
      left: 80,
      top: 100,
    });
    expect(target).toBeTruthy();
    expect(
      deck.some((s) => s.id === target!.slideId),
      "the target names a slide that is not in the deck",
    ).toBe(true);
  });

  it("still tries the id as given when the listing itself could not be read", async () => {
    // "Not listed" and "could not list" are different facts. A listing that did
    // not answer says nothing about the slide, which may well be there.
    const deck = deckOf(2);
    installHost(deck);
    faults.slideIdNeverReadable = true;
    const heard = listen();
    try {
      await insertSceneIntoSlide(buildChart(cfg), { slideId: "s2", tagData: JSON.stringify(cfg) }).catch(() => {});
      // Asserted on the decision rather than on the ink: with every slide id
      // unreadable the shared fake cannot resolve ANY id either (its own
      // `getItem` reads them), so where the shapes end up says nothing here.
      // What matters is that the id was not declared absent on an answer that
      // never read one.
      expect(
        heard.lines.some((l) => /not in the deck/.test(l.message)),
        "an unreadable listing was read as an absent slide",
      ).toBe(false);
      expect(heard.lines.find((l) => l.message === "resolved the target slide to an index")?.data?.index).toBe(-1);
    } finally {
      heard.stop();
      faults.slideIdNeverReadable = false;
    }
  });
});

describe("r1-52 — a deck scan that could not look is not a deck with no charts", () => {
  afterEach(() => {
    faults.hollowReads = 0;
    faults.failSyncOn = 0;
    _resetNamedShapeForTest();
  });

  /** One slide holding one tagged chart, with the host's record pointing at it. */
  async function deckWithANamedChart() {
    const deck = deckOf(1);
    installHost(deck);
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
    const before = namedShape();
    expect(before, "the fixture needs a recorded chart").toBeTruthy();
    return { deck, before };
  }

  it("keeps the record when the slide answered short", async () => {
    const { before } = await deckWithANamedChart();
    // The shape collection comes back empty while the slide's own count says
    // one: nothing threw, and `charts` is [] for a deck that holds a chart.
    faults.hollowReads = 1;
    const scan = await listChartsInDeck();
    expect(scan.charts.length, "the fixture did not blind the scan").toBe(0);
    expect(scanIsComplete(scan), "the fixture did not make the scan incomplete").toBe(false);
    faults.hollowReads = 1;
    const kept = await refreshNamedShapeFromDeck();
    expect(kept, "a scan that could not look cleared the record").toEqual(before);
    expect(namedShape()).toEqual(before);
  });

  it("keeps the record when the page would not read at all", async () => {
    const { before } = await deckWithANamedChart();
    // The slide count is the scan's first sync; the page's is the next.
    faults.failSyncOn = trips.syncs + 2;
    const kept = await refreshNamedShapeFromDeck();
    expect(kept, "a page that failed to read cleared the record").toEqual(before);
  });

  it("says the scan could not see, rather than that no chart would be named", async () => {
    await deckWithANamedChart();
    const heard = listen();
    try {
      faults.hollowReads = 1;
      await refreshNamedShapeFromDeck();
    } finally {
      heard.stop();
    }
    const said = heard.lines.find((l) => /could not see the whole deck/.test(l.message));
    expect(said, "the incomplete scan was acted on in silence").toBeTruthy();
    expect(said!.data!.kept).toBe(true);
  });

  it("still clears the record once a COMPLETE scan finds the deck empty", async () => {
    // The other half, so the fix cannot be "never clear": rounds 242-247 are
    // what a stale id costs.
    installHost(deckOf(1));
    _rememberNamedShapeForTest("s1", "shape-gone");
    expect(await refreshNamedShapeFromDeck()).toBeNull();
    expect(namedShape()).toBeNull();
  });
});

describe("r2-21 — a chart the user groups with their own shape", () => {
  const mk = (title: string, values: number[]) => ({
    ...sampleConfig("clustered"),
    title,
    data: { categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "x", values }] },
  });

  /** Three slides, one chart each, titled A, B, C. */
  async function threeCharts() {
    const slides = deckOf(3);
    installHost(slides, [], slides[0]);
    const cfgs = [mk("A", [400, 500, 450, 520]), mk("B", [10, 20, 30, 40]), mk("C", [50, 60, 70, 80])];
    for (let i = 0; i < 3; i++)
      await insertSceneIntoSlide(buildChart(cfgs[i] as never), {
        tagData: JSON.stringify(cfgs[i]),
        slideId: slides[i].id,
      });
    return slides;
  }

  /**
   * A real slide lists its TOP-LEVEL shapes and counts them the same way. The
   * shared fake lists a group's members beside the group, and its `getCount`
   * ignores `groupHidesChildren`, so under that fault every grouped chart would
   * read "short" and the control would fail for a reason that is the fake's.
   */
  function realisticListing(slides: FakeSlide[]) {
    faults.groupHidesChildren = true;
    for (const slide of slides) {
      const s = slide as unknown as {
        created: { deleted: boolean; id: string; grouped?: { id?: string }[] }[];
        shapes: { getCount: () => { value: number } };
      };
      s.shapes.getCount = () => {
        const hidden = new Set(s.created.flatMap((x) => (x.grouped ?? []).map((c) => String(c?.id))));
        return { value: s.created.filter((x) => !x.deleted && !hidden.has(String(x.id))).length };
      };
    }
  }

  const titlesOf = (scan: Awaited<ReturnType<typeof listChartsInDeck>>) =>
    scan.charts.map((c) => (JSON.parse(c.configJson) as { title: string }).title).sort();

  afterEach(() => {
    faults.groupHidesChildren = false;
    faults.refuseGroups = 0;
    faults.failSyncOn = 0;
    faults.refuseGroupRead = false;
  });

  it("control: a deck of untouched charts is still complete, and costs no extra sync", async () => {
    const slides = await threeCharts();
    realisticListing(slides);
    const before = trips.syncs;
    const scan = await listChartsInDeck();
    const spent = trips.syncs - before;
    expect(titlesOf(scan)).toEqual(["A", "B", "C"]);
    expect(scan.nested).toEqual([]);
    expect(scanIsComplete(scan)).toBe(true);
    // slide count, the page's slide read, the page's tag read — our own chart
    // groups are tagged, so none of them is opened.
    expect(spent, "opened groups of our own").toBe(3);
  });

  it("does not report a scan complete when a chart sits inside the user's group", async () => {
    const slides = await threeCharts();
    const shapes = (slides[0] as unknown as { shapes: any }).shapes;
    const created = (slides[0] as unknown as { created: any[] }).created;
    const chartA = created.find((x) => x.type === "group" && !x.deleted && x.tagStore?.get(CHART_TAG));
    expect(chartA, "chart A did not group").toBeTruthy();
    const caption = shapes.addTextBox("Source: finance", { left: 20, top: 420, width: 300, height: 20 });
    // The user's Ctrl+G — through the fake's own addGroup, as the host does it.
    shapes.addGroup([chartA, caption]);
    realisticListing(slides);

    const scan = await listChartsInDeck();
    expect(titlesOf(scan), "the other charts are still found").toEqual(["B", "C"]);
    expect(scan.nested, "the chart inside the user's group went unnoticed").toEqual([{ slide: 1, title: "A" }]);
    expect(scanIsComplete(scan), "a scan that cannot see chart A called itself complete").toBe(false);
    expect(scanGap(scan)).toMatch(/slide 1 "A"/);
    expect(scanGap(scan)).toMatch(/ungroup it/);
  });

  it("finds the chart whose loose parts the user grouped (an ungrouped chart's tag is on a leaf)", async () => {
    const slides = deckOf(3);
    installHost(slides, [], slides[0]);
    const cfgs = [mk("A", [400, 500, 450, 520]), mk("B", [10, 20, 30, 40]), mk("C", [50, 60, 70, 80])];
    // The host refused every grouping for chart A: its tag sits on a loose anchor.
    faults.refuseGroups = 50;
    await insertSceneIntoSlide(buildChart(cfgs[0] as never), { tagData: JSON.stringify(cfgs[0]), slideId: "s1" });
    faults.refuseGroups = 0;
    for (let i = 1; i < 3; i++)
      await insertSceneIntoSlide(buildChart(cfgs[i] as never), {
        tagData: JSON.stringify(cfgs[i]),
        slideId: slides[i].id,
      });
    const created = (slides[0] as unknown as { created: { deleted: boolean }[] }).created;
    const loose = created.filter((x) => !x.deleted);
    const shapes = (slides[0] as unknown as { shapes: any }).shapes;
    const caption = shapes.addTextBox("Source: finance", { left: 20, top: 420, width: 300, height: 20 });
    shapes.addGroup([...loose, caption]);
    realisticListing(slides);

    const scan = await listChartsInDeck();
    expect(titlesOf(scan)).toEqual(["B", "C"]);
    expect(scan.nested).toEqual([{ slide: 1, title: "A" }]);
    expect(scanIsComplete(scan)).toBe(false);
  });

  it("follows a chart down through a group inside a group", async () => {
    const slides = await threeCharts();
    const shapes = (slides[1] as unknown as { shapes: any }).shapes;
    const created = (slides[1] as unknown as { created: any[] }).created;
    const chartB = created.find((x) => x.type === "group" && !x.deleted && x.tagStore?.get(CHART_TAG));
    const inner = shapes.addGroup([chartB, shapes.addTextBox("Note", { left: 0, top: 0, width: 50, height: 10 })]);
    shapes.addGroup([inner, shapes.addTextBox("Logo", { left: 0, top: 0, width: 50, height: 10 })]);
    realisticListing(slides);
    const scan = await listChartsInDeck();
    expect(scan.nested).toEqual([{ slide: 2, title: "B" }]);
    expect(scanIsComplete(scan)).toBe(false);
  });

  it("opens a group of the user's that holds no chart, and still calls the scan complete", async () => {
    // The false-refusal half. A logo cluster is the commonest group on any deck;
    // looking inside it and finding nothing must not cost Same Scale its deck.
    const slides = await threeCharts();
    const shapes = (slides[2] as unknown as { shapes: any }).shapes;
    shapes.addGroup([
      shapes.addTextBox("Company", { left: 0, top: 0, width: 50, height: 10 }),
      shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 5, height: 5 }),
    ]);
    realisticListing(slides);
    const scan = await listChartsInDeck();
    expect(titlesOf(scan)).toEqual(["A", "B", "C"]);
    expect(scan.nested).toEqual([]);
    expect(scanIsComplete(scan), "an ordinary logo group made the scan incomplete").toBe(true);
  });

  it("marks a group that will not open as unknown, not as empty", async () => {
    const slides = await threeCharts();
    const shapes = (slides[2] as unknown as { shapes: any }).shapes;
    shapes.addGroup([
      shapes.addTextBox("Company", { left: 0, top: 0, width: 50, height: 10 }),
      shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 5, height: 5 }),
    ]);
    realisticListing(slides);
    // The slide count is the scan's first sync, the page's slide read and tag
    // read are the next two, and the fourth is the one that opens the group.
    faults.failSyncOn = trips.syncs + 4;
    const scan = await listChartsInDeck();
    expect(titlesOf(scan), "what the page had already found was thrown away").toEqual(["A", "B", "C"]);
    expect(scan.tagsUnread, "a group nobody could open was counted as having nothing in it").toBeGreaterThan(0);
    expect(scanIsComplete(scan)).toBe(false);
  });
});

describe("r2-19 — a same-slide duplicate of an ungrouped chart", () => {
  type Slide = ReturnType<typeof makeSlide>;
  type Shape = Slide["created"][number];
  const liveShapes = (slide: Slide) => slide.created.filter((s) => !s.deleted);
  /** What a shape looks like to a viewer — enough to say whether an update touched it. */
  const look = (s: Shape) =>
    JSON.stringify([
      s.deleted,
      +s.left.toFixed(2),
      +s.top.toFixed(2),
      +s.width.toFixed(2),
      s.text ?? null,
      s.fillColor,
    ]);

  /** Ctrl+D as desktop PowerPoint does it: new ids, the SAME tags, offset by `dx`. */
  function duplicate(slide: Slide, shapes: Shape[], dx: number) {
    return shapes.map((s) => {
      const c = makeShape(s.type, s.geo, { left: s.left + dx, top: s.top, width: s.width, height: s.box.height });
      c.name = s.name;
      c.text = s.text;
      c.fillColor = s.fillColor;
      for (const [k, v] of s.tagStore) c.tagStore.set(k, v);
      slide.created.push(c);
      return { from: s, copy: c };
    });
  }

  /** A chart the host refused to group, so it lands loose with a parts tag — as production writes one. */
  async function looseChart() {
    const slide = makeSlide("s1");
    installHost([slide], [], slide);
    faults.refuseGroups = 1000;
    const config = { ...sampleConfig("clustered"), title: "Original" };
    await insertSceneIntoSlide(buildChart(config), { tagData: JSON.stringify(config), left: 20, top: 20 });
    faults.refuseGroups = 0;
    const original = liveShapes(slide);
    const anchor = original.find((s) => s.tagStore.has(CHART_TAG))!;
    expect(
      original.some((s) => s.type === "group"),
      "the fixture needs a chart that is NOT grouped",
    ).toBe(false);
    expect(anchor.tagStore.get(CHART_PARTS_TAG), "the fixture needs a parts tag").toBeTruthy();
    return { slide, config, original, anchor };
  }

  /** One value changed: small enough that the update writes the changed shapes in place. */
  const smallEdit = (c: ReturnType<typeof sampleConfig>) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s, i) =>
        i === 0 ? { ...s, values: s.values.map((v, j) => (j === 0 ? (v ?? 0) + 5 : v)) } : s,
      ),
    },
  });
  /** A category more: every shape's place in the scene shifts, so the update redraws. */
  const bigEdit = (c: ReturnType<typeof sampleConfig>) => ({
    ...c,
    data: {
      categories: [...c.data.categories, "Q5"],
      series: c.data.series.map((s) => ({ ...s, values: [...s.values, 80] })),
    },
  });

  async function selectAndEdit(
    pick: Shape,
    config: ReturnType<typeof sampleConfig>,
    edit: (c: ReturnType<typeof sampleConfig>) => ReturnType<typeof sampleConfig>,
  ) {
    userClicksShape(pick.id);
    const selected = await loadChartFromSelection();
    expect(selected?.target.shapeId, "the click did not select the anchor").toBe(pick.id);
    const next = edit(config);
    return updateChartInSlide(buildChart(next), selected!.target, { tagData: JSON.stringify(next) });
  }

  afterEach(() => {
    faults.refuseGroups = 0;
    onTrace(undefined);
    setTracing(false);
  });

  it("an in-place edit of the COPY leaves the original's shapes as they were", async () => {
    const { slide, config, original, anchor } = await looseChart();
    const copies = duplicate(slide, original, 400);
    const copyAnchor = copies.find((x) => x.from === anchor)!.copy;
    const before = original.map(look);
    await selectAndEdit(copyAnchor, config, smallEdit);
    expect(original.map(look), "editing the copy rewrote the original's bars and labels").toEqual(before);
  });

  it("a redraw of the COPY does not delete the original's parts", async () => {
    const { slide, config, original, anchor } = await looseChart();
    const copies = duplicate(slide, original, 400);
    const copyAnchor = copies.find((x) => x.from === anchor)!.copy;
    await selectAndEdit(copyAnchor, config, bigEdit);
    expect(
      original.filter((s) => s.deleted).length,
      "the original chart was deleted out from under the user by an edit to its copy",
    ).toBe(0);
    expect(copyAnchor.deleted, "the copy's own anchor was not replaced").toBe(true);
  });

  it("says why it left the shared shapes alone", async () => {
    const { slide, config, original, anchor } = await looseChart();
    const copies = duplicate(slide, original, 400);
    const heard = listen();
    try {
      await selectAndEdit(copies.find((x) => x.from === anchor)!.copy, config, bigEdit);
    } finally {
      heard.stop();
    }
    const said = heard.lines.find((l) => /names shapes another chart on the slide names too/.test(l.message));
    expect(said, "the parts list was dropped in silence").toBeTruthy();
    expect(said!.data!.alsoNamedBy).toBe(anchor.id);
  });

  it("still edits the ORIGINAL in place while a copy exists", async () => {
    // The owner is the older anchor. Dropping the list from BOTH would turn the
    // one safe edit into a stranding redraw.
    const { slide, config, original, anchor } = await looseChart();
    const copies = duplicate(slide, original, 400);
    const copyLooks = copies.map((x) => look(x.copy));
    const heard = listen();
    let result: Awaited<ReturnType<typeof updateChartInSlide>>;
    try {
      result = await selectAndEdit(anchor, config, smallEdit);
    } finally {
      heard.stop();
    }
    expect(result?.shapeId, "the original was redrawn instead of updated in place").toBe(anchor.id);
    expect(
      heard.lines.some((l) => l.message === "updated only the shapes that changed"),
      "the original did not take the in-place route",
    ).toBe(true);
    expect(original.some((s) => s.deleted)).toBe(false);
    expect(
      copies.map((x) => look(x.copy)),
      "editing the original touched the copy",
    ).toEqual(copyLooks);
  });

  it("a lone ungrouped chart is still updated in place, and spends no search for a rival", async () => {
    const { config, anchor } = await looseChart();
    const heard = listen();
    let result: Awaited<ReturnType<typeof updateChartInSlide>>;
    try {
      result = await selectAndEdit(anchor, config, smallEdit);
    } finally {
      heard.stop();
    }
    expect(result?.shapeId).toBe(anchor.id);
    expect(heard.lines.some((l) => /names shapes another chart/.test(l.message))).toBe(false);
    expect(heard.lines.some((l) => l.message === "updated only the shapes that changed")).toBe(true);
  });
});
