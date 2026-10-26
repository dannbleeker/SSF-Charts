// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { installHost, makeSlide } from "./helpers/office-host";
import type { FakeSlide } from "./helpers/office-host";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EditTarget } from "../src/render/powerpoint";

/**
 * Pane writes racing each other and the selection.
 *
 * Every other pane test mocks the whole renderer (`pane-host-actions.test.ts`),
 * which is right for orchestration and wrong for these: a race between two
 * writes, or between a write and a load, is a question about WHICH chart a write
 * landed on, and a double that hands back whatever the test says cannot answer
 * it. So this file runs the REAL renderer on the fake host and wraps only the
 * three entry points a test has to be able to hold open or record.
 *
 * `hold` is the wrapper's remote control. By default every wrapped call is a
 * plain pass-through to the real function; a test arms a gate to freeze a write
 * mid-flight (the user is free to click through the deck while one runs — an
 * update can take tens of seconds on the web), or a canned answer to make the
 * renderer say something the fake cannot be made to (a chart that has gone, a
 * target the host would not name).
 */
type EditTargetLike = { slideId: string; shapeId: string; left: number; top: number; lost?: string };

const hold = vi.hoisted(() => ({
  /** Awaited before `updateChartInSlide` reaches the renderer. */
  updateGate: null as null | Promise<void>,
  /** Awaited before `updateChartsInSlides` (Same Scale's batch) reaches the renderer. */
  batchGate: null as null | Promise<void>,
  /** Awaited before `insertSceneIntoSlide` reaches the renderer. */
  insertGate: null as null | Promise<void>,
  /** When set, replaces the renderer's answer to `updateChartInSlide`. */
  updateAnswer: undefined as undefined | (() => unknown),
  /** When set, replaces the renderer's answer to `updateChartsInSlides`. */
  batchAnswer: undefined as undefined | ((items: { target: unknown }[]) => unknown[]),
  /** Every `updateChartInSlide` call, in the order the pane made it. */
  updates: [] as { target: EditTargetLike; tagData?: string; picture: boolean }[],
  /** Every `updateChartsInSlides` batch. */
  batches: [] as { target: EditTargetLike; tagData?: string }[][],
  /** Every `insertSceneIntoSlide` call made after a test's own set-up. */
  inserts: [] as { tagData?: string; slideId?: string }[],
}));

vi.mock("../src/render/powerpoint", async (importOriginal) => {
  const real = await importOriginal<typeof import("../src/render/powerpoint")>();
  return {
    ...real,
    updateChartInSlide: async (
      scene: unknown,
      target: EditTargetLike,
      opts: { tagData?: string; pictureBase64?: string } = {},
    ) => {
      hold.updates.push({ target, tagData: opts.tagData, picture: !!opts.pictureBase64 });
      if (hold.updateGate) await hold.updateGate;
      if (hold.updateAnswer) return hold.updateAnswer();
      return (real.updateChartInSlide as (...a: unknown[]) => unknown)(scene, target, opts);
    },
    updateChartsInSlides: async (
      items: { scene: unknown; target: EditTargetLike; opts?: { tagData?: string } }[],
      onFailed?: unknown,
    ) => {
      hold.batches.push(items.map((i) => ({ target: i.target, tagData: i.opts?.tagData })));
      if (hold.batchGate) await hold.batchGate;
      if (hold.batchAnswer) return hold.batchAnswer(items);
      return (real.updateChartsInSlides as (...a: unknown[]) => unknown)(items, onFailed);
    },
    insertSceneIntoSlide: async (
      scene: unknown,
      opts: { tagData?: string; slideId?: string } = {},
      onPhase?: unknown,
    ) => {
      hold.inserts.push({ tagData: opts.tagData, slideId: opts.slideId });
      if (hold.insertGate) await hold.insertGate;
      return (real.insertSceneIntoSlide as (...a: unknown[]) => unknown)(scene, opts, onPhase);
    },
  };
});

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** A gate a test opens by hand. */
function gate(): { wait: Promise<void>; open: () => void } {
  let open!: () => void;
  const wait = new Promise<void>((r) => (open = r));
  return { wait, open };
}

/** Poll until `cond` holds, or fail naming what never happened. */
async function until(what: string, cond: () => boolean, ms = 8000): Promise<void> {
  const end = Date.now() + ms;
  while (!cond()) {
    if (Date.now() > end)
      throw new Error(
        `timed out waiting for ${what} — the note read "${document.getElementById("host-note")?.textContent ?? ""}"`,
      );
    await sleep(5);
  }
}

/** The pane has nothing running: no locked button, no busy note, no Stop. */
const paneIdle = (): boolean =>
  !$<HTMLButtonElement>("insert").disabled &&
  $("host-note").className.indexOf("status-busy") < 0 &&
  $("status-stop").hasAttribute("hidden");

const idle = async (what = "the pane to go idle") => {
  await sleep(10);
  await until(what, paneIdle);
};

const noteText = (): string => $("host-note").textContent ?? "";

/** A stacked-column chart config, tagged the way an inserted chart carries it. */
const cfgOf = (title: string, values: number[], extra: Partial<ChartConfig> = {}): ChartConfig => ({
  kind: "stacked",
  title,
  width: 400,
  height: 260,
  data: { categories: values.map((_, i) => `C${i}`), series: [{ name: "S", values }] },
  ...extra,
});

/** Live shapes carrying the chart config tag on one slide, with the config parsed. */
function chartsOn(slide: FakeSlide): { id: string; cfg: ChartConfig }[] {
  const tagged: { id: string; cfg: ChartConfig }[] = [];
  for (const s of slide.created) {
    const tag = s.deleted ? undefined : s.tagStore?.get("POWERCHART_CONFIG");
    if (tag) tagged.push({ id: s.id as string, cfg: JSON.parse(tag) as ChartConfig });
  }
  return tagged;
}

/** Every live shape on a slide. */
const liveShapes = (slide: FakeSlide) => slide.created.filter((s) => !s.deleted);

interface Deck {
  slides: FakeSlide[];
  ctx: ReturnType<typeof installHost>;
  /** The shape id each set-up chart landed as, in the order they were given. */
  ids: string[];
}

/**
 * A fake deck, with a chart drawn onto each named slide by the REAL renderer —
 * exactly what an earlier session of the add-in would have left behind.
 */
async function deckWith(charts: { slide: string; cfg: ChartConfig }[], slideIds: string[]): Promise<Deck> {
  const slides = slideIds.map((id) => makeSlide(id));
  const ctx = installHost(slides);
  const renderer = await import("../src/render/powerpoint");
  const ids: string[] = [];
  for (const c of charts) {
    const target = await renderer.insertSceneIntoSlide(buildChart(c.cfg), {
      tagData: JSON.stringify(c.cfg),
      slideId: c.slide,
    });
    expect(target, `the set-up chart "${c.cfg.title}" did not land`).toBeTruthy();
    ids.push((target as EditTarget).shapeId);
  }
  hold.inserts.length = 0;
  return { slides, ctx, ids };
}

/** The pane's selection listener, once `bootPane` has captured it. */
let selectionListener: (() => unknown) | null = null;

/** Boot the real pane against the host `deckWith` installed. */
async function bootPane(): Promise<void> {
  // The Office bits app.ts needs at boot that the fake host does not carry.
  const O = (globalThis as unknown as { Office: Record<string, any> }).Office;
  O.onReady = (cb: () => void) => cb();
  O.PlatformType = { OfficeOnline: "OfficeOnline", PC: "PC", Mac: "Mac" };
  O.context.displayLanguage = "en-US";
  O.context.diagnostics = { platform: "PC" };
  selectionListener = null;
  const orig = O.context.document.addHandlerAsync;
  O.context.document.addHandlerAsync = (type: string, h: () => unknown) => {
    selectionListener = h;
    return orig(type, h);
  };
  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  const app = await import("../src/taskpane/app");
  await app.harnessReady;
  await sleep(20);
  expect(selectionListener, "the pane never registered its selection listener").toBeTypeOf("function");
}

/**
 * The user clicks a shape on a slide, the way PowerPoint reports it: the slide
 * becomes the selected one, the shape the selected shape, and the pane's
 * selection listener is told.
 *
 * Selected through the slide's OWN collection, not by looking the id up across
 * the deck — shape ids repeat from slide to slide on a real host, and a lookup
 * that returns the first match could not select the second one.
 */
function userSelects(deck: Deck, slide: string, shapeId: string | null): void {
  (deck.ctx as unknown as { presentation: { setSelectedSlides(ids: string[]): void } }).presentation.setSelectedSlides([
    slide,
  ]);
  const on = deck.slides.find((s) => s.id === slide)!;
  if (shapeId) on.setSelectedShapes([shapeId]);
  void selectionListener?.();
}

/** Type into the title box, as a user does. */
function typeTitle(text: string): void {
  const title = $("chart-title") as HTMLInputElement;
  title.value = text;
  title.dispatchEvent(new Event("input", { bubbles: true }));
}

/**
 * "Edit selected chart" for whatever is selected, then wait for the pane to show
 * it. Waits on the title box rather than on the pane going idle: the point of
 * several tests is to load a chart WHILE another action is still running.
 */
async function loadSelected(title: string): Promise<void> {
  $("load-selection").click();
  await until(`"${title}" to load into the pane`, () => ($("chart-title") as HTMLInputElement).value === title);
  await sleep(5);
}

/** Tick Auto-update chart once the load's own debounced preview render has passed. */
async function tickAutoUpdate(): Promise<void> {
  // The load re-renders the preview 80ms later, and that render arms the timer
  // if the box is already ticked. Ticking after it keeps the baseline clean: the
  // only push that follows is one the test asked for with an edit.
  await sleep(150);
  ($("auto-update") as HTMLInputElement).checked = true;
}

/** Canned `updateChartInSlide` answers: a fresh shape id on `slideId` for every call. */
function redrawnTargets(slideId: string): () => EditTargetLike {
  let n = 0;
  return () => ({ slideId, shapeId: `redrawn-${++n}`, left: 10, top: 20 });
}

/**
 * A rasteriser whose decode the test holds open — the window in which the pane
 * has built a picture update and has not yet sent it.
 */
function heldRaster(): { pending: (() => void)[]; restore: () => void } {
  const pending: (() => void)[] = [];
  class HeldImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_v: string) {
      pending.push(() => this.onload?.());
    }
  }
  vi.stubGlobal("Image", HeldImage);
  const spies = [
    vi
      .spyOn(HTMLCanvasElement.prototype, "getContext")
      .mockReturnValue({ scale() {}, drawImage() {} } as unknown as CanvasRenderingContext2D),
    vi.spyOn(HTMLCanvasElement.prototype, "toDataURL").mockReturnValue("data:image/png;base64,RASTER"),
    vi.spyOn(URL, "createObjectURL").mockImplementation(() => "blob:x"),
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {}),
  ];
  return {
    pending,
    restore: () => {
      for (const s of spies) s.mockRestore();
    },
  };
}

/** A browser that cannot decode the chart SVG: every rasterise fails at once. */
function failingRaster(): { restore: () => void } {
  class UndecodableImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_v: string) {
      Promise.resolve().then(() => this.onerror?.());
    }
  }
  vi.stubGlobal("Image", UndecodableImage);
  const spies = [
    vi.spyOn(URL, "createObjectURL").mockImplementation(() => "blob:x"),
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {}),
  ];
  return {
    restore: () => {
      for (const s of spies) s.mockRestore();
    },
  };
}

beforeEach(() => {
  hold.updateGate = null;
  hold.batchGate = null;
  hold.insertGate = null;
  hold.updateAnswer = undefined;
  hold.batchAnswer = undefined;
  hold.updates.length = 0;
  hold.batches.length = 0;
  hold.inserts.length = 0;
});

afterEach(() => {
  // A pane from this test must not go on pushing into the next one.
  const box = document.getElementById("auto-update") as HTMLInputElement | null;
  if (box) {
    box.checked = false;
    box.dispatchEvent(new Event("change"));
  }
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe("r2-3 — an Update in flight stays with the chart it started from", () => {
  it("does not re-point the pane at the redrawn chart when another was loaded meanwhile", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [100, 200, 300]) },
      ],
      ["s1", "s2"],
    );
    const [s1, s2] = deck.slides;
    const [aId, bId] = deck.ids;
    await bootPane();

    userSelects(deck, "s1", aId);
    await loadSelected("A");

    // The user edits A and presses Update; the host takes its time.
    typeTitle("A edited");
    const slow = gate();
    hold.updateGate = slow.wait;
    $("insert").click();
    await until("A's update to reach the renderer", () => hold.updates.length === 1);

    // Mid-write, they click chart B on the other slide and press Edit selected
    // chart. Nothing locks that button during an Update.
    userSelects(deck, "s2", bId);
    await loadSelected("B");

    // A's write finishes. It must not re-point the pane (now showing B) at A.
    hold.updateGate = null;
    slow.open();
    await idle("A's update to finish");
    expect(($("chart-title") as HTMLInputElement).value, "A's write overwrote the pane's chart").toBe("B");

    // The user edits B and pushes it.
    typeTitle("B edited");
    $("insert").click();
    await idle("B's update to finish");

    expect(hold.updates, "B's push never reached the renderer").toHaveLength(2);
    expect(hold.updates[1].target.slideId, "B's edit was sent to A's slide").toBe("s2");
    expect(
      chartsOn(s2).map((c) => c.cfg.title),
      "B was not updated",
    ).toEqual(["B edited"]);
    expect(
      chartsOn(s1).map((c) => c.cfg.title),
      "B's config was written over chart A",
    ).toEqual(["A edited"]);
  });

  it("keeps the loaded chart's Update button when a write that cleared its own target finishes", async () => {
    // The same blind write, by the branches that CLEAR the target: the chart the
    // write started from has gone, and the pane no longer holds it. Clearing
    // anyway turned the freshly loaded B into an "Insert into slide" — the next
    // push put a duplicate of B on the slide instead of updating it.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [100, 200, 300]) },
      ],
      ["s1", "s2"],
    );
    const [aId, bId] = deck.ids;
    await bootPane();
    userSelects(deck, "s1", aId);
    await loadSelected("A");

    typeTitle("A edited");
    const slow = gate();
    hold.updateGate = slow.wait;
    hold.updateAnswer = () => null; // A has been deleted from the slide meanwhile
    $("insert").click();
    await until("A's update to reach the renderer", () => hold.updates.length === 1);

    userSelects(deck, "s2", bId);
    await loadSelected("B");
    hold.updateGate = null;
    slow.open();
    await idle("A's update to finish");

    expect($("insert").textContent, "A's failure took B's edit target with it").toBe("Update chart");
    expect(($("chart-title") as HTMLInputElement).value).toBe("B");
  });

  it("sends the chart it rasterised to that chart's own shape, not to one loaded while it decoded", async () => {
    // The narrower half: the target was read AFTER the await on the raster, so a
    // chart loaded while the browser decoded A's picture received A's config.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3], { render: "image" }) },
        { slide: "s2", cfg: cfgOf("B", [100, 200, 300]) },
      ],
      ["s1", "s2"],
    );
    const [aId, bId] = deck.ids;
    const raster = heldRaster();
    try {
      await bootPane();
      userSelects(deck, "s1", aId);
      await loadSelected("A");
      hold.updateAnswer = redrawnTargets("s1");

      typeTitle("A edited");
      $("insert").click();
      await until("the raster to be waiting on the browser", () => raster.pending.length === 1);

      userSelects(deck, "s2", bId);
      await loadSelected("B");

      raster.pending.shift()!();
      await until("A's update to reach the renderer", () => hold.updates.length === 1);
      expect(hold.updates[0].target.slideId, "A's picture update was sent to the chart loaded meanwhile").toBe("s1");
      expect(JSON.parse(hold.updates[0].tagData!).title).toBe("A edited");
      expect(hold.updates[0].picture, "this was meant to be a picture update").toBe(true);
      await idle("A's update to finish");
    } finally {
      raster.restore();
    }
  });

  it("leaves a chart loaded while Insert as new was drawing its copy with its edit target", async () => {
    // Insert as new ends by letting go of the chart it copied, so the next push
    // cannot overwrite it. It cleared whatever the pane held when it ended, and
    // a copy takes seconds to draw.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [100, 200, 300]) },
      ],
      ["s1", "s2", "s3"],
    );
    const [aId, bId] = deck.ids;
    await bootPane();
    userSelects(deck, "s1", aId);
    await loadSelected("A");

    userSelects(deck, "s3", null);
    const slow = gate();
    hold.insertGate = slow.wait;
    $("insert-new").click();
    await until("Insert as new to reach the renderer", () => hold.inserts.length === 1);

    userSelects(deck, "s2", bId);
    await loadSelected("B");
    hold.insertGate = null;
    slow.open();
    // Not `idle`: opening B posted a settled note mid-insert, so the insert's own
    // closing "Done." is withheld and its last busy note stands — a separate,
    // older quirk of `guard`'s settle test, which is not what this asserts.
    await until(
      "Insert as new to release its buttons",
      () => !$<HTMLButtonElement>("insert").disabled && $("status-stop").hasAttribute("hidden"),
    );

    expect($("insert").textContent, "the copy's end cleared the chart opened meanwhile").toBe("Update chart");
    expect(($("chart-title") as HTMLInputElement).value).toBe("B");
  });

  it("does not let an Explode that lost its target take an unrelated chart's target with it", async () => {
    // Explode works on the SELECTION, which need not be the chart open in the
    // pane. Its `lost` branch cleared the pane's target whatever it held.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("P", [4, 5, 6]) },
      ],
      ["s1", "s2"],
    );
    const [aId, pId] = deck.ids;
    await bootPane();
    userSelects(deck, "s1", aId);
    await loadSelected("A");

    userSelects(deck, "s2", pId);
    hold.updateAnswer = () => ({ slideId: "s2", shapeId: pId, left: 10, top: 20, lost: "no-config" });
    $("explode").click();
    await idle("Explode to finish");

    expect(noteText()).toMatch(/would not save the chart's settings/);
    expect($("insert").textContent, "Explode of another chart cleared the pane's target").toBe("Update chart");
    expect(($("chart-title") as HTMLInputElement).value).toBe("A");
  });
});

describe("an update that throws after its delete", () => {
  /** A chart of its own plus a second one on the slide, so a slide swap is not offered. */
  async function crowdedSlide() {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s1", cfg: cfgOf("Z", [7, 8, 9]) },
      ],
      ["s1"],
    );
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    return deck;
  }

  it("lets go of the target it destroyed, as a Stop does", async () => {
    const raster = failingRaster();
    try {
      await crowdedSlide();
      // The renderer's own marker for "the old chart's delete has committed".
      hold.updateAnswer = () => {
        throw Object.assign(new Error("PowerPoint did not respond while drawing shapes 1-10"), {
          __powerchartWreckage: { slideId: "s1", at: { left: 40, top: 50 }, strayIds: [] },
        });
      };
      typeTitle("A edited");
      $("insert").click();
      await idle("the failed update to settle");
      expect(noteText()).toMatch(/^Failed:/);
      expect($("insert").textContent, "the pane still points an Update at the shape that update deleted").toBe(
        "Insert into slide",
      );
    } finally {
      raster.restore();
    }
  });

  it("keeps the target when the host refused before touching the chart", async () => {
    const raster = failingRaster();
    try {
      await crowdedSlide();
      hold.updateAnswer = () => {
        throw new Error("the host refused the redraw");
      };
      typeTitle("A edited");
      $("insert").click();
      await idle("the failed update to settle");
      expect(noteText()).toMatch(/^Failed:/);
      expect($("insert").textContent, "a failure that deleted nothing cost the pane its chart").toBe("Update chart");
    } finally {
      raster.restore();
    }
  });
});

describe("r2-15 — an auto-update queued behind a write does not outlive the edit target", () => {
  it("inserts nothing when the write it waited behind found the chart gone", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    await tickAutoUpdate();

    // An Update is open on the host…
    const slow = gate();
    hold.updateGate = slow.wait;
    hold.updateAnswer = () => null;
    $("insert").click();
    await until("the update to reach the renderer", () => hold.updates.length === 1);

    // …and an edit lands while it is, arming the auto-update with the target
    // still set. The write then ends with the chart gone and the target cleared.
    typeTitle("edited while the write ran");
    await sleep(100);
    hold.updateGate = null;
    slow.open();
    await idle("the update to finish");
    expect(noteText()).toMatch(/no longer on the slide/);
    expect($("insert").textContent).toBe("Insert into slide");

    // The timer fires into a pane with nothing to update. It used to take the
    // new-chart branch and put a chart on the slide that nobody asked for.
    await sleep(1400);
    expect(hold.inserts, "the auto-update inserted a brand-new chart").toHaveLength(0);
    expect(noteText(), "the note was overwritten by a timer-started insert").toMatch(/no longer on the slide/);
  });

  it("inserts nothing extra when Insert as new clears the target under a waiting edit", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1", "s2"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    await tickAutoUpdate();

    // Insert as new onto the empty slide, held open on the host.
    userSelects(deck, "s2", null);
    const slow = gate();
    hold.insertGate = slow.wait;
    $("insert-new").click();
    await until("Insert as new to reach the renderer", () => hold.inserts.length === 1);

    // An edit lands during it, then the insert ends and clears the target.
    typeTitle("edited while the copy was drawn");
    await sleep(100);
    hold.insertGate = null;
    slow.open();
    await idle("Insert as new to finish");
    expect($("insert").textContent).toBe("Insert into slide");

    await sleep(1400);
    expect(hold.inserts, "the waiting auto-update inserted a second chart").toHaveLength(1);
  });
});

describe("r2-17 — never two writes to one chart at once", () => {
  it("makes a clicked Update wait for an auto-update's write, then write to the target it returned", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    await tickAutoUpdate();
    hold.updateAnswer = redrawnTargets("s1");

    const slow = gate();
    hold.updateGate = slow.wait;
    typeTitle("first");
    await until("the auto-update to start writing", () => hold.updates.length === 1, 4000);

    // An auto-update is not a guarded action, so the Insert button is live.
    expect($<HTMLButtonElement>("insert").disabled).toBe(false);
    typeTitle("second");
    $("insert").click();
    await sleep(300);
    expect(hold.updates, "a second write to the same chart started while the first was open").toHaveLength(1);

    hold.updateGate = null;
    slow.open();
    await until("the clicked update to go out", () => hold.updates.length >= 2, 4000);
    expect(hold.updates[1].target.shapeId, "the second write used the target the first had already replaced").toBe(
      "redrawn-1",
    );
    expect(JSON.parse(hold.updates[1].tagData!).title).toBe("second");
    await idle("the clicked update to finish");
  });

  it("holds an auto-update off while Same Scale rewrites the chart being edited", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    await tickAutoUpdate();
    // The renderer's answer carries the per-chart list too (`UpdatedTargets.perItem`);
    // a plain array here made `followRescale` throw and the test pass anyway.
    hold.batchAnswer = (items) => {
      const list = items.map((i) => i.target);
      return Object.assign(list, { perItem: list });
    };
    hold.updateAnswer = redrawnTargets("s1");

    const slow = gate();
    hold.batchGate = slow.wait;
    $("same-scale").click();
    await until("Same Scale's batch to reach the renderer", () => hold.batches.length === 1);

    // An edit lands during the batch; the debounce expires with the batch open.
    typeTitle("edited during Same Scale");
    await sleep(1500);
    expect(hold.updates, "an auto-update wrote the chart while Same Scale was rewriting it").toHaveLength(0);

    hold.batchGate = null;
    slow.open();
    await until("the held edit to go out once Same Scale was done", () => hold.updates.length === 1, 4000);
    expect(JSON.parse(hold.updates[0].tagData!).title).toBe("edited during Same Scale");
    // And Same Scale itself finished: a throw in `followRescale` is swallowed by the
    // guard and shows only as a "Failed: …" note, which the assertions above cannot see.
    expect($("host-note").textContent ?? "").not.toMatch(/^Failed/);
  });

  it("holds an auto-update off while Explode rewrites the chart being edited", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    await tickAutoUpdate();
    hold.updateAnswer = redrawnTargets("s1");

    const slow = gate();
    hold.updateGate = slow.wait;
    $("explode").click();
    await until("Explode's update to reach the renderer", () => hold.updates.length === 1);

    typeTitle("edited during Explode");
    await sleep(1500);
    expect(hold.updates, "an auto-update wrote the chart while Explode was rewriting it").toHaveLength(1);

    hold.updateGate = null;
    slow.open();
    await idle("Explode to finish");
    // The pane is still holding the chart it held, so what Explode hands back is
    // only a new target: the title typed meanwhile is still the pane's, and the
    // timer that edit armed (it waited out the write, above) pushes it once
    // Explode is done. This used to find the pane replaced by the chart Explode
    // re-read, title "A", and nothing to push — the typed edit never reached the
    // slide (v2a, `fix-v2-pane-explode-locale.test.ts`).
    await until("the edit typed during Explode to go out", () => hold.updates.length === 2, 4000);
    expect(JSON.parse(hold.updates[1].tagData!).title).toBe("edited during Explode");
    expect(hold.updates[1].target.shapeId, "the edit went to the picture's dead id").toBe("redrawn-1");
  });

  it("makes Same Scale wait for an auto-update's write to one of its charts", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    await tickAutoUpdate();
    // The renderer's answer carries the per-chart list too (`UpdatedTargets.perItem`);
    // a plain array here made `followRescale` throw and the test pass anyway.
    hold.batchAnswer = (items) => {
      const list = items.map((i) => i.target);
      return Object.assign(list, { perItem: list });
    };
    hold.updateAnswer = redrawnTargets("s1");

    const slow = gate();
    hold.updateGate = slow.wait;
    typeTitle("pushed by the timer");
    await until("the auto-update to start writing", () => hold.updates.length === 1, 4000);

    $("same-scale").click();
    await sleep(400);
    expect(hold.batches, "Same Scale rewrote the chart an auto-update was still writing").toHaveLength(0);

    hold.updateGate = null;
    slow.open();
    await until("Same Scale to run once the write was done", () => hold.batches.length === 1, 4000);
    await idle("Same Scale to finish");
  });
});

describe("r2-16 — Same Scale and the chart the pane is editing", () => {
  /** Two charts, one open in the pane, then Same Scale across the deck. */
  async function rescale(a: number[], b: number[]) {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", a) },
        { slide: "s2", cfg: cfgOf("B", b) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    $("same-scale").click();
    await idle("Same Scale to finish");
    return { deck, aId: deck.ids[0] };
  }

  it("follows the chart to its new shape when the rescale had to redraw it", async () => {
    const { deck, aId } = await rescale([1, 2, 3], [10, 20, 30]);
    const [s1] = deck.slides;
    expect(noteText()).toMatch(/Same scale applied to 2 charts \(max 30\)/);
    const redrawn = chartsOn(s1);
    expect(redrawn, "set-up: this rescale is meant to redraw A under a new id").toHaveLength(1);
    expect(redrawn[0].id).not.toBe(aId);

    typeTitle("A edited");
    $("insert").click();
    await idle("the update to finish");

    expect(noteText(), "the pane lost the chart Same Scale had just redrawn").toBe("Done.");
    expect($("insert").textContent).toBe("Update chart");
    const after = chartsOn(s1);
    expect(after.map((c) => c.cfg.title)).toEqual(["A edited"]);
    expect(after[0].cfg.scale, "the update dropped the shared scale").toEqual({ max: 30 });
  });

  it("keeps the shared scale on a chart that was rescaled in place when it is then edited", async () => {
    const { deck, aId } = await rescale([10, 20, 29], [10, 20, 30]);
    const [s1] = deck.slides;
    expect(chartsOn(s1)[0].id, "set-up: this rescale is meant to land in place").toBe(aId);
    expect(chartsOn(s1)[0].cfg.scale).toEqual({ max: 30 });

    typeTitle("A edited");
    $("insert").click();
    await idle("the update to finish");

    expect(noteText()).toBe("Done.");
    const after = chartsOn(s1);
    expect(after.map((c) => c.cfg.title)).toEqual(["A edited"]);
    expect(after[0].cfg.scale, "the edit silently undid Same Scale").toEqual({ max: 30 });
  });

  it("shows the shared scale in the pane's own axis fields", async () => {
    await rescale([10, 20, 29], [10, 20, 30]);
    const inputs = [...document.querySelectorAll<HTMLInputElement>("input")].filter(
      (i) => i.type === "text" && i.placeholder === "auto",
    );
    expect(
      inputs.map((i) => i.value),
      "the pane still shows an automatic scale",
    ).toContain("30");
  });

  it("lets go of a chart Same Scale redrew but could not name again", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    // The renderer's own answer for a chart it redrew and could not tag: the OLD
    // id, marked lost. Keeping it made the next push resolve a deleted shape.
    // And the per-chart answer that goes with the list (`UpdatedTargets.perItem`):
    // every chart asked was written, so every entry is an answer.
    hold.batchAnswer = (items) => {
      const list = items.map((i) => ({ ...(i.target as EditTargetLike), lost: "no-config" }));
      return Object.assign(list, { perItem: list });
    };
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect($("insert").textContent, "the pane kept a target Same Scale had destroyed").toBe("Insert into slide");
  });

  it("leaves the pane alone when the chart it is editing was not part of the rescale", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
        { slide: "s3", cfg: cfgOf("C", [5, 5, 5], { kind: "pie" }) },
      ],
      ["s1", "s2", "s3"],
    );
    await bootPane();
    userSelects(deck, "s3", deck.ids[2]);
    await loadSelected("C");
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect($("insert").textContent).toBe("Update chart");
    expect(chartsOn(deck.slides[2])[0].id, "the pie was touched by a value-axis rescale").toBe(deck.ids[2]);
  });
});

describe("r2-18 — the selection banner knows two charts on different slides apart", () => {
  /** Two identical charts on two slides, with B's shape id forced equal to A's. */
  async function twinCharts() {
    const cfg = cfgOf("Original", [1, 2, 3]);
    const deck = await deckWith(
      [
        { slide: "s1", cfg },
        { slide: "s2", cfg },
      ],
      ["s1", "s2"],
    );
    const [, s2] = deck.slides;
    const [aId, bId] = deck.ids;
    // Shape ids are per slide on a real host (round 497's deck holds "26" as a
    // PowerChart on five slides); the fake numbers them deck-wide, so make it so.
    const b = s2.created.find((s) => !s.deleted && s.id === bId)!;
    b.id = aId;
    await bootPane();
    return { deck, aId };
  }

  it("still offers the selected chart when it has the loaded chart's shape id on another slide", async () => {
    const { deck, aId } = await twinCharts();
    userSelects(deck, "s1", aId);
    await sleep(40);
    expect($("selection-banner").style.display, "set-up: a chart is selected and nothing loaded").toBe("");
    $("selection-banner-load").click();
    await until("A to load", () => /Chart loaded/.test(noteText()) || noteText() === "Done.");
    await idle();
    expect($("selection-banner").style.display, "the banner stays up for the chart it just loaded").toBe("none");

    userSelects(deck, "s2", aId);
    await sleep(40);
    expect($("selection-banner").style.display, "the other slide's chart was mistaken for the loaded one").toBe("");
  });

  it("stays hidden for the loaded chart itself", async () => {
    const { deck, aId } = await twinCharts();
    userSelects(deck, "s1", aId);
    await sleep(40);
    $("selection-banner-load").click();
    await idle();
    userSelects(deck, "s1", aId);
    await sleep(40);
    expect($("selection-banner").style.display).toBe("none");
  });

  it("keeps the loaded chart's edit target when Explode fails on a same-id chart elsewhere", async () => {
    const { deck, aId } = await twinCharts();
    userSelects(deck, "s1", aId);
    await sleep(40);
    $("selection-banner-load").click();
    await idle();
    expect($("insert").textContent).toBe("Update chart");

    // Explode works on the SELECTION: the chart on s2, whose shape id matches the
    // pane's. Its update finds nothing, which must not clear A's target.
    userSelects(deck, "s2", aId);
    hold.updateAnswer = () => null;
    $("explode").click();
    await idle("Explode to finish");
    expect(noteText()).toMatch(/no longer on the slide/);
    expect($("insert").textContent, "an unrelated chart's failure cleared the pane's target").toBe("Update chart");
  });
});

describe("KILLED-stop — Stop during an Update", () => {
  /** A six-category clustered chart: 30-odd shapes, so a redraw takes several batches. */
  const clustered = (title: string): ChartConfig => ({
    kind: "clustered",
    title,
    width: 400,
    height: 260,
    data: {
      categories: ["C0", "C1", "C2", "C3", "C4", "C5"],
      series: [
        { name: "S", values: [1, 2, 3, 4, 5, 6] },
        { name: "T", values: [6, 5, 4, 3, 2, 1] },
      ],
    },
  });

  /**
   * Press Update on a chart the in-place path refuses (a size change), and press
   * Stop at the named moment of the redraw that follows.
   */
  async function stoppedRedraw(when: "after-first-batch" | "after-delete-before-draw") {
    const deck = await deckWith([{ slide: "s1", cfg: clustered("A") }], ["s1"]);
    const [s1] = deck.slides;
    const group = deck.ids[0];
    await bootPane();
    userSelects(deck, "s1", group);
    await loadSelected("A");
    // A size change: forces a full redraw rather than the in-place fast path.
    ($("chart-w") as HTMLInputElement).value = "520";
    $("chart-w").dispatchEvent(new Event("input", { bubbles: true }));
    await sleep(300);

    const ctx = deck.ctx as unknown as { sync: (...a: unknown[]) => Promise<unknown> };
    const origSync = ctx.sync.bind(ctx);
    const oldGroup = s1.created.find((s) => s.id === group)!;
    let pressedWith: number | null = null;
    ctx.sync = async (...args: unknown[]) => {
      const r = await origSync(...args);
      if (pressedWith === null && oldGroup.deleted) {
        const live = liveShapes(s1).length;
        if (when === "after-first-batch" ? live > 0 : live === 0) {
          pressedWith = live;
          ($("status-stop") as HTMLButtonElement).click();
        }
      }
      return r;
    };
    $("insert").click();
    await idle("the stopped update to settle");
    ctx.sync = origSync;
    expect(pressedWith, "set-up: Stop was never pressed mid-redraw").not.toBeNull();
    return { s1, oldGroup };
  }

  for (const when of ["after-first-batch", "after-delete-before-draw"] as const) {
    it(`says the chart was removed, and lets go of its dead target (${when})`, async () => {
      const { s1, oldGroup } = await stoppedRedraw(when);
      // The behaviour the manual documents and this change leaves alone: a
      // stopped redraw clears what it had drawn, and the original is gone.
      expect(oldGroup.deleted, "set-up: the original chart should have been replaced").toBe(true);
      expect(liveShapes(s1).length, "the stopped redraw left debris on the slide").toBe(0);

      expect(noteText(), "told the user the chart was kept").not.toMatch(/was kept/);
      expect(noteText()).toMatch(/removed/i);
      expect(noteText(), "did not say how to get it back").toMatch(/Ctrl\+Z/);
      expect($("host-note").className).toContain("status-err");
      expect($("insert").textContent, "the pane still points an Update at a deleted shape").toBe("Insert into slide");
    });
  }

  it("says so when a Stop took the chart an auto-update was rewriting", async () => {
    // An auto-update is not a guarded action, so its stop used to be silent — the
    // right call for a stop that cost nothing, and a chart that vanished with no
    // word said for one that took it.
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    await tickAutoUpdate();
    hold.updateAnswer = () => {
      throw Object.assign(new Error("Stopped."), {
        __powerchartStopped: true,
        __powerchartWreckage: { slideId: "s1", at: { left: 40, top: 50 }, strayIds: [] },
      });
    };
    typeTitle("A edited");
    await until("the auto-update to write", () => hold.updates.length === 1, 4000);
    await until("the pane to say the chart was removed", () => /removed/.test(noteText()));
    expect(noteText()).toMatch(/Ctrl\+Z/);
    expect($("insert").textContent, "the pane still points an Update at the removed chart").toBe("Insert into slide");
  });

  it("still says a stop that landed before anything was deleted kept the chart, and keeps its target", async () => {
    // The old wording is TRUE here, and so is the target: the stop took effect
    // before the first delete, so the chart is exactly as it was.
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    hold.updateAnswer = () => {
      throw Object.assign(new Error("Stopped."), { __powerchartStopped: true });
    };
    typeTitle("A edited");
    $("insert").click();
    await idle("the stopped update to settle");
    expect(noteText()).toBe("Stopped — anything already drawn was kept.");
    expect($("insert").textContent).toBe("Update chart");
  });
});

describe("KILLED-stop — Stop during Same Scale", () => {
  /**
   * Press Stop once `oldShapeId` has been deleted from `slide` and the redraw has
   * put its first batch back — the same moment the Update tests above press it.
   */
  function stopMidRedraw(deck: Deck, slide: string, oldShapeId: string): { pressed: () => number | null } {
    const on = deck.slides.find((s) => s.id === slide)!;
    const old = on.created.find((s) => s.id === oldShapeId)!;
    const ctx = deck.ctx as unknown as { sync: (...a: unknown[]) => Promise<unknown> };
    const origSync = ctx.sync.bind(ctx);
    let pressedWith: number | null = null;
    ctx.sync = async (...args: unknown[]) => {
      const r = await origSync(...args);
      if (pressedWith === null && old.deleted && liveShapes(on).length > 0) {
        pressedWith = liveShapes(on).length;
        ($("status-stop") as HTMLButtonElement).click();
      }
      return r;
    };
    return { pressed: () => pressedWith };
  }

  it("sweeps the half-drawn chart when the only chart it tried is the one stopped", async () => {
    // The batch THROWS when every chart it tried failed, and the sweep was written
    // after the call: Stop on the first chart left its debris on the slide.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
      ],
      ["s1", "s2"],
    );
    const [s1, s2] = deck.slides;
    const [aId, bId] = deck.ids;
    await bootPane();
    userSelects(deck, "s1", aId);
    await loadSelected("A");

    const stop = stopMidRedraw(deck, "s1", aId);
    $("same-scale").click();
    await idle("the stopped Same Scale to settle");

    expect(stop.pressed(), "set-up: Stop was never pressed mid-redraw").not.toBeNull();
    expect(liveShapes(s1).length, "the stopped redraw left a half-drawn chart on the slide").toBe(0);
    expect(noteText()).toMatch(/removed/i);
    expect(noteText()).toMatch(/Ctrl\+Z/);
    expect($("insert").textContent, "the pane still points an Update at the shape Same Scale deleted").toBe(
      "Insert into slide",
    );
    // The chart Same Scale never reached is as it was.
    expect(chartsOn(s2).map((c) => c.id)).toEqual([bId]);
    expect(chartsOn(s2)[0].cfg.scale).toBeUndefined();
  });

  it("says a chart was removed when the Stop cut one redraw short after others had landed", async () => {
    // X lands in place; Y's redraw is the one the Stop catches. The batch does
    // not throw — X landed — so this used to end on "anything already drawn was
    // kept" over a chart that had been swept.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("X", [10, 20, 30]) },
        { slide: "s2", cfg: cfgOf("Y", [1, 2, 3]) },
      ],
      ["s1", "s2"],
    );
    const [s1, s2] = deck.slides;
    const [, yId] = deck.ids;
    await bootPane();
    userSelects(deck, "s2", yId);
    await loadSelected("Y");

    const stop = stopMidRedraw(deck, "s2", yId);
    $("same-scale").click();
    await idle("the stopped Same Scale to settle");

    expect(stop.pressed(), "set-up: Stop was never pressed mid-redraw").not.toBeNull();
    expect(liveShapes(s2).length, "the stopped redraw left a half-drawn chart on the slide").toBe(0);
    expect(noteText(), "told the user nothing was lost").not.toMatch(/was kept/);
    expect(noteText()).toMatch(/removed/i);
    expect($("insert").textContent, "the pane still points an Update at the shape Same Scale deleted").toBe(
      "Insert into slide",
    );
    // The chart that had already landed keeps the shared scale, as documented.
    expect(chartsOn(s1)[0].cfg.scale).toEqual({ max: 30 });
  });
});
