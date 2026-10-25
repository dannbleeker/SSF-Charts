// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { installHost, makeSlide } from "./helpers/office-host";
import type { FakeSlide } from "./helpers/office-host";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EditTarget } from "../src/render/powerpoint";

/**
 * THE THIRD BUG HUNT, STREAM k5 — the pane's writes to the slide: what Auto-update
 * pushes, what a queued Update does when its chart is lost, and what Same Scale
 * does to the chart the pane holds and to a chart it cannot read.
 *
 * Built the way `fix-p1-write-races.test.ts` is, and for the same reason: each of
 * these is a question about WHICH chart a write landed on, and a double that hands
 * back whatever the test says cannot answer it. The REAL renderer runs on the fake
 * host; only the three entry points a test has to hold open or answer are wrapped.
 */
type EditTargetLike = { slideId: string; shapeId: string; left: number; top: number; lost?: string };

const hold = vi.hoisted(() => ({
  /** Awaited before `updateChartInSlide` reaches the renderer. */
  updateGate: null as null | Promise<void>,
  /** When set, replaces the renderer's answer to `updateChartInSlide`. */
  updateAnswer: undefined as undefined | (() => unknown),
  /** When set, replaces the renderer's answer to `updateChartsInSlides` (given the real one). */
  batchAnswer: undefined as
    undefined | ((items: { target: unknown }[], real: (items: unknown[]) => Promise<unknown[]>) => unknown),
  /** Every `updateChartInSlide` call, in the order the pane made it. */
  updates: [] as { target: EditTargetLike; tagData?: string }[],
  /** Every `updateChartsInSlides` batch. */
  batches: [] as { target: EditTargetLike; tagData?: string }[][],
  /** Every `insertSceneIntoSlide` call made after a test's own set-up. */
  inserts: [] as { tagData?: string; slideId?: string }[],
}));

vi.mock("../src/render/powerpoint", async (importOriginal) => {
  const real = await importOriginal<typeof import("../src/render/powerpoint")>();
  return {
    ...real,
    updateChartInSlide: async (scene: unknown, target: EditTargetLike, opts: { tagData?: string } = {}) => {
      hold.updates.push({ target, tagData: opts.tagData });
      if (hold.updateGate) await hold.updateGate;
      if (hold.updateAnswer) return hold.updateAnswer();
      return (real.updateChartInSlide as (...a: unknown[]) => unknown)(scene, target, opts);
    },
    updateChartsInSlides: async (
      items: { scene: unknown; target: EditTargetLike; opts?: { tagData?: string } }[],
      onFailed?: unknown,
    ) => {
      hold.batches.push(items.map((i) => ({ target: i.target, tagData: i.opts?.tagData })));
      const realBatch = (its: unknown[]) =>
        (real.updateChartsInSlides as (...a: unknown[]) => Promise<unknown[]>)(its, onFailed);
      if (hold.batchAnswer) return hold.batchAnswer(items, realBatch);
      return realBatch(items);
    },
    insertSceneIntoSlide: async (
      scene: unknown,
      opts: { tagData?: string; slideId?: string } = {},
      onPhase?: unknown,
    ) => {
      hold.inserts.push({ tagData: opts.tagData, slideId: opts.slideId });
      return (real.insertSceneIntoSlide as (...a: unknown[]) => unknown)(scene, opts, onPhase);
    },
  };
});

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

function gate(): { wait: Promise<void>; open: () => void } {
  let open!: () => void;
  const wait = new Promise<void>((r) => (open = r));
  return { wait, open };
}

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

const paneIdle = (): boolean =>
  !$<HTMLButtonElement>("insert").disabled &&
  $("host-note").className.indexOf("status-busy") < 0 &&
  $("status-stop").hasAttribute("hidden");
const idle = async (what = "the pane to go idle") => {
  await sleep(10);
  await until(what, paneIdle);
};
const noteText = (): string => $("host-note").textContent ?? "";

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
const liveShapes = (slide: FakeSlide) => slide.created.filter((s) => !s.deleted);

interface Deck {
  slides: FakeSlide[];
  ctx: ReturnType<typeof installHost>;
  /** The shape id each set-up chart landed as, in the order they were given. */
  ids: string[];
}

/**
 * A fake deck, with a chart drawn onto each named slide by the REAL renderer.
 * `tag` replaces the config the shape is tagged with — how a deck carrying a tag
 * that is not a chart (an agenda slide's, a hand-edited one) is made.
 */
async function deckWith(
  charts: { slide: string; cfg: ChartConfig; tag?: string }[],
  slideIds: string[],
): Promise<Deck> {
  const slides = slideIds.map((id) => makeSlide(id));
  const ctx = installHost(slides);
  const renderer = await import("../src/render/powerpoint");
  const ids: string[] = [];
  for (const c of charts) {
    const target = await renderer.insertSceneIntoSlide(buildChart(c.cfg), {
      tagData: c.tag ?? JSON.stringify(c.cfg),
      slideId: c.slide,
    });
    expect(target, `the set-up chart "${c.cfg.title}" did not land`).toBeTruthy();
    ids.push((target as EditTarget).shapeId);
  }
  hold.inserts.length = 0;
  return { slides, ctx, ids };
}

let selectionListener: (() => unknown) | null = null;

async function bootPane(): Promise<void> {
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

function userSelects(deck: Deck, slide: string, shapeIds: string[]): void {
  (deck.ctx as unknown as { presentation: { setSelectedSlides(ids: string[]): void } }).presentation.setSelectedSlides([
    slide,
  ]);
  const on = deck.slides.find((s) => s.id === slide)!;
  on.setSelectedShapes(shapeIds);
  void selectionListener?.();
}

function typeTitle(text: string): void {
  const title = $("chart-title") as HTMLInputElement;
  title.value = text;
  title.dispatchEvent(new Event("input", { bubbles: true }));
}

async function loadSelected(title: string): Promise<void> {
  $("load-selection").click();
  await until(`"${title}" to load into the pane`, () => ($("chart-title") as HTMLInputElement).value === title);
  await sleep(5);
}

const tickAutoUpdate = (on = true) => {
  ($("auto-update") as HTMLInputElement).checked = on;
};

function redrawnTargets(slideId: string): () => EditTargetLike {
  let n = 0;
  return () => ({ slideId, shapeId: `redrawn-${++n}`, left: 10, top: 20 });
}

beforeEach(() => {
  hold.updateGate = null;
  hold.updateAnswer = undefined;
  hold.batchAnswer = undefined;
  hold.updates.length = 0;
  hold.batches.length = 0;
  hold.inserts.length = 0;
  try {
    window.localStorage.clear();
  } catch {
    /* a jsdom without storage */
  }
});

afterEach(() => {
  const box = document.getElementById("auto-update") as HTMLInputElement | null;
  if (box) {
    box.checked = false;
    box.dispatchEvent(new Event("change"));
  }
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

/* ─────────────────────────────── r3-30 ───────────────────────────────────────── */

describe("opening a chart is not an edit: Auto-update does not write it back (r3-30)", () => {
  it("leaves the slide alone after 'Edit selected chart', whatever the box said before the load", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("Opened not edited", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    // Ticked BEFORE the load, which is the order that armed it: the load's own
    // debounced render was the push.
    tickAutoUpdate();
    userSelects(deck, "s1", [deck.ids[0]]);
    await loadSelected("Opened not edited");
    await sleep(1400);
    expect(hold.updates, "opening the chart wrote it back to the slide").toHaveLength(0);
    expect($("insert").textContent, "set-up: the chart is open for editing").toBe("Update chart");
  });

  it("does not recolour a chart whose stored config has no style when a browser style file exists", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("Stored plain", [1, 2, 3]) }], ["s1"]);
    window.localStorage.setItem(
      "ssf-charts-style",
      JSON.stringify({ palette: ["#112233", "#445566"], fontFamily: "Georgia" }),
    );
    await bootPane();
    tickAutoUpdate();
    userSelects(deck, "s1", [deck.ids[0]]);
    await loadSelected("Stored plain");
    await sleep(1400);
    expect(hold.updates, "the style file was pushed onto a chart nobody edited").toHaveLength(0);
    expect(chartsOn(deck.slides[0])[0].cfg.style, "the stored config gained a style").toBeUndefined();
  });

  it("still pushes the first real edit after a load", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    tickAutoUpdate();
    userSelects(deck, "s1", [deck.ids[0]]);
    await loadSelected("A");
    await sleep(400);
    expect(hold.updates).toHaveLength(0);
    typeTitle("A edited");
    await until("the edit to reach the slide", () => hold.updates.length === 1, 4000);
    expect(JSON.parse(hold.updates[0].tagData!).title).toBe("A edited");
    await idle("the update to finish");
  });

  it("drops an edit still waiting for the chart that was open when another is loaded", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    tickAutoUpdate();
    userSelects(deck, "s1", [deck.ids[0]]);
    await loadSelected("A");
    await sleep(100);
    typeTitle("A edited, not yet pushed");
    userSelects(deck, "s2", [deck.ids[1]]);
    await loadSelected("B");
    await sleep(1400);
    expect(hold.updates, "a pending push of A's edit went out against B").toHaveLength(0);
  });
});

/* ─────────────────────────────── r3-35 ───────────────────────────────────────── */

describe("a clicked Update that waited behind a write which lost the chart is not an Insert (r3-35)", () => {
  async function updateQueuedBehindAutoUpdate() {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]) }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", [deck.ids[0]]);
    await loadSelected("A");
    await sleep(150);
    tickAutoUpdate();
    const slow = gate();
    hold.updateGate = slow.wait;
    typeTitle("first edit");
    await until("the auto-update to start writing", () => hold.updates.length === 1, 4000);
    expect($<HTMLButtonElement>("insert").disabled, "set-up: the Update button is live during an auto-update").toBe(
      false,
    );
    typeTitle("second edit");
    $("insert").click();
    await sleep(200);
    return { deck, slow };
  }

  it("inserts nothing when the write ahead of it ends by losing the chart's settings", async () => {
    const { deck, slow } = await updateQueuedBehindAutoUpdate();
    hold.updateAnswer = () => ({ slideId: "s1", shapeId: "redrawn", left: 10, top: 20, lost: "no-config" });
    hold.updateGate = null;
    slow.open();
    await idle("both writes to finish");
    expect(hold.updates, "the clicked Update was written as an update after the target was lost").toHaveLength(1);
    expect(hold.inserts, "the clicked Update became a brand-new Insert").toHaveLength(0);
    expect(liveShapes(deck.slides[0]).length, "a second copy of the chart is on the slide").toBeLessThan(20);
    expect(noteText(), "the note the first write left was overwritten").toMatch(/would not save the chart's settings/);
    expect($("insert").textContent).toBe("Insert into slide");
  });

  it("inserts nothing when the write ahead of it finds the chart gone", async () => {
    const { slow } = await updateQueuedBehindAutoUpdate();
    hold.updateAnswer = () => null;
    hold.updateGate = null;
    slow.open();
    await idle("both writes to finish");
    expect(hold.inserts).toHaveLength(0);
    expect(noteText()).toMatch(/no longer on the slide/);
  });

  it("still writes the clicked Update to the target the write ahead returned when that write went well", async () => {
    const { slow } = await updateQueuedBehindAutoUpdate();
    hold.updateAnswer = redrawnTargets("s1");
    hold.updateGate = null;
    slow.open();
    await until("the clicked update to go out", () => hold.updates.length >= 2, 4000);
    expect(hold.updates[1].target.shapeId).toBe("redrawn-1");
    expect(JSON.parse(hold.updates[1].tagData!).title).toBe("second edit");
    await idle("the clicked update to finish");
    expect(hold.inserts).toHaveLength(0);
  });

  it("is still an Insert when there was no chart to update when the button was pressed", async () => {
    await deckWith([], ["s1"]);
    await bootPane();
    expect($("insert").textContent).toBe("Insert into slide");
    $("insert").click();
    await until("the insert to reach the renderer", () => hold.inserts.length === 1, 4000);
    await idle("the insert to finish");
  });
});

/* ─────────────────────────────── r3-36 ───────────────────────────────────────── */

describe("Same Scale that reaches only some charts leaves the pane on a live target (r3-36)", () => {
  /**
   * The real batch, with the charts NOT on `keep` handed over naming a shape the
   * slide does not answer for — the host declining the rest, inside the renderer.
   * Leaving them out of the items instead hides them from it, and its per-chart
   * answer (`UpdatedTargets.perItem`) is then about fewer charts than the pane asked.
   */
  const reachOnly =
    (keep: string[]) => (items: { target: unknown }[], real: (items: unknown[]) => Promise<unknown[]>) =>
      real(
        items.map((i) =>
          keep.includes((i.target as EditTargetLike).slideId)
            ? i
            : { ...i, target: { ...(i.target as object), shapeId: "no-such-shape" } },
        ),
      );

  async function threeCharts() {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
        { slide: "s3", cfg: cfgOf("C", [5, 10, 15]) },
      ],
      ["s1", "s2", "s3"],
    );
    await bootPane();
    userSelects(deck, "s1", [deck.ids[0]]);
    await loadSelected("A");
    return deck;
  }

  it("follows the chart to its new shape when ANOTHER chart was the one the host dropped", async () => {
    const deck = await threeCharts();
    const [s1] = deck.slides;
    hold.batchAnswer = reachOnly(["s1", "s2"]);
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect(noteText()).toMatch(/Same scale reached 2 of 3 charts/);
    const redrawn = chartsOn(s1);
    expect(redrawn, "set-up: this rescale is meant to redraw A under a new id").toHaveLength(1);
    expect(redrawn[0].id).not.toBe(deck.ids[0]);

    typeTitle("A edited");
    $("insert").click();
    await idle("the update to finish");
    expect(noteText(), "the pane kept the dead id of the chart Same Scale redrew").toBe("Done.");
    expect($("insert").textContent).toBe("Update chart");
    expect(chartsOn(s1).map((c) => c.cfg.title)).toEqual(["A edited"]);
    expect(hold.inserts, "a second copy of A was put on the slide").toHaveLength(0);
  });

  it("follows it when the dropped chart is the one on a slide the user deleted mid-run", async () => {
    const deck = await threeCharts();
    hold.batchAnswer = async (items, real) => {
      deck.slides[2].created.forEach((s) => ((s as { deleted?: boolean }).deleted = true));
      return real(items);
    };
    $("same-scale").click();
    await idle("Same Scale to finish");
    typeTitle("A edited");
    $("insert").click();
    await idle("the update to finish");
    expect(noteText()).toBe("Done.");
    expect(chartsOn(deck.slides[0]).map((c) => c.cfg.title)).toEqual(["A edited"]);
  });

  it("leaves the pane alone when its own chart was the one that was not reached", async () => {
    const deck = await threeCharts();
    hold.batchAnswer = reachOnly(["s2", "s3"]);
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect(noteText()).toMatch(/Same scale reached 2 of 3 charts/);
    expect(noteText(), "the pane was blamed for a chart that was never touched").not.toMatch(/open in this pane/);
    expect(chartsOn(deck.slides[0])[0].id, "set-up: A was not redrawn").toBe(deck.ids[0]);
    typeTitle("A edited");
    $("insert").click();
    await idle("the update to finish");
    expect(noteText()).toBe("Done.");
    expect(chartsOn(deck.slides[0]).map((c) => c.cfg.title)).toEqual(["A edited"]);
  });

  it("tells which of two charts on the pane's slide came back, so the pane is not left guessing", async () => {
    // This was "says so when it cannot tell": the renderer's list could not say
    // which of the two it had dropped. Its per-chart answer can (fix-m3 has the
    // whole matrix), so the pane's chart is left alone without a word about it.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s1", cfg: cfgOf("A2", [4, 5, 6]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    userSelects(deck, "s1", [deck.ids[0]]);
    await loadSelected("A");
    // Of the two on s1 the host answers for one (A2), and for B.
    hold.batchAnswer = (items, real) =>
      real([{ ...items[0], target: { ...(items[0].target as object), shapeId: "no-such-shape" } }, items[1], items[2]]);
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect(noteText()).toMatch(/Same scale reached 2 of 3 charts/);
    expect(noteText(), "the pane was blamed for a chart that was never touched").not.toMatch(/open in this pane/);
    typeTitle("A edited");
    $("insert").click();
    await idle("the update to finish");
    expect(noteText()).toBe("Done.");
    expect(chartsOn(deck.slides[0]).find((c) => c.cfg.title === "A edited")?.id).toBe(deck.ids[0]);
  });
});

/* ─────────────────────────────── r3-12 ───────────────────────────────────────── */

describe("Same scale (deck) leaves out a chart whose config cannot be built and does the rest (r3-12)", () => {
  const AGENDA = JSON.stringify({ kind: "agenda", chapters: ["Context", "Findings", "Recommendation"], highlight: 1 });

  it("skips the skill's agenda slide, rescales the two charts, and names what it skipped", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
        { slide: "s3", cfg: cfgOf("agenda stand-in", [1]), tag: AGENDA },
      ],
      ["s1", "s2", "s3"],
    );
    await bootPane();
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect(hold.batches, "Same Scale stopped before it rescaled anything").toHaveLength(1);
    expect(hold.batches[0], "the agenda tag was handed to the renderer as a chart").toHaveLength(2);
    expect(noteText()).toMatch(/Same scale applied to 2 charts \(max 30\)/);
    expect(noteText(), "the skipped chart was not named").toMatch(/Skipped 1 it could not read — kind "agenda"\./);
    expect($("host-note").className, "work that did not happen was reported in green").toContain("status-err");
    expect(chartsOn(deck.slides[0])[0].cfg.scale, "A was not rescaled").toEqual({ max: 30 });
    expect(chartsOn(deck.slides[1])[0].cfg.scale).toEqual({ max: 30 });
  });

  it("skips a tag that is not JSON and one that is null, naming them as unreadable", async () => {
    await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
        { slide: "s3", cfg: cfgOf("x", [1]), tag: "{not json" },
        { slide: "s4", cfg: cfgOf("y", [1]), tag: "null" },
      ],
      ["s1", "s2", "s3", "s4"],
    );
    await bootPane();
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect(hold.batches).toHaveLength(1);
    expect(hold.batches[0]).toHaveLength(2);
    expect(noteText()).toMatch(/Same scale applied to 2 charts/);
    expect(noteText()).toMatch(/Skipped 2 it could not read — an unreadable chart ×2\./);
  });

  it("says what it skipped when too few charts are left to share a scale", async () => {
    await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("agenda stand-in", [1]), tag: AGENDA },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect(hold.batches).toHaveLength(0);
    expect(noteText()).toMatch(/Same scale needs at least two value-axis charts in the deck\./);
    expect(noteText()).toMatch(/Skipped 1 it could not read — kind "agenda"\./);
  });

  it("still passes over a chart with no value axis without a word", async () => {
    await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
        { slide: "s3", cfg: cfgOf("C", [5, 5, 5], { kind: "pie" }) },
      ],
      ["s1", "s2", "s3"],
    );
    await bootPane();
    $("same-scale").click();
    await idle("Same Scale to finish");
    expect(noteText()).toBe("Same scale applied to 2 charts (max 30).");
  });
});
