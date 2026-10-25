// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { failSyncsOn, installHost, makeSlide } from "./helpers/office-host";
import type { FakeSlide } from "./helpers/office-host";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import { insertSceneIntoSlide, requestStop, resetStop, updateChartsInSlides } from "../src/render/powerpoint";
import type { EditTarget, UpdatedTargets } from "../src/render/powerpoint";

/**
 * THE THIRD BUG HUNT, FIX STREAM m3 — Same Scale that reaches only some charts
 * must say which.
 *
 * `updateChartsInSlides` used to hand back one target per chart it RESOLVED and
 * nothing for the rest, so a caller holding a list of N could not tell which of
 * the N it had been given back. The pane's `followRescale` matched them by
 * position inside the open chart's slide, which only works when nothing on that
 * slide was dropped: two asked charts and one answer left it unable to say
 * whether its own chart had been redrawn (a dead id) or never touched (a live
 * one), and it left the pane as it was and said so.
 *
 * Two halves, and each fails alone:
 *
 *   - the RENDERER's answer is per chart asked, in the order asked, `null` for a
 *     chart it never touched (the `perItem` describes below, against the fake
 *     host with no pane);
 *   - the PANE follows its own chart by that answer (the describes after them,
 *     the real pane and the real renderer on the fake host).
 *
 * The host DECLINING a chart is made the way the host does it, inside the real
 * renderer: the chart's target is handed over naming a shape the slide does not
 * answer for, which is exactly what its `alive` filter drops. Filtering the
 * items before the real call (what r3-36's tests did) hides the chart from the
 * renderer, which can then answer for only the charts it was shown.
 */
type EditTargetLike = { slideId: string; shapeId: string; left: number; top: number; lost?: string };

const hold = vi.hoisted(() => ({
  /** When set, replaces the renderer's answer to `updateChartsInSlides` (given the real one). */
  batchAnswer: undefined as
    undefined | ((items: { target: unknown }[], real: (items: unknown[]) => Promise<unknown>) => unknown),
  /** Every `insertSceneIntoSlide` call the pane made after a test's own set-up. */
  inserts: [] as { tagData?: string; slideId?: string }[],
}));

vi.mock("../src/render/powerpoint", async (importOriginal) => {
  const real = await importOriginal<typeof import("../src/render/powerpoint")>();
  return {
    ...real,
    updateChartsInSlides: async (
      items: { scene: unknown; target: EditTargetLike; opts?: { tagData?: string } }[],
      onFailed?: unknown,
    ) => {
      const realBatch = (its: unknown[]) =>
        (real.updateChartsInSlides as (...a: unknown[]) => Promise<unknown>)(its, onFailed);
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

/** What the pane's two axis-scale boxes show: "" is automatic. */
const scaleBoxes = (): string[] =>
  [...document.querySelectorAll<HTMLInputElement>("input")]
    .filter((i) => i.type === "text" && i.placeholder === "auto")
    .map((i) => i.value);

const cfgOf = (title: string, values: number[]): ChartConfig => ({
  kind: "stacked",
  title,
  width: 400,
  height: 260,
  data: { categories: values.map((_, i) => `C${i}`), series: [{ name: "S", values }] },
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

interface Deck {
  slides: FakeSlide[];
  ctx: ReturnType<typeof installHost>;
  /** The shape id each set-up chart landed as, in the order they were given. */
  ids: string[];
}

/** A fake deck, with a chart drawn onto each named slide by the REAL renderer. */
async function deckWith(charts: { slide: string; cfg: ChartConfig }[], slideIds: string[]): Promise<Deck> {
  const slides = slideIds.map((id) => makeSlide(id));
  const ctx = installHost(slides);
  const ids: string[] = [];
  for (const c of charts) {
    const target = await insertSceneIntoSlide(buildChart(c.cfg), { tagData: JSON.stringify(c.cfg), slideId: c.slide });
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
  deck.slides.find((s) => s.id === slide)!.setSelectedShapes(shapeIds);
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

beforeEach(() => {
  hold.batchAnswer = undefined;
  hold.inserts.length = 0;
  resetStop();
  try {
    window.localStorage.clear();
  } catch {
    /* a jsdom without storage */
  }
});

afterEach(() => {
  resetStop();
  failSyncsOn.clear();
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

/**
 * The real batch, with the charts at `which` handed over naming a shape the slide
 * does not answer for — the renderer's own `alive` filter drops them, the way it
 * drops a chart the host would not resolve. The chart itself stays on the slide.
 */
const hostDeclines =
  (...which: number[]) =>
  (items: { target: unknown }[], real: (items: unknown[]) => Promise<unknown>) =>
    real(
      items.map((it, i) =>
        which.includes(i) ? { ...it, target: { ...(it.target as object), shapeId: "no-such-shape" } } : it,
      ),
    );

/* ──────────────────────────────── the renderer's answer ───────────────────────────────── */

describe("updateChartsInSlides answers for every chart it was asked about, in the order asked", () => {
  const stacked = (v: number): ChartConfig => cfgOf(`T${v}`, [v, v + 1]);

  /** n tagless charts on one slide: each target names a real shape, so each is redrawn. */
  function loose(n: number) {
    const slide = makeSlide("s1");
    installHost([slide]);
    const items = Array.from({ length: n }, (_, i) => {
      const s = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
      return {
        scene: buildChart(stacked(i)),
        target: { slideId: "s1", shapeId: s.id as string, left: 10, top: 20 } as EditTarget,
        opts: { tagData: `{"i":${i}}` },
      };
    });
    return { slide, items };
  }

  it("marks a chart whose shape the host would not resolve, and keeps the list it always returned", async () => {
    const { items } = loose(3);
    const ghost = { ...items[1], target: { ...items[1].target, shapeId: "no-such-shape" } };
    const next = await updateChartsInSlides([items[0], ghost, items[2]]);
    // The list is what it was: one target per chart the host resolved.
    expect(next, "the dropped chart is not in the list").toHaveLength(2);
    // And the answer names which of the three each of those is.
    expect(next.perItem, "no per-chart answer at all").toHaveLength(3);
    expect(next.perItem[1], "a chart the host would not resolve was reported as reached").toBeNull();
    expect(next.perItem[0]?.shapeId, "the first chart was not handed its new shape").toBe(next[0].shapeId);
    expect(next.perItem[2]?.shapeId, "the third chart was handed the first's shape").toBe(next[1].shapeId);
    expect(next[0].shapeId).not.toBe(items[0].target.shapeId);
    expect(next[1].shapeId).not.toBe(items[2].target.shapeId);
  });

  it("marks a chart whose slide the host would not resolve", async () => {
    const { items } = loose(3);
    const lost = { ...items[0], target: { ...items[0].target, slideId: "no-such-slide" } };
    const next = await updateChartsInSlides([lost, items[1], items[2]]);
    expect(next).toHaveLength(2);
    expect(next.perItem.map((t) => (t ? "reached" : null))).toEqual([null, "reached", "reached"]);
  });

  it("answers null for every chart when none could be resolved", async () => {
    const { items } = loose(2);
    const none = items.map((it) => ({ ...it, target: { ...it.target, slideId: "no-such-slide" } }));
    const next = await updateChartsInSlides(none);
    expect(next).toHaveLength(0);
    expect(next.perItem).toEqual([null, null]);
  });

  it("answers an empty batch with an empty answer", async () => {
    installHost([makeSlide("s1")]);
    const next = await updateChartsInSlides([]);
    expect(next).toEqual([]);
    expect(next.perItem).toEqual([]);
  });

  it("gives a chart a Stop took the redraw of no target, while the list still carries the one it kept", async () => {
    const { items } = loose(3);
    const batches = Math.ceil(buildChart(stacked(0)).nodes.length / 10);
    // 2 resolve syncs, then per chart 1 delete + `batches` renders: this is the
    // first render batch of chart 2, so chart 2 is deleted and never drawn.
    failSyncsOn.add(2 + (1 + batches) + 1 + 1);
    // The Stop the person presses when they see the stall — before chart 3.
    const next = await updateChartsInSlides(items, () => requestStop());
    // Chart 1 redrew; chart 2's old shapes are gone and its target is spent, not
    // "not reached"; chart 3 was never touched.
    expect(next.perItem[0]?.shapeId).not.toBe(items[0].target.shapeId);
    expect(next.perItem[0]?.lost).toBeUndefined();
    expect(next.perItem[1]?.lost, "a chart whose delete committed and whose redraw failed is spent").toBe(
      "unknown-shape",
    );
    expect(next.perItem[2], "a chart the Stop never reached was reported as touched").toBeNull();
    // The list is unchanged in meaning: the untouched chart still comes back with its old target.
    expect(next).toHaveLength(3);
    expect(next[2].shapeId).toBe(items[2].target.shapeId);
  });

  it("counts a chart that was written in place as reached even though its target did not change", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg = cfgOf("In place", [10, 20, 29]);
    const target = (await insertSceneIntoSlide(buildChart(cfg), {
      tagData: JSON.stringify(cfg),
      slideId: "s1",
    })) as EditTarget;
    const next = cfgOf("In place", [10, 20, 29]);
    next.scale = { max: 30 };
    const out = await updateChartsInSlides([
      { scene: buildChart(next), target, opts: { tagData: JSON.stringify(next) } },
    ]);
    // Written where it stands: the shape survived and the tag now carries the scale.
    expect(chartsOn(slide).map((c) => c.id)).toEqual([target.shapeId]);
    expect(chartsOn(slide)[0].cfg.scale, "set-up: the scale was not written").toEqual({ max: 30 });
    expect(out.perItem, "an in-place edit looked like a chart nobody touched").toHaveLength(1);
    expect(out.perItem[0]?.shapeId).toBe(target.shapeId);
  });

  it("does not show the answer to anything that compares or spreads the list", async () => {
    const { items } = loose(2);
    const next = await updateChartsInSlides(items);
    expect(Object.keys(next)).toEqual(["0", "1"]);
    expect(next).toEqual([...next]);
    expect(JSON.stringify(next)).toBe(JSON.stringify([...next]));
  });
});

/* ──────────────────────────────────── the pane ───────────────────────────────────── */

describe("the pane follows its own chart through a Same Scale that reached only some", () => {
  /**
   * s1 holds A and A2 (two asked charts on the open chart's slide), s2 holds B.
   * The default values are chosen so the shared scale (max 30) redraws A under a
   * new id and writes A2 where it stands; each test says when it needs otherwise.
   */
  async function twoOnASlide(a = [1, 2, 3], a2 = [4, 5, 6]) {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", a) },
        { slide: "s1", cfg: cfgOf("A2", a2) },
        { slide: "s2", cfg: cfgOf("B", [10, 20, 30]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    return deck;
  }

  async function open(deck: Deck, title: string, id: string) {
    userSelects(deck, "s1", [id]);
    await loadSelected(title);
  }

  const sameScale = async () => {
    $("same-scale").click();
    await idle("Same Scale to finish");
  };

  /** Edit the title and press Update; the chart on s1 carrying `title` after. */
  async function editAndUpdate(title: string) {
    typeTitle(title);
    $("insert").click();
    await idle("the update to finish");
  }

  it("adopts the redrawn id when it is the FIRST of two charts on its slide and the SECOND was dropped", async () => {
    const deck = await twoOnASlide();
    await open(deck, "A", deck.ids[0]);
    hold.batchAnswer = hostDeclines(1);
    await sameScale();
    const said = noteText();
    const onS1 = chartsOn(deck.slides[0]);
    expect(onS1.find((c) => c.cfg.title === "A")?.id, "set-up: A was meant to be redrawn under a new id").not.toBe(
      deck.ids[0],
    );
    expect(onS1.find((c) => c.cfg.title === "A2")?.id, "set-up: A2 was meant to be left alone").toBe(deck.ids[1]);

    await editAndUpdate("A edited");
    expect(noteText(), "the pane kept the dead id of the chart Same Scale redrew").toBe("Done.");
    expect(
      chartsOn(deck.slides[0])
        .map((c) => c.cfg.title)
        .sort(),
    ).toEqual(["A edited", "A2"]);
    expect(hold.inserts, "a second copy was put on the slide").toHaveLength(0);
    expect(scaleBoxes(), "the pane did not take over the scale its chart now carries").toContain("30");
    expect(said).toMatch(/Same scale reached 2 of 3 charts/);
    expect(said, "the pane was told it could not tell, with the answer in hand").not.toMatch(/open in this pane/);
  });

  it("keeps its live target and says nothing alarming when it is the SECOND and that one was dropped", async () => {
    const deck = await twoOnASlide();
    await open(deck, "A2", deck.ids[1]);
    hold.batchAnswer = hostDeclines(1);
    await sameScale();
    expect(noteText()).toMatch(/Same scale reached 2 of 3 charts/);
    expect(noteText(), "the pane was blamed for a chart that was never touched").not.toMatch(/open in this pane/);
    expect(scaleBoxes(), "the pane took a scale its chart never got").not.toContain("30");
    expect(chartsOn(deck.slides[0]).find((c) => c.cfg.title === "A2")?.id, "set-up: A2 was not redrawn").toBe(
      deck.ids[1],
    );

    await editAndUpdate("A2 edited");
    expect(noteText()).toBe("Done.");
    const after = chartsOn(deck.slides[0]);
    expect(after.find((c) => c.cfg.title === "A2 edited")?.id, "the edit did not land on A2's own shape").toBe(
      deck.ids[1],
    );
    expect(hold.inserts).toHaveLength(0);
  });

  it("keeps its live target when it is the FIRST and that one was dropped", async () => {
    const deck = await twoOnASlide();
    await open(deck, "A", deck.ids[0]);
    hold.batchAnswer = hostDeclines(0);
    await sameScale();
    expect(noteText()).toMatch(/Same scale reached 2 of 3 charts/);
    expect(noteText()).not.toMatch(/open in this pane/);
    expect(scaleBoxes()).not.toContain("30");
    await editAndUpdate("A edited");
    expect(noteText()).toBe("Done.");
    expect(chartsOn(deck.slides[0]).find((c) => c.cfg.title === "A edited")?.id).toBe(deck.ids[0]);
  });

  it("adopts the redrawn id when it is the SECOND and the FIRST was dropped", async () => {
    const deck = await twoOnASlide([1, 2, 3], [1, 2, 3]);
    await open(deck, "A2", deck.ids[1]);
    hold.batchAnswer = hostDeclines(0);
    await sameScale();
    const said = noteText();
    expect(
      chartsOn(deck.slides[0]).find((c) => c.cfg.title === "A2")?.id,
      "set-up: A2 was meant to be redrawn",
    ).not.toBe(deck.ids[1]);
    await editAndUpdate("A2 edited");
    expect(noteText(), "the pane kept the dead id of the chart Same Scale redrew").toBe("Done.");
    expect(hold.inserts, "a second copy was put on the slide").toHaveLength(0);
    expect(scaleBoxes()).toContain("30");
    expect(said).toMatch(/Same scale reached 2 of 3 charts/);
    expect(said).not.toMatch(/open in this pane/);
  });

  it("takes over the scale of a chart that was written in place, which has no new id to adopt", async () => {
    // A [10,20,29] under a shared max of 30 is written where it stands: the answer for
    // it is the target it already had, and it is still an answer.
    const deck = await twoOnASlide([10, 20, 29]);
    await open(deck, "A", deck.ids[0]);
    hold.batchAnswer = hostDeclines(1);
    await sameScale();
    expect(chartsOn(deck.slides[0]).find((c) => c.cfg.title === "A")?.id, "set-up: A was meant to land in place").toBe(
      deck.ids[0],
    );
    expect(scaleBoxes(), "the pane's copy of the chart still shows an automatic scale").toContain("30");
    await editAndUpdate("A edited");
    expect(noteText()).toBe("Done.");
    expect(
      chartsOn(deck.slides[0]).find((c) => c.cfg.title === "A edited")?.cfg.scale,
      "the edit silently undid Same Scale on the chart",
    ).toEqual({ max: 30 });
  });

  it("leaves the pane on a chart a Stop kept Same Scale from reaching, and gives it no scale it never got", async () => {
    const deck = await twoOnASlide();
    userSelects(deck, "s2", [deck.ids[2]]);
    await loadSelected("B");
    // Stop lands before the first chart is touched: nothing is rescaled.
    hold.batchAnswer = async (items, real) => {
      (await import("../src/render/powerpoint")).requestStop();
      return real(items);
    };
    await sameScale();
    expect(chartsOn(deck.slides[1])[0].id, "set-up: B was not left alone").toBe(deck.ids[2]);
    expect(chartsOn(deck.slides[1])[0].cfg.scale, "set-up: B was rescaled").toBeUndefined();
    expect(scaleBoxes(), "the pane claimed a scale the chart on the slide does not have").not.toContain("30");
    expect($("insert").textContent, "the pane let go of a chart nothing had touched").toBe("Update chart");
    expect(noteText(), "a Stop that rescaled nothing was reported as a rescale").not.toMatch(/Same scale applied/);
    await editAndUpdate("B edited");
    expect(noteText()).toBe("Done.");
    expect(chartsOn(deck.slides[1])[0].id).toBe(deck.ids[2]);
  });

  it("counts the charts it reached the way the per-chart answer does", async () => {
    await twoOnASlide();
    let answered: UpdatedTargets | undefined;
    hold.batchAnswer = async (items, real) => {
      answered = (await hostDeclines(0, 1)(items, real)) as UpdatedTargets;
      return answered;
    };
    await sameScale();
    expect(answered?.perItem.filter(Boolean), "the answer reached a different number of charts").toHaveLength(1);
    expect(noteText()).toMatch(/Same scale reached 1 of 3 charts — PowerPoint would not answer for the other 2\./);
    expect($("host-note").className).toContain("status-err");
  });
});
