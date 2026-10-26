// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readdirSync, readFileSync } from "fs";
import { installHost, makeSlide } from "./helpers/office-host";
import type { FakeSlide } from "./helpers/office-host";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EditTarget } from "../src/render/powerpoint";
import { EN } from "../src/taskpane/i18n";

/**
 * THE THIRD BUG HUNT, SECOND FIX RUN, STREAM v2: three small pane faults the
 * last runs saw and did not fix (docs/BACKLOG.md, "The third bug hunt").
 *
 *  - V2a: Explode ended by replacing the whole pane with the chart it had re-read
 *    off the slide, so a change typed while it ran was thrown away.
 *  - V2b: the locale control showed a blank for a locale it does not list.
 *  - V2c: an i18n string no caller reaches.
 *
 * Like `fix-p1-write-races.test.ts` this runs the REAL renderer on the fake host
 * and wraps only the one entry point a test has to hold open: an Explode is
 * `updateChartInSlide` without a picture, and the window the fault lives in is
 * the seconds that call takes on the real host, with the user free to keep typing.
 */
type EditTargetLike = { slideId: string; shapeId: string; left: number; top: number; lost?: string };

const hold = vi.hoisted(() => ({
  /** Awaited before `updateChartInSlide` reaches the renderer. */
  updateGate: null as null | Promise<void>,
  /** When set, replaces the renderer's answer to `updateChartInSlide`. */
  updateAnswer: undefined as undefined | (() => unknown),
  /** Every `updateChartInSlide` call, in the order the pane made it. */
  updates: [] as { target: EditTargetLike; tagData?: string; picture: boolean }[],
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

/** A stacked-column chart config, as an inserted chart carries it. */
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

interface Deck {
  slides: FakeSlide[];
  ctx: ReturnType<typeof installHost>;
  /** The shape id each set-up chart landed as, in the order they were given. */
  ids: string[];
}

/**
 * A fake deck with a chart drawn onto each named slide by the REAL renderer. A
 * chart with `picture` goes in the way the pane's "Insert as picture" puts one:
 * a tagged picture, whose config says `render: "image"`.
 */
async function deckWith(
  charts: { slide: string; cfg: ChartConfig; picture?: boolean }[],
  slideIds: string[],
): Promise<Deck> {
  const slides = slideIds.map((id) => makeSlide(id));
  const ctx = installHost(slides);
  const renderer = await import("../src/render/powerpoint");
  const ids: string[] = [];
  for (const c of charts) {
    const cfg = c.picture ? { ...c.cfg, render: "image" as const } : c.cfg;
    const target = await renderer.insertSceneIntoSlide(buildChart(cfg), {
      tagData: JSON.stringify(cfg),
      slideId: c.slide,
      ...(c.picture ? { pictureBase64: "AAAA" } : {}),
    });
    expect(target, `the set-up chart "${c.cfg.title}" did not land`).toBeTruthy();
    ids.push((target as EditTarget).shapeId);
  }
  hold.updates.length = 0;
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
  try {
    window.localStorage.clear();
  } catch {
    /* a jsdom without storage */
  }
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
 * The user clicks a shape on a slide, the way PowerPoint reports it: through the
 * slide's OWN collection, because shape ids repeat from slide to slide.
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

const titleBox = (): string => ($("chart-title") as HTMLInputElement).value;
const pictureBox = (): HTMLInputElement => $("render-image") as HTMLInputElement;

/** "Edit selected chart" for whatever is selected, then wait for the pane to show it. */
async function loadSelected(title: string): Promise<void> {
  $("load-selection").click();
  await until(`"${title}" to load into the pane`, () => titleBox() === title);
  await sleep(5);
}

/** Tick Auto-update chart once the load's own debounced preview render has passed. */
async function tickAutoUpdate(): Promise<void> {
  await sleep(150);
  ($("auto-update") as HTMLInputElement).checked = true;
}

const exportConfig = (): ChartConfig => {
  $("json-export").click();
  return JSON.parse(($("json-io") as HTMLTextAreaElement).value);
};
const importConfig = (cfg: Partial<ChartConfig>) => {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(cfg);
  $("json-import").click();
};

beforeEach(() => {
  hold.updateGate = null;
  hold.updateAnswer = undefined;
  hold.updates.length = 0;
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

/* -------------------------------------------------------------------------- */

describe("V2a: Explode keeps what was typed while it ran", () => {
  /**
   * A picture chart, open in the pane, with Explode started and held inside the
   * renderer: the seconds a real host takes, with the person free to keep typing.
   */
  async function explodeHeld() {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]), picture: true }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    expect(pictureBox().checked, "set-up: the pane should hold a picture chart").toBe(true);
    const slow = gate();
    hold.updateGate = slow.wait;
    $("explode").click();
    await until("Explode's update to reach the renderer", () => hold.updates.length === 1);
    return {
      deck,
      /** Let the renderer finish and wait for the pane to settle. */
      finish: async () => {
        hold.updateGate = null;
        slow.open();
        await idle("Explode to finish");
      },
    };
  }

  it("keeps a title typed while Explode ran, and carries it to the exploded chart on the next Update", async () => {
    const { deck, finish } = await explodeHeld();
    typeTitle("typed during Explode");
    await finish();

    // The pane still holds what was typed, not the chart Explode re-read ("A").
    expect(titleBox(), "Explode replaced the pane with the chart it re-read off the slide").toBe(
      "typed during Explode",
    );
    expect(exportConfig().title).toBe("typed during Explode");
    // The slide shows the pre-edit chart, and the pane says it has not caught up.
    const exploded = chartsOn(deck.slides[0]);
    expect(exploded, "Explode left one chart on the slide").toHaveLength(1);
    expect(exploded[0].cfg.title).toBe("A");
    expect(noteText()).toMatch(/has not been applied yet/);
    // And what the pane pushes next is a chart again, not the picture it began as.
    expect(pictureBox().checked, "the pane still says it holds a picture").toBe(false);
    expect(exportConfig().render, "the pane would redraw the exploded chart as a picture").toBeUndefined();

    // Update: the pane's target is the exploded chart's, and the typed title lands.
    $("insert").click();
    await idle("Update to finish");
    expect(hold.updates, "Explode's own update and the one the person pressed").toHaveLength(2);
    expect(hold.updates[1].target.shapeId, "the pane kept the picture's dead id").toBe(exploded[0].id);
    expect(hold.updates[1].picture, "the typed edit went back to the slide as a picture").toBe(false);
    const after = chartsOn(deck.slides[0]);
    expect(after).toHaveLength(1);
    expect(after[0].cfg.title).toBe("typed during Explode");
    expect(after[0].cfg.render).toBeUndefined();
  });

  it("sends the edit itself when Auto-update is ticked, to the exploded chart and as shapes", async () => {
    const { deck, finish } = await explodeHeld();
    await tickAutoUpdate();
    typeTitle("typed during Explode");
    // Long enough for the timer the edit armed to fire at least once INSIDE the
    // write, find it in flight and re-arm itself (`maybeAutoUpdate`). Explode's end
    // used to cancel that timer along with the rest of the pane; it must not now.
    await sleep(1200);
    expect(hold.updates, "an auto-update wrote the chart while Explode was rewriting it").toHaveLength(1);
    await finish();
    const exploded = chartsOn(deck.slides[0]);
    await until("the held edit to go out once Explode was done", () => hold.updates.length === 2, 4000);
    await idle("the auto-update to finish");
    expect(hold.updates[1].target.shapeId).toBe(exploded[0].id);
    expect(hold.updates[1].picture).toBe(false);
    expect(JSON.parse(hold.updates[1].tagData!).title).toBe("typed during Explode");
    expect(chartsOn(deck.slides[0])[0].cfg.title).toBe("typed during Explode");
  });

  it("changes nothing and pushes nothing when nothing was typed: the note stays the ordinary one", async () => {
    const { deck, finish } = await explodeHeld();
    await tickAutoUpdate();
    await finish();
    expect(noteText()).toBe("Exploded to native shapes — edits now update them in place.");
    expect(titleBox()).toBe("A");
    expect(pictureBox().checked).toBe(false);
    expect($("insert").textContent).toBe("Update chart");
    // A chart the pane has just been handed back must not be written over with the
    // pane's own reconstruction of it (see `applyConfig`).
    await sleep(1400);
    expect(hold.updates, "a chart nothing had edited was written back").toHaveLength(1);
    expect(chartsOn(deck.slides[0])).toHaveLength(1);
  });

  it("takes the size the chart was kept at, as an ordinary update does", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]), picture: true }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    // The person had stretched the picture on the slide; the update drew the shapes
    // at that size and says so through `EditTarget.size`.
    hold.updateAnswer = () => ({
      slideId: "s1",
      shapeId: "kept-1",
      left: 10,
      top: 20,
      size: { width: 520, height: 300 },
    });
    $("explode").click();
    await idle("Explode to finish");
    expect(($("chart-w") as HTMLInputElement).value).toBe("520");
    expect(($("chart-h") as HTMLInputElement).value).toBe("300");
    expect(exportConfig()).toMatchObject({ width: 520, height: 300 });
    expect(noteText()).toBe("Exploded to native shapes — edits now update them in place.");
  });

  it("does not tell the person to press Update when Auto-update is applying the edit itself", async () => {
    // The note said "Press Update chart" and stayed on screen after the timer the
    // typed edit armed had already pushed it; following it redrew a right chart.
    const { deck, finish } = await explodeHeld();
    await tickAutoUpdate();
    typeTitle("typed during Explode");
    await finish();
    await until("the held edit to go out once Explode was done", () => hold.updates.length === 2, 4000);
    await idle("the auto-update to finish");
    expect(chartsOn(deck.slides[0])[0].cfg.title).toBe("typed during Explode");
    expect(noteText()).toMatch(/being applied by Auto-update/);
    expect(noteText(), "told the person to press Update for an edit Auto-update had sent").not.toMatch(
      /Press Update chart/,
    );
  });

  it("claims no edit when the pane only picked the exploded chart up while Explode ran", async () => {
    // The pane held B when Explode of A began; the person pressed Edit selected chart on
    // A mid-run. The pane then differs from how it STARTED (it held B), and nothing was
    // typed: the note must not say an edit is waiting.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]), picture: true },
        { slide: "s2", cfg: cfgOf("B", [4, 5, 6]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    userSelects(deck, "s2", deck.ids[1]);
    await loadSelected("B");
    userSelects(deck, "s1", deck.ids[0]);
    const slow = gate();
    hold.updateGate = slow.wait;
    $("explode").click();
    await until("Explode's update to reach the renderer", () => hold.updates.length === 1);
    await loadSelected("A");
    hold.updateGate = null;
    slow.open();
    await idle("Explode to finish");
    expect(titleBox()).toBe("A");
    expect(noteText(), "claimed an edit nobody made").toBe(
      "Exploded to native shapes — edits now update them in place.",
    );
  });

  it("keeps a size typed while Explode ran, instead of the size the slide was kept at", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]), picture: true }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    await loadSelected("A");
    hold.updateAnswer = () => ({
      slideId: "s1",
      shapeId: "kept-1",
      left: 10,
      top: 20,
      size: { width: 520, height: 300 },
    });
    const slow = gate();
    hold.updateGate = slow.wait;
    $("explode").click();
    await until("Explode's update to reach the renderer", () => hold.updates.length === 1);
    const width = $("chart-w") as HTMLInputElement;
    width.value = "700";
    width.dispatchEvent(new Event("input", { bubbles: true }));
    hold.updateGate = null;
    slow.open();
    await idle("Explode to finish");
    expect(width.value, "the typed size was replaced by the kept size").toBe("700");
    expect(exportConfig().width).toBe(700);
    expect(noteText()).toMatch(/has not been applied yet/);
  });

  it("still loads the exploded chart into a pane that held a different one", async () => {
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]), picture: true },
        { slide: "s2", cfg: cfgOf("B", [4, 5, 6]) },
      ],
      ["s1", "s2"],
    );
    await bootPane();
    userSelects(deck, "s2", deck.ids[1]);
    await loadSelected("B");
    typeTitle("B edited");
    // Explode works on the SELECTION, which is not the chart open in the pane.
    userSelects(deck, "s1", deck.ids[0]);
    $("explode").click();
    await idle("Explode to finish");
    const exploded = chartsOn(deck.slides[0]);
    expect(titleBox(), "Explode of another chart should show that chart").toBe("A");
    expect(noteText()).toBe("Exploded to native shapes — edits now update them in place.");
    $("insert").click();
    await idle("Update to finish");
    expect(hold.updates[1].target.slideId).toBe("s1");
    expect(hold.updates[1].target.shapeId).toBe(exploded[0].id);
  });

  it("still loads the exploded chart into a pane that held none", async () => {
    const deck = await deckWith([{ slide: "s1", cfg: cfgOf("A", [1, 2, 3]), picture: true }], ["s1"]);
    await bootPane();
    userSelects(deck, "s1", deck.ids[0]);
    expect($("insert").textContent).toBe("Insert into slide");
    $("explode").click();
    await idle("Explode to finish");
    expect(titleBox()).toBe("A");
    expect($("insert").textContent).toBe("Update chart");
    expect(noteText()).toBe("Exploded to native shapes — edits now update them in place.");
  });

  it("does not take a chart on another slide with the pane's shape id for the chart the pane holds", async () => {
    // Shape ids are per slide on a real host (round 497's deck holds "26" as a
    // PowerChart on five slides); the fake numbers them deck-wide, so make it so.
    const deck = await deckWith(
      [
        { slide: "s1", cfg: cfgOf("A", [1, 2, 3]) },
        { slide: "s2", cfg: cfgOf("P", [4, 5, 6]), picture: true },
      ],
      ["s1", "s2"],
    );
    const [aId, pId] = deck.ids;
    deck.slides[1].created.find((s) => !s.deleted && s.id === pId)!.id = aId;
    await bootPane();
    userSelects(deck, "s1", aId);
    await loadSelected("A");
    typeTitle("unsaved on A");

    userSelects(deck, "s2", aId);
    $("explode").click();
    await idle("Explode to finish");
    // The chart Explode worked on is P, on s2: the pane must show P, with its target.
    expect(titleBox(), "a chart on another slide was taken for the one the pane holds").toBe("P");
    $("insert").click();
    await idle("Update to finish");
    expect(hold.updates[1].target.slideId).toBe("s2");
  });
});

/* -------------------------------------------------------------------------- */

describe("V2b: the locale control shows the locale in force", () => {
  const localeSelect = (): HTMLSelectElement =>
    document.querySelector<HTMLSelectElement>('select[aria-label="Label locale"]')!;
  const optionTexts = (sel: HTMLSelectElement): string[] => [...sel.options].map((o) => o.textContent ?? "");
  const shown = (sel: HTMLSelectElement): string => sel.selectedOptions[0]?.textContent ?? "";
  const chartWith = (locale?: string): Partial<ChartConfig> => ({
    kind: "clustered",
    data: { categories: ["A", "B"], series: [{ name: "S", values: [1234.5, 2] }] },
    ...(locale === undefined ? {} : { numberFormat: { locale } }),
  });

  async function boot() {
    installHost([makeSlide("s1")]);
    await bootPane();
  }

  it("offers the four locales and selects en-US for a chart that says nothing", async () => {
    await boot();
    importConfig(chartWith());
    expect(optionTexts(localeSelect())).toEqual(["en-US", "de-DE", "fr-FR", "da-DK"]);
    expect(localeSelect().value).toBe("en-US");
    expect(exportConfig().numberFormat).toBeUndefined();
  });

  it("selects de-DE for a loaded de_DE, and exports it as it was written", async () => {
    await boot();
    importConfig(chartWith("de_DE"));
    expect(localeSelect().value, "the select is blank for a locale in force").toBe("de-DE");
    expect(shown(localeSelect())).toBe("de-DE");
    expect(localeSelect().options, "a listed locale needs no extra option").toHaveLength(4);
    expect(exportConfig().numberFormat?.locale, "an untouched control rewrote the locale").toBe("de_DE");
  });

  it.each(["en-GB", "sv-SE"])("adds a temporary option for %s, selected, and exports it unchanged", async (tag) => {
    await boot();
    importConfig(chartWith(tag));
    expect(localeSelect().value, "the select is blank for a locale in force").toBe(tag);
    expect(shown(localeSelect())).toBe(tag);
    expect(optionTexts(localeSelect())).toEqual(["en-US", "de-DE", "fr-FR", "da-DK", tag]);
    expect(exportConfig().numberFormat?.locale).toBe(tag);
  });

  it("spells the temporary option as a BCP-47 tag for a POSIX locale, and still exports what was written", async () => {
    await boot();
    importConfig(chartWith("en_GB"));
    expect(shown(localeSelect())).toBe("en-GB");
    expect(localeSelect().value).toBe("en-GB");
    expect(exportConfig().numberFormat?.locale).toBe("en_GB");
  });

  it("matches a listed locale whatever its case", async () => {
    await boot();
    importConfig(chartWith("da-dk"));
    expect(shown(localeSelect())).toBe("da-DK");
    expect(localeSelect().options).toHaveLength(4);
    expect(exportConfig().numberFormat?.locale).toBe("da-dk");
  });

  it.each(["", "not a locale", "de-DEE"])(
    "shows en-US for %j, which Intl rejects and the engine therefore prints as en-US",
    async (bad) => {
      await boot();
      importConfig(chartWith(bad));
      expect(shown(localeSelect())).toBe("en-US");
      expect(localeSelect().options).toHaveLength(4);
      // Untouched, the control rewrites nothing: a hand-written value is the person's.
      // Exactly what was written, '' included: `?? ""` here made the row pass for a
      // control that wrote en-US back into state too (an export of undefined).
      expect(exportConfig().numberFormat?.locale).toBe(bad);
    },
  );

  it("writes the new locale once the person changes the control, from a temporary option or a listed one", async () => {
    await boot();
    importConfig(chartWith("en-GB"));
    const sel = localeSelect();
    sel.value = "da-DK";
    sel.dispatchEvent(new Event("change", { bubbles: true }));
    expect(exportConfig().numberFormat?.locale).toBe("da-DK");
    // Loading another chart rebuilds the options from what the pane holds: the temporary one is gone.
    importConfig(chartWith("de_DE"));
    const again = localeSelect();
    again.value = "fr-FR";
    again.dispatchEvent(new Event("change", { bubbles: true }));
    expect(exportConfig().numberFormat?.locale).toBe("fr-FR");
    expect(localeSelect().options).toHaveLength(4);
  });
});

/* -------------------------------------------------------------------------- */

describe("V2c: an i18n string nothing reads", () => {
  const STALE = "The chart open in this pane may have been redrawn";

  it("has no catalogue entry for the Same Scale note that lost its only caller", () => {
    expect(Object.keys(EN).filter((k) => k.startsWith(STALE))).toEqual([]);
  });

  it("is read by nothing: no source file names it", () => {
    const files = readdirSync("src/taskpane").map((f) => `src/taskpane/${f}`);
    const readers = files.filter((f) => readFileSync(f, "utf8").includes(STALE));
    expect(readers).toEqual([]);
  });

  it("keeps the Explode strings the closing step now chooses between", () => {
    expect(Object.keys(EN)).toContain("Exploded to native shapes — edits now update them in place.");
    expect(Object.keys(EN)).toContain(
      "Exploded to native shapes — the edit you made while it ran has not been applied yet. Press Update chart to apply it.",
    );
  });
});
