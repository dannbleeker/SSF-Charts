// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { faults, hostSlideSize, installHost, makeSlide } from "./helpers/office-host";
import type { FakeSlide } from "./helpers/office-host";
import { MAX_ADD_RETRY_ROUNDS } from "../src/render/powerpoint";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { Scene } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * The pane's batch insert, its table element and its config round trip.
 *
 * Six findings from the second bug hunt, one file, because they share one way of
 * being checked: the REAL pane (`app.ts` against `taskpane.html`) and — for the
 * ones that reach PowerPoint — the REAL renderer on the fake host. A double that
 * answers whatever a test says cannot tell a chart that landed on its own slide
 * from one that landed on top of another, and "where did it land" is the whole
 * question for r2-25 and r2-26.
 *
 * `hold` is the only thing wrapped: it records what the pane handed the
 * renderer (the scene for the table, the options for the batch) and can make one
 * insert refuse, which is how a host that says no mid-batch is modelled.
 */
type Inserted = { tagData?: string; slideId?: string; left?: number; top?: number; pictureBase64?: string };

const hold = vi.hoisted(() => ({
  /** Every `insertSceneIntoSlide` call the pane made, with the scene it drew. */
  inserts: [] as { scene: unknown; opts: Inserted }[],
  /** The 1-based call that throws, or 0 for none. */
  refuseCall: 0,
  /** Held open before the 1-based call `gateCall` reaches the renderer — a draw the user can Stop. */
  gate: null as null | Promise<void>,
  gateCall: 0,
}));

vi.mock("../src/render/powerpoint", async (importOriginal) => {
  const real = await importOriginal<typeof import("../src/render/powerpoint")>();
  return {
    ...real,
    insertSceneIntoSlide: async (scene: unknown, opts: Inserted = {}, onPhase?: unknown) => {
      hold.inserts.push({ scene, opts });
      if (hold.gate && hold.inserts.length === hold.gateCall) await hold.gate;
      if (hold.refuseCall && hold.inserts.length === hold.refuseCall) throw new Error("host refused the insert");
      return (real.insertSceneIntoSlide as (...a: unknown[]) => unknown)(scene, opts, onPhase);
    },
  };
});

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

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
const noteIsRed = (): boolean => $("host-note").className.includes("status-err");

/** Boot the real pane against the host `installHost` put in place. */
async function bootPane(): Promise<void> {
  const O = (globalThis as unknown as { Office: Record<string, any> }).Office;
  O.onReady = (cb: () => void) => cb();
  O.PlatformType = { OfficeOnline: "OfficeOnline", PC: "PC", Mac: "Mac" };
  O.context.displayLanguage = "en-US";
  O.context.diagnostics = { platform: "PC" };
  window.history.replaceState({}, "", "/taskpane.html");
  try {
    window.localStorage.clear();
  } catch {
    /* a jsdom without storage is still a valid pane */
  }
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  const app = await import("../src/taskpane/app");
  await app.harnessReady;
  await sleep(20);
}

/** A fresh one-slide deck with the pane booted on it. */
async function deckWithPane(): Promise<FakeSlide[]> {
  const slides = [makeSlide("s1")];
  installHost(slides);
  await bootPane();
  return slides;
}

const dataOf = (values: number[]) => ({
  categories: values.map((_, i) => `C${i}`),
  series: [{ name: "S", values }],
});

/** A stacked-column config a batch can carry, sized as the caller says. */
const chartCfg = (title: string, width = 480, height = 300): ChartConfig => ({
  kind: "stacked",
  title,
  width,
  height,
  data: dataOf([3, 5, 8]),
});

/** Put `configs` in the JSON box and press Insert batch. */
async function insertBatch(configs: unknown): Promise<void> {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(configs);
  $("json-insert-batch").click();
  await idle("the batch to finish");
}

/** The config tag of every chart on a slide, in the order they were drawn. */
function chartsOn(slide: FakeSlide): ChartConfig[] {
  const out: ChartConfig[] = [];
  for (const s of slide.created) {
    const tag = s.deleted ? undefined : s.tagStore?.get("POWERCHART_CONFIG");
    if (tag) out.push(JSON.parse(tag) as ChartConfig);
  }
  return out;
}

/** The right-hand edge of the ink on a slide — the number that says "off the slide". */
function rightEdge(slide: FakeSlide): number {
  return Math.max(0, ...slide.created.filter((s) => !s.deleted).map((s) => s.box.left + s.box.width));
}

const textsOf = (scene: Scene): string[] =>
  scene.nodes.flatMap((n) => ("text" in n && typeof n.text === "string" ? [n.text] : []));

beforeEach(() => {
  hold.inserts.length = 0;
  hold.refuseCall = 0;
  hold.gate = null;
  hold.gateCall = 0;
});

afterEach(() => {
  faults.swallowAdds = 0;
  hostSlideSize.cx = 12192000;
  hostSlideSize.cy = 6858000;
  vi.unstubAllGlobals();
});

describe("r2-25 — Insert batch puts every chart on a slide of its own, fitted to it", () => {
  it("adds one slide per chart and shrinks a chart wider than the slide", async () => {
    // A 4:3 deck: 720pt wide. The third config is 900pt wide, which no slide
    // here holds. Set AFTER `installHost`, which resets the size to 16:9.
    const slides = await deckWithPane();
    hostSlideSize.cx = 9144000;
    hostSlideSize.cy = 6858000;
    await insertBatch([chartCfg("one"), chartCfg("two"), chartCfg("three", 900, 300)]);

    expect(slides.length, "the batch did not add a slide for each chart").toBe(4);
    expect(slides[0].created.length, "a chart was drawn on the slide the user was on").toBe(0);
    // One chart on each new slide, in the order they were pasted.
    expect(slides.slice(1).map((s) => chartsOn(s).map((c) => c.title))).toEqual([["one"], ["two"], ["three"]]);
    // Each chart was handed its own slide and the same origin a single insert uses.
    const sent = hold.inserts.map((i) => i.opts);
    expect(new Set(sent.map((o) => o.slideId)).size, "two charts were sent to one slide").toBe(3);
    expect(sent.map((o) => [o.left, o.top])).toEqual([
      [60, 90],
      [60, 90],
      [60, 90],
    ]);
    // The wide one was fitted to the slide, as the single path fits it (642x214),
    // and its tag carries the size it was drawn at.
    const wide = chartsOn(slides[3])[0];
    expect([wide.width, wide.height]).toEqual([642, 214]);
    for (const s of slides.slice(1)) expect(rightEdge(s), "ink ran off the slide").toBeLessThanOrEqual(720);
    expect(noteText()).toContain("3 chart");
    expect(noteIsRed()).toBe(false);
  });

  it("places a chart that cannot get a slide the way a single insert would, not on top of the last", async () => {
    // The host drops every slide add. The chart goes where the user is, which is
    // the single path's own fallback — and then the single path's own rule
    // applies: beside the last one, scaled to the room left.
    const slides = await deckWithPane();
    faults.swallowAdds = 100 * (1 + MAX_ADD_RETRY_ROUNDS);
    await insertBatch([chartCfg("one"), chartCfg("two")]);

    expect(slides.length, "a slide appeared although every add was dropped").toBe(1);
    expect(chartsOn(slides[0]).map((c) => c.title)).toEqual(["one", "two"]);
    const [first, second] = hold.inserts.map((i) => i.opts);
    expect([first.left, first.top]).toEqual([60, 90]);
    expect(
      second.left !== first.left || second.top !== first.top,
      "the second chart was drawn on exactly the same point as the first",
    ).toBe(true);
    // Said, in red: the user asked for a slide each and did not get one.
    expect(noteText()).toMatch(/current slide/i);
    expect(noteIsRed()).toBe(true);
  });
});

describe("r2-26 — one chart that cannot be inserted does not stop the rest", () => {
  it("skips an unreadable config, inserts the others and says which one it skipped", async () => {
    const slides = await deckWithPane();
    await insertBatch([
      chartCfg("first"),
      // No `data`: `normalizeData` refuses it on purpose, so a batch can isolate it.
      { kind: "clustered", title: "bad" },
      chartCfg("third"),
      chartCfg("fourth"),
    ]);

    expect(noteText(), "the batch aborted on the bad entry").not.toMatch(/^Failed/);
    expect(chartsOn(slides[0]), "a chart was drawn on the user's own slide").toEqual([]);
    expect(slides.slice(1).map((s) => chartsOn(s).map((c) => c.title))).toEqual([["first"], ["third"], ["fourth"]]);
    // Not a blank slide for the chart that never drew: nothing was added for it.
    expect(slides.length).toBe(4);
    expect(noteText()).toContain("3 of 4");
    expect(noteText(), "the note does not name the chart that was skipped").toContain("chart 2");
    expect(noteText(), "the note does not name the title").toContain("bad");
    expect(noteText(), "the note does not say why").toContain("data");
    expect(noteIsRed()).toBe(true);
  });

  it("carries on past a chart the host refused, and counts the slide it left behind", async () => {
    const slides = await deckWithPane();
    hold.refuseCall = 2;
    await insertBatch([chartCfg("a"), chartCfg("b"), chartCfg("c")]);

    expect(noteText()).not.toMatch(/^Failed/);
    const titles = slides.flatMap((s) => chartsOn(s).map((c) => c.title));
    expect(titles, "the chart after the refusal was never tried").toEqual(["a", "c"]);
    expect(noteText()).toContain("2 of 3");
    expect(noteText()).toContain("chart 2");
    expect(noteText()).toContain("host refused the insert");
    // The refused chart's slide had been added before the draw failed; the user
    // has to be told there is a slide to look at.
    expect(slides.length).toBe(4);
    expect(noteText()).toMatch(/1 slide/);
    expect(noteIsRed()).toBe(true);
  });

  it("stops between charts when Stop is pressed, and says it stopped instead of counting", async () => {
    // Stop is the one thing that must NOT be swallowed as "a chart that failed":
    // the user cancelled, the renderer's stopped error says so, and the rest of
    // the batch is theirs to keep or re-run.
    await deckWithPane();
    let open!: () => void;
    hold.gate = new Promise<void>((r) => (open = r));
    hold.gateCall = 1;
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify([chartCfg("a"), chartCfg("b"), chartCfg("c")]);
    $("json-insert-batch").click();
    await until("the first chart to reach the renderer", () => hold.inserts.length === 1);
    $("status-stop").click();
    open();
    await idle("the batch to stop");

    expect(hold.inserts, "a chart was started after Stop").toHaveLength(1);
    expect(noteText()).toMatch(/^Stopped/);
    expect(noteText(), "the stop was reported as skipped charts").not.toMatch(/Skipped/);
  });

  it("says so when there was nothing to insert, instead of reporting success", async () => {
    const slides = await deckWithPane();
    await insertBatch([]);
    expect(slides.length).toBe(1);
    expect(noteText()).toMatch(/^Nothing inserted/);
    expect(noteIsRed()).toBe(true);
  });

  it("reports every chart as skipped, and adds no slide, when none can be drawn", async () => {
    const slides = await deckWithPane();
    await insertBatch([{ kind: "clustered" }, { kind: "pie", title: "also bad" }]);
    expect(slides.length, "a blank slide was added for a chart that could not be built").toBe(1);
    expect(noteText()).toMatch(/^Nothing inserted/);
    expect(noteText()).toContain("chart 1");
    expect(noteText()).toContain("also bad");
    expect(noteIsRed()).toBe(true);
  });
});

/** Set a datasheet cell the way typing does. */
function typeCell(row: number, col: number, text: string): void {
  const input = document.querySelector<HTMLInputElement>(`#datasheet input[data-row="${row}"][data-col="${col}"]`)!;
  input.value = text;
  input.dispatchEvent(new Event("input", { bubbles: true }));
}

/** Load `cfg` through the JSON box, as a user does. */
function importConfig(cfg: unknown): void {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(cfg);
  $("json-import").click();
}

/** Read the pane's current config back out of the JSON box. */
function exportConfig(): ChartConfig {
  $("json-export").click();
  return JSON.parse(($("json-io") as HTMLTextAreaElement).value) as ChartConfig;
}

describe("r2-23 — Table from datasheet shows what the datasheet computes, not what was typed", () => {
  it("prints a formula's value, not its source", async () => {
    await deckWithPane();
    importConfig({
      kind: "clustered",
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "North", values: [10, 20] },
          { name: "South", values: [30, 40] },
          { name: "Total", values: [0, 0] },
        ],
      },
    });
    typeCell(3, 1, "=SUM(B2:B3)");
    typeCell(3, 2, "=SUM(C2:C3)");
    // The chart's own reading of the same cells, so the table is held to it.
    expect(exportConfig().data.series[2].values).toEqual([40, 60]);

    $("table-insert").click();
    await idle("the table to insert");
    const texts = textsOf(hold.inserts[0].scene as Scene);
    expect(texts, "the table printed formula source").not.toContain("=SUM(B2:B3)");
    expect(texts).toEqual(["", "Q1", "Q2", "North", "10", "20", "South", "30", "40", "Total", "40", "60"]);
  });

  it("prints a waterfall total as the total the chart draws, not as its 'e' token", async () => {
    await deckWithPane();
    importConfig(sampleConfig("waterfall"));
    // The pane seeds the token itself for every total in the config.
    const cells = [...document.querySelectorAll<HTMLInputElement>("#datasheet input")];
    expect(cells.some((i) => i.value === "e")).toBe(true);

    $("table-insert").click();
    await idle("the table to insert");
    const table = textsOf(hold.inserts[0].scene as Scene);
    const chart = textsOf(buildChart(exportConfig()));
    expect(table, "the table printed the token").not.toContain("e");
    // The last cell of the Delta row is the closing total, and the chart prints it.
    expect(table.at(-1)).toBe("93");
    expect(chart).toContain("93");
  });

  it("leaves cells it cannot evaluate, and every effect token, exactly as typed", async () => {
    await deckWithPane();
    importConfig({
      kind: "clustered",
      data: { categories: ["A", "B"], series: [{ name: "Row", values: [1, 2] }] },
    });
    typeCell(1, 1, "[up] +14%");
    typeCell(1, 2, "=SUM(B2:");
    $("table-insert").click();
    await idle("the table to insert");
    const scene = hold.inserts[0].scene as Scene;
    // The trend arrow is still an arrow: the effect token was read, not printed.
    expect(scene.nodes.some((n) => n.name === "cell-trend-1-1")).toBe(true);
    // A formula the evaluator cannot read is the user's text, and stays visible.
    expect(textsOf(scene)).toContain("=SUM(B2:");
  });
});

describe("r2-41 — a labelContent the engine draws survives opening and re-saving the chart", () => {
  const withLabels = (labelContent: unknown): ChartConfig => ({
    kind: "pie",
    width: 480,
    height: 300,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [50, 30, 20] }] },
    decorations: { labelContent } as ChartConfig["decorations"],
  });

  it("reads a bare string as a list of one, as the engine does", async () => {
    await deckWithPane();
    const original = withLabels("value");
    importConfig(original);
    const saved = exportConfig();
    expect(saved.decorations?.labelContent, "the bare string was dropped").toEqual(["value"]);
    expect(textsOf(buildChart(saved))).toEqual(textsOf(buildChart(original)));
  });

  it("keeps an empty list, which draws no label parts, rather than turning it into the default", async () => {
    await deckWithPane();
    const original = withLabels([]);
    importConfig(original);
    const saved = exportConfig();
    expect(saved.decorations?.labelContent, "[] was collapsed to the kind's default").toEqual([]);
    expect(textsOf(buildChart(saved))).toEqual(textsOf(buildChart(original)));
  });

  it("lets the user's own pick replace a carried empty list, and 'Default' mean the default", async () => {
    await deckWithPane();
    importConfig(withLabels([]));
    const select = [...document.querySelectorAll<HTMLLabelElement>("#options label")]
      .find((l) => l.textContent?.includes("Label content"))!
      .querySelector("select")!;
    select.value = "value";
    select.dispatchEvent(new Event("change"));
    expect(exportConfig().decorations?.labelContent).toEqual(["value"]);
    // Back to Default: the empty list the chart arrived with must not return.
    select.value = "";
    select.dispatchEvent(new Event("change"));
    expect(exportConfig().decorations?.labelContent).toBeUndefined();
  });

  it("still round-trips an ordinary list", async () => {
    await deckWithPane();
    importConfig(withLabels(["value", "percent"]));
    expect(exportConfig().decorations?.labelContent).toEqual(["value", "percent"]);
  });
});

describe("r2-42 — a waterfall's total column survives a '100%=' or 'X extent' row", () => {
  const waterfall = (extra: Partial<ChartConfig["data"]>): ChartConfig => ({
    kind: "waterfall",
    width: 480,
    height: 300,
    data: {
      categories: ["Start", "Up", "End"],
      series: [{ name: "Delta", values: [10, 5, 0] }],
      ...extra,
    },
    waterfall: { totalIndices: [2] },
  });

  it("puts the total token in the first series row, not in the '100%=' row above it", async () => {
    await deckWithPane();
    const original = waterfall({ hundredPercent: [100, 100, 100] });
    importConfig(original);
    const saved = exportConfig();
    expect(saved.waterfall?.totalIndices, "the closing total was lost").toEqual([2]);
    expect(saved.data.hundredPercent, "the 100%= row picked up a blank").toEqual([100, 100, 100]);
    expect(textsOf(buildChart(saved))).toEqual(textsOf(buildChart(original)));
  });

  it("does the same under an 'X extent' row, and under both", async () => {
    await deckWithPane();
    importConfig(waterfall({ xExtent: [1, 1, 1] }));
    let saved = exportConfig();
    expect(saved.waterfall?.totalIndices).toEqual([2]);
    expect(saved.data.xExtent).toEqual([1, 1, 1]);

    importConfig(waterfall({ hundredPercent: [100, 100, 100], xExtent: [1, 1, 1] }));
    saved = exportConfig();
    expect(saved.waterfall?.totalIndices).toEqual([2]);
    expect(saved.data.hundredPercent).toEqual([100, 100, 100]);
    expect(saved.data.xExtent).toEqual([1, 1, 1]);
  });

  it("still seeds row 1 for an ordinary waterfall", async () => {
    await deckWithPane();
    importConfig(waterfall({}));
    expect(exportConfig().waterfall?.totalIndices).toEqual([2]);
  });
});

describe("r2-8 — the difference arrow and value-line controls keep what they have no control for", () => {
  const base = (over: Partial<ChartConfig>): ChartConfig => ({
    kind: "stacked",
    width: 480,
    height: 300,
    data: { categories: ["A", "B"], series: [{ name: "S", values: [150, 100] }] },
    ...over,
  });

  const analysisLabel = (needle: string): HTMLLabelElement =>
    [...document.querySelectorAll<HTMLLabelElement>("#options label")].find((l) => l.textContent?.includes(needle))!;

  /** Press a difference spinner without changing what it holds. */
  const resendDifferenceSpinner = (index: number): void => {
    const spinner = analysisLabel("Difference arrow").querySelectorAll<HTMLInputElement>("input[type=number]")[index];
    spinner.dispatchEvent(new Event("input"));
  };

  const valueBox = (): HTMLInputElement =>
    analysisLabel("Value line").querySelector<HTMLInputElement>("input[type=text]")!;
  const meanBox = (): HTMLInputElement =>
    analysisLabel("Value line").querySelector<HTMLInputElement>("input[type=checkbox]")!;
  const typeValues = (text: string): void => {
    valueBox().value = text;
    valueBox().dispatchEvent(new Event("input"));
  };

  const diffLabel = (cfg: ChartConfig): string | undefined => {
    const node = buildChart(cfg).nodes.find((n) => n.name === "diff-label");
    return node && "text" in node ? node.text : undefined;
  };

  it("keeps percent:false when a spinner is pressed", async () => {
    await deckWithPane();
    const original = base({ decorations: { difference: { from: 0, to: 1, percent: false } } });
    importConfig(original);
    expect(diffLabel(original)).toBe("-50");
    resendDifferenceSpinner(0);
    const saved = exportConfig();
    expect(saved.decorations?.difference).toEqual({ from: 0, to: 1, percent: false });
    expect(diffLabel(saved), "the label flipped to a percentage").toBe("-50");
  });

  it("keeps an arrow anchored to a value line when a spinner is pressed", async () => {
    await deckWithPane();
    const original = base({
      decorations: {
        valueLines: [{ mode: "value", value: 100 }],
        difference: { from: 0, to: 0, fromValueLine: 0, percent: false },
      },
    });
    importConfig(original);
    expect(diffLabel(original)).toBe("+50");
    resendDifferenceSpinner(1);
    const saved = exportConfig();
    expect(saved.decorations?.difference).toEqual({ from: 0, to: 0, fromValueLine: 0, percent: false });
    expect(diffLabel(saved), "the arrow lost its value line and was not drawn").toBe("+50");
  });

  it("keeps the anchor and percent through unticking and re-ticking the arrow", async () => {
    await deckWithPane();
    importConfig(
      base({
        decorations: {
          valueLines: [{ mode: "value", value: 100 }],
          difference: { from: 0, to: 0, fromValueLine: 0, percent: false },
        },
      }),
    );
    const box = analysisLabel("Difference arrow").querySelector<HTMLInputElement>("input[type=checkbox]")!;
    box.checked = false;
    box.dispatchEvent(new Event("change"));
    expect(exportConfig().decorations?.difference).toBeUndefined();
    box.checked = true;
    box.dispatchEvent(new Event("change"));
    expect(exportConfig().decorations?.difference).toEqual({ from: 0, to: 0, fromValueLine: 0, percent: false });
  });

  it("keeps the lines in the order they arrived when the values box is touched", async () => {
    await deckWithPane();
    const original = base({
      // Column B stands at 180 and the mean of the two columns at 165, so an
      // arrow measured from the 100 line reads +80 and one from the mean +15.
      data: { categories: ["A", "B"], series: [{ name: "S", values: [150, 180] }] },
      decorations: {
        valueLines: [{ mode: "value", value: 100 }, { mode: "mean" }],
        difference: { from: 1, to: 1, fromValueLine: 0, percent: false },
      },
    });
    importConfig(original);
    const before = diffLabel(original);
    expect(before).toBe("+80");
    typeValues("100");
    const saved = exportConfig();
    expect(saved.decorations?.valueLines, "the mean line moved to the front").toEqual([
      { mode: "value", value: 100 },
      { mode: "mean" },
    ]);
    expect(saved.decorations?.difference?.fromValueLine).toBe(0);
    expect(diffLabel(saved), "the arrow now measures from the mean").toBe(before);
  });

  it("follows the anchored line when the values around it change", async () => {
    await deckWithPane();
    importConfig(
      base({
        decorations: {
          valueLines: [{ mode: "value", value: 100 }, { mode: "mean" }],
          // Anchored to the MEAN, which sits at index 1.
          difference: { from: 1, to: 1, fromValueLine: 1, percent: false },
        },
      }),
    );
    // A second fixed value goes in. The mean keeps its slot between the two, and
    // the arrow stays on the mean wherever that is.
    typeValues("100, 200");
    const saved = exportConfig();
    expect(saved.decorations?.valueLines).toEqual([
      { mode: "value", value: 100 },
      { mode: "mean" },
      { mode: "value", value: 200 },
    ]);
    expect(saved.decorations?.difference?.fromValueLine).toBe(1);
  });

  it("gives the anchor back when the line it was anchored to is typed out and back in", async () => {
    await deckWithPane();
    importConfig(
      base({
        decorations: {
          valueLines: [{ mode: "value", value: 100 }],
          difference: { from: 0, to: 0, fromValueLine: 0, percent: false },
        },
      }),
    );
    // Select-all, Backspace, retype: the line is gone for a moment.
    typeValues("");
    let saved = exportConfig();
    expect(saved.decorations?.valueLines).toBeUndefined();
    expect(saved.decorations?.difference?.fromValueLine, "a dead index was left pointing at nothing").toBeUndefined();
    typeValues("120");
    saved = exportConfig();
    expect(saved.decorations?.valueLines).toEqual([{ mode: "value", value: 120 }]);
    expect(saved.decorations?.difference?.fromValueLine, "the anchor was lost for good").toBe(0);
  });

  it("still puts the mean first for a chart the pane built itself", async () => {
    await deckWithPane();
    importConfig(base({}));
    meanBox().checked = true;
    meanBox().dispatchEvent(new Event("change"));
    typeValues("100");
    expect(exportConfig().decorations?.valueLines).toEqual([{ mode: "mean" }, { mode: "value", value: 100 }]);
  });
});

describe("r2-8 — layOutValueLines, directly", () => {
  type Lines = Parameters<typeof import("../src/taskpane/app").layOutValueLines>[0];
  const mean = { mode: "mean" } as const;
  const fixed = (value: number) => ({ mode: "value", value }) as const;

  async function layout(loaded: Lines, wantMean: boolean, text: string, anchor?: number) {
    await deckWithPane();
    const { layOutValueLines } = await import("../src/taskpane/app");
    return layOutValueLines(loaded, wantMean, text, anchor);
  }

  it("lays a chart the pane built out mean-first, with no anchor", async () => {
    expect(await layout([], true, "50, 100")).toEqual({ lines: [mean, fixed(50), fixed(100)], anchor: undefined });
  });

  it("keeps the order of the lines it was given, fixed values filling the value slots in turn", async () => {
    expect(await layout([fixed(1), mean, fixed(2)], true, "10, 20, 30")).toEqual({
      lines: [fixed(10), mean, fixed(20), fixed(30)],
      anchor: undefined,
    });
  });

  it("follows the anchored slot, and hands the anchor back when the slot is empty and filled again", async () => {
    const loaded: Lines = [fixed(1), mean, fixed(2)];
    expect((await layout(loaded, true, "1, 2", 2)).anchor).toBe(2);
    expect((await layout(loaded, true, "1", 2)).anchor, "the line is gone, so is the anchor").toBeUndefined();
    expect((await layout(loaded, false, "1, 2", 1)).anchor, "the mean is unticked").toBeUndefined();
    expect((await layout(loaded, false, "1, 2", 2)).anchor, "unticking the mean moves the line up one").toBe(1);
  });

  it("reads two mean lines as one mean, and anchors to it", async () => {
    expect(await layout([mean, mean], true, "", 1)).toEqual({ lines: [mean], anchor: 0 });
  });

  it("answers no anchor for an index that names nothing", async () => {
    // JSON is under no obligation to hold a number where the type says one.
    for (const bad of [-1, 5, 1.5, Number.NaN, "length" as unknown as number, "constructor" as unknown as number])
      expect((await layout([fixed(1)], false, "1", bad)).anchor, String(bad)).toBeUndefined();
  });
});
