// @vitest-environment jsdom
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import { spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import JSZip from "jszip";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { sceneFingerprint } from "../src/core/scene-diff";
import { setTracing, traceLog } from "../src/core/trace";
import * as lib from "../src/index";
import { listChartsInDeck, updateChartInSlide } from "../src/render/powerpoint";
import type { Scene } from "../src/core/scene";
import type { ChartConfig, ChartKind } from "../src/core/types";
import { fromEuropeanNumber, looksEuropean, sheetToData, typedNumber } from "../src/taskpane/datasheet";
import { installHost, makeSlide, type FakeShape, type FakeSlide } from "./helpers/office-host";

/**
 * THE THIRD BUG HUNT, STREAM v3: two leftovers the m2 review named.
 *
 * V3a: a thousands group cannot start with a zero. M2e (9d59b76) fixed
 * `EU_GROUPED` and `US_GROUPED`; every OTHER place the datasheet reads a thousands
 * group kept the head `\d{1,3}` or `\d{1,2}`, and BACKLOG.md called them
 * unreachable. They are reachable by any text a cell or a box can hold.
 *
 * V3b: a chart the headless skill writes carried no scene fingerprint, so its
 * first edit in the add-in was always a full redraw.
 */

afterEach(() => {
  vi.unstubAllGlobals();
  setTracing(false);
});

/** See `pane-state.test.ts`: the pane's last debounced render must not outlive jsdom. */
afterAll(async () => {
  await new Promise((r) => setTimeout(r, 250));
});

/** What the datasheet makes of one typed or pasted cell: the number, or null for a gap. */
const cell = (text: string): number | null =>
  sheetToData({
    cells: [
      ["", "A"],
      ["S", text],
    ],
  }).series[0].values[0];

/** A pasted block with one column of data under a header, the way `looksEuropean` is handed one. */
const block = (...cells: string[]) => [["Region", "N"], ...cells.map((c) => ["x", c])];

describe("V3a: the space and apostrophe group in a datasheet cell cannot open with a zero", () => {
  // Every spelling of the group gap `GROUP_GAP` names: plain, no-break, narrow
  // no-break (fr-FR), thin space, the ASCII apostrophe and U+2019 (de-CH, fr-CH).
  const GAPS = [" ", " ", " ", " ", "'", "’"];

  it("reads 0 500 as no number, not five hundred, whichever mark groups it", () => {
    for (const gap of GAPS) {
      expect(cell(`0${gap}500`), JSON.stringify(`0${gap}500`)).toBeNull();
      expect(cell(`00${gap}500`), "two zeros").toBeNull();
      expect(cell(`007${gap}500`), "a zero-led three-digit head").toBeNull();
      expect(cell(`0${gap}500${gap}000`), "two groups").toBeNull();
      expect(cell(`0${gap}000`), "a zero is not written as a group of zeros").toBeNull();
    }
  });

  it("refuses it with its dressing on too: a sign, parentheses, a currency, a percent", () => {
    expect(cell("-0 500")).toBeNull();
    expect(cell("(0 500)")).toBeNull();
    expect(cell("€ 0 500")).toBeNull();
    expect(cell("0 500 kr.")).toBeNull();
    expect(cell("0 500%")).toBeNull();
    expect(cell("0 500-")).toBeNull();
  });

  it("refuses a zero-led space group that carries a decimal, in either mark", () => {
    expect(cell("0 500.5")).toBeNull();
    expect(cell("0 500,5")).toBeNull();
  });

  it("still reads every group that opens with 1 to 9", () => {
    for (const gap of GAPS) {
      expect(cell(`1${gap}500`), `1${gap}500`).toBe(1500);
      expect(cell(`10${gap}500`)).toBe(10500);
      expect(cell(`999${gap}999`)).toBe(999999);
      expect(cell(`1${gap}234${gap}567`)).toBe(1234567);
      expect(cell(`-1${gap}500`)).toBe(-1500);
      expect(cell(`(1${gap}500)`)).toBe(-1500);
      expect(cell(`€ 1${gap}500`)).toBe(1500);
      expect(cell(`1${gap}500.5`)).toBe(1500.5);
    }
  });

  it("leaves a plain zero, a zero decimal and an unspaced leading zero as they were", () => {
    expect(cell("0")).toBe(0);
    expect(cell("0.5")).toBe(0.5);
    expect(cell("-0.5")).toBe(-0.5);
    expect(cell("0.500")).toBe(0.5);
    expect(cell(".5")).toBe(0.5);
    // No separator, so no group: a leading zero here is only a number written with a
    // zero in front, which `Number` has always read.
    expect(cell("0500")).toBe(500);
    expect(cell("007")).toBe(7);
  });

  it("casts no European vote and is not converted: the rest of a Danish block still is", () => {
    // `0 500,5` read as the core `0500,5`, which `EU_DECIMAL` took for a comma
    // decimal: it voted European and `fromEuropeanNumber` handed back 0500.5.
    expect(fromEuropeanNumber("0 500,5")).toBe("0 500,5");
    expect(fromEuropeanNumber("0'500,5")).toBe("0'500,5");
    expect(looksEuropean(block("0 500,5"))).toBe(false);
    // And a real European paste around it is unaffected.
    expect(looksEuropean(block("0 500,5", "1.234,5"))).toBe(true);
    expect(fromEuropeanNumber("1 234,5")).toBe("1234.5");
  });
});

describe("V3a: Indian grouping cannot open with a zero", () => {
  it("reads 0,12,345 as no number, not twelve thousand", () => {
    expect(cell("0,12,345")).toBeNull();
    expect(cell("01,23,456")).toBeNull();
    expect(cell("0,00,500")).toBeNull();
    expect(cell("00,12,345.5")).toBeNull();
    expect(cell("-0,12,345")).toBeNull();
    expect(cell("(0,12,345)")).toBeNull();
  });

  it("is no evidence for an American sheet, so a Danish block keeps its conversion", () => {
    // It matched `LAKH_GROUPED`, which `looksEuropean` counts as a comma that is a
    // thousands mark, and one American cell takes the conversion from the whole block.
    expect(looksEuropean(block("0,12,345", "1.234,5"))).toBe(true);
    expect(looksEuropean(block("01,23,456", "987,5"))).toBe(true);
  });

  it("still reads every Indian grouping that opens with 1 to 9", () => {
    expect(cell("1,23,456")).toBe(123456);
    expect(cell("12,34,567.89")).toBe(1234567.89);
    expect(cell("1,00,000")).toBe(100000);
    expect(cell("99,99,999")).toBe(9999999);
    expect(cell("-1,23,456")).toBe(-123456);
    expect(cell("+12,34,567")).toBe(1234567);
    expect(cell("(1,23,456)")).toBe(-123456);
    expect(looksEuropean(block("1,23,456", "987.5"))).toBe(false);
    expect(looksEuropean(block("12,34,567.89", "1.234,5"))).toBe(false);
  });

  it("leaves the US and the European groups as M2e left them", () => {
    expect(cell("1,500")).toBe(1500);
    expect(cell("12,345.6")).toBe(12345.6);
    expect(cell("0,500")).toBeNull();
    expect(fromEuropeanNumber("1.500")).toBe("1500");
    expect(fromEuropeanNumber("10.500")).toBe("10500");
    expect(fromEuropeanNumber("1.234.567")).toBe("1234567");
    expect(fromEuropeanNumber("0.500")).toBe("0.500");
  });
});

describe("V3a: typedNumber reads a zero-led group as a decimal or as nothing, never as a group", () => {
  const LOCALES = ["en-US", "da-DK", "de-DE", "fr-FR"];

  it("is NaN for a repeated mark that opens with a zero", () => {
    for (const locale of LOCALES) {
      for (const text of ["0.500.500", "0,500,500", "00.500.500", "01,234,567", "-0.500.500", "+0,500,500"]) {
        expect(typedNumber(text, locale), `${text} in ${locale}`).toBeNaN();
      }
    }
  });

  it("is NaN for both marks with a zero-led group: 0.500,5 and 0,500.5", () => {
    for (const locale of LOCALES) {
      for (const text of ["0.500,5", "0,500.5", "00.500,25", "0.500.500,5", "-0,500.5"]) {
        expect(typedNumber(text, locale), `${text} in ${locale}`).toBeNaN();
      }
    }
  });

  it("is NaN for a zero-led space or apostrophe group, the box's copy of the cell rule", () => {
    for (const locale of LOCALES) {
      for (const gap of [" ", " ", " ", " ", "'", "’"]) {
        expect(typedNumber(`0${gap}500`, locale), `0${gap}500 in ${locale}`).toBeNaN();
        expect(typedNumber(`00${gap}500`, locale)).toBeNaN();
        expect(typedNumber(`0${gap}500,5`, locale)).toBeNaN();
        expect(typedNumber(`0${gap}500.5`, locale)).toBeNaN();
        expect(typedNumber(`-0${gap}500`, locale)).toBeNaN();
      }
    }
  });

  it("still reads every group that opens with 1 to 9, in every locale", () => {
    for (const locale of LOCALES) {
      expect(typedNumber("1.234.567", locale), locale).toBe(1234567);
      expect(typedNumber("1,234,567", locale), locale).toBe(1234567);
      expect(typedNumber("10.500.250", locale)).toBe(10500250);
      expect(typedNumber("999,999,999", locale)).toBe(999999999);
      expect(typedNumber("1.234,5", locale)).toBe(1234.5);
      expect(typedNumber("1,234.5", locale)).toBe(1234.5);
      expect(typedNumber("1.234.567,89", locale)).toBe(1234567.89);
      expect(typedNumber("1,234,567.89", locale)).toBe(1234567.89);
      expect(typedNumber("-1.234.567", locale)).toBe(-1234567);
      expect(typedNumber("1 500", locale)).toBe(1500);
      expect(typedNumber("1'500", locale)).toBe(1500);
      expect(typedNumber("1’500", locale)).toBe(1500);
      expect(typedNumber("1 234 567", locale)).toBe(1234567);
      expect(typedNumber("-1 500", locale)).toBe(-1500);
    }
    expect(typedNumber("1 234,5", "da-DK")).toBe(1234.5);
    expect(typedNumber("1 234.5", "en-US")).toBe(1234.5);
  });

  it("keeps the single-mark decimals and the leading-zero decimals as they were", () => {
    for (const locale of LOCALES) {
      expect(typedNumber("0.500", locale), `0.500 in ${locale}`).toBe(0.5);
      expect(typedNumber("-0.5", locale)).toBe(-0.5);
      expect(typedNumber("0,5", locale)).toBe(0.5);
      expect(typedNumber("-0,5", locale)).toBe(-0.5);
      expect(typedNumber("0", locale)).toBe(0);
      expect(typedNumber("0.0", locale)).toBe(0);
      expect(typedNumber("10.500", "da-DK")).toBe(10500);
      expect(typedNumber("10.500", "en-US")).toBe(10.5);
    }
    // The ones the locale decides, as k4 pinned them.
    expect(typedNumber("1,500", "da-DK")).toBe(1.5);
    expect(typedNumber("1,500", "en-US")).toBe(1500);
    expect(typedNumber("3.000", "da-DK")).toBe(3000);
    expect(typedNumber("3.000", "en-US")).toBe(3);
  });
});

describe("V3a: the pane's own boxes refuse them, which is how a person can reach typedNumber", () => {
  const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const typeInto = (el: HTMLInputElement, value: string) => {
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const boxes = (prefix: string): HTMLInputElement[] => {
    const row = [...document.querySelectorAll<HTMLElement>("#options label")].find((l) =>
      (l.textContent ?? "").startsWith(prefix),
    );
    if (!row) throw new Error(`no options row starting "${prefix}"`);
    return [...row.querySelectorAll<HTMLInputElement>('input[type="text"]')];
  };

  async function bootPane() {
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
  }

  it("Axis scale min and max: 0 500 and 0.500.500 set nothing, 1 500 and 1.500.500 set a number", async () => {
    await bootPane();
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({
      kind: "clustered",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
      numberFormat: { locale: "da-DK" },
    } satisfies Partial<ChartConfig>);
    $("json-import").click();
    const exported = (): ChartConfig => {
      $("json-export").click();
      return JSON.parse(($("json-io") as HTMLTextAreaElement).value);
    };
    const [min, max] = boxes("Axis scale min");
    typeInto(min, "0 500");
    typeInto(max, "0.500.500");
    expect(exported().scale, "a zero-led group set an axis bound").toBeUndefined();
    typeInto(min, "1 500");
    typeInto(max, "1.500.500");
    expect(exported().scale).toEqual({ min: 1500, max: 1500500 });
  });
});

/* -------------------------------------------------------------------------- */

/**
 * THE SKILL AS IT SHIPS: `scripts/render-pptx.mjs` beside a `lib/ssf-charts.js`
 * built from THIS source tree, in a folder of its own.
 *
 * Built here and not read from `dist-lib/`, which is a build artifact that a
 * local run after a change to `src/` finds stale (CLAUDE.md, Gotchas): a test of
 * what the skill stamps cannot depend on whether somebody remembered
 * `npm run build:lib`. The packaged layout is the one `render-pptx.mjs` tries
 * first (`../lib/ssf-charts.js`), `node_modules` is linked in so its imports
 * (pptxgenjs, jszip, and resvg for image mode) resolve, and nothing is written
 * inside the repository.
 */
let builtLib: string | undefined;
/** Every folder this file made in the OS temp directory, so they are not left behind run after run. */
const made: string[] = [];
function discard(root: string): void {
  try {
    const link = join(root, "node_modules");
    if (existsSync(link)) {
      // The link goes first and alone: a recursive delete that followed it would
      // empty the repository's own node_modules. Anything that is not the link
      // this file made is left where it is, and so is the folder around it.
      if (!lstatSync(link).isSymbolicLink()) return;
      unlinkSync(link);
    }
    rmSync(root, { recursive: true, force: true });
  } catch {
    /* a temp folder the OS reaps in its own time */
  }
}

function buildLib(): string {
  if (builtLib) return builtLib;
  const outDir = mkdtempSync(join(tmpdir(), "pc-v3-lib-"));
  made.push(outDir);
  // Vite's entry, not the `.bin` shim (AppLocker on the owner's box); and a clean
  // environment, because vitest's NODE_ENV=test is not a mode vite builds in.
  const r = spawnSync(
    process.execPath,
    ["node_modules/vite/bin/vite.js", "build", "--config", "vite.config.lib.ts", "--outDir", outDir, "--emptyOutDir"],
    { encoding: "utf8", env: { ...process.env, NODE_ENV: "production" } },
  );
  const file = join(outDir, "ssf-charts.js");
  if (r.status !== 0 || !existsSync(file)) throw new Error(`building the lib failed: ${r.stderr}\n${r.stdout}`);
  return (builtLib = file);
}

/** A skill folder, with the lib as `transform` leaves it, and the root to run it from. */
function packageSkill(transform: (js: string) => string = (js) => js): string {
  const root = mkdtempSync(join(tmpdir(), "pc-v3-skill-"));
  made.push(root);
  mkdirSync(join(root, "skill", "scripts"), { recursive: true });
  mkdirSync(join(root, "skill", "lib"), { recursive: true });
  for (const f of ["render-pptx.mjs", "pptx-paint.mjs"])
    copyFileSync(join("skill", "scripts", f), join(root, "skill", "scripts", f));
  writeFileSync(join(root, "skill", "lib", "ssf-charts.js"), transform(readFileSync(buildLib(), "utf8")));
  symlinkSync(resolve("node_modules"), join(root, "node_modules"), "junction");
  return root;
}

/** Run the packaged skill over some configs and open the deck it wrote. */
async function writeDeck(root: string, configs: unknown[], name = "deck"): Promise<JSZip> {
  const input = join(root, `${name}.json`);
  const file = join(root, `${name}.pptx`);
  writeFileSync(input, JSON.stringify(configs));
  const r = spawnSync(process.execPath, [join(root, "skill", "scripts", "render-pptx.mjs"), input, file], {
    encoding: "utf8",
  });
  if (r.status !== 0) throw new Error(`render-pptx failed: ${r.stderr}`);
  return JSZip.loadAsync(readFileSync(file));
}

const unxml = (s: string) =>
  s
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");

/** The tags a slide's chart carries, read the way the host reads them: slide rels -> tags part -> `<p:tag name val>`. */
async function tagsOf(zip: JSZip, n: number): Promise<Record<string, string>> {
  const rels = (await zip.file(`ppt/slides/_rels/slide${n}.xml.rels`)?.async("string")) ?? "";
  const out: Record<string, string> = {};
  for (const target of rels.matchAll(/Target="\.\.\/tags\/(tag\d+\.xml)"/g)) {
    const part = await zip.file(`ppt/tags/${target[1]}`)!.async("string");
    for (const tag of part.matchAll(/<p:tag name="([^"]*)" val="([^"]*)"\/>/g)) out[tag[1]] = unxml(tag[2]);
  }
  return out;
}

/** One shape of the chart's group as the file spells it: name, kind and box in points. */
interface FileShape {
  name: string;
  textBox: boolean;
  left: number;
  top: number;
  width: number;
  height: number;
}

/** The children of a slide's one group, in the order the file lists them. */
async function groupMembers(zip: JSZip, n: number): Promise<FileShape[]> {
  const xml = await zip.file(`ppt/slides/slide${n}.xml`)!.async("string");
  const grp = new DOMParser().parseFromString(xml, "application/xml").getElementsByTagName("p:grpSp")[0];
  if (!grp) throw new Error(`slide ${n} carries no group`);
  const members: FileShape[] = [];
  for (const el of Array.from(grp.children)) {
    if (!["p:sp", "p:cxnSp", "p:pic"].includes(el.tagName)) continue;
    const off = el.getElementsByTagName("a:off")[0];
    const ext = el.getElementsByTagName("a:ext")[0];
    members.push({
      name: el.getElementsByTagName("p:cNvPr")[0].getAttribute("name") ?? "",
      textBox: el.getElementsByTagName("p:cNvSpPr")[0]?.getAttribute("txBox") === "1",
      left: Number(off.getAttribute("x")) / 12700,
      top: Number(off.getAttribute("y")) / 12700,
      width: Number(ext.getAttribute("cx")) / 12700,
      height: Number(ext.getAttribute("cy")) / 12700,
    });
  }
  return members;
}

/**
 * LOAD A SKILL-WRITTEN SLIDE AS THE PANE WOULD FIND IT: one group named
 * PowerChart, holding the file's shapes in the file's order under the file's
 * names and sizes, and carrying every tag the file gave it. The fake host has no
 * .pptx reader of its own (`insertSlidesFromBase64` makes one placeholder shape
 * and the config tag), and a fixture that built the group from the SCENE would
 * agree with the scene by construction and prove nothing about the file.
 */
function slideFromFile(members: FileShape[], tags: Record<string, string>): { slide: FakeSlide; group: FakeShape } {
  const slide = makeSlide("s1");
  installHost([slide], [], slide, () => true);
  const shapes = members.map((m) => {
    const box = { left: m.left, top: m.top, width: m.width, height: m.height };
    const s = m.textBox ? slide.shapes.addTextBox("", box) : slide.shapes.addGeometricShape("rectangle", box);
    s.name = m.name;
    return s;
  });
  const group = slide.shapes.addGroup(shapes);
  group.name = "PowerChart";
  for (const [k, v] of Object.entries(tags)) group.tags.add(k, v);
  // Everything above is the file's content, not a batch awaiting a sync.
  slide.pending.length = 0;
  return { slide, group };
}

const clustered = (): ChartConfig => JSON.parse(JSON.stringify(sampleConfig("clustered"))) as ChartConfig;
const edited = (cfg: ChartConfig, change: (c: ChartConfig) => void): ChartConfig => {
  const next = JSON.parse(JSON.stringify(cfg)) as ChartConfig;
  change(next);
  return next;
};

describe("V3b: the skill stamps the scene fingerprint, and the add-in treats it as the one it would have written", () => {
  let deck: JSZip;
  let pictureDeck: JSZip;
  // Every kind the sampler knows, so the equality of the two scenes is swept and not sampled.
  const KINDS: ChartKind[] = CHART_KINDS.map((k) => k.kind);
  const cfgOf = (kind: ChartKind): ChartConfig => ({ ...sampleConfig(kind), title: "Sales" });

  beforeAll(async () => {
    const root = packageSkill();
    deck = await writeDeck(root, KINDS.map(cfgOf));
    pictureDeck = await writeDeck(
      root,
      [
        { ...cfgOf("clustered"), render: "image" },
        { kind: "agenda", chapters: ["A", "B"], highlight: 0 },
      ],
      "picture",
    );
  }, 240000);

  afterAll(() => {
    for (const root of made) discard(root);
  });

  it("src/index.ts exports sceneFingerprint, and it is the one the add-in uses", () => {
    expect(typeof lib.sceneFingerprint).toBe("function");
    const scene = buildChart(clustered());
    expect(lib.sceneFingerprint(scene)).toBe(sceneFingerprint(scene));
  });

  it("stamps, on every chart drawn as shapes, the value the update path recomputes from the stored config", async () => {
    // A tag that is merely PRESENT would move the add-in's refusal one reason along
    // (`pptx-deck.test.ts`, the same check for the add-in's own deck builder). The
    // skill draws `buildChart({ ...DEFAULT_SIZE, ...cfg })` and stores `cfg`, and
    // the add-in rebuilds from the stored text alone: for every kind the two scenes
    // have to be the same one.
    for (let i = 0; i < KINDS.length; i++) {
      const tags = await tagsOf(deck, i + 1);
      const kind = KINDS[i];
      expect(Object.keys(tags), `${kind} lost its config or origin`).toEqual(
        expect.arrayContaining(["POWERCHART_CONFIG", "POWERCHART_ORIGIN"]),
      );
      expect(tags.POWERCHART_SCENE, `${kind} carries no scene fingerprint`).toBeTruthy();
      expect(tags.POWERCHART_SCENE, `${kind}: not what the update path recomputes`).toBe(
        sceneFingerprint(buildChart(JSON.parse(tags.POWERCHART_CONFIG) as ChartConfig)),
      );
      expect(tags.POWERCHART_SCENE, `${kind}: not the scene the skill drew`).toBe(
        sceneFingerprint(buildChart({ ...DEFAULT_SIZE, ...cfgOf(kind) })),
      );
    }
  });

  it("writes it beside the config in the one tags part, and changes nothing else in the slide", async () => {
    // The stamp is one more `<p:tag>` in the part the slide already references: the
    // slide XML and its relationships name the same part either way, so a deck read
    // by any other tool differs by that one tag and nothing it follows.
    const part = await deck.file("ppt/tags/tag1.xml")!.async("string");
    expect(part.match(/<p:tag /g)).toHaveLength(3);
    expect(await deck.file("ppt/slides/_rels/slide1.xml.rels")!.async("string")).toContain("../tags/tag1.xml");
  });

  it("leaves an image-mode chart and an agenda without one", async () => {
    // A picture is not in the scene (`tryInPlaceUpdate` will not let the scene decide
    // an update that draws one), and an agenda carries no tag at all.
    const tags = await tagsOf(pictureDeck, 1);
    expect(tags.POWERCHART_CONFIG, "the picture lost its config").toBeTruthy();
    expect(tags.POWERCHART_SCENE).toBeUndefined();
    expect(await tagsOf(pictureDeck, 2)).toEqual({});
  });

  /**
   * The chart the pane would find, edited, and which route the edit took.
   * `change` is applied to the stored config to make the pane's next one.
   */
  async function editedInPlace(
    n: number,
    change: (c: ChartConfig) => void,
    tamper?: (t: Record<string, string>) => void,
  ) {
    const tags = await tagsOf(deck, n);
    tamper?.(tags);
    const { slide, group } = slideFromFile(await groupMembers(deck, n), tags);
    const members = [...(group.grouped as FakeShape[])];
    const idsBefore = members.map((m) => m.id);
    const target = (await listChartsInDeck()).charts[0].target;
    const stored = JSON.parse(tags.POWERCHART_CONFIG) as ChartConfig;
    const next = edited(stored, change);
    setTracing(true);
    await updateChartInSlide(buildChart(next), target, { tagData: JSON.stringify(next) });
    const route = traceLog().entries.some((e) => e.message === "updated only the shapes that changed")
      ? "in place"
      : "redraw";
    const declined = traceLog().entries.find((e) => e.message === "not updating in place — redrawing instead");
    return {
      slide,
      group,
      members,
      next,
      route,
      why: String((declined?.data as { why?: string } | undefined)?.why ?? ""),
      idsKept: members.every((m, i) => m.id === idsBefore[i] && !m.deleted),
      origin: JSON.parse(tags.POWERCHART_ORIGIN) as number[],
    };
  }

  /** The members of an edited chart, against the scene's rects and texts, in the one-for-one order. */
  const offShapes = (r: Awaited<ReturnType<typeof editedInPlace>>): string[] => {
    const scene: Scene = buildChart(r.next);
    const off: string[] = [];
    scene.nodes.forEach((n, i) => {
      if (n.kind !== "rect") return;
      const m = r.members[i];
      const want = [r.origin[0] + n.x, r.origin[1] + n.y, Math.max(0.2, n.w), Math.max(0.2, n.h)];
      const got = [m.left, m.top, m.width, m.height];
      if (want.some((v, k) => Math.abs(v - got[k]) > 0.5))
        off.push(`${n.name}: want ${want.map((v) => v.toFixed(1))} got ${got.map((v) => v.toFixed(1))}`);
    });
    return off;
  };

  it("a value edit on a skill-made chart takes the in-place route, and lands every bar where the scene puts it", async () => {
    const r = await editedInPlace(KINDS.indexOf("clustered") + 1, (c) => (c.data.series[0].values[1] = 30));
    expect(r.why, "the update declined").toBe("");
    expect(r.route, "the first edit of a skill-made chart was a full redraw").toBe("in place");
    expect(r.idsKept, "the shapes the file drew were replaced").toBe(true);
    expect(offShapes(r)).toEqual([]);
  });

  it("a retitle takes it too, on the column, stacked, line and waterfall charts", async () => {
    for (const kind of ["clustered", "stacked", "line", "waterfall"] as ChartKind[]) {
      const r = await editedInPlace(KINDS.indexOf(kind) + 1, (c) => (c.title = "Renamed"));
      expect(r.why, `${kind}: the update declined`).toBe("");
      expect(r.route, `${kind}: the first edit was a full redraw`).toBe("in place");
      expect(r.idsKept, `${kind}: the shapes the file drew were replaced`).toBe(true);
      expect(r.members[0].text, `${kind}: the title was not written`).toBe("Renamed");
    }
  });

  it("on no kind is the fingerprint the reason an edit is refused, and an in-place edit leaves every bar where the scene puts it", async () => {
    // The whole claim rests on the skill drawing one shape per node, in node order
    // (`checkShapesAgainstScene` is silent about the names pptxgenjs gives), so it is
    // asked of every kind and not of the four above. A kind the differ cannot write
    // to is refused for THAT, which is as safe as it ever was: what must never be
    // the answer is the fingerprint, and what must never happen is a wrong write.
    const wrote: string[] = [];
    for (let i = 0; i < KINDS.length; i++) {
      const kind = KINDS[i];
      const edits: [string, (c: ChartConfig) => void][] = [
        ["retitle", (c) => (c.title = "Renamed")],
        [
          "value",
          (c) => {
            const values = c.data.series[0].values;
            if (typeof values[1] === "number") values[1] += 7;
          },
        ],
      ];
      for (const [what, change] of edits) {
        const r = await editedInPlace(i + 1, change);
        expect(r.why, `${kind} ${what}: refused for its fingerprint`).not.toMatch(/fingerprint|no longer renders/);
        if (r.route === "in place") {
          wrote.push(`${kind} ${what}`);
          expect(r.idsKept, `${kind} ${what}: shapes were replaced on the in-place route`).toBe(true);
          expect(offShapes(r), `${kind} ${what}: a bar is not where the scene puts it`).toEqual([]);
        }
      }
    }
    // A sweep that took no edit in place would pass the loop above and prove nothing.
    expect(wrote.length, "no kind took the in-place route").toBeGreaterThanOrEqual(KINDS.length);
  });

  it("is the fingerprint that lets it: the same chart without the tag is redrawn, and says why", async () => {
    const r = await editedInPlace(
      KINDS.indexOf("clustered") + 1,
      (c) => (c.title = "Renamed"),
      (t) => delete t.POWERCHART_SCENE,
    );
    expect(r.route).toBe("redraw");
    expect(r.why).toMatch(/no scene fingerprint/);
  });

  it("is judged by its VALUE: a fingerprint of some other scene sends the edit to a redraw", async () => {
    const r = await editedInPlace(
      KINDS.indexOf("clustered") + 1,
      (c) => (c.title = "Renamed"),
      (t) => (t.POWERCHART_SCENE = sceneFingerprint(buildChart(sampleConfig("stacked")))),
    );
    expect(r.route).toBe("redraw");
    expect(r.why).toMatch(/no longer renders to the scene that was drawn/);
  });

  it("a lib that does not export it still writes the deck, with no fingerprint, as before", async () => {
    // A packaged skill ships no build step, and a stale lib beside a newer script is
    // a state a person can reach: it must degrade to what the script always wrote
    // and not turn every chart into an error slide.
    const stale = packageSkill((js) => {
      // The one `export { … }` clause a bundled ES lib ends with, and only inside it.
      const out = js.replace(/export \{[^}]*\};?\s*$/, (clause) =>
        clause.replace(/(?:\w+ as )?sceneFingerprint,? ?/, ""),
      );
      if (out === js) throw new Error("the built lib has no sceneFingerprint export to take out");
      return out;
    });
    const zip = await writeDeck(stale, [cfgOf("clustered")]);
    const tags = await tagsOf(zip, 1);
    expect(Object.keys(tags).sort()).toEqual(["POWERCHART_CONFIG", "POWERCHART_ORIGIN"]);
    expect(await zip.file("ppt/slides/slide1.xml")!.async("string")).toContain('name="PowerChart"');
  });
});
