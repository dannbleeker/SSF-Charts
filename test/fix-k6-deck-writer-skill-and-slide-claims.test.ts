// @vitest-environment jsdom
/**
 * The generated deck, the headless skill and the slide it claims — the
 * regressions behind stream k6-deck-writer-skill-and-slide-claims (third bug
 * hunt). Each block names the finding it pins and why the older tests could not
 * see it.
 */
import { afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { layoutBackground } from "../src/core/color";
import { onTrace, setTracing } from "../src/core/trace";
import type { ChartConfig, ChartKind } from "../src/core/types";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { addSlideForChart } from "../src/render/powerpoint";
import { hexOr } from "../skill/scripts/pptx-paint.mjs";
import { faults, installHost, makeShape, makeSlide, type FakeSlide } from "./helpers/office-host";

afterEach(() => vi.unstubAllGlobals());

const LIB = "dist-lib/ssf-charts.js";

function ensureLib() {
  if (existsSync(LIB)) return;
  // Run vite's entry directly rather than the node_modules/.bin shim, which is
  // not executable on every dev box (Windows AppLocker).
  const r = spawnSync(process.execPath, ["node_modules/vite/bin/vite.js", "build", "--config", "vite.config.lib.ts"], {
    encoding: "utf8",
  });
  if (r.status !== 0) throw new Error(`build:lib failed: ${r.stderr}`);
}

const run = (args: string[]) => spawnSync(process.execPath, args, { encoding: "utf8" });
const PPTX = "skill/scripts/render-pptx.mjs";
const SVG = "scripts/render-batch.mjs";

const writeJson = (dir: string, name: string, value: unknown) => {
  const file = join(dir, name);
  writeFileSync(file, JSON.stringify(value));
  return file;
};

/** One slide's XML out of a written deck. */
const slideXmlOf = async (zip: JSZip, n: number) => zip.file(`ppt/slides/slide${n}.xml`)!.async("string");

/**
 * The tags a slide's shapes reference, by name — read the way the host reads
 * them: slide rels -> tags part -> `<p:tag name val>`.
 */
async function tagsOf(zip: JSZip, n: number): Promise<Record<string, string>> {
  const rels = (await zip.file(`ppt/slides/_rels/slide${n}.xml.rels`)?.async("string")) ?? "";
  const out: Record<string, string> = {};
  for (const target of rels.matchAll(/Target="\.\.\/tags\/(tag\d+\.xml)"/g)) {
    const part = await zip.file(`ppt/tags/${target[1]}`)!.async("string");
    for (const tag of part.matchAll(/<p:tag name="([^"]*)" val="([^"]*)"\/>/g)) {
      out[tag[1]] = tag[2].replace(/&quot;/g, '"').replace(/&amp;/g, "&");
    }
  }
  return out;
}

const bgOf = (xml: string) => /<p:bg>.*?<\/p:bg>/.exec(xml)?.[0] ?? "";

describe("r3-0 — a paint that draws nothing is not a slide background", () => {
  // Every spelling `layoutBackground` (src/core/color.ts) reads as "no canvas".
  // `rgba(0, 0, 0, 0)` is what a browser's computed style answers for it.
  const ZERO_ALPHA = [
    "rgba(0,0,0,0)",
    "rgba(0, 0, 0, 0)",
    "rgba(0,0,0,0%)",
    "rgb(0 0 0 / 0)",
    "rgb(0 0 0 / 0%)",
    "#00000000",
    "#0000",
    "hsla(0,0%,0%,0)",
    "hsl(0 0% 0% / 0)",
    "RGBA(0,0,0,0)",
    "transparent",
    "none",
  ];

  it("hexOr falls back for every zero-alpha spelling, not only the keyword", () => {
    for (const bg of ZERO_ALPHA) expect(hexOr(bg, "#ffffff"), `hexOr(${JSON.stringify(bg)})`).toBe("ffffff");
  });

  it("still paints a background that draws something, translucent or not", () => {
    expect(hexOr("rgba(0,0,0,0.5)", "#ffffff")).toBe("000000");
    expect(hexOr("#11111180", "#ffffff")).toBe("111111");
    expect(hexOr("#000000", "#ffffff")).toBe("000000");
    expect(hexOr("rgba(17,17,17,1)", "#ffffff")).toBe("111111");
    expect(hexOr("#111", "#ffffff")).toBe("111111");
  });

  it("agrees with the engine on every spelling of a canvas (layoutBackground)", () => {
    // The layouts decide the chart's ink from `layoutBackground`; the deck paints
    // the slide from `hexOr`. Two functions answering "what canvas is this" for
    // the same string is how a near-black chart ended up on a black slide, so the
    // sweep asks both about the same strings: whatever canvas the engine lays the
    // chart out against is the colour the slide must be.
    const spellings = [
      ...ZERO_ALPHA,
      "#111111",
      "#1e1e1e",
      "#ffffff",
      "red",
      "rgb(10, 20, 30)",
      "rgba(10, 20, 30, 0.4)",
      "hsl(210 50% 20%)",
      "off-white",
      "not-a-colour",
      "",
      "   ",
    ];
    for (const bg of spellings) {
      const laid = layoutBackground(bg);
      const canvas = laid === "transparent" ? "#ffffff" : laid;
      expect(hexOr(bg, "#ffffff"), `hexOr(${JSON.stringify(bg)}) vs the engine's canvas ${JSON.stringify(laid)}`).toBe(
        hexOr(canvas, "#ffffff"),
      );
    }
  });

  it("is white on the generated slide when the chart's background draws nothing", async () => {
    const base: ChartConfig = { ...sampleConfig("clustered"), ...DEFAULT_SIZE, title: "T" };
    const slideBg = async (background: string) => {
      const cfg = { ...base, style: { background } } as ChartConfig;
      const { base64 } = await buildDeckBase64([{ scene: buildChart(cfg), configJson: JSON.stringify(cfg) }]);
      return bgOf(await slideXmlOf(await JSZip.loadAsync(base64, { base64: true }), 1));
    };
    for (const bg of ["rgba(0,0,0,0)", "#00000000", "#0000", "hsla(0,0%,0%,0)", "rgb(0 0 0 / 0)", "transparent"]) {
      expect(await slideBg(bg), `slide for background ${JSON.stringify(bg)}`).toContain('val="FFFFFF"');
    }
    // And the dark canvas the same code path exists for is untouched.
    expect(await slideBg("#111111")).toContain('val="111111"');
  });
});

describe("r3-29 / r3-27 — an agenda is not a chart", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-k6-agenda-"));
  const out = join(dir, "agenda.pptx");
  let zip: JSZip;
  let status: number | null = null;

  beforeAll(async () => {
    ensureLib();
    const chart = { ...sampleConfig("clustered"), title: "Sales" };
    const input = writeJson(dir, "agenda.json", [
      { kind: "agenda", chapters: ["Context", "Findings", "Recommendation"], highlight: 1 },
      chart,
      // The deck-wide dark style copied onto an agenda, which takes no style.
      { kind: "agenda", chapters: ["A", "B"], highlight: 0, style: { background: "#111111" } },
      // `render: "image"` is for dense charts; an agenda is a handful of text boxes.
      { kind: "agenda", chapters: ["A", "B"], highlight: 0, render: "image" },
    ]);
    const r = run([PPTX, input, out]);
    status = r.status;
    if (r.status !== 0) throw new Error(`render-pptx failed: ${r.stderr}`);
    zip = await JSZip.loadAsync(readFileSync(out));
  }, 180000);

  it("writes no tag and no group on an agenda slide, so the pane has no chart to open there", async () => {
    expect(status).toBe(0);
    for (const n of [1, 3, 4]) {
      const xml = await slideXmlOf(zip, n);
      expect(await tagsOf(zip, n), `slide ${n} carries a tag`).toEqual({});
      expect(xml, `slide ${n} was grouped`).not.toContain("<p:grpSp>");
      expect(xml, `slide ${n} has a shape named like a chart`).not.toContain('name="PowerChart"');
    }
    // The agenda text is really there, as loose shapes.
    expect(await slideXmlOf(zip, 1)).toContain("Recommendation");
  });

  it("leaves the chart beside it grouped and tagged", async () => {
    const tags = await tagsOf(zip, 2);
    expect(Object.keys(tags)).toContain("POWERCHART_CONFIG");
    expect(await slideXmlOf(zip, 2)).toContain('name="PowerChart"');
    // No config in the file says "agenda": nothing the pane reads is data-less.
    const all = await Promise.all([1, 2, 3, 4].map((n) => tagsOf(zip, n)));
    expect(JSON.stringify(all)).not.toContain("agenda");
  });

  it("paints the agenda on white whatever style.background it was given", async () => {
    // `buildAgendaScene` takes no style and inks #0b0b0b / #52514E; a #111111
    // slide under it is near-black text on near-black.
    expect(bgOf(await slideXmlOf(zip, 3))).toContain('val="FFFFFF"');
  });

  it("draws an agenda as native shapes even when asked for an image", async () => {
    const xml = await slideXmlOf(zip, 4);
    expect(xml).not.toContain("<p:pic>");
    expect(xml).toContain("<p:sp>");
  });
});

describe("r3-31 / r3-33 / r3-69 — a skill-made chart carries the frame origin the add-in's deck writes", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-k6-origin-"));
  const out = join(dir, "origin.pptx");
  let zip: JSZip;
  // Untitled, so the ink starts INSIDE the frame: 2 to 112pt by kind, which is
  // the gap a frame-origin-less tag turned into a jump on the first edit.
  const KINDS: ChartKind[] = ["clustered", "stacked", "pie", "waterfall"];
  const cfgOf = (kind: ChartKind): ChartConfig => ({ ...sampleConfig(kind), title: undefined });

  beforeAll(async () => {
    ensureLib();
    const input = writeJson(dir, "origin.json", [
      ...KINDS.map(cfgOf),
      { ...sampleConfig("violin"), title: undefined, render: "image" },
    ]);
    const r = run([PPTX, input, out]);
    if (r.status !== 0) throw new Error(`render-pptx failed: ${r.stderr}`);
    zip = await JSZip.loadAsync(readFileSync(out));
  }, 180000);

  /** The tagged shape's own corner in points: the group, or the picture on an image slide. */
  const taggedShapeAt = (xml: string): [number, number] => {
    const grp = /<p:grpSp><p:nvGrpSpPr>.*?<p:grpSpPr><a:xfrm><a:off x="(-?\d+)" y="(-?\d+)"/.exec(xml);
    const pic = /<p:pic>.*?<a:off x="(-?\d+)" y="(-?\d+)"/s.exec(xml);
    const at = grp ?? pic!;
    return [Number(at[1]) / 12700, Number(at[2]) / 12700];
  };

  it("writes POWERCHART_ORIGIN, anchored at the group the file actually has", async () => {
    for (let i = 0; i < KINDS.length; i++) {
      const tags = await tagsOf(zip, i + 1);
      expect(Object.keys(tags), `${KINDS[i]}: the origin tag is missing`).toContain("POWERCHART_ORIGIN");
      const origin = JSON.parse(tags.POWERCHART_ORIGIN) as number[];
      const [left, top] = taggedShapeAt(await slideXmlOf(zip, i + 1));
      expect(origin[2], `${KINDS[i]}: anchor left is not the group's`).toBeCloseTo(left, 6);
      expect(origin[3], `${KINDS[i]}: anchor top is not the group's`).toBeCloseTo(top, 6);
    }
  });

  it("puts the frame origin where the add-in's own deck builder puts it (the same arithmetic)", async () => {
    for (let i = 0; i < KINDS.length; i++) {
      const cfg = cfgOf(KINDS[i]);
      const { base64 } = await buildDeckBase64([
        { scene: buildChart({ ...DEFAULT_SIZE, ...cfg }), configJson: JSON.stringify(cfg) },
      ]);
      const built = await JSZip.loadAsync(base64, { base64: true });
      const want = JSON.parse((await tagsOf(built, 1)).POWERCHART_ORIGIN) as number[];
      const got = JSON.parse((await tagsOf(zip, i + 1)).POWERCHART_ORIGIN) as number[];
      got.forEach((v, k) => expect(v, `${KINDS[i]}: origin[${k}]`).toBeCloseTo(want[k], 6));
    }
  });

  it("is not a frame-origin anchor in disguise: the untitled chart's ink starts inside the frame", async () => {
    // The precondition that makes the two tests above able to fail. If the
    // group started AT the frame origin, an origin of [frame, frame] would be
    // right and none of this would matter.
    const origin = JSON.parse((await tagsOf(zip, 1)).POWERCHART_ORIGIN) as number[];
    expect(Math.abs(origin[3] - origin[1]), "the fixture has no ink offset, so it proves nothing").toBeGreaterThan(10);
  });

  it("tags an image-mode chart's picture the same way", async () => {
    const n = KINDS.length + 1;
    const tags = await tagsOf(zip, n);
    expect(Object.keys(tags)).toContain("POWERCHART_ORIGIN");
    const origin = JSON.parse(tags.POWERCHART_ORIGIN) as number[];
    const [left, top] = taggedShapeAt(await slideXmlOf(zip, n));
    expect(origin[2]).toBeCloseTo(left, 6);
    expect(origin[3]).toBeCloseTo(top, 6);
    // A picture fills its own frame, so frame and anchor coincide — to the EMU
    // pptxgenjs rounds the picture's offset to (1/12700pt), not to the float.
    expect(origin[0]).toBeCloseTo(origin[2], 3);
    expect(origin[1]).toBeCloseTo(origin[3], 3);
  });
});

/** The slide-background rect of a preview, as the SVG writes it. */
const canvasRect = (svg: string) => /<rect width="100%" height="100%" fill="([^"]*)"\/>/.exec(svg)?.[1];

describe("r3-72 / r3-57 — the SVG preview is drawn on the chart's own background", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-k6-svg-"));
  const outDir = join(dir, "svg");
  const dark: ChartConfig = {
    ...sampleConfig("clustered"),
    title: "Dark deck",
    style: { background: "#111111", text: "#f2f2f2", mutedText: "#bbbbbb", axis: "#888888" },
  };

  beforeAll(() => ensureLib(), 180000);

  it("paints a dark-styled chart on its own canvas, not a hard-coded white", () => {
    const input = writeJson(dir, "dark.json", [
      dark,
      // Only the background named: the ink stays its default near-black, which is
      // unreadable on the slide, and the preview must show that rather than hide it.
      { ...sampleConfig("clustered"), title: "Only bg", style: { background: "#111111" } },
      { ...sampleConfig("clustered"), title: "Plain" },
      { kind: "agenda", chapters: ["A", "B"], highlight: 0, style: { background: "#111111" } },
    ]);
    const r = run([SVG, input, outDir]);
    expect(r.status, r.stderr).toBe(0);
    const svg = (name: string) => readFileSync(join(outDir, name), "utf8");
    expect(canvasRect(svg("01-dark-deck.svg"))).toBe("#111111");
    expect(canvasRect(svg("02-only-bg.svg"))).toBe("#111111");
    expect(canvasRect(svg("03-plain.svg")), "a chart with no background lost its white").toBe("#ffffff");
    expect(canvasRect(svg("04-agenda.svg")), "an agenda takes no style").toBe("#ffffff");
  });
});

describe("r3-73 / r3-75 — ordinary bad CLI input gets a message that says what to fix", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-k6-cli-"));
  const good = [{ ...sampleConfig("pie"), title: "Share" }];
  // By code point, not typed: an invisible character in a source file is the
  // very thing this test is about.
  const BOM = String.fromCharCode(0xfeff);

  beforeAll(() => ensureLib(), 180000);

  it("names an output folder that does not exist, instead of a Node stack dump", () => {
    const input = writeJson(dir, "good.json", good);
    const target = join(dir, "no-such-folder", "deck.pptx");
    const r = run([PPTX, input, target]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain("couldn't write");
    expect(r.stderr).toContain(target);
    expect(r.stderr).toContain("ENOENT");
    expect(r.stderr, "the raw uncaught-exception dump").not.toMatch(/node:fs|\n\s+at /);
    expect(existsSync(target)).toBe(false);
  });

  it("names an output path that is a folder", () => {
    const input = writeJson(dir, "good2.json", good);
    const asFolder = join(dir, "i-am-a-folder");
    mkdirSync(asFolder);
    const r = run([PPTX, input, asFolder]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain(asFolder);
    expect(r.stderr).toContain("folder");
    expect(r.stderr).not.toMatch(/node:fs|\n\s+at /);
  });

  it("names an svg output folder that is a file", () => {
    const input = writeJson(dir, "good3.json", good);
    const aFile = join(dir, "i-am-a-file");
    writeFileSync(aFile, "x");
    const r = run([SVG, input, aFile]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain(aFile);
    expect(r.stderr).not.toMatch(/node:fs|\n\s+at /);
  });

  it("reads a config saved with a UTF-8 byte-order mark", () => {
    const input = join(dir, "bom.json");
    writeFileSync(input, BOM + JSON.stringify(good), "utf8");
    const pptx = join(dir, "bom.pptx");
    const a = run([PPTX, input, pptx]);
    expect(a.status, a.stderr).toBe(0);
    expect(existsSync(pptx)).toBe(true);
    const svgDir = join(dir, "bom-svg");
    const b = run([SVG, input, svgDir]);
    expect(b.status, b.stderr).toBe(0);
  });

  it("says null when the file holds null, not 'a object'", () => {
    const input = join(dir, "null.json");
    writeFileSync(input, "null");
    const r = run([PPTX, input, join(dir, "null.pptx")]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain("got null");
    expect(r.stderr).not.toContain("a object");
    // The other primitives keep their wording.
    writeFileSync(input, "5");
    expect(run([PPTX, input, join(dir, "five.pptx")]).stderr).toContain("got a number");
  });

  it("prints usage for --help instead of looking for a file called --help", () => {
    for (const script of [PPTX, SVG]) {
      for (const flag of ["--help", "-h"]) {
        const r = run([script, flag]);
        expect(r.status, `${script} ${flag}`).toBe(0);
        expect(r.stdout).toContain("usage: node scripts/");
        expect(r.stderr).not.toContain("ENOENT");
      }
    }
  });
});

describe("r3-55 / r3-74 / r3-56 — the skill's docs say what the engine does", () => {
  const skillMd = readFileSync("skill/SKILL.md", "utf8");
  const reference = readFileSync("skill/reference.md", "utf8");
  const types = readFileSync("src/core/types.ts", "utf8");

  /** The members `interface Decorations` declares. */
  const decorationKeys = (() => {
    const start = types.indexOf("export interface Decorations {");
    const body = types.slice(start, types.indexOf("\n}\n", start));
    return new Set([...body.matchAll(/^ {2}([A-Za-z]+)\??:/gm)].map((m) => m[1]));
  })();

  it("lists only real `decorations` keys under the `decorations` object (segmentOrder is top level)", () => {
    // From the heading to wherever the list ends — the new "These go at the TOP
    // LEVEL" or, in the wording this replaced, the ", plus" that ran the
    // top-level keys on after the decorations ones.
    const sentence = /Decorations \((?:INSIDE the )?`decorations` object\):([\s\S]*?)(?:These go at the|, plus\s)/.exec(
      skillMd,
    )?.[1];
    expect(sentence, "SKILL.md no longer has the decorations list this test reads").toBeTruthy();
    const named = [...sentence!.matchAll(/`([A-Za-z]+)/g)].map((m) => m[1]);
    expect(named.length).toBeGreaterThan(8);
    for (const key of named) expect(decorationKeys.has(key), `\`${key}\` is not a key of Decorations`).toBe(true);
    expect(sentence).not.toContain("segmentOrder");
    // …and it is documented where the engine reads it.
    const topLevel = /These go at the\s+TOP LEVEL[\s\S]*?ignored without an error/.exec(skillMd)?.[0] ?? "";
    expect(topLevel).toContain("`segmentOrder`");
    expect(decorationKeys.has("segmentOrder")).toBe(false);
  });

  it("does not promise adjustable pie geometry the pptx does not have", () => {
    expect(reference).not.toMatch(/adjustable pie geometry/);
    expect(reference).toMatch(/custom-geometry polygon/);
  });

  it("says the world tile map is never auto-detected", () => {
    expect(reference).not.toMatch(/layout is auto-detected from the codes/);
    expect(reference).toMatch(/`world` is never\s+auto-detected/);
    expect(skillMd).toMatch(/`world` is never auto-detected/);
  });

  it("does not offer day numbers with data.dates to the calendar heatmap", () => {
    expect(reference).not.toMatch(/ISO strings or day numbers with `data\.dates`/);
    expect(reference).toMatch(/bare day number is NOT a date here/);
  });
});

describe("r3-34 — an extra slide is deleted only when a fresh read shows it blank", () => {
  afterEach(() => {
    faults.addLandsExtra = 0;
    faults.faultShapeGetCount = false;
    onTrace(undefined);
    setTracing(false);
  });

  const deckOf = (n: number): FakeSlide[] => Array.from({ length: n }, (_, i) => makeSlide(`s${i + 1}`));

  function listen() {
    const lines: { message: string; data?: Record<string, unknown> }[] = [];
    setTracing(true);
    onTrace((e) => lines.push({ message: e.message, data: e.data as Record<string, unknown> | undefined }));
    return lines;
  }

  /**
   * Make the FIRST slide the host appends arrive holding shapes — somebody
   * else's slide landing between the two listings, ahead of the one we added.
   * The deck array is the host's own, so intercepting `push` is intercepting
   * the add.
   */
  function firstNewSlideArrivesPopulated(deck: FakeSlide[]) {
    const push = deck.push.bind(deck);
    let done = false;
    deck.push = (...slides: FakeSlide[]) => {
      const n = push(...slides);
      for (const s of slides) {
        if (done || !s.id.startsWith("slide-")) continue;
        done = true;
        for (let i = 0; i < 25; i++)
          s.created.push(makeShape("geometric", "rectangle", { left: i, top: i, width: 10, height: 10 }));
      }
      return n;
    };
  }

  it("keeps a populated slide that the id diff names as new, and puts the chart on the last", async () => {
    const deck = deckOf(2);
    installHost(deck);
    faults.addLandsExtra = 1;
    firstNewSlideArrivesPopulated(deck);
    const lines = listen();
    const id = await addSlideForChart();
    expect(id, "no slide was claimed").toBeTruthy();
    expect(deck.length, "a slide that holds shapes was deleted").toBe(4);
    expect(deck[2].created.length, "the slide the user's co-author made lost its shapes").toBe(25);
    expect(deck[3].id, "the id names a slide other than the last").toBe(id);
    // Reported, not silent.
    const left = lines.find((l) => l.message.startsWith("left a new-looking slide in the deck"));
    expect(left?.data?.shapes).toBe(25);
    const claim = lines.find((l) => l.message.startsWith("claimed the appended slide for a chart"));
    expect(claim?.data).toMatchObject({ removed: 0, notBlank: 1 });
  });

  it("still takes back the blank twin, which is what the claim exists for", async () => {
    const deck = deckOf(2);
    installHost(deck);
    faults.addLandsExtra = 1;
    const lines = listen();
    const id = await addSlideForChart();
    expect(id).toBeTruthy();
    expect(deck.length, "the blank twin was left behind").toBe(3);
    expect(deck[2].id).toBe(id);
    const claim = lines.find((l) => l.message.startsWith("claimed the appended slide for a chart"));
    expect(claim?.data).toMatchObject({ removed: 1, notBlank: 0 });
  });

  it("deletes nothing it could not read: an unreadable shape count is not a zero", async () => {
    const deck = deckOf(2);
    installHost(deck);
    faults.addLandsExtra = 1;
    faults.faultShapeGetCount = true;
    const lines = listen();
    const id = await addSlideForChart();
    expect(id).toBeTruthy();
    expect(deck.length, "a slide was deleted without being shown to be blank").toBe(4);
    const left = lines.find((l) => l.message.startsWith("left a new-looking slide in the deck"));
    expect(left?.data?.shapes).toBe("unreadable");
  });
});
