// @vitest-environment jsdom
/**
 * The generated deck, the SVG preview and the live add-in drawing one chart —
 * the regressions behind stream r4-deck-and-preview-parity. Each block names the
 * finding it pins and why the older tests could not see it.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { annularSectorPoints, arrowheadBox, dashKind, symbolPreset } from "../src/core/geometry";
import type { Scene, TextNode } from "../src/core/scene";
import type { ChartConfig, ChartKind } from "../src/core/types";
import { sceneToSvg } from "../src/render/svg";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { chartOnItsOwnSlideAsFile, errorText } from "../src/render/powerpoint";
import { StaleBuildError, isStaleBuild, lazy } from "../src/render/lazy";
import { slideShapesBounds } from "../src/render/ooxml";
import { toHex6 } from "../src/core/color";
import { hex, makeAddNode } from "../skill/scripts/pptx-paint.mjs";
import { installHost, makeSlide } from "./helpers/office-host";

/** Every part of a base64 .pptx, by path. */
async function unzip(base64: string): Promise<JSZip> {
  return JSZip.loadAsync(base64, { base64: true });
}

/** One slide's XML. */
async function slideXml(base64: string, n = 1): Promise<string> {
  const zip = await unzip(base64);
  return zip.file(`ppt/slides/slide${n}.xml`)!.async("string");
}

/** The value of one tag on the deck's first chart group, as the pane reads it back. */
async function tagValue(base64: string, name: string): Promise<string | undefined> {
  const zip = await unzip(base64);
  for (const path of Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f))) {
    const part = await zip.file(path)!.async("string");
    const hit = new RegExp(`name="${name}" val="([^"]*)"`).exec(part);
    if (hit) return hit[1].replace(/&quot;/g, '"').replace(/&amp;/g, "&");
  }
  return undefined;
}

/** The fill attribute of the first rect in a one-rect scene, as the preview draws it. */
const previewFill = (fill: string): string | undefined =>
  /fill="([^"]*)"/.exec(
    sceneToSvg({ width: 10, height: 10, nodes: [{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill }] }),
  )?.[1];

describe("r2-33 — the preview's colour allow-list reads what the decks read", () => {
  // The forms `toRgb` (src/core/color.ts) and `readHex` (pptx-paint.mjs) both
  // draw correctly. CSS function names are case-insensitive, a channel may carry
  // a sign (CSS clamps it, and wraps a negative hue), and `deg` is the hue's
  // default unit spelled out.
  const DRAWN = [
    "RGB(255, 0, 0)",
    "Rgb(0,128,255)",
    "RGBA(1, 2, 3, 0.5)",
    "HSL(120, 50%, 50%)",
    "hsl(210deg 60% 45%)",
    "hsl(210DEG 60% 45%)",
    "hsl(-30, 50%, 50%)",
    "rgb(300,-20,50)",
    "rgb(+10, 20, 30)",
  ];

  it("draws every one of them instead of falling back to black", () => {
    for (const c of DRAWN) {
      expect(previewFill(c), `${JSON.stringify(c)} fell back to black in the preview; both decks draw it`).toBe(c);
    }
  });

  it("does not paint a bar black under a label inked for the real colour", () => {
    // The worse half. In-bar value labels pick their ink from the REAL colour, so
    // a mid-light fill that the preview turned black got near-black text on a
    // black bar: invisible in the preview and in every picture insert.
    const cfg = (color: string): ChartConfig => ({
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [30], color },
          { name: "T", values: [50] },
        ],
      },
    });
    for (const color of ["HSL(120, 50%, 50%)", "hsl(120deg 50% 50%)"]) {
      const svg = sceneToSvg(buildChart(cfg(color)));
      expect(svg, `${color}: the series bar was painted black`).toContain(`fill="${color}"`);
      expect(svg).not.toContain('fill="#000000"');
    }
  });

  it("still refuses every breakout, with the wider class", () => {
    // What the widening may NOT admit: a quote to close the attribute, a bracket
    // to open a tag, an entity, a comment, any word but the one unit, and an
    // angle unit the other sinks would misread.
    for (const attack of [
      'RGB(1,2,3)"><script>alert(1)</script><rect fill="rgb(1,2,3',
      'rgb(1,2,3deg)" onload="alert(1)',
      'hsl(-1deg)"><image href=x onerror=alert(1)>',
      "rgb(-1,2,3)/**/",
      "hsl(1deg)/**/",
      "hsl(1deg;fill:red)",
      "rgb(from red r g b)",
      "rgb(var(--x))",
      "URL(javascript:alert(1))",
      "hsl(0.5turn 50% 50%)",
      "rgb(1e2,0,0)",
      "rgb(1&amp;2,3,4)",
    ]) {
      expect(previewFill(attack), `${JSON.stringify(attack)} survived the allow-list`).toBe("#000000");
    }
  });

  it("admits an angle unit only where the other sinks read it as that unit", () => {
    // The allow-list names `deg` and not `turn`, `rad` or `grad` because the shared
    // parser (`toRgb`) and the deck's (`readHex`) take the NUMBER and drop the
    // unit: `hsl(0.5turn …)` is read as half a degree by both. Drawing the true
    // colour in the preview would add a third answer, not remove one. So the
    // preview draws a form exactly when the colour sinks read it as it means —
    // and this goes red the day one of them learns a unit the allow-list lacks.
    for (const [form, meaning] of [
      ["hsl(210deg 60% 45%)", "hsl(210 60% 45%)"],
      ["hsl(0.5turn 50% 50%)", "hsl(180 50% 50%)"],
      ["hsl(3.14159rad 50% 50%)", "hsl(180 50% 50%)"],
      ["hsl(200grad 50% 50%)", "hsl(180 50% 50%)"],
    ]) {
      const sinksReadIt = toHex6(form) === toHex6(meaning) && `#${hex(form)}` === toHex6(meaning);
      expect(previewFill(form) === form, `${form}: the preview and the colour sinks disagree about it`).toBe(
        sinksReadIt,
      );
    }
  });

  it("builds a resolvable pattern id from an upper-case, signed colour", () => {
    const svg = sceneToSvg({
      width: 100,
      height: 100,
      nodes: [{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "RGB(78,-121,167)", pattern: "diagonal", name: "p" }],
    });
    const ref = /fill="url\(#([^"]+)\)"/.exec(svg);
    expect(ref).toBeTruthy();
    expect(ref![1]).toMatch(/^[\w.-]+$/);
    expect(svg).toContain(`<pattern id="${ref![1]}"`);
  });
});

/** A one-chart deck, as the add-in's generated-file paths build it. */
async function deckFor(cfg: ChartConfig, slide?: { width: number; height: number }) {
  const scene = buildChart(cfg);
  const built = await buildDeckBase64([{ scene, title: cfg.title ?? "Chart", configJson: JSON.stringify(cfg) }], slide);
  return { scene, ...built };
}

describe("r2-34 — a generated slide takes the chart's own canvas colour", () => {
  const dark = (): ChartConfig => ({
    kind: "clustered",
    ...DEFAULT_SIZE,
    title: "Revenue",
    style: { background: "#1e1e1e", text: "#f2f2f2", mutedText: "#c8c8c8" } as ChartConfig["style"],
    data: { categories: ["Q1", "Q2", "Q3"], series: [{ name: "A", values: [42, 51, 38] }] },
  });

  it("paints the slide with style.background, as the preview and the skill renderer do", async () => {
    const { base64 } = await deckFor(dark());
    const bg = /<p:bg>.*?<\/p:bg>/.exec(await slideXml(base64))?.[0];
    expect(bg, "a dark-styled chart landed on the forced-white slide the add-in always wrote").toContain(
      'val="1E1E1E"',
    );
  });

  it("stays white when the chart names no background, or one it cannot read", async () => {
    const bgOf = async (cfg: ChartConfig) => /<p:bg>.*?<\/p:bg>/.exec(await slideXml((await deckFor(cfg)).base64))?.[0];
    expect(await bgOf({ ...dark(), style: undefined })).toContain('val="FFFFFF"');
    // `hexOr`'s contract: an unrecognised paint is WHITE, never `hex`'s black.
    expect(await bgOf({ ...dark(), style: { background: "off-white" } as ChartConfig["style"] })).toContain(
      'val="FFFFFF"',
    );
    expect(await bgOf({ ...dark(), style: { background: "transparent" } as ChartConfig["style"] })).toContain(
      'val="FFFFFF"',
    );
  });

  it("lets a caller's own background win, and survives a config that is not JSON", async () => {
    const scene = buildChart(dark());
    const own = await buildDeckBase64([{ scene, configJson: JSON.stringify(dark()), background: "#336699" }]);
    expect(await slideXml(own.base64)).toContain('<a:srgbClr val="336699"/>');
    const broken = await buildDeckBase64([{ scene, configJson: "{not json" }]);
    expect(/<p:bg>.*?<\/p:bg>/.exec(await slideXml(broken.base64))?.[0]).toContain('val="FFFFFF"');
  });

  it("is what the own-slide offer hands the host (the shipped chartOnItsOwnSlideAsFile)", async () => {
    // The add-in's own caller, driven rather than copied: nothing on this path
    // ever set `DeckItem.background`, and the stalled-update slide swap builds
    // its item the same way.
    const cfg = dark();
    installHost([makeSlide("s1")]);
    const captured: string[] = [];
    const PP = (globalThis as unknown as { PowerPoint: { run: (cb: (ctx: unknown) => unknown) => unknown } })
      .PowerPoint;
    const run = PP.run;
    PP.run = (cb) =>
      run(async (ctx: unknown) => {
        const presentation = (ctx as { presentation: Record<string, unknown> }).presentation;
        if (!presentation.__wrapped) {
          const insert = presentation.insertSlidesFromBase64 as (b64: string, opts?: unknown) => void;
          presentation.insertSlidesFromBase64 = (b64: string, opts?: unknown) => {
            captured.push(b64);
            return insert.call(presentation, b64, opts);
          };
          presentation.__wrapped = true;
        }
        return cb(ctx);
      });
    try {
      await chartOnItsOwnSlideAsFile(buildChart(cfg), cfg);
    } finally {
      PP.run = run;
    }
    expect(captured, "the shipped path never reached the host's file insert").toHaveLength(1);
    expect(/<p:bg>.*?<\/p:bg>/.exec(await slideXml(captured[0]))?.[0]).toContain('val="1E1E1E"');
  });
});

describe("r1-20 — the origin anchor is where the group actually sits", () => {
  /** The group's `<a:off>` in points, from the slide the deck wrote. */
  async function groupAt(base64: string): Promise<[number, number]> {
    const xml = await slideXml(base64);
    const off = /<p:grpSp><p:nvGrpSpPr>.*?<p:grpSpPr><a:xfrm><a:off x="(-?\d+)" y="(-?\d+)"/.exec(xml)!;
    return [Number(off[1]) / 12700, Number(off[2]) / 12700];
  }
  const cfgOf = (kind: ChartKind, titled: boolean): ChartConfig => ({
    ...sampleConfig(kind),
    title: titled ? "Title" : undefined,
  });
  const SLIDE = { width: 960, height: 540 };

  it("reads an untitled chart's first edit as no drag at all", async () => {
    const { scene, base64 } = await deckFor(cfgOf("clustered", false), SLIDE);
    const origin = JSON.parse((await tagValue(base64, "POWERCHART_ORIGIN"))!) as number[];
    const [groupLeft, groupTop] = await groupAt(base64);
    // The precondition that makes this test able to fail: the ink does NOT start
    // at the frame origin, so a frame-origin anchor is a lie.
    expect(Math.abs(groupTop - origin[1]), "the fixture has no ink offset, so it proves nothing").toBeGreaterThan(10);
    expect(origin[2]).toBeCloseTo(groupLeft, 6);
    expect(origin[3]).toBeCloseTo(groupTop, 6);
    // And the update path's own arithmetic (powerpoint.ts: origin + (live - anchor))
    // lands the chart exactly where it already is.
    expect(origin[0] + (groupLeft - origin[2])).toBeCloseTo(origin[0], 6);
    expect(origin[1] + (groupTop - origin[3])).toBeCloseTo(origin[1], 6);
    expect(scene.nodes.length).toBeGreaterThan(1);
  });

  it("holds for every kind, titled or not (a titled mekko is 2pt off, an untitled pie 112pt)", async () => {
    const kinds: ChartKind[] = [
      "clustered",
      "stacked",
      "pie",
      "doughnut",
      "waterfall",
      "mekko",
      "sunburst",
      "candlestick",
      "butterfly",
    ];
    for (const kind of kinds) {
      for (const titled of [true, false]) {
        const { base64 } = await deckFor(cfgOf(kind, titled), SLIDE);
        const origin = JSON.parse((await tagValue(base64, "POWERCHART_ORIGIN"))!) as number[];
        const [groupLeft, groupTop] = await groupAt(base64);
        const why = `${kind}${titled ? "" : " (untitled)"}`;
        expect(origin[2], `${why}: anchor left is not the group's`).toBeCloseTo(groupLeft, 6);
        expect(origin[3], `${why}: anchor top is not the group's`).toBeCloseTo(groupTop, 6);
      }
    }
  });

  it("measures a one-shape slide by that shape's own frame", () => {
    const sp = (x: number, y: number) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/>` +
      `<a:ext cx="100" cy="200"/></a:xfrm></p:spPr></p:sp>`;
    const tree = (shapes: string) =>
      `<p:sld><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree></p:cSld></p:sld>`;
    expect(slideShapesBounds(tree(sp(300, 400)))).toEqual({ x: 300, y: 400, cx: 100, cy: 200 });
    expect(slideShapesBounds(tree(sp(300, 400) + sp(50, 900)))).toEqual({ x: 50, y: 400, cx: 350, cy: 700 });
    expect(slideShapesBounds(tree(""))).toBeNull();
    expect(slideShapesBounds("not a slide")).toBeNull();
  });
});

describe("r1-47 — the generated deck carries the chart's alt text", () => {
  it("writes scene.desc as the group's descr", async () => {
    const cfg: ChartConfig = { ...sampleConfig("clustered"), title: undefined };
    const { scene, base64 } = await deckFor(cfg);
    expect(scene.desc, "the fixture has no text alternative, so it proves nothing").toBeTruthy();
    const xml = await slideXml(base64);
    const group = /<p:grpSp><p:nvGrpSpPr><p:cNvPr [^>]*>/.exec(xml)![0];
    expect(group).toContain(`name="PowerChart"`);
    expect(group, "the group has no descr: the one output of the three with no alt text").toContain(
      `descr="${scene.desc!.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}"`,
    );
  });
});

describe("r1-22 — a line break in a label is two lines in the preview too", () => {
  /** A scene holding one label, as a pasted Alt+Enter header (`Revenue`, break, `(EURm)`) reaches it. */
  const labelScene = (text: string, valign: TextNode["valign"] = "top"): Scene => ({
    width: 480,
    height: 300,
    nodes: [
      { kind: "text", x: 10, y: 100, w: 100, h: 14, text, fontSize: 10, color: "#111111", align: "left", valign },
    ],
  });
  /** The `<tspan>` texts and baselines of the label, or its bare text when it is one line. */
  const drawn = (text: string, valign: TextNode["valign"] = "top") => {
    const svg = sceneToSvg(labelScene(text, valign));
    const spans = [...svg.matchAll(/<tspan x="[^"]*" y="([^"]*)">([^<]*)<\/tspan>/g)].map((m) => ({
      y: Number(m[1]),
      text: m[2],
    }));
    return { svg, spans };
  };
  /** The paragraphs of the label in the generated deck. */
  async function paragraphs(text: string): Promise<string[]> {
    const { base64 } = await buildDeckBase64([{ scene: labelScene(text) }]);
    const xml = await slideXml(base64);
    return [...xml.matchAll(/<a:p>(.*?)<\/a:p>/g)].map((p) =>
      [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => t[1]).join(""),
    );
  }

  it("draws each line of a break as its own line, below the one before", () => {
    const { svg, spans } = drawn("FY24\n(Actual)");
    expect(
      spans.map((s) => s.text),
      "the preview put the break in one <text> and let SVG collapse it",
    ).toEqual(["FY24", "(Actual)"]);
    expect(spans[1].y - spans[0].y).toBeCloseTo(12, 6); // 1.2 x the 10pt font
    expect(svg).not.toContain("\n(Actual)");
    // A top-anchored block hangs from the top, exactly where a one-line label sits.
    expect(spans[0].y).toBe(110);
  });

  it("anchors the block the way a PowerPoint text box with that valign does", () => {
    const bottom = drawn("a\nb\nc", "bottom").spans.map((s) => s.y);
    const single = Number(/<text x="[^"]*" y="([^"]*)"/.exec(drawn("c", "bottom").svg)![1]);
    expect(bottom[2], "the last line sits where a one-line bottom label does").toBeCloseTo(single, 6);
    const middle = drawn("a\nb\nc", "middle").spans.map((s) => s.y);
    const centre = Number(/<text x="[^"]*" y="([^"]*)"/.exec(drawn("b", "middle").svg)![1]);
    expect(middle[1], "the middle line sits where a one-line middle label does").toBeCloseTo(centre, 6);
  });

  it("leaves a one-line label's markup exactly as it was", () => {
    const { svg, spans } = drawn("Q1");
    expect(spans).toEqual([]);
    expect(svg).toContain('y="110" font-size="10" fill="#111111" text-anchor="start">Q1</text>');
  });

  it("agrees with the deck on how many lines every spelling of a break makes", async () => {
    // EVERY spelling, not the one that was reported: Alt+Enter is LF (CHAR(10)),
    // a Word paste is VT, and the others are the rest of Unicode's line breaks.
    const spellings: [string, string][] = [
      ["LF", "FY24\n(Actual)"],
      ["CRLF", "FY24\r\n(Actual)"],
      ["CR", "FY24\r(Actual)"],
      ["VT", "FY24\u000b(Actual)"],
      ["FF", "FY24\u000c(Actual)"],
      ["NEL", "FY24\u0085(Actual)"],
      // Built, not written: the editor path that wrote this file turns a typed
      // U+2028/U+2029 escape into the raw character, which is a line terminator.
      ["LS", `FY24${String.fromCharCode(0x2028)}(Actual)`],
      ["PS", `FY24${String.fromCharCode(0x2029)}(Actual)`],
    ];
    for (const [name, text] of spellings) {
      expect(
        drawn(text).spans.map((s) => s.text),
        `${name}: preview`,
      ).toEqual(["FY24", "(Actual)"]);
      expect(await paragraphs(text), `${name}: deck`).toEqual(["FY24", "(Actual)"]);
    }
  });

  it("does not glue a Word line break's two words together", async () => {
    // `xmlText` deletes VT, so `Q1<VT>Report` was `Q1Report` in the preview and
    // the deck while the live add-in passed the break to PowerPoint.
    expect(drawn("Q1\u000bReport").svg).not.toContain("Q1Report");
    expect(await paragraphs("Q1\u000bReport")).toEqual(["Q1", "Report"]);
  });

  it("treats a trailing break as nothing, in both", async () => {
    // pptxgenjs does not split a text that ENDS in a newline, so the break used
    // to ride into one <a:t> as a raw CRLF.
    expect(drawn("abc\n").spans).toEqual([]);
    expect(drawn("abc\n").svg).toContain(">abc</text>");
    expect(await paragraphs("abc\n")).toEqual(["abc"]);
    const { base64 } = await buildDeckBase64([{ scene: labelScene("abc\r\n") }]);
    expect(await slideXml(base64)).not.toMatch(/<a:t>[^<]*[\r\n]/);
  });

  it("keeps an interior blank line in the deck and leaves the gap in the preview", async () => {
    expect(await paragraphs("a\n\nb")).toEqual(["a", "", "b"]);
    const { spans } = drawn("a\n\nb");
    expect(spans.map((s) => s.text)).toEqual(["a", "b"]);
    expect(spans[1].y - spans[0].y).toBeCloseTo(24, 6);
  });
});

describe("r2-65 — a full-turn ring is two circles, not an outline with a seam", () => {
  /** The `addShape` calls `makeAddNode` makes for one node. */
  function shapesFor(node: unknown): { type: string; opts: Record<string, unknown> }[] {
    const shapes: { type: string; opts: Record<string, unknown> }[] = [];
    const slide = {
      addShape: (type: string, opts: Record<string, unknown>) => void shapes.push({ type, opts }),
      addText: () => undefined,
    };
    makeAddNode({ dashKind, annularSectorPoints, symbolPreset, arrowheadBox })(slide, node, 0, 0);
    return shapes;
  }
  const wedge = (startAngle: number, endAngle: number) => ({
    kind: "wedge",
    cx: 100,
    cy: 100,
    r: 75,
    innerR: 40,
    startAngle,
    endAngle,
    fill: "#2a78d6",
    stroke: "#ffffff",
    strokeWidth: 1,
  });
  type Pt = { x: number; y: number; moveTo?: boolean; close?: boolean };

  it("draws a 360-degree annulus as two closed contours", () => {
    const [shape] = shapesFor(wedge(0, 360));
    expect(shape.type).toBe("custGeom");
    const pts = shape.opts.points as Pt[];
    expect(
      pts.filter((p) => p.moveTo),
      "one outline: the connecting segment is stroked",
    ).toHaveLength(2);
    expect(pts.filter((p) => p.close)).toHaveLength(2);
    // The first contour is closed BEFORE the second begins, so no segment joins
    // the outer ring to the inner one.
    const firstClose = pts.findIndex((p) => p.close);
    expect(pts[firstClose + 1].moveTo).toBe(true);
    // And the inner one is the hole: it sits at the inner radius and winds the other way.
    const px = (p: Pt) => Math.hypot(p.x / (1 / 72) - 75, p.y / (1 / 72) - 75);
    expect(px(pts[0])).toBeCloseTo(75, 3);
    expect(px(pts[firstClose + 1])).toBeCloseTo(40, 3);
  });

  it("leaves a partial annular sector as the single outline it is", () => {
    const [shape] = shapesFor(wedge(0, 90));
    const pts = shape.opts.points as Pt[];
    expect(pts.filter((p) => p.moveTo)).toHaveLength(1);
    expect(pts.filter((p) => p.close)).toHaveLength(1);
  });

  it("holds through the real deck: a one-group sunburst's ring has no seam", async () => {
    const cfg: ChartConfig = {
      kind: "sunburst",
      ...DEFAULT_SIZE,
      data: {
        categories: ["EMEA | Cloud", "EMEA | Licenses", "EMEA | Services"],
        series: [{ name: "Revenue", values: [40, 25, 35] }],
      },
    } as ChartConfig;
    const { scene, base64 } = await deckFor(cfg);
    const full = scene.nodes.filter((n) => n.kind === "wedge" && n.endAngle - n.startAngle >= 359.9 && n.innerR > 0);
    expect(full.length, "the fixture has no full-turn ring, so it proves nothing").toBeGreaterThan(0);
    const xml = await slideXml(base64);
    const paths = [...xml.matchAll(/<a:custGeom>[\s\S]*?<\/a:custGeom>/g)].map((m) => m[0]);
    const rings = paths.filter((p) => (p.match(/<a:moveTo>/g) ?? []).length === 2);
    expect(rings.length, "the full ring was written as one outline with a stroked radial seam").toBe(full.length);
    for (const ring of rings) expect((ring.match(/<a:close\s*\/>/g) ?? []).length).toBe(2);
  });
});

describe("r2-46 — a failed chunk fetch is not asserted to be a deploy", () => {
  // `import()` rejects with ONE message for a dropped connection, a DNS or proxy
  // failure, a 5xx and a 404 alike. Chromium's is word for word the stale-build
  // one, so no test on the message can tell them apart — which is the point.
  const chunkFailure = new TypeError(
    "Failed to fetch dynamically imported module: https://ssf-chart.example/assets/pptxgen.es-C8DOodSg.js",
  );

  it("names the connection as a cause, because the message cannot rule it out", () => {
    const said = errorText(chunkFailure);
    expect(said, "told the user a deploy happened when the browser said nothing of the kind").not.toMatch(
      /^SSF Charts has been updated/,
    );
    expect(said).toMatch(/connection/i);
    // Both causes, and the advice that is right for both.
    expect(said).toMatch(/was updated/i);
    expect(said).toMatch(/open it again/i);
    expect(said).toMatch(/slides are untouched/i);
    expect(said, "a chunk URL is not something the reader can act on").not.toMatch(/https?:|\.js\b/);
  });

  it("says the same through lazy() and the error it throws", async () => {
    const thrown = await lazy(() => Promise.reject(chunkFailure), "the deck writer").catch((e: unknown) => e);
    expect(thrown).toBeInstanceOf(StaleBuildError);
    expect((thrown as Error).message).toMatch(/connection/i);
    expect((thrown as Error).message).toContain("the deck writer");
  });

  it("still recognises all three families, and still leaves a bare fetch() failure alone", () => {
    expect(isStaleBuild(chunkFailure)).toBe(true);
    expect(isStaleBuild(new Error("error loading dynamically imported module"))).toBe(true);
    expect(isStaleBuild(new TypeError("Importing a module script failed."))).toBe(true);
    // That wording is `fetch()`'s, never `import()`'s: some other request failed.
    expect(isStaleBuild(new TypeError("Failed to fetch"))).toBe(false);
  });
});
