// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
// @ts-expect-error — plain .mjs auditor, no types.
import { readDeckBytes, faultsIn } from "../scripts/verify-deck.mjs";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import {
  canonicalSlideSize,
  parseSlideSizeEmu,
  injectGroupsAndTags,
  groupSlideShapes,
  tagFirstShape,
  tagPart,
  topLevelElements,
  xmlAttr,
} from "../src/render/ooxml";
import { sceneFingerprint } from "../src/core/scene-diff";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";

/** A minimal slide tree of the shape pptxgenjs writes. */
function slideXml(shapes: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `${shapes}</p:spTree></p:cSld></p:sld>`
  );
}

/** One shape with a frame, at EMU coordinates. */
function sp(id: number, x: number, y: number, cx: number, cy: number, inner = ""): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="part-${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>${inner}</p:sp>`
  );
}

describe("ooxml: splitting a shape tree", () => {
  it("returns each top-level shape, not the first closing tag it meets", () => {
    // A shape that CONTAINS a p:sp is exactly what breaks a textual
    // <p:sp>…</p:sp> match: it pairs the outer open with the inner close.
    const nested = `<p:sp><p:x>${sp(9, 0, 0, 1, 1)}</p:x></p:sp>`;
    const parts = topLevelElements(`${sp(2, 0, 0, 10, 10)}${nested}${sp(3, 0, 0, 10, 10)}`);
    expect(parts).toHaveLength(3);
    expect(parts[1]).toBe(nested);
  });

  it("keeps a self-closing element as its own top-level entry", () => {
    expect(topLevelElements(`<p:a/><p:b><p:c/></p:b>`)).toEqual(["<p:a/>", "<p:b><p:c/></p:b>"]);
  });
});

describe("ooxml: grouping a slide's shapes", () => {
  it("frames the group on its children's bounding box", () => {
    const xml = groupSlideShapes(slideXml(sp(2, 100, 200, 300, 400) + sp(3, 50, 900, 100, 100)));
    const off = /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff/.exec(xml)!;
    expect(off.slice(1)).toEqual(["50", "200", "350", "800"]); // x 50..400, y 200..1000
  });

  it("repeats the frame as chOff/chExt, so the group transform is an identity", () => {
    // PowerPoint's own grouping does this, and it is why children keep their
    // absolute coordinates. Any other child extent silently rescales the chart.
    const xml = groupSlideShapes(slideXml(sp(2, 100, 200, 300, 400) + sp(3, 50, 900, 100, 100)));
    const m =
      /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(\d+)" y="(\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/.exec(
        xml,
      )!;
    expect([m[5], m[6], m[7], m[8]]).toEqual([m[1], m[2], m[3], m[4]]);
  });

  it("leaves every child's own coordinates untouched", () => {
    const shapes = sp(2, 100, 200, 300, 400) + sp(3, 50, 900, 100, 100);
    const xml = groupSlideShapes(slideXml(shapes));
    expect(xml).toContain(shapes);
  });

  it("gives the group an id no shape on the slide is using", () => {
    const xml = groupSlideShapes(slideXml(sp(2, 0, 0, 1, 1) + sp(7, 0, 0, 1, 1)));
    expect(xml).toContain(`<p:cNvPr id="8" name="PowerChart"/>`);
  });

  it("hangs the config tag off the group, where the live renderer puts it", () => {
    const xml = groupSlideShapes(slideXml(sp(2, 0, 0, 1, 1) + sp(3, 0, 0, 1, 1)), "rId4");
    expect(xml).toContain(`<p:nvPr><p:custDataLst><p:tags r:id="rId4"/></p:custDataLst></p:nvPr>`);
  });

  it("does not group a single shape — there is nothing to hold together", () => {
    const one = slideXml(sp(2, 0, 0, 1, 1));
    expect(groupSlideShapes(one)).toBe(one);
  });

  it("refuses a tree it does not recognise instead of writing a broken slide", () => {
    expect(() => groupSlideShapes("<p:sld/>")).toThrow(/no <p:spTree>/);
  });

  it("takes the bounding box of a slide too big to spread", () => {
    // The box came from `Math.min(...frames.map(…))`, which passes one ARGUMENT
    // per shape — so the bound came from the data and a big enough slide threw
    // "Maximum call stack size exceeded" rather than returning a wrong answer.
    // Same class the datasheet's formula engine carries two comments about.
    //
    // The trigger is synthetic: it takes ~150k top-level shapes, and no chart
    // this engine lays out comes near that. It is guarded because the loop
    // costs the same and had no bound at all, and this module is handed decks
    // by a library caller.
    const n = 150_000;
    const body = Array.from({ length: n }, (_, i) => sp(i + 2, i * 100, 0, 100, 100)).join("");
    const xml = groupSlideShapes(slideXml(body));
    expect(xml).not.toBe(slideXml(body));
    // …and the box it took is right: x from 0 to the last shape's right edge.
    expect(xml).toContain(`<a:off x="0" y="0"/><a:ext cx="${(n - 1) * 100 + 100}" cy="100"/>`);
    // 0.8 s idle, measured 2026-09-13, building a deliberately oversized slide.
    // Covered by the 20s default in `vitest.config.ts`, which explains why that
    // default is not vitest's 5s.
  });
});

/**
 * The chart's text alternative, which the generated deck was the one output of
 * the three not to carry.
 *
 * `skill/reference.md` recorded it as a limit on the grounds that pptxgenjs
 * exposes alt text on pictures and native charts only — true of pptxgenjs, and
 * one step short: this module hand-patches the slide XML and writes the very
 * element alt text lives on, the group's `<p:cNvPr descr>`.
 */
describe("ooxml: the group carries the chart's alt text", () => {
  const shapes = sp(2, 100, 200, 300, 400) + sp(3, 50, 900, 100, 100);

  it("writes the description onto the group it creates", () => {
    const xml = groupSlideShapes(slideXml(shapes), undefined, "clustered column chart. 2 categories: Q1, Q2.");
    expect(xml).toContain(`name="PowerChart" descr="clustered column chart. 2 categories: Q1, Q2."`);
  });

  it("escapes it, because a chart title is user text", () => {
    const xml = groupSlideShapes(slideXml(shapes), undefined, 'Sales < Plan & "growth"');
    expect(xml).toContain(`descr="Sales &lt; Plan &amp; &quot;growth&quot;"`);
    expect(xml).not.toContain('descr="Sales <');
  });

  it("omits the attribute entirely when there is nothing to say", () => {
    // Not `descr=""`: an empty description is not "no alt text", it is alt text
    // that says nothing, which reads to a screen-reader user as a described
    // object whose description is blank.
    expect(groupSlideShapes(slideXml(shapes))).not.toContain("descr=");
    expect(groupSlideShapes(slideXml(shapes), undefined, "")).not.toContain("descr=");
  });
});

describe("ooxml: tags and deck size", () => {
  it("escapes a config JSON payload into an attribute", () => {
    const part = tagPart([["POWERCHART_CONFIG", `{"title":"A & <B>"}`]]);
    expect(part).toContain(`val="{&quot;title&quot;:&quot;A &amp; &lt;B&gt;&quot;}"`);
    expect(xmlAttr(`'`)).toBe("&apos;");
  });

  it("stamps the scene fingerprint into a generated deck", async () => {
    // WITHOUT THIS, EVERY DECK CHART IS PERMANENTLY REDRAW-ONLY. The in-place
    // update refuses a chart with no fingerprint — "it was drawn by an older
    // build" — and that message was wrong about these: they were drawn by THIS
    // builder, which had never written one.
    //
    // Rounds 143 and 144 both show charts 1/8, 2/8 and 3/8 getting past the
    // check while 4/8 through 8/8 refuse, and the split is exactly which charts
    // came from `insertSceneIntoSlide` and which arrived in the generated deck.
    // The fingerprint blocker only became visible at all once the group read
    // stopped the parts check short-circuiting ahead of it.
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    const built = await buildDeckBase64([{ scene, configJson: JSON.stringify(cfg) }]);
    // The tag parts are where both tags land, so read those rather than the
    // whole deck: a substring hit anywhere in the file would not prove the tag
    // reached a part a shape actually references.
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    const parts = await Promise.all(
      Object.keys(zip.files)
        // `zip.files` lists the DIRECTORY `ppt/tags/` too, and `zip.file()`
        // answers null for one — the prefix alone is not enough.
        .filter((n) => n.startsWith("ppt/tags/") && !zip.files[n].dir)
        .map((n) => zip.file(n)!.async("string")),
    );
    expect(parts.length, "the deck carries no tag parts at all").toBeGreaterThan(0);
    const xml = parts.join(" ");
    const valueOf = (name: string) => xml.match(new RegExp(`<p:tag name="${name}" val="([^"]*)"`))?.[1];

    const stored = valueOf("POWERCHART_SCENE");
    expect(stored, "the generated deck carries no scene fingerprint").toBeTruthy();
    // Beside the config, not instead of it — a shape may reference only one
    // tags part, so losing the config to gain the fingerprint would be a trade,
    // not a fix.
    const storedConfig = valueOf("POWERCHART_CONFIG");
    expect(storedConfig, "the config tag went missing").toBeTruthy();

    // AND IT HAS TO BE THE SAME VALUE THE UPDATE PATH RECOMPUTES. `tryInPlaceUpdate`
    // does `sceneFingerprint(buildChart(JSON.parse(tags.config)))` and refuses on a
    // mismatch, so a tag that is merely PRESENT would move the refusal one reason
    // along rather than end it — which is exactly how this bug hid behind the
    // parts check for 117 rounds.
    const decoded = storedConfig!.replace(/&quot;/g, '"').replace(/&amp;/g, "&");
    expect(
      sceneFingerprint(buildChart(JSON.parse(decoded) as ChartConfig)),
      "the stored fingerprint is not the one the update path will recompute",
    ).toBe(stored);
  });

  it("carries several tags in one part, since a shape may reference only one", () => {
    const part = tagPart([
      ["POWERCHART_CONFIG", "{}"],
      ["POWERCHART_ORIGIN", "[0,0,0,0]"],
    ]);
    expect(part.match(/<p:tag /g)).toHaveLength(2);
  });

  it("rewrites the slide size to PowerPoint's exact 16:9", () => {
    // pptxgenjs writes 12191695 for the same nominal 13.333in — 305 EMU short.
    // A deck whose size differs from the destination's invites a rescale on
    // insert, and a rescale moves every chart.
    const out = canonicalSlideSize(`<p:presentation><p:sldSz cx="12191695" cy="6858000"/></p:presentation>`);
    expect(out).toContain(`<p:sldSz cx="12192000" cy="6858000"/>`);
  });

  it("declares the DESTINATION's size when it is given one", () => {
    // The 16:9 default was applied unconditionally, which is the very mismatch
    // the test above says invites a rescale — just in the other direction. A
    // generated deck inserted into a 4:3 presentation declared 16:9, so the
    // host rescaled every slide and moved every chart, on every 4:3 deck.
    const out = canonicalSlideSize(`<p:presentation><p:sldSz cx="12191695" cy="6858000"/></p:presentation>`, {
      cx: 9144000,
      cy: 6858000,
    });
    expect(out).toContain(`<p:sldSz cx="9144000" cy="6858000"/>`);
  });

  it("will not WRITE a slide size it would refuse to READ", () => {
    // The two sides of this file disagreed and only one of them was checking.
    // `parseSlideSizeEmu` rejects a zero, negative, non-finite or non-integral
    // dimension — "a zero or non-finite dimension is not a slide" — while
    // `canonicalSlideSize` wrote whatever it was handed straight into the
    // attribute. None of these is a `ST_PositiveCoordinate`, so PowerPoint
    // offers to repair the deck, and `injectGroupsAndTags` is exported as
    // library API where this module's own comment says a contract must not be
    // kept by luck.
    //
    // 12192000.7 is the quiet one: a finite positive number that looks right in
    // a diff, and exactly what `points * EMU_PER_POINT` yields for a
    // non-integral point size.
    const pres = `<p:presentation><p:sldSz cx="12191695" cy="6858000"/></p:presentation>`;
    for (const emu of [
      { cx: NaN, cy: 6858000 },
      { cx: 0, cy: 0 },
      { cx: -12192000, cy: 6858000 },
      { cx: Infinity, cy: 6858000 },
      { cx: "wide", cy: 6858000 },
    ] as unknown as { cx: number; cy: number }[]) {
      const out = canonicalSlideSize(pres, emu);
      const label = JSON.stringify(emu);
      // The written size reads back as a slide, and the fallback is the
      // documented default rather than a throw that would fail an insert.
      expect(parseSlideSizeEmu(out), `${label} wrote a size that is not a slide`).toEqual({
        cx: 12192000,
        cy: 6858000,
      });
    }
    // A fractional EMU is rounded to the integer the schema wants, not discarded
    // — the caller's intent is legible and a rescale of half an EMU is nothing.
    expect(parseSlideSizeEmu(canonicalSlideSize(pres, { cx: 9144000.7, cy: 6858000.2 }))).toEqual({
      cx: 9144001,
      cy: 6858000,
    });
  });
});

describe("a generated deck matches the deck it is going into", () => {
  /** The <p:sldSz> a built deck declares. */
  async function declaredSize(base64: string) {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(base64, { base64: true });
    return parseSlideSizeEmu(await zip.file("ppt/presentation.xml")!.async("string"));
  }

  it("declares 16:9 when nothing says otherwise", async () => {
    const built = await buildDeckBase64([{ scene: buildChart(sampleConfig("line")), title: "A" }]);
    expect(await declaredSize(built.base64)).toEqual({ cx: 12192000, cy: 6858000 });
  }, 30_000);

  it("declares 4:3 — and re-centres for it — when the destination is 4:3", async () => {
    // Two bugs in one, and they compound. The file declared a size the
    // destination did not share, so the host rescaled every slide on insert;
    // and the centring divided by a 960pt width the destination did not have,
    // so the chart was off-centre before the rescale even started.
    const scene = buildChart(sampleConfig("line"));
    const wide = await buildDeckBase64([{ scene, title: "A" }], { width: 960, height: 540 });
    const narrow = await buildDeckBase64([{ scene, title: "A" }], { width: 720, height: 540 });
    expect(await declaredSize(narrow.base64)).toEqual({ cx: 9144000, cy: 6858000 });
    // Centred for 720pt, not 960: the chart sits further left than it would on
    // a wide slide. Comparing the two decks proves the geometry actually moved
    // rather than only the declaration changing.
    //
    // The RIGHTMOST offset, not the leftmost — both decks carry a background
    // shape at x=0, so a min() is 0 either way and would pass against a build
    // that never re-centred anything.
    const { default: JSZip } = await import("jszip");
    const rightmost = async (b64: string) => {
      const zip = await JSZip.loadAsync(b64, { base64: true });
      const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
      return Math.max(...[...xml.matchAll(/<a:off x="(\d+)"/g)].map((m) => Number(m[1])));
    };
    const narrowX = await rightmost(narrow.base64);
    const wideX = await rightmost(wide.base64);
    expect(narrowX).toBeLessThan(wideX);
    // And by the amount the centring arithmetic says: half the width difference.
    expect(wideX - narrowX).toBeCloseTo(((960 - 720) / 2) * 12700, -3);
  }, 30_000);
});

describe("ooxml: reading a deck's declared slide size", () => {
  it("reads cx and cy whatever order they are written in", () => {
    // This is how the add-in learns its own slide size on a host below
    // PowerPointApi 1.10 — off a .pptx the host exported. Attribute order is
    // not guaranteed by anything, so it is not assumed.
    expect(parseSlideSizeEmu(`<p:sldSz cx="9144000" cy="6858000"/>`)).toEqual({ cx: 9144000, cy: 6858000 });
    expect(parseSlideSizeEmu(`<p:sldSz cy="6858000" cx="12192000" type="screen16x9"/>`)).toEqual({
      cx: 12192000,
      cy: 6858000,
    });
  });

  it("answers null rather than a nonsense slide", () => {
    // A zero dimension would put a divide-by-zero into every aspect-ratio
    // calculation downstream, and "no answer" is a case the caller already
    // handles — it falls to the next rung.
    expect(parseSlideSizeEmu(`<p:presentation/>`)).toBeNull();
    expect(parseSlideSizeEmu(`<p:sldSz cx="0" cy="6858000"/>`)).toBeNull();
    expect(parseSlideSizeEmu(`<p:sldSz cx="12192000"/>`)).toBeNull();
  });
});

describe("ooxml: a slide with nothing to group", () => {
  it("tags the shape itself, so a one-picture chart is still re-editable", () => {
    // Image mode is one picture. There is no group to hang the config on, and
    // without this such a chart carries no POWERCHART_CONFIG at all — the
    // difference between a picture of a chart and a chart.
    const one = slideXml(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="img"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic>`);
    expect(groupSlideShapes(one)).toBe(one); // nothing to group
    expect(tagFirstShape(one, "rId7")).toContain(
      `<p:nvPr><p:custDataLst><p:tags r:id="rId7"/></p:custDataLst></p:nvPr>`,
    );
  });

  it("extends an nvPr that already has content rather than replacing it", () => {
    const withPh = slideXml(
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr></p:sp>`,
    );
    const out = tagFirstShape(withPh, "rId3");
    // The placeholder survives, and CT_ApplicationNonVisualDrawingProps wants
    // ph before custDataLst — which is the order this produces.
    expect(out).toContain(`<p:nvPr><p:ph type="body"/><p:custDataLst><p:tags r:id="rId3"/></p:custDataLst></p:nvPr>`);
  });

  it("leaves a slide it cannot read alone", () => {
    expect(tagFirstShape("<p:sld/>", "rId1")).toBe("<p:sld/>");
  });
});

describe("ooxml: dressing and slides must stay in step", () => {
  it("refuses a deck with more slides than dressing entries", async () => {
    // `dressing[i]` IS `slide{i+1}.xml` — the pairing is strictly positional.
    // A caller whose two counts drift apart puts every later chart's config on
    // the wrong slide, and the file still opens cleanly, so nothing surfaces
    // it until someone edits a chart and overwrites a different one. Too MANY
    // dressing entries already failed loudly on the missing slide part; too
    // few used to leave the trailing slides silently ungrouped and untagged.
    const built = await buildDeckBase64([
      { scene: buildChart(sampleConfig("line")), title: "A" },
      { scene: buildChart(sampleConfig("line")), title: "B" },
    ]);
    await expect(injectGroupsAndTags(built.base64, [{ title: "A" }])).rejects.toThrow(/2 slide\(s\) but 1 dressing/);
  }, 30_000);
});

describe("ooxml: parsing bytes another library wrote", () => {
  it("does not end a tag at a raw > inside an attribute value", () => {
    // pptxgenjs writes `typeface="…"` WITHOUT escaping it — the one unescaped
    // attribute in the library — and font names are free-form user strings.
    // Read naively, the `>` ends the tag early and the depth counter desyncs
    // for the rest of the slide: this returns nothing, `groupSlideShapes` sees
    // fewer than two children, and the chart is never grouped at all.
    const nasty = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/></p:nvSpPr><a:latin typeface="Ba>d"/></p:sp>`;
    const parts = topLevelElements(`${nasty}${sp(3, 0, 0, 10, 10)}`);
    expect(parts).toHaveLength(2);
    expect(parts[0]).toBe(nasty);
  });

  it("numbers new tag parts above the ones already in the deck", async () => {
    // `zip.file()` overwrites. Starting at tag1 unconditionally destroyed any
    // pre-existing tag part and every slide pointing at it — safe today only
    // because pptxgenjs happens to emit none.
    const built = await buildDeckBase64([
      { scene: buildChart(sampleConfig("line")), title: "A", configJson: `{"kind":"line"}` },
    ]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    zip.file("ppt/tags/tag1.xml", "<keep-me/>");
    const out = await injectGroupsAndTags(await zip.generateAsync({ type: "base64" }), [
      { configJson: `{"kind":"line"}`, slot: 0, title: "A" },
    ]);
    const after = await JSZip.loadAsync(out.base64, { base64: true });
    expect(await after.file("ppt/tags/tag1.xml")!.async("string")).toBe("<keep-me/>");
    expect(Object.keys(after.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f)).length).toBeGreaterThan(1);
  }, 30_000);
});

describe("building a deck in-process", () => {
  const cfg = (title: string): ChartConfig => ({ ...sampleConfig("line"), title });

  it("produces a .pptx whose charts are grouped, tagged and slot-stamped", async () => {
    const configJson = JSON.stringify(cfg("Line"));
    const built = await buildDeckBase64([
      { scene: buildChart(cfg("Line")), title: "Line", configJson, slot: 3 },
      { scene: buildChart(cfg("Second")), title: "Second", slot: 4 },
    ]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    // The count the caller verifies against has to be what THIS renderer drew,
    // not what the Office.js one would have.
    expect(built.shapesPerSlide).toHaveLength(2);
    expect(built.shapesPerSlide[0]).toBeGreaterThan(1);

    const slides = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    expect(slides).toHaveLength(2);

    const slide1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
    // One group per slide, holding the chart.
    expect(slide1.match(/<p:grpSp>/g)).toHaveLength(1);
    expect(slide1).toContain(`name="PowerChart"`);
    // The slot tag is a slide-level custDataLst, right after the shape tree.
    expect(slide1).toContain(`</p:spTree><p:custDataLst><p:tags r:id=`);

    // The config round-trips through the tag part unharmed.
    const tags = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
    expect(tags.length).toBeGreaterThanOrEqual(3); // config+origin, slot, slot
    const partTexts = await Promise.all(tags.map((t) => zip.file(t)!.async("string")));
    const configTag = partTexts.find((t) => t.includes("POWERCHART_CONFIG"))!;
    const val = /name="POWERCHART_CONFIG" val="([^"]*)"/.exec(configTag)![1];
    const decoded = val
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&");
    expect(JSON.parse(decoded)).toEqual(JSON.parse(configJson));

    // Every tag part is declared, or PowerPoint rejects the file.
    const types = await zip.file("[Content_Types].xml")!.async("string");
    for (const t of tags) {
      expect(types).toContain(`PartName="/${t}"`);
    }
    // …and reachable from its slide.
    const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    expect(rels.match(/relationships\/tags/g)).toHaveLength(2); // group tag + slot tag
    expect(rels).toContain("../tags/");

    const presentation = await zip.file("ppt/presentation.xml")!.async("string");
    expect(presentation).toContain(`<p:sldSz cx="12192000" cy="6858000"/>`);
  }, 30_000);

  it("leaves a chart without config ungrouped-but-slotted rather than inventing a tag", async () => {
    const built = await buildDeckBase64([{ scene: buildChart(cfg("Plain")), title: "Plain", slot: 0 }]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    const slide = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(slide).toContain("<p:grpSp>");
    // The group carries no tag reference — there was no config to write.
    expect(slide).toContain(`<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`);
    const tags = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
    expect(tags).toHaveLength(1);
    expect(await zip.file(tags[0])!.async("string")).toContain("POWERCHART_DEMO_SLOT");
  }, 30_000);
});

/**
 * The deck writer against strings a real title or config can carry.
 *
 * Everything here is assembled by string surgery into XML — the slot tag, the
 * config tag, the rels — so an unescaped metacharacter does not produce a bad
 * chart, it produces a file PowerPoint will not open. The bar is therefore not
 * "looks right": every part must parse as XML, carry no character XML forbids,
 * and audit clean under the same tool `npm run verify-deck` uses.
 *
 * This one found nothing, and is kept for that reason — it is a contract, and
 * deleting one escape from `xmlAttr` turns it red thirty-two times over.
 */
/** Strings a title or a config payload can really carry. */
const NASTY: [string, string][] = [
  ["xml metachars", `A & B <tag> "q" 'a' >`],
  ["cdata close", `]]> ends a CDATA`],
  ["entity-ish", `&amp;lt; &#x41; &notanentity;`],
  ["control chars", "bell \u0007 vtab \u000b formfeed \u000c"],
  ["null char", "before \u0000 after"],
  ["surrogate pair", `emoji 😀 and 🏳️‍🌈`],
  ["lone surrogate", "lone \ud800 half"],
  ["rtl override", "safe\u202egnip.exe"],
  ["xml decl", `<?xml version="1.0"?><root/>`],
  ["closing tag", `</a:t></a:r></a:p>`],
  ["rels injection", `" Target="../slides/slide1.xml" X="`],
  ["huge title", "x".repeat(50_000)],
  ["newlines", "line1\nline2\r\nline3"],
  ["tab", "a\tb"],
  ["only whitespace", "   "],
  ["empty", ""],
];

describe("ooxml: hostile strings in a title, a config or a run token", () => {
  it("still produces a well-formed, fault-free deck", async () => {
    const bad: string[] = [];
    for (const [name, s] of NASTY) {
      const cfg = { ...sampleConfig("clustered"), title: s };
      let built;
      try {
        built = await buildDeckBase64([
          { scene: buildChart(cfg), title: s, configJson: JSON.stringify(cfg), slot: 0, run: s },
        ]);
      } catch (e) {
        bad.push(`${name}: BUILD THREW ${(e as Error).message.slice(0, 60)}`);
        continue;
      }
      const bytes = Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0));
      // 1. It must be a readable zip whose parts are well-formed XML.
      let zip;
      try {
        zip = await JSZip.loadAsync(bytes);
      } catch (e) {
        bad.push(`${name}: NOT A ZIP ${(e as Error).message.slice(0, 40)}`);
        continue;
      }
      for (const path of Object.keys(zip.files).filter((f) => f.endsWith(".xml") || f.endsWith(".rels"))) {
        const xml = await zip.file(path)!.async("string");
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        if (doc.getElementsByTagName("parsererror").length) {
          bad.push(`${name}: ${path} is not well-formed XML`);
          break;
        }
        // A character XML forbids outright, however it got in.
        // eslint-disable-next-line no-control-regex -- the point is to catch these
        const illegal = xml.match(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/);
        if (illegal) bad.push(`${name}: ${path} carries a raw control char U+${illegal[0].charCodeAt(0).toString(16)}`);
      }
      // 2. The auditor must find no structural fault.
      try {
        const deck = await readDeckBytes(bytes);
        const faults = faultsIn(deck);
        if (faults.length) bad.push(`${name}: ${faults[0]}`);
      } catch (e) {
        bad.push(`${name}: AUDIT THREW ${(e as Error).message.slice(0, 50)}`);
      }
    }
    // Sliced so one broken escape prints a readable list, not 32 lines.
    expect(bad.slice(0, 20)).toEqual([]);
  }, 120_000);
});
