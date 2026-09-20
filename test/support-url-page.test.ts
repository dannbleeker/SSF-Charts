import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import { inflateSync } from "zlib";

/**
 * THE PAGE A CERTIFICATION REVIEWER OPENS FIRST.
 *
 * Both production manifests submit `https://github.com/dannbleeker/SSF-Charts`
 * as their `<SupportUrl>`, and GitHub serves `README.md` at that URL. So the
 * README is not an internal document: it is store-facing, by submission rather
 * than by directory, and the rules `docs/STORE-LISTING.md` writes down for
 * "anything store-facing" reach it.
 *
 * Nothing read it that way before. Two things had drifted by 2026-09-20:
 *
 *   - `docs/gallery.png`, the hero image, was captured when the add-in was
 *     called **PowerChart**, and the page heading it photographed said so. The
 *     rename swept `index.html` and the README prose; it could not sweep pixels,
 *     because no grep can see them. (PowerChart is also Oracle Health's EHR
 *     product, so the stale name collides with a live one.)
 *   - the feature table's first column was headed "think-cell feature", which
 *     framed 168 rows — most of the file — as a comparison against a named
 *     competitor, on the page Microsoft's reviewer lands on. `STORE-LISTING.md`
 *     already forbids the mark on store-facing surfaces and its own checklist
 *     line ("Listing copy above is trademark-clean") had only ever been applied
 *     to the copy inside that file.
 *
 * Both are the same defect as the PowerPointApi floor in `manifest.test.ts`:
 * a user-facing page nothing checked. These are the checks.
 */
const repoFile = (name: string) => readFileSync(fileURLToPath(new URL(`../${name}`, import.meta.url)));
const repoText = (name: string) => repoFile(name).toString("utf8");

const SUPPORT_URL = "https://github.com/dannbleeker/SSF-Charts";

/** The mark this project may describe but must not wear. */
const COMPETITOR = /think-cell/i;

/**
 * Enough of a PNG reader to answer one question: where is the first ink?
 *
 * Only 8-bit truecolour is handled, which is what `docs/gallery.png` is; a
 * different format throws rather than passing quietly, because a guard that
 * cannot read its subject must not report "fine". `scripts/build-icons.mjs`
 * writes PNGs and there is no reader anywhere in the repo to reuse.
 */
function readRgbRows(png: Buffer): { width: number; height: number; row: (y: number) => Buffer } {
  if (!png.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    throw new Error("not a PNG");
  }
  let i = 8;
  let width = 0;
  let height = 0;
  const idat: Buffer[] = [];
  while (i < png.length) {
    const length = png.readUInt32BE(i);
    const type = png.toString("latin1", i + 4, i + 8);
    const data = png.subarray(i + 8, i + 8 + length);
    if (type === "IHDR") {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      if (data[8] !== 8 || data[9] !== 2) {
        throw new Error(`gallery.png is bit depth ${data[8]}, colour type ${data[9]} — this reader only does 8/2`);
      }
    }
    if (type === "IDAT") idat.push(data);
    i += 12 + length;
    if (type === "IEND") break;
  }
  const raw = inflateSync(Buffer.concat(idat));
  const bpp = 3;
  const stride = width * bpp;
  // Un-filtering is sequential — row y needs row y-1 — so rows are decoded in
  // order and memoised, and the caller stops as soon as it has its answer.
  const done: Buffer[] = [];
  const row = (y: number): Buffer => {
    for (let n = done.length; n <= y; n++) {
      const filter = raw[n * (stride + 1)];
      const line = raw.subarray(n * (stride + 1) + 1, (n + 1) * (stride + 1));
      const cur = Buffer.alloc(stride);
      const up = done[n - 1] ?? Buffer.alloc(stride);
      for (let x = 0; x < stride; x++) {
        const a = x >= bpp ? cur[x - bpp] : 0;
        const b = up[x];
        const c = x >= bpp ? up[x - bpp] : 0;
        let v = line[x];
        if (filter === 1) v += a;
        else if (filter === 2) v += b;
        else if (filter === 3) v += (a + b) >> 1;
        else if (filter === 4) {
          const p = a + b - c;
          const pa = Math.abs(p - a);
          const pb = Math.abs(p - b);
          const pc = Math.abs(p - c);
          v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
        }
        cur[x] = v & 0xff;
      }
      done.push(cur);
    }
    return done[y];
  };
  return { width, height, row };
}

describe("the page both production manifests submit as their Support URL", () => {
  /**
   * The premise every other test here rests on. If the Support URL moves to a
   * hosted page, these checks are pointed at the wrong file and should say so
   * loudly rather than keep passing about a page nobody submits any more.
   */
  it.each(["manifest-prod.xml", "manifest-excel-prod.xml"])("%s still serves README.md as its Support URL", (name) => {
    const url = /<SupportUrl\s+DefaultValue="([^"]+)"/.exec(repoText(name))?.[1];
    expect(
      url,
      `${name}'s Support URL is ${url}, not the repository root. The checks in this file audit README.md ` +
        `because GitHub serves it at the repo root — re-point them at whatever page now answers the Support URL.`,
    ).toBe(SUPPORT_URL);
  });

  /**
   * THE HERO IMAGE MUST NOT CARRY A RENDERED PAGE HEADING.
   *
   * A heading baked into the screenshot is a product name no rename can reach:
   * this one said "PowerChart" for months after the rename, in the first
   * 89 rows of the image, above the chart grid. The rule that survives a future
   * re-capture is structural rather than textual — the image starts at the chart
   * grid, so there is no text band for a name to hide in. The title lives in the
   * README's alt text and caption instead, which the next test pins.
   *
   * "Ink" is any pixel that is not the page background, and the background is
   * taken from the image's own top-left corner (page padding) rather than a
   * hard-coded hex, so a re-styled page does not produce a false red.
   *
   * A card is told from a line of text by the LONGEST CONTIGUOUS run of ink in
   * the row, not by how much ink the row holds in total: the gallery's subtitle
   * ran the full width of the page and covered more than half the row's pixels,
   * so a share-of-row test called a sentence a chart. A card's top border is one
   * unbroken run of ~47% of the width (two columns with a gap); the longest run
   * inside a line of 15px text is a glyph stem.
   */
  it("starts at the chart grid, with no heading band above it", () => {
    const { width, height, row } = readRgbRows(repoFile("docs/gallery.png"));
    const bg = row(0).subarray(0, 3);
    const isBg = (r: Buffer, x: number) => r[x * 3] === bg[0] && r[x * 3 + 1] === bg[1] && r[x * 3 + 2] === bg[2];
    const SOLID = width * 0.2;

    const scan = (y: number) => {
      const r = row(y);
      let ink = 0;
      let run = 0;
      let longest = 0;
      for (let x = 0; x < width; x++) {
        if (isBg(r, x)) run = 0;
        else {
          ink++;
          longest = Math.max(longest, ++run);
        }
      }
      return { ink, longest };
    };

    let firstInk = -1;
    let firstSolid = -1;
    for (let y = 0; y < height && firstSolid < 0; y++) {
      const { ink, longest } = scan(y);
      if (ink > 0 && firstInk < 0) firstInk = y;
      if (longest >= SOLID) firstSolid = y;
    }

    expect(
      firstSolid,
      "no solid band found in docs/gallery.png — is it still the gallery screenshot?",
    ).toBeGreaterThanOrEqual(0);
    expect(
      firstInk,
      `docs/gallery.png has ink at row ${firstInk} but the chart grid only starts at row ${firstSolid}: ` +
        `${firstSolid - firstInk} rows of something else — a page heading — sit above it. A heading in the ` +
        `pixels is a product name no rename can grep for, which is exactly how "PowerChart" survived on the ` +
        `Support URL page. Crop to the grid and put the title in the README's alt text instead.`,
    ).toBe(firstSolid);
  });

  /**
   * ...and the title it replaced has to actually be there. Alt text is also the
   * only version of this image a screen reader gets.
   */
  it("names the product in the hero image's alt text", () => {
    const alt = /!\[([^\]]*)\]\(docs\/gallery\.png\)/.exec(repoText("README.md"))?.[1];
    expect(alt, "README.md no longer shows docs/gallery.png at all").toBeTruthy();
    expect(
      alt,
      `the hero image's alt text is "${alt}". It carries the title that used to be baked into the screenshot, ` +
        `so it has to name the product — that is the whole point of moving it out of the pixels.`,
    ).toMatch(/SSF Charts/);
  });

  /**
   * THE FEATURE TABLE IS A CAPABILITY LIST, NOT A COMPARISON.
   *
   * `docs/STORE-LISTING.md`: "the public listing, name, description, and
   * screenshots must **not** use the 'think-cell' mark as branding … internal
   * docs may keep it." A 168-row table headed with the competitor's name, on the
   * page submitted as the Support URL, is the mark used as framing.
   *
   * Scoped to table rows on purpose. The nominative mentions in the intro and
   * the Disclaimer are the owner's positioning call and are deliberately left
   * alone; the next test keeps them honest.
   */
  it("keeps the competitor's mark out of the feature table", () => {
    const offenders = repoText("README.md")
      .split("\n")
      .map((line, n) => [n + 1, line] as const)
      .filter(([, line]) => line.startsWith("|") && COMPETITOR.test(line));
    expect(
      offenders.map(([n, line]) => `${n}: ${line.slice(0, 80)}`),
      `README.md is the manifests' Support URL, so its feature table is store-facing copy. A row that names ` +
        `think-cell turns the table into a comparison against a competitor on the page Microsoft's reviewer ` +
        `opens. Describe the capability instead — docs/STORE-LISTING.md has the rule.`,
    ).toEqual([]);
  });

  /**
   * The flip side of allowing nominative mentions at all: they are only fair use
   * while the page says plainly that this is not that product. The Disclaimer
   * has to outlive any edit that keeps the mark.
   */
  it("keeps the disclaimer for as long as the mark appears anywhere on the page", () => {
    const readme = repoText("README.md");
    if (!COMPETITOR.test(readme)) return;
    expect(
      readme,
      `README.md still names think-cell but has lost the "not affiliated" disclaimer. The mentions are ` +
        `defensible as nominative use only while the page disclaims affiliation — drop one or the other, not ` +
        `the disclaimer alone.`,
    ).toMatch(/not affiliated with[\s\S]{0,120}think-cell/i);
  });
});
