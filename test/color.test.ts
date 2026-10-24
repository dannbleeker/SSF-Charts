import { describe, expect, it } from "vitest";
import { contrastInk } from "../src/core/scene";
import { toRgb, toHex6, alphaOf, lerpColor } from "../src/core/color";
import { hex as pptxHex, alphaOf as pptxAlphaOf, CSS_NAMES } from "../skill/scripts/pptx-paint.mjs";

/** Colour parsing and normalisation — every paint form the config admits. */

describe("contrastInk", () => {
  it("expands 3-digit hex and picks readable ink", () => {
    expect(contrastInk("#fff")).toBe(contrastInk("#ffffff"));
    expect(contrastInk("#000")).not.toBe(contrastInk("#fff"));
  });

  it("reads rgb()/hsl() fills, not as black", () => {
    // A hex-only parser returned NaN->0 (pure black) for functional colours, so
    // a near-white rgb() fill wrongly got WHITE ink. These forms are valid config
    // (Series.color etc. are plain strings the renderer's PAINT_OK admits).
    expect(contrastInk("rgb(250,250,250)")).toBe(contrastInk("#fafafa"));
    expect(contrastInk("rgb(250,250,250)")).toBe("#0b0b0b"); // dark ink on near-white
    expect(contrastInk("rgb(20,20,20)")).toBe("#ffffff"); // white ink on near-black
    expect(contrastInk("hsl(0,0%,100%)")).toBe("#0b0b0b"); // hsl white
  });
});

describe("toHex6 / alphaOf normalize colours for the PowerPoint renderers", () => {
  it("normalizes every allow-listed form to a 6-digit hex", () => {
    expect(toHex6("#4e79a7")).toBe("#4e79a7"); // identity for the hex the engine emits
    expect(toHex6("#abc")).toBe("#aabbcc");
    expect(toHex6("#4e79a780")).toBe("#4e79a7"); // alpha byte dropped
    expect(toHex6("rgb(78,121,167)")).toBe("#4e79a7");
    expect(toHex6("hsl(0,0%,100%)")).toBe("#ffffff");
    expect(toHex6("not-a-color")).toBe("#808080"); // named/unknown → grey, never black
  });

  it("reads the opacity a paint carries, 1 when opaque", () => {
    expect(alphaOf("#4e79a7")).toBe(1);
    expect(alphaOf("#4e79a780")).toBeCloseTo(128 / 255, 5);
    expect(alphaOf("rgba(1,2,3,0.5)")).toBe(0.5);
    expect(alphaOf("hsla(0,0%,0%,0.25)")).toBe(0.25);
    expect(alphaOf("rgb(1,2,3)")).toBe(1);
  });
});

describe("toRgb parses every allow-listed colour form", () => {
  it("matches hex for the equivalent rgb()/hsl(), strips alpha, expands short hex", () => {
    expect(toRgb("rgb(78,121,167)")).toEqual(toRgb("#4e79a7"));
    expect(toRgb("#abc")).toEqual(toRgb("#aabbcc"));
    expect(toRgb("#4e79a780")).toEqual(toRgb("#4e79a7")); // alpha byte dropped
    expect(toRgb("rgb(50%,50%,50%)")).toEqual([127, 127, 127]);
    expect(toRgb("hsl(0,0%,0%)")).toEqual([0, 0, 0]);
    // A malformed paint falls back to mid grey, never NaN/black.
    expect(toRgb("not-a-color")).toEqual([128, 128, 128]);
  });
});

/**
 * A hex of the wrong LENGTH is a typo, not a colour.
 *
 * CSS hex is 3, 4, 6 or 8 digits. `toRgb` accepted any length and let
 * `parseInt` read whatever was there, so the result looked perfectly valid:
 * `#12345` — a six-digit colour with a digit dropped, which is the mistake a
 * person actually makes — came back `#012345`, and `#1` came back `#000001`.
 *
 * That is worse than a visible fallback twice over. The user sees a plausible
 * wrong colour instead of the grey that would show them the typo, and the pptx
 * sink answers BLACK for the same string (its exit check demands exactly six
 * digits), so one config rendered two different ways depending on which
 * renderer you looked at.
 */
describe("a hex colour of an impossible length", () => {
  /** What `toRgb` answers for a paint it cannot read — mid grey, deliberately visible. */
  const UNREADABLE = "#808080";

  it("falls back instead of inventing a colour from the digits that are there", () => {
    // The exact shapes a dropped or doubled keystroke produces.
    expect(toHex6("#12345"), "a five-digit hex was read as a colour").toBe(UNREADABLE);
    expect(toHex6("#1"), "a one-digit hex was read as a colour").toBe(UNREADABLE);
    expect(toHex6("#12"), "a two-digit hex was read as a colour").toBe(UNREADABLE);
    expect(toHex6("#1234567"), "a seven-digit hex was read as a colour").toBe(UNREADABLE);
    expect(toHex6("#123456789"), "a nine-digit hex was read as a colour").toBe(UNREADABLE);
  });

  it("still parses every length CSS actually defines", () => {
    // The other half, or the rule above could be satisfied by rejecting
    // everything. 4 and 8 drop their alpha byte here; `alphaOf` reads it.
    expect(toHex6("#123")).toBe("#112233");
    expect(toHex6("#1234")).toBe("#112233");
    expect(toHex6("#123456")).toBe("#123456");
    expect(toHex6("#12345678")).toBe("#123456");
  });
});

describe("a colour string that is not a number where a number was expected", () => {
  /**
   * `parseFloat` is looser than the regex that feeds it. `[\d.]+` matches a
   * bare "." and a "..", and `parseFloat(".")` is NaN — so `hsl(., 50%, 50%)`
   * reached the maths as NaN.
   *
   * In the hsl branch that was a CRASH rather than a wrong colour:
   * `Math.floor(NaN / 60) % 6` is NaN, the hue sector table has no NaN entry,
   * and destructuring `undefined` throws. In the rgb branch it produced
   * `#NaNNaNNaN`, which is not a colour anywhere.
   *
   * Ordinary input, too: a config arrives from the JSON box, a saved template,
   * a shape tag written in another deck, and the skill's caller — the same
   * boundary that made `categories: [2023, 2024]` crash.
   */
  const BROKEN = [
    "hsl(., 50%, 50%)",
    "hsl(-.., 50%, 50%)",
    "hsla(., ., .)",
    "rgb(., ., .)",
    "rgba(-.., 1, 1)",
    "hsl(50, ., .)",
  ];

  it("never throws, whatever it is handed", () => {
    for (const c of BROKEN) expect(() => toRgb(c), `toRgb threw on ${c}`).not.toThrow();
  });

  it("falls back to mid grey rather than to NaN", () => {
    for (const c of BROKEN) {
      expect(toRgb(c), `${c} produced a non-colour`).toEqual([128, 128, 128]);
      expect(toHex6(c), `${c} produced a non-colour`).toBe("#808080");
    }
  });

  it("always produces six hex digits, from any channel value at all", () => {
    // Guarded in `rgbToHex` rather than at its four call sites: every producer
    // of a channel can arrive with a NaN for its own reasons, and `lerpColor`
    // does — an interpolation weight of NaN clamps to NaN, not to 0 or 1.
    expect(toHex6("rgb(., ., .)")).toMatch(/^#[0-9a-f]{6}$/);
    expect(lerpColor("#000000", "#ffffff", NaN)).toMatch(/^#[0-9a-f]{6}$/);
    expect(lerpColor("#000000", "#ffffff", Infinity)).toMatch(/^#[0-9a-f]{6}$/);
  });

  it("still reads the colours that were always fine", () => {
    // The negative control. A parser that answered grey to everything would
    // pass all three cases above and silently grey out every chart.
    expect(toRgb("hsl(-30, 50%, 50%)")).toEqual([191, 64, 128]);
    expect(toRgb("rgb(10, 20, 30)")).toEqual([10, 20, 30]);
    expect(toHex6("#abc")).toBe("#aabbcc");
  });
});

describe("a percentage ALPHA is not a percentage channel", () => {
  /**
   * `rgba(200,100,50,50%)` is legal CSS and legal config. The scale test was
   * `/%/.test(<whole string>)`, so the alpha's percent sign was read as
   * evidence that the CHANNELS were percentages: every one got multiplied by
   * 2.55 and clipped, turning a mid orange into near-white.
   *
   * The pptx sink was fixed for this and says so in a comment; this sink was
   * not — and this is the one `officeHex` calls, so the live add-in drew the
   * washed-out colour while the skill's headless deck drew the right one.
   */
  it("reads the channels the same with a % alpha as with a decimal one", () => {
    expect(toRgb("rgba(200,100,50,50%)")).toEqual(toRgb("rgba(200,100,50,0.5)"));
    expect(toHex6("rgba(200, 100, 50, 50%)")).toBe("#c86432");
    expect(toRgb("hsla(200, 50%, 40%, 50%)")).toEqual(toRgb("hsla(200, 50%, 40%, 0.5)"));
  });

  it("still reads the alpha, and still scales percentage CHANNELS", () => {
    // The negative control: dropping the % test altogether would pass the case
    // above and break every percentage-channel colour.
    expect(alphaOf("rgba(200,100,50,50%)")).toBe(0.5);
    expect(toRgb("rgba(100%, 0%, 0%, 50%)")).toEqual([255, 0, 0]);
    expect(toRgb("rgb(50%,50%,50%)")).toEqual([127, 127, 127]);
  });
});

describe("the preview sink and the pptx sink answer the same colour", () => {
  /**
   * CLAUDE.md: there are three colour sinks and they are separate code on
   * purpose — `src/core/color.ts` (preview, and `officeHex` in the live
   * renderer via `toHex6`), `skill/scripts/pptx-paint.mjs` (headless pptx) and
   * the Office pass-through. The same defect has now been found in all three
   * independently, and the percentage-alpha bug above lived in ONE of them for
   * as long as it did because nothing ever compared them.
   *
   * Named colours are the one declared divergence: the pptx sink carries the
   * CSS table, this one answers mid grey (documented on `toRgb`), and the live
   * renderer hands the name to Office untouched. Everything else must agree.
   */
  const FORMS = [
    "#4e79a7",
    "#ABC",
    "#4e79a780",
    "#abcd",
    "  #4e79a7  ",
    "rgb(78,121,167)",
    "rgb(78, 121, 167)",
    "rgba(78,121,167,0.5)",
    "rgba(200,100,50,50%)",
    "rgb(50%,50%,50%)",
    "rgba(100%,0%,0%,50%)",
    "hsl(0,0%,100%)",
    "hsl(210,60%,40%)",
    "hsla(210,60%,40%,0.25)",
    "hsla(210,60%,40%,25%)",
    "hsl(-30, 50%, 50%)",
    "hsl(400, 50%, 50%)",
    "rgb(300,-20,50)",
    "rgb(", // a truncated function is not a function to either sink: black both sides
    // CSS Color 4's space-separated channels with a SLASH alpha — the form MDN
    // documents first, and the one an agent writing a config is most likely to
    // produce. Both sinks parsed the CHANNELS of these correctly and dropped
    // the alpha: the split took the fourth comma-part, which does not exist
    // here, so a translucent colour reached the deck solid. `rgb`/`hsl` without
    // the trailing `a` is the same syntax by CSS Color 4 and was rejected
    // outright by the alpha readers while both channel readers accepted it.
    "rgb(78 121 167)",
    "rgb(78 121 167 / 0.5)",
    "rgba(78 121 167 / 0.5)",
    "rgba(78 121 167 / 50%)",
    "hsl(210 60% 40%)",
    "hsl(210 60% 40% / 0.25)",
    "hsla(210 60% 40% / 25%)",
    // A BARE hex string, `#` omitted — OOXML's own spelling. The pptx sink has
    // always taken these and this one did not, so `AABBCC` as a series colour
    // drew the real colour in the skill's deck and mid grey in the preview and
    // in the live add-in: two of the three sinks disagreeing with the third
    // about a perfectly ordinary paint, which is what this whole block exists
    // to catch. Pasting a hex out of a brand guide without its `#` is the
    // commonest way to produce one.
    "AABBCC",
    "aabbcc",
    "4e79a7",
    "abc",
    "ABC",
    "  4E79A7  ",
    // Named colours. This sink read every one of them as mid grey until the
    // table it shares with the pptx sink was carried here too — a gap that put
    // white ink on light fills and flipped `background: "white"` into
    // dark-canvas mode. Listed here so the two tables are pinned to each other.
    "white",
    "black",
    "rebeccapurple",
    "lightyellow",
    "  LightYellow  ",
    "STEELBLUE",
    "darkgrey",
  ];

  /** Paints neither sink can read — the one place they answer differently. */
  const UNREADABLE_FORMS = [
    "rgb(., ., .)",
    "hsl(., 50%, 50%)",
    "",
    "not-a-color",
    // A hex of a LEGAL length with a stray character in it. `parseInt` read the
    // digits in front of the typo, so this sink drew an invented colour
    // (`#aabbcc`, `#000ff0`, `#4e79a7`) while the pptx sink's strict regex and
    // the SVG allow-list drew black: the declared grey/black pair is the only
    // difference there is meant to be.
    "#abc;",
    "#00ff0o",
    "#4e79a7zz",
  ];

  it("normalises every functional and hex form identically", () => {
    // Case only: the pptx sink echoes a hex literal's own case, and OOXML does
    // not care. Everything else about the six digits must match.
    for (const c of FORMS) {
      expect(toHex6(c).slice(1), `toHex6 vs pptx hex for ${JSON.stringify(c)}`).toBe(pptxHex(c).toLowerCase());
    }
  });

  it("reads the same opacity from every form", () => {
    for (const c of [...FORMS, ...UNREADABLE_FORMS, "transparent", "TRANSPARENT"]) {
      expect(alphaOf(c), `alphaOf vs pptx alphaOf for ${JSON.stringify(c)}`).toBeCloseTo(pptxAlphaOf(c), 6);
    }
  });

  it("reads the opacity the author actually wrote, in either CSS syntax", () => {
    // Agreeing is not enough — both sinks agreed on 1 for every slash form,
    // which is what made the bug invisible to the check above. This pins the
    // VALUE, so "they agree" cannot be satisfied by both being wrong together.
    for (const [c, want] of [
      ["rgb(78 121 167 / 0.5)", 0.5],
      ["rgba(78 121 167 / 0.5)", 0.5],
      ["rgba(78 121 167 / 50%)", 0.5],
      ["hsl(210 60% 40% / 0.25)", 0.25],
      ["hsla(210 60% 40% / 25%)", 0.25],
      ["rgba(78,121,167,0.5)", 0.5],
      // …and a form with no alpha stays opaque, so the rule is not "find any
      // number and call it the alpha".
      ["rgb(78 121 167)", 1],
      ["hsl(210 60% 40%)", 1],
      ["rgb(78,121,167)", 1],
    ] as [string, number][]) {
      expect(alphaOf(c), `core alphaOf(${JSON.stringify(c)})`).toBeCloseTo(want, 6);
      expect(pptxAlphaOf(c), `pptx alphaOf(${JSON.stringify(c)})`).toBeCloseTo(want, 6);
    }
  });

  it("cannot reinterpret a named colour as bare hex, whatever the table holds", () => {
    // `toRgb` tries the name table BEFORE bare hex, so a collision would be
    // resolved in the name's favour either way. This pins the stronger fact the
    // comment there claims: there is no collision to resolve, so the ordering is
    // a property of the code rather than of what happens to be in the table —
    // and if somebody adds `"beef"` to CSS_NAMES one day, this says so.
    const hexy = Object.keys(CSS_NAMES).filter((n) => /^([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(n));
    expect(hexy, "a CSS colour name is also a valid bare hex string").toEqual([]);
  });

  it("differs on an UNREADABLE paint, and only there — grey here, black in the deck", () => {
    // Declared, not accidental: this sink answers mid grey on purpose (so
    // `contrastInk` still picks a sane ink), the pptx sink answers black on
    // purpose (its security note requires six hex digits and black is the
    // documented fallback). Pinned so the difference stays a decision.
    for (const c of UNREADABLE_FORMS) {
      expect(toHex6(c), `${JSON.stringify(c)}`).toBe("#808080");
      expect(pptxHex(c), `${JSON.stringify(c)}`).toBe("000000");
    }
  });
});
