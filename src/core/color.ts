/**
 * Color-scale math for heatmaps and tile maps: linear-light RGB
 * interpolation (fixes the muddy midpoint of naive sRGB lerp) and the
 * sequential / diverging value→color scales.
 */

/**
 * A paint as text, for anything at all that arrives claiming to be one.
 *
 * `(color ?? "").trim()` guarded null and undefined and nothing else, so a
 * palette of NUMBERS — `style.palette: [1, 2, 3]` — threw
 * `TypeError: (color ?? "").trim is not a function` straight out of the
 * renderer. Every route into this function is user JSON: the pane's style
 * import stores whatever parses, a chart config can be pasted in whole, and the
 * skill hands one over from an agent. The type says `string`; TypeScript checks
 * the code, not the file someone pastes.
 *
 * Anything that is not a string reads as the empty string, which the callers
 * already handle: mid grey from `toRgb`, opaque from `alphaOf`. That is the
 * same answer they give for a named CSS colour, so a bad paint degrades exactly
 * like an unrecognised one rather than taking the chart down.
 */
function paintText(color: unknown): string {
  return typeof color === "string" ? color.trim() : "";
}

/**
 * Parse any paint the renderer's allow-list accepts into RGB.
 *
 * A hex-only `parseInt(h, 16)` yields NaN for `rgb()`/`hsl()`, and the bitwise
 * ops then coerce that NaN to 0 — so every functional colour silently read as
 * PURE BLACK. Those forms are legitimate config (Series.color / Series.colors /
 * style.palette are plain strings, and svg.ts's PAINT_OK admits them), so the
 * mis-read reached contrastInk (white ink chosen for a near-white fill) and
 * lerpColor (tints lerping toward black instead of the real hue).
 *
 * Named CSS colours were a "known gap" here for as long as the table looked
 * like more weight than it was worth, and the gap cost more than the table:
 * every one of them read as MID GREY, which is not a sane fallback — it is a
 * confident wrong answer, and two helpers act on it.
 *
 *  - `contrastInk` chose WHITE ink for a `lightyellow` fill. Grey's luminance
 *    sits just under the threshold, so every light named colour got invisible
 *    labels.
 *  - `relLuminance(style.background)` decides whether the canvas is dark, so
 *    `background: "white"` put every canvas-aware helper into DARK mode: the
 *    same chart written `#ffffff` and `white` came out with different heatmap
 *    tints, zone fills and no-data greys.
 *
 * The table is the one the pptx sink has always carried. The two are separate
 * files on purpose (that one ships standalone inside the skill zip), and
 * `test/color.test.ts` now drives NAMED colours through both sinks, so they are
 * pinned to each other by a test rather than by hope.
 */
/**
 * What an unparseable colour becomes. Mid grey, which is what every branch
 * below already falls back to — named here so the guard and the fallbacks stay
 * the same answer.
 */
/**
 * The CSS Color 4 named colours as "name rrggbb" pairs — the same table the
 * headless pptx sink carries (`skill/scripts/pptx-paint.mjs`, which ships
 * standalone in the skill zip and so cannot import this one). `test/color.test.ts`
 * drives named colours through both, so a change to one fails against the other.
 */
const CSS_NAMES: Record<string, string> = Object.fromEntries(
  `aliceblue f0f8ff, antiquewhite faebd7, aqua 00ffff, aquamarine 7fffd4, azure f0ffff, beige f5f5dc,
  bisque ffe4c4, black 000000, blanchedalmond ffebcd, blue 0000ff, blueviolet 8a2be2, brown a52a2a,
  burlywood deb887, cadetblue 5f9ea0, chartreuse 7fff00, chocolate d2691e, coral ff7f50, cornflowerblue 6495ed,
  cornsilk fff8dc, crimson dc143c, cyan 00ffff, darkblue 00008b, darkcyan 008b8b, darkgoldenrod b8860b,
  darkgray a9a9a9, darkgreen 006400, darkgrey a9a9a9, darkkhaki bdb76b, darkmagenta 8b008b,
  darkolivegreen 556b2f, darkorange ff8c00, darkorchid 9932cc, darkred 8b0000, darksalmon e9967a,
  darkseagreen 8fbc8f, darkslateblue 483d8b, darkslategray 2f4f4f, darkslategrey 2f4f4f, darkturquoise 00ced1,
  darkviolet 9400d3, deeppink ff1493, deepskyblue 00bfff, dimgray 696969, dimgrey 696969, dodgerblue 1e90ff,
  firebrick b22222, floralwhite fffaf0, forestgreen 228b22, fuchsia ff00ff, gainsboro dcdcdc,
  ghostwhite f8f8ff, gold ffd700, goldenrod daa520, gray 808080, green 008000, greenyellow adff2f, grey 808080,
  honeydew f0fff0, hotpink ff69b4, indianred cd5c5c, indigo 4b0082, ivory fffff0, khaki f0e68c,
  lavender e6e6fa, lavenderblush fff0f5, lawngreen 7cfc00, lemonchiffon fffacd, lightblue add8e6,
  lightcoral f08080, lightcyan e0ffff, lightgoldenrodyellow fafad2, lightgray d3d3d3, lightgreen 90ee90,
  lightgrey d3d3d3, lightpink ffb6c1, lightsalmon ffa07a, lightseagreen 20b2aa, lightskyblue 87cefa,
  lightslategray 778899, lightslategrey 778899, lightsteelblue b0c4de, lightyellow ffffe0, lime 00ff00,
  limegreen 32cd32, linen faf0e6, magenta ff00ff, maroon 800000, mediumaquamarine 66cdaa, mediumblue 0000cd,
  mediumorchid ba55d3, mediumpurple 9370db, mediumseagreen 3cb371, mediumslateblue 7b68ee,
  mediumspringgreen 00fa9a, mediumturquoise 48d1cc, mediumvioletred c71585, midnightblue 191970,
  mintcream f5fffa, mistyrose ffe4e1, moccasin ffe4b5, navajowhite ffdead, navy 000080, oldlace fdf5e6,
  olive 808000, olivedrab 6b8e23, orange ffa500, orangered ff4500, orchid da70d6, palegoldenrod eee8aa,
  palegreen 98fb98, paleturquoise afeeee, palevioletred db7093, papayawhip ffefd5, peachpuff ffdab9,
  peru cd853f, pink ffc0cb, plum dda0dd, powderblue b0e0e6, purple 800080, rebeccapurple 663399, red ff0000,
  rosybrown bc8f8f, royalblue 4169e1, saddlebrown 8b4513, salmon fa8072, sandybrown f4a460, seagreen 2e8b57,
  seashell fff5ee, sienna a0522d, silver c0c0c0, skyblue 87ceeb, slateblue 6a5acd, slategray 708090,
  slategrey 708090, snow fffafa, springgreen 00ff7f, steelblue 4682b4, tan d2b48c, teal 008080, thistle d8bfd8,
  tomato ff6347, turquoise 40e0d0, violet ee82ee, wheat f5deb3, white ffffff, whitesmoke f5f5f5, yellow ffff00,
  yellowgreen 9acd32`
    .split(",")
    .map((pair) => pair.trim().split(" ") as [string, string]),
);

const UNREADABLE: [number, number, number] = [128, 128, 128];

/**
 * Is this a CSS colour NAME the table knows?
 *
 * Exported for the live Office renderer, which hands named colours to the host
 * verbatim — Office knows the same names — and needs to tell a name from any
 * other bare word. Own-property only, like every lookup on this table.
 */
export function isNamedColor(color: string): boolean {
  const c = paintText(color).toLowerCase();
  return Object.prototype.hasOwnProperty.call(CSS_NAMES, c);
}

export function toRgb(color: string): [number, number, number] {
  const c = paintText(color);
  // The `transparent` keyword is no colour at all: it is the slide behind it,
  // which every sink already takes to be white — `officeHex` answers "#ffffff"
  // for it, the pptx sink's `hexOr` falls back to white for it, and the SVG
  // canvas it draws is transparent over a white slide. `alphaOf` knew the
  // keyword and this did not, so it
  // read as the UNREADABLE mid grey below, whose luminance (0.216) is under the
  // 0.5 cut `relLuminance(style.background)` makes, and a chart written
  // `background: "transparent"` was laid out for a DARK canvas that is drawn
  // light: the heatmap's cell lightness range fell from 4.01:1 to 1.10:1, "no
  // data" came out #919191 — louder than the data — and every tint that mixes
  // toward the background (box, violin, treemap, sunburst, ribbons, waffle)
  // mixed toward grey. Ink on a fill that is not there goes against the slide
  // for the same reason.
  if (/^transparent$/i.test(c)) return [255, 255, 255];
  if (c.startsWith("#")) {
    const h = c.slice(1);
    // CSS hex is 3, 4, 6 or 8 digits and nothing else. Any other length is a
    // TYPO, and the length check is what stops it becoming a colour.
    //
    // Without it `parseInt` happily reads whatever is there and the result
    // looks perfectly valid: `#12345` — a six-digit colour with a digit dropped
    // — came back `#012345`, and `#1` came back `#000001`. So the one input a
    // user is most likely to produce by accident rendered as a plausible wrong
    // colour rather than as the grey that makes a mistake visible. The pptx
    // sink answers black for the same string (its exit check demands exactly
    // six digits), so a single config rendered two different ways depending on
    // which renderer you looked at — the three-sink divergence this file's
    // siblings keep being fixed for.
    //
    // 4 and 8 carry an alpha byte the colour maths has no use for; `alphaOf`
    // reads it separately.
    //
    // And every character has to be a hex digit, for the same reason as the
    // length. `parseInt` stops at the first character that is not one and
    // returns the digits in front of it, so NaN only ever came back when the
    // FIRST was bad: `#00ff0o` (a letter o for a zero) read as #000ff0, `#4e7g`
    // as #44ee77, and the alpha slot of a 4- or 8-digit form is never read at
    // all, so `#abc;` — a CSS paste with its semicolon — read as #aabbcc.
    // Those are the same plausible wrong colour the length check exists to
    // stop, and the sinks split on it: `officeHex` draws whatever this answers,
    // while the SVG allow-list and the pptx sink's `readHex` are both strict
    // and draw black. `contrastInk` then inked the segment for the invented
    // colour — dark on a light #aabbcc — over a preview and a deck that fill it
    // black, where an unreadable paint gets white ink and stays legible.
    if (!/^(?:[0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(h)) return UNREADABLE;
    const rgb = h.length === 3 || h.length === 4 ? h.slice(0, 3).replace(/./g, "$&$&") : h.slice(0, 6);
    const n = parseInt(rgb, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  // `parseFloat` is looser than the regex that fed it: `[\d.]+` matches a bare
  // "." and a "..", and `parseFloat(".")` is NaN. So a colour of `hsl(., 50%,
  // 50%)` — the sort of thing a hand-edited config or a template written in
  // another deck arrives with — used to reach the maths as NaN. In the hsl
  // branch that was a CRASH, not a wrong colour: `Math.floor(NaN / 60) % 6` is
  // NaN, the sector table has no NaN entry, and destructuring `undefined`
  // throws. In the rgb branch it was `#NaNNaNNaN`, which is not a colour and,
  // in the pptx sink, not the six hex digits that path's own security note
  // requires. Both are the same root: a `string` in the types is not a number
  // in the file someone pasted.
  const nums = (s: string) => (s.match(/-?[\d.]+%?/g) ?? []).map((v) => parseFloat(v));
  const finite = (vs: number[]) => vs.every((v) => Number.isFinite(v));
  if (/^rgba?\(/i.test(c)) {
    // Only the r/g/b components decide the 0–255 vs 0–100% scale. Testing the
    // whole string let a percentage ALPHA — the perfectly legal
    // `rgba(200,100,50,50%)` — multiply the CHANNELS by 2.55 and clip the
    // colour to near-white. The pptx sink (`skill/scripts/pptx-paint.mjs`) was
    // fixed for exactly this and carries a note saying so; this sink was not,
    // and it is the one `officeHex` calls — so the same config drew the right
    // colour in the headless deck and a washed-out one in the live add-in.
    const comps = (c.slice(c.indexOf("(") + 1).match(/-?[\d.]+%?/g) ?? []).slice(0, 3);
    const [r = 0, g = 0, b = 0] = comps.map((v) => parseFloat(v));
    if (!finite([r, g, b])) return UNREADABLE;
    const scale = comps.some((v) => v.endsWith("%")) ? 2.55 : 1; // percentages are 0–100, bare numbers 0–255
    return [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v * scale)))) as [number, number, number];
  }
  if (/^hsla?\(/i.test(c)) {
    // The hue is read as DEGREES, so a hue that says otherwise is not one this
    // function can read. `nums` keeps the digits and drops the letters after
    // them, which made `hsl(0.5turn 100% 50%)` — CSS cyan — half a degree, and
    // red; `3.14159rad` and `200grad` the same. It also split `1e2` into 1 and 2
    // and shifted every component along, and read `hsl(none 100% 50%)` with the
    // saturation as the hue. The unit is not ours to convert: the paint grammar
    // admits `deg` and nothing else (svg.ts's PAINT_OK), so the preview draws
    // these black, and learning `turn` here alone would leave the deck's parser
    // (`readHex` in `skill/scripts/pptx-paint.mjs`, which reads the number the
    // same way) and the preview as two other answers. Unreadable is the answer
    // the contract gives (`toRgb` parses what the allow-list accepts), and the
    // declared grey beats a confident wrong hue that `contrastInk` then inks for.
    //
    // `deg` stays: it was right all along, in this sink and the deck's, and
    // PAINT_OK admits it. A bare `hsl()` has no hue token and is left to the
    // maths below, as it always was.
    const hue = /^hsla?\(\s*([^\s,/)]*)/i.exec(c)?.[1] ?? "";
    if (hue !== "" && !/^[+-]?[\d.]+(?:deg)?$/i.test(hue)) return UNREADABLE;
    const [h = 0, s = 0, l = 0] = nums(c);
    if (!finite([h, s, l])) return UNREADABLE;
    return hslToRgb(((h % 360) + 360) % 360, Math.max(0, Math.min(1, s / 100)), Math.max(0, Math.min(1, l / 100)));
  }
  // A named CSS colour. Own-property only: a series coloured `__proto__` or
  // `constructor` would otherwise reach Object.prototype and come back an object
  // or a FUNCTION, which is how the same table broke the pptx sink and the SVG
  // pattern table before it.
  const named = Object.prototype.hasOwnProperty.call(CSS_NAMES, c.toLowerCase()) ? CSS_NAMES[c.toLowerCase()] : "";
  if (/^[0-9a-fA-F]{6}$/.test(named)) {
    const n = parseInt(named, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  // A BARE hex string, `#` omitted. The pptx sink has always taken these — it is
  // OOXML's own spelling, `<a:srgbClr val="AABBCC"/>` — and this sink and the
  // Office one did not, so `AABBCC` as a series colour drew mid grey in the
  // preview and in the live add-in and the real colour in the skill's deck. The
  // preview lying about the deck is the divergence class this file's siblings
  // keep being fixed for, and a user pasting a hex out of a brand guide without
  // its `#` is the commonest way to reach it.
  //
  // Last, so a NAMED colour always wins the string. Nothing is actually taken
  // from the name table by this — no CSS colour name is a valid bare hex string,
  // which is checked rather than assumed (`test/color.test.ts`) — but the order
  // is what makes that a property of the code and not of the table's contents.
  //
  // 3 and 6 digits only. The 4- and 8-digit alpha forms are deliberately left to
  // the `#` spelling: without the prefix, `aabbcc80` is as likely to be an id or
  // a truncated token as a colour, and the strictness above is there because the
  // input most worth catching is a TYPO.
  if (/^([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(c)) {
    const rgb = c.length === 3 ? c.replace(/./g, "$&$&") : c;
    const n = parseInt(rgb, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  return UNREADABLE;
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  const chroma = (1 - Math.abs(2 * l - 1)) * s;
  const x = chroma * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - chroma / 2;
  const [r, g, b] = [
    [chroma, x, 0],
    [x, chroma, 0],
    [0, chroma, x],
    [0, x, chroma],
    [x, 0, chroma],
    [chroma, 0, x],
  ][Math.floor(h / 60) % 6];
  return [r, g, b].map((v) => Math.round((v + m) * 255)) as [number, number, number];
}

const hexToRgb = toRgb;

/**
 * Any allow-listed paint → a plain 6-digit `#RRGGBB` (no alpha). The PowerPoint
 * renderers' colour sinks accept only `#RRGGBB` (or a named colour), so rgb()/
 * hsl()/3-digit/8-digit forms — all valid config — must be normalised here or
 * they mis-render (Office.js) or fall back to black (pptx). Named colours are a
 * known gap and resolve to mid grey via toRgb; the office renderer passes named
 * colours through natively instead of calling this.
 */
export function toHex6(color: string): string {
  return rgbToHex(toRgb(color));
}

/**
 * Opacity in [0, 1] carried by a paint (8-digit `#RRGGBBAA`, `rgba()`, `hsla()`);
 * 1 for every opaque form. The PowerPoint renderers split this into a shape's
 * `transparency` so an alpha authored in the config isn't silently dropped.
 */
export function alphaOf(color: string): number {
  const c = paintText(color);
  // The `transparent` KEYWORD, which the sibling in `skill/scripts/pptx-paint.mjs`
  // has always known and this one did not. It is the documented floating-segment
  // idiom, and only two layouts guarded it before it reached a renderer — mekko's
  // guard even says why: "Office.js hands 'transparent' to setSolidColor, which
  // it rejects". The guard was never swept to the SINK, so on every other kind
  // the bare word went to the live host.
  if (/^transparent$/i.test(c)) return 0;
  // 4- and 8-digit hex carry the alpha in their last digit / last byte.
  const hex = /^#([0-9a-fA-F]{4}|[0-9a-fA-F]{8})$/.exec(c);
  if (hex) {
    const h = hex[1];
    const aa = h.length === 4 ? h[3] + h[3] : h.slice(6, 8);
    return parseInt(aa, 16) / 255;
  }
  // `rgb`/`hsl` as well as `rgba`/`hsla`: CSS Color 4 made them aliases, and the
  // channel readers in all three sinks already accept the short spelling, so
  // rejecting it here dropped the alpha of a colour they were otherwise reading
  // correctly.
  const fn = /^(?:rgba?|hsla?)\(([^)]*)\)$/i.exec(c);
  if (fn) {
    const a = alphaArg(fn[1]);
    if (a !== undefined) return Number.isFinite(a) ? Math.max(0, Math.min(1, a)) : 1;
  }
  return 1;
}

/**
 * The alpha argument of an `rgb()`/`hsl()` body, in either CSS syntax.
 *
 * The old rule split on `[,/]` and took the FOURTH part, which reads the legacy
 * comma form and quietly fails the modern one: `rgb(70 130 180 / 0.5)` splits
 * into two parts, not four, so the alpha was dropped and the colour drawn
 * opaque. Both PowerPoint renderers already parsed that form's CHANNELS
 * correctly, so a translucent steel blue arrived in the deck as a solid one —
 * with the preview drawing it BLACK, because the SVG allow-list did not admit
 * the slash at all. One colour, three different results.
 *
 * The slash decides: CSS Color 4 puts the alpha after it and separates channels
 * with spaces, while the legacy form has no slash and four commas.
 */
function alphaArg(body: string): number | undefined {
  const slash = body.indexOf("/");
  const raw = slash >= 0 ? body.slice(slash + 1) : body.split(",")[3];
  if (raw === undefined) return undefined;
  const t = raw.trim();
  if (t === "") return undefined;
  return t.endsWith("%") ? parseFloat(t) / 100 : parseFloat(t);
}

/**
 * Three channels to `#RRGGBB`, always — six hex digits, whatever it is handed.
 *
 * The clamp used to be `Math.max(0, Math.min(255, Math.round(c)))`, which is
 * total for numbers and not for NaN: every one of those returns NaN, and
 * `NaN.toString(16)` is the STRING "NaN", so a single unreadable channel
 * produced `#NaNNaNNaN`. That is not a colour in SVG, and in the pptx sink it
 * is not the six hex digits that path's security note requires.
 *
 * Guarded here rather than at the four call sites, because every producer of a
 * channel — `toRgb`, `lerpColor`'s interpolation, both colour scales — can
 * arrive with a NaN for its own reasons, and a fix at the root cannot be
 * forgotten by the next one.
 */
const rgbToHex = (rgb: number[]): string =>
  "#" +
  rgb
    .map((c) =>
      Math.max(0, Math.min(255, Number.isFinite(c) ? Math.round(c) : 0))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("");

const toLin = (c: number) => (c / 255 <= 0.04045 ? c / 255 / 12.92 : Math.pow((c / 255 + 0.055) / 1.055, 2.4));
const toSrgb = (c: number) => 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055);

/** Interpolate two hex colors in linear-light RGB (t ∈ [0, 1]). */
export function lerpColor(c0: string, c1: string, t: number): string {
  const a = hexToRgb(c0);
  const b = hexToRgb(c1);
  return rgbToHex(
    [0, 1, 2].map((i) => toSrgb(toLin(a[i]) + (toLin(b[i]) - toLin(a[i])) * Math.max(0, Math.min(1, t)))),
  );
}

/** Fill for cells/tiles with no data, on a light canvas. */
export const NO_DATA = "#e6e6e6";

/** Perceptual (linear-light) relative luminance of a hex colour, 0..1. */
const relLuminance = (hex: string): number => {
  const [r, g, b] = hexToRgb(hex).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};

/**
 * The colour a chart is laid out AGAINST, from the `style.background` it was
 * given: that colour when it is one, the default white canvas when it is not.
 *
 * "Not" is a paint that draws nothing (`transparent`, or any form with zero
 * alpha — `rgba(0, 0, 0, 0)` is what a browser's computed style answers for it)
 * or one that cannot be read at all (`none`, a typo like `off-white`). The sinks
 * already treat both the same way: the SVG canvas is transparent over a white
 * slide, and `paintOr` (the preview) and `hexOr` (the deck) fall back to white
 * for a paint they cannot read. Here the mid grey `toRgb` answers for an
 * unreadable paint — chosen so that `contrastInk` still picks a sane INK — was
 * read as a canvas verdict, and a luminance of 0.216 laid the chart out for a
 * dark slide.
 *
 * Applied in the helpers below, which take `background` as a plain string, and
 * once for every layout by `layoutBackground`, which `resolveStyle` runs where
 * the style is built — the direct `lerpColor(style.background, …)` sites need no
 * edit of their own.
 */
const canvasOf = (background: string): string => {
  const laid = layoutBackground(background);
  return laid === "transparent" ? "#ffffff" : laid;
};

/**
 * `style.background` as the LAYOUTS read it: one answer for every spelling of "no
 * canvas", resolved once where the style is built (`resolveStyle`).
 *
 * WHY ONCE. A layout reads `style.background` two ways. Colour maths mixes toward
 * it (`lerpColor(style.background, fill, 0.22)` in the boxplot, violin, sunburst,
 * treemap, scatter, line ribbons, column marks), and a paint IS it (the hairline
 * between segments, the hollow of a forecast marker). `toRgb` has to answer
 * something for a paint it cannot read, and the answer is the mid grey that
 * `contrastInk` needs for INK — so a background it could not read, or one with no
 * alpha, laid every one of those mixes out against grey and then dark. 7ab702a
 * taught `toRgb` the `transparent` KEYWORD and gave the helpers in this file
 * `canvasOf`; the other spellings (`rgba(0,0,0,0)`, `#0000`, `#00000000`,
 * `hsla(…, 0)`, `none`, a typo, an empty string) were still grey at every direct
 * site, and a sweep of the sample kinds drew 14 paints differently from
 * `transparent` (boxplot 4, sunburst 7, violin 3). Editing the two dozen call sites
 * in a dozen layouts was the wrong shape for a fix that is one fact about one string.
 *
 * WHAT IT ANSWERS, so the maths reads one canvas and the paints change as little as
 * they can:
 *  - a colour that draws something → itself, byte for byte;
 *  - a paint that draws nothing (any form with zero alpha, and `none`, the SVG
 *    word for it) → the `transparent` keyword. `toRgb` reads that as the white
 *    slide behind it and `alphaOf` as 0, so the maths sees a light canvas and a
 *    stroke or fill painted in it stays invisible in every sink, exactly as
 *    `rgba(0,0,0,0)` was. White-with-no-alpha would do the same, but the keyword
 *    is already what the sinks and the layouts guard on;
 *  - a paint no sink can read (`off-white`, `""`, junk, a non-string) → `#ffffff`,
 *    which is what the canvas itself is drawn as for it (`paintOr`'s and
 *    `hexOr`'s fallback). A separator in the canvas's colour is the default
 *    chart's look; the sinks' own fallback for a stroke is black.
 *
 * The CANVAS PAINT is not touched. The pane and the deck read `cfg.style.background`
 * from the config (`canvasColor` in the pane, `backgroundOf` in the deck), not
 * from the style a layout was handed, so a transparent chart stays transparent in
 * the SVG and on the slide.
 */
export function layoutBackground(background: unknown): string {
  const c = paintText(background);
  if (/^none$/i.test(c)) return "transparent";
  if (toRgb(c) === UNREADABLE) return "#ffffff";
  return alphaOf(c) === 0 ? "transparent" : (background as string);
}

/**
 * A background panel/zone tint (weekend shading, marginal-total strips, scatter
 * quadrants) that stays subtle on either theme. The charts hardcode tints tuned
 * for a light canvas; on a LIGHT background this returns that literal UNCHANGED
 * (so default charts are byte-identical), but on a DARK background a light-grey
 * box would glare, so mirror the tint's small step away from white into an equal
 * step away from black — a faint lift off the dark canvas instead.
 */
export function zoneFill(background: string, lightFill: string): string {
  const canvas = canvasOf(background);
  if (relLuminance(canvas) >= 0.5) return lightFill;
  const drop = Math.min(1, Math.max(0, 1 - relLuminance(lightFill)));
  return lerpColor(canvas, "#ffffff", drop);
}

/**
 * "No data" fill for the canvas in play. A missing cell must read as ABSENT —
 * quieter than every value on the scale — and the light-canvas grey is instead
 * the brightest thing on a dark slide, which makes the gaps the loudest marks
 * in the grid. Light canvases keep the literal, so default charts are
 * byte-identical.
 *
 * Not zoneFill: that mirrors the tint's LINEAR-LIGHT step away from white into
 * an equal step away from black, which is accurate for the faint panel tints it
 * serves (#f5f5f5 and friends) but overshoots badly this far down — #e6e6e6
 * came back as #808080, a 4.3:1 block against a #1b1b1b slide where the light
 * pair is 1.25:1. Match the CONTRAST RATIO instead, so "absent" is equally
 * recessive on either canvas.
 */
export function noDataFill(background: string): string {
  const bg = relLuminance(canvasOf(background));
  if (bg >= 0.5) return NO_DATA;
  // The ratio NO_DATA keeps against a white canvas, re-solved against this one.
  const ratio = 1.05 / (relLuminance(NO_DATA) + 0.05);
  const target = Math.max(0, Math.min(1, ratio * (bg + 0.05) - 0.05));
  const c = Math.round(255 * (target <= 0.0031308 ? target * 12.92 : 1.055 * Math.pow(target, 1 / 2.4) - 0.055));
  return rgbToHex([c, c, c]);
}

/**
 * Sequential scale: 12% tint of the CANVAS (kept off the bare canvas so a low
 * value never reads as "no data") → the full color.
 *
 * The empty end is the background, not a hardcoded white: on a dark slide a
 * near-white low cell is the BRIGHTEST mark on the chart, so the scale reads
 * inverted — the smallest value shouts and the largest recedes. Defaulting to
 * white keeps every light-theme chart byte-identical.
 */
export function sequentialScale(
  min: number,
  max: number,
  color: string,
  background = "#ffffff",
): (v: number) => string {
  const lo = lerpColor(canvasOf(background), color, 0.12);
  const span = max - min || 1;
  return (v) => lerpColor(lo, color, (v - min) / span);
}

/**
 * Diverging scale through the canvas colour, symmetric around zero so equal
 * distances from zero get equal intensity on both sides. Zero must vanish into
 * the slide — a hardcoded white zero cell on a dark canvas is a glaring block
 * exactly where the data says "neutral".
 */
export function divergingScale(
  min: number,
  max: number,
  positive: string,
  negative: string,
  background = "#ffffff",
): (v: number) => string {
  const extent = Math.max(Math.abs(min), Math.abs(max)) || 1;
  const canvas = canvasOf(background);
  return (v) => (v >= 0 ? lerpColor(canvas, positive, v / extent) : lerpColor(canvas, negative, -v / extent));
}
