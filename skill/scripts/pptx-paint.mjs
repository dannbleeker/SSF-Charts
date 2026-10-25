/**
 * Pure paint + node-mapping helpers for the headless pptx renderer, split out of
 * render-pptx.mjs so they are importable and unit-testable IN-PROCESS. The CLI in
 * render-pptx.mjs runs as a subprocess (v8 can't measure it) and loads the engine
 * via top-level await; everything here is a pure function of its inputs — no
 * engine load, no argv, no file or pptx side effects — so it can be exercised and
 * measured directly. render-pptx.mjs imports these; build-skill.mjs ships this
 * file alongside it, so the relative import holds in both the repo and the zip.
 */

export const IN = 1 / 72; // points → inches

const to2 = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0");

export const hslToHex = (h, s, l) => {
  h = ((h % 360) + 360) % 360;
  s = Math.max(0, Math.min(1, s / 100));
  l = Math.max(0, Math.min(1, l / 100));
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
  return to2((r + m) * 255) + to2((g + m) * 255) + to2((b + m) * 255);
};

// The CSS Color 4 named colours as "name rrggbb" pairs. SVG renders a bare name
// natively and Office.js passes it through to the host, but pptxgenjs takes hex
// only — without this table every named colour collapsed to one grey, so two
// series coloured "red" and "steelblue" became indistinguishable in the deck.
export const CSS_NAMES = Object.fromEntries(
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
    .map((pair) => pair.trim().split(" ")),
);

// A paint as text, whatever actually turned up.
//
// `(c ?? "").trim()` guarded null and undefined and nothing else, so a palette
// of NUMBERS threw `TypeError: (c ?? "").trim is not a function` — out of
// `hex`, which the comment below correctly says must never fail to return six
// hex digits. A config reaches this renderer straight from an agent, where a
// numeric colour is a plausible mistake rather than an exotic one, and the
// throw lands outside every per-chart guard: one bad colour, no deck.
//
// Anything that is not a string reads as empty, which both callers already
// handle — black from `hex`, opaque from `alphaOf` — so a bad paint degrades
// like an unrecognised one. Same fix as `src/core/color.ts`; the two sinks are
// deliberately separate code and both had it.
const paintText = (c) => (typeof c === "string" ? c.trim() : "");

// Normalise ANY paint the SVG allow-list accepts to a validated 6-digit hex, so
// the headless pptx matches the preview instead of falling back to black for
// rgb()/hsl()/named colours. SECURITY: the colour is interpolated into OOXML
// unescaped, so this MUST always return exactly six hex digits — a value like
// `000"/><a:x` could otherwise inject markup. rgb()/hsl() are parsed, named
// colours resolve through CSS_NAMES; anything unrecognised is black.
export const hex = (c) => {
  // The guarantee, enforced where it is stated rather than trusted to every
  // branch below.
  //
  // Two of those branches broke it. `hsl(., 50%, 50%)` THREW — the regex that
  // finds the numbers matches a bare ".", `parseFloat(".")` is NaN, and the
  // hue sector table has no NaN entry — and `rgb(., ., .)` returned the
  // nine-character string "NaNNaNNaN". Neither is injection, but the six-digit
  // rule is what makes this safe, and a rule that holds "except for two inputs"
  // is not a rule. A throw is worse still: it happens inside pptxgenjs at
  // writeFile, outside every per-chart guard, so one bad colour destroys the
  // whole batch — which is exactly what the CSS_NAMES note below records
  // happening once already.
  let out;
  try {
    out = readHex(c);
  } catch {
    out = undefined;
  }
  return typeof out === "string" && /^[0-9a-fA-F]{6}$/.test(out) ? out : "000000";
};

/**
 * Like `hex`, but says "I did not recognise that" instead of guessing black.
 *
 * `hex`'s black fallback is right for INK — a label in an unknown colour is
 * still a label — and catastrophic for a slide BACKGROUND, which is the one
 * paint that decides whether everything else can be seen. `background:
 * "off-white"` (a typo) or `"transparent"` (a paint SSF Charts documents
 * elsewhere) produced `<a:srgbClr val="000000"/>` with the chart's own near-
 * black ink on top: a deck of black slides carrying invisible charts, opening
 * cleanly, validating, and reported as a success by the CLI.
 *
 * The SVG renderer already falls back to white here, so the two sinks disagreed
 * about the same config — the divergence class this project has now found in
 * all three colour sinks independently.
 *
 * Kept separate from `hex` rather than changing it: the six-hex-digit guarantee
 * `hex` makes is load-bearing for every other call site, and a paint that is
 * merely unrecognised is a different fact from one that is unreadable.
 */
export const hexOr = (c, fallbackHex) => {
  let out;
  try {
    out = readHex(c);
  } catch {
    out = undefined;
  }
  return typeof out === "string" && /^[0-9a-fA-F]{6}$/.test(out) ? out : hex(fallbackHex);
};

const readHex = (c) => {
  const raw = paintText(c);
  const h = raw.replace("#", "");
  if (/^[0-9a-fA-F]{3,4}$/.test(h))
    return h
      .slice(0, 3)
      .replace(/./g, "$&$&"); // #abc / #abcd → aabbcc
  if (/^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(h)) return h.slice(0, 6);
  const nums = (str) => (str.match(/-?[\d.]+%?/g) ?? []).map((v) => parseFloat(v));
  let m;
  if ((m = /^rgba?\(([^)]*)\)$/i.exec(raw))) {
    // Only the r/g/b components decide the 0–255 vs 0–100% scale: testing the
    // whole argument list let a percentage ALPHA (the legal `rgba(r,g,b,50%)`)
    // multiply the channels by 2.55 and clip every colour to white.
    const rgb = (m[1].match(/-?[\d.]+%?/g) ?? []).slice(0, 3);
    const sc = rgb.some((v) => v.endsWith("%")) ? 2.55 : 1;
    const [r = 0, g = 0, b = 0] = rgb.map((v) => parseFloat(v));
    return to2(r * sc) + to2(g * sc) + to2(b * sc);
  }
  if ((m = /^hsla?\(([^)]*)\)$/i.exec(raw))) {
    // The hue is DEGREES, and `nums` keeps the digits and drops the letters after
    // them — so `hsl(0.5turn 100% 50%)` (CSS cyan) was read as half a degree and
    // drawn red, `3.14159rad` and `200grad` the same, and `1e2` split into 1 and
    // 2 and shifted every component along. Same rule as `toRgb` in
    // src/core/color.ts and the SVG allow-list: a hue that is not a plain number
    // or `deg` is a paint nobody here can read, and UNDEFINED is this function's
    // answer for that (`hex` supplies the black one level up).
    const hue = /^\s*([^\s,/)]*)/.exec(m[1])?.[1] ?? "";
    if (hue !== "" && !/^[+-]?[\d.]+(?:deg)?$/i.test(hue)) return undefined;
    const [hh = 0, ss = 0, ll = 0] = nums(m[1]);
    return hslToHex(hh, ss, ll);
  }
  // Own-property only. CSS_NAMES is a plain object, so a colour literally named
  // "constructor" or "__proto__" reached Object.prototype and this returned a
  // FUNCTION (or an object) — violating the six-hex-digit guarantee above and
  // detonating inside pptxgenjs at writeFile, which is outside every per-chart
  // guard, so one bad colour destroyed the entire batch.
  const key = raw.toLowerCase();
  const named = Object.prototype.hasOwnProperty.call(CSS_NAMES, key) ? CSS_NAMES[key] : undefined;
  // UNDEFINED for "I did not recognise that", not "000000". Returning black here
  // made the failure indistinguishable from a genuine black, so `hexOr` — whose
  // entire job is to tell those apart — could not, and a slide background of
  // `off-white` still came out black. `hex` supplies the black fallback one
  // level up, so its six-hex-digit guarantee is unchanged.
  return typeof named === "string" ? named : undefined;
};

// Opacity 0..1 carried by a paint (8-digit #RRGGBBAA, rgba(), hsla(), or the
// `transparent` keyword); 1 when opaque. The SVG renderer honours the alpha
// natively; here it becomes OOXML transparency — or no paint at all.
export const alphaOf = (c) => {
  const raw = paintText(c);
  if (/^transparent$/i.test(raw)) return 0;
  const h = raw.replace("#", "");
  if (/^[0-9a-fA-F]{8}$/.test(h)) return parseInt(h.slice(6), 16) / 255;
  if (/^[0-9a-fA-F]{4}$/.test(h)) return parseInt(h[3] + h[3], 16) / 255;
  // `rgb`/`hsl` as well as `rgba`/`hsla` — CSS Color 4 made them aliases, and
  // `readHex` above already reads the short spelling's channels.
  const m = /^(?:rgba?|hsla?)\(([^)]*)\)$/i.exec(raw);
  if (m) {
    // The SLASH decides which syntax this is. Splitting on `[,/]` and taking the
    // fourth part reads the legacy comma form and quietly fails the modern one:
    // `rgb(70 130 180 / 0.5)` splits into two parts, not four, so the alpha was
    // dropped and a translucent colour reached the deck solid — while `readHex`
    // parsed its channels perfectly. Kept in step with `alphaOf` in
    // `src/core/color.ts`, which is the same rule for the live add-in; the two
    // are separate code on purpose and have now had the same defect twice.
    const body = m[1];
    const slash = body.indexOf("/");
    const rawAlpha = slash >= 0 ? body.slice(slash + 1) : body.split(",")[3];
    const t = rawAlpha === undefined ? "" : rawAlpha.trim();
    if (t !== "") {
      const a = t.endsWith("%") ? parseFloat(t) / 100 : parseFloat(t);
      return Number.isFinite(a) ? Math.max(0, Math.min(1, a)) : 1;
    }
  }
  return 1;
};

// A pptxgenjs solid fill folding an 8-digit-hex alpha and any scene fillOpacity
// into OOXML transparency (0 = opaque, 100 = clear). A zero transparency is
// dropped by pptxgenjs, so an opaque fill is byte-identical to the bare {color}.
// The STROKE twin of fillOf, and the reason it exists: every stroke site handed
// pptxgenjs a bare `hex(n.stroke)` with no transparency, so a translucent series
// colour rendered at two opacities inside one chart — fill honoured, outline
// opaque — and disagreed with both other renderers. SVG emits
// `stroke="#2a78d659"`; Office.js splits it through `strokeColor` into
// `transparency: 0.651`. Stroke alpha is not among the divergences the parity
// contract at the top of `src/core/scene.ts` declares intentional.
export const lineOf = (color) => {
  const t = Math.round((1 - alphaOf(color)) * 100);
  return t > 0 ? { color: hex(color), transparency: t } : { color: hex(color) };
};

// A text run's colour, with any alpha the paint carries. The TEXT twin of
// `fillOf`/`lineOf`, and the reason it exists: every fill and stroke on a slide
// honoured its alpha and text alone did not, so a deliberately muted label came
// out at full strength in the deck while the preview drew it faint. Zero
// transparency is omitted, so an opaque colour is byte-identical to the bare
// {color} this replaces.
export const textInk = (color) => {
  const t = Math.round((1 - alphaOf(color)) * 100);
  return t > 0 ? { color: hex(color), transparency: t } : { color: hex(color) };
};

export const fillOf = (color, fillOpacity = 1) => {
  const t = Math.round((1 - alphaOf(color) * fillOpacity) * 100);
  // Fully clear: emit no fill rather than a colour. `color: "transparent"` is the
  // documented floating-segment idiom, and it has no hex — painting it as one
  // put a solid block where the preview shows nothing.
  if (t >= 100) return { type: "none" };
  return t > 0 ? { color: hex(color), transparency: t } : { color: hex(color) };
};

/**
 * Characters XML 1.0 forbids outright — they cannot be escaped, only removed, and
 * a single one makes the slide unparseable. A U+000B (a Word line break) in a
 * chart label produced a .pptx that PowerPoint calls corrupt while this renderer
 * reported success. Unpaired surrogates go too; a well-formed pair (an emoji) is
 * a legal character and survives. Mirrors src/render/svg.ts — duplicated because
 * this module ships standalone inside the skill zip.
 */
/* eslint-disable no-control-regex -- matching control characters is the point */
const XML_ILLEGAL =
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
/* eslint-enable no-control-regex */

/** Text safe to place in an OOXML slide. */
export const xmlText = (s) => String(s ?? "").replace(XML_ILLEGAL, "");

/**
 * A label's LINES: its text split at every spelling of a line break.
 *
 * A pasted header can carry one (Alt+Enter in Excel, `FY24\n(Actual)`).
 * `parseClipboardGrid` keeps it on purpose so a wrapped header stays one row, and
 * the datasheet's own note calls the result "a label that the chart draws on two
 * lines". The break reaches the scene as-is and was read three ways: the live
 * renderer hands it to PowerPoint, which makes two paragraphs; pptxgenjs splits it
 * on `\n` into two `<a:p>`; and the SVG preview put it raw inside one `<text>`,
 * where it collapses to a space — one line. The preview was the odd one out, and
 * `src/render/svg.ts` now draws the lines this function names.
 *
 * The spellings were the other half. pptxgenjs splits on LF only, and `xmlText`
 * DELETES VT and FF, so a Word line break glued `Q1\vReport` into `Q1Report`
 * here while the live path passed it to PowerPoint, which breaks the line. EVERY
 * spelling is a break — CRLF, LF, CR, VT, FF, NEL, U+2028 and U+2029 — and the
 * lines leave joined by `\n`, the one thing pptxgenjs understands.
 *
 * Trailing empty lines go. pptxgenjs does not split a text that ENDS in a newline
 * at all, so the break would arrive as a raw CRLF inside one run.
 *
 * NOT folded onto one line, though that would fit a one-line box: folding it here
 * would only make the deck disagree with the live add-in, which draws both lines.
 * The layout sizes the label by its lines instead — `textWidth` reads the widest
 * line and the category axis, the value axes, the unit, the series names and the
 * legend grow their boxes by `labelLines`' count (`src/core/scene.ts`). The
 * other labels a layout boxes as one line are listed in docs/BACKLOG.md. Mirrors
 * `labelLines` there — duplicated because this module ships standalone inside
 * the skill zip.
 */
const LINE_BREAK = /\r\n|[\n\r\v\f\u0085\p{Zl}\p{Zp}]/u;
const labelLines = (s) => {
  const lines = String(s ?? "").split(LINE_BREAK);
  while (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  return lines;
};

/**
 * A font name safe to place in an OOXML ATTRIBUTE.
 *
 * pptxgenjs escapes every text run it writes, but not `fontFace`: it emits
 * `<a:latin typeface="${opts.fontFace}"/>` raw, the one unescaped attribute in
 * the library. A `>` in a font name therefore closes the tag early, and the
 * depth counter in `topLevelElements` — which is how a slide's shapes are
 * found at all — desyncs for the rest of the slide and returns nothing, so the
 * chart is never grouped.
 *
 * `ChartStyle.fontFamily` is a free-form string that round-trips through a
 * style file in localStorage, and scene nodes carry a `fontFamily` of their
 * own. Nothing currently copies one to the other, so there is no live path
 * today — but they are named the same on both ends and connecting them is
 * obviously intended, and the day someone does, every text node becomes an
 * XML-injection point. Strip the characters that cannot survive an attribute
 * rather than rely on that wiring staying unfinished.
 */
export const xmlFontName = (s) => xmlText(s).replace(/[<>&"']/g, "");

/** A paint that would actually draw: present, and not fully transparent. */
export const visible = (paint) => !!paint && alphaOf(paint) > 0;

/**
 * A point size OOXML will accept, enforced where it is WRITTEN.
 *
 * pptxgenjs emits `sz="${fontSize * 100}"` with no checking, and
 * `ST_TextFontSize` is 100..400000 — 1pt to 4000pt. A scene node carrying
 * 0.0001 wrote `sz="0"` and one carrying 1e6 wrote `sz="120000000"`, neither of
 * which is in the type, so PowerPoint offers to repair the deck while the CLI
 * reports a success. Same failure shape as the `x="Infinity"` that `finiteNodes`
 * exists to prevent.
 *
 * The engine clamps `style.fontSize` on the way in now, which closes the config
 * route. This is the other one: `makeAddNode` is a pure exported function that
 * takes a scene node, and this file sits outside `tsconfig.include` and is
 * checked by NOTHING — the same reason the polygon case below refuses to assume
 * `finiteNodes` ran. A guarantee about the bytes belongs at the point that
 * writes them, which is the stance `hex` already takes about its six digits.
 */
const fontPt = (v) => Math.min(4000, Math.max(1, typeof v === "number" && Number.isFinite(v) ? v : 12));

/**
 * Bind the four engine helpers a node mapping needs (dashKind, annularSectorPoints,
 * symbolPreset, arrowheadBox) and return `addNode(slide, n, dx, dy)` — a pure
 * function that maps one scene node to PptxgenJS calls at a slide offset (inches).
 * Taking the engine as a parameter keeps this module free of the top-level await
 * engine load, so it stays importable and measurable.
 */
export function makeAddNode({ dashKind, annularSectorPoints, symbolPreset, arrowheadBox }) {
  return function addNode(slide, n, dx, dy) {
    switch (n.kind) {
      case "rect": {
        slide.addShape("rect", {
          x: dx + n.x * IN,
          y: dy + n.y * IN,
          w: Math.max(0.003, n.w * IN),
          h: Math.max(0.003, n.h * IN),
          // A "none" fill is an outlined/hollow rect (IBCS plan/budget columns).
          fill: n.fill === "none" ? { type: "none" } : fillOf(n.fill),
          line:
            visible(n.stroke) && (n.strokeWidth ?? 0) > 0
              ? { ...lineOf(n.stroke), width: n.strokeWidth }
              : { type: "none" },
        });
        break;
      }
      case "line": {
        // PptxgenJS draws lines TL→BR of the box; flip for rising segments.
        const rising = (n.x2 - n.x1) * (n.y2 - n.y1) < 0;
        slide.addShape("line", {
          x: dx + Math.min(n.x1, n.x2) * IN,
          y: dy + Math.min(n.y1, n.y2) * IN,
          w: Math.abs(n.x2 - n.x1) * IN,
          h: Math.abs(n.y2 - n.y1) * IN,
          flipV: rising,
          line: {
            ...lineOf(n.stroke),
            width: n.strokeWidth ?? 1,
            // `dashKind` answers `none` for an array specifying no dash. Guarding
            // on `n.dash` being truthy put a dotted line in the deck where the
            // preview drew it solid, because `[]` is truthy.
            ...(dashKind(n.dash) === "none"
              ? {}
              : { dashType: dashKind(n.dash) === "dot" ? "sysDot" : "dash" }),
          },
        });
        break;
      }
      case "text": {
        slide.addText(xmlText(labelLines(n.text).join("\n")), {
          x: dx + n.x * IN,
          y: dy + n.y * IN,
          w: Math.max(0.05, n.w * IN),
          h: Math.max(0.05, n.h * IN),
          fontSize: fontPt(n.fontSize),
          // Alpha too. A text colour carrying one — `#0b0b0b80`, an `rgba()`,
          // an `hsla()` — was drawn translucent by the SVG renderer and OPAQUE
          // here, because this took `hex` and threw the alpha away while every
          // fill and stroke on the same slide went through `fillOf`/`lineOf`.
          // Muted ink is how a chart de-emphasises a label, so the two pictures
          // disagreed about which labels mattered. pptxgenjs takes the same
          // 0-100 transparency on a text run as on a shape.
          ...textInk(n.color),
          bold: !!n.bold,
          align: n.align,
          valign: n.valign,
          fontFace: xmlFontName(n.fontFamily) || "Segoe UI",
          margin: 0,
          wrap: false,
        });
        break;
      }
      case "ellipse": {
        slide.addShape("ellipse", {
          x: dx + (n.cx - n.rx) * IN,
          y: dy + (n.cy - n.ry) * IN,
          w: Math.max(0.003, n.rx * 2 * IN),
          h: Math.max(0.003, n.ry * 2 * IN),
          // fill "none" = stroke-only ring (radar circle grid).
          fill: n.fill === "none" ? { type: "none" } : fillOf(n.fill),
          line:
            visible(n.stroke) && (n.strokeWidth ?? 0) > 0
              ? { ...lineOf(n.stroke), width: n.strokeWidth }
              : { type: "none" },
        });
        break;
      }
      case "polygon": {
        // Real freeform geometry: filled, optionally translucent polygons
        // (radar series and grid webs) as native editable shapes.
        // Nothing to bound. `Math.min(...[])` is Infinity, and this writes
        // straight into the OOXML as an EMU offset — `x="Infinity"`, which is
        // not an Int64 and which Microsoft's validator rejects outright, so
        // the whole deck is one PowerPoint may refuse to open. The engine drops
        // these now (`finiteNodes`), and this file is checked by NOTHING —
        // it sits outside tsconfig — so it does not get to assume that.
        if (!Array.isArray(n.points) || n.points.length < 2) break;
        const xs = n.points.map((p) => p.x);
        const ys = n.points.map((p) => p.y);
        const x0 = Math.min(...xs);
        const y0 = Math.min(...ys);
        slide.addShape("custGeom", {
          x: dx + x0 * IN,
          y: dy + y0 * IN,
          w: Math.max(0.01, (Math.max(...xs) - x0) * IN),
          h: Math.max(0.01, (Math.max(...ys) - y0) * IN),
          points: [
            ...n.points.map((p, i) => ({ x: (p.x - x0) * IN, y: (p.y - y0) * IN, moveTo: i === 0 })),
            { close: true },
          ],
          fill: n.fill ? fillOf(n.fill, n.fillOpacity) : { type: "none" },
          line: visible(n.stroke) ? { ...lineOf(n.stroke), width: n.strokeWidth ?? 1 } : { type: "none" },
        });
        break;
      }
      case "wedge": {
        const span = n.endAngle - n.startAngle;
        const x0 = n.cx - n.r;
        const y0 = n.cy - n.r;
        const box = {
          x: dx + x0 * IN,
          y: dy + y0 * IN,
          w: n.r * 2 * IN,
          h: n.r * 2 * IN,
          fill: fillOf(n.fill),
          line: visible(n.stroke) ? { ...lineOf(n.stroke), width: n.strokeWidth ?? 1 } : { type: "none" },
        };
        if (span >= 359.9 && n.innerR <= 0) {
          slide.addShape("ellipse", box);
        } else {
          // Filled sector via custGeom for BOTH the doughnut/gauge ring (innerR>0)
          // and the solid pie wedge (innerR=0, which annularSectorPoints degenerates
          // to a fan from the centre — outer arc forward, then the centre point back).
          // The OOXML "pie" preset can't be used for the solid case: it takes two
          // independently-normalized angles and draws swAng = end − start, so any
          // slice crossing 3 o'clock (every pie has exactly one) gets a negative
          // sweep and renders the wrong wedge. custGeom samples polar() directly, so
          // it's correct across the boundary — the same reason the ring uses it.
          const rel = (p) => ({ x: (p.x - x0) * IN, y: (p.y - y0) * IN });
          const arc = annularSectorPoints(n.cx, n.cy, Math.max(0, n.innerR), n.r, n.startAngle, n.endAngle);
          const half = arc.length / 2; // outer points carry moveTo; inner points don't.
          // A FULL TURN with a hole (a one-group sunburst's inner ring, a one-slice
          // sunburst) is two circles, not one outline. As a single contour the
          // path runs the outer circle, then a straight segment from the outer
          // point at 12 o'clock to the inner one, the inner circle back, and a
          // close that retraces that segment — and an OOXML path strokes every
          // segment, so the wedge's 1pt separator showed as a radial line across
          // the band. The SVG preview draws two evenodd circles and the live fan
          // skips its edges for a full turn (`span >= 359.9`), so only the deck
          // had a seam. Each circle gets its own moveTo and close: the inner one
          // runs the other way round (`annularSectorPoints` returns it reversed),
          // so the fill still leaves the hole.
          const ring = span >= 359.9;
          const pts = arc.map((p, i) =>
            i < half ? { ...rel(p), moveTo: i === 0 } : ring && i === half ? { ...rel(p), moveTo: true } : rel(p),
          );
          if (ring) pts.splice(half, 0, { close: true });
          pts.push({ close: true });
          slide.addShape("custGeom", { ...box, points: pts });
        }
        break;
      }
      case "chevron": {
        slide.addShape(n.flatLeft ? "homePlate" : "chevron", {
          x: dx + n.x * IN,
          y: dy + n.y * IN,
          w: n.w * IN,
          h: n.h * IN,
          fill: fillOf(n.fill),
          line: { type: "none" },
        });
        break;
      }
      case "symbol": {
        // Native preset geometry, so the marker stays FILLED — a custGeom polygon
        // would render here but not in the live add-in. SYMBOL_PRESET names are
        // OOXML preset names, which is exactly what addShape takes.
        slide.addShape(symbolPreset(n.shape), {
          x: dx + (n.cx - n.size) * IN,
          y: dy + (n.cy - n.size) * IN,
          w: n.size * 2 * IN,
          h: n.size * 2 * IN,
          fill: fillOf(n.fill),
          line:
            visible(n.stroke) && (n.strokeWidth ?? 0) > 0
              ? { ...lineOf(n.stroke), width: n.strokeWidth ?? 1 }
              : { type: "none" },
        });
        break;
      }
      case "arrowhead": {
        // Rotated triangle whose tip is offset onto (n.x, n.y) about the box
        // centre — matching the SVG renderer, which anchors the tip. See arrowheadBox.
        const box = arrowheadBox(n.x, n.y, n.size, n.angle);
        slide.addShape("triangle", {
          x: dx + box.left * IN,
          y: dy + box.top * IN,
          w: box.size * IN,
          h: box.size * IN,
          fill: fillOf(n.fill),
          line: { type: "none" },
          rotate: Math.round(box.rotation),
        });
        break;
      }
    }
  };
}
