import { polar, symbolPoints, dashKind } from "../core/geometry";
import { isNamedColor } from "../core/color";
import { LINE_PITCH, labelLines, type Scene, type SceneNode } from "../core/scene";

/**
 * Characters XML 1.0 forbids OUTRIGHT — they cannot be escaped, only removed.
 * The C0 controls except tab/LF/CR, UNPAIRED surrogates (a well-formed pair is a
 * legal astral character — an emoji — and must survive), and the two permanently
 * unassigned code points. A single one makes the whole document unparseable, so
 * a chart label carrying (say) a U+000B line break from a Word paste produced an
 * .svg no strict parser would open, and a .pptx PowerPoint calls corrupt — while
 * the renderer reported success. Escaping `& < > "` never covered this: they are
 * not markup, they are simply illegal.
 */
/* eslint-disable no-control-regex -- matching control characters is the point */
const XML_ILLEGAL =
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
/* eslint-enable no-control-regex */

/**
 * Drop what XML cannot represent, then escape what it can.
 *
 * Coerces first, because the thing arriving is not reliably a string. A chart
 * `title` of `2024` — a number in a JSON config, which is what someone writes
 * for a year — reached here and threw `s.replace is not a function`, taking the
 * whole render with it. Same for `valueAxisTitle`, and for any text a config
 * supplies. The type says `string`; the config is a file someone pasted.
 *
 * `String()` rather than a rejection: a number, a boolean or a date is text the
 * author plainly meant to show, and rendering `2024` is the right answer rather
 * than a merciful one. An object comes out as `[object Object]`, which is
 * visible nonsense on the slide — and visible nonsense beats a chart that did
 * not draw, because the author can see it and fix it.
 */
export const xmlText = (s: string) => String(s ?? "").replace(XML_ILLEGAL, "");

const esc = (s: string) =>
  xmlText(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/*
 * A label's LINES (`labelLines`) and the pitch between them (`LINE_PITCH`) come
 * from `src/core/scene.ts`, where the LAYOUT sizes a label's box from the same
 * two — they were local to this file, and the layout read a two-line label as one
 * wide line and boxed it one line tall while this drew the second line below the
 * box. The twin of `labelLines` in `skill/scripts/pptx-paint.mjs`, whose note has
 * the whole story: a pasted header can carry a line break, the live renderer and
 * the deck both draw it as two lines, and this preview emitted it raw inside one
 * `<text>`, where SVG collapses it to a space. And VT/FF, which `xmlText`
 * deletes, glued `Q1<VT>Report` into `Q1Report` here and in the deck.
 */

/**
 * Colours reach the scene graph verbatim from ChartConfig (series color, custom
 * palette, per-point colors) and are interpolated straight into SVG paint
 * attributes. A value like `#000"><image href=x onerror=alert(1)>` would break
 * out of the attribute and inject nodes that execute when this SVG is assigned
 * via innerHTML (the pane preview) or opened as a standalone document — stored
 * XSS reachable cross-user through a POWERCHART_CONFIG shape tag. Allow only the
 * shapes a paint can legitimately take; anything else falls back to a safe ink.
 * Every colour the engine actually produces (hex, and internally-built url(#…))
 * passes unchanged, so valid charts render byte-identically.
 */
/**
 * The `/` is CSS Color 4's alpha separator — `rgb(70 130 180 / 0.5)`, the form
 * MDN now documents first and the one an agent writing a config is most likely
 * to produce. Without it that colour failed this test and fell back to BLACK in
 * the preview, while both PowerPoint renderers parsed it and drew steel blue:
 * one colour, two pictures, and only the wrong one on screen.
 *
 * It widens the allow-list by one character and cannot widen what escapes.
 * Everything inside the parentheses is still limited to digits, `.`, `,`,
 * whitespace, `%` and now `/` — no quote to close the attribute, no `<` or `>`
 * to open a tag, no `&` to start an entity, and no `*`, so it cannot even open
 * a CSS comment. The alternation is anchored, so a slash outside the function
 * form still fails.
 *
 * The same one-colour-two-pictures gap stayed open in three more spellings the
 * other sinks already read (`toRgb` in `src/core/color.ts`, and `readHex` in
 * `skill/scripts/pptx-paint.mjs`, both case-insensitive and both taking `-?` on a
 * channel): an UPPER or mixed-case function name (`RGB(…)` — CSS function names
 * are case-insensitive), a sign (`rgb(300,-20,50)`, `hsl(-30 …)`, which CSS clamps
 * and wraps the way those parsers do), and the `deg` unit on a hue
 * (`hsl(210deg 60% 45%)`, which is just the default unit spelled out). Each fell
 * back to BLACK here and drew the real colour in the deck. The bar was black and
 * its in-bar value label was inked from the REAL colour, so a mid-light fill got
 * near-black text on a black bar — invisible in the preview and in every picture
 * insert, which are drawn from this same string (`rasterizeScene`).
 *
 * Widened by the `i` flag, the two sign characters and the one unit word, so it
 * still cannot widen what escapes: inside the parentheses there is still no quote,
 * `<`, `>`, `&` or `*`, and the only letters allowed are the whole word `deg`.
 * The other angle units (`turn`, `rad`, `grad`) are deliberately NOT admitted:
 * `toRgb` reads the number and drops the unit, so those hues are misread by every
 * sink, and letting this one draw the true colour would only add a third answer.
 */
const PAINT_OK = /^(#[0-9a-f]{3,8}|(?:rgba?|hsla?)\((?:[\d.,\s%/+-]|deg)+\)|url\(#[\w.-]+\)|[a-z]{1,24})$/i;
/**
 * A paint the markup can carry, or black.
 *
 * TRIMMED first, which the three other colour parsers in this repo have always
 * done (`paintText` in `src/core/color.ts` and in `skill/scripts/pptx-paint.mjs`,
 * and `officeHex`) and this one did not — so a colour with a stray space around
 * it, exactly what a pasted palette entry or a hand-edited config carries,
 * rendered BLACK in the reference renderer while both PowerPoint renderers
 * rendered it correctly. Non-strings read as empty for the same reason they do
 * there: the type says `string`, and TypeScript checks the code, not the file
 * someone pasted.
 */
/**
 * A BARE hex string, `#` omitted — OOXML's own spelling, which the pptx sink has
 * always accepted and `toRgb` now does too.
 *
 * This sink needs more than permission to carry it: `AABBCC` is a run of letters,
 * so the allow-list's named-colour arm passed it through UNCHANGED and the markup
 * carried `fill="AABBCC"`, which is not a valid SVG paint — the attribute is
 * ignored and the shape renders black. `4e79a7` has digits, failed every arm, and
 * fell back to black explicitly. Two spellings of the same colour, both black,
 * while the deck drew the real one.
 *
 * So it is normalised rather than allowed: the `#` is put back, and the value
 * then goes through the same allow-list as everything else rather than around it.
 */
const BARE_HEX = /^([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;
const paintOr = (c: unknown, fallback: string): string => {
  const raw = typeof c === "string" ? c.trim() : "";
  // A named colour wins the string, as it does in `toRgb` — and as there, no CSS
  // name is valid bare hex, so this decides nothing today. It is the ordering
  // that keeps it true if the name table ever grows one.
  const t = !isNamedColor(raw) && BARE_HEX.test(raw) ? `#${raw}` : raw;
  return t && PAINT_OK.test(t) ? t : fallback;
};
const paint = (c: unknown): string => paintOr(c, "#000000");

/**
 * NUMERIC attributes are the second injection surface, and the one the colour
 * allow-list above does not cover. Sizes, widths and opacities reach the markup
 * by raw interpolation, and ChartConfig's numeric fields (`style.fontSize`,
 * `decorations.fillOpacity`, …) are only `number` in TypeScript — a type erased
 * at runtime. Untrusted JSON therefore carries a STRING straight into an
 * attribute: `fontSize: '10"><image href=x onerror=alert(1) /><text x="'` closes
 * the element and injects an executing node. That JSON is genuinely untrusted —
 * a `#c=` share link, an imported config, or a POWERCHART_CONFIG shape tag
 * authored in another deck — and the pane assigns this SVG via innerHTML.
 *
 * Coerce to a finite number or fall back. Every value the engine actually
 * produces is already numeric, so valid charts render byte-identically.
 */
const num = (v: unknown, fallback: number): number => {
  const n = typeof v === "number" ? v : Number(v);
  return Number.isFinite(n) ? n : fallback;
};

/** Hatch/dot pattern tiles: series-colored base with white strokes over it. */
const PATTERN_TILE: Record<string, (id: string, color: string) => string> = {
  diagonal: (id, c) =>
    `<pattern id="${id}" width="5" height="5" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="5" height="5" fill="${c}"/><line x1="0" y1="0" x2="0" y2="5" stroke="#ffffff" stroke-width="1.4" stroke-opacity="0.75"/></pattern>`,
  crosshatch: (id, c) =>
    `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="6" height="6" fill="${c}"/><path d="M 0 0 V 6 M 0 0 H 6" stroke="#ffffff" stroke-width="1.1" stroke-opacity="0.7"/></pattern>`,
  dots: (id, c) =>
    `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse"><rect width="6" height="6" fill="${c}"/><circle cx="3" cy="3" r="1.1" fill="#ffffff" fill-opacity="0.8"/></pattern>`,
  horizontal: (id, c) =>
    `<pattern id="${id}" width="5" height="5" patternUnits="userSpaceOnUse"><rect width="5" height="5" fill="${c}"/><line x1="0" y1="1" x2="5" y2="1" stroke="#ffffff" stroke-width="1.4" stroke-opacity="0.75"/></pattern>`,
};

// A fragment id must be a safe token. A hex paint yields one after dropping the
// "#", but an rgb()/hsl() paint (both accepted by PAINT_OK) carries "(", ")" and
// "," — which break the <pattern id> AND make url(#…) unresolvable, since the URL
// parser stops at the first ")". Mapping those chars to "-" is a no-op for hex
// (byte-identical output) and produces a resolvable id for the rest.
const patternId = (pattern: string, color: string) => `p-${pattern}-${color.replace("#", "").replace(/[^\w-]/g, "-")}`;

/**
 * The tile builder for a pattern name, or undefined.
 *
 * `PATTERN_TILE[name]` alone reaches Object.prototype: `pattern: "__proto__"`
 * came back as an object (truthy, not callable → the whole render threw), and
 * `"constructor"` / `"toString"` came back as functions that were then CALLED as
 * tile builders, writing garbage into <defs> and a url(#…) reference to a
 * <pattern> that was never emitted. The name is unvalidated config. The sibling
 * pptx renderer already guards its colour-name table the same way.
 */
const patternTile = (pattern: string) =>
  Object.prototype.hasOwnProperty.call(PATTERN_TILE, pattern) ? PATTERN_TILE[pattern] : undefined;

/** Render a scene to a standalone SVG string (1pt = 1px). */
export function sceneToSvg(scene: Scene, opts: { background?: string } = {}): string {
  const parts: string[] = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${num(scene.width, 480)}" height="${num(scene.height, 300)}" viewBox="0 0 ${num(scene.width, 480)} ${num(scene.height, 300)}" font-family="Segoe UI, Arial, sans-serif" role="img">`,
  );
  // Accessible name + text alternative as the first children of the root, so a
  // screen reader announces the chart under role="img" instead of skipping it
  // (WCAG 1.1.1). First-child <title>/<desc> are the accessible name/description
  // per the SVG mapping — no id/aria wiring needed, so multiple inline charts on
  // one page can't collide on ids.
  //
  // An UNTITLED chart (the pane's own gallery shape) has only the generated
  // summary, and <desc> maps to the description alone: role="img" with no name
  // is exactly what axe-core's role-img-alt reports as a 1.1.1 failure. So the
  // summary becomes the name in that case — announced once, not twice.
  if (scene.title) {
    parts.push(`<title>${esc(scene.title)}</title>`);
    if (scene.desc) parts.push(`<desc>${esc(scene.desc)}</desc>`);
  } else if (scene.desc) {
    parts.push(`<title>${esc(scene.desc)}</title>`);
  }
  // One <pattern> def per (pattern, color) pair used by the scene's rects.
  const defs = new Map<string, string>();
  for (const n of scene.nodes) {
    const tile = n.kind === "rect" && n.pattern ? patternTile(n.pattern) : undefined;
    if (tile && n.kind === "rect" && n.pattern) {
      const fill = paint(n.fill);
      const id = patternId(n.pattern, fill);
      if (!defs.has(id)) defs.set(id, tile(id, fill));
    }
  }
  if (defs.size) parts.push(`<defs>${[...defs.values()].join("")}</defs>`);
  if (opts.background) {
    // An UNREADABLE background is the one place `paint`'s black fallback is the
    // wrong answer: this rect covers the whole canvas and is drawn before every
    // node, so a background of "not-a-colour" — or of anything the allow-list
    // declines — painted the chart out entirely. Nothing else in the scene can
    // hide behind a bad value like that. Fall back to the canvas the engine
    // documents as its default instead.
    parts.push(`<rect width="100%" height="100%" fill="${paintOr(opts.background, "#ffffff")}"/>`);
  }
  for (const n of scene.nodes) parts.push(nodeToSvg(n));
  parts.push("</svg>");
  return parts.join("\n");
}

function nodeToSvg(n: SceneNode): string {
  switch (n.kind) {
    case "rect": {
      const stroke = n.stroke ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"` : "";
      const fill = n.pattern && patternTile(n.pattern) ? `url(#${patternId(n.pattern, paint(n.fill))})` : paint(n.fill);
      return `<rect x="${r(n.x)}" y="${r(n.y)}" width="${r(n.w)}" height="${r(n.h)}" fill="${fill}"${stroke}${name(n)}/>`;
    }
    case "line": {
      // `dashKind` decides, not `length`: an array with no positive finite value
      // is not a dash, and emitting `stroke-dasharray="0 0"` / `"-5 -5"` for one
      // is markup SVG ignores anyway. Asking the same question all three sinks
      // ask is what keeps the preview and the two decks agreeing.
      const dash = dashKind(n.dash) !== "none" ? ` stroke-dasharray="${n.dash!.map((d) => num(d, 0)).join(" ")}"` : "";
      return `<line x1="${r(n.x1)}" y1="${r(n.y1)}" x2="${r(n.x2)}" y2="${r(n.y2)}" stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"${dash}${name(n)}/>`;
    }
    case "text": {
      const anchor = n.align === "left" ? "start" : n.align === "right" ? "end" : "middle";
      const x = n.align === "left" ? n.x : n.align === "right" ? n.x + n.w : n.x + n.w / 2;
      // Approximate vertical alignment with baseline offsets.
      const lines = labelLines(n.text);
      // The block of lines hangs from the top, sits on the bottom or centres
      // on the middle, as a PowerPoint text box with that anchor lays out its
      // paragraphs — so the first baseline moves up by the lines beneath it.
      // Zero for the ordinary one-line label, whose markup is unchanged.
      const rise =
        (lines.length - 1) * n.fontSize * LINE_PITCH * (n.valign === "top" ? 0 : n.valign === "bottom" ? 1 : 0.5);
      const y =
        (n.valign === "top"
          ? n.y + n.fontSize
          : n.valign === "bottom"
            ? n.y + n.h - n.fontSize * 0.25
            : n.y + n.h / 2 + n.fontSize * 0.36) - rise;
      const weight = n.bold ? ` font-weight="600"` : "";
      const family = n.fontFamily ? ` font-family="${esc(n.fontFamily)}"` : "";
      // `num` substitutes for a non-finite size and nothing else, so zero and
      // negative reached the markup: `font-size="0"` renders nothing and
      // `font-size="-5"` is invalid, so the attribute is dropped and the label
      // is drawn at whatever the document inherits. Both are a label that
      // silently is not what the scene asked for, which is the preview lying
      // about what the deck will contain. The third sink's own floor, matching
      // the clamps `pptx-paint.mjs` and `officeFontPt` already apply.
      // One `<tspan>` per line, each placed absolutely: SVG never breaks a line
      // on a newline, so the break has to be drawn. A one-line label keeps the
      // bare text node it always had.
      const body =
        lines.length === 1
          ? esc(lines[0])
          : lines
              .map((line, i) =>
                line ? `<tspan x="${r(x)}" y="${r(y + i * n.fontSize * LINE_PITCH)}">${esc(line)}</tspan>` : "",
              )
              .join("");
      return `<text x="${r(x)}" y="${r(y)}" font-size="${Math.max(0.01, num(n.fontSize, 12))}" fill="${paint(n.color)}" text-anchor="${anchor}"${weight}${family}${name(n)}>${body}</text>`;
    }
    case "ellipse": {
      const stroke = n.stroke ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"` : "";
      return `<ellipse cx="${r(n.cx)}" cy="${r(n.cy)}" rx="${r(n.rx)}" ry="${r(n.ry)}" fill="${paint(n.fill)}"${stroke}${name(n)}/>`;
    }
    case "chevron": {
      // Fixed notch fraction. The PowerPoint renderers instead name the native
      // chevron/homePlate preset, whose default point depth differs — an
      // intentional preview-vs-deck approximation (see the parity contract in
      // scene.ts).
      const notch = n.h * 0.28;
      const left = n.flatLeft
        ? `M ${r(n.x)} ${r(n.y)}`
        : `M ${r(n.x)} ${r(n.y)} L ${r(n.x + notch)} ${r(n.y + n.h / 2)} L ${r(n.x)} ${r(n.y + n.h)}`;
      const d = n.flatLeft
        ? `M ${r(n.x)} ${r(n.y)} L ${r(n.x + n.w - notch)} ${r(n.y)} L ${r(n.x + n.w)} ${r(n.y + n.h / 2)} L ${r(n.x + n.w - notch)} ${r(n.y + n.h)} L ${r(n.x)} ${r(n.y + n.h)} Z`
        : `${left} L ${r(n.x)} ${r(n.y + n.h)} L ${r(n.x + n.w - notch)} ${r(n.y + n.h)} L ${r(n.x + n.w)} ${r(n.y + n.h / 2)} L ${r(n.x + n.w - notch)} ${r(n.y)} Z`;
      return `<path d="${d}" fill="${paint(n.fill)}"${name(n)}/>`;
    }
    case "wedge": {
      const stroke = n.stroke ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"` : "";
      // A full 360° span has coincident start/end points, so the arc collapses
      // and SVG draws nothing. Emit an explicit circle (or ring) instead.
      if (n.endAngle - n.startAngle >= 359.9) {
        const ring = (rad: number, sweep: number) => {
          const a = polar(n.cx, n.cy, rad, 0);
          const b = polar(n.cx, n.cy, rad, 180);
          return `M ${r(a.x)} ${r(a.y)} A ${r(rad)} ${r(rad)} 0 1 ${sweep} ${r(b.x)} ${r(b.y)} A ${r(rad)} ${r(rad)} 0 1 ${sweep} ${r(a.x)} ${r(a.y)} Z`;
        };
        const d = n.innerR > 0 ? `${ring(n.r, 1)} ${ring(n.innerR, 0)}` : ring(n.r, 1);
        return `<path d="${d}" fill="${paint(n.fill)}" fill-rule="evenodd"${stroke}${name(n)}/>`;
      }
      const large = n.endAngle - n.startAngle > 180 ? 1 : 0;
      const o1 = polar(n.cx, n.cy, n.r, n.startAngle);
      const o2 = polar(n.cx, n.cy, n.r, n.endAngle);
      let d: string;
      if (n.innerR > 0) {
        const i1 = polar(n.cx, n.cy, n.innerR, n.endAngle);
        const i2 = polar(n.cx, n.cy, n.innerR, n.startAngle);
        d = `M ${r(o1.x)} ${r(o1.y)} A ${r(n.r)} ${r(n.r)} 0 ${large} 1 ${r(o2.x)} ${r(o2.y)} L ${r(i1.x)} ${r(i1.y)} A ${r(n.innerR)} ${r(n.innerR)} 0 ${large} 0 ${r(i2.x)} ${r(i2.y)} Z`;
      } else {
        d = `M ${r(n.cx)} ${r(n.cy)} L ${r(o1.x)} ${r(o1.y)} A ${r(n.r)} ${r(n.r)} 0 ${large} 1 ${r(o2.x)} ${r(o2.y)} Z`;
      }
      return `<path d="${d}" fill="${paint(n.fill)}"${stroke}${name(n)}/>`;
    }
    case "polygon": {
      const pts = n.points.map((p) => `${r(p.x)},${r(p.y)}`).join(" ");
      const fill = n.fill
        ? ` fill="${paint(n.fill)}"${n.fillOpacity != null ? ` fill-opacity="${num(n.fillOpacity, 1)}"` : ""}`
        : ` fill="none"`;
      const stroke = n.stroke
        ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}" stroke-linejoin="round"`
        : "";
      return `<polygon points="${pts}"${fill}${stroke}${name(n)}/>`;
    }
    case "arrowhead": {
      // Narrow isosceles triangle with tip at (x, y), pointing along angle. The
      // PowerPoint renderers name the native `triangle` preset in a 2*size box,
      // which is broader — an intentional approximation (see scene.ts parity
      // contract); only the tip anchor and angle are shared exactly.
      const s = n.size;
      return `<path d="M 0 0 L ${-s * 1.8} ${-s * 0.7} L ${-s * 1.8} ${s * 0.7} Z" fill="${paint(n.fill)}" transform="translate(${r(n.x)} ${r(n.y)}) rotate(${r(n.angle)})"${name(n)}/>`;
    }
    case "symbol": {
      // The other two renderers name a preset; here we draw its outline.
      const pts = symbolPoints(n.shape, n.cx, n.cy, n.size)
        .map((p) => `${r(p.x)},${r(p.y)}`)
        .join(" ");
      const stroke = n.stroke
        ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}" stroke-linejoin="round"`
        : "";
      return `<polygon points="${pts}" fill="${paint(n.fill)}"${stroke}${name(n)}/>`;
    }
  }
}

const r = (v: number) => Math.round(v * 100) / 100;
const name = (n: SceneNode) => (n.name ? ` data-name="${esc(n.name)}"` : "");
