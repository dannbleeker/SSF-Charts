// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";

/**
 * EVERY TYPED-INTO CONTROL IN THE PANE, RESOLVED THROUGH THE CASCADE, IN BOTH
 * SCHEMES.
 *
 * The bug this exists for: `taskpane.css` states `color: var(--ink)` for every
 * text input and select, and leaves the FILL to each component's own rule.
 * Seven controls are matched by no component rule at all — the sibling rules
 * stop at `input[type="text"]` and `select`, so `#chart-w`, `#chart-h`
 * (the chart size boxes), `#flow-highlight`, `#check-state` and the three
 * Automation selects fell through. With no author background they took the
 * UA's, which is WHITE unless the page declares `color-scheme`. In an Office
 * dark theme that put `--ink` (#dce9f7) on #ffffff:
 *
 *     measured in Chromium under prefers-color-scheme: dark, 2026-09-20,
 *     getComputedStyle(#chart-w) → color rgb(220,233,247) / background
 *     rgb(255,255,255) — 1.23:1, against WCAG 1.4.3's 4.5:1.
 *
 * WHY THIS IS A CSS TEST AND NOT A DOM ONE. jsdom has no cascade worth the
 * name and no UA form-control rendering at all, so `getComputedStyle` on a
 * booted pane answers "rgba(0,0,0,0)" for every one of these and would pass
 * happily while the real pane was unreadable. The rules and the markup are
 * both static, so resolve them here: parse the sheet, match it against the
 * real `taskpane.html`, and do the cascade — specificity, then source order —
 * by hand. Verified against the browser the day it was written: every pair
 * below matches what Chromium computed for the same element.
 */

const css = readFileSync("src/taskpane/taskpane.css", "utf8");
const html = readFileSync("src/taskpane/taskpane.html", "utf8");

/**
 * What the UA paints when the author paints nothing. BOTH values measured in
 * Chromium (the engine behind Office's WebView2) on 2026-09-20, on this
 * markup, under prefers-color-scheme: dark:
 *   without `color-scheme` on :root  → background #ffffff, border #767676
 *   with    `color-scheme: light dark` → background #3b3b3b, border #858585
 * That difference is the whole bug, so the model has to carry it.
 */
const UA_FIELD_LIGHT = "#ffffff";
const UA_FIELD_DARK = "#3b3b3b";
const UA_TEXT_DARK = "#ffffff";

type Rule = { sel: string; body: string; dark: boolean; order: number };

/** Split a selector list on TOP-LEVEL commas: `:where(a, b)` is ONE selector. */
const splitSelectors = (head: string): string[] => {
  const out: string[] = [];
  let depth = 0;
  let cur = "";
  for (const ch of head) {
    if (ch === "(") depth++;
    else if (ch === ")") depth--;
    if (ch === "," && depth === 0) {
      out.push(cur.trim());
      cur = "";
    } else cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
};

/** Flatten the sheet to ordered rules, flagging the ones inside the dark block. */
const flatten = (text: string, dark: boolean, out: Rule[]): Rule[] => {
  let i = 0;
  while (i < text.length) {
    const brace = text.indexOf("{", i);
    if (brace < 0) break;
    const head = text.slice(i, brace).trim();
    let depth = 1;
    let j = brace + 1;
    while (j < text.length && depth > 0) {
      if (text[j] === "{") depth++;
      else if (text[j] === "}") depth--;
      j++;
    }
    const body = text.slice(brace + 1, j - 1);
    if (head.startsWith("@")) {
      // Only the dark block changes the answer; other at-rules (reduced
      // motion, keyframes) are passed through at the scheme they sit in.
      if (/prefers-color-scheme:\s*dark/.test(head)) flatten(body, true, out);
      else if (head.startsWith("@media")) flatten(body, dark, out);
    } else {
      for (const sel of splitSelectors(head)) out.push({ sel, body, dark, order: out.length });
    }
    i = j;
  }
  return out;
};

const rules = flatten(css.replace(/\/\*[\s\S]*?\*\//g, ""), false, []);

/** `prop: value` pairs of one block. Split, rather than match — see below. */
const declarations = (body: string): [string, string][] =>
  body
    .split(";")
    .map((d) => d.trim())
    .filter(Boolean)
    .map((d) => {
      const colon = d.indexOf(":");
      return [d.slice(0, colon).trim(), d.slice(colon + 1).trim()] as [string, string];
    });

const declaredValue = (body: string, prop: string): string | null => {
  let last: string | null = null;
  for (const [p, v] of declarations(body)) if (p === prop) last = v;
  return last;
};

/** `:root` tokens for one scheme — the dark block redefines a subset. */
const tokens = (dark: boolean): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const r of rules) {
    if (r.sel !== ":root" || (r.dark && !dark)) continue;
    for (const [p, v] of declarations(r.body)) if (p.startsWith("--")) out[p] = v;
  }
  return out;
};
const PALETTE = { light: tokens(false), dark: tokens(true) };

const resolve = (value: string, scheme: "light" | "dark"): string => {
  const m = value.match(/var\((--[\w-]+)(?:\s*,\s*([^)]+))?\)/);
  if (!m) return value;
  return (PALETTE[scheme][m[1]] ?? m[2] ?? "").trim();
};

/**
 * Specificity, enough of it for this sheet: ids, then class-likes (classes,
 * attribute selectors, pseudo-classes), then type selectors. `:where()`
 * contributes ZERO by definition, which is the entire point of the fill rule
 * in `taskpane.css` — strip it before counting, and count nothing inside it.
 */
const specificity = (sel: string): number => {
  const s = sel.replace(/:where\([^)]*\)/g, "");
  const ids = (s.match(/#[\w-]+/g) || []).length;
  const classish =
    (s.match(/\.[\w-]+/g) || []).length +
    (s.match(/\[[^\]]*\]/g) || []).length +
    (s.match(/:(?!:)[\w-]+/g) || []).length;
  const types = (s.match(/(^|[\s>+~])[a-z]+/g) || []).length;
  return ids * 10000 + classish * 100 + types;
};

/** Hover/focus/disabled rules describe a state, not the resting control. */
const STATE_SELECTOR = /:(hover|focus|focus-visible|active|disabled|checked|first-child|empty)|::/;

const linear = (channel: number): number => {
  const s = channel / 255;
  return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
};
const luminance = (hex: string): number => {
  const h = hex.replace("#", "");
  return (
    0.2126 * linear(parseInt(h.slice(0, 2), 16)) +
    0.7152 * linear(parseInt(h.slice(2, 4), 16)) +
    0.0722 * linear(parseInt(h.slice(4, 6), 16))
  );
};
/** WCAG 2.x contrast ratio. */
const contrast = (a: string, b: string): number => {
  const la = luminance(a);
  const lb = luminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
};

/** Does `:root` tell the UA this page has a dark rendering? */
const rootDeclaresDarkScheme = rules
  .filter((r) => r.sel === ":root" && !r.dark)
  .some((r) => /dark/.test(declaredValue(r.body, "color-scheme") ?? ""));

/**
 * Controls a user types into or reads a value out of. Checkboxes, ranges and
 * colour swatches are left out: they carry no text, so 1.4.3 does not apply to
 * them and their contrast is a different argument.
 */
const TEXTUAL_INPUT =
  'input[type="text"], input[type="number"], input[type="search"], input[type="email"], input[type="url"], input[type="tel"], input[type="date"], input[type="time"], input:not([type]), select, textarea';

const doc = new DOMParser().parseFromString(html, "text/html");
const controls = [...doc.querySelectorAll(TEXTUAL_INPUT)] as HTMLElement[];

/**
 * BOTH SHELLS THAT LOAD THIS SHEET. `src/excel/excel.html` links the same
 * `taskpane.css`, so a rule written for the PowerPoint pane is a rule the Excel
 * pane wears too — and the bug being fixed here was precisely a control nobody
 * had thought to style. Its three controls all sit in `.field` and were never
 * broken; they are here so the next one added is checked in the file it is
 * added to, rather than in the pane somebody remembered.
 */
const excelDoc = new DOMParser().parseFromString(readFileSync("src/excel/excel.html", "utf8"), "text/html");
const excelControls = [...excelDoc.querySelectorAll(TEXTUAL_INPUT)] as HTMLElement[];
const allControls = [
  ...controls.map((el) => ({ el, file: "taskpane.html" })),
  ...excelControls.map((el) => ({ el, file: "excel.html" })),
];

const name = (el: HTMLElement) =>
  el.id || `${el.tagName.toLowerCase()}[${el.getAttribute("type") ?? ""}] ${el.outerHTML.slice(0, 60)}`;

/** The winning declaration for one property, by specificity then source order. */
const winner = (el: HTMLElement, scheme: "light" | "dark", prop: string): string | null => {
  let best: { value: string; spec: number; order: number } | null = null;
  for (const r of rules) {
    if (r.dark && scheme !== "dark") continue;
    if (STATE_SELECTOR.test(r.sel)) continue;
    let matches: boolean;
    try {
      matches = el.matches(r.sel);
    } catch {
      matches = false; // a selector jsdom cannot parse cannot be reasoned about
    }
    if (!matches) continue;
    const value =
      declaredValue(r.body, prop) ?? (prop === "background" ? declaredValue(r.body, "background-color") : null);
    if (value == null) continue;
    const spec = specificity(r.sel);
    if (!best || spec > best.spec || (spec === best.spec && r.order > best.order))
      best = { value, spec, order: r.order };
  }
  return best ? best.value : null;
};

const resolved = (el: HTMLElement, scheme: "light" | "dark") => {
  const colorDecl = winner(el, scheme, "color");
  const bgDecl = winner(el, scheme, "background");
  const uaField = scheme === "dark" && rootDeclaresDarkScheme ? UA_FIELD_DARK : UA_FIELD_LIGHT;
  const uaText = scheme === "dark" && rootDeclaresDarkScheme ? UA_TEXT_DARK : "#000000";
  return {
    authored: bgDecl != null,
    color: colorDecl ? resolve(colorDecl, scheme) : uaText,
    background: bgDecl ? resolve(bgDecl, scheme) : uaField,
  };
};

describe("pane inputs are readable in both Office themes", () => {
  it("finds the pane's controls at all (a guard on the selector, not the product)", () => {
    // If the markup is reorganised out from under this file, the two tests
    // below would pass vacuously on an empty list. They must not.
    expect(controls.length).toBeGreaterThan(10);
    expect(controls.map((c) => c.id)).toContain("chart-w");
    expect(excelControls.map((c) => c.id)).toContain("kind");
  });

  /**
   * `color-scheme` is the only thing that reaches what CSS cannot paint: the
   * number spinners, the select's arrow, and the option list a `<select>` pops
   * open. Those options inherit the select's `color` — measured
   * rgb(220,233,247) — onto a UA-painted background, so without this
   * declaration a dark-theme user opens a white popup of near-white text. It
   * is also the backstop for any control the fill rule does not name.
   */
  it("declares a color-scheme, so the UA stops painting the rest of the pane light", () => {
    expect(rootDeclaresDarkScheme).toBe(true);
  });

  it("gives every typed-into control a fill of its own, never the UA's", () => {
    const orphans = allControls
      .filter(({ el }) => !resolved(el, "dark").authored)
      .map(({ el, file }) => `${file}: ${name(el)}`);
    // Not a style preference: an author fill is the only one that is the
    // pane's own colour. The UA's dark field (#3b3b3b) is readable but foreign
    // — a grey box among navy ones — and its light field is white, which is
    // what made the size boxes unreadable in the first place.
    expect(orphans).toEqual([]);
  });

  it("clears WCAG AA (4.5:1) for body text in light AND dark", () => {
    const failures: string[] = [];
    for (const scheme of ["light", "dark"] as const) {
      for (const { el, file } of allControls) {
        const { color, background } = resolved(el, scheme);
        const ratio = contrast(color, background);
        if (ratio < 4.5)
          failures.push(`${scheme}: ${file}: ${name(el)} ${color} on ${background} = ${ratio.toFixed(2)}:1`);
      }
    }
    expect(failures).toEqual([]);
  });

  it("measures the two pairs the fix turns on, so the numbers are on the record", () => {
    // The chart size box, which is what the audit reported as unreadable.
    const sizeBox = doc.getElementById("chart-w") as HTMLElement;
    const dark = resolved(sizeBox, "dark");
    expect(dark.color).toBe("#dce9f7");
    expect(dark.background).toBe("#12202f");
    expect(contrast(dark.color, dark.background)).toBeCloseTo(13.38, 1);
    // Light is unchanged to the pixel: `--panel` IS #ffffff, which is exactly
    // what the UA was painting there before the fill rule existed.
    const light = resolved(sizeBox, "light");
    expect(light.background).toBe("#ffffff");
    expect(contrast(light.color, light.background)).toBeCloseTo(16.32, 1);
  });

  it("leaves a component that states its own surface alone — the fill is a default", () => {
    // `.type-search input` deliberately recesses into the page with `--surface`
    // rather than `--panel`. The fill rule is `:where(...)`, so it has zero
    // specificity and cannot take that over; a plain selector list would tie
    // with `.type-search input` and win on source order. Confirmed in Chromium:
    // the search box still computes rgb(10,20,32) in dark.
    const search = doc.getElementById("type-search-input") as HTMLElement;
    expect(resolved(search, "dark").background).toBe("#0a1420");
  });
});
