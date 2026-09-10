/**
 * Pane localization. The app ships English-only, but every user-facing string is
 * catalogued here so a new language is a translation-only, type-checked drop-in.
 *
 * ── Adding a language ────────────────────────────────────────────────────────
 * 1. Add one entry to `DICTS`, e.g. `de: { ...every key of EN, translated }`.
 *    Its type is `Record<StringKey, string>`, so TypeScript refuses to compile
 *    the dict until EVERY key is present — a missing translation is a build error,
 *    not a silent English fallback. `test/i18n.test.ts` double-checks at runtime.
 * 2. That's it. `localizePane` auto-selects it from the host display language
 *    (or `?lang=`), the DOM sweep re-skins static markup, and `t()` covers the
 *    runtime status strings — including the `{placeholder}` templates.
 *
 * Chart OUTPUT is already language-neutral: its text is user data, its numbers
 * and dates localize via `NumberFormat.locale`, and the one engine-injected word
 * ("Other") is fed in through `ChartConfig.labels.other` — the pane sets it to
 * `t("Other")`, the pure engine just uses whatever it's given.
 */

/**
 * The canonical catalogue of every translatable string — UI chrome, runtime
 * status messages, `{placeholder}` templates, insert-phase words, and the one
 * engine label. English maps to itself; this object IS the key set a language
 * must cover.
 */
export const EN = {
  // Step headings, panel & tab titles.
  "1 · Chart type": "1 · Chart type",
  "2 · Data": "2 · Data",
  "3 · Decorations": "3 · Decorations",
  "Preview & size": "Preview & size",
  Elements: "Elements",
  "Automation (JSON)": "Automation (JSON)",
  Agenda: "Agenda",
  // Action buttons.
  "Insert into slide": "Insert into slide",
  "Insert as new": "Insert as new",
  "Edit selected chart": "Edit selected chart",
  "Same scale (deck)": "Same scale (deck)",
  "Same scale (selection)": "Same scale (selection)",
  "Download SVG": "Download SVG",
  "Download PNG": "Download PNG",
  "Copy chart link": "Copy chart link",
  "Update chart": "Update chart",
  "Insert agenda slides": "Insert agenda slides",
  "Export current": "Export current",
  Import: "Import",
  "Insert batch": "Insert batch",
  "Export style": "Export style",
  "Import style": "Import style",
  "Save style to deck": "Save style to deck",
  "Use deck style": "Use deck style",
  "Save as template": "Save as template",
  Delete: "Delete",
  Insert: "Insert",
  "Edit it": "Edit it",
  // Decoration & option labels.
  "Chart title": "Chart title",
  "Segment labels": "Segment labels",
  "Series labels": "Series labels",
  "Column totals": "Column totals",
  "Grand total": "Grand total",
  "Category labels": "Category labels",
  "Value axis": "Value axis",
  Gridlines: "Gridlines",
  "Horizontal (bar)": "Horizontal (bar)",
  "Connector lines": "Connector lines",
  "100% = note": "100% = note",
  "Auto-update chart": "Auto-update chart",
  "Total row": "Total row",
  "Datamark axis (ticks only)": "Datamark axis (ticks only)",
  "Use deck theme": "Use deck theme",
  // Chart-type families (grouped picker) + its search.
  "Columns & bars": "Columns & bars",
  "Line & area": "Line & area",
  "Parts of a whole": "Parts of a whole",
  Distribution: "Distribution",
  Correlation: "Correlation",
  "Matrix & spatial": "Matrix & spatial",
  "Search chart types…": "Search chart types…",
  "No chart type matches that search.": "No chart type matches that search.",
  // Format group names.
  Labels: "Labels",
  "Axes & scale": "Axes & scale",
  Analysis: "Analysis",
  "Colours & style": "Colours & style",
  // Datasheet help.
  "Paste straight from Excel — special data rows": "Paste straight from Excel — special data rows",

  // Runtime status — plain messages.
  "Working…": "Working…",
  "Done.": "Done.",
  Stop: "Stop",
  "Placed beside the last chart — drag or resize it as you like.":
    "Placed beside the last chart — drag or resize it as you like.",
  "Placed beside the last chart and scaled to fit — drag or resize it as you like.":
    "Placed beside the last chart and scaled to fit — drag or resize it as you like.",
  "Stopping…": "Stopping…",
  "Stopped — anything already drawn was kept.": "Stopped — anything already drawn was kept.",
  "That chart is no longer on the slide — insert it again.": "That chart is no longer on the slide — insert it again.",
  "Chart loaded — edits will update it in place.": "Chart loaded — edits will update it in place.",
  "Chart loaded from a shared link.": "Chart loaded from a shared link.",
  "Chart config loaded.": "Chart config loaded.",
  "Style exported — share the JSON as your corporate style file.":
    "Style exported — share the JSON as your corporate style file.",
  "Style imported — applied to every chart from this pane.": "Style imported — applied to every chart from this pane.",
  "Style saved into this deck — it travels with the file.": "Style saved into this deck — it travels with the file.",
  "This host cannot store a deck style (needs PowerPoint API 1.7).":
    "This host cannot store a deck style (needs PowerPoint API 1.7).",
  "Using this deck's own style.": "Using this deck's own style.",
  "This deck carries no style — using yours.": "This deck carries no style — using yours.",
  "The selection is not an SSF chart — select an inserted chart group first.":
    "The selection is not an SSF chart — select an inserted chart group first.",
  "Same scale needs at least two value-axis charts in the deck.":
    "Same scale needs at least two value-axis charts in the deck.",
  "Select two or more charts (Ctrl-click), then apply Same scale.":
    "Select two or more charts (Ctrl-click), then apply Same scale.",
  "Couldn't encode the PNG on this browser.": "Couldn't encode the PNG on this browser.",
  "Couldn't render the preview to PNG.": "Couldn't render the preview to PNG.",
  "Shareable chart link copied to the clipboard.": "Shareable chart link copied to the clipboard.",
  "Clipboard blocked — the link is in the JSON box, copy it from there.":
    "Clipboard blocked — the link is in the JSON box, copy it from there.",
  "Not running inside PowerPoint — use Download SVG, or sideload the manifest to insert native shapes.":
    "Not running inside PowerPoint — use Download SVG, or sideload the manifest to insert native shapes.",

  // Runtime status — {placeholder} templates (filled by t(key, params)).
  "Failed: {error}": "Failed: {error}",
  "Couldn't render PNG: {error}": "Couldn't render PNG: {error}",
  // The one note the datasheet raises on its own. A translator should keep the
  // two worked examples: they are what let a user check the reading at a glance
  // rather than take it on trust, and the digits carry that in any language.
  "Read as European numbers — 1.234 as 1234, 987,5 as 987.5. Rewrote {n} cell(s); edit any that are wrong.":
    "Read as European numbers — 1.234 as 1234, 987,5 as 987.5. Rewrote {n} cell(s); edit any that are wrong.",
  "Style import failed: {error}": "Style import failed: {error}",
  "Invalid JSON: {error}": "Invalid JSON: {error}",
  "That is valid JSON, but not a chart this pane can open: {error}":
    "That is valid JSON, but not a chart this pane can open: {error}",
  "Host answered late — {message}": "Host answered late — {message}",
  "Same scale applied to {n} charts (max {max}).": "Same scale applied to {n} charts (max {max}).",
  // The two clauses that qualify the line above. Catalogued because they are
  // now composed WITH it: uncatalogued, a localised pane got a sentence that
  // was translated up to the full stop and English after it.
  '{n} chart(s) were too dense for this host and went in as pictures — "Explode to native shapes" turns them back.':
    '{n} chart(s) were too dense for this host and went in as pictures — "Explode to native shapes" turns them back.',
  "{n} image chart(s) fell back to native shapes.": "{n} image chart(s) fell back to native shapes.",
  "Inserted {n} chart(s) on the current slide.": "Inserted {n} chart(s) on the current slide.",
  "Inserting demo slides… {done} of {total}": "Inserting demo slides… {done} of {total}",
  'Loaded chart 1 of {total} — use "Insert batch" for all.': 'Loaded chart 1 of {total} — use "Insert batch" for all.',
  "Working… {phase}": "Working… {phase}",
  "PowerPoint has not answered for {secs}s. Look at the slide area: if PowerPoint is showing *Sorry, we ran into a problem*, click Refresh there — nothing behind that dialog can answer. If a test run was in progress, its steps are saved and *Download the crashed run* will offer them.":
    "PowerPoint has not answered for {secs}s. Look at the slide area: if PowerPoint is showing *Sorry, we ran into a problem*, click Refresh there — nothing behind that dialog can answer. If a test run was in progress, its steps are saved and *Download the crashed run* will offer them.",
  // Insert-phase words (the {phase} fill above).
  "opening PowerPoint…": "opening PowerPoint…",
  "building shapes…": "building shapes…",
  "sending to PowerPoint…": "sending to PowerPoint…",
  "grouping…": "grouping…",
  done: "done",

  // Engine-injected chart label (threaded into the pure core via ChartConfig.labels.other).
  Other: "Other",
} as const;

/** Every translatable key. A language dictionary must cover all of these. */
export type StringKey = keyof typeof EN;

/**
 * Shipped languages. Empty today (English-only). Add `de: { …Record<StringKey> }`
 * to localize; the type makes an incomplete dictionary a compile error.
 */
const DICTS: Record<string, Record<StringKey, string>> = {};

let activeDict: Record<StringKey, string> | undefined;

/**
 * Register a language dictionary at runtime, keyed by its 2-letter code. A seam
 * for dynamically-loaded translations (and how the tests exercise the DOM sweep
 * without a language shipping). Still `Record<StringKey, string>`, so a caller
 * cannot register an incomplete dictionary.
 */
export function registerLanguage(code: string, dict: Record<StringKey, string>): void {
  DICTS[code.slice(0, 2).toLowerCase()] = dict;
}

// Static markup elements whose direct text (and input placeholders) get re-skinned.
const LOCALIZE_SELECTOR =
  "h2, button, label, .banner, option, .tagline, figcaption, summary, .acc-title, .group-label, .fgroup-name, .no-type-result";

/**
 * Own-property lookup. These dictionaries are plain objects and the key is
 * arbitrary text — DOM content, or a status string — so "toString" reached
 * Object.prototype and returned a FUNCTION: a template named `toString` rendered
 * as "[object Undefined]", a placeholder named `valueOf` became
 * "function valueOf() { [native code] }", and t() threw outright once it tried
 * to interpolate into one.
 */
const lookup = (dict: Record<string, string> | undefined, key: string): string | undefined =>
  dict && Object.prototype.hasOwnProperty.call(dict, key) && typeof dict[key] === "string" ? dict[key] : undefined;

function translateTree(root: ParentNode, dict: Record<string, string>): void {
  for (const el of root.querySelectorAll<HTMLElement>(LOCALIZE_SELECTOR)) {
    // Only translate an element's direct text, so child inputs/spans survive.
    for (const child of el.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        const key = child.textContent?.trim();
        const hit = key ? lookup(dict, key) : undefined;
        if (key && hit) child.textContent = child.textContent!.replace(key, hit);
      }
    }
  }
  for (const input of root.querySelectorAll<HTMLInputElement>("input[placeholder]")) {
    const p = input.placeholder.trim();
    const hit = lookup(dict, p);
    if (p && hit) input.placeholder = hit;
  }
}

/** Translate visible UI text in place; a no-op for English / unsupported languages. */
export function localizePane(language: string | undefined): void {
  activeDict = DICTS[(language ?? "").slice(0, 2).toLowerCase()];
  if (activeDict) translateTree(document, activeDict);
}

/** Re-apply the active translation to a freshly-rendered subtree (gallery, format groups). */
export function localizeTree(root: ParentNode): void {
  if (activeDict) translateTree(root, activeDict);
}

/**
 * Translate a runtime string built in code (a status message) and fill any
 * `{placeholder}` from `params`. Falls back to the English catalogue, then to the
 * raw key, so callers can wrap unconditionally and dynamic text passes through.
 */
export function t(key: string, params?: Record<string, string | number>): string {
  let s: string = lookup(activeDict, key) ?? lookup(EN as unknown as Record<string, string>, key) ?? key;
  if (params) for (const [k, v] of Object.entries(params)) s = s.split(`{${k}}`).join(String(v));
  return s;
}
