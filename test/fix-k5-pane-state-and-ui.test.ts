// @vitest-environment jsdom
import { afterAll, afterEach, describe, expect, it, vi } from "vitest";
import { readdirSync, readFileSync } from "fs";
import ts from "typescript";
import { PALETTES } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import { EN, localizePane, registerLanguage, t, type StringKey } from "../src/taskpane/i18n";

/**
 * THE THIRD BUG HUNT, STREAM k5: pane state that goes stale, keyboard and
 * screen-reader defects, and the localisation claims. One describe per finding.
 *
 * The pane tests boot the real `app.ts` against `taskpane.html` the way
 * `pane-state.test.ts` does. The writes that need PowerPoint are in
 * `fix-k5-pane-writes.test.ts`.
 */

/** See `pane-state.test.ts`: the pane's last debounced render must not outlive jsdom. */
afterAll(async () => {
  await new Promise((r) => setTimeout(r, 250));
});

afterEach(() => {
  vi.restoreAllMocks();
  localizePane(undefined);
});

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Boot the pane on a clean browser. */
async function bootPane() {
  try {
    window.localStorage.clear();
  } catch {
    /* a jsdom without storage */
  }
  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  const app = await import("../src/taskpane/app");
  await app.harnessReady;
}

const exportConfig = (): ChartConfig => {
  $("json-export").click();
  return JSON.parse(($("json-io") as HTMLTextAreaElement).value);
};
const cell = (r: number, c: number) =>
  document.querySelector<HTMLInputElement>(`#datasheet input[data-row="${r}"][data-col="${c}"]`)!;
const typeInto = (el: HTMLInputElement, value: string) => {
  el.value = value;
  el.dispatchEvent(new Event("input", { bubbles: true }));
};
const sheetButton = (label: string) =>
  [...document.querySelectorAll<HTMLButtonElement>("#datasheet .sheet-controls button")].find(
    (b) => b.textContent === label,
  )!;
/** Ctrl+Z / Ctrl+Y typed WHERE the person is (default: the page itself). */
const ctrl = (key: string, on: EventTarget = document) =>
  on.dispatchEvent(new KeyboardEvent("keydown", { key, ctrlKey: true, bubbles: true }));
/** An Options group by the name on its summary: "Labels", "Axes & scale", "Analysis", "Layout", "Colours & style". */
const group = (name: string) =>
  [...document.querySelectorAll<HTMLDetailsElement>("#options details.fgroup")].find(
    (d) => d.querySelector(".fgroup-name")?.textContent === name,
  )!;

/* ─────────────────────────── r3-23 / r3-32 ──────────────────────────────────── */

describe("the colour pickers follow the datasheet's series (r3-23, r3-32)", () => {
  const swatches = () => [...document.querySelectorAll<HTMLLabelElement>("#options .series-colors label")];
  const captions = () => swatches().map((l) => l.textContent);
  const seriesNames = () => exportConfig().data.series.map((s) => s.name);
  const coloursBySeries = () => exportConfig().data.series.map((s) => [s.name, s.color ?? null]);
  const pick = (caption: string, colour: string) => {
    const input = swatches()
      .find((l) => l.textContent === caption)!
      .querySelector("input")!;
    input.value = colour;
    input.dispatchEvent(new Event("input", { bubbles: true }));
  };

  it("starts level with the sheet", async () => {
    await bootPane();
    expect(captions()).toEqual(seriesNames());
    expect(captions()).toEqual(["Enterprise", "SMB", "Consumer"]);
  });

  it('"+ Row" adds a picker for the new series, and the one captioned SMB still paints SMB', async () => {
    await bootPane();
    sheetButton("+ Row").click();
    expect(seriesNames()).toEqual(["Enterprise", "Series 2", "SMB", "Consumer"]);
    expect(captions(), "the pickers did not follow the new row").toEqual(seriesNames());

    pick("SMB", "#00ff00");
    expect(coloursBySeries(), "the picker captioned SMB painted another series").toEqual([
      ["Enterprise", null],
      ["Series 2", null],
      ["SMB", "#00ff00"],
      ["Consumer", null],
    ]);
    pick("Series 2", "#0000ff");
    expect(coloursBySeries()[1]).toEqual(["Series 2", "#0000ff"]);
  });

  it('"− Row" takes the dead picker away, and the first one paints the first series that is left', async () => {
    await bootPane();
    sheetButton("− Row").click();
    expect(seriesNames()).toEqual(["SMB", "Consumer"]);
    expect(captions(), "a picker is left for a series that no longer exists").toEqual(["SMB", "Consumer"]);

    pick("SMB", "#00ff00");
    expect(coloursBySeries()).toEqual([
      ["SMB", "#00ff00"],
      ["Consumer", null],
    ]);
  });

  it("typing over a series name renames its picker", async () => {
    await bootPane();
    typeInto(cell(1, 0), "Sales");
    expect(seriesNames()).toEqual(["Sales", "SMB", "Consumer"]);
    expect(captions()).toEqual(["Sales", "SMB", "Consumer"]);
  });

  it("an undone row takes its picker with it", async () => {
    await bootPane();
    sheetButton("+ Row").click();
    expect(captions()).toHaveLength(4);
    ctrl("z");
    expect(captions()).toEqual(["Enterprise", "SMB", "Consumer"]);
    ctrl("y");
    expect(captions()).toEqual(["Enterprise", "Series 2", "SMB", "Consumer"]);
  });

  it("keeps the picker's own element, so a row added above it does not take its focus", async () => {
    await bootPane();
    const first = swatches()[0].querySelector("input")!;
    first.focus();
    sheetButton("+ Row").click();
    expect(swatches()[0].querySelector("input"), "the first picker was rebuilt").toBe(first);
    expect(first.isConnected).toBe(true);
  });

  it("shows the palette's colours after a Palette pick, not the old palette's", async () => {
    await bootPane();
    const select = [...document.querySelectorAll<HTMLSelectElement>("#options select")].find((s) =>
      [...s.options].some((o) => o.value === "Blues"),
    )!;
    const before = swatches().map((l) => l.querySelector("input")!.value);
    expect(before).toEqual(PALETTES.Default.slice(0, 3).map((c) => c.toLowerCase()));
    select.value = "Blues";
    select.dispatchEvent(new Event("change", { bubbles: true }));
    await sleep(150);
    expect(
      swatches().map((l) => l.querySelector("input")!.value),
      "the swatches stayed on the old palette while the chart drew the new one",
    ).toEqual(PALETTES.Blues.slice(0, 3).map((c) => c.toLowerCase()));
  });
});

/* ─────────────────────────────── r3-24 ───────────────────────────────────────── */

/**
 * THE [hidden] SWEEP. jsdom forces `[hidden]` to `display: none` whatever the
 * sheet says, so no DOM test can show the trap: an author `display: flex` beats
 * the user agent's `[hidden] { display: none }` and the attribute changes nothing
 * on screen. The rules and the markup are both static, so match them here.
 */
type CssRule = { sel: string; body: string; order: number };
const stripComments = (s: string) => s.replace(/\/\*[\s\S]*?\*\//g, "");
function flattenCss(text: string, out: CssRule[] = []): CssRule[] {
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
      if (head.startsWith("@media")) flattenCss(body, out);
    } else {
      let d = 0;
      let cur = "";
      const parts: string[] = [];
      for (const ch of head) {
        if (ch === "(") d++;
        else if (ch === ")") d--;
        if (ch === "," && d === 0) {
          parts.push(cur.trim());
          cur = "";
        } else cur += ch;
      }
      if (cur.trim()) parts.push(cur.trim());
      for (const sel of parts) out.push({ sel, body, order: out.length });
    }
    i = j;
  }
  return out;
}
const displayOf = (body: string): string | null => {
  let last: string | null = null;
  for (const d of body.split(";")) {
    const m = /^\s*display\s*:\s*([^!]+?)\s*(?:!important)?\s*$/.exec(d);
    if (m) last = m[1];
  }
  return last;
};
const specificityOf = (sel: string): number => {
  const s = sel.replace(/:where\([^)]*\)/g, "");
  const ids = (s.match(/#[\w-]+/g) || []).length;
  const classish =
    (s.match(/\.[\w-]+/g) || []).length +
    (s.match(/\[[^\]]*\]/g) || []).length +
    (s.match(/:(?!:)[\w-]+/g) || []).length;
  const types = (s.match(/(^|[\s>+~])[a-z]+/g) || []).length;
  return ids * 10000 + classish * 100 + types;
};
const matches = (el: Element, sel: string): boolean => {
  try {
    return el.matches(sel);
  } catch {
    return false;
  }
};

/**
 * Every element the pane hides that an author `display` rule would keep on screen:
 * a rule that sets `display` to anything but none, matches the element at rest,
 * and is not answered by a `[hidden]` rule that says `display: none` at least as
 * specifically and no earlier in the sheet.
 */
function hiddenButShown(cssText: string, doc: Document, toggled: string[]): string[] {
  const rules = flattenCss(stripComments(cssText));
  const REST = /:(hover|focus|focus-visible|active|disabled|checked|empty|first-child)|::/;
  const els = new Set<Element>([...doc.querySelectorAll("[hidden]")]);
  // The ones the code hides later are asked as they will be when it does.
  for (const sel of toggled)
    doc.querySelectorAll(sel).forEach((e) => {
      e.setAttribute("hidden", "");
      els.add(e);
    });
  const offenders: string[] = [];
  for (const el of els) {
    for (const r of rules) {
      const display = displayOf(r.body);
      if (!display || display === "none" || REST.test(r.sel) || !matches(el, r.sel)) continue;
      const answered = rules.some(
        (o) =>
          /\[hidden\]/.test(o.sel) &&
          displayOf(o.body) === "none" &&
          matches(el, o.sel) &&
          (specificityOf(o.sel) > specificityOf(r.sel) ||
            (specificityOf(o.sel) === specificityOf(r.sel) && o.order > r.order)),
      );
      if (!answered)
        offenders.push(`${el.tagName.toLowerCase()}${el.id ? `#${el.id}` : ""} by ${r.sel} {display: ${display}}`);
    }
  }
  return [...new Set(offenders)];
}

describe("an element the pane hides stays hidden (r3-24)", () => {
  const css = readFileSync("src/taskpane/taskpane.css", "utf8");
  const page = () => new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  /** Everything app.ts / status.ts / harness-ui.ts write `.hidden` or the attribute on. */
  const TOGGLED = [
    ".action-bar",
    "#actions-menu",
    "#template-name-row",
    "#slow-offer",
    "#slow-offer-picture",
    "#slow-offer-here",
    "#status-strip",
    "#status-bar",
    "#status-stop",
    "#testing-section",
    "#demo-crashlog",
  ];

  it("has a [hidden] answer for every element it hides", () => {
    expect(hiddenButShown(css, page(), TOGGLED)).toEqual([]);
  });

  it("is able to say so: the sheet without the template row's rule is rejected", () => {
    const broken = css.replace(/\.element-row\[hidden\]\s*\{[^}]*\}/, "");
    expect(broken).not.toBe(css);
    expect(hiddenButShown(broken, page(), TOGGLED).join("\n")).toMatch(/div#template-name-row by \.element-row/);
  });

  it("sees the other rules go too (the guard is not specific to one element)", () => {
    const broken = css.replace(/\.slow-offer\[hidden\]\s*\{[^}]*\}/, "");
    expect(hiddenButShown(broken, page(), TOGGLED).join("\n")).toMatch(/div#slow-offer by \.slow-offer/);
  });
});

/* ─────────────────────────────── r3-25 ───────────────────────────────────────── */

describe("the ⋯ menu can be used from the keyboard (r3-25)", () => {
  /**
   * What a browser does and jsdom does not: `focus()` on an element that is not
   * rendered, or is disabled, does nothing. Without it the old test passed on the
   * broken behaviour, because jsdom focuses a `display: none` button.
   */
  function browserFocusRule() {
    const real = HTMLElement.prototype.focus;
    vi.spyOn(HTMLElement.prototype, "focus").mockImplementation(function (this: HTMLElement) {
      const unfocusable =
        this.style.display === "none" ||
        this.hidden ||
        (this as HTMLButtonElement).disabled ||
        !!this.closest("[hidden], [style*='display: none']");
      if (unfocusable) return;
      real.call(this);
    });
  }
  const trigger = () => $<HTMLButtonElement>("more-actions");
  const menuItems = () => [...document.querySelectorAll<HTMLButtonElement>("#actions-menu button")];
  const arrow = (key: string) =>
    document.activeElement!.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
  const active = () => document.activeElement as HTMLElement;

  it("moves focus INTO the menu when it opens with 'Insert as new' hidden", async () => {
    await bootPane();
    browserFocusRule();
    expect($("insert-new").style.display, "set-up: no chart is open for editing").toBe("none");
    trigger().focus();
    trigger().click();
    expect(active().closest("#actions-menu"), "focus stayed on the ⋯ button").not.toBeNull();
    expect(active().id).toBe("download");
  });

  it("walks the items that can take focus, skipping hidden and disabled ones, both ways round", async () => {
    await bootPane();
    browserFocusRule();
    const usable = menuItems().filter((b) => !b.disabled && b.style.display !== "none");
    expect(
      usable.map((b) => b.id),
      "set-up: a browser tab has these three",
    ).toEqual(["download", "download-png", "copy-link"]);
    trigger().click();
    const walk: string[] = [active().id];
    for (let i = 0; i < 4; i++) {
      arrow("ArrowDown");
      walk.push(active().id);
    }
    expect(walk).toEqual(["download", "download-png", "copy-link", "download", "download-png"]);
    arrow("End");
    expect(active().id).toBe("copy-link");
    arrow("Home");
    expect(active().id).toBe("download");
    arrow("ArrowUp");
    expect(active().id).toBe("copy-link");
  });

  it("opens from the arrow keys on the button: Down onto the first item, Up onto the last", async () => {
    await bootPane();
    browserFocusRule();
    trigger().focus();
    trigger().dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }));
    expect($("actions-menu").hidden).toBe(false);
    expect(active().id).toBe("download");
    // Close it and try the other end.
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect($("actions-menu").hidden).toBe(true);
    trigger().dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true, cancelable: true }));
    expect(active().id).toBe("copy-link");
  });

  it("skips an item a running action has disabled in the host, too", async () => {
    await bootPane();
    browserFocusRule();
    // The host-connected pane: every item enabled but 'Insert as new'.
    for (const b of menuItems()) b.disabled = false;
    $<HTMLButtonElement>("load-selection").disabled = true; // a guard()ed click holds it
    trigger().click();
    expect(active().id).toBe("explode");
    arrow("ArrowUp");
    expect(active().id, "ArrowUp walked onto the disabled item and stopped").toBe("copy-link");
  });
});

/* ─────────────────────────────── r3-26 ───────────────────────────────────────── */

describe("the Excel pane can be read by a screen reader (r3-26)", () => {
  const page = () => new DOMParser().parseFromString(readFileSync("src/excel/excel.html", "utf8"), "text/html");
  /** The accessible name of a control, by the rules a screen reader uses (see `pane-host-actions.test.ts`). */
  const accessibleName = (doc: Document, el: Element): string => {
    const aria = el.getAttribute("aria-label");
    if (aria?.trim()) return aria.trim();
    const by = el.getAttribute("aria-labelledby");
    if (by) {
      const ref = by
        .split(/\s+/)
        .map((id) => doc.getElementById(id)?.textContent ?? "")
        .join(" ")
        .trim();
      if (ref) return ref;
    }
    const title = el.getAttribute("title");
    if (title?.trim()) return title.trim();
    const forLabel = el.id ? doc.querySelector(`label[for="${el.id}"]`) : null;
    if (forLabel?.textContent?.trim()) return forLabel.textContent.trim();
    const wrapping = el.closest("label");
    if (wrapping) {
      const first = wrapping.querySelector("input, select, textarea, button");
      if (first === el && wrapping.textContent?.trim()) return wrapping.textContent.trim();
    }
    if (el instanceof (doc.defaultView ?? window).HTMLButtonElement) return el.textContent?.trim() ?? "";
    return "";
  };

  it("names every control on the page, the JSON box included", () => {
    const doc = page();
    const controls = [...doc.querySelectorAll("input, select, textarea, button")];
    expect(controls.length, "the page was not parsed").toBeGreaterThan(4);
    const unnamed = controls
      .filter((el) => !accessibleName(doc, el))
      .map((el) => `${el.tagName.toLowerCase()}#${el.id}`);
    expect(unnamed).toEqual([]);
    expect(accessibleName(doc, doc.getElementById("output")!)).toBe("Chart configuration JSON");
  });

  it("announces what it says: the status line is a live region", () => {
    const note = page().getElementById("note")!;
    expect(note.getAttribute("role")).toBe("status");
    expect(note.getAttribute("aria-live")).toBe("polite");
  });
});

/* ─────────────────────────────── r3-66 ───────────────────────────────────────── */

describe("rebuilding a panel does not take the person's place with it (r3-66)", () => {
  const datamarkBox = () =>
    [...document.querySelectorAll<HTMLLabelElement>("#options label")]
      .find((l) => (l.textContent ?? "").startsWith("Datamark axis"))!
      .querySelector("input")!;
  const valueAxisBox = () =>
    [...document.querySelectorAll<HTMLLabelElement>("#options label")]
      .find((l) => l.textContent === "Value axis")!
      .querySelector("input")!;
  const tick = (box: HTMLInputElement, on: boolean) => {
    box.checked = on;
    box.dispatchEvent(new Event("change", { bubbles: true }));
  };

  it("ticking Datamark axis keeps the Axes group open and the box focused, and sets Value axis beside it", async () => {
    await bootPane();
    group("Axes & scale").open = true;
    const box = datamarkBox();
    box.focus();
    expect(document.activeElement).toBe(box);

    tick(box, true);
    expect(group("Axes & scale").open, "the group the person was in snapped shut").toBe(true);
    expect(box.isConnected, "the checkbox they just ticked was destroyed").toBe(true);
    expect(document.activeElement, "keyboard focus fell to the page").toBe(box);
    expect(valueAxisBox().checked, "the Value axis box did not follow the datamark axis").toBe(true);

    tick(box, false);
    expect(group("Axes & scale").open).toBe(true);
    expect(document.activeElement).toBe(box);
    expect(valueAxisBox().checked).toBe(false);
  });

  it("keeps a group open across the rebuild a chart type pick makes", async () => {
    await bootPane();
    expect(group("Labels").open, "set-up: Labels opens by default").toBe(true);
    group("Axes & scale").open = true;
    group("Labels").open = false;
    $("gallery").querySelector<HTMLButtonElement>('.thumb[data-kind="clustered"]')!.click();
    expect(group("Axes & scale").open, "Axes & scale was shut by the rebuild").toBe(true);
    expect(group("Labels").open, "a group the person closed was opened by the rebuild").toBe(false);
    expect(group("Analysis").open).toBe(false);
  });

  it("returns focus to the same cell after a multi-cell paste", async () => {
    await bootPane();
    const target = cell(1, 1);
    target.focus();
    const paste = new Event("paste", { bubbles: true, cancelable: true }) as Event & { clipboardData: unknown };
    paste.clipboardData = { getData: () => "5\t6\n7\t8" };
    target.dispatchEvent(paste);
    expect(cell(1, 1).value, "set-up: the paste landed").toBe("5");
    expect(target.isConnected, "set-up: the grid was rebuilt").toBe(false);
    expect(document.activeElement).toBe(cell(1, 1));
  });

  it("returns focus to the button that was pressed, and to the cell Ctrl+Z was typed in", async () => {
    await bootPane();
    const add = sheetButton("+ Row");
    add.focus();
    add.click();
    const again = sheetButton("+ Row");
    expect(again, "set-up: the controls were rebuilt").not.toBe(add);
    expect(document.activeElement, "the button was pressed from the keyboard and focus fell to the page").toBe(again);

    typeInto(cell(1, 1), "99");
    cell(1, 1).focus();
    ctrl("z", cell(1, 1));
    expect(cell(1, 1).value, "set-up: the edit was undone").not.toBe("99");
    expect(document.activeElement, "Ctrl+Z dropped focus out of the grid").toBe(cell(1, 1));
  });

  it("puts focus on the Chart type summary when a tile was picked from the keyboard", async () => {
    await bootPane();
    const tile = $("gallery").querySelector<HTMLButtonElement>('.thumb[data-kind="pie"]')!;
    tile.focus();
    tile.click();
    expect($<HTMLDetailsElement>("type-acc").open, "set-up: the picker folds away").toBe(false);
    expect(tile.isConnected, "set-up: the tile was rebuilt").toBe(false);
    expect(document.activeElement).toBe(document.querySelector("#type-acc > summary"));
  });
});

/* ─────────────────────────────── r3-67 ───────────────────────────────────────── */

describe("a data cell is named from the headers as they are NOW (r3-67)", () => {
  const name = (r: number, c: number) => cell(r, c).getAttribute("aria-label");

  it("renames the cells under a series the moment its name is typed", async () => {
    await bootPane();
    expect(name(1, 1)).toBe("Enterprise, 2022");
    typeInto(cell(1, 0), "Sales");
    expect(name(1, 1)).toBe("Sales, 2022");
    expect(name(1, 2)).toBe("Sales, 2023");
    expect(name(2, 1), "another series' cell was renamed").toBe("SMB, 2022");
  });

  it("renames the cells under a category the moment its heading is typed", async () => {
    await bootPane();
    typeInto(cell(0, 1), "FY99");
    expect(name(1, 1)).toBe("Enterprise, FY99");
    expect(name(2, 1)).toBe("SMB, FY99");
    expect(name(1, 2), "another column's cell was renamed").toBe("Enterprise, 2023");
  });

  it("keeps the header cells named by their axis, and falls back when a heading is cleared", async () => {
    await bootPane();
    typeInto(cell(1, 0), "Sales");
    expect(name(1, 0)).toBe("Row 1 heading");
    typeInto(cell(1, 0), "");
    expect(name(1, 1)).toBe("Row 1, 2022");
    typeInto(cell(0, 1), "");
    expect(name(1, 1)).toBe("Row 1, column 1");
  });
});

/* ─────────────────────────────── r3-68 ───────────────────────────────────────── */

describe("the Stop button reads when the pointer is on it, in both schemes (r3-68)", () => {
  const css = stripComments(readFileSync("src/taskpane/taskpane.css", "utf8"));
  const rules = (() => {
    // The dark block is the one @media that changes the answer.
    const dark = /@media\s*\(prefers-color-scheme:\s*dark\)\s*\{/.exec(css);
    const lightText = dark
      ? css.slice(0, dark.index) + css.slice(matchingBrace(css, dark.index + dark[0].length))
      : css;
    const darkText = dark
      ? css.slice(dark.index + dark[0].length, matchingBrace(css, dark.index + dark[0].length) - 1)
      : "";
    return { light: flattenCss(lightText), dark: flattenCss(darkText) };
  })();
  function matchingBrace(text: string, from: number): number {
    let depth = 1;
    let j = from;
    while (j < text.length && depth > 0) {
      if (text[j] === "{") depth++;
      else if (text[j] === "}") depth--;
      j++;
    }
    return j;
  }
  const declared = (body: string, prop: string): string | null => {
    let last: string | null = null;
    for (const d of body.split(";")) {
      const i = d.indexOf(":");
      if (i > 0 && d.slice(0, i).trim() === prop) last = d.slice(i + 1).trim();
    }
    return last;
  };
  const tokens = (scheme: "light" | "dark"): Record<string, string> => {
    const out: Record<string, string> = {};
    for (const r of [...rules.light, ...(scheme === "dark" ? rules.dark : [])])
      if (r.sel === ":root")
        for (const d of r.body.split(";")) {
          const i = d.indexOf(":");
          if (i > 0 && d.slice(0, i).trim().startsWith("--")) out[d.slice(0, i).trim()] = d.slice(i + 1).trim();
        }
    return out;
  };
  const resolve = (value: string, scheme: "light" | "dark"): string => {
    const m = value.match(/var\((--[\w-]+)(?:\s*,\s*([^)]+))?\)/);
    return m ? (tokens(scheme)[m[1]] ?? m[2] ?? "").trim() : value;
  };
  const lum = (hex: string): number => {
    const h = hex.replace("#", "");
    const c = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
    const lin = c.map((s) => (s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4)));
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2];
  };
  const contrast = (a: string, b: string) => {
    const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p);
    return (x + 0.05) / (y + 0.05);
  };

  it.each(["light", "dark"] as const)("hover text clears 4.5:1 on the button's own fill (%s)", (scheme) => {
    const all = [...rules.light, ...(scheme === "dark" ? rules.dark : [])];
    const hover = all.filter((r) => r.sel === ".status-stop:hover:not(:disabled)").pop()!;
    const rest = all.filter((r) => r.sel === ".status-stop").pop()!;
    const text = resolve(declared(hover.body, "color")!, scheme);
    const fill = resolve(declared(rest.body, "background")!, scheme);
    expect(text, "the hover colour did not resolve").toMatch(/^#[0-9a-f]{6}$/i);
    expect(contrast(text, fill), `${text} on ${fill}`).toBeGreaterThanOrEqual(4.5);
  });

  it("uses no custom property that nothing declares and no fixed colour in place of one", () => {
    const declaredProps = new Set([...css.matchAll(/(--[\w-]+)\s*:/g)].map((m) => m[1]));
    const bad: string[] = [];
    for (const m of css.matchAll(/var\(\s*(--[\w-]+)\s*(?:,\s*([^)]*(?:\([^)]*\))?[^)]*))?\)/g)) {
      if (declaredProps.has(m[1])) continue;
      // An undeclared name falls to its fallback in BOTH schemes, so a fallback that
      // is a literal colour is the light theme's colour in the dark one.
      if (m[2] === undefined || /#[0-9a-f]{3,8}\b|\brgba?\(/i.test(m[2])) bad.push(m[0]);
    }
    expect(bad).toEqual([]);
  });
});

/* ─────────────────────────────── r3-64 ───────────────────────────────────────── */

describe("the catalogue holds every string the pane says (r3-64)", () => {
  /** The Testing panel and the in-host battery are developer tools: English by design, not for a translator. */
  const DEVELOPER_ONLY = new Set(["harness-ui.ts", "selftest.ts", "crashlog.ts"]);

  type Call = { file: string; line: number; fn: string; keys: string[] | null; template: boolean };
  function noteAndTCalls(): Call[] {
    const dir = "src/taskpane";
    const found: Call[] = [];
    for (const file of readdirSync(dir).filter((f) => f.endsWith(".ts") && !DEVELOPER_ONLY.has(f))) {
      const sf = ts.createSourceFile(file, readFileSync(`${dir}/${file}`, "utf8"), ts.ScriptTarget.Latest, true);
      /** Every literal text the expression can be: a string, a concatenation, either arm of a conditional. */
      const texts = (e: ts.Expression): string[] | null => {
        if (ts.isStringLiteral(e) || ts.isNoSubstitutionTemplateLiteral(e)) return [e.text];
        if (ts.isParenthesizedExpression(e)) return texts(e.expression);
        if (ts.isBinaryExpression(e) && e.operatorToken.kind === ts.SyntaxKind.PlusToken) {
          const [l, r] = [texts(e.left), texts(e.right)];
          return l && r ? l.flatMap((a) => r.map((b) => a + b)) : null;
        }
        if (ts.isConditionalExpression(e)) {
          const [a, b] = [texts(e.whenTrue), texts(e.whenFalse)];
          return a && b ? [...a, ...b] : null;
        }
        return null;
      };
      const visit = (n: ts.Node): void => {
        if (ts.isCallExpression(n) && ts.isIdentifier(n.expression) && ["t", "note"].includes(n.expression.text)) {
          const arg = n.arguments[0];
          if (arg)
            found.push({
              file,
              line: sf.getLineAndCharacterOfPosition(n.getStart()).line + 1,
              fn: n.expression.text,
              keys: texts(arg),
              template: ts.isTemplateExpression(arg) || (ts.isBinaryExpression(arg) && /`/.test(arg.getText())),
            });
        }
        ts.forEachChild(n, visit);
      };
      visit(sf);
    }
    return found;
  }

  it("has an English entry for every literal t() and note() text outside the developer tools", () => {
    const missing = noteAndTCalls()
      .flatMap((c) => (c.keys ?? []).map((k) => ({ ...c, k })))
      .filter((c) => c.k !== "" && !Object.prototype.hasOwnProperty.call(EN, c.k))
      .map((c) => `${c.file}:${c.line} ${c.fn}(${JSON.stringify(c.k.slice(0, 70))})`);
    expect(missing).toEqual([]);
  });

  it("builds no catalogue key by interpolation (a key made at run time cannot be in the catalogue)", () => {
    const built = noteAndTCalls()
      .filter((c) => c.template)
      .map((c) => `${c.file}:${c.line} ${c.fn}(\`…\`)`);
    expect(built).toEqual([]);
  });

  it("has no entry no caller can reach: the batch's old count line was replaced, not kept", () => {
    expect(Object.keys(EN)).not.toContain("Inserted {n} chart(s) on the current slide.");
    expect(Object.keys(EN)).toContain("Inserted {n} of {total} chart(s).");
    expect(Object.keys(EN)).toContain("Inserted {n} chart(s), one per slide.");
  });

  it("lets a language translate those strings: a dictionary covering EN reaches them", () => {
    const dict = Object.fromEntries((Object.keys(EN) as StringKey[]).map((k) => [k, `DE<${k}>`])) as Record<
      StringKey,
      string
    >;
    registerLanguage("de", dict);
    localizePane("de-DE");
    expect(t("Nothing inserted.")).toBe("DE<Nothing inserted.>");
    expect(t("Inserted {n} of {total} chart(s).", { n: 3, total: 4 })).toBe("DE<Inserted 3 of 4 chart(s).>");
    expect(t("Skipped {n} it could not read — {which}.", { n: 1, which: "x" })).toBe(
      "DE<Skipped 1 it could not read — x.>",
    );
  });

  it("does not claim more than it does: chart output is English whatever the locale", () => {
    const header = readFileSync("src/taskpane/i18n.ts", "utf8").split("*/")[0];
    // The claim that was made: "its numbers and dates localize via NumberFormat.locale".
    expect(header).not.toMatch(/dates\s+localize/);
    expect(header).not.toMatch(/every user-facing string is catalogued/i);
    // And what is now said instead.
    expect(header).toMatch(/month/i);
    expect(header).toMatch(/alt text|description/i);
  });
});
