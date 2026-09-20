// @vitest-environment jsdom
import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { BUILTIN_TEMPLATES } from "../src/taskpane/templates";
import type { ChartConfig } from "../src/core/types";

/**
 * Saved chart templates — a whole user-facing feature with no tests at all.
 *
 * Save the chart you have set up, pick it again later, delete it when you are
 * done. It rides on `localStorage` and on a plain object keyed by whatever the
 * user typed, which is the combination this repo has been bitten by before:
 * "object lookups keyed by a config string must use
 * `Object.prototype.hasOwnProperty.call`" is in the project's own notes, with a
 * list of the tables it was applied to. This table was not on the list.
 */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const TEMPLATES_KEY = "ssf-charts-templates";
const LEGACY_TEMPLATES_KEY = "powerchart-templates";

async function bootPane() {
  window.localStorage.clear();
  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  await import("../src/taskpane/app");
}

/** Re-open the pane WITHOUT clearing storage — what a reload really is. */
async function reopenPane() {
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  await import("../src/taskpane/app");
}

/**
 * Save the pane's current chart under `name`, the way a user does.
 *
 * WAS A `prompt()` SPY. The pane called `prompt()` until 2026-09-20, which an
 * add-in must not do — a native modal blocks the Office HOST, not just the
 * pane — so the name now comes from an inline row that Save reveals. Driving
 * that row is closer to the real thing than stubbing a global ever was: this
 * goes through the same input, the same button and the same handler the user
 * does, and it would notice if the row stopped appearing.
 *
 * `async` because the handler awaits the answer now.
 */
async function saveAs(name: string) {
  $("template-save").click();
  const row = $("template-name-row");
  expect(row.hidden, "Save did not reveal the name row").toBe(false);
  $<HTMLInputElement>("template-name").value = name;
  $("template-name-ok").click();
  // Let the promise that click resolves settle before anything reads storage.
  await Promise.resolve();
}

/** The user-template names the picker is offering. */
const offered = () =>
  [...$<HTMLSelectElement>("template-list").querySelectorAll("option")]
    .map((o) => o.value)
    .filter((v) => v.startsWith("user:"))
    .map((v) => v.slice("user:".length));

const pick = (value: string) => {
  const sel = $<HTMLSelectElement>("template-list");
  sel.value = value;
  sel.dispatchEvent(new Event("change"));
};

const stored = (): Record<string, unknown> => JSON.parse(window.localStorage.getItem(TEMPLATES_KEY) ?? "{}");

beforeEach(async () => {
  await bootPane();
});
afterEach(() => {
  vi.restoreAllMocks();
  window.localStorage.clear();
});

describe("saved chart templates", () => {
  it("still finds templates saved under the pre-rename key", async () => {
    // THE RENAME COULD HAVE EATEN SOMEONE'S WORK. These keys moved from
    // `powerchart-*` to `ssf-charts-*` on 2026-08-27, and this one holds saved
    // chart templates — not a preference that can be set again, but work the
    // user did. A plain rename would have come up looking correct and emptied,
    // silently, on the first load after an update.
    //
    // Read-through, not a migration write: the old value is used when the new
    // key is absent, and the next save writes the new key. Nothing is deleted,
    // so a downgrade still finds its data.
    window.localStorage.removeItem(TEMPLATES_KEY);
    window.localStorage.setItem(LEGACY_TEMPLATES_KEY, JSON.stringify({ "from before": { kind: "waterfall" } }));
    await reopenPane();
    expect(offered(), "a template saved under the old key was lost by the rename").toContain("from before");
  });
  it("saves the chart you have, offers it back, and loads it", async () => {
    ($("chart-title") as HTMLInputElement).value = "Q3 revenue";
    ($("chart-title") as HTMLInputElement).dispatchEvent(new Event("input"));
    ($("chart-w") as HTMLInputElement).value = "640";
    ($("chart-w") as HTMLInputElement).dispatchEvent(new Event("input"));
    await saveAs("my layout");

    expect(offered(), "the saved template was not offered back").toContain("my layout");
    expect(Object.keys(stored())).toContain("my layout");

    // Change the pane, then pick the template — it must come back.
    ($("chart-w") as HTMLInputElement).value = "300";
    ($("chart-w") as HTMLInputElement).dispatchEvent(new Event("input"));
    pick("user:my layout");
    expect(($("chart-w") as HTMLInputElement).value, "picking a template did not restore its size").toBe("640");
  });

  it("survives a reload — the point of saving one", async () => {
    await saveAs("keeps");
    await reopenPane();
    expect(offered(), "the template did not survive reopening the pane").toContain("keeps");
  });

  it("deletes a user template, and refuses to delete a starter", async () => {
    await saveAs("throwaway");
    pick("user:throwaway");
    $("template-delete").click();
    expect(offered()).not.toContain("throwaway");
    expect(Object.keys(stored())).not.toContain("throwaway");

    // A starter is not the user's to delete, and the guard is a string prefix
    // — the kind that stops working the moment someone renames the option
    // values, so it is pinned.
    const starter = BUILTIN_TEMPLATES[0];
    pick(`builtin:${starter.name}`);
    $("template-delete").click();
    const still = [...$<HTMLSelectElement>("template-list").querySelectorAll("option")].map((o) => o.value);
    expect(still, "a built-in starter was deleted").toContain(`builtin:${starter.name}`);
  });

  it("loads a built-in starter", () => {
    const starter = BUILTIN_TEMPLATES.find((t) => (t.config as ChartConfig).kind);
    expect(starter, "no starter carries a kind to check against").toBeTruthy();
    pick(`builtin:${starter!.name}`);
    // The pane's own type summary is the visible proof it took.
    expect($("type-sub").textContent, "picking a starter changed nothing").toBeTruthy();
  });

  it("keeps a template whose name collides with a JavaScript builtin", async () => {
    // `all[name] = config` on an object from `JSON.parse` is a plain assignment
    // for every name but one. For `__proto__` it hits the inherited SETTER and
    // re-parents the object instead of storing anything — so the template reads
    // back correctly for the rest of the session, `JSON.stringify` writes it
    // out as `{}`, and it is gone the next time the pane opens. Saved,
    // apparently fine, silently lost.
    //
    // Nobody names a template `__proto__` on purpose. That is not the point:
    // the point is that this is the third table in this repo to be keyed by a
    // user-supplied string, and the project's own notes say to guard every one.
    await saveAs("__proto__");
    expect(offered(), "the odd name was not even offered in the same session").toContain("__proto__");

    await reopenPane();
    expect(offered(), "a saved template vanished on reload, with nothing said").toContain("__proto__");
  });

  it("cannot be tricked into loading Object.prototype's members as a chart", () => {
    // The read side of the same table. `loadTemplates()[name]` for a name that
    // is not stored — `constructor`, `toString` — reaches Object.prototype and
    // hands back a FUNCTION, which is truthy, so the pane would apply it as a
    // config. Nothing offers those names today because the picker is built
    // from `Object.keys`, which is exactly the sort of accident that stops
    // being true when someone changes how the list is built.
    window.localStorage.setItem(TEMPLATES_KEY, JSON.stringify({ real: { kind: "clustered" } }));
    // A width that is NOT the default, so "applied the wrong thing" and
    // "applied nothing" cannot look the same. They did in the first version of
    // this test, which read the default back and called it a pass.
    ($("chart-w") as HTMLInputElement).value = "640";
    ($("chart-w") as HTMLInputElement).dispatchEvent(new Event("input"));
    const before = ($("chart-w") as HTMLInputElement).value;
    expect(before).toBe("640");
    // Reach past the picker and ask for the dangerous name directly.
    const sel = $<HTMLSelectElement>("template-list");
    const opt = document.createElement("option");
    opt.value = "user:constructor";
    sel.appendChild(opt);
    pick("user:constructor");
    expect(($("chart-w") as HTMLInputElement).value, "applied something that was not a template").toBe(before);
  });
});
