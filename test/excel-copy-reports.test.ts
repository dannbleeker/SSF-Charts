// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from "vitest";

/**
 * THE COPY BUTTON SAYS WHETHER IT COPIED.
 *
 * Found 2026-09-29 by widening `no-silent-write-failure.test.ts`'s empty-catch
 * ban from `src/taskpane/app.ts` to all of `src/`. The Excel bridge's copy
 * handler read:
 *
 *     if (out) void navigator.clipboard?.writeText(out).catch(() => {});
 *
 * Two silences in one line. `navigator.clipboard.writeText` rejects routinely —
 * a document that is not focused, a permission the browser withholds, an add-in
 * frame without clipboard-write — and the user, who pressed Copy and is looking
 * at a textarea full of JSON they now believe is on the clipboard, was told
 * nothing and would paste whatever was there before.
 *
 * The `?.` hid the second: where `navigator.clipboard` is absent entirely,
 * optional chaining short-circuits the whole expression, so the `.catch` never
 * ran either and an unsupported browser also failed in silence.
 *
 * `generate()` in the same file already put both of its outcomes in `#note`.
 * This pins that the copy button does too.
 */
function pane(): void {
  document.body.innerHTML = `
      <select id="kind"><option value="stacked">stacked</option></select>
      <input id="title" value="">
      <button id="generate"></button>
      <button id="copy"></button>
      <textarea id="output">{"kind":"stacked"}</textarea>
      <p id="note"></p>`;
}

/** Let the handler's `.then`/`.catch` microtasks run before reading `#note`. */
async function settle(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

const note = (): string => document.getElementById("note")!.textContent ?? "";

describe("the Excel bridge's copy button reports its outcome", () => {
  beforeEach(() => {
    pane();
    vi.resetModules();
  });

  it("says it copied", async () => {
    Object.defineProperty(navigator, "clipboard", {
      value: { writeText: vi.fn(async () => {}) },
      configurable: true,
    });
    await import("../src/excel/excel");

    document.getElementById("copy")!.click();
    await settle();
    // Without the fix this is still the sideload message: a successful copy said
    // nothing, so the user could not tell it from a failed one.
    expect(note(), "a successful copy said nothing at all").toContain("Copied");
  });

  it("names a refusal and tells the user what to do instead", async () => {
    Object.defineProperty(navigator, "clipboard", {
      value: {
        writeText: vi.fn(async () => {
          throw new Error("Write permission denied.");
        }),
      },
      configurable: true,
    });
    await import("../src/excel/excel");

    document.getElementById("copy")!.click();
    await settle();
    // BOTH HALVES. The message has to carry the host's own reason — a bare
    // "could not copy" is the same non-answer as silence when the user is trying
    // to work out whether to grant a permission — and it has to name the way
    // out, because selecting the textarea by hand is a thing they can do.
    expect(note(), "the failure did not name its cause").toContain("Write permission denied.");
    expect(note(), "the failure did not name the way out").toMatch(/copy it by hand/i);
  });

  it("says so when the browser has no clipboard API at all", async () => {
    // THE `?.` CASE, which the old catch could not reach: optional chaining
    // short-circuits the whole chain, so `.catch` was never attached and nothing
    // ran. This is the one of the three that no catch body could have fixed.
    Object.defineProperty(navigator, "clipboard", { value: undefined, configurable: true });
    await import("../src/excel/excel");

    document.getElementById("copy")!.click();
    await settle();
    expect(note(), "an unsupported browser failed in silence").toMatch(/will not let an add-in write the clipboard/i);
  });

  it("still does nothing at all when there is nothing to copy", async () => {
    // The one silence that is correct, and it stays: an empty output box means
    // the user pressed Copy before pressing Generate, and a note about it would
    // be noise over a button that visibly did nothing.
    const writeText = vi.fn(async () => {});
    Object.defineProperty(navigator, "clipboard", { value: { writeText }, configurable: true });
    await import("../src/excel/excel");

    (document.getElementById("output") as HTMLTextAreaElement).value = "";
    const before = note();
    document.getElementById("copy")!.click();
    await settle();
    expect(writeText).not.toHaveBeenCalled();
    expect(note()).toBe(before);
  });
});
