import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

/**
 * THE FOUR STEPS A MICROSOFT REVIEWER IS TOLD TO CLICK.
 *
 * `docs/STORE-LISTING.md` carries a **Notes for certification** draft — the
 * field where you tell a stranger how to exercise the add-in. It walks them
 * through a ribbon path, a button, a banner and a section heading, naming each
 * in the exact words the product uses.
 *
 * Every one of those words was sourced from the code when the draft was
 * written, and the draft says so. That is the problem. `manifest.test.ts` and
 * `support-url-page.test.ts` both exist because a document that was accurate
 * when written drifted underneath a product that kept moving, and nothing
 * noticed. This note has the sharpest version of that failure mode: the reader
 * is a certification reviewer with no other picture of what the add-in does, on
 * a deadline, deciding whether to trust the rest of the submission. A step that
 * names a button which has since been relabelled does not just waste their
 * time — it makes every other claim look unchecked.
 *
 * The draft itself ends "Walk the four steps once on the web before pasting
 * it". That walk is a one-time act by a person. THIS keeps holding afterwards.
 *
 * WHAT IS DELIBERATELY NOT HERE. The 25 kinds and the Area / Tile map picture
 * behaviour are pinned by `web-density-census.test.ts`, which RE-DERIVES them
 * from `DEMO_SHAPE_BUDGET` rather than restating a number. Asserting them again
 * here would be a second copy to keep in step, which is the defect, not a check.
 */

const ROOT = resolve(__dirname, "..");
const text = (name: string) => readFileSync(resolve(ROOT, name), "utf8");

const NOTES = text("docs/STORE-LISTING.md");
const MANIFEST = text("manifest.xml");
const PANE = text("src/taskpane/taskpane.html");
const APP = text("src/taskpane/app.ts");

/**
 * The draft lives inside a blockquote in the checklist. Pulling it out means
 * these assertions are about the reviewer-facing text and not about some other
 * passage in a 284-line document that happens to use the same words.
 */
const draft = (() => {
  const start = NOTES.indexOf("SSF Charts needs no account");
  expect(
    start,
    "the notes-for-certification draft is gone from docs/STORE-LISTING.md — re-point this file",
  ).toBeGreaterThan(-1);
  const end = NOTES.indexOf("Every claim in that draft is sourced", start);
  const raw = NOTES.slice(start, end > start ? end : undefined);
  /**
   * Compare PROSE, not layout. The draft is markdown inside a blockquote, so
   * its sentences arrive wrapped across lines behind `>` prefixes with `**` and
   * backticks sitting between the words a reviewer reads as one phrase —
   * "Press **Insert into slide**" holds the button's name with markers inside
   * it. Partner Center gets the rendered text, so that is what to assert on.
   * Matching the raw source instead would fail on a reflow that changed
   * nothing, and the noise would teach whoever hit it to loosen the assertion.
   */
  return raw
    .split("\n")
    .map((l) => l.replace(/^\s*>?\s?/, ""))
    .join(" ")
    .split(/[*`_]/)
    .join("")
    .split(/\s+/)
    .join(" ");
})();

describe("the notes for certification still describe this product", () => {
  /**
   * STEP 1 — the way in. A reviewer who cannot find the ribbon entry never
   * reaches anything else, and the manifest is the only thing that decides
   * what it is called.
   */
  describe("step 1: the ribbon path", () => {
    it.each([
      ["the group label", "Group.Label"],
      ["the button that opens the pane", "OpenPane.Label"],
    ])("names %s exactly as the manifest sets it", (_what, id) => {
      const label = new RegExp(`id="${id}" DefaultValue="([^"]+)"`).exec(MANIFEST)?.[1];
      expect(label, `${id} is gone from manifest.xml`).toBeTruthy();
      expect(
        draft.includes(label as string),
        `the notes send a reviewer to "${label}" by a name the manifest no longer uses. ` +
          `A first step that does not work makes the other three unreadable.`,
      ).toBe(true);
    });
  });

  /**
   * STEP 2 — "Press Insert into slide", and what lands. The group name is what
   * the note tells them to look for in PowerPoint's own Selection pane, so it
   * is a claim about a string a user reads, not an internal identifier.
   */
  describe("step 2: insert, and what appears on the slide", () => {
    it("presses a button the pane actually has", () => {
      const label = /<button id="insert"[^>]*>([^<]+)</.exec(PANE)?.[1]?.trim();
      expect(label, "#insert is gone from taskpane.html").toBeTruthy();
      expect(draft).toContain(label as string);
    });

    it("promises the group name PowerPoint will really show", () => {
      // Both renderers have to agree, or the name depends on which path drew it.
      const names = ["src/render/ooxml.ts", "src/render/powerpoint.ts"].map(
        (f) => /const GROUP_NAME = "([^"]+)"/.exec(text(f))?.[1],
      );
      expect(new Set(names).size, `the two renderers group under different names: ${names.join(" vs ")}`).toBe(1);
      expect(draft).toContain(names[0] as string);
    });

    it("claims native shapes, which is what the pane's own copy claims", () => {
      expect(draft).toMatch(/native PowerPoint shapes/i);
    });
  });

  /**
   * STEP 3 — the round trip that distinguishes this from a picture generator,
   * and the one step whose wording is quoted rather than paraphrased.
   */
  describe("step 3: select the chart and edit it in place", () => {
    it("quotes the selection banner verbatim", () => {
      const banner = /<div id="selection-banner"[^>]*>([\s\S]*?)<button/.exec(PANE)?.[1];
      expect(banner, "#selection-banner is gone from taskpane.html").toBeTruthy();
      const sentence = (banner as string).replace(/\s+/g, " ").trim().replace(/\.$/, "");
      expect(
        draft.includes(sentence),
        `the notes quote a banner this pane does not show. It reads: "${sentence}". ` +
          `A quotation is the one kind of claim a reviewer can check at a glance.`,
      ).toBe(true);
    });

    it("names the button on that banner", () => {
      const button = /<button id="selection-banner-load"[^>]*>([^<]+)</.exec(PANE)?.[1]?.trim();
      expect(button, "#selection-banner-load is gone from taskpane.html").toBeTruthy();
      expect(draft).toContain(button as string);
    });

    it("promises an update in place rather than a second chart", () => {
      expect(draft).toMatch(/updates in place rather than adding a second one/i);
    });
  });

  /**
   * STEP 4 — the section the reviewer opens to see the range of the product.
   * The numbered heading is user-visible text and has been renumbered before.
   */
  describe("step 4: the chart type section", () => {
    it("names the section heading the pane renders", () => {
      const heading = /<span class="acc-title">([^<]*Chart type[^<]*)<\/span>/.exec(PANE)?.[1]?.trim();
      expect(heading, "no 'Chart type' accordion heading in taskpane.html").toBeTruthy();
      expect(draft).toContain(heading as string);
    });
  });

  /**
   * THE FIRST SENTENCE, and the one a reviewer acts on before clicking
   * anything: no account, no key, no sign-in. It is also a claim the code can
   * falsify, the way `privacy-terms-pages.test.ts` falsifies the network one.
   */
  describe("the no-credentials promise", () => {
    it("is what the notes lead with", () => {
      expect(draft).toMatch(/no account, licence key, sign-in or demo credentials/i);
    });

    it("is true of the shipped pane", () => {
      const asks = /\b(type=["']password["']|signIn|getAccessToken|OfficeRuntime\.auth|msal)\b/i.exec(APP + PANE);
      expect(
        asks?.[0],
        `the pane now has "${asks?.[0]}" in it, and the notes for certification promise a reviewer ` +
          `that nothing asks them to sign in. Amend the note in the same change that adds the call.`,
      ).toBe(undefined);
    });
  });

  /**
   * THE REQUIREMENT SET, which decides which hosts a reviewer may test on. It
   * moved twice in a fortnight (1.4 -> 1.8 -> 1.10) and the note states it in
   * prose, where nothing recomputes it.
   */
  describe("the minimum requirement set", () => {
    it("states the floor the manifest actually enforces", () => {
      const floor = /<Set Name="PowerPointApi" MinVersion="([\d.]+)"/.exec(MANIFEST)?.[1];
      expect(floor, "no PowerPointApi requirement in manifest.xml").toBeTruthy();
      expect(
        draft.includes(`PowerPointApi ${floor}`),
        `the notes name a floor the manifest does not enforce. The manifest says ${floor}.`,
      ).toBe(true);
    });

    /**
     * The two build numbers are the ones a reviewer checks their own install
     * against. They are researched values, not derived ones, so what this pins
     * is that the note and the manifest's own working tell the same story.
     */
    it.each(["19610.20002", "16.105"])("agrees with the manifest on build %s", (build) => {
      expect(MANIFEST, `manifest.xml no longer mentions build ${build}`).toContain(build);
      expect(draft, `the notes quote build ${build} and the manifest's reasoning no longer does`).toContain(build);
    });
  });
});
