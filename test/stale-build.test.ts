import { describe, it, expect } from "vitest";
import { lazy, isStaleBuild, StaleBuildError } from "../src/render/lazy";
import { errorText } from "../src/render/powerpoint";

/**
 * What the pane says when a release lands while it is open.
 *
 * FOUND BY A ROUND. The 4:3 leg of the 2026-08-28 cycle failed every scenario
 * after the first, and the first one's message was the browser's:
 *
 *     Failed to fetch dynamically imported module:
 *     https://ssf-chart.struktureretsundfornuft.dk/assets/pptxgen.es-C8DOodSg.js
 *
 * A build had been published while the pane was open. Vite hashes every chunk,
 * a deploy replaces the whole `assets/` directory, and the pane was holding an
 * `index.html` naming chunks the server no longer had.
 *
 * It is not a harness problem — it is the ordinary shape of a user's day. The
 * pane stays open for a PowerPoint session, releases go out during one, and the
 * next deck insert asks for a file that has been deleted. What that user saw was
 * a URL, from an add-in that looked broken.
 */
describe("a chunk the server no longer has", () => {
  it("recognises the wording of all three browser families", () => {
    // ALL THREE OR IT IS WORSE THAN NOTHING. Matching only Chrome would give the
    // honest message on Windows and the raw one on a Mac or an iPad, which is
    // the harder of the two bugs to diagnose — the same failure reported two
    // different ways depending on who hit it.
    expect(isStaleBuild(new Error("Failed to fetch dynamically imported module: https://x/assets/a-B1.js"))).toBe(true);
    expect(isStaleBuild(new Error("error loading dynamically imported module"))).toBe(true);
    expect(isStaleBuild(new TypeError("Importing a module script failed."))).toBe(true);
  });

  it("leaves a bare fetch() failure to whoever made the request", () => {
    // `TypeError: Failed to fetch` is `fetch()`'s wording and never `import()`'s,
    // so it is some OTHER request failing — and dressing it as a chunk that would
    // not load would send whoever debugs it to the wrong place. Only the
    // module-loading wording counts. (That wording, when a NETWORK failure
    // produces it, is word for word the deleted-chunk one: it cannot be told
    // apart here, so the sentence names both causes — see
    // `test/fix-r4-deck-and-preview-parity.test.ts`.)
    expect(isStaleBuild(new TypeError("Failed to fetch"))).toBe(false);
    expect(isStaleBuild(new Error("NetworkError when attempting to fetch resource."))).toBe(false);
    expect(isStaleBuild(new Error("An internal error has occurred."))).toBe(false);
    expect(isStaleBuild(undefined)).toBe(false);
    expect(isStaleBuild(null)).toBe(false);
  });

  it("says what happened and what to do, and never shows a URL", () => {
    const err = new StaleBuildError("the deck writer");
    // Names BOTH causes: the browser's error for a chunk that would not load is
    // the same whether the server replaced it or the connection dropped, so the
    // sentence may not assert the first. See `test/fix-r4-deck-and-preview-parity.test.ts`.
    expect(err.message).toContain("was updated");
    expect(err.message).toMatch(/connection/i);
    expect(err.message).toContain("the deck writer");
    expect(err.message, "the pane must say how to recover, not merely what broke").toMatch(/open it again/i);
    // A REASSURANCE, because the failure arrives mid-insert and the honest
    // question is whether the deck was damaged. It was not.
    expect(err.message).toMatch(/slides are untouched/i);
    expect(err.message, "a chunk URL is not something the reader can act on").not.toMatch(/https?:|\.js\b/);
  });

  it("does not dress a real bug up as a stale deploy", () => {
    // If the imported module itself throws, that is a defect in this codebase
    // and it must arrive intact — anything else sends the next person debugging
    // it to the deploy pipeline instead of to the stack trace.
    const real = new Error("Cannot read properties of undefined (reading 'addSlide')");
    return expect(lazy(() => Promise.reject(real), "the deck writer")).rejects.toBe(real);
  });

  it("passes the module through when nothing is wrong", async () => {
    expect(await lazy(() => Promise.resolve({ default: 42 }), "anything")).toEqual({ default: 42 });
  });

  it("replaces the browser's message wherever an error is reported", async () => {
    // `errorText` is the one funnel — the pane's notes, the round self-test's
    // verdicts and the traces all come through it — so the substitution belongs
    // there rather than at each call site.
    const raw = new Error("Failed to fetch dynamically imported module: https://x/assets/pptxgen.es-C8DOodSg.js");
    expect(errorText(raw), "the raw browser message reached the user").not.toMatch(/dynamically imported/i);
    expect(errorText(raw)).toMatch(/was updated/i);
    // And the wrapped form keeps the part name it was given.
    await expect(lazy(() => Promise.reject(raw), "the deck writer")).rejects.toThrow(/the deck writer/);
    await lazy(() => Promise.reject(raw), "the deck writer").catch((e) => {
      expect(errorText(e)).toMatch(/the deck writer/);
    });
  });

  it("keeps the Office.js detail on errors that are NOT this", () => {
    // The guard against over-reaching: `errorText` exists to surface `code` and
    // `debugInfo`, which a plain String(err) drops, and a stale-build shortcut
    // that swallowed those would trade one silent failure for another.
    const office = Object.assign(new Error("An internal error has occurred."), {
      code: "GeneralException",
      debugInfo: { errorLocation: "Slide.id" },
    });
    const said = errorText(office);
    expect(said).toContain("code=GeneralException");
    expect(said).toContain("Slide.id");
  });
});
