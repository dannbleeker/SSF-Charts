import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import {
  GROUP_NAME,
  IDLE_BUDGET_MS,
  INSERT_SENTINEL,
  IDLE_POLL_MS,
  SHOTS,
  SHOT_WIDTH,
  SLICE,
  busyScript,
  clearScript,
  readString,
  readValue,
  renderScript,
  setUpAndInsertScript,
  verifyScript,
  // @ts-expect-error — a .mjs tool script with no types, imported for its pure helpers.
} from "../scripts/store-shots.mjs";

/**
 * THE CAPTURE THAT REPORTED A WORKING PRODUCT BROKEN.
 *
 * `scripts/store-shots.mjs` renders the store listing images through
 * `Slide.getImageAsBase64`. Its first two runs printed "the pane never settled"
 * for every shot, and the pane had been answering `"idle:Done."` throughout:
 * `pw()` joined stdout to stderr, the CLI's update banner landed after the
 * value, and the tail-anchored parser read null on every poll forever.
 *
 * That is the same class as the Elements probe reporting a good host as broken
 * five times: **a capture tool's bug is indistinguishable from the failure it
 * exists to report**. Here it was indistinguishable from "the product cannot
 * draw a chart" — and the first diagnosis, that the host was merely slow, was
 * wrong and was believed for an hour because a leftover `PowerChart` group on
 * the slide seemed to corroborate it. It had been left by an earlier manual
 * run.
 *
 * So the parts that decide a verdict are pinned here: what the transfer check
 * checks, that no host call happens while a draw is running, and — since this
 * is where it actually broke — that a value is only ever read from stdout.
 */

describe("the store screenshot capture", () => {
  /**
   * THE RULE THE FILE HEADER IS BUILT AROUND. Polling the SLIDE during a draw
   * is what broke the Elements probe. This waits on a CSS class instead, which
   * costs the host nothing — and a future edit that reaches for `PowerPoint.run`
   * here would be reintroducing the bug by hand.
   */
  it("asks the DOM, never the host, while it is waiting", () => {
    const watcher = busyScript();
    expect(watcher).toContain("getElementById('host-note')");
    expect(
      /PowerPoint\.run|\.sync\(|getSelectedSlides/.test(watcher),
      "the busy check reaches into the host. Polling the slide during a draw is what made the " +
        "Elements probe report a working product broken five times — see this file's header.",
    ).toBe(false);
  });

  /**
   * The one direct measurement is 63s, for a 780x420 waterfall on a rested
   * session; a deep session costs more. This pins the budget above that with
   * room, NOT because 180s was ever shown to be too short — it never was, the
   * parser was broken — but because a budget under the only number anyone has
   * measured would be a red nobody could interpret.
   */
  it("gives a chart longer than the deepest measured draw", () => {
    expect(IDLE_BUDGET_MS).toBeGreaterThanOrEqual(300_000);
    expect(IDLE_POLL_MS).toBeGreaterThanOrEqual(1000);
    expect(IDLE_BUDGET_MS / IDLE_POLL_MS, "more polls than this is noise, not patience").toBeLessThan(1000);
  });

  describe("reading a value back out of the pane", () => {
    it("unescapes what the CLI escaped", () => {
      expect(readString('"idle:Done."')).toBe("idle:Done.");
      expect(readString('"a \\"quoted\\" word"')).toBe('a "quoted" word');
    });

    /**
     * The CLI prints a banner before the value. Taking the LAST quoted run is
     * what steps over it — an earlier version of this regex would have returned
     * a line of the update notice.
     */
    it("takes the value, not the banner above it", () => {
      const out = ["╔═══════════════════╗", "║ Update available  ║", "╚═══════════════════╝", "", '"idle:Done."'].join(
        "\n",
      );
      expect(readString(out)).toBe("idle:Done.");
    });

    /**
     * NULL IS A DISTINCT ANSWER and the caller depends on it. The first run
     * could not tell "the eval returned nothing" from "the host is still busy",
     * and reported both as the same sentence.
     */
    it("returns null rather than a guess when nothing came back", () => {
      expect(readString("")).toBe(null);
      expect(readString("Error: no such element")).toBe(null);
      expect(readString(null)).toBe(null);
    });

    /**
     * THE BUG THAT COST TWO RUNS, pinned at the seam it actually crossed.
     *
     * The CLI puts the VALUE on stdout and its "Update available" banner on
     * stderr. The caller used to hand `readString` the two concatenated, so the
     * string ended with a box-drawing frame instead of the value and every poll
     * of the pane read as null — for two full runs, while the pane had been
     * saying `"idle:Done."` the whole time.
     *
     * The reason it survived is the reason this test exists: the only check was
     * `readString` on a clean string, which no call path produces. A terminal
     * interleaves the streams so the banner lands FIRST and it all works by
     * hand; `spawnSync` captures them apart and joins them in a fixed order.
     * So the fixture below is the real shape — value on `out`, banner on `err`.
     */
    describe("with the streams the CLI really returns", () => {
      const BANNER =
        "\n╔════════════════════════════════════════════════════════════════════╗\n" +
        "║ Update available for @playwright/cli: 0.1.18 → 0.1.21              ║\n" +
        "╚════════════════════════════════════════════════════════════════════╝\n\n";

      it("reads the value even when stderr is noisy", () => {
        expect(readValue({ out: '"idle:Done."\n', err: BANNER, status: 0 })).toBe("idle:Done.");
      });

      it("never lets stderr stand in for a value", () => {
        // If the banner could reach the parser, this is what used to happen.
        expect(readValue({ out: "", err: BANNER, status: 0 })).toBe(null);
        expect(readValue({ out: "", err: '"idle:Done."', status: 1 })).toBe(null);
      });

      it("survives a missing result entirely", () => {
        expect(readValue(undefined)).toBe(null);
        expect(readValue({ out: undefined, err: "", status: 0 })).toBe(null);
      });
    });
  });

  /**
   * THE CALL SITE THAT WAS MISSED.
   *
   * `pw()` used to return a joined string and now returns `{out, err, status}`.
   * Four of its five call sites were updated; `refFor` was not, and the script
   * died on `pw(...).split is not a function` the first time it ran. Unit tests
   * could not see it — `refFor` lives inside `main()` and talks to a live host.
   *
   * So this reads the source instead. It is a sweep, not a behaviour test, and
   * that is the point: the defect class here is "one place still speaks the old
   * shape", which only something that looks at EVERY place can catch.
   */
  it("consumes every pw() result as the object pw() returns", () => {
    const src = readFileSync(resolve(__dirname, "..", "scripts/store-shots.mjs"), "utf8");
    const uses = [...src.matchAll(/\bpw\([^\n]*?\)(\s*\.\s*\w+|\s*[),;])/g)].map((m) => m[0]);
    expect(uses.length, "no pw() call sites found — has the helper been renamed?").toBeGreaterThan(3);
    for (const use of uses) {
      // Either handed to the reader, or reached into by stream name. Anything
      // else is treating the result as the string it stopped being.
      const ok = /\.\s*(out|err|status)\b/.test(use) || /[),;]$/.test(use.trim());
      expect(ok, `\`${use.trim()}\` uses a pw() result as if it were still a string`).toBe(true);
    }
    expect(
      /\bpw\([^\n]*?\)\s*\.\s*(split|trim|includes|match|replace|slice|startsWith)\b/.test(src),
      "a pw() result has a string method called on it — pw() returns {out, err, status}",
    ).toBe(false);
  });

  describe("the slice transfer", () => {
    /**
     * THE CHECK THAT MATTERS. A truncated base64 string still decodes to a
     * valid PNG *prefix* — header, dimensions and all — so a short transfer
     * produces a file that opens, reports the right size, and is corrupt below
     * the fold. Only the length catches it, and only if the length is asked for
     * before the slices rather than derived from them.
     */
    it("asks the page for the length before it asks for any of the content", () => {
      const render = renderScript();
      expect(render).toContain("window.__shot");
      expect(
        render,
        "the render must report the length it stashed, or the assembled string has nothing to match",
      ).toContain('"len:" + img.value.length');
    });

    it("slices in chunks small enough to survive a shell-quoted line", () => {
      expect(SLICE).toBeGreaterThan(0);
      expect(SLICE, "a slice this large is the single long line the slicing exists to avoid").toBeLessThanOrEqual(
        32_000,
      );
    });
  });

  describe("what gets captured", () => {
    it("renders at the size AppSource asks for", () => {
      expect(SHOT_WIDTH).toBe(1366);
      expect(renderScript()).toContain(`width: ${SHOT_WIDTH}`);
    });

    /**
     * A store image is the product's first impression. `test`, `foo` and an
     * empty title all read as unfinished software to a reviewer who has never
     * seen the pane.
     */
    it("uses titles a consultant would actually type", () => {
      expect(SHOTS.length).toBeGreaterThan(0);
      expect(SHOTS.length, "AppSource accepts at most five listing images").toBeLessThanOrEqual(5);
      for (const shot of SHOTS as { file: string; title: string; w: number; h: number }[]) {
        expect(shot.file).toMatch(/^\d\d-[a-z]+\.png$/);
        expect(shot.title.length, `"${shot.title}" is too short to read as a real deck's chart`).toBeGreaterThan(8);
        expect(
          /\b(test|foo|bar|baz|sample|untitled|lorem)\b/i.test(shot.title),
          `"${shot.title}" is placeholder text, and a listing image is the first thing a reviewer sees`,
        ).toBe(false);
        expect(shot.w).toBeGreaterThan(0);
        expect(shot.h).toBeGreaterThan(0);
      }
    });

    it("gives every shot its own file", () => {
      const files = (SHOTS as { file: string }[]).map((s) => s.file);
      expect(new Set(files).size, "two shots write to one path — the second would silently replace the first").toBe(
        files.length,
      );
    });

    /**
     * THE STALE-NOTE RACE. The pane sets `status-busy` in its own click
     * handler, so for a moment after the click the note still carries the
     * PREVIOUS action's "Done." — and a waiter that polls in that window reads
     * a success belonging to something else and captures a blank slide.
     *
     * Stamping the note busy before the click closes it. The order is the whole
     * point: a stamp that landed after `.click()` would guard nothing, and both
     * orderings look equally reasonable in a diff.
     */
    it("stamps the note busy BEFORE it clicks insert, not after", () => {
      const setup = setUpAndInsertScript({ kind: "waterfall", title: "Revenue bridge FY25", w: 780, h: 420 });
      const stamp = setup.indexOf("status-busy");
      const click = setup.indexOf('getElementById("insert").click()');
      expect(stamp, "the insert no longer stamps the note — the stale-note race is back").toBeGreaterThan(-1);
      expect(click).toBeGreaterThan(-1);
      expect(
        stamp < click,
        "the note is stamped AFTER the click, which leaves the window open: a poll in between " +
          "reads the previous action's success and captures a slide nothing has drawn on yet.",
      ).toBe(true);
      expect(setup).toContain(INSERT_SENTINEL);
    });

    /**
     * THE PICTURE THAT CONTRADICTED ITS OWN CAPTION.
     *
     * The first successful capture came out as a STACKED COLUMN chart titled
     * "Revenue bridge FY25", because the script set a title and never chose a
     * kind — so it drew the pane's default. A bridge is a waterfall. A store
     * image whose caption misdescribes the picture is the exact shape of an
     * AppSource "functionality does not match the offer description".
     *
     * The ORDER is the second half of it: the gallery tile's handler calls
     * `applyConfig(sampleConfig(kind))`, which replaces the whole config
     * including the title. Choosing the kind after typing the title throws the
     * title away, and the failure is invisible — you get a real chart with the
     * sample's caption.
     */
    it("chooses the chart kind BEFORE typing the title", () => {
      const setup = setUpAndInsertScript({ kind: "waterfall", title: "Revenue bridge FY25", w: 780, h: 420 });
      const kind = setup.indexOf('[data-kind="waterfall"]');
      const title = setup.indexOf("chart-title");
      expect(kind, "the setup never selects a chart kind — it will draw whatever the pane defaults to").toBeGreaterThan(
        -1,
      );
      expect(title).toBeGreaterThan(-1);
      expect(
        kind < title,
        "the kind is chosen AFTER the title is typed. The gallery tile applies a fresh sample " +
          "config, which overwrites the title — so the shot would carry the sample's caption.",
      ).toBe(true);
    });

    /**
     * And the filename has to agree with both. `01-waterfall.png` holding a
     * stacked chart is a mislabel that survives every other check here.
     */
    it("names each file after the kind it actually draws", () => {
      for (const shot of SHOTS as { file: string; kind: string }[]) {
        expect(shot.kind, `${shot.file} does not say which chart kind to draw`).toBeTruthy();
        expect(
          shot.file.includes(shot.kind),
          `${shot.file} draws a "${shot.kind}" chart. The filename is what a person picks from a ` +
            `folder, and it currently names a different chart.`,
        ).toBe(true);
      }
    });

    it("types the title into the pane the way a user does, so nothing is faked in", () => {
      const setup = setUpAndInsertScript({ kind: "waterfall", title: "Revenue bridge FY25", w: 780, h: 420 });
      expect(setup).toContain("Revenue bridge FY25");
      expect(setup).toContain('getElementById("insert").click()');
      expect(setup, "the width the shot asks for must reach the pane's own input").toContain('"780"');
    });
  });

  /**
   * THE BLANK-SLIDE BLIND SPOT, and the sibling of the length check above.
   *
   * An insert that quietly produced nothing leaves an empty slide.
   * `getImageAsBase64` renders that perfectly: the transfer length matches, the
   * PNG signature is right, the dimensions are right, and the file written is
   * 1366px of white. Every downstream check passes. Only asking the slide what
   * it holds catches it — and a store listing image of an empty slide is the
   * worst thing on this list to ship unnoticed.
   */
  describe("checking a chart actually landed", () => {
    it("asks the slide what is on it before rendering anything", () => {
      const check = verifyScript();
      expect(check).toContain("getSelectedSlides");
      expect(check, "the check must count the group, not just the shapes").toContain(GROUP_NAME);
    });

    /**
     * WHAT IT REFUSES, AND WHAT IT ONLY REPORTS.
     *
     * The first version rejected any slide with no `PowerChart` group, and
     * threw away a good waterfall — 45 shapes drawn, the pane reporting
     * success, and `shapes.addGroup` refused by the host, which
     * `powerpoint.ts` traces as "the host refused addGroup". That costs
     * re-editing, not appearance, and a listing image is an appearance.
     *
     * So the only refusal is the empty slide: the one outcome that renders to a
     * flawless PNG and that nothing downstream can catch.
     */
    it("refuses an empty slide, and an ungrouped one too", () => {
      const src = readFileSync(resolve(__dirname, "..", "scripts/store-shots.mjs"), "utf8");
      expect(src, "no empty-slide refusal — a blank slide would be written as a store image").toMatch(
        /loose === 0[\s\S]{0,200}not capturing a blank slide/,
      );
      /**
       * THIS REFUSAL WAS REMOVED ONCE, ON AN ARGUMENT, AND THE IMAGE REFUTED IT.
       *
       * The argument: the host traces "the host refused addGroup", a refusal
       * costs re-editing rather than appearance, and a listing image is only an
       * appearance. So it was softened to a printed note and the shot written.
       * The shot came back with its title drawn TWICE at two sizes and the axis
       * labels out of frame — a chart layered over the previous one's
       * leftovers. Ungrouped did not mean "drew fine, failed to group".
       *
       * Every other check in that file passed it. Only looking at it caught it.
       */
      expect(
        src,
        "an ungrouped insert is being captured again. It was tried: the picture had two titles on it.",
      ).toMatch(/counts\[2\]\)\s*<\s*1[\s\S]{0,400}failed\+\+/);
    });

    /**
     * THE CAUSE UNDER THAT PICTURE. `clearScript` is raced against a budget and
     * returns `"clear-failed"` when it loses; the caller used to discard that,
     * so the next chart drew on top of the last one's shapes. Both halves are
     * needed — the clear saying it worked, and the slide then reading empty,
     * because a host can report a delete it has not finished applying.
     */
    it("will not draw until the slide is checked empty", () => {
      const src = readFileSync(resolve(__dirname, "..", "scripts/store-shots.mjs"), "utf8");
      expect(src, "the clear's own result is ignored again — `clear-failed` would pass silently").toMatch(
        /cleared !== "cleared"/,
      );
      expect(src, "nothing re-reads the slide after clearing it").toMatch(/emptied/);
      const clearAt = src.indexOf('cleared !== "cleared"');
      const insertAt = src.indexOf("setUpAndInsertScript(shot)");
      expect(clearAt, "the clear is not checked before the insert is clicked").toBeLessThan(insertAt);
    });

    /**
     * NOT A SECOND COPY. The group name is decided by the renderers; restating
     * it here would be one more place to keep in step, which is the defect this
     * repo keeps finding rather than a check against it.
     */
    it("uses the name the renderers really group under", () => {
      const root = resolve(__dirname, "..");
      const names = ["src/render/ooxml.ts", "src/render/powerpoint.ts"].map(
        (f) => /const GROUP_NAME = "([^"]+)"/.exec(readFileSync(resolve(root, f), "utf8"))?.[1],
      );
      expect(new Set(names).size, `the renderers disagree: ${names.join(" vs ")}`).toBe(1);
      expect(
        GROUP_NAME,
        `the capture looks for "${GROUP_NAME}" and the renderers write "${names[0]}", so every ` +
          `shot would be rejected as a blank slide`,
      ).toBe(names[0]);
    });
  });

  /**
   * The clear, the check and the render all go through the host, and a host
   * that stops answering would hang the capture rather than fail it. Each is
   * raced against a budget; none may be sent bare.
   */
  it.each([
    ["the clear", clearScript],
    ["the render", renderScript],
    ["the did-it-land check", verifyScript],
  ])("%s cannot hang forever on a host that stops answering", (_name, make) => {
    const script = (make as () => string)();
    expect(script).toContain("Promise.race");
    expect(script).toMatch(/setTimeout\(\(\) => rej/);
  });
});
