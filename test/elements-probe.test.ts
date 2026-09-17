import { describe, expect, it } from "vitest";
// @ts-expect-error — a .mjs tool script with no types, imported for its pure helpers.
import { ELEMENTS, judge, readAlt } from "../scripts/elements-probe.mjs";

/**
 * The probe that asks whether the five Elements buttons land a shape carrying
 * the `desc` v0.6.1 added — `scene.desc -> opts.altText ->
 * Shape.altTextDescription`, a PowerPointApi 1.10 write.
 *
 * WHY THE VERDICT IS A PURE FUNCTION. A probe's own bugs look exactly like the
 * finding it exists to report: "nothing landed" is what you get from a real
 * failure AND from a ref that pointed at the wrong frame. This repo has paid
 * for that distinction more than once — `run-unsettled2.mjs` reported six
 * THREWs that were its own `items[listed - 1]` bug, and the conclusion only
 * changed when empty-read was separated from threw. So everything that decides
 * pass or fail lives above the driving code and is tested here without a host.
 */

const shape = (id: string, name: string, alt = "") => ({ id, name, alt });
const harvey = ELEMENTS.find((e: { el: string }) => e.el === "harvey")!;

describe("what the Elements probe concludes", () => {
  it("passes when a described shape appears that was not there before", () => {
    const before = [shape("1", "Title 1")];
    const after = [shape("1", "Title 1"), shape("2", "PowerChart", "Harvey ball, 75% filled.")];
    expect(judge(harvey, before, after, "clicked")).toMatchObject({ el: "harvey", ok: true, added: 1 });
  });

  it("fails when a shape lands with NO description — the regression it exists for", () => {
    // This is exactly what a v0.6.1 that lost its `desc` wiring would look like:
    // the element draws, and nothing carries alt text.
    const v = judge(harvey, [shape("1", "Title 1")], [shape("1", "Title 1"), shape("2", "PowerChart", "")], "clicked");
    expect(v.ok).toBe(false);
    expect(v.why).toMatch(/none with altTextDescription/);
  });

  it("fails when the description belongs to something else", () => {
    // A stale group from an earlier insert is described — just not as a harvey.
    // Without this, the probe would pass on the previous element's shape.
    const v = judge(harvey, [], [shape("2", "PowerChart", "Revenue by segment, stacked column.")], "clicked");
    expect(v.ok).toBe(false);
    expect(v.why).toMatch(/not as a harvey/);
  });

  it("does not count a shape that was already on the slide", () => {
    // The id set is the whole guard. Re-reading the same described shape after a
    // click that did nothing must not read as a pass.
    const both = [shape("2", "PowerChart", "Harvey ball, 50% filled.")];
    expect(judge(harvey, both, both, "clicked").ok).toBe(false);
  });

  it("reports the BUTTON's state rather than blaming the host", () => {
    for (const said of ["disabled", "no-button"]) {
      const v = judge(harvey, [], [], said);
      expect(v.ok).toBe(false);
      expect(v.why).toContain(said);
    }
  });

  it("says the host would not answer, rather than that nothing landed", () => {
    // null is a refused/failed read. Calling that "nothing landed" would file a
    // host refusal as a product defect, which is the confusion this repo keeps
    // meeting: a check that cannot tell "verified" from "not attempted".
    const v = judge(harvey, null, null, "clicked");
    expect(v.ok).toBe(false);
    expect(v.why).toMatch(/would not list the slide/);
  });

  it("parses the host's answer, and refuses a mangled one", () => {
    expect(readAlt('alt:[{"id":"1","name":"n","alt":"a"}]')).toEqual([{ id: "1", name: "n", alt: "a" }]);
    expect(readAlt("alt-failed:budget")).toBeNull();
    expect(readAlt("alt:[not json")).toBeNull();
    expect(readAlt(undefined)).toBeNull();
  });

  it("covers the five elements the manifest actually ships", () => {
    // If a sixth deep link is added to manifest.xml this list has to grow with
    // it, or the probe silently stops covering the new one.
    expect(ELEMENTS.map((e: { el: string }) => e.el)).toEqual(["harvey", "check", "flow", "kpi", "table"]);
  });
});
