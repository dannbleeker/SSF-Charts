import { describe, expect, it } from "vitest";
import {
  BATCH,
  KINDS,
  plan,
  readTrial,
  shouldAbort,
  summarise,
  trialScript,
  // @ts-expect-error — a .mjs tool script with no types, imported for its pure helpers.
} from "../scripts/textbox-cost-probe.mjs";

/**
 * The experiment behind BACKLOG item 24: does a batch of TEXT BOXES stall the
 * web host where the same number of geometric shapes does not?
 *
 * Everything that decides the answer is pure and tested here, for the reason
 * `elements-probe.test.ts` gives at length: a probe's own bugs look exactly
 * like the finding it exists to report, and this one is measuring a host that
 * has spent a day refusing to answer. A summariser that quietly averaged a
 * timeout in as a slow trial would manufacture the result it was built to test.
 */

const trial = (kind: string, ok: boolean, ms: number) => ({ kind, ok, ms });

describe("the text-box cost experiment", () => {
  it("alternates the kinds instead of running them in blocks", () => {
    // A host that tires would manufacture this finding if the kinds ran in
    // blocks: the later kind is slower because it is later. This archive has
    // already confused session age with a per-slide cost curve once.
    const order: string[] = plan(3).map((t: { kind: string }) => t.kind);
    expect(order).toEqual(["rect", "text", "mixed", "rect", "text", "mixed", "rect", "text", "mixed"]);
    // Each kind sits at the same average position, which is the property that
    // makes drift cancel rather than accumulate on one arm.
    const meanPos = (k: string) =>
      order
        .map((x, i) => [x, i] as const)
        .filter(([x]) => x === k)
        .reduce((a, [, i]) => a + i, 0) / 3;
    expect(meanPos("text") - meanPos("rect")).toBe(1);
    expect(meanPos("mixed") - meanPos("text")).toBe(1);
  });

  it("sends the same NUMBER of shapes in every kind — only composition varies", () => {
    // The whole experiment is void if the batches differ in size, because size
    // is the explanation it exists to rule out.
    for (const k of KINDS as { kind: string }[]) {
      const src = trialScript(k.kind);
      expect(src, k.kind).toContain(`i < ${BATCH}`);
    }
    expect(trialScript("text")).toContain("addTextBox");
    expect(trialScript("text")).not.toContain("addGeometricShape");
    expect(trialScript("rect")).toContain("addGeometricShape");
    expect(trialScript("rect")).not.toContain("addTextBox");
    // `mixed` is the table's real first batch: two rules, then cell texts.
    expect(trialScript("mixed")).toContain("addLine");
    expect(trialScript("mixed")).toContain("addTextBox");
    expect(trialScript("mixed")).toContain("i < 2");
  });

  it("never averages a timed-out trial in as a slow one", () => {
    // THE ERROR THIS REPO SPENT 2026-09-17 UNDOING, in a new place. A trial that
    // hit the budget is a NON-ANSWER; folding its 60s into the median would
    // invent a number nothing measured, and would make the hypothesis look
    // supported precisely when the evidence is weakest.
    const results = [
      trial("rect", true, 1000),
      trial("rect", true, 1200),
      trial("text", true, 1100),
      trial("text", false, 60000),
      trial("mixed", true, 1000),
      trial("mixed", true, 1000),
    ];
    const { by } = summarise(results);
    expect(by.text.medianMs, "the 60s non-answer entered the median").toBe(1100);
    expect(by.text.completed).toBe(1);
    expect(by.text.timedOut).toBe(1);
  });

  it("REFUSES to compare when too little completed", () => {
    // One completed trial per kind is not a comparison, and a verdict off it
    // would read as evidence. Exit 2 is "could not ask", the same meaning it
    // carries in the Elements probe.
    const { verdict } = summarise([trial("rect", true, 1000), trial("text", true, 9000), trial("mixed", true, 9000)]);
    expect(verdict.code).toBe(2);
    expect(verdict.says).toMatch(/NOT ENOUGH COMPLETED TRIALS/);
  });

  it("calls a total text failure the STRONGEST form of the result, not a gap", () => {
    // If no text batch ever completes while the rectangles all do, "text is
    // slower" understates it — the batch does not finish at all, which is
    // exactly what the table element does. That case must not fall through to
    // the median comparison, where `text.medianMs` is null.
    const { verdict } = summarise([
      trial("rect", true, 900),
      trial("rect", true, 1100),
      trial("text", false, 60000),
      trial("text", false, 60000),
      trial("mixed", false, 60000),
      trial("mixed", false, 60000),
    ]);
    expect(verdict.code).toBe(0);
    expect(verdict.says).toMatch(/STRONGER THAN SLOWER/);
  });

  it("REFUTES the hypothesis when the gap is not there", () => {
    // The direction that matters. This probe exists to be able to come back
    // "no" — a measurement that can only confirm is not a measurement, and the
    // cost of a false yes here is a fix aimed at the wrong thing.
    const { verdict } = summarise([
      trial("rect", true, 1000),
      trial("rect", true, 1100),
      trial("text", true, 1200),
      trial("text", true, 1300),
      trial("mixed", true, 1250),
      trial("mixed", true, 1250),
    ]);
    expect(verdict.code).toBe(1);
    expect(verdict.says).toMatch(/REFUTED/);
    expect(verdict.says).toMatch(/look elsewhere/);
  });

  it("supports it only on a gap wide enough to act on", () => {
    // 2x is the bar. A 30% difference between two medians of four trials on a
    // host this noisy is not something to redesign a renderer around.
    const near = summarise([
      trial("rect", true, 1000),
      trial("rect", true, 1000),
      trial("text", true, 1900),
      trial("text", true, 1900),
      trial("mixed", true, 1900),
      trial("mixed", true, 1900),
    ]);
    expect(near.verdict.code, "1.9x read as support").toBe(1);
    const wide = summarise([
      trial("rect", true, 1000),
      trial("rect", true, 1000),
      trial("text", true, 9000),
      trial("text", true, 9000),
      trial("mixed", true, 9000),
      trial("mixed", true, 9000),
    ]);
    expect(wide.verdict.code).toBe(0);
    expect(wide.verdict.says).toMatch(/9\.0x/);
  });

  it("gives up after one round in which nothing completed", () => {
    // 2026-09-18: all twelve trials spent their 60s budget against a pane that
    // had died — INCLUDING `rect`, the arm that is meant to be fast. Twelve
    // minutes for twelve identical non-answers, and the twelfth was worth no
    // more than the third.
    const dead = [trial("rect", false, 60000), trial("text", false, 60000), trial("mixed", false, 60000)];
    expect(shouldAbort(dead)).toBe(true);
    // One round is the unit. Two timeouts is this host on an ordinary bad day.
    expect(shouldAbort(dead.slice(0, 2))).toBe(false);
    // And ONE completed trial anywhere in that round means the host is alive,
    // so the run must continue — otherwise a single slow `rect` would abort an
    // experiment that was about to produce its answer.
    expect(shouldAbort([trial("rect", true, 900), trial("text", false, 60000), trial("mixed", false, 60000)])).toBe(
      false,
    );
    expect(shouldAbort([trial("rect", false, 60000), trial("text", true, 900), trial("mixed", false, 60000)])).toBe(
      false,
    );
    expect(shouldAbort([])).toBe(false);
  });

  it("reads the host's answer, including the CLI's escaped spelling", () => {
    expect(readTrial('trial:{"ok":true,"ms":1234}')).toEqual({ ok: true, ms: 1234 });
    expect(readTrial('trial:{\\"ok\\":false,\\"ms\\":60000,\\"why\\":\\"budget\\"}')).toMatchObject({
      ok: false,
      why: "budget",
    });
    expect(readTrial("nothing here")).toBeNull();
    expect(readTrial(undefined)).toBeNull();
  });
});
