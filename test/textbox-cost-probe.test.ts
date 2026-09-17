import { describe, expect, it } from "vitest";
import {
  BATCH,
  HYPOTHESIS,
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

/**
 * Every kind EXCEPT the one under test, all completing at the same time — so
 * the arm `only` supplies is the single thing that differs.
 *
 * It must EXCLUDE the test kind. Including it and appending the override left
 * that arm with four trials and a median mixing baseline with override, which
 * silently halved the multiple each test was trying to assert — two of these
 * tests passed a 2x threshold while asserting a 5x gap.
 */
const even = (ms: number) =>
  (KINDS as { kind: string }[])
    .filter((k) => k.kind !== HYPOTHESIS.test)
    .flatMap((k) => [trial(k.kind, true, ms), trial(k.kind, true, ms)]);
const only = (kind: string, a: number, b: number) => [trial(kind, true, a), trial(kind, true, b)];
const onlyFailed = (kind: string) => [trial(kind, false, 60000), trial(kind, false, 60000)];

describe("the text-box cost experiment", () => {
  it("alternates the kinds instead of running them in blocks", () => {
    // A host that tires would manufacture this finding if the kinds ran in
    // blocks: the later kind is slower because it is later. This archive has
    // already confused session age with a per-slide cost curve once.
    const order: string[] = plan(3).map((t: { kind: string }) => t.kind);
    expect(order).toEqual([...KINDS, ...KINDS, ...KINDS].map((k: { kind: string }) => k.kind));
    // Each kind sits at the same average position, which is the property that
    // makes drift cancel rather than accumulate on one arm.
    const meanPos = (k: string) =>
      order
        .map((x, i) => [x, i] as const)
        .filter(([x]) => x === k)
        .reduce((a, [, i]) => a + i, 0) / 3;
    expect(meanPos("text") - meanPos("rect")).toBe(1);
    expect(meanPos("mixed") - meanPos("text")).toBe(1);
    expect(meanPos("styled") - meanPos("mixed")).toBe(1);
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

  it("turns on the DECLARED pair, and on nothing else", () => {
    // The verdict compares HYPOTHESIS.test against HYPOTHESIS.control, which
    // differ in exactly one thing. Written against the declaration rather than
    // against kind names on purpose: this file has already carried three
    // hypotheses — composition (refuted at 1.3x), statement count (supported at
    // 3.4x, and 1.6s is not the 45s the table dies at), and now the empty cell.
    // Tests pinned to the current pair would have to be rewritten each time,
    // and rewriting a test to match a new answer is how a suite stops checking.
    const bad = summarise([...even(600), ...only(HYPOTHESIS.test, 5000, 5100)]);
    expect(bad.verdict.code).toBe(0);
    expect(bad.verdict.says, bad.verdict.says).toContain(HYPOTHESIS.about);
    // It quotes both numbers, so a reader cannot mistake which pair the
    // multiple came from.
    expect(bad.verdict.says).toContain("5050");
    expect(bad.verdict.says).toContain("600");
  });

  it("calls a total failure of the test arm the STRONGEST form, not a gap", () => {
    // If no test batch completes while its control does, "slower" understates
    // it — the batch does not finish, which is exactly what the table element
    // does. That case must not fall through to the median comparison, where the
    // test arm's median is null.
    const { verdict } = summarise([...even(600), ...onlyFailed(HYPOTHESIS.test)]);
    expect(verdict.code).toBe(0);
    expect(verdict.says).toMatch(/STRONGER THAN SLOWER/);
    expect(verdict.says).toContain(HYPOTHESIS.control);
  });

  it("REFUTES when the gap is not there", () => {
    // The direction that matters, and one this probe has already delivered:
    // composition was refuted at 555ms against 421ms. A measurement that can
    // only confirm is not a measurement, and the cost of a false yes is a fix
    // aimed at the wrong thing.
    const { verdict } = summarise([...even(600), ...only(HYPOTHESIS.test, 700, 720)]);
    expect(verdict.code).toBe(1);
    expect(verdict.says).toMatch(/REFUTED/);
    expect(verdict.says).toMatch(/look elsewhere/);
  });

  it("supports it only on a gap wide enough to act on", () => {
    // 2x is the bar. A 30% difference between medians of four trials on a host
    // this noisy is not something to redesign a renderer around — and 1.3x is
    // precisely what the refuted composition hypothesis measured.
    expect(summarise([...even(1000), ...only(HYPOTHESIS.test, 1900, 1900)]).verdict.code, "1.9x read as support").toBe(
      1,
    );
    const wide = summarise([...even(1000), ...only(HYPOTHESIS.test, 9000, 9000)]);
    expect(wide.verdict.code).toBe(0);
    expect(wide.verdict.says).toMatch(/9.0x/);
  });

  it("gives up after one round in which nothing completed", () => {
    // 2026-09-18: all twelve trials spent their 60s budget against a pane that
    // had died — INCLUDING `rect`, the arm that is meant to be fast. Twelve
    // minutes for twelve identical non-answers, and the twelfth was worth no
    // more than the third.
    const dead = KINDS.map((k: { kind: string }) => trial(k.kind, false, 60000));
    expect(shouldAbort(dead)).toBe(true);
    // One round is the unit. Three timeouts out of four is this host on an
    // ordinary bad day.
    expect(shouldAbort(dead.slice(0, KINDS.length - 1))).toBe(false);
    // And ONE completed trial anywhere in that round means the host is alive,
    // so the run must continue — otherwise a single slow `rect` would abort an
    // experiment that was about to produce its answer.
    for (let i = 0; i < KINDS.length; i++) {
      const round = dead.map((t: { kind: string; ok: boolean; ms: number }, j: number) =>
        j === i ? { ...t, ok: true, ms: 900 } : t,
      );
      expect(shouldAbort(round), `one live trial at position ${i} still aborted`).toBe(false);
    }
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
