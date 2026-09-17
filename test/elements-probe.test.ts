import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import {
  ELEMENTS,
  GROUP_NAME,
  QUIESCE_FLOOR_MS,
  SETTLE_BUDGET_MS,
  SETTLE_POLL_MS,
  hostSilent,
  judge,
  readAlt,
  readNote,
  settleReads,
  summarise,
  // @ts-expect-error — a .mjs tool script with no types, imported for its pure helpers.
} from "../scripts/elements-probe.mjs";

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

/**
 * Drive `settleReads` against a scripted host. `frames` is what each read
 * answers in order (a `null` is the host refusing); the clock advances one poll
 * per read, so a test that would take a minute of wall time takes none.
 */
const settled = (frames: unknown[], before: unknown) => {
  let at = 0;
  let clock = 0;
  return settleReads(
    before,
    () => frames[Math.min(at++, frames.length - 1)],
    async () => {
      clock += SETTLE_POLL_MS;
    },
    { now: () => clock },
  );
};
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
    const v = judge(harvey, [shape("1", "Title 1")], [shape("1", "Title 1"), shape("2", GROUP_NAME, "")], "clicked");
    expect(v.ok).toBe(false);
    expect(v.why).toMatch(/carries no altTextDescription/);
    expect(v.ungrouped, "the group formed — this is the write, not the grouping").toBeFalsy();
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
    // The flag, not the sentence, is what the exit code reads. A reworded
    // message must not be able to turn "could not ask" back into "asked, and
    // the answer was wrong".
    expect(v.silent).toBe(true);
  });

  it("calls the run unmeasured when the host answered NOTHING", () => {
    // 2026-09-17: all five came back silent and the probe printed "0 of 5
    // elements carried an altTextDescription", exit 1, against a build measured
    // working hours earlier. Exit 1 says the product was asked. It was not.
    const all = ELEMENTS.map((e: { el: string; expect: RegExp }) => judge(e, null, null, "clicked"));
    expect(all.every((v: { ok: boolean }) => !v.ok)).toBe(true);
    expect(hostSilent(all)).toBe(true);
  });

  it("does NOT excuse the run when one element answered", () => {
    // The dangerous direction. A real regression in one element, next to four
    // host stalls, must stay a finding — an environment excuse that swallows it
    // would make this probe unable to ever report the thing it exists for.
    const stalled = judge(harvey, null, null, "clicked");
    const landedBare = judge(harvey, [], [shape("2", "PowerChart", "")], "clicked");
    expect(hostSilent([stalled, landedBare])).toBe(false);
    expect(hostSilent([landedBare])).toBe(false);
    // A button the pane never offered is also an answer — about the pane.
    expect(hostSilent([stalled, judge(harvey, [], [], "disabled")])).toBe(false);
  });

  it("does not count an unasked element as one that failed", () => {
    // The live run on 2026-09-17: `check` answered and passed, the host went
    // quiet on the other four. It printed "1 of 5 elements carried an
    // altTextDescription" and exited 1 — the reading a user gets from that is
    // four broken elements. Four were never asked.
    const asked = judge(harvey, [], [shape("2", "PowerChart", "Harvey ball, 75% filled.")], "clicked");
    const quiet = ELEMENTS.slice(1).map((e: { el: string; expect: RegExp }) => judge(e, null, null, "clicked"));
    const s = summarise([asked, ...quiet]);
    expect(s.carried).toEqual(["harvey"]);
    expect(s.failed).toEqual([]);
    expect(s.unasked).toEqual(["check", "flow", "kpi", "table"]);
    expect(s.code, "nothing answered wrongly, so this is a re-run, not a finding").toBe(2);
  });

  it("does not file a click that landed nothing under the description write", () => {
    // 2026-09-17: `harvey` put no shape on the slide at all and the report said
    // "1 grouped and were not described: harvey" — because the bucket was
    // computed by subtracting the ungrouped ones from the failures rather than
    // read off the verdict. It sent the reader to `powerpoint.ts` for a click
    // whose shapes never arrived.
    const none = judge(harvey, [], [], "clicked");
    const loose = judge(harvey, [], [shape("2", "harvey-ring")], "clicked");
    const bare = judge(harvey, [], [shape("2", GROUP_NAME, "")], "clicked");
    const s = summarise([none, loose, bare]);
    expect(s.absent).toEqual(["harvey"]);
    // A button the pane never offered is not "the click was accepted".
    expect(summarise([judge(harvey, [], [], "disabled")]).unclickable).toEqual(["harvey"]);
    expect(summarise([judge(harvey, [], [], "disabled")]).absent).toEqual([]);
    expect(s.ungrouped).toEqual(["harvey"]);
    expect(s.undescribed).toEqual(["harvey"]);
    // All three are still failures. Naming the cause must not excuse any of them.
    expect(s.failed).toHaveLength(3);
    expect(s.code).toBe(1);
  });

  it("counts a wrongly-described element against the write, not the grouping", () => {
    const wrong = judge(harvey, [], [shape("2", GROUP_NAME, "Revenue by segment, stacked column.")], "clicked");
    expect(summarise([wrong]).undescribed).toEqual(["harvey"]);
    expect(summarise([wrong]).ungrouped).toEqual([]);
  });

  it("keeps a real failure a finding even when the host stalled around it", () => {
    // The direction that matters. If a stall could downgrade the run, a genuine
    // regression in one element would file itself as an environment problem.
    const broke = judge(harvey, [], [shape("2", "PowerChart", "")], "clicked");
    const quiet = judge(harvey, null, null, "clicked");
    expect(summarise([broke, quiet]).code).toBe(1);
    expect(summarise([broke, quiet]).failed).toEqual(["harvey"]);
  });

  it("is green only when every element was asked and answered", () => {
    const pass = judge(harvey, [], [shape("2", "PowerChart", "Harvey ball, 10% filled.")], "clicked");
    expect(summarise([pass, pass]).code).toBe(0);
    expect(summarise([pass, pass]).carried).toEqual(["harvey", "harvey"]);
  });

  it("names the grouping refusal rather than blaming the description", () => {
    // Measured on a real host, 2026-09-17: one slide held four correct
    // descriptions on four `PowerChart` groups AND sixty loose parts —
    // `harvey-fill-f0`, `step-2`, `cell-text-1-3` — from inserts that never
    // grouped. The description is written on the group, so an insert that never
    // grouped has nowhere to put it. Reporting that as "no altTextDescription"
    // sends the reader to the 1.10 write, which is working, instead of to the
    // grouping refusal, which is not.
    const loose = judge(harvey, [], [shape("2", "harvey-ring"), shape("3", "harvey-fill-f0")], "clicked");
    expect(loose.ok).toBe(false);
    expect(loose.ungrouped).toBe(true);
    expect(loose.why).toMatch(/NO GROUP FORMED/);
    // AND IT MUST NOT PICK A CAUSE. Loose parts with no group are equally an
    // insert the host refused to group and a draw that stopped before reaching
    // the grouping sync — `SHAPES_PER_SYNC` is 10, so a stalled 23-shape table
    // leaves exactly this. The first version of this message asserted the
    // refusal, and the evidence file then repeated it about a table whose ten
    // parts were one sync of a draw the pane had already reported dying in.
    expect(loose.why, "names both causes").toMatch(/refused to group/);
    expect(loose.why, "names both causes").toMatch(/draw stopped/);
    expect(loose.why, "does not assert one of them").not.toMatch(/this is the grouping refusal/);

    // The other side: the group DID form and carries nothing. That one really
    // is the regression this probe exists for, and must not be excused.
    const bare = judge(harvey, [], [shape("2", GROUP_NAME), shape("3", "harvey-ring")], "clicked");
    expect(bare.ok).toBe(false);
    expect(bare.ungrouped).toBeFalsy();
    expect(bare.why).toMatch(/group carries no altTextDescription/);
  });

  it("keeps its copy of the group name equal to the renderer's", () => {
    // Three copies of one lookup key: `ooxml.ts`, `powerpoint.ts`, and this
    // probe. If the renderer renames the group, a probe still looking for
    // "PowerChart" would report every insert as ungrouped.
    for (const file of ["src/render/powerpoint.ts", "src/render/ooxml.ts"]) {
      const src = readFileSync(new URL(`../${file}`, import.meta.url), "utf8");
      expect(src, file).toContain(`const GROUP_NAME = ${JSON.stringify(GROUP_NAME)}`);
    }
  });

  it("waits past the gap between the draw sync and the grouping sync", async () => {
    // THE WHOLE REASON `settleReads` EXISTS, and the case that made it wrong
    // twice in one hour. The renderer commits the loose parts on one sync and
    // only then commits addGroup + the name + the alt text on a second. An
    // outside reader in between sees a stable set of parts and no group, which
    // is indistinguishable from a grouping refusal.
    //
    // This host holds that state for three reads before the group arrives. A
    // two-read quiesce rule stops inside it and reports the refusal.
    const frames = [
      [shape("1", "Title 1")], // pre-existing, already in `before`
      [shape("1", "Title 1"), shape("2", "harvey-ring"), shape("3", "harvey-fill-f0")],
      [shape("1", "Title 1"), shape("2", "harvey-ring"), shape("3", "harvey-fill-f0")],
      [shape("1", "Title 1"), shape("2", "harvey-ring"), shape("3", "harvey-fill-f0")],
      [shape("1", "Title 1"), shape("4", GROUP_NAME, "Harvey ball, 75% filled.")],
    ];
    const r = await settled(frames, [shape("1", "Title 1")]);
    expect(r.settled).toBe(true);
    expect(
      r.after.some((s: { alt: string }) => s.alt),
      "stopped before the group was described",
    ).toBeTruthy();
    // And judged, end to end, it is a pass rather than the grouping refusal the
    // old rule reported.
    expect(judge(harvey, [shape("1", "Title 1")], r.after, "clicked").ok).toBe(true);
  });

  it("stops once the slide has been quiet long enough", async () => {
    // The other side: an insert that really did leave loose parts must not cost
    // the whole budget. Quiet for three reads AND past the floor, then it stops
    // — with `settled` true, because the answer is real.
    const loose = [shape("2", "harvey-ring"), shape("3", "harvey-fill-f0")];
    const r = await settled([[], loose, loose, loose, loose, loose, loose, loose, loose], []);
    expect(r.settled).toBe(true);
    expect(r.waitedMs, "left before the floor").toBeGreaterThanOrEqual(QUIESCE_FLOOR_MS);
    expect(r.waitedMs, "spent the whole budget on a slide that had stopped").toBeLessThan(SETTLE_BUDGET_MS);
    expect(judge(harvey, [], r.after, "clicked").ungrouped).toBe(true);
  });

  it("says STILL CHANGING rather than reporting the last read as final", async () => {
    // A slide that never stops growing. The verdict is read off the last
    // successful read either way, so without `settled` it is indistinguishable
    // from a finished insert — which is how "no group formed" gets reported for
    // an element that was still being drawn.
    const growing = Array.from({ length: 40 }, (_, i) =>
      Array.from({ length: i + 1 }, (_, k) => shape(String(k + 2), "step-" + k)),
    );
    const r = await settled([[], ...growing], []);
    expect(r.settled).toBe(false);
    expect(r.waitedMs).toBeGreaterThanOrEqual(SETTLE_BUDGET_MS);
    expect(r.after, "threw away the evidence it did have").not.toBeNull();
  });

  it("does not count a refused read toward the slide having gone quiet", async () => {
    // A null is the host DECLINING, not the slide holding still. If a stall
    // counted as "nothing changed", two of them plus one real read would satisfy
    // the quiesce rule — and on this host stalls come in runs of 45-60s.
    //
    // The sequence is built so the nulls are what decide. Five growing reads
    // (nothing stable), two refusals, then the same set again, and only THEN
    // the group with its description. Counting the refusals stops one read
    // early and reports loose parts; not counting them reaches the answer.
    const parts = (n: number) => Array.from({ length: n }, (_, i) => shape(String(i + 2), "step-" + i));
    const r = await settled(
      [
        [],
        parts(1),
        parts(2),
        parts(3),
        parts(4),
        parts(5),
        null,
        null,
        parts(5),
        [shape("9", GROUP_NAME, "Process flow: Scope, Design. Scope highlighted.")],
      ],
      [],
    );
    expect(r.settled).toBe(true);
    expect(
      r.after.some((s: { alt: string }) => s.alt),
      "two stalls were counted as quiet and it stopped before the group arrived",
    ).toBe(true);
  });

  it("does not diff against a `before` the host never gave", async () => {
    // `before ?? []` made every shape on the slide look new, so the first read
    // satisfied "something was added", the quiesce test compared a set that
    // never changes, and it stopped ~5s after the click having handed `judge` a
    // full slide as the added set. There is nothing to wait for — `judge` marks
    // a null `before` silent whatever comes back.
    const busy = [shape("1", "Title 1"), shape("2", GROUP_NAME, "Checkbox, yes.")];
    const r = await settled([busy, busy, busy, busy, busy, busy, busy, busy], null);
    expect(r.settled, "claimed a settled answer about a slide it could not diff").toBe(false);
    expect(r.waitedMs, "waited out a budget for an answer that cannot exist").toBeLessThan(QUIESCE_FLOOR_MS);
    expect(judge(harvey, null, r.after, "clicked").silent).toBe(true);
  });

  it("reads what the PANE said, including its escaped spelling", () => {
    // The instrument that decides between "the host refused to group" and "the
    // draw stopped before it got there". Same escaped-quote trap as `readAlt`:
    // the CLI returns `"` as `\"`, and a parser that cannot read the answer
    // reports the same thing as a pane that said nothing.
    const plain =
      'note:{"text":"PowerPoint did not respond while drawing shapes 1-10 of 23 (45s)","cls":"hint status-err"}';
    expect(readNote(plain)?.text).toMatch(/did not respond while drawing shapes 1-10 of 23/);
    const escaped = 'note:{\\"text\\":\\"Inserted 1 chart.\\",\\"cls\\":\\"hint status-ok\\"}';
    expect(readNote(escaped)?.text).toBe("Inserted 1 chart.");
    // An empty note is an ANSWER — the pane said nothing — and must not be
    // confused with the read failing.
    expect(readNote('note:{"text":"","cls":"hint"}')?.text).toBe("");
    expect(readNote("nothing like a note")).toBeNull();
    expect(readNote(undefined)).toBeNull();
  });

  it("does not call an empty run silent", () => {
    // Zero results means the loop never ran. Reporting that as "the host was
    // quiet" would give a reassuring reason for a probe that did nothing.
    expect(hostSilent([])).toBe(false);
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
