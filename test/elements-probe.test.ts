import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import {
  CLICK_SENTINEL,
  ELEMENTS,
  GROUP_NAME,
  SETTLE_BUDGET_MS,
  SETTLE_POLL_MS,
  clickElementScript,
  hostSilent,
  judge,
  busyScript,
  readAlt,
  readBusy,
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
 * Drive `settleReads` against a scripted PANE. `busyFrames` is what each DOM
 * read of the busy class answers, in order; the clock advances one poll per
 * read, so a minute of waiting takes none.
 */
const settledBy = (busyFrames: unknown[], read: () => unknown, onBusy?: () => void) => {
  let at = 0;
  let clock = 0;
  return settleReads(
    async () => read(),
    async () => {
      onBusy?.();
      return busyFrames[Math.min(at++, busyFrames.length - 1)];
    },
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

  it("makes NO host call until the pane says it has finished", async () => {
    // THE MOST EXPENSIVE DEFECT THIS PROBE HAS HAD. The previous `settleReads`
    // polled the SLIDE every 4s through the draw — a PowerPoint.run per poll,
    // interleaved with the renderer's own batches, on a host that forces a full
    // presentation save on every sync. It broke the insert it was measuring,
    // and reported `table` as broken five times.
    //
    // Measured, same session, same cleared slide, same element, polling the
    // only difference: "Done." without it, and "Failed: PowerPoint did not
    // respond while drawing shapes 11-20 of 23 (45s)" with it.
    //
    // So the COUNT is the assertion: one read, and not before the pane is done.
    // Everything else in this file is about reading the answer correctly; this
    // is about not destroying it.
    // THE ORDER IS THE ASSERTION, not the count — and the first version of this
    // test got that wrong. It checked `reads === 1` and `busyPolls === 4`, both
    // of which hold just as well if the read happens BEFORE the wait, which is
    // the defect itself. Mutation-checked: moving the read above the loop
    // passed all 23 tests in this file. So the calls are recorded in sequence
    // and the read must come last.
    const calls: string[] = [];
    const r = await settledBy(
      [true, true, true, false],
      () => {
        calls.push("read");
        return [shape("2", GROUP_NAME, "Harvey ball, 75% filled.")];
      },
      () => calls.push("busy"),
    );
    expect(calls, "the host was asked out of order, or more than once").toEqual([
      "busy",
      "busy",
      "busy",
      "busy",
      "read",
    ]);
    expect(r.settled).toBe(true);
    expect(judge(harvey, [], r.after, "clicked").ok).toBe(true);
  });

  it("keeps waiting when the pane cannot be read, rather than calling it finished", async () => {
    // null is "cannot tell". Treating an unknown as a finish would read the
    // slide mid-draw, which is the failure above in a subtler dress.
    const r = await settledBy([null, null, false], () => [shape("2", GROUP_NAME, "Checkbox, yes.")]);
    expect(r.settled).toBe(true);
    expect(r.waitedMs).toBe(3 * SETTLE_POLL_MS);
  });

  it("gives up on a pane that never settles, and says so", async () => {
    // A budget that runs out is not a finish. It still reads once — the
    // evidence it does have is worth keeping — but `settled` is false, so the
    // report says STILL CHANGING instead of presenting it as final.
    const r = await settledBy(
      Array.from({ length: 999 }, () => true),
      () => [shape("2", "harvey-ring")],
    );
    expect(r.settled).toBe(false);
    expect(r.waitedMs).toBeGreaterThanOrEqual(SETTLE_BUDGET_MS);
    expect(r.after, "threw away the evidence it had").not.toBeNull();
  });

  it("reads the pane's busy CLASS, and an absent note as cannot-tell", () => {
    expect(readBusy("busy:yes")).toBe(true);
    expect(readBusy("busy:no")).toBe(false);
    expect(readBusy("busy:unknown"), "an absent note must not read as finished").toBeNull();
    expect(readBusy("nothing")).toBeNull();
    expect(readBusy(undefined)).toBeNull();
    // The class, not the text: the text is translated and the class is not.
    expect(busyScript()).toContain("status-busy");
    // And it must stay a DOM read. The moment this touches the host it is the
    // old defect again.
    expect(busyScript(), "the busy check reached for the host").not.toContain("PowerPoint.run");
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

  it("stamps the note busy before clicking, not after", () => {
    // THE SECOND WAY THIS PROBE CAN MANUFACTURE A FALSE RED, found 2026-09-23.
    //
    // The first was polling the host during a draw, which `settleReads` above
    // now prevents. This one is quieter: `#host-note` holds ONE verdict, and
    // between `b.click()` and the pane setting `status-busy` in its own
    // handler, the note still reads the PREVIOUS element's `status-ok`. A poll
    // landing in that window says "settled", `settleReads` reads the slide
    // mid-draw, and a working element is reported broken.
    //
    // The archived 5-of-5 is not in doubt — a stale settle reads the previous
    // element's shapes, which fail this element's `expect`, so the race can
    // only produce failures and none were recorded. Fixed because it is latent.
    //
    // THE ORDER IS THE GUARANTEE. A stamp after the click guards nothing, and
    // the two spellings read identically in a diff.
    const script = clickElementScript("table");
    const stamp = script.indexOf("status-busy");
    // The INSERT click, spelled with what follows it — `b.click()` alone also
    // matches the `tab.click()` that switches tab, which happens first and
    // would make this pass while guarding nothing.
    const click = script.indexOf('b.click(); return "clicked"');
    expect(stamp, "the probe no longer stamps the note — a stale verdict reads as this insert's").toBeGreaterThan(-1);
    expect(click).toBeGreaterThan(-1);
    expect(
      stamp < click,
      "the note is stamped AFTER the click, which leaves the race open: a poll in between reads " +
        "the previous element's success, settles early, and reads the slide mid-draw.",
    ).toBe(true);
    // And it must stay a DOM write. Reaching for the host here would be the
    // expensive defect above, reintroduced by the fix for this one.
    expect(script, "the click script reached for the host").not.toContain("PowerPoint.run");
  });

  it("uses a sentinel no element could be mistaken for", () => {
    // The sentinel goes into the same slot the pane reports through, so a probe
    // that accepted it would be quoting itself back as the product's answer.
    expect(CLICK_SENTINEL).toBeTruthy();
    for (const { el, expect: pattern } of ELEMENTS as { el: string; expect: RegExp }[]) {
      expect(
        pattern.test(CLICK_SENTINEL),
        `the sentinel "${CLICK_SENTINEL}" matches ${el}'s pattern ${pattern}, so the probe's own ` +
          `placeholder would read as that element landing`,
      ).toBe(false);
    }
  });
});
