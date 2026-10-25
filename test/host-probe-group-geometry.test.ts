// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { groupGeometryVerdict, runHostProbes } from "../src/render/host-probe";
import { faults, installHost, makeSlide } from "./helpers/office-host";

/**
 * `group-members-report-geometry`: does a group's collection give back each
 * member's name, width and height, and the group its own width and height.
 *
 * The update path skips any field the host does not give, so a host that gives
 * none runs the in-place guard blind — and the success line's `checked` count
 * cannot say whether that is this host. This question says, and these tests hold
 * the words it says it in: `yes`, `partial`, `no` and `short-N`, with the raw
 * reading of what did not arrive in the detail.
 */

afterEach(() => vi.unstubAllGlobals());

const ID = "group-members-report-geometry";
const member = (id: string, w: number, h: number, name = "string:probe") => ({ id, name, width: w, height: h });
const whole = [member("1", 20, 20, "string:probe-0"), member("2", 30, 25, "string:probe-1")];
const group = { width: 55, height: 25 };

describe("groupGeometryVerdict", () => {
  it("says yes when all eight readings arrive and are what was drawn, in whatever order the host lists them", () => {
    expect(groupGeometryVerdict({ asked: 2, listed: whole, group }).answer).toBe("yes");
    expect(groupGeometryVerdict({ asked: 2, listed: [...whole].reverse(), group }).answer).toBe("yes");
  });

  it("says partial when one reading never came, and names it", () => {
    const noName = [{ ...whole[0], name: "threw PropertyNotLoaded" }, whole[1]];
    const v = groupGeometryVerdict({ asked: 2, listed: noName, group });
    expect(v.answer).toBe("partial");
    expect(v.detail).toContain("names 1/2");
    expect(v.detail).toContain("threw PropertyNotLoaded");
    const noGroupWidth = groupGeometryVerdict({ asked: 2, listed: whole, group: { ...group, width: "undefined" } });
    expect(noGroupWidth.answer).toBe("partial");
    expect(noGroupWidth.detail).toContain("group width no");
  });

  it("says partial, not yes, when a reading arrives and is not what was drawn", () => {
    // A host that reports a member at some other size is not "reporting" it.
    const v = groupGeometryVerdict({ asked: 2, listed: [whole[0], member("2", 31, 25, "string:probe-1")], group });
    expect(v.answer).toBe("partial");
    expect(v.detail).toContain("arrived 8 of 8");
    expect(groupGeometryVerdict({ asked: 2, listed: whole, group: { width: 110, height: 50 } }).answer).toBe("partial");
  });

  it("says no when the members answered an id and nothing else", () => {
    const idsOnly = whole.map((m) => ({ ...m, name: "undefined", width: "undefined", height: "undefined" }));
    const v = groupGeometryVerdict({ asked: 2, listed: idsOnly, group: { width: "undefined", height: "undefined" } });
    expect(v.answer).toBe("no");
    expect(v.detail).toContain("arrived 0 of 8");
  });

  it("tells a throw from an undefined in the detail, since they are different hosts", () => {
    const threw = whole.map((m) => ({ ...m, width: "threw PropertyNotLoaded" }));
    const undef = whole.map((m) => ({ ...m, width: "undefined" }));
    expect(groupGeometryVerdict({ asked: 2, listed: threw, group }).detail).toContain("threw PropertyNotLoaded");
    expect(groupGeometryVerdict({ asked: 2, listed: undef, group }).detail).toContain('"width":"undefined"');
  });

  it("says short-N when the group lists fewer members than it holds, however much they say", () => {
    expect(groupGeometryVerdict({ asked: 2, listed: [whole[0]], group }).answer).toBe("short-1");
    expect(groupGeometryVerdict({ asked: 2, listed: [], group }).answer).toBe("short-0");
  });

  it("does not take the empty string for a name", () => {
    const v = groupGeometryVerdict({ asked: 2, listed: whole.map((m) => ({ ...m, name: "string:" })), group });
    expect(v.answer).toBe("partial");
    expect(v.detail).toContain("names 0/2");
  });
});

describe("the question, asked of the fake host", () => {
  const ask = async () => {
    const sheet = await runHostProbes("fake", "test", { only: [ID], passes: 1 });
    return sheet.answers.find((a) => a.id === ID)!;
  };

  it("answers yes on a host that gives everything, and says the members were reached by id", async () => {
    installHost([makeSlide("s1")]);
    const a = await ask();
    expect(a.answer).toBe("yes");
    expect(a.detail).toContain("members via ids");
    expect(a.detail).toContain("arrived 8 of 8");
  });

  it("answers no on a host whose members give an id and nothing else, and the group only its place", async () => {
    installHost([makeSlide("s1")]);
    faults.groupMembersSilentOn = ["name", "width", "height"];
    faults.groupSilentOn = ["width", "height"];
    const a = await ask();
    expect(a.answer).toBe("no");
    expect(a.detail).toContain("threw PropertyNotLoaded");
    expect(a.detail).toContain("arrived 0 of 8");
    // The ids still came: the members were listed, which is what `short-N` is about.
    expect(a.detail).toContain("listed 2 of 2");
  });

  it("tells a load that lands as undefined from one that throws", async () => {
    installHost([makeSlide("s1")]);
    faults.groupMembersSilentOn = ["name", "width", "height"];
    faults.groupSilentOn = ["width", "height"];
    faults.silentReadsAs = "undefined";
    const a = await ask();
    expect(a.answer).toBe("no");
    expect(a.detail).not.toContain("PropertyNotLoaded");
    expect(a.detail).toContain('"name":"undefined"');
  });

  it("answers partial for each field a host can withhold on its own", async () => {
    for (const [silent, group] of [
      [["name"], []],
      [["width"], []],
      [["height"], []],
      [[], ["width"]],
      [[], ["height"]],
    ] as const) {
      installHost([makeSlide("s1")]);
      faults.groupMembersSilentOn = [...silent];
      faults.groupSilentOn = [...group];
      const a = await ask();
      expect(a.answer, `withholding members ${silent} / group ${group}`).toBe("partial");
    }
  });

  it("answers threw when the host refuses the group read, as it does for the neighbouring questions", async () => {
    installHost([makeSlide("s1")]);
    faults.refuseGroupRead = true;
    const a = await ask();
    expect(a.answer).toBe("threw");
  });

  it("is on the shortlist a run under slide pressure falls back to, because nothing has answered it yet", async () => {
    const { RESAMPLE_IDS } = await import("../src/render/host-probe");
    expect(RESAMPLE_IDS).toContain(ID);
  });
});
