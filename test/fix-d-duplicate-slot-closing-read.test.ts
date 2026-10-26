import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { runSelfTest } from "../src/taskpane/selftest";
import { _resetDeadlinesFiredForTest, _resetStallContextForTest } from "../src/render/powerpoint";
import { onTrace, setTracing } from "../src/core/trace";
import { faults, installHost, makeSlide } from "./helpers/office-host";

/**
 * `two slides claiming one slot`, and the closing read it takes straight after the
 * repair's deletes.
 *
 * Fifteen archived failures (of 484 traced runs, none of 469 passes) split in two.
 * Ten of them had the read-back right and both duplicates deleted, and then the
 * count and the listing, taken at once, still showed three or four slides:
 * `4 slides inserted, 3 kept, 3 of 2 still re-editable`. That is a deck that had
 * already lost two slides and had not yet said so. The scenario now takes the read
 * again once after the settle the count reads use when the repair says it deleted
 * and the read still shows more than the run set out to keep. A leftover that is
 * real is still there on the second look.
 */
const NAME = "two slides claiming one slot";

beforeEach(() => {
  _resetStallContextForTest();
  _resetDeadlinesFiredForTest();
});

afterEach(() => {
  setTracing(false);
  onTrace(undefined);
  faults.deletedSlidesLinger = 0;
  faults.refuseSlideDelete = false;
});

/** Arm a fault at the moment the scenario starts, so the probe set-up before it is untouched. */
const armWhenItStarts = (arm: () => void) => {
  setTracing(true);
  onTrace((e) => {
    if (e.message === "scenario starting" && e.data?.name === NAME) arm();
  });
};

const run = async () => {
  installHost([makeSlide("s1")]);
  const all = await runSelfTest("probe", NAME);
  return all.find((r) => r.name === NAME)!;
};

describe("the closing read of the duplicate-slot scenario", () => {
  it("passes on a clean host and reads once (the control)", async () => {
    setTracing(true);
    const result = await run();
    expect(result.ok, result.detail).toBe(true);
    expect(result.detail).toMatch(/2 kept, 2 of 2 still re-editable/);
  });

  it("takes the read again when the host still counts the deleted slides on the first look", async () => {
    // One count read late: the shape of rounds 148, 315, 375 and 422 (`4 kept, 2 of 2
    // still re-editable`). The fake lingers a deleted slide in the COUNT for a number
    // of reads, not a time, so one is what a single settle can answer; it cannot
    // model the listing lagging as well (the phantom slide is unreadable), which is
    // the `3 of 2` variant and is covered by the same re-read.
    armWhenItStarts(() => {
      faults.deletedSlidesLinger = 1;
    });
    const result = await run();
    expect(result.skipped, result.detail).toBeFalsy();
    expect(result.ok, `a host that was late to show the deletes was filed as a failure: ${result.detail}`).toBe(true);
    expect(result.detail).toMatch(/2 kept, 2 of 2 still re-editable/);
  });

  it("still FAILS when the slides really are still there", async () => {
    armWhenItStarts(() => {
      faults.refuseSlideDelete = true;
    });
    const result = await run();
    expect(result.ok, "a delete that never happened was read as the host being late").toBe(false);
    expect(result.skipped).toBeFalsy();
    expect(result.detail).toMatch(/[34] kept/);
  });
});
