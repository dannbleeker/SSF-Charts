import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types.
import { nextRoundNumber } from "../scripts/round.mjs";
// @ts-expect-error — a plain .mjs tool with no types.
import { loadRounds } from "../scripts/rounds-gate.mjs";

/**
 * THE ARCHIVE DOES NOT STOP AT 999.
 *
 * Found 2026-09-28, at round 462. Every reader of the round archive matched
 * `/^\d{3}-/` — exactly three digits. That pattern does not match
 * "1000-abc.json" AT ALL: the fourth character is a digit where the pattern
 * needs its dash. So at round 1000 the driver would go on writing files that
 * `nextRoundNumber`, `loadRounds`, triage, salvage and `rounds.test.ts` all
 * stopped seeing — at the same moment, with nothing failing anywhere. The
 * archive would split in two and every reading would be taken from the older
 * half.
 *
 * And a second half to the same trigger: the readers sorted filenames with a
 * bare `.sort()`, under which "1000-x.json" comes BEFORE "999-x.json" because
 * "1" < "9". Widening the pattern alone would only have moved the breakage —
 * the newest round would be found, and then sorted into the middle.
 *
 * The PAD deliberately stays three wide: `padStart` never truncates, so 29
 * keeps its "029" and 1000 is written "1000". Widening it would rename every
 * file already archived.
 */
describe("round numbers past 999", () => {
  it("counts on from a four-digit round", () => {
    expect(
      nextRoundNumber(["0998-x.json", "999-x.json"]),
      "a three-digit-only pattern cannot see a four-digit round, so the driver would reuse 1000 forever",
    ).toBe("1000");
    expect(nextRoundNumber(["999-x.json", "1000-x.json", "1001-x.json"])).toBe("1002");
  });

  it("still pads a small number to three, and never truncates a large one", () => {
    expect(nextRoundNumber([])).toBe("001");
    expect(nextRoundNumber(["028-x.json"])).toBe("029");
    expect(nextRoundNumber(["1234-x.json"])).toBe("1235");
  });

  it("orders the archive by round number, not by filename", () => {
    // A bare `.sort()` puts 1000 first, so the LAST entry — which every
    // "the newest round says…" reading depends on — would be 999.
    const files = ["998-a.json", "999-b.json", "1000-c.json", "1001-d.json"];
    const rounds = loadRounds(
      "rounds",
      () => files,
      (p: string) => JSON.stringify({ build: String(p).split("/").pop() }),
    ) as { build: string }[];
    expect(rounds, "loadRounds returned nothing — re-point this test").toHaveLength(4);
    expect(
      rounds[rounds.length - 1].build,
      `the newest round is not last — the archive came back ordered ${rounds.map((r) => r.build).join(", ")}`,
    ).toBe("1001-d.json");
  });

  it("still reads a purely three-digit archive exactly as before", () => {
    // The whole archive today is three digits. This fix must be invisible to it.
    const files = ["001-a.json", "029-b.json", "462-c.json"];
    const rounds = loadRounds(
      "rounds",
      () => files,
      (p: string) => JSON.stringify({ build: String(p).split("/").pop() }),
    ) as { build: string }[];
    expect(rounds.map((r) => r.build)).toEqual(["001-a.json", "029-b.json", "462-c.json"]);
  });
});
