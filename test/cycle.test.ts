import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import * as cycle from "../scripts/cycle.mjs";
const { cyclePlan, nextStep, readReceipt } = cycle;

/**
 * The cycle runner's whole job is knowing when to stop, so stopping is what is
 * worth testing. A runner that carries on past a regression wastes a night; one
 * that stops on an ordinary failed scenario throws away the second half of a
 * pair, which is the only thing that separates a real fault from this host's
 * 1-in-5 mood.
 */
describe("a night's cycle", () => {
  const receipt = (over = {}) => ({
    reason: "finished",
    codes: [],
    recoverable: false,
    roundFile: "082-x.json",
    ...over,
  });

  it("runs 16:9 twice and 4:3 once, in that order", () => {
    const plan = cyclePlan();
    expect(plan.map((p: { size: string }) => p.size)).toEqual(["16:9", "16:9", "4:3"]);
    // 4:3 LAST: a night that dies early should lose the validation, not the
    // measurement.
    expect(plan[plan.length - 1].size).toBe("4:3");
    // Different decks — a 4:3 leg on the 16:9 deck would file a round under a
    // profile it was not measured at.
    expect(plan[0].deck).not.toBe(plan[2].deck);

    // THE NAMES, PINNED, because "they differ" is what this test used to check
    // and it passed happily while the 4:3 default named `Presentation66` — a
    // deck that no longer existed. An unattended night would have refused its
    // whole validation leg with `deck-missing`.
    //
    // A unit test cannot know what is open in a browser, so this pin only makes
    // a change DELIBERATE. The live check is `node scripts/round.mjs --check`,
    // which now prints the fronted document and its measured slide size.
    expect(plan[0].deck).toBe("Presentation64");
    // `Presentation70` since 2026-08-26. `Presentation67` was still open and
    // still correctly named — what had gone stale was its SLIDE SIZE, reset to
    // 16:9, which stops the 4:3 leg with `wrong-size` and is the one thing
    // recovery is forbidden to fix. The name alone could not have told anyone.
    expect(plan[2].deck).toBe("Presentation70");
  });

  it("carries on after a round whose scenarios failed", () => {
    // A failing scenario is the measurement WORKING. Stopping here would throw
    // away the pair.
    expect(nextStep({ exitCode: 0, receipt: receipt(), gateStatus: 0 }).go).toBe(true);
  });

  it("stops dead on a regression, which is the one fatal check", () => {
    const step = nextStep({ exitCode: 0, receipt: receipt(), gateStatus: 1 });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/WAS passing/);
  });

  it("blames the refusal, not a gate verdict about a round this leg never ran", () => {
    /**
     * Cycle three on 2026-09-08 archived NOTHING — the round refused as
     * `deep-session`, round 5 of a session, where scenarios start being
     * skipped. The gate then ran on an unchanged archive, found round 435's
     * standing regression, and the cycle announced "a scenario that WAS passing
     * has stopped — it wants reading now". True of round 435 and nothing to do
     * with why this cycle stopped; what the reader needed was "rest 45 minutes".
     *
     * Third time this file has misnamed its own stop. The pattern never varies:
     * a signal that is TRUE gets reported as the CAUSE.
     */
    const refused = {
      reason: "not-ready",
      codes: ["pane-stale", "deep-session"],
      roundFile: null,
      recoverable: false,
    };
    const step = nextStep({ exitCode: 1, receipt: refused, gateStatus: 1 });
    expect(step.go).toBe(false);
    expect(step.why, "the stop did not name the reason nothing ran").toMatch(/deep-session/);
    expect(step.why, "it did not say that nothing was archived").toMatch(/nothing was archived/);
    // THE GATE IS NOT SUPPRESSED. A standing regression still matters — it is
    // just not why tonight was lost, and a message that dropped it would trade
    // one silence for another.
    expect(step.why, "the standing regression was swallowed").toMatch(/PREVIOUS round/);

    // AND WHEN THE ARCHIVE DID GROW, the gate's verdict IS the news.
    const archived = nextStep({ exitCode: 0, receipt: receipt(), gateStatus: 1 });
    expect(archived.why, "a fresh regression stopped naming itself").toMatch(/WAS passing/);
  });

  it("names the host-death check by itself, instead of blaming a verdict that did not fall", () => {
    // Round 428: 19 of 19 scenarios green, the gate red because `what a chart
    // kind costs` had killed PowerPoint for the first time in 35 runs. Both
    // fatal checks exited 1, so the night stopped saying "a scenario that WAS
    // passing has stopped" — and its reader went looking for a failed verdict
    // that was not in the round file.
    const step = nextStep({ exitCode: 0, receipt: receipt(), gateStatus: 3 });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/KILLING THE HOST/);
    // AND IT MUST NOT ALSO SAY THE OTHER THING. A message naming both checks is
    // the same failure with more words — the reader still cannot tell which one
    // fired.
    expect(step.why, "the host-death stop still claims a scenario stopped passing").not.toMatch(/WAS passing/);
  });

  it("stops when a leg finished but filed nothing", () => {
    // A ROUND THAT FINISHED IS NOT A ROUND THAT WAS FILED. `attempt` returns 0
    // when the pane says the run is done; archiving happens after and is
    // best-effort. Sailing past that costs the round twice: the gate re-judges
    // the PREVIOUS round and passes, so the night reads healthy, and the next
    // leg's download overwrites the only copy of the log.
    //
    // Exactly the state the archive-ENOENT bug produced. This is the layer that
    // should have caught it and did not.
    const step = nextStep({
      exitCode: 0,
      receipt: receipt({ reason: "finished", roundFile: null }),
      gateStatus: 0,
    });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/nothing was archived/);
  });

  it("does not mistake a --check for a round that lost its evidence", () => {
    // `--check` exits 0 and archives nothing by design, so the stop above is
    // keyed on the reason rather than on the exit code alone.
    expect(nextStep({ exitCode: 0, receipt: receipt({ reason: "checked", roundFile: null }), gateStatus: 0 }).go).toBe(
      true,
    );
  });

  it("does not call a gate that could not judge a regression", () => {
    // The two non-zero exits are opposite findings. Exit 2 is the gate saying it
    // could not read the archive at all, and reporting that as a fall sends
    // someone hunting a regression that never happened — which is exactly what
    // an interrupted write to a round file used to do, node exiting 1 on a
    // SyntaxError with the night stopping to blame the build.
    const step = nextStep({ exitCode: 0, receipt: receipt(), gateStatus: 2 });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/not a regression/);
    expect(step.why, "sent the reader after a scenario that never fell").not.toMatch(/WAS passing/);
  });

  it("stops on a refusal recovery does not address, and says it needs a person", () => {
    const step = nextStep({
      exitCode: 1,
      receipt: receipt({ reason: "not-ready", codes: ["wrong-size"], recoverable: false, roundFile: null }),
      gateStatus: 0,
    });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/needs a person/);
    expect(step.why).toMatch(/wrong-size/);
  });

  it("does not retry what the driver already retried", () => {
    // By the time a recoverable refusal reaches here the driver has exhausted
    // its own retries. Another attempt from this layer would be a second
    // implementation of shouldRetry, and there must only be one.
    const step = nextStep({
      exitCode: 1,
      receipt: receipt({ reason: "not-ready", codes: ["pane-stale"], recoverable: true, roundFile: null }),
      gateStatus: 0,
    });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/retried and still could not clear/);
  });

  it("stops when the driver left no account of itself at all", () => {
    const step = nextStep({ exitCode: 1, receipt: null, gateStatus: 0 });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/no \.round-outcome\.json/);
  });

  it("treats an unreadable receipt as no receipt rather than crashing on it", () => {
    expect(
      readReceipt(
        "x",
        () => true,
        () => "{ not json",
      ),
    ).toBeNull();
    expect(readReceipt("x", () => false)).toBeNull();
  });

  it("runs every leg but the first with --fresh", () => {
    // THE PAIR DISCIPLINE AND THE FRESH-PANE DISCIPLINE WERE IN DIRECT CONFLICT,
    // and `npm run cycle` is the only sanctioned way to run a pair, so the
    // conflict lived here: the brief has said "run the second round with
    // `--fresh`, not optional" since 2026-08-22, and this script never passed it.
    //
    // Pane age separates post-retry 0.4 from 4.6, and nothing in the driver
    // freshens the pane between rounds — the reload was removed because it
    // raises a beforeunload prompt over unsaved work and cost the sideload after
    // rounds 124 and 132. What freshened it was a MERGE, and a cycle runs its
    // legs back to back with no merge. Rounds 159 and 161 started on panes of
    // 696s and 666s and were both the worse half of their pair.
    const { roundArgs } = cycle as unknown as {
      roundArgs: (leg: { leg: number }, dir: string, retry: string) => string[];
    };
    expect(
      roundArgs({ leg: 1 }, ".pw", "6"),
      "the FIRST leg follows a merge, so its pane is already fresh",
    ).not.toContain("--fresh");
    expect(roundArgs({ leg: 1 }, ".pw", "6").slice(0, 5)).toEqual([
      "scripts/round.mjs",
      "--dir",
      ".pw",
      "--retry",
      "6",
    ]);
    /**
     * AND FOR THE SAME REASON IT WAITS FOR PAGES: leg 1 follows a merge, so it
     * is the leg whose deploy may not have landed yet.
     *
     * `driverRun.waitedForDeploy` was true in 37 of 195 archived rounds, all
     * between 166 and 240, and false in every one of the 104 since that carry
     * the field — because this function never passed the flag, while the driver
     * had implemented the wait the whole time.
     *
     * The readiness gate already keeps stale builds out, and no archived round
     * has ever run against one. What that refusal COSTS is the issue:
     * `site-behind` is deliberately not recoverable, so the leg stops and takes
     * the cycle with it, and a stop leaves no archived trace behind. Waiting is
     * the right answer to a deploy that has not landed; retrying is not, and
     * ending the night is not either.
     */
    expect(
      roundArgs({ leg: 1 }, ".pw", "6"),
      "the leg that follows a merge does not wait for the deploy it is testing",
    ).toContain("--wait-for-deploy");
    expect(roundArgs({ leg: 2 }, ".pw", "6")).toContain("--fresh");
    // Legs 2 and 3 run back to back after leg 1, so the deploy has necessarily
    // landed and waiting again would only spend the night's time.
    expect(roundArgs({ leg: 2 }, ".pw", "6"), "a later leg waited for a deploy that had already landed").not.toContain(
      "--wait-for-deploy",
    );
    expect(roundArgs({ leg: 3 }, ".pw", "6")).not.toContain("--wait-for-deploy");
    // LEG 3 TOO. It is a second-in-sequence round with the same aged pane; only
    // its deck differs, and the 4:3 validation is compared against the 16:9
    // pair, so letting it run on a stale pane reintroduces the confound the
    // other two just removed.
    expect(roundArgs({ leg: 3 }, ".pw", "6"), "the 4:3 leg inherits an aged pane like any other").toContain("--fresh");
  });
});
