#!/usr/bin/env node
/**
 * Run a night's rounds without a person between them.
 *
 *     npm run cycle
 *
 * Two rounds at 16:9 and one at 4:3, which is the schedule agreed with the repo
 * owner: 16:9 is the measurement, 4:3 is VALIDATION. The noise floor here is
 * 1-versus-5 for the same fault with nothing changed, so a single round is never
 * evidence and the pair is the unit — `docs/ROUNDS.md` says two per build
 * minimum, three where a claim depends on it.
 *
 * WHAT THIS DOES NOT DO is as important as what it does.
 *
 * It does not decide whether a round found something. It does not decide whether
 * a divergence is real. It does not set a slide size — `RECOVERABLE_STOPS`
 * excludes `wrong-size` precisely because changing the size would change what
 * the round measures rather than restore it, and a cycle runner that "helpfully"
 * fixed the profile would be doing the one thing recovery is forbidden to do.
 * And it never goes near sign-in: that needs a password, none of that is the
 * agent's to do, and the driver returns before any of this on that path.
 *
 * It also does not reimplement `shouldRetry`. Retrying is the driver's job and
 * there must be exactly one implementation of "is this worth another attempt";
 * this reads the receipt the driver leaves and stops when the driver gave up.
 */
import { spawnSync } from "child_process";
import { readFileSync, existsSync, rmSync } from "fs";
import { isMain } from "./is-main.mjs";
// Its own line: the grouped-import trap is documented in test/triage.test.ts and
// has been paid for four times.
import { RECEIPT_PATH } from "./round.mjs";

/**
 * The night's legs, in order.
 *
 * 4:3 LAST and only once. It is the cheap validation arm, and putting it first
 * would mean a night that dies early loses the measurement rather than the
 * check. `PW_EXPECT_SIZE` is what makes the driver refuse a deck in the wrong
 * profile — without it a 4:3 leg run against a 16:9 deck files a round under a
 * profile it was not measured at, and every later comparison inherits that.
 */
// `tall` IS THE DECK THAT EXISTS, and this default has already been wrong once.
//
// It said `Presentation66` until 2026-08-20, while the deck actually created,
// sideloaded and left open for the 4:3 leg was `Presentation67`. An unattended
// cycle would have fronted a tab that is not there and refused the whole leg
// with `deck-missing` — which is at least loud, but the night would still have
// lost its validation arm. The driver's own comments record the same trap one
// layer down (`recover` hard-coded `Presentation63` after the deck became
// `Presentation64`). A deck name is not a constant; it is a fact about the
// browser, and it goes stale every time a document is remade.
// `Presentation70` SINCE 2026-08-26, and this is the second time this default
// has gone stale — the comment above records the first. `Presentation67` was
// still open and still named correctly, so nothing looked broken; what had
// changed is that its slide size had been reset to 16:9, and a 4:3 leg against
// it stops the whole cycle with `wrong-size`, which recovery is forbidden to fix.
//
// A deck name is not the only fact that goes stale here. Its SIZE does too, and
// the name alone cannot tell you. `Presentation70` was created at 4:3 for this.
export function cyclePlan({ wide = "Presentation64", tall = "Presentation70" } = {}) {
  return [
    { leg: 1, deck: wide, size: "16:9", why: "the measurement" },
    { leg: 2, deck: wide, size: "16:9", why: "the pair — one round is never evidence" },
    { leg: 3, deck: tall, size: "4:3", why: "validation" },
  ];
}

/**
 * What to launch each leg with — `--fresh` on every leg but the first.
 *
 * THE PAIR DISCIPLINE AND THE FRESH-PANE DISCIPLINE WERE IN DIRECT CONFLICT,
 * and this is the only sanctioned way to run a pair, so the conflict lived
 * here. Measured 2026-08-22: pane age separates post-retry 0.4 from 4.6, and
 * NOTHING IN THE DRIVER FRESHENS THE PANE between rounds — the reload was
 * removed because it raises a beforeunload prompt over unsaved work and has
 * cost the sideload twice, after rounds 124 and 132. What actually freshened it
 * was a MERGE, which makes the pane stale so recovery reloads it.
 *
 * A cycle runs its legs back to back with no merge between them, so every leg
 * after the first inherited a pane hundreds of seconds old. Rounds 159 and 161
 * started at 696s and 666s and were both the worse half of their pair; the four
 * fresh-pane rounds beside them refused no group at all.
 *
 * `--fresh` closes the BROWSER, which the persistent profile survives, and costs
 * one extra attempt while recovery reopens the pane from the ribbon. Leg 3 gets
 * it too: it is a second-in-sequence round with the same aged pane, and only its
 * deck differs.
 */
export function roundArgs(leg, dir, retry) {
  const args = ["scripts/round.mjs", "--dir", dir, "--retry", retry];
  if (leg.leg > 1) args.push("--fresh");
  /**
   * AND LEG 1 WAITS FOR PAGES, because leg 1 is the one that follows a merge.
   *
   * `driverRun.waitedForDeploy` was true in 37 of 195 archived rounds, all
   * between 166 and 240, and false in every one of the 104 since that carry the
   * field — because this function never passed the flag, and this is the only
   * sanctioned way to run a pair. The driver has implemented the wait the whole
   * time.
   *
   * The readiness gate is not the problem: it refuses `site-behind` when the
   * site is not serving HEAD, and no archived round has ever run against a stale
   * deployment. The problem is what that refusal COSTS. `site-behind` sits
   * outside the recoverable set on purpose — `round.mjs` says a reload does not
   * make Pages deploy faster, and it is right — so the leg stops and takes the
   * cycle with it, and a stop leaves no archived trace behind. Waiting is the
   * correct answer to "the deploy has not landed yet". Retrying is not, and
   * ending the night is not either.
   *
   * Legs 2 and 3 do not need it: they run back to back after leg 1, by which
   * time the deploy under test has necessarily landed.
   */
  if (leg.leg === 1) args.push("--wait-for-deploy");
  return args;
}

/** The driver's account of the round that just ran, or null if it left none. */
export function readReceipt(path = RECEIPT_PATH, exists = existsSync, read = readFileSync) {
  if (!exists(path)) return null;
  try {
    return JSON.parse(read(path, "utf8"));
  } catch {
    return null;
  }
}

/**
 * Carry on, or stop and fetch a person?
 *
 * The three stopping conditions are different in kind and the messages have to
 * keep them apart, because the response to each is different: a build that
 * regressed wants reading, a refusal the driver could not clear wants a hand on
 * the machine, and a missing receipt means something went wrong beneath both.
 *
 * A round that FINISHED always continues, even when its scenarios failed. A
 * scenario failing is the measurement working — stopping the night on it would
 * throw away the second half of the pair, which is the only thing that could
 * tell a real fault from this host's 1-in-5 mood.
 */
/**
 * Did this receipt come from a round that RAN, as opposed to one that refused?
 *
 * `--check` also exits 0, and it archives nothing by design. Keying the
 * archived-nothing stop on the reason keeps it aimed at the case that matters.
 */
function reasonFinished(receipt) {
  return receipt?.reason === "finished";
}

export function nextStep({ exitCode, receipt, gateStatus }) {
  // TWO WAYS FOR THE GATE TO EXIT NON-ZERO, and they are opposite findings. 1 is
  // the thing it exists to say: a scenario that was passing has stopped. 2 is
  // the gate reporting that it could not judge at all — an unreadable archive,
  // an empty one — and calling that a regression would send someone hunting a
  // fall that never happened. Before the gate had its own code for this, an
  // interrupted write to a round file did exactly that: node exited 1 on a
  // SyntaxError and the night stopped blaming the build.
  if (gateStatus === 2) {
    return {
      go: false,
      why: "the gate could not judge this round — nothing was checked, and that is not a regression",
    };
  }
  // TWO FATAL CHECKS, TWO MESSAGES. Everything non-zero used to say "a scenario
  // that WAS passing has stopped", and round 428 showed what that costs: 19 of
  // 19 scenarios green, the gate red on a host DEATH, and a stop line sending
  // its reader to hunt a failed verdict that was not there. A stop that
  // misnames its own cause is worse than a terse one, because it is acted on.
  if (gateStatus === 3) {
    return {
      go: false,
      why:
        "a scenario is KILLING THE HOST more often than it did — no verdict fell, so read the death rate " +
        "the gate printed, not the scenario list",
    };
  }
  if (gateStatus !== 0) {
    return {
      go: false,
      why: "a scenario that WAS passing has stopped — that is the one fatal check, and it wants reading now",
    };
  }
  if (!receipt) {
    return { go: false, why: `the driver left no ${RECEIPT_PATH} — it did not reach the end of its own run` };
  }
  // A ROUND THAT FINISHED IS NOT A ROUND THAT WAS FILED. `attempt` returns 0
  // the moment the pane says the run is done; archiving happens after, is
  // best-effort by design, and sets `roundFile` only when a file was written.
  //
  // Sailing past that costs the round twice over: the gate re-judges the
  // PREVIOUS round and passes, so the night reads as healthy, and the next
  // leg's download overwrites `.playwright-cli/ssf-charts-run-log.json` — the
  // only copy of the evidence. That is exactly the state the archive-ENOENT
  // bug produced, and this is the layer that should have caught it.
  if (exitCode === 0 && reasonFinished(receipt) && !receipt.roundFile)
    return {
      go: false,
      why: "the round finished but nothing was archived — its log is still in the session directory and the next leg would overwrite it",
    };
  if (exitCode === 0) return { go: true, why: null };
  // The driver has ALREADY retried everything recovery addresses by this point.
  // A non-zero exit carrying recoverable codes means the retries ran out, not
  // that another one would help.
  const codes = receipt.codes?.length ? ` (${receipt.codes.join(", ")})` : "";
  return {
    go: false,
    why: receipt.recoverable
      ? `the driver retried and still could not clear ${receipt.reason}${codes}`
      : `${receipt.reason}${codes} is not something recovery addresses — it needs a person`,
  };
}

if (isMain(import.meta.url, process.argv[1])) {
  const dirArg = process.argv.indexOf("--dir");
  const dir = dirArg === -1 ? ".pw-session" : process.argv[dirArg + 1];
  const retryArg = process.argv.indexOf("--retry");
  const retry = retryArg === -1 ? "6" : process.argv[retryArg + 1];
  const plan = cyclePlan({
    wide: process.env.PW_DECK_16_9 ?? undefined,
    tall: process.env.PW_DECK_4_3 ?? undefined,
  });

  console.log(`  a cycle of ${plan.length} round(s): ${plan.map((p) => p.size).join(", ")}\n`);
  let ran = 0;
  for (const leg of plan) {
    console.log(`\n  ROUND ${leg.leg} of ${plan.length} — ${leg.deck} at ${leg.size}, ${leg.why}`);
    // A CHILD PROCESS PER ROUND, not an imported function. A round that wedges
    // its own process cannot then poison the next one, and the per-round
    // environment is the natural way to say which arm this is.
    // CLEAR IT FIRST, or the check below reads the LAST round's outcome. A leg
    // that dies before writing leaves the previous leg's receipt on disk, and
    // nothing distinguishes the two — which made the "the driver left no
    // receipt" branch, written for exactly this, unreachable from leg 2 onward.
    // A stale `finished` would have been read as this round's success.
    try {
      rmSync(RECEIPT_PATH, { force: true });
    } catch {
      /* nothing to clear is the normal case on the first leg */
    }
    // `--fresh` ON EVERY LEG BUT THE FIRST, which is the whole point of a pair
    // and was missing from the only sanctioned way to run one.
    //
    // Measured 2026-08-22. The pane is what separates post-retry 0.4 from 4.6,
    // and NOTHING IN THE DRIVER FRESHENS IT between rounds — the reload was
    // removed because it raises a beforeunload prompt over unsaved work and has
    // cost the sideload twice. What actually freshened it was a MERGE, which
    // makes the pane stale so recovery reloads it. So the pair discipline and
    // the fresh-pane discipline were in direct conflict: a cycle runs its legs
    // back to back with no merge, and every leg after the first inherited a
    // pane hundreds of seconds old. Rounds 159 and 161 started at 696s and 666s
    // and were both the worse half of their pair.
    //
    // `--fresh` closes the BROWSER, which the persistent profile survives, and
    // costs one extra attempt while recovery reopens the pane from the ribbon.
    // The brief has said "not optional" since that measurement; this is what
    // makes `npm run cycle` obey it.
    const r = spawnSync(process.execPath, roundArgs(leg, dir, retry), {
      stdio: "inherit",
      env: { ...process.env, PW_DECK: leg.deck, PW_EXPECT_SIZE: leg.size },
    });
    const receipt = readReceipt();
    // The gate runs after EVERY round rather than once at the end, so a
    // regression stops the night at the round that caused it instead of being
    // found three rounds later against a changed deck.
    const gate = spawnSync(process.execPath, ["scripts/rounds-gate.mjs"], { stdio: "inherit" });
    const step = nextStep({ exitCode: r.status, receipt, gateStatus: gate.status });
    if (receipt?.roundFile) ran++;
    if (!step.go) {
      console.error(`\n  CYCLE STOPPED after ${ran} archived round(s): ${step.why}`);
      console.error("  Nothing here is fixed by running it again — see docs/ROUNDS.md.");
      process.exit(1);
    }
  }
  console.log(`\n  cycle done — ${ran} round(s) archived. The rounds are the evidence; read them.`);
  process.exit(0);
}
