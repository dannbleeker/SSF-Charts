# Rounds that ran, and were never filed

Complete rounds recovered from `crashes/` by `scripts/salvage-crashed.mjs`.
Same shape as `rounds/`, and **deliberately not in it**.

## Where they came from

Every verdict of a round is in before `collectDeckEvidence` runs, and that scan
is where this host dies — 9 builds against 4 on the per-phase traces, in a band
of 441-572s. Until 2026-08-29 the pane assembled its run log _after_ that scan,
so a crash there left a complete round with nothing to file it from: the driver
found no log to download and wrote a crash record instead.

33 of 49 crash records hold every verdict. 22 of those can also say, from the
host's own `pageSetup` reading, what slide size they ran at. Those 22 are here.

`src/taskpane/app.ts` now banks the log _before_ the scan, so this should stop
accruing.

**THOSE NUMBERS ARE FROM THE FIRST SALVAGE PASS AND ARE STALE**, corrected
2026-09-07 after they were quoted into a `docs/BACKLOG.md` argument. `619d5a1`
on 2026-09-03 — "The salvage was not broken — its bars were refusing 47 real
rounds" — widened what could be recovered, and this directory grew. As it
stands:

    files                          69
    distinct runs                  68   041 and 043 are the same round,
                                        salvaged from two crash records five
                                        hours apart on 2026-08-29
    complete                       50
    partial (`salvagedPartial`)    18   reach as low as 3 of 14
    verified 4:3                   32
    verified 16:9                   7
    no verified size               29

So **only 39 of the 68 can enter an aspect arm at all**, and 18 of those must be
conditioned on reach before any rate is computed — the rule `salvage-crashed.mjs`
states in its own header, and which `BACKLOG.md` broke twice before anyone
checked.

## Why they are not in `rounds/`

`rounds/README.md` opens with the contract: `NNN-<build>.json`, **oldest first**.
The number is the chronology. Every ordering-sensitive instrument leans on it —
`poolFallbackRates` compares the oldest third against the newest third, and
"this round" everywhere means the highest number.

These are from 2026-08-26 to 08-29. Filed at the end they would take 311-332 and
claim to be the newest rounds in the archive. They cannot be interleaved instead:
**no archived round records when it ran**, so there is nothing to interleave
against.

Measured on a scratch copy rather than assumed — adding all 22 took
`poolGroupingOutcome` to `null` and `poolFallbackRates` from 5 signals to 0,
because the newest round became one carrying no trace. That is the stale-`now`
guard being honest, and it is also the gate going quiet about both its headline
numbers.

## What is missing from each file, and why it is missing rather than guessed

- **`deck`** — the scan never returned. An invented inventory is exactly the
  fabricated evidence `crashes/README.md` warns about, and everything downstream
  reads that field as measured.
- **`trace`** — the crash keeps `steps` as formatted strings; a round's trace is
  structured entries with `message` and `data`. Reshaping prose into structure
  would be the same lie in a different shape.

Both absences make the pooling functions skip these rounds, which is only safe
because the stale-`now` class was fixed first — see `isTheRoundBeingJudged` in
`scripts/triage.mjs`.

Each file carries `salvagedFrom`, naming the crash record it came from.

## What they are good for

Any question that does not depend on order: what a build's verdicts were, how
often a scenario has ever failed, whether the 4:3 arm agrees with 16:9.

**All 22 are 4:3**, which is itself the finding — the validation leg runs last
and deepest into a session, so it is the one that keeps dying in the scan. It
takes that arm from 26 rounds to 48.

**STALE AS ABOVE, and the "26 to 48" is the dangerous half.** Of the 68 runs
here, 32 are verified 4:3, 7 are 16:9 and 29 have no verified size — so "all
are 4:3" is no longer true, and the direction of the finding (the validation leg
dies deepest) survives while the count does not. The 26 + 22 = 48 arithmetic
also happens to equal the archived PRE 4:3 round count, which is a coincidence
and a live double-count hazard for anyone reading this beside `BACKLOG.md`.
`rounds/` is clean: no file in it carries `salvagedFrom`, checked.

Pooling them properly needs an ordering key that does not exist yet.
