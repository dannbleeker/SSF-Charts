import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
// @ts-expect-error — plain .mjs tool, no types.
import * as driver from "../scripts/round.mjs";
// A NAMESPACE import, and destructured below, because the directive only covers
// the line that carries `from`. A named-import list this long is wrapped by
// prettier onto ten lines, which moves `from` five lines down, silently uncovers
// the import and turns the directive itself into an "unused directive" error.
// This is the grouped-import trap in `triage.test.ts` reached from the other
// side: there the directive covered too much, here formatting moved it off.
const {
  readiness,
  buildOf,
  nextRoundNumber,
  stripImages,
  cliEntry,
  sessionDir,
  sessionPosition,
  sessionDepthWarning,
  SESSION_DEPTH_WARN_AT,
  readSession,
  SESSION_PATH,
  RECEIPT_PATH,
  readReceipt,
  ensureSessionDir,
  spawnFailureRemedy,
  pingScript,
  readPing,
  refFor,
  frontedDeck,
  archivableDeck,
  roundConfigArg,
  waitForRef,
  refreshPane,
  namePresent,
  setSlideSizeScript,
  readSlideSize,
  sawCrashDialog,
  quietStreak,
  archive,
  signedOut,
  signInIsPopup,
  shouldRetry,
  recover,
  cli,
  isOverflow,
  recoveryFor,
  noBrowser,
  profileHeldByOrphan,
  endOrphanCommand,
  endOrphanedBrowser,
  describePane,
  paneAnsweredNothing,
  browserDiedMidRound,
  onlyDirtyDeck,
  DEAD_BROWSER_POLLS,
  slideResolveScript,
  readSlideResolve,
  slideSizeScript,
  outcomeReceipt,
  RECOVERABLE_STOPS,
  selectDeck,
  sideloadAddIn,
  sweepDeck,
  DECK_HOME_URL,
} = driver;

const READY = { head: "abc1234", deployed: "abc1234", stamp: "abc1234", slides: 1, verbose: true, pictures: true };

/**
 * The driver's whole value is refusing to start a round that cannot prove
 * anything, so the refusals are what is worth testing. A round that runs on the
 * wrong build is worse than no round: it produces a file that looks like
 * evidence and is not, and nothing downstream can tell.
 */
describe("deciding whether a round is worth running", () => {
  const ready = READY;

  it("runs when the site, the pane and HEAD all agree and the deck is clean", () => {
    expect(readiness(ready).ok).toBe(true);
    expect(readiness({ ...ready, slideOk: true }).ok, "a resolved slide changes nothing").toBe(true);
  });

  it("refuses a host that answers the cheap call but will not resolve a slide", () => {
    // THE 2s CRASH, moved to where it is cheap. Four rounds running, PowerPoint
    // crashed two seconds into the FIRST attempt and never into one that followed
    // a recovery, and its own log named the call: `OnServerFindSucceeded could
    // not find target slide`. The ping cannot see that state — `getCount` is a
    // count, and this host answers it in single-digit ms while in it.
    const r = readiness({ ...ready, ping: { answered: true, ms: 3 }, slideOk: false });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("would not resolve slide 1");
    // And it must not fire on the state the ping already covers, or the two
    // messages contradict each other on the same sheet.
    expect(readiness({ ...ready, ping: { answered: false, ms: 8011 }, slideOk: null }).stop).toHaveLength(1);
  });

  it("asks for a slide by position, inside a budget", () => {
    const script = slideResolveScript(20000);
    expect(script, "the call the host dies on, not a count").toContain("getItemAt(0)");
    expect(script, "no budget means a wedged host hangs the check forever").toContain("20000");
    expect(readSlideResolve("slide:287#62081387")).toBe(true);
    expect(readSlideResolve("slide-failed:GeneralException")).toBe(false);
    // Silence is not a refusal: an empty answer means the eval never landed, and
    // calling that a refused slide would refuse rounds on a healthy host.
    expect(readSlideResolve("")).toBe(null);
  });

  it("refuses when the site has not published HEAD yet", () => {
    // The one that looks harmless. Pages lags a merge by a minute or two, and a
    // round started inside that window measures the PREVIOUS build while
    // appearing to test the new one.
    const r = readiness({ ...ready, deployed: "0000000" });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("Deploy Pages");
  });

  it("refuses when the pane is older than the site, and says how to clear it", () => {
    // `Cache-Control: max-age=600` on the pane HTML. Reopening the pane does not
    // clear it; the whole PowerPoint tab has to be hard-reloaded. Round 24 was
    // nearly run this way.
    const r = readiness({ ...ready, stamp: "0000000" });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toMatch(/hard-reload/);
  });

  it("refuses a deck that still holds the last round's slides", () => {
    // Not a correctness problem, a comparability one: rounds 24 and 25 differed
    // only in this, and were compared as though they did not.
    const r = readiness({ ...ready, slides: 7 });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("7 slides");
  });

  it("refuses when either toggle is off, because they decide what the round can prove", () => {
    expect(readiness({ ...ready, verbose: false }).stop.join(" ")).toContain("Verbose trace");
    expect(readiness({ ...ready, pictures: false }).stop.join(" ")).toContain("empty");
  });

  it("does not refuse merely because something could not be read", () => {
    // Unknown is not the same as wrong. A missing toggle reading means the pane
    // was not open at that moment, and turning that into a hard stop would make
    // the driver unusable while telling nobody anything true.
    expect(readiness({ ...ready, verbose: null, pictures: null, slides: null }).ok).toBe(true);
  });
});

describe("the Automation tab holds everything the driver needs", () => {
  it("does not treat an unread toggle as permission to run", () => {
    // `Verbose trace`, `Picture every slide` and the run button all live on the
    // Automation tab, and a pane does not open there — a freshly sideloaded or
    // reopened one comes up on `Chart`, where none of the three is in the DOM.
    //
    // 2026-08-20: the driver put the add-in back by itself for the first time,
    // reached `ready` with `verbose trace ?` in the same block, and then stopped
    // because the run button was not there either. Both readings had one cause.
    //
    // `null` means "could not read it". A round started on that produces a trace
    // too thin to mine, and the point of selecting the tab first is that the
    // answer becomes a reading.
    const unread = readiness({ ...READY, verbose: null });
    expect(unread.ok, "an unread Verbose trace is not evidence that it is on").toBe(true);
    // The refusal that DOES exist stays: an explicit false still stops the round.
    const off = readiness({ ...READY, verbose: false });
    expect(off.ok).toBe(false);
    expect(off.stop.join(" ")).toMatch(/Verbose trace is off/);
  });
});

describe("a sign-in TAB is not a sign-in PROMPT", () => {
  it("refuses when the host is silent AND the document never rendered", () => {
    // A genuine signed-out browser: no deck, no slide list, nothing to measure.
    // Only a person can clear it, and the message says so.
    const r = readiness({ ...READY, loggedOut: true, authPopup: true, ping: null, slides: null });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toMatch(/needs a password/);
  });

  it("does NOT blame a sign-in when the deck's slide list still reads", () => {
    // The hard case, and the one that would have cost a night. Chrome's crash
    // restore leaves a `login.live.com` tab open PERMANENTLY — it never goes
    // away on its own — so this refusal would fire on every future round the
    // moment the host went briefly quiet, which on this machine is the commonest
    // transient state there is.
    //
    // A readable slide list does not PROVE the session is fine: Office can raise
    // a re-auth prompt beside a loaded deck, which is exactly what this refusal
    // was written for. The two are indistinguishable from the tab list — so the
    // honest move is to report the silent host that was actually observed, let
    // the reload behind it run, and mention the tab rather than blaming it.
    // A silent host reports `{ answered: false }`, not null — null means the ping
    // was never readable, which is a different state and does not refuse.
    const r = readiness({ ...READY, loggedOut: true, authPopup: true, ping: { answered: false, ms: 8009 }, slides: 1 });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" "), "claimed a sign-in it could not distinguish").not.toMatch(/needs a password/);
    expect(r.codes, "a silent host is recoverable; a sign-in is not").toContain("host-silent");
  });

  it("carries on when the host is ANSWERING, because that proves the session is signed in", () => {
    // 2026-08-20: Chrome's "Restore pages?" brought back an `oauth20_authorize`
    // popup from a crashed session, so the tab list showed a login page beside a
    // deck that was working perfectly. The check printed
    // `host answered in 7ms · slide 1 resolved` and refused in the same breath,
    // telling the owner to enter a password they had already entered.
    //
    // A host that answers Office.js and resolves a slide cannot be
    // unauthenticated. The ping is better evidence than the tab list because it
    // asks the thing we care about rather than looking at the furniture.
    const r = readiness({ ...READY, loggedOut: true, authPopup: true, ping: { answered: true, ms: 7 }, slideOk: true });
    expect(r.ok, "a stale login tab blocked a working deck").toBe(true);
  });

  it("still refuses when the slide will not resolve, even if the ping answered", () => {
    // Half-answering is not answering. `slideOk === false` is the host taking a
    // call and failing the one that matters, which is exactly the state the
    // original guard existed to stop a round on.
    const r = readiness({
      ...READY,
      loggedOut: true,
      authPopup: true,
      ping: { answered: true, ms: 7 },
      slideOk: false,
    });
    expect(r.ok).toBe(false);
  });
});

describe("setting the deck's slide size, only when asked", () => {
  it("writes BOTH dimensions and reads them back rather than assuming", () => {
    // `PowerPoint.PageSetup.slideWidth`/`slideHeight` are writable at
    // PowerPointApi 1.10, which round 096's `environment` line shows this host
    // advertising — so the `wrong-size` refusal was asking a person to do
    // something the API can do.
    //
    // The read-back is the point. The driver's own refusal text has warned for
    // months that a size change "made while the document is loading is accepted
    // and does nothing", and an API write deserves the same suspicion.
    const four = setSlideSizeScript("4:3", 20_000);
    expect(four, "4:3 is 720x540 at 96dpi").toMatch(/slideWidth = 720/);
    expect(four).toMatch(/slideHeight = 540/);
    expect(four, "wrote without reading the result back").toMatch(/load\("slideWidth,slideHeight"\)/);
    expect(four).toMatch(/return "size:"/);

    const wide = setSlideSizeScript("16:9", 20_000);
    expect(wide).toMatch(/slideWidth = 960/);
    expect(wide).toMatch(/slideHeight = 540/);
  });

  it("carries a budget, so a host that will not answer cannot hang the leg", () => {
    expect(setSlideSizeScript("4:3", 12_345)).toMatch(/12345/);
  });

  it("round-trips through readSlideSize, or the driver cannot tell whether it took", () => {
    // The setter returns the same `size:WxH` shape the reader parses. If those
    // two ever drift, the driver would resize the deck and then report that it
    // could not read the result — and carry on with the old value.
    expect(readSlideSize("size:720x540")).toBe("4:3");
    expect(readSlideSize("size:960x540")).toBe("16:9");
    expect(readSlideSize("size-failed:budget"), "a failure must not read as a size").toBeNull();
  });
});

describe("telling a control that is absent from one that is merely disabled", () => {
  // Playwright hands out a `ref` only for something it could act on, so a
  // greyed-out button matches its line and carries none. `refFor` answers null
  // for both, and the driver read that single null as "the add-in is gone".
  const findReturning = (out: string) => ((cmd: string) => (cmd === "find" ? out : "")) as never;

  it("sees a disabled control that refFor cannot", () => {
    const sh = findReturning('        - button "Insert chart" [disabled]');
    expect(refFor(sh, "Insert chart", /button "Insert chart"/), "a disabled control has no ref").toBeNull();
    expect(namePresent(sh, "Insert chart", /button "Insert chart"/), "but it IS in the ribbon").toBe(true);
  });

  it("still says absent when the ribbon really does not carry it", () => {
    const sh = findReturning('        - button "Design Suggestions" [ref=f12e42]');
    expect(namePresent(sh, "Insert chart", /button "Insert chart"/)).toBe(false);
  });

  it("agrees with refFor when the control is usable", () => {
    const sh = findReturning('        - button "Insert chart" [ref=f12e99]');
    expect(refFor(sh, "Insert chart", /button "Insert chart"/)).toBe("f12e99");
    expect(namePresent(sh, "Insert chart", /button "Insert chart"/)).toBe(true);
  });
});

describe("talking to the browser at all", () => {
  it("bounds every CLI call, because the round's own deadline cannot", () => {
    // THE STALL THIS DRIVER COULD NOT SURVIVE. The poll loop checks a 30-minute
    // limit at the TOP of each pass, so it only ever fires if the call below it
    // returned. An unbounded spawn meant a wedged CLI — a browser that stopped
    // answering, a tab mid-crash — hung the driver with that deadline sitting
    // there unreachable and nothing printed since.
    const seen: Record<string, unknown>[] = [];
    const run = (_exe: string, _args: string[], opts: Record<string, unknown>) => {
      seen.push(opts);
      return { status: 0, stdout: "" };
    };
    cli(run, ".", "some-cli.js")("list");
    expect(seen).toHaveLength(1);
    expect(seen[0].timeout, "an unbounded call can hang the whole night").toBeGreaterThan(0);
    // Generous, or it kills the slow calls that are legitimate: an `eval`
    // carries a 20s page-side budget and `requests` can return tens of MB.
    expect(seen[0].timeout).toBeGreaterThanOrEqual(60_000);
  });

  it("fronts the deck a round was told to run against, and refuses when it is not open", () => {
    // `PW_DECK` used to reach only `recover`, in the branch that reopens a dead
    // browser — so a cycle setting it per leg chose which deck a RECOVERY would
    // hunt for and nothing else. The ordinary path never selected a tab, so the
    // nightly cycle's 4:3 leg measured whichever document leg two left open and
    // refused with `wrong-size` every single night.
    const calls: string[][] = [];
    const shWith = (tabs: string) => {
      const sh = ((...args: string[]) => {
        calls.push(args);
        return args[0] === "tab-list" ? tabs : "";
      }) as never as { (...a: string[]): string; state: unknown };
      return sh;
    };
    const open = shWith('0: https://x/ "Presentation64"\n1: https://x/ "Presentation66"');
    expect(selectDeck(open, "Presentation66")).toBe(true);
    expect(
      calls.some((c) => c[0] === "tab-select" && c[1] === "1"),
      "did not front the deck it was given",
    ).toBe(true);
    // Not open at all: say so rather than measure the wrong document.
    expect(selectDeck(shWith('0: https://x/ "Presentation64"'), "Presentation66")).toBe(false);
    // No deck asked for is the behaviour this has always had.
    expect(selectDeck(shWith(""), null)).toBe(true);
  });

  it("hands the next round a fresh pane, not just a clean deck", async () => {
    // A CLEAN DECK IS NOT A CLEAN PANE. The pane's age when a round starts
    // separates post-retry 0.43 from 4.57 and a 16-shape deck from a 60+ one
    // (rounds 110-123). Sweeping clears the SLIDES and leaves the pane exactly
    // as the last round left it, so every second round in this archive is a
    // degraded sample — a fact published three times as a property of the
    // profile, the position and the observer before anyone measured pane age.
    const calls: string[][] = [];
    const sh = ((...args: string[]) => {
      calls.push(args);
      if (args[0] === "find" && /Insert chart/.test(args[1] ?? "")) return 'button "Insert chart" [ref=r1]';
      if (args[0] === "find" && /Automation/.test(args[1] ?? "")) return 'tab "Automation" [ref=r2]';
      return "";
    }) as never as { (...a: string[]): string; state: unknown };

    const ok = await refreshPane(sh, async () => {});
    expect(ok, "reported no pane when one reopened").toBe(true);
    expect(
      calls.some((c) => c[0] === "reload"),
      "never reloaded the tab",
    ).toBe(true);
    // AND IT MUST ANSWER THE BEFOREUNLOAD MODAL. PowerPoint asks "changes you
    // made may not be saved" when a tab with unsaved work is reloaded, and until
    // that is answered EVERY command fails — `find` returns nothing at all, not
    // even its miss message, `tab-list` shows an empty title, and `screenshot`
    // refuses with "does not handle the modal state". The browser looks dead and
    // is not. Calling this straight after `sweepDeck` GUARANTEES unsaved changes,
    // and round 124 wedged exactly there until a human accepted the dialog.
    expect(
      calls.some((c) => c[0] === "dialog-accept"),
      "left a modal blocking every later call",
    ).toBe(true);
    // Reopened AND put on Automation: a pane always reopens on Chart, where
    // `Verbose trace` and the run button are not in the DOM at all, so a round
    // starting there reads `verbose trace ?` and cannot find its own button.
    expect(calls.filter((c) => c[0] === "eval").length, "clicked neither the pane nor the tab").toBeGreaterThanOrEqual(
      2,
    );

    // AND IT MUST NOT RELOAD TWICE. `recover` chooses between a crash dialog's
    // Refresh button and a plain reload before delegating here; a second reload
    // would throw away the dialog handling it just did.
    const after: string[][] = [];
    const sh2 = ((...args: string[]) => {
      after.push(args);
      return "";
    }) as never as { (...a: string[]): string; state: unknown };
    await refreshPane(sh2, async () => {}, { reloaded: true });
    expect(
      after.some((c) => c[0] === "reload"),
      "reloaded on top of the caller's own reload",
    ).toBe(false);
  });

  it("does NOT reload the pane at the end of a round", () => {
    // THIS GUARD USED TO ASSERT THE OPPOSITE, and inverting it is the point.
    //
    // Reloading between rounds worked: round 126 was the first second round in
    // this archive to score a first round's numbers. It also broke the session
    // roughly one round in four. A reload of a tab with unsaved work raises
    // PowerPoint's beforeunload prompt, and `dialog-accept` means LEAVE WITHOUT
    // SAVING — while the sweep immediately before it GUARANTEES unsaved work,
    // having just deleted slides. The add-in was gone from the ribbon after the
    // reload following round 124 and again after 132, with the deck grown back
    // to 79 slides.
    //
    // So the call is out until the driver can wait for the autosave and reload a
    // clean document, which needs a way to read PowerPoint's saved state that
    // does not exist here yet. A guard that merely stopped asserting would let
    // it drift back in unnoticed; this one states the cost.
    const src = readFileSync(new URL("../scripts/round.mjs", import.meta.url), "utf8");
    // ANCHOR FIRST, ASSERT SECOND — this guard silently degraded to checking a
    // SINGLE CHARACTER. `indexOf` returns -1 when the declaration is spelled any
    // other way (`const collectRound = async`, an added `export`), and
    // `src.slice(-1)` is then the last character of the file; `indexOf` for the
    // closing brace on that is -1, so `fn` fell back to `body.slice(0, 4000)` —
    // one character — and `not.toMatch` passes on it.
    //
    // A source-text guard that cannot find its anchor reports a floor, not a
    // count, and this one fences off the regression that COSTS THE SIDELOAD:
    // rounds 124 and 132, about one round in four. Silence reading as success,
    // inside the guard written to stop exactly that.
    const at = src.indexOf("async function collectRound");
    expect(at, "the guard could not find collectRound, so it was about to check one character").toBeGreaterThan(-1);
    const body = src.slice(at);
    const end = body.indexOf("\n}\n");
    expect(
      end,
      "collectRound has no closing brace at column 0 — the slice would be the rest of the file",
    ).toBeGreaterThan(0);
    const fn = body.slice(0, end);
    expect(fn.length, "the extracted function body is too short to contain anything").toBeGreaterThan(500);
    expect(fn, "collectRound reloads the pane again — that costs the sideload about 1 round in 4").not.toMatch(
      /^\s*(?:if \()?await refreshPane\(/m,
    );
    // And `refreshPane` itself must survive: `recover` depends on it, and it is
    // correct there because the page it reloads has usually already crashed and
    // has nothing left to save.
    expect(src, "refreshPane was deleted, taking recovery's reload with it").toMatch(
      /export async function refreshPane/,
    );
    // Same anchor discipline. A missing declaration slices the last character
    // and this `toMatch` would fail — loudly, but naming the wrong thing.
    const recoverAt = src.indexOf("export async function recover");
    expect(recoverAt, "recover is no longer declared this way — the guard is reading the wrong text").toBeGreaterThan(
      -1,
    );
    expect(src.slice(recoverAt), "recover stopped using it").toMatch(/refreshPane\(/);
  });

  it("will not call an add-in missing when the window is too narrow to show it", () => {
    // THE MOST EXPENSIVE MISREADING IN THIS PROJECT, and it cost a day.
    // PowerPoint collapses trailing ribbon commands into a `...` overflow on a
    // narrow window, and a collapsed command IS NOT IN THE ACCESSIBILITY TREE.
    // So the ribbon read answers false — truthfully, it is not rendered — and
    // every reader above it concludes the add-in is gone.
    //
    // Measured 2026-08-21, same browser, same document, seconds apart:
    //
    //     page width 1237   Insert chart: false   "the add-in is not loaded here"
    //     page width 2375   Insert chart: true    round runs
    //
    // On that reading the driver ran the sideload walk FOUR TIMES against a deck
    // that already had the add-in and sent the owner to do it by hand.
    const cramped = readiness({
      ...READY,
      stamp: null,
      canOpenPane: false,
      commandPresent: false,
      slides: 1,
      ribbonRoom: 1237,
    });
    expect(cramped.ok).toBe(false);
    expect(cramped.codes, "called it a missing add-in from a window that cannot show one").not.toContain(
      "addin-missing",
    );
    expect(cramped.codes).toContain("ribbon-cramped");
    // RECOVERABLE, unlike `addin-missing`: widening a window is not a change to
    // anything the round measures, so the driver fixes it instead of stopping.
    expect(RECOVERABLE_STOPS.has("ribbon-cramped"), "a cramped ribbon must not end the night").toBe(true);

    // 1996 IS THE WIDTH THAT GOT THROUGH, and it is pinned by number because a
    // threshold is only worth what the widest CRAMPED observation says. On
    // 2026-08-22 a 1996px window hid `Insert chart` exactly as 1237 had, but
    // 1996 cleared the old 1600 limit — so this branch was skipped, the fatal
    // `addin-missing` was taken instead, and the driver spent an evening
    // sideloading into a deck that had the add-in the whole time.
    expect(
      readiness({ ...READY, stamp: null, canOpenPane: false, commandPresent: false, slides: 1, ribbonRoom: 1996 })
        .codes,
      "1996px was called wide enough to judge, and it hides the command",
    ).toContain("ribbon-cramped");

    // AN UNREADABLE WIDTH IS NOT A WIDE ONE. A reader that cannot measure the
    // window cannot tell a missing add-in from a hidden one.
    expect(
      readiness({ ...READY, stamp: null, canOpenPane: false, commandPresent: false, slides: 1, ribbonRoom: null })
        .codes,
    ).toContain("ribbon-cramped");

    // AND THE REAL REFUSAL MUST SURVIVE. With room to render and still no
    // command, the add-in genuinely is not there — that stop is the point of
    // all this and must not be softened away.
    const roomy = readiness({
      ...READY,
      stamp: null,
      canOpenPane: false,
      commandPresent: false,
      slides: 1,
      ribbonRoom: 2375,
    });
    expect(roomy.codes, "a genuinely missing add-in stopped being reported").toContain("addin-missing");
    expect(RECOVERABLE_STOPS.has("addin-missing"), "addin-missing must stay a hard stop").toBe(false);
  });

  it("recovers a deck that is merely not open, but never a wrong slide size", async () => {
    // `recover` navigates to OneDrive and clicks the file whose link matches the
    // name — that branch has existed for months and `deck-missing` was never
    // wired to it, so the stop ended the night with the repair one call away. A
    // full cycle died this way on 2026-08-27 when leg 1's tab closed during a
    // `--fresh` restart.
    expect(RECOVERABLE_STOPS.has("deck-missing"), "recovery can open the named deck and should").toBe(true);

    // THE LINE THAT MUST NOT MOVE WITH IT. `wrong-size` is excluded because
    // setting the size would change WHAT THE ROUND MEASURES, which is the one
    // thing recovery is forbidden to do — `cycle.mjs` says so in its header.
    // Opening the deck the caller named is setup; resizing it is falsification.
    expect(RECOVERABLE_STOPS.has("wrong-size"), "recovery must never set the slide size").toBe(false);
  });

  it("does not judge a size it read off the wrong deck", () => {
    /**
     * THE TWO STOPS ABOVE MET, AND THE FATAL ONE WON.
     *
     * When the wanted deck is not fronted, `size` is the size of whatever other
     * document IS — so the size check refused `wrong-size` about a deck the
     * round was never going to run against. `deck-missing` is recoverable and
     * `wrong-size` deliberately is not, so the bogus companion turned a stop the
     * driver clears itself into one that ends the night.
     *
     * It bites the NORMAL sequence, which is why a single cycle never saw it:
     * leg 3 runs at 4:3 and leaves that deck in front, so the next cycle's leg 1
     * wants the 16:9 deck and measures the 4:3 one still sitting there. A full
     * cycle died this way on 2026-08-27 and again on 2026-08-29 — the second
     * time after the reorder above had already improved the MESSAGE, which is
     * what made it look fixed.
     */
    const away = readiness({
      ...READY,
      wantDeck: "Presentation64",
      deckFronted: false,
      size: "4:3",
      expectSize: "16:9",
    });
    expect(away.codes, "named the deck, which is the useful half").toContain("deck-missing");
    expect(away.codes, "judged the size of a deck it was not going to run against").not.toContain("wrong-size");
    // The point of the whole fix: every code left is one recovery can clear, so
    // the driver retries instead of ending the night.
    expect(
      away.codes.every((c: string) => RECOVERABLE_STOPS.has(c)),
      `a stop the driver cannot clear survived: ${away.codes.join(", ")}`,
    ).toBe(true);

    // AND THE REAL REFUSAL MUST SURVIVE. With the right deck in front at the
    // wrong size, the owner really has set the profile wrong, and filing a round
    // under it is worse than not running one.
    const here = readiness({
      ...READY,
      wantDeck: "Presentation64",
      deckFronted: true,
      size: "4:3",
      expectSize: "16:9",
    });
    expect(here.codes, "the refusal this check exists for was softened away").toContain("wrong-size");
    // Unchanged where no deck was named: whichever document is in front IS the
    // one being judged, so its size is the right size to judge.
    const anyDeck = readiness({ ...READY, size: "4:3", expectSize: "16:9" });
    expect(anyDeck.codes).toContain("wrong-size");
  });

  it("never reads `find`'s miss message as a hit", async () => {
    // THE WORST POLARITY A BUG CAN HAVE, and it caught me twice on 2026-08-20.
    // `playwright-cli find` answers a miss with:
    //
    //     No matches found for "Insert chart".
    //
    // which CONTAINS the query. So a probe built from the bare name matches the
    // miss, and reports a control present at the exact moment it is absent. It
    // said a crash dialog was up when none was, and then said the add-in was
    // back in the ribbon when it was not — which I passed to the owner as
    // confirmation of something they had said, so the bad reading did not even
    // look like a machine's mistake.
    const missing = ((...args: string[]) =>
      args[0] === "find" ? `No matches found for "${args[1]}".` : "") as never as {
      (...a: string[]): string;
      state: unknown;
    };
    expect(namePresent(missing, "Insert chart", /"Insert chart"/), "matched the miss message itself").toBe(false);
    expect(namePresent(missing, "Insert chart", /button "Insert chart"/)).toBe(false);
    expect(refFor(missing, "Insert chart", /"Insert chart"/), "took a ref out of a miss").toBeNull();

    // And a real hit still reads as one — a filter that swallowed everything
    // would pass every assertion above.
    const found = ((...args: string[]) => (args[0] === "find" ? 'button "Insert chart" [ref=r7]' : "")) as never as {
      (...a: string[]): string;
      state: unknown;
    };
    expect(namePresent(found, "Insert chart", /button "Insert chart"/)).toBe(true);
    expect(refFor(found, "Insert chart", /button "Insert chart"/)).toBe("r7");
  });

  it("waits for a sideloaded command instead of looking once", async () => {
    // A NOT-YET REPORTED AS A NEVER, and the most expensive version of this
    // repo's house defect so far. `sideloadAddIn` uploaded the manifest, slept
    // 12s, looked for `Insert chart` ONCE, and returned "uploaded, but no
    // SSF Charts command appeared". Round 117 refused on it and told the owner
    // only a person could put the add-in back. THE OWNER TOUCHED NOTHING AND
    // THE COMMAND APPEARED ANYWAY — the sideload had worked; the check was
    // early. A verdict that sends a person to redo work the machine already did
    // is worse than no verdict.
    let looks = 0;
    const sh = ((...args: string[]) => {
      if (args[0] !== "find") return "";
      looks++;
      // Absent for the first three polls, then the ribbon repopulates.
      return looks > 3 ? 'button "Insert chart" [ref=abc]' : "";
    }) as never as { (...a: string[]): string; state: unknown };

    let clock = 0;
    const sleep = async (ms: number) => {
      clock += ms;
    };
    const found = await waitForRef(sh, sleep, "Insert chart", /button "Insert chart"/, 90_000, 3000, () => clock);
    expect(found, "gave up on a command that did arrive").toBe("abc");
    expect(looks, "looked once — that is a sleep, not a wait").toBeGreaterThan(1);

    // AND IT MUST STILL GIVE UP. A poll with no end is how an overnight run is
    // found in the morning having done nothing, so the budget has to bite.
    let clock2 = 0;
    const never = ((...args: string[]) => (args[0] === "find" ? "" : "")) as never as {
      (...a: string[]): string;
      state: unknown;
    };
    const gone = await waitForRef(
      never,
      async (ms: number) => {
        clock2 += ms;
      },
      "Insert chart",
      /button "Insert chart"/,
      9_000,
      3000,
      () => clock2,
    );
    expect(gone, "polled forever for something that never arrives").toBeNull();
    expect(clock2, "the budget did not bound the wait").toBeLessThanOrEqual(12_000);

    // AND IT MUST STILL END WHEN THE CLOCK DOES NOT MOVE. A wait bounded only
    // by time assumes the sleep really sleeps; every stub in this file returns
    // immediately, so `now()` stands still and the loop spins as fast as the
    // process can allocate. That is not a polite hang — it took Vitest to
    // `FATAL ERROR: JavaScript heap out of memory` in 80 seconds. The attempt
    // counter is the bound that a frozen clock cannot defeat.
    let calls = 0;
    const counting = ((...args: string[]) => {
      if (args[0] === "find") calls++;
      return "";
    }) as never as { (...a: string[]): string; state: unknown };
    const stuck = await waitForRef(
      counting,
      async () => {},
      "Insert chart",
      /button "Insert chart"/,
      90_000,
      3000,
      () => 0, // a clock that never moves
    );
    expect(stuck).toBeNull();
    expect(calls, "spun without bound against a frozen clock").toBeLessThanOrEqual(40);
  });

  it("passes the round config to `open` and to nothing else", () => {
    // THE VIEWPORT FIX, AND THE REGRESSION IT SHIPPED WITH FOR ONE MINUTE.
    //
    // The round browser was opened with no config, leaving it on a fixed
    // 2880x1800 page viewport inside a 1280x752 window — the page laid out more
    // than twice as wide as the window could show, so the strip holding the task
    // pane fell outside the visible area and the owner could not watch a round.
    // `viewport: null` makes the page follow the window (measured: 1036 inner
    // against 1050 outer, versus 2880 against 1280).
    //
    // The first version passed the flag on EVERY CLI call, which reads as
    // harmless and is not: `--config` is an option of `open` alone, so every
    // other command answered `Unknown option: --config` and the driver reported
    // a healthy browser as `pane ?`, `deck ?`, "the pane is closed" — a config
    // flag whose failure mode impersonates the crash it exists to help someone
    // watch.
    const args: string[][] = [];
    const run = (_exe: string, a: string[]) => {
      args.push(a);
      return { status: 0, stdout: "" };
    };
    const sh = cli(run, ".", "some-cli.js");
    sh("tab-list");
    sh("find", "Chart");
    expect(
      args.every((a) => !a.some((x) => x.startsWith("--config"))),
      "an ordinary command carried --config, which playwright-cli rejects outright",
    ).toBe(true);

    // And the flag itself: present when the file is there, ABSENT when it is
    // not. A flag pointing at a path that does not exist would fail the one
    // call that matters most — the reopen after a browser death.
    expect(roundConfigArg(() => true)).toHaveLength(1);
    expect(roundConfigArg(() => true)[0].startsWith("--config=")).toBe(true);
    expect(
      roundConfigArg(() => false),
      "pointed at a config that is not there",
    ).toEqual([]);
  });

  it("names the document the check is actually about", () => {
    // THE CHECK NEVER SAID WHICH DECK IT CHECKED. Build stamps, slide counts,
    // host latency — every line described the round without naming the
    // document, so a check against the wrong deck read exactly like a check
    // against the right one. On 2026-08-20 the dedicated 4:3 deck was still
    // 960x540 and nothing in any output contradicted the note that called the
    // 4:3 leg "blocked on owner setup".
    const shWith = (tabs: string) =>
      ((...args: string[]) => (args[0] === "tab-list" ? tabs : "")) as never as {
        (...a: string[]): string;
        state: unknown;
      };
    const real =
      "- 0: [Home - OneDrive](https://onedrive.live.com/)\n" +
      "- 1: [Presentation64.pptx](https://onedrive.live.com/x)\n" +
      "- 3: (current) [Presentation67.pptx](https://onedrive.live.com/y)";
    expect(frontedDeck(shWith(real)), "took a name from the wrong line").toBe("Presentation67.pptx");

    // UNREADABLE IS NOT A NAME. Both of these used to be indistinguishable from
    // a successful read of a deck called nothing, which is this project's house
    // defect; the caller prints `?` for null and must be given the chance.
    expect(frontedDeck(shWith("")), "an empty tab list is not a deck").toBeNull();
    expect(
      frontedDeck(shWith("- 0: (current) [Home - OneDrive](https://onedrive.live.com/)")),
      "not a document",
    ).toBeNull();
  });

  /**
   * Putting the add-in back after a browser death took the sideload with it.
   *
   * Ten steps of Office ribbon automation, which will break the day Microsoft
   * moves a control. What these hold is not the walk but the two things that
   * decide whether a breakage is survivable: it must never leave a dialog over
   * the document, and it must never claim success it did not verify.
   */
  describe("putting the add-in back", () => {
    /** A fake `sh` whose answers are keyed by the query it is given. */
    const shWith = (answers: Record<string, string>, missing: string[] = []) => {
      const calls: string[][] = [];
      const sh = ((...args: string[]) => {
        calls.push(args);
        const q = args[1] ?? "";
        // EXACT, not substring. `"Manage My Add-ins".includes("Add-ins")` is
        // true, so a loose match hands three different queries the same answer
        // and the walk "passes" having clicked the wrong control twice.
        if (missing.includes(q)) return "";
        return answers[q] ?? "";
      }) as never as { (...a: string[]): string; calls: string[][] };
      sh.calls = calls;
      return sh;
    };
    // Every control the walk needs, each answering with a ref line `refFor` parses.
    const ALL: Record<string, string> = {
      "Slide List": 'listbox "Slide List" [ref=r1]',
      "Add-ins": 'button "Add-ins" [ref=r2]',
      "See all": 'menuitem "See all installed add-ins" [ref=r3]',
      "More Add-ins": 'menuitem "More Add-ins" [ref=r4]',
      "MY ADD-INS": 'tab "MY ADD-INS" [ref=r5]',
      "Manage My Add-ins": 'button "Manage My Add-ins" [ref=r6]',
      "Upload My Add-in": 'menuitem "Upload My Add-in" [ref=r7]',
      Browse: 'button "Browse..." [ref=r8]',
      Upload: 'button "Upload" [ref=r9]',
      "Insert chart": 'button "Insert chart" [ref=r10]',
    };
    const now = () => Promise.resolve();

    it("walks the whole flow and uploads the PROD manifest", async () => {
      const sh = shWith(ALL);
      expect(await sideloadAddIn(sh, now, "C:/x/manifest-prod.xml")).toBe(true);
      const uploaded = sh.calls.find((c) => c[0] === "upload");
      expect(uploaded?.[1], "sideloaded the dev manifest").toMatch(/manifest-prod\.xml$/);
      // The document must be focused FIRST — the ribbon ignores clicks until it
      // is, which cost three failed attempts to discover and reports itself as
      // `aria-expanded=true` over a menu that is plainly shut.
      const firstClick = sh.calls.findIndex((c) => c[0] === "eval");
      const listLookup = sh.calls.findIndex((c) => (c[1] ?? "").includes("Slide List"));
      expect(listLookup).toBeLessThan(firstClick);
    });

    it("never leaves a dialog over the document when a step is missing", async () => {
      // THE RISK THAT MATTERS MOST. A half-walked dialog is invisible to
      // `readiness` and fatal to every round after it — the deck reads fine and
      // nothing can be clicked.
      for (const step of ["More Add-ins", "Upload My Add-in", "Insert chart"]) {
        // Cancel has to be FINDABLE, or this asserts nothing: looking for a
        // dismiss control and clicking one are different acts, and the first
        // version of this test checked only that the lookup happened — it
        // passed with the clicks deleted.
        const sh = shWith({ ...ALL, Cancel: 'button "Cancel" [ref=rcancel]' }, [step]);
        expect(await sideloadAddIn(sh, now, "C:/x/m.xml"), `${step} should have failed`).toBe(false);
        const clickedCancel = sh.calls.some((c) => c[0] === "eval" && c[2] === "rcancel");
        expect(clickedCancel, `left a dialog open after failing at ${step}`).toBe(true);
      }
    });

    /**
     * TWO UIS ARE IN THE FIELD AND ONLY ONE HAS THE `See all` RUNG.
     *
     * Measured 2026-09-05 on a brand-new presentation: the Add-ins flyout came
     * up straight on the store panel, with no "See all installed add-ins"
     * anywhere in it and `More Add-ins` sitting right there as the next rung.
     * The walk refused with "the Add-ins menu did not open", which was false
     * twice over — the menu HAD opened, and the control it wanted is not on
     * that variant at all. It cost a sideload that then had to be walked by
     * hand.
     */
    it("carries on when this UI has no `See all` rung, and still stops when the menu is shut", async () => {
      const newUi = shWith(ALL, ["See all"]);
      expect(await sideloadAddIn(newUi, now, "C:/x/m.xml"), "refused a menu that had opened").toBe(true);

      // The rung AFTER it is what decides, because `More Add-ins` is on both
      // variants. A genuinely unopened menu still fails — one line later, with
      // a message that is true — and still leaves no dialog behind.
      const shut = shWith({ ...ALL, Cancel: 'button "Cancel" [ref=rcancel]' }, ["See all", "More Add-ins"]);
      expect(await sideloadAddIn(shut, now, "C:/x/m.xml"), "walked on with no menu open").toBe(false);
      expect(
        shut.calls.some((c) => c[0] === "eval" && c[2] === "rcancel"),
        "left a dialog open",
      ).toBe(true);
    });

    it("will not call it done when the host never accepted the manifest", async () => {
      // The Upload button is disabled until a file is accepted, so its enabled
      // state is the host's own receipt — better evidence than the upload call
      // not throwing.
      const sh = shWith({ ...ALL, Upload: 'button "Upload" [disabled] [ref=r9]' });
      expect(await sideloadAddIn(sh, now, "C:/x/m.xml")).toBe(false);
    });

    it("reloads before walking, because it cannot open a menu that is already open", async () => {
      // TWICE ON 2026-08-21 the first attempt died at "the Add-ins menu did not
      // open", and both times the cause was a `button "Add-ins" [expanded]`
      // left over from a PREVIOUS failed attempt. Clicking an open menu closes
      // it, so every retry toggled the menu instead of walking it — and the
      // driver reported the ribbon as unopenable when it was merely already
      // open.
      //
      // Escape is not enough, measured: it returned the button to `[active]`
      // and the next attempt still failed. Only a reload let the walk reach the
      // upload, both times.
      const calls: string[][] = [];
      const sh = ((...args: string[]) => {
        calls.push(args);
        return "";
      }) as never as { (...a: string[]): string; state: unknown };
      await sideloadAddIn(sh, async () => {}, "C:/x/m.xml");

      const reloadAt = calls.findIndex((c) => c[0] === "reload");
      const firstFind = calls.findIndex((c) => c[0] === "find");
      expect(reloadAt, "never reloaded — a stuck menu will defeat every attempt").toBeGreaterThanOrEqual(0);
      expect(reloadAt, "reloaded AFTER starting the walk, which is too late to clear the menu").toBeLessThan(firstFind);
      // And the modal a reload can raise must be answered, or every later call
      // fails in a way that impersonates a dead browser.
      expect(
        calls.some((c) => c[0] === "dialog-accept"),
        "left a beforeunload modal blocking the walk it just started",
      ).toBe(true);

      // AND IT MUST WAIT FOR THE RIBBON, not sleep at it. This was a fixed 55s
      // sleep and one look, and it failed on the first cold browser start it
      // met: round 133, after a machine restart, reported `no Add-ins button in
      // the ribbon` — OFFICE'S OWN button — because the document was still
      // loading. A minute later the whole ribbon was there.
      //
      // The tell is more than one `find` for the ribbon before the walk gives
      // up: a single look cannot tell "not there" from "not there YET".
      const docLooks = calls.filter((c) => c[0] === "find" && /Slide List/.test(c[1] ?? "")).length;
      expect(docLooks, "looked for the document once — that is a sleep, not a wait").toBeGreaterThan(1);
    });

    it("refuses to walk a document that is not up", async () => {
      // Ten ribbon steps against a loading tab would leave a dialog on it.
      const sh = shWith(ALL, ["Slide List"]);
      expect(await sideloadAddIn(sh, now, "C:/x/m.xml")).toBe(false);
      expect(
        sh.calls.some((c) => (c[1] ?? "").includes("Add-ins")),
        "started clicking anyway",
      ).toBe(false);
    });
  });

  it("believes the sweep's own answer rather than that it was called", () => {
    // "deck swept — the next round starts clean" printed whether the sweep had
    // cleaned the deck, failed outright, or left slides on it. The next round
    // then refuses with `deck-dirty` for a state the previous round's output
    // said could not exist. A tool that exits in silence reads as a pass.
    const shWith = (evalAnswer: string) =>
      ((...args: string[]) => {
        if (args[0] === "find") return 'tab "Chart" [ref=r1]';
        return evalAnswer;
      }) as never as (...a: string[]) => string;
    expect(sweepDeck(shWith("deck:1")), "a deck down to its one slide is clean").toBe(true);
    expect(sweepDeck(shWith("deck:7")), "seven slides left is not a swept deck").toBe(false);
    expect(sweepDeck(shWith("deck-failed")), "the sweep said it failed").toBe(false);
    expect(sweepDeck(shWith("")), "no answer is not a clean deck").toBe(false);
  });

  it("stops on the first attempt when the document has no add-in to open", () => {
    // SEVEN ATTEMPTS AND FIFTEEN MINUTES, on 2026-08-16. A 4:3 leg switched to a
    // deck SSF Charts was not registered for; `recover` reopens the pane from
    // the ribbon's `Insert chart` control, that deck offers no such control, and
    // the loop rediscovered this six more times. A refusal recovery cannot
    // address must not be retried.
    // `commandPresent: false` is what "no add-in here" actually means — the
    // ribbon does not carry the control at all. Saying only `canOpenPane: false`
    // is a weaker statement (it may be present and greyed out) and this fixture
    // used to make it, which is how the two states got conflated in the driver.
    const r = readiness({ ...READY, stamp: null, canOpenPane: false, commandPresent: false });
    expect(r.ok).toBe(false);
    expect(r.codes).toContain("addin-missing");
    expect(r.codes, "still reported as a merely-closed pane").not.toContain("pane-closed");
    expect(shouldRetry("not-ready", 0, 6, r.codes), "retrying this cannot help").toBe(false);
    // A DISABLED CONTROL IS NOT AN ABSENT ONE, and this is the case that cost
    // a round on 2026-08-19. PowerPoint's document went `Disconnected` — the
    // network, the same fault that crashed round 089 — which greys the WHOLE
    // ribbon. The accessibility tree still said `button "Insert chart"
    // [disabled]`, the add-in was loaded and fine, and the driver called it
    // `addin-missing`: the one refusal recovery may not retry, needing the
    // owner, for a state a reload clears.
    const greyed = readiness({ ...READY, stamp: null, canOpenPane: false, commandPresent: true });
    expect(greyed.codes, "a greyed ribbon is not a missing add-in").not.toContain("addin-missing");
    expect(greyed.codes).toContain("host-disconnected");
    expect(shouldRetry("not-ready", 0, 6, greyed.codes), "a reload clears this — it must be retried").toBe(true);

    // And a pane that is simply shut on a deck that CAN open one stays
    // recoverable, or the loop stops doing the thing it is good at.
    const shut = readiness({ ...READY, stamp: null, canOpenPane: true });
    expect(shut.codes).toContain("pane-closed");
    expect(shouldRetry("not-ready", 0, 6, shut.codes)).toBe(true);
  });

  it("will not call a tab that is merely mid-reload an absent add-in", () => {
    // THE DANGEROUS HALF OF THE CHECK ABOVE. A reloading tab answers nothing to
    // every read, so the ribbon looks as bare as a document with no add-in — and
    // `addin-missing` is deliberately un-retryable, so getting this wrong ends a
    // night on a state a reload would have cleared in twenty seconds.
    //
    // A readable slide list is the proof the document is actually up. The deck
    // that motivated the check had one; a loading tab does not.
    const loading = readiness({ ...READY, stamp: null, canOpenPane: false, slides: null });
    expect(loading.codes, "condemned a tab that was only still loading").not.toContain("addin-missing");
    expect(loading.codes).toContain("pane-closed");
    expect(shouldRetry("not-ready", 0, 6, loading.codes), "a reload would have fixed this").toBe(true);
  });

  it("names the missing DECK rather than blaming its slide size", () => {
    // The message matters as much as the refusal. "the deck is 16:9 and this
    // round was asked for 4:3" sends the owner to Design ▸ Slide Size for a
    // deck that was simply never opened.
    const r = readiness({ ...READY, wantDeck: "Presentation66", deckFronted: false });
    expect(r.ok).toBe(false);
    expect(r.codes).toContain("deck-missing");
    expect(r.stop.join(" ")).toContain("Presentation66");
    // And it must not fire when no deck was asked for.
    expect(readiness({ ...READY }).ok).toBe(true);
  });

  it("treats a timed-out call as nothing measured, not as an answer", () => {
    // A timeout arrives as `error`, which is the same shape as any other spawn
    // failure — and the difference that matters is already drawn: nothing was
    // measured, so the sweep must not trust anything it read.
    const run = () => ({ error: Object.assign(new Error("timed out"), { code: "ETIMEDOUT" }), status: null });
    const sh = cli(run, ".", "some-cli.js");
    expect(sh("list")).toBe("");
    expect(sh.state.unreachable, "a wedged CLI must not read as a healthy empty answer").toBe(true);
  });

  it("refuses everything, loudly, when the CLI could not be run", () => {
    // The failure this driver shipped with: every browser read came back empty
    // and it reported the pane as closed while the pane sat open on screen.
    // Unreachable has to short-circuit, because nothing below it was measured.
    const r = readiness({
      head: "a",
      deployed: "a",
      stamp: null,
      slides: null,
      verbose: null,
      pictures: null,
      reachable: false,
    });
    expect(r.ok).toBe(false);
    expect(r.stop).toHaveLength(1);
    expect(r.stop[0]).toContain("nothing below was actually measured");
  });

  it("tells a dead browser from a closed pane", () => {
    // SEVEN ATTEMPTS, 2026-08-15. A round wedged and the browser process died
    // with it; `recover` then reloaded and reopened a pane inside a window that
    // did not exist, seven times, while the check said "could not read the pane's
    // build stamp — is the add-in open?" The add-in was fine. There was nothing
    // to open it in.
    expect(noBrowser("  (no browsers)")).toBe(true);
    expect(noBrowser("### Browsers\n- ms:\n  - status: open")).toBe(false);
    expect(noBrowser("")).toBe(false);

    const r = readiness({ ...READY, browserGone: true, stamp: null, slides: null });
    expect(r.ok).toBe(false);
    expect(r.stop).toHaveLength(1);
    expect(r.stop[0], "sent the reader to the add-in again").not.toContain("is the add-in open?");
    expect(r.stop[0]).toContain("no browser");
    // Recoverable WITHOUT a password, because the profile keeps the sign-in —
    // so the loop reopens it rather than waiting for a person.
    expect(r.codes).toEqual(["browser-gone"]);
    expect(shouldRetry("not-ready", 0, 3, r.codes), "a dead browser is the loop's to fix").toBe(true);
  });

  it("tells an orphaned browser holding the profile from no browser at all", () => {
    // SEVEN ATTEMPTS AGAIN, 2026-09-06, and the same message was wrong in a new
    // way. A playwright daemon died leaving its Chrome running on
    // `C:/devtools/pw-profile`. `list` answered `(no browsers)` — truthfully,
    // the daemon knew of none — so `noBrowser` was true, the `close-all` guard
    // was skipped, and `open` refused a profile the driver had just decided was
    // free. Recovery then reported "the file list shows no `Presentation64` to
    // open", which named the one thing that was fine.
    //
    // The two states LOOK identical from `list` and differ only in what `open`
    // says, which is why the stderr had to stop being thrown away.
    expect(profileHeldByOrphan("Error: Browser is already in use for C:/devtools/pw-profile, use --isolated")).toBe(
      true,
    );
    expect(noBrowser("  (no browsers)"), "and `list` cannot tell them apart").toBe(true);
    // An unrelated failure is not this one.
    expect(profileHeldByOrphan("Error: Timeout 30000ms exceeded")).toBe(false);
    // A call that never ran left no stderr, and an absence must not read as a
    // diagnosis — `sh` returns null for a call that SUCCEEDED, too.
    expect(profileHeldByOrphan(null)).toBe(false);
    expect(profileHeldByOrphan(undefined)).toBe(false);
  });

  it("ends the orphan by its profile, never by being called chrome", () => {
    // THE ONE THING THAT MUST NOT GO WRONG. The driver's browser and the
    // operator's own are both `chrome.exe`; only `--user-data-dir` separates
    // them, so a matcher that named the process would close the window someone
    // is reading this in.
    const [cmd, args] = endOrphanCommand("C:/devtools/pw-profile", "win32");
    expect(cmd).toBe("powershell");
    expect(args.join(" "), "the profile is the whole filter").toContain("C:/devtools/pw-profile");
    expect(args.join(" ")).toContain("Stop-Process");

    // The platform is a PARAMETER for `is-main.mjs`'s reason: rounds run on
    // Windows and CI on ubuntu, so a test that could only exercise the platform
    // under it would go green against the case it exists to cover.
    const [posixCmd, posixArgs] = endOrphanCommand("/tmp/pw-profile", "linux");
    expect(posixCmd).toBe("pkill");
    expect(posixArgs.join(" "), "and the profile is still the filter").toContain("user-data-dir=/tmp/pw-profile");
  });

  it("reports whether the end-the-orphan command ran, not whether it worked", () => {
    // Whether the profile came free is answered by trying `open` again. An
    // inference here would be a third opinion, and a wrong one is how a
    // recovery loop declares success into a wedged host.
    const calls: string[][] = [];
    const ran = endOrphanedBrowser(
      "C:/p",
      (cmd: string, args: string[]) => {
        calls.push([cmd, ...args]);
        return {};
      },
      "win32",
    );
    expect(ran).toBe(true);
    expect(calls, "it has to actually spawn something").toHaveLength(1);

    // A spawn that never happened is not an ended browser.
    expect(endOrphanedBrowser("C:/p", () => ({ error: new Error("ENOENT") }), "win32")).toBe(false);
    expect(
      endOrphanedBrowser(
        "C:/p",
        () => {
          throw new Error("blocked by group policy");
        },
        "win32",
      ),
      "a throwing spawn must not take the round down with it",
    ).toBe(false);
  });

  it("says what the pane was showing when the run button is missing", () => {
    // `no-run-button` ended three cycles on 2026-09-06/07, each AFTER a
    // successful re-sideload — pane back, Automation tab selected, deck swept,
    // `ready`, then the refusal. The stop asks for a person and handed them
    // nothing, so the only question left ("what was on the pane?") had no
    // answer in any of the three logs.
    const paneOnTheWrongTab = [
      "### Elements",
      '- tab "Chart" [selected] [ref=f8e13]',
      '- tab "Elements" [ref=f8e14]',
      '- tab "Automation" [ref=f8e16]',
      '- button "Insert into slide" [ref=f8e99]',
    ].join("\n");
    const said = describePane(paneOnTheWrongTab);
    expect(said, "the tab list is what separates a wrong tab from a broken pane").toContain('tab "Automation"');
    expect(said).toContain('button "Insert into slide"');

    // A pane that answered nothing must not read as a pane with no controls —
    // they are different failures and this runs where both are possible.
    expect(describePane(""), "an empty answer").toContain("answered nothing at all");
    expect(describePane(undefined)).toContain("answered nothing at all");
    // Prose and refs in the same output are not controls.
    expect(describePane("### Chart\n- generic: some text\n  - img [ref=f1]")).toContain("answered nothing at all");

    // CAPPED, because this prints underneath the stop it explains. A full
    // accessibility dump would bury the message it is attached to.
    const many = Array.from({ length: 40 }, (_, i) => `- button "B${i}" [ref=f${i}]`).join("\n");
    expect(describePane(many).split("\n"), "an uncapped dump buries the stop").toHaveLength(13);
  });

  it("calls a pane with no tabs and no buttons CLOSED, which recovery can reopen", () => {
    /**
     * THE ANSWER THE DIAGNOSTIC WAS BUILT FOR, and it arrived on 2026-09-08.
     * `no-run-button` had ended five cycles, every one after a successful
     * recovery and a printed `ready`. When the pane was finally described it
     * was not on the wrong tab — it answered NOTHING, seconds after `host
     * answered in 3ms · slide 1 resolved`.
     *
     * A pane with no tabs and no buttons is a closed pane, and `pane-closed` is
     * already in `RECOVERABLE_STOPS`. Reading it as a missing BUTTON is what
     * made the stop terminal, and it cost the 4:3 leg on attempt 3 of 7.
     */
    expect(paneAnsweredNothing(""), "an empty answer is not a pane").toBe(true);
    expect(paneAnsweredNothing(undefined)).toBe(true);
    // Prose, headings and refs are not controls — the same parse `describePane`
    // uses, so the two can never disagree about what "nothing" means.
    expect(paneAnsweredNothing("### Chart\n- generic: some text\n  - img [ref=f1]")).toBe(true);

    // A pane on the WRONG TAB still has tabs, and must stay `no-run-button`.
    // That one really does want a person, and promoting it to a recoverable
    // stop would spend all seven attempts on a pane that was never going to
    // grow the button.
    expect(paneAnsweredNothing('- tab "Chart" [selected] [ref=f1]'), "a tab list is an answer").toBe(false);
    expect(paneAnsweredNothing('  - button "Insert into slide" [ref=f2]'), "indented controls still count").toBe(false);
  });

  it("keeps what a failed CLI call said, and nothing from one that worked", () => {
    // The stderr is the only place the two browser-absent states differ, so a
    // driver that discards it cannot tell them apart however carefully it reads
    // `list`. Recorded per call, like `lastFailed` beside it.
    const runs = [
      { status: 1, stdout: "", stderr: "Browser is already in use for C:/devtools/pw-profile" },
      { status: 0, stdout: "ok", stderr: "" },
    ];
    let n = 0;
    const sh = cli(() => runs[n++], ".", "C:/fake/playwright-cli.js");

    sh("open");
    expect(sh.state.lastStderr, "the reason the call failed went nowhere").toContain("already in use");
    sh("list");
    expect(sh.state.lastStderr, "a successful call must clear it, not inherit the last failure").toBeNull();
  });

  it("notices a browser that dies UNDER a running round, which the quiet counter cannot", async () => {
    // 24 MINUTES ON 2026-08-16, and the cause is a good guard with a hole in it.
    //
    // `quietStreak` resets to zero whenever a CLI call FAILED — deliberately,
    // because one failed call means nothing was measured, and folding that into
    // "the pane is gone" once killed a healthy round that went on to pass 10 of
    // 12. But a dead browser makes EVERY call fail, permanently: the quiet
    // counter can never reach its threshold, the crash dialog cannot be read
    // either, and the loop polls a corpse until the thirty-minute limit. `pw
    // list` said `(no browsers)` outright while the driver sat there.
    expect(browserDiedMidRound(DEAD_BROWSER_POLLS, "  (no browsers)"), "a dead browser went unnoticed").toBe(true);

    // BOTH CONDITIONS ARE LOAD-BEARING, and each guards against re-making the
    // other's mistake.
    //
    // A streak without the list check would end a round on the absence of
    // evidence — exactly the contention bug above.
    expect(browserDiedMidRound(DEAD_BROWSER_POLLS, "### Browsers\n- ms:\n  - status: open")).toBe(false);
    expect(browserDiedMidRound(DEAD_BROWSER_POLLS, ""), "an unreadable list is not a dead browser").toBe(false);
    // And the list check without a streak would fire on a single blip, which is
    // what a second terminal touching the CLI looks like.
    expect(browserDiedMidRound(1, "  (no browsers)"), "one failed poll is contention, not a death").toBe(false);
    expect(browserDiedMidRound(0, "  (no browsers)")).toBe(false);
  });

  it("retries a browser that died mid-round, because reopening needs no password", () => {
    // The profile keeps the sign-in — a dead browser is not a lost sign-in — so
    // `recover` reopens it unattended. This is also the CHEAPEST of the
    // retryable reasons: it costs the minute already spent, where a wedge costs
    // the full thirty before anyone knows.
    expect(shouldRetry("browser-gone", 0, 3, undefined), "an unattended run stopped dead on a recoverable state").toBe(
      true,
    );
    // Still bounded by the caller's --retry N, like every other reason.
    expect(shouldRetry("browser-gone", 3, 3, undefined)).toBe(false);
  });

  it("sweeps a dirty deck rather than refusing, but only when it is the ONLY thing wrong", () => {
    // A dirty deck is not a fault — it is the last round's slides, and the
    // driver already sweeps them on every recovery. Refusing over it made a
    // person do by hand the one step the machine does better, and by hand is
    // how a deck reached ZERO slides on 2026-08-16: a fixed count of deletes
    // against a deck that held fewer, leaving `slide 1 REFUSED`, which is the
    // state the 2s crash starts from.
    expect(onlyDirtyDeck(["deck-dirty"]), "the one stop the driver can clear itself").toBe(true);

    // ONLY, and that word is the whole guard. Dirty AND stale is a round that
    // would measure the wrong build; healing the cheap half moves it closer to
    // running while still being wrong.
    expect(onlyDirtyDeck(["deck-dirty", "pane-stale"]), "healed its way into measuring the wrong build").toBe(false);
    expect(onlyDirtyDeck(["pane-stale"])).toBe(false);
    // No codes is no evidence — the same rule `shouldRetry` already applies.
    expect(onlyDirtyDeck([])).toBe(false);
    expect(onlyDirtyDeck(undefined)).toBe(false);
  });

  it("refuses a deck that is not the slide size this round was asked for", () => {
    // A CLICK THAT DID NOTHING. Setting Widescreen during the 2026-08-16 control
    // run silently did not take — it landed while the document was in its greyed
    // "Loading" state, the menu accepted it, and nothing changed. It was caught
    // only by reopening the menu and reading which box was ticked.
    //
    // A round that believes it is 4:3 and is not proves nothing, which is the
    // same harm as a round on a stale pane — already a hard stop. With a nightly
    // cycle running 16:9 twice and 4:3 once, an unverified size files a round
    // under the wrong profile, which is worse than not running it.
    const r = readiness({ ...READY, size: "16:9", expectSize: "4:3" });
    expect(r.ok).toBe(false);
    expect(r.codes).toContain("wrong-size");
    expect(r.stop.join(" "), "did not warn that a click can be accepted and do nothing").toMatch(/CHECK IT TOOK/);

    // Matching is silent, and asking for nothing is silent — an ordinary 16:9
    // round must read exactly as it always has.
    expect(readiness({ ...READY, size: "4:3", expectSize: "4:3" }).ok).toBe(true);
    expect(readiness({ ...READY, size: null, expectSize: null }).ok).toBe(true);
    // AND A HOST THAT WOULD NOT ANSWER IS NOT A MISMATCH. Null is no evidence,
    // and refusing on no evidence is the mistake `reachable` exists to prevent.
    expect(readiness({ ...READY, size: null, expectSize: "4:3" }).ok, "refused on an unanswered question").toBe(true);

    // NOT RECOVERABLE, and that is deliberate: `recover` could set the size in
    // two clicks, and doing so would change what the round measures rather than
    // restore it.
    expect(shouldRetry("not-ready", 0, 3, ["wrong-size"]), "recovery tried to change the deck's slide size").toBe(
      false,
    );
  });

  it("reads a slide size the way the add-in does, and calls silence silence", () => {
    expect(readSlideSize("size:960x540")).toBe("16:9");
    expect(readSlideSize("size:720x540")).toBe("4:3");
    // Anything else is its own profile, never folded into the nearest named one.
    expect(readSlideSize("size:1000x500")).toBe("1000x500");
    // A failure or a silence answers nothing — see the null case above.
    expect(readSlideSize("size-failed:GeneralException")).toBe(null);
    expect(readSlideSize("")).toBe(null);
    // Points via pageSetup, which is what the add-in itself reads.
    expect(slideSizeScript(15000)).toContain("slideWidth");
    expect(slideSizeScript(15000), "no budget means a wedged host hangs the check").toContain("15000");
  });

  it("finds the CLI entry beside node, or says it cannot", () => {
    // THE OVERRIDE IS PART OF THIS FUNCTION, so the test has to control it.
    //
    // Leaving it to the ambient environment is what made this test green on CI
    // and RED on the only machine that runs rounds. `PLAYWRIGHT_CLI_JS` is set
    // there; `cliEntry` returns it on its first line, before `exists` is ever
    // consulted; so the `() => false` case never reached the branch it names and
    // the assertion was really testing "nobody exported a variable". That is the
    // same defect as the injected `entry` thirty lines below — learned once,
    // and not carried up the file.
    const saved = process.env.PLAYWRIGHT_CLI_JS;
    try {
      delete process.env.PLAYWRIGHT_CLI_JS;
      expect(cliEntry("C:/n/node.exe", () => true)).toContain("playwright-cli.js");
      expect(cliEntry("C:/n/node.exe", () => false)).toBe(null);
      // And when it IS set it wins, unconditionally — which is what makes it an
      // override rather than a second guess. Asserted, because the unchecked
      // return is load-bearing: it is why a bad path reaches `spawnSync` at all,
      // and `spawnFailureRemedy` is what makes that legible when it does.
      process.env.PLAYWRIGHT_CLI_JS = "C:/elsewhere/playwright-cli.js";
      expect(cliEntry("C:/n/node.exe", () => false)).toBe("C:/elsewhere/playwright-cli.js");
    } finally {
      if (saved === undefined) delete process.env.PLAYWRIGHT_CLI_JS;
      else process.env.PLAYWRIGHT_CLI_JS = saved;
    }
  });

  it("makes the session directory before it spawns anything into it", () => {
    // `.pw-session/` is gitignored, so a FRESH CLONE has none — and spawnSync
    // reports a missing `cwd` as ENOENT against the EXECUTABLE, which reads as a
    // broken node install and used to be answered with "install playwright-cli".
    // A round was lost to that on 2026-08-23 with the tool installed and node on
    // PATH. `scripts/pw.sh` has always done this; the driver assumed someone had
    // sourced it.
    const made: string[] = [];
    const sh = cli(
      () => ({ status: 0, stdout: "" }),
      "C:/nowhere/.pw-session",
      "C:/fake/playwright-cli.js",
      (d: string) => made.push(d),
    );
    sh("list");
    expect(made, "the directory has to be MADE, not assumed").toEqual(["C:/nowhere/.pw-session"]);
  });

  it("survives a session directory it cannot create", () => {
    // A driver that throws out of its own first line reports nothing at all. The
    // spawn below fails anyway, and `spawnFailureRemedy` says why in words.
    expect(() => ensureSessionDir("package.json/nope")).not.toThrow();
  });

  it("records WHICH of the two inputs was absent when a spawn fails", () => {
    // The errno cannot tell these apart — both are ENOENT, and both name the
    // executable. Checked at failure time instead of inferred.
    const sh = cli(
      () => ({ error: new Error("spawnSync C:/Program Files/nodejs/node.exe ENOENT") }),
      ".",
      "C:/fake/playwright-cli.js",
      () => {},
    );
    sh("find", "x");
    expect(sh.state.unreachableAt?.entryMissing, "C:/fake is not on this machine").toBe(true);
    expect(sh.state.unreachableAt?.cwdMissing, '"." is').toBe(false);
  });

  it("names the input that was missing instead of guessing at the install", () => {
    expect(spawnFailureRemedy({ cwdMissing: true, entryMissing: false })).toContain("session directory");
    expect(spawnFailureRemedy({ cwdMissing: true, entryMissing: false })).toContain("scripts/pw.sh");
    expect(spawnFailureRemedy({ cwdMissing: false, entryMissing: true })).toContain("npm i -g @playwright/cli");
    // A missing directory must NOT be answered with an install. That answer cost
    // a round; it is the whole reason this function exists.
    expect(
      spawnFailureRemedy({ cwdMissing: true, entryMissing: false }),
      "the install has never once been the problem",
    ).not.toContain("npm i -g");
    // And when both are present, saying either would be a guess. The comment
    // above the old message already admitted this and the message ignored it.
    const neither = spawnFailureRemedy({ cwdMissing: false, entryMissing: false });
    expect(neither).not.toContain("npm i -g");
    expect(neither).toContain("neither a missing install nor a missing directory");
  });

  it("normalises a session directory, because 8.3 short names are a different session", () => {
    // "C:/Users/DANN~1.PED/..." and "C:/Users/dann.pedersen/..." are one folder
    // and two strings, and the daemon keys its sessions by the string — so from
    // the short-name form it answers "(no browsers)" with the browser open.
    expect(sessionDir("C:/Users/DANN~1.PED/x", () => "C:/Users/dann.pedersen/x")).toBe("C:/Users/dann.pedersen/x");
  });

  it("falls back to the path it was given when it cannot be resolved", () => {
    expect(
      sessionDir("/nope", () => {
        throw new Error("ENOENT");
      }),
    ).toBe("/nope");
  });

  it("keeps a failed spawn's doubt inside the sweep that saw it", () => {
    // ROUND 044. `--retry` cleared the host's crash, `recover` reloaded the tab
    // and reopened the pane, and one of its spawns lost a race. The latch was
    // set for the rest of the PROCESS, so the next attempt read a healthy setup
    // — the host answered in 4ms, on the line above — and refused it with
    // "nothing below was actually measured", which was false of every value it
    // had just printed. One flaky spawn cost the whole round.
    let fail = true;
    const run = () => {
      if (fail) return { error: new Error("EAGAIN") };
      return { status: 0, stdout: "ok" };
    };
    // THE ENTRY IS INJECTED, and leaving it to the default is what turned this
    // test green here and red on CI. `cliEntry()` finds the real playwright-cli
    // on the machine that runs rounds and finds nothing on a CI runner, where
    // every call then takes the no-entry path and latches whatever the fake
    // `run` was going to do. A test about spawn results must not depend on
    // whether a tool is installed.
    const sh = cli(run, ".", "C:/fake/playwright-cli.js");
    // Sticky WITHIN a sweep is deliberate and stays: a call that never ran and a
    // page that answered nothing are the same empty string.
    sh("find", "x");
    expect(sh.state.unreachable).toBe(true);
    fail = false;
    sh("find", "x");
    expect(sh.state.unreachable, "a later success does not retro-clear this sweep").toBe(true);
    // A new sweep is a new question.
    sh.startSweep();
    expect(sh.state.unreachable).toBe(false);
    sh("find", "x");
    expect(sh.state.unreachable, "a sweep whose calls all ran is reachable").toBe(false);
  });

  it("still reports unreachable in every sweep when the tool is not installed at all", () => {
    // The permanent case the latch was written for. With no entry no spawn is
    // attempted at all, and `startSweep` must not paper over it.
    const sh = cli(
      () => {
        throw new Error("should never spawn");
      },
      ".",
      null,
    );
    sh.startSweep();
    sh("find", "x");
    expect(sh.state.unreachable).toBe(true);
  });

  it("does not call an answer too big to hold a tool that could not be run", () => {
    // ROUND 044's ROOT CAUSE. `requests` on a live PowerPoint tab is bigger than
    // spawnSync's 1 MiB default, so it came back ENOBUFS — an error, on a call
    // that ran perfectly. Read as "playwright-cli could not be run", it emptied
    // two crash reports and then refused the retry on a healthy host.
    const enobufs = Object.assign(new Error("spawnSync C:/node.exe ENOBUFS"), { code: "ENOBUFS" });
    expect(isOverflow(enobufs), "the code").toBe(true);
    expect(isOverflow(new Error("spawnSync C:/node.exe ENOBUFS")), "the message, when there is no code").toBe(true);
    expect(isOverflow(new Error("spawn ENOENT")), "a tool that is not there").toBe(false);
    expect(isOverflow(undefined)).toBe(false);

    const sh = cli(() => ({ error: enobufs }), ".", "C:/fake/playwright-cli.js");
    sh("requests");
    expect(sh.state.unreachable, "an overflow is not an unreachable tool").toBe(false);
    expect(sh.state.overflowed).toBe("requests");
    expect(sh.state.lastError).toBe("overflow");
  });

  it("asks for a buffer big enough to hold a live tab's request log", () => {
    // The fix for the cause rather than the symptom. Pinned because the default
    // is invisible: nothing about `spawnSync(...)` says 1 MiB, and the failure it
    // produces names the executable rather than the size.
    let opts: { maxBuffer?: number } = {};
    const sh = cli(
      (_file: string, _args: string[], o: { maxBuffer?: number }) => {
        opts = o;
        return { status: 0, stdout: "" };
      },
      ".",
      "C:/fake/playwright-cli.js",
    );
    sh("requests");
    expect(opts.maxBuffer).toBeGreaterThanOrEqual(64e6);
  });
});

describe("asking the host whether it is awake", () => {
  /**
   * The ping is a STRING sent to another JavaScript engine, so nothing the
   * compiler knows reaches it. These run the generated source against a stub
   * `PowerPoint` — the only way to find out whether what the driver sends is
   * the thing it means to send.
   */
  const runPing = (budgetMs: number, host: unknown) =>
    new Function("PowerPoint", `return (${pingScript(budgetMs)})()`)(host) as Promise<string>;

  it("comes back ok when the host answers", async () => {
    const ctx = { presentation: { slides: { getCount: () => {} } }, sync: async () => {} };
    expect(await runPing(2000, { run: async (cb: (c: unknown) => unknown) => cb(ctx) })).toMatch(/^ok:\d+$/);
  });

  it("comes back, not hangs, when the host never answers", async () => {
    // The whole point. A host that will not answer `getCount()` is exactly the
    // state rounds 24, 25 and 29 each spent most of an hour inside, and an
    // unraced `PowerPoint.run` would sit here as long as it sat there.
    //
    // Verdict only, never the number: Windows' clock quantises to ~15.6ms, so
    // asserting elapsed milliseconds would flake for reasons unrelated to this.
    const out = await runPing(20, { run: () => new Promise(() => {}) });
    expect(out).toMatch(/^no:\d+$/);
  });

  it("comes back when the host throws instead of hanging", async () => {
    // A 5010 or a torn-down proxy is a NO, not a crash in the driver.
    expect(
      await runPing(2000, {
        run: async () => {
          throw new Error("GeneralException");
        },
      }),
    ).toMatch(/^no:\d+$/);
  });

  it("takes the ref off the matching line, not the first one in the frame tree", () => {
    // What made the first live ping lie. `find` prints the frame hierarchy above
    // the hit, so the first ref in its output belongs to the OUTER iframe — the
    // OneDrive document, where `Office` and `PowerPoint` are both undefined.
    // Evaluating there returns "no" for every host, healthy or not, which is the
    // worst possible failure for a precondition: it refuses correct work and
    // teaches the reader to ignore it.
    const found = [
      "iframe [ref=f1a2b3c]",
      "  iframe [ref=f4d5e6f]",
      '    checkbox "Verbose trace" [checked] [ref=f7a8b9c]',
    ].join("\n");
    const sh = (() => found) as unknown as Parameters<typeof refFor>[0];
    expect(refFor(sh, "Verbose trace", /checkbox "Verbose trace"/)).toBe("f7a8b9c");
  });

  it("has no ref when nothing matched", () => {
    const sh = (() => "iframe [ref=f1a2b3c]") as unknown as Parameters<typeof refFor>[0];
    expect(refFor(sh, "Verbose trace", /checkbox "Verbose trace"/)).toBe(null);
  });

  it("reads a verdict back out of the CLI's quoting, and nothing out of noise", () => {
    expect(readPing('"ok:37"')).toEqual({ answered: true, ms: 37 });
    expect(readPing("no:8001")).toEqual({ answered: false, ms: 8001 });
    expect(readPing(""), "an empty result is not a dead host — see `cli`").toBe(null);
    expect(readPing("undefined")).toBe(null);
  });

  it("names the crash when PowerPoint's own dialog is up", () => {
    // The wedge, finally identified. Four rounds died against it and the report
    // they got was a 90-second timeout on whatever call came next; what had
    // actually happened is that PowerPoint crashed and put up a modal, behind
    // which every Office.js call — including an empty sync — hangs forever
    // without throwing. Saying so is the difference between a refusal the reader
    // can act on and one they wait out.
    const r = readiness({ ...READY, crashed: true, ping: { answered: false, ms: 8002 } });
    expect(r.ok).toBe(false);
    expect(r.stop[0], "the crash outranks the silence it causes").toContain("Refresh");
    expect(r.stop.join(" ")).toContain("Add-ins");
  });

  it("does not see a dialog in find's own echo of the query", () => {
    // `find` answers a miss with `No matches found for "<query>"` — the query
    // included. A detector that tested for the phrase it searched for would
    // report a crash on every healthy host forever. Same shape as the ref
    // `buildOf` used to read out of its own haystack; that one cost a round.
    expect(sawCrashDialog('No matches found for "Sorry, we ran into a problem".')).toBe(false);
    // The echo with the word `dialog` INSIDE the query, which is what a
    // contains-the-word check gets wrong. The first version of this test used a
    // query without it, so it passed against a detector that had no guard at
    // all — green, and proving nothing.
    expect(sawCrashDialog('No matches found for "crash dialog".')).toBe(false);
    expect(sawCrashDialog(""), "unreachable CLI is not a crash — see `reachable`").toBe(false);
    expect(sawCrashDialog('- dialog [ref=f21e737]:\n  - button "Refresh" [ref=f21e750]')).toBe(true);
  });

  it("does not end a running round because the CLI lost a race", () => {
    // What killed round 29's first archive-worthy run: the CLI serves one
    // command per session, an agent read the trace while the driver polled, one
    // poll exited non-zero, and the empty string that came back was reported as
    // "PowerPoint has probably crashed". The round was fine — it went on to pass
    // 10 of 12 scenarios. A failed call is not a silent page.
    expect(quietStreak(1, "", true), "a failed call measured nothing").toBe(0);
    expect(quietStreak(1, '- button "Download run log" [disabled]', false)).toBe(0);
    // Genuinely silent, twice, is the real thing — and it has to reach two.
    expect(quietStreak(0, "", false)).toBe(1);
    expect(quietStreak(1, "   ", false)).toBe(2);
  });

  it("retries what recovery can actually put right", () => {
    // A crash is the state the driver has always put right on its own, six times
    // in one night. The rest of this is 2026-08-15, when retrying ONLY a crash
    // cost two rounds: mid-round the quiet wedge exits as `silent` and the driver
    // went home, though `docs/ROUNDS.md` says a reload clears both forms of the
    // wedge and `recover` already does exactly that.
    expect(shouldRetry("crashed", 0, 3)).toBe(true);
    expect(shouldRetry("crashed", 3, 3), "the budget is a budget").toBe(false);
    expect(shouldRetry("silent", 0, 3), "the quiet wedge is a wedge").toBe(true);
    // A WEDGE MID-ROUND, which stopped an unattended run dead in its first hour
    // on 2026-08-15: the round timed out at thirty minutes, `recover` could have
    // cleared it in eighty seconds, and `timeout` was not on the list. The cost
    // of retrying is bounded by the `--retry` the caller chose; the cost of not
    // retrying was nine idle hours.
    expect(shouldRetry("timeout", 0, 3), "a wedged round is recoverable").toBe(true);
    expect(shouldRetry("timeout", 3, 3), "and still bounded by the budget").toBe(false);
    expect(shouldRetry("finished", 0, 3)).toBe(false);
  });

  it("retries a check whose every refusal a reload would clear, and no other", () => {
    // The hand-recovery of 2026-08-15, encoded. `--check` refused with a silent
    // host, a pane one build behind and an eight-slide deck; a person then did
    // the five steps `recover` does and every one of them went green. The driver
    // had the same five steps and would not take them.
    expect(shouldRetry("not-ready", 0, 3, ["host-silent", "pane-stale", "deck-dirty"])).toBe(true);
    expect(shouldRetry("not-ready", 0, 3, ["crashed"])).toBe(true);
    expect(shouldRetry("not-ready", 0, 3, ["slide-refused"])).toBe(true);

    // And the refusals a reload does NOT touch, which is what the old comment
    // was right about: retrying a stale build just repeats itself until the
    // night is gone. ONE unrecoverable code is enough to stop, even beside three
    // recoverable ones.
    expect(shouldRetry("not-ready", 0, 3, ["site-behind"]), "waiting on Pages is not a reload").toBe(false);
    expect(shouldRetry("not-ready", 0, 3, ["verbose-off"]), "a pane toggle is a person's choice").toBe(false);
    expect(shouldRetry("not-ready", 0, 3, ["host-silent", "deck-dirty", "site-behind"])).toBe(false);

    // No codes is not a licence: it means nothing was recorded about why the
    // check refused, and retrying on no evidence is how a loop spins.
    expect(shouldRetry("not-ready", 0, 3, []), "an empty list is not a reason to retry").toBe(false);
    expect(shouldRetry("not-ready", 0, 3), "and neither is no list at all").toBe(false);
    expect(shouldRetry("not-ready", 3, 3, ["deck-dirty"]), "the budget still binds").toBe(false);
  });

  it("names what it is recovering from, rather than announcing a crash that did not happen", () => {
    // ROUND 047 refused on a dirty deck alone and was told "clearing the crash
    // and starting again". The line was written when a crash was the only thing
    // retried; the moment the retry covered more, it began inventing one. A
    // recovery line is read while someone is debugging.
    expect(recoveryFor("crashed", ["crashed"])).toContain("clearing the crash");
    expect(recoveryFor("silent", undefined)).toContain("went quiet");
    expect(recoveryFor("timeout", undefined)).toContain("wedged");
    expect(recoveryFor("not-ready", ["deck-dirty"]), "the round-047 line").toBe(
      "recovering from a dirty deck, then starting again",
    );
    expect(recoveryFor("not-ready", ["host-silent", "pane-stale", "deck-dirty"])).toBe(
      "recovering from a silent host and a stale pane and a dirty deck, then starting again",
    );
    // An unrecoverable code never reaches this line (shouldRetry stops first),
    // so it must not be named as something being recovered from.
    expect(recoveryFor("not-ready", ["site-behind"])).toBe("recovering and starting again");
    expect(recoveryFor("not-ready", [])).toBe("recovering and starting again");
  });

  it("gives every refusal a code, and every code a meaning", () => {
    // The codes are what the retry decision reads, so a stop that forgets one is
    // a stop the driver will treat as no reason at all. Asserted against the
    // real function rather than a list, because a hand-kept list is the thing
    // that goes stale.
    const wedged = readiness({ ...READY, ping: { answered: false, ms: 8011 }, slides: 8, stamp: "old1234" });
    expect(wedged.ok).toBe(false);
    expect(wedged.codes).toEqual(["host-silent", "pane-stale", "deck-dirty"]);
    expect(wedged.codes.length, "a message without a code is invisible to the retry").toBe(wedged.stop.length);
    expect(shouldRetry("not-ready", 0, 3, wedged.codes), "the state a person fixed by hand").toBe(true);

    const stale = readiness({ ...READY, head: "abc1234", deployed: "def5678", stamp: "def5678" });
    expect(stale.codes).toEqual(["site-behind"]);
    expect(shouldRetry("not-ready", 0, 3, stale.codes)).toBe(false);
  });

  it("recovers in the order the host needs, not the order that reads well", async () => {
    // The sequence is load-bearing: Refresh reloads the document and closes the
    // pane, so the pane must be reopened AFTER it, the Automation tab must be
    // showing before the run button exists, and the deck has to be cleaned last
    // because cleaning needs the pane's frame to evaluate in.
    const calls: string[] = [];
    const sh = ((...args: string[]) => {
      calls.push(args[0] + (args[1] ? ` ${args[1].slice(0, 22)}` : ""));
      if (args[0] === "find" && args[1] === "Sorry, we ran into a problem") return "- dialog [ref=f1]:";
      if (args[0] === "find") return `button "${args[1]}" [ref=f2]\ntab "${args[1]}" [ref=f2]`;
      return "ok";
    }) as unknown as Parameters<typeof recover>[0];

    await recover(sh, async () => {});

    expect(calls.filter((c) => c.startsWith("eval")).length, "clicked and cleaned").toBeGreaterThan(2);
    const order = calls.join(" | ");
    expect(order.indexOf("find Insert chart"), "the pane is reopened after the reload").toBeGreaterThan(
      order.indexOf("find Sorry"),
    );
    expect(order.indexOf("find Automation")).toBeGreaterThan(order.indexOf("find Insert chart"));
  });

  it("goes to the file list before hunting for a deck that is not open", async () => {
    /**
     * THE LINK IS ONLY ON THE FILE LIST, and this branch searched whatever tab
     * happened to be current. When the deck is missing that is normally the other
     * leg's presentation, and a presentation page carries no
     * `link "PresentationNN"` — so the search found nothing, nothing was clicked,
     * and recovery slept 25 seconds and gave up. An absent ref is not an error,
     * so it did that silently, on every attempt.
     *
     * Measured on the live browser on 2026-08-29 with the deck genuinely absent:
     * "No matches found" from the presentation tab, "Found 2 matches" from the
     * OneDrive tab. It cost a full cycle — seven attempts, all `deck-missing`.
     *
     * It had been masked. This stop used to arrive beside `wrong-size` — the size
     * check read whichever deck WAS fronted — and `wrong-size` is deliberately
     * unrecoverable, so the cycle died on attempt one and recovery never reached
     * here. Fixing that guard this morning is what uncovered this one.
     */
    const calls: string[][] = [];
    const sh = ((...args: string[]) => {
      calls.push(args);
      // Two tabs, neither of them the wanted deck, the current one a presentation.
      if (args[0] === "tab-list")
        return `- 0: [Home - OneDrive](${DECK_HOME_URL})\n- 1: (current) [Presentation70.pptx](https://onedrive.live.com/x)`;
      if (args[0] === "find" && args[1] === "Sorry, we ran into a problem") return 'No matches found for "…".';
      if (args[0] === "find") return `link "${args[1]}" [ref=f9]`;
      return "ok";
    }) as unknown as Parameters<typeof recover>[0];

    process.env.PW_DECK = "Presentation64";
    try {
      await recover(sh, async () => {});
    } finally {
      delete process.env.PW_DECK;
    }
    const flat = calls.map((c) => c.join(" "));
    const selected = flat.indexOf("tab-select 0");
    const searched = flat.findIndex((c) => c.startsWith("find Presentation64"));
    expect(selected, "never fronted the file list").toBeGreaterThanOrEqual(0);
    expect(searched, "never looked for the deck at all").toBeGreaterThanOrEqual(0);
    expect(selected, "searched for the deck link before fronting the list that holds it").toBeLessThan(searched);
    // An existing OneDrive tab is REUSED rather than navigated to: a `goto` on
    // the current tab would close a presentation this cycle still needs.
    expect(flat, "opened a new tab when one was already on the file list").not.toContain(`tab-new ${DECK_HOME_URL}`);
  });

  it("reloads when there is no dialog to clear", async () => {
    // The quiet form of the wedge — no modal, session severed by the network. A
    // recovery that only knew how to click Refresh would do nothing at all here.
    const calls: string[] = [];
    const sh = ((...args: string[]) => {
      calls.push(args[0]);
      if (args[0] === "find" && args[1] === "Sorry, we ran into a problem") return 'No matches found for "…".';
      return `button "x" [ref=f2]`;
    }) as unknown as Parameters<typeof recover>[0];

    await recover(sh, async () => {});
    expect(calls).toContain("reload");
  });

  it("says SIGNED OUT rather than blaming a pane that was never there", async () => {
    // The state the owner walks back into after the browser dies. A signed-out
    // browser answers every pane read with nothing, which looks exactly like an
    // add-in nobody opened — and the check duly said "is the add-in open?" while
    // the tab was showing login.live.com, sending the reader to hunt for a pane
    // in a window with no document in it.
    //
    // It also short-circuits, for the same reason `reachable` does: if this is
    // true then nothing below it was measured and every other line is noise.
    const r = readiness({ ...READY, stamp: null, slides: null, loggedOut: true });
    expect(r.ok).toBe(false);
    expect(r.stop, "said more than the one thing that is true").toHaveLength(1);
    expect(r.stop[0]).toContain("sign-in page");
    expect(r.stop[0], "the reader has to know this one is theirs").toContain("needs a password");
  });

  it("knows a sign-in page from a document", () => {
    expect(signedOut("- 0: (current) [Sign in](https://login.live.com/login.srf?wa=wsignin1)")).toBe(true);
    expect(signedOut("- 0: (current) [x](https://login.microsoftonline.com/common/oauth2)")).toBe(true);
    expect(signedOut("- 0: (current) [Presentation63.pptx](https://onedrive.live.com/personal/x/doc.aspx)")).toBe(
      false,
    );
    expect(signedOut(""), "no tabs is not a sign-in page").toBe(false);
  });

  it("knows a sign-in POPUP beside a live deck from a browser sitting on a login page", () => {
    // What actually ended the overnight run of 2026-08-15/16. Office opened an
    // auth prompt NEXT TO a deck tab that was still there, and the driver
    // reported "the browser is on a Microsoft sign-in page" — while the screen
    // showed PowerPoint with a small dialog over it.
    //
    // Stopping the round is right either way: if the host is asking for
    // credentials, nothing measured past that point can be trusted. What was
    // wrong is a message that does not match what the reader sees, which on an
    // overnight run is the only account of the night they get.
    const popup = [
      "- 0: [Presentation63.pptx](https://onedrive.live.com/personal/x/doc.aspx)",
      "- 1: (current) [Sign in](https://login.live.com/login.srf?wa=wsignin1)",
    ].join("\n");
    expect(signInIsPopup(popup), "a deck tab was still open and this called it a login page").toBe(true);
    // The whole browser on a login page — no document tab anywhere.
    expect(signInIsPopup("- 0: (current) [Sign in](https://login.live.com/login.srf)")).toBe(false);
    // And it must not fire on a healthy deck with no prompt at all.
    expect(signInIsPopup("- 0: (current) [Presentation63.pptx](https://onedrive.live.com/x)")).toBe(false);
    expect(signInIsPopup("")).toBe(false);

    // The messages are genuinely different, and each names what the reader can
    // see. Both hand the job back — it needs a password either way.
    const asPopup = readiness({ ...READY, stamp: null, slides: null, loggedOut: true, authPopup: true });
    const asPage = readiness({ ...READY, stamp: null, slides: null, loggedOut: true });
    expect(asPopup.stop[0], "the popup case still described a browser on a login page").toContain("beside the deck");
    expect(asPage.stop[0]).toContain("sign-in page");
    expect(asPopup.stop[0]).not.toBe(asPage.stop[0]);
    for (const r of [asPopup, asPage]) {
      expect(r.ok, "a host asking for credentials cannot be measured either way").toBe(false);
      expect(r.stop[0], "the reader has to know this one is theirs").toContain("needs a password");
    }
  });

  it("refuses a round when the host will not answer the cheapest call there is", () => {
    const r = readiness({ ...READY, ping: { answered: false, ms: 8001 } });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("8001ms");
  });

  it("runs when the host answers, and when nobody asked it", () => {
    // Not asked is not unwell. The pane can be closed at check time for reasons
    // the driver already reports separately, and a second hard stop built on a
    // reading that was never taken would just be noise.
    expect(readiness({ ...READY, ping: { answered: true, ms: 120 } }).ok).toBe(true);
    expect(readiness({ ...READY, ping: null }).ok).toBe(true);
  });
});

describe("archive housekeeping", () => {
  it("refuses to file the same log twice under a new number", async () => {
    // Round 042 wedged, `Download run log` was disabled so the click did
    // nothing, and the PREVIOUS round's file was still sitting at the same path.
    // It was archived as 039 and was byte-identical to 038 — a whole round of
    // evidence that never happened, caught by a checksum and nothing else.
    //
    // A fabricated round is the worst thing that directory can hold. Everything
    // downstream pools these files — verdict histories, the rasterise arms, the
    // scenario flip detector — so one duplicate silently doubles the weight of
    // whatever the real round happened to say.
    const round = { build: "abc1234 · 2026-08-15 05:00Z", trace: { entries: [] } };
    const body = `${JSON.stringify(round, null, 2)}\n`;
    const read = () => body;
    expect(() =>
      archive("log.json", "rounds", read as never, (() => {}) as never, (() => ["038-abc1234.json"]) as never),
    ).toThrow(/byte-identical to 038-abc1234\.json/);
  });

  it("does not hand out a number a round committed elsewhere already has", () => {
    // WHAT ACTUALLY WENT WRONG on 2026-08-16, and it is not an overwrite. Round
    // 064 was archived on `main` and its findings committed to a branch;
    // `git checkout main` then removed the file from the working tree, because
    // it is tracked only on that branch. The next round was archived from
    // `main`, where the directory ends at 063, and was numbered 064 as well.
    //
    // Nothing was destroyed — `nextRoundNumber` is max+1, so it never lands on a
    // file it can SEE. The damage is two different rounds both called 064 in two
    // git contexts: a merge collision, and a pooled report that reads one of
    // them twice or not at all. Every downstream number here is a pool over this
    // directory.
    //
    // The fix is what the lister is given, so that is what this pins.
    const round = { build: "abc1234 · 2026-08-16 05:00Z", trace: { entries: [] } };
    // Distinct bodies, or the byte-identical twin guard fires first and this
    // passes without ever reaching the numbering under test.
    const read = ((p: string) =>
      p === "log.json"
        ? `${JSON.stringify(round, null, 2)}\n`
        : `${JSON.stringify({ ...round, trace: { entries: [{ message: p }] } }, null, 2)}\n`) as never;
    const written: string[] = [];
    const write = ((p: string) => written.push(p)) as never;

    // The working tree alone — the bug. It cannot see 064 and reissues it.
    archive("log.json", "rounds", read as never, write, (() => ["063-9dc3bc8.json"]) as never);
    expect(written[0], "the working tree alone reissued a number git already holds").toBe("rounds/064-abc1234.json");

    // The union of disk and git — the fix. 064 is committed on a branch, so the
    // next round is 065 even though the file is nowhere on disk.
    written.length = 0;
    archive("log.json", "rounds", read as never, write, (() => ["063-9dc3bc8.json", "064-bcd5773.json"]) as never);
    expect(written[0], "a round committed on another branch was numbered over").toBe("rounds/065-abc1234.json");
  });

  it("refuses to archive the PREVIOUS round's log when a wedge left the download doing nothing", () => {
    // THE FAILURE THE TWIN CHECK CANNOT SEE. `Download run log` is DISABLED
    // while a round is running, so clicking it after one that wedged does
    // nothing at all — and the previous round's file is still sitting at the
    // same path, waiting to be filed under a new number. Round 039 was archived
    // byte-identical to 038 that way: a whole round of evidence that never took
    // place, caught by a checksum and nothing else.
    //
    // The twin check only catches it when the two logs are IDENTICAL. A stale
    // log that merely DIFFERS — an older round, an earlier build — sails past
    // it. The pane's build stamp is what closes that, and it costs nothing: the
    // driver already read it to decide the round was worth running.
    const round = { build: "aaaaaaa · 2026-08-16 05:00Z", trace: { entries: [] } };
    const read = () => `${JSON.stringify(round, null, 2)}
`;
    const written: string[] = [];
    const write = ((p: string) => written.push(p)) as never;
    expect(
      () => archive("log.json", "rounds", read as never, write, (() => []) as never, "bbbbbbb"),
      "filed a log from a different build as this round",
    ).toThrow(/is the PREVIOUS round's file/);
    expect(written, "it refused and wrote anyway").toEqual([]);

    // And the matching case still files, or the guard would block every round.
    archive("log.json", "rounds", read as never, write, (() => []) as never, "aaaaaaa");
    expect(written[0]).toBe("rounds/001-aaaaaaa.json");
  });

  it("files a log that differs from everything already kept", () => {
    const round = { build: "abc1234 · 2026-08-15 05:00Z", trace: { entries: [] } };
    const written: string[] = [];
    const name = archive(
      "log.json",
      "rounds",
      ((p: string) => (String(p) === "log.json" ? JSON.stringify(round) : "{}")) as never,
      ((p: string) => written.push(String(p))) as never,
      (() => ["038-abc1234.json"]) as never,
    );
    expect(name).toBe("039-abc1234.json");
    expect(written).toHaveLength(1);
  });

  it("reads a build out of a stamp, and nothing out of prose", () => {
    expect(buildOf("32a6987 · 2026-08-14 08:03Z")).toBe("32a6987");
    expect(buildOf("no build here")).toBe(null);
  });

  it("does not mistake a playwright ref for a build", () => {
    // `f14e735` is an element ref, and the accessibility dump this reads is full
    // of them. Taking a bare 7-hex token reported the pane as showing a build
    // the site had never served and refused to start a round on a pane that was
    // showing exactly the right commit.
    expect(buildOf('button "Insert chart" [ref=f14e735]')).toBe(null);
    expect(buildOf("generic [ref=f9e12]: 2601703 · 2026-08-14 13:58Z")).toBe("2601703");
  });

  it("numbers the next round above the highest already kept", () => {
    expect(nextRoundNumber(["023-0e22b31.json", "028-32a6987.json", "README.md"])).toBe("029");
    expect(nextRoundNumber([])).toBe("001");
  });

  it("takes the highest number, not the last one listed", () => {
    // `readdirSync` order is not guaranteed, and taking the last entry would
    // renumber a round on top of an existing one the first time it differed.
    expect(nextRoundNumber(["028-32a6987.json", "023-0e22b31.json"])).toBe("029");
  });

  it("strips slide images but leaves the evidence alone", () => {
    const round = {
      build: "abc1234",
      deck: { picture: "A".repeat(3000) },
      hostAnswers: { answers: [{ id: "a", answer: "yes", detail: "short" }] },
    };
    const out = stripImages(structuredClone(round));
    expect(out.deck.picture).toMatch(/^<image stripped/);
    expect(out.hostAnswers.answers[0].detail, "a short string is not an image").toBe("short");
  });
});

/**
 * The driver exits 0 or 1 and nothing else, so whatever runs the next round can
 * either be told what happened or go parsing sentences. `rounds-gate.mjs`
 * already refused to parse prose once; this is the same refusal with a file
 * behind it.
 */
describe("the account the driver leaves of how a round ended", () => {
  it("says which file the round produced, not merely that one was produced", () => {
    // "The newest file in rounds/" is the obvious substitute and it is the
    // assumption that produced a wrong overwrite diagnosis in this repo once.
    const r = outcomeReceipt({ reason: "finished", codes: [], roundFile: "082-0f7eadc.json", build: "0f7eadc" });
    expect(r.roundFile).toBe("082-0f7eadc.json");
    expect(r.build).toBe("0f7eadc");
    expect(r.reason).toBe("finished");
  });

  it("decides recoverability with the same set the driver retries on", () => {
    // Two implementations of "is this worth another attempt" is one too many.
    const one = [...RECOVERABLE_STOPS][0];
    expect(outcomeReceipt({ reason: "not-ready", codes: [one] }).recoverable).toBe(true);
    expect(outcomeReceipt({ reason: "not-ready", codes: [one, "wrong-size"] }).recoverable).toBe(false);
    // `wrong-size` must never read as recoverable: recovery COULD set a slide
    // size, and doing so would change what the round measures rather than
    // restore it.
    expect(RECOVERABLE_STOPS.has("wrong-size")).toBe(false);
  });

  it("hands back an array of codes even when there were none", () => {
    // A reader doing `codes.includes(...)` on a finished round should get
    // `false`, not a crash.
    expect(outcomeReceipt({ reason: "finished" }).codes).toEqual([]);
    expect(outcomeReceipt({ reason: "finished" }).recoverable).toBe(false);
    expect(outcomeReceipt({ reason: "crashed", codes: undefined }).codes).toEqual([]);
  });

  it("records the slide size in the shape the driver actually produces", () => {
    // THE PROFILE STRING, not an object. This test used to hand in
    // `{ width, height }` and assert it came back — which tested the
    // assumption, not the driver. `readSlideSize` returns "16:9", "4:3", or
    // "960x540", so the receipt recorded `"size": {}` on every cycle leg and
    // the field naming the arm a round belonged to said nothing at all.
    for (const [raw, profile] of [
      ["960x540", "16:9"],
      ["720x540", "4:3"],
      // Neither named ratio — a custom deck must be visibly its own profile
      // rather than folded into whichever one it is nearest.
      ["1000x500", "1000x500"],
    ]) {
      expect(outcomeReceipt({ reason: "finished", size: readSlideSize(`size:${raw}`) }).size).toBe(profile);
    }
    expect(outcomeReceipt({ reason: "finished" }).size).toBeNull();
  });

  it("numbers a round against every round ever filed, not just the working tree", () => {
    // TWO ROUNDS WERE BOTH FILED AS 064 on 2026-08-16. `everyRoundEverFiled`
    // was written for it and wired to the `--archive` subcommand — the path a
    // person uses by hand. `collectRound`, which archives EVERY round, went on
    // passing `readdirSync`: the working tree alone, blind to a round committed
    // on a branch that is not checked out, which is the exact collision.
    //
    // Guarded at the DEFAULT rather than at a call site, so the next caller
    // cannot repeat the omission. Read off the function itself because the
    // behaviour is a default value, and a test that passed its own lister —
    // as every other test here does — could never see it.
    const src = String(driver.archive);
    expect(src, "the automatic path can reuse a number again").toMatch(/list\s*=\s*everyRoundEverFiled/);
    expect(src, "readdirSync as the default is what filed 064 twice").not.toMatch(/list\s*=\s*readdirSync/);
  });

  it("files a round even when the archive names one that git has and the disk does not", () => {
    // THE TWIN CHECK OPENS EVERY NAME THE LISTER RETURNS, and since
    // `everyRoundEverFiled` became the default that list deliberately includes
    // rounds committed on a branch nobody has checked out. Unguarded it threw
    // ENOENT on exactly the HEALTHY path — `.find` reaches a git-only name only
    // when nothing matched, which is the case of a genuinely new round — and
    // `collectRound` swallowed it, so the driver exited 0 claiming a finished
    // round with nothing written and the next leg overwrote the evidence.
    const onDisk = '{\n  "build": "aaa1111 · x"\n}\n';
    const read = ((p: string) => {
      if (p.endsWith("log.json")) return JSON.stringify({ build: "ccc3333 · a fresh round" });
      if (p.endsWith("063-aaa1111.json")) return onDisk;
      // The git-only name: on nobody's disk, so opening it throws exactly as
      // `readFileSync` does.
      throw Object.assign(new Error("ENOENT: no such file or directory"), { code: "ENOENT" });
    }) as never;
    const written: string[] = [];
    const write = ((p: string) => written.push(p)) as never;
    const list = (() => ["063-aaa1111.json", "064-bbb2222.json"]) as never;
    const name = archive("log.json", "rounds", read, write, list);
    // Numbered past the round git knows about, which is the whole reason the
    // union lister is the default.
    expect(name).toBe("065-ccc3333.json");
    expect(written).toHaveLength(1);
  });

  it("still refuses a byte-identical twin that IS on disk", () => {
    // The guard the fix must not cost: a log the pane never rewrote means the
    // round did not finish, and filing it would double the weight of whatever
    // the real round said.
    const body = '{\n  "build": "aaa1111 · x"\n}\n';
    const read = ((p: string) => {
      if (p.endsWith("log.json")) return '{"build":"aaa1111 · x"}';
      if (p.endsWith("063-aaa1111.json")) return body;
      throw Object.assign(new Error("ENOENT"), { code: "ENOENT" });
    }) as never;
    expect(() =>
      archive(
        "log.json",
        "rounds",
        read,
        (() => {}) as never,
        (() => ["063-aaa1111.json", "064-bbb2222.json"]) as never,
      ),
    ).toThrow(/byte-identical to 063-aaa1111\.json/);
  });

  it("names a round that threw where nothing expected it to", () => {
    // An unexpected exception used to kill the process outright: no receipt, no
    // retry, and a night that ended on its first surprise.
    const r = outcomeReceipt({ reason: "threw", codes: [], threw: "Cannot read properties of undefined" });
    expect(r.reason).toBe("threw");
    expect(r.threw).toMatch(/Cannot read properties/);
    // RECOVERY DOES ADDRESS IT, and the receipt has to say so. `recoverable`
    // answers one question — did the driver retry this — and the cycle reads it
    // only to choose its wording; it stops either way. So the honest answer is
    // the one `shouldRetry` gives.
    //
    // This asserted `false` while `recoverable` looked at the CODES alone, and
    // most stops carry none: a crash, a wedge, a dead browser and a throw are
    // all retried on their reason. A night that recovered from a crash six times
    // ended by printing "crashed is not something recovery addresses — it needs
    // a person", directly under six lines saying "clearing the crash and
    // starting again".
    expect(r.recoverable).toBe(true);
  });

  it("agrees with shouldRetry about what recovery addresses", () => {
    // One question, one implementation. Two of these carry no codes at all, so a
    // receipt reading only `codes` gets every one of them wrong.
    for (const reason of ["crashed", "silent", "timeout", "browser-gone", "threw"])
      expect(outcomeReceipt({ reason, codes: [] }).recoverable, `${reason} is retried by the driver`).toBe(true);
    // And a stop recovery genuinely cannot touch still reads false.
    expect(outcomeReceipt({ reason: "not-ready", codes: ["wrong-size"] }).recoverable).toBe(false);
    expect(outcomeReceipt({ reason: "finished", codes: [] }).recoverable).toBe(false);
  });

  it("leaves `threw` off a round that did not throw", () => {
    expect(outcomeReceipt({ reason: "finished" })).not.toHaveProperty("threw");
  });

  it("retries a round that threw, and stops once the retries are spent", () => {
    // Bounded by the caller's own --retry, so a deterministic bug fails that
    // many times and stops rather than spinning all night.
    expect(shouldRetry("threw", 0, 3, [])).toBe(true);
    expect(shouldRetry("threw", 3, 3, [])).toBe(false);
  });
});

describe("the build under test is pinned for the whole round", () => {
  it("does not ask git for HEAD when the caller pinned it", async () => {
    // A ROUND TESTS ONE BUILD. `attempt` used to re-read `git HEAD` on every
    // retry, so the identity of the thing being measured could change while it
    // was being measured — and it did: a commit made in this clone while a
    // round was retrying moved HEAD ahead of the deployed site, and the driver
    // then refused its own round as `site-behind` against a commit that had
    // never been deployed and was not what it was testing. Twice on 2026-08-21.
    //
    // `main` reads it once and hands it down; this pins the half that carries
    // the risk, since the other half is a single read at the top of the loop.
    const gitCalls: string[][] = [];
    const run = (cmd: string, args: string[]) => {
      gitCalls.push([cmd, ...args]);
      return { stdout: "", status: 0 };
    };
    const sh = Object.assign(() => "", {
      startSweep: () => {},
      state: { unreachable: false },
      dir: ".",
    });
    await (driver as unknown as { attempt: (...a: unknown[]) => Promise<unknown> }).attempt(
      ["--check"],
      { run, head: "abc1234", fetchBuild: async () => "", sleep: async () => {} },
      sh,
    );
    expect(
      gitCalls.filter((c) => c[1] === "rev-parse"),
      "attempt re-read HEAD despite being handed a pinned build",
    ).toHaveLength(0);
  });
});

describe("what the driver had to do to get the round", () => {
  const log = { build: "abc1234 · 2026-08-21 10:00Z", selftest: [], trace: { entries: [] } };
  const read = (() => `${JSON.stringify(log, null, 2)}\n`) as never;

  it("records the attempts and the stops that forced them", () => {
    // A SUCCESSFUL RECOVERY ERASES ITS OWN EVIDENCE. Once `recover` works, the
    // round looks exactly like one that never needed it — so a round run
    // against a host that was already unwell was indistinguishable from a clean
    // one, for all 149 rounds on file.
    //
    // Round 148 is why this exists: three attempts (a silent host, then a
    // closed pane), then two scenarios failed that had not failed once in 109
    // rounds, with no app-code change to explain it. "Was the host sick?" is
    // the first question and nothing archived could answer it.
    let out = "";
    const write = ((_p: string, body: string) => {
      out = body;
    }) as never;
    archive(
      "log.json",
      "rounds",
      read,
      write,
      (() => []) as never,
      null as never,
      null as never,
      { attempts: 3, recovered: ["host-silent", "pane-closed"] } as never,
    );
    const filed = JSON.parse(out) as { driverRun?: { attempts: number; recovered: string[] } };
    expect(filed.driverRun?.attempts, "the round does not say how many attempts it took").toBe(3);
    expect(filed.driverRun?.recovered, "the stops that forced the retries were not recorded").toEqual([
      "host-silent",
      "pane-closed",
    ]);
  });

  it("writes no driverRun at all when the driver did not say", () => {
    // An ABSENT reading must not read as "no recoveries". Rounds archived before
    // this existed carry no field, and a default of zero would quietly claim
    // every one of them ran clean.
    let out = "";
    const write = ((_p: string, body: string) => {
      out = body;
    }) as never;
    archive("log.json", "rounds", read, write, (() => []) as never);
    expect(JSON.parse(out), "invented a clean run for a round that never reported one").not.toHaveProperty("driverRun");
  });

  it("names the document the round ran against, and separates a look from a silence", () => {
    /**
     * WHICH DECK. 0 of the 408 files under `rounds/` and `crashes/` named one,
     * so deck identity and slide profile are perfectly confounded across the
     * whole archive: every 4:3 round ran one document and every 16:9 round
     * another. The sharpest contrast this repo owns — 4:3 crashing on 52 of 73
     * attempts against 0 of 51 at 16:9, over the 25 builds that ran both arms —
     * is therefore equally consistent with "that one file is sick", and no round
     * already filed can settle it.
     */
    const filedWith = (...extra: unknown[]) => {
      let out = "";
      const write = ((_p: string, body: string) => {
        out = body;
      }) as never;
      archive("log.json", "rounds", read, write, (() => []) as never, null as never, ...(extra as never[]));
      return JSON.parse(out) as { driverDeck?: string | null; driverSlideSize?: string | null };
    };

    expect(filedWith(null, null, "Presentation70.pptx").driverDeck).toBe("Presentation70.pptx");

    /**
     * THREE STATES, NOT TWO, and this is the assertion that makes the field
     * worth having. `null` means the driver looked at the tab list and it would
     * not say; the field being ABSENT means nobody offered a reading at all —
     * an older build, or `--archive` run by hand, which cannot know which tab
     * was in front. Collapsing them is this repo's most repeated defect.
     */
    expect(filedWith(null, null, null).driverDeck, "a look that failed was erased").toBeNull();
    expect(filedWith(), "invented a look nobody took").not.toHaveProperty("driverDeck");

    /**
     * THE SAME REPAIR ON THE SIZE BESIDE IT. `if (driverSize)` dropped the field
     * on a failed read, so 3 of the last 120 archived rounds carry no
     * `driverSlideSize` and nothing says whether that is "could not tell" or
     * "never asked".
     *
     * It matters most here because the remaining reading is the untrustworthy
     * one: on rounds 115 and 116 the pane fell to its last rung and reported a
     * saved 720x540 file for a deck the driver had twice measured live at
     * 960x540. When the driver's number is missing, the only number left has
     * already lied once — and the reader has to be able to see that.
     */
    expect(filedWith("4:3").driverSlideSize).toBe("4:3");
    expect(filedWith(null).driverSlideSize, "an unread size was erased rather than recorded").toBeNull();
    expect(filedWith(), "invented a size reading for a round that never took one").not.toHaveProperty(
      "driverSlideSize",
    );

    /**
     * AND THE CALL SITE, which is the half that gets missed.
     *
     * Everything above proves `archive` records a deck it is GIVEN. None of it
     * proves anything gives it one — and an argument nobody passes is a field
     * that is always absent, which would look exactly like the archive we are
     * trying to fix. This repo has shipped that shape four times: the fix
     * written, one call site left behind, and an export nobody calls reading as
     * a comment that compiles.
     *
     * Asserted as source because the alternative is standing up the whole
     * download-and-sweep path for one argument.
     */
    const src = readFileSync(new URL("../scripts/round.mjs", import.meta.url), "utf8");
    expect(src, "collectRound stopped telling archive which deck was in front").toMatch(
      /archive\([\s\S]{0,400}archivableDeck\(frontedDeck\(sh\)\),/,
    );
    // AND IT GOES THROUGH THE FILTER, not straight from the tab list. See
    // `archivableDeck`: these files are committed to a public repository, and
    // an unconfigured round runs against whatever document happens to be in
    // front.
    expect(src, "the fronted name reaches a tracked file unfiltered").not.toMatch(
      /archive\([\s\S]{0,400}[^(]frontedDeck\(sh\) \?\? null,/,
    );
  });

  it("publishes only the deck names the operator himself configured", () => {
    /**
     * `rounds/*.json` and `crashes/*.json` ARE COMMITTED, and this repository is
     * public. `.gitignore` already draws that line and draws it about this
     * exact field — crash reports are excluded because they carry "request
     * URLs, session ids, the document name. This repo is public." The `.md`
     * dumps were held back for it; the `.json` files beside them were not,
     * because until 2026-09-02 they carried no name.
     *
     * Then `driverDeck` was added and wrote whatever was fronted into a tracked
     * file. On the test decks that is harmless. The hazard is the ordinary
     * path: with no `PW_DECK` the driver runs against whatever tab is in front,
     * so one round started beside a real presentation publishes its filename.
     */
    const saved = { deck: process.env.PW_DECK, wide: process.env.PW_DECK_16_9, tall: process.env.PW_DECK_4_3 };
    try {
      delete process.env.PW_DECK;
      delete process.env.PW_DECK_16_9;
      delete process.env.PW_DECK_4_3;

      // A real document nobody configured is never named.
      expect(archivableDeck("Q3 Board Review FINAL.pptx"), "published a document nobody asked for").toBe("undisclosed");
      // The cycle's own decks are named — somebody chose them on purpose, and
      // the whole point of the field is telling the two arms apart.
      expect(archivableDeck("Presentation70.pptx")).toBe("Presentation70.pptx");
      expect(archivableDeck("Presentation64.pptx")).toBe("Presentation64.pptx");

      // A deck named through the environment is the operator naming it.
      process.env.PW_DECK = "Presentation71";
      expect(archivableDeck("Presentation71.pptx"), "the deck this round was told to use was withheld").toBe(
        "Presentation71.pptx",
      );
      // Matched with and without the extension, because `PW_DECK` carries a
      // bare name and the tab list reports a filename.
      process.env.PW_DECK = "Weekly Numbers.pptx";
      expect(archivableDeck("Weekly Numbers.pptx")).toBe("Weekly Numbers.pptx");

      /**
       * THREE STATES, and null keeps its own meaning. Null is "the tab list
       * would not say"; `undisclosed` is "we looked and will not publish it".
       * Collapsing those is the same defect this driver spent the day removing
       * from slide counts.
       */
      expect(archivableDeck(null), "an unread tab list became a withheld name").toBeNull();
      expect(archivableDeck(""), "an empty name became a withheld name").toBeNull();
    } finally {
      if (saved.deck === undefined) delete process.env.PW_DECK;
      else process.env.PW_DECK = saved.deck;
      if (saved.wide === undefined) delete process.env.PW_DECK_16_9;
      else process.env.PW_DECK_16_9 = saved.wide;
      if (saved.tall === undefined) delete process.env.PW_DECK_4_3;
      else process.env.PW_DECK_4_3 = saved.tall;
    }
  });
});

describe("the round the pane finishes after the host dies", () => {
  /**
   * 55% OF EVERY CRASH ON FILE IS RECOVERABLE AND NOTHING WAS RECOVERING IT.
   *
   * 41 of the 75 sound-build crash records end at the same trace line —
   * `collecting deck evidence — scanning`, the POST-round inventory pass that
   * runs after every verdict is already banked. The pane bounds it at 45s so a
   * host dying in there cannot take the round; on timeout the scan gives up,
   * the round files without deck evidence, and `Download run log` is enabled.
   *
   * That timeout has fired in **0 of 77** records — because the driver returned
   * the instant it saw PowerPoint's dialog and was gone before the pane's 45
   * seconds were up.
   */
  const paneSh = (script: Array<string>) => {
    let n = 0;
    const calls: string[][] = [];
    const fn = ((...args: string[]) => {
      calls.push(args);
      if (args[0] !== "find") return "ok";
      const out = script[Math.min(n, script.length - 1)];
      n++;
      fn.state.lastFailed = out === "FAIL";
      return out === "FAIL" ? "" : out;
    }) as unknown as ((...a: string[]) => string) & { state: { lastFailed: boolean } };
    fn.state = { lastFailed: false };
    return { fn, calls, polls: () => n };
  };
  const ENABLED = '  - button "Download run log" [ref=f2]';
  const DISABLED = '  - button "Download run log" [disabled] [ref=f2]';
  const sleep = async () => {};

  it("waits out the pane's budget and files the round instead of losing it", async () => {
    // Disabled for three polls, then the pane's 45s scan budget expires, the
    // round files itself, and the button comes alive.
    const { fn, polls } = paneSh([DISABLED, DISABLED, DISABLED, ENABLED]);
    expect(await driver.paneFinishedAnyway(fn, sleep, Date.now(), 90_000)).toBe(true);
    expect(polls(), "it stopped polling before the pane could answer").toBeGreaterThanOrEqual(4);
  });

  it("gives up when the pane really is gone, rather than waiting for ever", async () => {
    // The other half. A host that took the pane with it must still cost a
    // bounded amount of time — this runs inside a 30-minute wedge budget and
    // spending it here would trade one lost round for a lost night.
    const { fn, polls } = paneSh([DISABLED]);
    const began = Date.now();
    expect(await driver.paneFinishedAnyway(fn, (async () => {}) as never, began, 300)).toBe(false);
    // BOUNDED, and measured rather than assumed: it has to come back inside the
    // grace it was given, not merely come back. An unbounded version of this
    // would trade one lost round for a whole lost night.
    expect(Date.now() - began, "it waited past the grace it was handed").toBeLessThan(5_000);
    expect(polls(), "it gave up without reading the button at all").toBeGreaterThanOrEqual(2);
  });

  it("never reads a DISABLED button as a finished round", async () => {
    /**
     * The button exists from the moment the pane loads; `disabled` is the only
     * thing separating "a round is running" from "a round is done". Matching
     * the name alone would file the PREVIOUS round's log as this one's — the
     * exact defect `archive`'s build check exists to catch downstream, and it
     * should never get that far.
     */
    const { fn, polls } = paneSh([DISABLED]);
    expect(await driver.paneFinishedAnyway(fn, (async () => {}) as never, Date.now(), 300)).toBe(false);
    /**
     * AND IT HAS TO HAVE LOOKED. The first version of this passed a 60ms grace,
     * which expired before the loop could poll — so it asserted "returns false"
     * against a function that had not read anything at all, and survived the
     * mutant that deletes the `[disabled]` lookahead entirely. A test that
     * cannot reach the code it names is not a test.
     */
    expect(polls(), "the grace expired before it read the button even once").toBeGreaterThanOrEqual(2);
  });

  it("does not read a poll that never ran as a pane that is gone", async () => {
    /**
     * `sh.state.lastFailed` means the CLI call itself did not happen — a second
     * terminal touching the session, a busy CLI — and it is not evidence about
     * the pane. The main poll loop carries a long comment about the round this
     * confusion once killed. A failed poll here must cost patience, not the
     * round: the pane answers on the very next one.
     */
    const { fn } = paneSh(["FAIL", "FAIL", ENABLED]);
    expect(await driver.paneFinishedAnyway(fn, sleep, Date.now(), 90_000)).toBe(true);
  });

  it("is wired into the crash branch, ahead of giving the round up", () => {
    /**
     * The half a unit test cannot reach. `paneFinishedAnyway` can be perfect
     * and change nothing if the crash branch still returns first — which is
     * precisely the shape of the defect it exists to fix, and the shape this
     * repo has shipped five times.
     *
     * Order matters twice over: the crash evidence must still be captured
     * BEFORE the wait (it is the only window in which the host's own account
     * exists), and the wait must come before the `return`.
     */
    const src = readFileSync(new URL("../scripts/round.mjs", import.meta.url), "utf8");
    // ANCHORED ON THE MESSAGE, not on `sawCrashDialog` — that call appears three
    // times in this file and `indexOf` finds the readiness one, 300 lines from
    // the poll loop this test is about. The first draft asserted against the
    // wrong branch entirely and failed for the right-looking wrong reason.
    const branch = src.slice(src.indexOf("PowerPoint crashed ${secs}s in"));
    const body = branch.slice(0, branch.indexOf("\n    }"));
    expect(body, "the crash branch no longer waits for the pane").toMatch(/paneFinishedAnyway\(sh, sleep, started\)/);
    expect(body.indexOf("keepCrashEvidence"), "the host's account is no longer captured first").toBeGreaterThan(-1);
    expect(
      body.indexOf("keepCrashEvidence") < body.indexOf("paneFinishedAnyway"),
      "the wait now runs before the crash evidence is captured — recovery would reload the tab first",
    ).toBe(true);
    expect(
      body.indexOf("paneFinishedAnyway") < body.indexOf('return { code: 1, reason: "crashed" }'),
      "the round is given up before the pane is given its chance",
    ).toBe(true);
  });
});

describe("the steps a crashed round managed to write", () => {
  const OFFERED = '        - button "Download the crashed run" [ref=f7c1]';
  const shWith = (out: string) => {
    const calls: string[][] = [];
    const fn = ((...args: string[]) => {
      calls.push(args);
      return args[0] === "find" ? out : "ok";
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state?: Record<string, unknown> };
    fn.dir = ".pw";
    fn.state = {};
    return { fn, calls };
  };
  const sleep = async () => {};

  it("keeps them, because a crashed round archives nothing at all", async () => {
    // ROUND 150 CRASHED SIX TIMES AND LEFT NO TRACE. Placing the crash needed
    // the scenario start times of OTHER rounds to guess at — and it did not
    // have to. `crashlog.ts` flushes every step to localStorage as it happens,
    // and the next pane load un-hides "Download the crashed run". The driver's
    // own recovery reopens that pane against the same persistent profile, so
    // the button was there after every one of those six crashes.
    const { fn, calls } = shWith(OFFERED);
    const copied: string[][] = [];
    const to = await driver.keepCrashedRun(
      fn,
      sleep,
      ((from: string, dest: string) => copied.push([from, dest])) as never,
      (() => true) as never,
    );
    expect(copied, "the crashed run was offered and not saved").toHaveLength(1);
    expect(copied[0][0]).toBe(".pw/.playwright-cli/ssf-charts-crashed-run.json");
    expect(to, "kept it but did not say where").toMatch(/^crashes\/.*-crashed-run\.json$/);
    // It has to actually press the button, not merely notice it.
    expect(
      calls.some((c) => c[0] === "eval"),
      "found the button and never clicked it",
    ).toBe(true);
  });

  it("names the tab in front, and does not pretend it is the deck that crashed", async () => {
    /**
     * 160 crash records name no document, and they are the half of the archive
     * that matters: 61 of them are the 4:3 arm. Without a name, "4:3 crashes far
     * more" and "that one file is sick" cannot be told apart.
     *
     * THE FIELD IS DELIBERATELY NOT CALLED `driverDeck`. This runs at the START
     * of the next round, rescuing the previous one's remains before they are
     * overwritten, so the tab in front is the NEXT round's starting state. It
     * usually is the crashed round's deck and after a `--fresh` restart or a leg
     * change it is not — and a query written against a field called `driverDeck`
     * would never know the difference.
     */
    const fn = ((...args: string[]) => {
      if (args[0] === "find") return OFFERED;
      if (args[0] === "tab-list") return "0: https://x/ [Presentation70.pptx] (current)";
      return "ok";
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state?: Record<string, unknown> };
    fn.dir = ".pw";
    fn.state = {};

    let written = "";
    const copied: string[][] = [];
    await driver.keepCrashedRun(
      fn,
      sleep,
      ((from: string, dest: string) => copied.push([from, dest])) as never,
      (() => true) as never,
      (() => JSON.stringify({ kind: "ssf-charts-crash-log", steps: ["  1s  x"] })) as never,
      ((_p: string, body: string) => {
        written = body;
      }) as never,
    );
    const kept = JSON.parse(written) as Record<string, unknown>;
    expect(kept.frontedWhenRescued, "the rescued crash record still names no document").toBe("Presentation70.pptx");
    expect(kept, "renaming this to driverDeck would make every future query wrong").not.toHaveProperty("driverDeck");
    // The evidence itself must survive the stamping untouched.
    expect(kept.steps, "stamping the record cost it its steps").toEqual(["  1s  x"]);
    expect(copied, "stamped it AND copied over it").toHaveLength(0);
    // Two-space JSON with a trailing newline: `crashes/` is not in
    // `.prettierignore`, so the gate reads what this writes.
    expect(written.endsWith("\n"), "no trailing newline — prettier fails the directory").toBe(true);
    expect(written).toContain('\n  "kind"');
  });

  it("keeps the bytes exactly as they came when the record will not parse", async () => {
    // A CRASH RECORD IS THE ONLY COPY OF A ROUND THAT DIED. Losing one to a
    // stamping error would be the worst trade in this file, so anything
    // unexpected falls back to the byte copy that has always worked.
    const { fn } = shWith(OFFERED);
    const copied: string[][] = [];
    const to = await driver.keepCrashedRun(
      fn,
      sleep,
      ((from: string, dest: string) => copied.push([from, dest])) as never,
      (() => true) as never,
      (() => "{ this is not json") as never,
      (() => {
        throw new Error("must not write a half-stamped record");
      }) as never,
    );
    expect(copied, "a record that would not parse was dropped instead of copied").toHaveLength(1);
    expect(to).toMatch(/^crashes\/.*-crashed-run\.json$/);
  });

  it("does nothing when no crashed run is offered", async () => {
    // The ordinary case, every healthy round. `clearCrashLog` hides the button
    // once pressed, so this must not re-download in a loop either.
    const { fn } = shWith('        - button "Insert chart" [ref=f1]');
    const copied: string[][] = [];
    const to = await driver.keepCrashedRun(
      fn,
      sleep,
      ((a: string, b: string) => copied.push([a, b])) as never,
      (() => true) as never,
    );
    expect(to).toBeNull();
    expect(copied, "invented a crashed run out of a healthy pane").toHaveLength(0);
  });

  it("says so when the file never arrives, rather than reporting a save", async () => {
    // An ABSENT file must not read as a rescue. The download is a browser
    // save that can simply not happen, and `collectRound` already carries a
    // comment about exactly that on the run log.
    const { fn } = shWith(OFFERED);
    const copied: string[][] = [];
    const to = await driver.keepCrashedRun(
      fn,
      sleep,
      ((a: string, b: string) => copied.push([a, b])) as never,
      (() => false) as never,
    );
    expect(to, "claimed a save for a file that never arrived").toBeNull();
    expect(copied).toHaveLength(0);
  });
});

describe("what a retry recovered from", () => {
  it("names the stops, rather than bucketing them as not-ready", () => {
    // FOUR ROUNDS RUNNING RECORDED {"attempts":2,"recovered":["not-ready"]}.
    // That reads as "the host needed rescuing" and mostly does not mean it:
    // round 153's first attempt refused for `pane-stale` — a build had just
    // been deployed and the pane held the previous one, which is a property of
    // how rounds are RUN — and for `host-silent`, which is host health.
    //
    // Folded into one word they cannot be told apart, so any claim about the
    // host getting better or worse would be drawn from a bucket. This is the
    // residual-bucket mistake made inside the field added to stop it.
    expect(driver.recoveryLabel("not-ready", ["host-silent", "pane-stale"])).toBe("not-ready:host-silent+pane-stale");
  });

  it("sorts them, so the same pair counts as the same thing every round", () => {
    // Unsorted, `host-silent+pane-stale` and `pane-stale+host-silent` are two
    // different strings describing one situation, and pooling them across the
    // archive would report two causes where there is one.
    expect(driver.recoveryLabel("not-ready", ["pane-stale", "host-silent"])).toBe(
      driver.recoveryLabel("not-ready", ["host-silent", "pane-stale"]),
    );
  });

  it("falls back to the bare reason when there are no codes", () => {
    // `crashed` and `browser-gone` carry no codes — they are not readiness
    // stops. A trailing colon would make them look truncated.
    expect(driver.recoveryLabel("crashed", [])).toBe("crashed");
    expect(driver.recoveryLabel("browser-gone", undefined)).toBe("browser-gone");
  });
});

describe("starting a round from a fresh session", () => {
  it("closes the browser BEFORE recovering, not after", () => {
    // ORDER IS THE WHOLE THING. `recover` only opens a browser when there is
    // not one — `if (noBrowser(...))` — so recovering into a live browser
    // reloads the stale session instead of replacing it, which is exactly what
    // this exists to avoid. Recovering first and closing after would leave no
    // browser at all.
    const order: string[] = [];
    const sh = ((...args: string[]) => {
      order.push(args[0]);
      return "";
    }) as never;
    return driver
      .startFresh(
        sh,
        (async () => {}) as never,
        (async () => {
          order.push("recover");
        }) as never,
        // The no-op prune is REQUIRED, not tidiness. The real default deletes
        // real directories, and this test walked C:/devtools/pw-profile and
        // pruned 1.5GB of live caches before the seam was injectable.
        "C:/unused",
        (() => null) as never,
      )
      .then(() => {
        expect(order, "the browser was not closed before the rebuild").toEqual(["close", "recover"]);
      });
  });

  it("waits after closing, because the process does not go away instantly", async () => {
    // `close` returns before the browser is gone, and `recover` decides whether
    // to open one by asking whether a browser is there. Racing them reads the
    // dying session as a live one and reloads it.
    let waited = 0;
    await driver.startFresh(
      (() => "") as never,
      (async (ms: number) => {
        waited += ms;
      }) as never,
      (async () => {}) as never,
      "C:/unused",
      (() => null) as never,
    );
    expect(waited, "closed and rebuilt in the same breath").toBeGreaterThan(0);
  });
});

describe("recovering when the browser lives but the deck does not", () => {
  /** Records every CLI verb, and answers whatever the caller says. */
  const shFor = (answers: Record<string, string>) => {
    const calls: string[][] = [];
    const fn = ((...args: string[]) => {
      calls.push(args);
      return answers[args[0]] ?? "";
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state?: Record<string, unknown> };
    fn.dir = ".pw";
    fn.state = {};
    return { fn, calls };
  };
  const sleep = (async () => {}) as never;

  it("opens the deck when a browser is up but no tab holds it", async () => {
    // THE BUG THIS FIXES COST SEVEN ATTEMPTS. The deck-opening block sat inside
    // `if (noBrowser(...))`, which assumed a living browser implies an open
    // deck. It does not — after a tab crash, or anything that reopens the
    // browser on OneDrive's home, `recover` skipped the deck entirely, reloaded
    // the wrong page and hunted for a pane that had never been opened. Every
    // attempt reported "could not read the pane's build stamp" with
    // `deck ? slide(s)` beside it, while the deck sat in the file list.
    const { fn, calls } = shFor({
      list: "- default:\n  - status: open",
      "tab-list": "- 0: (current) [Home - OneDrive](https://onedrive.live.com/)",
    });
    await driver.recover(fn, sleep);
    // It has to go LOOKING for the deck. Without this the whole recovery is a
    // reload of whatever page happened to be in front.
    expect(
      calls.some((c) => c[0] === "find" && c.some((a) => a.includes(driver.DECK_NAME))),
      "a browser was up, the deck was not, and recover never went looking for it",
    ).toBe(true);
  });

  it("does not open a second tab on a deck that is already open", async () => {
    // A deck tab that is present needs SELECTING, not opening again. Clicking
    // the file a second time is how a recovery leaves two tabs on one document,
    // and the driver then measures whichever one it happens to front.
    const { fn, calls } = shFor({
      list: "- default:\n  - status: open",
      "tab-list": `- 0: (current) [${driver.DECK_NAME}.pptx](https://x)`,
    });
    await driver.recover(fn, sleep);
    expect(
      calls.some((c) => c[0] === "find" && c.some((a) => a.includes(driver.DECK_NAME))),
      "went hunting for a deck that was already open",
    ).toBe(false);
  });
});

describe("selecting a deck tab that is already open", () => {
  const shFor = (answers: Record<string, string>) => {
    const calls: string[][] = [];
    const fn = ((...args: string[]) => {
      calls.push(args);
      return answers[args[0]] ?? "";
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state?: Record<string, unknown> };
    fn.dir = ".pw";
    fn.state = {};
    return { fn, calls };
  };

  it("fronts a deck tab it did not have to open", async () => {
    // MY OWN REGRESSION, one commit old. Guarding the whole block on "the deck
    // tab is absent" meant a deck that WAS open but not fronted never got
    // selected, so the reload below refreshed whatever tab happened to be
    // current — OneDrive's home — and every attempt reported `deck ? slide(s)`
    // with the deck sitting one tab away.
    //
    // Opening and selecting are two jobs. Click the file only when no tab holds
    // it; select it either way.
    const { fn, calls } = shFor({
      list: "- default:\n  - status: open",
      "tab-list": `- 0: (current) [Home - OneDrive](https://onedrive.live.com/)\n- 1: [${driver.DECK_NAME}.pptx](https://x)`,
    });
    await driver.recover(fn, (async () => {}) as never);
    expect(
      calls.some((c) => c[0] === "tab-select" && c[1] === "1"),
      "the deck was open on tab 1 and never fronted",
    ).toBe(true);
    // …and it must not click the file again, which is how one document ends up
    // with two tabs and the driver measures whichever it fronts.
    expect(
      calls.some((c) => c[0] === "find" && c.some((a) => a.includes(driver.DECK_NAME))),
      "re-opened a deck that was already open",
    ).toBe(false);
  });
});

describe("a sign-in popup that outlived its flow", () => {
  const TABS =
    "### Result\n" +
    "- 0: [Home - OneDrive](https://onedrive.live.com/)\n" +
    "- 1: (current) [Presentation64.pptx](https://onedrive.live.com/personal/x/doc.aspx)\n" +
    "- 2: [Microsoft account](https://login.live.com/oauth20_authorize.srf?client_id=abc)";
  const shFor = (tabs: string) => {
    const calls: string[][] = [];
    const fn = ((...args: string[]) => {
      calls.push(args);
      return args[0] === "tab-list" ? tabs : "";
    }) as never;
    return { fn, calls };
  };

  it("closes it once the host has answered, which is the proof it is superfluous", () => {
    // MSAL opens the sign-in in a popup and closes it when the flow finishes.
    // On 2026-08-22 one did not: it sat as a third tab for three hours while
    // the deck worked normally beside it. Office.js answering is what proves
    // nothing is waiting on that window.
    const { fn, calls } = shFor(TABS);
    expect(driver.closeStaleAuthPopups(fn, { hostAnswered: true })).toEqual([2]);
    expect(calls).toContainEqual(["tab-close", "2"]);
  });

  it("closes NOTHING when the host has not answered", () => {
    // THE GATE THAT MATTERS. Without a live host there is no evidence the flow
    // is finished, and a sign-in window that might still be live is the one
    // surface this loop must not touch.
    const { fn, calls } = shFor(TABS);
    expect(driver.closeStaleAuthPopups(fn, { hostAnswered: false })).toEqual([]);
    expect(
      calls.some((c) => c[0] === "tab-close"),
      "closed a sign-in window with no proof it was done",
    ).toBe(false);
  });

  it("never closes the tab a person is looking at", () => {
    // A FRONTED sign-in window may be one someone is part-way through, and the
    // driver has no way to tell a finished flow from a half-typed one.
    const fronted =
      "- 0: [Home - OneDrive](https://onedrive.live.com/)\n" +
      "- 1: (current) [Microsoft account](https://login.live.com/oauth20_authorize.srf?x=1)";
    const { fn } = shFor(fronted);
    expect(driver.closeStaleAuthPopups(fn, { hostAnswered: true })).toEqual([]);
  });

  it("leaves every other tab alone", () => {
    // Only the two sign-in hosts, matched on the URL. A deck tab closed by
    // mistake would end the round it belongs to.
    const { fn } = shFor(
      "- 0: [Home - OneDrive](https://onedrive.live.com/)\n- 1: (current) [x](https://example.com/)",
    );
    expect(driver.closeStaleAuthPopups(fn, { hostAnswered: true })).toEqual([]);
  });

  it("closes the highest index first, because closing renumbers the rest", () => {
    const many =
      "- 0: (current) [deck](https://onedrive.live.com/x)\n" +
      "- 1: [Microsoft account](https://login.live.com/a)\n" +
      "- 2: [Microsoft account](https://login.microsoftonline.com/b)";
    const { fn } = shFor(many);
    expect(driver.closeStaleAuthPopups(fn, { hostAnswered: true })).toEqual([2, 1]);
  });
});

describe("a read that could not run", () => {
  it("is not reported as an absence", () => {
    // THE HOUSE DEFECT IN ITS PUREST FORM. `spawnSync` on this machine
    // intermittently answers ENOENT for a node.exe that is plainly there —
    // eight consecutive calls succeeded minutes later — and a failed spawn
    // returns "". Every reading is taken from that empty string.
    //
    // On 2026-08-22 one such failure was reported as `pane ?`, `deck ?
    // slide(s)`, "could not read the pane's build stamp" and finally "the
    // add-in is gone from this document". The deck was open, the stamp was
    // findable and `Insert chart` was in the ribbon throughout.
    const out = readiness({
      ...READY,
      stamp: null,
      slides: null,
      canOpenPane: false,
      commandPresent: false,
      readsFailed: true,
    });
    expect(out.ok).toBe(false);
    expect(out.codes, "a call that never ran was reported as a missing add-in").not.toContain("addin-missing");
    expect(out.codes, "a call that never ran was reported as a closed pane").not.toContain("pane-closed");
    expect(out.codes).toContain("reads-failed");
    // RECOVERABLE: the next attempt's calls usually run, so trying again is
    // exactly the right response — unlike `addin-missing`, which only a person
    // can clear.
    expect(RECOVERABLE_STOPS.has("reads-failed"), "a transient spawn failure must not end the night").toBe(true);
  });

  it("does not fire when the reads ran and simply found nothing", () => {
    // THE REAL REFUSALS MUST SURVIVE. An add-in that is genuinely absent, read
    // by calls that genuinely ran, still has to stop the round.
    const out = readiness({
      ...READY,
      stamp: null,
      slides: 1,
      canOpenPane: false,
      commandPresent: false,
      ribbonRoom: 3800,
      readsFailed: false,
    });
    expect(out.codes, "softened away the stop that exists to catch a real problem").toContain("addin-missing");
  });
});

describe("pruning the profile caches", () => {
  it("never names a directory that holds the sign-in or the add-in", () => {
    // THE ONE MISTAKE THAT MATTERS. `Network/` is Chrome's cookie store, which
    // is the sign-in — and only the owner can replace that. `Local Storage`
    // holds crashlog.ts's flushed steps AND the sideloaded manifest itself:
    // Microsoft's docs say "the add-in's manifest is stored in the browser's
    // local storage, so if you clear the browser's cache... you have to
    // sideload the add-in again".
    //
    // Asserted against the LIST rather than the behaviour, because the list is
    // the whole safety argument: a deny-list is one rename away from deleting
    // something new, an allow-list is not.
    const forbidden = ["Network", "Local Storage", "Cookies", "Preferences", "IndexedDB", "Sessions"];
    for (const dir of driver.PROFILE_CACHE_DIRS)
      for (const f of forbidden)
        expect(dir.includes(f), `${dir} would delete ${f}, which holds the sign-in or the add-in`).toBe(false);
  });

  it("does nothing at all while a browser is open", () => {
    // Chrome rewrites these files as it runs, and deleting underneath it
    // corrupts a profile whose sign-in only the owner can restore.
    const removed: string[] = [];
    expect(
      driver.pruneProfileCaches("C:/p", {
        browserOpen: true,
        rm: ((p: string) => removed.push(p)) as never,
        exists: (() => true) as never,
        sizeOf: () => 9e9,
      }),
    ).toBeNull();
    expect(removed, "deleted cache files with the browser still holding them").toHaveLength(0);
  });

  it("leaves a small profile alone, because the caches cost time to rebuild", () => {
    // PowerPoint re-downloads and re-JITs its bundles after a prune, so this is
    // a once-in-a-while reclaim and not a per-round tidy.
    const removed: string[] = [];
    expect(
      driver.pruneProfileCaches("C:/p", {
        browserOpen: false,
        rm: ((p: string) => removed.push(p)) as never,
        exists: (() => true) as never,
        sizeOf: () => 1000,
      }),
    ).toBeNull();
    expect(removed).toHaveLength(0);
  });

  it("prunes when the profile is genuinely large", () => {
    const removed: string[] = [];
    const out = driver.pruneProfileCaches("C:/p", {
      browserOpen: false,
      rm: ((p: string) => removed.push(p)) as never,
      exists: (() => true) as never,
      sizeOf: () => 9e9,
    });
    expect(out?.removed.length, "found a 9GB profile and pruned nothing").toBeGreaterThan(0);
    expect(
      removed.every((p) => p.includes("/Default/")),
      "deleted something outside the profile",
    ).toBe(true);
  });
});

describe("which stop wins when several are true", () => {
  it("names the missing browser rather than the reads it made fail", () => {
    // ORDERING IS A DIAGNOSIS. With no browser EVERY call exits non-zero —
    // "The browser 'default' is not open" — so `readsFailed` is true too. The
    // first version of that guard sat ABOVE this one and swallowed it,
    // reporting "a read could not be RUN" for a condition that has a name, a
    // cause and a recovery.
    //
    // Putting the vague stop ahead of the specific one is the same defect as
    // reporting an absence for an unmeasured thing, one layer up. It went into
    // this file and out again within the hour.
    const out = readiness({ ...READY, browserGone: true, readsFailed: true, stamp: null, slides: null });
    expect(out.codes, "the vaguer stop won").toContain("browser-gone");
    expect(out.codes, "reported a failed read for a browser that is simply not there").not.toContain("reads-failed");
  });

  it("still reports a failed read when the browser IS there", () => {
    // The stop has to survive the reordering, or the fix it exists for is gone.
    const out = readiness({ ...READY, browserGone: false, readsFailed: true, stamp: null, slides: null });
    expect(out.codes).toContain("reads-failed");
  });
});

describe("a browser this session cannot reach", () => {
  /** `list` sees it; every session-keyed call exits non-zero. */
  const unreachable = () => {
    const calls: string[][] = [];
    const fn = ((...args: string[]) => {
      calls.push(args);
      if (args[0] === "list") {
        fn.state.lastFailed = false;
        return "### Browsers\n- default:\n  - status: open";
      }
      fn.state.lastFailed = true;
      return "";
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state: Record<string, unknown> };
    fn.state = { lastFailed: false, lastError: null };
    fn.dir = ".pw";
    return { fn, calls };
  };

  it("is not reachable, however healthy `list` looks", () => {
    // `list` enumerates every browser the daemon knows, whatever session it
    // belongs to; every other command is session-keyed, and the daemon keys
    // sessions by the working-directory STRING. A browser opened from a shell
    // whose cwd was `C:/devtools/...` is invisible to a driver whose cwd is
    // `C:\devtools\...` — same folder, different string.
    const { fn } = unreachable();
    expect(driver.browserReachable(fn), "called a browser reachable that answers nothing").toBe(false);
  });

  it("trusts the failure flag over the output, if the two ever disagree", () => {
    // TODAY `cli` returns "" whenever the call exits non-zero, so the emptiness
    // check alone would catch this and the flag is belt-and-braces. Pinned
    // anyway: the day `cli` returns partial output on a failed call — a
    // truncated read, a half-written answer — the flag is the only thing that
    // still tells the truth, and a reachability check that believed the bytes
    // would call a broken session healthy.
    const fn = ((..._args: string[]) => "- 0: [something](https://x)") as unknown as ((...a: string[]) => string) & {
      state: Record<string, unknown>;
    };
    fn.state = { lastFailed: true, lastError: null };
    expect(driver.browserReachable(fn), "believed the output over the failure flag").toBe(false);
  });

  it("gets closed so one can be opened here", async () => {
    // A browser we cannot use STILL HOLDS THE PROFILE, and `open` refuses with
    // "Browser is already in use". `--isolated` is no use: a fresh profile has
    // no sign-in, and the sign-in is the one thing this loop cannot recreate.
    const { fn, calls } = unreachable();
    await driver.recover(fn, (async () => {}) as never);
    const closed = calls.findIndex((c) => c[0] === "close-all");
    const opened = calls.findIndex((c) => c[0] === "open");
    expect(closed, "left an unusable browser holding the profile").toBeGreaterThanOrEqual(0);
    expect(opened, "never opened a browser this session can use").toBeGreaterThanOrEqual(0);
    expect(closed, "opened before closing — the profile would still be locked").toBeLessThan(opened);
  });

  it("does not close anything when there is simply no browser", async () => {
    // The ordinary path must not pay for a close it does not need.
    const calls: string[][] = [];
    const fn = ((...args: string[]) => {
      calls.push(args);
      return args[0] === "list" ? "  (no browsers)" : "";
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state: Record<string, unknown> };
    fn.state = { lastFailed: false, lastError: null };
    fn.dir = ".pw";
    await driver.recover(fn, (async () => {}) as never);
    expect(
      calls.some((c) => c[0] === "close-all"),
      "closed browsers when there were none",
    ).toBe(false);
    expect(
      calls.some((c) => c[0] === "open"),
      "never opened one",
    ).toBe(true);
  });

  it("asks for a window at least as wide as a ribbon command needs", async () => {
    // THE WIDEN COULD NOT REACH THE BAR IT EXISTS TO CLEAR. `ensureRibbonRoom`
    // requested 1900px while `MIN_RIBBON_WIDTH` was 2375. 1900 was generous when
    // the threshold was 1600; raising the threshold left the request behind, and
    // the two are constants that have to move together.
    //
    // It hid behind the device scale factor. On this machine a request of 1900
    // usually ARRIVES as 3800 CSS px, so every round printed "widened to 3800px"
    // and nothing looked wrong. Round 181 reopened the browser with `--fresh`,
    // the scale came back 1:1, 1900 arrived as 1900, the Add-ins command stayed
    // in the overflow, and SEVEN consecutive attempts could not reopen the pane.
    // A whole round lost to a number that had been too small the entire time.
    //
    // This is the guard that would have caught it the moment the threshold moved.
    const calls: string[][] = [];
    let width = 1009;
    const sh = ((...args: string[]) => {
      calls.push(args);
      if (args[0] === "resize") {
        // The honest 1:1 case — the arrival equals the request. Anything that
        // relies on a scale factor to clear the bar is relying on luck.
        width = Number(args[1]);
        return "";
      }
      return `{"width":${width}}`;
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state: Record<string, unknown> };
    sh.state = { lastFailed: false, lastError: null };

    const after = await driver.ensureRibbonRoom(sh, async () => {});
    const resize = calls.find((c) => c[0] === "resize");
    expect(resize, "never asked for a wider window at all").toBeDefined();
    expect(
      Number(resize![1]),
      `asked for ${resize?.[1]}px, under the ${driver.MIN_RIBBON_WIDTH}px a ribbon command needs`,
    ).toBeGreaterThanOrEqual(driver.MIN_RIBBON_WIDTH);
    expect(after, "the widen returned a width still under the threshold").toBeGreaterThanOrEqual(
      driver.MIN_RIBBON_WIDTH,
    );
  });

  it("leaves a window that is already wide enough alone", async () => {
    // The early return matters: a resize costs three seconds of sleep on every
    // attempt, and a round that retries seven times would pay it seven times.
    const calls: string[][] = [];
    const sh = ((...args: string[]) => {
      calls.push(args);
      return `{"width":${driver.MIN_RIBBON_WIDTH + 100}}`;
    }) as unknown as ((...a: string[]) => string) & { dir?: string; state: Record<string, unknown> };
    sh.state = { lastFailed: false, lastError: null };
    await driver.ensureRibbonRoom(sh, async () => {});
    expect(
      calls.some((c) => c[0] === "resize"),
      "resized a window that was already wide enough",
    ).toBe(false);
  });
});

describe("where in a run of rounds this one sits", () => {
  const T0 = Date.parse("2026-08-24T08:00:00.000Z");
  it("counts up while rounds keep coming", () => {
    // Consecutive rounds land 12-20 minutes apart; 15 is an ordinary gap.
    const p = sessionPosition(
      { at: "2026-08-24T08:00:00.000Z", sessionIndex: 4, sessionStartedAt: "2026-08-24T07:00:00.000Z" },
      T0 + 15 * 60000,
    );
    expect(p.index).toBe(5);
    expect(p.sincePrevMs).toBe(15 * 60000);
    // MINUTES INTO THE SESSION is what the drift is indexed by — rounds run
    // 684s to 1085s, so the count is a proxy and the clock is the variable.
    expect(p.startedAt).toBe("2026-08-24T07:00:00.000Z");
    expect(p.elapsedMs, "07:00 to 08:15 is 75 minutes").toBe(75 * 60000);
  });

  it("starts over after a long enough gap", () => {
    // Came back after lunch. This is the only distinction the gap has to make.
    const p = sessionPosition({ at: "2026-08-24T08:00:00.000Z", sessionIndex: 9 }, T0 + 90 * 60000);
    expect(p.index, "a new session, not the tenth round of the old one").toBe(1);
    expect(p.sincePrevMs).toBe(null);
    // A new session starts its own clock at zero rather than inheriting one.
    expect(p.elapsedMs).toBe(0);
    expect(p.startedAt).toBe(new Date(T0 + 90 * 60000).toISOString());
  });

  it("keeps counting across the receipt the cycle clears before every leg", () => {
    /**
     * THE DEFECT THIS SPLIT EXISTS FOR, as an arithmetic statement.
     *
     * `cycle.mjs` deletes `RECEIPT_PATH` before each leg so its "the driver left
     * no receipt" check reads THIS leg's outcome and not the last one's. The
     * session chain used to live in that same file, so every leg of every cycle
     * started over: of 52 archived rounds under 45 minutes apart, 33 record the
     * later one as round 1, and rounds 334 and 335 — nine minutes apart, inside
     * one cycle — both say `sessionIndex: 1, 0m in`.
     *
     * The chain is `SESSION_PATH` now. Reading a session while the receipt is
     * gone has to keep counting, because that is precisely the state a cycle
     * leaves between its legs.
     */
    const p = sessionPosition(
      { at: "2026-08-24T08:00:00.000Z", sessionIndex: 2, sessionStartedAt: "2026-08-24T07:45:00.000Z" },
      T0 + 9 * 60000,
    );
    expect(p.index, "a cleared receipt must not reset the session").toBe(3);
    expect(p.elapsedMs, "07:45 to 08:09 is 24 minutes").toBe(24 * 60000);
  });

  it("reaches the depth at which the driver refuses to run", () => {
    // `sessionDepthWarning` refuses from SESSION_DEPTH_WARN_AT, and a refusal
    // needs an index that climbs. Pinned at 1 — which is what a cleared receipt
    // produced on every leg — the stop was unreachable and `--deep` had nothing
    // to override.
    const fifth = sessionPosition(
      { at: "2026-08-24T08:00:00.000Z", sessionIndex: 4, sessionStartedAt: "2026-08-24T07:00:00.000Z" },
      T0 + 15 * 60000,
    );
    expect(fifth.index).toBe(SESSION_DEPTH_WARN_AT);
    expect(sessionDepthWarning(fifth.index), "the depth stop never fires").not.toEqual([]);
    // …and a first round still runs, or the driver would refuse everything.
    expect(sessionDepthWarning(sessionPosition(null, T0).index)).toEqual([]);
  });

  it("keeps the chain in a file of its own, not the one the cycle clears", () => {
    // The two paths must not be the same string, and `readSession` must be
    // reading the session file rather than quietly falling back to the receipt
    // — which is how this would regress to exactly what it was.
    expect(SESSION_PATH).not.toBe(RECEIPT_PATH);
    const seen: string[] = [];
    const chained = readSession(SESSION_PATH, (p: string) => {
      seen.push(String(p));
      return JSON.stringify({ at: "2026-08-24T08:00:00.000Z", sessionIndex: 6, sessionStartedAt: "x" });
    });
    expect(seen, "read something other than the session file").toEqual([SESSION_PATH]);
    expect(chained?.sessionIndex).toBe(6);
    // A missing chain is "no previous round", never a throw: a first-ever run
    // has no file and must still be allowed to start.
    expect(
      readSession(SESSION_PATH, () => {
        throw new Error("ENOENT");
      }),
    ).toBeNull();
  });

  it("wires the session chain, not the receipt, into the position", () => {
    /**
     * READ FROM THE SOURCE, because `main` is not exported and the wiring is
     * the whole defect.
     *
     * Everything above tests `sessionPosition`, which was never wrong — it was
     * being handed the wrong file. Swapping `readSession` back to `readReceipt`
     * in `main` reproduces the original defect exactly and passed every other
     * test in this suite, so the one line that matters had no guard at all.
     *
     * The same idiom `host-probe.test.ts` uses for "the formula was inlined
     * again instead of calling the helper": crude, and it fails on the change
     * that would undo this.
     */
    const src = readFileSync(new URL("../scripts/round.mjs", import.meta.url), "utf8");
    expect(src, "the session chain is not read from its own file").toMatch(
      /const prevSession = \(deps\.readSession \?\? readSession\)\(\)/,
    );
    expect(src, "the position is taken from something other than the session chain").toMatch(
      /sessionPosition\(prevSession, startedAtMs\)/,
    );
    // And `main` no longer reads the receipt AT ALL — lint found that the
    // moment the chain moved, because the chain was its only reader there. The
    // receipt is written here and read by `cycle.mjs`: one question each, which
    // is the whole point of the split. Asserted so nothing quietly hands the
    // receipt a second reader with a second lifetime again.
    expect(src, "main is reading the receipt again — give it its own question or none").not.toMatch(
      /const prevReceipt = /,
    );
    expect(src, "the receipt stopped being written, so the cycle cannot see this leg").toMatch(/write\(RECEIPT_PATH/);
  });

  it("treats no previous round as the first", () => {
    expect(sessionPosition(null, T0).index).toBe(1);
    expect(sessionPosition({}, T0).index).toBe(1);
    expect(sessionPosition({ at: "not a date" }, T0).index).toBe(1);
  });

  it("refuses a receipt from the future rather than trusting a negative gap", () => {
    // A clock that moved is not a session. Round 2 with sincePrev of -40m would
    // be worse than no reading at all, because it looks like data.
    const p = sessionPosition({ at: "2026-08-24T09:00:00.000Z", sessionIndex: 3 }, T0);
    expect(p.index).toBe(1);
    expect(p.sincePrevMs).toBe(null);
  });

  it("never reports a negative time-into-session", () => {
    // Reachable: `at` is in the past and inside the gap, so the session
    // continues, while `sessionStartedAt` sits ahead of now because the clock
    // moved between the two writes. A negative "minutes in" is worse than no
    // reading, because it looks like data.
    const p = sessionPosition(
      { at: "2026-08-24T07:50:00.000Z", sessionIndex: 2, sessionStartedAt: "2026-08-24T09:00:00.000Z" },
      T0,
    );
    expect(p.elapsedMs).toBe(0);
  });

  it("recovers when the previous receipt has no index yet", () => {
    // Every receipt written before this field existed. The round after one of
    // those is the SECOND of its session, not the first.
    const p = sessionPosition({ at: "2026-08-24T08:00:00.000Z" }, T0 + 12 * 60000);
    expect(p.index).toBe(2);
  });
});

describe("reading the last round's receipt", () => {
  it("treats every unreadable receipt as no previous round", () => {
    // A missing file, a half-written one and a corrupt one must not stop a
    // round from running.
    expect(
      readReceipt("x", () => {
        throw new Error("ENOENT");
      }),
    ).toBe(null);
    expect(readReceipt("x", () => "{not json")).toBe(null);
    expect(readReceipt("x", () => JSON.stringify({ sessionIndex: 3 }))).toEqual({ sessionIndex: 3 });
  });
});

describe("a round deep in a session is refused unless asked for", () => {
  it("says nothing for the first four rounds", () => {
    // Ten back-to-back rounds skipped [0 0 0 0 1 2 2 2 0 3]. The first FOUR
    // skipped nothing, so warning earlier would cry wolf on rounds that are fine.
    expect(sessionDepthWarning(1)).toEqual([]);
    expect(sessionDepthWarning(4)).toEqual([]);
  });

  it("warns from the fifth, where the skips actually start", () => {
    const w = sessionDepthWarning(5);
    expect(w).toHaveLength(1);
    expect(w[0]).toContain("SKIPPED");
    expect(w[0], "the reader needs the remedy, not just the diagnosis").toContain("Rest 45+ minutes");
  });

  const deepRound = (extra = {}) =>
    readiness({
      head: "abc1234",
      deployed: "abc1234",
      stamp: "abc1234",
      slides: 1,
      verbose: true,
      pictures: true,
      sessionIndex: 9,
      ...extra,
    });

  it("refuses the round, having been ignored as a warning", () => {
    // THIS REVERSES AN EARLIER DECISION, on evidence the earlier decision did
    // not have. It read: "A deep round still produces a real sheet; refusing one
    // throws away work somebody asked for. What it must not do is pass for a
    // complete round."
    //
    // The second sentence is the one that failed. Round 244 was the fifth
    // back-to-back of a block; the warning printed in full, was read, and the
    // round was run and filed anyway — and its self-test half is worth less than
    // the three before it for precisely the reason the line predicted. A warning
    // that is printed and stepped over is not a label, it is decoration.
    //
    // The first sentence keeps its force, which is why this is a stop with a
    // door rather than a ban: `--deep` runs it and records the arm. No work is
    // thrown away, it just has to be asked for.
    const r = deepRound();
    expect(r.ok).toBe(false);
    expect(r.codes).toContain("deep-session");
    expect(r.stop[0], "the reader needs the way out, not just the refusal").toContain("--deep");
  });

  it("runs it when asked, and then labels it", () => {
    const r = deepRound({ allowDeepSession: true });
    expect(r.ok).toBe(true);
    // Still warned. Consenting to a shallower round does not make it a complete
    // one, and the round file is read by someone who was not there.
    expect(r.warn).toHaveLength(1);
  });

  it("does not say it twice", () => {
    // Once it is a stop it is not also a warning: one problem reported in two
    // registers reads as two problems, and a reader counting them is misled
    // about how much is wrong with the round.
    expect(deepRound().warn).toEqual([]);
  });

  it("cannot be recovered from, because nothing but waiting fixes it", () => {
    // Every other readiness stop names something the driver can act on — reopen
    // the pane, reload the tab, clean the deck. This one names the clock. A
    // retry would burn three more attempts and refuse three more times.
    expect(RECOVERABLE_STOPS.has("deep-session")).toBe(false);
    expect(shouldRetry("not-ready", 1, 4, ["deep-session"])).toBe(false);
  });

  it("reports what else is wrong with a deep round, not just its depth", () => {
    // Refusing on depth FIRST would hide a stale pane behind "you are tired",
    // and the owner would rest for an hour and then meet the pane.
    const r = deepRound({ stamp: "older12" });
    expect(r.codes).toContain("pane-stale");
    expect(r.codes).toContain("deep-session");
  });

  it("invents no warning from a missing or odd index", () => {
    expect(sessionDepthWarning(undefined)).toEqual([]);
    expect(sessionDepthWarning(null)).toEqual([]);
    expect(sessionDepthWarning("lots")).toEqual([]);
  });
});

describe("an attempt that ran no round keeps out of the session", () => {
  const prev = {
    at: "2026-08-25T18:00:00.000Z",
    sessionIndex: 4,
    sessionStartedAt: "2026-08-25T17:00:00.000Z",
  };
  const receipt = (over: Record<string, unknown>) =>
    outcomeReceipt({
      codes: [],
      build: null,
      size: null,
      at: "2026-08-25T18:24:00.000Z",
      sessionIndex: 5,
      sessionStartedAt: "2026-08-25T17:00:00.000Z",
      prev,
      ...over,
    });

  it("holds the previous round's timestamp when readiness refused", () => {
    // THE LIVELOCK THIS FIXES. `deep-session` refused round 256 as "round 5 of
    // this session" and stamped `at: <now>, sessionIndex: 5` for a round that
    // ran no battery. The next attempt would be round 6, refused and stamped
    // again, climbing — a caller retrying promptly never gets a round out of it,
    // because every refusal pushes the 45-minute gap forward from itself.
    const r = receipt({ reason: "not-ready", roundFile: null });
    expect(r.at).toBe(prev.at);
    expect(r.sessionIndex).toBe(4);
    expect(r.heldSessionFrom, "a timestamp older than the writer must say why").toBeTruthy();
  });

  it("lets a finished round take its own place", () => {
    const r = receipt({ reason: "finished", roundFile: "256-abc.json" });
    expect(r.at).toBe("2026-08-25T18:24:00.000Z");
    expect(r.sessionIndex).toBe(5);
    expect(r.heldSessionFrom).toBeUndefined();
  });

  it("lets a CRASHED round take its own place too", () => {
    // A crash ran the battery until the host died. It worked the host, so it
    // owns its place — the session is about load, not about success.
    const r = receipt({ reason: "crashed", roundFile: null });
    expect(r.at).toBe("2026-08-25T18:24:00.000Z");
    expect(r.sessionIndex).toBe(5);
  });

  it("stamps normally when there is no previous receipt to hold", () => {
    // First round on a fresh machine. Holding nothing would write nulls and
    // make the next round think the session had never started.
    const r = outcomeReceipt({
      reason: "not-ready",
      codes: [],
      roundFile: null,
      at: "2026-08-25T18:24:00.000Z",
      sessionIndex: 1,
      sessionStartedAt: "2026-08-25T18:24:00.000Z",
    });
    expect(r.at).toBe("2026-08-25T18:24:00.000Z");
    expect(r.sessionIndex).toBe(1);
  });

  it("does not let a held receipt look like a session that never ends", () => {
    // Held forward, the gap is measured from the last REAL round — so waiting
    // out the rest clears it, which is the behaviour the stop is asking for.
    const held = receipt({ reason: "not-ready", roundFile: null });
    const after = sessionPosition(held, Date.parse(prev.at) + 46 * 60 * 1000);
    expect(after.index, "waiting past the gap must start a fresh session").toBe(1);
  });
});

describe("which kind of missing the add-in is", () => {
  it("calls a browser death absent, and puts it back", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { needsSideload } = await import("../scripts/round.mjs");
    // No pane, no command, nothing to open — a web sideload does not survive
    // the browser process.
    expect(needsSideload({ pane: null, canOpenPane: false, commandPresent: false })).toBe("absent");
  });

  it("does not call a FIRST sighting of a dead pane stale", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { needsSideload } = await import("../scripts/round.mjs");
    // THIS IS THE WHOLE GUARD. A `Disconnected` document greys its ribbon and
    // answers nothing, and it looks exactly like a sideload whose host has gone
    // — on one reading. It reconnects; the dead host does not. Spending the
    // one-per-process sideload on the first sighting would burn it on a
    // document that merely had to wait, which is the mistake the absent-case
    // guard already carries a comment about.
    expect(needsSideload({ pane: null, canOpenPane: true, commandPresent: true, sightings: 1 })).toBe(null);
  });

  it("calls it stale once a second attempt has seen the same thing", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { needsSideload } = await import("../scripts/round.mjs");
    // The state the 2026-08-27 domain move left: `Insert chart` in the ribbon,
    // opening a pane pinned to a host that had stopped existing. Readiness read
    // it as transient and retried seven times.
    expect(needsSideload({ pane: null, canOpenPane: true, commandPresent: true, sightings: 2 })).toBe("stale");
  });

  it("never asks for a sideload when the pane is open", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { needsSideload } = await import("../scripts/round.mjs");
    // A sideload is the most expensive thing the driver can do to a document.
    // A pane that is already there settles every other question.
    expect(needsSideload({ pane: "ref_1", canOpenPane: true, commandPresent: true, sightings: 9 })).toBe(null);
    expect(needsSideload({ pane: "ref_1", canOpenPane: false, commandPresent: false, sightings: 9 })).toBe(null);
  });
});

/**
 * THE RESTORE BUBBLE, and why the tidy-looking fix is the one to avoid.
 *
 * `sh("close")` ends the browser without Chromium's own shutdown, so the profile
 * keeps `exit_type: "Crashed"` and the next launch offers to restore the pages
 * the driver deliberately closed. Every `--fresh` leg left one behind.
 *
 * Closing the tabs first would look neater and is the move this repo has already
 * paid for twice: a PowerPoint tab with unsaved work raises `beforeunload`, and
 * accepting it discards the work AND the per-document SIDELOAD — measured after
 * rounds 124 and 132, which then refused with `addin-missing`. The deck sweep
 * runs immediately before this and guarantees unsaved work. So the tabs are left
 * alone and the flag is repaired instead.
 */
describe("leaving the browser profile as if it had exited on purpose", () => {
  const prefs = (over: Record<string, unknown> = {}) =>
    JSON.stringify({ profile: { exit_type: "Crashed", ...over }, other: { kept: true } });

  it("marks a crashed profile as cleanly exited", () => {
    let written = "";
    const changed = driver.clearCrashFlag("P", {
      exists: () => true,
      read: () => prefs(),
      write: (_p: string, s: string) => {
        written = s;
      },
    });
    expect(changed, "did not repair a profile flagged as crashed").toBe(true);
    const out = JSON.parse(written);
    expect(out.profile.exit_type).toBe("Normal");
    expect(out.profile.exited_cleanly).toBe(true);
    // Everything else in Preferences survives — this file carries the sign-in.
    expect(out.other.kept, "rewrote the profile and lost the rest of it").toBe(true);
  });

  it("leaves an already-clean profile alone", () => {
    let wrote = false;
    const changed = driver.clearCrashFlag("P", {
      exists: () => true,
      read: () => prefs({ exit_type: "Normal", exited_cleanly: true }),
      write: () => {
        wrote = true;
      },
    });
    expect(changed).toBe(false);
    expect(wrote, "rewrote a Preferences file that needed nothing").toBe(false);
  });

  it("does not guess at a profile it cannot read", () => {
    // A nuisance bubble is worth less than the sign-in this file holds, so an
    // unparseable Preferences is left exactly as found.
    let wrote = false;
    const changed = driver.clearCrashFlag("P", {
      exists: () => true,
      read: () => "{ not json",
      write: () => {
        wrote = true;
      },
    });
    expect(changed).toBe(false);
    expect(wrote, "wrote over a Preferences file it could not parse").toBe(false);
    expect(driver.clearCrashFlag("P", { exists: () => false, read: () => "", write: () => {} })).toBe(false);
  });

  it("repairs the flag on the fresh path, and closes no tabs doing it", async () => {
    const calls: string[] = [];
    const sh = ((...args: string[]) => {
      calls.push(args[0]);
      return "";
    }) as never;
    let cleared = "";
    await driver.startFresh(
      sh,
      (async () => {}) as never,
      (async () => {}) as never,
      "PROFILE",
      () => null,
      (p: string) => {
        cleared = p;
        return true;
      },
    );
    expect(cleared, "the fresh path did not repair the profile it just crashed").toBe("PROFILE");
    // The whole point: no tab is closed, so no beforeunload can fire.
    expect(calls, "closed a tab — that is what cost the sideload after rounds 124 and 132").not.toContain("tab-close");
    expect(calls).toContain("close");
  });
});
