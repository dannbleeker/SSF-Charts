#!/usr/bin/env node
/**
 * Drive a real-host round, and refuse to start one that cannot prove anything.
 *
 * A round costs ten minutes of a real PowerPoint and is the only evidence this
 * project gets about the host. Three of the four rounds run on 2026-08-14 were
 * set up by hand, fifteen browser steps at a time, and the steps that matter are
 * the ones easiest to skip:
 *
 *   - The pane is served with `Cache-Control: max-age=600`. Open it too soon
 *     after a merge and the round tests code the host never fetched. There is no
 *     way to tell from the result; the round simply means something other than
 *     what you think.
 *   - A round starting on a deck full of the last round's slides is not the same
 *     experiment as one starting clean, and the two have been compared as though
 *     they were.
 *   - Verbose trace and Picture every slide decide what the round can prove.
 *     Off, the trace is thin and empty slides cannot be confirmed empty.
 *
 * So this checks first and runs second, and every precondition is a hard stop
 * with the fix in the message. A round that runs on the wrong build is worse
 * than no round: it produces a file that looks like evidence.
 *
 *   node scripts/round.mjs --check     # preconditions only, nothing driven
 *   node scripts/round.mjs             # check, run, poll, archive — then triage BY HAND
 *
 * It does NOT triage. This line said it did until 2026-08-19, and a session
 * planned around it: reading the archived round is `node scripts/triage.mjs
 * rounds/0NN-<build>.json` and `npm run rounds`, both separate commands. A
 * usage block that claims a step nobody runs is the same defect class as the
 * three stale slogans in `triage.mjs`.
 *
 * NOT part of any gate, and it cannot be. It needs a signed-in PowerPoint on
 * the web with the add-in sideloaded, which exists on the owner's machine and
 * nowhere else. The decisions are pure and unit-tested; the browser is injected.
 *
 * See `docs/ROUNDS.md` for what to do with the round once it exists, and
 * `CLAUDE.md` "Looking at the task pane" for the browser traps this encodes.
 */
import { spawnSync } from "child_process";
import {
  readFileSync,
  writeFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  realpathSync,
  copyFileSync,
  statSync,
  rmSync,
} from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { isMain } from "./is-main.mjs";
import { collectCrashEvidence } from "./crash-forensics.mjs";

/**
 * The pane's build stamp, e.g. `32a6987 · 2026-08-14 08:03Z` → `32a6987`.
 *
 * THE SEPARATOR IS REQUIRED, and leaving it out cost a good round. Seven hex
 * characters is not a rare shape: `playwright-cli` names every element `f14e735`
 * and friends, and the accessibility dump this function reads is full of them.
 * Matching a bare 7-hex token picked up a REF id, reported the pane as showing a
 * build the site had never served, and refused to start — on a pane that was
 * showing exactly the right commit.
 *
 * A precondition that blocks correct work is worse than no precondition at all,
 * because the fix it suggests (hard-reload the tab) makes the reader doubt a
 * machine that was right.
 */
export function buildOf(text) {
  const m = /\b([0-9a-f]{7})\s+·/.exec(String(text ?? ""));
  return m ? m[1] : null;
}

/**
 * What to DO about a spawn that never ran — off checked facts, not the errno.
 *
 * `ENOENT` from this driver's `spawnSync` has two causes and names neither:
 * Node reports a missing working directory as ENOENT AGAINST THE EXECUTABLE, so
 * a missing `.pw-session/` and a missing playwright-cli print the same line —
 *
 *     spawnSync <the node.exe that would have run it> ENOENT
 *
 * — and the old message answered both with "Install it". The comment above it
 * already said "the install has never once been the problem" and the text still
 * said to install it. It cost the first round of the 2026-08-23 session: the
 * tool was installed, node was on PATH, and the clone simply had no session
 * directory, because `.pw-session/` is gitignored and only `scripts/pw.sh`
 * created it.
 *
 * The two inputs are CHECKABLE, so they are checked — see `cli`. This function
 * only reads the answer. The last branch is the honest one: when both exist,
 * naming either would be the same guess in a new coat.
 */
export function spawnFailureRemedy(at) {
  if (at?.cwdMissing)
    return (
      "the session directory does not exist — nothing below was actually measured. " +
      "Create it (`. scripts/pw.sh` does, with `mkdir -p`), or pass an existing --dir."
    );
  if (at?.entryMissing)
    return (
      "playwright-cli is not at the path this driver was given — nothing below was actually measured. " +
      "Install it (`npm i -g @playwright/cli`), or point PLAYWRIGHT_CLI_JS at its entry point."
    );
  return (
    "playwright-cli could not be run — nothing below was actually measured. " +
    "Its entry point and the session directory both exist, so this is " +
    "neither a missing install nor a missing directory."
  );
}

/**
 * Is this round worth running, and if not, what has to change first.
 *
 * Every no here is a round that would have produced a file nobody could draw a
 * conclusion from. `deployed !== head` is the one that looks harmless and is
 * not: it means the site has not finished publishing the commit under test, so
 * the pane will serve the previous build and the round will quietly measure it.
 */
export function readiness({
  head,
  deployed,
  stamp,
  slides,
  verbose,
  pictures,
  reachable = true,
  unreachableAt = null,
  browserGone = false,
  ping = null,
  slideOk = null,
  crashed = false,
  loggedOut = false,
  authPopup = false,
  size = null,
  expectSize = null,
  wantDeck = null,
  deckFronted = true,
  canOpenPane = true,
  // Whether the ribbon carries the command AT ALL, apart from whether it can be
  // clicked. Defaults true so an absent argument cannot invent a missing add-in.
  commandPresent = true,
  /**
   * Did a read that feeds this judgement FAIL TO RUN?
   *
   * Not "did it come back empty" — did the call itself never happen. Defaults
   * false so an absent argument cannot invent a failure, and when true nothing
   * below may be read as an absence, because nothing below was measured.
   */
  readsFailed = false,
  /**
   * The page's width when the ribbon was read, or null if it could not be read.
   *
   * PowerPoint hides trailing ribbon commands in an overflow on a narrow
   * window, and a hidden command is not in the accessibility tree — so
   * `commandPresent: false` means "absent OR merely not rendered", and those
   * need different answers. Defaults to a width known to be wide enough,
   * because an absent argument must not manufacture a narrow window.
   */
  ribbonRoom = MIN_RIBBON_WIDTH,
  /**
   * Which round of this session this is — see `sessionPosition`.
   *
   * Not a refusal. A deep round still produces a real sheet; it just answers
   * fewer questions than it claims to, and the reader deserves to be told which
   * kind of round they are about to read. Defaults to 1 so an absent argument
   * cannot invent a warning.
   */
  sessionIndex = 1,
  /**
   * Run anyway, however deep in the session this is.
   *
   * The depth rule is a STOP rather than a warning as of round 245, and this is
   * the door out. It was a warning for as long as it existed, printed in full,
   * and it was read and ignored — round 244 was the fifth back-to-back of a
   * block and its self-test half is worth less than the three before it for
   * exactly the reason the line predicted.
   *
   * Readiness refuses when "a round now would not prove anything", and a round
   * that silently skips its heaviest scenarios is the clearest case of that in
   * this file. It is not a HOST fault, though, which is why it takes a flag and
   * not a recovery: nothing the driver can do fixes it except waiting.
   */
  allowDeepSession = false,
}) {
  const stop = [];
  /**
   * A CODE beside every message, because retrying has to be decided on what a
   * stop IS rather than on how it is worded.
   *
   * `shouldRetry` used to retry a crash and nothing else, so the driver stopped
   * dead on three states it already knows how to fix — and on 2026-08-15 a
   * person fixed exactly those three by hand, in the order `recover` does them:
   * the host was silent (the quiet wedge), the pane was a build behind, the deck
   * held eight slides. One reload and reopen cleared all three.
   *
   * Matching on the messages instead would tie the retry loop to prose that is
   * edited whenever a message is improved. See `RECOVERABLE_STOPS`.
   */
  const codes = [];
  const refuse = (code, message) => {
    codes.push(code);
    stop.push(message);
  };
  // First, and on its own: everything below reads as "the browser said nothing"
  // when the truth is that nobody asked it. See `cli`.
  if (!reachable) {
    // WHICH of the three unreachable states this is. A tool that is missing, a
    // session directory that is missing, and a machine too busy to start the
    // tool all produce the same empty string here, and they want different
    // handling: fail fast with one remedy, fail fast with another, or retry.
    //
    // Split in two passes because the two questions are different. Retry-or-not
    // is decided off the ERRNO, which is the only thing that separates "too busy
    // just now" from "not there at all". Which-remedy is decided off the two
    // paths, CHECKED — because the errno cannot tell a missing entry from a
    // missing cwd, Node reporting both as ENOENT against the executable.
    //
    // The headline names the CONDITION rather than guessing at a cause. This
    // comment said so while the message underneath it still opened with
    // "Install it" — and the install has never once been the problem.
    const transient = unreachableAt && spawnFailureIsTransient(unreachableAt.error);
    return {
      ok: false,
      codes: transient ? ["cli-busy"] : [],
      stop: [
        (transient
          ? "playwright-cli could not be STARTED — the machine was busy, not the tool missing. " +
            "Nothing below was actually measured; the next attempt is expected to clear it."
          : spawnFailureRemedy(unreachableAt)) +
          // The call and the errno, when there is one. Without them this message
          // points at the install every time, and the install is almost never it.
          (unreachableAt
            ? `\n      the call that could not be spawned: \`${unreachableAt.args}\` — ${unreachableAt.error}`
            : ""),
      ],
    };
  }
  // Before the sign-in check and everything under it: a browser that is not
  // there answers every read with nothing, and "is the add-in open?" is the
  // wrong question to send anyone to. See `noBrowser`.
  //
  // AND BEFORE `readsFailed`, which is the more general fact and therefore the
  // less useful one. With no browser EVERY call exits non-zero — "The browser
  // 'default' is not open" — so a `readsFailed` check placed first swallows
  // this one and reports "a read could not be RUN" for a condition that has a
  // name, a cause and a recovery. Ordering the vague stop ahead of the specific
  // one is the same defect as reporting an absence for an unmeasured thing,
  // committed one layer up; it went in and out of this file within the hour.
  if (browserGone)
    return {
      ok: false,
      codes: ["browser-gone"],
      stop: [
        "there is no browser — the process died, taking the tab with it. The persistent profile still " +
          "holds the sign-in, so this is recoverable without a password: " +
          "`pw open --persistent --profile=C:/devtools/pw-profile --headed https://onedrive.live.com/`, " +
          "then open the deck, select its tab, and reopen the pane from Home ▸ Add-ins ▸ Insert chart.",
      ],
    };
  // A READ THAT COULD NOT RUN IS NOT AN ABSENCE — but only once the browser has
  // been ruled out above, because a missing browser makes every call fail and
  // has a far better answer than this one.
  //
  // `spawnSync` on this machine intermittently answers ENOENT for a node.exe
  // that is plainly there — eight consecutive calls succeeded minutes later. On
  // 2026-08-22 one such failure was reported as `pane ?`, `deck ? slide(s)`,
  // "could not read the pane's build stamp" and finally "the add-in is gone
  // from this document". The deck was open, the stamp was findable and `Insert
  // chart` was in the ribbon throughout. It cost an evening and a sideload walk
  // against a document that never needed one.
  //
  // RECOVERABLE, because the next attempt's reads usually run: this is the one
  // stop where trying again is exactly the right response.
  if (readsFailed)
    return {
      ok: false,
      stop: [
        "a read that this judgement depends on could not be RUN — not empty, never executed. " +
          "Nothing below was measured, so none of it can be reported as missing. Trying again.",
      ],
      codes: ["reads-failed"],
    };
  // THE WEDGE, by its real name. Rounds 24, 25, 29 and 30 each spent most of an
  // hour on this and none of them said what it was: the host's editing session
  // dies, and every Office.js call after that hangs forever without ever throwing
  // — even an empty `context.sync()`. `PowerPoint.run` still ENTERS its callback,
  // which is why it reads as the host thinking rather than the host being gone.
  //
  // This branch is the loud form, where PowerPoint has raised its own error and
  // put a modal over the document; the deck reading back as `?` in the same
  // breath is the UI frozen behind it. The quiet form has no dialog and only the
  // ping catches it. Neither is something to wait out — a reload clears both in
  // seconds. See `docs/ROUNDS.md`, "The wedge".
  // Before everything the browser could tell us, because if this is true then
  // nothing else was measurable and every message below is noise.
  //
  // UNLESS THE HOST IS ANSWERING, and that exception is the whole point. A
  // sign-in TAB is not a sign-in PROMPT. Chrome's "Restore pages?" brings back
  // every tab a crashed session had open, including an `oauth20_authorize`
  // popup that was already finished — so the tab list shows a login page beside
  // a deck that is working perfectly.
  //
  // Observed 2026-08-20: the check printed `host answered in 7ms · slide 1
  // resolved` and refused in the same breath, telling the owner to go and enter
  // a password they had already entered. The deck was fine; the tab was debris.
  //
  // A host that answers Office.js and resolves a slide IS an authenticated
  // session — an unauthenticated one cannot do either. So the ping is the
  // discriminator, and it is better evidence than the tab list because it asks
  // the thing we actually care about instead of looking at the furniture around
  // it. If the host is silent, the refusal stands and reads exactly as before.
  const hostAnswering = Boolean(ping?.answered) && slideOk !== false;
  if (loggedOut && hostAnswering)
    console.log(
      "  a Microsoft sign-in tab is open, but the host is answering — treating it as leftover from a " +
        "browser restore rather than a live prompt",
    );
  // AND WHEN THE HOST IS SILENT BUT THE DOCUMENT IS UP, SAY SO INSTEAD OF
  // GUESSING. `slides !== null` means the deck rendered its slide list, which a
  // signed-out browser cannot do — but it does NOT prove the session is fine,
  // because Office can raise a re-auth prompt beside a loaded deck, which is the
  // case this refusal was written for.
  //
  // So the two are genuinely indistinguishable from the tab list, and claiming
  // "sign in" is a guess that costs the owner a trip to the machine. A stale
  // `login.live.com` tab restored by Chrome after a crash is PERMANENT debris:
  // it never goes away on its own, so this refusal would fire on every future
  // round the moment the host went briefly quiet — which on this machine is the
  // commonest transient state there is.
  //
  // A silent host is already a recoverable stop with a reload behind it. Let
  // that run, and mention the tab rather than blaming it. If the host is still
  // silent after recovery, the round fails as `host-silent`, which is what was
  // actually observed.
  if (loggedOut && !hostAnswering && slides !== null)
    console.log(
      "  a Microsoft sign-in tab is open and the host is silent, but the deck's slide list still reads — " +
        "treating this as a silent host (recoverable) rather than a sign-in prompt. If a reload does not " +
        "fix it, the tab may be a real prompt after all.",
    );
  if (loggedOut && !hostAnswering && slides === null)
    return {
      ok: false,
      stop: [
        authPopup
          ? "Office has opened a sign-in prompt beside the deck — the deck tab is still there, but the host is " +
            "asking for credentials and nothing measured past this point can be trusted. " +
            "Sign in on the prompt, check the add-in pane is still open, then check again. " +
            "None of that is the agent's to do: it needs a password."
          : "the browser is on a Microsoft sign-in page — there is no document to run a round in. " +
            "Sign in, open the deck, sideload the add-in from Home ▸ Add-ins, then check again. " +
            "None of that is the agent's to do: it needs a password.",
      ],
    };
  if (crashed)
    refuse(
      "crashed",
      'PowerPoint has crashed — its own "Sorry, we ran into a problem" dialog is up, and every Office.js ' +
        "call behind it hangs forever. Click Refresh in that dialog; the host answers again within the minute. " +
        "The pane closes with it — reopen from Home ▸ Add-ins ▸ Insert chart.",
    );
  // Before anything about builds or decks: is the host answering? A stale pane
  // and a dirty deck are both worth fixing, and neither matters if the host will
  // not talk — that is the state four rounds have burned an hour each on.
  if (ping && !ping.answered)
    refuse(
      "host-silent",
      `the host did not answer the cheapest possible call in ${ping.ms}ms — it is not going to answer a round. ` +
        "Its editing session is gone, usually because the network moved (look for net::ERR_NETWORK_CHANGED in " +
        "`playwright-cli console`). Reload the PowerPoint tab and reopen the pane: measured 8011ms silent before, " +
        "7ms after.",
    );
  // AFTER the ping, because a host that answered `getCount` and then refused a
  // slide is a different state from one that answered nothing, and the fix is
  // the same only by coincidence. See `slideResolveScript`: this is the call the
  // 2s crash dies on, moved to where it costs two seconds instead of a round.
  if (slideOk === false)
    refuse(
      "slide-refused",
      "the host answered the cheap call but would not resolve slide 1 — this is the state the 2s crash " +
        "starts from (`OnServerFindSucceeded could not find target slide` in its own log). Reload the " +
        "PowerPoint tab and reopen the pane; an attempt that follows a recovery has never crashed.",
    );
  // NAMES BOTH CAUSES, BECAUSE IT CANNOT TELL THEM APART. This used to read
  // "is Pages up?", which is a diagnosis, and it was the wrong one on
  // 2026-09-05: every round refused here while `curl` fetched the same URL in
  // 850ms. The fault was local — node's `fetch` gives each address 250ms, a TCP
  // connect to Pages from that machine took 282ms, so it abandoned IPv4 and
  // fell through to an IPv6 route the machine did not have. Two rounds were
  // spent looking at GitHub. A refusal that points confidently at the wrong
  // thing costs more than one naming both candidates and how to separate them.
  if (!deployed)
    refuse(
      "no-build",
      "could not read a build stamp from the site. Either Pages is down, or the network from HERE cannot " +
        "reach it — node's `fetch` gives each address 250ms, and a slower connect falls through to an address " +
        "this machine may have no route to. Tell them apart with " +
        "`curl https://ssf-chart.struktureretsundfornuft.dk/build.json`: if that answers, the site is fine and " +
        "the round wants `NODE_OPTIONS=--network-family-autoselection-attempt-timeout=2000`.",
    );
  else if (head && deployed !== head)
    refuse("site-behind", `the site is serving ${deployed} but HEAD is ${head} — wait for Deploy Pages to finish`);
  // `slides !== null` is the proof the DOCUMENT is up. Without it this fires on
  // a tab that is merely mid-reload — every `refFor` answers nothing there too —
  // and it is a refusal recovery is forbidden to retry, so a transient state
  // would end the night. Unknown is not the same as wrong; the 4:3 deck that
  // motivated this read its slide list perfectly well and simply had no add-in.
  if (!stamp && !canOpenPane && slides !== null && commandPresent)
    // PRESENT BUT UNUSABLE, which is the opposite conclusion from the one below
    // and was reaching it. A `Disconnected` document greys its whole ribbon —
    // `Insert chart`, `Add-ins`, everything — and that is transient: the tab
    // reconnects, or a reload fixes it. Recoverable on purpose.
    refuse(
      "host-disconnected",
      "the SSF Charts command is in the ribbon but DISABLED — the document is not connected, " +
        "so nothing in the ribbon can be clicked. This clears on its own or with a reload; " +
        "it is not a missing add-in.",
    );
  else if (!stamp && !canOpenPane && slides !== null && !wideEnoughToJudge(ribbonRoom))
    // TOO NARROW TO TELL, and this refusal is the one that must never guess.
    //
    // On 2026-08-21 a 1237px window hid `Insert chart` in PowerPoint's `...`
    // overflow, the driver read it as a missing add-in, ran the sideload walk
    // four times against a deck that already had it, and sent the owner to do
    // the job by hand. Widening the window to 2375px brought the command
    // straight back. NOTHING WAS EVER MISSING.
    //
    // Recoverable, unlike `addin-missing` below: the window is not a
    // measurement surface, so `ensureRibbonRoom` simply widens it and the next
    // attempt reads a ribbon that is allowed to show what it has.
    refuse(
      "ribbon-cramped",
      `the window is ${ribbonRoom === null ? "an unreadable width" : `${ribbonRoom}px wide`} and PowerPoint hides ribbon ` +
        `commands below about ${MIN_RIBBON_WIDTH}px — a command that is not RENDERED is not in the accessibility tree, ` +
        "so this cannot tell a missing add-in from a hidden one. Widening and re-reading.",
    );
  else if (!stamp && !canOpenPane && slides !== null)
    // A DIFFERENT REFUSAL, because recovery cannot touch it. `recover` reopens
    // the pane from the ribbon's `Insert chart` control; a document that does
    // not offer that control has no add-in to open, and reloading it forever
    // will not produce one. Deliberately absent from `RECOVERABLE_STOPS`, so
    // this stops on the first attempt instead of the seventh.
    //
    // GUARDED ON THE NAME BEING ABSENT NOW, not merely on there being no ref.
    // Those are different states and this fired on both, turning a disconnected
    // document into a stop only the owner could clear.
    refuse(
      "addin-missing",
      "this document has no SSF Charts command in its ribbon — the add-in is not loaded here. " +
        // "A reload will not bring it back" stood here until 2026-08-20 and was
        // FALSE. Round 117 said it; the owner touched nothing; `Insert chart`
        // and `Insert element` were both in the ribbon minutes later, because
        // the sideload this driver had already performed simply had not
        // surfaced yet. Telling someone a machine cannot fix something it has
        // in fact just fixed is the worst thing this file can print.
        "If the driver has just uploaded it, LOOK AGAIN before doing anything — the ribbon can take a " +
        "minute to show a freshly sideloaded command. Otherwise sideload it into this deck " +
        "(Add-ins ▸ Upload My Add-in), or run against a deck that already has it.",
    );
  else if (!stamp) refuse("pane-closed", "could not read the pane's build stamp — is the add-in open?");
  else if (deployed && stamp !== deployed)
    refuse(
      "pane-stale",
      `the pane is showing ${stamp} while the site serves ${deployed} — hard-reload the whole ` +
        `PowerPoint tab (the pane HTML is cached for ten minutes; reopening the pane alone does not clear it)`,
    );
  if (slides !== null && slides > 1)
    refuse(
      "deck-dirty",
      `the deck holds ${slides} slides — clean it, or this round is not comparable with one that started clean`,
    );
  // BEFORE the size check, because it explains a wrong size rather than
  // repeating it. A cycle names the deck each leg should run against, and until
  // the driver actually fronted that tab the third leg measured whichever
  // document leg two left open — refusing with `wrong-size`, which reads as
  // "the owner set the slide size wrong" when nothing had asked for the right
  // deck at all. Naming the deck is a far better message than naming its size.
  if (wantDeck && !deckFronted)
    refuse(
      "deck-missing",
      `no open tab is the deck \`${wantDeck}\` — this round was told to run against it and cannot. ` +
        "Open it, or unset PW_DECK to run against whichever document is in front.",
    );
  // THE SIZE THE DECK ACTUALLY IS, when a profile was asked for. Only when both
  // are known: a host that would not answer has said nothing, and refusing on no
  // evidence is what `reachable` exists to prevent.
  //
  // AND ONLY WHEN IT IS THE RIGHT DECK'S SIZE, which is the half the comment
  // above only half fixed. Reordering the two checks improved the MESSAGE and
  // left the refusal firing: when the wanted deck is not fronted, `size` is the
  // size of whatever other document is, so this refuses `wrong-size` about a
  // deck this round was never going to run against.
  //
  // That is fatal rather than cosmetic. `deck-missing` is in `RECOVERABLE_STOPS`
  // and `recover` knows how to open a deck; `wrong-size` is deliberately NOT,
  // because setting a size changes what a round measures rather than restoring
  // it. So the bogus companion turns a stop the driver can clear into one it
  // cannot, and the whole night ends.
  //
  // It cost a full cycle on 2026-08-27 and again on 2026-08-29, and it bites the
  // NORMAL sequence: leg 3 runs at 4:3 and leaves that deck fronted, so the next
  // cycle's leg 1 wants the 16:9 deck and measures the 4:3 one that is still in
  // front. A single cycle never sees it. On 2026-08-29 the fix by hand was to
  // open the deck — which is exactly what recovery would have done unasked.
  if (expectSize && size && (!wantDeck || deckFronted) && size !== expectSize)
    refuse(
      "wrong-size",
      `the deck is ${size} and this round was asked for ${expectSize} — rerun with \`PW_SET_SIZE=1\` to have ` +
        "the driver set it (a deck that EXISTS to be that profile only), or set it in Design ▸ Slide Size and " +
        "CHECK IT TOOK, because a click made while the document is loading is accepted and does nothing. " +
        "Filing a round under the wrong profile is worse than not running it.",
    );
  if (verbose === false) refuse("verbose-off", "Verbose trace is off — the round's trace will be too thin to mine");
  if (pictures === false)
    refuse("pictures-off", "Picture every slide is off — a slide that reads back empty cannot be confirmed empty");
  // LAST, so a deep round still reports everything else wrong with it. Refusing
  // early would hide a stale pane behind "you are tired", and the owner would
  // rest for an hour and then meet the pane.
  const deep = sessionDepthWarning(sessionIndex);
  if (deep.length && !allowDeepSession)
    refuse("deep-session", `${deep[0]} Pass \`--deep\` to run anyway and file it knowing what it is.`);
  return {
    ok: stop.length === 0,
    stop,
    codes,
    // Not also a warning once it is a stop — saying it twice in one report reads
    // as two separate problems.
    warn: allowDeepSession ? deep : [],
  };
}

/** The next round number, from the archive. */
export function nextRoundNumber(files) {
  const ns = files.map((f) => Number(/^(\d{3})-/.exec(f)?.[1])).filter((n) => Number.isFinite(n));
  return String((ns.length ? Math.max(...ns) : 0) + 1).padStart(3, "0");
}

/** Strip the base64 slide images — see `rounds/README.md`. */
export function stripImages(round) {
  const walk = (o) => {
    if (!o || typeof o !== "object") return;
    for (const k of Object.keys(o)) {
      const v = o[k];
      if (typeof v === "string" && v.length > 500 && /^[A-Za-z0-9+/=]+$/.test(v.slice(0, 200)))
        o[k] = "<image stripped for the archive — see rounds/README.md>";
      else walk(v);
    }
  };
  walk(round);
  return round;
}

/**
 * The CLI's own JavaScript, to be run by node directly.
 *
 * Three routes were tried on the owner's machine and two are dead ends:
 *
 *   spawnSync("playwright-cli", …)               ENOENT — it is not an .exe
 *   spawnSync("playwright-cli", …, {shell:true}) "This program is blocked by
 *                                                group policy"
 *
 * AppLocker blocks the `.cmd` shim, so every route through a shell is closed
 * here no matter how the arguments are quoted. `node <entry>` works, because
 * node.exe is allow-listed and it is what the shim would have run anyway. It
 * also sidesteps cmd.exe re-parsing arguments, which matters: the stamp lookup
 * passes a regex containing spaces and braces, and the click helper passes
 * JavaScript containing `=>` and `{ }` — and `>` is a redirect character.
 *
 * Global installs sit beside node itself in this layout, and
 * `PLAYWRIGHT_CLI_JS` overrides for anyone whose does not.
 */
export function cliEntry(execPath = process.execPath, exists = existsSync) {
  if (process.env.PLAYWRIGHT_CLI_JS) return process.env.PLAYWRIGHT_CLI_JS;
  const guess = join(dirname(execPath), "node_modules", "@playwright", "cli", "playwright-cli.js");
  return exists(guess) ? guess : null;
}

/**
 * `playwright-cli`, and a way to know it was never reached.
 *
 * A tool that cannot be invoked and a pane that is closed produce the same empty
 * string, and folding the first into the second sends the reader to fix a
 * browser that is fine. On this driver's first live run it reported "could not
 * read the pane's build stamp — is the add-in open?" while the pane sat open on
 * screen showing the stamp. So `unreachable` is tracked across the session and
 * reported as its own precondition, ahead of everything it would otherwise
 * poison.
 */
export function ensureSessionDir(dir) {
  try {
    mkdirSync(dir, { recursive: true });
  } catch {
    // Leave it. The spawn then fails with `cwdMissing`, and `spawnFailureRemedy`
    // says so in words a reader can act on — better than an ENOENT stack thrown
    // out of the driver's first line, before anything has been reported.
  }
}

export function sessionDir(dir, real = realpathSync.native) {
  // The daemon keys its sessions by the working directory STRING, and Windows
  // gives the same directory two names. Opened from a shell that resolved
  // `C:\Users\dann.pedersen\...`, the session is invisible to a process whose cwd
  // is the 8.3 form `C:\Users\DANN~1.PED\...` — same folder, different string,
  // and `list` answers "(no browsers)" while the browser sits on screen. That
  // cost most of an afternoon's debugging on this driver alone.
  try {
    return real(dir);
  } catch {
    return dir;
  }
}

/**
 * Did this call RUN and say too much, rather than fail to run at all?
 *
 * `spawnSync` reports both as `error`, and the driver treated both as "the tool
 * could not be run". They are opposite facts: an overflow means the browser
 * answered, and every other spawn error means nobody asked. Round 044 spent two
 * attempts and two empty crash reports on the difference.
 *
 * Matched on the CODE, with the message as a fallback, because Node stamps
 * `err.code = "ENOBUFS"` on the overflow path while the message it prints names
 * the executable rather than the reason.
 */
export function isOverflow(err) {
  if (!err) return false;
  return err.code === "ENOBUFS" || /ENOBUFS/.test(String(err.message ?? err));
}

/** Which manifest a re-sideload uploads. The PROD one — see `sideloadAddIn`. */
export const MANIFEST_PATH = process.env.PW_MANIFEST ?? "C:/devtools/SSF-Charts/manifest-prod.xml";

/**
 * Has this process already tried to put the add-in back? One attempt, ever.
 *
 * A latch rather than a retry budget, because the failure it guards is not
 * transient: if the ribbon walk did not work the first time, it is because
 * Microsoft moved something, and walking it six more times finds the same
 * missing control six more times. `addin-missing` stays out of
 * `RECOVERABLE_STOPS`; this runs BEFORE readiness concludes, so a success is
 * simply a round that starts and a failure is the refusal the driver already
 * had. Per process, and the cycle runner spawns one per leg.
 */
let sideloadAttempted = false;

/**
 * How many attempts have found a ribbon command whose pane will not open.
 *
 * THE STATE A DOMAIN MOVE LEAVES, and the driver had no name for it. A web
 * sideload pins the pane's URL: after the cutover to `ssf-chart` on 2026-08-27,
 * `Insert chart` was still in the ribbon and still opened a pane that could
 * never load, because the host it named had stopped existing. `commandPresent`
 * was true, so the walk below never ran, and readiness read it as the TRANSIENT
 * `host-disconnected` — a document that merely has to reconnect. It retried
 * seven times against a cause no amount of retrying can reach.
 *
 * Retrying is exactly what tells the two apart. A disconnected document comes
 * back; a sideload pointing at a dead host does not. So the count, not the
 * single reading, is what unlocks the sideload — the second attempt to see it
 * is evidence the first could not be.
 */
let commandWithoutPane = 0;

/**
 * Does the add-in need putting back, and which kind of missing is it?
 *
 * `"absent"` · `"stale"` · `null`, and the three are decided here rather than in
 * the driver's flow so the rule can be stated and tested instead of inferred
 * from a five-clause `if`.
 *
 * ABSENT is a browser death: no pane, no command, nothing to open. STALE is a
 * command that is still in the ribbon opening a pane that cannot load, which is
 * what a domain move leaves and what read as a transient disconnect for seven
 * attempts on 2026-08-27.
 *
 * `sightings` is what separates stale from transient, and nothing else can: a
 * disconnected document reconnects on the next attempt, a sideload pointing at
 * a host that no longer exists does not. One reading cannot tell them apart, so
 * the second is the evidence.
 */
export function needsSideload({ pane, canOpenPane, commandPresent, sightings = 0 }) {
  if (pane) return null;
  if (!canOpenPane && !commandPresent) return "absent";
  if (commandPresent && sightings >= 2) return "stale";
  return null;
}

/** Test-only: let a suite arm the latch more than once. */
export function _resetSideloadLatchForTest() {
  sideloadAttempted = false;
  commandWithoutPane = 0;
}

/**
 * Put the add-in back after a browser death took the sideload with it.
 *
 * A web sideload does not survive the browser process. `recover` restores the
 * window, the deck and its tab — and then finds no SSF Charts command in the
 * ribbon, because there is no add-in to open. On 2026-08-16 that ended a night
 * twice, about fifteen minutes each time.
 *
 * THE WALK IS THE ONE VERIFIED BY HAND on 2026-08-17, and the first step is the
 * one nobody would guess: the ribbon's `Add-ins` button ignores clicks until the
 * document surface has focus. Three attempts failed on that alone, all of them
 * reporting `aria-expanded=true` over a menu that was plainly shut in a
 * screenshot. The accessibility tree lied; the pixels did not.
 *
 * `manifest-prod.xml`, never `manifest.xml`: the latter points at
 * `https://localhost`, so it would sideload a pane that readiness then refuses
 * for disagreeing with the deployed site — a worse outcome than no pane, because
 * it looks like it worked.
 *
 * NO NATIVE FILE DIALOG EVER OPENS. Playwright intercepts the chooser, so
 * `upload` hands the path over directly. That is the only reason this is
 * automatable at all.
 *
 * Best-effort and self-cleaning. Every failure path dismisses whatever is open,
 * because a modal left over the document makes every later round refuse for a
 * reason that has nothing to do with the round.
 */
export async function sideloadAddIn(sh, sleep, manifest = MANIFEST_PATH) {
  const click = (ref) => sh("eval", "el => { el.click(); return 'ok'; }", ref);
  const step = (query, pattern) => refFor(sh, query, pattern);
  const giveUp = (why) => {
    console.error(`  could not put the add-in back: ${why}`);
    // ALWAYS, even when nothing looks open. A half-walked dialog is invisible
    // to `readiness` and fatal to every round after it.
    const cancel = step("Cancel", /button "Cancel"/);
    if (cancel) click(cancel);
    const close = step("Close", /button "Close"/);
    if (close) click(close);
    return false;
  };
  console.log("  the add-in is gone from this document — putting it back");

  // A RELOAD BEFORE THE WALK, because the walk cannot open a menu that is
  // already open. Twice on 2026-08-21 the first attempt died at "the Add-ins
  // menu did not open", and both times the cause was a `button "Add-ins"
  // [expanded]` left over from a PREVIOUS failed attempt: clicking an open menu
  // closes it, so every retry toggled the menu instead of walking it, and the
  // driver reported the ribbon as unopenable when it was merely already open.
  //
  // ESCAPE IS NOT ENOUGH — measured. It returns the button to `[active]` and
  // the very next attempt still failed; only a reload let the walk run to the
  // upload, both times.
  //
  // This can discard unsaved work, which is the fault that got the between-
  // rounds reload pulled. Acceptable HERE and nowhere else: this path only runs
  // when the add-in is already missing, and a document with no add-in cannot
  // produce a round whose work would be worth saving.
  sh("reload");
  sh("dialog-accept");

  // WAITED FOR, NOT SLEPT THROUGH. This was a fixed 55s sleep and one look, and
  // it failed on the first cold browser start it met: round 133, after a machine
  // restart, reported `no Add-ins button in the ribbon` — OFFICE'S OWN button,
  // not SSF Charts' — because the document was still loading 55 seconds after a
  // reload it had barely begun. Checked by hand a minute later, the whole ribbon
  // was there.
  //
  // ON THE SLIDE LIST, not the ribbon: that is the signal the walk already uses
  // for "the document is up", and probing the ribbon first would start touching
  // it before knowing there is a document to touch — which the walk is
  // deliberately built not to do.
  //
  // Fourth time this shape has been wrong in two days: `sideloadAddIn` after its
  // upload, `commandPresent` before deciding to sideload, `refreshPane` after
  // its reload, and now the reload this function does itself. A fixed sleep
  // encodes a guess about the host's speed; the host does not know about it.
  await waitForRef(sh, sleep, "Slide List", /listbox "Slide List"/, RIBBON_WAKE_BUDGET_MS * 3, 5000);

  // FIRST, and it is not optional: the ribbon will not open its menus while the
  // document surface is unfocused.
  const list = step("Slide List", /listbox "Slide List"/);
  if (!list) return giveUp("the deck's slide list is not readable — the document is not up");
  click(list);
  await sleep(2000);

  const addins = step("Add-ins", /button "Add-ins"/);
  if (!addins) return giveUp("no Add-ins button in the ribbon");
  click(addins);
  await sleep(5000);

  // OPTIONAL, BECAUSE TWO UIS ARE IN THE FIELD AND ONLY ONE HAS THIS RUNG.
  //
  // On 2026-09-05 a brand-new presentation opened the Add-ins flyout straight
  // onto the store panel — no "See all installed add-ins" anywhere in it, with
  // `More Add-ins` sitting right there as the next rung. This walk refused with
  // "the Add-ins menu did not open", which was false twice over: the menu HAD
  // opened, and the thing it was looking for is not on this variant at all.
  //
  // Absence is the ONLY case treated as "carry on". A `See all` that is present
  // is still clicked, because on the older UI the store panel is behind it and
  // skipping it lands the next step on nothing. What decides the walk is the
  // rung after this one — `More Add-ins` is on both variants, and its absence
  // is still fatal, so a genuinely unopened menu is caught one line later with
  // a message that is true.
  const seeAll = step("See all", /menuitem "See all installed add-ins"/);
  if (seeAll) {
    click(seeAll);
    await sleep(6000);
  }

  const more = step("More Add-ins", /menuitem "More Add-ins"/);
  if (!more)
    return giveUp(
      seeAll
        ? "no `More Add-ins` entry"
        : "the Add-ins menu did not open — neither `See all installed add-ins` nor `More Add-ins` is in it",
    );
  click(more);
  await sleep(9000);

  const mine = step("MY ADD-INS", /tab "MY ADD-INS"/);
  if (!mine) return giveUp("the Office Add-ins dialog did not open");
  click(mine);
  await sleep(5000);

  const manage = step("Manage My Add-ins", /button "Manage My Add-ins"/);
  if (!manage) return giveUp("no `Manage My Add-ins` control");
  click(manage);
  await sleep(5000);

  const upload = step("Upload My Add-in", /menuitem "Upload My Add-in"/);
  if (!upload) return giveUp("no `Upload My Add-in` entry");
  click(upload);
  await sleep(6000);

  const browse = step("Browse", /button "Browse\.\.\."/);
  if (!browse) return giveUp("the upload dialog has no Browse button");
  click(browse);
  await sleep(2000);
  sh("upload", manifest);
  await sleep(3000);

  // THE BUTTON'S OWN STATE IS THE RECEIPT. It is disabled until a file is
  // accepted, so finding it enabled is the host confirming it took the
  // manifest — better evidence than the upload call not erroring.
  const go = refFor(sh, "Upload", /button "Upload"(?! \[disabled\])/);
  if (!go) return giveUp("the manifest was not accepted — Upload stayed disabled");
  click(go);
  // POLLED, NOT SLEPT THROUGH. The ribbon does not repopulate on a schedule,
  // and a single look after a fixed sleep turns "not yet" into "never" — see
  // `SIDELOAD_COMMAND_BUDGET_MS` for the round that proved it.
  const open = await waitForRef(sh, sleep, "Insert chart", /button "Insert chart"/, SIDELOAD_COMMAND_BUDGET_MS);
  if (!open)
    return giveUp(
      `uploaded, but no SSF Charts command appeared within ${SIDELOAD_COMMAND_BUDGET_MS / 1000}s — ` +
        "the upload may still land, so re-check the ribbon before sideloading by hand",
    );
  click(open);
  await sleep(25000);
  return true;
}

/**
 * Bring the named deck's tab to the front, and say whether it is there at all.
 *
 * `PW_DECK` used to reach only `recover`, in the branch that reopens a browser
 * that died — so a cycle setting it per leg was choosing which deck a RECOVERY
 * would look for and nothing else. The ordinary path never selected a tab; it
 * ran against whatever happened to be fronted. The nightly cycle's third leg
 * therefore asked for a 4:3 deck, got the 16:9 one still open from leg two, and
 * refused with `wrong-size` every time — a stop no recovery addresses, so the
 * night ended there, reading as an operator's mistake when the runner had simply
 * never asked for the deck.
 *
 * Returns `true` when that deck is now fronted, `false` when no tab carries the
 * name. Never throws: a `tab-list` that could not be read is a reading not
 * taken, and `readiness` has its own, better-worded refusals for a browser that
 * is not answering.
 */
/**
 * The name of the document currently in front, or null.
 *
 * WHY A CHECK MUST SAY WHICH DECK IT CHECKED. Every line the readiness block
 * printed described the round — build stamps, slide counts, the host's latency —
 * without once naming the document it was describing. So a check run against
 * the wrong deck read exactly like a check run against the right one, and on
 * 2026-08-20 that is precisely what happened: the dedicated 4:3 deck was still
 * 960x540, the notes recorded the 4:3 leg as "blocked on owner setup", and no
 * output anywhere contradicted either belief.
 *
 * Null when the tab list could not be read or carries no document — a reading
 * not taken, printed as `?` rather than guessed at. `readiness` has its own,
 * better-worded refusals for a browser that is not answering.
 */
export function frontedDeck(sh) {
  const line = sh("tab-list")
    .split("\n")
    .find((l) => l.includes("(current)"));
  if (!line) return null;
  const m = /\[([^\]]+\.pptx)\]/i.exec(line);
  return m ? m[1] : null;
}

/**
 * The fronted deck's name IF THIS ARCHIVE IS ALLOWED TO PUBLISH IT.
 *
 * `rounds/*.json` and `crashes/*.json` are committed, and this repository is
 * public. `.gitignore` already says what that means, and says it about exactly
 * this field: crash reports are excluded because they carry "request URLs,
 * session ids, the document name. This repo is public. ... the reports never
 * leave the machine." The `.md` dumps were held back for it; the `.json` files
 * beside them were not, because until 2026-09-02 they never carried a name.
 *
 * Then `driverDeck` was added to break the deck/geometry confound and wrote
 * whatever `frontedDeck` read straight into a tracked file. On the test decks
 * that is harmless. The hazard is the ordinary path: with no `PW_DECK` the
 * driver runs against WHATEVER TAB IS IN FRONT, so one round started beside a
 * real presentation publishes its filename to a public repository.
 *
 * SO ONLY NAMES THE OPERATOR HIMSELF NAMED ARE PUBLISHED. `PW_DECK`,
 * `PW_DECK_16_9`, `PW_DECK_4_3` and the cycle's own defaults are decks somebody
 * configured on purpose; anything else is a document this archive has no
 * business naming, and is recorded as `undisclosed`.
 *
 * `undisclosed` is deliberately not null. Null already means "the tab list
 * would not say", and collapsing "we could not look" into "we looked and will
 * not tell" is the same defect this file spent the day removing from slide
 * counts. A reader can tell all three apart.
 */
export function archivableDeck(name) {
  if (!name) return null;
  const configured = [
    process.env.PW_DECK,
    process.env.PW_DECK_16_9,
    process.env.PW_DECK_4_3,
    // The cycle's own defaults, so an unconfigured nightly run still records
    // which arm it was on. Kept in step with `cyclePlan` in `cycle.mjs`.
    "Presentation64",
    "Presentation70",
  ].filter(Boolean);
  // Matched with and without the extension: `PW_DECK` carries a bare name
  // (`Presentation70`) and the tab list reports a filename.
  const base = name.replace(/\.pptx$/i, "").toLowerCase();
  return configured.some((d) => d.replace(/\.pptx$/i, "").toLowerCase() === base) ? name : "undisclosed";
}

export function selectDeck(sh, deckName) {
  if (!deckName) return true;
  const line = sh("tab-list")
    .split("\n")
    .find((l) => l.includes(deckName));
  if (!line) return false;
  const n = /(\d+):/.exec(line)?.[1];
  if (!n) return false;
  sh("tab-select", n);
  return true;
}

/**
 * How long a single CLI call may take before it counts as a wedge.
 *
 * Env-overridable because the one thing that would make this wrong is a host
 * slower than any yet seen, and nobody debugging that at 2am should have to
 * edit a script to get past it.
 */
export const CLI_TIMEOUT_MS = Number(process.env.PW_CLI_TIMEOUT_MS) || 180_000;

/**
 * The round browser's own playwright-cli config, or nothing.
 *
 * WHAT IT FIXES. The round browser was opened with no config at all, which left
 * it on a FIXED 2880x1800 page viewport inside a 1280x752 window — the page laid
 * out more than twice as wide as the window could show, so the right-hand strip
 * of PowerPoint, which is exactly where the task pane lives, fell outside the
 * visible area. The owner could not watch a round; maximising did nothing,
 * because the window was never the constraint. Measured, both ways:
 *
 *     no config      innerWidth 2880   outerWidth 1280   (2.25x mismatch)
 *     viewport null  innerWidth 1036   outerWidth 1050   (matches)
 *
 * `viewport: null` makes the page follow the window, which is what a headed
 * browser someone is meant to look at should always have done. Verified in a
 * throwaway session first, so the round browser was never at risk.
 *
 * NOT `.playwright/cli.config.json`. That one is 512x900 on purpose — the task
 * pane at its real width, for judging how the pane LOOKS (see CLAUDE.md). It has
 * nothing to do with a round, and pointing a round at it would be worse than the
 * bug: 512px is narrower than PowerPoint's own chrome.
 *
 * Absent file means no flag, rather than a flag to a path that is not there:
 * every CLI call carries this, and a bad path would fail all of them at once.
 *
 * TAKES EFFECT AT THE NEXT BROWSER OPEN. A context's viewport is fixed when it
 * is created, so a browser already running keeps whatever it started with.
 *
 * ONLY ON `open`. The first version of this passed the flag on EVERY CLI call,
 * which reads as harmless and is not: `--config` is an option of `open` alone,
 * so every other command answered `Unknown option: --config` and the driver
 * reported a healthy browser as `pane ?`, `deck ?`, "the pane is closed". A
 * config flag that breaks the readiness check would have looked exactly like
 * the crash it is meant to help someone watch.
 */
export function roundConfigArg(exists = existsSync) {
  const p = fileURLToPath(new URL("../.playwright/round.config.json", import.meta.url));
  return exists(p) ? [`--config=${p}`] : [];
}

export function cli(run, dir, entry = cliEntry(), ensure = ensureSessionDir) {
  const state = { unreachable: false };
  // MAKE THE SESSION DIRECTORY BEFORE ANYTHING IS SPAWNED INTO IT.
  //
  // `.pw-session/` is gitignored, so a FRESH CLONE does not have one — and
  // `spawnSync` reports a missing `cwd` as ENOENT against the EXECUTABLE:
  //
  //     spawnSync <the node.exe that would have run it> ENOENT
  //
  // which reads as a broken node install, and which the readiness message used
  // to answer with "install playwright-cli". Three wrong answers from one errno.
  // It cost the first round of the 2026-08-23 session, on a clone where the tool
  // was installed, node was on PATH, and the directory was the entire problem.
  // `scripts/pw.sh` has always done this with `mkdir -p`; the driver assumed
  // someone had sourced it first, which is true of every machine that has run a
  // round before and of no machine setting one up.
  //
  // BEFORE `sessionDir`, not after: `realpathSync.native` throws on a path that
  // does not exist, and the fallback hands back the un-normalised string — the
  // 8.3 short-name bug that function exists to prevent.
  ensure(dir);
  const cwd = sessionDir(dir);
  const sh = (...args) => {
    if (!entry) {
      state.unreachable = true;
      return "";
    }
    // NODE, on the CLI's own JavaScript. No shim, no shell.
    //
    // `maxBuffer` because the default is 1 MiB and `requests` on a live
    // PowerPoint tab is bigger than that — the document channel alone runs to
    // hundreds of POSTs with query strings on them. Over the line, `spawnSync`
    // returns ENOBUFS and throws the output away, which is how round 044 lost
    // two crash reports and then the round itself.
    // BOUNDED, because the round's own deadline cannot bound this. The poll
    // loop checks a 30-minute limit at the TOP of each pass, which only ever
    // runs if the call below returned — so a CLI that wedges (a browser that
    // stopped answering CDP, a tab mid-crash) hangs the driver with the
    // deadline sitting there unreachable and nothing on screen. That is the
    // exact shape of an overnight run that is found dead in the morning having
    // printed nothing since hour one.
    //
    // Generous on purpose: the slowest legitimate call is an `eval` carrying a
    // 20s page-side budget, and `requests` on a live tab can return tens of
    // megabytes. Three minutes is far above both and far below a night. A
    // timeout arrives as `r.error`, which the existing branch below already
    // reads as "nothing was measured" — which is exactly what it is.
    const r = run(process.execPath, [entry, "-s=ms", "--raw", ...args], {
      encoding: "utf8",
      cwd,
      maxBuffer: 64e6,
      timeout: CLI_TIMEOUT_MS,
    });
    // A CALL THAT RAN AND SAID TOO MUCH IS NOT A TOOL THAT COULD NOT BE RUN, and
    // conflating them is what sent two rounds' debugging at a healthy install.
    // ENOBUFS means the browser answered and the answer did not fit; every other
    // spawn error means nothing was measured. Only the second is `unreachable`.
    // Per CALL, beside the two latches, so a reader one layer out can tell
    // whether the empty string it just got was an answer or a failure. The crash
    // report is that reader, and without this it wrote "(nothing)" over a read
    // that never happened — twice, on the only two crashes of round 044.
    state.lastError = isOverflow(r.error) ? "overflow" : r.error ? "spawn" : null;
    if (isOverflow(r.error)) {
      state.overflowed = args[0];
    } else if (r.error) {
      state.unreachable = true;
      // WHICH call, and what the OS said. "playwright-cli could not be run" sent
      // two rounds' worth of debugging at an install that was fine, because the
      // message named the tool and the tool was never the problem. A spawn
      // failure has a subcommand and an errno and both were being thrown away.
      state.unreachableAt ??= {
        args: args.join(" ").slice(0, 80),
        error: String(r.error?.message ?? r.error),
        // WHICH OF THE TWO INPUTS WAS ABSENT — checked, not inferred from the
        // errno. A spawn ENOENT here has two causes, the entry and the working
        // directory, and the remedy for either is useless for the other. Read at
        // failure time because both can appear or vanish mid-run.
        entryMissing: !existsSync(entry),
        cwdMissing: !existsSync(cwd),
      };
    }
    // A failed CLI call and a page that answered with nothing are the same empty
    // string, and the difference decides whether a round is alive. Recorded, not
    // folded in — the same distinction `unreachable` exists to keep.
    state.lastFailed = !r.error && r.status !== 0;
    /**
     * WHAT THE FAILED CALL SAID, because throwing it away cost a cycle.
     *
     * A non-zero call returns "" here and the reason went nowhere. On
     * 2026-09-06 `open` failed on every attempt of a cycle with
     *
     *     Browser is already in use for C:/devtools/pw-profile
     *
     * and the driver reported `browser-gone` — "the process died" — then
     * "the file list shows no `Presentation64` to open". Both false: `list`
     * saw no browser while the profile was very much held, so every attempt
     * reopened nothing and blamed the file list, and the cycle only moved once
     * the Chrome holding the profile was ended by hand.
     *
     * WHY the browser stops answering while still running is NOT established —
     * a dead CLI daemon is the obvious guess and it is only a guess. This
     * field is about the sentence the tool printed, which is a fact whatever
     * the mechanism turns out to be.
     *
     * Kept only for the failing call, and only as text for a human to read.
     * Nothing branches on stdout being empty because of it.
     */
    state.lastStderr = r.status === 0 ? null : String(r.stderr ?? "");
    return r.status === 0 ? String(r.stdout ?? "") : "";
  };
  sh.state = state;
  // The session directory the CLI is rooted at, so callers can find what it
  // downloads without re-deriving it — `collectRound` needs the run log, and a
  // second `sessionDir()` call beside this one is a second place to get the 8.3
  // short-name normalisation wrong.
  sh.dir = cwd;
  /**
   * Start a fresh sweep, forgetting a spawn failure the last one saw.
   *
   * `unreachable` is deliberately sticky WITHIN a sweep: a call that never ran
   * and a page that answered with nothing are the same empty string, so once one
   * spawn has failed nothing that sweep read can be trusted. Across sweeps it is
   * a lie, and round 044 is what that costs.
   *
   * `--retry` survived the crash, `recover` reloaded the tab and reopened the
   * pane, and one of its ~8 spawns lost a race. The next attempt then measured a
   * perfectly healthy setup — `host answered in 4ms`, printed one line above —
   * and refused it with "playwright-cli could not be run — nothing below was
   * actually measured", which was false of every value on screen. A latch that
   * outlives its evidence turns a recoverable round into a stop, and this one
   * exited 0 while doing it.
   *
   * The same mistake as the poll that once ended a round on a single failed CLI
   * call (round 29), one call site further out. Fixed the same way: scope the
   * doubt to the thing it was actually about.
   */
  sh.startSweep = () => {
    state.unreachable = false;
    state.unreachableAt = undefined;
  };
  return sh;
}

/** A `ref_N` for the first element matching `pattern` in a `find` result. */
/**
 * Is the control THERE, whatever state it is in?
 *
 * `refFor` answers null for two situations that are not the same thing: no such
 * control, and a control that is present but DISABLED — a greyed-out button can
 * match the line and carry no ref.
 *
 * CAN, not does. "Playwright only issues a ref for something it could act on"
 * was the first version of this sentence and it is too strong: the add-in pane's
 * own `button "Use deck style" [disabled] [ref=f19e328]` has both, while the
 * PowerPoint ribbon's `button "Insert chart" [disabled]` has neither. Native
 * `disabled` and whatever Office marks its ribbon with are not treated alike.
 *
 * Which is the argument for asking the question directly rather than inferring
 * it: the rule to rely on is never "no ref means gone", in either direction.
 *
 * On 2026-08-19 that cost a round. PowerPoint's document went `Disconnected`
 * (the network again, the same fault that crashed round 089), which greys the
 * WHOLE ribbon. `canOpenPane` looked for `button "Insert chart"`, got no ref,
 * and concluded the add-in was gone — an un-retryable, owner-only refusal — for
 * a state a reload clears. The accessibility tree said `button "Insert chart"
 * [disabled]` the entire time, with the add-in loaded and fine.
 *
 * So: presence and usability are two questions, and the driver has to ask both.
 */
/**
 * `find`'s answer with its MISS MESSAGE removed.
 *
 * THE TRAP THIS CLOSES HAS CAUGHT ME TWICE IN ONE DAY. `playwright-cli find`
 * answers a miss with:
 *
 *     No matches found for "Insert chart".
 *
 * — which CONTAINS the query. So any pattern built from the bare name matches
 * the miss, and a probe for a control reports it present precisely when it is
 * absent. It is the worst possible polarity: the failure mode is a false
 * POSITIVE, on the exact reading a caller uses to decide something is fine.
 *
 * On 2026-08-20 it said a crash dialog was up when none was (harmless, caught
 * in minutes), and then said the add-in was back in the ribbon when it was not
 * — which I reported to the owner as confirmation of something they had told
 * me, so the bad reading did not even look like a machine's mistake.
 *
 * Stripped centrally rather than patched at each call site, because the two
 * call sites that got it right did so by accident of wording (`button "..."`
 * happens not to appear in the miss line), and the next caller has no reason
 * to know.
 */
export function findLines(sh, query) {
  return sh("find", query)
    .split("\n")
    .filter((l) => !/^\s*No matches found for /.test(l));
}

export function namePresent(sh, query, pattern) {
  return findLines(sh, query).some((l) => pattern.test(l));
}

/**
 * How long the ribbon may take to show a freshly sideloaded command.
 *
 * ONE SLEEP AND ONE LOOK IS NOT A WAIT, and on 2026-08-20 that cost a pair.
 * `sideloadAddIn` uploaded the manifest, slept 12s, checked for `Insert chart`
 * once, found nothing and returned "uploaded, but no SSF Charts command
 * appeared" — a permanent-sounding verdict. Round 117 refused, reporting that
 * only a person could put the add-in back. THE OWNER TOUCHED NOTHING AND THE
 * COMMAND APPEARED ANYWAY: `Insert chart` and `Insert element` were both in the
 * ribbon minutes later. The sideload had worked; the check was early.
 *
 * That is this project's house defect again — a NOT-YET reported as a NEVER —
 * and it is the most expensive version of it so far, because the verdict sends
 * a person to do a job the machine had already done.
 */
export const SIDELOAD_COMMAND_BUDGET_MS = 90_000;

/**
 * How long the ribbon may take to WAKE before its absence means anything.
 *
 * Different question from `SIDELOAD_COMMAND_BUDGET_MS`, hence a different
 * number. That one asks "did an upload land"; this one asks "is this tab awake
 * yet". An idle Office tab does not keep its ribbon in the accessibility tree,
 * and the first read after it wakes can arrive before the controls do.
 *
 * Short, because the answer is usually immediate and a genuinely missing add-in
 * should still be reported quickly.
 */
export const RIBBON_WAKE_BUDGET_MS = 20_000;

/**
 * Poll for a ref until it appears or the budget runs out.
 *
 * Returns the ref, or null once the budget is spent — and null here means
 * "still not there after N seconds", which is a different and weaker claim than
 * the single-look version it replaces. Callers must word it that way.
 *
 * `now` is injectable so a test can spend a budget without spending the time.
 */
export async function waitForRef(sh, sleep, query, pattern, budgetMs, every = 3000, now = Date.now) {
  const started = now();
  // BOUNDED TWO WAYS, and the second bound is not belt-and-braces.
  //
  // A wait bounded only by a clock assumes the clock moves, which assumes the
  // sleep really sleeps. Inject a stub that returns immediately — every test in
  // this file does — and `now()` never advances, so the loop spins as fast as
  // the process can allocate. It does not hang politely: it took Vitest to
  // `FATAL ERROR: JavaScript heap out of memory` in 80 seconds.
  //
  // Counting the attempts costs one variable and cannot be defeated by a clock
  // that stands still, which is exactly the condition under which a runaway
  // loop is hardest to notice.
  const maxAttempts = Math.max(1, Math.ceil(budgetMs / Math.max(1, every)) + 1);
  for (let attempt = 1; ; attempt++) {
    const ref = refFor(sh, query, pattern);
    if (ref) return ref;
    if (attempt >= maxAttempts || now() - started >= budgetMs) return null;
    await sleep(every);
  }
}

/**
 * The page width below which PowerPoint hides ribbon commands in an overflow.
 *
 * THE MOST EXPENSIVE MISREADING OF THIS WHOLE PROJECT, and it cost a day.
 * PowerPoint collapses trailing ribbon commands into a `...` menu when the
 * window is narrow, and a collapsed command IS NOT IN THE ACCESSIBILITY TREE.
 * So `namePresent(sh, "Insert chart", ...)` answers false — TRUTHFULLY, the
 * button is not rendered — and every reader above it concludes the add-in is
 * gone.
 *
 * Measured on 2026-08-21, same browser, same document, seconds apart:
 *
 *     page width 1237   Insert chart: false     "the add-in is not loaded here"
 *     page width 2375   Insert chart: true      round runs
 *
 * On that reading the driver ran the sideload walk FOUR TIMES against a deck
 * that already had the add-in, refused four rounds, and told the owner to
 * sideload by hand. Nothing was ever missing.
 *
 * It became reachable when the round browser moved to `viewport: null` so the
 * page would follow the window — which fixed the pane being invisible to a
 * PERSON and made commands invisible to the DRIVER. The window is not a
 * measurement surface, so it must not be allowed to vary.
 *
 * 1600 WAS BELOW A WINDOW THAT WAS GENUINELY CRAMPED, and the same false
 * "the add-in is gone" came back on 2026-08-22 because of it. Four measured
 * points now bracket the boundary:
 *
 *     1237   cramped   hid `Insert chart` (the 2026-08-21 incident)
 *     1996   cramped   hid it again, and 1996 >= 1600 so this said "wide
 *                      enough" and the driver reported the add-in MISSING
 *     2375   roomy     what brought the command back on 2026-08-21
 *     3800   roomy     what brought it back on 2026-08-22
 *
 * So the boundary is between 1996 and 2375, and the old threshold sat below
 * the highest width KNOWN to collapse. Chosen at the lowest width measured to
 * work rather than padded past it: a threshold above the evidence is a guess
 * in the safe direction, and this file has already paid for one guess.
 *
 * The number is in `window.innerWidth`, which on this machine is twice the
 * width `resize` is asked for — 1900 produced 3800. Do not compare it against
 * a figure taken from the resize call.
 *
 * IT WILL BE WRONG AGAIN ON ANOTHER MONITOR. A fixed pixel threshold is a
 * proxy for "can the ribbon show its commands", and the honest guard is that
 * being narrow is RECOVERABLE — `ensureRibbonRoom` widens and the next attempt
 * re-reads. What must never happen is the fatal `addin-missing` branch being
 * taken on a window that was never wide enough to judge.
 */
export const MIN_RIBBON_WIDTH = 2375;

/**
 * Is the window wide enough for an absent command to MEAN something?
 *
 * Null — the width could not be read — is NOT wide enough. A reader that cannot
 * measure the window cannot distinguish a missing add-in from a hidden one, and
 * this project's house defect is exactly that substitution.
 */
export function wideEnoughToJudge(width) {
  return typeof width === "number" && width >= MIN_RIBBON_WIDTH;
}

/** The page's own width in CSS pixels, or null if it could not be read. */
export function pageWidth(sh) {
  const out = sh("eval", "() => String(window.innerWidth)");
  const m = /(\d{3,5})/.exec(String(out ?? ""));
  return m ? Number(m[1]) : null;
}

/**
 * Make the window wide enough that the ribbon shows its commands.
 *
 * Returns the width finally measured, or null if it could not be read at all —
 * and NULL IS NOT A PASS. A caller that cannot measure the window cannot tell a
 * missing add-in from a hidden one, which is the whole failure this exists to
 * end.
 *
 * Widens rather than refuses, because the window is not something the round is
 * measuring: nothing in a round's result depends on how wide the tab is, so
 * correcting it silently costs nothing and asking a person to do it costs a
 * round.
 */
export async function ensureRibbonRoom(sh, sleep) {
  const before = pageWidth(sh);
  if (before !== null && before >= MIN_RIBBON_WIDTH) return before;
  // DERIVED FROM THE THRESHOLD, never a literal beside it. This asked for 1900
  // while `MIN_RIBBON_WIDTH` was 2375: **the widen could not reach the bar it
  // exists to clear.** 1900 was generous when the threshold was 1600, and
  // raising the threshold left the request behind — a pair of constants that
  // have to move together, one of which moved.
  //
  // It hid for weeks behind the device scale factor. On this machine a request
  // of 1900 usually ARRIVES as 3800 CSS px, comfortably over, so the rounds all
  // said "widened to 3800px" and nothing looked wrong. Round 181 reopened the
  // browser with `--fresh`, the scale came back 1:1, 1900 arrived as 1900, the
  // Add-ins command stayed in the overflow, and seven attempts in a row could
  // not reopen the pane. A round lost to a number that was too small for the
  // whole time it was too small.
  //
  // The margin is for the scale going the OTHER way — a request that arrives
  // smaller than asked — and the measurement below is what settles it either
  // way. Ask for plenty and MEASURE what arrived, which is what the previous
  // comment said and the previous number did not do.
  sh("resize", String(Math.round(MIN_RIBBON_WIDTH * 1.3)), "1000");
  await sleep(3000);
  const after = pageWidth(sh);
  if (after !== null && before !== null && after > before)
    console.log(`  the window was ${before}px and hides ribbon commands — widened to ${after}px`);
  // AND SAY SO WHEN IT DID NOT WORK. `after` was returned and never compared to
  // the threshold, so a widen that fell short was indistinguishable from one
  // that succeeded — the caller saw a number and read it as a fix. That is the
  // house defect in the one function whose entire job is to reach a number.
  if (after !== null && after < MIN_RIBBON_WIDTH)
    console.log(
      `  STILL ${after}px, under the ${MIN_RIBBON_WIDTH}px a ribbon command needs to render — ` +
        `expect "the add-in is missing" from a command that is merely collapsed`,
    );
  return after;
}

export function refFor(sh, query, pattern) {
  // Through `findLines`, so the miss message can never be matched as a hit —
  // see the note there. `refFor` was safe only because a ref-bearing pattern
  // has no ref to extract from the miss line, which is luck, not design.
  const line = findLines(sh, query).find((l) => pattern.test(l));
  // The ref on the MATCHING LINE, never the first ref in the output. `find`
  // prints the whole frame hierarchy above the hit, so the first ref belongs to
  // the outer iframe — evaluating against it lands in the OneDrive document
  // rather than in the pane, where `Office` and `PowerPoint` are both undefined
  // and a host ping silently reports a healthy host as dead.
  return line ? (/ref=([a-z0-9]+)/.exec(line)?.[1] ?? null) : null;
}

/**
 * Is the host answering AT ALL, before a round is offered?
 *
 * Rounds 24, 25 and 29 each cost most of an hour to discover that it was not
 * going to. Round 29 showed the host was already unwell before the probe's
 * fourth question, and a ping run afterwards showed the state persists: at
 * 17:54, nearly two hours after that round wedged and after a full tab reload,
 * `slides.getCount()` — the cheapest call Office.js has — did not come back
 * inside eight seconds.
 *
 * Detectable in seconds, and it survives a reload. Asking here turns a
 * sixty-six-minute wedge into a two-second refusal.
 *
 * Runs inside the PANE's frame, which is why the caller passes a ref belonging
 * to the pane: `PowerPoint` does not exist in the document frame around it.
 */
export function pingScript(budgetMs) {
  return (
    "async () => { const t = Date.now(); try { await Promise.race([ " +
    "PowerPoint.run(async (c) => { c.presentation.slides.getCount(); await c.sync(); }), " +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    'return "ok:" + (Date.now() - t); } catch (e) { return "no:" + (Date.now() - t); } }'
  );
}

/**
 * Touch a SLIDE before the round does, because the ping does not.
 *
 * THE EXPERIMENT, stated so a later reader can tell whether it worked.
 * PowerPoint crashed 2s into the FIRST attempt of rounds 043 and 044 — four for
 * four — and never into an attempt that followed a recovery, the difference
 * being about eighty seconds of reload and settling. Its own log says what it
 * was doing:
 *
 *     In OnDisconnect(), setting SlideViewNode.srcSlide to null
 *     Failed to restore selection after load content.
 *     OnServerFindSucceeded could not find target slide, time elapsed: 430 ms
 *     GlobalErrorHandler:DisplayErrorDialog: 5341289
 *
 * A document still settling when the round's first Office.js call lands. The
 * ping cannot see it: `slides.getCount()` is a COUNT, and this host answers it
 * in single-digit milliseconds while it is in exactly that state. Resolving a
 * slide is the cheapest call that goes down the path the host died on.
 *
 * The point is NOT to prevent the crash — it is to move it. If the theory holds,
 * this trips it here, where the check is two seconds and `--retry` recovers
 * before any round has been spent; today it costs a whole attempt plus the
 * recovery. If the theory is wrong this answers `ok` and the round crashes
 * anyway, which refutes it for the price of one extra call.
 *
 * `getItemAt(0)` and not the selection: a round starts on a one-slide deck, and
 * the selection API is itself one of the calls this host has wedged on
 * (`which selection call wedges the host`).
 */
export function slideResolveScript(budgetMs) {
  return (
    "async () => { try { return await Promise.race([ " +
    "PowerPoint.run(async (c) => { const s = c.presentation.slides.getItemAt(0); " +
    's.load("id"); await c.sync(); return "slide:" + (s.id || "?"); }), ' +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    '} catch (e) { return "slide-failed:" + (e && e.message ? String(e.message).slice(0, 80) : "?"); } }'
  );
}

/**
 * Ask the deck its slide size, so a round cannot lie about which one it ran.
 *
 * WHY THIS IS A PRECONDITION AND NOT A NICETY. Setting a deck to Widescreen
 * during the 2026-08-16 control run SILENTLY DID NOT TAKE — the click landed
 * while the document was in its greyed "Loading" state, the menu accepted it,
 * and nothing changed. It was caught only by reopening the menu and reading
 * which box was ticked.
 *
 * A round that believes it is 4:3 and is not proves exactly nothing, which is
 * the same harm as a round on a stale pane — and that is already a hard stop.
 * With a nightly cycle running 16:9 twice and 4:3 once, an unverified size means
 * filing a round under the wrong profile, which is worse than not running it.
 *
 * Points, via `pageSetup`, because that is what the add-in itself reads.
 */
/**
 * Set the deck's slide size, so a dedicated 4:3 deck can be made 4:3 unattended.
 *
 * `PowerPoint.PageSetup.slideWidth` and `slideHeight` are WRITABLE at
 * PowerPointApi 1.10, and round 096's `environment` line records this host
 * advertising 1.1 through 1.10. So the driver's `wrong-size` refusal — "set it
 * in Design ▸ Slide Size" — was asking a person to do something the API can do.
 *
 * OFF BY DEFAULT, AND THAT IS THE IMPORTANT HALF. Resizing the WRONG deck is a
 * quiet disaster: there is one 16:9 deck behind almost the whole archive,
 * `roundProfile` defaults to 16:9 for the 53 rounds that carry no size at all,
 * and `scenarioRegressions` compares within one profile — so a deck that
 * silently changed shape mid-series would split every comparison built on it.
 * A misaimed `PW_DECK` plus an automatic resize would do that to a real
 * presentation without anyone noticing.
 *
 * So it happens only when the owner asks for it with `PW_SET_SIZE=1`, and it
 * says so loudly when it does. The intended use is a deck that EXISTS to be 4:3,
 * where setting the size is idempotent and makes the deck what its name claims.
 */
export function setSlideSizeScript(size, budgetMs) {
  const [w, h] = size === "4:3" ? [720, 540] : [960, 540];
  return (
    "async () => { try { return await Promise.race([ " +
    "PowerPoint.run(async (c) => { const p = c.presentation.pageSetup; " +
    `p.slideWidth = ${w}; p.slideHeight = ${h}; await c.sync(); ` +
    'p.load("slideWidth,slideHeight"); await c.sync(); ' +
    'return "size:" + Math.round(p.slideWidth) + "x" + Math.round(p.slideHeight); }), ' +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    '} catch (e) { return "size-failed:" + (e && e.message ? String(e.message).slice(0, 60) : "?"); } }'
  );
}

export function slideSizeScript(budgetMs) {
  return (
    "async () => { try { return await Promise.race([ " +
    "PowerPoint.run(async (c) => { const p = c.presentation.pageSetup; " +
    'p.load("slideWidth,slideHeight"); await c.sync(); ' +
    'return "size:" + Math.round(p.slideWidth) + "x" + Math.round(p.slideHeight); }), ' +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    '} catch (e) { return "size-failed:" + (e && e.message ? String(e.message).slice(0, 60) : "?"); } }'
  );
}

/**
 * `"size:960x540"` → `"16:9"`; a failure or a silence → null.
 *
 * Null is NOT a mismatch. A host that would not answer has told us nothing, and
 * refusing a round on no evidence is the mistake `reachable` exists to prevent.
 */
export function readSlideSize(out) {
  const m = /size:(\d+)x(\d+)/.exec(String(out ?? ""));
  if (!m) return null;
  const w = Number(m[1]);
  const h = Number(m[2]);
  if (!w || !h) return null;
  const r = w / h;
  if (Math.abs(r - 16 / 9) < 0.01) return "16:9";
  if (Math.abs(r - 4 / 3) < 0.01) return "4:3";
  return `${w}x${h}`;
}

/** `"slide:287#62081387"` → true; a failure or a silence → false. */
export function readSlideResolve(out) {
  const s = String(out ?? "");
  if (/slide-failed/.test(s)) return false;
  return /slide:/.test(s) ? true : null;
}

/**
 * Is the browser sitting on a Microsoft sign-in page?
 *
 * Told apart from "the pane is closed" because the fix is completely different
 * and only one of them is the agent's to do. A signed-out browser answers every
 * pane read with nothing, which reads exactly like an add-in nobody opened —
 * and the check duly said "is the add-in open?" while the tab showed
 * `login.live.com`, sending the reader to look for a pane in a window that has
 * no document in it.
 *
 * Same distinction `reachable` exists for, a layer further out: there is a
 * difference between a question that was answered with nothing and a question
 * there was nobody to ask.
 */
export function signedOut(tabList) {
  return AUTH_TAB.test(String(tabList ?? ""));
}

const AUTH_TAB = /login\.live\.com|login\.microsoftonline\.com|login\.windows\.net/;

/**
 * Is that sign-in tab a POPUP beside a live deck, rather than the whole browser
 * sitting on a login page?
 *
 * `signedOut` is right to stop the round either way — if Office is asking for
 * credentials, nothing measured after it can be trusted — but the two look
 * completely different to the person who walks up to the screen, and until
 * 2026-08-16 they read the same sentence. What ended the overnight run of
 * 2026-08-15/16 was an auth popup that opened BESIDE a deck tab that was still
 * open, and the driver reported "the browser is on a Microsoft sign-in page".
 * The reader looked at a screen showing PowerPoint and a small dialog, and had
 * to work out for themselves that the driver was describing the dialog.
 *
 * A message that does not match what is on the screen is worse than a vague one:
 * it makes a reader doubt the whole report, which on an overnight run is the only
 * account of what happened.
 *
 * The signal is that some OTHER tab is still a document. Deliberately loose about
 * which — the driver's own recovery looks for `Presentation63` by name, and
 * hard-coding one deck's name into a diagnostic is how it would go quietly wrong
 * for the next deck.
 */
export function signInIsPopup(tabList) {
  const lines = String(tabList ?? "")
    .split("\n")
    .filter((l) => l.trim());
  return (
    lines.some((l) => AUTH_TAB.test(l)) &&
    lines.some((l) => !AUTH_TAB.test(l) && /officeapps\.live\.com|onedrive\.live\.com|sharepoint\.com|\.pptx/i.test(l))
  );
}

/**
 * Is there a BROWSER at all?
 *
 * The third thing that answers every pane read with nothing, after "the CLI
 * could not be run" and "the browser is signed out" — and the one that had no
 * name until it cost seven attempts on 2026-08-15. A round wedged, the browser
 * process died with it, and `recover` then reloaded and reopened a pane in a
 * window that did not exist, seven times, while the check reported "could not
 * read the pane's build stamp — is the add-in open?" The add-in was fine. There
 * was nothing to open it in.
 *
 * Worth its own precondition because the fix is specific and the loop can do it
 * unattended: reopen from the persistent profile, which still holds the
 * sign-in — **a dead browser is not a lost sign-in**, and believing otherwise
 * has now cost this project two separate stretches of hours.
 *
 * `pw list` answers `(no browsers)` in exactly this state, and that string is
 * what the daemon prints when it has no session for the working directory.
 */
/**
 * Where the browser that survives a session lives.
 *
 * `scripts/pw.sh` parks the SESSION in the repo and the PROFILE here, and the
 * profile is the half that holds the OneDrive sign-in across a browser death.
 */
export const PROFILE_DIR = process.env.PW_PROFILE_DIR ?? "C:/devtools/pw-profile";

/**
 * The deck a recovery reopens, when the browser has died and taken its tab.
 *
 * A DEFAULT rather than a constant, and overridable with `PW_DECK`, because the
 * deck's name changes and the old hard-coded `Presentation63` was already stale
 * on 2026-08-16 — the deck in use had become `Presentation64`. A sideload on
 * PowerPoint for the web is **per-document**, so a fresh deck is a fresh
 * sideload and a fresh name, and this will drift again.
 *
 * It fails in the worst possible way when wrong: `recover` reopens OneDrive,
 * finds no matching link, clicks nothing, and reports a closed pane — in exactly
 * the situation the function exists to rescue.
 */
export const DECK_NAME = process.env.PW_DECK ?? "Presentation64";

export function noBrowser(listOutput) {
  return /\(no browsers\)/i.test(String(listOutput ?? ""));
}

/**
 * Did `open` refuse because something this session cannot reach holds the
 * profile?
 *
 * The third state between "a browser this session can drive" and "no browser
 * at all": Chrome still running, `list` reporting `(no browsers)`, and the
 * profile directory locked by a process nothing here can address. `noBrowser`
 * cannot see it and `close-all` cannot end it.
 *
 * NAMED FOR THE SYMPTOM, not the cause. "The daemon died" is the obvious
 * explanation and stays an inference — it has not been caught happening, and a
 * function named after an unproven mechanism is a claim every caller then
 * repeats.
 *
 * Matched on the profile phrase rather than the `--isolated` advice beside it,
 * because isolated is not a remedy here and a message that changes its
 * suggestion should not change this answer. Anything that is not a string is
 * false: a call that never ran left no stderr, and an absence must never read
 * as a diagnosis.
 */
export function profileHeldByOrphan(stderr) {
  return /browser is already in use for/i.test(String(stderr ?? ""));
}

/**
 * What the pane is showing, for a stop that otherwise hands a person nothing.
 *
 * `no-run-button` ended three cycles on 2026-09-06/07, each after a browser
 * death and a SUCCESSFUL re-sideload: the log shows the pane back, the
 * Automation tab selected, the deck swept, `ready`, and then the refusal. The
 * mitigation in `attempt` — select the tab and look again — had already run and
 * already failed, so the only useful next question is what the pane actually
 * had on it, and that answer was discarded all three times.
 *
 * TABS AND BUTTONS ONLY, and capped. Enough to separate "on the wrong tab" from
 * "rendered without its controls" from "not a pane at all" — three different
 * next steps that were one message. A full accessibility dump would bury the
 * stop it is attached to.
 *
 * Anything unparseable reads as nothing rather than throwing: this runs on a
 * path that has already failed, and a diagnostic that can end a round is worse
 * than no diagnostic.
 */
export function describePane(findOutput, limit = 12) {
  const lines = String(findOutput ?? "")
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => /^- (tab|button) "/.test(l))
    .slice(0, limit);
  return lines.length ? `  the pane is showing:\n    ${lines.join("\n    ")}` : "  the pane answered nothing at all";
}

/**
 * Did the pane describe NOTHING — no tab, no button, nothing at all?
 *
 * The same parse as `describePane`, asked as a question, because the answer
 * decides whether a round can be recovered or the night ends.
 *
 * WHY IT EXISTS. `no-run-button` has ended five cycles, every one of them AFTER
 * a successful recovery: the pane reopens, the add-in re-sideloads, the deck is
 * swept, the driver prints `ready`, and then the button is not there. It has
 * been a terminal stop that whole time, on the theory that a pane sitting on
 * the wrong tab needs a person to look at it.
 *
 * On 2026-09-08 the diagnostic finally printed and the pane was not on the
 * wrong tab — it answered nothing at all. No tabs, no buttons, seconds after
 * `host answered in 3ms · slide 1 resolved`. That is not a pane wanting a
 * click; it is a pane that has gone, which `RECOVERABLE_STOPS` already handles
 * as `pane-closed`. Recovery reopens it and the attempt loop tries again,
 * instead of a night ending on attempt 3 of 7.
 *
 * KEPT SEPARATE FROM `describePane` rather than folded into it. That one
 * formats for a person and is called inside a swallowing try/catch, because a
 * diagnostic must never end a round. This one CHANGES CONTROL FLOW, so it has
 * to be callable and testable on its own and must not be wrapped in the same
 * swallow.
 */
export function paneAnsweredNothing(findOutput) {
  return !String(findOutput ?? "")
    .split("\n")
    .some((l) => /^\s*- (tab|button) "/.test(l));
}

/**
 * The command that ends an orphaned browser, for the platform this runs on.
 *
 * A FUNCTION OF THE PLATFORM RATHER THAN A READ OF `process.platform`, for the
 * reason `is-main.mjs` gives at length: rounds run on Windows and CI runs on
 * ubuntu, so a helper that could only be exercised on the platform under it
 * would go green against the bug it exists to catch.
 *
 * Matched on `--user-data-dir=<profile>`, which is the only thing that
 * separates this driver's browser from the operator's own Chrome. Both are
 * `chrome.exe`; matching the NAME would close the window someone is reading
 * this in.
 *
 * AND THE SPAWN ITSELF WAS CHECKED, not assumed. This file already carries a
 * long note about AppLocker closing every route through a shell — the
 * `playwright-cli` shim is blocked with "This program is blocked by group
 * policy" — so a new helper that shells out is exactly the kind of thing that
 * would be written, merged, and then quietly return false forever. Run against
 * a profile path matching nothing on the owner's machine on 2026-09-06 it
 * spawned and returned in 648ms, killing nothing.
 */
export function endOrphanCommand(profile, platform = process.platform) {
  if (platform === "win32")
    return [
      "powershell",
      [
        "-NoProfile",
        "-Command",
        `Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" | Where-Object { $_.CommandLine -like '*${profile}*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }`,
      ],
    ];
  return ["pkill", ["-f", `user-data-dir=${profile}`]];
}

/**
 * End a browser that holds the profile and answers to no daemon.
 *
 * ONLY EVER CALLED AFTER `open` HAS JUST REFUSED, and that guard is the whole
 * licence for it. The driver already closes browsers — `--fresh` does it every
 * leg, `close-all` does it in the branch above — so ending one it launched is
 * not a new authority. What would be new is ending one speculatively, so this
 * cannot be reached except from a refusal that names this exact condition.
 *
 * WHY IT IS WORTH DOING AT ALL. This was written as a message and left for a
 * person, and then the state recurred twice in one evening: the browser died
 * 746s into round 415, its daemon with it, and thirteen Chrome processes went
 * on holding `pw-profile` with nothing able to address them. A message is the
 * right answer for a condition an operator will see; this one arrives at 3am in
 * the middle of an unattended cycle, and every retry after it is wasted.
 *
 * Returns whether the command ran, not whether it worked — the caller finds
 * that out by trying `open` again, which is a fact rather than an inference.
 */
export function endOrphanedBrowser(profile, run = spawnSync, platform = process.platform) {
  const [cmd, args] = endOrphanCommand(profile, platform);
  try {
    const r = run(cmd, args, { encoding: "utf8", timeout: 30_000 });
    return !r?.error;
  } catch {
    return false;
  }
}

/**
 * Can THIS driver actually use a browser — not, is there one somewhere.
 *
 * `list` enumerates every browser the daemon knows about, whatever session it
 * belongs to. Every other command is SESSION-KEYED, and the daemon keys its
 * sessions by the working-directory STRING. So a browser opened from a shell
 * whose cwd was `C:/devtools/SSF-Charts/.pw-session` is invisible to a driver
 * whose cwd is `C:\devtools\SSF Charts\.pw-session` — same folder, different
 * string, different session. `sessionDir` normalises the driver's own path and
 * cannot do anything about a browser opened from outside it.
 *
 * On 2026-08-22 that produced a driver that could SEE a browser and not talk to
 * it: `list` succeeded while `tab-list`, `find` and every read exited non-zero.
 * `noBrowser` answered false, so `recover` never opened one it could use, and
 * seven attempts ran against a browser it was structurally unable to reach.
 *
 * Asked by making a session-keyed call and seeing whether it worked, because
 * that is the capability every later read depends on. An empty answer counts as
 * unreachable: `tab-list` on a live session always names at least one tab.
 */
export function browserReachable(sh) {
  const out = sh("tab-list");
  if (sh.state?.lastFailed || sh.state?.lastError) return false;
  return String(out ?? "").trim().length > 0;
}

/**
 * Did `find` actually FIND it, or is it echoing the query back?
 *
 * `playwright-cli find` answers a miss with `No matches found for "<query>"` —
 * which contains the query. Testing the output for the phrase searched for is
 * therefore always true, and a crash detector built that way would report a
 * crash on a perfectly healthy host, every time. Same family as the ref that
 * `buildOf` used to read out of its own haystack.
 */
/**
 * How long to keep watching the pane after PowerPoint's crash dialog appears.
 *
 * The pane's own budget is `DECK_EVIDENCE_TIMEOUT_MS` — 45 seconds — and its
 * timer starts when the SCAN starts, which is some seconds before the dialog
 * shows. 90 gives that a clear margin plus the local work that follows it, and
 * still costs a twentieth of the wedge budget it is spent against.
 *
 * Env-overridable for the same reason every other budget here is: the one thing
 * that would make it wrong is a host slower than any yet seen.
 */
export const PANE_FINISH_GRACE_MS = Number(process.env.PW_PANE_FINISH_GRACE_MS) || 90_000;

/**
 * Did the pane finish the round anyway, after the host died under it?
 *
 * Polls the one signal that means the round is complete and downloadable — the
 * same enabled-button test the main loop breaks on — and answers false when the
 * grace runs out. See the call site for why this exists at all.
 *
 * READS ONLY THE PANE, never the host: `find` is a DOM read against an iframe
 * that outlived Office.js, so it works precisely when nothing else does. An
 * Office.js call here would hang on the dead host and spend the grace waiting
 * for a single answer.
 *
 * A poll that FAILS TO RUN is not a "no" — `sh.state.lastFailed` says the call
 * never happened, and folding that into "the pane is gone" is the mistake the
 * main loop's `dlFailed` comment describes at length. Those polls are skipped,
 * so a busy CLI costs patience rather than a round.
 */
export async function paneFinishedAnyway(sh, sleep, started, graceMs = PANE_FINISH_GRACE_MS) {
  const until = Date.now() + graceMs;
  console.log(
    `  the round's verdicts are banked before the deck scan — watching the pane for ${Math.round(graceMs / 1000)}s ` +
      "in case it finishes without the host",
  );
  while (Date.now() < until) {
    await sleep(3000);
    const dl = sh("find", "Download run log");
    if (sh.state.lastFailed) continue;
    if (/button "Download run log"(?! \[disabled\])/.test(dl)) {
      const secs = Math.round((Date.now() - started) / 1000);
      console.log(
        `  THE PANE FINISHED ANYWAY, ${secs}s in — the host died in the deck scan and the round survived it. ` +
          "Filing it; its deck evidence will be absent and the round says so.",
      );
      return true;
    }
  }
  console.error(`  the pane did not finish inside ${Math.round(graceMs / 1000)}s — this round is lost`);
  return false;
}

export function sawCrashDialog(found) {
  const text = String(found ?? "").trim();
  // Deliberately NOT "does the output contain the word dialog". That version
  // passed its own test while proving nothing — the phrase searched for happens
  // not to contain the word, so the echo could never have tripped it and the
  // guard was decoration. Whether `find` matched at all is the real question,
  // and it stays true whatever the next query says. Empty is the third answer:
  // the CLI was never reached, which `reachable` reports on its own.
  return text !== "" && !/No matches found/.test(text);
}

/**
 * How many polls in a row the PAGE has answered with nothing.
 *
 * A CLI call that failed is not a page that said nothing, and only the second
 * kind means the round is over. `failed` resets rather than counts: a tool that
 * could not be run measured nothing, and treating it as evidence is the mistake
 * `reachable` was written to stop, arriving here by a different door.
 */
export function quietStreak(prev, text, failed) {
  return String(text ?? "").trim() || failed ? 0 : prev + 1;
}

/** `{ answered, ms }` from what `pingScript` returned, or null when unreadable. */
export function readPing(out) {
  const m = /"?(ok|no):(\d+)"?/.exec(String(out ?? ""));
  return m ? { answered: m[1] === "ok", ms: Number(m[2]) } : null;
}

/** Click through the element's own handler — a plain click does not reach the pane. */
function clickRef(sh, ref) {
  // The pane sits two iframes deep and `playwright-cli click` silently does
  // nothing there: `aria-selected` never moves and no error is raised. Firing
  // the element's own handler is the only thing that works, and finding that
  // out cost a round.
  return sh("eval", "el => { el.click(); return 'ok'; }", ref);
}

export async function attempt(argv, deps, sh, healed = false) {
  const run = deps.run ?? spawnSync;
  const sleep = deps.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  const fetchBuild = deps.fetchBuild ?? defaultFetchBuild;
  const checkOnly = argv.includes("--check");
  // This sweep's reachability is about THIS sweep. See `sh.startSweep`.
  sh.startSweep?.();

  // BEFORE ANYTHING IS MEASURED, because every reading below is about whichever
  // tab is fronted. Only when a deck was actually asked for: with no `PW_DECK`
  // this is the behaviour it has always had, running against the open document.
  const wantDeck = process.env.PW_DECK || null;
  const deckFronted = wantDeck ? selectDeck(sh, wantDeck) : true;

  // PINNED BY THE CALLER FOR THE WHOLE ROUND, not re-read per attempt. A round
  // tests ONE build; re-deriving its identity mid-run lets the thing under test
  // change while it is being measured. Concretely: a commit made in this clone
  // while a round was retrying moved HEAD ahead of the deployed site, and the
  // driver then refused its own round as `site-behind` against a commit that had
  // never been deployed and was not what it was testing. Happened twice on
  // 2026-08-21. Falls back to reading it, for callers that drive `attempt`
  // directly.
  const head =
    deps.head ??
    (String(run("git", ["rev-parse", "--short=7", "HEAD"], { encoding: "utf8" }).stdout ?? "").trim() || null);
  const deployed = buildOf(await fetchBuild());
  let stamp = buildOf(sh("find", "--regex", "/[0-9a-f]{7} ·/"));
  // CAPTURED AGAINST THE CALL THAT PRODUCED IT, exactly as the poll loop in
  // `collectRound` does for `dl`. `sh.state` belongs to the MOST RECENT call, so
  // reading it any later judges this find's emptiness against a different call.
  const stampFailed = !!(sh.state.lastFailed || sh.state.lastError);

  const listRef = refFor(sh, "Slide List", /listbox "Slide List"/);
  const listFailed = !!(sh.state.lastFailed || sh.state.lastError);
  const slides = listRef ? (sh("snapshot", listRef).match(/option "Slide"/g) ?? []).length : null;
  // A READ THAT COULD NOT RUN IS NOT AN ABSENCE.
  //
  // `spawnSync` on this machine intermittently answers ENOENT for a node.exe
  // that is plainly there — eight consecutive calls succeeded minutes later —
  // and a failed spawn returns "". Every reading below is taken from that empty
  // string, so on 2026-08-22 a transient failure was reported as `pane ?`,
  // `deck ? slide(s)`, "could not read the pane's build stamp", and finally
  // "the add-in is gone from this document". None of it was true: the deck was
  // open, the stamp was findable and `Insert chart` was in the ribbon the whole
  // time. It cost an evening.
  //
  // This is the house defect in its purest form — a measurement that did not
  // happen, reported as a measurement that came back negative.
  const readsFailed = stampFailed || listFailed;

  // UNREACHABLE COUNTS AS GONE, for readiness's purposes. A browser this
  // session cannot issue a read against is worth exactly as much to the round
  // as no browser at all, and `recover` clears both the same way — see
  // `browserReachable`. Reported under the same stop so the driver retries
  // rather than reporting the absences the unusable browser would produce.
  const browserGone = noBrowser(sh("list")) || !browserReachable(sh);
  // Read ONCE. Two `tab-list` calls could straddle the popup opening or closing
  // and produce a message describing neither state.
  const tabs = sh("tab-list");
  const loggedOut = signedOut(tabs);
  const authPopup = loggedOut && signInIsPopup(tabs);
  const crashed = sawCrashDialog(sh("find", "Sorry, we ran into a problem"));

  // A tab, not the Verbose trace checkbox. The checkbox only exists while the
  // pane is ON the Automation tab, so anchoring the ping there skipped it
  // silently on a pane sitting anywhere else — and a skipped ping reads as
  // `ready`, which is the one answer this must never give without asking.
  // BEFORE ANY RIBBON READ. A narrow window hides commands in an overflow, and
  // a hidden command is indistinguishable from an absent one — see
  // `MIN_RIBBON_WIDTH` for the day that cost.
  const ribbonRoom = await ensureRibbonRoom(sh, sleep);
  const paneRef = refFor(sh, "Chart", /tab "Chart"/);
  // CAN THE PANE BE OPENED AT ALL, if it is shut? `recover` reopens it from the
  // ribbon's `Insert chart` control, so a document that does not offer that
  // control is one recovery cannot help — and retrying it is pure waste.
  //
  // Measured on 2026-08-16: a 4:3 leg switched to a deck the add-in was not
  // registered for and spent SEVEN attempts, about fifteen minutes, rediscovering
  // that a pane it cannot open is not openable. Only asked when the pane is
  // actually shut, so a healthy round pays nothing for it.
  let canOpenPane = paneRef ? true : Boolean(refFor(sh, "Insert chart", /button "Insert chart"/));
  // THE SAME CONTROL, ASKED THE OTHER WAY. `canOpenPane` is "could I click it";
  // this is "is it there at all". A `Disconnected` document greys the whole
  // ribbon, so the first is false and the second is true — and reading only the
  // first turned a transient network state into `addin-missing`, which recovery
  // is forbidden to retry. Cost the round on 2026-08-19.
  // PATIENT, BECAUSE THIS DECIDES WHETHER TO SIDELOAD.
  //
  // A single look here is the same shape as the defect inside `sideloadAddIn`,
  // one layer up: this reading decides whether to walk the sideload dialog at
  // all, so a premature miss costs a round.
  //
  // AN EARLIER VERSION OF THIS COMMENT CLAIMED THE RIBBON HAD BEEN FINE AND
  // THE DRIVER IMPATIENT. That was wrong, and it was wrong because of the trap
  // documented on `findLines`: the probe I checked it with matched `find`'s own
  // miss message, so it reported `Insert chart` present at the very moment it
  // was absent. The driver had been right the whole time. Corrected before the
  // false version could be read as history.
  //
  // The patience is KEPT, on its own merits rather than that story: an idle
  // Office tab does not keep its ribbon in the accessibility tree, one look
  // cannot tell "not there" from "not there YET", and the asymmetry favours
  // waiting — a few seconds lost against a genuinely missing add-in, versus a
  // wasted sideload walk and a refused round.
  //
  // Short budget on purpose: this is "is the ribbon awake yet", not "did an
  // upload land", and the latter has its own, much longer wait.
  const commandPresent =
    Boolean(paneRef) ||
    canOpenPane ||
    Boolean(await waitForRef(sh, sleep, "Insert chart", /button "Insert chart"/, RIBBON_WAKE_BUDGET_MS, 4000));
  // PUT IT BACK, ONCE, rather than refuse a night over it. This is the state a
  // browser death leaves: the deck is up and readable and there is no add-in in
  // it, because a web sideload does not survive the process. `recover` restores
  // everything except that.
  //
  // Guarded on `slides` for the same reason the refusal below is: a tab that is
  // merely mid-reload answers nothing to every read and looks identical, and
  // walking ten ribbon steps against a loading document would leave a dialog
  // over it.
  let pane = paneRef;
  // NOT WHILE THE RIBBON IS GREYED OUT. The sideload walk needs to click the
  // Add-ins button, and on a disconnected document that button is disabled like
  // every other — so the walk fails at its first step and reports "no Add-ins
  // button in the ribbon", which reads as Microsoft having moved something. It
  // did exactly that on 2026-08-19, burning the one attempt the latch allows on
  // a document that simply had to reconnect.
  // AND NOT ON A CRAMPED RIBBON. The walk ran four times on 2026-08-21 against
  // a deck that already had the add-in, because a narrow window hid the command
  // it was looking for. A sideload is the most expensive thing this driver can
  // do to a document; it must never be triggered by a reading it cannot trust.
  //
  // TWO STATES, ONE WALK. The add-in can be ABSENT — a browser death took the
  // sideload with it, which is the case above — or PRESENT AND STALE, its
  // command still in the ribbon and its pane pinned to a URL that no longer
  // answers. The second is what a domain move leaves behind, and it looked
  // identical to a transient disconnect until it had been seen twice.
  //
  // The stale case deliberately does NOT require `!canOpenPane`: the command is
  // right there, which is precisely why the absent-case guard could not see it.
  // What it requires instead is a SECOND sighting, because that is the reading
  // the first one cannot give — a disconnected document reconnects, a dead host
  // does not. Both still share the one-per-process latch and the cramped-ribbon
  // guard, so this cannot spend more sideloads than before.
  if (!pane && commandPresent && slides !== null) commandWithoutPane++;
  const missing = needsSideload({ pane, canOpenPane, commandPresent, sightings: commandWithoutPane });
  if (missing && slides !== null && !sideloadAttempted && wideEnoughToJudge(ribbonRoom)) {
    if (missing === "stale")
      console.log("  the ribbon command opens a pane that will not load — re-sideloading, its URL may be stale");
    sideloadAttempted = true;
    if (await sideloadAddIn(sh, sleep)) {
      // RE-READ, because everything below was measured before the add-in
      // existed. The stamp especially: it is what says the pane is serving the
      // build under test, and a sideload that lands the wrong manifest must
      // still be caught by the ordinary check rather than assumed good.
      pane = refFor(sh, "Chart", /tab "Chart"/);
      canOpenPane = Boolean(pane) || Boolean(refFor(sh, "Insert chart", /button "Insert chart"/));
      stamp = buildOf(sh("find", "--regex", "/[0-9a-f]{7} ·/"));
      console.log(`  the add-in is back — pane ${stamp ?? "still not readable"}`);
    }
  }
  // WHICH DOCUMENT this check is about. Read from the tab list, so it costs
  // nothing and works even when the host is silent — the case where knowing
  // what you are looking at matters most.
  const fronted = frontedDeck(sh);
  const ping = pane ? readPing(sh("eval", pingScript(8000), pane)) : null;
  // Only when the host is already answering: a slide resolve on a host that did
  // not survive `getCount` tells us nothing the ping has not, and costs 20s.
  const slideOk = pane && ping?.answered ? readSlideResolve(sh("eval", slideResolveScript(20000), pane)) : null;
  // `PW_EXPECT_SIZE=4:3` is how a nightly cycle says which arm this round
  // belongs to.
  const expectSize = process.env.PW_EXPECT_SIZE || null;
  // MEASURED EVERY CHECK, not only when a profile was asked for.
  //
  // This used to read the size ONLY when `PW_EXPECT_SIZE` was set, on the
  // reasoning that an unasked question costs a round trip and answers nothing.
  // That reasoning was wrong in the way that matters: AN INSTRUMENT THAT ONLY
  // ANSWERS WHEN YOU TELL IT THE ANSWER CAN NEVER SURPRISE YOU. The dedicated
  // 4:3 deck sat at 960x540 — a second 16:9 deck — from the day it was made
  // until 2026-08-20, and every `--check` in between was silent about it,
  // because nobody had run one with `PW_EXPECT_SIZE` set. The docs recorded the
  // 4:3 leg as "blocked on owner setup" the whole time. It was not blocked; it
  // was unmeasured, and the two are indistinguishable from the outside.
  //
  // The expectation now controls the COMPARISON only. The read is unconditional.
  let size = pane && ping?.answered ? readSlideSize(sh("eval", slideSizeScript(15000), pane)) : null;
  // MAKE THE DECK WHAT THE LEG ASKED FOR, when the owner has opted in. See
  // `setSlideSizeScript` for why this is off by default: resizing the wrong deck
  // splits every comparison the archive rests on, and a misaimed `PW_DECK` would
  // do it to a real presentation silently.
  //
  // Only when the size is KNOWN and WRONG. A host that would not answer has said
  // nothing, and resizing on no evidence is the mistake `reachable` exists to
  // prevent — the same reason the refusal below needs both values.
  if (expectSize && size && size !== expectSize && process.env.PW_SET_SIZE && pane) {
    console.log(`  the deck is ${size} and this leg wants ${expectSize} — setting it (PW_SET_SIZE is on)`);
    const after = readSlideSize(sh("eval", setSlideSizeScript(expectSize, 20000), pane));
    console.log(`  slide size is now ${after ?? "unreadable"}`);
    // READ THE ANSWER BACK rather than assuming the write took. The driver's own
    // refusal text has warned about this for months — "CHECK IT TOOK, because a
    // click made while the document is loading is accepted and does nothing" —
    // and an API write deserves the same suspicion.
    size = after ?? size;
  }
  // RE-READ, after the slide touch rather than only before it. If the touch is
  // what trips the host, the dialog appears in the seconds that follow — and
  // reading the dialog only at the top of the sweep is how a crash the check
  // itself provoked would be carried into the round as `ready`.
  const crashedAfter = slideOk === false ? sawCrashDialog(sh("find", "Sorry, we ran into a problem")) : false;

  // THE AUTOMATION TAB HAS TO BE SHOWING BEFORE ANY OF THIS IS READ. `Verbose
  // trace`, `Picture every slide` and the run button all live on it, and a pane
  // does not open there — a freshly sideloaded or reopened one comes up on
  // `Chart`, where none of the three is in the DOM.
  //
  // 2026-08-20 is what this is for: the driver put the add-in back by itself for
  // the first time, reached `ready` with `verbose trace ?` in the same block,
  // and then stopped because the run button was not there either. Both readings
  // had the same cause and neither was acted on.
  //
  // Selected HERE rather than just before the click, so the two toggles are read
  // rather than guessed. A round that starts with verbose trace off produces a
  // trace too thin to mine, and `?` is not evidence that it is on.
  if (pane) {
    const autoTab = refFor(sh, "Automation", /tab "Automation"/);
    const already = /tab "Automation" \[selected\]/.test(sh("find", "Automation"));
    if (autoTab && !already) {
      console.log("  the pane is not on the Automation tab — selecting it");
      clickRef(sh, autoTab);
    }
  }
  const toggles = sh("find", "Verbose trace");
  const verbose = /checkbox "Verbose trace"/.test(toggles) ? /Verbose trace" \[checked\]/.test(toggles) : null;
  const pictures = /checkbox "Picture every slide"/.test(toggles)
    ? /Picture every slide" \[checked\]/.test(toggles)
    : null;

  const state = {
    head,
    deployed,
    stamp,
    slides,
    verbose,
    pictures,
    reachable: !sh.state.unreachable,
    unreachableAt: sh.state.unreachableAt ?? null,
    browserGone,
    ping,
    slideOk,
    crashed: crashed || crashedAfter,
    loggedOut,
    authPopup,
    size,
    expectSize,
    wantDeck,
    deckFronted,
    canOpenPane,
    commandPresent,
    ribbonRoom,
    readsFailed,
  };
  const { ok, stop, codes, warn } = readiness({
    ...state,
    sessionIndex: deps.driverRun?.sessionIndex ?? 1,
    allowDeepSession: deps.allowDeepSession ?? false,
  });
  console.log(
    `  HEAD ${head ?? "?"} · site ${deployed ?? "?"} · pane ${stamp ?? "?"} · deck ${slides ?? "?"} slide(s)` +
      // WHICH DOCUMENT. Every line above described the round without ever
      // saying what it was running against, so a check against the wrong deck
      // read exactly like a check against the right one.
      (fronted ? ` · ${fronted}` : ""),
  );
  console.log(
    `  verbose trace ${verbose ?? "?"} · picture every slide ${pictures ?? "?"}` +
      // PRINTED EVERY CHECK. It used to appear only when `PW_EXPECT_SIZE` was
      // set, which meant the one state worth catching — a deck that is not the
      // profile you believe it is — was invisible unless you already suspected
      // it. `?` when the host did not answer: unreadable is not a size.
      ` · slide size ${size ?? "?"}` +
      (expectSize ? ` (want ${expectSize})` : ""),
  );
  console.log(
    `  host ${ping ? (ping.answered ? `answered in ${ping.ms}ms` : `SILENT for ${ping.ms}ms`) : "not asked — the pane is closed"}` +
      // Printed every round, pass or fail, because the experiment needs the
      // rounds where it says `resolved` as much as the ones where it does not.
      (slideOk === null ? "" : slideOk ? " · slide 1 resolved" : " · slide 1 REFUSED") +
      (state.crashed ? " · PowerPoint's crash dialog is up" : ""),
  );
  // HERE, and only here, because the line above is the proof. A host that just
  // answered has a valid session, so any sign-in popup still open belongs to a
  // flow that finished — see `closeStaleAuthPopups` for the other three gates.
  const shut = closeStaleAuthPopups(sh, { hostAnswered: !!ping?.answered });
  if (shut.length)
    console.log(`  closed ${shut.length} sign-in popup(s) left over from a finished flow — the host is answering`);
  // SWEEP IT RATHER THAN REFUSE, when the deck is the only thing wrong. See
  // `onlyDirtyDeck` — this is the one stop the driver can clear better than a
  // person, and the person doing it by hand is what emptied a deck entirely.
  if (!ok && onlyDirtyDeck(codes) && !healed && sweepDeck(sh)) {
    // ONCE, never in a loop. The re-check reads the deck through the same call,
    // so a sweep that did not take refuses exactly as it would have — a failed
    // heal cannot read as a successful one.
    console.log("  the deck still holds the last round's slides — sweeping it rather than refusing");
    await sleep(8000);
    return attempt(argv, deps, sh, true);
  }
  // BEFORE the verdict, so it is read whether or not the round goes ahead. A
  // deep round is not refused — it is labelled.
  for (const w of warn ?? []) console.log(`\n  DEEP IN A SESSION — ${w}`);
  if (!ok) {
    console.error("\n  NOT READY — a round now would not prove anything:");
    for (const s of stop) console.error(`    - ${s}`);
    return { code: 1, reason: state.crashed ? "crashed" : "not-ready", codes };
  }
  console.log("  ready");
  if (checkOnly) return { code: 0, reason: "checked" };

  // BEFORE THIS ROUND OVERWRITES THE LAST ONE'S REMAINS. See `keepCrashedRun`.
  await keepCrashedRun(sh, sleep);

  // THE RUN BUTTON ONLY EXISTS ON THE AUTOMATION TAB, and a pane does not open
  // there. A freshly sideloaded or reopened pane comes up on `Chart`, so the
  // button is not merely hidden — it is not in the DOM at all.
  //
  // 2026-08-20: the driver put the add-in back by itself for the first time,
  // reached `ready`, and then stopped on `no-run-button` with the pane sitting
  // on Chart and `verbose trace ?` in the same block — the checkbox that reads
  // as `?` lives on the same tab as the button that was missing. Everything
  // needed to notice was on screen; nothing acted on it.
  //
  // So: select the tab, then look again. This is the cheapest possible recovery
  // and it was previously an un-retryable stop.
  // The Automation tab was selected before readiness read its toggles, so the
  // button should be here. One more try anyway — the pane can be re-rendered
  // between the two, and a second look is cheaper than a lost round.
  let runBtn = refFor(sh, "Probe, then self-test", /button "Probe, then self-test"/);
  if (!runBtn) {
    const automation = refFor(sh, "Automation", /tab "Automation"/);
    if (automation) {
      clickRef(sh, automation);
      runBtn = refFor(sh, "Probe, then self-test", /button "Probe, then self-test"/);
    }
  }
  if (!runBtn) {
    console.error("  could not find the run button, and the Automation tab did not bring it back");
    // Swallowed on purpose: this runs on a path that has already failed, and a
    // diagnostic that can end a round is worse than no diagnostic.
    //
    // The READING is taken outside the swallow and used below. An unreadable
    // pane and a pane that answered nothing are not the same state, so a throw
    // here must leave the stop exactly as it was rather than promote it.
    let described = null;
    try {
      described = sh("find", "--regex", '/(tab "|button ")/');
      console.error(describePane(described));
    } catch {
      /* the pane is past describing; the refusal above is the report */
    }
    // A PANE WITH NO TABS AND NO BUTTONS IS A CLOSED PANE, and this driver has
    // always known how to reopen one. `no-run-button` has ended five cycles,
    // every time after a successful recovery and a printed `ready`; when the
    // diagnostic finally spoke on 2026-09-08 it said "the pane answered nothing
    // at all", seconds after the host answered in 3ms. Calling that a missing
    // BUTTON is what made it terminal.
    if (described !== null && paneAnsweredNothing(described)) {
      console.error("  nothing at all is on the pane — treating it as a closed pane, which recovery can reopen");
      return { code: 1, reason: "pane-closed" };
    }
    return { code: 1, reason: "no-run-button" };
  }
  console.log("  running — this takes about ten minutes");
  clickRef(sh, runBtn);

  // Polled, never slept-through: a wedged host is the normal failure here and it
  // has to be distinguishable from a slow one.
  const started = Date.now();
  const limit = 30 * 60 * 1000;
  // An empty answer ENDS the round, so it has to be believed twice. The CLI
  // serves one command at a time per session: anything else touching it while
  // this polls — a second terminal, an agent looking at the trace — makes one
  // poll exit non-zero, and folding that into "the pane is gone" killed a round
  // that was running perfectly and went on to finish 10 of 12 scenarios. The
  // report was worse than the loss: it named a crash that had not happened.
  let quiet = 0;
  /** Consecutive polls whose CLI call could not be run — see `browserDiedMidRound`. */
  let failedPolls = 0;
  for (;;) {
    if (Date.now() - started > limit) {
      console.error("  the round has not finished in 30 minutes — the host is wedged; see docs/ROUNDS.md");
      await keepCrashEvidence(sh, "30 minutes with no finish");
      return { code: 1, reason: "timeout" };
    }
    const dl = sh("find", "Download run log");
    // CAPTURED HERE, against the call that produced `dl`. `sh.state.lastFailed`
    // belongs to the MOST RECENT call, and the crash-dialog read below overwrites
    // it before either counter is updated — so both were judging this read's
    // emptiness against a different call's success.
    //
    // That re-opened the exact regression the comment above this loop describes.
    // The `failed` argument exists because one poll exiting non-zero means
    // NOTHING was measured, and folding that into "the pane is gone" once killed
    // a round that went on to finish 10 of 12 scenarios. With the flag taken from
    // the crash-dialog read instead, a failed `dl` read paired with a successful
    // crash read counts as a genuine silence: two of those in a row and the
    // driver kills a healthy round, files a crash report for a crash that never
    // happened, and lets recovery reload the tab out from under it.
    const dlFailed = sh.state.lastFailed;
    if (/button "Download run log"(?! \[disabled\])/.test(dl)) break;
    // WATCH FOR THE CRASH, do not wait it out. Rounds 30 and 31 each died about
    // three minutes in and then held this loop for the full thirty, because the
    // only thing it knew how to notice was the finish. Twenty-seven wasted
    // minutes twice over is most of an hour of a night's throughput.
    //
    // A DOM read, deliberately, not a ping: the pane is mid-round and an
    // Office.js call from here would interleave with the round's own batches and
    // change what it measures. The dialog is in the document frame and costs
    // nothing to look at.
    if (sawCrashDialog(sh("find", "Sorry, we ran into a problem"))) {
      const secs = Math.round((Date.now() - started) / 1000);
      console.error(
        `  PowerPoint crashed ${secs}s in — its dialog is up and nothing behind it will answer. ` +
          'See docs/ROUNDS.md, "The wedge".',
      );
      // NOW, before recovery reloads the tab and takes the request log with it.
      // This is the only window in which the host's own account of the crash
      // exists, and three separate hand passes were spent reaching it.
      await keepCrashEvidence(sh, `${secs}s into a round`);
      /**
       * AND THEN WAIT, because the pane has a way out of this and has never
       * once been given the time to take it.
       *
       * 41 of the 75 crash records on file — 55% — end at the same trace line:
       * `collecting deck evidence — scanning`. That is the POST-ROUND inventory
       * pass, which runs after every verdict is already banked in `lastRunLog`.
       * The pane bounds it with `DECK_EVIDENCE_TIMEOUT_MS`, 45 seconds, exactly
       * so a host dying in there cannot take the round with it; on timeout the
       * scan returns undefined, the round files without deck evidence, and
       * `demo-log` is enabled. Between that timeout and the button there is not
       * one `await` — no host call at all, only local work — and the pane
       * outlives the host, which is the same fact `keepCrashedRun` relies on to
       * get a file out of every one of these.
       *
       * THAT TIMEOUT HAS FIRED IN 0 OF 77 RECORDS. Not because it does not
       * work: because this branch returned the instant it saw the dialog, and
       * the driver was gone before the pane's 45 seconds were up.
       *
       * So the cost of finding out is one minute against a 30-minute wedge
       * budget, and the prize is a complete round — every verdict, every probe
       * answer — where today there is a crash report and a gap in the archive.
       * `docs/BACKLOG.md` item 14 asks whether a STUB may be filed for these.
       * The better answer is not to lose them.
       *
       * The crash evidence above is kept either way: the host really did crash,
       * and its account is worth having whether or not the round survives it.
       */
      if (await paneFinishedAnyway(sh, sleep, started)) break;
      return { code: 1, reason: "crashed" };
    }
    // THE BROWSER, not just the pane. Checked before the quiet counter because
    // the counter cannot see this state at all — see `browserDiedMidRound`.
    failedPolls = dlFailed ? failedPolls + 1 : 0;
    if (failedPolls >= DEAD_BROWSER_POLLS && browserDiedMidRound(failedPolls, sh("list"))) {
      const secs = Math.round((Date.now() - started) / 1000);
      console.error(
        `  the browser died ${secs}s into the round — the process is gone, taking the tab with it. ` +
          "The persistent profile still holds the sign-in, so this is recoverable without a password.",
      );
      return { code: 1, reason: "browser-gone" };
    }
    quiet = quietStreak(quiet, dl, dlFailed);
    if (quiet >= 2) {
      console.error("  the pane stopped answering — PowerPoint has probably crashed; the trace is still in the DOM");
      await keepCrashEvidence(sh, `${Math.round((Date.now() - started) / 1000)}s into a round, pane silent`);
      return { code: 1, reason: "silent" };
    }
    await new Promise((r) => setTimeout(r, 20000));
  }
  console.log("  finished");
  // FINISH THE JOB, rather than hand back a round somebody has to collect. The
  // three steps below were done by hand after every one of the 48 archived
  // rounds, and two of them are where the mistakes were: a stale log filed as a
  // fresh round (039 was byte-identical to 038), and a deck cleared by hand down
  // to zero slides.
  //
  // Each is best-effort and none can fail the round. The round is DONE by this
  // point — its evidence is in the pane either way, and a driver that turned a
  // good round into a non-zero exit over housekeeping would be worse than the
  // housekeeping.
  const roundFile = await collectRound(sh, stamp, sleep, size, deps.driverRun ?? null);
  return { code: 0, reason: "finished", roundFile, build: stamp, size };
}

/**
 * Download the run log, archive it, and leave the deck ready for the next round.
 *
 * `stamp` is the build the PANE was serving, checked against the log's own build
 * before anything is filed — see `archive`. That is what stops the previous
 * round's file being archived as a new one when a wedge left the download button
 * disabled and the click did nothing.
 *
 * Returns the name it filed, or null. The caller needs it for the outcome
 * receipt: a cycle runner that has just been told a round finished has no other
 * way to name WHICH round, and guessing "the newest file in rounds/" is the
 * assumption that already produced one wrong overwrite diagnosis in this repo.
 */
/**
 * Download the steps a CRASHED round managed to write, before a new one buries
 * them.
 *
 * A crashed round archives nothing: it never reaches the download button, so
 * the only thing left is the host's own console in `crashes/*.md`, which says
 * what PowerPoint thought and nothing about what WE were doing. Round 150
 * crashed six times in a row at 255-284s, and placing the crash needed the
 * scenario start times of OTHER rounds to guess at it.
 *
 * It did not have to. `crashlog.ts` flushes every step to `localStorage` as it
 * happens, and the next pane load calls `recoverCrashLog()` and un-hides
 * "Download the crashed run". The driver's own recovery reopens that pane, with
 * the same persistent profile and the same storage — so the button was sitting
 * there after every one of those six crashes and nobody pressed it. The one
 * crashed run in this repo's session directory was downloaded BY HAND, on
 * 2026-08-15.
 *
 * Kept beside the crash report rather than filed as a round: a partial run is
 * not a round, and `archive` is right to refuse it. Best-effort throughout —
 * a round is not worth failing over the paperwork of the round before it.
 */
/**
 * What a retry actually recovered FROM, named rather than bucketed.
 *
 * `reason` for a readiness stop is always the single word "not-ready", and four
 * rounds running recorded exactly that — which reads as "the host needed
 * rescuing" and mostly does not mean it.
 *
 * Round 153's first attempt refused for TWO reasons. `pane-stale`, because a
 * build had just been deployed and the pane still held the previous one; and
 * `host-silent`, because the editing session had dropped while the browser sat
 * idle between rounds. The first is a property of how rounds are RUN and says
 * nothing about the host. The second is host health. Folded into one word they
 * cannot be told apart, and any claim about the host getting better or worse is
 * then drawn from a bucket.
 *
 * That is the residual-bucket mistake, made inside a field added three rounds
 * earlier to stop exactly that.
 *
 * Sorted, so the same pair of stops reads the same way in every round and can
 * be counted across the archive.
 */
/**
 * Close the browser and build a new one, before a round is attempted.
 *
 * EVERY ROUND SO FAR HAS STARTED FROM WHATEVER THE LAST ONE LEFT BEHIND, gone
 * stale. Four rounds running needed a second attempt, and round 154 finally
 * named both reasons rather than bucketing them:
 *
 *     driverRun {"attempts":2,"recovered":["not-ready:host-silent+pane-stale"]}
 *
 * `host-silent` is the editing session dropping while the browser sits idle
 * between rounds. `pane-stale` is a build having been deployed since the pane
 * loaded. A fresh session should answer both — a new session is not silent, and
 * a pane loaded a moment ago is not stale.
 *
 * The argument is DETERMINISM, not speed. A round that starts from the previous
 * round's leftovers is a round whose starting conditions are a variable, and
 * every cross-round comparison in this archive rests on those being stable.
 *
 * CLOSE FIRST, THEN RECOVER. `recover` only opens a browser when there is not
 * one — `if (noBrowser(...))` — so recovering into a live browser reloads the
 * stale session instead of replacing it, which is the thing being avoided.
 *
 * Opt-in for now (`--fresh`). Whether it earns its place is a question
 * `driverRun.attempts` can answer, and it should be answered before it becomes
 * the default.
 */
/**
 * The only directories `pruneProfileCaches` will ever delete.
 *
 * Caches, and nothing else. What must SURVIVE, and why:
 *
 *   Network/         Chrome's cookie store — this IS the sign-in. Deleting it
 *                    means the owner signing in again by hand, which is the one
 *                    thing this loop cannot do for itself.
 *   Local Storage/   the session's own state, `crashlog.ts`'s flushed steps —
 *                    the only evidence a crashed round leaves behind — AND THE
 *                    SIDELOADED MANIFEST ITSELF. Microsoft's sideloading docs
 *                    say it outright: "when you sideload an add-in on the web,
 *                    the add-in's manifest is stored in the browser's local
 *                    storage, so if you clear the browser's cache... you have
 *                    to sideload the add-in again" (learn.microsoft.com,
 *                    testing/sideload-office-add-ins-for-testing, read
 *                    2026-08-22). Deleting this directory is how a cache tidy
 *                    turns into a lost add-in.
 *   Preferences      profile settings.
 *
 * Named explicitly rather than derived by excluding those, because a list of
 * things to keep is one rename away from deleting something new, and a list of
 * things to delete is not.
 */
export const PROFILE_CACHE_DIRS = ["Service Worker/CacheStorage", "Code Cache", "Cache", "GPUCache"];

/** Prune only when the profile is genuinely large; the caches cost time to rebuild. */
export const PROFILE_PRUNE_ABOVE_BYTES = 1_000_000_000;

/**
 * Bytes under `dir`, giving up as soon as the total passes `cap`.
 *
 * The point is the THRESHOLD, not the number, and a full walk of a 1.6GB
 * profile is tens of thousands of stat calls for a figure nobody reads.
 */
export function bytesAtLeast(dir, cap, { readdir = readdirSync, stat = statSync } = {}) {
  let total = 0;
  const stack = [dir];
  while (stack.length) {
    const at = stack.pop();
    let entries;
    try {
      entries = readdir(at, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const e of entries) {
      const p = `${at}/${e.name}`;
      if (e.isDirectory()) stack.push(p);
      else {
        try {
          total += stat(p).size;
        } catch {
          /* a file Chrome removed mid-walk is not a measurement failure */
        }
        if (total > cap) return total;
      }
    }
  }
  return total;
}

/**
 * Delete the caches a persistent profile accumulates, keeping the sign-in.
 *
 * `C:/devtools/pw-profile` reached 1.6GB in nine days, and 1.55GB of it was
 * cache: 805MB of service-worker CacheStorage, 401MB of V8 Code Cache, 337MB of
 * HTTP cache. It is PowerPoint Web doing it, not this project — our own site
 * appears in that cache zero times while `officeapps.live.com` appears 136
 * times. Every round loads the full editor, Microsoft ships new bundles often,
 * and a persistent profile never evicts the old ones.
 *
 * ONLY WITH THE BROWSER CLOSED. Chrome rewrites these files while it runs, and
 * deleting underneath it corrupts a profile that holds a sign-in only the owner
 * can replace.
 *
 * ONLY WHEN IT IS WORTH IT. The caches cost real time to rebuild — PowerPoint
 * re-downloads and re-JITs its bundles on the next round — so this is a
 * once-in-a-while reclaim, not a per-round tidy.
 */
export function pruneProfileCaches(profile, { browserOpen, rm = rmSync, exists = existsSync, sizeOf } = {}) {
  if (browserOpen) return null;
  const root = `${profile}/Default`;
  if (!exists(root)) return null;
  const measure = sizeOf ?? ((d) => bytesAtLeast(d, PROFILE_PRUNE_ABOVE_BYTES));
  let total = 0;
  for (const d of PROFILE_CACHE_DIRS) {
    const at = `${root}/${d}`;
    if (exists(at)) total += measure(at);
    if (total > PROFILE_PRUNE_ABOVE_BYTES) break;
  }
  if (total <= PROFILE_PRUNE_ABOVE_BYTES) return null;
  const removed = [];
  for (const d of PROFILE_CACHE_DIRS) {
    const at = `${root}/${d}`;
    if (!exists(at)) continue;
    try {
      rm(at, { recursive: true, force: true });
      removed.push(d);
    } catch {
      /* a cache that will not delete is not worth failing a night over */
    }
  }
  return { removed, bytes: total };
}

/**
 * `prune` is INJECTABLE, and that is not decoration.
 *
 * It defaults to the real thing, which deletes real directories with real
 * `rmSync`. A test that exercised this function without overriding it walked
 * `C:/devtools/pw-profile` and pruned 1.5GB of live caches — from a unit test,
 * in a suite run. The sign-in and the sideloaded manifest survived because the
 * exclusion list held, which is the design working; the call happening at all
 * is the design failing.
 *
 * Anything with a destructive real-filesystem default has to be reachable only
 * through a seam a test can close.
 */
/**
 * Tell the profile it exited on purpose, so the next start does not offer to
 * restore the pages we deliberately closed.
 *
 * `sh("close")` ends the browser process without Chromium's own shutdown, so
 * `Default/Preferences` keeps `exit_type: "Crashed"` — which is exactly the flag
 * the next launch reads to decide whether to show "Restore pages?". Every
 * `--fresh` leg left one behind, and the owner meets the bubble when they next
 * open that profile by hand.
 *
 * THE OBVIOUS FIX IS THE DANGEROUS ONE. Closing the tabs before the browser
 * looks tidier and is the thing this repo has already paid for twice: a
 * PowerPoint tab with unsaved work raises `beforeunload`, and accepting it
 * discards the work AND — measured after rounds 124 and 132 — the per-document
 * SIDELOAD, which then refuses the next several rounds with `addin-missing`.
 * The sweep that runs just before this guarantees unsaved work. So the tabs are
 * left alone and the flag is repaired instead: no dialog, nothing to accept,
 * nothing to lose.
 *
 * Injectable for the reason the prune beside it is: a real-filesystem default
 * that a test could reach once walked a live 1.5GB profile from a unit run.
 */
export function clearCrashFlag(profile, { read = readFileSync, write = writeFileSync, exists = existsSync } = {}) {
  const file = `${profile}/Default/Preferences`;
  if (!exists(file)) return false;
  try {
    const prefs = JSON.parse(read(file, "utf8"));
    prefs.profile ??= {};
    // Already clean — say so rather than rewriting a large file for nothing.
    if (prefs.profile.exit_type === "Normal" && prefs.profile.exited_cleanly === true) return false;
    prefs.profile.exit_type = "Normal";
    prefs.profile.exited_cleanly = true;
    write(file, JSON.stringify(prefs));
    return true;
  } catch {
    // A profile we cannot parse is not one to guess at. The restore bubble is a
    // nuisance; a corrupted Preferences file costs the sign-in.
    return false;
  }
}

export async function startFresh(
  sh,
  sleep,
  recoverFn = recover,
  profile = PROFILE_DIR,
  prune = pruneProfileCaches,
  clearFlag = clearCrashFlag,
) {
  sh("close");
  await sleep(3000);
  // HERE, and only here: the browser is closed and has not been rebuilt yet,
  // which is the one moment in a round when these files are nobody's.
  const pruned = prune(profile, { browserOpen: false });
  if (pruned)
    console.log(
      `  pruned ${pruned.removed.length} profile cache(s) — over ${Math.round(PROFILE_PRUNE_ABOVE_BYTES / 1e9)}GB, ` +
        "sign-in and crash log untouched",
    );
  // Same moment, same reason: the process is gone, so the file is nobody's.
  if (clearFlag(profile)) console.log("  marked the profile as cleanly exited — no restore prompt next time");
  await recoverFn(sh, sleep);
}

/**
 * Hosts that serve a sign-in popup. Nothing else is ever closed by the function
 * below, and the list is deliberately short and explicit.
 */
/** Where the deck lives, and the one page recovery navigates back to. */
export const DECK_HOME_URL = process.env.PW_DECK_HOME ?? "https://onedrive.live.com/";

const AUTH_POPUP_HOSTS = /login\.live\.com|login\.microsoftonline\.com/;

/**
 * Is this page a sign-in PROMPT, a sign-in ERROR, or neither?
 *
 * THE DISTINCTION THE LOOP DID NOT MAKE, and it cost a six-hour unattended run.
 *
 * On 2026-08-23 the sole tab was
 * `login.microsoft.com/consumers/fido/get`, titled "Sign in to your account".
 * The URL and the title both said sign-in, so the loop stopped and notified the
 * owner that he was needed. **He was not.** The page rendered:
 *
 *     Sign in
 *     Sorry, but we're having trouble signing you in.
 *     AADSTS900561: The endpoint only accepts POST requests. Received a GET request.
 *     Request Id: ...  Correlation Id: ...  Timestamp: ...
 *
 * No field, no button to proceed, nothing to enter. An error page REPORTING that
 * a sign-in failed, produced by a bad GET — and the profile still held a live
 * session, which navigating back to OneDrive proved in one command.
 *
 * The same page had already caused one false stop eight hours earlier. The test
 * written after that one was "check the title", and this is the second failure
 * of that test: a title can say "Sign in" on a page whose only content is a
 * request id. **Read what is rendered, not what the tab is called.**
 *
 * CONSERVATIVE BY CONSTRUCTION. `prompt` is the default for anything on an auth
 * host that this cannot positively identify as an error, because the cost of
 * the two mistakes is not symmetric: treating a prompt as an error would have
 * the loop navigate away from something a person is part-way through, and
 * treating an error as a prompt only wastes a round. An `AADSTS` code with NO
 * input on the page is the one shape that cannot be awaiting entry.
 *
 * Reading only. Nothing here types, clicks or closes.
 */
export function authPageKind(url, snapshot) {
  if (!AUTH_POPUP_HOSTS.test(String(url ?? "")) && !/login\.microsoft\.com/.test(String(url ?? ""))) return "not-auth";
  const text = String(snapshot ?? "");
  // Anything that could accept a credential. A prompt awaiting entry always has
  // one of these; an error page has none.
  const takesInput = /textbox|passwordbox|button "Next"|button "Sign in"|button "Use your passkey"/i.test(text);
  const isError = /AADSTS\d+/.test(text) || /having trouble signing you in/i.test(text);
  if (isError && !takesInput) return "error";
  return "prompt";
}

/**
 * Close a sign-in popup that outlived the flow it belonged to.
 *
 * MSAL opens `login.live.com/oauth20_authorize.srf` in a popup and closes it
 * when the flow finishes. On 2026-08-22 one did not: it sat as a third tab for
 * three hours while the deck worked normally beside it, and every `tab-list`
 * the driver read carried it.
 *
 * FOUR GATES, because a sign-in window is the one surface this loop must never
 * meddle with:
 *
 *   1. THE HOST MUST HAVE ANSWERED. That is the proof the popup is superfluous
 *      — Office.js answering means the session is valid, so nothing is waiting
 *      on that window. Without it this closes nothing, which is the right
 *      answer for a flow that might still be live.
 *   2. NEVER THE CURRENT TAB. A fronted sign-in window may be one a PERSON is
 *      part-way through, and the driver cannot tell.
 *   3. ONLY these hosts, matched on the URL.
 *   4. NOTHING IS READ, TYPED OR CLICKED inside it. The only action is
 *      `tab-close` on an index. Closing a stale WINDOW is not the same as
 *      interacting with a credential PROMPT, and this must never become that.
 *
 * Highest index first, because closing a tab renumbers the ones after it.
 */
export function closeStaleAuthPopups(sh, { hostAnswered } = {}) {
  if (!hostAnswered) return [];
  const indices = [];
  for (const line of String(sh("tab-list") ?? "").split("\n")) {
    if (/\(current\)/.test(line)) continue;
    if (!AUTH_POPUP_HOSTS.test(line)) continue;
    const n = /^\s*-\s*(\d+):/.exec(line)?.[1];
    if (n !== undefined) indices.push(Number(n));
  }
  const closed = [];
  for (const n of indices.sort((a, b) => b - a)) {
    sh("tab-close", String(n));
    closed.push(n);
  }
  return closed;
}

export function recoveryLabel(reason, codes) {
  return Array.isArray(codes) && codes.length ? `${reason}:${[...codes].sort().join("+")}` : String(reason);
}

export async function keepCrashedRun(
  sh,
  sleep,
  copy = copyFileSync,
  exists = existsSync,
  read = readFileSync,
  write = writeFileSync,
  // Injected like the rest, so a test can prove the landing spot is cleared
  // BEFORE the click rather than infer it from the file that ends up archived.
  rm = rmSync,
) {
  try {
    const ref = refFor(sh, "Download the crashed run", /button "Download the crashed run"/);
    // Nothing kept, or a previous attempt already saved it — `clearCrashLog`
    // hides the button once pressed, so this does not re-download in a loop.
    if (!ref) return null;
    const from = `${sh.dir ?? "."}/.playwright-cli/ssf-charts-crashed-run.json`;
    /**
     * CLEAR THE LANDING SPOT FIRST, so "its file never arrived" means what it
     * says.
     *
     * The download lands on a fixed path. This used to click, wait, and archive
     * whatever was sitting there — so a click that downloaded NOTHING, because
     * the pane had died or the button was stale, re-archived the PREVIOUS
     * download under today's date.
     *
     * Not a worry, an event: `crashes/2026-09-08T13-42-37-crashed-run.json` was
     * written on 2026-09-08 and is byte-identical to
     * `2026-09-07T22-13-42-crashed-run.json` — the same build stamp `36916a5 ·
     * 2026-09-07 21:49Z`, the same `startedAt` to the millisecond, the same 855
     * steps. Yesterday's crash filed a second time as though it were today's,
     * in the very attempt that went on to find the pane empty.
     *
     * Nothing downstream was corrupted, and only because `loadCrashRecords`
     * dedupes on `build + startedAt + steps`. Its docstring — "ONE EVENT FILED
     * TWICE IS ONE EVENT" — was written for a different re-archive, and a death
     * rate protected by a coincidence is not protected.
     */
    try {
      rm(from, { force: true });
    } catch {
      // A landing spot that will not clear is no reason to skip the salvage.
      // The freshness check below simply becomes as weak as it used to be.
    }
    clickRef(sh, ref);
    await sleep(8000);
    if (!exists(from)) {
      console.error("  a crashed run was offered but its file never arrived");
      return null;
    }
    const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const to = `crashes/${stamp}-crashed-run.json`;
    /**
     * STAMPED WITH THE TAB IN FRONT, and named for exactly what that is.
     *
     * 160 crash records name no document, and they are the half of the archive
     * that matters most: 61 of them are the 4:3 arm, which is the arm the whole
     * enquiry is about. Without a document name, "4:3 crashes far more" and
     * "that one file is sick" are the same sentence.
     *
     * BUT THIS IS NOT THE DECK THAT CRASHED, and calling it one would be worse
     * than recording nothing. `keepCrashedRun` runs at the START of the next
     * round — see the call site — rescuing the previous round's remains before
     * they are overwritten. What is in front now is the NEXT round's starting
     * state. On the ordinary path recovery has reopened the same document and
     * the two agree; after a `--fresh` restart or a cycle changing legs they
     * need not. A field called `driverDeck` here would be read as the crashed
     * round's deck by every future query, including the ones that matter.
     *
     * So it is named for the reading actually taken. In a controlled run where
     * `PW_DECK` is pinned for the leg it is exact; when mining the archive it is
     * a strong hint that says so.
     *
     * PARSE-AND-WRITE, falling back to the byte copy on anything unexpected. A
     * crash record that arrives malformed is still the only copy of a round that
     * died, and losing it to a stamping error would be the worst trade in this
     * file. Two-space JSON with a trailing newline, because `crashes/` is not in
     * `.prettierignore` and the gate checks it.
     */
    let stamped = false;
    try {
      const log = JSON.parse(read(from, "utf8"));
      log.frontedWhenRescued = archivableDeck(frontedDeck(sh));
      log.deckAskedNextRound = process.env.PW_DECK ?? null;
      write(to, JSON.stringify(log, null, 2) + "\n");
      stamped = true;
    } catch (err) {
      console.error(`  (could not stamp the crashed run, keeping the bytes as they came: ${err?.message ?? err})`);
    }
    if (!stamped) copy(from, to);
    console.log(`  kept the steps of a round that never finished — ${to}`);
    return to;
  } catch (err) {
    console.error(`  (could not keep the crashed run: ${err?.message ?? err})`);
    return null;
  }
}

async function collectRound(sh, stamp, sleep, driverSize = null, driverRun = null) {
  let filed = null;
  try {
    const dl = refFor(sh, "Download run log", /button "Download run log"/);
    if (!dl) {
      // SAY SO. This was a bare `return null`, and it is the only failure on
      // this path that printed nothing — three lines below, the catch calls that
      // out as the rule: "Named, never swallowed."
      //
      // It cost a cycle on 2026-08-26. The round finished, the receipt recorded
      // `roundFile: null`, `cycle.mjs` correctly refused to run the next leg
      // rather than overwrite the log — and the reason it stopped appeared
      // nowhere, so the two remaining legs of the night were lost to a
      // diagnosis that had to be read out of the source.
      console.error("  the pane offered no `Download run log` button — the round ran but cannot be filed");
      console.error("  archive it by hand: node scripts/round.mjs --archive .playwright-cli/ssf-charts-run-log.json");
      return null;
    }
    clickRef(sh, dl);
    await sleep(12000);
    // WHERE THE BROWSER PUTS IT IS NOT WHERE THIS INVOCATION EXPECTS IT.
    //
    // The download directory belongs to the RUNNING BROWSER SESSION, fixed when
    // it was opened, and does not move because a later invocation passed a
    // different `--dir`. So a browser opened by `node scripts/round.mjs` (dir =
    // cwd) keeps downloading to `./.playwright-cli/` even when `cycle.mjs` then
    // spawns legs with `--dir .pw-session` and reads `.pw-session/.playwright-cli/`.
    //
    // That mismatch cost two cycles on 2026-08-27, and its symptom was actively
    // misleading: a STALE log from an earlier round sits at the expected path,
    // so `existsSync` is true and the round is "archived" from the wrong file.
    // What caught it was `archive`'s build check — "that log is build 67f4124
    // and the pane is serving e97699e" — which is the guard doing precisely the
    // job it was added for.
    //
    // Take the NEWEST of the candidates rather than trusting either path. Safe
    // because the build check still runs on whatever is chosen: a wrong pick is
    // refused, not filed.
    const candidates = [...new Set([`${sh.dir ?? "."}/.playwright-cli`, ".playwright-cli"])].map(
      (d) => `${d}/ssf-charts-run-log.json`,
    );
    const found = candidates
      .filter((p) => existsSync(p))
      .map((p) => ({ p, at: statSync(p).mtimeMs }))
      .sort((a, b) => b.at - a.at);
    if (!found.length) {
      console.error("  the run log did not arrive — archive it by hand once it does");
      return null;
    }
    const logPath = found[0].p;
    if (found.length > 1 && found[0].p !== candidates[0])
      console.log(`  the run log landed in ${logPath}, not where this invocation looked — using the newer one`);
    // `frontedDeck` READ HERE, not passed down from the readiness block, so the
    // field says which document the evidence was collected FROM rather than
    // which one was in front when the round was judged worth starting. On the
    // ordinary path they are the same tab; when they are not, this is the one
    // that describes the file being archived.
    filed = archive(
      logPath,
      "rounds",
      readFileSync,
      writeFileSync,
      everyRoundEverFiled,
      stamp,
      driverSize,
      driverRun,
      archivableDeck(frontedDeck(sh)),
    );
    console.log(`  archived as rounds/${filed}`);
  } catch (err) {
    // Named, never swallowed. A round whose log was not filed is a round that
    // will be filed by hand, and the person doing it needs to know why.
    console.error(`  could not archive this round: ${err instanceof Error ? err.message : String(err)}`);
  }
  // The deck last, so a failed archive still leaves it ready — the two are
  // independent and coupling them would cost the next round for the sake of
  // this one's paperwork.
  // SAY WHICH, because the two lead to different mornings. A sweep that did not
  // clean the deck leaves the next round to refuse with `deck-dirty`, and this
  // line used to claim the opposite whatever happened.
  if (sweepDeck(sh)) console.log("  deck swept — the next round starts clean");
  else console.error("  the deck was NOT swept — the next round will refuse until it is");

  // A CLEAN DECK IS NOT A CLEAN PANE, and until 2026-08-20 this line was the
  // only thing standing between one round and the next. The pane's age at a
  // round's start separates post-retry 0.43 from 4.57 and a 16-shape deck from
  // a 60+ one; sweeping clears the slides and leaves the pane exactly as the
  // last round left it. Every second round in this archive is a degraded
  // sample because of it.
  //
  // AFTER the sweep, because `sweepDeck` needs the pane it is about to replace.
  // Best-effort like everything else here: the round is already archived, and a
  // driver that turned a good round into a non-zero exit over housekeeping
  // would be worse than the housekeeping.
  // THE BETWEEN-ROUNDS RELOAD IS OFF, and this comment is the reason rather
  // than a TODO. It worked — round 126 was the first second round in this
  // archive to score a first round's numbers — and it cost the session roughly
  // one round in four.
  //
  // The mechanism: a reload of a tab with unsaved work raises PowerPoint's
  // beforeunload prompt, and `dialog-accept` means LEAVE WITHOUT SAVING. The
  // sweep immediately above guarantees unsaved work — it has just deleted
  // slides — so accepting discards it. It also appears to discard the
  // per-document SIDELOAD: the add-in was gone from the ribbon after the reload
  // that followed round 124 and again after round 132, and rounds 4, 5 and 6 of
  // that batch refused with `addin-missing` against a deck that had grown back
  // to 79 slides.
  //
  // Dismissing instead would cancel the reload rather than save. The real fix is
  // to WAIT FOR THE AUTOSAVE and reload only once the document is clean, so no
  // prompt appears at all — which needs a way to read PowerPoint's saved state
  // that this driver does not have yet, and which must not be guessed at with a
  // sleep on the one path whose failure costs the add-in.
  //
  // `refreshPane` stays: `recover` uses it, and it is correct there because the
  // page it reloads has usually already crashed and has nothing to save.
  // `paneAgeAtStartSeconds` stays too — the gate still reports whether a round
  // inherited a pane, which is what makes its counters readable.
  return filed;
}

/**
 * The stops a reload-and-reopen actually clears — which is to say, the stops
 * `recover` was written for.
 *
 * Derived from that function rather than from a judgement about which refusals
 * feel transient: `recover` clicks Refresh or reloads, waits out the reload,
 * reopens the pane from the ribbon, clicks the Automation tab, and cleans the
 * deck. Every code here is undone by one of those five steps, and every code
 * NOT here survives all of them.
 *
 * `wrong-size` is deliberately absent and must stay that way. `recover` could
 * set a deck's slide size — it is two clicks — and doing so would CHANGE WHAT
 * THE ROUND MEASURES rather than restore it, which is the one thing recovery is
 * not allowed to do. A deck in the wrong profile is a setup error for a person,
 * not a transient state to clear.
 *
 * `addin-missing` is absent for the same kind of reason and was learned the same
 * way. `recover` reopens the pane from the ribbon's `Insert chart` control, so a
 * document that does not carry that control has nothing for recovery to click;
 * on 2026-08-16 the driver spent seven attempts and about fifteen minutes
 * proving that twice in one night.
 *
 * "AND ONLY A PERSON CAN PUT IT BACK" STOOD HERE AND IS NO LONGER TRUE.
 * On 2026-08-20 a browser reopen took the sideload with it, `sideloadAddIn`
 * uploaded the manifest, and the ribbon carried `Insert chart` and `Insert
 * element` again with the owner touching nothing. The machine restored it. What
 * it got wrong was the WAIT: it checked once after a fixed sleep, called the
 * absence permanent, and sent a person to redo work already done.
 *
 * The stop stays out of `RECOVERABLE_STOPS` all the same, and for the ORIGINAL
 * reason only — recovery clicks `Insert chart` to reopen the pane, so a document
 * without that control gives recovery nothing to click, and looping on it burns
 * a night. The sideload path is where the patience belongs, and that is where it
 * now is (`SIDELOAD_COMMAND_BUDGET_MS`).
 *
 * Deliberately absent, and each for its own reason: `site-behind` and `no-build`
 * are waiting for Pages and a reload does not make it deploy faster — though
 * `no-build`'s reason is only half true, and knowing which half matters. It
 * fires both when Pages has nothing to say and when this machine cannot reach
 * it at all (2026-09-05: node's `fetch` timing out on a 250ms per-address
 * budget while `curl` answered in 850ms). Reloading cures neither, so it stays
 * out; but the refusal itself now names both, because the reader's next move
 * differs completely between them;
 * `verbose-off` and `pictures-off` are choices a person made in the pane and
 * silently re-making them would change what the round measures; the sign-in and
 * unreachable-CLI states never get here because they return before the codes do.
 */
/**
 * Did a spawn fail because the tool is not there, or because the machine was busy?
 *
 * These want OPPOSITE treatment and had one handler between them. A missing
 * binary is not going to appear on the second try, so failing fast is right. A
 * spawn that TIMED OUT is the machine being busy, and retrying is the whole
 * cure — but the unreachable-CLI stop returns before any code is assigned, so
 * `--retry 6` never got a chance at it.
 *
 * WHAT IT COST: round 197, 2026-08-23. `spawnSync … node.exe ETIMEDOUT` on
 * `eval () => String(window.innerWidth)` — the first call of the round, while the
 * full 124-file vitest suite was running in the same session. The tool was
 * installed and had run six rounds that day. The driver exited 1 on a condition
 * that would have cleared by itself, and the message sent the reader to the
 * install.
 *
 * Conservative on purpose: ONLY the errnos that mean "could not start it just
 * now" are transient. Anything unrecognised stays fatal, because a stop that
 * loops is worse than one that stops — this loop runs unattended for hours.
 */
export function spawnFailureIsTransient(errorText) {
  return /ETIMEDOUT|EAGAIN|EBUSY|EMFILE|ENFILE|ENOMEM|SIGTERM|SIGKILL/i.test(String(errorText ?? ""));
}
export const RECOVERABLE_STOPS = new Set([
  // A spawn that timed out rather than a tool that is missing. See
  // `spawnFailureIsTransient` — the two states used to share one handler and
  // the transient one exited the run.
  "cli-busy",
  "browser-gone",
  "crashed",
  "host-silent",
  "slide-refused",
  "pane-closed",
  "pane-stale",
  "deck-dirty",
  // THE DECK THE CYCLE NAMED IS NOT OPEN — and `recover` has always known how to
  // fix this. It navigates to OneDrive, clicks the file whose link matches the
  // name, and selects the tab either way; the "Click the file only when no tab
  // holds it; select it either way" branch exists for exactly this state.
  //
  // It was simply never wired to it. The stop said "Open it, or unset PW_DECK"
  // and ended the night, while the function that opens it sat one call away. A
  // full cycle died this way on 2026-08-27: leg 1 wanted `Presentation64`, its
  // tab had closed during a `--fresh` restart, and the driver stayed on the 4:3
  // deck and refused — correctly, but for something it could have repaired.
  //
  // NOT the same class as `addin-missing` or `wrong-size`. Those need a person
  // because the first cannot be sideloaded unattended and the second would
  // CHANGE WHAT THE ROUND MEASURES. Opening the deck the caller explicitly named
  // changes nothing about the measurement — it is the measurement being set up.
  //
  // Bounded by `--retry` like every other stop, so a `PW_DECK` naming a document
  // that does not exist costs that many navigations and then reports honestly.
  "deck-missing",
  // A window too narrow for PowerPoint to render its ribbon commands. Recovery
  // does not need to click anything to clear it — `ensureRibbonRoom` widens the
  // window on the next attempt — and the alternative is `addin-missing`, a stop
  // only a person can clear, fired on a command that was never absent.
  "ribbon-cramped",
  // A greyed-out ribbon on a `Disconnected` document. Transient by nature — the
  // tab reconnects or a reload clears it — and it used to be reported as
  // `addin-missing`, which recovery is forbidden to retry.
  "host-disconnected",
  // A call that never RAN. The reading it would have produced is unknown, not
  // negative, and the next attempt's calls usually run — a transient ENOENT
  // from `spawnSync` cost an evening on 2026-08-22 by being read as an absent
  // pane, an absent deck and finally an absent add-in.
  "reads-failed",
]);

/**
 * Is another attempt worth making?
 *
 * IT USED TO BE A CRASH AND NOTHING ELSE, and the cost of that showed up on
 * 2026-08-15 in two places on one afternoon. Mid-round, the QUIET form of the
 * wedge exits as `silent` — the host stops answering with no dialog — and the
 * driver went home, though `docs/ROUNDS.md` says in as many words that a reload
 * clears both forms and `recover` already does exactly that. At check time, a
 * person then hand-fixed a silent host, a pane one build behind and an
 * eight-slide deck, in the same order `recover` does them, because a `not-ready`
 * was never retried either.
 *
 * So the question is no longer "was it a crash" but "does recovery address
 * everything that refused". A `not-ready` whose codes are all recoverable is
 * worth another attempt; one carrying a single stop recovery cannot touch is
 * not, and stopping on it is the behaviour the old comment was right about — a
 * round that retries a stale build until the night is gone measures nothing.
 *
 * `codes` is optional so the two reasons that carry none (`crashed`, `silent`)
 * read the same as they always did.
 */
/**
 * What the recovery about to run is actually recovering FROM.
 *
 * One line, and it is read at the worst moment — mid-loop, by someone deciding
 * whether the round is worth watching. Saying "clearing the crash" when the deck
 * was merely dirty is how a debugging session starts by hunting a crash that
 * never happened.
 */
export function recoveryFor(reason, codes) {
  if (reason === "crashed") return "clearing the crash and starting again";
  if (reason === "silent") return "the host went quiet — reloading and starting again";
  if (reason === "timeout") return "the round wedged and did not finish — reloading and starting again";
  const named = (codes ?? []).filter((c) => RECOVERABLE_STOPS.has(c));
  if (!named.length) return "recovering and starting again";
  const say = {
    crashed: "a crash dialog",
    "host-silent": "a silent host",
    "slide-refused": "a host that would not resolve slide 1",
    "pane-closed": "a closed pane",
    "pane-stale": "a stale pane",
    "deck-dirty": "a dirty deck",
  };
  return `recovering from ${named.map((c) => say[c] ?? c).join(" and ")}, then starting again`;
}

/**
 * Has the BROWSER died under a round that is still polling?
 *
 * The hole this closes, and it cost 24 minutes on 2026-08-16. `quietStreak`
 * resets to zero whenever a CLI call FAILED, deliberately: one failed call means
 * nothing was measured, and treating it as "the pane is gone" once killed a
 * healthy round that went on to pass 10 of 12. That protection is right and is
 * left alone here.
 *
 * But a dead browser makes every call fail, permanently — so the quiet counter
 * can never reach its threshold, the crash dialog cannot be read either, and the
 * loop polls a corpse until the thirty-minute limit. `pw list` said
 * `(no browsers)` outright while the driver sat there.
 *
 * Two conditions, and both are needed. A STREAK of failures, so ordinary
 * contention (a second terminal, an agent reading the trace) cannot trigger it —
 * three consecutive misses is a minute of silence, far past any blip. And then
 * an affirmative `(no browsers)`, so the round is never ended on the absence of
 * evidence. Either alone would re-make the mistake the other guards against.
 */
export const DEAD_BROWSER_POLLS = 3;
export function browserDiedMidRound(failedStreak, listOutput) {
  return failedStreak >= DEAD_BROWSER_POLLS && noBrowser(listOutput);
}

/**
 * Is a dirty deck the ONLY thing standing in the way?
 *
 * A dirty deck is not a fault — it is the last round's slides, and the driver
 * already knows how to sweep them (`cleanDeckScript`, which `recover` runs every
 * time it fires). Refusing over it makes a person do by hand the one step the
 * machine does better: hand-clearing took a deck to ZERO slides on 2026-08-16
 * and produced `slide 1 REFUSED`, the state the 2s crash starts from.
 *
 * ONLY, and that word is the whole guard. A deck that is dirty AND on a stale
 * pane is a round that would measure the wrong build, and healing the cheap half
 * moves it one step closer to running while still being wrong. Anything beyond a
 * dirty deck refuses out loud, as before.
 */
export function onlyDirtyDeck(codes) {
  return Array.isArray(codes) && codes.length === 1 && codes[0] === "deck-dirty";
}

export function shouldRetry(reason, attempt, max, codes) {
  if (attempt >= max) return false;
  // A WEDGE MID-ROUND IS RECOVERABLE, and leaving `timeout` out of this stopped
  // an unattended run dead in its first hour. The quiet wedge that produces it is
  // the same state `silent` names — `docs/ROUNDS.md` says a reload clears both
  // forms — and `recover` already does exactly that. The argument for excluding
  // it was cost: a wedge has already burned thirty minutes and another may burn
  // thirty more. But that cost is bounded by `--retry N`, which the caller chose,
  // and the alternative is a ten-hour run that ends at hour one with the host
  // sitting idle and recoverable.
  // `browser-gone` joins them: the profile keeps the sign-in, so reopening needs
  // no password and `recover` already does it. A round that ends this way has
  // burned a minute, not thirty — it is the cheapest of these to retry.
  // `threw` joins them: an unexpected exception used to kill the process
  // outright, so `--retry` never saw it. Retrying is bounded by `max` — a
  // deterministic bug fails that many times and stops — and the alternative is
  // a night that ends on its first surprise with recovery never attempted.
  if (
    reason === "crashed" ||
    reason === "silent" ||
    reason === "timeout" ||
    reason === "browser-gone" ||
    reason === "threw"
  )
    return true;
  if (reason !== "not-ready") return false;
  // An EMPTY list is not a licence. It means nothing was recorded about why the
  // check refused, and retrying on no evidence is how a loop spins.
  return Array.isArray(codes) && codes.length > 0 && codes.every((c) => RECOVERABLE_STOPS.has(c));
}

/** Delete every slide but the first, so the next round starts where the last one did. */
export function cleanDeckScript(budgetMs) {
  return (
    "async () => { const budget = (p, ms) => Promise.race([p, new Promise((_, r) => " +
    `setTimeout(() => r(new Error("TIMEOUT")), ms))]); try { const n = await budget(PowerPoint.run(async (c) => { ` +
    'const s = c.presentation.slides; s.load("items/id"); await c.sync(); const count = s.items.length; ' +
    "for (let i = count - 1; i >= 1; i--) c.presentation.slides.getItemAt(i).delete(); await c.sync(); " +
    `s.load("items/id"); await c.sync(); return s.items.length; }), ${budgetMs}); return "deck:" + n; } ` +
    'catch (e) { return "deck-failed"; } }'
  );
}

/**
 * Put PowerPoint back on its feet — the recovery done by hand six times tonight.
 *
 * Every step was learned the expensive way and none is optional:
 *
 *   - The dialog's Refresh button is matched BY TEXT. Its accessible name is
 *     sometimes absent, the label sitting in a child `generic` instead, so
 *     `button "Refresh"` finds it on one crash and not on the next.
 *   - Refreshing reloads the document, which closes the pane. It comes back from
 *     Home ▸ Add-ins ▸ **Insert chart** — that is the `ShowTaskpane` control,
 *     despite the name.
 *   - The Automation tab has to be showing or the run button is not in the DOM.
 *   - The deck has to go back to one slide, or the next round is not comparable
 *     with one that started clean.
 *
 * The waits are generous on purpose: a document reload takes tens of seconds and
 * a step taken early lands on nothing and fails silently.
 */
export async function recover(sh, sleep, profile = PROFILE_DIR) {
  // NO BROWSER AT ALL comes first, because everything below reloads and clicks
  // inside a window that is not there. Not hypothetical: on 2026-08-15 a round
  // wedged, the browser process died with it, and this function then reloaded
  // nothing and reopened nothing seven times while the check reported "could not
  // read the pane's build stamp — is the add-in open?"
  //
  // Reopening is the loop's to do, not the owner's. The persistent profile still
  // holds the sign-in — a dead browser is not a lost sign-in — so it needs no
  // password, and the alternative is a ten-hour run ending in its first hour.
  // REACHABLE, not merely present. See `browserReachable` — `list` sees every
  // browser the daemon knows, including one belonging to another session that
  // this driver cannot issue a single read against.
  if (!browserReachable(sh)) {
    // A BROWSER WE CANNOT USE STILL HOLDS THE PROFILE, and `open` refuses with
    // "Browser is already in use for <profile>, use --isolated". Isolated is no
    // use here: a fresh profile has no sign-in, and the sign-in is the one
    // thing this loop cannot recreate. So the unusable session has to go first.
    //
    // Guarded on there BEING one, so the ordinary "no browser at all" path does
    // not pay for a close it does not need.
    if (!noBrowser(sh("list"))) {
      console.log("  a browser is open that this session cannot reach — closing it so one can be opened here");
      sh("close-all");
    }
    sh("open", ...roundConfigArg(), "--persistent", `--profile=${profile}`, "--headed", "https://onedrive.live.com/");
    // AND END THE ORPHAN WHEN THE CLOSE ABOVE COULD NOT HAVE HELPED.
    //
    // `close-all` reaches browsers the DAEMON knows about. An orphan — Chrome
    // still running after its daemon died — is not one of those, so `list`
    // reports nothing, the guard above skips the close, and `open` then refuses
    // for a profile the driver has just been told is free. Recovery walks on to
    // the deck and blames the file list, every attempt, until `--retry` is out.
    //
    // THIS WAS A MESSAGE FOR AN OPERATOR FOR ABOUT AN HOUR, on the argument
    // that ending an OS process is a decision about someone's machine. Then the
    // state recurred: the browser died 746s into round 415 and left thirteen
    // Chrome processes holding the profile. A message is the right answer to a
    // condition a person is watching; this one arrives mid-cycle, unattended,
    // and every attempt after it is wasted. The driver already closes browsers
    // — `--fresh` does it every leg — so ending one it launched, on its own
    // profile, only after `open` has refused for exactly this reason, is the
    // same authority rather than a new one.
    //
    // AND IT COSTS NOTHING THAT WAS NOT ALREADY LOST. The obvious objection is
    // that a browser death takes the sideload with it, so killing one throws
    // away a working add-in. It does not: this branch is only reached when
    // `list` reports no browser AND `open` refuses the profile, which together
    // mean nothing here can issue a single call against that Chrome. The
    // sideload inside it is unreachable whether or not the process keeps
    // running. What is thrown away is a process, not a capability.
    if (profileHeldByOrphan(sh.state?.lastStderr)) {
      console.log(`  a browser holds ${profile} but answers nothing — ending it, then opening again`);
      if (endOrphanedBrowser(profile)) {
        await sleep(3000);
        sh(
          "open",
          ...roundConfigArg(),
          "--persistent",
          `--profile=${profile}`,
          "--headed",
          "https://onedrive.live.com/",
        );
      }
      // NOT "it worked". The retry above either opened a browser or did not,
      // and the readiness check below is what says which — an inference here
      // would be a third opinion nobody asked for.
      if (profileHeldByOrphan(sh.state?.lastStderr))
        console.error(`  the profile is still held — end the Chrome on ${profile} by hand, then this recovers itself`);
    }
    await sleep(15000);
  }
  // THE DECK TAB, WHETHER OR NOT THE BROWSER IS NEW.
  //
  // This block used to sit inside the `noBrowser` branch above, which assumed a
  // living browser implies an open deck. It does not. A browser sitting on
  // OneDrive's home — after a tab crash, or after anything reopened the browser
  // without the deck — made `recover` skip the deck entirely, reload the wrong
  // page, and then hunt for a pane that had never been opened. Seven attempts
  // in a row reported "could not read the pane's build stamp — is the add-in
  // open?" with `deck ? slide(s)` beside it, while the deck sat in the file
  // list four minutes old.
  //
  // Asked of the TAB LIST rather than of the browser, because "is the deck
  // open" is the actual question and the browser's existence never answered it.
  {
    // The deck, and then ITS tab. Clicking the file opens a NEW tab while the
    // CLI stays on the old one, and skipping that is how a healthy setup reads
    // as a closed pane.
    // THE DECK'S NAME IS NOT A CONSTANT, and hard-coding it here was a trap that
    // fired the day it was written about. This said `Presentation63` while the
    // deck in use had become `Presentation64` — a new document, because a web
    // sideload is per-document — so a browser death would have reopened OneDrive
    // and then failed to find anything, silently, in exactly the situation this
    // function exists for. `PW_DECK` overrides it for a deck named anything else.
    // A SIGN-IN ERROR PAGE IS NOT A SIGN-IN PROMPT, and telling them apart is
    // what this loop got wrong twice in one day.
    //
    // On 2026-08-23 the sole tab was a login host titled "Sign in to your
    // account", and the run stopped and notified the owner. The page was an
    // AADSTS900561 error with no field on it — a bad GET — and the profile still
    // held a live session. Navigating back to OneDrive proved it in one command.
    //
    // So: classify, then act. An ERROR page is recovered from by navigating to
    // the deck, which is not meddling with a credential prompt because there is
    // no prompt. A real PROMPT is left completely alone and the round refuses
    // with a reason that names it, instead of "is the add-in open?" seven times.
    {
      const current = String(sh("tab-list") ?? "")
        .split("\n")
        .find((l) => /(current)/.test(l));
      const url = /((https?:[^)]+))/.exec(String(current ?? ""))?.[1];
      const kind = url ? authPageKind(url, sh("snapshot")) : "not-auth";
      if (kind === "error") {
        console.log("  the tab is on a sign-in ERROR page, not a prompt — navigating back to the deck");
        sh("goto", DECK_HOME_URL);
        await sleep(8000);
      } else if (kind === "prompt") {
        console.error("  the tab is on a LIVE sign-in prompt — leaving it untouched; this one needs the owner");
        return;
      }
    }
    const deckName = process.env.PW_DECK ?? DECK_NAME;
    // OPENING AND SELECTING ARE TWO JOBS, and conflating them was a regression
    // this function shipped with for one commit. Guarding the whole block on
    // "the deck tab is absent" meant a deck that WAS open but not fronted never
    // got selected — so `reload` below refreshed whatever tab happened to be
    // current, usually OneDrive's home, and every attempt then reported
    // `deck ? slide(s)` with the deck sitting one tab away.
    //
    // Click the file only when no tab holds it; select it either way.
    if (!sh("tab-list").includes(deckName)) {
      /**
       * LOOK AT THE FILE LIST FIRST — the link is only ON the file list.
       *
       * `refFor` searches whatever tab is CURRENT. When the deck is missing the
       * current tab is normally the other leg's presentation, and a presentation
       * page carries no `link "PresentationNN"` at all. So the search found
       * nothing, nothing was clicked, and this branch slept 25 seconds and gave
       * up — every attempt, silently, because an absent ref is not an error.
       *
       * Measured on the live browser, 2026-08-29, with the deck genuinely absent:
       *
       *     from the current tab (a presentation)   No matches found
       *     from the OneDrive tab                   Found 2 matches
       *
       * It cost a whole cycle tonight: seven attempts, `deck-missing` each time.
       * It had been invisible because this stop used to arrive with `wrong-size`
       * beside it — the size check read whichever deck WAS fronted — and
       * `wrong-size` is deliberately unrecoverable, so the cycle died on attempt
       * one and recovery never got this far. Fixing that guard this morning is
       * what exposed this.
       *
       * PREFER AN EXISTING TAB to navigating: the other leg's deck may be open in
       * another tab, and a `goto` on the current one would close a document this
       * cycle still needs. Navigate only when no tab is on the file list.
       */
      const home = String(sh("tab-list") ?? "")
        .split("\n")
        .find((l) => l.includes(DECK_HOME_URL));
      const homeIndex = home ? /(\d+):/.exec(home)?.[1] : null;
      if (homeIndex) sh("tab-select", homeIndex);
      else {
        sh("tab-new", DECK_HOME_URL);
        await sleep(8000);
      }
      const deckPattern = new RegExp(`link "${deckName}`);
      const deck = refFor(sh, deckName, deckPattern);
      if (deck) clickRef(sh, deck);
      // SAY WHICH WAY IT WENT. A ref that is simply absent looks exactly like a
      // click that did not take, and that is what hid this for as long as it hid.
      else console.error(`  the file list shows no \`${deckName}\` to open — recovery cannot reach it`);
      await sleep(25000);
    }
    const line = sh("tab-list")
      .split("\n")
      .find((l) => l.includes(deckName));
    const n = line ? /(\d+):/.exec(line)?.[1] : null;
    if (n) sh("tab-select", n);
    await sleep(20000);
  }
  const dialog = /dialog \[ref=([a-z0-9]+)\]/.exec(sh("find", "Sorry, we ran into a problem"))?.[1];
  if (dialog)
    sh(
      "eval",
      'el => { const b = [...el.querySelectorAll("button")].find(n => /^\\s*Refresh\\s*$/.test(n.textContent || "")); ' +
        'if (!b) return "no refresh"; b.click(); return "clicked"; }',
      dialog,
    );
  else sh("reload");

  const pane = await refreshPane(sh, sleep, { reloaded: true });

  sweepDeck(sh);
  return pane;
}

/**
 * Give the deck a FRESH PANE: reload the tab, reopen the pane, select Automation.
 *
 * WHY A ROUND NEEDS THIS AND A DECK SWEEP IS NOT ENOUGH. The pane's age when a
 * round starts is the best predictor of what that round reports. Rounds 110-123,
 * split on it:
 *
 *     fresh pane (<200s)   post-retry 0, 2, 0, 0, 0, 1, 0   deck mostly 16
 *     reused pane          post-retry 0, 5, 7, 3, 8, 7, 2   deck 45-97
 *
 * Mean 0.43 against 4.57. Sweeping the deck clears the SLIDES and leaves
 * whatever the pane itself accumulated, so the second round of every pair in
 * this archive has been a degraded sample — and that degradation was published
 * three times as a property of the profile, of the position, and of the
 * observer before anyone measured the pane's age.
 *
 * ONE IMPLEMENTATION, extracted from `recover` rather than copied beside it.
 * This repo has already paid for a second copy of a sweep: `recover` grew a
 * hardcoded deck name that went stale, and the fix was to have exactly one of
 * everything. `recover` still owns the browser-death and crash-dialog handling
 * that must happen BEFORE the reload; it hands the rest here.
 *
 * `reloaded` says whether the caller has already issued the reload — `recover`
 * has, because it must choose between a crash dialog's Refresh button and a
 * plain reload. Anyone else has not.
 */
export async function refreshPane(sh, sleep, { reloaded = false } = {}) {
  if (!reloaded) sh("reload");

  // A RELOAD RAISES A BEFOREUNLOAD MODAL, AND A MODAL BLOCKS EVERYTHING.
  //
  // PowerPoint asks "changes you made may not be saved" when a tab with unsaved
  // work is reloaded, and until that dialog is answered EVERY playwright-cli
  // command fails — `find` returns nothing at all (not even its miss message),
  // `tab-list` shows the tab with an empty title, and `screenshot` refuses with
  // "does not handle the modal state". The browser looks dead and is not.
  //
  // Calling this straight after `sweepDeck` is the worst possible moment: the
  // sweep just deleted slides, so there are ALWAYS unsaved changes. Round 124
  // wedged exactly there — reloaded, prompted, and sat behind the modal until a
  // human accepted it.
  //
  // `recover` has issued the same reload for months without hitting this,
  // because it runs when the page has usually already crashed and has no
  // beforeunload handler left to fire. That is luck, so the accept lives here,
  // on the shared path, and covers both callers.
  //
  // Unconditional and best-effort: accepting a dialog that is not there costs
  // one no-op call, and checking first would need a call that the modal blocks.
  sh("dialog-accept");
  await sleep(55000);

  // POLLED, NOT A SINGLE LOOK AFTER A FIXED SLEEP. Third time today this exact
  // shape has been wrong — `sideloadAddIn` after its upload, `commandPresent`
  // before deciding to sideload, and now here, in the function written to make
  // the NEXT round clean. A reloading Office tab does not repopulate its ribbon
  // on a schedule.
  const pane = await waitForRef(sh, sleep, "Insert chart", /button "Insert chart"/, RIBBON_WAKE_BUDGET_MS, 4000);
  if (pane) clickRef(sh, pane);
  await sleep(20000);

  // AUTOMATION, or the next round's readiness reads `verbose trace ?` and the
  // run button is not in the DOM. A pane always reopens on Chart.
  const automation = refFor(sh, "Automation", /tab "Automation"/);
  if (automation) clickRef(sh, automation);
  await sleep(5000);

  return Boolean(pane);
}

/**
 * Delete every slide but the first, through the pane's own Office.js context.
 *
 * Pulled out of `recover` so the readiness check can use it without the reload
 * and reopen that surround it there — see `onlyDirtyDeck`. Same call, same
 * budget, ONE implementation: a second sweep written beside this one is how the
 * hardcoded deck name and the three stale slogans in `triage.mjs` happened.
 */
export function sweepDeck(sh) {
  const anchor = refFor(sh, "Chart", /tab "Chart"/);
  if (!anchor) return false;
  // READ WHAT IT ANSWERED. This threw the result away and returned `true`
  // unconditionally, so "deck swept — the next round starts clean" printed
  // whether the sweep had cleaned the deck, failed outright (`deck-failed`), or
  // left slides behind — and the next round then refuses with `deck-dirty` for a
  // reason the previous round's output said could not have happened.
  //
  // `cleanDeckScript` returns `deck:N`, the slides remaining. One is clean: the
  // loop stops at index 1 on purpose, because a deck cannot have zero slides and
  // a fixed-count delete once took one to exactly that.
  const out = sh("eval", cleanDeckScript(90000), anchor);
  const left = /deck:(\d+)/.exec(out)?.[1];
  if (left === undefined) return false;
  return Number(left) <= 1;
}

/**
 * The last round's receipt, or null when there is not one worth having.
 *
 * Separate from `sessionPosition` so that one stays pure: a missing file, an
 * unreadable one and a half-written one are all "no previous round" here, and
 * none of them should stop a round from running.
 */
export function readReceipt(path = RECEIPT_PATH, read = readFileSync) {
  try {
    return JSON.parse(String(read(path, "utf8")));
  } catch {
    return null;
  }
}

/** Where the driver leaves its account of how the round ended. */
export const RECEIPT_PATH = ".round-outcome.json";

/**
 * Where the SESSION chain lives — and why it is not the receipt.
 *
 * `sessionPosition` chains through a file because each round is its own
 * process, and it chained through `RECEIPT_PATH`. That file has a second reader
 * with an incompatible lifetime: `cycle.mjs` clears it before EVERY leg, so its
 * "the driver left no receipt" branch reads THIS leg's outcome rather than the
 * last one's. Both readers are right about their own question, and one file
 * could not serve them both.
 *
 * SO THE INSTRUMENT WAS BEING DESTROYED BY THE ONLY SANCTIONED WAY TO RUN A
 * PAIR. Every leg of every cycle began as `sessionIndex: 1, 0m in`. The archive
 * says it plainly: of 52 consecutive rounds under 45 minutes apart, 33 record
 * the later one as the first of a fresh session, and `sessionElapsedMs` is 0 in
 * 76 of the 104 rounds that carry it — 89 minutes being the largest figure ever
 * recorded, for an effect indexed from 0 to 224.
 *
 * What that costs is not bookkeeping. `sessionDepthWarning` REFUSES a round from
 * the fifth of a session — measured: 10 skips across rounds 5-10 of a
 * back-to-back block against 0 across 9 rested ones — and a refusal needs an
 * index that counts. Pinned at 1 the stop is unreachable, `--deep` has nothing
 * to override, and a night of back-to-back rounds runs well past the point where
 * the host starts dropping scenarios, with nothing saying so.
 *
 * Two questions, two files. Nothing clears this one but a real gap in time.
 */
export const SESSION_PATH = ".round-session.json";

/**
 * Where the last round sat in its session, or null when there is no chain.
 *
 * Same shape and the same forgiveness as `readReceipt`: a missing, unreadable or
 * half-written file all mean "no previous round", and none of them should stop a
 * round from running.
 */
export function readSession(path = SESSION_PATH, read = readFileSync) {
  try {
    return JSON.parse(String(read(path, "utf8")));
  } catch {
    return null;
  }
}

/**
 * The round of a session at which scenarios start being skipped.
 *
 * Measured 2026-08-25, ten back-to-back rounds against nine rested ones:
 *
 *     BACK-TO-BACK (216-225)  10 rounds  10 skipped  [0 0 0 0 1 2 2 2 0 3]
 *     RESTED       (230-238)   9 rounds   0 skipped  [0 0 0 0 0 0 0 0 0]
 *
 * The first FOUR rounds of the back-to-back block skipped nothing. From the
 * fifth it never really stops. Nine rested rounds — each the first of its
 * session — skipped nothing at all.
 */
export const SESSION_DEPTH_WARN_AT = 5;

/**
 * A warning, never a refusal, for a round deep in a session.
 *
 * A skip is not a slow scenario: it is `the host stopped answering during this
 * scenario, so nothing was checked`. A round with three of them measured eleven
 * things instead of fourteen, and the three it lost are the heaviest — which is
 * also where all three archived crashes happened.
 *
 * NOT a refusal, deliberately. A deep round still produces a real sheet and
 * refusing one would throw away work someone asked for; what it cannot do is
 * quietly pass for a complete round. So this says which kind of round is about
 * to be read, and leaves the choice where it belongs.
 */
export function sessionDepthWarning(sessionIndex) {
  const n = Number(sessionIndex);
  if (!Number.isFinite(n) || n < SESSION_DEPTH_WARN_AT) return [];
  return [
    `round ${n} of this session — from about the fifth, scenarios start being SKIPPED ` +
      `(10 skips across rounds 5-10 of a back-to-back block; 0 across 9 rested rounds). ` +
      `A skip means nothing was checked, and the scenarios lost are the heaviest. ` +
      `Rest 45+ minutes for a complete round.`,
  ];
}

/**
 * How long a gap ends a SESSION of rounds.
 *
 * A round takes 11-18 minutes and the loop starts the next one immediately, so
 * consecutive rounds land 12-20 minutes apart. 45 minutes is comfortably past
 * that and comfortably short of "came back after lunch", which is the
 * distinction this needs to make and the only one it needs to make.
 */
export const SESSION_GAP_MS = 45 * 60 * 1000;

/**
 * WHERE IN A RUN OF ROUNDS THIS ONE SITS — and it is not cosmetic.
 *
 * Ten rounds on one build, 2026-08-24, every one `--fresh`:
 *
 *     minutes-in     0     65    116    186    224
 *     laterMed   19321  24423  40469  29023  40067
 *
 * The most repeated measurement in this harness roughly DOUBLES over two hours
 * of back-to-back rounds, and past ~90 minutes scenarios start being skipped
 * because the host stops answering. It is not the machine: `--fresh` rebuilds
 * the browser every round and 8.6 GB of 15.7 GB was free at the bottom of the
 * run.
 *
 * **No archived round records when it ran.** The file carries the BUILD stamp
 * and nothing else, so session position is unrecoverable from 190 rounds — and
 * it is worth 2x on the headline number. Every cross-round comparison this
 * archive has ever made is confounded by a variable nobody could see, and the
 * noise-floor note already had its finger on it without knowing the cause:
 * "the second run is usually the worse one, so a floor measured this way
 * includes an effect as well as noise."
 *
 * Chained through the receipt because each round is its own PROCESS — there is
 * no long-lived thing to hold a counter, and inferring position from file mtimes
 * works only until the files are committed and checked out again.
 *
 * Pure, and separate from the reading, so the arithmetic can be tested without
 * a filesystem or a clock.
 */
export function sessionPosition(prev, nowMs, gapMs = SESSION_GAP_MS) {
  const fresh = (at) => ({
    index: 1,
    sincePrevMs: null,
    startedAt: Number.isFinite(at) ? new Date(at).toISOString() : null,
    elapsedMs: 0,
  });
  const prevAt = Date.parse(prev?.at ?? "");
  if (!Number.isFinite(prevAt) || !Number.isFinite(nowMs)) return fresh(nowMs);
  const since = nowMs - prevAt;
  // A receipt from the FUTURE is a clock that moved, not a session. Treated as
  // a fresh start rather than trusted into a negative gap.
  if (since < 0 || since > gapMs) return fresh(nowMs);
  const prevIndex = Number(prev?.sessionIndex);
  // ELAPSED TIME, NOT JUST THE COUNT — and this is the number the finding is
  // actually indexed by. The 2026-08-24 drift table reads in MINUTES-IN
  // (0/65/116/186/224), not in rounds, because rounds vary from 684s to 1085s
  // and a count cannot tell a fast session from a slow one. Recording the index
  // alone would have measured a proxy for the variable and called it the
  // variable.
  const startedAt = Date.parse(prev?.sessionStartedAt ?? "");
  const start = Number.isFinite(startedAt) ? startedAt : nowMs;
  return {
    index: (Number.isFinite(prevIndex) && prevIndex > 0 ? prevIndex : 1) + 1,
    sincePrevMs: since,
    startedAt: new Date(start).toISOString(),
    elapsedMs: Math.max(0, nowMs - start),
  };
}

/**
 * What just happened, as structure rather than prose.
 *
 * The driver's exit code is BINARY — 0 for a round that finished, 1 for
 * everything else — so anything downstream that wants to know WHY a round
 * stopped has exactly two options: parse the console output, or be told. The
 * console output is prose written for a person at 2am and edited whenever a
 * message is improved, and `rounds-gate.mjs` already refused to parse prose once
 * for the same reason: a gate that reads sentences is a gate that breaks when
 * someone fixes a sentence.
 *
 * So the driver writes down what it knows. `reason` and `codes` are the same
 * values `shouldRetry` judges, which means a reader can apply `RECOVERABLE_STOPS`
 * itself rather than reimplementing the judgement — and there must only ever be
 * one implementation of "is this worth another attempt".
 *
 * `roundFile` matters as much as the reason. A caller told only that a round
 * finished would have to guess which file it produced, and "the newest file in
 * rounds/" is precisely the assumption that produced a wrong overwrite
 * diagnosis in this repo once already.
 *
 * Pure, and separate from the writing, so the shape can be tested without a
 * filesystem.
 */
export function outcomeReceipt({
  reason,
  codes,
  roundFile,
  build,
  size,
  threw,
  at,
  sessionIndex,
  sessionStartedAt,
  prev,
}) {
  // A ROUND THAT NEVER RAN DOES NOT EXTEND THE SESSION.
  //
  // The session exists to say how long the host has been under load, and it is
  // chained through this file because each round is its own process. A readiness
  // refusal runs no battery: it checks the pane, finds something wrong, and
  // stops. Stamping a fresh `at` for that moves the 45-minute gap forward from a
  // round that did nothing, and recording an index counts a round that never
  // happened into the drift table.
  //
  // It bit within hours of the `deep-session` stop landing. Round 256 refused as
  // "round 5 of this session", wrote `at: <now>, sessionIndex: 5`, and the next
  // attempt would have been round 6 — refused again, stamped again, climbing.
  // A caller that retried promptly would never have got a round out of it, and
  // the fix for the livelock is the same as the fix for the arithmetic: carry
  // the previous receipt's position forward instead of inventing a new one.
  //
  // ONLY when nothing ran. A round that finished, crashed, or wedged has worked
  // the host and owns its place in the session.
  // `not-ready` ALONE decides it. A readiness refusal never carries a round
  // file, so pairing the two reads as caution and is in fact unkillable — no
  // input can tell the two forms apart. `crashed` has no round file either and
  // must NOT hold, which is why the reason is the condition and the file is not.
  const ranNothing = reason === "not-ready";
  if (ranNothing && prev) {
    return {
      ...outcomeReceipt({ reason, codes, roundFile, build, size, threw, at, sessionIndex, sessionStartedAt }),
      at: prev.at ?? at ?? new Date().toISOString(),
      sessionIndex: prev.sessionIndex ?? null,
      sessionStartedAt: prev.sessionStartedAt ?? null,
      // Said in the file, because a receipt whose timestamp is older than the
      // process that wrote it is otherwise indistinguishable from a stale file.
      heldSessionFrom: "the previous round — this attempt ran none",
    };
  }
  return {
    reason: reason ?? null,
    // Always an array. A reader doing `codes.includes(...)` on a `finished`
    // round should get `false`, not a crash on undefined.
    codes: Array.isArray(codes) ? codes : [],
    roundFile: roundFile ?? null,
    build: build ?? null,
    // THE PROFILE STRING, which is what `readSlideSize` actually returns —
    // "16:9", "4:3", or "960x540" for anything else. This was written as
    // `{ width: size.width, height: size.height }` against a shape the driver
    // has never produced, so every cycle leg recorded `"size": {}` and the
    // field that says WHICH ARM a round belonged to said nothing at all. It
    // read as correct because its test passed an object in, which tested the
    // assumption rather than the driver.
    size: size ?? null,
    // Only on the path that has one — see the catch in `main`. Present means
    // the round failed in a way nothing anticipated, which is a different
    // thing from every named refusal and should not have to be inferred.
    ...(threw ? { threw } : {}),
    // Whether the reason is one recovery ADDRESSES, decided here by the same set
    // the driver retries on — never re-derived downstream.
    // THE SAME QUESTION `shouldRetry` ANSWERS, asked the same way. This read
    // only the codes, and most stops carry none: a crash, a silent host, a
    // wedge, a dead browser and an unexpected throw are all retried on their
    // REASON alone. So a night that recovered from a crash six times ended by
    // printing "crashed is not something recovery addresses — it needs a
    // person", directly under six lines saying "clearing the crash and starting
    // again". The receipt contradicted the console in the same output.
    recoverable: shouldRetry(reason, 0, 1, codes),
    at: at ?? new Date().toISOString(),
    // WHERE IN A RUN OF ROUNDS THIS ONE SAT. Chained here because each round is
    // its own process and there is nowhere else that outlives one — see
    // `sessionPosition`, and the 2x within-session drift that made it necessary.
    sessionIndex: sessionIndex ?? null,
    // The SESSION start, chained so every later round can say how many minutes
    // into the session it ran — the variable the drift is indexed by.
    sessionStartedAt: sessionStartedAt ?? null,
  };
}

/**
 * Block until the site serves the commit under test.
 *
 * WHY IT LIVES HERE. A round is worthless against a stale deploy — `readiness`
 * says so and stops — so every round is launched behind a wait loop, and that
 * loop has been retyped by hand at the shell each time. On 2026-08-22 the
 * hand-written one had TWO silent faults at once: it polled
 * `dannbleeker.github.io`, which GitHub Pages has 301'd to the custom domain
 * since July, and it read `.commit` from a document whose only field is
 * `build`. Either alone prints `?`; so does a deploy that has not landed. Ten
 * identical `?` lines are indistinguishable from patience, and the round sat
 * there for ten minutes after the deploy had already landed.
 *
 * The driver already knows the URL and already knows how to read the stamp.
 * A caller that has to re-derive either can get one of them wrong, and the
 * failure looks like waiting.
 *
 * Reports the served build every poll rather than a bare tick, so a poll that
 * cannot read anything is visibly different from one reading an old hash — the
 * distinction the `?` collapsed.
 */
export async function waitForDeploy(
  head,
  { fetchBuild = defaultFetchBuild, sleep, log = console.log, every = 20_000, max = 40 } = {},
) {
  log(`  waiting for ${head} to deploy`);
  for (let i = 1; i <= max; i++) {
    const served = buildOf(await fetchBuild());
    log(`    poll ${i}: ${served ?? "site did not answer with a build stamp"}`);
    if (served === head) {
      log("  deployed");
      return true;
    }
    if (i < max) await sleep(every);
  }
  // NOT A THROW. A round against a stale build is a bad round, and `readiness`
  // is the thing that decides that — it names the stale pane, it names the
  // hash, and `--retry` recovers from it. Dying here would replace a diagnosis
  // with a stack trace.
  log(`  gave up waiting after ${max} poll(s) — running anyway, and readiness will say if the site is behind`);
  return false;
}

async function main(argv, deps = {}) {
  const run = deps.run ?? spawnSync;
  const sleep = deps.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  const dirArg = argv.indexOf("--dir");
  const sh = cli(run, dirArg === -1 ? process.cwd() : argv[dirArg + 1]);
  const retryArg = argv.indexOf("--retry");
  const max = retryArg === -1 ? 0 : Number(argv[retryArg + 1]) || 0;

  const write = deps.write ?? writeFileSync;
  // EVERY STOP THAT FORCED A RETRY, in order. A round that needed two
  // recoveries before it could start is a round run against a host that was
  // already unwell, and until now the archive said nothing about it: round 148
  // took three attempts and then failed two scenarios that had never failed in
  // 109 rounds, and no archived field could connect those two facts.
  const recovered = [];
  // WHEN THIS ROUND BEGAN, and where in the session it sits. Read once, before
  // any attempt, so a round that needed three recoveries still reports the time
  // the work actually started rather than the time it finally succeeded.
  const startedAtMs = (deps.now ?? Date.now)();
  /**
   * THE SESSION CHAIN, FROM ITS OWN FILE — see `SESSION_PATH`.
   *
   * Read from the receipt until 2026-08-31, and `cycle.mjs` clears that before
   * every leg, so every leg of every cycle called itself round 1 of a fresh
   * session and the `deep-session` stop could never fire.
   *
   * The driver no longer reads the receipt AT ALL, which lint pointed out the
   * moment this moved — and that is the useful part: the session chain was the
   * receipt's only reader inside `main`. It is written here and read by
   * `cycle.mjs`. One question each, which is what the split was for.
   */
  const prevSession = (deps.readSession ?? readSession)();
  const session = sessionPosition(prevSession, startedAtMs);
  // A fresh browser before the first attempt, when asked for. See `startFresh`.
  if (argv.includes("--fresh")) {
    console.log("  closing the browser — this round starts from a fresh session");
    await startFresh(sh, sleep);
  }
  // Read ONCE, here, and handed to every attempt. See `attempt`'s `head`.
  const pinnedHead =
    deps.head ??
    (String(run("git", ["rev-parse", "--short=7", "HEAD"], { encoding: "utf8" }).stdout ?? "").trim() || null);

  // AFTER the head is pinned, so the wait and the round agree on what is under
  // test. See `waitForDeploy` for what a hand-rolled version of this cost.
  if (argv.includes("--wait-for-deploy") && pinnedHead)
    await (deps.waitForDeploy ?? waitForDeploy)(pinnedHead, { fetchBuild: deps.fetchBuild, sleep });

  for (let n = 0; ; n++) {
    if (n) console.log(`\n  attempt ${n + 1} of ${max + 1}`);
    // THE UNKNOWN FAILURE IS A FAILURE TOO. `attempt` is ~200 lines with one
    // try/catch in it, and everything it does not anticipate arrived here as an
    // unhandled rejection: the process died with a stack trace, `--retry`
    // covered none of it, and no receipt was written — so a night that had six
    // attempts left ended on the first one, and whatever ran it could not even
    // say why.
    //
    // Retried like any other reason, and bounded by the same `--retry` the
    // caller chose. A deterministic bug will simply fail `max` times and stop;
    // a transient one — the kind `recover` exists for — gets the same second
    // chance a crash dialog does.
    let outcome;
    try {
      // WHAT THE DRIVER HAD TO DO TO GET A ROUND AT ALL. See `driverRun`.
      //
      // AND HOW THE ROUND WAS LAUNCHED, which the archive could not say. On
      // 2026-08-22 six rounds established that a second-round-of-a-pair on an
      // aged pane refuses a group and one on a fresh pane does not — the whole
      // argument for running the second round with `--fresh`. Then the claim
      // could not be entered in the prediction ledger, because **nothing in the
      // round file records which arm a round was in.** `paneAgeAtStartSeconds`
      // gives the symptom; the flag that caused it was known only from the shell
      // history of whoever typed it.
      //
      // `recovered: ["not-ready:pane-closed"]` correlates — a `--fresh` round
      // closes the browser, so the pane comes back closed — but a correlate is
      // not the fact, and rounds 163 and 165 are labelled `--fresh` in this
      // project's journal on nothing but my word.
      //
      // The house defect wearing its plainest costume: an experiment whose ARM
      // was not recorded. Everything else about those rounds was.
      outcome = await attempt(
        argv,
        {
          ...deps,
          head: pinnedHead,
          // The override, if the caller asked for it. Recorded in `driverRun`
          // below as well as acted on here, for the reason the `--fresh` note
          // spells out: a round whose ARM is not in the file cannot be compared
          // with one that is, and "ran deep in a session" is an arm.
          allowDeepSession: deps.allowDeepSession ?? argv.includes("--deep"),
          driverRun: {
            attempts: n + 1,
            recovered: [...recovered],
            fresh: argv.includes("--fresh"),
            deep: argv.includes("--deep"),
            waitedForDeploy: argv.includes("--wait-for-deploy"),
            // WHEN, and WHERE IN THE SESSION. The archive had neither, and
            // without them "this build is slower" and "this round ran later in
            // the day" are the same sentence — see `sessionPosition`.
            startedAt: new Date(startedAtMs).toISOString(),
            sessionIndex: session.index,
            sincePrevRoundMs: session.sincePrevMs,
            sessionStartedAt: session.startedAt,
            sessionElapsedMs: session.elapsedMs,
            /**
             * WHICH DOCUMENT THIS LEG WAS TOLD TO RUN AGAINST.
             *
             * The other half of `driverDeck`, and not the same fact. That one
             * is a MEASUREMENT — the tab that was actually in front when the
             * evidence was collected. This one is the INSTRUCTION, and it
             * belongs beside `fresh`, `deep` and `waitedForDeploy` because like
             * them it says what the caller asked for rather than what happened.
             *
             * The two disagree in exactly the state the driver already refuses
             * on: `deck-missing`, where the named deck is not open and the round
             * would otherwise measure whichever document was. Recording only the
             * measurement would leave that disagreement invisible on every round
             * that did NOT trip the refusal — including any run where a cycle's
             * `PW_DECK` was misaimed and the right deck happened to be in front.
             *
             * Null when nothing was named, which is the ordinary path: the round
             * runs against whatever is fronted, and saying so is honest where
             * defaulting to `DECK_NAME` would invent an instruction nobody gave.
             */
            deck: process.env.PW_DECK ?? null,
          },
        },
        sh,
      );
    } catch (err) {
      console.error(`  the round threw where nothing expected it to: ${err?.message ?? err}`);
      outcome = { code: 1, reason: "threw", codes: [], threw: String(err?.message ?? err) };
    }
    const { code, reason, codes, roundFile, build, size, threw } = outcome;
    if (!shouldRetry(reason, n, max, codes)) {
      // ONLY THE OUTCOME THAT STANDS. Writing a receipt per attempt would leave
      // a caller reading the state of a round that recovery went on to fix,
      // which is the opposite of what the file is for.
      //
      // Best-effort, exactly like archiving: a round that ran is not undone by
      // a failed write, and turning one into a non-zero exit would make the
      // paperwork more important than the evidence.
      const receipt = outcomeReceipt({
        reason,
        codes,
        roundFile,
        build,
        size,
        threw,
        sessionIndex: session.index,
        sessionStartedAt: session.startedAt,
        // See `outcomeReceipt`: an attempt that ran no round keeps the previous
        // round's place in the session rather than taking one. Chained off the
        // SESSION file, which is where that place now lives.
        prev: prevSession,
      });
      try {
        write(RECEIPT_PATH, JSON.stringify(receipt, null, 2));
      } catch (err) {
        console.error(`  (could not write ${RECEIPT_PATH}: ${err?.message ?? err})`);
      }
      // AND THE CHAIN, SEPARATELY. Only the three fields `sessionPosition`
      // reads, so nothing else can grow a dependency on this file and give it a
      // second lifetime — which is the whole reason it exists apart from the
      // receipt. Written from the receipt so the "ran nothing holds its place"
      // rule has exactly one implementation.
      try {
        write(
          SESSION_PATH,
          JSON.stringify(
            { at: receipt.at, sessionIndex: receipt.sessionIndex, sessionStartedAt: receipt.sessionStartedAt },
            null,
            2,
          ),
        );
      } catch (err) {
        console.error(`  (could not write ${SESSION_PATH}: ${err?.message ?? err})`);
      }
      return code;
    }
    // NAMED FOR WHAT ACTUALLY HAPPENED. This said "clearing the crash" whatever
    // the reason was, and the moment the retry covered more than crashes it
    // started lying: round 047 refused on a dirty deck alone and was told a
    // crash was being cleared. A recovery line is read while someone is
    // debugging, and one that invents a crash sends them looking for it.
    console.log(`  ${recoveryFor(reason, codes)} — see docs/ROUNDS.md, "The wedge"`);
    // The REASON, recorded before the recovery that hides it. Once `recover`
    // succeeds the round looks like any other, and the only evidence it was ever
    // in trouble is this list.
    // THE CODES, NOT JUST THE BUCKET. `reason` for a readiness stop is always
    // the single word "not-ready", and four rounds running recorded exactly
    // that — which reads as "the host needed rescuing" and mostly is not.
    //
    // Round 153's first attempt refused for TWO reasons: `pane-stale`, because
    // a build had just been deployed and the pane still held the old one, and
    // `host-silent`, because the editing session had dropped while the browser
    // sat idle between rounds. The first is a property of how rounds are run
    // and says nothing about the host; the second is host health. Folded into
    // one word they cannot be told apart, and any claim about whether the host
    // is getting better or worse is drawn from a bucket.
    //
    // This is the residual-bucket mistake made inside a field added three
    // rounds ago to stop exactly that.
    recovered.push(recoveryLabel(reason, codes));
    await recover(sh, sleep);
  }
}

/**
 * Write the crash report beside the rounds, and never let it cost the driver.
 *
 * Named for what it is: the evidence outlives the tab only if something keeps
 * it. `crashes/` rather than `rounds/` on purpose — these are not rounds, and
 * everything downstream pools that directory.
 */
async function keepCrashEvidence(sh, at, write = writeFileSync) {
  try {
    const report = await collectCrashEvidence(sh, { at });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const path = `crashes/${stamp}.md`;
    write(path, report);
    console.error(`  the host's own account of it is in ${path}`);
  } catch (err) {
    console.error(`  (could not keep the crash evidence: ${err?.message ?? err})`);
  }
}

async function defaultFetchBuild() {
  // A network blip is a reading this tool could not take, not a reason to die
  // with a stack trace. Unguarded, one `getaddrinfo ENOTFOUND` took the whole
  // check down mid-investigation and printed nothing about the host, the pane or
  // the deck — every one of which had already been measured by then. `buildOf`
  // turns the empty string into `null` and `readiness` already says what a
  // missing build means.
  try {
    const r = await fetch(`https://ssf-chart.struktureretsundfornuft.dk/build.json?cb=${Date.now()}`);
    return await r.text();
  } catch {
    return "";
  }
}

/** Archive a downloaded round under the next number. See `rounds/README.md`. */
export function archive(
  logPath,
  dir = "rounds",
  read = readFileSync,
  write = writeFileSync,
  // THE DEFAULT, because the caller that mattered was passing the unsafe one.
  // `everyRoundEverFiled` was written after two different rounds were both
  // filed as 064, and it was wired to the `--archive` subcommand — the path a
  // person uses by hand, occasionally. `collectRound`, which archives EVERY
  // round automatically, went on passing `readdirSync`: the working tree alone,
  // blind to any round committed on a branch that is not checked out. That is
  // precisely the collision it was written to stop, on the path that runs a
  // hundred times more often.
  //
  // Defaulted rather than fixed at the call site so the next caller cannot make
  // the same omission. Tests pass their own lister and are unaffected.
  list = everyRoundEverFiled,
  expectBuild = null,
  /**
   * WHAT THE DRIVER MEASURED, from outside the pane.
   *
   * Rounds 115 and 116 filed themselves as `720x540` while running on a 960x540
   * deck. The driver had read live `PageSetup` twice and printed `slide size
   * 16:9 (want 16:9)` before each round; the pane, asked at an unlucky moment,
   * fell to its last rung and read a saved file PowerPoint had not updated yet.
   * BOTH NUMBERS EXISTED. Nothing compared them, so the round filed under a
   * profile it was not measured at and the guard that exists to prevent exactly
   * that passed.
   *
   * Recording the driver's reading beside the pane's is what makes the
   * comparison possible at all. Null when the driver could not read it — an
   * absent second opinion must never read as agreement.
   *
   * DEFAULTS TO `undefined`, NOT NULL, and the distinction is load-bearing
   * rather than pedantic. Null is now WRITTEN to the round file — it is how a
   * reader sees that the driver looked and the host would not answer — so a
   * default of null would stamp that same claim onto every caller that never
   * offered a reading, `--archive` by hand included. Passing null explicitly
   * still means what it always meant; omitting the argument now means nobody
   * asked, which is a third state the field did not used to have.
   */
  driverSize = undefined,
  /**
   * WHAT THE DRIVER HAD TO DO TO GET THIS ROUND.
   *
   * `{ attempts, recovered }` — how many times the driver tried, and the stop
   * reason behind each retry. A round that needed two recoveries before it
   * could start was run against a host that was already unwell, and that is a
   * property of the EVIDENCE, not of the driver.
   *
   * Round 148 is the case. It took three attempts — a silent host, then a
   * closed pane — and then failed two scenarios that had not failed once in 109
   * rounds, with no app-code change to explain it. "Was it a sick host?" is the
   * obvious question and the archive could not answer it for any of the 149
   * rounds on file, because a successful recovery leaves a round looking exactly
   * like one that never needed it.
   *
   * Null when unknown — an absent reading must never read as "no recoveries".
   */
  driverRun = null,
  /**
   * WHICH DOCUMENT WAS IN FRONT when this round's evidence was collected.
   *
   * Left `undefined` by default rather than null, because the two mean
   * different things and both will occur — see where it is written below.
   */
  driverDeck = undefined,
) {
  const round = stripImages(JSON.parse(read(logPath, "utf8")));
  const build = buildOf(round.build);
  if (!build) throw new Error("that file carries no build stamp — it is not a round log");
  // THE LOG MUST BE THE ROUND THAT JUST RAN, and this is the guard the twin
  // check cannot be. `Download run log` is DISABLED while a round is running, so
  // clicking it after one that wedged does nothing at all — and the previous
  // round's file is still sitting at the same path, waiting to be filed under a
  // new number. That happened: round 039 was archived byte-identical to 038, a
  // whole round of evidence that never took place.
  //
  // The twin check below catches that case only when the two logs are IDENTICAL.
  // A stale log that merely differs — an older round, a different build — sails
  // past it. Comparing the log's own build stamp to the pane's is what closes
  // it, and it costs nothing: the driver already read the stamp to decide the
  // round was worth running.
  if (expectBuild && build !== expectBuild)
    throw new Error(
      `that log is build ${build} and the pane is serving ${expectBuild} — it is the PREVIOUS round's file, ` +
        "which is what sits at that path when a round wedges and the download button does nothing. Nothing archived.",
    );
  // THE DOWNLOAD IS A FILE ON DISK THAT IS ONLY SOMETIMES REPLACED. `Download
  // run log` is disabled while a round is running, so clicking it after a round
  // that WEDGED does nothing at all — and the previous round's log is still
  // sitting at the same path, waiting to be archived a second time under a new
  // number. That happened: round 039 was filed as byte-identical to 038, a
  // whole extra round of evidence that never took place, and only a checksum
  // caught it.
  //
  // A fabricated round is the worst thing this directory can hold. Everything
  // downstream pools these files — verdict histories, the rasterise arms, the
  // scenario flip detector — so one duplicate quietly doubles the weight of
  // whatever the real round happened to say.
  const body = `${JSON.stringify(round, null, 2)}\n`;
  const twin = list(dir)
    .filter((f) => /^\d{3}-.*\.json$/.test(f))
    .find((f) => {
      // A NAME THIS LISTER RETURNS NEED NOT BE A FILE. `list` defaults to
      // `everyRoundEverFiled`, whose entire purpose is to name rounds that are
      // NOT in the working tree — committed on a branch nobody has checked out —
      // so that `nextRoundNumber` below cannot hand their number out twice. That
      // function only ever inspects the NAMES. This check opens them, and it
      // inherited the new lister as collateral when the default changed.
      //
      // Unguarded, that threw ENOENT on the healthy path and only there: `.find`
      // walks the on-disk names first and reaches a git-only name only when
      // nothing matched, which is exactly the case of a genuinely new round. The
      // guard was perfectly inverted — a duplicate was caught, a real round
      // crashed — and `collectRound` swallowed the throw, so the driver exited 0
      // reporting a finished round with nothing written to rounds/ and the next
      // leg's download then overwrote the evidence.
      //
      // A file that cannot be opened is not a twin of anything.
      try {
        return read(`${dir}/${f}`, "utf8") === body;
      } catch {
        return false;
      }
    });
  if (twin)
    throw new Error(
      `that log is byte-identical to ${twin} — the pane never wrote a new one, ` +
        "which is what happens when the round did not finish. Nothing archived.",
    );
  // THE NUMBER COMES FROM WHAT `list` CAN SEE, and on 2026-08-16 that was not
  // everything. Round 064 was archived on `main`, its findings committed to a
  // branch, and `git checkout main` then REMOVED the file from the working tree
  // — it is tracked only on the branch. The next round was archived from `main`,
  // where the directory ends at 063, and was numbered 064 as well.
  //
  // Nothing was overwritten (`nextRoundNumber` is max+1, so it never lands on a
  // file it can see) and no evidence was lost. What it produces is two DIFFERENT
  // rounds both called 064 in two git contexts, which collides the moment either
  // is merged — and a pooled report that silently reads one of them twice, or
  // not at all.
  //
  // So the caller passes what GIT knows as well as what is on disk; see the
  // `--archive` branch below. `archive` itself stays pure and takes the union.
  const name = `${nextRoundNumber(list(dir))}-${build}.json`;
  // TWO spaces, because prettier checks this directory and every round archived
  // at one space failed the gate until someone reformatted it by hand.
  // STAMPED, NOT MERGED. It goes in its own field so the pane's own reading is
  // left exactly as the pane reported it — an archive that quietly corrected
  // itself would destroy the evidence that the two ever disagreed, which is the
  // only thing that makes the disagreement findable.
  //
  // AND WRITTEN EVEN WHEN THE DRIVER COULD NOT READ IT. `if (driverSize)`
  // dropped the field entirely on a failed read, so a round nobody had verified
  // the arm of looked exactly like one from a build that never checked: 3 of the
  // last 120 archived rounds carry no `driverSlideSize`, and nothing on disk
  // says whether that means "looked and could not tell" or "never looked".
  //
  // That matters more here than almost anywhere, because the pane's own reading
  // is the one known to be wrong: on rounds 115 and 116 the pane fell to its
  // last rung and reported a saved 720x540 file for a deck the driver had twice
  // measured live at 960x540. When the driver's reading is missing, the only
  // number left is the one that has already lied once. A reader must be able to
  // see that, so `null` is written and the third state — the field absent —
  // keeps its own meaning: an older build, or `--archive` run by hand.
  if (driverSize !== undefined) round.driverSlideSize = driverSize ?? null;
  /**
   * WHICH DOCUMENT THIS ROUND RAN AGAINST.
   *
   * Read from the tab that is actually in front rather than from `PW_DECK`,
   * because `PW_DECK` says only what was ASKED for and is null on the ordinary
   * path — and the whole point of this field is to be true on the days nobody
   * was being careful.
   *
   * NOTHING IN THIS ARCHIVE HAS EVER RECORDED IT. 0 of 408 files under
   * `rounds/` and `crashes/` name a document. So deck identity and slide
   * profile are perfectly confounded: every 4:3 round ran one file and every
   * 16:9 round another, and the strongest contrast this repo owns — 4:3
   * crashing on 52 of 73 attempts against 0 of 51 at 16:9, over the 25 builds
   * that ran both arms — is equally consistent with "that one file is sick".
   * A whole line of enquiry cannot be settled from the archive for want of one
   * string per round.
   *
   * No round already filed can be repaired. This exists so the next comparison
   * is not born confounded too.
   *
   * WRITTEN EVEN WHEN NULL, against the `driverSlideSize` line above rather
   * than with it. Absent and null are different facts here and both will occur:
   * absent means a build that never looked, null means this build looked and
   * the tab list would not say. Collapsing them is this repo's most repeated
   * defect — unknown printed as nothing — and the field is new, so there is no
   * back-compatibility reason to inherit it.
   *
   * `undefined` is the third state and it is why this is not a `?? null`: it
   * means the caller never offered a reading. `--archive` by hand does not know
   * which tab was in front, and stamping `null` there would claim a look that
   * nobody took — the same collapse this field exists to avoid, one level up.
   */
  if (driverDeck !== undefined) round.driverDeck = driverDeck;
  if (driverRun) round.driverRun = driverRun;
  write(`${dir}/${name}`, JSON.stringify(round, null, 2) + "\n");
  return name;
}

/**
 * Every round this repo has ever filed — on disk AND in git — so the next number
 * cannot reuse one that is committed on a branch nobody has checked out.
 *
 * `readdirSync` alone sees the WORKING TREE, and a round's file lives on the
 * branch its findings were committed to until that merges. Check out `main` and
 * the file disappears; archive from there and the number is handed out twice.
 * That happened on 2026-08-16 and produced two different rounds both called 064.
 *
 * `git ls-files` covers the current branch's index; `git log --all` covers every
 * branch, which is what actually matters here — the collision is with a round
 * committed somewhere else. Falls back to the directory alone if git is not
 * available or the call fails, because refusing to archive a real round over a
 * numbering nicety would be the worse trade.
 */
function everyRoundEverFiled(dir) {
  const onDisk = readdirSync(dir);
  try {
    const inGit = spawnSync("git", ["log", "--all", "--name-only", "--pretty=format:", "--", dir], {
      encoding: "utf8",
      maxBuffer: 64e6,
    });
    if (inGit.status !== 0 || !inGit.stdout) return onDisk;
    const names = inGit.stdout
      .split("\n")
      .map((l) => l.trim())
      .filter(Boolean)
      .map((l) => l.slice(l.lastIndexOf("/") + 1));
    return [...new Set([...onDisk, ...names])];
  } catch {
    return onDisk;
  }
}

if (isMain(import.meta.url, process.argv[1])) {
  const argv = process.argv.slice(2);
  if (argv[0] === "--archive") {
    if (!argv[1] || !existsSync(argv[1])) {
      console.error("usage: node scripts/round.mjs --archive <ssf-charts-round.json>");
      process.exit(2);
    }
    console.log(`archived as rounds/${archive(argv[1], "rounds", readFileSync, writeFileSync, everyRoundEverFiled)}`);
    process.exit(0);
  }
  process.exit(await main(argv));
}
