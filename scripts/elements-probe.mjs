/**
 * THE FIVE ELEMENTS BUTTONS THE MANIFEST SHIPS, AGAINST A REAL HOST.
 *
 * `manifest.xml` ships five ribbon deep links — `?tab=elements&el=harvey`,
 * `check`, `flow`, `kpi`, `table` — and **no scenario in `selftest.ts` touches
 * Elements at all**: grep it for harvey/kpiCard/elements and the count is 0. So
 * the 19-of-19 a round prints and the Elements code have never once overlapped.
 *
 * That matters now rather than in the abstract. Elements is the ONLY
 * Office.js-reaching behaviour change between the last round-validated build
 * before the release (`454303d`, 2026-09-10) and v0.6.1 itself: `elements.ts`
 * gained `desc` on all five builders, and `scene.desc` flows
 * `-> opts.altText -> Shape.altTextDescription` at three call sites in
 * `powerpoint.ts`. `altTextDescription` is a PowerPointApi **1.10** write, gated
 * on `supports("1.10")`.
 *
 * One behaviour change, on a 1.10 write, on the buttons a certification tester
 * clicks first, with zero host evidence. Rounds 455-457 were all green and none
 * of them asked.
 *
 * WHAT IT ASSERTS, per element:
 *   1. the pane's own button for it exists and is enabled
 *   2. clicking it lands at least one shape on the slide
 *   3. something that landed carries a NON-EMPTY `altTextDescription`
 *   4. and that text describes THIS element rather than a chart
 *
 * (3) is the new behaviour; (4) is what stops it passing against a group left
 * over from an earlier insert.
 *
 * IT DRIVES THE REAL BUTTONS, not the engine. Calling `buildHarvey` in-process
 * would test the same `desc` strings this file already imports its expectations
 * from, and prove nothing about the deep links, the pane wiring or the 1.10
 * write. The buttons all share the accessible name "Insert", so they are
 * clicked by DOM id inside the pane's frame.
 *
 * IN `scripts/`, NEVER `src/`: it changes nothing a user receives and is not in
 * the vite bundle, so it can land mid-window without moving the deployed build.
 *
 * Usage:  node scripts/elements-probe.mjs
 * Exit:   0 all five carried their description · 1 at least one did not
 *         · 2 could not ask (no pane, no frame, or a host that never answered)
 */

import { spawnSync } from "node:child_process";
import { isMain } from "./is-main.mjs";

const CLI = process.env.PLAYWRIGHT_CLI_JS;
const DIR = "C:\\devtools\\SSF-Charts\\.pw-session";

/**
 * The five the manifest ships, with what their `desc` should say.
 *
 * The patterns are deliberately loose — `elements.ts` builds these from
 * sanitised user input ("Harvey ball, 75% filled.", "Table, 2 rows by 3
 * columns.") and pinning the exact string here would make this a change
 * detector for copy rather than a check that the description ARRIVED.
 */
/**
 * The name the renderer gives the group it forms, and where the description is
 * written. A LOOKUP KEY rather than a label — `powerpoint.ts` and `ooxml.ts`
 * each hold their own private copy, so this third one is checked against the
 * source by the test beside this file rather than trusted.
 */
export const GROUP_NAME = "PowerChart";

/** How long an element gets to land before the probe stops waiting for it. */
export const SETTLE_BUDGET_MS = 60_000;

/** Between reads. Short enough to catch a fast insert, long enough not to flood the host. */
export const SETTLE_POLL_MS = 4_000;

/**
 * Wait for the PANE to say it has finished, then read the slide ONCE.
 *
 * THE PREVIOUS VERSION OF THIS FUNCTION BROKE THE THING IT MEASURED, and that
 * is the whole reason this one looks so plain. It polled the slide every four
 * seconds THROUGH the draw — a `PowerPoint.run` per poll, interleaved with the
 * renderer's own batches, on a host that forces a full presentation save on
 * every sync (office-js#6329). Measured 2026-09-18, same fresh session, same
 * cleared slide, same element, the only difference being the polling:
 *
 *     table insert, no polling      -> "Done."
 *     table insert, polled every 4s -> "Failed: PowerPoint did not respond
 *                                       while drawing shapes 11-20 of 23 (45s)"
 *
 * and the poll results in the failing run read `X,20,20,20,20,...` — the first
 * read timing out against the draw, then the shape count frozen at 20 forever.
 * The other four Elements survived it because their draws are short enough; the
 * table's is long enough to be hit over and over. Every "table FAILS" this probe
 * ever reported was this.
 *
 * So: no host calls until the insert is done. `busy` is a DOM read of
 * `#host-note`'s class, which the pane sets to `status-busy` while working and
 * to `status-ok`/`status-err` when it settles — it costs the host nothing, and
 * it is a better signal than any amount of guessing at quiescence, which is
 * what the stable-id counting and the 24s floor were.
 */
export async function settleReads(read, busy, sleep, opts = {}) {
  const budgetMs = opts.budgetMs ?? SETTLE_BUDGET_MS;
  const pollMs = opts.pollMs ?? SETTLE_POLL_MS;
  const now = opts.now ?? (() => Date.now());
  const started = now();
  let settled = false;
  while (now() - started < budgetMs) {
    await sleep(pollMs);
    // `null` is "cannot tell" — the note was unreadable. Keep waiting rather
    // than treating an unknown as a finish.
    if ((await busy()) === false) {
      settled = true;
      break;
    }
  }
  return { after: await read(), waitedMs: now() - started, settled };
}

export const ELEMENTS = [
  { el: "harvey", expect: /harvey/i },
  { el: "check", expect: /checkbox/i },
  { el: "flow", expect: /process flow/i },
  { el: "kpi", expect: /kpi|\d/i },
  { el: "table", expect: /table/i },
];

const pw = (...args) => {
  const r = spawnSync(process.execPath, [CLI, "-s=ms", "--raw", ...args], {
    encoding: "utf8",
    cwd: DIR,
    maxBuffer: 64e6,
    timeout: 180_000,
  });
  if (r.error) throw r.error;
  return String(r.stdout ?? "") + String(r.stderr ?? "");
};

/**
 * Read every shape on the selected slide, with its alt text.
 *
 * `altTextDescription` is named in the load rather than taken off `items`: the
 * bare collection read is the one this repo has measured coming back SHORT
 * (`shapes-items-count-honest` answers `short-3` every round), and a shape
 * missing from a short read would be reported here as the product failing to
 * write a description it wrote perfectly well.
 */
export const readAltScript = (budgetMs = 20000) =>
  "async () => { try { return await Promise.race([ " +
  "PowerPoint.run(async (c) => { const s = c.presentation.getSelectedSlides().getItemAt(0); " +
  'const sh = s.shapes; sh.load("items/id,items/name,items/altTextDescription"); await c.sync(); ' +
  'return "alt:" + JSON.stringify((sh.items || []).map((x) => ({ id: x.id, name: x.name, alt: x.altTextDescription || "" }))); }), ' +
  `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
  '} catch (e) { return "alt-failed:" + (e && e.message ? String(e.message).slice(0, 90) : "?"); } }';

/**
 * What the PANE says about what just happened — `#host-note`, the `note()`
 * target, `role="status"`.
 *
 * THE INSTRUMENT THAT DECIDES, and this probe did not read it. When an insert
 * leaves loose parts and no group, a shape list cannot say whether the host
 * refused to group or the draw stopped before reaching the grouping sync — both
 * leave exactly the same slide. The pane knows: it writes `PowerPoint did not
 * respond while drawing shapes 1-10 of 23 (45s)` for the second, and something
 * quite different for the first.
 *
 * On 2026-09-17 the `table` element produced one sync's worth of a 23-shape
 * draw, the evidence file recorded that as a grouping refusal, and the sentence
 * correcting it said the deciding instrument "was not read". This is that
 * instrument. A pure DOM read, so it answers even while the host is silent —
 * which is exactly when it is worth the most.
 */
export const readNoteScript = () =>
  "() => { const n = document.getElementById('host-note'); " +
  "return 'note:' + JSON.stringify({ text: n ? n.textContent : null, cls: n ? n.className : null }); }";

/**
 * Is the pane still working? A DOM read, costing the host NOTHING.
 *
 * `note()` sets `#host-note`'s class to `hint status-busy` while an action runs
 * and to `status-ok`/`status-err` when it settles, so the class answers "has
 * the insert finished" without a single `PowerPoint.run`. That is the whole
 * point — see `settleReads` for what asking the HOST during a draw did.
 */
export const busyScript = () =>
  "() => { const n = document.getElementById('host-note'); " +
  "return 'busy:' + (n ? (/status-busy/.test(n.className) ? 'yes' : 'no') : 'unknown'); }";

/** `busy:yes|no` -> true/false, or null when the pane could not be read. */
export function readBusy(out) {
  const m = /busy:(yes|no|unknown)/.exec(String(out ?? ""));
  if (!m || m[1] === "unknown") return null;
  return m[1] === "yes";
}

/** `note:{...}` -> the note, or null when the pane would not answer. */
export function readNote(out) {
  const m = /note:(\{.*?\})\s*"?\s*$|note:(\{.*\})/s.exec(String(out ?? ""));
  const body = m?.[1] ?? m?.[2];
  if (!body) return null;
  for (const candidate of [body, body.replace(/\\"/g, '"')]) {
    try {
      const parsed = JSON.parse(candidate);
      if (parsed && typeof parsed === "object") return parsed;
    } catch {
      // try the next spelling — see `readAlt` for why both are attempted
    }
  }
  return null;
}

/**
 * What the note is stamped with between the click and the pane's own verdict.
 * If this is ever what the probe reports, the pane never answered at all.
 */
export const CLICK_SENTINEL = "waiting for this insert";

/**
 * Click one Elements button by id, inside the pane's frame.
 *
 * IT STAMPS THE NOTE BUSY BEFORE IT CLICKS. `#host-note` is a single slot
 * holding one verdict, and the pane only sets `status-busy` once its own click
 * handler reaches that line. Between the click and that moment the note still
 * carries the PREVIOUS element's `status-ok` — so `settleReads`, polling
 * `SETTLE_POLL_MS` later, can read a success belonging to the element before
 * this one and conclude this insert finished before it started. It would then
 * read the slide mid-draw and report the element as failed.
 *
 * That is the false red this probe was rewritten to stop, arriving by a second
 * route: the first was polling the HOST during the draw, this one is reading a
 * stale verdict from the PANE. The archived 5-of-5 run is not in doubt — a
 * stale settle reads the previous element's shapes, which fail this element's
 * `expect` regex, so this race can only manufacture failures and none were
 * recorded. It is fixed because it is latent, not because it fired.
 *
 * Found 2026-09-23, when `scripts/store-shots.mjs` hit the identical race and
 * was caught by its own guard holding a chart of 80 loose shapes that had not
 * finished grouping.
 */
export const clickElementScript = (el) =>
  "async () => { const tab = document.querySelector('[data-tab=\"elements\"]'); if (tab) tab.click(); " +
  "await new Promise((r) => setTimeout(r, 400)); " +
  `const b = document.getElementById(${JSON.stringify(el + "-insert")}); ` +
  'if (!b) return "no-button"; if (b.disabled) return "disabled"; ' +
  "const n = document.getElementById('host-note'); " +
  `if (n) { n.className = "hint status-busy"; n.textContent = ${JSON.stringify(CLICK_SENTINEL)}; } ` +
  'b.click(); return "clicked"; }';

/**
 * `alt:[...]` -> the array, or null when the host refused.
 *
 * THE CLI ESCAPES THE QUOTES IT PRINTS. `--raw` returns the eval's string with
 * `"` as `\"`, so the payload arrives as `[{\"id\":\"1\",...}]` and `JSON.parse`
 * throws on it. Unescaped on 2026-09-17 after all five elements reported "the
 * host would not list the slide" while the host was answering perfectly — the
 * same slide read back four described shapes when asked by hand.
 *
 * That is this probe's own third self-inflicted wound in one session, and the
 * worst of them: the other two refused to measure, and this one reported a
 * FAILURE against a product that was working. A parser that cannot read the
 * answer says the same thing as a feature that was never written.
 */
export function readAlt(out) {
  const m = /alt:(\[.*?\])\s*"?\s*$|alt:(\[.*\])/s.exec(String(out ?? ""));
  const body = m?.[1] ?? m?.[2];
  if (!body) return null;
  for (const candidate of [body, body.replace(/\\"/g, '"')]) {
    try {
      const parsed = JSON.parse(candidate);
      if (Array.isArray(parsed)) return parsed;
    } catch {
      // try the next spelling
    }
  }
  return null;
}

/**
 * The verdict for one element. Pure, so the part a test can reach is reachable
 * without a host — which is most of what goes wrong in a probe like this.
 */
export function judge({ el, expect }, before, after, clicked) {
  if (clicked !== "clicked") return { el, ok: false, cause: "button", why: `the pane's button said ${clicked}` };
  // `cause` and `silent` rather than a prose match: what the exit code and the
  // report turn on must not be sentences somebody can reword. See `summarise`.
  if (!Array.isArray(before) || !Array.isArray(after))
    return { el, ok: false, cause: "silent", silent: true, why: "the host would not list the slide" };
  const had = new Set(before.map((s) => s.id));
  const added = after.filter((s) => !had.has(s.id));
  if (!added.length) return { el, ok: false, cause: "nothing-landed", why: "nothing landed on the slide" };
  const described = added.filter((s) => String(s.alt ?? "").trim());
  if (!described.length) {
    // TWO DIFFERENT FINDINGS WEAR THIS FACE, and saying only "no description"
    // reports the wrong one. The description is written on the GROUP. When the
    // host refuses to group — which the archive has at roughly a quarter of
    // rounds — the parts land loose (`harvey-fill-f0`, `step-2`, `cell-text-1-3`)
    // and there is nothing to write it on. That is the grouping refusal, not a
    // regression in the 1.10 write, and the fix for it is in a different file.
    //
    // Measured on 2026-09-17: a slide holding four correct descriptions on four
    // `PowerChart` groups, and beside them sixty loose parts from inserts that
    // never grouped. This probe called that 5 of 5 failing to describe.
    // AND "NEVER GROUPED" IS NOT SOMETHING THIS CAN SEE. The first version of
    // this branch said it was, and the evidence file said it after that: loose
    // parts with no group were reported as the grouping refusal. They are
    // equally a draw that STOPPED — `SHAPES_PER_SYNC` is 10, so a 23-shape
    // element whose host went quiet after the first batch leaves exactly this.
    // The table element's ten parts on 2026-09-17 were read as a refusal and
    // were the first sync of a draw the pane had already reported dying in.
    //
    // So the verdict names the OBSERVATION — no group formed — and lists the
    // two causes without choosing. The pane's own status line is what chooses,
    // and this probe does not read it.
    const grouped = added.some((s) => s.name === GROUP_NAME);
    return {
      el,
      ok: false,
      cause: grouped ? "undescribed" : "ungrouped",
      ungrouped: !grouped,
      why: grouped
        ? `${added.length} shape(s) landed and grouped, and the group carries no altTextDescription`
        : `${added.length} loose shape(s) landed and NO GROUP FORMED, so there was nowhere to write the description — either the host refused to group, or the draw stopped before it got there. Read the pane's status line to tell those apart; this cannot.`,
      added: added.length,
    };
  }
  const match = described.find((s) => expect.test(s.alt));
  if (!match)
    return {
      el,
      ok: false,
      cause: "mismatched",
      why: `described, but not as a ${el}: ${JSON.stringify(described.map((s) => s.alt).slice(0, 2))}`,
    };
  return { el, ok: true, alt: match.alt, added: added.length };
}

/**
 * Did the host answer at all? A SILENT HOST IS NOT A FAILING PRODUCT, and this
 * file said it was.
 *
 * `Office` and `PowerPoint` can both be loaded objects in a pane that is on
 * screen and answering DOM questions, while `PowerPoint.run(...)` never
 * resolves — the host has stopped answering. That state is all over this week's
 * archive: 45s stalls blind-skipped a scenario in round 458 and another in 459,
 * and it defeated the `table` element twice.
 *
 * On 2026-09-17 this probe met it and printed "0 of 5 elements carried an
 * altTextDescription", exit 1, against a product measured working hours
 * earlier. Exit 1 means "asked and the answer was wrong". Nothing was asked.
 *
 * So when EVERY element came back silent the run exits 2 — could not ask —
 * rather than accusing the code. The distinction is the one this repo keeps
 * paying to relearn: a check that cannot tell "verified" from "not attempted"
 * reads as evidence.
 *
 * ALL of them, not some: one silent element among four answered ones is a real
 * result about that element, and downgrading the whole run for it would hide a
 * genuine failure behind an environment excuse.
 */
export const hostSilent = (results) => results.length > 0 && results.every((r) => r.silent === true);

/**
 * The run's three buckets and the exit code they add up to.
 *
 * Pure, and separate from the printing, because this is the part that was
 * WRONG: a silent element was counted in the denominator of "N of 5 carried an
 * altTextDescription" and pushed the exit to 1. Both said the product had been
 * asked and had answered badly.
 *
 *   `carried`  asked, and the description arrived   -> nothing owed
 *   `failed`   asked, and the answer was wrong      -> exit 1, a finding
 *   `unasked`  the host never answered              -> exit 2, a re-run
 *
 * A finding outranks a stall: one real failure among four stalls is still a
 * failure, and exiting 2 there would file a regression as an environment
 * excuse. A stall outranks a clean pass, because "4 of 5 passed and the fifth
 * was never asked" is not a green run — it is a partial one, and 0 would say
 * the fifth was fine.
 */
export function summarise(results) {
  const of = (...causes) => results.filter((r) => causes.includes(r.cause)).map((r) => r.el);
  const carried = results.filter((r) => r.ok).map((r) => r.el);
  const unasked = of("silent");
  const failed = results.filter((r) => !r.ok && !r.silent).map((r) => r.el);
  // Subsets of `failed`, not extra buckets — every one of these is still an
  // element that did not arrive described. They are named apart because they
  // send the reader to different files, and on 2026-09-17 a `nothing landed`
  // was reported as "grouped and not described", which points at the 1.10 write
  // for a click whose shapes never appeared at all.
  return {
    carried,
    failed,
    ungrouped: of("ungrouped"),
    undescribed: of("undescribed", "mismatched"),
    absent: of("nothing-landed"),
    // Kept apart from `absent`: "the pane would not offer the button" and "the
    // click was taken and nothing followed" were printed under one line that
    // asserted the click had been accepted, which is false of the first.
    unclickable: of("button"),
    unasked,
    code: failed.length ? 1 : unasked.length ? 2 : 0,
  };
}

async function main() {
  if (!CLI) {
    console.error("PLAYWRIGHT_CLI_JS is not set — nothing below was measured.");
    process.exit(2);
  }
  // ANY OF THE PANE'S OWN TABS WILL DO, because the ref is only used to say
  // WHICH FRAME to evaluate in — it is never clicked or read.
  //
  // This anchored on `tab "Chart"` alone, copied from `round.mjs`'s readiness
  // check, and found nothing on 2026-09-17 against a pane that was open and
  // healthy: the tabs on screen were `Elements`, `Agenda`, `Automation`. The
  // driver's anchor works for the driver and is not a general way to find this
  // pane. Reported as `no pane on screen`, which was false and would have sent
  // the next reader to restart a browser that was fine.
  //
  // `Elements` first because it is the tab this probe needs open anyway.
  // TWO TRAPS HERE, BOTH DOCUMENTED IN `round.mjs`'s `refFor`, AND THIS FILE
  // WALKED INTO BOTH on 2026-09-17 by copying the SHAPE of that call without
  // reading what it does.
  //
  //   1. `find` takes PLAIN TEXT. `--regex '/tab "Elements"/'` matches nothing,
  //      while `find "Elements"` returns the line. The driver passes a JS regex
  //      as a SECOND argument to `refFor` and applies it to the output itself;
  //      it is not a CLI flag.
  //   2. Take the ref from the MATCHING LINE, never the first ref in the
  //      output. `find` prints the whole frame hierarchy above its hit, so the
  //      first ref is the OUTER iframe — evaluating against it lands in the
  //      OneDrive document, where `PowerPoint` is undefined, and every question
  //      after that gets a confident wrong answer.
  //
  // Either bug alone produces "no pane" or "nothing landed" against a pane that
  // is open and working, which is indistinguishable from the defect this probe
  // exists to report.
  const anchors = [/tab "Elements"/, /tab "Chart"/, /tab "Automation"/, /tab "Agenda"/];
  let ref = null;
  for (const pattern of anchors) {
    const line = pw("find", String(pattern.source).replace(/^tab "|"$/g, ""))
      .split("\n")
      .find((l) => pattern.test(l));
    ref = line ? (/ref=([a-z0-9]+)/.exec(line)?.[1] ?? null) : null;
    if (ref) break;
  }
  if (!ref) {
    console.error(
      `no pane frame found — looked for ${anchors.map((a) => a.source).join(", ")}. Open the deck and the add-in first. Nothing was measured.`,
    );
    process.exit(2);
  }
  // PROVE THE FRAME BEFORE TRUSTING IT. A ref that resolves to the wrong frame
  // answers every later question with a plausible wrong answer — no Office, no
  // buttons, and a verdict that reads as "the product did not insert anything".
  const probe = pw(
    "eval",
    "() => (typeof PowerPoint !== 'undefined' ? 'office' : 'no-office') + '|' + (document.getElementById('harvey-insert') ? 'pane' : 'no-pane')",
    ref,
  );
  if (!/office\|pane/.test(probe)) {
    console.error(
      `the frame found is not the add-in pane (${/(\w+\|\w+)/.exec(probe)?.[1] ?? "?"}). Nothing was measured.`,
    );
    process.exit(2);
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  /** `settleReads` bound to this run's frame. The reasoning lives on it. */
  const settle = () =>
    settleReads(
      () => readAlt(pw("eval", readAltScript(), ref)),
      () => readBusy(pw("eval", busyScript(), ref)),
      sleep,
    );
  /**
   * EACH ELEMENT ON A SLIDE OF ITS OWN, because position was a confound and
   * this file spent a day reporting it as a defect.
   *
   * `ELEMENTS` runs in a fixed order and `table` is last, so every run measured
   * it on a slide already holding four inserted elements — and it failed, five
   * times, with the probe reporting "table" as the thing that was broken. It is
   * not. Measured 2026-09-18, on one healthy host, in one sitting:
   *
   *     table 1st  on an empty slide   -> Done, grouped, "Table, 4 rows by 5 columns."
   *     harvey 5th on 4 element groups -> Done
   *     table 6th  on 5 element groups -> Failed, shapes 11-20 of 23
   *     table 1st  again after that    -> Done
   *
   * So the element works and the SLIDE LOAD is the variable — which is BACKLOG
   * item 24's real subject and is not what this probe is for. This probe asks
   * whether each of the five carries its description; measuring each on a clean
   * slide is that question's honest scope, and leaving the accumulation in made
   * the answer depend on alphabetical luck.
   */
  const clearSlide = () =>
    pw(
      "eval",
      'async () => { try { await Promise.race([ PowerPoint.run(async (c) => { const sh = c.presentation.getSelectedSlides().getItemAt(0).shapes; sh.load("items/id"); await c.sync(); for (const x of (sh.items || [])) x.delete(); await c.sync(); }), new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), 60000)) ]); return "cleared"; } catch (e) { return "clear-failed"; } }',
      ref,
    );
  const results = [];
  for (const spec of ELEMENTS) {
    clearSlide();
    const before = readAlt(pw("eval", readAltScript(), ref));
    const clicked = /"?(clicked|disabled|no-button)"?/.exec(pw("eval", clickElementScript(spec.el), ref))?.[1] ?? "?";
    const { after, waitedMs, settled } = await settle();
    const verdict = judge(spec, before, after, clicked);
    results.push(verdict);
    const mark = verdict.ok ? "ok  " : verdict.silent ? "----" : "FAIL";
    // SAY WHETHER THE SLIDE HAD STOPPED CHANGING. Without it, a verdict read off
    // the last successful read of a budget that ran out looks exactly like one
    // read off a finished insert — which is how "no group formed" gets reported
    // for an element that was still being drawn.
    const secs = `${Math.round(waitedMs / 1000)}s${settled ? "" : ", STILL CHANGING when the budget ran out"}`;
    const said = verdict.ok ? `${JSON.stringify(verdict.alt)} (${secs})` : `${verdict.why} (waited ${secs})`;
    console.log(`  ${mark} ${spec.el.padEnd(7)} ${said}`);
    // AND WHAT THE PANE SAID, whenever the element did not simply work. This is
    // a DOM read and answers while the host is silent, so it is the one line
    // available in exactly the case the rest of this probe cannot resolve.
    if (!verdict.ok) {
      const n = readNote(pw("eval", readNoteScript(), ref));
      const text = n?.text?.trim();
      console.log(
        `       the pane said: ${text ? JSON.stringify(text) : "nothing — #host-note was empty or unreadable"}`,
      );
      verdict.paneSaid = text ?? null;
    }
  }
  console.log("");
  if (hostSilent(results)) {
    console.log(`The host answered none of the ${results.length} reads. NOTHING WAS MEASURED.`);
    console.log("The pane is on screen and `PowerPoint` is loaded, but `PowerPoint.run` never");
    console.log("resolves — the same stall that blind-skipped scenarios in rounds 458 and 459.");
    console.log("Reload the deck and the add-in, then run this again. This is not a verdict on");
    console.log("`altTextDescription`.");
    process.exit(2);
  }
  const { carried, ungrouped, undescribed, absent, unclickable, unasked, code } = summarise(results);
  console.log(`${carried.length} of ${results.length} elements carried an altTextDescription.`);
  if (unasked.length) {
    console.log(`${unasked.length} were never asked — the host went quiet on ${unasked.join(", ")}.`);
    console.log("Those are not verdicts. Re-run until every element has been asked once.");
  }
  if (ungrouped.length) {
    console.log(`${ungrouped.length} produced no group: ${ungrouped.join(", ")}.`);
    console.log("Their parts landed loose, so there was nowhere to write a description. Either the");
    console.log("host refused to group or the draw stopped before grouping — READ THE PANE'S STATUS");
    console.log("LINE, which says which. Both look identical from a shape list, and on 2026-09-17");
    console.log("one sync's worth of a stalled 23-shape table was recorded as a grouping refusal.");
  }
  if (unclickable.length) {
    console.log(`${unclickable.length} could not be clicked at all: ${unclickable.join(", ")}.`);
    console.log("The pane offered no usable button, so nothing was asked of the host. This is the");
    console.log("pane, not the 1.10 write — check the Elements tab rendered.");
  }
  if (absent.length) {
    console.log(`${absent.length} put nothing on the slide: ${absent.join(", ")}.`);
    console.log("The click was accepted and no shape followed it, so there is nothing yet to say");
    console.log("about descriptions. Check the pane's own status line for what it reported.");
  }
  if (undescribed.length) {
    console.log(`${undescribed.length} grouped and were not described: ${undescribed.join(", ")}.`);
    console.log("This is the 1.10 `Shape.altTextDescription` write in `powerpoint.ts` — the only");
    console.log("Office.js-reaching change in v0.6.1, and the thing no scenario covers.");
  }
  process.exit(code);
}

if (isMain(import.meta.url, process.argv[1])) {
  main().catch((err) => {
    console.error(err);
    process.exit(2);
  });
}
