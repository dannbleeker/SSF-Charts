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

/**
 * How long an insert must have been quiet before "it has stopped changing" is
 * believed. See `settle` — the gap between the draw sync and the grouping sync
 * is a window in which a finished-looking set of loose parts is the middle of
 * an insert, and this is the floor under how long that window is allowed to be.
 */
export const QUIESCE_FLOOR_MS = 24_000;

/** How long an element gets to land before the probe stops waiting for it. */
export const SETTLE_BUDGET_MS = 60_000;

/** Between reads. Short enough to catch a fast insert, long enough not to flood the host. */
export const SETTLE_POLL_MS = 4_000;

/**
 * Read until the slide has stopped changing, or the budget runs out.
 *
 * WAS A FLAT 9s WAIT, and on 2026-09-17 that reported `flow` as "nothing landed
 * on the slide" in a run where the host refused three of five reads outright. A
 * number chosen for a healthy host, applied to a stalling one, turns latency
 * into a defect report.
 *
 * IT MUST NOT STOP AT THE FIRST NEW SHAPE. The first replacement did, and
 * reported all five elements as landing shapes with no `altTextDescription` —
 * including `check`, which had passed with "Checkbox, yes." minutes earlier.
 * These elements draw their parts, group them, and set the description on the
 * group LAST, so the first new id is the middle of the insert. That is the same
 * false red as the flat wait, from the opposite side, inside the same hour.
 *
 * AND QUIESCENCE NEEDS MORE THAN TWO READS, plus a floor under the clock. The
 * renderer commits the parts on one sync, re-reads, and only then commits
 * `addGroup` + the name + the alt text on a SECOND sync. Between those, an
 * outside reader sees a stable set of loose parts and no group — indistinguishable
 * from a grouping refusal. Two consecutive reads is 8 seconds and this host's own
 * evidence records reads stalling 45-60s, so 8s sat inside the ordinary gap.
 *
 * EXPORTED AND INJECTABLE BECAUSE IT IS THE ONLY NEW LOGIC HERE. A review on
 * 2026-09-17 pointed out that reverting any of the above would be invisible —
 * it lived inside `main`, reachable by no test. `read`, `sleep` and `now` are
 * parameters so the whole thing runs in milliseconds against a scripted host.
 */
export async function settleReads(before, read, sleep, opts = {}) {
  const budgetMs = opts.budgetMs ?? SETTLE_BUDGET_MS;
  const pollMs = opts.pollMs ?? SETTLE_POLL_MS;
  const floorMs = opts.floorMs ?? QUIESCE_FLOOR_MS;
  const now = opts.now ?? (() => Date.now());
  const started = now();
  // NO DIFF IS POSSIBLE WITHOUT A BEFORE. `before ?? []` made every shape on the
  // slide look new, so the first read satisfied "something was added" and the
  // quiesce test compared a set that never changes — it stopped about 5s after
  // the click and handed `judge` a full slide as the added set. `judge` marks a
  // null `before` silent whatever comes back, so there is nothing to wait for:
  // take one read for the record and go.
  if (!Array.isArray(before)) {
    await sleep(pollMs);
    return { after: await read(), waitedMs: now() - started, settled: false };
  }
  const had = new Set(before.map((s) => s.id));
  let last = null;
  let previousIds = null;
  let stable = 0;
  let settled = false;
  while (now() - started < budgetMs) {
    await sleep(pollMs);
    const seen = await read();
    // A read the host refused is NOT an answer about the slide. Keep waiting —
    // treating it as "nothing changed" would count silence toward quiescence.
    if (!Array.isArray(seen)) continue;
    last = seen;
    const added = seen.filter((s) => !had.has(s.id));
    // The terminal state: a description arrived. More waiting cannot unset it.
    if (added.some((s) => String(s.alt ?? "").trim())) {
      settled = true;
      break;
    }
    const ids = added
      .map((s) => s.id)
      .sort()
      .join(",");
    stable = added.length && previousIds === ids ? stable + 1 : 0;
    previousIds = ids;
    if (stable >= 2 && now() - started >= floorMs) {
      settled = true;
      break;
    }
  }
  return { after: last, waitedMs: now() - started, settled };
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

/** Click one Elements button by id, inside the pane's frame. */
export const clickElementScript = (el) =>
  "async () => { const tab = document.querySelector('[data-tab=\"elements\"]'); if (tab) tab.click(); " +
  "await new Promise((r) => setTimeout(r, 400)); " +
  `const b = document.getElementById(${JSON.stringify(el + "-insert")}); ` +
  'if (!b) return "no-button"; if (b.disabled) return "disabled"; b.click(); return "clicked"; }';

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
  const settle = (before) => settleReads(before, () => readAlt(pw("eval", readAltScript(), ref)), sleep);
  const results = [];
  for (const spec of ELEMENTS) {
    const before = readAlt(pw("eval", readAltScript(), ref));
    const clicked = /"?(clicked|disabled|no-button)"?/.exec(pw("eval", clickElementScript(spec.el), ref))?.[1] ?? "?";
    const { after, waitedMs, settled } = await settle(before);
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
