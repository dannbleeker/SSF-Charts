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
 *         · 2 could not ask (no pane, no host)
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
  if (clicked !== "clicked") return { el, ok: false, why: `the pane's button said ${clicked}` };
  if (!Array.isArray(before) || !Array.isArray(after))
    return { el, ok: false, why: "the host would not list the slide" };
  const had = new Set(before.map((s) => s.id));
  const added = after.filter((s) => !had.has(s.id));
  if (!added.length) return { el, ok: false, why: "nothing landed on the slide" };
  const described = added.filter((s) => String(s.alt ?? "").trim());
  if (!described.length)
    return { el, ok: false, why: `${added.length} shape(s) landed, none with altTextDescription`, added: added.length };
  const match = described.find((s) => expect.test(s.alt));
  if (!match)
    return {
      el,
      ok: false,
      why: `described, but not as a ${el}: ${JSON.stringify(described.map((s) => s.alt).slice(0, 2))}`,
    };
  return { el, ok: true, alt: match.alt, added: added.length };
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
  const results = [];
  for (const spec of ELEMENTS) {
    const before = readAlt(pw("eval", readAltScript(), ref));
    const clicked = /"?(clicked|disabled|no-button)"?/.exec(pw("eval", clickElementScript(spec.el), ref))?.[1] ?? "?";
    await sleep(9000);
    const after = readAlt(pw("eval", readAltScript(), ref));
    const verdict = judge(spec, before, after, clicked);
    results.push(verdict);
    console.log(
      `  ${verdict.ok ? "ok  " : "FAIL"} ${spec.el.padEnd(7)} ${verdict.ok ? JSON.stringify(verdict.alt) : verdict.why}`,
    );
  }
  const bad = results.filter((r) => !r.ok);
  console.log("");
  console.log(`${results.length - bad.length} of ${results.length} elements carried an altTextDescription.`);
  if (bad.length) {
    console.log("This is the 1.10 `Shape.altTextDescription` write in `powerpoint.ts` — the only");
    console.log("Office.js-reaching change in v0.6.1, and the thing no scenario covers.");
  }
  process.exit(bad.length ? 1 : 0);
}

if (isMain(import.meta.url, process.argv[1])) {
  main().catch((err) => {
    console.error(err);
    process.exit(2);
  });
}
