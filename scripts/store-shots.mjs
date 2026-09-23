/**
 * PRODUCE THE STORE LISTING IMAGES — the blocker under `docs/STORE-LISTING.md`.
 *
 * AppSource wants 1-5 images at 1366x768 showing the add-in. The listing doc
 * records why none exist, and corrects itself halfway through: the slide-editing
 * canvas does not composite in a headless browser, so a browser capture shows
 * ribbon and pane correctly and a flat `#F5F5F5` where the slide should be —
 * but the SLIDE is capturable anyway, because `Slide.getImageAsBase64` renders
 * through Office.js rather than through the screen, at a caller-chosen width.
 * Its own docstring calls it "the only way an add-in can see its own output".
 *
 * So there are two genuine captures available and one that is not:
 *
 *   PANE      a browser screenshot. Real, composites fine, it is just HTML.
 *   SLIDE     this script. Real product output, rendered by PowerPoint itself.
 *   BOTH IN
 *   ONE FRAME not available headlessly, and the listing doc is right that
 *             compositing them would be a picture no user ever saw. That is a
 *             misrepresentation question, not a technical one, and it is the
 *             OWNER'S call — this script does not do it.
 *
 * This produces the second. Separate images is route 2 of the three the listing
 * doc lays out, and the only one that needs no decision from anybody.
 *
 * NO HOST CALLS WHILE THE INSERT RUNS. It waits on `#host-note`'s busy CLASS,
 * which is a DOM read costing the host nothing. Polling the slide during a draw
 * is what made the Elements probe report a working product broken five times —
 * see `settleReads` in `scripts/elements-probe.mjs`.
 *
 * IT DELETES EVERY SHAPE ON THE SELECTED SLIDE, once per shot, so each capture
 * starts from nothing. Point it at a scratch deck. PowerPoint's undo can bring
 * the shapes back, but nobody watching a headless run is there to press it.
 *
 * Usage:  node scripts/store-shots.mjs [outDir]
 * Exit:   0 wrote every shot · 1 a shot failed · 2 could not ask
 */

import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { isMain } from "./is-main.mjs";

const CLI = process.env.PLAYWRIGHT_CLI_JS;
const DIR = "C:\\devtools\\SSF-Charts\\.pw-session";

/** AppSource's listing image size. */
export const SHOT_WIDTH = 1366;

/**
 * The shots to take, in order.
 *
 * Chosen to answer the question a reviewer actually has — "what does this put on
 * my slide?" — with the kinds the listing copy names first. Titles are the ones
 * a consultant would type, not `test` or `foo`: a store image is the product's
 * first impression and placeholder text reads as unfinished.
 */
export const SHOTS = [
  { file: "01-waterfall.png", kind: "waterfall", title: "Revenue bridge FY25", w: 780, h: 420 },
  { file: "02-stacked.png", kind: "stacked", title: "Revenue by segment", w: 780, h: 420 },
];

/**
 * THE VALUE IS ON STDOUT. THE NOISE IS ON STDERR. NEVER CONCATENATE THEM.
 *
 * This returned `stdout + stderr` for its first two runs, and the cost was
 * every poll of the pane reading as a non-answer: the CLI prints an
 * "Update available for @playwright/cli" banner on stderr, so the joined
 * string ended with a box-drawing frame rather than with the value, and
 * `readString` — which anchors to the end — returned null forever. The pane
 * had said `"idle:Done."` on stdout the whole time.
 *
 * What made it survive two runs is that it works BY HAND. A terminal
 * interleaves the two streams in time order, so the banner comes first and the
 * value last, and every manual check confirmed a parser that could not work
 * under `spawnSync`, where the streams are captured apart and joined in a fixed
 * order. The check and the real call path disagreed about the input.
 *
 * So: `out` is the only thing a value is ever read from. `err` is kept for
 * reporting, because a CLI that fails on stderr must not read as silence.
 */
const pw = (...args) => {
  const r = spawnSync(process.execPath, [CLI, "-s=ms", "--raw", ...args], {
    encoding: "utf8",
    cwd: DIR,
    maxBuffer: 256e6,
    timeout: 300_000,
  });
  if (r.error) throw r.error;
  return { out: String(r.stdout ?? ""), err: String(r.stderr ?? ""), status: r.status };
};

/**
 * The value a `pw()` call produced, or null if it produced none.
 *
 * Exported so a test can feed it a REAL pair of streams. The bug above lived
 * through two runs because the only thing under test was `readString`, handed a
 * tidy string that no call path ever produces — a fake that could not fail the
 * way the real thing failed.
 */
export const readValue = (r) => readString(r?.out);
const said = readValue;

/** Pull a quoted string out of the CLI's `--raw` output, unescaping it. */
export function readString(out) {
  const m = /"((?:[^"\\]|\\.)*)"\s*$/.exec(String(out ?? "").trim());
  if (!m) return null;
  try {
    return JSON.parse(`"${m[1]}"`);
  } catch {
    return m[1];
  }
}

/**
 * Base64 comes back in SLICES, and that is not premature caution.
 *
 * A 1366px slide render is around 40,000 base64 characters and the CLI escapes
 * what it prints, so a single `eval` that returned the whole string would be one
 * very long shell-quoted line — the kind of thing that gets truncated somewhere
 * between the page and here without saying so, leaving a PNG that is a valid
 * prefix and a corrupt file. Stashing it on `window` and reading fixed-size
 * slices makes the transfer checkable: the length is asked for first, and the
 * assembled string has to match it.
 */
export const SLICE = 12_000;

/**
 * How long a chart is allowed to take, and how often to ask.
 *
 * A single 780x420 waterfall settled in 63s on a rested session — the only
 * direct measurement there is. Session depth is the variable that moves it
 * (`SESSION_GAP_MS` in `scripts/round.mjs`: a session deepens on wall-clock gap,
 * and a deep one makes every sync longer), so the budget covers a deep session
 * rather than the best case. Eight minutes is generous on purpose: overshooting
 * costs waiting, undershooting costs a false red about a shipped product.
 *
 * WHAT THIS NUMBER IS NOT. The first two runs of this script reported "the pane
 * never settled" and the budget had nothing to do with it — `pw()` was joining
 * stdout to stderr, so the parser never saw a value and no budget would have
 * been enough. Raising 180s to 480s was the right change for the wrong reason,
 * and the reason is recorded here because a number defended by a false story is
 * the one nobody re-derives.
 */
export const IDLE_BUDGET_MS = 480_000;
export const IDLE_POLL_MS = 3000;

/**
 * How many times to ask the slide what it holds, and how long between asks.
 *
 * Not retry-until-it-says-what-I-want: it exists because this host has a
 * measured window in which a shape that is ON the slide is not yet counted by
 * a read (BACKLOG item 5, draft D — 0 of 6 trials counted a drawn shape inside
 * two minutes on a clean deck). A single read therefore cannot distinguish
 * "not visible to a read yet" from "never grouped". Four asks over 24s
 * resolves the first and leaves the second reported.
 */
export const VERIFY_TRIES = 4;
export const VERIFY_GAP_MS = 8000;

/** Is the pane still working? DOM only — see the file header. */
export const busyScript = () =>
  "() => { const n = document.getElementById('host-note'); " +
  "return n ? (/status-busy/.test(n.className) ? 'busy' : 'idle:' + (n.textContent || '').slice(0, 120)) : 'no-note'; }";

/**
 * Set the chart up the way a user would, then press Insert.
 *
 * IT STAMPS THE NOTE BUSY BEFORE IT CLICKS, and that is not decoration. The
 * pane sets `status-busy` itself, but only once its click handler gets that
 * far, and the note left behind by the PREVIOUS action still reads "Done." in
 * the meantime. A waiter polling 3s later can therefore see the old success
 * and conclude this insert has finished before it has started — then capture a
 * blank or half-drawn slide and call it a shot.
 *
 * Writing the sentinel here closes the window: from the click onward the note
 * cannot say anything but "busy" until the pane's own `note()` replaces it with
 * a real outcome. The waiter is then reading THIS insert's verdict and no
 * other. The pane owns the element; this only pre-sets it to the state the pane
 * is about to set anyway.
 */
export const setUpAndInsertScript = (shot) =>
  "async () => { const t = document.querySelector('[data-tab=\"chart\"]'); if (t) t.click(); " +
  "await new Promise((r) => setTimeout(r, 800)); " +
  // THE KIND FIRST, AND THE TITLE AFTER IT. The gallery tile's handler calls
  // `applyConfig(sampleConfig(kind))`, which replaces the WHOLE config — the
  // title included. Typing the title first and then choosing the kind silently
  // discards it, which is how `01-waterfall.png` came out as a stacked column
  // chart captioned "Revenue bridge FY25": a picture that contradicted its own
  // label, and the exact shape of an AppSource "does not match the offer
  // description" rejection.
  `const k = document.querySelector('[data-kind="${shot.kind}"]'); ` +
  `if (!k) return "no-kind:${shot.kind}"; k.click(); ` +
  "await new Promise((r) => setTimeout(r, 800)); " +
  `const ti = document.getElementById("chart-title"); ti.value = ${JSON.stringify(shot.title)}; ti.dispatchEvent(new Event("input")); ` +
  `const w = document.getElementById("chart-w"), h = document.getElementById("chart-h"); ` +
  `w.value = "${shot.w}"; w.dispatchEvent(new Event("input")); h.value = "${shot.h}"; h.dispatchEvent(new Event("input")); ` +
  "await new Promise((r) => setTimeout(r, 800)); " +
  "const n = document.getElementById('host-note'); " +
  `if (n) { n.className = "hint status-busy"; n.textContent = ${JSON.stringify(INSERT_SENTINEL)}; } ` +
  'document.getElementById("insert").click(); return "clicked"; }';

/**
 * What the note is stamped with between the click and the pane's own verdict.
 * If a capture ever reports this as its outcome, the pane never answered at all.
 */
export const INSERT_SENTINEL = "waiting for this insert";

/** Wipe the selected slide so each shot starts from nothing. */
export const clearScript = (budgetMs = 120_000) =>
  "async () => { try { await Promise.race([ PowerPoint.run(async (c) => { " +
  "const sh = c.presentation.getSelectedSlides().getItemAt(0).shapes; " +
  'sh.load("items/id"); await c.sync(); for (const x of (sh.items || [])) x.delete(); await c.sync(); }), ' +
  `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
  'return "cleared"; } catch (e) { return "clear-failed"; } }';

/** The name the renderers group a finished chart under. */
export const GROUP_NAME = "PowerChart";

/**
 * Is the chart actually ON the slide?
 *
 * Without this the capture has a blind spot exactly like the truncated
 * transfer below, one level up: an insert that quietly produced nothing leaves
 * an EMPTY slide, `getImageAsBase64` renders it perfectly, and the script
 * writes a valid 1366px PNG of blank white and reports it as a success with
 * dimensions. Every check downstream — the length, the PNG signature, the
 * header — passes on that file. A store listing image of an empty slide is the
 * worst possible thing to not notice.
 *
 * One host call, made only after the pane has said it is done, which is what
 * the file header permits.
 */
export const verifyScript = (name = GROUP_NAME, budgetMs = 60_000) =>
  "async () => { try { return await Promise.race([ PowerPoint.run(async (c) => { " +
  "const sh = c.presentation.getSelectedSlides().getItemAt(0).shapes; " +
  'sh.load("items/name"); await c.sync(); const items = sh.items || []; ' +
  `return "shapes:" + items.length + ":" + items.filter((s) => s.name === ${JSON.stringify(name)}).length; }), ` +
  `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
  '} catch (e) { return "verify-failed:" + (e && e.message ? e.message : "?"); } }';

/** Render the selected slide and stash the base64 on `window`. */
export const renderScript = (width = SHOT_WIDTH, budgetMs = 120_000) =>
  "async () => { try { return await Promise.race([ PowerPoint.run(async (c) => { " +
  `const img = c.presentation.getSelectedSlides().getItemAt(0).getImageAsBase64({ width: ${width} }); ` +
  'await c.sync(); window.__shot = img.value; return "len:" + img.value.length; }), ' +
  `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
  '} catch (e) { return "render-failed:" + (e && e.message ? e.message : "?"); } }';

async function main() {
  if (!CLI) {
    console.error("PLAYWRIGHT_CLI_JS is not set — nothing was captured.");
    process.exit(2);
  }
  const outDir = process.argv[2] ?? "docs/store-shots";
  const refFor = (text, pattern) => {
    // `.out` — the accessibility tree is on stdout, checked rather than assumed.
    const line = pw("find", text)
      .out.split("\n")
      .find((l) => pattern.test(l));
    return line ? (/ref=([a-z0-9]+)/.exec(line)?.[1] ?? null) : null;
  };
  const ref = refFor("Chart", /tab "Chart"/) ?? refFor("Elements", /tab "Elements"/);
  if (!ref) {
    console.error("no pane frame found — open the deck and the add-in first. Nothing was captured.");
    process.exit(2);
  }
  mkdirSync(outDir, { recursive: true });
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  let failed = 0;
  let wrote = 0;

  /**
   * Wait for the pane to go quiet, and SAY WHAT IT LAST SAID if it does not.
   *
   * The first run of this script reported "the pane never settled" twice and
   * that sentence cost an hour: a stale ref, a missing `#host-note` and a host
   * that is genuinely still drawing all produce it, and it names none of them.
   * The last raw read separates them at a glance.
   */
  const waitForIdle = async (why) => {
    // WALL CLOCK, NOT POLLS. Every `pw()` spawns a fresh `node
    // playwright-cli.js` that has to start and attach, measured at ~1.9s
    // against a 3s sleep — so a budget counted in poll ITERATIONS runs for
    // roughly 1.6x the time it claims, and a run that promised to give up after
    // 480s was still going at 13 minutes. The number in the message has to be
    // the number on the clock, or a timeout gets argued about rather than read.
    const started = Date.now();
    let last = "";
    let polls = 0;
    while (Date.now() - started < IDLE_BUDGET_MS) {
      await sleep(IDLE_POLL_MS);
      polls++;
      last = said(pw("eval", busyScript(), ref)) ?? "(the eval returned nothing)";
      if (last.startsWith("idle:")) return { text: last.slice(5), waitedMs: Date.now() - started, polls };
    }
    return { text: null, last, why, waitedMs: Date.now() - started, polls };
  };

  for (const shot of SHOTS) {
    process.stdout.write(`  ${shot.file.padEnd(18)} `);

    // START FROM QUIET. Clicking Insert onto a busy host is what turned one
    // slow shot into two failures last run: the second shot queued behind the
    // first, so its budget was spent waiting for work it had not asked for.
    const quiet = await waitForIdle("before starting");
    if (quiet.text === null) {
      console.log(`the pane was still busy before this shot even began — last read ${quiet.last}`);
      failed++;
      break;
    }

    /**
     * THE CLEAR HAS TO BE CHECKED, AND ITS RESULT HAS TO BE SEEN.
     *
     * This used to fire `clearScript()` and ignore what came back. `clearScript`
     * is raced against a budget and returns `"clear-failed"` when it loses, and
     * the run carried on regardless — so the next chart drew ON TOP of the
     * previous one's leftovers.
     *
     * It produced a picture that named its own cause: `01-waterfall.png`
     * captured on 2026-09-23 held the title rendered TWICE at two sizes, one
     * from each chart, with the axis labels pushed out of frame. Every check in
     * this file passed it — shapes were present, the transfer was whole, the
     * PNG was valid. The only thing that caught it was looking at it.
     *
     * So: the clear must SAY it cleared, and the slide must then READ as empty.
     * The second half matters on its own, because the host can report a delete
     * it has not finished applying.
     */
    const cleared = said(pw("eval", clearScript(), ref));
    if (cleared !== "cleared") {
      console.log(`the slide would not clear — ${String(cleared).slice(0, 50)}. Not drawing on top of what is there.`);
      failed++;
      continue;
    }
    let emptied = false;
    for (let t = 0; t < VERIFY_TRIES; t++) {
      if (t) await sleep(VERIFY_GAP_MS);
      const after = /^shapes:(\d+):/.exec(said(pw("eval", verifyScript(), ref)) ?? "");
      if (after && Number(after[1]) === 0) {
        emptied = true;
        break;
      }
    }
    if (!emptied) {
      console.log(
        `the slide still holds shapes after clearing it. Drawing now would layer this chart over ` +
          `the last one — which is exactly how a capture came back with two titles on it.`,
      );
      failed++;
      continue;
    }

    const clicked = said(pw("eval", setUpAndInsertScript(shot), ref));
    if (clicked !== "clicked") {
      // Most often `no-kind:<kind>` — the gallery does not offer that tile, so
      // nothing was clicked and waiting would time out describing the wrong
      // thing. Say which step refused instead.
      console.log(`the pane would not set this chart up — ${String(clicked).slice(0, 60)}`);
      failed++;
      continue;
    }

    // Wait on the PANE, not the host. See the file header.
    const done = await waitForIdle("after insert");
    if (done.text === null) {
      console.log(`the pane never settled in ${IDLE_BUDGET_MS / 1000}s — last read ${done.last}`);
      failed++;
      // A host this slow will not be faster for the next shot, and a queued
      // click would land mid-draw. Stop rather than manufacture more reds.
      break;
    }
    const settled = done.text;
    if (/^Failed/i.test(settled)) {
      console.log(`the insert failed — ${settled.slice(0, 70)}`);
      failed++;
      continue;
    }

    /**
     * IS THERE A CHART THERE? An empty slide renders to a perfectly valid PNG.
     *
     * ASKED MORE THAN ONCE, because one read cannot tell the two answers apart.
     * This host has a measured read-recency window — BACKLOG item 5, draft D:
     * `shapes.load("items/id")` returns an EMPTY collection for a shape a
     * screenshot shows on the slide, for a while after it arrives. So a first
     * read finding loose parts and no group is equally "the group has not
     * become visible to a read yet" and "this insert never grouped". Re-asking
     * separates them: recency resolves, a real refusal does not.
     *
     * Measured here on 2026-09-23: a shot was rejected on `80 shapes, no
     * PowerChart` immediately after the pane said "Done.".
     */
    let counts = null;
    let seen = "";
    for (let t = 0; t < VERIFY_TRIES; t++) {
      if (t) await sleep(VERIFY_GAP_MS);
      seen = said(pw("eval", verifyScript(), ref)) ?? "";
      counts = /^shapes:(\d+):(\d+)$/.exec(seen);
      if (counts && Number(counts[2]) >= 1) break;
    }
    if (!counts) {
      console.log(`could not check what landed on the slide — ${seen.slice(0, 60)}`);
      failed++;
      continue;
    }
    /**
     * NO GROUP, NO CAPTURE — AND THAT RULE WAS REMOVED ONCE AND PUT BACK.
     *
     * The argument for removing it sounded good: `powerpoint.ts` traces "the
     * host refused addGroup", a refusal costs re-editing rather than
     * appearance, and a listing image is only an appearance. So the check was
     * softened to a printed note and the shot was written anyway.
     *
     * THE IMAGE REFUTED IT. That capture held the title drawn twice at two
     * sizes and the axis labels pushed out of frame — a chart layered on the
     * previous one's leftovers. Ungrouped did not mean "drew fine, failed to
     * group"; it meant the insert had not come out whole. Whether the missing
     * group is cause or symptom, it is the only signal available here that
     * separates a good render from that one, and it costs a re-run to honour.
     *
     * Kept as a refusal, on evidence, against a plausible argument that had
     * none.
     */
    const loose = Number(counts[1]);
    if (loose === 0) {
      console.log(`the pane said "${settled.slice(0, 40)}" and the slide is empty — not capturing a blank slide`);
      failed++;
      continue;
    }
    if (Number(counts[2]) < 1) {
      console.log(
        `${loose} shapes landed but never grouped into ${GROUP_NAME} over ` +
          `${(VERIFY_TRIES * VERIFY_GAP_MS) / 1000}s — the pane said "${settled.slice(0, 30)}". An insert ` +
          `that does not group has not come out whole; the last one captured this way had its title ` +
          `drawn twice. Re-run rather than ship it.`,
      );
      failed++;
      continue;
    }

    const lenSaid = said(pw("eval", renderScript(), ref)) ?? "";
    const expected = Number(/^len:(\d+)$/.exec(lenSaid)?.[1] ?? NaN);
    if (!Number.isFinite(expected)) {
      console.log(`the host would not render the slide — ${lenSaid.slice(0, 60)}`);
      failed++;
      continue;
    }
    let b64 = "";
    for (let at = 0; at < expected; at += SLICE) {
      const part = said(pw("eval", `() => window.__shot.slice(${at}, ${at + SLICE})`, ref));
      if (part === null) break;
      b64 += part;
    }
    // THE LENGTH IS THE CHECK. A truncated transfer still decodes to a valid
    // PNG prefix, which would be a corrupt image that looks like a success.
    if (b64.length !== expected) {
      console.log(`transfer came back short — ${b64.length} of ${expected} characters`);
      failed++;
      continue;
    }
    const png = Buffer.from(b64, "base64");
    if (png.subarray(0, 8).toString("hex") !== "89504e470d0a1a0a") {
      console.log("what came back is not a PNG");
      failed++;
      continue;
    }
    writeFileSync(`${outDir}/${shot.file}`, png);
    wrote++;
    console.log(`${png.length} bytes · ${png.readUInt32BE(16)}x${png.readUInt32BE(20)} · "${shot.title}"`);
  }

  console.log("");
  // COUNT WHAT WAS WRITTEN, not what did not fail. The loop stops early on a
  // busy host, so the shots after it were never attempted — reporting them as
  // failures would be a verdict on a host that was never asked.
  const skipped = SHOTS.length - wrote - failed;
  console.log(`${wrote} of ${SHOTS.length} written to ${outDir}/` + (skipped ? ` · ${skipped} never attempted` : ""));
  if (failed) console.log("A shot that failed is not a product verdict — re-run it on a host that is answering.");
  process.exit(failed ? 1 : 0);
}

if (isMain(import.meta.url, process.argv[1])) {
  main().catch((err) => {
    console.error(err);
    process.exit(2);
  });
}
