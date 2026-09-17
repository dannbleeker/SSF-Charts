/**
 * WHY DOES `table` NOT DRAW? A bench for BACKLOG item 24 — and a record of two
 * hypotheses it has already killed.
 *
 * `table` is the one Elements button that does not draw on PowerPoint on the
 * web. The pane names the failure itself, five times now, the last of them on a
 * host answering a bare read in 141ms where the other four Elements drew in
 * 11-43s and this bench completed 16 of 16 trials:
 *
 *   Failed: PowerPoint did not respond while drawing shapes 1-10 of 23 (45s)
 *   | at=drawing the chart's shapes
 *
 * So it is DETERMINISTIC, not the host's weather. That much is settled.
 *
 * WHAT THIS BENCH HAS ESTABLISHED, in order:
 *
 *   1. COMPOSITION — REFUTED. Ten text boxes against ten rectangles is 571ms
 *      against 419ms, 1.3x. Not a cause.
 *   2. STATEMENT COUNT — SUPPORTED, AND IT IS NOT THE ANSWER. The same ten
 *      shapes styled the way `addText` styles them cost 1641ms against 498ms
 *      unstyled: 3.4x for identical geometry. `SHAPES_PER_SYNC` counts SHAPES
 *      and the cost is statements, which is worth knowing on its own. But
 *      1.6 seconds is not the 45 the table dies at, so it explains a cost and
 *      not a failure.
 *   3. THE EMPTY CORNER CELL — REFUTED. `buildTableScene` emits a text node for
 *      every cell including the blank one at 0,0, and the failed insert's own
 *      shapes read back with `cell-text-0-0` carrying `""`. No other Element has
 *      one. Adding exactly that to the styled batch: 1628ms against 1641ms.
 *      Nothing.
 *
 * SO THE CAUSE IS NOT IN THE SHAPES. The table's first batch, reproduced as
 * faithfully as this bench can reproduce it, completes in 1.6 seconds; the real
 * one does not complete in 45. The difference is something the renderer does
 * around the adds — the slide reference it draws onto, the names it sets, the
 * box geometry, or something outside `addText` entirely. The next arm should
 * close one of those, and the one after that should stop modelling and
 * instrument the real path instead.
 *
 * WHY IT IS A SEPARATE SCRIPT AND NOT A HOST-PROBE QUESTION. `host-probe.ts`
 * ships, and its questions run in every round; a question earns that cost by
 * being something worth watching forever. This is a one-off that answers yes or
 * no. If the answer turns out to matter every round, it can graduate then — and
 * `scripts/` changes nothing a user receives, so it can land mid-window without
 * moving the deployed build.
 *
 * Usage:  node scripts/textbox-cost-probe.mjs [repeats]
 * Exit:   0 measured · 1 measured and the hypothesis is REFUTED · 2 could not ask
 */

import { spawnSync } from "node:child_process";
import { isMain } from "./is-main.mjs";

const CLI = process.env.PLAYWRIGHT_CLI_JS;
const DIR = "C:\\devtools\\SSF-Charts\\.pw-session";

/** What one sync is allowed before the trial is recorded as a non-answer. */
export const TRIAL_BUDGET_MS = 60_000;

/** The batch size the renderer actually uses — `powerpoint.ts` SHAPES_PER_SYNC. */
export const BATCH = 10;

/**
 * The three batches, all of size `BATCH`, differing only in composition.
 *
 * `mixed` is the table's real first batch, reproduced: two rules and eight cell
 * texts. It is here so the experiment cannot be answered "your synthetic case
 * is not what the table does" — if `mixed` behaves like `text` and unlike
 * `rect`, the composition claim carries from a contrived batch to the real one.
 */
export const KINDS = [
  { kind: "rect", what: "10 geometric rectangles — harvey's first batch in kind" },
  { kind: "text", what: "10 BARE text boxes — addTextBox and nothing else" },
  { kind: "mixed", what: "2 lines + 8 bare text boxes — the table's batch in SHAPES" },
  {
    kind: "styled",
    what: "2 lines + 8 text boxes STYLED the way addText styles them — the table's batch in STATEMENTS",
  },
  { kind: "empty", what: "the styled batch with ONE empty string — the table's batch EXACTLY, corner cell and all" },
];

/**
 * Which pair the verdict turns on, and what it is a hypothesis ABOUT.
 *
 * Declared rather than hard-coded into `summarise` because this file has now
 * carried three hypotheses and will carry more: composition (refuted at 1.3x),
 * statement count (supported at 3.4x, and 1.6s is not the 45s the table dies
 * at, so it explains a cost and not the failure), and now the empty cell.
 *
 * `test` against `control` must differ in ONE thing. Here that is the text of a
 * single box: `cell-text-0-0`, the blank corner of the table header, is the one
 * `addTextBox("")` in the first batch, and no other Element has one.
 */
export const HYPOTHESIS = {
  test: "empty",
  control: "styled",
  about: "an empty-string text box",
  soWhat: "the table's first batch is the styled batch plus one empty cell, and only the table has one",
};

/**
 * The properties `addText` sets on every text node, in its order.
 *
 * THE POINT OF THE FOURTH KIND, and the reason the first three refuted the
 * wrong hypothesis. `SHAPES_PER_SYNC` counts SHAPES; this repo has already
 * measured that a whole text node is **20 statements** and a rect **7**
 * (`docs/BACKLOG.md`, the in-place update work). So the table's first batch and
 * harvey's first batch are both "ten shapes" and are ~174 and ~70 statements —
 * two and a half times apart in the quantity that actually crosses to the host.
 *
 * A bare `addTextBox` is one statement, which is why `text` and `mixed` came
 * back at 555ms and 505ms and refuted composition. They were not the table's
 * batch; they were its shape COUNT wearing its name. This kind is the batch.
 */
export const STYLED_STATEMENTS = [
  "sh2.fill.clear()",
  "sh2.lineFormat.visible = false",
  "tf.wordWrap = false",
  "tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone",
  "tf.leftMargin = 0",
  "tf.rightMargin = 0",
  "tf.topMargin = 0",
  "tf.bottomMargin = 0",
  "tf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle",
  "font.size = 11",
  'font.color = "#1A202C"',
  "font.bold = false",
  'font.name = "Segoe UI"',
  "tf.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.left",
];

/**
 * Trial order: strictly alternating, never blocked.
 *
 * A HOST THAT TIRES WOULD MANUFACTURE THIS RESULT if the kinds ran in blocks —
 * run ten rects then ten text batches and the text ones are slower because they
 * are later, which is the shape of the finding and none of its substance. This
 * archive has already paid for that confusion once: `what makes a long run slow
 * down` measured a per-slide cost curve that turned out to be session age.
 *
 * Alternating puts each kind at the same average position in the session, so
 * drift hits all three equally and cancels in the comparison.
 */
export function plan(repeats) {
  const out = [];
  for (let r = 0; r < repeats; r++) for (const k of KINDS) out.push({ kind: k.kind, round: r });
  return out;
}

/** Build one trial's in-page script. Geometry keeps the shapes on a 960x540 slide. */
export const trialScript = (kind, n = BATCH, budgetMs = TRIAL_BUDGET_MS) => {
  const box = "{ left: 20 + (i % 5) * 180, top: 20 + Math.floor(i / 5) * 60, width: 160, height: 40 }";
  const addText = `sh.addTextBox("cell " + i, ${box})`;
  const addRect = `sh.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, ${box})`;
  // `mixed` reproduces rule-top/rule-header plus eight cell texts: the first two
  // are lines, the rest text, which is the order `buildTableScene` emits.
  const rule = "sh.addLine(PowerPoint.ConnectorType.straight, { left: 20, top: 20 + i * 30, width: 900, height: 0.5 })";
  const styledBody = (text) =>
    `{ const sh2 = sh.addTextBox(${text}, ${box}); const tf = sh2.textFrame; const font = tf.textRange.font; ` +
    STYLED_STATEMENTS.join("; ") +
    "; }";
  const styledText = styledBody('"cell " + i');
  // The corner cell, and only it: `buildTableScene` emits a text node for every
  // cell including the blank one at 0,0. Everything else in this batch is
  // identical to `styled`, so the pair differs in one string.
  const emptyText = `if (i === 2) ${styledBody('""')} else ${styledBody('"cell " + i')}`;
  const body =
    kind === "text"
      ? `for (let i = 0; i < ${n}; i++) { ${addText}; }`
      : kind === "rect"
        ? `for (let i = 0; i < ${n}; i++) { ${addRect}; }`
        : kind === "styled"
          ? `for (let i = 0; i < ${n}; i++) { if (i < 2) { ${rule}; } else ${styledText} }`
          : kind === "empty"
            ? `for (let i = 0; i < ${n}; i++) { if (i < 2) { ${rule}; } else { ${emptyText} } }`
            : `for (let i = 0; i < ${n}; i++) { if (i < 2) { ${rule}; } else { ${addText}; } }`;
  return (
    "async () => { const t0 = Date.now(); try { " +
    "await Promise.race([ PowerPoint.run(async (c) => { " +
    "const s = c.presentation.getSelectedSlides().getItemAt(0); const sh = s.shapes; " +
    body +
    " await c.sync(); }), " +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    'return "trial:" + JSON.stringify({ ok: true, ms: Date.now() - t0 }); } ' +
    'catch (e) { return "trial:" + JSON.stringify({ ok: false, ms: Date.now() - t0, why: e && e.message ? String(e.message).slice(0, 60) : "?" }); } }'
  );
};

/** Wipe the slide between trials, so shape count never becomes the variable. */
export const clearScript = (budgetMs = TRIAL_BUDGET_MS) =>
  "async () => { try { await Promise.race([ PowerPoint.run(async (c) => { " +
  "const s = c.presentation.getSelectedSlides().getItemAt(0); const sh = s.shapes; " +
  'sh.load("items/id"); await c.sync(); for (const x of (sh.items || [])) x.delete(); await c.sync(); }), ' +
  `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
  'return "cleared"; } catch (e) { return "clear-failed"; } }';

/** `trial:{...}` -> the object, tolerating the CLI's escaped quotes. See readAlt. */
export function readTrial(out) {
  const m = /trial:(\{.*?\})\s*"?\s*$|trial:(\{.*\})/s.exec(String(out ?? ""));
  const body = m?.[1] ?? m?.[2];
  if (!body) return null;
  for (const candidate of [body, body.replace(/\\"/g, '"')]) {
    try {
      const parsed = JSON.parse(candidate);
      if (parsed && typeof parsed === "object") return parsed;
    } catch {
      // try the next spelling
    }
  }
  return null;
}

/**
 * Stop after a whole round in which nothing completed.
 *
 * TWELVE TRIALS AT A 60s BUDGET IS TWELVE MINUTES, and on 2026-09-18 all twelve
 * spent it against a pane that had died — including `rect`, the arm that is
 * supposed to be fast. Every one of them was the same non-answer, and the
 * twelfth was worth no more than the third.
 *
 * One full round is the unit, not one trial: a single timeout is this host on
 * an ordinary bad day, three in a row covering all three kinds is the host not
 * answering at all. Abort then, and say so — the result is `could not ask`
 * either way, so the only thing the remaining nine minutes buys is a later
 * finish.
 */
export function shouldAbort(results, kinds = KINDS.length) {
  if (results.length < kinds) return false;
  return results.slice(0, kinds).every((r) => !r.ok);
}

const median = (xs) => {
  const s = [...xs].sort((a, b) => a - b);
  if (!s.length) return null;
  const mid = s.length >> 1;
  return s.length % 2 ? s[mid] : Math.round((s[mid - 1] + s[mid]) / 2);
};

/**
 * What the trials add up to — and what they are NOT allowed to add up to.
 *
 * A TIMED-OUT TRIAL IS NOT A SLOW ONE. It is a non-answer, and averaging its
 * budget in as though it were a measurement is the same error as counting a
 * silent host as a wrong answer, which this repo spent 2026-09-17 undoing. So
 * timeouts are counted separately and never enter a median.
 *
 * And the verdict REFUSES rather than guesses. With fewer than two completed
 * trials of any kind there is no comparison to make, and a run where every
 * `text` trial timed out is the most interesting outcome available — it says
 * the batch does not complete at all, which is stronger than "it is slower" —
 * so that case is named rather than dropped.
 */
export function summarise(results) {
  const by = {};
  for (const k of KINDS) {
    const mine = results.filter((r) => r.kind === k.kind);
    const done = mine.filter((r) => r.ok);
    by[k.kind] = {
      what: k.what,
      trials: mine.length,
      completed: done.length,
      timedOut: mine.length - done.length,
      medianMs: median(done.map((r) => r.ms)),
    };
  }
  // `styled` AGAINST `mixed` IS THE COMPARISON, because they are the same ten
  // shapes in the same composition and differ ONLY in the property writes —
  // which is the statement count. `rect` and `text` stay as the controls that
  // showed shape count and composition were not the variable.
  // The declared pair, differing in one thing. See HYPOTHESIS.
  const control = by[HYPOTHESIS.control];
  const test = by[HYPOTHESIS.test];
  const thin = KINDS.some((k) => by[k.kind].completed < 2 && k.kind !== HYPOTHESIS.test);
  let verdict;
  if (test.completed === 0 && control.completed >= 2)
    verdict = {
      code: 0,
      says:
        `STRONGER THAN SLOWER: no \`${HYPOTHESIS.test}\` batch completed while \`${HYPOTHESIS.control}\` — the same ` +
        `batch without ${HYPOTHESIS.about} — completed ${control.completed}/${control.trials} at ${control.medianMs}ms. ` +
        `The cause is ${HYPOTHESIS.about}, and ${HYPOTHESIS.soWhat}.`,
    };
  else if (thin || test.completed < 2)
    verdict = {
      code: 2,
      says: "NOT ENOUGH COMPLETED TRIALS to compare. Fewer than two of some kind finished; re-run on a host that is answering.",
    };
  else if (test.medianMs > control.medianMs * 2)
    verdict = {
      code: 0,
      says:
        `SUPPORTED: \`${HYPOTHESIS.test}\` ${test.medianMs}ms against \`${HYPOTHESIS.control}\` ${control.medianMs}ms — ` +
        `${(test.medianMs / control.medianMs).toFixed(1)}x, and the batches differ only in ${HYPOTHESIS.about}.`,
    };
  else
    verdict = {
      code: 1,
      says:
        `REFUTED: \`${HYPOTHESIS.test}\` ${test.medianMs}ms against \`${HYPOTHESIS.control}\` ${control.medianMs}ms is not ` +
        `the gap the hypothesis needs. ${HYPOTHESIS.about} is not what makes the table's batch fail — look elsewhere.`,
    };
  return { by, verdict };
}

async function main() {
  if (!CLI) {
    console.error("PLAYWRIGHT_CLI_JS is not set — nothing below was measured.");
    process.exit(2);
  }
  const repeats = Number(process.argv[2]) || 4;
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
  const refFor = (text, pattern) => {
    const line = pw("find", text)
      .split("\n")
      .find((l) => pattern.test(l));
    return line ? (/ref=([a-z0-9]+)/.exec(line)?.[1] ?? null) : null;
  };
  const ref =
    refFor("Elements", /tab "Elements"/) ?? refFor("Automation", /tab "Automation"/) ?? refFor("Chart", /tab "Chart"/);
  if (!ref) {
    console.error("no pane frame found. Open the deck and the add-in first. Nothing was measured.");
    process.exit(2);
  }
  if (
    !/office\|pane/.test(
      pw(
        "eval",
        "() => (typeof PowerPoint !== 'undefined' ? 'office' : 'no-office') + '|' + (document.getElementById('harvey-insert') ? 'pane' : 'no-pane')",
        ref,
      ),
    )
  ) {
    console.error("the frame found is not the add-in pane. Nothing was measured.");
    process.exit(2);
  }
  // PREFLIGHT: ASK THE HOST ONE CHEAP QUESTION FIRST. The frame check above
  // proves the PANE is there; it says nothing about whether `PowerPoint.run`
  // resolves, and those are different states this repo has confused more than
  // once. On 2026-09-18 a dead host turned the whole experiment into twelve
  // identical non-answers and twelve minutes.
  const alive = readTrial(
    pw(
      "eval",
      'async () => { const t0 = Date.now(); try { await Promise.race([ PowerPoint.run(async (c) => { const sh = c.presentation.getSelectedSlides().getItemAt(0).shapes; sh.load("items/id"); await c.sync(); }), new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), 30000)) ]); return "trial:" + JSON.stringify({ ok: true, ms: Date.now() - t0 }); } catch (e) { return "trial:" + JSON.stringify({ ok: false, ms: Date.now() - t0 }); } }',
      ref,
    ),
  );
  if (!alive?.ok) {
    console.error("the pane is on screen but the host did not answer a bare shape read in 30s. Nothing was measured.");
    process.exit(2);
  }
  console.log(`host answered a bare read in ${alive.ms}ms — starting.\n`);
  const order = plan(repeats);
  console.log(
    `${order.length} trials, ${repeats} of each kind, strictly alternating so session drift cannot become the finding.\n`,
  );
  const results = [];
  let aborted = false;
  for (const t of order) {
    pw("eval", clearScript(), ref);
    const got = readTrial(pw("eval", trialScript(t.kind), ref));
    const row = { kind: t.kind, ok: Boolean(got?.ok), ms: Number(got?.ms ?? 0), why: got?.why };
    results.push(row);
    console.log(
      `  r${t.round + 1} ${t.kind.padEnd(6)} ${row.ok ? `${row.ms}ms` : `DID NOT COMPLETE after ${row.ms}ms (${row.why ?? "?"})`}`,
    );
    if (shouldAbort(results)) {
      aborted = true;
      console.log("\n  a full round completed nothing — the host has stopped answering. Stopping rather than");
      console.log(`  spending the remaining ${order.length - results.length} trials on the same non-answer.`);
      break;
    }
  }
  pw("eval", clearScript(), ref);
  if (aborted) {
    console.error("\nNothing was measured. Re-run when the host is answering.");
    process.exit(2);
  }
  const { by, verdict } = summarise(results);
  console.log("");
  for (const k of KINDS) {
    const b = by[k.kind];
    console.log(
      `  ${k.kind.padEnd(6)} ${String(b.medianMs ?? "—").padStart(7)}ms median · ${b.completed}/${b.trials} completed${b.timedOut ? `, ${b.timedOut} timed out` : ""}  (${b.what})`,
    );
  }
  console.log("");
  console.log(verdict.says);
  process.exit(verdict.code);
}

if (isMain(import.meta.url, process.argv[1])) {
  main().catch((err) => {
    console.error(err);
    process.exit(2);
  });
}
