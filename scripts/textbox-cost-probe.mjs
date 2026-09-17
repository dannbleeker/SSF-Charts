/**
 * IS IT THE TEXT BOXES? The hypothesis behind BACKLOG item 24, tested.
 *
 * `table` is the one Elements button that does not draw on PowerPoint on the
 * web. The pane names the failure itself, four times across 2026-09-17 and
 * twice consecutively on a machine that had just run a clean three-leg cycle:
 *
 *   Failed: PowerPoint did not respond while drawing shapes 1-10 of 23 (45s)
 *   | at=drawing the chart's shapes
 *
 * IT IS NOT SHAPE COUNT. `harvey` is 24 shapes and finishes in 11 seconds;
 * `table` is 23 and dies on its first batch of ten. `SHAPES_PER_SYNC` is 10, so
 * both send the host the same number of shapes in the same first sync. What
 * differs is COMPOSITION — harvey's first ten are `harvey-ring` plus nine
 * `harvey-fill-*`, every one a geometric shape; the table's are `rule-top`,
 * `rule-header` and EIGHT `cell-text-*`, i.e. eight `addTextBox` calls.
 *
 * That is a hypothesis with one supporting comparison, on one host, where
 * nothing was varied deliberately. This varies it.
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
  { kind: "text", what: "10 text boxes" },
  { kind: "mixed", what: "2 lines + 8 text boxes — the table's real first batch" },
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
  const body =
    kind === "text"
      ? `for (let i = 0; i < ${n}; i++) { ${addText}; }`
      : kind === "rect"
        ? `for (let i = 0; i < ${n}; i++) { ${addRect}; }`
        : `for (let i = 0; i < ${n}; i++) { if (i < 2) { sh.addLine(PowerPoint.ConnectorType.straight, { left: 20, top: 20 + i * 30, width: 900, height: 0.5 }); } else { ${addText}; } }`;
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
  const rect = by.rect;
  const text = by.text;
  const thin = KINDS.some((k) => by[k.kind].completed < 2);
  let verdict;
  if (text.completed === 0 && rect.completed >= 2)
    verdict = {
      code: 0,
      says: "STRONGER THAN SLOWER: every text-box batch failed to complete while the rectangles did. The composition claim holds in its sharpest form.",
    };
  else if (thin)
    verdict = {
      code: 2,
      says: "NOT ENOUGH COMPLETED TRIALS to compare. Fewer than two of some kind finished; re-run on a host that is answering.",
    };
  else if (text.medianMs > rect.medianMs * 2)
    verdict = {
      code: 0,
      says: `SUPPORTED: text ${text.medianMs}ms against rect ${rect.medianMs}ms for the same batch size, a ${(text.medianMs / rect.medianMs).toFixed(1)}x difference.`,
    };
  else
    verdict = {
      code: 1,
      says: `REFUTED: text ${text.medianMs}ms against rect ${rect.medianMs}ms is not the gap the hypothesis needs. Composition is not what makes the table's batch fail — look elsewhere.`,
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
