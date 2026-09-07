/**
 * How long is a shape invisible after it is drawn on a freshly added slide?
 *
 *     node docs/evidence/bound-empty-read.mjs [trials]
 *
 * DRAFT D'S FIRST BLOCKER. "Some period after the slide was added" is not a bug
 * report. This adds a slide, draws one shape on it, then asks a FRESH
 * `PowerPoint.run` for `shapes.load("items/id")` every few seconds until the
 * collection is not empty, and records when it first answers.
 *
 * THE LOOP IS HERE, NOT IN THE PAGE. The first version waited inside a single
 * `eval` and the CLI killed it — a page-side loop that outlives the eval budget
 * reports NOTHING, not a partial answer, which is the worst of both. Every eval
 * is now one short question, the same shape the driver's own `slideResolveScript`
 * uses.
 *
 * A CONTROL ON EVERY READ, against a slide that existed before this session. If
 * the control ever reads empty then that reading says nothing about new slides —
 * it says the host was unwell, and without the control the two are the same
 * line.
 *
 * CENSORED TRIALS ARE REPORTED, not dropped. A trial that never fills inside the
 * cap is the most interesting kind, and averaging it away would hide it. Every
 * reading is printed: a mean here would be the next stale number in a file that
 * has already had five.
 */
import { spawnSync } from "child_process";

const CLI = "C:\\Users\\dann_\\AppData\\Roaming\\npm\\node_modules\\@playwright\\cli\\playwright-cli.js";
const DIR = "C:\\devtools\\SSF-Charts\\.pw-session";
const TRIALS = Number(process.argv[2] ?? 10);
const CAP_MS = 120_000;
const EVERY_MS = 4_000;
/** A slide that existed before this session — the control on every read. */
const CONTROL_INDEX = 5;

const pw = (...args) => {
  const r = spawnSync(process.execPath, [CLI, "-s=ms", "--raw", ...args], {
    encoding: "utf8",
    cwd: DIR,
    maxBuffer: 64e6,
    timeout: 90_000,
  });
  if (r.error) return { failed: String(r.error.code ?? r.error.message) };
  return { out: String(r.stdout ?? "") + String(r.stderr ?? "") };
};

/** The last JSON object the CLI echoed back, unescaped. */
const answer = (res) => {
  if (res.failed) return { cliFailed: res.failed };
  const text = res.out.replace(/\\"/g, '"');
  const m = /\{[^{}]*\}/g;
  const all = text.match(m);
  if (!all) return { unparsed: res.out.trim().slice(-160) };
  try {
    return JSON.parse(all[all.length - 1]);
  } catch {
    return { unparsed: all[all.length - 1] };
  }
};

const SETUP = `async () => {
  await PowerPoint.run(async (c) => { c.presentation.slides.add(); await c.sync(); });
  let index = -1, idAtAdd = "";
  await PowerPoint.run(async (c) => {
    const slides = c.presentation.slides;
    slides.load("items/id");
    await c.sync();
    index = slides.items.length - 1;
    idAtAdd = slides.items[index].id;
  });
  // THE COUNT BEFORE THE DRAW, because a slide from slides.add() arrives with
  // layout placeholders on it. Without this, "3 shapes" and "2 shapes" are both
  // unreadable: neither says whether the rectangle is among them, and the first
  // eight trials of this experiment produced exactly that ambiguity.
  // (No backticks in here: this comment lives inside a template literal, and
  // the first version of it closed the string three lines early.)
  let before = -1;
  await PowerPoint.run(async (c) => {
    const slide = c.presentation.slides.getItemAt(index);
    slide.shapes.load("items/id");
    await c.sync();
    before = slide.shapes.items.length;
  });
  await PowerPoint.run(async (c) => {
    const slide = c.presentation.slides.getItemAt(index);
    const s = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    s.left = 20; s.top = 20; s.width = 40; s.height = 30;
    await c.sync();
  });
  return JSON.stringify({ index: index, idAtAdd: idAtAdd, before: before });
}`;

const READ = (index) => `async () => {
  let seen = -1, control = -1, threw = "";
  try {
    await PowerPoint.run(async (c) => {
      const slide = c.presentation.slides.getItemAt(${index});
      slide.shapes.load("items/id");
      await c.sync();
      seen = slide.shapes.items.length;
    });
    await PowerPoint.run(async (c) => {
      const ctl = c.presentation.slides.getItemAt(${CONTROL_INDEX});
      ctl.shapes.load("items/id");
      await c.sync();
      control = ctl.shapes.items.length;
    });
  } catch (e) { threw = String(e && e.message).slice(0, 60); }
  return JSON.stringify({ seen: seen, control: control, threw: threw });
}`;

const found = pw("find", "Chart").out ?? "";
const ref = /\[ref=([a-z0-9]+)\]/.exec(found.split("\n").find((l) => /tab "Chart"/.test(l)) ?? "");
if (!ref) {
  console.error('the pane is not open — no `tab "Chart"` ref to evaluate in.');
  process.exit(1);
}
const PANE = ref[1];
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

for (let n = 0; n < TRIALS; n++) {
  const set = answer(pw("eval", SETUP, PANE));
  if (set.index === undefined) {
    console.log(JSON.stringify({ trial: n, setupFailed: set }));
    continue;
  }
  const drewAt = Date.now();
  const reads = [];
  let filledAtMs;
  while (Date.now() - drewAt < CAP_MS) {
    const r = answer(pw("eval", READ(set.index), PANE));
    reads.push({ at: Date.now() - drewAt, ...r });
    // THE DRAWN SHAPE, not "any shape". A slide that arrives with two
    // placeholders reads 2 before anything is drawn, so `seen > 0` would call
    // the rectangle visible the instant the slide exists.
    if (r.seen > set.before) {
      filledAtMs = Date.now() - drewAt;
      break;
    }
    await sleep(EVERY_MS);
  }
  console.log(
    JSON.stringify({
      trial: n,
      index: set.index,
      idAtAdd: set.idAtAdd,
      // THE THRESHOLD, WRITTEN DOWN. `before` is the sole thing `seen` is
      // compared against, and the first run of this harness left it out of the
      // output entirely — so "the slide had 2 placeholders before the draw" was
      // inferable from the data but not re-derivable from it, which is a weaker
      // thing than it reads as. A file that records a verdict without its
      // threshold is asking to be taken on trust.
      before: set.before,
      filledAtMs,
      censored: filledAtMs === undefined,
      reads,
    }),
  );
}
