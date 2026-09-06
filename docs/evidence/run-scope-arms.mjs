/**
 * Run the Draft A scope experiment inside the task pane's frame.
 *
 *     node run-scope-arms.mjs            # reports what it would do
 *     node run-scope-arms.mjs --go       # actually runs it
 *
 * The route is the driver's own: `pw eval <script> <ref>` evaluates in the
 * frame that owns `ref`, and the pane's ref is where Office.js lives. See
 * `slideResolveScript` in scripts/round.mjs — same mechanism, same shape.
 *
 * NO SHELL. Arguments go through `spawnSync` as an argv array, because this
 * script's payload contains `=>`, braces and quotes, and Git Bash rewrites
 * anything that looks like a path. Both have already cost time in this repo.
 */
import { spawnSync } from "child_process";

const CLI = "C:\\Users\\dann_\\AppData\\Roaming\\npm\\node_modules\\@playwright\\cli\\playwright-cli.js";
const DIR = "C:\\devtools\\SSF-Charts\\.pw-session";
const go = process.argv.includes("--go");

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
 * Three arms, one host. Each draws a shape and tags it from a run LATER than
 * the one that produced its slide — Draft A's failing configuration — and they
 * differ only in where the slide came from.
 *
 *   a  the document's own slide                      index 0
 *   b  a slide added in THIS session, settled first  the addSlides discipline
 *   c  a slide added in a PREVIOUS session           index 5
 *
 * PREDICTIONS, written before running:
 *   Draft A as titled   ->  a OK, b 5010, c 5010
 *   "session" variant   ->  a OK, b OK,   c 5010
 *   "unsettled" variant ->  a OK, b OK,   c OK   (and the draft is ours)
 */
const SCRIPT = `async () => {
  const out = [];
  const rect = PowerPoint.GeometricShapeType.rectangle;
  const drawAndTag = async (label, index) => {
    try {
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(index);
        const s = slide.shapes.addGeometricShape(rect);
        s.left = 20; s.top = 20; s.width = 40; s.height = 30;
        await context.sync();
        slide.shapes.load("items/id");
        await context.sync();
        const listed = slide.shapes.items.length;
        slide.shapes.items[listed - 1].tags.add("SCOPEPROBE", label);
        await context.sync();
        out.push({ arm: label, index, listed, tag: "OK" });
      });
    } catch (e) {
      out.push({ arm: label, index, tag: "THREW", code: e && e.code, message: String(e && e.message).slice(0, 140) });
    }
  };
  let total = 0;
  await PowerPoint.run(async (c) => {
    const slides = c.presentation.slides;
    slides.load("items/id");
    await c.sync();
    total = slides.items.length;
  });
  await drawAndTag("a-document-own", 0);
  await PowerPoint.run(async (c) => { c.presentation.slides.add(); await c.sync(); });
  let addedIndex = -1;
  await PowerPoint.run(async (c) => {
    const slides = c.presentation.slides;
    slides.load("items/id");
    await c.sync();
    addedIndex = slides.items.length - 1;
  });
  await drawAndTag("b-added-this-session", addedIndex);
  await drawAndTag("c-added-previous-session", 5);
  return JSON.stringify({ total, out });
}`;

const tabs = pw("tab-list");
console.log(tabs.split("\n").filter((l) => l.startsWith("- ")).join("\n"));
if (!go) {
  console.log("\n(dry run — pass --go to execute)");
  process.exit(0);
}

// The pane's ref. `tab "Chart"` is what the driver looks for, and the pane must
// be OPEN — a ref from a closed pane evaluates in the document frame, where
// `PowerPoint` is undefined and every arm would report the same false failure.
const found = pw("find", "Chart");
const ref = /\[ref=([a-z0-9]+)\]/.exec(found.split("\n").find((l) => /tab "Chart"/.test(l)) ?? "");
if (!ref) {
  console.error("the pane is not open — no `tab \"Chart\"` ref to evaluate in. Open it, then re-run.");
  console.error(found.slice(0, 800));
  process.exit(1);
}
console.log("pane ref", ref[1]);
console.log(pw("eval", SCRIPT, ref[1]));
