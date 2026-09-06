/**
 * Separate the two reads that failed together.
 *
 * The previous arm queued `getItemAt(index)` and `getItem(id)` into ONE sync,
 * so the 5010 from the id killed the sync before the index read could answer.
 * They have to be asked in different syncs or the answer is a coin toss about
 * which one spoke.
 *
 * The id that came back for a slide added this session is
 * `4123571127#123571113` — the ADD-TIME id space this repo already documented:
 * "`4123571114#123571113` at add time, `256#2587447327` a moment later for the
 * same slide". So the question is narrow:
 *
 *   byIndex reads N > 0  ->  the shapes are there; only the ID is bad, and both
 *                            tracker drafts are describing our own stale-id bug
 *   byIndex reads 0      ->  the shapes really are unreachable, and the finding
 *                            is bigger than an id
 */
import { spawnSync } from "child_process";

const CLI = "C:\\Users\\dann_\\AppData\\Roaming\\npm\\node_modules\\@playwright\\cli\\playwright-cli.js";
const DIR = "C:\\devtools\\SSF-Charts\\.pw-session";

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

const SCRIPT = `async () => {
  const out = [];
  const rect = PowerPoint.GeometricShapeType.rectangle;

  for (let i = 0; i < 3; i++) {
    const rec = { trial: i };
    await PowerPoint.run(async (c) => { c.presentation.slides.add(); await c.sync(); });
    let index = -1;
    await PowerPoint.run(async (c) => {
      const slides = c.presentation.slides;
      slides.load("items/id");
      await c.sync();
      index = slides.items.length - 1;
      rec.idAtAdd = slides.items[index].id;
    });
    rec.index = index;

    await PowerPoint.run(async (c) => {
      const slide = c.presentation.slides.getItemAt(index);
      const s = slide.shapes.addGeometricShape(rect);
      s.left = 20; s.top = 20; s.width = 40; s.height = 30;
      await c.sync();
    });

    try {
      await PowerPoint.run(async (c) => {
        const slide = c.presentation.slides.getItemAt(index);
        slide.load("id");
        slide.shapes.load("items/id");
        await c.sync();
        rec.freshByIndex = slide.shapes.items.length;
        rec.idNow = slide.id;
        if (slide.shapes.items.length > 0) {
          slide.shapes.items[0].tags.add("SCOPEPROBE", "split-" + i);
          await c.sync();
          rec.tagByIndex = "OK";
        }
      });
    } catch (e) { rec.freshByIndex = "THREW: " + String(e && e.message).slice(0, 70); }

    try {
      await PowerPoint.run(async (c) => {
        const slide = c.presentation.slides.getItem(rec.idAtAdd);
        slide.shapes.load("items/id");
        await c.sync();
        rec.byAddTimeId = slide.shapes.items.length;
      });
    } catch (e) { rec.byAddTimeId = "THREW: " + String(e && e.message).slice(0, 45); }

    if (rec.idNow && rec.idNow !== rec.idAtAdd) {
      try {
        await PowerPoint.run(async (c) => {
          const slide = c.presentation.slides.getItem(rec.idNow);
          slide.shapes.load("items/id");
          await c.sync();
          rec.bySettledId = slide.shapes.items.length;
        });
      } catch (e) { rec.bySettledId = "THREW: " + String(e && e.message).slice(0, 45); }
    }
    out.push(rec);
  }
  return JSON.stringify(out);
}`;

const found = pw("find", "Chart");
const ref = /\[ref=([a-z0-9]+)\]/.exec(found.split("\n").find((l) => /tab "Chart"/.test(l)) ?? "");
if (!ref) {
  console.error('the pane is not open — no `tab "Chart"` ref to evaluate in.');
  process.exit(1);
}
const raw = pw("eval", SCRIPT, ref[1]);
const json = /\[\{.*\}\]/s.exec(raw.replace(/\\"/g, '"'));
console.log(json ? JSON.parse(json[0]).map((r) => JSON.stringify(r)).join("\n") : raw);
