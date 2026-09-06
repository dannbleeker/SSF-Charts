/**
 * Arm D/E again, with the arm's own bug removed.
 *
 * The first attempt reported six THREWs reading "Cannot read properties of
 * undefined (reading 'tags')" — which is not the host refusing anything, it is
 * this script indexing an EMPTY shape collection. `listed` was 0 on a slide a
 * shape had just been drawn and synced onto.
 *
 * That is a finding rather than a nuisance: it is the short/empty collection
 * read Draft A lists as its second symptom, and it fired on every one of six
 * freshly added slides while arm B — one added slide, earlier, on a quieter
 * deck — read 1 and tagged fine.
 *
 * So this version separates the two outcomes the first one conflated:
 *
 *   empty-read      the collection came back with 0 shapes; nothing to tag
 *   empty-then-N    it came back 0, and a SECOND read found N (so it settles)
 *   tag OK / THREW  what the tag write actually did, when there was a shape
 *
 * Without that split, "the host refused the tag" and "we asked too early" are
 * the same line, and the first is what a tracker issue would have claimed.
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

  const drawAndTag = async (label, index) => {
    const rec = { arm: label, index };
    try {
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(index);
        const s = slide.shapes.addGeometricShape(rect);
        s.left = 20; s.top = 20; s.width = 40; s.height = 30;
        await context.sync();
        slide.shapes.load("items/id");
        await context.sync();
        rec.listed = slide.shapes.items.length;
        if (rec.listed === 0) {
          // ASK AGAIN before calling it a refusal. A collection that fills in
          // on a second read is a timing fact; one that stays empty is a
          // different claim entirely.
          slide.shapes.load("items/id");
          await context.sync();
          rec.listedAgain = slide.shapes.items.length;
        }
        const items = slide.shapes.items;
        if (items.length === 0) { rec.tag = "no-shape-to-tag"; return; }
        items[items.length - 1].tags.add("SCOPEPROBE", label);
        await context.sync();
        rec.tag = "OK";
      });
    } catch (e) {
      rec.tag = "THREW";
      rec.code = e && e.code;
      rec.message = String(e && e.message).slice(0, 120);
    }
    out.push(rec);
  };

  for (let i = 0; i < 3; i++) {
    let unsettled = -1;
    await PowerPoint.run(async (c) => {
      const slides = c.presentation.slides;
      slides.add();
      await c.sync();
      slides.load("items/id");
      await c.sync();
      unsettled = slides.items.length - 1;
    });
    await drawAndTag("d-unsettled-" + i, unsettled);

    await PowerPoint.run(async (c) => { c.presentation.slides.add(); await c.sync(); });
    let settled = -1;
    await PowerPoint.run(async (c) => {
      const slides = c.presentation.slides;
      slides.load("items/id");
      await c.sync();
      settled = slides.items.length - 1;
    });
    await drawAndTag("e-settled-" + i, settled);
  }

  // The controls that mattered last time, repeated at the END so they are
  // measured on the same tired deck the arms above just used rather than on a
  // fresh one. A control that only holds when things are calm is not a control.
  await drawAndTag("a-document-own", 0);
  await drawAndTag("c-previous-session", 5);
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
