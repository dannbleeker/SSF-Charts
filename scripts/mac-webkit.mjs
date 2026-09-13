/**
 * THE NEAREST THING TO A MAC THAT THIS PROJECT CAN RUN.
 *
 * Every measurement SSF Charts has ever taken is on a Chromium engine: 430+
 * archived rounds in PowerPoint on the web, and the single Windows desktop
 * reading in `docs/evidence/windows-desktop-2026-09-13.json`, which is Edge
 * WebView2. Microsoft's own table (`browsers-used-by-office-web-add-ins`) says
 * Office on Mac runs add-ins in **Safari with WKWebView** — JavaScriptCore and
 * WebKit. So Mac is not merely a platform nobody has clicked; it is the only
 * ENGINE FAMILY on which not one line of this code has ever executed.
 *
 * This script executes the chart engine on WebKit and diffs it against Node.
 *
 * WHAT IT DOES NOT DO, stated first so the result is not over-read:
 *
 *   - It is NOT PowerPoint on Mac. There is no Office.js, no host, no slide.
 *     Everything in `src/render/powerpoint.ts` past scene construction — every
 *     `context.sync()`, the group call, the picture fallback — is untouched.
 *     A green run here says the geometry and the labels agree, not that the
 *     add-in works on a Mac.
 *   - It is PLAYWRIGHT's WebKit, which tracks WebKit trunk, not the WKWebView
 *     shipped in any particular macOS. It therefore measures the CURRENT
 *     engine. It cannot tell you about the oldest Mac the manifest admits.
 *   - n = 1 engine, one build, one machine. Same caveat as the Windows file.
 *
 * WHAT IT IS GOOD FOR, and why it was written rather than assumed:
 *
 *   1. `Intl.NumberFormat` (src/core/format.ts) formats EVERY number label in
 *      every chart. It is backed by ICU, and the ICU in JavaScriptCore is not
 *      the ICU in V8. A separator or a space that differs by engine is a
 *      different chart on a Mac, silently, on every slide.
 *   2. REGEX LOOKBEHIND. `src/render/svg.ts` and `src/core/format.ts` both use
 *      it. Safari did not support lookbehind until 16.4, and an unsupported
 *      lookbehind is a SyntaxError at PARSE time — it does not throw where it
 *      is used, it takes the whole module down. That is the blank-pane failure
 *      mode, and it is the one most likely to fail an AppSource review on Mac.
 *   3. It re-derives the "below 1.10" counts instead of citing them. The
 *      published figure was 8 charts until 2026-09-13, when running this said
 *      9: showcase #107 is a radar whose eight radial bars are eight `wedge`
 *      nodes, and it had been filed for a fortnight under the charts that lose
 *      only their arrows. A recipe catches that; a remembered number does not.
 *
 * Usage:  npm run mac:webkit          (needs `npx playwright-core install webkit` once)
 * Exit:   0 agree · 1 divergence · 2 could not run
 */

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { isMain } from "./is-main.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
const LIB = resolve(ROOT, "dist-lib/ssf-charts.js");
const SHOWCASE = resolve(ROOT, "examples/showcase.json");

/**
 * The comparison unit. Geometry is rounded to 1e-6 because the two engines are
 * both IEEE-754 and a last-bit difference in a transcendental is not a finding
 * — but it is rounded rather than dropped, because a WEDGE ANGLE that differs
 * in the fourth decimal IS one, and that is the class of thing this exists to
 * catch. Text is compared verbatim: that is where Intl lives and where an
 * engine difference is user-visible.
 */
const DIGEST = `(scene) => scene.nodes.map((n) => {
  const num = (v) => (typeof v === "number" ? Math.round(v * 1e6) / 1e6 : v);
  const out = { k: n.kind };
  for (const key of Object.keys(n).sort()) {
    if (key === "kind") continue;
    const v = n[key];
    if (typeof v === "number") out[key] = num(v);
    else if (typeof v === "string") out[key] = v;
    else if (Array.isArray(v)) out[key] = v.map((e) => (typeof e === "number" ? num(e) : typeof e === "object" && e ? JSON.stringify(Object.fromEntries(Object.entries(e).map(([a, b]) => [a, typeof b === "number" ? num(b) : b]))) : e));
  }
  return out;
})`;

/** Text only — isolated so an Intl divergence reports as itself, not as "a node differs". */
const TEXTS = `(scene) => scene.nodes.filter((n) => typeof n.text === "string").map((n) => n.text)`;

async function main() {
  let showcase, samples, lib;
  try {
    showcase = JSON.parse(readFileSync(SHOWCASE, "utf8"));
    lib = readFileSync(LIB, "utf8");
  } catch (err) {
    console.error(`Cannot read inputs — run \`npm run build:lib\` first.\n  ${err.message}`);
    process.exit(2);
  }

  const mod = await import(`file://${LIB.replace(/\\/g, "/")}`);
  samples = mod.CHART_KINDS.map((k) => mod.sampleConfig(k));
  const CASES = [
    ...showcase.map((cfg, i) => ({ id: `showcase#${i}`, cfg })),
    ...samples.map((cfg, i) => ({ id: `sample:${mod.CHART_KINDS[i]}`, cfg })),
  ];

  // ---- Node side -----------------------------------------------------------
  const digestFn = eval(DIGEST);
  const textsFn = eval(TEXTS);
  const node = new Map();
  for (const c of CASES) {
    try {
      const scene = mod.buildChart(c.cfg);
      node.set(c.id, { d: JSON.stringify(digestFn(scene)), t: JSON.stringify(textsFn(scene)) });
    } catch (err) {
      node.set(c.id, { d: `THREW:${err.message}`, t: `THREW:${err.message}` });
    }
  }

  // ---- WebKit side ---------------------------------------------------------
  let webkit;
  try {
    ({ webkit } = await import("playwright-core"));
  } catch {
    console.error("playwright-core not resolvable. `npm install`.");
    process.exit(2);
  }

  let browser;
  try {
    browser = await webkit.launch();
  } catch (err) {
    console.error(
      "WebKit will not launch. Install it once with:\n" +
        "    npx playwright-core install webkit\n" +
        `  ${err.message}`,
    );
    process.exit(2);
  }

  const page = await browser.newPage();
  const pageErrors = [];
  page.on("pageerror", (e) => pageErrors.push(String(e)));
  await page.setContent("<!doctype html><title>mac-webkit</title>");

  const engine = await page.evaluate(() => ({
    ua: navigator.userAgent,
    // Lookbehind is a PARSE-time feature: it must be probed through a
    // constructor, because a literal in this function would fail to parse on an
    // engine that lacks it and take the probe down with it.
    lookbehind: (() => {
      try {
        return new RegExp("(?<!a)b").test("cb");
      } catch (e) {
        return `UNSUPPORTED: ${e.name}`;
      }
    })(),
  }));

  const wk = await page.evaluate(
    async ({ src, cases, digestSrc, textsSrc }) => {
      const url = URL.createObjectURL(new Blob([src], { type: "text/javascript" }));
      let m;
      try {
        m = await import(url);
      } catch (err) {
        return { fatal: `module failed to load: ${err && err.message}` };
      }
      const digest = eval(digestSrc);
      const texts = eval(textsSrc);
      const out = {};
      for (const c of cases) {
        try {
          const scene = m.buildChart(c.cfg);
          out[c.id] = { d: JSON.stringify(digest(scene)), t: JSON.stringify(texts(scene)) };
        } catch (err) {
          out[c.id] = { d: `THREW:${err.message}`, t: `THREW:${err.message}` };
        }
      }
      return { out };
    },
    { src: lib, cases: CASES, digestSrc: DIGEST, textsSrc: TEXTS },
  );

  // ---- The below-1.10 recount, in WebKit, so the doc figure is re-derivable --
  const drop = await page.evaluate(
    async ({ src, cases }) => {
      const url = URL.createObjectURL(new Blob([src], { type: "text/javascript" }));
      const m = await import(url);
      const subject = {},
        arrows = {},
        total = {};
      const hits = [];
      for (const c of cases) {
        if (!c.id.startsWith("showcase#")) continue;
        total[c.cfg.kind] = (total[c.cfg.kind] || 0) + 1;
        let scene;
        try {
          scene = m.buildChart(c.cfg);
        } catch {
          continue;
        }
        const w = scene.nodes.filter((n) => n.kind === "wedge").length;
        const a = scene.nodes.filter((n) => n.kind === "arrowhead").length;
        if (!w && !a) continue;
        hits.push({ id: c.id, kind: c.cfg.kind, w, a, title: c.cfg.title || "" });
        if (w) subject[c.cfg.kind] = (subject[c.cfg.kind] || 0) + 1;
        else arrows[c.cfg.kind] = (arrows[c.cfg.kind] || 0) + 1;
      }
      return { subject, arrows, total, hits };
    },
    { src: lib, cases: CASES },
  );

  await browser.close();

  // ---- Report --------------------------------------------------------------
  console.log("SSF Charts — the chart engine on WebKit (the Mac engine family)");
  console.log(`  engine     ${engine.ua}`);
  console.log(`  lookbehind ${engine.lookbehind === true ? "supported" : engine.lookbehind}`);
  console.log(`  cases      ${CASES.length} (${showcase.length} showcase + ${samples.length} kind samples)`);
  console.log("");

  if (wk.fatal) {
    console.error(`FATAL — the library did not even load under WebKit:\n  ${wk.fatal}`);
    console.error("  This is the blank-pane failure mode. Nothing below ran.");
    process.exit(1);
  }

  const geomDiff = [],
    textDiff = [];
  for (const c of CASES) {
    const a = node.get(c.id),
      b = wk.out[c.id];
    if (!b) {
      geomDiff.push({ id: c.id, why: "missing from WebKit run" });
      continue;
    }
    if (a.t !== b.t) textDiff.push({ id: c.id, node: a.t, wk: b.t });
    else if (a.d !== b.d) geomDiff.push({ id: c.id, why: "geometry differs" });
  }

  const fmt = (o) =>
    Object.entries(o)
      .sort((x, y) => y[1] - x[1])
      .map(([k, v]) => `${k} ${v}/${drop.total[k]}`)
      .join(", ");
  const nSubject = Object.values(drop.subject).reduce((a, b) => a + b, 0);
  const nArrows = Object.values(drop.arrows).reduce((a, b) => a + b, 0);

  console.log("BELOW PowerPointApi 1.10 (re-derived here, not cited):");
  console.log(`  ${nSubject + nArrows} of ${showcase.length} shipped charts lose ink`);
  console.log(`  ${nSubject} lose their SUBJECT   — ${fmt(drop.subject)}`);
  console.log(`  ${nArrows} lose only ANNOTATION  — ${fmt(drop.arrows)}`);
  console.log("");

  if (pageErrors.length) {
    console.log("UNCAUGHT PAGE ERRORS:");
    for (const e of pageErrors) console.log(`  ${e}`);
    console.log("");
  }

  if (!geomDiff.length && !textDiff.length) {
    console.log(`AGREES WITH NODE on all ${CASES.length} cases — geometry and every text label.`);
    console.log("Reminder: this is the ENGINE, not PowerPoint on Mac. No Office.js ran.");
    process.exit(0);
  }

  if (textDiff.length) {
    console.log(`TEXT DIVERGES on ${textDiff.length} case(s) — suspect Intl/ICU:`);
    for (const d of textDiff.slice(0, 10)) {
      console.log(`  ${d.id}`);
      console.log(`    node   ${d.node.slice(0, 300)}`);
      console.log(`    webkit ${d.wk.slice(0, 300)}`);
    }
    if (textDiff.length > 10) console.log(`  … and ${textDiff.length - 10} more`);
  }
  if (geomDiff.length) {
    console.log(`GEOMETRY DIVERGES on ${geomDiff.length} case(s):`);
    for (const d of geomDiff.slice(0, 20)) console.log(`  ${d.id} — ${d.why}`);
    if (geomDiff.length > 20) console.log(`  … and ${geomDiff.length - 20} more`);
  }
  process.exit(1);
}

if (isMain(import.meta.url, process.argv[1])) {
  main().catch((err) => {
    console.error(err);
    process.exit(2);
  });
}
