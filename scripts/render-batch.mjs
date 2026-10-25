#!/usr/bin/env node
/**
 * Batch automation (the .ppttc idea, open): render chart configs from a JSON
 * file to SVGs without PowerPoint.
 *
 *   npm run build:lib
 *   node scripts/render-batch.mjs examples/charts.json out/
 *
 * The JSON file holds one ChartConfig or an array of them (see README for
 * the schema; `sampleConfig` in the library produces valid examples).
 *
 * build-skill.mjs copies this file into the skill package as
 * scripts/render-svg.mjs, so every user-facing mention of the script name is
 * taken from argv — a usage line naming render-batch.mjs pointed the skill's
 * user at a file that does not exist in the zip.
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { basename, join } from "node:path";
import { buildChart, buildAgendaScene, sceneToSvg, DEFAULT_SIZE } from "../dist-lib/ssf-charts.js";

const [, self, input, outDir = "out"] = process.argv;
const usage = `usage: node scripts/${basename(self ?? "render-batch.mjs")} <charts.json> [outDir]`;
// `--help` was taken for the input file ("Could not read --help as JSON: ENOENT"),
// the usage line printing only when argv held nothing at all.
if (input === "--help" || input === "-h") {
  console.log(
    `${usage}\n\n  charts.json  one ChartConfig, or an array of them\n  outDir       where the .svg previews go (created if missing; default out)`,
  );
  process.exit(0);
}
if (!input) {
  console.error(usage);
  process.exit(1);
}

let raw;
try {
  // A UTF-8 byte-order mark is stripped first: Windows PowerShell 5.1's
  // `Set-Content -Encoding UTF8` and Notepad's "UTF-8 with BOM" write one,
  // `readFileSync(…, "utf8")` keeps it, and JSON.parse then fails on an
  // invisible U+FEFF that no editor shows. `render-pptx.mjs` reads the same way.
  // Tested by code point: an invisible character spelled out in a regex literal
  // is the same trap in this file.
  const text = readFileSync(input, "utf8");
  raw = JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);
} catch (err) {
  console.error(`Could not read ${input} as JSON: ${err.message}`);
  process.exit(1);
}
const configs = Array.isArray(raw) ? raw : [raw];
try {
  mkdirSync(outDir, { recursive: true });
} catch (err) {
  // An `outDir` that is an existing FILE (EEXIST), or sits under one (ENOTDIR),
  // ended the process in Node's raw stack dump; `render-pptx.mjs` words its own
  // write failure and this is the same step.
  console.error(`Could not create ${outDir}: ${err?.message ?? err} — give a folder path, not a file`);
  process.exit(1);
}

/** An agenda is a slide of text, not a chart — see `backgroundFor`. */
const isAgenda = (cfg) => !!cfg && cfg.kind === "agenda";

/**
 * The canvas a preview is drawn on: the chart's OWN `style.background`, as the
 * pane's preview (`canvasColor`) and the .pptx slide (`render-pptx.mjs`) both
 * paint it, and white when the chart names none.
 *
 * A hard-coded white put a dark-styled chart's near-white ink on a white rect
 * (about 1.1:1) — a blank-looking preview of a deck slide that was fine — and,
 * the other way round, hid that a config naming ONLY a dark `style.background`
 * keeps its default #0b0b0b ink and is unreadable on the slide it paints. This
 * is the QA output `SKILL.md` tells an agent to inspect and to fix the config
 * from, so it has to show what the deck shows. `sceneToSvg` falls back to white
 * itself for a paint it cannot read.
 *
 * An agenda is drawn on white whatever it is given: `buildAgendaScene` takes no
 * style and inks a fixed palette for a white canvas, and the .pptx paints it so.
 */
const backgroundFor = (cfg) => (isAgenda(cfg) ? "#ffffff" : (cfg?.style?.background ?? "#ffffff"));

/** Build the scene for one config: an agenda slide, or a chart. Mirrors render-pptx.mjs's sceneFor. */
function sceneFor(cfg) {
  if (isAgenda(cfg)) {
    // { kind:"agenda", chapters:[…], highlight?, title? } — buildChart has no
    // such kind, so routing it there threw and took the whole batch with it.
    return buildAgendaScene(Array.isArray(cfg.chapters) ? cfg.chapters : [], {
      highlight: cfg.highlight,
      title: cfg.title,
      width: cfg.width,
      height: cfg.height,
    });
  }
  return buildChart({ ...DEFAULT_SIZE, ...cfg });
}

let failed = 0;
configs.forEach((cfg, i) => {
  // Isolate each config, as the pptx renderer does: one bad chart in a batch
  // must not throw away the previews of all the others.
  try {
    const svg = sceneToSvg(sceneFor(cfg), { background: backgroundFor(cfg) });
    // `String(...)`, because a NUMERIC title is what someone writes for a year —
    // `svg.ts` and `chart-hostile-input.test.ts` both name it as ordinary rather
    // than exotic. This line called `.replace` on the raw field, and the chart
    // had ALREADY rendered by then: the throw is at the filename, so a perfectly
    // good preview was discarded and the CLI printed an engine message about
    // `.replace` not being a function. This file ships in the skill zip as
    // `render-svg.mjs`, the command SKILL.md tells an agent to run for previews,
    // so an agent titling a chart `2024` got its deck and no preview, with
    // nothing saying why.
    const name = String(cfg?.title ?? cfg?.kind ?? `chart-${i}`)
      .replace(/[^\w-]+/g, "-")
      .toLowerCase();
    const file = join(outDir, `${String(i + 1).padStart(2, "0")}-${name}.svg`);
    writeFileSync(file, svg);
    console.log(file);
  } catch (err) {
    failed++;
    console.error(`chart ${i + 1}: ${err?.message ?? err}`);
  }
});
if (failed) process.exit(1);
