#!/usr/bin/env node
/**
 * Audit a real PowerPoint deck for what SSF Charts actually left in it.
 *
 * Every hard diagnosis in this project was settled by reading the .pptx, not
 * by reading the add-in's own report — including the one where the report was
 * the thing that was wrong. A 39-slide web run said it had tagged 20 charts;
 * the file carried 31. The log agreed with itself and was false, and only the
 * bytes said so.
 *
 * So this reads the bytes. It knows nothing about what the run believed and
 * takes no input from it: slot tags, config tags, groups and frames as they
 * exist in the file. Point it at a deck saved out of a real host and it says
 * what is there.
 *
 *   node scripts/verify-deck.mjs <deck.pptx> [--json]
 *
 * Exit code is 0 when the deck is structurally sound (whatever the run's
 * outcome), 1 when a fault is found, 2 when the file cannot be read — so it
 * can gate a check as well as inform one.
 *
 * "Structurally sound" deliberately excludes run outcomes. A deck missing
 * eight charts because the host dropped them is a bad RUN and a fine FILE; a
 * deck whose tag part is unreferenced is a broken file. Only the second is a
 * fault here, because only the second is something this repo wrote wrong.
 */
import { readFileSync } from "fs";
import JSZip from "jszip";
import { isMain } from "./is-main.mjs";

const CONFIG_TAG = "POWERCHART_CONFIG";
const ORIGIN_TAG = "POWERCHART_ORIGIN";
const SLOT_TAG = "POWERCHART_DEMO_SLOT";
const GROUP_NAME = "PowerChart";
const NOT_COMPLETE = "PowerChart:not-complete";

/** Undo the five XML attribute-value escapes, in the one order that is safe. */
const unescapeAttr = (v) =>
  v
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");

/** Slide parts in DECK order — numeric, never lexicographic (slide10 < slide9). */
function slidePartsInOrder(zip) {
  return Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => Number(/(\d+)/.exec(a)[1]) - Number(/(\d+)/.exec(b)[1]));
}

/** The shape tree's direct children, with nesting handled — see ooxml.ts. */
function topLevel(xml) {
  const tree = /<p:spTree>([\s\S]*)<\/p:spTree>/.exec(xml);
  if (!tree) return [];
  const out = [];
  let depth = 0;
  let start = -1;
  const tag = /<(\/?)([A-Za-z0-9:_-]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let m;
  while ((m = tag.exec(tree[1]))) {
    if (m[4] === "/") {
      if (depth === 0) out.push(m[0]);
      continue;
    }
    if (m[1] === "/") {
      depth--;
      if (depth === 0 && start >= 0) {
        out.push(tree[1].slice(start, m.index + m[0].length));
        start = -1;
      }
      continue;
    }
    if (depth === 0) start = m.index;
    depth++;
  }
  // The first two children are the tree's own mandatory nvGrpSpPr/grpSpPr
  // boilerplate, not shapes.
  return out.filter((e) => !/^<p:(nvGrpSpPr|grpSpPr)[\s>]/.test(e));
}

export async function readDeck(path) {
  return readDeckBytes(readFileSync(path));
}

/** The same audit against bytes already in hand — what the tests drive. */
export async function readDeckBytes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const types = await zip.file("[Content_Types].xml")?.async("string");
  const slides = slidePartsInOrder(zip);
  const referencedTagParts = new Set();
  const rows = [];

  for (const [i, part] of slides.entries()) {
    const xml = await zip.file(part).async("string");
    const relsPath = part.replace("slides/", "slides/_rels/") + ".rels";
    const rels = (await zip.file(relsPath)?.async("string")) ?? "";
    // The relationship ID travels with the part name from here on. Which
    // element points at a tag part is the whole question below, and the old
    // code threw the rid away on this line.
    const tagRefs = [...rels.matchAll(/Id="([^"]+)"[^>]*Target="\.\.\/tags\/(tag\d+\.xml)"/g)].map((m) => ({
      rid: m[1],
      part: m[2],
    }));

    let slot = null;
    let slotMalformed = false;
    // ONE ENTRY PER CONFIG TAG ON THE SLIDE, not one for the slide. This kept a
    // single `config`, `configRaw` and `configRid` and overwrote them on every
    // tag part the relationships listed, so the audit judged whichever config
    // happened to be listed LAST and never looked at the others. A slide holds
    // several charts as a matter of course — the self-test draws two onto one
    // slide on purpose, and round 497's inventory has slides carrying 3, 2, 5
    // and 8 objects named PowerChart — and on those a corrupt or unreachable
    // config was cleared whenever a later chart's was healthy: "no structural
    // faults", exit 0, on a deck with a dead chart in it.
    const configs = [];
    let originSeen = false;
    const missingParts = [];
    for (const { rid, part: t } of tagRefs) {
      const file = zip.file(`ppt/tags/${t}`);
      if (!file) {
        missingParts.push(t);
        continue;
      }
      referencedTagParts.add(`ppt/tags/${t}`);
      const tx = await file.async("string");
      const s = new RegExp(`name="${SLOT_TAG}" val="([^"]*)"`).exec(tx);
      if (s) {
        try {
          slot = JSON.parse(unescapeAttr(s[1]));
        } catch {
          slot = { malformed: s[1] };
        }
        // Any one malformed slot, not whichever was read last: a healthy slot
        // listed after a corrupt one must not clear it.
        if (typeof slot?.malformed === "string") slotMalformed = true;
      }
      const c = new RegExp(`name="${CONFIG_TAG}" val="([^"]*)"`).exec(tx);
      if (c) {
        const raw = unescapeAttr(c[1]);
        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch {
          parsed = { malformed: true };
        }
        // `origin` is whether THIS chart's own tag part names an origin: one
        // shape references one tags part and the writer puts config and origin
        // together in it, so another chart's origin says nothing about this one.
        configs.push({ rid, raw, malformed: parsed?.malformed === true, origin: tx.includes(ORIGIN_TAG) });
      }
      if (tx.includes(ORIGIN_TAG)) originSeen = true;
    }

    const shapes = topLevel(xml);
    const groups = shapes.filter((e) => e.startsWith("<p:grpSp"));
    // WHICH element the config tag hangs off, not merely that the slide's
    // relationship part mentions it.
    //
    // Shape-level and slide-level tags are declared in the same `.rels` file,
    // so a check that reads only that file cannot tell "the tag is on the
    // chart's `p:grpSp`" (re-editable) from "the tag part exists and nothing in
    // the shape tree points at it" (dead). Those are the same measurement and
    // opposite outcomes, and this tool's "N re-editable" line is what a
    // real-host deck was cleared on. `ooxml.ts` has a branch that mints the rid,
    // writes the part, adds the relationship and then consumes it only
    // `if (item.group !== false)` — so the invisible case is reachable today.
    //

    // The chart object: a real group, or the single named shape a degraded
    // picture leaves behind. Both are "one object named PowerChart".
    const named = shapes.filter((e) => new RegExp(`name="${GROUP_NAME}"`).test(e));
    // Where the config tag is actually anchored. `configOnChart` is the one
    // that means re-editable; anything else is a tag the pane will never find.
    //
    // A substring test over the chart object's own markup, deliberately, rather
    // than a full parse: the writer emits
    // `<p:nvPr><p:custDataLst><p:tags r:id="…"/></p:custDataLst></p:nvPr>`
    // inside that object, so "does this object's markup reference the rid" IS
    // the question. Relationship ids are unique within a slide part, so a match
    // cannot belong to anything else.
    const anchoredOn = (e, rid) => e.includes(`<p:tags r:id="${rid}"/>`);
    // EVERY config tag has to be anchored for the slide to count as re-editable,
    // and ONE that is not makes it an orphan. Judging only the last tag listed
    // is how a dead chart beside a healthy one read as sound.
    const configOnChart = configs.length > 0 && configs.every((c) => named.some((e) => anchoredOn(e, c.rid)));
    const configOnSlide = configs.length > 0 && !configOnChart;
    // PER OBJECT, for the report's count. The rows are one per SLIDE and a slide
    // can hold several objects named PowerChart, so counting rows undercounted
    // them — and a chart that kept its group but lost its tag, the recurring
    // "1 of 2 new charts are re-editable" round failure (BACKLOG; rounds 102,
    // 467, 497), printed as "chart objects: 1 (1 re-editable, 0 NOT)".
    const chartsReEditable = named.filter((e) => configs.some((c) => anchoredOn(e, c.rid))).length;
    // How many shapes are INSIDE the chart object. `shapes` above counts the
    // slide's top-level children, where a 40-shape chart and a 1-shape
    // degraded picture both read as 1 — which is the number a reader most
    // often wants and the one this could not previously give.
    const chartShapes = named.reduce((n, e) => {
      const all = (e.match(/<p:(sp|pic|cxnSp|grpSp)[\s>]/g) ?? []).length;
      // A group's own opening tag matches too; a degraded picture IS the one
      // shape and has no container to discount.
      return n + (e.startsWith("<p:grpSp") ? all - 1 : all);
    }, 0);
    rows.push({
      index: i,
      slot: slot?.i ?? null,
      // Its own flag, exactly like `configMalformed` below, because the row
      // cannot carry it otherwise: a `{malformed}` marker has no `i`, so
      // `slot?.i ?? null` collapses it to null and the fault check —
      // `r.slot !== null && typeof r.slot !== "number"` — could never fire for
      // the one case it was written for. A deck this repo emitted with a corrupt
      // slot tag was certified "no structural faults" and exited 0; triage then
      // counted the slide as carrying no slot tag and reported its item LOST,
      // blaming the host for dropping a slide sitting right there.
      slotMalformed,
      title: slot?.title ?? null,
      run: slot?.run ?? null,
      shapes: shapes.length,
      chartShapes,
      groups: groups.length,
      chartObject: named.length > 0,
      /**
       * How many objects named PowerChart the slide holds, and how many of them
       * a config tag is anchored on. The two booleans around them say whether
       * the slide has ANY; these are what the report's totals are summed from.
       */
      chartObjects: named.length,
      chartsReEditable,
      picture: named.length > 0 && groups.length === 0,
      config: configs.length > 0,
      /**
       * The config tag is anchored on the chart object, so the pane can find it.
       *
       * `config` above only says the tag part EXISTS and the slide's
       * relationship part mentions it — which is equally true of a tag nothing
       * in the shape tree points at. Every "re-editable" count this tool prints
       * comes from here now. With several config tags on the slide it means ALL
       * of them are anchored.
       */
      configOnChart,
      /** The tag exists and hangs off nothing the pane reads — a dead config. */
      configOrphaned: configOnSlide,
      // The tag's actual payload, for callers that need to reproduce the deck
      // rather than audit it — the fake host in the test suite builds its
      // slides from this, so it cannot decode a generated deck differently
      // from the tool that checks one. Stripped from `--json`, which is a
      // report and would otherwise carry a full ChartConfig per slide.
      //
      // The LAST config listed, as it always was: a reproduction needs one
      // payload per slide and every consumer builds one chart per slide. The
      // AUDIT fields around it cover all of them.
      configJson: configs.at(-1)?.raw ?? null,
      configMalformed: configs.some((c) => c.malformed),
      origin: configs.length ? configs.every((c) => c.origin) : originSeen,
      stamped: xml.includes(NOT_COMPLETE),
      missingTagParts: missingParts,
      duplicateShapeIds: duplicateShapeIds(xml),
    });
  }

  const allTagParts = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
  return { rows, allTagParts, referencedTagParts, types: types ?? "" };
}

/**
 * Non-unique `<p:cNvPr id="…">` values on one slide.
 *
 * The classic trigger for PowerPoint's "repair" dialog, and invisible to
 * everything else here: the schema does not express uniqueness, so the OOXML
 * validator passes a slide with two shapes sharing an id (measured — injected
 * `id="1"` twice and it reported nothing), and a shape-tree walk that only
 * counts elements never compares them.
 *
 * Ids must be unique per slide, INCLUDING inside groups, so this reads every
 * `cNvPr` in the part rather than only the top level. Non-numeric or absent
 * ids are ignored — those are a different kind of wrong and not this check's.
 */
function duplicateShapeIds(xml) {
  const seen = new Set();
  const dupes = new Set();
  for (const m of xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)) {
    if (seen.has(m[1])) dupes.add(m[1]);
    else seen.add(m[1]);
  }
  return [...dupes];
}

/** Faults are things THIS REPO wrote wrong — never things the host did badly. */
export function faultsIn({ rows, allTagParts, referencedTagParts, types }) {
  const faults = [];
  for (const r of rows) {
    if (r.missingTagParts.length)
      faults.push(`slide ${r.index + 1}: references missing tag part(s) ${r.missingTagParts.join(", ")}`);
    if (r.configMalformed) faults.push(`slide ${r.index + 1}: ${CONFIG_TAG} is not valid JSON`);
    if (r.duplicateShapeIds?.length)
      faults.push(
        `slide ${r.index + 1}: shape id(s) ${r.duplicateShapeIds.join(", ")} used more than once — PowerPoint will offer to repair the file`,
      );
    if (r.slotMalformed || (r.slot !== null && typeof r.slot !== "number"))
      faults.push(`slide ${r.index + 1}: slot tag is malformed`);
    // A config with nothing to hang it on would be unreachable from the pane.
    if (r.config && !r.chartObject)
      faults.push(`slide ${r.index + 1}: carries a config tag but no "${GROUP_NAME}" object to load it from`);
    // And a config the chart object does not point at is just as unreachable,
    // while looking identical from the relationship part. This is the case the
    // old check could not see at all: it was satisfied by the mere existence of
    // an element named PowerChart anywhere at top level, so a regression that
    // put the r:id on the slide's own `custDataLst` instead of the group's
    // would have reported "re-editable", exited 0, kept CI green, and left
    // every generated chart in the deck dead to the pane.
    if (r.configOrphaned)
      faults.push(
        `slide ${r.index + 1}: ${CONFIG_TAG} is not referenced by the "${GROUP_NAME}" object — the pane cannot load it`,
      );
  }
  for (const p of allTagParts) {
    if (!referencedTagParts.has(p)) faults.push(`${p}: tag part is not referenced by any slide`);
    if (types && !types.includes(`PartName="/${p}"`))
      faults.push(`${p}: no [Content_Types].xml override — PowerPoint will reject the file`);
  }
  // Two runs' slides in one deck is NOT a fault — it is the case the run token
  // exists to survive — but a repeated (run, slot) pair means two slides claim
  // to be the same item of the same run, which nothing should produce.
  const seen = new Map();
  for (const r of rows) {
    if (r.run === null || r.slot === null) continue;
    const key = `${r.run}#${r.slot}`;
    if (seen.has(key)) faults.push(`slides ${seen.get(key) + 1} and ${r.index + 1}: same run and slot`);
    else seen.set(key, r.index);
  }
  return faults;
}

/**
 * The report's chart-object totals, summed over OBJECTS and not over slides.
 *
 * The rows are one per slide, and counting them reported "chart objects: 1" for
 * a slide holding two. Exported because the totals are what a real-host deck is
 * cleared on (`docs/REGRESSION.md` section 0), so they are worth testing without
 * spawning the CLI.
 */
export function chartTotals(rows) {
  const objects = rows.reduce((n, r) => n + r.chartObjects, 0);
  const reEditable = rows.reduce((n, r) => n + r.chartsReEditable, 0);
  return { objects, reEditable };
}

function report(deck, faults) {
  const { rows } = deck;
  const w = (s, n) =>
    String(s ?? "-")
      .padEnd(n)
      .slice(0, n);
  console.log(`${rows.length} slide(s)\n`);
  console.log(`  ${w("#", 4)}${w("title", 26)}${w("slot", 6)}${w("shapes", 7)}${w("object", 9)}${w("config", 7)}flags`);
  for (const r of rows) {
    const flags = [
      r.stamped && "NOT-COMPLETE",
      r.config && !r.origin && "no-origin",
      r.configOrphaned && "ORPHANED-CONFIG",
    ]
      .filter(Boolean)
      .join(" ");
    console.log(
      `  ${w(r.index + 1, 4)}${w(r.title, 26)}${w(r.slot, 6)}${w(r.shapes, 7)}` +
        `${w(r.picture ? "picture" : r.groups ? "group" : r.chartObject ? "shape" : "loose", 9)}` +
        // `2/3` on a slide holding three objects of which two are re-editable:
        // a single yes/no cannot say which of several charts lost its tag.
        `${w(r.chartObjects > 1 ? `${r.chartsReEditable}/${r.chartObjects}` : r.configOnChart ? "yes" : r.config ? "orphan" : "—", 7)}${flags}`,
    );
  }
  const runs = [...new Set(rows.map((r) => r.run).filter(Boolean))];
  const { objects, reEditable } = chartTotals(rows);
  console.log(
    `\n  runs in deck: ${runs.length ? runs.join(", ") : "none"}` +
      `\n  chart objects: ${objects} (${reEditable} re-editable, ${objects - reEditable} NOT re-editable)` +
      `\n  pictures: ${rows.filter((r) => r.picture).length}` +
      `\n  stamped NOT COMPLETE: ${rows.filter((r) => r.stamped).length}`,
  );
  if (faults.length) {
    console.log(`\n  ${faults.length} FAULT(S):`);
    for (const f of faults) console.log(`   - ${f}`);
  } else {
    console.log(`\n  no structural faults`);
  }
}

// CLI only when invoked directly; importable as a library for the tests.
//
// `isMain`, not `endsWith` — the identical guard as `scripts/triage.mjs` had,
// and found the same way on 2026-09-28: `test/is-main.test.ts` only admitted
// files mentioning `import.meta.url`, and this spelling never does, so the
// sweep written to catch hand-rolled CLI guards could see neither of them.
const invokedDirectly = isMain(import.meta.url, process.argv[1]);
const [, , path, ...flags] = process.argv;
if (!invokedDirectly) {
  // imported — nothing to do
} else if (!path) {
  console.error("usage: node scripts/verify-deck.mjs <deck.pptx> [--json]");
  process.exit(2);
} else {
  let deck;
  try {
    deck = await readDeck(path);
  } catch (err) {
    console.error(`could not read ${path}: ${err?.message ?? err}`);
    process.exit(2);
  }
  const faults = faultsIn(deck);
  if (flags.includes("--json")) {
    // `configJson` is for programmatic callers, not for a report — a full
    // ChartConfig per slide would bury everything else in the output.
    const slides = deck.rows.map(({ configJson: _payload, ...row }) => row);
    console.log(JSON.stringify({ slides, faults }, null, 2));
  } else {
    report(deck, faults);
  }
  process.exit(faults.length ? 1 : 0);
}
