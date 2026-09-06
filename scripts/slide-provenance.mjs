#!/usr/bin/env node
/**
 * Where does this add-in actually draw, and where do its 5010s land?
 *
 *     node scripts/slide-provenance.mjs
 *
 * WRITTEN BECAUSE A CLAIM LOOKED CONFIRMED AND WAS NOT. On 2026-09-06 a
 * controlled experiment pinned `InvalidParam passed to GetItem(id)` (code
 * 5010) to a slide's PROVENANCE — a shape drawn on a slide the add-in
 * introduced cannot be tagged from a later run, while the document's own
 * slides are fine. The archive appeared to agree: 1,361 of 1,369 5010-bearing
 * events sit on a slide the add-in had added that round, and none on a slide
 * the document already had.
 *
 * It agrees with everything, because the add-in has drawn on a slide it did
 * not itself add TWICE in ~9,600 batches. Every scenario adds a slide and
 * draws on it, so there is no second arm for the comparison to run against.
 * This script prints both halves together — the failures AND the denominator
 * — so the second can never be left out of the first again.
 *
 * NO ROUNDS NEEDED. The slide id was already on file inside
 * `debugInfo.fullStatements` (`slides.getItem("262#1236456497")`), which is why
 * this reads 385 archived rounds rather than waiting for new ones.
 */
import { readdirSync, readFileSync } from "fs";
import { isMain } from "./is-main.mjs";

/**
 * The id is double-escaped: `debugInfo` is a JSON string INSIDE a data field,
 * so a stringified event carries `getItem(\\"262#1236456497\\")`. Matching a
 * single backslash finds nothing, which is how the first pass of this read 0
 * of 1,369 and looked like an absence of evidence.
 */
const SLIDE_ID = /slides\.getItem\(\\*"(\d+#\d+)/;

/** Did this event mention Office.js error code 5010, however it was wrapped? */
export function mentions5010(data) {
  const s = JSON.stringify(data ?? null);
  return s.includes('"5010"') || s.includes("code=5010");
}

/** The slide an error's own statement list says it was working on. */
export function slideOf(data) {
  return SLIDE_ID.exec(JSON.stringify(data ?? null))?.[1] ?? null;
}

/**
 * Which of the two kinds of slide this is, for one round.
 *
 * `added` — the add-in put it there this round. `own` — the document had it.
 * `unknown` covers an id from neither list, which happens when a slide was
 * swept before the round's inventory was taken; it is reported rather than
 * folded into either arm, because an unplaceable id must not pad a count.
 */
export function classify(id, added, all) {
  if (!id) return "unnamed";
  if (added.has(id)) return "added";
  if (all.has(id)) return "own";
  return "unknown";
}

export function survey(dir = "rounds", list = readdirSync, read = readFileSync) {
  const files = list(dir)
    .filter((f) => /^\d+-/.test(f))
    .sort();
  const draws = { added: 0, own: 0, sentinel: 0, unknown: 0 };
  const failures = { added: 0, own: 0, unknown: 0, unnamed: 0 };
  let rounds = 0;
  let events = 0;

  for (const f of files) {
    let j;
    try {
      j = JSON.parse(read(`${dir}/${f}`, "utf8"));
    } catch {
      continue;
    }
    if (!j.deck?.inventory) continue;
    rounds++;
    const added = new Set(j.deck.newSlides || []);
    const all = new Set(j.deck.inventory.map((s) => s.slideId));

    for (const e of j.trace?.entries || []) {
      if (e.message === "batch issued") {
        const k = e.data?.onSlideKey;
        if (typeof k === "string" && k.startsWith("(")) draws.sentinel++;
        else if (added.has(k)) draws.added++;
        else if (all.has(k)) draws.own++;
        else draws.unknown++;
      }
      if (!mentions5010(e.data)) continue;
      events++;
      failures[classify(slideOf(e.data), added, all)]++;
    }
  }
  return { rounds, events, draws, failures };
}

// BOTH ARGUMENTS. `isMain(import.meta.url)` alone returns false on Windows and
// the script then prints nothing and exits 0, which reads as a pass — the exact
// failure `is-main.mjs` was written about, committed again on first use.
if (isMain(import.meta.url, process.argv[1])) {
  const { rounds, events, draws, failures } = survey();
  console.log(`${rounds} archived round(s)\n`);
  console.log("  WHERE THIS ADD-IN DRAWS (batch issued)");
  console.log(`    on a slide it added that round   ${String(draws.added).padStart(6)}`);
  console.log(`    on a slide the document had      ${String(draws.own).padStart(6)}`);
  console.log(`    the (visible) sentinel           ${String(draws.sentinel).padStart(6)}`);
  console.log(`    unplaceable                      ${String(draws.unknown).padStart(6)}\n`);
  console.log(`  WHERE ITS 5010s LAND (${events} event(s))`);
  console.log(`    on a slide it added that round   ${String(failures.added).padStart(6)}`);
  console.log(`    on a slide the document had      ${String(failures.own).padStart(6)}`);
  console.log(`    unplaceable                      ${String(failures.unknown).padStart(6)}`);
  console.log(`    no slide in the statement list   ${String(failures.unnamed).padStart(6)}\n`);
  console.log("  READ THE SECOND TABLE ONLY AGAINST THE FIRST. With a denominator of");
  console.log("  one, every failure lands on an added slide because there is nothing");
  console.log("  else here to land on.\n");
  console.log("  `unplaceable` is a real slide id absent from that round's final");
  console.log("  inventory — a slide swept before the deck was read, which scenarios do");
  console.log("  to their own scratch slides. Every one of them is therefore a slide the");
  console.log("  add-in added, and counting them as such only widens the gap.");
}
