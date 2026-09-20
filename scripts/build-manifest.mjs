#!/usr/bin/env node
/**
 * Generate the production add-in manifests from the dev manifests by swapping
 * the localhost dev-server origin for the hosted (GitHub Pages) origin. Keeps
 * the two in lockstep: the ONLY difference between dev and prod is the URL —
 * GUIDs, requirement sets and ribbon layout stay identical, so PowerPoint
 * treats them as the same add-in.
 *
 *   node scripts/build-manifest.mjs            # write manifest-*-prod.xml
 *   node scripts/build-manifest.mjs --check    # fail if the prod files are stale
 *
 * The site is served from a custom domain at its root, so the origin has no
 * path segment. Override with PAGES_ORIGIN if the host ever changes.
 */
import { readFileSync, writeFileSync } from "node:fs";

const DEV_ORIGIN = "https://localhost:3000";
const PAGES_ORIGIN = (process.env.PAGES_ORIGIN ?? "https://ssf-chart.struktureretsundfornuft.dk").replace(/\/$/, "");

const PAIRS = [
  ["manifest.xml", "manifest-prod.xml"],
  ["manifest-excel.xml", "manifest-excel-prod.xml"],
];

/**
 * The manifest the ROUND LOOP sideloads: identical to the production one except
 * that every task-pane URL carries `?harness=1`.
 *
 * WHY IT EXISTS. `app.ts` can hide the Automation ▸ Testing section unless the
 * pane is opened with `harness=1` — demo deck, self-test, host probe, download
 * run log, clean up the last round. Its own comment is the argument: "A stranger
 * cannot parse it, and two of those buttons change their document. That is fine
 * for a sideloaded tool and is a problem for a store listing, where a reviewer
 * opens every tab cold."
 *
 * THE GATE WAS BUILT AND LEFT OFF, because flipping it alone would hide the
 * section from the round driver as well and stop the loop —
 * `TESTING_UI_NEEDS_OPT_IN` says exactly that. This is the missing half: with a
 * harness manifest emitted and the driver sideloading it, flipping that constant
 * costs the loop nothing.
 *
 * NOTHING CHANGES TODAY. The constant is still `false`, so `harness=1` is read
 * and ignored. This only makes the flip available; whether to take it is the
 * owner's call, because it changes what a user receives.
 *
 * ONLY `taskpane.html` URLs. The gate lives in the task pane's bundle; icons and
 * the support URL are not pane loads and rewriting them would mean nothing.
 */
const HARNESS_FROM = "manifest-prod.xml";
const HARNESS_OUT = "manifest-harness.xml";

/** Add `harness=1` to a task-pane URL, respecting whatever query it already has. */
export function withHarness(url) {
  if (!url.includes("taskpane.html")) return url;
  if (/[?&](amp;)?harness=1(&|$)/.test(url)) return url;
  // `&amp;` because this is XML — the deep links already read
  // `?tab=elements&amp;el=harvey`, and a bare `&` would not parse.
  return url + (url.includes("?") ? "&amp;" : "?") + "harness=1";
}

/** Rewrite every task-pane `DefaultValue` in a manifest. */
export function harnessManifest(xml) {
  return xml.replace(/DefaultValue="([^"]*taskpane\.html[^"]*)"/g, (_m, url) => `DefaultValue="${withHarness(url)}"`);
}

const check = process.argv.includes("--check");
let stale = false;

for (const [dev, prod] of PAIRS) {
  const src = readFileSync(dev, "utf8");
  if (!src.includes(DEV_ORIGIN)) {
    throw new Error(`${dev} has no ${DEV_ORIGIN} origin to rewrite — did the dev manifest change shape?`);
  }
  const out = src.replaceAll(DEV_ORIGIN, PAGES_ORIGIN);
  if (out.includes(DEV_ORIGIN)) throw new Error(`${dev}: dev origin survived the swap`);
  if (check) {
    let current = null;
    try {
      current = readFileSync(prod, "utf8");
    } catch {
      /* missing → stale */
    }
    if (current !== out) {
      stale = true;
      console.error(`${prod} is stale — run \`npm run build:manifest\` and commit it`);
    } else {
      console.log(`${prod} is current`);
    }
  } else {
    writeFileSync(prod, out);
    console.log(`wrote ${prod} → ${PAGES_ORIGIN}`);
  }
}

// THE HARNESS MANIFEST, derived from the PROD one rather than from `manifest.xml`
// — so it can never drift from what ships in anything but the one parameter, and
// so it points at the deployed origin the driver already refuses to disagree with.
{
  const src = readFileSync(HARNESS_FROM, "utf8");
  const out = harnessManifest(src);
  if (out === src) {
    throw new Error(
      `${HARNESS_FROM}: no taskpane.html URL took the harness parameter — did the manifest change shape?`,
    );
  }
  if (check) {
    let current = null;
    try {
      current = readFileSync(HARNESS_OUT, "utf8");
    } catch {
      /* missing → stale */
    }
    if (current !== out) {
      stale = true;
      console.error(`${HARNESS_OUT} is stale — run \`npm run build:manifest\` and commit it`);
    } else {
      console.log(`${HARNESS_OUT} is current`);
    }
  } else {
    writeFileSync(HARNESS_OUT, out);
    console.log(`wrote ${HARNESS_OUT} → every taskpane.html URL carries ?harness=1`);
  }
}

if (check && stale) process.exit(1);
