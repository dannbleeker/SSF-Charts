/**
 * The manifest rules this project has actually been bitten by, in one place.
 *
 * Two callers, and the second is why this file exists. `test/manifest.test.ts`
 * checks the four manifests in the working tree. `check-published-install.mjs`
 * checks the ones a USER downloads, which is a different file and has twice not
 * been the same file:
 *
 * - v0.1.0 shipped the DEV manifests while the README told users to download
 *   `manifest-prod.xml`, which was not in the release at all. Twelve days.
 * - v0.3.0 (2026-07-31) shipped `<Version>0.1.0</Version>`. Microsoft's own
 *   validator called that an error — "Manifest Version Too Low: the manifest
 *   has unsupported version number less than 1.0" — the first time CI ran it,
 *   on 2026-08-06, and #289 fixed the file in the repo the same day. No release
 *   has been cut since, so the fix has never reached a single user.
 *
 * Both are the same shape: the committed artifact is gated, the PUBLISHED one
 * is not, and the gate going green is what makes the gap invisible.
 *
 * Pure, and no Node imports on purpose — `check-published-install.mjs` fetches
 * over the network and this must stay something a browser or a test can call
 * with a string.
 */

/** The GUIDs. Changing one makes it a different add-in and orphans every sideload. */
export const ADDIN_IDS = {
  powerpoint: "b7f6d3a2-4c1e-4e8a-9f2b-7d5c0a1e6f43",
  excel: "c8a7e4b3-5d2f-4f9b-8a3c-8e6d1b2f7a54",
};

/** Which add-in a manifest filename names. */
export function familyOf(name) {
  return /excel/i.test(name) ? "excel" : "powerpoint";
}

/**
 * Everything wrong with one manifest, as sentences a reader can act on.
 *
 * An empty array means it passes the rules below — NOT that Microsoft's
 * validator would accept it. That runs in CI as a job of its own and is the
 * authority; this is the subset that must hold even when the service is
 * unreachable, which for the whole life of this repo it was.
 *
 * @param {string} xml   the manifest's text
 * @param {string} name  its filename, used for the GUID and localhost rules
 * @param {{ expectProduction?: boolean }} [opts]  production manifests may not name localhost
 */
export function checkManifest(xml, name, opts = {}) {
  const problems = [];
  if (!xml || !/<OfficeApp/i.test(xml)) {
    problems.push(`${name}: this is not an Office add-in manifest`);
    return problems;
  }

  const version = /<Version>([^<]+)<\/Version>/.exec(xml)?.[1];
  if (!version) {
    problems.push(`${name}: has no <Version>`);
  } else {
    const parts = version.split(".").map((p) => Number(p));
    if (!parts.every((p) => Number.isInteger(p) && p >= 0)) {
      problems.push(`${name}: version "${version}" is not a dotted number`);
    } else if (parts.length < 2) {
      problems.push(`${name}: version "${version}" needs at least major.minor`);
    } else if (parts[0] < 1) {
      // The rule verbatim, and the one this repo has actually shipped wrong:
      // "unsupported version number less than 1.0".
      problems.push(`${name}: version "${version}" is below 1.0, which Office rejects outright`);
    }
  }

  const id = /<Id>([^<]+)<\/Id>/.exec(xml)?.[1];
  const expected = ADDIN_IDS[familyOf(name)];
  if (id !== expected) {
    problems.push(
      `${name}: add-in GUID is ${id ?? "missing"}, expected ${expected} — every sideload would be orphaned`,
    );
  }

  const production = opts.expectProduction ?? /-prod\.xml$/.test(name);
  if (production && /localhost/i.test(xml)) {
    problems.push(`${name}: a production manifest points at localhost, so nothing in it resolves for a user`);
  }

  // THE ICONS ARE A SIZE CONTRACT, AND NOTHING HERE HAD EVER READ IT.
  //
  // Microsoft's reference for `HighResolutionIconUrl`: "For content and task
  // pane apps, the image resolution must be 64 x 64 pixels", and both it and
  // `IconUrl` are "required for add-ins distributed in AppSource". All four
  // manifests pointed the high-res one at `icon-80.png` — measured 80x80 — from
  // the day they were written until 2026-09-10, and `icon-64.png` was sitting
  // in `assets/` and on the deployed site the whole time.
  //
  // WHY EVERY EXISTING GATE MISSED IT, which is the part worth keeping:
  // `check-published-install.mjs` HEAD-checks every URL, so it proved the icon
  // RESOLVES and reported green; `office-addin-manifest validate` checks the
  // schema, not the pixels; and the rules above never opened an image. **A URL
  // that 200s is not a URL that is right.**
  //
  // Checked by NAME rather than by measuring the file: these rules run against
  // an XML string, with no filesystem and no image decoder, and `npm run icons`
  // is what guarantees `icon-64.png` is 64x64. So what is enforced here is "the
  // declaration names the asset built for it" — the half that actually drifted.
  for (const [element, px] of [
    ["IconUrl", 32],
    ["HighResolutionIconUrl", 64],
  ]) {
    const url = new RegExp(`<${element}[^>]*DefaultValue="([^"]+)"`).exec(xml)?.[1];
    if (!url) {
      problems.push(`${name}: has no <${element}>, which AppSource requires`);
    } else if (!new RegExp(`icon-${px}\\.png$`, "i").test(url)) {
      problems.push(
        `${name}: <${element}> points at ${url.split("/").pop()}, but Office requires ${px}x${px} there — ` +
          `use icon-${px}.png (a URL that resolves is not a URL that is the right size)`,
      );
    }
  }

  return problems;
}

/**
 * Every absolute URL a manifest asks the host to FETCH.
 *
 * XML namespace declarations are not fetched, so they are dropped — but by
 * HOSTNAME, not by substring. `u.includes("schemas.microsoft.com")` was the
 * first attempt and CodeQL was right to fail it (`js/incomplete-url-substring-
 * sanitization`): `https://anything.example/?x=schemas.microsoft.com` matches
 * that test, and here the consequence is a real URL quietly excused from being
 * checked at all — the sanitiser deciding what the checker gets to see.
 *
 * A URL that will not parse is KEPT. It cannot be a namespace we recognise, and
 * a manifest carrying an unparseable URL is a finding rather than a thing to
 * skip quietly.
 */
const NOT_FETCHED = new Set(["schemas.microsoft.com", "schemas.openxmlformats.org", "www.w3.org"]);

export function urlsIn(xml) {
  const found = xml.match(/https?:\/\/[^"'<>\s]+/g) ?? [];
  return [
    ...new Set(
      found.filter((u) => {
        try {
          return !NOT_FETCHED.has(new URL(u).hostname);
        } catch {
          return true;
        }
      }),
    ),
  ];
}
