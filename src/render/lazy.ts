/**
 * Dynamic imports that say something useful when a deploy has moved underneath
 * the pane.
 *
 * FOUND BY A ROUND, not by review. The 4:3 leg of the 2026-08-28 cycle failed
 * every scenario after the first, and the first one's message was:
 *
 *     threw: Failed to fetch dynamically imported module:
 *     https://ssf-chart.struktureretsundfornuft.dk/assets/pptxgen.es-C8DOodSg.js
 *
 * A new build had been published while the pane was open. Vite gives every
 * chunk a content hash, a deploy replaces the whole `assets/` directory, and the
 * pane was still holding an `index.html` that names chunks the server no longer
 * has. Nothing is wrong with the code, the deck or the host — the page is simply
 * older than the site.
 *
 * IT IS NOT A TEST-HARNESS PROBLEM. Every user meets this: the pane stays open
 * for the length of a PowerPoint session, releases go out during the day, and
 * the first deck insert after one lands on a chunk that has been deleted. What
 * they saw was a raw browser message about a URL, from an add-in that looked
 * broken.
 *
 * So the failure is named. `isStaleBuild` recognises the three shapes browsers
 * use for it and the caller says what happened and what to do — "close the pane
 * and open it again" — which the reader can act on. It does NOT reload the pane
 * itself: `location.reload()` in an Office task pane raises a beforeunload
 * prompt over the user's unsaved deck, and this project has already lost a
 * sideload twice to that dialog.
 *
 * WHAT IT CANNOT SAY IS WHY. `import()` rejects with ONE message whatever went
 * wrong. Chromium's "Failed to fetch dynamically imported module: <url>" covers
 * a dropped connection, a DNS or proxy failure, a 5xx and a 404 alike (its
 * module-script fetcher hands back the same null for all of them), and Firefox's
 * and Safari's wordings are no more specific. The first version of this file
 * read the wording as "the chunk is gone" and told every user "SSF Charts has
 * been updated" — including the one whose connection had dropped, or whose proxy
 * was blocking the chunk, on a site that had not changed at all. The only
 * message-free ways to tell the two apart are to re-fetch `build.json` the way
 * `warnIfStale` does or to probe the chunk, and both are a new network request
 * from a pane that makes exactly one, so the sentence names BOTH causes. The
 * advice is the same for both, and right for both: the browser keeps a failed
 * module fetch for the life of the page, so only a fresh pane can load it.
 */

/** Thrown in place of the browser's own message when a chunk would not load. */
export class StaleBuildError extends Error {
  constructor(
    readonly what: string,
    readonly cause?: unknown,
  ) {
    super(
      `Part of SSF Charts (${what}) could not be loaded — either the connection dropped or ` +
        `SSF Charts was updated since this pane was opened. Close the pane and open it again — your slides are untouched.`,
    );
    this.name = "StaleBuildError";
  }
}

/**
 * Whether this error is a module chunk that would not load — because the
 * server no longer has it, or because it could not be reached; see the note
 * above for why the wording cannot say which.
 *
 * THREE SHAPES, because the browsers do not agree and this has to recognise all
 * of them or it is worse than nothing — a check that matches only Chrome would
 * report the honest message on one browser and the raw one on the others, which
 * is the harder bug to diagnose of the two.
 *
 * Deliberately NOT matching a bare `TypeError: Failed to fetch`. That is the
 * wording of `fetch()`, never of `import()`, so it is some OTHER request failing
 * — the pane's own, or the host's — and dressing it as a chunk that would not
 * load would send whoever debugs it to the wrong place. Only the module-loading
 * wording counts.
 */
export function isStaleBuild(err: unknown): boolean {
  const msg = err instanceof Error ? err.message : String(err ?? "");
  return (
    // Chrome, Edge, and every Chromium-based host — which is what PowerPoint on
    // the web runs inside on Windows.
    /failed to fetch dynamically imported module/i.test(msg) ||
    // Firefox.
    /error loading dynamically imported module/i.test(msg) ||
    // Safari, including the iPad host.
    /importing a module script failed/i.test(msg)
  );
}

/**
 * Load a chunk, and turn "the file would not load" into a sentence.
 *
 * Anything else is rethrown untouched: a genuine bug inside the imported module
 * must not be dressed up as a load failure, or the next person debugging it
 * starts from the wrong place.
 */
export async function lazy<T>(load: () => Promise<T>, what: string): Promise<T> {
  try {
    return await load();
  } catch (err) {
    if (isStaleBuild(err)) throw new StaleBuildError(what, err);
    throw err;
  }
}
