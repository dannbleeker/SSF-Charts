import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "fs";
import { posix, sep } from "path";

/**
 * The pane's test mock defines everything the pane imports from the renderer.
 *
 * TWICE IN ONE NIGHT this cost 110 tests at once. `app.ts` gained a call to a
 * new renderer export — `readDeckStyleWithReason`, then `warmCustomXmlSurface` —
 * and `test/pane-host-actions.test.ts` mocks that whole module, so the export
 * was simply absent. `app.ts` throws while it is being imported, every test in
 * the file dies, and the useful sentence
 *
 *     [vitest] No "warmCustomXmlSurface" export is defined on the
 *     "../src/render/powerpoint" mock
 *
 * is buried under a hundred and ten failures that all look like something else.
 * The information was there both times; the signal-to-noise was not.
 *
 * This turns that into one failure that names the missing export.
 *
 * It compares the IMPORT LIST rather than what is actually called, on purpose: a
 * value that is imported and not yet called today will be called eventually, and
 * a guard that waits for the call is a guard that fires on the day someone is
 * busy. Types are excluded — they vanish at runtime and a mock never needs them.
 */

/**
 * The value names a module imports from one specific module path — from EVERY
 * statement that names it, not the first. A second `import { … } from` the same
 * path (a split import is how a long list gets edited) used to be invisible, and
 * so did the shape of a type-only `import type { … }`, whose start the search for
 * `import {` walked straight past into the previous statement.
 */
export function importedValuesFrom(source: string, modulePath: string): string[] {
  const marker = `} from "${modulePath}";`;
  const names: string[] = [];
  for (let end = source.indexOf(marker); end >= 0; end = source.indexOf(marker, end + marker.length)) {
    const open = source.lastIndexOf("import {", end);
    // `import type { … }` is erased before anything runs, whole.
    if (open < 0 || source.lastIndexOf("import type {", end) > open) continue;
    const body = source.slice(open + "import {".length, end);
    names.push(
      ...body
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean)
        // `type X` and `type { X }` are erased before anything runs.
        .filter((s) => !s.startsWith("type "))
        // `x as y` — the mock has to define the exported name, not the local alias.
        .map((s) => s.split(/\s+as\s+/)[0].trim())
        .filter((s) => /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(s)),
    );
  }
  return names;
}

/** The keys a `vi.mock(path, () => ({ ... }))` factory defines. */
export function mockedKeys(source: string, modulePath: string): string[] {
  const at = source.indexOf(`vi.mock("${modulePath}"`);
  if (at < 0) return [];
  const body = source.slice(at);
  const keys = new Set<string>();
  // Property keys at any depth are fine: a name defined anywhere in the factory
  // is a name the module will expose. Over-matching here can only make this
  // guard quieter, never noisier, and the unit tests below pin that it still
  // catches an absent one.
  for (const m of body.matchAll(/(^|[\s{,])([A-Za-z_$][A-Za-z0-9_$]*)\s*:/g)) keys.add(m[2]);
  return [...keys];
}

/**
 * EVERY MODULE THE PANE LOADS THAT IMPORTS THE RENDERER, found rather than listed.
 *
 * 2026-09-29 this widened from `app.ts` alone to `app.ts` + `harness-ui.ts`,
 * because `vi.mock` replaces the module for EVERY importer and the testing panel
 * had just taken about thirty imports with it. It named two files again, and the
 * same thing had already happened one level down: `harness-ui.ts` imports
 * `selftest.ts`, `host-probe.ts` and `experiments.ts`, which load under the same
 * factory and which the guard never read. Between them they need 22 names the
 * factory does not define and 11 it defines for them alone (the sync counters,
 * the late-sync probe, `reconcileDeck`, the probe context), and some of those
 * eleven are read OUTSIDE a scenario's try: delete `lastLateSyncSeq` from the
 * mock and 11 pane tests die with `Cannot read properties of undefined (reading
 * 'text')`, the missing export named nowhere, which is the whole-suite failure
 * this file exists to turn into one sentence. The guard stayed green.
 *
 * So the corpus is every file under `src/` that imports a value from the
 * renderer, by the path it would spell it with. A module that starts importing
 * it is covered the day it does, and the list cannot narrow by someone forgetting
 * to extend it. What matters is the UNION: whichever file imports a name, the
 * single mock has to define it.
 */
const RENDERER = "src/render/powerpoint";

/** Every file under `src/` with the renderer values it imports, derived from the tree. */
function rendererImporters(): { file: string; names: string[] }[] {
  const found: { file: string; names: string[] }[] = [];
  for (const entry of readdirSync("src", { recursive: true }).map(String)) {
    if (!entry.endsWith(".ts")) continue;
    const file = `src/${entry.split(sep).join("/")}`;
    const rel = posix.relative(posix.dirname(file), RENDERER);
    const names = importedValuesFrom(readFileSync(file, "utf8"), rel.startsWith(".") ? rel : `./${rel}`);
    if (names.length) found.push({ file, names });
  }
  return found;
}

/**
 * The two modules the pane evaluates AT BOOT. A name either imports and the
 * factory omits is fatal to every test in `pane-host-actions.test.ts` at once, so
 * nothing is excused for them.
 */
const LOADED_AT_BOOT = ["src/taskpane/app.ts", "src/taskpane/harness-ui.ts"];

/**
 * RENDERER NAMES THE MOCK DOES NOT DEFINE AND THE PANE SUITE SURVIVES, 2026-10-24.
 *
 * Only for modules the pane reaches THROUGH `harness-ui.ts`: the self-test
 * battery and the host probe. Each scenario runs inside its own try, so a name
 * the factory omits arrives as vitest's `No "x" export is defined on the mock`
 * and becomes one failed verdict — the pane suite's one-click round does exactly
 * that 19 times and still passes, because it asserts that verdicts were bundled
 * into the file and not what they say. The battery itself is tested against the
 * real renderer and the fake host in `selftest.test.ts` and `host-probe.test.ts`,
 * which is where these names are exercised.
 *
 * It is a LIST, and closed, rather than 22 `vi.fn()` stubs: a stub changes what
 * the battery does inside the pane suite, for a battery whose real coverage is
 * elsewhere. Anything not listed is held to the mock like everything else, and
 * `staleTolerations` fails the day a listed name is mocked or no longer
 * imported — delete it then. Adding a name here is a decision that it is only
 * ever read inside a scenario; the ones that are not (the sync counters,
 * `lastLateSync`, `waitForLateSync`, `reconcileDeck`) are in the mock and must
 * stay.
 */
const TOLERATED_ABSENT: Record<string, string[]> = {
  "src/taskpane/selftest.ts": [
    "canSelectShapes",
    "clearShapeSelection",
    "stepOf",
    "readbackTimeoutMs",
    "selectShape",
    "selectionLadder",
    "awaitSelectedChart",
    "canWatchSelection",
    "slideImageBase64",
    "slideOccupancy",
    "showSlide",
    "settledSlideCount",
    "slideShapeList",
    "shapeGeometryByName",
    "timeShapeRounds",
    "moveShapeBy",
    "lastStall",
    "RASTERISE_OP",
    "rasterGap",
    "shapesDrawnOn",
  ],
  "src/render/host-probe.ts": ["scratchSlideIdByPosition", "deleteTrailingSlides"],
};

type Importer = { file: string; names: string[] };

/** What each importer needs that the mock lacks and nothing excuses, as `file: name` lines. */
function unaccountedFor(importers: Importer[], mocked: Set<string>): string[] {
  return importers.flatMap(({ file, names }) =>
    names
      .filter((n) => !mocked.has(n))
      .filter((n) => LOADED_AT_BOOT.includes(file) || !TOLERATED_ABSENT[file]?.includes(n))
      .map((n) => `${file}: ${n}`),
  );
}

/** Tolerations that no longer describe anything: the name is mocked now, or the file stopped importing it. */
function staleTolerations(importers: Importer[], mocked: Set<string>): string[] {
  return Object.entries(TOLERATED_ABSENT).flatMap(([file, names]) =>
    names
      .filter((n) => mocked.has(n) || !importers.find((i) => i.file === file)?.names.includes(n))
      .map((n) => `${file}: ${n}`),
  );
}

describe("the pane mock covers what the pane imports", () => {
  const paneTest = readFileSync("test/pane-host-actions.test.ts", "utf8");
  const importers = rendererImporters();
  const mocked = new Set(mockedKeys(paneTest, "../src/render/powerpoint"));
  const importedByThePane = (): string[] => [...new Set(importers.flatMap((i) => i.names))];

  it("parses both sides, or it is comparing nothing with nothing", () => {
    // The failure mode of a guard like this is matching zero on both sides and
    // passing forever. The pane imports dozens from the renderer; if this ever
    // reads near zero, the parser broke rather than the import list shrinking.
    expect(importedByThePane().length).toBeGreaterThan(30);
    expect(mocked.size).toBeGreaterThan(30);
    // AND EACH BOOT MODULE IS SEEN. A union that silently came from one file would
    // pass the floor above while covering half of what it claims.
    for (const file of LOADED_AT_BOOT)
      expect(
        importers.find((i) => i.file === file)?.names.length ?? 0,
        `${file} contributed no renderer imports — it moved, or the parser stopped seeing it`,
      ).toBeGreaterThan(0);
    // AND SO IS EVERYTHING BEHIND THEM. The scan has to reach past the two named
    // files, or it is the narrower guard this replaced: some importer must be
    // one the pane loads only through another module.
    expect(
      importers.filter((i) => !LOADED_AT_BOOT.includes(i.file)).length,
      "no module besides the two loaded at boot imports the renderer — the scan stopped reaching them",
    ).toBeGreaterThan(0);
  });

  it("defines every renderer value the pane imports", () => {
    const missing = unaccountedFor(importers, mocked);
    expect(
      missing,
      `test/pane-host-actions.test.ts mocks the renderer but does not define: ${missing.join(", ")}\n` +
        "Add it to the vi.mock factory, or the pane throws at import and every test in that file dies at once.",
    ).toEqual([]);
  });

  it("excuses only absences that are still true", () => {
    expect(staleTolerations(importers, mocked), "delete these from TOLERATED_ABSENT").toEqual([]);
  });

  it("notices each name the mock keeps by hand for the modules behind harness-ui going missing", () => {
    // THE CASE THE TWO-FILE GUARD COULD NOT SEE, run for every name it concerns.
    // These are renderer names an importer OTHER than the boot modules needs, which
    // the factory defines — the sync counters and the late-sync probe among them.
    // Take each out of the mock in turn (the key set, in memory) and the guard
    // has to name it. Derived, not listed: the next name the battery needs is
    // covered the day it is mocked.
    const keptByHand = [
      ...new Set(importers.filter((i) => !LOADED_AT_BOOT.includes(i.file)).flatMap((i) => i.names)),
    ].filter((n) => mocked.has(n) && !importers.some((i) => LOADED_AT_BOOT.includes(i.file) && i.names.includes(n)));
    expect(keptByHand.length, "nothing behind harness-ui needs a name only the mock supplies").toBeGreaterThan(0);
    for (const name of keptByHand) {
      const without = new Set(mocked);
      without.delete(name);
      expect(
        unaccountedFor(importers, without).map((line) => line.split(": ")[1]),
        `the guard did not notice "${name}" missing from the mock`,
      ).toContain(name);
    }
  });

  it("catches an absent export rather than shrugging", () => {
    // Proven against a synthetic pair, because the real pair passes — a guard
    // whose only evidence is that it is currently quiet has not been tested.
    const fakeApp = 'import {\n  alpha,\n  beta,\n  type Gamma,\n  delta as d,\n} from "../render/powerpoint";';
    expect(importedValuesFrom(fakeApp, "../render/powerpoint")).toEqual(["alpha", "beta", "delta"]);

    const fakeTest = 'vi.mock("../src/render/powerpoint", () => ({\n  alpha: vi.fn(),\n  beta: 1,\n}));';
    const keys = new Set(mockedKeys(fakeTest, "../src/render/powerpoint"));
    expect(
      ["alpha", "beta", "delta"].filter((n) => !keys.has(n)),
      "missed the absent one",
    ).toEqual(["delta"]);
  });

  it("reads every import statement that names the module, and none that is only a type", () => {
    // A second statement from the same path, and a type-only one between two
    // value ones: the first form was invisible and the second made the search
    // for `import {` land on the previous statement.
    const src =
      'import { alpha } from "../render/powerpoint";\n' +
      'import type { Shape } from "../render/powerpoint";\n' +
      'import { other } from "./elsewhere";\n' +
      'import { beta, gamma } from "../render/powerpoint";\n';
    expect(importedValuesFrom(src, "../render/powerpoint")).toEqual(["alpha", "beta", "gamma"]);
  });

  it("holds a module behind harness-ui to the mock unless it is listed", () => {
    // The scope fix itself, on a synthetic importer: a name the battery needs and
    // the factory lacks is reported for `selftest.ts`, a listed one is not, and
    // the boot modules are excused of nothing.
    const mocked = new Set(["alpha"]);
    const behind = (names: string[]) => [{ file: "src/taskpane/selftest.ts", names }];
    expect(unaccountedFor(behind(["alpha", "brandNewName"]), mocked)).toEqual([
      "src/taskpane/selftest.ts: brandNewName",
    ]);
    expect(unaccountedFor(behind(["alpha", "shapesDrawnOn"]), mocked)).toEqual([]);
    expect(unaccountedFor([{ file: "src/taskpane/harness-ui.ts", names: ["shapesDrawnOn"] }], mocked)).toEqual([
      "src/taskpane/harness-ui.ts: shapesDrawnOn",
    ]);
    // …and a toleration for something the mock now defines is reported as stale.
    expect(staleTolerations(behind(["shapesDrawnOn"]), new Set(["shapesDrawnOn"]))).toContain(
      "src/taskpane/selftest.ts: shapesDrawnOn",
    );
  });
});
