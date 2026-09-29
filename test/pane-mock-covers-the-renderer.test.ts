import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";

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

/** The value names a module imports from one specific module path. */
export function importedValuesFrom(source: string, modulePath: string): string[] {
  const marker = `} from "${modulePath}";`;
  const end = source.indexOf(marker);
  if (end < 0) return [];
  const open = source.lastIndexOf("import {", end);
  if (open < 0) return [];
  const body = source.slice(open + "import {".length, end);
  return (
    body
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean)
      // `type X` and `type { X }` are erased before anything runs.
      .filter((s) => !s.startsWith("type "))
      // `x as y` — the mock has to define the exported name, not the local alias.
      .map((s) => s.split(/\s+as\s+/)[0].trim())
      .filter((s) => /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(s))
  );
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
 * BOTH PANE MODULES, 2026-09-29.
 *
 * This read `app.ts` alone. When the testing panel moved to `harness-ui.ts` it
 * took about thirty renderer imports with it — and `vi.mock` replaces the module
 * for EVERY importer, so a name only `harness-ui.ts` uses is just as fatal if
 * the factory omits it. The guard would have kept passing over the half it could
 * still see: the same narrowing this suite has just been swept for.
 *
 * Read as one corpus rather than two checks, because what matters is the UNION —
 * whichever file imports a name, the single mock has to define it.
 */
const PANE_MODULES = ["src/taskpane/app.ts", "src/taskpane/harness-ui.ts"];

describe("the pane mock covers what the pane imports", () => {
  const panes = PANE_MODULES.map((f) => readFileSync(f, "utf8"));
  const paneTest = readFileSync("test/pane-host-actions.test.ts", "utf8");
  const importedByThePane = (): string[] => [
    ...new Set(panes.flatMap((src) => importedValuesFrom(src, "../render/powerpoint"))),
  ];

  it("parses both sides, or it is comparing nothing with nothing", () => {
    // The failure mode of a guard like this is matching zero on both sides and
    // passing forever. The pane imports dozens from the renderer; if this ever
    // reads near zero, the parser broke rather than the import list shrinking.
    expect(importedByThePane().length).toBeGreaterThan(30);
    expect(mockedKeys(paneTest, "../src/render/powerpoint").length).toBeGreaterThan(30);
    // AND EACH FILE IS SEEN. A union that silently came from one file would pass
    // the floor above while covering half of what it claims.
    for (const [i, src] of panes.entries())
      expect(
        importedValuesFrom(src, "../render/powerpoint").length,
        `${PANE_MODULES[i]} contributed no renderer imports — it moved, or the parser stopped seeing it`,
      ).toBeGreaterThan(0);
  });

  it("defines every renderer value the pane imports", () => {
    const imported = importedByThePane();
    const mocked = new Set(mockedKeys(paneTest, "../src/render/powerpoint"));
    const missing = imported.filter((n) => !mocked.has(n));
    expect(
      missing,
      `test/pane-host-actions.test.ts mocks the renderer but does not define: ${missing.join(", ")}\n` +
        "Add it to the vi.mock factory, or the pane throws at import and every test in that file dies at once.",
    ).toEqual([]);
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
});
