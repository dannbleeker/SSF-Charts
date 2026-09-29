import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { dirname, join, normalize } from "path";
import { sourceFiles } from "./helpers/module-source";

/**
 * NO RUNTIME IMPORT CYCLE IN `src/`.
 *
 * WHY THIS HAS TO EXIST, and why it did not before. There is no
 * `eslint-plugin-import` in this repo, so no `import/no-cycle`, and nothing else
 * looks at the shape of the module graph. A cycle does not fail `tsc` — circular
 * imports are legal TypeScript — and it does not fail the suite either, because
 * vitest loads modules in whatever order a test file asks for and usually gets
 * away with it. What it does is hand one of the two modules a binding that is
 * still `undefined` at evaluation time, in whichever order the BUNDLER picks.
 *
 * So the failure surfaces in Vite, in the pane, inside PowerPoint, as a
 * `TypeError: x is not a function` on a line that has not changed — with the
 * whole suite green. This project has a name for that class already: the bug
 * whose local reproduction cannot see it.
 *
 * AND IT FOUND ONE ON ITS FIRST RUN, which is the argument for writing it.
 *
 * Three readings of this graph by hand — two of them deliberate audits — all
 * reported it acyclic. They were answering a different question: whether
 * `src/core` imports the RENDERER, which is the architecture's stated rule and is
 * true. None looked for a cycle WITHIN `src/core`, and there are three, all
 * between `column.ts` and the combo base kinds it dispatches to. They are
 * deliberate, documented at `column.ts:38`, and listed below.
 *
 * The rest of the graph is a DAG: the `taskpane → render` edge is one-way, and
 * inside `src/render/` it is `{host-probe, experiments} → powerpoint → {lazy,
 * office-error-text, office-proxy-read, pptx-deck, ooxml}`.
 *
 * It is written NOW because a refactor that extracts siblings out of
 * `powerpoint.ts` is exactly the operation that creates a new one — pull a helper
 * out, have it reach back for something it used to sit beside, and the edge
 * closes. A guard added after the extraction would be a guard written to pass.
 *
 * ── WHAT IS DELIBERATELY NOT COUNTED ────────────────────────────────────────
 * `import type` and `import { type X }` are erased before anything runs, so a
 * type-only cycle cannot produce an undefined binding. Counting them would make
 * this fire on a shape that is harmless, which is how a guard gets switched off.
 * Dynamic `import()` is skipped for the same reason — it resolves after module
 * init, which is the whole point of `src/render/lazy.ts`.
 */

/** A module's runtime imports, as resolved repo-relative paths. */
export function runtimeImports(file: string, read: (f: string) => string = (f) => readFileSync(f, "utf8")): string[] {
  const src = read(file)
    // Comments first. A commented-out import is not an edge, and half this
    // source is prose — the same reason `module-source.ts` blanks them.
    .replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, " "))
    .split("\n")
    .map((l) => l.replace(/\/\/.*$/, ""))
    .join("\n");
  const out: string[] = [];
  // `import … from "x"` and bare `import "x"`, but not `import type … from "x"`.
  for (const m of src.matchAll(/^\s*import\s+(?!type\s)([\s\S]*?)from\s*["']([^"']+)["']/gm)) {
    const clause = m[1];
    const spec = m[2];
    if (!spec.startsWith(".")) continue;
    // `import { type A, type B } from "x"` erases completely too. Only skip when
    // EVERY named binding is type-only — one value binding keeps the edge.
    const named = clause.match(/\{([\s\S]*)\}/);
    if (named) {
      const parts = named[1]
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
      if (parts.length && parts.every((p) => p.startsWith("type "))) continue;
    }
    out.push(resolveSpec(file, spec));
  }
  return out;
}

/** `"./lazy"` from `src/render/powerpoint.ts` → `src/render/lazy.ts`. */
export function resolveSpec(fromFile: string, spec: string): string {
  const raw = normalize(join(dirname(fromFile), spec))
    .split("\\")
    .join("/");
  // Extensionless specifiers are the norm here (`moduleResolution: "bundler"`).
  // A specifier that already names a file — `pptx-paint.mjs` — is left alone and
  // simply will not match a `src/` node, which is correct: it is a leaf.
  return /\.(ts|mjs|js|json)$/.test(raw) ? raw : `${raw}.ts`;
}

/**
 * The first cycle in a graph, as the path that closes it, or null.
 *
 * Exported and pure so it can be proven against a synthetic graph. A cycle
 * detector whose only evidence is that the real graph is currently acyclic has
 * not been tested — it is indistinguishable from `return null`.
 */
export function findCycle(graph: Map<string, string[]>): string[] | null {
  const state = new Map<string, "open" | "done">();
  const stack: string[] = [];

  const walk = (node: string): string[] | null => {
    const seen = state.get(node);
    if (seen === "done") return null;
    if (seen === "open") return [...stack.slice(stack.indexOf(node)), node];
    state.set(node, "open");
    stack.push(node);
    for (const next of graph.get(node) ?? []) {
      const found = walk(next);
      if (found) return found;
    }
    stack.pop();
    state.set(node, "done");
    return null;
  };

  // Every node, not just roots: a cycle with no entry point from outside it
  // would be invisible to a roots-only walk.
  for (const node of [...graph.keys()].sort()) {
    const found = walk(node);
    if (found) return found;
  }
  return null;
}

/**
 * Edges that close a cycle and are ACCEPTED, with the reason and where the code
 * already says so.
 *
 * `src/core/layout/column.ts:38` carries the decision: *"Combo base kinds. These
 * modules import back from column (LayoutResult / horizontalChrome), but the
 * calls happen at runtime so the ESM cycle resolves."* That reasoning is sound —
 * every symbol crossing the cycle is a hoisted `function` declaration, and
 * nothing calls across it at module-init time, so neither side can observe an
 * undefined binding.
 *
 * It is recorded here rather than silently tolerated because the reasoning has a
 * FAILURE CONDITION and nothing was watching it: convert any of
 * `horizontalChrome`, `seriesLabelNodes`, `legendRow`, `layoutWaterfall`,
 * `layoutMekko`, `layoutLine` or `detailParents` to a `const` arrow function, or
 * add a top-level call across the boundary, and the cycle stops resolving — in
 * the bundle, not in `tsc` and not here.
 *
 * The other fifteen modules importing `type LayoutResult` from `column` are not
 * listed: a type import is erased before anything runs and cannot be a cycle.
 */
const ACCEPTED_CYCLE_EDGES: ReadonlyArray<readonly [string, string]> = [
  ["src/core/layout/waterfall.ts", "src/core/layout/column.ts"],
  ["src/core/layout/mekko.ts", "src/core/layout/column.ts"],
  ["src/core/layout/line.ts", "src/core/layout/column.ts"],
];

describe("the module graph of src/", () => {
  it("has no runtime import cycle beyond the accepted combo-base edges", () => {
    const files = sourceFiles("src", /\.ts$/);
    // THE GUARD ON THE GUARD. A walk that resolves to nothing reports "no
    // cycles" forever, which is the vacuous pass this repo has met in a tag-key
    // loop and a widened sweep.
    expect(files.length, "no source files found — the walk broke and this checks nothing").toBeGreaterThan(40);

    const graph = new Map<string, string[]>();
    let edges = 0;
    for (const f of files) {
      const key = f.split("\\").join("/");
      // Only edges that land on another file we actually scanned. An import of
      // `skill/scripts/pptx-paint.mjs` or a node builtin is a leaf, not an edge.
      const to = runtimeImports(key).filter((t) => files.some((g) => g.split("\\").join("/") === t));
      graph.set(key, to);
      edges += to.length;
    }
    expect(edges, "no internal edges found — the specifier parser broke").toBeGreaterThan(40);

    /**
     * THE ALLOWLIST IS CHECKED BEFORE IT IS APPLIED, because a stale entry is
     * worse than no entry: it silently permits a cycle that came back for a
     * different reason. This repo's standing complaint is a list in one file and
     * a fact in another drifting apart — so each accepted edge must still be a
     * real edge, or this fails and somebody deletes the line.
     */
    for (const [from, to] of ACCEPTED_CYCLE_EDGES) {
      expect(
        graph.get(from) ?? [],
        `${from} no longer imports ${to} — delete this entry from ACCEPTED_CYCLE_EDGES`,
      ).toContain(to);
    }
    for (const [from, to] of ACCEPTED_CYCLE_EDGES) {
      graph.set(
        from,
        (graph.get(from) ?? []).filter((t) => t !== to),
      );
    }

    const cycle = findCycle(graph);
    expect(
      cycle,
      cycle
        ? `import cycle: ${cycle.join(" → ")}\n\n` +
            `One of these modules will see an undefined binding at init, in whatever order the bundler picks. ` +
            `It will not fail here or in tsc — it fails in the pane, inside PowerPoint.`
        : "",
    ).toBeNull();
  });

  it("finds a cycle when there is one, rather than always answering null", () => {
    // Proven against synthetic graphs, because the real one passes.
    expect(
      findCycle(
        new Map([
          ["a", ["b"]],
          ["b", ["c"]],
          ["c", []],
        ]),
      ),
    ).toBeNull();
    expect(
      findCycle(
        new Map([
          ["a", ["b"]],
          ["b", ["a"]],
        ]),
      ),
    ).toEqual(["a", "b", "a"]);
    expect(
      findCycle(
        new Map([
          ["a", ["b"]],
          ["b", ["c"]],
          ["c", ["a"]],
        ]),
      ),
    ).toEqual(["a", "b", "c", "a"]);
    // Self-edge.
    expect(findCycle(new Map([["a", ["a"]]]))).toEqual(["a", "a"]);
    // A cycle nothing outside it points into — invisible to a roots-only walk.
    expect(
      findCycle(
        new Map([
          ["root", []],
          ["x", ["y"]],
          ["y", ["x"]],
        ]),
      ),
    ).not.toBeNull();
    // A diamond is not a cycle, and a detector that says it is would be useless.
    expect(
      findCycle(
        new Map([
          ["a", ["b", "c"]],
          ["b", ["d"]],
          ["c", ["d"]],
          ["d", []],
        ]),
      ),
    ).toBeNull();
  });

  it("reads value imports and skips the ones that are erased before runtime", () => {
    const fake = [
      'import { lazy } from "./lazy";',
      'import type { Scene } from "../core/scene";',
      'import { type PolygonNode, type TextNode } from "../core/scene";',
      'import { trace, type TraceKind } from "../core/trace";',
      'import JSZip from "jszip";',
      '// import { ghost } from "./ghost";',
      'const later = await import("./deferred");',
    ].join("\n");
    const got = runtimeImports("src/render/powerpoint.ts", () => fake);
    // `./lazy` and `../core/trace` are value edges. The type-only pair is erased,
    // `jszip` is not relative, the commented one is not code, and the dynamic
    // import resolves after init.
    expect(got).toEqual(["src/render/lazy.ts", "src/core/trace.ts"]);
  });

  it("resolves a specifier the way the bundler does", () => {
    expect(resolveSpec("src/render/powerpoint.ts", "./lazy")).toBe("src/render/lazy.ts");
    expect(resolveSpec("src/render/powerpoint.ts", "../core/scene")).toBe("src/core/scene.ts");
    // Already has an extension — left alone, so it simply will not match a node.
    expect(resolveSpec("src/render/pptx-deck.ts", "../../skill/scripts/pptx-paint.mjs")).toBe(
      "skill/scripts/pptx-paint.mjs",
    );
  });
});
