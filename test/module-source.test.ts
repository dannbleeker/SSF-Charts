import { describe, expect, it } from "vitest";
import { mkdtempSync, mkdirSync, writeFileSync, rmSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import {
  codeOf,
  describeOffenders,
  expectSweptSomething,
  sourceDeclaring,
  sourceFiles,
  sweep,
} from "./helpers/module-source";

/**
 * The helper that widens a source sweep from one file to a directory.
 *
 * PROVEN AGAINST PLANTED VIOLATIONS, not against the tree being currently
 * clean. This repo's flagship rule is that a regression test must be shown to
 * fail without its fix, and a sweep helper whose only evidence is that the
 * codebase happens to satisfy it today has not been tested at all — it is the
 * exact shape of the vacuous pass it exists to prevent.
 *
 * So every case below builds a tiny tree on disk with the violation IN it, and
 * asserts the sweep finds that violation at that line.
 */

/** A throwaway source tree. Returns its root; the caller removes it. */
function tree(files: Record<string, string>): string {
  const root = mkdtempSync(join(tmpdir(), "ssf-sweep-"));
  for (const [rel, body] of Object.entries(files)) {
    const full = join(root, rel);
    mkdirSync(join(full, ".."), { recursive: true });
    writeFileSync(full, body, "utf8");
  }
  return root;
}

describe("sourceFiles", () => {
  it("walks subdirectories and returns a stable order", () => {
    const root = tree({
      "b.ts": "",
      "a.ts": "",
      "nested/deep/c.ts": "",
      "skip.md": "",
      "nested/d.html": "",
    });
    try {
      const rel = sourceFiles(root).map((p) =>
        p
          .slice(root.length + 1)
          .split("\\")
          .join("/"),
      );
      // Sorted, and `.md` excluded. The order matters: an unsorted readdirSync
      // reports offenders differently on CI than locally, which reads as flake.
      expect(rel).toEqual(["a.ts", "b.ts", "nested/d.html", "nested/deep/c.ts"]);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });
});

describe("codeOf", () => {
  it("blanks block, HTML and line comments", () => {
    const root = tree({
      "x.ts": ["const a = 1; // trailing banned()", "/* banned() in a block */", "const b = 2;"].join("\n"),
      "y.html": "<!-- banned() in markup -->\n<div></div>",
    });
    try {
      const ts = codeOf(join(root, "x.ts"));
      expect(ts, "a line comment still carried its text").not.toMatch(/banned\(\)/);
      expect(ts, "the code beside the comment was lost").toContain("const a = 1;");
      expect(codeOf(join(root, "y.html")), "an HTML comment still carried its text").not.toMatch(/banned\(\)/);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it("keeps line numbers, which is the whole reason it blanks rather than deletes", () => {
    // A multi-line block comment must leave the same number of lines behind, or
    // every offender after it is reported at the wrong place — and a sweep that
    // names the wrong line is one nobody trusts twice.
    const root = tree({ "x.ts": ["/* one", "   two", "   three */", "const offender = 1;"].join("\n") });
    try {
      const lines = codeOf(join(root, "x.ts")).split("\n");
      expect(lines).toHaveLength(4);
      expect(lines[3]).toContain("const offender = 1;");
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });
});

describe("sweep", () => {
  it("finds a planted violation and names its file and line", () => {
    const root = tree({
      "a.ts": "const ok = 1;\n",
      "sub/b.ts": "const ok = 1;\nvoid thing().catch(() => {});\n",
    });
    try {
      const found = sweep(/\.catch\(\s*\(\s*\w*\s*\)\s*=>\s*\{\s*\}\s*\)/, root);
      expect(found).toHaveLength(1);
      expect(found[0].line, "reported the wrong line").toBe(2);
      expect(found[0].file).toContain("b.ts");
      expect(describeOffenders(found)).toContain(":2");
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it("does not match the same pattern inside a comment", () => {
    // The mistake this repo has made twice: a sweep matching its own prose. A
    // colour check flagged a comment about a deleted colour, and a slide-id
    // sweep matched the word in a comment beside the code it was meant to pin.
    const root = tree({ "a.ts": "// void thing().catch(() => {});\nconst ok = 1;\n" });
    try {
      expect(sweep(/\.catch\(\s*\(\s*\w*\s*\)\s*=>\s*\{\s*\}\s*\)/, root)).toEqual([]);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it("finds EVERY occurrence, not every other one", () => {
    // A caller passing a /g regex would otherwise carry `lastIndex` between
    // lines and skip alternate matches — a bug that only shows on the second
    // offender, which is exactly when a sweep matters.
    const root = tree({ "a.ts": "bad()\nbad()\nbad()\n" });
    try {
      expect(sweep(/bad\(\)/g, root)).toHaveLength(3);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });
});

describe("expectSweptSomething", () => {
  it("fails when the walk found nothing, rather than letting a ban pass over air", () => {
    const root = tree({ "notes.md": "no source here" });
    try {
      // THE GUARD ON THE GUARD. Without this, widening a sweep to a path that
      // does not exist any more turns every ban built on it green forever.
      expect(() => expectSweptSomething(root, 1)).toThrow();
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it("passes over the real tree, and the floor is well under it", () => {
    expect(expectSweptSomething("src", 40).length).toBeGreaterThan(40);
  });
});

describe("sourceDeclaring", () => {
  it("returns the one file that declares the name", () => {
    const root = tree({ "a.ts": "const other = 1;\n", "sub/b.ts": "export function target(): void {}\n" });
    try {
      expect(sourceDeclaring(/export function target\(/, root)).toContain("export function target");
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it("refuses zero matches — a body slice over nothing is not a guard", () => {
    const root = tree({ "a.ts": "const other = 1;\n" });
    try {
      expect(() => sourceDeclaring(/export function gone\(/, root)).toThrow();
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it("refuses two matches rather than picking one arbitrarily", () => {
    const root = tree({ "a.ts": "export function dup(): void {}\n", "b.ts": "export function dup(): void {}\n" });
    try {
      expect(() => sourceDeclaring(/export function dup\(/, root)).toThrow();
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });
});
