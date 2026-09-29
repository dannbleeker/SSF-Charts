import { readFileSync } from "fs";
import ts from "typescript";
import { sourceFiles } from "./module-source";

/**
 * How big each function in `src/` is, in CODE lines.
 *
 * ── WHY CODE LINES AND NOT LINES ────────────────────────────────────────────
 * This repo is 48% comment — 27,458 prose lines against 27,092 of code, and
 * `powerpoint.ts` is 62% prose. That is the project's memory and the thing a
 * refactor must not lose. A size ratchet measured in raw lines would therefore
 * fire every time somebody wrote down why, punishing the one habit here most
 * worth keeping, and it would be switched off within a month. The rule this
 * enforces is about COMPLEXITY; comments are not complexity.
 *
 * So both comment-only lines and blanks are dropped before counting, and a
 * function's doc block is not part of it at all — the span is the declaration
 * itself, taken from the TypeScript AST rather than by matching braces.
 *
 * ── WHY THE AST ─────────────────────────────────────────────────────────────
 * Brace counting is what I used to survey this codebase and it is wrong in two
 * ways that matter for a gate: a `{` inside a string or a regex literal shifts
 * the depth, and a declaration whose body starts on a later line is measured
 * from the wrong place. `typescript` is already a dependency and answers both
 * exactly.
 */
export interface FunctionSize {
  /** `src/taskpane/app.ts` — posix separators, so the fixture is platform-stable. */
  file: string;
  name: string;
  /** Lines inside the declaration that are neither blank nor comment-only. */
  code: number;
}

/** Comment-only and blank lines, by line number (1-based), for one source file. */
function nonCodeLines(text: string): Set<number> {
  const out = new Set<number>();
  const lines = text.split("\n");
  // Blank the block comments first, preserving newlines, then judge each line.
  const blanked = text.replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, " ")).split("\n");
  for (let i = 0; i < lines.length; i++) {
    const withoutLineComment = (blanked[i] ?? "").replace(/\/\/.*$/, "");
    if (withoutLineComment.trim() === "") out.add(i + 1);
  }
  return out;
}

/** Every named function-like declaration in a file, with its code-line count. */
export function functionSizes(file: string): FunctionSize[] {
  const text = readFileSync(file, "utf8");
  const src = ts.createSourceFile(file, text, ts.ScriptTarget.ES2020, true);
  const skip = nonCodeLines(text);
  const posix = file.split("\\").join("/");
  const found: FunctionSize[] = [];

  const nameOf = (node: ts.Node): string | undefined => {
    if (ts.isFunctionDeclaration(node) || ts.isMethodDeclaration(node)) return node.name?.getText(src);
    // `const x = () => {}` and `const x = function () {}` — the useful name is
    // the variable's, which is what a reader greps for.
    if (ts.isVariableDeclaration(node) && node.initializer) {
      if (ts.isArrowFunction(node.initializer) || ts.isFunctionExpression(node.initializer))
        return node.name.getText(src);
    }
    return undefined;
  };

  const bodyOf = (node: ts.Node): ts.Node | undefined => {
    if (ts.isFunctionDeclaration(node) || ts.isMethodDeclaration(node)) return node.body;
    if (ts.isVariableDeclaration(node) && node.initializer) return node.initializer;
    return undefined;
  };

  const visit = (node: ts.Node): void => {
    const name = nameOf(node);
    const body = bodyOf(node);
    if (name && body) {
      const from = src.getLineAndCharacterOfPosition(body.getStart(src)).line + 1;
      const to = src.getLineAndCharacterOfPosition(body.getEnd()).line + 1;
      let code = 0;
      for (let l = from; l <= to; l++) if (!skip.has(l)) code++;
      found.push({ file: posix, name, code });
    }
    ts.forEachChild(node, visit);
  };
  visit(src);
  return found;
}

/**
 * Every function in `src/` at or above `floor` code lines, worst first.
 *
 * NESTED FUNCTIONS ARE COUNTED TOO, and their lines also count toward the
 * parent — which is correct for a ratchet. A thousand-line function does not
 * become simple by holding its complexity in a closure, and the parent is the
 * thing a reader has to hold in their head.
 */
export function largeFunctions(floor = 150, dir = "src"): FunctionSize[] {
  return sourceFiles(dir, /\.ts$/)
    .flatMap(functionSizes)
    .filter((f) => f.code >= floor)
    .sort((a, b) => b.code - a.code || (a.file + a.name < b.file + b.name ? -1 : 1));
}

/** The fixture's shape: `file::name` -> recorded code lines. */
export type SizeMark = Record<string, number>;

export const keyOf = (f: FunctionSize): string => `${f.file}::${f.name}`;

export const toMark = (list: FunctionSize[]): SizeMark => Object.fromEntries(list.map((f) => [keyOf(f), f.code]));
