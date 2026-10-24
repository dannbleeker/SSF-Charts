import ts from "typescript";

/**
 * Every place a self-test scenario reads a slide count and then does arithmetic
 * or a comparison on it, instead of handing it to `deckGrowth`.
 *
 * ── WHY THIS IS AN AST WALK AND NOT A REGEX ─────────────────────────────────
 * `test/selftest.test.ts` ("fires the guards it used to silence") swept the
 * count sites with three regexes keyed on the NAMES the first repair happened to
 * meet: `after - before`, `afterInsert - before`, `settled - before`. Three
 * scenarios named their counts `slideCountBefore`/`slideCountAfter` and
 * `countBefore`/`countAfter`/`countSwept`, compared them raw, and the guard that
 * claimed to count every site stayed green over all five comparisons. A guard
 * keyed on spellings covers those spellings — the same defect, a second time,
 * one rename over.
 *
 * So this finds the VARIABLES instead: every `const x = await slideCount()` (or
 * `settledSlideCount(...)`), and then every binary expression in the same
 * function that has `x` for an operand. What an unread count does to each of
 * those is the whole bug — `undefined >= n` is false and drops a problem
 * silently, `undefined - n` is NaN and files "grew by NaN" as a hard failure —
 * and `deckGrowth` is the one place the file turns "unread" into a skip.
 *
 * An AST rather than text so that comments and template literals, which quote
 * these expressions freely, are not mistaken for code.
 *
 * ── THE ONE ALLOWED USE ─────────────────────────────────────────────────────
 * `settledSlideCount(before + wanted)`: the sum only sizes a wait for the host to
 * catch up, and `duplicateSlot` takes its verdict from `deckGrowth` afterwards.
 * An unread `before` makes that argument NaN, which costs one extra settle and
 * decides nothing.
 */
export function rawSlideCountUses(text: string, file = "selftest.ts"): string[] {
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.ES2020, true);
  const COUNTERS = new Set(["slideCount", "settledSlideCount"]);
  const OPERATORS = new Set<ts.SyntaxKind>([
    ts.SyntaxKind.PlusToken,
    ts.SyntaxKind.MinusToken,
    ts.SyntaxKind.LessThanToken,
    ts.SyntaxKind.LessThanEqualsToken,
    ts.SyntaxKind.GreaterThanToken,
    ts.SyntaxKind.GreaterThanEqualsToken,
    ts.SyntaxKind.EqualsEqualsToken,
    ts.SyntaxKind.EqualsEqualsEqualsToken,
    ts.SyntaxKind.ExclamationEqualsToken,
    ts.SyntaxKind.ExclamationEqualsEqualsToken,
  ]);

  const readsACount = (init: ts.Expression | undefined): boolean => {
    if (!init || !ts.isAwaitExpression(init)) return false;
    const call = init.expression;
    return ts.isCallExpression(call) && ts.isIdentifier(call.expression) && COUNTERS.has(call.expression.text);
  };
  const isSizingAWait = (node: ts.Node): boolean =>
    ts.isCallExpression(node.parent) &&
    ts.isIdentifier(node.parent.expression) &&
    node.parent.expression.text === "settledSlideCount";
  const enclosingFunction = (node: ts.Node): ts.Node => {
    let at: ts.Node = node;
    while (at.parent && !ts.isFunctionLike(at)) at = at.parent;
    return at;
  };

  // Keyed by position: `countAfter - countBefore` has a count on BOTH sides, so
  // it is reached once from each declaration and is still one site.
  const found = new Map<number, string>();
  const visit = (node: ts.Node) => {
    if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name) && readsACount(node.initializer)) {
      const name = node.name.text;
      const scope = enclosingFunction(node);
      const walk = (inner: ts.Node) => {
        if (
          ts.isBinaryExpression(inner) &&
          OPERATORS.has(inner.operatorToken.kind) &&
          ((ts.isIdentifier(inner.left) && inner.left.text === name) ||
            (ts.isIdentifier(inner.right) && inner.right.text === name)) &&
          !isSizingAWait(inner)
        )
          found.set(inner.pos, inner.getText(source));
        ts.forEachChild(inner, walk);
      };
      walk(scope);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return [...found.values()];
}
