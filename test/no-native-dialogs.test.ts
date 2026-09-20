import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

/**
 * THE PANE MUST NOT OPEN A NATIVE MODAL.
 *
 * It runs inside an Office iframe, where `alert`, `confirm` and `prompt` block
 * the HOST as well as the pane — PowerPoint itself stops responding until the
 * user dismisses a box they may not even be able to see. Microsoft's add-in
 * guidance says not to, and a store reviewer meeting a frozen PowerPoint is not
 * going to file a nuanced bug report.
 *
 * `app.ts` already made this argument in `offerOwnSlide` — "A PROMISE ROUND A
 * PAIR OF BUTTONS rather than `confirm()`" — and shipped a `prompt()` anyway,
 * on the Save-as-template button, for as long as templates have existed. The
 * reasoning was written down and one call site was missed, which is a shape
 * this repo keeps meeting.
 *
 * So the rule gets a test rather than a comment. A comment is an argument; this
 * is the thing that notices.
 */

/** Every shipped source file — `src/`, excluding nothing, because nothing there is exempt. */
function shippedFiles(dir = "src"): string[] {
  const out: string[] = [];
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, e.name);
    if (e.isDirectory()) out.push(...shippedFiles(p));
    else if (/\.(ts|html)$/.test(e.name)) out.push(p);
  }
  return out;
}

describe("the pane opens no native dialogs", () => {
  it("calls neither prompt, alert nor confirm anywhere in src/", () => {
    const offenders: string[] = [];
    for (const file of shippedFiles()) {
      // BLOCK COMMENTS GO FIRST, and HTML ones are why. The markup carries a
      // `<!-- … -->` explaining why `prompt()` was removed, and the first
      // version of this test reported those three prose lines as violations —
      // a checker that cannot tell a mention from a call, which is the trap
      // this repo has already met in a colour check that flagged a comment
      // about a deleted colour. Newlines are preserved so line numbers survive.
      const blanked = readFileSync(file, "utf8")
        .replace(/<!--[\s\S]*?-->/g, (m) => m.replace(/[^\n]/g, " "))
        .replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, " "));
      blanked.split("\n").forEach((line, i) => {
        // Then line comments, for the `//` form.
        const code = line.replace(/\/\/.*$/, "");
        // `window.prompt(` or a bare `prompt(` — but not `.prompt(` on some
        // other object, and not an identifier that merely ends in those letters.
        if (/(^|[^.\w])(window\s*\.\s*)?(prompt|alert|confirm)\s*\(/.test(code)) {
          offenders.push(`${file}:${i + 1}  ${code.trim().slice(0, 100)}`);
        }
      });
    }
    expect(offenders, `native dialog call(s) in shipped source:\n${offenders.join("\n")}`).toEqual([]);
  });

  it("still finds one if somebody adds it back", () => {
    // The guard on the guard. A matcher this narrow is easy to file down until
    // it matches nothing, and a rule that cannot fail is not a rule.
    const sample = [
      'const name = prompt("Template name?", state.title);',
      "window.confirm('are you sure');",
      "  alert(msg);",
    ];
    for (const line of sample) {
      expect(/(^|[^.\w])(window\s*\.\s*)?(prompt|alert|confirm)\s*\(/.test(line), line).toBe(true);
    }
    // And things that must NOT match: a method on something else, a name that
    // merely contains the word, and prose about the rule.
    for (const line of [
      "await page.prompt('x');",
      "const confirmed = await confirmTheThing();",
      "const alerts = readAlerts();",
      " * rather than `confirm()`. The pane runs in an Office iframe",
    ]) {
      const code = line.replace(/\/\/.*$/, "").replace(/^\s*\*.*$/, "");
      expect(/(^|[^.\w])(window\s*\.\s*)?(prompt|alert|confirm)\s*\(/.test(code), line).toBe(false);
    }
  });

  it("offers the inline replacement the Save button needs", () => {
    // The markup has to exist for `askTemplateName` to find, and it has to be
    // reachable: an input with no accessible name is a different rejection.
    const html = readFileSync("src/taskpane/taskpane.html", "utf8");
    expect(html).toContain('id="template-name-row"');
    expect(html).toContain('id="template-name-ok"');
    expect(html).toContain('id="template-name-cancel"');
    expect(html, "the name box has no accessible name").toMatch(
      /id="template-name"[^>]*aria-label="Template name"|aria-label="Template name"[^>]*id="template-name"/,
    );
    // Hidden until asked for, like `slow-offer` above it.
    expect(html).toMatch(/id="template-name-row"[^>]*hidden/);
  });
});
