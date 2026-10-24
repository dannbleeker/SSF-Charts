import { beforeAll, describe, expect, it } from "vitest";
import { applyNodeInPlace } from "../src/render/powerpoint";
import type { NodeChange } from "../src/core/scene-diff";
import type { RectNode, TextNode } from "../src/core/scene";

/**
 * WHAT AN IN-PLACE WRITE ACTUALLY SENDS, counted.
 *
 * `applyNodeInPlace` wrote every property of every changed node
 * unconditionally. The host's own statement list puts one text node at roughly
 * twenty statements, so a retitle — the single edit the fast path exists for —
 * sent twenty to change one string, and `same scale across the deck` sent some
 * three hundred and sixty to move eighteen nodes.
 *
 * `planSceneUpdate` now says which property groups differ and the applier
 * writes only those. **Nothing else in the suite can see that.** Correctness
 * tests cannot: writing MORE than necessary is still correct, just slow, so a
 * mutant that ignores `parts` and writes everything leaves every other test
 * green. That mutant is exactly the regression this file exists to catch — the
 * optimisation being quietly reverted.
 *
 * So this counts. The shape is a recording proxy rather than the office-host
 * fake, because the fake stores values and cannot tell "wrote the same number"
 * from "did not write".
 */

/**
 * The enums only, and they matter more than they look.
 *
 * Six of the applier's writes sit inside `try` blocks that read
 * `PowerPoint.TextVerticalAlignment` and friends. With no such global the enum
 * READ throws, the catch swallows it, and those six writes never happen — so
 * the first version of this file counted thirteen writes for a whole text node
 * and silently tested a path with the margins and alignment missing.
 *
 * With the enums present it counts twenty, which is the figure `powerpoint.ts`
 * quotes from the host's own statement list. That agreement is the check that
 * this file is counting the right thing.
 */
beforeAll(() => {
  (globalThis as Record<string, unknown>).PowerPoint = {
    ShapeAutoSize: { autoSizeNone: "autoSizeNone" },
    TextVerticalAlignment: { top: "top", middle: "middle", bottom: "bottom" },
    ParagraphHorizontalAlignment: { left: "left", center: "center", right: "right" },
  };
});

/** Every property set and method called, as dotted paths. */
function recorder() {
  const seen: string[] = [];
  const make = (path: string): unknown =>
    new Proxy(function () {} as unknown as Record<string, unknown>, {
      get(_t, prop) {
        // Symbols reach here whenever something coerces the proxy — `toPrimitive`
        // from a template string, `Symbol.iterator` from a spread. Returning
        // another proxy for those makes the coercion throw, so they answer
        // undefined and only named properties are followed.
        if (typeof prop === "symbol") return undefined;
        return make(path ? `${path}.${prop}` : prop);
      },
      set(_t, prop: string, _v) {
        seen.push(path ? `${path}.${prop}` : prop);
        return true;
      },
      apply(_t, _this, _args) {
        seen.push(`${path}()`);
        return make(path);
      },
    });
  return { shape: make("") as never, seen };
}

const rect = (over: Partial<RectNode> = {}): RectNode => ({
  kind: "rect",
  x: 1,
  y: 2,
  w: 10,
  h: 20,
  fill: "#123456",
  name: "seg-0-0",
  ...over,
});
const text = (over: Partial<TextNode> = {}): TextNode => ({
  kind: "text",
  x: 1,
  y: 2,
  w: 40,
  h: 12,
  text: "hello",
  fontSize: 10,
  color: "#000000",
  align: "center",
  valign: "middle",
  name: "title",
  ...over,
});

const wrote = (n: RectNode | TextNode, parts?: NodeChange[]) => {
  const { shape, seen } = recorder();
  applyNodeInPlace(shape, n, 0, 0, {}, parts);
  return seen;
};

describe("an in-place write sends only the properties that changed", () => {
  it("sends the whole node when no diff is supplied — twenty statements for text", () => {
    // The old behaviour, still the right answer for a caller without a plan,
    // and the baseline every saving below is measured against. Twenty is the
    // figure `powerpoint.ts` quotes from the host's own statement list, which
    // is how this file knows it is counting the same thing the host charges for.
    expect(wrote(text()), "a whole text node is no longer twenty statements").toHaveLength(20);
    // Eight, not seven, since the applier started writing the fill's alpha even
    // when it is 0: a bar recoloured from translucent to opaque would otherwise
    // stay translucent. See the next two cases.
    expect(wrote(rect()), "a whole rect is no longer eight statements").toHaveLength(8);
  });

  /**
   * A shape that is being RECOLOURED has a past, and the adder's shortcuts are
   * only right for a new one. `setSolidColor` is documented as setting the colour
   * and says nothing of an alpha already there, so the transparency is written
   * even when it is 0 — on the in-place route only.
   */
  it("writes the fill's transparency on a recolour, and the outline's too", () => {
    expect(wrote(rect(), ["fill"])).toContain("fill.transparency");
    const outlined = wrote(rect({ stroke: "#2a78d6", strokeWidth: 1.5 }), ["line"]);
    expect(outlined).toContain("lineFormat.transparency");
  });

  /**
   * An outline gained in place has to be switched ON, and BEFORE it is coloured:
   * the adder never says so because a new rectangle starts with its outline
   * showing, but a column drawn as an actual has `visible = false`.
   */
  it("switches an outline on before it colours it", () => {
    const seen = wrote(rect({ stroke: "#2a78d6", strokeWidth: 1.5 }), ["line"]);
    expect(seen).toContain("lineFormat.visible");
    expect(seen.indexOf("lineFormat.visible"), "the outline is coloured before it is shown").toBeLessThan(
      seen.indexOf("lineFormat.color"),
    );
    // …and losing the stroke still hides it, once.
    expect(wrote(rect({ stroke: undefined }), ["line"]).filter((s) => s === "lineFormat.visible")).toHaveLength(1);
  });

  /** The edit the fast path exists for. Twenty statements to change one string. */
  it("sends TWO statements for a retitle, not twenty", () => {
    const seen = wrote(text(), ["text"]);
    expect(seen).toContain("textFrame.textRange.text");
    expect(seen).toContain("name");
    expect(seen, `a retitle sent ${seen.length} statements: ${seen.join(", ")}`).toHaveLength(2);
  });

  it("sends no geometry, font or alignment for a retitle", () => {
    const seen = wrote(text(), ["text"]).join(" ");
    for (const banned of ["left", "top", "width", "height", "font", "wordWrap", "paragraphFormat", "fill"])
      expect(seen, `a retitle wrote ${banned}`).not.toContain(banned);
  });

  it("sends only geometry when a bar is rescaled", () => {
    const seen = wrote(rect(), ["box"]);
    expect(seen.sort()).toEqual(["height", "left", "name", "top", "width"]);
  });

  it("sends only the fill when a series is recoloured", () => {
    const seen = wrote(rect(), ["fill"]);
    expect(seen).toContain("name");
    expect(seen.join(" ")).toContain("fill");
    expect(seen.join(" "), "a recolour moved the bar").not.toContain("left");
  });

  it("sends only the font when a label is resized", () => {
    const seen = wrote(text(), ["font"]);
    const joined = seen.join(" ");
    expect(joined).toContain("font.size");
    expect(joined).toContain("font.color");
    expect(joined, "a font change rewrote the string").not.toContain("textRange.text");
    expect(joined, "a font change moved the box").not.toContain("left");
  });

  /**
   * The frame constants ride with `box`, and this is the assertion that says
   * so. Nothing in a scene can change `wordWrap` or the margins, so a group of
   * their own would never fire — but they are what make a text box's geometry
   * mean what the layout intended, so a node whose geometry moves rewrites
   * them and a pure retitle does not.
   */
  it("rewrites the frame constants when the geometry moves, and not otherwise", () => {
    expect(wrote(text(), ["box"]).join(" ")).toContain("wordWrap");
    expect(wrote(text(), ["text"]).join(" "), "a retitle rewrote the frame constants").not.toContain("wordWrap");
  });

  /**
   * THE COUNT IS THE POINT. A rescale moves eighteen of twenty-four nodes, and
   * this is the per-node arithmetic behind the whole change.
   */
  it("costs less per node than writing the node whole", () => {
    const whole = wrote(text()).length;
    for (const [why, parts] of [
      ["retitle", ["text"]],
      ["moved label", ["box", "text"]],
      ["restyled label", ["font"]],
    ] as [string, NodeChange[]][]) {
      const some = wrote(text(), parts).length;
      expect(some, `${why} sent ${some} of ${whole} — no saving`).toBeLessThan(whole);
    }
  });
});
