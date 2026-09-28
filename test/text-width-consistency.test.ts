import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { ellipsize, textWidth } from "../src/core/scene";
import { clipToWidth } from "../src/core/elements";
import type { TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * `textWidth` MEASURES A CODE POINT, NOT A STRING IT HAPPENS TO SIT IN.
 *
 * `text-width-cjk.test.ts` next door pins what a full-width glyph costs. This
 * file pins the property underneath that one, which is the property the whole
 * engine assumes and which the function did not have: the width of a piece of
 * text is the sum of the widths of its parts, and nothing else about the string
 * changes the answer.
 *
 * Three ways it was not true, all found by the same question — "does this
 * character cost the same in every string?":
 *
 *   1. An ASTRAL code point outside the full-width set (a flag's regional
 *      indicators, a playing card, a mathematical alphanumeric) was charged two
 *      narrow units when the string held no full-width glyph and ONE when it
 *      did, because the sum was written two ways. `ellipsize` seeds a running
 *      width from `textWidth` and subtracts per dropped code point, so a clip
 *      that removed the last CJK character flipped the accounting underneath
 *      itself and returned text up to TWICE the width it was asked to fit.
 *   2. A COMBINING MARK was charged a full narrow unit though it is drawn on
 *      its base and adds no advance. `"é"` is one code point composed and two
 *      decomposed, and macOS hands out the decomposed form — so the same
 *      pasted name measured 8% wider on a Mac than on a PC, and the fits
 *      followed.
 *   3. The emoji rule started at U+1F300, so the forty-odd emoji that live in
 *      the BMP — ⭐ ✅ ❌ ⚡ ⬛ — were charged the Latin average while their
 *      astral twins 🌟 🟢 🔴 🔥 🟥 were charged a full em.
 *
 * All three are the same defect the CJK pass fixed, wearing different clothes:
 * the engine was measuring something adjacent to what it meant to measure.
 */
describe("a code point costs the same in every string", () => {
  const FS = 12;

  /** One of each interesting class, as code points so a slice never splits a pair. */
  const ALPHABET = [
    ..."abZ9 ",
    ..."売上高",
    ..."ｱｲ",
    ..."📊📈", // astral, full-width
    ..."🇩🇰", // astral, NOT full-width — a flag's two regional indicators
    ..."🃏🀄🈁", // astral, NOT full-width — cards, mahjong, enclosed ideograph
    ..."𝐀𝐁", // astral, NOT full-width — mathematical alphanumerics
    ..."⭐✅❌", // BMP emoji
    ..."éà",
    "é", // the same "é", decomposed
  ];

  it("is additive: the whole is the sum of its parts", () => {
    /**
     * The property the clip depends on, asserted directly. It failed on any
     * string mixing a full-width glyph with an astral code point that is not
     * one: `textWidth("🇩🇰")` was 25.92 and `textWidth("日")` 12, but
     * `textWidth("🇩🇰日")` came back 24.96 — less than the flag alone.
     */
    for (const a of ALPHABET)
      for (const b of ALPHABET)
        for (const bold of [false, true])
          expect(textWidth(a + b, FS, bold), `${JSON.stringify(a)} + ${JSON.stringify(b)}`).toBeCloseTo(
            textWidth(a, FS, bold) + textWidth(b, FS, bold),
            9,
          );
  });

  it("never lets a clip return text wider than the width it was given", () => {
    /**
     * The consequence, swept. `clipToWidth` and `ellipsize` are the one
     * implementation behind every shrink-then-clip in the engine and behind
     * `clipTextToFrame`, the backstop that keeps ink on the chart — so a clip
     * that overshoots is a label drawn onto the slide beside the chart.
     *
     * `clipToWidth("🇩🇰🇸🇪🇳🇴 北欧の売上", 12, 40)` came back 71.3pt wide.
     */
    const rnd = (() => {
      let s = 20260928;
      return () => (s = (s * 1103515245 + 12345) % 2147483648) / 2147483648;
    })();
    for (let i = 0; i < 20000; i++) {
      let s = "";
      const len = 1 + Math.floor(rnd() * 8);
      for (let p = 0; p < len; p++) s += ALPHABET[Math.floor(rnd() * ALPHABET.length)];
      const bold = rnd() < 0.5;
      const maxW = rnd() * 90;
      for (const out of [ellipsize(s, FS, bold, maxW), clipToWidth(s, FS, maxW, bold)]) {
        if (!out) continue;
        expect(
          textWidth(out, FS, bold),
          `${JSON.stringify(s)} clipped to ${maxW.toFixed(2)} came back ${JSON.stringify(out)}`,
        ).toBeLessThanOrEqual(maxW + 1e-9);
      }
    }
  });

  it("keeps a mixed CJK-and-emoji label inside its chart", () => {
    /**
     * End to end, because the unit property only matters through a layout. A
     * heatmap's row labels are right-aligned in a gutter fitted to them: at
     * 160x120 the clipped name was drawn 29.4 points past the LEFT edge of a
     * 160-point chart, and neither PowerPoint renderer clips a text box.
     */
    const name = "🇩🇰🇸🇪🇳🇴 北欧の売上";
    const scene = buildChart({
      kind: "heatmap",
      width: 160,
      height: 120,
      title: name,
      data: {
        categories: [name, name, name],
        series: [
          { name, values: [1, 2, 3] },
          { name: `${name}2`, values: [3, 2, 1] },
        ],
      },
    } as unknown as ChartConfig);
    for (const n of scene.nodes) {
      if (n.kind !== "text") continue;
      const t = n as TextNode;
      const w = textWidth(t.text, t.fontSize, t.bold);
      const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
      expect(x, `${t.name} starts left of the chart: ${JSON.stringify(t.text)}`).toBeGreaterThan(-0.5);
      expect(x + w, `${t.name} runs past the chart: ${JSON.stringify(t.text)}`).toBeLessThan(scene.width + 0.5);
    }
  });

  it("measures a Mac paste and a Windows paste the same", () => {
    /**
     * NFD is not an exotic input: it is what macOS gives you, so the same
     * category name typed once arrives composed from one machine and decomposed
     * from another. A combining mark charged 0.54em made those two different
     * charts — at 12pt on a 120x90 frame the title was drawn at 9.5pt composed
     * and 9.0 decomposed, a treemap drew four labels composed and two
     * decomposed, and 252 charts of a 25-kind sweep differed with no visible
     * difference in their text.
     */
    for (const s of ["é", "Årsomsætning", "Trésorerie", "Ứng dụng"])
      expect(textWidth(s.normalize("NFD"), FS), s).toBeCloseTo(textWidth(s.normalize("NFC"), FS), 9);

    const build = (form: "NFC" | "NFD") =>
      buildChart({
        kind: "treemap",
        width: 120,
        height: 90,
        title: "Årsomsætning i Europa".normalize(form),
        data: {
          categories: ["Årsomsætning", "Trésorerie", "Événement", "Português"].map((c) => c.normalize(form)),
          series: [{ name: "Ökonomie".normalize(form), values: [40, 30, 20, 10] }],
        },
      } as unknown as ChartConfig);
    const composed = build("NFC").nodes.filter((n): n is TextNode => n.kind === "text");
    const decomposed = build("NFD").nodes.filter((n): n is TextNode => n.kind === "text");
    expect(decomposed.length, "a decomposed paste drew a different number of labels").toBe(composed.length);
    composed.forEach((t, i) => {
      expect(decomposed[i].fontSize, `${t.name} was drawn at a different size`).toBe(t.fontSize);
      expect(decomposed[i].text.normalize("NFC"), `${t.name} was clipped differently`).toBe(t.text.normalize("NFC"));
    });
  });

  it("charges a BMP emoji what it charges its astral twin", () => {
    // Same glyph size on the slide, and the table that says "emoji are one em"
    // simply started above them.
    for (const [bmp, astral] of [
      ["⭐", "🌟"],
      ["✅", "🟢"],
      ["❌", "🔴"],
      ["⚡", "🔥"],
      ["⬛", "🟥"],
      ["❓", "🔷"],
    ])
      expect(textWidth(bmp, FS), `${bmp} vs ${astral}`).toBeCloseTo(textWidth(astral, FS), 9);
  });

  it("leaves the narrow symbols this engine draws itself alone", () => {
    /**
     * THE GUARD ON THE FIX ABOVE. `buildCheckbox` draws ✓ and ✗, the funnel's
     * conversion note draws ▴ and ▾, and every clip in the engine ends in an
     * ellipsis — all of them ambiguous-width, all of them rendered at a Latin
     * advance. Widening the emoji table onto those would double the measured
     * width of a glyph this product puts on charts today.
     */
    for (const s of ["✓", "✗", "▴", "▾", "…", "–", "−", "€", "⇄", "⋯", "→", "•"])
      expect(textWidth(s, FS), s).toBeCloseTo(FS * 0.54, 9);
  });

  it("still measures Latin exactly as it always did", () => {
    // The reason any of this could be changed at all — see the note in
    // `text-width-cjk.test.ts`. Sixty-odd call sites and every shipped chart's
    // fits were measured against this number.
    for (const s of ["Revenue", "Q1", "", "Ünïcödé àccents", "1,234.56", "EBITDA bridge"]) {
      expect(textWidth(s, FS), s).toBeCloseTo(s.length * FS * 0.54, 10);
      expect(textWidth(s, FS, true), `${s} bold`).toBeCloseTo(s.length * FS * 0.58, 10);
    }
  });
});
