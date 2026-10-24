import { describe, expect, it } from "vitest";
import { buildChart, normalizeConfig, DEFAULT_SIZE } from "../src/core/chart";
import { divergingScale, noDataFill, NO_DATA, sequentialScale, toHex6, toRgb, zoneFill } from "../src/core/color";
import { formatP, parseDateToken } from "../src/core/format";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { contrastInk } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Epoch day of a calendar date, the unit `parseDateToken` answers in. */
const day = (y: number, m: number, d: number) => Math.floor(Date.UTC(y, m - 1, d) / 86400000);

/**
 * r1-3. The day-existence check (#422) predates the slash ambiguity gate (#432),
 * so it was written when a slash date named no day for certain and was left out.
 * After the gate every slash date that reaches `Date.parse` names its day
 * unambiguously, and `Date.parse` rolls an impossible one into the next month:
 * `04/31/2026` became 1 May while `Apr 31 2026`, `2026-04-31` and `30.02.2026`
 * were all refused.
 */
describe("parseDateToken refuses a slash date whose day does not exist (r1-3)", () => {
  it.each([
    "04/31/2026", // US order, April has 30 days
    "02/29/2023", // not a leap year
    "09/31/2026",
    "11/31/2026",
    "02/30/26", // two-digit year
    "4/31/26",
    "2026/04/31", // year first
    "2026/02/30",
    "2026/2/29",
  ])("%s is not a date", (t) => {
    expect(parseDateToken(t), `${t} rolled into the next month`).toBeNull();
  });

  it.each([
    ["04/30/2026", day(2026, 4, 30)],
    ["02/29/2024", day(2024, 2, 29)], // a real leap day
    ["02/28/2023", day(2023, 2, 28)],
    ["12/31/2026", day(2026, 12, 31)],
    ["01/15/2026", day(2026, 1, 15)], // unambiguous US order, untouched
    ["03/03/2026", day(2026, 3, 3)], // both readings agree, untouched
    ["2026/04/30", day(2026, 4, 30)],
    ["2026/4/5", day(2026, 4, 5)],
    ["2024/02/29", day(2024, 2, 29)],
    ["04/15/2026 23:30", day(2026, 4, 15)],
    ["04/15/2026 11:30 PM", day(2026, 4, 15)],
    ["2026/04", day(2026, 4, 1)], // a year-month label names no day, so none can be absent
  ])("%s still reads as the date it names", (t, want) => {
    expect(parseDateToken(t), `${t} stopped parsing`).toBe(want);
  });

  it("is refused on the config path too, so a Gantt bar is not drawn to a day nobody wrote", () => {
    const gantt = (end: string) =>
      ({
        kind: "gantt",
        ...DEFAULT_SIZE,
        data: {
          categories: ["Design", "Build"],
          series: [
            { name: "Start", values: ["04/01/2026", "04/15/2026"] },
            { name: "End", values: ["04/10/2026", end] },
          ],
        },
      }) as unknown as ChartConfig;
    const endOf = (end: string) => normalizeConfig(gantt(end)).data.series[1].values[1];
    expect(endOf("04/30/2026")).toBe(day(2026, 4, 30));
    expect(endOf("04/31/2026"), "04/31/2026 became 1 May").toBeNull();
  });
});

/**
 * r1-36. The `#` branch checked only the LENGTH, then `parseInt` read the hex
 * PREFIX: it stops at the first non-hex character, so NaN appeared only when the
 * very first digit was bad. `#00ff0o` came back #000ff0 and `#4e7g` #44ee77, and
 * the alpha digits of a 4- or 8-digit form were never read at all, so `#abc;`
 * came back #aabbcc. The live add-in (`officeHex` -> `toHex6`) filled that
 * invented colour while the SVG sink and the pptx sink, both strict, drew black,
 * and `contrastInk` chose its ink for the invented colour.
 */
describe("a hex colour with a stray non-hex character is unreadable, not a prefix (r1-36)", () => {
  const GREY = [128, 128, 128];

  it.each([
    "#abc;", // a CSS paste with its semicolon: 4 characters, so the alpha slot was never read
    "#abcz",
    "#4e79a7zz", // 8 characters: junk where the alpha byte goes
    "#00ff0o", // a letter o for a zero, mid-string
    "#4e7g",
    "#4e79a7g1",
    "#12 456", // a space inside
    "#4e-9a7",
  ])("%s", (c) => {
    expect(toRgb(c), `${c} was read as the hex digits in front of the typo`).toEqual(GREY);
    expect(toHex6(c)).toBe("#808080");
  });

  it("still reads every well-formed length, in either case", () => {
    expect(toRgb("#ABC")).toEqual([0xaa, 0xbb, 0xcc]);
    expect(toRgb("#abcd")).toEqual([0xaa, 0xbb, 0xcc]);
    expect(toRgb("#4E79A7")).toEqual([0x4e, 0x79, 0xa7]);
    expect(toRgb("#4e79a780")).toEqual([0x4e, 0x79, 0xa7]);
    expect(toRgb("#4E79A7FF")).toEqual([0x4e, 0x79, 0xa7]);
    expect(toRgb("  #4e79a7  ")).toEqual([0x4e, 0x79, 0xa7]);
  });

  it("gives a segment label the ink a genuinely unreadable paint gets, not the ink of an invented colour", () => {
    // #aabbcc is light, so the invented colour chose DARK ink, and the preview
    // and the deck fill that segment black: near-black text on black. An
    // unreadable paint gets white ink, which stays visible on black.
    expect(contrastInk("#abc;")).toBe(contrastInk("#xyz"));
    expect(contrastInk("#abc;")).toBe("#ffffff");
  });
});

/**
 * r2-64. The hsl branch took the bare numbers and treated the first as degrees
 * whatever followed it, so `hsl(0.5turn 100% 50%)` — CSS cyan — was read as half
 * a degree and drawn red, with `contrastInk` choosing ink for the red. A unit
 * the paint grammar does not admit (the SVG allow-list names `deg` and nothing
 * else) is unreadable, the declared grey, not a confident wrong hue.
 */
describe("an hsl() hue with a unit other than deg is unreadable, not read as degrees (r2-64)", () => {
  const GREY = [128, 128, 128];

  it.each([
    "hsl(0.5turn 100% 50%)", // CSS cyan, read as 0.5 degrees: red
    "hsl(200grad 100% 50%)",
    "hsl(3.14159rad 100% 50%)",
    "hsla(0.5turn, 100%, 50%, 0.5)",
    "HSL(0.5TURN 100% 50%)",
    "hsl(1e2 50% 50%)", // exponent: the regex split it into 1 and 2 and shifted every component
    "hsl(none 100% 50%)", // no hue token at all: the saturation was read as the hue
    "hsl(from red h s l)",
  ])("%s", (c) => {
    expect(toRgb(c), `${c} was read with its unit dropped`).toEqual(GREY);
    expect(toHex6(c)).toBe("#808080");
  });

  it("still reads a unitless hue and a hue in degrees, which were never wrong", () => {
    const cyan = [0, 255, 255];
    expect(toRgb("hsl(180 100% 50%)")).toEqual(cyan);
    expect(toRgb("hsl(180deg 100% 50%)")).toEqual(cyan);
    expect(toRgb("HSL(180DEG, 100%, 50%)")).toEqual(cyan);
    expect(toRgb("hsl( 180 , 100% , 50% )")).toEqual(cyan);
    expect(toRgb("hsl(-180 100% 50%)")).toEqual(cyan); // a negative hue wraps
    expect(toRgb("hsl(540deg 100% 50%)")).toEqual(cyan);
    expect(toRgb("hsla(180 100% 50% / 0.5)")).toEqual(cyan);
    expect(toRgb("hsla(180, 100%, 50%, 0.5)")).toEqual(cyan);
  });
});

/**
 * r1-4. `relLuminance(style.background)` decides light versus dark canvas for
 * every canvas-aware helper, and `toRgb` did not know `transparent` (only
 * `alphaOf` did), so it answered the UNREADABLE mid grey, luminance 0.216, which
 * is under the 0.5 cut: the chart was laid out for a DARK canvas that the SVG
 * and the slide draw as the light one. The heatmap's cell lightness range
 * collapsed from 4.01:1 to 1.10:1, so the scale encoded value only through
 * saturation, and "no data" became #919191, louder than the data. Commit 7b7f872
 * fixed the same flip for `background: "white"` by teaching `toRgb` the CSS
 * names; `transparent` is not one.
 */
describe("a canvas that paints nothing, or cannot be read, lays out as the light default (r1-4)", () => {
  const white = "#ffffff";
  // Written as a user, an agent, or a web tool's computed style would write them.
  const CLEAR = [
    "transparent",
    "TRANSPARENT",
    "  transparent  ",
    "rgba(0,0,0,0)",
    "#00000000",
    "#0000",
    "hsla(0,0%,0%,0)",
  ];
  // Words and shapes that name no colour at all: `render-pptx.mjs` cites "off-white".
  const UNREADABLE = ["none", "off-white", "not-a-color", "", "#abc;", "hsl(0.5turn 50% 50%)"];

  it("reads the transparent keyword as the slide behind it, which every sink takes to be white", () => {
    expect(toRgb("transparent")).toEqual([255, 255, 255]);
    expect(toRgb("TRANSPARENT")).toEqual([255, 255, 255]);
    expect(toHex6("transparent")).toBe(white);
    // Ink for a label that sits on a fill that is not there goes against the slide.
    expect(contrastInk("transparent")).toBe("#0b0b0b");
  });

  it.each([...CLEAR, ...UNREADABLE])("%j draws the light canvas's no-data grey and zone tint", (bg) => {
    expect(noDataFill(bg)).toBe(NO_DATA);
    expect(zoneFill(bg, "#f5f5f5")).toBe("#f5f5f5");
    expect(zoneFill(bg, "#efe7e7")).toBe("#efe7e7");
  });

  it.each([...CLEAR, ...UNREADABLE])("%j gives the sequential and diverging scales a light empty end", (bg) => {
    const seq = sequentialScale(0, 10, "#2a78d6", bg);
    const seqWhite = sequentialScale(0, 10, "#2a78d6", white);
    for (const v of [0, 2.5, 5, 10]) expect(seq(v), `sequential at ${v}`).toBe(seqWhite(v));
    const div = divergingScale(-10, 10, "#2a78d6", "#e34948", bg);
    const divWhite = divergingScale(-10, 10, "#2a78d6", "#e34948", white);
    for (const v of [-10, -3, 0, 3, 10]) expect(div(v), `diverging at ${v}`).toBe(divWhite(v));
  });

  it("still lays a genuinely dark canvas out as dark", () => {
    // The other half, or "treat it as light" could be satisfied by never asking.
    expect(noDataFill("#1b1b1b")).not.toBe(NO_DATA);
    expect(zoneFill("#1b1b1b", "#f5f5f5")).not.toBe("#f5f5f5");
    expect(noDataFill("rgba(0,0,0,0.5)"), "a half-transparent dark canvas is still a dark one").not.toBe(NO_DATA);
  });

  it("draws every chart kind the same on background transparent as on #ffffff", () => {
    // The whole class, not the two layouts the report named: boxplot, violin,
    // waffle, treemap, sunburst, the line fill-between ribbon and the column mark
    // paint all tint toward `style.background`, directly or through a helper.
    // Only a paint that IS the background literal may differ, and only in its
    // spelling.
    const withBackground = (kind: ChartConfig["kind"], background: string) => {
      const base = sampleConfig(kind);
      return buildChart({ ...base, style: { ...(base.style ?? {}), background } } as ChartConfig);
    };
    const differing: string[] = [];
    let compared = 0;
    for (const { kind } of CHART_KINDS) {
      const clear = withBackground(kind, "transparent");
      const solid = withBackground(kind, white);
      expect(clear.nodes.length, `${kind}: node count`).toBe(solid.nodes.length);
      clear.nodes.forEach((a, i) => {
        const b = solid.nodes[i] as unknown as Record<string, unknown>;
        for (const p of ["fill", "stroke", "color"]) {
          const av = (a as unknown as Record<string, unknown>)[p];
          if (av === b[p] || (av === "transparent" && b[p] === white)) continue;
          differing.push(`${kind}/${a.name ?? i}.${p}: ${String(av)} vs ${String(b[p])}`);
        }
        compared++;
      });
    }
    expect(compared, "the sweep compared nothing").toBeGreaterThan(1000);
    expect(differing.slice(0, 12), `${differing.length} paints differ`).toEqual([]);
  });
});

/**
 * r1-38. `formatP` took no locale and used `toFixed`, and both scatter trend
 * captions build `R² = …` with a bare `toFixed(2)`, so a scatter chart with
 * `numberFormat.locale: "de-DE"` printed comma-decimal axes (`2,0`) beside a
 * dot-decimal caption (`R² = 0.60, p = 0.12`): the "two number systems in one
 * chart" that `formatPercent` and the percent-label sweep exist to prevent.
 * The thresholds `< 0.001`, `< 0.01` and `< 0.05` were literals and kept the dot
 * in every locale too.
 *
 * This half is `formatP`. The caption itself is built in `layout/scatter.ts`
 * (`R²` at the two `trend-stats` call sites), which must pass the chart's
 * `numberFormat.locale` to it and format `R²` with `formatNumber`.
 */
describe("formatP prints the p-value in the chart's locale (r1-38)", () => {
  it.each(["de-DE", "fr-FR", "da-DK"])("%s uses a decimal comma, in the value and in the thresholds", (locale) => {
    expect(formatP(0.31, locale), locale).toBe("= 0,31");
    expect(formatP(0.0004, locale), locale).toBe("< 0,001");
    expect(formatP(0.004, locale), locale).toBe("< 0,01");
    expect(formatP(0.03, locale), locale).toBe("< 0,05");
  });

  it("keeps the dot with no locale, with en-US, and with a malformed tag", () => {
    for (const locale of [undefined, "en-US", "not a locale"]) {
      expect(formatP(0.31, locale), String(locale)).toBe("= 0.31");
      expect(formatP(0.0004, locale), String(locale)).toBe("< 0.001");
      expect(formatP(0.004, locale), String(locale)).toBe("< 0.01");
      expect(formatP(0.03, locale), String(locale)).toBe("< 0.05");
    }
  });

  it("rounds to two decimals and prints a whole p-value in full", () => {
    expect(formatP(0.05)).toBe("= 0.05");
    expect(formatP(0.123456)).toBe("= 0.12");
    expect(formatP(0.999)).toBe("= 1.00");
    expect(formatP(1)).toBe("= 1.00");
  });
});

/**
 * r1-38, END TO END. `formatP` taking a locale proves nothing about the chart:
 * the caption is composed in `layout/scatter.ts`, at two call sites, and until
 * both pass the chart's locale a de-DE scatter printed comma axes beside
 * `R² = 0.31, p = 0.07`. This drives `buildChart` the way a user's config does
 * and reads the text node the caption lands in.
 */
describe("the scatter trend caption is printed in the chart's locale (r1-38)", () => {
  // Scattered enough that the fit is weak and p lands in the "= 0.xx" branch.
  const ys = [5, 2, 8, 1, 9, 3, 7, 4];
  const cloud = (locale: string | undefined, degree?: number): ChartConfig =>
    ({
      kind: "scatter",
      ...DEFAULT_SIZE,
      ...(locale ? { numberFormat: { locale } } : {}),
      ...(degree ? { scatter: { trendDegree: degree } } : {}),
      data: {
        categories: ys.map((_, i) => `p${i}`),
        series: [
          { name: "X", values: ys.map((_, i) => i + 1) },
          { name: "Y", values: ys },
          { name: "Trend", values: ys.map(() => 1) },
        ],
      },
    }) as unknown as ChartConfig;
  const caption = (c: ChartConfig): string => {
    const n = buildChart(c).nodes.find((x) => x.name === "trend-stats");
    expect(n, "no trend caption was drawn").toBeTruthy();
    return (n as unknown as { text: string }).text;
  };

  it("uses the decimal comma for the straight fit's R² and p", () => {
    const text = caption(cloud("de-DE"));
    expect(text, text).toMatch(/^R² = \d,\d\d, p [<=] \d,\d+$/);
    expect(text, text).not.toMatch(/\d\.\d/);
  });

  it("uses the decimal comma for a polynomial fit's R²", () => {
    const text = caption(cloud("de-DE", 2));
    expect(text, text).toMatch(/^R² = \d,\d\d · quadratic$/);
  });

  it("keeps the dot with no locale", () => {
    expect(caption(cloud(undefined))).toMatch(/^R² = \d\.\d\d, p [<=] \d\.\d+$/);
    expect(caption(cloud(undefined, 2))).toMatch(/^R² = \d\.\d\d · quadratic$/);
  });
});
