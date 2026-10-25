import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart, normalizeConfig } from "../src/core/chart";
import { axisTickLabel, formatNumber, formatP, formatPercent, parseDateToken } from "../src/core/format";
import { EUROPE_TILES, US_TILES, WORLD_TILES, resolveTileLayout } from "../src/core/layout/tilemap-layouts";
import { sampleConfig } from "../src/core/samples";
import { ellipsize, textWidth, type SceneNode, type TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * Third bug hunt, stream k3: the dates, numbers and locale spellings the engine
 * reads, the ellipsis it appends, and the grid a tile map draws.
 */

const DAY_MS = 86400000;
const iso = (d: number | null) => (d == null ? null : new Date(d * DAY_MS).toISOString().slice(0, 10));
const read = (t: string, locale?: string) => iso(parseDateToken(t, locale));
const texts = (nodes: SceneNode[]) => nodes.filter((n): n is TextNode => n.kind === "text");

/**
 * r3-20 / r3-21. `Date.parse` has ONE order for a numeric date, month first, and
 * only the slash form was ever fenced off from it. Every other spelling of the same
 * digits was guessed (`03-01-2026` was 1 March, and so was `03.01.2026 08:00`
 * while `03.01.2026` was 3 January), and every spelling whose first number is above
 * 12 was refused although it has exactly one reading.
 */
describe("one rule for every numeric day-month-year spelling (r3-20, r3-21)", () => {
  it("takes the only reading there is, whichever way round it is written", () => {
    // 15 is not a month, so nothing is being guessed.
    for (const t of [
      "15-01-2026",
      "15/01/2026",
      "15.01.2026",
      "15 01 2026",
      "15. 1. 2026",
      "15-01-2026 08:00",
      "15/01/2026 08:00:30",
      "15.01.2026 08:00",
      "Mon 15/01/2026",
      "01-15-2026",
      "01/15/2026",
      "1 15 2026",
      "01-15-2026 3:00 PM",
      "01/15/2026 15:00",
    ]) {
      expect(read(t), `${t} should be 15 January`).toBe("2026-01-15");
    }
    expect(read("31/12/2025"), "31/12/2025").toBe("2025-12-31");
    expect(read("31-12-2025"), "31-12-2025").toBe("2025-12-31");
    expect(read("25.12.2026 23:59"), "25.12.2026 23:59").toBe("2026-12-25");
  });

  it("refuses a date that reads two ways, in every spelling and with or without a time", () => {
    // 03-01-2026 is 3 January or 1 March; Danish Excel's default short date.
    for (const t of [
      "03-01-2026",
      "3-1-2026",
      "03/01/2026",
      "03 01 2026",
      "03-01-26",
      "12-11-2026",
      "03-01-2026 08:00",
      "03/01/2026 08:00",
      "03/01/2026, 08:00",
      "Mon 03/01/2026",
    ]) {
      expect(read(t), `${t} was guessed at instead of refused`).toBeNull();
    }
  });

  it("keeps the dotted form day-first, bare or with a time of day or a space", () => {
    // The dot is the European separator: there is no mm.dd.yyyy, so one reading.
    expect(read("03.01.2026")).toBe("2026-01-03");
    expect(read("01.03.2026")).toBe("2026-03-01");
    expect(read("03.01.2026 08:00"), "was 1 March").toBe("2026-01-03");
    expect(read("03.01.2026, 08:00"), "was 1 March").toBe("2026-01-03");
    expect(read("3. 1. 2026"), "was 1 March").toBe("2026-01-03");
    // …in every locale, a month-first one included.
    expect(read("03.01.2026 08:00", "en-US")).toBe("2026-01-03");
    // A dot with a short year is the version-number shape, as it is bare.
    expect(read("03.01.26 08:00")).toBeNull();
    expect(read("1.2.26")).toBeNull();
  });

  it("reads a two-way date day-first only when the chart's locale writes the day first", () => {
    for (const locale of ["da-DK", "de-DE", "en-GB", "nl-NL", "fr-FR", "en-IN", "da_DK"]) {
      expect(read("03-01-2026", locale), `03-01-2026 in ${locale}`).toBe("2026-01-03");
      expect(read("05/02/2026", locale), `05/02/2026 in ${locale}`).toBe("2026-02-05");
      expect(read("03 01 2026 08:00", locale), `03 01 2026 08:00 in ${locale}`).toBe("2026-01-03");
    }
    // A month-first locale does not settle it either: only a day-first one was offered.
    // Year-first locales, an unknown tag and a malformed one pick no winner.
    for (const locale of ["en-US", "ja-JP", "sv-SE", "hu-HU", "xx-YY", "not a locale!!", "", undefined]) {
      expect(read("03-01-2026", locale), `03-01-2026 in ${String(locale)}`).toBeNull();
    }
  });

  it("does not let a locale move a date that has one reading, or that reads the same either way", () => {
    for (const locale of [undefined, "da-DK", "en-US", "ja-JP"]) {
      expect(read("01-15-2026", locale)).toBe("2026-01-15");
      expect(read("15-01-2026", locale)).toBe("2026-01-15");
      expect(read("03-03-2026", locale)).toBe("2026-03-03");
      expect(read("12/12/2026", locale)).toBe("2026-12-12");
      expect(read("1/1/2026", locale)).toBe("2026-01-01");
    }
  });

  it("gives a day that does not exist no reading at all, in any spelling", () => {
    for (const t of [
      "31/04/2026",
      "04/31/2026",
      "30-02-2026",
      "02-30-2026",
      "29/02/2023",
      "02/29/2023",
      "31 06 2026",
    ]) {
      expect(read(t), t).toBeNull();
      expect(read(t, "da-DK"), `${t} in da-DK`).toBeNull();
    }
    expect(read("29/02/2024"), "a leap day").toBe("2024-02-29");
    expect(read("02/29/2024"), "a leap day").toBe("2024-02-29");
  });

  it("reads a two-digit year as it always did, and a day-first one only on the locale's word", () => {
    // A short year leaves a third reading, year first (`26/01/03`), that a day-first-only token cannot rule out.
    expect(read("01/15/26")).toBe("2026-01-15");
    expect(read("03/03/26")).toBe("2026-03-03");
    expect(read("1/1/26")).toBe("2026-01-01");
    expect(read("1/15/75")).toBe("1975-01-15");
    expect(read("1/15/45")).toBe("2045-01-15");
    expect(read("15/01/26"), "15/01/26 could be 2015-01-26").toBeNull();
    expect(read("15/01/26", "da-DK")).toBe("2026-01-15");
    expect(read("03-01-26", "da-DK")).toBe("2026-01-03");
  });

  it("refuses the shape inside a token it cannot place, rather than handing it to Date.parse", () => {
    expect(read("03/01/2026 Mon")).toBeNull();
    expect(read("03/01/2026 hello")).toBeNull();
    // …and a month-word date or an ISO one is not that shape.
    expect(read("Mon, 05 Jan 2026")).toBe("2026-01-05");
    expect(read("2026-01-15 10:30:00")).toBe("2026-01-15");
  });

  it("is read from a config's Gantt rows with the chart's own locale", () => {
    const gantt = (numberFormat?: ChartConfig["numberFormat"]): ChartConfig => ({
      kind: "gantt",
      ...DEFAULT_SIZE,
      numberFormat,
      data: {
        categories: ["Design", "Build", "Test"],
        series: [
          { name: "Start", values: ["03-01-2026", "05-02-2026", "02-03-2026"] as unknown as number[] },
          { name: "End", values: ["09-01-2026", "11-02-2026", "06-03-2026"] as unknown as number[] },
        ],
      },
    });
    const day = (s: string) => Date.parse(`${s}T00:00:00Z`) / DAY_MS;
    const dated = normalizeConfig(gantt({ decimals: "auto", locale: "da-DK" }));
    expect(dated.data.series[0].values).toEqual([day("2026-01-03"), day("2026-02-05"), day("2026-03-02")]);
    expect(dated.data.series[1].values).toEqual([day("2026-01-09"), day("2026-02-11"), day("2026-03-06")]);
    expect(dated.data.dates).toBe(true);
    // The whole chart is the one the same dates written as ISO draw: three bars, 3-9 Jan, 5-11 Feb, 2-6 Mar.
    const isoGantt: ChartConfig = {
      ...gantt({ decimals: "auto", locale: "da-DK" }),
      data: {
        categories: ["Design", "Build", "Test"],
        series: [
          { name: "Start", values: ["2026-01-03", "2026-02-05", "2026-03-02"] as unknown as number[] },
          { name: "End", values: ["2026-01-09", "2026-02-11", "2026-03-06"] as unknown as number[] },
        ],
      },
    };
    const bars = (c: ChartConfig) => buildChart(c).nodes.filter((n) => n.kind === "rect" && n.name?.startsWith("bar"));
    expect(bars(isoGantt).length, "the fixture draws no bars").toBeGreaterThanOrEqual(3);
    expect(buildChart(gantt({ decimals: "auto", locale: "da-DK" })).nodes).toEqual(buildChart(isoGantt).nodes);
    // No locale, no winner: a visible gap, never the bars "1 Mar–1 Sep" it used to draw.
    const gap = normalizeConfig(gantt());
    expect(gap.data.series[0].values).toEqual([null, null, null]);
    expect(gap.data.series[1].values).toEqual([null, null, null]);
    expect(bars(gantt())).toEqual([]);
  });
});

/**
 * r3-40. `iso` and the day-existence check both wanted a two-digit month, so a
 * one-digit one fell to `Date.parse`, which rolls an impossible day into the next
 * month, with nothing to compare it against.
 */
describe("a nonexistent day is no date in any spelling (r3-40)", () => {
  it.each(["2026-2-30", "2023-2-29", "2023-4-31", "2023-6-31", "2023-9-31", "2023-11-31", "2026-2-30 10:00"])(
    "refuses %s",
    (t) => {
      expect(read(t), `${t} rolled into the next month`).toBeNull();
    },
  );

  it("refuses a year-first date with a space the same way", () => {
    expect(read("2026 02 30")).toBeNull();
    expect(read("2026 2 30")).toBeNull();
    expect(read("2026/2/30")).toBeNull();
    expect(read("2026 2 3")).toBe("2026-02-03");
  });

  it("still reads a one-digit month and day that exist", () => {
    expect(read("2026-2-3")).toBe("2026-02-03");
    expect(read("2024-2-29")).toBe("2024-02-29");
    expect(read("2026-1-5")).toBe("2026-01-05");
    expect(read("2026-1-15T08:00:00Z")).toBe("2026-01-15");
    expect(read("2026-2")).toBe("2026-02-01");
    expect(read("2025-12")).toBe("2025-12-01");
  });

  it("agrees with the calendar for every spelling, month and day 29-32 of a leap and a common year", () => {
    const spellings: [string, (y: number, m: number, d: number) => string][] = [
      ["y-m-d", (y, m, d) => `${y}-${m}-${d}`],
      ["y-mm-dd", (y, m, d) => `${y}-${String(m).padStart(2, "0")}-${d}`],
      ["y/m/d", (y, m, d) => `${y}/${m}/${d}`],
      ["y m d", (y, m, d) => `${y} ${m} ${d}`],
      ["d.m.y", (y, m, d) => `${d}.${m}.${y}`],
      ["d-m-y", (y, m, d) => `${d}-${m}-${y}`],
      ["d/m/y", (y, m, d) => `${d}/${m}/${y}`],
      ["d m y", (y, m, d) => `${d} ${m} ${y}`],
      ["d. m. y", (y, m, d) => `${d}. ${m}. ${y}`],
      ["m/d/y", (y, m, d) => `${m}/${d}/${y}`],
      ["m-d-y", (y, m, d) => `${m}-${d}-${y}`],
      ["d-m-y time", (y, m, d) => `${d}-${m}-${y} 08:00`],
    ];
    const wrong: string[] = [];
    for (const year of [2023, 2024]) {
      for (let m = 1; m <= 12; m++) {
        for (const d of [29, 30, 31, 32]) {
          // Day 13 and up is no month, so every spelling above has ONE reading.
          const exists = d <= new Date(Date.UTC(year, m, 0)).getUTCDate();
          const want = exists ? Math.floor(Date.UTC(year, m - 1, d) / DAY_MS) : null;
          for (const [name, spell] of spellings) {
            const t = spell(year, m, d);
            const got = parseDateToken(t);
            if (got !== want) wrong.push(`${name} ${t}: ${iso(got)} (want ${iso(want)})`);
          }
        }
      }
    }
    expect(wrong, `${wrong.length} spellings disagree with the calendar`).toEqual([]);
  });
});

/**
 * r3-61. The axis wrote `formatNumber(t * 100)` and a literal "%", so the digits
 * followed the chart's locale and the sign did not.
 */
describe("a share axis prints its percent the way the segments do (r3-61)", () => {
  const hundred = (locale?: string): ChartConfig => ({
    kind: "stacked100",
    ...DEFAULT_SIZE,
    numberFormat: { decimals: 0, locale },
    data: {
      categories: ["Q1"],
      series: ["A", "B", "C", "D"].map((name) => ({ name, values: [1] })),
    },
    decorations: { segmentLabels: true, valueAxis: true },
  });

  it.each(["de-DE", "fr-FR", "da-DK", "es-ES", "tr-TR", "ar-EG", "fa-IR", "en-US", undefined])(
    "%s: the 25% tick reads like the 25% segment",
    (locale) => {
      const nodes = texts(buildChart(hundred(locale)).nodes);
      const segment = nodes.find((n) => n.name === "label-0-0")!.text;
      const ticks = nodes.filter((n) => n.name === "value-axis").map((n) => n.text);
      expect(segment, "the fixture draws no segment label").toBe(formatPercent(0.25, 0, false, locale));
      expect(ticks, `${String(locale)}: axis ${ticks.join(" | ")}`).toContain(segment);
      expect(ticks).toContain(formatPercent(1, 0, false, locale));
    },
  );

  it("is exactly the old text in en-US, with the suffix dropped and the sign kept", () => {
    const ticks = [0, 0.25, 0.5, 0.75, 1];
    expect(ticks.map(axisTickLabel(ticks, true, {}))).toEqual(["0%", "25%", "50%", "75%", "100%"]);
    expect(ticks.map(axisTickLabel(ticks, true, { suffix: " m" }))).toEqual(["0%", "25%", "50%", "75%", "100%"]);
    expect(ticks.map(axisTickLabel(ticks, true, { decimals: 1 }))).toEqual([
      "0.0%",
      "25.0%",
      "50.0%",
      "75.0%",
      "100.0%",
    ]);
    expect(ticks.map(axisTickLabel(ticks, true, { forceSign: true }))).toEqual(["0%", "+25%", "+50%", "+75%", "+100%"]);
    expect(axisTickLabel([-0.5, 0, 0.5], true, {})(-0.5)).toBe("-50%");
    // A value axis is untouched.
    expect(axisTickLabel([0, 50, 100], false, { locale: "de-DE", suffix: " m" })(50)).toBe("50 m");
  });
});

/**
 * r3-63. A hand- or agent-written config spells a locale `de_DE` and a count "2",
 * and both were silently ignored.
 */
describe("numberFormat spellings are read as what they spell (r3-63)", () => {
  const clustered = (numberFormat: unknown): ChartConfig => ({
    kind: "clustered",
    ...DEFAULT_SIZE,
    numberFormat: numberFormat as ChartConfig["numberFormat"],
    data: { categories: ["A", "B"], series: [{ name: "S", values: [1234.567, 2345.678] }] },
    decorations: { segmentLabels: true },
  });
  const labels = (numberFormat: unknown) =>
    texts(buildChart(clustered(numberFormat)).nodes)
      .filter((n) => n.name?.startsWith("label-"))
      .map((n) => n.text);

  it('reads decimals: "2" as 2', () => {
    expect(labels({ decimals: 2 })).toEqual(["1,234.57", "2,345.68"]);
    expect(labels({ decimals: "2" }), "a numeric string printed no decimals").toEqual(["1,234.57", "2,345.68"]);
    expect(labels({ decimals: " 1 " })).toEqual(["1,234.6", "2,345.7"]);
    expect(normalizeConfig(clustered({ decimals: "2" })).numberFormat).toEqual({ decimals: 2 });
  });

  it('leaves "auto" and everything that is no number exactly as written', () => {
    expect(normalizeConfig(clustered({ decimals: "auto" })).numberFormat).toEqual({ decimals: "auto" });
    expect(normalizeConfig(clustered({ decimals: "abc" })).numberFormat).toEqual({ decimals: "abc" });
    expect(normalizeConfig(clustered({ decimals: "" })).numberFormat).toEqual({ decimals: "" });
    expect(normalizeConfig(clustered({ decimals: 1, locale: "de-DE" })).numberFormat).toEqual({
      decimals: 1,
      locale: "de-DE",
    });
    expect(normalizeConfig(clustered(undefined)).numberFormat).toBeUndefined();
    expect(normalizeConfig(clustered(null)).numberFormat).toBeNull();
  });

  it("reads the POSIX spelling de_DE as de-DE, in every formatter", () => {
    expect(formatNumber(1234.567, { decimals: 1, locale: "de_DE" })).toBe(
      formatNumber(1234.567, { decimals: 1, locale: "de-DE" }),
    );
    expect(formatNumber(1234.567, { decimals: 1, locale: "de_DE" })).toBe("1.234,6");
    expect(formatPercent(0.25, 0, false, "de_DE")).toBe(formatPercent(0.25, 0, false, "de-DE"));
    expect(formatP(0.31, "da_DK")).toBe("= 0,31");
    expect(labels({ decimals: 1, locale: "de_DE" })).toEqual(["1.234,6", "2.345,7"]);
  });

  it("still falls back to en-US for a tag Intl rejects", () => {
    expect(formatNumber(1234.567, { decimals: 1, locale: "not-a-real-locale!!" })).toBe("1,234.6");
    expect(formatNumber(1234.567, { decimals: 1, locale: "de__DE__" })).toBe("1,234.6");
  });
});

/**
 * r3-65. `ellipsize` walks code points and stops the moment the rest fits, and a
 * ZWJ sequence and a flag each have a piece that fits on its own.
 */
describe("ellipsize never stops inside a cluster it cannot cut (r3-65)", () => {
  // Written as escapes: both are invisible, and an invisible character in a source line is a bug nobody can see.
  const ZWJ = "\u{200d}";
  const RLM = "\u{200f}";
  /** Why a text cut at an ellipsis ends in half a cluster, or "" when it does not. */
  const halfCluster = (text: string): string => {
    const body = text.endsWith(RLM) ? text.slice(0, -1) : text;
    if (!body.endsWith("…")) return "";
    const kept = body.slice(0, -1);
    if (kept.endsWith(ZWJ)) return "ends in a joiner";
    // A run of regional indicators pairs from its start, so an odd run ends in half a flag.
    const run = [...(/[\u{1f1e6}-\u{1f1ff}]+$/u.exec(kept)?.[0] ?? "")].length;
    return run % 2 ? `${run} trailing regional indicators` : "";
  };
  const SOURCES = [
    "🇩🇰🇸🇪🇳🇴🇫🇮 Nordics",
    "Team 👨‍👩‍👧 growth",
    "Team 👨‍👩‍👧‍👦 growth plan for 2026",
    "Ops 🏳️‍🌈 and 👩🏽‍💻 devs",
    "ABC🇩🇰🇸🇪 trailing text",
  ];

  it("leaves no dangling joiner and no half flag in front of the ellipsis, at any width", () => {
    const bad: string[] = [];
    for (const src of SOURCES) {
      for (const bold of [false, true]) {
        for (let maxW = 1; maxW <= textWidth(src, 12, bold) + 4; maxW++) {
          const out = ellipsize(src, 12, bold, maxW);
          if (!out) continue;
          const why = halfCluster(out);
          if (why) bad.push(`${src} @${maxW}: ${why}`);
          const kept = out.replace(new RegExp(`${RLM}?$`, "u"), "").replace(/…$/u, "");
          if (!src.startsWith(kept)) bad.push(`${src} @${maxW}: ${kept} is not a prefix`);
          if (textWidth(out, 12, bold) > maxW) bad.push(`${src} @${maxW}: ${textWidth(out, 12, bold)} wide`);
        }
      }
    }
    expect(bad.slice(0, 8), `${bad.length} cuts through a cluster`).toEqual([]);
  });

  it("answers the cases that were cut", () => {
    const flags = ellipsize("🇩🇰🇸🇪🇳🇴 Nordics", 12, false, 30);
    expect(flags, "was a lone regional indicator").not.toBe("🇩…");
    expect(halfCluster(flags)).toBe("");
    const family = ellipsize("Team 👨‍👩‍👧 growth", 12, false, 80);
    expect(family.endsWith(`${ZWJ}…`), "ended in a zero-width joiner").toBe(false);
    expect(family.startsWith("Team ")).toBe(true);
  });

  it("changes nothing for text that fits, or that has no cluster to cut", () => {
    expect(ellipsize("Revenue", 12, false, 500)).toBe("Revenue");
    expect(ellipsize("Revenue by region", 12, false, 60)).toBe("Revenue …");
    expect(ellipsize("🇩🇰🇸🇪", 12, false, 500)).toBe("🇩🇰🇸🇪");
    // A skin tone is a base emoji and a modifier, each with a width: the walk may stop between them and the result is still whole.
    expect(ellipsize("a👍🏽b👍🏽c👍🏽d", 12, false, 40)).toBe("a👍🏽…");
  });

  it("puts a right-to-left mark after the ellipsis of a right-to-left label, and only then", () => {
    const arabic = "المبيعات حسب المنطقة والشهر";
    const hebrew = "מכירות לפי אזור וחודש";
    for (const src of [arabic, hebrew]) {
      const out = ellipsize(src, 12, false, 70);
      expect(out.endsWith(`…${RLM}`), `${src} -> ${out}`).toBe(true);
      expect(textWidth(out, 12, false), "the mark must be paid for").toBeLessThanOrEqual(70);
    }
    // Digits after the RTL run carry its direction on, so the last LETTER decides.
    expect(ellipsize("المبيعات حسب المنطقة 2026 2027", 12, false, 110).endsWith(`…${RLM}`)).toBe(true);
    // Latin text, a Latin cut of mixed text, and a CJK label take no mark.
    expect(ellipsize("Revenue by region and month", 12, false, 70).endsWith(RLM)).toBe(false);
    expect(ellipsize("المبيعات Sales by region and month", 12, false, 140).endsWith(RLM)).toBe(false);
    expect(ellipsize("売上高の地域別内訳と月次推移", 12, false, 70).endsWith(RLM)).toBe(false);
    // The mark is only ever added after the ellipsis, never to text that fits.
    expect(ellipsize(arabic, 12, false, 500)).toBe(arabic);
  });

  it("reaches the scene: no label of a built chart ends in a half cluster", () => {
    // 160x120 is where the labels are cut: a pie's legend and the category axis of every cartesian kind.
    const data = {
      categories: ["🇩🇰🇸🇪🇳🇴 Nordic markets", "Team 👨‍👩‍👧 growth plan", "🇸🇪🇳🇴🇫🇮🇮🇸 Region"],
      series: [{ name: "Team 👨‍👩‍👧 growth plan", values: [3, 5, 4] }],
    };
    const cut: string[] = [];
    for (const kind of ["pie", "column", "bar", "line", "table"] as const) {
      const cfg = { kind, width: 160, height: 120, data, decorations: { legend: true, segmentLabels: true } };
      for (const n of texts(buildChart(cfg as unknown as ChartConfig).nodes)) {
        if (halfCluster(n.text)) cut.push(`${kind} ${n.name}: ${n.text} (${halfCluster(n.text)})`);
      }
    }
    expect(cut).toEqual([]);
  });
});

/**
 * r3-14. The pane has no control for `map` and its tile-map sample pins "europe",
 * so US codes typed over the sample drew the Europe grid with nothing on it.
 */
describe("a pinned tile map yields to codes that name another (r3-14)", () => {
  const US = ["CA", "TX", "NY", "FL", "WA", "OR", "IL", "OH", "PA", "GA", "NC"];
  // No code of it is also a US postal code (DE, MD, ME, MT and AL are), so the US pin matches none of it.
  const EUROPE = ["DK", "SE", "NO", "FI", "NL", "FR", "GB", "ES", "IT", "PL"];
  const WORLD = ["NA", "LATAM", "EU", "MEA", "SSA", "CIS", "SA", "EA", "SEA", "OCE"];

  it("draws the grid the codes name when the pin matches none of them", () => {
    expect(resolveTileLayout("europe", US)).toBe(US_TILES);
    expect(resolveTileLayout("us", EUROPE)).toBe(EUROPE_TILES);
    expect(resolveTileLayout("world", US)).toBe(US_TILES);
    expect(resolveTileLayout("eu", US)).toBe(US_TILES);
  });

  it("keeps a pin that matches any code, or that nothing contradicts", () => {
    expect(resolveTileLayout("europe", EUROPE)).toBe(EUROPE_TILES);
    expect(resolveTileLayout("world", WORLD)).toBe(WORLD_TILES);
    // A few typos in a Europe set: the pin is the better guess.
    expect(resolveTileLayout("europe", ["DK", "SE", "XX", "YY", "ZZ"])).toBe(EUROPE_TILES);
    // Nothing is detected from these, so there is nothing to move the pin to.
    expect(resolveTileLayout("europe", ["XX", "YY"])).toBe(EUROPE_TILES);
    expect(resolveTileLayout("world", US.slice(0, 1).concat(["XX", "YY", "ZZ"]))).toBe(WORLD_TILES);
  });

  it("detects when nothing is pinned, and never answers world", () => {
    expect(resolveTileLayout(undefined, US)).toBe(US_TILES);
    expect(resolveTileLayout(undefined, EUROPE)).toBe(EUROPE_TILES);
    expect(resolveTileLayout(undefined, WORLD)).toBeUndefined();
    expect(resolveTileLayout(undefined, [])).toBeUndefined();
  });

  it("reads only the layouts it has: a pin named like an Object.prototype member is no layout", () => {
    expect(resolveTileLayout("constructor", US)).toBe(US_TILES);
    expect(resolveTileLayout("__proto__", ["XX"])).toBeUndefined();
    expect(resolveTileLayout("toString", ["XX"])).toBeUndefined();
  });

  it("draws the US grid when US codes are typed over the pane's Europe-pinned sample", () => {
    const sample = sampleConfig("tilemap");
    expect(sample.map, "the sample is the pinned one this finding is about").toBe("europe");
    const typed: ChartConfig = {
      ...sample,
      data: { categories: US, series: [{ name: "Revenue", values: [84, 61, 45, 32, 140, 66, 98, 112, 54, 72, 38] }] },
    };
    const names = buildChart(typed).nodes.map((n) => n.name);
    expect(names, "no grid at all").not.toContain("tilemap-error");
    expect(names).toContain("tile-CA");
    expect(names).toContain("tile-value-CA");
    expect(names, "the Europe grid was drawn").not.toContain("tile-IS");
    // …and the sample itself still draws Europe.
    expect(buildChart(sample).nodes.map((n) => n.name)).toContain("tile-value-DK");
  });
});
