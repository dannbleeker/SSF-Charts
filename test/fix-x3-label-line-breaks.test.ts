import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { legendRowCount } from "../src/core/layout/frame";
import { sceneToSvg } from "../src/render/svg";
import {
  LINE_PITCH,
  ellipsize,
  labelLines,
  lineCount,
  textWidth,
  type LineNode,
  type Scene,
  type TextNode,
} from "../src/core/scene";
import { clipToWidth } from "../src/core/elements";
import type { ChartConfig } from "../src/core/types";

/**
 * x3 / r1-22: A LABEL'S LINE BREAK SIZES ITS BOX.
 *
 * `0e9500f` made the preview and the generated deck DRAW a break in a category or
 * series name as two lines, as PowerPoint does (a datasheet cell takes Alt+Enter,
 * and MANUAL.md says the break stays a break). The LAYOUT was left reading the
 * string as one run: `textWidth` counted the break as a character, so `Q1\nReport`
 * was judged 9 characters wide, and every box was one line tall. Measured at
 * 480x300 on a clustered chart with a category named `Q1\nReport`: the box was 14
 * tall, the plot reserved one line for the strip, and the second line's ink ended
 * at y=307 — seven points BELOW the chart, clipped by the preview's viewBox and
 * drawn over whatever sits under the chart on a slide.
 *
 * Every assertion below is about what a person sees: the lines of a label stay
 * inside the box that holds them, the box stays on the chart, and the ink is
 * where the PREVIEW puts it (the second `<tspan>` is read back out of the SVG, so
 * the layout is tied to the thing that actually draws the lines).
 */

/** Every spelling of a line break the sinks split (see `labelLines`). */
const BREAKS: [string, string][] = [
  ["CRLF", "\r\n"],
  ["LF", "\n"],
  ["CR", "\r"],
  ["VT", "\v"],
  ["FF", "\f"],
  ["NEL", String.fromCharCode(0x85)],
  ["U+2028", String.fromCharCode(0x2028)],
  ["U+2029", String.fromCharCode(0x2029)],
];

const FS = 12;

/** The right-hand series name's one-line box: `fontSize * 1.35`, at the samples' 10pt. */
const GUTTER_LINE_H = 10 * 1.35;

const clustered = (over: Partial<ChartConfig> = {}): ChartConfig =>
  ({ ...JSON.parse(JSON.stringify(sampleConfig("clustered"))), ...DEFAULT_SIZE, ...over }) as ChartConfig;

/** The sample with its first category replaced by `name`. */
const withFirstCategory = (name: string, over: Partial<ChartConfig> = {}): ChartConfig => {
  const cfg = clustered(over);
  cfg.data.categories[0] = name;
  return cfg;
};

const textNodes = (scene: Scene, prefix: string): TextNode[] =>
  scene.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith(prefix));
const byName = (scene: Scene, name: string): TextNode => {
  const n = scene.nodes.find((x): x is TextNode => x.kind === "text" && x.name === name);
  expect(n, `no ${name} node`).toBeTruthy();
  return n!;
};
const lineNamed = (scene: Scene, name: string): LineNode => {
  const n = scene.nodes.find((x): x is LineNode => x.kind === "line" && x.name === name);
  expect(n, `no ${name} line`).toBeTruthy();
  return n!;
};

/**
 * Where the PREVIEW draws a label's last baseline, read out of the SVG: the `y` of
 * the last `<tspan>` of the text carrying `data-name`. The layout's box has to
 * hold this, so the number comes from the renderer and not from the layout's own
 * arithmetic.
 */
function lastBaselineInSvg(scene: Scene, name: string): number {
  const svg = sceneToSvg(scene);
  const text = new RegExp(`<text [^>]*data-name="${name}"[^>]*>(.*?)</text>`).exec(svg);
  expect(text, `${name} is not in the SVG`).toBeTruthy();
  const ys = [...text![1].matchAll(/<tspan [^>]*y="([-\d.]+)"/g)].map((m) => Number(m[1]));
  expect(ys.length, `${name} was not drawn as tspans`).toBeGreaterThan(1);
  return ys[ys.length - 1];
}

/** Descent below a baseline, as a share of the font — the preview's own 0.25. */
const DESCENT = 0.25;

describe("a label's lines, measured", () => {
  it("splits every spelling of a break, with CRLF as one", () => {
    for (const [what, b] of BREAKS) {
      expect(labelLines(`Q1${b}Report`), what).toEqual(["Q1", "Report"]);
      expect(lineCount(`Q1${b}Report`), what).toBe(2);
      // A trailing break has no ink on its last line; a blank line between two
      // others is a line the author typed.
      expect(lineCount(`Q1${b}`), `${what} trailing`).toBe(1);
    }
    expect(lineCount("A\n\nB")).toBe(3);
    expect(lineCount("Q1 Report")).toBe(1);
    expect(lineCount(2024)).toBe(1);
    expect(lineCount(null)).toBe(1);
  });

  it("reads the WIDEST line, whichever line it is and whatever the spelling", () => {
    for (const [what, b] of BREAKS) {
      const widest = textWidth("Report", FS);
      expect(textWidth(`Q1${b}Report`, FS), `${what}, widest last`).toBeCloseTo(widest, 9);
      expect(textWidth(`Report${b}Q1`, FS), `${what}, widest first`).toBeCloseTo(widest, 9);
      expect(textWidth(`Q1${b}Report${b}Q3`, FS), `${what}, widest between`).toBeCloseTo(widest, 9);
      expect(textWidth(`Q1${b}Report`, FS, true), `${what} bold`).toBeCloseTo(textWidth("Report", FS, true), 9);
    }
    // The red this guards: the break was charged as a character, so the label
    // came out as wide as the one run it is NOT drawn as.
    expect(textWidth("Q1\nReport", FS)).toBeLessThan(textWidth("Q1 Report", FS));
  });

  it("leaves text with no break exactly as it was measured", () => {
    // The sum the function has always computed, bit for bit.
    expect(textWidth("Report", FS)).toBe(6 * 0.54 * FS);
    expect(textWidth("Report", FS, true)).toBe(6 * 0.58 * FS);
    expect(textWidth("", FS)).toBe(0);
    expect(textWidth("売上高", FS)).toBe(3 * FS);
  });

  it("clips each LINE to the room, keeping the breaks as spelled", () => {
    const room = textWidth("AAAAA", 10);
    for (const [what, b] of BREAKS) {
      const cut = clipToWidth(`AAAAAAAAAAAAAAAA${b}BB`, 10, room);
      // The second line fits and must survive whole; the first is cut.
      expect(cut.endsWith(`${b}BB`), `${what}: ${JSON.stringify(cut)}`).toBe(true);
      expect(cut.startsWith("AAA"), what).toBe(true);
      expect(cut, what).toContain("…");
      expect(textWidth(cut, 10), what).toBeLessThanOrEqual(room + 1e-9);
    }
    // Nothing fits on any line: nothing is kept, as for one line.
    expect(ellipsize("AAAA\nBBBB", 10, false, 1)).toBe("");
    // A line that already fits is returned untouched.
    expect(clipToWidth("AB\nCD", 10, 100)).toBe("AB\nCD");
  });
});

describe("the upright category axis", () => {
  for (const [what, b] of BREAKS) {
    it(`holds a name broken by ${what}: box, strip and ink all inside the chart`, () => {
      const scene = buildChart(withFirstCategory(`Q1${b}Report`));
      const broken = byName(scene, "category-0");
      const plain = byName(scene, "category-1");

      // The box is one pitch of its own font taller than a one-line name's...
      expect(broken.h - broken.fontSize * 1.4).toBeCloseTo(broken.fontSize * LINE_PITCH, 9);
      // ...and a name with no break keeps the box it always had, beside it.
      expect(plain.h).toBe(plain.fontSize * 1.4);
      expect(broken.y).toBe(plain.y);

      // The last line's ink, where the preview draws it, is inside the box that
      // holds it AND inside the chart.
      const inkBottom = lastBaselineInSvg(scene, "category-0") + broken.fontSize * DESCENT;
      expect(inkBottom, "the second line is drawn below its own box").toBeLessThanOrEqual(broken.y + broken.h + 1e-9);
      expect(inkBottom, "the second line is drawn below the chart").toBeLessThanOrEqual(scene.height);
    });
  }

  it("pays for the extra line out of the plot, so the strip starts under the baseline", () => {
    const one = buildChart(clustered());
    const two = buildChart(withFirstCategory("Q1\nReport"));
    const pitch = one.nodes.length && byName(one, "category-0").fontSize * LINE_PITCH;
    // The zero line is the plot's foot. One line more in the strip lifts it by
    // exactly that line, at the chart font the strip was priced at.
    expect(lineNamed(one, "baseline").y1 - lineNamed(two, "baseline").y1).toBeCloseTo(pitch, 9);
    // And the names sit below it, the way the one-line strip's always did.
    expect(byName(two, "category-0").y).toBeGreaterThan(lineNamed(two, "baseline").y1);
  });

  it("fits a broken name to its slot by its widest line, not by the two lines laid end to end", () => {
    // Narrow enough that the names must shrink. `AAAAAA\nAAAAAA` is drawn as
    // two 6-letter lines, so it needs the room `AAAAAA` needs; read as one run it
    // is 13 letters wide and pulled the whole axis down to a smaller size.
    const narrow = (names: string[]): ChartConfig => {
      const cfg = clustered({ width: 150 });
      cfg.data.categories = names;
      return cfg;
    };
    const oneLine = buildChart(narrow(["AAAAAA", "AAAAAA", "AAAAAA", "AAAAAA"]));
    const twoLines = buildChart(narrow(["AAAAAA\nAAAAAA", "AAAAAA", "AAAAAA", "AAAAAA"]));
    const shrunk = byName(oneLine, "category-0").fontSize;
    expect(shrunk, "the fixture is not narrow enough to shrink").toBeLessThan(10);
    expect(byName(twoLines, "category-0").fontSize).toBe(shrunk);
  });

  it("drops the whole strip, as for a one-line name, when the lines cannot fit inside the chart", () => {
    // 60pt tall: one line of names fits under the plot, three do not. A name
    // printed under the chart is on the slide, not on the chart.
    const tall = clustered({ height: 60, width: 300 });
    tall.title = undefined;
    tall.data.categories[0] = "A\nB\nC\nD\nE";
    const scene = buildChart(tall);
    for (const n of textNodes(scene, "category-")) {
      const ink = n.y + n.fontSize + (lineCount(n.text) - 1) * n.fontSize * LINE_PITCH + n.fontSize * DESCENT;
      expect(ink, `${n.name} is drawn below a ${scene.height}pt chart`).toBeLessThanOrEqual(scene.height + 1e-9);
    }
  });
});

describe("the upright value axis", () => {
  const withSuffix = (suffix: string): ChartConfig =>
    clustered({
      numberFormat: { decimals: 0, suffix },
      decorations: { valueAxis: true, gridlines: true, seriesLabels: false },
    });

  it("grows a tick label that holds a break and keeps it centred on its tick", () => {
    const flat = textNodes(buildChart(withSuffix(" €m")), "value-axis").filter((n) => n.name === "value-axis");
    const broken = textNodes(buildChart(withSuffix("\n€m")), "value-axis").filter((n) => n.name === "value-axis");
    expect(flat.length).toBeGreaterThan(2);
    expect(broken.length).toBe(flat.length);
    broken.forEach((n, i) => {
      expect(lineCount(n.text), n.text).toBe(2);
      expect(n.fontSize).toBe(flat[i].fontSize);
      // One pitch of the font it is drawn at, on top of the one-line box...
      expect(n.h - flat[i].h).toBeCloseTo(n.fontSize * LINE_PITCH, 9);
      // ...grown both ways, so its middle is still on the tick.
      expect(n.y + n.h / 2).toBeCloseTo(flat[i].y + flat[i].h / 2, 9);
    });
  });

  it("keeps the unit's last line where a one-line unit sits and lets the rest rise", () => {
    // Totals band above the plot, a title above that: room for the lines to rise.
    const unit = (valueAxisTitle: string): TextNode | undefined =>
      buildChart(clustered({ valueAxisTitle, decorations: { valueAxis: true, totals: true, seriesLabels: false } }))
        .nodes.filter((n): n is TextNode => n.kind === "text")
        .find((n) => n.name === "value-axis-title");
    const flat = unit("EUR m");
    const broken = unit("Revenue\nEUR m");
    expect(flat, "the one-line unit is missing").toBeTruthy();
    expect(broken, "the two-line unit was dropped").toBeTruthy();
    const rise = flat!.fontSize * LINE_PITCH;
    expect(broken!.h - flat!.h).toBeCloseTo(rise, 9);
    // Bottom-anchored: the bottom edge, and so the last line, has not moved.
    expect(broken!.y + broken!.h).toBeCloseTo(flat!.y + flat!.h, 9);
    expect(broken!.y).toBeCloseTo(flat!.y - rise, 9);
  });

  it("drops a unit whose raised lines would leave the chart, rather than printing above it", () => {
    // No title and no totals band: the unit already starts at the top edge.
    const bare = clustered({ valueAxisTitle: "Revenue\nEUR m", decorations: { valueAxis: true, seriesLabels: false } });
    bare.title = undefined;
    const unit = (cfg: ChartConfig) =>
      buildChart(cfg).nodes.find((n) => n.kind === "text" && n.name === "value-axis-title") as TextNode | undefined;
    // Its first line would start 7pt above the chart (bandTop 0, ink top -7.4).
    expect(unit(bare), "a unit drawn above the chart").toBeUndefined();
    // The same unit on ONE line is untouched: only the lines it cannot hold go.
    const oneLine = clustered({ valueAxisTitle: "Revenue EUR m", decorations: bare.decorations });
    oneLine.title = undefined;
    expect(unit(oneLine), "the one-line unit was dropped too").toBeTruthy();
  });
});

describe("the legend", () => {
  const sideways = (name: string): ChartConfig => {
    const cfg = clustered({ horizontal: true });
    cfg.data.series[0].name = name;
    return cfg;
  };

  it("counts the pitches a broken name takes, and the wrap row under it starts lower", () => {
    const x0 = 40;
    const maxX = 476;
    expect(legendRowCount(["Revenue", "2025"], 12, x0, maxX)).toBe(1);
    expect(legendRowCount(["Revenue\n(EUR m)", "2025"], 12, x0, maxX)).toBe(2);
    // Three lines: 1.4 + 2 * 1.2 = 3.8 font heights, which is three 1.6 pitches.
    expect(legendRowCount(["A\nB\nC", "2025"], 12, x0, maxX)).toBe(3);
    // Wrapped: the second wrap row starts under the TALLEST label of the first.
    const long = "Wholesale revenue ".repeat(3).trim();
    expect(legendRowCount([`${long}\nsecond`, long], 12, x0, 200)).toBeGreaterThan(
      legendRowCount([long, long], 12, x0, 200),
    );
  });

  it("boxes a broken series name by its lines and reserves them above the plot", () => {
    const flat = buildChart(sideways("Revenue (EUR m)"));
    const broken = buildChart(sideways("Revenue\n(EUR m)"));
    const f = byName(flat, "legend-0");
    const g = byName(broken, "legend-0");
    expect(g.h - f.h).toBeCloseTo(g.fontSize * LINE_PITCH, 9);
    // The neighbouring name keeps its one-line box.
    expect(byName(broken, "legend-1").h).toBe(byName(flat, "legend-1").h);
    // Reserved: the plot starts lower by whole legend row pitches (1.6 font).
    const lowered = lineNamed(broken, "baseline").y1 - lineNamed(flat, "baseline").y1;
    expect(lowered).toBeCloseTo(g.fontSize * 1.6, 9);
    // And the box is above the plot it was made room for.
    expect(g.y + g.h).toBeLessThanOrEqual(lineNamed(broken, "baseline").y1 + 1e-9);
    // The chip rides the middle of the block.
    const chip = broken.nodes.find((n) => n.kind === "rect" && n.name === "legend-chip-0") as unknown as {
      y: number;
      h: number;
    };
    expect(chip.y + chip.h / 2).toBeCloseTo(g.y + g.h / 2, 9);
  });
});

describe("the right-hand series names", () => {
  /** A sample with its first two series named over two lines each. */
  const broken = (kind: "clustered" | "stacked"): ChartConfig => {
    const cfg = { ...JSON.parse(JSON.stringify(sampleConfig(kind))), ...DEFAULT_SIZE } as ChartConfig;
    cfg.data.series[0].name = "Revenue\n(EUR m)";
    cfg.data.series[1].name = "Cost\n(EUR m)";
    return cfg;
  };

  for (const kind of ["clustered", "stacked"] as const) {
    it(`pitches ${kind} names by their own height: no block runs into the next, and each holds its ink`, () => {
      const scene = buildChart(broken(kind));
      const labels = textNodes(scene, "series-label-").sort((a, b) => a.y - b.y);
      expect(labels.length).toBeGreaterThan(1);
      const twoLines = labels.filter((l) => lineCount(l.text) === 2);
      expect(twoLines.length, "the fixture lost its broken names").toBeGreaterThan(1);
      labels.forEach((l, i) => {
        // The box is the one-line box plus a pitch for each further line...
        const extra = (lineCount(l.text) - 1) * l.fontSize * LINE_PITCH;
        expect(l.h, l.name).toBeCloseTo(GUTTER_LINE_H + extra, 9);
        // ...and a name's block does not run into the next name's.
        if (i > 0) {
          const above = labels[i - 1];
          expect(above.y + above.h, `${above.name} runs into ${l.name}`).toBeLessThanOrEqual(l.y + 1e-9);
        }
        // Its last line, where the preview draws it, is inside its own box and the chart.
        if (lineCount(l.text) > 1) {
          const ink = lastBaselineInSvg(scene, l.name!) + l.fontSize * DESCENT;
          expect(ink, `${l.name} ink leaves its box`).toBeLessThanOrEqual(l.y + l.h + 1e-9);
          expect(ink).toBeLessThanOrEqual(scene.height);
        }
        expect(l.y, `${l.name} is above the chart`).toBeGreaterThanOrEqual(0);
      });
    });
  }

  it("keeps a one-line name its one-line box beside broken ones", () => {
    const scene = buildChart(broken("stacked"));
    const one = textNodes(scene, "series-label-").filter((l) => lineCount(l.text) === 1);
    expect(one.length).toBeGreaterThan(0);
    for (const l of one) expect(l.h, l.name).toBe(GUTTER_LINE_H);
  });
});

describe("the sideways value axis", () => {
  const sidewaysWith = (suffix: string): ChartConfig =>
    clustered({
      horizontal: true,
      numberFormat: { decimals: 0, suffix },
      decorations: { valueAxis: true, categoryAxis: true, seriesLabels: false },
    });

  it("reserves the extra line under the plot and grows each tick label by it", () => {
    const flat = buildChart(sidewaysWith(" m"));
    const broken = buildChart(sidewaysWith("\nm"));
    const ticksOf = (s: Scene) => textNodes(s, "value-axis");
    const foot = (s: Scene) => lineNamed(s, "baseline").y2;
    expect(ticksOf(broken).length, "the ticks were dropped").toBe(ticksOf(flat).length);
    ticksOf(broken).forEach((n, i) => {
      const was = ticksOf(flat)[i];
      expect(lineCount(n.text), n.text).toBe(2);
      expect(n.h - was.h).toBeCloseTo(n.fontSize * LINE_PITCH, 9);
      // Hung from the foot of the plot as the one-line labels are...
      expect(n.y).toBeCloseTo(foot(broken) + 2, 9);
      // ...and the second line, where the preview draws it, is inside the box and
      // the chart.
      const ink = n.y + n.fontSize + n.fontSize * LINE_PITCH + n.fontSize * DESCENT;
      expect(ink).toBeLessThanOrEqual(n.y + n.h + 1e-9);
      expect(ink).toBeLessThanOrEqual(broken.height);
    });
    // The strip is paid for out of the plot: its foot rises by one pitch of the
    // tick font the strip was priced at (0.9 of the chart font).
    expect(foot(flat) - foot(broken)).toBeCloseTo(10 * 0.9 * LINE_PITCH, 9);
  });
});

describe("the sideways category axis", () => {
  it("grows a broken name about its row's middle and keeps its neighbours' boxes", () => {
    const flat = buildChart(clustered({ horizontal: true }));
    const broken = buildChart(withFirstCategory("Q1\nReport", { horizontal: true }));
    const a = byName(flat, "category-0");
    const b = byName(broken, "category-0");
    expect(b.h - a.h).toBeCloseTo(b.fontSize * LINE_PITCH, 9);
    expect(b.y + b.h / 2).toBeCloseTo(a.y + a.h / 2, 9);
    expect(byName(broken, "category-1").h).toBe(byName(flat, "category-1").h);
    // The two lines of ink sit inside the box that was grown for them.
    const inkBottom = lastBaselineInSvg(broken, "category-0") + b.fontSize * DESCENT;
    expect(inkBottom).toBeLessThanOrEqual(b.y + b.h + 1e-9);
  });
});
