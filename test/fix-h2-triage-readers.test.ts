import { afterEach, describe, expect, it, vi } from "vitest";
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";
// @ts-expect-error - plain .mjs tools, no types.
import * as triage from "../scripts/triage.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as pools from "../scripts/round-pools.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as gate from "../scripts/rounds-gate.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as hostHistory from "../scripts/host-history.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as regimes from "../scripts/host-regimes.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as issueStatus from "../scripts/issue-status.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as mutation from "../scripts/mutation-triage.mjs";

// Destructured after the imports so each `@ts-expect-error` stays attached to
// the one short line it suppresses.
const { dormantInstruments, roundsToJudgeOn, stakedWithoutATime, reportUpdateShortfalls } = triage;
const { poolNoiseFloor, poolWithinRoundSpread, poolDrawCostCurve, poolUpdateShortfalls, poolTagFaults } = pools;
const { spanAgainstUsual } = gate;
const { byRoundTime, readRound, history: historyOf } = hostHistory;
const { flippedTogether, readRoundSamples, explainByRegime } = regimes;
const { citedIssues } = issueStatus;
const { survivorsOf, reportBody } = mutation;

/**
 * Eleven defects in the triage and pooled-reader tools, found by a bug hunt on
 * 2026-10-24 and verified twice. Every one of them is a REPORT that said
 * something the archive does not: none touches the product, and each is pinned
 * here against the reading that was wrong.
 */

// ── r1-61 ────────────────────────────────────────────────────────────────────
describe("a round number is the whole number in the file name (r1-61)", () => {
  const log = (name: string, ...messages: string[]) => ({
    roundName: name,
    trace: { entries: messages.map((message) => ({ message })) },
  });

  it("counts silence from round 1060, not from round 106", () => {
    // `1060-x.json` was read as round 106 by `.slice(0, 3)` while `newest` stayed
    // at 999. A line that stopped at 995 then had a gap of 4 and was hidden, and
    // a line firing in every round from 1000 on was reported as 893 rounds quiet.
    const rows = dormantInstruments(
      [
        log("990-a.json", "retired at 995"),
        log("995-a.json", "retired at 995"),
        log("999-a.json", "still firing"),
        log("1000-a.json", "still firing", "introduced at 1000"),
        log("1060-a.json", "still firing", "introduced at 1000"),
      ],
      40,
      "",
    );
    expect(rows.map((r: { message: string }) => r.message)).toEqual(["retired at 995"]);
    expect(rows[0].lastRound).toBe(995);
    expect(rows[0].gap, "1060 - 995").toBe(65);
  });

  it("still reads a three-digit archive the way it always did", () => {
    const rows = dormantInstruments([log("065-a.json", "q"), log("225-b.json", "x")], 40, "");
    expect(rows[0]).toMatchObject({ message: "q", lastRound: 65, gap: 160 });
  });

  it("labels rounds 1000 and 1005 differently in the pooled readers", () => {
    // Label-only, but "1000" through "1009" all printing as "100" made the
    // NOISE FLOOR's session list unreadable.
    const entry = (ms: number, i: number) => ({
      message: "updated only the shapes that changed",
      data: { chart: `${i + 2}/8`, changed: 18, of: 24, ms },
    });
    const round = (name: string) => ({
      roundName: name,
      build: "aaaaaaa · 2026-10-24 09:00Z",
      driverRun: { sessionIndex: 1 },
      selftest: [],
      trace: { entries: [19000, 19100, 19200, 19300].map(entry) },
    });
    const floor = poolNoiseFloor([round("1000-a.json"), round("1005-a.json")]).get("aaaaaaa");
    expect(floor.map((r: { round: string }) => r.round)).toEqual(["1000", "1005"]);
    const spread = poolWithinRoundSpread([round("1000-a.json"), round("1005-a.json")]);
    expect(spread.map((r: { round: string }) => r.round)).toEqual(["1000", "1005"]);
  });
});

// ── r1-62 ────────────────────────────────────────────────────────────────────
describe("rounds of one build are ordered by their own start time (r1-62)", () => {
  const sameBuild = "9a8fb38 · 2026-09-29 20:42Z";
  const round = (n: string, startedAt?: string) => ({ build: sameBuild, path: n, startedAt, answers: {} });

  it("breaks the tie inside a build by startedAt, whatever order it is handed", () => {
    // The build stamp is when vite built the bundle, to the minute, and every
    // round of a cycle shares it. A stable sort then kept argv order.
    const a = round("495", "2026-09-29T20:42:41.718Z");
    const b = round("496", "2026-09-29T21:04:10.100Z");
    const c = round("497", "2026-09-29T21:26:15.511Z");
    const forward = [a, b, c].sort(byRoundTime).map((r) => r.path);
    const reversed = [c, b, a].sort(byRoundTime).map((r) => r.path);
    expect(forward).toEqual(["495", "496", "497"]);
    expect(reversed, "argv order leaked into the result").toEqual(forward);
  });

  it("still puts an earlier build first, whatever the start times say", () => {
    const old = { build: "aaaaaaa · 2026-09-29 19:00Z", path: "old", startedAt: "2026-09-29T23:00:00Z", answers: {} };
    const fresh = { build: sameBuild, path: "fresh", startedAt: "2026-09-29T20:50:00Z", answers: {} };
    expect([fresh, old].sort(byRoundTime).map((r) => r.path)).toEqual(["old", "fresh"]);
  });

  it("puts a round with no start time after the ones that have one, keeping its place among its own kind", () => {
    const x = round("x");
    const y = round("y");
    const dated = round("dated", "2026-09-29T21:00:00Z");
    expect([x, y, dated].sort(byRoundTime).map((r) => r.path)).toEqual(["dated", "x", "y"]);
    expect([y, x, dated].sort(byRoundTime).map((r) => r.path)).toEqual(["dated", "y", "x"]);
  });

  it("reads startedAt off the round file, and `latest` no longer depends on argv", () => {
    const dir = mkdtempSync(join(tmpdir(), "pc-hist-"));
    try {
      const write = (name: string, startedAt: string, answer: string) => {
        const p = join(dir, name);
        writeFileSync(
          p,
          JSON.stringify({
            build: sameBuild,
            driverRun: { startedAt },
            hostAnswers: {
              kind: "ssf-charts-host-answers",
              build: sameBuild,
              answers: [{ id: "q", answer, samples: [] }],
            },
          }),
        );
        return p;
      };
      const p1 = write("495-a.json", "2026-09-29T20:42:41.718Z", "yes");
      const p2 = write("496-a.json", "2026-09-29T21:04:10.100Z", "yes");
      const p3 = write("497-a.json", "2026-09-29T21:26:15.511Z", "commit-threw");
      expect(readRound(p3).startedAt).toBe("2026-09-29T21:26:15.511Z");
      const seqOf = (paths: string[]) => historyOf(paths.map(readRound).sort(byRoundTime))[0];
      const fwd = seqOf([p1, p2, p3]);
      const rev = seqOf([p3, p2, p1]);
      // Newest-first argv used to print "commit-threw yes yes" and name the
      // OLDEST round as `latest`.
      expect(rev.seq).toEqual(["yes", "yes", "commit-threw"]);
      expect(rev).toEqual(fwd);
      expect(rev.latest).toBe("commit-threw");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ── r1-63 ────────────────────────────────────────────────────────────────────
describe("a slow round is judged against rounds of the same scenario count (r1-63)", () => {
  const round = (scenarios: number, seconds: number) => ({
    selftest: Array.from({ length: scenarios }, () => ({ ok: true })),
    trace: { entries: [{ ms: 5000 }, { ms: 5000 + seconds * 1000 }] },
  });

  it("does not call an ordinary 19-scenario round slow because 13-scenario rounds are shorter", () => {
    // The archive at round 497, in miniature: a long history of 13-scenario
    // rounds near 615s, then 19-scenario rounds near 905s. The whole-archive
    // median is ~615s, so a normal 928s round read 1.5x "the usual".
    const archive = [
      ...Array.from({ length: 12 }, () => round(13, 615)),
      ...[890, 905, 910, 900].map((s) => round(19, s)),
      round(19, 928),
    ];
    const u = spanAgainstUsual(archive);
    expect(u.scenarios).toBe(19);
    expect(u.prior, "only the rounds that ran the same list").toBe(4);
    expect(u.median).toBe(905);
    expect(u.ratio, "928s against 905s is an ordinary round").toBeLessThan(1.5);
  });

  it("still flags a genuinely slow round of the same size", () => {
    const archive = [...[890, 905, 910, 900].map((s) => round(19, s)), round(19, 1500)];
    expect(spanAgainstUsual(archive).ratio).toBeGreaterThanOrEqual(1.5);
  });

  it("claims no usual when too few earlier rounds ran the same list, rather than borrowing another list's", () => {
    // The first rounds after the scenario list changes. A median of the old
    // list's rounds would be the wrong yardstick; none is the honest answer.
    const archive = [
      ...Array.from({ length: 30 }, () => round(13, 615)),
      round(19, 900),
      round(19, 905),
      round(19, 1400),
    ];
    const u = spanAgainstUsual(archive);
    expect(u.prior).toBe(2);
    expect(u.median).toBeNull();
    expect(u.ratio).toBeNull();
    expect(u.span).toBe(1400);
  });

  it("says nothing for a newest round that carries no trace", () => {
    expect(spanAgainstUsual([round(19, 900), { selftest: [], trace: { entries: [] } }])).toBeNull();
  });
});

// ── r1-64 ────────────────────────────────────────────────────────────────────
describe("a prediction staked with an ISO timestamp keeps that day's rounds (r1-64)", () => {
  const logs = [
    { build: "aaaaaaa · 2026-08-18 22:00Z" },
    { build: "bbbbbbb · 2026-08-19 13:58Z" },
    { build: "ccccccc · 2026-08-19 21:30Z" },
  ];
  const buildOf = (l: { build: string }) => l.build.split(" ")[0];
  const eligible = (madeOn: string) =>
    roundsToJudgeOn(logs, "zzzzzzz", buildOf, madeOn).map((l: { build: string }) => buildOf(l));

  it("reads the T form as the same instant as the space form", () => {
    // 'T' is 0x54 and ' ' is 0x20, so "2026-08-19 13:58" > "2026-08-19T10:00Z" was
    // false: every round later that same day compared as older and the ledger
    // said `no round yet`.
    expect(eligible("2026-08-19 10:00Z")).toEqual(["bbbbbbb", "ccccccc"]);
    expect(eligible("2026-08-19T10:00Z")).toEqual(["bbbbbbb", "ccccccc"]);
    expect(eligible("2026-08-19T10:00:00.000Z"), "toISOString() output").toEqual(["bbbbbbb", "ccccccc"]);
  });

  it("converts a numeric offset to UTC instead of comparing its digits", () => {
    // `date -Iseconds` on a Danish clock: 12:00+02:00 is 10:00Z, 16:00+02:00 is 14:00Z.
    expect(eligible("2026-08-19T12:00:00+02:00")).toEqual(["bbbbbbb", "ccccccc"]);
    expect(eligible("2026-08-19T16:00:00+02:00")).toEqual(["ccccccc"]);
  });

  it("still excludes a round taken at the very minute of the claim", () => {
    expect(eligible("2026-08-19T13:58:00Z")).toEqual(["ccccccc"]);
  });

  it("still reads a bare date as the start of that day, and a space-form time as it always did", () => {
    expect(eligible("2026-08-19")).toEqual(["bbbbbbb", "ccccccc"]);
    expect(eligible("2026-08-19 21:30Z")).toEqual([]);
  });

  it("warns about a zoneless T time with seconds, and stays quiet about a stamp that carries its zone", () => {
    const p = (madeOn: string) => ({ madeOn, afterBuild: "zzzzzzz" });
    expect(stakedWithoutATime(p("2026-08-19T10:00:00"), logs, buildOf)).toBe(true);
    expect(stakedWithoutATime(p("2026-08-19T10:00"), logs, buildOf)).toBe(true);
    expect(stakedWithoutATime(p("2026-08-19T10:00:00.000Z"), logs, buildOf)).toBe(false);
    expect(stakedWithoutATime(p("2026-08-19T10:00:00+02:00"), logs, buildOf)).toBe(false);
  });
});

// ── r2-55 ────────────────────────────────────────────────────────────────────
describe("flippedTogether drops an uninformative answer the way explainBy does (r2-55)", () => {
  const seq = (...answers: string[]) => answers.map((answer, i) => ({ answer, pass: i + 1 }));

  it("does not let `unreadable` make a question look as though it moved", () => {
    // Round 290: shape-add-fresh-getitem-slide goes yes, threw, yes and
    // picture-then-shape-read goes yes, unreadable, threw. Only the second
    // changed between its two REAL passes (1 and 3), so the two did not flip in
    // lockstep. Shipped: lockstep over 3 passes.
    const t = flippedTogether(seq("yes", "threw", "yes"), seq("yes", "unreadable", "threw"));
    expect(t.verdict).toBe("independent");
    expect(t.shared, "the unreadable pass is not a pass both answered").toBe(2);
  });

  it("and does not hide a lockstep behind one", () => {
    // Round 290's other pair: short-3, short-0, short-0 beside yes, unreadable,
    // threw. Over the two shared real passes both moved. Shipped: independent.
    const t = flippedTogether(seq("short-3", "short-0", "short-0"), seq("yes", "unreadable", "threw"));
    expect(t.verdict).toBe("lockstep");
    expect(t.shared).toBe(2);
  });

  it("still drops a pass the run never put", () => {
    const t = flippedTogether(seq("yes", "no-scratch-slide", "no"), seq("a", "b", "b"));
    expect(t.shared).toBe(2);
  });
});

// ── r2-58 ────────────────────────────────────────────────────────────────────
describe("host-regimes reads a pre-rename sheet in today's words (r2-58)", () => {
  it("does not report which-end-a-short-read-drops as a flip when it said `all` then `none`", () => {
    // Before 2026-09-01 the probe said `all` and `none` for "this question did
    // not arise"; today those are `not-a-short-read` and `none-of-ours`, both
    // uninformative. Read raw they were two real faces: 84 archived rounds listed
    // the question as EXPLAINED by regime and 3 as COIN.
    const dir = mkdtempSync(join(tmpdir(), "pc-regimes-"));
    try {
      const sample = (answer: string, pass: number, regime: string) => ({ answer, pass, atMs: pass * 1000, regime });
      const path = join(dir, "255-a.json");
      writeFileSync(
        path,
        JSON.stringify({
          hostAnswers: {
            kind: "ssf-charts-host-answers",
            build: "5bb643a · 2026-08-25 18:04Z",
            answers: [
              {
                id: "which-end-a-short-read-drops",
                answer: "all",
                samples: [
                  sample("all", 1, "healthy"),
                  sample("none", 2, "collection-refused"),
                  sample("none", 3, "collection-refused"),
                ],
              },
            ],
          },
        }),
      );
      const round = readRoundSamples(path);
      const q = round.answers[0];
      expect(q.samples.map((s: { answer: string }) => s.answer)).toEqual([
        "not-a-short-read",
        "none-of-ours",
        "none-of-ours",
      ]);
      expect(q.answer).toBe("not-a-short-read");
      expect(explainByRegime(q.samples).verdict, "a manufactured flip").toBe("steady");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("leaves every other word, and every other question, as recorded", () => {
    const dir = mkdtempSync(join(tmpdir(), "pc-regimes-"));
    try {
      const path = join(dir, "300-a.json");
      const answers = [
        { id: "other-question", answer: "all", samples: [{ answer: "all", pass: 1, regime: "healthy" }] },
        { id: "which-end-a-short-read-drops", answer: "head", samples: [{ answer: "head", pass: 1 }] },
      ];
      writeFileSync(path, JSON.stringify({ kind: "ssf-charts-host-answers", build: "b", answers }));
      const round = readRoundSamples(path);
      expect(round.answers[0].samples[0].answer, "a different question may say `all`").toBe("all");
      expect(round.answers[1].samples[0].answer).toBe("head");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ── r2-56 ────────────────────────────────────────────────────────────────────
describe("the draw-cost curve does not bin a placeholder 0 as an empty slide (r2-56)", () => {
  const issued = (upTo: number, onSlide: number, onSlideKey: string, prevBatchMs?: number) => ({
    message: "batch issued",
    data: { upTo, onSlide, onSlideKey, ...(prevBatchMs === undefined ? {} : { prevBatchMs }) },
  });

  it("leaves a sentinel-keyed first batch out of the empty-target band", () => {
    // A draw's FIRST batch is issued before the slide is named: `onSlideKey` is
    // "(visible)" and `onSlide` is a placeholder 0, though the slide may already
    // hold dozens of this run's shapes (it took 12.9s here, against 3.6s for a
    // genuinely empty slide). Binned as 0 it lifted the empty band's median.
    const entries: unknown[] = [];
    for (let i = 0; i < 12; i++)
      entries.push(
        issued(10, 0, "(visible)"),
        issued(20, 10, "slide-a", 12000), // times the sentinel batch
        issued(30, 20, "slide-a", 4000),
      );
    for (let i = 0; i < 12; i++)
      entries.push(
        issued(10, 0, "slide-b"),
        issued(20, 10, "slide-b", 3000), // times a genuinely empty slide's batch
        issued(30, 20, "slide-b", 3000),
      );
    const c = poolDrawCostCurve([{ trace: { entries } }]);
    const band0 = c.rows.find((r: { lo: number }) => r.lo === 0);
    expect(band0.median, "the sentinel rows were binned as an empty slide").toBe(3000);
    expect(band0.n).toBe(12);
    // Nothing else moved: the batch SIZES still read every timed row.
    expect(c.sizes).toEqual([10]);
  });
});

// ── r2-57 ────────────────────────────────────────────────────────────────────
describe("the update-shortfall report speaks only over readings it pooled (r2-57)", () => {
  const reading = (over: Record<string, unknown>) => ({
    scope: "update",
    message: "shapes left on the slide after an in-place update",
    data: { slideId: "s1", before: 1, after: 1, growth: 0, settled: true, atRisk: 0, ...over },
  });
  const notGrouping = (chart: string | undefined, message: string) => ({
    scope: "group",
    message,
    data: chart === undefined ? { index: 0 } : { chart, index: 0 },
  });

  const printed = (logs: unknown[]) => {
    const lines: string[] = [];
    const spy = vi.spyOn(console, "log").mockImplementation((...a: unknown[]) => void lines.push(a.join(" ")));
    try {
      reportUpdateShortfalls(logs);
    } finally {
      spy.mockRestore();
    }
    return lines.join("\n");
  };
  afterEach(() => vi.restoreAllMocks());

  it("counts the readings it pooled, so an unsettled-only pool has none", () => {
    const o = poolUpdateShortfalls([
      { trace: { entries: [reading({ settled: false, atRisk: 1, growth: 3 }), reading({ settled: false })] } },
    ]);
    expect(o.updates).toBe(2);
    expect(o.unsettledKept).toBe(2);
    expect(o.pooled).toBe(0);
    expect(o.unsettledBlind, "an at-risk chart among the disputed readings").toBe(1);
    expect(o.atRisk, "disputed readings must not feed the all-clear arithmetic").toBe(0);
    expect(poolUpdateShortfalls([{ trace: { entries: [reading({}), reading({ settled: false })] } }]).pooled).toBe(1);
  });

  it("prints no growth figures and no verdict over a pool of nothing but disputed readings", () => {
    // Round 497 in miniature: two in-place updates, both `settled: false`. It
    // printed "GREW by 0" and "NOT AN ALL-CLEAR ... (0 at risk — every one was
    // grouped...)" over zero pooled measurements.
    const out = printed([{ trace: { entries: [reading({ settled: false }), reading({ settled: false })] } }]);
    expect(out).toContain("no usable reading yet");
    expect(out).toContain("2 had two host reads that disagreed");
    expect(out).not.toContain("the slide GREW");
    expect(out).not.toContain("NOT AN ALL-CLEAR");
  });

  it("does not say 'predates the instrument's fix' about a reading that does not", () => {
    const out = printed([{ trace: { entries: [reading({ settled: false })] } }]);
    expect(out).not.toContain("predate the instrument's fix");
  });

  it("does not print '0 at risk' as the whole story when a disputed reading touched an at-risk chart", () => {
    // Rounds 132 and 140: a pooled settled reading beside two disputed ones that
    // carried `atRisk: 1`.
    const out = printed([
      {
        trace: {
          entries: [reading({}), reading({ settled: false, atRisk: 1 }), reading({ settled: false, atRisk: 1 })],
        },
      },
    ]);
    expect(out).toContain("2 of those touched a chart AT RISK");
    expect(out).toContain("bar the 2 disputed reading(s)");
    expect(out).toContain("NOT AN ALL-CLEAR");
  });

  it("counts a chart declined twice once", () => {
    // Every `the positional guess named no shape of ours` decline is followed by
    // `no member handle this host will accept` for the SAME chart.
    const o = poolUpdateShortfalls([
      {
        trace: {
          entries: [
            reading({}),
            notGrouping("draw-5", "not grouping: the positional guess named no shape of ours"),
            notGrouping("draw-5", "not grouping: no member handle this host will accept"),
            notGrouping("draw-6", "not grouping: no member handle this host will accept"),
          ],
        },
      },
    ]);
    expect(o.ungroupedCharts).toBe(2);
  });

  it("still counts a line with no chart label one by one, as the early rounds always were", () => {
    const o = poolUpdateShortfalls([
      {
        trace: {
          entries: [
            reading({}),
            notGrouping(undefined, "not grouping: no member handle this host will accept"),
            notGrouping(undefined, "not grouping: no member handle this host will accept"),
          ],
        },
      },
    ]);
    expect(o.ungroupedCharts).toBe(2);
  });
});

// ── r2-59 ────────────────────────────────────────────────────────────────────
describe("the 5010 tag-fault columns count one 5010 once (r2-59)", () => {
  const failure = (step: string, code: string | null) => {
    const error = `InvalidParam passed to GetItem(id) | at=${step}${code ? ` | code=${code}` : ""} | debugInfo={}`;
    return [
      // step()'s own line: the message is the step's label.
      { scope: "error", message: step, data: { error } },
      // and the catch's echo of the same error.
      { scope: "tag", message: "tagging failed — charts are not re-editable until repaired", data: { error } },
    ];
  };
  const faults = (entries: unknown[]) =>
    poolTagFaults([{ build: "aaaaaaa · 2026-10-24 09:00Z", trace: { entries } }]).get("aaaaaaa")[0];

  it("counts one failed config-tag write as one, not two", () => {
    // Round 497: one ErrorPointer failure printed cfg-tag-5010 = 2. This one is
    // a real 5010, and still a single event.
    const f = faults(failure("writing the chart's config tag", "5010"));
    expect(f["cfg-tag-5010"]).toBe(1);
    expect(f["tagging-failed"], "the echo keeps its own column").toBe(1);
  });

  it("does not count another error code, or a timeout, as a 5010", () => {
    // Round 282 had four config-tag failures and no 5010: three 45s timeouts
    // (no code at all) and one GeneralException.
    const f = faults([
      ...failure("writing the chart's config tag", "ErrorPointer"),
      ...failure("writing the chart's config tag", "GeneralException"),
      ...failure("writing the chart's config tag", null),
    ]);
    expect(f["cfg-tag-5010"]).toBe(0);
    expect(f["tagging-failed"]).toBe(3);
  });

  it("holds group-5010 to the same code", () => {
    const f = faults([
      {
        scope: "error",
        message: "grouping the chart's shapes",
        data: { error: "x | at=grouping the chart's shapes | code=5010 | d" },
      },
      {
        scope: "error",
        message: "grouping the chart's shapes",
        data: { error: "x | at=grouping the chart's shapes | code=InvalidArgument | d" },
      },
    ]);
    expect(f["group-5010"]).toBe(1);
  });
});

// ── r2-60 ────────────────────────────────────────────────────────────────────
describe("issue-status finds an issue cited by its URL (r2-60)", () => {
  const write = (root: string, path: string, text: string) => {
    mkdirSync(join(root, path, ".."), { recursive: true });
    writeFileSync(join(root, path), text);
  };

  it("sees the markdown-link form, the bare form, and a root file, and ignores an uncited #NNNN", () => {
    const root = mkdtempSync(join(tmpdir(), "pc-issues-"));
    try {
      write(
        root,
        "docs/RESEARCH.md",
        "Rejected: [#5896](https://github.com/OfficeDev/office-js/issues/5896) and office-js#2780; see also #4444.\n",
      );
      write(root, "CLAUDE.md", "See https://github.com/OfficeDev/office-js/issues/6266 for the rest.\n");
      write(root, "src/a.ts", "// office-js #3014 says so\n");
      write(root, "docs/evidence/x.json", '{"note": "office-js#3565"}');
      const found = citedIssues(root);
      expect([...found.keys()].sort()).toEqual(["2780", "3014", "5896", "6266"]);
      expect([...found.get("5896")]).toEqual([join(root, "docs/RESEARCH.md").replace(/\\/g, "/")]);
      expect([...found.get("6266")]).toEqual(["CLAUDE.md"]);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it("agrees with test/office-js-watch.test.ts about what this repo cites", () => {
    // The two tools disagreed: that test counted the URL form and this one did
    // not, so #2881, #3309, #5896 and #6266 were never re-read upstream. Its
    // definition, repeated here on purpose, must be a subset of this tool's.
    const mine = citedIssues(".");
    const theirs = new Set<string>();
    const walk = (dir: string) => {
      for (const name of readdirSync(dir)) {
        if (name === "node_modules" || name === ".git" || name === "dist-lib") continue;
        const path = `${dir}/${name}`;
        if (statSync(path).isDirectory()) walk(path);
        else if (/\.(ts|mjs|md)$/.test(name)) {
          for (const m of readFileSync(path, "utf8").matchAll(/office-js[#/](?:issues\/)?(\d{3,6})/g)) theirs.add(m[1]);
        }
      }
    };
    for (const dir of ["src", "docs", "scripts", "test"]) walk(dir);
    const missing = [...theirs].filter((n) => !mine.has(n)).sort();
    expect(missing, "cited per office-js-watch.test.ts but never re-read by issue-status").toEqual([]);
  });
});

// ── r2-61 ────────────────────────────────────────────────────────────────────
describe("mutation-triage counts a mutant no test ran (r2-61)", () => {
  const report = (status: string) => ({
    files: {
      "src/core/new-feature.ts": {
        mutants: [
          { status, mutatorName: "EqualityOperator", location: { start: { line: 12 } }, replacement: "a > b" },
          { status, mutatorName: "BooleanLiteral", location: { start: { line: 20 } }, replacement: "false" },
        ],
      },
      "src/core/chart.ts": {
        mutants: [{ status: "Killed", mutatorName: "BooleanLiteral", location: { start: { line: 1 } } }],
      },
    },
  });

  it("treats NoCoverage as a survivor, marked as one no test runs", () => {
    const rows = survivorsOf(report("NoCoverage"));
    expect(rows.map((r: { key: string }) => r.key)).toEqual([
      "src/core/new-feature.ts:12:EqualityOperator",
      "src/core/new-feature.ts:20:BooleanLiteral",
    ]);
    expect(rows[0].uncovered).toBe(true);
    const body = reportBody(rows, rows, []);
    expect(body).toContain("2 surviving mutant(s) nobody has ruled on, of 2");
    expect(body).toContain("`a > b` — no test runs this line");
  });

  it("leaves a Survived row exactly as it was", () => {
    const rows = survivorsOf(report("Survived"));
    expect(rows).toHaveLength(2);
    expect(rows[0]).not.toHaveProperty("uncovered");
    expect(reportBody(rows, rows, [])).not.toContain("no test runs this line");
  });

  it("exits 3, the finding, for a week whose only undetected mutants are NoCoverage", () => {
    // The workflow files an issue on exit 3 and only then. A week of new code
    // with no tests used to exit 0 with "0 surviving mutant(s)".
    const dir = mkdtempSync(join(tmpdir(), "pc-mut-"));
    try {
      const path = join(dir, "mutation.json");
      writeFileSync(path, JSON.stringify(report("NoCoverage")));
      const r = spawnSync(process.execPath, ["scripts/mutation-triage.mjs", path], { encoding: "utf8" });
      expect(r.status).toBe(3);
      expect(r.stdout).toContain("2 surviving mutant(s) nobody has ruled on");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
