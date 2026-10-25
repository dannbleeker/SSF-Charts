import { afterEach, describe, expect, it, vi } from "vitest";
import { createHash } from "node:crypto";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
// @ts-expect-error - plain .mjs tools, no types.
import * as pools from "../scripts/round-pools.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as triage from "../scripts/triage.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import { CLAIMS } from "../scripts/claims.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as deploy from "../scripts/await-deploy.mjs";
// @ts-expect-error - plain .mjs tools, no types.
import * as buildSkill from "../scripts/build-skill.mjs";

// Destructured after the imports so each `@ts-expect-error` stays attached to
// the one short line it suppresses.
const { poolFreshVsEstablished } = pools;
const { probeFlipsWithinBuild, passPositionBias } = triage;
const { builtShaOf, judgeDeploy, judgeCi, awaitDeploy } = deploy;
const { zipPackage } = buildSkill;

/**
 * Regressions behind stream k7-harness-gaps-and-docs (third bug hunt): readers
 * the earlier fixes missed, and sentences that stated something false. Every
 * block pins the SIBLING of a defect that was fixed in one reader and left in
 * another, so each says which fix it completes.
 */

type Claim = { id: string; check: (logs: unknown[]) => { ok: boolean | null; actual: string } };
const claim = (id: string): Claim => {
  const found = (CLAIMS as Claim[]).find((c) => c.id === id);
  if (!found) throw new Error(`no claim ${id}`);
  return found;
};

// ── r3-45 ────────────────────────────────────────────────────────────────────
describe("a chart whose slide was not yet named is not 'freshly added, empty' (r3-45)", () => {
  /**
   * A chart's batches the way the draw writes them. The first batch is issued
   * before the host has named the slide, so it is keyed `(visible)` and carries a
   * placeholder `onSlide: 0`; the second batch carries the count on the real key,
   * which includes the first batch's own shapes.
   */
  const named = "257#1";
  const batch = (chart: string, upTo: number, onSlide: number, onSlideKey: string | undefined) => ({
    message: "batch issued",
    data: { chart, upTo, total: 16, onSlide, ...(onSlideKey === undefined ? {} : { onSlideKey }) },
  });
  const outcome = (chart: string, grouped: boolean) => ({
    message: grouped ? "grouped the chart's shapes" : "not grouping: refused",
    data: { chart },
  });
  /** Named from the start: `prior` is what the run had drawn on the slide. */
  const onNamedSlide = (chart: string, prior: number, grouped = true) => [
    batch(chart, 10, prior, named),
    batch(chart, 16, prior + 10, named),
    outcome(chart, grouped),
  ];
  /** Sentinel first batch, then a named one: the true prior is `second - first.upTo`. */
  const sentinelThenNamed = (chart: string, prior: number, grouped = true) => [
    batch(chart, 10, 0, "(visible)"),
    batch(chart, 16, prior + 10, named),
    outcome(chart, grouped),
  ];

  it("places a sentinel-keyed chart by its second batch, on either side", () => {
    const f = poolFreshVsEstablished([
      {
        trace: {
          entries: [
            ...onNamedSlide("a", 0),
            // 16 shapes of this run's own were already on the slide it landed on.
            ...sentinelThenNamed("b", 16),
            ...sentinelThenNamed("c", 0, false),
          ],
        },
      },
    ]);
    // `b` is NOT fresh: before the fix it was counted in the empty-slide row
    // because its placeholder read 0, so fresh was 3 and established 0.
    expect(f.established, "a chart on a slide this run had already drawn on").toBe(1);
    expect(f.establishedGrouped).toBe(1);
    expect(f.fresh, "the named empty slide and the recovered empty one").toBe(2);
    expect(f.freshGrouped).toBe(1);
    expect(f.unresolved).toBe(0);
  });

  it("counts a chart it cannot place in neither row", () => {
    const f = poolFreshVsEstablished([
      {
        trace: {
          entries: [
            // A draw that fits one batch never meets a named slide.
            batch("only", 10, 0, "(visible)"),
            outcome("only", true),
            // A second batch that is still unnamed says nothing either.
            batch("unnamed", 10, 0, "(visible)"),
            batch("unnamed", 16, 10, "(visible)"),
            outcome("unnamed", true),
            // Numbers that do not subtract to a count are not a reading.
            batch("odd", 10, 0, "(visible)"),
            batch("odd", 16, 4, named),
            outcome("odd", true),
          ],
        },
      },
    ]);
    expect(f).toMatchObject({ established: 0, fresh: 0, unresolved: 3 });
    expect(f.freshGrouped, "an unplaced chart must not inflate the fresh row").toBe(0);
  });

  it("never reads a sentinel's own count, whatever it says", () => {
    // A single-batch draw banks under the sentinel for good, so the shared count
    // can be above 0 and still name no slide: that is not an established slide.
    const f = poolFreshVsEstablished([
      { trace: { entries: [batch("shared", 10, 30, "(visible)"), outcome("shared", true)] } },
    ]);
    expect(f).toMatchObject({ established: 0, fresh: 0, unresolved: 1 });
  });

  it("still reads a row that carries no slide key as written", () => {
    // Fixtures and the oldest rounds have no `onSlideKey`; the absence says
    // nothing about which kind they are, as `poolDrawCostCurve` already holds.
    const f = poolFreshVsEstablished([
      {
        trace: {
          entries: [
            { message: "batch issued", data: { chart: "x", onSlide: 32 } },
            outcome("x", true),
            { message: "batch issued", data: { chart: "y", onSlide: 0 } },
            outcome("y", false),
          ],
        },
      },
    ]);
    expect(f).toMatchObject({ established: 1, establishedGrouped: 1, fresh: 1, freshGrouped: 0 });
  });

  it("holds fresh-slides-group to the charts that really were on an empty slide", () => {
    // 20 rounds. Each has one chart on a named empty slide (it groups in 17 of
    // the 20, so 85%) and one or two charts on a slide the run had already drawn
    // on, keyed `(visible)` on their first batch, which always group. Read raw,
    // the second kind padded the fresh row to 51 of 54 = 94% and the claim held
    // above its 90% line while the real rate sat at 85%.
    const logs = Array.from({ length: 20 }, (_, r) => ({
      trace: {
        entries: [
          ...onNamedSlide("real", 0, r < 17),
          ...sentinelThenNamed("busy-1", 16),
          ...(r < 14 ? sentinelThenNamed("busy-2", 40) : []),
        ],
      },
    }));
    const verdict = claim("fresh-slides-group").check(logs);
    expect(verdict.actual).toMatch(/^17\/20 = 85%/);
    expect(verdict.ok, "the claim reads the real fresh-slide rate, not one the placeholders lift").toBe(false);
  });

  it("says how many charts the claim left out", () => {
    const logs = Array.from({ length: 20 }, () => ({
      trace: { entries: [...onNamedSlide("real", 0), batch("lone", 10, 0, "(visible)"), outcome("lone", true)] },
    }));
    const verdict = claim("fresh-slides-group").check(logs);
    expect(verdict).toMatchObject({ ok: true });
    expect(verdict.actual).toContain("20/20 = 100%");
    expect(verdict.actual, "a smaller denominator should not be a surprise").toContain("20 more not counted");
  });
});

// ── r3-46 ────────────────────────────────────────────────────────────────────
describe("an archived answer is read in today's words (r3-46)", () => {
  const ID = "which-end-a-short-read-drops";
  const round = (build: string, answers: Record<string, string>) => ({
    build,
    hostAnswers: { answers: Object.entries(answers).map(([id, answer]) => ({ id, answer })) },
  });

  it("does not list a rename as a probe that flipped within a build", () => {
    // `all` and `none` became `not-a-short-read` and `none-of-ours` on
    // 2026-09-01, both uninformative. Read raw, two rounds of one build that
    // recorded them were "a probe that flipped", and the gate invited a reader
    // to declare it in UNSTABLE_ANSWERS as host non-determinism.
    const flips = probeFlipsWithinBuild([
      round("55630e7 · 2026-08-31", { [ID]: "all", other: "yes" }),
      round("55630e7 · 2026-08-31", { [ID]: "none", other: "yes" }),
      round("55630e7 · 2026-08-31", { [ID]: "all", other: "yes" }),
    ]);
    expect(flips.flips.map((f: { id: string }) => f.id)).not.toContain(ID);
    expect(flips.slots, "the renamed question has no informative answer to compare, so it is no slot").toBe(1);
    expect(flips.differing).toBe(0);
  });

  it("still reports a real flip, renamed question or not", () => {
    const flips = probeFlipsWithinBuild([
      round("bbbbbbb · x", { q: "yes", [ID]: "all" }),
      round("bbbbbbb · x", { q: "no", [ID]: "none" }),
    ]);
    expect(flips.flips.map((f: { id: string }) => f.id)).toEqual(["q"]);
  });

  it("pools one answer under one row in the pass-position table", () => {
    // 60 samples per round, 30 on pass 1 and 30 later. The old round spells the
    // two answers `all`/`none`; the new one `not-a-short-read`/`none-of-ours`.
    const samples = (a: string, b: string) => [
      ...Array.from({ length: 20 }, () => ({ pass: 1, answer: a })),
      ...Array.from({ length: 10 }, () => ({ pass: 1, answer: b })),
      ...Array.from({ length: 5 }, () => ({ pass: 2, answer: a })),
      ...Array.from({ length: 25 }, () => ({ pass: 2, answer: b })),
    ];
    const logs = [
      { hostAnswers: { answers: [{ id: ID, answer: "all", samples: samples("all", "none") }] } },
      {
        hostAnswers: {
          answers: [{ id: ID, answer: "not-a-short-read", samples: samples("not-a-short-read", "none-of-ours") }],
        },
      },
    ];
    const rows = passPositionBias(logs) as { answer: string; p1: number; later: number }[];
    // Four rows (two spellings of each answer) became two, with the shift the
    // data has: 67% on pass 1 against 17% later, and the mirror image.
    expect(rows.map((r) => r.answer).sort()).toEqual(["none-of-ours", "not-a-short-read"]);
    const lead = rows.find((r) => r.answer === "not-a-short-read");
    expect(Math.round(lead!.p1)).toBe(67);
    expect(Math.round(lead!.later)).toBe(17);
  });
});

// ── r3-48 ────────────────────────────────────────────────────────────────────
describe("a claim reads the whole round number in a file name (r3-48)", () => {
  const queued = { message: "a chart's tag could not even be queued" };
  const round = (name: string, ...entries: object[]) => ({ roundName: name, trace: { entries } });

  it("counts the silence from round 065, not from round 100", () => {
    // `1000-x.json` was cut to `100` by `.slice(0, 3)`, so the newest round read
    // as 100 and the gap since the trace last fired (065) as 35: the claim
    // printed STALE for rounds 1000 to 1059 on a trace that has been silent
    // for a thousand rounds.
    const verdict = (newest: string) =>
      claim("no-queue-trace-is-dead").check([round("065-a.json", queued), round("999-a.json"), round(newest)]);
    expect(verdict("999-b.json")).toMatchObject({ ok: true, actual: "last round 065, 934 rounds ago" });
    expect(verdict("1000-a.json")).toMatchObject({ ok: true, actual: "last round 065, 935 rounds ago" });
    expect(verdict("1059-a.json")).toMatchObject({ ok: true, actual: "last round 065, 994 rounds ago" });
  });

  it("goes STALE when the trace fires again past round 999", () => {
    const stale = claim("no-queue-trace-is-dead").check([round("065-a.json", queued), round("1002-a.json", queued)]);
    expect(stale).toMatchObject({ ok: false, actual: "last round 1002, 0 rounds ago" });
  });

  it("does not count a file that is not a round as round zero", () => {
    // `rounds/predictions.json` lives beside the rounds. Its name has no number,
    // so it can neither be the newest round nor the one the trace fired in.
    const verdict = claim("no-queue-trace-is-dead").check([
      round("065-a.json", queued),
      round("500-a.json"),
      round("predictions.json", queued),
    ]);
    expect(verdict).toMatchObject({ ok: true, actual: "last round 065, 435 rounds ago" });
  });
});

// ── r3-79 ────────────────────────────────────────────────────────────────────
describe("await-deploy names the run by the commit it built (r3-79)", () => {
  // THE 2026-09-27 INCIDENT, as the run list recorded it. A (bbfcf5b) was pushed
  // first and B (dcf21b4) a minute later; CI for B finished first. Both Deploy
  // Pages runs carry B as `headSha` — main's tip when each was created — and the
  // later one published A.
  const A = "bbfcf5b41ca30669418b1db81d1bd6688eb708b0";
  const B = "dcf21b42c0b1e2a76483f0d9f46f910f8e70bc68";
  const at = (hms: string) => `2026-09-27T${hms}Z`;
  type Row = {
    databaseId: number;
    name: string;
    headSha: string;
    event: string;
    status: string;
    conclusion: string | null;
    createdAt: string;
    updatedAt: string;
  };
  const row = (over: Partial<Row>): Row => ({
    databaseId: 1,
    name: "Deploy Pages",
    headSha: B,
    event: "workflow_run",
    status: "completed",
    conclusion: "success",
    createdAt: at("12:17:22"),
    updatedAt: at("12:18:05"),
    ...over,
  });
  const ciA = row({
    databaseId: 10,
    name: "CI",
    headSha: A,
    event: "push",
    createdAt: at("12:13:17"),
    updatedAt: at("12:17:33"),
  });
  const ciB = row({
    databaseId: 11,
    name: "CI",
    headSha: B,
    event: "push",
    createdAt: at("12:14:14"),
    updatedAt: at("12:17:20"),
  });
  const p1 = row({ databaseId: 21, createdAt: at("12:17:22"), updatedAt: at("12:18:05") }); // built B
  const p2 = row({ databaseId: 22, createdAt: at("12:17:35"), updatedAt: at("12:18:20") }); // built A, over B
  const logs: Record<number, string> = {
    21: "build\tUNKNOWN STEP\t2026-09-27T12:17:40Z HEAD is now at dcf21b4 Merge: the next change",
    22: "build\tUNKNOWN STEP\t2026-09-27T12:17:53Z HEAD is now at bbfcf5b Merge: the capture route works",
  };
  const builtOf = (r: Row) => builtShaOf(r, (id: number) => logs[id] ?? null);
  const since = (ci: Row) => Date.parse(ci.updatedAt) - 10_000;

  it("reads the commit a workflow_run build checked out, and trusts headSha only for a hand-started run", () => {
    expect(builtOf(p1), "headSha says B for both runs; the log says what was built").toBe("dcf21b4");
    expect(builtOf(p2)).toBe("bbfcf5b");
    expect(builtShaOf(row({ event: "workflow_dispatch", headSha: A }), () => "nothing to read")).toBe(A);
    expect(builtOf(row({ databaseId: 99 })), "a log that cannot be read says nothing").toBeNull();
  });

  it("does not call B live while the site serves A", () => {
    const rows = [p2, p1, ciB, ciA];
    // The newest run with headSha B is p2, which published A. Matching on that
    // column said "B is live" in 15 of 15 poll phases of this incident.
    const verdict = judgeDeploy({ rows, sha: B, since: since(ciB), builtOf });
    expect(verdict.state).toBe("replaced");
    expect(verdict.note).toContain("bbfcf5b");
    // And A, which p2 did publish, is live.
    expect(judgeDeploy({ rows, sha: A, since: since(ciA), builtOf }).state).toBe("live");
  });

  it("does not count a deploy that started before CI for the commit concluded", () => {
    // CI for A concluded at 12:17:33, so p1 (created 12:17:22) was started by
    // something else; if it counted, A would have two runs and no clear latest.
    const rows = [p2, p1];
    const verdict = judgeDeploy({
      rows,
      sha: A,
      since: since(ciA),
      builtOf: (r: Row) => (r.databaseId === 21 ? null : builtOf(r)),
    });
    expect(verdict.state, "p1 is out of scope, so its unreadable log must not hold the answer").toBe("live");
  });

  it("keeps waiting when only another commit's deploy exists yet", () => {
    // CI(A) finished three seconds before CI(B): the run it started is created
    // after B was pushed, so it also carries B as headSha, and `find` on that
    // column returned it as "B's deploy" the moment it succeeded.
    const early = row({ databaseId: 31, createdAt: at("12:17:19"), updatedAt: at("12:17:50") });
    const verdict = judgeDeploy({
      rows: [early],
      sha: B,
      since: since(ciB),
      builtOf: () => "bbfcf5b",
    });
    expect(verdict).toMatchObject({ state: "waiting", note: "no Deploy Pages run has built it yet" });
  });

  it("holds at waiting while any deploy is in flight, because it may land on top", () => {
    const queued = row({
      databaseId: 41,
      status: "queued",
      conclusion: null,
      createdAt: at("12:19:00"),
      updatedAt: at("12:19:00"),
    });
    const verdict = judgeDeploy({ rows: [queued, p1], sha: B, since: since(ciB), builtOf });
    expect(verdict.state).toBe("waiting");
    expect(verdict.note).toContain("41");
  });

  it("ends in failed when the run that built the commit failed, and ignores skipped and cancelled runs", () => {
    const failed = row({ databaseId: 51, conclusion: "failure" });
    const skipped = row({ databaseId: 52, conclusion: "skipped", createdAt: at("12:17:25") });
    const cancelled = row({ databaseId: 53, conclusion: "cancelled", createdAt: at("12:17:26") });
    const built = (r: Row) => (r.databaseId === 51 ? "dcf21b4" : null);
    const verdict = judgeDeploy({ rows: [skipped, cancelled, failed], sha: B, since: since(ciB), builtOf: built });
    expect(verdict.state).toBe("failed");
    expect(verdict.note).toContain("51");
  });

  it("never reports live on a run whose commit it cannot read", () => {
    const verdict = judgeDeploy({ rows: [p1], sha: B, since: since(ciB), builtOf: () => null });
    expect(verdict.state).toBe("waiting");
    expect(verdict.note).toContain("could not read");
  });

  it("judges CI by its own conclusion", () => {
    expect(judgeCi(null).state).toBe("waiting");
    expect(judgeCi(row({ status: "in_progress", conclusion: null })).state).toBe("waiting");
    expect(judgeCi(ciB).state).toBe("success");
    expect(judgeCi(row({ conclusion: "failure" }))).toMatchObject({ state: "failed", note: "failed: failure" });
  });

  // The whole wait, driven through a fake `gh` that answers `run list` from a
  // script of what GitHub would say poll by poll, and `run view` from `logs`.
  const drive = async (script: Row[][], sha: string, minutes = 30) => {
    let polls = 0;
    let clock = 0;
    const said: string[] = [];
    const warned: string[] = [];
    const code = await awaitDeploy({
      sha,
      minutes,
      gh: (args: string[]) => {
        if (args[0] === "run" && args[1] === "list")
          return JSON.stringify(script[Math.min(polls++, script.length - 1)]);
        if (args[0] === "run" && args[1] === "view") return logs[Number(args[2])] ?? "";
        throw new Error(`unexpected gh ${args.join(" ")}`);
      },
      sleep: async () => {
        clock += 15_000;
      },
      now: () => clock,
      say: (m: string) => said.push(String(m)),
      warn: (m: string) => warned.push(String(m)),
      tick: () => {},
    });
    return { code, said: said.join("\n"), warned: warned.join("\n") };
  };

  it("exits 0 only when the most recent deploy built the commit", async () => {
    const running = row({ databaseId: 21, status: "in_progress", conclusion: null });
    const { code, said } = await drive([[], [ciB], [running, ciB], [p1, ciB]], B);
    expect(code).toBe(0);
    expect(said).toContain("dcf21b4 is live");
  });

  it("exits 2 when B is replaced, and says how to settle it", async () => {
    const { code, warned } = await drive([[p2, p1, ciB, ciA]], B);
    expect(code).toBe(2);
    expect(warned).toContain("replaced");
    expect(warned).toContain("gh workflow run");
  });

  it("exits 1 on a red CI and does not claim the commit may be live", async () => {
    const { code, warned } = await drive([[row({ ...ciB, conclusion: "failure" })]], B);
    expect(code).toBe(1);
    // This message said "It may well be LIVE — deploys do not wait for CI" for
    // a month after `pages.yml` began to wait for it.
    expect(warned).not.toMatch(/may well be LIVE|do not wait for CI/);
    expect(warned).toContain("never deployed");
  });

  it("exits 3 when CI never appears, and when the deploy never does", async () => {
    expect((await drive([[]], B, 1)).code).toBe(3);
    expect((await drive([[ciB]], B, 1)).code).toBe(3);
  });
});

// ── r3-77 ────────────────────────────────────────────────────────────────────
describe("the skill zip is the same bytes on every build (r3-77)", () => {
  const made: string[] = [];
  afterEach(() => {
    vi.useRealTimers();
    for (const dir of made.splice(0)) rmSync(dir, { recursive: true, force: true });
  });

  /** A package laid out like `skill-dist/ssf-charts`, under a scratch directory. */
  const packageDir = () => {
    const base = mkdtempSync(join(tmpdir(), "pc-skill-"));
    made.push(base);
    const root = join(base, "ssf-charts");
    for (const sub of ["lib", "scripts", "examples"]) mkdirSync(join(root, sub), { recursive: true });
    writeFileSync(join(root, "SKILL.md"), "# skill\n");
    writeFileSync(join(root, "lib", "ssf-charts.js"), "export const x = 1;\n");
    writeFileSync(join(root, "scripts", "render-svg.mjs"), "import '../lib/ssf-charts.js';\n");
    writeFileSync(join(root, "examples", "charts.json"), "{}\n");
    return root;
  };

  /**
   * Every entry's name and RAW DOS timestamp, straight from the central
   * directory, so no reader library decides how a date is shown. The end record
   * has no comment, so it is the last 22 bytes.
   */
  const entries = (zip: Buffer) => {
    const end = zip.length - 22;
    expect(zip.readUInt32LE(end), "end of central directory").toBe(0x06054b50);
    const count = zip.readUInt16LE(end + 10);
    let at = zip.readUInt32LE(end + 16);
    const out: { name: string; time: number; date: number }[] = [];
    for (let i = 0; i < count; i++) {
      const nameLen = zip.readUInt16LE(at + 28);
      const extraLen = zip.readUInt16LE(at + 30);
      const commentLen = zip.readUInt16LE(at + 32);
      out.push({
        name: zip.toString("utf8", at + 46, at + 46 + nameLen),
        time: zip.readUInt16LE(at + 12),
        date: zip.readUInt16LE(at + 14),
      });
      at += 46 + nameLen + extraLen + commentLen;
    }
    return out;
  };
  const sha = (b: Buffer) => createHash("sha256").update(b).digest("hex");

  it("dates every entry 1980-01-01, the first day a zip can hold", async () => {
    const zip: Buffer = await zipPackage(packageDir());
    const found = entries(zip);
    expect(found.map((e) => e.name)).toEqual([
      "ssf-charts/",
      "ssf-charts/examples/",
      "ssf-charts/lib/",
      "ssf-charts/scripts/",
      "ssf-charts/SKILL.md",
      "ssf-charts/examples/charts.json",
      "ssf-charts/lib/ssf-charts.js",
      "ssf-charts/scripts/render-svg.mjs",
    ]);
    // 0x0021 is year 0 (1980), month 1, day 1. `new Date(0)` wrote a year of -10
    // into a 7-bit field, which is 0xec21: 2098-01-01.
    for (const e of found)
      expect({ name: e.name, date: e.date, time: e.time }).toEqual({ name: e.name, date: 0x0021, time: 0 });
  });

  it("does not depend on the clock", async () => {
    const root = packageDir();
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(new Date("2026-10-25T13:09:08Z"));
    const first: Buffer = await zipPackage(root);
    // The folder entries used to take the wall clock, so two builds seconds
    // apart differed from byte 11 on.
    vi.setSystemTime(new Date("2031-02-03T04:05:06Z"));
    const second: Buffer = await zipPackage(root);
    expect(sha(second)).toBe(sha(first));
  });

  it("still unpacks to the same files under ssf-charts/", async () => {
    const loaded = await JSZip.loadAsync(await zipPackage(packageDir()));
    expect(Object.keys(loaded.files).sort()).toEqual(
      [
        "ssf-charts/",
        "ssf-charts/SKILL.md",
        "ssf-charts/examples/",
        "ssf-charts/examples/charts.json",
        "ssf-charts/lib/",
        "ssf-charts/lib/ssf-charts.js",
        "ssf-charts/scripts/",
        "ssf-charts/scripts/render-svg.mjs",
      ].sort(),
    );
    expect(await loaded.file("ssf-charts/lib/ssf-charts.js")!.async("string")).toBe("export const x = 1;\n");
  });
});
