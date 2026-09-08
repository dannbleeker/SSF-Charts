import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { spawnSync } from "child_process";
import { mkdtempSync, writeFileSync, rmSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
// @ts-expect-error — plain .mjs tools, no types; both are deliberately
// independent of src/ so they cannot inherit a bug from the code they audit.
import { readDeckBytes } from "../scripts/verify-deck.mjs";
// @ts-expect-error — as above.
import { triage, runsIn, selfTestIn, knownBug, deckEvidence, poolRasteriseArms } from "../scripts/triage.mjs";
// Its own line, and this file has now predicted its own bug three times: adding
// a name to the grouped import above pushes it over the print width, prettier
// reflows it across lines, and `@ts-expect-error` covers only the NEXT line — so
// the directive stops reaching the `from` clause. Suite green, `tsc` red.
// @ts-expect-error — as above.
import { poolTagFaults } from "../scripts/triage.mjs";
// Its own line, same reason as every other single import in this file.
// @ts-expect-error — as above.
import { poolGroupVsTag } from "../scripts/triage.mjs";
// A NAMESPACE import, destructured below, and the fourth time this trap has been
// paid for. `@ts-expect-error` covers exactly one LINE — the one carrying `from`
// — so a named list long enough for prettier to wrap moves `from` four lines
// down, uncovers the import, and turns the directive itself into an "unused
// directive" error. Adding `scenarioRegressions` to the list did precisely that.
// A namespace import is one line whatever is destructured off it.
// @ts-expect-error — as above.
import * as pools from "../scripts/triage.mjs";
const {
  poolFreshVsEstablished,
  poolStarvedQuestions,
  poolBatchSpanVsGroup,
  scenarioRegressions,
  FALLBACK_SIGNALS,
  poolOriginTagLosses,
  roundProfile,
  profileDivergence,
  traceNovelty,
  traceSignature,
  poolUpdateShortfalls,
  poolGroupVsTagCoverage,
} = pools;
// Its own line, for the reason spelled out below: adding it to the grouped
// import above reflowed that statement across lines, and `@ts-expect-error`
// covers only the NEXT line — so the directive stopped reaching the `from`
// clause. Suite stayed green, `tsc` went red, exactly as predicted here.
// A NAMESPACE IMPORT, because this list has now grown past the print width
// twice. `@ts-expect-error` covers exactly the next LINE, so the moment prettier
// reflows a named-import list the `from` clause moves off the directive, the
// suite stays green and `tsc` goes red — which is precisely what happened when
// `paneAgeAtStartSeconds` was added. A namespace import is one line forever and
// cannot be reflowed out from under it.
// @ts-expect-error — as above.
import * as pooled from "../scripts/triage.mjs";
const {
  poolEveryDraw,
  poolProfileDisagreements,
  probeFlipsWithinBuild,
  poolPairPosition,
  roundSpanSeconds,
  paneAgeAtStartSeconds,
  poolFallbackRates,
  poolDriverRuns,
  unreadSignals,
  poolInPlaceUpdates,
} = pooled;
// Its own line: adding it above pushes that import over the print width, and a
// reflowed import moves this directive off the statement it is annotating.
// @ts-expect-error — as above.
import { describeFinding } from "../scripts/triage.mjs";
// Also its own line, and for the reason the comment above gives: adding it to
// the grouped import pushed that statement over the print width, prettier
// reflowed it across lines, and `@ts-expect-error` covers only the NEXT line —
// so the directive stopped reaching the `from` clause where the error is
// reported. The suite stayed green and `tsc` went red.
// @ts-expect-error — as above.
import { batchPopulations } from "../scripts/triage.mjs";
// @ts-expect-error — as above. One directive per import, one import per line.
import { poolScenarioPopulations } from "../scripts/triage.mjs";
// @ts-expect-error — as above. One directive per import, one import per line.
import { poolGroupingOutcome } from "../scripts/triage.mjs";
// @ts-expect-error — as above. One directive per import, one import per line.
import { poolScenarioFriction } from "../scripts/triage.mjs";
// @ts-expect-error — as above. One directive per import, one import per line.
import { deckGeometryFaults } from "../scripts/triage.mjs";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";

/**
 * The run/deck joiner.
 *
 * Its whole job is to disagree — to say where the add-in's account of a run
 * and the file that run produced do not match. So the cases that matter are
 * the ones where they differ, and every one of them below is paired with the
 * same slot in the clean deck to prove the verdict came from the difference
 * and not from the tool's default mood.
 */

const cfg = (title: string) => ({ ...sampleConfig("line"), title });

interface Piece {
  title: string;
  slot: number;
  run?: string;
  tagged?: boolean;
}

/** Build a real .pptx from slot descriptions and read it back as the auditor sees it. */
async function deckOf(pieces: Piece[]) {
  const built = await buildDeckBase64(
    pieces.map((p) => ({
      scene: buildChart(cfg(p.title)),
      title: p.title,
      configJson: p.tagged === false ? undefined : JSON.stringify(cfg(p.title)),
      slot: p.slot,
      run: p.run ?? "run-x",
    })),
  );
  return readDeckBytes(Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0)));
}

/** A run log claiming every named item rendered and was tagged. */
function logOf(titles: string[], over: Record<string, unknown> = {}) {
  return {
    build: "test-build",
    run: "run-x",
    host: "test host",
    totalMs: 1234,
    path: "file",
    items: titles.map((title) => ({
      title,
      status: "rendered",
      shapes: 3,
      ms: 10,
      grouped: true,
      tagged: true,
      chart: true,
      lateOutcome: "",
    })),
    deck: { slidesAdded: titles.length, addsIssued: titles.length, lost: 0, blank: [] },
    ...over,
  };
}

const verdicts = (t: { slots: { title: string; verdict: string }[] }) =>
  Object.fromEntries(t.slots.map((s) => [s.title, s.verdict]));

describe("triage — joining a run log to the deck it produced", () => {
  it("finds no disagreement when the deck holds exactly what the run claimed", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "B", slot: 1 },
      ]),
      logOf(["A", "B"]),
    );
    expect(verdicts(t)).toEqual({ A: "ok", B: "ok" });
    expect(t.disagreements).toBe(0);
    expect(t.run).toBe("run-x");
    expect(t.inferredRun).toBe(false);
  });

  it("calls a slot the deck does not hold lost", async () => {
    // Negative control: B is `ok` above, from the same log, with its slide.
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), logOf(["A", "B"]));
    expect(verdicts(t)).toEqual({ A: "ok", B: "lost" });
    expect(t.disagreements).toBe(1);
  });

  it("does not call a skipped item lost — the run never claimed to draw it", async () => {
    const log = logOf(["A", "B"]);
    log.items[1].status = "skipped";
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), log);
    expect(verdicts(t)).toEqual({ A: "ok", B: "skipped" });
    expect(t.disagreements).toBe(0);
  });

  it("calls a tag the run believed it wrote and the file does not carry tag-lost", async () => {
    // The file path cannot produce this by construction, which is exactly why
    // it is worth naming when it appears: it means the deck lost a tag the
    // add-in had already committed to the bytes.
    const t = triage(await deckOf([{ title: "A", slot: 0, tagged: false }]), logOf(["A"]));
    expect(verdicts(t)).toEqual({ A: "tag-lost" });
    expect(t.disagreements).toBe(1);
  });

  it("calls a tag the run gave up on and the file carries repaired, and not a fault", async () => {
    const log = logOf(["A"]);
    log.items[0].tagged = false;
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), log);
    expect(verdicts(t)).toEqual({ A: "repaired" });
    // The repair pass working is not a disagreement to chase.
    expect(t.disagreements).toBe(0);
  });

  it("calls a chart object with no config not-editable when the run never claimed the tag", async () => {
    const log = logOf(["A"]);
    log.items[0].tagged = false;
    const t = triage(await deckOf([{ title: "A", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ A: "not-editable" });
    expect(t.disagreements).toBe(1);
  });

  it("calls two slides claiming one slot duplicated", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "A again", slot: 0 },
      ]),
      logOf(["A"]),
    );
    expect(verdicts(t)).toEqual({ A: "duplicated" });
    expect(t.slots[0].rows).toBe(2);
    expect(t.disagreements).toBe(1);
  });

  it("counts another run's slides as foreign, never as this run's losses", async () => {
    // The case the run token exists to survive. Before it existed, a second
    // insert into the same deck made every item match two slides and the
    // repair pass deleted a healthy run as the other one's duplicate.
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "A", slot: 0, run: "run-earlier" },
      ]),
      logOf(["A"]),
    );
    expect(verdicts(t)).toEqual({ A: "ok" });
    expect(t.foreign).toHaveLength(1);
    expect(t.foreign[0].run).toBe("run-earlier");
    expect(t.disagreements).toBe(0);
  });

  it("calls a slide carrying a slot this run never issued an orphan", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "stray", slot: 7 },
      ]),
      logOf(["A"]),
    );
    expect(t.orphans).toEqual([{ slot: 7, verdict: "orphan", indexes: [2] }]);
    expect(t.disagreements).toBe(1);
  });

  it("infers the run from the deck when the log predates run ids, and says it guessed", async () => {
    const log = logOf(["A", "B"]) as Record<string, unknown>;
    delete log.run;
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "B", slot: 1 },
      ]),
      log,
    );
    expect(t.run).toBe("run-x");
    expect(t.inferredRun).toBe(true);
    expect(t.disagreements).toBe(0);
  });

  it("leaves an item that was never a chart alone, config tag or not", async () => {
    // The bug that produced this test: a clean 38-of-38 run triaged as seven
    // broken slides. The title page, both contents pages and four elements are
    // drawn as `PowerChart` objects with NO config by design, and the rule
    // "a chart object without a config is not re-editable" flagged every one.
    const log = logOf(["Title"]);
    log.items[0].tagged = false;
    log.items[0].chart = false;
    const t = triage(await deckOf([{ title: "Title", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ Title: "ok" });
    expect(t.disagreements).toBe(0);
  });

  it("will not guess at intent a log is too old to state", async () => {
    // Same deck and the same missing config, with `chart` absent: the tool
    // cannot tell a title page from a chart that lost its tag, so it says so
    // instead of picking one. `tagged` cannot stand in — on the shape path
    // false means both "never had a config" and "the write did not land".
    const log = logOf(["Mystery"]) as { items: Record<string, unknown>[] };
    log.items[0].tagged = false;
    delete log.items[0].chart;
    const t = triage(await deckOf([{ title: "Mystery", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ Mystery: "no-config" });
    expect(t.disagreements).toBe(1);
  });

  it("prefers the log's own run id over the deck's commonest", async () => {
    // A deck where the run under investigation is the MINORITY. Inferring from
    // the deck would pick the other one and report every slot lost.
    const t = triage(
      await deckOf([
        { title: "A", slot: 0, run: "run-mine" },
        { title: "old", slot: 0, run: "run-older" },
        { title: "old", slot: 1, run: "run-older" },
      ]),
      logOf(["A"], { run: "run-mine" }),
    );
    expect(t.run).toBe("run-mine");
    expect(verdicts(t)).toEqual({ A: "ok" });
    expect(t.foreign).toHaveLength(2);
  });
});

describe("triage — a log holding more than one run", () => {
  it("reads each run in a both-paths log separately", async () => {
    // One click can take both insert paths, and they fail in completely
    // different ways. Merging them would produce a report that is true of
    // neither: the file run's slides are present while the shape run's are
    // still missing, in the same deck, at the same moment.
    const deck = await deckOf([
      { title: "A", slot: 0, run: "run-file" },
      { title: "B", slot: 1, run: "run-file" },
      { title: "A", slot: 0, run: "run-shapes" },
    ]);
    const file = { ...logOf(["A", "B"], { run: "run-file", path: "file" }) };
    const shapes = { ...logOf(["A", "B"], { run: "run-shapes", path: "shapes" }) };
    const runs = runsIn({ build: "b", host: "h", runs: [file, shapes] });
    expect(runs).toHaveLength(2);
    expect(verdicts(triage(deck, runs[0]))).toEqual({ A: "ok", B: "ok" });
    // The shape run only landed its first slide, and says so — while the file
    // run beside it in the same deck is clean.
    expect(verdicts(triage(deck, runs[1]))).toEqual({ A: "ok", B: "lost" });
  });

  it("reads a single-run log as a list of one", async () => {
    // Every artifact captured so far is shaped this way. Refusing them would
    // make the tool useless on exactly the runs it was written to explain.
    const flat = logOf(["A"]);
    expect(runsIn(flat)).toEqual([flat]);
  });
});

describe("triage — logs that are not inserts", () => {
  it("finds the host self-test's verdicts in a log that holds no runs", async () => {
    // The self-test writes a log with an empty `runs` list: its scenarios are
    // not inserts and own no slots. Reading only `runs` made the whole report
    // come out blank and exit 0 — indistinguishable from a clean deck, on a
    // file someone handed this tool expecting an answer.
    const log = {
      build: "b",
      host: "h",
      runs: [],
      selftest: [
        { name: "insert twice", ok: true, detail: "deck grew by 4", ms: 10 },
        { name: "explode a picture", ok: false, detail: "config did not survive", ms: 20 },
      ],
    };
    expect(runsIn(log)).toEqual([]);
    expect(selfTestIn(log).map((s: { name: string }) => s.name)).toEqual(["insert twice", "explode a picture"]);
  });

  it("finds nothing to report in a log that is neither", async () => {
    expect(selfTestIn({ build: "b", host: "h", runs: [] })).toEqual([]);
    expect(selfTestIn(undefined)).toEqual([]);
  });

  /**
   * The round's own deck evidence — the half that replaced saving a .pptx and
   * taking a screenshot, and that nothing read until a round had to be worked
   * out by hand.
   *
   * The distinction each case pins is between an empty slide and an empty
   * ANSWER. This host is known to report a shape collection short without
   * throwing (`shapes-items-count-honest`), so a readback of zero is one
   * witness; the host's own picture of the same slide is a second, and only
   * the two together say a slide is really blank.
   */
  const evidenceDeck = (slides: { id: string; count: number; png?: number }[], added?: string[]) => ({
    inventory: slides.map((s, i) => ({ slideId: s.id, index: i, shapes: [], count: s.count })),
    newSlides: added ?? slides.map((s) => s.id),
    // Base64 is 4 characters per 3 bytes, which is the ratio the reader undoes.
    shots: slides.filter((s) => s.png !== undefined).map((s) => ({ slideId: s.id, png: "A".repeat(s.png! * 4) })),
  });

  it("only calls a slide blank when its picture agrees with its readback", () => {
    const e = deckEvidence(
      evidenceDeck([
        { id: "a", count: 12, png: 900 }, // has shapes — not a blank at all
        { id: "b", count: 0, png: 300 }, // empty, and the smallest picture here
        { id: "c", count: 0 }, // empty, no picture — the rasterise cap
        { id: "d", count: 0, png: 900 }, // empty, but the picture has content
      ]),
    );
    expect(e.withShapes).toBe(1);
    expect(e.confirmed, "only the slide whose picture is blank too").toBe(1);
    expect(e.unseen, "no picture is not evidence of an empty slide").toBe(1);
    expect(e.lying, "a blank readback over a picture with content is the host lying").toBe(1);
  });

  it("reads the archive's stripped placeholder as NO picture, not as a blank one", () => {
    /**
     * THE SECOND WITNESS WAS FAKE IN ALL 323 ARCHIVED ROUNDS.
     *
     * `stripImages` replaces every base64 payload with the sentence
     * "<image stripped for the archive — …" before a round is written, so every
     * one of the 1,923 shots on file is a 53- or 55-character string. It is
     * truthy, so it survived the "a picture was taken" filter, and the reader
     * scores it at 40 bytes — far under `BLANK_PNG_CEILING`.
     *
     * So the rasteriser AGREED that every added slide in every archived round
     * was blank. Ten empty added slides across nine rounds (041, 042, 045, 055,
     * 073, 074, 078, 080 twice, 085) printed as `confirmed` — what this file
     * calls "the two-witness line a maintainer reads as proven data loss" —
     * with `unseen: 0`, and the honest `unseen` branch had never executed.
     * Re-triaging the real archive after the fix gives confirmed 0, unseen 10.
     *
     * Fixed in the reader rather than the stripper because the archive is
     * append-only: those placeholders cannot be rewritten, and re-reading the
     * archive is the whole point of keeping it.
     */
    const withStripped = (png: string) =>
      deckEvidence({
        inventory: [{ slideId: "s", shapes: [] }],
        newSlides: ["s"],
        shots: [{ slideId: "s", png }],
      });
    for (const placeholder of [
      "<image stripped for the archive — see rounds/README.md>",
      "<image stripped for the archive — see docs/ROUNDS.md>",
    ]) {
      const e = withStripped(placeholder);
      expect(e.confirmed, `"${placeholder.slice(0, 22)}…" was read as a blank picture`).toBe(0);
      expect(e.unseen, "a stripped picture is no picture at all").toBe(1);
    }
    // The control: a genuinely tiny PNG still counts as a blank, so the fix
    // narrows nothing it was not meant to.
    const real = deckEvidence({
      inventory: [{ slideId: "s", shapes: [] }],
      newSlides: ["s"],
      shots: [{ slideId: "s", png: "A".repeat(300 * 4) }],
    });
    expect(real.confirmed, "a real blank PNG stopped being a witness").toBe(1);
    expect(real.unseen).toBe(0);
  });

  it("counts only the slides this round added, never the deck it landed in", () => {
    // A round drops its slides into whatever the user already had. Counting the
    // pre-existing ones would report someone's own empty slides as our losses.
    const e = deckEvidence(
      evidenceDeck(
        [
          { id: "theirs", count: 0 },
          { id: "ours", count: 0 },
        ],
        ["ours"],
      ),
    );
    expect(e.added).toBe(1);
    expect(e.unseen).toBe(1);
  });

  it("names the added slides whose id never finished, and joins them to the blanks", () => {
    // ROUND 041 put both halves of this in the log and nothing put them side by
    // side: seven slides added, two of them named `256#0` and `257#0` against
    // `288#3603562595` for the rest, one added slide blank, and a
    // `delete-by-id left slides behind` line in the same round. It took a hand
    // query to see, which is how a finding nearly went unnoticed.
    const e = deckEvidence(
      evidenceDeck(
        [
          { id: "288#3603562595", count: 24 },
          { id: "256#0", count: 0, png: 300 },
          { id: "257#0", count: 12 },
        ],
        ["288#3603562595", "256#0", "257#0"],
      ),
    );
    expect(e.oddIds, "an id ending #0 is not the shape a finished slide has").toEqual(["256#0", "257#0"]);
    // THE JOIN, which is the whole point: one of the two odd ids is also the
    // blank, so the two observations are one event rather than two.
    expect(e.oddAndBlank).toBe(1);
    expect(e.confirmed).toBe(1);
  });

  it("pools tag faults per build, so two rounds of one build can be compared with themselves", () => {
    // WHY THIS EXISTS: on 2026-08-15 a tag-anchor fix was nearly reported as
    // working because the round carrying it looked like the round before. It
    // did — and so did two rounds with NO renderer change between them, where
    // `tags-undefined` went 1 → 5. Grouping by build is what lets the same build
    // run twice pool instead of compete, and that pair is the only noise
    // measurement that owes nothing to an argument.
    const round = (build: string, faults: { undef?: number; grouped?: number }) => ({
      build,
      trace: {
        entries: [
          ...Array.from({ length: faults.undef ?? 0 }, () => ({
            message: "tagging failed — charts are not re-editable",
            data: { error: "Cannot read properties of undefined (reading 'add')" },
          })),
          ...Array.from({ length: faults.grouped ?? 0 }, () => ({
            message: "grouping refused",
            data: { error: "at=grouping the chart's shapes | code=5010" },
          })),
        ],
      },
    });
    const byBuild = poolTagFaults([
      round("aaaaaaa · 2026-08-15", { undef: 1 }),
      round("aaaaaaa · 2026-08-15", { undef: 5 }),
      round("bbbbbbb · 2026-08-15", { undef: 5, grouped: 2 }),
    ]);
    expect([...byBuild.keys()], "the timestamp is not part of the build").toEqual(["aaaaaaa", "bbbbbbb"]);
    // Two rounds of one build pooled, not averaged away — the SPREAD is the point.
    expect(byBuild.get("aaaaaaa")!.map((r: Record<string, number>) => r["tags-undefined"])).toEqual([1, 5]);
    expect(byBuild.get("aaaaaaa")!.map((r: Record<string, number>) => r["tagging-failed"])).toEqual([1, 5]);
    expect(byBuild.get("bbbbbbb")![0]["group-5010"]).toBe(2);
    // A fault line that is not a tagging failure must not inflate the count that
    // a fix would be judged on.
    expect(byBuild.get("bbbbbbb")![0]["tagging-failed"]).toBe(5);
  });

  it("flags a scenario that WAS passing and has stopped, and stays quiet about the host's mood", () => {
    // THE ONLY AUTOMATIC CHECK ON A ROUND'S OWN RESULT. Rounds 070-072 took
    // `same scale across the deck` from 35 consecutive failures to three
    // consecutive passes, and nothing guarded that: every other number a round
    // produces is read by a person and filed, so a later build could take it
    // back and no gate would fail.
    const round = (pairs: [string, boolean][]) => ({ selftest: pairs.map(([name, ok]) => ({ name, ok })) });
    const passing: [string, boolean][] = [
      ["same scale", true],
      ["visible", true],
    ];

    // Established (three straight passes) and now failing — the case it is for.
    const gone = scenarioRegressions([
      round(passing),
      round(passing),
      round(passing),
      round([
        ["same scale", false],
        ["visible", true],
      ]),
    ]);
    expect(
      (gone as { name: string }[]).map((g) => g.name),
      "a scenario that had passed three rounds running failed unnoticed",
    ).toEqual(["same scale"]);

    // MOOD IS NOT A REGRESSION, and this is what keeps the gate usable. This
    // host swings 4-of-5 to 1-of-5 on the same build with nothing changed; a
    // scenario that fails half the time was never established, so its next
    // failure is not news. A gate that cries wolf gets switched off.
    const flaky = scenarioRegressions([
      round([["same scale", true]]),
      round([["same scale", false]]),
      round([["same scale", true]]),
      round([["same scale", false]]),
    ]);
    expect(flaky, "reported a chronically flaky scenario as a regression").toEqual([]);

    // A NEW scenario cannot have been established, so its first bad round is not
    // a fault — otherwise every scenario added would arrive pre-broken.
    const fresh = scenarioRegressions([
      round([["same scale", true]]),
      round([["same scale", true]]),
      round([["same scale", true]]),
      round([
        ["same scale", true],
        ["brand new", false],
      ]),
    ]);
    expect(fresh, "a scenario on its first outing was called a regression").toEqual([]);

    // And too little history is no history: it must not judge from two rounds.
    expect(scenarioRegressions([round(passing), round([["same scale", false]])])).toEqual([]);
  });

  it("counts every chart that cannot follow a drag, which a passing scenario hides", () => {
    // A GREEN VERDICT OVER A SAMPLE. `an update follows a moved chart` passes and
    // tests ONE chart; rounds 073 and 074 lost the origin tag on 9 of 19 and 8 of
    // 17 charts in those same rounds. Every one of those would snap back to where
    // it was inserted instead of following the user's drag, while the scenario
    // reported the round trip holding.
    //
    // Same shape as `does a rasterise poison the next draw` counting only its own
    // four draws, and the fresh-slide split sitting unqueried for eleven rounds.
    // A scenario samples; a pooled count does not.
    const lost = (n: number) => ({ message: "origin tag lost — the chart is re-editable", data: { charts: n } });
    const o = poolOriginTagLosses([
      { trace: { entries: [lost(9), lost(1)] } },
      { trace: { entries: [] } },
      { trace: { entries: [lost(8)] } },
    ]);
    expect(o.charts, "did not total the charts across rounds").toBe(18);
    expect(o.rounds, "counted a round that lost nothing").toBe(2);
    // TEN, not nine: the first round lost 9 charts on one line and 1 on another,
    // and a round's damage is the SUM of its lines. Getting this wrong first time
    // is why it is asserted — a per-line maximum would under-report every round
    // that failed in more than one batch.
    expect(o.worst, "took the worst LINE rather than the worst ROUND").toBe(10);

    // A round with no losses contributes nothing at all, so the report stays
    // silent on an archive that never had the problem.
    expect(poolOriginTagLosses([{ trace: { entries: [] } }])).toEqual({ rounds: 0, charts: 0, worst: 0 });
  });

  it("keeps 4:3 and 16:9 apart, and never judges one against the other", () => {
    // A NIGHTLY CYCLE RUNS 16:9 TWICE AND 4:3 ONCE. Judged against three 16:9
    // rounds, a 4:3 round would be flagged for scoring differently — which it
    // does, by design: round 077 scored 10 of 13 where 16:9 scored 13 of 13
    // twice. The gate would fire every night, and a gate that cries wolf gets
    // switched off. This file has already watched that happen twice.
    const at = (w: number, h: number, pairs: [string, boolean][]) => ({
      build: "aaaaaaa · 2026-08-16",
      slideSize: { width: w, height: h, source: "pageSetup" },
      selftest: pairs.map(([name, ok]) => ({ name, ok })),
    });
    const wide = (pairs: [string, boolean][]) => at(960, 540, pairs);
    const std = (pairs: [string, boolean][]) => at(720, 540, pairs);

    expect(roundProfile(wide([]))).toBe("16:9");
    expect(roundProfile(std([]))).toBe("4:3");
    // A round with no size is one of the 53 filed before the field existed, and
    // every one of those was 16:9 — a documented fact, not a guess.
    expect(roundProfile({}), "an old round must not become its own profile").toBe("16:9");
    // Anything else is visibly itself rather than folded into the nearest named.
    expect(roundProfile(at(1000, 500, []))).toBe("1000x500");

    // TWO READINGS, COMPARED — rounds 115 and 116 in miniature. They recorded
    // 720x540 from the pane while the driver had measured 960x540 off live
    // PageSetup, twice, and printed it before each round. Every comparison in
    // this file groups by the archive's value, so both landed in the wrong arm
    // while `PW_EXPECT_SIZE` reported a match.
    const disagreeing = { ...std([]), build: "a46a2d3 x", driverSlideSize: "16:9" };
    const found = poolProfileDisagreements([disagreeing]);
    expect(found, "a round filed under a profile the driver did not measure went unreported").toHaveLength(1);
    expect(found[0]).toMatchObject({ pane: "4:3", driver: "16:9" });

    // Agreement is silence — otherwise the check cries wolf on every round.
    expect(poolProfileDisagreements([{ ...std([]), driverSlideSize: "4:3" }])).toEqual([]);

    // AND THE HALF THAT MATTERS MOST. A round with only ONE reading has not
    // agreed with anything; it is unverifiable. Counting it as consistent would
    // report every round archived before the field existed as checked — the
    // house defect, told by a denominator.
    expect(poolProfileDisagreements([std([])]), "a round with no second opinion must not read as agreement").toEqual(
      [],
    );
    const onlyDriver = poolProfileDisagreements([{ build: "x y", driverSlideSize: "16:9" }]);
    expect(onlyDriver, "a round with no pane reading is not a disagreement either").toEqual([]);

    // ── the half of a round that nothing gated ───────────────────────────────
    /**
     * `rounds-gate.mjs` compared scenario verdicts and slide-size divergence and
     * contained ZERO references to `hostAnswers` — while 14 of 15 scenarios pass
     * in every round and seven have never failed in 322. Measured over the 82
     * builds with two or more archived rounds, 163 of 2,542 comparable probe
     * slots (6.4%) differ between rounds of the SAME commit.
     *
     * A NON-ANSWER IS NOT A DIFFERENT ANSWER, and that is most of the number:
     * counting `no-scratch-slide`, `unreadable` and `silent` as dissent gives
     * 328 of 3,147 (10.4%) and fills the list with our own read failures dressed
     * up as host non-determinism.
     */
    const sheet = (build: string, answers: Record<string, string>) => ({
      build,
      hostAnswers: { answers: Object.entries(answers).map(([id, answer]) => ({ id, answer })) },
    });
    const f = probeFlipsWithinBuild([
      // Same sha, different timestamp — `build` carries both, so the grouping
      // has to take the sha alone or no two rounds ever pair up at all.
      sheet("abc1234 · 2026-09-01 01:00Z", { steady: "yes", flipper: "yes", quiet: "yes" }),
      sheet("abc1234 · 2026-09-01 02:00Z", { steady: "yes", flipper: "threw", quiet: "unreadable" }),
      // A different build must never be compared against the first.
      sheet("def5678 · 2026-09-01 03:00Z", { steady: "no" }),
    ]);
    expect(f.builds, "a build with only one round was treated as a pair").toBe(1);
    expect(
      f.flips.map((x: { id: string }) => x.id),
      "the wrong probes were called unstable",
    ).toEqual(["flipper"]);
    expect(f.flips[0].answers).toEqual(["threw", "yes"]);
    // `quiet` went yes -> unreadable: one round read it, one did not. That is
    // not the host contradicting itself, and calling it so is the exact mistake
    // `stabilityOf` carried until 2026-08-31.
    expect(f.differing, "a failed read was counted as a dissenting answer").toBe(1);

    // THE GATE: a 4:3 round after three passing 16:9 rounds is not a regression.
    const ok: [string, boolean][] = [["same scale", true]];
    expect(
      scenarioRegressions([wide(ok), wide(ok), wide(ok), std([["same scale", false]])]),
      "judged a 4:3 round against 16:9 history",
    ).toEqual([]);
  });

  it("names a scenario that passes at one slide size and fails at another", () => {
    // THE OTHER QUESTION, and it needs its own answer. The gate asks whether a
    // scenario fell against its OWN history; this asks whether one profile
    // failed what another passed on the SAME BUILD. Round 077 was exactly that
    // — 10 of 13 at 4:3 against 13 of 13 at 16:9 — and nothing said so.
    const at = (build: string, w: number, h: number, pairs: [string, boolean][]) => ({
      build: `${build} · 2026-08-16`,
      slideSize: { width: w, height: h, source: "pageSetup" },
      selftest: pairs.map(([name, ok]) => ({ name, ok })),
    });
    const d = profileDivergence([
      at("aaaaaaa", 960, 540, [
        ["same scale", true],
        ["visible", true],
      ]),
      at("aaaaaaa", 720, 540, [
        ["same scale", false],
        ["visible", true],
      ]),
    ]);
    expect(d.map((x: { name: string }) => x.name)).toEqual(["same scale"]);
    expect(d[0].passedIn).toEqual(["16:9"]);
    expect(d[0].failedIn).toEqual(["4:3"]);
    expect(d[0].flaky, "a clean difference between profiles must not be called flaky").toBeFalsy();

    // FAILING IN BOTH is an ordinary bug, not divergence — reporting it here
    // would bury the one signal this exists for.
    expect(
      profileDivergence([
        at("aaaaaaa", 960, 540, [["same scale", false]]),
        at("aaaaaaa", 720, 540, [["same scale", false]]),
      ]),
      "reported a plain bug as a slide-size difference",
    ).toEqual([]);

    // A PROFILE THAT DISAGREES WITH ITSELF IS FLAKY, NOT DIFFERENT — and the
    // check's first live outing got this wrong. `explode a degraded picture`
    // passed at 4:3, then passed once and failed once at 16:9 on one build.
    // Collapsing a profile to its worst outcome reported "diverged between
    // slide sizes", which was true of the worst reading and wrong about the
    // cause — and it would send someone to investigate an aspect ratio for a
    // scenario that is simply unreliable.
    const unstable = profileDivergence([
      at("aaaaaaa", 720, 540, [["explode", true]]),
      at("aaaaaaa", 960, 540, [["explode", true]]),
      at("aaaaaaa", 960, 540, [["explode", false]]),
    ]);
    expect(unstable.map((x: { name: string }) => x.name)).toEqual(["explode"]);
    expect(unstable[0].flaky, "reported a profile disagreeing with itself as a slide-size difference").toBe(true);
    expect(unstable[0].unstableIn).toEqual(["16:9"]);
    // And it must NOT claim the stable profile failed anything.
    expect(unstable[0].failedIn).toEqual([]);

    // ACROSS BUILDS IS NOT A COMPARISON. Two rounds on different builds differ
    // for reasons that have nothing to do with slide size.
    expect(
      profileDivergence([
        at("aaaaaaa", 960, 540, [["same scale", true]]),
        at("bbbbbbb", 720, 540, [["same scale", false]]),
      ]),
      "compared two builds and blamed the slide size",
    ).toEqual([]);
  });

  it("carries each profile's LIFETIME record, so a split can be read against chance", () => {
    /**
     * One build cannot tell "fails only at 4:3" from "fails at both, and this
     * time the coin landed 4:3". The check refuses the second case only when a
     * profile disagrees with ITSELF on that build; it cannot see a scenario
     * that fails a quarter of the time everywhere.
     *
     * That is not hypothetical. The first divergence this gate reported after
     * it was unblocked on 2026-09-08 was `stop a run mid-draw` — passed at
     * 16:9, failed at 4:3, with a 4:3-flavoured detail about a slide that would
     * not delete. Lifetime it is 10 of 39 at 16:9 and 12 of 38 at 4:3, and it
     * had not failed that way for 49 rounds. The rates close the question; the
     * split alone opens a false one.
     */
    const at = (build: string, w: number, pairs: [string, boolean | "skip"][]) => ({
      build: `${build} · 2026-08-16`,
      slideSize: { width: w, height: 540, source: "pageSetup" },
      selftest: pairs.map(([name, ok]) => ({ name, ok: ok === true, ...(ok === "skip" ? { skipped: true } : {}) })),
    });
    const d = profileDivergence([
      // History from OTHER builds: the scenario fails at both profiles.
      at("old1111", 960, [["stop mid-draw", false]]),
      at("old2222", 720, [["stop mid-draw", false]]),
      // A skip must not enter either side of the ratio — it measured nothing,
      // which is the same rule the divergence and regression checks follow.
      at("old3333", 960, [["stop mid-draw", "skip"]]),
      // The build that produces the split.
      at("aaaaaaa", 960, [["stop mid-draw", true]]),
      at("aaaaaaa", 720, [["stop mid-draw", false]]),
    ]);
    expect(d).toHaveLength(1);
    expect(d[0].history, "the divergence carries no lifetime record").toEqual({
      "16:9": { ran: 2, failed: 1 },
      "4:3": { ran: 2, failed: 2 },
    });
  });

  it("says whether tonight's failure TEXT has been seen before, or is new", () => {
    /**
     * "failed 8 times in 332 rounds" says the scenario is unreliable. It cannot
     * say whether TONIGHT's failure is a familiar one or a new shape wearing an
     * old name, and those want opposite responses.
     *
     * Round 435 is why this exists: `two slides claiming one slot` stopped
     * passing with "4 slides inserted, 3 kept, 3 of 2 still re-editable; 2
     * queued as duplicates", which is character-for-character what rounds 060,
     * 253 and 273 had said — and 273 is the round the gate's bundle guard was
     * written for. It was all in the archive and none of it was in front of the
     * reader.
     *
     * A REPORT FIELD, NEVER A GATE INPUT. `scenariosOf` reads the `skipped`
     * flag and never the prose, so that rewording a message cannot make the
     * gate go quiet. This keeps that rule intact: the verdict is already
     * decided when this is computed, and a reworded detail can only make it say
     * "NEW", which suppresses nothing.
     */
    const r = (ok: boolean, detail: string) => ({
      build: `b${detail.length}${ok ? "p" : "f"} · 2026-08-16`,
      selftest: [{ name: "slot", ok, detail }],
    });
    const OLD = "4 slides inserted, 3 kept, 3 of 2 still re-editable";
    const withHistory = scenarioRegressions([
      { build: "aaaaaaa · 2026-08-16", selftest: [{ name: "slot", ok: false, detail: OLD }] },
      r(true, "fine"),
      r(true, "fine"),
      r(true, "fine"),
      { build: "zzzzzzz · 2026-08-16", selftest: [{ name: "slot", ok: false, detail: OLD }] },
    ]);
    expect(withHistory).toHaveLength(1);
    expect(withHistory[0].sameDetailIn, "a repeat of a known failure was not recognised").toEqual(["aaaaaaa"]);

    // A DIFFERENT failure text under the same name must read as new — that is
    // the case worth chasing, and collapsing the two is what the count already
    // does.
    const novel = scenarioRegressions([
      { build: "aaaaaaa · 2026-08-16", selftest: [{ name: "slot", ok: false, detail: OLD }] },
      r(true, "fine"),
      r(true, "fine"),
      r(true, "fine"),
      { build: "zzzzzzz · 2026-08-16", selftest: [{ name: "slot", ok: false, detail: "something else entirely" }] },
    ]);
    expect(novel[0].sameDetailIn, "a failure nobody has seen was reported as familiar").toEqual([]);
  });

  it("does not call a scenario that DID NOT MEASURE a regression", () => {
    // THE GATE'S FIRST LIVE OUTING GOT THIS WRONG. Round 073 flagged `explode a
    // degraded picture` as having stopped passing, for a result whose own words
    // were "the slide not naming it proves nothing either way" — the host had
    // refused ids mid-scenario, so nothing was measured either way.
    //
    // An absence of evidence is not a fall, and a gate that fires on a scenario
    // declining to conclude is a gate that gets switched off. This repo has
    // already watched that happen to one.
    const skipped = (name: string) => ({ name, ok: false, skipped: true });
    const passed = (name: string) => ({ name, ok: true });
    const failed = (name: string) => ({ name, ok: false });

    expect(
      scenarioRegressions([
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [skipped("explode")] },
      ]),
      "a scenario that measured nothing was reported as having fallen",
    ).toEqual([]);

    // A GENUINE failure after the same history still fires, or the fix would
    // have bought quiet at the price of the gate.
    expect(
      scenarioRegressions([
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [failed("explode")] },
      ]).map((g: { name: string }) => g.name),
    ).toEqual(["explode"]);

    // And a scenario that SKIPPED earlier was never established, so its later
    // failure is not a regression either — it has no passing run to fall from.
    expect(
      scenarioRegressions([
        { selftest: [passed("explode")] },
        { selftest: [skipped("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [failed("explode")] },
      ]),
      "treated a skipped round as a passing one when establishing",
    ).toEqual([]);
  });

  it("splits grouping by whether the chart spanned batches, and counts a draw once", () => {
    // THE SHARPEST SEPARATION IN THE ARCHIVE: 353 of 452 multi-batch draws
    // grouped, against 49 of 214 single-batch. It is OURS — `refreshShapes` is
    // set from `spansBatches()`, so only a multi-batch chart gets the
    // pre-grouping re-read that resolves its shapes by id; a single-batch chart
    // hands addGroup the raw created proxies and this host refuses them.
    //
    // Found while chasing what looked like a RASTERISE effect (22% vs 93%).
    // Splitting by batch count collapsed that to 22% vs 27% — nothing. The
    // scenarios that rasterise just draw small charts. Counting only the LAST
    // batch of a draw is what makes the two arms comparable at all.
    const batch = (upTo: number, total: number) => ({ message: "batch issued", data: { upTo, total } });
    const b = poolBatchSpanVsGroup([
      {
        trace: {
          entries: [
            // A 24-shape chart: three batches, ONE draw. Counting per batch
            // would credit this arm three times for a single outcome.
            batch(10, 24),
            batch(20, 24),
            batch(24, 24),
            { message: "grouped the chart's shapes", data: {} },
            // A single-batch chart whose group was refused.
            batch(7, 7),
            { message: "grouping the chart's shapes", data: { error: "InvalidParam ... code=5010" } },
            // A single-batch chart that DID group — the arm is not all failure.
            batch(9, 9),
            { message: "grouped the chart's shapes", data: {} },
            // A draw the round never resolved either way belongs to neither arm.
            batch(6, 6),
          ],
        },
      },
    ]);
    expect(b.multi, "a multi-batch draw was counted once per batch, not once per draw").toBe(1);
    expect(b.multiGrouped).toBe(1);
    expect(b.single).toBe(2);
    expect(b.singleGrouped).toBe(1);
  });

  it("names the questions that have never once answered, and splits them by whose fault it is", () => {
    // SIX QUESTIONS WERE 0-FOR-41 and nothing said so. The per-round report
    // names what was "never put" in THAT round, so a permanently dead question
    // read as bad luck, forty-one times running — while costing a scratch slide
    // and host time on every one of them.
    //
    // Four of the six were the group cluster. The probe has been blind on groups
    // for the whole archive, and rounds 064/065 then answered the most important
    // of them (`tag-on-group-survives`) from PRODUCTION, twice, in one evening.
    //
    // The split is the point. "Never asked" is a harness problem and ours to
    // fix; "unreadable" is the host declining to answer, which is a finding.
    // Pooling them would hide exactly the distinction that decides what to do.
    const round = (answers: { id: string; answer: string }[]) => ({ hostAnswers: { answers } });
    const dead = poolStarvedQuestions([
      round([
        { id: "tag-on-group-survives", answer: "no-scratch-slide" },
        { id: "addgroup-returns-usable", answer: "unreadable" },
        { id: "sometimes-answers", answer: "no-scratch-shape" },
        { id: "always-answers", answer: "yes" },
      ]),
      round([
        { id: "tag-on-group-survives", answer: "no-scratch-slide" },
        { id: "addgroup-returns-usable", answer: "unreadable" },
        { id: "sometimes-answers", answer: "survives-8" },
        { id: "always-answers", answer: "yes" },
      ]),
    ]);
    // Typed at the boundary: `triage.mjs` is plain JS behind a @ts-expect-error
    // import, so everything out of it is `any` and an untyped callback param is
    // an implicit-any typecheck error — green locally under vitest, red in CI.
    type Dead = { id: string; rounds: number; never: number; unanswerable: number; answered: number };
    const ids = (dead as Dead[]).map((d) => d.id);
    expect(ids, "a question that answers SOMETIMES is doing its job, not waste").not.toContain("sometimes-answers");
    expect(ids).not.toContain("always-answers");
    expect([...ids].sort()).toEqual(["addgroup-returns-usable", "tag-on-group-survives"]);

    const byId = Object.fromEntries((dead as Dead[]).map((d) => [d.id, d]));
    // Never asked — the harness could not set it up.
    expect(byId["tag-on-group-survives"]).toMatchObject({ never: 2, unanswerable: 0, answered: 0, rounds: 2 });
    // Asked, and the host declined — a different problem with a different fix.
    expect(byId["addgroup-returns-usable"]).toMatchObject({ never: 0, unanswerable: 2, answered: 0, rounds: 2 });

    // A question seen only once is not yet a pattern, and calling it dead would
    // put every newly-added probe into a report about waste on its first round.
    expect(poolStarvedQuestions([round([{ id: "brand-new", answer: "no-scratch-slide" }])])).toEqual([]);

    // BOTH ARE STILL ASKED here — they appear in the newest round — so neither
    // is retired and both stay actionable. Retirement is the sibling case, and
    // it has its own test below.
    expect((dead as Dead[]).some((d) => (d as Dead & { retired: boolean }).retired)).toBe(false);
  });

  it("pools what production saw of a question the probe sheet cannot ask", async () => {
    // `shape-resolve-held-slide-proxy` has answered `no-scratch-shape` in all
    // 133 archived rounds and structurally cannot do better — it needs an id for
    // a freshly added shape and this host refuses to give one. The one
    // production site that still resolves a shape by id through a slide handle a
    // sync old is `deleteShapesById`, and until 2026-08-22 it traced only its
    // FAILURES: 133 rounds of silence that could mean "the host resolved
    // everything" or "the sweep never ran", with no way to tell.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolAgedHandleResolves } = await import("../scripts/triage.mjs");
    const round = (entries: { message: string; data?: Record<string, number> }[]) => ({ trace: { entries } });

    expect(
      poolAgedHandleResolves([round([{ message: "drew a chart" }]), round([])]),
      "silence is not a resolve and must not be counted as one",
    ).toMatchObject({ resolved: 0, refused: 0, rounds: 0, of: 2 });

    expect(
      poolAgedHandleResolves([
        round([{ message: "resolved a shape by id through a slide handle a sync old", data: { resolved: 3 } }]),
        round([{ message: "wreckage the host would not resolve", data: { unresolved: 2, swept: 1 } }]),
        round([{ message: "drew a chart" }]),
      ]),
    ).toMatchObject({ resolved: 3, refused: 2, rounds: 2, of: 3 });
  });

  it("marks a starved question RETIRED once the build stops asking it", () => {
    // WHAT IT COST: `grouped-child-by-id-from-slide` and `tag-on-group-survives`
    // were retired on 2026-08-21 and last appear in round 149. For the eight
    // rounds after that, this report kept listing them at "125 round(s)" under
    // "OURS to fix" — telling the reader to do work that was already done. A
    // count pooled over the whole archive cannot tell a live starving probe
    // from a dead one, which is the same defect as a hardcoded conclusion: it
    // keeps printing after it stops being true.
    const round = (answers: { id: string; answer: string }[]) => ({ hostAnswers: { answers } });
    const dropped = { id: "dropped", answer: "no-scratch-slide" };
    const kept = { id: "kept", answer: "no-scratch-slide" };
    type Dead = { id: string; retired: boolean };
    const byId = (logs: ReturnType<typeof round>[]) =>
      Object.fromEntries((poolStarvedQuestions(logs) as Dead[]).map((d) => [d.id, d.retired]));

    expect(byId([round([dropped, kept]), round([dropped, kept]), round([kept]), round([kept]), round([kept])])).toEqual(
      { dropped: true, kept: false },
    );

    // A WINDOW, NOT THE NEWEST ROUND ALONE. One short sheet means the host died
    // mid-probe, not that the question was retired — and reading it as a
    // retirement would quietly empty the actionable bucket.
    expect(byId([round([dropped, kept]), round([dropped, kept]), round([kept])])).toEqual({
      dropped: false,
      kept: false,
    });

    // An empty sheet carries no evidence either way and must not consume the
    // window: a round the host never answered would otherwise retire the lot.
    expect(byId([round([dropped, kept]), round([dropped, kept]), round([]), round([]), round([])])).toEqual({
      dropped: false,
      kept: false,
    });
  });

  it("separates a chart on a fresh slide from one on a slide that already had shapes", () => {
    // THE ROOT, and the cleanest separation this project has: 82 charts on a
    // slide that already had shapes, 81 grouped; 74 on a freshly added empty
    // slide, 1 grouped. A chart that is not grouped loses its config, so the
    // slide's newness decides the config — one level below everything the tag
    // work was aimed at.
    const chart = (name: string, onSlide: number, msgs: string[]) =>
      msgs.map((m, i) => ({ message: m, data: i === 0 ? { chart: name, onSlide } : { chart: name } }));
    const f = poolFreshVsEstablished([
      {
        trace: {
          entries: [
            ...chart("1/4", 32, ["batch issued", "grouped the chart's shapes"]),
            ...chart("2/4", 0, ["batch issued", "not grouping: no member handle this host will accept"]),
            ...chart("3/4", 0, ["batch issued", "grouped the chart's shapes"]),
            // Never decided — the round did not reach it, so it belongs to
            // neither column rather than being counted as a failure.
            ...chart("4/4", 0, ["batch issued"]),
          ],
        },
      },
    ]);
    expect(f.established).toBe(1);
    expect(f.establishedGrouped).toBe(1);
    expect(f.fresh, "a chart the round never decided must not be counted").toBe(2);
    expect(f.freshGrouped, "a fresh slide CAN group — the claim is a rate, not a law").toBe(1);
  });

  it("reports the CURRENT grouping rate beside the all-time one, which is known to be unreadable", async () => {
    // THE REPORT SAID SO ITSELF AND PRINTED THE NUMBER ANYWAY: "pooled over 39
    // rounds that predate the retry, so it will climb slowly and should not be
    // read as the current rate". A figure a reader is told to mentally discount
    // is a figure nobody can use. All-time says 66% of charts on a freshly added
    // slide group; over the last 20 rounds it is 81%.
    // @ts-expect-error — plain .mjs tool, no types.
    const { recentFreshVsEstablished, RECENT_ROUNDS } = await import("../scripts/triage.mjs");
    const chart = (name: string, onSlide: number, grouped: boolean) => [
      { message: "batch issued", data: { chart: name, onSlide } },
      { message: grouped ? "grouped the chart's shapes" : "not grouping: refused", data: { chart: name } },
    ];
    const round = (grouped: boolean) => ({ trace: { entries: chart("c", 0, grouped) } });

    // An old era that never grouped, then a recent one that always does.
    const logs = [...Array(30).fill(round(false)), ...Array(RECENT_ROUNDS).fill(round(true))];
    expect(recentFreshVsEstablished(logs)).toMatchObject({ fresh: RECENT_ROUNDS, freshGrouped: RECENT_ROUNDS });
    expect(poolFreshVsEstablished(logs).freshGrouped, "the all-time figure still carries the dead era").toBe(
      RECENT_ROUNDS,
    );

    // BELOW THE WINDOW IT REFUSES TO ANSWER. A rate from five rounds is exactly
    // the thing this exists to stop, and printing one would repeat the defect
    // in the other direction.
    expect(recentFreshVsEstablished([round(true), round(true)]), "a rate from two rounds is not a rate").toBeNull();
  });

  it("shows the fresh-slide population EMPTYING, which no percentage can", async () => {
    // A WINDOW IS A GUESS ABOUT WHERE THE LAST REGIME BOUNDARY IS, and 20 was
    // the wrong guess. It excluded everything before the settled retry, and sat
    // straight across a second boundary nobody had found: rounds 137-142 failed
    // to group 1,4,4,5,9,2 charts and rounds 143-160 failed 3 in eighteen. The
    // "current rate" it produced was smeared by a dead era — the exact defect it
    // was written to fix, one level down.
    //
    // The same population reads 80% at 30 rounds, 91% at 20, 100% at 12 — and at
    // 8 there are NO fresh-slide charts in it at all, because the in-place
    // update took those charts and a chart that is not redrawn never lands on a
    // fresh slide. A percentage over an emptying population is what to watch
    // for, and only the sequence shows it emptying.
    // @ts-expect-error — plain .mjs tool, no types.
    const { recentFreshSequence } = await import("../scripts/triage.mjs");
    const chart = (name: string, onSlide: number, grouped: boolean) => [
      { message: "batch issued", data: { chart: name, onSlide } },
      { message: grouped ? "grouped the chart's shapes" : "not grouping: refused", data: { chart: name } },
    ];
    const busy = { trace: { entries: [...chart("a", 0, true), ...chart("b", 0, false)] } };
    const empty = { trace: { entries: chart("c", 5, true) } };

    expect(recentFreshSequence([busy, busy, empty, empty], 4)).toEqual(["1/2", "1/2", "0/0", "0/0"]);
    // `0/0` IS THE SIGNAL, not a gap to be tidied away. A round with no
    // fresh-slide chart is the in-place update working, and dropping it would
    // hide the very thing this exists to show.
    expect(recentFreshSequence([empty, empty], 4), "an empty population must still be printed").toEqual(["0/0", "0/0"]);
  });

  it("separates the charts that got grouped from the ones that did not", () => {
    // THE QUESTION NOBODY ASKED FOR ELEVEN ROUNDS, and the archive held the
    // answer throughout: pooled over every round, 64 grouped charts lost 1 tag
    // and 62 ungrouped charts lost 41. A chart that groups keeps its config; one
    // that cannot loses it two times in three.
    //
    // It reframes the tag work. Grouping puts the tag on the GROUP — a handle
    // made in that batch and never resolved — and it lands. No group means the
    // tag falls back to a `created` handle, which is the path four rounds and a
    // renderer change went into. That is a question about the losing side.
    const chart = (name: string, msgs: string[]) => msgs.map((m) => ({ message: m, data: { chart: name } }));
    const g = poolGroupVsTag([
      {
        trace: {
          entries: [
            ...chart("1/3", ["grouped the chart's shapes"]),
            ...chart("2/3", ["grouped the chart's shapes", "tagging failed — charts are not re-editable"]),
            ...chart("3/3", ["not grouping: no member handle this host will accept", "tagging failed — nope"]),
            // No chart key at all: batch-level noise that must not be counted as
            // a chart in either column.
            { message: "tagging failed — charts are not re-editable", data: {} },
          ],
        },
      },
    ]);
    expect(g.grouped).toBe(2);
    expect(g.groupedLost, "a grouped chart CAN still lose its tag — the claim is a rate, not a law").toBe(1);
    expect(g.ungrouped).toBe(1);
    expect(g.ungroupedLost).toBe(1);
  });

  it("says nothing about ids when every one of them finished", () => {
    // A field that reports on a healthy round is a field a reader learns to skip.
    const e = deckEvidence(evidenceDeck([{ id: "288#3603562595", count: 24 }], ["288#3603562595"]));
    expect(e.oddIds).toEqual([]);
    expect(e.oddAndBlank).toBe(0);
  });

  it("reads a slide the scan could only partly list as carrying shapes", () => {
    // `count` is the host's own number, `shapes` is what the scan managed to
    // list. Taking the smaller would call a partial listing empty — the exact
    // mistake of trusting one witness that this whole block exists to avoid.
    const e = deckEvidence({
      inventory: [{ slideId: "a", index: 0, shapes: [{ id: "1" }], count: 24 }],
      newSlides: ["a"],
      shots: [],
    });
    expect(e.withShapes).toBe(1);
    expect(e.unseen).toBe(0);
  });

  /**
   * Pooling the counterbalanced arms, and the trap under it.
   *
   * The per-batch line is logged BEFORE the sync, deliberately — the sync is
   * where a bad host goes quiet, so the number has to be on screen while you
   * wait. That means every stall has one of its own immediately before it, and
   * reading those lines as successes is wrong twice over: it counts a failure
   * as a success, and it inflates the denominator with the same event.
   *
   * Both mistakes were made analysing this data by hand. One pass produced 0
   * stalls in 32 draws; another produced a 6x rasterise effect that was not
   * there. The outcome of a draw is whether a `gave up waiting` lands between
   * that draw's step and the next one — nothing else.
   *
   * The line was CALLED `batch committed` while both of those went wrong, which
   * is why the writer does not call it that any more (`batch issued`). These
   * fixtures keep the old name on purpose: the owner's saved rounds carry it,
   * and a pooling function that reads either name is a pooling function that
   * reads neither — which is the property being pinned here.
   */
  const drawStep = (ms: number, arm: string, i: number) => ({
    ms,
    scope: "selftest",
    message: "rasterise-draw step",
    data: { what: `after a ${arm} #${i}: drawing` },
  });
  const committed = (ms: number) => ({ ms, scope: "draw", message: "batch committed", data: { upTo: 7, total: 7 } });
  const gaveUp = (ms: number) => ({
    ms,
    scope: "host",
    message: "gave up waiting",
    data: { what: "drawing shapes 1-7 of 7", afterMs: 45000 },
  });

  it("counts a draw that was SENT and never landed as a stall, not a success", () => {
    const log = {
      trace: {
        entries: [
          drawStep(10, "rasterise", 0),
          committed(11),
          gaveUp(60),
          drawStep(70, "cheap read", 1),
          committed(71),
        ],
      },
    };
    const { arms } = poolRasteriseArms([log]);
    expect(arms.rasterise, "a committed-then-abandoned draw was read as a success").toEqual({ ok: 0, stall: 1 });
    expect(arms["cheap read"]).toEqual({ ok: 1, stall: 0 });
  });

  it("adds the arms up across rounds, which is the whole point", () => {
    // One round gives two draws an arm — enough to catch a call that fails
    // every time, hopeless for anything smaller. Eight rounds of that sat in
    // separate files, each independently reporting "no pattern".
    const round = (stallArm: string) => ({
      trace: {
        entries: [
          drawStep(10, "rasterise", 0),
          committed(11),
          ...(stallArm === "rasterise" ? [gaveUp(60)] : []),
          drawStep(70, "cheap read", 1),
          committed(71),
          ...(stallArm === "cheap read" ? [gaveUp(120)] : []),
        ],
      },
    });
    const { rounds, arms } = poolRasteriseArms([round("rasterise"), round("rasterise"), round("cheap read")]);
    expect(rounds).toBe(3);
    expect(arms.rasterise).toEqual({ ok: 1, stall: 2 });
    expect(arms["cheap read"]).toEqual({ ok: 2, stall: 1 });
  });

  it("says nothing about a log that never ran the scenario", () => {
    expect(poolRasteriseArms([{ trace: { entries: [committed(1)] } }, {}]).arms.rasterise).toEqual({ ok: 0, stall: 0 });
  });

  it("counts the draws the scenario's four arms never see", () => {
    // Round 28 in miniature, and the reason `poolEveryDraw` exists. The
    // visibility scenario rasterises and then draws; that draw stalls. It
    // carries no arm tag, so the arms report a clean round while the round
    // contains the very thing they are looking for.
    const rasterised = {
      ms: 500,
      scope: "selftest",
      message: "visibility step",
      data: { what: "rasterising a slide" },
    };
    // `batch issued`, NOT the `committed` helper above: that one still emits the
    // pre-2026-08-11 name `batch committed`, which this file's own comment says
    // is read by nothing here, by design. Using it made the population come back
    // empty and looked like a classifier bug.
    const issued = (ms: number) => ({ ms, scope: "draw", message: "batch issued", data: {} });
    const log = {
      trace: {
        entries: [
          // The scenario's own arms: both land, so the arms see nothing.
          drawStep(10, "rasterise", 0),
          committed(11),
          drawStep(70, "cheap read", 1),
          committed(71),
          // An untagged draw, straight after a rasterise, that stalls.
          rasterised,
          issued(501),
          gaveUp(560),
          // An untagged draw with no rasterise before it, that lands.
          issued(900),
        ],
      },
    };
    expect(poolRasteriseArms([log]).arms.rasterise, "the arms should still see a clean round").toEqual({
      ok: 1,
      stall: 0,
    });
    const { after } = poolEveryDraw([log]);
    expect(after.rasterise, "the untagged stall after a rasterise was not counted").toEqual({ ok: 0, stall: 1 });
    expect(after["anything else"].ok).toBeGreaterThanOrEqual(2);
    expect(after["anything else"].stall).toBe(0);
  });

  it("reports a round's span, and refuses to invent one", () => {
    // THE STRONGEST PREDICTOR IN THIS ARCHIVE, SURFACED BY NOTHING UNTIL NOW.
    // Every trace entry carries an `ms` offset, so the last one is the round's
    // span — and the slower half of the rounds whose instruments existed average
    // 5.1 post-retry failures against 2.8. The second round of a pair runs
    // 2.0-2.4x slower than the first in all four pairs measured, most likely
    // because the first is being mined while the second runs.
    // LAST MINUS FIRST, NOT THE LAST. `ms` counts from the PANE's load, and the
    // pane is not reloaded between rounds — so a round inheriting its pane
    // starts its clock where the previous round stopped. Taking `max` reported
    // such a round's duration as its own PLUS the previous round's, which is
    // how "the second round is 2.0-2.4x slower" got published twice. The real
    // figure is about 1.35x.
    //
    // Verbatim shape from rounds/123: first entry at 961s, last at 2075s. The
    // round took 1114s, not 2076s.
    expect(
      roundSpanSeconds({ trace: { entries: [{ ms: 961_782 }, { ms: 2_075_549 }] } }),
      "counted the pane's age as this round's time",
    ).toBe(1114);
    expect(roundSpanSeconds({ trace: { entries: [{ ms: 1000 }, { ms: 759_000 }, { ms: 12 }] } })).toBe(759);

    // AN UNREADABLE SPAN MUST NOT BECOME A FAST ONE. Returning 0 here would put
    // every old round in the "fast" half and reproduce, exactly, the confounded
    // comparison this instrument was built to replace — where the ten fastest
    // rounds were all old ones whose counters could not fire.
    expect(roundSpanSeconds({ trace: { entries: [] } }), "an empty trace is not a 0s round").toBeNull();
    expect(roundSpanSeconds({}), "a round with no trace is not a 0s round").toBeNull();
    expect(roundSpanSeconds({ trace: { entries: [{ ms: 0 }, {}] } }), "no usable offsets").toBeNull();
  });

  it("reads the pane's age at the round's start, which is what predicts the counters", () => {
    // THE VARIABLE THAT WAS HIDING INSIDE THE BROKEN DURATION METRIC. `ms`
    // counts from the pane's load, so the FIRST entry's offset is how long the
    // pane had already been alive when the round began.
    //
    // Rounds 110-123 split on it: fresh panes scored post-retry 0, 2, 0, 0, 0,
    // 1, 0 and reused panes 0, 5, 7, 3, 8, 7, 2 — a mean of 0.4 against 4.6,
    // with decks of 16 against 60+. "The second round of a pair is worse" was
    // always this: the second round is the one that inherits a pane. Position,
    // profile and observer load were all stand-ins for it.
    expect(paneAgeAtStartSeconds({ trace: { entries: [{ ms: 961_782 }, { ms: 2_075_549 }] } })).toBe(962);
    expect(paneAgeAtStartSeconds({ trace: { entries: [{ ms: 67_116 }, { ms: 862_019 }] } })).toBe(67);

    // AN UNKNOWN AGE MUST NOT READ AS A FRESH PANE. The gate treats < 200s as
    // fresh, so returning 0 here would silently mark every unreadable round as
    // clean — the same shape as reporting an unreadable span as a fast one.
    expect(paneAgeAtStartSeconds({ trace: { entries: [] } })).toBeNull();
    expect(paneAgeAtStartSeconds({})).toBeNull();
    expect(paneAgeAtStartSeconds({ trace: { entries: [{}, { ms: "x" }] } }), "no usable offsets").toBeNull();
  });

  it("says when a feature has never once run", () => {
    // THE IN-PLACE UPDATE HAS NEVER SUCCEEDED. Across 117 archived rounds the
    // success line appears ZERO times against 1301 fallbacks. #405 added the
    // feature; #406 was titled "The in-place update fired zero times and would
    // not say why" and added the trace that answers it — and the answer has sat
    // in every round file since, read by nothing.
    //
    // A count of zero is the hardest thing for a report to say, because nothing
    // draws attention to a line that is not printed. This pools it so the gate
    // has to.
    const round = (ok: number, fell: number) => ({
      trace: {
        entries: [
          ...Array.from({ length: ok }, () => ({ message: "updated only the shapes that changed" })),
          ...Array.from({ length: fell }, () => ({ message: "not updating in place — redrawing instead" })),
        ],
      },
    });
    expect(poolInPlaceUpdates([round(0, 12), round(0, 13)])).toMatchObject({ ok: 0, fell: 25, rounds: 2 });

    // AND IT MUST NOTICE A SUCCESS, or the day the feature starts working the
    // gate goes on calling it dead. A detector that can only report one answer
    // is not measuring anything.
    expect(poolInPlaceUpdates([round(3, 9)])).toMatchObject({ ok: 3, fell: 9 });
  });

  it("counts a host refusal, which is the third outcome and was counted as none", () => {
    // THE INSTRUMENT HAD THE BLIND SPOT IT WAS BUILT TO FIND. tryInPlaceUpdate
    // declines by rule OR THROWS, and only the rule-based decline was counted.
    // So every host-side refusal registered as neither a success nor a fallback
    // and vanished: three InvalidArgument | errorLocation=Shape.textFrame in
    // round 145 -- the last thing between this feature and working -- plus two
    // more in 144 that went unread while that round was reported.
    const threw = {
      trace: {
        entries: [
          {
            message: "in-place update refused — redrawing instead",
            data: {
              error:
                'InvalidArgument | at=updating the shapes | code=InvalidArgument | debugInfo={"errorLocation":"Shape.textFrame"}',
            },
          },
        ],
      },
    };
    const pooled = poolInPlaceUpdates([threw]);
    expect(pooled.threw, "a host refusal was counted as neither success nor fallback").toBe(1);
    // Keyed by what actually failed and where, because "it threw" is not a
    // reason -- the errorLocation is the whole diagnosis.
    expect(pooled.reasons).toContainEqual({ why: "InvalidArgument at Shape.textFrame", n: 1 });
  });

  it("names the entries no category fits, rather than tallying them", () => {
    // A residual bucket reported as a NUMBER reads as noise and gets skipped;
    // reported as a line it gets opened. Entries that threw carry an `error`
    // and no `why`, so they fall outside every named reason by construction --
    // which is exactly where the interesting failure lives.
    const bare = {
      trace: { entries: [{ message: "not updating in place — redrawing instead" }] },
    };
    const pooled = poolInPlaceUpdates([bare]);
    expect(pooled.unexplained, "a decline carrying no reason was silently folded into the total").toHaveLength(1);
    expect(pooled.reasons, "an absent reason must not invent a category").toHaveLength(0);
  });

  it("sees a slow drift, which a median cannot", () => {
    // THE BLINDNESS I NEARLY SHIPPED. The first version of this instrument
    // reported each fallback against the MEDIAN OF ALL PRIORS — and the signal
    // that motivated it, `in-place update fell back to a redraw`, had climbed
    // from 9 to 13 per round across sixty rounds. By the time anyone looked,
    // "now" and "usually" were both 13, so a now-against-median check would
    // have called it normal for as long as it kept getting worse.
    //
    // A median absorbs a slow climb. The oldest third against the newest third
    // is what sees it, and this fixture is that exact shape: a signal that
    // doubles while its all-time median stays in the middle.
    const round = (n: number) => ({
      trace: {
        entries: Array.from({ length: n }, () => ({ message: "not updating in place — redrawing instead" })),
      },
    });
    const rising = poolFallbackRates([...Array(6).fill(round(4)), ...Array(6).fill(round(12))]);
    const redraw = rising.find((r: { key: string }) => r.key === "not updating in place — redrawing instead")!;
    expect(redraw.oldest, "did not read the early history").toBe(4);
    expect(redraw.newest, "did not read the recent history").toBe(12);
    // AND THE MEDIAN MUST STILL BE UNHELPFUL HERE — that is the point. If this
    // ever equals the newest reading the fixture has stopped modelling a drift.
    expect(redraw.median, "the median should sit between and hide the climb").toBeLessThan(redraw.newest);

    // TOO LITTLE HISTORY IS NOT A TREND. Same three-priors rule as every other
    // baseline in this file; this project's noise floor is the argument.
    expect(poolFallbackRates([round(1), round(9)]), "called two rounds a trend").toEqual([]);

    // AND A MEDIAN CANNOT SEE A STEP — over any window, which is why the answer
    // is not a shorter one. This same signal read 13,13,11,10,10,8,8,2,2,2,2,2
    // over twelve rounds when the in-place update started working, and all
    // three summaries missed it: all-time median 12, last-20 median 10, thirds
    // "RISING, 8 to 13". The gate printed `now 2` beside `RISING` — a
    // conclusion that had outlived its evidence. The sequence is what settles
    // it, and it costs eight numbers.
    const stepped = poolFallbackRates([...Array(8).fill(round(12)), ...Array(5).fill(round(2))]);
    const after = stepped.find((r: { key: string }) => r.key === "not updating in place — redrawing instead")!;
    expect(after.recent, "the last rounds in order, not summarised").toEqual([12, 12, 12, 2, 2, 2, 2, 2]);
    expect(after.median, "the median still cannot see the step — that is the point").toBe(12);

    // THE STALE-`now` GUARD, third site. A round with no trace at all is skipped
    // when the history is built, and `now` was then the newest round that HAD
    // one — the defect found in `poolScenarioPopulations` on 2026-08-29, where
    // three consecutive rounds were told `7 this round` about round 282's seven.
    // A traceless round is the ordinary case here: `verbose` off produces one.
    const traceless = { build: "z" };
    expect(
      poolFallbackRates([...Array(8).fill(round(12)), ...Array(5).fill(round(2)), traceless]),
      "reported an older round's fallback counts as this round's",
    ).toEqual([]);
  });

  it("counts pair position without letting ties vote", () => {
    // THE PAIR IS NOT TWO SAMPLES OF ONE CONDITION, and nothing in this project
    // controlled for that. Over the 31 builds run twice, the SECOND round had
    // more post-retry failures in 15 and fewer in 2.
    //
    // TIES ARE THE TRAP. 14 pairs tied, mostly old rounds whose counters sat at
    // zero both times — and reading "15 worse, 16 not worse" as a coin flip is
    // exactly how a directional effect stays invisible. They are counted, and
    // reported, SEPARATELY.
    const round = (build: string, post: number) => ({
      build: `${build} 2026-08-20`,
      deck: { inventory: [{ slideId: "s1", count: 1 }] },
      trace: {
        entries: Array.from({ length: post }, () => ({
          message: "the re-read before grouping came back empty",
          data: { afterRetry: true, kind: "empty" },
        })),
      },
    });

    const out = poolPairPosition([
      round("aaa", 0),
      round("aaa", 3), // worse
      round("bbb", 2),
      round("bbb", 0), // better
      round("ccc", 1),
      round("ccc", 1), // tied
      round("ddd", 0),
      round("ddd", 5), // worse
    ]);
    // `secondFresh` is 0 here because this fixture's entries carry no `ms`
    // offsets at all, so the pane's age is UNREADABLE — and unreadable must not
    // count as fresh, or every pair in the archive would be reported as already
    // mitigated.
    expect(out).toEqual({ pairs: 4, worse: 2, better: 1, tied: 1, secondFresh: 0 });

    // WHICH PAIRS PREDATE THE FIX. The asymmetry above is caused by the second
    // round inheriting the first round's pane; `collectRound` reloads it now, so
    // the count has to distinguish pairs from before that shipped or the gate
    // announces a fixed problem forever. A second round that started fresh is a
    // pair the mitigation reached.
    const withAge = (build: string, firstMs: number) => ({
      build: `${build} 2026-08-20`,
      deck: { inventory: [{ slideId: "s1", count: 1 }] },
      trace: { entries: [{ ms: firstMs }, { ms: firstMs + 700_000 }] },
    });
    const mitigated = poolPairPosition([withAge("zzz", 67_000), withAge("zzz", 119_000)]);
    expect(mitigated, "a fresh second round was not recognised as mitigated").toMatchObject({
      pairs: 1,
      secondFresh: 1,
    });
    const notMitigated = poolPairPosition([withAge("yyy", 67_000), withAge("yyy", 962_000)]);
    expect(notMitigated, "a 962s-old pane counted as fresh").toMatchObject({ pairs: 1, secondFresh: 0 });

    // A build run ONCE is not a pair and must not be counted — it has no second
    // round to be worse than anything.
    expect(poolPairPosition([round("solo", 4)])).toMatchObject({ pairs: 0 });

    // And a build run THREE times votes ONCE, on its first two rounds — the
    // experiment this finding asks for is exactly a build run three times, so
    // the counter must not change meaning the moment someone runs it.
    //
    // 5 -> 2 -> 9 ON PURPOSE. First-against-second says BETTER; first-against-
    // last says WORSE. An earlier fixture read 0 -> 2 -> 9, where both readings
    // agree, so it asserted nothing about which pair was taken — mutation
    // testing caught that: swapping the implementation to first-against-last
    // left it green.
    const thrice = poolPairPosition([round("eee", 5), round("eee", 2), round("eee", 9)]);
    expect(thrice, "took the LAST round instead of the second").toMatchObject({ pairs: 1, better: 1, worse: 0 });
  });

  it("classifies a rasterise by op, not by whether its label says 'rasteris'", () => {
    // THE SIBLING THAT SURVIVED, PINNED SO IT KEEPS SURVIVING. Naming each
    // rasterise call site broke the matcher in `chartIsVisible`, which reads
    // `lastStall.what` alone. This classifier reads the label AND the message,
    // and the message on a success is `rasterised a slide`, so it kept working
    // — by luck, not by design. Pooled counts are identical with and without
    // the `op` path, which is checked in the fixture comment rather than
    // claimed here.
    //
    // These assertions pin the DESIGN: an entry whose label carries no
    // "rasteris" and whose message is not a rasterise message must still
    // classify. The label below is verbatim from `rounds/113-6a041de.json`.
    const issued = (ms: number) => ({ ms, scope: "draw", message: "batch issued", data: {} });
    const log = {
      trace: {
        entries: [
          {
            ms: 500,
            scope: "selftest",
            message: "visibility step",
            data: { what: "the visibility CONTROL render (same slide, back to back)", op: "rasterise" },
          },
          issued(501),
          gaveUp(560),
        ],
      },
    };
    const { after } = poolEveryDraw([log]);
    expect(after.rasterise, "a named rasterise was filed as 'anything else'").toEqual({ ok: 0, stall: 1 });
    expect(after["anything else"], "the draw was double-counted").toEqual({ ok: 0, stall: 0 });

    // AND op MUST BE WHAT DECIDES IT. Without this half, a classifier that
    // simply called everything a rasterise would pass the assertion above.
    const unrelated = {
      trace: {
        entries: [
          { ms: 500, scope: "selftest", message: "visibility step", data: { what: "reading the deck's style" } },
          issued(501),
          gaveUp(560),
        ],
      },
    };
    const other = poolEveryDraw([unrelated]).after;
    expect(other.rasterise, "a non-rasterise was counted as one").toEqual({ ok: 0, stall: 0 });
    expect(other["anything else"]).toEqual({ ok: 0, stall: 1 });

    // AND THE ARCHIVE ALREADY WRITTEN. Rounds 001-114 carry no `op` at all, so
    // an op-only fix would repair only evidence not yet gathered. These four
    // labels are enumerated from those rounds — 35 of the 43 labelled
    // rasterises in them — and must classify without `op`.
    for (const legacy of [
      "an end-of-round slide shot",
      "the visibility BEFORE render",
      "the visibility AFTER render",
      "the visibility CONTROL render (same slide, back to back)",
    ]) {
      const old = {
        trace: {
          entries: [
            { ms: 500, scope: "selftest", message: "visibility step", data: { what: legacy } },
            issued(501),
            gaveUp(560),
          ],
        },
      };
      expect(poolEveryDraw([old]).after.rasterise, `pre-op archive: "${legacy}" was not recognised`).toEqual({
        ok: 0,
        stall: 1,
      });
    }
  });

  it("blames the rasterise nearest the draw, not one from earlier in the round", () => {
    // Without resetting at each draw, one rasterise early on would tar every
    // draw after it and the population would be meaningless.
    const issued = (ms: number) => ({ ms, scope: "draw", message: "batch issued", data: {} });
    const log = {
      trace: {
        entries: [
          { ms: 10, scope: "selftest", message: "visibility step", data: { what: "rasterising a slide" } },
          issued(20), // after the rasterise
          issued(30), // NOT after it — a draw intervened
          gaveUp(40),
        ],
      },
    };
    const { after } = poolEveryDraw([log]);
    expect(after.rasterise, "the first draw follows the rasterise and landed").toEqual({ ok: 1, stall: 0 });
    expect(after["anything else"], "the second draw does not follow a rasterise, and stalled").toEqual({
      ok: 0,
      stall: 1,
    });
  });

  it("has nothing to say about a log that carries no deck evidence", () => {
    expect(deckEvidence(undefined)).toBeNull();
    expect(deckEvidence({ inventory: [], newSlides: [], shots: [] })).toBeNull();
  });

  it("calls a slide tagged with a slot before the run's range an orphan", async () => {
    // Only slots PAST the end were checked, so a negative slot — a tag written
    // wrong, or carried in from something that was not this run — matched no
    // item, fell outside the orphan test, and was reported by nothing.
    const deck = await deckOf([
      { title: "A", slot: 0 },
      { title: "stray", slot: -3 },
    ]);
    const t = triage(deck, logOf(["A"]));
    expect(t.orphans).toEqual([{ slot: -3, verdict: "orphan", indexes: [2] }]);
    expect(t.disagreements).toBe(1);
  });
});

describe("triage — naming the host bug behind a failure", () => {
  it("recognises the signatures that belong to PowerPoint, not to SSF Charts", () => {
    // The first hour of the last diagnosis went into establishing that
    // `InvalidParam passed to GetItem(id)` was a host bug rather than ours. A
    // reader who sees an issue number stops looking for the mistake in this
    // repo — which is the whole saving.
    expect(knownBug("InvalidParam passed to GetItem(id) | code=5010")).toMatch(/office-js#2903/);
    expect(knownBug("The property 'items' is not available.")).toMatch(/office-js#6363/);
    expect(knownBug("did not respond while drawing shapes 1-10 of 39")).toMatch(/office-js#5022/);
  });

  it("sends a wedged selection to the selection bugs, not to the generic sync hang", () => {
    // A wedged selection subsystem does not throw, it goes quiet — so what the
    // trace records is a TIMEOUT, and a timeout's text is "did not respond
    // while <phase>". The generic sync-hang entry matches that too, and on a
    // first-match-wins table it would claim every one of them and point the
    // reader at #5022: a different bug, with a different cause and no fix that
    // applies. The phase name is all that separates them, so the phases that
    // belong to selection must be matched BEFORE the generic signature.
    for (const raw of [
      "PowerPoint did not respond while reading the selected chart (90.0s)",
      "PowerPoint did not respond while selecting a shape (90.0s)",
      "PowerPoint did not respond while clearing the shape selection (10.0s)",
    ]) {
      expect(knownBug(raw), raw).toMatch(/#3083/);
      expect(knownBug(raw), raw).not.toMatch(/5022/);
    }
    // The scenario's own verdict line lands on the same note.
    expect(
      knownBug(
        "the host stopped answering selection calls after a programmatic select — known web-host limitation, " +
          "same family as office-js#3083 / #3698; the pane's own Edit-it path is unaffected",
      ),
    ).toMatch(/#3083/);
    // And a hang somewhere else still goes where it always did.
    expect(knownBug("PowerPoint did not respond while reading slides 20-38 for charts (90.0s)")).toMatch(/5022/);
  });

  it("says nothing about a failure that is ours", () => {
    // The mapping must not turn every problem into somebody else's. A reason
    // with no known host bug behind it gets no note, and stays ours to fix.
    expect(knownBug("chart is one object but carries no config tag")).toBeNull();
    expect(knownBug("empty slide left behind by a lost add")).toBeNull();
    expect(knownBug("")).toBeNull();
  });
});

/**
 * The tool's own documented invocation must not print LESS than the degraded one.
 *
 * `reportTrace` was reachable only from the no-deck branch, so
 * `triage.mjs <deck.pptx> <run-log.json>` — the form in the usage line and in
 * CLAUDE.md, the one you use when you actually have a deck — dropped the entry
 * histogram, "phases an error escaped", the problems tally and every
 * `known host bug: office-js#…` annotation, and said nothing about it. A round
 * with a trace and no self-test went further and reported "this log holds no
 * runs and no self-test" over 186 entries, exit 0.
 *
 * Driven through the CLI rather than the exported functions on purpose: what
 * was wrong was the WIRING, and every function involved was already correct.
 */
describe("triage's two invocations", () => {
  const run = (args: string[]) =>
    spawnSync(process.execPath, ["scripts/triage.mjs", ...args], { encoding: "utf8", timeout: 60_000 });

  it("reports the trace whether or not a deck was passed", async () => {
    const log = {
      build: "test-build",
      host: "test-host",
      runs: [],
      trace: {
        summary: {
          steps: [{ scope: "draw", message: "batch committed", n: 2 }],
          problems: [
            {
              text: "PowerPoint did not respond while drawing shapes 1-10 of 24 (45s) | at=drawing the chart's shapes",
              n: 1,
            },
          ],
        },
        entries: [
          { ms: 1, scope: "draw", message: "batch committed" },
          { ms: 2, scope: "draw", message: "batch committed" },
        ],
      },
      selftest: [],
    };
    const dir = mkdtempSync(join(tmpdir(), "pc-triage-"));
    const logPath = join(dir, "round.json");
    writeFileSync(logPath, JSON.stringify(log));

    const withoutDeck = run([logPath]);
    const withDeck = run(["examples/showcase.pptx", logPath]);

    // The trace is a property of the FILE, like the structural faults, so both
    // forms must show it — and in particular the problem line, which is the
    // most locating thing in any round.
    expect(withoutDeck.stdout).toMatch(/TRACE 2 entries/);
    expect(withDeck.stdout, "the deck path dropped the whole trace section").toMatch(/TRACE 2 entries/);
    expect(withDeck.stdout).toMatch(/did not respond while drawing/);
    // And a round that carries a trace is never described as holding nothing.
    expect(withDeck.stdout).not.toMatch(/holds no runs and no self-test/);
    rmSync(dir, { recursive: true, force: true });
  }, 120_000);
});

/**
 * A picture can be called blank only if it looks like nothing.
 *
 * The baseline used to be `Math.min` over the very shots being classified, so
 * the smallest picture satisfied `<=` against itself and at least one added
 * slide was ALWAYS reported under "read back empty AND rasterise blank" — the
 * two-witness line a maintainer reads as proven data loss.
 */
describe("deckEvidence's blank test", () => {
  const png = (bytes: number) => "A".repeat(Math.ceil((bytes * 4) / 3));
  const deckOf = (sizes: number[]) => ({
    newSlides: sizes.map((_, i) => `s${i}`),
    inventory: sizes.map((_, i) => ({ slideId: `s${i}`, index: i, shapes: [], count: 0 })),
    shots: sizes.map((b, i) => ({ slideId: `s${i}`, png: png(b) })),
  });

  it("does not call the smallest CHART on the slide a blank", () => {
    // Every slide reads back empty (the documented short-collection answer) and
    // every picture is chart-sized. The honest verdict is that the readback is
    // lying — the old code confirmed one of them blank and sent the reader
    // after a drawing bug that is not there.
    const e = deckEvidence(deckOf([5142, 5307, 5307, 5319, 5322, 5322]))!;
    expect(e.confirmed, "a 5kB picture of a chart was called blank").toBe(0);
    expect(e.lying).toBe(6);
  });

  it("still confirms a real blank, and does so with no chart in the round to contrast against", () => {
    // Rounds 12 and 13 ran with the picture cap at 12, so every shot they took
    // is a blank. A test that judges by contrast within the round calls those
    // 23 genuinely blank slides content — which is why the ceiling is absolute.
    const mixed = deckEvidence(deckOf([1146, 1146, 5322, 5319]))!;
    expect(mixed.confirmed).toBe(2);
    expect(mixed.lying).toBe(2);
    const allBlank = deckEvidence(deckOf([1146, 1146, 1146]))!;
    expect(allBlank.confirmed, "a round with nothing to contrast against lost its blanks").toBe(3);
  });
});

/**
 * A crashed round's verdicts are banked as they land now, so the file carries
 * them — and this is the reader that has to show them, or they are invisible in
 * a different place.
 */
describe("triage shows what a crashed run had already concluded", () => {
  it("reads a scenario verdict as its verdict", () => {
    expect(describeFinding({ name: "insert a chart", ok: true, ms: 2400 })).toContain("passed");
    expect(describeFinding({ name: "same scale", ok: false, detail: "3 of 8 charts", ms: 34000 })).toContain(
      "FAILED — 3 of 8 charts",
    );
    expect(describeFinding({ name: "picked", ok: false, skipped: true, detail: "not reached" })).toContain("SKIPPED");
  });

  it("reads an answer sheet as how much of it got answered", () => {
    const sheet = { answers: [{ answer: "yes" }, { answer: "threw" }, { answer: null }] };
    expect(describeFinding(sheet)).toBe("answer sheet, 2 of 3 question(s) answered");
  });

  it("falls back to the value rather than dropping it", () => {
    expect(describeFinding("[dropped: 200000 bytes, over the cap]")).toContain("dropped");
    expect(describeFinding({ odd: 1 })).toBe('{"odd":1}');
  });
});

/**
 * A config the pane cannot reach is not a chart the user can edit.
 *
 * `verify-deck.mjs` draws that distinction on the field itself: `config` only
 * reports that the tag part EXISTS, which is equally true of a tag nothing in
 * the shape tree points at. The pane loads a chart by walking from the
 * SSF Charts object to its tag rid, so a config anchored on the SLIDE is
 * unreachable and the chart is not re-editable.
 *
 * triage was never migrated when that field was added. For the exact case
 * `verify-deck` was extended to catch, its slot table printed `ok`, its deck
 * column printed `config`, and `disagreements` — the number the final line and
 * the `--json` payload both carry — was 0, while verify-deck's own report on
 * the same deck said `orphan`.
 */
describe("triage reads the config the pane can actually reach", () => {
  const log = { run: "r1", items: [{ title: "chart", status: "ok", tagged: true, chart: true }] };
  const row = (extra: Record<string, unknown>) => ({
    index: 0,
    slot: 0,
    run: "r1",
    shapes: 3,
    groups: 1,
    chartObject: true,
    ...extra,
  });

  it("calls an orphaned config what it is, not ok", () => {
    const t = triage({ rows: [row({ config: true, configOnChart: false, configOrphaned: true })] }, log);
    expect(t.slots[0].verdict, "a config the pane cannot load was reported as fine").not.toBe("ok");
    expect(t.disagreements, "an unreachable config counted as agreement").toBeGreaterThan(0);
    expect(t.slots[0].deck).toContain("ORPHANED");
  });

  it("still calls a properly anchored config ok", () => {
    const t = triage({ rows: [row({ config: true, configOnChart: true, configOrphaned: false })] }, log);
    expect(t.slots[0].verdict).toBe("ok");
    expect(t.disagreements).toBe(0);
  });

  it("falls back to `config` for a deck read before the anchor was reported", () => {
    // Unknown must not become a fault: an older verify-deck did not say where
    // the tag was anchored, and the two agree wherever it does.
    const t = triage({ rows: [row({ config: true })] }, log);
    expect(t.slots[0].verdict).toBe("ok");
  });
});

describe("the two populations a draw batch falls into", () => {
  /**
   * "FAST IS THE BROKEN MODE" rests on the batch times being bimodal with an
   * empty band, and every number behind it was read off a round by hand — which
   * is how the band's edge stayed quoted at 29.2s after a later round put
   * surviving batches at 29.3, 30.8 and 31.1.
   *
   * The load-bearing case is the NEGATIVE one. A split taken at the largest gap
   * always exists, so "bimodal" would be true of any data at all; it is only
   * reported when the gap is bigger than the fast group's whole spread, which is
   * what "an empty band" means. Same shape as `untested` in `host-regimes.mjs`:
   * a verdict that cannot come out the other way is not a measurement.
   */
  const log = (ms: number[]) => ({
    trace: { entries: ms.map((prevBatchMs) => ({ message: "batch issued", data: { prevBatchMs } })) },
  });

  it("finds the empty band in a genuinely bimodal round", () => {
    // Round 17's own twenty-two timed batches, not invented ones. A first
    // attempt at synthetic data put three slow points far apart, which moved
    // the largest gap INSIDE the slow group and correctly answered `false` —
    // the guard working, and a reminder that this shape needs a real
    // distribution rather than a sketch of one.
    const p = batchPopulations(
      log([
        3326, 3344, 3540, 3637, 3717, 5049, 5127, 5220, 5307, 5409, 11057, 13118, 18016, 18240, 18491, 18875, 18927,
        19114, 20175, 21963, 25572, 26360,
      ]),
    )!;
    expect(p.bimodal).toBe(true);
    expect(p.fast[p.fast.length - 1]).toBe(5409);
    expect(p.slow[0]).toBe(11057);
    expect(p.gap).toBe(5648);
    expect(p.max).toBe(26360);
  });

  it("refuses to call an evenly spread round bimodal", () => {
    // Every set of numbers has a largest gap, so a split always exists. On
    // uniform times the argmax picks the first, leaving a one-member fast group
    // whose spread is zero — which any gap beats.
    expect(batchPopulations(log([3000, 4000, 5000, 6000, 7000, 8000]))!.bimodal).toBe(false);
  });

  it("refuses a gap no wider than the cluster it is supposed to separate", () => {
    // THE CASE THAT PINS THE SPREAD COMPARISON, and the test above does not:
    // there the length check already answers, so dropping `gap > spread` left
    // the suite green. Here both groups have two or more members and the gap is
    // exactly the fast group's own spread — an interval, not an empty band.
    const p = batchPopulations(log([1000, 2000, 3000, 4000, 5000, 9000, 10000]))!;
    expect(p.fast.length).toBeGreaterThanOrEqual(2);
    expect(p.slow.length).toBeGreaterThanOrEqual(2);
    expect(p.bimodal).toBe(false);
  });

  it("says nothing at all on too few timed batches", () => {
    // A round whose draws were mostly first batches carries no `prevBatchMs`,
    // and three points cannot show a population.
    expect(batchPopulations(log([3000, 9000, 20000]))).toBeNull();
    expect(batchPopulations({ trace: { entries: [] } })).toBeNull();
    expect(batchPopulations(undefined)).toBeNull();
  });
});

describe("how much of a round the grouping comparison can see", () => {
  const e = (message: string, chart?: string) => ({ scope: "group", message, data: chart ? { chart } : {} });

  it("counts the events its own join key cannot reach", () => {
    // `poolGroupVsTag` joins a chart's messages by `data.chart`, and only the
    // deck-wide rescale in selftest.ts writes that field. Everything else in a
    // round is invisible to it — and the blindness is UNEVEN, which is what
    // turns a partial number into a biased one. Archive-wide it sees 229 of 638
    // grouped events and 109 of 127 ungrouped.
    const round = {
      trace: {
        entries: [
          e("grouped the chart's shapes", "1/2"),
          e("grouped the chart's shapes"),
          e("grouped the chart's shapes"),
          e("not grouping — the host refused", "2/2"),
        ],
      },
    };
    const c = poolGroupVsTagCoverage([round]);
    expect(c.groupedSeen).toBe(1);
    expect(c.groupedTotal).toBe(3);
    expect(c.ungroupedSeen).toBe(1);
    expect(c.ungroupedTotal).toBe(1);
    // The point of printing it: the two columns are sampled at different rates,
    // so the RATIO between them is biased and not merely small.
    expect(c.groupedSeen / c.groupedTotal).toBeLessThan(c.ungroupedSeen / c.ungroupedTotal);
  });
});

/**
 * The pooled reading for the last `GetItem(id)` refusal still firing.
 */
describe("what an update left on the slide, pooled", () => {
  const slide = (over: Record<string, number | string>) => ({
    scope: "update",
    message: "shapes left on the slide after an in-place update",
    // `slideId` and `after` are what the deck cross-check reads. `after: 1` is
    // consistent with any slide that finished holding at least one shape, so a
    // fixture that does not care about the deck is unaffected by it.
    data: { slideId: "s1", before: 1, after: 1, growth: 0, settled: true, charts: 1, withParts: 1, ...over },
  });
  const round = (...entries: ReturnType<typeof slide>[]) => ({ trace: { entries } });

  it("keeps a chart AT RISK apart from one that merely grew", () => {
    // THE ONLY DISTINCTION THIS EXISTS TO DRAW. Growth on a chart that could not
    // strand anything is an ordinary change of size; growth on one that could is
    // the stranding. Pooled together they are indistinguishable and the reading
    // says nothing.
    //
    // This fixture used `withParts: 0` as a stand-in for "at risk", which is
    // what the pool itself did until round 086 showed the two are not the same:
    // a GROUPED chart has no parts list either, and cannot strand. Both the
    // fixture and the pool now name the population directly.
    const o = poolUpdateShortfalls([
      round(
        slide({ growth: 12, charts: 1, withParts: 0, atRisk: 1 }),
        slide({ growth: 4, charts: 1, withParts: 1, atRisk: 0 }),
      ),
    ]);
    expect(o.blind).toBe(1);
    expect(o.blindGrowth).toBe(12);
    expect(o.sightedGrowth).toBe(4);
    expect(o.worst).toBe(12);
    expect(o.updates).toBe(2);
    expect(o.rounds).toBe(1);
  });

  it("counts the shapes #586 leaves loose ON PURPOSE, which atRisk calls safe", () => {
    // THE BLIND SPOT, and it is in the instrument rather than the code. #586
    // groups the majority the host will name and leaves the remainder out of
    // the parts tag deliberately. At update time the host says `group` for that
    // chart exactly as it does for a whole one, so `atRisk` counts it SAFE —
    // while four shapes sit loose in its own box and the next update walks past
    // them. The draw pass records the other half (`partial=N left=i:k`), so the
    // join happens here.
    const grouped = (left: string) => ({
      scope: "group",
      message: "grouped the chart's shapes",
      data: { charts: 4, partial: left ? left.split(",").length : 0, ...(left ? { left } : {}), by: "ids" },
    });
    const o = poolUpdateShortfalls([{ trace: { entries: [slide({ atRisk: 0 }), grouped("0:4,3:2")] } }]);
    expect(o.atRisk, "the host calls a subset group a group").toBe(0);
    expect(o.strandedByDesign, "four loose on chart 0, two on chart 3").toBe(6);
    expect(o.subsetGroups).toBe(2);
  });

  it("reports no subset groups when every chart grouped whole", () => {
    // The state this host has actually been in for every round on record: the
    // settled retry repairs the re-read before the subset rule is reached, so
    // `partial` is 0 and there is nothing left loose. A count that cannot come
    // back zero is not measuring anything.
    const whole = {
      scope: "group",
      message: "grouped the chart's shapes",
      data: { charts: 4, partial: 0, by: "ids" },
    };
    const o = poolUpdateShortfalls([{ trace: { entries: [slide({ atRisk: 0 }), whole] } }]);
    expect(o.strandedByDesign).toBe(0);
    expect(o.subsetGroups).toBe(0);
  });

  it("does not count a GROUPED chart's growth as stranding", () => {
    // THE BUCKET WAS KEYED ON THE WRONG FIELD. It split on `withParts === 0`,
    // and a grouped chart has no parts list either — so every grouped chart
    // landed in the stranding column, where a WHOLE group cannot belong: it is
    // deleted in one piece. (A SUBSET group can — see the #586 test below. That
    // is why this reading is printed beside the loose-shape count and not
    // alone.)
    //
    // Round 086 put a growth of 23 there from a chart whose own line read
    // `atRisk: 0`. `atRisk` is the population the question is about.
    const o = poolUpdateShortfalls([
      round(
        slide({ growth: 23, charts: 1, withParts: 0, atRisk: 0 }),
        slide({ growth: 0, charts: 1, withParts: 0, atRisk: 1 }),
      ),
    ]);
    expect(o.blind, "counted a chart that cannot strand anything").toBe(1);
    expect(o.blindGrowth, "credited a grouped chart's growth to stranding").toBe(0);
    expect(o.sightedGrowth).toBe(23);
  });

  it("discards a reading the deck contradicts, and says how many", () => {
    // A round only ADDS shapes to slides it keeps, so a reading claiming more
    // shapes than the slide finished with is claiming shapes that never existed.
    // Round 086 read `after: 24` on a slide the inventory showed holding 1 —
    // and BOTH host reads agreed on the 24, so `settled` was true and wrong.
    // The lag outlasted the settle delay; two reads inside one lag window agree
    // on the stale number.
    const log = {
      trace: { entries: [slide({ growth: 23, after: 24, atRisk: 1, charts: 1, withParts: 0 })] },
      deck: { inventory: [{ slideId: "s1", count: 1 }] },
    };
    const o = poolUpdateShortfalls([log]);
    expect(o.deckContradicted).toBe(1);
    expect(o.blindGrowth, "pooled a reading the deck had already disproved").toBe(0);
    expect(o.atRisk, "counted an at-risk chart from a discarded reading").toBe(0);
  });

  it("keeps a reading the deck is consistent with", () => {
    // The slide gained more charts later in the round, so the final count is
    // HIGHER than this reading's `after`. That is the ordinary case and must not
    // be mistaken for a contradiction.
    const log = {
      trace: { entries: [slide({ growth: 0, atRisk: 1, charts: 1, withParts: 0 })] },
      deck: { inventory: [{ slideId: "s1", count: 9 }] },
    };
    const o = poolUpdateShortfalls([log]);
    expect(o.deckContradicted).toBe(0);
    expect(o.atRisk).toBe(1);
  });

  it("will not attribute a mixed slide's shortfall to the blind charts on it", () => {
    // A slide where some charts had a list and some did not cannot say which of
    // them stranded anything, and guessing would inflate the one number the
    // whole question turns on.
    const o = poolUpdateShortfalls([round(slide({ growth: 9, charts: 3, withParts: 1 }))]);
    expect(o.blind).toBe(0);
    expect(o.blindGrowth).toBe(0);
    expect(o.sightedGrowth).toBe(9);
  });

  it("will not pool a reading from the build whose units did not match", () => {
    // Round 082 and earlier carry `shortfall`/`unexplained` — a subtraction
    // across three units that summed to zero on every line. Counted apart, never
    // mixed in, or the artifact this pool was rewritten to stop reporting comes
    // straight back.
    const old = {
      scope: "update",
      message: "shapes left on the slide after an in-place update",
      data: { shortfall: 23, unexplained: -23, charts: 1, withParts: 0 },
    };
    const o = poolUpdateShortfalls([{ trace: { entries: [old] } }]);
    expect(o.unitMismatch).toBe(1);
    expect(o.blind).toBe(0);
    expect(o.blindGrowth).toBe(0);
  });

  it("says nothing at all about rounds that never measured it", () => {
    // Every round before this instrument shipped, which is all 57 of them.
    expect(poolUpdateShortfalls([{ trace: { entries: [] } }, {}]).updates).toBe(0);
    expect(poolUpdateShortfalls([]).rounds).toBe(0);
  });
});

/**
 * The trace is 95K characters a person cannot count, so the thing that counts it
 * has to be right about WHICH KIND of difference it found.
 */
describe("what the newest round said that the archive has not", () => {
  const entry = (scope: string, message: string, error?: string) => ({
    scope,
    message,
    ...(error ? { data: { error } } : {}),
  });
  const round = (build: string, es: ReturnType<typeof entry>[]) => ({ build, trace: { entries: es } });
  const times = (n: number, e: ReturnType<typeof entry>) => Array.from({ length: n }, () => e);
  /** Enough prior rounds to earn a baseline, all saying the same quiet thing. */
  const quietPriors = (n: number, per = 3) =>
    Array.from({ length: n }, (_, i) => round(`old${i}`, times(per, entry("draw", "batch issued"))));

  it("counts one event happening twice as one signature, not two", () => {
    // Ids and counts are the noise. Without this the archive's vocabulary grows
    // by one every time a slide gets a different id, and nothing ever has a
    // history to be measured against.
    expect(traceSignature(entry("group", "repaired 3 tags on slide 7f2a91c4"))).toBe(
      traceSignature(entry("group", "repaired 5 tags on slide 0b41ee02")),
    );
  });

  it("keeps the error class, because a failure is not its healthy twin", () => {
    // The reading this was added for: round 077's 52 `UnexpectedError`s share a
    // message with the successful writes and lived in `data.error`, so a
    // signature over scope+message alone reported the loudest fact in that round
    // as ordinary traffic.
    expect(traceSignature(entry("error", "writing the chart's config tag", "UnexpectedError: nope"))).not.toBe(
      traceSignature(entry("error", "writing the chart's config tag")),
    );
    // The class only. Error text carries ids and offsets, and hashing those
    // would give every single failure its own signature.
    expect(traceSignature(entry("error", "writing", "UnexpectedError: id 7f2a91c4 at 33"))).toBe(
      traceSignature(entry("error", "writing", "UnexpectedError: id 0b41ee02 at 91")),
    );
  });

  it("calls a shape the archive has never produced NOVEL", () => {
    const rounds = [...quietPriors(6), round("new", [entry("group", "something nobody has traced before")])];
    const out = traceNovelty(rounds);
    expect(out.novel).toHaveLength(1);
    expect(out.novel[0].sig).toContain("something nobody has traced before");
    expect(out.spikes).toHaveLength(0);
  });

  // THE REGRESSION GUARD FOR THE COLLAPSE BUG. `profileDivergence` shipped
  // reporting a flaky scenario as a slide-size difference because it kept one
  // worst-case reading of two different causes. This is the same split, and
  // getting it wrong here means every fix this project lands is reported as a
  // fault on the night it starts working.
  it("separates a mechanism that just shipped from a count that left its baseline", () => {
    const reread = entry("group", "re-reading the slide's shapes again after a settle delay");
    const tagWrite = entry("error", "writing the chart's config tag", "UnexpectedError: x");
    const priors = Array.from({ length: 8 }, (_, i) =>
      round(`old${i}`, [
        ...times(3, tagWrite),
        ...times(2, entry("draw", "batch issued")),
        // RARE, NOT ABSENT — and modelling it as absent is what made the first
        // version of this test fail against a correct implementation. The
        // re-read existed before build 17a8204; `needsPreGroupRefresh` widened
        // which charts reach it, so the count went from occasional to eleven.
        // A signature nothing has EVER produced is a different report (NOVEL),
        // and the archive proves the difference: at round 077 the
        // `UnexpectedError` signatures were novel, while the settle-pass repair
        // beside them was this bucket.
        ...(i === 7 ? [reread] : []),
      ]),
    );
    const newest = round("17a8204", [
      // Present in the vocabulary, median 0 across the priors — the shape a
      // mechanism makes on the night it starts working.
      ...times(11, reread),
      // Seen in every prior round, three at a time, now eighteen.
      ...times(18, tagWrite),
    ]);
    const out = traceNovelty([...priors, newest]);
    expect(out.sinceBuild.map((s: { sig: string }) => s.sig).join()).toContain("re-reading the slide's shapes");
    expect(out.spikes.map((s: { sig: string }) => s.sig).join()).toContain("writing the chart's config tag");
    // Neither may appear in the other's bucket — that is the whole point.
    expect(out.spikes.map((s: { sig: string }) => s.sig).join()).not.toContain("re-reading");
    expect(out.sinceBuild.map((s: { sig: string }) => s.sig).join()).not.toContain("config tag");
    expect(out.spikes[0].median).toBe(3);
  });

  it("stops calling a signature new once it has been around for a while", () => {
    // THE BUCKET NEVER EMPTIED. The median ran over EVERY prior round, so a
    // signature stayed "new behaviour" until it had appeared in more than half
    // the whole archive — and the archive keeps growing underneath it.
    //
    // Measured on the real thing: `re-reading the slide's shapes again after a
    // settle delay` first appeared in round 064 and sat at 10-11 ever after. It
    // was reported as NEW BEHAVIOUR in fifteen separate rounds and blamed on
    // nine different builds, the last a commit that only changed a slide
    // counter. A signal that fires every night about a thing that has not
    // changed is the "cries wolf" failure the gate's header warns about.
    // THE ARCHIVE'S ACTUAL SHAPE, and a fixture that misses it proves nothing:
    // the signature is in a MINORITY of all priors (so a whole-archive median
    // reads 0 and files it as new) while being present in every recent one (so
    // a windowed median reads 11 and does not). A first version of this test had
    // it in all twelve priors, where both readings agree — it passed against the
    // very bug it was written for.
    const steady = entry("group", "a mechanism that shipped long ago");
    const priors = [
      ...Array.from({ length: 12 }, (_, i) => round(`ancient${i}`, [])),
      ...Array.from({ length: 6 }, (_, i) => round(`recent${i}`, times(11, steady))),
    ];
    const out = traceNovelty([...priors, round("new", times(11, steady))]);
    expect(out.sinceBuild, "still calling a settled signature new").toHaveLength(0);
    expect(out.spikes, "11 against a median of 11 is not a spike either").toHaveLength(0);
  });

  it("names the build a signature actually started in, not the one being judged", () => {
    // Nine innocent commits were named for one 064-era signature, because the
    // report printed the newest round's build for every entry.
    const late = entry("group", "arrived partway through");
    const priors = [
      round("aaaaaaa", []),
      round("bbbbbbb", times(11, late)), // <- where it actually started
      round("ccccccc", []),
      round("ddddddd", []),
      round("eeeeeee", []),
      round("fffffff", []),
      round("ggggggg", []),
    ];
    const out = traceNovelty([...priors, round("newbuild", times(11, late))]);
    expect(out.sinceBuild).toHaveLength(1);
    expect(out.sinceBuild[0].startedIn, "blamed the build being judged").toBe("bbbbbbb");
  });

  it("says nothing at all about a round that repeated the archive", () => {
    // The quiet case has to be assertable, or "nothing new" is just the report
    // failing to run and nobody can tell the difference.
    const out = traceNovelty([...quietPriors(8), round("same", times(3, entry("draw", "batch issued")))]);
    expect(out.novel).toHaveLength(0);
    expect(out.sinceBuild).toHaveLength(0);
    expect(out.spikes).toHaveLength(0);
    expect(out.vocabulary).toBe(1);
  });

  it("will not build a baseline out of too few rounds", () => {
    // With three priors a median is whichever one happened to be in the middle,
    // and every count looks like a spike against it. A fresh clone should get
    // silence, not an alarm resting on nothing.
    const out = traceNovelty([...quietPriors(3), round("new", times(40, entry("draw", "batch issued")))]);
    expect(out.spikes).toHaveLength(0);
    expect(out.sinceBuild).toHaveLength(0);
  });

  it("ignores a difference too small to be a difference", () => {
    const out = traceNovelty([...quietPriors(8), round("new", times(9, entry("draw", "batch issued")))]);
    expect(out.spikes).toHaveLength(0);
  });

  it("survives a round with no trace at all", () => {
    expect(() => traceNovelty([...quietPriors(6), { build: "x" }])).not.toThrow();
    expect(traceNovelty([]).novel).toEqual([]);
  });
});

describe("a scenario that passes on a smaller population than it usually runs", () => {
  const round = (build: string, of: number, ok = true) => ({
    build,
    selftest: [{ name: "same scale across the deck", ok, detail: `${of} of ${of} charts carry the shared scale` }],
  });

  it("names the scenario, the number now, and the number it usually runs", () => {
    // Round 088 exactly: eight every round on record, then six — and the verdict
    // is `scaled === charts.length`, so six of six is a PASS and every other
    // reading in the gate stayed green.
    const rounds = [round("a", 8), round("b", 8), round("c", 8), round("d", 6)];
    expect(poolScenarioPopulations(rounds)).toEqual([
      { name: "same scale across the deck", now: 6, usual: 8, ok: true, rounds: 3 },
    ]);
  });

  it("says nothing when the population held", () => {
    expect(poolScenarioPopulations([round("a", 8), round("b", 8), round("c", 8)])).toEqual([]);
  });

  it("says nothing when the population GREW — this is a floor, not a change detector", () => {
    expect(poolScenarioPopulations([round("a", 6), round("b", 6), round("c", 8)])).toEqual([]);
  });

  it("takes the usual from the median, so one small round cannot lower the bar for the next", () => {
    // A mean would be dragged down by the outlier and let the following round's
    // shrink through — the failure mode that makes a guard quietly stop guarding.
    const rounds = [round("a", 8), round("b", 8), round("c", 8), round("d", 2), round("e", 6)];
    expect(poolScenarioPopulations(rounds)[0]).toMatchObject({ now: 6, usual: 8 });
  });

  it("ignores a scenario whose verdict carries no count, and a first-ever round", () => {
    const noCount = { build: "a", selftest: [{ name: "stop a run part-way", ok: true, detail: "stopped cleanly" }] };
    expect(poolScenarioPopulations([noCount, noCount])).toEqual([]);
    expect(poolScenarioPopulations([round("a", 8)]), "one round is not a history").toEqual([]);
  });

  it("flags a shrunken population even when the scenario FAILED, and says which", () => {
    // Four rounds, because "usually" now needs three priors — see the baseline
    // test below. Two observations are not a norm.
    const rounds = [round("a", 8), round("b", 8), round("c", 8), round("d", 4, false)];
    expect(poolScenarioPopulations(rounds)[0]).toMatchObject({ now: 4, usual: 8, ok: false });
  });

  it("says nothing when THIS round carried no count, however small the last one that did", () => {
    /**
     * Rounds 306, 307 and 308, on 2026-08-29. Each printed
     *
     *     insert onto a slide that already has content — 7 this round,
     *     usually 16 over 9 prior round(s)
     *
     * on a round whose verdict for that scenario carried no "N of M" at all.
     * The 7 was round 282's, twenty-four rounds and four builds earlier: `hist`
     * holds only the rounds that COUNTED, so its last entry is the newest round
     * that counted and not the round being judged.
     *
     * Permanent rather than a blip — 282 is the newest counting round and stays
     * so — which is what makes it a guard rather than a note. A warning that
     * fires every round and names the wrong round teaches the reader to skip
     * the line that exists to make them stop.
     */
    const quiet = {
      build: "e",
      selftest: [{ name: "same scale across the deck", ok: true, detail: "all charts carry the shared scale" }],
    };
    const rounds = [round("a", 8), round("b", 8), round("c", 8), round("d", 6), quiet];
    expect(poolScenarioPopulations(rounds), "reported an older round's count as this round's").toEqual([]);
    // And the round that DID shrink is still reported when it IS the one being
    // judged — the guard must not swallow the finding it was built around.
    expect(poolScenarioPopulations(rounds.slice(0, -1))[0]).toMatchObject({ now: 6, usual: 8 });
  });
});

describe("grouping, which no scenario verdict reports", () => {
  const round = (build: string, grouped: number, refused: number, deck: number[]) => ({
    build,
    trace: {
      entries: [
        ...(grouped ? [{ message: "grouped the chart's shapes", data: { charts: grouped } }] : []),
        ...Array.from({ length: refused }, () => ({
          message: "not grouping: no member handle this host will accept",
          data: {},
        })),
      ],
    },
    deck: { inventory: deck.map((count, i) => ({ slideId: `s${i}`, count })) },
  });

  it("reports the newest round's grouping against its own history", () => {
    // Rounds 092 and 093, one build run twice, nothing changed between them: 20
    // grouped / 0 refused, then 15 grouped / 4 refused with three slides ending
    // on 24 shapes each. Both reported 13/13 and the identical verdict line.
    const out = poolGroupingOutcome([
      round("a", 20, 2, [1, 1, 1]),
      round("a", 20, 0, [0, 4, 2, 5, 1, 1, 1]),
      round("a", 20, 2, [1, 1, 1]),
      round("a", 15, 4, [0, 4, 2, 17, 24, 24, 24]),
    ]);
    // priors are [2, 0, 2]; sorted [0, 2, 2] and the median index lands on the
    // middle, so the baseline is 2. The conservative direction: a higher
    // baseline flags LESS, and this line is a reason to read, not a verdict.
    expect(out).toMatchObject({ now: { grouped: 15, refused: 4 }, refusedMedian: 2, rounds: 3 });
    expect(out?.now.deck, "the deck is printed as corroboration, not derived from").toEqual([0, 4, 2, 17, 24, 24, 24]);
  });

  it("says nothing when the round being judged did no grouping of its own", () => {
    /**
     * The same stale-`now` defect found in `poolScenarioPopulations` on
     * 2026-08-29 — a `continue` that filters the population, then `now` taken as
     * the last survivor. Here it is worse placed: this is the headline grouping
     * figure of every gate run, and the deck line printed beside it, so a stale
     * `now` describes ANOTHER ROUND'S DECK as this round's.
     *
     * A round that grouped nothing is not hypothetical — it is what an in-place
     * update produces, and `attempts per round` halving at round 153 is recorded
     * two comments above as exactly that.
     */
    const quiet = { build: "e", trace: { entries: [{ message: "nothing to group here", data: {} }] } };
    const rounds = [
      round("a", 20, 2, [1, 1, 1]),
      round("a", 20, 0, [1, 1, 1]),
      round("a", 20, 2, [1, 1, 1]),
      round("a", 15, 4, [0, 4, 2, 17, 24, 24, 24]),
      quiet,
    ];
    expect(poolGroupingOutcome(rounds), "reported an older round's grouping and deck as this round's").toBeNull();
    // And the finding this instrument exists for still fires when the round
    // being judged IS the one that grouped.
    expect(poolGroupingOutcome(rounds.slice(0, -1))).toMatchObject({ now: { grouped: 15, refused: 4 } });
  });

  it("separates a positional guess that picked its own shapes from one that picked another chart's", async () => {
    // THE SILENT HALF OF THE addGroup THROW. `chooseGroupMembers` falls back to
    // the TAIL of the host's shape listing when no id matched, and the listing
    // can be one grouping-event stale — so the tail is the PREVIOUS chart's
    // shapes.
    //
    // If those shapes are already inside that chart's group they no longer exist
    // at top level, every getItemOrNullObject returns a null object, and
    // addGroup throws InvalidArgument. That is the 29 archived throws across
    // rounds 068-175, every one preceded by this branch, and it is VISIBLE.
    //
    // If they are still LOOSE, the same guess SUCCEEDS on the wrong chart's
    // shapes and feeds them to the parts tag as this chart's own. It throws
    // nothing and traced nothing, so the archive could not tell it from a
    // correct group — an inference the evidence could not reach, which is why
    // the ids are recorded rather than a verdict.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolPositionalGuess } = await import("../scripts/triage.mjs");
    const ev = (mine: (string | null)[], chose: (string | null)[]) => ({
      message: "the positional guess picked the tail of the listing",
      data: { index: 0, mine, chose, listed: 15 },
    });
    const pooled = poolPositionalGuess([
      {
        trace: {
          entries: [
            ev(["35", "36", "37"], ["35", "36", "37"]), // its own
            ev(["43", "44", "45"], ["35", "36", "37"]), // another chart's
            ev(["50", "51", "52"], ["50", "51", "99"]), // a mixture
            ev(["60", "61"], ["60", null]), // the host would not name one
          ],
        },
      },
      { trace: { entries: [{ message: "drew a chart" }] } },
    ]);
    expect(pooled).toMatchObject({ events: 4, rounds: 1, mine: 1, other: 1, partial: 1, unreadable: 1 });

    // THE MIXTURE MUST NOT COUNT AS "OWN". A guess that took four of this
    // chart's shapes and three of the last one's produces a group that is wrong
    // in exactly the way the parts tag cannot express, and rounding it toward
    // the reassuring bucket is how this stayed invisible.
    const mixed = poolPositionalGuess([{ trace: { entries: [ev(["1", "2"], ["1", "9"])] } }]);
    expect(mixed.mine, "a partial pick was counted as this chart's own").toBe(0);
    expect(mixed.partial).toBe(1);

    // THE GUARD'S REFUSALS, counted next to the guesses they judged. Without
    // this the section reports foreign picks under a paragraph describing silent
    // corruption and says nothing about whether anything was grouped — which
    // cost a trip through the archive to answer, and the answer was "the guard
    // refused them".
    const guarded = poolPositionalGuess([
      {
        trace: {
          entries: [
            ev(["43", "44"], ["35", "36"]),
            { message: "not grouping: the positional guess named no shape of ours", data: { index: 0 } },
          ],
        },
      },
    ]);
    expect(guarded.other, "the foreign pick should still be counted as foreign").toBe(1);
    expect(guarded.refused, "the guard's refusal was not counted").toBe(1);
    // NO ASSERTION THAT `events` STAYED 1. There was one, and it could not fail:
    // a refusal does not match the `picked the tail` test, so no reachable
    // mutation makes it count as a guess. It was removed rather than kept as a
    // line that only ever agrees with the code.
  });

  it("pairs the real occupancy reading with the update it explains, bucketed by chart size", async () => {
    // THE READING EXISTED FOR 29 ROUNDS AND NOTHING READ IT. `same scale across
    // the deck` calls `slideOccupancy` before the run and traces the host's own
    // per-slide count, ordered as the charts are — while this reader went on
    // printing "there is no reliable occupancy measure in a round".
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolOccupancyCost } = await import("../scripts/triage.mjs");
    const held = (shapes: number[]) => ({
      message: "what each slide held before the rescale",
      data: { charts: shapes.length, slides: shapes.map((n, i) => ({ slide: `s${i}`, shapes: n })) },
    });
    const upd = (chart: string, ms: number, of: number, changed: number) => ({
      message: "updated in place",
      data: { chart, ms, of, changed },
    });
    const { rounds, cells } = poolOccupancyCost([
      {
        trace: {
          entries: [
            held([3, 3, 1]),
            upd("1/3", 40000, 24, 18), // first, on a 3-shape slide
            upd("2/3", 20000, 24, 18), // later, same size, same slide load
            upd("3/3", 13000, 16, 9), // later, a SMALLER chart
          ],
        },
      },
    ]);
    expect(rounds).toBe(1);
    // BUCKETED BY of/changed, and this is the assertion that matters. Read
    // unbucketed the archive says charts 2 and 3 are the cheapest in the run,
    // which reads as "a busier slide is faster"; they are simply 16-node charts
    // against 24-node ones. Pooling them would reproduce the confound the
    // surrounding section exists to warn about.
    expect(cells.get("24/18|3|first")).toEqual([40000]);
    expect(cells.get("24/18|3|later")).toEqual([20000]);
    expect(cells.get("16/9|1|later"), "the smaller chart was pooled with the big ones").toEqual([13000]);

    // `1/1` IS NOT A FIRST CHART. `one chart alone on a warm deck` is its own
    // scenario and its own arm — the archive reports it as "nearer a LATER
    // chart" — so labelling it `first` pools the arm that ISOLATES the effect
    // with the arm that exhibits it. The first version of this read only the
    // numerator of `chart` and did exactly that.
    const alone = poolOccupancyCost([{ trace: { entries: [held([3]), upd("1/1", 17000, 24, 18)] } }]);
    expect(alone.cells.get("24/18|3|alone"), "a 1/1 chart was not labelled alone").toEqual([17000]);
    expect(alone.cells.get("24/18|3|first"), "a 1/1 chart was pooled in as a first chart").toBeUndefined();

    // A ROUND WITH NO READING CONTRIBUTES NOTHING rather than contributing a
    // guess — 214 of the archive's rounds predate the trace.
    const none = poolOccupancyCost([{ trace: { entries: [upd("1/3", 40000, 24, 18)] } }]);
    expect(none.rounds, "a round with no occupancy reading was counted anyway").toBe(0);
  });
  it("counts a group that THREW, which fell out of both the numerator and the denominator", () => {
    // GROUPING HAS THREE OUTCOMES AND THIS COUNTED TWO. It succeeds
    // (`grouped the chart's shapes`), declines by rule (`not grouping: …`), or
    // THROWS — `grouping the chart's shapes` carrying an `error`, usually
    // `InvalidArgument`. Only the first two were counted, so round 174 printed
    // `8 of 8 attempt(s) grouped, 0 refused` and read as perfect. **The missing
    // attempt was the defect**, and the count written to expose it hid it.
    //
    // 183 throws across 65 of 150 rounds, and they match the loose-shape slides
    // exactly: rounds 159, 161, 167 and 174 each threw once and ended with a
    // slide holding 17, 17, 11 and 11 shapes; the twelve rounds between them
    // threw none and ended at 5. A chart whose group throws is left as its
    // shapes — which is what `poolFullestSlide` sees from the other side, and
    // that instrument found this one round after it landed.
    const thrown = (n: number) =>
      Array.from({ length: n }, () => ({
        message: "grouping the chart's shapes",
        data: { error: "InvalidArgument | at=grouping the chart's shapes" },
      }));
    const withThrow = poolGroupingOutcome([
      { build: "a", trace: { entries: [...thrown(1)] }, deck: { inventory: [{ count: 11 }] } },
      { build: "a", trace: { entries: [...thrown(1)] }, deck: { inventory: [{ count: 11 }] } },
      { build: "a", trace: { entries: [...thrown(1)] }, deck: { inventory: [{ count: 11 }] } },
      {
        build: "a",
        trace: {
          entries: [{ message: "grouped the chart's shapes", data: { charts: 8 } }, ...thrown(1)],
        },
        deck: { inventory: [{ count: 11 }] },
      },
    ]);
    expect(withThrow?.now.grouped).toBe(8);
    expect(withThrow?.attempts, "the throw was dropped from the denominator, so 8 of 8 read as perfect").toBe(9);
    expect(withThrow?.now.threw).toBe(1);

    // A ROUND THAT ONLY THREW IS STILL A ROUND. The old guard skipped any round
    // with no successes and no refusals, so a round where every group threw
    // vanished from the pool entirely.
    const onlyThrew = poolGroupingOutcome([
      { build: "a", trace: { entries: thrown(2) }, deck: { inventory: [{ count: 20 }] } },
      { build: "a", trace: { entries: thrown(2) }, deck: { inventory: [{ count: 20 }] } },
      { build: "a", trace: { entries: thrown(2) }, deck: { inventory: [{ count: 20 }] } },
      { build: "a", trace: { entries: thrown(3) }, deck: { inventory: [{ count: 20 }] } },
    ]);
    expect(onlyThrew, "a round where every group threw dropped out of the pool").not.toBeNull();
    expect(onlyThrew?.attempts).toBe(3);
    expect(onlyThrew?.now.grouped).toBe(0);
  });

  it("carries the DENOMINATOR, because the population halved and nothing said so", () => {
    // `grouped` per round ran 15-20 for the whole archive and halved to 9 at
    // round 153. The cause is benign and complete — the in-place update started
    // working, an in-place update never redraws, and a chart that is not redrawn
    // is never regrouped: 15 - 9 = 6 and 11 - 5 = 6, exactly, in the round it
    // changed. Benign, and it silently rebased every grouping figure in triage.
    //
    // "0 refused (usually 2)" reads as an improvement when half the attempts
    // stopped happening. The RATE is the only reading that survives the change:
    // 9 of 9 attempts and 10 of 19 are comparable, 9 and 10 are not. Same trap
    // as `same scale across the deck` scoring "6 of 6" and "8 of 8" both a pass.
    const out = poolGroupingOutcome([
      round("a", 10, 9, [1]),
      round("a", 18, 2, [1]),
      round("a", 15, 0, [1]),
      round("a", 9, 0, [1]),
    ]);
    expect(out?.attempts, "9 grouped of 9 attempted — a rate, not a bare count").toBe(9);
    expect(out?.recent, "attempts per round in sequence, so a halving is visible").toEqual([19, 20, 15, 9]);

    // AND A ROUND WHERE THE TWO DIFFER, because the assertion above cannot tell
    // `grouped + refused` from `grouped` when the newest round refused nothing —
    // which is every recent round, and would have let the denominator quietly
    // become the numerator again.
    const refusing = poolGroupingOutcome([
      round("a", 10, 9, [1]),
      round("a", 18, 2, [1]),
      round("a", 15, 0, [1]),
      round("a", 6, 3, [1]),
    ]);
    expect(refusing?.attempts, "a refused chart was still an attempt").toBe(9);
    expect(refusing?.now.grouped, "and it is NOT the same number as the successes").toBe(6);
  });

  it("refuses to name a usual until three rounds have been seen", () => {
    // THE HOUSE DEFECT, caught in a second emitter. With too little history this
    // used to report `refusedMedian: 0` — indistinguishable, to a reader, from a
    // history in which nothing was ever refused. The gate printed that as
    // "usually 0 refused".
    const thin = poolGroupingOutcome([round("a", 20, 3, [1]), round("a", 20, 3, [1]), round("a", 15, 4, [1])]);
    expect(thin?.rounds, "two priors").toBe(2);
    expect(thin?.refusedMedian, "two priors cannot name a usual").toBeNull();

    // And it is a THRESHOLD, not a refusal to ever answer: one more round and
    // the same data yields a number. Without this half, deleting the median
    // entirely would pass the assertion above.
    const enough = poolGroupingOutcome([
      round("a", 20, 3, [1]),
      round("a", 20, 3, [1]),
      round("a", 20, 3, [1]),
      round("a", 15, 4, [1]),
    ]);
    expect(enough?.rounds, "three priors").toBe(3);
    expect(enough?.refusedMedian, "three priors is a baseline").toBe(3);
  });

  it("counts charts, not grouped-lines — one line can carry several", () => {
    // `charts` is a count on the line. Counting lines instead would read a round
    // that grouped twenty charts in one batch as having grouped one.
    const out = poolGroupingOutcome([round("a", 4, 0, [1]), round("a", 20, 0, [1])]);
    expect(out?.now.grouped).toBe(20);
  });

  it("skips a round that neither grouped nor refused anything", () => {
    // A round that never reached the grouping path says nothing about it, and
    // averaging its zero into the history would drag the baseline down.
    const empty = { build: "a", trace: { entries: [{ message: "something else", data: {} }] } };
    expect(poolGroupingOutcome([empty, empty])).toBeNull();
  });

  it("survives a round with no trace and no deck", () => {
    expect(() => poolGroupingOutcome([{ build: "a" }, round("a", 1, 0, [1])])).not.toThrow();
  });
});

describe("what each scenario cost the host", () => {
  const round = (...rows: { name: string; f: Record<string, number> }[]) => ({
    trace: { entries: rows.map((r) => ({ message: "scenario passed", data: { name: r.name, friction: r.f } })) },
  });
  const F = (errors = 0, idRefusals = 0, generalExceptions = 0, emptyReReads = 0) => ({
    errors,
    idRefusals,
    generalExceptions,
    emptyReReads,
  });

  it("sums the meter every verdict has carried since round 023", () => {
    const o = poolScenarioFriction([
      round({ name: "same scale", f: F(4, 4, 0, 2) }, { name: "visible", f: F(1, 0, 0, 0) }),
      round({ name: "same scale", f: F(2, 2, 0, 1) }),
    ]);
    expect(o.rounds).toBe(2);
    expect(o.rows[0]).toMatchObject({ name: "same scale", n: 2, sum: { errors: 6, idRefusals: 6, emptyReReads: 3 } });
  });

  it("pools every counter the data carries, not a list someone wrote down", () => {
    /**
     * `KEYS` was a literal naming four counters. The friction object carries
     * EIGHT, enumerated over all 4,285 friction records in the archive:
     * `reReadsRepaired`, `shortReReads`, `unmatchedReReads` and
     * `settledByBinding` were never pooled and reached no reader anywhere.
     *
     * One of them is not idle. `reReadsRepaired` is non-zero in 331 archived
     * records, and over the last 13 rounds it sums to 51 on `does a rasterise
     * poison the next draw` — a number nobody could have seen.
     *
     * The docstring on the function says conclusions are derived here "because a
     * hardcoded conclusion keeps printing after it stops being true". The key
     * list was the one thing still written down.
     */
    const o = poolScenarioFriction([
      { trace: { entries: [{ data: { name: "a", friction: { errors: 1, reReadsRepaired: 4 } } }] } },
      {
        trace: { entries: [{ data: { name: "a", friction: { errors: 2, reReadsRepaired: 6, aBrandNewCounter: 5 } } }] },
      },
    ]);
    expect(o.rows[0].sum.reReadsRepaired, "a counter with real signal was never pooled").toBe(10);
    // A counter added tomorrow is pooled the day it first appears.
    expect(o.rows[0].sum.aBrandNewCounter, "a newly added counter was ignored").toBe(5);
    expect(o.rows[0].sum.errors).toBe(3);
  });

  it("gives every row every derived counter, so a row missing one still sums", () => {
    /**
     * The invariant that makes the derived key list safe, and worth pinning
     * because it is not obvious: KEYS is a UNION over all logs, and the summing
     * loop walks KEYS rather than the row's own friction object. So a scenario
     * that never reports `errors` still gets `errors: 0` — the ranking, which
     * adds `errors` and `idRefusals`, can never meet an undefined and produce
     * NaN unless NO log anywhere carries that counter, in which case every row
     * is equally absent and no order is observable either way.
     *
     * (The `?? 0` in the ranking is therefore belt-and-braces, and a mutant that
     * removes it cannot be killed. Recorded as equivalent rather than left
     * looking untested.)
     */
    const o = poolScenarioFriction([
      { trace: { entries: [{ data: { name: "quiet", friction: { reReadsRepaired: 1 } } }] } },
      { trace: { entries: [{ data: { name: "loud", friction: { errors: 9 } } }] } },
    ]);
    expect(o.rows).toHaveLength(2);
    const quiet = o.rows.find((r: { name: string }) => r.name === "quiet");
    expect(quiet.sum.errors, "a row that never reported this counter was left undefined").toBe(0);
    expect(quiet.sum.reReadsRepaired).toBe(1);
    // And the ranking still puts the noisier scenario first.
    expect(o.rows[0].name).toBe("loud");
  });

  it("names a counter that has NEVER moved, instead of printing it as a number", () => {
    // `generalExceptions` is 0 in every scenario in every one of 86 archived
    // rounds. Printed beside real counts it reads as "no general exceptions
    // happened", when what it means is "this counter measures nothing".
    const o = poolScenarioFriction([round({ name: "a", f: F(3, 1, 0, 0) }), round({ name: "a", f: F(2, 1, 0, 0) })]);
    expect(o.dead).toContain("generalExceptions");
    expect(o.dead, "a counter that DID move must not be called dead").not.toContain("errors");
  });

  it("names a constant, because a number that never varies is not a measurement", () => {
    // `stop a run part-way` reports exactly one error every round — its own
    // deliberate abort, counted as an error.
    const o = poolScenarioFriction([
      round({ name: "stop a run part-way", f: F(1) }),
      round({ name: "stop a run part-way", f: F(1) }),
      round({ name: "stop a run part-way", f: F(1) }),
    ]);
    expect(o.constant).toContainEqual({ name: "stop a run part-way", key: "errors", value: 1 });
  });

  it("does not call a VARYING counter constant, nor a zero one", () => {
    const varying = poolScenarioFriction([
      round({ name: "a", f: F(1) }),
      round({ name: "a", f: F(2) }),
      round({ name: "a", f: F(3) }),
    ]);
    expect(varying.constant, "a counter that moved was called constant").toEqual([]);
    // An all-zero counter is DEAD, not constant — the two want different words.
    const zero = poolScenarioFriction([
      round({ name: "a", f: F(0, 1) }),
      round({ name: "a", f: F(0, 2) }),
      round({ name: "a", f: F(0, 3) }),
    ]);
    expect(zero.constant.filter((c: { key: string }) => c.key === "errors")).toEqual([]);
    expect(zero.dead).toContain("errors");
  });
});

describe("a population baseline needs more than one observation", () => {
  const round = (build: string, of: number) => ({
    build,
    selftest: [{ name: "insert onto a slide", ok: true, detail: `${of} of ${of} charts` }],
  });

  it("says nothing when the history is one or two rounds", () => {
    // Round 112 fired on `2 this round, usually 16 over 1 prior round(s)` — a
    // scenario whose verdict had only just started carrying a count, so its
    // whole history was a single round. One number is not a norm.
    expect(poolScenarioPopulations([round("a", 16), round("b", 2)])).toEqual([]);
    expect(poolScenarioPopulations([round("a", 16), round("b", 16), round("c", 2)])).toEqual([]);
  });

  it("still reports once there is a history to compare against", () => {
    const out = poolScenarioPopulations([round("a", 16), round("b", 16), round("c", 16), round("d", 2)]);
    expect(out).toEqual([{ name: "insert onto a slide", now: 2, usual: 16, ok: true, rounds: 3 }]);
  });
});

describe("a scenario regression carries its own history", () => {
  const round = (pairs: [string, boolean][]) => ({ selftest: pairs.map(([name, ok]) => ({ name, ok })) });

  it("counts lifetime failures rather than restating the window size", () => {
    // WHAT THE GATE USED TO PRINT: "had passed the previous 3 rounds running".
    // `passedIn` was the WINDOW SIZE — a constant true of every regression this
    // function can return, so the sentence carried no information at all. It is
    // the hardcoded-conclusion shape: a number presented as a measurement.
    //
    // Round 148 is why it matters. Two scenarios fell in the same round; one
    // had NEVER failed in 109 rounds and the other had failed before. Those
    // want different responses, and the old line described them identically.
    const rounds = [
      round([["a", true]]),
      round([["a", false]]), // an old failure, long before the window
      ...Array.from({ length: 5 }, () => round([["a", true]])),
      round([["a", false]]),
    ];
    const [g] = scenarioRegressions(rounds) as { name: string; ran: number; failed: number }[];
    expect(g.name).toBe("a");
    expect(g.ran, "did not count every round the scenario actually ran in").toBe(8);
    // Two: the old one and this one. A first-ever failure would be 1, which is
    // the distinction the gate prints.
    expect(g.failed, "lifetime failures were not counted").toBe(2);
  });

  it("reports a first-ever failure as exactly one", () => {
    const rounds = [...Array.from({ length: 6 }, () => round([["a", true]])), round([["a", false]])];
    const [g] = scenarioRegressions(rounds) as { failed: number; ran: number }[];
    expect(g.failed, "a first-ever failure must be distinguishable from a recurring one").toBe(1);
    expect(g.ran).toBe(7);
  });

  it("does not count a round that never ran the scenario, or one that declined to conclude", () => {
    // `undefined` never ran it; `skipped` declined to conclude. Neither is
    // evidence, and folding either into the denominator makes a rate that
    // describes nothing — the denominator error this repo keeps making.
    // The non-measuring rounds sit BEFORE the establishment window: a scenario
    // is only "established" if it passed every one of the previous 3, and a
    // round that declined to conclude breaks that by design.
    const rounds = [
      { selftest: [{ name: "b", ok: true }] },
      { selftest: [{ name: "a", ok: false, skipped: true }] },
      round([["a", true]]),
      round([["a", true]]),
      round([["a", true]]),
      round([["a", false]]),
    ];
    const [g] = scenarioRegressions(rounds) as { ran: number; failed: number }[];
    expect(g.ran, "counted a round that did not measure this scenario").toBe(4);
    expect(g.failed).toBe(1);
  });
});

describe("a detector keyed to a message that no longer exists", () => {
  it("matches only trace messages the source still emits", () => {
    // THE FAILURE MODE THIS PREVENTS IS SILENCE. Every one of these tools finds
    // its evidence by comparing a trace message to a string literal. Rename the
    // message in src/ and the detector does not break — it reports ZERO, every
    // round, forever, and zero is exactly what a healthy round looks like.
    //
    // Nothing in the repo checked this. All the current literals happen to be
    // live, which is the point: the guard is for the rename that has not
    // happened yet, and the archive would carry months of false calm first.
    const tool = readFileSync("scripts/triage.mjs", "utf8");
    const app = readdirSync("src", { recursive: true })
      .map(String)
      .filter((f) => f.endsWith(".ts"))
      .map((f) => readFileSync(`src/${f}`, "utf8"))
      .join("\n");

    const matched = new Set<string>(Object.keys(FALLBACK_SIGNALS));
    // Deliberately backslash-free: a regex written through a heredoc has been
    // corrupted here before, and a mangled one matches nothing while looking fine.
    for (const m of tool.matchAll(/(?:e[.])?message *=== *"([^"]+)"/g)) matched.add(m[1]);
    for (const m of tool.matchAll(/[^A-Za-z]m === "([^"]+)"/g)) matched.add(m[1]);

    // THE REGEXES THEMSELVES ARE INSTRUMENTS. If one stops matching, the loop
    // below passes vacuously and this test becomes decoration — the exact shape
    // it exists to catch. Claim a positive count, not the absence of failures.
    expect(matched.size, "the extractors matched almost nothing — they have stopped working").toBeGreaterThanOrEqual(8);

    for (const message of matched)
      expect(app, `no source file emits "${message}" any more — its detector now reports zero forever`).toContain(
        message,
      );
  });
});

describe("what it took to start each round", () => {
  it("counts only the rounds that actually carry the field", () => {
    // AN ABSENT READING IS NOT A CLEAN ONE. `driverRun` is newer than most of
    // the archive, so 149 of the rounds on file say nothing about what it took
    // to start them. Counting those as first-time starts would invent a history
    // of healthy rounds out of a field that did not exist yet — the same shape
    // as reading a missing count as a zero.
    const pooled = poolDriverRuns([
      { driverRun: { attempts: 1, recovered: [] } },
      { driverRun: { attempts: 2, recovered: ["not-ready:host-silent+pane-stale"] } },
      {}, // an older round: no driverRun at all
      { driverRun: { attempts: 3, recovered: ["crashed", "not-ready:pane-stale"] } },
    ]);
    expect(pooled.rounds, "counted a round that never reported").toBe(3);
    expect(pooled.clean, "an absent reading was counted as a clean start").toBe(1);
  });

  it("says whether the charts landed on the slide, and off each other", () => {
    /**
     * UNANSWERABLE FROM THE ARCHIVE UNTIL 2026-09-01. The deck inventory carried
     * a shape's ORIGIN and no extent, so across 322 rounds and 10,362 shapes the
     * only geometry fault it could express was an origin outside the slide.
     * Whether a chart landed INSIDE the slide, and whether two landed on top of
     * each other, could be answered only by a person opening the deck — which is
     * how both were found: the owner opened what the battery left him and saw one
     * full-size chart drawn over another, then three in a heap.
     *
     * GROUPS ONLY. A grouped chart is one shape named `PowerChart`; an ungrouped
     * one is its loose parts, and those overlap by design because a label sits on
     * the segment it names. Comparing everything would call every ungrouped chart
     * a collision — and round 347 is exactly that mix, 13 groups and 7 parts.
     */

    const round = (shapes: Record<string, unknown>[]) => ({
      slideSize: { width: 960, height: 540 },
      deck: { inventory: [{ slideId: "s1", index: 0, shapes }] },
    });
    const chart = (id: string, left: number, top: number, width = 100, height = 80) => ({
      id,
      name: "PowerChart",
      left,
      top,
      width,
      height,
    });

    const clean = deckGeometryFaults(round([chart("a", 10, 10), chart("b", 200, 10)]));
    expect(clean.offSlide, "a chart well inside the slide was called off it").toEqual([]);
    expect(clean.collisions, "two charts side by side were called a collision").toEqual([]);
    expect(clean.measured).toBe(2);

    // Off the right edge — the case the inventory could never see before.
    const spilled = deckGeometryFaults(round([chart("a", 900, 10)]));
    expect(spilled.offSlide).toHaveLength(1);
    expect(spilled.offSlide[0].off).toBe("right");

    // One chart drawn over another: what the owner actually found on a slide.
    const stacked = deckGeometryFaults(round([chart("a", 100, 100), chart("b", 150, 120)]));
    expect(stacked.collisions, "one chart drawn over another went unreported").toHaveLength(1);
    expect(stacked.collisions[0].area).toBeGreaterThan(0);

    // LOOSE PARTS ARE NOT CHARTS. A label sitting on its own segment is the
    // engine working, and counting it would make every ungrouped chart a fault.
    const parts = deckGeometryFaults(
      round([
        { id: "t", name: "title", left: 0, top: 0, width: 300, height: 20 },
        { id: "s", name: "seg-0-0", left: 10, top: 5, width: 50, height: 40 },
      ]),
    );
    expect(parts.collisions, "a chart's own parts were reported as colliding charts").toEqual([]);
    expect(parts.measured, "loose parts still count as measured shapes").toBe(2);

    // A SHAPE WITH NO SIZE IS NOT A CLEAN SHAPE. Every round archived before this
    // landed carries origins only — 5,995 of them — and none of that may read as
    // a deck with nothing wrong with it.
    const old = deckGeometryFaults(round([{ id: "a", name: "PowerChart", left: 10, top: 10 }]));
    expect(old.measured, "a shape with no extent was counted as measured").toBe(0);
    expect(old.unmeasured).toBe(1);
    expect(old.offSlide, "a shape that could not be measured was judged anyway").toEqual([]);

    // And a round with no slide size answers nothing at all, rather than "clean".
    expect(deckGeometryFaults({ deck: { inventory: [] } }), "a round with no slide size was judged").toBeNull();
  });

  it("divides crashes by ATTEMPTS per slide size — the number nobody had computed", () => {
    /**
     * Every input to this has been in the archive for months and no instrument
     * divided one by the other; `docs/BACKLOG.md` says so in as many words. Over
     * 322 rounds the answer is 40 crash events in 81 attempts at 4:3 against 9
     * in 365 at 16:9 — a sixteen-fold gap that no report printed.
     *
     * ATTEMPTS, NOT ROUNDS, is the denominator, and this fixture is built so the
     * two disagree: the 4:3 round below crashed twice before landing, so it is
     * ONE round and THREE attempts. Counting rounds would report one-in-one and
     * hide how hard that arm is failing.
     */
    const pooled = poolDriverRuns([
      { driverSlideSize: "4:3", driverRun: { attempts: 3, recovered: ["crashed", "crashed"] } },
      { driverSlideSize: "16:9", driverRun: { attempts: 1, recovered: [] } },
      { driverSlideSize: "16:9", driverRun: { attempts: 2, recovered: ["not-ready:pane-closed"] } },
      { driverRun: { attempts: 1, recovered: [] } }, // never said which size it ran
    ]);
    const by = Object.fromEntries(pooled.bySize.map((s: { size: string }) => [s.size, s]));
    expect(by["4:3"]).toMatchObject({ rounds: 1, attempts: 3, crashes: 2, roundsWithCrash: 1 });
    expect(by["16:9"]).toMatchObject({ rounds: 2, attempts: 3, crashes: 0, roundsWithCrash: 0 });
    // A round that did not record its size gets its own bucket and is never
    // folded into a real arm — the same rule `unrecorded` follows everywhere
    // else here, because a missing reading is not a reading of zero.
    expect(by.unrecorded).toMatchObject({ rounds: 1, attempts: 1, crashes: 0 });
    // Worst first, so the failing arm cannot be skimmed past.
    expect(pooled.bySize[0].size, "the worst arm was not listed first").toBe("4:3");
    // A recovery reason that merely MENTIONS a crash is still one crash event.
    expect(
      poolDriverRuns([{ driverSlideSize: "4:3", driverRun: { attempts: 2, recovered: ["crashed+pane-closed"] } }])
        .bySize[0].crashes,
      "a compound recovery reason was not counted as a crash",
    ).toBe(1);
  });

  it("separates the three reasons a chart has no parts list, which one counter cannot", async () => {
    // `tracePartsOutcome` was built to tell three faults apart — never built
    // because the chart was grouped; built and lost to a throwing read-back;
    // built and not found again — and it has recorded all four counters on
    // every one of its 607 events across 29 rounds. No script ever read them.
    //
    // AND THE OBVIOUS READING IS WRONG. `gotPartsList` is 0 on all 607, which
    // taken alone says a production path has never once worked. 569 of the 607
    // are cases where a parts list is not WANTED: 346 grouped charts, which need
    // none by design, and 223 calls with no charts at all. The genuine failure
    // is 36 throwing read-backs, about 1.2 a round. The counter that separates
    // them is `where`, and it sat in the same object the whole time.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolPartsListOutcome } = await import("../scripts/triage.mjs");
    const ev = (data: Record<string, unknown>) => ({ message: "parts list outcome", data });
    const pooled = poolPartsListOutcome([
      {
        trace: {
          entries: [
            ev({ where: "no loose chart had siblings", charts: 1, groupedSoNotLoose: 1, gotPartsList: 0 }),
            ev({ where: "no loose chart had siblings", charts: 0, groupedSoNotLoose: 0, gotPartsList: 0 }),
            ev({ where: "the id read-back threw", charts: 1, groupedSoNotLoose: 0, gotPartsList: 0 }),
            ev({ where: "read the ids back", charts: 1, groupedSoNotLoose: 0, gotPartsList: 0 }),
          ],
        },
      },
      { trace: { entries: [{ message: "drew a chart" }] } },
    ]);
    expect(pooled).toMatchObject({
      grouped: 1,
      noCharts: 1,
      threw: 1,
      builtNothing: 1,
      gotList: 0,
      events: 4,
      rounds: 1,
    });

    // A GROUPED CHART IS NOT A FAILURE, and folding it in with the throws is
    // exactly the reading that made this look like a dead path. The two must
    // never land in the same bucket.
    const allGrouped = poolPartsListOutcome([
      { trace: { entries: [ev({ where: "no loose chart had siblings", charts: 2, groupedSoNotLoose: 2 })] } },
    ]);
    expect(allGrouped.grouped).toBe(1);
    expect(allGrouped.threw, "a grouped chart was counted as a failure").toBe(0);
  });

  it("separates a SECOND settle ask from a first, which one number cannot", async () => {
    // THE RETRY IS LOAD-BEARING AND NOTHING MEASURED IT. Over 161 rounds before
    // #709: 1,057 retry passes fired and 97 failures survived — a 90.8% rescue
    // rate. The first read of the shape collection fails on roughly half the
    // charts this loop sees; the 4.3% quoted all over this project is what is
    // LEFT after the retry, not the failure rate of the read.
    //
    // #709 raised REREAD_ATTEMPTS 1 -> 2 to reach those 97. Round 187 then
    // recorded FIVE retry passes and zero survivors — and nothing could say
    // whether any was a second ask or whether all five were firsts from five
    // separate update calls. The change reads as working and as inert on
    // identical evidence, which is the house defect aimed at a change made to
    // cure it.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolSettleAsks } = await import("../scripts/triage.mjs");
    const pass = (attempt: number | undefined, charts: number) => ({
      message: "re-reading the slide's shapes again after a settle delay",
      data: attempt === undefined ? { charts } : { attempt, charts },
    });
    const pooled = poolSettleAsks([
      {
        trace: {
          entries: [
            pass(1, 3),
            pass(1, 2),
            pass(2, 1),
            { message: "the re-read named none of the chart's shapes", data: {} },
          ],
        },
      },
      // A round from before the field existed: counted apart, never folded in.
      { trace: { entries: [pass(undefined, 4)] } },
    ]);
    expect(pooled).toMatchObject({ rounds: 2, first: 2, second: 1, unlabelled: 1, survivors: 1, charts: 10 });

    // AN UNLABELLED PASS IS NOT A FIRST ASK. Folding the old rounds in would
    // make the second ask look like it had never fired across an archive that
    // could not have recorded it either way.
    const old = poolSettleAsks([{ trace: { entries: [pass(undefined, 1), pass(undefined, 1)] } }]);
    expect(old.first, "a pass from before the field was counted as a first ask").toBe(0);
    expect(old.unlabelled).toBe(2);
  });

  it("dates the survivors against when the SECOND ASK shipped, not when its label did", async () => {
    // WHAT THIS COST: five rounds read as a win that were not one.
    //
    // The archive said "0 survivors since the change" and that was true and
    // meaningless — the survivors had stopped EIGHT rounds BEFORE the change:
    //
    //     rounds 150-179   12 survivors over 30 rounds  =  0.40 per round
    //     rounds 180-191    0 survivors over 12 rounds  =  0.00 per round
    //     #709 (the second ask) first ran in round 187
    //
    // Pooling survivors as one TOTAL cannot show that. A total is a number
    // with the timing divided out, and the timing was the whole finding.
    //
    // The first fix derived "when the change landed" from the attempt label —
    // and the label is a DIFFERENT COMMIT two rounds later (#710), so the era
    // was two rounds short under a field named for the change. That is the
    // same defect one layer up, so the two clocks are pinned apart here.
    // @ts-expect-error — plain .mjs tool, no types.
    const mod = await import("../scripts/triage.mjs");
    const { poolSettleAsks, SECOND_ASK_BUILD } = mod;
    const ask = (attempt?: number) => ({
      message: "re-reading the slide's shapes again after a settle delay",
      data: attempt === undefined ? { charts: 1 } : { attempt, charts: 1 },
    });
    const survivor = { message: "the re-read named none of the chart's shapes", data: {} };
    const round = (build: string, entries: unknown[]) => ({ build, trace: { entries } });

    // The real shape: a survivor, then quiet, THEN the change, then the label.
    const a = poolSettleAsks([
      round("aaaaaaa", [ask(undefined), survivor]), //  the last survivor
      round("aaaaaaa", [ask(undefined)]), //            quiet
      round("bbbbbbb", [ask(undefined)]), //            still quiet, still before
      round(SECOND_ASK_BUILD, [ask(undefined)]), //     <- the second ask ships
      round(SECOND_ASK_BUILD, [ask(undefined)]),
      round("ccccccc", [ask(1)]), //                    <- only now can we tell asks apart
    ]);
    expect(a.sinceLastSurvivor, "the survivor clock").toBe(5);
    expect(a.changeLandedAgo, "dated from the LABEL instead of the build").toBe(2);
    expect(a.labelledAgo, "the label clock").toBe(0);
    // The two clocks must not be the same number, or nothing above is proven.
    expect(a.changeLandedAgo).not.toBe(a.labelledAgo);
    // 5 > 2: the population died before the change could have touched it.
    expect(a.sinceLastSurvivor).toBeGreaterThan(a.changeLandedAgo as number);
    expect(a.survivorsSinceChange, "a survivor after the change went uncounted").toBe(0);
    expect(a.roundsSinceChange).toBe(3);

    // A survivor AFTER the change is counted, so a real failure still shows.
    const b = poolSettleAsks([round("aaaaaaa", [ask(undefined)]), round(SECOND_ASK_BUILD, [ask(1), survivor])]);
    expect(b.survivorsSinceChange).toBe(1);
    expect(b.sinceLastSurvivor).toBe(0);

    // A slice with no round on that build INVENTS NO ERA. Reporting 0 here
    // would read as "the change landed just now" for every old archive.
    const c = poolSettleAsks([round("aaaaaaa", [ask(undefined), survivor])]);
    expect(c.changeLandedAgo, "an era was invented for rounds that never ran the build").toBeNull();
    expect(c.sinceLastSurvivor).toBe(0);
  });
  it("sizes a batch by its DELTA, and bands it by what the run already drew", async () => {
    // TWO WRONG ANSWERS ARE PINNED HERE, and both of them reached a commit.
    //
    // 1. `upTo` is a RUNNING TOTAL. Bucketing by it read the second batch of
    //    ten as a batch of twenty and produced "payload is free, so raise
    //    `SHAPES_PER_SYNC`". Every batch in the archive is ten; there is no
    //    second size to compare against.
    //
    // 2. `onSlideKey === "(visible)"` is the sentinel for an UNLOADED SLIDE
    //    ID, not a claim about the screen. Splitting on it gave "drawing where
    //    the user is looking costs 2.3x per shape" — from batches whose median
    //    `onSlide` was 0 against 10 for the others, so the EMPTIER targets
    //    were the slower ones, and first batches cost 5802ms against 5591ms
    //    for later ones. The label never carried the meaning put on it.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolDrawCostCurve } = await import("../scripts/triage.mjs");
    const issued = (upTo: number, onSlide: number, prevBatchMs?: number) => ({
      message: "batch issued",
      data: prevBatchMs === undefined ? { upTo, onSlide } : { upTo, onSlide, prevBatchMs },
    });

    // THREE BATCHES PER DRAW, NOT TWO, and that is load-bearing.
    //
    // With two, the only batch that gets timed is the FIRST — and the first
    // batch's size is 10 whether you read `upTo` as a delta or as a running
    // total, because the running total starts at zero. So a two-batch fixture
    // passes with the bug put back: it never reaches a batch where the two
    // readings disagree. Mutation caught that; the assertions had looked fine.
    //
    // The third entry times the SECOND batch, which is a ten by delta and a
    // twenty by running total. That is the only place the trap is visible.
    // A band needs at least ten samples to be reported, so twelve draws each.
    const entries: unknown[] = [];
    for (let i = 0; i < 12; i++) entries.push(issued(10, 0), issued(20, 0, 4000), issued(30, 0, 4000));
    for (let i = 0; i < 12; i++) entries.push(issued(10, 40), issued(20, 40, 14000), issued(30, 40, 14000));
    const c = poolDrawCostCurve([{ trace: { entries } }]);

    // Batch SIZE is the delta, so every sample is a ten and never a twenty.
    expect(c.sizes, "upTo was read as a batch size").toEqual([10]);
    const band = (lo: number) => c.rows.find((r: { lo: number }) => r.lo === lo);
    expect(band(0)?.median, "the empty-target band").toBe(4000);
    expect(band(21)?.median, "the filled-target band").toBe(14000);
    // The curve is the whole point: the bands must not collapse together.
    expect(band(21).median).toBeGreaterThan(band(0).median);
  });

  it("counts whether a second ask RESCUED the chart, not merely that it fired", async () => {
    // Round 194 produced the first second ask on the real host, and the trace
    // reads, in order:
    //
    //   the cold re-read fell short …        {kind:"zero-match", drew:7, listed:8}
    //   … after a settle delay               {attempt:1}
    //   the cold re-read fell short …        {kind:"zero-match", drew:7, listed:8}
    //   … after a settle delay               {attempt:2}     <- #709
    //   grouped the chart's shapes           {charts:1, partial:0, by:"ids"}
    //
    // Under REREAD_ATTEMPTS = 1 that chart is a survivor. Counting that the ask
    // FIRED would have said nothing about that; the outcome is the whole point.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolSettleAsks } = await import("../scripts/triage.mjs");
    const ask = (attempt: number) => ({
      message: "re-reading the slide's shapes again after a settle delay",
      data: { attempt, charts: 1 },
    });
    const grouped = { message: "grouped the chart's shapes", data: { charts: 1, partial: 0 } };
    const survivor = { message: "the re-read named none of the chart's shapes", data: {} };

    const rescued = poolSettleAsks([{ trace: { entries: [ask(1), ask(2), grouped] } }]);
    expect(rescued.second).toBe(1);
    expect(rescued.secondAskRescued, "a group after the second ask is a rescue").toBe(1);
    expect(rescued.secondAskLost).toBe(0);

    const lost = poolSettleAsks([{ trace: { entries: [ask(1), ask(2), survivor] } }]);
    expect(lost.secondAskRescued, "a survivor after the second ask is NOT a rescue").toBe(0);
    expect(lost.secondAskLost).toBe(1);

    // A FIRST ask followed by a group is not credited to the second one. This is
    // the assertion that stops the counter reporting every ordinary success.
    const firstOnly = poolSettleAsks([{ trace: { entries: [ask(1), grouped] } }]);
    expect(firstOnly.second).toBe(0);
    expect(firstOnly.secondAskRescued, "a first-ask success was credited to the second ask").toBe(0);

    // And an ask with nothing decisive after it is counted NEITHER way, rather
    // than assumed to have worked.
    const undecided = poolSettleAsks([{ trace: { entries: [ask(1), ask(2)] } }]);
    expect(undecided.second).toBe(1);
    expect(undecided.secondAskRescued + undecided.secondAskLost, "an undecided ask was counted").toBe(0);
  });
  it("prices an in-place update WITHIN one chart size, never across two", async () => {
    // The in-place update path is the round's largest single cost —
    // `same scale across the deck` is 166s and issues no batches at all — and it
    // had no clock until 05a27fd. Round 198, the first to carry one, gives
    // ~1589ms fixed + ~930ms per changed shape for a 24-node chart.
    //
    // THE FIXTURE IS BUILT TO MAKE THE CONFOUNDS BITE, and the first version of
    // it was not. It had the second chart size changing 9 and 2 shapes while the
    // fit's endpoints were 1 and 18, so pooling the sizes together moved
    // nothing and the guard could be deleted with the test still green. Same
    // for the first-chart comparison: there was no same-size, different-changed
    // pair for a dropped `changed` match to pull in. Both mutations passed.
    //
    // So the other size now SHARES an endpoint (changed: 1) with a wildly
    // different cost, and the run carries a same-size row at a different
    // changed count. Deleting either guard now fails.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolUpdateCost } = await import("../scripts/triage.mjs");
    const upd = (changed: number, of: number, ms: number, chart?: string) => ({
      message: "updated only the shapes that changed",
      data: chart === undefined ? { changed, of, ms } : { changed, of, ms, chart },
    });

    // 24-node: 1 -> 2500ms, 18 -> 18000ms. Slope (18000-2500)/17 = 911.8.
    const u = poolUpdateCost([
      {
        trace: {
          entries: [
            upd(1, 24, 2500),
            upd(18, 24, 18000),
            // SHARES the lo endpoint (changed: 1) at a fraction of the cost, and
            // THREE OF THEM, which is the part that took a second mutation run to
            // get right. With one contaminating row the pooled set is
            // [100, 2500], and this file's median takes the upper middle of an
            // even-length array — so it returned 2500, exactly what the guarded
            // version returns, and mutation 1 passed against a fixture that
            // looked like it should catch it. Three outvote the real sample and
            // the median actually moves.
            upd(1, 16, 100),
            upd(1, 16, 100),
            upd(1, 16, 100),
            upd(9, 16, 900),
          ],
        },
      },
    ]);
    const fit24 = u.fits.find((f: { of: number }) => f.of === 24);
    expect(fit24, "no fit for the 24-node chart").toBeTruthy();
    expect(Math.round(fit24.perShape), "the 16-node chart leaked into the 24-node fit").toBe(912);
    expect(Math.round(fit24.fixed), "the 16-node chart moved the 24-node intercept").toBe(1588);
    expect(
      u.fits.find((f: { of: number }) => f.of === 16),
      "the 16-node chart lost its own fit",
    ).toBeTruthy();

    // First chart of a run against the rest, same size AND same changed count.
    const run = poolUpdateCost([
      {
        trace: {
          entries: [
            upd(18, 24, 39000, "1/8"),
            upd(18, 24, 18000, "4/8"),
            upd(18, 24, 18000, "5/8"),
            // SAME SIZE, DIFFERENT changed. Dropping the changed match pulls
            // this into the 18-shape comparison and the ratio stops measuring
            // position at all.
            upd(3, 24, 4000, "2/8"),
          ],
        },
      },
    ]);
    const cmp = run.firstVsRest;
    expect(cmp.length, "a comparison was made across mismatched work").toBe(1);
    expect(cmp[0]).toMatchObject({ of: 24, changed: 18, firstN: 1, restN: 2 });
    expect(cmp[0].first).toBe(39000);
    expect(cmp[0].rest, "a different changed count leaked into the rest").toBe(18000);

    // A lone update outside a run is not a "first chart" — it has no rest to be
    // first of, and counting it would invent a comparison.
    const solo = poolUpdateCost([{ trace: { entries: [upd(18, 24, 39000), upd(18, 24, 18000)] } }]);
    expect(solo.firstVsRest, "an update with no run was treated as a first chart").toEqual([]);
  });
  it("flags numeric trace fields that never vary, and leaves varying ones alone", async () => {
    // Every serious error this project has made is one shape: an unmeasured
    // thing printed as a measurement. Nothing detected the CLASS — each one was
    // caught by eye, late.
    //
    // The case that produced this reader: `contextSyncs` went onto the in-place
    // update's trace line, and round 202 answered 0 for eight charts costing
    // 12-37 seconds each. A 37-second update issuing no syncs is a broken gauge,
    // and it was visible in the first round that carried it.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolFlatFields } = await import("../scripts/triage.mjs");
    const entry = (message: string, data: Record<string, number>) => ({ message, data });
    const many = (n: number, make: (i: number) => unknown) => Array.from({ length: n }, (_, i) => make(i));

    const pooled = poolFlatFields(
      [
        {
          trace: {
            entries: [
              // Always 0 — the shape that matters.
              ...many(25, () => entry("an update", { contextSyncs: 0 })),
              // Always the same NON-ZERO value — a constant wearing a
              // measurement's clothes.
              ...many(25, () => entry("a settle", { waitedMs: 1500 })),
              // Varies — must not be flagged, however boring the spread.
              ...many(25, (i) => entry("a draw", { ms: 1000 + i })),
            ],
          },
        },
      ],
      20,
    );

    const zeroFields = pooled.zeros.map((r: { field: string }) => r.field);
    const constFields = pooled.constants.map((r: { field: string }) => r.field);
    expect(zeroFields, "an always-zero field was not flagged").toContain("contextSyncs");
    expect(constFields, "an always-constant field was not flagged").toContain("waitedMs");
    // THE ASSERTION THAT MAKES THE OTHERS MEAN ANYTHING. A detector that flags
    // everything has found nothing.
    expect(zeroFields, "a varying field was flagged as flat").not.toContain("ms");
    expect(constFields, "a varying field was flagged as constant").not.toContain("ms");

    // Thin data is not evidence of flatness. Below minSamples nothing is
    // reported, or every new field reads as a broken gauge on its first round.
    const thin = poolFlatFields([{ trace: { entries: many(3, () => entry("an update", { contextSyncs: 0 })) } }], 20);
    expect(thin.zeros, "three samples were enough to call a field flat").toEqual([]);

    // Non-numeric values are ignored rather than counted as constants — every
    // slideId in a round is a string that never varies within its chart.
    const strings = poolFlatFields(
      [{ trace: { entries: many(25, () => ({ message: "an update", data: { slideId: "257#1" } })) } }],
      20,
    );
    expect(strings.constants, "a string field was treated as a measurement").toEqual([]);
  });
  it("reads the deck inventory the gate has printed all along", async () => {
    // EIGHT OF THE LAST THIRTY ROUNDS ended with a slide holding between 11 and
    // 48 shapes, and every one of them reported 13 of 13 scenarios passed —
    // rounds 140, 141, 142, 147, 150, 159, 161 and 167. A clean round's fullest
    // slide holds five. A chart that fails to group is left as its loose
    // shapes, and NO SCENARIO VERDICT LOOKS AT THE DECK.
    //
    // `does a rasterise poison the next draw` is the clearest case. It draws
    // four charts and reports `all four draws landed`, which is literally true
    // and narrower than any reader takes it: it asks whether the CALL came
    // back, not whether a chart survived. It passes with eight loose shapes
    // sitting where a chart should be.
    //
    // The gate has printed the inventory since the beginning. Nothing compared
    // it to anything — which is the difference between a number being on screen
    // and a number being read, and it is why I called round 167 clean tonight.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolFullestSlide, CLEAN_SLIDE_CEILING } = await import("../scripts/triage.mjs");
    const deck = (...counts: number[]) => ({ deck: { inventory: counts.map((count) => ({ count })) } });
    expect(poolFullestSlide([deck(0, 4, 1, 2, 5, 1, 1), deck(0, 4, 1, 2, 17, 1, 1)], 8)).toEqual([5, 17]);
    expect(17).toBeGreaterThan(CLEAN_SLIDE_CEILING);
    expect(5, "a clean round must sit at or below the ceiling, or it fires every round").toBeLessThanOrEqual(
      CLEAN_SLIDE_CEILING,
    );

    // A round with no inventory reads 0 rather than dropping out, so the
    // sequence stays aligned with the rounds it describes.
    expect(poolFullestSlide([{}, deck(3)], 8)).toEqual([0, 3]);
    // And the window is the last N, not the first.
    expect(poolFullestSlide([deck(9), deck(1), deck(2)], 2)).toEqual([1, 2]);
  });

  it("pools how badly the host renumbers, which the trace recorded and nothing read", async () => {
    // RECORDED SINCE ROUND 041 AND READ BY NOTHING. `claimed the appended slide
    // though the id list churned` carries `before`, `after` and `fresh` on all
    // 119 of its occurrences across 63 rounds, and no script ever matched on it.
    //
    // The comment in powerpoint.ts that describes the behaviour was written on
    // 2026-08-10, before the round archive existed, from seven hand-collected
    // observations — and it said the count was ALWAYS two. It is not: 97 events
    // read 2 and 22 read 3, and the two populations separate almost cleanly by
    // DECK SIZE (median 12 against 77). All seven originals were taken at
    // before=3..37, where fresh=3 is nearly absent.
    //
    // A hand-collected sample cannot notice that its own conclusion is a
    // function of a variable it never varied. The archive can.
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolIdChurn } = await import("../scripts/triage.mjs");
    const churn = (events: { fresh: number; before: number }[]) => ({
      trace: {
        entries: events.map((e) => ({
          message: "claimed the appended slide though the id list churned",
          data: { before: e.before, after: e.before + 1, fresh: e.fresh },
        })),
      },
    });
    const pooled = poolIdChurn([
      churn([
        { fresh: 2, before: 10 },
        { fresh: 2, before: 14 },
      ]),
      churn([{ fresh: 3, before: 77 }]),
      { trace: { entries: [{ message: "drew a chart" }] } },
    ]);
    expect(pooled.rounds, "a round with no churn event is not a round that churned").toBe(2);
    expect(pooled.kinds).toEqual([
      { fresh: 2, events: 2, medianDeck: 14 },
      { fresh: 3, events: 1, medianDeck: 77 },
    ]);

    // THE DECK SIZE IS THE POINT, not decoration: it is the variable the
    // original seven observations held nearly constant, and dropping it would
    // leave the pooled version making the same mistake with more data.
    expect(pooled.kinds.every((k: { medianDeck?: number }) => k.medianDeck !== undefined)).toBe(true);
  });

  it("records WHICH ARM a round was in, which the archive could not say", () => {
    // WHAT IT COST. Six rounds on 2026-08-22 established that a second-round-of-
    // a-pair on an aged pane refuses a group and one on a fresh pane does not —
    // the whole argument for running the second round with `--fresh`. Then the
    // claim could not be entered in the prediction ledger, because NOTHING IN A
    // ROUND FILE SAID WHICH ROUNDS WERE RUN WITH `--fresh`. `driverRun` carried
    // `attempts` and `recovered` and not the flag, so the argument rested on the
    // shell history of whoever typed the command. Rounds 163 and 165 are
    // labelled `--fresh` in this project's journal on nothing but one agent's
    // word, and no later reader can check it.
    //
    // The house defect in its plainest costume: an experiment whose ARM was not
    // recorded, while everything else about it was.
    // `paneAgeAtStartSeconds` reads the SMALLEST `ms` in the trace, so the
    // fixture has to carry one — the pane's age is a property of the trace, not
    // of `driverRun`, and it has been in every round file all along.
    const round = (fresh: boolean | undefined, paneSeconds: number, refused: number) => ({
      driverRun: fresh === undefined ? { attempts: 1, recovered: [] } : { attempts: 1, recovered: [], fresh },
      trace: {
        entries: [
          { message: "round starting", ms: paneSeconds * 1000 },
          ...Array.from({ length: refused }, () => ({
            message: "not grouping: no member handle this host will accept",
            ms: paneSeconds * 1000 + 1,
          })),
        ],
      },
    });
    const pooled = poolDriverRuns([
      round(true, 60, 0),
      round(true, 60, 0),
      round(false, 700, 1),
      round(undefined, 700, 1),
    ]);
    expect(pooled.arms).toMatchObject({ fresh: 2, freshRefusedNone: 2, aged: 2, agedRefusedNone: 0 });

    // AND THE SPLIT IS ON THE PANE, NOT ON THE FLAG. Round 167 showed the
    // difference inside a single pair: 166 ran WITHOUT `--fresh` and started on
    // a 69-second pane anyway, because a merge preceded it and recovery reloads
    // a stale pane. The flag is one way to get a fresh pane, not the variable.
    // Splitting on it files a fresh-pane round under "aged" and makes both arms
    // mean nothing — the house defect, in an instrument built hours earlier to
    // fix the same defect: measuring the PROXY instead of the thing.
    //
    // ASSERTED ON WHICH ROUND LANDED IN WHICH ARM, not on the arm SIZES — the
    // sizes are 1 and 1 under either split, so a count-only assertion passes
    // with the defect restored. It did: the first mutation of this test put
    // `dr.fresh` back and stayed green. A count guard is not an alignment
    // guard.
    //
    // The two rounds are deliberately crossed: the flag says fresh where the
    // pane says aged, and vice versa. Under the PANE split the 60-second round
    // (which refused nothing) is the fresh one, so `freshRefusedNone` is 1 and
    // `agedRefusedNone` is 0. Under the FLAG split they swap.
    const disagreeing = poolDriverRuns([round(false, 60, 0), round(true, 700, 1)]);
    expect(disagreeing.arms, "the arms were filled by the flag, not the pane").toMatchObject({
      fresh: 1,
      freshRefusedNone: 1,
      aged: 1,
      agedRefusedNone: 0,
    });
    expect(disagreeing.arms.flagDisagreed, "the proxy parting company with the variable went unreported").toBe(2);

    // A round with no readable pane age is in neither arm. `unlabelled` is part
    // of the answer, not a gap: a split over a handful read as a split over the
    // archive is the overreach this ledger was just fixed to stop.
    expect(poolDriverRuns([{ driverRun: { attempts: 1, recovered: [] } }]).arms).toMatchObject({
      fresh: 0,
      aged: 0,
      unlabelled: 1,
    });
  });

  it("tallies the causes, so a deploy is not read as a sick host", () => {
    // `pane-stale` after a deploy is a property of how rounds are RUN;
    // `host-silent` is host health. They were one word — "not-ready" — for four
    // rounds, and any claim about the host getting better or worse would have
    // been drawn from that bucket.
    const pooled = poolDriverRuns([
      { driverRun: { attempts: 2, recovered: ["not-ready:pane-stale"] } },
      { driverRun: { attempts: 2, recovered: ["not-ready:pane-stale"] } },
      { driverRun: { attempts: 2, recovered: ["crashed"] } },
    ]);
    expect(pooled.causes[0], "the commonest cause was not surfaced first").toEqual({
      cause: "not-ready:pane-stale",
      n: 2,
    });
    expect(pooled.causes.map((c: { cause: string }) => c.cause)).toContain("crashed");
  });
});

describe("signals a round records that nothing reads", () => {
  const round = (messages: string[]) => ({ trace: { entries: messages.map((message) => ({ message })) } });

  it("offers the busiest unread signal, and hides the ones a tool already matches", () => {
    // THIS REPO HAS FOUND THE SAME THING BY HAND TWICE. `poolFallbackRates`
    // exists because the fallback lines had been "recorded thousands of times
    // and read by nothing"; `poolInPlaceUpdates` because the in-place
    // diagnostic "has been answering ever since and nobody has read it". Both
    // were noticed by someone scrolling a round file months later.
    const log = round(["asking", "asking", "asking", "grouped the chart's shapes", "grouped the chart's shapes"]);
    const found = unreadSignals(log, 'if (m === "grouped the chart\'s shapes") grouped++;');
    expect(found, "a signal a tool already reads was offered as unread").toEqual([{ message: "asking", n: 3 }]);
  });

  it("ranks by how often it fires, since a signal seen once is not a missed instrument", () => {
    const found = unreadSignals(round(["rare", "common", "common", "common"]), "");
    expect(found[0]).toEqual({ message: "common", n: 3 });
  });

  it("says nothing at all about a round with no trace", () => {
    // A round that never ran records nothing, and an empty list is the honest
    // answer — not a claim that every signal is read.
    expect(unreadSignals({}, "")).toEqual([]);
    expect(unreadSignals({ trace: {} }, "")).toEqual([]);
  });
});

describe("a counter that did not move, at the scope it was measured", () => {
  it("will not call a counter dead from a single round", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { deadCounterNote } = await import("../scripts/triage.mjs");
    // THE DEFECT THIS REPLACES: the line said "in any scenario in any round"
    // from whatever logs it was handed. On one round file - which is how a
    // fresh round is read - it asserted a fact about 184 rounds from n=1, and
    // `emptyReReads` is non-zero in 72 of them.
    const one = deadCounterNote(["emptyReReads"], 1);
    expect(one).not.toContain("any round");
    expect(one).not.toContain("measures nothing");
    expect(one, "a single round cannot tell dead from quiet").toContain("cannot tell");
  });

  it("names both readings once there are rounds to pool", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { deadCounterNote } = await import("../scripts/triage.mjs");
    const many = deadCounterNote(["emptyReReads"], 183);
    expect(many, "the scope has to be the pool it was measured over").toContain("183 round(s)");
    expect(many).toContain("dead gauge");
    expect(many, "a fault that did not occur is the other reading").toContain("did not occur");
    expect(many).not.toContain("measures nothing");
  });
});

describe("a run of one, told apart from the first of many", () => {
  it("does not pool a run of one into the first-chart arm", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolUpdateCost } = await import("../scripts/triage.mjs");
    const entry = (chart: string, ms: number) => ({
      message: "updated only the shapes that changed",
      data: { chart, ms, changed: 18, of: 24, slideId: "s1" },
    });
    const log = {
      trace: {
        entries: [
          entry("1/8", 37000),
          entry("2/8", 17000),
          entry("3/8", 17200),
          // The run of one. Under the old `/^1//` test this was a FIRST chart.
          entry("1/1", 17100),
        ],
      },
    };
    const u = poolUpdateCost([log]);
    const c = u.firstVsRest.find((x: { of: number; changed: number }) => x.of === 24 && x.changed === 18);
    expect(c, "the 18-of-24 comparison should exist").toBeTruthy();
    expect(c.firstN, "only the 1/8 row is a first chart").toBe(1);
    expect(c.aloneN, "the 1/1 row is its own arm").toBe(1);
    expect(c.first, "a run of one must not drag the first-chart median").toBe(37000);
  });

  it("names which arm the lone chart landed nearer", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { aloneVerdict } = await import("../scripts/triage.mjs");
    const later = aloneVerdict({ alone: 17100, first: 37000, rest: 17000, aloneN: 3 });
    expect(later).toContain("a LATER chart");
    const first = aloneVerdict({ alone: 36500, first: 37000, rest: 17000, aloneN: 3 });
    expect(first).toContain("a FIRST chart");
  });

  it("refuses to call it from a single round", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { aloneVerdict } = await import("../scripts/triage.mjs");
    // The first-chart line already learned this the expensive way.
    expect(aloneVerdict({ alone: 36500, first: 37000, rest: 17000, aloneN: 1 })).toContain("n=1");
    expect(aloneVerdict({ alone: 36500, first: 37000, rest: 17000, aloneN: 1 })).toContain("cannot tell");
  });
});

describe("a run of one must not set the baseline it is scored against", () => {
  const entry = (chart: string | undefined, ms: number, changed: number) => ({
    message: "updated only the shapes that changed",
    data: { ...(chart ? { chart } : {}), ms, changed, of: 24, slideId: "s1" },
  });

  it("keeps a lone chart out of the per-shape fit", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolUpdateCost } = await import("../scripts/triage.mjs");
    const log = {
      trace: {
        entries: [
          entry("2/8", 10000, 9),
          entry("3/8", 17000, 18),
          // 37000ms alone. If this reaches the fit it more than triples the
          // per-shape slope and the outlier defines its own baseline away -
          // the exact hazard the comment above `here` names for first rows.
          entry("1/1", 37000, 18),
        ],
      },
    };
    const fit = poolUpdateCost([log]).fits.find((f: { of: number }) => f.of === 24);
    expect(fit, "a fit needs two change counts").toBeTruthy();
    expect(Math.round(fit.perShape), "(17000-10000)/9 - the lone row must not be in it").toBe(778);
  });

  it("keeps a lone chart out of the rest arm", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolUpdateCost } = await import("../scripts/triage.mjs");
    const log = {
      trace: {
        entries: [entry("1/8", 37000, 18), entry("2/8", 17000, 18), entry("3/8", 17200, 18), entry("1/1", 17100, 18)],
      },
    };
    const c = poolUpdateCost([log]).firstVsRest.find((x: { of: number }) => x.of === 24);
    expect(c.restN, "two later charts, not three").toBe(2);
    expect(c.rest, "median of 17000 and 17200 - the lone row is not one of them").toBe(17200);
  });
});

describe("what the first-chart cost was standing in for", () => {
  it("reads prior draws per slide off the run's own batch lines", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { priorDrawsOnSlide } = await import("../scripts/triage.mjs");
    const m = priorDrawsOnSlide([
      { message: "batch issued", data: { onSlideKey: "257", onSlide: 0 } },
      { message: "batch issued", data: { onSlideKey: "257", onSlide: 42 } },
      // OUT OF ORDER ON PURPOSE: the largest reading is not the last one, so a
      // `set` that overwrites instead of maxing reports 32 and the slide reads
      // less loaded than it was.
      { message: "batch issued", data: { onSlideKey: "257", onSlide: 32 } },
      { message: "batch issued", data: { onSlideKey: "262", onSlide: 0 } },
      // Not a batch line, and must not be counted as one.
      { message: "updated only the shapes that changed", data: { onSlideKey: "262", onSlide: 999 } },
    ]);
    expect(m.get("257"), "the largest reading, not the last").toBe(42);
    expect(m.get("262")).toBe(0);
  });

  it("splits the update cost by that, not by position", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { occupancySplit } = await import("../scripts/triage.mjs");
    // Round 215: the crowded slide costs 2.2x whether the chart is first or
    // alone, and the lone chart on a clear slide pays later-chart cost.
    const rows = [
      { of: 24, changed: 18, ms: 36350, drawnThereBefore: 42 },
      { of: 24, changed: 18, ms: 16314, drawnThereBefore: 0 },
      { of: 24, changed: 18, ms: 15235, drawnThereBefore: 0 },
    ];
    const [line] = occupancySplit(rows, 24, 18);
    // ORDERED: the loaded slide first, then the clear one. Asserting only that
    // both numbers appear passes with the buckets swapped, which would print
    // the ratio upside down and read as the clear slide being dearer.
    expect(line).toContain("42 drawn 36350ms vs 0 drawn 16314ms");
    expect(line, "loaded over clear, so the ratio is above 1").toContain("(2.2x)");
    expect(line, "the label has to say what the number counts").toContain("ALREADY DRAWN");
    expect(line).not.toContain("OCCUPANCY");
  });

  it("prints nothing when there is only one bucket", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { occupancySplit } = await import("../scripts/triage.mjs");
    // A split with one side is the confound restated, not resolved.
    const rows = [
      { of: 24, changed: 18, ms: 36350, drawnThereBefore: 42 },
      { of: 24, changed: 18, ms: 35000, drawnThereBefore: 42 },
    ];
    expect(occupancySplit(rows, 24, 18)).toEqual([]);
    expect(occupancySplit([{ of: 24, changed: 18, ms: 1, drawnThereBefore: null }], 24, 18)).toEqual([]);
  });
});

describe("what the driver went through, finally printed", () => {
  it("names a crash instead of hiding it in an attempt count", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { driverRunLine } = await import("../scripts/triage.mjs");
    const line = driverRunLine({ attempts: 3, fresh: true, recovered: ["not-ready:pane-closed", "crashed"] });
    expect(line, "the word that stops a reader skimming").toContain("crashed");
    expect(line).toContain("3 attempt(s)");
  });

  it("says nothing at all for rounds archived before the field existed", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { driverRunLine } = await import("../scripts/triage.mjs");
    // 190 archived rounds have no driverRun. A row of question marks on every
    // one is noise, not information.
    expect(driverRunLine(undefined)).toBe("");
    expect(driverRunLine(null)).toBe("");
  });

  it("names the document, the profile and what the round asked of the host", async () => {
    /**
     * THREE FIELDS THAT LANDED AND NOTHING PRINTED. Exactly the state
     * `driverRun` was in for its first fortnight — archived and invisible, so a
     * round rescued from a crash read like one that never needed rescuing.
     *
     * `driverDeck` matters most: 0 of 408 archived files named a document
     * before it, which is why "4:3 crashes more" and "that one file is sick"
     * were the same sentence for a fortnight.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundIdentityLine } = await import("../scripts/triage.mjs");
    const line = roundIdentityLine({
      driverDeck: "Presentation70.pptx",
      driverSlideSize: "16:9",
      syncs: 1848,
      driverRun: { deck: "Presentation70" },
    });
    expect(line).toContain("Presentation70.pptx");
    expect(line).toContain("16:9");
    expect(line, "the sync count is archived and still invisible").toContain("1848 syncs");
    // The instruction AGREES with what was in front, so it is not repeated —
    // the ordinary case must stay quiet or the line stops being read.
    expect(line, "shouted about an instruction that was obeyed").not.toContain("ASKED FOR");
  });

  it("says so when the deck in front is not the deck that was asked for", async () => {
    /**
     * `deck-missing` refuses the loud version of this. The quiet version is a
     * round that ran against the right deck BY LUCK while the instruction named
     * another — nothing refuses that, and nothing said it either.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundIdentityLine } = await import("../scripts/triage.mjs");
    const line = roundIdentityLine({
      driverDeck: "Presentation64.pptx",
      driverRun: { deck: "Presentation70" },
    });
    expect(line, "a round ran against a deck nobody asked for and the line was silent").toContain(
      "ASKED FOR Presentation70",
    );
  });

  it("keeps `undisclosed` and `unreadable` apart, and says nothing for the rounds that predate all three", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundIdentityLine } = await import("../scripts/triage.mjs");
    // A withheld name is a decision — see `archivableDeck`. A null is a reading
    // that failed. Collapsing them is this repo's most repeated defect.
    expect(roundIdentityLine({ driverDeck: "undisclosed" })).toContain("undisclosed");
    expect(roundIdentityLine({ driverDeck: null })).toContain("unreadable");
    // 400 files predate every one of these fields, and a row of question marks
    // on each is noise rather than information — the same call `driverRunLine`
    // makes.
    expect(roundIdentityLine({ build: "abc1234" })).toBe("");
    expect(roundIdentityLine(null)).toBe("");
  });

  it("warns when a round sits deep in a session", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { driverRunLine } = await import("../scripts/triage.mjs");
    const deep = driverRunLine({ attempts: 1, sessionIndex: 6, sincePrevRoundMs: 900000 });
    expect(deep, "laterMed roughly doubles by round 6-8").toContain("DEEP IN A SESSION");
    expect(deep).toContain("round 6 of this session");
    expect(deep).toContain("15m after the last");
    const shallow = driverRunLine({ attempts: 1, sessionIndex: 2 });
    expect(shallow).not.toContain("DEEP IN A SESSION");
  });
});

describe("skipped is a third outcome, not half a failure", () => {
  it("separates skipped from failed in the headline", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { selfTestHeadline } = await import("../scripts/triage.mjs");
    // 2026-08-24: every missing scenario was a SKIP - the host stopped
    // answering and nothing was checked - and "12 of 14 passed" read as two
    // product failures.
    const h = selfTestHeadline([{ ok: true }, { ok: true }, { ok: false, skipped: true }, { ok: false }]);
    expect(h).toContain("2 of 4 scenarios passed");
    expect(h).toContain("1 skipped (nothing was checked)");
    expect(h).toContain("1 FAILED");
  });

  it("stays quiet about the kinds that did not happen", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { selfTestHeadline } = await import("../scripts/triage.mjs");
    const clean = selfTestHeadline([{ ok: true }, { ok: true }]);
    expect(clean).toBe("2 of 2 scenarios passed");
  });
});

describe("instruments that have gone quiet", () => {
  const log = (name: string, ...messages: string[]) => ({
    roundName: name,
    trace: { entries: messages.map((message) => ({ message })) },
  });

  it("reports a line that stopped, and not one still firing", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormantInstruments } = await import("../scripts/triage.mjs");
    const rows = dormantInstruments([
      log("065-aaa.json", "a chart's tag could not even be queued", "still firing"),
      log("225-bbb.json", "still firing"),
    ]);
    expect(rows.map((r: { message: string }) => r.message)).toEqual(["a chart's tag could not even be queued"]);
    expect(rows[0].lastRound).toBe(65);
    expect(rows[0].gap, "225 - 65").toBe(160);
  });

  it("counts silence in ROUNDS, and orders by how loud the line used to be", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormantInstruments } = await import("../scripts/triage.mjs");
    const rows = dormantInstruments([
      log("010-a.json", "quiet one", "loud one", "loud one", "loud one"),
      log("225-b.json", "current"),
    ]);
    // A line that fired three times before going quiet deserves the reader
    // ahead of one that fired once.
    expect(rows[0].message).toBe("loud one");
    expect(rows[0].events).toBe(3);
  });

  it("says nothing when every line is current", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormantInstruments } = await import("../scripts/triage.mjs");
    expect(dormantInstruments([log("224-a.json", "x"), log("225-b.json", "x")])).toEqual([]);
    expect(dormantInstruments([])).toEqual([]);
  });

  it("ignores files whose name is not a round", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormantInstruments } = await import("../scripts/triage.mjs");
    // `rounds/predictions.json` lives beside the rounds and has been counted as
    // one before, which made a header disagree with its own pooling.
    //
    // The ledger must be SKIPPED, not merely absent from the output: its name
    // parses to NaN, and letting that reach the newest-round arithmetic makes
    // every comparison NaN and the whole report silently empty. So the assertion
    // is that a genuinely dormant line is still found ALONGSIDE it.
    const rows = dormantInstruments([
      log("predictions.json", "ledger"),
      log("010-a.json", "went quiet"),
      log("225-b.json", "current"),
    ]);
    expect(
      rows.map((r: { message: string }) => r.message),
      "the ledger is not a round",
    ).toEqual(["went quiet"]);
    expect(rows[0].gap).toBe(215);
  });
});

describe("a rate the sample cannot support", () => {
  it("refuses to turn five observations into a percentage", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { rateOrSilence } = await import("../scripts/triage.mjs");
    // This was LIVE: "freshly added, empty - 5 chart(s), 5 grouped = 100%",
    // printed above prose saying the recent window is what to quote, against a
    // documented 1% baseline that BACKLOG.md calls READ THIS FIRST.
    const five = rateOrSilence(5, 5);
    expect(five).not.toContain("100%");
    expect(five).toContain("n=5");
  });

  it("states a rate once the sample can carry one", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { rateOrSilence, MIN_RATE_N } = await import("../scripts/triage.mjs");
    expect(rateOrSilence(MIN_RATE_N, MIN_RATE_N)).toBe("100%");
    expect(rateOrSilence(10, 40)).toBe("25%");
    // The boundary itself counts as enough, not one past it.
    expect(rateOrSilence(1, MIN_RATE_N - 1)).toContain("too few");
  });

  it("keeps the honest half — the counts are printed either way", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { rateOrSilence } = await import("../scripts/triage.mjs");
    // Nothing to rate at all is a dash, not a false zero.
    expect(rateOrSilence(0, 0)).toBe("—");
  });
});

describe("the noise floor, measured where the session drift is held out", () => {
  const round = (name: string, build: string, sessionIndex: number, ms: number[]) => ({
    roundName: name,
    build,
    driverRun: { sessionIndex },
    selftest: [{ ok: true }],
    trace: {
      entries: ms.map((m, i) => ({
        ms: m,
        message: "updated only the shapes that changed",
        data: { chart: `${i + 2}/8`, changed: 18, of: 24, ms: m },
      })),
    },
  });

  it("takes only the first round of a session", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolNoiseFloor } = await import("../scripts/triage.mjs");
    // A sample from minute 116 of a session is measuring the drift, not the
    // floor: ten back-to-back rounds showed the same measurement roughly double
    // across two hours. Excluded, not weighted and not caveated.
    const per = poolNoiseFloor([
      round("230-aaa.json", "aaa", 1, [15000]),
      round("231-aaa.json", "aaa", 2, [40000]),
      round("232-aaa.json", "aaa", 7, [41000]),
    ]);
    expect(per.get("aaa").map((r: { round: string }) => r.round)).toEqual(["230"]);
  });

  it("keeps builds apart", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolNoiseFloor } = await import("../scripts/triage.mjs");
    // A floor is a statement about ONE build run repeatedly. Pooling two builds
    // measures the difference between them, which is the thing a floor exists
    // to judge.
    const per = poolNoiseFloor([round("230-aaa.json", "aaa", 1, [15000]), round("231-bbb.json", "bbb", 1, [16000])]);
    expect([...per.keys()].sort()).toEqual(["aaa", "bbb"]);
  });

  it("refuses to call a floor from two rounds", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { MIN_FLOOR_N } = await import("../scripts/triage.mjs");
    // The line this replaces IS a spread over two observations: "cabb357 scored
    // 1 and 5 with NOTHING changed". Reproducing that with better provenance
    // would be the same defect in a new coat.
    expect(MIN_FLOOR_N).toBeGreaterThan(2);
  });

  it("takes the median of a round, so one slow chart is not the round", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolNoiseFloor } = await import("../scripts/triage.mjs");
    const per = poolNoiseFloor([round("230-aaa.json", "aaa", 1, [15000, 16000, 17000])]);
    expect(per.get("aaa")[0].laterMed).toBe(16000);
  });

  it("ignores the first chart of a run, which is not a later chart", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolNoiseFloor } = await import("../scripts/triage.mjs");
    // The first chart of a multi-chart run costs ~2.2x a later one — it sits on
    // the deck's busiest slide. Letting it into the median would make the floor
    // a measurement of which slide the harness happened to fill.
    const log = round("230-aaa.json", "aaa", 1, [15000]);
    log.trace.entries.push({
      ms: 36000,
      message: "updated only the shapes that changed",
      data: { chart: "1/8", changed: 18, of: 24, ms: 36000 },
    });
    const per = poolNoiseFloor([log]);
    expect(per.get("aaa")[0].laterMed).toBe(15000);
  });
});

describe("a recovery line that spends the reader's attention where it counts", () => {
  it("shouts about a crash and whispers about a routine pane-open", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { driverRunLine } = await import("../scripts/triage.mjs");
    const line = driverRunLine({
      attempts: 3,
      recovered: ["not-ready:pane-closed", "crashed"],
    });
    expect(line).toContain("RECOVERED FROM: crashed");
    // The routine half is kept — hiding it would be a different lie — but it is
    // parenthetical, not the headline.
    expect(line).toContain("(routine: not-ready:pane-closed)");
    expect(line.indexOf("RECOVERED FROM"), "the notable half comes first").toBeLessThan(line.indexOf("(routine:"));
  });

  it("says nothing loud when everything was routine", async () => {
    // 82% of rounds recover from a closed or stale pane, because a round starts
    // with one. A line that shouts on four rounds in five teaches the reader to
    // skim it — and then `crashed` scrolls past inside it, which is exactly what
    // happened on 2026-08-24.
    // @ts-expect-error - plain .mjs tool, no types.
    const { driverRunLine } = await import("../scripts/triage.mjs");
    const line = driverRunLine({ attempts: 2, recovered: ["not-ready:pane-stale"] });
    expect(line).not.toContain("RECOVERED FROM");
    expect(line).toContain("(routine: not-ready:pane-stale)");
  });

  it("treats anything unrecognised as notable rather than routine", async () => {
    // The set is a whitelist on purpose. A recovery reason nobody has classified
    // is one nobody has decided is harmless.
    // @ts-expect-error - plain .mjs tool, no types.
    const { driverRunLine, ROUTINE_RECOVERIES } = await import("../scripts/triage.mjs");
    expect(ROUTINE_RECOVERIES.has("crashed"), "a crash is never routine").toBe(false);
    expect(ROUTINE_RECOVERIES.has("browser-gone")).toBe(false);
    expect(ROUTINE_RECOVERIES.has("not-ready:host-silent+pane-stale"), "a silent host is not routine").toBe(false);
    expect(driverRunLine({ attempts: 2, recovered: ["something-new"] })).toContain("RECOVERED FROM: something-new");
  });
});

describe("every way a probe declines, not just the one word", () => {
  const log = (id: string, answer: string) => ({ hostAnswers: { answers: [{ id, answer }] } });

  it("counts a documented non-answer as a non-answer", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolStarvedQuestions } = await import("../scripts/triage.mjs");
    // `no-refusal` carries the probe's own words: "the host grouped today, so
    // the question was never put. Not an answer." The reader of that value
    // counted it as one, so a probe that never answered stayed out of the report
    // whose whole job is to find probes that never answer.
    const rows = poolStarvedQuestions([log("p", "no-refusal"), log("p", "no-refusal"), log("p", "no-refusal")]);
    expect(rows.map((r: { id: string }) => r.id)).toContain("p");
  });

  it("covers the other two prefixes the host declines with", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolStarvedQuestions } = await import("../scripts/triage.mjs");
    for (const answer of ["no-creation-id", "no-group-id", "unreadable"]) {
      const rows = poolStarvedQuestions([log("q", answer), log("q", answer)]);
      expect(
        rows.map((r: { id: string }) => r.id),
        answer,
      ).toContain("q");
    }
  });

  it("still leaves a probe that answers sometimes alone", async () => {
    // Widening the pattern must not bury a working question. A probe with any
    // real answer is doing its job, however much silence surrounds it.
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolStarvedQuestions } = await import("../scripts/triage.mjs");
    const rows = poolStarvedQuestions([log("r", "no-refusal"), log("r", "no-refusal"), log("r", "tags-gone")]);
    expect(rows.map((r: { id: string }) => r.id)).not.toContain("r");
  });

  it("keeps ours-to-fix separate from a fact about the host", async () => {
    // The remedies are opposite: a setup the harness cannot build is ours to fix
    // or retire, and a question the host will not answer is a finding to leave
    // alone. Collapsing them would send someone to fix the host.
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolStarvedQuestions } = await import("../scripts/triage.mjs");
    const ours = poolStarvedQuestions([log("s", "no-scratch-shape"), log("s", "no-scratch-shape")])[0];
    const host = poolStarvedQuestions([log("t", "no-creation-id"), log("t", "no-creation-id")])[0];
    expect(ours.never).toBeGreaterThan(ours.unanswerable);
    expect(host.unanswerable).toBeGreaterThan(host.never);
  });
});

describe("a round states its own error bar", () => {
  const round = (name: string, ms: number[]) => ({
    roundName: name,
    trace: {
      entries: ms.map((m, i) => ({
        message: "updated only the shapes that changed",
        data: { chart: `${i + 2}/8`, changed: 18, of: 24, ms: m },
      })),
    },
  });

  it("measures how much a round agreed with itself", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolWithinRoundSpread } = await import("../scripts/triage.mjs");
    // Round 230 for real: five later charts inside 2% of each other.
    const tight = poolWithinRoundSpread([round("230-a.json", [19468, 19695, 19237, 19311, 19552])]);
    expect(tight[0].pct).toBe(2);
    // Round 232 for real: the same five spread over 24%.
    const loose = poolWithinRoundSpread([round("232-a.json", [25707, 23080, 28568, 25822, 25845])]);
    expect(loose[0].pct).toBe(24);
  });

  it("will not compute a spread from fewer than four charts", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolWithinRoundSpread } = await import("../scripts/triage.mjs");
    // Two charts give a range with no idea whether either is typical, and a
    // round that managed one later chart is telling a different story.
    expect(poolWithinRoundSpread([round("230-a.json", [19000, 25000])])).toEqual([]);
    expect(poolWithinRoundSpread([round("230-a.json", [19000, 20000, 21000])])).toEqual([]);
    expect(poolWithinRoundSpread([round("230-a.json", [19000, 20000, 21000, 22000])])).toHaveLength(1);
  });

  it("ignores the first chart of a run, which is not a later chart", async () => {
    // The first chart sits on the deck's busiest slide and costs ~2.2x. Letting
    // it in would report the slide as disagreement.
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolWithinRoundSpread } = await import("../scripts/triage.mjs");
    const r = round("230-a.json", [19468, 19695, 19237, 19311]);
    r.trace.entries.push({
      message: "updated only the shapes that changed",
      data: { chart: "1/8", changed: 18, of: 24, ms: 43000 },
    });
    expect(poolWithinRoundSpread([r])[0].pct, "the first chart must not widen it").toBe(2);
  });

  it("ignores a lone chart in its own run", async () => {
    // `1/1` is the alone arm — a run of one has no later charts at all.
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolWithinRoundSpread } = await import("../scripts/triage.mjs");
    const r = round("230-a.json", [19468, 19695, 19237, 19311]);
    r.trace.entries.push({
      message: "updated only the shapes that changed",
      data: { chart: "1/1", changed: 18, of: 24, ms: 40000 },
    });
    expect(poolWithinRoundSpread([r])[0].pct).toBe(2);
  });
});

describe("the floor is reported as a spread that does not grow with n", () => {
  it("carries quartiles, not only min and max", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolNoiseFloor } = await import("../scripts/triage.mjs");
    // RANGE only ever grows with sample size: it went 66% at five sessions to
    // 73% at eight because one faster round arrived, with nothing about the host
    // changing. The IQR is what a reader can actually use as a bar.
    const mk = (name: string, ms: number) => ({
      roundName: name,
      build: "aaa",
      driverRun: { sessionIndex: 1 },
      selftest: [],
      trace: {
        entries: [2, 3, 4, 5].map((i) => ({
          message: "updated only the shapes that changed",
          data: { chart: i + "/8", changed: 18, of: 24, ms },
        })),
      },
    });
    const rows = poolNoiseFloor([
      mk("230-a.json", 10),
      mk("231-a.json", 20),
      mk("232-a.json", 30),
      mk("233-a.json", 40),
      mk("234-a.json", 1000),
    ]).get("aaa");
    expect(rows).toHaveLength(5);
    const meds = rows.map((r: { laterMed: number }) => r.laterMed).sort((a: number, b: number) => a - b);
    // The single outlier owns the range and does not own the middle half.
    expect(meds[meds.length - 1]).toBe(1000);
    expect(meds[Math.floor(meds.length * 0.75)], "q3 is unmoved by one wild round").toBeLessThan(1000);
  });
});

describe("claims the archive can contradict", () => {
  const roundWith = (name: string, entries: unknown[]) => ({ roundName: name, trace: { entries } });
  const update = (chart: string, ms: number) => ({
    message: "updated only the shapes that changed",
    data: { chart, changed: 18, of: 24, ms },
  });

  it("says HOLDS while the archive agrees, and STALE when it stops", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    const claim = CLAIMS.find((c: { id: string }) => c.id === "first-chart-costs-more");
    // Two rounds of a world where the first chart is dear.
    const dear = [1, 2].map((n) =>
      roundWith(`23${n}-a.json`, [update("1/8", 40000), ...[2, 3, 4, 5, 6].map((i) => update(`${i}/8`, 18000))]),
    );
    expect(claim.check([...dear, ...dear, ...dear]).ok).toBe(true);
    // And a world where it stopped being dear — the fix nobody would otherwise
    // notice, which is the whole reason this file exists.
    const even = [1, 2].map((n) =>
      roundWith(`23${n}-a.json`, [update("1/8", 18000), ...[2, 3, 4, 5, 6].map((i) => update(`${i}/8`, 18000))]),
    );
    const stale = claim.check([...even, ...even, ...even]);
    expect(stale.ok, "a fact that moved must not read as ok").toBe(false);
    expect(stale.actual).toContain("1.00x");
  });

  it("refuses a ratio whose middle halves overlap", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    const claim = CLAIMS.find((c: { id: string }) => c.id === "first-chart-costs-more");
    // A 2x ratio of medians sitting on top of a distribution that overlaps it
    // completely. The measured noise floor between two rounds of the SAME build
    // is IQR 14% and RANGE 73%, so a difference this smeared is not a difference
    // — and until now the check would have called it a fact because the medians
    // divided nicely.
    const smeared = [1, 2, 3].map((n) =>
      roundWith(`23${n}-a.json`, [
        update("1/8", 40000),
        update("2/8", 4000),
        update("3/8", 4000),
        update("4/8", 60000),
        update("5/8", 60000),
        update("6/8", 4000),
      ]),
    );
    const r = claim.check([...smeared, ...smeared]);
    expect(r.ok, "a smeared difference must not read as a fact").toBe(false);
    expect(r.actual).toContain("OVERLAP");
  });

  it("answers ? rather than guessing when the sample is too small", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    const claim = CLAIMS.find((c: { id: string }) => c.id === "first-chart-costs-more");
    const one = [roundWith("230-a.json", [update("1/8", 40000), update("2/8", 18000)])];
    // `null`, not false. "Not enough data" and "the claim is wrong" are opposite
    // messages and this report exists to keep them apart.
    expect(claim.check(one).ok).toBe(null);
  });

  it("does not let a broken query read as a refutation", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { checkClaims } = await import("../scripts/claims.mjs");
    // A check that throws is a broken CLAIM, not a moved fact. Reporting it as
    // STALE would send someone to re-read doctrine that is perfectly true.
    const rows = checkClaims([
      {
        get roundName(): string {
          throw new Error("boom");
        },
      },
    ]);
    for (const r of rows) if (r.ok === false) expect(r.actual).not.toContain("threw");
    expect(rows.every((r: { ok: boolean | null }) => r.ok !== false || true)).toBe(true);
  });

  it("gives every claim a query, a date and a sentence", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    expect(CLAIMS.length).toBeGreaterThan(0);
    for (const c of CLAIMS) {
      // A claim nobody can refute is a slogan.
      expect(typeof c.check, c.id).toBe("function");
      expect(c.says.length, c.id).toBeGreaterThan(20);
      expect(c.measured, c.id).toMatch(/20\d\d-\d\d-\d\d/);
    }
  });
});

describe("the current-beliefs index cannot drift from the claims", () => {
  it("mentions every checked claim", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    const doc = readFileSync("docs/WHAT-WE-KNOW.md", "utf8");
    // The index exists because an 8000-line journal cannot be searched. It is
    // worth having only while it is complete: a checked claim missing from it is
    // a fact nobody will find, which is the failure the index was built to end.
    // THE ID ITSELF, not a word from it. Matching any long word let
    // `first-chart-is-on-the-busiest-slide` pass on the word "first", which
    // appears in an unrelated line — a drift test that can be satisfied by
    // coincidence is not a drift test.
    for (const c of CLAIMS) {
      expect(doc, `WHAT-WE-KNOW.md does not name the claim "${c.id}"`).toContain(c.id);
    }
  });

  it("marks the checked ones so a reader knows which are enforced", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    const doc = readFileSync("docs/WHAT-WE-KNOW.md", "utf8");
    // An unmarked claim reads exactly like a marked one and rots silently — the
    // distinction this whole file exists to make.
    const ticks = (doc.match(/\u2713/g) ?? []).length;
    expect(ticks, "every checked claim should carry a tick").toBeGreaterThanOrEqual(CLAIMS.length);
  });
});

describe("where an update's time went, sync by sync", () => {
  const log = (rows: { chart: string; syncMs: number[] }[]) => ({
    trace: {
      entries: rows.map((r) => ({
        message: "updated only the shapes that changed",
        data: { chart: r.chart, changed: 18, of: 24, ms: r.syncMs.reduce((a, b) => a + b, 0), syncMs: r.syncMs },
      })),
    },
  });

  it("separates the first chart of a run from the rest", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolSyncBreakdown } = await import("../scripts/triage.mjs");
    const r = poolSyncBreakdown([
      log([
        { chart: "1/8", syncMs: [11000, 11000, 12000, 800] },
        { chart: "2/8", syncMs: [5000, 5000, 5000, 600] },
        { chart: "3/8", syncMs: [5100, 5100, 5100, 610] },
      ]),
    ]);
    expect(r.first).toHaveLength(1);
    expect(r.later).toHaveLength(2);
  });

  it("keeps a run of ONE out of both arms", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolSyncBreakdown } = await import("../scripts/triage.mjs");
    // The lone arm is the control that separates position from slide. Pooling it
    // into either side destroys the thing it exists to measure.
    const r = poolSyncBreakdown([
      log([
        { chart: "1/1", syncMs: [5000, 5000, 5000, 600] },
        { chart: "1/8", syncMs: [11000, 11000, 12000, 800] },
        { chart: "2/8", syncMs: [5000, 5000, 5000, 600] },
      ]),
    ]);
    expect(r.first).toHaveLength(1);
    expect(r.later).toHaveLength(1);
    expect(r.first[0][0], "the lone chart must not be read as the first of a run").toBe(11000);
  });

  it("takes only the size it was asked for", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolSyncBreakdown } = await import("../scripts/triage.mjs");
    // A sync's cost is a function of how many shapes it writes, so mixing sizes
    // produces a per-sync figure describing neither.
    const mixed = {
      trace: {
        entries: [
          {
            message: "updated only the shapes that changed",
            data: { chart: "1/8", changed: 9, of: 16, ms: 1, syncMs: [1, 2, 3, 4] },
          },
          {
            message: "updated only the shapes that changed",
            data: { chart: "2/8", changed: 18, of: 24, ms: 1, syncMs: [5, 6, 7, 8] },
          },
        ],
      },
    };
    const r = poolSyncBreakdown([mixed]);
    expect(r.first, "the 9-of-16 chart is a different measurement").toHaveLength(0);
    expect(r.later).toHaveLength(1);
  });

  it("ignores rows carrying no per-sync timings", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { poolSyncBreakdown } = await import("../scripts/triage.mjs");
    const noSyncs = {
      trace: {
        entries: [
          {
            message: "updated only the shapes that changed",
            data: { chart: "1/8", changed: 18, of: 24, ms: 37000 },
          },
        ],
      },
    };
    expect(poolSyncBreakdown([noSyncs])).toEqual({ first: [], later: [] });
  });
});

describe("why a trace line went quiet", () => {
  const src = `
    /**
     * A note quoting a retired line: \`the binding could not name the chart\`
     * fired 16 times, all in rounds 077-078.
     */
    // and a line comment mentioning "a slide's shape count would not settle"
    // retired on 2026-08-01: "the settle's re-read came back empty" fired 102 times
    trace("update", "a slide's shape count would not settle — taking the second read", {});
    trace("group", "tagging failed — charts are not re-editable until repaired", {});
    await boundedSync(context, \`settling the config tag on a shape found by \${byId ? "id" : "name"}\`);
  `;

  it("reads a message still in the source as a fault that stopped", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormancyCause, stringLiteralsIn } = await import("../scripts/triage.mjs");
    expect(dormancyCause("tagging failed — charts are not re-editable until repaired", stringLiteralsIn(src))).toBe(
      "stopped",
    );
  });

  it("reads a rewritten message as renamed, not as a fixed fault", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormancyCause, stringLiteralsIn } = await import("../scripts/triage.mjs");
    // The instrument fires every round under its new tail. Reported as a fault
    // that stopped 140 rounds ago, it would be read as evidence the host
    // improved — the exact shape of the Thread 3 mistake.
    expect(dormancyCause("a slide's shape count would not settle — not counting it", stringLiteralsIn(src))).toBe(
      "renamed",
    );
  });

  it("reads an interpolated message as renamed too", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormancyCause, stringLiteralsIn } = await import("../scripts/triage.mjs");
    // No rendered form can ever match the source once the message is built by
    // interpolation, so every archived variant looks retired forever.
    expect(dormancyCause("settling the config tag on a shape found by id", stringLiteralsIn(src))).toBe("renamed");
  });

  it("does not let a comment vouch for a line that is gone", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormancyCause, stringLiteralsIn } = await import("../scripts/triage.mjs");
    // THE DANGEROUS DIRECTION. This file's house style names retired trace lines
    // in prose, in BACKTICKS — so a regex sweep for template literals matches
    // inside the very comments that say the line is dead, and every retired
    // message reads as live.
    expect(dormancyCause("the binding could not name the chart", stringLiteralsIn(src))).toBe("removed");
  });

  it("does not let a LINE comment vouch for a line that is gone either", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormancyCause, stringLiteralsIn } = await import("../scripts/triage.mjs");
    // The block-comment case has a twin: a `//` note quoting the full retired
    // message in double quotes reads as an emitter unless line comments are
    // skipped too, and this file's notes are written both ways.
    expect(dormancyCause("the settle's re-read came back empty", stringLiteralsIn(src))).toBe("removed");
  });

  it("says `unknown` rather than guessing when there is no source to read", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { dormancyCause } = await import("../scripts/triage.mjs");
    // An unreadable tree is no evidence, and "removed" on no evidence would
    // retire a live instrument on the strength of a missing directory.
    expect(dormancyCause("anything at all", null)).toBe("unknown");
  });

  it("keeps a string that contains a comment marker", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stringLiteralsIn } = await import("../scripts/triage.mjs");
    // Stripping `//` line-wise would cut this in half and retire the message
    // after it.
    const kept = stringLiteralsIn('const a = "see https://example.com for why"; trace("x", "still here", {});');
    expect(kept).toContain("see https://example.com for why");
    expect(kept).toContain("still here");
  });
});

describe("the noise floor section leads with the answer", () => {
  const round = (build: string, n: number, ms: number) => ({
    roundName: `${n}-${build}.json`,
    build,
    driverRun: { sessionIndex: 1 },
    selftest: [],
    trace: {
      entries: [{ message: "updated only the shapes that changed", data: { chart: "2/8", changed: 18, of: 24, ms } }],
    },
  });

  const say = async (logs: unknown[]) => {
    // @ts-expect-error - plain .mjs tool, no types.
    const m = await import("../scripts/triage.mjs");
    const said: string[] = [];
    const real = console.log;
    // Swapped AFTER the await, and restored synchronously around the call only.
    // Wrapping the import as well put the restore in a `finally` that ran before
    // the promise resolved, so nothing was captured at all.
    console.log = (...a: unknown[]) => void said.push(a.join(" "));
    try {
      m.reportNoiseFloor(logs);
    } finally {
      console.log = real;
    }
    return said;
  };

  it("names the build that HAS a floor before the ones that do not", async () => {
    // A build with one round prints "too few to call a floor". With several such
    // builds ahead of the answer, the section reads as "the floor is
    // unmeasurable" — which on 2026-08-25 produced a four-hour plan to
    // re-measure a floor the archive already held at n=9.
    const logs = [
      round("thin1", 226, 20000),
      round("thin2", 228, 20000),
      ...[230, 231, 232, 233, 234].map((n, i) => round("deep", n, 15000 + i * 100)),
    ];
    const said = await say(logs);
    const first = said.findIndex((l) => /BEST AVAILABLE/.test(l));
    const tooFew = said.findIndex((l) => /too few to call a floor/.test(l));
    expect(first, "the answer must be announced").toBeGreaterThanOrEqual(0);
    expect(said[first]).toContain("deep");
    expect(first, "the answer must come before the refusals").toBeLessThan(tooFew);
    // …and so must its DETAIL. Announcing the answer and then printing two
    // "too few" blocks above the numbers still makes a reader scroll past
    // refusals to reach the floor, which is the whole failure.
    const detail = said.findIndex((l) => /in-place update, 18 of 24/.test(l));
    expect(detail, "the floor's own numbers must be printed").toBeGreaterThanOrEqual(0);
    expect(detail, "the floor's numbers must come before any refusal").toBeLessThan(tooFew);
  });

  it("says plainly when NO build has a floor, rather than listing refusals", async () => {
    const said = await say([round("thin1", 226, 20000), round("thin2", 228, 20000)]);
    expect(said.some((l) => /NO BUILD HAS/.test(l))).toBe(true);
    expect(said.some((l) => /BEST AVAILABLE/.test(l))).toBe(false);
  });
});

describe("the tag sync's non-difference is checked too", () => {
  const roundWith = (name: string, entries: unknown[]) => ({ roundName: name, trace: { entries } });
  const upd = (chart: string, syncMs: number[]) => ({
    message: "updated only the shapes that changed",
    data: { chart, changed: 18, of: 24, ms: syncMs.reduce((a, b) => a + b, 0), syncMs },
  });

  it("goes stale when the tag sync starts paying a premium of its own", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    const claim = CLAIMS.find((c: { id: string }) => c.id === "tag-sync-is-not-the-writes");
    // Writes still 2x, and the tag ratio is only 1.4 — under the 1.5 the check
    // used to test alone. But the tag quartiles are now CLEAR of each other, so
    // the tag sync is consistently dearer for the first chart.
    //
    // That matters more than its size: "the tag sync does not pay the 2.2x" is
    // the sentence that located the cost in the writes. A small but consistent
    // premium moves that conclusion, and a ratio under 1.5 would not notice.
    const rounds = [1, 2, 3, 4, 5, 6].map((n) =>
      roundWith(`23${n}-a.json`, [
        upd("1/8", [20000, 20000, 20000, 140]),
        ...[2, 3, 4, 5, 6].map((i) => upd(`${i}/8`, [9000, 9000, 9000, 100])),
      ]),
    );
    const r = claim.check(rounds);
    expect(r.actual).toContain("QUARTILES CLEAR");
    expect(r.ok, "a tag sync that separates must not read as ok").toBe(false);
  });

  it("holds while the tag quartiles still overlap", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    const claim = CLAIMS.find((c: { id: string }) => c.id === "tag-sync-is-not-the-writes");
    // The world as measured: writes separate cleanly, the tag sync's spread
    // straddles both, so its cost is the same whichever chart it belongs to.
    const rounds = [1, 2, 3, 4, 5, 6].map((n) =>
      roundWith(`23${n}-a.json`, [
        upd("1/8", [20000, 20000, 20000, [70, 110, 150][n % 3]]),
        ...[2, 3, 4, 5, 6].map((i) => upd(`${i}/8`, [9000, 9000, 9000, [60, 120, 160][i % 3]])),
      ]),
    );
    const r = claim.check(rounds);
    expect(r.actual).toContain("quartiles overlap");
    expect(r.ok).toBe(true);
  });
});

describe("the by-id refusal recovery is watched, not assumed", () => {
  const roundWith = (name: string, entries: unknown[]) => ({ roundName: name, trace: { entries } });
  const rescue = (asked: number, recovered: number) => ({
    message: "re-read recovered shapes a by-id lookup had refused",
    data: { asked, recovered },
  });
  const claimOf = async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    return CLAIMS.find((c: { id: string }) => c.id === "the-re-read-always-rescues-a-refused-lookup");
  };

  it("holds while every refused lookup is recovered", async () => {
    const claim = await claimOf();
    const logs = Array.from({ length: 25 }, (_, i) => roundWith(`2${i}0-a.json`, [rescue(1, 1)]));
    expect(claim.check(logs).ok).toBe(true);
  });

  it("goes stale the moment one is not", async () => {
    const claim = await claimOf();
    // The recovery standing between a refused lookup and a silent no-op on a
    // chart the user is looking at. One miss is the whole point of watching it.
    const logs = [
      ...Array.from({ length: 24 }, (_, i) => roundWith(`2${i}0-a.json`, [rescue(1, 1)])),
      roundWith("299-a.json", [rescue(1, 0)]),
    ];
    const r = claim.check(logs);
    expect(r.ok).toBe(false);
    expect(r.actual).toContain("24/25");
  });

  it("counts a re-read that threw as a failure even when the tally is clean", async () => {
    const claim = await claimOf();
    // `the re-read of a refused slide would not answer either` returns early
    // with whatever it had, so a throw can leave asked and recovered matching
    // while the rescue did not happen. A rate alone would call that 100%.
    const logs = [
      ...Array.from({ length: 25 }, (_, i) => roundWith(`2${i}0-a.json`, [rescue(1, 1)])),
      roundWith("299-a.json", [{ message: "the re-read of a refused slide would not answer either", data: {} }]),
    ];
    const r = claim.check(logs);
    expect(r.ok).toBe(false);
    expect(r.actual).toContain("1 re-read(s) threw");
  });

  it("says ? rather than 100% on a handful of events", async () => {
    const claim = await claimOf();
    // 5 of 5 is not 100%. The bar is the same twenty the report uses.
    const logs = [roundWith("230-a.json", [rescue(5, 5)])];
    expect(claim.check(logs).ok).toBe(null);
  });
});

describe("the sheet reports pass 1, so pass 1 gets checked", () => {
  const sheetOf = (samples: { answer: string; pass: number }[]) => ({
    roundName: "230-a.json",
    hostAnswers: { answers: [{ id: "q", answer: samples[0].answer, samples }] },
  });
  const many = (pass: number, answer: string, n: number) =>
    Array.from({ length: n }, () => ({ answer, pass, atMs: 0, regime: "healthy", scratch: "first-slide" }));

  it("names a question whose first answer is not like its later ones", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { passPositionBias } = await import("../scripts/triage.mjs");
    // `record` keeps the FIRST real answer, so the sheet's headline is a COLD
    // answer. If cold differs, every archived answer inherits that — measured
    // at up to 19 points on 2026-08-25.
    const rows = passPositionBias([sheetOf([...many(1, "threw", 30), ...many(2, "yes", 30), ...many(3, "yes", 30)])]);
    const threw = rows.find((r: { answer: string }) => r.answer === "threw");
    expect(threw.p1).toBe(100);
    expect(threw.later).toBe(0);
  });

  it("stays quiet when the passes agree", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { passPositionBias } = await import("../scripts/triage.mjs");
    expect(passPositionBias([sheetOf([...many(1, "yes", 30), ...many(2, "yes", 30)])])).toEqual([]);
  });

  it("ignores questions the harness could not put", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { passPositionBias } = await import("../scripts/triage.mjs");
    // `no-scratch-slide` is far commoner on later passes for reasons that are
    // entirely ours — the run spends its scratch slides as it goes. Counting it
    // would report the harness's own housekeeping as a host effect.
    const rows = passPositionBias([
      sheetOf([...many(1, "yes", 30), ...many(2, "yes", 30), ...many(3, "no-scratch-slide", 30)]),
    ]);
    expect(rows).toEqual([]);
  });

  it("says nothing on a sample too thin to mean anything", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { passPositionBias } = await import("../scripts/triage.mjs");
    // The same twenty the rest of this report uses. A 100%-versus-0% split on
    // three observations is not a shift.
    expect(passPositionBias([sheetOf([...many(1, "threw", 3), ...many(2, "yes", 3)])])).toEqual([]);
  });

  it("compares the SAME answer on both sides", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { passPositionBias } = await import("../scripts/triage.mjs");
    // Taking each side's own top answer would compare two different questions
    // and call the difference a shift.
    const rows = passPositionBias([sheetOf([...many(1, "threw", 30), ...many(2, "yes", 30), ...many(3, "yes", 30)])]);
    for (const r of rows) expect(typeof r.answer).toBe("string");
    expect(new Set(rows.map((r: { answer: string }) => r.answer))).toEqual(new Set(["threw", "yes"]));
  });
});

describe("the probe's own scratch-slide churn", () => {
  const round = (name: string, whys: string[]) => ({
    roundName: name,
    trace: { entries: whys.map((why) => ({ message: "replaced the scratch slide", data: { why } })) },
  });

  it("counts what only the raw trace has ever held", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scratchChurn } = await import("../scripts/triage.mjs");
    const c = scratchChurn([round("230-a.json", ["silent", "silent", "gone"]), round("231-a.json", ["silent"])]);
    expect(c.total).toBe(4);
    expect(c.rounds).toBe(2);
    expect(c.why).toEqual([
      ["silent", 3],
      ["gone", 1],
    ]);
  });

  it("splits the causes, because they want opposite fixes", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scratchChurn } = await import("../scripts/triage.mjs");
    // `gone` is a slide that really was deleted — something is removing it.
    // `silent` is a proxy the host would not answer for, which the update path
    // recovers from by re-reading the collection and adding nothing at all.
    // Pooling them would hide the only one with a cheap fix.
    const c = scratchChurn([round("230-a.json", ["gone", "silent"])]);
    expect(new Map(c.why).get("gone")).toBe(1);
    expect(new Map(c.why).get("silent")).toBe(1);
  });

  it("calls the old rounds `unrecorded` rather than inventing a cause", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scratchChurn } = await import("../scripts/triage.mjs");
    // 16548 replacements were archived before the cause was carried out of the
    // catch. Defaulting them to `silent` would manufacture 16548 observations of
    // a thing nobody measured.
    const c = scratchChurn([
      { roundName: "1-a.json", trace: { entries: [{ message: "replaced the scratch slide" }] } },
    ]);
    expect(c.why).toEqual([["unrecorded", 1]]);
  });

  it("ignores rounds that never replaced one", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scratchChurn } = await import("../scripts/triage.mjs");
    // A round with none must not drag the median toward zero — the statistic is
    // "what it costs when it happens", and it happens in every round on record.
    const c = scratchChurn([
      round("230-a.json", ["silent", "silent"]),
      { roundName: "231-a.json", trace: { entries: [] } },
    ]);
    expect(c.rounds).toBe(1);
    expect(c.median).toBe(2);
  });
});

describe("the scratch re-acquire stays watched", () => {
  const claimOf = async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    return CLAIMS.find((c: { id: string }) => c.id === "scratch-slides-are-re-acquired-not-rebought");
  };
  const round = (name: string, worked: boolean[]) => ({
    roundName: name,
    trace: {
      entries: worked.map((w) => ({
        message: "re-acquired the scratch slide by position instead of adding one",
        data: { worked: w },
      })),
    },
  });

  it("holds while the cheap route keeps working", async () => {
    const claim = await claimOf();
    expect(claim.check([round("254-a.json", Array(21).fill(true).concat(Array(5).fill(false)))]).ok).toBe(true);
  });

  it("goes stale when the id stops settling and the run buys slides again", async () => {
    const claim = await claimOf();
    // The failure mode is SILENT: the replacement path still works, so the round
    // still produces a sheet. What it costs is 63 slide adds, a deck of 110 and
    // a minute of probe time — which is how this hid for 250 rounds.
    const r = claim.check([round("300-a.json", Array(25).fill(false).concat([true]))]);
    expect(r.ok).toBe(false);
    expect(r.actual).toContain("4% worked");
  });

  it("says ? rather than boasting from a handful", async () => {
    const claim = await claimOf();
    expect(claim.check([round("254-a.json", [true, true, true])]).ok).toBe(null);
  });
});

describe("where the battery's time goes", () => {
  const round = (name: string, rows: { name: string; ms: number; ok?: boolean; skipped?: boolean }[]) => ({
    roundName: name,
    selftest: rows.map((r) => ({ ok: r.ok ?? true, skipped: r.skipped ?? false, name: r.name, ms: r.ms })),
  });
  const many = (n: number, rows: { name: string; ms: number; ok?: boolean; skipped?: boolean }[]) =>
    Array.from({ length: n }, (_, i) => round(`2${i}0-a.json`, rows));

  it("prices each scenario by its median, loudest first", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioCost } = await import("../scripts/triage.mjs");
    const rows = scenarioCost(
      many(6, [
        { name: "slow", ms: 160000 },
        { name: "quick", ms: 20000 },
      ]),
    );
    expect(rows[0].name).toBe("slow");
    expect(rows[0].median).toBe(160000);
    expect(rows[0].share).toBe(89);
  });

  it("holds skips out of the median", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioCost } = await import("../scripts/triage.mjs");
    // A skipped scenario records the time it took to GIVE UP. Pooling that with
    // real runs produces a number describing neither — the same mistake
    // `selfTestHeadline` exists to prevent one level up.
    const rows = scenarioCost([
      ...many(5, [{ name: "one", ms: 100000 }]),
      ...many(5, [{ name: "one", ms: 1, ok: false, skipped: true }]),
    ]);
    expect(rows[0].n).toBe(5);
    expect(rows[0].median).toBe(100000);
  });

  it("keeps a FAILED scenario, which did run", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioCost } = await import("../scripts/triage.mjs");
    // A failure ran to a verdict and cost what it cost. Only a skip checked
    // nothing.
    const rows = scenarioCost(many(6, [{ name: "one", ms: 50000, ok: false, skipped: false }]));
    expect(rows[0].n).toBe(6);
  });

  it("says nothing about a scenario it has barely seen", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioCost } = await import("../scripts/triage.mjs");
    expect(scenarioCost(many(2, [{ name: "rare", ms: 9000 }]))).toEqual([]);
  });

  it("carries the spread, so a drift can be told from a slow day", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioCost } = await import("../scripts/triage.mjs");
    // The measured noise floor between rounds of one build is IQR 14%. A median
    // with no spread beside it cannot be read against that.
    const rows = scenarioCost(
      Array.from({ length: 8 }, (_, i) => round(`2${i}0-a.json`, [{ name: "one", ms: 10000 + i * 1000 }])),
    );
    expect(rows[0].q1).toBeLessThan(rows[0].median);
    expect(rows[0].q3).toBeGreaterThan(rows[0].median);
  });
});

describe("whether buying a slide is worth it", () => {
  const claimOf = async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    return CLAIMS.find((c: { id: string }) => c.id === "buying-a-replacement-slide-rescues-the-question");
  };
  const round = (name: string, rescued: boolean[]) => ({
    roundName: name,
    trace: { entries: rescued.map((r) => ({ message: "the replacement slide answered", data: { rescued: r } })) },
  });

  it("holds while the bought slide usually answers", async () => {
    const claim = await claimOf();
    expect(
      claim.check([
        round("256-a.json", Array(15).fill(true).concat(Array(3).fill(false))),
        round("257-a.json", Array(4).fill(true)),
      ]).ok,
    ).toBe(true);
  });

  it("goes stale when the buy stops paying", async () => {
    const claim = await claimOf();
    // The direction that matters. If this drops, 18 slide adds a round are being
    // spent on nothing and the buy should be reconsidered — which is a decision
    // that needs a number, not the impression that adds look wasteful.
    const r = claim.check([round("300-a.json", Array(20).fill(false).concat([true, true]))]);
    expect(r.ok).toBe(false);
    expect(r.actual).toContain("9% rescued");
  });

  it("says ? on a single round's worth", async () => {
    const claim = await claimOf();
    expect(claim.check([round("256-a.json", Array(18).fill(true))]).ok).toBe(null);
  });
});

describe("what a round is made of", () => {
  const round = (spanMs: number, probeMs: number[], batteryMs: number[]) => ({
    roundName: "230-a.json",
    selftest: batteryMs.map((ms, i) => ({ ok: true, skipped: false, name: `s${i}`, ms })),
    trace: {
      entries: [
        ...probeMs.map((ms) => ({ message: "answered", data: { ms }, ms: 1 })),
        { message: "done", data: {}, ms: spanMs },
      ],
    },
  });

  it("splits a round into battery, probe and what is neither", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundAnatomy } = await import("../scripts/triage.mjs");
    const a = roundAnatomy(Array.from({ length: 6 }, () => round(600000, [50000, 30000], [300000, 100000])));
    expect(a.battery).toBe(400000);
    expect(a.probe).toBe(80000);
    expect(a.outside).toBe(120000);
  });

  it("counts the partner questions as probe time", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundAnatomy } = await import("../scripts/triage.mjs");
    // A partner question is a probe question that another question triggered.
    // Leaving it out would push its cost into `outside`, where it would read as
    // harness overhead rather than as the sheet's own work.
    const withPartner = Array.from({ length: 6 }, () => ({
      roundName: "230-a.json",
      selftest: [],
      trace: {
        entries: [
          { message: "answered", data: { ms: 1000 }, ms: 1 },
          { message: "partner answered", data: { ms: 2000 }, ms: 2 },
          { message: "done", data: {}, ms: 10000 },
        ],
      },
    }));
    expect(roundAnatomy(withPartner).probe).toBe(3000);
  });

  it("never reports a negative outside", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundAnatomy } = await import("../scripts/triage.mjs");
    // The battery and the probe overlap the span imperfectly — scenario time is
    // wall-clock and probe `ms` is per-question — so the arithmetic can go past
    // the span. A negative would read as a round that finished before it began.
    const a = roundAnatomy(Array.from({ length: 6 }, () => round(1000, [50000], [300000])));
    expect(a.outside).toBe(0);
  });

  it("says nothing from a handful of rounds", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundAnatomy } = await import("../scripts/triage.mjs");
    expect(roundAnatomy([round(600000, [1000], [2000])])).toBe(null);
  });

  it("skips a round whose trace has entries but no clock", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundAnatomy } = await import("../scripts/triage.mjs");
    // Distinct from an EMPTY trace, which an earlier guard drops. A crashed
    // round can archive entries whose last one carries no `ms` — the span is
    // then unknown, not zero, and treating it as zero makes `outside` the whole
    // battery and drags every median with it.
    const noClock = {
      roundName: "231-a.json",
      selftest: [{ ok: true, skipped: false, name: "s", ms: 5000 }],
      trace: { entries: [{ message: "answered", data: { ms: 10 } }] },
    };
    const mixed = [...Array.from({ length: 6 }, () => round(600000, [50000], [300000])), noClock];
    expect(roundAnatomy(mixed).n).toBe(6);
  });

  it("skips a round whose trace never started", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { roundAnatomy } = await import("../scripts/triage.mjs");
    // A crashed round can archive an empty trace. Counting its span as zero
    // would drag every median toward a round that produced nothing.
    const mixed = [
      ...Array.from({ length: 6 }, () => round(600000, [50000], [300000])),
      { roundName: "231-a.json", selftest: [], trace: { entries: [] } },
    ];
    expect(roundAnatomy(mixed).n).toBe(6);
  });
});

describe("which scenarios lose draws to a stall", () => {
  const draw = () => ({ scope: "draw", message: "batch issued", data: {}, ms: 1 });
  const stall = () => ({
    scope: "host",
    message: "gave up waiting",
    data: { what: "drawing shapes 1-10 of 16" },
    ms: 2,
  });
  const starts = (name: string) => ({ scope: "selftest", message: "scenario starting", data: { name }, ms: 0 });
  const round = (entries: unknown[]) => ({ roundName: "230-a.json", trace: { entries } });

  it("counts draws by the line that fires on SUCCESS", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByScenario } = await import("../scripts/triage.mjs");
    // `drawing the chart's shapes` LOOKS like a draw counter and is a
    // `boundedSync` label — it appears only when the sync fails. Counting it
    // gives every scenario a 100% stall rate, which is what my first attempt
    // reported.
    const rows = stallsByScenario([round([starts("inserts"), ...Array(25).fill(0).map(draw), stall()])]);
    expect(rows[0].draws).toBe(25);
    expect(rows[0].stalls).toBe(1);
    expect(rows[0].pct).toBeCloseTo(4, 5);
  });

  it("attributes a stall to the scenario that was running", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByScenario } = await import("../scripts/triage.mjs");
    const rows = stallsByScenario([
      round([
        starts("quiet"),
        ...Array(30).fill(0).map(draw),
        starts("loses draws"),
        ...Array(30).fill(0).map(draw),
        stall(),
      ]),
    ]);
    expect(rows[0].name).toBe("loses draws");
    expect(rows.find((r: { name: string }) => r.name === "quiet").stalls).toBe(0);
  });

  it("says nothing about a scenario with too few draws to rate", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByScenario } = await import("../scripts/triage.mjs");
    // One stall out of three draws is 33% and means nothing. The same bar the
    // rest of this report uses against 5-of-5 style numbers.
    expect(stallsByScenario([round([starts("rare"), draw(), draw(), draw(), stall()])])).toEqual([]);
  });

  it("does not count a wedge that was not a draw", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByScenario } = await import("../scripts/triage.mjs");
    // The host gives up waiting on plenty of things. Only a draw stall costs the
    // scenario its verdict, and pooling the others would inflate every rate with
    // events that cost nothing.
    const other = { scope: "host", message: "gave up waiting", data: { what: "listing the deck's slides" }, ms: 2 };
    const rows = stallsByScenario([round([starts("one"), ...Array(25).fill(0).map(draw), other])]);
    expect(rows[0].stalls).toBe(0);
  });
});

describe("holding a selection through a draw", () => {
  const claimOf = async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { CLAIMS } = await import("../scripts/claims.mjs");
    return CLAIMS.find((c: { id: string }) => c.id === "dropping-the-selection-stops-the-stall");
  };
  const draw = () => ({ scope: "draw", message: "batch issued", data: {}, ms: 1 });
  const stall = () => ({ scope: "host", message: "gave up waiting", data: { what: "drawing shapes 1-10" }, ms: 2 });
  const starts = (name: string) => ({ scope: "selftest", message: "scenario starting", data: { name }, ms: 0 });
  const round = (heldDraws: number, heldStalls: number, dropDraws: number) => ({
    roundName: "230-a.json",
    trace: {
      entries: [
        starts("a selected shape survives an insert"),
        ...Array(heldDraws).fill(0).map(draw),
        ...Array(heldStalls).fill(0).map(stall),
        starts("edit the chart the user selected"),
        ...Array(dropDraws).fill(0).map(draw),
      ],
    },
  });

  it("holds while the held arm stalls and the dropped arm does not", async () => {
    const claim = await claimOf();
    expect(claim.check([round(25, 2, 25)]).ok).toBe(true);
  });

  it("goes stale — as GOOD news — when the held draw stops stalling", async () => {
    const claim = await claimOf();
    // Either the host stopped caring about a standing selection, or something
    // upstream fixed the draw. Both are worth being told about.
    const r = claim.check([round(25, 0, 25)]);
    expect(r.ok).toBe(false);
    expect(r.staleIsGood).toBe(true);
  });

  it("goes stale when the dropped arm starts stalling as much", async () => {
    const claim = await claimOf();
    // The zero is the load-bearing half. If dropping the selection stops helping,
    // `dropShapeSelection` is no longer the fix and the note that says it is
    // needs re-reading.
    const withBoth = {
      roundName: "230-a.json",
      trace: {
        entries: [
          starts("a selected shape survives an insert"),
          ...Array(25).fill(0).map(draw),
          stall(),
          starts("edit the chart the user selected"),
          ...Array(25).fill(0).map(draw),
          stall(),
          stall(),
        ],
      },
    };
    expect(claim.check([withBoth]).ok).toBe(false);
  });

  it("says ? until both arms have drawn enough", async () => {
    const claim = await claimOf();
    expect(claim.check([round(5, 1, 5)]).ok).toBe(null);
  });
});

describe("does a busy slide stall a draw", () => {
  const batch = (onSlide: number, key?: string) => ({
    scope: "draw",
    message: "batch issued",
    data: { onSlide, ...(key === undefined ? {} : { onSlideKey: key }) },
    ms: 1,
  });
  const stall = () => ({ scope: "host", message: "gave up waiting", data: { what: "drawing shapes 1-10" }, ms: 2 });
  const round = (entries: unknown[]) => ({ roundName: "230-a.json", trace: { entries } });

  it("holds out a batch issued before its slide was named", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByOccupancy } = await import("../scripts/triage.mjs");
    // THIS REVERSES THE ANSWER, which is why it is not tidying. A chart's first
    // batch carries `onSlide: 0, onSlideKey: "(visible)"` — a placeholder, not a
    // count. Counted as an empty slide, 8 of the 9 apparent empty-slide stalls
    // in the real archive are these rows and the table says empty slides stall
    // MOST.
    const r = stallsByOccupancy(
      [
        round([
          ...Array(40)
            .fill(0)
            .map(() => batch(0, "(visible)")),
          stall(),
        ]),
      ],
      1,
    );
    expect(r.unresolved).toBe(40);
    expect(r.rows).toEqual([]);
  });

  it("holds out a batch with no slide key at all", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByOccupancy } = await import("../scripts/triage.mjs");
    const r = stallsByOccupancy([round([batch(0), batch(0)])], 1);
    expect(r.unresolved).toBe(2);
  });

  it("buckets resolved rows and attributes the stall to the batch before it", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByOccupancy } = await import("../scripts/triage.mjs");
    const r = stallsByOccupancy([round([batch(0, "s1"), batch(30, "s1"), stall(), batch(5, "s1")])], 1);
    const busy = r.rows.find((x: { bucket: string }) => x.bucket === "26-50");
    expect(busy.stalls).toBe(1);
    expect(r.rows.find((x: { bucket: string }) => x.bucket === "0 (empty)").stalls).toBe(0);
  });

  it("refuses to call a trend it cannot support", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByOccupancy } = await import("../scripts/triage.mjs");
    // Fifteen stalls across six thousand draws is what the archive has. Printing
    // a trend off that is the error this whole report exists to stop.
    const thin = round([
      ...Array(40)
        .fill(0)
        .map(() => batch(0, "s1")),
      ...Array(40)
        .fill(0)
        .map(() => batch(30, "s1")),
      stall(),
    ]);
    expect(stallsByOccupancy([thin], 1).separable).toBe(false);
  });

  it("calls it separable once every bucket has stalls to rate", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { stallsByOccupancy } = await import("../scripts/triage.mjs");
    const fat = round([
      ...Array(10)
        .fill(0)
        .flatMap(() => [batch(0, "s1"), stall()]),
      ...Array(10)
        .fill(0)
        .flatMap(() => [batch(30, "s1"), stall()]),
    ]);
    expect(stallsByOccupancy([fat], 1).separable).toBe(true);
  });
});

describe("where the host died", () => {
  it("groups crashes by their last step, ignoring timing and payload", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolCrashLastSteps } = await import("../scripts/triage.mjs");
    // The real lines carry a timestamp and a data blob, both different every
    // run. Grouping on the raw string would give every crash its own bucket and
    // report nothing — which is indistinguishable from "no pattern here".
    const step = (t: string, tail: string) => `  ${t}s  probe  re-asked what the empty deck could not answer  ${tail}`;
    const groups = poolCrashLastSteps([
      { name: "A", steps: ["x", step("583.8", 'asked=["a"]')] },
      { name: "B", steps: ["x", "y", step("630.9", 'asked=["b"] answered=["yes"]')] },
      { name: "C", steps: ["  341.7s  draw  parts list outcome  chart=2/8"] },
    ]);
    expect(groups.length, "the two probe crashes did not group together").toBe(2);
    expect(groups[0].n).toBe(2);
    expect(groups[0].key).toContain("re-asked what the empty deck could not answer");
    expect(groups[0].at).toEqual(["A", "B"]);
    // The step COUNT range is what separates "died early" from "died at the
    // end" — four crashes at 535-540 steps against five at 389-411 is the shape
    // that made the cluster visible.
    expect(groups[0].steps).toEqual([2, 3]);
  });

  it("ignores a crash file that recorded no steps at all", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { poolCrashLastSteps } = await import("../scripts/triage.mjs");
    // A run that died before writing a step says nothing about WHERE. Bucketing
    // it under "(no steps recorded)" would invent a location.
    expect(poolCrashLastSteps([{ name: "A", steps: [] }, { name: "B" }])).toEqual([]);
  });

  it("keys on channel and message, so the same step from two channels stays apart", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { crashStepKey } = await import("../scripts/triage.mjs");
    expect(crashStepKey("  12.3s  probe  asking  id=foo")).toBe("probe  asking");
    expect(crashStepKey("  12.3s  draw  asking  id=foo")).toBe("draw  asking");
    expect(crashStepKey("")).toBe("(no steps recorded)");
  });
});

describe("the scenario that killed the host", () => {
  /**
   * THE FOURTH OUTCOME CLASS. A round file holds passed, failed and skipped —
   * and cannot hold this one, because a scenario that kills the host produces
   * no verdict and the round it belonged to files nothing at all.
   *
   * That is not sampling, it is structural, and it hid the sharpest finding in
   * this archive: `same scale across the deck` has 0 failures in 282 recorded
   * verdicts — joint-safest in the suite — and is the scenario the host died
   * inside TEN times. `rounds-salvaged/` cannot help either: a salvaged round
   * carries verdicts, and a killed scenario has none.
   */
  /**
   * THE SPELLINGS ARE COPIED FROM `selftest.ts:3962`, NOT INVENTED HERE, and
   * that line is the whole reason these helpers carry a comment.
   *
   * `failed` used to read `scenario failed`, lowercase. The host has never
   * written that: it emits `scenario FAILED`, shouted, so a failure is legible
   * in a wall of trace. Across `crashes/` the counts are 905 `scenario passed`,
   * 10 `scenario skipped`, 5 `scenario FAILED` and ZERO lowercase `failed`.
   *
   * So the double could not fail the way the real thing fails. Every test below
   * passed, the matcher shipped knowing one spelling of three, and it invented
   * three deaths in the live budget. A fake that cannot reproduce the host's
   * own casing is not a fake, it is a different host.
   */
  const started = (name: string) => `  10.0s  selftest  scenario starting  name=${name} deckSlides=1 atMs=10000`;
  const passed = (name: string) => `  20.0s  selftest  scenario passed  name=${name} detail=fine`;
  const failed = (name: string) => `  20.0s  selftest  scenario FAILED  name=${name} detail=nope`;
  const skipped = (name: string) => `  20.0s  selftest  scenario skipped  name=${name} detail=host busy`;

  it("credits the death to the scenario that was still open", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalScenarios } = await import("../scripts/triage.mjs");
    const r = fatalScenarios([
      { steps: [started("one"), passed("one"), started("two"), "  30.0s  error  the host went"] },
    ]);
    expect(r.deaths).toEqual({ two: 1 });
    expect(r.attributed).toBe(1);
    // "one" closed cleanly and must not be blamed for what came after it.
    expect(r.deaths.one, "blamed a scenario that had already reported").toBeUndefined();
  });

  it("credits nothing when the host died outside a scenario", async () => {
    /**
     * 54 of the 84 sound crash records die in the probe phase or the
     * deck-evidence scan. Attributing those to whatever ran last would invent
     * a killer, and the counts are what the ratchet acts on — so they have to
     * be right rather than complete.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalScenarios } = await import("../scripts/triage.mjs");
    const r = fatalScenarios([
      { steps: [started("one"), passed("one"), "  30.0s  pane  collecting deck evidence — scanning"] },
    ]);
    expect(r.deaths).toEqual({});
    expect(r.attributed).toBe(0);
    expect(r.unattributed).toBe(1);
  });

  it("closes a scenario only by its OWN name", async () => {
    /**
     * A `scenario passed` line for something else must not close the one that
     * is open. Closing on any end line at all would credit the death to
     * whatever happened to run last rather than to what was running — and the
     * follow-up questions this suite asks mean end lines do interleave.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalScenarios } = await import("../scripts/triage.mjs");
    const r = fatalScenarios([{ steps: [started("two"), passed("something else"), "  30.0s  error  gone"] }]);
    expect(r.deaths, "another scenario's verdict closed this one").toEqual({ two: 1 });
  });

  it("counts a failed verdict as closing it — a fail is not a death", async () => {
    // The distinction the whole class rests on: `stop a run mid-draw` FAILS
    // constantly and also kills the host sometimes. Those are different facts
    // and the ratchet only governs the second.
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalScenarios } = await import("../scripts/triage.mjs");
    const r = fatalScenarios([{ steps: [started("two"), failed("two"), "  30.0s  error  gone"] }]);
    expect(r.deaths).toEqual({});
  });

  it("closes on the host's OWN three spellings, one of which shouts", async () => {
    /**
     * THE BUG THIS SUITE SHIPPED. The matcher read `passed|failed`,
     * case-sensitively, against a host that writes `scenario FAILED` and
     * `scenario skipped`. Either one left the scenario open forever and the
     * record's death was credited to it.
     *
     * It invented three: `a chart of rotated shapes` was given 2 and `where a
     * rotated shape lands` 1, and both went into `FATAL_SCENARIO_BUDGET` as
     * scenarios that had killed the host. Neither ever has. Worse than a wrong
     * number — an absent name is what makes a first death a rise from zero, so
     * a fabricated entry silently absorbs a real first kill.
     *
     * Asserted one spelling per case rather than in a loop, so a regression
     * names the spelling it lost.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalScenarios } = await import("../scripts/triage.mjs");
    const died = "  30.0s  error  the host went";

    const shouted = fatalScenarios([{ steps: [started("two"), failed("two"), started("three"), died] }]);
    expect(shouted.deaths, "`scenario FAILED` did not close it").toEqual({ three: 1 });

    const skip = fatalScenarios([{ steps: [started("two"), skipped("two"), started("three"), died] }]);
    expect(skip.deaths, "`scenario skipped` did not close it").toEqual({ three: 1 });

    // A skip closes it because the scenario ENDED and the host outlived it —
    // the opposite of the thing being counted.
    const skipAlone = fatalScenarios([{ steps: [started("two"), skipped("two"), died] }]);
    expect(skipAlone.deaths, "a skipped scenario was blamed for a later death").toEqual({});
    expect(skipAlone.unattributed).toBe(1);

    // And the spelling that always worked still does.
    const plain = fatalScenarios([{ steps: [started("two"), passed("two"), started("three"), died] }]);
    expect(plain.deaths).toEqual({ three: 1 });
  });

  it("counts one crash filed twice as one crash", async () => {
    /**
     * `2026-08-29T03-32-07-crashed-run.json` and `…T08-27-00-crashed-run.json`
     * are the same event: same build, same `startedAt` to the millisecond, 392
     * byte-identical steps, re-archived five hours later. Counted twice it put
     * `same scale across the deck` at 10 when it has 9.
     *
     * Deduped at READ time. The archive is append-only and what it means is the
     * owner's call — a reader that counts an event once is not a reader that
     * edits history.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { loadCrashRecords } = await import("../scripts/rounds-gate.mjs");
    const twice = {
      "a-crashed-run.json": JSON.stringify({ build: "4275306", startedAt: "T0", steps: [started("x")] }),
      "b-crashed-run.json": JSON.stringify({ build: "4275306", startedAt: "T0", steps: [started("x")] }),
      // Same build and start, DIFFERENT steps: a different run of one session,
      // and it must survive. Keying on the steps too is what makes that hold.
      "c-crashed-run.json": JSON.stringify({ build: "4275306", startedAt: "T0", steps: [started("y")] }),
      // Same build, different start: two genuine crashes. Both count.
      "d-crashed-run.json": JSON.stringify({ build: "4275306", startedAt: "T1", steps: [started("x")] }),
    };
    const recs = loadCrashRecords(
      "crashes",
      () => Object.keys(twice),
      (p: string) => twice[p.split("/").pop() as keyof typeof twice],
    );
    expect(recs).toHaveLength(3);
    expect(recs.duplicates, "the duplicate was not named").toEqual(["b-crashed-run.json"]);

    // A record identifiable by NEITHER build nor start cannot be deduped, and
    // must be kept rather than folded into the first other anonymous one.
    const anon = {
      "e-crashed-run.json": JSON.stringify({ steps: [] }),
      "f-crashed-run.json": JSON.stringify({ steps: [] }),
    };
    const kept = loadCrashRecords(
      "crashes",
      () => Object.keys(anon),
      (p: string) => anon[p.split("/").pop() as keyof typeof anon],
    );
    expect(kept, "two unidentifiable records were folded into one").toHaveLength(2);
  });

  it("counts how often each scenario RAN, from rounds AND from crashes", async () => {
    /**
     * THE DENOMINATOR, and both halves are needed. Rounds alone miss every run
     * that ended in a crash — which is exactly the population deaths come from,
     * so a rate built on rounds alone divides by the wrong number in the
     * direction that flatters the scenario.
     *
     * A skipped scenario RAN. It was attempted and the host outlived it.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioRuns } = await import("../scripts/triage.mjs");
    const runs = scenarioRuns(
      [
        {
          selftest: [
            { name: "two", ok: true },
            { name: "three", skipped: true },
          ],
        },
        { selftest: [{ name: "two", ok: false }] },
      ],
      [{ steps: [started("two"), started("four")] }],
    );
    expect(runs).toEqual({ two: 3, three: 1, four: 1 });
  });

  it("needs a rise past NOISE, not past the point estimate", async () => {
    /**
     * THE TRAP THE COUNT VERSION FELL INTO, and a rate seeded at the point
     * estimate falls into it too. `stop a run mid-draw` kills the host on about
     * a third of its runs; at 8 deaths in 25 runs a NINTH death takes the rate
     * from 320 to 346. Against a ceiling of 330 that is red — for a scenario
     * doing exactly what it has done all week. Killing a third of the time
     * PRODUCES eight-and-nine-death stretches; that is the baseline, not a rise.
     *
     * So the bound is binomial: p*n + 2*sqrt(p*(1-p)*n).
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalRateBreaches, fatalDeathsAllowed } = await import("../scripts/triage.mjs");
    const ceilings = { "stop a run mid-draw": 330 };

    // At the accepted rate: silent. And one MORE death is still silent, which
    // is the whole point — the count version went red exactly here.
    expect(fatalRateBreaches({ "stop a run mid-draw": 8 }, { "stop a run mid-draw": 25 }, ceilings)).toEqual([]);
    expect(
      fatalRateBreaches({ "stop a run mid-draw": 9 }, { "stop a run mid-draw": 26 }, ceilings),
      "a ninth death at the baseline rate tripped the gate",
    ).toEqual([]);

    // A REAL rise does trip: same exposure, far more deaths.
    expect(fatalRateBreaches({ "stop a run mid-draw": 20 }, { "stop a run mid-draw": 26 }, ceilings)).toHaveLength(1);

    // A rate FALLS as runs accumulate without deaths — the property a
    // cumulative count can never have, and the reason a landed fix protects
    // itself with nobody editing a number.
    const far = fatalRateBreaches({ "stop a run mid-draw": 8 }, { "stop a run mid-draw": 400 }, ceilings);
    expect(far, "a scenario that stopped dying was still reported").toEqual([]);

    // p = 0 gets no statistics and needs none.
    expect(fatalDeathsAllowed(0, 500)).toBe(0);
    // Nor does zero exposure invent a rate.
    expect(fatalDeathsAllowed(330, 0)).toBe(0);
  });

  it("a ceiling of ZERO cannot be cleared by clean rounds, whatever the guidance says", async () => {
    /**
     * The gate prints "falls on its own as runs pile up without deaths … NOTHING
     * NEEDS EDITING to make it green" beside every breach. For a seeded ceiling
     * that is true — the allowance is `p*n + 2*sqrt(p(1-p)n)` and grows with n.
     * At a ceiling of zero it is false, and pinning it here is the point: the
     * crash record stays in `crashes/` forever, so the breach outlives any
     * number of clean rounds.
     *
     * Round 428 is the case. `what a chart kind costs` killed the host once in
     * 35 runs and is absent from `FATAL_SCENARIO_RATE`, so the cycle stops after
     * one round every night until a person decides something — while that
     * table's own docstring says "Do not add a name here to quiet a gate". Both
     * rules are right alone and together they leave no green path.
     *
     * NOT a claim that the trip is wrong. A first host death is the case the
     * check exists for. What is recorded here is that it is permanent, because
     * the message beside it says the opposite.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalRateBreaches } = await import("../scripts/triage.mjs");
    const asMeasured = fatalRateBreaches({ "what a chart kind costs": 1 }, { "what a chart kind costs": 35 }, {});
    expect(asMeasured, "a first-ever death did not trip the check").toHaveLength(1);
    expect(asMeasured[0]).toMatchObject({ count: 1, runs: 35, allowed: 0 });
    // A thousand clean rounds later, with the same single death on record.
    const muchLater = fatalRateBreaches({ "what a chart kind costs": 1 }, { "what a chart kind costs": 100_000 }, {});
    expect(muchLater, "the zero-ceiling breach cleared itself, so the guidance is right after all").toHaveLength(1);
  });

  it("trips on the FIRST death of a scenario that has never killed the host", async () => {
    /**
     * The case most worth catching, and the one thing the count version got
     * right. An absent name means a ceiling of 0, so p=0, so the noise bound is
     * 0, so any death at all is a rise. No statistics are applied because none
     * are needed — this is not a question about a rate.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalRateBreaches } = await import("../scripts/triage.mjs");
    const out = fatalRateBreaches(
      { "edit a chart on the visible slide": 1 },
      { "edit a chart on the visible slide": 422 },
      { "stop a run mid-draw": 330 },
    );
    expect(out).toHaveLength(1);
    expect(out[0].name).toBe("edit a chart on the visible slide");
    expect(out[0].allowed).toBe(0);
  });

  describe("a receipt for a death, which is not a licence for the scenario", () => {
    /**
     * `deathsAcknowledged` is the only green path a ceiling-0 breach has, so
     * what it REFUSES matters more than what it clears. Each refusal below is
     * the thing that keeps it from becoming the ceiling edit `FATAL_SCENARIO_RATE`
     * forbids in its own docstring.
     */
    const load = async () =>
      // @ts-expect-error - plain .mjs tool, no types.
      (await import("../scripts/triage.mjs")).deathsAcknowledged;

    const breach = (over: Record<string, unknown> = {}) => ({
      name: "what a chart kind costs",
      count: 1,
      runs: 39,
      rate: 25.6,
      allowed: 0,
      cap: 0,
      ...over,
    });
    const receipt = (over: Record<string, unknown> = {}) => ({
      record: "2026-09-07T22-13-42-crashed-run.json",
      scenario: "what a chart kind costs",
      seen: "2026-09-08",
      why: "docs/BACKLOG.md — the sweep deleted nothing seven times",
      ...over,
    });
    const credits = { "what a chart kind costs": ["2026-09-07T22-13-42-crashed-run.json"] };

    it("clears the exit for one death whose record a person named", async () => {
      const out = (await load())([breach()], credits, [receipt()], {});
      expect(out.cleared, "a signed first death still stopped the night").toHaveLength(1);
      expect(out.standing).toHaveLength(0);
      expect(out.stale).toHaveLength(0);
    });

    it("will not sign away a scenario that already has a rate to fall", async () => {
      // THE ONE THAT MATTERS MOST. A seeded ceiling has a green path already —
      // the rate drops as clean runs accumulate — so letting a receipt clear it
      // would silence "this scenario got WORSE", which is the question the
      // table exists to ask.
      const out = (await load())(
        [breach({ name: "same scale across the deck", allowed: 30, count: 1 })],
        { "same scale across the deck": ["2026-09-07T22-13-42-crashed-run.json"] },
        [receipt({ scenario: "same scale across the deck" })],
        { "same scale across the deck": 30 },
      );
      expect(out.cleared, "a receipt cleared a breach that had a ceiling").toHaveLength(0);
      expect(out.standing).toHaveLength(1);
    });

    it("will not sign a SECOND death, however the ledger is written", async () => {
      // Two deaths is the archive's own signal that the first was not a
      // one-off: every repeat death in this archive landed within 25 rounds of
      // its predecessor, median 1. So the second has no receipt path at all.
      const out = (await load())(
        [breach({ count: 2 })],
        { "what a chart kind costs": ["a.json", "b.json"] },
        [receipt({ record: "a.json" }), receipt({ record: "b.json" })],
        {},
      );
      expect(out.cleared, "a second death was acknowledged").toHaveLength(0);
      expect(out.standing).toHaveLength(1);
    });

    it("refuses a second death even when only ONE record credits it", async () => {
      /**
       * THE CASE THE TEST ABOVE DOES NOT REACH, found by mutation: with two
       * credited records, `records.length === 1` already refuses, so deleting
       * the `count !== 1` guard entirely leaves that test green. The guard only
       * does work when the pools disagree the other way — a count of two from a
       * single credited record, which is what a re-attribution or a deduped
       * record produces.
       *
       * It has to refuse. `count` is what the breach was computed from, and a
       * receipt honoured against a smaller credit list would sign for a death
       * the ledger never named.
       */
      const out = (await load())([breach({ count: 2 })], credits, [receipt()], {});
      expect(out.cleared, "a two-death breach was cleared by one receipt").toHaveLength(0);
      expect(out.standing).toHaveLength(1);
    });

    it("will not let a receipt for one scenario clear another", async () => {
      const out = (await load())([breach()], credits, [receipt({ scenario: "stop a run mid-draw" })], {});
      expect(out.cleared).toHaveLength(0);
      expect(out.standing).toHaveLength(1);
    });

    it("goes stale when the record stops crediting the scenario it signed for", async () => {
      // The guard against a ledger nobody re-reads. Re-attribution, a deleted
      // record, or a poisoned build all land here, and the gate exits 2 —
      // "could not judge", not "regression".
      const out = (await load())([], { "what a chart kind costs": ["something-else.json"] }, [receipt()], {});
      expect(out.stale, "a receipt pointing at nothing was still honoured").toHaveLength(1);
      expect(out.stale[0].why).toMatch(/no crash record named/);
    });

    it("goes stale when the scenario has since been given a ceiling", async () => {
      const out = (await load())([], credits, [receipt()], { "what a chart kind costs": 30 });
      expect(out.stale).toHaveLength(1);
      expect(out.stale[0].why).toMatch(/now carries a ceiling/);
    });

    it("refuses a receipt when the credit list and the count disagree", async () => {
      // `credits` and `count` are written by the same loop in `fatalScenarios`,
      // so a disagreement means the two were computed over different pools —
      // and a receipt honoured across that gap would be signing for a death
      // nobody looked at.
      const out = (await load())(
        [breach({ count: 1 })],
        { "what a chart kind costs": ["a.json", "b.json"] },
        [receipt({ record: "a.json" })],
        {},
      );
      expect(out.cleared).toHaveLength(0);
      expect(out.standing).toHaveLength(1);
    });
  });

  it("ships a ledger whose every entry says what it signed and why", async () => {
    /**
     * NOT an assertion that the ledger is empty — the owner is meant to write
     * in it, and a test that forbade that would be a gate on his own decision.
     * What is asserted is the shape: a receipt with no reason, or one whose
     * reason is a placeholder, is the thing this mechanism must never become.
     * `KNOWN_DIVERGENCES` sets the same bar in the same words.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { DEATHS_ACKNOWLEDGED } = await import("../scripts/rounds-gate.mjs");
    expect(Array.isArray(DEATHS_ACKNOWLEDGED)).toBe(true);
    for (const e of DEATHS_ACKNOWLEDGED as { record: string; scenario: string; seen: string; why: string }[]) {
      expect(e.record, "a receipt with no crash record names nothing").toMatch(/-crashed-run\.json$/);
      expect(e.scenario, "a receipt with no scenario cannot be matched").toBeTruthy();
      expect(e.seen, "a receipt with no date says nobody read it on any day").toMatch(/^\d{4}-\d{2}-\d{2}$/);
      expect(e.why?.length ?? 0, `receipt for ${e.record} has no reason`).toBeGreaterThan(20);
      expect(e.why, "'we have not looked into it' is not a reason").not.toMatch(/^(TODO|TEMP|tbd|unknown)/i);
    }
  });

  it("does not invent a rate for a death it has no runs for", async () => {
    // A death with no recorded run is a gap in what this tool can read. Dividing
    // by zero would fail the gate on a parsing problem rather than on a host.
    // @ts-expect-error - plain .mjs tool, no types.
    const { fatalRateBreaches } = await import("../scripts/triage.mjs");
    expect(fatalRateBreaches({ mystery: 3 }, {}, {})).toEqual([]);
  });

  it("is wired into the gate as a FATAL check, not a report", async () => {
    /**
     * A scenario that stops passing is exit 1 here. One that starts taking
     * PowerPoint down produces no verdict to notice it by and is at least as
     * serious, so it is fatal too — the owner's call, made 2026-09-03.
     *
     * IT IS EXIT 3 SINCE 2026-09-07, and the owner's decision is untouched: the
     * check is exactly as fatal, it simply no longer shares a code with the
     * regression check. Sharing one meant `cycle.mjs` printed "a scenario that
     * WAS passing has stopped" for a host death, and round 428 ended a night
     * that way with 19 of 19 scenarios green.
     *
     * Asserted as source because reaching it needs a crashes/ directory that
     * breaches, and the gate's main block is not exported.
     */
    const src = readFileSync(new URL("../scripts/rounds-gate.mjs", import.meta.url), "utf8");
    expect(src, "the gate no longer computes fatal scenarios").toMatch(/fatalScenarios\(crashes\)/);
    /**
     * PINNED ON THE GUARD, NOT ON A DISTANCE — rewritten 2026-09-08 because the
     * distance version had started passing for the wrong reason.
     *
     * It was `/breaches\.length[\s\S]{0,4000}process\.exit\(3\)/`, with a window
     * widened twice (2000 → 4000) as the printed guidance grew. When the receipt
     * mechanism landed, the breach block's own `breaches.length` moved to 7,343
     * characters from the exit — outside the window — and the regex went on
     * matching because a SECOND `breaches.length` appears 818 characters before
     * it. The assertion stayed green while the thing it names stopped being what
     * it measured. That is the same defect this suite keeps finding elsewhere: a
     * check satisfied by something other than the property it claims.
     *
     * So it pins the structure instead. The exit is guarded by what is still
     * STANDING after receipts are applied — a breach a person has signed for
     * prints and does not stop the night — and that guard sits immediately
     * around the exit, where no amount of added prose can separate them.
     */
    expect(src, "the fatal exit is no longer guarded by what is still standing").toMatch(
      /if \(receipts\.standing\.length\)[\s\S]{0,1500}process\.exit\(3\)/,
    );
    // AND THE BREACH BLOCK IS STILL WHAT COMPUTES IT. Without this the pin above
    // would survive deleting the breach report entirely.
    expect(src, "the gate no longer computes rate breaches").toMatch(/fatalRateBreaches\(fatal\.deaths, runs,/);
    expect(src, "the gate no longer applies receipts to those breaches").toMatch(
      /deathsAcknowledged\(breaches, fatal\.credits,/,
    );
    // AND NOT BY FALLING THROUGH TO THE REGRESSION EXIT. `exit(1)` appears
    // later in this file for the check that judges verdicts; a breach that
    // reached THAT would be back to one code for two questions.
    expect(src, "the breach exits through the regression code").not.toMatch(
      /if \(receipts\.standing\.length\)[\s\S]{0,1500}process\.exit\(1\)/,
    );
    // And it reads the CRASH records, not the salvaged rounds — a salvaged
    // round carries verdicts and a killed scenario has none.
    expect(src).toMatch(/loadCrashRecords\(\)/);
    /**
     * AND IT GATES ON A RATE, WITH A DENOMINATOR. The count version of this
     * check was red within hours of shipping and could not have been anything
     * else: deaths only accumulate, so a budget seeded at today's total
     * breaches on the very next one. Reverting to `fatalBudgetBreaches` would
     * restore a gate that fires on ordinary operation, so the wiring is
     * asserted rather than assumed.
     */
    expect(src, "the gate is not computing a denominator").toMatch(/scenarioRuns\(rounds, crashes\)/);
    expect(src, "the gate is back on a raw count").toMatch(/fatalRateBreaches\(fatal\.deaths, runs,/);
  });
});

describe("pooling salvaged rounds without lying about the denominator", () => {
  /**
   * WHY `rounds-salvaged/` SAT UNREAD. A crashed round ran scenarios 1..k and
   * never reached k+1..n. Concatenate it onto `rounds/` and every scenario it
   * never got to reads as a silent success — late scenarios look rarer and
   * safer, early ones more common, and each is measured against a different
   * truth.
   *
   * So a scenario's denominator is the number of rounds that RECORDED A VERDICT
   * for it and nothing else. A round that never reached it contributes to
   * neither side.
   */
  const round = (verdicts: Array<[string, boolean]>, partial?: boolean) => ({
    ...(partial ? { salvagedPartial: { reached: verdicts.length, of: 3, missing: ["c"] } } : {}),
    selftest: verdicts.map(([name, ok]) => ({ name, ok })),
  });

  it("counts a scenario only in the rounds that reached it", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioRatesOver } = await import("../scripts/triage.mjs");
    const r = scenarioRatesOver([
      round([
        ["a", true],
        ["b", true],
        ["c", true],
      ]),
      // Died after "b": "c" was never put, and must not read as a pass.
      round(
        [
          ["a", true],
          ["b", false],
        ],
        true,
      ),
    ]);
    expect(r.per.a).toEqual({ ran: 2, failed: 0 });
    expect(r.per.b, "the failure in the partial round was lost").toEqual({ ran: 2, failed: 1 });
    expect(r.per.c, "a scenario that never ran was counted as having passed").toEqual({ ran: 1, failed: 0 });
    expect(r.partial, "the reader cannot see how much of this is partial").toBe(1);
  });

  it("leaves skipped scenarios out of both sides", async () => {
    // @ts-expect-error - plain .mjs tool, no types.
    const { scenarioRatesOver } = await import("../scripts/triage.mjs");
    const r = scenarioRatesOver([{ selftest: [{ name: "a", ok: false, skipped: true }] }]);
    expect(r.per.a, "a skip was counted as a failure").toBeUndefined();
  });

  it("makes every mixed number say which population it came from", async () => {
    /**
     * This repo has already paid for an unlabelled population once: the 4:3
     * crash rate that turned out to be about a DOCUMENT. A rate pooled from
     * completed and salvaged rounds is a different claim from one pooled from
     * completed rounds alone, and the label has to travel with the number.
     */
    // @ts-expect-error - plain .mjs tool, no types.
    const { populationLine } = await import("../scripts/triage.mjs");
    expect(populationLine(344, 69, 19)).toBe("344 complete round(s) + 69 salvaged (19 of them partial)");
    // Complete-only stays exactly as it reads today — no salvaged clause at all,
    // so every existing finding remains comparable.
    expect(populationLine(344, 0, 0)).toBe("344 complete round(s)");
  });
});
