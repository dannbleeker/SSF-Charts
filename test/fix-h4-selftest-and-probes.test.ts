// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import {
  installHost,
  makeSlide,
  makeShape,
  faults,
  unansweredShapeReads,
  type FakeShape,
  type FakeSlide,
} from "./helpers/office-host";
import { EXPERIMENTS, runExperiment } from "../src/render/experiments";
import { runSelfTest, setSelfTestRasterizer, type ScenarioResult } from "../src/taskpane/selftest";
import { runHostProbes } from "../src/render/host-probe";
import { setTracing, traceLog } from "../src/core/trace";
import { rawSlideCountUses } from "./helpers/slide-count-sites";

/**
 * The instruments that judge the host: the on-demand experiments, the self-test
 * battery's verdicts, and the probe sheet's slide handling.
 *
 * Every defect in this file is the same shape wearing five costumes: an
 * instrument that reports on the host and cannot be trusted to say so when it
 * did not measure, or that does to the document what it is only meant to ask
 * about. The tests therefore run the REAL scenarios and experiments on the fake
 * host and bend only what a test must be able to bend, through the wrapper
 * below — a canned answer for a call the fake cannot be made to get wrong the
 * way the real host does (ids that repeat across slides, a slide count that
 * comes back with no value, a draw that reports success and lands nowhere).
 */
type EditTargetLike = { slideId: string; shapeId: string; lost?: string; [k: string]: unknown };
type AnyFn = (...args: unknown[]) => unknown;

const hold = vi.hoisted(() => ({
  /** The fake's deck, by reference: `installHost` mutates the array it is given. */
  deck: [] as { delete(): void; created: unknown[] }[],

  /** r2-31 — set once the explode scenario's collapse to a picture has run. */
  sawPicture: false,
  /** r2-31 — what to do to the explode-back call, when this test is about it. */
  explode: null as null | {
    /** Refuse tag writes during the explode-back, so the chart really loses its config. */
    refuseTags?: boolean;
    /** Give another slide's chart the exploded chart's shape id — ids repeat across slides. */
    collide?: boolean;
    /** Replace the update's own report of what it did. */
    report?: (out: EditTargetLike) => EditTargetLike;
  },

  /** r2-32 — what the subject draw of `a big chart on a slide of its own` does. */
  draw: "real" as "real" | "nowhere" | "vanishes",

  /** r2-62 — answer every slide count with no value once a scenario's own slide exists. */
  blindAfterAdd: false,
  blind: false,

  /** r2-30 — a host that refuses a child by id and still honours the group. */
  liftRefusalForDeletes: false,
  /** r2-30 — every delete an experiment asked for. */
  deletes: [] as { slideId: string; ids: string[] }[],
}));

vi.mock("../src/render/powerpoint", async (importOriginal) => {
  const real = await importOriginal<typeof import("../src/render/powerpoint")>();
  return {
    ...real,
    updateChartInSlide: async (scene: unknown, target: EditTargetLike, opts: { pictureBase64?: string } = {}) => {
      const picture = !!opts.pictureBase64;
      // The explode-back is the first update after the collapse to a picture.
      const back = !picture && hold.sawPicture;
      if (picture) hold.sawPicture = true;
      const arm = hold.explode;
      if (back && arm?.refuseTags) faults.refuseTagWrites = 100000;
      let out = (await (real.updateChartInSlide as AnyFn)(scene, target, opts)) as EditTargetLike | null;
      if (back) {
        hold.sawPicture = false;
        faults.refuseTagWrites = 0;
        if (out && arm?.collide) {
          const scan = await real.listChartsInDeck();
          const other = scan.charts.find((c) => c.target.slideId !== out!.slideId);
          for (const slide of hold.deck)
            for (const shape of slide.created as { id: string }[])
              if (other && shape.id === other.target.shapeId) shape.id = out.shapeId;
        }
        if (out && arm?.report) out = arm.report(out);
      }
      return out;
    },
    insertSceneIntoSlide: async (scene: unknown, opts: { slideId?: string } = {}, onPhase?: AnyFn) => {
      // Only the SUBJECT draw names a slide; the scenario's control draws on the visible one.
      if (opts.slideId && hold.draw === "nowhere") {
        for (let n = 0; n < 11; n++) onPhase?.("commit", `${n} of 11`);
        return { slideId: opts.slideId, shapeId: "ghost", left: 0, top: 0, width: 1, height: 1 };
      }
      if (opts.slideId && hold.draw === "vanishes") {
        // The host drops the slide it just added, and the insert's own recovery
        // draws on the visible one — round 373.
        hold.deck[hold.deck.length - 1].delete();
        return (real.insertSceneIntoSlide as AnyFn)(scene, { ...opts, slideId: undefined }, onPhase);
      }
      return (real.insertSceneIntoSlide as AnyFn)(scene, opts, onPhase);
    },
    slideCount: async () => (hold.blind ? undefined : real.slideCount()),
    addSlideForChart: async () => {
      const id = await real.addSlideForChart();
      if (hold.blindAfterAdd) hold.blind = true;
      return id;
    },
    chartOnItsOwnSlideAsFile: async (...args: unknown[]) => {
      const landed = await (real.chartOnItsOwnSlideAsFile as AnyFn)(...args);
      if (hold.blindAfterAdd) hold.blind = true;
      return landed;
    },
    deleteShapesById: async (slideId: string, ids: string[]) => {
      hold.deletes.push({ slideId, ids: [...ids] });
      const refusing = faults.refuseShapeById;
      if (hold.liftRefusalForDeletes) faults.refuseShapeById = false;
      try {
        return await real.deleteShapesById(slideId, ids);
      } finally {
        faults.refuseShapeById = refusing;
      }
    },
  };
});

afterEach(() => {
  hold.sawPicture = false;
  hold.explode = null;
  hold.draw = "real";
  hold.blindAfterAdd = false;
  hold.blind = false;
  hold.liftRefusalForDeletes = false;
  hold.deletes = [];
  hold.deck = [];
  faults.refuseTagWrites = 0;
  faults.refuseShapeById = false;
  faults.refuseSlideDelete = false;
  faults.newSlideIdSettlesAfter = null;
  setTracing(false);
});

const byName = (results: ScenarioResult[], name: string): ScenarioResult => {
  const found = results.find((r) => r.name === name);
  if (!found) throw new Error(`no verdict for "${name}" in: ${results.map((r) => r.name).join(" | ")}`);
  return found;
};

/* ───────────────────────────────── r2-30 ───────────────────────────────── */

/**
 * Two experiments grouped EVERY shape on the deck's first slide, painted the
 * first one red, and deleted the group. The first slide is the user's, not a
 * scratch slide: the comment that said "the scratch slide is empty until this
 * experiment fills it" stopped being true when they moved to slide 0, and the
 * three siblings written after that had the repair (a set of what the slide
 * already held) and these two never did.
 */
describe("experiments leave the user's own slide as they found it", () => {
  /** A deck whose first slide holds a logo and a title, as a real deck's does. */
  function usersDeck() {
    const first = makeSlide("users-title-slide");
    const logo = makeShape("geometric", "rectangle", { left: 500, top: 10, width: 80, height: 40 });
    logo.name = "Company logo";
    logo.fillColor = "0055AA";
    const title = makeShape("text", undefined, { left: 40, top: 200, width: 400, height: 60 });
    title.name = "Title 1";
    title.text = "Q3 results";
    first.created.push(logo, title);
    const deck = [first, makeSlide("users-second-slide")];
    hold.deck = deck;
    installHost(deck);
    return { first, logo, title };
  }

  const liveOn = (slide: FakeSlide): FakeShape[] => (slide.created as FakeShape[]).filter((s) => !s.deleted);
  /** What is on the slide, by name — a failing diff of whole fake shapes is unreadable. */
  const heldOn = (slide: FakeSlide): string[] => liveOn(slide).map((s) => s.name ?? s.id);
  const USERS = ["Company logo", "Title 1"];
  /** Whether any live group on the slide holds this shape — directly or through another group. */
  const swallowed = (slide: FakeSlide, shape: FakeShape): boolean => {
    const holds = (g: FakeShape): boolean =>
      ((g.grouped ?? []) as FakeShape[]).some((c) => c.id === shape.id || (c.grouped ? holds(c) : false));
    return liveOn(slide).some((g) => !!g.grouped && holds(g));
  };

  const HOSTS: Record<string, (first: FakeSlide) => void> = {
    "an honest host": () => {},
    "a host that refuses shapes by id": () => {
      faults.refuseShapeById = true;
    },
    "a host whose new slide ids settle underneath it": () => {
      faults.newSlideIdSettlesAfter = 2;
    },
    // The slide would not list its shapes the first time it was asked. That is
    // not an empty slide, and an experiment that read it as one took every shape
    // on it for its own.
    "a host that will not list the slide's shapes at first": (first) => {
      unansweredShapeReads.add(first.id);
    },
  };

  for (const [hostName, arm] of Object.entries(HOSTS)) {
    for (const e of EXPERIMENTS) {
      it(`${e.id} does not touch the user's shapes on ${hostName}`, async () => {
        const { first, logo, title } = usersDeck();
        arm(first);
        await runExperiment(e.id);
        expect(logo.deleted, "the user's logo was deleted").toBe(false);
        expect(title.deleted, "the user's title was deleted").toBe(false);
        expect(logo.fillColor, "the user's logo was recoloured").toBe("0055AA");
        expect(title.text, "the user's title text changed").toBe("Q3 results");
        expect(swallowed(first, logo), "the user's logo was grouped with the experiment's shapes").toBe(false);
        expect(swallowed(first, title), "the user's title was grouped with the experiment's shapes").toBe(false);
      });
    }
  }

  it("draws nothing at all when the slide would not say what it already holds", async () => {
    // Asked of every experiment that puts shapes on the slide: with no list of
    // what was there first, nothing can be told apart from the user's own, so
    // the honest move is to not start — and not to leave rectangles behind.
    for (const e of EXPERIMENTS.filter((x) => x.id !== "slide-tag-round-trip")) {
      const { first } = usersDeck();
      unansweredShapeReads.add(first.id);
      const r = await runExperiment(e.id);
      expect(r.answer, `${e.id} asked its question without knowing what the slide held`).toBe("no-child-ids");
      expect(heldOn(first), `${e.id} drew on a slide it could not read`).toEqual(USERS);
    }
  });

  for (const id of ["grouped-child-by-id", "group-members-real-slide"]) {
    it(`${id} gives back everything it drew, on an honest host`, async () => {
      const { first } = usersDeck();
      const r = await runExperiment(id);
      expect(r.answer, `${id} never reached its question: ${r.detail}`).not.toMatch(/^(no-child-ids|no-group|threw)$/);
      expect(heldOn(first), `${id} left its own shapes on the user's first slide`).toEqual(USERS);
    });
  }

  it("grouped-child-by-id takes its group back off the slide on every answer, not only yes", async () => {
    // The child lookup is refused by id and the group itself is not: the answer is
    // `refused-by-id`, which used to return with the group — and whatever was
    // inside it — still on slide 1. Real-host record: `child 2 of group 5`, then
    // `child 5 of group 9`, then `child 9 of group 13`, each run picking up the
    // group the one before had left.
    const { first } = usersDeck();
    faults.refuseShapeById = true;
    hold.liftRefusalForDeletes = true;
    const r = await runExperiment("grouped-child-by-id");
    expect(r.answer, "this host should refuse the child by id").toBe("refused-by-id");
    expect(hold.deletes.length, "the experiment never tried to take its group back").toBeGreaterThan(0);
    expect(heldOn(first), `the ${r.answer} path left the group on the user's slide`).toEqual(USERS);
  });

  it("does not fold a group an earlier run left behind into the next run's group", async () => {
    // The host refuses the delete too, so run 1 leaves its group. Run 2 must see
    // that group as something the slide already held, not as one of its own three.
    const { first, logo, title } = usersDeck();
    faults.refuseShapeById = true;
    await runExperiment("grouped-child-by-id");
    await runExperiment("grouped-child-by-id");
    const nested = liveOn(first).filter((g) => ((g.grouped ?? []) as FakeShape[]).some((c) => !!c.grouped));
    expect(
      nested.map((g) => g.id),
      "a leftover group was grouped into the next run's group",
    ).toEqual([]);
    expect(swallowed(first, logo)).toBe(false);
    expect(swallowed(first, title)).toBe(false);
  });
});

/* ───────────────────────────────── r2-31 ───────────────────────────────── */

/**
 * `explode a degraded picture` decided that the exploded chart kept its config
 * with `rescan.found.find((c) => c.target.shapeId === exploded.shapeId)`: shape
 * id alone. Shape ids are per-slide integers on the real host (round 497's
 * inventory has `26` on five slides), the update had already said `no-config`,
 * and round 026-162f80a filed "config intact" over a chart that lost its tag
 * while the deck scan dropped from six charts to five.
 */
describe("explode a degraded picture", () => {
  const NAME = "explode a degraded picture";
  async function explode(arm: typeof hold.explode): Promise<ScenarioResult> {
    hold.explode = arm;
    const deck = [makeSlide("s1")];
    hold.deck = deck;
    installHost(deck);
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    return byName(await runSelfTest("probe", NAME), NAME);
  }

  it("passes when the chart really came back with its config", async () => {
    const r = await explode(null);
    expect(r.skipped, r.detail).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail).toContain("config intact");
  });

  it("fails when the config was lost and nothing else could be mistaken for the chart", async () => {
    const r = await explode({ refuseTags: true });
    expect(r.ok, "called a chart with no config re-editable").toBe(false);
    expect(r.detail).toMatch(/did not survive/);
  });

  it("fails when the config was lost and another slide's chart has the same shape id", async () => {
    // Round 026-162f80a: the tag write failed, the update said `no-config`, and
    // the id the picture had on its old slide is also an id on another.
    const r = await explode({ refuseTags: true, collide: true });
    expect(r.ok, "another slide's chart vouched for the exploded one").toBe(false);
    expect(r.detail).toMatch(/did not survive/);
  });

  it("believes the update when it says the config was lost, whatever the rescan finds", async () => {
    // Nothing is refused and the chart is on the slide with its tag — so the
    // rescan finds it by slide AND shape. The update's own report is the
    // first-hand account and it says `no-config`.
    const r = await explode({ report: (out) => ({ ...out, lost: "no-config" }) });
    expect(r.ok, "overruled the update's own `no-config`").toBe(false);
    expect(r.detail).toMatch(/did not survive/);
    expect(r.detail).toContain("no-config");
  });

  it("does not take a chart on ANOTHER slide for the exploded one just because the shape id repeats", async () => {
    // The update reports nothing wrong (so `lost` cannot save it) and the chart
    // genuinely has no config; only the slide tells the two apart.
    const r = await explode({
      refuseTags: true,
      collide: true,
      report: (out) => {
        const { lost: _lost, ...rest } = out;
        return rest as EditTargetLike;
      },
    });
    expect(r.ok, "matched a chart on another slide by shape id alone").toBe(false);
    expect(r.detail).toMatch(/did not survive/);
  });
});

/* ───────────────────────────────── r2-32 ───────────────────────────────── */

/**
 * `a big chart on a slide of its own` read what landed with the id
 * `addSlideForChart` returned — the add-time form, which stops resolving once the
 * host re-keys the slide. By the time eleven batches had drawn (~100s) the read
 * was always null, `landed` was undefined, and the one content check
 * (`typeof landed === "number" && landed === 0`) could not fire: 121 of 121
 * archived passes print "undefined shapes". Round 373 is the case it exists for:
 * the slide was gone, the insert drew all 103 shapes on the VISIBLE slide, and
 * the verdict still read "slide swept".
 */
describe("a big chart on a slide of its own", () => {
  const NAME = "a big chart on a slide of its own";
  async function bigChart(): Promise<ScenarioResult> {
    const deck = [makeSlide("s1")];
    hold.deck = deck;
    installHost(deck);
    return byName(await runSelfTest("probe", NAME), NAME);
  }

  it("counts the shapes that landed, once the host has re-keyed the slide", async () => {
    faults.newSlideIdSettlesAfter = 1;
    const r = await bigChart();
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail, "reported a count it never took").not.toMatch(/undefined/);
    expect(r.detail).toMatch(/, \d+ shapes, /);
  });

  it("still counts them on a host that does not re-key", async () => {
    const r = await bigChart();
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail).toMatch(/, \d+ shapes, /);
  });

  it("fails a draw that reports success and leaves the slide empty, once the host has re-keyed it", async () => {
    hold.draw = "nowhere";
    faults.newSlideIdSettlesAfter = 1;
    const r = await bigChart();
    expect(r.ok, "a draw that landed nowhere passed").toBe(false);
    expect(r.detail).toMatch(/came back empty/);
  });

  it("fails a draw that reports success and leaves the slide empty, before the host re-keys it", async () => {
    // The half that always worked — kept so the repair cannot trade it away.
    hold.draw = "nowhere";
    const r = await bigChart();
    expect(r.ok, "a draw that landed nowhere passed").toBe(false);
    expect(r.detail).toMatch(/came back empty/);
  });

  it("fails when the slide it added is gone and the chart was drawn somewhere else", async () => {
    hold.draw = "vanishes";
    faults.newSlideIdSettlesAfter = 1;
    const r = await bigChart();
    expect(r.ok, "passed with the chart on a slide other than the one it added").toBe(false);
    expect(r.detail).toMatch(/no longer in the deck/);
  });
});

/* ───────────────────────────────── r2-62 ───────────────────────────────── */

/**
 * Three own-slide scenarios compared raw `slideCount()` results. A count that
 * comes back with no value turns `x >= y` into false and drops "the slide could
 * not be deleted" from the problems (the quiet costume), and turns `x - y !== 1`
 * into a hard FAILED "grew by NaN" (the loud one). `deckGrowth` and
 * `countlessSkip` were written for exactly this and swept six sites; these three
 * name their counts differently and the guard that claimed to count every site
 * keyed on the names it had met.
 */
describe("the own-slide scenarios and an unread slide count", () => {
  async function run(name: string, opts: { blind: boolean; refuseDelete: boolean }) {
    const deck = [makeSlide("s1")];
    hold.deck = deck;
    installHost(deck);
    hold.blindAfterAdd = opts.blind;
    faults.refuseSlideDelete = opts.refuseDelete;
    const r = byName(await runSelfTest("probe", name), name);
    return { r, deckSize: deck.length };
  }
  const readAsAPass = (r: ScenarioResult) => r.ok && !r.skipped;

  for (const name of ["stop a run mid-draw", "a big chart on a slide of its own"]) {
    it(`${name}: a slide that would not delete is still a failure while the count can be read`, async () => {
      const { r } = await run(name, { blind: false, refuseDelete: true });
      expect(r.ok, "the arm is vacuous — a refused delete was not reported").toBe(false);
      expect(r.skipped).toBeFalsy();
    });

    it(`${name}: an unread count does not turn a slide left in the deck into a pass`, async () => {
      const { r } = await run(name, { blind: true, refuseDelete: true });
      expect(readAsAPass(r), `passed with its slide still in the deck: ${r.detail}`).toBe(false);
      expect(r.skipped, "an unread count is a skip, like every other site").toBe(true);
      expect(r.blind, "lost sight of the deck must read as blind, not as a missing feature").toBe(true);
    });
  }

  it("stop a run mid-draw: its count is also the check that a reported delete really took, so an unread one is not a pass", async () => {
    const { r } = await run("stop a run mid-draw", { blind: true, refuseDelete: false });
    expect(readAsAPass(r), `passed on no evidence: ${r.detail}`).toBe(false);
    expect(r.skipped).toBe(true);
  });

  it("a big chart on a slide of its own: a delete the deck listing confirmed needs no count to stand", async () => {
    // The count only decides the case where the delete said no. When it said yes
    // it was checked against a fresh listing of the deck, so an unread count has
    // nothing to add — and skipping on it would discard a real pass.
    const { r } = await run("a big chart on a slide of its own", { blind: true, refuseDelete: false });
    expect(r.skipped, r.detail).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
  });

  it("a chart handed over as a file: an unread count is not a growth of NaN", async () => {
    const { r } = await run("a chart handed over as a file", { blind: true, refuseDelete: false });
    expect(r.detail, "unknown arrived dressed as a measurement").not.toMatch(/NaN/);
    expect(r.skipped, "an unread count is a skip").toBe(true);
    expect(r.blind).toBe(true);
  });

  it("a chart handed over as a file: still passes when the counts are readable", async () => {
    const { r } = await run("a chart handed over as a file", { blind: false, refuseDelete: false });
    expect(r.skipped, r.detail).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
  });

  describe("the guard that claims to count every site", () => {
    const live = readFileSync("src/taskpane/selftest.ts", "utf8");

    it("finds no raw comparison or subtraction of a slide count in the self-test", () => {
      expect(rawSlideCountUses(live)).toEqual([]);
    });

    it("can see the shapes the old guard could not (a deliberately broken copy)", () => {
      // The pre-fix lines, verbatim in shape. Each names its counts differently
      // from the three spellings the regex guard keyed on.
      const broken = {
        "stop a run mid-draw": `const s = async () => { const slideCountBefore = await slideCount(); const slideCountAfter = await slideCount(); return slideCountAfter >= slideCountBefore && "x"; };`,
        "a big chart on a slide of its own": `const s = async () => { const slideCountBefore = await slideCount(); const slideCountAfter = await slideCount(); return !swept && slideCountAfter > slideCountBefore && "x"; };`,
        "a chart handed over as a file (growth)": `const s = async () => { const countBefore = await slideCount(); const countAfter = await slideCount(); return countAfter - countBefore !== 1 && "x"; };`,
        "a chart handed over as a file (swept)": `const s = async () => { const countBefore = await slideCount(); const countSwept = await slideCount(); return !swept && countSwept > countBefore && "x"; };`,
      };
      for (const [site, text] of Object.entries(broken))
        expect(rawSlideCountUses(text), `the guard cannot see: ${site}`).not.toEqual([]);
    });

    it("lets a count be handed to deckGrowth, printed, and used to size a wait", () => {
      const fine = `const s = async () => {
        const before = await slideCount();
        const after = await slideCount();
        const grew = deckGrowth(before, after);
        const settled = await settledSlideCount(before + 4);
        return \`deck \${before} to \${after} to \${settled ?? "?"}\${grew}\`;
      };`;
      expect(rawSlideCountUses(fine)).toEqual([]);
    });
  });
});

/* ───────────────────────────────── r2-63 ───────────────────────────────── */

/**
 * `shape-add-held-slide-proxy-again` declares `burnsTheSlide: true` — "it gives
 * up its slide too rather than handing the wreckage to whatever runs next" — and
 * nothing ever read the flag on a partner. The main loop checks it for the
 * trigger only, so `binding-names-shape-later` was asked first on the slide the
 * partner had just written through a dead proxy. Round 497's first sample of it
 * (`commit-threw`, on `reused-slide/replaced`) is the one that locked the row.
 */
describe("a partner question that wrecks the slide", () => {
  const PARTNER = "shape-add-held-slide-proxy-again";
  const NEXT = "binding-names-shape-later";

  async function sheetWithTrace() {
    installHost([makeSlide("s1")]);
    setTracing(true);
    const sheet = await runHostProbes("fake", "test", { passes: 3 });
    return { sheet, entries: traceLog().entries.filter((e) => e.scope === "probe") };
  }

  it("gives up the slide it wrecked, as its trigger does", async () => {
    const { entries } = await sheetWithTrace();
    const asked = entries.filter(
      (e) => e.message === "partner answered" && (e.data as { id?: string })?.id === PARTNER,
    );
    const gaveUp = entries.filter(
      (e) =>
        e.message === "giving up the scratch slide this question wrecked" &&
        (e.data as { id?: string })?.id === PARTNER,
    );
    expect(asked.length, "the partner never fired — the run is vacuous").toBeGreaterThan(0);
    expect(gaveUp.length, "the partner's burnsTheSlide flag was never read").toBe(asked.length);
  });

  it("does not ask the next question on the partner's slide", async () => {
    const { sheet } = await sheetWithTrace();
    const next = sheet.answers.find((a) => a.id === NEXT)!;
    expect(next.samples?.length, "the next question was never sampled").toBeGreaterThan(0);
    const states = next.samples!.map((s) => `${s.scratch}/${s.acquired}`);
    expect(
      states.filter((s) => s.startsWith("reused-slide")),
      `${NEXT} started on a slide that had already been used: ${states.join(", ")}`,
    ).toEqual([]);
  });
});
