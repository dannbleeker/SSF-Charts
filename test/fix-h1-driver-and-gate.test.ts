import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
// @ts-expect-error — plain .mjs tool, no types. A NAMESPACE import and a
// destructure below, for the reason `round-driver.test.ts` gives at its own
// import: a wrapped named-import list moves `from` off the line the directive
// covers.
import * as driver from "../scripts/round.mjs";
// @ts-expect-error — plain .mjs tool, no types; deliberately independent of
// src/ so it cannot inherit a bug from the code it audits.
import * as verifyDeck from "../scripts/verify-deck.mjs";
// @ts-expect-error — plain .mjs tool, no types.
import * as installCheck from "../scripts/check-published-install.mjs";
// @ts-expect-error — plain .mjs module, no types.
import * as manifestRules from "../scripts/manifest-rules.mjs";
import { readFileSync } from "fs";
import JSZip from "jszip";
import { topLevelElements } from "../src/render/ooxml";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";

const { archive, DRIVER_STAMPS, main, outcomeReceipt, readiness, SESSION_PATH } = driver;
const { readDeckBytes, faultsIn, chartTotals } = verifyDeck;
const { main: checkInstall, sweepUrls } = installCheck;
const { checkManifest, urlsIn } = manifestRules;

/**
 * An in-memory `rounds/`, so `archive` can be driven the way `collectRound`
 * drives it — a log file at one path, a directory it lists and reads back —
 * without touching the real archive.
 */
function memoryArchive(files: Record<string, string>) {
  const store = new Map(Object.entries(files));
  return {
    store,
    read: ((p: string) => {
      const body = store.get(String(p));
      if (body === undefined) throw new Error(`ENOENT: ${p}`);
      return body;
    }) as never,
    write: ((p: string, body: string) => {
      store.set(String(p), body);
    }) as never,
    list: ((dir: string) =>
      [...store.keys()].filter((k) => k.startsWith(`${dir}/`)).map((k) => k.slice(dir.length + 1))) as never,
  };
}

describe("r1-32: the twin check recognises a round the driver filed", () => {
  const raw = {
    build: "abc1234 · 2026-10-24 10:00Z",
    host: { name: "PowerPoint for the web" },
    trace: { entries: [{ message: "first leg" }] },
  };
  const download = `${JSON.stringify(raw, null, 2)}\n`;

  // `collectRound`'s own argument shape: the pane's build to expect, then the
  // three readings the driver stamps. Passed positionally because that is how it
  // calls, and the order is the thing that must stay true.
  const fileAsTheDriverDoes = (m: ReturnType<typeof memoryArchive>, over: { size?: unknown; deck?: unknown } = {}) =>
    archive(
      "log.json",
      "rounds",
      m.read,
      m.write,
      m.list,
      "abc1234",
      "size" in over ? over.size : { width: 960, height: 540 },
      { attempts: 1, recovered: [] },
      "deck" in over ? over.deck : "Presentation70",
    );

  it("refuses to file the same download a second time after the driver filed it", () => {
    // THE 039/038 SHAPE, on the path that runs every cycle. Round 039 was filed
    // byte-identical to 038 because `Download run log` did nothing after a
    // wedge and the previous round's file was still at the landing path. The
    // twin check exists for exactly that, and it compared the raw download with
    // a file `collectRound` had already stamped — so it could not match. The
    // build stamp does not help: a stale log from the SAME build passes it.
    const m = memoryArchive({ "log.json": download });
    expect(fileAsTheDriverDoes(m)).toBe("001-abc1234.json");
    expect(
      () => fileAsTheDriverDoes(m),
      "a stale same-build log was filed as a second round because its twin carries the driver's stamps",
    ).toThrow(/byte-identical to 001-abc1234\.json/);
    expect([...m.store.keys()].filter((k) => k.startsWith("rounds/"))).toEqual(["rounds/001-abc1234.json"]);
  });

  it("refuses it whatever the NEXT leg would have stamped on it", () => {
    // Within a cycle the duplicate would have carried the next leg's slide size
    // and deck, and been counted in the wrong arm as well as twice — so the
    // refusal cannot depend on the stamps agreeing.
    const m = memoryArchive({ "log.json": download });
    fileAsTheDriverDoes(m);
    expect(() => fileAsTheDriverDoes(m, { size: { width: 720, height: 540 }, deck: "Presentation43" })).toThrow(
      /byte-identical/,
    );
    // And a leg that could read neither: null is stamped (the driver looked and
    // the host would not say), which is still not a different round.
    expect(() => fileAsTheDriverDoes(m, { size: null, deck: null })).toThrow(/byte-identical/);
  });

  it("refuses it from a hand `--archive`, which passes no build and no stamps", () => {
    // The path round.mjs itself tells a person to take when the pane offers no
    // download button, and the one with NO other guard: no `expectBuild`.
    const m = memoryArchive({ "log.json": download });
    fileAsTheDriverDoes(m);
    expect(() => archive("log.json", "rounds", m.read, m.write, m.list)).toThrow(/byte-identical to 001-abc1234\.json/);
  });

  it("does not depend on how the filed copy was formatted", () => {
    // `rounds/` is checked by prettier and files have been reformatted by hand,
    // so a comparison of bytes would turn on whitespace nobody chose.
    const filed = { ...raw, driverSlideSize: null, driverDeck: "Presentation70", driverRun: { attempts: 1 } };
    const m = memoryArchive({ "log.json": download, "rounds/007-abc1234.json": JSON.stringify(filed) });
    expect(() => archive("log.json", "rounds", m.read, m.write, m.list)).toThrow(/byte-identical to 007-abc1234\.json/);
  });

  it("still files a different log from the same build", () => {
    // The guard must not become "one round per build": a cycle runs two or
    // three legs on one build and every one of them is a real round.
    const m = memoryArchive({ "log.json": download });
    fileAsTheDriverDoes(m);
    m.store.set(
      "log.json",
      `${JSON.stringify({ ...raw, trace: { entries: [{ message: "second leg" }] } }, null, 2)}\n`,
    );
    expect(fileAsTheDriverDoes(m)).toBe("002-abc1234.json");
  });

  it("is told about every field `archive` stamps, not only the ones it remembers", () => {
    // THE LIST CANNOT BE TRUSTED TO STAY WHOLE. Three stamps were added over
    // two weeks and each went in below the comparison; a fourth that is added
    // to `archive` and not to `DRIVER_STAMPS` is the same defect again. So the
    // list is checked against what `archive` really writes, with every
    // argument given a value so no stamp is skipped for being undefined.
    const m = memoryArchive({ "log.json": download });
    fileAsTheDriverDoes(m);
    const written = JSON.parse(m.store.get("rounds/001-abc1234.json")!) as Record<string, unknown>;
    const stamped = Object.keys(written).filter((k) => !(k in raw));
    expect([...DRIVER_STAMPS].sort()).toEqual(stamped.sort());
  });
});

describe("r2-29: an attempt that ran no battery takes no place in the session", () => {
  const prev = {
    at: "2026-10-24T08:00:00.000Z",
    sessionIndex: 3,
    sessionStartedAt: "2026-10-24T07:00:00.000Z",
  };
  const receipt = (over: Record<string, unknown>) =>
    outcomeReceipt({
      codes: [],
      build: null,
      size: null,
      at: "2026-10-24T08:20:00.000Z",
      sessionIndex: 4,
      sessionStartedAt: "2026-10-24T07:00:00.000Z",
      prev,
      ...over,
    });

  it.each(["checked", "not-ready", "no-run-button"])("holds the previous round's place for `%s`", (reason) => {
    // `not-ready` alone held it. A `--check` that passes returns before any
    // battery, and `no-run-button` stops before the click: neither put load on
    // the host, so neither owns a place in a session that measures load.
    const r = receipt({ reason });
    expect(r.at).toBe(prev.at);
    expect(r.sessionIndex).toBe(3);
    expect(r.sessionStartedAt).toBe(prev.sessionStartedAt);
    expect(r.heldSessionFrom, "a timestamp older than the writer must say why").toBeTruthy();
  });

  it.each(["checked", "not-ready", "no-run-button"])(
    "founds no session for `%s` when there is none to hold",
    (reason) => {
      // THE HALF A HOLD ALONE DOES NOT FIX. On a fresh clone `.round-session.json`
      // does not exist (it is gitignored), so there is nothing to hold and the old
      // code stamped `sessionIndex: 1` — the first `--check` of the day founded a
      // session and the first real round was recorded as 2.
      const r = receipt({ reason, prev: null });
      expect(r.sessionIndex).toBeNull();
      expect(r.sessionStartedAt).toBeNull();
      expect(r.heldSessionFrom).toBeTruthy();
      // The receipt keeps its own outcome and timestamp; only the CHAIN has
      // nothing to take from it.
      expect(r.reason).toBe(reason);
      expect(r.at).toBe("2026-10-24T08:20:00.000Z");
    },
  );

  it.each(["finished", "crashed", "timeout", "silent"])("still lets `%s` take its own place", (reason) => {
    // The other direction: a round that worked the host owns its place, and the
    // hold must not widen to swallow it.
    const r = receipt({ reason, roundFile: "x.json" });
    expect(r.at).toBe("2026-10-24T08:20:00.000Z");
    expect(r.sessionIndex).toBe(4);
    expect(r.heldSessionFrom).toBeUndefined();
  });
});

describe("r2-29: the documented check-then-run loop", () => {
  // Driven through the real `main`, because the fault is in what `main` WRITES
  // after an attempt and no pure function sees it. `attempt` is the only stub: it
  // asks the real `readiness` the question the real one asks, with the session
  // index `main` hands it, so the `deep-session` stop is the driver's own.
  const READY = { head: "abc1234", deployed: "abc1234", stamp: "abc1234", slides: 1, verbose: true, pictures: true };

  beforeEach(() => {
    // Date only: nothing here sleeps, and the chain is a function of the clock.
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(new Date("2026-10-24T09:00:00.000Z"));
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  function session() {
    const files = new Map<string, string>();
    const seen: { argv: string; index: number; outcome: string }[] = [];
    const attempt = async (
      argv: string[],
      deps: { driverRun?: { sessionIndex: number }; allowDeepSession?: boolean },
    ) => {
      const index = deps.driverRun?.sessionIndex ?? 1;
      const r = readiness({ ...READY, sessionIndex: index, allowDeepSession: deps.allowDeepSession ?? false });
      const checking = argv.includes("--check");
      // A check is a half-minute of probes; a round is twelve minutes of load.
      vi.setSystemTime(Date.now() + (checking ? 30_000 : 12 * 60_000));
      const outcome = !r.ok
        ? { code: 1, reason: "not-ready", codes: r.codes }
        : checking
          ? { code: 0, reason: "checked" }
          : { code: 0, reason: "finished", roundFile: "001-abc1234.json", build: "abc1234", size: "16:9" };
      seen.push({ argv: checking ? "--check" : "round", index, outcome: `${outcome.reason}` });
      return outcome;
    };
    const go = async (...argv: string[]) => {
      const code = await main(argv, {
        run: () => ({ stdout: "", status: 0 }),
        head: "abc1234",
        attempt,
        write: (p: string, body: string) => {
          files.set(p, body);
        },
        readSession: () => {
          const body = files.get(SESSION_PATH);
          return body ? JSON.parse(body) : null;
        },
      });
      // Seconds between one command and the next, as a person typing them.
      vi.setSystemTime(Date.now() + 5_000);
      return code;
    };
    return { files, seen, go };
  }

  it("runs four rounds before the deep-session stop, not two", async () => {
    // ROUND-LOOP-BRIEF "Each cycle": step 1 `--check`, step 2 the round. At HEAD
    // each check took a place, so the indices ran 1 (check), 2 (round), 3, 4 and
    // the third check landed on 5 and was refused — halving a stop that is
    // measured to matter from the fifth round (10 skips across rounds 5-10 of a
    // back-to-back block, 0 across 9 rested ones).
    const s = session();
    for (let k = 1; k <= 4; k++) {
      expect(await s.go("--check"), `check before round ${k}`).toBe(0);
      expect(await s.go(), `round ${k}`).toBe(0);
    }
    expect(
      s.seen.filter((x) => x.argv === "round").map((x) => x.index),
      "a check took a place in the session",
    ).toEqual([1, 2, 3, 4]);
    // And the stop is where it was designed to be: the fifth.
    expect(await s.go("--check")).toBe(1);
    expect(s.seen.at(-1)).toEqual({ argv: "--check", index: 5, outcome: "not-ready" });
  });

  it("leaves a fresh machine's chain unwritten by a check", async () => {
    const s = session();
    expect(await s.go("--check")).toBe(0);
    expect(s.files.has(SESSION_PATH), "a --check founded a session").toBe(false);
    // So the first real round is the first of its session.
    await s.go();
    expect(s.seen.map((x) => x.index)).toEqual([1, 1]);
    expect(JSON.parse(s.files.get(SESSION_PATH)!).sessionIndex).toBe(1);
  });

  it("does not move the chain's timestamp, so waiting out the gap still ends the session", async () => {
    const s = session();
    await s.go();
    const filed = JSON.parse(s.files.get(SESSION_PATH)!) as { at: string; sessionIndex: number };
    await s.go("--check");
    const after = JSON.parse(s.files.get(SESSION_PATH)!) as { at: string; sessionIndex: number };
    expect(after.at, "a check pushed the 45-minute gap forward from itself").toBe(filed.at);
    expect(after.sessionIndex).toBe(filed.sessionIndex);
  });
});

/**
 * ONE SLIDE, TWO CHARTS, built from a real one-chart deck by copying its group
 * and tag part — what a self-test does by inserting twice onto one slide, and
 * what round 497's inventory shows on four slides (3, 2, 5 and 8 objects named
 * PowerChart). Chart A is the original, listed FIRST in the slide's
 * relationships; chart B is appended, so it is listed LAST — which is the one
 * the old audit judged and the only one it ever looked at.
 */
async function twoChartSlide() {
  const cfg = { ...sampleConfig("line"), title: "A" };
  const built = await buildDeckBase64([
    { scene: buildChart(cfg), title: "A", configJson: JSON.stringify(cfg), slot: 0, run: "run-x" },
  ]);
  const zip = await JSZip.loadAsync(Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0)));
  const slidePath = "ppt/slides/slide1.xml";
  const relsPath = "ppt/slides/_rels/slide1.xml.rels";
  const slide = await zip.file(slidePath)!.async("string");
  const rels = await zip.file(relsPath)!.async("string");
  const types = await zip.file("[Content_Types].xml")!.async("string");

  const treeStart = slide.indexOf("</p:grpSpPr>") + "</p:grpSpPr>".length;
  const groupA = topLevelElements(slide.slice(treeStart, slide.indexOf("</p:spTree>"))).find((e) =>
    e.includes('name="PowerChart"'),
  )!;
  const ridA = /<p:tags r:id="([^"]+)"\/>/.exec(groupA)![1];
  const relA = new RegExp(`<Relationship [^>]*Id="${ridA}"[^>]*/>`).exec(rels)![0];
  const partA = /tags\/(tag\d+\.xml)/.exec(relA)![1];
  const parts = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
  const partB = `tag${Math.max(...parts.map((f) => Number(/(\d+)\.xml$/.exec(f)![1]))) + 1}.xml`;
  const ridB = "rId90";
  // Ids bumped so the duplicate-shape-id check stays quiet: these tests are
  // about the config tags, and a second fault would only muddy which one fired.
  const groupB = groupA
    .replace(/(<p:cNvPr\b[^>]*\bid=")(\d+)(")/g, (_m, a, id, c) => `${a}${Number(id) + 1000}${c}`)
    .replace(`<p:tags r:id="${ridA}"/>`, `<p:tags r:id="${ridB}"/>`);
  zip.file(
    slidePath,
    slide.replace(groupA, () => groupA + groupB),
  );
  zip.file(
    relsPath,
    rels.replace("</Relationships>", () => `${relA.replace(ridA, ridB).replace(partA, partB)}</Relationships>`),
  );
  zip.file(`ppt/tags/${partB}`, await zip.file(`ppt/tags/${partA}`)!.async("string"));
  const overrideA = new RegExp(`<Override PartName="/ppt/tags/${partA}"[^>]*/>`).exec(types)![0];
  zip.file(
    "[Content_Types].xml",
    types.replace("</Types>", () => `${overrideA.replace(partA, partB)}</Types>`),
  );

  const edit = async (path: string, change: (s: string) => string) => {
    const was = await zip.file(path)!.async("string");
    const now = change(was);
    expect(now, `the edit to ${path} changed nothing`).not.toBe(was);
    zip.file(path, now);
  };
  const corrupt = (part: string) =>
    edit(`ppt/tags/${part}`, (s) => s.replace(/(name="POWERCHART_CONFIG" val=")[^"]*"/, '$1{not json"'));
  return {
    partA,
    partB,
    ridA,
    ridB,
    edit,
    corrupt,
    slidePath,
    audit: async () => readDeckBytes(await zip.generateAsync({ type: "uint8array" })),
    zip,
  };
}

describe("r1-33: verify-deck audits every chart on a slide, not the last one listed", () => {
  it("reads two healthy charts as two charts, both re-editable", async () => {
    // The control, and the proof the fixture is a real two-chart slide: without
    // it every case below could be failing for a reason that is the fixture's.
    const audit = await (await twoChartSlide()).audit();
    expect(audit.rows).toHaveLength(1);
    expect(audit.rows[0]).toMatchObject({ chartObjects: 2, chartsReEditable: 2, configOnChart: true });
    expect(faultsIn(audit)).toEqual([]);
    expect(chartTotals(audit.rows)).toEqual({ objects: 2, reEditable: 2 });
  }, 30_000);

  it("catches a corrupt config on the FIRST chart when the last is healthy", async () => {
    // The finding, verbatim: chart A's config set to `{not json` reported
    // "no structural faults" and exited 0, because only chart B was judged.
    const s = await twoChartSlide();
    await s.corrupt(s.partA);
    const audit = await s.audit();
    expect(audit.rows[0].configMalformed, "the first chart's corrupt config was not seen").toBe(true);
    expect(faultsIn(audit).join(" ")).toMatch(/not valid JSON/i);
  }, 30_000);

  it("catches the same corruption on the LAST chart — the one it always saw", async () => {
    // Pins that the fix did not trade one blind spot for another.
    const s = await twoChartSlide();
    await s.corrupt(s.partB);
    expect(faultsIn(await s.audit()).join(" ")).toMatch(/not valid JSON/i);
  }, 30_000);

  it("catches a config the FIRST chart does not point at", async () => {
    const s = await twoChartSlide();
    await s.edit(s.slidePath, (x) => x.replace(`<p:tags r:id="${s.ridA}"/>`, ""));
    const audit = await s.audit();
    expect(audit.rows[0].configOrphaned, "an unreachable first config read as sound").toBe(true);
    expect(audit.rows[0].configOnChart).toBe(false);
    expect(faultsIn(audit).join(" ")).toMatch(/not referenced by the "PowerChart" object/i);
  }, 30_000);

  it("catches an unreachable config on the LAST chart too", async () => {
    const s = await twoChartSlide();
    await s.edit(s.slidePath, (x) => x.replace(`<p:tags r:id="${s.ridB}"/>`, ""));
    expect(faultsIn(await s.audit()).join(" ")).toMatch(/not referenced by the "PowerChart" object/i);
  }, 30_000);

  it("counts chart OBJECTS, so a chart that lost its tag is not hidden behind its neighbour", async () => {
    // THE RECURRING ROUND FAILURE, "1 of 2 new charts are re-editable" (BACKLOG;
    // rounds 102, 467, 497): chart A keeps its group and loses its tag part,
    // relationship and override. The report printed "chart objects: 1 (1
    // re-editable, 0 NOT re-editable)" and exited 0 — the truth being two
    // objects, one of them not re-editable — so the very line that failure is
    // read from could not show it.
    const s = await twoChartSlide();
    await s.edit(s.slidePath, (x) => x.replace(`<p:tags r:id="${s.ridA}"/>`, ""));
    await s.edit("ppt/slides/_rels/slide1.xml.rels", (x) =>
      x.replace(new RegExp(`<Relationship [^>]*Id="${s.ridA}"[^>]*/>`), ""),
    );
    await s.edit("[Content_Types].xml", (x) =>
      x.replace(new RegExp(`<Override PartName="/ppt/tags/${s.partA}"[^>]*/>`), ""),
    );
    s.zip.remove(`ppt/tags/${s.partA}`);
    const audit = await s.audit();
    expect(chartTotals(audit.rows)).toEqual({ objects: 2, reEditable: 1 });
    expect(audit.rows[0]).toMatchObject({ chartObjects: 2, chartsReEditable: 1 });
    // Not a FAULT: a chart object with no config is also what a title page or an
    // element is, by design — the exit code is unchanged and only the count is
    // now true.
    expect(faultsIn(audit)).toEqual([]);
  }, 30_000);

  it("counts a one-chart-per-slide deck as it always did", async () => {
    // The negative control for the totals: a deck with a chart-less slide still
    // reads 3 objects, 2 of them re-editable, exactly as the per-slide count did.
    const cfg = (title: string) => ({ ...sampleConfig("line"), title });
    const built = await buildDeckBase64([
      { scene: buildChart(cfg("A")), title: "A", configJson: JSON.stringify(cfg("A")), slot: 0, run: "run-x" },
      { scene: buildChart(cfg("B")), title: "B", configJson: JSON.stringify(cfg("B")), slot: 1, run: "run-x" },
      { scene: buildChart(cfg("C")), title: "C", slot: 2, run: "run-x" },
    ]);
    const audit = await readDeckBytes(Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0)));
    expect(chartTotals(audit.rows)).toEqual({ objects: 3, reEditable: 2 });
    expect(faultsIn(audit)).toEqual([]);
  }, 30_000);

  it("catches a malformed slot tag listed before a healthy one", async () => {
    // The same overwrite on the slot: a corrupt slot tag was cleared whenever a
    // later tag part carried a good one.
    const s = await twoChartSlide();
    const slotPart = (
      await Promise.all(
        Object.keys(s.zip.files)
          .filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f))
          .map(async (f) => [f, await s.zip.file(f)!.async("string")] as const),
      )
    ).find(([, text]) => text.includes("POWERCHART_DEMO_SLOT"))!;
    const slotRel = new RegExp(`<Relationship [^>]*Target="\\.\\./tags/${slotPart[0].split("/").pop()}"[^>]*/>`).exec(
      await s.zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string"),
    )![0];
    // A second slot tag, healthy, appended so it is listed AFTER the original;
    // then the original (listed first) is the one corrupted.
    const extra = "tag91.xml";
    s.zip.file(`ppt/tags/${extra}`, slotPart[1]);
    await s.edit("ppt/slides/_rels/slide1.xml.rels", (x) =>
      x.replace(
        "</Relationships>",
        `${slotRel.replace(/Id="[^"]+"/, 'Id="rId91"').replace(/tag\d+\.xml/, extra)}</Relationships>`,
      ),
    );
    await s.edit("[Content_Types].xml", (x) =>
      x.replace(
        "</Types>",
        `<Override PartName="/ppt/tags/${extra}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/></Types>`,
      ),
    );
    await s.edit(slotPart[0], (x) =>
      x.replace(/(name="POWERCHART_DEMO_SLOT" val=")[^"]*"/, '$1{&quot;i&quot;:0,&quot;title&quot;:&quot;X"'),
    );
    const audit = await s.audit();
    expect(audit.rows[0].slotMalformed, "a later healthy slot tag cleared an earlier corrupt one").toBe(true);
    expect(faultsIn(audit).join(" ")).toMatch(/slot tag is malformed/i);
  }, 30_000);

  it("reports a chart with no origin even when its neighbour has one", async () => {
    // `origin` was true if ANY tag part named one, so a second chart that lost
    // its origin was hidden by the first. One shape references one tags part and
    // the writer puts config and origin together in it, so each chart answers
    // for itself.
    const s = await twoChartSlide();
    // The deck builder gives every config-bearing chart an origin, so both start
    // with one — and the slide says so.
    expect((await s.audit()).rows[0].origin).toBe(true);
    await s.edit(`ppt/tags/${s.partB}`, (x) => x.replace(/<p:tag name="POWERCHART_ORIGIN" val="[^"]*"\/>/, ""));
    expect((await s.audit()).rows[0].origin, "chart B has no origin and the slide said it did").toBe(false);
  }, 30_000);
});

describe("r2-48: a network failure in the URL sweep is 'could not check', not a finding", () => {
  // The real manifests, so the sweep is reached exactly as the weekly run reaches
  // it: published and committed identical, every rule passing. Nothing here
  // touches the network — `get` is injected, and the only thing that can fail is
  // what a test makes fail.
  const manifest = (name: string) => readFileSync(name, "utf8");
  const urls = urlsIn(manifest("manifest-prod.xml")) as string[];
  const victim = urls.find((u) => /icon-16\.png$/.test(u))!;

  function run(head: (url: string) => Promise<{ ok: boolean; status: number }>, api?: () => Promise<unknown>) {
    const heads: string[] = [];
    const out: string[] = [];
    const warned: string[] = [];
    const get = async (url: string, init?: { method?: string }) => {
      if (init?.method === "HEAD") {
        heads.push(url);
        return head(url);
      }
      if (url.startsWith("https://api.github.com/")) {
        if (api) await api();
        return {
          ok: true,
          status: 200,
          json: async () => ({
            tag_name: "v9.9.9",
            assets: ["manifest-prod.xml", "manifest-excel-prod.xml"].map((name) => ({
              name,
              browser_download_url: `https://download.test/${name}`,
            })),
          }),
        };
      }
      const name = url.split("/").pop()!;
      return { ok: true, status: 200, text: async () => manifest(name) };
    };
    return {
      heads,
      out,
      warned,
      code: checkInstall({
        get: get as never,
        out: (s: string) => out.push(s),
        warn: (s: string) => warned.push(s),
      }) as Promise<number>,
    };
  }
  const answers = async () => ({ ok: true, status: 200 });

  it("reaches the sweep and calls an all-200 install path sound (the control)", async () => {
    // Non-vacuity: every case below is about the requests this one makes.
    const r = run(answers);
    expect(await r.code).toBe(0);
    expect(r.heads.length, "the sweep was never reached").toBeGreaterThan(5);
    expect(r.out.join("")).toContain("sound");
  });

  it("exits 2, not 3, when a HEAD request fails outright", async () => {
    // The finding: one `TypeError: fetch failed` among the ~23 requests became
    // "icon-16.png answers fetch failed", exit 3, and quality-sweep.yml filed or
    // reopened the install-watch issue with advice to cut a release.
    const r = run(async (url) => {
      if (url === victim) throw new TypeError("fetch failed");
      return { ok: true, status: 200 };
    });
    expect(await r.code, "a dropped connection read as a broken install path").toBe(2);
    expect(r.out, "nothing may reach stdout: the workflow files that as the issue body").toEqual([]);
    expect(r.warned.join("\n")).toContain(victim);
    expect(r.warned.join("\n")).toContain("fetch failed");
  });

  it("still reports a URL that ANSWERS 404 as a finding", async () => {
    // The other half: the fix must not turn the sweep into one that cannot fail.
    const r = run(async (url) => (url === victim ? { ok: false, status: 404 } : { ok: true, status: 200 }));
    expect(await r.code).toBe(3);
    expect(r.out.join("")).toContain(`points at ${victim}, which answers 404`);
  });

  it("keeps a real finding when another request fails, and does not blame the failure on the manifest", async () => {
    const other = urls.find((u) => u !== victim && /^https:\/\/ssf-chart\./.test(u))!;
    const r = run(async (url) => {
      if (url === other) throw new TypeError("fetch failed");
      return url === victim ? { ok: false, status: 404 } : { ok: true, status: 200 };
    });
    expect(await r.code).toBe(3);
    expect(r.out.join("")).toContain("which answers 404");
    expect(r.out.join(""), "a failed request was written into the issue body").not.toContain("fetch failed");
    expect(r.warned.join("\n")).toContain(other);
  });

  it("keeps exit 2 for a failed release fetch", async () => {
    // Already true before the change; pinned so the refactor to a returned exit
    // code did not drop it.
    const r = run(answers, async () => {
      throw new TypeError("fetch failed");
    });
    expect(await r.code).toBe(2);
    expect(r.warned.join("\n")).toContain("could not read the published install path");
    expect(r.heads).toEqual([]);
  });

  it("separates what was answered from what was not, per URL", async () => {
    const pairs = [{ name: "manifest-prod.xml", published: manifest("manifest-prod.xml") }];
    const swept = await sweepUrls(pairs, async (url: string) => {
      if (url === victim) throw new Error("ENOTFOUND");
      return { ok: false, status: 500 };
    });
    expect(swept.unreachable).toEqual([`manifest-prod.xml → ${victim}: ENOTFOUND`]);
    expect(swept.problems.length, "every other URL answered 500, which IS an answer").toBe(urls.length - 1);
    expect(swept.problems.join(" ")).not.toContain("ENOTFOUND");
  });
});

describe("r2-49: the manifest version rule enforces Microsoft's format", () => {
  const good = readFileSync("manifest-prod.xml", "utf8");
  const withVersion = (v: string) => good.replace(/<Version>[^<]+<\/Version>/, `<Version>${v}</Version>`);
  const problemsFor = (v: string) => checkManifest(withVersion(v), "manifest-prod.xml") as string[];

  it("passes the committed manifest, so every case below is about the version alone", () => {
    expect(checkManifest(good, "manifest-prod.xml")).toEqual([]);
  });

  it.each([
    ["1..0", "an empty part"],
    ["1.0.", "a trailing dot"],
    ["1.0.1.", "a trailing dot after three parts"],
    [".1.0", "a leading dot"],
    ["1e3.0", "an exponent"],
    ["0x1.0", "a hex part"],
    ["1.-1", "a negative part"],
    ["1.+1", "a signed part"],
    ["1. 0", "a space inside the number"],
    ["1.0.0.0.0", "five parts"],
    ["1.0.1.0.1", "five parts"],
    ["1.0.123456.0", "a six-digit part"],
    ["123456.0", "a six-digit major"],
  ])("rejects %s (%s)", (version) => {
    // `Number("")` is 0, `Number("1e3")` is 1000 and `Number("0x1")` is 1, so the
    // per-part `Number()` check passed every one of these. Microsoft's format is
    // 1 to 4 parts of at most 5 digits.
    const problems = problemsFor(version);
    expect(problems, `"${version}" passed the version rule`).toHaveLength(1);
    expect(problems[0]).toContain(`version "${version}"`);
  });

  it.each(["1.0", "1.0.0", "1.0.0.0", "1.0.3.0", "10.20.30.40", "12345.12345.12345.12345", "1.00"])(
    "still accepts %s",
    (version) => {
      // The other half: a tighter rule must not start refusing real versions.
      expect(problemsFor(version)).toEqual([]);
    },
  );

  it("does not make padding around the number a finding", () => {
    // Whitespace between the tags is not part of the number Office reads.
    expect(problemsFor("\n  1.0.0.0\n")).toEqual([]);
  });

  it("keeps the verdicts the existing tests pin", () => {
    expect(problemsFor("0.1.0")[0]).toContain("below 1.0");
    expect(problemsFor("1")[0]).toContain("needs at least major.minor");
    expect(problemsFor("abc")[0]).toContain("not a dotted number");
    expect(problemsFor("1.0b")[0]).toContain("not a dotted number");
  });

  it("is applied to the published manifest too", () => {
    // `check-published-install` reuses this rule, so a release that carried the
    // same malformed version as main used to read as sound.
    const bad = withVersion("1.0.0.0.0");
    const problems = installCheck.judgePublished([{ name: "manifest-prod.xml", published: bad, committed: bad }]);
    expect(problems.some((p: string) => p.includes("at most four"))).toBe(true);
  });
});
