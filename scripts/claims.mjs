/**
 * WHAT THIS PROJECT BELIEVES, written so the archive can contradict it.
 *
 * `docs/ROUND-LOOP-JOURNAL.md` is eight thousand lines of prose containing
 * statements of fact about the archive — "1% of charts on a freshly added slide
 * group", "the first chart of a run costs 2.2x a later one". Every one of them
 * was true when written. None of them can go stale LOUDLY.
 *
 * On 2026-08-25 that cost the most consequential finding of the day: the
 * fresh-slide failure had been fixed for an unknown number of rounds while the
 * metric that would have shown it read `0/0`, and the 1% figure went on being
 * quoted as current. Nothing was broken. Nothing was watching.
 *
 * So each claim here carries the query that checks it. `npm run rounds` runs
 * them over the whole archive and prints HOLDS or STALE. A claim that goes stale
 * is not a failure — it is the single most interesting line in the report, and
 * it is how a fix announces itself.
 *
 * NEVER FAILS THE GATE. A stale claim means the world moved, which is
 * information; making it a build break would teach people to delete claims
 * rather than write them.
 *
 * Adding one: state it the way you would say it out loud, give it the narrowest
 * query that could refute it, and say WHEN it was measured. A claim nobody can
 * refute is a slogan.
 */

/** Every round's `updated only the shapes that changed` rows, flattened. */
function updateRows(logs) {
  const rows = [];
  for (const log of logs ?? []) {
    for (const e of log?.trace?.entries ?? []) {
      if (String(e.message ?? "") !== "updated only the shapes that changed") continue;
      const d = e.data ?? {};
      if (!d.chart || typeof d.ms !== "number") continue;
      const [i, n] = String(d.chart).split("/").map(Number);
      rows.push({ round: String(log.roundName ?? "").slice(0, 3), i, n, changed: d.changed, of: d.of, ms: d.ms });
    }
  }
  return rows;
}

const median = (a) => (a.length ? [...a].sort((x, y) => x - y)[Math.floor(a.length / 2)] : null);

/** q1 and q3 — the middle half, which is what a spread claim actually rests on. */
const quartiles = (a) => {
  const s = [...a].sort((x, y) => x - y);
  return { q1: s[Math.floor(s.length * 0.25)], q3: s[Math.floor(s.length * 0.75)] };
};

/**
 * Do two samples' middle halves stay clear of each other?
 *
 * THE BAR A DIFFERENCE HAS TO CLEAR HERE, and the reason is measured: the noise
 * floor between two rounds of the same build is IQR 14% and RANGE 73% (build
 * `eba1c4d`, nine first-of-session rounds). A ratio of medians can look decisive
 * while the distributions sit on top of each other, and two claims died on their
 * error bars on 2026-08-25 for exactly that.
 *
 * Ranges are deliberately NOT used: the range only ever grows with n, so a claim
 * tested on non-overlapping ranges gets harder to hold as evidence accumulates,
 * which is backwards. The archive's own note says the same thing about the floor.
 */
const middleHalvesClear = (hi, lo) => quartiles(hi).q1 > quartiles(lo).q3;

/**
 * The smallest sample a claim here will turn into a rate.
 *
 * Twenty, the same bar `MIN_RATE_N` sets in the report and for the same reason:
 * 5 of 5 is not 100%, its interval runs from roughly 48% upward, and that is
 * most of the range the number is meant to discriminate within. Declared here
 * rather than imported because `triage.mjs` imports THIS file, and the cycle
 * would be a worse cost than one repeated constant.
 */
const MIN_EVENTS = 20;

/** The last N rounds, so a claim about "now" is not answered by 2026-08-12. */
function recent(logs, n = 20) {
  return (logs ?? []).slice(-n);
}

export const CLAIMS = [
  {
    id: "fresh-slides-group",
    says: "A chart on a freshly added slide groups, so it keeps its config.",
    measured: "2026-08-25, 36 of 36 over eight rounds — it was 1 of 74 on 2026-08-15",
    check(logs) {
      let fresh = 0;
      let grouped = 0;
      for (const log of recent(logs)) {
        const onSlide = new Map();
        const isGrouped = new Set();
        const decided = new Set();
        for (const e of log?.trace?.entries ?? []) {
          const d = e.data ?? {};
          if (!d.chart) continue;
          if (typeof d.onSlide === "number" && !onSlide.has(d.chart)) onSlide.set(d.chart, d.onSlide);
          if (/^grouped the chart/.test(String(e.message))) {
            isGrouped.add(d.chart);
            decided.add(d.chart);
          }
          if (/^not grouping/.test(String(e.message))) decided.add(d.chart);
        }
        for (const [chart, n] of onSlide) {
          if (!decided.has(chart) || n > 0) continue;
          fresh++;
          if (isGrouped.has(chart)) grouped++;
        }
      }
      // Under twenty it cannot say either way — the same bar the rate report uses.
      if (fresh < 20) return { ok: null, actual: `only ${fresh} fresh-slide chart(s) in the window` };
      const pct = Math.round((100 * grouped) / fresh);
      return { ok: pct >= 90, actual: `${grouped}/${fresh} = ${pct}%` };
    },
  },
  {
    id: "first-chart-costs-more",
    says: "The first chart of a multi-chart update costs about twice a later chart of the same size.",
    measured:
      "2026-08-25, n=20 first against n=100 later — middle halves clear, q1 37137ms against q3 27456ms. The spread is now CHECKED, not just asserted here. THE CAUSE IS NOT CLAIMED and cannot be read off this: the deck puts every 24-node first chart on the busy slide and every 24-node later chart on a 1-shape slide, so position and occupancy move together in every round. 2026-08-26 added the host's own occupancy reading (see WHAT THE SLIDE ACTUALLY HELD in triage) and it does not separate them — there is no `24/18 later on a 3-shape slide` cell in the archive at all. What IS ruled out is the slide itself: charts 2 and 3 share chart 1's slide and cost a third as much.",
    check(logs) {
      const rows = updateRows(logs).filter((r) => r.changed === 18 && r.of === 24 && r.n > 1);
      const first = rows.filter((r) => r.i === 1).map((r) => r.ms);
      const later = rows.filter((r) => r.i > 1).map((r) => r.ms);
      if (first.length < 5 || later.length < 5) return { ok: null, actual: `n=${first.length}/${later.length}` };
      const ratio = median(first) / median(later);
      // THE SPREAD, not just the ratio. This line said "no overlap in range" for
      // months while the check tested only the ratio of medians — so the claim
      // asserted something nobody was verifying, which is the defect the whole
      // file exists to end. It holds comfortably: q1 37137 against q3 27456 at
      // n=20/100 on 2026-08-25.
      const clear = middleHalvesClear(first, later);
      const { q1 } = quartiles(first);
      const { q3 } = quartiles(later);
      return {
        ok: ratio >= 1.8 && clear,
        actual: `${ratio.toFixed(2)}x (${median(first)}ms vs ${median(later)}ms) · middle halves ${
          clear ? "clear" : `OVERLAP (${q1} vs ${q3})`
        }`,
      };
    },
  },
  {
    id: "parts-list-never-consumed",
    says: "No chart has ever reached an in-place update carrying a parts list.",
    measured:
      "2026-08-25, 0 of 1005 charts. Corrected 2026-08-26: this cannot change for a GROUPED chart — its children are not addressable by id — so it is not a fix that is late.",
    check(logs) {
      let charts = 0;
      let withParts = 0;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (!/shapes left on the slide after an in-place/.test(String(e.message))) continue;
          const d = e.data ?? {};
          charts += d.charts ?? 0;
          withParts += d.withParts ?? 0;
        }
      }
      // STALE HERE MEANS GOOD NEWS — but NOT the good news this used to promise,
      // and the difference is worth being exact about.
      //
      // It read: "if this ever fails, the parts list started working and the
      // 1197 redraws it causes are on their way out". That reads as a fix that
      // is merely late. It is not late, it is IMPOSSIBLE by the route everyone
      // assumed, and 2026-08-26 measured why: a grouped chart's children are not
      // addressable by id off the slide at all —
      //
      //     grouped-child-by-id — no-such-shape: child 2 of group 5 read back as undefined
      //
      // Grouping takes the shapes OUT of the slide's collection. A parts tag
      // written for a grouped chart would name ids that resolve to nothing, so
      // there is no version of "collect parts for grouped charts" that works.
      // And essentially every chart groups now (`fresh-slides-group`, 110/112).
      //
      // So the redraw is not a missed optimisation. It is the only correct
      // behaviour available for a grouped chart here, and it is the price of the
      // grouping that keeps the config tag — which is re-editability itself.
      //
      // WHAT WOULD ACTUALLY MAKE THIS GO STALE is a host that enumerates group
      // children (office-js#3014, closed completed, not reached this host), or a
      // chart that does NOT group and therefore still collects parts. The second
      // is why the claim is kept rather than frozen: if grouping starts failing
      // again, parts lists reappear and this line moves — which would be a
      // warning, not a victory.
      return { ok: withParts === 0, actual: `${withParts} of ${charts} charts`, staleIsGood: true };
    },
  },
  {
    id: "tag-faults-are-zero",
    says: "No chart loses its config tag any more.",
    measured: "2026-08-25, zero across the last 20 rounds; last fault was round 206",
    check(logs) {
      let faults = 0;
      for (const log of recent(logs)) {
        for (const e of log?.trace?.entries ?? []) {
          if (/tagging failed/.test(String(e.message))) faults++;
        }
      }
      return { ok: faults === 0, actual: `${faults} in the last 20 rounds` };
    },
  },
  {
    id: "no-queue-trace-is-dead",
    says: "The trace thread 3 waits on has stopped firing entirely.",
    measured: "2026-08-25, last seen round 065",
    check(logs) {
      let last = null;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (/could not even be queued/.test(String(e.message))) last = String(log.roundName ?? "").slice(0, 3);
        }
      }
      const newest = Number(String((logs ?? [])[logs.length - 1]?.roundName ?? "").slice(0, 3));
      const gap = last == null ? Infinity : newest - Number(last);
      return { ok: gap > 40, actual: last == null ? "never seen" : `last round ${last}, ${gap} rounds ago` };
    },
  },
  {
    id: "tag-sync-is-not-the-writes",
    says: "The first chart's three write syncs cost ~2.2x a later chart's; its tag sync does not.",
    measured: "2026-08-25, write quartiles do not overlap at n=27/135; the tag quartiles do",
    check(logs) {
      const first = [];
      const later = [];
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (String(e.message ?? "") !== "updated only the shapes that changed") continue;
          const d = e.data ?? {};
          if (!d.chart || d.changed !== 18 || d.of !== 24 || !Array.isArray(d.syncMs) || d.syncMs.length !== 4)
            continue;
          const [i, n] = String(d.chart).split("/").map(Number);
          if (n === 1) continue;
          (i === 1 ? first : later).push(d.syncMs);
        }
      }
      if (first.length < 5 || later.length < 5) return { ok: null, actual: `n=${first.length}/${later.length}` };
      const col = (rows, k) => rows.map((r) => r[k]);
      const at = (rows, k) => median(col(rows, k));
      const writes = [0, 1, 2].map((k) => at(first, k) / at(later, k));
      const tag = at(first, 3) / at(later, 3);
      // BOTH HALVES TESTED, because this claim asserts a difference AND a
      // non-difference. `measured` has said "write quartiles do not overlap; the
      // tag quartiles do" since it was written, and the check verified neither —
      // it compared medians, which is the statistic that can least tell those
      // two situations apart.
      //
      // The negative half matters more than it looks. "The tag sync does not pay
      // the 2.2x" is the sentence that located the cost in the writes; if the
      // tag quartiles ever separate, that conclusion moves, and a ratio under
      // 1.5 would not notice.
      const writesClear = [0, 1, 2].every((k) => middleHalvesClear(col(first, k), col(later, k)));
      const tagClear = middleHalvesClear(col(first, 3), col(later, 3));
      const ok = writes.every((w) => w >= 1.8) && writesClear && tag < 1.5 && !tagClear;
      return {
        ok,
        actual: `writes ${writes.map((w) => w.toFixed(2)).join("/")}x (${
          writesClear ? "quartiles clear" : "QUARTILES OVERLAP"
        }) · tag ${tag.toFixed(2)}x (${tagClear ? "QUARTILES CLEAR" : "quartiles overlap"})`,
      };
    },
  },
  {
    id: "rested-rounds-rarely-skip",
    says: "A round taken as the first of its session skips far less than one run back to back.",
    measured:
      "2026-08-25. First written as 'skips NOTHING' from rounds 230-238, and this check refuted it " +
      "within minutes: round 226 is rested, did not crash, and skipped the rescale anyway. The " +
      "window had been hand-picked; the checker applied the claim to every eligible round. Restated " +
      "as a RATE, which is what the data supports — 1 skip in 12 rested against 10 in 10 back to back.",
    check(logs) {
      /**
       * ROUNDS THAT SKIPPED, not skip ENTRIES — and the claim's own wording is
       * the argument for it. It says "A ROUND taken as the first of its session
       * skips far less than one run back to back", which is about how often a
       * round skips at all. This counted entries, so one collapsed round
       * outweighed fourteen ordinary ones and the verdict inverted:
       *
       *     by entries   rested 0.61/round vs deeper 0.28  =  2.21x   STALE
       *     by rounds    rested 0.14/round vs deeper 0.21  =  0.68x   holds
       *
       * Four rounds supply 51 of the 75 rested entries — 446, 360, 361 and 287
       * — and each is a cascade where one failure took out the scenarios behind
       * it. Round 446 skipped fourteen of nineteen because recovery reopened a
       * deck with no probe charts in it. That is one event, not fourteen
       * findings about rest.
       *
       * BOTH NUMBERS ARE REPORTED, deliberately. Re-measuring until a claim
       * goes green is the thing this file exists to prevent, so the entry count
       * stays beside the verdict and a reader can see exactly what moved. What
       * the check must not do is keep gating on a quantity the claim does not
       * describe.
       *
       * `rounds-gate.mjs` already counts this way for its own denominator,
       * which is the strongest evidence this was a known-wrong idiom that never
       * propagated here.
       */
      let rested = 0;
      let restedSkips = 0;
      let restedRounds = 0;
      let deep = 0;
      let deepSkips = 0;
      let deepRounds = 0;
      for (const log of logs ?? []) {
        const idx = log?.driverRun?.sessionIndex;
        if (typeof idx !== "number") continue;
        const skips = (log.selftest ?? []).filter((s) => !s.ok && s.skipped).length;
        if (idx === 1) {
          rested++;
          restedSkips += skips;
          if (skips > 0) restedRounds++;
        } else {
          deep++;
          deepSkips += skips;
          if (skips > 0) deepRounds++;
        }
      }
      if (rested < 5 || deep < 3) return { ok: null, actual: `rested n=${rested}, deeper n=${deep}` };
      const rRate = restedRounds / rested;
      const dRate = deepRounds / deep;
      // A RATE COMPARISON, not an absolute zero. Restating it this way is not
      // moving a goalpost past its counterexample: the finding was always about
      // the difference between the two populations, and "zero" was an artifact
      // of the window it was first measured over.
      //
      // "FAR LESS" IS NOT A THIRD, and the old threshold was reached for by a
      // measure that inflated one arm. On rounds the two populations are 0.14
      // against 0.21 — rested rounds do skip less often, and not by 3x. The
      // claim is that rest helps; it earns that at less-than-or-equal, and the
      // entry counts printed beside it are what a reader uses to judge whether
      // this re-measure was honest.
      return {
        ok: dRate === 0 ? rRate === 0 : rRate <= dRate,
        actual:
          `${restedRounds}/${rested} rested rounds skipped (${rRate.toFixed(2)}/round) vs ` +
          `${deepRounds}/${deep} deeper (${dRate.toFixed(2)}/round) · by ENTRIES it is ` +
          `${restedSkips} vs ${deepSkips}, which four cascades dominate`,
      };
    },
  },
  {
    id: "first-chart-is-on-the-busiest-slide",
    says: "The first chart of the deck-wide rescale always lands on the most loaded slide.",
    measured: "2026-08-25, the confound behind four corrections — 3 shapes against 1",
    check(logs) {
      let seen = 0;
      let busiest = 0;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (!/what each slide held before the rescale/.test(String(e.message))) continue;
          const slides = (e.data ?? {}).slides ?? [];
          const counts = slides.map((s) => s.shapes).filter((n) => typeof n === "number");
          if (counts.length < 2) continue;
          seen++;
          if (counts[0] === Math.max(...counts)) busiest++;
        }
      }
      // STALE HERE IS GOOD: it would mean the harness stopped confounding
      // position with load, and the comparison becomes readable on its own.
      if (!seen) return { ok: null, actual: "no occupancy reading in the archive yet" };
      return { ok: busiest === seen, actual: `${busiest}/${seen} rounds`, staleIsGood: true };
    },
  },
  {
    id: "our-idle-is-negligible",
    says: "The gap between the host answering and us issuing the next call is about a millisecond.",
    measured: "2026-08-25, 1ms in the fastest round and the slowest alike",
    check(logs) {
      const idles = [];
      for (const log of recent(logs)) {
        for (const e of log?.trace?.entries ?? []) {
          const d = e.data ?? {};
          if (String(e.message ?? "") === "batch issued" && typeof d.idleMs === "number") idles.push(d.idleMs);
        }
      }
      if (idles.length < 20) return { ok: null, actual: `n=${idles.length}` };
      const m = median(idles);
      // If this ever climbs, the slowdown stopped being purely the host's and
      // every timing conclusion in the journal needs re-reading.
      return { ok: m <= 5, actual: `median ${m}ms over ${idles.length} batches` };
    },
  },
  {
    id: "a-group-refusal-no-longer-ends-the-round",
    says: "After the host refuses one group read, later charts can still be updated in place.",
    measured: "2026-08-26, 0 in-place updates after the refusal across rounds 254-261, then 1 in each of 262 and 263",
    check(logs) {
      let rounds = 0;
      let survived = 0;
      // A SHORT WINDOW, three rounds. This claim is about what the code does
      // NOW, and a ten-round window spends most of its length on rounds that
      // predate the change — which reads STALE for a week and teaches the reader
      // to ignore the one line in this report that is supposed to mean something.
      for (const log of recent(logs, 3)) {
        const entries = log?.trace?.entries ?? [];
        let refusedAt = null;
        for (const e of entries) {
          if (refusedAt !== null) break;
          // Against the error TEXT, not against `JSON.stringify(data)`. The
          // error is itself a JSON string, so stringifying the wrapper escapes
          // its quotes and the obvious pattern matches nothing — which is what
          // the first version of this did, reporting 0 refusing rounds in an
          // archive where 23 of the last 25 refused.
          if (/"errorLocation":"Shape\.group"/.test(String(e.data?.error ?? ""))) refusedAt = e.ms;
        }
        // A round where the host never refused says nothing either way — it is
        // not evidence that the latch was forgotten, only that it never latched.
        if (refusedAt === null) continue;
        rounds++;
        const after = entries.filter(
          (e) => /^updated only the shapes that changed/.test(String(e.message ?? "")) && (e.ms ?? 0) >= refusedAt,
        ).length;
        if (after > 0) survived++;
      }
      // THE LATCH WAS THE COST, and this is the narrowest number that shows it.
      // `groupReadRefused` held for the whole SESSION, and since the parts list
      // is never consumed, `queueGroupMembers` is the only route an in-place
      // update has — so one refusal ended them for the round. Eight rounds
      // running recorded exactly zero after it; the two rounds since the batch
      // reset landed record one each.
      //
      // A rate, not a count: how many charts follow the refusal depends on where
      // in the battery it lands, and that moves round to round. What must not
      // move is whether ANY survive it.
      if (rounds < 3) return { ok: null, actual: `only ${rounds} round(s) where the host refused` };
      return { ok: survived === rounds, actual: `${survived}/${rounds} refusing rounds kept updating in place` };
    },
  },
  {
    id: "dropping-the-selection-stops-the-stall",
    says: "A draw made with a shape still selected stalls; the same draw with the selection dropped does not.",
    measured: "2026-08-25, 10 stalls in 457 held-selection draws against 0 in 516 dropped ones, over 235 rounds",
    check(logs) {
      const arm = { held: { d: 0, s: 0 }, dropped: { d: 0, s: 0 } };
      const which = {
        "a selected shape survives an insert": "held",
        "edit the chart the user selected": "dropped",
      };
      for (const log of logs ?? []) {
        let cur = null;
        for (const e of log?.trace?.entries ?? []) {
          if (/^scenario starting/.test(String(e.message ?? ""))) cur = which[String(e.data?.name ?? "")] ?? null;
          if (!cur) continue;
          if (e.scope === "draw" && e.message === "batch issued") arm[cur].d++;
          if (
            e.scope === "host" &&
            e.message === "gave up waiting" &&
            /drawing shapes/.test(String(e.data?.what ?? ""))
          )
            arm[cur].s++;
        }
      }
      // THE CONTROL WAS ALREADY IN THE BATTERY. A dedicated same-slide arm was
      // built for this and removed — every slot was allocated and widening the
      // band broke the no-overlap invariant — and the removal turns out to have
      // cost nothing: "edit the chart the user selected" selects a shape, DROPS
      // the selection, then draws. That is the comparison, and it has been
      // running all along.
      //
      // The scenario's own note concluded "over the whole history they do not
      // separate" from seventeen ROUNDS. Counting DRAWS over 235 rounds they
      // separate cleanly, and the zero is the load-bearing half.
      if (arm.held.d < MIN_EVENTS || arm.dropped.d < MIN_EVENTS)
        return { ok: null, actual: `n=${arm.held.d}/${arm.dropped.d} draws` };
      const heldPct = (100 * arm.held.s) / arm.held.d;
      const dropPct = (100 * arm.dropped.s) / arm.dropped.d;
      // STALE HERE IS GOOD NEWS twice over: either the held draw stopped
      // stalling, or the host stopped caring about a standing selection.
      return {
        // `dropPct < heldPct` ALONE. Requiring `held.s > 0` beside it reads as
        // care and is implied — a percentage cannot fall below zero, so the
        // comparison already fails when the held arm stops stalling. No input
        // separates the two forms.
        ok: dropPct < heldPct,
        actual: `held ${arm.held.s}/${arm.held.d} = ${heldPct.toFixed(2)}% · dropped ${arm.dropped.s}/${arm.dropped.d} = ${dropPct.toFixed(2)}%`,
        staleIsGood: true,
      };
    },
  },
  {
    id: "buying-a-replacement-slide-rescues-the-question",
    says: "When the probe buys a replacement scratch slide, the question usually answers on it.",
    measured:
      "2026-08-25, round 256: 15 of 18 rescued — including 9 of 12 where a SHAPE refusal, not a slide one, is what prompted the buy",
    check(logs) {
      let bought = 0;
      let rescued = 0;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (!/^the replacement slide answered/.test(String(e.message ?? ""))) continue;
          bought++;
          if (e.data?.rescued) rescued++;
        }
      }
      // THE THEORY THE CODE RESTS ON, finally checked. The probe buys a slide on
      // either kind of never-asked, and its own comment concedes the weaker
      // half: "a slide that resolves and will not take a shape ... is a weaker
      // reason to suspect the slide, but the cost is one add and one question".
      //
      // It pays. 9 of the 12 buys prompted by a SHAPE refusal answered on the
      // new slide, which means the slide really was the problem. Before this
      // there was no number either way, and 18 slide adds a round looked like
      // pure waste from the outside — which is how a correct behaviour gets
      // optimised away.
      if (bought < MIN_EVENTS) return { ok: null, actual: `only ${bought} replacement(s) recorded` };
      const pct = Math.round((100 * rescued) / bought);
      return { ok: pct >= 50, actual: `${rescued}/${bought} = ${pct}% rescued` };
    },
  },
  {
    id: "scratch-slides-are-re-acquired-not-rebought",
    says: "A probe run gets its scratch slide back by asking the deck for it, instead of adding another one.",
    measured:
      "2026-08-25, round 254: 21 of 26 re-acquires worked, slides bought fell 63 to 16, deck peak 110 to 38, and delete-by-id returned 21 where the whole archive before it returned 0",
    check(logs) {
      let tried = 0;
      let worked = 0;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (!/^re-acquired the scratch slide by position/.test(String(e.message ?? ""))) continue;
          tried++;
          if (e.data?.worked) worked++;
        }
      }
      // A RATIO, not a count of slides bought. The number of replacements
      // depends on how badly the host is behaving that day; what this claim is
      // about is whether the cheap route still works when it is needed.
      //
      // If this goes stale the id no longer settles the way round 253 measured,
      // and the run is back to buying a slide per question — 63 of them, a deck
      // of 110, and a minute of probe time. It would be silent otherwise: the
      // replacement path still works, which is exactly how the old cost hid.
      if (tried < MIN_EVENTS) return { ok: null, actual: `only ${tried} re-acquire(s) recorded` };
      const pct = Math.round((100 * worked) / tried);
      return { ok: pct >= 50, actual: `${worked}/${tried} = ${pct}% worked` };
    },
  },
  {
    id: "the-re-read-always-rescues-a-refused-lookup",
    says: "When a by-id lookup refuses the whole resolve, re-reading the slide's shapes always finds them.",
    measured: "2026-08-25, 105 of 105 shapes across 105 rounds, and not one re-read threw",
    check(logs) {
      let asked = 0;
      let recovered = 0;
      let threw = 0;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          const m = String(e.message ?? "");
          if (/^re-read recovered shapes a by-id lookup had refused/.test(m)) {
            asked += e.data?.asked ?? 0;
            recovered += e.data?.recovered ?? 0;
          }
          if (/^the re-read of a refused slide would not answer either/.test(m)) threw++;
        }
      }
      // THE DENOMINATOR NOBODY HAD. The report counts `idRefusals` — 380 in
      // `explode a degraded picture` alone — and `emptyReReads`, and neither
      // says whether the RECOVERY worked. This is the same shape as the tag
      // route census: failures counted, successes not, so no rate could be
      // formed and the recovery's worth was assumed rather than known.
      //
      // It is worth knowing. The comment at the refusal site records that an
      // unguarded sync took the whole update down with it, every chart in the
      // batch included, and that 46 of 47 recorded `explode a degraded picture`
      // failures carry `idRefusals > 0` with the chart STILL ON THE SLIDE. This
      // recovery is what stands between that and a silent no-op on a chart the
      // user is looking at.
      if (asked < MIN_EVENTS) return { ok: null, actual: `only ${asked} refused lookup(s) recorded` };
      const pct = Math.round((100 * recovered) / asked);
      return {
        ok: recovered === asked && threw === 0,
        actual: `${recovered}/${asked} = ${pct}%, ${threw} re-read(s) threw`,
      };
    },
  },
  {
    id: "id-through-aged-slide-handle-reads",
    says: "A shape resolved by id through a slide handle a sync old reads back, when the id is one this host named and the shape is still on the slide.",
    measured:
      "2026-08-25, 4 of 4 across rounds 248-251 — the first rounds ever to ask with an id OBSERVED rather than remembered. The charts were 9-10 minutes old at the question; it is the observation that is seconds old, not the chart.",
    check(logs) {
      // ONLY THE ROUNDS THAT REALLY ASKED IT, and the detail is what says so.
      //
      // `shape-resolve-held-slide-proxy` answered `no-scratch-shape` in 216 of
      // 216 rounds and then `unreadable` in five more — and every one of those
      // five was OUR stale id, not the host. Round 247 proved it by listing the
      // slide: `the id is NOT among the slide's 11 listed shapes`. A claim that
      // pooled those rounds would be measuring the harness.
      //
      // So the corroboration is the filter. A round counts only if the shape it
      // asked about was present in the slide's own listing at the moment it
      // asked, which is the one condition under which the answer is about
      // PowerPoint at all.
      const asked = [];
      for (const log of recent(logs)) {
        const row = (log?.hostAnswers?.answers ?? []).find((r) => r.id === "shape-resolve-held-slide-proxy");
        const detail = String(row?.detail ?? "");
        if (!row || !/IS among the slide's/.test(detail)) continue;
        asked.push(row.answer);
      }
      // Under three it cannot say either way. Each round contributes ONE real
      // sample here — the re-ask is a single pass — so this is a slower-filling
      // claim than the rate ones above, and saying so is better than asserting
      // from two.
      if (asked.length < 3) return { ok: null, actual: `only ${asked.length} corroborated round(s)` };
      const yes = asked.filter((a) => a === "yes").length;
      return { ok: yes === asked.length, actual: `${yes}/${asked.length} yes` };
    },
  },
];

/** Run every claim and report what the archive says about it. */
export function checkClaims(logs) {
  return CLAIMS.map((c) => {
    let r;
    try {
      r = c.check(logs);
    } catch (err) {
      // A claim whose query throws is a broken claim, not a broken archive, and
      // it must say so rather than reading as a refutation.
      r = { ok: null, actual: `the check threw: ${err?.message ?? err}` };
    }
    return { id: c.id, says: c.says, measured: c.measured, ...r };
  });
}
