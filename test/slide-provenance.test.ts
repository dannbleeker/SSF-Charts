import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import * as survey from "../scripts/slide-provenance.mjs";

const { mentions5010, slideOf, classify, survey: run } = survey;

/**
 * The archive query that nearly produced a confirmed-looking wrong answer.
 *
 * Two things in `scripts/slide-provenance.mjs` have already been wrong once
 * each, and both failed SILENTLY — one by finding nothing, one by leaving out
 * the denominator. Neither shows up as an error; both show up as a confident
 * number. So each gets a test that fails when it regresses.
 */
describe("slide provenance over the round archive", () => {
  /**
   * `debugInfo` is a JSON string INSIDE a data field, so a stringified event
   * carries two backslashes before the quote. The first version of the regex
   * allowed at most one and matched 0 of 1,369 events — an absence that reads
   * exactly like "no slide was ever recorded", which was the opposite of true.
   */
  it("finds the slide id through double escaping, not just single", () => {
    const doubled = {
      err: 'InvalidParam | code=5010 | debugInfo={"fullStatements":["var slide = slides.getItem(\\"262#1236456497\\");"]}',
    };
    expect(slideOf(doubled), "the escaping the archive actually uses").toBe("262#1236456497");

    // The plain form still works — the fix widened the match, it did not move it.
    expect(slideOf({ err: 'slides.getItem("259#490753504")' })).toBe("259#490753504");
    // A shape id is not a slide id: only `slides.getItem` counts.
    expect(slideOf({ err: 'var shape = shapes.getItem("28");' }), "a shape id is not a slide").toBeNull();
    expect(slideOf(null)).toBeNull();
  });

  it("recognises 5010 in both shapes the archive writes it", () => {
    expect(mentions5010({ debugInfo: { code: "5010" } })).toBe(true);
    expect(mentions5010({ err: "InvalidParam | code=5010 | at=x" })).toBe(true);
    expect(mentions5010({ code: "5011" }), "a neighbouring code is not this one").toBe(false);
    expect(mentions5010(undefined)).toBe(false);
  });

  it("puts an added slide in the added arm even though it is also in the inventory", () => {
    // `deck.inventory` lists EVERY slide, added ones included, so a classifier
    // that tested `all` first would report every draw as the document's own and
    // invert the entire finding.
    const added = new Set(["257#1"]);
    const all = new Set(["256#0", "257#1"]);
    expect(classify("257#1", added, all)).toBe("added");
    expect(classify("256#0", added, all)).toBe("own");
    expect(classify("999#9", added, all), "an id from neither list is not evidence for either").toBe("unknown");
    expect(classify(null, added, all)).toBe("unnamed");
  });

  /**
   * THE DENOMINATOR IS THE POINT. A survey that counted only failures would
   * have reported "1,361 of 1,369 on added slides" and read as confirmation,
   * when the real finding is that there is almost nothing else to draw on.
   */
  it("counts where draws went, not only where failures landed", () => {
    const round = {
      deck: { inventory: [{ slideId: "256#0" }, { slideId: "257#1" }], newSlides: ["257#1"] },
      trace: {
        entries: [
          { message: "batch issued", data: { onSlideKey: "257#1" } },
          { message: "batch issued", data: { onSlideKey: "256#0" } },
          { message: "batch issued", data: { onSlideKey: "(visible)" } },
          { message: "parts list outcome", data: { err: 'code=5010 slides.getItem("257#1")' } },
        ],
      },
    };
    const out = run(
      "rounds",
      () => ["001-abc.json"],
      () => JSON.stringify(round),
    );
    expect(out.rounds).toBe(1);
    expect(out.draws, "the sentinel is its own state, not an added slide").toEqual({
      added: 1,
      own: 1,
      sentinel: 1,
      unknown: 0,
    });
    expect(out.events).toBe(1);
    expect(out.failures.added).toBe(1);
    expect(out.failures.own).toBe(0);
  });
});
