// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { EN, localizePane, t, type StringKey } from "../src/taskpane/i18n";

/**
 * The app ships English-only, but the i18n plumbing is live so a language is a
 * drop-in. These lock the contract that keeps that drop-in safe: t() resolves the
 * catalogue, fills {placeholder} templates, and passes dynamic text through; and
 * the EN catalogue is well-formed so a future translation can't silently break it.
 */

describe("t — runtime string resolution", () => {
  it("returns catalogued English strings unchanged", () => {
    localizePane(undefined);
    expect(t("Done.")).toBe("Done.");
    expect(t("Working…")).toBe("Working…");
    expect(t("Other")).toBe("Other");
  });

  it("fills {placeholder} templates from params", () => {
    expect(t("Inserted {n} of {total} chart(s).", { n: 3, total: 4 })).toBe("Inserted 3 of 4 chart(s).");
    expect(t("Same scale applied to {n} charts (max {max}).", { n: 4, max: 90 })).toBe(
      "Same scale applied to 4 charts (max 90).",
    );
    expect(t("Failed: {error}", { error: "boom" })).toBe("Failed: boom");
  });

  it("passes an unknown / dynamic string through unchanged", () => {
    expect(t("A one-off diagnostic the catalogue never saw")).toBe("A one-off diagnostic the catalogue never saw");
  });

  it("stays English for any language while no dictionaries ship", () => {
    localizePane("de-DE");
    expect(t("Done.")).toBe("Done.");
    localizePane("fr");
    expect(t("Working…")).toBe("Working…");
  });
});

describe("EN catalogue is well-formed (a translation can build against it safely)", () => {
  it("maps every key to a non-empty string, English to itself", () => {
    for (const [k, v] of Object.entries(EN)) {
      expect(typeof v, k).toBe("string");
      expect(v.length, k).toBeGreaterThan(0);
      expect(v, `${k} must map to itself in the English catalogue`).toBe(k);
    }
  });

  it("templates and their fills use matching {placeholders}", () => {
    // Every {name} in a key also appears in its value (English is self-consistent)
    // — the per-placeholder shape a future language must preserve.
    for (const [k, v] of Object.entries(EN)) {
      const keyPlaceholders = (k.match(/\{[a-z]+\}/g) ?? []).sort();
      const valPlaceholders = (v.match(/\{[a-z]+\}/g) ?? []).sort();
      expect(valPlaceholders, k).toEqual(keyPlaceholders);
    }
  });

  it("a complete language dictionary is exactly the EN key set (the drop-in contract)", () => {
    // Simulate adding a language: a valid dict is Record<StringKey, string>, i.e.
    // its keys equal EN's — what TypeScript enforces at compile time.
    const fake: Record<StringKey, string> = Object.fromEntries(
      (Object.keys(EN) as StringKey[]).map((k) => [k, `‹${k}›`]),
    ) as Record<StringKey, string>;
    expect(Object.keys(fake).sort()).toEqual(Object.keys(EN).sort());
  });
});

describe("prototype-named keys (regression)", () => {
  // The catalogue and the dictionaries are plain objects, and the key is
  // arbitrary text — DOM content or a status string. "toString" reached
  // Object.prototype and t() returned a FUNCTION; with params it threw outright.
  it.each(["toString", "constructor", "valueOf", "hasOwnProperty", "__proto__"])(
    "t(%s) returns a string and never throws",
    (key) => {
      expect(typeof t(key)).toBe("string");
      expect(t(key)).toBe(key); // unknown key passes through unchanged
      expect(() => t(key, { n: 1 })).not.toThrow();
      expect(typeof t(key, { n: 1 })).toBe("string");
    },
  );
});
