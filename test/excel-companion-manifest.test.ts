import { describe, expect, it } from "vitest";
import { readdirSync, readFileSync } from "fs";
import { fileURLToPath } from "url";

/**
 * WHAT THE EXCEL MANIFEST CLAIMS, CHECKED AGAINST WHAT THE EXCEL CODE DOES.
 *
 * The companion was written by copying the PowerPoint manifest and editing the
 * parts that were obviously host-specific, and two things survived that edit
 * unnoticed until 2026-09-20 — both of them assertions about behaviour:
 *
 *   1. The requirement-set comment argued the floor from
 *      `ShapeCollection.addGeometricShape / addLine / addTextBox`, which this
 *      add-in never calls. In Excel those arrive at ExcelApi **1.9**, nine sets
 *      above the 1.1 written beside it, so the only reading of that block was
 *      that nobody had derived the number.
 *   2. The first-run callout said "…and insert your first chart". The companion
 *      inserts nothing anywhere — its entire Excel surface is a read of the
 *      selected range — and that string is what a validation tester meets
 *      first, because the host pops it once on install.
 *
 * Neither is the kind of thing `office-addin-manifest validate` can see: both
 * files are perfectly valid XML making perfectly well-formed false claims. So
 * this file re-derives the claims from the source rather than reading them back
 * off the manifest, in the same spirit as `test/below-1-10-census.test.ts` —
 * which is how the PowerPoint floor's published figure was caught at 8 when the
 * real count was 9.
 *
 * `manifest-excel-prod.xml` is generated from `manifest-excel.xml` by a URL
 * swap, so both are asserted: forgetting `node scripts/build-manifest.mjs`
 * fails here as well as in the staleness check.
 */

const read = (name: string) => readFileSync(fileURLToPath(new URL(`../${name}`, import.meta.url)), "utf8");

const EXCEL_MANIFESTS = ["manifest-excel.xml", "manifest-excel-prod.xml"];

const SRC = read("src/excel/excel.ts");
const HTML = read("src/excel/excel.html");

/**
 * EVERY EXCEL OBJECT-MODEL MEMBER THE COMPANION IS AUDITED TO TOUCH, with the
 * requirement set that introduced it. Versions are quoted member by member from
 * the `@remarks [Api set: ...]` annotations in
 * `node_modules/@types/office-js/index.d.ts` — the generated mirror of
 * Microsoft's API reference, and the only copy of it that is readable offline.
 *
 * `null` means "not a member of any ExcelApi set, so it cannot raise the floor":
 * `load`/`sync` are `OfficeExtension.ClientObject` / `ClientRequestContext`,
 * shared by every Office host, and `Excel.run` is the entry point to the
 * namespace itself — wherever `Excel` exists at all, ExcelApi 1.1 does.
 */
const AUDITED: Record<string, string | null> = {
  "Excel.run": null,
  "RequestContext.workbook": "1.1",
  "RequestContext.sync": null,
  "Workbook.getSelectedRange": "1.1",
  "Range.load": null,
  "Range.values": "1.1",
  "Range.address": "1.1",
};

/**
 * Re-derive the Excel surface from the source.
 *
 * The object model is reachable from exactly three named handles in this file —
 * the `Excel` namespace, the `context` `Excel.run` hands the callback, and the
 * `range` taken off it — so sweeping those three names finds every call. The
 * test below asserts the handles are still called that before trusting the
 * sweep; a rename fails loudly rather than returning a quietly empty set, which
 * is the failure mode that makes a check indistinguishable from a check that
 * never ran.
 */
function excelSurface(src: string): Set<string> {
  const found = new Set<string>();
  const sweep = (re: RegExp, owner: string) => {
    for (const m of src.matchAll(re)) found.add(`${owner}.${m[1]}`);
  };
  sweep(/\bExcel\.(\w+)\b/g, "Excel");
  sweep(/\bcontext\.(\w+)\b/g, "RequestContext");
  sweep(/\bcontext\.workbook\.(\w+)\b/g, "Workbook");
  sweep(/\brange\.(\w+)\b/g, "Range");
  // `load("values,address")` names properties fetched from the host that the
  // code may never mention again — they are part of the surface either way.
  for (const m of src.matchAll(/\.load\("([^"]*)"\)/g)) {
    for (const prop of m[1].split(",")) found.add(`Range.${prop.trim()}`);
  }
  return found;
}

describe("the Excel companion's ExcelApi floor", () => {
  it("still reaches the object model through the handles the sweep keys on", () => {
    // Guard the guard. If these names change, `excelSurface` goes blind and
    // every assertion below passes vacuously.
    expect(SRC, "`Excel.run` no longer takes a parameter named `context`").toContain("Excel.run(async (context)");
    expect(SRC, "the selected range is no longer bound to a const named `range`").toMatch(
      /\brange = context\.workbook\.getSelectedRange\(\)/,
    );
  });

  it("calls nothing outside the audited set — anywhere in src/excel", () => {
    // The sweep reads one file, so prove one file is all there is. A second
    // module reaching the object model would otherwise be invisible to it.
    const ts = readdirSync(fileURLToPath(new URL("../src/excel", import.meta.url))).filter((f) => f.endsWith(".ts"));
    expect(ts, "src/excel has grown a second module — re-derive the floor from it too").toEqual(["excel.ts"]);

    expect(
      [...excelSurface(SRC)].sort(),
      "the Excel API surface changed. The manifest's floor is derived from exactly this list — add the new " +
        "member to AUDITED with the `[Api set: ...]` line from @types/office-js, and raise the manifest if it " +
        "is higher than 1.1.",
    ).toEqual(Object.keys(AUDITED).sort());
  });

  it("pins the manifest floor to the highest set that surface needs", () => {
    const needed = Object.values(AUDITED).filter((v): v is string => v !== null);
    const highest = needed.sort((a, b) => {
      const [aj, an] = a.split(".").map(Number);
      const [bj, bn] = b.split(".").map(Number);
      return aj - bj || an - bn;
    })[needed.length - 1];
    expect(highest).toBe("1.1");

    for (const name of EXCEL_MANIFESTS) {
      const floor = /<Set\s+Name="ExcelApi"\s+MinVersion="([\d.]+)"/.exec(read(name))?.[1];
      expect(floor, `${name}: no ExcelApi requirement at all`).toBeDefined();
      expect(
        floor,
        `${name}: declares ExcelApi ${floor} but the code needs ${highest}. Both directions cost something — ` +
          `Partner Center derives the certified platform list from the Requirements block, so too LOW enlarges ` +
          `what Microsoft certifies and too HIGH locks out hosts the pane runs on. See the comment in ` +
          `manifest-excel.xml.`,
      ).toBe(highest);
    }
  });

  it("would notice a call that needs a newer set", () => {
    // The sweep is only worth running if it can still fail. `Shape.textFrame`
    // and `Worksheet.shapes` are ExcelApi 1.9 — the very APIs the old comment
    // claimed this add-in used — so this is the exact regression that comment
    // would have waved through.
    const mutated = SRC.replace(
      "range.load(",
      "context.workbook.worksheets.getActiveWorksheet().shapes;\n      range.load(",
    );
    expect([...excelSurface(mutated)]).toContain("Workbook.worksheets");
    expect([...excelSurface(mutated)].sort()).not.toEqual(Object.keys(AUDITED).sort());
  });
});

describe("the Excel companion's first-run callout", () => {
  /** The pane's primary button, read from the pane rather than assumed. */
  const button = /<button id="generate"[^>]*>([^<]+)<\/button>/.exec(HTML)?.[1].trim();

  const description = (xml: string) => /<bt:String id="GetStarted\.Description" DefaultValue="([^"]*)"/.exec(xml)?.[1];

  it("finds the pane's primary button to check against", () => {
    expect(button, "src/excel/excel.html no longer has a labelled `#generate` button").toBeTruthy();
  });

  it.each(EXCEL_MANIFESTS)("%s tells a first-time user something they can actually do", (name) => {
    const text = description(read(name));
    expect(text, `${name}: no GetStarted.Description at all — the callout would render empty`).toBeTruthy();

    // THE ORIGINAL DEFECT. The PowerPoint manifest may say "insert your first
    // chart" because PowerPoint inserts one; this add-in has no insert path,
    // which the ExcelApi surface above re-derives rather than assumes.
    expect(
      text!.toLowerCase(),
      `${name}: the callout promises an insert. The Excel companion never writes to the workbook — its ` +
        `whole Excel surface is reading the selected range (see the ExcelApi tests above).`,
    ).not.toMatch(/\binsert\b/);

    // And it has to name the control the user is being sent to. A callout that
    // describes a button by a name the button does not carry is the same defect
    // one rename later.
    expect(text, `${name}: the callout does not name the pane's button ("${button}")`).toContain(button!);

    // The config is inert until it reaches the other half of the product, so
    // the one instruction a new user needs is where to take it.
    expect(text, `${name}: the callout never says where the generated config goes`).toContain("PowerPoint");
  });

  it("would catch the string that shipped", () => {
    // Verbatim, as it stood until 2026-09-20 — proof the assertions above fire
    // on the real defect and not merely on a strawman.
    const shipped = "Find SSF Charts on the Home tab and insert your first chart.";
    expect(shipped.toLowerCase()).toMatch(/\binsert\b/);
    expect(shipped).not.toContain(button!);
    expect(shipped).not.toContain("PowerPoint");
  });
});
