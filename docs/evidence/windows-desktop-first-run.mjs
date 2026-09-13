#!/usr/bin/env node
/**
 * WHAT THIS POWERPOINT ON WINDOWS ACTUALLY SUPPORTS, AND WHAT A CHART LANDS AS.
 *
 * Every one of the 430 archived rounds is PowerPoint on the WEB — `platform`
 * reads `OfficeOnline` in all of them. Microsoft's validator says a submission
 * is tested on Windows and Mac as well, so the two claims the store listing
 * makes there had never been measured at all: that the host supports what the
 * product needs, and that a chart arrives as native shapes rather than a
 * picture.
 *
 * This takes both readings in one pass. It is not a round and writes nothing to
 * `rounds/`: a round is a battery against the web host, and this is a single
 * question put to a different platform.
 *
 * PRECONDITIONS, and the third one needs a person at the screen:
 *
 *   1. PowerPoint must have been launched with remote debugging, set
 *      PROCESS-SCOPED in the launching shell and never persisted:
 *
 *        $env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=9444"
 *        Start-Process "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" -ArgumentList '"<deck>"'
 *
 *      PowerPoint must be fully exited first — it reuses one process for every
 *      document, so a second deck open elsewhere silently keeps the old process
 *      and the flag is ignored.
 *
 *   2. The manifest must be in a trusted catalog. On AITEST that is
 *      `\\AITEST\OfficeAddins`, registered with `Flags=1`. Take it from the
 *      RELEASE asset rather than the working tree, so what is measured is what
 *      a user installs, and check the `<Id>` rather than the filename — three
 *      add-ins there ship a file called `manifest-prod.xml`.
 *
 *   3. **Someone has to add it once, by hand**: Home → Add-ins → More Add-ins →
 *      SHARED FOLDER → pick it → Add, then open the pane. There is no COM or
 *      command-line route to that, and the CDP route cannot bootstrap itself —
 *      the WebView only SPAWNS when the ribbon button is clicked, so there is
 *      no target to connect to until a human has clicked one.
 *
 * Run:  node docs/evidence/windows-desktop-first-run.mjs [--insert]
 *
 * `--insert` presses "Insert into slide" and reports what the pane says. Verify
 * what LANDED with COM rather than trusting the message — a picture and a group
 * both report success:
 *
 *     $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
 *     $s = $pp.ActivePresentation.Slides.Item(1)
 *     foreach ($sh in $s.Shapes) { "$($sh.Name) type=$($sh.Type) items=$($sh.GroupItems.Count)" }
 *
 * msoGroup is 6 and msoPicture is 13. A group named `PowerChart` is the answer
 * the listing promises; a picture is the fallback and means the host dropped a
 * mark the chart carries.
 */
import { chromium } from "playwright-core";

const CDP = process.env.SSF_CDP ?? "http://127.0.0.1:9444";
const PANE = "ssf-chart.struktureretsundfornuft.dk";

const browser = await chromium.connectOverCDP(CDP);
const pages = browser.contexts().flatMap((c) => c.pages());
const pane = pages.find((p) => p.url().includes(PANE));
if (!pane) {
  console.error(`No SSF Charts pane on ${CDP}. Targets:`);
  for (const p of pages) console.error(`  ${p.url().slice(0, 100)}`);
  console.error("Open the pane in PowerPoint — precondition 3 above.");
  process.exit(2);
}

const reading = await pane.evaluate(() => {
  const supported = (set) => {
    try {
      return Office?.context?.requirements?.isSetSupported("PowerPointApi", set);
    } catch (err) {
      return `threw: ${err.message}`;
    }
  };
  const sets = {};
  for (const s of ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"]) sets[s] = supported(s);
  return {
    host: Office?.context?.host ?? null,
    platform: Office?.context?.platform ?? null,
    version: Office?.context?.diagnostics?.version ?? null,
    sets,
    // The pane prints the deployed build; a reading against a stale pane is a
    // reading about the wrong code.
    paneBuild: document.body.innerText.match(/[0-9a-f]{7} · [^\n]*/)?.[0] ?? null,
  };
});
console.log(JSON.stringify(reading, null, 1));

if (process.argv.includes("--insert")) {
  const button = pane.locator('button:has-text("Insert into slide")').first();
  if (!(await button.isVisible())) {
    console.error("no `Insert into slide` button — is the Chart tab open?");
    process.exit(2);
  }
  await button.click();
  const deadline = Date.now() + 120_000;
  let last = "";
  while (Date.now() < deadline) {
    const said = await pane.evaluate(() =>
      (document.querySelector(".note, #note, [role=status]")?.innerText ?? "").trim(),
    );
    if (said && said !== last) {
      console.log(`  ${said.slice(0, 220)}`);
      last = said;
    }
    if (/inserted|added|scaled|failed|could not|error/i.test(said)) break;
    await new Promise((r) => setTimeout(r, 2000));
  }
  console.log("Now check what LANDED with COM — the message cannot tell a group from a picture.");
}

await browser.close();
