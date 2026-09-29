# SSF Charts — project memory

Open-source think-cell clone: a PowerPoint charting add-in whose charts are
**native, editable shapes** (never pictures), plus a Claude Agent Skill that
renders the same charts headlessly. 25 chart kinds, think-cell's signature
decorations, and "the good chart" design formalia are all implemented — see
the README feature table for the authoritative list.

## Architecture in one paragraph

`ChartConfig` (pure JSON) → `buildChart()` in `src/core` (pure TypeScript, no
Office imports) → renderer-agnostic scene graph (`rect | line | text | ellipse
| wedge | chevron | arrowhead | polygon | symbol` nodes, coordinates in points) → three
renderers: SVG (`src/render/svg.ts`, preview + tests), Office.js
(`src/render/powerpoint.ts`, live add-in; config persisted in a
`POWERCHART_CONFIG` shape tag for re-editing), and PptxgenJS
(`skill/scripts/render-pptx.mjs`, the skill's headless output). Details:
`docs/ARCHITECTURE.md`; the research that shaped it: `docs/RESEARCH.md`.

## Where things live

`test/README.md` maps the test suite; this is the source side.

| directory                       | what it owns                                                                                                                                                                                                                                                                                                                           |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/core/`                     | the engine, **zero Office imports**: `chart.ts` (`buildChart`), `layout/<kind>.ts` one per chart kind, `decor.ts`, `scene.ts` (the node contract **and** the three-renderer parity rules), `geometry.ts`, `format.ts`, `color.ts`, `collide.ts`, `samples.ts`                                                                          |
| `src/core/`, run-time reasoning | `placement.ts` (where a chart goes on a slide that already has content), `reconcile.ts` (what a run ACTUALLY produced, once the host stopped moving), `trace.ts` (the ordered record a run nobody watched leaves behind)                                                                                                               |
| `src/render/`                   | `svg.ts` (the reference renderer), `powerpoint.ts` (Office.js, the live add-in), `ooxml.ts` (post-processes a `.pptx` to carry the slot tags and groups pptxgenjs cannot write), `pptx-deck.ts` (builds a whole deck in the browser, handed over in one call), `host-probe.ts` (the answer sheet — what THIS PowerPoint actually does) |
| `src/taskpane/`                 | `app.ts` (the product's own pane), `harness-ui.ts` (the Testing panel and every helper only it uses — see below), `datasheet.ts`, `selftest.ts` (the in-host battery), `crashlog.ts` (the record that survives a run which never ends), `i18n.ts`, `templates.ts`                                                                      |
| `skill/scripts/`                | `render-pptx.mjs` + `pptx-paint.mjs`, the headless renderer. **Outside `tsconfig.include` (`["src", "test"]`), so never typechecked**                                                                                                                                                                                                  |
| `scripts/`                      | `triage.mjs`, `verify-deck.mjs`, `validate-ooxml.mjs`, `host-baseline.mjs`, `visible-charts.mjs`, `office-js-watch.mjs`, and the `build-*` set                                                                                                                                                                                         |

**`app.ts` and `harness-ui.ts`, and which one a change belongs in.** The pane
ships its own Testing panel, and until 2026-09-29 it was mixed into `app.ts` —
`wireInsert` was 1,369 lines wiring the product's controls and the `demo-*` /
`experiment-*` ones from the same brace. The panel and the ~1,890 lines of
helpers only it used now live in `harness-ui.ts`. Anything a USER presses is
`app.ts`; anything the round driver presses is `harness-ui.ts`.

`harness-ui.ts` takes its seven dependencies as an argument rather than importing
them, and that is structural: `app.ts` imports it, so importing back would close
an ESM cycle — which `tsc` and the whole suite would pass while the bundle handed
one side an undefined binding at init, inside PowerPoint. `test/import-cycles.test.ts`
is what notices. The same rule applies to anything else extracted from the pane.

**Function size is ratcheted.** `test/function-size.test.ts` records every
function in `src/` at or over 150 CODE lines (comments and blanks excluded,
measured from the TypeScript AST) and fails when one grows or a new one appears.
Re-record deliberately with
`UPDATE_FUNCTION_SIZES=1 node ./node_modules/vitest/vitest.mjs run test/function-size.test.ts`
— growth is allowed, hiding it is not. Note that raw lines and code lines rank
very differently here: `groupAndTagAll` is 827 lines and 267 of code.

**Adding a `SceneNode` kind: six seams fail loudly, one does not.** Four fail
at the SOURCE, as compile errors — `nodeToSvg` (`src/render/svg.ts`), `addNode`
(`src/render/powerpoint.ts`), the coordinate chain in `test/fuzz.test.ts`, and
`translateNodes` in `src/core/chart.ts`, which carries an explicit `never`
guard for exactly this.

Two more fail in CI rather than at the source, each through a
`Record<SceneNode["kind"], …>` in a test that cannot be filled in without
answering the question the seam poses. **This is a weaker guarantee and worth
knowing as one**: the build still succeeds, and a `--exclude` on the wrong file
takes the guard with it.

- `shiftNodeX` (`src/core/scene.ts`) — a duck-typed `Record<string, number>`
  cast over a fixed key list, so a new coordinate field is simply left
  unshifted. `points` was already missing when this moved out of
  `src/demo/demo.ts`, so a polygon stayed put while the scene around it shifted.
  `test/demo.test.ts` maps every kind to the coordinates a shift must move; each
  key is proven load-bearing by removing it.
- `makeAddNode` (`skill/scripts/pptx-paint.mjs`) — the headless pptx mapping,
  outside `tsconfig.include` and typechecked by nothing, with no `default` case
  (an unknown kind is deliberately ignored, not thrown on). A new kind therefore
  renders as NOTHING in the skill's .pptx, in a file that opens cleanly and is
  reported as a success. `test/pptx-paint.test.ts` asserts every kind draws
  something. **Not `render-pptx.mjs`**, which this list used to name and which
  holds no node mapping at all: somebody adding a kind would open it, find no
  `switch`, and conclude the seam did not exist.

Genuinely silent, one:

- `src/core/collide.ts` — text-only, matched by `MOVABLE` name prefixes, so a
  new label-like node is invisible to the de-collision pass. Not closed the same
  way on purpose: the match is on a runtime NAME, not on a kind, so there is no
  union to be exhaustive over — and which labels should move is a design
  decision per node, not something a map can answer.

## The lockstep rule (CI-enforced — do not skip)

Any feature change must update, in the same PR:

1. **Skill docs** — `skill/SKILL.md` + `skill/reference.md`
   (`test/skill-docs.test.ts` fails on missing kinds/rows/keys).
2. **Showcase deck** — extend `scripts/build-showcase.mjs`, run
   `npm run showcase`, commit the regenerated `examples/showcase.json` +
   `showcase.pptx` (`test/showcase.test.ts` + a CI byte-diff staleness gate).
3. **User manual** — `docs/MANUAL.md` (`test/manual.test.ts` fails on
   missing kinds/rows/controls/elements).
4. **README feature table.**

## Commands

```bash
npm run dev        # gallery + pane at localhost:3000
npm test           # full suite; npm run coverage enforces thresholds
npm run typecheck
npm run showcase   # regenerate the showcase deck (required after feature work)
npm run skill      # build skill-dist/ssf-charts.zip
npm run triage     # join a real run's deck + run log, and say where they differ
npm run host-diff  # line a probe answer sheet up against the fake
```

The suite's size lives in `test/fixtures/test-count.json` and is gated by
`scripts/test-count.mjs` — read it there. This line used to restate the number,
and drifted to claiming 1500 against a recorded floor of 2224: the same failure
the backlog and `UNSTABLE_ANSWERS` paragraphs below already warn about.

### Running the gate on the owner's Windows box

Environment only — CI on ubuntu is unaffected — but locally the commands above
fail in ways that read as code bugs:

- **`npx` is dead under PowerShell Constrained Language Mode.** `npx.ps1` does
  method invocation and throws
  `MethodInvocationNotSupportedInConstrainedLanguage` before reaching the tool.
  Call the entry point with `node`, or run it from git-bash.
- **AppLocker blocks an npm script that NESTS `npm run`** ("blocked by group
  policy"). That kills `showcase`, `skill`, `build:pages` and `visible-charts`
  — every one starts with `npm run build:lib`. Split them into their two flat
  halves. **Any new npm script must stay flat** or it is unrunnable there.

```bash
node ./node_modules/typescript/bin/tsc --noEmit                 # typecheck
node ./node_modules/vitest/vitest.mjs run                       # whole suite, nothing excluded
npm run build:lib && node scripts/build-showcase.mjs            # = npm run showcase
npm run build:lib && node scripts/build-skill.mjs               # = npm run skill
npm run build:lib && node scripts/visible-charts.mjs            # = npm run visible-charts
```

**The tool CLIs run there now, and until 2026-08-10 three of them did not.**
`test-count.mjs`, `flaky.mjs` and `visible-charts.mjs` each guarded `main()`
with `import.meta.url.endsWith(process.argv[1].split("/").pop())`, which never
splits a backslashed path — so every invocation printed nothing and exited 0.
Not a crash: a clean exit from the suite-shrink guard, the flake sweep and the
visual gate, which reads exactly like a pass. The predicate is shared now
(`scripts/is-main.mjs`) and `test/is-main.test.ts` drives BOTH platforms,
because the broken form works perfectly on the ubuntu runner and a same-platform
test could never have caught it. `visible-charts` also had no browser candidate
on Windows; it takes installed Chrome or Edge now.

### Looking at the task pane

The pane is the one surface the suite cannot judge: `dom-pane.test.ts` pins its
behaviour in jsdom, which has no layout and no colours, and the snapshot suite
never renders it at all. Look at it with `playwright-cli`
(`npm i -g @playwright/cli`), which drives the installed Chrome:

```bash
node ./node_modules/vite/bin/vite.js --port 3000 --strictPort &   # the pane's own dev server
playwright-cli open http://localhost:3000/src/taskpane/taskpane.html
playwright-cli click '.tabs .tab[data-tab="elements"]'
playwright-cli screenshot --filename=/tmp/pane.png
playwright-cli close-all
```

`.playwright/cli.config.json` sizes the window to **512×900**, which is the pane
at its real width — a chart picker judged at 1280 wide is three columns the host
will never show. `.playwright/dark.config.json` is the same window with
`colorScheme: "dark"`, and it is the only way to see dark mode here: headless
Chrome has no `prefers-color-scheme` switch, so the `@media` block never fires
and the pane looks permanently light. Give it a session of its own so both can
be open at once — `playwright-cli -s=dark open
--config=.playwright/dark.config.json <url>`.

Three things that will otherwise cost time. Config precedence is
`~/.playwright/cli.config.json` → this repo's `.playwright/` → env → flags, so a
personal global config is overridden here rather than fighting it. `--filename`
is relative to the **cwd**, not to `.playwright-cli/`, so a bare
`--filename=pane.png` drops a PNG in the repo root — write to a temp path. And
two console messages are expected outside the host and mean nothing: Office.js
warns it is `loaded outside of Office client`, and `/favicon.ico` 404s.

Two limits carry over from the jsdom harness. The pane localizes on
`Office.onReady`, which never fires in a browser, so `?lang=de` will not
translate — i18n stays a unit test. And nothing here reaches Office.js, so a
pane bug that only exists mid-`sync()` is still invisible; that is what the
real-host round is for.

**The whole suite runs there now — do not exclude anything.** This paragraph
used to say `test/skill.test.ts` could not pass and to run with
`--exclude '**/skill.test.ts'`, which was true and cost more than it looked: the
file that checks what the SKILL ships was gated by CI alone. Both reasons are
gone. It read a `.pptx` by interpolating an OS temp path into a `python3 -c`
string, where `\rings.pptx` is a carriage return, and `build-skill.mjs` zipped
through `python3 -m zipfile`, which on Windows is the Microsoft Store alias stub
and simply fails. Both now use `jszip`, already a dependency — so the zipper and
the reader are one library, no interpreter is involved, and the zip is
byte-identical run to run.

`node scripts/build-showcase.mjs` **alone renders with a stale engine** — it
imports `dist-lib/`, so without `build:lib` first you are diffing the showcase
against your own uncompiled change, and it reads as non-determinism.

## Working conventions (established with the repo owner)

- **Answer in caveman style** (the `caveman` skill, `full` level) from the first
  reply of every session in this repo, without being asked. Terse: drop
  articles and filler, fragments fine, keep every technical fact. Code,
  commits, PR bodies and user-facing docs stay normal prose — the style is for
  chat only. Drop it for security warnings, destructive-action confirmations,
  and any multi-step sequence where dropping conjunctions could be misread.
  "stop caveman" turns it off for that session.
- **Branch flow**: develop on the session's designated `claude/*` branch (the
  task prompt names it); after each merge, reset it onto `origin/main`
  (`git checkout -B <branch> origin/main`) — never stack on merged history.
  One PR per increment.
- **Auto-merge is authorized**: once CI is green on the exact pushed commit
  (verify `head_sha` matches local HEAD), merge the PR to main without asking.
- **Snapshots** (`test/snapshots.test.ts`) freeze every sample chart's SVG.
  Only update (`vitest -u`) after reviewing renders visually — screenshot via
  Playwright (`/opt/pw-browsers/chromium`) and inspect before accepting. CI now
  does the part a machine can: `npm run visible-charts` rasterises every sample
  in a real browser and fails on a chart that is drawn but not visible
  (white-on-white, collapsed to zero, off-canvas, one flat block, or missing over
  half its usual ink). Not a pixel diff — those numbers are measurements, and
  `--update` re-records `test/fixtures/chart-ink.json` after a deliberate change.
- **Visual QA is part of done**: render new features to SVG → PNG and look at
  them; several real bugs were only caught this way.
- **A regression test must be proven to fail without its fix.** Stash the
  source file, re-run, confirm the new test goes red, restore. A guard that
  passes against the pre-fix file is not a guard — it is decoration. Check
  _which_ assertion went red, too: a guard can fail for the wrong reason and
  still look proven (one here failed on a slide-swap rung that records nothing
  either way; another compared a `min()` both sides answered 0 for).
- **When two explanations fit the evidence, ask — do not reason.** A probe
  answer, a verdict, a log line can be about the host or about the thing that
  asked. Reasoning about which cost this project two full sheets and a session;
  adding the variant that separates them settled it in one run
  (`shape-add-fresh-slide-proxy` / `-held-` / `-positional-`, and again with
  `shapes-items-via-positional-slide`). Any answer that could be about the probe
  gets a partner question, and `KNOWN_DIVERGENCES` is where one waits while its
  re-run is pending. A probe can now carry its own partner: `Probe.follow` fires
  a second question in the same run when the first answer admits two readings,
  which is the rule automated — the reasoning was never the expensive part, the
  round trip was.
- **One answer sheet is not evidence about this host — it is evidence about
  this host in that minute.** Several probe questions give different answers on
  different runs of the SAME build, minutes apart —
  `shape-add-held-slide-proxy` and `shape-add-positional-slide-proxy` most
  sharply, and for a while in lockstep and always opposite. That is not a
  sequence of corrections where the newest value is the true one; it is a coin,
  and a majority across samples is not a mechanism either. A question that has
  been asked ONCE has not been answered, it has been sampled.

  **The sheet's YIELD is a sample too, and it is noisier than anything you will
  ship.** Four rounds on 2026-08-09 answered 15, 21, 18 and 24 of 28 — and the
  last two are the SAME BUILD, `1fa0509`, six questions apart. So a round that
  answers more than the one before it is not evidence that the change between
  them worked, and a round that answers fewer is not evidence that it broke:
  the swing on identical code is larger than any fix measured here so far. The
  probe's end-of-run second pass is the case in point — its real contribution is
  4 rescues in one round, well inside that noise. What DOES attribute is the
  per-question tag: `answered on a second pass at the end of the run` names the
  rung that produced the answer, one row at a time, and cannot be confused with
  the host having a good minute. Read those, not the total. **And read the build
  stamp first** — rounds 15 and 16 look like a before/after and are not one.

  The run log says why, and it is not subtle: `scratch slide landed but its id
will not resolve` several times mid-run, replacement scratch slides taken, and
  every question inside that window answering `no-scratch-slide` before the host
  came back. The host's ability to resolve a freshly added slide's id comes and
  goes within a single run.

  **A partner can be the coin while its trigger holds — 2026-08-12
  (`89675b6`).** `shape-add-held-slide-proxy-again` came off
  `PENDING_QUESTIONS` on `756682e` as stable across three passes, and this round
  flipped it inside ONE round (`threw` healthy, then `yes` in slide-trouble)
  while the TRIGGER answered `threw` both times and reported stable. This file
  names the trigger as the sharpest flipper on record; the pair has now been
  seen the other way round. It is in `UNSTABLE_ANSWERS` now — which a follow-up
  could not be until the shortlist invariant learned that a partner rides its
  trigger instead of carrying its own mark.

  **`UNSTABLE_ANSWERS` in `scripts/host-baseline.mjs` is the authoritative list,
  with every observation on every entry — do not restate its counts here.** A
  tally in this file and a list in that one drift the moment either changes, and
  this paragraph spent several rounds claiming "five sheets" while the table had
  six and eight; the same mistake the backlog paragraph further down already
  warns about. Read the table.

  What the table is FOR is stopping the next reader building on whichever answer
  a sheet happens to carry. `shape-add-positional-slide-proxy: yes` is exactly
  what would make a positional slide handle look like the safe way out of the
  by-id refusals, and it is what most of its samples say. It would not help
  anyway: `shapes-items-via-positional-slide` answers as short as the by-id form
  does in every run that put both, so a positional handle reads a shape
  collection no better.

- **The fake is gated against a real host in CI** — `test/host-contract.test.ts`
  diffs `FAKE_BASELINE` against the committed sheet in
  `test/fixtures/host-answers-web.json`. A new divergence fails there unless it
  is declared in `KNOWN_DIVERGENCES` with a reason. When a fresh sheet arrives:
  replace the fixture, run the suite, deal with what goes red. Do not edit the
  fixture by hand — it is a recording, not a preference.
  **A newer sheet is not automatically a better one.** The `a546897` round put
  only 18 of 27 questions, and swapping it in would have turned eight committed
  ANSWERS — `delete-then-lookup: reports-gone`, `group-reports-its-children:
threw`, `tags-on-fresh-shape: yes`, `untrack-available: no` among them — into
  `no-scratch-slide`, which the gate correctly reads as "unknown". That is
  deleting knowledge on the strength of a bad ten minutes, the same mistake the
  gate itself used to make one level up. Replace when the new sheet answers at
  least as much; otherwise keep the fixture, and put what genuinely moved into
  `UNSTABLE_ANSWERS` where it belongs.

  **The rule has now refused a sheet, so it is not theoretical.** The 2026-08-09
  evening round (`619d24b`) is the best-LOOKING sheet this project has taken —
  the probe's end-of-run second pass got it to 21 answers where the two rounds
  before it managed 15 and 13 — and it still must not be committed. It answers
  **21 of 28** against the fixture's **24 of 27**, and swapping it would turn
  `tags-on-fresh-shape: yes`, `group-reports-its-children: threw` and
  `group-of-existing-shape-readable: no-group-id` into `no-scratch-slide`, which
  the gate correctly reads as unknown. Count the ANSWERS, not how much better
  the round went.

  **Refused a second time on 2026-08-11** (`7027f96`, 23 of 28 against the
  committed 24 of 28): one ANSWER — `group-of-existing-shape-readable` — would
  have become unknown, and none would have changed. A one-question drop is
  inside this host's yield noise in both directions, which is exactly why the
  rule is arithmetic rather than judgement: "the newer sheet" and "the better
  sheet" are different things, and only one of them is checkable. Do the
  subtraction before swapping; it takes a minute and it has now said no twice.

  **The counts in this paragraph are historical and the fixture has moved since
  — read `test/fixtures/host-answers-web.json`'s own `build` and count its
  answers, never a number written here.** As of 2026-08-11 the committed sheet
  is `756682e`, answering 25 of 29. The round after it (`4feb5be`) answered 25
  of 29 as well, lost no ANSWER, and changed two —
  `shapes-items-via-positional-slide: not-listed → short-0` and
  `scratch-slides-returned: none → all`. That clears the bar and was still not
  swapped, deliberately: the count is identical, the first of the two changes is
  between two spellings of "the collection did not answer", and that same
  question CHANGED ITS ANSWER MID-ROUND — so committing its final value would be
  recording a coin-toss as a recording. Clearing the arithmetic makes a swap
  permitted, not obligatory.

  **SWAPPED AGAIN on 2026-08-12 to `1789749`, and this one the arithmetic did
  NOT decide — it is the first swap made on a judgement the rule does not
  cover.** The sheet answers 26 of 31 against the committed 26 of 30, so the
  count clears, but it is a ONE-FOR-ONE TRADE: it loses
  `binding-names-shape-later` and gains `untrack-available-on-shape`. Every
  earlier refusal cited "it would turn an ANSWER into unknown", and by that test
  alone this should have been refused too.

  What changed the reading is that the trade is not symmetric in
  REPRODUCIBILITY. `binding-names-shape-later` has been never-put for SEVEN
  consecutive rounds — the fixture was holding an answer no current sheet can
  re-earn — while `untrack-available-on-shape` answered in two consecutive
  rounds and could not reach the fixture while the old one stayed. A recording
  that cannot be reproduced still serves the gate, but it was blocking one that
  can. Both findings survive in prose either way: the bindings reading is in
  `KNOWN_DIVERGENCES` and is not lost by the swap.

  Read this as the precedent for a TRADE, not as the count rule loosening. A
  sheet that answers less is still refused, and it has been five times.

  **A swap to `957aca0` was WRITTEN UP HERE AND NEVER LANDED, and this paragraph
  claimed for a day that it had.** It was argued as the first sheet since these
  rules were written to make a swap obligatory rather than merely permitted —
  answering more than the sheet it replaced, losing no ANSWER, and finally
  putting `binding-names-shape-later`, which twelve attempts across nine rounds
  had never reached. The reasoning may well have been right. The commit is not
  there: the last change to the fixture is #419 (`3d984d9`), and nothing since
  has touched it.

  So the committed sheet is **`1789749`, answering 26 of 31**, and the five it
  never put are `binding-names-shape-later`, `grouped-child-by-id-from-slide`,
  `shape-resolve-held-slide-proxy`, `tag-on-group-survives` and
  `tags-add-same-key-twice` — including the very question the swap was supposed
  to bank. Count the fixture yourself before quoting either number; the round
  that would let this be redone is gone, so the arithmetic is all that is left.

  This is the failure mode the whole section warns about, committed by the
  section itself, one paragraph after telling the reader not to trust it: **read
  the fixture's own `build` and count its answers, never a sentence written
  here.** A swap is a commit, not a note.

- **The office-js tracker is swept weekly** by
  `.github/workflows/office-js-watch.yml`, which reports only issues touching
  APIs this repo calls that are not yet in `KNOWN_ISSUES`
  (`scripts/office-js-watch.mjs`). When one is triaged, add it to that table
  **with what was done about it** — including "no exposure", which records that
  somebody checked. Anything left out comes back next Monday. Five of the guards
  in this repo came from one manual sweep; this is so the next one is not luck.
- **Count the rounds before you write down what this host "usually" does.**
  `scripts/host-history.mjs <round*.json>` prints every question's answer per
  round, ignores the words that mean the run never put it, and reports the
  streak from the END. It also does the fixture-swap arithmetic this file asks
  for ("replace it when a new sheet answers at least as much") rather than
  leaving somebody to eyeball it. `UNSTABLE_ANSWERS` was written by hand from
  whatever rounds were open at the time and four of its five entries had gone
  stale by 2026-08-10 — one called a question "a coin" that had answered the
  same way ten rounds running, which tells the next reader they may not build on
  an answer that has in fact never moved.
- **"Changed its answer mid-round" is not the same as "a coin", and until
  2026-08-13 nothing could tell them apart.** The probe has stamped the host's
  regime (`healthy` / `slide-trouble` / `collection-refused`) on every sample
  since #390, and those stamps had only ever been read BY HAND, one entry at a
  time, in prose — the same way `UNSTABLE_ANSWERS` went stale.
  `scripts/host-regimes.mjs <round*.json>` reads them: for every question that
  moved within a round it says whether host state accounts for it, whether two
  flippers moved at the SAME pass boundary (one mechanism sampled twice, or
  two), and which regime the never-put questions were attempted in.

  **The verdict that matters is `untested`.** With three passes landing in three
  different regimes, "every regime maps to one answer" is true by construction —
  it cannot fail, so it is not evidence. A mapping is only reported as
  `explained` when some regime was sampled MORE THAN ONCE and agreed with
  itself. Without that split the tool would manufacture a finding every time a
  question flipped, which is the rasterise control's fixed-arm-order bug and the
  frame gate's invented exceptions for a third time.

  On its first run — against the committed fixture, so anyone can reproduce it —
  it reclassified two of that sheet's three flippers.
  `shape-add-positional-slide-proxy` reads `yes` in `slide-trouble` (twice,
  agreeing) and `not-listed` in `collection-refused`; `picture-then-shape-read`
  the same shape. Neither is a coin: they are degradations a caller can test
  for. `shapes-items-via-positional-slide` IS one — a single regime answered
  both `short-0` and `not-listed`, which no amount of regime sampling explains.

  **A verdict is about ONE round**, and the footer says so. `explained` is a
  reason to re-read an `UNSTABLE_ANSWERS` entry, not on its own a reason to
  rewrite it; the same verdict on a second round is what makes it a mechanism.
  `COIN` is the strong direction — it survives however many regimes were
  sampled, because one regime already answered two ways.

  **Both halves earned their keep on the first real round (`cd3b60c`,
  2026-08-13), and one of them by being WRONG.** `shape-add-positional-slide-proxy`
  read `explained` against the committed fixture and `coin` one round later, the
  same regime having produced both faces. So the mapping was one round's shape,
  and the rule that an `explained` verdict may not rewrite an entry is what
  stopped a day-old finding being written down as a mechanism. A tool whose
  caution is vindicated inside a day is worth the caution.

  **The pairing half ANSWERED a question that had been open in prose.**
  `shape-add-held-slide-proxy` and its partner `-again` came back LOCKSTEP —
  same pass boundary, three pairs, three agreements — so they are one mechanism
  sampled twice and a flip in either is a flip in both. That round also caught
  the pair the entry said could only be caught, never scheduled: `yes`/`yes`
  half a second apart on pass 2. Seven pairs on record now, seven agreements.

  **And the two verdicts together say more than either alone.** The pair agrees
  perfectly, so at any instant this host has a DEFINITE answer; and the same
  question is a `coin` by regime, because `collection-refused` produced both
  faces. There is a state, it changes during a run, and `regime` is not it. The
  open question is no longer "is it a coin" — it is "what is the state", with
  one candidate eliminated.

  **Two more candidates died on round 17's own numbers, before anything was
  built.** `threw` at 16.3s, `yes` at 33.9s, `threw` at 55.6s is non-monotonic,
  so neither ELAPSED TIME nor the PASS NUMBER is the variable — the "pass 1
  threw, later passes yes" split that R12 and R13 supported is broken by pass 3
  reverting. Read those three numbers before proposing either again.

  **The candidate left is the SCRATCH SLIDE, and every sample carries it now**
  (`ScratchState` in `src/render/host-probe.ts`): `first-slide`, `fresh-slide`,
  `reused-slide`, `no-slide`. `UNSTABLE_ANSWERS` already fingers it — this
  question WRECKS its own scratch slide every time it is asked, so pass 1 meets
  a deck with no scratch history and later passes do not. `host-regimes.mjs`
  reads it beside the regime and prints both, so a round where one says `coin`
  and the other `explained` is the answer arriving.

  **The stamp was tested RETROSPECTIVELY against round 17 before it ever ran,
  and it would have been blind there.** That question replaces its scratch slide
  0.2-0.5s before every ask (`replaced the scratch slide` at 16.1s, 33.7s and
  55.1s against samples at 16.3s, 33.9s and 55.6s), so all three samples are
  `fresh-slide` and the answers still disagree. Do not expect this field to
  answer on its own; expect it to be one of several.

  That retrospective is also what found a real defect in the READER. A stamp
  that never moves for a question produced `coin` — one bucket holding two
  answers — which is indistinguishable in the data from a genuine coin and means
  the opposite thing: `coin` is a fact about the HOST, `blind` is a fact about
  the STAMP. There is a fifth verdict now, and it fires on exactly this case.

  **And a saturation check, because this is the sixth time.**
  `stampSpread` reports what share of a round's samples carry a stamp's
  commonest value, and says so out loud past `SATURATED_AT`. On its first run it
  found `regime` at **88% `collection-refused`** in round 17 — worse than the 85%
  sticky flag `regimeFrom` was written to replace, and past the failure
  criterion that function's own docstring sets. A saturated field produces
  `untested` and `blind`, and both read as caution rather than as a broken
  instrument.

  **That is now fixed, and the fix is the one the docstring always claimed.**
  `regimeFrom` promised "the most RECENT thing this run watched it do" and
  implemented priority order, so one refusal painted every sample for twenty
  seconds however well the collection answered in between — and in round 17 the
  collection ANSWERED 14 of the 28 times it was asked, interleaved throughout. A
  later COLLECTION answer clears an earlier refusal now, which takes that round
  from 88% to 72% in simulation. Only a collection answer: `lastGoodAt` is set
  by almost every question, so recency across all three signals would saturate
  on `healthy` instead — the same defect mirrored, and there is a test holding
  that direction too.

  It does NOT explain the flip. Under the fix the pair reads `healthy`,
  `slide-trouble`, `slide-trouble` against `threw`, `yes`, `threw` — still a
  coin. The instrument got better; the question is still open.

  **And the round after that, reading the collection DIRECTLY says the regime
  model may be wrong at the root.** `collectionTimeline` prints every question
  that asks the shape collection, in time order, with what it said — the single
  most load-bearing fact about this host, reasoned about from batch timings and
  grouping traces for months and never simply READ.

  In time order round 17 looks exactly like the received story: the collection
  answers, refuses, comes back, refuses again, in three clean bursts. Grouped by
  QUESTION it is nothing of the kind. **Seven of its eight collection questions
  gave the SAME verdict on every pass** — `shapes-items-count-honest` refused all
  three times, `shapes-items-via-positional-slide` refused all three,
  `shape-add-fresh-slide-proxy` answered all three — and only
  `shape-add-positional-slide-proxy` ever varied. The interleaving is the fixed
  question order cycling through the passes.

  So on this round "the collection is refusing" is mostly a fact about WHICH
  QUESTION was asked recently, not about when. That also explains the 88%
  saturation without any appeal to host state: the always-refusing questions are
  asked every pass, so the flag is nearly always set. **Reads that list items
  fail; calls that add a shape answer** — which is a claim about two different
  operations, not about a host degrading over time.

  One round, and it should be confirmed on a second before anything is built on
  it — but note which way the caution runs. The existing model is the one with
  a round's evidence against it. `variesOverTime` is deliberately false unless
  some question gave BOTH verdicts: a bare "an answer came after a refusal" is
  satisfied by question order alone and reported recovery on a host that never
  changed, which is what the first version of this did before the by-question
  grouping caught it.

  Three properties of the stamp are worth knowing, because each is a mistake
  this file has recorded before. It is CATEGORICAL — an age or a counter gives
  every sample its own value, and "every value maps to one answer" is then true
  for any data at all, which is the `untested` shape. It is on EVERY sample, not
  only the odd ones. And every one of the five places that takes a scratch slide
  goes through `takeScratch`, held by a SOURCE SCAN in `test/host-probe.test.ts`
  — bypassing one leaves the counter stale and stamps `reused-slide` on a
  brand-new slide, and the behavioural test cannot see it, because `fresh-slide`
  still arrives from another path. Proven by doing it.

- **Dependabot's banner gets read, and the reading gets written down** in
  `docs/DEPENDENCY-ALERTS.md`. Same rule as the table above: an entry says what
  was decided, "no exposure" included. It sat at 9 alerts for weeks with nobody
  looking, which is the state a scanner is worth nothing in — a standing red
  count trains everyone to scroll past the one that matters. `npm audit fix`
  without `--force` cleared four; the rest are upstream pins with the analysis
  recorded. **Never take `npm audit`'s advice unread**: its proposed remedy for
  `image-size` is to move `pptxgenjs` back three major versions.
- **Stale documentation is a defect — fix it when you find it, in that turn.**
  Don't file it, don't mention it and move on. This applies to comments that
  justify a design with a claim that is no longer true, which is the expensive
  kind: `placement.ts` asserted that no requirement set exposes slide
  dimensions, and that single false sentence is why charts could only ever be
  placed vertically. If the fix is genuinely out of scope, say so explicitly
  rather than leaving the claim standing.
- **Test files are named by topic, never by increment** (`test/README.md` has
  the map). No `batch-N` / `bug-hunt-N` / `coverage-*` grab-bags: a test
  belongs with the thing it tests, not with the reason it was written.
- **This scenario has now produced FOUR false destruction claims, each by a
  different mechanism, and the pattern is the lesson.** `explode a degraded
picture` reports what became of a chart, and on this host every route to that
  answer is defeated by the host refusing to name things:

  1. `the slide went from 1 to 0` — an empty collection read believed as an
     empty slide (both signals agreed at zero).
  2. `the picture vanished while being exploded back` — a null update read as
     destruction when the host had merely refused the id.
  3. `the picture vanished while being redrawn` — the fix for (2) keyed on
     THROWN id refusals, and this host also fails to resolve a target quietly.
  4. `the picture is GONE from the slide` (round `1789749`) — the fix for (3)
     asked the slide and compared ids, and the id it compared with was one the
     host had refused a step earlier. The settle's `withId: 0` on that same
     chart is the tell: it had to find it by NAME. The deck inventory shows the
     slide holding one shape named `PowerChart` — the picture, untouched.

  What survives all four: **an id refusal anywhere in a scenario makes every id
  in it suspect**, so a shape the slide will not name is not a shape that is
  gone. Destruction is claimed only from a positive read on a scenario where
  this host refused nothing, and the refusal count is measured from the
  scenario's START because the refusal that poisons an id is routinely in an
  earlier step than the call that fails.

  Do not add a fifth mechanism. If a future round claims destruction, check the
  deck inventory before believing it — that is what caught all four.

- **Two style stores, and the DECK wins — decided 2026-08-18 when deck-level
  style shipped.** The imported style file and the saved templates live in
  `localStorage`, so they follow the BROWSER; a deck's style lives in a
  presentation-scoped custom XML part (`src/core/deck-style.ts`,
  `readDeckStyle` / `writeDeckStyle`), so it follows the FILE. When both exist
  the deck's wins, because the feature exists for the colleague who opens your
  deck — and that colleague is exactly the person whose browser holds a
  different style. Importing a style in the pane flips the session to the
  browser's copy (an explicit act should not be silently undone) and `Use deck
style` flips back. `resolveStyleFile` is the whole decision, and it never
  MERGES the two: a deck's palette under a browser's font is a style neither
  party chose.

  **Chart config stays in shape tags and must.** A config has to travel with its
  shape through copy/paste into another deck and through PowerPoint's own
  Duplicate Slide; a presentation-scoped part follows neither. The two stores
  answer different questions — "what is this shape" and "what does this deck
  look like".

  **The trap for anything that GENERATES a deck:** Office.js enumerates only
  parts related from `ppt/presentation.xml`. A part written at the package root
  (`/customXml/itemN.xml`) is invisible to the API, so a generated deck that
  drops one there produces a style no add-in can ever read. Nothing in
  `pptx-deck.ts` or `ooxml.ts` writes one today — the add-in's own write goes
  through `customXmlParts.add`, which relates it correctly — and if that
  changes, relate it from the presentation.

  **`getOnlyItemOrNullObject` REFUSES a namespace holding two parts**, it does
  not pick one, so the write deletes every part in the namespace before adding.
  A deck that quietly accumulated three styles would otherwise have no style at
  all, and the fake models the refusal so that stays true.

- **A chart cut and pasted on the WEB loses its config, permanently and
  silently — office-js#3784, triaged 2026-08-12.** Shape tags do not survive
  cut/paste on PowerPoint web (desktop keeps them), and every chart's config
  lives in a `POWERCHART_CONFIG` shape tag. So a user who cuts a chart and
  pastes it back has a chart that is no longer re-editable, with nothing in any
  log — the tag ANSWERED and said there is no config, which `tagsUnread` does
  not cover because that counts tags the host would not answer either way.

  **Deliberately not detected.** A count of shapes named `PowerChart` carrying
  no config would find them, and on this host it would be swamped by the tag
  writes the host itself refuses every round — the number would report "this
  host is unwell" rather than "your paste broke a chart". `same scale across the
deck` already says how many charts in a deck are re-editable, which is the
  same fact without the false precision. Recorded so the next person to meet an
  un-editable pasted chart does not go hunting in our tag writer.

- **The self-test headline counts OUR defects, not red scenarios.** Those are
  different numbers on this host, and conflating them made the summary useless
  for the only question worth asking of a series of rounds: is the add-in
  getting better? Most red here is the shape collection dying part-way through
  — `same scale across the deck` is largely a measurement of WHEN that happens
  — so a single "N of M passed" moved with the host's mood and never with the
  work. Across four rounds it read 10/11, 8/10, 9/11 while the defects that
  were actually ours were 0, 0, 0 and 1.

  **"Part-way through" is measured now, and it is not about position — it is
  **SUPERSEDED 2026-08-25: that 1% is now 36 of 36.** Kept because it is why the
  — see `scripts/claims.mjs`, claim `fresh-slides-group`, re-checked every round.
  code exists, not because it describes the host today — see
  `docs/WHAT-WE-KNOW.md`, checked every round by `scripts/claims.mjs`.

  about the SLIDE.** Pooled over 36 rounds (`npm run rounds`, _WHICH SLIDE THE
  CHART LANDED ON_):

      slide already had shapes  109 chart(s), 108 grouped = 99%
      freshly added, empty       92 chart(s),   1 grouped =  1%

  A chart drawn onto a slide this run has just added gets a short or empty
  pre-grouping re-read, so it is not grouped; an ungrouped chart's config tag
  then falls back to a `created` handle, which this host refuses about seven
  times in ten. That is the whole of `same scale`'s failure, and it is the
  freshly-added slide this repo has known about since #108 — not the tag writer,
  which four rounds of work aimed at before anyone joined `onSlide` to the
  grouping outcome. See `docs/BACKLOG.md`.

  **THE 1% IS HISTORY — corrected 2026-08-23.** The settled retry (rounds
  064/065) fixed it, and those charts group now: over the last 12 rounds it is
  20 of 20. The all-time figure still reads 66% because it is pooled over 39
  rounds that predate the fix, which is why `npm run rounds` now prints a recent
  window and a per-round sequence beside it — **quote those, never the pooled
  percentage.** The diagnosis above is kept because it is correct and it is what
  the fix was built from; only the rate is stale.

  Read the sequence rather than either percentage. Over the last eight rounds
  the fresh-slide population is empty — `[0/0 0/0 …]` — because the in-place
  update now handles those charts and a chart that is not redrawn never lands on
  a fresh slide. A percentage over an emptying population is the thing to watch
  for here.

  `scenarioBlame` is the split, and it is evidence rather than judgement: a
  failure is host-degraded only when the run recorded the host refusing
  something INSIDE that scenario (an id it would not resolve, an empty
  collection read, a GeneralException). Everything else is ours.

  **The default direction matters more than the split.** Unproven lands on US,
  because getting it backwards turns this into a way to make failures disappear,
  which is worse than not splitting at all — and there is a guard that fails if
  someone inverts it. Host-degraded failures are still NAMED in the line; they
  are not hidden, they are just not evidence about the product.

  Validated against a round whose answer was known before the rule existed. On
  `89675b6` the picture regression failed with `errors: 0, idRefusals: 0,
emptyReReads: 0` — a pure logic bug of ours — while `same scale` failed with
  six id refusals and four empty re-reads. Ours and the host's, separated
  correctly, from data recorded before anyone was looking for it.

- **Four sweeps run on a schedule, none of them gating a PR** — the office-js
  tracker (above), plus three in `.github/workflows/quality-sweep.yml`: a
  **flake hunt** (the suite three times under CPU load, reporting any test that
  disagreed with itself — `scripts/flaky.mjs` tells that apart from a suite that
  is red the same way every time), a **mutation run** over `src/core`
  (`npx stryker run`, scoped by `vitest.mutation.config.ts`), and an
  **install-path check** (`scripts/check-published-install.mjs`). Mutation is
  the rule below, automated: it answers "which assertions are decorative" for
  the whole engine at once, where the stash-and-re-run answers it for one. Not
  required checks on purpose — minutes-long jobs in front of every merge get
  switched off after the first bad week.
- **Every gate here reads the working tree; a user downloads the RELEASE.**
  Those two have diverged twice, and both times a fully green repo said nothing.
  v0.1.0 shipped the dev manifests while the README pointed at a
  `manifest-prod.xml` that was not in the release at all — twelve days, with
  `release.yml` sitting correct and un-run. Then v0.3.0 (2026-07-31) shipped
  `<Version>0.1.0</Version>`, the version Microsoft's validator rejects outright;
  #289 fixed the repo on 2026-08-06.

  **That one is SHIPPED — v0.3.1 went out on 2026-08-10 carrying it** (the bump
  commit says so in its subject). This paragraph said "the fix has reached
  nobody" for two days after it stopped being true, which is the failure the
  bullet is about happening to the bullet: a claim about the RELEASE, written
  from the working tree, and never re-checked. Read `list_releases` before
  repeating any sentence here about what has or has not shipped.

  The weekly check compares what is published against what is committed and says
  "cut a release" when they part. **After a fix that changes a manifest, ask the
  owner for a release** — merging it is not shipping it.

- **When moving tests between files, pin the total first and check it after.**
  A reorg once silently deleted 43 tests — the suite still went green because
  the count was never compared. `npx vitest run | grep "Tests "` before and
  after; if the number moved and you did not add or remove a case, stop. CI now
  holds the floor: `scripts/test-count.mjs` fails when the total drops below the
  recorded mark, which rises on its own. A deliberate drop is re-recorded with
  `--update`, so it lands in the diff where a reviewer sees it.
- **Releases**: merges to main refresh the rolling `skill-latest` prerelease.
  Versioned releases via the Release workflow's manual dispatch (the git proxy
  rejects tag pushes — dispatch creates the tag server-side).
- **Flag manifest re-installs to the owner**: the add-in is hosted on GitHub
  Pages (`ssf-chart.struktureretsundfornuft.dk`) and the sideloaded
  `manifest-prod.xml` only points at stable URLs — so code/pane/chart changes
  ship through `main` → Pages with **no** re-install. But when a change touches
  the **manifest itself** (ribbon buttons/menu items, `Permissions`,
  requirement sets, `DisplayName`/`Description`, icon references, or — never do
  this — the `<Id>` GUID), the owner MUST re-sideload the updated manifest in
  PowerPoint for it to take effect. Always tell the owner explicitly in that
  PR/turn: "⚠️ this needs a manifest re-install in PowerPoint" and say why.
  CI validates all four manifests with Microsoft's own tool (the `manifest` job,
  kept out of `test` because it calls a Microsoft service and `test` is the only
  required check). It found `<Version>0.1.0` on its first run — Office rejects
  anything below 1.0 — so the version is `1.0.0.0` and independent of the npm
  package version; `test/manifest.test.ts` pins that offline, along with the
  `<Id>` GUIDs and "no localhost in a prod manifest".

## Gotchas

- Office.js has **no freeform paths**: pies are triangle fans, radar/polygon
  fills degrade to outlines in the live add-in (the skill's pptx output gets
  real filled `custGeom` polygons), pattern fills are SVG-only (solid in PPT).
- **A chart's chrome is a fixed number of points, and a small frame cannot pay
  for it — so the plot goes NEGATIVE rather than small.** Every layout computes
  its plot by subtracting title, legend, axis and footnote from the frame, and
  each of those is priced in font sizes regardless of how much frame there is.
  Scatter at 120x90 — a thumbnail — computed `h: -8`. A negative height is not a
  squashed plot: `toY` maps the value domain through it, so the axis INVERTS
  (larger values plot downward) and the plot's own bottom edge lands below the
  chart. `fitPlot` (`src/core/layout/frame.ts`) is the floor, and every layout
  that builds its own plot rect goes through it.

  **Which way it grows is the load-bearing part.** It grows UP from the bottom
  edge the layout gave it, because that edge is the category axis and the value
  baseline — moving it moves what the chart claims — while everything above it
  is chrome. Anchoring to the frame's bottom instead was tried first and is
  measurably worse: it pins the plot to the foot of the frame and every label
  drawn beneath a mark spills out (14pt on a 120x90 bubble).

  The same shape recurs one level out, in four places that reserve room for a
  ring or band of labels and then floor the radius past the reservation — radar,
  sunburst, tilemap, pie. There the answer is not a clamp: when the reservation
  cannot be met the labels are DROPPED, because a label drawn off the chart is
  not there anyway and a floor that ignores its own reservation is the thing
  putting it there.

  **The pie is the one where the floor won, and it cost the whole chart.** Its
  side margin is a flat `fs * 7` — 70pt of a 120pt-wide frame — so a pie under
  ~140pt wide had nothing left and fell to its 1pt floor: a **2pt dot**, 0.1% of
  a thumbnail in ink against 38% at 200x150, with four labels drawn around it as
  though there were a chart there. Not an overflow, so no frame gate could see
  it; it was found by measuring what share of its own frame each kind covers
  across sizes, which is worth re-running after any layout change.

  **The first fix for it was all-or-nothing, and that is its own bug: a chart
  must not get SMALLER as its frame grows.** Taking the reservation in full the
  moment the frame could afford it collapsed the arc at the threshold — growing
  a pie from 160 to 170 points wide took its radius from 75 to 15, and a 280pt
  pie was no bigger than a 160pt one. Nothing overflowed and every gate was
  green; it is visible only in a SWEEP of one dimension, which is now
  `never shrinks when the frame grows` in `test/pie.test.ts`.

  The two axes do not admit the same answer, and that is the transferable part.
  **Horizontally the labels are SOFT** — they are clipped to the room they get —
  so the margin is simply capped at a share of the half-width and the arc grows
  smoothly; above ~280pt the flat `fs * 7` fits inside the cap and nothing moves.
  **Vertically they are HARD**: a label above the ring has nowhere to go and
  clipping its text does not make it shorter. What does is drawing it SMALLER,
  since the band is `outerFs * 2.2` — so the ring's own font size pays for the
  band, and below a 5pt floor the ring is dropped outright. That leaves one
  residual step of about 6 points on a frame ~66pt tall, which is inherent to
  having a floor at all and replaces one of 60.

  **That measurement is the one to reach for, because this failure is invisible
  to every other gate.** Nothing goes negative, nothing leaves the frame, the
  snapshots are green — the chrome simply eats the chart. It found three more
  the same day: a funnel whose gaps took 59 of a 120x90 chart's 64 points of
  plot, leaving five HAIRLINE bands at the 1pt floor (the split was "a point per
  band, gaps take the rest", which is backwards — the bands are the chart and the
  gaps are chrome for a label); and a butterfly whose two value strips and centre
  gutter took 84 of 120 points, leaving both sets of bars 36 between them. Both
  now scale the chrome into a budget the way `layoutGantt` already scales its
  three text gutters, and the butterfly's category names shrink WITH the gutter —
  they are centred in it, so a name wider than it is drawn across the bars it
  names.

  **A label is fitted to the MARK it sits on, not to the frame.** The pie's
  inside labels were the last ones fitted to nothing: drawn at the chart font in
  the middle of their slice, so a name wider than its own wedge ran across the
  neighbouring slices and, on a small frame, past the edge of the chart — where
  the frame clip cut it to an ellipsis and both the preview and the deck showed
  `mericas 38…`. The room a label in a wedge has is the CHORD of that wedge at
  the radius the label sits on, and `insideChord` is the one arithmetic the
  layout and its guard share, so the test cannot check a bound the layout never
  promised. Same shrink-together-then-clip as the funnel's rows and the
  butterfly's names; only slices big enough to get an inside label take part, so
  a thin one cannot drag the rest down.

  **Dropping a label inside a `forEach` is where this bites back.** The pie's
  slice loop advances its running `angle` at the END of the callback, so the
  `return` that skipped an outer label skipped the advance too and every slice
  after the first started at zero — a doughnut showing the wrong data at every
  small size, inside its frame, with the snapshots green because they are taken
  at one size. `test/pie.test.ts` pins the invariant that no frame or ink check
  can see: the slices TILE the circle, each starting where the last ended,
  covering 360° exactly once. Check what follows a guard before writing it, and
  prefer a condition on the block to an early return.

  **The FONT is the other axis a chart gets squeezed along, and the gate held it
  fixed for a week.** Every layout prices its chrome in font sizes, so a big font
  does what a small frame does. Seven overflows were sitting at 24 and 32pt, all
  one shape: a label centred on a row, a ring or a legend line, in a box
  `fontSize * 1.2` to `* 1.5` tall. Once the font outgrew the SPACING the labels
  overlapped each other at any frame size, and the last one left the chart at a
  small one. The funnel had been fixed for this; the butterfly, the gantt, the
  radar's ticks and the scatter's legend never had. Each is bounded by the space
  it actually has now — the row pitch, the ring gap, the frame — and the gate
  sweeps fonts as well as frames.

  **A shrink must move the box and the font TOGETHER, and the showcase caught me
  getting that wrong.** The radar's tick box is `fs * 1.2` for a font of
  `fs * 0.85`, deliberately taller than its text; collapsing both onto the new
  bound shifted every tick by 0.9pt on charts where nothing needed to shrink, and
  three showcase slides moved. Written as a RATIO instead — exactly 1 when the
  font is untouched — the geometry is byte-identical. Any "last resort" that
  changes an ordinary chart is not one, and the deck diff is what says so.

  **A label INSIDE the frame can still be unreadable, and no gate here could see
  that** — the frame sweeps pass a chart whose every label is stacked on its
  neighbour. Diffing the text ink boxes against EACH OTHER found 73
  kind/font/frame combinations with overlapping text, and one shape was 30 of
  them: adjacent CATEGORY AXIS labels, which is one defect rather than seven
  because that axis is shared by every cartesian kind. Fitting it to its slot
  took the total to 56, and four more fits (cascade, mekko, butterfly, radar)
  to 46. Two of those were defects at the DEFAULT font and size, which is the
  part worth remembering: the sweep is not only about extremes.

  Two things it taught. **A box wider than the mark it labels bleeds into the
  neighbour**: the mekko's was `width + 8` against a 2pt column gap, so every
  label could run 4pt into each side; the room is the mark plus its own share of
  the gap, so adjacent boxes ABUT. And **a label drawn below the plot must be
  reserved for in the plot**: the cascade's drop caption was drawn 7.5pt into the
  footnote's band at 480x300, and clamping it up instead stacks two captions at
  the same y where two thin blocks end together — which is a worse collision than
  the one it fixes.

  **The count is ZERO at the default font, in BOTH orientations, and the rest is
  at 18pt and above.** Fitting each label to the space it actually has — tick
  labels to their tick spacing, heatmap rows to their row, sunburst labels to
  their wedge, the cascade's in-bar lines to the gap between them — took the
  upright count from 237 to 76.

  **Then the same sweep was run SIDEWAYS and found the fits had only ever been
  written for the upright chart.** `horizontalChrome` drew its category names at
  the chart font in a box `fs * 1.5` tall centred on the row; the mekko's own fit
  answered `false` outright when horizontal, so the whole thing was skipped; the
  totals, the waterfall's value labels and the combo line's point labels were
  bounded by nothing at all. One shared bound — the ROW PITCH — closed nine
  kinds at once, and the combo's point labels alone were 222 of the pairs.
  Rotating a chart rotates which side of a label is crowded; it does not excuse
  the label from being fitted, and for a year nothing measured the rotated case.

  Do not quote a total from this paragraph — measure it. The sweep is ink
  boxes against each other over kinds x orientations x fonts x frames, and its
  answer moves with the ranges you give it. What is stable is the SHAPE of what
  is left: point labels against axis chrome in scatter and bubble (decided, see
  below), and charts whose chrome simply exceeds the frame at 26pt and up.

  **A fit needs a FLOOR as well as a bound, and none of them had one.** Every
  axis fit shrinks a label to the room it has and answers whatever the
  arithmetic says — six y ticks 1.6pt apart on a 200x150 scatter at a 26pt font
  produced six ONE-POINT labels, ink no reader can resolve, from a fit that
  reported success. Below `MIN_TICK_FS` the labels are dropped and the gridlines
  kept, which is the answer the radar, sunburst, tilemap and pie reservations
  already gave when their band could not be met.

  **A CLAMP and a FLIP are not the same move, and the difference decided a case
  this file had already recorded as closed.** Flooring the CAGR caption at the
  title's bottom was tried and reverted: it turned five `title x cagr-label`
  overlaps into EIGHT against the column totals, because a floor moves the label
  whether or not the destination is free. The de-collision pass can now flip a
  label BELOW its mark when the way up is blocked, and that is worth 13 pairs
  with nothing traded — because a flip only takes a position already clear of
  everything settled, and the totals rank ahead of the caption and settle first.
  So it cannot buy its way off the title by landing on a total; it stays put.
  Same idea, opposite result, and the revert was still right about the clamp.

  Restricted to labels anchored to a single mark (`FLIPPABLE` in `collide.ts`).
  A series label may NOT go down: moving it crosses the label beneath it and the
  two then name each other's lines — every label on the canvas, each one wrong,
  which is the failure the upward-only rule was written for in the first place.

  **Two attempted fixes were REVERTED because the measurement said they made
  things worse, and that is the useful part.** Giving the scatter's label placer
  the axis labels as obstacles removes 35 pairs and makes it DROP point labels it
  can no longer position — on a comfortable 480x300 chart, because the y axis
  owns the left margin: a point's label is data and a tick label is chrome.
  Confining that placer's band to the plot — which is the same trade by another
  route, since the band deliberately overhangs into the tick strip — takes the
  count for those two kinds from 889 to 599 and drops 56 of 301 point labels.
  Same verdict, and the reason is recorded at the call site so the next person
  to notice the overhang finds out it was measured. And a third
  nearly went the same way on a misread: the combo render LOOKED like a total had
  landed on a line marker, and a text-versus-mark sweep says it has not. That
  sweep found nothing at first because it matched no marker at all — a combo's
  are `rect` nodes named `combo-marker`, not symbols — which is the vacuous
  measurement this file keeps warning about, caught only by asking what it
  matched.

  **The overlap half of that gate was checking TWO of the eight frames, and its
  name did not say so.** The overflow sweep runs 80x60 through 960x540; the
  overlap sweep ran 200x150 and 480x300, and asserted "no chart overlaps its own
  text at the default font". Sweeping the other six found **146 overlapping pairs
  at the DEFAULT font** — 54 at 80x60, 36 at 60x300, 29 at 300x60, 27 across
  120x90 and 160x120, none at 960x540. A gate whose name is wider than its
  coverage is the `batch committed` mistake in a test.

  Fifty of those are the scatter/bubble trade this file already records as
  measured and refused twice, and it is DECLARED in the gate now
  (`acceptedTrade`) rather than avoided by not looking: a point label may touch
  an axis TICK, only those two kinds, only against a numbered point label. Every
  other pair fails. Stating an exception and dodging the frames that would reveal
  it are not the same thing.

  **A second sweep took the gate to 60x300, and found the same defect class six
  more times.** A tall sidebar carried 16 pairs and an 80x60 thumbnail 52. What
  they share is a band drawn without being paid for: the funnel's conversion
  rate at a fixed font in a gap that collapses, the heatmap's column headers and
  the butterfly's series names at the chart font in boxes a few points wide, the
  combo's horizontal point label in a fixed 60pt box clamped to `cfg.width - 60`
  (x=0 on a 60pt chart, straight across the category names), and a bubble size
  legend whose numbers are wider than the circles they label.

  Two are worth knowing beyond their own layout. **`fitPlot` may overrun a band
  the layout reserved above the plot** — it grows UP from the bottom edge, which
  is the baseline and may not move, so on a frame that cannot pay for its chrome
  the plot rises back through its own reservation. The cascade's group headers
  ended up under the bars that way, and the answer is the one every reservation
  here already gives: chrome that cannot be paid for is not drawn. **And a
  conservative bound must stay conservative at the small end** — the scatter
  legend reserved rows using `max(40, cfg.width - axisW - 8)`, and at 80pt wide
  that floor made the bound WIDER than the walk gets, so it counted fewer rows
  than it drew and the legend wrapped into the x-axis strip. Flooring a WIDTH up
  reserves less, not more.

  **The reservation and the draw share ONE predicate now, in the one place they
  had disagreed twice.** `horizontalLegendFits` is read by
  `computeFrameHorizontal`, by `horizontalChrome` and — since 2026-08-13 — by
  the horizontal MEKKO, which draws its own legend and was the one call site the
  shared-predicate change missed. It gated on `seriesLabels && series.length > 1`
  alone, so on a frame the predicate refuses it drew a three-row legend into a
  reservation of ZERO rows: four overlapping pairs at 120x90 rotated, two at
  80x60, one at 300x60, every one of them inside the frame where no overflow gate
  could see it. **A shared predicate is only shared once every call site asks it**
  — extracting one and wiring two of three leaves the third in the state this
  paragraph already names as the worst of the three options. Gating it cleared
  every rotated frame to zero, so 80x60, 120x90 and 300x60 all joined the rotated
  sweep, which is now WIDER than the upright one. The band that is
  reserved and the legend that is drawn cannot come apart. They had: the mekko
  drew at a widened `frame.x` the reservation never counted for, and a scatter
  fix that gated only the reservation left the legend drawn over a band reserved
  as zero — which MEASURED WORSE than not gating at all, 55 overlapping pairs to 63. Reserving nothing and drawing anyway is the worst of the three options,
  and it is the one a half-fix produces. A shared function is the only version
  of "these two agree" that cannot rot.

  **300x60 and 80x60 are CLOSED, and each needed a BOUND rather than the
  de-collision work this paragraph predicted for two rounds.** The prediction was
  wrong three times in a row — the sunburst, then these — and the pattern is
  worth keeping: "the labels are already movable, so `collide.ts` has looked and
  found nowhere" is a statement about POSITION, and the answer was always SIZE.

  - Upright **column totals** were the one family never fitted to anything:
    centred on their category slot at the full chart font, so `"113"` at 17.4pt
    sat in a 13pt slot and ran into its neighbour. The horizontal branch had
    bounded itself against its row pitch for a year; the upright one had not.
  - The combo line's **series name** shares the right margin with the column
    series names and was taking a width FLOOR (`fs * 3.4`) even where a real
    gutter existed — fitted to that inflated width it held the chart font while
    its neighbours shrank, and was drawn across them.
  - The **CAGR caption** was the last label in the engine still drawn at the full
    chart font, in a FIXED 90pt box that hung off the left edge of any narrower
    chart (x = -17 at 80x60, -24 at 60x300). It is fitted to the gap between the
    title's INK and the arrow now, and dropped past the floor — keeping the
    ARROW, which carries the anchors and the slope where the caption is only a
    caption.

  **Clamping the CAGR caption to the title's bottom is still refused**, and this
  did not overturn it: a clamp moves a label whether or not the destination is
  free, which is how it turned five `title x cagr-label` overlaps into eight
  against the totals. Shrink-then-drop is not a clamp.

  **The floor on the combo's series name is LOAD-BEARING and removing it broke
  the showcase.** The measurement that said it only ever bound at 80x60 was taken
  on `sampleConfig("combo")` alone; the showcase's combo slide has a plot that
  reaches the right edge — exactly the case the floor was written for — and
  dropping the floor removed "Margin %" from the deck. It is taken only when the
  real gutter cannot hold a readable name now. **Measure a layout claim on more
  than the sample before acting on it**: the deck diff is what caught this, and
  nothing else would have.

  Three of the remaining shapes are fixed and the gate now covers 160x120:
  the tilemap's two scale ends (each owns half a colour bar a few points wide and
  their ink met in the middle), the mekko's legend (drawn at the widened
  `frame.x` while `computeFrameHorizontal` had counted its rows for the
  pre-widening x — the reservation and `legendRow` disagreed, which that
  function's own comment says they must not, so it wrapped onto rows nobody had
  reserved and landed on the bars), and the clustered in-bar labels (fitted with
  `textWidth <= across + 2`, so ink stood a point proud of the mark on each side
  while the neighbouring bar's label did the same — the mekko's `width + 8`
  lesson, unlearned one layout over).

  **120x90 is CLOSED, and the prediction in this paragraph was wrong about what
  it would take.** It said the sunburst's adjacent outside labels needed
  `collide.ts` vertical de-confliction rather than another bound. They did not,
  and `collide.ts` was the one place the fix could NOT go: those labels are named
  `label-N`, the same name scatter and bubble use for POINT labels, and moving
  those is a trade this file already records as measured and refused twice. A
  `MOVABLE` entry would have taken the refused trade by accident.

  What it needed was the bound the arc fit is not. Each label was sized by its
  OWN wedge's arc, which says nothing about where the NEIGHBOUR's midpoint falls
  — a wide wedge beside a narrow one earns a tall label and still sits close to
  it. Every outside label is on one circle (`r + fs*0.7`), so once they are all
  placed the vertical gap between two of them is a fact the layout can simply
  read. Sizing both to `gap / INK_RATIO` makes their two half-heights sum to
  exactly the gap.

  Two details worth keeping. **Past the floor it drops ONE of the pair, not
  both** — dropping both loses a label the survivor's room could have carried,
  and the smaller font is the narrower wedge, i.e. the smaller share of the data.
  And **the sides are independent**: a left label and a right label never meet,
  so pooling them would shrink labels with no neighbour near them.

  The general lesson is the one the pie's inside labels already taught, one step
  further out: **a label is fitted to the mark it sits on, and where marks are
  adjacent, to the NEIGHBOUR as well.** The arc fit was necessary and not
  sufficient.

  **The rotated sweep is now WIDER than the upright one.** It was four pairs at
  120x90 rotated, all one shape — a mekko's legend against its totals and its
  category names — which turned out not to be a fit written for the upright
  chart at all, but the legend's shared predicate having a call site nobody
  wired. Gating it took every rotated frame to zero, so the rotated list carries
  80x60, 120x90 and 300x60 while the upright one stops at 120x90. Do not expect
  the two lists to match: rotating a chart rotates which side of a label is
  crowded, and what is left upright at those two frames is de-collision work on
  the stacked and combo kinds.

  `test/frame-fit.test.ts` is the standing gate: nothing a chart draws leaves
  its own box, over every kind × eight frame sizes × seven fonts, plus no chart
  overlapping its own text at the default font in EITHER orientation, at every
  one of those frames — 60x300, 80x60, 120x90, 160x120, 200x150, 300x60 and
  480x300. The overlap sweep and the overflow sweep finally cover the same list.

  **And BOTH swept `sampleConfig(kind)`, so every decoration a sample does not
  switch on was swept by nothing — on a product whose selling point is
  think-cell's annotations.** Third time this file has recorded a gate whose name
  is wider than its coverage, and this one was the widest. Turning each
  decoration on in turn found five defect families, one of them at 480x300:

  - a **difference arrow** on a kind with no category axis. Seven kinds — pie,
    doughnut, funnel, sunburst, tilemap, treemap, waffle — publish a PLACEHOLDER
    anchor, every category at the middle of the plot with `categoryWidth` the
    whole plot, so that decorations which only need somewhere to POINT still
    work. An arrow between two of them has the same "here" and "there": it drew
    a zero-length arrow at `plot centre + plot width / 2 + 10` — ten points past
    the right edge — labelled `-5…` by the frame clip. Not drawn now. **One arrow
    legitimately shares a category with itself** and the suite said so before it
    shipped: a difference arrow anchored at a VALUE LINE compares one column with
    that line. That is why the CAGR arrow and the difference arrow ask slightly
    different questions instead of sharing a predicate.
  - a **callout** whose text is wider than the chart. Its clamp answers the
    CENTRE of the frame when the box cannot fit between its own bounds — written
    for the vertical case — so an oversized bubble hung off both edges at once,
    18.9pt each side on twelve kinds. Shrink-then-drop, like every other label.
  - a **callout on the TITLE.** The same clamp floors at `h / 2`, and y=0 is
    where the title is: on a 100% chart, a mekko, a treemap or a sunburst —
    which publish no `valueToY`, so every callout anchors on the plot ceiling —
    the bubble landed on the chart's own name. The ceiling is `titleInkBottom`
    now. Fourth appearance of "a clamp that keeps a label on the canvas is not
    automatically a fix — ask what is at the destination".
  - a **quadrant label** taking `Math.max(20, zone - 8)` — a floor that RAISES a
    width, which is the scatter legend's own bug in another layout. Fitted to its
    zone and dropped past the floor.
  - a **band label** priced in font sizes with nothing bounding it, drawn below
    the foot of a squeezed plot. Bounded by the FRAME rather than by the band,
    deliberately: a band can legitimately be a sliver, and fitting the label to
    that would drop the name of every thin band where today it merely spills over
    its own edge onto the data it names.
  - a **lollipop's value label**, the one horizontal label that never got the
    row-pitch bound the other nine kinds were given.

  The decoration sweep is two fonts rather than seven, because what it found were
  bounds missing outright rather than bounds that bit at a size — and the font
  axis is already swept in full.

  **The TOP-LEVEL options were the third hole, and the largest defect in this
  repo's layout code was sitting in it.** `axisBreak`, `multiples`, `pareto`,
  `pie.semi`, `pie.explode`, `logScale`, `secondaryAxis` and the rest are neither
  a decoration nor a font, so nothing swept them either. Seven more families, two
  of them at ordinary sizes:

  - the SEMI-doughnut gauge, three ways: its labels went through a clamp that
    cannot place a box wider than the frame (60.6pt off a 60x300 chart, the
    callout's bug in another layout), its leader lines ran past the arc's own
    edge, and its headline total hung below the chart;
  - an EXPLODED slice moves `r * 0.08` along its midline and takes its outside
    label with it, past a band sized for the un-exploded ring — 9.6pt under a
    960x540 doughnut. The offset comes out of the RADIUS now, so the moved slice
    lands back inside the envelope the un-exploded one had;
  - `axisBreak`'s zigzag overhangs the plot by a fixed 4 points, which is more
    than the plot's own inset on a small chart;
  - `pareto`'s secondary axis labels are centred on their ticks, with nothing
    stopping the topmost one leaving the chart — the per-tick `printsOnTitle`
    answer the primary axis already gives;
  - `multiples` builds one CHART PER SERIES at `(height - chrome) / rows`, so a
    titled 300x60 chart asks for a 150x25 panel and then a 7.6pt one. Two general
    bounds came out of that: a TITLE is now bounded by the frame's height as well
    as its width (its ink reaches `fontSize * 1.21` down from y=0, so a chart
    shorter than that draws its own name off the bottom), and the category strip
    yields at the FOOT of the chart the same way it already yielded to the title;
  - a `footnote` or an explicit `scale` squeezes the plot until a line's or a
    combo's markers cross the frame — **a mark is drawn AROUND its position**,
    and only scatter and bubble had ever been given that. `markInFrame` is the
    answer for a marker that sits on a LINE: moving it would detach it from its
    own series, so it gives up size instead and is dropped below a point of it.

  **The wedge measurement had to be fixed before any of this could be read.**
  `inkBox` bounded a wedge by the disc that contains it, so a semi-doughnut —
  which draws half a circle — reported 291pt of overflow past the bottom of a
  960x540 chart where the ink stops at the centre line. A conservative bound does
  not merely miss defects, it INVENTS them, which this file already records for
  the arrowhead and for the text-width estimate. Read the arc, not the disc.

  **And the CAGR arrow's bounds are now a clamp between two edges rather than a
  floor at zero**, which is worth stating because the first version of it was
  wrong in a way the suite could not see: it dropped the arrow with a bare
  `return`, from inside `decorationNodes`, which builds ONE list for every
  decoration — so the difference arrow, the callouts and the bands went with it.
  A condition on the block, never an early return; the pie's slice loop already
  carries that lesson one file over.

  **The DATA is the third axis, and it found one thing — which is worth as much
  as a long list.** Every sweep here holds `sampleConfig`'s own data: four to six
  short category names, one to three series, two-digit values. Transforming it
  (long names, 24 categories, 10 series, values in the billions, all-negative,
  tiny fractions) turned up a single overflow family: a FUNNEL with 24 stages on
  a 60pt chart drew its last stage 10.4pt below the chart, because `Math.max(1,
…)` floored each band and 24 floored bands plus their gaps measured 36 points of
  a 24-point plot. The floor comes out of the GAP now — the bands are the chart,
  the gaps are chrome for a label, which is what that layout's own comment
  already said — and where even hairlines will not fit, the floor is abandoned
  rather than the stages.

  **The one defect that arrives with the DATA rather than with the frame** is
  worth its own line, because no sweep over sizes could ever have found it. The
  scatter's x and y tick strips were fitted by one rule, `gap >= fs * 1.4`, and
  only one of them is about a line height: down the Y the labels stack, across
  the X they sit side by side, so what the gap has to hold there is the WIDEST
  LABEL. `1,234,567,890` is 60 points wide at the default font, sails through a
  14-point test, and is drawn straight through its neighbour. Fitted by width
  now. Ordinary two-digit samples hide it completely.

  **The overlap half of the decoration sweep was measured and NOT gated**, which
  is the more interesting half. Thirty-three pairs, and after the callout fix
  every one that remains is a decoration doing its job: a callout is an OPAQUE
  bubble deliberately placed over the chart (the ink test cannot see a filled
  box), a band label sits over the data it annotates, and the rest are the
  scatter/bubble point-label trade this file already records as measured and
  refused twice. A gate asserting zero there would be asserting that decorations
  do not overlap the chart, which is what they are for.

  **That sentence was ASPIRATIONAL until 2026-08-18, and the cross-product it
  described was the one thing nobody had run.** The overflow sweep asked eight
  frames at the default font and seven fonts at THREE frames; a large font on a
  SMALL frame — the corner where every overflow of the preceding three PRs had
  been living — was never asked. Giving the font sweep the full frame list turned
  up **42 nodes drawing outside their own chart**, over ten kinds, none of them
  at the default font. All 42 are closed, and the shapes were the ones this file
  already names: a band priced in font sizes on a frame that cannot pay for it
  (heatmap and butterfly headers, the heatmap's gradient legend, the mekko's
  upright totals, the bubble size key), and a label hung above its own mark with
  nothing over it but the edge of the chart (combo, waterfall, boxplot).

  Two of them are general enough to state on their own.

  **An empty text node is not nothing.** `clipTextToFrame` used to answer `""`
  for a label with no room and leave the node in place — and the only labels it
  can empty are the ones whose anchor is ALREADY outside the frame, so the pass
  that exists to keep ink inside the chart was the last thing putting a shape
  outside it. Layouts do the same one level up: their own fits clip a category
  name to nothing on a narrow chart, 203 such nodes across the kind/frame/font
  sweep. Every one is a text box PowerPoint has to create, on a host whose draw
  cost grows with the shapes already on the slide. They are dropped now, and
  `emits no text node the clip has emptied` holds it.

  That drop also caught a test asserting nothing: `a sunburst keeps its INSIDE
labels when the outer ring will not fit` passed on three nodes carrying the
  empty string. The node existed; no label was drawn. It went red the moment
  empties were dropped, which is the only reason anyone found out — the same
  vacuous-measurement shape this file records for the rasterise control and the
  text-versus-mark sweep.

  **The OVERLAP half was widened the same day, to the fonts either side of the
  default, and it found 98 pairs over eleven kinds.** Same shape as the overflow
  sweep: the gate swept seven frames at ONE font while its name claimed the
  default font generally. 6, 8, 14 and 18pt are in it now — the range a deck
  actually uses — and 24 and 32pt deliberately are not (282 pairs, and at that
  size the chrome genuinely exceeds the frame, which is a decision about how
  small a chart SSF Charts claims to draw rather than another bound).

  **Half of the 98 were ONE mechanism, and it is the mirror of the marker
  clamp.** A value drawn above a mark — a column total, a waterfall value, a
  boxplot median, a mekko total, a combo point label — was CLAMPED to y=0 when
  its mark reached the top of the plot. y=0 is where the TITLE is, so the clamp
  did not fix the defect, it moved it onto the one label that names the chart.
  `aboveMarkFontSize` in `layout/frame.ts` is the shared answer: the band between
  the title and the mark, and no label where that band cannot carry one. A clamp
  that keeps a label on the canvas is not automatically a fix — ask what is at
  the destination.

  The rest were bounds that measured the wrong thing, each worth one line:

  - the funnel's row labels stopped shrinking at 6pt whatever the band held, so
    a floor that ignores its own reservation put every row's name through its
    neighbours' — dropped past `MIN_LABEL_FS` now, like every other fit here;
  - the heatmap's column headers were anchored to `titleH` while `fitPlot` moved
    the grid, so on a short frame the headers sat below their own columns and on
    top of the legend. A header names the column under it and has to move with
    it;
  - the radar's tick fit divided `r` by the RING COUNT, and the rings drawn are
    the ticks above the minimum — so the outermost radius is not one of them and
    the "gap" was a quarter wider than the rings actually are;
  - the bubble size key was pushed off one edge to clear the other (title, then
    x-axis strip) until it was asked to fit INSIDE the plot or not be drawn, and
    it clears the group legend only where the two actually meet — clearing the
    legend's whole BAND moved the key on an ordinary 480x300 chart, which the
    snapshot caught;
  - the horizontal combo's series name was the last label drawn at the full
    chart font while every number around it had been fitted to its row. It takes
    half a row now, because `collide.ts` has nowhere to move it sideways: every
    row is occupied.

  **24 and 32pt joined the gate on 2026-08-19, and the entry that deferred them
  was wrong about what they held.** It said the remaining pairs were a chart
  with more chrome than room, and therefore a `MIN_READABLE` decision rather
  than a bound. Measured: **185 pairs, 165 of them on ONE frame** (300x60), and
  the mechanism was neither. `fitPlot` grows a squeezed plot UP from its bottom
  edge, and the category strip, the axis ticks and the line's own name are all
  placed FROM that edge — so they climb into the TITLE together. The category
  axis alone asks for `fs * 1.5 + 3`, which is 51 of the 60 points a 300x60
  chart has at 24pt.

  `printsOnTitle` is the shared answer, and what it protects is the title:
  chrome that cannot be paid for is not drawn, and the title is the one label
  that says what the reader is looking at. Per-tick where the labels are
  independent numbers, per-strip where they share a y. No ratio was needed, and
  the `MIN_READABLE` question is still unasked — which is the better outcome,
  because it was never a question about readability, it was a plot growing
  through its own chrome.

  Two smaller corrections came with it. The combo line's sideways name was
  bounded to half a row, and half a row is not the room it has: every row
  carries a category name and a total CENTRED on it, so the clear air is the
  pitch less one of those labels, and the name is aimed at the boundary between
  two rows rather than just under its point. And `uprightNameY`'s last-resort
  clamp was removed in favour of drawing nothing — its destination on a frame
  that short is the title, and a series name that reads as part of the title
  names nothing.

  **A gate can be widened into a test that was asserting nothing.** `drops a
point label the name has taken the room from` measured 120x90, where the name
  is now dropped outright — so the trade it names cannot happen there and the
  test passed on an empty premise. It measures 160x120 now. Same failure mode as
  the sunburst assertion that passed on empty text nodes, one round later.

  **A mark is drawn AROUND its position**, so a point at the very edge of a plot
  puts half its marker outside it — harmless while chrome sits over the plot, and
  not harmless once `fitPlot` brings the plot's edge up against the frame's. The
  showcase's own overlap-relief slide had a bubble hanging 11.7pt past the right
  edge of its box at 520x300, a perfectly ordinary size, and the deck diff is
  what found it. Scatter and bubble hold a mark far enough inside the frame for
  its GLYPH to fit (a star reaches 1.67x its data radius, so the bound is the
  drawn extent, not `r`). It is a clamp and it moves data — the same trade
  `fitPlot` makes for the plot itself — and it is the identity for any mark
  already inside the chart, which is asserted rather than assumed.

  It measures
  INK, not boxes — a first version measured boxes and produced four false
  positives and one false negative in one run, because a label's box is routinely
  wider than its text and is anchored by `align`.

  **Measure with the metric the layouts measure with.** A scratch sweep using a
  characters × font-size estimate reported two scatter point labels sitting in
  the x-axis strip at the default font, and the gate written from it asserted
  those two as known exceptions. The engine's own `textWidth` says they clear —
  so the exceptions were fiction and the gate now asserts zero. An approximate
  metric does not merely miss defects, it invents them, and an invented one gets
  written into a test as a permanent allowance.

- **A shape the user selected must be let go of before anything is drawn.** On
  the web, `addTextBox` DELETES the selected shape (office-js#2775) and a picture
  cannot be inserted while one is selected (#3698) — and the insert path reads
  the selection's bounds precisely so it can place the chart there. Reading the
  selection and holding it are different things; `dropShapeSelection` runs
  between them. Ordering is the whole protection, and the guard asserts it.
- **Anything that calls `setSelectedShapes` must run after the ladder.** That
  call wedges the web host's whole selection subsystem, so the ladder
  (`which selection call wedges the host`) has to be the FIRST such call in a
  run — that, and not being alone in a run, is what lets it be routine. Two
  orderings were tried and are wrong, both recorded in `docs/PUBLISHING.md`:
  last (the wedge happens before the ladder asks) and third (the ladder wedges
  six scenarios instead of two). The test states it as the property, because
  adjacency to one named scenario was only ever a proxy for it.
- **A slide id captured early in a run may not be the id the deck answers to
  later — and that is why the probe's clean-up leaves slides behind.** The
  2026-08-11 round (`96461eb`) reported `0 of 45 scratch slide(s) deleted; 45
left in the deck (the deletes reported 45 but the deck only shrank by 0)`.
  Every one of the 45 `deleteSlideById` calls returned true and the deck went
  from 53 slides to 53. The route is `deleteSlideByPosition`'s
  `indexOf(id) < 0`, which reads "not in the deck's list" as "already gone" —
  sound only while the id we hold and the ids the deck lists are the same
  strings. In that round the scratch ids read `4123571115#123571113` while the
  deck listed `256#109857222` through `314#195537992`, and BOTH come from the
  same `slideIds()` projection minutes apart.

  **ANSWERED 2026-08-11 (`756682e`): `stillListed` came back ZERO of 62.** The
  measurement was built to separate two readings — the ids are stable and the
  deletes fail, or the id we hold is not the id the deck answers to — and it
  chose the second, unambiguously. Delete-by-id on a slide this run added is not
  failing, it is **structurally impossible**: `indexOf(id) < 0` is true for every
  one of them, so every delete takes the "already gone" branch and removes
  nothing. Both id lists come from the SAME `slideIds()` projection minutes
  apart — the add captured its id from it, the clean-up reads it back — which is
  what makes this a fact about the host rather than about two readers.

  `deleteSlideByPosition` now returns **false** for an unfindable id rather than
  true. That claims less and deletes no more, which is the only safe direction
  on a call that removes slides from someone's presentation; for an id that
  genuinely never existed it is a shade pessimistic, and an under-count costs a
  line in a report where an over-count costs sixty blank slides.

  **The repair is positional, and it is written now** (authorised 2026-08-11).
  The probe's slides are appended by `slides.add()`, so they are the last N in
  the deck and need no id at all. `deleteTrailingSlides` is the hands;
  `positionalSweepPlan` is the decision, kept pure and away from any host call
  because on this path the decision IS the safety. It returns a plan only when
  both deck counts are known, the deck actually grew, and the count is at most
  the smaller of "what this run added, less what already went back" and "how
  much the deck grew" — which together guarantee the first index to delete is at
  or after the deck's size when the run started. Everything below that index was
  the user's before the run began, and that floor is asserted a second time
  rather than trusted to the arithmetic. Deletion runs highest index first so
  removing one cannot shift another, and the sweep fires ONLY when delete-by-id
  left something behind, so a host whose ids work never reaches it.

  `test/host-probe.test.ts` proves each clamp is load-bearing: with the "no more
  than the deck grew" clamp removed a plan reaches `from: -54`, i.e. into the
  user's own slides, and the guard names it.

  **IT WORKS, and a clamp fired on its first outing** (2026-08-11, `c792072`).
  The deck went from 1 slide to 71 during the probe and back to 3:
  `sweeping the run's own slides by position {from: 3, count: 68, deckAtStart: 1,
deckNow: 71}` then `swept: 68`. The round left the owner **8 slides, 2 of them
  empty**, against 70 and 61 the round before.

  The two it left are the interesting part. The deck grew by SEVENTY while the
  run could only account for sixty-eight scratch ids — `addScratchSlide` returns
  null when an add lands but cannot be claimed — so the "never more than this
  run added" clamp held the count at 68 and left the other two alone. That is
  the clamp doing exactly its job on live data: two blank slides is the price of
  never deleting a slide this run cannot prove it created, and it is the right
  price.

  `stillListed: 0 of 68` again, so the by-id finding reproduces and the sweep is
  not papering over an intermittent fault — it is the only mechanism that works
  here.

  **Clean on 2026-08-12 (`957aca0`): `69 of 69 scratch slide(s) deleted … 69
removed by a positional sweep after delete-by-id took none`, and the deck it
  left holds SEVEN slides, none of them blank.** No leftovers at all this time,
  where the round before left two — and the two were never a flaw in the clamp,
  only the price of an add the run could not claim, which this round did not
  have. `scratch-slides-returned` therefore answered `all` and came OFF
  `KNOWN_DIVERGENCES`, exactly as that entry's own text said it should ("this
  entry stays until the clean-up is fixed"). The clean-up is fixed.

- **`getItemOrNullObject` is not the last word on whether a slide exists.**
  PowerPoint on the web resolved a freshly-added slide's id once and refused it
  ever after, while still listing that id in `slides.load("items/id")`. So
  "the host will not resolve it" ≠ "it is gone": `deleteSlideById` read it that
  way and reported clean-ups it had not done, and a probe run left fourteen
  blank slides in the deck. The deck's own id list is the stronger question —
  use it to identify a slide (`addScratchSlide` diffs it) and to confirm a
  delete. Any new code that acts on a slide id needs the same doubt.
- **A slide handle is good for ONE sync when the slide was just added.**
  Resolving a `getItemOrNullObject(id)` proxy is what makes Office.js rewrite
  its object path to `getItem(id)`, and a freshly-added slide's id does not
  round-trip through `getItem` on the web — so the handle that just passed its
  liveness check answers `GeneralException`
  (`errorLocation: SlideCollection.getItem`) the next time it is used. The
  `SlideThunk` comment says this for `getItemAt` and it is just as true by id:
  re-acquire per sync-batch, never hold one across a `context.sync()`. It cost
  the host probe eight of its fourteen questions (each one recorded as a host
  divergence that had never been asked) and `slideImageBase64` the self-test's
  whole visibility scenario. Pre-existing slides are unaffected — their ids
  round-trip — which is why editing a chart in place has always worked.
  **Confirmed head-on 2026-08-08, and it is not about the handle at all.** A
  FRESH `slides.getItem(id)` — nothing held, nothing resolved earlier — answered
  `threw` (`GeneralException`) for a slide added moments before, while its
  follow-up partner asked the same call about a pre-existing slide in the same
  run and got `yes`. So the rule is not "do not hold a slide handle"; it is
  **`getItem(id)` does not work on a NEW slide, by any route**. That is the call
  `getTargetSlide` makes. The partner is `Probe.follow` doing its job: one run,
  two readings separated, no reasoning.
- **An update writes only what changed, when it can prove that is the same
  thing as redrawing.** The add-in used to delete every shape and add every
  shape back for any edit at all, and on the web that is ~50 seconds for a
  24-shape chart. The diffs say almost none of it was needed: a retitle changes
  ONE node of twenty-four, a single edited data point changes two, a full
  rescale eighteen.

  `src/core/scene-diff.ts` is the decision, pure and testable without a
  PowerPoint. It answers null to everything it is not sure of — the frame must
  be identical, and the node count, order, kind and name at every index, because
  shapes are found POSITIONALLY (anchor is node 0, the parts tag lists the rest
  in drawing order) and a structural difference would write a bar's geometry
  onto a label. Only `rect` and `text` may be written to: they are the two kinds
  whose whole appearance is a closed set of property writes, where a wedge is a
  fan of rotated triangles and an arrowhead a rotated triangle, both baked at
  creation with no freeform path to edit.

  **`CHART_SCENE_TAG` is what makes it sound rather than merely fast.** The
  update rebuilds the OLD scene from the config stored on the chart, and "what
  that config renders to today" is "what is on that slide" only while the engine
  has not changed. Ship a nudged label offset and the diff would call those
  nodes unchanged, skip them, and leave the stale rendering there FOREVER —
  where redraw-everything repairs it on the next edit. So a fingerprint of the
  scene as drawn is written with the config and checked before any diff is
  trusted. A consequence worth knowing: the first edit of every chart already in
  a deck is a full redraw that stamps the fingerprint, and only later edits are
  fast.

  **It applies to UNGROUPED charts only, which is the case that matters.** A
  grouped chart's shapes are inside the group and its parts tag does not list
  them, so there is no node-to-shape mapping at all — and this host ungroups
  every chart it cannot group, which is where the fifty seconds live. A healthy
  host keeps its groups and keeps redrawing, which it can afford.

  Every refusal falls through to the redraw, and the fast path runs BEFORE the
  delete, so a refusal costs time and nothing else. The tag writes ride in the
  same sync as the property writes deliberately: a chart whose picture is new
  and whose config is old would silently revert the user's edit next time they
  opened it, and that is the one outcome worse than being slow.

  `applyNodeInPlace` must set every property the adders set, and a source scan
  in `test/office-render.test.ts` holds them in lockstep — it names the
  forgotten property, because a missing line is a chart that draws right and
  edits wrong with nothing in any log to say so.

  **It fired ZERO times on its first real round (`b7e183d`), and the round could
  not say why — the SIXTH failure-only field in this repo, written by the same
  session that wrote the rule down.** There was no success line and no refusal
  line, which is indistinguishable from the code not being there; the reason had
  to be reasoned out of grouping traces and a deck inventory instead of read.
  It declines with `not updating in place — redrawing instead` and a `why` now,
  because the reasons are not interchangeable: a grouped chart has no
  node-to-shape mapping BY DESIGN and never will, a missing fingerprint means an
  older build drew it and this very redraw fixes it, and a refused id readback is
  a fact about the host. Three different next steps behind one silent `return
false`.

  **ANSWERED on the next round (`53ec985`), in words, unanimously: ELEVEN
  refusals and every one of them `the chart has no parts list, so its nodes
cannot be mapped to shapes`.** Not one other reason in the round — no
  fingerprint mismatch, no refused shape, no plan the differ declined. So the
  inference above is confirmed and can be stated: on this host the fast path's
  precondition is unavailable in BOTH branches. A grouped chart's shapes are
  inside the group and its parts tag does not list them; an ungrouped chart's
  parts tag needs an id readback this host refuses. The in-place update is for
  healthier hosts, and here it costs one cheap check per chart and nothing else.

  That is not a reason to remove it. It is correct, it is guarded, and the
  moment a chart arrives with a parts list — a desktop host, or this one on a
  good day — an edit stops costing fifty seconds. What it is a reason for is
  not counting on it in any reasoning about THIS host's numbers.

  **A PICTURE is not in the scene, so the scene may not decide the update —
  found 2026-08-12 (`89675b6`), and it is a regression this fast path
  introduced.** `render: "image"` on a config does not produce a picture; the
  renderer takes that path only when handed `pictureBase64`. So collapsing a
  chart to a picture builds the SAME scene the chart already has: the differ
  compared 24 nodes to 24 identical ones, answered "nothing changed", wrote
  nothing, and reported success — `updated only the shapes that changed
{changed: 0, of: 24}`, with the slide still holding its 24 native shapes.

  The self-test is what caught it (`the collapse added 0 shapes — the slide went
from 24 to 24`), but the cost is not the self-test's. The auto-picture fallback
  is what the add-in reaches for when this host has ALREADY failed to draw
  shapes, so the one path that exists to rescue a struggling host was the path
  being skipped, silently, with success reported to the user.
  `tryInPlaceUpdate` refuses a picture update outright rather than learning to
  handle one: it writes a closed set of `rect` and `text` properties, and a
  picture fill is neither.

  **The guard had to move before it meant anything.** Written against the fake's
  ordinary insert it passed with the fix removed — the fake groups, so the fast
  path was already refusing a step earlier for want of a parts list, and the new
  refusal was never reached. It lives with the other fast-path tests now, on the
  `drawLoose` harness that produces the ungrouped chart this path accepts. Third
  time this session that a guard needed its setup fixed before it could fail.

- **Two obvious ways round the id refusals are already closed, and both were
  checked against Microsoft's own surface on 2026-08-13 rather than guessed at.**
  Anyone meeting `InvalidParam passed to GetItem(id)` reaches for one of these
  within a minute, so they are written down with what killed them.

  - **Looking the shape up by NAME instead.** Not available. PowerPoint's
    `ShapeCollection.getItem(key: string)` is documented in the office-js
    typings as _"Gets a shape using its unique ID"_, `@param key The ID of the
shape` — ID only, on both the slide's collection and `ShapeScopedCollection`.
    The `docs/powerpoint/shapes.md` article shows `slide.shapes.getItem("StatusCard")`
    and reads as a name lookup, which is what makes this worth a line: the
    ARTICLE is the misleading source and the typings are the authority. Read
    `node_modules/@types/office-js` before building on a docs example.
  - **Reading the shapes back through the SELECTION** (`getSelectedShapes()`,
    which returns a `ShapeScopedCollection` and is a different host path from
    the slide's collection). Corroborated broken on the web by two open issues:
    office-js#3083 (`setSelectedShapes([])` does not unselect on web, though it
    does on desktop) and #3698 (after that call, `getSelectedShapes()` never
    resolves its `context.sync()`). Both are already in `KNOWN_ISSUES`; what is
    new is that they close this as an ALTERNATIVE rather than only as a hazard.
    It would also mean selecting the user's shapes to read our own, which the
    selection gotcha above already forbids.

  And a third that this repo measured before anyone else published it:
  office-js#5022's commenter works around a hung sync with _"a 1-2 second delay…
  though this sometimes still fails"_. That is the same remedy office-js#2903
  proposes for fresh slides, and the scratch-slide gotcha below records what it
  cost here — 1 of 25 questions answered against 19 of 26. External agreement
  that the delay is unreliable; local measurement that it is actively worse.

- **A chart the drawing context could not tag is not finished.** On the web the
  tag write goes through a shape proxy several syncs old and the host refuses it
  (`InvalidParam passed to GetItem(id)`, 46 times in one 38-item run), leaving a
  chart on the slide with no config — visibly a chart, and not re-editable. The
  demo path always survived it because `insertDemoDeck` re-reads the settled
  deck and plans a `retag`; the ordinary insert and update paths had no such
  pass, which is what `same scale across the deck` was reporting as "3 of 8
  charts carry the shared scale". `settleAndTagChart` gives them the same
  second chance from a fresh context. A settled write is not the "retry against
  a host that just dropped a sync" that gave this project duplicate slides —
  the distinction is a re-read, and it is the one recovery this host honours.
  **The second chance is not a guarantee**, and the deck from 2026-08-06 is the
  counter-example: four charts redrawn to a new scale, ungrouped and carrying no
  config at all, after the settle had had its turn. The fake had no host on
  which the settle also fails — `faults.refuseTagWrites` is now that host, and
  it had to be armed on BOTH tag writers, because the refreshed-proxy path is
  the one the settle uses and a fault on the other only ever refuses the write
  the settle repairs. Which of the two failed on the real host is not yet
  known: `settleUntaggedCharts` traces `{charts, settled, lost}`, so the run log
  says, and the deck alone does not. **Answered on 2026-08-06: neither — the
  settle was never invoked.** A run with FIVE `tagging failed` events carried no
  `settled the config tag` trace at all, and that trace is unconditional, so its
  absence proves the pass never ran. The hole was `updateChartsInSlides`
  returning the caller's old target _bare_ when group-and-tag produced no target
  at all: `lost` stayed undefined, the `no-config` filter matched nothing, and
  the settle returned before it could trace. "Never asked" and "asked and
  failed" look identical from a deck and want different fixes — the trace is
  what separates them, so never reason from a deck about which one happened.
  `targetWithNoTagResult` is that decision, extracted so it can be checked
  without a PowerPoint. **And on 2026-08-08 the next question was answered too:
  the settle now RUNS and still loses.** Five `settle pass:` lines in one round,
  every one of them `settled: 0, lost: 1`, each preceded by `the settle's
re-read came back empty`. So the fall-through to a collection read does not
  rescue a chart the run has just drawn — it meets the same empty collection
  everything else does. That is not in tension with the 23 retags an earlier run
  landed by collection read: those were pre-existing slides read by the repair
  pass, and this is a slide the run drew seconds ago. The distinction to hold on
  to is not by-id versus collection, it is **fresh versus settled**.

  **Partly overturned on 2026-08-09 by round 8: the settle CAN rescue a
  freshly-drawn chart, and did.** One of five settles in that round reported
  `settled: 1, lost: 0` — the first success on record. What separates it from
  the four that failed is not freshness: it is whether the chart was GROUPED.
  A grouped chart gives the settle a shape id to write through; an ungrouped one
  leaves it a collection read, and the collection is what this host will not
  answer. So the rule is narrower than "the settle does not rescue fresh
  charts": it rescues the ones grouping survived for, which on a degrading host
  is the chart or two either side of the flip. The
  messages carry their outcome now — they all begin `settle pass:` so absence
  still reads as "never invoked" — because for one round the log said
  "settled the config tag…" five times while its own numbers said nothing was
  settled.

- **Do NOT wait after adding a slide — it was tried and it cost 18 of 19 probe
  answers.** office-js#2903 says a slide added on Online is unusable for a
  couple of seconds and its reporter's fix is to wait; `addScratchSlide` did
  that on 2026-08-07 and the next round answered **1 of 25** questions against
  19 of 26 the build before. The add landed, the wait ran, and the liveness
  check after it found nothing, so every question came back `no-scratch-slide`.
  This host is not the host that issue describes: it resolves a fresh slide's
  id ONCE and refuses it ever after, so waiting spends the one resolution later
  rather than buying time. `web-host.test.ts` guards against reintroducing it.
- **Two `slides.add()` calls in quick succession kill PowerPoint on the web.**
  Named on 2026-08-08 by the step tracing added the same day, in one round:
  `adding the first scratch slide` at 33.2s, `adding the second scratch slide`
  at 33.6s, and the tab was gone. The first add SURVIVED — the second step is
  only written if it did — and nothing after the second was reached, which
  exonerates the ninety-six shapes the scenario was about to draw. Four tenths
  of a second apart is all it took.
  Read with the rasterise gotcha below, the pattern is one thing: whatever this
  host does after `slides.add()` it does slowly and badly, and a second
  operation arriving before it settles takes the process with it. Rasterising a
  fresh slide killed five rounds; adding a second slide killed two more.
  **Nothing in this repo takes a scratch slide any more.** Both scenarios that
  did have stopped, and `test/selftest.test.ts` asserts it for the whole battery
  by slide count, so a new one cannot quietly reintroduce it. The host probe
  still takes them — it has no alternative, every question is about a slide it
  owns — which is why it churns through so many and why its rounds are the ones
  that see `no-scratch-slide` windows.
- **Only an ID may cross a sync — and a proxy's PARENT counts.** A shape proxy
  carries its parent's object path, so members from a re-read collection and
  members from `created` are equally poisoned once Office.js has rewritten that
  slide handle to `slides.getItem(id)`. Handing either to `addGroup` does not
  merely fail to group: it THROWS, and the throw takes the batch's tagging with
  it, so the chart loses its group AND its config. Five charts in one run.
  `chooseGroupMembers` is that decision — group by re-resolved ids, or group
  nothing — and "group nothing" is better because the THROW costs every chart in
  the batch where declining costs one. Not because declining is cheap: the
  archive says an ungrouped chart loses its config 66% of the time, which is why
  a partial match is now grouped one level up rather than declined. **It worked
  on 2026-08-07**
  (`not grouping: no member handle this host will accept index=0 refreshed=0`)
  and the chart lost its config anyway — for the reason in the next bullet, not
  this one.
- **The host may refuse to name a SHAPE by id, and the recovery is a collection
  read.** Sixty-six errors in the 2026-08-07 run log, every one `InvalidParam
passed to GetItem(id)`, code 5010, at `errorLocation: ShapeCollection.getItem`
  — the slide answered, the shape did not. It took the drawing context's tag
  write, the readback, `ungroupedFallback`'s id read, **and
  `settleAndTagChart`'s own fresh-context write**, which is exactly what "the
  update reported 5×no-config" was. The settle then gave up, on the reasoning
  that a collection search "would only find a DIFFERENT shape to put this
  chart's config on" — sound with no id, wrong with one, because the read loads
  `items/id` and the caller's id picks its own shape out of the answer.
  `settleByCollectionRead` is that fall-through, and a collection read is the
  pattern this host DOES honour: 23 retags landed that way in the run that lost
  46 tag writes. `faults.refuseShapeById` is the fake being this host — and it
  must be paired with a refused in-context tag write, or the settle never runs
  and the guard passes against the unfixed file, which is how the first version
  of that test proved nothing.
- **A printed `getItem` does NOT mean a held handle — ASKED AND ANSWERED.**
  Office.js rewrites a resolved `getItemOrNullObject` proxy's path to
  `getItem(id)`, so for months every log read as though this code were holding
  handles across syncs. It was not. With `extendedErrorLogging` on, the host
  annotates the call each path was CREATED by, and the answer is unambiguous:
  `var slide = slides.getItem("282#543504795") /* originally
getItemOrNullObject("282#543504795") */`. The rewrite is just how a resolved
  null-object proxy prints. Do not re-derive this from an excerpt; the
  annotation is the evidence, and it took one build to get after two sessions
  of reasoning that got nowhere. The same annotation is what identifies a
  genuinely poisoned proxy: `shapes.getItem("27") /* originally addTextBox(...)
*/` is a `created` proxy, and those the host really does refuse.
- **When the excerpt cannot settle it, make the host say more.** Every
  `debugInfo` before 2026-08-07 ends `"fullStatements":["Please enable
config.extendedErrorLogging to see full statements."]`, so all a reader got was
  `surroundingStatements`, an excerpt — and the excerpt could not tell a held
  handle from a rewritten one, which is the question above. Turning the flag on
  answered it in a single round. `enableExtendedErrorLogging` is the
  two-explanations rule applied to the tooling: stop reasoning, make the host
  say. `trimDebugInfo` keeps BOTH ends of the statement list so the file stays
  sendable — tail-only was the first version and it was wrong, because
  `surroundingStatements` centres its `>>>>>` marker on the failing statement
  and in that round the marker sat on the batch's FIRST line while the log said
  "… 37 earlier statement(s) dropped".
- **`load("items")` does not load the items' properties — Microsoft says so.**
  "You must explicitly specify each property you need from collection items, as
  they won't be loaded by default, including scalar properties." Every
  collection load in this file names its properties; the re-read before grouping
  was the one that did not, and `id` is the only thing it reads. Changed to
  `items/id` on 2026-08-07 — after being changed and REVERTED once the same day,
  which is the part worth keeping: the first attempt had no source behind it and
  the suite went red, so it was traded back rather than argued for. The red test
  turned out to be asserting the wrong property (which writer landed the tag,
  rather than whether the chart is re-editable), and it only surfaced because
  the fake's `hollowReads` keys on the projection string. Guarded by a source
  scan in `web-host.test.ts`.
- **`untrack()` is not available on this host, and that is now measured rather
  than assumed.** Microsoft's performance guidance names untracking as the
  remedy for our exact symptom — "large batch operations may generate a lot of
  proxy objects… Calling untrack() after your add-in is done with the object
  should yield a noticeable performance benefit when using large numbers of
  proxy objects" — and `renderShapesChunked` holds one proxy per shape for a
  whole draw, hundreds a run, untracking none.

  `untrack-available` had answered `no` for months, but it asks a NULL-OBJECT
  slide proxy, which is the one kind most likely to lack the method; the probe's
  own comment suspected exactly that and nobody had put the partner. On
  2026-08-12 (`89675b6`) `untrack-available-on-shape` asked a proxy
  `addGeometricShape` had just returned — the kind the draw loop actually holds
  — and it answers `no` too. So the confound is gone and the `no` is about the
  host. The draw path is not getting untracking, and that is a decision on
  evidence rather than an omission. Do not re-propose it.

- **Drawing cost grows with the shapes ALREADY on the slide — and the request
  context has nothing to do with it.** `what makes a long run slow down` finally
  measured something on 2026-08-09, after two outings that died in setup. Eight
  rounds of twelve rectangles per arm, one arm holding a single context open and
  the other taking a fresh one each round:

      one context     2339 3177 3704 4363 5010 5775 6372 6769 ms   +101%
      fresh contexts  2852 3397 3936 4636 5165 5635 6736 7352 ms   +105%

  Both arms grew, within four points of each other, so the CONTEXT is not the
  variable — which is what the experiment was built to decide, and it is now
  decided. What each round shares is that it adds twelve shapes to a slide that
  already holds twelve times the round number: the per-round increments are
  roughly constant (+838, +527, +659, +647, +765, +597, +397), so a round costs
  about `2339 + 630(n-1)` ms and the TOTAL is quadratic in shapes per slide.
  Ninety-six shapes on one slide cost 37.5 seconds; the same ninety-six over
  eight slides would cost about nineteen. That is the arithmetic behind the
  17-second batches, and behind dense charts being the ones that fail.
  **REPRODUCED on a second round (`87dc418`), and the second round corrects the
  first's reading of it.** The curves land on top of each other:

      one context     2395 3095 3713 4400 5072 5734 6398 7455 ms   +107%
      fresh contexts  2428 3056 3676 4356 5557 5619 6345 7264 ms   +106%

  Round for round within a few percent of the run above, on a different build,
  with the same `deckBefore=7`. The per-round increment is ~660ms both times.
  That part is solid.

  **And it is PER-SLIDE, not per-deck — the two arms already separated that.**
  This file said a third arm on a fresh slide per round would be needed and that
  the host would not give one out. Wrong: the arms run in SEQUENCE on DIFFERENT
  slides, so the second arm's first round is the measurement. It starts after
  the first arm has put ninety-six shapes on another slide in the same deck, and
  if the cost were deck-wide it would start slow. It does not: 2395 against 2428,
  a 1.4% gap. The 22% gap in the first round (2339 against 2852) that this file
  called "the only hint" of a deck effect did not reproduce, and one sample of a
  22% gap next to one of 1.4% is noise, not a mechanism.

  So: adding shapes to a slide makes THAT slide expensive to draw on, and costs
  the next slide nothing. Which is why spreading a deck's charts one per slide
  is not merely tidier — it is the difference between linear and quadratic.

  **The routine rounds add a second hint, and it points the other way.** Two
  consecutive rounds stalled in the same two scenarios, on the same call, at
  much the same times — `drawing shapes 1-10 of 24` giving up after 45 seconds
  at 481s and 591s on `1fd6aa3`, and at 456s and 567s on `40b5e44`, the latter
  on a deck that started with ONE slide (`round starting deckSlides: 1`). Those
  draws go onto slides holding a couple of charts, not ninety-six shapes, so
  per-slide cost does not explain them.

  And per-deck SHAPE COUNT does not either, now that the second experiment round
  has ruled it out — ninety-six shapes on a neighbouring slide cost the next
  slide 1.4%. What is left is the TAB: elapsed time, memory, or whatever else a
  long-lived Office.js session accumulates that a shape count cannot see. That
  is a narrower answer than "the deck or the tab", and it is narrower because a
  second sample was taken rather than argued for.

  **But elapsed time is not the whole of it either, and the 2026-08-09 sweep of
  seven rounds says why.** Read as a distribution rather than as anecdotes, the
  stalls have two properties nothing here had noticed:

  - **Every stall is the FIRST batch of a scenario's draw. Eight for eight,
    `drawing shapes 1-10 of 24`, never batch 2 or 3.** If the host were simply
    fatigued, stalls would land anywhere in a chart's three batches. They land
    on the first sync after some other operation and nowhere else.

    **OVERTURNED on 2026-08-12 (`275a76a`), by the first counter-example in
    nine rounds:** `drawing shapes 11-20 of 24` stalled — batch TWO — and its
    predecessor is the tell. Batch one of the same chart had just answered in
    **25.8 seconds**, near the top of the surviving range, so that draw was
    already crawling before it died. Read the eight-for-eight as a pattern that
    held while the host was healthy, not as a law.

  - **Elapsed time cannot be the variable on its own**, because in every one of
    those rounds the scenario immediately AFTER the stalled one — older tab,
    more elapsed time — drew its full twenty-four shapes and passed. r6 is the
    clean case: `a selected shape survives an insert` stalls at 410s, `edit the
chart the user selected` passes at 459-514s, `the chart is actually visible`
    passes at 517-586s.

  What every stalled scenario had in common was its PREDECESSOR: the selection
  ladder (×4), `stop a run part-way`, which does NOT abort a draw mid-flight —
  it asks for the stop BEFORE the insert, and across 69 archived rounds it has
  committed zero batches, so no shape is ever queued (×3), and
  a scenario that itself selected and stalled (×1). That was 8 of 8.

  **RETRACTED on 2026-08-09 by round 7, which is why it was written down as a
  correlation.** `a selected shape survives an insert` PASSED — in the routine
  round, at 407s, eighth in the list and immediately after the ladder, which is
  the exact position and the exact predecessor that had stalled it four times
  out of four. Same build family, same battery, same order. A clean
  counter-example at the identical position kills the predecessor account
  outright; nothing about "what ran before it" survives.

  The picked round ran the same day and agrees. Alone on the deck, with only the
  two head-of-round inserts in front of it, it drew all three batches (8.2s,
  12.2s, 12.4s) and passed — at 750s of tab life, thirteen minutes in, which
  takes elapsed time down with it for that scenario.

  So of the four candidates, three are now out for `a selected shape survives an
insert`: the scenario, its predecessor, and the tab's age. What is left is
  that this host stalls INTERMITTENTLY, and the run of four was a run.

  **Both things that were built to settle it have now reported, and both did.**

  - **The call-level record fired on its first outing.** The log used to have
    nothing at all between a scenario announcing itself and its first `batch
committed` — three to five seconds of probe reads, deck inventories and
    selection calls, all invisible — which is the only reason the account was
    ever at scenario level. Round 7's one draw stall says:

        gave up waiting  what="drawing shapes 1-10 of 24"
                         afterAnswering="rasterising a slide"  idleMs=1

    A rasterise answered, the draw's first sync went out a millisecond later,
    and it never came back. Every future round carries this for free.

    **`idleMs` IS NOT THE VARIABLE — asked and answered on the next round.**
    The gap was recorded for the first batch of every draw, stalled or not, and
    round 8 (`d812d0c`) reported four batches that all SURVIVED:

        82.6s  idleMs=2182   edit a chart on the visible slide

    100.1s idleMs=2 insert onto a slide that already has content
    129.9s idleMs=1 insert onto a slide that already has content
    188.3s idleMs=854 same scale across the deck

    A surviving first batch went out **one millisecond** after the previous
    answer, which is exactly what the stall reported. A value that occurs in
    both populations cannot separate them, so the gap is dead as a lead and
    nothing should be built on it — least of all a wait before drawing, which
    this host has already punished once (see the scratch-slide gotcha).

    The instrumentation stays. It cost one trace field, it killed a plausible
    hypothesis in a single round, and the half that is still live rides along
    with it: `afterAnswering` names the CALL, and the identity of the call is
    the lead the gap turned out not to be.

    **The name alone was still not enough, and `afterAnsweringMs` is the third
    field on that line.** Round 16's control arms take 22-29 seconds each and a
    rasterise plus a seven-shape draw sits inside one number, so "which half is
    growing" — the question the arm exists to answer — could not be asked.
    `withTimeout` already stamps when a named call is issued and when it
    answers, so the duration is a subtraction at a seam that was already there.
    It is written on the first batch of EVERY draw and on every stall, which is
    the rule two rounds above paid for.

    **It answered on its first outing, and the answer is that the rasterise is
    not the expensive half.** Round 17 (`4feb5be`) puts a rasterise at 915ms,
    1246ms and 2387ms, against 2-3ms for `counting the deck's slides` and
    ~700-900ms for a tag write. The control's arms take 22-29 seconds, so the
    rasterise is at most a tenth of one and the rest is the seven-shape draw.
    Anything still hunting the arms' cost should be looking at the draw.

  - **The picked round ran and answered.** The criterion was set in advance:
    stalls → the scenario itself; passes → predecessor or position. It PASSED,
    so the scenario is out, and the routine round passing it at the same
    position takes the predecessor with it.

  **Round 9 removed the last of that too.** It produced TWO stalls, and they
  named two DIFFERENT predecessors:

      496.6s  drawing shapes 1-10 of 24   afterAnswering="selecting a shape"     idleMs=1
      608.2s  drawing shapes 1-10 of 24   afterAnswering="rasterising a slide"   idleMs=1

  So "the sync after a rasterise" is not the pattern — `a selected shape survives
an insert` stalled after a selection call, having passed the two rounds before.
  And `edit the chart the user selected` ALSO selects a shape and then draws, in
  the same round, and survived.

  **And the predecessor's name is in exactly the condition the gap was in: no
  baseline.** It is recorded on stalls and nowhere else, so thirteen draws
  survived in that round without saying what they followed. That is the same
  mistake `idleMs` cost two rounds to kill, made a second time with a different
  field — which is why the first batch of every draw now records BOTH. One round
  decides whether the identity of the preceding call discriminates or goes the
  way the gap did.

  **Round 10 gave that field its baseline, and unlike the gap it SURVIVED.**
  With `afterAnswering` on every first batch rather than only on stalls, the two
  populations separate cleanly:

      survivors followed   moving the view to a slide, counting the deck's slides,
                           writing the chart's origin tag (x5), re-reading a slide
                           to tag the chart it would not tag (x4), selecting a
                           shape, reading the selected chart
      the stall followed   rasterising a slide

  Twenty-nine surviving first batches across rounds 9 and 10, not one of them
  after a rasterise; two stalls after a rasterise, in consecutive rounds. And
  `selecting a shape` — round 9's other stall — turns up among round 10's
  survivors, which is what a non-cause looks like.

  **And round 11 killed it, the same way every other candidate died.** That
  round contains two draws after a rasterise, 150 seconds apart:

      584.5s  pass   after="rasterising a slide"    (the chart is actually visible — PASSED)
      734.8s  STALL  after="rasterising a slide"    (the control's rasterise arm)

  A value that occurs in both populations cannot separate them. So
  `afterAnswering` has gone the way of `idleMs`: **no call before the draw
  predicts the stall.** Every candidate is now eliminated — the scenario, the
  preceding scenario, the tab's age, the idle gap, and the identity of the
  preceding call. What is left is a host that stalls the first sync of a draw
  intermittently, at roughly one or two draws in fifteen, and nothing yet
  distinguishes which.

  **The control's own verdict that round was WRONG, and that is the more useful
  lesson.** It reported `the draw after a RASTERISE did not land` — a claim the
  same log contradicts — because the two arms ran in a fixed order, cheap first
  and rasterise second, so the rasterise arm was always the later one. A
  diagnostic that manufactures a finding is worse than no diagnostic, and this
  one would have manufactured that finding every round it stalled.

  It is counterbalanced now: rasterise, cheap, cheap, rasterise, with each call
  type running once early and once late. `rasteriseArmVerdict` names the CALL
  only when every rasterise arm fails and every cheap arm draws; it names
  POSITION when both late arms fail and both early ones draw; and it says "no
  pattern" for anything else, which on eleven rounds of evidence is the answer
  to expect. The arms draw eight-shape charts rather than the battery's
  twenty-four, so four of them cost about what two of the old ones did.

  This is the same shape as the two experiments that already paid: `the chart is
actually visible` and `what makes a long run slow down` each spent four rounds
  saying "it crashed again" and one picked round saying which call.

  **Round 12 (`3223293`, 2026-08-11) ran the counterbalanced control for the
  first time, and it said "no pattern" — which is the design working.** Both
  rasterise arms drew and both cheap arms drew, no stall in any of the four
  (`poolRasteriseArms` on that round: `rasterise 2/0, cheap read 2/0`). The
  round's one draw stall was somewhere else entirely, and under the OLD fixed
  ordering the verdict would have had nothing to manufacture from either. Run a
  few more of these and "no pattern" repeated IS the finding; stop instrumenting
  at that point.

  **Round 16 (`1fc21b9`) is the third "no pattern", and the control turned out
  to be measuring something else for free.** All four arms drew; the verdict was
  `a draw straight after a rasterise is no worse than one after a cheap read`.
  That is three counterbalanced rounds, twelve arms, and this paragraph's own
  criterion says the instrumenting can stop.

  The by-product is the useful part. The four arms draw the same seven shapes
  onto the SAME slide in sequence, so with `onSlide` recorded they are four
  points on the per-slide cost curve, in an ordinary round, with no experiment
  to schedule:

      onSlide 23 → 22.7s      onSlide 30 → 25.6s      onSlide 37 → 28.9s

  About **+0.44s per shape already on the slide**, for a seven-shape draw. The
  dedicated experiment (`what makes a long run slow down`) is `pickedOnly` and
  costs a whole round; this arrives every time the battery runs. If the curve is
  ever in doubt again, read these three numbers before scheduling anything.

  **The SLOPE reproduces and the INTERCEPT does not.** Round `393e6e4`
  (2026-08-12) puts `onSlide 24 → 9.7s, 31 → 11.5s, 38 → 16.5s` — +0.49s per
  shape, within a tenth of the figure above, on less than half the absolute
  cost. So "+0.44s per shape" travels between rounds and "a seven-shape draw
  costs 23s" does not; quote the slope, never the level. These arms are an
  add-only path, which is the one place `onSlide` was honest even before the
  counter was made net, so the two measurements are comparable.

  **Round 15 (`756682e`, 2026-08-11) said POSITION, and that is the design
  earning its keep — and contradicting round 14.** Round 14 reported `every draw
after a RASTERISE failed and every one after a cheap read landed, interleaved
so position cannot account for it`; round 15 reported `both LATER draws failed
and both earlier ones landed, whichever call preceded them — this is position
or elapsed time, not the rasterise`. Same battery, same counterbalancing, two
  rounds, opposite verdicts. Pooled across the four counterbalanced rounds the
  arms are `rasterise 5 ok / 3 stall` against `cheap 7 ok / 1 stall`, which is a
  direction and not a result.

  Read the disagreement as the finding rather than as one of the two rounds
  being wrong. It also settles a question about the WORDING: these verdicts were
  nearly softened to "consistent with" on the grounds that four arms cannot
  support a causal claim, and that would have been a mistake — each verdict is a
  true statement about its own round, the cross-round uncertainty belongs in
  `poolRasteriseArms`, and hedging both would have hidden exactly the
  disagreement that is informative here.

  **Round 13 (`7027f96`, 2026-08-11) ran it again and said "no pattern" again**
  — `all four draws landed — a draw straight after a rasterise did not stall`.
  Two counterbalanced rounds, four arms each, zero stalls. That is two of the
  "few more" this paragraph asks for; one more like it and the instrumenting
  should stop, because "no pattern" repeated IS the finding.

  That round also re-killed `afterAnswering` by a second route, without anyone
  designing for it:

      152.7s  pass   after="moving the view to a slide"   idleMs=1109
      351.6s  STALL  after="moving the view to a slide"   idleMs=14108

  Same predecessor, both populations, 200 seconds apart in one round. Every
  candidate stays dead.

  **Rounds `957aca0` and `393e6e4` (2026-08-12) are the fourth and fifth "no
  pattern", so STOP READING THIS CONTROL as an open question.** Twenty arms
  across five counterbalanced rounds; the only rounds that ever said otherwise
  are 14 and 15, which said opposite things. The control stays because it is
  four cheap draws that double as cost-curve points, not because the question is
  live. `393e6e4` also supplied yet another distinct stall predecessor —
  `writing the chart's origin tag` — which is simply one more name on a list
  that has never predicted anything.

  **The one genuinely new number is that `idleMs=14108`, and it does not bring
  the gap back.** It is seven times the largest gap ever recorded on a surviving
  first batch (survivors span 1ms to 2182ms across all rounds; that round's own
  span 1ms to 2054ms), so it is tempting to read as sufficient-though-not-
  necessary. Do not, yet: it is a single observation, and 1ms sits in both
  populations, so the field still cannot classify a draw. Worth one line in the
  next round's read — if a second stall arrives with a gap in the tens of
  seconds while survivors stay under three, that is a real signal. One is an
  anecdote, and this file has a paragraph about exactly that mistake.

- **The self-test piles almost everything onto ONE slide, and the per-slide cost
  curve then charges it for that.** The `onSlideKey` fix landed and paid on its
  first round (`275a76a`): slide `257#3695341871` is the target of EIGHT of the
  ten scenarios, and this run's own shape count on it climbs 20 → 68 → 92 → 116
  → 140 → 144 → 165 as the battery proceeds. Nothing else in the deck goes past 34.

  That number is the input to the quadratic cost this file measured at about
  **+0.44s per shape already on the slide**. At `onSlide: 144` the overhead
  alone is ~63 seconds, which is more than `BATCH_TIMEOUT_MS` — so a draw there
  is expected to stall, and one did: `the chart is actually visible` gave up on
  `drawing shapes 1-9 of 9`, a NINE-shape chart, at exactly that count.

  So some of what has been read as "this host stalls intermittently" is the
  battery's own arithmetic arriving. Not all of it — the round's other stall was
  at `onSlide: 34`, and the rasterise control drew fine at 165 — but a scenario
  ordered late is measurably harder to pass than the same scenario ordered
  early, and no round before this one could see that.

  **Read those counts as an upper bound, not as the slide's contents.**
  `onSlide` counts what this run DREW and is never decremented, so a chart
  redrawn in place adds its whole node count again while the slide stays the
  same size — 92 on the counter against 3 in the deck inventory, in round
  `957aca0`. The battery's climb above is mostly genuine (its scenarios insert
  rather than redraw) but the deck-wide rescale's is not, and the
  `FAST IS THE BROKEN MODE` gotcha carries the case that turns on it.

  The slide-sharing is deliberate and the reason is good: `chartIsVisible` must
  never rasterise a slide the run just added. Sharing a slide that is not the
  BUSIEST one costs nothing, and `leastLoadedChart` is that: every scenario that
  needs a chart now takes the one on the slide this run has loaded least,
  counted from `shapesDrawnOn` — a number the renderer already kept, so no host
  call is added.

  **Ties keep the deck's order**, which is what makes it safe to drop in: on a
  fresh run every load is zero, so it picks exactly what `found[0]` picked and
  every existing expectation holds. It diverges only once this run has actually
  loaded a slide.

  Measured against the fake, the same battery concentrates **0.619** of its
  draws on one slide with `found[0]` and **0.369** with the rule — 195 shapes
  on the worst slide against 120. The fake spreads more than the real host does
  to begin with (its probe charts land one per slide, where the real deck had
  nine on one), so that understates the effect rather than overstating it.

  **It landed and it worked on the real host** (2026-08-12, `957aca0`). The
  round's loads read **96, 45, 44, 24, 20, 20** across six slides, against 165
  on one slide the round before — and the round went from two draw stalls to
  one. At ~+0.44s per shape present, taking the worst slide from 165 to 96 is
  about thirty seconds of overhead off the busiest draw, which is the difference
  between sitting inside `BATCH_TIMEOUT_MS` and not.

- **A field can be recorded on both populations and STILL be useless, if it is
  rarely populated.** `onSlide` — how many shapes this run had already put on
  the slide, the input to the quadratic cost curve — was added on every batch,
  stalled or not, exactly as the rule below demands. Its first real round
  (`1fc21b9`) carried it on **7 of 46 batches**, and on exactly ONE alongside
  `prevBatchMs`, because it was conditional on `opts.slideId` and most draws do
  not pass one. So the pair that was supposed to measure the cost curve produced
  a single point.

  The rule below asks "what is this value on the runs that WORK". Ask the second
  question too: **on how many of them is it there at all?** A field present a
  sixth of the time answers nothing, and the shortfall is invisible until a
  round is read — the numbers that ARE there look perfectly healthy. It keys on
  a `(visible)` sentinel when the caller names no slide now, and emits
  `onSlideKey` beside the count so a reader can see when a total may span more
  than one slide rather than having to assume it does not.

  **The next round caught the field being wrong, and the KEY is the only reason
  it did.** With `onSlide` on every batch, round 17 (`4feb5be`) showed the
  deck-wide rescale climbing 72, 82, 92, 96 … 260 — a smooth rising curve, on a
  deck whose fullest slide held 24 shapes. `updateChartsInSlides` never filled
  `slideId` in, so eight charts on eight slides all keyed on `(visible)` and the
  counter pooled them. The numbers looked perfectly healthy and described no
  slide in the deck; `onSlideKey: "(visible)"` on every line is what gave it
  away. The update path names its slide now.

  Two rules out of one field: emit a key beside any POOLED total, and when a
  diagnostic starts reporting the shape you expected, check what it is keyed on
  before believing it.

  **The INSERT path was still on the sentinel a round later, and by then the
  counter had a safety job.** Round `393e6e4` shows every batch of `insert onto
a slide that already has content` keyed `(visible)` while every other scenario
  named its slide — the last caller that never filled `slideId` in, and the
  commonest path in the add-in, since the pane's Insert button names no slide.
  That is not untidy any more: `slideHoldsOnlyChart` reads this counter to
  decide whether an empty read of a slide is believable, and it authorises
  DELETING the user's slide. A chart inserted the ordinary way banked its
  twenty-four shapes under the sentinel, so `shapesDrawnOn(realId)` answered
  ZERO for a slide the run had just filled and the guard would have believed the
  host on exactly the slide it exists to refuse.

  `slideKeyFor` reads the slide's OWN id when the caller named none. It is free
  — `insertSceneIntoSlide` already queues `slide.load("id")` before the first
  batch — but only from the SECOND batch, because the first batch's line is
  written before the sync that answers it. So the first batch's shapes are moved
  across the moment a real key appears, and only this draw's own shapes move:
  the sentinel is shared, and taking its whole total would steal another draw's
  count.

  **Reproduced exactly on `89675b6`**: the sentinel again appears on exactly
  two batches in the round and the carry-over again reports ten shapes twice.
  Two rounds, same numbers, so this is settled rather than a good minute.

  **It landed and both halves worked on the real host** (2026-08-12,
  `ee1741e`). The `(visible)` sentinel drops from every insert batch to TWO in
  the whole round — the first batch of each unnamed draw, before the host has
  answered the id — and the carry-over says so out loud twice: `moved this
draw's shape count onto the slide the host finally named {from: "(visible)",
to: "257#1897035307", shapes: 10}`. Ten shapes each time, which is the batch
  size, which is exactly what the arithmetic predicted.

  **That carry-over is the interesting part of the change, because nothing in
  the fake could make it fail.** The fake answers `slide.id` from the first
  read, so the key never transitioned and the block was decoration — the exact
  thing this file has a rule about. `faults.slideIdUnreadableBeforeFirstSync`
  is the fake being a real host for one window: the id throws
  `PropertyNotLoaded` until the run's first sync. With it armed the guard fails
  `expected 14 to be 24` — a ten-shape first batch stranded on the sentinel —
  and passes with the carry-over in place. Coarser than the real per-property
  load rule, and enough for the only window that matters.

- **A value recorded only on FAILURES cannot be compared against anything, and
  this project keeps building them.** Four in one session: `idleMs` and
  `afterAnswering` were written on stalls but not on the draws that survived;
  the settle's two writes shared one label so a refusal named neither; and
  `listChartsInDeck` traced only when a scan came back SHORT, which made every
  deck scan in every round invisible. That last one has a price tag: round 10's
  `stop a run part-way` took 39.4 seconds against 2.6-3.2s in the eight rounds
  before it, and the log had a 39-second hole where the scan was. The stop
  itself was instant — the verification after it was not, and nothing said so.
  The scan is also the operation the quadratic per-slide cost predicts should
  grow worst, since it reads every slide's shapes on a deck the battery keeps
  adding to, and it was the one operation never measured. Before adding a
  diagnostic field, ask what its value is on the runs that WORK; if the answer is
  "it is not written", the field cannot discriminate and is not yet a
  measurement.

  **A fifth, found 2026-08-11: GROUPING spoke only when it refused.** Round 16
  left a slide carrying one `PowerChart` group (id 51) **plus four loose
  shapes** — `label-1-3`, `baseline`, `series-label-0`, `series-label-1`, ids 47
  to 50, every one inside the chart's own box and every one with a lower id than
  the group. Two readings fit and no log could choose between them: the group
  took a subset of the chart, or four shapes from an earlier draw outlived a
  redraw. Both are real defects and they want different fixes.

  A partial group is a designed outcome, not an accident — the re-read
  deliberately KEEPS a partial match rather than falling back, because every
  shape in it is provably ours where the positional rule is a guess. What was
  missing is that it never announced itself. `grouped the chart's shapes` is
  written after the grouping sync (so the name is an outcome it knows) and
  carries `charts`, `partial`, `left` and `by`. A partial group is worth naming
  on its own account: the loose remainder does not move with the group, so the
  user drags the chart and leaves its baseline behind.

  The fake could not model it either, and its own comment said so wrongly.
  `hollowReads` describes a host that "asked about 19 shapes and got 3 back" and
  then returns `[]` — the limit case, never the one it documents. An empty read
  makes `chooseGroupMembers` say "group nothing" and the chart stays loose in
  one piece; a SHORT read is kept and grouped, so the chart is split. Opposite
  branches. `faults.readsMissing` is the short case.

  **ANSWERED on the line's first outing (`4feb5be`, 2026-08-11): the group took
  a SUBSET.** Fifteen successful groups in that round, fourteen of them
  `partial: 0`, and exactly one — chart 4 of 8 —
  `grouped the chart's shapes charts=1 partial=1 left=0:4 by=ids`. The deck it
  left carries the same slide as the round before, with the same four names
  loose beside the group: `label-1-3`, `baseline`, `series-label-0`,
  `series-label-1`. So the other reading — four shapes from an earlier draw
  outliving a redraw — is dead, and the defect is the one that was designed in:
  a re-read that matched 20 of 24 ids groups the 20 and leaves 4 on the slide.
  One round, no reasoning, because the success path finally spoke.

  **A partial match is GROUPED, and the chart is split — the owner's call, taken
  2026-08-19, reversing what this paragraph said for eight days.** The branch had
  refused because a grouped-and-split chart looks like one object and is not:
  drag it and the baseline stays behind, with nothing said. What refused it was
  the belief that the alternative was free — "an ungrouped chart is still tagged,
  still re-editable" — and the archive refutes exactly that:

      grouped      64 chart(s),  1 lost the tag =  2%
      NOT grouped  62 chart(s), 41 lost the tag = 66%

  So the choice is between four loose shapes and a chart that is not a chart. The
  config tag goes on the group, which is a handle made in the grouping batch and
  never resolved; without a group it falls back to a `created` handle this host
  refuses two times in three.

  **One bound, and it is NOT in the call as the owner stated it.** He weighed 20
  of 24, which is what every round on record produces; the code takes the subset
  only when it is a strict MAJORITY of the chart. A group holding one shape while
  twenty-three sit around it saves the config and destroys the object — dragging
  the chart moves a label. Majority rather than a tuned share, because "more of
  the chart is inside the group than outside it" is a statement about the object
  and any other number would be invented.

  **The stranded remainder is deliberately NOT written into the parts tag**, and
  that was the first thing tried. It sounds strictly better — the next update
  would take the loose shapes with the group — and it is how the trade becomes
  data loss: the only ids we hold for those shapes are the ones CREATION
  returned, and creation ids are precisely what this host has been seen not to
  answer to (`withOwnId 7 of 7 … matched 0`, rounds 068/069). The parts tag is a
  list the update path DELETES BY. A shape the host would not name is not a shape
  we may delete.

  **Two seams had to move with it, and the second is the one nothing would have
  caught.** `freshMembers` now carries partial lists, and two readers must have
  whole ones: `ungroupedFallback` builds the parts tag off that map — a short
  parts tag is worse than none, because the next update deletes what it can name
  and leaves the rest, so the chart grows on every edit — and the single-member
  tag-target swap assumes `fresh[0]` is the anchor. `wholeMatch` is the set that
  separates them, and the parts-tag case is reachable: grouping can still throw
  after a subset was chosen.

  `chooseGroupMembers` deliberately did NOT change. It is handed the re-read's
  members and nothing else, so `["a", undefined]` could be two of a two-shape
  chart or two of twenty-four — it cannot see `created.length`, so it cannot know
  whether the group would still be the chart. The caller that can see both
  numbers is the one that takes the trade.

  The partial branch does NOT fall through to the positional rule. That branch
  is safe only when nothing matched by id: a slide holding the user's own shapes
  can satisfy `items.length >= created.length` while the chart itself read
  short, and "the last N" would then reach past the chart into the user's
  content and group it in — to be deleted with the chart on the next update.

  `partial` on the success line is an OUTCOME again rather than an invariant, and
  it is the only line that reports what the call costs: `left=0:4` says chart 0
  has four shapes loose inside its own box. A round archive spans both meanings
  — `grouping:` on the short-read line is what dates a round.

- **A trace may not be NAMED for an outcome it is written before knowing.** The
  per-batch draw line was called `batch committed` and is emitted one statement
  before the sync it describes — on purpose, because the sync is where a bad
  host goes quiet and the number has to be on screen while you wait. The
  ordering was right and the name was false, so every stall on record left
  behind a line saying the batch it killed had committed. It cost two hand
  analyses of the same data: one paired the lines with draws and reported 0
  stalls in 32, the other counted them as successes and produced a 6x rasterise
  effect that was not there. `scripts/triage.mjs` and `test/triage.test.ts` both
  carry comment blocks whose whole job was to warn the next reader off it — a
  workaround in every reader is the tell that the writer is wrong. It is `batch
issued` now, which is what the line actually knows, and there is no commit line
  at all: the next batch's `issued` implies it and the last one's is the draw
  returning. The triage fixtures deliberately keep the old spelling, because
  rounds saved before 2026-08-11 carry it and the pooling function must go on
  reading neither name.

- **A stall is DEATH, not slowness — do not raise `BATCH_TIMEOUT_MS` hoping for
  an answer.** Seventeen abandoned calls across nine rounds, and not one of them
  ever came back: `a call we gave up on finally answered` appears zero times in
  any round file, including the 100-200 seconds each round keeps running
  afterwards. Round 9 looked like a counter-example and is not — two probe
  questions there answered at 8429ms and 8399ms against an 8000ms budget, which
  reads as a late answer and is the probe's RETRY: it replaced the scratch slide
  and re-asked, and the elapsed time is measured from the first attempt. And the population is bimodal with an empty band — of 327 batches
  that DID answer the slowest took **29.2s**, against a 45-second budget, so
  nothing has ever landed between 29s and 45s. A batch answers within ~29s or
  never.

  **The 29.2s figure is stale as of 2026-08-12** (`275a76a`), the slowest round
  on record: three surviving batches over it, at 29.3s, 30.8s and **31.1s**. The
  empty band is therefore 31s-45s rather than 29s-45s. The argument it supports
  is unchanged — a batch answers well inside the budget or never — but do not
  re-quote 29.2 as the maximum.

  Two things follow. Raising the budget buys nothing, and lowering it to ~35s
  would save ten seconds a round for a 1.2× margin instead of 1.5× — not worth
  it either. Leave it alone; the number is fine and the finding is what it means.

  **That was read out of a trace line's ABSENCE, which is this project's most
  expensive habit** (`settleUntaggedCharts` was "ran and failed" for two
  sessions when it had never run). So it is measured now rather than inferred:
  a stalled scenario waits `LATE_ANSWER_WAIT_MS` and its verdict says which
  happened, in words, every time (`stallDetail`). The plumbing behind that had
  Round 12 is the first round where that measurement did the talking rather
  than an absence: `same scale across the deck` came back
  `the abandoned call had still not answered 3s later`, in the verdict, in
  words. Two more abandoned calls that round (one probe question, one draw
  batch) and still no `a call we gave up on finally answered` anywhere — but
  the point is that the round SAYS so now instead of leaving a reader to infer
  it from a line that is not there. (The count above is not re-tallied here on
  purpose; it is the one number this file owns, and adding to it from a single
  round file is how it would start to drift.)

  never been tested on the DRAW path either — every existing late-answer test
  goes through `insertDemoDeck` — so there is now one that stalls a real draw
  batch and checks the report carries the late answer.

- **FAST IS THE BROKEN MODE, not the healthy one.** The draw times are bimodal —
  ~17s per batch or ~3-5s — and this file said until 2026-08-08 that the host
  "recovers mid-run and goes again". It does not. Within ONE run of `same scale
across the deck`, eight updates of the same chart at the same size: charts 1-4
  took ~17s per batch with no grouping failures at all, and charts 5-8 took
  ~3-5s and every one of them carries `the re-read before grouping came back
empty` and `not grouping`. The boundary is exact and the correlation is
  8-for-8. Batches got cheap because the host stopped answering for the shape
  collection, so there was nothing left to group or read back — it was not
  recovering, it was degrading further. Do not read a fast stretch as health,
  and do not "fix" the slow one: `what makes a long run slow down` was written
  on the assumption that slow means degraded, and it has that backwards.
  (AutoSave was the obvious candidate for a periodic slowdown and is ruled out:
  the regimes are long contiguous blocks — eight slow, then ten fast — not
  anything alternating on a timer.)

  **The bimodality is now MEASURED rather than inferred, and the heading needs
  narrowing.** Round 16 (`1fc21b9`) recorded `prevBatchMs` on 27 batches:

      fast   3020 3178 3281 3303 3623 3700 4789 5124 5165 5372 5413 6036
      slow   11723 14236 16294 16488 16603 16708 17070 17263 17312 17359
             17714 17791 21195 23022 23428

  Nothing between 6.0s and 11.7s. Two populations with an empty band, which is
  what every earlier round asserted from wall-clock arithmetic and none had ever
  put a number on.

  And the same round contains a fast batch that is not broken at all: `edit a
chart on the visible slide` drew at 3.0s and 4.8s, passed, round-tripped its
  config, and reported `errors: 0`. So FAST ALONE IS NOT THE SIGNAL — a healthy
  host draws fast onto a slide with little on it, which the per-slide cost curve
  already predicts. What the heading is really about is a flip to fast INSIDE a
  scenario that was drawing slowly, which arrives with the collection going
  quiet a chart later. Do not read a fast batch as broken; read a flip as
  broken.

  **Round `957aca0` (2026-08-12) put a second candidate under the flip, and it
  cannot yet be chosen between — do not reason about it, and do not build on
  either.** The rescale's batch lines finally carry `chart`, `onSlideKey` and
  `prevBatchMs` together, and they say:

      chart 1/8   slide 257   17714 18726 ms   SLOW
      chart 2/8   slide 257   17840 17250 ms   SLOW
      chart 3/8   slide 257   21097 16643 ms   SLOW
      chart 4/8   slide 258    4674  5506 ms   fast
      chart 5/8   slide 259    3783  5263 ms   fast
      chart 6/8   slide 260    3300  5163 ms   fast

  The flip at chart 3/4 is exactly the point where the charts stop sharing one
  slide and start having one each. It is also, in the same round, the point
  where grouping stops surviving — so **the slide and the collection's health
  change together**, and one round cannot say which the clock is following.

  What it DOES kill is elapsed time, again and from a new direction: the fast
  charts run at 395-434s and the slow ones at 242-381s, so the LATER draws are
  the quick ones inside a single scenario.

  **The measurement that would separate them is not available yet, because
  `onSlide` is not what its readers think it is.** It counts what this run has
  DRAWN on a slide and is never decremented, so a chart redrawn in place adds
  24 to it every time while the slide's real contents do not move: slide 257's
  counter reached **92** in this round, and the deck inventory at the end of the
  same run shows that slide holding **3** shapes. Every reading of "shapes
  already on the slide" taken from this field on an UPDATE path is therefore
  inflated — including the +0.44s/shape slope, whose own source (the rasterise
  arms) happens to be an add-only path where the field is honest.

  **VALIDATED against the deck on 2026-08-12 (`1789749`).** Three slides in
  that round held ungrouped charts, and each one's last `batch issued` reads
  `onSlide: 20` against a deck inventory of **24** — 20 being the count BEFORE
  that slide's final four-shape batch. Exact, three for three. The grouped slide
  reads 34 against 11 for the reason above: the inventory counts top-level
  shapes and grouping collapses twenty-four into one. Compare these two numbers
  only on a slide whose charts stayed ungrouped, and when you do, they agree.

  **The counter is net now, so the next round answers this.** Both places the
  run takes its own shapes off a slide give the count back — the redraw's
  delete and `deleteShapesById`'s stray sweep — and `replacedShapeCount` is the
  part that is not arithmetic: deleting a GROUP removes its children in one
  call, and a grouped chart's parts tag does not list them, so the call count
  says 1 for a chart that occupied twenty-four. It uses the calls when they
  enumerate the chart and the size of the chart going back when they do not, so
  a same-size redraw nets to zero. Read the next round's `onSlide` on the
  rescale against the deck inventory; if they agree, the flip is answerable.

  **ANSWERED on the next round (`393e6e4`, 2026-08-12), and the answer is
  GROUPING, not the slide — the paragraph above backed the wrong horse.** The
  counter fix did what it was for: every chart in the rescale now starts its
  redraw at `onSlide: 0`, because a redraw replaces a chart and nets to zero.
  With that variable held flat, the flip is still there:

      chart 1/7   slide 257   onSlide 0   18088 21688 ms   SLOW   GROUPED
      chart 2/7   slide 257   onSlide 0   19650 19442 ms   SLOW   GROUPED
      chart 3/7   slide 258   onSlide 0    5323  7507 ms   fast   not grouped
      chart 4/7   slide 259   onSlide 0    6479  8381 ms   fast   not grouped
      chart 5/7   slide 260   onSlide 0    6845  9579 ms   fast   not grouped
      chart 6/7   slide 261   onSlide 0    5072  9665 ms   fast   not grouped
      chart 7/7   slide 262   onSlide 0    6015 10175 ms   fast   not grouped

  Identical counter, 3.5× spread, and the boundary sits exactly on
  `grouped the chart's shapes` giving way to `the re-read before grouping came
back empty`. So the flip follows the COLLECTION, which is what the heading
  said before the previous round muddied it.

  **The slide reading is not merely unsupported now, it is contradicted.** That
  round's deck ends with slide 257 — the slow one — holding **3** shapes, and
  slides 258 and 259 — fast — holding **25** and **52**. The fewest-shapes
  slide is the slow one, which is the opposite of the crowding account. Round
  `957aca0` made the two look joined only because its deck happened to put the
  shared slide and the surviving groups on the same charts; this round's deck
  separates them and they come apart cleanly.

  **The counter itself checks out where the comparison is fair.** `onSlide`
  against the deck inventory: slide 258 reads 24 against 25, slide 259 reads 45
  against 52 (the rasterise arms drew 7 more after the last reading). Slide 257
  reads 34 against 3 and is NOT a discrepancy — the inventory counts top-level
  shapes, and grouping collapses twenty-four into one. Only compare these two
  numbers on a slide whose charts stayed ungrouped.

  **The friction is astonishingly local.** That round logged 15 errors in 818
  seconds and **14 of them belong to `same scale across the deck`** (14
  idRefusals, 4 emptyReReads); one to `stop a run part-way`; every other
  scenario reported `errors: 0, idRefusals: 0, generalExceptions: 0,
emptyReReads: 0`. This host is not generally degraded — it fails in exactly one
  shape, and the per-scenario friction delta is what makes that visible at a
  glance instead of by counting error lines.

  **Reproduced exactly on 2026-08-09, and it makes the scenario's SCORE
  readable.** In the `40b5e44` round, `same scale across the deck` redrew its
  eight charts between 134.7s and 342.4s. Charts 1-3 ran at 16-18s per batch;
  chart 4 straddled the boundary; charts 5-8 ran at 0.6-5s. The five grouping
  failures in that window land at 291.8, 303.8, 316.5, 328.9 and 341.1s — one
  per fast chart, none for a slow one, each the full chain (`the re-read before
grouping came back empty` → `not grouping` → `tagging failed` → `settle pass:
could not repair any`). The verdict was **3 of 8**, which is the three slow
  charts.

  So this scenario's score is not a variable defect. It is a measurement of WHEN
  the host flips regime: 3 of 8 and 4 of 8 across rounds is the flip landing one
  chart earlier or later. Read it that way, and stop treating a move from 4 to 3
  as a regression.

  **Round 8 (`d812d0c`) decomposes the score exactly, and the degradation turns
  out to have THREE stages rather than two.** Per-chart wall clock:

      charts 1-3   ~35s each   SLOW    clean: grouped, tagged, done
      chart  4       8.1s      fast    grouped; tag refused; SETTLE REPAIRED IT
      chart  5       8.3s      fast    grouped; tag refused; settle refused too
      charts 6-8   ~8.2s each  fast    NOT grouped; tag refused; settle empty

  4 of 8 = the three slow charts plus the one the settle rescued.

  **Round 9 (`448ffc6`) narrows that: the three stages were one round's shape,
  not a structure.** Same scenario, same build family, and the two boundaries
  land together — charts 1-3 slow and grouped, charts 4-8 fast and NOT grouped,
  five settles all `settled: 0, lost: 1`, score 3 of 8:

      charts 1-3   ~35s each   SLOW    grouped, clean
      charts 4-8   ~8.4s each  fast    NOT grouped, all lost

  So the speed flip and the grouping death are INDEPENDENT and may or may not
  coincide: a chart apart in round 8, together in round 9. What survives both is
  narrower and more useful than either — **the score counts the charts that
  GROUPED and kept a tag, and grouping is what the collection read decides.**
  Speed is a symptom sitting near that boundary, not the thing that sets it.

  **And the settle repaired a chart for the first time on record** —
  `settle pass: repaired every config tag the drawing context lost, charts=1
settled=1 lost=0`. Every previous observation was `settled: 0, lost: 1`, and
  this file said so. The recovery works.

  **What it needs is NOT the chart having been grouped — that reading is
  corrected by round 16 (`1fc21b9`, 2026-08-11), which reproduces round 8's
  three stages exactly:**

      charts 1-3   ~34s each   SLOW    grouped, clean
      chart  4       9.6s      fast    grouped; tag refused; SETTLE REPAIRED IT
      charts 5-8   ~8.5s each  fast    NOT grouped; tag refused; settle empty

  Score 4 of 8 = three slow-clean plus one settle-rescue, and the rescue lands
  on chart 4 in both rounds. So the three-stage shape is not one round's
  accident after all; round 9's two-stage round is the variant.

  The correction is what the settle's rescue turned on. Chart 4 had been
  grouped, and its repair still did not go through a shape id: the call that
  landed it was `settling the config tag on a shape found by NAME`, because the
  id readback had been refused for chart 4 too. What separates it from charts
  5-8 is only that its fresh-context collection read ANSWERED, where theirs
  reported `the settle's re-read came back empty`. Grouping is correlated
  because the same collection decides both — it is not the mechanism. The settle
  pass now carries `withId` beside `charts/settled/lost` so the two routes are a
  count rather than a hand pass over `afterAnswering` strings.

  **The flip's INDEX is the stable thing.** Charts 1-3 slow and 4-onward fast in
  all three rounds that decomposed — 8, 9 and 16. What varies between a 3 and a
  4 is one binary event, whether the settle catches chart 4, not a boundary
  sliding along the scenario. Read a move from 4 to 3 as that coin, not as a
  regression and not as the flip landing earlier.

  **The scenario reports the index itself now, and stops once it has it.**
  Every one of those decompositions was done by hand, from a log where all
  eight charts say `index: 0` — within a single-chart update each chart IS the
  first one. Two changes end that:

  - `traceAbout({ chart: "4/8" }, …)` in `src/core/trace.ts` attaches a subject
    to every line written inside a span — draw batches, grouping, tag writes,
    the settle, errors. Merged UNDER the payload so a call site naming the same
    key still wins, restored on the way out including on a throw, and the deck
    path uses the same mechanism for `item: "3/38"`.
  - `rescaleShouldStop` ends the scenario once TWO charts in a row have lost
    their config, and the verdict carries `the host flipped at chart N of M`.
    Two rather than one because the first degraded chart is the one the settle
    may still rescue, which is exactly the 3-versus-4 coin. On round 16's data
    it would have stopped after chart 5 and saved ~38s of an 818s round.

  **On a healthy host the stop never fires**, so nothing is skipped and the
  scenario still proves every chart in the deck takes the shared scale. That
  property is what makes the shortcut safe, and it is guarded from both sides —
  one test that it fires when the host refuses twice, one that it does not fire
  when the host behaves (proven against a build whose rule always says stop).

  **Round 17 (`4feb5be`, 2026-08-11) scored 7 OF 8 — the best this scenario has
  ever recorded — and it separates two things this file had been treating as
  one.** Charts 1-3 slow (17.0, 17.7, 18.5s per batch) and clean; chart 4 fast,
  grouped PARTIALLY, tag refused, settle repaired it by id; chart 5 the only
  loss; charts 6-8 fast, ungrouped, and chart 7 rescued by the settle again.

  So the SPEED flip is still after chart 3 — four rounds, four times — while the
  config-loss index was 5 and only one chart was lost at all. Those are not the
  same boundary and this file's "the flip's index is stable" paragraph was about
  the first. Read the speed flip as the stable one; the losses are a coin the
  settle gets to toss each time, and this round it won six of seven.

  **The index is NOT stable, and 2026-08-12 (`393e6e4`) is the counter-example:
  the speed flip landed after chart 2.** Five rounds had put it after chart 3,
  which was enough to write "four rounds, four times" above; the sixth moved it.
  What did not move is the thing it tracks — grouping — and that round is the
  one where the two can finally be told apart (see `FAST IS THE BROKEN MODE`).
  So read the flip as following the collection, and read its INDEX as wherever
  the collection happened to give way that round. Nothing should key on 3.

  Same round scored **5 of 7** with two settle rescues, both `withId: 1`, which
  is the most this pass has ever repaired in one round. The stop fired at chart
  6 of 7 after two consecutive losses, as designed.

  **The verdict's wording was wrong on its first outing and is fixed.** One lost
  chart printed `the host flipped at chart 5 of 8` — the word three earlier
  rounds use for a host that degraded and never came back, on the best round on
  file. `rescaleLossNote` now reserves "flipped" for the two-consecutive case
  the scenario actually stops on, and says `1 of the 8 charts redrawn lost its
config, the first at chart 5 of 8` otherwise. A verdict and the decision it
  reports on must not be able to disagree.

- **The web host does not LIST the shapes a run just added.** The finding the
  extended log produced, and the one everything else downstream hangs off:
  `the re-read before grouping came back empty index=0 drew=24`, four times in
  one round, on slides that had just taken 24 shapes each. The probe says the
  same thing three more ways in the same run — `shapes-items-count-honest:
unreadable`, `shape-proxy-survives-one-sync: unreadable`,
  `shapes-items-via-positional-slide: not-listed`. So `refreshed=0` is the EMPTY
  case, not a throw and not a failed match, which is exactly why that trace was
  split three ways. The consequence chain is the whole of `same scale across the
deck`: nothing to group with, no fresh tag target, so the tag goes through a
  `created` proxy, the host refuses it, and the settle has nothing better to
  offer. Slides are fine — `getitem-durable-slide: yes`,
  `shape-add-positional-slide-proxy: yes`. It is the shape COLLECTION that will
  not answer. **Settled 2026-08-08, twice over.** The parent handle was never
  the problem: `shapes-items-count-honest` and its positional partner
  `shapes-items-via-positional-slide` both answered `short-0` in the same run,
  so renaming a readback positionally buys nothing. And the collection is not
  merely quiet, it is inconsistent with itself —
  `getcount-populates-same-sync: yes, value=8` next to `items=0`, same host,
  same minute. The count is right and the list is empty. That is why every
  collection read in this repo has to be corroborated against the slide's own
  count (`slideShapeNames`) rather than believed.

  **And the corroboration has a hole, found on 2026-08-12 (`957aca0`): both
  signals can agree AT ZERO, and both be wrong.** `explode a degraded picture`
  read the slide it had just drawn on and reported `the collapse added 0 shapes
(none) — the slide went from 1 to 0`, while the deck inventory taken at the end
  of the same run shows that slide holding one shape named `PowerChart`. The
  chart never moved. `slideShapeList` did exactly what this paragraph asks —
  loaded `items/id,items/name`, called `getCount()`, compared them — and got
  zero from both, so its return value could not carry the doubt.

  The consequence is that **an empty read is not evidence of an empty slide**,
  and every caller has to decide that for itself from something outside the
  collection. Two do now. The scenario has a floor argument: an update that
  just handed back a target cannot have drawn onto a slide holding nothing.
  `slideHoldsOnlyChart` — the gate that authorises DELETING the user's slide —
  uses `shapesDrawnOn`, which splits along this pathology's own line: the host
  does not list the shapes a run just ADDED, so a zero on a slide this run drew
  on is the refusal, and a zero on a slide it never touched is an answer. That
  keeps the bare-slide case the swap fallback exists for while refusing the case
  that would replace a logo, a title and the speaker notes with a generated
  slide carrying none of them.

  `faults.slideReadsEmpty` is the fake being this host, with two arming times.
  The later one (`"after-a-picture"`) is not decoration: armed from the first
  read it starves `probeCharts`, the scenario skips for want of a chart, and the
  guard passes against the unfixed build.

  **A consequence of that swap worth knowing: `picture-then-shape-read` came
  OFF `KNOWN_DIVERGENCES`.** The new sheet's row for it is `yes`, which is what
  the fake says, so the gate correctly reports no divergence — but the host gave
  `unreadable` on the same round's other two passes. Nothing is lost, because
  that entry's content (office-js#5022, and `drawDemoItem` performing exactly
  that sequence) moved into `UNSTABLE_ANSWERS`, which is the right table for a
  coin. Do not read the absence of a divergence there as the host having
  settled.

  **CONFIRMED ON LIVE DATA the round after that (`ee1741e`, 2026-08-12):**
  `the slide read back EMPTY right after the collapse — not believing it
{slide: 261#2230304510, was: 1}`. The pathology recurred, the guard caught it,
  and the scenario continued to the config round-trip instead of claiming data
  loss. The round before had passed WITHOUT entering the branch, which is why
  that pass proved nothing and this line does.

  **And the same scenario then failed the same way one step later, which is the
  more useful finding.** Its verdict read `the picture vanished while being
exploded back to shapes`, and the deck inventory from the same run shows that
  slide holding one shape named `PowerChart`. Nothing vanished: the update had
  just logged three `InvalidParam passed to GetItem(id)`, so the host would not
  name the picture and the redraw could not work on it. `updateLossNote` reads
  the run's own `idRefusals` either side of the call and says which fact it
  means. Still a failure — but "the add-in deleted a chart" and "the host would
  not answer for it" send a maintainer to opposite ends of the codebase.

  **The previous round (`393e6e4`) passed the scenario and was NOT evidence the
  fix works.** `explode a degraded picture` came back `collapsed to a picture and
exploded back, config intact` — and the trace carries no `the slide read back
EMPTY right after the collapse` line at all, so the host answered the
  collection normally and the new branch was never entered. A green scenario on
  a host having a good minute says nothing about a guard for its bad ones; what
  says the fix works is the test that goes red without it. Worth stating because
  the temptation to read the next round as vindication is exactly how this
  project has misread a round before.

- **The battery's one recurring stall now says what it is.** `a selected shape
survives an insert` stalled its first draw batch in four of the last five
  rounds (`957aca0`, `ee1741e`, `89675b6`, `47a80c8`) after passing eight
  running before that, and every round reported it with the runner's generic
  "the host got in the way" — a specific, repeating observation thrown away each
  time.

  What the note may claim is bounded by the round files, and the tempting answer
  is dead: every one of those stalls reads `afterAnswering: "selecting a shape",
idleMs: 2-3`, and `selecting a shape` sits in the SURVIVING population in all
  four of the same rounds, because `edit the chart the user selected` draws
  after it and lands. So the preceding call is not the variable — the same way
  it was not for any earlier candidate. What is left, and all the note says, is
  the one way this draw differs from every other in the battery: it is the only
  one made with a selection STANDING, which is #2775's repro and exactly what
  `dropShapeSelection` exists to avoid.

  **A control arm was built for it and then removed, which is the more useful
  half.** Matching this draw needs a same-size chart on the same slide; every
  slot is allocated; sharing one broke `every chart the battery draws has a slot
to itself`, and widening the band broke the degradation grid's fit. Two
  invariants pushing back is the codebase saying the change is wrong-shaped —
  and the product does not turn on the answer anyway, since the add-in already
  drops the selection before drawing. So it is an observation, stated once,
  rather than an experiment worth deforming the battery for.
  `faults.stallDrawAfterSelect` is the fake being this host for it.

  **MEASURED across every round on file, and the effect is RECENT — which the
  first version of this note implied it was not.** This draw against `edit the
chart the user selected`, which selects a shape, DROPS the selection, then
  draws:

      all 17 rounds        held 10 ok / 7 stalled    control 14 ok / 3 not ok
      since `957aca0` (8)  held  2 ok / 6 stalled    control  8 ok / 0

  Over the whole history they do not separate, and in the earliest rounds the
  pattern ran the OTHER way — the control failing while the held draw passed.
  Restricted to the last eight rounds they separate sharply.

  So this is not "drawing while selected stalls this host", which is what a
  four-of-five rate looked like. Something changed around `957aca0`, and two
  readings fit: the host is different, or the slide-spreading landed in that
  build and moved which slide this scenario works on. It takes `found[0]` and
  not `leastLoadedChart`, so it did not change directly — but what the other
  scenarios leave on which slide did. Do not build on either reading.

- **A self-test scenario that ends the run costs the whole report, even last.**
  `the chart is actually visible` ran dead last precisely so its crash could not
  take other scenarios' verdicts — and it still cost every round, because the
  report is written when the battery RETURNS. Four rounds, four builds, always
  within a step or two of `adding a scratch slide`, never once a verdict. It is
  `pickedOnly` now. "Last" protects the verdicts; only "out of the routine list"
  protects the report. Whether the scenario or the ten minutes in front of it is
  the killer has never been separated — running it alone is that experiment.
- **Do not rasterise a slide the run just added. It kills PowerPoint.**
  `getImageAsBase64` on a freshly-added slide has now failed on the web FIVE
  distinct ways: `GeneralException` at `SlideCollection.getItem`; taking the
  call and silently producing nothing; never answering the sync; sitting on the
  full ninety-second readback budget; and — 2026-08-08 — killing the tab
  outright. `rasteriseTimeoutMs` is twenty seconds, capped by the readback
  budget so a test can still shorten it, and it does not help against the fifth:
  a timeout cannot save a process that is gone.

  **The isolating experiment ran on 2026-08-08 and it is the scenario, not the
  load.** `the chart is actually visible`, picked alone on a fresh deck, was
  reached at 61.5s with only its two inserts in front of it — against ten
  minutes and nine scenarios in the four rounds before. It added a scratch slide
  at 61.5s, logged `rasterising the empty slide` at 61.8s, and the tab died.
  Those same two inserts head every routine round and kill nothing, so elapsed
  time and volume of drawing are both out. Four rounds of "it crashed again"
  said nothing; one picked round said this.

  `chartIsVisible` no longer takes a scratch slide at all — it does its
  before/after on a slide the run added EARLIER, which drops the scratch add,
  the fresh-slide rasterise and the delete (that delete had killed a round of
  its own).

  **And the re-run closed the last reading.** On `e49cca8` the step named
  `rasterising a slide that already existed` answered without incident, the
  scenario drew its 24 shapes, rasterised again and returned a verdict — the
  first in six rounds. So this host rasterises perfectly well; it was the FRESH
  slide, both directions now measured rather than inferred. A slide the run just
  added remains the worst surface this host offers, and nothing may rasterise
  one.

  **It PASSED on `c7d91d5` and is routine again** — `drawing the chart changed
what the slide looks like (10064 → 15652 bytes)`, through PowerPoint's own
  rasteriser. That is the first time this project has confirmed a chart it drew
  is VISIBLE anywhere but in a human's eyes. Sharing a slide is the price of
  never rasterising a fresh one, and the bill came the same round: two full-size
  charts drawn over each other, which a rasteriser reads fine and a human calls
  broken. The scenario's chart is 30% of the slide in the bottom-right corner
  now — this battery leaves its slides in the deck for someone to look at, so
  "the measurement still works" is not the bar.

  **The three passes on record are not equally strong, and until 2026-08-11 the
  verdict could not say so.** 10064 → 15652 is +55%; 15704 → 16580 is +5.6%;
  round 16's is 14868 → 14976 — **a hundred and eight bytes, 0.7%** — and all
  three read identically in the round file. The gate asserts the two renders
  DIFFER, which a re-encode satisfies on its own. The verdict now carries the
  delta and the share, and flags anything under `THIN_VISIBILITY_RATIO`.
  Deliberately NOT a failure: a change is still a change, and turning a thin one
  red would fail a round on a judgement no measurement here supports. Reporting
  it lets a reader tell 0.7% from 55% without opening the file, which is all
  that was actually missing.

  **And the very next round made the flag look prophetic. `14988 → 15096` is
  +108 bytes AGAIN** — the same delta as `14868 → 14976`, from a different
  starting size, two rounds running. A chart appearing in a rasterised slide
  does not cost the same 108 bytes twice by coincidence; something constant is
  moving and the chart is not visibly in it. Nothing has been changed on the
  strength of two samples, and the verdict still passes, but a third +108 should
  be read as this gate measuring a header rather than a chart. Watch for it.

  **The third arrived on the next round: `14856 → 14964`, +108 again.** Three
  consecutive rounds, three different starting sizes, one delta. That is no
  longer an anecdote, and by this paragraph's own criterion the gate is
  measuring something constant.

  A length cannot tell a header from a picture, so the verdict now reports WHERE
  the two renders first differ and how many bytes differ at all
  (`renderDifference`). An encoder header, a timestamp or a counter differs
  early and in a handful of places; a chart drawn into the image differs across
  the body of the data.

  **The next round KILLED the header reading, and the answer was not what this
  paragraph predicted.** `53ec985` reported `14864 → 14976 bytes, +112, 0.8%,
first differ at 1% in, 14540 byte(s) differing` — **97% of the image differs.**
  The near-constant length delta is a coincidence of compression, not a constant
  header, and three rounds of `+108` had been read as a mechanism when they were
  a red herring. A length is not a measurement of difference; this file spent
  three rounds treating it as one.

  **Two readings survive, and a before/after pair cannot separate them**: the
  chart is in the picture, or this host's rasteriser does not produce the same
  bytes twice. Both make a drawn chart and an untouched slide look equally
  different. So the scenario now takes a CONTROL — the same slide rasterised
  twice with nothing drawn between — and the verdict says
  `two renders of the UNCHANGED slide also differed, so this proves NOTHING
about the chart` when the rasteriser is unstable. If that fires, every "the
  chart is visible" verdict on record is worth nothing; if the control comes
  back clean, this gate finally means what it has been claiming for six rounds.
  One extra call on a slide the run has already rasterised safely.

  **The control ran on 2026-08-12 (`957aca0`) and came back CLEAN, so the gate
  means what it says.** No `proves NOTHING` caveat in the verdict: two renders
  of the unchanged slide were identical, this rasteriser is deterministic, and
  the difference a drawn chart produces is the chart. Six rounds of verdicts are
  retrospectively worth what they claimed.

  **And the same round explains the thin deltas without any of the mechanisms
  this paragraph guessed at.** It reported `9052 → 10864 bytes, +1812, 20%` —
  twenty percent, against the 0.7-0.8% of the three `+108` rounds — and the
  variable is the SLIDE, not the rasteriser and not the encoder. The `+108`
  rounds drew onto the battery's overloaded slide, where a chart is a small
  addition to a crowded picture; this round drew onto a slide the spreading fix
  had kept light. A thin delta is a crowded slide, and the flag is worth keeping
  for exactly that reason: it reports how much of the picture the chart is, on a
  gate whose whole subject is whether a human would see it.

  **But it is a render, not a screen — office-js#6498, triaged 2026-08-12.**
  On the web an inserted shape can appear in the slide PREVIEW and not in the
  main view, and `getImageAsBase64` renders precisely that preview. So a pass
  here means "the chart is in PowerPoint's own render of the slide", which is
  strictly weaker than "a human looking at the slide would see it", and the
  verdict says the weaker thing now instead of the stronger one. No add-in API
  reads the canvas, so this cannot be closed by measuring harder — only by
  somebody looking, which is one more reason the battery leaves its slides in
  the deck. The report is about slide MASTERS and ours are slides, so the match
  is on the mechanism rather than on the repro; that is why it bounds a claim
  instead of predicting a failure.

- **The showcase deck's SLIDE XML is deterministic; its BYTES are not.** Always
  commit the regenerated deck with the code that changed it — but read a
  modified `examples/showcase.pptx` in `git status` as no evidence at all that
  the deck moved. Two builds of identical code differ at byte 11, because
  `docProps/core.xml` carries a wall-clock build timestamp and the zip carries
  entry timestamps. Every file under `ppt/` is byte-identical, which is why the
  CI gate compares `ppt/slides/slide*` and says so at the step.

  This line used to read "the showcase build is byte-deterministic", and that
  cost time twice in one session: once nearly committing pure timestamp churn as
  though it were a real diff, once treating a modified file as a regression to
  hunt. To check whether a change actually moved the deck, extract both zips and
  diff `ppt/` — the raw file comparison answers a different question and always
  says yes.

      git show HEAD:examples/showcase.pptx > /tmp/old.pptx
      npm run build:lib && node scripts/build-showcase.mjs
      mkdir -p /tmp/a /tmp/b
      (cd /tmp/a && unzip -qo /tmp/old.pptx)
      (cd /tmp/b && unzip -qo examples/showcase.pptx)
      diff -r /tmp/a/ppt /tmp/b/ppt      # this is the answer

  And diff `ppt/`, not the whole extraction: `docProps/core.xml` differs on every
  build, so diffing the extraction root reports a change that is never real.

- The pane rebuilds `ChartConfig` from UI state: new **decoration** keys
  round-trip automatically; new **top-level** config keys need a state field
  or the `state.extras` passthrough in `src/taskpane/app.ts`. **Forgetting is a
  CI failure now, not a silent one** — `test/pane-state.test.ts` reads the key
  list out of the `ChartConfig` interface and carries every one of them through
  import → export, so a new field fails there until it is either given a sample
  or declared as owned by a pane control. Miss it and the key was dropped on
  import and destroyed on the next re-save, in a pane that looked like it had
  loaded the chart fine.
- All sample/showcase data is invented dummy data (`src/core/samples.ts`,
  `scripts/build-showcase.mjs`) — keep it that way; the repo will go public.
- GitHub MCP `actions_list` and `list_pull_requests` responses exceed the token
  cap — parse the saved JSON file with python instead of reading it. In
  `list_pull_requests` the `merged` field is unreliable; read `merged_at`.
- **Squash-merge hides ancestry.** A merged `claude/*` branch is NOT an
  ancestor of `main`, so `git branch --merged` reports nothing and
  `--contains` proves nothing. Confirm a branch is safe to delete from its
  PR's `merged_at`, not from git.
- **Prune after a merge** (`git fetch --prune`, or set `fetch.prune=true`).
  GitHub deletes the branch on merge, but the local `origin/claude/*` tracking
  ref survives pointing at the pre-merge head. The stop hook picks that stale
  ref as its baseline and reports the squash-merge commit as an unverified
  commit "on your branch" — which it is not, and amending it as the hook
  suggests would force-push a rewrite of `main`.
- Object lookups keyed by a config string must use
  `Object.prototype.hasOwnProperty.call` — a pattern/colour/marker named
  `__proto__` or `constructor` otherwise reaches `Object.prototype` and either
  crashes the renderer or gets CALLED. **Read the guard's current reach from
  `grep -rl hasOwnProperty src/ skill/` rather than from a list here** — this
  sentence named four files, then five, and was wrong by three within a week,
  which is the same drift the backlog and `UNSTABLE_ANSWERS` paragraphs warn
  about. Apply it to any new table.

  Two things a grep cannot tell you, so they stay written down:

  - **Check the write side too.** The saved-templates table in `app.ts` was
    missed for months: `all[name] = value` where name is `__proto__` hits the
    inherited setter and re-parents the object instead of storing, and the entry
    then vanishes with nothing said. That one is guarded by a null prototype,
    which fixes both directions at the root instead of at each call site.
  - **A table does not have to be keyed by a CONFIG string to be reachable.**
    `buildCheckbox`'s glyph and colour tables are keyed by a `CheckState` union
    and were the seventh instance of this class — the pane feeds them from a
    `<select>`, but `src/index.ts` exports the builder, so the value is whatever
    a library caller passed. Every public export in `src/index.ts` is a boundary
    of the same kind as the JSON box.

- `Date.parse` is far looser than a date cell: `parseDateToken` therefore
  gates on shape (date punctuation + month/weekday words only) before parsing.
  Don't route new cell input around it.
- **`dist-lib/` is a build artifact, and `skill-scripts.test.ts` runs against
  it.** A local full-suite run after a change to `src/core` tests the OLD core
  through that path while CI builds fresh — so the suite goes green locally and
  red in CI, for a real regression. Run `npm run build:lib` before trusting a
  local run that touched core.
- **There are THREE colour sinks, and they are separate code on purpose:**
  `src/core/color.ts` (preview), `skill/scripts/pptx-paint.mjs` (headless
  pptx), `officeHex` in `powerpoint.ts` (the live add-in). The same bug has now
  been found in all three independently — each was fixed when a sweep aimed at
  _that_ renderer found it. Change one, check the other two.
  **Done deliberately once, on 2026-08-09, and it paid immediately**: probing
  all three with the same six inputs found the same defect in all three in one
  pass, rather than one at a time over three sessions. The defect: `parseFloat`
  is looser than the regex feeding it — `[\d.]+` matches a bare `.` and
  `parseFloat(".")` is NaN — so `hsl(., 50%, 50%)` **threw** (the hue sector
  table has no NaN entry, so destructuring `undefined` blew up) and
  `rgb(., ., .)` returned `#NaNNaNNaN`, which in the pptx sink is not the six
  hex digits that path's own security note requires. Both are now enforced at
  the root — `rgbToHex` for the preview, an exit check in `hex()` for pptx —
  rather than per branch, and `officeHex` inherits the first.

  **Done deliberately a second time on 2026-08-12, and the useful part is what
  it found: a DISAGREEMENT, not a crash.** Running 48 paint strings through the
  preview sink and the pptx sink and diffing the answers turned up exactly two
  real divergences, both the same fact — a BARE hex string, `#` omitted. The
  pptx sink has always taken those (it is OOXML's own spelling,
  `<a:srgbClr val="AABBCC"/>`); the other two did not. So `AABBCC` as a series
  colour drew the real colour in the skill's deck and mid grey in the preview
  and the live add-in, and pasting a hex out of a brand guide without its `#` is
  the commonest way to reach it.

  The SVG sink needed more than permission. `AABBCC` is a run of LETTERS, so the
  allow-list's named-colour arm passed it through unchanged and the markup
  carried `fill="AABBCC"` — not a valid SVG paint, so the attribute is ignored
  and the shape renders black — while `4e79a7` has digits, matched no arm, and
  fell back to black explicitly. Two spellings of one colour, both black. It is
  normalised now (the `#` is put back) rather than allowed, so the value still
  goes through the allow-list instead of around it.

  Two things worth keeping from how this was decided. **The other 16
  disagreements were not defects** — the grey-vs-black fallback for an
  UNREADABLE paint is deliberate on both sides and `hexOr` exists for the
  background case; reading that comment before "fixing" it is what stopped a
  regression. And **the ordering claim is checked rather than argued**: bare hex
  is tried after the name table, and `test/color.test.ts` asserts that no CSS
  colour name is a valid bare hex string, so the ordering is a property of the
  code rather than of what happens to be in the table.

  The comparison itself is the reusable part — `toHex6` against `hex` over a
  corpus of paint strings, with the answers diffed rather than merely checked
  for throwing. `test/color.test.ts` holds that corpus and both sinks read from
  it, so a form added for one is checked against the other.

  **Done a third time on 2026-08-13, for DASH ARRAYS, and it is the sharpest
  result of the three: 10 of 16 inputs disagreed.** The cause is one line each
  side. The two PowerPoint sinks guarded on `dash` being TRUTHY (`if (!s.dash)`,
  `n.dash ? … : {}`) where SVG guarded on its CONTENT — so an array carrying no
  positive length drew a SOLID line in the preview and a DOTTED one in both
  decks. `[]` is the sharp case: truthy, produced by ordinary caller code like
  `dash: xs.filter(…)`, and `dashKind` answered `dot` for it under a test that
  called that "a harmless default". It was not.

  **The fix is to make the SHARED FUNCTION total rather than to guard three call
  sites** — `dashKind` returns `"none" | "dot" | "dash"` and takes `unknown`.
  Every sink already imported it, so nothing had to be threaded and there is no
  second predicate to drift. Same lesson as `horizontalLegendFits`: a shared
  function is the only version of "these two agree" that cannot rot. Do not
  re-guard on `n.dash` being truthy at a call site.

  Two things worth keeping. **The corpus found a case reasoning would not**:
  `[-1, 4]` HAS a positive length, so "some positive value" called it a dash —
  but a negative entry is an ERROR in SVG and the whole attribute is dropped, so
  the preview draws it solid. The rule had to match what SVG does, not what
  sounded right, and only running both sinks over one corpus said so. And **the
  first metric was wrong in the direction this file keeps warning about**: the
  probe's "is it dashed" predicate tested `!s.includes("-")`, which matches
  `stroke-dasharray` itself, so it reported the two CORRECT cases as
  disagreements. An approximate metric invents defects as well as missing them —
  it was fixed before anything was read off it.

- **A `string` in the types is not a string in the file someone pasted.** A
  config arrives from the JSON box, a saved template, a shape tag written in
  another deck, and the skill's caller. `categories: [2023, 2024]` and
  `title: 2024` are ordinary things to write and both used to crash. Coerce at
  the boundary (`normalizeConfig` / `normalizeData` for config,
  `paintText` / `xmlText` for text) rather than at each consumer.

## Out of scope (decided, don't revisit without the owner)

- think-cell's "impossible trio" for a sandboxed add-in: live Excel data
  links, in-canvas drag manipulation, and the slide-layout engine — think-cell
  memory-patches Office binaries for these.
- True geographic maps and 3D surface charts (no freeform paths; the deck this
  project follows argues against 3D anyway). Built instead: tile-grid
  cartograms (`tilemap`) and heatmaps.
- **Guessing at an ambiguous slash date.** `03/01/2026` is 3 January in Europe
  and 1 March in the US, and nothing in a cell says which; `parseDateToken`
  refuses it rather than picking one, and refuses only when both numbers could
  be a month (`01/15/2026` and `03/03/2026` still parse). Decided by the owner
  on 2026-08-12 with the cost on the table — a US author who writes
  `03/01/2026` for 1 March gets an empty cell — because the alternative is a
  Gantt that draws perfectly and is two months wrong, which is what it used to
  do. Reading them day-first, or from `style.locale`, were both offered and
  declined: they pick a winner where the data does not. This is the same rule
  `numericValue` applies to a European "1,5" against an American "1,500", and
  it is settled, not a default awaiting a better idea.

## Backlog

`docs/BACKLOG.md` is the single curated backlog (researched candidates with
feasibility/priority, plus a rejected list — don't re-propose those). Items
graduate from there into PRs and are removed when shipped.

Do **not** restate the backlog's contents here — a count in this file and a
list in that one drift the moment either changes, and this paragraph spent
several releases claiming the open candidate was image-output render mode
after that shipped. Read `docs/BACKLOG.md` for what is open; read the README
feature table and git for what has shipped.

## Pending / user-gated

`docs/PUBLISHING.md` is the go-live runbook. Phases 0, 1 and 4 are done: the
repo is public and protected, the add-in is hosted at
<https://ssf-chart.struktureretsundfornuft.dk/>, and **v0.2.0** ships the
prod manifests, the skill zip and the showcase deck.

What is left needs the owner, not the agent:

- **Phase 2 — keep validating in real PowerPoint.** It HAS run now: three
  answer sheets, several run logs and decks, and the fixes they produced are in
  `git`. `test/fixtures/host-answers-web.json` is one of those recordings and is
  what the CI contract gate diffs against. So the standing question is no longer
  "does any of this work" but "what does the newest build answer".

  Both experiments have now run. _the chart is actually visible_ PASSED on
  `c7d91d5` and is routine again. _what makes a long run slow down_ ran on
  `25407ed`, picked alone, and KILLED THE TAB at 26.9 seconds with two scenarios
  behind it — announcing itself and then nothing, which is exactly the state
  `chartIsVisible` spent four rounds in. Two calls fit and the log cannot choose
  between them: taking a scratch slide, or drawing ninety-six shapes onto one
  the run had just added. It is instrumented now (`degradation step`), so the
  next picked round names the call instead of the scenario. Do not reason about
  which; that is what the steps are for. `PENDING_QUESTIONS` in
  `scripts/host-baseline.mjs` is the authoritative list of unasked probe
  questions — **read the table rather than a count here**, which is the rule
  the `UNSTABLE_ANSWERS` paragraph above already states and which this sentence
  broke within a day of being written: it said "nine down to one" and a probe
  added on 2026-08-09 (`binding-names-shape-later`) made it two.
  _edit the chart YOU click_ came off this list on 2026-08-08 — it PASSED.

  **Bindings are not the way out of the id refusals — asked and answered on
  2026-08-12 (`957aca0`).** The idea was that a binding made from the live Shape
  proxy inside the batch that created it needs neither an id round trip nor a
  collection read, so `settleAndTagChart` would have a route that does not go
  through the `ShapeCollection.getItem(id)` every 5010 comes from. The host
  answers `commit-threw`: it REJECTS the batch carrying the binding
  (`ErrorPointer`). It counts as an answer rather than as one more bad minute
  because the probe's control arm committed the same batch WITHOUT a binding
  seconds earlier and it landed — which is exactly what that arm was added for,
  after an earlier round produced the same signal with nothing to compare it to.
  The reading now lives in `KNOWN_DIVERGENCES`, where the fake is deliberately
  left saying `yes`: nothing here makes a binding, so there is no caller to
  protect, and the entry's job is to be what the next person reaching for
  bindings finds first.

  **A never-asked answer is not an answer, and the gate used to think it was.**
  `no-scratch-slide` and `no-scratch-shape` mean the run could not put the
  question. The contract gate read their absence from `differ` as agreement and
  their presence in the sheet as an answer, so the first sheet carrying them
  demanded that two `KNOWN_DIVERGENCES` entries be deleted and a
  `PENDING_QUESTIONS` entry retired — deleting knowledge on the strength of a
  setup failure. Both directions now treat `notAsked` as unknown.

  **Nothing is owed to the owner right now.** The manifest re-install he was
  asked for was done on 2026-08-06, and nothing since has touched a manifest —
  do not ask again unless a PR actually changes one. What to ask him to click is
  written down: "The standing test run" in `docs/PUBLISHING.md`. Don't improvise
  a new one per session, and don't ask for the deck or a screenshot — the round's
  own file has carried both since the deck-evidence change.

  **Eleven rounds have now run, and the battery reports.** The committed fixture
  is whichever sheet last answered at least as much as the one before it — read
  its `build` field rather than trusting a commit named here, because this
  paragraph has already been wrong about that once. `layouts-readable` and
  `slide-layout-readable` both answer `yes`, which retires office-js#4906 and
  #3826 as exposures on this host and this deck.

  **Both picked-alone experiments ran and answered, and neither is pending.**
  _the chart is actually visible_ is routine again and has PASSED twice
  (`c7d91d5`, and again on `1cd7ea3` at 15704 → 16580 bytes). _what makes a long
  run slow down_ has MEASURED — twice, on two builds — and the quadratic
  per-slide cost written up above is its result; it is `pickedOnly` and does not
  need running again unless that number is in doubt.

  **The stall investigation is closed as far as candidates go.** Eleven rounds
  eliminated the scenario, the preceding scenario, the tab's age, the idle gap
  before the sync, and the identity of the preceding call — each by a later
  round putting the candidate in both populations. What remains is a host that
  stalls the first sync of a draw intermittently, one or two draws in fifteen,
  with no known discriminator. `does a rasterise poison the next draw` is the
  counterbalanced control that will say "no pattern" until something changes;
  a few rounds of that IS the finding, and is the point at which to stop
  instrumenting and live with it.

- **Phase 3 — activate the Claude skill** (upload the zip on claude.ai).

Follow it phase by phase; retire items from it and from this list as they
complete.
