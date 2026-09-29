# Tests

Vitest discovers every `test/*.test.ts` automatically — there is no manifest to
update. Run the suite with `npm test`, thresholds with `npm run coverage`.

## Naming convention

**Name a test file by the topic it covers, never by the increment that shipped
it.** A chart kind's tests live in `<kind>.test.ts`; a cross-cutting feature's
tests live in `<feature>.test.ts`. If you add a kind or feature, add or extend
the matching file — don't start a `batch-N` / `backlog-x` grab-bag. (The old
`backlog-a…t`, `bug-hunt`, `hunt-*`, `r2-*` files were exactly that, and were
split back out into the topic files below. So were `coverage-core` /
`coverage-branches`, which named a _mechanism_ rather than a topic — a test
belongs with the thing it tests, not with the reason it was written.)

## Where things live

- **Chart kinds** — one file each: `waterfall`, `column`, `line`, `combo`,
  `scatter`, `pie`, `sunburst`, `radar`, `radial-bar`, `boxplot`, `violin`,
  `candlestick`, `heatmap`, `tilemap`, `treemap`, `gantt`, `funnel`, `butterfly`,
  `waffle`, `gauge`, `bump`, `pareto`, `sparkline`, `bullet`, `cascade`,
  `elements`.
- **Cross-cutting features** — `axis-features` / `axis-scale` (axes, scales, log
  floors), `format` / `format-edge` (number & label formatting), `dates`,
  `palette`, `data-sorting`, `legend-layout`, `decor-guards` /
  `decoration-layout` (decoration clipping & anchoring), `value-extent`
  (cross-kind extent/auto-scale invariants), `geometry`, `color` (paint parsing
  & contrast ink), `collide` (label collision resolution), `good-chart*`,
  `frame-fit` (nothing a chart draws leaves the box the chart is, swept over
  every kind × eight frame sizes — the property, because every fix in this area
  had been found by looking at one chart at one size while the same defect sat
  in four other layouts; `elements-fit` is the same question for the non-chart
  elements).
- **Hostile input** — `chart-hostile-input`, now four sweeps rather than one:
  every value a **cell** can hold (huge, subnormal, NaN, infinite, empty,
  degenerate sizes), every **style** field with the wrong type, every
  **top-level key** with the wrong type through both offline renderers, every
  **decoration** key, and eighteen malformed **data shapes**. The cell sweep's
  bar is _termination_ — both bugs it found were loops whose bound came from the
  data. The type sweeps' bar is simply _not throwing_: their bugs were all one
  shape, a `string` API meeting a non-string out of user JSON, and they were
  spread across every layer.
  The live Office renderer gets the same treatment in `office-render`, which
  matters because it is the one no earlier sweep had ever pushed a malformed
  config through — and it held its own copy of the colour bug.
- **Repair planner** — `reconcile` (the rules), `reconcile-fuzz` (four thousand
  generated decks against the invariants that cost a user their work: never
  delete another run's slide, never delete every copy of an item, never act on
  a slide that is not there).
- **Renderers & app** — `office-render` (Office.js against the fake host),
  `web-host` (the same renderer against that host at its WORST — every
  misbehaviour a real PowerPoint on the web has shown us, on at once), `pptx-paint`
  (headless pptx node mapping), `ooxml-ranges` (the sweep beside it: every number
  that mapping writes, checked against the SCHEMA RANGE of the attribute it lands
  in, over every kind × a hostile config matrix — `finiteNodes` covers
  `x="Infinity"`, and this covers the rest, because an out-of-range attribute is
  equally a deck PowerPoint offers to repair and equally reported as a success),
  `svg-render` (SVG node emission — paths,
  polygons, options), `pane-state` / `pane-host-actions` / `pane-widgets` /
  `dom-pane` (task pane — **every boot helper must `await app.harnessReady`**
  since 2026-09-29, because the Testing panel is a dynamic import and its
  listeners attach a microtask after the module finishes; a test that clicks
  `#demo-insert` before that silently does nothing, and an import still in flight
  across the next `vi.resetModules()` wires the PREVIOUS module instance onto the
  shared document. A test that drives `demo-*` also boots with `?harness=1`,
  because since the opt-in gate was flipped that is the only pane those buttons
  exist on), `crashlog` (the record that outlives a run that
  never ends), `templates` (saving and re-picking a chart setup — a whole
  feature that had no tests until one of them turned up a bug),
  `host-probe` (the fake's own frozen answer sheet — what it CLAIMS about the
  host it stands for, so a real PowerPoint can be diffed against it),
  `host-contract` (the other half of that: `FAKE_BASELINE` diffed against a real
  PowerPoint's committed sheet, so a new divergence fails in CI in seconds
  instead of waiting for somebody with the app open),
  `host-regimes` (whether a question that changed its answer MID-round is a coin
  or a function of the host's state — the probe has stamped a regime on every
  sample since #390 and nothing read those stamps mechanically. Its load-bearing
  case is `untested`: with three passes in three regimes "every regime maps to
  one answer" cannot fail, so it is not evidence, and reporting it as `explained`
  would manufacture a finding every time a question flipped. Its reader is
  parameterised over the stamp, so the scratch-slide state added after round 17 is
  read by the same code),
  `selftest` (the in-host
  battery's own logic — that every scenario reports, that a blind deck scan is
  never read as an empty deck, and that a wedged host is attributed to the host),
  `skill*`, `parity`, `snapshots`, `a11y-svg`, `security-*`, `dark-theme`,
  `fuzz`, `hardening`, `degenerate-inputs`.
- **Things CI can check that are not code** — `manifest` (the four add-in
  manifests: Office's version floor, the `<Id>` GUIDs, no localhost surviving
  into a production manifest — pinned offline because the real validator calls a
  Microsoft service and cannot run everywhere), `office-js-watch` (the weekly
  tracker sweep's matching half, including that its `KNOWN_ISSUES` table covers
  every office-js issue the codebase cites and that every watched term is a call
  this repo actually makes), `visible-charts` (the verdict half of the visual
  gate — the rasterising half runs in a real browser and cannot run here),
  `host-history` (the arithmetic behind `UNSTABLE_ANSWERS` and the
  fixture-swap decision — a streak counted from the end, with the words that
  mean "never put" ignored so one bad night does not read as instability),
  `is-main` (whether a tool script was RUN or merely imported — three CLIs
  answered that wrong on Windows and exited 0 in silence, so both platforms are
  driven here rather than whichever one the runner happens to be),
  `runbook` (the controls `docs/PUBLISHING.md` tells the owner to click are
  spelled the way the pane spells them — the starred row said "Run the whole
  round" for weeks and no control has ever carried that text, which cost a
  round).
- **The deck a run produces, audited from its bytes** — `verify-deck` (did
  SSF Charts write what it meant to: slot tags, groups, config parts, shape ids
  unique per slide), `ooxml-validate` (is it a legal `.pptx` at all, against the
  OOXML grammar, plus the one baselined finding it is allowed to have),
  `triage` (joining a saved deck to the run log that produced it). The first two
  catch nearly disjoint sets and both gate CI on `examples/showcase.pptx`.
  `triage` also holds the tests for the 33 POOLED READERS, which moved to
  `scripts/round-pools.mjs` on 2026-09-29 so Stryker could mutate them. They
  stayed in that file rather than moving with the code: they are forty
  interleaved `it`s inside one 1,300-line describe that tests the readers and the
  tool around them from the same brace, and a botched test split is worse than a
  file whose name is broader than one module. What mattered was that they RUN
  under mutation — see `triage-repo.test.ts` below.

## The fake PowerPoint host

`helpers/office-host.ts` is the Office.js double: recording proxies for shapes,
slides, tags and groups, plus a `faults` object of misbehaviours. Every fault in
it was added AFTER a real host taught us the behaviour — a stale shape proxy
refused by `getItem(id)`, a shape collection reading back shorter than it is
without throwing, a refused `addGroup`, a sync that answers minutes late, a
collection read the host never answers at all, a shape whose position stays
unreadable until the load that asked for it lands.

One family of them models the same thing at four levels — the host taking a
load and answering nothing: `unansweredShapeReads` (a whole shape collection),
`unansweredNullChecks` (one `getItemOrNullObject` proxy), `unansweredTagLoads`
(a tag), and `faults.strictShapeReads` (a shape's own `id`/`left`/`top`).
Reading any unanswered proxy is `PropertyNotLoaded` on a real host and was a
plain value here, which is how three self-test scenarios could fail in a real
PowerPoint against a fully green suite. `strictShapeReads` is off by default
only because the fake's shape objects double as the surface tests assert
geometry against — the others are always in force once armed.

`newSlideResolvesTimes` is the one that came from an answer sheet rather than a
crash. PowerPoint on the web resolved a freshly-added slide's id **once** and
refused it ever after — while still listing that same id in
`slides.load("items/id")`, so the slide was plainly there and only the lookup
was broken. `null` means ids stay good (every other host); a number is how many
`getItemOrNullObject` calls each added slide answers before it starts reporting
gone, and `0` means never, starting now. That `0` matters: a count-based version
of the delete test passed against the very code it was written to falsify,
because the fix spends one lookup more than the bug did and the lease outlasted
both.

`selectionWedgesHost` is the odd one out and the newest: it models a call that
is taken and _poisons later ones_. A programmatic `setSelectedShapes` succeeds,
and every selection sync after it then never settles at all — neither resolving
nor rejecting, which is what PowerPoint on the web actually does and the one
failure shape no `catch` can see. Only a bounded wait survives it, so it is the
fault that proves the bounds work.

The newest family came from answer sheets rather than from crashes, and each
models something a real PowerPoint on the web was measured doing: a freshly-added
slide's handle dying at the next sync (`expiringSlideHandle`, unconditional
because it is established); a shape that cannot be NAMED in the batch that
created it (`noIdInCreatingSync` — implied by five questions being the only five
never put); `getItem` refusing a slide this run added (`refuseGetItemOnNewSlide`);
and `addTextBox` deleting the selected shape (`textBoxDeletesSelection`,
office-js#2775). The last two are OFF by default and say so at their definitions:
nothing has established that the build the owner runs still does either, and a
fake that asserts unasked host behaviour turns every real answer into a
divergence.

`syncCostMs` is not a misbehaviour at all — it is a clock. It charges each sync
by a function of `syncsInContext` or `syncsTotal`, which are exactly the two
hypotheses the degradation experiment separates, and without it a measurement
could only be tested against a fake that is always instant.

Faults are opt-in per test. `applyWebProfile()` turns on the set a real web host
shows at once; call it AFTER `installHost`, which resets every fault. It is not
the default, because applied everywhere it would fail hundreds of tests for
reasons that have nothing to do with what they assert.

**It can take a generated `.pptx`.** `insertSlidesFromBase64` really decodes the
bytes — through `scripts/verify-deck.mjs`, the same decoder the audit tool uses,
so the fake cannot read a generated deck differently from the tool that checks
one — and materialises each slide with its slot tag and a `PowerChart` group
holding as many children as the file holds. The child count matters: a readback
measures a chart by what is inside its group, and a fake that put one shape
there made every generated chart read back as wreckage.

**`shapes.items` hands back fresh handles and leaves earlier ones stale**, the
way real Office.js does. The fake used to refresh the shape objects themselves,
so a re-fetch anywhere healed a stale proxy held anywhere — and that one
kindness is why a whole class of stale-proxy bug could only be found by a human
running the add-in in a real PowerPoint. Do not "simplify" it back.

## The structure of `src/` is checked too, and the sweeps are tree-wide

Added 2026-09-29, after a survey found four source sweeps enforcing a general
rule while reading a single file by name.

- **`helpers/module-source.ts`** — the shared sweep. `sweep(pattern, dir)` walks
  a directory, blanks block, HTML and line comments (newlines preserved, so the
  reported line is the one a reader opens the file to), and returns
  `file:line  code` for every offender. `expectSweptSomething` is the guard on
  the guard: a widened sweep that resolves to zero files passes every ban built
  on it, forever. `sourceDeclaring` finds whichever file holds a declaration and
  throws on zero matches or on more than one, so a slice-based guard fails loudly
  when its target moves instead of slicing the wrong text.

  **A ban pinned to one path is defeated by moving the code — it does not go
  red, it goes quiet.** That is the failure this closes, and the first tree-wide
  run of the empty-catch ban found five violations nobody had seen.

- **`import-cycles.test.ts`** — there is no `eslint-plugin-import` here, so
  nothing else looks at the module graph. A cycle is legal TypeScript and the
  suite loads modules in an order that usually survives it; it fails in the
  BUNDLE, as an undefined binding at init, inside PowerPoint. Three combo-base
  edges into `column.ts` are declared in `ACCEPTED_CYCLE_EDGES` with the reason
  `column.ts:38` already gives, and each is checked to still BE an edge, so a
  stale entry fails rather than quietly permitting something else.

  It also holds **the harness/product boundary**, added 2026-09-29 when `app.ts`
  started loading `harness-ui.ts` with `import()`. That took the pane chunk from
  319,021 to 164,912 bytes, and ONE static edge from the product path to any of
  the six harness modules pulls all ~13,400 lines back — silently, because
  nothing in `npm test` measures a bundle. This file is the right home because it
  already reasons about the difference between a static edge and a dynamic one:
  it deliberately ignores the latter, which is exactly why the boundary that
  creates has to be asserted from the other side.

- **`round-pools.test.ts`** — asserts what makes `scripts/round-pools.mjs`
  MUTATABLE: no imports at all, no filesystem, no clock, no `process`. One
  `readFileSync` in there and its tests stop being runnable in Stryker's sandbox,
  which puts the 33 pooled readers back outside `mutate` — and everything would
  still be green. `stryker.config.json` records what that cost the last time:
  four tests written for those readers passed against the code they were written
  to catch. **This file is itself excluded from the mutation suite**, and found
  that out by failing: Stryker's instrumentation reads `process.env`, so a guard
  that greps the source was right about the bytes and wrong about which file
  they were. Same incompatibility as `secondary-axis-ticks.test.ts`.

- **`triage-repo.test.ts`** — the two `triage.test.ts` blocks that genuinely
  cannot run under mutation: a sweep reading every file under `src/`, which
  Stryker rewrites, and a pair of 120-second child-process spawns that can never
  kill a mutant because a child does not inherit the active-mutant global. They
  are here so that `triage.test.ts` — which holds every pooled reader's tests —
  can stay IN the mutation suite. Keep that division: a test moved here because
  it is slow rather than because it cannot run is signal thrown away.

- **`function-size.test.ts`** — a ratchet over every function in `src/` at or
  over 150 **code** lines, measured from the TypeScript AST. Code lines and not
  raw lines because this source is 48% comment: a ratchet that fired when
  somebody documented a function would be switched off within a month. Growth is
  allowed, hiding it is not — re-record with `UPDATE_FUNCTION_SIZES=1`.

## The CI configuration is code too

`workflows.test.ts` reads `.github/workflows/` and pins the few settings this
repo has had to learn the hard way. It exists because nothing else in the suite
looks at those files, so every default in them stays unexamined until it costs
something — and on 2026-08-06 one did: `actions/deploy-pages@v4`'s ten-minute
default timeout cancelled six consecutive Pages deployments while GitHub was
still reporting `deployment_in_progress`, and the live site served an
eight-hour-old build through four merged pull requests. Add a case here when a
workflow default turns out to have been a bet.

## Lockstep-gated files — do not rename

These enforce the feature-set lockstep (see `CONTRIBUTING.md`) and are referenced
by name from `CONTRIBUTING.md`, the PR template, and build scripts:
`skill-docs.test.ts`, `showcase.test.ts`, `manual.test.ts`, `snapshots.test.ts`.
