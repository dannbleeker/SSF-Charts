# Dependency alerts, and what was decided about each

Dependabot has been posting a banner on every push for weeks — "9 vulnerabilities
(3 high, 6 moderate)" — and nobody had read it. A standing red count that nobody
triages is worse than no scanner at all: it trains everyone to scroll past the
one that matters.

This file is the same idea as `KNOWN_ISSUES` in `scripts/office-js-watch.mjs`.
An entry says what was checked and what was decided, **including "no exposure"**,
which records that somebody looked. An alert not listed here has not been read.

**Open as of 2026-11-23** (read from `npm audit`; GitHub's alert API answers this
session's token with 403): **0 findings**. The `image-size`, `qs`, `fast-uri` and
`minimatch` entries further down are held by `overrides` in `package.json`, each
with its `_why` key.

Re-read it with:

```bash
npm audit
gh api repos/dannbleeker/SSF-Charts/dependabot/alerts --jq '.[] | select(.state=="open") | "\(.security_advisory.severity)\t\(.dependency.package.name)\t\(.dependency.scope)"'
```

## Cleared 2026-08-10

`npm audit fix` (no `--force`, so nothing semver-major moved) took the count from
8 to 4: **brace-expansion**, **nanoid**, **postcss** and **undici**, all
development-scope transitives. `pptxgenjs` stayed at 4.0.1, which was the thing
to check — see below.

## Cleared 2026-11-23

`npm audit` had grown to 14 findings (1 high, 13 moderate). Taken on one branch:

- **`zod`** 4.4.3 → 4.6.7 (GHSA-mrhx-c7xc-m3w5, high), under
  `@stryker-mutator/core → mutation-server-protocol`, inside its `^4.1.12`.
- **`minimatch` under the linter** (GHSA-g3hv-75wg-m249): `eslint` 10.15.0,
  `@eslint/config-array` and `@typescript-eslint/typescript-estree` now declare
  `^11` themselves, so raising the `eslint` floor to `^10.15.0` and
  `typescript-eslint` to `^8.73.0` cleared it with no override.
- **`qs`** (GHSA-x5fp-wj9c-mxmx, GHSA-4mjr-xmp4-gh2g, GHSA-6pq8-jhm2-2vcw): the
  existing override moved from `^6.15.3` to `^6.16.1`, same major.

## Cleared 2026-11-23, second pass: `minimatch` under Stryker

`@stryker-mutator/core` 9.6.1 declares `minimatch ~10.2.4`, and so does the
newest Stryker (10.1.1), so neither a refresh nor the Stryker major clears
GHSA-g3hv-75wg-m249. The first pass of the day left it open because an override
across a major on a tool the PR gate never runs "could not be verified here".
The weekly Dependabot review ran the tool instead of reasoning about it:

- Stryker uses minimatch in three files (`file-matcher.js`, `project-reader.js`,
  `options-validator.js`), all of them matching this repo's own `mutate` and
  ignore globs.
- `npx stryker run --dryRunOnly` on 10.2.5 and on 11.0.9 reports the same match:
  `Found 44 of 1150 file(s) to be mutated`, `Instrumented 44 source file(s) with
  24325 mutant(s)`.
- `npx stryker run --mutate src/core/format.ts --dryRunOnly`, the shape the
  weekly sweep runs, completes on 11.0.9: `Initial test run succeeded. Ran 3529
  tests`.

So `overrides` now carries `"@stryker-mutator/core": { "minimatch": "^11.0.1" }`,
scoped to Stryker, and `npm audit` reports 0. Development scope, never shipped.
Remove it once Stryker declares `minimatch ^11`.

**Not caused by this, and not fixed by it:** the FULL-scope dry run (no
`--mutate`) fails its initial test run on this machine with the override and
without it, on the same test (`test/fix-k5-pane-state-and-ui.test.ts`, "starts
level with the sheet": `Invalid regular expression: /\r\n|[\n\r\V\f\u0085\p{Zl}\p{Zp}]/u`).
The test passes under `vitest.mutation.config.ts` outside Stryker and the scoped
run above passes, so it is something about instrumenting all 44 files at once.
Measured 2026-11-23 on Node 22.23.5; not yet explained.

## Held by overrides (the original triage, kept for the reasoning)

### `image-size` — 2 × high — NO EXPOSURE

> ICNS parser allows denial of service through an infinite loop
> JXL and HEIF parsers allow denial of service through infinite loops

The only alert scoped **runtime**, which is what makes it worth writing down
rather than shrugging at.

- It arrives one level down: `pptxgenjs@4.0.1 → image-size@1.2.1`.
- **There is no forward fix.** 1.2.1 is the newest 1.x, the patch is in 2.x, and
  pptxgenjs — already on its newest release — pins `^1.2.1`. `npm audit`'s
  proposed remedy is to move `pptxgenjs` to **1.1.5**, three majors backwards,
  which would take the whole pptx renderer with it. That is not a fix.
- **It is not reachable.** `image-size` runs when pptxgenjs embeds a picture.
  This repo calls `addImage` in exactly one place — `skill/scripts/render-pptx.mjs`,
  on the `render: "image"` path — and what it hands over is a **PNG this code
  rasterised itself** from its own scene graph through resvg. The vulnerable
  parsers are ICNS, JXL and HEIF. No user file, of any format, reaches it.
- **It is not shipped to the browser either.** pptxgenjs is loaded by dynamic
  `import()` so it stays out of the pane's first-load bundle, and the built
  `dist/assets/pptxgen.es-*.js` carries Vite's `__vite-browser-external` stub
  where the Node-only imports were — none of the parsers are in the pane at all.

Revisit when pptxgenjs takes `image-size@2`. Do not downgrade pptxgenjs.

### `qs` — 1 × moderate — NO EXPOSURE

> `qs.stringify` crashes with a TypeError — a remotely triggerable DoS

It arrives under `@stryker-mutator/core`, the mutation-testing runner:
`@stryker-mutator/core@9.6.1 → typed-rest-client@2.3.1 → qs@6.15.1`. Stryker is a
devDependency, it runs in one weekly scheduled job on a GitHub runner, and it is
in neither the add-in nor the `test` job that gates a PR. Nothing here parses a
query string or makes an HTTP request at runtime.

**Re-read 2026-08-23, and two things had changed since this entry was written.**

- `typed-rest-client` no longer carries an alert of its own — it is not open, and
  it is not in the fixed list either. The heading used to say "2 × moderate" and
  claimed a package that is no longer named. Corrected.
- **A forward fix now exists** (6.15.2), where before there was none. That makes
  this a different decision from the `image-size` one above, which is still
  blocked: here the patch is reachable, and the answer is still no.

**An `overrides` block was added and reverted the same hour.** `{"qs": "^6.15.2"}`
resolves to 6.15.3 and clears the alert, and the gate stays green — so the
temptation is real and it is worth writing down why it was undone. Taking it
would pin a transitive that Stryker has not tested against, and add the repo's
first `overrides` block — a permanent maintenance surface every future install
reads — to remove an exposure that is **nil**, in a tool that never runs in CI's
gating path. A clean alert list is not worth a standing lie about what this
project depends on.

Revisit when Stryker updates its own tree. Nothing to do meanwhile — and
"nothing to do" here now means *a fix was available and declined*, not *no fix
existed*.
