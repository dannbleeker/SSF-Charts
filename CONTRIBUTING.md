# Contributing to SSF Charts

## Setup & daily loop

```bash
npm install
npm run dev          # demo gallery at http://localhost:3000, pane at /src/taskpane/taskpane.html
npm test             # 300+ tests incl. SVG snapshots, fuzzing, and mocked Office.js hosts
npm run coverage     # same suite + report; CI enforces the thresholds in vitest.config.ts
npm run typecheck
```

The layout engine (`src/core`) is pure TypeScript — most changes are testable
without PowerPoint. The Office.js renderer (`src/render/powerpoint.ts`) is a
thin mapping layer; changes there need a sideloaded host to verify.

## The lockstep rule

Three artifacts must always match the feature set, and CI enforces all of them:

1. **The Claude Agent Skill** (`skill/`) — `test/skill-docs.test.ts` fails if a
   chart kind, datasheet row, or decoration key is missing from the docs. The
   zip is rebuilt by CI from these sources on every push.
2. **The showcase deck** (`examples/showcase.pptx`) — extend
   `scripts/build-showcase.mjs` for new features, run `npm run showcase`, and
   commit the regenerated json+pptx. `test/showcase.test.ts` checks coverage
   and CI diffs the slide XML against a fresh build.
3. **The user manual** (`docs/MANUAL.md`) — `test/manual.test.ts` fails if a
   chart kind, datasheet row, pane control, or element is missing from it.

## Adding a language

The pane ships English-only, but the localization is live and type-gated — a new
language is translation-only, no code changes:

1. In `src/taskpane/i18n.ts`, add one entry to `DICTS`, e.g.
   `de: { ...every key of EN, translated }`. The type is `Record<StringKey, string>`,
   so **TypeScript refuses to compile until every catalogue key is translated** —
   a missing string is a build error, not a silent English fallback.
2. Run `npm test` (`test/i18n.test.ts` checks the catalogue and interpolation).

`localizePane` auto-selects the dictionary from the host display language (or
`?lang=`); `t()` covers runtime status strings, including `{placeholder}`
templates; and the one engine-injected word ("Other") is passed into the pure
core via `ChartConfig.labels.other`. Chart _numbers_ localize via
`NumberFormat.locale`; everything else a chart prints does not — month names,
"Today", "Less"/"More", "no data", "CAGR n/a", "Gap" and the alt text are
English whatever the language, and so are the pane's JS-built control labels
and its `title`/`aria-label` attributes. The header of `src/taskpane/i18n.ts`
lists what a language reaches. Right-to-left scripts need engine layout work and
are out of scope for a translation-only add.

## Snapshots

`test/snapshots.test.ts` freezes every sample chart's SVG. If your change
intentionally alters layout, review the gallery visually first, then
`npx vitest run -u` and mention the change in the PR.

## Releasing

Merges to `main` refresh the rolling `skill-latest` release. For a versioned
release, push a tag: `git tag v0.x.0 && git push origin v0.x.0` — the Release
workflow attaches the skill zip, both manifests, and the showcase deck.
