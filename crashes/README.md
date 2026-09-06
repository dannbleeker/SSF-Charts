Crash reports written by `scripts/round.mjs` when a round wedges.

NOT rounds. Everything that pools `rounds/` — verdict histories, the rasterise
arms, the flip detector — must never see these, which is why they live here.

Each file is PowerPoint's own account of why it died: the console errors, where
the document's data channel stopped, and the ULS window around the fatal entry.
The browser discards all of it when the tab reloads, and recovery reloads the
tab, so this is the only copy.

## Some of these ARE complete rounds

"NOT rounds" is about the directory, not about every record in it. The host dies
in `collectDeckEvidence`, which runs after every verdict is in — 9 builds against
4 on the per-phase traces — so a `-crashed-run.json` written there holds a full
scenario result that was never filed. **33 of 49 do.**

`scripts/salvage-crashed.mjs` turns those into rounds, and refuses the rest by
name. It writes to `rounds-salvaged/`, never here and never to `rounds/`; see
that directory's README for why they cannot join the numbered archive.

The rule above still holds for the pooling functions: nothing should read this
directory as evidence. Read the salvages instead.

## Why the `.md` files are gitignored and the `.json` files are not

This repository is public, and the two halves of this directory are not equally
safe to publish.

`crashes/*.md` is written from the PAGE — console errors, network activity, the
ULS window — so it carries request URLs, session ids and the document's name.
It is excluded in `.gitignore`, and `crashes/README.md` is negated back in.

`crashes/*-crashed-run.json` is written from the PANE's own trace buffer, which
never sees any of that. All 98 committed records were scanned on 2026-09-06 for
host URLs, `resid`/`cid` query ids, personal-drive paths, email addresses and
bearer tokens. None carried any. `test/crash-records-carry-no-identifiers.test.ts`
is that scan, kept, because the pane's payload can grow a field at any time and
the first person to notice would otherwise be a stranger reading the repo.

**If that test ever fails, do not delete the file to make it pass.** Find out
what started writing an identifier into the run log; the record is the only copy
of what the host said as it died, and the leak is upstream of it.
