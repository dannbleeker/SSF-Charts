/**
 * What the fake claims, what a real host said, and the diff between them —
 * with no Node in it.
 *
 * Split out of `host-diff.mjs` so three callers can share one copy: the CLI,
 * the CI gate, and the TASK PANE. The pane is why the split exists — it can now
 * tell the owner whether a probe run found anything before they send the file,
 * and it could not import a module that opens with `import ... from "fs"`.
 *
 * Nothing here may reach for the filesystem, the network, or `process`. The
 * moment it does, the pane stops building and the reason will not be obvious.
 */
/** The fake's frozen answers, as asserted in `test/host-probe.test.ts`. */
export const FAKE_BASELINE = {
  // The fake's collection answers honestly, so both routes work there and the
  // interesting answer — index answering where the list does not — is a
  // divergence by construction. That is the point: this question exists to be
  // answered by the REAL host, and a double that reproduced the host's deafness
  // would be modelling a bug rather than a contract.
  "shapes-by-index-vs-items": "both-answer",
  // THE FAKE DOES NOT MODEL `creationId`, AND DELIBERATELY DOES NOT. A double
  // that handed back a plausible id would make these read `yes`, `stable` and
  // `kept` in every suite run while saying nothing whatever about the real host
  // — which is the only thing they exist to ask. `absent` is the truthful answer
  // for an object without the property, and a real host answering anything else
  // is exactly the divergence worth seeing.
  //
  // See `reference_fake_populates_on_request`: a fake that fills in what the
  // host refuses cannot fail, and it hid a 56-round defect once already.
  "creationid-on-fresh-shape": "absent",
  "creationid-survives-a-sync": "no-creation-id",
  "creationid-survives-grouping": "no-creation-id",
  "load-isnullobject-populates": "unreadable",
  "load-id-populates-isnullobject": "yes",
  "getitemornullobject-missing": "null-object",
  "shape-add-fresh-slide-proxy": "yes",
  "shape-add-held-slide-proxy": "threw",
  // Its partner does the identical thing a moment later on another fresh slide.
  // The fake refuses held proxies consistently, so the pair AGREES here — which
  // is the "the host has a definite behaviour" outcome the question exists to
  // tell apart from a coin. A real host that answered these two differently in
  // one run would be saying the opposite, and that is the finding.
  "shape-add-held-slide-proxy-again": "threw",
  "shape-resolve-held-slide-proxy": "yes",
  "shape-add-fresh-getitem-slide": "yes",
  "shape-add-positional-slide-proxy": "yes",
  "shape-proxy-survives-one-sync": "yes",
  "shapes-items-count-honest": "at-least-5",
  "shapes-items-via-positional-slide": "at-least-5",
  "getcount-populates-same-sync": "yes",
  "tags-add-same-key-twice": "overwrites",
  "tags-on-fresh-shape": "yes",
  "tag-through-refetched-shape": "yes",
  "how-many-syncs-a-creation-handle-survives": "survives-8",
  // `yes` here is the fake saying its collection read did NOT poison the
  // creation handle — on the scratch slide, where the shape has no other
  // handles onto it. Production's does, which is the whole question; see
  // PENDING_QUESTIONS.
  "collection-read-poisons-the-creation-handle": "yes",
  // The healthy fake answers every read, so nothing degrades and it says so.
  "how-many-collection-reads-a-context-survives": "survives-12",
  // The fake's collection read drops from the TAIL (`faults.readsMissing`), but
  // unarmed it drops nothing, so a healthy fake keeps every shape.
  //
  // AND THAT WAS NEVER AGREEMENT, IT WAS TWO SHRUGS MATCHING. This read `all`
  // until 2026-09-01 and so did the host, in 87 rounds, and `host-diff` compares
  // only the answer — so the sheet recorded 87 rounds of host-agrees-with-fake
  // about which end a short read drops, on a question neither had ever been in a
  // position to answer. The renamed word says what it always meant: the
  // precondition did not arise. It is weak now, so a real answer displaces it.
  "which-end-a-short-read-drops": "not-a-short-read",
  // The fake DOES refuse a group through a slide handle two syncs old — that is
  // `parentWindowOk`, modelled from a real host naming the parent — and the tag
  // written afterwards still lands. So the fake's answer is "the context
  // survives a refused group", and the real host is what this question is for.
  "does-a-failed-group-poison-the-tag": "yes",
  "delete-then-lookup": "reports-gone",
  "addgroup-returns-usable": "yes",
  "group-children-via-getcount": "two",
  "group-reports-its-children": "two",
  "binding-names-shape-later": "yes",
  "getitemat-past-end": "threw",
  "picture-then-shape-read": "yes",
  "group-of-existing-shape-readable": "2",
  "slide-layout-readable": "yes",
  "layouts-readable": "yes",
  "untrack-available": "no",
  // The fake models `untrack` on a SHAPE and not on a null-object slide, which
  // is why this pair is worth asking as a pair: the fake already behaves the
  // way the confound predicts, and only the real host can say whether it does
  // too. Written from what the fake actually answers, after a guess at `no` was
  // corrected by the gate.
  "untrack-available-on-shape": "yes",
  "scratch-slides-returned": "all",
  // THE FAKE ANSWERS THE WAY THE PRODUCT ASSUMES, which is the point of asking.
  // Every diagonal we draw is a rectangle of the segment's LENGTH rotated about
  // its midpoint, so `unrotated-box` is not an observation here — it is a
  // restatement of what `addSegment` and `arrowheadBox` are built on. A real
  // host answering `POST-ROTATION-BOX` is therefore a divergence saying every
  // diagonal in every line, scatter, radar and violin chart is drawn at the
  // wrong size and in the wrong place, which is the finding no round could
  // reach while the question was asked through a chart.
  "rotation-keeps-the-unrotated-box": "unrotated-box",
  // `draws` only since the fake's GeometricShapeType learned `hexagon`. It
  // threw before that, and the throw was right: the hex tile map had just moved
  // onto a preset name that nothing rendered through this renderer.
  "named-preset-resolves": "draws",
};

/**
 * Divergences between the fake and the committed real-host sheet that are
 * EXPECTED, each with the reason it is allowed to stand.
 *
 * The point of the table is what it makes impossible. `test/host-contract.test.ts`
 * diffs the fake's baseline against `test/fixtures/host-answers-web.json` and
 * fails on any divergence that is not declared here — so a change to the fake
 * that contradicts a real PowerPoint stops being something a human discovers by
 * running PowerPoint, which is how every host bug in this project has been
 * found so far, and starts being a red test in two seconds.
 *
 * A divergence belongs here for exactly one of two reasons, and the note has to
 * say which:
 *
 * - the fake models a DIFFERENT host on purpose (its happy path is a host that
 *   behaves, which is what makes most tests readable), or
 * - the real answer is known to be about the probe rather than the host, and
 *   the question has been re-asked but not yet re-run.
 *
 * "We have not looked into it" is not a reason. An entry with a note like that
 * is a to-do wearing a passing test's clothes.
 */
export const KNOWN_DIVERGENCES = {
  "untrack-available-on-shape":
    "The fake models the host `untrack()` was designed for — one where a shape proxy carries the method — and PowerPoint on the " +
    "web is not that host. ANSWERED 2026-08-12 (`89675b6`, reproduced on `1789749`): `no`, asked of a proxy " +
    "`addGeometricShape` had just returned, which is the kind `renderShapesChunked` holds hundreds of. That removes the " +
    "confound its trigger carried — `untrack-available` asks a NULL-OBJECT slide proxy, the one kind most likely to lack any " +
    "method — so the `no` is about the platform rather than about the probe. " +
    "Microsoft's performance guidance names untracking as the remedy for our exact symptom (\"large batch operations may " +
    'generate a lot of proxy objects... a noticeable performance benefit when using large numbers of proxy objects"), so this ' +
    "closes that idea on evidence rather than leaving it as an omission in the draw path. Do not re-propose it. The fake keeps " +
    "the method because `untrack` is best-effort everywhere this repo calls it and a host that has it is a real host.",
  "load-isnullobject-populates":
    "The fake models the host `queueNullCheck` was written for, where loading the flag by name populates nothing. PowerPoint on the web does populate it. Both hosts are real; the workaround is harmless on this one rather than necessary.",
  // ---------------------------------------------------------------------------
  // THE THREE BELOW ARRIVED WITH THE 2026-09-01 FIXTURE SWAP, and the number is
  // the news: the previous capture would have brought SIX. Three went away by
  // themselves when `unreadable`, `silent` and a probe's own `all` stopped
  // locking rows against the real answers underneath them — the answer-ranking
  // work of 2026-08-31 and 2026-09-01. These three are what is actually left.
  // ---------------------------------------------------------------------------
  "rotation-keeps-the-unrotated-box":
    "The fake models the host this renderer is BUILT for — one that keeps reporting the unrotated box after `rotation` is " +
    "set, which is what `addSegment` and `arrowheadBox` both assume. This host will not say. Measured across 36 probe " +
    "passes in 12 rounds: `unreadable` 31 times, `unrotated-box` 4, and the contradicting answer NOT ONCE. The detail " +
    "names the cause — 'neither shape's width came back' — and it is the documented behaviour of this host, which answers " +
    "no geometry for a shape a probe has just added. Note the deck INVENTORY reads width and height perfectly for shapes " +
    "already on a slide (2026-09-01, 49 of 49 across three rounds), so this is about a just-added proxy and not about " +
    "geometry in general. The four readings all agree with the fake, so nothing here contradicts the assumption the " +
    "product rests on; what is missing is confirmation, not evidence against. If a future capture reads `rotated-box`, " +
    "every diagonal in every line, scatter, radar and violin chart is drawn wrong and this entry is the first thing to " +
    "revisit.",
  "shape-proxy-survives-one-sync":
    "The fake's happy path hands back a shape proxy that is still usable one sync later. This host throws. It is the same " +
    "5010 refusal every other id route meets, and it is why `renderShapesChunked` re-reads rather than holding proxies " +
    "across a sync. Declared rather than modelled: teaching the fake to poison a proxy would change what every other " +
    "test in this repo is standing on, for a path the product already avoids.",
  "shapes-by-index-vs-items":
    "The fake's happy path answers both routes into a shape collection. This host answers `index-beats-items` or " +
    "`index-unreadable` depending on the capture — it is already declared in UNSTABLE_ANSWERS as a coin, and the two " +
    "faces are what the 2026-09-01 pair caught (`index-unreadable` in round 350, `index-beats-items` in 351). Either " +
    "way the by-index route is the one the product uses.",
  // ---------------------------------------------------------------------------
  // THE NINE BELOW ARRIVED TOGETHER, ON 2026-08-29, WITH THE FIXTURE SWAP.
  //
  // Not nine new host behaviours. Nine that this gate has never been able to
  // see: the fixture it compared against was captured 2026-08-12 and answered
  // 31 questions, while every round since round ~254 answers 39. All nine sat
  // in `PENDING_QUESTIONS` as "the sheet predates this question" while the host
  // answered them, consistently, for twenty-five rounds and more.
  //
  // Each note carries the tally over the 25 rounds to 2026-08-28, because one
  // sheet is one sample and several of these are coins. Where the count is not
  // unanimous the question is ALSO in `UNSTABLE_ANSWERS`, and the rule there
  // holds: do not build on either face.
  // ---------------------------------------------------------------------------
  "tags-on-fresh-shape":
    "The fake's happy path writes a tag on a shape it has just made. THIS HOST WILL NOT: `threw`, 25 of 25 rounds to " +
    "2026-08-28, with no second face anywhere in that window. It is the sharpest of the nine — unanimous, and about the " +
    "mechanism the whole product identifies charts by. It does not contradict `tag-the-creation-proxy-a-sync-later` " +
    "(`yes`, 25 of 25): a tag lands on that proxy A SYNC LATER and not in the batch that made the shape, which is " +
    "precisely the order `settleAndTagChart` already uses. So the product is on the right side of this and the fake is " +
    "not. Kept as a divergence rather than modelled, because a fake that refused fresh-shape tags would make several " +
    "hundred ordinary tests read as tag-timing exercises.",
  "scratch-slides-returned":
    "The fake's happy path gives every scratch slide back. This host gives SOME back: 25 of 25 rounds say `some`, with " +
    "no other face. That is the id-namespace finding seen from the clean-up end — a slide taken by id and returned by " +
    "position is not always the same slide to this host — and 77f9ca4 is the commit that stopped the probe paying for " +
    "the difference. The probe already sweeps by position for exactly this, so `some` is survivable; it is declared " +
    "because the fake being the tidy one is the direction that misleads.",
  "collection-read-poisons-the-creation-handle":
    "The fake's happy path answers `yes` — its collection read leaves the creation handle usable — and says so from a " +
    "SCRATCH slide, where the shape has no other handle onto it. Production's does, which was always the real question. " +
    "This host answers `refused`, 25 of 25. So the answer is in and it is the pessimistic one: reading a slide's shape " +
    "collection costs you the handles you were holding. Every batching decision in `renderShapesChunked` is on the right " +
    "side of that already.",
  "does-a-failed-group-poison-the-tag":
    "The fake models the host this code was written against, where a refused group leaves the context usable and the tag written afterwards lands. This " +
    "host answers `tags-gone`, 25 of 25 — a failed `addGroup` takes the tag with it. That is what `settleAndTagChart` " +
    "exists for and why it opens a FRESH context rather than continuing in the one that just failed. Declared rather " +
    "than modelled: the fake's version is what makes an ordinary grouping test about grouping.",
  "picture-then-shape-read":
    "office-js#5022. The fake's happy path reads a shape collection after a picture insert; this host mostly will not. " +
    "`unreadable` 21 of 25, `yes` 4 — A COIN, and its `UNSTABLE_ANSWERS` entry is the one to read before acting. One " +
    "pass in five is not permission: `drawDemoItem` performs exactly this sequence, so a `yes` would read as licence to " +
    "stop guarding it.",
  "shape-add-held-slide-proxy":
    "The fake's happy path refuses a held slide proxy consistently, which is what makes its pair with `-again` a readable test. This " +
    "host is a COIN and has flipped the pair's direction: `yes` 21 of 25, `threw` 4. Its partner answers `threw` 24 of " +
    "25 in the same rounds, so the two now disagree ROUTINELY rather than exceptionally — which is itself the finding, " +
    "and the reason both stay in `UNSTABLE_ANSWERS`. Do not build on either face.",
  "how-many-collection-reads-a-context-survives":
    "The fake's happy path answers every read and reports `survives-12`. This host answers `short-at-1` 23 of 25 — the " +
    "FIRST collection read in a context already comes back short — with `survives-12` twice. The pessimistic reading is " +
    "the one the product is built on, and the two `survives-12` rounds are what `UNSTABLE_ANSWERS` is for.",
  "how-many-syncs-a-creation-handle-survives":
    "The fake's happy path keeps a creation handle alive for eight syncs. This host answers `refused-after-1` 23 of 25, " +
    "`survives-8` twice. One sync is the budget, and `renderShapesChunked` spends it that way. The two optimistic " +
    "rounds are recorded rather than averaged away.",
  "tag-through-refetched-shape":
    "The fake's happy path re-fetches a shape by id and tags it. This host answers `no-id` 23 of 25 — the re-fetch has " +
    "no id to tag through — with `yes` twice. It is the same 5010 refusal every other id route meets, seen from the tag " +
    "side, and it is why the config tag is written on the GROUP rather than on a re-fetched member.",
  "shapes-items-count-honest":
    "The fake's happy path answers a shape collection honestly, and `faults.hollowReads` is where the refusal lives. RE-ASKED AND ANSWERED 2026-08-08: `short-0`, items=0, with `getcount-populates-same-sync` answering `yes, value=8` in the same run. Same host, same minute: the count is right and the list is empty. That is the sharpest form this bug has taken, and it is what `slideShapeNames`' corroboration check exists to catch. The earlier WITHDRAWN note (the `items`-undefined answer being about the handle, not the collection) is settled by the partner below.",
  "tags-add-same-key-twice":
    "WITHDRAWN, awaiting a re-run. The real `other — value=undefined` was one shape proxy held across four syncs, not an opinion about tag keys.",
  "addgroup-returns-usable":
    "The fake's happy path hands back a group usable in the sync that made it. RE-ASKED 2026-08-08 and still `unreadable`, but read the detail before calling it settled: `members via same-batch`. The strict id route could not supply members — the collection answers empty, see above — so this asked with same-batch proxies, which is the weaker form of the question. What it establishes is that the weak form fails; the strong form has still never been put on this host.",
  "group-reports-its-children":
    "The fake's happy path lists a group's children. RE-ASKED 2026-08-08 with the load queued in the sync that MADE the group — the friendliest form there is — and this host still answered `threw`, \"The property 'items' is not available\": office-js#6363's signature. Its sibling `group-children-via-getcount` was added to ask the other way and answered `unreadable`. Both routes refused, so `contentShapes` returning UNKNOWN_CONTENT for a grouped slide is the permanent answer rather than a gap.",
  "group-children-via-getcount":
    "The fake's happy path counts a group's children; the web host's refusal lives in a named fault rather than the default. ANSWERED 2026-08-08: `unreadable`. Read with `group-reports-its-children` above — both ways into a group's children are refused on this host. Also carried in UNSTABLE_ANSWERS, because it has been asked once and once is a sample.",
  "group-of-existing-shape-readable":
    "The fake's happy path names a group it has just made, so the later-batch question can be put at all. This host would not: `no-group-id`. That is an answer and not a setup failure — a host that will not name a fresh group cannot be asked about resolving one from the deck afterwards, and the fact belongs in the sheet. It also means `countGroupChildrenPage`, which swallows failures per shape, produces no error and no measurement here.",
};

/**
 * Questions this host has answered DIFFERENTLY on different runs of one build.
 *
 * Deliberately not a `KNOWN_DIVERGENCES` entry, and the contract gate rejecting
 * the first attempt at that is the reason this list exists. A divergence is
 * "the fake and the real host disagree", which is a fact about two systems and
 * is either true or it is not. This is a fact about ONE system: the real host
 * gave answer A, then gave answer B, minutes apart, same build. The gate has
 * nothing to say about that and should not be made to.
 *
 * What it is for is stopping the next reader — human or agent — from building
 * on whichever answer a sheet happens to carry. `shape-add-positional-slide-
 * proxy: yes` is exactly the answer that makes a positional slide handle look
 * like the safe route out of the by-id refusals, and it is one of the two that
 * flipped.
 *
 * The mechanism is in the run log rather than inferred: three `scratch slide
 * landed but its id will not resolve` lines mid-run, two replacement scratch
 * slides taken, and every question asked inside that window answering
 * `no-scratch-slide` before the host came back. The host's ability to resolve a
 * freshly added slide's id comes and goes within a single 37-second run — the
 * same reversible bimodality the draw times show.
 *
 * A question in here has been SAMPLED, not answered. Removing an entry needs
 * several runs agreeing, not one.
 */
/**
 * HOW OFTEN EACH SCENARIO KILLS THE HOST — a rate, and a ratchet.
 *
 * A round file has three verdict states and cannot hold a fourth: a scenario
 * that KILLS the host produces no verdict, and the round it belonged to files
 * nothing at all. So the worst thing a scenario can do is the one thing
 * `rounds/` cannot say, and it has hidden this:
 *
 *     same scale across the deck — 0 failures in 282 recorded verdicts,
 *     joint-safest in the whole suite, and the scenario the host died inside
 *     NINE times.
 *
 * `rounds-salvaged/` does not help: a salvaged round carries verdicts, and a
 * killed scenario has none. Only the crash record's own steps know.
 *
 * DEATHS PER 1000 RUNS, AND THE FIRST VERSION OF THIS TABLE WAS A RAW COUNT.
 *
 * That version shipped on the morning of 2026-09-03 and was red by the
 * afternoon. It could not have been anything else. `crashes/` is append-only,
 * so counts only ever grow; a budget seeded at today's total breaches on the
 * very next death, and for a scenario that has been dying all week the next
 * death is Tuesday. It fired on the episode continuing rather than on a
 * regression, and a gate that cries wolf gets switched off — which
 * `docs/BACKLOG.md` records happening to an earlier one.
 *
 * It also RANKED WRONG. Deaths arrive in episodes, not at a steady drip:
 * `same scale` took all nine of its deaths between 2026-08-24 and 08-29 and
 * none since, and `stop a run mid-draw` took all eight on 09-02 and 09-03. So a
 * cumulative count mostly measures how long a scenario has been exposed. Put
 * against how often each actually runs, the order inverts:
 *
 *     a big chart on a slide of its own    5 deaths /  11 runs = 455
 *     stop a run mid-draw                  8 deaths /  27 runs = 296
 *     same scale across the deck           9 deaths / 421 runs =  21
 *     every other listed scenario          1 death  / ~420     =   2.4
 *
 * THOSE ARE THE 2026-09-03 READINGS AND THE CEILINGS ARE STILL SEEDED FROM
 * THEM. Re-derived 2026-09-08, and every one has fallen — not because anything
 * was fixed, but because the denominators grew while the deaths did not:
 *
 *     a big chart on a slide of its own   12 deaths /  84 runs = 143  (ceiling 460)
 *     stop a run mid-draw                  8 deaths /  93 runs =  86  (ceiling 330)
 *     same scale across the deck           9 deaths / 495 runs =  18  (ceiling  30)
 *
 * So the two worst are now running at roughly a third of what they are allowed,
 * and the gate would not notice either of them tripling. That is the cost of a
 * ceiling seeded once from a point estimate, and it is the mirror of the zero
 * problem below rather than a separate defect. Whether to re-seed is the
 * owner's: lowering a ceiling is deciding a scenario may kill PowerPoint LESS
 * often than it used to be allowed to, which is a claim about the product, not
 * about the instrument.
 *
 * The scenario that led on raw count is a nine-times-baseline one. The two that
 * kill the host on a THIRD to a HALF of their runs looked smaller only because
 * they are young — `a big chart on a slide of its own` was written on
 * 2026-09-03 and has run eleven times. Both are the OWN-SLIDE pair, which is
 * the defect `6438dd1` fixed at 16:9 and round 370 showed still open at 4:3.
 *
 * AND A RATE CAN FALL. This is the property the instrument needed and the one a
 * count can never have: runs keep accumulating, so a scenario that stops dying
 * sinks on its own and a landed fix protects itself without anyone editing a
 * number. That is what `overlap-budget` does — it measures a CURRENT quantity —
 * and it is why the analogy to it finally holds. An earlier comment here
 * promised a fall over a cumulative count and had to be retracted; this is what
 * it should have said.
 *
 * SEEDED AT THE RATES MEASURED ON 2026-09-03, rounded up to the next ten. At
 * zero this would fail every night for damage already known and already on the
 * backlog. At today's rates it is silent while a known-bad scenario stays as
 * bad as it is, and loud the moment one gets WORSE — which is the question
 * worth asking, and the one the count could not ask.
 *
 * A NAME ABSENT FROM THIS TABLE HAS NEVER KILLED THE HOST, so its first death
 * is a rise from zero and is caught. That is the case most worth catching, and
 * it is the one thing the count version got right. Do not add a name here to
 * quiet a gate: a new name is a new scenario killing PowerPoint.
 *
 * AND SINCE 2026-09-08 THAT SENTENCE HAS A SECOND HALF. Absent now means never
 * killed the host, OR killed it once with a receipt beside it in
 * `DEATHS_ACKNOWLEDGED` (`scripts/rounds-gate.mjs`). The receipt names one
 * crash record, clears the gate's exit and not the count, cannot touch a
 * scenario that has a ceiling, and cannot cover a second death. It exists
 * because the rule above had no green path at all: at p=0 the allowance is 0
 * for every denominator, so a first death breached for ever and the nightly
 * cycle stopped after one round from round 428 on. Adding the name here would
 * have been the thing this paragraph forbids; a receipt is the thing it does
 * not.
 *
 * Attribution is narrow on purpose: a scenario counts only when its
 * `scenario starting` line was never closed. **36 of 99** sound crash records
 * attribute this way and the other 63 died in the probe phase or the deck
 * scan, credited to nothing (re-derived 2026-09-08; it read "26 of 83"). See
 * `fatalScenarios` in `scripts/triage.mjs`, and `scenarioRuns` for the
 * denominator.
 */
export const FATAL_SCENARIO_RATE = {
  /**
   * IT KILLS POWERPOINT ON NEARLY HALF ITS RUNS. Five deaths in eleven, and it
   * has only existed since 2026-09-03 — written that day to isolate the
   * own-slide defect from `stop a run mid-draw`, which it did.
   *
   * This is the worst number in the suite by a factor of 190 over baseline, and
   * it is not a harness problem: the scenario draws a chart onto a slide the
   * product just added, which is a thing users do.
   */
  "a big chart on a slide of its own": 460,
  /**
   * Ours, and known: `ca138f8` moved it onto a freshly-added slide and it has
   * failed every round since. Same defect as the entry above — the two share
   * it, which is why they share an order of magnitude.
   *
   * Nine seconds long and 296 per 1000. Duration is not danger.
   */
  "stop a run mid-draw": 330,
  /**
   * THE ONE THIS TABLE EXISTS FOR, and the reason it is a rate. Nine deaths and
   * not one failed verdict in 282 — every reader of this archive has been told
   * for weeks that this scenario is among the safest in the suite.
   *
   * Nine times baseline, not the worst in the suite. All nine landed between
   * 2026-08-24 and 08-29, inside `updated only the shapes that changed`, in the
   * 420-600s window where this host dies of session age. Nothing since.
   */
  "same scale across the deck": 30,
  /**
   * SEEDED BY THE OWNER ON 2026-09-09, at his instruction, after two deaths.
   *
   * This is the first name added to this table since it was written, and the
   * docstring above forbids adding one "to quiet a gate". It is here because the
   * scenario reached TWO deaths, which the receipt mechanism refuses to sign —
   * one death is a coin, two is a pattern — and a ceiling is then the only green
   * path there is. That refusal working as designed is what made this a decision
   * rather than a workaround.
   *
   * 2 deaths in 45 runs = 44.4 per 1000. The floor that holds is 12; 50 is the
   * current rate rounded up to the next ten, which is how every other entry here
   * was seeded and which matches this table's stated purpose: silent while a
   * known-bad scenario stays as bad as it is, loud the moment it gets worse. A
   * ceiling of 12 would have re-fired on the third death.
   *
   * WHAT IT ADMITS: this scenario may kill PowerPoint on about one run in
   * twenty. That is worse than `same scale across the deck` and better than the
   * two own-slide names above it.
   *
   * WHY IT DIES, which is known and is not a mystery to be solved by this
   * number: `kindCostSpread` deletes each specimen before drawing the next, and
   * on this host that delete does not land — all four routes to sweeping a fresh
   * shape are measured and closed in `docs/BACKLOG.md`. So occupancy climbs
   * 0, 7, 15, 24, 34, 44, 53, 61 and the tab dies on the eighth. The fix is to
   * draw each specimen on a scratch slide that can be deleted whole, which
   * changes the occupancy the experiment measures at — the reason it is the
   * owner's call and not a cleanup. When it lands, this number should come back
   * down, and lowering it is his call too.
   */
  "what a chart kind costs": 50,
  // The tail: one death each, on 200-420 runs. At 2.4 per 1000 these are the
  // background rate of a host that falls over sometimes, not scenarios with a
  // problem. They are listed so their FIRST rise is measured against something.
  "one chart alone on a warm deck": 10,
  "edit the chart the user selected": 10,
  "explode a degraded picture": 10,
  "insert onto a slide that already has content": 10,
  "two slides claiming one slot": 10,
  // TWO NAMES WERE HERE ON 2026-09-03 AND SHOULD NEVER HAVE BEEN.
  //
  // `a chart of rotated shapes` (2) and `where a rotated shape lands` (1) were
  // credited three deaths by a matcher that knew one spelling of three: the
  // host writes `scenario FAILED` and `scenario skipped`, and the close regex
  // read `passed|failed`. The scenario that closed was left open and the
  // record's death landed on it. Corrected in `fatalScenarios`; both have
  // killed the host exactly zero times.
  //
  // Their ABSENCE is load-bearing, per the rule above.
};

export const UNSTABLE_ANSWERS = {
  // ---------------------------------------------------------------------------
  // ADDED 2026-09-01, WHEN IT LEFT `PENDING_QUESTIONS` FOR THE COMMITTED SHEET.
  // ---------------------------------------------------------------------------
  "rotation-keeps-the-unrotated-box":
    "A COIN THAT LANDS ON ITS EDGE almost every time: `unreadable` 31 of 36 probe passes across 12 rounds, " +
    "`unrotated-box` 4, and the contradicting answer never. It is not that the host disagrees with itself about where a " +
    "rotated shape sits — it is that it will not answer at all for a shape a probe has just added, and occasionally it " +
    "does. That is why the question keeps its `resample` mark: a 1-in-9 read rate is exactly the case a single pass " +
    "cannot settle, and the four reads it has managed all agree with the fake. Do not read a round of `unreadable` as " +
    "the answer, and do not read one `unrotated-box` as proof either — the useful number is that nothing has ever come " +
    "back saying the box is the POST-rotation one, which is the reading that would mean every diagonal in the product " +
    "is drawn wrong.",
  // ---------------------------------------------------------------------------
  // THE SEVEN BELOW WERE ADDED 2026-08-29 FROM THE ARCHIVE, NOT FROM A SHEET.
  //
  // `scripts/host-history.mjs` exists because entries here go stale silently
  // when they are written from whatever rounds the author had open. These seven
  // were written the way that tool tells you to: one pass over the 25 rounds to
  // 2026-08-28, in round-time order, counting faces. Re-run it before editing
  // any of them —
  //
  //     node scripts/host-history.mjs --fixture rounds/2*.json
  //
  // Five of the seven had never been in this table because they were sitting in
  // `PENDING_QUESTIONS` as "the sheet predates this question", which they had
  // stopped being some hundred rounds earlier. See that table's note.
  // ---------------------------------------------------------------------------
  "how-many-collection-reads-a-context-survives":
    "A COIN, and the cheap face is the common one: `short-at-1` 23 of the 25 rounds to 2026-08-28, `survives-12` twice. " +
    "The first collection read in a context usually comes back short, which is what every readback in `powerpoint.ts` is " +
    "written for. The danger is the other face: `survives-12` says a context can be trusted for a dozen reads, and two " +
    "rounds in twenty-five is not a licence to batch on that assumption.",
  "how-many-syncs-a-creation-handle-survives":
    "A COIN: `refused-after-1` 23 of 25, `survives-8` twice. One sync is the budget `renderShapesChunked` spends, and it " +
    "is the right one. Same warning as its neighbour above — the optimistic face is the one that would licence holding " +
    "handles across a whole chart, and it has appeared twice.",
  "tag-through-refetched-shape":
    "A COIN, heavily weighted: `no-id` 23 of 25, `yes` twice. The re-fetch usually has no id to tag through, which is the " +
    "5010 refusal every other id route meets. The two `yes` rounds are why the config tag is still written on the GROUP " +
    "rather than on a re-fetched member: a route that works one round in twelve is not a route.",
  "shapes-by-index-vs-items":
    "THREE FACES over 25 rounds — `index-unreadable` 15, `both-answer` 9, `no-count` 1. The interesting one is " +
    "`both-answer`, because it says the collection answered BOTH ways in that round, and it is common enough (better " +
    "than one round in three) that a sheet catching it looks like good news about the host. It is not: the modal answer " +
    "is that the index route answers where the list does not, which is what `shapes-by-index` in `powerpoint.ts` is for.",
  "shape-proxy-survives-one-sync":
    "office-js#2903, and A COIN rather than the flat refusal this project has described: `threw` 18 of 25, `yes` 7. " +
    "Moved here from `KNOWN_DIVERGENCES` on 2026-08-29, where it had read as a settled refusal — the fixture happened to " +
    "carry `threw`, and nothing compared it against the archive. Seven rounds in twenty-five say a proxy DOES survive a " +
    "sync here. Neither face may be built on, and the pessimistic one is what the draw path already assumes.",
  // `group-of-existing-shape-readable` was considered here and deliberately left
  // out: `threw` 20 of 25 and `no-group-id` 5, but BOTH faces are refusals and
  // the conclusion is identical either way, so a scarce scratch slide spent
  // re-asking it buys nothing. This table's membership costs probe slides — it
  // is what `RESAMPLE_IDS` is derived from — so a question only belongs here
  // when its faces would lead to DIFFERENT decisions. Its divergence entry
  // carries the two-face detail.
  "binding-names-shape-later":
    "REOPENED 2026-08-29, AND THIS IS THE ONE TO READ TWICE. The entry that stood in `KNOWN_DIVERGENCES` said the host " +
    "answers `commit-threw` — the batch carrying a binding is REJECTED — and that PowerPointApi 1.8 bindings were " +
    "therefore retired as an idea. Over the 25 rounds to 2026-08-28 the tally is `yes` 15, `silent` 8, `unreadable` 1, " +
    "`commit-threw` 1. So the rejection is the RARE face now, and the common one is that the binding works.\n" +
    "That does not make bindings a route — `silent` in a third of rounds is its own problem, and one commit-threw is " +
    "still a commit-threw — but it does retire the flat 'cannot be asked on this host' this project has been carrying, " +
    "including in docs/BACKLOG.md §2. It is the last untried way out of the id refusals and the evidence against it has " +
    "weakened. Nothing in the repo makes a binding, so there is no caller at risk either way; what changed is whether " +
    "the idea deserves another look, and it does.",
  "picture-then-shape-read":
    "office-js#5022's question — can a shape collection be read after a picture insert — and a COIN, seen flipping inside one " +
    "round on 2026-08-12 (`1789749`): `yes` on pass 1 while the host was in slide-trouble, then `unreadable` on passes 2 and 3 " +
    "once the collection was refusing. The committed fixture carries `unreadable`. " +
    'It is in this table because of what the other value would licence: a `yes` reads as "#5022 does not affect this host", ' +
    "and `drawDemoItem` performs exactly that sequence — insert a picture, then read shapes — so somebody would reasonably " +
    "stop guarding it. One pass in three is not permission. Note the regimes: both readings came from the same round minutes " +
    "apart, so this is the host moving rather than two builds disagreeing.",
  "shape-add-held-slide-proxy-again":
    "RETIRED FROM STABLE ON 2026-08-12 (`89675b6`), by the mechanism it exists to be. It came off `PENDING_QUESTIONS` on " +
    "`756682e` as stable across three passes; this round flipped it inside ONE round — `threw` on pass 1 with the host " +
    "healthy, `yes` on pass 2 in slide-trouble — while its TRIGGER answered `threw` both times and reported stable. So the " +
    "pair has been seen the other way round from the way this project describes it: the partner is the coin here and the " +
    "trigger held. Do not build on either value. Worth keeping as a pair rather than collapsing: a round where the two " +
    "disagree is exactly the evidence that this host's refusal of a held proxy is a state it moves through, not a rule.",
  "shapes-items-count-honest":
    "MEASURED AGAIN 2026-08-10: `unreadable` on the last SEVEN consecutive rounds (1fa0509 through 3d17165), not the three this " +
    "entry claimed. `short-0` has not appeared since 619d24b. The ANSWER never moved — this host will not tell a caller what is on " +
    "a slide — and the FORM now looks settled too, though `slideShapeList` still handles both and should, because nothing explains " +
    "why the form ever changed. Original note follows. " +
    "`short-0` again on 2026-08-09 (`619d24b`), against `unreadable` in the committed fixture — nine samples now, still alternating between the two forms with no trend. " +
    "The ANSWER is stable; its FORM is not. `unreadable` (2026-08-05), `short-0` (2026-08-08 on 2f1e8c4), `unreadable` again " +
    "(2026-08-08 on a546897), `short-0` again (2026-08-09 on 8bb9e8f), then `unreadable` on the last three (448ffc6, cfa1f50 and the round after it, which answered identically). Seven samples, two forms, no trend — though the last three held still, which is the longest it has. Every one of those says the same thing — this host will not tell a caller what is on a slide — and they " +
    "differ only in how the refusal arrives: a collection that throws, versus one that answers with zero items. Worth keeping apart " +
    "because the two want different code (a catch versus a corroborated count), and `slideShapeList` handles both for exactly this reason.",
  "shapes-items-via-positional-slide":
    "As its by-id partner above, and moving in step with it: `short-0` (2f1e8c4), then `not-listed` on every run since (a546897, d812d0c, 448ffc6, cfa1f50). Every run agrees with the " +
    "by-id form in the same run, which is the finding — the parent handle is not the variable. What varies is the host, run to run.",
  "group-children-via-getcount":
    "NO LONGER ONCE: eight of eight `unreadable` across the ten rounds to 2026-08-10 (the other two never put the question). " +
    'The hedge below — "two consistent answers from two routes is a strong hint, not a finding" — was right when it was written and ' +
    "is now overtaken: both routes into a group's children are refused, consistently, on every round that could ask. `contentShapes` " +
    "returning UNKNOWN_CONTENT for a grouped slide is the permanent answer. Original note follows. " +
    "ASKED AND ANSWERED ONCE, on a degraded host: `unreadable`, 2026-08-08. Its sibling `group-reports-its-children` answered `threw` " +
    "(\"The property 'items' is not available\") on a healthy round the same day, so BOTH routes into a group's children have now been " +
    "refused and `contentShapes` returning UNKNOWN_CONTENT for a grouped slide looks permanent rather than a gap. " +
    "Listed here rather than treated as settled because that round put only 17 of 27 questions — `getcount-populates-same-sync` itself " +
    "came back `no-scratch-slide` in it, having answered `yes, value=9` the round before. One sample from a host in that state is a " +
    "sample. Two consistent answers from two routes is a strong hint, not a finding.",
  "shape-add-held-slide-proxy":
    "THE `yes` PAIR HAS BEEN CAUGHT, AND THE REGIME IS NOT THE STATE — 2026-08-13 (`cd3b60c`). The note below " +
    "names its own honest limit: four agreeing pairs, all of them `threw` pairs, so what was shown was consistency " +
    "in the `threw` state rather than consistency in general, and a pair taken while the host answers `yes` 'cannot " +
    "be scheduled — it has to be caught'. This round caught it, and two more besides:\n" +
    "  pass 1   trigger `threw` 16.3s / partner `threw` 17.0s   (healthy)\n" +
    "  pass 2   trigger `yes`   33.9s / partner `yes`   34.4s   (collection-refused)   <- the missing one\n" +
    "  pass 3   trigger `threw` 55.6s / partner `threw` 56.8s   (collection-refused)\n" +
    "Seven pairs now, seven agreements, zero disagreements, and the `yes` state is paired half a second apart. A " +
    "fifty-fifty coin agrees half the time, so seven for seven is p=0.0078 against one. The reading the entry " +
    "already reached — at any instant this host has a DEFINITE answer, and the variation across a run is a state " +
    "changing rather than a coin landing — is now supported in BOTH states rather than one. " +
    "AND THE NEW HALF: `scripts/host-regimes.mjs` calls this question a COIN, correctly, because " +
    "`collection-refused` produced BOTH `yes` (pass 2) and `threw` (pass 3). Those two facts are not in tension — " +
    "together they say something neither says alone. There IS a definite state, it changes during a run, and the " +
    "probe's `regime` stamp is NOT that state. So the next question is not 'is it a coin' (it is not) but 'what is " +
    "the state', and `regime` has been eliminated as the answer. " +
    "The pair is also LOCKSTEP by that tool's other reading — trigger and partner move at the same pass boundary — " +
    "so they are one mechanism sampled twice, not two questions. Treat a flip in either as a flip in both. " +
    "Original notes follow. THE PARTNER HAS ANSWERED, AND IT IS NOT A COIN — 2026-08-11 (`756682e`). Three paired asks in one round, " +
    "TWO of them on later passes, and all three agreed:\n" +
    "  pass 1   trigger `threw` 37.8s  / partner `threw` 39.2s\n" +
    "  pass 2   trigger `threw` 89.0s  / partner `threw` 90.1s\n" +
    "  pass 3   trigger `threw` 127.5s / partner `threw` 129.1s\n" +
    "With the agreeing pair from `96461eb` that is four pairs, four agreements, zero disagreements. A fifty-fifty " +
    "coin agrees half the time, so four for four is p=0.0625 against it — not proof, and much stronger than anything " +
    "fifteen one-sample rounds could say. Read it as: at any given instant this host has a DEFINITE answer, and the " +
    "variation across a run is a state changing, not a coin landing. " +
    "The honest limit: all four pairs are `threw` pairs. Nothing yet pairs a `yes`, so what is shown is consistency " +
    "in the `threw` state rather than consistency in general. A pair taken while the host is answering `yes` is the " +
    "one still missing, and it cannot be scheduled — it has to be caught. " +
    "The later-pass tally moved too, and against the old story: this round answered `threw` on ALL THREE passes, so " +
    "later passes now stand at 3 x `yes` and 3 x `threw`. Pass 1 remains near-deterministic (18 of 19 `threw`). " +
    "Original notes follow. THE CLEAN SPLIT IS GONE — 2026-08-11 (`96461eb`), the round that first carried the partner. A LATER pass " +
    "answered `threw` for the first time (pass 3, 46.7s), so the tally below is no longer 3-of-3 either side:\n" +
    "  pass 1 : 18 x `threw`, 1 x `yes`   (19 observations)\n" +
    "  later  :  3 x `yes`,   1 x `threw` ( 4 observations)\n" +
    "Read that as: pass 1 is near-deterministic and the LATER passes are where the variability lives — which is a " +
    "different claim from two clean populations, and a better fit to every round on file. " +
    "THE PARTNER ANSWERED ONCE AND AGREED: trigger `threw` at 7.7s, `shape-add-held-slide-proxy-again` `threw` at " +
    "8.4s, seven tenths of a second apart on two different fresh slides. One agreeing pair is weak evidence against " +
    "a fast coin (a coin agrees half the time) and it is the LEAST informative pair available, because pass 1 is the " +
    "condition that barely varies. The pair worth having is on a LATER pass, and that round did not get one — the " +
    "pass-3 partner came back `no-scratch-slide`. The pair costs two scratch slides on the host least willing to " +
    "give them, which is the thing to fix if this is still open after another round or two. " +
    "Original notes follow. THE POPULATIONS WERE NEVER SEPARATED — 2026-08-11 (`7027f96`), and this supersedes both readings below. " +
    "Every observation this entry has ever recorded, sorted by WHICH PASS asked it:\n" +
    "  pass 1 (asked ~1-2s into the probe):  17 x `threw`, 1 x `yes`   (18 observations)\n" +
    "  any later pass (34s-80s in):           3 x `yes`,   0 x `threw` ( 3 observations)\n" +
    "The second row did not exist until 2026-08-11, because `PROBE_PASSES` shipped that day (780cf02) and every " +
    "round before it asked each question exactly ONCE. So the sixteen pre-3x observations are not sixteen samples " +
    "of a coin — they are sixteen samples of ONE condition, and the ten-in-a-row `threw` that the note below calls " +
    "a trend is ten rounds asking at the same moment and getting the same answer. Both 3x rounds agree: `threw` " +
    "first, `yes` on every later pass (R12 threw@2.1s then yes@79.4s; R13 threw@1.1s then yes@34.4s and yes@47.7s). " +
    "DO NOT read that as settled. Three readings still fit and this round cannot choose between them: the age of " +
    "the RUN, the pass number itself, or the state of the SCRATCH SLIDE — and the third is live, because the run " +
    "log shows this question WRECKS its own scratch slide every time it is asked (`giving up the scratch slide " +
    "this question wrecked`, three times in R13), so pass 1 meets a deck with no scratch history and later passes " +
    "do not. What is settled is that the variable is not the round, which is what fifteen one-sample rounds were " +
    "implicitly testing. The partner question that separates run-age from slide-state is the `Probe.follow` case " +
    "to write next: ask it twice in quick succession LATE in a run, once on a fresh scratch slide and once on one " +
    "the run has already used. One answer to each, one round, no reasoning. " +
    "The regime stamps ARE readable now (`regimeFrom` shipped in #390 and R13 is the first round on a build " +
    "carrying it — 24/19/6 across three regimes, moving back and forth, against 55-of-65-identical before). They " +
    "say `healthy` for the `threw` and `slide-trouble` for both `yes`. That is one round; it is consistent with " +
    "the scratch-slide reading and does not establish it. Original notes follow. SETTLED 2026-08-11 (`3223293`), and it is a coin after all. The first round to ask every question three times " +
    "got BOTH answers out of one host, one tab, 77 seconds apart: `threw` at 2.1s and `yes` at 79.4s, run never " +
    "restarted, build never changed. That is the observation this entry has wanted since the single `yes` of " +
    "2026-08-08, and it retires the reading below: a mechanism with one early outlier cannot reproduce its outlier " +
    "on demand inside one run. It is not a claim of fifty-fifty — two samples are two samples — but the variation is " +
    "WITHIN a session, so re-running whole rounds was never going to settle it, and the fifteen observations below " +
    "were fifteen one-sample rounds rather than a trend. " +
    "The regime stamped on those two samples (`healthy` for the `threw`, `collection-refused` for the `yes`) is NOT " +
    "yet evidence: the flag it came from latched 8.9s into a 110s probe, so `collection-refused` may mean no more " +
    "than `later`. Fixed the same day (`regimeFrom`), so the next round's stamps can be read. Original note follows. " +
    "MEASURED AGAIN 2026-08-10 over ten consecutive rounds (`scripts/host-history.mjs`): `threw` TEN times out of ten, " +
    "551ad42 through 3d17165. That makes fifteen of sixteen observations, with the single `yes` now eleven rounds and two " +
    "days behind — this entry called it a coin, and a coin does not do that. Read it as a mechanism with one early outlier " +
    "rather than a fifty-fifty, and keep the entry only because the outlier has never been explained. Original note follows. " +
    "ALTERNATES. Six observations: `threw` (2026-08-05), `threw` (2026-08-07), `yes` (2026-08-08 run a), `threw` (2026-08-08 run b), " +
    "`threw` (2026-08-08 run c, the sheet now committed), `threw` (2026-08-09 on d812d0c). Five of the six are `threw`. Earlier wording here said it flipped once, which " +
    "reads as though the newer value were the true one and the old one a mistake — it is not a sequence of corrections, it is a coin, and " +
    "five of six landings do not make the seventh a mechanism. " +
    "The fake keeps refusing held proxies, which is the safe direction: code that never holds one across a sync is correct whichever way the coin lands.",
  "shape-add-positional-slide-proxy":
    "A REGIME MAPPING THAT DID NOT REPRODUCE — 2026-08-13 (`cd3b60c`), and the failure is the useful part. Run " +
    "against the COMMITTED fixture (`1789749`), `scripts/host-regimes.mjs` read this question as EXPLAINED by " +
    "regime: `yes` in `slide-trouble` twice and `not-listed` in `collection-refused`, a mapping with a way to come " +
    "out wrong that did not. One round later it is a COIN — `collection-refused` produced BOTH `not-listed` " +
    "(pass 1, 19.3s) and `yes` (pass 2, 38.7s) in the same regime.\n" +
    "So the mapping was one round's shape, not a mechanism, which is exactly what that tool's footer warns and why " +
    "an `explained` verdict is a reason to RE-READ an entry rather than to rewrite it. Had this entry been rewritten " +
    "as 'a degradation you can test for' on the strength of the first run, it would have been wrong inside a day. " +
    "Two rounds of `explained` is the bar; one is a reason to look again. " +
    "Original notes follow. MEASURED AGAIN 2026-08-10 over ten rounds: `yes` on the last five consecutively, `not-listed` on three before that, and `threw` " +
    "NOT ONCE. So the face that made this dangerous — `threw`, the one that says the question was put and refused — has not appeared " +
    "in ten rounds, while `yes` is on its longest run. That does NOT retire the warning: `yes` is still exactly the answer that makes " +
    "a positional slide handle look like a way around the by-id refusals, and `shapes-items-via-positional-slide` (genuinely " +
    "unstable, still alternating) is why it would not help anyway. Original note follows. " +
    "A ninth observation on 2026-08-09 (`619d24b`): `not-listed` again, the third face, on a run that answered 21 of 28 — so it is not a symptom of a badly degraded round either. " +
    "THREE-SIDED, which is worth saying because this entry called it a coin. Eight observations: `yes`, `yes`, `threw`, `yes`, `yes`, " +
    "`threw`, `not-listed` (2026-08-08, `1fd6aa3`), `yes` (2026-08-09, `d812d0c`). The third face is not a variant of the other two — `not-listed` means the DECK'S " +
    "OWN SLIDE LIST did not contain the scratch slide's id, so the question was never really put, while `threw` means it was put and " +
    "refused. A host that will not list a slide it has just added is the same story as everything else here, arriving at a different door. " +
    "For the first six observations it did alternate in lockstep with its partner above and always opposite to it: `yes`, `yes`, `threw`, `yes`, `yes`, `threw`. " +
    "The more dangerous of the two, because `yes` is exactly the answer that makes a positional slide handle look like a way around the " +
    "by-id refusals — and four of the five samples say `yes`. A majority is not a mechanism. Whatever decides these two flips within a " +
    "single run (see the `no-scratch-slide` windows in any probe log), and until that is understood neither answer may be built on. " +
    "`shapes-items-via-positional-slide` is the reason it would not help anyway: a positional handle reads a shape collection exactly as " +
    "short-0 as a by-id one does.",
};

/**
 * Questions the committed real-host sheet cannot answer, because they were
 * added after it was taken.
 *
 * Declared for the same reason divergences are: a hole in the comparison is
 * worth seeing. `host-contract.test.ts` fails on an UNdeclared hole, so a probe
 * added without a re-run cannot quietly reduce what the gate covers, and fails
 * on a stale entry too, so the list shrinks the moment a newer sheet arrives.
 *
 * Every entry here is a reason to ask the owner for a probe run.
 *
 * EMPTY SINCE 2026-08-29, and it emptied all at once. The fixture had been the
 * 2026-08-12 capture, which answers 31 questions; every round since ~254
 * answers 39. So all ELEVEN entries had been answered — several of them for
 * more than a hundred rounds — while this table went on calling them unasked
 * and the probe went on spending scarce scratch slides re-asking them.
 *
 * That is the failure this table was built to prevent, arriving through the
 * one door it does not watch: it shrinks when a newer sheet lands, and nothing
 * made a newer sheet land. **The list is only as current as the fixture.**
 *
 * What the eleven turned out to say, over the 25 rounds to 2026-08-28:
 *
 *     AGREE WITH THE FAKE — no declaration needed
 *       shapes-by-index-vs-items            both-answer (a coin; see UNSTABLE)
 *       creationid-on-fresh-shape           absent          25 of 25
 *       creationid-survives-a-sync          no-creation-id  25 of 25
 *       creationid-survives-grouping        no-creation-id  25 of 25
 *       which-end-a-short-read-drops        all             25 of 25
 *       shape-resolve-held-slide-proxy      yes             24 of 24
 *
 *     CONTRADICT THE FAKE — now declared in KNOWN_DIVERGENCES, with tallies
 *       how-many-collection-reads-a-context-survives   short-at-1      23 of 25
 *       does-a-failed-group-poison-the-tag             tags-gone       25 of 25
 *       collection-read-poisons-the-creation-handle    refused         25 of 25
 *       how-many-syncs-a-creation-handle-survives      refused-after-1 23 of 25
 *       tag-through-refetched-shape                    no-id           23 of 25
 *
 * The three `creationid` answers are the load-bearing ones and they close a
 * design: this host reports PowerPointApi **1.10** in `requirementSets` and
 * does not populate `Shape.creationId` on any of the three questions, on any of
 * 25 rounds. Retiring the positional group-member mapping in favour of creation
 * ids is therefore not available here, and is not a matter of gating on 1.10.
 */
export const PENDING_QUESTIONS = {
  // EMPTY, AND THAT IS THE POINT OF THE REGISTER WORKING.
  //
  // Both entries added on 2026-08-30 are gone as of 2026-09-01, retired the way
  // the register asks: the code deployed, rounds 347-351 asked, and the fixture
  // was refreshed from round 351. `named-preset-resolves` answered `draws` in
  // 14 of 14 rounds and now lives in the committed sheet.
  //
  // `rotation-keeps-the-unrotated-box` answered too, and its answer is
  // `unreadable` — so it moved to KNOWN_DIVERGENCES rather than out of the
  // registers altogether, because the fake says `unrotated-box` and a
  // disagreement has to be declared somewhere. Read that entry before assuming
  // the question is closed: it is not answered, it is unanswerable on this
  // host, and the four passes that did read all agreed with the fake.
  //
  // A new probe belongs here between the commit that adds it and the round that
  // answers it, and nowhere else. `rounds-gate.mjs` prints any id left here
  // that the ARCHIVE has already answered, which is what stopped
  // `named-preset-resolves` sitting here for another fifty rounds.
};

/**
 * What each divergence would MEAN, so the report says something actionable
 * rather than just "these differ".
 *
 * Keyed by probe id. Absent is fine — the diff still reports the divergence,
 * it just cannot say what rests on it.
 */
const WHAT_IT_MEANS = {
  "addgroup-returns-usable":
    "Whether a group can be used in the sync that made it. WITHDRAWN: the 2026-08-04 'unreadable' came from asking for the id one sync after the group was made, and from grouping members that were themselves a sync old — so it measured proxy age, which three other questions already establish. Members are resolved in the grouping batch now and the id is asked for in it.",
  "shapes-items-via-positional-slide":
    "The partner to the question above, and the only one of the four contaminated answers that could not be cleaned up by re-resolving: a collection load is queued in one batch and read in the next by definition. If this reads back and the by-id form does not, the collection was never the problem and the parent handle was — which decides how every readback in `powerpoint.ts` should name its slide.",
  "load-isnullobject-populates":
    "`queueNullCheck` loads 'id' instead of 'isNullObject' precisely because the flag cannot be loaded by name. If a real host populates it, that whole comment is wrong for this host — and the workaround is merely harmless rather than necessary. ANSWERED: PowerPoint on the web (2026-08-04) said yes, and read the flag back as false. The negative is host-specific; the workaround stays because the host it was written for is real too.",
  "load-id-populates-isnullobject":
    "If a real host does NOT populate the flag from a real property load, `queueNullCheck` does not work and every `isLive` check is answering 'not live' for live objects. `isLive` treats unreadable as NOT live, so the failure mode is refusing to act on slides that are fine.",
  "shape-add-fresh-slide-proxy":
    "Whether this host will take a shape at all on a slide added moments ago, asked through a slide proxy resolved in the same sync as the add. ANSWERED: PowerPoint on the web (build a609c9c) said YES. A new slide takes shapes perfectly well — it is never the slide's newness that fails.",
  "shape-add-held-slide-proxy":
    "The same add through a slide proxy resolved one sync earlier — what Office.js has by then rewritten to `slides.getItem(id)`. ANSWERED: THREW, GeneralException, while the fresh and positional forms of the same add both worked. So it is the HOLDING that fails, not the id and not the slide. The fake models this unconditionally now (`expiringSlideHandle`), which is why the baseline beside it reads `threw` — the one question here whose expected answer is a refusal.",
  "shape-add-positional-slide-proxy":
    "The third way to name the same slide. ANSWERED: YES — so by-index is not the fix for anything; by-id was never the problem. Kept as the control that makes the pair above readable.",
  "shape-proxy-survives-one-sync":
    "office-js#2903. The fake keeps proxies alive by default, which is the kindness that hid a whole class of stale-proxy bug until a human found it in a real host. If a real host refuses a one-sync-old proxy, `applyWebProfile` should be the default rather than a named profile. ANSWERED, sideways: the 2026-08-04 self-test run threw `InvalidParam passed to GetItem(id)` at `ShapeCollection.getItem` while grouping a chart's shapes, five charts in a row — so on that host the answer is no. The probe's own attempt never reached the question.",
  "shapes-items-count-honest":
    "`faults.hollowReads` models a host answering SHORT without throwing — a readback asked about 19 shapes and was told 3. WITHDRAWN: the 2026-08-04 answer (`items` undefined) was about the handle the collection hangs off, not the collection. A collection read cannot avoid crossing a sync — the load is queued in one batch and read in the next — so this one could not be cleaned up by re-resolving, and got a partner instead: `shapes-items-via-positional-slide` asks the same thing through a positional parent. Read the two together or neither.",
  "tags-add-same-key-twice":
    "Re-editing a chart rewrites POWERCHART_CONFIG on the same shape every time. If a host appends rather than overwrites, a chart edited ten times carries ten configs and the reader picks one arbitrarily. WITHDRAWN: the 2026-08-04 'other — value=undefined' was the probe holding one shape proxy across four syncs, not an opinion about tag keys. Every write now goes through a shape resolved in its own batch, so the next sheet's answer is the first real one.",
  "tags-on-fresh-shape":
    "`faults.tagsUndefinedOn` models `.tags` coming back undefined, where reading `.add` throws SYNCHRONOUSLY and escapes the tagging loop — losing the config for every chart after it in the batch, not just the one.",
  "tag-through-refetched-shape":
    "The production path `finishCharts` actually uses: write POWERCHART_CONFIG through `shapes.getItemOrNullObject(id)`, where the id was read off a shape created a sync earlier. Rounds 29 and 30 both failed `same scale across the deck` with `InvalidParam passed to GetItem(id)` (5010) at `writing the chart's config tag`. `tags-on-fresh-shape` says the fresh shape's `.tags` is usable; this asks whether the id round trip is what the host refuses. `threw` means stop re-fetching; `yes` means the 5010 is about something else.",
  "delete-then-lookup":
    "`deleteSlideById` re-checks from a FRESH context because the same-context answer was not trusted. If a host answers honestly here, that second round trip is removable.",
  "group-children-via-getcount":
    "Added 2026-08-08, in response to the sheet taken that day. Its sibling `group-reports-its-children` asked through `group/shapes/items/id` and this host answered `threw` — \"The property 'items' is not available\", office-js#6363's exact signature — with the load queued in the sync that MADE the group, which is the friendliest form the question has. But the same sheet says `getcount-populates-same-sync: yes, value=9`: this host COUNTS a shape collection it will not LIST. Nobody has asked whether that holds for a GROUP's collection, and the answer decides something concrete — `contentShapes` returns UNKNOWN_CONTENT for every grouped slide, which is what makes the reconcile report a slide complete without counting it. A count is all it needs.",
  "group-reports-its-children":
    "The single most load-bearing answer here. A chart IS a group, and the readback measures whether a chart survived by counting what is inside it — so a host that groups successfully and then reports no children makes every chart read back as wreckage, and the repair pass 'fixes' charts that were never broken. WITHDRAWN: the 2026-08-04 PropertyNotLoaded was a nested load queued a sync after the group was made. It is queued in the grouping batch now.",
  "binding-names-shape-later":
    "Whether the repair pass can be given a handle that does not go through `ShapeCollection.getItem(id)`. Every 5010 this host throws is at that call, and it is what leaves a chart drawn and nameless — no group, no tag, nothing to settle one onto. A binding is made from the live Shape proxy inside the batch that created it, so it needs neither an id round trip nor a collection read, and the document persists it. A real `yes` means `settleAndTagChart` has a route it does not have today; a real `no` retires the idea. Watch the answer WORD, because five of them are different facts: `no-binding-api` is a missing 1.8 surface and says nothing about the idea; `add-threw` is `bindings.add` objecting on the spot; `commit-threw` is the host rejecting the batch that carried it, which counts as an answer ONLY because the probe commits the same batch without a binding first — see below; `unreadable` means the binding was made and then would not name its shape, the same refusal wearing a new coat; `yes` is the one that changes what can be built. FIRST REAL SIGNAL, 2026-08-09 evening (`2a44f64`): the commit came back `UnexpectedError` in 1.3 seconds while `shape-add-fresh-slide-proxy` answered `yes` in the same sheet. That points at the binding, but it is cross-question inference on a host that flaps between minutes, so the probe was given its own control arm rather than the reading being written down as fact. One sample, and the question stays open until a sheet answers it with the control in place. WHAT MICROSOFT'S OWN DOCS ADD (searched 2026-08-13, `bind-shapes-in-presentation`): the API is real and current — `bindings.add(shape, BindingType.shape, id)` and `Binding.getShape()` are both PowerPointApi 1.8, and this host reports 1.10, which is consistent with `commit-threw` rather than `no-binding-api` and means the surface is not the problem. The documented FLOW differs from the probe's in one respect worth a variant: Microsoft creates the shape, sets its fill, and retrieves through `bindings.getItem(id).getShape()` in a SEPARATE `PowerPoint.run`, whereas the probe binds inside the batch that created the shape. So `commit-threw` may be the host refusing to bind a shape the document does not yet have, rather than refusing bindings at all. A variant that binds in a LATER batch than the one that drew the shape is untried, and is the partner question this entry wants next. Do not read that as a prediction — it is the one difference between a flow Microsoft publishes and the flow that failed here.",
  "getitemat-past-end":
    "Nothing in this repo currently depends on the answer — it is here to find out before something does.",
  "untrack-available":
    "The fake does not implement `untrack`, so every `untrack()` call in `powerpoint.ts` is a no-op under test and the proxy-release path is entirely unexercised. A real host saying 'yes' does not fix that; it means the path is real and still untested.",
};

/**
 * Find the answer sheet inside whatever arrived — the sheet itself, or a whole
 * round's file with the sheet nested in it.
 *
 * Split out of `answersOf` because the ANSWERS were being unwrapped and the
 * HEADER was not: `host-diff.mjs` read `source` and `requirementSets` off the
 * outer object, so a round file — the shape the pane actually writes, and the
 * one `answersOf` exists to accept — printed `REAL HOST ?` and
 * `requirement sets: unknown` above a page of real answers. Which host, and
 * which API versions it offers, is not decoration here: the whole binding lead
 * turns on whether PowerPointApi 1.8 is present, and the file said 1.1 through
 * 1.10 while the report said it did not know.
 *
 * "Run the whole round" writes one file for both halves precisely so there is
 * one thing to send; every reader of it has to unwrap the same way.
 */
/**
 * Does this look like an answer sheet, under either name it has had?
 *
 * BOTH SPELLINGS, and the old one is not legacy cruft — 39 of the last 40
 * archived rounds carry `powerchart-host-answers`. Accepting only the new name
 * would make this file reject 257 rounds of evidence, which is the archive the
 * whole project reasons from. Duplicated from `host-probe.ts` rather than
 * imported because this is a plain `.mjs` tool with no build step; the two are
 * kept in step by the test that reads a real archived round.
 */
function isHostAnswersKind(kind) {
  return kind === "ssf-charts-host-answers" || kind === "powerchart-host-answers";
}

export function sheetOf(file) {
  if (isHostAnswersKind(file?.kind)) return file;
  if (file?.hostAnswers) return sheetOf(file.hostAnswers);
  return file ?? null;
}

/** Read a sheet, whichever shape it arrived in. */
/**
 * Answers this project has RENAMED, keyed by the question that used to give them.
 *
 * READ, NEVER WRITTEN BACK. A captured sheet is evidence, and this file already
 * settles the principle one screen away, about the 2026-08-27 `kind` rename:
 * "Rewriting a captured artefact to match a rename falsifies the evidence." So
 * an old capture keeps saying `all`, and this table is how today's code
 * RECOGNISES that as the same observation rather than a disagreement.
 *
 * Without it, renaming an answer silently reports a host divergence that never
 * happened — the fake would claim `not-a-short-read` against a fixture saying
 * `all` and `KNOWN_DIVERGENCES` would gain an entry describing a rename as a
 * behavioural difference, which is worse than the staleness it replaced.
 *
 * Entries can be dropped once no committed or archived sheet still carries the
 * old word. Nothing depends on that happening.
 */
export const RENAMED_ANSWERS = {
  // 2026-09-01. `all` and `none` were this probe's way of saying its question
  // did not arise — nothing was dropped, or the collection listed shapes that
  // were not ours. They ranked as named answers and locked the row, and `all`
  // matched the fake's `all`, so 87 rounds recorded agreement about which end a
  // short read drops on a question neither side had answered.
  "which-end-a-short-read-drops": { all: "not-a-short-read", none: "none-of-ours" },
};

export function answersOf(file) {
  const sheet = sheetOf(file);
  if (isHostAnswersKind(sheet?.kind) && Array.isArray(sheet.answers)) {
    return Object.fromEntries(sheet.answers.map((a) => [a.id, RENAMED_ANSWERS[a.id]?.[a.answer] ?? a.answer]));
  }
  // A bare map, e.g. the committed baseline.
  if (sheet && typeof sheet === "object" && !Array.isArray(sheet)) return sheet;
  return null;
}

/**
 * Answers that mean the question was never put, so they are never divergences.
 *
 * The probe's own vocabulary, kept in step with `NOT_ASKED` in
 * `src/render/host-probe.ts`. No probe can produce any of these words as an
 * answer, which is what makes them safe to read this way.
 *
 * "Kept in step" was a hope, and it drifted: `not-asked` — what the probe's mute
 * breaker records when it abandons the rest of a sheet — was added on the probe
 * side and never here. Every question that breaker gives up on was therefore
 * compared against the fake and reported as a real-host DIVERGENCE, which is the
 * exact failure the third bullet of `diffAnswers` was written for ("PowerPoint
 * refused eight setups on 2026-08-04 and this tool reported eight host
 * divergences from questions nobody had asked"). The same bug, wearing a word
 * that did not exist yet. No round has tripped the breaker so far, so nothing
 * published has been wrong — it would have fired the first time a host degraded
 * badly enough to matter. `test/host-probe.test.ts` now asserts the two sets are
 * equal, so the next word cannot be added to one side alone.
 */
export const NEVER_ASKED = new Set(["no-scratch-slide", "no-scratch-shape", "no-named-slide", "not-asked"]);

/**
 * Answers that mean the question WAS put and produced nothing to name.
 *
 * The second tier, and the same arrangement as `NEVER_ASKED` above for the same
 * reason: this is a plain .mjs tool that imports nothing from the TypeScript, so
 * the vocabulary is copied and a test asserts the copies stay equal rather than
 * trusting anyone to remember. `UNINFORMATIVE` in `src/render/host-probe.ts` is
 * the original; each of these words cost this project rounds of real answers
 * before it was classified there.
 */
export const UNINFORMATIVE_ANSWERS = new Set(["other", "unreadable", "silent", "not-a-short-read", "none-of-ours"]);

/**
 * Compare two answer sheets.
 *
 * Three ways a question can fail to be a match, and all three used to be one:
 *
 * - `onlyReal` / `onlyFake` — one side was never asked. A gap, not agreement.
 * - `notAsked` — the real host never got far enough to answer. The probe run
 *   says so in a word no probe can produce, because it learned the hard way
 *   what happens otherwise: PowerPoint on the web refused eight setups on
 *   2026-08-04 and this tool reported eight host divergences from questions
 *   nobody had asked.
 * - `differ` — the only one that is actually a finding.
 */
/**
 * Which pending questions the ARCHIVE says the host has already answered.
 *
 * `PENDING_QUESTIONS` instructs, in its own comment, that an entry be deleted
 * once the host answers — "an id left here after the host has answered is the
 * fixture going stale in writing, which is the failure this register exists to
 * make impossible to do quietly". Its enforcing gate could never notice: it
 * compares against the committed fixture, and an id is legitimately pending
 * precisely BECAUSE the fixture predates it. So the gate is green while the
 * question is unanswered AND while it has been answered fifty times, and cannot
 * tell those apart — the verified-versus-not-attempted failure again, in the one
 * register built to prevent it.
 *
 * Measured 2026-09-01: `named-preset-resolves` had answered `draws` in 11 of 11
 * rounds, 33 of 33 samples, stable in every one, across five builds — as settled
 * as anything in the archive — while the register still called it unanswered and
 * every gate stayed green.
 *
 * The archive can see what the fixture cannot, so it is asked here. A `stable`
 * answer is not required: this reports what the host SAID, and whether it said
 * it consistently is the reader's next question, not a reason to stay quiet.
 */
export function pendingAlreadyAnswered(rounds, pending = PENDING_QUESTIONS) {
  const seen = new Map();
  for (const round of rounds ?? []) {
    const answers = answersOf(round?.hostAnswers);
    if (!answers) continue;
    for (const id of Object.keys(pending)) {
      const a = answers[id];
      // A weak word is not an answer — `unreadable` and `silent` mean the
      // question was put and produced nothing, which is why they were made
      // displaceable in the first place. Retiring a question on one of those
      // would be the forgetting this register exists to prevent, arriving
      // through the door marked "the archive says so".
      if (a === undefined || NEVER_ASKED.has(a) || UNINFORMATIVE_ANSWERS.has(a)) continue;
      const row = seen.get(id) ?? { id, answers: new Map(), rounds: 0 };
      row.rounds++;
      row.answers.set(a, (row.answers.get(a) ?? 0) + 1);
      seen.set(id, row);
    }
  }
  return [...seen.values()].map((r) => ({
    id: r.id,
    rounds: r.rounds,
    answers: [...r.answers].sort((a, b) => b[1] - a[1]).map(([answer, n]) => ({ answer, n })),
  }));
}

export function diffAnswers(real, fake) {
  const ids = [...new Set([...Object.keys(real), ...Object.keys(fake)])].sort();
  const agree = [];
  const differ = [];
  const notAsked = [];
  const onlyReal = [];
  const onlyFake = [];
  for (const id of ids) {
    const r = real[id];
    const f = fake[id];
    if (r === undefined) onlyFake.push(id);
    else if (f === undefined) onlyReal.push(id);
    else if (NEVER_ASKED.has(r)) notAsked.push({ id, why: r });
    else if (r === f) agree.push(id);
    else differ.push({ id, real: r, fake: f, means: WHAT_IT_MEANS[id] });
  }
  return { agree, differ, notAsked, onlyReal, onlyFake };
}
