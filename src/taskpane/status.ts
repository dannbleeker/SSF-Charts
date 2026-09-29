/**
 * The pane's status strip: the note, its colour, the progress bar, the Stop
 * button and the elapsed readout.
 *
 * ── WHY THIS IS ITS OWN MODULE ──────────────────────────────────────────────
 * Both `app.ts` and `harness-ui.ts` write to this strip, and until this file
 * existed the second of them reached it through a module-level `wired` holder:
 * `harness-ui.ts` took `note`, `setProgress` and `noteHostActivity` as injected
 * dependencies, three of the seven it asked for, purely because they lived in
 * `app.ts` and importing them back would have made a cycle. Three of the seven
 * were therefore parameters describing a fact about file layout rather than a
 * fact about the panel.
 *
 * A module both of them import instead of one of them owning it removes that.
 * `i18n.ts` imports nothing, so this file's only import cannot close a cycle,
 * and the dependency direction is now stated rather than inverted.
 *
 * ── KEEP IT THAT WAY ────────────────────────────────────────────────────────
 * `test/import-cycles.test.ts` deliberately ignores dynamic `import()`, so once
 * `app.ts` loads `harness-ui.ts` dynamically a static `import { note } from
 * "./app"` inside the panel would be a real runtime cycle that the guard cannot
 * see. Nothing here may import `./app`, and neither may `harness-ui.ts`.
 *
 * ── WHAT DID NOT MOVE ───────────────────────────────────────────────────────
 * `guard()` stayed in `app.ts`. It is the pane's action wrapper — it disables
 * buttons, counts host work and runs the insert — and it USES the strip rather
 * than being part of it. Everything below is DOM plus the decisions about what
 * the DOM is allowed to say.
 */
import { t } from "./i18n";

/**
 * The pane's live region. Declared here rather than beside the other DOM
 * handles in `app.ts` because `note` is its only reader — it was 62 lines above
 * the code that used it, in a block about the gallery.
 */
const hostNote = document.getElementById("host-note") as HTMLElement;

/**
 * The strip itself, its progress bar, its elapsed readout and its Stop button.
 * Private: everything that writes to them is below, which is the point.
 */
const statusStrip = document.getElementById("status-strip");
const statusBar = document.getElementById("status-bar");
const statusElapsed = document.getElementById("status-elapsed");
const statusStop = document.getElementById("status-stop") as HTMLButtonElement | null;

/**
 * Show or hide the Stop button for the action in flight.
 *
 * Reset on every show: the button is reused across actions, and one left
 * disabled reading "Stopping…" would greet the next action as already stopping.
 */
export function showStop(on: boolean) {
  if (!statusStop) return;
  if (on) {
    statusStop.disabled = false;
    statusStop.textContent = t("Stop");
  }
  statusStop.toggleAttribute("hidden", !on);
}

/**
 * How many times the pane has SETTLED — posted a note that is not "busy".
 *
 * `guard()` prints "Done." only when the action it ran did not report an end
 * state of its own, and it used to detect that by comparing the note text
 * against the busy text it had posted. Any progress note broke the comparison,
 * and an insert always ends on one: `phaseNote("done")` writes "Working… done",
 * which is still busy. So the text no longer matched, "Done." was skipped, and
 * the pane was left showing a blue busy note above a progress bar that slid on
 * forever — the action had finished, and nothing said so.
 *
 * Counting settlements asks the question that actually matters — "did this
 * action reach an end state?" — instead of inferring it from wording, which
 * also makes it immune to the language the note is rendered in.
 */
export let settledNotes = 0;

/**
 * Write the host note together with its status colour. The colour is a
 * parameter rather than an afterthought because only guard() used to set it,
 * so every other message inherited whatever the previous action left behind —
 * an "Invalid JSON" error rendered in the success green.
 */
export function note(
  text: string,
  status: "ok" | "err" | "busy" | "none" = "none",
  params?: Record<string, string | number>,
) {
  /**
   * THE STRIP OPENS FIRST, AND THE ORDER IS THE WHOLE POINT.
   *
   * `#host-note` is the pane's live region (`role="status"`, `aria-live`), and
   * `.status-strip[hidden]` is `display: none` in taskpane.css — an author rule,
   * so the collapse is real and not just the UA default. Writing the text before
   * lifting `hidden` therefore mutated a live region inside a `display: none`
   * subtree, and a live region is only announced for a change made while it is
   * rendered: content that appears together WITH the region counts as its
   * initial content, not as an update. The strip ships `hidden` in the markup
   * and `note("")` at boot leaves it that way, so the message this cost was the
   * FIRST of every session — "Working…" on the user's first insert.
   *
   * Un-hiding first makes the write a change to a rendered region, which is what
   * every platform guidance asks for. NOT VERIFIED AGAINST A SCREEN READER —
   * there is none in this harness; what is tested is the DOM ordering, which is
   * the half that is ours. The same ordering is applied to the other live region
   * this pane collapses, `#slow-offer` — see `offerOwnSlide`.
   *
   * Clearing runs the other way round on purpose: hide, then blank. There is
   * nothing to announce about a note being taken away.
   *
   * The text itself goes through the runtime translator so a localized pane
   * announces it in the user's language. `text` is the English source string (an
   * EN catalogue key), optionally carrying {placeholders} that `params` fills
   * after translation.
   */
  statusStrip?.toggleAttribute("hidden", !text);
  hostNote.textContent = t(text, params);
  hostNote.className = status === "none" ? "hint" : `hint status-${status}`;
  statusBar?.toggleAttribute("hidden", status !== "busy");
  if (status !== "busy") {
    setProgress(null);
    settledNotes++;
  }
}

/**
 * How far along, when we honestly know: a fraction for work we complete in
 * steps, "busy" for work we hand to PowerPoint in one go.
 *
 * A single insert is ONE context.sync() and Office.js reports nothing until it
 * lands, so any percentage there would be invented — and a bar that sticks at
 * 99% is a worse lie than no bar. Chunked work (the demo deck) really does know,
 * so it gets a real one.
 */
export function setProgress(p: number | "busy" | null) {
  if (!statusBar) return;
  const fill = statusBar.querySelector("i");
  if (p === "busy") {
    statusBar.classList.add("indeterminate");
    if (fill) fill.style.width = "";
  } else {
    statusBar.classList.remove("indeterminate");
    if (fill) fill.style.width = p == null ? "0" : `${Math.round(Math.max(0, Math.min(1, p)) * 100)}%`;
  }
}

/**
 * The last moment anything was heard from a run — see `noteHostActivity`.
 *
 * Module-level rather than passed in, because the two things that need it sit
 * on opposite sides of a module boundary: the trace subscriber in
 * `harness-ui.ts`, which sees every step a run produces, and the elapsed ticker
 * below, which is the only thing on screen while a sync is outstanding.
 */
let lastHostActivity = 0;

/** A run just did something. Bumped per traced step. */
export function noteHostActivity(now = Date.now()): void {
  lastHostActivity = now;
}

/**
 * How long a run may go silent before the pane stops implying it is working.
 *
 * Generous on purpose. A single draw batch on PowerPoint web has been measured
 * at ~17 seconds and a stalled-but-alive sync at 45, so anything under a minute
 * would cry wolf on a host that is merely slow — the failure mode this replaces
 * cuts the other way and is worse, but a warning nobody believes is no warning.
 */
const SILENT_RUN_MS = 60_000;

/**
 * What the elapsed readout should say, given how long the run has been going
 * and how long since anything happened.
 *
 * Pure, and exported, because it is the whole of a decision this project has
 * now watched go wrong on a real host three times: PowerPoint dies, the task
 * pane survives — it is a separate frame, and the `PowerPoint.run` promise it
 * is waiting on simply never settles — and the pane counts upward forever under
 * the word "Working…". The owner is left watching a number climb with no way to
 * tell it from a slow chart. There is no error to catch here and no timeout
 * that helps, because nothing rejects; silence is the only evidence there is.
 */
export function elapsedLabel(elapsedMs: number, silentMs: number): string {
  const secs = `${Math.round(elapsedMs / 1000)}s`;
  return silentMs >= SILENT_RUN_MS ? `${secs} · silent for ${Math.round(silentMs / 1000)}s` : secs;
}

/**
 * Count the seconds while the host works. It is the only number we can report
 * mid-sync, and on a host that takes 20s to draw a chart, a number that moves
 * is the difference between "working" and "dead".
 *
 * Except when it is not, which is why the readout also carries silence. See
 * `elapsedLabel`.
 */
let elapsedTimer: ReturnType<typeof setInterval> | undefined;
let saidSilent = false;
export function startElapsed() {
  const t0 = Date.now();
  stopElapsed();
  noteHostActivity(t0);
  saidSilent = false;
  const tick = () => {
    const now = Date.now();
    const silentMs = now - lastHostActivity;
    if (statusElapsed) statusElapsed.textContent = elapsedLabel(now - t0, silentMs);
    // Said once, not every second: the note area is shared, and a message that
    // rewrites itself every tick is one nobody reads. It also must not be an
    // "err" — nothing has failed as far as this pane knows, and claiming
    // otherwise over a merely slow host is the mistake in the other direction.
    if (!saidSilent && silentMs >= SILENT_RUN_MS) {
      saidSilent = true;
      note(
        // "IF a test run was in progress" — this fires from `guard()`, which
        // wraps EVERY pane action, an ordinary Insert included. It used to end
        // "The run's steps are saved either way and *Download the crashed run*
        // will offer them", which names a run an ordinary user never started and
        // a button they cannot see: `beginCrashLog` is called only from the
        // three harness paths, and the control lives in Automation ▸ Testing.
        // Being pointed at a missing button while your slide is frozen is worse
        // than being told nothing. The first two sentences are the ones that
        // help, and they are unchanged.
        "PowerPoint has not answered for {secs}s. Look at the slide area: if PowerPoint is showing *Sorry, we ran into a problem*, click Refresh there — nothing behind that dialog can answer. If a test run was in progress, its steps are saved and *Download the crashed run* will offer them.",
        "busy",
        { secs: Math.round(silentMs / 1000) },
      );
    }
  };
  tick();
  elapsedTimer = setInterval(tick, 1000);
}
export function stopElapsed() {
  clearInterval(elapsedTimer);
  elapsedTimer = undefined;
  saidSilent = false;
  if (statusElapsed) statusElapsed.textContent = "";
}

/**
 * Wire the Stop button, and let it say so the moment it is clicked.
 *
 * A registrar rather than an exported DOM handle, so all four elements of the
 * strip stay private to this file — an exported `statusStop` would put the
 * button's appearance back under two owners, which is the shape this module
 * exists to end. Whether the button is guarded is the caller's decision and
 * stays at the call site.
 *
 * The label changes immediately, before the handler's work can finish. The
 * batch already handed to PowerPoint still has to come back — up to
 * BATCH_TIMEOUT_MS — and without this the pane looks like it ignored the click
 * for as long as that takes.
 */
export function onStopRequested(handler: () => void): void {
  statusStop?.addEventListener("click", () => {
    handler();
    if (statusStop) {
      statusStop.disabled = true;
      statusStop.textContent = t("Stopping…");
    }
  });
}
