/**
 * Turning what the host threw into a line a person can act on.
 *
 * Lifted out of `powerpoint.ts` on 2026-09-06, BYTE-FOR-BYTE — the only edit
 * in the move is `const STEP_KEY` becoming `export const STEP_KEY`, because
 * `step()` stayed behind and needs it. Nothing here talks to Office.js: every
 * function takes an `unknown` and returns text, which is why it can be read,
 * tested and argued about without a host in front of you. `powerpoint.ts` was
 * over 13,000 lines and these 149 of them had no reason to be among them.
 *
 * `errorText`, `stepOf` and `trimDebugInfo` are still re-exported from
 * `powerpoint.ts`, so no caller had to change. The move is the whole change,
 * and the test count either side of it is the same.
 */
import { isStaleBuild, StaleBuildError } from "./lazy";
/**
 * Where an error was when it escaped — attached to the error itself.
 *
 * Read off by `errorText`, so every place that reports an error says WHICH
 * phase it came out of, not only what the host called it.
 */
export const STEP_KEY = "__powerchartStep";

/**
 * How many of `fullStatements` a trace keeps at each end. See `trimDebugInfo`.
 *
 * Ten and ten, against the batch this project actually issues: a chart is drawn
 * `SHAPES_PER_SYNC` at a time and each shape costs several statements, so ten at
 * the head carries the handles the batch opened with and ten at the tail carries
 * what it was doing when it stopped, without the other eight shapes' noise.
 */
const FULL_STATEMENT_ENDS = 10;

/**
 * Cut `debugInfo.fullStatements` down to something a run log can carry.
 *
 * `extendedErrorLogging` (see `traceEnvironment`) is what makes Office.js fill
 * that array in at all, and it fills it in with the WHOLE batch. One round held
 * 66 of these errors; at a full batch each that is a file nobody can send.
 *
 * BOTH ends, and the first version's tail-only rule is the reason. It assumed
 * the failing statement is last. It is not: Office.js reports it separately in
 * `statement`, and `surroundingStatements` centres its `>>>>>` marker on it —
 * and in the 2026-08-07 round that marker sat on the FIRST statement of the
 * batch while `fullStatements` came back "… 37 earlier statement(s) dropped".
 * The one line worth reading was the one line thrown away.
 *
 * The head is also where the batch's opening handles are, and those are the
 * whole question. A slide printed as `slides.getItem("282#…")` carrying the
 * annotation "originally getItemOrNullObject" is what finally settled whether a
 * printed `getItem` means a held handle. It does not — Office.js annotates the
 * call each path was created by, and the rewrite is just how a resolved proxy
 * prints.
 *
 * Pure, and separate from `errorText`, so the trimming can be tested without a
 * PowerPoint anywhere near it.
 */
/**
 * The longest a single statement is allowed to be in a recorded error.
 *
 * ONE statement can carry a whole deck. Round 148's `insert on top of an
 * earlier run` failed with `GeneralException | errorLocation=
 * Microsoft.Office.PrivateApiService`, and the diagnosis — the location, the
 * code, the call that failed — arrived behind ~100KB of base64, because
 * `insertSlidesFromBase64` takes the file as an argument and Office echoes the
 * argument back in the statement. Trimming the COUNT of statements did nothing
 * about it: the payload was in one of the ten that were kept.
 *
 * The call is the diagnosis. The bytes it was handed are not, and every one of
 * them was also being written into the round archive.
 */
const STATEMENT_MAX_CHARS = 200;

function trimStatement(s: unknown): unknown {
  if (typeof s !== "string" || s.length <= STATEMENT_MAX_CHARS) return s;
  return `${s.slice(0, STATEMENT_MAX_CHARS)}… ${s.length - STATEMENT_MAX_CHARS} more char(s)`;
}

export function trimDebugInfo(info: unknown): unknown {
  if (!info || typeof info !== "object") return info;
  const rec = info as Record<string, unknown>;
  const full = Array.isArray(rec.fullStatements) ? rec.fullStatements : undefined;
  const tooMany = full !== undefined && full.length > FULL_STATEMENT_ENDS * 2 + 1;
  const long = (v: unknown) => typeof v === "string" && v.length > STATEMENT_MAX_CHARS;
  const tooLong =
    long(rec.statement) ||
    (Array.isArray(rec.surroundingStatements) && rec.surroundingStatements.some(long)) ||
    (full?.some(long) ?? false);
  // UNTOUCHED WHEN THERE IS NOTHING TO TOUCH — a caller may compare identity,
  // and "trimmed" should mean something was actually dropped.
  if (!tooMany && !tooLong) return info;

  const kept = tooMany
    ? [
        ...full!.slice(0, FULL_STATEMENT_ENDS),
        `… ${full!.length - FULL_STATEMENT_ENDS * 2} statement(s) dropped from the middle`,
        ...full!.slice(full!.length - FULL_STATEMENT_ENDS),
      ]
    : full;
  return {
    ...rec,
    ...(rec.statement !== undefined ? { statement: trimStatement(rec.statement) } : {}),
    ...(Array.isArray(rec.surroundingStatements)
      ? { surroundingStatements: rec.surroundingStatements.map(trimStatement) }
      : {}),
    ...(kept ? { fullStatements: kept.map(trimStatement) } : {}),
  };
}

/**
 * The phase label `step` attached to an error, if it carries one.
 *
 * Exported because a caller sometimes has to tell WHICH operation timed out,
 * not merely that one did. `isTimeout` cannot: every bounded wait produces the
 * same kind of error, so a scenario that catches one and explains it is
 * explaining a guess. The self-test's `edit the chart the user selected` did
 * exactly that — it reported "the host stopped answering selection calls" for a
 * run whose trace says `gave up waiting what=drawing shapes 1-10 of 24`, and
 * sent the reader to two selection bugs that had nothing to do with it.
 */
export function stepOf(err: unknown): string | undefined {
  if (!err || typeof err !== "object") return undefined;
  const at = (err as Record<string, unknown>)[STEP_KEY];
  return typeof at === "string" ? at : undefined;
}

/**
 * Everything an Office.js error knows. A RichApi.Error's `message` is usually
 * generic ("An internal error has occurred"); the useful part — the failing
 * command and why — lives in `code` and `debugInfo`, which a plain String(err)
 * silently drops.
 */
export function errorText(err: unknown): string {
  if (!err || typeof err !== "object") return String(err);
  // A CHUNK THE SERVER NO LONGER HAS, said as a sentence and nothing else.
  //
  // Handled at the funnel rather than at each call site, so the pane, the round
  // self-test and the trace all say the same thing. `StaleBuildError` already
  // carries the sentence; the raw form is caught too, because an import added
  // later without `lazy` would otherwise report a URL to somebody who cannot do
  // anything with one.
  //
  // Returned WITHOUT `code=` or `debugInfo=`: those exist to place an Office.js
  // refusal, and there is no host refusal here. Appending them to a message
  // that is already an instruction just buries it.
  if (isStaleBuild(err)) {
    return err instanceof StaleBuildError ? err.message : new StaleBuildError("one of its parts", err).message;
  }
  const e = err as { message?: string; code?: string; debugInfo?: unknown };
  const bits = [e.message ?? String(err)];
  // The phase the add-in was in, when a `step` recorded one. Office.js says
  // what it refused; this says what we were doing — and a report with only the
  // first half is what made three real-host failures take a session to place.
  const at = (err as Record<string, unknown>)[STEP_KEY];
  if (typeof at === "string") bits.push(`at=${at}`);
  if (e.code) bits.push(`code=${e.code}`);
  if (e.debugInfo) {
    try {
      bits.push(`debugInfo=${JSON.stringify(trimDebugInfo(e.debugInfo))}`);
    } catch {
      /* not serialisable — the message and code still carry */
    }
  }
  return bits.join(" | ");
}
