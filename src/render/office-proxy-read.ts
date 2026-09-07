/**
 * Reading an Office.js proxy without letting it take the draw down.
 *
 * Lifted out of `powerpoint.ts` on 2026-09-07, BYTE-FOR-BYTE except that each
 * function gained an `export`. Every one of them exists for the same reason:
 * an unloaded property THROWS `PropertyNotLoaded` on access rather than
 * answering undefined, and a host that half-answers is the normal case here,
 * not the exception. So each turns a refusal into a value a caller can branch
 * on.
 *
 * NO OFFICE.JS CALLS AND NO MODULE STATE. They take a shape of object and
 * return a value; that is what makes them safe to move and safe to test
 * without a host. The stall/timeout cluster was the first candidate for this
 * extraction and was rejected for the opposite reason — `withTimeout`
 * increments `deadlinesFired`, which `host-probe.ts` also reads, so moving it
 * would have put a mutable counter across three modules.
 *
 * 155 call sites in `powerpoint.ts` and none anywhere else, so nothing is
 * re-exported: the import list at the top of that file is the whole change.
 */
/**
 * Release a proxy object's client-side memory once its loaded values have been
 * read. Office.js keeps every proxy touched in a `run` alive until the context
 * is disposed, and the docs call out a "noticeable performance benefit when
 * using large numbers of proxy objects" from untracking — a deck-wide scan
 * creates one shape proxy and one tag proxy PER shape ON EVERY SLIDE, which is
 * exactly that case. Best-effort: a null-object proxy may not expose untrack.
 */
export function untrack(obj: unknown): void {
  try {
    (obj as { untrack?: () => void })?.untrack?.();
  } catch {
    /* not a tracked proxy on this host */
  }
}

/**
 * Put a `getItemOrNullObject` proxy into the next sync, so that `isNullObject`
 * answers once that sync lands.
 *
 * `load("isNullObject")` does NOT do this, and that is the whole reason this
 * function exists. It reads like it should — the property is right there on the
 * proxy and `load` takes any name — but `isNullObject` is not a property the
 * host holds. It is a flag Office.js sets from the RESPONSE to a load of real
 * properties. Ask for it by name and the load selects nothing the host knows,
 * the proxy takes no part in the sync, and reading the flag afterwards throws
 * `PropertyNotLoaded` — with no `errorLocation`, because the getter lives on
 * Office.js's base class and does not know which type it is standing on.
 *
 * That is what the self-test's "edit a chart on the visible slide" died on in
 * PowerPoint on the web: an in-place update resolved the chart's slide, asked
 * for `isNullObject`, and threw reading it back — before it had deleted
 * anything, so the update simply refused. Every other resolve in this file
 * happens to load a REAL property (`value` on a tag, `left,top` on a shape) and
 * has always worked, which is why the failure was confined to the paths that
 * had nothing else to ask for.
 *
 * `id` because every proxy resolved this way is a Slide or a Shape, and both
 * carry one. The value is not used — being in the sync is the point.
 *
 * One correction from a real host, since the paragraphs above are a claim about
 * Office.js and it turns out not to hold everywhere. The host probe's first
 * question asks exactly this, and PowerPoint on the web answered **yes**: a
 * `load("isNullObject")` there populated the flag and read back `false`. So the
 * negative is host-specific rather than universal, and on that host this
 * function is a harmless no-op instead of a necessary workaround. Keep it: the
 * host where it was necessary is also a real one, loading a real property is
 * correct on both, and the cost is a property nobody reads.
 */
export function queueNullCheck(proxy: { load(propertyNames: string): void }): void {
  proxy.load("id");
}

/**
 * A collection's `items`, or `undefined` when the host never answered the load
 * that was queued for it.
 *
 * PowerPoint on the web answers a shape-collection read short (see the repair
 * pass's `shapesSeen` counters, and `faults.hollowReads`) — and at the limit it
 * answers with nothing at all. Office.js then leaves the collection unpopulated
 * and the plain `.items` read throws `PropertyNotLoaded` at
 * `ShapeCollection.items`, which took the whole deck-wide chart scan down with
 * it: one silent slide, and Same Scale reported a thrown error instead of
 * rescaling the charts it could see.
 *
 * A collection that did not answer is a collection we know nothing about, so
 * every caller has to decide what to do with the gap rather than inherit a
 * throw. None of them may report it as "no charts here" without saying so.
 */
export function loadedItems<T>(collection: { items: T[] }): T[] | undefined {
  try {
    const items = collection.items;
    return Array.isArray(items) ? items : undefined;
  } catch {
    return undefined;
  }
}

/**
 * A loaded property, or `undefined` when the host did not answer for it.
 *
 * Same hazard as `loadedItems`, one level down: a proxy whose load was queued
 * in a sync that failed — or that the host simply did not answer — throws on
 * every property read. Reading through this turns that into a missing value,
 * which callers can fall back from.
 */
export function loadedValue<T>(read: () => T): T | undefined {
  try {
    return read();
  } catch {
    return undefined;
  }
}

/**
 * Whether the host CONFIRMED this object exists.
 *
 * Three states collapse to two here, and the direction matters: "the host said
 * it is there" is true, while both "the host said it is gone" and "the host
 * never answered" are false. A caller that cannot tell must behave as if the
 * object is absent — every use of this guards a delete or a redraw, and doing
 * either to a shape we cannot see is how an edit destroys something.
 */
export function isLive(proxy: { isNullObject: boolean }): boolean {
  return loadedValue(() => proxy.isNullObject) === false;
}

/**
 * Did the host answer for a shape that arrived as a COLLECTION ITEM?
 *
 * `isLive` asks `isNullObject === false`, and that protocol exists only for a
 * `…OrNullObject` lookup. A shape handed back as an item of a loaded collection
 * has no null-object flag at all — it reads `undefined`, so `isLive` calls it
 * dead and refuses to write to a shape the host has already produced.
 *
 * Confirm it the way it actually arrived instead: `queueGroupMembers` asks for
 * `items/id`, so an id is the host's positive answer that this member exists.
 * A positive claim, not the absence of a negative one.
 */
export function isConfirmedMember(proxy: { id: string }): boolean {
  return Boolean(loadedValue(() => proxy.id));
}

/**
 * Did the host ANSWER for this object at all — either way?
 *
 * The other half of `isLive`, and the distinction it deliberately throws away.
 * `isLive` is right to fold silence in with absence wherever the next step is a
 * delete. But an UPDATE that folds them together stops being conservative and
 * starts being wrong: a chart the host merely would not resolve is still on the
 * slide, in front of the user, and doing nothing to it is a silent no-op on
 * something they can see.
 *
 * The archive says that is the common case rather than the exotic one. Forty-six
 * of the forty-seven recorded `explode a degraded picture` failures carry
 * `idRefusals > 0`, and the two that could report a verdict at all both said the
 * same thing: "the host would not work on the chart again, but it is STILL ON
 * THE SLIDE — nothing was lost".
 */
export function hostAnswered(proxy: { isNullObject: boolean }): boolean {
  return loadedValue(() => proxy.isNullObject) !== undefined;
}

/**
 * A tag's value, or undefined for "absent, or the host did not say".
 *
 * The `!tag.isNullObject && tag.value` pair was written out at eight call
 * sites, and every one of them read BOTH properties raw. Either read throws if
 * the load did not land, which is a `PropertyNotLoaded` in the middle of a scan
 * — the same shape of failure as the shape-collection one, one level down. A
 * tag nobody can read is a chart that is not re-editable; that is a fact to
 * degrade to, never to throw over.
 */
export function tagValue(tag: { isNullObject: boolean; value: string }): string | undefined {
  if (!isLive(tag)) return undefined;
  const v = loadedValue(() => tag.value);
  return typeof v === "string" && v ? v : undefined;
}
