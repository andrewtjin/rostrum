// Safety gates and host capability detection (decisions #14, #18).
//
// Two independent concerns live here, both pure:
//   * Track-Changes hard gate — Hide must not run while TC is on, or every hidden
//     run becomes a tracked deletion and a partial Undo can strand the document.
//   * Feature detection — split the manifest *floor* (hard requirements) from
//     runtime feature flags, so the add-in fails loudly on an unsupported host
//     rather than throwing deep in the engine.

import { FeatureSupport, TrackChangesMode } from "./types";

/** Thrown by the engine when Hide is attempted while Track Changes is on. */
export class TrackChangesActiveError extends Error {
  constructor(public readonly mode: TrackChangesMode) {
    super(
      `Track Changes is "${mode}". Rostrum hides text only with Track Changes off, ` +
        `so a partial Undo cannot strand the document. Turn it off and retry, or let ` +
        `Rostrum toggle it off for this operation.`
    );
    this.name = "TrackChangesActiveError";
    Object.setPrototypeOf(this, TrackChangesActiveError.prototype);
  }
}

/** Thrown when the host lacks a hard-required capability (e.g. Word on the web). */
export class UnsupportedHostError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UnsupportedHostError";
    Object.setPrototypeOf(this, UnsupportedHostError.prototype);
  }
}

/** Gate: throw unless Track Changes is fully off. */
export function assertTrackChangesOff(mode: TrackChangesMode): void {
  if (mode !== "Off") throw new TrackChangesActiveError(mode);
}

/** The minimal port surface the Track-Changes gate needs (satisfied by WordPort AND RangeScopedPort). */
export interface TrackChangesPort {
  getChangeTrackingMode(): Promise<TrackChangesMode>;
  setChangeTrackingMode(mode: TrackChangesMode): Promise<void>;
}

/**
 * The Track-Changes POLICY over an ALREADY-READ mode (decision #14), with the throw/toggle/finally-
 * restore behavior that both gate entry points share. Kept as the one copy of that policy (DRY): the
 * standard `withTrackChangesGate` issues its own TC read then delegates here; the Hide engine's
 * read-fused path (Loop 002 B2 / 002-S4) primes the TC mode in `readParagraphs`' FIRST sync and calls
 * this directly, so a clean Hide spends NO separate Word.run on the TC read.
 *
 * Semantics (identical to the legacy gate): with TC off, run directly. With TC on: throw
 * `TrackChangesActiveError` unless the caller opted into auto-toggle, in which case turn TC off, run,
 * and restore the prior mode in `finally` — so a partial Undo can never strand the document with
 * half-tracked revisions. Returns the body's result plus whether it toggled.
 *
 * NOTE the ordering contract the Hide path depends on: when TC is on and auto-toggle is off, this
 * throws BEFORE invoking `body`, so a caller that has already read the document (to prime the mode in
 * the same sync) classifies/writes NOTHING — the abort happens before any parse or host write.
 */
export async function withPrefetchedTrackChangesGate<T>(
  port: TrackChangesPort,
  autoToggle: boolean,
  mode: TrackChangesMode,
  body: () => Promise<T>
): Promise<{ result: T; toggled: boolean }> {
  if (mode === "Off") {
    return { result: await body(), toggled: false };
  }
  if (!autoToggle) {
    assertTrackChangesOff(mode); // throws TrackChangesActiveError (before body runs)
  }
  await port.setChangeTrackingMode("Off");
  try {
    return { result: await body(), toggled: true };
  } finally {
    await port.setChangeTrackingMode(mode);
  }
}

/**
 * Run a mutation under the Track-Changes gate (decision #14), shared by the whole-body Hide engine AND
 * the range-scoped Condense & Shrink controller so the policy lives in ONE place. Reads the current TC
 * mode in its own Word.run, then applies the shared throw/toggle/finally-restore policy
 * (`withPrefetchedTrackChangesGate`). Returns the body's result plus whether it toggled.
 *
 * The Hide path no longer uses THIS entry point (it primes TC inside its read sync and calls the
 * prefetched variant); Condense & Shrink still do, so their behavior is byte-for-byte unchanged.
 */
export async function withTrackChangesGate<T>(
  port: TrackChangesPort,
  autoToggle: boolean,
  body: () => Promise<T>
): Promise<{ result: T; toggled: boolean }> {
  const mode = await port.getChangeTrackingMode();
  return withPrefetchedTrackChangesGate(port, autoToggle, mode, body);
}

/**
 * The capability probe we need — `Office.context.requirements` satisfies this,
 * and tests pass a fake. Keeping it injectable means feature detection is unit
 * tested without an Office host.
 */
export interface RequirementsLike {
  isSetSupported(name: string, version?: string): boolean;
}

/** Map host requirement sets to Rostrum's capability flags (decision #18). */
export function detectFeatureSupport(req: RequirementsLike): FeatureSupport {
  return {
    canHide: req.isSetSupported("WordApiDesktop", "1.2"),
    canCustomXml: req.isSetSupported("WordApi", "1.4"),
    canChangeTracking: req.isSetSupported("WordApi", "1.4"),
    canStyleBorders: req.isSetSupported("WordApiDesktop", "1.1"),
    canStyleFormat: req.isSetSupported("WordApi", "1.5"),
    // `Document.getStyles()` — the METHOD officeStyles.ts actually calls — is WordApi 1.5,
    // NOT WordApiDesktop 1.4. The desktop set gates the *synchronous* `Document.styles`
    // property, which Rostrum doesn't use. Gating on the desktop set falsely disabled
    // Apply Styles on real current-Word builds that report WordApi 1.5 (canStyleFormat:true)
    // but lack WordApiDesktop 1.4 (canGetStyles was wrongly false). See LESSONS #42.
    canGetStyles: req.isSetSupported("WordApi", "1.5")
  };
}

/**
 * Assert the host can run Rostrum at all. The two hard requirements are hiding
 * text (WordApiDesktop 1.2) and the manifest store (WordApi 1.4); everything else
 * degrades gracefully.
 */
export function assertCanRun(support: FeatureSupport): void {
  if (!support.canHide) {
    throw new UnsupportedHostError(
      "Rostrum runs on Word for Windows or Mac (desktop). This host lacks " +
        "WordApiDesktop 1.2, which Rostrum needs to hide text. To collapse a doc " +
        "without the add-in, hidden text is still reversible from Word's Font dialog."
    );
  }
  if (!support.canCustomXml) {
    throw new UnsupportedHostError(
      "This host lacks WordApi 1.4 (custom XML parts), which Rostrum needs to " +
        "remember the ON-state and keep-colors per document."
    );
  }
}

/**
 * One-time warning text for OneDrive/SharePoint co-authoring sessions (decision
 * #14). There is no reliable Office.js signal for live co-authoring, so this is
 * surfaced advisorily by the task pane rather than enforced.
 */
export const COAUTHORING_WARNING =
  "If others are co-authoring this document right now, hide/show changes may " +
  "merge unpredictably. Rostrum's Show All is convergent, but it's safest to run " +
  "invisibility when you're the only editor.";
