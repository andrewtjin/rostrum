// Shared host-level shell pieces used by BOTH surfaces (the task pane and the workspace dialog):
// host readiness/capability detection (useHost), the unsupported-host fallback, the metadata-only
// ComingSoon, and the deep-link hash parser. Kept apart from the feature-scoped Shell so the pane
// and dialog import the same readiness logic and the same presentational fallbacks — no dup.
import * as React from "react";
import { useEffect, useState } from "react";
import { assertCanRun, detectFeatureSupport, RequirementsLike, UnsupportedHostError } from "../core/guards";
import { FeatureSupport } from "../core/types";
import { logger } from "../core/debug";
import { RostrumFeature } from "../features/types";

const log = logger("shell");

// ===========================================================================
// Host readiness — a SUITE concern (not invisibility's). No feature surface mounts until the
// host is both ready and supported, so feature panels can assume a live, capable host.
// ===========================================================================
export type HostPhase = "loading" | "unsupported" | "ready";

export interface HostState {
  phase: HostPhase;
  features: FeatureSupport | null;
  unsupportedMessage: string | null;
}

/** Options for {@link useHost} / {@link resolveHostState}. */
export interface UseHostOptions {
  /**
   * Hard-gate on host capability (default `true`). This is the load-bearing flag behind the
   * dialog bug fix.
   *
   * The TASK PANE and ribbon run in the document-bound runtime, where `Office.context.requirements`
   * reports the host's real API sets — so they gate (`true`): an honestly-unsupported host (Word on
   * the web / old perpetual) gets the UnsupportedHost screen with the native-reverse steps.
   *
   * The DIALOG window runs in a SEPARATE Office runtime that is detached from the document, where
   * the requirements probe is unreliable and routinely UNDER-reports the desktop sets
   * (`WordApiDesktop *`). Re-running the hard gate there falsely declared a perfectly capable
   * desktop host "lacks WordApiDesktop 1.2". The dialog is only ever opened from an already-capable
   * host and (today) only shows ComingSoon placeholders that need NO host capability, so it opts
   * OUT (`false`): capability is a per-FEATURE concern (each feature's `isAvailable` / ComingSoon),
   * not a reason to refuse to render the suite shell at all.
   */
  requireSupport?: boolean;
}

/** Host capabilities when the runtime exposes no usable requirements probe (e.g. the dialog window). */
const NO_CAPABILITIES: FeatureSupport = {
  canHide: false,
  canCustomXml: false,
  canChangeTracking: false,
  canStyleBorders: false,
  canStyleFormat: false,
  canGetStyles: false,
};

/**
 * Pure readiness resolution: detect capabilities and decide the phase with NO React/Office wiring,
 * so the gate logic — and the dialog's opt-out — is unit-testable against a fake requirements probe
 * (`__tests__/hostState.test.ts`). In soft mode (`requireSupport: false`) a missing/throwing probe
 * degrades to "no capabilities, still ready" rather than a false unsupported / failed-to-start
 * screen — exactly the dialog-runtime case.
 */
export function resolveHostState(req: RequirementsLike | undefined, opts?: UseHostOptions): HostState {
  const requireSupport = opts?.requireSupport ?? true;
  let features: FeatureSupport;
  try {
    features = detectFeatureSupport(req as RequirementsLike);
  } catch (e) {
    // The dialog runtime can lack a usable requirements object entirely. Soft mode doesn't gate on
    // capability, so come up ready with nothing detected instead of failing to start; hard mode
    // (pane/ribbon) treats a missing probe as a genuine startup problem and rethrows.
    if (!requireSupport) return { phase: "ready", features: NO_CAPABILITIES, unsupportedMessage: null };
    throw e;
  }
  if (requireSupport) {
    try {
      assertCanRun(features); // throws UnsupportedHostError on web / old perpetual
    } catch (e) {
      if (e instanceof UnsupportedHostError) {
        return { phase: "unsupported", features, unsupportedMessage: e.message };
      }
      throw e;
    }
  }
  return { phase: "ready", features, unsupportedMessage: null };
}

export function useHost(opts?: UseHostOptions): HostState {
  const requireSupport = opts?.requireSupport ?? true;
  const [state, setState] = useState<HostState>({
    phase: "loading",
    features: null,
    unsupportedMessage: null,
  });

  useEffect(() => {
    let cancelled = false;
    Office.onReady()
      .then(() => {
        if (cancelled) return;
        setState(resolveHostState(Office.context.requirements, { requireSupport }));
      })
      .catch((e) => {
        log.caught("host bootstrap failed", e);
        if (cancelled) return;
        setState({
          phase: "unsupported",
          features: null,
          unsupportedMessage: `Rostrum failed to start: ${String((e as Error)?.message ?? e)}`,
        });
      });
    return () => {
      cancelled = true;
    };
  }, [requireSupport]);

  return state;
}

/** Parse the deep-link feature id from a URL hash (`#invisibility` → "invisibility"). */
export function featureIdFromHash(hash: string): string | null {
  return hash && hash.length > 1 ? decodeURIComponent(hash.slice(1)) : null;
}

// ===========================================================================
// Presentational fallbacks shared by both surfaces.
// ===========================================================================

/**
 * Rendered from feature metadata alone for `planned` tools (or any feature lacking a surface).
 * This IS the surface a debater sees for an unbuilt tool, so it does real work: it advertises what
 * the tool will do (`highlights`), where it will live (pane vs full-window workspace), and its
 * maturity — turning "coming soon" from a dead end into a preview of the suite. No host capability
 * is consulted here: a planned tool needs none, and (crucially) the dialog runtime can't detect
 * them reliably — so this surface must never gate on them (that was the bug).
 */
export function ComingSoon(props: { feature: RostrumFeature }): React.ReactElement {
  const { feature } = props;
  // Label off the maturity so a future stable-but-surfaceless feature isn't mislabeled "Planned".
  const badge =
    feature.status === "stable" ? "Available" : feature.status === "preview" ? "Preview" : "Planned";
  const surfaceNote =
    feature.primarySurface === "dialog"
      ? "Opens in a full-window workspace."
      : "Opens in the Rostrum pane.";
  return (
    <div className="r-soon">
      <div className="r-soon__glyph" aria-hidden="true">
        {feature.glyph}
      </div>
      <h2>{feature.title}</h2>
      <p className="r-soon__tagline">{feature.tagline}</p>
      {feature.highlights && feature.highlights.length > 0 ? (
        <>
          <p className="r-soon__label">What it’ll do</p>
          <ul className="r-soon__list">
            {feature.highlights.map((h) => (
              <li key={h}>{h}</li>
            ))}
          </ul>
        </>
      ) : null}
      <p className="r-hint">
        {feature.unavailableReason ?? "Coming soon."} {surfaceNote}
      </p>
      <span className="r-badge r-badge--planned">{badge}</span>
    </div>
  );
}

/** Unsupported-host panel (web / old perpetual) with the native, add-in-free reverse steps. */
export function UnsupportedHost(props: { message: string }): React.ReactElement {
  return (
    <div className="r-unsupported">
      <h2>Rostrum needs desktop Word</h2>
      <p>{props.message}</p>
      <h3>Reveal hidden text without the add-in</h3>
      <ol>
        <li>Select the affected text (or the whole document with Ctrl+A).</li>
        <li>
          Open Home ▸ Font dialog (Ctrl+D) and clear the <b>Hidden</b> checkbox, or toggle Home ▸ ¶
          (Show/Hide) to view hidden text.
        </li>
      </ol>
    </div>
  );
}
