// Shared host-level shell pieces used by BOTH surfaces (the task pane and the workspace dialog):
// host readiness/capability detection (useHost), the unsupported-host fallback, the metadata-only
// ComingSoon, and the deep-link hash parser. Kept apart from the feature-scoped Shell so the pane
// and dialog import the same readiness logic and the same presentational fallbacks — no dup.
import * as React from "react";
import { useEffect, useState } from "react";
import { assertCanRun, detectFeatureSupport, UnsupportedHostError } from "../core/guards";
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

export function useHost(): HostState {
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
        const features = detectFeatureSupport(Office.context.requirements);
        try {
          assertCanRun(features); // throws UnsupportedHostError on web / old perpetual
        } catch (e) {
          if (e instanceof UnsupportedHostError) {
            setState({ phase: "unsupported", features, unsupportedMessage: e.message });
            return;
          }
          throw e;
        }
        setState({ phase: "ready", features, unsupportedMessage: null });
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
  }, []);

  return state;
}

/** Parse the deep-link feature id from a URL hash (`#invisibility` → "invisibility"). */
export function featureIdFromHash(hash: string): string | null {
  return hash && hash.length > 1 ? decodeURIComponent(hash.slice(1)) : null;
}

// ===========================================================================
// Presentational fallbacks shared by both surfaces.
// ===========================================================================

/** Rendered from feature metadata alone for `planned` tools (or any feature lacking a surface). */
export function ComingSoon(props: { feature: RostrumFeature }): React.ReactElement {
  const { feature } = props;
  return (
    <div className="r-soon">
      <div className="r-soon__glyph" aria-hidden="true">
        {feature.glyph}
      </div>
      <h2>{feature.title}</h2>
      <p>{feature.tagline}</p>
      <p className="r-hint">
        {feature.unavailableReason ?? "Coming soon."}
        {feature.primarySurface === "dialog" ? " Opens in a full-window workspace." : ""}
      </p>
      <span className="r-badge r-badge--planned">Planned</span>
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
