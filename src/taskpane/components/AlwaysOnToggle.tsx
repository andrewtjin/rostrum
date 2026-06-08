// Always-On toggle — the Settings switch for "load Rostrum on every document."
//
// This is a SUITE-level concern (not invisibility's), so it's a self-contained component with its own
// state over `core/alwaysOn` + the Office startup seam — deliberately NOT routed through any feature
// controller. It lives in the dedicated Settings pane (src/features/settings/panel.tsx), the suite's
// home for general app-wide settings; it shares no feature's engine.
//
// Cap-gate: when the host lacks the shared runtime (`SharedRuntime 1.1`) — e.g. a manifest that doesn't
// activate it, or desktop Word older than ~mid-2022 — `readAlwaysOn` reports `supported: false`. Rather
// than vanish (which would leave the dedicated Settings pane looking empty) the control then renders a
// short "not available here" note; the live checkbox shows only on a shared-runtime host. The three
// render-states are decided by the pure `alwaysOnView` in core/alwaysOn.ts.
import * as React from "react";
import { useCallback, useEffect, useRef, useState } from "react";
import { alwaysOnView, readAlwaysOn, setAlwaysOn, StartupBehaviorHost } from "../../core/alwaysOn";
import { createOfficeStartupHost, startupStorage } from "../../core/officeStartup";
import { StorageLike } from "../../core/settings";
import { logger } from "../../core/debug";

const log = logger("alwaysOn");

interface AlwaysOnUi {
  /** Null until the first read resolves (render nothing while loading — avoids a flash). */
  supported: boolean | null;
  on: boolean;
  busy: boolean;
  toggle: (next: boolean) => void;
}

/**
 * Own the toggle's state over the injected seam. The host + storage are injectable so the hook is
 * testable, but default to the live Office adapter for production.
 */
function useAlwaysOn(host: StartupBehaviorHost, storage: StorageLike): AlwaysOnUi {
  const [supported, setSupported] = useState<boolean | null>(null);
  const [on, setOn] = useState(true);
  const [busy, setBusy] = useState(false);
  // Guards every async state-write so a settle after the pane closes is a no-op (the pane can be
  // dismissed mid-toggle). One ref for the whole hook keeps the read effect + toggle consistent.
  const mounted = useRef(true);
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);

  useEffect(() => {
    readAlwaysOn(host, storage)
      .then((state) => {
        if (!mounted.current) return;
        setSupported(state.supported);
        setOn(state.on);
      })
      .catch((e) => {
        log.caught("reading always-on state failed (hiding toggle)", e);
        if (mounted.current) setSupported(false);
      });
  }, [host, storage]);

  const toggle = useCallback(
    (next: boolean) => {
      if (busy) return;
      setBusy(true);
      // Optimistic: reflect the user's choice immediately, reconcile from the real result.
      setOn(next);
      setAlwaysOn(host, storage, next)
        .then((state) => {
          if (!mounted.current) return;
          setSupported(state.supported);
          setOn(state.on);
        })
        .catch((e) => {
          // The persist+drive failed: revert the optimistic flip so the checkbox never shows a state
          // that wasn't saved. (setAlwaysOn persists the intent BEFORE driving Office, so a throw here
          // is rare — but the UI must never claim a setting stuck when it didn't.)
          log.caught("setting always-on failed", e);
          if (mounted.current) setOn(!next);
        })
        .finally(() => {
          if (mounted.current) setBusy(false);
        });
    },
    [busy, host, storage]
  );

  return { supported, on, busy, toggle };
}

/**
 * The rendered toggle. Renders nothing while loading or when the shared runtime is unavailable, so the
 * pane stays clean on hosts where Always-On can't work. Props default to the live Office seam; tests
 * inject a fake host/storage.
 */
export function AlwaysOnToggle(props?: {
  host?: StartupBehaviorHost;
  storage?: StorageLike;
}): React.ReactElement | null {
  // Build the live Office seam once per mount (cheap; just closes over the Office globals).
  const hostRef = useRef<StartupBehaviorHost>(props?.host ?? createOfficeStartupHost());
  const storageRef = useRef<StorageLike>(props?.storage ?? startupStorage());
  const ui = useAlwaysOn(hostRef.current, storageRef.current);

  const view = alwaysOnView(ui.supported);

  // Still resolving support → render nothing transient. This is the ONLY null case; the Settings pane's
  // own intro fills the frame, so there's no blank flash.
  if (view === "loading") return null;

  // No shared runtime here → don't vanish (that would leave the dedicated Settings pane empty); explain it.
  if (view === "unavailable") {
    return (
      <div className="r-section">
        <h2 className="r-section__title">Always-On</h2>
        <p className="r-note">
          Always-On isn&apos;t available on this version of Word. To open Rostrum in a document, use
          Home ▸ Add-ins.
        </p>
      </div>
    );
  }

  // Supported → the live switch. Rendered expanded inline (not a collapsed <details>) because on the
  // dedicated Settings pane this is the headline control, not an advanced/optional disclosure.
  return (
    <div className="r-section">
      <h2 className="r-section__title">Always-On {ui.on ? "(on)" : "(off)"}</h2>
      <p className="r-hint">
        When on, Rostrum loads on the ribbon automatically every time you open Word — no need to
        relaunch it from Home ▸ Add-ins per document. Turning it <strong>off</strong> keeps Rostrum
        installed but stops it loading automatically; relaunch from Home ▸ Add-ins anytime, or run
        the uninstaller to remove it fully.
      </p>
      <label className="r-live">
        <input
          type="checkbox"
          checked={ui.on}
          disabled={ui.busy}
          onChange={(e) => ui.toggle(e.target.checked)}
        />
        Load Rostrum on every document (recommended)
      </label>
    </div>
  );
}
