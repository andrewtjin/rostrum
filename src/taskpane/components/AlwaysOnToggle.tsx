// Always-On toggle — the Settings switch for "load Rostrum on every document."
//
// This is a SUITE-level concern (not invisibility's), so it's a self-contained component with its own
// state over `core/alwaysOn` + the Office startup seam — deliberately NOT routed through the
// invisibility controller. It's rendered inside the Invisibility "Settings" pane only because that is
// the settings surface the ribbon "Settings" button opens; it shares none of that feature's engine.
//
// Cap-gate: when the host lacks the shared runtime (`SharedRuntime 1.0`) — which is every build until
// the shared-runtime manifest is sideloaded — `readAlwaysOn` reports `supported: false` and this
// renders NOTHING. So the toggle simply doesn't appear where the lever can't work, and today's build
// looks identical. On a shared-runtime host it shows a checkbox wired to `setAlwaysOn`.
import * as React from "react";
import { useCallback, useEffect, useRef, useState } from "react";
import { readAlwaysOn, setAlwaysOn, StartupBehaviorHost } from "../../core/alwaysOn";
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
        .catch((e) => log.caught("setting always-on failed", e))
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

  // Loading or unsupported → render nothing (the cap-gate: no lever, no UI).
  if (ui.supported !== true) return null;

  return (
    <details className="r-section">
      <summary>Always on {ui.on ? "(on)" : "(OFF)"}</summary>
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
    </details>
  );
}
