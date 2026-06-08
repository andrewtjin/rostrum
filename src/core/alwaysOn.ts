// Always-On reconciliation — the host-free brain behind "load Rostrum on every document."
//
// The user's reported "huge bug" was that a sideloaded add-in shows under Home ▸ Add-ins but its
// custom tab does NOT auto-appear on every new document — they had to relaunch it per document. The
// Office-native cure is a SHARED RUNTIME driven by `Office.addin.setStartupBehavior(load | none)`.
// This module owns the *policy* around that lever — pure logic over an injected seam — so the
// decision ("on first launch, register for auto-load; honor an explicit opt-out; no-op where the
// host can't") is unit-tested in Node with a fake host, exactly like the rest of `core/`.
//
// The Office-touching adapter that satisfies {@link StartupBehaviorHost} lives next door in
// `core/officeStartup.ts` (the only file that reads `Office.addin`), keeping this one host-free.

import { StorageLike, loadAlwaysOn, saveAlwaysOn } from "./settings";

/** The two startup states Office exposes, normalized to a plain string so this module needs no Office. */
export type StartupBehavior = "load" | "none";

/**
 * The slice of `Office.addin` we depend on, as an injectable seam. `isSupported()` folds together the
 * two facts that gate everything: the host advertises `SharedRuntime 1.1` (the set's first real version)
 * AND `Office.addin` actually exposes the startup-behavior API. When false, every operation below is a
 * graceful no-op — a build whose manifest doesn't activate the shared runtime reports false, so this
 * whole feature stays dormant. Production wiring: {@link import("./officeStartup").createOfficeStartupHost}.
 */
export interface StartupBehaviorHost {
  /** True only when the shared-runtime startup API is genuinely callable on this host. */
  isSupported(): boolean;
  /** Office's current source-of-truth startup behavior. Only called when `isSupported()`. */
  getStartupBehavior(): Promise<StartupBehavior>;
  /** Set Office's startup behavior. Only called when `isSupported()`. */
  setStartupBehavior(behavior: StartupBehavior): Promise<void>;
}

/** What a reconcile / set / read reports back to its caller (UI or startup bootstrap). */
export interface AlwaysOnState {
  /** Whether the shared-runtime startup API is available — drives the toggle's cap-gate. */
  supported: boolean;
  /** The effective on/off state (Office truth when supported; the persisted intent otherwise). */
  on: boolean;
}

/** Map the persisted boolean intent to the Office behavior token. Pure. */
function intentToBehavior(on: boolean): StartupBehavior {
  return on ? "load" : "none";
}

/**
 * The persisted per-device intent as an Office behavior — the seed the first-launch reconciliation
 * drives Office toward. Default ON (decision #3): absent storage ⇒ `"load"`.
 */
export function desiredBehavior(storage: StorageLike): StartupBehavior {
  return intentToBehavior(loadAlwaysOn(storage, true));
}

/**
 * First-launch self-register: make Office's real startup behavior match the persisted intent. This is
 * the "the first manual launch is the last one" logic — the first time the user opens Rostrum (which
 * they already do via Home ▸ Add-ins), this registers `load` so every later document has the tab with
 * no action. Idempotent and self-healing (only writes when Office disagrees), so it's safe to run on
 * every shared-runtime startup. A no-op — and never throws — where the host can't do startup behavior.
 */
export async function reconcileStartupBehavior(
  host: StartupBehaviorHost,
  storage: StorageLike
): Promise<AlwaysOnState> {
  const desired = desiredBehavior(storage);
  if (!host.isSupported()) {
    // Can't drive Office here; report the persisted intent so callers still render a sensible state.
    return { supported: false, on: desired === "load" };
  }
  try {
    const current = await host.getStartupBehavior();
    if (current !== desired) {
      await host.setStartupBehavior(desired);
    }
    return { supported: true, on: desired === "load" };
  } catch {
    // A live host that still failed the probe/set: degrade to "supported but report intent" rather
    // than throwing out of the startup path (a failed auto-register must never break the add-in).
    return { supported: true, on: desired === "load" };
  }
}

/**
 * Read the current always-on state for the Settings toggle. Prefers Office's `getStartupBehavior()`
 * (the source of truth) when supported, falling back to the persisted intent otherwise — so the
 * toggle reflects reality on a shared-runtime host and a sensible default elsewhere.
 */
export async function readAlwaysOn(host: StartupBehaviorHost, storage: StorageLike): Promise<AlwaysOnState> {
  if (!host.isSupported()) {
    return { supported: false, on: loadAlwaysOn(storage, true) };
  }
  try {
    const current = await host.getStartupBehavior();
    return { supported: true, on: current === "load" };
  } catch {
    return { supported: true, on: loadAlwaysOn(storage, true) };
  }
}

/**
 * Apply a toggle change: persist the new intent AND drive Office to match (when supported). Persisting
 * first means an opt-out survives even if the Office call fails. Returns the resulting state.
 */
export async function setAlwaysOn(
  host: StartupBehaviorHost,
  storage: StorageLike,
  on: boolean
): Promise<AlwaysOnState> {
  saveAlwaysOn(storage, on);
  if (!host.isSupported()) {
    return { supported: false, on };
  }
  try {
    await host.setStartupBehavior(intentToBehavior(on));
  } catch {
    // Intent is persisted; Office will be reconciled on the next launch. Don't surface a hard failure.
  }
  return { supported: true, on };
}
