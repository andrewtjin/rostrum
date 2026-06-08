// The Office-touching adapter for Always-On — the ONLY file that reads `Office.addin`.
//
// `core/alwaysOn.ts` holds the pure policy over an injected {@link StartupBehaviorHost} seam; this
// file is the production implementation of that seam. Isolating the Office calls here keeps the
// policy host-free + unit-tested, and means the capability probe (the make-or-break "is the shared
// runtime actually here?") lives in exactly one place.
//
// Everything is defensive: `Office.addin` and `Office.StartupBehavior` only exist on a shared-runtime
// host, and `Office.context.requirements` may be unusable in some runtimes — so `isSupported()` folds
// all of that into one boolean, and the getters/setters are only ever called after it returns true.

import { StartupBehavior, StartupBehaviorHost } from "./alwaysOn";

/* eslint-disable @typescript-eslint/no-explicit-any */

/** A throwaway in-memory storage for hosts without a real localStorage (privacy-locked WebView). */
function memoryStorage(): { getItem(k: string): string | null; setItem(k: string, v: string): void } {
  const m = new Map<string, string>();
  return { getItem: (k) => m.get(k) ?? null, setItem: (k, v) => void m.set(k, v) };
}

/** The device storage Always-On persists its intent to — real localStorage when present, else memory. */
export function startupStorage(): { getItem(k: string): string | null; setItem(k: string, v: string): void } {
  return typeof localStorage !== "undefined" ? localStorage : memoryStorage();
}

/**
 * Build the production {@link StartupBehaviorHost} from the live `Office` globals. Safe to call on any
 * host: when the shared-runtime startup API is absent, `isSupported()` returns false and the add-in's
 * Always-On feature stays fully dormant (no error, no UI).
 */
export function createOfficeStartupHost(): StartupBehaviorHost {
  return {
    isSupported(): boolean {
      try {
        const addin = (Office as any)?.addin;
        const reqs = (Office as any)?.context?.requirements;
        const sb = (Office as any)?.StartupBehavior;
        // SharedRuntime 1.1 is the FIRST published version of the set (there is no 1.0) and the one
        // that exposes the Office.addin.* startup-behavior APIs — so gate on 1.1, not a non-existent 1.0.
        const sharedRuntime = !!reqs && typeof reqs.isSetSupported === "function" && reqs.isSetSupported("SharedRuntime", "1.1");
        return (
          sharedRuntime &&
          !!addin &&
          typeof addin.getStartupBehavior === "function" &&
          typeof addin.setStartupBehavior === "function" &&
          // The StartupBehavior enum must exist too — get/setStartupBehavior below read sb.load/sb.none,
          // so an SDK that exposes the methods but not the enum would otherwise throw past the gate.
          !!sb &&
          typeof sb.load !== "undefined" &&
          typeof sb.none !== "undefined"
        );
      } catch {
        return false;
      }
    },
    async getStartupBehavior(): Promise<StartupBehavior> {
      const behavior = await (Office as any).addin.getStartupBehavior();
      // Office returns the StartupBehavior enum; treat anything that isn't `load` as `none`.
      return behavior === (Office as any).StartupBehavior?.load ? "load" : "none";
    },
    async setStartupBehavior(behavior: StartupBehavior): Promise<void> {
      const SB = (Office as any).StartupBehavior;
      await (Office as any).addin.setStartupBehavior(behavior === "load" ? SB.load : SB.none);
    },
  };
}
