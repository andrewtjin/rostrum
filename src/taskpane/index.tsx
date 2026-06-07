// Shared-runtime entry point (manifest <FunctionFile> + long-lived <Runtime> page, since 0.3.0).
//
// This ONE page does three things in the single persistent runtime:
//   1. Mounts the React pane immediately. `useRostrum` awaits `Office.onReady()` internally and shows
//      a "loading" view until the host is ready, so the pane is correct whether shown now or later.
//   2. Wires the ribbon command handlers (`associateAll`) — the job the old ephemeral commands.js did.
//   3. Reconciles Always-On on first launch: registers `setStartupBehavior(load)` so the tab auto-loads
//      on every document (a no-op where the shared runtime / startup API is unavailable).
// Office.js itself is loaded by the <script> tag in taskpane.html before this bundle runs.

import { createRoot } from "react-dom/client";
import { App } from "./App";
import { associateAll } from "../commands/commands";
import { reconcileStartupBehavior } from "../core/alwaysOn";
import { createOfficeStartupHost, startupStorage } from "../core/officeStartup";
import { logger } from "../core/debug";

const log = logger("runtime");

const container = document.getElementById("root");
if (container) {
  createRoot(container).render(<App />);
}

// Ribbon wiring + Always-On reconciliation run once the host is ready. Guarded so importing this
// bundle in a non-host environment never touches `Office`.
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(() => {
    associateAll();
    // First manual launch self-registers auto-load; later launches are idempotent. Never throws.
    reconcileStartupBehavior(createOfficeStartupHost(), startupStorage()).catch((e) =>
      log.caught("always-on reconciliation failed (ignored)", e)
    );
  });
}
