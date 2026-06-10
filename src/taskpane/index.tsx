// Task-pane entry point — the manifest's <FunctionFile> + task-pane page (one bundle).
//
// This ONE page does two things, in whichever runtime Office loads it:
//   1. Mounts the React pane immediately. `useRostrum` awaits `Office.onReady()` internally and shows
//      a "loading" view until the host is ready, so the pane is correct whether shown now or later.
//   2. Wires the ribbon command handlers (`associateAll`) — so a ribbon ExecuteFunction button finds
//      its function in the ephemeral function-file runtime Office spins up for it.
// Office.js itself is loaded by the <script> tag in taskpane.html before this bundle runs.
//
// (Pre-0.3.0 the ribbon used a separate ephemeral commands.html; the 0.3.x always-on spike made this
// page a long-lived shared runtime. Always-On is retired — see the removal plan — so this is back to a
// plain TaskPaneApp: taskpane.html is both the on-demand pane and the function file, no shared runtime.)

import { createRoot } from "react-dom/client";
import { App } from "./App";
import { associateAll } from "../commands/commands";
// Pane stylesheet — bundled so it ships content-hashed (the WebView caches add-in assets
// aggressively; a stable CSS URL could go stale against new JS after an update).
import "./taskpane.css";

const container = document.getElementById("root");
if (container) {
  createRoot(container).render(<App />);
}

// Ribbon wiring runs once the host is ready. Guarded so importing this bundle in a non-host
// environment never touches `Office`.
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(() => {
    associateAll();
  });
}
