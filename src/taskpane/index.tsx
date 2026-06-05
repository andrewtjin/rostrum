// Task-pane entry point. We mount React immediately and let `useRostrum` await
// `Office.onReady()` internally (it renders a "loading" view until the host is
// ready), which keeps the bootstrap in one tested place. Office.js itself is loaded
// by the <script> tag in taskpane.html before this bundle runs.

import { createRoot } from "react-dom/client";
import { App } from "./App";

const container = document.getElementById("root");
if (container) {
  createRoot(container).render(<App />);
}
