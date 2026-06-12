// The suite's roadmap, expressed as DATA — the Verbatim-rivalling tools Rostrum will grow.
// Each entry is registered as a first-class, capability-gated contribution so the RIBBON can
// advertise the full suite and the extension path is proven end-to-end (a planned tool gets its
// own ribbon group that opens a ComingSoon surface). When a tool is built, swap its `status` to
// "stable"/"preview", attach a `panel`/`dialog` in its feature module, fill in real `commands`,
// and run `npm run gen:manifest` — nothing in the shell/ribbon/dialog changes.
//
// The roadmap is currently EMPTY: Format & Condense, Flow, and Cite & Paste were removed to be
// re-added when their real implementations land. The pattern for re-adding them lives in git
// history (see the `plannedPane` / `plannedDialog` factories that used to live here) — a planned
// pane is one ribbon "Open" → its ComingSoon pane; a planned dialog is an `openWorkspaceDialog`
// action → a full-window ComingSoon workspace for space-heavy surfaces.
import { FeatureContribution } from "./types";

/** The planned suite, in ribbon group order after the shipped tools. Empty until the next tool lands. */
export const plannedContributions: FeatureContribution[] = [];
