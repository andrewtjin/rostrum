// The assembled suite for the REACT shells. This is the ONE place features (with their rendered
// surfaces) are wired into the registry the pane + dialog read. Registration order = ribbon /
// listing order, with the headline convenience tool (Invisibility Mode) first, then the planned
// roadmap tools.
//
// To add a tool: build its contribution + (optional) React surfaces, register it here, add it to
// `contributions.ts`, and run `npm run gen:manifest`. Nothing in the shell/ribbon/dialog changes.
import { FeatureRegistry } from "./registry";
import { invisibilityFeature } from "./invisibility/feature";
import { condenseFeature } from "./condense/feature";
import { plannedContributions } from "./planned";
import { settingsFeature } from "./settings/feature";

/** The app-wide feature registry singleton (React-augmented features). */
export const registry = new FeatureRegistry();

// Feature #1: the headline time-differential tool. Real, stable, ships today.
registry.register(invisibilityFeature);

// Feature #2: Condense & Shrink — the lossless answer to Verbatim's Shrink + Condense.
registry.register(condenseFeature);

// The roadmap: planned contributions render a ComingSoon from metadata alone (no React surface).
for (const feature of plannedContributions) {
  registry.register(feature);
}

// Suite-level settings, registered LAST so its group sits rightmost on the Rostrum tab. Must mirror the
// position of `settingsContribution` in contributions.ts (a parity test guards that the two lists agree).
registry.register(settingsFeature);

// Re-export the contribution surface so consumers import from one place.
export { FeatureRegistry } from "./registry";
export * from "./types";
