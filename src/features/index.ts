// The assembled suite for the REACT shells. This is the ONE place features (with their rendered
// surfaces) are wired into the registry the pane + dialog read. Registration order = ribbon /
// listing order, with the headline convenience tool (Invisibility Mode) first, then the planned
// roadmap tools.
//
// To add a tool: build its contribution + (optional) React surfaces, register it here, add it to
// `contributions.ts`, and run `npm run gen:manifest`. Nothing in the shell/ribbon/dialog changes.
import { FeatureRegistry } from "./registry";
import { invisibilityFeature } from "./invisibility/feature";
import { plannedContributions } from "./planned";

/** The app-wide feature registry singleton (React-augmented features). */
export const registry = new FeatureRegistry();

// Feature #1: the headline time-differential tool. Real, stable, ships today.
registry.register(invisibilityFeature);

// The roadmap: planned contributions render a ComingSoon from metadata alone (no React surface).
for (const feature of plannedContributions) {
  registry.register(feature);
}

// Re-export the contribution surface so consumers import from one place.
export { FeatureRegistry } from "./registry";
export * from "./types";
