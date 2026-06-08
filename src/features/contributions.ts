// The HEADLESS suite — every feature's contribution (metadata + ribbon descriptor + commands),
// React-free. This is the single list the ribbon runtime (commands.ts) associates and the Node
// manifest generator (manifestGen.ts / tools/gen-manifest.ts) turns into manifest XML. The React
// shells use the parallel `registry` in index.ts (same features, with their rendered surfaces
// attached). Registration order here = ribbon group order, headline tool (Invisibility) first.
import { FeatureContribution } from "./types";
import { invisibilityContribution } from "./invisibility/contribution";
import { condenseContribution } from "./condense/contribution";
import { plannedContributions } from "./planned";
import { settingsContribution } from "./settings/contribution";

export const contributions: FeatureContribution[] = [
  // Suite-level settings group, FIRST (leftmost) on the Rostrum tab — global config sits ahead of the tools.
  settingsContribution,
  invisibilityContribution,
  condenseContribution,
  ...plannedContributions,
];
