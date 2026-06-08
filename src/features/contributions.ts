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
  invisibilityContribution,
  condenseContribution,
  ...plannedContributions,
  // Suite-level settings group, last (rightmost) on the Rostrum tab — a config destination, not a tool.
  settingsContribution,
];
