// Invisibility Mode assembled for the React shells: the headless contribution plus its rendered
// pane surface. The registry (src/features/index.ts) consumes THIS; the ribbon runtime and the
// manifest generator consume the headless `invisibilityContribution` directly, so neither pulls
// in the React panel.
import { invisibilityContribution } from "./contribution";
import { InvisibilityPanel } from "./panel";
import { RostrumFeature } from "../types";

export const invisibilityFeature: RostrumFeature = {
  ...invisibilityContribution,
  panel: InvisibilityPanel,
};
