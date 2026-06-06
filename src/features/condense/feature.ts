// Condense & Shrink assembled for the React shells: the headless contribution plus its rendered pane
// surface. The registry (src/features/index.ts) consumes THIS; the ribbon runtime and the manifest
// generator consume the headless `condenseContribution` directly, so neither pulls in the React panel.
import { condenseContribution } from "./contribution";
import { CondensePanel } from "./panel";
import { RostrumFeature } from "../types";

export const condenseFeature: RostrumFeature = {
  ...condenseContribution,
  panel: CondensePanel,
};
