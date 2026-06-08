// Settings assembled for the React shells: the headless contribution plus its rendered pane surface.
// The registry (src/features/index.ts) consumes THIS; the ribbon runtime and the manifest generator
// consume the headless `settingsContribution` directly, so neither pulls in the React panel.
import { settingsContribution } from "./contribution";
import { SettingsPanel } from "./panel";
import { RostrumFeature } from "../types";

export const settingsFeature: RostrumFeature = {
  ...settingsContribution,
  panel: SettingsPanel,
};
