// Task-pane root — a thin mount point for the feature-agnostic suite Shell. The Shell
// owns host readiness and renders the launcher → the selected feature's surface, so this
// file has no feature knowledge at all: Invisibility Mode and every future tool reach the
// screen via the registry (src/features), never via edits here.
import * as React from "react";
import { Shell } from "./Shell";

export function App(): React.ReactElement {
  return <Shell />;
}
