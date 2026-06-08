// Settings — the REACT pane surface (the compact deep-linked task pane). The shell mounts this when
// `taskpane.html#settings` is open. It is the suite's home for GENERAL, app-wide settings; today that is
// the single Always-On switch, relocated here out of the Invisibility pane where it never belonged.
//
// Two deliberate UX rules encoded here:
//   * NEVER BLANK. `AlwaysOnToggle` renders nothing while it loads and on hosts without the shared
//     runtime (it self-cap-gates). A Settings pane whose only child is that widget would show an empty
//     page. So this panel ALWAYS renders its own intro line; the toggle handles its own absent state.
//   * TEST SEAM. `host` / `storage` are optional overrides forwarded straight to `AlwaysOnToggle` so the
//     wiring is exercisable with the fake Office seam; production omits them (live Office adapter).
import * as React from "react";
import { AlwaysOnToggle } from "../../taskpane/components/AlwaysOnToggle";
import { StartupBehaviorHost } from "../../core/alwaysOn";
import { StorageLike } from "../../core/settings";
import { FeaturePanelProps } from "../types";

export interface SettingsPanelProps extends FeaturePanelProps {
  /** Optional Office startup seam override (tests inject a fake; production uses the live adapter). */
  host?: StartupBehaviorHost;
  /** Optional storage override (tests inject a fake store; production uses localStorage). */
  storage?: StorageLike;
}

export function SettingsPanel(props: SettingsPanelProps): React.ReactElement {
  return (
    <div className="r-feature">
      {/* Always rendered — guarantees the pane is never blank, even while the toggle below loads or
          self-hides on a host without the shared runtime. */}
      <p className="r-hint">Control how Rostrum works across all your documents.</p>
      <AlwaysOnToggle host={props.host} storage={props.storage} />
    </div>
  );
}
