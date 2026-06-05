// The task pane's root component: a thin composition of the (tested) hook's state
// and the presentational panels. The only "logic" here is choosing which top-level
// view to render for the current phase.

import * as React from "react";
import { useRostrum } from "./useRostrum";
import {
  ApplyStyles,
  Buttons,
  KeepColorPicker,
  WholeBodyModeToggle,
  StatusBar,
  TrackChangesPrompt,
  UnsupportedHost
} from "./components/Panels";
import { DiagnosticsPanel } from "./components/DiagnosticsPanel";

export function App(): React.ReactElement {
  const ui = useRostrum();

  if (ui.phase === "loading") {
    return <div className="r-loading">Loading Rostrum…</div>;
  }
  if (ui.phase === "unsupported") {
    return <UnsupportedHost message={ui.unsupportedMessage ?? "This host is not supported."} />;
  }

  return (
    <div className="r-app">
      <header className="r-header">
        <h1>Rostrum</h1>
        <span className="r-subtitle">Invisibility Mode</span>
      </header>

      <Buttons
        status={ui.status}
        busy={ui.busy}
        onHide={ui.hide}
        onReHide={ui.reHide}
        onShowAll={ui.showAll}
        onCancel={ui.cancel}
      />

      <StatusBar status={ui.status} busy={ui.busy} progress={ui.progress} banner={ui.banner} />

      {/* Live mode (keep the paragraph I'm typing in visible) is HIDDEN for now. The engine stays
          in src/liveMode.ts and is still wired through useRostrum (ui.liveOn / ui.toggleLive); to
          re-enable, restore the control here:
            {ui.features?.canHide && (
              <label className="r-live">
                <input type="checkbox" checked={ui.liveOn} onChange={ui.toggleLive} />
                Live mode — keep the paragraph I'm typing in visible
              </label>
            )} */}

      <KeepColorPicker keepColors={ui.status.keepColors} busy={ui.busy} onChange={ui.setKeepColors} />

      <WholeBodyModeToggle status={ui.status} busy={ui.busy} onChange={ui.setPureWholeBody} />

      {ui.features && <ApplyStyles features={ui.features} busy={ui.busy} onApply={ui.applyStyles} />}

      <DiagnosticsPanel features={ui.features} status={ui.status} />

      {ui.trackChangesMode && (
        <TrackChangesPrompt
          mode={ui.trackChangesMode}
          onConfirm={ui.confirmTrackChanges}
          onDismiss={ui.dismissTrackChanges}
        />
      )}
    </div>
  );
}
