// The full-window workspace — the suite's opt-in space-heavy surface, for tools that genuinely
// need room (Flow's speech-doc surface). It runs in a SEPARATE runtime from the task pane (an
// Office Dialog iframe), re-detects the host via the shared `useHost`, and renders exactly the ONE
// feature it was opened for (routed by `dialog.html#<id>`). Like the pane, it is feature-SCOPED:
// no launcher, no in-pane navigation — a feature is opened directly from its ribbon button.
import * as React from "react";
import { useState } from "react";
import { createRoot } from "react-dom/client";
import { FeatureSupport } from "../core/types";
import { registry } from "../features";
import { RostrumFeature } from "../features/types";
import { ComingSoon, featureIdFromHash, UnsupportedHost, useHost } from "../taskpane/host";

/** Ask the parent (task pane / ribbon) to close this dialog. messageParent only works in a dialog. */
function closeWorkspace(): void {
  if (typeof Office !== "undefined" && Office.context?.ui?.messageParent) {
    Office.context.ui.messageParent("close");
  }
}

/** Renders the active feature's DIALOG surface, or ComingSoon until that surface exists. */
function WorkspaceSurface(props: { feature: RostrumFeature; features: FeatureSupport }): React.ReactElement {
  const Dialog = props.feature.dialog;
  if (props.feature.status === "planned" || !Dialog) {
    return <ComingSoon feature={props.feature} />;
  }
  return <Dialog features={props.features} />;
}

function DialogWorkspace(): React.ReactElement {
  // SOFT gate (`requireSupport: false`): the dialog runs in a document-DETACHED Office runtime
  // whose `requirements` probe under-reports the desktop API sets, so the pane's hard capability
  // gate would falsely declare a capable host "lacks WordApiDesktop 1.2" here. The dialog is only
  // opened from an already-capable host and only shows ComingSoon placeholders (which need no
  // capability), so it must not refuse to render on a detection it can't trust. (When a real
  // capability-using dialog surface ships, it should receive caps from the OPENER — see
  // dialog/open.ts — rather than re-detecting them in this runtime.)
  const host = useHost({ requireSupport: false });
  // The routed feature is fixed when the dialog opens (one tool per workspace window).
  const [featureId] = useState<string | null>(() => featureIdFromHash(window.location.hash));

  if (host.phase === "loading") {
    return <div className="r-loading">Loading workspace…</div>;
  }
  if (host.phase === "unsupported" || !host.features) {
    return <UnsupportedHost message={host.unsupportedMessage ?? "This host is not supported."} />;
  }

  const features = host.features;
  const feature = featureId ? registry.get(featureId) ?? null : null;

  return (
    <div className="r-workspace">
      <header className="r-header">
        <div className="r-header__brand">
          <h1>Rostrum</h1>
          <span className="r-subtitle">{feature ? feature.title : "Workspace"}</span>
        </div>
        <div className="r-header__actions">
          <button className="r-btn" onClick={closeWorkspace}>
            Done
          </button>
        </div>
      </header>

      <main className="r-main" aria-label={feature ? feature.title : "Rostrum workspace"}>
        {feature ? (
          <WorkspaceSurface feature={feature} features={features} />
        ) : (
          <div className="r-soon">
            <p className="r-hint">Open a tool from the Rostrum tab on the ribbon.</p>
          </div>
        )}
      </main>
    </div>
  );
}

const container = document.getElementById("root");
if (container) {
  createRoot(container).render(<DialogWorkspace />);
}
