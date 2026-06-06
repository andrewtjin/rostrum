// The task-pane shell — feature-SCOPED, not a launcher. Each Rostrum tool opens its OWN pane
// directly from the ribbon (`taskpane.html#<id>`); this shell reads the hash and renders exactly
// that ONE feature's panel. There is deliberately NO launcher grid and NO in-pane navigation:
// the ribbon is the suite's home, the pane is a focused surface that opens only when a feature
// needs settings/progress, so it never sits open eating the debater's reading space.
//
// It imports the registry but NO concrete feature — Invisibility and every future tool reach the
// screen only by registering, so growing the suite never means editing this file.
import * as React from "react";
import { useEffect, useState } from "react";
import { FeatureSupport } from "../core/types";
import { registry } from "../features";
import { RostrumFeature } from "../features/types";
import { ComingSoon, featureIdFromHash, UnsupportedHost, useHost } from "./host";
import { DiagnosticsPanel } from "./components/DiagnosticsPanel";

/** Renders the active feature's pane surface, or a ComingSoon when it has no panel yet. */
function FeatureHost(props: { feature: RostrumFeature; features: FeatureSupport }): React.ReactElement {
  const Panel = props.feature.panel;
  if (props.feature.status === "planned" || !Panel) {
    return <ComingSoon feature={props.feature} />;
  }
  return <Panel features={props.features} />;
}

/** Shown only when the pane is opened without a feature route (an edge fallback, not a hub). */
function NoFeature(): React.ReactElement {
  return (
    <div className="r-empty">
      <p className="r-hint">Pick a tool from the <b>Rostrum</b> tab on the ribbon.</p>
    </div>
  );
}

export function Shell(): React.ReactElement {
  const host = useHost();
  // The active feature is the hash route — set by which ribbon pane-button opened this pane.
  const [featureId, setFeatureId] = useState<string | null>(() => featureIdFromHash(window.location.hash));
  useEffect(() => {
    const onHash = (): void => setFeatureId(featureIdFromHash(window.location.hash));
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, []);

  if (host.phase === "loading") {
    return <div className="r-loading">Loading Rostrum…</div>;
  }
  if (host.phase === "unsupported" || !host.features) {
    return <UnsupportedHost message={host.unsupportedMessage ?? "This host is not supported."} />;
  }

  const features = host.features;
  const feature = featureId ? registry.get(featureId) ?? null : null;

  return (
    <div className="r-app">
      <header className="r-header">
        <div className="r-header__brand">
          <h1>Rostrum</h1>
          <span className="r-subtitle">{feature ? feature.title : "Debating suite"}</span>
        </div>
      </header>

      <main className="r-main" aria-label={feature ? feature.title : "Rostrum"}>
        {feature ? <FeatureHost feature={feature} features={features} /> : <NoFeature />}
      </main>

      {/* Diagnostics is cross-cutting suite chrome (host-capability matrix + live tracer),
          not a feature — always available, collapsed by default. */}
      <DiagnosticsPanel features={features} />
    </div>
  );
}
