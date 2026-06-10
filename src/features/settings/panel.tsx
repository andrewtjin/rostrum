// Settings — the REACT pane surface (the compact deep-linked task pane). The shell mounts this when
// `taskpane.html#settings` is open. It is the suite's home for GENERAL, app-wide settings.
//
// Today it is purely INFORMATIONAL. "Load Rostrum on every document" used to be an in-pane Always-On
// switch (Office.addin.setStartupBehavior); that was retired 2026-06-08 because the platform can't do
// it (setStartupBehavior is per-document, never reaching a brand-new doc) and because the
// Trusted-Catalog sideload already loads Rostrum on every document for free. An add-in also cannot
// un-sideload itself (sandbox: no registry/filesystem), so the real OFF lever is external — the
// Trust Center / a catalog-registry toggle — never an in-pane button. This pane therefore STATES how
// Rostrum is installed and POINTS at the external OFF steps, rather than pretending to toggle it.
import * as React from "react";
import { PRODUCT_VERSION } from "../version";

export function SettingsPanel(): React.ReactElement {
  return (
    <div className="r-feature">
      <p className="r-hint">General Rostrum settings — these apply across all your documents.</p>

      {/* Status (static). The pane has no API to read the catalog install state, so this is honest
          informational copy, not a live readout. */}
      <div className="r-section">
        <h2 className="r-section__title">On for every document</h2>
        <div className="r-chips">
          <span className="r-chip r-chip--on">Catalog install</span>
        </div>
        <p className="r-hint">
          Rostrum loads on the ribbon of every Word document on this PC because it&apos;s installed via a
          trusted add-in catalog. There&apos;s nothing to switch on here — opening one document with
          Rostrum is enough; it stays on the ribbon for the rest.
        </p>
      </div>

      {/* OFF is external by necessity — an add-in can't un-sideload itself, so this links to the steps
          rather than offering a switch that couldn't actually work. */}
      <div className="r-section">
        <h2 className="r-section__title">Turn Rostrum off for every document</h2>
        <p className="r-hint">
          Removing Rostrum from every document is done in Word&apos;s Trust Center, not from this pane:
        </p>
        <ol className="r-list">
          <li>File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Trusted Add-in Catalogs.</li>
          <li>Untick <strong>Show in Menu</strong> for the Rostrum catalog.</li>
          <li>
            Tick <strong>Next time Office starts, clear all previously-started web add-ins</strong>, then
            restart Word.
          </li>
        </ol>
        <p className="r-note">
          The clear-on-restart step affects <em>every</em> sideloaded web add-in, not just Rostrum. To
          turn Rostrum back on, re-tick <strong>Show in Menu</strong> and restart.
        </p>
      </div>

      {/* The one place a student can read their running version — support asks and the site's
          Updates section both start from "what version are you on". */}
      <p className="r-note">Rostrum v{PRODUCT_VERSION}</p>
    </div>
  );
}
