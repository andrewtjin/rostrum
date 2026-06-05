// Presentational pieces of the task pane. These are pure-render components driven
// entirely by props from `useRostrum` — no Office.js, no engine logic — so the
// visual layer stays trivial to reason about. The interesting orchestration lives
// in the (tested) controller/hook.

import * as React from "react";
import { FeatureSupport } from "../../core/types";
import { HIGHLIGHT_COLORS } from "../../core/styles";
import { COAUTHORING_WARNING } from "../../core/guards";
import { REFLOW_WARNING } from "../../core/officeStyles";
import { Banner } from "../useRostrum";
import { ProgressInfo } from "../../core/officeWordPort";
import { ControllerStatus } from "../controller";

// ---------------------------------------------------------------------------
// Primary actions: Hide / Re-hide / Show All (+ Cancel while busy)
// ---------------------------------------------------------------------------
export function Buttons(props: {
  status: ControllerStatus;
  busy: boolean;
  onHide: () => void;
  onReHide: () => void;
  onShowAll: () => void;
  onCancel: () => void;
}): React.ReactElement {
  const { status, busy } = props;
  return (
    <div className="r-buttons">
      <button className="r-btn r-btn--primary" disabled={busy} onClick={props.onHide}>
        {status.armed ? "Hide (re-arm)" : "Hide"}
      </button>
      <button className="r-btn" disabled={busy || !status.armed} onClick={props.onReHide}>
        Re-hide
      </button>
      <button className="r-btn" disabled={busy} onClick={props.onShowAll}>
        Show All
      </button>
      {busy && (
        <button className="r-btn r-btn--ghost" onClick={props.onCancel}>
          Cancel
        </button>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Status bar: armed/live indicators, progress, the action banner, co-auth note
// ---------------------------------------------------------------------------
export function StatusBar(props: {
  status: ControllerStatus;
  busy: boolean;
  progress: ProgressInfo | null;
  banner: Banner | null;
}): React.ReactElement {
  const { status, progress, banner } = props;
  return (
    <div className="r-status">
      <div className="r-chips">
        <span className={`r-chip ${status.armed ? "r-chip--on" : ""}`}>
          {status.armed ? "Invisibility ON" : "Invisibility off"}
        </span>
        {/* Live-mode chip removed while Live mode is hidden (engine kept in src/liveMode.ts). */}
      </div>

      {progress && (
        <div className="r-progress" title={`${progress.phase} ${progress.done}/${progress.total}`}>
          <div
            className="r-progress__bar"
            style={{ width: `${progress.total ? Math.round((progress.done / progress.total) * 100) : 0}%` }}
          />
          <span className="r-progress__label">
            {progress.phase === "read" ? "Scanning" : "Writing"} {progress.done}/{progress.total}
          </span>
        </div>
      )}

      {banner && <div className={`r-banner r-banner--${banner.kind}`}>{banner.text}</div>}

      {status.armed && <p className="r-note">{COAUTHORING_WARNING}</p>}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Keep-color picker: which highlight colors count as "keep"
// ---------------------------------------------------------------------------
export function KeepColorPicker(props: {
  keepColors: string[];
  busy: boolean;
  onChange: (colors: string[]) => void;
}): React.ReactElement {
  const set = new Set(props.keepColors.map((c) => c.toLowerCase()));
  const toggle = (color: string): void => {
    const next = new Set(set);
    const key = color.toLowerCase();
    if (next.has(key)) next.delete(key);
    else next.add(key);
    props.onChange([...next]);
  };
  return (
    <details className="r-section">
      <summary>
        Keep colors ({set.size} of {HIGHLIGHT_COLORS.length})
      </summary>
      <p className="r-hint">Highlighted runs in these colors stay visible. Applies on the next Hide / Re-hide.</p>
      <div className="r-colorgrid">
        {HIGHLIGHT_COLORS.map((color) => {
          const key = color.toLowerCase();
          return (
            <label key={key} className="r-color">
              <input type="checkbox" checked={set.has(key)} disabled={props.busy} onChange={() => toggle(color)} />
              <span className="r-swatch" style={{ background: swatch(key) }} />
              {color}
            </label>
          );
        })}
      </div>
    </details>
  );
}

/** Map an OOXML highlight name to a CSS color for the little swatch. */
function swatch(name: string): string {
  const map: Record<string, string> = {
    yellow: "#ffff00",
    green: "#00ff00",
    cyan: "#00ffff",
    magenta: "#ff00ff",
    blue: "#0000ff",
    red: "#ff0000",
    darkblue: "#000080",
    darkcyan: "#008080",
    darkgreen: "#008000",
    darkmagenta: "#800080",
    darkred: "#800000",
    darkyellow: "#808000",
    darkgray: "#808080",
    lightgray: "#c0c0c0",
    black: "#000000",
    white: "#ffffff"
  };
  return map[name] ?? "#cccccc";
}

// ---------------------------------------------------------------------------
// Whole-body mode (avenue ⑦) — the DEFAULT Hide path (one getOoxml + one insertOoxml; fast +
// losslessly reversible, wet-confirmed). This is an OPT-OUT: unchecking it falls back to the slower,
// maximally-compatible per-paragraph commit. (The old lossy "ultra-fast" mode was removed — ⑦ is
// faster AND lossless, so it superseded it.)
// ---------------------------------------------------------------------------
export function WholeBodyModeToggle(props: {
  status: ControllerStatus;
  busy: boolean;
  onChange: (on: boolean) => void;
}): React.ReactElement {
  return (
    <details className="r-section">
      <summary>Speed: Whole-body mode {props.status.pureWholeBody ? "(on)" : "(OFF — compatibility)"}</summary>
      <p className="r-hint">
        Rostrum hides the whole document in a single pass — the fastest path, with formatting and
        highlights fully preserved and a lossless reverse. This is <strong>on by default</strong>;
        uncheck it only if a document&apos;s structure (unusual numbering, section breaks, fields)
        doesn&apos;t survive a Hide → Show All, to use the slower paragraph-by-paragraph path instead.
        Takes effect on the next Hide / Re-hide.
      </p>
      <label className="r-live">
        <input
          type="checkbox"
          checked={props.status.pureWholeBody}
          disabled={props.busy}
          onChange={(e) => props.onChange(e.target.checked)}
        />
        Whole-body mode (fastest) — uncheck for maximum compatibility
      </label>
    </details>
  );
}

// ---------------------------------------------------------------------------
// Track-Changes prompt (decision #14): offer to toggle off + restore
// ---------------------------------------------------------------------------
export function TrackChangesPrompt(props: {
  mode: string;
  onConfirm: () => void;
  onDismiss: () => void;
}): React.ReactElement {
  return (
    <div className="r-modal">
      <div className="r-modal__card">
        <h3>Track Changes is on ({props.mode})</h3>
        <p>
          Rostrum hides text only with Track Changes off, so a partial Undo can&apos;t strand the
          document. Turn it off for this operation? Rostrum will restore your setting afterward.
        </p>
        <div className="r-modal__actions">
          <button className="r-btn r-btn--primary" onClick={props.onConfirm}>
            Turn off &amp; continue
          </button>
          <button className="r-btn r-btn--ghost" onClick={props.onDismiss}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Apply Rostrum styles (gated + reflow warning)
// ---------------------------------------------------------------------------
export function ApplyStyles(props: {
  features: FeatureSupport;
  busy: boolean;
  onApply: () => void;
}): React.ReactElement | null {
  // Hide the control entirely when the host can't do it (keeps the pane honest).
  if (!props.features.canGetStyles || !props.features.canStyleFormat) return null;
  // A single-click action, NOT a collapsible <details> dropdown: the reflow caveat
  // rides along as a hover tooltip (the button's `title`) instead of expanded text,
  // so the pane stays compact and the warning is one hover away.
  return (
    <div className="r-section">
      <button
        className="r-btn"
        disabled={props.busy}
        onClick={props.onApply}
        title={REFLOW_WARNING}
      >
        Apply Rostrum styles
      </button>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Unsupported-host panel (web / old perpetual) with the native fallback
// ---------------------------------------------------------------------------
export function UnsupportedHost(props: { message: string }): React.ReactElement {
  return (
    <div className="r-unsupported">
      <h2>Rostrum needs desktop Word</h2>
      <p>{props.message}</p>
      <h3>Reveal hidden text without the add-in</h3>
      <ol>
        <li>Select the affected text (or the whole document with Ctrl+A).</li>
        <li>
          Open Home ▸ Font dialog (Ctrl+D) and clear the <b>Hidden</b> checkbox, or toggle Home ▸ ¶
          (Show/Hide) to view hidden text.
        </li>
      </ol>
    </div>
  );
}
