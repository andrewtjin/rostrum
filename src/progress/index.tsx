// The ribbon progress pop-out PAGE (runs inside the Office dialog opened by progress/host.ts).
// It renders the live scan/write bar from parent→child messages and offers Cancel. It holds no
// engine state: the driver in the ribbon runtime owns the operation and streams ticks here, then
// closes this dialog on completion. The op's label arrives in the URL hash (#Hide).
import * as React from "react";
import { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import type { ProgressInfo } from "../core/officeWordPort";
import { formatProgress, progressPercent } from "../core/progress";
// Pop-out stylesheet — bundled so it ships content-hashed (same staleness rationale as the pane).
import "./progress.css";

/** Parent→child messages, mirrored from progress/host.ts (parsed defensively). */
type ParentMsg =
  | { kind: "progress"; phase: ProgressInfo["phase"]; done: number; total: number }
  | { kind: "done"; status: string; message?: string; diagnostics?: string };

type UiState =
  | { phase: "working"; progress: ProgressInfo | null }
  | { phase: "done"; status: string; message?: string; diagnostics?: string };

/** The op label is passed in the URL hash so the window can title itself before any message. */
function labelFromHash(): string {
  const h = window.location.hash;
  return h && h.length > 1 ? decodeURIComponent(h.slice(1)) : "Working";
}

/** Post a control message to the parent (the ribbon driver). No-op outside a dialog host. */
function messageParent(m: string): void {
  if (typeof Office !== "undefined" && Office.context?.ui?.messageParent) {
    try {
      Office.context.ui.messageParent(m);
    } catch {
      // Outside a real dialog host there's no parent — nothing to do.
    }
  }
}

function ProgressApp(): React.ReactElement {
  const label = labelFromHash();
  const [state, setState] = useState<UiState>({ phase: "working", progress: null });
  const [cancelRequested, setCancelRequested] = useState(false);
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    let cancelled = false;
    void Office.onReady().then(() => {
      if (cancelled) return;
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
        const raw = (arg as { message?: string }).message;
        if (!raw) return;
        let msg: Partial<ParentMsg>;
        try {
          msg = JSON.parse(raw) as Partial<ParentMsg>;
        } catch {
          return; // ignore anything that isn't our JSON protocol
        }
        if (msg.kind === "progress") {
          setState({ phase: "working", progress: { phase: msg.phase!, done: msg.done!, total: msg.total! } });
        } else if (msg.kind === "done") {
          setState({ phase: "done", status: msg.status ?? "ok", message: msg.message, diagnostics: msg.diagnostics });
        }
      });
      // Signal readiness so the driver replays the current progress / final state.
      messageParent("ready");
    });
    return () => {
      cancelled = true;
    };
  }, []);

  const onCancel = (): void => {
    setCancelRequested(true);
    messageParent("cancel");
  };

  const done = state.phase === "done";
  const status = done ? state.status : null;
  // Keep the window open (with a Close button) only for outcomes the user must read + act on
  // (an error, or a Track-Changes "blocked" with instructions). Clean outcomes auto-close.
  const keepOpen = done && (status === "error" || status === "blocked");
  const progress = state.phase === "working" ? state.progress : null;
  const pct = progress ? progressPercent(progress) : done ? 100 : 0;
  const line = done
    ? state.message ?? (status === "ok" ? "Done." : status ?? "")
    : progress
      ? formatProgress(progress)
      : cancelRequested
        ? "Cancelling…"
        : "Working…";
  // Indeterminate while we have no numeric total yet (and aren't done): a calmer animated track.
  const indeterminate = !done && !progress;

  const onClose = (): void => messageParent("close");

  // Diagnostics arrive ONLY on a kept-open failure (error/blocked). The bug report is the Office
  // error code + debugInfo a ribbon op can surface nowhere else, so the pop-out shows it with a
  // one-click Copy (the user pastes it back). A manual text-selection fallback covers any host
  // whose dialog webview blocks the async clipboard API.
  const diagnostics = state.phase === "done" ? state.diagnostics : undefined;
  const selectDiagnostics = (): void => {
    const pre = document.getElementById("p-diag");
    if (!pre) return;
    const range = document.createRange();
    range.selectNodeContents(pre);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
  };
  const onCopy = (): void => {
    if (!diagnostics) return;
    const clip = (navigator as Navigator & { clipboard?: { writeText?: (t: string) => Promise<void> } }).clipboard;
    if (clip?.writeText) {
      clip.writeText(diagnostics).then(() => setCopied(true), selectDiagnostics);
    } else {
      selectDiagnostics();
    }
  };

  return (
    <div className="p-card">
      <div className="p-title">{label}</div>
      <div
        className={`p-track${indeterminate ? " p-track--idle" : ""}`}
        role="progressbar"
        aria-label={label}
        aria-valuemin={0}
        aria-valuemax={100}
        {...(indeterminate ? {} : { "aria-valuenow": pct })}
      >
        <div className="p-bar" style={{ width: `${pct}%` }} />
      </div>
      <div className={`p-line${keepOpen ? " p-line--error" : ""}`}>{line}</div>
      {/* Visually-hidden live region: announces the FINAL result once (not every numeric tick —
          the progressbar role already conveys ongoing progress), so a screen-reader user hears
          completion even though the window may auto-close right after. */}
      <div className="p-sr" role="status" aria-live="polite" aria-atomic="true">
        {done ? line : ""}
      </div>
      {!done && (
        <button className="p-btn" onClick={onCancel} disabled={cancelRequested}>
          {cancelRequested ? "Cancelling…" : "Cancel"}
        </button>
      )}
      {keepOpen && (
        <button className="p-btn" onClick={onClose} autoFocus>
          Close
        </button>
      )}
      {keepOpen && diagnostics && (
        <details className="p-diag-wrap">
          <summary>Diagnostics</summary>
          <pre id="p-diag" className="p-diag">
            {diagnostics}
          </pre>
          <button className="p-btn" onClick={onCopy}>
            {copied ? "Copied ✓" : "Copy diagnostics"}
          </button>
        </details>
      )}
    </div>
  );
}

const container = document.getElementById("root");
if (container) {
  createRoot(container).render(<ProgressApp />);
}
