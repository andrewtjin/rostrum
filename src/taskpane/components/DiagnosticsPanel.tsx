// The "make debugging insanely easy" surface, in the pane itself.
//
// Inside a live Word host you can't attach a debugger, and the JS console is buried.
// So this panel exposes the shared tracer directly to the user: the capability
// matrix the add-in detected, the manifest/armed state, a verbosity dial, a LIVE
// rolling log of every namespaced, timed, correlated event, and a one-click "Copy
// bug report" that bundles the whole timeline + host info onto the clipboard. A user
// who hits a problem can hand you a complete trace without leaving Word.

import * as React from "react";
import { FeatureSupport } from "../../core/types";
import { formatEntry, LogEntry, LogLevel, tracer } from "../../core/debug";

const LEVELS: LogLevel[] = ["debug", "info", "warn", "error"];

// Cross-cutting suite chrome: the host-capability matrix + the live tracer. This is NOT a
// feature — it's diagnostics for the whole add-in — so it takes only the detected host
// capabilities. Per-feature state (e.g. invisibility's armed/keep-colors) lives in that
// feature's own panel, not here.
export function DiagnosticsPanel(props: {
  features: FeatureSupport | null;
}): React.ReactElement {
  const [entries, setEntries] = React.useState<LogEntry[]>(() => tracer.getBuffer());
  const [level, setLevel] = React.useState<LogLevel>(tracer.getMinLevel());
  const [copied, setCopied] = React.useState(false);

  // Subscribe to the live tracer; each new entry appends to the rolling view.
  React.useEffect(() => {
    setEntries(tracer.getBuffer());
    const off = tracer.subscribe((e) => setEntries((prev) => capTail([...prev, e])));
    return off;
  }, []);

  const changeLevel = (next: LogLevel): void => {
    tracer.setMinLevel(next);
    setLevel(next);
  };

  const clear = (): void => {
    tracer.clear();
    setEntries([]);
  };

  const copyReport = (): void => {
    const report = tracer.bugReport(hostHeader(props));
    void writeClipboard(report).then((ok) => {
      setCopied(ok);
      window.setTimeout(() => setCopied(false), 2000);
    });
  };

  return (
    <details className="r-section r-diag">
      <summary>Diagnostics</summary>

      <FeatureMatrix features={props.features} />

      <div className="r-diag__controls">
        <label>
          Log level{" "}
          <select value={level} onChange={(e) => changeLevel(e.target.value as LogLevel)}>
            {LEVELS.map((l) => (
              <option key={l} value={l}>
                {l}
              </option>
            ))}
          </select>
        </label>
        <button className="r-btn r-btn--ghost" onClick={copyReport}>
          {copied ? "Copied!" : "Copy bug report"}
        </button>
        <button className="r-btn r-btn--ghost" onClick={clear}>
          Clear log
        </button>
      </div>

      <div className="r-log" role="log" aria-live="polite">
        {entries.length === 0 ? (
          <div className="r-log__empty">No events yet. Run Hide / Show All to populate the trace.</div>
        ) : (
          entries
            .slice()
            .reverse()
            .map((e) => (
              <div key={e.seq} className={`r-log__row r-log__row--${e.level}`}>
                <code>{formatEntry(e)}</code>
                {e.data !== undefined && <code className="r-log__data">{safeJson(e.data)}</code>}
              </div>
            ))
        )}
      </div>
    </details>
  );
}

// ---------------------------------------------------------------------------
function FeatureMatrix(props: { features: FeatureSupport | null }): React.ReactElement {
  const f = props.features;
  const rows: Array<[string, boolean | undefined]> = [
    ["Hide (font.hidden, WordApiDesktop 1.2)", f?.canHide],
    ["Manifest (customXml, WordApi 1.4)", f?.canCustomXml],
    ["Track Changes (WordApi 1.4)", f?.canChangeTracking],
    ["Style sizes (WordApi 1.5)", f?.canStyleFormat],
    ["Style borders (WordApiDesktop 1.1)", f?.canStyleBorders],
    ["getStyles (WordApiDesktop 1.4)", f?.canGetStyles]
  ];
  return (
    <table className="r-matrix">
      <tbody>
        {rows.map(([label, ok]) => (
          <tr key={label}>
            <td>{label}</td>
            <td className={ok ? "r-yes" : "r-no"}>{ok === undefined ? "?" : ok ? "yes" : "no"}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Keep only the most recent N entries in the live view (the tracer buffer caps too). */
function capTail(entries: LogEntry[], max = 300): LogEntry[] {
  return entries.length > max ? entries.slice(entries.length - max) : entries;
}

function hostHeader(props: { features: FeatureSupport | null }): Record<string, unknown> {
  const nav = typeof navigator !== "undefined" ? navigator.userAgent : "n/a";
  return { userAgent: nav, features: props.features, when: new Date().toISOString() };
}

function safeJson(v: unknown): string {
  try {
    return JSON.stringify(v);
  } catch {
    return String(v);
  }
}

/** Clipboard write with a textarea fallback for older WebViews; resolves success. */
async function writeClipboard(text: string): Promise<boolean> {
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // fall through to the legacy path
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  } catch {
    return false;
  }
}
