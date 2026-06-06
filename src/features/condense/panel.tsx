// Condense & Shrink — the REACT surface (the deep-linked task pane). The shell mounts this when
// `taskpane.html#condense` is open. It owns the condense-specific hook, its presentational controls, and
// the panel; the headless commands + ribbon descriptor live next door in `contribution.ts` so the ribbon
// runtime and the manifest generator never pull React. The pure engines (`core/shrink`, `core/condense`)
// and the tested `CondenseController` are unchanged.
import * as React from "react";
import { useCallback, useEffect, useRef, useState } from "react";
import { OmissionPattern, TrackChangesMode } from "../../core/types";
import { logger } from "../../core/debug";
import { CondenseController, CondenseStatus, formatSize } from "../../taskpane/condenseController";
import { OpOutcome } from "../../taskpane/controller";
import { DEFAULT_CONDENSE_SETTINGS } from "../../core/settings";
import { FeaturePanelProps } from "../types";

const log = logger("condensePane");

/** A transient banner under the buttons. */
interface Banner {
  kind: "ok" | "info" | "warn" | "error";
  text: string;
}

/** An action that can be retried past the Track-Changes gate (the boolean = auto-toggle TC). */
type GatedAction = (c: CondenseController, autoToggleTC: boolean) => Promise<OpOutcome>;

interface CondenseUi {
  ready: boolean;
  status: CondenseStatus;
  busy: boolean;
  banner: Banner | null;
  trackChangesMode: TrackChangesMode | null;
  shrink: () => void;
  unshrink: () => void;
  condense: () => void;
  condensePilcrows: () => void;
  fullCondense: () => void;
  retainParagraphs: () => void;
  uncondense: () => void;
  setUsePilcrows: (on: boolean) => void;
  setRetainParagraphs: (on: boolean) => void;
  setShrinkParagraphMarks: (on: boolean) => void;
  setLossless: (on: boolean) => void;
  setOmissionPatterns: (patterns: OmissionPattern[]) => void;
  confirmTrackChanges: () => void;
  dismissTrackChanges: () => void;
}

function useCondense(): CondenseUi {
  const [ready, setReady] = useState(false);
  const [status, setStatus] = useState<CondenseStatus>({ settings: DEFAULT_CONDENSE_SETTINGS });
  const [busy, setBusy] = useState(false);
  const [banner, setBanner] = useState<Banner | null>(null);
  const [trackChangesMode, setTrackChangesMode] = useState<TrackChangesMode | null>(null);

  const controllerRef = useRef<CondenseController | null>(null);
  const pendingRef = useRef<GatedAction | null>(null);

  useEffect(() => {
    const controller = new CondenseController({});
    controllerRef.current = controller;
    setStatus(controller.status());
    setReady(true);
  }, []);

  const runAction = useCallback(
    async (fn: GatedAction, autoToggleTC = false) => {
      const controller = controllerRef.current;
      if (!controller || busy) return;
      setBusy(true);
      setBanner(null);
      try {
        const out = await fn(controller, autoToggleTC);
        switch (out.status) {
          case "ok":
            setStatus(controller.status());
            setBanner({ kind: "ok", text: out.message });
            break;
          case "trackChanges":
            pendingRef.current = fn;
            setTrackChangesMode(out.mode);
            break;
          case "cancelled":
            break;
          case "error":
            setBanner({ kind: "error", text: out.message });
            break;
        }
      } catch (e) {
        log.caught("condense action failed", e);
        setBanner({ kind: "error", text: String((e as Error)?.message ?? e) });
      } finally {
        setBusy(false);
      }
    },
    [busy]
  );

  const update = useCallback((patch: Parameters<CondenseController["setSettings"]>[0]) => {
    const controller = controllerRef.current;
    if (!controller) return;
    controller.setSettings(patch);
    setStatus(controller.status());
  }, []);

  const confirmTrackChanges = useCallback(() => {
    const fn = pendingRef.current;
    pendingRef.current = null;
    setTrackChangesMode(null);
    if (fn) void runAction(fn, true);
  }, [runAction]);

  const dismissTrackChanges = useCallback(() => {
    pendingRef.current = null;
    setTrackChangesMode(null);
    setBanner({ kind: "warn", text: "Turn Track Changes off in Word (Review ▸ Track Changes), then try again." });
  }, []);

  return {
    ready,
    status,
    busy,
    banner,
    trackChangesMode,
    shrink: () => void runAction((c, a) => c.shrink(a)),
    unshrink: () => void runAction((c, a) => c.unshrink(a)),
    condense: () => void runAction((c, a) => c.condense(a)),
    condensePilcrows: () => void runAction((c, a) => c.condenseWithPilcrows(a)),
    fullCondense: () => void runAction((c, a) => c.fullCondense(a)),
    retainParagraphs: () => void runAction((c, a) => c.retainParagraphsCondense(a)),
    uncondense: () => void runAction((c, a) => c.uncondense(a)),
    setUsePilcrows: (on) => update({ usePilcrows: on }),
    setRetainParagraphs: (on) => update({ retainParagraphs: on }),
    setShrinkParagraphMarks: (on) => update({ shrinkParagraphMarks: on }),
    setLossless: (on) => update({ reversal: on ? "marker" : "none" }),
    setOmissionPatterns: (patterns) => update({ omissionPatterns: patterns }),
    confirmTrackChanges,
    dismissTrackChanges,
  };
}

// ===========================================================================
// Presentational controls
// ===========================================================================

function ShrinkRow(props: {
  busy: boolean;
  lastShrink: number | null | undefined;
  onShrink: () => void;
  onUnshrink: () => void;
}): React.ReactElement {
  const sizeLabel = props.lastShrink === undefined ? "—" : formatSize(props.lastShrink);
  return (
    <div className="r-section">
      <div className="r-buttons">
        <button className="r-btn r-btn--primary" disabled={props.busy} onClick={props.onShrink}>
          Shrink
        </button>
        <button className="r-btn" disabled={props.busy} onClick={props.onUnshrink}>
          Unshrink
        </button>
      </div>
      <p className="r-hint">
        Current shrink size: <b>{sizeLabel}</b>. Shrink cycles non-underlined text down a size each press
        (8→7→6→5→4→Normal); the underlined cut, highlights, cites, and headings stay full-size.
      </p>
    </div>
  );
}

function CondenseRow(props: {
  busy: boolean;
  onCondense: () => void;
  onPilcrows: () => void;
  onFull: () => void;
  onRetain: () => void;
  onUncondense: () => void;
}): React.ReactElement {
  return (
    <div className="r-section">
      <div className="r-buttons">
        <button className="r-btn r-btn--primary" disabled={props.busy} onClick={props.onCondense}>
          Condense
        </button>
        <button className="r-btn" disabled={props.busy} onClick={props.onUncondense}>
          Uncondense
        </button>
      </div>
      <p className="r-hint">Condense uses your settings below. Or pick a mode directly:</p>
      <div className="r-buttons">
        <button className="r-btn r-btn--ghost" disabled={props.busy} onClick={props.onPilcrows}>
          Condense w/ pilcrows
        </button>
        <button className="r-btn r-btn--ghost" disabled={props.busy} onClick={props.onFull}>
          Full condense
        </button>
        <button className="r-btn r-btn--ghost" disabled={props.busy} onClick={props.onRetain}>
          Retain paragraphs
        </button>
      </div>
    </div>
  );
}

function ModeToggles(props: {
  busy: boolean;
  status: CondenseStatus;
  onUsePilcrows: (on: boolean) => void;
  onRetain: (on: boolean) => void;
  onShrinkMarks: (on: boolean) => void;
  onLossless: (on: boolean) => void;
}): React.ReactElement {
  const s = props.status.settings;
  return (
    <details className="r-section" open>
      <summary>Condense settings</summary>
      <label className="r-live">
        <input type="checkbox" checked={s.usePilcrows} disabled={props.busy} onChange={(e) => props.onUsePilcrows(e.target.checked)} />
        Use pilcrows (show a ¶ at each former paragraph break)
      </label>
      <label className="r-live">
        <input type="checkbox" checked={s.retainParagraphs} disabled={props.busy} onChange={(e) => props.onRetain(e.target.checked)} />
        Retain paragraphs (keep structure; only drop blank lines)
      </label>
      <label className="r-live">
        <input type="checkbox" checked={s.shrinkParagraphMarks} disabled={props.busy} onChange={(e) => props.onShrinkMarks(e.target.checked)} />
        Shrink ¶ marks to 6pt when shrinking (tightens line spacing)
      </label>
      <label className="r-live">
        <input type="checkbox" checked={s.reversal === "marker"} disabled={props.busy} onChange={(e) => props.onLossless(e.target.checked)} />
        Losslessly reversible (uncheck only for a faster, one-way condense)
      </label>
    </details>
  );
}

function OmissionEditor(props: {
  busy: boolean;
  patterns: OmissionPattern[];
  onChange: (patterns: OmissionPattern[]) => void;
}): React.ReactElement {
  const { patterns } = props;
  const setAt = (i: number, patch: Partial<OmissionPattern>): void => {
    const next = patterns.map((p, j) => (j === i ? { ...p, ...patch } : p));
    props.onChange(next);
  };
  const removeAt = (i: number): void => props.onChange(patterns.filter((_, j) => j !== i));
  const add = (): void => props.onChange([...patterns, { open: "[", close: "]", keyword: "Omitted" }]);
  return (
    <details className="r-section">
      <summary>Omission markers ({patterns.length})</summary>
      <p className="r-hint">
        Shrink restores text inside these markers to full size, so an “[…Omitted…]” indicator stays
        readable. A bracketed span counts only when it contains the keyword.
      </p>
      {patterns.map((p, i) => (
        <div key={i} className="r-omission">
          <input
            aria-label="open delimiter"
            className="r-omission__delim"
            value={p.open}
            disabled={props.busy}
            onChange={(e) => setAt(i, { open: e.target.value })}
          />
          <input
            aria-label="keyword"
            className="r-omission__kw"
            value={p.keyword}
            placeholder="keyword"
            disabled={props.busy}
            onChange={(e) => setAt(i, { keyword: e.target.value })}
          />
          <input
            aria-label="close delimiter"
            className="r-omission__delim"
            value={p.close}
            disabled={props.busy}
            onChange={(e) => setAt(i, { close: e.target.value })}
          />
          <button className="r-btn r-btn--ghost" disabled={props.busy} aria-label="remove marker" onClick={() => removeAt(i)}>
            ✕
          </button>
        </div>
      ))}
      <button className="r-btn r-btn--ghost" disabled={props.busy} onClick={add}>
        Add marker
      </button>
    </details>
  );
}

function TrackChangesPrompt(props: { mode: string; onConfirm: () => void; onDismiss: () => void }): React.ReactElement {
  const confirmRef = useRef<HTMLButtonElement | null>(null);
  useEffect(() => {
    confirmRef.current?.focus();
  }, []);
  return (
    <div className="r-modal" onClick={props.onDismiss}>
      <div
        className="r-modal__card"
        role="dialog"
        aria-modal="true"
        aria-label="Track Changes is on"
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === "Escape") props.onDismiss();
        }}
      >
        <h3>Track Changes is on ({props.mode})</h3>
        <p>
          Rostrum condenses text only with Track Changes off, so a partial Undo can’t strand the document.
          Turn it off for this operation? Rostrum restores your setting afterward.
        </p>
        <div className="r-modal__actions">
          <button className="r-btn r-btn--primary" ref={confirmRef} onClick={props.onConfirm}>
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

// ===========================================================================
// The feature's pane panel
// ===========================================================================
export function CondensePanel(_props: FeaturePanelProps): React.ReactElement {
  const ui = useCondense();

  if (!ui.ready) {
    return <div className="r-loading">Preparing Condense & Shrink…</div>;
  }

  return (
    <div className="r-feature">
      <ShrinkRow busy={ui.busy} lastShrink={ui.status.lastShrinkHalfPts} onShrink={ui.shrink} onUnshrink={ui.unshrink} />
      <CondenseRow
        busy={ui.busy}
        onCondense={ui.condense}
        onPilcrows={ui.condensePilcrows}
        onFull={ui.fullCondense}
        onRetain={ui.retainParagraphs}
        onUncondense={ui.uncondense}
      />

      {ui.banner && (
        <div
          className={`r-banner r-banner--${ui.banner.kind}`}
          role={ui.banner.kind === "error" || ui.banner.kind === "warn" ? "alert" : "status"}
          aria-live={ui.banner.kind === "error" || ui.banner.kind === "warn" ? "assertive" : "polite"}
        >
          {ui.banner.text}
        </div>
      )}

      <ModeToggles
        busy={ui.busy}
        status={ui.status}
        onUsePilcrows={ui.setUsePilcrows}
        onRetain={ui.setRetainParagraphs}
        onShrinkMarks={ui.setShrinkParagraphMarks}
        onLossless={ui.setLossless}
      />
      <OmissionEditor busy={ui.busy} patterns={ui.status.settings.omissionPatterns} onChange={ui.setOmissionPatterns} />
      <p className="r-hint">Tip: select the card text first, or place the cursor in a card. Shrink &amp; Condense act on your selection (or the current paragraph).</p>

      {ui.trackChangesMode && (
        <TrackChangesPrompt mode={ui.trackChangesMode} onConfirm={ui.confirmTrackChanges} onDismiss={ui.dismissTrackChanges} />
      )}
    </div>
  );
}
