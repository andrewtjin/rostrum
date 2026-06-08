// Invisibility Mode — the REACT surface (the compact deep-linked task pane). The shell mounts
// this when `taskpane.html#invisibility` is open. It owns the invisibility-specific hook, its
// presentational controls, and the panel itself; the headless commands + ribbon descriptor live
// next door in `contribution.ts` so the ribbon runtime and the manifest generator never pull
// React. The pure engine (`core/`) and the tested `RostrumController` are unchanged.
import * as React from "react";
import { useCallback, useEffect, useRef, useState } from "react";
import { FeatureSupport, TrackChangesMode } from "../../core/types";
import { HIGHLIGHT_COLORS } from "../../core/styles";
import { COAUTHORING_WARNING } from "../../core/guards";
import { REFLOW_WARNING } from "../../core/officeStyles";
import { ProgressInfo } from "../../core/officeWordPort";
import { formatProgress, progressPercent } from "../../core/progress";
import { LiveMode } from "../../liveMode";
import { logger } from "../../core/debug";
import { ControllerStatus, OpOutcome, RostrumController } from "../../taskpane/controller";
import { FeaturePanelProps } from "../types";

const log = logger("invis");

// ===========================================================================
// Hook — the invisibility feature's state + actions over the tested controller.
// The shell guarantees a ready, supported host before mounting this, so there is no
// host-bootstrap/unsupported phase here: just `ready`.
// ===========================================================================

/** A transient banner under the buttons. */
export interface Banner {
  kind: "ok" | "info" | "warn" | "error";
  text: string;
}

export interface InvisibilityUi {
  ready: boolean;
  status: ControllerStatus;
  busy: boolean;
  progress: ProgressInfo | null;
  banner: Banner | null;
  /** Non-null when reading the document at init failed — the panel shows an error, not buttons. */
  initError: string | null;
  trackChangesMode: TrackChangesMode | null;
  liveOn: boolean;
  hide: () => void;
  showAll: () => void;
  applyStyles: () => void;
  cancel: () => void;
  setKeepColors: (colors: string[]) => void;
  setPureWholeBody: (on: boolean) => void;
  confirmTrackChanges: () => void;
  dismissTrackChanges: () => void;
  toggleLive: () => void;
}

function useInvisibility(features: FeatureSupport): InvisibilityUi {
  const [ready, setReady] = useState(false);
  const [status, setStatus] = useState<ControllerStatus>({
    armed: false,
    keepColors: [],
    pureWholeBody: true,
  });
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState<ProgressInfo | null>(null);
  const [banner, setBanner] = useState<Banner | null>(null);
  const [initError, setInitError] = useState<string | null>(null);
  const [trackChangesMode, setTrackChangesMode] = useState<TrackChangesMode | null>(null);
  const [liveOn, setLiveOn] = useState(false);

  const controllerRef = useRef<RostrumController | null>(null);
  const liveRef = useRef<LiveMode | null>(null);

  // ---- build the controller once (host is ready when this panel mounts) ----
  useEffect(() => {
    let cancelled = false;
    const controller = new RostrumController({ features, onProgress: (p) => setProgress(p) });
    controllerRef.current = controller;
    liveRef.current = new LiveMode({});
    controller
      .init()
      .then((initial) => {
        if (cancelled) return;
        setStatus(initial);
        setReady(true);
        if (initial.armed) {
          setBanner({ kind: "info", text: "This document has Rostrum invisibility ON." });
        }
      })
      .catch((e) => {
        log.caught("invisibility init failed", e);
        if (cancelled) return;
        // The host is fine (the shell already gated that), but reading THIS document failed.
        // Surface an error state — NOT the action buttons, which would otherwise operate on a
        // controller whose init never completed. Show All still works from the ribbon.
        setInitError(msg(e));
        setReady(true);
      });
    return () => {
      cancelled = true;
      void liveRef.current?.stop();
    };
  }, [features]);

  // ---- external-change re-sync (ribbon uses a separate controller instance) ----
  // A Hide / Re-hide / Show All run from the RIBBON lands in the document manifest but not
  // in this pane's in-memory state, so the green "Invisibility ON" chip goes stale. Re-read
  // whenever the pane regains focus/visibility (i.e. the user returns from the ribbon).
  useEffect(() => {
    const resync = (): void => {
      if (typeof document !== "undefined" && document.hidden) return;
      const controller = controllerRef.current;
      if (!controller) return;
      controller
        .refreshFromDocument()
        .then(setStatus)
        .catch((e) => log.caught("status re-sync from document failed (ignored)", e));
    };
    window.addEventListener("focus", resync);
    document.addEventListener("visibilitychange", resync);
    return () => {
      window.removeEventListener("focus", resync);
      document.removeEventListener("visibilitychange", resync);
    };
  }, []);

  // ---- action plumbing ----------------------------------------------------
  const runAction = useCallback(
    async (fn: (c: RostrumController) => Promise<OpOutcome>) => {
      const controller = controllerRef.current;
      if (!controller || busy) return;
      setBusy(true);
      setBanner(null);
      setProgress(null);
      try {
        const out = await fn(controller);
        switch (out.status) {
          case "ok":
            setStatus(controller.status());
            setBanner({ kind: "ok", text: out.message });
            break;
          case "trackChanges":
            setTrackChangesMode(out.mode);
            break;
          case "cancelled":
            setBanner({ kind: "warn", text: "Cancelled — no changes were written." });
            break;
          case "error":
            setBanner({ kind: "error", text: out.message });
            break;
        }
      } finally {
        setBusy(false);
        setProgress(null);
      }
    },
    [busy]
  );

  const hide = useCallback(() => void runAction((c) => c.hide()), [runAction]);
  const showAll = useCallback(() => void runAction((c) => c.showAll()), [runAction]);
  const applyStyles = useCallback(() => void runAction((c) => c.applyStyles()), [runAction]);
  const cancel = useCallback(() => controllerRef.current?.cancel(), []);

  const setKeepColors = useCallback((colors: string[]) => {
    const controller = controllerRef.current;
    if (!controller) return;
    controller.setKeepColors(colors);
    setStatus(controller.status());
  }, []);

  const setPureWholeBody = useCallback((on: boolean) => {
    const controller = controllerRef.current;
    if (!controller) return;
    controller.setPureWholeBody(on);
    setStatus(controller.status());
  }, []);

  const confirmTrackChanges = useCallback(() => {
    // Hide is the only Track-Changes–gated pane op, so "Turn off & continue" always retries Hide.
    setTrackChangesMode(null);
    void runAction((c) => c.hide(true));
  }, [runAction]);

  const dismissTrackChanges = useCallback(() => {
    setTrackChangesMode(null);
    setBanner({
      kind: "warn",
      text: "Turn Track Changes off in Word (Review ▸ Track Changes), then try again.",
    });
  }, []);

  const toggleLive = useCallback(() => {
    const live = liveRef.current;
    if (!live) return;
    if (live.isActive) {
      void live.stop().then(() => setLiveOn(false));
    } else {
      void live
        .start()
        .then(() => setLiveOn(true))
        .catch((e) => setBanner({ kind: "warn", text: `Live mode unavailable: ${msg(e)}` }));
    }
  }, []);

  return {
    ready,
    status,
    busy,
    progress,
    banner,
    initError,
    trackChangesMode,
    liveOn,
    hide,
    showAll,
    applyStyles,
    cancel,
    setKeepColors,
    setPureWholeBody,
    confirmTrackChanges,
    dismissTrackChanges,
    toggleLive,
  };
}

// ===========================================================================
// Presentational controls (pure render; props in, no Office.js / engine logic).
// ===========================================================================

function Buttons(props: {
  status: ControllerStatus;
  features: FeatureSupport;
  busy: boolean;
  onHide: () => void;
  onShowAll: () => void;
  onApplyStyles: () => void;
  onCancel: () => void;
}): React.ReactElement {
  const { status, busy, features } = props;
  // Apply Styles is host-gated (needs getStyles + styleFormat); omit it where unsupported so the
  // row stays honest. It sits next to Show All as a single-click action; the reflow caveat rides
  // along as a hover tooltip (keeps the pane compact — no extra section).
  const canApplyStyles = features.canGetStyles && features.canStyleFormat;
  return (
    <div className="r-buttons">
      <button className="r-btn r-btn--primary" disabled={busy} onClick={props.onHide}>
        {status.armed ? "Hide (re-arm)" : "Hide"}
      </button>
      <button className="r-btn" disabled={busy} onClick={props.onShowAll}>
        Show All
      </button>
      {canApplyStyles && (
        <button className="r-btn" disabled={busy} onClick={props.onApplyStyles} title={REFLOW_WARNING}>
          Apply Rostrum styles
        </button>
      )}
      {busy && (
        <button className="r-btn r-btn--ghost" onClick={props.onCancel}>
          Cancel
        </button>
      )}
    </div>
  );
}

function StatusBar(props: {
  status: ControllerStatus;
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
      </div>

      {progress && (
        <div
          className="r-progress"
          title={formatProgress(progress)}
          role="progressbar"
          aria-label="Rostrum progress"
          aria-valuemin={0}
          aria-valuemax={100}
          {...(progress.total > 0 ? { "aria-valuenow": progressPercent(progress) } : {})}
        >
          <div className="r-progress__bar" style={{ width: `${progressPercent(progress)}%` }} />
          <span className="r-progress__label">{formatProgress(progress)}</span>
        </div>
      )}

      {banner && (
        <div
          className={`r-banner r-banner--${banner.kind}`}
          role={banner.kind === "error" || banner.kind === "warn" ? "alert" : "status"}
          aria-live={banner.kind === "error" || banner.kind === "warn" ? "assertive" : "polite"}
        >
          {banner.text}
        </div>
      )}
      {status.armed && <p className="r-note">{COAUTHORING_WARNING}</p>}
    </div>
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
    white: "#ffffff",
  };
  return map[name] ?? "#cccccc";
}

function KeepColorPicker(props: {
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
              <span className="r-swatch" aria-hidden="true" style={{ background: swatch(key) }} />
              {color}
            </label>
          );
        })}
      </div>
    </details>
  );
}

function WholeBodyModeToggle(props: {
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

function TrackChangesPrompt(props: {
  mode: string;
  onConfirm: () => void;
  onDismiss: () => void;
}): React.ReactElement {
  // A true blocking modal, so make it a real dialog: labelled, focus moved in on open, Tab trapped
  // between the two actions, Escape to dismiss, and focus restored to the trigger on close.
  const cardRef = useRef<HTMLDivElement | null>(null);
  const confirmRef = useRef<HTMLButtonElement | null>(null);
  const previouslyFocused = useRef<Element | null>(null);

  useEffect(() => {
    previouslyFocused.current = typeof document !== "undefined" ? document.activeElement : null;
    confirmRef.current?.focus();
    return () => {
      (previouslyFocused.current as HTMLElement | null)?.focus?.();
    };
  }, []);

  const onKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === "Escape") {
      e.preventDefault();
      props.onDismiss();
      return;
    }
    if (e.key !== "Tab") return;
    const focusables = cardRef.current?.querySelectorAll<HTMLButtonElement>("button");
    if (!focusables || focusables.length === 0) return;
    const first = focusables[0];
    const last = focusables[focusables.length - 1];
    if (e.shiftKey && document.activeElement === first) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && document.activeElement === last) {
      e.preventDefault();
      first.focus();
    }
  };

  return (
    <div className="r-modal" onClick={props.onDismiss}>
      <div
        className="r-modal__card"
        role="dialog"
        aria-modal="true"
        aria-labelledby="r-tc-title"
        aria-describedby="r-tc-desc"
        ref={cardRef}
        onClick={(e) => e.stopPropagation()}
        onKeyDown={onKeyDown}
      >
        <h3 id="r-tc-title">Track Changes is on ({props.mode})</h3>
        <p id="r-tc-desc">
          Rostrum hides text only with Track Changes off, so a partial Undo can&apos;t strand the
          document. Turn it off for this operation? Rostrum will restore your setting afterward.
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

/** The invisibility manifest state, shown in this feature's panel. */
function ManifestState(props: { status: ControllerStatus }): React.ReactElement {
  return (
    <p className="r-hint">
      Armed: <b>{props.status.armed ? "yes" : "no"}</b> · keep-colors:{" "}
      <b>{props.status.keepColors.length ? props.status.keepColors.join(", ") : "(all)"}</b>
    </p>
  );
}

// ===========================================================================
// The feature's pane panel — what the shell mounts when Invisibility is selected.
// ===========================================================================
export function InvisibilityPanel({ features }: FeaturePanelProps): React.ReactElement {
  const ui = useInvisibility(features);

  if (!ui.ready) {
    return <div className="r-loading">Preparing Invisibility Mode…</div>;
  }

  // Init failed reading this document — show an error, not actionable buttons over a
  // controller that never finished initializing. Show All still works from the ribbon.
  if (ui.initError) {
    return (
      <div className="r-feature">
        <div className="r-banner r-banner--error" role="alert">Couldn&apos;t read this document: {ui.initError}</div>
        <p className="r-hint">Reload the pane to try again.</p>
      </div>
    );
  }

  return (
    <div className="r-feature">
      <Buttons
        status={ui.status}
        features={features}
        busy={ui.busy}
        onHide={ui.hide}
        onShowAll={ui.showAll}
        onApplyStyles={ui.applyStyles}
        onCancel={ui.cancel}
      />
      <StatusBar status={ui.status} progress={ui.progress} banner={ui.banner} />

      {/* Live mode (keep the paragraph I'm typing in visible) is HIDDEN for now. The engine
          stays in src/liveMode.ts and is still wired through useInvisibility (ui.liveOn /
          ui.toggleLive); to re-enable, restore the control here. */}

      <KeepColorPicker keepColors={ui.status.keepColors} busy={ui.busy} onChange={ui.setKeepColors} />
      <WholeBodyModeToggle status={ui.status} busy={ui.busy} onChange={ui.setPureWholeBody} />
      {/* "Load Rostrum on every document" is NOT invisibility's concern: it's now the Trusted-Catalog
          install (the in-app Always-On toggle was retired). The Settings pane explains it. */}
      <ManifestState status={ui.status} />

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

/** Local error-to-message helper (kept private to the panel). */
function msg(e: unknown): string {
  return String((e as Error)?.message ?? e);
}
