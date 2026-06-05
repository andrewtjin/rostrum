// The single React hook the task pane is built on. It owns the controller +
// live-mode instances (in refs, so they survive re-renders) and projects their
// state into render-ready values. All the real logic is in the (tested)
// RostrumController; this hook is the thin React boundary.

import { useCallback, useEffect, useRef, useState } from "react";
import { assertCanRun, detectFeatureSupport, UnsupportedHostError } from "../core/guards";
import { FeatureSupport, TrackChangesMode } from "../core/types";
import { ProgressInfo } from "../core/officeWordPort";
import { LiveMode } from "../liveMode";
import { logger } from "../core/debug";
import { ControllerStatus, OpOutcome, RostrumController } from "./controller";

/** Top-level UI phase the pane renders. */
export type AppPhase = "loading" | "unsupported" | "ready";

/** A transient banner under the buttons. */
export interface Banner {
  kind: "ok" | "info" | "warn" | "error";
  text: string;
}

export interface RostrumUi {
  phase: AppPhase;
  unsupportedMessage: string | null;
  features: FeatureSupport | null;
  status: ControllerStatus;
  busy: boolean;
  progress: ProgressInfo | null;
  banner: Banner | null;
  /** Non-null while the Track-Changes prompt should be shown. */
  trackChangesMode: TrackChangesMode | null;
  liveOn: boolean;

  hide: () => void;
  reHide: () => void;
  showAll: () => void;
  applyStyles: () => void;
  cancel: () => void;
  setKeepColors: (colors: string[]) => void;
  setPureWholeBody: (on: boolean) => void;
  confirmTrackChanges: () => void;
  dismissTrackChanges: () => void;
  toggleLive: () => void;
}

const log = logger("pane");

export function useRostrum(): RostrumUi {
  const [phase, setPhase] = useState<AppPhase>("loading");
  const [unsupportedMessage, setUnsupportedMessage] = useState<string | null>(null);
  const [features, setFeatures] = useState<FeatureSupport | null>(null);
  const [status, setStatus] = useState<ControllerStatus>({
    armed: false,
    keepColors: [],
    pureWholeBody: true
  });
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState<ProgressInfo | null>(null);
  const [banner, setBanner] = useState<Banner | null>(null);
  const [trackChangesMode, setTrackChangesMode] = useState<TrackChangesMode | null>(null);
  const [liveOn, setLiveOn] = useState(false);

  const controllerRef = useRef<RostrumController | null>(null);
  const liveRef = useRef<LiveMode | null>(null);
  // Which mutation triggered the Track-Changes prompt, so "Turn off & continue"
  // retries the SAME action with auto-toggle.
  const pendingTcAction = useRef<"hide" | "reHide" | null>(null);

  // ---- one-time host bootstrap --------------------------------------------
  useEffect(() => {
    let cancelled = false;
    // Office.onReady resolves once the host + office.js are ready.
    Office.onReady()
      .then(async () => {
        const support = detectFeatureSupport(Office.context.requirements);
        setFeatures(support);
        try {
          assertCanRun(support); // throws UnsupportedHostError on web / old perpetual
        } catch (e) {
          if (e instanceof UnsupportedHostError) {
            setUnsupportedMessage(e.message);
            setPhase("unsupported");
            return;
          }
          throw e;
        }
        const controller = new RostrumController({
          features: support,
          onProgress: (p) => setProgress(p)
        });
        controllerRef.current = controller;
        liveRef.current = new LiveMode({});
        const initial = await controller.init();
        if (cancelled) return;
        setStatus(initial);
        setPhase("ready");
        if (initial.armed) {
          setBanner({ kind: "info", text: "This document has Rostrum invisibility ON." });
        }
      })
      .catch((e) => {
        log.caught("pane bootstrap failed", e);
        setUnsupportedMessage(`Rostrum failed to start: ${String((e as Error)?.message ?? e)}`);
        setPhase("unsupported");
      });
    return () => {
      cancelled = true;
      void liveRef.current?.stop();
    };
  }, []);

  // ---- external-change re-sync --------------------------------------------
  // A Hide / Re-hide / Show All run from the RIBBON uses a SEPARATE controller instance, so its
  // change lands in the document manifest but not in this pane's in-memory state — the green
  // "Invisibility ON" indicator goes stale. Re-read the document whenever the pane regains focus or
  // visibility (i.e. the user returns to it after using the ribbon), so the indicator catches up.
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
    async (
      fn: (c: RostrumController) => Promise<OpOutcome>,
      onOk?: (out: Extract<OpOutcome, { status: "ok" }>) => void,
      tcAction?: "hide" | "reHide"
    ) => {
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
            onOk?.(out);
            break;
          case "trackChanges":
            pendingTcAction.current = tcAction ?? null;
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

  const hide = useCallback(() => void runAction((c) => c.hide(), undefined, "hide"), [runAction]);
  const reHide = useCallback(() => void runAction((c) => c.reHide(), undefined, "reHide"), [runAction]);
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
    const action = pendingTcAction.current;
    setTrackChangesMode(null);
    pendingTcAction.current = null;
    // Retry the same mutation with auto-toggle on.
    void runAction((c) => (action === "reHide" ? c.reHide(true) : c.hide(true)));
  }, [runAction]);

  const dismissTrackChanges = useCallback(() => {
    setTrackChangesMode(null);
    pendingTcAction.current = null;
    setBanner({
      kind: "warn",
      text: "Turn Track Changes off in Word (Review ▸ Track Changes), then try again."
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
        .catch((e) =>
          setBanner({ kind: "warn", text: `Live mode unavailable: ${String((e as Error)?.message ?? e)}` })
        );
    }
  }, []);

  return {
    phase,
    unsupportedMessage,
    features,
    status,
    busy,
    progress,
    banner,
    trackChangesMode,
    liveOn,
    hide,
    reHide,
    showAll,
    applyStyles,
    cancel,
    setKeepColors,
    setPureWholeBody,
    confirmTrackChanges,
    dismissTrackChanges,
    toggleLive
  };
}
