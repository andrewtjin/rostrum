// Ribbon progress pop-out — the headless driver. A ribbon command runs with NO task pane (to
// keep the document's reading space), so this is how a long Hide/Show/Apply op gives live
// feedback: a tiny Office dialog that appears only while the op runs and closes itself on
// completion. Word add-ins can't draw a moving bar on the ribbon itself, so a transient dialog
// is the real option — it floats over Word and never docks, so it costs zero persistent space.
//
// DEFERRED OPEN (perceived-latency fix): the op starts IMMEDIATELY and the pop-out is opened
// ONLY if the op outlives a grace period sized so the COMMON case (Shrink/Condense, and a Hide on
// a typical tournament doc) shows NO dialog at all — no open latency, no end-linger — while a
// genuinely slow op (huge docs) still gets a real bar with a working Cancel. A fast op that
// ERRORS/blocks opens the pop-out at completion, because the ribbon has no other channel to report
// a failure the user must act on. Only fast *successes* go silent, and those have a visible
// document effect anyway.
//
// React-free (the dialog PAGE is React; this driver is not) so it stays in the ribbon command
// bundle. `Office`/`window` are touched only inside functions, so importing this in Node (the
// manifest generator) is safe.
import type { ProgressInfo } from "../core/officeWordPort";
import type { CommandResult } from "../features/types";
import { logger } from "../core/debug";
import { appPageUrl } from "../core/appUrl";

const log = logger("progress");

/** How long an op may run before the pop-out is opened. Sub-grace ops never spawn a dialog.
 *  3500ms (up from 300) so a typical tournament-doc Hide (~2–3s engine, plus variance on slower
 *  machines) shows NO dialog at all — the document visibly changing is the completion signal,
 *  matching the felt speed of the zap macros debaters compare against. Sized ABOVE the typical
 *  band's upper edge on purpose: a grace inside the band (e.g. 2500) makes its upper half blink a
 *  dialog for under a second. Long ops (huge docs) still get the bar — on a 45s run Cancel goes
 *  live ~4s in, ~90% of the run. */
export const GRACE_OPEN_MS = 3500;
/** Auto-close delay after a SUCCESSFUL op that DID show a dialog. 300ms (down from 900, originally
 *  1400) — product decision 2026-06-10: snappier close wins over the guarantee that the `polite`
 *  aria-live "Done" finishes announcing before `dialog.close()` tears down its a11y tree (it may be
 *  cut short on some screen readers). Error/blocked outcomes are unaffected: those keep the dialog
 *  open until dismissed, so every announcement a user must act on still lands. */
export const AUTO_CLOSE_OK_MS = 300;
/** Minimum on-screen time for a dialog that DID open before a clean finish may close it. Without a
 *  floor, an op that finishes shortly after the grace fires produces an appear-then-vanish blink —
 *  the exact artifact deferred-open exists to kill, just moved to the (grace, grace+~2s) band. The
 *  close delay is max(AUTO_CLOSE_OK_MS, MIN_SHOW_MS − time-already-visible), so a dialog either
 *  never appears or reads as deliberate. Error/blocked keep-open paths ignore this entirely. */
export const MIN_SHOW_MS = 1200;
/** Hard close for a pop-out whose page never initialized, so a broken dialog can't linger. */
const WATCHDOG_MS = 4000;
/** Office DialogEventReceived code for a normal user dismiss (anything else is a real problem). */
const DIALOG_CLOSED_BY_USER = 12006;

/** Hard cap on the diagnostics text streamed to the pop-out: messageChild has a payload ceiling,
 *  and the TAIL of the bug report (where the failure is) is what matters. Keep the last N chars. */
const MAX_DIAGNOSTICS_CHARS = 6000;

export interface WithProgressOptions {
  /** Called when the user clicks Cancel in the pop-out (wire to controller.cancel()). */
  onCancel?: () => void;
  /** Run the op; call `forward(progress)` to stream ticks into the pop-out. Returns the result. */
  run: (forward: (p: ProgressInfo) => void) => Promise<CommandResult>;
  /**
   * Bounded ops that stream NO progress (e.g. Condense's range ops — a single read+write round-trip):
   * when true the pop-out NEVER opens for progress or a clean finish, so an instant op shows no
   * indeterminate "Working…" bar at all. It still opens at completion on an error/blocked outcome —
   * the ribbon's only channel to report a failure the user must act on (e.g. Track Changes).
   */
  quiet?: boolean;
  /**
   * Optional full diagnostics (e.g. `tracer.bugReport()`), surfaced in the pop-out ONLY when the op
   * ends in a keep-open outcome (error/blocked). A ribbon command runs in its OWN runtime, so the
   * task pane's diagnostics view can never see its log — the pop-out is the only place a user can
   * read WHY a ribbon Hide/Show/Apply failed. Read lazily at completion so it captures the whole op.
   */
  getDiagnostics?: () => string;
}

/** Messages the driver pushes to the dialog page (parent → child). */
type ToChild =
  | { kind: "progress"; phase: ProgressInfo["phase"]; done: number; total: number }
  | { kind: "done"; status: CommandResult["status"]; message?: string; diagnostics?: string };

/**
 * Run `opts.run` immediately and, only if it outlives GRACE_OPEN_MS, open a progress pop-out that
 * streams its ticks; auto-close on a clean finish, keep open (with diagnostics) on error/blocked.
 *
 * Async-open is the whole subtlety: `displayDialogAsync` resolves on a callback, so the op can
 * settle BEFORE the dialog handle exists. Every open-completion decision (close-if-settled-clean,
 * show-the-result, arm-the-watchdog, idempotency) therefore lives in `openAndWire` and reads LIVE
 * state on resolve — never in a code path that assumes a dialog is already present.
 */
export async function withProgressDialog(label: string, opts: WithProgressOptions): Promise<CommandResult> {
  let dialog: Office.Dialog | null = null; // null until (and unless) the pop-out is opened
  let opening = false; // displayDialogAsync is in flight — idempotency guard for openAndWire
  let childReady = false; // the page reported "ready" and can receive messages
  let latest: ProgressInfo | null = null; // most recent tick, replayed once the page is ready
  // Initialized to a real result (not null) so the return type is honest even on a path that
  // somehow skips the assignment below; the try/catch always overwrites it in practice.
  let final: CommandResult = { status: "error", message: "Operation did not complete." };
  let settled = false; // the op has finished and `final` holds its real result
  let donePushed = false; // the "done" message has been shown (idempotency guard)
  let openedAt = 0; // Date.now() when the dialog handle arrived — anchors the MIN_SHOW_MS floor
  let closed = false;
  let watchdog: number | null = null;
  let graceTimer: number | null = null;

  const clearWatchdog = (): void => {
    if (watchdog !== null) {
      window.clearTimeout(watchdog);
      watchdog = null;
    }
  };
  const clearGrace = (): void => {
    if (graceTimer !== null) {
      window.clearTimeout(graceTimer);
      graceTimer = null;
    }
  };

  const close = (): void => {
    if (closed) return;
    closed = true;
    clearWatchdog();
    clearGrace();
    if (!dialog) return; // never opened (fast path) — nothing to tear down
    try {
      dialog.close();
    } catch (e) {
      log.caught("progress dialog close failed (ignored)", e);
    }
  };

  // messageChild is DialogApi 1.2; absent on older hosts → streaming is a no-op (indeterminate UI).
  const send = (msg: ToChild): void => {
    if (!dialog) return;
    const mc = (dialog as unknown as { messageChild?: (m: string) => void }).messageChild;
    if (typeof mc !== "function") return;
    try {
      mc.call(dialog, JSON.stringify(msg));
    } catch (e) {
      log.caught("messageChild failed (ignored)", e);
    }
  };

  // Show the final result in an OPEN pop-out. Auto-close ONLY on a clean outcome; for error/blocked
  // keep it open so the user can read the (often actionable, e.g. Track-Changes) message and dismiss
  // it — a ribbon op has no other surface to report failure. Idempotent; assumes the page is ready.
  const showDone = (): void => {
    if (closed || donePushed || !dialog) return;
    donePushed = true;
    clearWatchdog();
    const keepOpen = final.status === "error" || final.status === "blocked";
    // Gather diagnostics ONLY for a keep-open outcome (the user is about to read them). Guarded so a
    // throwing provider can never break the done message itself; truncated to the failure-bearing
    // tail so it stays under the messageChild payload ceiling.
    let diagnostics: string | undefined;
    if (keepOpen && opts.getDiagnostics) {
      try {
        const full = opts.getDiagnostics();
        diagnostics =
          full.length > MAX_DIAGNOSTICS_CHARS
            ? `…(diagnostics truncated to last ${MAX_DIAGNOSTICS_CHARS} chars)\n${full.slice(-MAX_DIAGNOSTICS_CHARS)}`
            : full;
      } catch (e) {
        log.caught("progress getDiagnostics threw (ignored)", e);
      }
    }
    send({ kind: "done", status: final.status, message: final.message, diagnostics });
    // Clean finish: close after the linger, but never let total visible time dip under the
    // MIN_SHOW_MS floor — a dialog that appeared moments ago must not blink out (see MIN_SHOW_MS).
    if (!keepOpen) {
      const visibleFor = Date.now() - openedAt;
      window.setTimeout(close, Math.max(AUTO_CLOSE_OK_MS, MIN_SHOW_MS - visibleFor));
    }
    // else: the page renders a Close button (posts "close" → close()).
  };

  // Wire handlers onto a freshly opened dialog, then reconcile against state that may have changed
  // WHILE the dialog was opening. This is the single place all "the op already finished" logic lives.
  const wire = (d: Office.Dialog): void => {
    dialog = d;
    openedAt = Date.now(); // the OS window is visible from here — the floor measures from this point
    d.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      const message = (arg as { message?: string }).message;
      if (message === "ready") {
        // The page finished loading and can receive now — replay the current state.
        childReady = true;
        if (settled) showDone();
        else if (latest) send({ kind: "progress", phase: latest.phase, done: latest.done, total: latest.total });
      } else if (message === "cancel") {
        try {
          opts.onCancel?.();
        } catch (e) {
          log.caught("progress onCancel threw (ignored)", e);
        }
      } else if (message === "close") {
        close();
      }
    });
    d.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      // Fires for a user dismiss (12006) AND for problems like a page that failed to load (12002).
      // Log the non-dismiss codes so a broken progress.html deployment is diagnosable, then mark
      // the dialog gone so we never try to close one that already closed itself.
      const code = (arg as { error?: number }).error;
      if (code && code !== DIALOG_CLOSED_BY_USER) {
        log.warn("progress dialog closed unexpectedly (not a user dismiss)", { error: code });
      }
      closed = true;
      clearWatchdog();
      clearGrace();
    });

    // Reconcile: the op may have settled while displayDialogAsync was resolving.
    if (settled) {
      const keepOpen = final.status === "error" || final.status === "blocked";
      if (keepOpen) {
        // Surface the error once the page reports ready; arm the watchdog NOW so a page that never
        // initializes (broken deploy) still tears down instead of lingering forever.
        if (watchdog === null) watchdog = window.setTimeout(close, WATCHDOG_MS);
        if (childReady) showDone();
      } else {
        // Settled CLEAN while opening: we no longer want this dialog. Close it immediately rather
        // than flash a needless "Done" — prevents both the flicker and a leaked open dialog.
        close();
      }
    }
    // else: op still running — progress streams to the page once it reports ready.
  };

  // Open the pop-out and wire it. Idempotent and async-safe: a second caller (grace timer vs.
  // completion) is dropped, and a dialog that resolves after we've already closed is closed at once.
  const openAndWire = (): void => {
    if (dialog || opening || closed) return;
    opening = true;
    void openDialog(label).then((d) => {
      opening = false;
      if (!d) return; // open failed (no host / popup blocked / dialog already open) — best-effort
      if (closed) {
        try {
          d.close();
        } catch (e) {
          log.caught("progress dialog close failed (ignored)", e);
        }
        return;
      }
      wire(d);
    });
  };

  // The forwarder handed to the op. It always buffers `latest`; it only streams when a ready page
  // exists, so ticks fired before the (possibly deferred) dialog opens are simply replayed later.
  const forward = (p: ProgressInfo): void => {
    latest = p;
    if (dialog && childReady) send({ kind: "progress", phase: p.phase, done: p.done, total: p.total });
  };

  // Start the op immediately — the dialog is the deferred part, never a precondition for running.
  const runPromise: Promise<CommandResult> = (async () => {
    try {
      return await opts.run(forward);
    } catch (e) {
      return { status: "error", message: String((e as Error)?.message ?? e) };
    }
  })();

  // Open the pop-out only if the op is still running after the grace period. Re-check `settled`
  // inside the callback: clearGrace() can lose a same-tick race against settle, so the flag is truth.
  // `quiet` ops skip this entirely — they can ONLY open via the completion error/blocked path below.
  if (!opts.quiet) {
    graceTimer = window.setTimeout(() => {
      graceTimer = null;
      if (!settled) openAndWire();
    }, GRACE_OPEN_MS);
  }

  final = await runPromise;
  settled = true;
  clearGrace();

  if (dialog) {
    // Slow op: the pop-out is already open. Drive it to its end state (or arm the watchdog if the
    // page never reported ready and so can't display anything).
    if (childReady) showDone();
    else if (watchdog === null) watchdog = window.setTimeout(close, WATCHDOG_MS);
  } else if (opening) {
    // The dialog is mid-open (grace just fired); `wire`'s settled-reconcile handles it on resolve.
  } else {
    // Fast path: no dialog was opened. Surface an error/blocked outcome (the ribbon's only failure
    // channel); a clean outcome stays silent — the document already shows the result.
    if (final.status === "error" || final.status === "blocked") openAndWire();
  }
  return final;
}

/** Open progress.html as a small dialog; resolves null (logged) when it can't open. */
function openDialog(label: string): Promise<Office.Dialog | null> {
  return new Promise((resolve) => {
    if (typeof Office === "undefined" || !Office.context?.ui?.displayDialogAsync) {
      resolve(null);
      return;
    }
    try {
      // Resolve progress.html RELATIVE to the current page so it loads on a project-Pages subpath
      // (location.origin would drop the `/rostrum` segment → 404). See core/appUrl.ts. Kept INSIDE
      // the try so a (theoretical) missing-`window` throw becomes the graceful "run without pop-out"
      // fallback rather than an unhandled rejection.
      const url = appPageUrl("progress.html", encodeURIComponent(label));
      Office.context.ui.displayDialogAsync(url, { height: 18, width: 28, displayInIframe: false }, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve(res.value);
        } else {
          log.warn("progress dialog failed to open (running without it)", { error: res.error?.message });
          resolve(null);
        }
      });
    } catch (e) {
      log.caught("displayDialogAsync threw (running without progress dialog)", e);
      resolve(null);
    }
  });
}
