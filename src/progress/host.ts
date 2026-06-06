// Ribbon progress pop-out — the headless driver. A ribbon command runs with NO task pane (to
// keep the document's reading space), so this is how a long Hide/Show/Apply op gives live
// feedback: a tiny Office dialog that appears only while the op runs and closes itself on
// completion. Word add-ins can't draw a moving bar on the ribbon itself, so a transient dialog
// is the real option — it floats over Word and never docks, so it costs zero persistent space.
//
// React-free (the dialog PAGE is React; this driver is not) so it stays in the ribbon command
// bundle. `Office`/`window` are touched only inside functions, so importing this in Node (the
// manifest generator) is safe.
import type { ProgressInfo } from "../core/officeWordPort";
import type { CommandResult } from "../features/types";
import { logger } from "../core/debug";

const log = logger("progress");

/** Auto-close delay after a SUCCESSFUL op — long enough for the page's aria-live to announce "Done". */
const AUTO_CLOSE_OK_MS = 1400;
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
 * Open a progress pop-out, run `opts.run` while forwarding progress into it, then auto-close.
 * Best-effort: if the dialog can't open (no host, popup blocked, a dialog already open), the op
 * still runs with a no-op forwarder. Live numbers need parent→child messaging (DialogApi 1.2);
 * without it the dialog shows an indeterminate "Working…" and is still closed on completion.
 */
export async function withProgressDialog(label: string, opts: WithProgressOptions): Promise<CommandResult> {
  const dialog = await openDialog(label);
  if (!dialog) {
    // No pop-out — run the op anyway; feedback just isn't shown for this invocation.
    return opts.run(() => undefined);
  }

  let childReady = false;
  let latest: ProgressInfo | null = null;
  // Initialized to a real result (not null) so the return type is honest even on a path that
  // somehow skips the assignment below; the try/catch always overwrites it in practice.
  let final: CommandResult = { status: "error", message: "Operation did not complete." };
  let settled = false; // the op has finished and `final` holds its real result
  let donePushed = false; // the "done" message has been shown (idempotency guard)
  let closed = false;
  let watchdog: number | null = null;

  const clearWatchdog = (): void => {
    if (watchdog !== null) {
      window.clearTimeout(watchdog);
      watchdog = null;
    }
  };

  const close = (): void => {
    if (closed) return;
    closed = true;
    clearWatchdog();
    try {
      dialog.close();
    } catch (e) {
      log.caught("progress dialog close failed (ignored)", e);
    }
  };

  // messageChild is DialogApi 1.2; absent on older hosts → streaming is a no-op (indeterminate UI).
  const send = (msg: ToChild): void => {
    const mc = (dialog as unknown as { messageChild?: (m: string) => void }).messageChild;
    if (typeof mc !== "function") return;
    try {
      mc.call(dialog, JSON.stringify(msg));
    } catch (e) {
      log.caught("messageChild failed (ignored)", e);
    }
  };

  // Show the final result. Auto-close ONLY on a clean outcome; for error/blocked keep the pop-out
  // open so the user can read the (often actionable, e.g. Track-Changes) message and dismiss it —
  // a ribbon op has no other surface to report failure. Idempotent.
  const showDone = (): void => {
    if (closed || donePushed) return;
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
    if (!keepOpen) window.setTimeout(close, AUTO_CLOSE_OK_MS);
    // else: the page renders a Close button (posts "close" → close()).
  };

  // Reconcile after the op settles: show the result if the page is ready, else arm a watchdog so a
  // page that never initialized still gets torn down (it has nothing to display anyway).
  const settle = (): void => {
    if (closed || !settled) return;
    if (childReady) showDone();
    else if (watchdog === null) watchdog = window.setTimeout(close, WATCHDOG_MS);
  };

  dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
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
  dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
    // Fires for a user dismiss (12006) AND for problems like a page that failed to load (12002).
    // Log the non-dismiss codes so a broken progress.html deployment is diagnosable, then mark
    // the dialog gone so we never try to close one that already closed itself.
    const code = (arg as { error?: number }).error;
    if (code && code !== DIALOG_CLOSED_BY_USER) {
      log.warn("progress dialog closed unexpectedly (not a user dismiss)", { error: code });
    }
    closed = true;
    clearWatchdog();
  });

  const forward = (p: ProgressInfo): void => {
    latest = p;
    if (childReady) send({ kind: "progress", phase: p.phase, done: p.done, total: p.total });
  };

  try {
    final = await opts.run(forward);
  } catch (e) {
    final = { status: "error", message: String((e as Error)?.message ?? e) };
  }
  settled = true;
  settle();
  return final;
}

/** Open progress.html as a small dialog; resolves null (logged) when it can't open. */
function openDialog(label: string): Promise<Office.Dialog | null> {
  return new Promise((resolve) => {
    if (typeof Office === "undefined" || !Office.context?.ui?.displayDialogAsync) {
      resolve(null);
      return;
    }
    const url = `${window.location.origin}/progress.html#${encodeURIComponent(label)}`;
    try {
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
