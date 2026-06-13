// Lifecycle/race tests for the ribbon progress driver (src/progress/host.ts). The driver carries
// every stateful invariant of the pop-out — and with DEFERRED OPEN those invariants got subtler:
// the op runs first, the dialog opens late (or never), and `displayDialogAsync` resolving AFTER the
// op settled is a first-class case. These are exactly the ordering bugs a human can't eyeball, so we
// drive it with a fake Office.Dialog + fake timers in Node (no host). Mirrors the fake-injection
// discipline used elsewhere in the suite.
import { withProgressDialog, GRACE_OPEN_MS, AUTO_CLOSE_OK_MS, MIN_SHOW_MS } from "../src/progress/host";
import type { CommandResult } from "../src/features/types";
import type { ProgressInfo } from "../src/core/officeWordPort";

type Handler = (arg: unknown) => void;

/** Let queued microtasks (the `.then` after displayDialogAsync resolves, the run() IIFE) drain. */
const flush = async (): Promise<void> => {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
};

/** A fake Office surface: a capturable dialog + the event handlers the driver registers. `openCount`
 *  proves the deferred-open contract (instant ops must NOT open a dialog). `deferOpen` holds the
 *  open callback so a test can interleave "op settles WHILE the dialog is still opening". */
function makeFakeOffice() {
  const handlers: Record<string, Handler> = {};
  const dialog = {
    closeCount: 0,
    sent: [] as Array<Record<string, unknown>>,
    close(): void {
      this.closeCount++;
    },
    messageChild(m: string): void {
      this.sent.push(JSON.parse(m));
    },
    addEventHandler(type: string, h: Handler): void {
      handlers[type] = h;
    },
  };
  let failOpen = false;
  let deferOpen = false;
  let pendingOpen: ((res: unknown) => void) | null = null;
  let openCount = 0;
  const Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    EventType: {
      DialogMessageReceived: "DialogMessageReceived",
      DialogEventReceived: "DialogEventReceived",
    },
    context: {
      ui: {
        displayDialogAsync(_url: string, _opts: unknown, cb: (res: unknown) => void): void {
          openCount++;
          const deliver = (): void =>
            failOpen ? cb({ status: "failed", error: { message: "blocked" } }) : cb({ status: "succeeded", value: dialog });
          if (deferOpen) pendingOpen = deliver;
          else deliver();
        },
      },
    },
  };
  return {
    Office,
    dialog,
    openCount: () => openCount,
    setFailOpen: (v: boolean) => {
      failOpen = v;
    },
    setDeferOpen: (v: boolean) => {
      deferOpen = v;
    },
    resolveOpen: () => {
      const p = pendingOpen;
      pendingOpen = null;
      p?.(undefined);
    },
    fire: (message: string) => handlers["DialogMessageReceived"]?.({ message }),
    fireDialogEvent: (code?: number) => handlers["DialogEventReceived"]?.({ error: code }),
  };
}

/** A run() that stays pending until the test resolves it — the only way to force an op PAST the
 *  grace period so a dialog actually opens. `forward` proxies to whatever forwarder the driver
 *  handed in, so a test can stream a tick before resolving. */
function pendingRun() {
  let resolve!: (r: CommandResult) => void;
  let captured: ((p: ProgressInfo) => void) | undefined;
  const run = (fwd: (p: ProgressInfo) => void): Promise<CommandResult> => {
    captured = fwd;
    return new Promise<CommandResult>((r) => {
      resolve = r;
    });
  };
  return {
    run,
    resolve: (r: CommandResult): void => resolve(r),
    forward: (p: ProgressInfo): void => captured?.(p),
  };
}

describe("withProgressDialog (deferred open)", () => {
  let fake: ReturnType<typeof makeFakeOffice>;

  beforeEach(() => {
    jest.useFakeTimers();
    fake = makeFakeOffice();
    (globalThis as Record<string, unknown>).Office = fake.Office;
    (globalThis as Record<string, unknown>).window = {
      setTimeout: (fn: () => void, ms?: number) => setTimeout(fn, ms),
      clearTimeout: (id: unknown) => clearTimeout(id as ReturnType<typeof setTimeout>),
      location: { origin: "https://localhost:3000", href: "https://localhost:3000/taskpane.html" },
    };
  });

  afterEach(() => {
    jest.useRealTimers();
    delete (globalThis as Record<string, unknown>).Office;
    delete (globalThis as Record<string, unknown>).window;
  });

  // ── Fast path: the whole point of the change ──────────────────────────────────────────────────
  it("a fast clean op opens NO dialog at all (zero perceived bloat)", async () => {
    const res = await withProgressDialog("Shrink", { run: async () => ({ status: "ok" }) });
    jest.runAllTimers(); // the grace timer would fire here — it must be a no-op (op already settled)
    expect(res).toEqual({ status: "ok" });
    expect(fake.openCount()).toBe(0);
    expect(fake.dialog.closeCount).toBe(0);
    expect(fake.dialog.sent).toHaveLength(0);
  });

  it("a fast clean op runs to completion even if its grace timer never gets a chance to fire", async () => {
    // No timer advance at all — proves the op does not DEPEND on the dialog/grace machinery.
    const res = await withProgressDialog("Condense", { run: async () => ({ status: "ok", message: "merged 3" }) });
    expect(res).toEqual({ status: "ok", message: "merged 3" });
    expect(fake.openCount()).toBe(0);
  });

  // ── Fast ERROR/BLOCKED: must still get a surface (ribbon's only failure channel) ────────────────
  it("a fast ERROR op opens the pop-out at completion and keeps it open with the message", async () => {
    const res = await withProgressDialog("Hide", { run: async () => ({ status: "error", message: "boom" }) });
    await flush(); // let the completion-triggered open resolve + wire
    fake.fire("ready"); // page reports ready → showDone runs now (it was deferred until ready)
    jest.runAllTimers();
    expect(res).toEqual({ status: "error", message: "boom" });
    expect(fake.openCount()).toBe(1);
    expect(fake.dialog.closeCount).toBe(0); // kept open so the user can read + dismiss
    expect(fake.dialog.sent).toContainEqual({ kind: "done", status: "error", message: "boom" });
  });

  it("a fast BLOCKED op (e.g. Track Changes) opens at completion and stays open", async () => {
    const res = await withProgressDialog("Hide", { run: async () => ({ status: "blocked", message: "Track Changes is on." }) });
    await flush();
    fake.fire("ready");
    jest.runAllTimers();
    expect(res.status).toBe("blocked");
    expect(fake.openCount()).toBe(1);
    expect(fake.dialog.closeCount).toBe(0);
  });

  it("converts a thrown op into an error result and opens the pop-out to surface it", async () => {
    const res = await withProgressDialog("Hide", {
      run: async () => {
        throw new Error("kaboom");
      },
    });
    await flush();
    fake.fire("ready");
    jest.runAllTimers();
    expect(res).toEqual({ status: "error", message: "kaboom" });
    expect(fake.openCount()).toBe(1);
    expect(fake.dialog.closeCount).toBe(0);
  });

  // ── quiet ops (Condense range ops): NEVER open on success, even when slow ───────────────────────
  it("a quiet op never opens the pop-out on a clean finish, even past the grace period", async () => {
    const op = pendingRun();
    const p = withProgressDialog("Shrink", { quiet: true, run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS * 4); // well past grace — a non-quiet op would have opened
    await flush();
    op.resolve({ status: "ok" });
    const res = await p;
    jest.runAllTimers();
    expect(res).toEqual({ status: "ok" });
    expect(fake.openCount()).toBe(0); // the whole point: no "Working…" bar on an instant op
  });

  it("a quiet op STILL opens the pop-out at completion on a blocked outcome (Track Changes surfaces)", async () => {
    const res = await withProgressDialog("Shrink", {
      quiet: true,
      run: async () => ({ status: "blocked", message: "Track Changes is on." }),
    });
    await flush();
    fake.fire("ready");
    jest.runAllTimers();
    expect(res.status).toBe("blocked");
    expect(fake.openCount()).toBe(1); // failure must still reach the user
    expect(fake.dialog.closeCount).toBe(0);
  });

  // ── Slow path: dialog appears after the grace, streams, closes cleanly ──────────────────────────
  it("a slow op opens the pop-out after the grace period and closes it once after Done", async () => {
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS); // grace fires → dialog opens
    await flush(); // wire the freshly opened dialog
    fake.fire("ready");
    op.resolve({ status: "ok" });
    await p;
    // The dialog JUST appeared, so the close honors the MIN_SHOW_MS floor, not the bare linger.
    jest.advanceTimersByTime(MIN_SHOW_MS);
    expect(fake.openCount()).toBe(1);
    expect(fake.dialog.closeCount).toBe(1);
    expect(fake.dialog.sent).toContainEqual({ kind: "done", status: "ok" });
  });

  // ── The 2026-06-10 snappiness decision, made load-bearing ───────────────────────────────────────
  // The whole suite asserts RELATIVE to these constants, so a silent revert (or a fat-fingered
  // 25000) would pass every other test. Pinning the values turns the approved product tuning into
  // a contract: grace ABOVE the typical Hide band (~2-3s engine) so common ops show no dialog;
  // short success linger; a min-show floor so a dialog that did appear never blinks out.
  it("pins the approved timing constants (product decision 2026-06-10)", () => {
    expect(GRACE_OPEN_MS).toBe(3500);
    expect(AUTO_CLOSE_OK_MS).toBe(300);
    expect(MIN_SHOW_MS).toBe(1200);
  });

  it("a dialog that appears just before the op finishes stays up for MIN_SHOW_MS (no blink)", async () => {
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS);
    await flush();
    fake.fire("ready");
    op.resolve({ status: "ok" }); // finishes moments after the dialog appeared
    await p;
    jest.advanceTimersByTime(AUTO_CLOSE_OK_MS); // the bare linger alone must NOT close it…
    expect(fake.dialog.closeCount).toBe(0);
    jest.advanceTimersByTime(MIN_SHOW_MS - AUTO_CLOSE_OK_MS); // …the floor does
    expect(fake.dialog.closeCount).toBe(1);
  });

  it("a dialog already visible past the floor closes after just the short linger", async () => {
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS);
    await flush();
    fake.fire("ready");
    jest.advanceTimersByTime(MIN_SHOW_MS * 2); // dialog has been up well past the floor
    op.resolve({ status: "ok" });
    await p;
    jest.advanceTimersByTime(AUTO_CLOSE_OK_MS); // floor already satisfied → bare linger closes
    expect(fake.dialog.closeCount).toBe(1);
  });

  it("replays the latest progress tick once the page reports ready (no lost first tick)", async () => {
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    op.forward({ phase: "read", done: 3, total: 9 }); // tick fired while the dialog is still deferred
    jest.advanceTimersByTime(GRACE_OPEN_MS);
    await flush();
    fake.fire("ready"); // op not settled, latest tick buffered → replay it
    op.resolve({ status: "ok" });
    await p;
    jest.runAllTimers();
    expect(fake.dialog.sent).toContainEqual({ kind: "progress", phase: "read", done: 3, total: 9 });
  });

  it("routes a child cancel message to onCancel (slow path, dialog present)", async () => {
    const onCancel = jest.fn();
    const op = pendingRun();
    const p = withProgressDialog("Hide", { onCancel, run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS);
    await flush();
    fake.fire("cancel");
    expect(onCancel).toHaveBeenCalledTimes(1);
    op.resolve({ status: "cancelled" });
    await p;
    jest.runAllTimers();
  });

  // ── The async-open races the reviewers flagged ──────────────────────────────────────────────────
  it("settled-CLEAN while the dialog is still opening → closes it with no 'Done' flash, no leak", async () => {
    fake.setDeferOpen(true);
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS); // grace fires → displayDialogAsync called but DEFERRED
    await flush();
    op.resolve({ status: "ok" }); // op settles clean BEFORE the open resolves
    await p;
    fake.resolveOpen(); // now the dialog finishes opening — must be closed immediately
    await flush();
    expect(fake.openCount()).toBe(1);
    expect(fake.dialog.closeCount).toBe(1);
    expect(fake.dialog.sent.find((m) => m.kind === "done")).toBeUndefined(); // no needless flash
  });

  it("does not double-open: a grace timer that fires after the op settled is a no-op", async () => {
    const res = await withProgressDialog("Hide", { run: async () => ({ status: "ok" }) });
    jest.advanceTimersByTime(GRACE_OPEN_MS * 5); // grace + any other timer — still nothing opens
    await flush();
    expect(res).toEqual({ status: "ok" });
    expect(fake.openCount()).toBe(0);
  });

  // ── Watchdog: a completion-opened error dialog whose page never initializes still tears down ────
  it("watchdog closes a completion-opened error dialog whose page never reports ready", async () => {
    const res = await withProgressDialog("Hide", { run: async () => ({ status: "error", message: "boom" }) });
    await flush(); // open + wire (which arms the watchdog because the outcome is keep-open)
    jest.runAllTimers(); // page never fires "ready" → only the watchdog is armed
    expect(res.status).toBe("error");
    expect(fake.openCount()).toBe(1);
    expect(fake.dialog.closeCount).toBe(1);
  });

  it("watchdog tears down a slow-op pop-out whose page never reports ready", async () => {
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS);
    await flush();
    op.resolve({ status: "ok" }); // settles, but the page never reported ready
    await p;
    jest.runAllTimers(); // watchdog
    expect(fake.dialog.closeCount).toBe(1);
  });

  it("does NOT re-close a dialog the user already dismissed", async () => {
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS);
    await flush();
    fake.fireDialogEvent(12006); // user closed it (the normal dismiss code)
    op.resolve({ status: "ok" });
    await p;
    jest.runAllTimers();
    expect(fake.dialog.closeCount).toBe(0);
  });

  // ── Best-effort: open can fail; the op must still complete ──────────────────────────────────────
  it("runs the op to completion when the (deferred) dialog can't open", async () => {
    fake.setFailOpen(true);
    const op = pendingRun();
    const p = withProgressDialog("Hide", { run: op.run });
    await flush();
    jest.advanceTimersByTime(GRACE_OPEN_MS); // grace → open attempt → fails (null)
    await flush();
    op.resolve({ status: "ok" });
    const res = await p;
    jest.runAllTimers();
    expect(res).toEqual({ status: "ok" });
    expect(fake.openCount()).toBe(1); // attempted…
    expect(fake.dialog.closeCount).toBe(0); // …but nothing to close
  });

  // ── Diagnostics: gathered only for a kept-open failure, truncated to the tail ───────────────────
  it("attaches getDiagnostics output to the done message on a kept-open failure", async () => {
    const getDiagnostics = jest.fn(() => "=== Rostrum diagnostics ===\nadapter commit | ✗ insertOoxml");
    const res = await withProgressDialog("Hide", { getDiagnostics, run: async () => ({ status: "error", message: "boom" }) });
    await flush();
    fake.fire("ready");
    jest.runAllTimers();
    expect(res.status).toBe("error");
    expect(getDiagnostics).toHaveBeenCalledTimes(1);
    expect(fake.dialog.sent).toContainEqual({
      kind: "done",
      status: "error",
      message: "boom",
      diagnostics: "=== Rostrum diagnostics ===\nadapter commit | ✗ insertOoxml",
    });
  });

  it("does NOT gather diagnostics for a successful op (no dialog, nothing to dump)", async () => {
    const getDiagnostics = jest.fn(() => "SHOULD-NOT-APPEAR");
    await withProgressDialog("Hide", { getDiagnostics, run: async () => ({ status: "ok" }) });
    jest.runAllTimers();
    expect(getDiagnostics).not.toHaveBeenCalled();
    expect(fake.openCount()).toBe(0);
  });

  it("truncates oversized diagnostics to the failure-bearing tail (messageChild payload cap)", async () => {
    const getDiagnostics = (): string => "STARTMARK" + "H".repeat(6000) + "ENDMARK";
    await withProgressDialog("Hide", { getDiagnostics, run: async () => ({ status: "error", message: "x" }) });
    await flush();
    fake.fire("ready");
    jest.runAllTimers();
    const done = fake.dialog.sent.find((m) => m.kind === "done") as { diagnostics?: string };
    expect(done.diagnostics).toContain("ENDMARK");
    expect(done.diagnostics).not.toContain("STARTMARK");
    expect(done.diagnostics).toContain("truncated");
  });
});
