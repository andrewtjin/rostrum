// Lifecycle/race tests for the ribbon progress driver (src/progress/host.ts). The driver carries
// every stateful invariant of the pop-out — open, stream, keep-open-on-error, watchdog, no
// double-close — and these are exactly the ordering-dependent bugs a human can't eyeball. We drive
// it with a fake Office.Dialog + fake timers in Node (no host needed). Mirrors the fake-injection
// discipline used elsewhere in the suite.
import { withProgressDialog } from "../src/progress/host";
import type { CommandResult } from "../src/features/types";

type Handler = (arg: unknown) => void;

/** A fake Office surface: a capturable dialog + the event handlers the driver registers. */
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
  const Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    EventType: {
      DialogMessageReceived: "DialogMessageReceived",
      DialogEventReceived: "DialogEventReceived",
    },
    context: {
      ui: {
        displayDialogAsync(_url: string, _opts: unknown, cb: (res: unknown) => void): void {
          if (failOpen) cb({ status: "failed", error: { message: "blocked" } });
          else cb({ status: "succeeded", value: dialog });
        },
      },
    },
  };
  return {
    Office,
    dialog,
    setFailOpen: (v: boolean) => {
      failOpen = v;
    },
    fire: (message: string) => handlers["DialogMessageReceived"]?.({ message }),
    fireDialogEvent: (code?: number) => handlers["DialogEventReceived"]?.({ error: code }),
  };
}

describe("withProgressDialog", () => {
  let fake: ReturnType<typeof makeFakeOffice>;

  beforeEach(() => {
    jest.useFakeTimers();
    fake = makeFakeOffice();
    (globalThis as Record<string, unknown>).Office = fake.Office;
    (globalThis as Record<string, unknown>).window = {
      setTimeout: (fn: () => void, ms?: number) => setTimeout(fn, ms),
      clearTimeout: (id: unknown) => clearTimeout(id as ReturnType<typeof setTimeout>),
      // `href` (not just `origin`) because the pop-out URL is now resolved RELATIVE to the
      // current page (core/appUrl.ts) so it survives a project-Pages subpath. The ribbon
      // command runtime that opens the pop-out is served from commands.html.
      location: { origin: "https://localhost:3000", href: "https://localhost:3000/commands.html" },
    };
  });

  afterEach(() => {
    jest.useRealTimers();
    delete (globalThis as Record<string, unknown>).Office;
    delete (globalThis as Record<string, unknown>).window;
  });

  it("closes the dialog exactly once after a successful op", async () => {
    const res = await withProgressDialog("Hide", { run: async () => ({ status: "ok" }) });
    fake.fire("ready"); // page reports ready after the op settled → showDone schedules auto-close
    jest.runAllTimers();
    expect(res).toEqual({ status: "ok" });
    expect(fake.dialog.closeCount).toBe(1);
    expect(fake.dialog.sent).toContainEqual({ kind: "done", status: "ok" });
  });

  it("still closes when the page never reports ready (watchdog tears down a dead pop-out)", async () => {
    await withProgressDialog("Hide", { run: async () => ({ status: "ok" }) });
    jest.runAllTimers(); // only the watchdog is armed
    expect(fake.dialog.closeCount).toBe(1);
  });

  it("does NOT re-close a dialog the user already dismissed", async () => {
    await withProgressDialog("Hide", { run: async () => ({ status: "ok" }) });
    fake.fireDialogEvent(12006); // user closed it (the normal dismiss code)
    jest.runAllTimers();
    expect(fake.dialog.closeCount).toBe(0);
  });

  it("keeps the pop-out open on error (no auto-close) and returns the error result", async () => {
    const res = await withProgressDialog("Hide", { run: async () => ({ status: "error", message: "boom" }) });
    fake.fire("ready");
    jest.runAllTimers();
    expect(res).toEqual({ status: "error", message: "boom" });
    expect(fake.dialog.closeCount).toBe(0);
    expect(fake.dialog.sent).toContainEqual({ kind: "done", status: "error", message: "boom" });
  });

  it("keeps the pop-out open on a blocked result (e.g. Track Changes) so its instruction is read", async () => {
    const res = await withProgressDialog("Hide", {
      run: async () => ({ status: "blocked", message: "Track Changes is on." }),
    });
    fake.fire("ready");
    jest.runAllTimers();
    expect(res.status).toBe("blocked");
    expect(fake.dialog.closeCount).toBe(0);
  });

  it("converts a thrown op into an error result and keeps the pop-out open", async () => {
    const res = await withProgressDialog("Hide", {
      run: async () => {
        throw new Error("kaboom");
      },
    });
    fake.fire("ready");
    jest.runAllTimers();
    expect(res).toEqual({ status: "error", message: "kaboom" });
    expect(fake.dialog.closeCount).toBe(0);
  });

  it("runs the op with a no-op forwarder when the dialog can't open (best-effort)", async () => {
    fake.setFailOpen(true);
    const res = await withProgressDialog("Hide", { run: async () => ({ status: "ok" }) });
    expect(res).toEqual({ status: "ok" });
    expect(fake.dialog.closeCount).toBe(0);
  });

  it("routes a child cancel message to onCancel", async () => {
    const onCancel = jest.fn();
    await withProgressDialog("Hide", { onCancel, run: async () => ({ status: "ok" }) });
    fake.fire("cancel");
    expect(onCancel).toHaveBeenCalledTimes(1);
    jest.runAllTimers();
  });

  it("replays the latest progress tick once the page reports ready (no lost first tick)", async () => {
    let resolveRun!: (r: CommandResult) => void;
    const p = withProgressDialog("Hide", {
      run: (forward) => {
        forward({ phase: "read", done: 3, total: 9 }); // tick BEFORE the page is ready
        return new Promise<CommandResult>((r) => {
          resolveRun = r;
        });
      },
    });
    await Promise.resolve();
    await Promise.resolve(); // let openDialog resolve, handlers register, and run() fire its tick
    fake.fire("ready"); // op not settled yet, latest tick present → replay it
    resolveRun({ status: "ok" });
    await p;
    jest.runAllTimers();
    expect(fake.dialog.sent).toContainEqual({ kind: "progress", phase: "read", done: 3, total: 9 });
  });

  it("attaches getDiagnostics output to the done message on a kept-open failure (ribbon error visibility)", async () => {
    const getDiagnostics = jest.fn(() => "=== Rostrum diagnostics ===\nadapter commit | ✗ insertOoxml");
    const res = await withProgressDialog("Hide", {
      getDiagnostics,
      run: async () => ({ status: "error", message: "boom" }),
    });
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

  it("does NOT gather diagnostics for a successful op (clean outcomes auto-close, no dump)", async () => {
    const getDiagnostics = jest.fn(() => "SHOULD-NOT-APPEAR");
    await withProgressDialog("Hide", { getDiagnostics, run: async () => ({ status: "ok" }) });
    fake.fire("ready");
    jest.runAllTimers();
    expect(getDiagnostics).not.toHaveBeenCalled();
    expect(fake.dialog.sent).toContainEqual({ kind: "done", status: "ok" });
  });

  it("truncates oversized diagnostics to the failure-bearing tail (messageChild payload cap)", async () => {
    // >6000 chars: the START marker must be dropped, the END marker (where the failure lives) kept.
    const getDiagnostics = (): string => "STARTMARK" + "H".repeat(6000) + "ENDMARK";
    await withProgressDialog("Hide", { getDiagnostics, run: async () => ({ status: "error", message: "x" }) });
    fake.fire("ready");
    jest.runAllTimers();
    const done = fake.dialog.sent.find((m) => m.kind === "done") as { diagnostics?: string };
    expect(done.diagnostics).toContain("ENDMARK");
    expect(done.diagnostics).not.toContain("STARTMARK");
    expect(done.diagnostics).toContain("truncated");
  });
});
