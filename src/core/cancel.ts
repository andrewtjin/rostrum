// Cooperative cancellation + macrotask pacing for the engine's long pure-JS loops.
//
// WHY THIS FILE: the pure whole-body Hide path (avenue ⑦) runs ONE host read sync, then a
// long stretch of synchronous JS — package assembly, then per-paragraph classification —
// before the commit sync. On the add-in's single-threaded runtime that stretch blocks both
// paint and input: progress ticks coalesce unpainted and a Cancel click queues unprocessed,
// so a bare `isCancelled()` poll inside the loop can never observe a click made after the
// read sync resolved (the click handler cannot run while the loop does). A `Pacer` fixes
// both with one primitive: a cheap per-iteration `tick()` that (a) throws `CancelledError`
// when cancellation was requested and (b) yields ONE macrotask whenever the time budget
// elapses, letting queued input/paint/postMessage run between work slices.
//
// `CancelToken`/`CancelledError` live here (not in officeWordPort.ts, their old home) so the
// host-agnostic engine (invisibility.ts) can pace/cancel without importing the Office
// adapter; the adapter re-exports both, so every existing import keeps compiling.

/** Cooperative cancellation, checked between read chunks (never mid-commit). */
export interface CancelToken {
  isCancelled(): boolean;
}

/** Thrown when the caller cancels during the (pre-write, always-safe) read phase. */
export class CancelledError extends Error {
  constructor() {
    super("Rostrum operation cancelled before any changes were written.");
    this.name = "CancelledError";
    Object.setPrototypeOf(this, CancelledError.prototype);
  }
}

/**
 * Paces a long synchronous loop: `await pacer.tick()` once per iteration. Almost every call
 * is just a clock check that resolves without leaving the microtask queue; only when the
 * budget has elapsed does it await one MACROTASK, which is what actually lets the host
 * runtime paint and deliver the Cancel click.
 */
export interface Pacer {
  /** Throws `CancelledError` if cancelled; yields one macrotask when the budget elapsed. */
  tick(): Promise<void>;
}

export interface PacerOptions {
  /** Polled on every tick AND re-checked after each yield (a Cancel click lands DURING a yield). */
  cancel?: CancelToken;
  /**
   * How long a loop may run between yields. The 50ms default keeps the pane visibly alive
   * (~20 paints/s worst case) while bounding overhead to one (possibly ~4ms-clamped) timer
   * per ≥50ms of real work — ≤ ~8% wall-clock. 0 = yield on every tick (tests);
   * Infinity = never yield (a pure cancel poll with the old bare-token semantics).
   */
  budgetMs?: number;
  /** The macrotask yield. Injectable so tests can count/instrument yields deterministically. */
  yieldFn?: () => Promise<void>;
  /** Clock used to meter the budget; injectable for deterministic tests. */
  now?: () => number;
}

/** One macrotask (`setTimeout 0`): lets the event loop paint and deliver queued input. */
const defaultYield = (): Promise<void> => new Promise<void>((resolve) => setTimeout(resolve, 0));

/** Build a `Pacer`. The defaults suit the live task pane; see `PacerOptions` for the knobs. */
export function createPacer(options: PacerOptions = {}): Pacer {
  const { cancel, budgetMs = 50, yieldFn = defaultYield, now = Date.now } = options;
  // The budget window opens at creation and re-arms after every yield. A stale window (pacer
  // created long before the op starts) costs at most one extra yield on the first tick.
  let sliceStart = now();
  return {
    async tick(): Promise<void> {
      // Cancel beats yielding: an already-cancelled op must not burn another work slice.
      if (cancel?.isCancelled()) throw new CancelledError();
      if (now() - sliceStart >= budgetMs) {
        await yieldFn();
        sliceStart = now();
        // Re-check AFTER the yield — the load-bearing line: the Cancel click is processed
        // during the yielded macrotask, so this is where it is finally observed.
        if (cancel?.isCancelled()) throw new CancelledError();
      }
    }
  };
}
