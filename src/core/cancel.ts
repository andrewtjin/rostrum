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

// ── The macrotask yield: a clamp-free tier chain (Loop 002 A4 / 002-S5) ──────────────────────
//
// WHY A TIER CHAIN: the live task pane runs on Chromium (the WebView2/Edge host). Chromium
// CLAMPS nested `setTimeout` to ≥4ms once the timer-nesting depth reaches 5 — and a paced Hide
// fires far more than five yields in a row, so every `setTimeout(0)` slice past the fifth pays a
// forced ~4ms idle stall. Over a large doc that is tens of ms of pure dead wall-clock. The fix is
// a macrotask that ISN'T a timer (so it never enters the nested-timer clamp) yet still returns
// control to the host's event loop so paint runs and the queued Cancel click is delivered:
//   tier 1 · `scheduler.yield()`  — the standards-track "yield to the host" primitive; a true
//            task-source continuation, no timer clamp. Newest Chromium only.
//   tier 2 · a PERSISTENT `MessageChannel` postMessage round-trip — a task-queue macrotask that
//            predates `scheduler` and is NOT a timer, so it dodges the ≥4ms clamp. The channel is
//            allocated ONCE (first browser yield) and reused for every later yield — never per tick.
//   tier 3 · `setTimeout(0)` — the universal fallback; the original, unchanged behavior.
//
// THE S-008 CONTRACT each tier must honor: the yield returns control to the loop that delivers
// the Cancel click, so the pacer's post-yield `isCancelled()` re-check still observes a cancel
// that landed during the yield. tiers 1+3 satisfy this in every host. Tier 2 satisfies it ONLY in
// a BROWSER, where `MessageChannel` is an HTML task-source macrotask interleaved with input and
// timers. In Node (`worker_threads` MessageChannel) port messages drain on an internal queue that
// runs AHEAD of the timer phase, so a timer-delivered cancel would starve — which is exactly why
// tier 2 is gated on a browser global below. Under Node/jest there is no `window`, so detection
// falls straight through to tier 3 (`setTimeout`), preserving the FIFO/cancel-landing semantics
// the pacing e2e tests pin (a `setTimeout`-modeled Cancel click is observed after the first yield).
//
// KNOWN WET RISK (carried to the wet packet, NOT settled headlessly): `scheduler.yield()` (and a
// browser MessageChannel) continuations can run AHEAD of a due `setTimeout`, so on the xlarge doc
// the 3500ms grace-dialog timer or a timer-delivered Cancel could fire late. If the live xlarge
// run shows that starvation, the documented fix is to DEMOTE `scheduler.yield` below
// MessageChannel — a ONE-LINE edit: reorder `YIELD_TIERS` below.

/** A tier's macrotask yield, plus whether it is usable in the current host (feature/host probe). */
interface YieldTier {
  /** Human-readable name — for the (commented) demote point and any future diagnostics. */
  readonly name: string;
  /** True iff this tier's primitive exists AND returns control to the cancel-delivering loop here. */
  isAvailable(): boolean;
  /** Build the actual `() => Promise<void>` yield. Called ONCE, only after `isAvailable()` is true. */
  create(): () => Promise<void>;
}

/**
 * True in a real browser/task-pane runtime, false under Node/jest. Used to gate the
 * MessageChannel tier: a Node `worker_threads` MessageChannel does NOT return control to the timer
 * phase (it would starve a timer-delivered Cancel), whereas a browser MessageChannel is a proper
 * HTML task-source macrotask. `window`/`self` are present in the Office.js pane and absent in the
 * `testEnvironment:"node"` jest run, so this cleanly routes Node to the `setTimeout` fallback.
 */
const isBrowserHost = (): boolean =>
  typeof window !== "undefined" || typeof self !== "undefined";

/**
 * The ordered tier chain — FIRST available tier wins. **This list is the demote point**: to make
 * MessageChannel tier 1 (the documented fallback if `scheduler.yield()` starves the grace timer on
 * the live xlarge doc — see KNOWN WET RISK above), swap the first two entries. Nothing else changes.
 */
const YIELD_TIERS: readonly YieldTier[] = [
  // Tier 1 — scheduler.yield(): the host's native "yield to the event loop", no timer clamp.
  {
    name: "scheduler.yield",
    isAvailable: () =>
      typeof (globalThis as { scheduler?: { yield?: unknown } }).scheduler?.yield === "function",
    // Bind to the live `scheduler` so we call the host's real implementation each yield. The cast
    // goes through `unknown` because `scheduler` is not in this project's DOM lib typings yet.
    create: () => () =>
      (globalThis as unknown as { scheduler: { yield: () => Promise<void> } }).scheduler.yield()
  },
  // Tier 2 — persistent MessageChannel: a non-timer macrotask, allocated ONCE and reused.
  {
    name: "MessageChannel",
    isAvailable: () => isBrowserHost() && typeof MessageChannel === "function",
    create: () => {
      // Allocate the channel a SINGLE time (lazily, on first browser yield) and reuse both ports
      // for every subsequent yield — allocating per tick would defeat the point and churn memory.
      // A queue of pending resolvers makes the single port reentrancy-safe: each `yield()` posts a
      // message and parks its resolver; each delivered message resolves the OLDEST waiter (FIFO).
      const channel = new MessageChannel();
      const waiters: Array<() => void> = [];
      channel.port1.onmessage = () => {
        // Shift FIFO so N concurrent yields resolve in post order (the pacer is sequential, so in
        // practice there is at most one waiter — the queue is purely defensive against reentrancy).
        const resolve = waiters.shift();
        if (resolve) resolve();
      };
      // `start()` is a no-op when `onmessage` is set (it auto-starts), but call it explicitly so the
      // contract is clear and the port is open even if a future refactor uses addEventListener.
      channel.port1.start?.();
      // In a browser the open channel is harmless (it lives for the page). Under a Node TEST host
      // (`__makeYieldForTier`) the same open ports would keep the event loop alive and trip jest's
      // "did not exit" guard — so `unref()` them when present (Node MessagePort only; the browser
      // MessagePort has no `unref`, so this is a no-op there). This keeps the channel persistent and
      // fully functional while letting the process exit once nothing else is pending.
      (channel.port1 as { unref?: () => void }).unref?.();
      (channel.port2 as { unref?: () => void }).unref?.();
      return () =>
        new Promise<void>((resolve) => {
          waiters.push(resolve);
          channel.port2.postMessage(0);
        });
    }
  },
  // Tier 3 — setTimeout(0): the universal fallback and the original, unchanged behavior.
  {
    name: "setTimeout",
    isAvailable: () => typeof setTimeout === "function",
    create: () => () => new Promise<void>((resolve) => setTimeout(resolve, 0))
  }
];

/**
 * The resolved yield, computed ONCE on first use (lazy so module load never touches a host
 * primitive). Caching also means the persistent MessageChannel — allocated inside the winning
 * tier's `create()` — is built exactly once and shared by every `defaultYield()` call.
 */
let resolvedYield: (() => Promise<void>) | undefined;

/**
 * One macrotask that yields to the host without Chromium's nested-timer ≥4ms clamp: the first
 * available tier of `scheduler.yield()` → persistent `MessageChannel` → `setTimeout(0)`. Lets the
 * event loop paint and deliver the queued Cancel click; the pacer re-checks cancellation after it.
 */
const defaultYield = (): Promise<void> => {
  if (!resolvedYield) {
    // `find` is guaranteed to match — tier 3 (`setTimeout`) is available in every JS host — but
    // keep a defensive fallback so a hypothetical setTimeout-less host still yields a microtask.
    const tier = YIELD_TIERS.find((t) => t.isAvailable());
    resolvedYield = tier ? tier.create() : () => Promise.resolve();
  }
  return resolvedYield();
};

/**
 * Test-only: the NAME of the tier `defaultYield` selects in the CURRENT host (without allocating
 * the channel or caching a yield). Lets the pacing tests assert which tier a given environment
 * exercises — under `testEnvironment:"node"` this is `"setTimeout"` (no browser global → tier 2 is
 * skipped), pinning that the FIFO/cancel-landing semantics the e2e tests rely on are unchanged.
 * Not part of the runtime contract; the engine never calls it.
 */
export function __selectedYieldTierName(): string {
  return (YIELD_TIERS.find((t) => t.isAvailable()) ?? { name: "noop" }).name;
}

/**
 * Test-only: build a FRESH yield from the named tier's real `create()` factory, bypassing host
 * detection and the cached singleton. Tiers 1 (`scheduler.yield`) and 2 (`MessageChannel`) only
 * auto-select in a browser, so this is the one way to exercise their actual production bodies —
 * the persistent-channel allocation/round-trip and the scheduler binding — under jest's node env.
 * Returns `undefined` if the tier's primitive is genuinely absent here (then the test skips it).
 * Not part of the runtime contract; the engine always goes through `defaultYield`.
 */
export function __makeYieldForTier(name: string): (() => Promise<void>) | undefined {
  const tier = YIELD_TIERS.find((t) => t.name === name);
  if (!tier) return undefined;
  // `MessageChannel` exists in Node (worker_threads) even though `isBrowserHost()` is false, so its
  // `create()` runs here; `scheduler.yield` is genuinely absent in Node, so guard on the primitive.
  if (name === "scheduler.yield" && !tier.isAvailable()) return undefined;
  return tier.create();
}

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
