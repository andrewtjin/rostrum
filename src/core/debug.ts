// Rostrum's diagnostics backbone — the thing that makes a live Office.js add-in
// debuggable at all.
//
// WHY THIS EXISTS: the Stage-1 engine is pure and trivially testable, but Stage 2
// runs inside a real Word host where you cannot attach a normal debugger, console
// output is easy to miss, and the most interesting failures (a proxy invalidating
// across a sync, an `insertOoxml` that silently drops list numbering, a Track-
// Changes restore that throws) happen deep inside an async `Word.run`. So every
// host-touching module logs through ONE tracer that:
//   * keeps a ring buffer of recent entries (survives the moment they scroll past),
//   * timestamps + sequences + namespaces + correlates every entry to an operation,
//   * times operations with spans,
//   * fans out to pluggable sinks (console + the task pane's live log viewer),
//   * clamps big payloads so a 200-page OOXML string can't OOM the buffer,
//   * produces a one-click bug report.
//
// It is pure and dependency-free (no Office.js, no DOM, no Node APIs) so it runs
// identically in unit tests, the task-pane browser, and the ribbon command runtime,
// and is itself fully unit-tested. Callers inject a clock/console/storage in tests.

import { StorageLike } from "./settings";

/** Severity, ordered debug < info < warn < error. */
export type LogLevel = "debug" | "info" | "warn" | "error";

/** Numeric rank for threshold comparisons. Higher = more severe. */
const LEVEL_RANK: Readonly<Record<LogLevel, number>> = {
  debug: 10,
  info: 20,
  warn: 30,
  error: 40
};

/** One recorded event. `data` is structured context (clamped on capture). */
export interface LogEntry {
  /** Monotonic per-tracer sequence number (stable ordering even at equal times). */
  seq: number;
  /** Epoch milliseconds from the injected clock. */
  time: number;
  level: LogLevel;
  /** Module namespace, e.g. "adapter", "live", "pane". */
  namespace: string;
  /** Operation correlation id, e.g. "hide#7" — ties child logs to one user action. */
  op?: string;
  msg: string;
  /** Small structured payload (long strings truncated, depth-bounded). */
  data?: unknown;
  /** Elapsed ms, present only on span-end entries. */
  durationMs?: number;
}

/** A consumer of entries (console, the UI live viewer, a test spy). */
export type LogSink = (entry: LogEntry) => void;

/** localStorage key for the persisted verbosity level (`.v1` allows reshaping). */
export const DEBUG_LEVEL_KEY = "rostrum.debugLevel.v1";

export interface TracerOptions {
  /** Entries below this level are dropped entirely (not buffered, not emitted). */
  minLevel?: LogLevel;
  /** Ring-buffer capacity; oldest entries evict first. */
  bufferSize?: number;
  /** Per-string clamp inside `data` (chars). Guards against buffering huge OOXML. */
  maxDataChars?: number;
  /** Injected clock; defaults to Date.now. Tests pass a deterministic counter. */
  clock?: () => number;
  /**
   * Console to mirror entries to, or null to stay silent (ribbon runtime / tests).
   * Defaults to the global console when available.
   */
  console?: Pick<Console, "debug" | "info" | "warn" | "error"> | null;
  /** Storage for persisting the level across sessions; null to skip. */
  storage?: StorageLike | null;
}

/** A running timer opened by `Logger.span`; call `end`/`fail` to log its duration. */
export interface Span {
  /** Close the span, logging an info entry with `durationMs` and extra context. */
  end(data?: Record<string, unknown>): number;
  /** Close the span at error level (a failed operation), preserving the cause. */
  fail(err: unknown, data?: Record<string, unknown>): number;
}

/**
 * Recursively clamp a value for safe capture: long strings are truncated with a
 * `…(+N)` marker, objects/arrays are walked to a bounded depth, and cycles are
 * cut. This is what lets callers pass an OOXML snippet or an Office error without
 * risking an unbounded or circular buffer.
 */
export function clampData(value: unknown, maxChars: number, depth = 4, seen = new Set<unknown>()): unknown {
  if (typeof value === "string") {
    return value.length > maxChars ? `${value.slice(0, maxChars)}…(+${value.length - maxChars})` : value;
  }
  if (value === null || typeof value !== "object") return value;
  if (seen.has(value)) return "[Circular]";
  if (depth <= 0) return Array.isArray(value) ? "[Array]" : "[Object]";
  seen.add(value);
  if (Array.isArray(value)) {
    return value.slice(0, 50).map((v) => clampData(v, maxChars, depth - 1, seen));
  }
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>).slice(0, 50)) {
    out[key] = clampData((value as Record<string, unknown>)[key], maxChars, depth - 1, seen);
  }
  return out;
}

/**
 * Extract the diagnostically useful fields from an unknown throw — especially an
 * `OfficeExtension.Error`, whose `code`/`debugInfo` are where the real cause lives
 * (a bare `.message` is often just "GeneralException"). Never throws itself.
 */
export function describeError(err: unknown): Record<string, unknown> {
  if (err == null || typeof err !== "object") return { error: String(err) };
  const e = err as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  if (typeof e.name === "string") out.name = e.name;
  if (typeof e.message === "string") out.message = e.message;
  // OfficeExtension.Error surface (present on Word.run failures).
  if (e.code != null) out.code = e.code;
  if (e.traceMessages != null) out.traceMessages = e.traceMessages;
  if (e.debugInfo != null && typeof e.debugInfo === "object") {
    const d = e.debugInfo as Record<string, unknown>;
    out.debugInfo = {
      code: d.code,
      message: d.message,
      errorLocation: d.errorLocation,
      statement: d.statement
    };
  }
  // Keep the stack last and short — it's noisy but occasionally decisive.
  if (typeof e.stack === "string") out.stack = e.stack;
  return out;
}

/**
 * The central tracer. One instance is shared process-wide (see the `tracer`
 * singleton below); modules obtain namespaced `Logger`s from it. Holding the
 * buffer and sinks in one place is what lets the task pane render a unified,
 * cross-module timeline of everything that happened during an operation.
 */
export class Tracer {
  private seq = 0;
  private minRank: number;
  private readonly buffer: LogEntry[] = [];
  private readonly sinks = new Set<LogSink>();
  private readonly opCounters = new Map<string, number>();
  private readonly maxDataChars: number;
  private readonly bufferSize: number;
  private readonly clock: () => number;
  private readonly out: Pick<Console, "debug" | "info" | "warn" | "error"> | null;
  private readonly storage: StorageLike | null;

  constructor(opts: TracerOptions = {}) {
    this.bufferSize = opts.bufferSize ?? 500;
    this.maxDataChars = opts.maxDataChars ?? 2000;
    this.clock = opts.clock ?? (() => Date.now());
    this.storage = opts.storage ?? null;
    // Default to the global console only if one exists (it does in browser + Node).
    this.out =
      opts.console !== undefined
        ? opts.console
        : typeof console !== "undefined"
          ? console
          : null;
    // A persisted level wins over the constructor default so a user who turned on
    // verbose logging keeps it across reloads.
    const persisted = this.readPersistedLevel();
    this.minRank = LEVEL_RANK[persisted ?? opts.minLevel ?? "info"];
  }

  /** The tracer's clock — used by spans to measure deltas without buffering. */
  now(): number {
    return this.clock();
  }

  /** Read the saved verbosity level, tolerating absent/garbled storage. */
  private readPersistedLevel(): LogLevel | null {
    if (!this.storage) return null;
    try {
      const raw = this.storage.getItem(DEBUG_LEVEL_KEY);
      return raw && raw in LEVEL_RANK ? (raw as LogLevel) : null;
    } catch {
      return null;
    }
  }

  /** Change the threshold at runtime and persist it (best-effort). */
  setMinLevel(level: LogLevel): void {
    this.minRank = LEVEL_RANK[level];
    if (this.storage) {
      try {
        this.storage.setItem(DEBUG_LEVEL_KEY, level);
      } catch {
        // A full / blocked storage must never break logging.
      }
    }
  }

  /** The current threshold as a level string. */
  getMinLevel(): LogLevel {
    return (Object.keys(LEVEL_RANK) as LogLevel[]).find((l) => LEVEL_RANK[l] === this.minRank) ?? "info";
  }

  /** True when an entry at `level` would be recorded — lets callers skip building data. */
  isEnabled(level: LogLevel): boolean {
    return LEVEL_RANK[level] >= this.minRank;
  }

  /**
   * Mint a unique operation id like `hide#3`, so every log line from one user
   * action shares a correlation tag the UI can group on.
   */
  nextOpId(name: string): string {
    const n = (this.opCounters.get(name) ?? 0) + 1;
    this.opCounters.set(name, n);
    return `${name}#${n}`;
  }

  /** A namespaced logger. Optionally pre-bound to an operation id. */
  logger(namespace: string, op?: string): Logger {
    return new Logger(this, namespace, op);
  }

  /** Subscribe a sink; returns an unsubscribe function. */
  subscribe(sink: LogSink): () => void {
    this.sinks.add(sink);
    return () => {
      this.sinks.delete(sink);
    };
  }

  /** A copy of the ring buffer, oldest-first. */
  getBuffer(): LogEntry[] {
    return this.buffer.slice();
  }

  /** Drop all buffered entries (the UI "clear log" action). */
  clear(): void {
    this.buffer.length = 0;
  }

  /**
   * Core record path. Builds a clamped entry, evicts the oldest if the buffer is
   * full, mirrors to the console, and fans out to sinks. A throwing sink is
   * isolated so one bad subscriber can't blind the others.
   */
  record(
    level: LogLevel,
    namespace: string,
    msg: string,
    data?: unknown,
    op?: string,
    durationMs?: number
  ): LogEntry | null {
    if (LEVEL_RANK[level] < this.minRank) return null;
    const entry: LogEntry = {
      seq: ++this.seq,
      time: this.clock(),
      level,
      namespace,
      msg
    };
    if (op !== undefined) entry.op = op;
    if (data !== undefined) entry.data = clampData(data, this.maxDataChars);
    if (durationMs !== undefined) entry.durationMs = durationMs;

    this.buffer.push(entry);
    if (this.buffer.length > this.bufferSize) this.buffer.shift();

    this.emitToConsole(entry);
    for (const sink of this.sinks) {
      try {
        sink(entry);
      } catch {
        // Sink failures are swallowed: diagnostics must never throw into callers.
      }
    }
    return entry;
  }

  /** Mirror an entry to the injected console in a compact, scannable form. */
  private emitToConsole(entry: LogEntry): void {
    if (!this.out) return;
    const line = formatEntry(entry);
    const fn =
      entry.level === "error"
        ? this.out.error
        : entry.level === "warn"
          ? this.out.warn
          : entry.level === "debug"
            ? this.out.debug
            : this.out.info;
    if (entry.data !== undefined) fn.call(this.out, line, entry.data);
    else fn.call(this.out, line);
  }

  /**
   * A human-readable bug report: a header line plus the recent timeline. The UI
   * appends host/platform info; keeping the tracer DOM-free means this is just the
   * captured events, safe to paste into an issue.
   */
  bugReport(header?: Record<string, unknown>): string {
    const lines: string[] = ["=== Rostrum diagnostics ==="];
    if (header) lines.push(JSON.stringify(header, null, 2));
    lines.push(`buffer: ${this.buffer.length} entries (min level ${this.getMinLevel()})`);
    lines.push("---");
    for (const e of this.buffer) {
      lines.push(formatEntry(e) + (e.data !== undefined ? ` ${safeStringify(e.data)}` : ""));
    }
    return lines.join("\n");
  }
}

/** JSON.stringify that can't throw (cycles already cut by clampData on capture). */
function safeStringify(v: unknown): string {
  try {
    return JSON.stringify(v);
  } catch {
    return String(v);
  }
}

/** `12:34:56.789 LEVEL ns op | msg (+Nms)` — fixed shape so logs grep/scan cleanly. */
export function formatEntry(e: LogEntry): string {
  const ts = formatTime(e.time);
  const lvl = e.level.toUpperCase().padEnd(5);
  const where = e.op ? `${e.namespace} ${e.op}` : e.namespace;
  const dur = e.durationMs !== undefined ? ` (+${Math.round(e.durationMs)}ms)` : "";
  return `${ts} ${lvl} ${where} | ${e.msg}${dur}`;
}

/** HH:MM:SS.mmm in UTC — locale-independent so logs compare across machines. */
function formatTime(ms: number): string {
  const d = new Date(ms);
  const p = (n: number, w = 2): string => String(n).padStart(w, "0");
  return `${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}.${p(d.getUTCMilliseconds(), 3)}`;
}

/**
 * A namespaced, optionally op-bound view over the tracer. This is what modules
 * actually hold. `child(op)` derives a new logger sharing the namespace but tagged
 * with a fresh operation id, so an adapter call started by the pane and one started
 * by the ribbon are distinguishable in the same buffer.
 */
export class Logger {
  constructor(
    private readonly tracer: Tracer,
    private readonly namespace: string,
    private readonly op?: string
  ) {}

  /** Derive a child logger bound to a new operation id minted from `name`. */
  child(name: string): Logger {
    return new Logger(this.tracer, this.namespace, this.tracer.nextOpId(name));
  }

  /** The operation id this logger carries, if any (for surfacing in results). */
  get operation(): string | undefined {
    return this.op;
  }

  /**
   * The tracer's clock — the SAME monotonic source `span()` brackets its `durationMs` with.
   * Exposed so a host adapter can take per-stage `now()` readings AT the actual stage boundaries
   * (read sync ↔ parse ↔ serialize ↔ commit sync) and emit ONE aggregate `debug()` line built from
   * those deltas (Loop 002 A1 / 002-S7), WITHOUT opening a span per stage (a span-end is info-level
   * and stamps `durationMs`, which the aggregate must NOT carry). Routing through the tracer means a
   * test's injected fake clock drives these readings deterministically, exactly like it drives spans.
   */
  now(): number {
    return this.tracer.now();
  }

  debug(msg: string, data?: unknown): void {
    this.tracer.record("debug", this.namespace, msg, data, this.op);
  }
  info(msg: string, data?: unknown): void {
    this.tracer.record("info", this.namespace, msg, data, this.op);
  }
  warn(msg: string, data?: unknown): void {
    this.tracer.record("warn", this.namespace, msg, data, this.op);
  }
  error(msg: string, data?: unknown): void {
    this.tracer.record("error", this.namespace, msg, data, this.op);
  }

  /** Log an error-level entry from a caught throw, expanding Office error fields. */
  caught(msg: string, err: unknown, extra?: Record<string, unknown>): void {
    this.tracer.record("error", this.namespace, msg, { ...describeError(err), ...extra }, this.op);
  }

  /**
   * Open a timing span. The returned handle logs a debug "start" immediately and
   * an info/error "end" with `durationMs` when closed — so a slow whole-body
   * commit shows up as one timed, correlated pair in the log.
   */
  span(name: string, data?: Record<string, unknown>): Span {
    const start = this.tracer.now();
    this.tracer.record("debug", this.namespace, `▶ ${name}`, data, this.op);
    const { tracer, namespace, op } = this;
    return {
      end: (endData?: Record<string, unknown>): number => {
        const dur = tracer.now() - start;
        tracer.record("info", namespace, `✔ ${name}`, endData, op, dur);
        return dur;
      },
      fail: (err: unknown, endData?: Record<string, unknown>): number => {
        const dur = tracer.now() - start;
        tracer.record("error", namespace, `✗ ${name}`, { ...describeError(err), ...endData }, op, dur);
        return dur;
      }
    };
  }
}

/**
 * The process-wide tracer. Modules import `logger(ns)` to get a namespaced logger
 * that writes into the one shared buffer the task pane renders. Tests construct
 * their own `new Tracer({clock})` for determinism and isolation.
 */
export const tracer = new Tracer({
  // In the browser, persist the chosen level so "verbose" survives a reload.
  storage: typeof localStorage !== "undefined" ? localStorage : null
});

/** Convenience: a namespaced logger off the shared tracer. */
export function logger(namespace: string, op?: string): Logger {
  return tracer.logger(namespace, op);
}
