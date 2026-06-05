import {
  Tracer,
  clampData,
  describeError,
  formatEntry,
  LogEntry,
  DEBUG_LEVEL_KEY
} from "../src/core/debug";

/** A clock the tests advance by hand, so durations and timestamps are deterministic. */
function fakeClock(): { now: () => number; advance: (ms: number) => void } {
  let t = 1_000_000; // arbitrary fixed epoch
  return { now: () => t, advance: (ms) => (t += ms) };
}

/** A tracer with no console mirror, deterministic clock, small buffer for eviction tests. */
function mkTracer(opts: Partial<ConstructorParameters<typeof Tracer>[0]> = {}) {
  const clock = fakeClock();
  const tracer = new Tracer({ console: null, clock: clock.now, bufferSize: 3, minLevel: "debug", ...opts });
  return { tracer, clock };
}

describe("clampData", () => {
  it("truncates long strings with a +N marker", () => {
    const out = clampData("x".repeat(50), 10) as string;
    expect(out).toBe("xxxxxxxxxx…(+40)");
  });

  it("leaves short strings, numbers, booleans, null untouched", () => {
    expect(clampData("hi", 10)).toBe("hi");
    expect(clampData(42, 10)).toBe(42);
    expect(clampData(null, 10)).toBeNull();
  });

  it("walks objects/arrays and clamps nested strings", () => {
    const out = clampData({ a: "y".repeat(20), b: [1, 2, "z".repeat(20)] }, 5) as Record<string, unknown>;
    expect(out.a).toBe("yyyyy…(+15)");
    expect((out.b as unknown[])[2]).toBe("zzzzz…(+15)");
  });

  it("cuts cycles instead of overflowing", () => {
    const a: Record<string, unknown> = {};
    a.self = a;
    expect(() => clampData(a, 10)).not.toThrow();
    expect((clampData(a, 10) as Record<string, unknown>).self).toBe("[Circular]");
  });

  it("bounds recursion depth", () => {
    const deep = { l1: { l2: { l3: { l4: { l5: "too deep" } } } } };
    const out = JSON.stringify(clampData(deep, 100, 2));
    expect(out).toContain("[Object]"); // depth cut renders a placeholder
  });
});

describe("describeError", () => {
  it("extracts OfficeExtension.Error fields (code + debugInfo)", () => {
    const officeErr = {
      name: "OfficeExtension.Error",
      code: "GeneralException",
      message: "Something failed",
      debugInfo: { code: "GeneralException", message: "deep msg", errorLocation: "Body.insertOoxml" }
    };
    const out = describeError(officeErr);
    expect(out.code).toBe("GeneralException");
    expect((out.debugInfo as Record<string, unknown>).errorLocation).toBe("Body.insertOoxml");
  });

  it("handles non-objects and null without throwing", () => {
    expect(describeError("boom")).toEqual({ error: "boom" });
    expect(describeError(null)).toEqual({ error: "null" });
  });
});

describe("Tracer recording + level threshold", () => {
  it("drops entries below the min level", () => {
    const { tracer } = mkTracer({ minLevel: "warn" });
    tracer.logger("t").info("ignored");
    tracer.logger("t").warn("kept");
    expect(tracer.getBuffer().map((e) => e.msg)).toEqual(["kept"]);
  });

  it("evicts oldest entries when the ring buffer is full", () => {
    const { tracer } = mkTracer({ bufferSize: 2 });
    const log = tracer.logger("t");
    log.info("a");
    log.info("b");
    log.info("c");
    expect(tracer.getBuffer().map((e) => e.msg)).toEqual(["b", "c"]);
  });

  it("assigns monotonic sequence numbers and the clock's timestamp", () => {
    const { tracer, clock } = mkTracer();
    tracer.logger("t").info("first");
    clock.advance(5);
    tracer.logger("t").info("second");
    const [a, b] = tracer.getBuffer();
    expect(b.seq).toBe(a.seq + 1);
    expect(b.time).toBe(a.time + 5);
  });

  it("isEnabled reflects the threshold", () => {
    const { tracer } = mkTracer({ minLevel: "info" });
    expect(tracer.isEnabled("debug")).toBe(false);
    expect(tracer.isEnabled("error")).toBe(true);
  });
});

describe("subscribers", () => {
  it("delivers entries to sinks and stops after unsubscribe", () => {
    const { tracer } = mkTracer();
    const seen: string[] = [];
    const off = tracer.subscribe((e) => seen.push(e.msg));
    tracer.logger("t").info("one");
    off();
    tracer.logger("t").info("two");
    expect(seen).toEqual(["one"]);
  });

  it("isolates a throwing sink so others still receive", () => {
    const { tracer } = mkTracer();
    const seen: string[] = [];
    tracer.subscribe(() => {
      throw new Error("bad sink");
    });
    tracer.subscribe((e) => seen.push(e.msg));
    expect(() => tracer.logger("t").info("hi")).not.toThrow();
    expect(seen).toEqual(["hi"]);
  });
});

describe("operation ids + spans", () => {
  it("mints incrementing op ids per name", () => {
    const { tracer } = mkTracer();
    expect(tracer.nextOpId("hide")).toBe("hide#1");
    expect(tracer.nextOpId("hide")).toBe("hide#2");
    expect(tracer.nextOpId("show")).toBe("show#1");
  });

  it("child loggers carry their op id onto every entry", () => {
    const { tracer } = mkTracer();
    const op = tracer.logger("adapter").child("hide");
    op.info("reading");
    expect(tracer.getBuffer()[0].op).toBe("hide#1");
  });

  it("span end logs an info entry with the measured duration", () => {
    const { tracer, clock } = mkTracer();
    const span = tracer.logger("adapter").span("commit");
    clock.advance(42);
    const dur = span.end({ paragraphs: 3 });
    expect(dur).toBe(42);
    const end = tracer.getBuffer().find((e) => e.msg.includes("commit") && e.durationMs !== undefined);
    expect(end?.durationMs).toBe(42);
    expect(end?.level).toBe("info");
  });

  it("span fail logs an error entry expanding the cause", () => {
    const { tracer, clock } = mkTracer();
    const span = tracer.logger("adapter").span("commit");
    clock.advance(7);
    span.fail({ code: "GeneralException", message: "nope" });
    const fail = tracer.getBuffer().find((e) => e.level === "error");
    expect(fail?.durationMs).toBe(7);
    expect((fail?.data as Record<string, unknown>).code).toBe("GeneralException");
  });
});

describe("level persistence", () => {
  it("saves the level to storage and reads it back on construction", () => {
    const store = new Map<string, string>();
    const storage = {
      getItem: (k: string) => store.get(k) ?? null,
      setItem: (k: string, v: string) => void store.set(k, v)
    };
    const t1 = new Tracer({ console: null, storage });
    t1.setMinLevel("debug");
    expect(store.get(DEBUG_LEVEL_KEY)).toBe("debug");

    // A fresh tracer pointed at the same storage starts at the persisted level.
    const t2 = new Tracer({ console: null, storage });
    expect(t2.getMinLevel()).toBe("debug");
  });

  it("survives storage that throws", () => {
    const storage = {
      getItem: () => {
        throw new Error("blocked");
      },
      setItem: () => {
        throw new Error("blocked");
      }
    };
    expect(() => new Tracer({ console: null, storage }).setMinLevel("warn")).not.toThrow();
  });
});

describe("formatEntry + bugReport", () => {
  const entry: LogEntry = {
    seq: 1,
    time: Date.UTC(2026, 5, 2, 12, 34, 56, 789),
    level: "info",
    namespace: "adapter",
    op: "hide#1",
    msg: "done",
    durationMs: 123
  };

  it("formats a stable, scannable line", () => {
    expect(formatEntry(entry)).toBe("12:34:56.789 INFO  adapter hide#1 | done (+123ms)");
  });

  it("bugReport includes a header and the recent timeline", () => {
    const { tracer } = mkTracer();
    tracer.logger("t").error("kaboom", { code: "X" });
    const report = tracer.bugReport({ host: "Word", platform: "PC" });
    expect(report).toContain("Rostrum diagnostics");
    expect(report).toContain('"host": "Word"');
    expect(report).toContain("kaboom");
  });
});
