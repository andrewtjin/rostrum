// Loop 002 A1 / 002-S7 — Step-0 per-stage instrumentation spy test (review tests-6).
//
// The node-direct Hide emits ONE per-op aggregate timing line so the engine's per-stage cost is
// observable for the wet-packet Step-0 re-baseline. This pins that contract precisely:
//   * the line is ABSENT at the default `info` level (production is unaffected — verbose-only);
//   * at `debug` it appears EXACTLY ONCE, carrying ALL eight numeric stage fields, at `level:"debug"`;
//   * it carries NO `durationMs` — the span-end fingerprint — which PROVES it is an `op.debug()` call,
//     not a span-end (a span-end is info-level AND stamps durationMs; the aggregate must be neither);
//   * pure modules (ooxml / keepers / ooxmlPackage) emit NO tracer calls — the timing reads all live in
//     the port/engine layer, never in the pure modules.
//
// The fake clock auto-increments on every read, so each stage boundary samples a distinct, non-zero
// value deterministically (the FakeContext's `sync()` is synchronous and would otherwise leave every
// delta at 0). This drives the per-stage durations without any real wall-clock dependency.

import { Tracer, LogEntry } from "../src/core/debug";
import { createOfficeWordPort } from "../src/core/officeWordPort";
import { hide } from "../src/core/invisibility";
import { para, run, mkDoc, buildPackage, harness, settings } from "./fakeWord";

const lvl0 = `<w:pPr><w:outlineLvl w:val="0"/></w:pPr>`; // inline Heading 1 outline level

/** The exact set of numeric stage fields the aggregate must carry (every one present). */
const STAGE_FIELDS = [
  "tcGateMs",
  "readSyncMs",
  "parseMs",
  "classifyMs",
  "applyMs",
  "serializeMs",
  "commitSyncMs",
  "engineTotalMs"
] as const;

/** The aggregate line's message — the single stable identifier the spy keys on. */
const AGG_MSG = "engine stage timing";

/**
 * A monotonic auto-incrementing fake clock. Each call returns the next integer, so consecutive
 * `now()` readings at stage boundaries differ by a known amount — every per-stage delta is non-zero
 * and deterministic without touching the wall clock.
 */
function autoClock(): () => number {
  let t = 0;
  return () => ++t;
}

/**
 * Build a node-direct Hide setup: a fresh fake doc, a frozen whole-body package the read returns, a
 * tracer at the requested level wired to a sink that records every entry, and a `pureWholeBody` port
 * driven off that tracer's clock (so the fake clock drives the stage timings).
 *
 * `harness()` supplies the in-memory runner; we IGNORE its internal tracer and hand the port our own
 * (level-controlled, clock-controlled) logger — the port logs through the logger it is given, so the
 * spy observes exactly what production would at that level.
 */
function setup(minLevel: "info" | "debug") {
  const doc = mkDoc([
    para(run("Heading"), { outlineNumber: 10, pPr: lvl0 }),
    para(run("a long card body")),
    para(run("Author 2019", { cite: true })),
    para(run("intro ") + run("warrant", { highlight: "yellow" }))
  ]);
  const pkg = buildPackage(doc.paragraphs);
  const h = harness(doc, pkg); // frozen package: getOoxml returns the same well-formed body each read
  const tracer = new Tracer({ console: null, clock: autoClock(), minLevel });
  const entries: LogEntry[] = [];
  tracer.subscribe((e) => entries.push(e));
  const port = createOfficeWordPort({
    runner: h.runner,
    pureWholeBody: true,
    logger: tracer.logger("adapter")
  });
  return { doc, port, entries };
}

describe("Step-0 per-stage instrumentation (Loop 002 A1 / 002-S7)", () => {
  it("emits NO aggregate timing line at the default info level", async () => {
    const { port, entries } = setup("info");
    await hide(port, settings(["yellow"]));
    const agg = entries.filter((e) => e.msg === AGG_MSG);
    expect(agg).toHaveLength(0);
  });

  it("emits EXACTLY ONE aggregate line at debug — all stage fields, level debug, NO durationMs", async () => {
    const { port, entries } = setup("debug");
    await hide(port, settings(["yellow"]));

    const agg = entries.filter((e) => e.msg === AGG_MSG);
    expect(agg).toHaveLength(1);
    const line = agg[0];

    // It is a debug-level entry...
    expect(line.level).toBe("debug");
    // ...and carries NO durationMs — the span-end fingerprint. Its ABSENCE pins the line to op.debug(),
    // not a span-end (which would be info-level AND stamp durationMs). This is the crux of tests-6.
    expect(line.durationMs).toBeUndefined();
    expect("durationMs" in line).toBe(false);

    // Every stage field is present and a finite number.
    const data = line.data as Record<string, unknown>;
    for (const field of STAGE_FIELDS) {
      expect(typeof data[field]).toBe("number");
      expect(Number.isFinite(data[field] as number)).toBe(true);
    }
    // The line carries ONLY the stage fields (no stray keys leaking into the Step-0 row schema).
    expect(Object.keys(data).sort()).toEqual([...STAGE_FIELDS].sort());

    // engineTotalMs is the end-to-end engine span and must be ≥ every sub-stage (it brackets them all).
    const total = data.engineTotalMs as number;
    for (const field of STAGE_FIELDS) {
      if (field === "engineTotalMs") continue;
      expect(total).toBeGreaterThanOrEqual(data[field] as number);
    }
  });

  it("never emits a span-end masquerading as the aggregate (no entry with the stage shape AND durationMs)", async () => {
    const { port, entries } = setup("debug");
    await hide(port, settings(["yellow"]));
    // No entry both LOOKS like the aggregate (carries the stage fields) AND has durationMs.
    const masquerade = entries.filter(
      (e) =>
        e.durationMs !== undefined &&
        e.data != null &&
        typeof e.data === "object" &&
        "engineTotalMs" in (e.data as Record<string, unknown>)
    );
    expect(masquerade).toHaveLength(0);
  });

  it("emits exactly ONE line per Hide op (a re-hide on the same port emits a second, not a duplicate)", async () => {
    const { port, entries } = setup("debug");
    await hide(port, settings(["yellow"])); // frozen package → re-hide reads the same body
    await hide(port, settings(["yellow"]));
    const agg = entries.filter((e) => e.msg === AGG_MSG);
    expect(agg).toHaveLength(2);
  });
});

describe("pure modules stay tracer-free (Loop 002 A1 — pure-layer containment)", () => {
  it("only the adapter + engine namespaces ever record; ooxml/keepers/ooxmlPackage never do", async () => {
    const { port, entries } = setup("debug");
    await hide(port, settings(["yellow"]));
    // A pure module taking a tracer call would surface as an entry whose namespace is that module.
    const pureNamespaces = new Set(["ooxml", "keepers", "ooxmlPackage", "ooxmlpackage"]);
    const offenders = entries.filter((e) => pureNamespaces.has(e.namespace));
    expect(offenders).toHaveLength(0);
    // Sanity: the adapter DID record (so the test isn't vacuously green on an empty buffer).
    expect(entries.some((e) => e.namespace === "adapter")).toBe(true);
  });
});
