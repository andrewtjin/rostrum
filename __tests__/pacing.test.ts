// Pacing + cancellation for the long pure-JS phases (core/cancel.ts).
//
// The default pure whole-body Hide (avenue ⑦) runs ONE host read sync, then a long stretch
// of synchronous JS — package assembly + per-paragraph classify — before the commit sync.
// On the live single-threaded runtime that stretch froze the pane: progress could not paint
// and a Cancel click queued unprocessed, so the old bare `isCancelled()` poll could never
// observe a click made after the read sync resolved. The `Pacer` folds the cancel poll into
// a per-iteration `tick()` that also yields a macrotask on a time budget.
//
// These tests pin four things: (1) the pacer's own contract (budget metering, cancel
// observed before AND after a yield); (2) BYTE-IDENTITY — a paced hide writes exactly the
// bytes an unpaced hide writes (pacing may change scheduling, never output), on synthetic
// fixtures AND a real .docx; (3) a cancel landing in the yield window aborts PRE-WRITE:
// document byte-unchanged, nothing flushed, manifest never armed, Track Changes restored;
// (4) end-to-end through the RostrumController, a Cancel "click" queued AFTER Hide (or
// Re-hide — both wire the controller's one pacer) started actually lands mid-operation —
// the exact user story that was dead before pacing.

import {
  CancelledError,
  createPacer,
  __selectedYieldTierName,
  __makeYieldForTier
} from "../src/core/cancel";
import {
  CancelledError as ReExportedCancelledError,
  createOfficeWordPort
} from "../src/core/officeWordPort";
import { hide } from "../src/core/invisibility";
import * as invisibility from "../src/core/invisibility";
import { RostrumController } from "../src/taskpane/controller";
import { FeatureSupport } from "../src/core/types";
import { FakeDoc, mkDoc, para, run, harness, settings } from "./fakeWord";
import { discoverSamples, paragraphsFromDocumentXml, readDocxParts } from "./realDocs";

// ---------------------------------------------------------------------------
// The pacer contract itself (no engine involved).
// ---------------------------------------------------------------------------
describe("createPacer", () => {
  it("re-exports the SAME CancelledError class from the adapter (instanceof must keep working)", () => {
    // The controller maps `e instanceof CancelledError` → the "cancelled" outcome; if the
    // adapter's re-export ever became a copy instead of the same class, that check would
    // silently break. Identity, not just shape.
    expect(ReExportedCancelledError).toBe(CancelledError);
  });

  it("yields once per elapsed budget and re-arms the window (fake clock)", async () => {
    let t = 0;
    let yields = 0;
    const pacer = createPacer({ budgetMs: 50, now: () => t, yieldFn: async () => void yields++ });
    await pacer.tick(); // 0ms elapsed → stays synchronous
    t = 49;
    await pacer.tick(); // still inside the budget
    expect(yields).toBe(0);
    t = 50;
    await pacer.tick(); // budget elapsed → one yield, window re-arms at t=50
    expect(yields).toBe(1);
    t = 99;
    await pacer.tick(); // re-armed window not yet elapsed
    expect(yields).toBe(1);
    t = 100;
    await pacer.tick();
    expect(yields).toBe(2);
  });

  it("budget 0 yields on every tick; budget Infinity never yields", async () => {
    let always = 0;
    const eager = createPacer({ budgetMs: 0, yieldFn: async () => void always++ });
    await eager.tick();
    await eager.tick();
    expect(always).toBe(2);

    let never = 0;
    const lazy = createPacer({ budgetMs: Number.POSITIVE_INFINITY, yieldFn: async () => void never++ });
    await lazy.tick();
    await lazy.tick();
    expect(never).toBe(0);
  });

  it("throws CancelledError when already cancelled — without burning another yield", async () => {
    let yields = 0;
    const pacer = createPacer({
      cancel: { isCancelled: () => true },
      budgetMs: 0,
      yieldFn: async () => void yields++
    });
    await expect(pacer.tick()).rejects.toBeInstanceOf(CancelledError);
    expect(yields).toBe(0); // cancel is checked BEFORE the yield
  });

  it("observes a cancel that lands DURING the yield — the Cancel-click window", async () => {
    // This is the load-bearing re-check: the click handler runs inside the yielded
    // macrotask, so the pacer must look again AFTER the yield resolves.
    let cancelled = false;
    const pacer = createPacer({
      cancel: { isCancelled: () => cancelled },
      budgetMs: 0,
      yieldFn: async () => {
        cancelled = true; // the "click"
      }
    });
    await expect(pacer.tick()).rejects.toBeInstanceOf(CancelledError);
  });

  it("Infinity-budget pacer still observes a pre-cancelled token (the adapter's bare-poll fallback)", async () => {
    // The adapter wraps a plain `cancel` option in a never-yielding pacer; the old bare
    // `isCancelled()` poll semantics must survive that wrap exactly.
    const pacer = createPacer({
      cancel: { isCancelled: () => true },
      budgetMs: Number.POSITIVE_INFINITY
    });
    await expect(pacer.tick()).rejects.toBeInstanceOf(CancelledError);
  });
});

// ---------------------------------------------------------------------------
// The clamp-free yield tier chain (Loop 002 A4 / 002-S5).
//
// `defaultYield` picks the first available tier of scheduler.yield → persistent MessageChannel →
// setTimeout(0). These tests pin (a) which tier THIS environment exercises — under jest's
// `testEnvironment:"node"` it MUST be `setTimeout`, because the MessageChannel tier is gated on a
// browser global so a Node worker_threads MessageChannel (which starves timer-delivered cancels)
// is never selected, which is exactly why the FIFO/cancel-landing e2e tests above stay green; and
// (b) the S-008 guarantee for the MessageChannel tier itself — even though it isn't auto-selected
// in Node, a real MessageChannel round-trip still returns control such that a cancel landing during
// the yield is observed by the pacer's post-yield re-check (proven here with the cancel delivered
// on the MessageChannel task source, the way a real browser Cancel click arrives).
// ---------------------------------------------------------------------------
describe("clamp-free yield tier chain", () => {
  it("selects the setTimeout tier under jest's node env (no browser global → MessageChannel skipped)", () => {
    // This is the load-bearing reason the controller e2e cancel tests are byte-for-byte unchanged:
    // Node routes to setTimeout, so a setTimeout-modeled Cancel click still lands FIFO after the
    // first yield. (In the live Chromium task pane the same chain selects scheduler.yield /
    // MessageChannel instead — the clamp-free path — but the cancel contract is identical.)
    expect(__selectedYieldTierName()).toBe("setTimeout");
  });

  it("MessageChannel tier (S-008): a cancel landing DURING the real production yield is seen post-yield", async () => {
    // Drive the ACTUAL production MessageChannel tier (`__makeYieldForTier` runs its real
    // `create()` — the persistent channel + FIFO waiter queue), not a hand-rolled copy. Inject it
    // into a pacer and deliver the "Cancel click" on the MessageChannel task source DURING the
    // yield — modeling a real browser Cancel click arriving while the host is yielded. The pacer's
    // post-yield `isCancelled()` re-check must observe it: the same S-008 guarantee setTimeout
    // gives, now proven for the clamp-free tier the live task pane actually selects.
    const mcYield = __makeYieldForTier("MessageChannel");
    expect(mcYield).toBeDefined(); // MessageChannel exists in Node (worker_threads)

    let cancelled = false;
    const pacer = createPacer({
      cancel: { isCancelled: () => cancelled },
      budgetMs: 0, // yield on the very first tick
      yieldFn: () => {
        // Schedule the "click" on the MessageChannel task source so it lands during the yield
        // window — modeling a real browser Cancel click delivered while the host is yielded.
        const clickChannel = new MessageChannel();
        clickChannel.port1.onmessage = () => {
          cancelled = true;
          clickChannel.port1.close();
          clickChannel.port2.close();
        };
        clickChannel.port2.postMessage(0);
        return mcYield!();
      }
    });

    await expect(pacer.tick()).rejects.toBeInstanceOf(CancelledError);
  });

  it("the production MessageChannel yield round-trips repeatedly on ONE persistent channel (reuse, not per-tick alloc)", async () => {
    // The production tier allocates the channel ONCE in `create()` and reuses it for every yield.
    // `__makeYieldForTier` builds that closure once; calling the returned yield many times proves
    // the SAME channel round-trips each time (a stuck/closed port would hang this test) without
    // re-allocating — the persistent-allocation contract the mission requires.
    const mcYield = __makeYieldForTier("MessageChannel");
    expect(mcYield).toBeDefined();
    let yields = 0;
    const pacer = createPacer({
      budgetMs: 0,
      yieldFn: async () => {
        await mcYield!();
        yields++;
      }
    });
    await pacer.tick();
    await pacer.tick();
    await pacer.tick();
    expect(yields).toBe(3);
  });

  it("scheduler.yield tier (when present): the real tier body delegates to the host scheduler", async () => {
    // scheduler.yield is absent in Node, so its production `create()` body would otherwise be
    // uncovered. Stub a minimal `scheduler.yield` on globalThis, build the REAL tier yield, and
    // assert it both selects scheduler.yield AND calls through to the host primitive (the clamp-free
    // path the newest Chromium task pane takes). Restore the global so no other test is affected.
    const g = globalThis as unknown as { scheduler?: { yield: () => Promise<void> } };
    const had = "scheduler" in g;
    let calls = 0;
    g.scheduler = { yield: () => (calls++, Promise.resolve()) };
    try {
      // With scheduler.yield present AND a browser-less host, scheduler.yield is still tier 1
      // (its availability does not depend on a browser global — only MessageChannel does).
      expect(__selectedYieldTierName()).toBe("scheduler.yield");
      const schedYield = __makeYieldForTier("scheduler.yield");
      expect(schedYield).toBeDefined();
      await schedYield!();
      await schedYield!();
      expect(calls).toBe(2); // each yield delegated to the host scheduler
    } finally {
      if (!had) delete g.scheduler;
    }
  });
});

// ---------------------------------------------------------------------------
// Through the full adapter + engine (pure whole-body, the DEFAULT Hide path).
// ---------------------------------------------------------------------------

const lvl0 = `<w:pPr><w:outlineLvl w:val="0"/></w:pPr>`; // package-resolvable Heading 1

/** A representative mix: kept heading, hidden card, kept cite, partial highlight. Built
 *  fresh per call so paced/unpaced runs mutate independent (but identical) documents. */
const makeDoc = (tc: Parameters<typeof mkDoc>[1] = "Off"): FakeDoc =>
  mkDoc(
    [
      para(run("Heading"), { pPr: lvl0 }),
      para(run("a long card body")),
      para(run("Author 2019", { cite: true })),
      para(run("intro ") + run("warrant", { highlight: "yellow" }))
    ],
    tc
  );

describe("paced hide through the pure whole-body path", () => {
  it("writes BYTE-IDENTICAL output to an unpaced hide — pacing changes scheduling, never bytes", async () => {
    const docA = makeDoc();
    const docB = makeDoc();
    const hA = harness(docA);
    const hB = harness(docB);
    const portA = createOfficeWordPort({
      runner: hA.runner,
      pureWholeBody: true,
      logger: hA.tracer.logger("adapter")
    });
    // Budget 0 = a yield on EVERY tick — the maximally-perturbed schedule; output must not care.
    let yields = 0;
    const pacer = createPacer({ budgetMs: 0, yieldFn: async () => void yields++ });
    const portB = createOfficeWordPort({
      runner: hB.runner,
      pureWholeBody: true,
      pacer,
      logger: hB.tracer.logger("adapter")
    });

    const resA = await hide(portA, settings(["yellow"]));
    const resB = await hide(portB, settings(["yellow"]), { pacing: pacer });

    expect(docB.paragraphs.map((p) => p.xml)).toEqual(docA.paragraphs.map((p) => p.xml));
    expect(docB.manifest?.xml).toBe(docA.manifest?.xml);
    expect(resB).toEqual(resA);
    // Call-count guard (parseCount-style): with budget 0 (yield on every tick) the pacer is ticked
    // once per paragraph in EACH of the THREE paced loops — the pure read (port), Phase A classify,
    // and Phase B apply (Loop 002 B1: the apply stretch is now paced too, so a big-doc node-direct
    // apply can paint progress and land a mid-Hide Cancel). Exactly 3N. If a loop stops ticking (pane
    // freezes again) or starts double-ticking, this fails loudly.
    expect(yields).toBe(3 * docA.paragraphs.length);
  });

  it("a cancel landing mid-CLASSIFY aborts pre-write: doc untouched, nothing flushed, TC restored", async () => {
    // TC starts ON + autoToggle, so the gate toggles Off then MUST restore on the throw.
    const doc = makeDoc("TrackAll");
    const before = doc.paragraphs.map((p) => p.xml);
    const h = harness(doc);
    // The port gets NO pacer (read runs straight through); only the classify loop is paced,
    // and the first yield flips the flag — pinning cancellation INSIDE the classify phase.
    const port = createOfficeWordPort({
      runner: h.runner,
      pureWholeBody: true,
      logger: h.tracer.logger("adapter")
    });
    let cancelled = false;
    const pacing = createPacer({
      cancel: { isCancelled: () => cancelled },
      budgetMs: 0,
      yieldFn: async () => {
        cancelled = true; // the "click" during the classify loop's first yield
      }
    });

    await expect(
      hide(port, settings(["yellow"]), { autoToggleTrackChanges: true, pacing })
    ).rejects.toBeInstanceOf(CancelledError);

    expect(doc.paragraphs.map((p) => p.xml)).toEqual(before); // byte-unchanged
    expect(doc.manifest).toBeNull(); // never armed
    // The buffered updates were never flushed: no document/manifest write of any kind hit
    // the host. (tc.set entries are expected — the gate toggled Off and restored.)
    const writeOps = h.ctx.commitLog.filter((c) => /insertOoxml|setXml|add|font\.hidden/.test(c.op));
    expect(writeOps).toEqual([]);
    expect(doc.tcMode).toBe("TrackAll"); // the gate's finally restored the prior mode
  });

  it("a cancel landing mid-READ (the port's pacer) aborts the pure read the same way", async () => {
    const doc = makeDoc();
    const before = doc.paragraphs.map((p) => p.xml);
    const h = harness(doc);
    let cancelled = false;
    const pacer = createPacer({
      cancel: { isCancelled: () => cancelled },
      budgetMs: 0,
      yieldFn: async () => {
        cancelled = true;
      }
    });
    const port = createOfficeWordPort({
      runner: h.runner,
      pureWholeBody: true,
      pacer,
      logger: h.tracer.logger("adapter")
    });
    await expect(port.readParagraphs()).rejects.toBeInstanceOf(CancelledError);
    expect(doc.paragraphs.map((p) => p.xml)).toEqual(before);
  });
});

// ---------------------------------------------------------------------------
// End-to-end: the pane's Cancel button through the RostrumController.
// ---------------------------------------------------------------------------
describe("controller Cancel lands mid-Hide / mid-Re-hide (end-to-end)", () => {
  const FULL: FeatureSupport = {
    canHide: true,
    canCustomXml: true,
    canChangeTracking: true,
    canStyleBorders: true,
    canStyleFormat: true,
    canGetStyles: true
  };
  const memStorage = () => {
    const m = new Map<string, string>();
    return {
      getItem: (k: string) => m.get(k) ?? null,
      setItem: (k: string, v: string) => void m.set(k, v)
    };
  };

  it("a cancel() queued AFTER hide() started produces the 'cancelled' outcome, doc untouched", async () => {
    const doc = makeDoc();
    const before = doc.paragraphs.map((p) => p.xml);
    const h = harness(doc);
    // The injected clock advances 100ms per reading, so EVERY pacer tick sees an elapsed
    // budget and yields a real macrotask — the deterministic stand-in for a doc large
    // enough to overrun the 50ms budget.
    let t = 0;
    const ctrl = new RostrumController({
      features: FULL,
      runner: h.runner,
      storage: memStorage(),
      logger: h.tracer.logger("pane"),
      now: () => (t += 100)
    });
    await ctrl.init();

    const op = ctrl.hide(); // pure whole-body default — do NOT await yet
    // Queue the Cancel "click" as a macrotask registered AFTER hide() started, BEFORE its
    // first paced yield resumes (same-delay timers fire FIFO). Pre-pacing, the whole op ran
    // to completion in one JS turn — this timer would have fired too late and the outcome
    // would be "ok", so this test genuinely fails without the fix.
    await new Promise<void>((resolve) =>
      setTimeout(() => {
        ctrl.cancel();
        resolve();
      }, 0)
    );

    expect(await op).toEqual({ status: "cancelled" });
    expect(doc.paragraphs.map((p) => p.xml)).toEqual(before); // nothing written
    expect(doc.manifest).toBeNull(); // never armed
    expect(ctrl.status().armed).toBe(false);
  });

  it("a cancel() during reHide() lands the same way — the armed doc and manifest stay put", async () => {
    // Mirrors the hide() case for the OTHER paced mutation. What "nothing written" must
    // mean HERE is different: reHide re-derives over an ARMED doc, so a cancelled pass has
    // to leave the armed state — hidden bytes, manifest, armed flag — exactly as the first
    // Hide committed it. (This proves the cancel user story; the engine-side `pacing`
    // option gets its own dedicated pin in the next test, because the port's read-pacer
    // alone — the SAME shared pacer — can land this cancel even if reHide dropped it.)
    const doc = makeDoc();
    const h = harness(doc);
    let t = 0;
    const ctrl = new RostrumController({
      features: FULL,
      runner: h.runner,
      storage: memStorage(),
      logger: h.tracer.logger("pane"),
      now: () => (t += 100)
    });
    await ctrl.init();
    // Arm via a real prior Hide — reHide's user story is re-derivation over an ARMED doc
    // (catching newly typed text). Running it to completion also proves resetCancel()
    // re-arms the one SHARED pacer between operations.
    expect((await ctrl.hide()).status).toBe("ok");
    const armedParas = doc.paragraphs.map((p) => p.xml);
    const armedManifest = doc.manifest?.xml;

    const op = ctrl.reHide(); // do NOT await yet
    // Same FIFO-timer trick as the hide case above: the Cancel "click" macrotask is
    // registered after reHide() started, so it fires before the first paced yield resumes.
    await new Promise<void>((resolve) =>
      setTimeout(() => {
        ctrl.cancel();
        resolve();
      }, 0)
    );

    expect(await op).toEqual({ status: "cancelled" });
    // The cancelled re-derivation flushed NOTHING: document bytes and manifest are still
    // exactly the armed state the first Hide committed, and the controller stays armed.
    expect(doc.paragraphs.map((p) => p.xml)).toEqual(armedParas);
    expect(doc.manifest?.xml).toBe(armedManifest);
    expect(ctrl.status().armed).toBe(true);
  });

  it("reHide() hands the controller's pacer to the engine (the classify-loop wiring)", async () => {
    // The e2e above cannot catch reHide dropping `pacing: this.pacer` — the port's
    // read-pacer is the same shared pacer and lands the cancel on its own. But WITHOUT
    // the engine option the classify loop reverts to one unpaced synchronous stretch
    // (paint + Cancel dead mid-classify on a large doc), so the wiring itself is pinned:
    // the engine must receive a tickable pacing option. (spyOn calls through, so the op
    // still runs for real.)
    const spy = jest.spyOn(invisibility, "reHide");
    const doc = makeDoc();
    const h = harness(doc);
    const ctrl = new RostrumController({
      features: FULL,
      runner: h.runner,
      storage: memStorage(),
      logger: h.tracer.logger("pane")
    });
    await ctrl.init();
    expect((await ctrl.reHide()).status).toBe("ok");
    expect(spy).toHaveBeenCalledTimes(1);
    expect(typeof spy.mock.calls[0][2]?.pacing?.tick).toBe("function");
    spy.mockRestore();
  });
});

// ---------------------------------------------------------------------------
// Real-document byte-identity + honest overhead numbers (skipped when samples/ absent, e.g. CI).
// ---------------------------------------------------------------------------
describe("paced vs unpaced hide on a real document", () => {
  const KEEP_COLORS = ["cyan", "yellow", "green", "lightGray", "magenta", "red"];
  const lightSamples = discoverSamples().filter((s) => s.tier === "small" || s.tier === "medium");

  /** Hide the same real doc twice — unpaced and paced — and assert byte-identity. Returns
   *  timings + yield count so the perf-gated case can report honest overhead numbers. */
  async function pacedVsUnpaced(
    fullPath: string,
    pacerBudgetMs: number,
    makeYieldFn: (bump: () => void) => () => Promise<void>
  ): Promise<{ paras: number; unpacedMs: number; pacedMs: number; yields: number }> {
      const { documentXml, stylesXml } = await readDocxParts(fullPath);
      const base = paragraphsFromDocumentXml(documentXml, stylesXml);
      // Two independent copies of the same real paragraphs (hide mutates xml in place).
      const docA = mkDoc(base.map((p) => ({ ...p })));
      const docB = mkDoc(base.map((p) => ({ ...p })));
      const hA = harness(docA);
      const hB = harness(docB);
      const portA = createOfficeWordPort({
        runner: hA.runner,
        pureWholeBody: true,
        logger: hA.tracer.logger("adapter")
      });
      let yields = 0;
      const pacer = createPacer({ budgetMs: pacerBudgetMs, yieldFn: makeYieldFn(() => yields++) });
      const portB = createOfficeWordPort({
        runner: hB.runner,
        pureWholeBody: true,
        pacer,
        logger: hB.tracer.logger("adapter")
      });

      // PACED runs FIRST, deliberately. In-process timings are heap-state-dominated: whichever
      // hide runs second inherits the first's ballooned heap. Measured both orders (2026-06):
      // unpaced-first → 10.4s unpaced / 12.2s paced (a fake "+17%"); paced-first → 5.8s paced /
      // 5.8s unpaced (a tie — and BOTH faster, because the paced run's yields let V8 GC
      // incrementally instead of monopolizing the thread). So this order biases AGAINST the
      // pacer the least and the numbers below honestly show overhead within noise.
      const t1 = Date.now();
      const resB = await hide(portB, settings(KEEP_COLORS), { pacing: pacer });
      const pacedMs = Date.now() - t1;
      const t0 = Date.now();
      const resA = await hide(portA, settings(KEEP_COLORS));
      const unpacedMs = Date.now() - t0;

      // Byte-identity on REAL OOXML — the core guarantee that pacing can't corrupt output.
      expect(docB.paragraphs.map((p) => p.xml).join(" ")).toBe(
        docA.paragraphs.map((p) => p.xml).join(" ")
      );
      expect(resB).toEqual(resA);
      expect(yields).toBeGreaterThan(0); // the loops really sliced
      return { paras: base.length, unpacedMs, pacedMs, yields };
  }

  // Always-run correctness: the SMALLEST sample (fast), with budget 0 — a yield between
  // EVERY paragraph is the maximally-perturbed schedule, the strongest identity probe.
  const smallest = lightSamples[0];
  (smallest ? it : it.skip)(
    "byte-identical under a yield-every-tick schedule (smallest sample)",
    async () => {
      await pacedVsUnpaced(smallest!.fullPath, 0, (bump) => async () => bump());
    },
    300000
  );

  // Perf measurement: the LARGEST non-heavy sample with the PRODUCTION pacer settings (50ms
  // budget, real timer yields). Gated like realDocs' heavy tiers — in-process jest timings
  // swing seconds run-to-run (GC/JIT), so this is for deliberate measurement, not the gate.
  const largest = lightSamples[lightSamples.length - 1];
  (largest && process.env.ROSTRUM_PERF === "1" ? it : it.skip)(
    "overhead with PRODUCTION pacer settings (50ms budget, real timer yields) — ROSTRUM_PERF=1",
    async () => {
      const r = await pacedVsUnpaced(
        largest!.fullPath,
        50,
        (bump) => () =>
          new Promise<void>((resolve) =>
            setTimeout(() => {
              bump();
              resolve();
            }, 0)
          )
      );
      // eslint-disable-next-line no-console
      console.log(
        `[pacing ${largest!.tier}] ${largest!.file}: ${r.paras} paras | ` +
          `unpaced ${r.unpacedMs}ms | paced(50ms budget) ${r.pacedMs}ms | ${r.yields} yields`
      );
    },
    300000
  );
});
