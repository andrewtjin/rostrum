// Task-pane orchestration tests — the RostrumController on the fake host. These
// cover the outcomes the React layer renders: arming, keep-color persistence, the
// Track-Changes prompt path, friendly timing, and Apply-Styles gating.

import { RostrumController } from "../src/taskpane/controller";
import { serializeManifest, parseManifestOrNull } from "../src/core/manifest";
import { FeatureSupport } from "../src/core/types";
import { para, run, mkDoc, harness, hiddenFlags, FakeDoc } from "./fakeWord";

const FULL: FeatureSupport = {
  canHide: true,
  canCustomXml: true,
  canChangeTracking: true,
  canStyleBorders: true,
  canStyleFormat: true,
  canGetStyles: true
};

function memStorage() {
  const m = new Map<string, string>();
  return {
    getItem: (k: string) => m.get(k) ?? null,
    setItem: (k: string, v: string) => void m.set(k, v)
  };
}

function controller(
  doc: FakeDoc = mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run("card body"))]),
  over: Partial<FeatureSupport> = {}
) {
  const h = harness(doc);
  const storage = memStorage();
  const ctrl = new RostrumController({
    features: { ...FULL, ...over },
    runner: h.runner,
    storage,
    // Force the per-paragraph path so these legacy fixtures (which carry PROXY outline levels that
    // pure whole-body ⑦ ignores) keep exercising proxy-based classification. Pure ⑦ is covered in its
    // own describe block + officeWordPortPure.test.ts.
    pureWholeBody: false,
    logger: h.tracer.logger("pane")
  });
  return { ctrl, h, doc, storage };
}

describe("init", () => {
  it("reports a fresh document as idle with all colors kept by default", async () => {
    const { ctrl } = controller();
    const status = await ctrl.init();
    expect(status.armed).toBe(false);
    expect(status.keepColors.length).toBeGreaterThan(0); // built-in default = all colors
  });

  it("reads armed state + keep-colors from an existing manifest", async () => {
    const { ctrl, doc } = controller();
    doc.manifest = {
      id: "part-1",
      xml: serializeManifest({ active: true, keepColors: ["green"], schemaVersion: 1 })
    };
    const status = await ctrl.init();
    expect(status.armed).toBe(true);
    expect(status.keepColors).toEqual(["green"]);
  });
});

describe("hide / showAll", () => {
  it("hides, arms, and reports a friendly summary", async () => {
    const { ctrl, doc } = controller();
    await ctrl.init();
    const out = await ctrl.hide();
    expect(out.status).toBe("ok");
    if (out.status === "ok") {
      expect(out.message).toMatch(/Hid 1 of 2 paragraphs/);
      expect(typeof out.tookMs).toBe("number");
    }
    expect(ctrl.status().armed).toBe(true);
    expect(parseManifestOrNull(doc.manifest!.xml)!.active).toBe(true);
  });

  it("showAll reveals and disarms", async () => {
    const { ctrl, doc } = controller();
    await ctrl.init();
    await ctrl.hide();
    const out = await ctrl.showAll();
    expect(out.status).toBe("ok");
    expect(ctrl.status().armed).toBe(false);
    expect(doc.manifest).toBeNull();
  });
});

describe("Track Changes prompt path", () => {
  it("returns a trackChanges outcome instead of throwing when TC is on", async () => {
    const doc = mkDoc([para(run("card"))], "TrackAll");
    const h = harness(doc);
    const ctrl = new RostrumController({ features: FULL, runner: h.runner, storage: memStorage() });
    await ctrl.init();
    const out = await ctrl.hide();
    expect(out).toEqual({ status: "trackChanges", mode: "TrackAll" });
  });

  it("retries past the gate when the user opts into auto-toggle", async () => {
    const doc = mkDoc([para(run("card"))], "TrackAll");
    const h = harness(doc);
    const ctrl = new RostrumController({ features: FULL, runner: h.runner, storage: memStorage() });
    await ctrl.init();
    const out = await ctrl.hide(true);
    expect(out.status).toBe("ok");
    expect(doc.tcMode).toBe("TrackAll"); // restored
  });
});

describe("keep-colors", () => {
  it("persists to device storage and writes them into the manifest on hide", async () => {
    const { ctrl, doc, storage } = controller();
    await ctrl.init();
    ctrl.setKeepColors(["green", "cyan"]);
    expect(ctrl.keepColors).toEqual(["green", "cyan"]);
    expect(storage.getItem("rostrum.deviceDefaults.v1")).toContain("green");

    await ctrl.hide();
    expect(parseManifestOrNull(doc.manifest!.xml)!.keepColors).toEqual(["green", "cyan"]);
  });
});

describe("applyStyles gating", () => {
  it("errors clearly when the host lacks the style APIs AND there are no cites to repair", async () => {
    // Default fixture (heading + plain body) has no mis-styled cite, so with sizing
    // unsupported nothing useful happens → a clear error.
    const { ctrl } = controller(undefined, { canGetStyles: false });
    await ctrl.init();
    const out = await ctrl.applyStyles();
    expect(out.status).toBe("error");
    if (out.status === "error") expect(out.message).toMatch(/can't apply Rostrum styles/);
  });

  it("repairs cites even when the size phase fails — one phase's failure never aborts the other", async () => {
    // The fake host does not model `getStyles()`, so the size phase throws here; that must NOT
    // abort cite-repair (the resilience guarantee). The op stays OK and reports the cite.
    // Schema-valid shape: outlineLvl lives inside <w:pPr> (headingLevel reads ONLY there).
    const tagPPr = `<w:pPr><w:outlineLvl w:val="3"/></w:pPr>`;
    const boldAuthor = `<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">Smith 20</w:t></w:r>`;
    const doc = mkDoc([
      para(run("Tag heading"), { pPr: tagPPr }),
      para(`${boldAuthor}<w:r><w:t xml:space="preserve"> [descriptor]</w:t></w:r>`)
    ]);
    const { ctrl } = controller(doc); // FULL features (size phase ATTEMPTED, fails on the fake)
    await ctrl.init();
    const out = await ctrl.applyStyles();
    expect(out.status).toBe("ok"); // cite-repair succeeded → not a hard error
    if (out.status === "ok") expect(out.message).toMatch(/repaired 1 cite\(s\)/);
  });

  it("still repairs cites and reports them even when sizing is unsupported", async () => {
    // Schema-valid shape: outlineLvl lives inside <w:pPr> (headingLevel reads ONLY there).
    const tagPPr = `<w:pPr><w:outlineLvl w:val="3"/></w:pPr>`;
    const boldAuthor = `<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">Smith 20</w:t></w:r>`;
    const doc = mkDoc([
      para(run("Tag heading"), { pPr: tagPPr }),
      para(`${boldAuthor}<w:r><w:t xml:space="preserve"> [descriptor]</w:t></w:r>`)
    ]);
    const { ctrl } = controller(doc, { canGetStyles: false }); // sizing unsupported
    await ctrl.init();
    const out = await ctrl.applyStyles();
    // Cite-repair only needs getOoxml/insertOoxml, so it runs → the op is OK and reports the cite.
    expect(out.status).toBe("ok");
    if (out.status === "ok") expect(out.message).toMatch(/repaired 1 cite\(s\)/);
  });
});

describe("external-change re-sync (refreshFromDocument)", () => {
  it("reflects a Hide done OUTSIDE this controller (e.g. the ribbon) after a document re-read", async () => {
    const { ctrl, storage, h, doc } = controller();
    await ctrl.init();
    expect(ctrl.status().armed).toBe(false);

    // A second controller on the SAME document/runner stands in for the ribbon command's SEPARATE
    // RostrumController instance: its Hide arms the document manifest...
    const ribbon = new RostrumController({
      features: FULL,
      runner: h.runner,
      storage,
      logger: h.tracer.logger("pane")
    });
    await ribbon.init();
    await ribbon.hide();
    expect(doc.manifest).not.toBeNull(); // the document is now armed

    // ...but THIS controller's in-memory armed flag stays stale (the green indicator bug) until it
    // re-reads the document — which the pane now does whenever it regains focus.
    expect(ctrl.status().armed).toBe(false);
    const refreshed = await ctrl.refreshFromDocument();
    expect(refreshed.armed).toBe(true);
    expect(ctrl.status().armed).toBe(true);
  });

  it("reflects an external Show All (disarm) too", async () => {
    const { ctrl, storage, h } = controller();
    await ctrl.init();
    await ctrl.hide(); // armed by THIS controller
    expect(ctrl.status().armed).toBe(true);

    // The ribbon's separate controller disarms via Show All...
    const ribbon = new RostrumController({ features: FULL, runner: h.runner, storage, logger: h.tracer.logger("pane") });
    await ribbon.init();
    await ribbon.showAll();

    // ...and the pane catches up on the next re-read.
    const refreshed = await ctrl.refreshFromDocument();
    expect(refreshed.armed).toBe(false);
  });
});

describe("re-entrancy guard — interleaved ops cannot overlap (002-F8)", () => {
  // 002-F8: B1's node-direct Hide holds a parsed, partially-mutated DOM in the port's `lastRead.pkg`
  // across the read→commit window. A SECOND concurrent op interleaving there is the half-mutated-package
  // losslessness exposure (002-F1/F4). `runMutation`/`applyStyles` already SET `inFlight`; the CHECK
  // ported here REFUSES the second op before it starts any engine work. The block surfaces through the
  // existing `status:"error"` "still running" channel (same as CondenseController.runRangeOp), so the
  // ribbon/pane render the kept-open pop-out the codebase already has (S-009), with no new UI.

  // A runner that, ONCE armed, holds the next Word.run (op A) open at a gate so op B is invoked while A
  // is provably in flight, then releases A. `arm()` is called AFTER init (so init's own runs pass
  // through); the next run after arming parks until `release()`. `engineReads()` counts whole-body reads
  // = engine invocations. (Init runs are not engine reads — it reads the manifest, not body OOXML.)
  function gatedHarness(doc: FakeDoc) {
    const h = harness(doc);
    let release!: () => void;
    const gate = new Promise<void>((res) => (release = res));
    let armed = false;
    let parked = false;
    const runner = async <T,>(b: (c: Word.RequestContext) => Promise<T>): Promise<T> => {
      if (armed && !parked) {
        parked = true; // park only the FIRST run after arming (op A); later runs pass through
        await gate;
      }
      return h.runner(b);
    };
    const engineReads = (): number => h.ctx.commitLog.filter((c) => c.op === "body.getOoxml").length;
    return {
      h,
      runner,
      release,
      engineReads,
      arm: (): void => {
        armed = true;
      }
    };
  }

  it("blocks a second Hide invoked while the first is in flight; the first still completes", async () => {
    const { h, runner, release, engineReads, arm } = gatedHarness(
      mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run("card body"))])
    );
    const ctrl = new RostrumController({
      features: FULL,
      runner,
      storage: memStorage(),
      pureWholeBody: false, // proxy path (proxy-outline fixture) — same path the other legacy tests use
      logger: h.tracer.logger("pane")
    });
    await ctrl.init();
    arm(); // from here, the next Word.run (op A's) parks at the gate

    // Op A: do NOT await. `inFlight` is set synchronously in runMutation before the engine is awaited,
    // and the gated runner parks A's first Word.run, so A is genuinely in flight from here on.
    const opA = ctrl.hide();

    // Op B: invoked WHILE A is in flight → must be refused with the blocked outcome and run NO engine
    // work (no second concurrent mutation). Awaiting B resolves immediately (it never touches the host).
    const outB = await ctrl.hide();
    expect(outB).toEqual({
      status: "error",
      message: "Another Rostrum operation is still running — let it finish first."
    });
    // The guard fired before any engine work for B: A has not even read yet (still parked at the gate),
    // so the read count is still 0 — B definitively did not start a second mutation.
    expect(engineReads()).toBe(0);

    // Release A and let it finish: it completes normally (exactly one engine read+commit), arming the doc.
    release();
    const outA = await opA;
    expect(outA.status).toBe("ok");
    if (outA.status === "ok") expect(outA.message).toMatch(/Hid 1 of 2 paragraphs/);
    expect(engineReads()).toBe(1); // exactly ONE whole-body read total — A's; B never read
    expect(ctrl.status().armed).toBe(true);

    // After A settles the guard is clear again — a subsequent op runs normally (flag cleared in finally).
    const outC = await ctrl.showAll();
    expect(outC.status).toBe("ok");
    expect(ctrl.status().armed).toBe(false);
  });

  it("blocks applyStyles while a Hide is in flight (the guard covers applyStyles too)", async () => {
    const { h, runner, release, arm } = gatedHarness(
      mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run("card body"))])
    );
    const ctrl = new RostrumController({
      features: FULL,
      runner,
      storage: memStorage(),
      pureWholeBody: false,
      logger: h.tracer.logger("pane")
    });
    await ctrl.init();
    arm();

    const opA = ctrl.hide(); // parked at the gate, in flight
    const outB = await ctrl.applyStyles(); // second op, different verb → still blocked
    expect(outB).toEqual({
      status: "error",
      message: "Another Rostrum operation is still running — let it finish first."
    });

    release();
    expect((await opA).status).toBe("ok");
  });
});

describe("Show All — single whole-doc reveal (performance default)", () => {
  it("reveals via ONE whole-doc font.hidden clear, not a per-paragraph loop", async () => {
    // The default commit strategy is whole-body, so Show All is one whole-doc font.hidden clear.
    const { ctrl, h, doc } = controller();
    await ctrl.init();
    await ctrl.hide();
    h.ctx.commitLog.length = 0;
    await ctrl.showAll();
    expect(h.ctx.commitLog.some((c) => c.op === "body.font.hidden")).toBe(true);
    expect(doc.manifest).toBeNull();
  });
});

describe("pure-whole-body (avenue ⑦) — the DEFAULT Hide path", () => {
  const lvl0 = `<w:pPr><w:outlineLvl w:val="0"/></w:pPr>`; // package-resolvable heading

  it("defaults ON in production (no override, fresh storage)", async () => {
    const h = harness(mkDoc([para(run("card"))]));
    const ctrl = new RostrumController({
      features: FULL,
      runner: h.runner,
      storage: memStorage(),
      logger: h.tracer.logger("pane")
    });
    await ctrl.init();
    expect(ctrl.status().pureWholeBody).toBe(true);
    expect(ctrl.pureWholeBodyOn).toBe(true);
  });

  it("persists an explicit opt-OUT and restores it next session", async () => {
    const h = harness(mkDoc([para(run("card"))]));
    const storage = memStorage();
    const ctrl = new RostrumController({ features: FULL, runner: h.runner, storage, logger: h.tracer.logger("pane") });
    await ctrl.init();
    expect(ctrl.status().pureWholeBody).toBe(true); // default on

    ctrl.setPureWholeBody(false);
    expect(ctrl.status().pureWholeBody).toBe(false);
    expect(storage.getItem("rostrum.pureWholeBody.v1")).toBe("false");

    // Next session on the SAME device storage reads the opt-out.
    const ctrl2 = new RostrumController({ features: FULL, runner: h.runner, storage, logger: h.tracer.logger("pane") });
    expect((await ctrl2.init()).pureWholeBody).toBe(false);
  });

  it("ignores a toggle while an operation is in flight (no mid-op adapter rebuild)", async () => {
    const { ctrl } = controller(); // helper forces pure OFF
    await ctrl.init();
    const inFlight = ctrl.hide(); // do NOT await
    ctrl.setPureWholeBody(true);
    await inFlight;
    expect(ctrl.status().pureWholeBody).toBe(false); // toggle refused mid-op
  });

  it("commits the whole body in ONE insertOoxml by default (heading kept, body hidden)", async () => {
    const h = harness(mkDoc([para(run("Heading"), { pPr: lvl0 }), para(run("card body"))]));
    const ctrl = new RostrumController({
      features: FULL,
      runner: h.runner,
      storage: memStorage(),
      logger: h.tracer.logger("pane")
    });
    await ctrl.init(); // pure ON by default — no toggle needed
    await ctrl.hide();

    expect(h.ctx.commitLog.filter((c) => c.op === "body.insertOoxml")).toHaveLength(1);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.insertOoxml")).toBe(false);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.font.hidden")).toBe(false);
    expect(hiddenFlags(h.doc.paragraphs[0].xml)).toEqual([false]); // heading kept (package outline)
    expect(hiddenFlags(h.doc.paragraphs[1].xml)).toEqual([true]); // body hidden
  });
});
