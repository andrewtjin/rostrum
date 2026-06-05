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
    const tagPPr = `<w:outlineLvl w:val="3"/>`;
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
    const tagPPr = `<w:outlineLvl w:val="3"/>`;
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
