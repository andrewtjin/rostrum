// Controller tests — the CondenseController on a FAKE RangeScopedPort (no Word host). These cover the
// outcomes the React layer renders: shrink/condense dispatch, the Track-Changes gate, heading refusal,
// settings-driven mode selection, the lossless round-trip through the port, and settings persistence.
import { CondenseController } from "../src/taskpane/condenseController";
import { RangeRead, RangeScopedPort, TrackChangesMode } from "../src/core/types";
import { CONDENSE_SETTINGS_KEY, StorageLike } from "../src/core/settings";
import { readFragmentParagraphs } from "../src/core/ooxmlCondense";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (inner: string): string => `<w:p>${inner}</w:p>`;
const bareP = (inner: string): string => `<w:p xmlns:w="${W_NS}">${inner}</w:p>`;
const body = (...ps: string[]): string => `<w:body xmlns:w="${W_NS}">${ps.join("")}</w:body>`;
const paraTexts = (xml: string): string[] =>
  readFragmentParagraphs(xml).map((runs) => runs.map((r) => r.text).join(""));

/** A fake range port backed by an in-memory fragment + TC mode. */
class FakeRangePort implements RangeScopedPort {
  xml: string;
  collapsed: boolean;
  outlineLevels: (number | null)[];
  tcMode: TrackChangesMode;
  writes = 0;

  constructor(
    xml: string,
    opts: { collapsed?: boolean; outlineLevels?: (number | null)[]; tc?: TrackChangesMode } = {}
  ) {
    this.xml = xml;
    this.collapsed = opts.collapsed ?? false;
    this.outlineLevels = opts.outlineLevels ?? [null];
    this.tcMode = opts.tc ?? "Off";
  }
  async getChangeTrackingMode(): Promise<TrackChangesMode> {
    return this.tcMode;
  }
  async setChangeTrackingMode(mode: TrackChangesMode): Promise<void> {
    this.tcMode = mode;
  }
  async readActiveRangeOoxml(): Promise<RangeRead> {
    return { ooxml: this.xml, collapsed: this.collapsed, outlineLevels: this.outlineLevels };
  }
  async replaceActiveRangeOoxml(xml: string): Promise<void> {
    this.xml = xml;
    this.writes++;
  }
}

function memStorage(): StorageLike {
  const m = new Map<string, string>();
  return { getItem: (k) => m.get(k) ?? null, setItem: (k, v) => void m.set(k, v) };
}

describe("shrink", () => {
  it("shrinks, writes back, and records the applied size", async () => {
    const port = new FakeRangePort(bareP(run("body text")));
    const ctrl = new CondenseController({ port, storage: memStorage() });
    const out = await ctrl.shrink();
    expect(out.status).toBe("ok");
    expect(port.writes).toBe(1);
    expect(ctrl.status().lastShrinkHalfPts).toBe(16); // 11pt → 8pt
    expect(readFragmentParagraphs(port.xml)[0][0].sizeHalfPts).toBe(16);
  });

  it("refuses a collapsed heading WITHOUT writing", async () => {
    const port = new FakeRangePort(bareP(run("Heading")), { collapsed: true, outlineLevels: [0] });
    const ctrl = new CondenseController({ port, storage: memStorage() });
    const out = await ctrl.shrink();
    expect(out.status).toBe("ok");
    if (out.status === "ok") expect(out.message).toMatch(/headings/);
    expect(port.writes).toBe(0);
  });
});

describe("Track-Changes gate", () => {
  it("returns a trackChanges outcome and does not write when TC is on", async () => {
    const port = new FakeRangePort(bareP(run("body")), { tc: "TrackAll" });
    const ctrl = new CondenseController({ port, storage: memStorage() });
    const out = await ctrl.shrink();
    expect(out).toEqual({ status: "trackChanges", mode: "TrackAll" });
    expect(port.writes).toBe(0);
  });

  it("auto-toggles past the gate and restores the prior mode", async () => {
    const port = new FakeRangePort(bareP(run("body")), { tc: "TrackAll" });
    const ctrl = new CondenseController({ port, storage: memStorage() });
    const out = await ctrl.shrink(true);
    expect(out.status).toBe("ok");
    expect(port.writes).toBe(1);
    expect(port.tcMode).toBe("TrackAll"); // restored after the op
  });
});

describe("condense — settings-driven dispatch", () => {
  it("uses pilcrows when the setting is on", async () => {
    const port = new FakeRangePort(body(p(run("AAA")), p(run("BBB"))));
    const ctrl = new CondenseController({ port, storage: memStorage() });
    ctrl.setSettings({ usePilcrows: true });
    await ctrl.condense();
    expect(port.xml).toContain("¶");
  });

  it("retain-paragraphs mode keeps structure and drops blank lines", async () => {
    const port = new FakeRangePort(body(p(run("AAA")), p(run("  ")), p(run("BBB"))));
    const ctrl = new CondenseController({ port, storage: memStorage() });
    ctrl.setSettings({ retainParagraphs: true });
    const out = await ctrl.condense();
    expect(out.status).toBe("ok");
    if (out.status === "ok") expect(out.message).toMatch(/blank line/);
    expect(readFragmentParagraphs(port.xml)).toHaveLength(3); // retained
  });

  it("explicit fullCondense merges regardless of the saved mode", async () => {
    const port = new FakeRangePort(body(p(run("AAA")), p(run("BBB")), p(run("CCC"))));
    const ctrl = new CondenseController({ port, storage: memStorage() });
    ctrl.setSettings({ retainParagraphs: true }); // saved mode is retain...
    await ctrl.fullCondense(); // ...but the explicit button merges
    expect(paraTexts(port.xml)).toEqual(["AAA BBB CCC"]);
  });
});

describe("lossless round-trip through the controller + port", () => {
  it("condense then uncondense restores the paragraphs", async () => {
    const port = new FakeRangePort(body(p(run("One")), p(run("Two")), p(run("Three"))));
    const ctrl = new CondenseController({ port, storage: memStorage() });
    await ctrl.condense();
    expect(paraTexts(port.xml)).toHaveLength(1); // merged
    await ctrl.uncondense();
    expect(paraTexts(port.xml)).toEqual(["One", "Two", "Three"]);
    expect(port.writes).toBe(2);
  });

  it("unshrink reverses shrink", async () => {
    const port = new FakeRangePort(bareP(run("body")));
    const ctrl = new CondenseController({ port, storage: memStorage() });
    await ctrl.shrink();
    expect(readFragmentParagraphs(port.xml)[0][0].sizeHalfPts).toBe(16);
    await ctrl.unshrink();
    expect(readFragmentParagraphs(port.xml)[0][0].sizeHalfPts).toBeNull();
    expect(ctrl.status().lastShrinkHalfPts).toBeNull();
  });
});

describe("settings persistence", () => {
  it("persists settings to storage and a fresh controller reads them back", async () => {
    const storage = memStorage();
    const ctrl = new CondenseController({ port: new FakeRangePort(bareP(run("x"))), storage });
    ctrl.setSettings({ usePilcrows: true, reversal: "none" });
    expect(storage.getItem(CONDENSE_SETTINGS_KEY)).toContain("usePilcrows");

    const ctrl2 = new CondenseController({ port: new FakeRangePort(bareP(run("x"))), storage });
    expect(ctrl2.getSettings().usePilcrows).toBe(true);
    expect(ctrl2.getSettings().reversal).toBe("none");
  });
});
