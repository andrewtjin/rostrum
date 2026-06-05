import {
  assertTrackChangesOff,
  detectFeatureSupport,
  assertCanRun,
  TrackChangesActiveError,
  UnsupportedHostError,
  RequirementsLike
} from "../src/core/guards";
import { FeatureSupport } from "../src/core/types";

/** Fake `Office.context.requirements`: supports exactly the listed "name version" pairs. */
class FakeRequirements implements RequirementsLike {
  private supported: Set<string>;
  constructor(pairs: string[]) {
    this.supported = new Set(pairs);
  }
  isSetSupported(name: string, version?: string): boolean {
    return this.supported.has(`${name} ${version ?? ""}`.trim());
  }
}

describe("assertTrackChangesOff (decision #14)", () => {
  it("passes when Track Changes is off", () => {
    expect(() => assertTrackChangesOff("Off")).not.toThrow();
  });

  it.each(["TrackAll", "TrackMineOnly"] as const)("throws when mode is %s", (mode) => {
    try {
      assertTrackChangesOff(mode);
      throw new Error("expected to throw");
    } catch (e) {
      expect(e).toBeInstanceOf(TrackChangesActiveError);
      expect((e as TrackChangesActiveError).mode).toBe(mode);
    }
  });
});

describe("detectFeatureSupport (decision #18)", () => {
  it("maps every capability when the full surface is present", () => {
    const req = new FakeRequirements([
      "WordApiDesktop 1.2",
      "WordApi 1.4",
      "WordApiDesktop 1.1",
      "WordApi 1.5",
      "WordApiDesktop 1.4"
    ]);
    expect(detectFeatureSupport(req)).toEqual<FeatureSupport>({
      canHide: true,
      canCustomXml: true,
      canChangeTracking: true,
      canStyleBorders: true,
      canStyleFormat: true,
      canGetStyles: true
    });
  });

  it("reports a web-like host (no desktop sets) as unable to hide", () => {
    const req = new FakeRequirements(["WordApi 1.4", "WordApi 1.5"]);
    const support = detectFeatureSupport(req);
    expect(support.canHide).toBe(false);
    expect(support.canStyleBorders).toBe(false);
    expect(support.canCustomXml).toBe(true);
  });

  // Regression (LESSONS #42): `Document.getStyles()` is WordApi 1.5, not WordApiDesktop
  // 1.4. A real current-Word build can report WordApi 1.5 (canStyleFormat:true) yet lack
  // WordApiDesktop 1.4 — the exact host where Apply Styles "did nothing". canGetStyles must
  // come from WordApi 1.5 so Apply Styles isn't falsely disabled on a capable host.
  it("derives canGetStyles from WordApi 1.5, not WordApiDesktop 1.4", () => {
    const req = new FakeRequirements(["WordApiDesktop 1.2", "WordApi 1.4", "WordApi 1.5"]);
    const support = detectFeatureSupport(req);
    expect(support.canGetStyles).toBe(true);
    expect(support.canStyleFormat).toBe(true);
    // Sanity: the desktop-only set is genuinely absent on this host shape.
    expect(req.isSetSupported("WordApiDesktop", "1.4")).toBe(false);
  });
});

describe("assertCanRun", () => {
  const support = (over: Partial<FeatureSupport>): FeatureSupport => ({
    canHide: true,
    canCustomXml: true,
    canChangeTracking: true,
    canStyleBorders: true,
    canStyleFormat: true,
    canGetStyles: true,
    ...over
  });

  it("passes when hide + custom XML are supported", () => {
    expect(() => assertCanRun(support({}))).not.toThrow();
  });

  it("throws UnsupportedHostError when hiding is unavailable", () => {
    expect(() => assertCanRun(support({ canHide: false }))).toThrow(UnsupportedHostError);
  });

  it("throws UnsupportedHostError when custom XML is unavailable", () => {
    expect(() => assertCanRun(support({ canCustomXml: false }))).toThrow(UnsupportedHostError);
  });
});
