import {
  assertTrackChangesOff,
  detectFeatureSupport,
  assertCanRun,
  withTrackChangesGate,
  withPrefetchedTrackChangesGate,
  TrackChangesActiveError,
  UnsupportedHostError,
  RequirementsLike,
  TrackChangesPort
} from "../src/core/guards";
import { FeatureSupport, TrackChangesMode } from "../src/core/types";

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

// ---------------------------------------------------------------------------
// Track-Changes gate — the prefetched-mode variant (Loop 002 B2 / 002-S4) and the
// standard gate that now delegates to it (DRY: one copy of the throw/toggle/restore policy).
// ---------------------------------------------------------------------------
describe("withPrefetchedTrackChangesGate (B2 read fusion)", () => {
  /** A TC port that records every set + how many times the mode was READ (must be ZERO here). */
  class RecordingPort implements TrackChangesPort {
    reads = 0;
    readonly sets: TrackChangesMode[] = [];
    constructor(public mode: TrackChangesMode) {}
    async getChangeTrackingMode(): Promise<TrackChangesMode> {
      this.reads++;
      return this.mode;
    }
    async setChangeTrackingMode(mode: TrackChangesMode): Promise<void> {
      this.sets.push(mode);
      this.mode = mode;
    }
  }

  it("runs the body directly when the primed mode is Off — and NEVER reads the mode itself", async () => {
    const port = new RecordingPort("Off");
    const out = await withPrefetchedTrackChangesGate(port, false, "Off", async () => 42);
    expect(out).toEqual({ result: 42, toggled: false });
    expect(port.reads).toBe(0); // the whole point: no TC-read run
    expect(port.sets).toEqual([]);
  });

  it("throws BEFORE the body runs when primed mode is on and auto-toggle is off (abort-before-parse)", async () => {
    const port = new RecordingPort("TrackAll");
    let bodyRan = false;
    await expect(
      withPrefetchedTrackChangesGate(port, false, "TrackAll", async () => {
        bodyRan = true;
        return 1;
      })
    ).rejects.toBeInstanceOf(TrackChangesActiveError);
    expect(bodyRan).toBe(false); // body never ran → zero classify/writes for the engine
    expect(port.sets).toEqual([]); // no toggle either
    expect(port.reads).toBe(0);
  });

  it("toggles off, runs, and restores the primed mode in finally (S-005)", async () => {
    const port = new RecordingPort("TrackMineOnly");
    const out = await withPrefetchedTrackChangesGate(port, true, "TrackMineOnly", async () => "done");
    expect(out).toEqual({ result: "done", toggled: true });
    expect(port.sets).toEqual(["Off", "TrackMineOnly"]); // toggled off, then restored
    expect(port.reads).toBe(0);
  });

  it("restores the prior mode even when the body throws", async () => {
    const port = new RecordingPort("TrackAll");
    await expect(
      withPrefetchedTrackChangesGate(port, true, "TrackAll", async () => {
        throw new Error("boom");
      })
    ).rejects.toThrow("boom");
    expect(port.sets).toEqual(["Off", "TrackAll"]); // finally restored despite the throw
  });

  it("the standard withTrackChangesGate delegates to it (reads the mode ONCE, same outcome)", async () => {
    const port = new RecordingPort("TrackAll");
    const out = await withTrackChangesGate(port, true, async () => "x");
    expect(out).toEqual({ result: "x", toggled: true });
    expect(port.reads).toBe(1); // the standard gate issues its own one TC read
    expect(port.sets).toEqual(["Off", "TrackAll"]);
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
