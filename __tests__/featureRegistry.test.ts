// Unit tests for the feature contribution model + the assembled default registry. These
// run in Node (no DOM): they assert the registry's pure logic and the shape of the LIVE
// suite — that Invisibility Mode is feature #1 with the four ribbon commands, the planned
// tools are registered + capability-gated, and command ids are globally unique (so the
// ribbon can associate each to its manifest FunctionName unambiguously).
import { FeatureRegistry } from "../src/features/registry";
import { RostrumFeature } from "../src/features/types";
import { registry } from "../src/features";
import { contributions } from "../src/features/contributions";
import { FeatureSupport } from "../src/core/types";

const desktopCaps: FeatureSupport = {
  canHide: true,
  canCustomXml: true,
  canChangeTracking: true,
  canStyleBorders: true,
  canStyleFormat: true,
  canGetStyles: true,
};
const webCaps: FeatureSupport = {
  ...desktopCaps,
  canHide: false, // Word-for-web: no font.hidden (WordApiDesktop 1.2)
  canStyleFormat: false,
  canGetStyles: false,
};

/** Build a throwaway feature with only the fields a test cares about. */
function feature(id: string, extra: Partial<RostrumFeature> = {}): RostrumFeature {
  return {
    id,
    title: id,
    tagline: "",
    glyph: "",
    status: "planned",
    primarySurface: "pane",
    isAvailable: () => true,
    ribbon: { label: id, controls: [] },
    commands: [],
    ...extra,
  };
}

describe("FeatureRegistry", () => {
  it("registers, lists in registration order, and looks up by id", () => {
    const r = new FeatureRegistry();
    r.register(feature("a")).register(feature("b"));
    expect(r.all().map((f) => f.id)).toEqual(["a", "b"]);
    expect(r.get("b")?.id).toBe("b");
    expect(r.get("missing")).toBeUndefined();
  });

  it("rejects a duplicate feature id", () => {
    const r = new FeatureRegistry();
    r.register(feature("dup"));
    expect(() => r.register(feature("dup"))).toThrow(/duplicate feature id/);
  });

  it("rejects a duplicate command id across features", () => {
    const r = new FeatureRegistry();
    const cmd = { id: "shared", title: "x", run: async () => ({ status: "noop" as const }) };
    r.register(feature("f1", { commands: [cmd] }));
    expect(() => r.register(feature("f2", { commands: [cmd] }))).toThrow(/duplicate command id/);
  });

  it("filters features by host capability", () => {
    const r = new FeatureRegistry();
    r.register(feature("desktopOnly", { isAvailable: (f) => f.canHide }));
    r.register(feature("anywhere", { isAvailable: () => true }));
    expect(r.available(desktopCaps).map((f) => f.id)).toEqual(["desktopOnly", "anywhere"]);
    expect(r.available(webCaps).map((f) => f.id)).toEqual(["anywhere"]);
  });

  it("flattens commands across features and finds one by id", () => {
    const r = new FeatureRegistry();
    const c1 = { id: "c1", title: "1", run: async () => ({ status: "ok" as const }) };
    const c2 = { id: "c2", title: "2", run: async () => ({ status: "ok" as const }) };
    r.register(feature("f", { commands: [c1, c2] }));
    expect(r.commands().map((c) => c.id)).toEqual(["c1", "c2"]);
    expect(r.command("c2")?.title).toBe("2");
    expect(r.command("missing")).toBeUndefined();
  });
});

describe("default suite registry", () => {
  it("registers Invisibility Mode as feature #1, stable, with the three ribbon commands", () => {
    expect(registry.all()[0]?.id).toBe("invisibility");
    const invis = registry.get("invisibility");
    expect(invis?.status).toBe("stable");
    // Re-hide was folded into Hide (idempotent → pressing Hide again re-hides), so the ribbon
    // surfaces three commands. The engine/controller `reHide()` stays and is tested separately.
    expect(invis?.commands.map((c) => c.id).sort()).toEqual(
      ["applyStyles", "hide", "showAll"].sort()
    );
  });

  it("gates Invisibility Mode on desktop hidden-text support", () => {
    const invis = registry.get("invisibility")!;
    expect(invis.isAvailable(desktopCaps)).toBe(true);
    expect(invis.isAvailable(webCaps)).toBe(false);
  });

  it("registers the planned suite tools as capability-gated slots", () => {
    for (const id of ["format", "flow", "cite"]) {
      const f = registry.get(id);
      expect(f).toBeTruthy();
      expect(f?.status).toBe("planned");
      // Planned tools list everywhere but aren't yet actionable on any host.
      expect(f?.isAvailable(desktopCaps)).toBe(false);
    }
  });

  it("keeps every command id globally unique (ribbon association is unambiguous)", () => {
    const ids = registry.commands().map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("gives every feature a non-empty ribbon group whose action controls map to real commands", () => {
    for (const f of registry.all()) {
      expect(f.ribbon.controls.length).toBeGreaterThan(0);
      const ids = new Set(f.commands.map((c) => c.id));
      for (const control of f.ribbon.controls) {
        // No dangling ribbon button: every ExecuteFunction control names a real command id.
        if (control.kind === "action") expect(ids.has(control.commandId)).toBe(true);
      }
    }
  });

  it("only gives a pane ribbon control to a feature that has a panel or is planned (ComingSoon)", () => {
    for (const f of registry.all()) {
      const hasPaneControl = f.ribbon.controls.some((c) => c.kind === "pane");
      if (hasPaneControl) expect(Boolean(f.panel) || f.status === "planned").toBe(true);
    }
  });
});

// The headless `contributions` (consumed by the ribbon + manifest generator) and the React
// `registry` (consumed by the panes) are assembled independently — guard against silent drift.
describe("contributions ↔ registry parity", () => {
  it("lists the same feature ids in the same order", () => {
    expect(contributions.map((c) => c.id)).toEqual(registry.all().map((f) => f.id));
  });

  it("keeps command ids globally unique across the headless contributions too", () => {
    const ids = contributions.flatMap((c) => c.commands.map((cmd) => cmd.id));
    expect(new Set(ids).size).toBe(ids.length);
  });
});
