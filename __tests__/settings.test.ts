import {
  resolveSettings,
  loadDeviceDefaults,
  saveDeviceDefaults,
  loadPureWholeBody,
  savePureWholeBody,
  DEFAULT_KEEP_COLORS,
  DEVICE_DEFAULTS_KEY,
  PURE_WHOLE_BODY_KEY,
  StorageLike
} from "../src/core/settings";
import { RostrumManifest } from "../src/core/types";

/** In-memory localStorage stand-in. */
class FakeStorage implements StorageLike {
  private map = new Map<string, string>();
  getItem(key: string): string | null {
    return this.map.has(key) ? this.map.get(key)! : null;
  }
  setItem(key: string, value: string): void {
    this.map.set(key, value);
  }
}

const manifest = (keepColors: string[]): RostrumManifest => ({
  active: true,
  keepColors,
  schemaVersion: 1
});

describe("resolveSettings precedence (decision #15)", () => {
  it("uses the document manifest when present", () => {
    const r = resolveSettings(manifest(["yellow"]), { keepColors: ["green"] });
    expect([...r.keepColors]).toEqual(["yellow"]);
  });

  it("falls back to device defaults when there is no manifest", () => {
    const r = resolveSettings(null, { keepColors: ["green"] });
    expect([...r.keepColors]).toEqual(["green"]);
  });

  it("falls back to all highlight colors when nothing is set", () => {
    const r = resolveSettings(null, null);
    expect(r.keepColors.size).toBe(DEFAULT_KEEP_COLORS.length);
    expect(r.keepColors.has("yellow")).toBe(true);
  });

  it("honors an explicit empty keep-set in the manifest", () => {
    const r = resolveSettings(manifest([]), { keepColors: ["green"] });
    expect(r.keepColors.size).toBe(0);
  });

  it("normalizes (lower-cases + drops unknown) manifest colors", () => {
    const r = resolveSettings(manifest(["Yellow", "bogus"]), null);
    expect([...r.keepColors]).toEqual(["yellow"]);
  });
});

describe("device defaults cache", () => {
  it("returns null when nothing is stored", () => {
    expect(loadDeviceDefaults(new FakeStorage())).toBeNull();
  });

  it("returns null on malformed JSON", () => {
    const s = new FakeStorage();
    s.setItem(DEVICE_DEFAULTS_KEY, "{not json");
    expect(loadDeviceDefaults(s)).toBeNull();
  });

  it("returns null when keepColors is not an array", () => {
    const s = new FakeStorage();
    s.setItem(DEVICE_DEFAULTS_KEY, JSON.stringify({ keepColors: "yellow" }));
    expect(loadDeviceDefaults(s)).toBeNull();
  });

  it("round-trips normalized defaults through storage", () => {
    const s = new FakeStorage();
    saveDeviceDefaults(s, { keepColors: ["Yellow", "bogus", "green"] });
    expect(loadDeviceDefaults(s)).toEqual({ keepColors: ["yellow", "green"] });
  });
});

describe("pure-whole-body flag (avenue ⑦, per-device) — the default Hide path", () => {
  it("honors the caller's default when unset (production passes true)", () => {
    expect(loadPureWholeBody(new FakeStorage())).toBe(false); // function default
    expect(loadPureWholeBody(new FakeStorage(), true)).toBe(true); // production default ⇒ ⑦ on
  });

  it("an explicitly stored value WINS over the default (an opt-out sticks)", () => {
    const s = new FakeStorage();
    savePureWholeBody(s, false);
    expect(s.getItem(PURE_WHOLE_BODY_KEY)).toBe("false");
    expect(loadPureWholeBody(s, true)).toBe(false); // user opt-out beats the default-on
    savePureWholeBody(s, true);
    expect(loadPureWholeBody(s, false)).toBe(true);
  });

  it("returns the default when storage throws on read", () => {
    const throwing: StorageLike = {
      getItem() {
        throw new Error("storage blocked");
      },
      setItem() {
        throw new Error("storage blocked");
      }
    };
    expect(loadPureWholeBody(throwing, true)).toBe(true);
  });
});
