import {
  serializeManifest,
  parseManifest,
  parseManifestOrNull,
  loadManifest,
  saveManifest,
  isArmed,
  MANIFEST_NAMESPACE,
  MANIFEST_SCHEMA_VERSION
} from "../src/core/manifest";
import {
  ParagraphUpdate,
  RawParagraph,
  RostrumManifest,
  TrackChangesMode,
  WordPort
} from "../src/core/types";

/** Minimal WordPort that only exercises the manifest part. */
class ManifestPort implements WordPort {
  manifestXml: string | null = null;
  async getChangeTrackingMode(): Promise<TrackChangesMode> {
    return "Off";
  }
  async setChangeTrackingMode(): Promise<void> {}
  async readParagraphs(): Promise<RawParagraph[]> {
    return [];
  }
  async writeParagraphs(_updates: ParagraphUpdate[]): Promise<void> {}
  async readManifest(): Promise<string | null> {
    return this.manifestXml;
  }
  async writeManifest(xml: string): Promise<void> {
    this.manifestXml = xml;
  }
  async clearManifest(): Promise<void> {
    this.manifestXml = null;
  }
  async clearHidden(): Promise<{ paragraphsScanned: number; paragraphsChanged: number }> {
    return { paragraphsScanned: 0, paragraphsChanged: 0 };
  }
}

describe("serialize / parse round-trip", () => {
  it("round-trips active + keep-colors + schema version", () => {
    const m: RostrumManifest = {
      active: true,
      keepColors: ["yellow", "green"],
      schemaVersion: MANIFEST_SCHEMA_VERSION
    };
    const parsed = parseManifest(serializeManifest(m));
    expect(parsed).toEqual(m);
  });

  it("emits the Rostrum namespace", () => {
    const xml = serializeManifest({ active: false, keepColors: [], schemaVersion: 1 });
    expect(xml).toContain(`xmlns="${MANIFEST_NAMESPACE}"`);
  });
});

describe("parseManifest", () => {
  const wrap = (inner: string, attrs = `schemaVersion="2"`): string =>
    `<rostrum xmlns="${MANIFEST_NAMESPACE}" ${attrs}>${inner}</rostrum>`;

  it("lower-cases, de-dupes, and drops unknown colors", () => {
    const xml = wrap(
      `<active>true</active><keepColors><color>Yellow</color><color>bogus</color><color>yellow</color><color>green</color></keepColors>`
    );
    const m = parseManifest(xml);
    expect(m.keepColors).toEqual(["yellow", "green"]);
    expect(m.schemaVersion).toBe(2);
    expect(m.active).toBe(true);
  });

  it("flags a missing schema version as unknown (0) for future migration guards", () => {
    const m = parseManifest(`<rostrum xmlns="${MANIFEST_NAMESPACE}"><active>true</active></rostrum>`);
    expect(m.schemaVersion).toBe(0);
  });

  it.each([
    ["<active>false</active>", false],
    ["<active>0</active>", false],
    ["", false],
    ["<active>true</active>", true],
    ["<active>1</active>", true]
  ])("reads active from %s", (inner, expected) => {
    expect(parseManifest(wrap(inner)).active).toBe(expected);
  });

  it("throws on a non-manifest root", () => {
    expect(() => parseManifest("<other/>")).toThrow();
  });
});

describe("parseManifestOrNull", () => {
  it.each([[null], [""], ["   "], ["<other/>"], ["not xml at all <"]])(
    "returns null for %s",
    (input) => {
      expect(parseManifestOrNull(input as string | null)).toBeNull();
    }
  );
});

describe("WordPort adapter", () => {
  it("saves, loads, and reports armed state", async () => {
    const port = new ManifestPort();
    expect(await loadManifest(port)).toBeNull();
    expect(await isArmed(port)).toBe(false);

    await saveManifest(port, { active: true, keepColors: ["yellow"], schemaVersion: 1 });
    const loaded = await loadManifest(port);
    expect(loaded!.active).toBe(true);
    expect(loaded!.keepColors).toEqual(["yellow"]);
    expect(await isArmed(port)).toBe(true);
  });

  it("is not armed when active is false", async () => {
    const port = new ManifestPort();
    await saveManifest(port, { active: false, keepColors: [], schemaVersion: 1 });
    expect(await isArmed(port)).toBe(false);
  });
});
