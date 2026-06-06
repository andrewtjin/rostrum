// Tests for the manifest generator (src/features/manifestGen.ts) and the committed manifest.xml.
// The generator turns the headless feature contributions into the ribbon; these assert that
// projection is faithful (a group + buttons per feature, no dangling FunctionNames, deep-linked
// panes) and that the committed file hasn't drifted from the generator (run `npm run gen:manifest`).
import { readFileSync } from "fs";
import { resolve } from "path";
import { buildManifestXml, manifestConfig } from "../src/features/manifestGen";
import { contributions } from "../src/features/contributions";
import { FeatureContribution } from "../src/features/types";

/** A throwaway contribution for exercising the generator's escaping / id logic in isolation. */
function syntheticFeature(over: Partial<FeatureContribution> = {}): FeatureContribution {
  return {
    id: "synthetic",
    title: "Synthetic",
    tagline: "",
    glyph: "",
    status: "stable",
    primarySurface: "pane",
    isAvailable: () => true,
    ribbon: { label: "Synthetic", controls: [] },
    commands: [],
    ...over,
  };
}

describe("buildManifestXml", () => {
  const xml = buildManifestXml(contributions, manifestConfig);

  it("emits exactly one <Group> per feature", () => {
    for (const f of contributions) {
      expect(xml).toContain(`<Group id="Rostrum.Group.${f.id}">`);
    }
    expect((xml.match(/<Group /g) ?? []).length).toBe(contributions.length);
  });

  it("emits a <FunctionName> for every action control and never a dangling one", () => {
    for (const f of contributions) {
      const commandIds = new Set(f.commands.map((c) => c.id));
      for (const control of f.ribbon.controls) {
        if (control.kind === "action") {
          // The button names a real command id (== the manifest FunctionName the ribbon associates).
          expect(commandIds.has(control.commandId)).toBe(true);
          expect(xml).toContain(`<FunctionName>${control.commandId}</FunctionName>`);
        }
      }
    }
  });

  it("deep-links every pane control to its own taskpane URL + per-feature TaskpaneId", () => {
    for (const f of contributions) {
      const hasPane = f.ribbon.controls.some((c) => c.kind === "pane");
      if (hasPane) {
        expect(xml).toContain(`<TaskpaneId>Rostrum.Tp.${f.id}</TaskpaneId>`);
        expect(xml).toContain(`${manifestConfig.origin}/taskpane.html#${f.id}`);
      }
    }
  });

  it("gives every control a label + tip resource", () => {
    const controlCount = contributions.flatMap((f) => f.ribbon.controls).length;
    // `\bid="..."` matches the resource DEFINITIONS, not the `resid="..."` references.
    expect((xml.match(/\bid="Lbl\d+"/g) ?? []).length).toBe(controlCount);
    expect((xml.match(/\bid="Tip\d+"/g) ?? []).length).toBe(controlCount);
  });

  it("keeps every resid reference within Office's 32-char cap", () => {
    // Every resid points at a bt:* resource (all capped at 32); element ids use `id=` and are
    // allowed up to 125, so we check the references, which transitively cover the definitions.
    const resids = [...xml.matchAll(/resid="([^"]+)"/g)].map((m) => m[1]);
    expect(resids.length).toBeGreaterThan(0);
    for (const r of resids) expect(r.length).toBeLessThanOrEqual(32);
  });

  it("bumps the version (ribbon structure changed → Office must re-register)", () => {
    expect(xml).toContain(`<Version>${manifestConfig.version}</Version>`);
    expect(manifestConfig.version).not.toBe("1.0.0.1");
  });
});

describe("buildManifestXml — robustness", () => {
  const xml = buildManifestXml(contributions, manifestConfig);

  it("escapes XML metacharacters in labels and tips (no invalid XML from a stray & or <)", () => {
    const f = syntheticFeature({
      ribbon: {
        label: 'A & B <C> "D" \'E\'',
        controls: [{ kind: "action", commandId: "x", label: "L & <b>", tip: 'T "q" & <r>' }],
      },
      commands: [{ id: "x", title: "x", run: async () => ({ status: "ok" }) }],
    });
    const out = buildManifestXml([f], manifestConfig);
    expect(out).toContain("A &amp; B &lt;C&gt; &quot;D&quot; &apos;E&apos;");
    expect(out).toContain("L &amp; &lt;b&gt;");
    expect(out).not.toContain("A & B <C>"); // the raw, unescaped form must not survive
  });

  it("generates unique resource ids (no resid collision across the counters)", () => {
    const ids = [...xml.matchAll(/<bt:(?:String|Url|Image) id="([^"]+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("has no dangling resid — every resid reference resolves to a defined resource", () => {
    const defined = new Set([...xml.matchAll(/<bt:(?:String|Url|Image) id="([^"]+)"/g)].map((m) => m[1]));
    const referenced = [...xml.matchAll(/resid="([^"]+)"/g)].map((m) => m[1]);
    expect(referenced.length).toBeGreaterThan(0);
    for (const r of referenced) expect(defined.has(r)).toBe(true);
  });
});

describe("manifest.xml drift guard", () => {
  it("the committed manifest.xml equals the generator output — run `npm run gen:manifest`", () => {
    const committed = readFileSync(resolve(__dirname, "..", "manifest.xml"), "utf8");
    const generated = buildManifestXml(contributions, manifestConfig);
    // Normalize line endings so git autocrlf can't trip the byte comparison.
    const norm = (s: string): string => s.replace(/\r\n/g, "\n");
    expect(norm(committed)).toBe(norm(generated));
  });
});
