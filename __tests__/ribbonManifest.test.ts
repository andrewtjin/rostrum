// Tests for the manifest generator (src/features/manifestGen.ts) and the committed manifest.xml.
// The generator turns the headless feature contributions into the ribbon; these assert that
// projection is faithful (a group + buttons per feature, no dangling FunctionNames, deep-linked
// panes) and that the committed file hasn't drifted from the generator (run `npm run gen:manifest`).
import { readFileSync } from "fs";
import { resolve } from "path";
import { buildManifestXml, manifestConfig, prodConfig, PROD_ID } from "../src/features/manifestGen";
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

describe("shared runtime (Always-On enablement, 0.3.0)", () => {
  const xml = buildManifestXml(contributions, manifestConfig);

  it("declares the SharedRuntime 1.0 requirement", () => {
    expect(xml).toContain(`<Set Name="SharedRuntime" MinVersion="1.0" />`);
  });

  it("emits a long-lived <Runtime> pointing at the taskpane runtime page", () => {
    expect(xml).toContain(`<Runtime resid="Rostrum.Taskpane.Url" lifetime="long" />`);
    expect(xml).toContain(`<bt:Url id="Rostrum.Taskpane.Url" DefaultValue="${manifestConfig.origin}/taskpane.html" />`);
  });

  it("points the FunctionFile at the runtime page (not the removed commands.html)", () => {
    expect(xml).toContain(`<FunctionFile resid="Rostrum.Taskpane.Url" />`);
    expect(xml).not.toContain("commands.html");
    expect(xml).not.toContain("Rostrum.Commands.Url");
  });

  it("bumped to the 0.3.0 shared-runtime product version", () => {
    expect(manifestConfig.version).toBe("0.3.0.0");
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

describe("prodConfig (hosted-origin override layer)", () => {
  const ORIGIN = "https://andrewtjin.github.io/rostrum";

  it("rebases EVERY hosted URL field onto the prod origin", () => {
    const cfg = prodConfig({ origin: ORIGIN });
    expect(cfg.origin).toBe(ORIGIN);
    expect(cfg.iconUrl).toBe(`${ORIGIN}/assets/icon-32.png`);
    expect(cfg.highResolutionIconUrl).toBe(`${ORIGIN}/assets/icon-80.png`);
    // No localhost survives anywhere in the override (the bug we're shipping to avoid).
    expect(JSON.stringify(cfg)).not.toContain("localhost");
  });

  it("uses PROD_ID by default — DISTINCT from the dev id so both can be sideloaded together", () => {
    const cfg = prodConfig({ origin: ORIGIN });
    expect(cfg.id).toBe(PROD_ID);
    expect(cfg.id).not.toBe(manifestConfig.id);
  });

  it("tolerates a trailing slash on the origin (no `//assets` doubling)", () => {
    const cfg = prodConfig({ origin: `${ORIGIN}/` });
    expect(cfg.origin).toBe(ORIGIN);
    expect(cfg.iconUrl).toBe(`${ORIGIN}/assets/icon-32.png`);
  });

  it("defaults support + learn-more to the repo issues page and the Pages landing root", () => {
    const cfg = prodConfig({ origin: ORIGIN });
    expect(cfg.supportUrl).toBe("https://github.com/andrewtjin/rostrum/issues");
    expect(cfg.getStarted.learnMoreUrl).toBe(`${ORIGIN}/`);
  });

  it("honors explicit overrides for id / support / learn-more", () => {
    const cfg = prodConfig({
      origin: ORIGIN,
      id: "11111111-2222-3333-4444-555555555555",
      supportUrl: "https://example.com/support",
      learnMoreUrl: "https://example.com/learn",
    });
    expect(cfg.id).toBe("11111111-2222-3333-4444-555555555555");
    expect(cfg.supportUrl).toBe("https://example.com/support");
    expect(cfg.getStarted.learnMoreUrl).toBe("https://example.com/learn");
  });

  it("does NOT mutate the shared dev manifestConfig (pure override)", () => {
    const beforeId = manifestConfig.id;
    const beforeOrigin = manifestConfig.origin;
    const beforeLearn = manifestConfig.getStarted.learnMoreUrl;
    prodConfig({ origin: ORIGIN });
    expect(manifestConfig.id).toBe(beforeId);
    expect(manifestConfig.origin).toBe(beforeOrigin);
    // The nested getStarted object must be cloned, not aliased — else the spread would leak.
    expect(manifestConfig.getStarted.learnMoreUrl).toBe(beforeLearn);
    expect(manifestConfig.origin).toBe("https://localhost:3000");
  });

  it("projects through buildManifestXml: every emitted URL is prod, none localhost", () => {
    const xml = buildManifestXml(contributions, prodConfig({ origin: ORIGIN }));
    expect(xml).toContain(`<Id>${PROD_ID}</Id>`);
    expect(xml).toContain(`DefaultValue="${ORIGIN}/taskpane.html"`);
    // The shared-runtime page (taskpane.html) is now the FunctionFile + Runtime; commands.html is gone.
    expect(xml).not.toContain("commands.html");
    expect(xml).toContain(`DefaultValue="${ORIGIN}/assets/icon-32.png"`);
    expect(xml).not.toContain("localhost");
    // Pane deep-links must also carry the prod origin.
    for (const f of contributions) {
      if (f.ribbon.controls.some((c) => c.kind === "pane")) {
        expect(xml).toContain(`${ORIGIN}/taskpane.html#${f.id}`);
      }
    }
  });
});
