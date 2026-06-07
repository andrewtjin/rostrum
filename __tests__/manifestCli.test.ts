// Tests for the gen-manifest CLI logic (tools/manifestCli.ts). The dev-vs-prod branching is the
// dangerous part: an adversarial review found that a bare `--origin`, an empty `--origin=`, or the
// `gen:manifest:prod` script run WITHOUT `--origin` would each silently write a broken or
// wrong-location manifest. These tests pin the loud-failure contract so those footguns stay fixed.
import { parseFlags, resolveManifestPlan } from "../tools/manifestCli";
import { manifestConfig, PROD_ID } from "../src/features/manifestGen";

describe("parseFlags", () => {
  it("parses --k=v (and keeps `=` inside the value, e.g. a query string)", () => {
    expect(parseFlags(["--origin=https://h/r"])).toEqual({ origin: "https://h/r" });
    expect(parseFlags(["--origin=https://h/r?a=1&b=2"])).toEqual({ origin: "https://h/r?a=1&b=2" });
  });

  it("parses --k v (space form), consuming the next non-flag token", () => {
    expect(parseFlags(["--origin", "https://h/r"])).toEqual({ origin: "https://h/r" });
  });

  it("treats a bare --flag (no value, or followed by another flag) as the string 'true'", () => {
    expect(parseFlags(["--origin"])).toEqual({ origin: "true" });
    expect(parseFlags(["--origin", "--out", "x"])).toEqual({ origin: "true", out: "x" });
  });

  it("parses --k= (equals with empty value) as the empty string", () => {
    expect(parseFlags(["--origin="])).toEqual({ origin: "" });
  });

  it("ignores positional (non --) tokens", () => {
    expect(parseFlags(["positional", "--x=1"])).toEqual({ x: "1" });
  });

  it("lets a later duplicate flag win", () => {
    expect(parseFlags(["--origin=a", "--origin=b"])).toEqual({ origin: "b" });
  });
});

describe("resolveManifestPlan — DEV (no prod signal)", () => {
  it("returns the committed-manifest dev plan and the SHARED dev config", () => {
    const plan = resolveManifestPlan({}, {});
    expect(plan.mode).toBe("dev");
    expect(plan.outRelative).toBe("manifest.xml");
    expect(plan.config).toBe(manifestConfig); // same object → byte-identical, drift-safe
  });
});

describe("resolveManifestPlan — PROD (valid origin)", () => {
  it("rebases onto the origin, defaults the out path, and uses PROD_ID", () => {
    const plan = resolveManifestPlan({ origin: "https://andrewtjin.github.io/rostrum" }, {});
    expect(plan.mode).toBe("prod");
    expect(plan.outRelative).toBe("dist/manifest.xml");
    expect(plan.config.origin).toBe("https://andrewtjin.github.io/rostrum");
    expect(plan.config.id).toBe(PROD_ID);
  });

  it("honors an explicit --out and the id/support/learn overrides", () => {
    const plan = resolveManifestPlan(
      {
        origin: "https://example.com/app",
        out: "build/manifest.xml",
        id: "11111111-2222-3333-4444-555555555555",
        support: "https://example.com/support",
        learn: "https://example.com/learn",
      },
      {}
    );
    expect(plan.outRelative).toBe("build/manifest.xml");
    expect(plan.config.id).toBe("11111111-2222-3333-4444-555555555555");
    expect(plan.config.supportUrl).toBe("https://example.com/support");
    expect(plan.config.getStarted.learnMoreUrl).toBe("https://example.com/learn");
  });

  it("reads the origin from ROSTRUM_ORIGIN env when no flag is given", () => {
    const plan = resolveManifestPlan({}, { ROSTRUM_ORIGIN: "https://h.example/app" });
    expect(plan.mode).toBe("prod");
    expect(plan.config.origin).toBe("https://h.example/app");
  });
});

describe("resolveManifestPlan — prod-intent footguns must FAIL LOUDLY (not silently degrade)", () => {
  it("throws on a bare --origin (parsed as 'true'), instead of writing `true/assets/...`", () => {
    expect(() => resolveManifestPlan({ origin: "true" }, {})).toThrow(/requires --origin/);
  });

  it("throws on an empty --origin= ", () => {
    expect(() => resolveManifestPlan({ origin: "" }, {})).toThrow(/requires --origin/);
  });

  it("throws when --out is given but --origin is missing (the `npm run gen:manifest:prod` footgun)", () => {
    // This is exactly what the prod npm script passes; without `-- --origin=…` it must NOT silently
    // fall back to a dev write at the wrong path.
    expect(() => resolveManifestPlan({ out: "dist/manifest.xml" }, {})).toThrow(/requires --origin/);
  });

  it("throws when ROSTRUM_MANIFEST_OUT is set but no origin is provided", () => {
    expect(() => resolveManifestPlan({}, { ROSTRUM_MANIFEST_OUT: "dist/manifest.xml" })).toThrow(
      /requires --origin/
    );
  });

  it("rejects a non-https origin (Office only loads https manifests)", () => {
    expect(() => resolveManifestPlan({ origin: "http://evil.example/app" }, {})).toThrow(/https/);
  });

  it("does not mutate the shared dev config on the prod path", () => {
    const before = manifestConfig.origin;
    resolveManifestPlan({ origin: "https://h.example/app" }, {});
    expect(manifestConfig.origin).toBe(before);
    expect(manifestConfig.origin).toBe("https://localhost:3000");
  });
});
