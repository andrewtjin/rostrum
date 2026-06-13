// gdocsBuild.test.ts — the Code.gs build pipeline (plan D9/S12 + A11.vii).
//
// WHAT IS PROVEN HERE:
//   * the pure post-processing helpers exported by tools/build-gdocs.mjs
//     (emitGlobalShims / assertNoModuleTokens / makeBanner / call-map
//     extraction) behave on happy AND failure paths;
//   * buildTo(dir) — run programmatically into a TEMP dir, never against the
//     gitignored gdocs/dist/ (plan A11.vii: committed-state tests must not
//     depend on uncommitted artifacts) — produces a non-empty Code.gs with
//     zero module tokens, the version banner, no tool attribution, one
//     top-level shim per core/adapterPure.ts call-map entry, the STRINGS menu
//     copy, and a byte-exact appsscript.json alongside.
//
// HOW THE .mjs IS LOADED (documented per S12 because it is non-obvious):
// ts-jest compiles this file to CommonJS, downleveling `import()` to
// `require()`, and jest's CJS loader cannot execute real ESM (.mjs) — while
// jest's sandbox lacks --experimental-vm-modules, so an in-process native
// dynamic import is also unavailable (and package.json/jest config are frozen
// for this change). The TRUE `await import()` of the build module therefore
// happens inside a one-shot Node child process (the same way `npm run
// build:gdocs` executes it in production), driven from these async tests with
// JSON in/out. Helper calls stay cheap (~0.1s each) and error messages round-
// trip verbatim, so `expect(...).toThrow(/.../)`` works unchanged.

import { execFileSync } from "child_process";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { pathToFileURL } from "url";
import { GDOCS_VERSION } from "../gdocs/src/core/constants";
import { STRINGS } from "../gdocs/src/core/strings";

// ---------------------------------------------------------------------------
// Child-process driver for tools/build-gdocs.mjs (see header comment).
// ---------------------------------------------------------------------------

const MJS_PATH = path.resolve(__dirname, "../tools/build-gdocs.mjs");
const MJS_URL = pathToFileURL(MJS_PATH).href;

/** What buildTo(dir) resolves with (mirrors the .mjs contract). */
interface GdocsBuildResult {
  outDir: string;
  codeGsPath: string;
  appsscriptJsonPath: string;
  codeGsBytes: number;
  shimNames: string[];
  version: string;
}

/**
 * The child program: dynamic-imports the build module, calls one exported
 * helper with JSON args (read from a file — env/argv would cap large inputs
 * like a whole Code.gs text), and prints a JSON envelope. Failures are
 * CAUGHT and reported in-band so the parent can rethrow with the original
 * message — that keeps toThrow(/…/) assertions exact.
 */
const DRIVER_SRC = [
  'import { readFileSync } from "node:fs";',
  "let payload;",
  "try {",
  "  const mod = await import(process.env.GDOCS_BUILD_MJS_URL);",
  "  const fn = mod[process.env.GDOCS_BUILD_FN];",
  '  if (typeof fn !== "function") throw new Error("build module has no export named " + process.env.GDOCS_BUILD_FN);',
  '  const args = JSON.parse(readFileSync(process.env.GDOCS_BUILD_ARGS_FILE, "utf8"));',
  "  const result = await fn(...args);",
  "  payload = { ok: true, result: result === undefined ? null : result };",
  "} catch (e) {",
  "  payload = { ok: false, error: e instanceof Error ? e.message : String(e) };",
  "}",
  "process.stdout.write(JSON.stringify(payload));"
].join("\n");

/** Suite-wide temp root: holds the per-call args files AND the build output
 * dir. Created once, removed in afterAll (Windows-safe: children have exited
 * by then, so nothing holds locks). */
let tmpRoot = "";
let argsFileCounter = 0;

/**
 * Call one exported helper of build-gdocs.mjs in a child Node process.
 * Synchronous on purpose: execFileSync keeps assertion style plain
 * (expect(() => …).toThrow()) without rejects plumbing everywhere.
 */
function callBuild(fnName: string, ...args: unknown[]): unknown {
  const argsFile = path.join(tmpRoot, `args-${argsFileCounter++}.json`);
  fs.writeFileSync(argsFile, JSON.stringify(args), "utf8");
  const stdout = execFileSync(process.execPath, ["--input-type=module", "-e", DRIVER_SRC], {
    env: {
      ...process.env,
      GDOCS_BUILD_MJS_URL: MJS_URL,
      GDOCS_BUILD_FN: fnName,
      GDOCS_BUILD_ARGS_FILE: argsFile
    },
    encoding: "utf8",
    // A whole Code.gs can round-trip through the envelope; default 1MB is
    // too close for comfort on a bundle that grows with the engine.
    maxBuffer: 32 * 1024 * 1024,
    windowsHide: true
  });
  const payload = JSON.parse(stdout) as { ok: boolean; result?: unknown; error?: string };
  if (!payload.ok) throw new Error(payload.error ?? "unknown build-module failure");
  return payload.result;
}

/**
 * Load core/adapterPure.ts (a normal jest-transformed TS module). The
 * specifier is deliberately NON-literal: Wave D lands the adapter module in
 * PARALLEL with this build module (LOOPS.md dispatch record), and a literal
 * import would fail root tsc while S11 is still in flight. With the variable
 * specifier, tsc stays green either way; at runtime this test fails with the
 * honest message below until the file lands, then goes green untouched.
 * (ts-jest downlevels this import() to require(), which is exactly right for
 * a .ts module inside jest's transform pipeline.)
 */
const ADAPTER_PURE_SPECIFIER = "../gdocs/src/core/adapterPure";
async function loadAdapterPureNamespace(): Promise<Record<string, unknown>> {
  try {
    return (await import(ADAPTER_PURE_SPECIFIER)) as Record<string, unknown>;
  } catch (e) {
    throw new Error(
      "gdocs/src/core/adapterPure.ts is not loadable yet (plan S11 — Wave D adapter module): " +
        "the call-map/shim cross-check cannot run until it lands. Original error: " +
        String(e)
    );
  }
}

beforeAll(() => {
  tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), "gdocs-build-test-"));
});

afterAll(() => {
  if (tmpRoot) fs.rmSync(tmpRoot, { recursive: true, force: true });
});

// ---------------------------------------------------------------------------
// Module surface — the documented contract this repo's npm script and these
// tests both depend on (the prompt-level requirement that buildTo(dir) is an
// export, not CLI-only behavior).
// ---------------------------------------------------------------------------

describe("build-gdocs.mjs module surface", () => {
  it("exports the documented helper surface including buildTo", () => {
    const lister = [
      "const mod = await import(process.env.GDOCS_BUILD_MJS_URL);",
      "process.stdout.write(JSON.stringify(Object.keys(mod).sort()));"
    ].join("\n");
    const out = execFileSync(process.execPath, ["--input-type=module", "-e", lister], {
      env: { ...process.env, GDOCS_BUILD_MJS_URL: MJS_URL },
      encoding: "utf8",
      windowsHide: true
    });
    const keys = JSON.parse(out) as string[];
    for (const expected of [
      "buildTo",
      "emitGlobalShims",
      "assertNoModuleTokens",
      "extractCallMapNames",
      "findCallMapExport",
      "assertEntriesPresentInBundle",
      "makeBanner"
    ]) {
      expect(keys).toContain(expected);
    }
  });

  it("importing the module never triggers a build as a side effect", () => {
    // The CLI block is guarded by an invoked-directly check; if that guard
    // broke, the import above would have written gdocs/dist as a side effect
    // of every test run. Assert the import alone created nothing new in tmp.
    callBuild("makeBanner", "1.2.3");
    expect(fs.existsSync(path.join(tmpRoot, "Code.gs"))).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Pure helper: makeBanner.
// ---------------------------------------------------------------------------

describe("makeBanner", () => {
  it("renders the install banner with the version", () => {
    const banner = callBuild("makeBanner", "1.2.3") as string;
    expect(banner).toContain("Rostrum for Google Docs v1.2.3");
    expect(banner).toContain("generated file, paste into Extensions > Apps Script");
    expect(banner).toContain("You don't need to read this");
    // Every banner line must be a comment — anything else would execute.
    for (const line of banner.trimEnd().split("\n")) {
      expect(line.startsWith("//")).toBe(true);
    }
  });

  it("carries no tool attribution", () => {
    const banner = callBuild("makeBanner", "1.2.3") as string;
    expect(banner).not.toMatch(/claude/i);
    expect(banner).not.toMatch(/anthropic/i);
    expect(banner).not.toMatch(/ai-generated/i);
    expect(banner).not.toMatch(/\bAI\b/);
  });

  it("rejects a non-semver version (guards against vundefined banners)", () => {
    expect(() => callBuild("makeBanner", "")).toThrow(/semver/);
    expect(() => callBuild("makeBanner", "junk")).toThrow(/semver/);
    expect(() => callBuild("makeBanner", null)).toThrow(/semver/);
  });
});

// ---------------------------------------------------------------------------
// Pure helpers: call-map extraction + shim emission.
// ---------------------------------------------------------------------------

describe("extractCallMapNames / emitGlobalShims", () => {
  it("accepts a plain array of function names", () => {
    expect(callBuild("extractCallMapNames", ["onOpen", "rstmHide"])).toEqual(["onOpen", "rstmHide"]);
  });

  it("accepts an array of entry objects carrying a function-name field", () => {
    const map = [{ fn: "onOpen" }, { functionName: "rstmHide" }, { name: "rstmShowAll" }];
    expect(callBuild("extractCallMapNames", map)).toEqual(["onOpen", "rstmHide", "rstmShowAll"]);
  });

  it("accepts a {key: fnName} record, using the values", () => {
    const names = callBuild("extractCallMapNames", { open: "onOpen", hide: "rstmHide" }) as string[];
    expect(names.sort()).toEqual(["onOpen", "rstmHide"]);
  });

  it("falls back to keys for an {fnName: label} record (labels are not identifiers)", () => {
    const names = callBuild("extractCallMapNames", {
      rstmHide: "Hide",
      rstmApplyStyles: "Apply debate styles"
    }) as string[];
    expect(names.sort()).toEqual(["rstmApplyStyles", "rstmHide"]);
  });

  it("accepts a record of entry objects, preferring the fn field over the key", () => {
    const names = callBuild("extractCallMapNames", {
      hide: { fn: "rstmHide", label: "Hide" },
      onOpen: { label: "ignored — key wins when no fn field" }
    }) as string[];
    expect(names.sort()).toEqual(["onOpen", "rstmHide"]);
  });

  it("rejects unusable call maps with diagnosable messages (failure paths)", () => {
    expect(() => callBuild("extractCallMapNames", null)).toThrow(/unusable call map/);
    expect(() => callBuild("extractCallMapNames", 42)).toThrow(/unusable call map/);
    expect(() => callBuild("extractCallMapNames", [])).toThrow(/onOpen/);
    expect(() => callBuild("extractCallMapNames", {})).toThrow(/empty/);
    expect(() => callBuild("extractCallMapNames", ["onOpen", "onOpen"])).toThrow(/duplicate/);
    expect(() => callBuild("extractCallMapNames", ["1bad"])).toThrow(/not a valid global function name/);
    expect(() => callBuild("extractCallMapNames", ["delete"])).toThrow(/not a valid global function name/);
    expect(() => callBuild("extractCallMapNames", ["Rostrum"])).toThrow(/shadow the bundle global/);
    // BOTH keys and values must fail the identifier check to reach the throw:
    // identifier keys with label-ish values take the documented fall-back-to-
    // KEYS path instead (the {fn: label} reading, asserted above).
    expect(() => callBuild("extractCallMapNames", { "x y": "a b", "if!": "2bad" })).toThrow(/identifiers/);
    expect(() => callBuild("extractCallMapNames", { a: "ok", b: {} })).toThrow(/mix of shapes/);
    expect(() => callBuild("extractCallMapNames", [7])).toThrow(/no function name/);
  });

  it("emits one top-level delegate per entry, plus an explanatory header", () => {
    const shims = callBuild("emitGlobalShims", ["onOpen", "rstmHide"]) as string;
    expect(shims).toContain("function onOpen(...args) { return Rostrum.onOpen(...args); }");
    expect(shims).toContain("function rstmHide(...args) { return Rostrum.rstmHide(...args); }");
    // The why-comment ships into Code.gs — the one place a curious user looks.
    expect(shims).toContain("TOP-LEVEL function declarations");
  });

  it("emitGlobalShims rejects a bad call map instead of emitting broken code", () => {
    expect(() => callBuild("emitGlobalShims", ["bad name"])).toThrow(/unusable call map/);
  });
});

// ---------------------------------------------------------------------------
// Pure helper: findCallMapExport (locating the constant on the adapterPure
// module namespace without hard-coding its export name).
// ---------------------------------------------------------------------------

describe("findCallMapExport", () => {
  it("finds a call-map-named export", () => {
    expect(callBuild("findCallMapExport", { CALL_MAP: ["onOpen"], other: 1 })).toEqual(["onOpen"]);
    expect(callBuild("findCallMapExport", { gdocsCallMap: { open: "onOpen" } })).toEqual({ open: "onOpen" });
  });

  it("falls back to a UNIQUE export that parses as a call map", () => {
    expect(callBuild("findCallMapExport", { entries: ["onOpen"], unrelated: 7 })).toEqual(["onOpen"]);
  });

  it("refuses to guess (failure paths)", () => {
    // Two call-map-named exports: ambiguous by name.
    expect(() => callBuild("findCallMapExport", { CALL_MAP: ["a"], callMap: ["b"] })).toThrow(/exactly one/);
    // Two anonymous parseable exports: ambiguous by shape.
    expect(() => callBuild("findCallMapExport", { a: ["onOpen"], b: ["alsoParses"] })).toThrow(
      /cannot locate the call-map export/
    );
    // Nothing usable at all — message lists what IS exported for diagnosis.
    expect(() => callBuild("findCallMapExport", { foo: 1 })).toThrow(/exports: foo/);
  });
});

// ---------------------------------------------------------------------------
// Pure helper: assertNoModuleTokens. The planted "export const x" case is the
// prompt-mandated canary; the rest cover every token class the helper guards.
// ---------------------------------------------------------------------------

describe("assertNoModuleTokens", () => {
  const BENIGN = [
    "// a comment that mentions exports and imports in prose",
    'var Rostrum = (() => { var x = 1; return { onOpen: () => x }; })();',
    "var exporter = 1;",
    "function reimport() { return exporter; }",
    "function onOpen(...args) { return Rostrum.onOpen(...args); }"
  ].join("\n");

  it("passes clean script text (no false positives on look-alike words)", () => {
    expect(() => callBuild("assertNoModuleTokens", BENIGN)).not.toThrow();
  });

  it("catches a planted export statement", () => {
    expect(() => callBuild("assertNoModuleTokens", BENIGN + "\nexport const x = 1;")).toThrow(
      /module token survived/
    );
  });

  it("catches a planted import statement", () => {
    expect(() => callBuild("assertNoModuleTokens", 'import { a } from "b";\n' + BENIGN)).toThrow(
      /import statement/
    );
  });

  it("catches dynamic import calls and import.meta", () => {
    expect(() => callBuild("assertNoModuleTokens", BENIGN + '\nlazy(() => import("./x"));')).toThrow(
      /dynamic import/
    );
    expect(() => callBuild("assertNoModuleTokens", BENIGN + "\nvar u = import.meta.url;")).toThrow(
      /import\.meta/
    );
  });

  it("catches CommonJS leftovers (require, module.exports, __require shim)", () => {
    expect(() => callBuild("assertNoModuleTokens", BENIGN + '\nconst f = require("x");')).toThrow(/require/);
    expect(() => callBuild("assertNoModuleTokens", BENIGN + "\nmodule.exports = {};")).toThrow(
      /module\.exports/
    );
    expect(() => callBuild("assertNoModuleTokens", BENIGN + '\n__require("fs");')).toThrow(/CJS dependency/);
  });

  it("reports the offending line number for diagnosis", () => {
    expect(() => callBuild("assertNoModuleTokens", BENIGN + "\nexport const x = 1;")).toThrow(/line 6/);
  });

  it("rejects non-string input (failure path)", () => {
    expect(() => callBuild("assertNoModuleTokens", null)).toThrow(/artifact text/);
  });
});

// ---------------------------------------------------------------------------
// Pure helper: assertEntriesPresentInBundle — the build-time guard that a
// call-map entry the adapter never exported fails the BUILD instead of
// becoming a shim that throws on the user's first menu click.
// ---------------------------------------------------------------------------

describe("assertEntriesPresentInBundle", () => {
  const BUNDLE = 'var Rostrum = (() => { function onOpen(e) {} return { onOpen: () => onOpen }; })();';

  it("passes when every entry appears in the bundle", () => {
    expect(() => callBuild("assertEntriesPresentInBundle", BUNDLE, ["onOpen"])).not.toThrow();
  });

  it("names every missing entry (failure path)", () => {
    expect(() => callBuild("assertEntriesPresentInBundle", BUNDLE, ["onOpen", "rstmHide", "rstmShowAll"])).toThrow(
      /rstmHide, rstmShowAll/
    );
  });
});

// ---------------------------------------------------------------------------
// The build itself — programmatic, into the suite temp dir (plan A11.vii).
// Requires the Wave D adapter module (gdocs/src/adapter/docsAdapter.ts +
// core/adapterPure.ts); until it lands these fail with the build's own
// actionable plan-S11 message, then go green untouched.
// ---------------------------------------------------------------------------

describe("buildTo(dir) artifact", () => {
  let result: GdocsBuildResult;
  let code = "";

  beforeAll(() => {
    // One real build shared by every assertion below — the artifact is
    // deterministic, and esbuild + child startup is the slow part.
    result = callBuild("buildTo", path.join(tmpRoot, "out")) as GdocsBuildResult;
    code = fs.readFileSync(result.codeGsPath, "utf8");
  }, 120000);

  it("writes a non-empty Code.gs exposing the Rostrum IIFE global", () => {
    expect(result.codeGsBytes).toBeGreaterThan(5000);
    expect(result.codeGsBytes).toBe(Buffer.byteLength(code, "utf8"));
    expect(code).toMatch(/var Rostrum\s*=/);
  });

  it("ships a byte-exact appsscript.json next to Code.gs (one-folder packet)", () => {
    const shipped = fs.readFileSync(result.appsscriptJsonPath);
    const source = fs.readFileSync(path.resolve(__dirname, "../gdocs/appsscript.json"));
    expect(shipped.equals(source)).toBe(true);
  });

  it("contains zero module tokens (Apps Script executes plain script)", () => {
    expect(() => callBuild("assertNoModuleTokens", code)).not.toThrow();
  });

  it("opens with the version banner and reports the same version", () => {
    const firstLine = code.split("\n")[0];
    expect(firstLine).toContain(`Rostrum for Google Docs v${GDOCS_VERSION}`);
    expect(firstLine).toContain("generated file, paste into Extensions > Apps Script");
    expect(result.version).toBe(GDOCS_VERSION);
  });

  it("carries no tool attribution anywhere in the artifact", () => {
    expect(code).not.toMatch(/claude/i);
    expect(code).not.toMatch(/anthropic/i);
    expect(code).not.toMatch(/ai-generated/i);
    expect(code).not.toMatch(/\bAI\b/);
  });

  it("emits exactly one top-level shim per adapterPure call-map entry", async () => {
    const ns = await loadAdapterPureNamespace();
    // JSON round-trip drops function-valued exports — harmless here, because
    // the call map is a data constant and that is all the finder needs.
    const callMap = callBuild("findCallMapExport", JSON.parse(JSON.stringify(ns)));
    const names = callBuild("extractCallMapNames", callMap) as string[];
    expect(names.length).toBeGreaterThan(0);
    // onOpen is load-bearing for the whole product: no onOpen, no menu.
    expect(names).toContain("onOpen");
    for (const n of names) {
      // Anchored at column 0: Apps Script only discovers TOP-LEVEL functions.
      expect(code).toMatch(new RegExp(`^function ${n}\\(`, "m"));
      expect(code).toContain(`function ${n}(...args) { return Rostrum.${n}(...args); }`);
    }
    // No stray or duplicated delegates beyond the call map.
    const delegateCount = (code.match(/^function [A-Za-z_$][A-Za-z0-9_$]*\(\.\.\.args\) \{ return Rostrum\./gm) ?? [])
      .length;
    expect(delegateCount).toBe(names.length);
    expect(result.shimNames.sort()).toEqual([...names].sort());
  });

  it("carries the STRINGS copy deck (menu-label spot check)", () => {
    // Two distinctive, ASCII-safe labels prove strings.ts made it into the
    // bundle (the deck is the adapter's only legal source of user copy).
    expect(code).toContain(STRINGS.menu.applyStyles);
    expect(code).toContain(STRINGS.menu.markCite);
  });
});

describe("buildTo(dir) failure paths", () => {
  it("rejects an empty target directory", () => {
    expect(() => callBuild("buildTo", "")).toThrow(/target directory/);
  });

  it("fails loudly when the target path is an existing file", () => {
    const blocked = path.join(tmpRoot, "blocked-out");
    fs.writeFileSync(blocked, "occupied", "utf8");
    expect(() => callBuild("buildTo", blocked)).toThrow();
  });
});
