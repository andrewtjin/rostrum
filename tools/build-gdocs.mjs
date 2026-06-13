#!/usr/bin/env node
// tools/build-gdocs.mjs — the single-file Code.gs build for the gdocs surface
// (plan D9/S12 + A11.vii).
//
// WHY THIS EXISTS: the gdocs distribution path is "paste one file into
// Extensions > Apps Script" (plan D10) — no clasp, no Marketplace, no dev-side
// Google account. That demands ONE artifact with zero module syntax (Apps
// Script's editor executes plain top-level script, not ES modules), so this
// script bundles the adapter entry (which transitively pulls the whole pure
// core + the inlined sidebar HTML) into a single IIFE under the global name
// "Rostrum", then post-processes it:
//
//   1. prepend a human banner carrying GDOCS_VERSION (constants.ts is the one
//      source of that number — plan D14);
//   2. append top-level function-declaration shims for every public entry in
//      core/adapterPure.ts's call map — Apps Script only discovers TOP-LEVEL
//      function declarations (menu callbacks, simple triggers like onOpen,
//      google.script.run targets), and the bundle hides everything inside the
//      IIFE. Because the shims and the menu are generated from the SAME
//      call-map constant, they can never drift apart (plan S12);
//   3. assert no module tokens survived (import/export/require) — a leaked
//      token would make the pasted file die at parse time on the user's
//      machine, the one place we cannot debug (fail at BUILD time instead);
//   4. write google-docs/dist/Code.gs plus a byte-exact copy of google-docs/appsscript.json
//      so the install packet is one folder.
//
// The post-processing steps are EXPORTED PURE HELPERS so __tests__/
// gdocsBuild.test.ts can unit-test them and can run the whole build
// programmatically into a temp dir (never asserting on the gitignored dist/ —
// plan A11.vii). `npm run build:gdocs` runs `typecheck:gdocs` first, so by the
// time esbuild (which does NO type checking) runs, the sources are known-sound.

import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

// ---------------------------------------------------------------------------
// Paths — resolved from this file's location so the build works from any cwd
// (the test invokes it programmatically; npm invokes it from the repo root).
// ---------------------------------------------------------------------------

const TOOLS_DIR = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(TOOLS_DIR, "..");
const ADAPTER_ENTRY = path.join(REPO_ROOT, "google-docs", "src", "adapter", "docsAdapter.ts");
const CONSTANTS_TS = path.join(REPO_ROOT, "google-docs", "src", "core", "constants.ts");
const ADAPTER_PURE_TS = path.join(REPO_ROOT, "google-docs", "src", "core", "adapterPure.ts");
const APPSSCRIPT_JSON = path.join(REPO_ROOT, "google-docs", "appsscript.json");
const DEFAULT_OUT_DIR = path.join(REPO_ROOT, "google-docs", "dist");

// ---------------------------------------------------------------------------
// Identifier validation — a bad call-map entry must fail the BUILD, never ship
// a Code.gs that dies at parse time in the user's Apps Script editor.
// ---------------------------------------------------------------------------

/**
 * Words that cannot be used as function-declaration names. Includes the
 * strict-mode-only set (arguments/eval/yield/let/static and the future
 * reserved words) because esbuild may emit a "use strict" directive and we
 * refuse to depend on the artifact's strictness either way.
 */
const RESERVED_WORDS = new Set([
  "break", "case", "catch", "class", "const", "continue", "debugger", "default",
  "delete", "do", "else", "enum", "export", "extends", "false", "finally",
  "for", "function", "if", "import", "in", "instanceof", "new", "null",
  "return", "super", "switch", "this", "throw", "true", "try", "typeof",
  "var", "void", "while", "with", "yield", "let", "static", "await",
  "implements", "interface", "package", "private", "protected", "public",
  "arguments", "eval"
]);

/** ASCII-identifier check — Apps Script entry points are ASCII by convention,
 * and a stricter charset here means the generated shims can never surprise. */
function isValidIdentifier(name) {
  return typeof name === "string"
    && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(name)
    && !RESERVED_WORDS.has(name);
}

// ---------------------------------------------------------------------------
// Call-map handling. core/adapterPure.ts owns the call-map constant (one
// source for both the onOpen menu and these shims — plan S12). The extractor
// tolerates the reasonable shapes a call map can take so the build and the
// adapter module stay decoupled at the SHAPE level while staying rigidly
// coupled at the CONTENT level (every entry must be a valid global fn name).
// ---------------------------------------------------------------------------

/** Pull a function name out of one object-shaped call-map entry. Checked in
 * priority order; `name`/`label`-style fields come last because they may hold
 * menu labels rather than identifiers. */
function nameFromEntryObject(entry) {
  for (const field of ["fn", "fnName", "functionName", "global", "name"]) {
    if (typeof entry[field] === "string") return entry[field];
  }
  return null;
}

/**
 * Normalize ANY accepted call-map shape into a validated list of global
 * function names. Accepted shapes:
 *   * string[]                          — names directly;
 *   * Array<{fn|fnName|functionName|global|name}> — names from the field;
 *   * Record<string, string>            — VALUES preferred (the
 *     `{menuKey: "rstmHide"}` reading); falls back to KEYS when values are
 *     not identifiers (an fn→label map: labels like "Mark cite" never pass
 *     the identifier check, so the fallback is unambiguous in practice);
 *   * Record<string, object>            — field from the value object, else
 *     the key itself.
 * Throws (never returns garbage) on anything else: an unusable call map must
 * stop the build, because the shims ARE the product's entire entry surface.
 */
export function extractCallMapNames(callMap) {
  const fail = (why) => {
    throw new Error(`gdocs build: unusable call map from core/adapterPure.ts — ${why}`);
  };
  let names = null;
  if (Array.isArray(callMap)) {
    names = callMap.map((entry, i) => {
      if (typeof entry === "string") return entry;
      if (entry && typeof entry === "object") {
        const n = nameFromEntryObject(entry);
        if (n !== null) return n;
      }
      return fail(`array entry ${i} carries no function name (expected a string or {fn|fnName|functionName|global|name})`);
    });
  } else if (callMap && typeof callMap === "object") {
    const keys = Object.keys(callMap);
    if (keys.length === 0) fail("the object is empty");
    const values = keys.map((k) => callMap[k]);
    if (values.every((v) => typeof v === "string")) {
      if (values.every(isValidIdentifier)) names = values;
      else if (keys.every(isValidIdentifier)) names = keys;
      else fail("neither all keys nor all values are valid identifiers");
    } else if (values.every((v) => v && typeof v === "object" && !Array.isArray(v))) {
      names = keys.map((k) => nameFromEntryObject(callMap[k]) ?? k);
    } else {
      fail("object values are a mix of shapes (expected all strings or all objects)");
    }
  } else {
    fail(`expected an array or object, got ${callMap === null ? "null" : typeof callMap}`);
  }
  if (names.length === 0) fail("it lists no entries — even an empty menu needs onOpen");
  const seen = new Set();
  for (const n of names) {
    if (!isValidIdentifier(n)) fail(`"${n}" is not a valid global function name`);
    // A shim named after the bundle global would shadow it and break every
    // other shim — structurally impossible to want.
    if (n === "Rostrum") fail(`an entry is named "Rostrum", which would shadow the bundle global`);
    if (seen.has(n)) fail(`duplicate entry "${n}"`);
    seen.add(n);
  }
  return names;
}

/**
 * Locate the call-map export on the adapterPure module namespace. Prefers an
 * export whose NAME says call map (/call.?map/i); otherwise accepts a UNIQUE
 * export that parses as one (adapterPure's other exports are functions, which
 * never parse). Anything ambiguous throws with the export list, because
 * guessing here could silently shim the wrong surface.
 */
export function findCallMapExport(moduleNamespace) {
  if (!moduleNamespace || typeof moduleNamespace !== "object") {
    throw new Error("gdocs build: adapterPure module namespace is not an object");
  }
  const keys = Object.keys(moduleNamespace);
  const preferred = keys.filter((k) => /call.?map/i.test(k));
  if (preferred.length === 1) return moduleNamespace[preferred[0]];
  if (preferred.length > 1) {
    throw new Error(`gdocs build: multiple call-map-named exports in core/adapterPure.ts (${preferred.join(", ")}) — keep exactly one`);
  }
  const parseable = keys.filter((k) => {
    try {
      extractCallMapNames(moduleNamespace[k]);
      return true;
    } catch {
      return false;
    }
  });
  if (parseable.length === 1) return moduleNamespace[parseable[0]];
  throw new Error(
    `gdocs build: cannot locate the call-map export in core/adapterPure.ts ` +
    `(exports: ${keys.join(", ") || "none"}) — export one constant matching /call.?map/i (plan S12)`
  );
}

// ---------------------------------------------------------------------------
// Pure post-processing helpers (unit-tested in __tests__/gdocsBuild.test.ts).
// ---------------------------------------------------------------------------

/**
 * The banner the user actually sees at the top of the pasted file. Lives here
 * (not in core/strings.ts) deliberately: STRINGS is the IN-PRODUCT copy deck
 * audited by the lexicon test; this is artifact packaging, and the version
 * interpolation must happen at build time. No tool attribution of any kind —
 * the file presents as what it is, a generated build of the Rostrum repo.
 */
export function makeBanner(version) {
  // Strict semver guard: a wrong import or refactor upstream would otherwise
  // print "vundefined" into the one line every installer reads.
  if (typeof version !== "string" || !/^\d+\.\d+\.\d+(-[A-Za-z0-9.-]+)?$/.test(version)) {
    throw new Error(`gdocs build: GDOCS_VERSION is not a semver string (got ${JSON.stringify(version)})`);
  }
  return [
    `// Rostrum for Google Docs v${version} — generated file, paste into Extensions > Apps Script. You don't need to read this.`,
    `// Built from the rostrum-addin repo (npm run build:gdocs); hand edits are overwritten by the next build.`
  ].join("\n") + "\n";
}

/**
 * Render the top-level shim block for every public entry in the call map.
 * Apps Script discovers ONLY top-level function declarations; the bundle's
 * implementations live on the Rostrum IIFE global, so each shim is a one-line
 * delegate. Rest args (ES2019-safe on the V8 runtime) keep one shim shape
 * correct for zero-arg menu handlers, onOpen(e), and google.script.run calls
 * with any arity.
 */
export function emitGlobalShims(callMap) {
  const names = extractCallMapNames(callMap);
  const header = [
    "// ---- Apps Script entry points ----------------------------------------------",
    "// Apps Script only sees TOP-LEVEL function declarations (menu callbacks,",
    "// simple triggers, google.script.run targets). Implementations live on the",
    "// Rostrum global above; these delegates are generated from the SAME call-map",
    "// constant the menu is built from, so menu and entry points can never drift."
  ];
  const shims = names.map((n) => `function ${n}(...args) { return Rostrum.${n}(...args); }`);
  return header.concat(shims).join("\n") + "\n";
}

/**
 * Module-syntax tokens that must NOT survive bundling — Apps Script executes
 * plain scripts, so any of these makes the pasted file fail at parse time (or,
 * for require/__require, at first click) on the user's machine. Patterns are
 * line-anchored where the token is only dangerous as a statement, so prose
 * like "exporter" or "reimport(" can never false-positive. Known tradeoff: a
 * multi-line template literal whose LINE starts with "export"/"import" would
 * trip this — acceptable, because failing visibly at build beats shipping a
 * dead artifact, and our own copy keeps those words off line starts.
 */
const MODULE_TOKEN_PATTERNS = [
  { re: /^[ \t]*import[ \t("'`]/m, what: "an import statement" },
  { re: /^[ \t]*export[ \t{]/m, what: "an export statement" },
  { re: /\bimport[ \t]*\(/, what: "a dynamic import call" },
  { re: /\bimport\.meta\b/, what: "an import.meta reference" },
  { re: /\brequire[ \t]*\(/, what: "a CommonJS require call" },
  { re: /\bmodule\.exports\b/, what: "a CommonJS module.exports assignment" },
  // esbuild only emits its __require interop shim when a CommonJS dependency
  // got bundled — which would throw at runtime on Apps Script. Catch it here.
  { re: /\b__require\b/, what: "esbuild's CommonJS interop shim (a CJS dependency was bundled)" }
];

/**
 * Throw if any module token survived in the final artifact text. Reports the
 * token kind, line number, and the offending line so a failure is diagnosable
 * from the build log alone.
 */
export function assertNoModuleTokens(code) {
  if (typeof code !== "string") {
    throw new Error("gdocs build: assertNoModuleTokens expects the artifact text");
  }
  for (const { re, what } of MODULE_TOKEN_PATTERNS) {
    const m = re.exec(code);
    if (m) {
      const line = code.slice(0, m.index).split("\n").length;
      const lineText = code.split("\n")[line - 1].trim().slice(0, 120);
      throw new Error(
        `gdocs build: module token survived bundling — ${what} at line ${line}: "${lineText}". ` +
        `Code.gs must be plain script (Apps Script has no module loader).`
      );
    }
  }
}

/**
 * Guard that every call-map entry textually appears in the bundle. esbuild
 * preserves export property names in the IIFE's return object, so an entry
 * the adapter never defined/exported is detectable here — at build time —
 * instead of becoming a shim that throws on the user's first menu click.
 * (Textual, not semantic: the strong guarantee comes from typecheck:gdocs,
 * which runs before this script; this catches call-map/export drift that
 * types alone cannot, e.g. a renamed export still referenced by the map.)
 */
export function assertEntriesPresentInBundle(bundleText, names) {
  const missing = names.filter((n) => !new RegExp(`\\b${n}\\b`).test(bundleText));
  if (missing.length > 0) {
    throw new Error(
      `gdocs build: call-map entries not found anywhere in the bundle: ${missing.join(", ")} — ` +
      `does google-docs/src/adapter/docsAdapter.ts export them?`
    );
  }
}

// ---------------------------------------------------------------------------
// Build-input loading. GDOCS_VERSION and the call map live in TypeScript, and
// this script is plain Node — so a tiny esbuild meta-bundle converts the two
// constants to importable ESM in memory (data: URL import, no temp files).
// ---------------------------------------------------------------------------

/** Forward-slash a path for use inside generated import specifiers (esbuild
 * resolves these on Windows either way; forward slashes avoid escape noise). */
function toPosix(p) {
  return p.split(path.sep).join("/");
}

/** Bundle + import the two TS constants the post-processing needs. */
async function loadBuildInputs() {
  if (!fs.existsSync(ADAPTER_PURE_TS)) {
    throw new Error(
      `gdocs build: ${ADAPTER_PURE_TS} is missing — the adapter module (plan S11) ` +
      `supplies the call-map constant the shims are generated from. Land it, then rebuild.`
    );
  }
  const stdin = [
    `export { GDOCS_VERSION } from ${JSON.stringify(toPosix(CONSTANTS_TS))};`,
    `export * as adapterPure from ${JSON.stringify(toPosix(ADAPTER_PURE_TS))};`
  ].join("\n");
  const result = await esbuild.build({
    stdin: { contents: stdin, resolveDir: REPO_ROOT, sourcefile: "gdocs-build-meta.ts", loader: "ts" },
    bundle: true,
    write: false,
    format: "esm",
    platform: "neutral",
    target: "es2022",
    logLevel: "silent"
  });
  const dataUrl = "data:text/javascript;base64," + Buffer.from(result.outputFiles[0].text).toString("base64");
  const mod = await import(dataUrl);
  return { version: mod.GDOCS_VERSION, callMap: findCallMapExport(mod.adapterPure) };
}

// ---------------------------------------------------------------------------
// The build itself.
// ---------------------------------------------------------------------------

/**
 * Bundle, post-process, and write Code.gs + appsscript.json into `outDir`.
 * Exported so gdocsBuild.test.ts can build into a temp dir (plan A11.vii: the
 * committed-state tests never look at the gitignored dist/). Returns a summary
 * the CLI and the tests both report from, so byte counts can't be hand-rolled.
 */
export async function buildTo(outDir) {
  if (typeof outDir !== "string" || outDir.length === 0) {
    throw new Error("gdocs build: buildTo(outDir) needs a target directory path");
  }
  if (!fs.existsSync(ADAPTER_ENTRY)) {
    throw new Error(
      `gdocs build: ${ADAPTER_ENTRY} is missing — the adapter module (plan S11) is the ` +
      `bundle entry. Land it, then rebuild.`
    );
  }
  const { version, callMap } = await loadBuildInputs();
  const bundle = await esbuild.build({
    entryPoints: [ADAPTER_ENTRY],
    bundle: true,
    write: false,
    // IIFE + globalName: everything stays off the top level except the one
    // "Rostrum" var the shims delegate to (Apps Script files share one global
    // namespace — a flat bundle would collide with the user's other scripts).
    format: "iife",
    globalName: "Rostrum",
    // Apps Script's V8 runtime is es2019-safe; newer syntax is downleveled.
    target: "es2019",
    // Closest platform model: no Node builtins, no filesystem. (Apps Script
    // has no DOM either, but "browser" only steers module resolution here.)
    platform: "browser",
    // Default charset (ascii) escapes non-ASCII as \uXXXX so the artifact
    // survives any clipboard/transport encoding on its way to the paste.
    charset: "ascii",
    // No minification: the pasted file is what the user debugs against if an
    // Apps Script error names a line — keep identifiers and structure honest.
    minify: false,
    // Strip @license/@preserve passthrough comments — the artifact opens with
    // OUR banner and nothing else speaks for the file.
    legalComments: "none",
    logLevel: "silent"
  });
  const bundleText = bundle.outputFiles[0].text;
  const shimNames = extractCallMapNames(callMap);
  // Order matters: check entry presence against the BUNDLE only (the shims
  // appended below would trivially satisfy the regex and mask a miss).
  assertEntriesPresentInBundle(bundleText, shimNames);
  const code = makeBanner(version) + "\n" + bundleText + "\n" + emitGlobalShims(callMap);
  assertNoModuleTokens(code);
  fs.mkdirSync(outDir, { recursive: true });
  const codeGsPath = path.join(outDir, "Code.gs");
  fs.writeFileSync(codeGsPath, code, "utf8");
  // Byte-exact copy: appsscript.json is hand-maintained next to the sources;
  // shipping a copy keeps the install packet one folder (plan S12).
  const appsscriptJsonPath = path.join(outDir, "appsscript.json");
  fs.copyFileSync(APPSSCRIPT_JSON, appsscriptJsonPath);
  return {
    outDir,
    codeGsPath,
    appsscriptJsonPath,
    codeGsBytes: Buffer.byteLength(code, "utf8"),
    shimNames,
    version
  };
}

// ---------------------------------------------------------------------------
// CLI entry — `npm run build:gdocs` lands here. Guarded so importing this
// module (the test does) never triggers a build as a side effect.
// ---------------------------------------------------------------------------

const invokedDirectly =
  process.argv[1] !== undefined &&
  import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;

if (invokedDirectly) {
  // Optional --out <dir> so humans can build elsewhere; default is the
  // gitignored google-docs/dist/ packet folder.
  const outFlag = process.argv.indexOf("--out");
  const outDir = outFlag >= 0 && process.argv[outFlag + 1]
    ? path.resolve(process.argv[outFlag + 1])
    : DEFAULT_OUT_DIR;
  buildTo(outDir).then(
    (r) => {
      console.log(`gdocs build OK (v${r.version})`);
      console.log(`  ${r.codeGsPath} — ${r.codeGsBytes} bytes, ${r.shimNames.length} entry shims`);
      console.log(`  ${r.appsscriptJsonPath} — copied`);
    },
    (e) => {
      console.error(e instanceof Error ? e.message : String(e));
      // exitCode (not process.exit) lets stdio flush before the process ends.
      process.exitCode = 1;
    }
  );
}
