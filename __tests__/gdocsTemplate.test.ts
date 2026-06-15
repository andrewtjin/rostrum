// gdocsTemplate.test.ts — the release-attestation + install-page guards for the
// gdocs "Make a copy" path (Loop 004, Stream 2).
//
// WHAT IS PROVEN HERE (and, just as importantly, what is NOT):
//   A. the committed google-docs/template.descriptor.json matches a fresh build
//      — so the attestation can never silently fall behind a Code.gs change;
//   B. the install page hero + README "Make a copy" links point at the Worker's
//      counted /gdocs-copy redirect (so the recommended install is measured, like
//      the Word manifest download — Loop 006);
//   C. the Worker's copy target (wrangler.toml + handler.js, which must agree) is a
//      REAL docs.google.com `.../copy` URL, not a release placeholder — the dead-CTA
//      guard, relocated to where the template doc id now lives;
//   D. the built Code.gs makes no network call — backing the "sends nothing
//      back" privacy promise the new install page leans on.
//
// HOW tools/build-gdocs.mjs IS LOADED (replicated from __tests__/gdocsBuild.test.ts,
// the shared origin of this harness — see that file's header for the full
// rationale): ts-jest compiles this file to CommonJS and downlevels `import()`
// to `require()`, which cannot execute a real ESM .mjs, and jest's sandbox has
// no --experimental-vm-modules. So the true `await import()` of the build module
// runs inside a one-shot Node child process (exactly how `npm run build:gdocs`
// runs it in production), driven over JSON in/out. A small replicated helper
// (not a shared module) is intentional: the two suites stay independently
// readable, and the harness is ~40 lines, not worth a cross-test abstraction.

import { execFileSync } from "child_process";
import * as crypto from "crypto";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { pathToFileURL } from "url";
import { GDOCS_VERSION } from "../google-docs/src/core/constants";

// The Worker's hard-coded fallback copy target — the redirect destination behind the
// hero CTA. require()d directly (handler.js is CommonJS, exactly how worker.test.ts
// loads it) so the relocated dead-CTA guard checks the SAME value the Worker ships.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { DEFAULT_GDOCS_COPY_TARGET } = require("../worker/src/handler.js") as {
  DEFAULT_GDOCS_COPY_TARGET: string;
};

// ---------------------------------------------------------------------------
// Child-process driver for tools/build-gdocs.mjs (origin: gdocsBuild.test.ts).
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
 * like a whole Code.gs text), and prints a JSON envelope. Failures are CAUGHT
 * and reported in-band so the parent can rethrow the original message verbatim.
 * (Replicated from gdocsBuild.test.ts's DRIVER_SRC.)
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
 * Synchronous on purpose: execFileSync keeps assertion style plain without
 * rejects plumbing everywhere. (Replicated from gdocsBuild.test.ts's callBuild.)
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
    // A whole Code.gs can round-trip through the envelope; default 1MB is too
    // close for comfort on a bundle that grows with the engine.
    maxBuffer: 32 * 1024 * 1024,
    windowsHide: true
  });
  const payload = JSON.parse(stdout) as { ok: boolean; result?: unknown; error?: string };
  if (!payload.ok) throw new Error(payload.error ?? "unknown build-module failure");
  return payload.result;
}

// ---------------------------------------------------------------------------
// Repo-relative paths to the committed files under test.
// ---------------------------------------------------------------------------

const DESCRIPTOR_PATH = path.resolve(__dirname, "../google-docs/template.descriptor.json");
const INSTALL_PAGE_PATH = path.resolve(__dirname, "../site/google-docs.html");
const README_PATH = path.resolve(__dirname, "../google-docs/README.md");
// The repo-root README is the SECOND hardcoded gdocs-version surface (the brand hub
// chooser row + the "Google Docs port" section). Unlike google-docs/README.md it is not
// processed by webpack's __GDOCS_VERSION__ substitution (it is shipped as raw markdown),
// so it can drift silently on a version bump — exactly what happened (0.2.1 left stale
// while every derived surface advanced to 0.2.2). Guarded below.
const MAIN_README_PATH = path.resolve(__dirname, "../README.md");
const WRANGLER_PATH = path.resolve(__dirname, "../worker/wrangler.toml");

/** The counted-redirect endpoint the install page + README point at for "Make a
 * copy": the Worker (worker/src/handler.js) counts the click, then 302-redirects to
 * the template's Copy dialog. Anchored once so the two surfaces can't drift. */
const WORKER_COPY_ROUTE = "https://rostrum-downloads.rostrum.workers.dev/gdocs-copy";

/** The canonical placeholder a Worker copy target would carry if a maintainer wired
 * the template before creating the Doc. Since Loop 006 the site's copy href is a fixed
 * /gdocs-copy route and the doc id lives in the Worker (wrangler.toml + handler.js), so
 * this guards THOSE values. Defined once so the dead-CTA guard is anchored to ONE
 * string instead of magic substrings scattered through assertions. */
const TEMPLATE_DOC_ID_SENTINEL = "REPLACE_WITH_TEMPLATE_DOC_ID";

/** Lowercase sha256 hex of a buffer/string — the descriptor's hash convention
 * (matches tools/gen-gdocs-descriptor.mjs, which hashes the bytes on disk). */
function sha256(data: Buffer | string): string {
  return crypto.createHash("sha256").update(data).digest("hex");
}

/**
 * Extract the hero "Make a copy" CTA href from the install page. The hero anchor
 * carries class="cta"; google-docs.html has exactly one (the Advanced download uses
 * .button-link, not .cta). We return that single href and let the caller assert it is
 * the counted /gdocs-copy route. Throws a diagnosable error if zero or more than one
 * cta anchor exists, so a missing/duplicated hero fails loudly here rather than
 * yielding a misleading "" that silently passes a weak assertion.
 */
function heroCtaHref(html: string): string {
  // Walk every <a …class="cta"…> anchor and read its href attribute. A tolerant
  // attribute scan (class and href in either order, single or double quotes)
  // keeps this robust to harmless markup reshuffles on the install page.
  const anchorRe = /<a\b[^>]*\bclass\s*=\s*["'][^"']*\bcta\b[^"']*["'][^>]*>/gi;
  const hrefs: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = anchorRe.exec(html)) !== null) {
    const hrefMatch = /\bhref\s*=\s*["']([^"']*)["']/i.exec(m[0]);
    if (hrefMatch) hrefs.push(hrefMatch[1]);
  }
  if (hrefs.length === 0) {
    throw new Error(`no hero CTA (<a class="cta">) found in site/google-docs.html`);
  }
  if (hrefs.length > 1) {
    throw new Error(
      `multiple cta anchors in site/google-docs.html — expected exactly one hero CTA: ${JSON.stringify(hrefs)}`
    );
  }
  return hrefs[0];
}

/**
 * Read the deployed copy target (GDOCS_COPY_TARGET) out of worker/wrangler.toml — the
 * value the live Worker actually 302-redirects to. Throws if absent so a missing var
 * fails loudly. (A line-anchored regex, not a TOML parser: one scalar string, no deps.)
 */
function wranglerCopyTarget(): string {
  const toml = fs.readFileSync(WRANGLER_PATH, "utf8");
  const m = /^\s*GDOCS_COPY_TARGET\s*=\s*"([^"]+)"/m.exec(toml);
  if (!m) throw new Error("GDOCS_COPY_TARGET not found in worker/wrangler.toml");
  return m[1];
}

// ---------------------------------------------------------------------------
// One real build shared by every assertion that needs the artifact — the build
// is deterministic, and esbuild + child startup is the slow part. (Until the
// Wave D adapter module lands, callBuild("buildTo", …) throws the build's own
// actionable plan-S11 message and this beforeAll surfaces it on every test.)
// ---------------------------------------------------------------------------

let buildResult: GdocsBuildResult;
let codeGs = "";

beforeAll(() => {
  tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), "gdocs-template-test-"));
  buildResult = callBuild("buildTo", path.join(tmpRoot, "out")) as GdocsBuildResult;
  codeGs = fs.readFileSync(buildResult.codeGsPath, "utf8");
}, 120000);

afterAll(() => {
  if (tmpRoot) fs.rmSync(tmpRoot, { recursive: true, force: true });
});

// ---------------------------------------------------------------------------
// Test A — the committed descriptor rides with the build.
// ---------------------------------------------------------------------------

describe("template.descriptor.json release attestation", () => {
  it("the committed descriptor matches a fresh build", () => {
    // WHY: the descriptor is the repo's record of exactly which Code.gs was
    // published to the live template. Asserting committed == fresh build forces
    // the descriptor to be regenerated (via tools/gen-gdocs-descriptor.mjs) with
    // ANY Code.gs change — a commit-hygiene tripwire that reds the suite the
    // moment the attestation drifts from the engine.
    //
    // HONEST SCOPE: this proves descriptor == build. It does NOT prove the LIVE
    // template Doc actually contains this build — nothing in CI can reach that
    // Doc (no Google credential; standing posture). That residual gap is closed
    // OPERATIONALLY by the maintainer smoke test (make a copy, open the panel,
    // confirm the footer version), documented in docs/install-gdocs-template.md.
    expect(fs.existsSync(DESCRIPTOR_PATH)).toBe(true);
    const descriptor = JSON.parse(fs.readFileSync(DESCRIPTOR_PATH, "utf8")) as {
      gdocsVersion: string;
      codeGsSha256: string;
    };
    // The version line independently guards the committed JSON's gdocsVersion
    // field against a hand-edit typo (a corruption the sha check cannot see);
    // the sha256 line below is the load-bearing build-drift tripwire.
    expect(descriptor.gdocsVersion).toBe(GDOCS_VERSION);
    expect(descriptor.codeGsSha256).toBe(sha256(fs.readFileSync(buildResult.codeGsPath)));
  });
});

// ---------------------------------------------------------------------------
// README version literal — the one hardcoded version surface with no other guard.
// ---------------------------------------------------------------------------

describe("README version literal", () => {
  it("tracks GDOCS_VERSION", () => {
    // WHY: footer / Diagnostics / site token / build banner / descriptor all
    // DERIVE from GDOCS_VERSION and are each test-guarded. google-docs/README.md:14
    // is a HARDCODED literal with no other guard — pin it so a future bump cannot
    // silently leave it stale while every derived surface advances.
    const readme = fs.readFileSync(README_PATH, "utf8");
    expect(readme).toContain(`v${GDOCS_VERSION}`);
  });

  it("the repo-root README's gdocs version surfaces track GDOCS_VERSION", () => {
    // WHY: the brand-hub README carries the gdocs version in TWO spots that are raw
    // markdown (no webpack token substitution), so they drift silently on a bump —
    // which is precisely how "early (v0.2.1)" survived the move to 0.2.2. We extract
    // whatever version each spot declares and assert it equals GDOCS_VERSION (DERIVED,
    // not a hardcoded expectation), so the test self-updates on a legitimate bump and
    // only reds when a surface is left stale.
    const readme = fs.readFileSync(MAIN_README_PATH, "utf8");

    // Spot 1 — the chooser/status row: "... | early (v0.2.2) |"
    const statusCell = /early \(v(\d+\.\d+\.\d+)\)/.exec(readme);
    expect(statusCell).not.toBeNull();
    expect(statusCell![1]).toBe(GDOCS_VERSION);

    // Spot 2 — the "Google Docs port" section: "An early **v0.2.2 MVP**: ..."
    const mvpLine = /\bv(\d+\.\d+\.\d+) MVP\b/.exec(readme);
    expect(mvpLine).not.toBeNull();
    expect(mvpLine![1]).toBe(GDOCS_VERSION);
  });
});

// ---------------------------------------------------------------------------
// Tests B + C — the install-page hero copy link.
// ---------------------------------------------------------------------------

describe("install-page + README copy link → the counted /gdocs-copy redirect", () => {
  it("the hero CTA points at the Worker's /gdocs-copy route (not Google directly)", () => {
    // WHY: the primary install is "click this, then make a copy". Routing the click
    // through the Worker's /gdocs-copy is what makes the recommended install COUNTED
    // (anonymously, like the Word manifest download — the whole point of Loop 006). If
    // the hero pointed straight at docs.google.com the copy would be invisible to the
    // counter. The Worker then 302-redirects to the real template Copy dialog.
    const html = fs.readFileSync(INSTALL_PAGE_PATH, "utf8");
    expect(heroCtaHref(html)).toBe(WORKER_COPY_ROUTE);
  });

  it("the README copy link uses the same /gdocs-copy route", () => {
    // The user-facing README must route through the same counted endpoint, or installs
    // that start from GitHub go uncounted and the two surfaces drift apart.
    const readme = fs.readFileSync(README_PATH, "utf8");
    expect(readme).toContain(WORKER_COPY_ROUTE);
  });
});

describe("Worker copy target — the relocated dead-CTA guard", () => {
  // The doc id moved off the install page (which now links to /gdocs-copy) and onto the
  // WORKER, which holds the redirect destination in two places that must agree:
  // wrangler.toml's GDOCS_COPY_TARGET (the deployed value) and handler.js's
  // DEFAULT_GDOCS_COPY_TARGET (the in-code fallback). The dead-CTA risk relocated with
  // it — a placeholder in EITHER would 302 users to a broken page — so these guards
  // keep both a real /copy URL AND identical, so neither a committed placeholder nor a
  // silent drift between the two layers can ship.
  //
  // The id class requires >=25 chars: real Google doc ids are ~44 chars of
  // [A-Za-z0-9_-], so this floor rejects short fake placeholders (TODO, xxxxxxxx,
  // YOUR_DOC_ID_HERE) that a looser pattern would wave through.
  const COPY_URL_RE = /^https:\/\/docs\.google\.com\/document\/d\/[A-Za-z0-9_-]{25,}\/copy$/;

  it("wrangler.toml GDOCS_COPY_TARGET is a well-formed template /copy URL", () => {
    expect(wranglerCopyTarget()).toMatch(COPY_URL_RE);
  });

  it("handler.js DEFAULT_GDOCS_COPY_TARGET is a well-formed template /copy URL", () => {
    expect(DEFAULT_GDOCS_COPY_TARGET).toMatch(COPY_URL_RE);
  });

  it("the deployed var and the in-code default agree (no silent drift)", () => {
    // If they diverge, prod uses the var while the fallback rots — assert equality so a
    // template-id change must update both layers in the same commit.
    expect(wranglerCopyTarget()).toBe(DEFAULT_GDOCS_COPY_TARGET);
  });

  it("the copy target is a real doc id, not a release placeholder", () => {
    // WHY: a placeholder doc id would 302 every visitor to a dead page. This is the
    // deliberate dead-CTA guard, relocated from the install page to the Worker target.
    const target = wranglerCopyTarget();
    expect(target).not.toContain(TEMPLATE_DOC_ID_SENTINEL);
    // Belt-and-suspenders: reject any obviously-placeholder id wording too, so a future
    // reworded sentinel cannot slip a dead CTA past the line above.
    const idSeg = /\/document\/d\/([^/]+)\/copy/.exec(target)?.[1] ?? "";
    expect(idSeg).not.toMatch(/REPLACE|TEMPLATE|PENDING|TODO|YOUR|EXAMPLE|SAMPLE|PLACEHOLDER|XXX/i);
  });
});

// ---------------------------------------------------------------------------
// Test D — the built Code.gs phones nothing home.
// ---------------------------------------------------------------------------

describe("built Code.gs network posture", () => {
  it("makes no network call", () => {
    // WHY: UrlFetchApp is the ONLY Apps Script network primitive. Asserting it
    // never appears in the artifact backs the binding "sends nothing back"
    // privacy promise the new install page leans on. It also pre-empts the new
    // version-pointer string in Help ever quietly becoming a fetch-latest
    // phone-home — any such regression reds here, conspicuously, before ship.
    expect(codeGs).not.toMatch(/UrlFetchApp/);
  });
});
