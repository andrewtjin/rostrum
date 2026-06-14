// gdocsTemplate.test.ts — the release-attestation + install-page guards for the
// gdocs "Make a copy" path (Loop 004, Stream 2).
//
// WHAT IS PROVEN HERE (and, just as importantly, what is NOT):
//   A. the committed google-docs/template.descriptor.json matches a fresh build
//      — so the attestation can never silently fall behind a Code.gs change;
//   B. the install page's hero "Make a copy" link is a well-formed
//      docs.google.com `.../copy` URL;
//   C. that same link is a REAL doc id, not the release placeholder (this one is
//      EXPECTED TO FAIL until the real template doc id is wired — it is the
//      deliberate dead-CTA deploy guard, see its comment);
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

/** The canonical placeholder that site/google-docs.html and google-docs/README.md
 * both ship in the copy href until a human wires the real template doc id at
 * release. Defined once so the dead-CTA guard is anchored to ONE string instead
 * of magic substrings scattered through assertions. */
const TEMPLATE_DOC_ID_SENTINEL = "REPLACE_WITH_TEMPLATE_DOC_ID";

/** Lowercase sha256 hex of a buffer/string — the descriptor's hash convention
 * (matches tools/gen-gdocs-descriptor.mjs, which hashes the bytes on disk). */
function sha256(data: Buffer | string): string {
  return crypto.createHash("sha256").update(data).digest("hex");
}

/**
 * Extract the hero "Make a copy" CTA href from the install page. The hero anchor
 * carries class="cta"; we select the one whose href is a docs.google.com link
 * (the page may grow secondary CTAs — word.html already has two — so "the cta
 * anchor" is matched by ITS destination, not by being the only one). Throws a
 * diagnosable error if no such anchor exists, so a missing hero fails loudly
 * here rather than yielding a misleading "" that silently passes a weak regex.
 */
function heroCopyHref(html: string): string {
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
  const docsHrefs = hrefs.filter((h) => /docs\.google\.com/i.test(h));
  if (docsHrefs.length === 0) {
    throw new Error(
      `no hero copy CTA found in site/google-docs.html — expected an <a class="cta"> whose ` +
        `href points at docs.google.com (found ${hrefs.length} cta anchor(s): ${JSON.stringify(hrefs)})`
    );
  }
  if (docsHrefs.length > 1) {
    throw new Error(
      `multiple docs.google.com cta anchors in site/google-docs.html — expected exactly one hero ` +
        `copy link: ${JSON.stringify(docsHrefs)}`
    );
  }
  return docsHrefs[0];
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
});

// ---------------------------------------------------------------------------
// Tests B + C — the install-page hero copy link.
// ---------------------------------------------------------------------------

describe("install-page hero copy link", () => {
  it("is a well-formed template /copy URL", () => {
    // WHY: the entire install path is "click this link, then File > Make a copy".
    // If the hero href is not a docs.google.com `.../copy` URL, the primary CTA
    // does not open the Copy dialog and the page is broken for every visitor.
    const html = fs.readFileSync(INSTALL_PAGE_PATH, "utf8");
    const href = heroCopyHref(html);
    // The id class requires >=25 chars: real Google doc ids are ~44 chars of
    // [A-Za-z0-9_-], so this floor rejects short fake placeholders (TODO,
    // xxxxxxxx, YOUR_DOC_ID_HERE) that a bare "+" would wave through.
    expect(href).toMatch(/^https:\/\/docs\.google\.com\/document\/d\/[A-Za-z0-9_-]{25,}\/copy$/);
  });

  it("is a real doc id, not a release placeholder", () => {
    // WHY: push-to-master DEPLOYS the install page. A placeholder doc id would
    // ship a DEAD primary CTA to production. This is the deliberate dead-CTA
    // deploy guard.
    //
    // EXPECTED TO FAIL until the real template doc id is wired at release: the
    // hero href ships a REPLACE_WITH_TEMPLATE_DOC_ID sentinel on purpose, and
    // this red is the signal that the id is still un-wired. It is written as a
    // normal assertion (NOT .skip/.todo) precisely so it stays loud — the suite
    // must not go green until a human swaps in the live doc id.
    const html = fs.readFileSync(INSTALL_PAGE_PATH, "utf8");
    const href = heroCopyHref(html);
    // Primary: the exact canonical sentinel the page ships until release.
    expect(href).not.toContain(TEMPLATE_DOC_ID_SENTINEL);
    // Belt-and-suspenders: reject any obviously-placeholder id wording too, so a
    // future reworded sentinel cannot slip a dead CTA past the line above.
    const idSeg = /\/document\/d\/([^/]+)\/copy/.exec(href)?.[1] ?? "";
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
