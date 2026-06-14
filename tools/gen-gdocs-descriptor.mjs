#!/usr/bin/env node
// tools/gen-gdocs-descriptor.mjs — the per-release descriptor regenerator for
// the gdocs "Make a copy" install path (Loop 004, Stream 2).
//
// WHY THIS EXISTS: the gdocs distribution is a single Google Doc the maintainer
// publishes by hand — paste Code.gs into the template's Apps Script, then share
// a `.../copy` link end-users click. Nothing in that loop is reproducible by CI
// (the standing posture is no clasp, no CI Google credential, no dev-side
// Google account — see docs/install-gdocs-template.md). The one thing we CAN
// pin in the repo is an ATTESTATION of exactly which build was published:
// google-docs/template.descriptor.json records the gdocs version and the
// sha256 of the Code.gs bytes that went into the live template.
//
// That descriptor is the release artifact gdocsTemplate.test.ts asserts against
// (committed descriptor == fresh build). If a human hand-typed the hash it would
// be an entire error class waiting to happen — a fat-fingered nibble reds the
// suite for a reason that has nothing to do with the code. So the descriptor is
// NEVER hand-edited: it is GENERATED here, straight from the same byte-
// deterministic buildTo() the npm build and the tests use. Run this once per
// release, right after pasting the new Code.gs into the template (the ordering
// is documented in docs/install-gdocs-template.md § Per-release).
//
// The version is taken from buildTo()'s returned `.version` — the SINGLE source.
// We deliberately do NOT re-parse constants.ts here: the build already resolves
// GDOCS_VERSION (and guards it as semver), so re-reading the constant would just
// open a second, silently-divergeable path to the same number.

import { buildTo } from "./build-gdocs.mjs";
import * as crypto from "node:crypto";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

// ---------------------------------------------------------------------------
// Paths — resolved from THIS file's location (mirrors build-gdocs.mjs's
// TOOLS_DIR/REPO_ROOT computation) so the script works from any cwd: npm runs
// it from the repo root, a maintainer might run it from tools/.
// ---------------------------------------------------------------------------

const TOOLS_DIR = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(TOOLS_DIR, "..");
const DESCRIPTOR_PATH = path.join(REPO_ROOT, "google-docs", "template.descriptor.json");

/**
 * Build into a throwaway temp dir, hash the produced Code.gs, and write the
 * descriptor. A temp dir (never google-docs/dist/) keeps this decoupled from
 * whatever the maintainer's working tree holds and matches the test's own
 * "build into a fresh dir" discipline — the descriptor reflects a clean build,
 * not leftover bytes.
 */
async function main() {
  // mkdtempSync gives a unique dir under the OS temp root, so concurrent runs
  // (or a left-behind dist/) can never cross-contaminate the hash input.
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "gdocs-descriptor-"));
  try {
    const result = await buildTo(tmpDir);
    // Hash the bytes on disk, not the in-memory string: the descriptor attests
    // to the FILE the maintainer pastes, so we read back exactly what was
    // written (utf8, byte-deterministic per build-gdocs.mjs's contract).
    const codeGsBytes = fs.readFileSync(result.codeGsPath);
    const codeGsSha256 = crypto.createHash("sha256").update(codeGsBytes).digest("hex");
    // Version comes from the build's single source — see header WHY.
    const gdocsVersion = result.version;

    // 2-space indent + trailing newline: matches the repo's JSON convention and
    // keeps the committed file diff-clean (no spurious whitespace churn when a
    // later release regenerates it).
    const descriptor = { gdocsVersion, codeGsSha256 };
    fs.writeFileSync(DESCRIPTOR_PATH, JSON.stringify(descriptor, null, 2) + "\n", "utf8");

    // Log what landed so the maintainer can eyeball it against the release they
    // think they are cutting (a wrong version here is a loud, visible mistake).
    console.log(`wrote ${DESCRIPTOR_PATH}`);
    console.log(`  gdocsVersion: ${gdocsVersion}`);
    console.log(`  codeGsSha256: ${codeGsSha256}`);
  } finally {
    // Always clean up the temp build, even on failure — nothing downstream
    // reads it, and leaving it would litter the temp root across releases.
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

main().catch((e) => {
  console.error(e instanceof Error ? e.message : String(e));
  // exitCode (not process.exit) lets stdout/stderr flush before exit.
  process.exitCode = 1;
});
