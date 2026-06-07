// CLI: regenerate the add-in manifest from the feature registry. The dev-vs-prod DECISION lives in
// the pure, unit-tested `resolveManifestPlan` (tools/manifestCli.ts); this file is only the I/O
// shell around it, so the dangerous branching is provable in tests rather than via the filesystem.
//
//   • DEV  (default — no prod signal):  writes the COMMITTED ../manifest.xml from the localhost
//          `manifestConfig`. Byte-for-byte what the drift test pins. `npm run gen:manifest`.
//   • PROD (--origin <https url>):       writes a BUILD ARTIFACT (default dist/manifest.xml) whose
//          URLs target a real host. Never overwrites the committed file, so the drift test is
//          unaffected. `npm run gen:manifest:prod -- --origin=https://andrewtjin.github.io/rostrum`.
//
// A prod-intent invocation with a missing/empty/non-https origin EXITS NON-ZERO with guidance,
// instead of silently writing a broken or dev manifest.
import { writeFileSync, mkdirSync } from "fs";
import { resolve, dirname } from "path";
import { contributions } from "../src/features/contributions";
import { buildManifestXml } from "../src/features/manifestGen";
import { parseFlags, resolveManifestPlan } from "./manifestCli";

function main(): void {
  const plan = resolveManifestPlan(parseFlags(process.argv.slice(2)), process.env);
  const out = resolve(__dirname, "..", plan.outRelative);
  mkdirSync(dirname(out), { recursive: true }); // dist/ may not exist yet on a clean checkout/CI
  writeFileSync(out, buildManifestXml(contributions, plan.config), "utf8");
  // eslint-disable-next-line no-console
  console.log(
    plan.mode === "prod"
      ? `PROD manifest.xml written: ${out} (origin=${plan.config.origin}, id=${plan.config.id})`
      : `manifest.xml written: ${out} — ${contributions.length} feature group(s) from the registry.`
  );
}

try {
  main();
} catch (e) {
  // eslint-disable-next-line no-console
  console.error(`ERROR: ${(e as Error).message}`);
  process.exit(1);
}
