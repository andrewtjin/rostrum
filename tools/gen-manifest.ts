// CLI: regenerate manifest.xml from the feature registry. Run via `npm run gen:manifest` (tsx).
// The actual XML is built by the pure, unit-tested `buildManifestXml`; this wrapper only supplies
// the real contributions + config and writes the file. After changing any feature's ribbon
// descriptor, run this — the drift test (__tests__/ribbonManifest.test.ts) enforces that you did.
import { writeFileSync } from "fs";
import { resolve } from "path";
import { contributions } from "../src/features/contributions";
import { buildManifestXml, manifestConfig } from "../src/features/manifestGen";

const xml = buildManifestXml(contributions, manifestConfig);
const out = resolve(__dirname, "..", "manifest.xml");
writeFileSync(out, xml, "utf8");
// eslint-disable-next-line no-console
console.log(`manifest.xml written: ${contributions.length} feature group(s) from the registry.`);
