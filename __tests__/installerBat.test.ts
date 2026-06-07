// Drift guard for the Stage A.5 Windows installer scripts (site/install.bat + site/uninstall.bat).
//
// The installer hard-codes two things that MUST track the deployed production manifest: the add-in
// <Id> GUID (the registry value name Word loads the add-in by) and the manifest URL it fetches. If
// either drifts from prod, the installer silently registers a DEAD key — Word loads nothing, with no
// error and no clue why. These tests turn that silent failure into a red build; because the Pages
// deploy is gated on `npm test` (.github/workflows/deploy-pages.yml), a drifted installer can never
// reach a student. Style mirrors manifestCli.test.ts (pin the prod contract, fail loudly on drift).
import { readFileSync } from "fs";
import { join } from "path";
import { PROD_ID, prodConfig } from "../src/features/manifestGen";

// The canonical Pages origin the deploy workflow builds the prod manifest from
// (`gen:manifest:prod -- --origin=…`). The installer must fetch from this exact origin.
const PROD_ORIGIN = "https://andrewtjin.github.io/rostrum";

// The per-user developer-sideload key (Office 16.0 hive, HKCU → no admin). Both scripts target it.
const WEF_KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";

const SITE = join(__dirname, "..", "site");
const installBat = readFileSync(join(SITE, "install.bat"), "utf8");
const uninstallBat = readFileSync(join(SITE, "uninstall.bat"), "utf8");

/** Extract a `set "KEY=VALUE"` assignment from a .bat file (the value may contain backslashes). */
function batVar(source: string, key: string): string | undefined {
  const m = source.match(new RegExp(`set\\s+"${key}=([^"]*)"`, "i"));
  return m?.[1];
}

describe("install.bat stays in sync with the deployed prod manifest", () => {
  it("registers the prod add-in <Id> (PROD_ID), never the localhost-dev id", () => {
    expect(batVar(installBat, "ADDIN_ID")).toBe(PROD_ID);
  });

  it("fetches the manifest from the prod origin", () => {
    // prodConfig is the single source the deploy uses; assert our literal origin matches it, then
    // that the installer URL is that origin's manifest.xml.
    expect(prodConfig({ origin: PROD_ORIGIN }).origin).toBe(PROD_ORIGIN);
    expect(batVar(installBat, "MANIFEST_URL")).toBe(`${PROD_ORIGIN}/manifest.xml`);
  });

  it("writes the per-user WEF Developer key (correct Office hive, no admin)", () => {
    expect(batVar(installBat, "WEF_KEY")).toBe(WEF_KEY);
  });
});

describe("uninstall.bat removes exactly what install.bat created", () => {
  it("deletes the same add-in <Id>", () => {
    expect(batVar(uninstallBat, "ADDIN_ID")).toBe(PROD_ID);
  });

  it("targets the same WEF Developer key", () => {
    expect(batVar(uninstallBat, "WEF_KEY")).toBe(WEF_KEY);
  });

  it("removes the exact install folder install.bat created (no orphaned files)", () => {
    // install.bat derives the folder via MANIFEST_DIR; uninstall.bat hard-codes the literal
    // path in its rmdir. Pin both to the same location so a future relocation of one can't
    // silently leave the other deleting/keeping the wrong directory.
    expect(batVar(installBat, "MANIFEST_DIR")).toBe("%LOCALAPPDATA%\\Rostrum");
    expect(uninstallBat).toContain('rmdir /s /q "%LOCALAPPDATA%\\Rostrum"');
  });
});

describe("installer scripts are safe to double-click", () => {
  // A UTF-8 BOM or a smart quote at the top makes cmd.exe choke on the first line, so the scripts
  // must be pure 7-bit ASCII. (readFileSync('utf8') would surface a leading BOM as U+FEFF, which is
  // outside the ASCII range, so this assertion doubles as a BOM guard.)
  it("install.bat is pure ASCII", () => {
    expect(/^[\x00-\x7F]*$/.test(installBat)).toBe(true);
  });

  it("uninstall.bat is pure ASCII", () => {
    expect(/^[\x00-\x7F]*$/.test(uninstallBat)).toBe(true);
  });
});
