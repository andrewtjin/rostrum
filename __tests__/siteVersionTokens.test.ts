// Guards the version-INDEPENDENCE invariant for the two-surface site.
//
// webpack's CopyWebpackPlugin transform stamps __ROSTRUM_VERSION__ → pkg.version
// (the Word add-in) and __GDOCS_VERSION__ → GDOCS_VERSION (the Google Docs port)
// across every .html copied from site/. Because the two surfaces ship on separate
// semver, a page must carry ONLY its own token: if the Google Docs page ever picked
// up __ROSTRUM_VERSION__ (or vice-versa) it would silently stamp the WRONG product's
// version — a dishonest, hard-to-spot regression. The transform can't prevent that;
// this test makes the "each page carries only its own token" rule load-bearing.

import * as fs from "fs";
import * as path from "path";

const SITE = path.join(__dirname, "..", "site");
const read = (f: string) => fs.readFileSync(path.join(SITE, f), "utf8");

describe("site version-token discipline (version independence)", () => {
  test("the Word install page carries only the Word version token", () => {
    const html = read("word.html");
    expect(html).toContain("__ROSTRUM_VERSION__");
    expect(html).not.toContain("__GDOCS_VERSION__");
  });

  test("the Google Docs install page carries only the gdocs version token", () => {
    const html = read("google-docs.html");
    expect(html).toContain("__GDOCS_VERSION__");
    expect(html).not.toContain("__ROSTRUM_VERSION__");
  });

  test("the brand landing carries BOTH version tokens (it lists both surfaces)", () => {
    // The platform chooser shows each surface's own version next to its name, so the
    // landing legitimately carries both tokens — webpack stamps each to its product's
    // version. This is the one page where both appearing is correct (the per-surface
    // pages above must still carry ONLY their own token).
    const html = read("index.html");
    expect(html).toContain("__ROSTRUM_VERSION__");
    expect(html).toContain("__GDOCS_VERSION__");
  });

  test("the comparison and privacy pages carry no version token", () => {
    // These pages make no per-surface version claim, so neither token should appear
    // (a stray token here would render literally on the deployed page).
    for (const f of ["comparison.html", "privacy.html"]) {
      const html = read(f);
      expect(html).not.toContain("__ROSTRUM_VERSION__");
      expect(html).not.toContain("__GDOCS_VERSION__");
    }
  });
});
