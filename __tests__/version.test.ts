// Guards the version single-source (features/version.ts). Before it existed, package.json and
// the manifest's 4-part <Version> were manually synced with NO drift guard — this suite makes
// a mismatch a test failure instead of a silent ship.
import { MANIFEST_VERSION, PRODUCT_VERSION, RIBBON_REVISION } from "../src/features/version";
import { manifestConfig } from "../src/features/manifestGen";
import { SettingsPanel } from "../src/features/settings/panel";
import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";

describe("product version single-source", () => {
  it("matches package.json (npm and the add-in can't drift)", () => {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const pkg = require("../package.json") as { version: string };
    expect(pkg.version).toBe(PRODUCT_VERSION);
  });

  it("derives the 4-part Office manifest version", () => {
    expect(MANIFEST_VERSION).toBe(`${PRODUCT_VERSION}.${RIBBON_REVISION}`);
    expect(manifestConfig.version).toBe(MANIFEST_VERSION);
  });

  it("is visible to the user in the Settings pane", () => {
    const html = renderToStaticMarkup(React.createElement(SettingsPanel));
    expect(html).toContain(`Rostrum v${PRODUCT_VERSION}`);
  });
});
