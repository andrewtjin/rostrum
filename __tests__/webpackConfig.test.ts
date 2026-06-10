// Guards the production minimizer wiring in webpack.config.js.
//
// Webpack's DEFAULT minimizer (Terser) covers JS only — extracted CSS ships
// byte-for-byte verbatim unless a CSS minimizer is explicitly listed. That
// shipped ~42% dead bytes (comment banners + inline dev rationale) in every
// content-hashed sheet until CssMinimizerPlugin was added. This suite pins the
// two load-bearing facts of that fix, because BOTH regress silently (the build
// stays green either way, only the shipped bytes change):
//
//   1. A CssMinimizerPlugin instance is present → extracted sheets are minified
//      and dev commentary stops shipping in public assets.
//   2. The "..." token is present → webpack KEEPS its default Terser for JS.
//      Declaring `minimizer` REPLACES the default set, so dropping "..." while
//      editing the array would ship ~350KB of unminified JS with no error.
//
// The factory runs in production mode only: dev mode fetches local HTTPS certs
// (office-addin-dev-certs), which must not be a requirement for the test gate.
const CssMinimizerPlugin = require("css-minimizer-webpack-plugin");
const configFactory = require("../webpack.config.js");

describe("webpack production minimizer", () => {
  // The config module exports an async factory (it awaits dev certs in dev
  // mode); resolve it once with production argv and share across assertions.
  let minimizer: unknown[];

  beforeAll(async () => {
    const config = await configFactory({}, { mode: "production" });
    minimizer = config.optimization?.minimizer ?? [];
  });

  it('keeps webpack\'s default JS minifier via the "..." token', () => {
    expect(minimizer).toContain("...");
  });

  it("minifies extracted CSS with CssMinimizerPlugin", () => {
    expect(minimizer.some((m) => m instanceof CssMinimizerPlugin)).toBe(true);
  });
});
