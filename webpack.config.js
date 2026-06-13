/* eslint-disable */
// Webpack config for the task pane + ribbon command bundles.
//
// Stage 1 ships the core engine + tests; the React entry points referenced here
// (taskpane, commands) are filled in during Stage 2. The config is included now
// so the scaffold is complete and `npm run build` has a real target. Office
// add-ins must be served over HTTPS, hence the dev-server https block.
const path = require("path");
const fs = require("fs");
// package.json.version is pinned to PRODUCT_VERSION by __tests__/version.test.ts,
// so it's the single, drift-proof source for stamping the version into the static
// site pages (the download page shows it) at build time.
const pkg = require("./package.json");

// The Google Docs port ships on its OWN semver (GDOCS_VERSION in
// gdocs/src/core/constants.ts), independent of pkg.version — the Word add-in is
// 0.3.x while the Docs MVP is 0.1.x. We stamp it onto the gdocs install page via
// a SEPARATE __GDOCS_VERSION__ token. Extract the literal by REGEX over the file
// TEXT — never require()/import the .ts. CI runs Node 20, which cannot strip
// TypeScript types, so requiring a file containing `export const X: T = ...`
// throws "Missing initializer in const declaration" at parse time (it only
// "works" on a newer local Node with strip-mode — a works-here/breaks-in-CI
// trap). A regex is Node-version-agnostic; a missing match throws HERE so a
// rename fails the build loudly instead of stamping "undefined" onto the page.
const GDOCS_CONSTANTS = path.resolve(__dirname, "gdocs/src/core/constants.ts");
const gdocsVersionMatch = fs
  .readFileSync(GDOCS_CONSTANTS, "utf8")
  .match(/GDOCS_VERSION\s*=\s*["']([^"']+)["']/);
if (!gdocsVersionMatch) {
  throw new Error(`webpack: GDOCS_VERSION not found in ${GDOCS_CONSTANTS} (renamed or moved?)`);
}
const gdocsVersion = gdocsVersionMatch[1];

const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const CssMinimizerPlugin = require("css-minimizer-webpack-plugin");
// Serve the SAME localhost cert that office-addin-dev-certs installs + trusts in the
// OS root store. Without this, webpack-dev-server self-signs a throwaway cert that
// Word's WebView2 does NOT trust -> "add-in doesn't contain a valid security
// certificate". getHttpsServerOptions() reads ~/.office-addin-dev-certs/localhost.*.
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, argv) => {
  const dev = argv.mode === "development";
  // Only the dev server needs the trusted cert; a production `build` must not require
  // certs to be installed on the machine, so fetch them lazily and only in dev.
  const httpsOptions = dev ? await devCerts.getHttpsServerOptions() : null;
  return {
    devtool: dev ? "source-map" : false,
    entry: {
      // ONE page (taskpane.html) that hosts the React pane AND wires the ribbon command handlers
      // (it imports ../commands/commands). Since 0.3.0 there is no separate commands.js entry — the
      // manifest's <FunctionFile> points here (no shared runtime; the Always-On spike was retired).
      taskpane: "./src/taskpane/index.tsx",
      // The full-window workspace dialog (opt-in space-heavy surface) — its own bundle so
      // it loads independently of the pane.
      dialog: "./src/dialog/index.tsx",
      // The ribbon progress pop-out — a tiny dialog shown only while a ribbon op runs.
      progress: "./src/progress/index.tsx"
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      // Content-hash the JS so Office's WebView (which caches add-in assets very aggressively)
      // is FORCED to refetch when the bundle changes — without this, users get stuck on stale
      // code after an update. HTML filenames stay stable (taskpane/commands/dialog/progress.html)
      // so the manifest + appPageUrl keep pointing at fixed pages; HtmlWebpackPlugin rewrites the
      // hashed <script> tags inside them automatically.
      filename: "[name].[contenthash].js",
      clean: true
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"]
    },
    optimization: {
      // The three pages (pane, dialog, progress) each bundled their own copy of React
      // (x3) and the core engine + xmldom (x2). Splitting shared chunks dedupes them:
      // total JS shipped drops ~55%, and a chunk fetched for one page is already in the
      // WebView's cache when another page loads. HtmlWebpackPlugin injects each page's
      // chunk set into its HTML automatically; filenames stay content-hashed.
      splitChunks: { chunks: "all" },
      // Webpack's default minimizer (Terser) handles JS ONLY — without an explicit CSS
      // minimizer the extracted sheets ship byte-for-byte verbatim, including the dev
      // comment banners and inline rationale notes (~42% of shipped CSS bytes). The
      // "..." token KEEPS the default Terser for JS (omitting it would silently ship
      // unminified JS); CssMinimizerPlugin (cssnano, safe default preset) runs beside
      // it. Minimizers only run with `optimization.minimize` on, i.e. production mode —
      // dev builds/watch are untouched.
      minimizer: ["...", new CssMinimizerPlugin()]
    },
    module: {
      rules: [
        // transpileOnly: the full semantic check already runs as its own gate (`npm run
        // typecheck` locally and as the CI step right before build), so re-checking inside
        // webpack only duplicates work (~28% slower prod build, worse watch-mode rebuilds).
        { test: /\.tsx?$/, use: { loader: "ts-loader", options: { transpileOnly: true } }, exclude: /node_modules/ },
        // Per-surface stylesheets are imported by each entry and EXTRACTED to content-hashed
        // files (same staleness rationale as the hashed JS — the WebView cache must be forced
        // to refetch changed styles). HtmlWebpackPlugin injects each page's <link> itself.
        { test: /\.css$/, use: [MiniCssExtractPlugin.loader, "css-loader"] }
      ]
    },
    plugins: [
      new MiniCssExtractPlugin({ filename: "[name].[contenthash].css" }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/dialog/dialog.html",
        chunks: ["dialog"]
      }),
      new HtmlWebpackPlugin({
        filename: "progress.html",
        template: "./src/progress/progress.html",
        chunks: ["progress"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets", noErrorOnMissing: true },
          // The static landing/sideload page → dist root, so the Pages site serves it at
          // `…/rostrum/` (alongside the prod manifest written there post-build). It's plain
          // HTML with no bundle, so it isn't a webpack entry — just copied verbatim.
          {
            from: "site",
            to: ".",
            noErrorOnMissing: true,
            // Stamp BOTH product versions into the static pages at build time so each
            // install page shows its own version without a hand-maintained string (or
            // any JS): __ROSTRUM_VERSION__ → the Word add-in (pkg.version) on the Word
            // pages, __GDOCS_VERSION__ → the Google Docs port on google-docs.html. The
            // two are version-independent; this single transform handles both so a page
            // referencing the wrong token would just no-op, never cross-leak a value.
            // Only .html is templated; the tokens are harmless no-ops where absent.
            transform(content, absPath) {
              return absPath.endsWith(".html")
                ? content
                    .toString()
                    .replace(/__ROSTRUM_VERSION__/g, pkg.version)
                    .replace(/__GDOCS_VERSION__/g, gdocsVersion)
                : content;
            }
          },
          // The Google Docs single-file deliverable (Code.gs + appsscript.json), built by
          // `npm run build:gdocs` into gdocs/dist/ in a SEPARATE prior step (the deploy
          // workflow runs build:gdocs BEFORE build). webpack's `clean` wipes the OUTPUT
          // dir (dist/), NOT the copy SOURCE (gdocs/dist/), so this lands them at
          // dist/gdocs/Code.gs + dist/gdocs/appsscript.json — where the site's gdocs
          // install page links them and the download-counter Worker proxies Code.gs.
          // noErrorOnMissing keeps a plain local `npm run build` (no prior build:gdocs)
          // green. Deliberately NO transform: the .gs must stay BYTE-IDENTICAL to the
          // built artifact (it embeds its own GDOCS_VERSION; a token pass could corrupt
          // it and would break version-independence).
          { from: "gdocs/dist", to: "gdocs", noErrorOnMissing: true }
        ]
      })
    ],
    devServer: {
      static: { directory: path.join(__dirname, "dist") },
      // Present the trusted office-addin-dev-certs cert in dev. The bare "https"
      // fallback only applies to non-dev modes, where the dev server isn't used.
      server: httpsOptions
        ? { type: "https", options: { key: httpsOptions.key, cert: httpsOptions.cert, ca: httpsOptions.ca } }
        : "https",
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" }
    }
  };
};
