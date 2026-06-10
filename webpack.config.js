/* eslint-disable */
// Webpack config for the task pane + ribbon command bundles.
//
// Stage 1 ships the core engine + tests; the React entry points referenced here
// (taskpane, commands) are filled in during Stage 2. The config is included now
// so the scaffold is complete and `npm run build` has a real target. Office
// add-ins must be served over HTTPS, hence the dev-server https block.
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
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
    module: {
      rules: [
        // transpileOnly: the full semantic check already runs as its own gate (`npm run
        // typecheck` locally and as the CI step right before build), so re-checking inside
        // webpack only duplicates work (~28% slower prod build, worse watch-mode rebuilds).
        { test: /\.tsx?$/, use: { loader: "ts-loader", options: { transpileOnly: true } }, exclude: /node_modules/ }
      ]
    },
    plugins: [
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
          { from: "site", to: ".", noErrorOnMissing: true }
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
