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
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts"
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"]
    },
    module: {
      rules: [
        { test: /\.tsx?$/, use: "ts-loader", exclude: /node_modules/ }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"]
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets", noErrorOnMissing: true }]
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
