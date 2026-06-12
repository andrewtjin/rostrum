// Cloudflare Worker entry — ES module format (the runtime's `export default`
// shape). All logic lives in handler.js (CommonJS) so the Jest suite can require
// it without ESM/transform friction; wrangler/esbuild bundles this import fine.
import handler from "./handler.js";

// Defensive: esbuild's CJS-default interop gives us the exports object directly,
// but if a future bundler change wrapped it as { default: … } instead, the call
// below would silently 500 with observability off. Resolve both shapes and fail
// LOUDLY at startup (visible in staging) rather than per-request in production.
const handleRequest =
  (handler && handler.handleRequest) ||
  (handler && handler.default && handler.default.handleRequest);

if (typeof handleRequest !== "function") {
  throw new Error("download-counter: handler.handleRequest missing after bundle");
}

export default {
  async fetch(request, env, _ctx) {
    return handleRequest(request, env);
  }
};
