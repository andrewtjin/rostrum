// handler.js — pure request logic for the Rostrum download counter.
//
// CommonJS on purpose: the Jest suite (ts-jest, CommonJS) requires this file
// directly so the logic is unit-tested without any ESM/transform friction. The
// Cloudflare entry (index.js, an ES module) imports it and wraps `handleRequest`
// in the `export default { fetch }` shape the Workers runtime expects.
//
// One job: serve the canonical manifest.xml while bumping a SINGLE anonymous
// integer in KV. Privacy contract (must stay true to site/privacy.html): the
// only thing ever retained is that counter — no request body is read, and no IP
// or header is stored or logged (observability is disabled in wrangler.toml).

// Canonical manifest, used when MANIFEST_ORIGIN isn't set in the environment.
const DEFAULT_MANIFEST_ORIGIN = "https://andrewtjin.github.io/rostrum/manifest.xml";

// The lone KV key. The whole datastore is this one counter.
const COUNTER_KEY = "downloads";

// CORS for the public /count read only: the README badge and any internal
// dashboard fetch it cross-origin. We expose the aggregate number and nothing
// per-request, so a wildcard origin is safe here.
const COUNT_CORS = { "Access-Control-Allow-Origin": "*" };

/**
 * Best-effort increment. MUST NOT throw into the response path: a KV hiccup must
 * never break a user's download. The count is a deliberately soft lower bound
 * (non-atomic read-modify-write can drop a write under concurrency, and that is
 * acceptable for a rough "how many people installed" signal).
 */
async function bumpCount(env) {
  try {
    const current = parseInt((await env.COUNTER.get(COUNTER_KEY)) || "0", 10) || 0;
    await env.COUNTER.put(COUNTER_KEY, String(current + 1));
  } catch {
    // Swallow — serving the file matters more than a perfectly accurate tally.
  }
}

/** Read the current tally; 0 when unset or garbled. Never throws. */
async function readCount(env) {
  try {
    return parseInt((await env.COUNTER.get(COUNTER_KEY)) || "0", 10) || 0;
  } catch {
    return 0;
  }
}

/**
 * Route a request. Only GET/HEAD are meaningful (a static download + a counter
 * read). The download path counts first, then proxies the canonical manifest
 * with an attachment header; if the origin is unreachable it 302-redirects so an
 * install is NEVER blocked by this Worker being unhappy.
 */
async function handleRequest(request, env) {
  const url = new URL(request.url);

  // CORS preflight: a browser fetching /count cross-origin will preflight with
  // OPTIONS. shields.io and curl use a plain GET and never hit this, but answer
  // it so a browser-side dashboard can read the count too.
  if (request.method === "OPTIONS") {
    return new Response(null, {
      status: 204,
      headers: { ...COUNT_CORS, "Access-Control-Allow-Methods": "GET, HEAD, OPTIONS" }
    });
  }

  if (request.method !== "GET" && request.method !== "HEAD") {
    return new Response("Method Not Allowed", { status: 405 });
  }

  // Internal / badge read — just the number.
  if (url.pathname === "/count") {
    const downloads = await readCount(env);
    return new Response(JSON.stringify({ downloads }), {
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store",
        ...COUNT_CORS
      }
    });
  }

  // The download itself. Bare "/" is a convenience alias for the manifest so a
  // pasted Worker URL still serves the file.
  if (url.pathname === "/manifest.xml" || url.pathname === "/") {
    // HEAD probes (link prefetchers, badge/health checks) must not inflate the
    // tally or pull the full manifest body — answer headers-only, uncounted.
    if (request.method === "HEAD") {
      return new Response(null, {
        headers: {
          "Content-Type": "application/xml; charset=utf-8",
          "Content-Disposition": 'attachment; filename="manifest.xml"',
          "Cache-Control": "no-store"
        }
      });
    }

    // Count the intent-to-download up front (best-effort, never blocks).
    await bumpCount(env);

    const origin = env.MANIFEST_ORIGIN || DEFAULT_MANIFEST_ORIGIN;
    try {
      // cacheTtl: 0 so we always serve the latest deployed manifest, not a stale edge copy.
      const upstream = await fetch(origin, { cf: { cacheTtl: 0 } });
      if (!upstream.ok) throw new Error("upstream " + upstream.status);
      const body = await upstream.text();
      return new Response(body, {
        headers: {
          "Content-Type": "application/xml; charset=utf-8",
          // Force a download with a stable filename even cross-origin. The HTML
          // `download` attribute is ignored across origins; this header is what
          // actually makes the browser save the file instead of rendering it.
          "Content-Disposition": 'attachment; filename="manifest.xml"',
          "Cache-Control": "no-store"
        }
      });
    } catch {
      // Origin unreachable → don't block: redirect to the canonical Pages copy.
      // (The site's own download buttons also carry a direct ./manifest.xml
      // fallback, so a transient Worker hiccup never strands a real installer.)
      return Response.redirect(origin, 302);
    }
  }

  return new Response("Not Found", { status: 404 });
}

module.exports = { handleRequest, bumpCount, readCount, COUNTER_KEY, DEFAULT_MANIFEST_ORIGIN };
