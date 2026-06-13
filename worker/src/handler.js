// handler.js — pure request logic for the Rostrum download counter.
//
// CommonJS on purpose: the Jest suite (ts-jest, CommonJS) requires this file
// directly so the logic is unit-tested without any ESM/transform friction. The
// Cloudflare entry (index.js, an ES module) imports it and wraps `handleRequest`
// in the `export default { fetch }` shape the Workers runtime expects.
//
// One job: serve the two install deliverables (Word's manifest.xml and Google
// Docs' Code.gs) while bumping a SINGLE anonymous integer PER SURFACE in KV.
// Privacy contract (must stay true to site/privacy.html): the only things ever
// retained are those two counters — no request body is read, and no IP or header
// is stored or logged (observability is disabled in wrangler.toml).

// Canonical install deliverables, used when the *_ORIGIN vars aren't set in the
// environment. Both live on the Pages site next to their install page.
const DEFAULT_MANIFEST_ORIGIN = "https://andrewtjin.github.io/rostrum/manifest.xml";
const DEFAULT_CODE_GS_ORIGIN = "https://andrewtjin.github.io/rostrum/google-docs/Code.gs";

// The two KV keys — one anonymous counter per install surface. The Word key
// keeps its original name ("downloads") so the historical tally is preserved
// across this change; the Google Docs key ("google_docs_downloads") is new and
// starts at 0. The whole datastore is these two integers.
const COUNTER_KEY = "downloads"; // Microsoft Word — manifest.xml downloads
const GOOGLE_DOCS_KEY = "google_docs_downloads"; // Google Docs — Code.gs downloads

// CORS for the public /count read only: the README badge and any internal
// dashboard fetch it cross-origin. We expose the aggregate numbers and nothing
// per-request, so a wildcard origin is safe here.
const COUNT_CORS = { "Access-Control-Allow-Origin": "*" };

/**
 * Best-effort increment of one counter key. MUST NOT throw into the response
 * path: a KV hiccup must never break a user's download. The count is a
 * deliberately soft lower bound (non-atomic read-modify-write can drop a write
 * under concurrency, and that is acceptable for a rough "how many people
 * installed" signal). `key` defaults to the Word counter so pre-existing callers
 * keep their behavior; the Google Docs route passes GOOGLE_DOCS_KEY.
 */
async function bumpCount(env, key = COUNTER_KEY) {
  try {
    const current = parseInt((await env.COUNTER.get(key)) || "0", 10) || 0;
    await env.COUNTER.put(key, String(current + 1));
  } catch {
    // Swallow — serving the file matters more than a perfectly accurate tally.
  }
}

/** Read one counter; 0 when unset or garbled. Never throws. */
async function readCount(env, key = COUNTER_KEY) {
  try {
    return parseInt((await env.COUNTER.get(key)) || "0", 10) || 0;
  } catch {
    return 0;
  }
}

/**
 * Serve one counted install deliverable — the DRY core shared by the manifest
 * and Code.gs routes. Count the intent-to-download first (best-effort, never
 * blocks), then proxy the canonical origin with an attachment header so the
 * browser SAVES the file instead of rendering it (the HTML `download` attribute
 * is ignored cross-origin; this header is what actually forces the save). If the
 * origin is unreachable, 302-redirect to it so an install is NEVER blocked by
 * this Worker being unhappy.
 */
async function serveCountedDownload(env, { key, origin, filename, contentType }) {
  await bumpCount(env, key);
  try {
    // cacheTtl: 0 so we always serve the latest deployed file, not a stale edge copy.
    const upstream = await fetch(origin, { cf: { cacheTtl: 0 } });
    if (!upstream.ok) throw new Error("upstream " + upstream.status);
    const body = await upstream.text();
    return new Response(body, {
      headers: {
        "Content-Type": contentType,
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store"
      }
    });
  } catch {
    // Origin unreachable → don't block: redirect to the canonical Pages copy.
    // (The site's own download buttons also carry a direct fallback link, so a
    // transient Worker hiccup never strands a real installer.)
    return Response.redirect(origin, 302);
  }
}

/**
 * Headers-only response for HEAD probes (link prefetchers, badge/health checks):
 * the attachment headers a real download would carry, but NO counter bump and NO
 * upstream body fetch.
 */
function headOnly(contentType, filename) {
  return new Response(null, {
    headers: {
      "Content-Type": contentType,
      "Content-Disposition": `attachment; filename="${filename}"`,
      "Cache-Control": "no-store"
    }
  });
}

/**
 * Route a request. Only GET/HEAD are meaningful (a static download + a counter
 * read). Two download routes (manifest.xml for Word, code.gs for Google Docs)
 * each count their own surface; /count returns both plus the unified total.
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

  // Internal / badge read — the per-surface tallies plus the unified
  // cross-platform total (what the public README badge shows).
  if (url.pathname === "/count") {
    const downloads = await readCount(env, COUNTER_KEY);
    const googleDocs = await readCount(env, GOOGLE_DOCS_KEY);
    return new Response(
      // `downloads` retained for back-compat (the pre-bifurcation badge query);
      // `google_docs` is the new surface; `total` is the unified cross-platform number.
      JSON.stringify({ downloads, google_docs: googleDocs, total: downloads + googleDocs }),
      {
        headers: {
          "Content-Type": "application/json",
          "Cache-Control": "no-store",
          ...COUNT_CORS
        }
      }
    );
  }

  // Word install deliverable. Bare "/" stays a convenience alias for the manifest
  // so a pasted Worker URL still serves the original file.
  if (url.pathname === "/manifest.xml" || url.pathname === "/") {
    if (request.method === "HEAD") {
      return headOnly("application/xml; charset=utf-8", "manifest.xml");
    }
    return serveCountedDownload(env, {
      key: COUNTER_KEY,
      origin: env.MANIFEST_ORIGIN || DEFAULT_MANIFEST_ORIGIN,
      filename: "manifest.xml",
      contentType: "application/xml; charset=utf-8"
    });
  }

  // Google Docs install deliverable — the single-file Code.gs users paste into
  // Extensions ▸ Apps Script. Counts the Google Docs surface, mirrors the manifest path.
  if (url.pathname === "/code.gs") {
    if (request.method === "HEAD") {
      return headOnly("text/javascript; charset=utf-8", "Code.gs");
    }
    return serveCountedDownload(env, {
      key: GOOGLE_DOCS_KEY,
      origin: env.CODE_GS_ORIGIN || DEFAULT_CODE_GS_ORIGIN,
      filename: "Code.gs",
      contentType: "text/javascript; charset=utf-8"
    });
  }

  return new Response("Not Found", { status: 404 });
}

module.exports = {
  handleRequest,
  bumpCount,
  readCount,
  COUNTER_KEY,
  GOOGLE_DOCS_KEY,
  DEFAULT_MANIFEST_ORIGIN,
  DEFAULT_CODE_GS_ORIGIN
};
