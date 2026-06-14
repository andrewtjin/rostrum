// handler.js — pure request logic for the Rostrum download counter.
//
// CommonJS on purpose: the Jest suite (ts-jest, CommonJS) requires this file
// directly so the logic is unit-tested without any ESM/transform friction. The
// Cloudflare entry (index.js, an ES module) imports it and wraps `handleRequest`
// in the `export default { fetch }` shape the Workers runtime expects.
//
// One job: measure how many people install Rostrum, anonymously. Two surfaces —
// Word (serve manifest.xml; bump "downloads") and Google Docs, whose PRIMARY install
// is the "Make a copy" template (the /gdocs-copy counted redirect) with the legacy
// Advanced Code.gs download as a fallback. The whole datastore is three anonymous
// integers in KV. Privacy contract (must stay true to site/privacy.html): no request
// body is read, and no IP or header is stored or logged (observability off in wrangler.toml).

// Canonical install deliverables, used when the *_ORIGIN vars aren't set in the
// environment. Both live on the Pages site next to their install page.
const DEFAULT_MANIFEST_ORIGIN = "https://andrewtjin.github.io/rostrum/manifest.xml";
const DEFAULT_CODE_GS_ORIGIN = "https://andrewtjin.github.io/rostrum/google-docs/Code.gs";

// The Google Docs PRIMARY install target — the maintainer template's "Make a copy"
// dialog. /gdocs-copy counts the click then 302-redirects here, so the recommended
// install is measured exactly like the Word manifest download. Used when
// GDOCS_COPY_TARGET is unset (wrangler.toml sets it). This is the SINGLE on-Worker
// source of the template doc id — the install page links to /gdocs-copy, never to
// Google directly — so recreating the template is a one-line edit here + a redeploy.
// gdocsTemplate.test.ts asserts this is a real /copy URL (not a placeholder), which
// relocates the old dead-CTA guard to where the doc id now actually lives.
const DEFAULT_GDOCS_COPY_TARGET =
  "https://docs.google.com/document/d/1DCR2b0sETwjCa_8VOjuEJ7BRatj8PbyPadnyYkhryQw/copy";

// The KV keys — anonymous counters per install action. The Word key keeps its
// original KV name ("downloads") so the historical tally is preserved across renames;
// /count surfaces it as "word". The Google Docs surface has TWO actions: the primary
// template copy ("google_docs_copies") and the Advanced Code.gs download
// ("google_docs_downloads"); /count reports their SUM as "google_docs" — the
// apples-to-apples gdocs install number — plus each as a breakdown field. The whole
// datastore is these three integers.
const COUNTER_KEY = "downloads"; // Microsoft Word manifest.xml tally (KV key; /count field = "word")
const GOOGLE_DOCS_KEY = "google_docs_downloads"; // Google Docs — Advanced Code.gs downloads
const GOOGLE_DOCS_COPIES_KEY = "google_docs_copies"; // Google Docs — primary "Make a copy" template installs

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
 * Route a request. Only GET/HEAD are meaningful (downloads, a counted redirect, and a
 * counter read). manifest.xml (Word) and code.gs (Google Docs Advanced) are counted
 * downloads; /gdocs-copy is the counted redirect to the Google Docs template Copy
 * dialog; /count returns every tally plus the unified total.
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
    const word = await readCount(env, COUNTER_KEY);
    const googleDocsCopies = await readCount(env, GOOGLE_DOCS_COPIES_KEY);
    const googleDocsDownloads = await readCount(env, GOOGLE_DOCS_KEY);
    // `google_docs` is the apples-to-apples Google Docs INSTALL number: the primary
    // template copies plus the Advanced Code.gs downloads. `google_docs_copies` and
    // `google_docs_downloads` expose the split. `total` = Word + Google Docs (unchanged
    // contract — the README badge reads $.total). `word`'s KV key is still "downloads"
    // (COUNTER_KEY), so its historical count carried across the field rename. Each tally
    // is read independently and degrades to 0 on a KV hiccup, so `total` is never NaN.
    const googleDocs = googleDocsCopies + googleDocsDownloads;
    return new Response(
      JSON.stringify({
        word,
        google_docs: googleDocs,
        google_docs_copies: googleDocsCopies,
        google_docs_downloads: googleDocsDownloads,
        total: word + googleDocs
      }),
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

  // Google Docs PRIMARY install — "Make a copy" of the maintainer template. Unlike the
  // two file routes above, this 302-redirects to Google's own Copy dialog (the bound
  // script travels with the copy); we count the intent-to-copy first so the recommended
  // path is measured like the Word manifest download. HEAD (link prefetch / health
  // checks) redirects WITHOUT counting, mirroring the download routes' HEAD handling.
  if (url.pathname === "/gdocs-copy") {
    const target = env.GDOCS_COPY_TARGET || DEFAULT_GDOCS_COPY_TARGET;
    if (request.method !== "HEAD") {
      await bumpCount(env, GOOGLE_DOCS_COPIES_KEY);
    }
    return Response.redirect(target, 302);
  }

  return new Response("Not Found", { status: 404 });
}

module.exports = {
  handleRequest,
  bumpCount,
  readCount,
  COUNTER_KEY,
  GOOGLE_DOCS_KEY,
  GOOGLE_DOCS_COPIES_KEY,
  DEFAULT_MANIFEST_ORIGIN,
  DEFAULT_CODE_GS_ORIGIN,
  DEFAULT_GDOCS_COPY_TARGET
};
