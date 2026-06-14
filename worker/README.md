# Rostrum download counter (Cloudflare Worker)

A tiny Worker that sits behind the site's two install deliverables: the **Word
`manifest.xml`** and the **Google Docs `Code.gs`**. It bumps one anonymous integer
per surface in KV, then serves the canonical file as a download. No request
metadata is ever read or stored. The entire datastore is two integers.

- `GET /manifest.xml` (or `/`) → `downloads++`, then serve the manifest with a
  `Content-Disposition: attachment` header. If the origin is unreachable it
  302-redirects to the canonical Pages copy, so an install is **never blocked**.
- `GET /gdocs-copy` → `google_docs_copies++`, then **302-redirect** to the Google Docs
  template's `.../copy` dialog (`GDOCS_COPY_TARGET`, defaulting to the live template).
  This is the PRIMARY Google Docs install — counting the click here is what makes "Make
  a copy" measured, exactly like the Word manifest download.
- `GET /code.gs` → `google_docs_downloads++`, then serve the Google Docs `Code.gs` the
  same way as the manifest (attachment; 302-fallback to `CODE_GS_ORIGIN` if unreachable).
  This is the Advanced / fallback install.
- `GET /count` → `{ "word": N, "google_docs": M, "google_docs_copies": C,
  "google_docs_downloads": D, "total": N+M }` (CORS-open) for internal checks + the
  README badge. `word` is the Word tally (KV key still `downloads`, so history is
  preserved); `google_docs` is the apples-to-apples Google Docs install number
  (`= C + D`: the template copies plus the Advanced downloads); the two `google_docs_*`
  fields expose the split; `total = word + google_docs` is the unified number the badge
  reads (`$.total`). Each tally is read independently and degrades to 0 on a KV hiccup,
  so `total` is never `NaN`.
- `HEAD` on a download path returns the attachment headers (uncounted, no upstream pull);
  `HEAD /gdocs-copy` returns its 302 without counting — both for link prefetchers /
  health checks.

## One-time deploy

Prereqs: a free [Cloudflare](https://dash.cloudflare.com/sign-up) account and
[`wrangler`](https://developers.cloudflare.com/workers/wrangler/) (`npm i -g wrangler`,
or use `npx wrangler`).

```sh
cd worker
wrangler login                         # interactive browser auth
wrangler kv namespace create COUNTER   # prints an id …
#   → paste that id into wrangler.toml's  id = "REPLACE_WITH_KV_NAMESPACE_ID"
wrangler deploy                        # prints the live URL, e.g.
#   https://rostrum-downloads.rostrum.workers.dev
```

The two proxied origins are set in `wrangler.toml` (`MANIFEST_ORIGIN`,
`CODE_GS_ORIGIN`); both default to the Pages site if unset. `CODE_GS_ORIGIN` points
at `…/rostrum/google-docs/Code.gs`, which the deploy workflow publishes via its
`build:gdocs` step + the webpack copy of `google-docs/dist` → `dist/google-docs`.

### Deploy ordering (important)

The README badge reads `$.total`, which every current Worker build returns, so the
`downloads` → `word` field rename does not affect it — there is no deploy-order
constraint for the badge. The Word tally itself carries over untouched because its
KV key is still `downloads` (only the `/count` output field was renamed).

The `/code.gs` route proxies `CODE_GS_ORIGIN` (the Pages copy at
`…/rostrum/google-docs/Code.gs`), which exists only after the master push that
publishes it. So the Google Docs download route is fully live only once that push
completes. If you deploy the Worker first, a `GET /code.gs` in the gap falls back
to a not-yet-published origin (a redirect to a 404) and records a few phantom
`google_docs` counts. This is harmless (the Google Docs install page is not advertised until
that same push) and never affects the Word `/manifest.xml` route, whose origin is
already live. To avoid the gap, confirm `curl -sI …/code.gs` redirects to a live
200 before announcing the Google Docs surface.

**`/gdocs-copy` must be live before the install page points at it.** The install page's
hero CTA links to `…/gdocs-copy`. The Pages deploy (push to master) and this Worker deploy
(`wrangler deploy`) are independent pipelines, so to avoid a dead hero, deploy the **Worker
first**, confirm `curl -sI …/gdocs-copy` 302s to the template `.../copy` URL, and only then
push the install-page change. (`GDOCS_COPY_TARGET` ships in `wrangler.toml`; if it is unset,
the route falls back to `handler.js`'s `DEFAULT_GDOCS_COPY_TARGET`, so a var-less deploy
still lands users on the Copy dialog.)

Then wire the site to the Worker (only if your Worker host differs from the
`rostrum-downloads.rostrum.workers.dev` placeholder):

1. In `site/word.html`, replace the placeholder host on the `manifest.xml` links;
   in `site/google-docs.html`, replace it on the `/gdocs-copy` hero CTA and the
   `/code.gs` Advanced link; in `google-docs/README.md`, on the `/gdocs-copy` copy link.
2. In `README.md`, the badge points at `/count` via the same placeholder, so update it
   too.
3. **Rate-limit cap, deferred (optional).** A `*.workers.dev` subdomain is not a
   Cloudflare *zone*, so the dashboard WAF "Rate limiting rules" UI does NOT apply
   here; there is nothing to configure in the dashboard. To add a cap later, use
   the native **Workers rate-limiting binding** (declare it in `wrangler.toml` and
   call `env.RATE_LIMITER.limit(...)` in `handler.js`, fail-open). It keys on IP
   ephemerally at the edge and stores nothing, so the privacy promise holds.
   Alternatively, attach a custom domain to unlock the WAF rate-limiting UI.

## Verify

```sh
curl -s https://rostrum-downloads.rostrum.workers.dev/count
#   → {"word":0,"google_docs":0,"google_docs_copies":0,"google_docs_downloads":0,"total":0}
# HEAD probes verify headers WITHOUT counting (link prefetch / health checks):
curl -sI https://rostrum-downloads.rostrum.workers.dev/manifest.xml | grep -i content-disposition
#   → content-disposition: attachment; filename="manifest.xml"
curl -sI https://rostrum-downloads.rostrum.workers.dev/gdocs-copy | grep -i location
#   → location: https://docs.google.com/document/d/<DOC_ID>/copy
curl -sI https://rostrum-downloads.rostrum.workers.dev/code.gs | grep -i content-disposition
#   → content-disposition: attachment; filename="Code.gs"
# Real GET installs (a browser following each link) bump the counters; after one of each:
curl -s https://rostrum-downloads.rostrum.workers.dev/count
#   → {"word":1,"google_docs":2,"google_docs_copies":1,"google_docs_downloads":1,"total":3}
```

## Tests

The handler logic is unit-tested by the add-in's Jest suite
(`__tests__/worker.test.ts`). Run `npm test` from the repo root. worker/ is outside
the coverage globs, so these tests are the SOLE correctness gate: every route and
failure path (including both downloads' 302 fallbacks, HEAD-uncounted, KV-fail-safe,
cross-counter isolation, and the per-key `/count` degrade) is enumerated there.
