# Rostrum download counter (Cloudflare Worker)

A tiny Worker that sits behind the site's two install deliverables: the **Word
`manifest.xml`** and the **Google Docs `Code.gs`**. It bumps one anonymous integer
per surface in KV, then serves the canonical file as a download. No request
metadata is ever read or stored. The entire datastore is two integers.

- `GET /manifest.xml` (or `/`) → `downloads++`, then serve the manifest with a
  `Content-Disposition: attachment` header. If the origin is unreachable it
  302-redirects to the canonical Pages copy, so an install is **never blocked**.
- `GET /code.gs` → `gdocs_downloads++`, then serve the Google Docs `Code.gs` the
  same way (attachment; 302-fallback to `CODE_GS_ORIGIN` if unreachable).
- `GET /count` → `{ "downloads": N, "gdocs": M, "total": N+M }` (CORS-open) for
  internal checks + the README badge. `downloads` is the Word tally (its key name
  is unchanged so the historical count is preserved); `gdocs` is the Google Docs
  tally; `total` is the unified cross-platform number the badge displays. Each is
  read independently and degrades to 0 on a KV hiccup, so `total` is never `NaN`.
- `HEAD` on either download path returns the attachment headers but is **not**
  counted and does not pull the upstream body (link prefetchers / health checks).

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
at `…/rostrum/gdocs/Code.gs`, which the deploy workflow publishes via its
`build:gdocs` step + the webpack copy of `gdocs/dist` → `dist/gdocs`.

### Deploy ordering (important)

The README badge reads `$.total`. A live Worker that predates this change returns
only `{downloads}`, so the badge would render "no data" until redeploy. **Deploy
this Worker first, confirm `/count` returns `total` (see Verify), and only then push
the site/README change** that points the badge at `$.total`. Because `downloads`
keeps its key, the existing Word tally carries over untouched.

The `/code.gs` route proxies `CODE_GS_ORIGIN` (the Pages copy at
`…/rostrum/gdocs/Code.gs`), which exists only after the master push that
publishes it. So the Google Docs download route is fully live only once that push
completes. If you deploy the Worker first, a `GET /code.gs` in the gap falls back
to a not-yet-published origin (a redirect to a 404) and records a few phantom
`gdocs` counts. This is harmless (the gdocs install page is not advertised until
that same push) and never affects the Word `/manifest.xml` route, whose origin is
already live. To avoid the gap, confirm `curl -sI …/code.gs` redirects to a live
200 before announcing the Google Docs surface.

Then wire the site to the Worker (only if your Worker host differs from the
`rostrum-downloads.rostrum.workers.dev` placeholder):

1. In `site/word.html`, replace the placeholder host on the `manifest.xml` links;
   in `site/google-docs.html`, replace it on the `/code.gs` link.
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
#   → {"downloads":0,"gdocs":0,"total":0}
curl -sI https://rostrum-downloads.rostrum.workers.dev/manifest.xml | grep -i content-disposition
#   → content-disposition: attachment; filename="manifest.xml"
curl -sI https://rostrum-downloads.rostrum.workers.dev/code.gs | grep -i content-disposition
#   → content-disposition: attachment; filename="Code.gs"
curl -s https://rostrum-downloads.rostrum.workers.dev/count
#   → {"downloads":1,"gdocs":1,"total":2}
```

## Tests

The handler logic is unit-tested by the add-in's Jest suite
(`__tests__/worker.test.ts`). Run `npm test` from the repo root. worker/ is outside
the coverage globs, so these tests are the SOLE correctness gate: every route and
failure path (including both downloads' 302 fallbacks, HEAD-uncounted, KV-fail-safe,
cross-counter isolation, and the per-key `/count` degrade) is enumerated there.
