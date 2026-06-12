# Rostrum download counter (Cloudflare Worker)

A tiny Worker that sits behind the site's **manifest.xml** download link. It bumps
one anonymous integer in KV, then serves the canonical manifest as a download. No
request metadata is ever read or stored — the entire datastore is a single number.

- `GET /manifest.xml` (or `/`) → `downloads++`, then serve the manifest with a
  `Content-Disposition: attachment` header. If the origin is unreachable it
  302-redirects to the canonical Pages copy, so an install is **never blocked**.
- `GET /count` → `{ "downloads": N }` (CORS-open) for internal checks + the README badge.

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

Then wire the site to the Worker:

1. In `site/index.html`, replace the placeholder host
   `https://rostrum-downloads.rostrum.workers.dev` with your real URL
   (3 download links — Windows step 1, Mac step 1, and the Updates re-download).
2. In `README.md`, the badge already points at `/count` via the same placeholder —
   update it too.
3. **Rate-limit cap — deferred (optional).** A `*.workers.dev` subdomain is not a
   Cloudflare *zone*, so the dashboard WAF "Rate limiting rules" UI does NOT apply
   here — there is nothing to configure in the dashboard. To add a cap later, use
   the native **Workers rate-limiting binding** (declare it in `wrangler.toml` and
   call `env.RATE_LIMITER.limit(...)` in `handler.js`, fail-open) — it keys on IP
   ephemerally at the edge and stores nothing, so the privacy promise holds.
   Alternatively, attach a custom domain to unlock the WAF rate-limiting UI.

## Verify

```sh
curl -s https://rostrum-downloads.rostrum.workers.dev/count
#   → {"downloads":0}
curl -sI https://rostrum-downloads.rostrum.workers.dev/manifest.xml | grep -i content-disposition
#   → content-disposition: attachment; filename="manifest.xml"
curl -s https://rostrum-downloads.rostrum.workers.dev/count
#   → {"downloads":1}
```

## Tests

The handler logic is unit-tested by the add-in's Jest suite
(`__tests__/worker.test.ts`) — run `npm test` from the repo root.
