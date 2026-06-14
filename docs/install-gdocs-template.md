# Publish & refresh the Rostrum Google Docs template

This is the **maintainer-only** runbook for the Google Docs install path. End users
never see it: they click a **Make a copy** link, get their own copy of a template Doc that
already carries the Rostrum Apps Script, and use it. This page is how you *publish* that
template once and *refresh* it on every release. (Installing the Word add-in instead? See
[`install-word.md`](install-word.md). The user-facing usage guide is
[`google-docs/README.md`](../google-docs/README.md).)

> **Why a template Doc, not Marketplace or clasp.** The standing posture is no clasp, no CI
> Google credential, and no dev-side Google account (see *Why not clasp* below). So the
> distribution unit is a single Google Doc you own: its Apps Script holds the built
> `Code.gs`, and a `.../copy` share link hands every debater their own copy in one click.

---

## What this is

A copy of one Google Doc — the **template** — is what end users actually install. The
template owns:

- the built **`Code.gs`** (pasted into the Doc's bound Apps Script project), and
- the **`appsscript.json`** manifest alongside it.

When a user opens the `.../copy` link and clicks through, Google clones the Doc *and its
bound script* into their Drive. They reload, the **Rostrum** menu appears, and they are
running the exact build you published. Your job here is to keep that template's script
current and to keep the `.../copy` link live.

---

## One-time setup

Do this once, when the template Doc first goes live.

1. **Create a BLANK Google Doc.** Empty, neutral body — **no real prep, no PII, no card
   text.** A template seeded from a real debate doc would leak that prep (and any names in
   it) into *every* student copy forever. Start from a truly empty document.
2. **Open the script editor:** **Extensions ▸ Apps Script**.
3. **Paste the two built files.** Run `npm run build:gdocs`, then paste
   `google-docs/dist/Code.gs` into the script's `Code.gs`, and the contents of
   `google-docs/dist/appsscript.json` into the manifest (enable **Show "appsscript.json"
   manifest file in editor** under Project Settings if it is hidden). Save.
4. **Clear all history.** Remove every comment, suggestion, and revision-history trace —
   the simplest guarantee is that you created the Doc fresh in step 1, so there is none to
   begin with. (A copied template carries the original's comment threads with it.)
5. **Share Anyone with the link ▸ Viewer.** Viewer is enough: the `.../copy` flow only
   needs read access to clone. Do **not** grant Editor — that would let anyone edit the
   master.
6. **Record the doc id and the copy URL.** From the Doc URL
   `https://docs.google.com/document/d/<DOC_ID>/edit`, the copy target is
   `https://docs.google.com/document/d/<DOC_ID>/copy`. Save the `<DOC_ID>` — you will wire
   it into the **Worker** (see *TEMPLATE_DOC_ID* below), not the install page; the site
   links to the Worker's `/gdocs-copy` route, which counts the install then redirects here.

> **Ownership is public.** Whichever Google account OWNS the template, its **display name
> AND email** are visible to every viewer who opens the Doc. Choose the owning account
> deliberately (a project/role account, not a personal one you would rather not surface).

---

## Per-release

Every time the engine changes and you cut a new gdocs version:

1. **Build:** `npm run build:gdocs` → produces a fresh `google-docs/dist/Code.gs`.
2. **Paste** that new `Code.gs` into the template's Apps Script (**Extensions ▸ Apps
   Script**), replacing the old script body. Save. (Re-paste `appsscript.json` too if the
   manifest changed.)
3. **Regenerate the descriptor:** `node tools/gen-gdocs-descriptor.mjs`. This rebuilds into
   a temp dir, hashes the produced `Code.gs`, and rewrites
   `google-docs/template.descriptor.json` with the new version + sha256. **Never hand-edit
   that file** — the hash is generated precisely so no one transcribes it by hand. Commit
   the regenerated descriptor.
4. **Confirm the install link still opens the Copy dialog** — open
   `https://rostrum-downloads.rostrum.workers.dev/gdocs-copy` in a logged-out/incognito
   window and verify it redirects to Google's "make a copy" prompt. (The Worker only
   needs redeploying if the doc id changed — see *TEMPLATE_DOC_ID* — not every release.)
5. **Test:** `npm test`.

**What each guard actually catches — and what it does NOT:**

- The descriptor + `gdocsTemplate.test.ts` catch **descriptor/build DRIFT**: if you change
  `Code.gs` and forget to regenerate the descriptor (step 3), Test A reds. That is the only
  thing it can prove — that the committed attestation equals a fresh build.
- It does **NOT** prove the LIVE template Doc is fresh. Nothing in CI can reach that Doc (no
  Google credential — standing posture). The ONLY thing that proves the live Doc carries the
  new build is the **manual smoke test**: open the `.../copy` link, make a copy, open the
  Rostrum panel, and confirm the **footer version is the new one**. Do this every release —
  it is the irreducibly-human step the automation cannot cover.

---

## TEMPLATE_DOC_ID (single source — on the Worker)

The template doc id lives in the **download-counter Worker**, not the install page. The
site's **Make a copy** links (the hero CTA in `site/google-docs.html` and the link in
`google-docs/README.md`) all point at the Worker's fixed `/gdocs-copy` route, which counts
the install click then 302-redirects to the real `.../copy` URL. The doc id therefore
appears in exactly two Worker places that **must stay identical**:

- `worker/wrangler.toml` → `GDOCS_COPY_TARGET` (the deployed value), and
- `worker/src/handler.js` → `DEFAULT_GDOCS_COPY_TARGET` (the in-code fallback).

**Reuse the SAME doc id across releases.** The template Doc is long-lived; a new release
re-pastes its script, it does not mint a new Doc, and the site needs no change. If the id
genuinely must change (you had to recreate the template), edit **both Worker values above**
in one commit and **`wrangler deploy`** the Worker — no site/Pages change is involved.

A `REPLACE_WITH_TEMPLATE_DOC_ID` sentinel guards it: `gdocsTemplate.test.ts` asserts both
Worker values are a real `.../copy` URL, equal to each other, and free of any placeholder
wording. That red is the guard against shipping a **dead primary CTA** (the Worker would
otherwise 302 every visitor to a broken page) and against the two values silently drifting.

---

## Why not clasp / why the nudge is static

- **No clasp, no CI push.** There is no CI Google credential and no dev-side Google account
  (the project's standing posture). So publishing is a manual paste into the template's
  Apps Script, and freshness is proven by the manual smoke test above — not by an automated
  deploy.
- **The "newer releases" pointer in Help is deliberately STATIC.** It is a fixed string, not
  a fetch-latest-version check. A live version check would mean the Doc's script reaches out
  over the network (`UrlFetchApp`) — which would break the binding **"sends nothing back"**
  privacy promise the install page makes and force a privacy-page edit. Keeping it static
  means any future attempt to add a phone-home is conspicuous: it would have to introduce
  `UrlFetchApp`, which `gdocsTemplate.test.ts` ("makes no network call") reds on sight.
