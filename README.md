# Rostrum — Invisibility Mode

A desktop **Word add-in** that hides debate-card *body* text while keeping headings,
cites, analytics, and highlighted runs visible — and is **natively reversible** even
without the add-in. Built for collapsing long policy/LD/PF briefs into a scannable
"blocks-only" view in round, then expanding them back losslessly.

> **Desktop only.** Rostrum hides text with Word's *Hidden font* attribute
> (`<w:vanish/>`), which is `WordApiDesktop 1.2` — present on **Word for Windows and
> Mac**, absent on Word for the web and Office 2016–2021 perpetual. On an unsupported
> host the pane explains the native fallback instead of failing silently.

---

## What it does

| Action | Effect |
|--------|--------|
| **Hide** | Hides every non-keeper body run (and collapses fully-hidden paragraphs), then *arms* the document. Keeps: paragraphs at outline level 0–3 (Heading 1–4 + the navy Analytics style), any paragraph containing a cite-styled run, and runs highlighted in a keep-color (extended to the whole word). |
| **Re-hide** | Deterministically re-derives over the whole document — catches anything newly typed or pasted. |
| **Show All** | Reveals everything Rostrum hid and disarms. Convergent: safe to run from any partial state. |
| **Keep colors** | Choose which highlight colors count as "keep" (default: all). Stored per-document (travels with the file) and as a per-device default. |
| **Apply Rostrum styles** | *(gated)* Sets the template's heading/cite sizes and boxes the Pocket. Reflows the document — reversible only with Ctrl+Z. |
| **Live mode** | *(best-effort, desktop)* Keeps the paragraph you're typing in visible while invisibility is ON. Re-hide is the real guarantee. |

The ON-state + keep-colors live in a single document-level **custom XML part** (the
"manifest"), so *any* machine that opens the file re-derives the same view
deterministically — no per-span tracking, no bloat, survives format round-trips.

---

## Requirements & install

- **Word for Windows or Mac (desktop).** Manifest floor: `WordApiDesktop 1.2` +
  `WordApi 1.4` (the hide mechanism + the manifest store + the Track-Changes gate).
- Node ≥ 18 for development.

### Sideload (development)

```bash
npm install
npm start          # builds, starts the HTTPS dev server on :3000, and sideloads
```

`npm start` (office-addin-debugging) trusts the dev certificate, serves the bundles
over HTTPS (Office requires HTTPS), and sideloads `manifest.xml` into Word. Then in
Word: **Home ▸ Rostrum pane**, or use the **Hide / Re-hide / Show All** ribbon
buttons directly.

To stop: `npm stop`. To rebuild without sideloading: `npm run build` (production) or
`npm run build:dev`.

> Before distributing, regenerate `<Id>` in `manifest.xml` per environment and point
> `SourceLocation` / icon URLs at your host instead of `https://localhost:3000`.

---

## Reversibility (the core guarantee)

Rostrum hides text with the **Hidden font attribute**, the same one Word's own Font
dialog writes. So a recipient who *doesn't have the add-in* can still recover
everything:

1. Select the text (Ctrl+A for the whole document).
2. **Home ▸ Font dialog (Ctrl+D) ▸ clear the _Hidden_ checkbox**, or toggle
   **Home ▸ ¶ (Show/Hide)** to view hidden text.

Caveat — **"Show hidden text" display setting:** if a reader has *File ▸ Options ▸
Display ▸ Hidden text* enabled, hidden runs render (greyed/dotted-underlined) rather
than disappearing. That's a viewer preference, not a Rostrum bug; Show All still
removes the hiding.

---

## Debugging (this is a first-class feature)

You cannot attach a normal debugger inside a live Word host, so Rostrum ships its own
diagnostics. Open the **Diagnostics** section at the bottom of the pane:

- **Capability matrix** — exactly which Office.js requirement sets this host
  advertises (and therefore which features are on).
- **Manifest state** — armed yes/no, and the active keep-colors.
- **Log level** — `debug / info / warn / error`, persisted across reloads.
- **Live log** — every host round-trip is a **namespaced, timed, correlated** tracer
  event (e.g. `adapter hide#3 | ✔ writeManifest (+812ms)`). Operations share an
  operation id so you can follow one user action end-to-end.
- **Copy bug report** — bundles the whole recent timeline + host/UA + capability
  matrix to the clipboard. Paste it into an issue and the failure is fully reproduced
  in text.

Under the hood (`src/core/debug.ts`): a ring-buffered tracer with pluggable sinks,
payload clamping (a 200-page OOXML string can't OOM the buffer), and
`OfficeExtension.Error` expansion (`.code` / `.debugInfo.errorLocation` are captured,
not just the generic `.message`). Every adapter/styles/live-mode/ribbon failure flows
through it. In the browser console you'll see the same lines mirrored.

---

## Architecture

```
src/
├─ core/                      # the engine — PURE, Office.js-free, 100% unit-tested
│  ├─ types.ts                #   WordPort contract + domain model
│  ├─ ooxml.ts                #   per-paragraph <w:vanish/> transforms
│  ├─ keepers.ts invisibility.ts manifest.ts settings.ts styles.ts guards.ts
│  ├─ debug.ts                #   the tracer (host-free, tested)
│  ├─ ooxmlPackage.ts         #   PURE whole-body split/splice + outline normalization
│  ├─ officeWordPort.ts       #   the REAL WordPort over Word.run  ← the only deep Office.js
│  └─ officeStyles.ts         #   ensureRostrumStyles (host glue + pure plan)
├─ liveMode.ts                # best-effort selection-visible (Common API)
├─ taskpane/
│  ├─ controller.ts           #   RostrumController — all UI orchestration, TESTED
│  ├─ useRostrum.ts           #   thin React hook over the controller
│  ├─ App.tsx + components/   #   presentational
│  └─ index.tsx taskpane.html
└─ commands/                  # ribbon handlers (reuse the controller)
```

The design principle (from Stage 1) holds: **pure policy behind a narrow port.** The
engine reasons about paragraphs as `(outline level, OOXML string)` and never touches
Word; one adapter (`officeWordPort.ts`) turns that into `Word.run` calls. Result: the
*entire* engine + adapter sequencing + UI orchestration are tested in plain Node with
a fake host — no Office mock, no browser. `npm test` → **159 tests**.

### Commit strategy (and the Step-0 fidelity spike)

The adapter supports two write mechanisms:

- **`per-paragraph`** *(default)* — read and write each paragraph through its own
  range. Index alignment is exact; no whole-document re-serialization. **Safe.**
- **`whole-body`** *(opt-in)* — one `body.getOoxml()` → splice only `<w:vanish/>` →
  one `body.insertOoxml("Replace")`. Fewer host round-trips at 200–300 pages, **but**
  its correctness depends on (a) xmldom re-serializing the whole body acceptably to
  Word and (b) the `<w:p>` document order aligning with `body.paragraphs`. Both need a
  **live-host fidelity spike** that can't run headless, so it is opt-in and carries a
  **structural alignment guard** that auto-falls-back to per-paragraph on any count
  mismatch (logged loudly).

To run the spike: build a throwaway doc with a TOC, section breaks, headers/footers,
tables, numbered lists, and fields; flip one run's visibility via `whole-body`; verify
numbering/sections/headers/fields/bookmarks/TOC are intact. If clean, flip the default
(`createOfficeWordPort({ commitStrategy: "whole-body" })`). The in-process perf test
shows whole-body's xmldom cost is real (≈9× the per-paragraph JS time on 5,000
paragraphs), so per-paragraph is also the better default until the spike justifies the
switch.

---

## Caveats

- **Live mode** is desktop best-effort: there is no per-keystroke event (it hooks the
  Common-API *selection-changed*), so a fast typist may briefly see hidden text.
  **Re-hide always reconciles.** Unavailable on web/Mac-without-`font.hidden`.
- **Co-authoring:** there's no reliable Office.js signal for a live co-authoring
  session. Hiding/showing while others edit may merge unpredictably; Show All is
  convergent, but run invisibility when you're the sole editor. The pane warns when
  armed.
- **Track Changes:** Hide refuses to run while Track Changes is on (a partial Undo
  could otherwise strand the document). The pane offers to toggle it off for the
  operation and restore it after.
- **Apply Rostrum styles** reflows pagination and is **not** undone by Show All — use
  Ctrl+Z. The pocket box needs `Style.borders` (desktop); the per-paragraph `w:pBdr`
  fallback is documented but not auto-applied.

---

## Testing

```bash
npm test           # 159 unit + integration tests (Node, no host)
npm run typecheck  # tsc --noEmit
npm run build      # production webpack bundle
```

Adapter/integration tests run the real engine through a fake `Word.RequestContext`
(`__tests__/fakeWord.ts`) that models Office.js's queue-then-`sync()` semantics, so
they assert the hard invariants directly: single-sync atomic commit, manifest
set-vs-add, the Track-Changes restore error, the multi-`<w:p>` guard, outline
normalization, and the whole-body alignment fallback.
