# Rostrum — a Word add-in suite for debaters

A desktop **Word add-in suite** for collapsing, condensing, and formatting debate
briefs in round. Its flagship tool, **Invisibility Mode**, hides debate-card *body*
text while keeping headings, cites, analytics, and highlighted runs visible — and is
**natively reversible** even without the add-in. Built for turning long policy/LD/PF
briefs into a scannable "blocks-only" view in round, then expanding them back
losslessly.

> **Status — v0.3.0.1, live.** Two tools ship today (**Invisibility Mode** and
> **Condense & Shrink**) alongside a **Settings** pane and a ribbon-advertised roadmap
> (Format & Condense, Flow, Cite & Paste). The production build is hosted on GitHub
> Pages — install instructions below.

> **Desktop only.** Rostrum hides text with Word's *Hidden font* attribute
> (`<w:vanish/>`), which is `WordApiDesktop 1.2` — present on **Word for Windows and
> Mac**, absent on Word for the web and Office 2016–2021 perpetual. On an unsupported
> host the pane explains the native fallback instead of failing silently.

---

## The suite

Everything lives on a single **Rostrum** ribbon tab. Each tool is its own ribbon group
with a deep-linked pane; the suite is generated from one feature registry, so adding a
tool is "register it + regenerate the manifest."

| Group (left → right) | Status | What it is |
|---|---|---|
| **Settings** | live | App-wide settings — an informational pane covering how Rostrum loads on every document (the Trusted-Catalog install) and where to turn that off. |
| **Invisibility** | live | Hide card bodies to a headings/cites/highlights-only view; natively reversible. |
| **Condense** | live | Shrink card font size and condense paragraph spacing — losslessly reversible. |
| **Format & Condense** | planned | One-click styles cleanup, condensing, and legacy-file migration. |
| **Flow** | planned | Track arguments across a round in a speech-doc surface. |
| **Cite & Paste** | planned | Generate citations and paste cards with clean formatting. |

Planned tools already appear on the ribbon — each opens a "coming soon" surface — so the
extension path is proven end-to-end. When a tool is built, its status flips and its real
commands fill in; nothing else moves.

---

## Invisibility Mode

| Action | Effect |
|--------|--------|
| **Hide** | Hides every non-keeper body run (and collapses fully-hidden paragraphs), then *arms* the document. Keeps: paragraphs at outline level 0–3 (Heading 1–4 + the navy Analytics style), any paragraph containing a cite-styled run, and runs highlighted in a keep-color (extended to the whole word). **Idempotent + convergent — press it again after editing to re-derive over the whole document and catch newly typed or pasted text** (there is no separate "Re-hide" button). |
| **Show All** | Reveals everything Rostrum hid and disarms. Convergent: safe to run from any partial state. |
| **Apply Styles** | *(gated)* Sets the template's heading/cite sizes, boxes the Pocket, and repairs mis-styled cites. Reflows the document; needs desktop Word 1.5+ and is reversible only with Ctrl+Z. |
| **Options** | The deep-linked pane: keep-color settings, whole-body commit mode, and the Diagnostics console (below). |

The ON-state + keep-colors live in a single document-level **custom XML part** (the
"manifest"), so *any* machine that opens the file re-derives the same view
deterministically — no per-span tracking, no bloat, survives format round-trips.

---

## Condense & Shrink

Rostrum's lossless answer to Verbatim's Shrink + Condense. The ribbon surfaces four
direct verbs plus an Options pane (mode checkboxes, one-click mode buttons, a live
shrink-size readout, and the omission-marker editor).

| Action | Effect |
|--------|--------|
| **Shrink** | Cycles the selected card's non-underlined text down a font size (8→7→6→5→4→Normal), keeping the underlined cut, highlights, cites, and headings full-size. Press again to shrink further. |
| **Condense** | Collapses the selection per your Condense settings (merge paragraphs / pilcrows / retain paragraphs). |
| **Uncondense** | Reverses Condense — restores every paragraph break Rostrum marked. |
| **Unshrink** | Reverses Shrink — resets the selected card's text back to its Normal size. |
| **Options** | Modes, the omission-marker editor, and one-click mode buttons. |

Reversal is lossless and add-in-free: Condense records what it merged in
**self-describing OOXML markers** (no sidecar), so Uncondense/Unshrink reconstruct the
original from the document itself. Needs only the core OOXML round-trip, so it runs
wherever the suite loads.

---

## Settings

A first-class ribbon group (gear icon, leftmost) opening a deep-linked **informational**
pane. It shows how Rostrum loads on every document — via the **Trusted-Catalog install**,
not a per-document toggle — and where to turn that off (the host's Trust Center). It
contributes no document-mutating commands.

---

## Requirements & install

- **Word for Windows or Mac (desktop).** Manifest floor: `WordApiDesktop 1.2` +
  `WordApi 1.4` (the hide mechanism + the manifest store + the Track-Changes gate).
  Apply Styles additionally needs Word 1.5+.
- Node ≥ 18 for development.

### Sideload (development)

```bash
npm install
npm start          # builds, starts the HTTPS dev server on :3000, and sideloads
```

`npm start` (office-addin-debugging) trusts the dev certificate, serves the bundles
over HTTPS (Office requires HTTPS), and sideloads `manifest.xml` into Word. Then in
Word, use the **Rostrum** ribbon tab — the **Hide / Show All / Apply Styles** buttons
directly, or **Options** on any group to open that tool's pane.

To stop: `npm stop`. To rebuild without sideloading: `npm run build` (production) or
`npm run build:dev`.

### Install for the public (sideload from GitHub Pages)

End users **host nothing** — the add-in runs from the hosted bundle; they just register one
manifest file in Word, once. The install page with step-by-step instructions for **Windows and
Mac** lives at:

> **https://andrewtjin.github.io/rostrum/**

It links the production `manifest.xml` and walks through the Trusted-Add-in-Catalog (Windows) /
`wef` folder (Mac) sideload. The live build is **v0.3.0.1**. This is **Stage A** distribution
(motivated early adopters); a future **Stage B** adds one-click AppSource install.

### Producing the production manifest

The committed `manifest.xml` is the **dev** manifest (`https://localhost:3000`). The production
manifest is a build artifact, generated against the hosted origin — never committed, so the
manifest **drift test stays green**:

```bash
npm run build                                                            # dist/ (hashed bundles, assets, landing page)
npm run gen:manifest:prod -- --origin=https://andrewtjin.github.io/rostrum   # → dist/manifest.xml
```

`gen:manifest:prod` rebases every `SourceLocation` / icon / support URL onto the origin and stamps
the production `<Id>` (distinct from dev, so both can be sideloaded on one machine). It **fails
loudly** if `--origin` is missing or not `https://`. CI (`.github/workflows/deploy-pages.yml`)
runs exactly these steps on every push to `master` and publishes `dist/` to GitHub Pages — gated
on green tests + a clean typecheck. (One-time: repo **Settings ▸ Pages ▸ Source = GitHub
Actions**.)

---

## Reversibility (the core guarantee)

Invisibility Mode hides text with the **Hidden font attribute**, the same one Word's own
Font dialog writes. So a recipient who *doesn't have the add-in* can still recover
everything:

1. Select the text (Ctrl+A for the whole document).
2. **Home ▸ Font dialog (Ctrl+D) ▸ clear the _Hidden_ checkbox**, or toggle
   **Home ▸ ¶ (Show/Hide)** to view hidden text.

Caveat — **"Show hidden text" display setting:** if a reader has *File ▸ Options ▸
Display ▸ Hidden text* enabled, hidden runs render (greyed/dotted-underlined) rather
than disappearing. That's a viewer preference, not a Rostrum bug; Show All still
removes the hiding.

Condense & Shrink are likewise reversible from the document alone (Uncondense /
Unshrink), via the self-describing markers Condense writes.

---

## Debugging (this is a first-class feature)

You cannot attach a normal debugger inside a live Word host, so Rostrum ships its own
diagnostics. Open the **Diagnostics** section at the bottom of the Invisibility pane:

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
not just the generic `.message`). Every adapter/styles/ribbon failure flows through it.
In the browser console you'll see the same lines mirrored.

---

## Architecture

```
src/
├─ core/                      # the engine — PURE, Office.js-free, unit-tested
│  ├─ types.ts                #   WordPort contract + domain model
│  ├─ ooxml.ts                #   per-paragraph <w:vanish/> transforms
│  ├─ keepers.ts invisibility.ts manifest.ts settings.ts styles.ts guards.ts
│  ├─ condense.ts shrink.ts ooxmlCondense.ts   #   Condense & Shrink engine + markers
│  ├─ citeRepair.ts outline.ts progress.ts     #   cite repair, outline normalization, progress
│  ├─ debug.ts                #   the tracer (host-free, tested)
│  ├─ ooxmlPackage.ts         #   PURE whole-body split/splice + outline normalization
│  ├─ officeWordPort.ts       #   the REAL WordPort over Word.run  ← the only deep Office.js
│  └─ officeStyles.ts         #   ensureRostrumStyles (host glue + pure plan)
├─ features/                  # the SUITE registry — one headless contribution per tool
│  ├─ contributions.ts registry.ts types.ts    #   the single feature list + ribbon descriptors
│  ├─ ribbonRuntime.ts manifestGen.ts          #   shared in-flight guard / progress + manifest gen
│  ├─ invisibility/ condense/ settings/         #   live tools (contribution + pane)
│  └─ planned.ts              #   roadmap tools, advertised on the ribbon as data
├─ taskpane/                  # React shells over the controllers
│  ├─ controller.ts condenseController.ts        #   UI orchestration, TESTED
│  ├─ App.tsx Shell.tsx components/ host.tsx
│  └─ index.tsx taskpane.html
├─ commands/                  # ribbon handlers (reuse the controllers)
└─ dialog/                    # the full-window workspace surface (planned dialog tools)
```

The design principle (from Stage 1) holds: **pure policy behind a narrow port.** The
engine reasons about paragraphs as `(outline level, OOXML string)` and never touches
Word; one adapter (`officeWordPort.ts`) turns that into `Word.run` calls. Result: the
*entire* engine + adapter sequencing + UI orchestration are tested in plain Node with
a fake host — no Office mock, no browser. `npm test` → **469 tests**.

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

- **Co-authoring:** there's no reliable Office.js signal for a live co-authoring
  session. Hiding/showing while others edit may merge unpredictably; Show All is
  convergent, but run invisibility when you're the sole editor. The pane warns when
  armed.
- **Track Changes:** Hide refuses to run while Track Changes is on (a partial Undo
  could otherwise strand the document). The pane offers to toggle it off for the
  operation and restore it after.
- **Apply Styles** reflows pagination and is **not** undone by Show All — use
  Ctrl+Z. The pocket box needs `Style.borders` (desktop); the per-paragraph `w:pBdr`
  fallback is documented but not auto-applied.

---

## Testing

```bash
npm test           # 469 unit + integration tests (Node, no host)
npm run typecheck  # tsc --noEmit
npm run build      # production webpack bundle
```

Adapter/integration tests run the real engine through a fake `Word.RequestContext`
(`__tests__/fakeWord.ts`) that models Office.js's queue-then-`sync()` semantics, so
they assert the hard invariants directly: single-sync atomic commit, manifest
set-vs-add, the Track-Changes restore error, the multi-`<w:p>` guard, outline
normalization, the whole-body alignment fallback, and Condense/Shrink round-trip
losslessness.
