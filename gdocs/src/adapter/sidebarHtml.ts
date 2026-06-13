// HTML for the sidebar, the help dialog, the receipt/refusal dialog shell, and
// the diagnostics dialog (plan D11, frontendDraft Steps 4-5 + 10; step S11).
//
// These are TS template literals interpolating STRINGS (and the settings
// module's DEFAULT_KEEP_HEXES) AT BUILD TIME: the bundle ships finished HTML
// strings, so the sidebar and the menu dialogs can never drift from the copy
// deck, and the lexicon test that audits STRINGS effectively audits this UI.
// Only RUNTIME data (the verb receipts, the diagnostics report) flows in
// through function parameters — always HTML-escaped, never trusted.
//
// CONVENTIONS (frontendDraft Step 2/10):
//   * every class is r- prefixed (the Word taskpane.css precedent);
//   * amber = #b45309 on #fef3c7, reserved for degraded-but-healthy; red is
//     reserved for nothing-was-applied refusals; everything else is neutral;
//   * receipts land in aria-live="polite" regions (busy states announce via a
//     visually-hidden live mirror — .r-sr-only, ported from taskpane.css);
//   * dialogs close on Esc AND Enter with initial focus on the OK button;
//   * the sidebar's tab order is its DOM order: header -> verbs -> receipt ->
//     settings -> cheat sheet -> footer (Step 10c).
//
// The inline <script> JS itself uses NO template literals (no backticks, no
// JS-level "${"), so every "${...}" seen inside a script block is a BUILD-TIME
// TS interpolation — one rule keeps the nested-language file readable.

import { STRINGS } from "../core/strings";
import { DEFAULT_KEEP_HEXES } from "../core/settings";
import { DEFAULT_CITE_MIN_PT } from "../core/constants";
import type { RoutedDialog } from "../core/adapterPure";

// ---------------------------------------------------------------------------
// Escaping — every interpolation goes through here, including build-time
// STRINGS values ("Help & shortcuts" carries a literal ampersand).
// ---------------------------------------------------------------------------

/** Minimal HTML escape for text content and attribute values. */
function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Escape for embedding a string as a JS literal inside an inline script —
 * JSON.stringify quotes it; "<" is broken so "</script>" can never terminate
 * the surrounding block no matter what the string contains. */
function jsString(s: string): string {
  return JSON.stringify(s).replace(/</g, "\\u003c");
}

// ---------------------------------------------------------------------------
// Shared style block — inlined into every surface so dialogs and the sidebar
// render identically (there is no external stylesheet in a paste-in script).
// ---------------------------------------------------------------------------

const BASE_CSS = `
  body { font: 13px/1.45 Arial, sans-serif; color: #202124; margin: 0; padding: 10px 12px; }
  h1, h2 { font-weight: bold; margin: 0 0 4px; }
  .r-title { font-size: 16px; }
  .r-card-title { font-size: 13px; }
  .r-state { margin: 0 0 10px; color: #5f6368; }
  .r-micro { color: #5f6368; font-size: 12px; margin: 6px 0 0; }
  .r-btn { min-height: 32px; padding: 4px 14px; font: inherit; cursor: pointer;
           border: 1px solid #dadce0; border-radius: 4px; background: #fff; }
  .r-btn:disabled { color: #9aa0a6; cursor: default; }
  .r-btn-primary { background: #1a73e8; border-color: #1a73e8; color: #fff; }
  .r-btn-primary:disabled { background: #9aa0a6; border-color: #9aa0a6; }
  .r-link { background: none; border: none; padding: 0; font: inherit; color: #1a73e8;
            cursor: pointer; text-decoration: underline; }
  /* Amber = degraded but healthy; red = nothing was applied (Step 2 tokens). */
  .r-amber { background: #fef3c7; color: #b45309; }
  .r-red { background: #fdeeee; color: #a50e0e; }
  /* Visually hidden but exposed to assistive tech (taskpane.css precedent). */
  .r-sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px;
               overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; border: 0; }
`;

const SIDEBAR_CSS = `
  .r-verbs { display: flex; gap: 8px; margin-bottom: 10px; }
  .r-verbs .r-btn { flex: 1; }
  .r-receipt { border: 1px solid #dadce0; border-radius: 4px; padding: 8px; margin-bottom: 12px; }
  .r-receipt p { margin: 0; }
  .r-receipt details { margin-top: 6px; }
  .r-card { border-top: 1px solid #dadce0; padding: 10px 0; }
  /* 10 columns mirrors the Docs picker grid; 24px targets (Step 10b). */
  .r-grid { display: grid; grid-template-columns: repeat(10, 1fr); gap: 2px; margin-top: 6px; }
  .r-swatch { display: flex; align-items: center; justify-content: center;
              min-width: 24px; min-height: 24px; cursor: pointer; }
  .r-chip { display: inline-block; width: 9px; height: 14px; margin-left: 1px;
            border: 1px solid #bbb; }
  .r-check { margin: 0; }
  .r-master, .r-toggle { display: block; margin-top: 4px; }
  .r-cs-table { border-collapse: collapse; margin-top: 6px; width: 100%; }
  .r-cs-table th, .r-cs-table td { text-align: left; padding: 3px 6px 3px 0; vertical-align: middle; }
  /* Each left cell renders IN its style, scaled ~60% to fit 300px (Step 4.5):
     Pocket 26pt/Hat 22pt/Block 16pt/Tag 14pt/Normal 11pt -> 16/14/12/11/11px. */
  .r-cs-pocket { font-size: 16px; font-weight: bold; border: 1px solid #444; padding: 0 4px; }
  .r-cs-hat { font-size: 14px; font-weight: normal; }
  .r-cs-block { font-size: 12px; font-weight: normal; }
  .r-cs-tag { font-size: 11px; font-weight: bold; }
  .r-cs-normal { font-size: 11px; font-weight: normal; }
  .r-key-mac { display: none; }
  .r-mac .r-key-mac { display: inline; }
  .r-mac .r-key-win { display: none; }
  .r-footer { border-top: 1px solid #dadce0; padding-top: 8px; margin-top: 4px;
              display: flex; gap: 10px; align-items: center; color: #5f6368; font-size: 12px; }
`;

const DIALOG_CSS = `
  .r-dialog-body { margin: 4px 0 14px; }
  .r-banner { padding: 8px; border-radius: 4px; }
  .r-actions { text-align: right; }
  textarea.r-report { width: 100%; height: 240px; box-sizing: border-box; font: 12px/1.4 monospace; }
`;

/** Dialog keyboard contract (Step 5): initial focus on OK; Esc AND Enter
 * close. Shared verbatim by every modal this module emits. */
const DIALOG_SCRIPT = `
<script>
  (function () {
    var ok = document.getElementById('r-ok');
    if (ok) ok.focus();
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' || e.key === 'Enter') google.script.host.close();
    });
  })();
</script>`;

// ---------------------------------------------------------------------------
// Cheat sheet (shared by the sidebar card and the help dialog)
// ---------------------------------------------------------------------------

/** Left-cell style class per row key — the card doubles as a visual spec, so
 * each name renders in (a scaled version of) the style it teaches. */
const CHEAT_ROW_CLASS: Record<keyof typeof STRINGS.sidebar.cheatSheet.rows, string> = {
  pocket: "r-cs-pocket",
  hat: "r-cs-hat",
  block: "r-cs-block",
  tag: "r-cs-tag",
  normal: "r-cs-normal"
};

/**
 * The 5-row shortcut table. `bothPlatforms` = help-dialog mode (plan A15: help
 * lists Windows AND Mac side by side; the sidebar platform-detects and shows
 * one column with a manual toggle).
 */
function cheatSheetTable(bothPlatforms: boolean): string {
  const rows = STRINGS.sidebar.cheatSheet.rows;
  const body = (Object.keys(rows) as Array<keyof typeof rows>)
    .map((key) => {
      const row = rows[key];
      const chord = bothPlatforms
        ? `<td>${escapeHtml(row.windows)}</td><td>${escapeHtml(row.mac)}</td>`
        : `<td><span class="r-key-win">${escapeHtml(row.windows)}</span>` +
          `<span class="r-key-mac">${escapeHtml(row.mac)}</span></td>`;
      return (
        `<tr><th scope="row" class="${CHEAT_ROW_CLASS[key]}">${escapeHtml(row.name)}</th>` +
        `<td>${escapeHtml(row.verbatim)}</td>${chord}</tr>`
      );
    })
    .join("");
  // Header row kept sr-only in the sidebar (the card is self-evident and 300px
  // is tight) but visible in the roomier help dialog.
  const head = bothPlatforms
    ? "<tr><th>Style</th><th>Verbatim</th><th>Windows</th><th>Mac</th></tr>"
    : `<tr class="r-sr-only"><th>Style</th><th>Verbatim</th><th>Shortcut</th></tr>`;
  return `<table class="r-cs-table">${head}${body}</table>`;
}

// ---------------------------------------------------------------------------
// Keep-color grid
// ---------------------------------------------------------------------------

/** Accessible names for the swatches (STRINGS-owned; unnamed tints fall back
 * to their hex — see the deck's rationale). */
const SWATCH_NAMES: Record<string, string> = STRINGS.sidebar.keepColors.names;

/**
 * One labeled checkbox per default keep hex (plan A8: the closed set IS the
 * UI). The checkbox itself conveys state (color-blind safe — Step 10b); the
 * chip only previews the color; the name rides in the accessible label and
 * the hover title. Insertion order of DEFAULT_KEEP_HEXES is the Docs picker's
 * grid order, which the 10-column layout mirrors.
 */
function keepColorGrid(): string {
  return [...DEFAULT_KEEP_HEXES]
    .map((hex) => {
      const name = SWATCH_NAMES[hex] ?? hex;
      return (
        `<label class="r-swatch" title="${escapeHtml(name)}">` +
        `<input type="checkbox" class="r-check" data-hex="${escapeHtml(hex)}" checked>` +
        `<span class="r-chip" style="background:${escapeHtml(hex)}"></span>` +
        `<span class="r-sr-only">${escapeHtml(name)}</span></label>`
      );
    })
    .join("");
}

// ---------------------------------------------------------------------------
// The sidebar
// ---------------------------------------------------------------------------

/**
 * The full 300px sidebar document (frontendDraft Step 4, top-to-bottom by
 * frequency of need). All copy is STRINGS; all state is fetched at runtime
 * via google.script.run (the HTML itself is baked at build time).
 */
export function sidebarHtml(): string {
  const S = STRINGS.sidebar;
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>${BASE_CSS}${SIDEBAR_CSS}</style>
</head>
<body>
  <header>
    <h1 class="r-title">${escapeHtml(S.title)}</h1>
    <p id="r-state" class="r-state" aria-live="polite"></p>
  </header>

  <div class="r-verbs">
    <button id="r-hide" class="r-btn r-btn-primary" type="button">${escapeHtml(S.hideButton)}</button>
    <button id="r-showall" class="r-btn" type="button">${escapeHtml(S.showAllButton)}</button>
  </div>
  <!-- Busy announcements go to a screen-reader mirror so progress is never
       spinner-only (Step 10f). -->
  <div id="r-live" class="r-sr-only" aria-live="polite"></div>

  <!-- The receipt region doubles as the consumed-first-click teaching slot
       while empty (plan A15). -->
  <div id="r-receipt" class="r-receipt" aria-live="polite">
    <p id="r-receipt-text">${escapeHtml(S.emptyState)}</p>
    <details id="r-amber-details" hidden>
      <summary>${escapeHtml(S.amberDetails)}</summary>
      <p>${escapeHtml(S.amberExplain)}</p>
    </details>
  </div>

  <section class="r-card" aria-labelledby="r-keep-title">
    <h2 id="r-keep-title" class="r-card-title">${escapeHtml(S.keepColors.title)}</h2>
    <label class="r-master"><input type="checkbox" id="r-any"> ${escapeHtml(S.keepColors.anyHighlight)}</label>
    <div class="r-grid" role="group" aria-labelledby="r-keep-title">${keepColorGrid()}</div>
    <p class="r-micro">${escapeHtml(S.keepColors.microcopy)}</p>
    <button id="r-savedefault" class="r-link" type="button">${escapeHtml(S.keepColors.saveDefault)}</button>
  </section>

  <section class="r-card" aria-labelledby="r-spacing-title">
    <h2 id="r-spacing-title" class="r-card-title">${escapeHtml(S.spacing.title)}</h2>
    <label class="r-toggle"><input type="checkbox" id="r-collapse"> ${escapeHtml(S.spacing.toggle)}</label>
    <p class="r-micro">${escapeHtml(S.spacing.microcopy)}</p>
  </section>

  <section class="r-card" aria-labelledby="r-cs-title">
    <h2 id="r-cs-title" class="r-card-title">${escapeHtml(S.cheatSheet.title)}</h2>
    ${cheatSheetTable(false)}
    <button id="r-platform" class="r-link" type="button">${escapeHtml(S.cheatSheet.showMac)}</button>
    <p class="r-micro">${escapeHtml(S.cheatSheet.screenReaderNote)}</p>
  </section>

  <footer class="r-footer">
    <span>${escapeHtml(S.footer.version)}</span>
    <button id="r-help" class="r-link" type="button">${escapeHtml(S.footer.help)}</button>
    <!-- "Report a problem" routes to Diagnostics on purpose: the copy-paste
         report IS the remote-user reporting channel (plan D12). -->
    <button id="r-report" class="r-link" type="button">${escapeHtml(S.footer.reportProblem)}</button>
  </footer>

<script>
  (function () {
    // Build-time injected copy the script needs at runtime.
    var FALLBACK_ERROR = ${jsString(STRINGS.errors.unknown.body)};
    var LABEL_SHOW_MAC = ${jsString(STRINGS.sidebar.cheatSheet.showMac)};
    var LABEL_SHOW_WIN = ${jsString(STRINGS.sidebar.cheatSheet.showWindows)};
    var BUSY = {
      'r-hide': { label: ${jsString(STRINGS.sidebar.hideBusy)}, live: ${jsString(STRINGS.sidebar.hideBusyLive)} },
      'r-showall': { label: ${jsString(STRINGS.sidebar.showAllBusy)}, live: ${jsString(STRINGS.sidebar.showAllBusyLive)} }
    };

    function byId(id) { return document.getElementById(id); }

    // ---- platform detect + manual override (plan A15) ----
    var mac = /mac/i.test(navigator.userAgent || '');
    function renderPlatform() {
      document.body.className = mac ? 'r-mac' : '';
      byId('r-platform').textContent = mac ? LABEL_SHOW_WIN : LABEL_SHOW_MAC;
    }
    byId('r-platform').addEventListener('click', function () { mac = !mac; renderPlatform(); });
    renderPlatform();

    // ---- settings round-trip ----
    // The last server-resolved settings: fields without UI (citeMinPt,
    // structuralCite) round-trip through saves untouched.
    var current = null;
    function applySettings(s) {
      current = s;
      byId('r-any').checked = s.keepMode === 'anyHighlight';
      byId('r-collapse').checked = !!s.collapseSpacing;
      var boxes = document.querySelectorAll('.r-grid input');
      for (var i = 0; i < boxes.length; i++) {
        boxes[i].checked = s.keepColors.indexOf(boxes[i].getAttribute('data-hex')) >= 0;
      }
    }
    function collectSettings() {
      var hexes = [];
      var boxes = document.querySelectorAll('.r-grid input');
      for (var i = 0; i < boxes.length; i++) {
        if (boxes[i].checked) hexes.push(boxes[i].getAttribute('data-hex'));
      }
      return JSON.stringify({
        keepMode: byId('r-any').checked ? 'anyHighlight' : 'set',
        keepColors: hexes,
        citeMinPt: current ? current.citeMinPt : ${DEFAULT_CITE_MIN_PT},
        structuralCite: current ? current.structuralCite : true,
        collapseSpacing: byId('r-collapse').checked
      });
    }
    function saveSettings(alsoDefault) {
      google.script.run.rostrumSaveSettings(collectSettings(), alsoDefault === true);
    }

    // ---- state line ----
    function applyState(s) {
      if (!s) return;
      byId('r-state').textContent = s.stateLine;
      try { applySettings(JSON.parse(s.settingsJson)); } catch (e) { /* stale blob: keep UI as-is */ }
    }
    function refreshState() {
      google.script.run.withSuccessHandler(applyState).rostrumSidebarState();
    }

    // ---- receipts ----
    function renderReceipt(d) {
      var box = byId('r-receipt');
      box.className = 'r-receipt' + (d.severity === 'amber' ? ' r-amber' : d.severity === 'red' ? ' r-red' : '');
      byId('r-receipt-text').textContent = d.body;
      var details = byId('r-amber-details');
      if (d.severity === 'amber') { details.hidden = false; } else { details.hidden = true; details.open = false; }
    }

    // ---- verbs (busy states kill the double-click problem — Step 4.2) ----
    var busy = false;
    function runVerb(buttonId, serverFn) {
      if (busy) return;
      busy = true;
      var hideBtn = byId('r-hide');
      var showBtn = byId('r-showall');
      var btn = byId(buttonId);
      var original = btn.textContent;
      hideBtn.disabled = true;
      showBtn.disabled = true;
      btn.textContent = BUSY[buttonId].label;
      byId('r-live').textContent = BUSY[buttonId].live;
      var done = function (d) {
        busy = false;
        hideBtn.disabled = false;
        showBtn.disabled = false;
        btn.textContent = original;
        byId('r-live').textContent = '';
        // A null result means the platform dropped the return value (or the
        // call failed in transport): fall back to the unknown-error copy —
        // the engine's invariant (Show All recovers) keeps it truthful.
        renderReceipt(d || { body: FALLBACK_ERROR, severity: 'red' });
        refreshState();
      };
      google.script.run
        .withSuccessHandler(done)
        .withFailureHandler(function () { done({ body: FALLBACK_ERROR, severity: 'red' }); })[serverFn]();
    }

    byId('r-hide').addEventListener('click', function () { runVerb('r-hide', 'rostrumHideFromSidebar'); });
    byId('r-showall').addEventListener('click', function () { runVerb('r-showall', 'rostrumShowAllFromSidebar'); });
    byId('r-savedefault').addEventListener('click', function () { saveSettings(true); });
    byId('r-any').addEventListener('change', function () { saveSettings(false); });
    byId('r-collapse').addEventListener('change', function () { saveSettings(false); });
    var grid = document.querySelectorAll('.r-grid input');
    for (var i = 0; i < grid.length; i++) {
      grid[i].addEventListener('change', function () { saveSettings(false); });
    }
    byId('r-help').addEventListener('click', function () { google.script.run.rostrumHelp(); });
    byId('r-report').addEventListener('click', function () { google.script.run.rostrumDiagnostics(); });

    refreshState();
  })();
</script>
</body>
</html>`;
}

// ---------------------------------------------------------------------------
// Dialogs
// ---------------------------------------------------------------------------

/** Shared modal page shell: one style block, one keyboard contract. */
function dialogPage(inner: string): string {
  return `<!DOCTYPE html>
<html>
<head><base target="_top"><style>${BASE_CSS}${DIALOG_CSS}</style></head>
<body>
${inner}
<div class="r-actions"><button id="r-ok" class="r-btn r-btn-primary" type="button" onclick="google.script.host.close()">${escapeHtml(
    STRINGS.dialogs.ok
  )}</button></div>
${DIALOG_SCRIPT}
</body>
</html>`;
}

/**
 * The receipt/refusal modal (frontendDraft Step 5: the two verb dialogs share
 * one shell; severity drives the banner — red for refusals, amber for
 * degraded-but-healthy receipts, none for plain receipts). The title also
 * rides in the modal's chrome (the adapter passes it to showModalDialog); a
 * refusal repeats it as the red header the spec calls for.
 */
export function dialogHtml(d: RoutedDialog): string {
  const banner = d.severity === "red" ? "r-banner r-red" : d.severity === "amber" ? "r-banner r-amber" : "";
  const heading = d.dialog === "refusal" ? `<h1 class="r-title">${escapeHtml(d.title)}</h1>` : "";
  return dialogPage(
    `<div class="${banner}">${heading}<p class="r-dialog-body">${escapeHtml(d.body)}</p></div>`
  );
}

/** The Help & shortcuts dialog (560x480): first-click rule FIRST (plan A15),
 * honest mechanism + escape hatch (plan D15), both platforms' chords (A15). */
export function helpHtml(): string {
  const H = STRINGS.help;
  return dialogPage(
    `<h1 class="r-title">${escapeHtml(STRINGS.dialogs.helpTitle)}</h1>
<p><strong>${escapeHtml(H.firstClick)}</strong></p>
<p>${escapeHtml(H.whatHideDoes)}</p>
<p>${escapeHtml(H.readingGuidance)}</p>
<p>${escapeHtml(H.escapeHatch)}</p>
<p>${escapeHtml(H.teamNorm)}</p>
<h2 class="r-card-title">${escapeHtml(STRINGS.sidebar.cheatSheet.title)}</h2>
${cheatSheetTable(true)}
<p class="r-micro">${escapeHtml(STRINGS.sidebar.cheatSheet.screenReaderNote)}</p>`
  );
}

/** The Diagnostics modal: one copy-me textarea (plan D12 — the report is the
 * wet round's report-back payload). Click selects all, so "copy everything
 * below" is one click + Ctrl+C. */
export function diagnosticsHtml(reportText: string): string {
  return dialogPage(
    `<h1 class="r-title">${escapeHtml(STRINGS.menu.diagnostics)}</h1>
<p>${escapeHtml(STRINGS.diagnostics.copyHint)}</p>
<textarea class="r-report" readonly onclick="this.select()">${escapeHtml(reportText)}</textarea>`
  );
}
