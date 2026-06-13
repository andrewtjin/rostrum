// The ONE file that touches Google globals (plan D8/A4, step S11).
//
// Everything decidable lives in core (controller for the verbs, adapterPure
// for adapter-side decisions, sidebarHtml for markup): this module only
//   (a) implements DocsPort over the Advanced Docs Service + Properties,
//   (b) extracts host facts the pure code cannot reach (the user's selection,
//       UI dialog answers), and
//   (c) dispatches results/errors to dialogs via try/catch + routeError.
// It is excluded from unit coverage (it cannot run without the Apps Script
// host) and proven by typecheck:gdocs, the build pipeline, and the wet round.
//
// ASYNC NOTE: core's verbs are async over the DocsPort seam, but every port
// method here wraps a SYNCHRONOUS Apps Script call in an already-resolved
// promise — the await chains drain as microtasks within the same execution,
// and the V8 runtime keeps an execution alive until its returned promise
// settles. google.script.run resolution of promise-returning server functions
// is a wet-checklist item; the sidebar tolerates a dropped return value (it
// falls back to the unknown-error receipt and re-reads the state line).

import {
  buildDiagnostics,
  CALL_MAP,
  classifyBatchError,
  countDebateHeadings,
  diagnosticsReadFacts,
  FONT_FLOOR_TRY_SIZES_PT,
  hideDialog,
  markCiteDialog,
  planMarkCiteFromPicks,
  probeReadbackSizePt,
  renderDiagnosticsText,
  routeError,
  SETTINGS_PROPERTY_KEY,
  showAllDialog,
  sidebarState,
  STATE_FIELDS_MASK,
  stylesDialog
} from "../core/adapterPure";
import type { FontFloorReading, RoutedDialog, SelectionPick } from "../core/adapterPure";
import { applyStyles, hide, showAll } from "../core/controller";
import { chunkGroups } from "../core/guards";
import { DOC_FIELDS_MASK } from "../core/parse";
import { resolveSettings, serializeSettings } from "../core/settings";
import { consentPrompt, docStateLine, STRINGS, stylesConfirm } from "../core/strings";
import { RevisionMismatchError } from "../core/types";
import type { DocsPort, DocsRequest, GShowAllResult } from "../core/types";
import { dialogHtml, diagnosticsHtml, helpHtml, sidebarHtml } from "./sidebarHtml";

// ---------------------------------------------------------------------------
// DocsPort over the Advanced Docs Service (plan D4/A13)
// ---------------------------------------------------------------------------

/** Resolved per call, never at module load: the bundle's top level runs under
 * onOpen's restricted auth mode, where host calls may be unavailable. */
function activeDocId(): string {
  return DocumentApp.getActiveDocument().getId();
}

/**
 * One revision-guarded batchUpdate. Returns the POST-apply revision id from
 * the response's writeControl (plan A13 — never a fresh get); a response that
 * omits it falls back to the guard we sent, which can only fail VISIBLY (a
 * later chunk would mismatch and refuse), never corrupt. Raw throws are
 * classified here so the controller's retry protocol sees exactly the two
 * error classes it understands.
 */
function rawBatchUpdate(requests: ReadonlyArray<object>, requiredRevisionId: string): string {
  try {
    const res = Docs.Documents.batchUpdate({ requests, writeControl: { requiredRevisionId } }, activeDocId());
    return res.writeControl?.requiredRevisionId ?? requiredRevisionId;
  } catch (e) {
    throw classifyBatchError(e);
  }
}

/** The verb read: ONE masked documents.get (plan A13; the mask lives in
 * parse.ts — the single place that knows what the engine reads). */
function fetchVerbDocument(): unknown {
  return Docs.Documents.get(activeDocId(), { includeTabsContent: true, fields: DOC_FIELDS_MASK });
}

/**
 * The live port. Settings read at BLOB granularity: this doc's properties win
 * whole, else the device default UserProperties blob (the "Save as my
 * default" tier) — per-FIELD precedence across tiers happens only where both
 * blobs are visible (rostrumSidebarState), a documented simplification: the
 * sidebar always saves complete blobs, so a partial doc blob never occurs in
 * practice.
 */
const PORT: DocsPort = {
  fetchDocument: () => Promise.resolve(fetchVerbDocument()),
  applyBatch: (requests: DocsRequest[], requiredRevisionId: string) =>
    Promise.resolve({ revisionId: rawBatchUpdate(requests, requiredRevisionId) }),
  readSettingsJson: () =>
    Promise.resolve(
      PropertiesService.getDocumentProperties().getProperty(SETTINGS_PROPERTY_KEY) ??
        PropertiesService.getUserProperties().getProperty(SETTINGS_PROPERTY_KEY)
    ),
  writeSettingsJson: (json: string) => {
    PropertiesService.getDocumentProperties().setProperty(SETTINGS_PROPERTY_KEY, json);
    return Promise.resolve();
  }
};

// ---------------------------------------------------------------------------
// Dialog plumbing — render what core routed, decide nothing.
// ---------------------------------------------------------------------------

/** Receipt/refusal modal (~360x220, frontendDraft Step 5 sizing). */
function showRouted(d: RoutedDialog): void {
  DocumentApp.getUi().showModalDialog(HtmlService.createHtmlOutput(dialogHtml(d)).setWidth(360).setHeight(220), d.title);
}

/** The menu-verb shell: run, route, show — the whole failure model is
 * try/catch -> routeError (plan S11's "zero branching" rule). `run` may
 * resolve null to mean "show nothing" (a cancelled confirm). */
async function menuVerb(run: () => Promise<RoutedDialog | null>): Promise<void> {
  let routed: RoutedDialog | null;
  try {
    routed = await run();
  } catch (e) {
    routed = routeError(e);
  }
  if (routed !== null) showRouted(routed);
}

// ---------------------------------------------------------------------------
// Shared verb flows (menu and sidebar render the SAME routed objects)
// ---------------------------------------------------------------------------

/**
 * Show All with the consent handshake (plan A14): a needsConsent outcome asks
 * via a native YES/NO alert (works from both the menu and a sidebar-initiated
 * server call), then re-invokes with the answer — the prescribed two-step, the
 * one branch the adapter owns because only it can ask a question.
 */
async function runShowAll(): Promise<GShowAllResult> {
  const first = await showAll(PORT);
  if (first.kind === "done") return first.result;
  const ui = DocumentApp.getUi();
  const answer = ui.alert(STRINGS.dialogs.consentTitle, consentPrompt(first.unrecordedTinyCount), ui.ButtonSet.YES_NO);
  const second = await showAll(PORT, { sweepUnrecorded: answer === ui.Button.YES });
  // An answered re-invoke cannot ask again (controller contract); the throw is
  // a belt-and-suspenders rail that routes to the unknown-error copy.
  if (second.kind !== "done") throw new Error("showAll consent loop did not settle");
  return second.result;
}

/**
 * Apply styles with the first-run confirm (frontendDraft Step 5): counted
 * blast radius when existing H1-4 headings exist, zero-friction otherwise.
 * Null = the user cancelled (no dialog at all). The confirm needs its own
 * read before the verb's own fetch — a deliberate second get on this one
 * verb (the count must describe the doc the user is agreeing about).
 */
async function runApplyStyles(): Promise<RoutedDialog | null> {
  const headingCount = countDebateHeadings(await PORT.fetchDocument());
  if (headingCount > 0) {
    const ui = DocumentApp.getUi();
    const answer = ui.alert(STRINGS.dialogs.stylesConfirmTitle, stylesConfirm(headingCount), ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) return null;
  }
  return stylesDialog(await applyStyles(PORT));
}

// ---------------------------------------------------------------------------
// Selection extraction for Mark cite (the one host fact pure code can't reach)
// ---------------------------------------------------------------------------

/**
 * Minimal structural view of the DocumentApp element tree. The official
 * typings split Element/ContainerElement/Paragraph into sibling interfaces
 * that don't cross-assign, so the tree walk goes through this local seam
 * (method-parameter bivariance makes the real objects assignable).
 */
interface DocTreeNode {
  getParent(): DocTreeContainer | null;
}
interface DocTreeContainer extends DocTreeNode {
  getChildIndex(child: DocTreeNode): number;
  getChild(childIndex: number): GoogleAppsScript.Document.Element;
}

/** Structural path from the body root to `el` ("3/0/2") — the identity key
 * for matching host elements (Apps Script wrappers have no object identity:
 * two lookups of the same paragraph are different objects). */
function pathKey(el: DocTreeNode): string {
  const path: number[] = [];
  let cur: DocTreeNode = el;
  for (;;) {
    const parent = cur.getParent();
    if (parent === null) break;
    path.unshift(parent.getChildIndex(cur));
    cur = parent;
  }
  return path.join("/");
}

/**
 * The flattened-paragraph ordinal of `paragraph` — Body.getParagraphs()
 * returns every paragraph (list items and table cells included) in document
 * order, which matches parse.ts's flattening order, so positions align with
 * GDoc.paragraphs. (Order parity is a documented platform assumption on the
 * wet checklist.) -1 = not found; the pure side drops such picks.
 */
function paragraphOrdinalOf(paragraph: DocTreeNode, body: GoogleAppsScript.Document.Body): number {
  const key = pathKey(paragraph);
  const all = body.getParagraphs();
  for (let i = 0; i < all.length; i++) {
    if (pathKey(all[i] as unknown as DocTreeNode) === key) return i;
  }
  return -1;
}

/** API-index units between a paragraph's start and `textEl`: preceding TEXT
 * siblings count their length; every other sibling (inline image, chip,
 * break) occupies exactly 1 index in the Docs API model. */
function apiUnitsBefore(textEl: DocTreeNode): number {
  const parent = textEl.getParent();
  if (parent === null) return 0;
  const index = parent.getChildIndex(textEl);
  let units = 0;
  for (let i = 0; i < index; i++) {
    const sibling = parent.getChild(i);
    units += sibling.getType() === DocumentApp.ElementType.TEXT ? sibling.asText().getText().length : 1;
  }
  return units;
}

/**
 * Lower the host selection to SelectionPicks (paragraph ordinal + API-unit
 * offsets) — all the index math and the A9 whitelist split then happen in
 * adapterPure.planMarkCiteFromPicks where they are tested. Pieces that are
 * neither text nor a paragraph (an image-only selection) are skipped: never
 * markable, exactly the whitelist's stance.
 */
function extractPicks(selection: GoogleAppsScript.Document.Range, body: GoogleAppsScript.Document.Body): SelectionPick[] {
  const picks: SelectionPick[] = [];
  for (const re of selection.getRangeElements()) {
    const el = re.getElement();
    const type = el.getType();
    if (type === DocumentApp.ElementType.TEXT) {
      const node = el as unknown as DocTreeNode;
      const parent = node.getParent();
      if (parent === null) continue;
      const ordinal = paragraphOrdinalOf(parent, body);
      if (ordinal < 0) continue;
      const before = apiUnitsBefore(node);
      if (re.isPartial()) {
        // RangeElement offsets are inclusive char positions within the Text.
        picks.push({
          paragraphOrdinal: ordinal,
          startOffset: before + re.getStartOffset(),
          endOffset: before + re.getEndOffsetInclusive() + 1
        });
      } else {
        picks.push({ paragraphOrdinal: ordinal, startOffset: before, endOffset: before + el.asText().getText().length });
      }
    } else if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
      const ordinal = paragraphOrdinalOf(el as unknown as DocTreeNode, body);
      if (ordinal >= 0) picks.push({ paragraphOrdinal: ordinal, startOffset: 0, endOffset: null });
    }
  }
  return picks;
}

// ---------------------------------------------------------------------------
// Public entry points — one exported function per CALL_MAP entry (the build's
// shims delegate to these through the Rostrum bundle global).
// ---------------------------------------------------------------------------

/** Simple trigger: menu ONLY (AuthMode-restricted — never a sidebar from
 * onOpen, frontendDraft Step 3), built from the same CALL_MAP the build shims
 * come from, labels from STRINGS alone. */
export function onOpen(_e?: unknown): void {
  const menu = DocumentApp.getUi().createMenu(STRINGS.sidebar.title);
  for (const entry of CALL_MAP) {
    if (entry.label === null) continue;
    if (entry.separatorBefore) menu.addSeparator();
    menu.addItem(entry.label, entry.fn);
  }
  menu.addToUi();
}

export async function rostrumHide(): Promise<void> {
  await menuVerb(async () => hideDialog(await hide(PORT)));
}

export async function rostrumShowAll(): Promise<void> {
  await menuVerb(async () => showAllDialog(await runShowAll()));
}

export async function rostrumApplyStyles(): Promise<void> {
  await menuVerb(runApplyStyles);
}

export async function rostrumMarkCite(): Promise<void> {
  await menuVerb(async () => {
    const docApp = DocumentApp.getActiveDocument();
    const selection = docApp.getSelection() as GoogleAppsScript.Document.Range | null;
    const picks = selection === null ? [] : extractPicks(selection, docApp.getBody());
    const plan = planMarkCiteFromPicks(await PORT.fetchDocument(), picks);
    // Apply through the soft chunk cap like every other verb (plan A11.viii):
    // each cite write is independent, so one group per request lets chunkGroups
    // pack them into <=CHUNK_MAX batches with safe boundaries, revision-chained
    // across chunks. The common (small) selection is a single chunk = one apply;
    // zero requests (nothing markable) yields zero chunks — apply nothing and
    // teach via the receipt.
    let revisionId = plan.revisionId;
    for (const chunk of chunkGroups(plan.requests.map((r) => ({ requests: [r] })))) {
      ({ revisionId } = await PORT.applyBatch(chunk, revisionId));
    }
    return markCiteDialog(plan.citedParagraphs);
  });
}

/** 300px sidebar (Docs fixes sidebar width; 300 is the platform constant). */
export function rostrumOpenPanel(): void {
  DocumentApp.getUi().showSidebar(HtmlService.createHtmlOutput(sidebarHtml()).setTitle(STRINGS.sidebar.title));
}

/** Help & shortcuts modal (560x480 per frontendDraft Step 5). */
export function rostrumHelp(): void {
  DocumentApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(helpHtml()).setWidth(560).setHeight(480),
    STRINGS.dialogs.helpTitle
  );
}

/**
 * Diagnostics (plan D12/A11.ii): self-cleaning probes + read-only stats,
 * rendered into a copy-me modal. Menu-invoked ONLY — the one flow allowed to
 * insert content (a single probe char at the very end, deleted afterwards).
 * Interpretation is all adapterPure; this is the I/O choreography.
 */
export async function rostrumDiagnostics(): Promise<void> {
  try {
    const t0 = Date.now();
    const raw = await PORT.fetchDocument();
    const fetchLatencyMs = Date.now() - t0;
    const facts = diagnosticsReadFacts(raw);

    const applyLatenciesMs: number[] = [];
    /** Timed, revision-chained raw batch (diagnostics needs insertText /
     * deleteContentRange, which the engine's DocsRequest union rightly
     * excludes — case 001-F1 — so probes go through the raw seam). */
    const timed = (requests: ReadonlyArray<object>, rev: string): string => {
      const t = Date.now();
      const next = rawBatchUpdate(requests, rev);
      applyLatenciesMs.push(Date.now() - t);
      return next;
    };

    let rev = facts.revisionId;
    const probeRange = { startIndex: facts.probeIndex, endIndex: facts.probeIndex + 1 };
    const fontFloor: FontFloorReading[] = [];
    let namedStyleProbe: "ok" | "rejected" | "skipped" = "skipped";
    let inheritClear: "ok" | "failed" | "skipped" = "skipped";
    let inserted = false;
    let cleanupOk = true;
    try {
      // The probe char (self-cleaning: deleted in the finally below).
      rev = timed([{ insertText: { location: { index: facts.probeIndex }, text: "." } }], rev);
      inserted = true;

      // Font floor: try each sub-1pt size, read back what the doc reports.
      for (const size of FONT_FLOOR_TRY_SIZES_PT) {
        try {
          rev = timed(
            [{ updateTextStyle: { range: probeRange, textStyle: { fontSize: { magnitude: size, unit: "PT" } }, fields: "fontSize" } }],
            rev
          );
          fontFloor.push({ triedPt: size, appliedOk: true, readBackPt: probeReadbackSizePt(fetchVerbDocument(), facts.probeIndex) });
        } catch (e) {
          // A revision mismatch means someone is editing — abort the whole
          // probe cleanly; a plain rejection is itself the probe's answer.
          if (e instanceof RevisionMismatchError) throw e;
          fontFloor.push({ triedPt: size, appliedOk: false, readBackPt: null });
        }
      }

      // Named-style acceptance (plan D13's one unverified request): rewrite
      // HEADING_6 with its CURRENT size — a no-op when accepted. Skipped when
      // the read stated no size (a guessed size could visibly change H6 text).
      if (facts.headingSixSizePt !== null) {
        try {
          rev = timed(
            [
              {
                updateNamedStyle: {
                  namedStyle: {
                    namedStyleType: "HEADING_6",
                    textStyle: { fontSize: { magnitude: facts.headingSixSizePt, unit: "PT" } }
                  },
                  fields: "textStyle.fontSize"
                }
              }
            ],
            rev
          );
          namedStyleProbe = "ok";
        } catch (e) {
          if (e instanceof RevisionMismatchError) throw e;
          namedStyleProbe = "rejected";
        }
      }

      // Clear-to-inherit (plan D13's restore semantics), on the probe char
      // before cleanup: after the clear the read must report NO stated size.
      try {
        rev = timed([{ updateTextStyle: { range: probeRange, textStyle: {}, fields: "fontSize" } }], rev);
        inheritClear = probeReadbackSizePt(fetchVerbDocument(), facts.probeIndex) === null ? "ok" : "failed";
      } catch (e) {
        if (e instanceof RevisionMismatchError) throw e;
        inheritClear = "failed";
      }
    } finally {
      // Self-cleaning: best effort even when a probe aborted. A failed delete
      // strands one "." char; the report says so and how to fix it by hand.
      if (inserted) {
        try {
          timed([{ deleteContentRange: { range: probeRange } }], rev);
        } catch {
          cleanupOk = false;
        }
      }
    }

    const report = renderDiagnosticsText(
      buildDiagnostics({ raw, fetchLatencyMs, applyLatenciesMs, fontFloor, namedStyleProbe, inheritClear, cleanupOk })
    );
    DocumentApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(diagnosticsHtml(report)).setWidth(520).setHeight(440),
      STRINGS.menu.diagnostics
    );
  } catch (e) {
    showRouted(routeError(e));
  }
}

// ---------------------------------------------------------------------------
// google.script.run targets (sidebar). These RETURN their routed objects (the
// sidebar renders receipts inline — frontendDraft Step 5) instead of showing
// modals; failures are returned, not thrown, so copy is never lost to the
// transport's generic failure handler.
// ---------------------------------------------------------------------------

export async function rostrumHideFromSidebar(): Promise<RoutedDialog> {
  try {
    return hideDialog(await hide(PORT));
  } catch (e) {
    return routeError(e);
  }
}

export async function rostrumShowAllFromSidebar(): Promise<RoutedDialog> {
  try {
    return showAllDialog(await runShowAll());
  } catch (e) {
    return routeError(e);
  }
}

/** The sidebar's load/refresh read: state line from the tiny masked get (plan
 * A13) + the FULLY-resolved settings blob (both tiers visible here, so the
 * field-level precedence triangle applies before serialization). */
export function rostrumSidebarState(): { stateLine: string; armed: boolean; settingsJson: string } {
  const state = sidebarState(Docs.Documents.get(activeDocId(), { fields: STATE_FIELDS_MASK }));
  const settings = resolveSettings(
    PropertiesService.getDocumentProperties().getProperty(SETTINGS_PROPERTY_KEY),
    PropertiesService.getUserProperties().getProperty(SETTINGS_PROPERTY_KEY)
  );
  return {
    stateLine: docStateLine(state.hiddenRegionCount),
    armed: state.armed,
    settingsJson: serializeSettings(settings)
  };
}

/**
 * Persist sidebar settings: normalize through the same resolve/serialize pair
 * the engine reads with (corruption-tolerant by construction), write this
 * doc's properties, and optionally the device default (the "Save as my
 * default" tier). Returns the normalized blob so the sidebar can resync.
 */
export function rostrumSaveSettings(json: string, alsoDeviceDefault?: boolean): string {
  const normalized = serializeSettings(resolveSettings(typeof json === "string" ? json : null, null));
  PropertiesService.getDocumentProperties().setProperty(SETTINGS_PROPERTY_KEY, normalized);
  if (alsoDeviceDefault === true) {
    PropertiesService.getUserProperties().setProperty(SETTINGS_PROPERTY_KEY, normalized);
  }
  return normalized;
}
