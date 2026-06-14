// Master copy deck — EVERY user-visible string for the gdocs surface lives in
// this one module (plan D11 / A11.v / frontendDraft Step 11): menu labels,
// dialog titles, sidebar microcopy, help prose, receipts, refusals, and the
// sweep-consent prompt. The adapter and the inlined sidebar HTML render ONLY
// from here, so the menu-dialog and sidebar channels can never drift, and one
// mechanical test (gdocsStrings.test.ts) audits the entire deck at once.
//
// VOICE RULES (binding — frontendDraft Phase A, plan A11.v):
//   * Sentence case; no exclamation points; no emoji.
//   * Counted receipts in the shipped Word taskpane voice (condenseController
//     "Nothing to shrink here (already smallest, or all kept)." is the no-op
//     pattern). Units are PARAGRAPHS and "passages" — never engine nouns.
//   * BANNED tokens (whole word, any case): run/runs, namedrange, manifest,
//     sentinel, batchupdate, card/cards. The phrase "named style" is banned
//     too. This rules out even the VERB "to run" — copy says "use Show All
//     again", never "run Show All again". The lexicon test scans mechanically.
//   * Amber copy always ends on the doc's healthy state and never blames the
//     user's formatting ("lost records" is banned framing — plan A14).
//   * Red copy states truthfully what was (not) applied plus the escape route.
//   * Honesty posture (plan D15): hiding is shrinking, not secrecy — help says
//     so outright and carries the no-add-on escape hatch.

import {
  DocsApiError,
  GHideResult,
  GShowAllResult,
  GStylesResult,
  HiddenStateError,
  MultiTabError,
  PartialApplyError,
  RevisionConflictError,
  RevisionMismatchError,
  SuggestionsActiveError
} from "./types";
import { GDOCS_VERSION } from "./constants";

// ---------------------------------------------------------------------------
// Analytics result type — a local shape so strings.ts compiles independently
// of whichever wave-B/C agent adds the canonical GDeleteAnalyticsResult to
// types.ts.  The fields are the minimum needed by the receipt helper;
// controller.ts and adapterPure.ts will use the authoritative type.
// ---------------------------------------------------------------------------

/** Minimum shape returned by planDeleteAnalytics/deleteAnalytics that the
 *  receipt helper needs. Keeps this module decoupled from the wave-B delivery. */
interface DeleteAnalyticsResult {
  paragraphsAffected: number;
  runsDeleted: number;
}

// ---------------------------------------------------------------------------
// Shared fragments — defined once so surfaces that legally repeat a string
// (menu label == dialog title == sidebar button) can never drift apart (DRY).
// ---------------------------------------------------------------------------

/** The two product verbs are proper nouns (Word add-in parity) — always
 * capitalized exactly like this, even mid-sentence. */
const HIDE = "Hide";
const SHOW_ALL = "Show All";
const APPLY_STYLES = "Apply debate styles";
const HELP_SHORTCUTS = "Help & shortcuts";

// Analytics verb labels — shared between the menu entries and the error copy
// so the displayed verb in error messages is always the same string the user
// clicked in the menu (no drift if we rename later).
const ANALYTICIFY = "Analytic-ify";
const DELETE_ANALYTICS = "Delete analytics";
/** Dialog title for the delete confirm — no ellipsis (the confirm IS the dialog). */
const DELETE_ANALYTICS_CONFIRM_TITLE = "Delete analytics";

/**
 * The consumed-first-click rule (plan A15): Google silently spends the user's
 * FIRST menu click on the authorization prompt — the single most common silent
 * install failure. This line must appear both as the help dialog's first line
 * and as the sidebar's empty state, so it is one constant, not two strings.
 */
const FIRST_CLICK = "First time? Google uses your first click to ask permission — just click again.";

// ---------------------------------------------------------------------------
// The deck. Static strings only — anything counted goes through the typed
// format helpers below so grammar (zero/singular/plural) is handled in ONE
// place and the lexicon test can exercise every output mechanically.
// ---------------------------------------------------------------------------

export const STRINGS = {
  /** onOpen menu labels, in menu order (verbs first — frontendDraft Step 3).
   * Analytics verbs come last in the Tools group (spec §3 adapterPure.ts):
   * Analytic-ify is additive/safe; Delete analytics is destructive and
   * visually separated by a separator before it. */
  menu: {
    hide: HIDE,
    showAll: SHOW_ALL,
    applyStyles: APPLY_STYLES,
    markCite: "Mark cite",
    /** "Analytic-ify" — the user's chosen word (spec §3 strings.ts); the
     * hyphen signals that a whole paragraph is reformatted, making it a
     * more deliberate action than plain "Analyticify". */
    analyticify: ANALYTICIFY,
    /** Ellipsis signals the menu item opens a confirm dialog before acting
     * (Docs/macOS HIG convention — plan §3 adapterPure.ts). */
    deleteAnalytics: `${DELETE_ANALYTICS}…`,
    openPanel: "Open Rostrum panel…",
    helpShortcuts: HELP_SHORTCUTS,
    diagnostics: "Diagnostics"
  },

  /** Titles + shared button labels for the four dialogs (frontendDraft Step 5)
   * plus the sweep-consent confirm (plan A14) and the delete-analytics confirm
   * (spec §3). Bodies are composed: receipts come from the helpers, refusals
   * from errorMessage(). */
  dialogs: {
    receiptTitle: "Rostrum",
    helpTitle: HELP_SHORTCUTS,
    stylesConfirmTitle: APPLY_STYLES,
    consentTitle: SHOW_ALL,
    /** Title for the destructive Delete analytics confirm dialog. Body is
     * generated by deleteAnalyticsConfirm(n) — never hard-coded here so the
     * counted paragraph number is always truthful. */
    deleteAnalyticsConfirmTitle: DELETE_ANALYTICS_CONFIRM_TITLE,
    ok: "OK",
    apply: "Apply",
    cancel: "Cancel",
    /** Consent buttons are deliberately count-neutral ("Restore" not "Restore
     * them") so one label is grammatical for any passage count. */
    consentYes: "Restore",
    consentNo: "Leave as is"
  },

  /** Sidebar labels and microcopy, top-to-bottom (frontendDraft Step 4). */
  sidebar: {
    title: "Rostrum",
    /** Empty receipt region doubles as the first-time teaching slot (A15). */
    emptyState: FIRST_CLICK,
    stateNothingHidden: "This doc: nothing hidden.",
    hideButton: HIDE,
    showAllButton: SHOW_ALL,
    /** Short in-button busy labels (kills the double-click problem)… */
    hideBusy: "Hiding…",
    showAllBusy: "Restoring…",
    /** …and the aria-live announcements (busy must not be spinner-only). */
    hideBusyLive: "Hiding — this can take a few seconds on long docs.",
    showAllBusyLive: "Restoring — this can take a few seconds on long docs.",
    /** Amber banner expander: explains degraded restores WITHOUT calling the
     * user's formatting "lost" (plan A14 rewording). */
    amberDetails: "What happened?",
    amberExplain:
      "Moving or copying hidden text can separate it from its saved sizes. Tip: Show All before reorganizing.",
    keepColors: {
      title: "Keep colors",
      microcopy: "Highlighted text in checked colors stays visible when you Hide.",
      /** The explicit open-predicate master toggle (plan A8). */
      anyHighlight: "Keep any highlight color",
      saveDefault: "Save as my default",
      /**
       * Human names for the swatches a debater actually recognizes: the Docs
       * picker's base row, the classic Word highlight hexes a .docx import
       * materializes, and the light-yellow tint the UX draft calls out by
       * name (frontendDraft Step 4). Unnamed tints fall back to their hex in
       * the grid's accessible labels — a name invented for "#d5a6bd" would be
       * noise, not clarity. Keys are parse.ts's canonical lower-case
       * "#rrggbb" shape, so lookups never need re-normalization.
       */
      names: {
        "#980000": "Red berry",
        "#ff0000": "Red",
        "#ff9900": "Orange",
        "#ffff00": "Yellow",
        "#00ff00": "Green",
        "#00ffff": "Cyan",
        "#4a86e8": "Cornflower blue",
        "#0000ff": "Blue",
        "#9900ff": "Purple",
        "#ff00ff": "Magenta",
        "#ffd966": "Light yellow",
        "#00008b": "Dark blue",
        "#008b8b": "Dark cyan",
        "#006400": "Dark green",
        "#8b008b": "Dark magenta",
        "#8b0000": "Dark red",
        "#808000": "Dark yellow",
        "#808080": "Dark gray",
        "#c0c0c0": "Light gray",
        "#000000": "Black"
      }
    },
    /** Spacing-collapse switch (plan A12, default OFF): surfaced in the panel
     * so the ONE wet round can flip it after measuring the real condensation
     * ratio, without a new build. */
    spacing: {
      title: "Spacing",
      toggle: "Tighten space around hidden lines",
      microcopy:
        "Collapses extra space above and below fully hidden lines so pages condense further. Show All puts the spacing back."
    },
    /** Verbatim-muscle-memory cheat sheet. Both platforms ship in the deck;
     * the sidebar platform-detects which column to show (plan A15). */
    cheatSheet: {
      title: "Style shortcuts",
      showMac: "Show Mac shortcuts",
      showWindows: "Show Windows shortcuts",
      rows: {
        // Verbatim F-keys verified against the live Verbatim template (wet round
        // 2026-06-12): Pocket F4 / Hat F5 / Block F6 / Tag F7 / Normal-clear F12.
        // (Cite is Verbatim F8 but is not a heading chord — it rides the
        // "Mark cite" action, so it is documented in prose, not in this card,
        // whose Windows/Mac columns are native Docs heading chords only.)
        pocket: { name: "Pocket", verbatim: "F4", windows: "Ctrl+Alt+1", mac: "Cmd+Option+1" },
        hat: { name: "Hat", verbatim: "F5", windows: "Ctrl+Alt+2", mac: "Cmd+Option+2" },
        block: { name: "Block", verbatim: "F6", windows: "Ctrl+Alt+3", mac: "Cmd+Option+3" },
        tag: { name: "Tag", verbatim: "F7", windows: "Ctrl+Alt+4", mac: "Cmd+Option+4" },
        normal: { name: "Normal", verbatim: "F12", windows: "Ctrl+Alt+0", mac: "Cmd+Option+0" }
      },
      /** AT users get a documented non-keystroke path (frontendDraft Step 10e). */
      screenReaderNote:
        "Ctrl+Alt shortcuts can collide with screen readers — the Rostrum menu and panel do everything the shortcuts do."
    },
    footer: {
      version: `Rostrum v${GDOCS_VERSION}`,
      help: HELP_SHORTCUTS,
      reportProblem: "Report a problem"
    }
  },

  /** Help dialog prose. firstClick renders FIRST (plan A15). */
  help: {
    firstClick: FIRST_CLICK,
    /** Honest mechanism statement — shrink, not secrecy (plan D15). */
    whatHideDoes:
      "Hide shrinks everything except headings, cites, and highlighted text, so a long doc reads like a speech doc. Show All brings every word back.",
    readingGuidance: "Tiny seams between kept words are hidden text — Show All brings it back.",
    /** The no-add-on escape hatch (plan D15; Word reveal-panel analog). */
    escapeHatch:
      "No add-on handy? Select All (Ctrl+A, or Cmd+A on Mac), then set the font size to 11 — everything becomes visible again.",
    teamNorm: "Working with teammates: don't edit a doc while it's hidden — Show All first.",
    /**
     * The scoped content-deletion honesty (exec-review MAJOR): the F5 fallback
     * copy was made deliberately universal — "Rostrum never shrinks your text
     * away" — so it could not promise "nothing was deleted" after a partial
     * Delete analytics. That trade moved the truthful, SCOPED explanation
     * (Delete analytics is the ONLY verb that removes content, and only the
     * Analytics text you styled) out of the fallback and into the dedicated
     * surfaces. The confirm dialog carries it at the destructive moment; Help
     * is the standing reference, so it must carry it too — this line is that
     * standing reference. It also disambiguates the hide/show loop AND names the
     * real recovery path: a deleted Analytics passage comes back through File >
     * Version history (003-F10 wet finding — server-side deletes are NOT in the
     * client Ctrl+Z undo stack on the live host), never through Show All, which
     * only un-shrinks.
     */
    analyticsVerbs:
      "Analytic-ify marks whole lines as Analytics (navy, 14pt) so Hide keeps them. Delete analytics is the only action that removes content — and only the Analytics text you styled. To bring it back, use File > Version history; Show All does not."
  },

  /** Diagnostics dialog chrome. The report BODY is composed by
   * core/adapterPure.renderDiagnosticsText — a technical instrument destined
   * for the report-back chat, not product copy — so the deck carries only the
   * one product-voiced line around it (the dialog title reuses menu.diagnostics). */
  diagnostics: {
    copyHint: "Copy everything below and paste it into the chat."
  },

  /** Static receipt fragments the counted helpers compose from. They live in
   * the deck (not inline in the helpers) so the leaf walk audits them even on
   * branches a sample grid might miss. */
  receipts: {
    /** The Word idempotent-no-op pattern, verbatim shape (condenseController). */
    hideNoop: "Nothing to hide here (already hidden, or all kept).",
    /** Reconcile pass that only SURFACED keepers — distinct from the pure no-op
     * because the doc did change (plan A1). */
    hideNothingNew: "Nothing new to hide here.",
    showAllNoop: "Nothing to show here (no hidden text found).",
    markCiteNoop: "Nothing to mark here (select the cite line first).",
    /** Zero-case is a teaching moment, not a failure (frontendDraft Step 7). */
    stylesZero:
      "0 paragraphs restyled — this doc has no headings yet. Put your cursor on a line and press Ctrl+Alt+1 to make it a Pocket.",
    stylesNext: "Now style lines with Ctrl+Alt+1–4 (see the cheat sheet in the Rostrum panel).",
    /** The documented degraded path (plan A5/D13): the retro pass styled what
     * exists; only FUTURE paragraphs may miss the look. Truthful, not alarming. */
    stylesDegraded:
      "Docs declined part of the style update — existing paragraphs were styled directly, but new lines you type may not pick up the debate sizes automatically.",

    // -----------------------------------------------------------------------
    // Analytics receipts (Loop 003 — spec §3 strings.ts).
    // Counted variants live in the helpers below; static fragments live here
    // so the leaf walk audits them even when the helper is called with n=0.
    // -----------------------------------------------------------------------

    /** Analytic-ify noop: cursor was not on a paragraph (empty ordinals set).
     * Teaches the right action rather than issuing a bare "nothing to do". */
    analyticifyNoop: "Put your cursor on a line first, then use Analytic-ify.",

    /** Delete analytics noop: shown when no Analytics text exists in the doc
     * (count read before the confirm returned 0 — no confirm shown, per spec). */
    deleteAnalyticsNoop: "No Analytics text found."
  },

  /** Refusal copy — exactly one entry per error class in types.ts, keyed so
   * errorMessage() is a pure lookup. Red bodies always say what was (not)
   * applied and end with the escape route. */
  errors: {
    /** Gates Hide and Apply debate styles ONLY — Show All is deliberately
     * ungated (plan A7), so this copy never needs a Show All variant. */
    suggestions: {
      title: "Resolve suggestions first",
      body: "Resolve all suggestions first — accept or reject them (Tools > Review suggested edits), then try again. Nothing was changed."
    },
    multiTab: {
      title: "This doc has multiple tabs",
      body: "This doc uses tabs — Rostrum works in single-tab docs for now. Move your speech to a doc without extra tabs, then try again. Nothing was changed."
    },
    hiddenState: {
      title: "Hidden text present",
      body: "This doc still has hidden text — Show All first, then re-apply styles. Nothing was changed."
    },
    /** A raw adapter-level mismatch that escaped the controller's retry wrap:
     * the guarded batch did not land, so "nothing was applied" stays truthful. */
    revisionMismatch: {
      title: "The doc changed",
      body: "The doc changed while Rostrum was working — nothing was applied. Try again."
    },
    /** Per-verb because truthfulness differs: a failed Show All leaves text
     * HIDDEN — claiming "your doc is untouched" would read as success.
     * Analytics verbs (Loop 003): all retries exhausted = nothing applied, so
     * "nothing was changed; try again" is truthful for both. The deleteAnalytics
     * variant is phrased precisely: if a conflict fires, the delete never landed,
     * so Show All's unavailability does NOT need disclaiming here. */
    revisionConflict: {
      hide: {
        title: "Someone edited the doc",
        body: "Someone edited the doc while hiding — nothing was applied; your doc is untouched. Try again."
      },
      showAll: {
        title: "Someone edited the doc",
        body: "Someone edited the doc while restoring — nothing was changed yet, so your text is still hidden. Try again."
      },
      applyStyles: {
        title: "Someone edited the doc",
        body: "Someone edited the doc while styling — nothing was applied; your doc is untouched. Try again."
      },
      /** All retries exhausted while analytic-ify was waiting to land: the
       * styling write never touched the doc, so the doc is exactly as left. */
      analyticify: {
        title: "Someone edited the doc",
        body: "Someone edited the doc while applying analytics — nothing was changed; try again."
      },
      /** All retries exhausted while Delete analytics was waiting to land:
       * because not a single chunk was applied, no content was removed. */
      deleteAnalytics: {
        title: "Someone edited the doc",
        body: "Someone edited the doc while deleting analytics — nothing was changed; try again."
      }
    },
    /** Per-verb because the recovery route differs: an interrupted Hide is
     * undone by Show All; an interrupted Show All just needs finishing; an
     * interrupted styles pass is safe to repeat (idempotent). Analytics verbs
     * have asymmetric recovery: analytic-ify is safe to repeat (idempotent
     * character formatting); Delete analytics is DESTRUCTIVE — the partial
     * delete CANNOT be undone via Show All and "Delete analytics again" must
     * finish the job (spec §3 strings.ts + controller.ts plan). */
    partialApply: {
      hide: {
        title: "Hide was interrupted",
        body: "Hide was interrupted partway — some text is hidden, some is not. Use Show All to bring everything back, then try again."
      },
      showAll: {
        title: "Show All was interrupted",
        body: "Show All was interrupted partway — some text is already back. Use Show All again to finish."
      },
      applyStyles: {
        title: "Styling was interrupted",
        body: "Apply debate styles was interrupted partway — some paragraphs are restyled. Use Apply debate styles again to finish; it is safe to repeat."
      },
      /** Analytic-ify is pure character formatting (idempotent): applying it
       * again repairs the partial result with no risk of side effects. */
      analyticify: {
        title: "Analytic-ify was interrupted",
        body: "Analytic-ify was interrupted partway — some paragraphs are already styled. Use Analytic-ify again to finish; it is safe to repeat."
      },
      /** Delete analytics is the sole destructive verb: a partial delete means
       * some Analytics text was already removed and Show All cannot bring it
       * back (only File > Version history can — server-side deletes are NOT in
       * the client Ctrl+Z undo stack on the live host, 003-F10). The recovery
       * path for the rest is Delete analytics again. "Show All will not bring it
       * back" is critical safety copy — it disambiguates the usual hide/show
       * loop and prevents data confusion. */
      deleteAnalytics: {
        title: "Delete analytics was interrupted",
        body: "Delete analytics was interrupted partway — some Analytics text was already removed, and Show All will not bring it back. Use Delete analytics again to remove the rest."
      }
    },
    /** Unmapped 400s: the atomic batch means nothing landed (plan A9). */
    docsApi: {
      title: "Docs rejected the change",
      body: "Docs rejected the change — nothing was applied."
    },
    /** Truly unknown failures can't truthfully claim "nothing was applied" and
     * can't claim "content was never deleted" — because deleteAnalytics IS a
     * content-deleter and an unknown error after a partial delete could mean
     * content was removed.  This entry is kept UNIVERSAL (no verb-specific
     * claims, no "analytics" jargon) and leans on the one invariant that
     * holds for EVERY verb: Rostrum never SHRINKS text invisibly — font-size
     * damage is always visible, and Show All restores any hide-verb damage.
     * The scoped "only Delete analytics removes content" explanation lives in
     * the Delete analytics confirm body (deleteAnalyticsConfirm) and in Help
     * (help.analyticsVerbs) — the two places where it is true and timely —
     * never in this universal fallback. */
    unknown: {
      title: "Something went wrong",
      body: "Something unexpected went wrong. Rostrum never shrinks your text away — if anything looks tiny, Show All brings it back."
    }
  }
} as const;

// ---------------------------------------------------------------------------
// Grammar/format internals (private — callers depend on the helpers below,
// never on these mechanics).
// ---------------------------------------------------------------------------

/**
 * Thousands-grouped integer ("2431" → "2,431") WITHOUT Intl: receipts must
 * format identically under Node (tests) and the Apps Script V8 runtime, and
 * locale-dependent output would make the counted-pattern tests environment-
 * sensitive. Counts are non-negative integers by contract.
 */
function fmt(n: number): string {
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/** "1 paragraph" / "3 paragraphs" — pluralizes the LAST word of a phrase by
 * appending "s" (every counted noun in this deck pluralizes regularly). */
function counted(n: number, noun: string): string {
  return `${fmt(n)} ${n === 1 ? noun : noun + "s"}`;
}

/** Possessive/verb agreement bits for the Show All receipts. */
function their(n: number): string {
  return n === 1 ? "its" : "their";
}
function sizesWord(n: number): string {
  return n === 1 ? "size" : "sizes";
}
function wereWord(n: number): string {
  return n === 1 ? "was" : "were";
}

// ---------------------------------------------------------------------------
// Typed format helpers — the ONLY producers of counted user copy. Each takes
// the engine's structured result so receipt numbers can never be hand-rolled
// (and so the lexicon test can drive every branch with representative inputs).
// ---------------------------------------------------------------------------

/**
 * Hide receipt. Branches, in order of specificity:
 *   * pure no-op (nothing hidden AND nothing surfaced) — the Word pattern;
 *   * reconcile that only surfaced keepers (changed 0, newlyKept > 0);
 *   * re-hide on an armed doc — "new paragraphs" + already-hidden passages;
 *   * fresh hide — "Hid text in N of M paragraphs."
 * A newly-kept line is APPENDED whenever the reconcile surfaced keepers
 * (plan A1). preexistingTinyCount is deliberately NOT surfaced here — orphan
 * tiny text is Diagnostics material; Hide never touches it (edge row 8).
 */
export function hideReceipt(r: GHideResult): string {
  if (r.paragraphsChanged === 0 && r.newlyKeptRestored === 0) {
    return STRINGS.receipts.hideNoop;
  }
  const parts: string[] = [];
  if (r.paragraphsChanged === 0) {
    parts.push(STRINGS.receipts.hideNothingNew);
  } else if (r.regionsAlreadyHidden > 0) {
    // Re-hide: the parenthetical counts already-hidden PASSAGES (regions in
    // engine terms) because the result has no already-hidden paragraph count.
    parts.push(
      `Hid ${counted(r.paragraphsChanged, "new paragraph")} (${counted(
        r.regionsAlreadyHidden,
        "passage"
      )} already hidden).`
    );
  } else {
    parts.push(`Hid text in ${fmt(r.paragraphsChanged)} of ${counted(r.paragraphsScanned, "paragraph")}.`);
  }
  if (r.newlyKeptRestored > 0) {
    parts.push(`Brought ${counted(r.newlyKeptRestored, "newly kept passage")} back to full size.`);
  }
  return parts.join(" ");
}

/**
 * Show All receipt. "Normalized" merges the two reset-to-style buckets —
 * RLE-mismatched segments and swept orphans — because to the user both mean
 * the same thing: that passage came back at its style's normal size (plan D3
 * counts both in the amber receipt). Every variant ends on the healthy state
 * ("All your text is visible.") and NEVER calls user formatting "lost"
 * (plan A14 rewording) — EXCEPT when a newer engine version's records were
 * skipped (edge row 16): hidden text then genuinely remains, so an amber
 * warning line is appended LAST and the receipt ends on it instead of a
 * healthy-state claim that would not be the whole truth.
 */
export function showAllReceipt(r: GShowAllResult): string {
  const exact = r.segmentsRestoredExact;
  const normalized = r.segmentsNormalized + r.sweptOrphans;
  const skipped = r.rangesSkippedNewerVersion;
  if (exact === 0 && normalized === 0 && skipped === 0) return STRINGS.receipts.showAllNoop;
  const parts: string[] = [];
  if (exact === 0 && normalized === 0) {
    // Only newer-version state was found: the no-op line's "no hidden text
    // found" would contradict the warning below, so the warning stands alone.
  } else if (normalized === 0) {
    parts.push(
      `Restored ${counted(exact, "passage")} to ${their(exact)} saved ${sizesWord(exact)}. All your text is visible.`
    );
  } else if (exact === 0) {
    // Pure normalize (e.g. a consented sweep on a copied doc): nothing had a
    // usable record, so "Restored … to their saved sizes" would be untrue.
    parts.push(
      `Reset ${counted(normalized, "passage")} to ${their(normalized)} style's normal size. All your text is visible.`
    );
  } else {
    parts.push(
      `Restored ${counted(exact, "passage")} to ${their(exact)} saved ${sizesWord(exact)}; ` +
        `${fmt(normalized)} ${wereWord(normalized)} reset to ${their(normalized)} style's normal size. ` +
        `All your text is visible.`
    );
  }
  if (skipped > 0) {
    // Edge row 16: records this Rostrum cannot read are not ours to touch —
    // truthful, blame-free, and actionable (the newer Rostrum can undo them).
    parts.push(
      `${counted(skipped, "hidden section")} ${wereWord(skipped)} made by a newer Rostrum and ${wereWord(skipped)} left untouched.`
    );
  }
  return parts.join(" ");
}

/**
 * Apply-styles receipt: counted per style (Pockets/Hats/Blocks/Tags, the
 * debate nouns), then the optional repair lines, then the teaching tail.
 * The all-zero outcome is the teaching string, not a failure (Step 7). The
 * degraded line is appended on EVERY branch when updateNamedStyle was
 * rejected — including the zero case, where the teaching string's Ctrl+Alt+1
 * promise needs the caveat most (plan A5/D13).
 */
export function stylesReceipt(r: GStylesResult): string {
  const { pocket, hat, block, tag } = r.restyled;
  const nothingDone = pocket + hat + block + tag === 0 && r.citesRepaired === 0 && r.spacingCleared === 0;
  const parts: string[] = [];
  if (nothingDone) {
    parts.push(STRINGS.receipts.stylesZero);
  } else {
    parts.push(
      `Styles applied — restyled ${counted(pocket, "Pocket")}, ${counted(hat, "Hat")}, ` +
        `${counted(block, "Block")}, ${counted(tag, "Tag")}.`
    );
    if (r.citesRepaired > 0) parts.push(`Repaired ${counted(r.citesRepaired, "cite")}.`);
    if (r.spacingCleared > 0) parts.push(`Cleared extra spacing on ${counted(r.spacingCleared, "paragraph")}.`);
    parts.push(STRINGS.receipts.stylesNext);
  }
  if (!r.namedStylesApplied) parts.push(STRINGS.receipts.stylesDegraded);
  return parts.join(" ");
}

/**
 * Pure-sweep consent prompt (plan A14): asked ONLY when zero rstm ranges
 * exist, so the tiny text might be the user's own formatting — Rostrum asks
 * before normalizing what it didn't hide. count > 0 by contract (the
 * controller never asks about nothing); 0 still renders grammatically as a
 * defensive plural.
 */
export function consentPrompt(count: number): string {
  const subject = count === 1 ? "passage has" : "passages have";
  const object = count === 1 ? "it" : "them";
  return `${fmt(count)} tiny-text ${subject} no Rostrum record. Restore ${object} to normal size too?`;
}

/**
 * First-run styles confirm body (frontendDraft Step 5): counted so the user
 * knows the blast radius before agreeing; the adapter skips the dialog
 * entirely at 0 headings (zero-friction first run).
 */
export function stylesConfirm(headingCount: number): string {
  return `This restyles ${counted(headingCount, "existing heading")} in this doc. Undo with Ctrl+Z.`;
}

/** Sidebar header state line (frontendDraft Step 4): counts hidden PASSAGES
 * (one per rstm range — the fields-masked state read, plan A13). */
export function docStateLine(hiddenPassages: number): string {
  if (hiddenPassages === 0) return STRINGS.sidebar.stateNothingHidden;
  return `This doc: ${counted(hiddenPassages, "hidden passage")}.`;
}

/** Mark-cite receipt (frontendDraft Step 7 — "Marked 1 cite."). Zero means
 * the selection held no markable text: teach, don't scold. */
export function markCiteReceipt(count: number): string {
  if (count === 0) return STRINGS.receipts.markCiteNoop;
  return `Marked ${counted(count, "cite")}.`;
}

/**
 * Analytic-ify receipt (Loop 003 — spec §3 strings.ts).
 *
 * Uses the NOUN form ("Made N paragraphs Analytics") rather than a verb form
 * ("Analytic-ified N paragraphs") because "analytic-ified" is awkward
 * grammar, and the noun form naturally teaches the unit: a whole paragraph is
 * the granularity.  "Analytics" is the style's proper name (the user's chosen
 * word, capitalized like Hide / Show All); the parenthetical keeps the color
 * hint so the user knows what changed without opening the format sidebar.
 *
 * Zero (empty ordinals, cursor not on a paragraph) → teaching noop string.
 */
export function analyticifyReceipt(n: number): string {
  if (n === 0) return STRINGS.receipts.analyticifyNoop;
  return `Made ${counted(n, "paragraph")} Analytics (navy, 14pt).`;
}

/**
 * Delete analytics receipt (Loop 003 — spec §3 strings.ts).
 *
 * Counts PARAGRAPHS affected (primary) and ranges deleted (secondary detail)
 * so the user understands both scope (which paragraphs) and granularity
 * (how many separate text spans were removed).  Zero-affected case is the
 * noop string (TOCTOU: analytics vanished between count-read and verb-read).
 *
 * "Deleted" (not "removed" or "cleared") is the accurate verb: this is the
 * one destructive action in the engine and copy should be unambiguous.
 */
export function deleteAnalyticsReceipt(r: DeleteAnalyticsResult): string {
  if (r.paragraphsAffected === 0) return STRINGS.receipts.deleteAnalyticsNoop;
  // runsDeleted counts the actual range ranges removed (partial + whole
  // paragraphs), giving users something to cross-check against the doc.
  return (
    `Deleted Analytics text in ${counted(r.paragraphsAffected, "paragraph")}` +
    ` (${counted(r.runsDeleted, "range")} removed).`
  );
}

/**
 * Delete analytics confirm body (Loop 003 — spec §3 strings.ts).
 *
 * CALM + TRUTHFUL voice (spec requirement): does NOT use "permanent" or
 * "irreversible" — those words create anxiety disproportionate to the action.
 * Names the REAL recovery path (003-F10 wet finding): on the live host a
 * server-side delete is NOT in the client Ctrl+Z undo stack, so the confirm
 * must NOT promise Ctrl+Z — it points to File > Version history (the recovery
 * that actually works) and states that Show All does not, disambiguating the
 * normal hide/show loop users are trained to reach for. One sentence of scope +
 * one of recovery, matching the stylesConfirm voice.
 *
 * n > 0 by contract (the adapter shows this confirm only when count > 0).
 */
export function deleteAnalyticsConfirm(n: number): string {
  return (
    `This removes the Analytics text in ${counted(n, "paragraph")}.` +
    ` To recover it, use File > Version history — Show All will not bring it back.`
  );
}

// ---------------------------------------------------------------------------
// Failure mapping — every error class in types.ts → exactly one deck entry.
// ---------------------------------------------------------------------------

/** What the refusal dialog / sidebar banner renders for a failed verb. */
export interface UserErrorMessage {
  title: string;
  body: string;
  /** "amber" is reserved for healthy-doc outcomes (all text visible). NO
   * error class qualifies — every refusal below is red — but the union is the
   * shared severity vocabulary the receipt banner also uses, so the adapter
   * renders both channels from one scale (frontendDraft Step 2). */
  severity: "red" | "amber";
}

/** Local shorthand: every mapped refusal is red (see UserErrorMessage). */
function red(entry: { title: string; body: string }): UserErrorMessage {
  return { title: entry.title, body: entry.body, severity: "red" };
}

/**
 * Map ANY thrown value to truthful user copy. instanceof checks (not name
 * strings) because the engine and adapter share these exact classes; anything
 * unrecognized — including non-Error throws — falls through to the unknown
 * entry, which promises only what the engine guarantees (text is never
 * deleted; Show All always recovers).
 */
export function errorMessage(e: unknown): UserErrorMessage {
  const E = STRINGS.errors;
  if (e instanceof SuggestionsActiveError) return red(E.suggestions);
  if (e instanceof MultiTabError) return red(E.multiTab);
  if (e instanceof HiddenStateError) return red(E.hiddenState);
  // Per-verb maps: PartialApplyError before the generic mismatch so an
  // interrupted multi-chunk apply never reads as "nothing was applied".
  if (e instanceof PartialApplyError) return red(E.partialApply[e.verb]);
  if (e instanceof RevisionConflictError) return red(E.revisionConflict[e.verb]);
  if (e instanceof RevisionMismatchError) return red(E.revisionMismatch);
  if (e instanceof DocsApiError) return red(E.docsApi);
  return red(E.unknown);
}
