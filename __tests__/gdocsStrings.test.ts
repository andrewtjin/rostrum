// gdocs copy-deck suite (plan A11.v / S10b). strings.ts is the ONE module user
// copy flows from, so this suite is deliberately MECHANICAL where the rules are
// mechanical: it walks every STRINGS leaf AND the output of every format helper
// across a representative input grid, then audits the whole corpus at once for
// the banned engine lexicon, exclamation points, and emoji. Pattern/grammar
// tests pin the counted receipt shapes exactly (the strings are a product
// surface — drift IS a regression), and every error class in types.ts must map
// to a distinct, truthful refusal. No snapshots: a snapshot would bless a
// lexicon violation as readily as a fix.

import {
  STRINGS,
  hideReceipt,
  showAllReceipt,
  stylesReceipt,
  consentPrompt,
  stylesConfirm,
  docStateLine,
  markCiteReceipt,
  analyticifyReceipt,
  deleteAnalyticsReceipt,
  deleteAnalyticsConfirm,
  errorMessage,
  UserErrorMessage
} from "../google-docs/src/core/strings";
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
} from "../google-docs/src/core/types";
import { GDOCS_VERSION } from "../google-docs/src/core/constants";
// helpHtml renders STRINGS.help.* — imported so the rendered-surface assertion
// proves the analytics-verbs explanation actually reaches the Help dialog,
// not just the deck (exec-review MAJOR: the explanation must hit a surface).
import { helpHtml } from "../google-docs/src/adapter/sidebarHtml";

// ---------------------------------------------------------------------------
// Result builders — zeroed defaults so each case states only what it tests.
// ---------------------------------------------------------------------------

function hide(over: Partial<GHideResult> = {}): GHideResult {
  return {
    paragraphsScanned: 0,
    paragraphsChanged: 0,
    regionsHidden: 0,
    regionsAlreadyHidden: 0,
    newlyKeptRestored: 0,
    preexistingTinyCount: 0,
    ...over
  };
}

function show(over: Partial<GShowAllResult> = {}): GShowAllResult {
  return {
    segmentsRestoredExact: 0,
    segmentsNormalized: 0,
    sweptOrphans: 0,
    rangesDeleted: 0,
    rangesSkippedNewerVersion: 0,
    ...over
  };
}

function styles(
  restyled: Partial<GStylesResult["restyled"]> = {},
  over: Partial<Omit<GStylesResult, "restyled">> = {}
): GStylesResult {
  return {
    namedStylesApplied: true,
    restyled: { pocket: 0, hat: 0, block: 0, tag: 0, ...restyled },
    spacingCleared: 0,
    citesRepaired: 0,
    ...over
  };
}

// ---------------------------------------------------------------------------
// Representative input grids — chosen to drive EVERY branch of every helper
// (zero / singular / plural / thousands / combined), so the lexicon audit sees
// every sentence the helpers can compose, not just the static deck.
// ---------------------------------------------------------------------------

const HIDE_GRID: GHideResult[] = [
  hide(), // pure no-op
  hide({ paragraphsScanned: 41, paragraphsChanged: 12, regionsHidden: 12 }), // fresh hide
  hide({ paragraphsScanned: 1, paragraphsChanged: 1, regionsHidden: 1 }), // fresh, singular
  hide({ paragraphsScanned: 4100, paragraphsChanged: 2431, regionsHidden: 900 }), // thousands
  hide({ paragraphsScanned: 50, paragraphsChanged: 12, regionsHidden: 12, regionsAlreadyHidden: 5 }), // re-hide
  hide({ paragraphsScanned: 50, paragraphsChanged: 1, regionsHidden: 1, regionsAlreadyHidden: 1 }), // re-hide, singular
  hide({ paragraphsScanned: 50, newlyKeptRestored: 3 }), // reconcile only surfaced keepers
  hide({ paragraphsScanned: 50, newlyKeptRestored: 1 }), // surfaced, singular
  hide({ paragraphsScanned: 50, paragraphsChanged: 2, regionsHidden: 2, regionsAlreadyHidden: 4, newlyKeptRestored: 1 }), // everything at once
  hide({ preexistingTinyCount: 7 }) // orphan tiny text present (must NOT surface)
];

const SHOW_GRID: GShowAllResult[] = [
  show(), // no-op
  show({ segmentsRestoredExact: 257, rangesDeleted: 41 }), // exact only
  show({ segmentsRestoredExact: 1, rangesDeleted: 1 }), // exact, singular
  show({ segmentsRestoredExact: 2431, rangesDeleted: 900 }), // thousands
  show({ segmentsRestoredExact: 257, segmentsNormalized: 80, sweptOrphans: 6, rangesDeleted: 41 }), // amber mixed
  show({ segmentsRestoredExact: 2, segmentsNormalized: 1, rangesDeleted: 2 }), // amber, singular normalized
  show({ segmentsNormalized: 3 }), // pure normalize (copied doc)
  show({ sweptOrphans: 1 }), // pure sweep, singular
  show({ segmentsRestoredExact: 2, rangesDeleted: 2, rangesSkippedNewerVersion: 3 }), // newer-version skip, appended (row 16)
  show({ rangesSkippedNewerVersion: 1 }) // ONLY newer-version state found, singular
];

const STYLES_GRID: GStylesResult[] = [
  styles(), // zero teaching case
  styles({}, { namedStylesApplied: false }), // zero case, degraded
  styles({ pocket: 12, hat: 18, block: 40, tag: 96 }), // the Step-7 journey counts
  styles({ pocket: 1, hat: 1, block: 1, tag: 1 }), // singulars
  styles({ pocket: 12, hat: 18, block: 40, tag: 96 }, { namedStylesApplied: false }), // counted + degraded
  styles({ pocket: 2 }, { citesRepaired: 3, spacingCleared: 20 }), // repair + spacing lines
  styles({}, { citesRepaired: 2 }), // repair-only pass (not the teaching case)
  styles({ pocket: 1 }, { citesRepaired: 1, spacingCleared: 1 }) // singular repair lines
];

/** Every error class in types.ts, every per-verb variant, plus unknown throws.
 * Analytics verb rows (Loop 003) are appended after the existing rows so the
 * "distinct body" audit catches them as well — each must have a unique body. */
const ERROR_GRID: { label: string; e: unknown }[] = [
  { label: "suggestions", e: new SuggestionsActiveError() },
  { label: "multiTab", e: new MultiTabError(3) },
  { label: "hiddenState", e: new HiddenStateError() },
  { label: "revisionMismatch", e: new RevisionMismatchError() },
  { label: "revisionConflict.hide", e: new RevisionConflictError("hide") },
  { label: "revisionConflict.showAll", e: new RevisionConflictError("showAll") },
  { label: "revisionConflict.applyStyles", e: new RevisionConflictError("applyStyles") },
  // Analytics revision-conflict rows (Loop 003): both land = "nothing changed",
  // but the phrasing differs so the distinct-body audit catches any future merge.
  { label: "revisionConflict.analyticify", e: new RevisionConflictError("analyticify") },
  { label: "revisionConflict.deleteAnalytics", e: new RevisionConflictError("deleteAnalytics") },
  { label: "partialApply.hide", e: new PartialApplyError("hide", 1, 3) },
  { label: "partialApply.showAll", e: new PartialApplyError("showAll", 2, 5) },
  { label: "partialApply.applyStyles", e: new PartialApplyError("applyStyles", 1, 2) },
  // Analytics partial-apply rows (Loop 003): recovery routes differ critically —
  // analyticify is safe to repeat; deleteAnalytics partial is irreversible via
  // Show All (requires "Delete analytics again").  Bodies must be distinct.
  { label: "partialApply.analyticify", e: new PartialApplyError("analyticify", 1, 3) },
  { label: "partialApply.deleteAnalytics", e: new PartialApplyError("deleteAnalytics", 2, 4) },
  { label: "docsApi", e: new DocsApiError("bad field mask") },
  { label: "unknown.error", e: new Error("boom") },
  { label: "unknown.string", e: "boom" },
  { label: "unknown.null", e: null }
];

// ---------------------------------------------------------------------------
// Corpus assembly — every string the user can ever see, labeled by origin so
// audit failures point at the offending leaf/branch directly.
// ---------------------------------------------------------------------------

interface CorpusEntry {
  at: string;
  text: string;
}

/** Recursively collect every string leaf of the STRINGS deck. */
function collectLeaves(node: unknown, path: string, out: CorpusEntry[]): void {
  if (typeof node === "string") {
    out.push({ at: path, text: node });
    return;
  }
  if (node !== null && typeof node === "object") {
    for (const [key, value] of Object.entries(node)) {
      collectLeaves(value, path === "" ? key : `${path}.${key}`, out);
    }
  }
}

/** The full audited corpus: deck leaves + every helper output over the grids. */
function buildCorpus(): CorpusEntry[] {
  const corpus: CorpusEntry[] = [];
  collectLeaves(STRINGS, "STRINGS", corpus);
  HIDE_GRID.forEach((r, i) => corpus.push({ at: `hideReceipt[${i}]`, text: hideReceipt(r) }));
  SHOW_GRID.forEach((r, i) => corpus.push({ at: `showAllReceipt[${i}]`, text: showAllReceipt(r) }));
  STYLES_GRID.forEach((r, i) => corpus.push({ at: `stylesReceipt[${i}]`, text: stylesReceipt(r) }));
  for (const n of [0, 1, 2, 86, 1200]) {
    corpus.push({ at: `consentPrompt(${n})`, text: consentPrompt(n) });
    corpus.push({ at: `stylesConfirm(${n})`, text: stylesConfirm(n) });
    corpus.push({ at: `docStateLine(${n})`, text: docStateLine(n) });
    corpus.push({ at: `markCiteReceipt(${n})`, text: markCiteReceipt(n) });
  }
  // Analytics helpers (Loop 003): drive every branch so the lexicon audit sees
  // the noun-form receipt, the noop string, and every branch of the confirm.
  // n=0 exercises the noop path; n=1 singular; n=3/86 plural; n=1200 thousands.
  for (const n of [0, 1, 3, 86, 1200]) {
    corpus.push({ at: `analyticifyReceipt(${n})`, text: analyticifyReceipt(n) });
    corpus.push({ at: `deleteAnalyticsConfirm(${n})`, text: deleteAnalyticsConfirm(n) });
  }
  // deleteAnalyticsReceipt: zero-affected noop + singular + plural + thousands.
  for (const r of [
    { paragraphsAffected: 0, runsDeleted: 0 },
    { paragraphsAffected: 1, runsDeleted: 1 },
    { paragraphsAffected: 3, runsDeleted: 5 },
    { paragraphsAffected: 86, runsDeleted: 120 },
    { paragraphsAffected: 1200, runsDeleted: 2400 }
  ]) {
    corpus.push({
      at: `deleteAnalyticsReceipt(${r.paragraphsAffected},${r.runsDeleted})`,
      text: deleteAnalyticsReceipt(r)
    });
  }
  for (const { label, e } of ERROR_GRID) {
    const m = errorMessage(e);
    corpus.push({ at: `errorMessage(${label}).title`, text: m.title });
    corpus.push({ at: `errorMessage(${label}).body`, text: m.body });
  }
  return corpus;
}

const CORPUS = buildCorpus();

/** Lower-cased word tokens of a string ("batchUpdate" -> ["batchupdate"]),
 * so the banned scan is whole-token and case-insensitive by construction. */
function tokens(text: string): string[] {
  return text.toLowerCase().match(/[a-z0-9]+/g) ?? [];
}

// ---------------------------------------------------------------------------
// The mechanical lexicon audit (plan A11.v).
// ---------------------------------------------------------------------------

describe("lexicon audit — every deck leaf and every helper output", () => {
  it("collected a non-trivial corpus (the walk itself must not silently break)", () => {
    // Guard the auditor: if a refactor made collectLeaves miss the deck, every
    // audit below would pass vacuously. ~60 deck leaves + ~60 helper outputs.
    expect(CORPUS.length).toBeGreaterThan(100);
  });

  it("never uses a banned engine token (whole word, any case)", () => {
    // The engine nouns the user must never see (plan A11.v + frontendDraft
    // Step 1). "run/runs" also bans the VERB — copy says "use Show All again".
    const banned = new Set(["run", "runs", "namedrange", "manifest", "sentinel", "batchupdate", "card", "cards"]);
    const offenders = CORPUS.filter((e) => tokens(e.text).some((t) => banned.has(t)));
    expect(offenders).toEqual([]);
  });

  it("never uses the banned phrases 'named style', 'named range', or 'lost' framing", () => {
    // Two-word engine phrases the token scan can't catch, plus plan A14's
    // explicit rewording rule: user formatting is never called "lost".
    const offenders = CORPUS.filter((e) => /named style|named range|\blost\b/i.test(e.text));
    expect(offenders).toEqual([]);
  });

  it("never uses exclamation points or emoji (voice rule, frontendDraft Step 2)", () => {
    const offenders = CORPUS.filter((e) => e.text.includes("!") || /\p{Extended_Pictographic}/u.test(e.text));
    expect(offenders).toEqual([]);
  });

  it("has no empty or whitespace-padded leaves (a blank label is a rendering bug)", () => {
    const offenders = CORPUS.filter((e) => e.text.trim() === "" || e.text !== e.text.trim());
    expect(offenders).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// Deck completeness — the surfaces named by the plan must exist verbatim.
// ---------------------------------------------------------------------------

describe("deck completeness (plan D11 / frontendDraft Steps 3-5)", () => {
  it("pins the nine menu labels exactly (these are the product's front door)", () => {
    // Two analytics labels added in Loop 003 (spec §3 strings.ts):
    //   analyticify — no ellipsis (non-destructive, acts immediately);
    //   deleteAnalytics — ellipsis signals a confirm dialog precedes the action.
    // Key order matches the STRINGS.menu declaration so diffs are readable.
    expect(STRINGS.menu).toEqual({
      hide: "Hide",
      showAll: "Show All",
      applyStyles: "Apply debate styles",
      markCite: "Mark cite",
      analyticify: "Analytic-ify",
      deleteAnalytics: "Delete analytics…",
      openPanel: "Open Rostrum panel…",
      helpShortcuts: "Help & shortcuts",
      diagnostics: "Diagnostics"
    });
  });

  it("carries the no-add-on escape hatch with the exact recovery recipe (plan D15)", () => {
    expect(STRINGS.help.escapeHatch).toContain("Select All");
    expect(STRINGS.help.escapeHatch).toContain("font size to 11");
  });

  it("carries the scoped content-deletion honesty in Help that the F5 fallback gave up (exec-review MAJOR)", () => {
    // The universal F5/unknown-error fallback was deliberately stripped of any
    // "only Delete analytics removes content" promise (it must hold for every
    // verb). That scoped truth has to live SOMEWHERE standing — Help is it.
    // These assertions target the MEANING, not the exact prose, so reworded
    // copy still passes as long as the safety facts survive:
    const line = STRINGS.help.analyticsVerbs;
    // (a) it is about the analytics verbs at all:
    expect(line).toContain("Analytic-ify");
    expect(line).toContain("Delete analytics");
    // (b) it scopes content deletion: Delete analytics is the ONLY content
    //     remover, and only the analytics text the user styled:
    expect(line).toMatch(/only .*removes content/i);
    // (c) it disambiguates the hide/show loop: Ctrl+Z undoes the delete, but
    //     Show All (which only un-shrinks) does NOT bring deleted content back.
    expect(line).toContain("Ctrl+Z");
    expect(line).toContain("Show All does not");
  });

  it("renders the analytics-verbs explanation in the rendered Help dialog, not just the deck", () => {
    // The deck leaf existing is not enough — the exec-review MAJOR was that the
    // explanation never reached a SURFACE. helpHtml() must actually emit the
    // line so a future refactor that drops the <p> is caught here. The line has
    // no HTML-special characters, so it survives escapeHtml() verbatim; assert
    // on the two safety clauses (not the whole string) so a reword still passes
    // as long as the scoped-deletion fact and the Show All caveat reach the UI.
    const html = helpHtml();
    expect(html).toContain("only the analytics text you styled");
    expect(html).toContain("Show All does not");
  });

  it("states the consumed-first-click rule once, shared by help and the sidebar (plan A15)", () => {
    expect(STRINGS.help.firstClick).toBe(
      "First time? Google uses your first click to ask permission — just click again."
    );
    // One constant, two surfaces — referential equality is the anti-drift proof.
    expect(STRINGS.sidebar.emptyState).toBe(STRINGS.help.firstClick);
  });

  it("ships Windows AND Mac chords for all five cheat-sheet rows (plan A15)", () => {
    const rows = Object.values(STRINGS.sidebar.cheatSheet.rows);
    expect(rows).toHaveLength(5);
    for (const row of rows) {
      expect(row.windows).toMatch(/^Ctrl\+Alt\+\d$/);
      expect(row.mac).toMatch(/^Cmd\+Option\+\d$/);
      expect(row.name.length).toBeGreaterThan(0);
      expect(row.verbatim.length).toBeGreaterThan(0);
    }
    // The debate-noun order is the teaching order (Pocket -> Normal).
    expect(rows.map((r) => r.name)).toEqual(["Pocket", "Hat", "Block", "Tag", "Normal"]);
  });

  it("surfaces the gdocs artifact version in the footer (plan D14)", () => {
    expect(STRINGS.sidebar.footer.version).toContain(GDOCS_VERSION);
  });
});

// ---------------------------------------------------------------------------
// Hide receipts — counted patterns and zero/singular/plural grammar.
// ---------------------------------------------------------------------------

describe("hideReceipt", () => {
  it("formats a fresh hide as 'Hid text in N of M paragraphs.' (Word's unit)", () => {
    expect(hideReceipt(hide({ paragraphsScanned: 41, paragraphsChanged: 12, regionsHidden: 12 }))).toBe(
      "Hid text in 12 of 41 paragraphs."
    );
  });

  it("groups thousands in counts (locale-independent)", () => {
    expect(hideReceipt(hide({ paragraphsScanned: 4100, paragraphsChanged: 2431, regionsHidden: 900 }))).toBe(
      "Hid text in 2,431 of 4,100 paragraphs."
    );
  });

  it("keeps the singular grammatical on a one-paragraph doc", () => {
    expect(hideReceipt(hide({ paragraphsScanned: 1, paragraphsChanged: 1, regionsHidden: 1 }))).toBe(
      "Hid text in 1 of 1 paragraph."
    );
  });

  it("formats a re-hide with the already-hidden parenthetical", () => {
    expect(
      hideReceipt(hide({ paragraphsScanned: 50, paragraphsChanged: 12, regionsHidden: 12, regionsAlreadyHidden: 5 }))
    ).toBe("Hid 12 new paragraphs (5 passages already hidden).");
  });

  it("keeps the re-hide singulars grammatical", () => {
    expect(
      hideReceipt(hide({ paragraphsScanned: 50, paragraphsChanged: 1, regionsHidden: 1, regionsAlreadyHidden: 1 }))
    ).toBe("Hid 1 new paragraph (1 passage already hidden).");
  });

  it("returns the Word idempotent no-op pattern verbatim when nothing happened", () => {
    expect(hideReceipt(hide({ paragraphsScanned: 33 }))).toBe(
      "Nothing to hide here (already hidden, or all kept)."
    );
    expect(hideReceipt(hide({ paragraphsScanned: 33 }))).toBe(STRINGS.receipts.hideNoop);
  });

  it("reports a reconcile that ONLY surfaced keepers distinctly from the no-op (plan A1)", () => {
    // The doc DID change (hidden text came back to full size) — claiming
    // "nothing to hide" alone would contradict what the user just watched.
    expect(hideReceipt(hide({ paragraphsScanned: 50, newlyKeptRestored: 3 }))).toBe(
      "Nothing new to hide here. Brought 3 newly kept passages back to full size."
    );
    expect(hideReceipt(hide({ paragraphsScanned: 50, newlyKeptRestored: 1 }))).toBe(
      "Nothing new to hide here. Brought 1 newly kept passage back to full size."
    );
  });

  it("appends the newly-kept line to a counted re-hide (reconcile + hide in one pass)", () => {
    expect(
      hideReceipt(
        hide({ paragraphsScanned: 50, paragraphsChanged: 2, regionsHidden: 2, regionsAlreadyHidden: 4, newlyKeptRestored: 1 })
      )
    ).toBe("Hid 2 new paragraphs (4 passages already hidden). Brought 1 newly kept passage back to full size.");
  });

  it("never surfaces preexistingTinyCount (orphan tiny text is Diagnostics material, edge row 8)", () => {
    const base = { paragraphsScanned: 41, paragraphsChanged: 12, regionsHidden: 12 };
    expect(hideReceipt(hide({ ...base, preexistingTinyCount: 7 }))).toBe(
      hideReceipt(hide({ ...base, preexistingTinyCount: 0 }))
    );
    // Including on the no-op path — tiny text Hide didn't touch is not a result.
    expect(hideReceipt(hide({ preexistingTinyCount: 7 }))).toBe(STRINGS.receipts.hideNoop);
  });
});

// ---------------------------------------------------------------------------
// Show All receipts — exact, amber, and the healthy-state guarantee.
// ---------------------------------------------------------------------------

describe("showAllReceipt", () => {
  it("formats the all-exact restore", () => {
    expect(showAllReceipt(show({ segmentsRestoredExact: 257, rangesDeleted: 41 }))).toBe(
      "Restored 257 passages to their saved sizes. All your text is visible."
    );
  });

  it("keeps the exact-restore singular grammatical ('its saved size')", () => {
    expect(showAllReceipt(show({ segmentsRestoredExact: 1, rangesDeleted: 1 }))).toBe(
      "Restored 1 passage to its saved size. All your text is visible."
    );
  });

  it("formats the amber mixed restore with the plan's exact wording", () => {
    // 80 RLE-mismatched segments + 6 swept orphans merge into ONE normalized
    // count (86) — to the user both mean "came back at the style's normal
    // size" (plan D3); splitting them would leak engine mechanics.
    expect(
      showAllReceipt(show({ segmentsRestoredExact: 257, segmentsNormalized: 80, sweptOrphans: 6, rangesDeleted: 41 }))
    ).toBe(
      "Restored 257 passages to their saved sizes; 86 were reset to their style's normal size. All your text is visible."
    );
  });

  it("keeps the amber singular grammatical ('1 was reset to its...')", () => {
    expect(showAllReceipt(show({ segmentsRestoredExact: 2, segmentsNormalized: 1, rangesDeleted: 2 }))).toBe(
      "Restored 2 passages to their saved sizes; 1 was reset to its style's normal size. All your text is visible."
    );
  });

  it("reports a pure normalize truthfully (no saved sizes existed to restore)", () => {
    // E.g. a consented sweep on a copied doc — "Restored ... to their saved
    // sizes" would be untrue, so the verb changes to "Reset".
    expect(showAllReceipt(show({ segmentsNormalized: 3 }))).toBe(
      "Reset 3 passages to their style's normal size. All your text is visible."
    );
    expect(showAllReceipt(show({ sweptOrphans: 1 }))).toBe(
      "Reset 1 passage to its style's normal size. All your text is visible."
    );
  });

  it("returns the no-op pattern when there was nothing to show", () => {
    expect(showAllReceipt(show())).toBe("Nothing to show here (no hidden text found).");
  });

  it("appends the newer-version amber line after the restore counts (edge row 16)", () => {
    expect(showAllReceipt(show({ segmentsRestoredExact: 2, rangesDeleted: 2, rangesSkippedNewerVersion: 3 }))).toBe(
      "Restored 2 passages to their saved sizes. All your text is visible. " +
        "3 hidden sections were made by a newer Rostrum and were left untouched."
    );
  });

  it("keeps the newer-version singular grammatical and never claims the no-op when only skips happened", () => {
    // Only newer-version state was found: "no hidden text found" would be a
    // lie, so the warning line stands alone as the whole receipt.
    expect(showAllReceipt(show({ rangesSkippedNewerVersion: 1 }))).toBe(
      "1 hidden section was made by a newer Rostrum and was left untouched."
    );
  });

  it("ends on the healthy state for every non-noop variant — EXCEPT when newer-version skips remain (rows 16/A14)", () => {
    for (const r of SHOW_GRID) {
      const text = showAllReceipt(r);
      if (text === STRINGS.receipts.showAllNoop) continue;
      if (r.rangesSkippedNewerVersion > 0) {
        // Hidden text genuinely remains — the receipt must end on the warning,
        // never on a healthy-state claim that is not the whole truth.
        expect(text.endsWith("left untouched.")).toBe(true);
        expect(text.endsWith("All your text is visible.")).toBe(false);
      } else {
        expect(text.endsWith("All your text is visible.")).toBe(true);
      }
    }
  });
});

// ---------------------------------------------------------------------------
// Apply-styles receipts — per-style counts, teaching zero, degraded line.
// ---------------------------------------------------------------------------

describe("stylesReceipt", () => {
  it("counts each debate style by its noun (the Step-7 journey numbers)", () => {
    expect(stylesReceipt(styles({ pocket: 12, hat: 18, block: 40, tag: 96 }))).toBe(
      "Styles applied — restyled 12 Pockets, 18 Hats, 40 Blocks, 96 Tags. " +
        "Now style lines with Ctrl+Alt+1–4 (see the cheat sheet in the Rostrum panel)."
    );
  });

  it("keeps the per-style singulars grammatical", () => {
    expect(stylesReceipt(styles({ pocket: 1, hat: 1, block: 1, tag: 1 }))).toContain(
      "restyled 1 Pocket, 1 Hat, 1 Block, 1 Tag."
    );
  });

  it("teaches instead of failing on the zero case (frontendDraft Step 7)", () => {
    expect(stylesReceipt(styles())).toBe(STRINGS.receipts.stylesZero);
    expect(stylesReceipt(styles())).toContain("press Ctrl+Alt+1 to make it a Pocket");
  });

  it("appends the cite-repair and spacing lines only when those passes did work", () => {
    const text = stylesReceipt(styles({ pocket: 2 }, { citesRepaired: 3, spacingCleared: 20 }));
    expect(text).toContain("Repaired 3 cites.");
    expect(text).toContain("Cleared extra spacing on 20 paragraphs.");
    // And not when their counts are zero:
    const bare = stylesReceipt(styles({ pocket: 2 }));
    expect(bare).not.toContain("Repaired");
    expect(bare).not.toContain("spacing");
  });

  it("treats a repair-only pass as work done, not the zero teaching case", () => {
    const text = stylesReceipt(styles({}, { citesRepaired: 2 }));
    expect(text).not.toBe(STRINGS.receipts.stylesZero);
    expect(text).toContain("Repaired 2 cites.");
  });

  it("appends the degraded line whenever updateNamedStyle was rejected (plan A5/D13)", () => {
    const counted = stylesReceipt(styles({ pocket: 12, hat: 18, block: 40, tag: 96 }, { namedStylesApplied: false }));
    expect(counted.endsWith(STRINGS.receipts.stylesDegraded)).toBe(true);
    // The zero case needs the caveat MOST: its teaching string promises
    // Ctrl+Alt+1 works, which is exactly what degraded mode undermines.
    const zero = stylesReceipt(styles({}, { namedStylesApplied: false }));
    expect(zero).toBe(`${STRINGS.receipts.stylesZero} ${STRINGS.receipts.stylesDegraded}`);
  });
});

// ---------------------------------------------------------------------------
// Small helpers — consent, confirm, state line, mark cite.
// ---------------------------------------------------------------------------

describe("consentPrompt (pure-sweep consent, plan A14)", () => {
  it("asks before normalizing unrecorded tiny text, counted", () => {
    expect(consentPrompt(86)).toBe("86 tiny-text passages have no Rostrum record. Restore them to normal size too?");
  });

  it("keeps the singular grammatical ('passage has ... it')", () => {
    expect(consentPrompt(1)).toBe("1 tiny-text passage has no Rostrum record. Restore it to normal size too?");
  });

  it("groups thousands and stays a question", () => {
    const text = consentPrompt(1200);
    expect(text.startsWith("1,200 ")).toBe(true);
    expect(text.endsWith("?")).toBe(true);
  });

  it("renders the defensive zero grammatically (controller never asks, but copy must not break)", () => {
    expect(consentPrompt(0)).toBe("0 tiny-text passages have no Rostrum record. Restore them to normal size too?");
  });
});

describe("stylesConfirm / docStateLine / markCiteReceipt", () => {
  it("states the styles blast radius before applying (frontendDraft Step 5)", () => {
    expect(stylesConfirm(166)).toBe("This restyles 166 existing headings in this doc. Undo with Ctrl+Z.");
    expect(stylesConfirm(1)).toBe("This restyles 1 existing heading in this doc. Undo with Ctrl+Z.");
  });

  it("renders the sidebar state line for armed and clean docs", () => {
    expect(docStateLine(0)).toBe(STRINGS.sidebar.stateNothingHidden);
    expect(docStateLine(343)).toBe("This doc: 343 hidden passages.");
    expect(docStateLine(1)).toBe("This doc: 1 hidden passage.");
  });

  it("counts marked cites and teaches on an empty selection", () => {
    expect(markCiteReceipt(1)).toBe("Marked 1 cite.");
    expect(markCiteReceipt(2)).toBe("Marked 2 cites.");
    expect(markCiteReceipt(0)).toBe(STRINGS.receipts.markCiteNoop);
  });
});

// ---------------------------------------------------------------------------
// Failure mapping — every error class -> distinct, truthful, red copy.
// ---------------------------------------------------------------------------

describe("errorMessage", () => {
  it("maps every error class (and unknown throws) to red — amber is reserved for healthy-doc outcomes", () => {
    // No refusal leaves the doc in a healthy "all visible" state, so by the
    // severity contract (frontendDraft Step 2) every mapping must be red.
    for (const { label, e } of ERROR_GRID) {
      expect({ label, severity: errorMessage(e).severity }).toEqual({ label, severity: "red" });
    }
  });

  it("gives every distinct failure a distinct body (no two errors share copy)", () => {
    const bodies = ERROR_GRID.filter((g) => !g.label.startsWith("unknown")).map((g) => errorMessage(g.e).body);
    expect(new Set(bodies).size).toBe(bodies.length);
  });

  it("routes the suggestions gate to the escape route, Hide voice (plan A7)", () => {
    const m = errorMessage(new SuggestionsActiveError());
    expect(m.body).toContain("accept or reject them (Tools > Review suggested edits)");
    expect(m.body).toContain("Nothing was changed.");
  });

  it("refuses multi-tab docs with a way forward (plan A3)", () => {
    const m = errorMessage(new MultiTabError(3));
    expect(m.body).toContain("tabs");
    expect(m.body).toContain("Nothing was changed.");
  });

  it("tells the styles verb to Show All first while hidden state exists (plan A5)", () => {
    expect(errorMessage(new HiddenStateError()).body).toContain("Show All first, then re-apply styles");
  });

  it("is truthful per verb on revision conflicts — a failed Show All leaves text HIDDEN", () => {
    const hideM = errorMessage(new RevisionConflictError("hide"));
    expect(hideM.body).toContain("nothing was applied");
    expect(hideM.body).toContain("your doc is untouched");
    // The Show All variant must NOT claim the doc is fine — text is still tiny.
    const showM = errorMessage(new RevisionConflictError("showAll"));
    expect(showM.body).toContain("still hidden");
    expect(showM.body).not.toContain("untouched");
    const stylesM = errorMessage(new RevisionConflictError("applyStyles"));
    expect(stylesM.body).toContain("nothing was applied");
  });

  it("gives each interrupted verb its own truthful recovery route (plan A11.iv)", () => {
    // Interrupted Hide: some text IS hidden — Show All is the way back.
    const hideM = errorMessage(new PartialApplyError("hide", 1, 3));
    expect(hideM.body).toContain("some text is hidden");
    expect(hideM.body).toContain("Use Show All to bring everything back");
    // Interrupted Show All: some text already restored — finish, don't undo.
    const showM = errorMessage(new PartialApplyError("showAll", 2, 5));
    expect(showM.body).toContain("some text is already back");
    expect(showM.body).toContain("Use Show All again to finish");
    // Interrupted styles: safe to repeat (idempotent), and says so.
    const stylesM = errorMessage(new PartialApplyError("applyStyles", 1, 2));
    expect(stylesM.body).toContain("safe to repeat");
  });

  it("maps an adapter-level revision mismatch to its own truthful refusal", () => {
    const m = errorMessage(new RevisionMismatchError());
    expect(m.body).toContain("nothing was applied");
  });

  it("maps unmapped Docs rejections to the generic atomic-batch refusal (plan A9)", () => {
    const m = errorMessage(new DocsApiError("bad field mask"));
    expect(m.body).toBe("Docs rejected the change — nothing was applied.");
    // Raw API detail must never leak into user copy.
    expect(m.title + m.body).not.toContain("bad field mask");
  });

  it("falls back to the unknown entry for any unrecognized throw, promising only what the engine guarantees", () => {
    const expected = errorMessage(new Error("boom"));
    expect(errorMessage("boom")).toEqual(expected);
    expect(errorMessage(null)).toEqual(expected);
    expect(errorMessage(42)).toEqual(expected);
    // An unknown failure CANNOT truthfully claim "nothing was applied" — it
    // leans on the hard invariant instead: Show All always recovers font-size
    // damage (the one thing Rostrum guarantees universally across all verbs).
    expect(expected.body).not.toContain("nothing was applied");
    expect(expected.body).toContain("Show All");
    // The body must NOT make the "never deleted" absolute claim: deleteAnalytics
    // IS a content-deleter, so "text is never deleted" would be false after a
    // partial delete whose error lands here.  The universal form says "never
    // SHRINKS your text away" (font-size invariant) instead, which is always true.
    expect(expected.body).not.toContain("never deleted");
    expect(expected.body).not.toContain("Your text is never");
    // No analytics jargon in the universal fallback — it must apply to every verb.
    expect(expected.body).not.toContain("analytics");
    // Raw internal error text never reaches the user.
    expect(expected.title + expected.body).not.toContain("boom");
  });

  it("returns the shared severity vocabulary shape (title/body/severity)", () => {
    const m: UserErrorMessage = errorMessage(new DocsApiError("x"));
    expect(Object.keys(m).sort()).toEqual(["body", "severity", "title"]);
  });

  // -------------------------------------------------------------------------
  // Analytics error entries (Loop 003 — spec §3 strings.ts).
  // Each assertion targets the SAFETY-critical property of its copy, not the
  // exact wording (exact wording is pinned by the distinct-body audit above).
  // -------------------------------------------------------------------------

  it("routes analyticify revision conflict to 'nothing was changed' (idempotent styling, nothing landed)", () => {
    const m = errorMessage(new RevisionConflictError("analyticify"));
    // All retries exhausted: no styling write ever landed, doc is exactly as left.
    expect(m.body).toContain("nothing was changed");
    // Must NOT claim Show All is needed — no font-size damage occurred.
    expect(m.body).not.toContain("Show All");
  });

  it("routes deleteAnalytics revision conflict to 'nothing was changed' (no delete landed)", () => {
    const m = errorMessage(new RevisionConflictError("deleteAnalytics"));
    // All retries exhausted means not a single chunk applied: no content removed.
    expect(m.body).toContain("nothing was changed");
    // Since nothing was deleted, the "Show All will not bring it back" caveat
    // is irrelevant here — it must NOT appear (it belongs only in the partial path).
    expect(m.body).not.toContain("Show All will not bring it back");
  });

  it("routes analyticify partial apply to the idempotent 'safe to repeat' recovery", () => {
    const m = errorMessage(new PartialApplyError("analyticify", 1, 3));
    // Analytic-ify is pure character formatting: identical writes are idempotent,
    // so repeating the verb on already-styled paragraphs is harmless.
    expect(m.body).toContain("safe to repeat");
    // Must NOT direct the user to Show All — no font-size hide was involved.
    expect(m.body).not.toContain("Use Show All");
  });

  it("routes deleteAnalytics partial apply to the destructive partial-delete copy (SAFETY-CRITICAL)", () => {
    const m = errorMessage(new PartialApplyError("deleteAnalytics", 2, 4));
    // Safety-critical: the partial delete CANNOT be recovered via Show All.
    // This must be stated explicitly so users don't assume the normal recovery path.
    expect(m.body).toContain("Show All will not bring it back");
    // Some content was already removed — say so truthfully.
    expect(m.body).toContain("already removed");
    // Recovery: run the verb again to finish, not undo.
    expect(m.body).toContain("Use Delete analytics again");
    // Must NOT claim "safe to repeat" — a second delete of already-deleted text
    // is safe, but framing it as idempotent would obscure the destructive nature.
    expect(m.body).not.toContain("safe to repeat");
  });

  it("analytics error bodies are distinct from each other AND from all existing verb bodies", () => {
    // The four new analytics rows must not share a body with any other row —
    // enforced here explicitly for the safety-critical delete-analytics copy.
    const analyticifyConflict = errorMessage(new RevisionConflictError("analyticify")).body;
    const deleteConflict = errorMessage(new RevisionConflictError("deleteAnalytics")).body;
    const analyticifyPartial = errorMessage(new PartialApplyError("analyticify", 1, 3)).body;
    const deletePartial = errorMessage(new PartialApplyError("deleteAnalytics", 2, 4)).body;
    // All four distinct from each other.
    const four = [analyticifyConflict, deleteConflict, analyticifyPartial, deletePartial];
    expect(new Set(four).size).toBe(4);
    // None duplicates any of the original five verb bodies.
    const origBodies = [
      errorMessage(new RevisionConflictError("hide")).body,
      errorMessage(new RevisionConflictError("showAll")).body,
      errorMessage(new RevisionConflictError("applyStyles")).body,
      errorMessage(new PartialApplyError("hide", 1, 2)).body,
      errorMessage(new PartialApplyError("showAll", 1, 2)).body,
      errorMessage(new PartialApplyError("applyStyles", 1, 2)).body
    ];
    for (const newBody of four) {
      expect(origBodies).not.toContain(newBody);
    }
  });
});

// ---------------------------------------------------------------------------
// Analytics receipt helpers (Loop 003 — spec §3 strings.ts).
// ---------------------------------------------------------------------------

describe("analyticifyReceipt", () => {
  it("names the unit (paragraph) and the applied attributes (navy, 14pt) for clarity", () => {
    expect(analyticifyReceipt(3)).toBe("Made 3 paragraphs analytics (navy, 14pt).");
  });

  it("keeps the singular grammatical", () => {
    expect(analyticifyReceipt(1)).toBe("Made 1 paragraph analytics (navy, 14pt).");
  });

  it("groups thousands", () => {
    expect(analyticifyReceipt(1200)).toBe("Made 1,200 paragraphs analytics (navy, 14pt).");
  });

  it("returns the noop teaching string when n=0 (empty selection / cursor not on a paragraph)", () => {
    // Zero ordinals means the cursor was in a context where no paragraph was
    // selected — teach the right action rather than showing a bare failure.
    expect(analyticifyReceipt(0)).toBe(STRINGS.receipts.analyticifyNoop);
    expect(analyticifyReceipt(0)).toContain("Analytic-ify");
  });
});

describe("deleteAnalyticsReceipt", () => {
  it("counts paragraphs affected (primary) and ranges removed (secondary)", () => {
    expect(deleteAnalyticsReceipt({ paragraphsAffected: 3, runsDeleted: 5 })).toBe(
      "Deleted analytics text in 3 paragraphs (5 ranges removed)."
    );
  });

  it("keeps the paragraph singular grammatical", () => {
    expect(deleteAnalyticsReceipt({ paragraphsAffected: 1, runsDeleted: 1 })).toBe(
      "Deleted analytics text in 1 paragraph (1 range removed)."
    );
  });

  it("keeps the range singular grammatical when paragraph count > 1 but ranges = 1", () => {
    // Degenerate but grammatically correct: e.g. two paragraphs merged into one
    // contiguous range after coalescing.
    expect(deleteAnalyticsReceipt({ paragraphsAffected: 2, runsDeleted: 1 })).toBe(
      "Deleted analytics text in 2 paragraphs (1 range removed)."
    );
  });

  it("groups thousands in both counts", () => {
    expect(deleteAnalyticsReceipt({ paragraphsAffected: 1200, runsDeleted: 2400 })).toBe(
      "Deleted analytics text in 1,200 paragraphs (2,400 ranges removed)."
    );
  });

  it("returns the noop string when paragraphsAffected=0 (TOCTOU: analytics vanished between confirm and verb)", () => {
    // countAnalyticsParagraphs > 0 showed a confirm, but the analytics were
    // deleted (e.g. by a collaborator) before the verb ran.  planDeleteAnalytics
    // returns 0 groups → the controller emits the noop receipt path.
    expect(deleteAnalyticsReceipt({ paragraphsAffected: 0, runsDeleted: 0 })).toBe(
      STRINGS.receipts.deleteAnalyticsNoop
    );
  });
});

describe("deleteAnalyticsConfirm", () => {
  it("states the paragraph count and names both recovery paths (Ctrl+Z and NOT Show All)", () => {
    expect(deleteAnalyticsConfirm(5)).toBe(
      "This removes the analytics text in 5 paragraphs. Undo with Ctrl+Z; Show All will not bring it back."
    );
  });

  it("keeps the singular grammatical", () => {
    expect(deleteAnalyticsConfirm(1)).toBe(
      "This removes the analytics text in 1 paragraph. Undo with Ctrl+Z; Show All will not bring it back."
    );
  });

  it("groups thousands", () => {
    expect(deleteAnalyticsConfirm(1200)).toBe(
      "This removes the analytics text in 1,200 paragraphs. Undo with Ctrl+Z; Show All will not bring it back."
    );
  });

  it("is calm — never uses 'permanent' or 'irreversible' (spec CALM+TRUTHFUL voice requirement)", () => {
    const text = deleteAnalyticsConfirm(10);
    expect(text).not.toContain("permanent");
    expect(text).not.toContain("irreversible");
  });

  it("explicitly disclaims Show All as a recovery path (safety-critical disambiguation)", () => {
    // Users trained on the hide/show loop must see that Show All does NOT undo
    // a content delete — this assertion guards that copy from being edited away.
    expect(deleteAnalyticsConfirm(2)).toContain("Show All will not bring it back");
  });

  it("names Ctrl+Z as the undo path (the real recovery that does work)", () => {
    expect(deleteAnalyticsConfirm(2)).toContain("Ctrl+Z");
  });
});
