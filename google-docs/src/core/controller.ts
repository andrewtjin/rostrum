// The verb orchestrator — the ONE place the pure planners meet the DocsPort
// (plan D5/A5/A7/A13/A14, step S9; Word reference: src/core/invisibility.ts).
//
// Every verb runs the SAME apply protocol (the cross-module pin):
//
//   fetch -> parseDocument -> assertSingleTab -> verb gates -> resolveSettings
//   -> plan -> chunkGroups -> apply chunks chained on the revisionId RETURNED
//   BY EACH applyBatch.
//
// ONE documents.get per verb (plan A13): revision chaining reads writeControl
// from each batchUpdate RESPONSE, never from a fresh get — the only re-fetch
// is the silent re-plan after a first-chunk revision mismatch, where a fresh
// read is the whole point.
//
// Failure discipline (plan D5):
//   * RevisionMismatch BEFORE anything landed (chunk 0 — more precisely,
//     applied === 0, see runVerb) -> silently re-plan from a fresh fetch, up
//     to MAX_REPLAN_ATTEMPTS times, then RevisionConflictError(verb). Retries
//     are IMMEDIATE — no sleeps in core, so the engine stays deterministic
//     and flake-free under test (plan A11.vi).
//   * RevisionMismatch AFTER >= 1 chunk landed -> PartialApplyError(verb,
//     applied, total). Re-planning here would be a lie — the doc already
//     changed under a half-applied verb; the user copy owns the recovery
//     route (Show All unwinds a torn Hide; Show All/styles re-run to finish).
//   * A DocsApiError aborts the verb UNLESS it hit applyStyles' named-style
//     batch — the ONE degradable batch (plan A5/D13): consumer accounts may
//     400 updateNamedStyle, so that batch alone is allowed to fail, flagged
//     namedStylesApplied=false, and the retro pass still runs. The degraded
//     path lives HERE, in core, precisely so it is unit-testable without a
//     host.
//
// Gate placement is deliberate and verb-specific (the A7 asymmetry):
//   * assertSingleTab guards EVERY verb — the parse only carries one tab's
//     body, so no verb may act on a multi-tab doc (plan A3, case 001-F9);
//   * Hide and Apply-styles gate on suggestions (indexes are only
//     trustworthy suggestion-free — they CREATE state);
//   * Apply-styles additionally refuses armed docs (assertNotHidden, plan A5);
//   * Show All gates on NOTHING beyond the tab check: it is the
//     always-available reverse (Word parity), and its one consent boundary —
//     the PURE-sweep path — is a question, not a gate (plan A14).
//
// Everything below is still host-free: the DocsPort seam is the only I/O, so
// the flagship suite drives these verbs against an in-memory fake
// (__tests__/fakeDocs.ts) end to end.

import { planMarkCiteFromPicks } from "./adapterPure";
import type { SelectionPick } from "./adapterPure";
import { MAX_REPLAN_ATTEMPTS } from "./constants";
import { planDeleteAnalytics } from "./deleteAnalytics";
import { assertNoSuggestions, assertNotHidden, assertSingleTab, chunkGroups } from "./guards";
import { parseDocument } from "./parse";
import { planHide } from "./planner";
import { planShowAll, pureSweepConsentNeeded } from "./restore";
import { resolveSettings } from "./settings";
import { planAnalyticify, planApplyStyles } from "./styles";
import {
  DocsApiError,
  DocsPort,
  DocsRequest,
  GDoc,
  GdocsSettings,
  GHideResult,
  GShowAllOutcome,
  GStylesResult,
  PartialApplyError,
  RequestGroup,
  RevisionConflictError,
  RevisionMismatchError
} from "./types";

// ---------------------------------------------------------------------------
// Private protocol machinery
// ---------------------------------------------------------------------------

/** The verb names the per-verb error classes carry (types.ts contract). The
 * Loop-003 analytics verbs widen this to the full set the PartialApplyError /
 * RevisionConflictError unions already accept, so strings.errorMessage indexes
 * each verb's own truthful failure copy (notably the delete-specific partial
 * copy — Show All cannot undo a torn delete). */
type Verb = "hide" | "showAll" | "applyStyles" | "analyticify" | "deleteAnalytics";

/** One batchUpdate payload, already chunked. `degradable` marks applyStyles'
 * named-style batch ONLY: a DocsApiError there flips the degraded flag and
 * the verb continues; on any other batch it propagates (the batch was atomic,
 * so the generic refusal copy stays truthful for it). */
interface PlannedBatch {
  requests: DocsRequest[];
  degradable: boolean;
}

/** What one plan attempt hands the apply loop. `finish` builds the verb's
 * result AFTER the apply outcome is known — applyStyles' namedStylesApplied
 * is an apply fact, unknowable at plan time (styles.ts module header). A
 * zero-batch plan is the legitimate no-op/short-circuit shape: the loop
 * applies nothing and `finish(false)` is returned untouched (Hide's "nothing
 * to hide", Show All's clean doc AND its consent question all ride this). */
interface VerbPlan<R> {
  batches: PlannedBatch[];
  finish: (degraded: boolean) => R;
}

/**
 * The shared protocol runner (see module header). `gate` carries the
 * verb-specific gates beyond the universal tab check; `makePlan` is invoked
 * once per attempt with the FRESH parsed view — a re-plan must never reuse
 * stale indexes, that is its entire reason to exist.
 *
 * Settings are read ONCE, outside the attempt loop: they live in document
 * properties, not document content, so they are not revision-coupled and a
 * re-plan gains nothing from re-reading them.
 *
 * The "first chunk" retry criterion is `applied === 0`, not a literal chunk
 * index: applyStyles' named-style batch may be REJECTED (degraded path)
 * without anything landing, and a mismatch on the next batch is then still a
 * clean nothing-applied state — silently re-planning it is truthful, while
 * keying on index 0 would wrongly report a partial apply.
 */
async function runVerb<R>(
  port: DocsPort,
  verb: Verb,
  gate: (doc: GDoc) => void,
  makePlan: (doc: GDoc, settings: GdocsSettings) => VerbPlan<R>
): Promise<R> {
  // Device defaults are an adapter concern (it embeds UserProperties JSON);
  // core passes null and the doc-properties tier + built-ins resolve (plan S4).
  const settings = resolveSettings(await port.readSettingsJson(), null);

  // Initial attempt + up to MAX_REPLAN_ATTEMPTS silent re-plans (plan D5).
  for (let attempt = 0; ; attempt++) {
    const doc = parseDocument(await port.fetchDocument());
    assertSingleTab(doc);
    gate(doc);
    const plan = makePlan(doc, settings);

    // Revision chain: starts at the fetched revision, then advances to each
    // applyBatch RESPONSE's revision (plan A13 — one get per verb).
    let revisionId = doc.revisionId;
    let degraded = false;
    let applied = 0;
    try {
      for (const batch of plan.batches) {
        try {
          const res = await port.applyBatch(batch.requests, revisionId);
          revisionId = res.revisionId;
          applied++;
        } catch (e) {
          if (batch.degradable && e instanceof DocsApiError) {
            // The named-style 400 (plan A5/D13): the rejected batch was
            // atomic, so the revision did not move — the chain continues
            // from the same id and the retro batches still land.
            degraded = true;
            continue;
          }
          throw e;
        }
      }
    } catch (e) {
      // ANY mid-sequence failure AFTER >= 1 chunk landed is a partial apply,
      // whatever the error class (exec-review BLOCKER fix): for the destructive
      // Delete-analytics verb a raw DocsApiError on a non-first chunk would map
      // to "nothing was applied" while content was ALREADY deleted — a data-loss
      // lie. PartialApplyError routes to truthful per-verb copy ("some was
      // removed; Show All will not bring it back; re-run to finish"). It is the
      // honest read for the style/reconcile verbs too (a torn Hide → "use Show
      // All"). This must precede the nothing-landed handling below.
      if (applied > 0) throw new PartialApplyError(verb, applied, plan.batches.length);
      // Nothing landed. A first-chunk revision mismatch is silently retryable;
      // every other clean-nothing-applied failure propagates for errorMessage
      // to map (e.g. a single atomic batch the API rejected — errors.docsApi's
      // "nothing was applied" is truthful in that case).
      if (!(e instanceof RevisionMismatchError)) throw e;
      if (attempt >= MAX_REPLAN_ATTEMPTS) throw new RevisionConflictError(verb);
      continue; // nothing landed — silent, immediate re-plan from a fresh fetch
    }
    return plan.finish(degraded);
  }
}

/** Chunk RequestGroups into non-degradable batches — the shape every verb
 * except applyStyles' named-style batch uses (DRY over the three verbs). */
function regularBatches(groups: RequestGroup[]): PlannedBatch[] {
  return chunkGroups(groups).map((requests) => ({ requests, degradable: false }));
}

// ---------------------------------------------------------------------------
// Public verbs
// ---------------------------------------------------------------------------

/**
 * Hide: the A1 reconcile over the whole doc. Gated on suggestions (plan D5)
 * because Hide writes against absolute indexes; refusal copy owns the escape
 * route. The result is the SUCCESSFUL attempt's planner result — exactly the
 * plan that applied, so the receipt's counts can never describe a stale view.
 */
export async function hide(port: DocsPort): Promise<GHideResult> {
  return runVerb(port, "hide", assertNoSuggestions, (doc, settings) => {
    const { groups, result } = planHide(doc, settings);
    return { batches: regularBatches(groups), finish: () => result };
  });
}

/**
 * Show All: restore + convergence sweep, deliberately ungated beyond the tab
 * check (plan A7 — a teammate's pending suggestion must never lock a doc in
 * the hidden state).
 *
 * The consent handshake (plan A14): when the doc carries ZERO rstm ranges but
 * shows tiny text, the tiny text might be the user's own formatting — so when
 * the caller has not answered (`opts.sweepUnrecorded` undefined) the verb
 * returns `needsConsent` WITHOUT writing anything, and the adapter re-invokes
 * with the answer. An explicit `false` proceeds and leaves the unrecorded
 * text untouched (planShowAll(doc, false) on a pure-sweep doc is a clean
 * zero-group no-op by design); `true` adopts it into the sweep. On an ARMED
 * doc the flag is irrelevant — orphans next to our own state are ours.
 */
export async function showAll(
  port: DocsPort,
  opts?: { sweepUnrecorded?: boolean }
): Promise<GShowAllOutcome> {
  return runVerb<GShowAllOutcome>(
    port,
    "showAll",
    () => undefined, // no verb gates: the always-available reverse (plan A7)
    (doc) => {
      const consent = pureSweepConsentNeeded(doc);
      if (consent !== null && opts?.sweepUnrecorded === undefined) {
        // Ask first — zero batches means zero writes on this path.
        return {
          batches: [],
          finish: (): GShowAllOutcome => ({ kind: "needsConsent", unrecordedTinyCount: consent })
        };
      }
      const plan = planShowAll(doc, opts?.sweepUnrecorded ?? false);
      // Edge row 16 rides INSIDE plan.result (rangesSkippedNewerVersion), so
      // passing the result through is the whole pass-through: the receipt's
      // amber newer-version line renders from the same counted contract as
      // every other Show All outcome.
      return {
        batches: regularBatches(plan.groups),
        finish: (): GShowAllOutcome => ({ kind: "done", result: plan.result })
      };
    }
  );
}

/**
 * Apply debate styles: the two-batch styles pass (plan A5). Gates: suggestions
 * (indexes) AND hidden state (restyling an armed doc would write run-level
 * sizes that collide with the RLE restore records — refusal copy says "Show
 * All first"). The named-style batch goes FIRST and is the one degradable
 * batch; the retro batches then run regardless of its fate, so existing
 * paragraphs are always styled and only future typing depends on the
 * named-style write landing (the documented degraded path, plan D13).
 */
export async function applyStyles(port: DocsPort): Promise<GStylesResult> {
  return runVerb(
    port,
    "applyStyles",
    (doc) => {
      assertNoSuggestions(doc);
      assertNotHidden(doc);
    },
    (doc, settings) => {
      const plan = planApplyStyles(doc, settings);
      const named = chunkGroups(plan.namedStyleGroups).map(
        (requests): PlannedBatch => ({ requests, degradable: true })
      );
      return {
        batches: [...named, ...regularBatches(plan.retroGroups)],
        finish: (degraded): GStylesResult => ({
          namedStylesApplied: !degraded,
          restyled: plan.counts.restyled,
          spacingCleared: plan.counts.spacingCleared,
          citesRepaired: plan.counts.citesRepaired
        })
      };
    }
  );
}

/**
 * Analytic-ify (Loop 003): turn the user's touched paragraphs navy 14pt — pure
 * CHARACTER formatting, so it shares the styles lane's gates exactly. Gates:
 * suggestions (the writes target absolute indexes, untrustworthy under pending
 * suggestions) AND hidden state (run-level size writes over an armed doc would
 * collide with the RLE restore records — the same reason Apply-styles refuses;
 * the refusal copy says "Show All first"). assertSingleTab is the universal
 * runVerb check.
 *
 * `ordinals` is the set of paragraph ordinals the adapter lowered the selection
 * / bare cursor down to. It is CAPTURED in the makePlan closure rather than
 * passed through the parsed view: ordinals are a host fact (the live selection),
 * stable across a silent re-plan because the verb only restyles by ordinal and
 * never moves content, so a fresh fetch re-plans the SAME paragraphs (the
 * killed-finding confirms capturing it here is correct — plan §3 controller.ts).
 *
 * regularBatches: one delete-free updateTextStyle group per paragraph, packed
 * into CHUNK_MAX batches. A torn multi-chunk apply throws PartialApplyError
 * ("analyticify"), whose copy says "safe to repeat" because the writes are
 * idempotent (003-S5). finish reports paragraphsStyled for the receipt.
 */
export async function analyticify(
  port: DocsPort,
  ordinals: ReadonlySet<number>
): Promise<{ paragraphsStyled: number }> {
  return runVerb(
    port,
    "analyticify",
    (doc) => {
      assertNoSuggestions(doc);
      assertNotHidden(doc);
    },
    (doc) => {
      const { groups, paragraphsStyled } = planAnalyticify(doc, ordinals);
      return { batches: regularBatches(groups), finish: () => ({ paragraphsStyled }) };
    }
  );
}

/**
 * Delete analytics (Loop 003) — the engine's SOLE content-deleter, confirm-
 * gated by the adapter before it ever reaches here. Same gates as analytic-ify
 * (single-tab via runVerb + no-suggestions + not-hidden): indexes must be
 * trustworthy because deleteContentRange writes against absolute positions, and
 * deleting analytics out of an armed doc would desync the RLE restore records.
 *
 * regularBatches over planDeleteAnalytics' groups — ONE deleteContentRange group
 * per range, already sorted DESCENDING by start index by the planner. That
 * descending order is LOAD-BEARING (a Docs batchUpdate applies requests
 * sequentially and each delete shifts every downstream index), and chunkGroups
 * preserves group order verbatim across revision-chained chunks, so the
 * guarantee holds end to end — including when the chunker splits the sequence
 * into multiple batches.
 *
 * A torn multi-chunk delete throws PartialApplyError("deleteAnalytics"): because
 * each chunk is a clean PREFIX of the descending sequence, the highest-index
 * ranges are already gone and the doc is coherent. Its copy is delete-SPECIFIC
 * (errorMessage indexes by verb) — it must NOT read "nothing was applied" and
 * must warn that Show All cannot bring the removed text back, the one place the
 * generic refusal copy would be a lie. finish reports paragraphsAffected /
 * runsDeleted. A TOCTOU window where the analytics vanished between the adapter's
 * confirm-count read and this verb's fetch simply plans zero groups: the loop
 * applies nothing and finish returns { 0, 0 }, which the receipt renders as the
 * graceful no-op string.
 */
export async function deleteAnalytics(
  port: DocsPort
): Promise<{ paragraphsAffected: number; runsDeleted: number }> {
  return runVerb(
    port,
    "deleteAnalytics",
    (doc) => {
      assertNoSuggestions(doc);
      assertNotHidden(doc);
    },
    (doc) => {
      const { groups, result } = planDeleteAnalytics(doc);
      return { batches: regularBatches(groups), finish: () => result };
    }
  );
}

/**
 * Mark cite — the same silent reconcile the other verbs get (plan D5), so a
 * second Mark cite fired before the first commits (or a teammate's edit
 * mid-flight) no longer rejects with the revision-conflict dialog. Re-marking
 * is IDEMPOTENT: planMarkCiteFromPicks emits the same cite-style writes for the
 * same picks, and applying a cite style that is already present is a no-op — so
 * unlike Hide a conflict here is safe to retry WHOLESALE (no PartialApply
 * distinction). The picks' paragraph ordinals/offsets are stable under the
 * style-only edits cites and Hide make, so re-planning from a FRESH fetch lands
 * the same marks on the new revision.
 *
 * Mark cite lives here (not via runVerb) because its plan comes from a HOST
 * fact — the selection, lowered to picks by the adapter — instead of the parsed
 * doc alone; the gates and the A9 whitelist still live inside
 * planMarkCiteFromPicks. Returns the cited-paragraph count for the receipt.
 *
 * On exhausting MAX_REPLAN_ATTEMPTS the raw RevisionMismatchError stands (the
 * "doc changed — nothing applied — try again" copy): a single-chunk mark is
 * atomic so that is truthful, and a multi-chunk mark's idempotent re-run makes
 * "try again" the correct recovery regardless.
 */
export async function markCite(port: DocsPort, picks: readonly SelectionPick[]): Promise<number> {
  for (let attempt = 0; ; attempt++) {
    const plan = planMarkCiteFromPicks(await port.fetchDocument(), picks);
    // Each cite write is independent: one group per request lets chunkGroups
    // pack them into <=CHUNK_MAX batches with safe boundaries, revision-chained
    // across chunks (the Mark-cite shape the adapter used before, now retried).
    let revisionId = plan.revisionId;
    try {
      for (const requests of chunkGroups(plan.requests.map((r) => ({ requests: [r] })))) {
        ({ revisionId } = await port.applyBatch(requests, revisionId));
      }
      return plan.citedParagraphs;
    } catch (e) {
      if (!(e instanceof RevisionMismatchError) || attempt >= MAX_REPLAN_ATTEMPTS) throw e;
      // Nothing to lose: re-read the fresh revision and re-mark (idempotent).
    }
  }
}
