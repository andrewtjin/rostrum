// The orchestrator: hide / re-hide / show-all over a WordPort, plus the pure
// per-paragraph classifier they share. This realizes the keeper algorithm from
// plan §1 — whole-document, idempotent, convergent.
//
// `classifyParagraph` is pure (RawParagraph + settings -> new OOXML) and carries
// all the policy; the async functions are thin: gate Track Changes, map the
// classifier over every paragraph, write the changed ones, update the manifest.
//
// LOOP 002 B1 — the Hide pass is TWO-PHASE so the P1 node-direct pipeline can run on a
// SINGLE parsed package (whole body parsed ONCE; each paragraph read + mutated through its
// live node; the package serialized ONCE at commit). Phase A (read-only, paced) builds a
// per-paragraph `VisibilityPlan` WITHOUT mutating; Phase B (apply, paced) mutates each
// paragraph IN PLACE. The pure keeper POLICY (`decideParagraph`) is shared by both phases
// AND by the legacy string-mode `classifyParagraph`, so there is exactly one copy of the
// rules (DRY). A node-backed paragraph (`p.parsed`, set only by the pure whole-body read)
// takes the in-place path; every other caller (no `.parsed`) keeps today's serialize path
// byte-for-byte, so `parseCount.test.ts`'s string invariant and the per-paragraph commit
// path are untouched.

import {
  HideOptions,
  HideResult,
  ParagraphAction,
  ParagraphUpdate,
  RawParagraph,
  ResolvedSettings,
  ShowAllResult,
  WordPort
} from "./types";
import {
  computeRunKeepFlags,
  isHeadingKept,
  paragraphHasCiteRun,
  planCrossGapSeparators
} from "./keepers";
import { ParsedParagraph, VisibilityPlan } from "./ooxml";
import { withTrackChangesGate } from "./guards";
import { MANIFEST_SCHEMA_VERSION, clearManifestPart, saveManifest } from "./manifest";

export interface ParagraphPlan {
  index: number;
  action: ParagraphAction;
  /** True only when the OOXML actually changed (avoids spurious writes). */
  changed: boolean;
  /** New OOXML (equals the input when `changed` is false). */
  ooxml: string;
}

/**
 * Inline images / embedded objects (`<w:drawing>`, `<w:object>`, `<w:pict>`) reference an
 * INTERNAL package part (e.g. an `r:embed` to `/word/media/…`). A per-paragraph commit
 * fragment can't carry that binary without re-bloating every fragment (which would erase the
 * whole-body read's speed), so a `hidePartial` OOXML write of such a paragraph would dangle
 * the reference and the live host rejects it ("we found a problem with its contents"). We keep
 * those paragraphs whole via the native `font.hidden` toggle (no OOXML write) — decision #16
 * keeps images untouched regardless. A cheap string probe (no extra parse) keeps classify fast.
 */
const HAS_INTERNAL_PART = /<w:(?:drawing|object|pict)[\s>/]/;

/**
 * The PURE per-paragraph DECISION — the shared policy, with NO mutation and NO serialize.
 *
 * It is what Phase A computes and Phase B applies, and what the legacy string-mode
 * `classifyParagraph` consumes too, so the keeper rules (table/image keep-whole, heading/cite
 * keep-whole, highlight keep + cross-gap bridge) live in EXACTLY one place (DRY — the three
 * review contracts forbid duplicating them across the node and string paths).
 *
 *   * `reveal` — the keepWhole self-heal: clear all vanish (table / internal-part / heading /
 *     cite). The apply side calls `makeAllVisible` (string) or `makeAllVisibleInPlace` (node).
 *   * `plan` — a `VisibilityPlan` (hideFlags + hideParaMark + splits). The apply side calls
 *     `applyVisibility` (string) or `applyVisibilityInPlace` (node).
 *
 * `hasInternalPart` is read from the already-built `RunView[]` (the fused node read fills it for
 * free) so the node path never re-serializes to run the `HAS_INTERNAL_PART` string probe — the
 * P1 rider. The string path passes the cheap string-probe result (computed by the caller, which
 * already holds `p.ooxml`) so its behavior is byte-for-byte unchanged.
 */
type ParagraphDecision =
  | { kind: "reveal"; action: "keepWhole" }
  | { kind: "plan"; action: ParagraphAction; plan: VisibilityPlan };

function decideParagraph(
  parsed: ParsedParagraph,
  headingLevel: number | null,
  inTable: boolean,
  hasInternalPart: boolean,
  settings: ResolvedSettings
): ParagraphDecision {
  const runs = parsed.runs;

  // Tables/images/equations are kept untouched (decision #16). Forcing visible is a no-op on a
  // normal one and self-heals one wrongly hidden before. Internal-part detection comes from the
  // run views on the node path (whole-subtree scan, byte-identical to the string probe).
  if (inTable || hasInternalPart) return { kind: "reveal", action: "keepWhole" };

  // Heading rule (#7) or cite rule (#6b): keep the whole paragraph visible.
  if (isHeadingKept(headingLevel) || paragraphHasCiteRun(runs)) {
    return { kind: "reveal", action: "keepWhole" };
  }

  // Body paragraph: keep highlighted (whole-word) + structural runs; hide the rest.
  const keep = computeRunKeepFlags(runs, settings.keepColors);

  // Bridge separators across hidden gaps so isolated highlighted chunks don't fuse
  // ("radiation" + hidden " are a constant threat, " + "would" -> "radiation would",
  // not "radiationwould" — wet-test bug 1). `extraKeep` promotes an existing
  // whitespace-only run; `splits` move one existing space out of a hidden run into a
  // visible sibling. Neither inserts text, so reversibility stays lossless.
  const { extraKeep, splits } = planCrossGapSeparators(runs, keep);
  for (const i of extraKeep) keep[i] = true;

  const hideFlags = keep.map((k) => !k);
  const anyRuns = runs.length > 0;
  const allHidden = anyRuns && hideFlags.every((h) => h);

  // Collapse the paragraph mark only when an entire *content* paragraph was hidden (decision #5).
  // Empty separator paragraphs (no runs) are left alone.
  const action: ParagraphAction = allHidden
    ? "hideWhole"
    : hideFlags.some((h) => h)
      ? "hidePartial"
      : "keepWhole";
  return { kind: "plan", action, plan: { hideFlags, hideParaMark: allHidden, splits } };
}

/**
 * PURE. Classify one paragraph and return its (possibly unchanged) OOXML — the STRING-MODE path
 * (the legacy contract every existing caller and `parseCount.test.ts` depend on). It parses the
 * paragraph ONCE into a string-backed `ParsedParagraph`, computes the shared `decideParagraph`
 * policy, then applies it through the serializing string-mode methods (`makeAllVisible` /
 * `applyVisibility`). The node-direct Hide path does NOT call this — it runs the same
 * `decideParagraph` over node-backed paragraphs and applies IN PLACE (see `hide`).
 *
 * Order of rules mirrors the plan: table content and headings/cites are kept whole (and forced
 * visible, so re-hide un-hides anything wrongly hidden); otherwise hide every run that isn't a
 * highlighted keeper or structural content, and collapse the paragraph mark when the whole
 * paragraph went hidden.
 */
export function classifyParagraph(
  p: RawParagraph,
  settings: ResolvedSettings
): ParagraphPlan {
  // Parse the paragraph ONCE and read + mutate through the same tree (string mode). The internal-part
  // signal uses the cheap string probe here — byte-for-byte the legacy behavior — so a paragraph whose
  // image part lives outside the `<w:body>` story we scope to is still caught the way it always was.
  const parsed = new ParsedParagraph(p.ooxml);
  const decision = decideParagraph(parsed, p.headingLevel, p.inTable, HAS_INTERNAL_PART.test(p.ooxml), settings);

  if (decision.kind === "reveal") {
    const res = parsed.makeAllVisible();
    return { index: p.index, action: decision.action, changed: res.changed, ooxml: res.xml };
  }
  const { hideFlags, hideParaMark, splits } = decision.plan;
  const res = parsed.applyVisibility(hideFlags, hideParaMark, splits);
  return { index: p.index, action: decision.action, changed: res.changed, ooxml: res.xml };
}

/**
 * One paragraph carried from Phase A (classify) to Phase B (apply): the node-backed
 * `ParsedParagraph` to mutate, the pure decision to apply, and its body index.
 */
interface PhaseAItem {
  index: number;
  parsed: ParsedParagraph;
  decision: ParagraphDecision;
}

/** Hide all non-keeper body text and arm the document (write manifest). */
export async function hide(
  port: WordPort,
  settings: ResolvedSettings,
  opts: HideOptions = {}
): Promise<HideResult> {
  // Reuses the shared Track-Changes gate (guards.ts) — the same one the range-scoped Condense &
  // Shrink controller runs under, so the TC policy lives in exactly one place.
  const { result, toggled } = await withTrackChangesGate(port, opts.autoToggleTrackChanges ?? false, async () => {
    const paras = await port.readParagraphs();

    // ── PHASE A — read-only classify (paced). Build each paragraph's pure decision WITHOUT
    // mutating its tree. A node-backed paragraph (`p.parsed`, the pure whole-body read) is read
    // through its LIVE node (zero parse); every other paragraph parses its own string ONCE — the
    // compat shim `p.parsed ?? new ParsedParagraph(p.ooxml)`, so `parseCount.test.ts`'s meaning is
    // preserved. Per-paragraph skip-on-throw: one malformed paragraph is left visible and counted,
    // never aborting the pass. Pacing/cancel are OUTSIDE the try so a CancelledError aborts the
    // whole pass (pre-write — nothing is buffered yet), not miscounted as a skip.
    const items: PhaseAItem[] = [];
    let skipped = 0;
    for (const p of paras) {
      if (opts.pacing) await opts.pacing.tick();
      try {
        const parsed = (p.parsed as ParsedParagraph | undefined) ?? new ParsedParagraph(p.ooxml);
        const decision = decideParagraph(
          parsed,
          p.headingLevel,
          p.inTable,
          // Node path reads `hasInternalPart` from the run views (free, whole-subtree scan); the
          // string/compat path uses the cheap string probe — byte-for-byte the legacy behavior.
          p.parsed ? parsed.runs.some((r) => r.hasInternalPart) : HAS_INTERNAL_PART.test(p.ooxml),
          settings
        );
        items.push({ index: p.index, parsed, decision });
      } catch {
        // A single malformed paragraph must not abort the whole pass.
        skipped++;
      }
    }

    // ── PHASE B — apply (paced). Mutate each paragraph IN PLACE (node mode) or serialize a new
    // fragment (string/compat mode). CONTRACT (C / 002-F4): ANY throw here aborts the WHOLE op
    // BEFORE a single host write. We catch ONLY to discard the half-mutated in-memory package via
    // `discardPreparedWrite` (so no later commit can serialize it — the node-direct path mutated the
    // port's cached `lastRead.pkg` by reference) and then RE-THROW, so nothing is written, the TC
    // gate's `finally` restores the prior mode, and the manifest stays unarmed. The pacer is
    // re-threaded so a mid-Hide Cancel (a CancelledError from `tick()`) still lands during the apply
    // stretch — and it takes the same discard-then-propagate path. There is NO per-item catch: unlike
    // Phase A's per-paragraph skip, a Phase-B mutation throw means the package is no longer trustworthy.
    const updates: ParagraphUpdate[] = [];
    try {
      for (const it of items) {
        if (opts.pacing) await opts.pacing.tick();
        const nodeMode = it.parsed.isNodeBacked;
        if (it.decision.kind === "reveal") {
          // keepWhole self-heal: clear all vanish. Node mode mutates in place (no serialize);
          // string mode serializes a fresh fragment for the per-paragraph commit.
          if (nodeMode) {
            const res = it.parsed.makeAllVisibleInPlace();
            if (res.changed) updates.push({ index: it.index, action: "keepWhole", ooxml: NODE_DIRECT_OOXML });
          } else {
            const res = it.parsed.makeAllVisible();
            if (res.changed) updates.push({ index: it.index, action: "keepWhole", ooxml: res.xml });
          }
        } else {
          const { plan, action } = it.decision;
          if (nodeMode) {
            const res = it.parsed.applyVisibilityInPlace(plan);
            if (res.changed) updates.push({ index: it.index, action, ooxml: NODE_DIRECT_OOXML });
          } else {
            const res = it.parsed.applyVisibility(plan.hideFlags, plan.hideParaMark, plan.splits);
            if (res.changed) updates.push({ index: it.index, action, ooxml: res.xml });
          }
        }
      }
    } catch (e) {
      // Half-mutated package: abandon it. The port nulls lastRead/pending so the on-disk doc
      // (untouched — nothing was written) is re-read fresh next time. Then propagate (002-F4).
      port.discardPreparedWrite?.();
      throw e;
    }

    await port.writeParagraphs(updates);
    await saveManifest(port, {
      active: true,
      keepColors: [...settings.keepColors],
      schemaVersion: MANIFEST_SCHEMA_VERSION
    });
    return { scanned: paras.length, changed: updates.length, skipped };
  });
  return {
    paragraphsScanned: result.scanned,
    paragraphsChanged: result.changed,
    paragraphsSkipped: result.skipped,
    trackChangesToggled: toggled
  };
}

/**
 * The `ooxml` placeholder for a node-direct update. In node mode the paragraph was already mutated
 * IN PLACE inside the port's cached whole-body package, so the commit serializes the WHOLE package
 * (never a per-paragraph fragment) and the `ooxml` field is unused for the write. It still flows
 * through `writeParagraphs` only to carry the changed-count + index range. A bare `<w:p/>` keeps the
 * single-`<w:p>` write guard happy on the (node-direct) path that never consults it for the actual
 * splice — defense in depth, since the pure commit ignores it entirely.
 */
const NODE_DIRECT_OOXML = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

/**
 * Re-derive over the whole document, catching newly typed/pasted text. This is
 * exactly `hide` again (deterministic re-derivation — decision #10), exported
 * under its own name so callers and the ribbon read clearly.
 */
export async function reHide(
  port: WordPort,
  settings: ResolvedSettings,
  opts: HideOptions = {}
): Promise<HideResult> {
  return hide(port, settings, opts);
}

/**
 * Reveal everything Rostrum hid and disarm the document. Convergent/idempotent:
 * safe to run from any partial state, so a stray Undo can't strand the doc
 * (decision #14). Not TC-gated — Show All is the always-available reverse.
 */
export async function showAll(port: WordPort): Promise<ShowAllResult> {
  // Fast path (Stage 4): clear font.hidden NATIVELY over the whole body story in one/two
  // host round-trips, instead of reading + rewriting every paragraph's OOXML (the old
  // path did `insertOoxml` per changed paragraph — thousands of reflows on a long doc,
  // ~3 min on the extremely-large doc). Behaviorally identical to the per-paragraph
  // makeAllVisible pass: it reveals the SAME set, including any run the user hid manually
  // (decision #10), and is convergent — a second Show All clears nothing new. Malformed
  // OOXML is moot here: nothing is parsed, so `paragraphsSkipped` is always 0. Not TC-
  // gated; Show All is the always-available reverse.
  const res = await port.clearHidden();
  await clearManifestPart(port);
  return {
    paragraphsScanned: res.paragraphsScanned,
    paragraphsChanged: res.paragraphsChanged,
    paragraphsSkipped: 0
  };
}
