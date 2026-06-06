// The contribution model that makes Rostrum a debating SUITE rather than a single
// Invisibility pane. This is the load-bearing interface of the whole UI: the shell
// (pane / ribbon / dialog) renders whatever FEATURES register here and knows nothing
// about any specific tool. Invisibility Mode is just feature #1 — the headline
// convenience tool, not the core — so adding the next tool (Flow, Cite & Paste, …)
// means registering a new feature, never editing the shell.
//
// A feature is declared in two halves:
//   * FeatureContribution — the HEADLESS half (metadata + ribbon descriptor + commands).
//     React-free, so the ribbon runtime AND the Node manifest generator can import it
//     without pulling in the component tree.
//   * RostrumFeature — the contribution PLUS its optional rendered surfaces (pane panel /
//     dialog), as the React shells consume it.
//
// Why a `type`-only React import: the contribution layer must stay a pure, host-free module
// so it is unit-testable in Node and importable from the ribbon runtime; we reference React's
// component TYPE for the panel/dialog surfaces but pull in no React runtime here.
import type { ComponentType } from "react";
import type { FeatureSupport } from "../core/types";

/** Maturity of a feature, surfaced in ComingSoon / badges so users know what's real vs. coming. */
export type FeatureStatus = "stable" | "preview" | "planned";

/** Where a feature primarily lives — drives the ComingSoon hint and which deep-link it opens. */
export type FeatureSurface = "pane" | "dialog";

/**
 * The outcome of a feature command, normalized so the ribbon, pop-out, and pane can all
 * report it uniformly without knowing a feature's internal result shape (e.g. invisibility's
 * `OpOutcome`). Features adapt their native result into this at the contribution boundary.
 */
export interface CommandResult {
  status: "ok" | "blocked" | "cancelled" | "error" | "noop";
  /** Human-facing one-liner (status bar / ribbon log / progress pop-out). */
  message?: string;
}

/**
 * A command a feature contributes. The SAME definition drives both the ribbon (associated
 * by `id`, which must equal the manifest `<FunctionName>` for ribbon-exposed commands) and
 * any in-pane invocation, so the two surfaces can never drift. `run` builds whatever it
 * needs lazily in the runtime that calls it (the ribbon function-file and the task pane are
 * separate JS contexts — they legitimately hold separate engine instances).
 */
export interface FeatureCommand {
  /** Unique across ALL features; for ribbon commands this is the manifest FunctionName. */
  id: string;
  /** Short label for menus / programmatic invocation. */
  title: string;
  /** Optional supertip / hint. */
  description?: string;
  /** Host-capability gate; defaults to "always available" when omitted. */
  isAvailable?: (features: FeatureSupport) => boolean;
  /** Execute the command in the current runtime, returning a normalized result. */
  run: () => Promise<CommandResult>;
}

/**
 * One ribbon button. Exactly two kinds, mirroring the only two Office manifest actions the
 * generator emits — so the static ribbon stays a faithful projection of the registry:
 *   * "action" → `ExecuteFunction`: runs the registered command `commandId` in the ribbon
 *     runtime with NO task pane (the real-estate-friendly default). `commandId` MUST be one of
 *     the owning feature's `commands` ids (== the manifest `<FunctionName>`).
 *   * "pane"   → `ShowTaskpane`: opens THIS feature's deep-linked pane (`taskpane.html#<id>`).
 *     The generator derives the URL + per-feature TaskpaneId from the owning feature id, so a
 *     "pane" control carries no URL of its own.
 * (Opening the full-window dialog is just an "action" whose command calls openWorkspaceDialog —
 * the manifest never needs a third control kind.)
 */
export type RibbonControl =
  | { kind: "action"; commandId: string; label: string; tip: string }
  | { kind: "pane"; label: string; tip: string };

/** A feature's ribbon group — emitted as one `<Group>` under the Rostrum custom tab. */
export interface RibbonGroup {
  /** Group label shown on the ribbon (e.g. "Invisibility"). */
  label: string;
  /** Buttons in the group, left-to-right. */
  controls: RibbonControl[];
}

/** Props the shell injects into every feature surface (pane panel or dialog workspace). */
export interface FeaturePanelProps {
  /** The detected host capabilities (already confirmed runnable by the shell). */
  features: FeatureSupport;
}

/**
 * The HEADLESS half of a self-describing Rostrum tool: everything the ribbon runtime and the
 * manifest generator need, with NO React. `contributions.ts` lists these, `commands.ts`
 * associates their commands to the ribbon, and `manifestGen.ts` turns their `ribbon` groups
 * into manifest XML. The suite grows by adding one of these (open/closed principle).
 */
export interface FeatureContribution {
  /** Stable unique id (registry key + pane/dialog hash route + per-feature TaskpaneId). */
  id: string;
  /** Display name in the pane header / ComingSoon. */
  title: string;
  /** One-line value proposition (ComingSoon copy). */
  tagline: string;
  /** A compact glyph/emoji (ComingSoon). Kept dependency-free for the scaffold. */
  glyph: string;
  /** Maturity — `planned` features render a ComingSoon surface from metadata alone. */
  status: FeatureStatus;
  /** Primary surface (drives the ComingSoon hint + which deep-link the "open" control targets). */
  primarySurface: FeatureSurface;
  /** Capability gate: is this tool usable on the current host? */
  isAvailable: (features: FeatureSupport) => boolean;
  /** Shown on ComingSoon when `isAvailable` is false or `status` is planned. */
  unavailableReason?: string;
  /**
   * A few concrete things this tool will do, rendered as bullets on its ComingSoon placeholder so
   * the suite PREVIEWS what's coming instead of just asserting that something is. Optional and
   * metadata-only — never affects the ribbon or the generated manifest. Keep to 2–4 short,
   * plain-language phrases (a debater who's never seen Verbatim should get it).
   */
  highlights?: string[];
  /** This feature's ribbon group — the static ribbon is generated from these. */
  ribbon: RibbonGroup;
  /** Ribbon + headless-invokable commands this feature contributes. */
  commands: FeatureCommand[];
}

/**
 * A feature as the React shells see it: the headless contribution plus its optional rendered
 * surfaces. `panel` is the compact deep-linked task pane; `dialog` is the opt-in full-window
 * workspace. A `planned` feature (or one with neither surface) renders a ComingSoon from its
 * metadata alone — so a planned tool needs zero React.
 */
export interface RostrumFeature extends FeatureContribution {
  /** Compact in-pane surface (contextual settings / actions). Optional. */
  panel?: ComponentType<FeaturePanelProps>;
  /** Space-heavy full-window surface for the Office Dialog workspace. Optional. */
  dialog?: ComponentType<FeaturePanelProps>;
}
