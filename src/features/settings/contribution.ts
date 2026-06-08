// Settings — the HEADLESS contribution (no React). The suite's home for GENERAL, app-wide settings
// that belong to no single tool — starting with Always-On ("load Rostrum on every document"). It is a
// first-class feature so it gets its own ribbon group + deep-linked pane via the same registry seam as
// every tool; the React pane lives in `panel.tsx` and is attached in `feature.ts`.
//
// Unlike Invisibility / Condense this feature contributes NO ribbon ExecuteFunction commands — it is a
// single deep-linked pane (the same pattern the planned panes use), so its one ribbon button is a
// `kind: "pane"` control and `commands` is empty (nothing for `associateAll` to wire).
import { FeatureContribution, RibbonGroup } from "../types";

const ribbon: RibbonGroup = {
  label: "Settings",
  // The standard settings GEAR (assets/gear-16|32|80.png) instead of the shared Rostrum "R" logo — a
  // feature carrying its own ribbon glyph via the generator's per-feature icon support.
  icon: "gear",
  controls: [
    {
      kind: "pane",
      label: "Open",
      tip: "Open Rostrum's general settings — including Always-On (load Rostrum on every document).",
    },
  ],
};

export const settingsContribution: FeatureContribution = {
  id: "settings",
  title: "Settings",
  // `tagline` / `glyph` are ComingSoon metadata; a stable feature with a panel never renders ComingSoon,
  // so these are inert but the type requires them. Keep them sensible for completeness.
  tagline: "General Rostrum settings.",
  glyph: "⚙", // text-presentation gear (not the ⚙️ emoji variant, which renders inconsistently small)
  status: "stable",
  primarySurface: "pane",
  // Always openable wherever the suite loads — the pane itself needs no special capability. The
  // SharedRuntime cap-gate lives INSIDE the Always-On widget (it self-hides where the lever can't work),
  // never on the whole feature: gating visibility on an optional cap is the trap LESSONS already records.
  isAvailable: () => true,
  ribbon,
  commands: [],
};
