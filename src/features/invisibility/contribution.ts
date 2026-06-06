// Invisibility Mode — the HEADLESS contribution (no React). This is what the ribbon runtime
// associates and what the manifest generator turns into a ribbon group; the React pane lives in
// `panel.tsx` and is attached in `feature.ts`. Keeping this React-free is why the ribbon command
// bundle (commands.js) and the Node manifest generator never pull in the component tree.
import { assertCanRun, detectFeatureSupport } from "../../core/guards";
import { RostrumController } from "../../taskpane/controller";
import { createRibbonRunner, emitProgress } from "../ribbonRuntime";
import { FeatureContribution, RibbonGroup } from "../types";

// ===========================================================================
// Ribbon-runtime controller (separate JS context from the task pane).
// ===========================================================================
//
// The in-flight guard, progress pop-out plumbing, and OpOutcome→CommandResult mapping live in the
// SHARED ribbon runtime (`../ribbonRuntime`) so this feature and Condense don't duplicate them — and
// so the in-flight guard is shared across BOTH features (they mutate the one open document). Here we
// only supply how to build invisibility's controller and how to cancel it.
const runner = createRibbonRunner({
  feature: "invisibility",
  build: async () => {
    const features = detectFeatureSupport(Office.context.requirements);
    assertCanRun(features); // throws UnsupportedHostError on web / old perpetual
    const controller = new RostrumController({ features, onProgress: emitProgress });
    await controller.init();
    return controller;
  },
  cancel: (c) => c.cancel(),
});

// ===========================================================================
// The ribbon group + commands. Command ids == the manifest <FunctionName>s the generator emits.
// ===========================================================================

const ribbon: RibbonGroup = {
  label: "Invisibility",
  controls: [
    // Hide is idempotent + convergent, so it doubles as Re-hide: pressing it again re-derives over
    // the whole document and catches newly typed/pasted text. We surface ONE button (no separate
    // Re-hide control) and say so in the tip — the engine `reHide()` is just `hide()` either way.
    { kind: "action", commandId: "hide", label: "Hide", tip: "Hide all non-keeper body text (cards), keeping headings, cites, analytics, and highlighted runs. Run it again after editing to hide newly typed or pasted text." },
    { kind: "action", commandId: "showAll", label: "Show All", tip: "Reveal everything Rostrum hid. Convergent — safe to run from any state." },
    { kind: "action", commandId: "applyStyles", label: "Apply Styles", tip: "Apply Rostrum heading/cite sizes and the pocket box, and repair mis-styled cites. Reflows the document; needs desktop Word 1.5+." },
    { kind: "pane", label: "Settings", tip: "Open the Rostrum pane for keep-color settings, whole-body mode, and diagnostics." },
  ],
};

export const invisibilityContribution: FeatureContribution = {
  id: "invisibility",
  title: "Invisibility Mode",
  tagline: "Collapse a brief to headings, cites, and highlights — instantly, reversibly.",
  glyph: "🙈",
  status: "stable",
  primarySurface: "pane",
  isAvailable: (f) => f.canHide,
  unavailableReason: "Needs desktop Word (hidden-text formatting is WordApiDesktop 1.2).",
  ribbon,
  commands: [
    {
      id: "hide", // == manifest <FunctionName>
      title: "Hide",
      description: "Hide card bodies; keep headings, cites, analytics, and highlights.",
      isAvailable: (f) => f.canHide,
      run: () => runner.run("Hide", (c) => c.hide()),
    },
    {
      id: "showAll",
      title: "Show All",
      description: "Reveal everything Rostrum hid and disarm the document.",
      isAvailable: (f) => f.canHide,
      run: () => runner.run("Show All", (c) => c.showAll()),
    },
    {
      id: "applyStyles",
      title: "Apply Styles",
      description: "Apply Rostrum heading/cite sizes + pocket box, and repair mis-styled cites.",
      isAvailable: (f) => f.canGetStyles && f.canStyleFormat,
      run: () => runner.run("Apply Styles", (c) => c.applyStyles()),
    },
  ],
};
