// Condense & Shrink — the HEADLESS contribution (no React). This is what the ribbon runtime associates
// and what the manifest generator turns into a ribbon group; the React pane lives in `panel.tsx` and is
// attached in `feature.ts`. Keeping this React-free is why the ribbon command bundle and the Node
// manifest generator never pull in the component tree.
//
// Feature #2 of the suite: Rostrum's lossless answer to Verbatim's Shrink + Condense. The ribbon
// surfaces the four direct verbs (Shrink / Condense / Uncondense / Unshrink) plus a Settings pane; the
// pane adds the mode checkboxes, the one-click mode buttons, the live shrink-size readout, and the
// omission-marker editor (see panel.tsx).
import { CondenseController } from "../../taskpane/condenseController";
import { createRibbonRunner } from "../ribbonRuntime";
import { FeatureContribution, RibbonGroup } from "../types";

// The shared ribbon runtime owns the in-flight guard + progress pop-out + result mapping; here we only
// say how to build the Condense controller. No `init()` (it has no document manifest to read) and no
// progress wiring (range ops are a single tiny read+write round-trip, nothing to stream).
//
// `quiet`: these are bounded, instant ops that emit no progress, so a transient "Working…" pop-out is
// pure noise (it can even outlast the op). Suppress it on success — the document visibly changes — and
// let the pop-out appear ONLY when the op errors/blocks (e.g. Track Changes), the ribbon's one channel
// to report a failure the user must act on.
const runner = createRibbonRunner({
  feature: "condense",
  build: async () => new CondenseController({}),
  quiet: true,
});

const ribbon: RibbonGroup = {
  label: "Condense",
  controls: [
    {
      kind: "action",
      commandId: "condenseShrink",
      label: "Shrink",
      tip: "Cycle the selected card's non-underlined text down a font size (8→7→6→5→4→Normal), keeping the underlined cut, highlights, cites, and headings full-size. Press again to shrink further.",
    },
    {
      kind: "action",
      commandId: "condenseRun",
      label: "Condense",
      tip: "Collapse the selection per your Condense settings (merge paragraphs / pilcrows / retain paragraphs). Losslessly reversible with Uncondense.",
    },
    {
      kind: "action",
      commandId: "condenseUncondense",
      label: "Uncondense",
      tip: "Reverse Condense — restore every paragraph break Rostrum marked.",
    },
    {
      kind: "action",
      commandId: "condenseUnshrink",
      label: "Unshrink",
      tip: "Reverse Shrink — reset the selected card's text back to its Normal size.",
    },
    { kind: "pane", label: "Options", tip: "Open Condense & Shrink options — modes, the omission-marker editor, and one-click mode buttons. (App-wide settings live in the Settings group.)" },
  ],
};

export const condenseContribution: FeatureContribution = {
  id: "condense",
  title: "Condense & Shrink",
  tagline: "Shrink cards and condense spacing — losslessly reversible.",
  glyph: "🗜️",
  status: "stable",
  primarySurface: "pane",
  // Needs only the core OOXML round-trip (getSelection/getOoxml/insertOoxml), already guaranteed by the
  // manifest's WordApiDesktop 1.2 + WordApi 1.4 floor — so it is usable wherever Rostrum runs at all.
  isAvailable: () => true,
  ribbon,
  commands: [
    {
      id: "condenseShrink", // == manifest <FunctionName>
      title: "Shrink",
      description: "Cycle non-underlined card text down one font size; keep the underlined cut readable.",
      isAvailable: () => true,
      run: () => runner.run("Shrink", (c) => c.shrink()),
    },
    {
      id: "condenseRun",
      title: "Condense",
      description: "Condense the selection per the current Condense settings (lossless by default).",
      isAvailable: () => true,
      run: () => runner.run("Condense", (c) => c.condense()),
    },
    {
      id: "condenseUncondense",
      title: "Uncondense",
      description: "Restore every paragraph break Rostrum's Condense marked.",
      isAvailable: () => true,
      run: () => runner.run("Uncondense", (c) => c.uncondense()),
    },
    {
      id: "condenseUnshrink",
      title: "Unshrink",
      description: "Reset shrunk text back to its Normal size.",
      isAvailable: () => true,
      run: () => runner.run("Unshrink", (c) => c.unshrink()),
    },
  ],
};
