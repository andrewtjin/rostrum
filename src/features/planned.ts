// The suite's roadmap, expressed as DATA — the Verbatim-rivalling tools Rostrum will grow.
// Registered now as first-class, capability-gated contributions so the RIBBON advertises the
// full suite and the extension path is proven end-to-end: each planned tool already gets its
// own ribbon group that opens its own surface (a ComingSoon pane, or a ComingSoon dialog for
// the space-heavy ones). When a tool is built, swap its `status` to "stable"/"preview", attach
// a `panel`/`dialog` in its feature module, and fill in real `commands` — nothing else changes.
import { FeatureSupport } from "../core/types";
import { CommandResult, FeatureContribution, RibbonControl } from "./types";
import { openWorkspaceDialog } from "../dialog/open";

/** A pane-first planned tool: one ribbon "Open" button → its deep-linked ComingSoon pane. */
function plannedPane(
  fields: Pick<FeatureContribution, "id" | "title" | "tagline" | "glyph" | "highlights"> & { unavailableReason: string }
): FeatureContribution {
  return {
    status: "planned",
    primarySurface: "pane",
    // Planned tools are listed/advertised but not yet a usable feature on any host.
    isAvailable: () => false,
    ribbon: {
      label: fields.title,
      controls: [{ kind: "pane", label: "Open", tip: fields.unavailableReason }],
    },
    commands: [],
    ...fields,
  };
}

/** A dialog-first planned tool: one ribbon "Open" button → an ExecuteFunction that pops its
 *  full-window ComingSoon workspace. The open command is host-runnable (it just shows the
 *  ComingSoon); the `planned` status is what makes the surface say "coming soon". */
function plannedDialog(
  fields: Pick<FeatureContribution, "id" | "title" | "tagline" | "glyph" | "highlights"> & { unavailableReason: string }
): FeatureContribution {
  const openId = `open_${fields.id}`;
  const openControl: RibbonControl = { kind: "action", commandId: openId, label: "Open", tip: fields.unavailableReason };
  return {
    status: "planned",
    primarySurface: "dialog",
    isAvailable: () => false,
    ribbon: { label: fields.title, controls: [openControl] },
    commands: [
      {
        id: openId,
        title: `Open ${fields.title}`,
        description: fields.unavailableReason,
        // Opening a dialog needs no special host capability.
        isAvailable: (_f: FeatureSupport) => true,
        run: async (): Promise<CommandResult> => {
          openWorkspaceDialog(fields.id);
          return { status: "ok" };
        },
      },
    ],
    ...fields,
  };
}

/**
 * The planned suite. Flow is dialog-hosted (a space-heavy speech-doc surface); Format & Condense +
 * Cite & Paste are pane-first. Order here is the ribbon group order after Invisibility Mode.
 * (Card Library — a full-window browse/import workspace — was deferred; re-add a `plannedDialog`
 * entry here when it's ready, and run `npm run gen:manifest`.)
 */
export const plannedContributions: FeatureContribution[] = [
  plannedPane({
    id: "format",
    title: "Format & Condense",
    tagline: "One-click styles cleanup, condensing, and legacy-file migration.",
    glyph: "🧹",
    highlights: [
      "Collapse a brief's tangle of one-off styles down to a clean, canonical set",
      "Merge duplicate and typo'd styles (e.g. Analytic → Analytics)",
      "Condense spacing and bring legacy files up to today's format",
    ],
    unavailableReason: "Planned — extends today's Apply-Styles + condense work.",
  }),
  plannedDialog({
    id: "flow",
    title: "Flow",
    tagline: "Track arguments across a round in a speech-doc surface.",
    glyph: "🗒️",
    highlights: [
      "Lay out a round's arguments side by side in a speech-ready view",
      "Pull cards and tags straight from your open brief",
      "Track what's been answered and what's still standing",
    ],
    unavailableReason: "Planned — the flowing / speech-doc surface.",
  }),
  plannedPane({
    id: "cite",
    title: "Cite & Paste",
    tagline: "Generate citations and paste cards with clean formatting.",
    glyph: "🔖",
    highlights: [
      "Build a clean citation from a source in seconds",
      "Paste cards with formatting that survives the move",
      "Flag cites that are missing fields before a round",
    ],
    unavailableReason: "Planned — citation + paste-with-formatting helpers.",
  }),
];
