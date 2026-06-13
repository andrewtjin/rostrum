// "Apply Rostrum styles": set the debate template's heading + cite sizes and box
// the Pocket, idempotently, behind a feature gate + an explicit reflow warning.
//
// This is host glue, so the DECISION of what to apply (which style, what size, how
// to draw the box) is a PURE, unit-tested function (`planStyleApplications`), and
// only the execution touches Office.js. The pure plan also encodes the verified
// requirement-set facts: sizes need `Style.font` (WordApi 1.5 = canStyleFormat),
// enumerating styles needs the `document.getStyles()` METHOD (WordApi 1.5 = canGetStyles —
// NOT the WordApiDesktop 1.4 `Document.styles` property, which Rostrum doesn't call),
// and the box prefers `Style.borders` (WordApiDesktop 1.1 = canStyleBorders) with the
// hand-authored `w:pBdr` OOXML as the documented fallback.

import { STYLE_MAP, POCKET_BORDER } from "./styles";
import { FeatureSupport } from "./types";
import { WordRunner, defaultWordRunner } from "./officeWordPort";
import { Logger, logger as rootLogger } from "./debug";

/* eslint-disable @typescript-eslint/no-explicit-any */
// Office Style/Border proxies, like the adapter, are reached through a localized
// `any` — the surface we touch is tiny and version-robust as string literals.

/** How the pocket box should be drawn for the current host. */
export type BorderMethod = "style" | "ooxml" | "none";

/** One concrete style edit the host should make. */
export interface StyleApplication {
  /** STYLE_MAP key: "pocket" | "hat" | "block" | "tag" | "cite". */
  key: string;
  /** Character styles apply to runs; paragraph styles to paragraphs. */
  isCharacterStyle: boolean;
  /** The name Word knows the style by (local heading name, or the cite styleId). */
  styleName: string;
  /** Target point size (decision #9). */
  sizePt: number;
  /** How to draw the boxed border (pocket only). */
  border: BorderMethod;
}

/** The result of an apply pass — structured for the diagnostics panel. */
export interface EnsureStylesResult {
  applied: Array<{ key: string; sizePt: number; border: BorderMethod }>;
  skipped: Array<{ key: string; reason: string }>;
  /** True when the host lacked the style APIs entirely (nothing attempted). */
  unsupported: boolean;
}

/** User-facing warning shown before applying (sizes reflow the whole document). */
export const REFLOW_WARNING =
  "Applying Rostrum styles changes heading and cite sizes across the document, which " +
  "reflows pagination. This can't be undone with Show All — use Word's Undo (Ctrl+Z) to revert.";

/** Built-in heading keys map to the LOCAL style names `getStyles().getByName` expects. */
const BUILTIN_LOCAL_NAME: Readonly<Record<string, string>> = {
  Heading1: "Heading 1",
  Heading2: "Heading 2",
  Heading3: "Heading 3",
  Heading4: "Heading 4"
};

/**
 * PURE. Turn the STYLE_MAP + the host's capabilities into a concrete list of edits.
 * Border method is chosen here: `Style.borders` when available, else the `w:pBdr`
 * OOXML fallback, else none. Unit-tested against the capability matrix.
 */
export function planStyleApplications(features: Pick<FeatureSupport, "canStyleBorders">): StyleApplication[] {
  return Object.entries(STYLE_MAP).map(([key, spec]) => {
    const isCharacterStyle = !!spec.isCharacterStyle;
    const styleName = isCharacterStyle
      ? spec.styleId ?? key
      : BUILTIN_LOCAL_NAME[spec.builtIn ?? ""] ?? spec.builtIn ?? key;
    const border: BorderMethod = spec.boxed ? (features.canStyleBorders ? "style" : "ooxml") : "none";
    return { key, isCharacterStyle, styleName, sizePt: spec.sizePt, border };
  });
}

export interface EnsureStylesOptions {
  features: FeatureSupport;
  runner?: WordRunner;
  logger?: Logger;
}

/**
 * Apply the planned style edits to the live document, idempotently. Each style is
 * isolated in its own try so one missing/locked style can't abort the rest; every
 * outcome is logged and returned for the diagnostics panel. Requires `getStyles()`
 * (canGetStyles) + `Style.font` (canStyleFormat); without them the whole op is a
 * no-op flagged `unsupported`, and the UI hides the button.
 */
export async function ensureRostrumStyles(options: EnsureStylesOptions): Promise<EnsureStylesResult> {
  const { features } = options;
  const run = options.runner ?? defaultWordRunner;
  const log = options.logger ?? rootLogger("styles");
  const span = log.span("ensureRostrumStyles", { features: pickStyleFeatures(features) });

  if (!features.canGetStyles || !features.canStyleFormat) {
    log.warn("host lacks style APIs (getStyles / Style.font) — skipping Apply Styles", {
      canGetStyles: features.canGetStyles,
      canStyleFormat: features.canStyleFormat
    });
    span.end({ unsupported: true });
    return { applied: [], skipped: [], unsupported: true };
  }

  const plan = planStyleApplications(features);
  const applied: EnsureStylesResult["applied"] = [];
  const skipped: EnsureStylesResult["skipped"] = [];

  try {
    await run(async (ctx) => {
      const styles = (ctx.document as any).getStyles();
      // Resolve every target style first, then sync once, then edit + sync once.
      const handles = plan.map((app) => ({ app, style: styles.getByNameOrNullObject(app.styleName) }));
      for (const h of handles) h.style.load("isNullObject");
      await ctx.sync();

      for (const { app, style } of handles) {
        if (style.isNullObject) {
          log.warn("style not found — skipped", { key: app.key, styleName: app.styleName });
          skipped.push({ key: app.key, reason: `style "${app.styleName}" not found` });
          continue;
        }
        try {
          // Size is set via Style.font.size (points).
          style.font.size = app.sizePt;
          // Box: only the Pocket. Prefer Style.borders; the OOXML fallback would be a
          // per-paragraph w:pBdr injection (documented limitation — not applied here).
          let border: BorderMethod = "none";
          if (app.border === "style") {
            applyBoxBorder(style);
            border = "style";
          } else if (app.border === "ooxml") {
            log.warn("pocket box needs Style.borders (WordApiDesktop 1.1); host lacks it — sizes only", {
              key: app.key
            });
            border = "none";
          }
          applied.push({ key: app.key, sizePt: app.sizePt, border });
          log.debug("style applied", { key: app.key, styleName: app.styleName, sizePt: app.sizePt, border });
        } catch (e) {
          log.caught("failed to apply one style (continuing)", e, { key: app.key });
          skipped.push({ key: app.key, reason: "edit failed (see diagnostics)" });
        }
      }
      await ctx.sync();
    });
    span.end({ applied: applied.length, skipped: skipped.length });
  } catch (e) {
    span.fail(e);
    throw new Error(`Rostrum couldn't apply styles. ${describeBriefly(e)}`);
  }

  return { applied, skipped, unsupported: false };
}

/** Set a single 3pt box on all four sides of a style via the Border collection. Width, color,
 *  and style all come from POCKET_BORDER so the live box matches the OOXML fragment and Verbatim's
 *  pocket (3pt). */
function applyBoxBorder(style: any): void {
  for (const side of ["Top", "Left", "Bottom", "Right"]) {
    const border = style.borders.getByLocation(side);
    border.type = "Single";
    // `Word.Border.width` is a BorderWidth STRING enum, not a number — assigning a number is
    // rejected at `ctx.sync()`. "Pt300" is the 3pt member (WordApiDesktop 1.1), matching Verbatim.
    border.width = POCKET_BORDER.borderWidthToken;
    border.color = `#${POCKET_BORDER.color}`;
  }
}

function pickStyleFeatures(f: FeatureSupport): Record<string, boolean> {
  return {
    canGetStyles: f.canGetStyles,
    canStyleFormat: f.canStyleFormat,
    canStyleBorders: f.canStyleBorders
  };
}

function describeBriefly(e: unknown): string {
  if (e && typeof e === "object") {
    const o = e as Record<string, unknown>;
    return `${o.code ? `${o.code}: ` : ""}${o.message ?? "unknown error"}`;
  }
  return String(e);
}
