// Shared fixture builders for the gdocs engine suites (NOT a .test.ts file —
// same convention as fakeWord.ts). Pure-view builders: tests describe content
// as compact (text, style) tuples and get index-consistent GDoc views back,
// mirroring the Word suites' rv()/seq()/mkPara() culture. The Docs API's
// load-bearing index rules are centralized HERE so every suite inherits them:
// body content starts at index 1, every paragraph ends with a newline char
// included in its range, and indexes are UTF-16 code units (surrogate pairs
// count 2 — `.length` semantics).

import {
  GDoc,
  GElement,
  GNamedRange,
  GNamedStyleType,
  GParagraph,
  GRange
} from "../gdocs/src/core/types";

/** Spec for one element of a paragraph under construction. */
export interface GeSpec {
  text: string;
  /** Explicit size in pt; omit/null = inherits from the named style. */
  size?: number | null;
  bold?: boolean;
  /** Lower-case "#rrggbb"; omit/null = no background. */
  bg?: string | null;
  /** kind "other" models chips/breaks/objects (whitelist rule — never styled).
   * Its `text` still occupies index space (chips have length in the real API). */
  kind?: "text" | "other";
}

/** Spec for one paragraph under construction. */
export interface GpSpec {
  style?: GNamedStyleType;
  inTable?: boolean;
  spaceAbovePt?: number | null;
  spaceBelowPt?: number | null;
  elements: GeSpec[];
}

export interface BuildDocOptions {
  revisionId?: string;
  tabCount?: number;
  suggestionsPresent?: boolean;
  namedRanges?: GNamedRange[];
  namedStyleSizesPt?: GDoc["namedStyleSizesPt"];
}

/** Convenience: a paragraph of a single plain text element. */
export function para(text: string, style: GNamedStyleType = "NORMAL_TEXT"): GpSpec {
  return { style, elements: [{ text }] };
}

/**
 * Build an index-consistent GDoc from paragraph specs. API REALITY RULE: a
 * paragraph's trailing newline lives INSIDE its final textRun's content (the
 * Docs API has no standalone paragraph-mark node), so the builder appends
 * "\n" to the last text element — or, when the paragraph ends in an "other"
 * element (the API never puts the newline in a non-text element), as its own
 * 1-char text element. The LAST paragraph is marked isLastInSegment so planner
 * clamp tests get the real can't-style-the-final-newline constraint for free.
 */
export function buildDoc(paras: GpSpec[], opts: BuildDocOptions = {}): GDoc {
  let cursor = 1; // Docs body content begins at index 1, not 0.
  const paragraphs: GParagraph[] = paras.map((p, i) => {
    const startIndex = cursor;
    const specs: GeSpec[] = [...p.elements];
    const last = specs[specs.length - 1];
    if (last !== undefined && (last.kind ?? "text") === "text") {
      specs[specs.length - 1] = { ...last, text: last.text + "\n" };
    } else {
      // Paragraph ends in a chip/object (or is empty): the newline is its own run,
      // inheriting the paragraph's default size.
      specs.push({ text: "\n" });
    }
    const elements: GElement[] = specs.map((e) => {
      const el: GElement = {
        startIndex: cursor,
        endIndex: cursor + e.text.length,
        kind: e.kind ?? "text",
        text: e.kind === "other" ? "" : e.text,
        fontSizePt: e.size ?? null,
        bold: e.bold ?? false,
        backgroundHex: e.bg ?? null
      };
      cursor += e.text.length;
      return el;
    });
    return {
      index: i,
      startIndex,
      endIndex: cursor,
      namedStyleType: p.style ?? "NORMAL_TEXT",
      inTable: p.inTable ?? false,
      isLastInSegment: i === paras.length - 1,
      spaceAbovePt: p.spaceAbovePt ?? null,
      spaceBelowPt: p.spaceBelowPt ?? null,
      elements
    };
  });
  return {
    revisionId: opts.revisionId ?? "rev-1",
    tabCount: opts.tabCount ?? 1,
    paragraphs,
    namedRanges: opts.namedRanges ?? [],
    namedStyleSizesPt: opts.namedStyleSizesPt ?? { NORMAL_TEXT: 11 },
    suggestionsPresent: opts.suggestionsPresent ?? false
  };
}

/** Shorthand for a named range with one segment. */
export function range(id: string, name: string, startIndex: number, endIndex: number): GNamedRange {
  return { id, name, segments: [{ startIndex, endIndex }] };
}

/** Shorthand for a bare GRange. */
export function r(startIndex: number, endIndex: number): GRange {
  return { startIndex, endIndex };
}
