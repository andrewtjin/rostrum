// Settings resolution + the per-device defaults cache (decision #15).
//
// Two tiers, with the document always winning:
//   1. Per-doc manifest keep-colors (travels with the file — decision #11/#15).
//   2. Per-device defaults in localStorage (labeled per-device, NOT roaming —
//      Word has no roaming-settings API; true roaming is roadmap).
//   3. Built-in default = ALL highlight colors count as keep (decision #11).
//
// Pure precedence logic + a tiny injectable Storage seam, so both are unit tested
// with a fake store (no browser).

import {
  CondenseSettings,
  DeviceDefaults,
  OmissionPattern,
  ResolvedSettings,
  ReversalStrategy,
  RostrumManifest
} from "./types";
import { HIGHLIGHT_COLORS, HIGHLIGHT_COLOR_SET } from "./styles";

/** localStorage key; the `.v1` suffix lets us migrate the cache shape later. */
export const DEVICE_DEFAULTS_KEY = "rostrum.deviceDefaults.v1";

/** Built-in default keep-set: every highlight color (decision #11). */
export const DEFAULT_KEEP_COLORS: readonly string[] = HIGHLIGHT_COLORS.map((c) =>
  c.toLowerCase()
);

/** The slice of `window.localStorage` we use, injectable for tests. */
export interface StorageLike {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

/** Lower-case, validate against known highlight names, and de-duplicate. */
function normalizeColors(colors: readonly string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const c of colors) {
    const v = c.trim().toLowerCase();
    if (v && HIGHLIGHT_COLOR_SET.has(v) && !seen.has(v)) {
      seen.add(v);
      out.push(v);
    }
  }
  return out;
}

/**
 * Resolve the keep-colors the engine runs with, honoring precedence
 * doc-manifest > device-default > built-in. A manifest present with an *empty*
 * keep-set is an explicit user choice and is honored (keep nothing by color).
 */
export function resolveSettings(
  docManifest: RostrumManifest | null,
  deviceDefaults: DeviceDefaults | null
): ResolvedSettings {
  let colors: string[];
  if (docManifest) {
    colors = normalizeColors(docManifest.keepColors);
  } else if (deviceDefaults) {
    colors = normalizeColors(deviceDefaults.keepColors);
  } else {
    colors = [...DEFAULT_KEEP_COLORS];
  }
  return { keepColors: new Set(colors) };
}

/** Read + validate per-device defaults from storage; null when absent/corrupt. */
export function loadDeviceDefaults(storage: StorageLike): DeviceDefaults | null {
  const raw = storage.getItem(DEVICE_DEFAULTS_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (
      typeof parsed === "object" &&
      parsed !== null &&
      Array.isArray((parsed as { keepColors?: unknown }).keepColors)
    ) {
      const colors = (parsed as { keepColors: unknown[] }).keepColors.filter(
        (c): c is string => typeof c === "string"
      );
      return { keepColors: normalizeColors(colors) };
    }
  } catch {
    // fall through to null on malformed JSON
  }
  return null;
}

/** Persist per-device defaults (normalized) to storage. */
export function saveDeviceDefaults(storage: StorageLike, defaults: DeviceDefaults): void {
  const payload: DeviceDefaults = { keepColors: normalizeColors(defaults.keepColors) };
  storage.setItem(DEVICE_DEFAULTS_KEY, JSON.stringify(payload));
}

// The "pure whole-body" Hide path (avenue ⑦) — the DEFAULT as of the 2026-06-05 wet-test (fast +
// losslessly reversible). It reads the whole story in ONE getOoxml, classifies each package paragraph
// directly (outline from the package's own styles.xml, no proxies, no text alignment), and commits
// with ONE insertOoxml("Replace") — fastest path, perfect formatting by construction (the full package
// round-trips). This per-DEVICE flag is the OPT-OUT: when explicitly set false, Rostrum uses the
// slower, maximally-compatible per-paragraph path instead. Absent ⇒ the default (on), see
// `loadPureWholeBody`. (The earlier "ultra-fast" lossy mode was REMOVED — ⑦ is faster AND lossless.)
export const PURE_WHOLE_BODY_KEY = "rostrum.pureWholeBody.v1";

/**
 * Read the per-device pure-whole-body flag. Returns the stored value when the user has explicitly
 * set it (so an opt-OUT sticks); otherwise `defaultValue`. Wet-test 2026-06-05 confirmed ⑦ is fast
 * (2–4s on real docs, 45s on a 960k-word doc) AND losslessly reversible, so production now passes
 * `defaultValue = true` (⑦ is the default Hide path); a user can still uncheck it for the slower,
 * maximally-compatible per-paragraph path. Garbled/throwing storage falls back to `defaultValue`.
 */
export function loadPureWholeBody(storage: StorageLike, defaultValue = false): boolean {
  try {
    const v = storage.getItem(PURE_WHOLE_BODY_KEY);
    return v === null ? defaultValue : v === "true";
  } catch {
    return defaultValue;
  }
}

/** Persist the per-device pure-whole-body flag (best-effort). */
export function savePureWholeBody(storage: StorageLike, on: boolean): void {
  storage.setItem(PURE_WHOLE_BODY_KEY, on ? "true" : "false");
}

// (Always-On removed 2026-06-08: "load Rostrum on every document" is now the Trusted-Catalog
// sideload, not an in-app `setStartupBehavior` toggle — so the per-device always-on intent key and
// its load/save helpers are gone.)

// ---------------------------------------------------------------------------
// Condense & Shrink — per-device settings (localStorage, like the keep-color cache)
// ---------------------------------------------------------------------------

/** localStorage key for the Condense & Shrink settings blob. `.v1` lets us migrate the shape later. */
export const CONDENSE_SETTINGS_KEY = "rostrum.condense.v1";

/**
 * The default omission markers Shrink restores to full size — Verbatim's set ([…Omitted…],
 * [[…Omitted…]], <…Omitted…>), modeled as delimiter/keyword triples. A bracketed span counts as an
 * omission only when it contains the keyword, so ordinary "[sic]"/"[their]" brackets are NOT un-shrunk.
 */
export const DEFAULT_OMISSION_PATTERNS: readonly OmissionPattern[] = [
  { open: "[[", close: "]]", keyword: "Omitted" },
  { open: "[", close: "]", keyword: "Omitted" },
  { open: "<", close: ">", keyword: "Omitted" }
];

/** Built-in Condense & Shrink defaults: lossless markers, merge-all, hidden breaks, no ¶ shrink. */
export const DEFAULT_CONDENSE_SETTINGS: CondenseSettings = {
  usePilcrows: false,
  retainParagraphs: false,
  reversal: "marker",
  shrinkParagraphMarks: false,
  omissionPatterns: DEFAULT_OMISSION_PATTERNS.map((p) => ({ ...p }))
};

/** Validate + normalize an omission pattern list (drop non-string / empty-delimiter entries). */
function normalizeOmissionPatterns(value: unknown): OmissionPattern[] {
  if (!Array.isArray(value)) return DEFAULT_OMISSION_PATTERNS.map((p) => ({ ...p }));
  const out: OmissionPattern[] = [];
  for (const raw of value) {
    if (raw && typeof raw === "object") {
      const o = raw as { open?: unknown; close?: unknown; keyword?: unknown };
      const open = typeof o.open === "string" ? o.open : "";
      const close = typeof o.close === "string" ? o.close : "";
      const keyword = typeof o.keyword === "string" ? o.keyword : "";
      // An omission needs both delimiters; the keyword may be empty (matches any bracketed span).
      if (open && close) out.push({ open, close, keyword });
    }
  }
  return out;
}

/** Coerce a stored value to a valid reversal strategy, defaulting to the lossless "marker". */
function normalizeReversal(value: unknown): ReversalStrategy {
  return value === "none" ? "none" : "marker";
}

/**
 * Read + validate the per-device Condense & Shrink settings from storage, filling any missing/invalid
 * field from the built-in defaults. Garbled or throwing storage falls back to the full defaults — the
 * feature must never be bricked by a bad cache. Pure given an injected store (fake-store tested).
 */
export function loadCondenseSettings(storage: StorageLike): CondenseSettings {
  let parsed: Record<string, unknown> = {};
  try {
    const raw = storage.getItem(CONDENSE_SETTINGS_KEY);
    if (raw) {
      const candidate = JSON.parse(raw) as unknown;
      if (candidate && typeof candidate === "object") parsed = candidate as Record<string, unknown>;
    }
  } catch {
    return { ...DEFAULT_CONDENSE_SETTINGS, omissionPatterns: DEFAULT_OMISSION_PATTERNS.map((p) => ({ ...p })) };
  }
  return {
    usePilcrows: typeof parsed.usePilcrows === "boolean" ? parsed.usePilcrows : DEFAULT_CONDENSE_SETTINGS.usePilcrows,
    retainParagraphs:
      typeof parsed.retainParagraphs === "boolean" ? parsed.retainParagraphs : DEFAULT_CONDENSE_SETTINGS.retainParagraphs,
    reversal: parsed.reversal === undefined ? DEFAULT_CONDENSE_SETTINGS.reversal : normalizeReversal(parsed.reversal),
    shrinkParagraphMarks:
      typeof parsed.shrinkParagraphMarks === "boolean"
        ? parsed.shrinkParagraphMarks
        : DEFAULT_CONDENSE_SETTINGS.shrinkParagraphMarks,
    omissionPatterns:
      parsed.omissionPatterns === undefined
        ? DEFAULT_OMISSION_PATTERNS.map((p) => ({ ...p }))
        : normalizeOmissionPatterns(parsed.omissionPatterns)
  };
}

/** Persist the per-device Condense & Shrink settings (normalized) to storage (best-effort). */
export function saveCondenseSettings(storage: StorageLike, settings: CondenseSettings): void {
  const payload: CondenseSettings = {
    usePilcrows: !!settings.usePilcrows,
    retainParagraphs: !!settings.retainParagraphs,
    reversal: normalizeReversal(settings.reversal),
    shrinkParagraphMarks: !!settings.shrinkParagraphMarks,
    omissionPatterns: normalizeOmissionPatterns(settings.omissionPatterns)
  };
  storage.setItem(CONDENSE_SETTINGS_KEY, JSON.stringify(payload));
}
