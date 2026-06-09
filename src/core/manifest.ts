// Document-level manifest: the single source of truth for "is invisibility ON,
// and with which keep-colors" (decisions #10, #13, #15).
//
// Stored as ONE custom XML part (NOT per-span content controls — those bloat
// 200-page docs and are silently stripped by .doc/RTF/Google-Docs conversion).
// Because it travels with the file and is scanned on open, ANY user/machine
// re-derives the same ON-state and keep-set deterministically (decision #13).
//
// Pure serialize/parse below; the *Manifest(port) helpers are the thin adapter.

import { DOMParser } from "@xmldom/xmldom";
import { RostrumManifest, WordPort } from "./types";
import { HIGHLIGHT_COLOR_SET } from "./styles";

/** Namespace that identifies our custom XML part among all parts in the document. */
export const MANIFEST_NAMESPACE = "https://rostrum.app/invisibility/manifest";

/** Current manifest schema version (decision #10: `{active, keepColors, schemaVersion}`). */
export const MANIFEST_SCHEMA_VERSION = 1;

const ROOT = "rostrum";

/** Minimal XML escaping for text/attribute values. */
function esc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Serialize a manifest to the custom-XML-part string. Pure. */
export function serializeManifest(m: RostrumManifest): string {
  const colors = m.keepColors.map((c) => `<color>${esc(c)}</color>`).join("");
  return (
    `<${ROOT} xmlns="${MANIFEST_NAMESPACE}" schemaVersion="${m.schemaVersion}">` +
    `<active>${m.active ? "true" : "false"}</active>` +
    `<keepColors>${colors}</keepColors>` +
    `</${ROOT}>`
  );
}

/* eslint-disable @typescript-eslint/no-explicit-any */
// xmldom nodes typed as `any` for the same version-robustness reason as ooxml.ts.

function localName(node: any): string {
  // nodeName may be "rostrum" or "ns:rostrum"; compare on the local part.
  const n: string = node.nodeName || "";
  const i = n.indexOf(":");
  return i >= 0 ? n.slice(i + 1) : n;
}

function firstText(parent: any, tag: string): string | null {
  const els = parent.getElementsByTagName(tag);
  if (!els || els.length === 0) return null;
  const t = els.item(0).textContent;
  return t == null ? null : t;
}

/**
 * Parse a manifest string. Throws if it is not a Rostrum manifest. Unknown
 * keep-colors are dropped (forward-compatibility); colors are lower-cased and
 * de-duplicated.
 */
export function parseManifest(xml: string): RostrumManifest {
  // Malformed input is an expected, handled path (parseManifestOrNull catches and
  // returns null). Swallow xmldom's default console.error logging so a doc with
  // no/garbled manifest doesn't spam the task-pane console; we detect failure via
  // the missing/incorrect root element below.
  const doc = new DOMParser({ onError: () => {} }).parseFromString(xml, "text/xml");
  const root = doc.documentElement;
  if (!root || localName(root) !== ROOT) {
    throw new Error("Not a Rostrum manifest");
  }

  // Missing/invalid version => 0 ("unknown"), distinguishable from a real v1 so a
  // future migration guard can detect a hand-edited or legacy manifest.
  const schemaRaw = parseInt(root.getAttribute("schemaVersion") || "", 10);
  const schemaVersion = Number.isFinite(schemaRaw) ? schemaRaw : 0;

  const activeText = (firstText(root, "active") || "").trim().toLowerCase();
  const active = activeText === "true" || activeText === "1";

  const colorEls = root.getElementsByTagName("color");
  const seen = new Set<string>();
  const keepColors: string[] = [];
  for (let i = 0; i < colorEls.length; i++) {
    const el = colorEls.item(i);
    const raw = ((el && el.textContent) || "").trim().toLowerCase();
    if (raw && HIGHLIGHT_COLOR_SET.has(raw) && !seen.has(raw)) {
      seen.add(raw);
      keepColors.push(raw);
    }
  }

  return { active, keepColors, schemaVersion };
}

/** Parse, returning null for null/empty/non-manifest input instead of throwing. */
export function parseManifestOrNull(xml: string | null): RostrumManifest | null {
  if (!xml || xml.trim() === "") return null;
  try {
    return parseManifest(xml);
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// WordPort adapter
// ---------------------------------------------------------------------------

/** Read + parse the document's manifest, or null when absent/invalid. */
export async function loadManifest(port: WordPort): Promise<RostrumManifest | null> {
  const xml = await port.readManifest();
  return parseManifestOrNull(xml);
}

/** Serialize + write the manifest to the document. */
export async function saveManifest(port: WordPort, m: RostrumManifest): Promise<void> {
  await port.writeManifest(serializeManifest(m));
}

/** Remove the manifest from the document (no-op when absent). */
export async function clearManifestPart(port: WordPort): Promise<void> {
  await port.clearManifest();
}

/**
 * On-open detection (decision #13): a document is "armed" (Show All / live-mode
 * UI should appear) iff it carries a manifest with active=true.
 */
export async function isArmed(port: WordPort): Promise<boolean> {
  const m = await loadManifest(port);
  return m !== null && m.active;
}
