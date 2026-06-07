// Build an absolute URL to one of the add-in's OWN html pages (dialog.html, progress.html),
// correct at ANY base — including a GitHub project-Pages subpath like `…github.io/rostrum/`.
//
// Why this exists: the obvious `${window.location.origin}/dialog.html` is WRONG on a project
// Pages site. `location.origin` is just the scheme+host (`https://andrewtjin.github.io`) and
// DROPS the `/rostrum` path segment, so the URL resolves to `/dialog.html` at the domain root
// → 404. Resolving the page name RELATIVE to the current document URL instead keeps whatever
// base path the bundle is actually served from:
//
//   base = https://andrewtjin.github.io/rostrum/taskpane.html   → https://andrewtjin.github.io/rostrum/dialog.html   ✓
//   base = https://localhost:3000/taskpane.html                 → https://localhost:3000/dialog.html                  ✓
//
// `new URL(relative, base)` replaces the LAST path segment of `base`, which is exactly the
// "sibling page in the same folder" semantics we want. Office serves all four html pages from
// one directory, so every caller is a sibling of its target.
//
// INVARIANT: `base` (window.location.href) must end in an explicit filename segment (e.g.
// `…/rostrum/taskpane.html`), NOT a bare directory (`…/rostrum`). With a filename, the last
// segment is replaced → `…/rostrum/dialog.html` ✓. With a bare directory and no trailing slash,
// `rostrum` itself would be treated as the segment to replace → `…/dialog.html` ✗. This holds in
// practice because Office always loads the add-in at a concrete `*.html` page (the manifest's
// SourceLocation / FunctionFile), never at a directory.

/**
 * Resolve `page` (e.g. `"dialog.html"`) against the current document URL and return the absolute
 * href, optionally with a fragment. `hash` is appended verbatim after `#` (callers pass an
 * already-`encodeURIComponent`-ed value); the URL API will not double-encode existing `%XX`
 * sequences, so the fragment round-trips byte-for-byte with the previous hand-built strings.
 */
export function appPageUrl(page: string, hash?: string): string {
  const url = new URL(page, window.location.href);
  // Assign via `.hash` (not string concat) so the URL serializer owns separator placement; it
  // strips a leading `#` if present and adds one otherwise. Skip empty/undefined to avoid a
  // trailing bare `#`.
  if (hash) url.hash = hash;
  return url.href;
}
