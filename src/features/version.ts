// The single source of truth for Rostrum's product version. Before this module existed the
// version lived in TWO manually-synced places — package.json ("0.3.0") and manifestGen.ts
// ("0.3.0.1") — with no guard against drift. Now the manifest derives from these constants and
// __tests__/version.test.ts pins package.json to PRODUCT_VERSION, so they can't drift silently.
//
// Versioning rules (see LESSONS): stay in 0.x until the FULL suite is wired; MAJOR = suite-wide
// revamp, MINOR = a new tool wired, PATCH = in-feature changes + bugfixes (never a suffix).

/** The product semver (mirrors package.json; the first three digits of the Office <Version>). */
export const PRODUCT_VERSION = "0.3.2";

/**
 * The Office manifest's 4th ("revision") digit: bumped ONLY to force Office to re-read ribbon
 * STRUCTURE within the same product version — Office caches the registered ribbon by Id+Version,
 * so removed/renamed groups linger until the Version string changes. NOT a bugfix digit.
 * Resets to 0 on every PRODUCT_VERSION bump — a new product version already changes the
 * <Version> string, so Office re-reads the ribbon without needing the revision counter.
 */
export const RIBBON_REVISION = 0;

/** The full 4-part Office manifest <Version>. */
export const MANIFEST_VERSION = `${PRODUCT_VERSION}.${RIBBON_REVISION}`;
