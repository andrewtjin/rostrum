// PURE CLI-argument logic for gen-manifest, split out from the file I/O so the dev-vs-prod
// branching can be UNIT-TESTED. The branching is dangerous: a prod-intent invocation with a
// missing/empty/non-https origin must FAIL LOUDLY, because the alternative — silently writing a
// broken prod manifest (e.g. `true/assets/icon.png`) or silently falling back to a dev write the
// caller never asked for — ships an add-in that won't load and gives no signal why. gen-manifest.ts
// keeps only the fs writes; everything decided here is a pure function of (flags, env).
import type { ManifestConfig } from "../src/features/manifestGen";
import { manifestConfig, prodConfig } from "../src/features/manifestGen";

/** A flat bag of parsed `--flag` values (everything is a string; bare flags are "true"). */
export type Flags = Record<string, string>;

/** Parse `--k=v`, `--k v`, and bare `--flag` (→ "true") from an argv slice into a flat record.
 *  Using argv (not just env) keeps invocation identical on Windows PowerShell and Linux CI. */
export function parseFlags(argv: string[]): Flags {
  const flags: Flags = {};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (!arg.startsWith("--")) continue;
    const body = arg.slice(2);
    const eq = body.indexOf("=");
    if (eq >= 0) {
      // `--k=v` (v may itself contain `=`, e.g. a URL with a query — split on the FIRST `=` only).
      flags[body.slice(0, eq)] = body.slice(eq + 1);
      continue;
    }
    // `--k v` form: consume the next token only if it isn't itself a flag; else treat as boolean.
    const next = argv[i + 1];
    if (next !== undefined && !next.startsWith("--")) {
      flags[body] = next;
      i++;
    } else {
      flags[body] = "true";
    }
  }
  return flags;
}

/** What gen-manifest.ts should write, fully decided. `outRelative` is relative to the repo root. */
export interface ManifestPlan {
  mode: "dev" | "prod";
  config: ManifestConfig;
  outRelative: string;
}

/**
 * Decide what manifest to write from parsed flags + env.
 *
 * DEV (no prod signal): writes the committed manifest.xml from the localhost `manifestConfig` —
 * byte-identical to legacy behavior, so the drift test stays green.
 *
 * PROD: requested by ANY of `--origin` / `--out` / `ROSTRUM_ORIGIN` / `ROSTRUM_MANIFEST_OUT`. We
 * detect prod *intent* SEPARATELY from origin *validity* so that an intent signal with a bad origin
 * THROWS (caller exits non-zero) instead of silently degrading to a dev write. Office only loads
 * https manifests, so a non-https origin is rejected here too — catching the typo at build time
 * rather than at sideload time.
 *
 * @throws Error when prod is requested but the origin is missing, empty, or not https://.
 */
export function resolveManifestPlan(flags: Flags, env: NodeJS.ProcessEnv): ManifestPlan {
  const origin = flags.origin ?? env.ROSTRUM_ORIGIN;
  const wantsProd =
    origin !== undefined || flags.out !== undefined || env.ROSTRUM_MANIFEST_OUT !== undefined;

  if (!wantsProd) {
    return { mode: "dev", config: manifestConfig, outRelative: "manifest.xml" };
  }

  // Prod was requested — the origin must be a real https URL. `!origin` rejects both undefined
  // (e.g. `--out` given without `--origin`) and "" (the `--origin=` empty form); the regex rejects
  // the bare-`--origin` footgun ("true") and any http:// typo.
  if (!origin || !/^https:\/\//i.test(origin)) {
    throw new Error(
      `prod manifest requires --origin=https://… (got ${JSON.stringify(origin)}). ` +
        `Example: npm run gen:manifest:prod -- --origin=https://andrewtjin.github.io/rostrum`
    );
  }

  const config = prodConfig({
    origin,
    id: flags.id ?? env.ROSTRUM_ID,
    supportUrl: flags.support ?? env.ROSTRUM_SUPPORT_URL,
    learnMoreUrl: flags.learn ?? env.ROSTRUM_LEARN_URL,
  });
  const outRelative = flags.out ?? env.ROSTRUM_MANIFEST_OUT ?? "dist/manifest.xml";
  return { mode: "prod", config, outRelative };
}
