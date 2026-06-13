# Rostrum repository architecture

This note records one structural decision and the reasoning behind it, so the
layout is not re-litigated later. It is written for contributors and is not part
of the public site.

## Decision: one repository, two published surfaces

Rostrum ships on two platforms from this single repository:

- the **Word add-in** at the repository root (`src/`, Office.js over Word's
  OOXML), and
- the **Google Docs port** under `gdocs/` (Apps Script over the Docs object
  model).

`gdocs/` is a sibling sub-project of the root project, exactly like the existing
`worker/` (the Cloudflare download counter). It is deliberately **not** an npm
sibling package, and the repository is **not** a monorepo of published packages.

## Why a sibling sub-project, not sibling packages

1. **No shared engine code.** The two engines share nothing at runtime. The Word
   add-in in `src/` is untouched by the port; `gdocs/` carries its own core and
   its own platform adapter. There is no internal library for the two to depend
   on, so a package boundary would have no contract to enforce. It would add
   structure without protecting anything real.
2. **An established precedent.** `worker/` already lives in this repository as a
   self-contained sub-project: its own `.gitignore` and `README.md`, its own
   source tree, its own deploy config (`wrangler.toml`), outside the root build
   and coverage globs, and unit-tested by the root Jest suite. `gdocs/` follows
   the same shape (`gdocs/.gitignore`, `gdocs/README.md`, `gdocs/tsconfig.json`,
   `gdocs/src/`, `gdocs/appsscript.json`). One precedent applied twice is simpler
   than two layout conventions.
3. **One toolchain, one lockfile, one CI pipeline.** A single `npm ci` installs
   everything; the deploy workflow typechecks the Word engine, typechecks and
   builds the Google Docs deliverable, then builds and publishes the site in one
   job. Sibling packages would add a workspace root, per-package lockfile
   surface, and cross-package wiring for no functional gain.

## Type isolation across the two surfaces

The Google Docs adapter references Apps Script ambient globals (`DocumentApp`,
`Docs`, `PropertiesService`, `HtmlService`). Those globals must not leak into the
Word add-in's type space. Isolation is enforced by the two `tsconfig` `types`
arrays:

- The **root `tsconfig.json`** declares a closed `"types": ["office-js", "jest",
  "node"]`, with no `google-apps-script`. Adding it there would make the Apps
  Script globals visible across the entire Word add-in.
- **`gdocs/tsconfig.json`** declares `"types": ["google-apps-script"]` and
  typechecks the whole `gdocs/` tree (core and adapter) with those globals in
  scope.

### Deliberate double-inclusion of the Google Docs core

The root `tsconfig.json` `include` also lists `gdocs/src/core/**/*`. This is
intentional. The Google Docs **core** is pure and platform-agnostic; checking it
under the root config, whose `types` array does not include `google-apps-script`,
proves the core compiles with no Google types in scope. The **adapter** is
excluded from the root config and is checked only under `gdocs/tsconfig.json`,
where the Google globals are available. The core is therefore typechecked twice,
on purpose, to keep the platform boundary honest.

## Accepted tradeoff

Because both surfaces share one `package.json`, the root over-declares the Google
Docs-only `@types/google-apps-script` as a dev dependency. This is accepted.
Splitting into sibling packages would scope that dependency to `gdocs/` alone,
but at the cost of a workspace root, more lockfile and CI surface, and a second
layout convention. That is a poor trade for removing one over-declared type
package. If the Google Docs surface ever grows runtime dependencies of its own,
revisit this note.
