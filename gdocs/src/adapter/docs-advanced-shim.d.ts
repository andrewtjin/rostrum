// Ambient typing for the Advanced Docs Service global (plan A4).
//
// WHY THIS FILE EXISTS: @types/google-apps-script (v1.0.x) ships typings for
// the built-in services (DocumentApp, PropertiesService, HtmlService, ...) but
// NOT for advanced-service globals — `Docs` (enabled in gdocs/appsscript.json,
// userSymbol "Docs") has no declaration there, so typecheck:gdocs would fail
// on the adapter's two calls. This shim types EXACTLY those two calls and
// nothing more (deliberately minimal: any other Docs.* use should fail the
// typecheck until someone consciously widens this file).
//
// Request/response bodies stay loose on purpose: the REAL contracts live in
// core (types.ts DocsRequest for requests; parse.ts owns interpreting the get
// response, which is why fetchDocument returns `unknown` end to end). Typing
// them richly here would create a second, unverified source of wire truth.

declare const Docs: {
  Documents: {
    /**
     * documents.get — returns the raw Document JSON. `fields` carries
     * parse.DOC_FIELDS_MASK (the verbs) or adapterPure.STATE_FIELDS_MASK (the
     * sidebar state line); `includeTabsContent` is set only on the verb read
     * (plan A3/A13).
     */
    get(
      documentId: string,
      optionalArgs?: { includeTabsContent?: boolean; fields?: string }
    ): unknown;
    /**
     * documents.batchUpdate — one atomic, revision-guarded write (plan D4/D5).
     * The response's writeControl carries the post-apply revision id the
     * controller chains on (plan A13: never a fresh get).
     */
    batchUpdate(
      resource: {
        requests: ReadonlyArray<object>;
        writeControl?: { requiredRevisionId?: string };
      },
      documentId: string
    ): { writeControl?: { requiredRevisionId?: string } };
  };
};
