// Manifest generator — turns the headless feature contributions into the add-in manifest's ribbon
// (one <Group> per feature) and its resources. This is the extensibility payoff: the static
// manifest Office reads at sideload becomes a faithful PROJECTION of the registry, so adding a
// tool never means hand-editing XML — you register it and run `npm run gen:manifest`.
//
// Pure + host-free: `buildManifestXml` takes the contributions + a static config and returns the
// full XML string (no fs, no Office), so it is unit-tested in Node and the committed manifest.xml
// is guarded against drift by a test that re-generates and compares.
//
// Office constraint encoded here: a manifest **resid** is capped at 32 chars, so generated resids
// are short opaque counters (Lbl0/Tip0/Url0/Grp0) while the human text lives in the DefaultValue
// (uncapped) and element `id`s (cap 125) stay descriptive for debugging.
import { FeatureContribution } from "./types";
import { MANIFEST_VERSION } from "./version";

/** The non-feature-derived manifest values (mirrors the current manifest's static header). */
export interface ManifestConfig {
  id: string;
  version: string;
  providerName: string;
  defaultLocale: string;
  displayName: string;
  description: string;
  /** Top-level catalog icons (separate from the in-app ribbon icons). */
  iconUrl: string;
  highResolutionIconUrl: string;
  supportUrl: string;
  /** Dev/prod origin that all SourceLocation + asset URLs are built from. */
  origin: string;
  getStarted: { title: string; description: string; learnMoreUrl: string };
  tab: { id: string; label: string };
}

/** The live config — the static half of the manifest. Bump `version` when the ribbon changes. */
export const manifestConfig: ManifestConfig = {
  id: "16270306-fc6b-47f5-809b-e15b151475e6",
  // PRODUCT VERSION (semver; mirrors package.json). 0.1.0 = the initial GitHub ship of Invisibility
  // Mode ALONE. 1.0.0 is RESERVED for the full suite — every tool (Format & Condense, Flow, Cite &
  // Paste, …) actually WIRED rather than "coming soon" — the build that truly competes with Verbatim,
  // and the first AppSource-submittable build (AppSource requires Version ≥ 1.0). Until then we stay
  // in 0.x. The Office <Version> is 4-part: keep the first three == the semver and bump the 4th
  // ("revision") for a ribbon-STRUCTURE re-register within the same product version — Office caches
  // the ribbon by Id+Version, so only a Version change drops a removed group on re-sideload. (Earlier
  // builds were mis-numbered 1.x; correcting DOWN to 0.x is a one-time WEF cache clear before the next
  // re-sideload, since Office won't treat a lower version as an update.)
  // 0.2.0 = the SECOND tool wired (Condense & Shrink), per the project's feature-milestone versioning
  // (MINOR digit climbs as the suite fills in toward the 1.0.0 full-suite milestone — see LESSONS). The
  // 4th digit re-registers the ribbon within a product version; bumping the MINOR already forces Office
  // to re-read the new "Condense" group on re-sideload, so the revision stays 0.
  // 0.3.0 = the Settings group (a suite-level, app-wide settings home with its own gear-iconned ribbon
  // group + deep-linked pane). NOTE: 0.3.x began as a SHARED-RUNTIME "Always-On" spike (auto-load on
  // every document via Office.addin.setStartupBehavior) — that feature was RETIRED 2026-06-08: it can
  // never reach a brand-new doc (setStartupBehavior is per-document), and "Rostrum on every document"
  // is now delivered by the Trusted-Catalog sideload instead. So this manifest is a PLAIN TaskPaneApp
  // again — no <Runtimes lifetime="long">, no SharedRuntime requirement — and only the Settings group
  // survives from the spike.
  //   .1 — the always-on removal: <Runtimes> + the SharedRuntime requirement stripped, the Settings
  //        group kept (leftmost, gear icon). This RESETS the abandoned .2–.5 shared-runtime revisions
  //        (each of which only existed to make Always-On work). Because the structure shrank back to a
  //        plain TaskPaneApp, re-sideloading over a machine that saw the old .5 spike needs a one-time
  //        WEF cache clear (Office won't treat a lower revision as an update — see LESSONS).
  // The digits themselves live in features/version.ts (single source; guarded against package.json
  // drift by version.test.ts) — bump them THERE.
  version: MANIFEST_VERSION,
  providerName: "Rostrum",
  defaultLocale: "en-US",
  displayName: "Rostrum",
  description:
    "Rostrum — a debating suite for Microsoft Word: hide card bodies, condense and shrink cards, and apply debate styles, all while keeping headings, cites, analytics, and highlighted runs — and natively reversible, even without the add-in.",
  iconUrl: "https://localhost:3000/assets/icon-32.png",
  highResolutionIconUrl: "https://localhost:3000/assets/icon-80.png",
  supportUrl: "https://example.com/rostrum/support",
  origin: "https://localhost:3000",
  getStarted: {
    title: "Rostrum is ready.",
    description: "Open a tool from the Rostrum tab — Invisibility hides card bodies (Show All reverses), Condense & Shrink compress them. Every change is natively reversible.",
    learnMoreUrl: "https://example.com/rostrum",
  },
  tab: { id: "Rostrum.Tab", label: "Rostrum" },
};

// ── Production hosting override layer ────────────────────────────────────────
// `manifestConfig` above is the DEV manifest (localhost:3000), and the committed
// manifest.xml is its byte-exact projection — the drift test pins that. To ship the
// add-in on a real host (GitHub Pages) we must NOT mutate that const (it would red the
// drift test) and must NOT overwrite the committed file. Instead `prodConfig()` returns
// a NEW config with every hosted-URL field rebased onto a real origin; the generator
// writes that to a build artifact (dist/manifest.xml), never to the committed file.

/** Prod add-in identity — DISTINCT from the dev `id` so the localhost-dev manifest and the
 *  Pages-hosted prod manifest can be sideloaded on the SAME machine without their ribbons
 *  colliding (Office caches the ribbon by Id+Version; a shared Id makes the two clobber each
 *  other in the WEF cache). Minted ONCE with `[guid]::NewGuid()`; never change it — changing
 *  it orphans every prior install (Office treats a new Id as a different add-in). */
export const PROD_ID = "ea3fb238-6832-4f91-9654-b9e7ef24d926";

/** The overrides a caller supplies to target a hosted origin. Only `origin` is required;
 *  the rest default to sensible repo-derived values so a one-arg call is enough. */
export interface ProdOverrides {
  /** Hosted origin, e.g. `https://andrewtjin.github.io/rostrum` (a trailing slash is tolerated). */
  origin: string;
  /** Add-in `<Id>`; defaults to {@link PROD_ID}. Override only for a separate staging install. */
  id?: string;
  /** `<SupportUrl>`; defaults to the repo issues page. */
  supportUrl?: string;
  /** GetStarted "Learn more" link; defaults to the Pages landing page (`origin/`). */
  learnMoreUrl?: string;
}

/**
 * Project the dev `manifestConfig` onto a hosted `origin`, returning a NEW config (the dev
 * const is never mutated). `buildManifestXml` already builds taskpane/commands/ribbon-icon
 * URLs from `config.origin`, so rebasing `origin` covers those automatically; the top-level
 * catalog icons + support/learn URLs are hardcoded to localhost in the dev const, so they are
 * the fields this helper must explicitly rewrite.
 */
export function prodConfig(o: ProdOverrides): ManifestConfig {
  // Tolerate a trailing slash so callers can paste a Pages URL verbatim; all URL building
  // below assumes no trailing slash (otherwise we'd emit `…/rostrum//assets/...`).
  const origin = o.origin.replace(/\/+$/, "");
  return {
    ...manifestConfig,
    id: o.id ?? PROD_ID,
    origin,
    iconUrl: `${origin}/assets/icon-32.png`,
    highResolutionIconUrl: `${origin}/assets/icon-80.png`,
    supportUrl: o.supportUrl ?? "https://github.com/andrewtjin/rostrum/issues",
    getStarted: {
      ...manifestConfig.getStarted,
      learnMoreUrl: o.learnMoreUrl ?? `${origin}/`,
    },
  };
}

/** Escape a value for an XML attribute / text node. */
function xml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Build the complete manifest XML from the contributions + config. Deterministic (stable order +
 * counters) so the output is byte-stable for the drift test.
 */
export function buildManifestXml(features: FeatureContribution[], config: ManifestConfig): string {
  const taskpaneBase = `${config.origin}/taskpane.html`;
  const icon = (size: number): string => `${config.origin}/assets/icon-${size}.png`;

  // Resource accumulators (resids stay ≤32 chars; text goes in DefaultValue).
  const shortStrings: Array<{ id: string; value: string }> = [];
  const longStrings: Array<{ id: string; value: string }> = [];
  const urls: Array<{ id: string; value: string }> = [];
  // Per-feature ribbon-icon images (populated when a feature declares `ribbon.icon`); appended to the
  // shared Rostrum logo images in the Resources section. Empty when every feature uses the shared icon.
  const featureImages: Array<{ id: string; value: string }> = [];
  let grpN = 0;
  let lblN = 0;
  let tipN = 0;
  let urlN = 0;

  const groupsXml = features
    .map((feature) => {
      const groupLabelId = `Grp${grpN++}`;
      shortStrings.push({ id: groupLabelId, value: feature.ribbon.label });

      // Per-feature ribbon icon: a feature may declare `ribbon.icon` (an assets base name) to carry its
      // OWN glyph (Settings → a gear); otherwise the group + its controls reference the shared Rostrum
      // logo resids. Any custom-icon image resources are registered ONCE here and reused by this group
      // and all of its controls (resid `Rostrum.Icon.<id>.<size>` is ≤32 chars for the current ids).
      const iconSizes = [16, 32, 80];
      const iconResids = iconSizes.map((s) =>
        feature.ribbon.icon ? `Rostrum.Icon.${feature.id}.${s}` : `Rostrum.Icon.${s}`
      );
      if (feature.ribbon.icon) {
        iconSizes.forEach((s, i) => {
          featureImages.push({ id: iconResids[i], value: `${config.origin}/assets/${feature.ribbon.icon}-${s}.png` });
        });
      }
      const renderIcon = (pad: string): string =>
        [
          `${pad}<Icon>`,
          ...iconSizes.map((s, i) => `${pad}  <bt:Image size="${s}" resid="${iconResids[i]}" />`),
          `${pad}</Icon>`,
        ].join("\n");

      const controlsXml = feature.ribbon.controls
        .map((control) => {
          const labelId = `Lbl${lblN++}`;
          const tipId = `Tip${tipN++}`;
          shortStrings.push({ id: labelId, value: control.label });
          longStrings.push({ id: tipId, value: control.tip });

          const disc = control.kind === "action" ? control.commandId : "pane";
          const controlId = `Rostrum.${feature.id}.${disc}`;

          let actionXml: string;
          if (control.kind === "action") {
            actionXml = [
              `                  <Action xsi:type="ExecuteFunction">`,
              `                    <FunctionName>${xml(control.commandId)}</FunctionName>`,
              `                  </Action>`,
            ].join("\n");
          } else {
            const urlId = `Url${urlN++}`;
            urls.push({ id: urlId, value: `${taskpaneBase}#${feature.id}` });
            // Distinct TaskpaneId per pane-bearing feature → each feature is its own pane.
            const taskpaneId = `Rostrum.Tp.${feature.id}`;
            actionXml = [
              `                  <Action xsi:type="ShowTaskpane">`,
              `                    <TaskpaneId>${xml(taskpaneId)}</TaskpaneId>`,
              `                    <SourceLocation resid="${urlId}" />`,
              `                  </Action>`,
            ].join("\n");
          }

          return [
            `                <Control xsi:type="Button" id="${xml(controlId)}">`,
            `                  <Label resid="${labelId}" />`,
            `                  <Supertip>`,
            `                    <Title resid="${labelId}" />`,
            `                    <Description resid="${tipId}" />`,
            `                  </Supertip>`,
            renderIcon("                  "),
            actionXml,
            `                </Control>`,
          ].join("\n");
        })
        .join("\n");

      return [
        `              <Group id="Rostrum.Group.${xml(feature.id)}">`,
        `                <Label resid="${groupLabelId}" />`,
        renderIcon("                "),
        controlsXml,
        `              </Group>`,
      ].join("\n");
    })
    .join("\n");

  const imagesXml = [
    ...[16, 32, 80].map((s) => `        <bt:Image id="Rostrum.Icon.${s}" DefaultValue="${xml(icon(s))}" />`),
    ...featureImages.map((im) => `        <bt:Image id="${im.id}" DefaultValue="${xml(im.value)}" />`),
  ].join("\n");

  const urlsXml = [
    // taskpane.html is BOTH the ribbon function file (Office loads it in an ephemeral runtime to run an
    // ExecuteFunction command's handler) AND the deep-linked pane shown on demand — so the <FunctionFile>
    // points here. (It replaced the old separate commands.html in 0.3.0; there is no shared runtime.)
    `        <bt:Url id="Rostrum.Taskpane.Url" DefaultValue="${xml(taskpaneBase)}" />`,
    `        <bt:Url id="Rostrum.GetStarted.LearnMoreUrl" DefaultValue="${xml(config.getStarted.learnMoreUrl)}" />`,
    ...urls.map((u) => `        <bt:Url id="${u.id}" DefaultValue="${xml(u.value)}" />`),
  ].join("\n");

  const shortStringsXml = [
    `        <bt:String id="Rostrum.GetStarted.Title" DefaultValue="${xml(config.getStarted.title)}" />`,
    `        <bt:String id="Rostrum.Tab.Label" DefaultValue="${xml(config.tab.label)}" />`,
    ...shortStrings.map((s) => `        <bt:String id="${s.id}" DefaultValue="${xml(s.value)}" />`),
  ].join("\n");

  const longStringsXml = [
    `        <bt:String id="Rostrum.GetStarted.Description" DefaultValue="${xml(config.getStarted.description)}" />`,
    ...longStrings.map((s) => `        <bt:String id="${s.id}" DefaultValue="${xml(s.value)}" />`),
  ].join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!--
  Rostrum add-in manifest. GENERATED by tools/gen-manifest.ts from the feature registry
  (src/features/contributions.ts) — DO NOT EDIT BY HAND. To change the ribbon, edit a feature's
  ribbon descriptor and run \`npm run gen:manifest\`. A drift test (__tests__/ribbonManifest.test.ts)
  fails if the committed file and the generator diverge.

  Desktop-only: the hide engine relies on font-hidden / <w:vanish/> (WordApiDesktop 1.2), so the
  <Requirements> floor hard-blocks Word for the web and old perpetual builds. (No SharedRuntime: the
  0.3.x Always-On spike that needed it was retired — this is a plain TaskPaneApp again.)
-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <Id>${xml(config.id)}</Id>
  <Version>${xml(config.version)}</Version>
  <ProviderName>${xml(config.providerName)}</ProviderName>
  <DefaultLocale>${xml(config.defaultLocale)}</DefaultLocale>
  <DisplayName DefaultValue="${xml(config.displayName)}" />
  <Description DefaultValue="${xml(config.description)}" />
  <IconUrl DefaultValue="${xml(config.iconUrl)}" />
  <HighResolutionIconUrl DefaultValue="${xml(config.highResolutionIconUrl)}" />
  <SupportUrl DefaultValue="${xml(config.supportUrl)}" />

  <Hosts>
    <Host Name="Document" />
  </Hosts>

  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="WordApiDesktop" MinVersion="1.2" />
      <Set Name="WordApi" MinVersion="1.4" />
    </Sets>
  </Requirements>

  <DefaultSettings>
    <SourceLocation DefaultValue="${xml(taskpaneBase)}" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="Rostrum.GetStarted.Title" />
            <Description resid="Rostrum.GetStarted.Description" />
            <LearnMoreUrl resid="Rostrum.GetStarted.LearnMoreUrl" />
          </GetStarted>
          <FunctionFile resid="Rostrum.Taskpane.Url" />

          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="${xml(config.tab.id)}">
${groupsXml}
              <Label resid="Rostrum.Tab.Label" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
${imagesXml}
      </bt:Images>
      <bt:Urls>
${urlsXml}
      </bt:Urls>
      <bt:ShortStrings>
${shortStringsXml}
      </bt:ShortStrings>
      <bt:LongStrings>
${longStringsXml}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}
