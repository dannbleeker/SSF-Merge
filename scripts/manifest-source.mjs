/**
 * The add-in's manifest, once, in both spellings Office accepts.
 *
 * Two formats and two environments is four files, and four hand-maintained
 * files is four chances for the production one — the only one anybody
 * sideloads — to say something the others do not. They are generated from the
 * definition below instead, and `test/manifest.test.ts` fails when a committed
 * file stops matching what this produces.
 *
 * Plain `.mjs`, not TypeScript, deliberately. The generated files are read by
 * the test AND by `scripts/build-manifests.mjs`; routing that through
 * `dist-lib/` would mean a manifest built from a stale engine whenever somebody
 * forgot `build:lib` first, which is a failure mode a sibling project already
 * records. Nothing here needs types that a test cannot check by reading the
 * output.
 */

/**
 * The GUID, and the one line in this file that must never change.
 *
 * Office identifies an add-in by it. A new GUID is a NEW add-in: every existing
 * sideload is orphaned, every user has to remove the old entry by hand, and
 * nothing anywhere says why. `test/manifest.test.ts` pins this exact string.
 */
export const ID = "43ebbbac-44ad-42b2-a582-0ef079093e6c";

/**
 * The manifest version, and NOT the npm package version.
 *
 * Office rejects anything below 1.0 outright — "Manifest Version Too Low" — and
 * a sibling project shipped `0.1.0` in four manifests for the whole life of the
 * repo because every one of its own tests passed and nothing had ever asked
 * Microsoft. Four parts, because the XML manifest wants `a.b.c.d`.
 *
 * **Bumped when the MANIFEST changes, not when the project releases.** That is
 * the moment Office has to be told an installed add-in has been updated, and it
 * is the moment AppSource requires a new number: a submission whose version has
 * not moved is refused as an update.
 *
 * 1.0.0.0 → 1.0.1.0 on 2026-08-30, for the first change to this file that a
 * user's PowerPoint would care about: `HighResolutionIconUrl` moved from the
 * 80px ribbon icon to the 64px one AppSource actually asks for, and the support
 * link stopped being a GitHub repository, which a submission refuses.
 */
export const VERSION = "1.0.1.0";

/**
 * The requirement floor, as the manifests' comment states it.
 *
 * A second spelling of `API_FLOOR` in `src/host/capability.ts`, because a build
 * script cannot import TypeScript and importing `dist-lib/` would make writing
 * a manifest depend on a build. `test/manifest.test.ts` holds the two together,
 * which is what the previous arrangement lacked: the code corrected 1.3 to 1.2
 * and both manifests kept 1.3 for weeks, so the file a user installs disagreed
 * with the check that runs.
 */
export const FLOOR = "1.2";

export const PROD_ORIGIN = "https://ssf-merge.struktureretsundfornuft.dk";
export const DEV_ORIGIN = "https://localhost:3000";

const DESCRIPTION =
  "Mail merge for PowerPoint. Mark a block of slides as your template, attach a table, and get one set of slides per row with every font, colour and placement intact.";

export const DEFINITION = {
  id: ID,
  version: VERSION,
  provider: "StruktureretSundFornuft",
  displayName: "SSF Merge",
  shortDescription: "One template block of slides, one set per row of your data.",
  description: DESCRIPTION,
  /**
   * All three were GitHub links, and all three would have failed a submission.
   *
   * Microsoft's submission FAQ is explicit about the support URL: it must be a
   * public web page that does not require authentication, and "you can't use
   * personal social media pages or GitHub repositories", nor "links to files
   * hosted online". That rules out a repository AND a file inside one, which is
   * what all three of these were.
   *
   * The privacy link was the worst of them. It pointed at `SECURITY.md` — a
   * policy for reporting vulnerabilities, not a privacy policy — so a reviewer
   * following it would have read the wrong document entirely and rejected the
   * submission for not having one. AppSource does not certify without a valid
   * privacy link, and missing or invalid links are among the five most common
   * reasons a submission fails.
   *
   * `terms` pointed at the MIT licence, which governs the SOURCE. A licence
   * telling a developer they may fork the repository is not a licence telling a
   * user what they may do with the add-in.
   *
   * GitHub is still linked FROM the support page, which is allowed: what the
   * rule forbids is a repository being the destination.
   */
  support: "https://ssf-merge.struktureretsundfornuft.dk/support.html",
  privacy: "https://ssf-merge.struktureretsundfornuft.dk/privacy.html",
  /**
   * Microsoft's own standard EULA, which they offer to publishers who have no
   * lawyer of their own and have taken the point up with one. Chosen
   * deliberately on 2026-08-30 rather than left as a placeholder.
   */
  terms: "https://support.office.com/client/61994a3b-2c87-41c4-a88d-a6455efa362d",
  /** Navy, the pane's own heading colour. */
  accent: "#00254C",
  button: {
    id: "SsfMerge.OpenPane",
    group: "SsfMerge.Group",
    taskpaneId: "SsfMergePane",
    label: "Mail merge",
    tooltip: "Open SSF Merge to mark a template block and merge your data into it.",
  },
};

/** Where a build points. `origin` is the only thing that differs. */
export function urls(origin) {
  return {
    taskpane: `${origin}/taskpane.html`,
    icon16: `${origin}/assets/icon-16.png`,
    icon32: `${origin}/assets/icon-32.png`,
    icon64: `${origin}/assets/icon-64.png`,
    icon80: `${origin}/assets/icon-80.png`,
    outline32: `${origin}/assets/icon-outline-32.png`,
  };
}

/** XML text, escaped. Every value below reaches an attribute or an element. */
function esc(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * The XML manifest — the one a person sideloads today.
 *
 * There is DELIBERATELY no `<Requirements>` block. The add-in needs
 * PowerPointApi 1.2 — `getFileAsync` is a Common API that PowerPointApi does
 * not gate — and that floor is checked at RUNTIME by `checkFloor` in
 * `src/host/capability.ts`, which is the one place it is stated.
 *
 * This said 1.3 until 2026-08-27, and 1.4 before that. Both were justified by
 * `slide.tags`, which nothing in this add-in calls: merge metadata rides in the
 * package. `capability.ts` corrected the number and the manifests kept the old
 * one, so the file a user installs disagreed with the check that runs. The
 * comment names `API_FLOOR` now and a test holds the two together.
 * Declaring it here would be worse than useless: a host that does not meet a
 * declared requirement set does not show the add-in AT ALL — no ribbon entry,
 * no error, nothing for the user to report — where the runtime check can say
 * which version is missing and what it costs them. `test/manifest.test.ts`
 * holds the absence, because it is the kind of thing somebody adds back as a
 * tidy-up.
 */
export function xmlManifest(origin) {
  const u = urls(origin);
  const d = DEFINITION;
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <Id>${d.id}</Id>
  <Version>${d.version}</Version>
  <ProviderName>${esc(d.provider)}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="${esc(d.displayName)}" />
  <Description DefaultValue="${esc(d.description)}" />
  <IconUrl DefaultValue="${esc(u.icon32)}" />
  <!-- 64, not 80. For a TASK PANE add-in the store's icon must be 32x32 and its
       high-resolution partner 64x64; 128 is the Outlook size and 80 is the
       RIBBON size, which is a different image in a different element below.
       This said 80 and would have been caught by AppSource validation rather
       than by anyone here — the ribbon uses 80 correctly a few lines down, so
       the wrong number looked like the right one. -->
  <HighResolutionIconUrl DefaultValue="${esc(u.icon64)}" />
  <SupportUrl DefaultValue="${esc(d.support)}" />

  <Hosts>
    <Host Name="Presentation" />
  </Hosts>

  <!-- No <Requirements>. The floor is PowerPointApi ${FLOOR} and it is checked at
       runtime by checkFloor, because a DECLARED set the host lacks makes the
       add-in vanish from the ribbon with no diagnostic at all. -->

  <DefaultSettings>
    <SourceLocation DefaultValue="${esc(u.taskpane)}" />
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title" />
            <Description resid="GetStarted.Description" />
            <LearnMoreUrl resid="Support.Url" />
          </GetStarted>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="${d.button.group}">
                <Label resid="Group.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16" />
                  <bt:Image size="32" resid="Icon.32" />
                  <bt:Image size="80" resid="Icon.80" />
                </Icon>
                <Control xsi:type="Button" id="${d.button.id}">
                  <Label resid="OpenPane.Label" />
                  <Supertip>
                    <Title resid="OpenPane.Label" />
                    <Description resid="OpenPane.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>${d.button.taskpaneId}</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${esc(u.icon16)}" />
        <bt:Image id="Icon.32" DefaultValue="${esc(u.icon32)}" />
        <bt:Image id="Icon.80" DefaultValue="${esc(u.icon80)}" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${esc(u.taskpane)}" />
        <bt:Url id="Support.Url" DefaultValue="${esc(d.support)}" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="${esc(d.displayName)}" />
        <bt:String id="OpenPane.Label" DefaultValue="${esc(d.button.label)}" />
        <bt:String id="GetStarted.Title" DefaultValue="${esc(d.displayName)} is ready" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="OpenPane.Tooltip" DefaultValue="${esc(d.button.tooltip)}" />
        <bt:String id="GetStarted.Description" DefaultValue="Find ${esc(d.displayName)} on the Home tab and mark the slides that repeat." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

/**
 * The unified (JSON) manifest, for Microsoft 365 deployment.
 *
 * Same definition, and the same deliberate omission of a declared FLOOR:
 * `extensions[].requirements.capabilities` is where one would go, and it is
 * left out for the reason above.
 *
 * `requirements.scopes` is a different field in the same block and it IS
 * declared. It is the JSON spelling of the XML's `<Hosts><Host
 * Name="Presentation"/></Hosts>` — which host this add-in runs in — and without
 * it the unified manifest, the file a tenant administrator deploys, did not say
 * it was a PowerPoint add-in at all while the XML did. Microsoft's own
 * validator accepts `presentation` and rejects a scope that is not in its enum,
 * which is how this was checked rather than assumed.
 *
 * `authorization.permissions.resourceSpecific` is the JSON spelling of
 * `<Permissions>ReadWriteDocument</Permissions>`, and it is the whole of what
 * this add-in asks for: it reads the deck's bytes and inserts slides.
 */
export function jsonManifest(origin) {
  const u = urls(origin);
  const d = DEFINITION;
  const host = new URL(origin).host;
  return (
    JSON.stringify(
      {
        $schema: "https://developer.microsoft.com/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
        manifestVersion: "1.17",
        id: d.id,
        version: d.version,
        name: { short: d.displayName, full: `${d.displayName} for PowerPoint` },
        description: { short: d.shortDescription, full: d.description },
        developer: {
          name: d.provider,
          websiteUrl: "https://struktureretsundfornuft.dk",
          privacyUrl: d.privacy,
          termsOfUseUrl: d.terms,
        },
        icons: { outline: "assets/icon-outline-32.png", color: "assets/icon-64.png" },
        accentColor: d.accent,
        localizationInfo: { defaultLanguageTag: "en-us" },
        validDomains: [host],
        authorization: {
          permissions: { resourceSpecific: [{ name: "Document.ReadWrite.User", type: "Delegated" }] },
        },
        extensions: [
          {
            // WHICH HOST, and nothing else. `scopes` is the JSON spelling of
            // the XML's <Hosts><Host Name="Presentation"/></Hosts>; the sibling
            // field `capabilities` is where a requirement-set floor would go
            // and is deliberately absent, for the same reason the XML has no
            // <Requirements>: the floor is checked at runtime by `checkFloor`,
            // and a declared one that is wrong makes the add-in vanish from the
            // ribbon with nothing to say why.
            requirements: { scopes: ["presentation"] },
            runtimes: [
              {
                id: "TaskPaneRuntime",
                type: "general",
                code: { page: u.taskpane },
                lifetime: "short",
                actions: [{ id: "ShowTaskpane", type: "openPage", pinnable: false, view: "dashboard" }],
              },
            ],
            ribbons: [
              {
                contexts: ["default"],
                tabs: [
                  {
                    builtInTabId: "TabHome",
                    groups: [
                      {
                        id: d.button.group,
                        label: d.displayName,
                        icons: [
                          { size: 16, url: u.icon16 },
                          { size: 32, url: u.icon32 },
                          { size: 80, url: u.icon80 },
                        ],
                        controls: [
                          {
                            id: d.button.id,
                            type: "button",
                            label: d.button.label,
                            icons: [
                              { size: 16, url: u.icon16 },
                              { size: 32, url: u.icon32 },
                              { size: 80, url: u.icon80 },
                            ],
                            supertip: { title: d.button.label, description: d.button.tooltip },
                            actionId: "ShowTaskpane",
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
      null,
      2,
    ) + "\n"
  );
}

/** Every file a build writes, by name. One place, so nothing is half-generated. */
export function allManifests() {
  return {
    "manifest.xml": xmlManifest(DEV_ORIGIN),
    "manifest-prod.xml": xmlManifest(PROD_ORIGIN),
    "manifest.json": jsonManifest(DEV_ORIGIN),
    "manifest-prod.json": jsonManifest(PROD_ORIGIN),
  };
}
