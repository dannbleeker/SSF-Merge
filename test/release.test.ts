import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tools with no types. The rules live THERE so the
// workflow and this test cannot read different ones.
import { RELEASE_ASSETS, assetsPromisedByDocs, releaseProblems } from "../scripts/release-assets.mjs";
// @ts-expect-error — as above.
import { withoutHashComments } from "../scripts/without-prose.mjs";

/**
 * What a release ships.
 *
 * Every other gate in this repo reads the WORKING TREE; a user downloads the
 * RELEASE. A sibling project had those diverge twice — once shipping the dev
 * manifests, once with the README pointing at a `manifest-prod.xml` that was
 * not in the release at all, for twelve days, with a correct release workflow
 * sitting un-run. Both directions are checked here.
 */
const read = (name: string) => readFileSync(name, "utf8");
const assets = RELEASE_ASSETS as string[];

describe("the release as it stands", () => {
  it("has nothing wrong with it", () => {
    expect(releaseProblems(read)).toEqual([]);
  });

  it("ships the manifests and nothing else, because the pane is hosted", () => {
    // A release of this repo is small on purpose: the only thing a user
    // downloads is a manifest. Anything else here would be a second copy of
    // something Pages already serves.
    expect(assets).toEqual(["manifest-prod.xml", "manifest-prod.json"]);
  });
});

describe("the documentation and the release cannot disagree", () => {
  it("attaches every production manifest the docs tell people to download", () => {
    // Read out of the PROSE, not from a list — the failure being guarded
    // against is exactly a second list that drifts from the first.
    const promised = (assetsPromisedByDocs() as string[]).filter((n) => n.includes("-prod"));
    expect(promised.length).toBeGreaterThan(0);
    for (const name of promised) expect(assets, `${name} is promised by the docs`).toContain(name);
  });

  it("catches the docs promising something the release does not carry", () => {
    // The twelve-day failure, in one line.
    expect(releaseProblems(read, ["manifest-prod.xml"], ["manifest-prod.xml", "manifest-prod.json"])).toEqual([
      expect.stringContaining("does not attach it"),
    ]);
  });

  it("does not demand a DEV manifest just because the docs mention one", () => {
    // `manifest.xml` appears in contributor documentation. It is for people
    // running the pane locally, and shipping it would point every installer at
    // a localhost port nothing is listening on.
    expect(releaseProblems(read, assets, ["manifest.xml", "manifest-prod.xml", "manifest-prod.json"])).toEqual([]);
  });
});

describe("what a release must never ship", () => {
  it("refuses a development manifest outright", () => {
    expect(releaseProblems(read, ["manifest.xml"], [])).toEqual([expect.stringContaining("not a production manifest")]);
  });

  it("refuses a production manifest that points at localhost", () => {
    // A release can be perfectly current and still be built from an origin that
    // stopped being the production one.
    const localhost = (name: string) =>
      read(name).replaceAll("https://ssf-merge.struktureretsundfornuft.dk", "https://localhost:3000");
    expect(releaseProblems(localhost, ["manifest-prod.xml"], [])).toEqual([expect.stringContaining("localhost")]);
  });

  it("refuses a manifest Office would reject", () => {
    // Every rule the offline checker holds applies to the file being shipped,
    // not only to the one in the tree.
    const broken = (name: string) => read(name).replace("<Version>1.0.0.0</Version>", "<Version>0.1.0</Version>");
    expect(releaseProblems(broken, ["manifest-prod.xml"], [])).toEqual([expect.stringContaining("below 1.0")]);
  });

  it("refuses an asset that is not there at all", () => {
    const missing = () => {
      throw new Error("no such file");
    };
    expect(releaseProblems(missing, ["manifest-prod.xml"], [])).toEqual([expect.stringContaining("is not there")]);
  });
});

describe("the release workflow", () => {
  // Stripped of its comments: the header explains why the tag is created by
  // `gh release create`, and the first version of these tests matched that
  // sentence instead of the step — so the ORDER check compared the comment's
  // position with the check's and reported it backwards.
  const workflow = withoutHashComments(read(".github/workflows/release.yml")) as string;

  it("is manual only — a release is a decision, not a consequence of merging", () => {
    expect(workflow).toContain("workflow_dispatch");
    expect(workflow).not.toMatch(/^\s+push:/m);
  });

  it("runs the pre-flight before it creates anything", () => {
    // The order is the guarantee. A tag created before the checks is a release
    // that has to be yanked rather than refused.
    const check = workflow.indexOf("scripts/check-release.mjs");
    const create = workflow.indexOf("gh release create");
    expect(check).toBeGreaterThan(-1);
    expect(create).toBeGreaterThan(check);
  });

  it("regenerates the manifests and fails if the tree was stale", () => {
    expect(workflow).toContain("scripts/build-manifests.mjs");
    expect(workflow).toContain("git diff --exit-code");
  });

  it("validates with Microsoft's own tool, on EVERY file it is shipping", () => {
    // Both, not just the XML. The release attaches `manifest-prod.json` too and
    // the note points administrators deploying to a whole tenant at it, so the
    // highest-consequence artifact here was the one no external authority read.
    expect(workflow).toContain("office-addin-manifest validate");
    for (const asset of RELEASE_ASSETS) {
      expect(workflow, `${asset} is shipped and not validated`).toMatch(
        new RegExp(`for m in[^\\n]*${asset.replace(".", "\\.")}`),
      );
    }
  });

  it("fetches the URLs the manifest points at, so a release cannot ship a dead host", () => {
    // `checkManifest` asks whether a production manifest points at localhost,
    // which is a different question and passes cleanly for one pointing at a
    // domain that 404s — an add-in that installs perfectly and shows a blank
    // ribbon button and an empty pane.
    expect(workflow).toContain("scripts/manifest-urls.mjs");
    expect(workflow).toContain("curl --fail");
    // Into a variable, never `… | while read`: a loop on the right of a pipe
    // runs in a subshell, and a curl that fails inside one is a step that goes
    // green while the host is down.
    expect(workflow, "a piped loop swallows the failure").not.toMatch(/manifest-urls\.mjs[^\n]*\|\s*while/);
    // And an empty list is a failure rather than a silent pass.
    expect(workflow).toContain('test -n "$urls"');
  });

  it("keeps the URL check OUT of CI, where a third-party outage would block merges", () => {
    // The same reasoning that already keeps Microsoft's validator out of the
    // `test` job. A gate that fails for somebody else's uptime is a gate that
    // gets switched off after the first bad week.
    expect(readFileSync(".github/workflows/ci.yml", "utf8")).not.toContain("manifest-urls.mjs");
  });

  it("attaches exactly the assets the rules name", () => {
    // The workflow's own upload list, against the list every check above used.
    const upload = workflow.slice(workflow.indexOf("gh release create"));
    expect(upload.length, "the create step is in the file").toBeGreaterThan(20);
    for (const name of assets) expect(upload, `${name} is uploaded`).toContain(name);
  });
});

describe("the deploy waits for the same gate CI runs", () => {
  /**
   * Pages ran on every push to `main` with nothing between the push and the
   * live add-in. CI ran too, but CONCURRENTLY — so a commit could be serving
   * from the production origin before its tests had finished, and if they then
   * failed, the broken pane was already what PowerPoint loaded.
   *
   * Every commit on main has in fact been green on both, checked over the run
   * history rather than assumed. It is the ORDERING that was missing, not the
   * tests.
   *
   * The five checks are listed in both workflows rather than wrapped in one
   * `npm run gate`, because a script chaining `npm run a && npm run b` cannot
   * run on the maintainer's own machine: npm spawns a shell for the `&&` and
   * AppLocker refuses it. Two lists is a drift risk, so this holds them against
   * each other.
   */
  /** The `npm run` commands of ONE job, bounded by the next job at the same indent. */
  const commands = (yaml: string, job: string): string[] => {
    const lines = yaml.split("\n");
    const at = lines.findIndex((l) => l === `  ${job}:`);
    expect(at, `${job} is not a job in this workflow`).toBeGreaterThan(-1);
    // The next sibling job, or the end. Without this the deploy job's own
    // `npm run build` reads as one of the gate's checks, which is how the
    // first version of this test failed against a workflow that was correct.
    const rest = lines.slice(at + 1);
    const until = rest.findIndex((l) => /^ {2}[A-Za-z0-9_-]+:$/.test(l));
    const body = (until === -1 ? rest : rest.slice(0, until)).join("\n");
    // EVERY `run:` command, not only the `npm run` ones. Narrowed to those, a
    // step that is not an npm script — the check that the test-count floor was
    // actually committed — could be in one gate and not the other and this
    // would say they agreed.
    return [...body.matchAll(/run: (.+)/g)].map((m) => (m[1] as string).trim());
  };

  it("runs the same checks, in the same order, in both", () => {
    const ci = readFileSync(".github/workflows/ci.yml", "utf8");
    const pages = readFileSync(".github/workflows/pages.yml", "utf8");

    const inCi = commands(ci, "test");
    expect(inCi.length, "the CI gate stopped naming any checks").toBeGreaterThan(3);
    expect(commands(pages, "gate"), "the deploy gate and CI have drifted").toEqual(inCi);
  });

  it("and the deploy will not start without it", () => {
    // The ordering itself. Without `needs`, both jobs start together and the
    // artifact can be live before the checks finish.
    const pages = readFileSync(".github/workflows/pages.yml", "utf8");
    expect(pages, "deploy does not wait for the gate").toMatch(/deploy:\s*\n\s*needs: gate/);
  });
});
