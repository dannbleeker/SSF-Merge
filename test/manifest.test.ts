import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { DOMParser } from "@xmldom/xmldom";
import { API_FLOOR } from "../src/host/capability.js";
// @ts-expect-error — plain .mjs tools with no types. The definition lives THERE
// so the generator and this test cannot read different sources.
import { DEFINITION, FLOOR, allManifests, urls, PROD_ORIGIN, DEV_ORIGIN } from "../scripts/manifest-source.mjs";
// @ts-expect-error — as above.
import { checkManifest, urlsIn, REQUIRED_ID } from "../scripts/manifest-rules.mjs";
// @ts-expect-error — as above.
import { fetchable as fetchableJs } from "../scripts/manifest-urls.mjs";

// `.mjs`, so the suite sees `any` across the boundary. Named once here rather
// than asserted at every call site, which is where an unsafe-return lint lands.
const fetchable = fetchableJs as (urls: string[]) => string[];

/**
 * The manifests, checked offline.
 *
 * CI runs Microsoft's own `office-addin-manifest validate` in a job of its own
 * and that is the authority. This is not a second copy of it — the validator
 * calls a Microsoft SERVICE, which answers 403 through this project's egress
 * proxy, so in the environment the code is actually written in it cannot run at
 * all. These pin the small number of rules whose violation costs something
 * specific, and the staleness gate that keeps four generated files honest.
 */
const NAMES = ["manifest.xml", "manifest-prod.xml", "manifest.json", "manifest-prod.json"];
const read = (name: string) => readFileSync(name, "utf8");
const generated = allManifests() as Record<string, string>;

describe("the committed manifests are the generated ones", () => {
  // They are committed because a manifest is a file a PERSON sideloads: it has
  // to be downloadable from the repo at a stable path with no toolchain. That
  // is only safe while "committed" and "generated" cannot come apart.
  it.each(NAMES)("%s matches what manifest-source.mjs produces", (name) => {
    expect(read(name)).toBe(generated[name]);
  });

  it("writes all four, so none can be quietly dropped", () => {
    expect(Object.keys(generated).sort()).toEqual([...NAMES].sort());
  });
});

describe("the rules this project would be bitten by", () => {
  it.each(NAMES)("%s breaks none of them", (name) => {
    expect(checkManifest(read(name), name)).toEqual([]);
  });

  it("would still catch each of them being broken", () => {
    // The rules are worth importing only if they can still fail. Every branch
    // below is a thing shipped somewhere or one edit from being.
    const xml = read("manifest-prod.xml");
    const json = read("manifest-prod.json");

    // Office rejects a version below 1.0 outright — "Manifest Version Too Low"
    // — and a sibling project shipped 0.1.0 in four manifests for months with
    // a fully green suite, because nothing had ever asked Microsoft.
    expect(
      checkManifest(xml.replace("<Version>1.0.0.0</Version>", "<Version>0.1.0</Version>"), "manifest-prod.xml"),
    ).toEqual([expect.stringContaining("below 1.0")]);

    // A changed GUID is a different add-in: every sideload orphaned, with
    // nothing anywhere saying why.
    expect(
      checkManifest(xml.replace(REQUIRED_ID, "00000000-0000-0000-0000-000000000000"), "manifest-prod.xml"),
    ).toEqual([expect.stringContaining("orphaned")]);
    expect(
      checkManifest(json.replace(REQUIRED_ID, "00000000-0000-0000-0000-000000000000"), "manifest-prod.json"),
    ).toEqual([expect.stringContaining("orphaned")]);

    // A prod manifest can be perfectly CURRENT and full of localhost, if the
    // origin the generator is handed ever stops being the production one.
    expect(checkManifest(xml.replaceAll(PROD_ORIGIN, DEV_ORIGIN), "manifest-prod.xml")).toEqual([
      expect.stringContaining("localhost"),
    ]);
    // And the other direction, which wastes an afternoon rather than a release:
    // a dev manifest pointing at production tests every edit against the last
    // deploy.
    expect(checkManifest(read("manifest.xml").replaceAll(DEV_ORIGIN, PROD_ORIGIN), "manifest.xml")).toEqual([
      expect.stringContaining("no local origin"),
    ]);

    expect(checkManifest("<html>not a manifest</html>", "manifest-prod.xml")).toEqual([
      expect.stringContaining("not an Office add-in manifest"),
    ]);
    expect(checkManifest("{ not json", "manifest-prod.json")).toEqual([expect.stringContaining("not valid JSON")]);

    // Inserting slides needs it, and an add-in that does not ask is one that
    // fails at the one call it exists to make.
    expect(checkManifest(xml.replace("<Permissions>ReadWriteDocument</Permissions>", ""), "manifest-prod.xml")).toEqual(
      [expect.stringContaining("ReadWriteDocument")],
    );
  });
});

describe("the requirement floor is checked at runtime, never declared", () => {
  // The load-bearing omission, and the kind of thing somebody adds back as a
  // tidy-up: "surely the manifest should say which API it needs". A host that
  // does not meet a DECLARED requirement set does not show the add-in at all —
  // no ribbon entry, no error, nothing for the user to report. checkFloor can
  // say which version is missing and what it costs them.
  it("names a floor in the code", () => {
    // 1.2, and corrected from 1.3: that was justified by `slide.tags`, which
    // nothing in this add-in calls. See `capability.ts`.
    expect(API_FLOOR).toBe("1.2");
  });

  it("states the SAME floor in the manifests as the code checks", () => {
    // Two spellings, held together. The code corrected 1.3 to 1.2 — the 1.3 was
    // justified by `slide.tags`, which nothing here calls — and both manifests
    // kept 1.3, so the file a user installs disagreed with the check that runs.
    // Nothing was watching, because the comment is prose to every other rule.
    expect(FLOOR).toBe(API_FLOOR);
    for (const name of NAMES.filter((n) => n.endsWith(".xml"))) {
      expect(read(name), name).toContain(`PowerPointApi ${API_FLOOR}`);
    }
  });

  it("declares no requirement set in any manifest", () => {
    for (const name of NAMES) {
      const text = read(name);
      // The XML manifest EXPLAINS the omission in a comment holding the word,
      // so the rule reads markup rather than prose — checked here through the
      // rule itself rather than by a second regex that could disagree with it.
      expect(checkManifest(text, name), name).toEqual([]);
      if (name.endsWith(".json")) {
        for (const extension of JSON.parse(text).extensions) {
          expect(extension.requirements, name).toBeUndefined();
        }
      }
    }
  });

  it("catches a production manifest served over http, in either format", () => {
    /**
     * Office fetches every address a manifest names and requires HTTPS for all
     * of them. A production manifest on `http://` fails Microsoft's validator,
     * and sideloaded anyway it fails the way the requirement-set rule above
     * describes: no ribbon entry, no error, nothing to report.
     *
     * `PROD_ORIGIN` is one constant, which is this rule set's own test for
     * whether a rule earns its place: one edit away from being shipped.
     */
    const xml = read("manifest-prod.xml").replace(/https:\/\/ssf-merge/g, "http://ssf-merge");
    expect(checkManifest(xml, "manifest-prod.xml")).toEqual([expect.stringContaining("insecure address")]);

    const json = read("manifest-prod.json").replace(/https:\/\/ssf-merge/g, "http://ssf-merge");
    expect(checkManifest(json, "manifest-prod.json")).toEqual([expect.stringContaining("insecure address")]);
  });

  it("does not mistake a namespace for an address", () => {
    /**
     * The whole difficulty of the rule above. `xmlns="http://schemas.microsoft
     * .com/..."` is an IDENTIFIER: never fetched, http by definition, and not
     * ours to change. A rule that read those would fire on every manifest ever
     * written and would have been deleted rather than fixed.
     *
     * Asserted from the real file, so it holds against whatever namespaces the
     * generator emits rather than against a fixture written to pass.
     */
    const text = read("manifest-prod.xml");
    expect(text, "the fixture stopped carrying an http namespace").toContain('xmlns="http://');
    expect(checkManifest(text, "manifest-prod.xml")).toEqual([]);
  });

  it("catches a requirement set being added back, in either format", () => {
    const xml = read("manifest-prod.xml").replace(
      "<Hosts>",
      '<Requirements><Sets><Set Name="PowerPointApi" MinVersion="1.3" /></Sets></Requirements><Hosts>',
    );
    expect(checkManifest(xml, "manifest-prod.xml")).toEqual([expect.stringContaining("checkFloor")]);

    const doc = JSON.parse(read("manifest-prod.json"));
    doc.extensions[0].requirements = { capabilities: [{ name: "PowerPointApi", minVersion: "1.3" }] };
    expect(checkManifest(JSON.stringify(doc), "manifest-prod.json")).toEqual([expect.stringContaining("checkFloor")]);
  });
});

describe("the XML manifests are well-formed", () => {
  // Not a schema check — that is the validator's job — but a manifest that
  // does not PARSE fails in PowerPoint with a dialog naming a line number, and
  // the generator interpolates strings a person will one day edit.
  it.each(["manifest.xml", "manifest-prod.xml"])("%s parses", (name) => {
    const problems: string[] = [];
    const doc = new DOMParser({
      onError: (level: string, message: string) => {
        if (level !== "warning") problems.push(message);
      },
    }).parseFromString(read(name), "text/xml");
    expect(problems, name).toEqual([]);
    expect(doc.documentElement?.nodeName).toBe("OfficeApp");
  });

  it("escapes a value that would otherwise break the markup", async () => {
    // Every field below reaches an attribute. `&` and `"` are the two that end
    // a document, and a description is exactly the kind of prose that grows an
    // ampersand.
    // @ts-expect-error — the .mjs source again, imported lazily here because
    // this is the one case that wants a manifest for an origin no build makes.
    const { xmlManifest } = await import("../scripts/manifest-source.mjs");
    const text = xmlManifest("https://example.test/a?x=1&y=2") as string;
    expect(text).toContain("&amp;y=2");
    expect(text).not.toMatch(/DefaultValue="[^"]*&(?!amp;|lt;|gt;|quot;)/);
  });
});

describe("every icon a manifest names is a file this repo builds", () => {
  it("has each one under public/assets", () => {
    // A manifest pointing at an icon nobody generated is a blank square in the
    // ribbon, and the validator does not fetch them.
    const named = new Set(
      NAMES.flatMap((n) => urlsIn(read(n)) as string[])
        .filter((u) => u.includes("/assets/"))
        .map((u) => u.slice(u.indexOf("/assets/") + 1)),
    );
    // Plus the two the JSON manifest names as package-relative paths.
    for (const relative of ["assets/icon-outline-32.png", "assets/icon-64.png"]) named.add(relative);
    expect(named.size).toBeGreaterThan(3);
    for (const relative of named) {
      expect(existsSync(`public/${relative}`), `public/${relative} is missing`).toBe(true);
    }
  });
});

describe("the manifest version is not the package version", () => {
  it("is four parts and at least 1.0", () => {
    // Office wants a.b.c.d and rejects below 1.0; npm wants semver and this
    // package is 0.0.0. Tying them together would have shipped an invalid
    // manifest on the first release.
    expect(DEFINITION.version).toMatch(/^\d+\.\d+\.\d+\.\d+$/);
    expect(Number(DEFINITION.version.split(".")[0])).toBeGreaterThanOrEqual(1);
    const pkg = JSON.parse(read("package.json")) as { version: string };
    expect(DEFINITION.version).not.toBe(pkg.version);
  });

  it("pins the GUID as a literal, so a regenerate cannot change it", () => {
    expect(DEFINITION.id).toBe("43ebbbac-44ad-42b2-a582-0ef079093e6c");
    expect(REQUIRED_ID).toBe(DEFINITION.id);
  });
});

describe("dev and prod differ in the origin and nothing else", () => {
  it("is the same manifest with one substitution", () => {
    // If they ever differ in more than the origin, one of them is being edited
    // by hand — which is the whole thing generating them prevents.
    expect(read("manifest.xml").replaceAll(DEV_ORIGIN, PROD_ORIGIN)).toBe(read("manifest-prod.xml"));
  });

  it("points the pane and the icons at that origin", () => {
    const u = urls(PROD_ORIGIN) as Record<string, string>;
    for (const value of Object.values(u)) expect(value.startsWith(PROD_ORIGIN)).toBe(true);
    expect(read("manifest-prod.xml")).toContain(u.taskpane);
  });
});

describe("the icons the manifests point at", () => {
  it("are byte-identical to what scripts/build-icons.mjs draws", async () => {
    // Committed because Pages serves them and a manifest names them by URL, and
    // gated because a committed binary is the one kind of file a reviewer
    // cannot read. If this fails, either the drawing changed and the PNGs were
    // not rebuilt, or a PNG was edited by hand.
    // @ts-expect-error — plain .mjs with no types.
    const { SIZES, png, markPixel } = await import("../scripts/build-icons.mjs");
    for (const size of SIZES as number[]) {
      const drawn = png(size, markPixel(size)) as Buffer;
      expect(Buffer.compare(drawn, readFileSync(`public/assets/icon-${size}.png`)), `icon-${size}.png`).toBe(0);
    }
  });

  it("are real PNGs of the size they claim", () => {
    // A truncated or zero-byte file still satisfies existsSync, and a blank
    // square in the ribbon is what a user sees.
    for (const size of [16, 32, 64, 80]) {
      const bytes = readFileSync(`public/assets/icon-${size}.png`);
      expect(bytes.subarray(0, 8).toString("hex"), `icon-${size}.png signature`).toBe("89504e470d0a1a0a");
      expect(bytes.readUInt32BE(16), `icon-${size}.png width`).toBe(size);
      expect(bytes.readUInt32BE(20), `icon-${size}.png height`).toBe(size);
    }
  });

  it("draws the same picture at every size — one orange row and two copies", async () => {
    // Asserted on `markPixel`, the pure function, rather than by decoding the
    // PNG: test one above already pins the file to exactly what this draws, so
    // checking the drawing is checking the file. A first version asserted the
    // file was longer than sixty bytes, which is true of any PNG at all.
    //
    // Everything in the mark is in SIXTEENTHS, so the proportions must not
    // drift between the ribbon's smallest icon and its largest.
    // @ts-expect-error — plain .mjs with no types.
    const { markPixel, NAVY, ORANGE } = await import("../scripts/build-icons.mjs");
    for (const size of [16, 32, 64, 80]) {
      const at = markPixel(size) as (x: number, y: number) => number[];
      const middle = Math.round(size / 2);
      const column = Array.from({ length: size }, (_, y) => at(middle, y));
      const kind = (px: number[]) =>
        px[3] !== 255 ? "clear" : px[0] === ORANGE[0] ? "orange" : px[0] === NAVY[0] ? "navy" : "pale";

      // Down the middle: ground, then orange, then two pale rows, ground
      // between each. Collapsed to the sequence of colours it passes through.
      const runs = column.map(kind).filter((k, i, all) => k !== all[i - 1]);
      expect(runs, `icon-${size}.png down the middle`).toEqual([
        "navy",
        "orange",
        "navy",
        "pale",
        "navy",
        "pale",
        "navy",
      ]);
      // The rows are inset, so the mark has a ground on both sides at every
      // size — the thing that stops it reading as a full-width dash.
      expect(kind(at(0, Math.round(size * 0.3))), `icon-${size}.png left margin`).not.toBe("orange");
    }
  });
});

describe("the URLs a release checks before it ships", () => {
  /**
   * Nothing else in this repo asks whether the host a manifest names is
   * SERVING. `checkManifest` asks whether a production manifest points at
   * localhost, which is a different question and passes cleanly for a manifest
   * pointing at a domain that 404s — so a release could tell people to sideload
   * a file that installs perfectly and shows a blank ribbon button and an empty
   * pane, with nothing anywhere saying why.
   *
   * The fetching is in the release workflow, because a network call is not
   * testable here. WHICH urls get fetched is the decision, and it is this.
   */
  it("takes the ones this project serves, and no more", () => {
    const urls = fetchable([
      "https://ssf-merge.example.dk/taskpane.html",
      "https://ssf-merge.example.dk/assets/icon-32.png",
      "https://github.com/dannbleeker/SSF-Merge",
      "https://learn.microsoft.com/office/dev/add-ins/",
    ]);
    // Same origin as the pane. A release must not fail because github.com rate
    // limited a runner or a documentation page moved — that is somebody else's
    // uptime, and this step would be switched off after the first bad week.
    expect(urls).toEqual([
      "https://ssf-merge.example.dk/assets/icon-32.png",
      "https://ssf-merge.example.dk/taskpane.html",
    ]);
  });

  it("answers NOTHING when it cannot tell which origin is ours", () => {
    // Rather than guessing at the first URL it sees. The workflow fails on an
    // empty list, so a manifest whose shape this no longer understands stops a
    // release instead of silently checking nothing.
    expect(fetchable(["https://github.com/dannbleeker/SSF-Merge"])).toEqual([]);
  });

  it("finds every URL the real production manifests serve", () => {
    // Against the committed files, so a renamed asset or a moved origin shows
    // up here rather than in somebody's ribbon.
    const urls = fetchable([...urlsIn(read("manifest-prod.xml")), ...urlsIn(read("manifest-prod.json"))]);
    expect(urls.length, "the release would check no URLs at all").toBeGreaterThan(0);
    expect(
      urls.some((u) => u.endsWith("/taskpane.html")),
      "the pane itself is not checked",
    ).toBe(true);
    expect(
      urls.some((u) => u.includes("icon")),
      "no icon is checked",
    ).toBe(true);
    expect(
      urls.every((u) => u.startsWith("https://")),
      "an insecure URL would be shipped",
    ).toBe(true);
  });
});
