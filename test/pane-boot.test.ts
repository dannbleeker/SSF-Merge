import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * The pane's entry point, as the browser sees it.
 *
 * Every other test here reaches past this file. `pane-render.test.ts` calls
 * `render` in jsdom; `scripts/pane-shots.mjs` imports `render.ts` directly in a
 * real browser and photographs the result. Both are useful and neither loads
 * `main.ts`, so for the whole life of this repo the pane's actual entry point
 * had never run anywhere — and it could not have, because the page never loaded
 * Office.js. `Office.onReady` was a `ReferenceError` on the first statement, the
 * pane painted its header and stopped, and fourteen perfect screenshots said
 * nothing was wrong.
 *
 * The lesson generalises past this bug: a harness that mounts a component
 * bypasses the document that mounts it in production, so the document needs a
 * guard of its own.
 */
const html = readFileSync("src/pane/taskpane.html", "utf8");

describe("the pane loads Office.js", () => {
  it("references Microsoft's documented CDN path", () => {
    // The only supported way to load it. A vendored copy is not serviceable and
    // Microsoft's guidance is against it.
    expect(html).toContain("https://appsforoffice.microsoft.com/lib/1/hosted/office.js");
  });

  it("loads it before the module that uses it", () => {
    // `main.ts` is type="module" and therefore deferred, so a BLOCKING script
    // tag is what guarantees `Office` exists by the time it runs.
    const office = html.indexOf("appsforoffice.microsoft.com");
    const entry = html.indexOf('src="./main.ts"');
    expect(office).toBeGreaterThan(-1);
    expect(entry).toBeGreaterThan(-1);
    expect(office).toBeLessThan(entry);
  });

  it("does not mark it async or defer", () => {
    // Either one turns the ordering above into a race the pane loses
    // intermittently — the worst version of this bug, because it works on the
    // machine you test it on.
    const tag = /<script[^>]*appsforoffice[^>]*>/.exec(html)?.[0] ?? "";
    // Asserted, so this cannot pass vacuously on a file with no tag at all —
    // which is exactly the state the revert that proves these tests puts it in.
    expect(tag, "no Office.js script tag to check").not.toBe("");
    expect(tag).not.toContain("async");
    expect(tag).not.toContain("defer");
  });
});

describe("the pane says so when Office.js did not arrive", () => {
  /**
   * A blocked CDN, an offline laptop or a proxy rule all end as a blank pane
   * with no diagnostic — nothing to report and nothing to search for. This
   * environment reproduces it exactly (its egress proxy refuses the CDN), which
   * is how the fallback was checked against a real failure rather than a mocked
   * one.
   */
  it("renders a message naming the address that failed", () => {
    expect(html).toContain("appsforoffice.microsoft.com), so it cannot talk to PowerPoint");
  });

  it("paints only when Office is absent AND nothing else rendered", () => {
    // Conditioned on both facts. On the timer alone it would paint over a pane
    // `main.ts` had already drawn.
    expect(html).toContain('typeof Office !== "undefined"');
    expect(html).toContain('pane.innerHTML !== ""');
  });
});
