import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import type { Browser, BrowserType, LaunchOptions } from "playwright";

// @ts-expect-error — a plain .mjs tool with no types. Narrowed here so every
// call below is checked against the shape the scripts rely on.
import { launchChromium as untypedLaunch, newestChromium as untypedNewest } from "../scripts/launch-chromium.mjs";

type Found = { path: string; revision: number } | undefined;
const newestChromium = untypedNewest as (dir: string) => Found;
const launchChromium = untypedLaunch as (
  chromium: Pick<BrowserType, "launch">,
  options?: LaunchOptions,
  env?: Record<string, string | undefined>,
) => Promise<{ browser: Browser; source: string }>;

/**
 * How the screenshot scripts pick a browser.
 *
 * Both used to name `chromium-1194` outright. By 2026-10-26 the container held
 * 1243 and 1248 and no 1194, so one script quietly skipped its pin and the
 * other could not launch. These tests hold the replacement to two promises:
 * it never needs a revision written down, and it never uses a fallback without
 * saying so.
 */

/** A browsers directory with these entries, each a layout holding a binary or "empty". */
function browsers(entries: Record<string, "linux64" | "linux" | "empty">): string {
  const root = mkdtempSync(join(tmpdir(), "pw-browsers-"));
  for (const [name, layout] of Object.entries(entries)) {
    if (layout === "empty") {
      mkdirSync(join(root, name));
      continue;
    }
    const at = join(root, name, `chrome-${layout}`);
    mkdirSync(at, { recursive: true });
    writeFileSync(join(at, "chrome"), "");
  }
  return root;
}

const fakeBrowser = { version: () => "0.0" } as unknown as Browser;
const MISSING = new Error("browserType.launch: Executable doesn't exist at /nowhere/chrome-headless-shell");

/** A `chromium` whose own download behaves as `own` says, recording each launch. */
function fakeChromium(own: "present" | "missing" | Error) {
  const calls: LaunchOptions[] = [];
  const chromium: Pick<BrowserType, "launch"> = {
    launch: (options: LaunchOptions = {}) => {
      calls.push(options);
      if (options.executablePath || own === "present") return Promise.resolve(fakeBrowser);
      return Promise.reject(own === "missing" ? MISSING : own);
    },
  };
  return { chromium, calls };
}

describe("the newest Chromium on disk", () => {
  it("compares revisions as numbers, not as strings", () => {
    // As strings "999" sorts after "1248", which would pick the oldest.
    const root = browsers({ "chromium-999": "linux", "chromium-1248": "linux64", "chromium-1243": "linux64" });
    expect(newestChromium(root)).toEqual({
      path: join(root, "chromium-1248", "chrome-linux64", "chrome"),
      revision: 1248,
    });
  });

  it("reads the older chrome-linux layout too", () => {
    const root = browsers({ "chromium-1194": "linux" });
    expect(newestChromium(root)?.path).toBe(join(root, "chromium-1194", "chrome-linux", "chrome"));
  });

  it("skips a directory with no binary in it rather than choosing a name", () => {
    const root = browsers({ "chromium-1300": "empty", "chromium-1243": "linux64" });
    expect(newestChromium(root)?.revision).toBe(1243);
  });

  it("ignores the headless shell and the bare symlink beside the real builds", () => {
    const root = browsers({ "chromium_headless_shell-1300": "linux64", "chromium-1243": "linux64" });
    expect(newestChromium(root)?.revision).toBe(1243);
  });

  it("answers nothing for a directory that holds no browser or does not exist", () => {
    expect(newestChromium(browsers({}))).toBeUndefined();
    expect(newestChromium(join(tmpdir(), "no-such-dir-for-launch-chromium"))).toBeUndefined();
  });
});

describe("launching", () => {
  const root = browsers({ "chromium-1243": "linux64" });

  it("uses playwright's own download when it is there, and says so", async () => {
    const { chromium, calls } = fakeChromium("present");
    const { source } = await launchChromium(chromium, {}, { PLAYWRIGHT_BROWSERS_PATH: root });
    expect(source).toBe("playwright's own download");
    expect(calls).toEqual([{}]);
  });

  it("falls back to what is on disk only when the download is missing, and names it a fallback", async () => {
    const { chromium, calls } = fakeChromium("missing");
    const { source } = await launchChromium(chromium, {}, { PLAYWRIGHT_BROWSERS_PATH: root });
    expect(calls.at(-1)?.executablePath).toBe(join(root, "chromium-1243", "chrome-linux64", "chrome"));
    expect(source).toMatch(/^FALLBACK chromium-1243/);
    expect(source).toContain("npx playwright install chromium");
  });

  it("lets CHROMIUM win over everything", async () => {
    const { chromium, calls } = fakeChromium("present");
    const { source } = await launchChromium(chromium, {}, { CHROMIUM: "/my/chrome", PLAYWRIGHT_BROWSERS_PATH: root });
    expect(calls).toEqual([{ executablePath: "/my/chrome" }]);
    expect(source).toBe("/my/chrome (CHROMIUM)");
  });

  it("rethrows any other launch failure instead of hiding it behind another browser", async () => {
    const crash = new Error("Target page, context or browser has been closed");
    const { chromium, calls } = fakeChromium(crash);
    await expect(launchChromium(chromium, {}, { PLAYWRIGHT_BROWSERS_PATH: root })).rejects.toBe(crash);
    expect(calls).toHaveLength(1);
  });

  it("rethrows the missing-download error when there is nothing to fall back to", async () => {
    const { chromium } = fakeChromium("missing");
    await expect(launchChromium(chromium, {}, { PLAYWRIGHT_BROWSERS_PATH: browsers({}) })).rejects.toBe(MISSING);
  });
});
