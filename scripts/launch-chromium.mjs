/**
 * One way for the screenshot scripts to get a Chromium, and a sentence saying
 * which one they got.
 *
 * Both callers used to name `/opt/pw-browsers/chromium-1194/...` outright,
 * because the development container's bundled browser and the installed
 * playwright disagreed on revision. A pinned revision goes dead the first time
 * either side moves, and both did: on 2026-10-26 that container held 1243 and
 * 1248 and no 1194 at all, playwright wanted 1248, `pane-shots.mjs` skipped its
 * pin only because it checked for the file first, and `build-aspect-probe.mjs`
 * did not check and could not launch at all. So no revision is written down
 * here. The fallback DISCOVERS what is on disk instead.
 *
 * The order:
 *
 * 1. `CHROMIUM`, when set. A person naming a binary wins.
 * 2. playwright's own download, the one its version expects. That is what CI
 *    uses and what `npx playwright install chromium` provides.
 * 3. Only when that binary is MISSING, the newest `chromium-<rev>` under
 *    `PLAYWRIGHT_BROWSERS_PATH`. A sweep one browser behind is still worth
 *    having. One that is behind without saying so is not, so `source` names
 *    the revision and says it is a fallback.
 *
 * Any other launch failure is rethrown as it is. A crash is not a missing
 * download, and hiding it behind a different browser would be a fallback
 * reading recorded as a measurement.
 */
import { existsSync, readdirSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

/** Where playwright keeps browsers when nobody has said otherwise, on Linux. */
const DEFAULT_BROWSERS = join(homedir(), ".cache", "ms-playwright");

/**
 * The newest full Chromium under `dir`, or undefined.
 *
 * Revisions are compared as NUMBERS: as strings "999" sorts after "1243".
 * Both layouts are accepted, `chrome-linux64/` (current) and `chrome-linux/`
 * (older builds), and a directory holding neither is skipped rather than
 * chosen. A name is not a browser.
 *
 * @param {string} dir
 * @returns {{ path: string, revision: number } | undefined}
 */
export function newestChromium(dir) {
  let names;
  try {
    names = readdirSync(dir);
  } catch {
    return undefined;
  }
  return names
    .map((name) => /^chromium-(\d+)$/.exec(name))
    .filter((m) => m !== null)
    .map((m) => ({ revision: Number(m[1]), dir: join(dir, m[0]) }))
    .sort((a, b) => b.revision - a.revision)
    .map(({ revision, dir: at }) => {
      const path = ["chrome-linux64", "chrome-linux"].map((sub) => join(at, sub, "chrome")).find((p) => existsSync(p));
      return path ? { path, revision } : undefined;
    })
    .find((hit) => hit !== undefined);
}

/** playwright's own wording when the binary for its version is not on disk. */
export function isMissingExecutable(error) {
  return error instanceof Error && error.message.includes("Executable doesn't exist");
}

/**
 * @param {import("playwright").BrowserType} chromium
 * @param {import("playwright").LaunchOptions} [options]
 * @returns {Promise<{ browser: import("playwright").Browser, source: string }>}
 */
export async function launchChromium(chromium, options = {}, env = process.env) {
  if (env.CHROMIUM) {
    const browser = await chromium.launch({ ...options, executablePath: env.CHROMIUM });
    return { browser, source: `${env.CHROMIUM} (CHROMIUM)` };
  }
  try {
    return { browser: await chromium.launch(options), source: "playwright's own download" };
  } catch (error) {
    if (!isMissingExecutable(error)) throw error;
    const found = newestChromium(env.PLAYWRIGHT_BROWSERS_PATH || DEFAULT_BROWSERS);
    if (!found) throw error;
    const browser = await chromium.launch({ ...options, executablePath: found.path });
    return {
      browser,
      source:
        `FALLBACK chromium-${found.revision}: playwright's own download is missing. ` +
        "`npx playwright install chromium` fetches the one it expects",
    };
  }
}
