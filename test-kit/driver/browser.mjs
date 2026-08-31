/**
 * Start the browser a round drives, and leave it running.
 *
 * This was the hole in the middle of the kit. Every other script here defaults
 * to `127.0.0.1:9333` and none of them opened it, so the browser had to be
 * started by hand with the right flags and the next session had no way to learn
 * which ones. A round that cannot start is not repeatable however good the rest
 * of it is.
 *
 * `web-login.mjs` is NOT this. It launches through Playwright, which owns the
 * process, so the browser dies the moment that script returns — including the
 * moment it detects a successful sign-in. It is for establishing the session in
 * the profile, once. This is for using it afterwards, and it spawns Edge
 * detached so the browser outlives the command.
 *
 * Usage:
 *
 *     node test-kit/driver/browser.mjs                 # open, wait for CDP
 *     node test-kit/driver/browser.mjs <url>           # and go somewhere
 *
 * Signing in is never the agent's to do. If the profile has no session this
 * opens a signed-out browser and says so, rather than pretending otherwise.
 */
import { spawn } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
import { resolve } from "node:path";

/**
 * AppLocker on this machine refuses Playwright's bundled Chromium — it lives
 * under %LOCALAPPDATA% and the policy only permits binaries in Program Files.
 * The system Edge is in Program Files, which is why the round drives that one.
 */
const EDGES = [
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
];

const PORT = process.env.SSF_CDP_PORT ?? "9333";
const PROFILE = resolve(process.env.SSF_PROFILE ?? "test-kit/out/browser-profile");
const BASE = `http://127.0.0.1:${PORT}`;

const alive = async () => {
  try {
    return (await (await fetch(`${BASE}/json/version`)).json()).Browser;
  } catch {
    return null;
  }
};

const already = await alive();
if (already) {
  console.log(`already up: ${already} on ${BASE}`);
  process.exit(0);
}

const exe = EDGES.find((p) => existsSync(p));
if (!exe) throw new Error(`no Edge found in Program Files; looked at:\n  ${EDGES.join("\n  ")}`);
mkdirSync(PROFILE, { recursive: true });

// detached + unref, or the browser goes down with this process and every later
// script in the round finds a refused connection instead of a browser.
const child = spawn(
  exe,
  [`--remote-debugging-port=${PORT}`, `--user-data-dir=${PROFILE}`, "--start-maximized", ...process.argv.slice(2)],
  { detached: true, stdio: "ignore" },
);
child.unref();

const deadline = Date.now() + 45_000;
while (Date.now() < deadline) {
  await new Promise((r) => setTimeout(r, 1000));
  const version = await alive();
  if (version) {
    console.log(`${version} on ${BASE}`);
    console.log(`profile: ${PROFILE}`);
    process.exit(0);
  }
}
throw new Error(`Edge started but ${BASE} never answered within 45s`);
