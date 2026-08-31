#!/usr/bin/env node
/**
 * Open a headed Chromium on a PERSISTENT profile and wait for Dann to sign in.
 *
 * The agent never types the credentials: this script only opens the window and
 * watches for the signed-in markers to appear. The profile lives outside the
 * repo, so the session survives between runs and the round does not need a
 * fresh login each time.
 *
 * Usage: node test-kit/driver/web-login.mjs [--minutes 15]
 */
import { chromium } from "playwright";
import { mkdirSync } from "node:fs";

/**
 * AppLocker on this machine refuses to execute Playwright's bundled Chromium —
 * it lives under %LOCALAPPDATA%, and the policy only permits binaries in
 * Program Files ("This program is blocked by group policy", spawn UNKNOWN).
 * The system browsers ARE in Program Files, so drive one of those instead.
 * Edge first: the machine is AAD-joined and Edge is likeliest to carry the
 * work-account session without a password prompt.
 */
export const CHANNEL = process.env.SSF_CHANNEL ?? "msedge";

const PROFILE = process.env.SSF_PROFILE ?? "test-kit/out/browser-profile";

const minutesIdx = process.argv.indexOf("--minutes");
const MINUTES = minutesIdx === -1 ? 15 : Number(process.argv[minutesIdx + 1]);

/**
 * The debugging port every other script in here talks to.
 *
 * This was the one thing in the kit with no script behind it. `cdp-eval`,
 * `drive`, `pane`, `shot` and the rest all default to `127.0.0.1:9333`, and
 * nothing opened it: the browser had to be started by hand with the flag, and
 * the next session had no way to know that. Opening it here costs nothing when
 * it is unused and means the documented entry point is actually sufficient to
 * run a round.
 */
const CDP_PORT = process.env.SSF_CDP_PORT ?? "9333";

mkdirSync(PROFILE, { recursive: true });

console.log(`browser : channel=${CHANNEL}`);
console.log(`cdp     : http://127.0.0.1:${CDP_PORT}`);

const ctx = await chromium.launchPersistentContext(PROFILE, {
  headless: false,
  viewport: null,
  channel: CHANNEL,
  args: ["--start-maximized", `--remote-debugging-port=${CDP_PORT}`],
});

const page = ctx.pages()[0] ?? (await ctx.newPage());
await page.goto("https://www.office.com/", { waitUntil: "domcontentloaded" });

console.log(`profile : ${PROFILE}`);
console.log(`Waiting up to ${MINUTES} minutes for sign-in. Sign in in the window that just opened.`);
console.log("Watching for the app launcher / account manager to appear...\n");

/**
 * Signed-in markers that only render once a session exists.
 *
 * Two things made this miss a sign-in that had plainly worked, on 2026-08-30,
 * and cost the round a detour into "did the login fail?".
 *
 * The chrome MOVED. `office.com` now redirects to `m365.cloud.microsoft/chat`,
 * a Copilot shell that renders none of the five markers below — so the watcher
 * sat there reporting nothing while OneDrive was one click away and fully
 * signed in. Landing on a known signed-in HOST is therefore an answer in its
 * own right: those hosts do not serve a signed-out page, they redirect to a
 * login one, and the line above already refuses login hosts.
 *
 * And it watched ONE tab. Sign-in can land in a new tab or leave the first on
 * an interstitial, and the first tab is not the one to ask by then, so every
 * open page is asked and any of them may answer.
 */
const SIGNED_IN_HOST = /m365\.cloud\.microsoft|onedrive\.live\.com|-my\.sharepoint\.com|office\.com\/launch/;

async function pageSignedIn(p) {
  try {
    const url = p.url();
    if (/login\.microsoftonline\.com|login\.live\.com/.test(url)) return false;
    if (SIGNED_IN_HOST.test(url)) return true;
    return await p
      .locator(
        [
          '[aria-label="App launcher"]',
          "#O365_MainLink_NavMenu",
          '[data-automationid="waffle"]',
          '[aria-label*="Account manager"]',
          "#O365_MainLink_Me",
        ].join(", "),
      )
      .first()
      .isVisible()
      .catch(() => false);
  } catch {
    return false;
  }
}

async function signedIn() {
  for (const p of ctx.pages()) {
    if (await pageSignedIn(p)) return true;
  }
  return false;
}

const deadline = Date.now() + MINUTES * 60_000;
let ok = false;
while (Date.now() < deadline) {
  if (await signedIn()) {
    ok = true;
    break;
  }
  await page.waitForTimeout(3000);
}

if (ok) {
  // Let the tokens settle into the profile before the context closes.
  await page.waitForTimeout(4000);
  let who = "(name not read)";
  try {
    who =
      (await page.locator("#O365_MainLink_Me, [aria-label*='Account manager']").first().getAttribute("aria-label")) ??
      who;
  } catch {
    /* the label is a convenience, not the verdict */
  }
  console.log(`SIGNED IN. Session saved to the profile. account-manager label: ${who}`);
  console.log(`current url: ${page.url()}`);
} else {
  console.log("NOT SIGNED IN within the window. The profile was still saved; re-run to continue.");
}

await ctx.close();
process.exit(ok ? 0 : 1);
