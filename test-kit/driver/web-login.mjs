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

mkdirSync(PROFILE, { recursive: true });

console.log(`browser : channel=${CHANNEL}`);

const ctx = await chromium.launchPersistentContext(PROFILE, {
  headless: false,
  viewport: null,
  channel: CHANNEL,
  args: ["--start-maximized"],
});

const page = ctx.pages()[0] ?? (await ctx.newPage());
await page.goto("https://www.office.com/", { waitUntil: "domcontentloaded" });

console.log(`profile : ${PROFILE}`);
console.log(`Waiting up to ${MINUTES} minutes for sign-in. Sign in in the window that just opened.`);
console.log("Watching for the app launcher / account manager to appear...\n");

/** Signed-in markers that only render once a session exists. */
async function signedIn() {
  try {
    const url = page.url();
    if (/login\.microsoftonline\.com|login\.live\.com/.test(url)) return false;
    const hit = await page
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
    return hit;
  } catch {
    return false;
  }
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
