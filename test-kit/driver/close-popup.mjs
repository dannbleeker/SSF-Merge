#!/usr/bin/env node
/**
 * Close any page whose URL matches the argument, and report what is left.
 *
 * The Office Add-ins store opens an MSA consent popup that this account is not
 * permitted to complete; it sits on top blocking the dialog underneath.
 *
 * Usage: node test-kit/driver/close-popup.mjs <url-substring>
 */
import { chromium } from "playwright";

const match = process.argv[2] ?? "login.live.com";
const browser = await chromium.connectOverCDP(process.env.SSF_CDP ?? "http://127.0.0.1:9333");
const ctx = browser.contexts()[0];

let closed = 0;
for (const p of ctx.pages()) {
  if (p.url().includes(match)) {
    await p.close().catch(() => {});
    closed++;
  }
}
console.log(`closed ${closed} page(s) matching "${match}"`);
for (const p of ctx.pages()) console.log(`  remains: ${p.url().slice(0, 120)}`);

await browser.close();
