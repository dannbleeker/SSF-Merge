#!/usr/bin/env node
/**
 * Shut the CDP-attached browser down the way a person does.
 *
 * Killing the process instead — `Stop-Process -Force`, `taskkill /F` — works,
 * and leaves the profile marked as having crashed. Edge then opens the next
 * session with a "Restore pages?" bubble over the window, which is somebody
 * else's dialog sitting on top of the round: a driver that clicks by coordinate
 * hits the wrong thing, and a human who reopens the profile is asked to clean
 * up after a crash that never happened.
 *
 * `Browser.close` is the graceful path. It runs the same shutdown the window's
 * X button does — sessions flushed, exit recorded as clean — so the next launch
 * comes up on a fresh tab with nothing over it.
 *
 * Usage: node test-kit/driver/close-browser.mjs
 */
const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";

let version;
try {
  version = await (await fetch(`${BASE}/json/version`)).json();
} catch {
  console.log(`nothing listening on ${BASE} — already closed`);
  process.exit(0);
}

const endpoint = version.webSocketDebuggerUrl;
if (!endpoint) {
  console.error("no browser-level websocket: this build will not take Browser.close");
  process.exit(3);
}

// The tabs first, so anything mid-write gets its beforeunload rather than
// having the window pulled from under it.
try {
  const targets = await (await fetch(`${BASE}/json/list`)).json();
  for (const t of targets.filter((x) => x.type === "page")) {
    await fetch(`${BASE}/json/close/${t.id}`).catch(() => {});
  }
  console.log(`closed ${targets.filter((x) => x.type === "page").length} tab(s)`);
} catch {
  /* the browser can go without this; it is a courtesy to the pages */
}

const ws = new WebSocket(endpoint);
await new Promise((r) => ws.addEventListener("open", r, { once: true }));
ws.send(JSON.stringify({ id: 1, method: "Browser.close" }));

// It answers by going away, so waiting for a reply would hang. Wait for the
// endpoint to stop answering instead, which is the thing actually being asked.
const deadline = Date.now() + 15_000;
let gone = false;
while (Date.now() < deadline) {
  await new Promise((r) => setTimeout(r, 500));
  try {
    await (await fetch(`${BASE}/json/version`)).json();
  } catch {
    gone = true;
    break;
  }
}
try {
  ws.close();
} catch {
  /* it is already gone, which is the point */
}

console.log(gone ? "browser closed cleanly — no crash flag, no restore prompt" : "STILL ANSWERING after 15s");
process.exit(gone ? 0 : 1);
