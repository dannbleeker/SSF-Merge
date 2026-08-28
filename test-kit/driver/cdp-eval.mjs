#!/usr/bin/env node
/**
 * Evaluate JavaScript inside a specific CDP target, by URL substring.
 *
 * PowerPoint for the web nests the editor and the Office Add-ins store in
 * out-of-process iframes. Playwright's connectOverCDP does not attach to them
 * here — frame.evaluate reports "Frame was detached" — but each OOPIF is its
 * own CDP target with its own websocket, and talking to that directly works.
 *
 * Usage:
 *   node test-kit/driver/cdp-eval.mjs list
 *   node test-kit/driver/cdp-eval.mjs <url-substring> "<expression>"
 */
const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";

const targets = await (await fetch(`${BASE}/json/list`)).json();

if (process.argv[2] === "list") {
  for (const t of targets) {
    if (t.type !== "page" && t.type !== "iframe") continue;
    console.log(`${t.type.padEnd(6)} ${t.url.slice(0, 110)}`);
  }
  process.exit(0);
}

const match = process.argv[2];
const expr = process.argv[3];
const hits = targets.filter((t) => (t.type === "page" || t.type === "iframe") && t.url.includes(match));
if (!hits.length) {
  console.error(`no page/iframe target whose url contains "${match}"`);
  process.exit(3);
}
const target = hits[0];
console.error(`target: ${target.type} ${target.url.slice(0, 90)}`);

const ws = new WebSocket(target.webSocketDebuggerUrl);
let id = 0;
const pending = new Map();

function send(method, params = {}) {
  const msgId = ++id;
  ws.send(JSON.stringify({ id: msgId, method, params }));
  return new Promise((resolve, reject) => pending.set(msgId, { resolve, reject }));
}

ws.addEventListener("message", (e) => {
  const msg = JSON.parse(e.data);
  if (msg.id && pending.has(msg.id)) {
    const { resolve, reject } = pending.get(msg.id);
    pending.delete(msg.id);
    if (msg.error) reject(new Error(JSON.stringify(msg.error)));
    else resolve(msg.result);
  }
});

await new Promise((resolve, reject) => {
  ws.addEventListener("open", resolve);
  ws.addEventListener("error", () => reject(new Error("websocket failed to open")));
});

try {
  const res = await send("Runtime.evaluate", {
    expression: expr,
    returnByValue: true,
    awaitPromise: true,
    userGesture: true,
  });
  if (res.exceptionDetails) {
    console.log("EXCEPTION:", JSON.stringify(res.exceptionDetails.exception?.description ?? res.exceptionDetails));
  } else {
    const v = res.result.value;
    console.log(typeof v === "object" ? JSON.stringify(v, null, 1) : String(v));
  }
} finally {
  ws.close();
}
