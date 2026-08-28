#!/usr/bin/env node
/**
 * Put a local file on a file input that lives in a nested same-origin frame.
 *
 * The Office Add-ins "Upload My Add-in" dialog is a frame inside the editor
 * OOPIF, so Playwright cannot reach its <input type=file>. Resolve the element
 * to a CDP objectId with Runtime.evaluate, then hand it to DOM.setFileInputFiles.
 *
 * Usage:
 *   node test-kit/driver/cdp-setfile.mjs <target-url-substring> "<js returning the input>" <absolute-file>
 */
const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";
const [, , match, expr] = process.argv;
const files = process.argv.slice(4);

const targets = await (await fetch(`${BASE}/json/list`)).json();
const target = targets.find((t) => (t.type === "page" || t.type === "iframe") && t.url.includes(match));
if (!target) {
  console.error(`no target matching "${match}"`);
  process.exit(3);
}

const ws = new WebSocket(target.webSocketDebuggerUrl);
let id = 0;
const pending = new Map();
const send = (method, params = {}) => {
  const msgId = ++id;
  ws.send(JSON.stringify({ id: msgId, method, params }));
  return new Promise((resolve, reject) => pending.set(msgId, { resolve, reject }));
};
ws.addEventListener("message", (e) => {
  const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) {
    const { resolve, reject } = pending.get(m.id);
    pending.delete(m.id);
    if (m.error) reject(new Error(JSON.stringify(m.error)));
    else resolve(m.result);
  }
});
await new Promise((res, rej) => {
  ws.addEventListener("open", res);
  ws.addEventListener("error", () => rej(new Error("ws open failed")));
});

try {
  await send("DOM.enable");
  await send("Runtime.enable");

  const ev = await send("Runtime.evaluate", { expression: expr, returnByValue: false });
  if (ev.exceptionDetails) throw new Error(ev.exceptionDetails.exception?.description ?? "evaluate threw");
  const objectId = ev.result.objectId;
  if (!objectId) throw new Error(`expression did not return an element (got ${ev.result.type})`);

  await send("DOM.setFileInputFiles", { files, objectId });
  console.log(`set ${files.length} file(s): ${files.join(", ")}`);

  // Read the input back: a file input that took the file reports it here.
  const check = await send("Runtime.evaluate", {
    expression: `(() => { const i = (${expr}); return i && i.files && i.files.length ? [...i.files].map(f => f.name + "(" + f.size + "B)").join(", ") : "NO FILE ON INPUT"; })()`,
    returnByValue: true,
  });
  console.log(`input now holds: ${check.result.value}`);
} finally {
  ws.close();
}
