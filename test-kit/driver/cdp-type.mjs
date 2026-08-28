#!/usr/bin/env node
/**
 * Focus an element in a CDP target and type into it with real input events.
 *
 * Setting `.value` through the native setter is the usual way to satisfy React's
 * value tracker, but this pane's Step 1 button stayed disabled afterwards — the
 * state never moved. Real key events go through the same path a person's typing
 * does, so there is nothing left to disagree about.
 *
 * Usage: node test-kit/driver/cdp-type.mjs <target-substring> "<js returning element>" <text>
 */
const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";
const [, , match, expr, text] = process.argv;

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
  // Focus, and clear whatever is there, without touching React's value directly.
  const focused = await send("Runtime.evaluate", {
    expression: `(() => { const el = (${expr}); if (!el) return 'no element'; el.focus(); el.select && el.select(); return 'focused ' + el.tagName; })()`,
    returnByValue: true,
    userGesture: true,
  });
  console.log(focused.result.value);

  await send("Input.dispatchKeyEvent", { type: "keyDown", key: "Delete", code: "Delete", windowsVirtualKeyCode: 46 });
  await send("Input.dispatchKeyEvent", { type: "keyUp", key: "Delete", code: "Delete", windowsVirtualKeyCode: 46 });

  if (process.argv[5] === "insert") {
    // Tab-separated, multi-line content cannot be typed key-by-key: a Tab moves
    // focus and an Enter submits. insertText delivers it the way a paste does.
    await send("Input.insertText", { text });
  } else {
    for (const ch of text) {
      await send("Input.dispatchKeyEvent", { type: "keyDown", text: ch, key: ch });
      await send("Input.dispatchKeyEvent", { type: "keyUp", key: ch });
    }
  }

  const after = await send("Runtime.evaluate", {
    expression: `(() => { const el = (${expr}); return el ? el.value : '(gone)'; })()`,
    returnByValue: true,
  });
  console.log(`value now: ${JSON.stringify(after.result.value)}`);
} finally {
  ws.close();
}
