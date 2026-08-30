#!/usr/bin/env node
/**
 * Send REAL key events to an element in a CDP target.
 *
 * `cdp-type.mjs` exists because React ignored the native-setter trick on this
 * pane's text boxes. The SAME trap is waiting one control along, on the
 * `<select>` behind "A blank cell leaves a blank": setting `.value` and
 * dispatching a `change` moved the DOM and not React, so the forecast stayed at
 * "Add 6 slides" while the select read `skip`.
 *
 * That failure looks exactly like a product bug — the round of 2026-08-30 was
 * one screenshot away from reporting "the number on the button does not change"
 * as a defect, which is the very thing that round was asked to check. It is not
 * a defect: driven with real ArrowDown keys the line becomes "2 of 3 rows x 2
 * slides", the button becomes "Add 4 slides", and the merge then adds 4.
 *
 * Keyboard interaction with a focused control is handled by the browser itself
 * and raises a TRUSTED `change`, so there is nothing left for React to
 * disbelieve.
 *
 * Usage: node test-kit/driver/cdp-key.mjs <target-substring> "<js returning element>" <Key> [repeat]
 *
 *   node test-kit/driver/cdp-key.mjs ssf-merge.struktureretsundfornuft.dk \
 *     "[...document.querySelectorAll('select')].find(x=>/Leave the whole row out/.test(x.textContent))" \
 *     ArrowDown 2
 */
const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";
const [, , match, expr, key, repeatArg] = process.argv;

if (!match || !expr || !key) {
  console.error('usage: cdp-key.mjs <target-substring> "<js returning element>" <Key> [repeat]');
  process.exit(2);
}
const repeat = Number(repeatArg ?? 1);

/**
 * Only the keys a select needs, spelled the way CDP wants them.
 *
 * `windowsVirtualKeyCode` is not optional here: without it Chrome delivers the
 * event and the default action never runs, so the select does not move and the
 * whole exercise silently proves nothing.
 */
const KEYS = {
  ArrowDown: { code: "ArrowDown", windowsVirtualKeyCode: 40 },
  ArrowUp: { code: "ArrowUp", windowsVirtualKeyCode: 38 },
  Enter: { code: "Enter", windowsVirtualKeyCode: 13 },
  Home: { code: "Home", windowsVirtualKeyCode: 36 },
  End: { code: "End", windowsVirtualKeyCode: 35 },
};
const spec = KEYS[key];
if (!spec) {
  console.error(`unknown key "${key}" — known: ${Object.keys(KEYS).join(", ")}`);
  process.exit(2);
}

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
ws.addEventListener("message", (ev) => {
  const m = JSON.parse(ev.data);
  const p = pending.get(m.id);
  if (!p) return;
  pending.delete(m.id);
  if (m.error) p.reject(new Error(m.error.message ?? JSON.stringify(m.error)));
  else p.resolve(m.result);
});

await new Promise((r) => ws.addEventListener("open", r, { once: true }));
await send("Runtime.enable");

const evaluate = async (expression) =>
  (await send("Runtime.evaluate", { expression, returnByValue: true })).result.value;

// Focus first, or the keys land on whatever had focus and the element under
// test is not the one that moved.
const before = await evaluate(
  `(function(){ const el = ${expr}; if(!el) return "NOT FOUND"; el.focus(); return el.tagName + " value=" + (el.value ?? ""); })()`,
);
if (before === "NOT FOUND") {
  console.error(`the expression matched no element`);
  process.exit(4);
}
console.log(`focused ${before}`);

for (let i = 0; i < repeat; i++) {
  await send("Input.dispatchKeyEvent", { type: "rawKeyDown", key, ...spec });
  await send("Input.dispatchKeyEvent", { type: "keyUp", key, ...spec });
  // The pane re-renders on the change; give React the tick it needs before the
  // next key, or a burst of them lands on a control that is being replaced.
  await new Promise((r) => setTimeout(r, 300));
}

const after = await evaluate(`(function(){ const el = ${expr}; return el ? String(el.value) : "gone"; })()`);
console.log(`value now: ${after}`);

// The DOM moving is NOT the thing being tested — React's state is. Say so, so a
// reader does not take this line as proof the pane agreed.
console.log(`(check the pane's own text: the DOM moving does not prove the state did)`);
ws.close();
