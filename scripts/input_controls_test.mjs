// Run the real input initializer with deterministic DOM/socket/paint boundaries.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const html = fs.readFileSync(new URL('../zadoo_vnc/templates/index.html', import.meta.url), 'utf8');
const start = html.indexOf('    function initializeInputControls() {');
const end = html.indexOf("    if (typeof window.connectInput !== 'function')", start);
assert.ok(start > 0 && end > start);
const listeners = new Map();
const windowListeners = new Map();
const frames = new Map();
let nextFrame = 0;
let current;
const remoteTarget = { tagName: 'CANVAS', closest: () => null };
const localTarget = { tagName: 'BUTTON', closest: () => ({}) };
const element = { style: {}, classList: { contains: () => false }, getBoundingClientRect: () => ({ left: 0, top: 0, width: 1000, height: 800 }) };
const canvas = { ...element, width: 1000, height: 800 };
const add = (map, name, fn) => map.set(name, [...(map.get(name) || []), fn]);
class Socket {
  static OPEN = 1;
  static CONNECTING = 0;
  static CLOSED = 3;
  constructor() { this.readyState = 1; this.bufferedAmount = 0; this.sent = []; current = this; }
  send(data) { this.sent.push(JSON.parse(data)); }
}
const context = vm.createContext({
  window: { addEventListener: (name, fn) => add(windowListeners, name, fn) },
  document: { body: { style: {} }, hidden: false,
    getElementById: id => id === 'screen' ? canvas : element,
    querySelector: () => element,
    addEventListener: (name, fn) => add(listeners, name, fn) },
  canvas, location: { protocol: 'http:', host: 'localhost' }, WebSocket: Socket,
  blockMouse: false, blockKeyboard: false, isKeyBlockedByUser: () => false,
  requestAnimationFrame: fn => { frames.set(++nextFrame, fn); return nextFrame; },
  cancelAnimationFrame: id => frames.delete(id), setTimeout: () => 0,
  updateStatus() {}, updateConnectionStatus() {}, fetchPublicUrl() {}, requestCaptureStats() {},
});
vm.runInContext(html.slice(start, end), context);
const connect = () => vm.runInContext('initializeInputControls()', context);
const fire = (name, overrides = {}) => {
  const event = { target: remoteTarget, clientX: 500, clientY: 400, button: 0,
    key: 'a', code: 'KeyA', preventDefault() {}, ...overrides };
  for (const fn of listeners.get(name) || []) fn(event);
};
const paint = () => { const batch = [...frames.values()]; frames.clear(); for (const fn of batch) fn(); };
connect();
for (let i = 0; i < 10; i++) { current.readyState = Socket.CLOSED; current.onclose({ code: 1006, reason: 'test' }); connect(); }
for (const event of ['mousedown', 'mouseup', 'mousemove', 'wheel', 'keydown', 'keyup']) assert.equal(listeners.get(event).length, 1, event);
fire('keydown'); fire('keyup');
assert.deepEqual(current.sent.map(e => e.state), ['down', 'up']);
console.log('INPUT_RECONNECT: PASS 10 reconnects, exactly one listener and key pair');
current.sent = [];
for (let i = 0; i < 1000; i++) fire('mousemove', { clientX: i });
assert.equal(current.sent.length, 1, 'first movement must send immediately');
assert.equal(frames.size, 1, 'only one pending paint');
paint();
assert.equal(current.sent.length, 2);
assert.equal(current.sent[1].x, .999);
paint(); assert.equal(frames.size, 0);
console.log('INPUT_BURST: PASS 1000 moves -> 2 sends, immediate first and exact final position');
current.sent = []; current.bufferedAmount = 65536;
for (let i = 0; i < 1000; i++) fire('mousemove', { clientX: i });
for (let i = 0; i < 10; i++) { paint(); assert.equal(frames.size, 1); }
assert.equal(current.sent.length, 0);
current.bufferedAmount = 0; paint();
assert.equal(current.sent.length, 1); assert.equal(current.sent[0].x, .999);
paint();
console.log('INPUT_BACKPRESSURE: PASS one latest motion, no queued flood');
current.sent = [];
fire('mousedown'); fire('mousemove', { clientX: 650 });
fire('mouseup', { target: localTarget, clientX: 1100 });
assert.deepEqual(current.sent.at(-1), { action: 'click', x: .65, y: .5, button: 'left', state: 'up' });
fire('keydown', { key: 'A' }); fire('keyup', { target: localTarget, key: 'a' });
assert.deepEqual(current.sent.slice(-2).map(e => e.key), ['A', 'A']);
current.sent = []; fire('keydown', { target: localTarget }); fire('keyup', { target: localTarget });
assert.equal(current.sent.length, 0);
fire('keydown'); fire('mousedown');
for (const fn of windowListeners.get('blur')) fn();
assert.deepEqual(current.sent.slice(-2).map(e => e.state), ['up', 'up']);
assert.equal(frames.size, 0);
console.log('INPUT_RELEASE: PASS drag over controls, shifted keys, local UI isolation, blur cleanup');
