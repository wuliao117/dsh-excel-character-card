import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const sample = path.join(here, '..', 'sample', '角色卡示例.xlsx');
const tempHome = mkdtempSync(path.join(os.tmpdir(), 'dsh-ecc-'));
process.env.DSH_HOME = tempHome;

const { apply } = await import('../lib/index.js');

const registered = { tool: null, section: null, route: null };
const disposers = [];
const ctx = {
  logger: () => ({ info() {}, warn() {}, error() {} }),
  effect(fn, label) {
    const dispose = fn();
    if (typeof dispose === 'function') disposers.push(dispose);
    else disposers.push(() => {});
    return dispose;
  },
  tools: {
    register(tool) {
      registered.tool = tool;
      return () => {
        registered.tool = null;
      };
    },
  },
  systemPrompt: {
    getSectionOrder() {
      return 10200;
    },
    section(section) {
      registered.section = section;
      return () => {
        registered.section = null;
      };
    },
  },
  webServer: {
    register(route) {
      registered.route = route;
      return () => {
        registered.route = null;
      };
    },
  },
};

apply(ctx);
assert.ok(registered.tool, 'tool registered');
assert.ok(registered.section, 'prompt section registered');
assert.ok(registered.route, 'web route registered');
assert.equal(registered.tool.name, 'excel_character_card');
assert.equal(registered.route.path, '/excel-character-card/api');

const exec = { agent: { id: 'session-test' }, signal: new AbortController().signal };

let result = await registered.tool.execute({ action: 'list' }, exec);
assert.equal(result.ok, true);
assert.equal(result.total, 0);

result = await registered.tool.execute({ action: 'import', path: sample, activate: true }, exec);
assert.equal(result.ok, true);
assert.equal(result.report.added, 4);
assert.ok(result.activeCardId);

const prompt = registered.section.text({ agent: { id: 'session-test' } });
assert.match(prompt, /角色扮演模式/);
assert.match(prompt, /林晚/);
assert.match(prompt, /168 cm/);
assert.match(prompt, /独立咖啡师/);

// Other sessions should not inherit a session-scoped activation.
assert.equal(registered.section.text({ agent: { id: 'other-session' } }), '');

result = await registered.tool.execute({ action: 'get', cardId: result.activeCardId }, exec);
assert.equal(result.ok, true);
assert.equal(result.card.name, '林晚');
assert.match(result.prompt, /晚风/);

result = await registered.tool.execute({ action: 'deactivate' }, exec);
assert.equal(result.ok, true);
assert.equal(registered.section.text({ agent: { id: 'session-test' } }), '');

result = await registered.tool.execute({ action: 'list' }, exec);
const firstId = result.cards[0].id;
result = await registered.tool.execute({ action: 'activate', cardId: firstId }, exec);
assert.equal(result.ok, true);
assert.match(registered.section.text({ agent: { id: 'session-test' } }), /角色扮演模式/);

// Global activation (no session) should apply to every session.
await registered.tool.execute({ action: 'deactivate' }, exec);
result = await registered.tool.execute({ action: 'activate', cardId: firstId }, { agent: undefined, signal: exec.signal });
assert.equal(result.ok, true);
assert.match(registered.section.text({ agent: { id: 'any-session' } }), /角色扮演模式/);

// State file exists and survives a fresh plugin instance.
const stateFile = path.join(tempHome, 'excel-character-card', 'state.json');
assert.ok(existsSync(stateFile), 'state persisted');

const second = { tool: null, section: null, route: null };
const ctx2 = {
  logger: () => ({ info() {}, warn() {}, error() {} }),
  effect(fn) {
    const dispose = fn();
    return typeof dispose === 'function' ? dispose : () => {};
  },
  tools: { register(tool) { second.tool = tool; return () => {}; } },
  systemPrompt: { getSectionOrder() { return 10200; }, section(section) { second.section = section; return () => {}; } },
  webServer: { register(route) { second.route = route; return () => {}; } },
};
const mod2 = await import('../lib/index.js?reload=1');
mod2.apply(ctx2);
const list2 = await second.tool.execute({ action: 'list' }, { agent: { id: 'session-test' }, signal: exec.signal });
assert.equal(list2.total, 4);
assert.match(second.section.text({ agent: { id: 'fresh-session' } }), /角色扮演模式/);

rmSync(tempHome, { recursive: true, force: true });
console.log('OK: host tool, prompt section, per-session/global activation, persistence');
