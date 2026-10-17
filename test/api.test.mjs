import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { readFileSync, mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const sample = path.join(here, '..', 'sample', '角色卡示例.xlsx');
const tempHome = mkdtempSync(path.join(os.tmpdir(), 'dsh-ecc-api-'));
process.env.DSH_HOME = tempHome;

const { apply } = await import('../lib/index.js');

let route = null;
const ctx = {
  logger: () => ({ info() {}, warn() {}, error() {} }),
  effect(fn) {
    fn();
  },
  tools: { register() { return () => {}; } },
  systemPrompt: { getSectionOrder() { return 10200; }, section() { return () => {}; } },
  webServer: {
    register(value) {
      route = value;
      return () => {};
    },
  },
};
apply(ctx);
assert.ok(route);

const server = createServer((req, res) => {
  route.handler(req, res).catch((error) => {
    res.statusCode = 500;
    res.end(String(error));
  });
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const { port } = server.address();
const base = `http://127.0.0.1:${port}/excel-character-card/api`;

async function json(url, init) {
  const response = await fetch(url, init);
  const body = await response.json();
  return { status: response.status, body };
}

const health = await json(`${base}/health`);
assert.equal(health.status, 200);
assert.equal(health.body.ok, true);

let state = await json(`${base}/state?sessionId=s1`);
assert.equal(state.body.result.total, 0);

const dataBase64 = readFileSync(sample).toString('base64');
const imported = await json(`${base}/import`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ sessionId: 's1', fileName: '角色卡示例.xlsx', dataBase64 }),
});
assert.equal(imported.status, 200, JSON.stringify(imported.body));
assert.equal(imported.body.result.total, 4);
assert.equal(imported.body.result.report.added, 4);

state = await json(`${base}/state?sessionId=s1`);
assert.equal(state.body.result.cards.length, 4);
assert.equal(state.body.result.activeCardId, null);

const first = state.body.result.cards[0];
const activated = await json(`${base}/activate`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ sessionId: 's1', cardId: first.id }),
});
assert.equal(activated.status, 200);
assert.equal(activated.body.result.activeCardId, first.id);

const detail = await json(`${base}/card?id=${encodeURIComponent(first.id)}`);
assert.equal(detail.status, 200);
assert.equal(detail.body.result.card.name, first.name);
assert.match(detail.body.result.prompt, /角色扮演模式/);

const deactivated = await json(`${base}/deactivate`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ sessionId: 's1' }),
});
assert.equal(deactivated.status, 200);
assert.equal(deactivated.body.result.activeCardId, null);

const deleted = await json(`${base}/delete`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ sessionId: 's1', cardId: first.id }),
});
assert.equal(deleted.status, 200);
assert.equal(deleted.body.result.total, 3);

await new Promise((resolve) => server.close(resolve));
rmSync(tempHome, { recursive: true, force: true });
console.log('OK: HTTP API upload/state/activate/detail/deactivate/delete');
