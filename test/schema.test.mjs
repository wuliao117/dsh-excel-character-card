import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const sample = path.join(here, '..', 'sample', '角色卡示例.xlsx');
const tempHome = mkdtempSync(path.join(os.tmpdir(), 'dsh-ecc-schema-'));
process.env.DSH_HOME = tempHome;

const { apply } = await import('../lib/index.js');

const TYPES = new Set(['object', 'array', 'string', 'number', 'integer', 'boolean', 'null']);
const ALLOWED = new Set([
  'type',
  'oneOf',
  'properties',
  'required',
  'additionalProperties',
  'items',
  'enum',
  'const',
  'description',
  'title',
  '$comment',
  'default',
  'examples',
]);

/** Validate the raw-schema subset DSH enforces (no author-only `json`). */
function checkSchema(schema, at = 'schema') {
  assert.ok(schema && typeof schema === 'object' && !Array.isArray(schema), `${at} must be an object`);
  for (const key of Object.keys(schema)) assert.ok(ALLOWED.has(key), `${at}.${key} is not in the supported subset`);
  if (schema.oneOf !== undefined) {
    assert.ok(Array.isArray(schema.oneOf) && schema.oneOf.length >= 2, `${at}.oneOf must have >=2 branches`);
    assert.equal(schema.type, undefined, `${at} cannot have type and oneOf`);
    schema.oneOf.forEach((branch, index) => checkSchema(branch, `${at}.oneOf[${index}]`));
    return;
  }
  assert.ok(TYPES.has(schema.type), `${at}.type must be one of the supported primitive/container types`);
  if (schema.type === 'object') {
    assert.equal(typeof schema.additionalProperties, 'boolean', `${at}.additionalProperties must be boolean`);
    for (const [key, child] of Object.entries(schema.properties ?? {})) checkSchema(child, `${at}.properties.${key}`);
    for (const key of schema.required ?? []) assert.ok(Object.hasOwn(schema.properties ?? {}, key), `${at}.required names ${key}`);
  }
  if (schema.type === 'array' && schema.items !== undefined) checkSchema(schema.items, `${at}.items`);
}

function checkValue(schema, value, at = 'value') {
  if (schema.oneOf) {
    assert.ok(schema.oneOf.some((branch) => {
      try {
        checkValue(branch, value, at);
        return true;
      } catch {
        return false;
      }
    }), `${at} did not match any oneOf branch`);
    return;
  }
  switch (schema.type) {
    case 'object': {
      assert.ok(value !== null && typeof value === 'object' && !Array.isArray(value), `${at} must be an object`);
      for (const key of schema.required ?? []) assert.ok(Object.hasOwn(value, key), `${at}.${key} is required`);
      for (const [key, child] of Object.entries(schema.properties ?? {})) {
        if (Object.hasOwn(value, key)) checkValue(child, value[key], `${at}.${key}`);
      }
      if (schema.additionalProperties === false) {
        for (const key of Object.keys(value)) assert.ok(Object.hasOwn(schema.properties ?? {}, key), `${at}.${key} is not allowed`);
      }
      break;
    }
    case 'array':
      assert.ok(Array.isArray(value), `${at} must be an array`);
      if (schema.items) value.forEach((item, index) => checkValue(schema.items, item, `${at}[${index}]`));
      break;
    case 'string':
      assert.equal(typeof value, 'string', `${at} must be a string`);
      break;
    case 'number':
      assert.equal(typeof value, 'number', `${at} must be a number`);
      break;
    case 'integer':
      assert.ok(Number.isInteger(value), `${at} must be an integer`);
      break;
    case 'boolean':
      assert.equal(typeof value, 'boolean', `${at} must be a boolean`);
      break;
    case 'null':
      assert.equal(value, null, `${at} must be null`);
      break;
    default:
      assert.fail(`${at} has unknown schema type ${schema.type}`);
  }
  if (Array.isArray(schema.enum)) assert.ok(schema.enum.includes(value), `${at} must be one of the enum`);
}

let tool = null;
const ctx = {
  logger: () => ({ info() {}, warn() {}, error() {} }),
  effect(fn) { fn(); },
  tools: { register(value) { tool = value; return () => {}; } },
  systemPrompt: { getSectionOrder() { return 10200; }, section() { return () => {}; } },
  webServer: { register() { return () => {}; } },
};
apply(ctx);
assert.ok(tool);

checkSchema(tool.parameters, 'parameters');
checkSchema(tool.output.schema, 'output.schema');
assert.equal(tool.parameters.type, 'object');
assert.deepEqual(tool.parameters.required, ['action']);
assert.ok(!JSON.stringify(tool.output.schema).includes('"json"'));

const exec = { agent: { id: 'schema-session' }, signal: new AbortController().signal };
const results = [];
results.push(await tool.execute({ action: 'list' }, exec));
results.push(await tool.execute({ action: 'import', path: sample, activate: true }, exec));
results.push(await tool.execute({ action: 'preview', path: sample }, exec));
results.push(await tool.execute({ action: 'get', cardId: results[1].activeCardId }, exec));
results.push(await tool.execute({ action: 'activate', cardId: results[1].activeCardId }, exec));
results.push(await tool.execute({ action: 'deactivate' }, exec));
results.push(await tool.execute({ action: 'delete', cardId: results[1].activeCardId }, exec));

for (const result of results) {
  checkValue(tool.output.schema, result, `result(${result.action})`);
  const blocks = tool.output.render({}, result);
  assert.ok(Array.isArray(blocks) && blocks.length > 0);
  assert.equal(blocks[0].type, 'text');
  assert.equal(typeof blocks[0].text, 'string');
}

rmSync(tempHome, { recursive: true, force: true });
console.log('OK: raw tool schemas and every result variant satisfy the DSH subset');
