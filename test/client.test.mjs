import assert from 'node:assert/strict';

let definition = null;
globalThis.window = {
  __ModuleLoader__: {
    load(value) {
      definition = value;
    },
  },
};

const ReactMock = {
  Fragment: Symbol('Fragment'),
  createElement(type, props, ...children) {
    return { type, props, children };
  },
  useState(initial) {
    return [typeof initial === 'function' ? initial() : initial, () => {}];
  },
  useEffect() {},
  useRef(value) {
    return { current: value };
  },
};

const ReactDOMMock = {
  createPortal(node) {
    return node;
  },
};

await import('../lib/client.js');
assert.ok(definition, 'client bundle registered with __ModuleLoader__');
assert.equal(definition.id, 'dsh-excel-character-card');

const mod = definition.factory((name) => {
  if (name === 'react') return ReactMock;
  if (name === 'react-dom') return ReactDOMMock;
  throw new Error(`unexpected require: ${name}`);
});
assert.equal(typeof mod.apply, 'function');
assert.deepEqual(mod.inject, ['slots']);

let registration = null;
let injectedKey = null;
const ctx = {
  slots: {
    inject(key, callback) {
      injectedKey = key;
      return callback();
    },
    register(options, component) {
      registration = { options, component };
      return () => {};
    },
  },
};
mod.apply(ctx);
assert.equal(injectedKey, 'conversation.input.right');
assert.equal(registration.options.name, 'conversation.input.right');
assert.equal(registration.options.id, 'excel-character-card');
assert.equal(typeof registration.component, 'function');

console.log('OK: client bundle registers the composer trigger');
