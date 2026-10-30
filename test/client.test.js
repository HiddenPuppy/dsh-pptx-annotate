/**
 * Tests for the browser half as shipped: the module loads through the DSH
 * module loader contract and registers the right previewer with the right
 * options, plus a render smoke test through a stubbed React.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

/** @type {any} */
let mod;
/** @type {any[]} */
const registered = [];
/** @type {any[]} */
const slotInjects = [];
/** @type {any[]} */
const slotRegistrations = [];

const fakeReact = {
  createElement: (type, props, ...children) => ({ type, props: props ?? {}, children }),
  useState: (initial) => [typeof initial === 'function' ? initial() : initial, () => {}],
  useEffect: () => {},
  useRef: () => ({ current: null }),
  useMemo: (fn) => fn(),
};

globalThis.window = {
  __ModuleLoader__: {
    load(m) {
      mod = m;
    },
  },
};

await import('../lib/client.js');

/**
 * Require shim mirroring the loader's externals.
 * @param {string} id module id
 * @returns {any} module
 */
function fakeRequire(id) {
  if (id === 'react') return fakeReact;
  throw new Error(`unexpected external: ${id}`);
}

const facade = mod.factory(fakeRequire);

/**
 * Fake client context capturing registrations.
 * @returns {any} context
 */
function makeContext() {
  const ctx = {
    effect: (fn) => fn(),
    documentPreviews: {
      register(def) {
        registered.push(def);
        return () => {};
      },
    },
    slots: {
      inject(name, factory) {
        slotInjects.push(name);
        factory();
        return () => {};
      },
      register(spec, component) {
        slotRegistrations.push({ spec, component });
        return () => {};
      },
    },
    sessions: {
      scope: (id) =>
        ctx.scopeMissing
          ? undefined
          : {
              get: (key) =>
                key === "conversation"
                  ? {
                      send: (text) => {
                        ctx.sent = text;
                        return Promise.resolve();
                      },
                    }
                  : undefined,
            },
    },
  };
  return ctx;
}

test('client: module follows the loader contract', () => {
  assert.equal(mod.id, 'dsh-pptx-annotate');
  assert.equal(typeof mod.factory, 'function');
  assert.equal(typeof facade.apply, 'function');
  assert.deepEqual(facade.inject, ['slots', 'documentPreviews', 'sessions']);
});

test('client: registers a pptx previewer that outranks the builtin', () => {
  const ctx = makeContext();
  facade.apply(ctx);
  assert.equal(registered.length, 1);
  const def = registered[0];
  assert.equal(def.id, 'dsh-pptx-annotate');
  assert.deepEqual(def.extensions, ['pptx']);
  assert.deepEqual(def.binaryExtensions, ['pptx']);
  assert.equal(def.priority, 'extension');
  assert.equal(def.loading, 'renderer');
  assert.equal(def.wrap, false);
  assert.equal(typeof def.title(), 'string');
});

test('client: mounts its body in the document slot under the same id', () => {
  assert.deepEqual(slotInjects, ['conversation.composer.dock', 'sidebar.right.tab.document']);
  assert.equal(slotRegistrations.length, 2);
  const doc = slotRegistrations.find((r) => r.spec.name === 'sidebar.right.tab.document');
  assert.ok(doc, 'document body registered');
  assert.equal(doc.spec.key, 'dsh-pptx-annotate');
  assert.equal(typeof doc.component, 'function');

  const activity = slotRegistrations.find((r) => r.spec.name === 'conversation.composer.dock');
  assert.ok(activity, 'composer activity chip registered');
  assert.equal(typeof activity.spec.inject, 'function');
  assert.equal(typeof activity.component, 'function');
});

test('client: submits a formatted prompt through the session scope', async () => {
  const ctx = makeContext();
  facade.apply(ctx);
  const { spec } = slotRegistrations.filter((r) => r.spec.name === 'sidebar.right.tab.document').pop();
  const props = spec.inject('session-42');
  assert.equal(typeof props.onSubmit, 'function');

  await props.onSubmit(facade.formatPrompt('deck.pptx', [
    { slide: 2, elementId: '#4', label: '图片', note: '换一张', order: 1 },
  ]));
  assert.match(ctx.sent, /【PPT 批注】`deck\.pptx` · 共 1 条/);
  assert.match(ctx.sent, /第 2 页 · #4 · 图片：换一张/);
});

test('client: submission fails loudly when the session scope is gone', () => {
  const ctx = makeContext();
  facade.apply(ctx);
  const { spec } = slotRegistrations.filter((r) => r.spec.name === 'sidebar.right.tab.document').pop();
  const props = spec.inject('session-404');
  const before = ctx.scopeMissing;
  ctx.scopeMissing = true;
  assert.throws(() => props.onSubmit('x'), /尚未绑定/);
  ctx.scopeMissing = before;
});

test('client: composer chip renders nothing until a submission is pending', () => {
  const tree = facade.AnnotateActivity({ sessionId: 'session-quiet' });
  assert.equal(tree, null);
});

test('client: prompt builder drops empty notes and numbers the rest', () => {
  const text = facade.formatPrompt('x.pptx', [
    { slide: 1, elementId: '#1', label: 'A', note: '  ', order: 1 },
    { slide: 1, elementId: '#2', label: 'B', note: '改字号', order: 2 },
  ]);
  assert.equal(text.split('\n').filter((l) => l.startsWith('- ')).length, 1);
});

test('client: pickAt agrees with the tested reference implementation', async () => {
  const { pickAt } = await import('../src/annotations.js');
  const shapes = [
    { x: 0.0, y: 0.0, w: 0.9, h: 0.9, kind: 'shape', text: '' },
    { x: 0.1, y: 0.1, w: 0.3, h: 0.2, kind: 'text', text: '标题' },
    { x: 0.1, y: 0.1, w: 0.05, h: 0.05, kind: 'shape', text: '' },
  ];
  for (const point of [[0.2, 0.15], [0.5, 0.5], [0.95, 0.95], [0.12, 0.12]]) {
    const a = facade.pickAt(shapes, point[0], point[1]);
    const b = pickAt(shapes, point[0], point[1]);
    assert.deepEqual(a && a.index, b && b.index, `same pick at ${point}`);
  }
});

test('client: body renders a placeholder while the path is unresolved', () => {
  const { component } = slotRegistrations.find((r) => r.spec.name === 'sidebar.right.tab.document');
  const tree = component({
    content: { kind: 'renderer', revision: 1, loaded() {}, failed() {}, reload() {} },
    resourceAddress: 'dsh-resource://file/abc/./deck.pptx',
    useTabInfo: () => ({ tab: { id: 't1', signal: new AbortController().signal } }),
    useResource: () => ({ value: undefined }),
    onSubmit: () => Promise.resolve(),
  });
  assert.ok(tree, 'renders something');
  assert.equal(tree.props.className, 'pa-state');
  assert.match(tree.children.join(''), /等待文件路径/);
});
