/**
 * End-to-end smoke test for the shipped browser bundle.
 *
 * Builds a self-contained page with a minimal React-compatible runtime, loads
 * `lib/client.js` through the real module-loader contract, serves the deck over
 * HTTP the way the Host routes do, then drives the UI in headless Chrome at a
 * wide and a narrow width: outline visibility, element picking, the toggle-off
 * rule, and the prompt that reaches the agent.
 *
 * Usage: node tools/smoke.mjs [--shot <png>]
 */
import { execFile } from 'node:child_process';
import { createServer } from 'node:http';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';

import { readDeck } from '../src/pptx-geom.js';
import { renderSlides } from '../src/render.js';

const run = promisify(execFile);
const here = fileURLToPath(new URL('..', import.meta.url));
const deckPath = join(here, 'fixtures/demo.pptx');
const chrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

/** Minimal React runtime: createElement, hooks, refs, DOM mounting. */
const MINI_REACT = String.raw`
let rerender = null;
const hookTable = new Map();
let renderPath = '';
let hookIndex = 0;

function createElement(type, props, ...children) {
  return { type, props: props || {}, children: children.flat(Infinity) };
}
function slotFor(i) {
  if (!hookTable.has(renderPath)) hookTable.set(renderPath, []);
  const arr = hookTable.get(renderPath);
  arr[i] = arr[i] || {};
  return arr[i];
}
function useState(initial) {
  const holder = slotFor(hookIndex++);
  if (!('value' in holder)) holder.value = typeof initial === 'function' ? initial() : initial;
  const set = (next) => {
    holder.value = typeof next === 'function' ? next(holder.value) : next;
    rerender();
  };
  return [holder.value, set];
}
function useRef(initial) {
  const holder = slotFor(hookIndex++);
  if (!('ref' in holder)) holder.ref = { current: initial };
  return holder.ref;
}
function useMemo(fn) {
  const holder = slotFor(hookIndex++);
  if (!('value' in holder)) holder.value = fn();
  return holder.value;
}
function useEffect(fn, deps) {
  const holder = slotFor(hookIndex++);
  const same = holder.deps && deps && holder.deps.length === deps.length &&
    holder.deps.every((d, i) => d === deps[i]);
  holder.deps = deps;
  if (!same) holder.pending = fn;
}
const React = { createElement, useState, useRef, useMemo, useEffect };

function runPending() {
  for (const [, slots] of hookTable) {
    for (const slot of slots) {
      if (slot && slot.pending) {
        const fn = slot.pending;
        slot.pending = null;
        try { fn(); } catch (error) { console.warn('effect failed', error); }
      }
    }
  }
}
function applyProps(el, props) {
  for (const [key, value] of Object.entries(props || {})) {
    if (key === 'children') continue;
    if (key === 'ref') { if (typeof value === 'function') value(el); continue; }
    if (key === 'className') el.className = value;
    else if (key === 'style') {
      for (const [prop, val] of Object.entries(value)) el.style[prop] = val;
    }
    else if (key === 'onChange' && typeof value === 'function') {
      const handler = () => value({ target: el });
      el.addEventListener('input', handler);
      el.addEventListener('change', handler);
    }
    else if (key.startsWith('on') && typeof value === 'function') {
      el.addEventListener(key.slice(2).toLowerCase(), value);
    }
    else if (key === 'value') el.value = value;
    else if (key === 'src') el.src = value;
    else if (key === 'disabled') el.disabled = !!value;
    else if (value !== undefined && value !== null) el.setAttribute(key, value);
  }
}
function mount(vnode, parent, path) {
  if (vnode === null || vnode === undefined || vnode === false || vnode === true) return;
  if (typeof vnode === 'string' || typeof vnode === 'number') {
    parent.appendChild(document.createTextNode(String(vnode)));
    return;
  }
  if (typeof vnode.type === 'function') {
    const prevPath = renderPath, prevIndex = hookIndex;
    renderPath = path; hookIndex = 0;
    const out = vnode.type(vnode.props || {});
    renderPath = prevPath; hookIndex = prevIndex;
    mount(out, parent, path);
    return;
  }
  const el = document.createElement(vnode.type);
  applyProps(el, vnode.props);
  (vnode.children || []).forEach((child, i) => mount(child, el, path + '.' + i));
  parent.appendChild(el);
}
function render(vnode, container) {
  rerender = () => {
    container.innerHTML = '';
    mount(vnode, container, 'root');
    runPending();
  };
  rerender();
}
`;

/** Harness page script: mount the shipped body and drive the UI. */
const HARNESS = String.raw`
(async () => {
  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail: detail === undefined ? '' : String(detail) });
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));

  const loaded = window.__MOD__;
  if (!loaded) throw new Error('client module never registered with the loader');
  const mod = loaded.factory((id) => {
    if (id === 'react') return React;
    throw new Error('unexpected external ' + id);
  });

  let reg = null;
  let body = null;
  const ctx = {
    effect: (fn) => fn(),
    documentPreviews: { register: (def) => { reg = def; return () => {}; } },
    slots: {
      inject: (name, factory) => { factory(); return () => {}; },
      register: (spec, component) => {
        if (spec.name === 'sidebar.right.tab.document') body = component;
        return () => {};
      },
    },
    sessions: { scope: () => ({ conversation: { send: () => Promise.resolve() } }) },
  };
  mod.apply(ctx);
  check('registers pptx previewer', reg && reg.extensions[0] === 'pptx' && reg.priority === 'extension', reg && reg.priority);
  check('registers body', typeof body === 'function');

  let submitted = '';
  render(React.createElement(body, {
    content: { kind: 'renderer', revision: 1, loaded: () => {}, failed: () => {}, reload: () => {} },
    resourceAddress: 'dsh-resource://file/s/./demo.pptx',
    useTabInfo: () => ({ tab: { id: 't1', signal: new AbortController().signal } }),
    useResource: () => ({ value: { absolutePath: '/fixtures/demo.pptx' } }),
    onSubmit: (text) => { submitted = text; return Promise.resolve(); },
  }), document.getElementById('host'));

  await wait(400);
  const wide = window.innerWidth >= 700;

  check('deck fetched and rendered', !!document.querySelector('[data-pptx-annotate]'));
  const slides = document.querySelectorAll('.pa-slide');
  check('renders every slide', slides.length === DECK.slides.length, slides.length + '/' + DECK.slides.length);

  const first = slides[0];
  const box = first.getBoundingClientRect();
  check('slide keeps the deck aspect ratio', Math.abs(box.width / box.height - DECK.width / DECK.height) < 0.05,
    (box.width / box.height).toFixed(3));
  check('slide never overflows its column', box.right <= window.innerWidth + 1, box.right + ' <= ' + window.innerWidth);

  const outline = document.querySelector('.pa-outline');
  const outlineShown = getComputedStyle(outline).display !== 'none';
  check(wide ? 'outline shown on a wide pane' : 'outline hidden on a narrow pane', outlineShown === wide, outlineShown);
  if (wide) check('outline lists every slide', outline.querySelectorAll('.pa-thumb').length === DECK.slides.length,
    outline.querySelectorAll('.pa-thumb').length);

  const hots = document.querySelectorAll('.pa-hot');
  check('overlays clickable elements', hots.length > 50, hots.length);
  check('submit disabled while empty', document.querySelector('.pa-btn.primary').disabled);

  const s2 = document.querySelectorAll('.pa-slide')[1];
  // The biggest text element on slide 2 is the heading we try to hit.
  const target = DECK.slides[1].shapes.map((s, i) => ({ s, i }))
    .filter((x) => (x.s.text || '').trim())
    .sort((a, b) => b.s.w * b.s.h - a.s.w * a.s.h)[0];
  const snippet = (target.s.text || '').slice(0, 6);
  const rect = s2.getBoundingClientRect();
  const cx = rect.left + (target.s.x + target.s.w / 2) * rect.width;
  const cy = rect.top + (target.s.y + target.s.h / 2) * rect.height;
  const hit = document.elementFromPoint(cx, cy);
  check('click lands on the heading element', hit && hit.classList.contains('pa-hot') && hit.title.includes(snippet),
    hit && hit.title);

  // Re-query before every click: a re-render replaces the mounted nodes.
  const clickTitle = () => document.elementFromPoint(cx, cy).click();

  clickTitle();
  await wait(60);
  const notes = document.querySelectorAll('.pa-note');
  check('click creates a note', notes.length === 1);
  check('note label names the picked element',
    notes.length > 0 && notes[0].textContent.includes(snippet),
    notes.length > 0 ? notes[0].textContent.slice(0, 40) : 'no note');

  clickTitle();
  await wait(60);
  check('second click clears an empty note', document.querySelectorAll('.pa-note').length === 0);

  clickTitle();
  await wait(60);
  const ta = document.querySelector('.pa-note textarea');
  check('note reopens after the toggle-off', !!ta);
  if (ta) {
    ta.value = '标题太大，缩到 40pt';
    ta.dispatchEvent(new Event('input', { bubbles: true }));
  }
  await wait(60);
  check('typing enables submit', !document.querySelector('.pa-btn.primary').disabled);
  clickTitle();
  await wait(60);
  check('second click keeps a written note', document.querySelectorAll('.pa-note').length === 1);

  if (location.search.includes('clean')) {
    // Screenshot mode: stop with one written note on screen.
    document.getElementById('result').textContent = JSON.stringify(results);
    return;
  }

  document.querySelector('.pa-btn.primary').click();
  await wait(150);
  check('submit reaches onSubmit', submitted.includes('标题太大'), submitted.slice(0, 60));
  check('prompt asks for point-by-point edits', submitted.includes('只改批注点到的元素'));
  check('prompt orders: change first, explain after', submitted.includes('先改后说'));
  check('prompt asks for optional suggestions afterwards', submitted.includes('可选优化建议'));
  check('notes cleared after submit', document.querySelectorAll('.pa-note').length === 0);

  const out = document.getElementById('result');
  out.textContent = JSON.stringify(results);
  out.setAttribute('data-done', '1');
})().catch((error) => {
  const out = document.getElementById('result');
  out.textContent = JSON.stringify([{ name: 'harness crashed', ok: false, detail: error && error.stack }]);
  out.setAttribute('data-done', '1');
});
`;

const body = await readFile(join(here, 'lib/client.js'), 'utf8');
const deck = readDeck(await readFile(deckPath));
const slides = await renderSlides(deckPath);
const images = [];
for (const page of slides.pages) {
  images.push(`data:image/png;base64,${(await readFile(page)).toString('base64')}`);
}

const html = `<!DOCTYPE html><meta charset="utf-8"><title>smoke</title>
<style>html,body{margin:0;height:100%} #host{height:100vh;width:100%;display:flex;flex-direction:column}</style>
<div id="host"></div><pre id="result"></pre>
<script>${MINI_REACT}</script>
<script>window.React = React;</script>
<script>window.__ModuleLoader__ = { load: (m) => { window.__MOD__ = m; } };</script>
<script>${body}</script>
<script>const DECK = ${JSON.stringify(deck)};</script>
<script>${HARNESS}</script>`;

const deckJson = JSON.stringify(deck);
const server = createServer((req, res) => {
  const url = new URL(req.url, 'http://127.0.0.1');
  if (url.pathname === '/api/pptx-annotate/deck') {
    res.writeHead(200, { 'content-type': 'application/json' });
    res.end(deckJson);
    return;
  }
  if (url.pathname === '/api/pptx-annotate/slide') {
    const n = Number(url.searchParams.get('n'));
    const data = Buffer.from(images[n - 1].split(',')[1], 'base64');
    res.writeHead(200, { 'content-type': 'image/png' });
    res.end(data);
    return;
  }
  res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
  res.end(html);
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const origin = `http://127.0.0.1:${server.address().port}`;

const dir = await mkdtemp(join(tmpdir(), 'pptx-smoke-'));
const file = join(dir, 'harness.html');
await writeFile(file, html, 'utf8');

if (!existsSync(chrome)) throw new Error('Chrome not found; cannot run the smoke test');

/**
 * Run the harness at one window width and parse its assertions.
 * @param {number} width window width in CSS pixels
 * @param {string} [shotPath] capture a screenshot instead of dumping the DOM
 * @returns {Promise<any[]>} assertions
 */
async function drive(width, shotPath) {
  const profile = join(dir, `chrome-${width}-${shotPath ? 'shot' : 'dom'}`);
  const args = [
    '--headless=new', '--no-sandbox', '--disable-gpu',
    '--disable-crash-reporter', '--disable-breakpad', '--no-first-run',
    '--no-default-browser-check', '--hide-scrollbars',
    `--user-data-dir=${profile}`, `--window-size=${width},900`,
    '--virtual-time-budget=8000',
  ];
  if (shotPath) args.push(`--screenshot=${shotPath}`);
  else args.push('--dump-dom');
  args.push(`${origin}/harness.html${shotPath ? '?clean=1' : ''}`);

  const { stdout } = await run(chrome, args, { maxBuffer: 64 * 1024 * 1024 })
    .catch((error) => ({ stdout: error.stdout ?? '' }));
  if (shotPath) return [];

  const match = /<pre id="result"[^>]*>([\s\S]*?)<\/pre>/.exec(stdout);
  if (!match) throw new Error('smoke harness produced no result block');
  const text = match[1]
    .replace(/&quot;/g, '"').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  return JSON.parse(text);
}

let failed = 0;
for (const width of [900, 380]) {
  const results = await drive(width);
  console.log(`\n— ${width}px —`);
  for (const r of results) {
    if (!r.ok) failed += 1;
    console.log(`${r.ok ? '✔' : '✖'} ${r.name}${r.detail ? `  (${r.detail})` : ''}`);
  }
  console.log(`${results.filter((r) => r.ok).length}/${results.length} checks passed`);
}

const shotAt = process.argv.indexOf('--shot');
if (shotAt > 0 && process.argv[shotAt + 1]) {
  await drive(760, process.argv[shotAt + 1]);
  console.log(`\nscreenshot → ${process.argv[shotAt + 1]}`);
}

server.close();
console.log(failed ? `\n${failed} check(s) failed` : '\nall checks passed');
process.exit(failed ? 1 : 0);
