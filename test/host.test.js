/**
 * Integration tests for the Host half: route registration and real responses
 * against the bundled LibreOffice Kit.
 */
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';

import { apply } from '../lib/index.js';
import { clearRenderCache } from '../src/render.js';

const deck = fileURLToPath(new URL('../fixtures/demo.pptx', import.meta.url));
const groups = fileURLToPath(new URL('../fixtures/groups.pptx', import.meta.url));

/** @type {Map<string, any>} */
const routes = new Map();

before(() => {
  apply({
    connection: {
      fetch: {
        register(route) {
          routes.set(route.path, route);
        },
      },
    },
  });
});

after(async () => {
  await clearRenderCache();
});

/**
 * Invoke a registered route the way the connection layer does.
 * @param {string} path registered path
 * @param {string} url full request URL
 * @returns {Promise<Response>} response
 */
function call(path, url) {
  const route = routes.get(path);
  assert.ok(route, `route ${path} registered`);
  return route.fetch(new Request(url));
}

test('host: both routes register on the shared API channel', () => {
  assert.deepEqual([...routes.keys()].sort(), [
    '/api/pptx-annotate/deck',
    '/api/pptx-annotate/slide',
    '/api/pptx-annotate/version',
  ]);
  for (const route of routes.values()) {
    assert.deepEqual(route.methods, ['GET']);
    assert.equal(route.requestBody, 'buffered');
  }
});

test('host: /deck returns the geometry model', async () => {
  const response = await call(
    '/api/pptx-annotate/deck',
    `http://localhost/api/pptx-annotate/deck?path=${encodeURIComponent(deck)}`,
  );
  assert.equal(response.status, 200);
  const body = await response.json();
  assert.equal(body.slides.length, 4);
  assert.equal(body.width, 12191695);
  assert.ok(body.slides[0].shapes.length > 3);
  assert.equal(body.slides[0].shapes[1].x.toFixed(3), '0.075');
});

test('host: /deck rejects a non-pptx path', async () => {
  const response = await call(
    '/api/pptx-annotate/deck',
    'http://localhost/api/pptx-annotate/deck?path=' + encodeURIComponent('/etc/hosts'),
  );
  assert.equal(response.status, 400);
  const body = await response.json();
  assert.match(body.error.message, /only \.pptx/);
});

test('host: /deck reports a missing file as 404', async () => {
  const response = await call(
    '/api/pptx-annotate/deck',
    'http://localhost/api/pptx-annotate/deck?path=' + encodeURIComponent('/tmp/nope.pptx'),
  );
  assert.equal(response.status, 404);
});

test('host: /version reports the file stamp used for cache busting', async () => {
  const first = await call(
    '/api/pptx-annotate/version',
    `http://localhost/api/pptx-annotate/version?path=${encodeURIComponent(deck)}`,
  );
  assert.equal(first.status, 200);
  const body = await first.json();
  assert.ok(body.version > 0, 'version is a positive stamp');
  assert.ok(body.bytes > 1000, 'size reported');

  const again = await (await call(
    '/api/pptx-annotate/version',
    `http://localhost/api/pptx-annotate/version?path=${encodeURIComponent(deck)}`,
  )).json();
  assert.equal(again.version, body.version, 'stable while the file is unchanged');
});

test('host: /deck carries the same version stamp', async () => {
  const deckBody = await (await call(
    '/api/pptx-annotate/deck',
    `http://localhost/api/pptx-annotate/deck?path=${encodeURIComponent(deck)}`,
  )).json();
  const version = await (await call(
    '/api/pptx-annotate/version',
    `http://localhost/api/pptx-annotate/version?path=${encodeURIComponent(deck)}`,
  )).json();
  assert.equal(deckBody.version, version.version);
});

test('host: /slide renders a real PNG for the requested page', async () => {
  const response = await call(
    '/api/pptx-annotate/slide',
    `http://localhost/api/pptx-annotate/slide?path=${encodeURIComponent(groups)}&n=1`,
  );
  assert.equal(response.status, 200);
  assert.equal(response.headers.get('content-type'), 'image/png');
  assert.equal(response.headers.get('x-slide-count'), '1');
  const bytes = Buffer.from(await response.arrayBuffer());
  assert.equal(bytes.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
  assert.ok(bytes.length > 5000);
});

test('host: /slide rejects an out-of-range page', async () => {
  const response = await call(
    '/api/pptx-annotate/slide',
    `http://localhost/api/pptx-annotate/slide?path=${encodeURIComponent(groups)}&n=9`,
  );
  assert.equal(response.status, 500);
  const body = await response.json();
  assert.match(body.error.message, /out of range/);
});
