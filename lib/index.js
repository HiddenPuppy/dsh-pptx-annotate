/**
 * Host half of the PPTX annotation previewer.
 *
 * Two authenticated Fetch routes sit on the shared API channel:
 *   GET /api/pptx-annotate/deck?path=<absolute .pptx>
 *   GET /api/pptx-annotate/slide?path=<absolute .pptx>&n=<1-based>
 *
 * The browser half renders slides as images and overlays the per-element boxes
 * this module reports, so clicks land on real shapes rather than pixels.
 */
import { stat } from 'node:fs/promises';
import { appendFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { readDeck } from '../src/pptx-geom.js';
import { renderSlide } from '../src/render.js';

/** Diagnostic trace file, appended on every apply and route failure. */
const TRACE = join(tmpdir(), 'dsh-pptx-annotate.log');

/**
 * Append one diagnostic line.
 * @param {string} message line
 */
function trace(message) {
  try {
    appendFileSync(TRACE, `${new Date().toISOString()} ${message}\n`);
  } catch {
    /* diagnostics must never break the plugin */
  }
}

trace('module: evaluated');

/** Cordis plugin name. */
export const name = 'pptx-annotate';

/** Required Host services: the authenticated API channel. */
export const inject = ['connection'];

/** Refuse absurd inputs rather than reading a huge file into memory. */
const MAX_DECK_BYTES = 256 * 1024 * 1024;

/**
 * Validate a requested deck path.
 * @param {string|null} path absolute path from the query string
 * @returns {Promise<string>} the same path when usable
 */
async function assertDeck(path) {
  if (!path) throw new HttpError(400, 'missing path');
  if (!/\.pptx$/i.test(path)) throw new HttpError(400, 'only .pptx files are supported');
  let info;
  try {
    info = await stat(path);
  } catch {
    throw new HttpError(404, 'file not found');
  }
  if (!info.isFile()) throw new HttpError(404, 'not a regular file');
  if (info.size > MAX_DECK_BYTES) throw new HttpError(413, 'file too large');
  return path;
}

/** HTTP error carrying the status the route answers with. */
class HttpError extends Error {
  /**
   * @param {number} status HTTP status
   * @param {string} message human-readable reason
   */
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

/**
 * Answer one route request, mapping failures to JSON errors.
 * @param {(url: URL) => Promise<Response>} handler route body
 * @returns {(request: Request) => Promise<Response>} Fetch handler
 */
function route(handler) {
  return async (request) => {
    try {
      return await handler(new URL(request.url, 'http://localhost'));
    } catch (error) {
      const status = error instanceof HttpError ? error.status : 500;
      if (status >= 500) trace(`route error: ${error?.stack ?? error}`);
      return Response.json(
        { error: { status, message: error?.message ?? String(error) } },
        { status },
      );
    }
  };
}

/**
 * Host plugin body: register both routes for this plugin's lifetime.
 * @param {any} ctx Host context carrying `connection`
 */
export function apply(ctx) {
  trace('apply: start');
  try {
    registerRoutes(ctx);
    trace('apply: ok');
  } catch (error) {
    trace(`apply: FAILED ${error?.stack ?? error}`);
    throw error;
  }
}

/**
 * Register both routes on the shared API channel.
 * @param {any} ctx Host context carrying `connection`
 */
function registerRoutes(ctx) {
  ctx.connection.fetch.register({
    path: '/api/pptx-annotate/deck',
    methods: ['GET'],
    requestBody: 'buffered',
    fetch: route(async (url) => {
      const path = await assertDeck(url.searchParams.get('path'));
      trace(`deck: ${path}`);
      const info = await stat(path);
      const deck = readDeck(await readFileBuffer(path));
      return Response.json(
        { ...deck, version: info.mtimeMs, bytes: info.size },
        { headers: { 'cache-control': 'no-store' } },
      );
    }),
  });

  ctx.connection.fetch.register({
    path: '/api/pptx-annotate/version',
    methods: ['GET'],
    requestBody: 'buffered',
    fetch: route(async (url) => {
      const path = await assertDeck(url.searchParams.get('path'));
      const info = await stat(path);
      return Response.json(
        { version: info.mtimeMs, bytes: info.size },
        { headers: { 'cache-control': 'no-store' } },
      );
    }),
  });

  ctx.connection.fetch.register({
    path: '/api/pptx-annotate/slide',
    methods: ['GET'],
    requestBody: 'buffered',
    fetch: route(async (url) => {
      const path = await assertDeck(url.searchParams.get('path'));
      const n = Number(url.searchParams.get('n') ?? '1');
      if (!Number.isInteger(n) || n < 1) throw new HttpError(400, 'bad slide number');
      const slide = await renderSlide(path, n);
      return new Response(slide.data, {
        headers: {
          'content-type': 'image/png',
          'content-length': String(slide.data.length),
          'cache-control': 'no-store',
          'x-slide-count': String(slide.count),
        },
      });
    }),
  });
}

/**
 * Read a file into a Buffer.
 * @param {string} path file path
 * @returns {Promise<Buffer>} bytes
 */
async function readFileBuffer(path) {
  const { readFile } = await import('node:fs/promises');
  return readFile(path);
}
