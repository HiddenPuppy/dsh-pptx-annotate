/**
 * Slide rasterisation through the bundled LibreOffice Kit.
 *
 * The kit ships as a Node CLI inside the application (or the npm install), so
 * this module shells out to it instead of loading a native addon into the Host
 * process. Renders are cached per file version.
 */
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readdir, readFile, rm, stat } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { join } from 'node:path';

const CACHE_LIMIT = 8;
const cache = new Map();

/**
 * Candidate Node executables, most specific first.
 * @returns {Array<{command: string, env: Record<string, string>}>} candidates
 */
function nodeCandidates() {
  const out = [];
  if (process.versions.electron) {
    out.push({ command: process.execPath, env: { ELECTRON_RUN_AS_NODE: '1' } });
  }
  out.push({ command: process.execPath, env: {} });
  const bundled = join(homedir(), '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node');
  if (existsSync(bundled)) out.push({ command: bundled, env: {} });
  out.push({ command: 'node', env: {} });
  return out;
}

/**
 * Candidate LibreOffice Kit CLI entry points.
 * @returns {string[]} existing paths
 */
function cliCandidates() {
  const roots = [];
  if (process.resourcesPath) roots.push(join(process.resourcesPath, 'app.asar.unpacked/dsh/node_modules'));
  roots.push('/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/dsh/node_modules');
  roots.push(join(homedir(), '.local/share/dsh/node_modules'));
  roots.push(join(homedir(), '.dsh/node_modules'));
  return roots
    .map((r) => join(r, '@deepseek-ai/libreoffice-kit/lib/cli.js'))
    .filter((p) => existsSync(p));
}

/**
 * Run the kit CLI to completion.
 * @param {string} cli CLI path
 * @param {string[]} args CLI arguments
 * @returns {Promise<{stdout: string, stderr: string}>} captured output
 */
async function runCli(cli, args) {
  let lastError;
  for (const candidate of nodeCandidates()) {
    try {
      return await new Promise((resolve, reject) => {
        const child = spawn(candidate.command, [cli, ...args], {
          env: { ...process.env, ...candidate.env },
        });
        let stdout = '';
        let stderr = '';
        child.stdout.on('data', (d) => { stdout += d; });
        child.stderr.on('data', (d) => { stderr += d; });
        child.on('error', reject);
        child.on('close', (code) => {
          if (code === 0) resolve({ stdout, stderr });
          else reject(new Error(`libreoffice-kit exited ${code}: ${stderr.slice(-400)}`));
        });
      });
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError ?? new Error('no node executable available to run libreoffice-kit');
}

/**
 * Render every slide of a deck to PNG files, reusing a cached render when the
 * file version is unchanged.
 * @param {string} file absolute path to a .pptx
 * @returns {Promise<{dir: string, pages: string[], width: number, height: number}>} render
 */
export async function renderSlides(file) {
  const info = await stat(file);
  const key = `${file}:${info.size}:${info.mtimeMs}`;
  const hit = cache.get(key);
  if (hit) {
    cache.delete(key);
    cache.set(key, hit);
    return hit;
  }

  const cli = cliCandidates()[0];
  if (!cli) throw new Error('libreoffice-kit CLI not found');

  const dir = await mkdtemp(join(tmpdir(), 'pptx-annotate-'));
  const output = join(dir, 'render');
  const { stdout } = await runCli(cli, ['render', '--input', file, '--output-dir', output]);
  const manifest = JSON.parse(stdout.trim().split('\n').pop());
  const pages = manifest.images.map((i) => i.path);
  const first = manifest.images[0];
  const entry = { dir: output, pages, width: first.width, height: first.height };

  cache.set(key, entry);
  while (cache.size > CACHE_LIMIT) {
    const [oldestKey, oldest] = cache.entries().next().value;
    cache.delete(oldestKey);
    await rm(oldest.dir, { recursive: true, force: true }).catch(() => {});
  }
  return entry;
}

/**
 * Read one rendered slide PNG.
 * @param {string} file absolute path to a .pptx
 * @param {number} pageNumber 1-based slide number
 * @returns {Promise<{data: Buffer, width: number, height: number, count: number}>} image
 */
export async function renderSlide(file, pageNumber) {
  const { pages, width, height } = await renderSlides(file);
  const target = pages[pageNumber - 1];
  if (!target) throw new Error(`slide ${pageNumber} out of range (${pages.length})`);
  return { data: await readFile(target), width, height, count: pages.length };
}

/**
 * Drop cached renders (used by tests).
 * @returns {Promise<void>} completion
 */
export async function clearRenderCache() {
  const dirs = [...cache.values()].map((v) => v.dir);
  cache.clear();
  await Promise.all(dirs.map((d) => rm(d, { recursive: true, force: true }).catch(() => {})));
}

/**
 * List rendered files without parsing, for diagnostics.
 * @param {string} dir render directory
 * @returns {Promise<string[]>} file names
 */
export async function listRendered(dir) {
  return (await readdir(dir)).sort();
}

/**
 * Short content hash used in cache keys by callers.
 * @param {string} value input
 * @returns {string} hex digest
 */
export function shortHash(value) {
  return createHash('sha256').update(value).digest('hex').slice(0, 12);
}
