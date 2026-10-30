/**
 * Minimal ZIP reader for PPTX containers.
 *
 * Reads the central directory and extracts members on demand. Stored (0) and
 * deflated (8) members are supported — the only two methods PowerPoint emits.
 */
import { inflateRawSync } from 'node:zlib';

const EOCD_SIG = 0x06054b50;
const CD_SIG = 0x02014b50;
const LFH_SIG = 0x04034b50;

/**
 * Parse the ZIP central directory.
 * @param {Uint8Array|Buffer} bytes complete archive bytes
 * @returns {{names: string[], has: (name: string) => boolean, read: (name: string) => Buffer|null}}
 */
export function openZip(bytes) {
  const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  const eocd = findEocd(buf);
  if (eocd < 0) throw new Error('not a zip archive: end-of-central-directory not found');

  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const entries = new Map();

  for (let i = 0; i < count; i += 1) {
    if (buf.readUInt32LE(off) !== CD_SIG) throw new Error(`corrupt central directory at entry ${i}`);
    const method = buf.readUInt16LE(off + 10);
    const compressedSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOffset = buf.readUInt32LE(off + 42);
    const name = buf.toString('utf8', off + 46, off + 46 + nameLen);
    entries.set(name, { method, compressedSize, localOffset });
    off += 46 + nameLen + extraLen + commentLen;
  }

  return {
    names: [...entries.keys()],
    has: (name) => entries.has(name),
    read(name) {
      const entry = entries.get(name);
      if (!entry) return null;
      const local = entry.localOffset;
      if (buf.readUInt32LE(local) !== LFH_SIG) throw new Error(`corrupt local header for ${name}`);
      const nameLen = buf.readUInt16LE(local + 26);
      const extraLen = buf.readUInt16LE(local + 28);
      const start = local + 30 + nameLen + extraLen;
      const raw = buf.subarray(start, start + entry.compressedSize);
      if (entry.method === 0) return Buffer.from(raw);
      if (entry.method === 8) return inflateRawSync(raw);
      throw new Error(`unsupported zip method ${entry.method} for ${name}`);
    },
  };
}

/**
 * Locate the end-of-central-directory record by scanning backwards.
 * @param {Buffer} buf archive bytes
 * @returns {number} offset, or -1
 */
function findEocd(buf) {
  const min = Math.max(0, buf.length - 66_000);
  for (let i = buf.length - 22; i >= min; i -= 1) {
    if (buf.readUInt32LE(i) === EOCD_SIG) return i;
  }
  return -1;
}
