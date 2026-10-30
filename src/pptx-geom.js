/**
 * PPTX slide geometry extraction.
 *
 * Produces, per slide, the box of every drawable element in normalized slide
 * coordinates (0..1), so a previewer can overlay clickable regions on a
 * rendered image without understanding OOXML itself.
 */
import { openZip } from './zip.js';
import { parseXml, localName, textContent } from './xml.js';

const SHAPE_KINDS = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp']);

/** Identity transform: every coordinate maps to itself. */
const IDENTITY = { x: (v) => v, y: (v) => v, cx: (v) => v, cy: (v) => v };

/**
 * Read a deck into a slide → elements model.
 * @param {Uint8Array|Buffer} bytes `.pptx` archive bytes
 * @returns {{width: number, height: number, slides: Array<{index: number, shapes: any[]}>}}
 */
export function readDeck(bytes) {
  const zip = openZip(bytes);
  const read = (name) => {
    const buf = zip.read(name);
    return buf ? buf.toString('utf8') : null;
  };

  const presentationXml = read('ppt/presentation.xml');
  if (!presentationXml) throw new Error('not a pptx: ppt/presentation.xml missing');
  const pres = parseXml(presentationXml);
  const size = findFirst(pres, 'sldSz')?.attrs ?? {};
  const width = Number(size.cx ?? 0);
  const height = Number(size.cy ?? 0);
  if (!width || !height) throw new Error('pptx slide size missing');

  const slideParts = zip.names
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => slideNo(a) - slideNo(b));

  const slides = slideParts.map((part, i) => {
    const xml = read(part);
    const root = parseXml(xml);
    const tree = findFirst(root, 'spTree');
    if (!tree) return { index: i + 1, shapes: [] };

    const layout = loadRelated(root, zip, part, 'slideLayout');
    const master = layout ? loadRelated(layout.root, zip, layout.part, 'slideMaster') : null;

    const shapes = extractShapes(tree, { width, height }, IDENTITY);
    const inherited = resolvePlaceholders(shapes, layout, master, width, height);

    return {
      index: i + 1,
      shapes: inherited.filter((s) => s.w > 0 && s.h > 0),
    };
  });

  return { width, height, slides };
}

/**
 * Extract drawable elements from a shape tree, flattening groups.
 * @param {any} tree `p:spTree` node
 * @param {{width: number, height: number}} size slide size in EMU
 * @param {{x: number, y: number, cx: number, cy: number}} transform ancestor group transform
 * @returns {any[]} shapes
 */
function extractShapes(tree, size, transform) {
  const out = [];
  for (const node of tree.children ?? []) {
    if (node.name === '#text') continue;
    const kind = localName(node.name);
    if (!SHAPE_KINDS.has(kind)) continue;

    if (kind === 'grpSp') {
      const own = groupTransform(node, transform);
      out.push(...extractShapes(node, size, own));
      continue;
    }

    const ph = placeholderOf(node);
    const rect = shapeRect(node, kind, transform, size);
    if (!rect && !ph) continue;

    const box = rect ?? { x: 0, y: 0, cx: 0, cy: 0 };
    out.push({
      name: shapeName(node),
      kind: shapeKind(node, kind),
      ph,
      text: collapse(textContent(findFirst(node, 'txBody') ?? node)),
      x: box.x / size.width,
      y: box.y / size.height,
      w: box.cx / size.width,
      h: box.cy / size.height,
    });
  }
  return out;
}

/**
 * Resolve the absolute EMU box of one shape.
 * @param {any} node shape node
 * @param {string} kind local element name
 * @param {{x: number, y: number, cx: number, cy: number}} transform ancestor transform
 * @param {{width: number, height: number}} size slide size
 * @returns {{x: number, y: number, cx: number, cy: number}|null} box
 */
function shapeRect(node, kind, transform, size) {
  const xfrm = kind === 'graphicFrame'
    ? findFirst(node, 'xfrm')
    : findFirst(findFirst(node, 'spPr') ?? node, 'xfrm');
  const off = xfrm && findFirst(xfrm, 'off');
  const ext = xfrm && findFirst(xfrm, 'ext');
  if (!off || !ext) return null;
  const x = Number(off.attrs.x ?? 0);
  const y = Number(off.attrs.y ?? 0);
  const cx = Number(ext.attrs.cx ?? 0);
  const cy = Number(ext.attrs.cy ?? 0);
  return applyTransform({ x, y, cx, cy }, transform);
}

/**
 * Compose a group's child-coordinate transform with its ancestors.
 * @param {any} group `p:grpSp` node
 * @param {{x: number, y: number, cx: number, cy: number}} parent ancestor transform
 * @returns {{x: number, y: number, cx: number, cy: number}} composed transform
 */
function groupTransform(group, parent) {
  const xfrm = findFirst(findFirst(group, 'grpSpPr') ?? group, 'xfrm');
  const off = xfrm && findFirst(xfrm, 'off');
  const ext = xfrm && findFirst(xfrm, 'ext');
  const chOff = xfrm && findFirst(xfrm, 'chOff');
  const chExt = xfrm && findFirst(xfrm, 'chExt');
  if (!off || !ext) return parent;
  const own = {
    x: Number(off.attrs.x ?? 0),
    y: Number(off.attrs.y ?? 0),
    cx: Number(ext.attrs.cx ?? 0),
    cy: Number(ext.attrs.cy ?? 0),
    chX: Number(chOff?.attrs.x ?? off.attrs.x ?? 0),
    chY: Number(chOff?.attrs.y ?? off.attrs.y ?? 0),
    chCx: Number(chExt?.attrs.cx ?? ext.attrs.cx ?? 0),
    chCy: Number(chExt?.attrs.cy ?? ext.attrs.cy ?? 0),
  };
  const scaleX = own.chCx ? own.cx / own.chCx : 1;
  const scaleY = own.chCy ? own.cy / own.chCy : 1;
  return {
    x: (v) => parent.x(own.x + (v - own.chX) * scaleX),
    y: (v) => parent.y(own.y + (v - own.chY) * scaleY),
    cx: (v) => parent.cx(v * scaleX),
    cy: (v) => parent.cy(v * scaleY),
  };
}

/**
 * Apply a transform chain to a raw box.
 * @param {{x: number, y: number, cx: number, cy: number}} box raw box
 * @param {any} t transform function chain
 * @returns {{x: number, y: number, cx: number, cy: number}} transformed box
 */
function applyTransform(box, t) {
  return { x: t.x(box.x), y: t.y(box.y), cx: t.cx(box.cx), cy: t.cy(box.cy) };
}

/**
 * Give shapes without their own transform the position their placeholder
 * inherits from the slide layout, then from the master.
 * @param {any[]} shapes extracted slide shapes
 * @param {any} layout loaded layout part
 * @param {any} master loaded master part
 * @param {number} width slide width in EMU
 * @param {number} height slide height in EMU
 * @returns {any[]} shapes with boxes
 */
function resolvePlaceholders(shapes, layout, master, width, height) {
  return shapes.map((shape) => {
    if (shape.w > 0 || !shape.ph) return shape;
    const rect = lookupPlaceholder(shape.ph, layout) ?? lookupPlaceholder(shape.ph, master);
    if (!rect) return shape;
    return {
      ...shape,
      x: rect.x / width,
      y: rect.y / height,
      w: rect.cx / width,
      h: rect.cy / height,
    };
  });
}

/**
 * Find a layout/master placeholder box matching a slide placeholder.
 * @param {{type?: string, idx?: string}} ph slide placeholder reference
 * @param {any} part loaded layout or master
 * @returns {{x: number, y: number, cx: number, cy: number}|null} box in EMU
 */
function lookupPlaceholder(ph, part) {
  if (!part) return null;
  const list = part.placeholders ?? [];
  const byIdx = ph.idx !== undefined && list.find((p) => p.ph.idx === ph.idx && p.ph.type === ph.type);
  const byIdxLoose = ph.idx !== undefined && list.find((p) => p.ph.idx === ph.idx);
  const byType = ph.type && list.find((p) => p.ph.type === ph.type);
  const hit = byIdx || byIdxLoose || byType;
  return hit && hit.rect.cx > 0 && hit.rect.cy > 0 ? hit.rect : null;
}

/**
 * Load a slide's related layout or master and index its placeholders.
 * @param {any} root parsed part XML
 * @param {any} zip archive
 * @param {string} partPath part path, e.g. `ppt/slides/slide1.xml`
 * @param {string} relKind trailing relationship type segment
 * @returns {{root: any, placeholders: any[], part: string}|null} loaded part
 */
function loadRelated(root, zip, partPath, relKind) {
  const dir = partPath.slice(0, partPath.lastIndexOf('/'));
  const base = partPath.slice(partPath.lastIndexOf('/') + 1);
  const relsBuf = zip.read(`${dir}/_rels/${base}.rels`);
  if (!relsBuf) return null;
  const rels = parseXml(relsBuf.toString('utf8'));
  const rel = (rels.children ?? []).find((n) =>
    n.name !== '#text' && String(n.attrs.Type ?? '').endsWith(`/${relKind}`));
  if (!rel) return null;

  const target = normalizeTarget(dir, rel.attrs.Target ?? '');
  const buf = zip.read(target);
  if (!buf) return null;
  const partRoot = parseXml(buf.toString('utf8'));
  const tree = findFirst(partRoot, 'spTree');
  // width/height of 1 keeps EMU values as-is; the caller normalizes.
  const placeholders = tree
    ? extractShapes(tree, { width: 1, height: 1 }, IDENTITY)
        .filter((s) => s.ph && s.w > 0 && s.h > 0)
        .map((s) => ({ ph: s.ph, rect: { x: s.x, y: s.y, cx: s.w, cy: s.h } }))
    : [];
  return { root: partRoot, placeholders, part: target };
}

/**
 * Resolve a relationship target against its part directory.
 * @param {string} dir owning part directory
 * @param {string} target relationship target
 * @returns {string} absolute part path
 */
function normalizeTarget(dir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = `${dir}/${target}`.split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

/**
 * Placeholder reference of a shape, if any.
 * @param {any} node shape node
 * @returns {{type?: string, idx?: string}|null} placeholder
 */
function placeholderOf(node) {
  const ph = findFirst(node, 'ph');
  if (!ph) return null;
  const out = {};
  if (ph.attrs.type) out.type = ph.attrs.type;
  if (ph.attrs.idx !== undefined) out.idx = ph.attrs.idx;
  return Object.keys(out).length ? out : null;
}

/**
 * Human name of a shape from its non-visual properties.
 * @param {any} node shape node
 * @returns {string} name
 */
function shapeName(node) {
  const nv = findFirst(node, 'cNvPr');
  return nv?.attrs.name ?? '';
}

/**
 * Classify a shape for UI grouping.
 * @param {any} node shape node
 * @param {string} kind local element name
 * @returns {string} kind
 */
function shapeKind(node, kind) {
  if (kind === 'pic') return 'picture';
  if (kind === 'graphicFrame') {
    if (findFirst(node, 'tbl')) return 'table';
    if (findFirst(node, 'chart')) return 'chart';
    return 'frame';
  }
  if (kind === 'cxnSp') return 'connector';
  const ph = placeholderOf(node);
  if (ph && (ph.type === 'title' || ph.type === 'ctrTitle')) return 'title';
  return findFirst(node, 'txBody') ? 'text' : 'shape';
}

/**
 * Find the first descendant with a local name.
 * @param {any} node root
 * @param {string} name local name
 * @returns {any|null} match
 */
function findFirst(node, name) {
  if (!node || node.name === '#text') return null;
  for (const child of node.children ?? []) {
    if (child.name === '#text') continue;
    if (localName(child.name) === name) return child;
    const hit = findFirst(child, name);
    if (hit) return hit;
  }
  return null;
}

/**
 * Collapse whitespace in extracted text.
 * @param {string} text raw text
 * @returns {string} tidy text
 */
function collapse(text) {
  return text.replace(/\s+/g, ' ').trim();
}

/**
 * Numeric suffix of a slide part name.
 * @param {string} part path
 * @returns {number} slide number
 */
function slideNo(part) {
  return Number(/slide(\d+)\.xml$/.exec(part)?.[1] ?? 0);
}
