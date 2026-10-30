/**
 * Annotation model: hit testing over slide geometry and the prompt text an
 * agent receives when the user submits a review.
 *
 * The browser bundle cannot import this file (it ships as a single module), so
 * `lib/client.js` carries the same two pure functions. `test/client.test.js`
 * asserts both implementations agree.
 */

/** Element kinds that carry editable text, in label priority order. */
const TEXT_KINDS = new Set(['title', 'text', 'table']);

/**
 * Candidate elements under a point, most specific first.
 *
 * Text-bearing shapes win over decoration, and the smallest such shape wins, so
 * clicking a heading lands on the heading rather than the box behind it.
 * @param {Array<{x: number, y: number, w: number, h: number}>} shapes slide shapes
 * @param {number} x normalized x
 * @param {number} y normalized y
 * @param {number} [pad] hit tolerance as a fraction of slide size
 * @returns {{shape: any, index: number}|null} hit
 */
export function pickAt(shapes, x, y, pad = 0.004) {
  const hits = [];
  for (let i = 0; i < shapes.length; i += 1) {
    const s = shapes[i];
    if (x >= s.x - pad && x <= s.x + s.w + pad && y >= s.y - pad && y <= s.y + s.h + pad) {
      hits.push({ shape: s, index: i, area: s.w * s.h });
    }
  }
  if (!hits.length) return null;

  const textHits = hits.filter((h) => TEXT_KINDS.has(h.shape.kind) || (h.shape.text ?? '').trim());
  const pool = textHits.length ? textHits : hits;
  pool.sort((a, b) => a.area - b.area);
  return { shape: pool[0].shape, index: pool[0].index };
}

/**
 * Topmost element containing a point, in normalized slide coordinates.
 * @param {Array<{x: number, y: number, w: number, h: number}>} shapes slide shapes
 * @param {number} x normalized x
 * @param {number} y normalized y
 * @param {number} [pad] hit tolerance as a fraction of slide size
 * @returns {{shape: any, index: number}|null} hit
 */
export function hitTest(shapes, x, y, pad = 0.004) {
  for (let i = shapes.length - 1; i >= 0; i -= 1) {
    const s = shapes[i];
    if (x >= s.x - pad && x <= s.x + s.w + pad && y >= s.y - pad && y <= s.y + s.h + pad) {
      return { shape: s, index: i };
    }
  }
  return null;
}

/**
 * Short human label for an element.
 * @param {any} shape slide shape
 * @param {number} index element index on the slide
 * @returns {string} label
 */
export function shapeLabel(shape, index) {
  const text = (shape?.text ?? '').trim();
  if (text) return text.length > 40 ? `${text.slice(0, 40)}…` : text;
  const kindNames = {
    picture: '图片',
    table: '表格',
    chart: '图表',
    frame: '图形框',
    connector: '连接线',
    title: '标题',
    text: '文本框',
    shape: '形状',
    group: '组合',
  };
  return kindNames[shape?.kind] ?? `元素 ${index + 1}`;
}

/**
 * Sort annotations by slide then insertion order.
 * @param {Array<{slide: number, order?: number}>} annotations annotations
 * @returns {any[]} sorted copy
 */
export function sortAnnotations(annotations) {
  return [...annotations].sort(
    (a, b) => (a.slide - b.slide) || ((a.order ?? 0) - (b.order ?? 0)),
  );
}

/**
 * Build the prompt the agent receives for a review submission.
 * @param {{file: string, annotations: Array<{slide: number, note: string, label?: string, elementId?: string, order?: number}>}} input submission
 * @returns {string} prompt text
 */
export function formatPrompt({ file, annotations }) {
  const list = sortAnnotations(annotations).filter((a) => (a.note ?? '').trim());
  if (!list.length) return '';

  const lines = list.map((a, i) => {
    const where = [`第 ${a.slide} 页`];
    if (a.elementId) where.push(a.elementId);
    if (a.label) where.push(a.label);
    return `${i + 1}. [${where.join(' · ')}] ${a.note.trim().replace(/\s+/g, ' ')}`;
  });

  return [
    `【PPT 批注】\`${file}\` · 共 ${list.length} 条`,
    '',
    ...lines,
    '',
    '要求：只改这几处，不要动其他元素。每改完一条，立刻用 `✅ 第 X 页 · 元素：改了什么` 回报一次；全部改完再重新渲染复核一遍。',
  ].join('\n');
}
