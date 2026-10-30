/**
 * Small non-validating XML reader.
 *
 * PPTX parts are machine-generated and well-formed, so a tokenizer that keeps
 * element names, attributes, and text is enough — and avoids a DOM dependency
 * in both the Host process and tests.
 */

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/**
 * Decode the five predefined XML entities plus numeric references.
 * @param {string} text raw text
 * @returns {string} decoded text
 */
export function decodeEntities(text) {
  if (!text.includes('&')) return text;
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X'
        ? parseInt(body.slice(2), 16)
        : parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
    }
    return ENTITIES[body] ?? whole;
  });
}

/**
 * Parse an XML document into a lightweight tree.
 * @param {string} source XML text
 * @returns {{name: string, attrs: Record<string, string>, children: any[]}}
 */
export function parseXml(source) {
  const root = { name: '#document', attrs: {}, children: [] };
  const stack = [root];
  const len = source.length;
  let i = 0;

  while (i < len) {
    const lt = source.indexOf('<', i);
    if (lt < 0) {
      appendText(source.slice(i));
      break;
    }
    if (lt > i) appendText(source.slice(i, lt));

    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      i = end < 0 ? len : end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', lt)) {
      const end = source.indexOf(']]>', lt + 9);
      const text = source.slice(lt + 9, end < 0 ? len : end);
      appendText(text, false);
      i = end < 0 ? len : end + 3;
      continue;
    }
    if (source[lt + 1] === '?' || source[lt + 1] === '!') {
      const end = source.indexOf('>', lt + 2);
      i = end < 0 ? len : end + 1;
      continue;
    }

    const gt = findTagEnd(source, lt + 1);
    if (gt < 0) break;
    const raw = source.slice(lt + 1, gt);
    i = gt + 1;

    if (raw[0] === '/') {
      const name = raw.slice(1).trim();
      popTo(name);
      continue;
    }

    const selfClosing = raw.endsWith('/');
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const space = body.search(/[\s]/);
    const name = space < 0 ? body : body.slice(0, space);
    const node = { name, attrs: space < 0 ? {} : parseAttrs(body.slice(space)), children: [] };
    stack[stack.length - 1].children.push(node);
    if (!selfClosing) stack.push(node);
  }

  return root.children.find((n) => n.name !== '#text') ?? root;

  function appendText(text, decode = true) {
    if (!text) return;
    stack[stack.length - 1].children.push({
      name: '#text',
      attrs: {},
      children: [],
      text: decode ? decodeEntities(text) : text,
    });
  }

  function popTo(name) {
    for (let s = stack.length - 1; s > 0; s -= 1) {
      if (stack[s].name === name) {
        stack.length = s;
        return;
      }
    }
  }
}

/**
 * Find the closing bracket of a tag, skipping quoted attribute values.
 * @param {string} source XML text
 * @param {number} from scan start
 * @returns {number} index of '>', or -1
 */
function findTagEnd(source, from) {
  let quote = '';
  for (let i = from; i < source.length; i += 1) {
    const ch = source[i];
    if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '>') {
      return i;
    }
  }
  return -1;
}

/**
 * Parse an attribute list.
 * @param {string} source text following the element name
 * @returns {Record<string, string>} attributes
 */
function parseAttrs(source) {
  const attrs = {};
  const re = /([^\s=/]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(source))) {
    attrs[m[1]] = decodeEntities(m[3] ?? m[4] ?? '');
  }
  return attrs;
}

/**
 * Local name of a possibly namespaced element (`a:off` → `off`).
 * @param {string} name element name
 * @returns {string} local name
 */
export function localName(name) {
  const colon = name.indexOf(':');
  return colon < 0 ? name : name.slice(colon + 1);
}

/**
 * Collect element children matching a predicate.
 * @param {any} node parent node
 * @param {(node: any) => boolean} predicate filter
 * @returns {any[]} matches
 */
export function childrenOf(node, predicate) {
  return (node?.children ?? []).filter((c) => c.name !== '#text' && predicate(c));
}

/**
 * All descendants (depth-first) with a local name in the given set.
 * @param {any} node root node
 * @param {string[]} names local names
 * @returns {any[]} matches
 */
export function findByLocalName(node, names) {
  const wanted = new Set(names);
  const out = [];
  const walk = (n) => {
    for (const child of n.children ?? []) {
      if (child.name === '#text') continue;
      if (wanted.has(localName(child.name))) out.push(child);
      walk(child);
    }
  };
  walk(node);
  return out;
}

/**
 * Concatenated text of an element subtree.
 * @param {any} node root node
 * @returns {string} text
 */
export function textContent(node) {
  let out = '';
  const walk = (n) => {
    for (const child of n.children ?? []) {
      if (child.name === '#text') out += child.text ?? '';
      else walk(child);
    }
  };
  if (node) walk(node);
  return out;
}
