/**
 * Unit tests for the PPTX geometry pipeline (zip → xml → normalized boxes).
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { openZip } from '../src/zip.js';
import { parseXml, textContent, decodeEntities } from '../src/xml.js';
import { readDeck } from '../src/pptx-geom.js';

const fixture = (name) => readFileSync(fileURLToPath(new URL(`../fixtures/${name}`, import.meta.url)));

test('zip: lists parts and inflates XML members', () => {
  const zip = openZip(fixture('demo.pptx'));
  assert.ok(zip.names.includes('ppt/presentation.xml'));
  assert.ok(zip.names.includes('ppt/slides/slide1.xml'));
  const xml = zip.read('ppt/presentation.xml').toString('utf8');
  assert.match(xml, /^<\?xml/);
  assert.equal(zip.read('nope.xml'), null);
});

test('xml: parses attributes, self-closing tags, comments and entities', () => {
  const root = parseXml(
    '<?xml version="1.0"?><!-- c --><a:root x="1"><b y=\'two\'/><c>a &amp; b &#65;</c></a:root>',
  );
  assert.equal(root.name, 'a:root');
  assert.equal(root.attrs.x, '1');
  const b = root.children.find((n) => n.name === 'b');
  assert.equal(b.attrs.y, 'two');
  assert.match(textContent(root), /a & b A/);
  assert.equal(decodeEntities('&#x42;'), 'B');
});

test('geometry: demo deck yields 4 slides and 99 elements', () => {
  const deck = readDeck(fixture('demo.pptx'));
  assert.equal(deck.slides.length, 4);
  assert.equal(deck.slides.reduce((n, s) => n + s.shapes.length, 0), 99);
  assert.equal(deck.width, 12191695);
});

test('geometry: every box is a sane fraction of the slide', () => {
  const deck = readDeck(fixture('demo.pptx'));
  for (const slide of deck.slides) {
    for (const s of slide.shapes) {
      for (const v of [s.x, s.y, s.w, s.h]) {
        assert.ok(Number.isFinite(v), 'finite');
      }
      assert.ok(s.x >= -0.01 && s.x <= 1, `x in range: ${s.x}`);
      assert.ok(s.y >= -0.01 && s.y <= 1, `y in range: ${s.y}`);
      assert.ok(s.w > 0 && s.w <= 1, `w in range: ${s.w}`);
      assert.ok(s.h > 0 && s.h <= 1, `h in range: ${s.h}`);
    }
  }
});

test('geometry: slide-1 title text and box survive the round trip', () => {
  const deck = readDeck(fixture('demo.pptx'));
  const title = deck.slides[0].shapes.find((s) => s.text.includes('真正的差距'));
  assert.ok(title, 'title found');
  assert.ok(title.x > 0.05 && title.x < 0.1, `x ≈ 0.075, got ${title.x}`);
  assert.equal(title.kind, 'text');
});

test('geometry: placeholders inherit boxes from layout and master', () => {
  const deck = readDeck(fixture('placeholders.pptx'));
  assert.equal(deck.slides.length, 2);
  const first = deck.slides[0].shapes;
  const title = first.find((s) => s.kind === 'title');
  assert.ok(title, 'title placeholder present');
  assert.equal(title.text, '占位符标题');
  assert.ok(title.w > 0.5 && title.h > 0.05, 'title got a real box');
  const body = first.find((s) => s.text === '副标题占位符');
  assert.ok(body && body.w > 0.4, 'subtitle placeholder got a real box');
});

test('geometry: grouped shapes are flattened into slide coordinates', () => {
  const deck = readDeck(fixture('groups.pptx'));
  const shapes = deck.slides[0].shapes;
  assert.equal(shapes.length, 3, 'group children surface individually');
  const rect = shapes.find((s) => s.text === '组内方块');
  assert.ok(rect, 'group child found');
  assert.ok(Math.abs(rect.x - 1 / 13.333) < 0.01, `x ≈ 0.075, got ${rect.x}`);
  assert.ok(Math.abs(rect.w - 2 / 13.333) < 0.01, `w ≈ 0.150, got ${rect.w}`);
  const outside = shapes.find((s) => s.text === '组外文本');
  assert.ok(Math.abs(outside.x - 8 / 13.333) < 0.01, 'sibling outside the group keeps its own box');
});
