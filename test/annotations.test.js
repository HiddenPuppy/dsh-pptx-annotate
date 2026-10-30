/**
 * Unit tests for the annotation model: hit testing, labels, prompt shape.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';

import { hitTest, pickAt, shapeLabel, sortAnnotations, formatPrompt } from '../src/annotations.js';

const shapes = [
  { x: 0, y: 0, w: 1, h: 1, kind: 'shape', text: '' },
  { x: 0.1, y: 0.1, w: 0.2, h: 0.2, kind: 'title', text: '标题' },
  { x: 0.1, y: 0.1, w: 0.2, h: 0.2, kind: 'text', text: '覆盖在上面的文本' },
];

test('hitTest: later shapes win, matching paint order', () => {
  const hit = hitTest(shapes, 0.15, 0.15);
  assert.ok(hit);
  assert.equal(hit.index, 2);
  assert.equal(hit.shape.text, '覆盖在上面的文本');
});

test('hitTest: misses outside every box', () => {
  assert.equal(hitTest(shapes.slice(1), 0.9, 0.9), null);
});

test('hitTest: tolerance makes thin elements clickable', () => {
  const thin = [{ x: 0.5, y: 0.5, w: 0.001, h: 0.001 }];
  assert.ok(hitTest(thin, 0.503, 0.5, 0.004));
});

test('pickAt: a heading beats the decorative box behind it', () => {
  const slide = [
    { x: 0.05, y: 0.05, w: 0.9, h: 0.3, kind: 'shape', text: '' },
    { x: 0.06, y: 0.06, w: 0.02, h: 0.03, kind: 'shape', text: '' },
    { x: 0.06, y: 0.1, w: 0.8, h: 0.12, kind: 'text', text: '大标题' },
  ];
  const hit = pickAt(slide, 0.5, 0.15);
  assert.equal(hit.shape.text, '大标题');
});

test('pickAt: without text, the smallest decoration wins', () => {
  const slide = [
    { x: 0, y: 0, w: 0.9, h: 0.9, kind: 'shape', text: '' },
    { x: 0.1, y: 0.1, w: 0.2, h: 0.2, kind: 'picture', text: '' },
  ];
  assert.equal(pickAt(slide, 0.15, 0.15).shape.kind, 'picture');
});

test('pickAt: empty space returns null', () => {
  assert.equal(pickAt([{ x: 0, y: 0, w: 0.1, h: 0.1, kind: 'shape', text: '' }], 0.8, 0.8), null);
});

test('shapeLabel: prefers text, falls back to kind', () => {
  assert.equal(shapeLabel({ text: '季度复盘' }, 0), '季度复盘');
  assert.equal(shapeLabel({ text: '', kind: 'table' }, 3), '表格');
  assert.equal(shapeLabel({ text: '' }, 4), '元素 5');
});

test('shapeLabel: truncates very long text', () => {
  const label = shapeLabel({ text: 'x'.repeat(100) }, 0);
  assert.ok(label.length <= 41);
  assert.ok(label.endsWith('…'));
});

test('sortAnnotations: slide order then insertion order', () => {
  const sorted = sortAnnotations([
    { slide: 2, order: 2 },
    { slide: 1, order: 3 },
    { slide: 1, order: 1 },
  ]);
  assert.deepEqual(sorted.map((a) => `${a.slide}:${a.order}`), ['1:1', '1:3', '2:2']);
});

test('formatPrompt: numbered, localized, and states the constraint', () => {
  const text = formatPrompt({
    file: 'deck.pptx',
    annotations: [
      { slide: 3, elementId: '#7', label: '表格', note: '换成三列卡片', order: 2 },
      { slide: 1, elementId: '#2', label: '标题', note: '  字号太大\n缩到 40pt ', order: 1 },
    ],
  });
  assert.match(text, /【PPT 批注】`deck\.pptx` · 共 2 条/);
  assert.match(text, /1\. \[第 1 页 · #2 · 标题\] 字号太大 缩到 40pt/);
  assert.match(text, /2\. \[第 3 页 · #7 · 表格\] 换成三列卡片/);
  assert.match(text, /只改这几处/);
});

test('formatPrompt: empty input yields empty string', () => {
  assert.equal(formatPrompt({ file: 'a.pptx', annotations: [] }), '');
  assert.equal(formatPrompt({ file: 'a.pptx', annotations: [{ slide: 1, note: '   ' }] }), '');
});
