'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { G, W } = require('./helpers');

// A tiny assessment that exercises each kind of question.
const QUIZ = {
  id: 'demo', title: 'Demo',
  workbook: { sheets: [{ name: 'Sales', rows: [['Region', 'Units', 'Price', 'Revenue'], ['East', 2, 1.5], ['West', 4, 2], ['East', 1, 3]] }] },
  questions: [
    { id: 'rev', type: 'range', sheet: 'Sales', range: 'D2:D4', solution: ['=B2*C2', '=B3*C3', '=B4*C4'], requireFormula: true, points: 3 },
    { id: 'total', type: 'cell', sheet: 'Sales', cell: 'F1', solution: '=SUMIF(A2:A4,"East",D2:D4)', functions: [['SUMIF', 'SUMIFS']], points: 2 },
    { id: 'fmt', type: 'format', sheet: 'Sales', range: 'D2:D4', format: 'currency', points: 1 },
    { id: 'tbl', type: 'table', sheet: 'Sales', range: 'A1:D4', name: 'SalesTbl', points: 1 },
    { id: 'val', type: 'validation', sheet: 'Sales', range: 'B2:B4', solution: { type: 'whole', min: 0, max: 99 }, points: 1 },
    { id: 'cf', type: 'condFormat', sheet: 'Sales', range: 'D2:D4', solution: { op: '>', value: 5, style: 'green' }, points: 1 },
    { id: 'srt', type: 'sorted', sheet: 'Sales', range: 'A1:C4', column: 'B', order: 'desc', points: 1 },
    { id: 'mc', type: 'choice', options: ['a', 'b', 'c'], answer: 2, points: 1 },
    { id: 'multi', type: 'choice', multi: true, options: ['a', 'b', 'c', 'd'], answer: [0, 2], points: 2 },
    { id: 'num', type: 'number', answer: 3.5, tolerance: 0.01, points: 1 },
    { id: 'txt', type: 'text', accept: ['primary key', /^PK$/i], points: 1 },
    { id: 'match', type: 'match', pairs: [['Row', 'A record'], ['Column', 'A field']], points: 2 },
  ],
};
const answersRight = { mc: 2, multi: [0, 2], num: '3.5', txt: 'Primary  Key', match: { 0: 0, 1: 1 } };
const item = (r, id) => r.items.find((i) => i.id === id);

test('the answer key earns 100% and a blank submission earns nothing', () => {
  const key = G.buildKey(QUIZ);
  const full = G.grade(QUIZ, { workbook: key, answers: G.perfectAnswers(QUIZ) });
  assert.equal(full.percent, 100, JSON.stringify(full.items.filter((i) => i.status !== 'correct')));
  assert.equal(full.max, 17);
  const blank = G.grade(QUIZ, { answers: {} });
  assert.equal(blank.score, 0);
});

test('typed-in values where a formula is required lose half credit', () => {
  const wb = G.buildKey(QUIZ);
  wb.setInput('Sales', 'D3', 3, { force: true }); // the key is sorted by Units: D3 is East 2 × 1.5
  const r = item(G.grade(QUIZ, { workbook: wb }), 'rev');
  assert.equal(r.earned, 2.5);
  assert.match(r.feedback, /asks for a formula/);
});

test('right value with the wrong function earns partial credit and says why', () => {
  const wb = G.buildKey(QUIZ);
  wb.setInput('Sales', 'F1', '=D3+D4', { force: true });
  const r = item(G.grade(QUIZ, { workbook: wb }), 'total');
  assert.equal(r.status, 'partial');
  assert.match(r.feedback, /SUMIF or SUMIFS/);
});

test('wrong values are named; reveal shows the expected value after submission', () => {
  const wb = G.buildKey(QUIZ);
  wb.setInput('Sales', 'F1', 99, { force: true });
  const hidden = item(G.grade(QUIZ, { workbook: wb }), 'total');
  assert.match(hidden.feedback, /not correct yet/);
  assert.doesNotMatch(hidden.feedback, /expected/);
  const shown = item(G.grade(QUIZ, { workbook: wb }, { reveal: true }), 'total');
  assert.match(shown.feedback, /expected 6/);
});

test('structure questions explain what is missing', () => {
  const wb = G.buildTemplate(QUIZ);
  wb.addTable('Sales', 'A1:D3', 'Partial');
  wb.addValidation('Sales', 'B2:B4', { type: 'whole', min: 1, max: 10 });
  const r = G.grade(QUIZ, { workbook: wb });
  assert.match(item(r, 'tbl').feedback, /only covers A1:D3/);
  assert.equal(item(r, 'tbl').earned, 0.5);
  assert.match(item(r, 'val').feedback, /allow 0 to 99/);
  assert.match(item(r, 'cf').feedback, /No matching conditional formatting/);
  assert.match(item(r, 'srt').feedback, /not sorted/);
});

test('sorting that splits rows apart is caught', () => {
  const wb = G.buildTemplate(QUIZ);
  wb.sortRange('Sales', 'B1:B4', 0, false); // sorts only the Units column
  assert.match(item(G.grade(QUIZ, { workbook: wb }), 'srt').feedback, /changed or split apart/);
});

test('non-workbook questions: choice, multi-select, number, text and match', () => {
  const r = G.grade(QUIZ, { workbook: G.buildKey(QUIZ), answers: answersRight });
  for (const id of ['mc', 'multi', 'num', 'txt', 'match']) assert.equal(item(r, id).status, 'correct', id);
  const wrong = G.grade(QUIZ, { answers: { mc: 0, multi: [0, 1], num: 'abc', txt: 'foreign key', match: { 0: 1, 1: 1 } } });
  assert.equal(item(wrong, 'mc').earned, 0);
  assert.equal(item(wrong, 'multi').earned, 0); // one right, one wrong cancels out
  assert.match(item(wrong, 'num').feedback, /not a number/);
  assert.equal(item(wrong, 'txt').earned, 0);
  assert.equal(item(wrong, 'match').earned, 1);
  assert.equal(item(G.grade(QUIZ, { answers: { txt: 'pk' } }), 'txt').status, 'correct');
});

test('a missing sheet is reported instead of crashing', () => {
  const q = { id: 'x', questions: [{ id: 'c', type: 'cell', sheet: 'Nope', cell: 'A1', expect: 1 }] };
  const r = G.grade(q, { workbook: W.Workbook.fromSpec({ sheets: [{ name: 'S' }] }) });
  assert.match(r.items[0].feedback, /no sheet named "Nope"/);
});

test('valuesMatch is forgiving about case and spacing, strict about types', () => {
  assert.equal(G.valuesMatch('  East  Side', 'east side', {}), true);
  assert.equal(G.valuesMatch('5', 5, {}), false);
  assert.equal(G.valuesMatch(0.1 + 0.2, 0.3, {}), true);
  assert.equal(G.valuesMatch(10.004, 10, { tolerance: 0.005 }), true);
  assert.equal(G.valuesMatch('East', 'east', { caseSensitive: true }), false);
});

test('letter grades follow the usual scale', () => {
  assert.deepEqual([95, 91, 88, 84, 81, 78, 74, 71, 68, 62, 40].map(G.letterGrade), ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'F']);
});
