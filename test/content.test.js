'use strict';
// Course integrity: every assessment is solvable to exactly 100% by its own
// answer key, a blank submission earns (almost) nothing, and the calendar
// matches the syllabus.
const test = require('node:test');
const assert = require('node:assert/strict');
const { F, G, C, allAssessments } = require('./helpers');

const ALL = allAssessments();

test('fifteen weekly lessons in three units, each complete', () => {
  assert.equal(C.lessons.length, 15);
  C.lessons.forEach((l, i) => {
    assert.equal(l.week, i + 1);
    assert.equal(l.unit, Math.ceil((i + 1) / 5));
    for (const k of ['id', 'title', 'summary']) assert.ok(l[k], l.id + ' ' + k);
    assert.ok(l.objectives.length >= 3, l.id + ' objectives');
    assert.ok(l.sections.length >= 3, l.id + ' sections');
    assert.ok(l.keyTerms.length >= 4, l.id + ' key terms');
    assert.ok(l.practice.length >= 1, l.id + ' practice workbook');
    assert.ok(l.assignment, l.id + ' assignment');
    l.sections.forEach((s) => assert.ok(s.heading && s.html.length > 200, l.id + ' section ' + s.heading));
  });
});

test('ids are unique across the whole course', () => {
  const ids = ALL.map(([, a]) => a.id);
  assert.equal(new Set(ids.filter((x) => !/^p\d/.test(x))).size, ids.filter((x) => !/^p\d/.test(x)).length);
  for (const [label, a] of ALL) {
    const qids = (a.questions || []).map((q) => q.id);
    assert.equal(new Set(qids).size, qids.length, label + ' has duplicate question ids');
  }
});

test('quiz and exam calendar matches the syllabus', () => {
  const quizWeeks = C.quizzes.map((q) => q.week).sort((a, b) => a - b);
  assert.deepEqual(quizWeeks, [2, 3, 4, 5, 7, 8, 9, 10, 12, 13, 14, 15]); // none in week 1 or the Monday after an exam
  assert.deepEqual(C.exams.map((e) => e.week), [5, 10, 15]);
  C.exams.forEach((e, i) => {
    assert.equal(e.unit, i + 1);
    assert.ok(e.covers.every((w) => w > i * 5 && w <= (i + 1) * 5), e.id + ' covers only its own unit');
    assert.equal(e.kind, 'exam');
  });
  C.quizzes.forEach((q) => { assert.equal(q.kind, 'quiz'); assert.ok(q.minutes > 0); });
});

test('primary and foreign keys are taught in week 7', () => {
  const w7 = C.lessons[6];
  const text = JSON.stringify(w7).toLowerCase();
  assert.match(text, /primary key/);
  assert.match(text, /foreign key/);
});

test('every formula in the course parses', () => {
  const formulas = [];
  const visit = (x, where) => {
    if (typeof x === 'string' && x.startsWith('=') && x.length > 1) formulas.push([x, where]);
    else if (Array.isArray(x)) x.forEach((y) => visit(y, where));
    else if (x && typeof x === 'object' && !(x instanceof RegExp)) Object.values(x).forEach((y) => visit(y, where));
  };
  for (const [label, a] of ALL) { visit(a.workbook, label + ' workbook'); (a.questions || []).forEach((q) => visit(q.solution, label + '/' + q.id)); }
  assert.ok(formulas.length > 100);
  for (const [f, where] of formulas) assert.doesNotThrow(() => F.parse(f.slice(1)), where + ': ' + f);
});

for (const [label, a] of ALL) {
  test('answer key scores 100%: ' + label, () => {
    const key = G.buildKey(a);
    const r = G.grade(a, { workbook: key, answers: G.perfectAnswers(a) });
    const off = r.items.filter((i) => i.status !== 'correct').map((i) => i.id + ': ' + i.feedback);
    assert.deepEqual(off, []);
    assert.equal(r.percent, 100);
    const blank = G.grade(a, { answers: {} });
    assert.ok(blank.percent < 10, 'a blank submission scored ' + blank.percent + '%');
    for (const q of a.questions || []) assert.ok(q.prompt || q.type === 'match', label + '/' + q.id + ' has a prompt');
  });
}

test('the authoring template (quiz-template.js) is a working, disabled example of every question type', () => {
  const before = C.quizzes.length;
  const tpl = require('../assets/js/content/quiz-template.js');
  assert.equal(C.quizzes.length, before, 'the template must not add itself to the course');
  const types = new Set(tpl.questions.map((q) => q.type));
  for (const t of ['choice', 'number', 'text', 'match', ...G.WORKBOOK_TYPES]) assert.ok(types.has(t), 'template shows ' + t);
  const r = G.grade(tpl, { workbook: G.buildKey(tpl), answers: G.perfectAnswers(tpl) });
  assert.equal(r.percent, 100, JSON.stringify(r.items.filter((i) => i.status !== 'correct')));
  assert.equal(G.grade(tpl, { answers: {} }).percent, 0);
});
