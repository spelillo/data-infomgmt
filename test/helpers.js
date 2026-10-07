// Shared helpers for the test suite. Every module is plain browser JS that
// also exports itself for Node, so tests load the real files the site uses.
'use strict';
const path = require('path');

const ROOT = path.join(__dirname, '..');
const js = (p) => path.join(ROOT, 'assets', 'js', p);

const F = require(js('core/formula.js'));
const W = require(js('core/workbook.js'));
const Pivot = require(js('core/pivot.js'));
const Q = require(js('core/query.js'));
const X = require(js('core/xlsx.js'));
const G = require(js('core/grader.js'));
const C = require(js('content/course.js'));
['unit1', 'unit2', 'unit3', 'quizzes', 'exams'].forEach((f) => require(js('content/' + f + '.js')));

// Every graded or practice item in the course, as [label, assessment].
function allAssessments() {
  const out = [];
  for (const l of C.lessons) {
    for (const p of l.practice || []) out.push([l.id + '/' + p.id, p]);
    if (l.assignment) out.push([l.id + '/' + l.assignment.id, l.assignment]);
  }
  for (const q of C.quizzes || []) out.push([q.id, q]);
  for (const e of C.exams || []) out.push([e.id, e]);
  return out;
}

function findAssessment(id) {
  const hit = allAssessments().find(([label, a]) => a.id === id || label === id);
  if (!hit) throw new Error('No assessment ' + id);
  return hit[1];
}

// Quick workbook from rows: wb({ Data: [['A', 'B'], [1, 2]] })
function wb(sheets) {
  return W.Workbook.fromSpec({ sheets: Object.entries(sheets).map(([name, rows]) => ({ name, rows })) });
}

const enc = new TextEncoder();
const dec = new TextDecoder();

// Add or replace parts inside an .xlsx and return the new bytes. `patch` maps
// a part name to new text, or to a function (oldText) => newText.
async function patchXlsx(bytes, patch) {
  const files = await X.unzip(bytes);
  for (const [name, v] of Object.entries(patch)) {
    const old = files[name] ? dec.decode(files[name]) : '';
    const text = typeof v === 'function' ? v(old) : v;
    files[name] = typeof text === 'string' ? enc.encode(text) : text;
  }
  return X.zipStore(Object.entries(files).map(([name, data]) => ({ name, data })));
}

module.exports = { ROOT, F, W, Pivot, Q, X, G, C, allAssessments, findAssessment, wb, patchXlsx, enc, dec };
