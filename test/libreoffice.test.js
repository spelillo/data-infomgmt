'use strict';
// Files written by the site open in a real spreadsheet program, and files that
// program saves are read and graded correctly. LibreOffice stands in for Excel
// here because it runs headless in CI; the test is skipped when it is absent.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { X, G, allAssessments } = require('./helpers');

function findSoffice() {
  for (const bin of ['soffice', 'libreoffice']) {
    try { execFileSync(bin, ['--version'], { stdio: 'ignore', timeout: 60000 }); return bin; } catch (e) { /* try next */ }
  }
  return null;
}
const SOFFICE = findSoffice();
const WRITABLE = new Set(['cell', 'range', 'cellType', 'filled', 'format', 'table', 'validation', 'condFormat', 'sorted']);

test('every answer key → .xlsx → LibreOffice re-save → read back still earns full marks', { skip: !SOFFICE && 'LibreOffice is not installed', timeout: 300000 }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dim-lo-'));
  const inDir = path.join(dir, 'in'), outDir = path.join(dir, 'out');
  fs.mkdirSync(inDir); fs.mkdirSync(outDir);
  const list = allAssessments().filter(([, a]) => (a.questions || []).some((q) => WRITABLE.has(q.type)));
  list.forEach(([, a]) => fs.writeFileSync(path.join(inDir, a.id + '.xlsx'), X.write(G.buildKey(a))));
  execFileSync(SOFFICE, ['--headless', '-env:UserInstallation=file://' + path.join(dir, 'profile'), '--calc', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', outDir, ...list.map(([, a]) => path.join(inDir, a.id + '.xlsx'))], { stdio: 'ignore', timeout: 240000 });
  const problems = [];
  for (const [label, a] of list) {
    const file = path.join(outDir, a.id + '.xlsx');
    if (!fs.existsSync(file)) { problems.push(label + ': LibreOffice did not produce a file'); continue; }
    const back = await X.read(fs.readFileSync(file));
    back.sources = { ...((a.workbook && a.workbook.sources) || {}) };
    const r = G.grade(a, { workbook: back, answers: G.perfectAnswers(a) });
    r.items.filter((i) => WRITABLE.has(i.type) && i.earned < i.points).forEach((i) => problems.push(label + '/' + i.id + ': ' + i.feedback));
  }
  fs.rmSync(dir, { recursive: true, force: true });
  assert.deepEqual(problems, []);
});
